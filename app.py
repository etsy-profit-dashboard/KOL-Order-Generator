import io
import re
import unicodedata
from pathlib import Path
from difflib import SequenceMatcher
from datetime import datetime

import requests
import pandas as pd
import streamlit as st

st.set_page_config(page_title="WMS Thai Address Converter", page_icon="📦", layout="wide")

POSTAL_DATA_URL = (
    "https://raw.githubusercontent.com/"
    "thailand-geography-data/thailand-geography-json/main/src/geography.json"
)
FUZZY_THRESHOLD = 0.86
AUTO_ACCEPT_FUZZY_THRESHOLD = 0.93


def clean_text(x):
    if x is None or pd.isna(x):
        return ""
    s = unicodedata.normalize("NFC", str(x))
    # Remove invisible / zero-width characters from WMS / Excel.
    for ch in [
        "\u200b", "\u200c", "\u200d", "\u200e", "\u200f",
        "\u202a", "\u202b", "\u202c", "\u202d", "\u202e",
        "\u2060", "\ufeff",
    ]:
        s = s.replace(ch, "")
    s = s.replace("\xa0", " ").replace("\n", " ").replace("\r", " ").replace("\t", " ")
    return re.sub(r"\s+", " ", s).strip()


def norm_thai(x):
    s = clean_text(x).lower()
    if not s:
        return ""
    replacements = {
        "เเ": "แ", "ํา": "ำ", "ํ": "",
        "": "่", "": "้", "": "๊", "": "๋", "": "์",
    }
    for old, new in replacements.items():
        s = s.replace(old, new)

    for pattern in [
        r"^จังหวัด", r"^กรุงเทพมหานครจังหวัด", r"^อำเภอ", r"^เขต",
        r"^แขวง", r"^ตำบล", r"^ตําบล", r"^จ\.", r"^อ\.", r"^ต\.",
    ]:
        s = re.sub(pattern, "", s)

    aliases = {
        "กรุงเทพฯ": "กรุงเทพมหานคร",
        "กรุงเทพ": "กรุงเทพมหานคร",
        "กทม": "กรุงเทพมหานคร",
        "กทม.": "กรุงเทพมหานคร",
        "กรุงเทพเทพมหานคร": "กรุงเทพมหานคร",
        "จังหวัดกรุงเทพเทพมหานคร": "กรุงเทพมหานคร",
        "อยุธยา": "พระนครศรีอยุธยา",
        "จังหวัดอยุธยา": "พระนครศรีอยุธยา",
        "ลําปาง": "ลำปาง",
        "ลาปาง": "ลำปาง",
    }
    s = aliases.get(s, s)
    return re.sub(r"[\s\-_–—/\\.,;:|()\[\]{}]+", "", s)


def clean_location_name(text):
    if text is None or pd.isna(text):
        return ""
    text = clean_text(text)
    text = re.sub(
        r"^(จังหวัด|จ\.|อำเภอ|อ\.|เขต|แขวง|ตำบล|ต\.|ตําบล)\s*",
        "", text, flags=re.IGNORECASE
    )
    return {
        "กรุงเทพ": "กรุงเทพมหานคร",
        "กทม": "กรุงเทพมหานคร",
        "กทม.": "กรุงเทพมหานคร",
        "กรุงเทพฯ": "กรุงเทพมหานคร",
        "กรุงเทพเทพมหานคร": "กรุงเทพมหานคร",
        "Bangkok": "กรุงเทพมหานคร",
        "BANGKOK": "กรุงเทพมหานคร",
        "bangkok": "กรุงเทพมหานคร",
    }.get(text.strip(), text.strip())


def thai_part(x):
    x = clean_text(x)
    return x.split("/", 1)[0].strip() if "/" in x else x


def english_part(x):
    x = clean_text(x)
    return x.split("/", 1)[1].strip() if "/" in x else ""


def make_bilingual(th, en):
    th, en = clean_text(th), clean_text(en)
    if not th:
        return ""
    return f"{th}/ {en}" if en else th


def clean_mobile_number(val):
    if val is None or pd.isna(val):
        return ""
    text = str(val).strip()
    if "e+" in text.lower():
        try:
            text = str(int(float(text)))
        except Exception:
            pass
    else:
        text = text.split(".")[0]
    text = re.sub(r"\D", "", text)
    if text.startswith("66") and len(text) >= 11:
        text = "0" + text[2:]
    elif len(text) == 9:
        text = "0" + text
    return text


def valid_postal(x):
    if x is None or pd.isna(x):
        return ""
    m = re.search(r"(?<!\d)(\d{5})(?!\d)", clean_text(x))
    return m.group(1) if m else ""


def extract_postal_code(address):
    return valid_postal(address)


def get_wms_postal(row):
    # WMS already contains the receiver postal code in many exports.
    for col in ["Receiver Zipcode", "Receiver Zip", "Postal Code", "Zipcode", "Zip Code"]:
        if col in row.index:
            code = valid_postal(row.get(col, ""))
            if code:
                return code
    return ""


def extract_address_components(address):
    address = clean_text(address)
    result = {"province": "", "district": "", "subdistrict": ""}
    if not address:
        return result

    thai_name = r"([ก-๙เแโใไฤฦะาิีึืุูํ่้๊๋์ฯ\-]+)"

    for pattern in [rf"(?:จังหวัด|จ\.)\s*{thai_name}"]:
        matches = re.findall(pattern, address)
        if matches:
            result["province"] = max(matches, key=len)
            break

    for pattern in [
        rf"(?:อำเภอ|อ\.)\s*{thai_name}",
        rf"(?:เขต)\s*{thai_name}",
    ]:
        matches = re.findall(pattern, address)
        if matches:
            result["district"] = max(matches, key=len)
            break

    for pattern in [
        rf"(?:ตำบล|ต\.)\s*{thai_name}",
        rf"(?:ตําบล)\s*{thai_name}",
        rf"(?:แขวง)\s*{thai_name}",
    ]:
        matches = re.findall(pattern, address)
        if matches:
            result["subdistrict"] = max(matches, key=len)
            break
    return result


def similarity(a, b):
    a, b = norm_thai(a), norm_thai(b)
    if not a or not b:
        return 0.0
    if a == b:
        return 1.0
    return SequenceMatcher(None, a, b).ratio()


def exact_match(items, value):
    key = norm_thai(clean_location_name(value))
    if not key:
        return None
    for item in items:
        if item.get("key") == key:
            return item
    return None


def fuzzy_match(items, candidates):
    best_item, best_score, best_candidate = None, 0.0, ""
    usable = [clean_location_name(x) for x in candidates if clean_location_name(x)]
    if not usable:
        return None, 0.0, ""
    for item in items:
        if not item.get("key"):
            continue
        for candidate in usable:
            score = similarity(item["key"], candidate)
            if score > best_score:
                best_score, best_item, best_candidate = score, item, candidate
    if best_item is not None and best_score >= FUZZY_THRESHOLD:
        return best_item, best_score, best_candidate
    return None, 0.0, ""


@st.cache_data(ttl=86400, show_spinner=False)
def load_postal_data():
    r = requests.get(POSTAL_DATA_URL, timeout=30)
    r.raise_for_status()
    return r.json()


def load_address_master(uploaded=None):
    if uploaded is not None:
        return pd.read_excel(uploaded)
    base_path = Path(__file__).parent / "data"
    for path in [
        base_path / "adress.xlsx",
        base_path / "adress(1).xlsx",
        base_path / "address.xlsx",
        base_path / "address(1).xlsx",
    ]:
        if path.exists():
            return pd.read_excel(path)
    raise FileNotFoundError(
        "ไม่พบไฟล์ adress.xlsx ในโฟลเดอร์ data\nPath: " + str(base_path)
    )


def prepare_master(df):
    df = df.copy()
    required_columns = ["จังหวัด", "เขต/อำเภอ"]
    if not set(required_columns).issubset(df.columns):
        raise ValueError("ไฟล์ adress.xlsx ต้องมีคอลัมน์ จังหวัด และ เขต/อำเภอ")

    df["จังหวัด"] = df["จังหวัด"].ffill()
    df["province_th"] = df["จังหวัด"].map(thai_part)
    df["province_en"] = df["จังหวัด"].map(english_part)
    df["district_th"] = df["เขต/อำเภอ"].map(thai_part)
    df["district_en"] = df["เขต/อำเภอ"].map(english_part)
    df["pkey"] = df["province_th"].map(norm_thai)
    df["dkey"] = df["district_th"].map(norm_thai)

    province_master = (
        df[["pkey", "province_th", "province_en"]]
        .drop_duplicates(subset=["pkey"], keep="first")
        .set_index("pkey")
    )
    district_master = (
        df[["dkey", "district_th", "district_en"]]
        .drop_duplicates(subset=["dkey"], keep="first")
        .set_index("dkey")
    )
    return df, province_master, district_master


def build_geo_tables(records):
    geo = pd.DataFrame(records)
    columns = [
        "provinceNameTh", "provinceNameEn", "districtNameTh",
        "districtNameEn", "subdistrictNameTh", "subdistrictNameEn",
    ]
    for c in columns:
        if c not in geo.columns:
            geo[c] = ""
        geo[c] = geo[c].fillna("").astype(str).map(clean_text)

    geo["pkey"] = geo["provinceNameTh"].map(norm_thai)
    geo["dkey"] = geo["districtNameTh"].map(norm_thai)
    geo["skey"] = geo["subdistrictNameTh"].map(norm_thai)

    if "postalCode" not in geo.columns:
        geo["postalCode"] = ""
    geo["postalCode"] = geo["postalCode"].fillna("").astype(str).map(valid_postal)
    return geo


def build_provinces(geo):
    out = []
    for (pkey, pth), g in geo.groupby(["pkey", "provinceNameTh"], sort=False):
        if pkey:
            out.append({"key": pkey, "th": pth, "en": g.iloc[0]["provinceNameEn"]})
    return out


def build_districts(gprov):
    out = []
    for (dkey, dth), g in gprov.groupby(["dkey", "districtNameTh"], sort=False):
        if dkey:
            out.append({"key": dkey, "th": dth, "en": g.iloc[0]["districtNameEn"]})
    return out


def build_subdistricts(gprov):
    out = []
    for (skey, sth), g in gprov.groupby(["skey", "subdistrictNameTh"], sort=False):
        if skey:
            out.append({"key": skey, "th": sth, "en": g.iloc[0]["subdistrictNameEn"]})
    return out


def resolve_province(geo, prov_raw, city_raw, area_raw, address):
    provinces = build_provinces(geo)
    parts = extract_address_components(address)
    addr_norm = norm_thai(address)

    p = exact_match(provinces, prov_raw)
    if p is not None:
        return p, "province_exact"

    p = exact_match(provinces, parts["province"])
    if p is not None:
        return p, "address_province"

    embedded = [p for p in provinces if p["key"] and p["key"] in addr_norm]
    if embedded:
        return max(embedded, key=lambda x: len(x["key"])), "address_province_embedded"

    for candidate in [city_raw, area_raw]:
        p = exact_match(provinces, candidate)
        if p is not None:
            return p, "province_from_city_area_exact"

    if prov_raw:
        p, score, _ = fuzzy_match(provinces, [prov_raw])
        if p is not None:
            return p, f"province_fuzzy_auto_{score:.2f}" if score >= AUTO_ACCEPT_FUZZY_THRESHOLD else f"province_fuzzy_check_{score:.2f}"

    if parts["province"]:
        p, score, _ = fuzzy_match(provinces, [parts["province"]])
        if p is not None:
            return p, f"address_province_fuzzy_auto_{score:.2f}" if score >= AUTO_ACCEPT_FUZZY_THRESHOLD else f"address_province_fuzzy_check_{score:.2f}"

    if not prov_raw:
        p, score, _ = fuzzy_match(provinces, [city_raw, area_raw])
        if p is not None:
            return p, f"fallback_province_fuzzy_auto_{score:.2f}" if score >= AUTO_ACCEPT_FUZZY_THRESHOLD else f"fallback_province_fuzzy_check_{score:.2f}"

    return None, "province_not_found"


def resolve_district(gprov, city_raw, area_raw, address):
    districts = build_districts(gprov)
    subdistricts = build_subdistricts(gprov)
    parts = extract_address_components(address)

    province_name = gprov.iloc[0]["provinceNameTh"] if not gprov.empty else ""
    if norm_thai(city_raw) == "เมือง" and province_name:
        city_raw = "เมือง" + province_name

    for candidate, status in [
        (city_raw, "city_exact"),
        (area_raw, "area_exact"),
        (parts["district"], "address_district_exact"),
    ]:
        d = exact_match(districts, candidate)
        if d is not None:
            return d, status

    sub_candidates = [area_raw, city_raw, parts["subdistrict"]]
    for candidate in sub_candidates:
        sub = exact_match(subdistricts, candidate)
        if sub is None:
            continue
        matched = gprov[gprov["skey"] == sub["key"]]
        if matched.empty:
            continue
        dgroups = list(matched.groupby(["dkey", "districtNameTh"], sort=False))
        if len(dgroups) == 1:
            (dkey, dth), g = dgroups[0]
            return {
                "key": dkey, "th": dth, "en": g.iloc[0]["districtNameEn"]
            }, "from_subdistrict_exact"

    for candidate, prefix in [(city_raw, "city"), (area_raw, "area")]:
        if not candidate:
            continue
        d, score, _ = fuzzy_match(districts, [candidate])
        if d is not None:
            addr_norm = norm_thai(address)
            confirmed = d["key"] in addr_norm or similarity(d["th"], parts["district"]) >= 0.95
            if score >= AUTO_ACCEPT_FUZZY_THRESHOLD or confirmed:
                return d, f"{prefix}_fuzzy_confirmed_{score:.2f}"
            return d, f"{prefix}_fuzzy_check_{score:.2f}"

    if parts["district"]:
        d, score, _ = fuzzy_match(districts, [parts["district"]])
        if d is not None:
            return d, f"address_district_fuzzy_auto_{score:.2f}" if score >= AUTO_ACCEPT_FUZZY_THRESHOLD else f"address_district_fuzzy_check_{score:.2f}"

    sub, score, _ = fuzzy_match(subdistricts, sub_candidates)
    if sub is not None:
        matched = gprov[gprov["skey"] == sub["key"]]
        if not matched.empty:
            dgroups = list(matched.groupby(["dkey", "districtNameTh"], sort=False))
            if len(dgroups) == 1:
                (dkey, dth), g = dgroups[0]
                return {
                    "key": dkey, "th": dth, "en": g.iloc[0]["districtNameEn"]
                }, f"from_subdistrict_fuzzy_{score:.2f}"

    return None, "district_not_found"


def resolve_subdistrict(gd, area_raw, city_raw, address):
    subs = build_subdistricts(gd)
    parts = extract_address_components(address)
    candidates = [area_raw, city_raw, parts["subdistrict"]]

    for candidate in candidates:
        sub = exact_match(subs, candidate)
        if sub is not None:
            return sub, "exact"

    sub, score, _ = fuzzy_match(subs, candidates)
    if sub is not None:
        confirmed = sub["key"] in norm_thai(address)
        if score >= AUTO_ACCEPT_FUZZY_THRESHOLD or confirmed:
            return sub, f"fuzzy_confirmed_{score:.2f}"
        return sub, f"fuzzy_check_{score:.2f}"

    return None, "not_found"


def resolve_postal(gd, sub, address, wms_postal=""):
    address_postal = extract_postal_code(address)

    # Priority: WMS explicit zip > zip in address > subdistrict > district.
    if wms_postal:
        return wms_postal, "WMS_EXPLICIT"

    if address_postal:
        return address_postal, "ADDRESS_EXPLICIT"

    if sub is not None:
        vals = sorted({
            str(x).strip()
            for x in gd.loc[gd["skey"] == sub["key"], "postalCode"].dropna()
            if re.fullmatch(r"\d{5}", str(x).strip())
        })
        if len(vals) == 1:
            return vals[0], "EXACT_SUBDISTRICT"

    vals = sorted({
        str(x).strip()
        for x in gd["postalCode"].dropna()
        if re.fullmatch(r"\d{5}", str(x).strip())
    })
    if len(vals) == 1:
        return vals[0], "DISTRICT_UNIQUE"
    if len(vals) > 1:
        return "", "DISTRICT_HAS_MULTIPLE_POSTAL_CODES"
    return "", "NO_POSTAL"


def resolve_row(row, geo, province_master, district_master):
    prov_raw = clean_location_name(row.get("Receipt Province", ""))
    city_raw = clean_location_name(row.get("Receipt City", ""))
    area_raw = clean_location_name(row.get("Receipt Area", ""))
    address = clean_text(row.get("Consignee Addr", ""))
    wms_postal = get_wms_postal(row)

    province, pstatus = resolve_province(geo, prov_raw, city_raw, area_raw, address)

    if province is None:
        postal = wms_postal or extract_postal_code(address)
        return {
            "จังหวัด": "", "เขต/อำเภอ": "", "รหัสไปรษณีย์": postal,
            "สถานะตรวจสอบ": "PROVINCE_NOT_FOUND",
            "จังหวัดสถานะ": pstatus, "อำเภอสถานะ": "",
            "ตำบลสถานะ": "", "รหัสไปรษณีย์สถานะ": "WMS_EXPLICIT" if wms_postal else "ADDRESS_EXPLICIT" if postal else "NO_POSTAL",
        }

    pkey = province["key"]
    gprov = geo[geo["pkey"] == pkey].copy()

    district, dstatus = resolve_district(gprov, city_raw, area_raw, address)

    if district is None:
        province_out = make_bilingual(province.get("th", ""), province.get("en", ""))
        if pkey in province_master.index:
            pm = province_master.loc[pkey]
            province_out = make_bilingual(pm["province_th"], pm["province_en"])
        postal = wms_postal or extract_postal_code(address)
        return {
            "จังหวัด": province_out, "เขต/อำเภอ": "", "รหัสไปรษณีย์": postal,
            "สถานะตรวจสอบ": "DISTRICT_NOT_FOUND",
            "จังหวัดสถานะ": pstatus, "อำเภอสถานะ": dstatus,
            "ตำบลสถานะ": "", "รหัสไปรษณีย์สถานะ": "WMS_EXPLICIT" if wms_postal else "ADDRESS_EXPLICIT" if postal else "NO_POSTAL",
        }

    dkey = district["key"]
    gd = gprov[gprov["dkey"] == dkey].copy()

    sub, sstatus = resolve_subdistrict(gd, area_raw, city_raw, address)
    postal, postal_status = resolve_postal(gd, sub, address, wms_postal)

    province_out = make_bilingual(province["th"], province["en"])
    if pkey in province_master.index:
        pm = province_master.loc[pkey]
        province_out = make_bilingual(pm["province_th"], pm["province_en"])

    district_out = make_bilingual(district["th"], district["en"])
    if dkey in district_master.index:
        dm = district_master.loc[dkey]
        district_out = make_bilingual(dm["district_th"], dm["district_en"])

    unresolved = (
        pstatus.endswith("_check")
        or "_fuzzy_check_" in pstatus
        or dstatus.endswith("_check")
        or "_fuzzy_check_" in dstatus
        or sstatus.startswith("fuzzy_check")
    )

    status = "FUZZY_MATCH_CHECK" if unresolved else (
        "OK" if postal_status in {
            "WMS_EXPLICIT", "ADDRESS_EXPLICIT",
            "EXACT_SUBDISTRICT", "DISTRICT_UNIQUE"
        } else postal_status
    )

    return {
        "จังหวัด": province_out,
        "เขต/อำเภอ": district_out,
        "รหัสไปรษณีย์": postal,
        "สถานะตรวจสอบ": status,
        "จังหวัดสถานะ": pstatus,
        "อำเภอสถานะ": dstatus,
        "ตำบลสถานะ": sstatus,
        "รหัสไปรษณีย์สถานะ": postal_status,
    }



# =========================================================
# KOL Excel output
# =========================================================

from copy import copy
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter


DEFAULT_KOL_TEMPLATE = "ไฟล์ทำ KOL ระบบใหม่ update app(1).xlsx"


def find_kol_template(uploaded=None):
    """Use uploaded template when provided; otherwise use the template next to app.py."""
    if uploaded is not None:
        return uploaded
    path = Path(__file__).parent / DEFAULT_KOL_TEMPLATE
    if path.exists():
        return path
    raise FileNotFoundError(
        f"ไม่พบไฟล์ต้นแบบ {DEFAULT_KOL_TEMPLATE} ในโฟลเดอร์เดียวกับ app.py"
    )


def _clear_values(ws, start_row, end_row, start_col=1, end_col=None):
    end_col = end_col or ws.max_column
    for row in ws.iter_rows(
        min_row=start_row, max_row=end_row,
        min_col=start_col, max_col=end_col
    ):
        for cell in row:
            cell.value = None


def _copy_cell_style(src, dst):
    """Copy cell formatting safely between different workbooks."""
    if src.font:
        dst.font = copy(src.font)
    if src.fill:
        dst.fill = copy(src.fill)
    if src.border:
        dst.border = copy(src.border)
    if src.alignment:
        dst.alignment = copy(src.alignment)
    if src.protection:
        dst.protection = copy(src.protection)
    dst.number_format = src.number_format


def _copy_row_style(ws, source_row, target_row, max_col=None):
    """Copy formatting only so newly added rows look like the template."""
    max_col = max_col or ws.max_column
    ws.row_dimensions[target_row].height = ws.row_dimensions[source_row].height
    for col in range(1, max_col + 1):
        src = ws.cell(source_row, col)
        dst = ws.cell(target_row, col)
        if src.has_style:
            dst._style = copy(src._style)
        if src.number_format:
            dst.number_format = src.number_format
        if src.alignment:
            dst.alignment = copy(src.alignment)
        if src.protection:
            dst.protection = copy(src.protection)


def _extend_formula_rows(ws, source_row, first_new_row, last_new_row):
    """Copy formula/style row and translate relative references for extra rows."""
    from openpyxl.formula.translate import Translator

    if last_new_row < first_new_row:
        return

    for r in range(first_new_row, last_new_row + 1):
        _copy_row_style(ws, source_row, r)
        for c in range(1, ws.max_column + 1):
            src = ws.cell(source_row, c)
            dst = ws.cell(r, c)
            if isinstance(src.value, str) and src.value.startswith("="):
                try:
                    dst.value = Translator(
                        src.value, origin=src.coordinate
                    ).translate_formula(dst.coordinate)
                except Exception:
                    dst.value = src.value


def write_order_data_to_template(template_source, final_df):
    """
    Create the NEW KOL workbook:
    - Keep the original workbook structure/styles/formulas.
    - Replace only the data area of 'ข้อมูลออเดอร์' with the converted WMS data.
    - Keep all other template sheets and formulas.
    """
    wb = load_workbook(template_source)

    if "ข้อมูลออเดอร์" not in wb.sheetnames:
        raise ValueError("ไฟล์ต้นแบบไม่มีชีท 'ข้อมูลออเดอร์'")
    ws = wb["ข้อมูลออเดอร์"]

    headers = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]
    missing = [h for h in final_df.columns if h not in headers]
    if missing:
        raise ValueError(
            "คอลัมน์จากไฟล์ WMS ไม่ตรงกับไฟล์ต้นแบบ: " + ", ".join(map(str, missing))
        )

    # Remove old order values but keep the template's formatting.
    if ws.max_row >= 2:
        _clear_values(ws, 2, ws.max_row, 1, ws.max_column)

    data_rows = len(final_df)
    needed_last_row = data_rows + 1

    # Copy the template row style to any rows beyond the original template size.
    style_source = min(max(2, ws.max_row), 2)
    if needed_last_row > ws.max_row:
        _copy_row_style(ws, style_source, ws.max_row + 1, ws.max_column)
        for r in range(ws.max_row + 1, needed_last_row + 1):
            _copy_row_style(ws, style_source, r, ws.max_column)

    # Write values by matching header names, not by hard-coded positions.
    header_to_col = {headers[c - 1]: c for c in range(1, len(headers) + 1)}
    for r_idx, (_, row) in enumerate(final_df.iterrows(), start=2):
        for col_name, value in row.items():
            c = header_to_col[col_name]
            cell = ws.cell(r_idx, c)
            if pd.isna(value):
                value = None
            cell.value = value

    # Make Excel recalculate all template formulas when the file is opened.
    try:
        wb.calculation.fullCalcOnLoad = True
        wb.calculation.forceFullCalc = True
        wb.calculation.calcMode = "auto"
    except Exception:
        pass

    return wb


def _build_product_lookup(wb):
    """
    Match the template's Excel VLOOKUP behavior.
    In 'ข้อมูลน้ำหนักสินค้า':
      D = sku
      H = Barcode
      L = single-item weight

    Excel VLOOKUP returns the FIRST matching SKU, so do not overwrite
    an existing SKU when duplicates are found.
    """
    result = {}
    if "ข้อมูลน้ำหนักสินค้า" not in wb.sheetnames:
        return result

    ws = wb["ข้อมูลน้ำหนักสินค้า"]
    for r in range(2, ws.max_row + 1):
        sku = clean_text(ws.cell(r, 4).value)
        if not sku or sku in result:
            continue

        barcode = ws.cell(r, 8).value
        weight = ws.cell(r, 12).value

        try:
            weight_value = float(weight) if weight not in (None, "") else 0.0
        except Exception:
            weight_value = 0.0

        result[sku] = {
            "barcode": "" if barcode is None else str(barcode).strip(),
            "weight": weight_value,
        }
    return result


def _raw_value(x):
    """Keep receiver-name text as supplied by WMS, including intentional spaces/NBSP."""
    if x is None or pd.isna(x):
        return ""
    return str(x)


def _derive_receiver_name(address):
    """
    Fallback used only when WMS has no receiver name.
    For addresses beginning with a village name, keep the village name before
    the house-number section. This covers the observed KOL/Eastlync pattern
    such as 'หมู่บ้านพฤษชาติวิลล่า 2 บ้านเลขที่ ...'.
    """
    text = _raw_value(address)
    if not text.strip():
        return ""

    m = re.match(r"^(.*?)\s+(?:บ้านเลขที่|เลขที่)\s*", text.strip())
    if m:
        name = m.group(1).strip()
        # Remove a trailing standalone house/village number when it is
        # immediately before the house-number phrase.
        name = re.sub(r"\s+\d+$", "", name).strip()
        return name

    return ""


def _build_old_kol_values(final_df, template_wb):
    """
    Convert the OLD KOL upload formulas to values while matching the
    template's Excel VLOOKUP behavior.

    Important:
    - SKU weight and barcode come from the FIRST matching SKU in
      'ข้อมูลน้ำหนักสินค้า', exactly like Excel VLOOKUP.
    - Package weight = SUM(line weights for the same address) + 0.2.
    - placedAt/payTime use the file-generation timestamp, matching NOW().
    - Receiver name is preserved from WMS without clean_text() changes.
    """
    product_lookup = _build_product_lookup(template_wb)
    now_value = datetime.now().replace(microsecond=0)

    rows = []
    helper_weights = []

    for _, row in final_df.iterrows():
        erp = clean_text(row.get("ERP No.", ""))
        product_code = clean_text(row.get("Product Code", ""))

        product_info = product_lookup.get(
            product_code, {"barcode": "", "weight": 0.0}
        )

        # The old KOL formula uses the template/master barcode, not WMS
        # Product Barcode. Fall back to WMS only when the master has none.
        barcode = product_info["barcode"] or clean_text(
            row.get("Product Barcode", "")
        )

        qty_raw = row.get("Allocated Qty", "")
        try:
            qty = float(qty_raw) if clean_text(qty_raw) else 0
            if qty.is_integer():
                qty = int(qty)
        except Exception:
            qty = 0

        raw_address = row.get("Consignee Addr", "")
        address = _raw_value(raw_address)

        product_weight = product_info["weight"]
        line_weight = round(product_weight * qty, 3)
        helper_weights.append((address, line_weight))

        comment = clean_text(row.get("Comment by seller", ""))
        if comment == "Express delivery":
            special_mark = "NORMAL"
        elif comment == "Lalamove":
            special_mark = "B2B_OFFLINE"
        else:
            special_mark = ""

        receiver_name = _raw_value(row.get("Receive Name", ""))
        if not receiver_name.strip():
            receiver_name = _derive_receiver_name(address)

        out = [
            "EastlyncOffline" if erp else "",
            "OFFLINE" if erp else "",
            "SALE" if erp else "",
            erp,
            special_mark,
            "EXPRESS" if erp else "",
            product_code,
            barcode,
            product_code,
            qty if erp else "",
            0 if erp else "",
            0 if erp else "",
            now_value if erp else "",
            now_value if erp else "",
            receiver_name,
            clean_mobile_number(row.get("Receiver Mobile Number", "")) if erp else "",
            clean_mobile_number(row.get("Receiver Mobile Number", "")) if erp else "",
            "TH" if erp else "",
            clean_text(row.get("จังหวัด", "")),
            clean_text(row.get("เขต/อำเภอ", "")),
            clean_text(row.get("รหัสไปรษณีย์", "")),
            clean_text(row.get("รหัสไปรษณีย์", "")),
            address,
            "COD" if erp else "",
            None,
            None,
            None,
            None,
            None,
            None,
            line_weight if erp else "",
        ]
        rows.append(out)

    address_totals = {}
    for address, weight in helper_weights:
        address_totals[address] = address_totals.get(address, 0.0) + weight

    for row in rows:
        address = row[22]
        if row[3]:
            row[26] = round(address_totals.get(address, 0.0) + 0.2, 3)
        else:
            row[26] = ""

    return rows



def create_old_kol_file(template_wb, final_df):
    """Create a standalone OLD KOL upload workbook with values only (no formulas).

    The upload file contains the 30 real upload columns A:AD.
    Column AE is only a hidden calculation helper in the template and is
    intentionally not exported to the old KOL upload file.
    """
    src_ws = template_wb["ไฟล์อัพโหลด KOL ระบบเก่า"]
    upload_max_col = 30  # A:AD; AE is an internal weight helper only.

    from openpyxl import Workbook
    out_wb = Workbook()
    out_ws = out_wb.active
    out_ws.title = "ไฟล์อัพโหลด KOL ระบบเก่า"

    # Copy widths / formatting for upload columns only.
    for c in range(1, upload_max_col + 1):
        letter = openpyxl.utils.get_column_letter(c)
        dim = src_ws.column_dimensions[letter]
        out_ws.column_dimensions[letter].width = dim.width
        out_ws.column_dimensions[letter].hidden = dim.hidden

    # Copy header + instruction rows.
    for r in (1, 2):
        for c in range(1, upload_max_col + 1):
            s = src_ws.cell(r, c)
            d = out_ws.cell(r, c)
            d.value = s.value
            _copy_cell_style(s, d)

    values = _build_old_kol_values(final_df, template_wb)

    for i, row_values in enumerate(values, start=3):
        for c, value in enumerate(row_values[:upload_max_col], start=1):
            cell = out_ws.cell(i, c)
            cell.value = value

    # Apply template data-row style to all generated rows.
    # M/N (placedAt/payTime) must use exactly the same font and date format
    # as the template data row, not the default workbook font.
    style_row = 3
    for r in range(3, 3 + len(values)):
        if r != style_row:
            for c in range(1, upload_max_col + 1):
                s = src_ws.cell(style_row, c)
                d = out_ws.cell(r, c)
                _copy_cell_style(s, d)

    out_ws.freeze_panes = src_ws.freeze_panes
    out_ws.auto_filter.ref = f"A1:AD{max(2, 2 + len(values))}"

    try:
        out_wb.calculation.calcMode = "manual"
    except Exception:
        pass

    return out_wb



def workbook_bytes(wb):
    buffer = io.BytesIO()
    wb.save(buffer)
    return buffer.getvalue()


# =========================================================
# UI
# =========================================================

st.title("📦 WMS → KOL Excel Generator")
st.caption("แปลงข้อมูล WMS → ใส่ลงไฟล์ต้นแบบ KOL ระบบใหม่ → สร้างไฟล์ KOL ระบบเก่าแบบไม่มีสูตร Excel")

with st.sidebar:
    st.header("ตั้งค่า")

    master_file = st.file_uploader(
        "Address master (ไม่บังคับ)", type=["xlsx"],
        help="ค่าเริ่มต้นใช้ adress.xlsx ในโฟลเดอร์ data"
    )

    template_file = st.file_uploader(
        "ไฟล์ต้นแบบ KOL ระบบใหม่ (ไม่บังคับ)",
        type=["xlsx"],
        help=f"ถ้าไม่อัปโหลด ระบบจะใช้ {DEFAULT_KOL_TEMPLATE} ที่อยู่ข้าง app.py"
    )

    st.info(
        "ไฟล์ที่ 1: ใช้ไฟล์ต้นแบบ KOL ระบบใหม่ โดยวางข้อมูล WMS ที่แปลงแล้วลงในชีท "
        "'ข้อมูลออเดอร์' และคงสูตร/รูปแบบของไฟล์ต้นแบบไว้\n\n"
        "ไฟล์ที่ 2: สร้างชีท 'ไฟล์อัพโหลด KOL ระบบเก่า' แยกออกมา โดยแปลงสูตรเป็นค่าข้อมูลจริง"
    )

uploaded = st.file_uploader(
    "1) อัปโหลดไฟล์ Excel จาก WMS",
    type=["xlsx", "xls"]
)

if uploaded is None:
    st.markdown("""
### วิธีใช้งาน
1. อัปโหลดไฟล์ **Outbound Detail Export จาก WMS**
2. ระบบตรวจสอบ `Receipt Province`, `Receipt City`, `Receipt Area`
3. ตรวจสอบ `Consignee Addr`
4. ตรวจสอบ `Receiver Zipcode`
5. สร้าง `จังหวัด`, `เขต/อำเภอ`, `รหัสไปรษณีย์`
6. นำข้อมูลที่ได้ไปใส่ในชีท **ข้อมูลออเดอร์** ของไฟล์ต้นแบบโดยอัตโนมัติ
7. สร้างไฟล์ **KOL ระบบใหม่** 1 ไฟล์ และ **KOL ระบบเก่าแบบไม่มีสูตร** อีก 1 ไฟล์
""")
    st.stop()

try:
    raw = pd.read_excel(uploaded, dtype=str).fillna("")
except Exception as e:
    st.error(f"อ่านไฟล์ไม่สำเร็จ: {e}")
    st.stop()

required = ["Receipt Province", "Receipt City", "Receipt Area", "Consignee Addr"]
missing = [c for c in required if c not in raw.columns]
if missing:
    st.error("ไม่พบคอลัมน์ที่จำเป็น: " + ", ".join(missing))
    st.stop()

if raw.empty:
    st.warning("⚠️ ไฟล์ Excel ไม่มีข้อมูล Order")
    st.stop()

try:
    master, province_master, district_master = prepare_master(load_address_master(master_file))
except Exception as e:
    st.error(f"อ่าน address master ไม่สำเร็จ: {e}")
    st.stop()

try:
    geo = build_geo_tables(load_postal_data())
except Exception as e:
    st.error(f"ไม่สามารถโหลดฐานข้อมูลรหัสไปรษณีย์ได้ในขณะนี้ ({e})")
    st.stop()

with st.spinner("กำลังประมวลผลข้อมูล WMS..."):
    results = [
        resolve_row(row, geo, province_master, district_master)
        for _, row in raw.iterrows()
    ]
    result_df = pd.DataFrame(results)

insert_at = list(raw.columns).index("Receipt Area") + 1
final = pd.concat(
    [
        raw.iloc[:, :insert_at].copy(),
        result_df[["จังหวัด", "เขต/อำเภอ", "รหัสไปรษณีย์"]],
        raw.iloc[:, insert_at:].copy(),
    ],
    axis=1,
)

if "Receiver Mobile Number" in final.columns:
    final["Receiver Mobile Number"] = final["Receiver Mobile Number"].apply(clean_mobile_number)

final["รหัสไปรษณีย์"] = (
    final["รหัสไปรษณีย์"].fillna("").astype(str)
    .str.extract(r"(\d{5})", expand=False).fillna("")
)

st.success(f"ประมวลผลแล้ว {len(final):,} แถว")
c1, c2, c3, c4 = st.columns(4)
c1.metric("ทั้งหมด", f"{len(final):,}")
c2.metric("สำเร็จ", f"{(result_df['สถานะตรวจสอบ'] == 'OK').sum():,}")
c3.metric("ต้องตรวจสอบ", f"{(result_df['สถานะตรวจสอบ'] != 'OK').sum():,}")
c4.metric("รหัสไปรษณีย์ว่าง", f"{result_df['รหัสไปรษณีย์'].fillna('').astype(str).str.strip().eq('').sum():,}")

empty_district = result_df["เขต/อำเภอ"].fillna("").astype(str).str.strip().eq("")
empty_province = result_df["จังหวัด"].fillna("").astype(str).str.strip().eq("")
empty_postal = result_df["รหัสไปรษณีย์"].fillna("").astype(str).str.strip().eq("")
issues = result_df[result_df["สถานะตรวจสอบ"] != "OK"]

issue_indices = sorted(
    set(issues.index)
    | set(result_df.index[empty_district])
    | set(result_df.index[empty_province])
    | set(result_df.index[empty_postal])
)

if issue_indices:
    with st.expander("⚠️ รายการที่ควรตรวจสอบก่อนใช้งาน", expanded=True):
        preview_cols = [
            "จังหวัด", "เขต/อำเภอ", "รหัสไปรษณีย์",
            "สถานะตรวจสอบ", "จังหวัดสถานะ", "อำเภอสถานะ",
            "ตำบลสถานะ", "รหัสไปรษณีย์สถานะ",
        ]
        left_cols = ["Receipt Province", "Receipt City", "Receipt Area", "Consignee Addr"]
        if "Receiver Zipcode" in raw.columns:
            left_cols.append("Receiver Zipcode")

        issue_preview = pd.concat(
            [
                raw.loc[issue_indices, left_cols].reset_index(drop=True),
                result_df.loc[issue_indices, preview_cols].reset_index(drop=True),
            ],
            axis=1,
        )
        st.dataframe(issue_preview, use_container_width=True, hide_index=True)

st.subheader("ตัวอย่างข้อมูลที่แปลงแล้ว")
show_cols = [
    "ERP No.", "Receipt Province", "Receipt City", "Receipt Area",
    "จังหวัด", "เขต/อำเภอ", "รหัสไปรษณีย์",
    "Product Code", "Product Barcode", "Allocated Qty",
]
st.dataframe(
    final[[c for c in show_cols if c in final.columns]].head(100),
    use_container_width=True,
    hide_index=True,
)

try:
    template_source = find_kol_template(template_file)

    with st.spinner("กำลังสร้างไฟล์ KOL ระบบใหม่และ KOL ระบบเก่า..."):
        # Load once for both outputs.
        template_wb = load_workbook(template_source)

        # File 1: full template + converted WMS data in 'ข้อมูลออเดอร์'.
        new_kol_wb = write_order_data_to_template(template_source, final)
        new_kol_bytes = workbook_bytes(new_kol_wb)

        # File 2: standalone old KOL sheet, values only.
        old_kol_wb = create_old_kol_file(template_wb, final)
        old_kol_bytes = workbook_bytes(old_kol_wb)

    st.success("สร้างไฟล์เรียบร้อยแล้ว")

    col1, col2 = st.columns(2)

    with col1:
        st.download_button(
            "⬇️ ดาวน์โหลดไฟล์ 1 — KOL ระบบใหม่",
            data=new_kol_bytes,
            file_name="KOL_New_WMS_Updated.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True,
        )
        st.caption(
            "ใช้ไฟล์ต้นแบบเดิม + ใส่ข้อมูล WMS ลงชีท 'ข้อมูลออเดอร์' "
            "โดยคงสูตร Excel และรูปแบบของต้นแบบ"
        )

    with col2:
        st.download_button(
            "⬇️ ดาวน์โหลดไฟล์ 2 — KOL ระบบเก่า (ไม่มีสูตร)",
            data=old_kol_bytes,
            file_name="KOL_Old_Upload_No_Formulas.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="secondary",
            use_container_width=True,
        )
        st.caption(
            "แยกเฉพาะชีท 'ไฟล์อัพโหลด KOL ระบบเก่า' และแปลงสูตรเป็นค่าข้อมูลจริง "
            "ไม่มีสูตร Excel"
        )

except Exception as e:
    st.error(f"สร้างไฟล์ KOL ไม่สำเร็จ: {e}")
    st.exception(e)
