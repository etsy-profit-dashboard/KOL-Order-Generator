import io
import base64
from copy import copy
import re
import unicodedata
from pathlib import Path
from difflib import SequenceMatcher
from datetime import datetime

import requests
import pandas as pd
import streamlit as st

try:
    import openpyxl
    from openpyxl import load_workbook
    from openpyxl.styles import Font, Alignment
except Exception:
    openpyxl = None

st.set_page_config(
    page_title="WMS → KOL Upload Converter",
    page_icon="📦",
    layout="wide",
)

POSTAL_DATA_URL = (
    "https://raw.githubusercontent.com/"
    "thailand-geography-data/thailand-geography-json/main/src/geography.json"
)
FUZZY_THRESHOLD = 0.86

# =========================================================
# TEXT / NUMBER HELPERS
# =========================================================

def clean_text(x):
    if x is None:
        return ""
    try:
        if pd.isna(x):
            return ""
    except Exception:
        pass
    s = unicodedata.normalize("NFC", str(x))
    s = s.replace("\xa0", " ").replace("\n", " ").replace("\r", " ").replace("\t", " ")
    s = re.sub(r"\s+", " ", s)
    return s.strip()


def norm_thai(x):
    s = clean_text(x).lower()
    s = unicodedata.normalize("NFC", s)
    replacements = {
        "เเ": "แ",
        "ํา": "ำ",
        "ํ": "",
        "กรุงเทพฯ": "กรุงเทพมหานคร",
        "กรุงเทพเทพมหานคร": "กรุงเทพมหานคร",
        "จังหวัดกรุงเทพเทพมหานคร": "กรุงเทพมหานคร",
        "กรุงเทพ": "กรุงเทพมหานคร",
        "กทม.": "กรุงเทพมหานคร",
        "กทม": "กรุงเทพมหานคร",
        "อยุธยา": "พระนครศรีอยุธยา",
        "จังหวัดอยุธยา": "พระนครศรีอยุธยา",
        "ลําปาง": "ลำปาง",
        "ลาปาง": "ลำปาง",
    }
    for old, new in replacements.items():
        s = s.replace(old, new)
    for prefix in ["จังหวัด", "อำเภอ", "เขต", "แขวง", "ตำบล", "ตําบล", "จ.", "อ.", "ต."]:
        while s.startswith(prefix):
            s = s[len(prefix):]
    return s.replace(" ", "")


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


def clean_location_name(x):
    s = clean_text(x)
    if not s:
        return ""
    # Do not remove a prefix buried inside a composite string such as
    # 'แขวงลาดยาว เขตจตุจักร'; composite parsing is handled separately.
    s = re.sub(r"^(จังหวัด|จ\.|อำเภอ|อ\.|เขต|แขวง|ตำบล|ต\.|ตําบล)\s*", "", s, flags=re.I)
    s = s.replace(".", "").strip()
    mapping = {
        "กรุงเทพ": "กรุงเทพมหานคร",
        "กทม": "กรุงเทพมหานคร",
        "กรุงเทพมหานคร": "กรุงเทพมหานคร",
        "กรุงเทพเทพมหานคร": "กรุงเทพมหานคร",
        "Bangkok": "กรุงเทพมหานคร",
        "อยุธยา": "พระนครศรีอยุธยา",
        "จังหวัดอยุธยา": "พระนครศรีอยุธยา",
        "ลําปาง": "ลำปาง",
        "ลาปาง": "ลำปาง",
    }
    return mapping.get(s, s)


def clean_mobile_number(x):
    s = clean_text(x)
    if not s:
        return ""
    if re.fullmatch(r"\d+\.0", s):
        s = s[:-2]
    s = re.sub(r"\D", "", s)
    if s.startswith("66") and len(s) >= 11:
        s = "0" + s[2:]
    elif s.startswith("00") and len(s) == 11:
        # Some WMS exports can already contain an extra leading 0.
        s = s[1:]
    elif len(s) == 9:
        s = "0" + s
    # Final safety rule: Thai mobile numbers must be 10 digits in the KOL file.
    if len(s) == 11 and s.startswith("00"):
        s = s[1:]
    return s


def clean_barcode(x):
    s = clean_text(x)
    if not s:
        return ""
    try:
        if re.fullmatch(r"[+-]?\d+(?:\.0+)?", s):
            return str(int(float(s)))
        if re.fullmatch(r"[+-]?\d+(?:\.\d+)?[eE][+-]?\d+", s):
            return format(float(s), ".0f")
    except Exception:
        pass
    return s


def to_number(x, default=0):
    if x is None:
        return default
    try:
        if pd.isna(x):
            return default
    except Exception:
        pass
    s = clean_text(x).replace(",", "")
    if not s:
        return default
    try:
        f = float(s)
        return int(f) if f.is_integer() else f
    except Exception:
        return default


def valid_postcode(x):
    s = clean_text(x)
    if re.fullmatch(r"\d{5}", s):
        return s
    m = re.findall(r"(?<!\d)(\d{5})(?!\d)", s)
    return m[-1] if m else ""

# =========================================================
# ADDRESS PARSING
# =========================================================

def extract_address_components(address):
    address = clean_text(address)
    out = {"province": "", "district": "", "subdistrict": "", "postal": ""}
    if not address:
        return out

    posts = re.findall(r"(?<!\d)(\d{5})(?!\d)", address)
    if posts:
        out["postal"] = posts[-1]

    # Thai administrative names are normally a single token. Keep the parser
    # simple so 'เขตคลองสามวา กรุงเทพ 10510' is parsed correctly.
    patterns = {
        "province": r"(?:จังหวัด|จ\.)\s*([ก-๙A-Za-z]+)",
        "district": r"(?:อำเภอ|อ\.|เขต)\s*([ก-๙A-Za-z]+)",
        "subdistrict": r"(?:ตำบล|ต\.|ตําบล|แขวง)\s*([ก-๙A-Za-z]+)",
    }
    for key, pattern in patterns.items():
        matches = re.findall(pattern, address, flags=re.I)
        vals = [clean_text(v) for v in matches if clean_text(v)]
        if vals:
            out[key] = vals[-1]

    return out


def composite_admin_values(text):
    """Extract district/subdistrict names from a field like 'แขวงลาดยาว เขตจตุจักร'."""
    s = clean_text(text)
    if not s:
        return {"district": [], "subdistrict": []}
    districts = [clean_text(v) for v in re.findall(r"(?:เขต|อำเภอ|อ\.)\s*([ก-๙A-Za-z]+)", s, flags=re.I)]
    subs = [clean_text(v) for v in re.findall(r"(?:แขวง|ตำบล|ต\.|ตําบล)\s*([ก-๙A-Za-z]+)", s, flags=re.I)]
    return {"district": districts, "subdistrict": subs}


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


def fuzzy_match(items, value):
    if not value:
        return None, 0.0
    best, best_score = None, 0.0
    for item in items:
        score = similarity(item.get("key", ""), value)
        if score > best_score:
            best, best_score = item, score
    if best is not None and best_score >= FUZZY_THRESHOLD:
        return best, best_score
    return None, 0.0


# =========================================================
# EMBEDDED SUPPORT DATA
# =========================================================
# This version is self-contained: address master and weight master
# are embedded in app.py, so the user only needs to upload the WMS file.
import base64
import zlib

EMBEDDED_ADDRESS_B64 = (
    "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"
)

EMBEDDED_WEIGHT_B64 = (
    "eNrtfctuZMmV2L6/ItFAD+BUVireD++yWFlV6Upm0kmyUqUdVUV1E6omGyy2pN4bsL0YeDXwxoa9s+GFPQZmMTAG8M+M5PFqfsHnRNxHRNy4eW+Q2TPydFWL"
    "FJm8N+KciBPnFedx/uZyslmcLmeXt7dX315/+OcT2v7IZp9+8/3s5YvF5Oz+7sP37x8mG/jDz5eb2Zf44eZu/sXl7Yfr+8ny4fr+9urj5Ie77x+++XL25d//"
    "zb/ZXP/u7//m31aPTernvvvmbnJ/dXP7w5ez51f37+8+XM/+8Od/8bd//d/+8Jf/6w//46/hvfffwmuzP/zlv/s///1/Jh/+8S/+Kv7wn1Uv/99//efw+5tX"
    "7ZvhJ/618JO//d//8e/+6r/+8T/8J/jkh+ur+/apv/sv/wo+Ozs5x4/aJVHBz/qL8zerzcXq5erkGSN0ZiihVBhNjdIzMidq1v4d/ix/MTlZ7hanqxfLyXq7"
    "n5y9npysl4vN+XI3o+QZe6YUIYRRLWT1u4HfiSCczbSUQlAqlSRUzKiayqmcMTKXU6qnAJTgU87hN44TUzabUT63asZn8A1+DeEIgOa2gZloGIXMo0eDXybi"
    "xeT1u8X6crfdwG+Lk9WLyfPFbrda7ibL8/Pl5mQ5udhuEkwo0TP4JmwAvxUAkZgKABa+ZhQwsFMqZ4JMOfyPIfxqxhAJqucyiwMLcCCyHVzBDx7/4O98Bv/a"
    "R4wyOPqUzmFaNuczaqYAD+UzbqacTpnBrYPJEQw7NwYAEABWNGS48Uw2i8ikoekiimARKflqslnB4m08FTzfrV69vlhuVptXE6CCy9No9YTUAR1QSjUN1pED"
    "oBQWbG7wC0mCsimlsKRcTLlCWpC4FoAIA2zcqmaWUpTQsAxwCam5poTT7er84nK3nLxariNUSIOKht+5IbBP8a4hwPgFVC0cGoCOo2oyZdLvKXwIpM/mJoeH"
    "DPAQBGadCWM1FVpZyrVxKxE8rkJMvposXi8mz+HrDL6Wv3i5Xa8WF7gpXZIWQJEpSSsgD1j8hqTxByRq7ogaSMrWCADtwfHNgK/CUyl0CDwnKfA6AJ59NTlf"
    "rFcn79b1sVzAX+CnzTJHUjomKQljNXgAB0BCmcLhBHJCHHhFUHLK9JSpEeQUskQuaYgH40hPJMLEBJicXr7ab3cXHvST9eLd5HRx/ibhJzZdfMUECYmHusMM"
    "E8MhwH1gjiEAQnPNcGdYlnpMCUOxAdBckcnJ7t35xWKN4J4vXsExWK9eOupZvgNEdstFuAewDJbHx1qpYA9grZAjOV4OS89gC6QjpxmuHpDP3AAedA4kB8MI"
    "Acclgm08Z6IkwGONjMgTzOT8cnMOYC83k/OzlxNJfjY5W/wM/oVYUAtYtMKDwIHiczitczY1yI0ksCU4qsCRcFIQTQLoB35l8JlnK2SuYd9Y7jRQEmABXI/Q"
    "TTOVUEYk54FGfOlid/ny5Rp40mp7CkcAPvd7gDypZlCrXyanmlsdYCMUcKepRpLCDcD1NzOhQThNqfHkANArAB/PgDAp9CEvolyQFnoGu4/oaxu+QEMZwX/R"
    "HuNn7izkAcdFp5LhkXUnomHcQoGqAUtdHQrUEryEg5MLHBV3ws4VHIy5O9BAWXDGpWOsubMBg4ZnmuIJbk81zk/moZAIX46ExPl2e/E6y1SNRcKOz7Xh1kRM"
    "FZWEmqUyL6XdsVYqz1Ipi4ioHdlK1jnTlJHMFsDvy8356u2yYqdn24tw8YHti2TxiQHqhIWc4hce3BmsFSw2B21iypCXOtoHyhEV+4yOL4AxWjujLFzoxb+8"
    "XFSH+MVitX6XHGIu/RkO9UopNU/BtwY+4k65rBYcltY2MozgqlvQRrMyGCAqpBQ+O1ueXSBxbLZA3OeOPNars0hy+bWWIAYTQpcENTkGPB6/AXlTOK3IcqzT"
    "HmCxBcCNR4Fpv/KgZsLHGtXrOQeFb8aA5HWMAy+SYPA8PYgE6EbP303OVps3oQIkGAsFMej0vFW9GPxXitgIvELFn8Hhajiqlsx0DwQImIN4nWx3IPN22/Nl"
    "ZKyo2HgBxEiAGBeFiPExOxaf9GA+w7kzJ3T8PB/A7PVyt3s32S1fJJiJCDOuwy0T9MfAjCeCkAWCEJhCKghhH2eBHLPKzvAIO80IrQNnKnI1FXC29ZwgHKjM"
    "zTnqtKjooTCLByyTxMzmJFktxBJr1x1sy5RND7a2dlbxfQB4rpxKjTIMv+m5nSkQZZ6nwvo65dptzhx0HGC6FPCAhbQxYKF9IE17rMG87qDBB+ycjmINEkwk"
    "Sl0gcYyFZW8Ua2+nzRhIBV7ZNoeVahpRAXDoVtQT2zm4oG+cIBU0CCI0MzNFmSQrvR6/GafYOzPR6ZeeKlHBBOPXoJYpMpDA4AUGo1d+2rU8W+5eLk9Az7wA"
    "7f7tcv12eQFa88UFqMmXcAq3mwkhm8lmAZoO8Jaz7e5kuV6sNgG9MIlyL9angCIQtylSRogfR1KHZWa0UrkQNQk2PcmvMsD6Rcwx6WW7jBzMjw6dCLIvRG8/"
    "2S92p3nchDEyxA10hOPhtk9wUyFuYOd0caOFW0fbrXu72KzW60W8ccAuAuQkPx5ydFNgv+HzhbtGq13LoCUMjfZMkSOiFe4ZHMiGsTPrWL+JH2eF+8Xa/Vq9"
    "3e7exbsF8ARoaX08tNimTMWCNwr3i1X71UFLGFD0A7SAJI+H1r5IDeGCF+4XD1jj8mKxjvdLRmzR0uMhxjdlLJ8X7havdmuxPt1uXsTbxUOewckRqZCH28WZ"
    "Ap4QGCtGeWkYvSIKN0y0G/Z8uXq1jDcM1McANXpEShTRATOBssO5Zl1GLwo3TFQb1kEKMAmZIYjMQPEAdCqlA1QOdGiF6oY2Pajsy0xKrkMNmFLuvP+IMfxD"
    "lhY9m4hEEol70lULQaW8vv1083Dz2+vJ6d3Np4fv768nr64/hnYBvG6DNRGEmMBOMALlA8VNps5fWfv36Bydvk4ldLxdojrLUZ3l2ZWxqWK+DxRz29FUxOGr"
    "nNfvXuy8x7vrM9bA+wNdVhHW0WUd1B0HcXZLRbTmib/HuXTCh0mw4Od3dw/fXE9OPl5f3X66vk/W3ERrTlnApQglYE7glZPAxQbD2+DdT2VNMK8FM/S4MjJ3"
    "mw5LiHdEVrorqKx3SYSeGjiTLVUq3dGlROix714ATZ5vt+cX3ud9frE6CV0FoDVpG5oWoGKowAIAgSngiAElAfE4HtH84DfD3WLlj5fQRRIr9LLWlpCDHP7/"
    "JeAEbG67iu59mER7svW8MdgJz9kMmnd41cMI+lQlEBRYdMZf/lFkDNbf/SUGqSwxhWREPtf33387Of/+9tP7++vr2wknXyckFHlkCIAcrIeE0znnU/wCax5d"
    "9Gh8ojMVv+HVIUOUvJcMLdo5dU4VnXdwSxI5Y0RorbH0sgpFe2BVaz15sdi9ce7I7N2rIxOQE9E1A9Blu78CiMhZ2N7BiqcVYDeIC/BnCntR4+JvTowzwLMW"
    "oWQRJjqYw9NQ9DCPuDMxcG6m1pEC3jVpvC1oKNmdxJadML+O8XChcAC9v11FqRF6FglwddxbAgKaK94SmBlj6Cq1Nf8mft3I3DqZzrpUrGLVA2zkln1QKzqq"
    "h6LhVc325M3k4jXGEGyBCpYXnm+f7YB7GPIqUjIM153rAdhN5IPtmiu8MUa3erPheMUkm6uyGJIE9Ehr0rYDesj5zt+d7l+vQJOIabjj7NXMyuSa0iT+e4oe"
    "tCmfa/xygii8q3T3fvPm2rgiINFFRyc6YLgRjJAuNmS2u/pwc3X7/nry/O7u08P1fc1Yvru/+iFkIEzQ8AASoE8ZM1gxx9PHcAMoXhJQ94vGSyULjIUTvDMg"
    "/pZGoUcYKNzMjczdOWlSqM1qQ2ar29/eAHe8+/79N5PLjw/3V5Pzh6sfJmtA7Hpydvc7jNMhZOL37GK32JyvL0+Wm4tZgBfjioaCA2MUnA6L7jtHRrLScBxv"
    "cfeaeKtf7w3e6Of4CsAXY8SivaE6gxEdhxGtbztfL9frFBce4IIO2ePgkpwaFu2O5Rlc2DhcWGNhnC/AdkqQCSSYpUdDhsXI8GhjmMggw8chwyevQYy9q+yK"
    "BBcW4ALWz5Fw4Qku4cZwkmFnYIKEVuCLq5uPP0xe3nxETnB2ff/r6/cPk/Pf3NxOLm5uH7w78wrMhKuPgOX9++uPVze3gQgUSuvYQEkAxOliEEW43JyqHIi0"
    "BETagvj26vbm48erCED02h0EMHK+VQsU4ZcoAf4aYzx8rIVv9du7+x9C6PAO5DB0bFN2gwFPAEFHczA6MAfflBlj8IQoWQHRrsCvrm++vo6g42YAOrEpu0Ty"
    "8BfQz36yv7r/dvLbDPHYoaWj+xS4YOk0zW0PKwGOVcB1KYcO0TVLQeObmChi0LgU8bULwTDMWeXOdP+ih7/oEES0p8ngtBokdn644VVncHz4i0TWRxI6MZS4"
    "VAnoXB0AXX2RSN1IYqZGmDlo+vrYx74IUtCZaCdGgQoeBoZUZssUjXzmQkhJGBoC1llGjY2vNQM70smWDgphcMhuCfr2du1tht1ys9yj8E1VWIlGXaJ5o+vU"
    "RWz60K7K7wP2A6jfXmeVXmeFL+NU3Dr6NQYmNLiEDbgtx9jaVByY0Hg8PX8dbUIa8PUK45L/rHZLnJ+97ER/+WgMLTvYaSEau0K48Ec1Y3aKQbW8dkUY5xfK"
    "4MRinAL+CTZXV9UzfGhDeuKlBJEksSxs7ENwYYRNbDXDABgXhYrbo5I41IxJGikUgAkLMXGSIMEE92/ctehu+XK1Tpz+LDIxlAncXmB+m5nzvU5Nj/cVuasS"
    "M9lgkpC9jXQPwCbkgISIHDb72bhb0C42whgVYYN36i02Qj4Zm32CTSBqCM3tDR21NzS/Nyk2oT+JS/ZUbCK1i5lQr+HWBWp1sBm1NzS/N1QklBYYs1zZJ2Oz"
    "j7FhITZUZrBho/aG5ffG6oTSQmy0eio28Q0nUaQVMNzo3N6wUXvD8nvDaLI3gbjnKNyeiM0+wSbYGylIBhs+am94fm8kP4CNJU/FJlbWh24gDHKJcXeYuZ3h"
    "KQ8IrFhBnkxn0R3m0CWEsWLUtoj8tmiSbEuICn0ykSVWytBdBLwwal9Efl90KmlCZNiTaSy6tRy8pghj5s5AZ3n5DnXi+oZldbJ85l3TPnzO/219GoXyahOF"
    "nxEL56QmbDRlmPNUS++pdmpnFc/e5tZUMYAZZdmW3FVYGbn4NdeoFQrv4/dx6KDU4l0Jquoi0gxJdjmtLLhhsOFigvK3xSup/MpVt2umG0AP6p3Lq5MIMV5Q"
    "oV4uxFTJxrlf+8hFxqlsbcm1RBR5Lr8CbXtz8XqJCq1PhAGAEfD9Aq/ZkrwqgyplvO8trQFZ85maqileISMWzLvHMBAVrz998CdDI4NiMkZ6S8FI2S0FC6N2"
    "qnP4cpVNKslaFYwB+OleYISJcNdUzqpAAnJWhcYkBiLrVBLvqe1aS4zQMg8/g+Hwhvl+svj4sbprvrn9evL86uO3AV3DSNqFNoDViYGcCCJGaUjurQQ+4wgT"
    "VbjXouOuZ0SVuetZlCn1dnUB5ttmcjIuYcfH2dowbpIadIhObXDx5q7d8B68Sc8xYAxplllTU+aZZ1HGVAt86AFwhmdo/juKICgIAqBZEBfMXIaC+zHOsxNx"
    "nl0MSJkfntEDtnMIfm/io9aR3mxDDYBQrtrUNYwYJ44/MnRkAJF37M0Md2S0zIHNoqyv/F5kw7RtGOZIDXK1uZkKn+ai0Fx2ogWJyScjT3kLvvPD0K5oYZ1M"
    "r8O+YxYlF4FQXE/OdtsL0AFQ2C/W68kL0ASq253mUJzC3sRcZ6LJabJPhnccGlq2d/4uxbZKVQYeqoAhUcTduoARYoVT6jXm5NLMVTVLspGG3NCMRYH8QCmB"
    "szIamBW5nlkkbYa8QWX+HwzkxGvYilPPpefVKOkNcuuGWWPEts6cTkaKHNUsSoA6u3u4vn2YvL15uPr25hYp+v7m62/gM+Terz7e/c7f00Y6rdZdeWMosG6F"
    "HjrlmYyq/EB1nJBqPEF4QzXP4cGKfNqYphkkcgWZyGv8/+Vy7e73t+tLpPIQAS1Eer3v8jwrD6NLwGz8jHjBzPFuGdNSWBUL5ZPKGGpe2gcJdUQVkyWuZMxr"
    "Or4SANTSXJYLF31T+XwBHYAbY280pvu61FKNN/5OncwwHKZLPL8gDPsyw88WL86jSEOX4Q5ydGpnmLNOMR4Ec5CBUaBCAGqAy0C2SPoZjaDMb8g4csqTu9+G"
    "tx4uDuHk7vb99dVHd3O/ieMRdJyDb0gbR8aowXQkivlL1JkHriQC0gyWSECErGfwxhtA7k4VsfJJyFylwBU5p+ANuqFD+OAjEUJSRghpG+4mENAREYK5ixxU"
    "+Mp+EKF9go+I8FFBoB9j9KgblPrbAgcVKopdbNggNizBhsbkxsMMR+Ysj6Nh0/FQHfa3wRt88PTw5PTQmNhA1wpOjz0qOnxT5HDDN/aDpwcfiRFSEUKRz8VK"
    "dVSE9nS8owqeF4O7I9LdiY+OsYFDh7k84aMhU+apCvO7wTrYXKyXzxbNnVXu/tBXicBRAgWQkajYC0ZdV4F7mBiQJBFmdBJR4MPxl4qBalVtgo98e/Da1Mn3"
    "n765ubvtiXAJA2sZjZxtRkoT31j6FAF4SKMlatGOrgOYfYIAXv2A1tgN4HGQFjlbhI0z9g6jVgUP9OEVmXZGCBLgJfkT8doXel+EpeM3LRPzE6MmE9R0gBoG"
    "Ej8FNbop9Mm4C+HRqFWblseLJ3jx+KL5aXgVZhShmjd+yzphUDFiJEGMBoiBefkkxNim0Mng9NfRiEUhOjFWLMaKm1grfhpW+zLfg7v+HosUD7ji9cNVnCGU"
    "FH+A3VLxvfqT8Eqi0YYMNncPPhqtaq8WH7+9u/2QIJWSoIyv15+G1L7IbhNWjN+rILzuuQuvi2kwUptMqFsKwszT0BKbosAXfGP8ZolqszJIEZsgRQOkuHwi"
    "UkmchdjHg6dIyaGCQX/W3IK4imDxDQhXgZqEvgPnrsE8EgSeu+KIzOVl2EpLQjOLuzAkjtnAISRlere0B/KLZFF+kRKtr1Fg/mebX4QZa7WXw1V5DC/NKEaK"
    "ZdzC0n6R5vJHqfdpDSqnYLTbcLJ9u9w5h2p9xzq2QAOFQ9HxDUq8LAaS4hFBVbdpSFBATDmtVaXBz4NakIrUu0NodAox/PHf/+cO3FiC6VFw7wtVnCjD5wDY"
    "NKlB4KMkFei9aZSkJJmiH4AAuuzdt9p37dHIXZ5FyT5DZRUckuNwOFQnAulHdvahshKK9yGJTT9cBMxpbOMwSOpB5EinMgB4ykMDkEUW5H1RdQQMAx4FM4uq"
    "IVTRtcZ26EaLp9NNUUEmd4E4DoX+ghUY2dlJdZOVpVJMNqyk4oG7RhwHf09ZCsrdNWECfGWOHAI+S0BsX3jJqfg48HlUdqLmPKpDQdY8nYJ4QQUDeHoxEoNF"
    "XGSiwoGlpeWIMMfAYRFthLvPCspOSeMrt0WvjDwJB0qBwElQpkNMlQlYTEyxK3Kw/oKzDcdh0Ff0Aw+D6sBfGXvFJzmyGwYKojIlxsEuosoXNRGlcURYr/bp"
    "RCSKRLAYST8HKpMg/XQ4qSL6cdIsMnAGShjB0yOJp68ACdCO7cJOxeNoR5RKYjkOeulzNQPqEUKYDvU0ZtgTqEeWCWI1DgHl8zQT+qcpBpIcgf5VkSS2SbH1"
    "MyzQko8J8PYa/fpn9uvfs+/ef4oucy2W13ExgFWRClezYibx/zG6xddw9HGAPswrl18K8BR6BPVIK2y3PHm92L1a/twHtz7OKlNheZ76aMj2aLgw15mc88zh"
    "0KSkxhA8vn8kXmPMNMXNExApqzDENH0cJuMtN6VHHhuHFwPEZA4xWlbdCF54JO2V23NKyMdvGC3LtoUXHkl7I6w8JdkTENmXyHbNHofFWLvP0KNQHSuQ95o9"
    "kuJKDUGl7OO3qbBcIbzwSHorMg+VVuNQYjmUikoVMs0fh9BIg5ERcRTS4yWqgh5pM3aRGm1DWnkctBalSgR/5LEqtSqV4Y8/VrxMgeCPPFRldqay5AkYlWoS"
    "4nEojTU/mTkK/YlSPUI8kvxKjVJN1OM3S5RqEeKRBFhkq2r6hBMlyvQJ+Th8Rlqvgh2F+GSJ/16rx6E01p6lxzlQqkyfMBiYcHZ/11Q58dWSXt59f/vh6qE/"
    "BivIVnLhcFEhVuu6hInICaKyNqyJg3mHtQUX7XsY3oMlh1yA8OOhpUW+exfIexjYA/WHXOTv40FlpYFQGKd7GFhehZvM4njex8NY5Ig3GJd5GMBOsEUIKbdP"
    "gFQUuqodsD/S9COaB4RxiUdspmcNBq62RWasz/vECqWCVPVJXXNDTNOpWmWQmSIRZGW5Sjaqueyycqtc0ynH7kHYEEXOJPb0E7Tq8IhNknANXWJEpw8SDFki"
    "yqxxvTWv76++vflwPXl+dX9/A2QXVmh+Nlmj629ycf17/CiqBIXtb4IOck0SnkiaD2KTP5+90dl+awrkFCc/xs6D6oAp3K7Zy5zhl8v2Vj6ERLjipOjRrNN+"
    "nUnpiyPFeb+cFMknjLk9ub59uAYuPnlxff1dkDW7vfkY1ZXCiFzfbWSK6g3z/XFcR5oqughrjiJNYC3MTJoMJ0UebTQ3gXh9wcEqojwnjWZx8Wgugrgfqata"
    "gy1jwOZ2GP3EXTy2C9Kxc9X0x8IU9XkuKseDU5Y8i++wQRRYB4UgsFViO7sjosBKLDefjTEAPg+hhzGCmmIuf+OY0PMiKYFviMHlF53lZ8HyNyVRj4NAHDN4"
    "+ErTB7ENQE9S6BmL496OCX2hlERFPJbSLTXLAUktuvdJbriiLNMfGYAh2cbD2wbM1d76gv2X64vd4pnPJwnKW5xuL8/Pl0lmM0t6K2IEoCttgunaWByiag2m"
    "Z0L5blC8zgvF/qIES4N08KAlAo+S2ebm6v3NbSWgg2s4X8Dx9OrTbyJJAQvtlYipxnQdvIOTWPhBWN/iGoQxd+kuHEuWuC6oKp6xTIgdpdV1u8IEGMtMuyIy"
    "SCIo6DAneiZc2ROXPkyqJCOqcbFQCEexqGVN7Xhfx7KhPuMByJhD5WqxTzEtClfcYrKU9KUVZK04YjsBV6gFG+9GIJQlCXCmZ4PKT6CmCeDNqmqmB6oaqCyg"
    "oEk1rUUmekW8xkDiWUqEFazjkCYZwiS1t1SVa1qtnJ7tM7Ksv+xw+i11RmdGmeGFsgj0kiHoJjDK1xGIRs7s1OIxck06qAyrMmKdEmBn2H49A54oEDQmuXk/"
    "UOajk+wmeNd7qS2LimXqJgk/12qk62bhplDQ2OFC9Rm93Bf7tDbt/mfi1q2uDI8rteJLlfq6pfg/14OxqpHBUKbOYD+ic2KLytCK0LhYrBfnbxaTt9v1yQIk"
    "h2sw8mK5PMPr9OV5IDfy7bbxdjPZFtDOMt22saoQb9iE7xmBKjzp7kvUN2VI+gk2YCmBEHy7XE/erC5CwK3gKkyeNCzoOoL5k1Wzl6pcuMHeI8IO9R4RJQ2F"
    "hRhR9jZTaIUoyWlSMaa1RsDks05wV2vvmjcId0awm2Tb8cAVXOlKbSGKpKKQj2vlHNikSkRyUGMJJ7DWAVrrezj7fFusDeYaHSVSUBSZeyL0/75avzvZrmsB"
    "7psiA5nX5dW6Elw0Ht9WUZLYXjtsQj2Xac0h04YC+XrDJrPuJVFMwkQn+HBNm3x7dl/PJtgG63JefNMmnxUjXFMx46slB32du9fFwpR5KWXUSvsr1zhi+Qsg"
    "nGorFi/eLkCHetFUGc7kPquEnYbXqr7QRiMYmBO8desPlYoG1t2NqPnOkEogQwU8XvBOAaLt5cWL7XaH2+T7sYRYSRlldKP6w5jrLS99m3nmq+Y4hcu1R8JN"
    "QfCRLTGVqTQjaZH+IHs7sx8szhIArTmZVSlK2E4Lk2mnVZ9zYEd4IHjNgfBuzh8FkGcREGVyWapHAR2yJNNpHY7ZV6AViar8mXL9eI0z3+LyZxniUUXOURmW"
    "w6kbqfjWQ630rayK6P5JWNLhRq5spZ3WfeM0GptYxhE9abXBhuq5FR3hJfUXaQfUqGdpmg+mKMlwoabgplv7l9vLzYvFhU+n6oQdVleDvHM1aLAuD/a3Qb8B"
    "VjXzDFWAXBAK+zozV0Vs7kiJO8ZKc9XcVGTjDfV0hqcXZRgtqgYxAToKE2hTdEB5PhI6iyLpjO1dShDKBiBSozsF9rG063EwogViD54u2yC6wJZdp8tkg7ro"
    "gGp/JHTCDRIsIDhsWdUhOFaETiZOj2I/hRQbdCodBZuEEQfJFr5AU9jv75mKeqIPY5MJ/aLGdrGR/EjYFLWCg8fLaC0b9EWt6RIb8NIjIbQoE5Qq6vE5jFIm"
    "igjorRMXarQ+EkIlmSwqajY5jEw3KIValcZMEkuOhYscH+Wqor6Dw6hkglEE74S0EotwHQUXVWJiKXYUdSzQxDT6p9BA9HWZ0Kjy1ZYbbYy6ggOgjTnvCIkM"
    "xCjVdFDEuKqarccTy3Uea27uspcC9Zt3ivZF6WUdW6G/2qa1ndwNNB9+Ho0c2hKuUDIIcWdJTF1PGBgBvUtMoFnUoQFRYBAp2e9gQrdSt41tte2ExkZdQGoE"
    "SymYwDy1lXXqy237WgcWU206oMsiE0iZuNY6QLvcXICWt4irbXf9GliML9kD5SIYAudA5d1zRbaDgsRoDmHpR5FLkVNmvJNVR8VbL1/tt7uLoGI8WNiIyc7X"
    "Co/VcNPRWy1HA4dNlS+qCfRhsDAogIoudYm1JoS/q5+5GyDh3Bt23g0YiHpPjpBUmh3JwZEGXKNyjnC3/Xs8ArgfrgSIaKu3C1cXtXsWNCtqzKXV44qXON+e"
    "IEG/HkJJHVbIXOQM+ic1OmokxkbUd3K1X0N0HZRRYOGQhNU6Q0v+DOTPr7Zpi4EAdo7FTEL3Uv4A29ytsNYFftWoFRfeM/T5Vrt9zIXUCQMKEMAcBxABiIF1"
    "CIBRp11LZOMFOjZ0Bgx417Q2ZVeNJrTa2FeguKxXJ+8aR+UC/vLM7UMGA91RAxgsSAU3tiqPQW8uoZi72es2Ija0RPZHncSe1Ms30pdq8LlPuKQuHgjvHnUD"
    "Pjo1MpFKpkT4m6j3NOGkr6R22hN0yN/pG2S2KCnVN3CJmDVy4OqjQx4G+2XFblMd4EtJQODOge2ou4rHMz4kML/MZTLWhCxxvTi5qEk7atAW1lJOiYZo0ZG2"
    "8M2VwHSkourLwKT8pWual2lbbVSZhLLkEXX/g+bbI2r/c1sH601lHYChXJCe8jyHsSo+gPomDPGeWFLkhrSPaWQwUQX4CIwJaL3y1F23uWsG0zRl9IGnThNy"
    "zbh1XH7LlpQbslHx96ZnShURg9DnYh6MYR0Wqhlvu0oo15jRHRAgIzSjqrboeEC02zKtHOgR5CVXg5bH1VVHNf87WGJVGab6+wBq2YbTg/6PdDY3LoJCYlF4"
    "Ll0L8nlT241i7nuu1wQCXiTq4IV9IaaHKq4qHZfvjBsEaiWPhmZReqTltBDLfInSutGe7G+0pzU7Fo5xAfAB2Qk8spBkD1UrhX2U/e33tLHHwrE0vwLeKNzJ"
    "3tql8F30t7HTVh0Nx32BUgGfFe7joTqmSsu0J1xQltOQo9FqYREieKFwG/vLmgKOrL+Fn6FHo9W4wOkIFYWLwq08VOYUtk70t/dD8XIkNEVZO/aTu0/fXj/c"
    "vH9+9fWs+pdRrIOnCjSI89fbszPQeZ4vXh0YO3iqRChcfvur+2uf19Y3cv1IGY/a33x3/enAqO7vI47LfnW2xAYm1Q8YaDuFwcBy5HWakHINbmTz0BgCfcVe"
    "PTu5uv/wK79h4a/UJdaA3iTYVDJvdwCPQJXKm0zBw4PbiM++2a5/dff7b66v7h/q15OPmvVJ/jaorZ2t1uvtHueGf9EvzZDNpyNUoupZpn4RvFn9lg4IHxd1"
    "woR/MQXUn2SIA/9Ukh35HBYs6TdVf5TrRYV/K6NmJeIh/O+ZsYtc1AnX6P3DaFYyLHPe+6LL333/61/HAyZ/yEwZPFEmAl7d/Pqhs0HBh5m5qr8W8eA31z+8"
    "/8YnFPd8mpmo/nMBPz55jcEj8W64j3zX0Nkz+M9xjGfNfwfnG7LDTm/u7+/u4wnbzzI4+T8WWUBn27N4jOqDzOjwl5Jzf9ahtLN+EjuLaas37QzZJHFvG0Os"
    "klpIDnZJ23yLYbfLqvsWFqWWzsfrQyAZdgR0Ha5qN+/s1cerDwyLkj+L+3Bmz6+bnDWTgx0A6oJkBZPX1xPJ1GzESXaT82Zyl7bFMWp/Tqf4ha3UqrZj2n1D"
    "vwZr+tq50CuHu2BdAPjQFY+bXcxerRcvmAuRb26lEu/w2XaPRQ0a9/CE6tP1LGrNhzEjNQ4WcMBYHoXt7KYWg7K96wxb3FTOs6ZLo8sHbWrwBeCLsYIfTniL"
    "wtn2dPlqt9gsQAcNvU0ehTD2s4uEMIpWJGiAOrnACB7QTOCrQQI3BRu9oSqB/XwqH6DDQmUIUI5I+3JI6ACJ5+vL5fPlbvcOO/M2LtjKSbb6JSDCyauoIL0l"
    "WKsaQWcMyBc43axKx52KKiFXTw2SKfZardLFMOycY2ZsF2w9ShUitiFdajiwO5FOW1USiac0QQ/rcFI7yEZxUtqcVuzUysGodUWQpmLKgTc7p1+TPv0Mt8VW"
    "qfPBVDQ+nT0M1c/WHk9uBTGuiQF1brtgPhXO59XKYDI+zGL9XKKZq8LMdjBLZupiJgZVLT+XSlYRo+EP45VbRzVG83ITspDT+MD4hqKdOz4O4vC3rxPREro/"
    "o66Pk2M02B9IKu4afQO5PZv5RAtRffO4PsPbZtEFnIlRkoFyGkAddmWEd5eb89XbZadX5oTKmLEwFGc1dyREiqpZpo8nr7zWrkUYc2mQ2Ciz4ivoecR45g78"
    "nI4UL5gU3MgXXDK8ksF0O5/dRH0tGelyHUkdhazcEcVp40nZWJ6MyezBsQGxinmedRoD7g5SF6nbZLAqhNtJPnenmKG1KDGsT4F0k+OlRXBmMXnWnyOkEiEw"
    "ZNz3Q0X6wJoLzrEeTBUlH/TzP+x2EE3kekMGE3GfudFMBYfIdGeLMgOycV1+MhWSor8zXy52SQAGiAbsMhxARQSi79N3nvm7HrfpLr7CnZqq5kEEkxrLIOE8"
    "tmCd45m4wFNxstgth0FzseUhaLqJ/HCQ8RxoZiQ71YF0wogvF8Pq5TnFNAjfRdrduQin12GSQVVk10kqiXfvTqpjR74QBj2mQYgDwgRACOlOH8OuKlNTHQKX"
    "mMdd1BT3x4Cweu4m/CuY2oxiXKw99Wh+C8znguHx/PEqoYhjqB5KfyyPQqqMVkRbYIgO76pkbKxKi1Fo9ezCSGtAJcQsH5fOz72xgPDxiL2wiKf1aqzMHzo/"
    "s0B+ykmTGe/qj3BU0uRceU0NqQrTXOa6TrWzLuJFmA6LYZKMZXAsPIzvLv/FIheR0FUxsUNspWKC1DVwCpq2yWhEeBWT+cQiX/bZdTjn81hf1l3tiakRfcUc"
    "5KGmWYfQ8Re1ph+F0k2YTMUvVityVAV7CpuLKjKQD3wpNFU4CjWNFMVcJKhqJAqop91zxMbpmaAG5TV8+dXQlXKqLfvMWaFFoGYRrLFdqa0uptWFQ9FGX0bf"
    "cyobMxswTn2NFIo6AiqNxokDosJdqGIusMin2wVtQAWypn8X/OVyXR0O06Nc0fBkI2J14gC75yJrp6SZ/+FNuG97H60/JsrUZoNWUvDAbAhXHytD0Cokpwko"
    "6qx8ohj0CgQuD5/ZQySDXoHm4ALhWxqBjAC7gL9xAMtROjoPZAdsvGs61kyKbLyZlvtaFaxplObS4Lo0ys1YdZ23ph0ob8C14KRjDpitvSHKuSJsrWpj3EJd"
    "GaOj4UTJ5P0yS5CcL2IB3zYX29WmDW5E6go3S8qONVwZ8lpYSSwICEyFb0UfRl3PBKl8OSyutZUx46P08UOST9DEIsawjUDVZ7Wir5ya3+QtYg0t9z0VSYKO"
    "lUg/8tQH9G3Bwn177RjC8/X2/HyL9UcvFqtNnyYY84TKsLPGaM6dK8FFljsp4guJzVv1kPjc7VQ7jNLl+yUKcJyAC6/Wb4IqN+fZbHkXZxUaddyFDLgDIhlA"
    "LHyhkCoWm7lMee7Vf+GrmjQy3NUUsVV0VQA7H2MHiNCyksZgcmDq65FZX4/AjALStaxG5ee7qcd42uLFO6BJYONiryFarlGJa2UYn3ntNJRijQwT2DmlK8LG"
    "JO07JFqzwEg0jrQZ47RSeakvRlsCMhT8Z8vFyeumH3efU5Vi4I1bIgVEZo2cfenLG6K2qL+smk7CmUYFVwZeAxe/x7oKrhznD5eR8a4NxypIGOTrZY5Fmgby"
    "Zryp7Mhce0vi4zhdf8tgUjGSeerWmobjZCz2OXUtx5nLcrDOF0+n0laljOp6gkAjqZUUJQf0mhI6YZkKad9rHyjVnmE+hRCRr8BRXizd9GgWrdkIZ7tT/jL6"
    "Uy3QgA40MDZel7rEDvO08qxg7KCQ1WFpQrAzlk4n36DPWtAho1yc7VYnoCjFRS4wpnZyforCufGFIfwyUaFU4wXjSgK7BAkivD8IHYxw8Fz2Suu2012o+Sjm"
    "rgPyxZkEdVH2uKu4pxx76TjDt/IzWCReV8wrnEuMUue1TOZCZ7Q7lm4uhkVenDp4YCY5kv2akHzerParxmxzNHNIfVWMVsKKcUWZ5q10bXleXYfNWZ91qTAB"
    "D9rOThg2jtka3sv0BsCtlAHJFApa2sOiRciiEVTd9fUaPkrbNqEfq7ooujybAKm/Xi6bwlJefTnZLRen6Kfm3k/tEajUFziiysVZoyOQOj6CqQW6KcD2DG8N"
    "u/qWGaWcn919erh2d8WUdmIu/B/HebZp5LuULqyu8gv7dKA+v3DuHoCS0e5oGU4rJGb6Yji6Y1/I67GuISYJYNE6XrtsfAy0q7lnYndYZEsddtn4O+bWcURd"
    "MPy0znrybng8s6yqWFuVlNJz1dEzGRmrHTN/X9VOix+4Moju7HkRDrwODRNXqchf7TqvI/O8O52aDVag8BOLBF/NqnwXkKBVPTbq3H8SJq9Zr1tsnTnzLLpa"
    "OjSvSeYFRdmXH3OxfM+wOrMkfserAyFz2goMNG5CmngEGVB9VfuMOWPV3eFjepJsfIAKM/FsZtIR9T28j5NsWkoWElTYtt208mX5qrI6whNWpfwb3z+7e/5Z"
    "HOh+aGq6Cf2r6F7lT5yajp56H0+NhT7zU8tayA7OvR85N9sEqju3qLY1c9eNEdr5A4PrwORsLOJsH07u/Oo9iI/Fm+1HCn6GAdTN5KinMsEPYF7ZzIMAjGiG"
    "46cX7fSEGOXuubK4y/lo7OOCwpYHNTeNaucWCROTYL8pEFOuJhHWI0JHjnPj0Jpju5KSmRtEFhm6+UoqThMgrWi0oG1jUNBMT7VXtutIINbWVDW1a5V3hTqJ"
    "LBEVxH+4RtGNeWiCICTEEnaAP35SNsb8MS62IbTJhZJ1FAlqBKAePRoCPlYsG9JqBBqdwTD346eVoywEE4SbGQPbi61KHj0njYtcBVvMmDOlsMiVn7a1TLAw"
    "mwJr4vHTjjNQwrDo5uUoOLoTG/0sCI5OXykqDwwnbGzWX1RzV9ueiELOTGEde5AQmy1YGefOcbxenVXOlueL5+8mZ6vNm6j+M9F9kcAwUGFbbe4ykIdSxcPZ"
    "temd3ZYlw6L50Z+9G81pae+c4Y2iIOhhCQSUTmsScPSbZ7Kyw9mqCKfs+oqCsmoc7MqUsobKCbZQcN2Lc2QrDkMhZkPFaLZ5CIzphaBoHSyctybkIjlP7XSG"
    "6N5N1gXT/fru44dvM/HL8ec+h6PviYLpkJfuwyx3xuoT4k2MWfpwyeDe6dWf0uIfKIq7rszAw2PSkj4WVYu4gRH3Q5pUJwb18ICxTjxQcc69MARitDXOtdFq"
    "QaLTS7KqN314SF7QkKKqqHJ4wHGtlH55hplqPgYl/Nlx/1n46aAm4h6Ej5qXqp+ToWhU9wuraIYKhuSVguGf9SG54c/paCPCbKtpSQACyQJGhtUB/yRX7Vtc"
    "5cbioyJJq2cDyHgWsribEHPHMuAGgvo3/MOMty8ynhuOjSmZ7/fdmpYGrMlRRtS1aYRcNXYWi2zeqx4ZWyI8mCL7aGjKas2ry1/d02Wj03R0ow6MTgtHZ+no"
    "lh8YfYRsqCiHBlRU6Qlea2k/LVzkTQwo3goeWORN4SInozNCDi1y4egsHZ3qQ4tcOLpIt7BZ8NzoopBAeAo7O7TuvEzUY/eIeHyiGt4hOuMLXta+yjUyiaDn"
    "5NDR5AUqAACTkCRYL+oA6KREHYDnk3NPsBZl/+ix9oJ4t6qB7VQzxJVMh/ciILfm+4JWIoLzdGApDsA9LnW+YvkiYP8iw1Tg05L2TILTlPga8iA0t8ibWIUI"
    "VplxojuFuPG4JceH00PcO8mwH2iZK6OqWJz3moXwZJHHQXCeLozQhzZxU9g+SPBEZhJpDx37SGaiIsIjRUS5G5BQEZU6WXbLejeWRvXVR7gFFIlhx/bWffY4"
    "KdNSNE35IbY16hkcHy5h5vB8supckAOD78sGTxkWl7p/cDpWk1CBzptUZ6s/LAFTsWTvTK8XQ7GCgcOaEflPx1WRGJonKXzR+4fR1TCGJmxqVmQ+GlHCIj/8"
    "+ua7q++++3jz/urBezs6HzRDR38ZSTOGteRhWIZmTMnWAv8RqUzTvZ4nfLhw8JTXKnpg8KITL2nCBhvlKiPfZIlG7nlZPHjvoriHS7xKoAfE55Rz2sep3MMF"
    "g3/65u77jx+u73+VFphI/pCZKnhiUJfzhCZblYVKkaFEWeKoFB2dkxrTty5C8CIrSHSEMrVcHBh8XzZ4amIZq/sHjy0ssOdCmcy5ND921ZARc4qO8UJtb2FS"
    "kdguQLU61HSpzixZantZemCzeTK8iYY3tDt8qptadmBHOn49EinqJLfjKTkJemDHC2toiY4JQ600B1Y/HZ+H4yuWW/50fCUPLH86vorGlzKz/un4mh1Y//2I"
    "SKndw913O+L9nwqGQ5gVBVqubxSYH7l+7ouOj7wd0tJgSJ9E0w6J+FRD8nBIPhxr5B/0zLAZkAjK4luP+rHC+0IhRcJomGZ9RX3dw6NcnNE7JJ2A2v4JCi8/"
    "4I2ErJnuveZyD0fGuQ7CMh1HITp+nqew91oA7uGRYk4F/nJlc0q6LfDb6NhVC8Zp37nTkReetfaywWw+F8ji6YjE5EZTdaJ+ajjKpKJzlZwHTjPjRd74rOek"
    "eo7Pon/1p+GagQEeUA4nwVkiyXnnCR+vnypksF4CxlpYP4d1T38RO9xjb3s8PLMiVnsp6xubFbl1mK8c0jAyIljvuGpMzEnYCtKS5HRy0XsN7p8uUZJwFdLh"
    "e41o/3RZIAe+k3AALvrZi3u6yDeCDC9lX7114dzDRf4RZNjp8P3rjw+XMl+eDm/lAf5YyHwTgc+qAKr84LEZOeAvRkmQcvZex5F7uIAfIzDp4P1kiQ8PcuWY"
    "yEhKkv2gu6fjW/DQ1e3vcVOSp+n4/ULPPT3e3+3OeDp6P8G7pwe5ekfPjoZnvW5e/3QZl/d2SDQB1Qcm0JcFPL6aPwa/v8Cle7qA01dafDS8YodWh1+WddL0"
    "ZlS8OuYA/OayKPDJJNK1V//nnayWfXIjnMZ+eAMkXhopDyyNuixzgovUw857XclRducIJzizkQAXttcDzkpu6oGrJO4wZXoNLl50Py0qL33AoonsH7okPE7I"
    "mErcI32s1xQ5B03iCpfxDhYcdO419PaRSgnLjTXcN94kS8ls31gjjnAV0BDd5PaMNiTBuU5UVmZ6IRtxRpkU8f1dzRlJbgMGrrplsgGE2d7BhkRyZQbHVCcO"
    "ADfMLRhNxaQz3vtGHJTq1R1gPKB3z2S29fCV9SFsu6BpFjctc2sX/xuhc/a+kKfB/vGZjRllA3Lu6bwW2T94luh6H8+T1QFMM9vSD0tWUzoAS4/m0/uGJXF/"
    "H9fEpx8c0raiBGgGH8/yz97H8yyyf1uznL4fmJ7j2j9+j8g+FkCfH///4vEfl4gLGV8hsymm+ULu1BeS8pluflKPf1YO/mkqBz8x+f1jE+afFqlhDHxra2gd"
    "BGf+5EnTJ49FmWPewD4aMbvlj/MPDq//n5z6+hM7Ldmss8+n5WBW2RFPSz6j67Ow+IcifxIKC24/k39E/iQUFp45HJf8yT5JtPtHJ/8/JdWw8Kz8uEr5T0yJ"
    "zAULfeYLByNrjqlEZiNfPovFz1rhZ63wH4P8P3vCPnvCPnvCPnvCPnvCPnvC/sm6AkKdp24b8/m01K6AUOfxjpIjuwKyTSL+wU/L/wO1AczW"
)

def _embedded_df(encoded_text):
    raw = zlib.decompress(base64.b64decode(encoded_text))
    return pd.read_csv(io.BytesIO(raw), dtype=object)

def find_address_file():
    # Optional external override: if a local master exists, use it.
    base = Path(__file__).parent
    roots = [base / "data", base]
    exact_names = ["adress.xlsx", "adress(1).xlsx", "address.xlsx", "address(1).xlsx"]
    for root in roots:
        for name in exact_names:
            p = root / name
            if p.exists():
                return p
    for root in roots:
        if not root.exists():
            continue
        for p in root.rglob("*.xlsx"):
            if p.name.startswith("~$"):
                continue
            try:
                cols = pd.read_excel(p, nrows=0).columns
                if "จังหวัด" in cols and "เขต/อำเภอ" in cols:
                    return p
            except Exception:
                pass
    return None

def prepare_master(df):
    """Build province and district lookup lists from the address master.

    Province output and district English spelling must come from the user's
    adress.xlsx master exactly. Districts are keyed globally by district name
    because the source workbook is not guaranteed to be sorted/grouped by
    province.
    """
    df = df.copy()
    if not {"จังหวัด", "เขต/อำเภอ"}.issubset(df.columns):
        raise ValueError("ไฟล์ address master ต้องมีคอลัมน์ จังหวัด และ เขต/อำเภอ")

    province_items = []
    seen_prov = set()
    for raw in df["จังหวัด"].fillna("").map(clean_text):
        if not raw:
            continue
        th, en = thai_part(raw), english_part(raw)
        key = norm_thai(th)
        if key and key not in seen_prov:
            province_items.append({
                "key": key, "th": th, "en": en, "source": raw
            })
            seen_prov.add(key)

    district_items = []
    seen_dist = set()
    for raw in df["เขต/อำเภอ"].fillna("").map(clean_text):
        if not raw:
            continue
        th, en = thai_part(raw), english_part(raw)
        key = norm_thai(th)
        if key and key not in seen_dist:
            district_items.append({
                "key": key, "th": th, "en": en, "source": raw
            })
            seen_dist.add(key)

    if not province_items:
        raise ValueError("address master ไม่มีข้อมูลจังหวัดที่ใช้งานได้")
    if not district_items:
        raise ValueError("address master ไม่มีข้อมูลเขต/อำเภอที่ใช้งานได้")

    return df, province_items, district_items


def load_address_master():
    p = find_address_file()
    if p is not None:
        return pd.read_excel(p), p
    return _embedded_df(EMBEDDED_ADDRESS_B64), Path("embedded: adress(1).xlsx")

def find_weight_workbook():
    base = Path(__file__).parent
    roots = [base / "data", base]
    names = [
        "ข้อมูลน้ำหนักสินค้า.xlsx",
        "ไฟล์ทำ KOL ระบบใหม่ update app.xlsx",
        "KOL template.xlsx",
        "KOL_template.xlsx",
    ]
    for root in roots:
        if not root.exists():
            continue
        for name in names:
            p = root / name
            if p.exists():
                try:
                    if "ข้อมูลน้ำหนักสินค้า" in pd.ExcelFile(p).sheet_names:
                        return p
                except Exception:
                    pass
    for root in roots:
        if not root.exists():
            continue
        for p in root.rglob("*.xlsx"):
            if p.name.startswith("~$"):
                continue
            try:
                if "ข้อมูลน้ำหนักสินค้า" in pd.ExcelFile(p).sheet_names:
                    return p
            except Exception:
                pass
    return None

def load_weight_master():
    p = find_weight_workbook()
    if p is not None:
        try:
            return pd.read_excel(p, sheet_name="ข้อมูลน้ำหนักสินค้า"), p
        except Exception:
            pass
    return _embedded_df(EMBEDDED_WEIGHT_B64), Path("embedded: ข้อมูลน้ำหนักสินค้า.xlsx")

# =========================================================
# POSTAL DATA (OPTIONAL: WMS ZIP IS ALWAYS A FALLBACK)
# =========================================================

@st.cache_data(ttl=86400, show_spinner=False)
def load_postal_data():
    r = requests.get(POSTAL_DATA_URL, timeout=30)
    r.raise_for_status()
    return r.json()


def build_geo_tables(records):
    geo = pd.DataFrame(records)
    for c in [
        "provinceNameTh", "provinceNameEn", "districtNameTh", "districtNameEn",
        "subdistrictNameTh", "subdistrictNameEn"
    ]:
        if c not in geo.columns:
            geo[c] = ""
        geo[c] = geo[c].fillna("").astype(str).map(clean_text)
    if "postalCode" not in geo.columns:
        geo["postalCode"] = ""
    geo["postalCode"] = (
        geo["postalCode"].astype(str).str.extract(r"(\d{5})", expand=False).fillna("")
    )
    geo["pkey"] = geo["provinceNameTh"].map(norm_thai)
    geo["dkey"] = geo["districtNameTh"].map(norm_thai)
    geo["skey"] = geo["subdistrictNameTh"].map(norm_thai)
    return geo

# =========================================================
# RESOLUTION
# =========================================================

def resolve_province(province_items, geo, prov_raw, city_raw, area_raw, address, wms_postal):
    parts = extract_address_components(address)
    address_province = clean_location_name(parts["province"])

    # 1. A valid Receipt Province wins.
    item = exact_match(province_items, prov_raw)
    if item:
        return item, "province_exact"

    # 2. Explicit province in address.
    item = exact_match(province_items, address_province)
    if item:
        return item, "address_province"

    # 3. Postal -> province. This is critical when WMS fields are scrambled.
    if wms_postal and not geo.empty:
        pg = geo[geo["postalCode"] == wms_postal]
        pkeys = pg["pkey"].dropna().unique().tolist()
        if len(pkeys) == 1:
            for p in province_items:
                if p["key"] == pkeys[0]:
                    return p, "postal_province"

    # 4. WMS City / Area may actually be the province.
    for candidate, status in [
        (city_raw, "city_as_province"),
        (area_raw, "area_as_province"),
    ]:
        item = exact_match(province_items, candidate)
        if item:
            return item, status

    # 5. Fuzzy fallback.
    for candidate, status in [
        (prov_raw, "province_fuzzy"),
        (address_province, "address_province_fuzzy"),
        (city_raw, "city_as_province_fuzzy"),
        (area_raw, "area_as_province_fuzzy"),
    ]:
        item, score = fuzzy_match(province_items, candidate)
        if item:
            return item, f"{status}_{score:.2f}"

    return None, "province_not_found"


def resolve_district(district_items, geo, province_item, city_raw, area_raw, address, wms_postal):
    parts = extract_address_components(address)
    candidates = []

    def add(v):
        v = clean_location_name(v)
        if not v:
            return
        # WMS sometimes appends an operational zone to Bangkok districts,
        # e.g. 'หนองจอกโซน1'. The administrative district is 'หนองจอก'.
        v2 = re.sub(r"โซน\s*\d+$", "", v).strip()
        for candidate in [v2, v]:
            if candidate and norm_thai(candidate) not in {norm_thai(x) for x in candidates}:
                candidates.append(candidate)

    # Composite fields such as 'แขวงลาดยาว เขตจตุจักร' are highly reliable.
    for raw in [city_raw, area_raw]:
        comp = composite_admin_values(raw)
        for v in comp["district"]:
            add(v)

    # Prefer explicit WMS City/Area. If it is exactly 'เมือง', expand it to
    # the province-specific Mueang district instead of the generic master row.
    for raw in [city_raw, area_raw]:
        cleaned = clean_location_name(raw)
        if norm_thai(cleaned) == "เมือง":
            add("เมือง" + province_item["th"])
        else:
            add(cleaned)

    # Address district is a fallback. Do not let a generic 'เมือง' from the
    # free-text address override an explicit 'เมืองชลบุรี' / 'เมืองสระบุรี'.
    address_district = clean_location_name(parts["district"])
    if norm_thai(address_district) != "เมือง":
        add(address_district)
    else:
        add("เมือง" + province_item["th"])

    # Exact match against adress.xlsx, which is the output source of truth.
    for cand in candidates:
        item = exact_match(district_items, cand)
        if item:
            return item, "district_exact"

    # Postal + geo can disambiguate the district.
    if wms_postal and not geo.empty:
        pg = geo[(geo["pkey"] == province_item["key"]) & (geo["postalCode"] == wms_postal)]
        dkeys = pg["dkey"].dropna().unique().tolist()
        if len(dkeys) == 1:
            for item in district_items:
                if item["key"] == dkeys[0]:
                    return item, "postal_district"

    # Fuzzy district.
    for cand in candidates:
        item, score = fuzzy_match(district_items, cand)
        if item:
            return item, f"district_fuzzy_{score:.2f}"

    # City/Area may be a subdistrict. Use geo only for this fallback.
    if not geo.empty:
        gprov = geo[geo["pkey"] == province_item["key"]]
        subs = []
        for (skey, sth), g in gprov.groupby(["skey", "subdistrictNameTh"], sort=False):
            subs.append({"key": skey, "th": sth})
        sub_candidates = []
        for raw in [area_raw, city_raw, parts["subdistrict"]]:
            comp = composite_admin_values(raw)
            for v in comp["subdistrict"]:
                sub_candidates.append(v)
            if raw:
                sub_candidates.append(raw)
        sub_candidates.append(parts["subdistrict"])
        for cand in sub_candidates:
            sub = exact_match(subs, cand)
            if sub:
                matched = gprov[gprov["skey"] == sub["key"]]
                dgroups = list(matched.groupby(["dkey", "districtNameTh"], sort=False))
                if len(dgroups) == 1:
                    (dkey, dth), g = dgroups[0]
                    # Find exact English from adress.xlsx by district key.
                    item = next((d for d in district_items if d["key"] == dkey), None)
                    if item:
                        return item, "from_subdistrict"
                    return {"key": dkey, "th": dth, "en": g.iloc[0]["districtNameEn"]}, "from_subdistrict"

    return None, "district_not_found"


def resolve_address_row(row, province_items, district_items, geo):
    prov_raw = clean_location_name(row.get("Receipt Province", ""))
    city_raw = clean_text(row.get("Receipt City", ""))
    area_raw = clean_text(row.get("Receipt Area", ""))
    address = clean_text(row.get("Consignee Addr", ""))
    wms_postal = valid_postcode(row.get("Receiver Zipcode", ""))

    province, pstatus = resolve_province(
        province_items, geo, prov_raw, city_raw, area_raw, address, wms_postal
    )
    if province is None:
        return {
            "จังหวัด": "", "เขต/อำเภอ": "", "รหัสไปรษณีย์": "",
            "สถานะตรวจสอบ": "PROVINCE_NOT_FOUND",
            "จังหวัดสถานะ": pstatus, "อำเภอสถานะ": "", "ตำบลสถานะ": "",
        }

    district, dstatus = resolve_district(
        district_items, geo, province, city_raw, area_raw, address, wms_postal
    )
    if district is None:
        return {
            "จังหวัด": make_bilingual(province["th"], province["en"]),
            "เขต/อำเภอ": "", "รหัสไปรษณีย์": "",
            "สถานะตรวจสอบ": "DISTRICT_NOT_FOUND",
            "จังหวัดสถานะ": pstatus, "อำเภอสถานะ": dstatus, "ตำบลสถานะ": "",
        }

    # Postal priority:
    # 1) address subdistrict in geo, 2) unique district in geo, 3) WMS zipcode,
    # 4) postcode written in address. This reproduces the old converter behavior
    # while still guaranteeing a fallback when WMS already provides a valid ZIP.
    parts = extract_address_components(address)
    postal = ""
    postal_status = "NO_POSTAL"
    gprov = geo[geo["pkey"] == province["key"]] if not geo.empty else pd.DataFrame()
    sub = None

    if not gprov.empty:
        subs = []
        for (skey, sth), g in gprov.groupby(["skey", "subdistrictNameTh"], sort=False):
            subs.append({"key": skey, "th": sth})

        sub_candidates = []
        comp_values = [area_raw, city_raw]
        for raw in comp_values:
            comp = composite_admin_values(raw)
            sub_candidates.extend(comp["subdistrict"])
        sub_candidates.append(parts["subdistrict"])
        sub_candidates.extend([area_raw, city_raw])

        for cand in sub_candidates:
            sub = exact_match(subs, cand)
            if sub:
                break

        if sub:
            vals = sorted({str(v).strip() for v in gprov.loc[gprov["skey"] == sub["key"], "postalCode"] if re.fullmatch(r"\d{5}", str(v).strip())})
            if len(vals) == 1:
                postal = vals[0]
                postal_status = "EXACT_SUBDISTRICT"
            elif len(vals) > 1:
                for candidate in [wms_postal, parts["postal"]]:
                    if candidate in vals:
                        postal = candidate
                        postal_status = "POSTAL_MATCH"
                        break

        if not postal:
            gd = gprov[gprov["dkey"] == district["key"]]
            vals = sorted({str(v).strip() for v in gd["postalCode"] if re.fullmatch(r"\d{5}", str(v).strip())})
            if len(vals) == 1:
                postal = vals[0]
                postal_status = "DISTRICT_UNIQUE"
            elif len(vals) > 1:
                for candidate in [wms_postal, parts["postal"]]:
                    if candidate in vals:
                        postal = candidate
                        postal_status = "POSTAL_MATCH"
                        break

    if not postal and wms_postal:
        postal = wms_postal
        postal_status = "WMS_POSTAL_FALLBACK"
    if not postal and parts["postal"]:
        postal = parts["postal"]
        postal_status = "ADDRESS_POSTAL_FALLBACK"

    status = "OK" if postal else postal_status
    if "fuzzy" in pstatus.lower() or "fuzzy" in dstatus.lower():
        status = "FUZZY_MATCH_CHECK"

    return {
        "จังหวัด": make_bilingual(province["th"], province["en"]),
        "เขต/อำเภอ": make_bilingual(district["th"], district["en"]),
        "รหัสไปรษณีย์": postal,
        "สถานะตรวจสอบ": status,
        "จังหวัดสถานะ": pstatus,
        "อำเภอสถานะ": dstatus,
        "ตำบลสถานะ": "exact" if sub is not None else "not_found",
    }

# =========================================================
# WEIGHT MASTER: SEARCH BY SHEET NAME, NOT FILE NAME
# =========================================================

def find_weight_workbook():
    base = Path(__file__).parent
    roots = [base / "data", base]
    # Exact names first.
    names = [
        "ข้อมูลน้ำหนักสินค้า.xlsx",
        "ไฟล์ทำ KOL ระบบใหม่ update app.xlsx",
        "KOL template.xlsx",
        "KOL_template.xlsx",
    ]
    for root in roots:
        if not root.exists():
            continue
        for name in names:
            p = root / name
            if p.exists():
                try:
                    if "ข้อมูลน้ำหนักสินค้า" in pd.ExcelFile(p).sheet_names:
                        return p
                except Exception:
                    pass
    # Any Excel workbook containing the exact weight sheet.
    for root in roots:
        if not root.exists():
            continue
        for p in root.rglob("*.xlsx"):
            if p.name.startswith("~$"):
                continue
            try:
                if "ข้อมูลน้ำหนักสินค้า" in pd.ExcelFile(p).sheet_names:
                    return p
            except Exception:
                pass
    return None


def load_weight_master():
    p = find_weight_workbook()
    if p is None:
        return pd.DataFrame(), None
    try:
        df = pd.read_excel(p, sheet_name="ข้อมูลน้ำหนักสินค้า")
        return df, p
    except Exception:
        return pd.DataFrame(), p


def build_weight_lookup(df):
    """Exactly reproduce the old Excel weight lookup.

    Old Excel:
      A = KOL skuCode
      C = VLOOKUP(A, D:L, 9, 0)

    Therefore the authoritative key is column D ('sku') and the weight is
    column L ('单个重（KG）'). The first matching row wins, just like VLOOKUP.
    """
    lookup = {}
    barcode_lookup = {}
    if df.empty:
        return lookup, barcode_lookup

    sku_col = "sku" if "sku" in df.columns else (df.columns[3] if len(df.columns) > 3 else None)
    weight_col = "单个重（KG）" if "单个重（KG）" in df.columns else (df.columns[11] if len(df.columns) > 11 else None)
    barcode_col = "Barcode" if "Barcode" in df.columns else (df.columns[7] if len(df.columns) > 7 else None)

    if sku_col is not None and weight_col is not None:
        for _, r in df.iterrows():
            sku = clean_text(r.get(sku_col, ""))
            if sku and norm_thai(sku) not in lookup:
                wt = to_number(r.get(weight_col, ""), None)
                if wt is not None:
                    lookup[norm_thai(sku)] = float(wt)

    if barcode_col is not None and weight_col is not None:
        for _, r in df.iterrows():
            barcode = clean_barcode(r.get(barcode_col, ""))
            if barcode and barcode not in barcode_lookup:
                wt = to_number(r.get(weight_col, ""), None)
                if wt is not None:
                    barcode_lookup[barcode] = float(wt)

    return lookup, barcode_lookup

# =========================================================
# KOL OUTPUT
# =========================================================

KOL_HEADERS = [
    "shopName/店铺名称", "platform/订单平台编码", "orderType/订单类型", "orderSn/订单编号",
    "specialOrderMark/订单标记", "deliveryType/发货类型", "skuCode/sku编码", "barcode/卖家条码",
    "skuName/卖家商品名称", "quantity/数量", "actualPrice/商品实付价", "retailPrice/商品应收价",
    "placedAt/下单时间", "payTime/支付时间", "receiverName/收件人名称", "receiverTel/收件人电话号",
    "receiverMobile/收件人手机号", "receiverCountry/收件人国家", "receiverProvince/收件人省份",
    "receiverCity/收件人城市", "receiverDistrict/收件人区域", "receiverZipcode/收件人邮政编码",
    "receiverDetailAddress/收件人明细住址", "isCod/支付方式", "taxFee/税费", "freight/运费",
    "weight/包裹重量", "length/包裹长度", "width/包裹宽度", "height/包裹高度",
]


def create_kol(raw, addr_results, sku_weights, barcode_weights):
    now = datetime.now()
    rows = []
    line_weights = []
    raw_addresses = []
    missing_weights = []

    def get(name):
        return name if name in raw.columns else None

    order_col = get("ERP No.")
    sku_col = get("Product Code")
    barcode_col = get("Product Barcode")
    qty_col = get("Expect Qty")
    receiver_col = get("Receive Name")
    mobile_col = get("Receiver Mobile Number")
    address_col = get("Consignee Addr")
    comment_col = get("Comment by seller")

    for i, (_, r) in enumerate(raw.iterrows()):
        order = clean_text(r.get(order_col, "")) if order_col else ""
        sku = clean_text(r.get(sku_col, "")) if sku_col else ""
        barcode = clean_barcode(r.get(barcode_col, "")) if barcode_col else ""
        qty = to_number(r.get(qty_col, 0), 0) if qty_col else 0
        raw_receiver = r.get(receiver_col, "") if receiver_col else ""
        receiver = "" if pd.isna(raw_receiver) else str(raw_receiver)
        mobile = clean_mobile_number(r.get(mobile_col, "")) if mobile_col else ""
        # KOL phone must be the normal Thai 10-digit format.
        # WMS values are normalized first, then keep exactly one leading 0.
        kol_mobile = mobile
        raw_addr = r.get(address_col, "") if address_col else ""
        output_addr = "" if pd.isna(raw_addr) else str(raw_addr)
        comment = clean_text(r.get(comment_col, "")) if comment_col else ""

        if comment == "Express delivery":
            mark = "NORMAL"
        elif comment == "Lalamove":
            mark = "B2B_OFFLINE"
        else:
            mark = ""

        ar = addr_results.iloc[i]
        sku_key = norm_thai(sku)
        barcode_key = clean_barcode(barcode)
        weight_each = sku_weights.get(sku_key)
        found_by = "sku"
        if weight_each is None and barcode_key:
            weight_each = barcode_weights.get(barcode_key)
            found_by = "barcode"
        if weight_each is None:
            weight_each = 0.0
            found_by = "missing"
            if sku:
                missing_weights.append({
                    "row": i + 2,
                    "orderSn": order,
                    "skuCode": sku,
                    "barcode": barcode,
                    "quantity": qty,
                })

        line_weight = float(weight_each) * float(qty)
        line_weights.append(line_weight)
        raw_addresses.append(output_addr)

        rows.append({
            "shopName/店铺名称": "EastlyncOffline",
            "platform/订单平台编码": "OFFLINE",
            "orderType/订单类型": "SALE",
            "orderSn/订单编号": order,
            "specialOrderMark/订单标记": mark,
            "deliveryType/发货类型": "EXPRESS",
            "skuCode/sku编码": sku,
            "barcode/卖家条码": barcode,
            "skuName/卖家商品名称": sku,
            "quantity/数量": qty,
            "actualPrice/商品实付价": 0,
            "retailPrice/商品应收价": 0,
            "placedAt/下单时间": now if order else "",
            "payTime/支付时间": now if order else "",
            "receiverName/收件人名称": receiver,
            "receiverTel/收件人电话号": kol_mobile,
            "receiverMobile/收件人手机号": kol_mobile,
            "receiverCountry/收件人国家": "TH",
            "receiverProvince/收件人省份": ar["จังหวัด"],
            "receiverCity/收件人城市": ar["เขต/อำเภอ"],
            # Keep this exactly like the old Excel template: both U and V use postcode.
            "receiverDistrict/收件人区域": ar["รหัสไปรษณีย์"],
            "receiverZipcode/收件人邮政编码": ar["รหัสไปรษณีย์"],
            "receiverDetailAddress/收件人明细住址": output_addr,
            "isCod/支付方式": "COD",
            "taxFee/税费": "",
            "freight/运费": "",
            "weight/包裹重量": "",
            "length/包裹长度": "",
            "width/包裹宽度": "",
            "height/包裹高度": "",
        })

    # Old Excel formula: SUMIF(W:W, W[row], AE:AE) + 0.2
    package_weights = {}
    for addr, line_weight in zip(raw_addresses, line_weights):
        package_weights[addr] = package_weights.get(addr, 0.0) + line_weight
    for row, addr in zip(rows, raw_addresses):
        if row["orderSn/订单编号"]:
            row["weight/包裹重量"] = round(package_weights.get(addr, 0.0) + 0.2, 3)

    return pd.DataFrame(rows, columns=KOL_HEADERS), pd.DataFrame(missing_weights)

# =========================================================
# EXPORT / STYLE
# =========================================================

TEMPLATE_XLSX_B64 = """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"""

# These columns are intentionally text in the KOL template.
KOL_TEXT_COLS = ["E", "F", "G", "H", "P", "Q", "R", "S", "T", "W", "X", "Y", "Z", "AA", "AB", "AC", "AD"]


def load_embedded_kol_template():
    """Load the known-good KOL template embedded inside this single app.py.

    This avoids depending on an external template file and preserves the
    original header/instruction rows, colors, widths, fonts and number formats.
    """
    if not openpyxl:
        raise RuntimeError("ต้องติดตั้ง openpyxl ก่อนใช้งาน")
    raw = base64.b64decode(TEMPLATE_XLSX_B64)
    wb = load_workbook(io.BytesIO(raw))
    ws = wb[wb.sheetnames[0]]
    ws.title = "ไฟล์อัพโหลด KOL ระบบเก่า"
    return wb, ws


def style_kol_sheet(ws):
    """Apply final KOL display rules without changing the template look."""
    # Never hide rows. This fixes the WPS/Excel view that appeared to start at row 478.
    for r in range(1, ws.max_row + 1):
        ws.row_dimensions[r].hidden = False

    # Keep the two original rows visible; data starts on row 3.
    ws.row_dimensions[1].height = 33
    ws.row_dimensions[2].height = 54

    # Do NOT freeze panes. The source file has no freeze pane and this avoids
    # the WPS split-view problem where the lower pane opens around row 478.
    ws.freeze_panes = None

    # Always open at A1, not the old template's remembered Q484 / I478 view.
    sv = ws.sheet_view
    sv.topLeftCell = "A1"
    if sv.selection:
        sel = sv.selection[0]
        sel.activeCell = "A1"
        sel.sqref = "A1"
        sel.pane = "topLeft"

    # Filter starts on the real header row and includes the instruction row,
    # matching the original template structure.
    ws.auto_filter.ref = f"A1:AD{ws.max_row}"

    # Ensure all rows are expanded/visible and no outline level hides them.
    for r in range(1, ws.max_row + 1):
        ws.row_dimensions[r].hidden = False
        ws.row_dimensions[r].outlineLevel = 0

    # Exact source-template column widths are retained. Add consistent row data alignment.
    for row in ws.iter_rows(min_row=3, max_row=ws.max_row, min_col=1, max_col=30):
        for cell in row:
            cell.alignment = copy(ws.cell(3, cell.column).alignment)

    # Phone fields: text, 10 digits. Postal code: real Excel Number.
    for cell in ws["P"][2:]:
        cell.number_format = "@"
    for cell in ws["Q"][2:]:
        cell.number_format = "@"
    for cell in ws["V"][2:]:
        cell.number_format = "0"

    # Date/time fields exactly follow the known-good source template.
    for col in ("M", "N"):
        for cell in ws[col][2:]:
            cell.number_format = "m/d/yyyy;@"


def write_kol_into_template(kol_df):
    """Return a workbook using the known-good KOL template as the visual/layout base."""
    wb, ws = load_embedded_kol_template()

    # Capture the original data-row styles before deleting sample data.
    row3_styles = {}
    if ws.max_row >= 3:
        for c in range(1, 31):
            cell = ws.cell(3, c)
            row3_styles[c] = {
                "font": copy(cell.font),
                "fill": copy(cell.fill),
                "border": copy(cell.border),
                "alignment": copy(cell.alignment),
                "number_format": cell.number_format,
                "protection": copy(cell.protection),
            }

    # Remove all old sample data while keeping row 1/2 exactly as the source template.
    if ws.max_row >= 3:
        ws.delete_rows(3, ws.max_row - 2)

    for excel_row, (_, data) in enumerate(kol_df.iterrows(), start=3):
        ws.cell(excel_row, 1, data.get("shopName/店铺名称", ""))
        for col_idx, header in enumerate(KOL_HEADERS, start=1):
            if col_idx == 1:
                continue
            value = data.get(header, "")
            # Never write NaN into the workbook.
            try:
                if pd.isna(value):
                    value = ""
            except Exception:
                pass
            cell = ws.cell(excel_row, col_idx, value)
            if col_idx in row3_styles:
                stl = row3_styles[col_idx]
                cell.font = copy(stl["font"])
                cell.fill = copy(stl["fill"])
                cell.border = copy(stl["border"])
                cell.alignment = copy(stl["alignment"])
                cell.number_format = stl["number_format"]
                cell.protection = copy(stl["protection"])

            if col_idx in (16, 17):  # P/Q phone
                cell.number_format = "@"
                cell.value = clean_mobile_number(value)
            elif col_idx == 22:  # V postal code
                postal = valid_postcode(value)
                cell.value = int(postal) if postal else ""
                cell.number_format = "0"
            elif col_idx in (13, 14):  # M/N datetime
                cell.number_format = "m/d/yyyy;@"
            elif col_idx == 27:  # AA weight
                cell.number_format = "0.00"
            elif col_idx in (11, 12):
                cell.number_format = "0.00"

    style_kol_sheet(ws)
    return wb

# =========================================================
# UI
# =========================================================

st.title("📦 WMS → KOL Upload Converter")
st.caption("อัปโหลด WMS เพียงไฟล์เดียว → แปลงที่อยู่ → สร้างไฟล์อัปโหลด KOL พร้อมใช้")

with st.sidebar:
    st.header("ตั้งค่า")
    address_path = find_address_file()
    weight_path = find_weight_workbook()
    if address_path:
        st.success(f"Address master: {address_path.name}")
    else:
        st.success("Address master: ใช้ข้อมูลที่ฝังใน app.py")
    if weight_path:
        st.success(f"ข้อมูลน้ำหนักสินค้า: {weight_path.name}")
    else:
        st.success("ข้อมูลน้ำหนักสินค้า: ใช้ข้อมูลที่ฝังใน app.py")
    st.info("อัปโหลดเฉพาะไฟล์ WMS เท่านั้น")

uploaded = st.file_uploader(
    "1) อัปโหลดไฟล์ Excel จาก WMS",
    type=["xlsx", "xls"],
    accept_multiple_files=False,
)

if uploaded is None:
    st.markdown("""
    ### วิธีใช้งาน
    1. โปรแกรมมี Address master และข้อมูลน้ำหนักสำรองฝังอยู่ใน `app.py` แล้ว
    2. หากมี `adress.xlsx` หรือไฟล์น้ำหนักใน `data` ระบบจะใช้เป็นข้อมูลภายนอกแทนได้
    3. อัปโหลด **Outbound Detail Export จาก WMS เพียงไฟล์เดียว**
    3. ระบบแปลงจังหวัด / อำเภอ / รหัสไปรษณีย์
    4. ระบบคำนวณน้ำหนักตามสูตร Excel เดิม
    5. ระบบสร้าง `KOL_Upload_Ready.xlsx` พร้อมใช้งาน
    """)
    st.stop()

try:
    raw = pd.read_excel(uploaded)
except Exception as e:
    st.error(f"อ่านไฟล์ WMS ไม่สำเร็จ: {e}")
    st.stop()

required = [
    "ERP No.", "Receipt Province", "Receipt City", "Receipt Area",
    "Consignee Addr", "Product Code", "Product Barcode", "Expect Qty",
    "Receive Name", "Receiver Mobile Number", "Receiver Zipcode",
]
missing = [c for c in required if c not in raw.columns]
if missing:
    st.error("ไม่พบคอลัมน์ที่จำเป็น: " + ", ".join(missing))
    st.stop()
if raw.empty:
    st.warning("ไฟล์ WMS ไม่มีข้อมูล Order")
    st.stop()

try:
    _, province_items, district_items = prepare_master(load_address_master()[0])
except Exception as e:
    st.error(f"อ่าน address master ไม่สำเร็จ: {e}")
    st.stop()

# Postal API is helpful but not mandatory because WMS postcode/address are fallbacks.
geo = pd.DataFrame(columns=[
    "provinceNameTh","provinceNameEn","districtNameTh","districtNameEn",
    "subdistrictNameTh","subdistrictNameEn","postalCode","pkey","dkey","skey"
])
postal_note = ""
try:
    geo = build_geo_tables(load_postal_data())
except Exception as e:
    postal_note = f"ฐานข้อมูลไปรษณีย์ออนไลน์โหลดไม่ได้ จึงใช้ ZIP จาก WMS/ที่อยู่เป็น fallback ({e})"

weight_master, weight_path = load_weight_master()
if weight_master.empty:
    st.error(
        "ไม่พบข้อมูลน้ำหนักสินค้า แม้ข้อมูลสำรองที่ฝังใน app.py ก็โหลดไม่ได้ กรุณาใช้ app.py เวอร์ชันล่าสุด"
    )
    st.stop()

sku_weights, barcode_weights = build_weight_lookup(weight_master)
if not sku_weights and not barcode_weights:
    st.error("พบไฟล์น้ำหนักสินค้า แต่ไม่พบข้อมูล SKU/น้ำหนักที่ใช้งานได้")
    st.stop()

with st.spinner("กำลังแปลงที่อยู่และสร้างไฟล์ KOL..."):
    address_results = [
        resolve_address_row(r, province_items, district_items, geo)
        for _, r in raw.iterrows()
    ]
    result_df = pd.DataFrame(address_results)
    kol_df, missing_weight_df = create_kol(raw, result_df, sku_weights, barcode_weights)

# Summary
c1, c2, c3, c4 = st.columns(4)
c1.metric("WMS ทั้งหมด", f"{len(raw):,}")
c2.metric("KOL ทั้งหมด", f"{len(kol_df):,}")
c3.metric("ที่อยู่ต้องตรวจ", f"{(result_df['สถานะตรวจสอบ'] != 'OK').sum():,}")
c4.metric("รหัสไปรษณีย์ว่าง", f"{result_df['รหัสไปรษณีย์'].fillna('').astype(str).str.strip().eq('').sum():,}")

if postal_note:
    st.warning(postal_note)

# Address issues
issues = result_df[result_df["สถานะตรวจสอบ"] != "OK"].copy()
if not issues.empty:
    with st.expander("⚠️ รายการที่ควรตรวจสอบก่อนใช้งาน", expanded=True):
        preview_cols = [
            "จังหวัด", "เขต/อำเภอ", "รหัสไปรษณีย์",
            "สถานะตรวจสอบ", "จังหวัดสถานะ", "อำเภอสถานะ", "ตำบลสถานะ"
        ]
        preview = pd.concat([
            raw.loc[issues.index, [
                "ERP No.", "Receipt Province", "Receipt City", "Receipt Area", "Consignee Addr", "Receiver Zipcode"
            ]].reset_index(drop=True),
            issues[preview_cols].reset_index(drop=True),
        ], axis=1)
        st.dataframe(preview, use_container_width=True, hide_index=True)

# Weight issues
if not missing_weight_df.empty:
    with st.expander(f"⚠️ SKU ไม่มีข้อมูลน้ำหนัก ({len(missing_weight_df):,} รายการ)", expanded=False):
        st.dataframe(missing_weight_df, use_container_width=True, hide_index=True)

# Preview
st.subheader("ตัวอย่าง KOL ที่จะดาวน์โหลด")
st.dataframe(kol_df.head(100), use_container_width=True, hide_index=True)

# Export
buffer = io.BytesIO()

# Build the main KOL sheet from the embedded known-good template.
kol_wb = write_kol_into_template(kol_df)

# Add the diagnostic sheets to the same workbook.
address_result = pd.concat([
    raw.reset_index(drop=True),
    result_df.reset_index(drop=True),
], axis=1)
# openpyxl can safely append the other sheets to the template workbook.
address_result.to_excel(kol_wb, index=False, sheet_name="Address_Result")
if not issues.empty:
    check_df = pd.concat([
        raw.loc[issues.index].reset_index(drop=True),
        issues.reset_index(drop=True),
    ], axis=1)
else:
    check_df = pd.DataFrame(columns=list(raw.columns) + list(result_df.columns))
check_df.to_excel(kol_wb, index=False, sheet_name="Check_Required")
if not missing_weight_df.empty:
    missing_weight_df.to_excel(kol_wb, index=False, sheet_name="Weight_Check")
else:
    pd.DataFrame(columns=["row","orderSn","skuCode","barcode","quantity"]).to_excel(
        kol_wb, index=False, sheet_name="Weight_Check"
    )

# Remove any duplicate old diagnostic sheets if present and ensure main sheet first.
main_ws = kol_wb["ไฟล์อัพโหลด KOL ระบบเก่า"]
kol_wb._sheets.remove(main_ws)
kol_wb._sheets.insert(0, main_ws)

# Re-apply final view/format rules after adding sheets.
style_kol_sheet(main_ws)

kol_wb.save(buffer)

st.success("สร้างไฟล์ KOL สำเร็จ — ใช้งานได้โดยไม่ต้องเปิด Excel สูตรเดิม")
st.download_button(
    "⬇️ ดาวน์โหลด KOL_Upload_Ready.xlsx",
    data=buffer.getvalue(),
    file_name="KOL_Upload_Ready.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    type="primary",
)
