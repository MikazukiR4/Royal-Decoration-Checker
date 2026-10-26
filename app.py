# -*- coding: utf-8 -*-
"""ระบบตรวจสอบสิทธิ์การขอพระราชทานเครื่องราชอิสริยาภรณ์และเหรียญจักรพรรดิมาลา (ร.จ.พ.)
อ้างอิง: ระเบียบ สปน. พ.ศ. 2564
ปรับปรุง: ตัดข้อความอ้างอิงระเบียบในแถบ Hero ออกเรียบร้อย
"""
import re
import io
from datetime import date, timedelta

import openpyxl
import pandas as pd
import streamlit as st
from openpyxl.styles import PatternFill, Alignment, Font
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import get_column_letter

# ============================ CONFIG ============================
CONFIG = {
    "org": "สำนักงานศึกษาธิการจังหวัดสมุทรสงคราม",
    "approver": "ศึกษาธิการจังหวัด หรือผู้รักษาราชการแทน",
    "hr_unit": "เจ้าหน้าที่ทะเบียนประวัติของ สศจ.",
    "year_be": 2570,
    "cutoff": (5, 29),
    "salary_ref": (4, 1),
    "rjp_years": 25,
    "rjp_date": (7, 28),
    "min_service": {"T": 5, "O": 5, "C": 5},
    "salary": {
        "cs3": 19860,
        "expert": 22140,
        "op": 10190,
    },
    "docs_sets": {"sash": 3, "below": 1, "rjp": 3},
    "docs": {
        "sash": ["สำเนาบัตรข้าราชการ", "สำเนาบัตรประจำตัวประชาชน",
                 "คำสั่งวิทยฐานะชำนาญการพิเศษ / เชี่ยวชาญ",
                 "สำเนาประกาศราชกิจจานุเบกษา ชั้นรอง ({prior})",
                 "คำสั่งเลื่อนเงินเดือนฉบับปัจจุบัน",
                 "สำเนา ก.พ.7 / ก.ค.ศ.16 ตั้งแต่บรรจุถึงปัจจุบัน ทุกหน้า — {hr} รับรอง",
                 "หนังสือรับรองความประพฤติ ({approver} ลงนามเท่านั้น)", "แบบ คส.2"],
        "civil": ["สำเนาบัตรประจำตัวประชาชน",
                  "สำเนา ก.พ.7 / ก.ค.ศ.16 ตั้งแต่บรรจุถึงปัจจุบัน ทุกหน้า — {hr} รับรอง", "แบบ คส.2"],
        "contract": ["สำเนาบัตรประจำตัวประชาชน",
                     "สำเนาสัญญาจ้างต่อเนื่องตั้งแต่เริ่มจ้างจนถึงปัจจุบัน", "แบบ คส.2"],
        "rjp": ["แบบประวัติ ร.จ.พ. (พิมพ์เท่านั้น เซ็นปากกาน้ำเงิน 3 ฉบับ)", "แบบ คส.2",
                "สำเนาบัตรประจำตัวประชาชน",
                "สำเนา ก.พ.7 / ก.ค.ศ.16 ตั้งแต่บรรจุถึงปัจจุบัน ทุกหน้า ({hr} รับรอง)"],
    },
}

TH_M = ["ม.ค.", "ก.พ.", "มี.ค.", "เม.ย.", "พ.ค.", "มิ.ย.",
        "ก.ค.", "ส.ค.", "ก.ย.", "ต.ค.", "พ.ย.", "ธ.ค."]
TH_FULL = ["มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน",
           "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม"]
LADDER = ["บ.ม.", "บ.ช.", "จ.ม.", "จ.ช.", "ต.ม.", "ต.ช.", "ท.ม.", "ท.ช.", "ป.ม.", "ป.ช.", "ม.ว.ม.", "ม.ป.ช."]
NAMES = {"บ.ม.": "เบญจมาภรณ์มงกุฎไทย", "บ.ช.": "เบญจมาภรณ์ช้างเผือก", "จ.ม.": "จัตุรถาภรณ์มงกุฎไทย",
         "จ.ช.": "จัตุรถาภรณ์ช้างเผือก", "ต.ม.": "ตริตาภรณ์มงกุฎไทย", "ต.ช.": "ตริตาภรณ์ช้างเผือก",
         "ท.ม.": "ทวีติยาภรณ์มงกุฎไทย", "ท.ช.": "ทวีติยาภรณ์ช้างเผือก", "ป.ม.": "ประถมาภรณ์มงกุฎไทย",
         "ป.ช.": "ประถมาภรณ์ช้างเผือก", "ม.ว.ม.": "มหาวชิรมงกุฎ", "ม.ป.ช.": "มหาปรมาภรณ์ช้างเผือก"}
SASH = {"ป.ม.", "ป.ช.", "ม.ว.ม.", "ม.ป.ช."}
PRIOR = {"ป.ม.": "ท.ช.", "ป.ช.": "ป.ม.", "ม.ว.ม.": "ป.ช.", "ม.ป.ช.": "ม.ว.ม."}

T, O, C = "T", "O", "C"
ACCOUNTS = {
    T: "ศึกษานิเทศก์ (บัญชี 16)",
    O: "ข้าราชการสำนักงาน ม.38 ค.(2) (บัญชี 15)",
    C: "พนักงานราชการ (บัญชี 26)"
}
ACC_NO = {T: "16", O: "15", C: "26"}
CHAIN13 = [("ป.ช.", [("held", "ป.ม.", 3)]), ("ม.ว.ม.", [("held", "ป.ช.", 3)]), ("ม.ป.ช.", [("held", "ม.ว.ม.", 5)])]

RULES = {
    T: {
        "ศึกษานิเทศก์ (ไม่มีวิทยฐานะ คศ.1)": {"ranks": [("ต.ม.", [])]},
        "ศึกษานิเทศก์ (ชำนาญการ / คศ.2)": {"ranks": [("ต.ช.", []), ("ท.ม.", [("min", "cs3")]), ("ท.ช.", [("since", "cs3", 5)])]},
        "ศึกษานิเทศก์ชำนาญการพิเศษ (คศ.3)": {"ranks": [("ท.ช.", [])]},
        "ศึกษานิเทศก์เชี่ยวชาญ (คศ.4)": {"ranks": [("ท.ช.", []), ("ป.ม.", [("held", "ท.ช.", 3)]), ("ป.ช.", [("held", "ป.ม.", 5)])]},
        "ศึกษานิเทศก์เชี่ยวชาญพิเศษ (เงินวิทยฐานะ 13,000)": {"yearly": {"ต.ม.", "ต.ช.", "ท.ม.", "ท.ช.", "ป.ม."},
            "ranks": [("ต.ม.", []), ("ต.ช.", []), ("ท.ม.", []), ("ท.ช.", []), ("ป.ม.", [])] + CHAIN13},
        "ศึกษานิเทศก์เชี่ยวชาญพิเศษ (เงินวิทยฐานะ 15,600)": {"ranks": [("ป.ช.", [("held", "ป.ม.", 3)]), ("ม.ว.ม.", [("held", "ป.ช.", 3)]), ("ม.ป.ช.", [("held", "ม.ว.ม.", 3)])]},
    },
    O: {
        "ระดับปฏิบัติงาน": {"ranks": [("บ.ม.", []), ("บ.ช.", [("below", "op"), ("pos", 10)]), ("จ.ม.", [("min", "op")]), ("จ.ช.", [("min", "op"), ("pos", 10)])]},
        "ระดับชำนาญงาน": {"ranks": [("ต.ม.", []), ("ต.ช.", [("pos", 5)])]},
        "ระดับอาวุโส": {"ranks": [("ท.ม.", []), ("ท.ช.", [("pos", 5)])]},
        "ระดับปฏิบัติการ": {"ranks": [("ต.ม.", [])]},
        "ระดับชำนาญการ": {"ranks": [("ต.ช.", []), ("ท.ม.", [("min", "expert")]), ("ท.ช.", [("since", "expert", 5)])]},
        "ระดับชำนาญการพิเศษ": {"ranks": [("ท.ช.", []), ("ป.ม.", [("flag", "salmax"), ("held", "ท.ช.", 5)])]},
        "ระดับเชี่ยวชาญ": {"ranks": [("ท.ช.", []), ("ป.ม.", [("held", "ท.ช.", 3)]), ("ป.ช.", [("held", "ป.ม.", 5)])]},
        "ประเภทอำนวยการ ระดับต้น": {"ranks": [("ท.ช.", []), ("ป.ม.", [("flag", "salmax"), ("held", "ท.ช.", 3)])]},
        "ประเภทอำนวยการ ระดับสูง": {"ranks": [("ท.ช.", []), ("ป.ม.", [("held", "ท.ช.", 3)]), ("ป.ช.", [("held", "ป.ม.", 3)]), ("ม.ว.ม.", [("held", "ป.ช.", 5)])]},
    },
    C: {
        "กลุ่มงานบริการ / ทั่วไป": {"ranks": [("บ.ม.", []), ("บ.ช.", [("held", "บ.ม.", 5)]), ("จ.ม.", [("held", "บ.ช.", 5)]), ("จ.ช.", [("held", "จ.ม.", 5)])]},
    },
}

def opt_label(acc, lv):
    return f"[บัญชี {ACC_NO[acc]}] {lv}"

OPTS = {opt_label(a, lv): (a, lv) for a in RULES for lv in RULES[a]}

def be(d):
    return f"{d.day} {TH_M[d.month - 1]} {d.year + 543}"

def limit(ce, n, md):
    return date(ce - n, md[0], md[1])

def make_ctx(acc, pos, **kw):
    c = dict(acc=acc, rule=RULES[acc][pos], pos_name=pos, Y=Y, ce=Y - 543, start=None, pos_date=None,
             since_date=None, salary=0, retire=None, salmax=None, contract=True, recv_idx=-1, last_year=0,
             gazette=None, req_prev=False, rjp_got=False, E=0)
    c.update(kw)
    return c

def check(req, c):
    t, E, ce, Y_ = req[0], c["E"], c["ce"], c["Y"]
    if t == "pos":
        n = req[1] + E
        lim = limit(ce, n, CONFIG["cutoff"])
        msg = f"ดำรงตำแหน่งปัจจุบันตั้งแต่ {be(lim)} หรือก่อนหน้า (ครบ {n} ปี)"
        return (None if not c["pos_date"] else c["pos_date"] <= lim), msg
    if t in ("min", "below"):
        a = CONFIG["salary"][req[1]]
        msg = f"เงินเดือนไม่น้อยกว่า {a:,} บาท" if t == "min" else f"เงินเดือนต่ำกว่า {a:,} บาท"
        if not c["salary"]:
            return None, msg
        return (c["salary"] >= a) == (t == "min"), msg
    if t == "since":
        a, n = CONFIG["salary"][req[1]], req[2] + E
        lim = limit(ce, n, CONFIG["salary_ref"])
        msg = f"เงินเดือนถึงเกณฑ์ {a:,} บาท ภายในวันที่ {be(lim)} (ครบกำหนด {n} ปี)"
        if c["salary"] and c["salary"] < a:
            return False, msg
        return (None if not c["since_date"] else c["since_date"] < lim), msg
    if t == "held":
        n, idx = req[2] + E, LADDER.index(req[1])
        if c["recv_idx"] < idx:
            return False, f"ต้องเคยได้รับชั้น {req[1]} มาแล้วไม่น้อยกว่า {n} ปี"
        latest = Y_ - n
        msg = f"ได้รับชั้น {req[1]} ภายในปี พ.ศ. {latest} (ครองเกณฑ์ครบ {n} ปี)"
        if not c["last_year"]:
            return None, msg
        ok = c["last_year"] <= latest
        return ok, (msg if ok else f"ครองชั้น {req[1]} ยังไม่ครบ {n} ปี (ได้รับปี พ.ศ. {c['last_year']} จะครบกำหนดในปี {c['last_year'] + n})")
    if t == "flag":
        v = c[req[1]]
        msg = {"retire": f"ต้องขอในปีที่เกษียณอายุราชการ ({Y_})",
               "salmax": "ได้รับเงินเดือนขั้นสูงของระดับตำแหน่งแล้ว"}[req[1]]
        return (None if v is None else bool(v)), msg
    return None, "เงื่อนไขอื่นๆ"

def evaluate(c):
    Y_, acc, rule = c["Y"], c["acc"], c["rule"]
    out = {"status": "bad", "target": None, "reasons": [], "warns": [], "conds": []}
    n = CONFIG["min_service"][acc] + c["E"]
    lim = limit(c["ce"], n, CONFIG["cutoff"])
    base_unknown = []
    if not c["start"]:
        base_unknown.append(f"ต้องมีวันบรรจุไม่เกินวันที่ {be(lim)} (รับราชการครบ {n} ปีบริบูรณ์)")
    elif c["start"] > lim:
        s = c["start"]
        out["reasons"].append(f"เวลาราชการยังไม่ครบ {n} ปี (บรรจุวันที่ {be(s)} หลังเกณฑ์กำหนด {be(lim)}) — จะครบเกณฑ์ในปี พ.ศ. {s.year + n + 543}")
        return out
    if acc == C and not c["contract"]:
        out["reasons"].append("สัญญาจ้างต้องต่อเนื่องตั้งแต่เริ่มบรรจุจนถึงปัจจุบัน")
        return out
    
    ranks = [r for r in rule["ranks"] if LADDER.index(r[0]) > c["recv_idx"]]
    if not ranks:
        out["reasons"].append("ได้รับเครื่องราชอิสริยาภรณ์ชั้นสูงสุดตามเกณฑ์ของตำแหน่งนี้แล้ว")
        return out

    first_fail = None
    for rank, reqs in ranks:
        fails, unknown = [], list(base_unknown)
        if rank not in rule.get("yearly", set()):
            if c["req_prev"] or c["last_year"] == Y_ - 1:
                fails.append(f"ติดเงื่อนไขขอพระราชทานปีติดต่อกัน (ได้รับ/เสนอขอในปี {Y_ - 1})")
            elif c["recv_idx"] >= 0 and not c["last_year"]:
                unknown.append(f"โปรดยืนยันว่าไม่ได้เสนอขอเครื่องราชฯ ในปี {Y_ - 1} (เพื่อป้องกันการขอปีติดต่อกัน)")
        for rq in reqs:
            ok, msg = check(rq, c)
            if ok is False:
                fails.append(msg)
            elif ok is None:
                unknown.append(msg)
        if fails:
            first_fail = first_fail or (rank, fails)
            continue
        if not out["target"]:
            out["target"] = rank
            out["status"] = "cond" if unknown else "ok"
            if not unknown:
                break
            out["conds"].append((rank, unknown))
        else:
            out["conds"].append((rank, unknown))
            if not unknown:
                break

    if out["target"]:
        tg = out["target"]
        if tg in PRIOR:
            if LADDER.index(PRIOR[tg]) != c["recv_idx"]:
                out["warns"].append(f"การขอชั้น {tg} จำเป็นต้องเคยได้รับชั้น {PRIOR[tg]} มาก่อน — โปรดตรวจสอบความถูกต้องของทะเบียน")
                if out["status"] == "ok":
                    out["status"] = "warn"
            if not c["gazette"]:
                out["warns"].append(f"ต้องแนบสำเนาราชกิจจานุเบกษาหน้าประกาศของชั้น {PRIOR[tg]}")
                if c["gazette"] is False and out["status"] == "ok":
                    out["status"] = "warn"
    else:
        fail_msgs = [m for m in first_fail[1]]
        out["reasons"] = [f"ยังไม่ผ่านเกณฑ์สำหรับชั้น {first_fail[0]} ({', '.join(fail_msgs)})"]
    return out

def next_year(c, horizon=10):
    if c.get("req_prev"):
        return None
    for k in range(1, horizon + 1):
        y = c["Y"] + k
        r = evaluate(dict(c, Y=y, ce=y - 543))
        if r["status"] in ("ok", "warn", "cond"):
            return y, r["target"], r["status"] == "cond"
    return None

def evaluate_rjp(c):
    lim = date(c["ce"] - CONFIG["rjp_years"], *CONFIG["rjp_date"])
    r, w = [], []
    if not (c["start"] and c["start"] <= lim):
        r.append(f"เวลาราชการยังไม่ครบ 25 ปีบริบูรณ์ภายในวันที่ {be(date(c['ce'], *CONFIG['rjp_date']))}")
    return r, w

def sanity(c):
    w, s, today = [], c["start"], date.today()
    if s:
        if s > today:
            w.append(f"วันบรรจุ ({be(s)}) อยู่ในอนาคต โปรดตรวจสอบปี พ.ศ.")
        elif today.year - s.year > 45:
            w.append(f"วันบรรจุ ({be(s)}) ห่างมากกว่า 45 ปี โปรดตรวจสอบความถูกต้อง")
    return w

def documents(c, target, rjp=False):
    D, S = CONFIG["docs"], CONFIG["docs_sets"]
    if rjp:
        return S["rjp"], [d.format(hr=CONFIG["hr_unit"]) for d in D["rjp"]]
    if c["acc"] == C:
        return S["below"], D["contract"]
    if target in SASH:
        prior = PRIOR[target] if target in PRIOR else "ชั้นที่เคยได้รับ"
        return S["sash"], [d.format(prior=prior, approver=CONFIG["approver"], hr=CONFIG["hr_unit"]) for d in D["sash"]]
    return S["below"], [d.format(hr=CONFIG["hr_unit"]) for d in D["civil"]]

# ============================ UI (Wide Layout / Adaptive) ============================
st.set_page_config(page_title="ตรวจสอบสิทธิ์เครื่องราชอิสริยาภรณ์ สศจ.", page_icon="🏛️", layout="wide")
st.markdown("""<style>
:root{--bg:#0e1726;--panel:#1a2638;--line:#33476a;--accent:#5b8def;--txt:#ffffff}
.stApp, header[data-testid="stHeader"]{background:var(--bg) !important}
.stApp, .stApp p, .stApp span, .stApp label, .stApp li, .stApp div, .stApp small, .stApp b, .stApp strong,
.stApp h1, .stApp h2, .stApp h3, .stApp h4, .stApp h5, .stApp h6,
[data-testid="stCaptionContainer"], [data-testid="stWidgetLabel"] *, [data-testid="stMarkdownContainer"] *,
[data-testid="stMetricLabel"] *, [data-testid="stMetricValue"]{color:var(--txt) !important}
[data-testid="stSidebar"]{background:#0b2a5b !important}
[data-testid="stSidebar"] *{color:var(--txt) !important}
.block-container{max-width:100% !important;padding:2rem 3rem !important}
.hero{background:linear-gradient(135deg,#0b2a5b,#1e4d9b);color:#fff;padding:22px 26px;border-radius:10px;margin-bottom:18px;border:1px solid var(--line)}
.hero h2{color:#fff;margin:0 0 4px 0;font-size:1.45rem}.hero p{margin:0;opacity:.9;font-size:.92rem}
.pdpa-notice{background:#1a2638;border:1px dashed var(--accent);padding:10px 16px;border-radius:8px;margin-bottom:16px;font-size:0.88rem;color:#d1d5db !important;}
.sec{border-left:4px solid var(--accent);padding-left:10px;margin:18px 0 6px;font-weight:700;color:#fff;font-size:1.05rem}
.res{border-radius:10px;padding:18px 22px;margin:10px 0;border:1px solid;color:#fff}
.res *{color:#fff !important}
.ok{background:#123d27;border-color:#2e8b57}
.warn{background:#4a3a0a;border-color:#d99a00}
.bad{background:#4a1714;border-color:#c0392b}
.cond{background:#3d3410;border-color:#e0a800}
.stButton>button, .stDownloadButton>button{background:#1e4d9b !important;color:#fff !important;border:1px solid var(--accent) !important;border-radius:8px;font-weight:600}
.stButton>button{width:100%}
</style>""", unsafe_allow_html=True)

with st.sidebar:
    Y = int(st.number_input("ปีที่ขอพระราชทาน (พ.ศ.)", 2560, 2620, CONFIG["year_be"], 1))

st.markdown(f"""<div class="hero"><h2>🏛️ ระบบตรวจสอบสิทธิ์การขอพระราชทานเครื่องราชอิสริยาภรณ์</h2>
<p>ประจำปี พ.ศ. {Y} · {CONFIG["org"]}</p></div>""", unsafe_allow_html=True)

# แจ้งเตือน PDPA
st.markdown("""<div class="pdpa-notice">
🔒 <b>ประกาศความเป็นส่วนตัว (PDPA Notice):</b> ระบบนี้ประมวลผลข้อมูลในหน่วยความจำชั่วคราวเพื่อตรวจสอบสิทธิ์เท่านั้น <b>ไม่มีการจัดเก็บ บันทึก หรืออัปโหลดข้อมูลส่วนบุคคลและไฟล์ทะเบียนของท่านลงบนเซิร์ฟเวอร์</b> เพื่อความปลอดภัยและความเป็นส่วนตัวสูงสุดตามกฎหมายคุ้มครองข้อมูลส่วนบุคคล
</div>""", unsafe_allow_html=True)

def li(xs):
    return ("<ul>" + "".join(f"<li>{x}</li>" for x in xs) + "</ul>") if xs else ""

def be_date(label, key, y0, opt=False):
    st.caption(label)
    a, b, cc = st.columns([1, 1.2, 1.3])
    d = a.selectbox("วัน", range(1, 32), key=key + "d", label_visibility="collapsed")
    m = b.selectbox("เดือน", range(1, 13), format_func=lambda i: TH_M[i - 1], key=key + "m", label_visibility="collapsed")
    years = (["ไม่ทราบ"] if opt else []) + list(range(Y, 2479, -1))
    yr = cc.selectbox("ปี พ.ศ.", years, index=0 if opt else max(0, Y - y0), key=key + "y", label_visibility="collapsed")
    if yr == "ไม่ทราบ":
        return None
    try:
        return date(yr - 543, m, d)
    except ValueError:
        st.error("วันที่ไม่ถูกต้อง")
        return None

def tri(label, key):
    v = st.radio(label, ["ไม่ทราบ", "ใช่", "ไม่ใช่"], horizontal=True, key=key)
    return {"ใช่": True, "ไม่ใช่": False}.get(v)

def render(c):
    acc, Yc = c["acc"], c["Y"]
    res = evaluate(c)
    st.markdown('<div class="sec">ผลการตรวจสอบ</div>', unsafe_allow_html=True)
    if res["status"] in ("ok", "warn", "cond"):
        tg = res["target"]
        if res["status"] == "cond":
            main = res["conds"][0][1]
            st.markdown(f'<div class="res cond"><h3>🟡 ขอพระราชทานชั้น {tg} ({NAMES[tg]}) ประจำปี {Yc} ได้ หากเงื่อนไขผ่าน:</h3>{li(main)}</div>', unsafe_allow_html=True)
        else:
            cls, head = ("ok", "✅ มีสิทธิ์ขอพระราชทาน") if res["status"] == "ok" else ("warn", "⚠️ มีสิทธิ์ แต่ต้องตรวจสอบเพิ่มเติม")
            st.markdown(f'<div class="res {cls}"><h3>{head} ชั้น {tg} ({NAMES[tg]}) ประจำปี {Yc}</h3></div>', unsafe_allow_html=True)
        n, docs = documents(c, tg)
        st.markdown(f"**เอกสารที่ต้องจัดเตรียม จำนวน {n} ชุด**")
        for i, d in enumerate(docs, 1):
            st.markdown(f"{i}. {d}")
    else:
        st.markdown(f'<div class="res bad"><h3>❌ ไม่มีสิทธิ์ขอพระราชทานในปี {Yc}</h3>{li(res["reasons"])}</div>', unsafe_allow_html=True)

def single_mode():
    st.markdown('<div class="sec">1) ประเภทบุคลากรและตำแหน่ง</div>', unsafe_allow_html=True)
    acc = st.selectbox("ประเภทบุคลากร", list(ACCOUNTS), format_func=ACCOUNTS.get)
    pos = st.selectbox("ระดับตำแหน่ง", list(RULES[acc]))
    rule = RULES[acc][pos]
    all_reqs = [q for _, rs in rule["ranks"] for q in rs]
    kinds = {q[0] for q in all_reqs} | {q[1] for q in all_reqs if q[0] == "flag"}

    st.markdown('<div class="sec">2) ประวัติการรับราชการและเงินเดือน</div>', unsafe_allow_html=True)
    start = be_date("วันบรรจุ / วันเริ่มจ้าง", "st", Y - 15)
    pos_date = since_date = None
    salary = 0
    retire = salmax = None
    if "pos" in kinds:
        pos_date = be_date("วันที่ดำรงตำแหน่ง/ระดับปัจจุบัน (เลือกปี 'ไม่ทราบ' ได้)", "pd", Y - 6, opt=True)
    if kinds & {"min", "below", "since"}:
        salary = st.number_input("เงินเดือนปัจจุบัน (บาท) — 0 = ยังไม่ทราบ", 0, 300000, 0, 10)
    if "since" in kinds:
        since_date = be_date("วันที่เงินเดือนถึงขั้นต่ำตามเกณฑ์ ครั้งแรก (เลือกปี 'ไม่ทราบ' ได้)", "sd", Y - 6, opt=True)
    if "retire" in kinds:
        retire = tri(f"ขอในปีที่เกษียณอายุราชการ (ปี {Y})", "rt")
    if "salmax" in kinds:
        salmax = tri("ได้รับเงินเดือนขั้นสูงของระดับ/อันดับแล้ว", "sm")
    contract = acc != C or st.checkbox("สัญญาจ้างต่อเนื่องตั้งแต่เริ่มบรรจุจนถึงปัจจุบัน")

    st.markdown('<div class="sec">3) เครื่องราชอิสริยาภรณ์ที่เคยได้รับ</div>', unsafe_allow_html=True)
    recv = st.selectbox("ชั้นตราสูงสุดที่เคยได้รับ", ["ไม่เคยได้รับ"] + LADDER,
                        format_func=lambda x: x if x == "ไม่เคยได้รับ" else f"{x} — {NAMES[x]}")
    last_year = 0
    gazette = False
    if recv != "ไม่เคยได้รับ":
        unk = st.checkbox("ไม่ทราบปีที่ได้รับ")
        if not unk:
            last_year = st.number_input("ปี พ.ศ. ที่เสนอขอ/ได้รับครั้งล่าสุด", 2500, Y, Y - 3)
        gazette = st.checkbox(f"มีสำเนาราชกิจจานุเบกษาของชั้น {recv}")
    req_prev = st.checkbox(f"ได้เสนอขอเครื่องราชฯ ในปี {Y - 1} ไปแล้ว แต่ยังไม่ประกาศในราชกิจจาฯ")

    if st.button("🔍 ตรวจสอบสิทธิ์"):
        if not start:
            st.error("กรุณาระบุวันบรรจุ/วันเริ่มจ้างให้ถูกต้อง")
            st.stop()
        render(make_ctx(acc, pos, start=start, pos_date=pos_date, since_date=since_date, salary=salary,
                        retire=retire, salmax=salmax, contract=contract,
                        recv_idx=LADDER.index(recv) if recv in LADDER else -1, last_year=int(last_year),
                        gazette=gazette, req_prev=req_prev))

# ============================ นำเข้าสมุดทะเบียน ============================
TH_DIGIT = str.maketrans("๐๑๒๓๔๕๖๗๘๙", "0123456789")
MONTH_KEY = {}
for _i, (_ab, _full) in enumerate(zip(TH_M, TH_FULL), 1):
    MONTH_KEY[_ab.replace(".", "")] = _i
    MONTH_KEY[_full] = _i
RANK_KEY = {re.sub(r"[.\s]", "", r): r for r in LADDER}
LEVEL_O = {"ปฏิบัติการ": "ระดับปฏิบัติการ", "ปฏิบัติงาน": "ระดับปฏิบัติงาน", "ชำนาญงาน": "ระดับชำนาญงาน",
           "อาวุโส": "ระดับอาวุโส", "ชำนาญการ": "ระดับชำนาญการ", "ชำนาญการพิเศษ": "ระดับชำนาญการพิเศษ",
           "เชี่ยวชาญ": "ระดับเชี่ยวชาญ", "อำนวยการต้น": "ประเภทอำนวยการ ระดับต้น", "อำนวยการสูง": "ประเภทอำนวยการ ระดับสูง"}
LEVEL_T = {"ชำนาญการ": "ศึกษานิเทศก์ (ชำนาญการ / คศ.2)", "คศ2": "ศึกษานิเทศก์ (ชำนาญการ / คศ.2)",
           "คศ1": "ศึกษานิเทศก์ (ไม่มีวิทยฐานะ คศ.1)", "เชี่ยวชาญ": "ศึกษานิเทศก์เชี่ยวชาญ (คศ.4)"}
LV_ORDER = [("ชำนาญการพิเศษ", "ชพ"), ("เชี่ยวชาญพิเศษ", "ชชพ"), ("เชี่ยวชาญ", "ชช"), ("ชำนาญการ", "ชก"),
            ("ชำนาญงาน", "ชง"), ("ปฏิบัติการ", "ปก"), ("ปฏิบัติงาน", "ปง"), ("อาวุโส", "อส"),
            ("อำนวยการสูง", ""), ("อำนวยการต้น", "")]
POSCOL = "ประเภท/ระดับ (ตามบัญชี)"
EX = ["เงินเดือนปัจจุบัน", "วันที่ดำรงตำแหน่ง/ระดับปัจจุบัน", "วันที่เงินเดือนถึงเกณฑ์", "ได้เงินเดือนขั้นสูง",
      "เกษียณปีนี้", "เสนอขอปีที่แล้ว (รอประกาศ)"]
EDITABLE = [POSCOL] + EX
STATUSES = ["มีสิทธิ์", "ขอได้ ถ้า...", "ข้อมูลไม่ครบ", "ไม่มีสิทธิ์"]
FILL = {"มีสิทธิ์": "C6EFCE", "ขอได้ ถ้า...": "FFEB9C", "ข้อมูลไม่ครบ": "FCE4D6", "ไม่มีสิทธิ์": "FFC7CE"}
UI_FILL = {"มีสิทธิ์": "1E5631", "ขอได้ ถ้า...": "6B5300", "ข้อมูลไม่ครบ": "7A3E12", "ไม่มีสิทธิ์": "7A1C1C"}

def _t(x):
    try:
        if x is None or pd.isna(x):
            return ""
    except (TypeError, ValueError):
        pass
    return " ".join(str(x).split())

def num(v):
    s = _t(v).translate(TH_DIGIT).replace(",", "")
    try:
        return float(s) if s else 0.0
    except ValueError:
        return 0.0

def truthy(v):
    return _t(v).lower() in ("true", "1", "1.0", "ใช่", "yes", "y")

def _y(y):
    y = 2500 + y if y < 100 else y
    return y - 543 if y > 2200 else y

def parse_date(v):
    try:
        if v is None or _t(v) == "":
            return None
        if hasattr(v, "year") and hasattr(v, "month"):
            return date(_y(v.year), v.month, v.day)
        if isinstance(v, (int, float)) and not isinstance(v, bool):
            v = int(v)
            if 2400 <= v <= 2700:
                return date(v - 543, 7, 28)
            if 1900 <= v <= 2200:
                return date(v, 7, 28)
            if 1000 < v < 80000:
                return date(1899, 12, 30) + timedelta(days=v)
            return None
        s = str(v).translate(TH_DIGIT).replace("พ.ศ.", " ").replace("พศ", " ").strip()
        if re.fullmatch(r"\d{4}", s):
            return parse_date(int(s))
        m = re.match(r"(\d{4})-(\d{1,2})-(\d{1,2})", s)
        if m:
            y, mo, d = map(int, m.groups())
            return date(_y(y), mo, d)
        m = re.match(r"(\d{1,2})\s*[/\-.]\s*(\d{1,2})\s*[/\-.]\s*(\d{2,4})", s)
        if m:
            d, mo, y = map(int, m.groups())
            return date(_y(y), mo, d)
        m = re.match(r"(\d{1,2})\s*([^\d]+?)\s*(\d{2,4})", s)
        if m:
            mo = MONTH_KEY.get(re.sub(r"[.\s]", "", m.group(2)))
            if mo:
                return date(_y(int(m.group(3))), mo, int(m.group(1)))
    except Exception:
        pass
    return None

def norm_rank(v):
    s = _t(v)
    k = re.sub(r"[.\s]", "", s)
    if k.startswith("รจพ"):
        return "ร.จ.พ."
    return RANK_KEY.get(k, s.replace(" ", ""))

def norm_level(level):
    s = re.sub(r"[.\s]|ระดับ|ประเภท|วิทยฐานะ", "", _t(level))
    for full, ab in LV_ORDER:
        if full in s or (ab and s == ab):
            return full
    if "อำนวยการ" in s:
        return "อำนวยการสูง" if "สูง" in s else ("อำนวยการต้น" if "ต้น" in s else "")
    for k in ("คศ1", "คศ2"):
        if k in s:
            return k
    return ""

def map_pos(title, level):
    t, lv = re.sub(r"\s", "", title), norm_level(level)
    if "พนักงานราชการ" in t or "จ้าง" in t:
        return opt_label(C, "กลุ่มงานบริการ / ทั่วไป")
    if "ศึกษานิเทศก์" in t:
        if lv == "ชำนาญการพิเศษ":
            key = "ศึกษานิเทศก์ชำนาญการพิเศษ (คศ.3)"
        else:
            key = LEVEL_T.get(lv, "ศึกษานิเทศก์ (ชำนาญการ / คศ.2)")
        return opt_label(T, key)
    key = LEVEL_O.get(lv)
    return opt_label(O, key) if key else opt_label(O, "ระดับชำนาญการ")

def read_register(ws):
    people, cur = [], None
    for row in ws.iter_rows(min_row=1, max_col=9, values_only=True):
        row = list(row) + [None] * (9 - len(row))
        v0 = row[0]
        a = str(int(v0)) if isinstance(v0, float) and v0.is_integer() else _t(v0).translate(TH_DIGIT)
        if a.isdigit():
            cur = dict(no=int(a), name="", title="", level="", start_raw=None, rank="", rank_raw=None,
                       rjp=False, rjp_raw=None, note="")
            people.append(cur)
        elif a:
            cur = None
        if cur is None:
            continue
        cur["name"] = (cur["name"] + " " + _t(row[1])).strip()
        cur["title"] = (cur["title"] + " " + _t(row[2])).strip()
        cur["level"] = cur["level"] or _t(row[3])
        cur["start_raw"] = cur["start_raw"] if cur["start_raw"] not in (None, "") else row[4]
        cur["note"] = (cur["note"] + " " + _t(row[8])).strip()
        rk = norm_rank(row[5])
        if rk == "ร.จ.พ.":
            cur["rjp"], cur["rjp_raw"] = True, row[6]
        elif rk:
            cur["rank"], cur["rank_raw"] = rk, row[6]
    return people

def init_df(people):
    n = len(people)
    return pd.DataFrame({
        "ลำดับ": [p["no"] for p in people], "ชื่อ-สกุล": [p["name"] for p in people],
        "ตำแหน่ง/ระดับในทะเบียน": [f"{p['title']} {p['level']}".strip() for p in people],
        POSCOL: [map_pos(p["title"], p["level"] or p["title"]) for p in people],
        EX[0]: pd.Series([None] * n, dtype="float"), EX[1]: [""] * n, EX[2]: [""] * n,
        EX[3]: [None] * n, EX[4]: [None] * n, EX[5]: [False] * n})

def _coerce(col, v):
    if col == EX[0]:
        return num(v) or None
    if col in (EX[3], EX[4], EX[5]):
        return truthy(v)
    return _t(v) or None

def merge_prev(base, old):
    if "ลำดับ" not in old.columns:
        return base, 0
    old = old.copy()
    old["ลำดับ"] = pd.to_numeric(old["ลำดับ"], errors="coerce")
    old = old.dropna(subset=["ลำดับ"]).drop_duplicates("ลำดับ").set_index("ลำดับ")
    n = 0
    for i, r in base.iterrows():
        if r["ลำดับ"] not in old.index:
            continue
        o = old.loc[r["ลำดับ"]]
        if "ชื่อ-สกุล" in o.index and _t(o["ชื่อ-สกุล"]) and _t(o["ชื่อ-สกุล"]) != _t(r["ชื่อ-สกุล"]):
            continue
        for col in EDITABLE:
            if col in o.index and _t(o[col]) != "":
                if col == POSCOL and _t(o[col]) not in OPTS:
                    continue
                base.at[i, col] = _coerce(col, o[col])
        n += 1
    return base, n

def analyse(p, ex):
    sel = _t(ex.get(POSCOL))
    start, rd = parse_date(p["start_raw"]), parse_date(p["rank_raw"])
    rank = p["rank"]
    recv_idx = LADDER.index(rank) if rank in LADDER else -1
    miss = []
    if not start:
        miss.append("วันบรรจุ")
    if rank and recv_idx < 0:
        miss.append(f"อ่านชั้นตรา '{rank}' ไม่ได้")
    if recv_idx >= 0 and not rd:
        miss.append(f"วันที่ได้รับ {rank}")
    if sel not in OPTS:
        return None, None, miss + [POSCOL], []
    acc, pos = OPTS[sel]
    rule = RULES[acc][pos]
    yn = lambda v: {"ใช่": True, "ไม่ใช่": False}.get(_t(v))
    sal, pos_d, since = num(ex.get(EX[0])), parse_date(ex.get(EX[1])), parse_date(ex.get(EX[2]))
    salmax, retire = yn(ex.get(EX[3])), yn(ex.get(EX[4]))
    c = make_ctx(acc, pos, start=start, pos_date=pos_d, since_date=since, salary=sal, retire=retire,
                 salmax=salmax, recv_idx=recv_idx, last_year=(rd.year + 543) if rd else 0, gazette=None,
                 req_prev=truthy(ex.get(EX[5])), rjp_got=p["rjp"])
    res = evaluate(c)
    cand = [q for r, rs in rule["ranks"] if LADDER.index(r) > recv_idx for q in rs]
    kinds = {q[0] for q in cand} | {q[1] for q in cand if q[0] == "flag"}
    need = []
    if kinds & {"min", "below", "since"} and not sal:
        need.append(EX[0])
    if "pos" in kinds and not pos_d:
        need.append(EX[1])
    if "since" in kinds and not since:
        need.append(EX[2])
    if "salmax" in kinds and salmax is None:
        need.append(EX[3])
    if "retire" in kinds and retire is None:
        need.append(EX[4])
    return c, res, miss, need

def summary(df):
    tgt = f"ขอได้ปี {Y}"
    s = df[df["สถานะ"].isin(["มีสิทธิ์", "ขอได้ ถ้า..."])]
    rows = []
    for rank in reversed(LADDER):
        for stt in ("มีสิทธิ์", "ขอได้ ถ้า..."):
            g = s[(s[tgt] == rank) & (s["สถานะ"] == stt)]
            if len(g):
                rows.append({"ชั้นตรา": f"{rank} ({NAMES[rank]})", "สถานะ": stt, "จำนวน": len(g),
                             "รายชื่อ": ", ".join(g["ชื่อ-สกุล"])})
    return pd.DataFrame(rows, columns=["ชั้นตรา", "สถานะ", "จำนวน", "รายชื่อ"])

def style_sheet(ws, status_col=None):
    head = PatternFill("solid", fgColor="0B2A5B")
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = head
        cell.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
    ws.freeze_panes = "C2"
    for col in ws.columns:
        L = max((len(str(c.value)) for c in col if c.value is not None), default=4)
        ws.column_dimensions[col[0].column_letter].width = max(8, min(55, L + 2))
    for row in ws.iter_rows(min_row=2):
        fill = None
        if status_col:
            k = FILL.get(row[status_col - 1].value)
            fill = PatternFill("solid", fgColor=k) if k else None
        for cell in row:
            cell.alignment = Alignment(wrap_text=True, vertical="top")
            if fill:
                cell.fill = fill

# ============================ แม่แบบ Excel ============================
XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
TPL_SHEET = "กรอกข้อมูล"
TPL_COLS = ["ลำดับ", "ชื่อ-สกุล", "ตำแหน่ง", POSCOL, "วันบรรจุ/วันเริ่มจ้าง", "ชั้นตราสูงสุดที่ได้รับ",
            "ปี/วันที่ได้รับชั้นตรา", "ได้รับ ร.จ.พ. แล้ว", "ปี/วันที่ได้รับ ร.จ.พ."] + EX + ["หมายเหตุ"]
TPL_DATE_COLS = ["วันบรรจุ/วันเริ่มจ้าง", "ปี/วันที่ได้รับชั้นตรา", "ปี/วันที่ได้รับ ร.จ.พ.", EX[1], EX[2]]
TPL_REQUIRED = ["ชื่อ-สกุล", "วันบรรจุ/วันเริ่มจ้าง"]
NO_RANK = {"ไม่เคยได้รับ", "ไม่มี", "-"}

def _xl(v):
    if hasattr(v, "item"):
        v = v.item()
    if isinstance(v, bool):
        return "ใช่" if v else "ไม่ใช่"
    if isinstance(v, (int, float)):
        if v != v:
            return None
        return int(v) if float(v).is_integer() else v
    return _t(v) or None

def make_template(rows=None):
    rows = rows or []
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = TPL_SHEET
    ws.append(TPL_COLS)
    for r in rows:
        ws.append([_xl(r.get(h)) for h in TPL_COLS])
    style_sheet(ws)
    N = max(500, len(rows) + 200)
    col = lambda h: get_column_letter(TPL_COLS.index(h) + 1)
    ws.column_dimensions[col(POSCOL)].width = 52
    ws.column_dimensions[col("ชื่อ-สกุล")].width = 28
    for h in TPL_DATE_COLS:
        ws.column_dimensions[col(h)].width = 18
        for i in range(2, N + 1):
            ws[f"{col(h)}{i}"].number_format = "@"

    ins = wb.create_sheet("วิธีกรอก")
    for s in [f"วิธีกรอกแม่แบบตรวจสิทธิ์เครื่องราชฯ สศจ. ประจำปี พ.ศ. {Y}",
              "1. กรอกในแผ่นงาน 'กรอกข้อมูล' — 1 คน = 1 แถว เริ่มแถวที่ 2",
              "2. คอลัมน์ 'ประเภท/ระดับ' และ 'ชั้นตราสูงสุด' มี Dropdown ให้คลิกเลือก",
              "3. ช่องที่ไม่ทราบให้เว้นว่าง — ระบบจะสรุปเป็น 'ขอได้ ถ้า...'", ""]:
        ins.append([s])
    ins["A1"].font = Font(bold=True, size=13)
    hdr_row = ins.max_row + 1
    ins.append(["คอลัมน์", "วิธีกรอก", "ตัวอย่าง"])
    for g in [("ลำดับ", "เลขลำดับ", "1"), ("ชื่อ-สกุล", "ชื่อ นามสกุล", "นางสาวตัวอย่าง ใจดี"),
              (POSCOL, "เลือกจากรายการ Dropdown", opt_label(O, "ระดับชำนาญการ")),
              ("วันบรรจุ/วันเริ่มจ้าง", "ว/ด/ปี พ.ศ.", "1/6/2555"),
              ("ชั้นตราสูงสุดที่ได้รับ", "เลือกจากรายการ Dropdown", "ต.ช."),
              (EX[0], "เงินเดือนปัจจุบัน", "28000"), (EX[5], "ใช่/ไม่ใช่", "ไม่ใช่")]:
        ins.append(list(g))
    for cell in ins[hdr_row]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="0B2A5B")

    lst = wb.create_sheet("ตัวเลือก")
    opts, ranks = list(OPTS), ["ไม่เคยได้รับ"] + LADDER
    lst.append(["ประเภท/ระดับ", "ชั้นตรา"])
    for i in range(max(len(opts), len(ranks))):
        lst.append([opts[i] if i < len(opts) else None, ranks[i] if i < len(ranks) else None])
    lst.sheet_state = "hidden"

    def dv(h, **kw):
        v = DataValidation(allow_blank=True, showErrorMessage=True, errorTitle="ค่าไม่ถูกต้อง", **kw)
        ws.add_data_validation(v)
        v.add(f"{col(h)}2:{col(h)}{N}")

    dv(POSCOL, type="list", formula1=f"'ตัวเลือก'!$A$2:$A${len(opts) + 1}", error="เลือกจากรายการเท่านั้น")
    dv("ชั้นตราสูงสุดที่ได้รับ", type="list", formula1=f"'ตัวเลือก'!$B$2:$B${len(ranks) + 1}", error="เลือกชั้นตราจากรายการเท่านั้น")
    for h in ("ได้รับ ร.จ.พ. แล้ว", EX[3], EX[4], EX[5]):
        dv(h, type="list", formula1='"ใช่,ไม่ใช่"', error="เลือก ใช่ หรือ ไม่ใช่")
    dv(EX[0], type="decimal", operator="between", formula1="0", formula2="300000", error="ใส่ตัวเลขเงินเดือน (บาท)")

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()

def read_template(ws):
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        raise ValueError("แผ่นงานว่างเปล่า")
    idx = {_t(h): i for i, h in enumerate(rows[0]) if _t(h)}
    miss = [h for h in TPL_REQUIRED if h not in idx]
    if miss:
        raise ValueError("ไม่พบคอลัมน์ " + ", ".join(miss))
    get = lambda r, h: r[idx[h]] if h in idx and idx[h] < len(r) else None
    people, recs, seen = [], [], set()
    for r in rows[1:]:
        name = _t(get(r, "ชื่อ-สกุล"))
        if not name:
            continue
        no = int(num(get(r, "ลำดับ")))
        if no <= 0 or no in seen:
            no = max(seen, default=0) + 1
        seen.add(no)
        rank = norm_rank(get(r, "ชั้นตราสูงสุดที่ได้รับ"))
        if rank in NO_RANK:
            rank = ""
        rjp_raw = get(r, "ปี/วันที่ได้รับ ร.จ.พ.")
        people.append(dict(no=no, name=name, title=_t(get(r, "ตำแหน่ง")), level="",
                           start_raw=get(r, "วันบรรจุ/วันเริ่มจ้าง"), rank=rank,
                           rank_raw=get(r, "ปี/วันที่ได้รับชั้นตรา") if rank else None,
                           rjp=truthy(get(r, "ได้รับ ร.จ.พ. แล้ว")) or bool(_t(rjp_raw)),
                           rjp_raw=rjp_raw, note=_t(get(r, "หมายเหตุ"))))
        recs.append({"ลำดับ": no, "ชื่อ-สกุล": name, **{h: get(r, h) for h in EDITABLE}})
    return people, pd.DataFrame(recs, columns=["ลำดับ", "ชื่อ-สกุล"] + EDITABLE)

def to_tpl_rows(people, ed):
    out = []
    for p, (_, ex) in zip(people, ed.iterrows()):
        s, rd, rj = parse_date(p["start_raw"]), parse_date(p["rank_raw"]), parse_date(p["rjp_raw"])
        r = {"ลำดับ": p["no"], "ชื่อ-สกุล": p["name"], "ตำแหน่ง": f"{p['title']} {p['level']}".strip(),
             "วันบรรจุ/วันเริ่มจ้าง": be(s) if s else _t(p["start_raw"]),
             "ชั้นตราสูงสุดที่ได้รับ": p["rank"] or None,
             "ปี/วันที่ได้รับชั้นตรา": be(rd) if rd else _t(p["rank_raw"]),
             "ได้รับ ร.จ.พ. แล้ว": "ใช่" if p["rjp"] else "ไม่ใช่",
             "ปี/วันที่ได้รับ ร.จ.พ.": be(rj) if rj else _t(p["rjp_raw"]),
             "หมายเหตุ": p["note"]}
        for h in EDITABLE:
            r[h] = ex.get(h)
        out.append(r)
    return out

def _row_css(r):
    col = UI_FILL.get(r["สถานะ"])
    return [f"background-color:#{col};color:#FFFFFF" if col else "color:#FFFFFF"] * len(r)

def batch_mode():
    st.markdown('<div class="sec">นำเข้าสมุดทะเบียน (Excel)</div>', unsafe_allow_html=True)
    t1, t2 = st.columns([1, 2])
    t1.download_button("📄 ดาวน์โหลดแม่แบบ Excel เปล่า", make_template(), f"แม่แบบตรวจเครื่องราชฯ_{Y}.xlsx",
                       mime=XLSX_MIME, use_container_width=True)
    t2.caption("กรอกข้อมูลในแม่แบบ แล้วอัปโหลดกลับเข้ามาในช่องด้านล่าง")
    up = st.file_uploader("เลือกไฟล์สมุดทะเบียน หรือไฟล์แม่แบบ (.xlsx)", type=["xlsx"])
    if not up:
        st.info("อัปโหลดไฟล์แม่แบบหรือสมุดทะเบียนเพื่อเริ่มต้น")
        return
    try:
        wb = openpyxl.load_workbook(up, data_only=True)
    except Exception as e:
        st.error(f"อ่านไฟล์ไม่ได้: {e}")
        return

    tdf = None
    if TPL_SHEET in wb.sheetnames:
        sheet = TPL_SHEET
        people, tdf = read_template(wb[sheet])
    else:
        sheet = st.selectbox("แผ่นงาน (Sheet)", wb.sheetnames) if len(wb.sheetnames) > 1 else wb.sheetnames[0]
        people = read_register(wb[sheet])
    if not people:
        st.warning("ไม่พบรายชื่อ")
        return

    key = f"ex_{up.name}_{up.size}_{sheet}"
    if key not in st.session_state:
        base = init_df(people)
        if tdf is not None:
            base, _ = merge_prev(base, tdf)
        st.session_state[key] = base

    yn_col = st.column_config.SelectboxColumn(options=["ใช่", "ไม่ใช่"])
    cfg = {
        POSCOL: st.column_config.SelectboxColumn(options=list(OPTS), width="large"),
        EX[0]: st.column_config.NumberColumn(min_value=0, step=10, format="%d"),
        EX[1]: st.column_config.TextColumn(help="เช่น 1/9/2565"),
        EX[2]: st.column_config.TextColumn(help="เช่น 1/4/2560"),
        EX[3]: yn_col, EX[4]: yn_col, EX[5]: st.column_config.CheckboxColumn()
    }
    ed = st.data_editor(st.session_state[key], hide_index=True, use_container_width=True,
                        disabled=["ลำดับ", "ชื่อ-สกุล", "ตำแหน่ง/ระดับในทะเบียน"], column_config=cfg, key=key + "_ed")

    rows, ctxs = [], {}
    tgt_col = f"ขอได้ปี {Y}"
    for i, (p, (_, ex)) in enumerate(zip(people, ed.iterrows())):
        c, res, miss, need = analyse(p, ex)
        rd = parse_date(p["rank_raw"])
        got = f"{p['rank']} ({rd.year + 543})" if rd else (p["rank"] or "-")
        target, reasons, status, nxt = "-", "", "ไม่มีสิทธิ์", "-"
        if c is None:
            status, reasons = "ข้อมูลไม่ครบ", f"โปรดเลือก '{POSCOL}' ให้ถูกต้อง"
        else:
            ctxs[i] = c
            if res["status"] in ("ok", "warn"):
                status, target = "มีสิทธิ์", res["target"]
                reasons = ", ".join(res["warns"]) if res["warns"] else "ผ่านเกณฑ์ครบถ้วน"
            elif res["status"] == "cond":
                status, target = "ขอได้ ถ้า...", res["target"]
                cond_list = res["conds"][0][1]
                reasons = f"รอ Data เพิ่มเติม ({'; '.join(cond_list)})"
                if res["warns"]:
                    reasons += " | " + ", ".join(res["warns"])
            else:
                reasons = ", ".join(res["reasons"])
                ny = next_year(c)
                if c["req_prev"]:
                    nxt = f"รอประกาศราชกิจจาฯ ปี {Y - 1}"
                elif ny:
                    nxt = f"ปี พ.ศ. {ny[0]} (ชั้น {ny[1]})"
        
        rows.append({
            "ลำดับ": p["no"], "ชื่อ-สกุล": p["name"], "ตำแหน่ง/ระดับ": f"{p['title']} {p['level']}".strip(),
            "ชั้นตราที่มี (ปี)": got, "สถานะ": status, tgt_col: target,
            "ข้อมูลที่ขาด": ", ".join(miss + need) if (miss or need) else "ครบถ้วน",
            "เหตุผล / ข้อควรตรวจ": reasons, "หมายเหตุเดิม": p["note"] or "-"
        })
    df = pd.DataFrame(rows)

    st.markdown('<div class="sec">ผลการตรวจ</div>', unsafe_allow_html=True)
    st.dataframe(df.style.apply(_row_css, axis=1), hide_index=True, use_container_width=True)

    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        df.to_excel(w, sheet_name="ผลตรวจ", index=False)
        style_sheet(w.sheets["ผลตรวจ"], df.columns.get_loc("สถานะ") + 1)
    st.download_button("⬇️ ดาวน์โหลดผลตรวจ (Excel)", buf.getvalue(), f"ผลตรวจเครื่องราชฯ_{Y}.xlsx", mime=XLSX_MIME)

mode = st.radio("โหมดการทำงาน", ["📥 นำเข้าสมุดทะเบียน (Excel)", "✍️ กรอกรายคน"], horizontal=True)
if mode.startswith("📥"):
    batch_mode()
else:
    single_mode()