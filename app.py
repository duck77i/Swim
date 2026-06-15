# -*- coding: utf-8 -*-
"""수영 대회 기록 뷰어 (Streamlit 웹앱) — 같은 폴더의 swim_records.csv 사용"""
import os, re, csv, io
import streamlit as st

st.set_page_config(page_title="수영 대회 기록", layout="wide")

def _check_password():
    """st.secrets['app_password'] 가 설정돼 있으면 비밀번호 잠금.
    (로컬 실행 등 secrets 없으면 잠금 없이 통과)"""
    try:
        pw = st.secrets["app_password"]
    except Exception:
        return True
    if not pw:
        return True
    if st.session_state.get("_auth_ok"):
        return True
    st.markdown("### 🔒 접근 비밀번호")
    entered = st.text_input("비밀번호를 입력하세요", type="password",
                            label_visibility="collapsed")
    if entered == pw:
        st.session_state["_auth_ok"] = True
        st.rerun()
    elif entered:
        st.error("비밀번호가 틀렸습니다.")
    st.stop()


_check_password()
_HERE = os.path.dirname(os.path.abspath(__file__))
CSV_GZ = os.path.join(_HERE, "swim_records.csv.gz")
CSV_PATH = os.path.join(_HERE, "swim_records.csv")

VIEW_COLS = ["대회명", "종별", "세부종목", "구분", "라운드", "순위",
             "이름", "학년", "소속", "시도", "예선기록", "기록"]


# ───────── 공통 유틸 ─────────
def parse_time(t):
    t = (t or "").strip()
    if not t or not re.match(r"^[\d:.]+$", t):
        return None
    try:
        p = t.split(":"); s = float(p[-1])
        if len(p) >= 2: s += int(p[-2]) * 60
        if len(p) >= 3: s += int(p[-3]) * 3600
        return s
    except Exception:
        return None


def rank_num(cell):
    m = re.match(r"(\d+)위", cell or "")
    return int(m.group(1)) if m else 9999


@st.cache_data(show_spinner="데이터 불러오는 중…")
def load_rows():
    import gzip, io, urllib.request
    rows = None
    # 1순위: Secrets의 data_url(비공개 구글드라이브 등)에서 다운로드
    url = None
    try:
        url = st.secrets["data_url"]
    except Exception:
        url = None
    if url:
        req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
        data = urllib.request.urlopen(req, timeout=90).read()
        if data[:2] == b"\x1f\x8b":          # gzip
            text = gzip.decompress(data).decode("utf-8-sig")
        else:                                  # 비압축 CSV거나 확인페이지
            text = data.decode("utf-8-sig", errors="replace")
            if "<html" in text[:200].lower():
                raise RuntimeError("데이터 링크가 파일을 직접 내려주지 않습니다. "
                                   "구글드라이브 '링크가 있는 모든 사용자' 공유로 설정하세요.")
        rows = list(csv.DictReader(io.StringIO(text)))
    # 2순위: 같은 폴더의 gz/CSV (로컬 실행용)
    elif os.path.exists(CSV_GZ):
        with gzip.open(CSV_GZ, "rt", encoding="utf-8-sig", newline="") as f:
            rows = list(csv.DictReader(f))
    else:
        with open(CSV_PATH, "r", encoding="utf-8-sig", newline="") as f:
            rows = list(csv.DictReader(f))
    out = []
    for r in rows:
        e = (r.get("세부종목") or "").replace("M", "m")
        if "비등록" in e:
            continue
        k = r.get("종별") or ""
        g = r.get("구분") or ""
        if g not in ("남자", "여자", "혼성"):
            g = "남자" if k.startswith("남자") else "여자" if k.startswith("여자") else ""
        kf = k.replace("초등부", "유년부") if ("4학년이하" in e and "초등부" in k) else k
        r["세부종목"] = e
        r["구분"] = g
        r["_종목f"] = e.replace("[4학년이하부]", "")
        r["_종별f"] = kf
        r["_년도"] = (r.get("대회기간") or "")[:4]
        r["_pid"] = r.get("선수ID") or (r.get("이름", "") + "|" + r.get("소속", ""))
        out.append(r)
    return out


def year_of(r):
    return r.get("_년도", "")


@st.cache_data(show_spinner=False)
def rankmaps(stroke, gender):
    """(년도,종별f,종목f)별 연간 종합순위. {(key,pid):rank}"""
    rows = load_rows()
    groups = {}
    for r in rows:
        if not r["_종목f"].startswith(stroke) or r["구분"] != gender:
            continue
        y = r["_년도"]
        if not y.isdigit():
            continue
        key = (y, r["_종별f"], r["_종목f"])
        pid = r["_pid"]
        for fld in ("예선기록", "기록"):
            t = parse_time(r.get(fld))
            if t is not None:
                d = groups.setdefault(key, {})
                if pid not in d or t < d[pid]:
                    d[pid] = t
    rm = {}
    for key, d in groups.items():
        srt = sorted(d.items(), key=lambda kv: kv[1])
        prev, pr = None, 0
        for i, (pid, t) in enumerate(srt, 1):
            rk = pr if t == prev else i
            prev, pr = t, rk
            rm[(key, pid)] = rk
    return rm


def to_excel_bytes(headers, data_rows):
    import xlsxwriter
    buf = io.BytesIO()
    wb = xlsxwriter.Workbook(buf, {"in_memory": True})
    ws = wb.add_worksheet("data")
    ws.write_row(0, 0, headers)
    for i, row in enumerate(data_rows, 1):
        ws.write_row(i, 0, [str(v) for v in row])
    wb.close()
    return buf.getvalue()


rows = load_rows()
st.title("🏊 수영 대회 기록 뷰어")
st.caption(f"총 {len(rows):,}건 · 대회 {len(set(r['대회명'] for r in rows))}개")

tab1, tab2, tab3, tab4 = st.tabs(["대회 목록", "기록 보기", "종합 순위", "성장 추적"])

# ───────── 대회 목록 ─────────
with tab1:
    agg = {}
    for r in rows:
        d = agg.setdefault(r["대회명"], {"년도": year_of(r), "기간": r.get("대회기간", ""),
                                        "장소": r.get("개최장소", ""), "n": 0, "ev": set()})
        d["n"] += 1
        d["ev"].add(r["_종목f"])
    table = [{"년도": d["년도"], "대회명": nm, "대회기간": d["기간"], "개최장소": d["장소"],
              "기록수": d["n"], "종목수": len(d["ev"])}
             for nm, d in agg.items()]
    table.sort(key=lambda x: x["대회기간"], reverse=True)
    st.dataframe(table, use_container_width=True, hide_index=True)

# ───────── 기록 보기 ─────────
with tab2:
    years = ["(전체)"] + sorted({year_of(r) for r in rows if year_of(r)}, reverse=True)
    c1, c2, c3 = st.columns(3)
    fy = c1.selectbox("년도", years, key="v_y")
    comp_pool = [r["대회명"] for r in rows
                 if fy in ("(전체)", "") or year_of(r) == fy]
    comps = ["(전체)"] + sorted(set(comp_pool))
    fc = c2.selectbox("대회", comps, key="v_c")
    txt = c3.text_input("검색 (이름/소속)", key="v_t")
    c4, c5, c6 = st.columns(3)
    kinds = sorted({r["_종별f"] for r in rows if r["_종별f"]})
    events = sorted({r["_종목f"] for r in rows if r["_종목f"]})
    fk = c4.multiselect("종별", kinds, key="v_k")
    fe = c5.multiselect("세부종목", events, key="v_e")
    frd = c6.multiselect("라운드", ["예선", "결승"], key="v_r")

    res = []
    for r in rows:
        if fy not in ("(전체)", "") and year_of(r) != fy:
            continue
        if fc not in ("(전체)", "") and r["대회명"] != fc:
            continue
        if fk and r["_종별f"] not in fk:
            continue
        if fe and r["_종목f"] not in fe:
            continue
        if frd and r.get("라운드") not in frd:
            continue
        if txt and txt not in (r.get("이름", "") + " " + r.get("소속", "")):
            continue
        res.append({c: r.get(c, "") for c in VIEW_COLS})
    st.write(f"**{len(res):,}건**")
    st.dataframe(res[:5000], use_container_width=True, hide_index=True)
    if res:
        st.download_button("엑셀로 추출",
                           to_excel_bytes(VIEW_COLS, [[d[c] for c in VIEW_COLS] for d in res]),
                           "기록.xlsx", key="v_dl")

# ───────── 종합 순위 ─────────
def promote(kind, grade):
    rules = [("유년부", 4, "초등부", "5"), ("초등부", 6, "중학부", "1"),
             ("중학부", 3, "고등부", "1"), ("고등부", 3, "대학부", "1"),
             ("대학부", 4, "일반부", "")]
    if not str(grade).isdigit():
        return kind, grade
    g = int(grade)
    for tok, mx, nxt, st0 in rules:
        if tok in kind:
            return (kind, str(g + 1)) if g + 1 <= mx else (kind.replace(tok, nxt), st0)
    return kind, grade


with tab3:
    yrs = sorted({year_of(r) for r in rows if year_of(r)}, reverse=True)
    yopts = ([str(int(yrs[0]) + 1) + "(예상)"] + yrs) if yrs else []
    c1, c2, c3 = st.columns(3)
    ry = c1.selectbox("년도", yopts, key="r_y")
    rk_kind = c2.selectbox("종별", ["(선택)"] + sorted({r["_종별f"] for r in rows if r["_종별f"]}), key="r_k")
    rk_ev = c3.selectbox("세부종목", ["(선택)"] + sorted({r["_종목f"] for r in rows if r["_종목f"]}), key="r_e")
    st.caption("그 해 모든 대회의 개인 최고기록(예선·결승 중 빠른 쪽)으로 만든 순위. "
               "'(예상)'은 전년도 기록을 진급 종별로 재배치.")
    if rk_kind != "(선택)" and rk_ev != "(선택)" and ry:
        pred = ry.endswith("(예상)")
        base = str(int(ry[:4]) - 1) if pred else ry
        best = {}
        for r in rows:
            if year_of(r) != base:
                continue
            kk, gg = (promote(r["_종별f"], r.get("학년", "")) if pred
                      else (r["_종별f"], r.get("학년", "")))
            if kk != rk_kind or r["_종목f"] != rk_ev:
                continue
            pid = r["_pid"]
            for fld in ("예선기록", "기록"):
                t = parse_time(r.get(fld))
                if t is None:
                    continue
                if pid not in best or t < best[pid]["s"]:
                    best[pid] = {"s": t, "기록": r.get(fld, ""), "이름": r.get("이름", ""),
                                 "학년": gg, "소속": r.get("소속", ""), "시도": r.get("시도", ""),
                                 "대회": r.get("대회명", "")}
        srt = sorted(best.values(), key=lambda b: b["s"])
        out = []
        prev, pr = None, 0
        for i, b in enumerate(srt, 1):
            rk = pr if b["s"] == prev else i
            prev, pr = b["s"], rk
            out.append({"순위": rk, "이름": b["이름"], "학년": b["학년"], "소속": b["소속"],
                        "시도": b["시도"], "기록": b["기록"], "달성대회": b["대회"]})
        st.write(f"**{len(out)}명**")
        st.dataframe(out, use_container_width=True, hide_index=True)
        if out:
            cols = ["순위", "이름", "학년", "소속", "시도", "기록", "달성대회"]
            st.download_button("엑셀로 추출",
                               to_excel_bytes(cols, [[d[c] for c in cols] for d in out]),
                               f"종합순위_{ry}_{rk_kind}_{rk_ev}.xlsx", key="r_dl")
    else:
        st.info("종별과 세부종목을 선택하세요.")

# ───────── 성장 추적 ─────────
def birth_year(rs):
    offs = [("유년부", 6), ("초등부", 6), ("중학부", 12), ("고등부", 15), ("대학부", 18)]
    votes = {}
    for r in rs:
        g, y = r.get("학년", ""), year_of(r)
        if not (g.isdigit() and y.isdigit()):
            continue
        for tok, off in offs:
            if tok in r["_종별f"]:
                b = int(y) - int(g) - off
                votes[b] = votes.get(b, 0) + 1
                break
    return str(max(votes, key=votes.get)) if votes else ""


with tab4:
    c1, c2 = st.columns(2)
    g_gender = c1.selectbox("성별", ["남자", "여자"], key="g_sex")
    g_stroke = c2.selectbox("종목(영법)", ["(선택)", "자유형", "배영", "평영", "접영", "개인혼영"], key="g_str")
    st.markdown("**순위 조건 필터** — 무관 / 이내 N / 밖 N (모두 만족하는 선수만)")
    stages = ["초등(초4/초6)", "중학부", "고등부", "대학부", "일반부", "최근"]
    cols = st.columns(len(stages))
    conds = {}
    for i, sname in enumerate(stages):
        with cols[i]:
            st.caption(sname)
            mode = st.selectbox("조건", ["무관", "이내", "밖"],
                                index=(1 if sname == "초등(초4/초6)" else 0),
                                key="gc_" + str(i), label_visibility="collapsed")
            n = st.number_input("N", 1, 99, (3 if sname == "초등(초4/초6)" else 8),
                                key="gn_" + str(i), label_visibility="collapsed")
            conds[sname] = (mode, n)

    if g_stroke != "(선택)":
        rm = rankmaps(g_stroke, g_gender)
        latest = max((int(year_of(r)) for r in rows if year_of(r).isdigit()), default=0)
        last_act = {}
        for r in rows:
            y = year_of(r)
            if y.isdigit():
                last_act[r["_pid"]] = max(last_act.get(r["_pid"], 0), int(y))
        srows = [r for r in rows if r["_종목f"].startswith(g_stroke) and r["구분"] == g_gender]
        bypid = {}
        for r in srows:
            bypid.setdefault(r["_pid"], []).append(r)

        def is_elem(r, gr):
            return r.get("학년") == gr and ("초등부" in r["_종별f"] or "유년부" in r["_종별f"])

        def stage_cell(pid, rs, pred, tag=""):
            cand = []
            for r in rs:
                if pred(r):
                    rk = rm.get(((year_of(r), r["_종별f"], r["_종목f"]), pid))
                    if rk:
                        dist = r["_종목f"][len(g_stroke):]
                        dn = int(re.sub(r"[^0-9]", "", dist) or 0)
                        cand.append((rk, dn, dist, year_of(r)))
            if not cand:
                return ""
            rk, _d, dist, y = min(cand)
            suf = " (" + ((tag + "·") if tag else "") + y + ")"
            return f"{rk}위 {dist}{suf}"

        def level_cell(pid, rs, level):
            tag = {"중학부": "중", "고등부": "고"}.get(level, "")
            for gr in ("3", "2", "1"):
                c = stage_cell(pid, rs,
                               lambda r, gg=gr: r.get("학년") == gg and level in r["_종별f"],
                               tag=("" if gr == "3" else tag + gr))
                if c:
                    return c
            return stage_cell(pid, rs, lambda r: level in r["_종별f"])

        def fcond(name, rk):
            mode, n = conds[name]
            if mode == "무관":
                return True
            return (rk <= n) if mode == "이내" else (rk > n)

        out = []
        for pid, rs in bypid.items():
            g4 = stage_cell(pid, rs, lambda r: is_elem(r, "4"))
            g6 = stage_cell(pid, rs, lambda r: is_elem(r, "6"))
            m3 = level_cell(pid, rs, "중학부")
            h3 = level_cell(pid, rs, "고등부")
            uv = stage_cell(pid, rs, lambda r: "대학부" in r["_종별f"])
            gn = stage_cell(pid, rs, lambda r: "일반부" in r["_종별f"])
            lastyr = max((int(year_of(r)) for r in rs if year_of(r).isdigit()), default=0)
            lc = stage_cell(pid, rs, lambda r: year_of(r) == str(lastyr))
            if not (fcond("초등(초4/초6)", min(rank_num(g4), rank_num(g6))) and
                    fcond("중학부", rank_num(m3)) and fcond("고등부", rank_num(h3)) and
                    fcond("대학부", rank_num(uv)) and fcond("일반부", rank_num(gn)) and
                    fcond("최근", rank_num(lc))):
                continue
            byr = birth_year(rs)
            la = last_act.get(pid, 0)
            if byr.isdigit():
                b = int(byr)
                dues = {"g4": b+10, "g6": b+12, "m3": b+13, "h3": b+16, "uv": b+19, "gn": b+19}
                vals = {"g4": g4, "g6": g6, "m3": m3, "h3": h3, "uv": uv, "gn": gn}
                for k2 in vals:
                    if not vals[k2] and latest >= dues[k2]:
                        vals[k2] = "은퇴(예상)" if la < dues[k2] else "-"
                g4, g6, m3, h3, uv, gn = (vals["g4"], vals["g6"], vals["m3"],
                                          vals["h3"], vals["uv"], vals["gn"])
            last = max(rs, key=lambda r: r.get("대회기간", ""))
            out.append((int(byr) if byr.isdigit() else 9999,
                        {"출생(추정)": byr, "이름": last.get("이름", ""),
                         "최근소속": last.get("소속", ""), "초4": g4, "초6": g6,
                         "중학부(3학년)": m3, "고등부(3학년)": h3, "대학부": uv,
                         "일반부": gn, "최근 연간 종합순위": lc}))
        out.sort(key=lambda x: x[0])
        data = [d for _b, d in out]
        st.write(f"**{len(data)}명**  ·  '은퇴(예상)'=이후 활동 없음 / '-'=그 단계 기록만 없음 / 빈칸=아직 해당 나이 아님")
        st.dataframe(data, use_container_width=True, hide_index=True)
        if data:
            gc = list(data[0].keys())
            st.download_button("엑셀로 추출",
                               to_excel_bytes(gc, [[d[c] for c in gc] for d in data]),
                               f"성장추적_{g_gender}_{g_stroke}.xlsx", key="g_dl")
    else:
        st.info("종목(영법)을 선택하세요.")
