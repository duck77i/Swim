# -*- coding: utf-8 -*-
"""수영 대회 기록 뷰어 (Streamlit, 메모리 효율 pandas 버전)"""
import os, re, io
import pandas as pd
import streamlit as st

st.set_page_config(page_title="수영 대회 기록", layout="wide")


def _check_password():
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
VIEW_COLS = ["대회명", "종별", "세부종목", "구분", "라운드", "순위",
             "이름", "학년", "소속", "시도", "예선기록", "기록"]


def parse_time(t):
    if not isinstance(t, str):
        return None
    t = t.strip()
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


@st.cache_data(show_spinner="데이터 불러오는 중…", max_entries=1)
def load_df():
    import gzip, urllib.request
    raw = None
    url = None
    try:
        url = st.secrets["data_url"]
    except Exception:
        url = None
    if url:
        req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
        data = urllib.request.urlopen(req, timeout=90).read()
        if data[:2] == b"\x1f\x8b":
            buf = io.BytesIO(data)
            df = pd.read_csv(buf, compression="gzip", dtype=str, encoding="utf-8-sig")
        else:
            head = data[:200].decode("utf-8", "replace").lower()
            if "<html" in head:
                raise RuntimeError("데이터 링크가 파일을 직접 내려주지 않습니다. "
                                   "구글드라이브 '링크가 있는 모든 사용자' 공유로 설정하세요.")
            df = pd.read_csv(io.BytesIO(data), dtype=str, encoding="utf-8-sig")
    elif os.path.exists(CSV_GZ):
        df = pd.read_csv(CSV_GZ, compression="gzip", dtype=str, encoding="utf-8-sig")
    else:
        df = pd.read_csv(os.path.join(_HERE, "swim_records.csv"),
                         dtype=str, encoding="utf-8-sig")
    df = df.fillna("")
    # 정규화
    e = df["세부종목"].str.replace("M", "m", regex=False)
    df = df[~e.str.contains("비등록")].copy()
    e = df["세부종목"].str.replace("M", "m", regex=False)
    k = df["종별"].astype(str)
    # 구분 복원
    g = df["구분"].astype(str)
    bad = ~g.isin(["남자", "여자", "혼성"])
    g = g.mask(bad & k.str.startswith("남자"), "남자")
    g = g.mask(bad & k.str.startswith("여자"), "여자")
    g = g.mask(~g.isin(["남자", "여자", "혼성"]), "")
    df["구분"] = g
    # 종별f: 소년체전 4학년이하 → 유년부
    kf = k.where(~(e.str.contains("4학년이하") & k.str.contains("초등부")),
                 k.str.replace("초등부", "유년부", regex=False))
    df["세부종목"] = e
    df["_종목f"] = e.str.replace(r"\[4학년이하부\]", "", regex=True)
    df["_종별f"] = kf
    df["_년도"] = df["대회기간"].str.slice(0, 4)
    pid = df["선수ID"].astype(str)
    df["_pid"] = pid.where(pid != "", df["이름"] + "|" + df["소속"])
    # 메모리 절약: category 변환
    for col in ["종별", "_종별f", "_종목f", "_년도", "구분", "라운드", "대회명"]:
        df[col] = df[col].astype("category")
    return df


@st.cache_data(show_spinner=False, max_entries=4)
def rankmaps(stroke, gender):
    """(년도,종별f,종목f)별 연간 종합순위. dict[(년도,종별f,종목f,pid)] = rank"""
    df = load_df()
    sub = df[df["_종목f"].astype(str).str.startswith(stroke) & (df["구분"] == gender)].copy()
    sub = sub[sub["_년도"].astype(str).str.isdigit()]
    # 행별 best 초
    t1 = sub["예선기록"].map(parse_time)
    t2 = sub["기록"].map(parse_time)
    sub["_sec"] = pd.concat([t1, t2], axis=1).min(axis=1)
    sub = sub.dropna(subset=["_sec"])
    # (년도,종별f,종목f,pid)별 최소 기록
    g = sub.groupby(["_년도", "_종별f", "_종목f", "_pid"], observed=True)["_sec"].min().reset_index()
    g["rank"] = g.groupby(["_년도", "_종별f", "_종목f"], observed=True)["_sec"].rank(method="min").astype(int)
    rm = {}
    yr = g["_년도"].astype(str).tolist()
    kd = g["_종별f"].astype(str).tolist()
    ev = g["_종목f"].astype(str).tolist()
    pd_ = g["_pid"].astype(str).tolist()
    rk = g["rank"].astype(int).tolist()
    for i in range(len(g)):
        rm[(yr[i], kd[i], ev[i], pd_[i])] = rk[i]
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


df = load_df()
st.title("🏊 수영 대회 기록 뷰어")
st.caption(f"총 {len(df):,}건 · 대회 {df['대회명'].nunique()}개")

tab1, tab2, tab3, tab4 = st.tabs(["대회 목록", "기록 보기", "종합 순위", "성장 추적"])

# ───────── 대회 목록 ─────────
with tab1:
    grp = df.groupby("대회명", observed=True)
    comp = pd.DataFrame({
        "년도": grp["_년도"].first().astype(str),
        "대회기간": grp["대회기간"].first(),
        "개최장소": grp["개최장소"].first(),
        "기록수": grp.size(),
        "종목수": grp["_종목f"].nunique(),
    }).reset_index().sort_values("대회기간", ascending=False)
    st.dataframe(comp, use_container_width=True, hide_index=True)

# ───────── 기록 보기 ─────────
with tab2:
    years = ["(전체)"] + sorted([y for y in df["_년도"].astype(str).unique() if y], reverse=True)
    c1, c2, c3 = st.columns(3)
    fy = c1.selectbox("년도", years, key="v_y")
    pool = df if fy == "(전체)" else df[df["_년도"].astype(str) == fy]
    comps = ["(전체)"] + sorted(pool["대회명"].astype(str).unique())
    fc = c2.selectbox("대회", comps, key="v_c")
    txt = c3.text_input("검색 (이름/소속)", key="v_t")
    c4, c5, c6 = st.columns(3)
    fk = c4.multiselect("종별", sorted(df["_종별f"].astype(str).unique()), key="v_k")
    fe = c5.multiselect("세부종목", sorted(df["_종목f"].astype(str).unique()), key="v_e")
    frd = c6.multiselect("라운드", ["예선", "결승"], key="v_r")

    m = pd.Series(True, index=df.index)
    if fy != "(전체)":
        m &= df["_년도"].astype(str) == fy
    if fc != "(전체)":
        m &= df["대회명"].astype(str) == fc
    if fk:
        m &= df["_종별f"].astype(str).isin(fk)
    if fe:
        m &= df["_종목f"].astype(str).isin(fe)
    if frd:
        m &= df["라운드"].astype(str).isin(frd)
    if txt:
        m &= (df["이름"].str.contains(txt, na=False) | df["소속"].str.contains(txt, na=False))
    res = df.loc[m, VIEW_COLS]
    st.write(f"**{len(res):,}건**")
    st.dataframe(res.head(5000), use_container_width=True, hide_index=True)
    if len(res):
        st.download_button("엑셀로 추출",
                           to_excel_bytes(VIEW_COLS, res.values.tolist()),
                           "기록.xlsx", key="v_dl")

# ───────── 종합 순위 ─────────
def promote(kind, grade):
    rules = [("유년부", 4, "초등부", "5"), ("초등부", 6, "중학부", "1"),
             ("중학부", 3, "고등부", "1"), ("고등부", 3, "대학부", "1"),
             ("대학부", 4, "일반부", "")]
    if not str(grade).isdigit():
        return kind, grade
    gg = int(grade)
    for tok, mx, nxt, st0 in rules:
        if tok in kind:
            return (kind, str(gg + 1)) if gg + 1 <= mx else (kind.replace(tok, nxt), st0)
    return kind, grade


with tab3:
    yrs = sorted([y for y in df["_년도"].astype(str).unique() if y], reverse=True)
    yopts = ([str(int(yrs[0]) + 1) + "(예상)"] + yrs) if yrs else []
    c1, c2, c3 = st.columns(3)
    ry = c1.selectbox("년도", yopts, key="r_y")
    rk_kind = c2.selectbox("종별", ["(선택)"] + sorted(df["_종별f"].astype(str).unique()), key="r_k")
    rk_ev = c3.selectbox("세부종목", ["(선택)"] + sorted(df["_종목f"].astype(str).unique()), key="r_e")
    st.caption("그 해 모든 대회의 개인 최고기록 순위. '(예상)'은 전년도 기록을 진급 종별로 재배치.")
    if rk_kind != "(선택)" and rk_ev != "(선택)" and ry:
        pred = ry.endswith("(예상)")
        base = str(int(ry[:4]) - 1) if pred else ry
        sub = df[(df["_년도"].astype(str) == base) & (df["_종목f"].astype(str) == rk_ev)].copy()
        if pred:
            kg = sub.apply(lambda r: promote(str(r["_종별f"]), r["학년"]), axis=1)
            sub["_k2"] = [x[0] for x in kg]
            sub["_g2"] = [x[1] for x in kg]
        else:
            sub["_k2"] = sub["_종별f"].astype(str)
            sub["_g2"] = sub["학년"]
        sub = sub[sub["_k2"] == rk_kind]
        if len(sub):
            sub["_sec"] = pd.concat([sub["예선기록"].map(parse_time),
                                     sub["기록"].map(parse_time)], axis=1).min(axis=1)
            sub = sub.dropna(subset=["_sec"])
            idx = sub.groupby("_pid", observed=True)["_sec"].idxmin()
            best = sub.loc[idx].sort_values("_sec").reset_index(drop=True)
            ranks, prev, pr = [], None, 0
            for i, s in enumerate(best["_sec"], 1):
                rk = pr if s == prev else i
                prev, pr = s, rk
                ranks.append(rk)
            out = pd.DataFrame({
                "순위": ranks, "이름": best["이름"], "학년": best["_g2"],
                "소속": best["소속"], "시도": best["시도"],
                "기록": best[["예선기록", "기록"]].apply(
                    lambda r: r["기록"] if parse_time(r["기록"]) == best.loc[r.name, "_sec"]
                    else r["예선기록"], axis=1),
                "달성대회": best["대회명"].astype(str),
            })
            st.write(f"**{len(out)}명**")
            st.dataframe(out, use_container_width=True, hide_index=True)
            st.download_button("엑셀로 추출", to_excel_bytes(list(out.columns), out.values.tolist()),
                               f"종합순위_{ry}_{rk_kind}_{rk_ev}.xlsx", key="r_dl")
        else:
            st.info("해당 조건의 기록이 없습니다.")
    else:
        st.info("종별과 세부종목을 선택하세요.")

# ───────── 성장 추적 ─────────
def birth_year(sub):
    offs = {"유년부": 6, "초등부": 6, "중학부": 12, "고등부": 15, "대학부": 18}
    votes = {}
    for _, r in sub.iterrows():
        g, y = r["학년"], str(r["_년도"])
        if not (str(g).isdigit() and y.isdigit()):
            continue
        for tok, off in offs.items():
            if tok in str(r["_종별f"]):
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

    st.caption("각 칸 = 그 시점 연간 종합순위(예: 1위 50m (2025)). 중·고등부는 3학년 우선, 없으면 최근 학년. "
               "'은퇴(예상)'=이후 활동 없음 / '-'=그 단계 기록만 없음 / 빈칸=아직 해당 나이 아님. "
               "국가대표 선발전은 전원 일반부로 기재됨.")

    if g_stroke != "(선택)":
        rm = rankmaps(g_stroke, g_gender)
        latest = max((int(y) for y in df["_년도"].astype(str).unique() if y.isdigit()), default=0)
        last_act = (df[df["_년도"].astype(str).str.isdigit()]
                    .assign(_y=lambda d: d["_년도"].astype(str).astype(int))
                    .groupby("_pid", observed=True)["_y"].max().to_dict())
        sub = df[df["_종목f"].astype(str).str.startswith(g_stroke) & (df["구분"] == g_gender)].copy()
        sub["_종별f"] = sub["_종별f"].astype(str)
        sub["_종목f"] = sub["_종목f"].astype(str)
        sub["_년도"] = sub["_년도"].astype(str)

        def cell_for(g, pred_mask, tag=""):
            d = g[pred_mask(g)]
            best = None
            yrs_ = d["_년도"].tolist(); kds = d["_종별f"].tolist()
            evs = d["_종목f"].tolist(); pds = d["_pid"].tolist()
            for i in range(len(d)):
                rk = rm.get((yrs_[i], kds[i], evs[i], pds[i]))
                if rk:
                    dist = evs[i][len(g_stroke):]
                    dn = int(re.sub(r"[^0-9]", "", dist) or 0)
                    if best is None or (rk, dn) < (best[0], best[1]):
                        best = (rk, dn, dist, yrs_[i])
            if not best:
                return ""
            rk, _dn, dist, y = best
            suf = " (" + ((tag + "·") if tag else "") + y + ")"
            return str(rk) + "위 " + dist + suf

        def is_elem(d, gr):
            return (d["학년"] == gr) & (d["_종별f"].str.contains("초등부") | d["_종별f"].str.contains("유년부"))

        def level_cell(g, level):
            tag = {"중학부": "중", "고등부": "고"}.get(level, "")
            for gr in ("3", "2", "1"):
                c = cell_for(g, lambda d, gg=gr: (d["학년"] == gg) & d["_종별f"].str.contains(level),
                             tag=("" if gr == "3" else tag + gr))
                if c:
                    return c
            return cell_for(g, lambda d: d["_종별f"].str.contains(level))

        def fcond(name, rk):
            mode, n = conds[name]
            if mode == "무관":
                return True
            return (rk <= n) if mode == "이내" else (rk > n)

        rows_out = []
        for pid, g in sub.groupby("_pid", observed=True):
            g4 = cell_for(g, lambda d: is_elem(d, "4"))
            g6 = cell_for(g, lambda d: is_elem(d, "6"))
            m3 = level_cell(g, "중학부")
            h3 = level_cell(g, "고등부")
            uv = cell_for(g, lambda d: d["_종별f"].str.contains("대학부"))
            gn = cell_for(g, lambda d: d["_종별f"].str.contains("일반부"))
            yrs2 = [int(y) for y in g["_년도"].unique() if y.isdigit()]
            lastyr = max(yrs2) if yrs2 else 0
            lc = cell_for(g, lambda d: d["_년도"] == str(lastyr))
            if not (fcond("초등(초4/초6)", min(rank_num(g4), rank_num(g6))) and
                    fcond("중학부", rank_num(m3)) and fcond("고등부", rank_num(h3)) and
                    fcond("대학부", rank_num(uv)) and fcond("일반부", rank_num(gn)) and
                    fcond("최근", rank_num(lc))):
                continue
            byr = birth_year(g)
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
            last = g.sort_values("대회기간").iloc[-1]
            rows_out.append((int(byr) if byr.isdigit() else 9999,
                             [byr, last["이름"], last["소속"], g4, g6, m3, h3, uv, gn, lc]))
        rows_out.sort(key=lambda x: x[0])
        cols_g = ["출생(추정)", "이름", "최근소속", "초4", "초6", "중학부(3학년)",
                  "고등부(3학년)", "대학부", "일반부", "최근 연간 종합순위"]
        data = [r for _b, r in rows_out]
        st.write(f"**{len(data)}명**")
        st.dataframe(pd.DataFrame(data, columns=cols_g), use_container_width=True, hide_index=True)
        if data:
            st.download_button("엑셀로 추출", to_excel_bytes(cols_g, data),
                               f"성장추적_{g_gender}_{g_stroke}.xlsx", key="g_dl")
    else:
        st.info("종목(영법)을 선택하세요.")
