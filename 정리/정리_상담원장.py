# -*- coding: utf-8 -*-
"""상담원장 정리 — ../messy/상담원장_2025.xlsx (원장 시트) → 정리본/

따르는 결정 기록: d-005-상담원장-정리-규칙
  ① «종결»→«완료», «대기»→«보류»로 합친다 (CS기획팀 답)
  ② 처리시간 음수는 지우지 않고 값을 그대로 둔 채 _표시 칸에 이유를 적는다
  ③ 만족도 6~10점은 지우지 않고 _표시 칸에 이유를 적는다. 평균 만족도에서는 이 줄을 뺀다
  ④ 칸 이름은 원본 그대로 둔다 (_표시 칸 하나만 맨 끝에 더한다)
  ⑤ 고객번호 앞뒤 공백을 없앤다
그 밖의 처리(테스트 접수 제외 · 중복 · 초→분 · 표기 통일 · 만족도 미응답 등)는
d-005 이전부터 이 스크립트가 하던 그대로다. 근거는 원본의 «작성기준» 시트와 사용자 지시.

원본은 읽기만 한다. 쓰는 파일: 정리본/상담원장_정리.csv · 정리본/지운줄.csv · 정리본/변경요약.md
돌리기: python 정리/정리_상담원장.py   (cleaner 폴더 어디서 돌려도 같다. 결정 기록은 ../my-wiki/decisions/)
"""
import re
import sys
from pathlib import Path

import pandas as pd

HERE = Path(__file__).resolve().parent
WIKI = HERE.parent
SRC = (WIKI / ".." / "messy" / "상담원장_2025.xlsx").resolve()
OUT = WIKI / "정리본"
TEL_SWITCH = pd.Timestamp("2025-07-01")   # 작성기준: 이날부터 전화 채널은 초 단위
REQ_START = pd.Timestamp("2025-03-01")    # 작성기준: 재문의는 이날 접수분부터 기록

CH = {**dict.fromkeys(["전화", "유선", "전화상담", "TEL"], "전화"),
      **dict.fromkeys(["챗봇", "chatbot", "Chatbot"], "챗봇"),
      **dict.fromkeys(["앱", "모바일앱", "app", "APP"], "앱"),
      **dict.fromkeys(["이메일", "메일", "e-mail", "E-mail"], "이메일")}
ST = {"완료": "완료", "처리완료": "완료", "종결": "완료",          # ① 종결 = 완료
      "진행중": "진행중", "처리중": "진행중",
      "보류": "보류", "대기": "보류"}                              # ① 대기 = 보류
RQ = {"Y": "예", "y": "예", "예": "예", "TRUE": "예",
      "N": "아니오", "n": "아니오", "아니오": "아니오", "FALSE": "아니오"}
MARK_NEG = "처리시간 음수 — 틀린 값으로 보이나 진짜 값을 몰라 그대로 둠 (d-005 ②)"
MARK_SAT = "만족도 6~10점 — NPS를 잘못 넣은 것으로 보이나 확인 못 함, 평균 만족도에서 뺌 (d-005 ③)"


def main():
    raw = pd.read_excel(SRC, sheet_name="원장", dtype=object)
    cols = list(raw.columns)                       # ④ 칸 이름 그대로
    n = raw.copy()
    changes = []                                   # (칸, 무엇을, 바뀐 줄 마스크)

    def note(col, what, mask):
        changes.append((col, what, mask.fillna(False).astype(bool)))

    # 접수일 표기
    d_raw = n["접수일"].astype(str)
    note("접수일", "«/»·«.» 표기를 «-»로", d_raw.str.contains(r"[/.]"))
    n["접수일"] = pd.to_datetime(d_raw.str.replace(r"[/.]", "-", regex=True)).dt.strftime("%Y-%m-%d")
    day = pd.to_datetime(n["접수일"])

    # ⑤ 고객번호 앞뒤 공백
    c = n["고객번호"].astype(str)
    note("고객번호", "앞뒤 공백 없앰 (d-005 ⑤)", c != c.str.strip())
    n["고객번호"] = c.str.strip()

    # 지역 · 월정액 · 접수채널 · 상담원팀 표기
    reg = {"경기도": "경기", "서울시": "서울", "서울특별시": "서울"}
    note("지역", "«경기도»→경기, «서울시»·«서울특별시»→서울", n["지역"].isin(list(reg)))
    n["지역"] = n["지역"].replace(reg)
    note("월정액", "«33,000원» 같은 글자를 숫자로", n["월정액"].map(lambda x: isinstance(x, str)))
    n["월정액"] = n["월정액"].map(lambda x: int(re.sub(r"[^\d]", "", str(x))) if pd.notna(x) else None).astype("Int64")
    ch = n["접수채널"].str.replace(" ", "", regex=False)
    if not set(ch) <= set(CH):
        sys.exit(f"모르는 접수채널: {set(ch) - set(CH)}")
    note("접수채널", "17가지 표기를 전화·챗봇·앱·이메일 넷으로", n["접수채널"] != ch.map(CH))
    n["접수채널"] = ch.map(CH)
    note("상담원팀", "팀 이름 띄어쓰기 없앰", n["상담원팀"].str.contains(" "))
    n["상담원팀"] = n["상담원팀"].str.replace(" ", "", regex=False)

    # 처리상태 ①
    st = n["처리상태"].str.strip()
    if not set(st) <= set(ST):
        sys.exit(f"모르는 처리상태: {set(st) - set(ST)}")
    for frm in ["완료 ", "처리완료", "종결", "처리중", "대기"]:
        note("처리상태", f"«{frm}»→«{ST[frm.strip()]}»" + (" (d-005 ①)" if frm in ("종결", "대기") else ""),
             n["처리상태"] == frm)
    n["처리상태"] = st.map(ST)

    # 재문의
    rq = n["재문의"].map(lambda x: RQ[str(x).strip()] if pd.notna(x) else None)
    note("재문의", "Y·y·TRUE→예, N·n·FALSE→아니오", n["재문의"].notna() & (n["재문의"].astype(str) != rq.astype(str)))
    pre = day < REQ_START
    note("재문의", "3월 1일 전 접수인데 값이 있어 비움", rq.notna() & pre)
    rq[pre] = None
    n["재문의"] = rq

    # 만족도 — 미응답 표기는 빈칸, «N점»은 숫자, 6~10은 값 그대로 두고 표시 ③
    s_raw = n["만족도"]
    s_str = s_raw.map(lambda x: str(x).strip() if pd.notna(x) else None)
    noresp = s_str.isin(["0", "-", "미응답"])
    note("만족도", "미응답 표기(0·-·미응답)를 빈칸으로", noresp)
    note("만족도", "«5점» 같은 글자를 숫자로", s_str.fillna("").str.endswith("점"))
    sat = s_str.map(lambda s: None if not isinstance(s, str) or s in ("0", "-", "미응답") else int(float(s.replace("점", ""))))
    mark_sat = sat.map(lambda v: v is not None and not pd.isna(v) and v > 5)
    n["만족도"] = sat

    # 상담시간 — 7/1 이후 전화는 초→분, 9999·0은 빈칸
    t = n["상담시간"].astype(float)
    post_tel = (n["접수채널"] == "전화") & (day >= TEL_SWITCH)
    note("상담시간", "7월 1일 이후 전화 줄을 초→분 (÷60)", post_tel)
    t[post_tel] = t[post_tel] / 60
    bad = n["상담시간"].astype(float).isin([9999, 0])
    note("상담시간", "9999·0을 빈칸으로", bad)
    t[bad] = None
    n["상담시간"] = t.round(2)

    # ② 처리시간 음수 표시
    mark_neg = pd.to_numeric(n["처리시간"], errors="coerce") < 0

    marks = []
    for neg, s6 in zip(mark_neg, mark_sat):
        marks.append("; ".join(m for m, on in ((MARK_NEG, neg), (MARK_SAT, s6)) if on))
    n["_표시"] = marks

    # 지울 줄 고르기
    reason = pd.Series("", index=raw.index)
    reason[raw["상담원ID"] == "AG999"] = "테스트 접수 (상담원ID AG999)"
    alive = reason == ""
    dup = n.loc[alive, cols].duplicated(keep="first")
    raw_same = raw.loc[alive, cols].duplicated(keep="first")
    for i in dup[dup].index:
        reason[i] = "완전히 같은 줄 (원본 그대로 같음)" if raw_same[i] else "완전히 같은 줄 (날짜·채널 등 표기만 달랐음)"
    alive = reason == ""
    a = n[alive].sort_values(["접수번호", "최종수정일"], kind="stable")
    if a.duplicated(["접수번호", "최종수정일"], keep=False).any():
        sys.exit("같은 접수번호 · 같은 최종수정일인데 내용이 다른 줄이 있어 멈춤")
    old = a.duplicated("접수번호", keep="last")
    reason[old[old].index] = "같은 접수번호의 옛 상태 줄 (최종수정일이 더 이른 줄)"
    keep = reason == ""

    clean = n.loc[keep, cols + ["_표시"]].sort_values("접수번호", kind="stable").reset_index(drop=True)
    removed = raw.loc[~keep].copy()
    removed.insert(0, "원본행", removed.index + 2)
    removed["지운이유"] = reason[~keep]

    OUT.mkdir(exist_ok=True)
    clean.to_csv(OUT / "상담원장_정리.csv", index=False, encoding="utf-8-sig")
    removed.to_csv(OUT / "지운줄.csv", index=False, encoding="utf-8-sig")

    # 다섯 숫자 + 고객 수
    N = len(clean)
    rqv = clean["재문의"].dropna()
    sat_ok = clean.loc[~clean["_표시"].str.contains("d-005 ③", regex=False), "만족도"].dropna()
    nums = [
        ("상담 건수", f"{N:,}건", "접수번호 하나가 한 건"),
        ("평균 상담시간", f"{clean['상담시간'].mean():.2f}분", f"값 있는 {clean['상담시간'].notna().sum():,}건"),
        ("재문의율", f"{(rqv == '예').mean() * 100:.2f}%", f"예 {(rqv == '예').sum():,} ÷ 재문의가 적힌 {len(rqv):,}건"),
        ("평균 만족도", f"{sat_ok.mean():.3f}", f"응답 {len(sat_ok):,}건 (d-005 ③ 표시 줄 뺌)"),
        ("완료율", f"{(clean['처리상태'] == '완료').mean() * 100:.2f}%", f"완료 {(clean['처리상태'] == '완료').sum():,} ÷ {N:,}건"),
        ("고객 수", f"{clean['고객번호'].nunique():,}명", "고객번호 종류 수"),
    ]

    # 변경요약.md
    L = ["# 상담원장 정리 — 변경요약", "",
         "- 원본: `../messy/상담원장_2025.xlsx` 원장 시트 (읽기만 함)",
         "- 스크립트: `정리/정리_상담원장.py` · 따르는 결정: [[d-005-상담원장-정리-규칙]]",
         "- 줄 수는 모두 이 스크립트가 돌 때 센 값이다.", "",
         "## 지운 줄", "", "| 단계 | 줄 수 |", "|---|---:|", f"| 원본 | {len(raw):,} |"]
    for k, v in removed["지운이유"].value_counts().items():
        L.append(f"| − {k} | {v:,} |")
    L += [f"| **정리본** | **{N:,}** |", "",
          f"빼기 확인: {len(raw):,} − {len(removed):,} = {len(raw) - len(removed):,} → 정리본 {N:,}줄과 "
          + ("같다." if len(raw) - len(removed) == N else "**다르다.**"),
          "지운 줄은 원본 값 그대로 `지운줄.csv`에 있다 (맨 앞 `원본행`, 맨 끝 `지운이유`).", "",
          "## 값을 바꾼 것 (정리본에 남은 줄 기준)", "", "| 칸 | 무엇을 | 줄 수 |", "|---|---|---:|"]
    for col, what, m in changes:
        L.append(f"| {col} | {what} | {int((m & keep).sum()):,} |")
    L += ["", "상담시간 칸 단위는 모두 **분**이다. 7월 1일 이후 전화 줄은 원본에서 초 단위였고, 60으로 나눠 분으로 바꿨다 (작성기준 시트).", "",
          "## 값은 그대로 두고 표시한 것 (`_표시` 칸, 맨 끝)", "", "| 이유 | 줄 수 |", "|---|---:|",
          f"| {MARK_NEG} | {int(clean['_표시'].str.contains('d-005 ②', regex=False).sum()):,} |",
          f"| {MARK_SAT} | {int(clean['_표시'].str.contains('d-005 ③', regex=False).sum()):,} |", "",
          "칸 이름은 원본 24칸 그대로이고, `_표시` 칸 하나만 더했다 (d-005 ④).", "",
          "## 다섯 숫자와 고객 수", "", "| 항목 | 값 | 어떻게 셌나 |", "|---|---|---|"]
    L += [f"| {a} | {b} | {c_} |" for a, b, c_ in nums]
    (OUT / "변경요약.md").write_text("\n".join(L) + "\n", encoding="utf-8")


if __name__ == "__main__":
    main()
