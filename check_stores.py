"""
카페24 신규 쇼핑몰 자동 점검 스크립트
사용법:
  pip install requests pandas openpyxl
  python check_stores.py 카페24_신규쇼핑몰_수도권_6개월_리스트.xlsx
결과: 점검결과_<원본파일명>.xlsx (점수 높은 순 = 기본 스킨으로 운영 중일 가능성 높은 순)
"""
import re, sys, time
from collections import Counter
from concurrent.futures import ThreadPoolExecutor
import pandas as pd
import requests

# ▼ 본인 테스트몰에 기본 스킨을 막 적용한 상태에서, 소스(Ctrl+U)에 보이는
#   기본 문구·샘플 이미지 파일명을 여기에 몇 개 넣으면 정확도가 크게 올라갑니다.
DEFAULT_MARKERS = [
    # "샘플 배너 문구", "main_banner_sample.jpg",
]

CLOSED_WORDS = ["준비중", "준비 중", "오픈 예정", "존재하지 않는", "운영하지 않", "서비스가 중지",
                "이용이 제한", "폐쇄", "coming soon", "under construction", "만료된 도메인"]
CAFE24_SIGNS = ["cafe24", "EC_FRONT", "/ind-script/", "echosting", "EC_SDE"]
UA = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                    "(KHTML, like Gecko) Chrome/126.0 Safari/537.36"}
CATEGORY_SHEETS = ["패션", "뷰티", "인테리어·리빙", "유아", "펫", "기타"]
# 원본 시트에서 결과 파일로 가져갈 열 (노란 칸 포함 — 결과 파일에서 바로 체크 가능)
KEEP_COLS = ["업종", "우선순위", "지역", "상호", "쇼핑몰 주소", "신고일", "취급품목", "관할구",
             "판단 근거", "기본스킨 여부", "모바일 상태", "연락일", "반응", "메모"]


def analyze(url):
    r = {"접속": "", "최종주소": "", "카페24확인": "", "스킨코드": "", "업로드이미지수": 0,
         "메인상품수": 0, "기본문구": 0, "점수": 0, "판단": ""}
    try:
        res = requests.get(url, headers=UA, timeout=12, allow_redirects=True)
        res.encoding = res.apparent_encoding or "utf-8"
        html = res.text
        r["최종주소"] = res.url
    except Exception as e:
        r["접속"], r["판단"] = "접속불가", type(e).__name__
        return r
    low = html.lower()
    if res.status_code >= 400:
        r["접속"], r["판단"] = f"오류 {res.status_code}", "운영 안 함 추정"
        return r
    if len(html) < 3000 or (any(w.lower() in low for w in CLOSED_WORDS) and "product" not in low):
        r["접속"], r["판단"] = "준비중/닫힘", "운영 안 함 추정"
        return r
    r["접속"] = "정상"
    r["카페24확인"] = "O" if any(s.lower() in low for s in CAFE24_SIGNS) else "X"

    skins = Counter(re.findall(r"/(skin-[a-z]*\d+)/", html))
    r["스킨코드"] = ", ".join(k for k, _ in skins.most_common(2))
    r["업로드이미지수"] = len(set(re.findall(r"/web/upload/[^\"'\s)]+\.(?:jpg|jpeg|png|gif|webp)", html, re.I)))
    r["메인상품수"] = len(set(re.findall(r"product_no=(\d+)", html)) |
                       set(re.findall(r"/product/[^\"'/]+/(\d+)/", html)))
    r["기본문구"] = sum(1 for m in DEFAULT_MARKERS if m and m.lower() in low)

    # 점수: 높을수록 '기본 스킨 그대로 + 막 시작한' 쇼핑몰
    s = 0
    if r["카페24확인"] == "O": s += 2
    if "skin1" in r["스킨코드"]: s += 3          # 최초 생성 스킨일 가능성
    if r["업로드이미지수"] <= 3: s += 3            # 직접 올린 배너·이미지가 거의 없음
    elif r["업로드이미지수"] <= 8: s += 1
    if 0 < r["메인상품수"] <= 12: s += 2           # 상품이 적음 = 초기
    s += r["기본문구"] * 3
    r["점수"] = s
    r["판단"] = "기본스킨 유력" if s >= 8 else ("확인 필요" if s >= 5 else "커스텀 추정")
    return r


def main(path):
    frames = []
    xls = pd.ExcelFile(path)
    for sh in CATEGORY_SHEETS:
        if sh in xls.sheet_names:
            df = pd.read_excel(xls, sh)
            if len(df):
                df.insert(0, "업종", sh)
                frames.append(df)
    data = pd.concat(frames, ignore_index=True)
    urls = data["쇼핑몰 주소"].astype(str).str.strip().tolist()
    print(f"{len(urls)}곳 점검 시작 (몇 분 걸립니다)")
    t = time.time()
    results = []
    with ThreadPoolExecutor(max_workers=8) as ex:
        for i, r in enumerate(ex.map(analyze, urls), 1):
            results.append(r)
            if i % 50 == 0 or i == len(urls):
                print(f"  {i}/{len(urls)} ({time.time()-t:.0f}초)")
    res = pd.DataFrame(results)
    cols = [c for c in KEEP_COLS if c in data.columns]
    out = pd.concat([data[cols], res], axis=1)
    out = out.sort_values(["점수", "신고일"], ascending=[False, False])
    name = "점검결과_" + path.split("/")[-1].split("\\")[-1]
    with pd.ExcelWriter(name, engine="openpyxl") as w:
        out[out["판단"] == "기본스킨 유력"].to_excel(w, sheet_name="기본스킨 유력", index=False)
        out[out["판단"] == "확인 필요"].to_excel(w, sheet_name="확인 필요", index=False)
        out.to_excel(w, sheet_name="전체", index=False)
        pd.crosstab(out["업종"], out["판단"], margins=True, margins_name="합계").to_excel(w, sheet_name="업종별 요약")
    print(f"완료 ({time.time()-t:.0f}초) → {name}")
    print(out["판단"].value_counts().to_string())


if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1 else "카페24_신규쇼핑몰_수도권_6개월_리스트.xlsx")
