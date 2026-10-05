# 기업분석 PPT 만들기

투자 전에 공부하는 기업을 '구글 기업분석' 예시 덱과 같은 양식으로 정리하는 도구입니다. 양식 규칙은 [STYLE.md](STYLE.md)에 있습니다.

## 폴더 구성

```
stock-study/
├─ build_deck.py        원고(deck.yaml) → .pptx
├─ fetch_data.py        SEC 재무 데이터 + 주가 → data/data.json, data/summary.md
├─ fetch_media.py       위키백과·위키미디어·PDF 이미지를 고화질로 받기 (출처는 CREDITS.md에 기록)
├─ render_slide.py      지난 덱의 슬라이드를 이미지로 (F/U 덱의 '지난 분석 회고'용)
├─ STYLE.md             글꼴·색·슬라이드 구성 규칙 (신규 분석)
├─ STYLE_FU.md          F/U(실적 리뷰) 양식
├─ companies/<TICKER>/  종목별 원고, 데이터, 이미지
└─ output/              완성된 .pptx
```

## 사용법

```bash
pip install -r requirements.txt

# 1) 재무·주가 데이터 받기
python fetch_data.py NVDA                  # companies/NVDA/data/ 에 저장

# 2) 이미지 받기
python fetch_media.py commons "Nvidia logo.svg" companies/NVDA/img/logo.png --width 800
python fetch_media.py wiki "Jensen Huang" companies/NVDA/img/ceo.jpg

# 3) 원고 쓰기: companies/GOOGL_sample/deck.yaml을 복사해 내용 바꾸기
# 4) PPT 만들기
python build_deck.py companies/NVDA/deck.yaml   # output/NVDA_기업분석.pptx
```

## 원고(deck.yaml) 형식

```yaml
company: 엔비디아           # 표지 이름
ticker: NVDA
logo: img/logo.png         # 제목 탭 로고
cover_image: img/logo.png  # 표지 큰 이미지
data: data/data.json       # 차트 source: 에 사용
slides:
  - type: cover
    notes: 오늘 공부할 기업은 엔비디아 입니다.
  - title: 사업한줄평                       # type을 생략하면 일반 슬라이드
    lines:
      - "[[GPU]]와 [[데이터센터]] 솔루션으로 수익 창출"
      - "- 둘째 줄은 '- '로 시작"
    image: img/hq.jpg                      # images: [a.png, b.png] 로 여러 장
    source: NVIDIA 10-K                    # 오른쪽 아래 출처
    notes: |
      발표 대본…
```

| 슬라이드 종류 | 쓰는 곳 | 주요 키 |
|---|---|---|
| `cover` | 표지 | `image`, `name`, `subtitle` |
| (생략) | 일반 슬라이드 | `title`, `lines`, `image(s)`, `chart(s)`, `table`, `layout: below/right/left` |
| `segments` | 매출 구성 구조도 | `segments: [{name, items: [...]}]`, `images` |
| `section` | 흰 구역 머리 | `text`, `subtitle` |
| `qa` | 어닝콜 Q&A | `q`, `a` (문자열 또는 목록) |
| `end` | 감사합니다 | `image`, `text` |
| `fu_cover` | F/U 표지 (주제 사진 + 회사명 + 후킹 문구) | `image`, `hook: "[[$1,300]] 간다?"`, `dim`, `logo_small` |
| `full_image` | 사진 한 장 전면 (IR 자료 표지, 행사 사진, 끝 장) | `image`, `fit`, `tab: 감사합니다` |
| `quote` | CEO 한마디 (오른쪽 인물 사진 + 인용문) | `image`, `quote: [줄, …]`, `by` |
| `guidance` | 최근 분기 + 다음 분기 가이던스 차트 | `quarters`, `revenue`, `revenue_labels`, `growth`, `margin`, `eps`, `guidance_index` |

모든 슬라이드에 `date: "2026-05-10"`을 주면 오른쪽 위에 기준일이 붙습니다. 차트 `source:`에 `until: "2026-05-08"`을 주면 그날까지의 데이터만 씁니다 (실적 발표 시점 기준 F/U 덱). F/U 양식은 [STYLE_FU.md](STYLE_FU.md), 예시는 `companies/LITE_fu_sample/deck_fu.yaml`에 있습니다.

차트는 직접 값을 넣거나 `data.json`에서 자동으로 채울 수 있습니다.

```yaml
chart: {source: price_1y}                                  # 1년 주가 카드 (price_5y도 가능)
charts:                                                    # 여러 개를 나란히
  - {source: annual.revenue, n: 4, title: 매출}
  - {source: quarterly.eps_diluted, n: 8, title: 희석 EPS}
chart:                                                     # 직접 입력 + 전망 구간(살구색)
  categories: ["2023", "2024", "2025", "2026"]
  values: [323, 525, 914, 1850]
  labels: ["$323억", "$525억", "$914억", "$1,800~1,900억"]
  top_labels: ["10.5%", "15%", "23%", "40%"]              # 막대 위 노란 라벨
  future: [3]
```

자동으로 채울 수 있는 지표는 `revenue`, `gross_profit`, `operating_income`, `net_income`, `eps_diluted`, `diluted_shares`, `rnd`, `cfo`, `capex`, `fcf`, `buyback`, `dps`입니다(`annual.` 또는 `quarterly.`를 앞에 붙입니다).

표는 다음처럼 씁니다.

```yaml
table:
  style: grid            # grid(비교표) | statement(재무제표 카드)
  header: ["", "Class A", "Class C"]
  rows: [["의결권", "1표", "없음"]]
  merge: [[2, 1, 2, 2]]  # (행, 열)부터 (행, 열)까지 병합
  arrows: [1]            # 노란 화살표로 가리킬 행
```

## 확인하기

```bash
python /path/to/pptx/scripts/office/validate.py output/NVDA_기업분석.pptx
```

이 환경에서 화면으로 확인하려면 LibreOffice Impress와 한글 글꼴을 설치하세요 (`apt-get install libreoffice-impress fonts-noto-cjk`). 미리보기는 에스코어 드림 대신 다른 글꼴로 보이니 글자 폭이 조금 다를 수 있습니다.
