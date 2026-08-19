# 오프라인 초대장 (A4 단면)

## 파일
- `가브리엘라_축복미사_초대장_A4.pdf` — **인쇄용 최종본**
- `초대장_A4단면.html` — 원본 (수정 시 이 파일을 고친 뒤 아래 명령으로 PDF 재생성)
- `proof.png` — 시안 이미지

## 규격
- A4 세로 210×297mm, **1페이지 단면**, 컬러
- 여백: 좌우 20mm / 상하 12mm — 가정용 프린터 여백 안에 들어감
- 배경이 종이 끝까지 닿지 않으므로 **재단여백(bleed) 불필요**

## 인쇄 설정
- 용지: A4
- 배율: **100% / 실제 크기** (「페이지에 맞춤」 끄기 — 켜면 축소됨)
- 컬러 또는 흑백 어느 쪽이든 무방

## PDF 재생성
```
cd print
"/Applications/Google Chrome.app/Contents/MacOS/Google Chrome" \
  --headless=new --disable-gpu --no-pdf-header-footer --virtual-time-budget=12000 \
  --print-to-pdf="가브리엘라_축복미사_초대장_A4.pdf" "file://$PWD/초대장_A4단면.html"
```

## 온라인 초대장과 다른 점
- 참석 접수 폼 없음 (오프라인이라 제외)
- 지도·캘린더 버튼 없음
- 나머지 문안·식순·안내는 온라인판과 동일
