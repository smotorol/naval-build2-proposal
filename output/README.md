# 제안서 자동 생성 산출물

## 실행 방법
```powershell
py build_ppt.py
```

계획 문서만 갱신하려면 다음을 실행한다.

```powershell
py build_ppt.py --plan-only
```

## 주요 산출물
- `proposal_outline.md`: 제안서 개요
- `slide_plan.md`: 슬라이드별 목적 및 메시지
- `extracted_requirements.md`: 요구사항 대응 매트릭스
- `image_catalog.md`: PDF 이미지 후보 분류
- `assets/`: PDF 렌더링 이미지와 접촉시트
- `해군교전분석모델_BuildII_결과분석_제안서_초안.pptx`: 제안서 PPT