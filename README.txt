이지스 고객 가이드 - GitHub Pages 배포 파일
1) index.html (세무일정) 2) calculator.html (간편 계산기) 3) forms.html (세무·노무 서식) 4) contract.html (근로계약서 자동 작성)
- 모든 파일을 저장소 루트에 업로드 (CNAME 포함)
- Settings > Pages > Branch: main / root, Enforce HTTPS
- 예스닉 DNS: guide CNAME -> 계정명.github.io
- 사업자번호 조회: Cloudflare Worker ALLOWED_ORIGIN = https://guide.aegistax.kr
