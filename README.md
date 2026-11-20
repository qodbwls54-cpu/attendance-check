# 근태 적법성 점검

**사이트:** https://attendance-check-ten.vercel.app/

사업장별 출근부 엑셀을 올리면 주 52시간 초과, 연장·야간·휴일근로수당 지급 대상, 휴일대체 처리 여부를 직원별로 판정하는 웹 도구입니다.

- **엑셀 원본과 판정 결과는 브라우저 안에서만 처리**하고 어디로도 보내지 않습니다.
- **담당자가 지정한 설정만 Supabase에 저장해 공유**합니다: 근태 코드 의미·시간, 공휴일 목록, 사업장 설정, 직원별 감단·교대·초단기 지정.
- Supabase를 연결하지 않으면 설정은 그 브라우저(localStorage)에만 저장됩니다.

## 구성

| 파일 | 내용 |
| --- | --- |
| `index.html` | 앱 전체 (정적 페이지) |
| `config.js` | Supabase URL·anon 키 |
| `supabase/migrations/…_att_settings.sql` | 테이블·RLS 정책 |
| `vercel.json` | Vercel 배포 설정 (`main` 푸시 시 자동 배포) |
| `templates/` | 표준 양식 2종과 생성 스크립트 |

## Supabase 설정

1. Supabase 프로젝트 → **SQL Editor**에서 `supabase/migrations/20261120000000_att_settings.sql` 내용을 실행합니다.
   (Supabase CLI를 쓰면 `supabase link` 후 `supabase db push`)
2. 담당자 이메일을 등록합니다. 등록된 사람만 설정을 읽고 쓸 수 있습니다.
   ```sql
   insert into public.att_members (email) values ('담당자@회사.com');
   ```
3. **Authentication → URL Configuration**
   - Site URL: `https://attendance-check-ten.vercel.app/`
   - Redirect URLs에도 같은 주소를 추가합니다 (로컬 테스트 시 `http://localhost:8000/` 도).
4. **Project Settings → API**에서 Project URL과 `anon` `public` 키를 `config.js`에 넣습니다.
   anon 키는 공개용 키라 저장소에 올려도 됩니다. 데이터 접근은 RLS 정책(`att_members` 등록 여부)이 막습니다.
   `service_role` 키는 절대 넣지 마세요.

로그인은 이메일 매직 링크입니다. 페이지 상단에 이메일을 넣고 메일함의 링크를 누르면 공유 설정을 불러옵니다.

## Vercel 배포

1. https://vercel.com/new 에서 **Import Git Repository** → `attendance-check` 선택 → **Deploy**.
   빌드 설정은 `vercel.json` 에 들어 있어 따로 바꿀 것이 없습니다 (`index.html`·`config.js` 만 배포).
2. 배포 주소(`https://attendance-check-ten.vercel.app/`)를 Supabase **Authentication → URL Configuration** 의
   Redirect URLs 에 추가합니다. Vercel 을 주 주소로 쓸 거면 Site URL 도 이 주소로 바꿉니다.
3. 이후 `main` 에 푸시하면 Vercel 이 자동으로 다시 배포합니다.

## 로컬 실행

```bash
python3 -m http.server 8000
```

이후 `http://localhost:8000` 에서 엽니다. `config.js`를 상대 경로로 읽으므로 파일을 더블클릭해 여는 방식(file://)이 아니라 이렇게 띄워야 합니다.

## 표준 양식

사업장마다 출근부 양식이 달라 자동 인식만으로는 오판이 생깁니다 (월 추정, 다중행 양식의 연장·야간 행 누락, 숫자가 일수인지 시간인지 불명확 등).
정확한 판정이 필요하면 아래 양식으로 옮겨 올립니다. 웹 화면 상단에서도 내려받을 수 있습니다.

| 파일 | 용도 |
| --- | --- |
| `templates/attendance-std-daily.xlsx` | ① 일별 근태표 — 직원마다 근태/근로시간/연장/야간 4줄 |
| `templates/attendance-std-clock.xlsx` | ② 출퇴근 기록 — 직원·날짜마다 한 줄 (출퇴근 기록기 자료 붙여넣기용) |
| `templates/attendance-std-status.xlsx` | ③ 근무현황 — 직원마다 한 줄(○·/·연차…) + 시간외 시트. 기존 근무현황 양식을 그대로 붙여 쓰기 좋음 |

양식을 고치려면 `templates/make_templates.py`를 수정하고 `python3 templates/make_templates.py`로 다시 만듭니다 (openpyxl 필요).

## 판정 기준

앱의 "판정 기준 · 공휴일" 탭에 정리돼 있습니다. 출근부에 적힌 내용만으로 판정하므로 최종 판단은 노무사 검토를 권합니다.
