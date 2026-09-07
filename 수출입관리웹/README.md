# DSS 수출입(외자) 관리 — 자바 웹판

> 공개 저장소이므로 **거래처·담당자·경로 기본값은 예시로 바꿔** 두었습니다.
> 실제 값은 실행 후 **[설정 · 기준정보]** 화면에서 입력합니다.

파이썬 데스크톱 프로그램(`../수출입관리프로그램`)을 **통합 포털(자바) 아래 웹 모듈**로 옮긴 것입니다.
업무 규칙은 `../수출입관리프로그램/spec/` 을 기준으로 그대로 구현했습니다.

| 항목 | 내용 |
|---|---|
| 스택 | Java 17 · Spring Boot 3.3 · Thymeleaf(서버 렌더링) · Spring Data JPA · Flyway |
| DB | 기본 H2 파일(`data/fxdb`) — 포털 DB로 옮길 때 `mariadb` / `postgresql` 프로필 |
| 테이블 | 공통 `com_company` `com_code` · 외자 `fx_deal` `fx_stage_log` `fx_lc_rule` `fx_setting` `fx_flow_node` `fx_flow_edge` |
| 관리번호 | `FX-YYYY-NNNN` (회사 전체 유일). `ref_no` 로 다른 모듈 건과 연결 |
| 로그인 | **없음.** 포털이 헤더로 넘겨주는 사용자를 받는다 (`SsoFilter`). 개발 중엔 `/dev-login` |

## 실행

```
빌드.bat     ← 처음 한 번 (또는 소스 수정 후).  target/fx-web.jar 생성
실행.bat     ← 서버 시작 후 브라우저가 http://localhost:8080 을 연다
```
JDK 17 과 Maven 이 필요합니다. `JAVA_HOME` 을 JDK 17 로 맞춘 뒤 `mvn -DskipTests package` 로 빌드하고
`java -jar target/fx-web.jar` 로 실행하면 됩니다.

처음 접속하면 **임시 로그인** 화면이 나옵니다(개발 모드). 로그인ID·이름·역할을 넣고 들어가면 됩니다.
DB가 없으면 자동 생성되고 예시 거래처·선택목록·업무 플로우 기본값이 들어갑니다.
실제 거래처와 고객사별 신용장 규칙은 **[설정 · 기준정보]** 에서 등록하세요.

## 화면

| 메뉴 | 내용 |
|---|---|
| 대시보드 | 진행중/임박/지연/커미션 미수령 카드 · 기한 알림 · 고객사별 현황 |
| 안건 관리 | 목록(보기 5종: 전체·발주납기·신용장·선적서류·커미션) + 상세(단계별 진행 / 전체 항목 / 비고) |
| 자료방 | 설정된 루트(로컬 또는 NAS) 탐색 · 검색 · 다운로드 · 올리기 · 새 폴더 · 이름 변경 · 삭제(휴지통) |
| 업무 플로우 | 자유 배치 플로우차트(끌기·연결·라벨·도형·색·확대축소) + 항목 클릭 시 관련 자료·안건 |
| 설정 · 기준정보 | 거래처 · 고객사별 신용장 규칙 · 기본값/알림 기준 · 선택 목록 · 자료방 경로 |

그 외: Offer Sheet 메일 문안(`/deals/mail?ids=`), 엑셀 5시트(`/deals/export.xlsx`), 신용장 자동계산 API(`/deals/lc-calc`).

## 포털 연동 (강한 통합) 시 할 일

1. **사용자 전달** — `application.yml` 의 `dss.sso.header-*` 를 포털이 쓰는 헤더 이름으로 맞추고 `dev-mode: false`.
   토큰/세션 방식이면 `auth/SsoFilter.java` 한 파일만 바꾸면 됩니다. 나머지 코드는 `UserContext.current()` 만 봅니다.
2. **DB** — `--spring.profiles.active=mariadb`(또는 postgresql) + 접속정보. Flyway 가 `db/migration/{vendor}/V1__init.sql` 로 테이블을 만듭니다.
   포털 DB에 이미 `com_company` 등 공통 테이블이 있으면 V1 에서 해당 CREATE 를 빼고 컬럼을 맞춥니다.
3. **모듈로 편입** — 포털 프로젝트에 `kr.co.dss.fx` 패키지·템플릿·정적파일을 옮기고, 포털 메뉴에 `/`, `/deals`, `/archive`, `/flow`, `/settings` 를 등록.
   레이아웃(`templates/layout.html`)의 사이드바/상단바를 포털 것으로 바꿉니다.
4. **자료방** — `dss.archive.root` 또는 설정 화면의 자료방 경로를 NAS(`//서버주소/공유폴더`)로. 웹서버가 사내망에 있어야 합니다.
5. **데이터 이관** — 이전 데스크톱판(SQLite)의 안건은 화면에서 다시 입력하거나 별도 스크립트로 옮깁니다.

## 구조

```
src/main/java/kr/co/dss/fx/
  FxWebApplication.java     시작점
  Stages.java               0~8단계 정의 (spec/stages.csv)
  auth/                     UserInfo · UserContext · SsoFilter   ← 포털 연동 지점
  domain/                   Deal · StageLog · Company · CodeItem · Setting · LcRule · FlowNode · FlowEdge
  repo/Repos.java           JPA 리포지토리
  service/                  DealService(채번·기한·단계) · AlertService(알림 6종) · ViewService(보기 5종)
                            ExcelService · MailService · ArchiveService · FlowService · MasterService · SettingService
  web/                      화면 컨트롤러 + FlowController(JSON API) + DevLoginController
src/main/resources/
  application.yml           설정 (H2 기본 / mariadb / postgresql 프로필)
  db/migration/{h2,mariadb,postgresql}/V1__init.sql
  templates/                layout · dashboard · deals/list · deals/form · deals/mail · archive · flow · settings · dev-login
  static/css/app.css  static/js/app.js  static/js/flow.js
```

## 검증 기준 (spec/rules.md §2)

고객사 전용 규칙(+7/+14일)일 때: FCA 2026-11-08 → Latest shipment **2026-11-15** / Expiry **2026-11-22** / 개설요청 기한 **2026-11-01**.
기본 규칙(+14/+28일)일 때: FCA 2026-11-08 → 2026-11-22 / 2026-12-06.
