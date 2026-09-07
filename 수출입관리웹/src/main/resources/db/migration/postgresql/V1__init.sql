-- =====================================================================
-- DSS 외자(수출입) 모듈 — 초기 스키마 (PostgreSQL)
-- 규약: 공통 com_*, 외자 fx_*. 모든 업무 테이블에 감사 컬럼.
-- =====================================================================

CREATE TABLE com_company (
  code        VARCHAR(20)  NOT NULL PRIMARY KEY,
  name        VARCHAR(100) NOT NULL,
  short_name  VARCHAR(50),
  kind        VARCHAR(20)  NOT NULL,          -- 고객사 / 최종고객사 / 공급사 / 기타
  biz_no      VARCHAR(20),
  active      CHAR(1)      NOT NULL DEFAULT 'Y',
  sort        INT          NOT NULL DEFAULT 0
);

CREATE TABLE com_code (
  id        BIGSERIAL PRIMARY KEY,
  category  VARCHAR(50)  NOT NULL,             -- 예: fx.구분, com.통화
  code_value VARCHAR(100) NOT NULL,
  sort      INT          NOT NULL DEFAULT 0,
  CONSTRAINT uq_com_code UNIQUE (category, code_value)
);

CREATE TABLE fx_setting (
  skey  VARCHAR(50)  NOT NULL PRIMARY KEY,
  sval  VARCHAR(500)
);

CREATE TABLE fx_lc_rule (
  company_code  VARCHAR(20) NOT NULL PRIMARY KEY,
  ship_days     INT NOT NULL,
  expiry_days   INT NOT NULL
);

CREATE TABLE fx_deal (
  id               BIGSERIAL PRIMARY KEY,
  doc_no           VARCHAR(20)  NOT NULL,
  ref_no           VARCHAR(20),
  status           VARCHAR(10)  NOT NULL DEFAULT '진행중',
  customer_code    VARCHAR(20),
  customer_pic     VARCHAR(50),
  end_user_code    VARCHAR(20),
  site             VARCHAR(100),
  model            VARCHAR(200),
  qty_num          DECIMAL(12,2),
  qty_unit         VARCHAR(10),
  kind             VARCHAR(20),
  need_date        DATE,
  po_no            VARCHAR(50),
  po_date          DATE,
  debit_no         VARCHAR(50),
  offer_no         VARCHAR(50),
  fca_date         DATE,
  latest_shipment  DATE,
  expiry_date      DATE,
  lc_no            VARCHAR(50),
  lc_request_date  DATE,
  lc_open_date     DATE,
  lc_bank          VARCHAR(50),
  lc_amend         VARCHAR(500),
  amount           DECIMAL(18,2),
  currency         CHAR(3),
  comm_rate        DECIMAL(5,2),
  comm_amount      DECIMAL(18,2),
  comm_invoice_no  VARCHAR(100),
  comm_billed      DATE,
  comm_received    DATE,
  note             TEXT,
  archive_dir      VARCHAR(300),
  owner            VARCHAR(50),
  created_at       TIMESTAMP,
  created_by       VARCHAR(50),
  updated_at       TIMESTAMP,
  updated_by       VARCHAR(50),
  deleted_at       TIMESTAMP,
  CONSTRAINT uq_fx_deal_doc_no UNIQUE (doc_no)
);
CREATE INDEX ix_fx_deal_status ON fx_deal (status);
CREATE INDEX ix_fx_deal_customer ON fx_deal (customer_code);

CREATE TABLE fx_stage_log (
  id         BIGSERIAL PRIMARY KEY,
  deal_id    BIGINT NOT NULL,
  stage_no   INT    NOT NULL,
  done_date  DATE,
  memo       VARCHAR(300),
  CONSTRAINT uq_fx_stage UNIQUE (deal_id, stage_no),
  CONSTRAINT fk_fx_stage_deal FOREIGN KEY (deal_id) REFERENCES fx_deal (id)
);

CREATE TABLE fx_flow_node (
  id        BIGSERIAL PRIMARY KEY,
  seq       INT,
  title     VARCHAR(100),
  descr     TEXT,
  tip       VARCHAR(500),
  store     VARCHAR(100),
  folder    VARCHAR(300),
  stage_no  INT,
  x         INT,
  y         INT,
  shape     VARCHAR(10),
  color     VARCHAR(10)
);

CREATE TABLE fx_flow_edge (
  id     BIGSERIAL PRIMARY KEY,
  src    BIGINT NOT NULL,
  dst    BIGINT NOT NULL,
  label  VARCHAR(50)
);

-- ---------------------------------------------------------------- 기본값
INSERT INTO fx_setting (skey, sval) VALUES
  ('company',      '㈜디에스에스'),
  ('ship_days',    '14'),
  ('expiry_days',  '28'),
  ('lc_lead',      '7'),
  ('comm_due',     '30'),
  ('horizon',      '21'),
  ('archive_path', '');

INSERT INTO com_code (category, code_value, sort) VALUES
  ('fx.구분', '일반', 0), ('fx.구분', '예비기', 1), ('fx.구분', '케이블', 2), ('fx.구분', '부품', 3), ('fx.구분', '기타', 9),
  ('com.통화', 'JPY', 0), ('com.통화', 'USD', 1), ('com.통화', 'KRW', 2), ('com.통화', 'EUR', 3),
  ('fx.수량단위', 'CH', 0), ('fx.수량단위', 'SET', 1), ('fx.수량단위', 'EA', 2), ('fx.수량단위', 'Unit', 3);

INSERT INTO com_company (code, name, short_name, kind, sort) VALUES
  ('C001', '고객사 A',    'A',  '고객사', 0),
  ('C002', '고객사 B',    'B',  '고객사', 1),
  ('V001', '공급사(해외 제조사)', 'Vendor', '공급사', 0),
  ('E001', '최종고객사 A', 'EA', '최종고객사', 0),
  ('E002', '최종고객사 B', 'EB', '최종고객사', 1);

-- 고객사별 신용장 기한 예외는 [설정] 화면에서 등록한다.
-- 예: 특정 고객사만 +7일 / +14일  ->  INSERT INTO fx_lc_rule VALUES ('C001', 7, 14);
