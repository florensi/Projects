-- LOGS (Table) 
--
CREATE TABLE LOGS
(
  DT        DATE,
  DOMAIN    VARCHAR2(25 BYTE),
  USERSZPD  VARCHAR2(25 BYTE),
  PROG      VARCHAR2(20 BYTE),
  TEXT      VARCHAR2(300 BYTE),
  BLOK      NUMBER(1),
  ZEX       VARCHAR2(6 BYTE)
)
TABLESPACE USERS
PCTUSED    0
PCTFREE    10
INITRANS   1
MAXTRANS   255
STORAGE    (
            INITIAL          64K
            NEXT             1M
            MINEXTENTS       1
            MAXEXTENTS       UNLIMITED
            PCTINCREASE      0
            BUFFER_POOL      DEFAULT
           )
LOGGING 
NOCOMPRESS 
NOCACHE
NOPARALLEL
MONITORING;

COMMENT ON COLUMN LOGS.BLOK IS 'код блокировки (1- заблокирован, 2- разблокирован)';

COMMENT ON COLUMN LOGS.ZEX IS 'блокируемый цех';


