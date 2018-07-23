-- VU_859_N  (Table) 
--
CREATE TABLE VU_859_N
(
  ZEX           NUMBER(3),
  TN            NUMBER(4),
  FIO           VARCHAR2(200 BYTE),
  N_DOGOVORA    VARCHAR2(20 BYTE),
  KOD_DOGOVORA  NUMBER(1),
  DATA_S        DATE,
  DATA_PO       DATE,
  SUM           NUMBER(10,2),
  PRIZNAK       NUMBER(1),
  MES           NUMBER(2),
  GOD           NUMBER(4),
  INN           VARCHAR2(10 BYTE)
)
TABLESPACE ASUZPL8
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


