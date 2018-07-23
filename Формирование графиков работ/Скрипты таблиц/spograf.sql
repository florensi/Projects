-- SPOGRAF  (Table) 
--
CREATE TABLE SPOGRAF
(
  OGRAF   NUMBER(4),
  NAME    VARCHAR2(120 BYTE),
  DLIT    NUMBER(4,2),
  OTCHET  NUMBER(1),
  BR      NUMBER(1)
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

COMMENT ON COLUMN SPOGRAF.BR IS '���-�� ������';

COMMENT ON COLUMN SPOGRAF.OGRAF IS '����� ����� �������';

COMMENT ON COLUMN SPOGRAF.NAME IS '������������ �������';

COMMENT ON COLUMN SPOGRAF.DLIT IS '������������ ����� �����';

COMMENT ON COLUMN SPOGRAF.OTCHET IS '����������� ������  (1- ������, 2 - ������������ �����)';


