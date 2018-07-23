DROP PROCEDURE CALC_OCENKA;

CREATE OR REPLACE PROCEDURE calc_ocenka( 
                                       pFio_ocen in varchar, --��� ��������
                                       pKat in varchar,      --��������� ���������
                                       pFunct_g in varchar,  --�������������� ������
                                       pMin in number,        --����������� �������� �������� � ������
                                       pMax in number,        --������������ �������� �������� � ������ 
                                       pA5 in number,         --5 ��������� �� ��������� �
                                       pA20 in number,        --20 ��������� �� ��������� �
                                       pB60 in number,        --60 ��������� �� ��������� �
                                       pC20 in number,        --20 ��������� �� ��������� �
                                       pC5 in number,         --5 ��������� �� ��������� �
                                       pKpe in number         --������ �� ��� 
) as
                                        
/*
DECLARE

pFio_ocen  varchar2(200); --��� ��������
pKat  varchar2(200);      --��������� ���������
pFunct_g  varchar2(200);  --�������������� ������
pMin  number;        --����������� �������� �������� � ������
pMax  number;        --������������ �������� �������� � ������ 
pA5  number;         --5 ��������� �� ��������� �
pA20  number;        --20 ��������� �� ��������� �
pB60  number;        --60 ��������� �� ��������� �
pC20  number;        --20 ��������� �� ��������� �
pC5  number;          --5 ��������� �� ��������� �*/

pKol_C   number; --���������� ���������� � ������� �
pKol_A   number; --���������� ���������� � ������� A
pKol_ob  number; --���������� ���������� ���������� ������
pRaznica number;
pNULLreit number; --���������� ������� � �� ������������� ���������
  
BEGIN

  execute immediate 'alter session set NLS_NUMERIC_CHARACTERS='||''''||'.,'||'''';

--2.�������� �� ������� ������� ��� ������
select sum(ob_kol1), sum(kol_C1), (sum(ob_kol1)-sum(kol_C1)) 
into pKol_ob, pKol_C, pRaznica 
from (   
       select  count(*) ob_kol1, 0 kol_C1 
       from ocenka o 
       where nvl(efect,0)!=0 and upper(fio_ocen)=pFio_ocen and upper(kat)=pKat and upper(decode(funct_g,NULL,'1',funct_g))=pFunct_g and decode(nvl(kpe_ocen,0),0,0,1)=pKpe
       union all 
       select  0 ob_kol1, count(*) kol_C1 
       from ocenka o 
       where nvl(efect,0)!=0 and upper(fio_ocen)=pFio_ocen and upper(kat)=pKat and upper(decode(funct_g,NULL,'1',funct_g))=pFunct_g and avt_reit='C' and decode(nvl(kpe_ocen,0),0,0,1)=pKpe
     );
   

 dbms_output.put_line('������ ��� ������, ������� = '||pRaznica);
 dbms_output.put_line('���-�� � ='||pKol_C);
  dbms_output.put_line('5% �� ���-�� = '||pC5);
  --�������� �� ������� ������ � ������ 5%
  if (pKol_C<=pC5 and pRaznica>0)
  then 
    dbms_output.put_line('C ������ 5%');  
    --�������� �������������� ������� � ����������� ���������, ��� �������������� ������� �� ���������� �� ������ �
    update ocenka set avt_reit='C'
    where (zex,tn) in (select zex, tn from (
                                             select zex, tn, efect,
                                             min(efect) over (partition by kat,funct_g,fio_ocen) a_min
                                             from ocenka o 
                                             where nvl(efect,0)!=0
                                             and upper(fio_ocen)=pFio_ocen and upper(kat)=pKat and upper(decode(funct_g,NULL,'1',funct_g))=pFunct_g and  avt_reit is NULL  and decode(nvl(kpe_ocen,0),0,0,1)=pKpe)    
                       where a_min=efect);

    commit;
    dbms_output.put_line('���������� C ������ 5%');  
 
    --������� ���������� ������� 
    select sum(ob_kol1), sum(kol_C1), (sum(ob_kol1)-sum(kol_C1)) 
    into pKol_ob, pKol_C, pRaznica 
    from (   
           select  count(*) ob_kol1, 0 kol_C1 
           from ocenka o 
           where nvl(efect,0)!=0 and upper(fio_ocen)=pFio_ocen and upper(kat)=pKat and upper(decode(funct_g,NULL,'1',funct_g))=pFunct_g  and decode(nvl(kpe_ocen,0),0,0,1)=pKpe
           union all 
           select  0 ob_kol1, count(*) kol_C1 
           from ocenka o 
           where nvl(efect,0)!=0 and upper(fio_ocen)=pFio_ocen and upper(kat)=pKat and upper(decode(funct_g,NULL,'1',funct_g))=pFunct_g and avt_reit='C'  and decode(nvl(kpe_ocen,0),0,0,1)=pKpe
         );
  
    while (pKol_C<pC5 and pRaznica>0)
    loop
      dbms_output.put_line('C ������ 5%');  
      --�������� �������������� ������� � ����������� ���������, ��� �������������� ������� �� ���������� �� ������ �
      update ocenka set avt_reit='C'
      where (zex,tn) in (select zex, tn from (
                                               select zex, tn, efect,
                                               min(efect) over (partition by kat,funct_g,fio_ocen) a_min
                                               from ocenka o 
                                               where nvl(efect,0)!=0
                                               and upper(fio_ocen)=pFio_ocen and upper(kat)=pKat and upper(decode(funct_g,NULL,'1',funct_g))=pFunct_g and avt_reit is NULL  and decode(nvl(kpe_ocen,0),0,0,1)=pKpe)    
                         where a_min=efect);

      commit;
      dbms_output.put_line('���������� C ������ 5%');  
   
      --������� ���������� ������� 
      select sum(ob_kol1), sum(kol_C1), (sum(ob_kol1)-sum(kol_C1)) 
      into pKol_ob, pKol_C, pRaznica 
      from (   
             select  count(*) ob_kol1, 0 kol_C1 
             from ocenka o 
             where nvl(efect,0)!=0 and upper(fio_ocen)=pFio_ocen and upper(kat)=pKat and upper(decode(funct_g,NULL,'1',funct_g))=pFunct_g  and decode(nvl(kpe_ocen,0),0,0,1)=pKpe
             union all 
             select  0 ob_kol1, count(*) kol_C1 
             from ocenka o 
             where nvl(efect,0)!=0 and upper(fio_ocen)=pFio_ocen and upper(kat)=pKat and upper(decode(funct_g,NULL,'1',funct_g))=pFunct_g and avt_reit='C' and decode(nvl(kpe_ocen,0),0,0,1)=pKpe
           );
  
    end loop;
  end if;

dbms_output.put_line(' ������� = '||pRaznica);
 dbms_output.put_line('���-�� � ='||pKol_C);
 dbms_output.put_line('20% �� ���-�� = '||pC20);
 

  --�������� �� ������� � ������ 20%
  if (pKol_C<=pC20 and pRaznica>0)
  then    
      --�������� �������������� ������� � ����������� ���������, ��� �������������� ������� �� ���������� �� ������ �
    update ocenka set avt_reit='B-'
    where (zex,tn) in (select zex, tn from (
                                             select zex, tn, efect,
                                             min(efect) over (partition by kat,funct_g,fio_ocen) a_min
                                             from ocenka o 
                                             where nvl(efect,0)!=0
                                             and upper(fio_ocen)=pFio_ocen and upper(kat)=pKat and upper(decode(funct_g,NULL,'1',funct_g))=pFunct_g and avt_reit is NULL and decode(nvl(kpe_ocen,0),0,0,1)=pKpe)    
                       where a_min=efect);
    commit;
    dbms_output.put_line('���������� �-');  
 
    --������� ���������� ������� 
    select sum(ob_kol1), sum(kol_C1), sum(ob_kol1)-sum(kol_C1)  
    into pKol_ob, pKol_c, pRaznica 
    from (   
           select  count(*) ob_kol1, 0 kol_C1 
           from ocenka o 
           where nvl(efect,0)!=0 and upper(fio_ocen)=pFio_ocen and upper(kat)=pKat and upper(decode(funct_g,NULL,'1',funct_g))=pFunct_g and decode(nvl(kpe_ocen,0),0,0,1)=pKpe
           union all 
           select  0 ob_kol1, count(*) kol_C1 
           from ocenka o 
           where nvl(efect,0)!=0 and upper(fio_ocen)=pFio_ocen and upper(kat)=pKat and upper(decode(funct_g,NULL,'1',funct_g))=pFunct_g and (avt_reit='C'or avt_reit='B-') and decode(nvl(kpe_ocen,0),0,0,1)=pKpe
         );
    dbms_output.put_line(' ������� = '||pRaznica);
    dbms_output.put_line('���-�� � ='||pKol_C);
    dbms_output.put_line('20% �� ���-�� = '||pC20); 
 
    --�������� �� ������� � ������ 20%
    while (pKol_C<pC20 and pRaznica>0)
    loop
    dbms_output.put_line('C ������ 20%');  
      --�������� �������������� ������� � ����������� ���������, ��� �������������� ������� �� ���������� �� ������ �
      update ocenka set avt_reit='B-'
      where (zex,tn) in (select zex, tn from (
                                              select zex, tn, efect,
                                              min(efect) over (partition by kat,funct_g,fio_ocen) a_min
                                              from ocenka o 
                                              where nvl(efect,0)!=0
                                              and upper(fio_ocen)=pFio_ocen and upper(kat)=pKat and upper(decode(funct_g,NULL,'1',funct_g))=pFunct_g and avt_reit is NULL and decode(nvl(kpe_ocen,0),0,0,1)=pKpe)    
                         where a_min=efect);
      commit;
      dbms_output.put_line('���������� �-');  
 
      --������� ���������� ������� 
      select sum(ob_kol1), sum(kol_C1), sum(ob_kol1)-sum(kol_C1)  
      into pKol_ob, pKol_c, pRaznica 
      from (   
             select  count(*) ob_kol1, 0 kol_C1 
             from ocenka o 
             where nvl(efect,0)!=0 and upper(fio_ocen)=pFio_ocen and upper(kat)=pKat and upper(decode(funct_g,NULL,'1',funct_g))=pFunct_g and decode(nvl(kpe_ocen,0),0,0,1)=pKpe
             union all 
             select  0 ob_kol1, count(*) kol_C1 
             from ocenka o 
             where nvl(efect,0)!=0 and upper(fio_ocen)=pFio_ocen and upper(kat)=pKat and upper(decode(funct_g,NULL,'1',funct_g))=pFunct_g and (avt_reit='C'or avt_reit='B-') and decode(nvl(kpe_ocen,0),0,0,1)=pKpe
           );
      dbms_output.put_line(' ������� = '||pRaznica);
      dbms_output.put_line('���-�� � ='||pKol_C);
      dbms_output.put_line('20% �� ���-�� = '||pC20); 
    end loop;
  end if;

    --�������� �������� �� ��� ������ ��� ��������
    if pRaznica>0 
    then
    dbms_output.put_line('���� ��� ������'); 
     
      --�������� �� ������� ������� � ������� �+
      if pA5>=1
      then
        dbms_output.put_line('���� �+');       
        --���� ������ � �+
        --����� ���-�� ���������� � ������ ��������� �������
         select sum(ob_kol1), sum(kol_A1), (sum(ob_kol1)-sum(kol_A1)) 
         into pKol_ob, pKol_A, pRaznica 
         from (   
                select  count(*) ob_kol1, 0 kol_A1 
                from ocenka o 
                where nvl(efect,0)!=0 and upper(fio_ocen)=pFio_ocen and upper(kat)=pKat and upper(decode(funct_g,NULL,'1',funct_g))=pFunct_g and decode(nvl(kpe_ocen,0),0,0,1)=pKpe
                union all 
                select  0 ob_kol1, count(*) kol_A1 
                from ocenka o 
                where nvl(efect,0)!=0 and upper(fio_ocen)=pFio_ocen and upper(kat)=pKat and upper(decode(funct_g,NULL,'1',funct_g))=pFunct_g and avt_reit is null and efect=pMax and decode(nvl(kpe_ocen,0),0,0,1)=pKpe
             );


         --�������� ���-�� �������� � ������������ �������������� � ������ ��������� <5%
         while (pKol_A<=pA5 and pRaznica>0)
         loop
           --���� ������ 5%
           dbms_output.put_line('�+ ������ 5%');  
            --�������� �������������� ������� � ������������ ���������, ��� �������������� ������� �� ���������� �� ������ A+
           update ocenka set avt_reit='A+'
           where (zex,tn) in ( select zex, tn from (
                                             select zex, tn, efect,
                                             max(efect) over (partition by kat,funct_g,fio_ocen) a_max
                                             from ocenka o 
                                             where nvl(efect,0)!=0
                                             and upper(fio_ocen)=pFio_ocen and upper(kat)=pKat and upper(decode(funct_g,NULL,'1',funct_g))=pFunct_g and avt_reit is NULL and decode(nvl(kpe_ocen,0),0,0,1)=pKpe)    
                               where a_max=efect);
           commit; 
          dbms_output.put_line('���������� �+');  

           --������� ���������� ������� 
           select sum(ob_kol1), sum(kol_A1), (sum(ob_kol1)-sum(kol_A1)) 
           into pKol_ob, pKol_A, pRaznica 
           from (   
                  select  count(*) ob_kol1, 0 kol_A1 
                  from ocenka o 
                  where nvl(efect,0)!=0 and upper(fio_ocen)=pFio_ocen and upper(kat)=pKat and upper(decode(funct_g,NULL,'1',funct_g))=pFunct_g and decode(nvl(kpe_ocen,0),0,0,1)=pKpe
                  union all 
                  select  0 ob_kol1, count(*) kol_A1 
                  from ocenka o 
                  where nvl(efect,0)!=0 and upper(fio_ocen)=pFio_ocen and upper(kat)=pKat and upper(decode(funct_g,NULL,'1',funct_g))=pFunct_g and (avt_reit='A' or avt_reit='A+') and decode(nvl(kpe_ocen,0),0,0,1)=pKpe
                  union all 
                  select 0 ob_kol1, count(*) kol_A1 from (
                                               select zex, tn, efect,
                                               max(efect) over (partition by kat,funct_g,fio_ocen) a_max
                                               from ocenka o 
                                               where nvl(efect,0)!=0
                                               and upper(fio_ocen)=pFio_ocen and upper(kat)=pKat and upper(decode(funct_g,NULL,'1',funct_g))=pFunct_g and avt_reit is NULL and decode(nvl(kpe_ocen,0),0,0,1)=pKpe)    
                                  where a_max=efect
                );
   
         end loop;

         --�������� ���-�� �������� � ������������ �������������� � ������ ��������� <20%
         while (pKol_A<=pA20 and pRaznica>0)
         loop
          dbms_output.put_line('� ������ 20%');  
           --�������� �������������� ������� � ������������ ���������, ��� �������������� ������� �� ���������� �� ������ A+
           update ocenka set avt_reit='A'
           where (zex,tn) in ( select zex, tn from (
                                             select zex, tn, efect,
                                             max(efect) over (partition by kat,funct_g,fio_ocen) a_max
                                             from ocenka o 
                                             where nvl(efect,0)!=0
                                             and upper(fio_ocen)=pFio_ocen and upper(kat)=pKat and upper(decode(funct_g,NULL,'1',funct_g))=pFunct_g and avt_reit is NULL and decode(nvl(kpe_ocen,0),0,0,1)=pKpe)    
                               where a_max=efect);
           commit; 
           dbms_output.put_line('����������� �');  

           --������� ���������� ������� 
           select sum(ob_kol1), sum(kol_A1), (sum(ob_kol1)-sum(kol_A1)) 
           into pKol_ob, pKol_A, pRaznica 
           from (   
                  select  count(*) ob_kol1, 0 kol_A1 
                  from ocenka o 
                  where nvl(efect,0)!=0 and upper(fio_ocen)=pFio_ocen and upper(kat)=pKat and upper(decode(funct_g,NULL,'1',funct_g))=pFunct_g and decode(nvl(kpe_ocen,0),0,0,1)=pKpe
                  union all 
                  select  0 ob_kol1, count(*) kol_A1 
                  from ocenka o 
                  where nvl(efect,0)!=0 and upper(fio_ocen)=pFio_ocen and upper(kat)=pKat and upper(decode(funct_g,NULL,'1',funct_g))=pFunct_g and (avt_reit='A' or avt_reit='A+' or avt_reit='C' or avt_reit='B-') and decode(nvl(kpe_ocen,0),0,0,1)=pKpe
                );


         end loop;


         if (pRaznica>0)
         then
            --�������� �������������� ������� � ������������ ���������, ��� �������������� ������� �� ���������� �� ������ B
           update ocenka set avt_reit='B'
           where (zex,tn) in ( select zex, tn
                               from ocenka o 
                               where nvl(efect,0)!=0
                               and upper(fio_ocen)=pFio_ocen and upper(kat)=pKat and upper(decode(funct_g,NULL,'1',funct_g))=pFunct_g and avt_reit is NULL and decode(nvl(kpe_ocen,0),0,0,1)=pKpe);    
   
           commit;

         dbms_output.put_line('���� �������� ������ �������� �');  
         end if; 


      --��� ������� � �+
      else
        dbms_output.put_line('��� ������� � �+');  
        --����� ���-�� ���������� � ������ ��������� �������
         select sum(ob_kol1), sum(kol_A1), (sum(ob_kol1)-sum(kol_A1)) 
         into pKol_ob, pKol_A, pRaznica 
         from (   
                  select  count(*) ob_kol1, 0 kol_A1 
                  from ocenka o 
                  where nvl(efect,0)!=0 and upper(fio_ocen)=pFio_ocen and upper(kat)=pKat and upper(decode(funct_g,NULL,'1',funct_g))=pFunct_g and decode(nvl(kpe_ocen,0),0,0,1)=pKpe
                  union all 
                  select  0 ob_kol1, count(*) kol_A1 
                  from ocenka o 
                  where nvl(efect,0)!=0 and upper(fio_ocen)=pFio_ocen and upper(kat)=pKat and upper(decode(funct_g,NULL,'1',funct_g))=pFunct_g and avt_reit is null and efect=pMax and decode(nvl(kpe_ocen,0),0,0,1)=pKpe
                );
         

         dbms_output.put_line('���-�� � ='||pKol_A); 
         dbms_output.put_line('� �� 20% =' ||pA20); 

         --�������� ���-�� �������� � ������������ �������������� � ������ ��������� <20%
         if (pKol_A<=pA20 and pRaznica>0)
         then
            --���� ������ 20%
             --�������� �������������� ������� � ������������ ���������, ��� �������������� ������� �� ���������� �� ������ A
             update ocenka set avt_reit='A'
             where (zex,tn) in ( select zex, tn from (
                                               select zex, tn, efect,
                                               max(efect) over (partition by kat,funct_g,fio_ocen) a_max
                                               from ocenka o 
                                               where nvl(efect,0)!=0
                                               and upper(fio_ocen)=pFio_ocen and upper(kat)=pKat and upper(decode(funct_g,NULL,'1',funct_g))=pFunct_g and avt_reit is NULL)    
                                  where efect=pMax);
             commit; 
             dbms_output.put_line('����������� �');  
  
             --������� ���������� ������� 
             select sum(ob_kol1), sum(kol_C1), (sum(ob_kol1)-sum(kol_C1)) 
             into pKol_ob, pKol_A, pRaznica 
             from (   
                    select  count(*) ob_kol1, 0 kol_C1 
                    from ocenka o 
                    where nvl(efect,0)!=0 and upper(fio_ocen)=pFio_ocen and upper(kat)=pKat and upper(decode(funct_g,NULL,'1',funct_g))=pFunct_g
                    union all 
                    select  0 ob_kol1, count(*) kol_C1 
                    from ocenka o 
                    where nvl(efect,0)!=0 and upper(fio_ocen)=pFio_ocen and upper(kat)=pKat and upper(decode(funct_g,NULL,'1',funct_g))=pFunct_g and avt_reit='A'
                    union all
                    select 0 ob_kol1, count(*) kol_A1 from (
                                               select zex, tn, efect,
                                               max(efect) over (partition by kat,funct_g,fio_ocen) a_max
                                               from ocenka o 
                                               where nvl(efect,0)!=0
                                               and upper(fio_ocen)=pFio_ocen and upper(kat)=pKat and upper(decode(funct_g,NULL,'1',funct_g))=pFunct_g and avt_reit is NULL)    
                                  where a_max=efect
                  );
    
           
        
        
           while (pKol_A<=pA20 and pRaznica>0)
           loop
             dbms_output.put_line('� ������ 20%');  
             --���� ������ 20%
             --�������� �������������� ������� � ������������ ���������, ��� �������������� ������� �� ���������� �� ������ A
             update ocenka set avt_reit='A'
             where (zex,tn) in ( select zex, tn from (
                                               select zex, tn, efect,
                                               max(efect) over (partition by kat,funct_g,fio_ocen) a_max
                                               from ocenka o 
                                               where nvl(efect,0)!=0
                                               and upper(fio_ocen)=pFio_ocen and upper(kat)=pKat and upper(decode(funct_g,NULL,'1',funct_g))=pFunct_g and avt_reit is NULL and decode(nvl(kpe_ocen,0),0,0,1)=pKpe)    
                                  where a_max=efect);
             commit; 
             dbms_output.put_line('����������� �');  
  
             --������� ���������� ������� 
             select sum(ob_kol1), sum(kol_A1), (sum(ob_kol1)-sum(kol_A1)) 
             into pKol_ob, pKol_A, pRaznica 
             from (   
                    select  count(*) ob_kol1, 0 kol_A1 
                    from ocenka o 
                    where nvl(efect,0)!=0 and upper(fio_ocen)=pFio_ocen and upper(kat)=pKat and upper(decode(funct_g,NULL,'1',funct_g))=pFunct_g and decode(nvl(kpe_ocen,0),0,0,1)=pKpe
                    union all 
                    select  0 ob_kol1, count(*) kol_A1 
                    from ocenka o 
                    where nvl(efect,0)!=0 and upper(fio_ocen)=pFio_ocen and upper(kat)=pKat and upper(decode(funct_g,NULL,'1',funct_g))=pFunct_g and (avt_reit='A' or avt_reit='A+' or avt_reit='C' or avt_reit='B-' ) and decode(nvl(kpe_ocen,0),0,0,1)=pKpe
                    
                  );
                dbms_output.put_line('���-�� � + ��� �������� = '||pKol_A);    
    
           end loop;
        
         end if;   
 
         if (pRaznica>0)
         then
          dbms_output.put_line('���� �������� ������, ����������� �');  
            --�������� �������������� ������� � ������������ ���������, ��� �������������� ������� �� ���������� �� ������ B
           update ocenka set avt_reit='B'
           where (zex,tn) in ( select zex, tn
                               from ocenka o 
                               where nvl(efect,0)!=0
                               and upper(fio_ocen)=pFio_ocen and upper(kat)=pKat and upper(decode(funct_g,NULL,'1',funct_g))=pFunct_g and avt_reit is NULL and decode(nvl(kpe_ocen,0),0,0,1)=pKpe);    
   
           commit;
             

         end if;

      end if; 
  
    end if;

END;
/
