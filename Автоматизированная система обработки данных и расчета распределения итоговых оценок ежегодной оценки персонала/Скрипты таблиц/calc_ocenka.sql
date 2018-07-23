DROP PROCEDURE CALC_OCENKA;

CREATE OR REPLACE PROCEDURE calc_ocenka( 
                                       pFio_ocen in varchar, --ФИО оценщика
                                       pKat in varchar,      --категория персонала
                                       pFunct_g in varchar,  --функциональная группа
                                       pMin in number,        --минимальное значение рейтинга в группе
                                       pMax in number,        --максимальное значение рейтинга в группе 
                                       pA5 in number,         --5 процентов по категории А
                                       pA20 in number,        --20 процентов по категории А
                                       pB60 in number,        --60 процентов по категории В
                                       pC20 in number,        --20 процентов по категории С
                                       pC5 in number,         --5 процентов по категории С
                                       pKpe in number         --оценка по КПЕ 
) as
                                        
/*
DECLARE

pFio_ocen  varchar2(200); --ФИО оценщика
pKat  varchar2(200);      --категория персонала
pFunct_g  varchar2(200);  --функциональная группа
pMin  number;        --минимальное значение рейтинга в группе
pMax  number;        --максимальное значение рейтинга в группе 
pA5  number;         --5 процентов по категории А
pA20  number;        --20 процентов по категории А
pB60  number;        --60 процентов по категории В
pC20  number;        --20 процентов по категории С
pC5  number;          --5 процентов по категории С*/

pKol_C   number; --количество работников с оценкой С
pKol_A   number; --количество работников с оценкой A
pKol_ob  number; --количество работников подлежащих оценке
pRaznica number;
pNULLreit number; --количество записей с не проставленным рейтингом
  
BEGIN

  execute immediate 'alter session set NLS_NUMERIC_CHARACTERS='||''''||'.,'||'''';

--2.проверка на наличие записей без оценок
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
   

 dbms_output.put_line('записи без оценок, разница = '||pRaznica);
 dbms_output.put_line('кол-во С ='||pKol_C);
  dbms_output.put_line('5% от кол-ва = '||pC5);
  --проверка на наличие оценок С меньше 5%
  if (pKol_C<=pC5 and pRaznica>0)
  then 
    dbms_output.put_line('C меньше 5%');  
    --обновить автоматический рейтинг с минимальным значением, где автоматический рейтинг не проставлен на оценку С
    update ocenka set avt_reit='C'
    where (zex,tn) in (select zex, tn from (
                                             select zex, tn, efect,
                                             min(efect) over (partition by kat,funct_g,fio_ocen) a_min
                                             from ocenka o 
                                             where nvl(efect,0)!=0
                                             and upper(fio_ocen)=pFio_ocen and upper(kat)=pKat and upper(decode(funct_g,NULL,'1',funct_g))=pFunct_g and  avt_reit is NULL  and decode(nvl(kpe_ocen,0),0,0,1)=pKpe)    
                       where a_min=efect);

    commit;
    dbms_output.put_line('обновление C меньше 5%');  
 
    --выбрать количество записей 
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
      dbms_output.put_line('C меньше 5%');  
      --обновить автоматический рейтинг с минимальным значением, где автоматический рейтинг не проставлен на оценку С
      update ocenka set avt_reit='C'
      where (zex,tn) in (select zex, tn from (
                                               select zex, tn, efect,
                                               min(efect) over (partition by kat,funct_g,fio_ocen) a_min
                                               from ocenka o 
                                               where nvl(efect,0)!=0
                                               and upper(fio_ocen)=pFio_ocen and upper(kat)=pKat and upper(decode(funct_g,NULL,'1',funct_g))=pFunct_g and avt_reit is NULL  and decode(nvl(kpe_ocen,0),0,0,1)=pKpe)    
                         where a_min=efect);

      commit;
      dbms_output.put_line('обновление C меньше 5%');  
   
      --выбрать количество записей 
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

dbms_output.put_line(' разница = '||pRaznica);
 dbms_output.put_line('кол-во С ='||pKol_C);
 dbms_output.put_line('20% от кол-ва = '||pC20);
 

  --проверка на наличие С меньше 20%
  if (pKol_C<=pC20 and pRaznica>0)
  then    
      --обновить автоматический рейтинг с минимальным значением, где автоматический рейтинг не проставлен на оценку С
    update ocenka set avt_reit='B-'
    where (zex,tn) in (select zex, tn from (
                                             select zex, tn, efect,
                                             min(efect) over (partition by kat,funct_g,fio_ocen) a_min
                                             from ocenka o 
                                             where nvl(efect,0)!=0
                                             and upper(fio_ocen)=pFio_ocen and upper(kat)=pKat and upper(decode(funct_g,NULL,'1',funct_g))=pFunct_g and avt_reit is NULL and decode(nvl(kpe_ocen,0),0,0,1)=pKpe)    
                       where a_min=efect);
    commit;
    dbms_output.put_line('обновление В-');  
 
    --выбрать количество записей 
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
    dbms_output.put_line(' разница = '||pRaznica);
    dbms_output.put_line('кол-во С ='||pKol_C);
    dbms_output.put_line('20% от кол-ва = '||pC20); 
 
    --проверка на наличие С меньше 20%
    while (pKol_C<pC20 and pRaznica>0)
    loop
    dbms_output.put_line('C меньше 20%');  
      --обновить автоматический рейтинг с минимальным значением, где автоматический рейтинг не проставлен на оценку С
      update ocenka set avt_reit='B-'
      where (zex,tn) in (select zex, tn from (
                                              select zex, tn, efect,
                                              min(efect) over (partition by kat,funct_g,fio_ocen) a_min
                                              from ocenka o 
                                              where nvl(efect,0)!=0
                                              and upper(fio_ocen)=pFio_ocen and upper(kat)=pKat and upper(decode(funct_g,NULL,'1',funct_g))=pFunct_g and avt_reit is NULL and decode(nvl(kpe_ocen,0),0,0,1)=pKpe)    
                         where a_min=efect);
      commit;
      dbms_output.put_line('обновление В-');  
 
      --выбрать количество записей 
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
      dbms_output.put_line(' разница = '||pRaznica);
      dbms_output.put_line('кол-во С ='||pKol_C);
      dbms_output.put_line('20% от кол-ва = '||pC20); 
    end loop;
  end if;

    --проверка остались ли еще записи без рейтинга
    if pRaznica>0 
    then
    dbms_output.put_line('есть еще записи'); 
     
      --проверка на наличие записей с оценкой А+
      if pA5>=1
      then
        dbms_output.put_line('есть А+');       
        --есть записи с А+
        --выбор кол-во оставшихся с пустым рейтингом записей
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


         --проверка кол-во значений с максимальной эффективностью и пустым рейтингом <5%
         while (pKol_A<=pA5 and pRaznica>0)
         loop
           --если меньше 5%
           dbms_output.put_line('А+ меньше 5%');  
            --обновить автоматический рейтинг с максимальным значением, где автоматический рейтинг не проставлен на оценку A+
           update ocenka set avt_reit='A+'
           where (zex,tn) in ( select zex, tn from (
                                             select zex, tn, efect,
                                             max(efect) over (partition by kat,funct_g,fio_ocen) a_max
                                             from ocenka o 
                                             where nvl(efect,0)!=0
                                             and upper(fio_ocen)=pFio_ocen and upper(kat)=pKat and upper(decode(funct_g,NULL,'1',funct_g))=pFunct_g and avt_reit is NULL and decode(nvl(kpe_ocen,0),0,0,1)=pKpe)    
                               where a_max=efect);
           commit; 
          dbms_output.put_line('обновление А+');  

           --выбрать количество записей 
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

         --проверка кол-во значений с максимальной эффективностью и пустым рейтингом <20%
         while (pKol_A<=pA20 and pRaznica>0)
         loop
          dbms_output.put_line('А меньше 20%');  
           --обновить автоматический рейтинг с максимальным значением, где автоматический рейтинг не проставлен на оценку A+
           update ocenka set avt_reit='A'
           where (zex,tn) in ( select zex, tn from (
                                             select zex, tn, efect,
                                             max(efect) over (partition by kat,funct_g,fio_ocen) a_max
                                             from ocenka o 
                                             where nvl(efect,0)!=0
                                             and upper(fio_ocen)=pFio_ocen and upper(kat)=pKat and upper(decode(funct_g,NULL,'1',funct_g))=pFunct_g and avt_reit is NULL and decode(nvl(kpe_ocen,0),0,0,1)=pKpe)    
                               where a_max=efect);
           commit; 
           dbms_output.put_line('обновляется А');  

           --выбрать количество записей 
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
            --обновить автоматический рейтинг с максимальным значением, где автоматический рейтинг не проставлен на оценку B
           update ocenka set avt_reit='B'
           where (zex,tn) in ( select zex, tn
                               from ocenka o 
                               where nvl(efect,0)!=0
                               and upper(fio_ocen)=pFio_ocen and upper(kat)=pKat and upper(decode(funct_g,NULL,'1',funct_g))=pFunct_g and avt_reit is NULL and decode(nvl(kpe_ocen,0),0,0,1)=pKpe);    
   
           commit;

         dbms_output.put_line('если остались записи обновить В');  
         end if; 


      --нет записей с А+
      else
        dbms_output.put_line('нет записей с А+');  
        --выбор кол-во оставшихся с пустым рейтингом записей
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
         

         dbms_output.put_line('кол-во А ='||pKol_A); 
         dbms_output.put_line('А от 20% =' ||pA20); 

         --проверка кол-во значений с максимальной эффективностью и пустым рейтингом <20%
         if (pKol_A<=pA20 and pRaznica>0)
         then
            --если меньше 20%
             --обновить автоматический рейтинг с максимальным значением, где автоматический рейтинг не проставлен на оценку A
             update ocenka set avt_reit='A'
             where (zex,tn) in ( select zex, tn from (
                                               select zex, tn, efect,
                                               max(efect) over (partition by kat,funct_g,fio_ocen) a_max
                                               from ocenka o 
                                               where nvl(efect,0)!=0
                                               and upper(fio_ocen)=pFio_ocen and upper(kat)=pKat and upper(decode(funct_g,NULL,'1',funct_g))=pFunct_g and avt_reit is NULL)    
                                  where efect=pMax);
             commit; 
             dbms_output.put_line('обновляется А');  
  
             --выбрать количество записей 
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
             dbms_output.put_line('А меньше 20%');  
             --если меньше 20%
             --обновить автоматический рейтинг с максимальным значением, где автоматический рейтинг не проставлен на оценку A
             update ocenka set avt_reit='A'
             where (zex,tn) in ( select zex, tn from (
                                               select zex, tn, efect,
                                               max(efect) over (partition by kat,funct_g,fio_ocen) a_max
                                               from ocenka o 
                                               where nvl(efect,0)!=0
                                               and upper(fio_ocen)=pFio_ocen and upper(kat)=pKat and upper(decode(funct_g,NULL,'1',funct_g))=pFunct_g and avt_reit is NULL and decode(nvl(kpe_ocen,0),0,0,1)=pKpe)    
                                  where a_max=efect);
             commit; 
             dbms_output.put_line('обновляется А');  
  
             --выбрать количество записей 
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
                dbms_output.put_line('кол-во А + без рейтинга = '||pKol_A);    
    
           end loop;
        
         end if;   
 
         if (pRaznica>0)
         then
          dbms_output.put_line('если остались записи, обновляется В');  
            --обновить автоматический рейтинг с максимальным значением, где автоматический рейтинг не проставлен на оценку B
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
