/*****************************************************************

RepoRTFM : Модуль подготовки отчетов в формате RTF.

Автор  : A.PL. simple
E-mail : aplmail@box.vsi.ru
WWW    : http://www.vsi.ru/~apl
         http://aplsimple.boom.ru   (здесь могут лежать более свежие версии)

Подробное описание см. в _usergid.rtf и в _read_me.txt

*****************************************************************/

#include <vcl.h>
#include <stdlib.h>
#include <stdio.h>
#include <algorithm>
using std::min;
using std::max;
#include "RepoRTFM.h"
#pragma hdrstop


#pragma package(smart_init)

namespace RepoRTF_innerface {

//Спецсимвол
AnsiString rtf_asSpec;
int rtf_lenSpec; //и его длина

bool rtf_RTFpic;

//максим.кол-во уровней
const rtf_maxLev=999;

//максим.длина строки для чтения/записи
const rtf_iSSize=4096;

//коды меток
const rtf_LevMarker=2;
const rtf_ParMarker=3;
AnsiString rtf_asParMarker, rtf_asLevMarker;

//входной файл
FILE *rtf_Instream;
//выходной файл
FILE *rtf_Outstream;

//Объявления функций внутреннего интерфейса RepoRTF_innerface

void __fastcall rtf_SpecSymbols(
  AnsiChar acSpec, AnsiString asRepl, AnsiString* asParvalue);

void __fastcall rtf_GetParameter(
  AnsiString* asDat, AnsiString* asParname,
  AnsiString* asParvalue, int* curlev);

void __fastcall rtf_StrIntoReport(
  AnsiString* asStr, int* kz);

void __fastcall rtf_ParIntoReport(AnsiString* rtf_StrOut,
  AnsiString asParname, AnsiString asParvalue);

void __fastcall rtf_CloseFiles(char* szErr);

AnsiString __fastcall rtf_SymbCode(AnsiString asSymb);

AnsiString __fastcall rtf_Marker(int lev);

char rtf_ErrStruc[] =
  "Ошибка в структуре файла шаблона и/или данных";
}

using namespace RepoRTF_innerface;

AnsiString rtf_SubstOld="\\fcharset0", rtf_SubstNew="\\fcharset204";

//========================================================
//Процедура обработки спец.символов формата rtf в
//значении параметра (они должны быть заменены кодами)

//Входные параметры:
//acSpec - спец.символ
//asRepl - его замена

//Входные/выходные параметры:
//asParvalue- значение параметра

void __fastcall RepoRTF_innerface::rtf_SpecSymbols(
  AnsiChar acSpec, AnsiString asRepl, AnsiString* asParvalue)
{
  for (int ib=1; ib<=(*asParvalue).Length(); ib++)
    if ((*asParvalue)[ib]==acSpec)
      *asParvalue = (*asParvalue).SubString(1,ib-1)+asRepl+
        (*asParvalue).SubString(ib+1,(*asParvalue).Length());
}

//========================================================
//Процедура получения текущего параметра из строки данных

//Входные параметры:
//asDat - строка данных

//Выходные параметры:
//asParname - имя параметра
//asParvalue - значение параметра
//curlev - уровень строки данных

void __fastcall RepoRTF_innerface::rtf_GetParameter(
  AnsiString* asDat, AnsiString* asParname,
  AnsiString* asParvalue, int* curlev)
{
int ib, ilen=(*asDat).Length();
  *curlev = 0;
  *asParname = "";
  ib = (*asDat).AnsiPos(":");
  if (ib<2)  return;
AnsiString asTmp=(*asDat).SubString(1,ib);
  if (asTmp[1]!='|') return;
   *curlev = atoi(asTmp.SubString(2,ib).c_str());
//избавляемся от первых спец.символов-разделителей параметров,
//например,"|1:"
 *asDat = (*asDat).SubString(ib+1,ilen);
//ищем конец имени параметра
 ib = (*asDat).AnsiPos(":");
 if (ib>1)
   {
//получаем имя параметра
   *asParname = (*asDat).SubString(1,ib-1);
   bool boolSpec = (*asParname).AnsiPos("_SYM")==1;
   *asParname = rtf_asParMarker+(*asParname)+rtf_asParMarker;
//ищем следующий
   *asDat = (*asDat).SubString(ib+1,ilen);
   ib=(*asDat).AnsiPos(asTmp);
//выделяем значение параметра
   if (ib)
     {
     *asParvalue = (*asDat).SubString(1,ib-1);
     *asDat = (*asDat).SubString(ib,ilen);
     }
   else
     *asParvalue = *asDat;
//для имени типа "спецсимвол" получить набор символов (шрифта Symbol)
   if (boolSpec)
     {
     asTmp = "{\\f3\\lang1033\\langfe1049\\langnp1033 ";
     for (int i=1;i<=(*asParvalue).Length();i++)
       asTmp = asTmp+
         rtf_SymbCode((*asParvalue).SubString(i,1));
     *asParvalue = asTmp+" }";
     return;  //остальные проверки не интересны
     }
   else
//надо учесть спец.символы управляющих слов формата rtf
     if (rtf_RTFpic)
       {
       rtf_SpecSymbols('\\',"\\'5C",asParvalue);
       rtf_SpecSymbols('{', "\\'7B",asParvalue);
       rtf_SpecSymbols('}', "\\'7D",asParvalue);
       }
//если значение содержит комбинацию |EOL:,
// создаем в нем прогоны строк,
//если значение содержит комбинацию |EOP:,
// создаем в нем прогоны страниц
   for (ib=1; ib;)
     {
     ib = (*asParvalue).AnsiPos("|EOL:");
     if (ib==0)
      {
      ib = (*asParvalue).AnsiPos("|EOP:");
      if (ib)
        *asParvalue = (*asParvalue).SubString(1,ib-1)+
          (rtf_RTFpic?"{\\page}":"\014")+
          (*asParvalue).SubString(ib+5,ilen);
       }
     else
       *asParvalue = (*asParvalue).SubString(1,ib-1)+
         (rtf_RTFpic?"{\\par}":"\n")+
         (*asParvalue).SubString(ib+5,ilen);
     }
   }
 else
   *asParname = "";
}

//========================================================
//Процедура выгрузки выходной строки в файл отчета

//Входные параметры:
//asStr - обрабатываемая строка

//Выходные параметры:
//kz - код завершения

void __fastcall RepoRTF_innerface::rtf_StrIntoReport(
  AnsiString* asStr, int* kz)
{
int ib,ib1,ib2,ib3;
unsigned char ch;
//удаляем из выходной строки все незаполненные параметры и уровни
//(это могут быть и комментарии)
  for (ib=1; ib<=(ib1=(*asStr).Length()); )
    {
    ch = (*asStr)[ib];
    if (((ch==rtf_LevMarker))||(ch==rtf_ParMarker))
      {
      if (ch==rtf_LevMarker)
        ib3=4;
      else
        ib3=1;
      ib2 = ((*asStr).SubString(ib+1,ib1)).AnsiPos(
        (*asStr).SubString(ib,ib3));
      (*asStr) = (*asStr).SubString(1,ib-1)+
        (*asStr).SubString(ib+ib2+ib3,ib1);
      }
    else
      ib++;
    }
//пишем выходную строку в файл отчета
 for (ib=1; (*kz==0)&&ib;)
   {
   ib = (*asStr).Length();
   if (ib)
     {
     ib1 = min(rtf_iSSize,ib);
     ib2 = fwrite((*asStr).SubString(1,rtf_iSSize).c_str(),
           ib1,1,rtf_Outstream);
     if (ib2==0) *kz=6;
     (*asStr) = (*asStr).SubString(rtf_iSSize+1,ib);
     }
   }
 (*asStr) = "";
}

//========================================================
//Процедура записи значения параметра в выходную строку

//Входные параметры:
//rtf_StrOut - обрабатываемая строка
//asParname  - название параметра
//asParvalue - значение параметра

//Выходные параметры:
//rtf_StrOut - обрабатываемая строка

void __fastcall RepoRTF_innerface::rtf_ParIntoReport(
  AnsiString* rtf_StrOut, AnsiString asParname,
  AnsiString asParvalue)
{
//одноименные параметры полосы детализации в шаблоне -
//одно значение в данных (как в test8)
   for (int ib=1; ib;)
     if ((ib=(*rtf_StrOut).AnsiPos(asParname))>0)
       (*rtf_StrOut) = (*rtf_StrOut).SubString(1,ib-1)+
        asParvalue+(*rtf_StrOut).SubString(ib+
        asParname.Length(),(*rtf_StrOut).Length());
}

//========================================================
// Закрытие файлов

void __fastcall RepoRTF_innerface::rtf_CloseFiles(char* szErr)
{
 fclose(rtf_Instream);
 fclose(rtf_Outstream);
 if (strlen(szErr)) throw RepoRTF_Error(szErr);
}

//========================================================
// маркер уровня

AnsiString __fastcall RepoRTF_innerface::rtf_Marker(int lev)
{
  return rtf_asLevMarker+(AnsiString(lev)+"**").SubString(1,3);
}

//========================================================
// получение кода символа в 16-ричном виде

AnsiString __fastcall RepoRTF_innerface::rtf_SymbCode
 (AnsiString asSymb)
{
int ic=WideChar(asSymb[1])%256;
  return "\\'"+AnsiString::IntToHex(ic,2);
}

//========================================================
//Процедура создания отчета из файлов данных и шаблона

void __fastcall RepoRTF_interface::rtf_CreateReport
(
const AnsiString asDatFileName, //имя файла данных
const AnsiString asPicFileName, //имя файла шаблона
const AnsiString asRepFileName, //имя файла отчета
callingProc cProc, //имя процедуры термометра
int* istrd         //номер текущей строки в файле данных
)
{
AnsiString rtf_StrInp;   //исходная строка
AnsiString rtf_StrOut;   //выходная строка
char szIO[rtf_iSSize+1]; //строка для вв/выв
char *pos;               //адрес позиции символа в строке
AnsiString asDat;        //строка файла данных
AnsiString asParname;    //имя параметра
AnsiString asParvalue;   //значение параметра
int curlev;            //номер текущего уровня данных
int prevlev=0;             //номер предыдущего уровня данных
int kz=0;                //код завершения
bool boolDOSdata;        //флаг "данные в кодировке DOS"
bool boolLev;            //флаг "найден уровень"

AnsiString asTmp;        //рабочие переменные
int ib, ib0, ib1, ib2, ilen, ilen2;
unsigned char ch;

  if (asPicFileName==asRepFileName)
   throw RepoRTF_Error("Имена файлов шаблона и отчета совпадают.");

  *istrd=0;
  rtf_asParMarker = AnsiString(char(rtf_ParMarker));
  rtf_asLevMarker = AnsiString(char(rtf_LevMarker));
  rtf_asSpec="#";
  rtf_lenSpec=1;

//Открытие файлов
//пытаемся открыть файл шаблона и прочитать его
 rtf_Instream = fopen(asPicFileName.c_str(),"r");
 if (rtf_Instream==NULL)
   throw RepoRTF_Error("Не открыт файл шаблона");
 rtf_StrInp = "";

//Чтение файла шаблона в исходную строку
 while (!feof(rtf_Instream))
   {
   ib = fread(szIO,1,rtf_iSSize,rtf_Instream);
   szIO[ib]=0;
   rtf_StrInp = rtf_StrInp+AnsiString(szIO);
   }

//пытаемся открыть файл отчета
 rtf_Outstream = fopen(asRepFileName.c_str(),"w");
 if (rtf_Outstream==NULL)
   {
   fclose(rtf_Instream);
   throw RepoRTF_Error("Не открыт файл отчета.\n\
ВОЗМОЖНО НЕ ЗАКРЫТ СТАРЫЙ ОТЧЕТ.");
   }
 fclose(rtf_Instream);

//пытаемся открыть файл данных
 rtf_Instream = fopen(asDatFileName.c_str(),"r");
 if (rtf_Instream==NULL)
   throw RepoRTF_Error("Не открыт файл данных");

//ПРОВЕРКА ШАБЛОНА НА RTF - НУЖНА, ЕСЛИ БУДУТ СОЗДАВАТЬСЯ
//ОТЧЕТЫ В ТЕКСТОВОМ ФОРМАТЕ
 rtf_RTFpic = (rtf_StrInp.AnsiPos("{\\rtf")==1);

//вставка 16.11.2002 :
//заменяем подстроки старые на новые (например, набор символов -
//да мало ли чего?)
 asParname  = rtf_SubstOld;
 asParvalue = rtf_SubstNew;
 for (;asParname.Length();)
   {
   ib = asParname.AnsiPos(rtf_asSpec);
   if (!ib)
    ib = ib0 = rtf_iSSize;
   else
    ib0 = asParvalue.AnsiPos(rtf_asSpec);
   asTmp = asParname.SubString(1,ib-1);
   asParname = asParname.SubString(ib+rtf_lenSpec,rtf_iSSize);
   asDat = asParvalue.SubString(1,ib0-1);
   asParvalue = asParvalue.SubString(ib0+rtf_lenSpec,rtf_iSSize);
   if (asDat!=asTmp)
    while ((ib=rtf_StrInp.AnsiPos(asTmp))>0)
     rtf_StrInp = rtf_StrInp.SubString(1,ib-1)+asDat+
       rtf_StrInp.SubString(ib+asTmp.Length(),rtf_StrInp.Length());
   }

//подготовка исходной строки
 for (ib=ib0=1; ib;)
   {
   ilen = rtf_StrInp.Length();
   ib = rtf_StrInp.SubString(ib0,ilen).AnsiPos(rtf_asSpec);
   if (ib)
     {
     ib=ib0=ib0+ib-1;
//нашли # - второй # обязан быть, иначе ошибка
     ib1 = rtf_StrInp.SubString(ib+rtf_lenSpec,ilen).AnsiPos(rtf_asSpec);
     if (!ib1)
       rtf_CloseFiles(rtf_ErrStruc);
//выделяем имя параметра
     asTmp = rtf_StrInp.SubString(ib+rtf_lenSpec,ib1-1);
     if (asTmp.AnsiPos("FONT=")==1) //загрузка фонта
        {
        asTmp = asTmp.SubString(6,rtf_iSSize);
        AddFontResource(asTmp.c_str());
        }
     ilen2 = asTmp.Length();
//здесь хитрость : в имя могут влезть доп.управляющие слова,
//прогоны строк и т.п. - всё это надо убрать (а комментарии
//пропустить сразу)
     ib2 = ib1+ib+rtf_lenSpec*2-1;
     if (ilen2)
      {
      if (asTmp[1]=='@') // замена спецсимвола
        {
        rtf_asSpec = asTmp.SubString(2,ilen2);
        rtf_lenSpec = rtf_asSpec.Length();
        rtf_StrInp = rtf_StrInp.SubString(1,ib-1)+
         rtf_StrInp.SubString(ib1+ib+ilen2-rtf_lenSpec,ilen);
        continue;
        }
//символы \' для русских букв (в комментариях) отсекаем сразу
      if (asTmp.AnsiPos("\\'")==0)
       for (ib1=1; ib1<=ilen2;)
        {
        ch = asTmp[ib1];
        if (!(isalnum(ch)||(ch=='_')))  //алфавитно-цифровой символ?
         {                         //не он :
         if ((ch=='{')||(ch=='\\'))
           {
           ilen2= asTmp.AnsiPos(" ");
           if (ilen2==0)
             ilen2= asTmp.AnsiPos("}");//удалим подстроку
           if (ilen2==0) ilen2=ib1; //если надо (27.01.03)
           }
         else
           ilen2 = ib1;             //или один символ
         asTmp = asTmp.SubString(1,ib1-1)+
           asTmp.SubString(ilen2+1,rtf_iSSize);
         ilen2 = asTmp.Length();
         }
        else
         ib1++;
        }
//Вместо #D1#,#D2#...,#8D#,#9D# в исходную строку пишутся символы с
//кодами rtf_LevMarker+1,2,... метки уровней детализации)
      ib1=asTmp.Length();
      boolLev=ib1;
      for (int c=0,i=1; i<=ib1; i++)
        {
        asParvalue=asTmp.SubString(i,1);
        if (asParvalue=="D")
          {c++; if ((c>1)||((i>1)&&(i<ib1))) {boolLev=false; break;}}
        else
          if ((asParvalue<"0")||(asParvalue>"9")) {boolLev=false; break;}
        }
      if (boolLev)
        {
//Разделение на D1 и 1D, D2 и 2D и т.д. не обязательно (можно
//было бы обойтись и парами 1D-1D,2D-2D,...) и даже в этом
//месте сказывается на эффективности программы, но сделано
//для наглядности и удобства создания полос детализации
        if (asTmp[1]=='D')
//начало блока
         curlev = atoi(asTmp.SubString(2,7).c_str());
        else
//конец блока
         curlev = atoi(asTmp.c_str());
        asTmp = rtf_StrInp.SubString(ib2,ilen);
//вынесем метку блока к правой скобке '}', избавляясь от \par
        if (rtf_RTFpic)
          {
          ib2 = asTmp.AnsiPos("}");
          if (ib2==1)
            {
            asTmp = asTmp.SubString(2,ilen);
            ib2 = asTmp.AnsiPos("}");
            }
          ib1 = asTmp.AnsiPos(rtf_asSpec );
          if (ib1&&(ib1<ib2)) ib2 = ib1;
          }
        else
          ib2 = 1;
        if (curlev>0)
          asParvalue = rtf_Marker(curlev);
        else
          asParvalue = "";
        rtf_StrInp = rtf_StrInp.SubString(1,ib-1)+asParvalue+
          asTmp.SubString(ib2,ilen);
        continue;
        }
      else
//Вместо # в исходную строку пишутся символы с кодом
//метки параметров
        {
//проверим: параметр не специальный ли
        if (asTmp=="DATE")// текущая дата
         asParvalue = FormatDateTime("dd.mm.yyyy",Date());
        else
        if (asTmp=="TIME")// текущее время
         asParvalue = FormatDateTime("hh:mm",Time());
        else
        if (asTmp=="EOP") // конец страницы
         asParvalue = (rtf_RTFpic?"{\\page}":"\014");
        else
         asParvalue = rtf_asParMarker+asTmp+rtf_asParMarker;
        rtf_StrInp = rtf_StrInp.SubString(1,ib-1)+
         asParvalue+rtf_StrInp.SubString(ib2,ilen);
        }
      }
     else
//наткнулись на ## - оставим только один #
      rtf_StrInp = rtf_StrInp.SubString(1,ib-1)+
         rtf_SymbCode(rtf_asSpec)+rtf_StrInp.SubString(ib2,ilen);
     }
   }

//цикл обработки файла данных
 rtf_StrOut = "";
 boolDOSdata = false;
 while (!feof(rtf_Instream))
   {
//Чтение строки файла данных
   pos = NULL;
   for (asDat=""; pos==NULL;)
     {
     szIO[0]=0;
//строка может быть и длиннее, чем rtf_iSSize
     fgets(szIO,rtf_iSSize,rtf_Instream);
     if (strlen(szIO)<1) break;
     pos = strstr(szIO,"\n");
     if (pos!=NULL) *pos = 0;
//если в файле данных задана кодировка DOS, переводим данные в
//кодировку Windows
     if (boolDOSdata) OemToAnsi(szIO,szIO);
     asDat = asDat+AnsiString(szIO);
     }

//номер тек.строки файла данных для термометра
//(и возможного анализа ошибки)
   (*istrd)++;
//и сам термометр
   if (cProc!=NULL)
     {
     kz = cProc(*istrd);      //если термометр захотел прервать
     if (kz)                  //обработку, прерываем
       rtf_CloseFiles("Создание отчета прервано");
     }

//пустые строки и комментарии не обрабатываем
   if ((asDat.Length()<1)||(asDat[1]=='#'))
     continue;

//учтем указание кодировки
   if (asDat.AnsiPos("|DOS:")==1)
     {
     boolDOSdata = true;
     continue;
     }
   if (asDat.AnsiPos("|WIN:")==1)
     {
     boolDOSdata = false;
     continue;
     }

//получение параметра из строки данных
   rtf_GetParameter(&asDat,&asParname,&asParvalue,&curlev);
   if ((curlev<0)||(curlev>rtf_maxLev))
     rtf_CloseFiles(rtf_ErrStruc);

   ilen = rtf_StrInp.Length();
//Если уровень строки файла данных равен 0, то
   if (curlev==0)
     {
     //ищем параметр в выходной строке
     ib1 = rtf_StrOut.AnsiPos(asParname);
     if (!ib1)
       {
       //если не находим, ищем в исходной
       ib1 = rtf_StrInp.AnsiPos(asParname);
       if (ib1)
         {
         //если находим, весь нулевой уровень с этим
         //параметром удаляем из исходной строки в выходную
         ib2 = ilen+1;
         for (ib=ib1; ib<ib2; ib++)
           {
           ch = rtf_StrInp[ib];
           if (ch==rtf_LevMarker)
             {         //надо проверять символ за символом,
             ib2=ib;   //чтобы не было ограничения -
             break;    //следования за 0-м непременно
             }         //1-го уровня
           }
         rtf_StrOut = rtf_StrOut+rtf_StrInp.SubString(1,ib2-1);
         rtf_StrInp = rtf_StrInp.SubString(ib2,ilen);
         }
       }
     //ВСЕ имена параметра в вых.строке заменяем на их значения
     rtf_ParIntoReport(&rtf_StrOut,asParname,asParvalue);
     }

// Если уровень строки файла данных > 0, то
   else //if (curlev>0)
     {
     //ищем уровень в шаблоне
     asTmp = rtf_Marker(curlev);
     ib1 = rtf_StrInp.AnsiPos(asTmp);
     ib2 = rtf_StrInp.SubString(ib1+4,ilen).AnsiPos(asTmp);
//вставка 23.04.02 :
//уровень не найден в шаблоне - пропускаем строку данных
     if (ib2==0) continue;
//извлекаем уровень из шаблона
     asTmp = rtf_StrInp.SubString(ib1+4,ib2-1);
//Запись уровня в выходную строку (без меток)
     if (prevlev==0)
//если был нулевой уровень, удаляем в выходную строку
//всё, что есть в ней до ПЕРВОГО уровня: считаем, что начата
//новая порция уровней
       {
       if (curlev>1)
         ib1 = rtf_StrInp.AnsiPos(rtf_Marker(1));
       if (ib1>1)
         {
         rtf_StrOut = rtf_StrOut+rtf_StrInp.SubString(1,ib1-1);
         rtf_StrInp = rtf_StrInp.SubString(ib1,ilen);
         }
       }
     rtf_StrOut = rtf_StrOut+asTmp;
//Запись значений параметров из строки данных в вых.строку
     while (asParname.Length())
       {
       rtf_ParIntoReport(&rtf_StrOut,asParname,asParvalue);
       rtf_GetParameter(&asDat,&asParname,&asParvalue,&ib1);
       }
//строка данных уровня с номером >0 обработана :
//выгружаем выходную строку в файл отчета
     rtf_StrIntoReport(&rtf_StrOut,&kz);
     if (kz)
       rtf_CloseFiles("Ошибка записи в файл отчета");
     }
//номер текущего уровня запомним, чтоб выгрузить 0-й уровень
//впоследствии
   prevlev = curlev;
   }

//остаток исходной строки надо добавить в выходную (подписи и т.п.)
 rtf_StrOut = rtf_StrOut+rtf_StrInp;

//Запись выходной строки в файл отчета
 rtf_StrIntoReport(&rtf_StrOut,&kz);
 rtf_CloseFiles(kz?"Ошибка записи в файл отчета":"");
}
