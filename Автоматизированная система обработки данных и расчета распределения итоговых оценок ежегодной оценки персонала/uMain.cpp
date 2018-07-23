//---------------------------------------------------------------------------
#define NO_WIN32_LEAN_AND_MEAN
#pragma link "EhLibADO"
#include <stdio.h>

#include <vcl.h>
#pragma hdrstop

#include "uMain.h"
#include "uDM.h"
#include "uReiting.h"
#include "uZagruzka.h"
#include "uVvod.h"
#include "uSprav.h"
#include "FuncUser.h"
#include "uLogs.h"
#include "uZameshenie.h"
#include "RepoRTFM.h"
#include "RepoRTFO.h"
#include "uData.h"
//---------------------------------------------------------------------------
#pragma package(smart_init)
#pragma link "DBGridEh"
#pragma resource "*.dfm"
TMain *Main;
AnsiString Mes[13]={"", "Январь", "Февраль", "Март", "Апрель", "Май", "Июнь", "Июль", "Август", "Сентябрь", "Октябрь", "Ноябрь","Декабрь"};
//---------------------------------------------------------------------------
__fastcall TMain::TMain(TComponent* Owner)
        : TForm(Owner)
{
}
//---------------------------------------------------------------------------
Variant toExcel(Variant AppEx,const char *Exc, int off, double data)
{
  try {
    AppEx.OlePropertyGet("Range", Exc).OlePropertyGet("Offset", off).OlePropertySet("Value", data);
  } catch(...) { ; }
}/* toExcel() */
//---------------------------------------------------------------------------
Variant toExcel(Variant AppEx,const char *Exc, int off, String data)
{
  try {
    AppEx.OlePropertyGet("Range", Exc).OlePropertyGet("Offset", off).OlePropertySet("Value", data.c_str());
  } catch(...) { ; }
}/* toExcel() */
//---------------------------------------------------------------------------
Variant  toExcel(Variant AppEx,const char *Exc, double data)
{
  try {
    AppEx.OlePropertyGet("Range", Exc).OlePropertySet("Value", data);
  } catch(...) { ; }
}/* toExcel() */

//---------------------------------------------------------------------------
Variant  toExcel(Variant AppEx,const char *Exc, int data)
{
  try {
    AppEx.OlePropertyGet("Range", Exc).OlePropertySet("Value", data);
  } catch(...) { ; }
}/* toExcel() */

//---------------------------------------------------------------------------
Variant  toExcel(Variant AppEx,const char *Exc, AnsiString data)
{
  try {
    Variant  cur = AppEx.OlePropertyGet("Range", Exc);
    cur.OlePropertySet("Value", data.c_str());
  } catch(...) { ; }
}/* toExcel() */
//---------------------------------------------------------------------------
void __fastcall TMain::FormCreate(TObject *Sender)
{
  Word Year, Month, Day;

  // Получение данных о пользователе из домена
  TStringList *SL_Groups = new TStringList();
  // TStringList *SL_Groups2 = new TStringList();

  // Получение данных о пользователе из домена
  // Переменные UserName, DomainName, UserFullName должны быть объявлены как AnsiString
  if (!GetFullUserInfo(UserName, DomainName, UserFullName))
    {
      MessageBox(Handle,"Ошибка получения данных о пользователе","Ошибка",8208);
      Application->Terminate();
      Abort();
    }

  //получение групп доступа из АД
  if (!GetUserGroups(UserName, DomainName, SL_Groups))
    {
      MessageBox(Handle,"Ошибка получения данных о пользователе","Ошибка",8208);
      Application->Terminate();
      Abort();
    }

  //проверка на доступ к услуге
  if ((SL_Groups->IndexOf("mmk-itsvc-hocn-admin")<=-1) && (SL_Groups->IndexOf("mmk-itsvc-hocn")<=-1))
    {
      MessageBox(Handle,"У вас нет прав для работы с\nпрограммой 'Оценка персонала'!!!","Права доступа",8208);
      Application->Terminate();
      Abort();
    }

 
 DecimalSeparator = '.';


  //Год для выборки данных из таблицы Ocenka
  DecodeDate(Date(),Year, Month, Day);

  if (Date()>=StrToDate("01.11."+IntToStr(Year)) && Date()<=StrToDate("31.12."+IntToStr(Year)))
    {
      god=Year+1;
      god_t=Year+1;
    }
  else
    {
      god=Year;
      god_t=Year;
    }

  DM->qOcenka->Close();
  DM->qOcenka->Parameters->ParamByName("pgod")->Value = IntToStr(god);
  DM->qOcenka->Parameters->ParamByName("pgod1")->Value = IntToStr(god);
  DM->qOcenka->Parameters->ParamByName("pgod2")->Value = IntToStr(god);
  DM->qOcenka->Parameters->ParamByName("pgod3")->Value = IntToStr(god);
  DM->qOcenka->Active=true;

  //Справочник цехов и дирекций
  DM->qSprav->Close();
  DM->qSprav->Parameters->ParamByName("pgod")->Value = god;
  DM->qSprav->Parameters->ParamByName("pgod1")->Value = god;
  try
    {
      DM->qSprav->Open();
    }
  catch(...)
    {
      Application->MessageBox("Ошибка при получении данных из справочника SP_PDIREKT","Ошибка соединения",
                              MB_OK + MB_ICONERROR);
      Application->Terminate();
      Abort();
    }

  //Год в таблице резервистов
  DM->qRezerv->Parameters->ParamByName("pgod")->Value = IntToStr(god);
  DM->qRezerv->Parameters->ParamByName("ptn")->Value = NULL;
  DM->qRezerv->Parameters->ParamByName("pshtat")->Value = NULL;
  DM->qRezerv->Active = true;


  StatusBar1->SimplePanel = true;
  StatusBar1->SimpleText ="Отчетный период: "+IntToStr(god)+" год";

   //Фильтрация автоматическая без нажатия Enter
  DBGridEh1->Style->FilterEditCloseUpApplyFilter =true;

  //Определение разрешения экрана
  AnsiString width = Screen->Width;     //ширина
  AnsiString height = Screen->Height;   //высота

 /*  //Установка автоматически растягивать грид в зависимости от разрешения
  if (width >= 1280 && height >= 1024 ||
      width >=1600 && height >= 900)
    {
      DBGridEh1->AutoFitColWidths = true;
    }
  else
    {
      DBGridEh1->AutoFitColWidths = false;
    }

 //Установка размера шрифта в зависимости от разрешения экрана
  if ( width >=1600 && height >= 900)
    {
      DBGridEh1->Font->Size = 11;
    }
  else
    {
      DBGridEh1->Font->Size = 10;
    }
    //Развернуть на весь экран окно главной формы
//  ShowWindow(Main->Handle, SW_MAXIMIZE);

  Main->WindowState = wsMaximized; */

  if (!GetMyDocumentsDir(DocPath))
    {
      MessageBox(Handle,"Ошибка доступа к папке документов","Ошибка",8208);
      Application->Terminate();
      Abort();
    }

  if (!GetTempDir(TempPath))
    {
      MessageBox(Handle,"Ошибка доступа к временной папке","Ошибка",8208);
      Application->Terminate();
      Abort();
    }

  WorkPath = DocPath + "\\Оценка персонала";
  Path = GetCurrentDir();
  FindWordPath();

  // Создание ProgressBar на StatusBar
  ProgressBar = new TProgressBar ( StatusBar1 );
  ProgressBar->Parent = StatusBar1;
  ProgressBar->Position = 0;
  ProgressBar->Left = Main->Width-ProgressBar->Width-13;//StatusBar1->Width-ProgressBar->Width-10;//StatusBar1->Panels->Items[0]->Width+StatusBar1->Panels->Items[1]->Width - ProgressBar->Width;//Width*18 + 81;
  //ProgressBar->Anchors = ProgressBar->Anchors << akRight << akTop << akLeft << akBottom;
  ProgressBar->Top = StatusBar1->Height/6;
  ProgressBar->Height = StatusBar1->Height-3;
  PostMessage(ProgressBar->Handle,0x0409,0,clRed);
  ProgressBar->Visible = false;
}
//---------------------------------------------------------------------------

//загрузка списка работников в картотеку
void __fastcall TMain::N5Click(TObject *Sender)
{
 AnsiString Sql;

  if (Application->MessageBox(("Загрузка списка работников выполняется 1 раз в год (в ноябре) \nперед формированием списков для оценки персонала. "
                               "Вы действительно хтите удалить все уже существующие данные \nпо работникам за "+IntToStr(god)+" год?\n"+
                               "Продолжить???").c_str(),
                               "Загрузка списка работников",
                               MB_YESNO + MB_ICONINFORMATION)==ID_NO)
    {
      Abort();
    }


//ЗАГРУЗКА СПРАВОЧНИКА ДИРЕКЦИЙ
//*****************************
  StatusBar1->SimpleText ="Загрузка справочника по дирекциям/управлениям...";

  //Проверка на наличие записей в справочнике дирекций
  Sql="select * from sp_direkt where god="+IntToStr(god);

  DM->qObnovlenie->Close();
  DM->qObnovlenie->SQL->Clear();
  DM->qObnovlenie->SQL->Add(Sql);
  try
    {
      DM->qObnovlenie->Open();
    }
  catch(Exception &E)
    {
      Application->MessageBox(("Возникла ошибка при получении данных из справочника дирекций (SP_DIREKT)" + E.Message).c_str(),"Ошибка",
                              MB_OK+MB_ICONERROR);
      StatusBar1->SimpleText ="Отчетный период: "+IntToStr(god)+" год";
      Abort();
    }

  if (DM->qObnovlenie->RecordCount>0)
    {
      //Удаление ранее загруженных записей в справочник дирекций
      Sql="delete from sp_direkt where god="+IntToStr(god);

      DM->qObnovlenie->Close();
      DM->qObnovlenie->SQL->Clear();
      DM->qObnovlenie->SQL->Add(Sql);

      try
        {
          DM->qObnovlenie->ExecSQL();
        }
      catch(Exception &E)
        {
          Application->MessageBox(("Возникла ошибка при удалении данных из справочника дирекций (SP_DIREKT)" + E.Message).c_str(),"Ошибка",
                                   MB_OK+MB_ICONERROR);
          StatusBar1->SimpleText ="Отчетный период: "+IntToStr(god)+" год";
          Abort();
        }
    }

  //Вставка новых записей в справочник дирекций
  Sql = "insert into sp_direkt (kod, naim, otchet, god) \
         select kod,                                    \
                naim,                                   \
                otchet,                                 \
                "+IntToStr(god)+"                       \
         from sp_direkt where god="+IntToStr(god-1);

  DM->qObnovlenie->Close();
  DM->qObnovlenie->SQL->Clear();
  DM->qObnovlenie->SQL->Add(Sql);

  try
    {
      DM->qObnovlenie->ExecSQL();
    }

  catch(Exception &E)
    {
      Application->MessageBox(("Возникла ошибка при попытке вставить данные в справочник дирекций (SP_DIREKT)" + E.Message).c_str(),"Вставка записей",
                              MB_OK + MB_ICONERROR);
      StatusBar1->SimpleText ="Отчетный период: "+IntToStr(god)+" год";
      Abort();
    }


  DM->qSprav->Requery();

  //Логи при загрузке
  InsertLog("Загрузка справочника дирекций по оценке персонала за "+IntToStr(god)+" год выполнена =)");
  DM->qLogs->Requery();


//ЗАГРУЗКА СПРАВОЧНИКА СООТВЕТСТВИЙ ДИРЕКЦИЙ И ЦЕХОВ
//***************************************************
  StatusBar1->SimpleText ="Загрузка справочника соответствий дирекций и цехов...";

  //Проверка на наличие записей в справочнике дирекций
  Sql="select * from sp_pdirekt where god="+IntToStr(god);

  DM->qObnovlenie->Close();
  DM->qObnovlenie->SQL->Clear();
  DM->qObnovlenie->SQL->Add(Sql);
  try
    {
      DM->qObnovlenie->Open();
    }
  catch(Exception &E)
    {
      Application->MessageBox(("Возникла ошибка при получении данных из справочника соответствий дирекций и цехов (SP_PDIREKT)" + E.Message).c_str(),"Ошибка",
                              MB_OK+MB_ICONERROR);
      StatusBar1->SimpleText ="Отчетный период: "+IntToStr(god)+" год";
      Abort();
    }

  if (DM->qObnovlenie->RecordCount>0)
    {
      //Удаление ранее загруженных записей в справочник соответствий дирекций и цехов
      Sql="delete from sp_pdirekt where god="+IntToStr(god);

      DM->qObnovlenie->Close();
      DM->qObnovlenie->SQL->Clear();
      DM->qObnovlenie->SQL->Add(Sql);

      try
        {
          DM->qObnovlenie->ExecSQL();
        }
      catch(Exception &E)
        {
          Application->MessageBox(("Возникла ошибка при удалении данных из справочника соответствий дирекций и цехов (SP_PDIREKT)" + E.Message).c_str(),"Ошибка",
                                   MB_OK+MB_ICONERROR);
          StatusBar1->SimpleText ="Отчетный период: "+IntToStr(god)+" год";
          Abort();
        }
    }

  //Вставка новых записей в справочник соответствий дирекций и цехов
  Sql = "insert into sp_pdirekt (zex, kod_d, naim_zex, zex_n, zex_s, funct, funct_g, god) \
         select zex,                                    \
                kod_d,                                  \
                naim_zex,                               \
                zex_n,                                  \
                zex_s,                                  \
                funct,                                  \
                funct_g,                                \
                "+IntToStr(god)+"                       \
         from sp_pdirekt where god="+IntToStr(god-1);

  DM->qObnovlenie->Close();
  DM->qObnovlenie->SQL->Clear();
  DM->qObnovlenie->SQL->Add(Sql);

  try
    {
      DM->qObnovlenie->ExecSQL();
    }

  catch(Exception &E)
    {
      Application->MessageBox(("Возникла ошибка при попытке вставить данные в справочник соответствий дирекций и цехов (SP_PDIREKT)" + E.Message).c_str(),"Вставка записей",
                              MB_OK + MB_ICONERROR);
      StatusBar1->SimpleText ="Отчетный период: "+IntToStr(god)+" год";
      Abort();
    }


  DM->qSprav->Requery();

  //Логи при загрузке
  InsertLog("Загрузка справочника соответствий дирекций и цехов по оценке персонала за "+IntToStr(god)+" год выполнена =)");
  DM->qLogs->Requery();


//ЗАГРУЗКА ОСНОВНОЙ ТАБЛИЦЫ
//*************************
  StatusBar1->SimpleText ="Загрузка списка работников для ежегодной оценки персонала...";

  //Проверка на наличие записей в таблице
  Sql="select * from Ocenka where god="+IntToStr(god);

  DM->qObnovlenie->Close();
  DM->qObnovlenie->SQL->Clear();
  DM->qObnovlenie->SQL->Add(Sql);
  try
    {
      DM->qObnovlenie->Open();
    }
  catch(Exception &E)
    {
      Application->MessageBox(("Возникла ошибка при получении данных из таблицы Ocenka" + E.Message).c_str(),"Ошибка",
                              MB_OK+MB_ICONERROR);
      StatusBar1->SimpleText ="Отчетный период: "+IntToStr(god)+" год";
      Abort();
    }

  if (DM->qObnovlenie->RecordCount>0)
    {
      //Удаление ранее загруженных записей
      Sql="delete from Ocenka where god="+IntToStr(god);

      DM->qObnovlenie->Close();
      DM->qObnovlenie->SQL->Clear();
      DM->qObnovlenie->SQL->Add(Sql);

      try
        {
          DM->qObnovlenie->ExecSQL();
        }
      catch(Exception &E)
        {
          Application->MessageBox(("Возникла ошибка при удалении данных из таблицы Ocenka" + E.Message).c_str(),"Ошибка",
                                   MB_OK+MB_ICONERROR);
          StatusBar1->SimpleText ="Отчетный период: "+IntToStr(god)+" год";
          Abort();
        }
    }

  //Вставка новых записей в таблицу
  Sql = "insert into ocenka (god,              \
                             zex,              \
                             tn,               \
                             fio,              \
                             dolg,             \
                             funct,            \
                             uu,               \
                             funct_g,          \
                             kat,              \
                             direkt,           \
                             direkt2,          \
                             uch,              \
                             dat_job,          \
                             nuch              \
                             )                 \
         (select "+IntToStr(god)+",            \
                 zex,                          \
                 tn_sap,                       \
                 initcap(fam||' '||im||' '||ot) as fio,                                                                 \
                 name_dolg_ru as dolg,                                                                                  \
                 (select naim from sp_funct where kod=(select funct from sp_pdirekt pdir where pdir.zex=o.zex and pdir.god="+IntToStr(god)+")) as funct,         \
                 ur_upr as uu,                                                                                          \
                 (select naim from sp_functg where kod=(select funct_g from sp_pdirekt pdir1 where pdir1.zex=o.zex and pdir1.god="+IntToStr(god)+")) as funct_g,     \
                 'сотрудник',                                                                                           \
                 zex as direkt,                                                                                         \
                 (select kod_d from sp_pdirekt pdir2 where pdir2.zex=o.zex and pdir2.god="+IntToStr(god)+") as direkt2,                                             \
                 (case when o.ur1 is null then o.zex                                                                    \
                       when o.ur2 is null then o.ur1                                                                    \
                       when o.ur3 is null then o.ur2                                                                    \
                       when o.ur4 is null then o.ur3                                                                    \
                  end) as uch,                                                                                          \
                  priem_date as dat_job,                                                                                \
                  name_ur1 as nuch                                                                                      \
         from sap_osn_sved o )";                                                                                          \
       //  where priem_date<=to_date('01.11."+IntToStr(god-1)+"', 'dd.mm.yyyy'))";         //Принятые на комбинат до ноября



  DM->qObnovlenie->Close();
  DM->qObnovlenie->SQL->Clear();
  DM->qObnovlenie->SQL->Add(Sql);

  try
    {
      DM->qObnovlenie->ExecSQL();
    }

  catch(Exception &E)
    {
      Application->MessageBox(("Возникла ошибка при попытке вставить данные в таблицу Ocenka" + E.Message).c_str(),"Вставка записей",
                              MB_OK + MB_ICONERROR);
      StatusBar1->SimpleText ="Отчетный период: "+IntToStr(god)+" год";
      Abort();
    }


  DM->qOcenka->Requery();

  //Логи при загрузке
  InsertLog("Загрузка ежегодного списка работников по оценке персонала за "+IntToStr(god)+" год выполнена =)");
  DM->qLogs->Requery();

  StatusBar1->SimpleText ="Отчетный период: "+IntToStr(god)+" год";

  Application->MessageBox("Загрузка ежегодного списка работников по оценке персонала выполнена =)","Загрузка данных",
                          MB_OK + MB_ICONINFORMATION);
}
//---------------------------------------------------------------------------


//Формирование списка по цехам в Excel
void __fastcall TMain::N3Click(TObject *Sender)
{
  if (god<god_t) Spisok_po_zex2016();
  else Spisok_po_zex2017();
}
//---------------------------------------------------------------------------

//Формирование списка по цехам в Excel до 2017 года
void __fastcall TMain::Spisok_po_zex2016()
{
  AnsiString sFile, Sql, zex, zex1;
  int n=18;
  double  rezult, komp, effekt;
  Variant AppEx, Sh;


  StatusBar1->SimpleText=" Идет формирование списков работников по цехам в Excel...";

  Sql="select initcap(fio) as fio, initcap(fio_ocen) as fio_ocen,                                                                                   \
              (select naim_zex from sp_pdirekt pdir where o.direkt=pdir.zex and pdir.god="+IntToStr(god)+") as naim_zex,                                      \
              (select naim from sp_direkt where god="+IntToStr(god)+" and kod = (select kod_d from sp_pdirekt pdir1 where pdir1.zex=o.direkt and pdir1.god="+IntToStr(god)+")) naim_direkt,\
              (select name_ur1 from sap_osn_sved where tn_sap=o.tn                                                   \
               union all                                                                                             \
               select name_ur1 from sap_sved_uvol where tn_sap=o.tn) as uch,                                        \
               (select distinct nazv_cex from ssap_cex where id_cex=substr(orez.zex_rez,1,2) and nazv_cex not like '%(устар.)')  naim_zex_rez,        \
               (select stext from p1000@sapmig_buffdb where otype='O' and langu='R' and short=orez.zex_rez and stext not like '%(уст%') as naim_uch_rez, \
              orez.zex_rez as zex_rez, \
              orez.shifr_rez as shifr_rez, \
              o.*                                                                                                    \
       from ocenka o left join ocenka_rez orez where o.god="+IntToStr(god)+" and orez.god="+IntToStr(god)+" and o.tn=orez.tn \                                                                                            \
       order by direkt";                                                                                                                                                  \
                                                                                                                                                                        \
  DM->qObnovlenie->Close();
  DM->qObnovlenie->SQL->Clear();
  DM->qObnovlenie->SQL->Add(Sql);
  try
    {
      DM->qObnovlenie->Open();
    }
  catch(Exception &E)
    {
      Application->MessageBox(("Возникла ошибка при получении данных из таблицы Ocenka" + E.Message).c_str(),"Ошибка",
                              MB_OK+MB_ICONERROR);
      InsertLog("Возникла ошибка при формировании списка работников за "+IntToStr(god)+" год по цехам в Excel");
      DM->qLogs->Requery();
      StatusBar1->SimpleText ="Отчетный период: "+IntToStr(god)+" год";
      Abort();
    }

  sFile = Path+"\\RTF\\ocenka.xlt";
  DeleteFile(WorkPath+"\\Формирование списка по цехам "+IntToStr(god)+" год");

  //Создание папки, если ее не существует
  ForceDirectories(WorkPath+"\\Формирование списка по цехам "+IntToStr(god)+" год");


  Cursor = crHourGlass;
  ProgressBar->Position = 0;
  ProgressBar->Visible = true;
  ProgressBar->Max=DM->qObnovlenie->RecordCount;


  // Открываем Excel
  try
    {
      AppEx=CreateOleObject("Excel.Application");
    }
  catch (...)
    {
      Application->MessageBox("Невозможно открыть Microsoft Excel!"
                              " Возможно это приложение на компьютере не установлено.","Ошибка",MB_OK+MB_ICONERROR);
      StatusBar1->SimpleText ="Отчетный период: "+IntToStr(god)+" год";
      ProgressBar->Visible = false;
      Cursor = crDefault;
    }

  while (DM->qObnovlenie->RecordCount>0 && !DM->qObnovlenie->Eof)
    {

      //Если возникает ошибка во время формирования отчета
      try
        {
          try
            {
              AppEx.OlePropertySet("AskToUpdateLinks",false);
              AppEx.OlePropertySet("DisplayAlerts",false);
              AppEx.OlePropertyGet("WorkBooks").OleProcedure("Open",sFile.c_str())    ;  //открываем книгу, указав её имя

              Sh=AppEx.OlePropertyGet("WorkSheets",1);                               //выбираем № активного листа книги
              //Sh=AppEx.OlePropertyGet("WorkSheets","Расчет");                      //выбираем лист по наименованию
            }
          catch(...)
            {
              Application->MessageBox("Ошибка открытия книги Microsoft Excel!","Ошибка",MB_OK+MB_ICONERROR);
              StatusBar1->SimpleText ="Отчетный период: "+IntToStr(god)+" год";
              ProgressBar->Visible = false;
              Cursor = crDefault;
            }

          int i=1;
          n=19;

          zex1 = DM->qObnovlenie->FieldByName("direkt")->AsString;
          zex = DM->qObnovlenie->FieldByName("direkt")->AsString;
          StatusBar1->SimpleText= "Идет формирование отчета: по цеху "+ zex;

          Variant Massiv;
          Massiv = VarArrayCreate(OPENARRAY(int,(0,30)),varVariant); //массив на 31 элементов

          while (!DM->qObnovlenie->Eof && zex==zex1)
            {

              Massiv.PutElement(i, 0);
              Massiv.PutElement(DM->qObnovlenie->FieldByName("fio")->AsString.c_str(), 1);
              Massiv.PutElement(DM->qObnovlenie->FieldByName("tn")->AsString.c_str(), 2);
              Massiv.PutElement(DM->qObnovlenie->FieldByName("dolg")->AsString.c_str(), 3);
              Massiv.PutElement(DM->qObnovlenie->FieldByName("funct")->AsString.c_str(), 4);
              Massiv.PutElement(DM->qObnovlenie->FieldByName("uu")->AsString.c_str(), 5);
              Massiv.PutElement(DM->qObnovlenie->FieldByName("funct_g")->AsString.c_str(), 6);
              Massiv.PutElement(DM->qObnovlenie->FieldByName("kat")->AsString.c_str(), 7);
              Massiv.PutElement(DM->qObnovlenie->FieldByName("naim_direkt")->AsString.c_str(), 8);
              Massiv.PutElement(DM->qObnovlenie->FieldByName("zex")->AsString.c_str(), 9);
              Massiv.PutElement(DM->qObnovlenie->FieldByName("naim_zex")->AsString.c_str(), 10);
              Massiv.PutElement(DM->qObnovlenie->FieldByName("nuch")->AsString.c_str(), 11);
              if (DM->qObnovlenie->FieldByName("rezult_ocen")->AsFloat==0) Massiv.PutElement(DM->qObnovlenie->FieldByName("rezult_ocen")->AsString.c_str(), 12);
              else Massiv.PutElement(DM->qObnovlenie->FieldByName("rezult_ocen")->AsFloat, 12);
              Massiv.PutElement(DM->qObnovlenie->FieldByName("rezult_proc")->AsFloat, 13);
              if (DM->qObnovlenie->FieldByName("kpe_ocen")->AsFloat==0) Massiv.PutElement(DM->qObnovlenie->FieldByName("kpe_ocen")->AsString.c_str(), 14);
              else Massiv.PutElement(DM->qObnovlenie->FieldByName("kpe_ocen")->AsFloat, 14);
              if (DM->qObnovlenie->FieldByName("komp_ocen")->AsFloat==0) Massiv.PutElement(DM->qObnovlenie->FieldByName("komp_ocen")->AsString.c_str(), 15);
              else Massiv.PutElement(DM->qObnovlenie->FieldByName("komp_ocen")->AsFloat, 15);
              Massiv.PutElement(DM->qObnovlenie->FieldByName("komp_proc")->AsFloat, 16);
              Massiv.PutElement(DM->qObnovlenie->FieldByName("efect")->AsFloat, 17);
              Massiv.PutElement(DM->qObnovlenie->FieldByName("avt_reit")->AsString.c_str(), 18);
              Massiv.PutElement(DM->qObnovlenie->FieldByName("skor_reit")->AsString.c_str(), 19);
              Massiv.PutElement(DM->qObnovlenie->FieldByName("kom_reit")->AsString.c_str(), 20);
              if (DM->qObnovlenie->FieldByName("rezerv")->AsString==1) Massiv.PutElement("да", 21);
              else Massiv.PutElement("нет", 21);
              Massiv.PutElement(DM->qObnovlenie->FieldByName("dolg_rezerv")->AsString.c_str(), 22);

              Massiv.PutElement(DM->qObnovlenie->FieldByName("naim_zex_rez")->AsString.c_str(), 23);
              Massiv.PutElement(DM->qObnovlenie->FieldByName("naim_uch_rez")->AsString.c_str(), 24);
              Massiv.PutElement(DM->qObnovlenie->FieldByName("zex_rez")->AsString.c_str(), 25);
              Massiv.PutElement(DM->qObnovlenie->FieldByName("shifr_rez")->AsString.c_str(), 26);

              Massiv.PutElement(DM->qObnovlenie->FieldByName("fio_ocen")->AsString.c_str(), 27);
              Massiv.PutElement(DM->qObnovlenie->FieldByName("dolg_ocen")->AsString.c_str(), 28);
              Massiv.PutElement(DM->qObnovlenie->FieldByName("data_ocen")->AsString.c_str(), 29);

              Sh.OlePropertyGet("Range", ("A" + IntToStr(n) + ":AE" + IntToStr(n)).c_str()).OlePropertySet("Value", Massiv); //строка с данными с ячейки A по ячейку АВ

              i++;
              n++;
              DM->qObnovlenie->Next();
              zex1 =DM->qObnovlenie->FieldByName("direkt")->AsString;
              ProgressBar->Position++;
            }

          //окрашивание ячеек
          Sh.OlePropertyGet("Range",("N18:N"+IntToStr(n-1)).c_str()).OlePropertyGet("Interior").OlePropertySet("Color",14277081);
          Sh.OlePropertyGet("Range",("Q18:S"+IntToStr(n-1)).c_str()).OlePropertyGet("Interior").OlePropertySet("Color",14277081);

          Sh.OlePropertyGet("Range",("B18:L"+IntToStr(n-1)).c_str()).OlePropertyGet("Interior").OlePropertySet("Color",14408946);
          Sh.OlePropertyGet("Range",("O18:O"+IntToStr(n-1)).c_str()).OlePropertyGet("Interior").OlePropertySet("Color",14408946);

          //рисуем сетку
          Sh.OlePropertyGet("Range",("A18:AE"+IntToStr(n-1)).c_str()).OlePropertyGet("Borders").OlePropertySet("LineStyle", xlContinuous);

          //Сохранить книгу в папке в файле по указанию
          if (zex==NULL) zex="Нет цеха";
          AnsiString vAsCurDir1=WorkPath+"\\Формирование списка по цехам "+IntToStr(god)+" год\\"+zex+".xlsx";

          Sh.OleProcedure("SaveAs",vAsCurDir1.c_str());
          //AppEx.OlePropertyGet("WorkBooks",1).OlePropertyGet("WorkSheets",1).OleProcedure("SaveAs",vAsCurDir1.c_str());


         //Закрыть книгу Excel с шаблоном для вывода информации
          AppEx.OlePropertyGet("WorkBooks",1).OleProcedure("Close");
          //Закрыть открытое приложение Excel
          //   AppEx.OleProcedure("Quit");
         // AppEx = Unassigned;
          //AppEx.OlePropertySet("AskToUpdateLinks",true);
          //AppEx.OlePropertySet("DisplayAlerts",true);

          StatusBar1->SimpleText= "Формирование отчета: по цеху "+ zex+" выполнено.";


        }
      catch (...)
        {
          AppEx.OleProcedure("Quit");
          AppEx = Unassigned;
          Cursor = crDefault;
          ProgressBar->Position=0;
          StatusBar1->SimpleText ="Отчетный период: "+IntToStr(god)+" год";
          ProgressBar->Visible=false;
          InsertLog("Возникла ошибка при формировании списка работников за "+IntToStr(god)+" год по цехам в Excel");
          Abort();
        }
    }

  AppEx.OleProcedure("Quit");
  AppEx = Unassigned;
  Cursor = crDefault;
  ProgressBar->Position=0;
  StatusBar1->SimpleText ="Отчетный период: "+IntToStr(god)+" год";
  ProgressBar->Visible=false;
  InsertLog("Формирование списка работников за "+IntToStr(god)+" год по цехам в Excel успешно завершено");
  DM->qLogs->Requery();
  Application->MessageBox("Формирование файлов успешно завершено","Формирование файлов",
                           MB_OK+MB_ICONINFORMATION);
}
//---------------------------------------------------------------------------

//Формирование списка по цехам в Excel с 2017 года
void __fastcall TMain::Spisok_po_zex2017()
{
  AnsiString sFile, Sql, zex, zex1;
  int n=18;
  double  rezult, komp, effekt;
  Variant AppEx, Sh;


  StatusBar1->SimpleText=" Идет формирование списков работников по цехам в Excel...";

  Sql="select initcap(fio) as fio, initcap(fio_ocen) as fio_ocen,                                                                                   \
              (select naim_zex from sp_pdirekt pdir where o.direkt=pdir.zex and pdir.god="+IntToStr(god)+") as naim_zex,                                      \
              (select naim from sp_direkt where god="+IntToStr(god)+" and kod = (select kod_d from sp_pdirekt pdir1 where pdir1.zex=o.direkt and pdir1.god="+IntToStr(god)+")) naim_direkt,\
              (select name_ur1 from sap_osn_sved where tn_sap=o.tn                                                   \
               union all                                                                                             \
               select name_ur1 from sap_sved_uvol where tn_sap=o.tn) as uch,                                        \
               (select distinct nazv_cex from ssap_cex where id_cex=substr(orez.zex_rez,1,2) and nazv_cex not like '%(устар.)')  naim_zex_rez,        \
               (select stext from p1000@sapmig_buffdb where otype='O' and langu='R' and short=orez.zex_rez and stext not like '%(уст%') as naim_uch_rez, \
               orez.zex_rez as zex_rez,    \
               orez.shifr_rez as shifr_rez, \
               orez.dolg_rez, \
              o.*                                                                                                    \
       from (                                                        \
             (select * from ocenka where god="+IntToStr(god)+" ) o   \
             left join                                               \
             (select * from ocenka_rez where god="+IntToStr(god)+" and (tn,tn_sap_rez) in (select tn, min(tn_sap_rez) from ocenka_rez where god="+IntToStr(god)+" group by tn)) orez\
             on o.tn=orez.tn                                                                                                                                                        \
            )                                                                                                                                                                       \
        order by direkt";
       //from ocenka o left join ocenka_rez orez where o.god="+IntToStr(god)+" and orez.god="+IntToStr(god)+" and o.tn=orez.tn \                                                    \                                          \
                                                                                                                                                       \
                                                                                                                                                                        \
  DM->qObnovlenie->Close();
  DM->qObnovlenie->SQL->Clear();
  DM->qObnovlenie->SQL->Add(Sql);
  try
    {
      DM->qObnovlenie->Open();
    }
  catch(Exception &E)
    {
      Application->MessageBox(("Возникла ошибка при получении данных из таблицы Ocenka" + E.Message).c_str(),"Ошибка",
                              MB_OK+MB_ICONERROR);
      InsertLog("Возникла ошибка при формировании списка работников за "+IntToStr(god)+" год по цехам в Excel");
      DM->qLogs->Requery();
      StatusBar1->SimpleText ="Отчетный период: "+IntToStr(god)+" год";
      Abort();
    }

  sFile = Path+"\\RTF\\ocenka2017.xlsx";
  DeleteFile(WorkPath+"\\Формирование списка по цехам "+IntToStr(god)+" год");

  //Создание папки, если ее не существует
  ForceDirectories(WorkPath+"\\Формирование списка по цехам "+IntToStr(god)+" год");


  Cursor = crHourGlass;
  ProgressBar->Position = 0;
  ProgressBar->Visible = true;
  ProgressBar->Max=DM->qObnovlenie->RecordCount;


  // Открываем Excel
  try
    {
      AppEx=CreateOleObject("Excel.Application");
    }
  catch (...)
    {
      Application->MessageBox("Невозможно открыть Microsoft Excel!"
                              " Возможно это приложение на компьютере не установлено.","Ошибка",MB_OK+MB_ICONERROR);
      StatusBar1->SimpleText ="Отчетный период: "+IntToStr(god)+" год";
      ProgressBar->Visible = false;
      Cursor = crDefault;
    }

  while (DM->qObnovlenie->RecordCount>0 && !DM->qObnovlenie->Eof)
    {

      //Если возникает ошибка во время формирования отчета
      try
        {
          try
            {
              AppEx.OlePropertySet("AskToUpdateLinks",false);
              AppEx.OlePropertySet("DisplayAlerts",false);
              AppEx.OlePropertyGet("WorkBooks").OleProcedure("Open",sFile.c_str())    ;  //открываем книгу, указав её имя

              Sh=AppEx.OlePropertyGet("WorkSheets",1);                               //выбираем № активного листа книги
              //Sh=AppEx.OlePropertyGet("WorkSheets","Расчет");                      //выбираем лист по наименованию
            }
          catch(...)
            {
              Application->MessageBox("Ошибка открытия книги Microsoft Excel!","Ошибка",MB_OK+MB_ICONERROR);
              StatusBar1->SimpleText ="Отчетный период: "+IntToStr(god)+" год";
              ProgressBar->Visible = false;
              Cursor = crDefault;
            }

          int i=1;
          n=19;

          zex1 = DM->qObnovlenie->FieldByName("direkt")->AsString;
          zex = DM->qObnovlenie->FieldByName("direkt")->AsString;
          StatusBar1->SimpleText= "Идет формирование отчета: по цеху "+ zex;

          Variant Massiv;
          Massiv = VarArrayCreate(OPENARRAY(int,(0,40)),varVariant); //массив на 41 элементов

          while (!DM->qObnovlenie->Eof && zex==zex1)
            {
              rezult = (DM->qObnovlenie->FieldByName("realizac")->AsFloat+
                        DM->qObnovlenie->FieldByName("kachestvo")->AsFloat+
                        DM->qObnovlenie->FieldByName("resurs")->AsFloat)/12*100;
              komp = (DM->qObnovlenie->FieldByName("potreb")->AsFloat+
                      DM->qObnovlenie->FieldByName("stand")->AsFloat+
                      DM->qObnovlenie->FieldByName("kach")->AsFloat+
                      DM->qObnovlenie->FieldByName("eff")->AsFloat+
                      DM->qObnovlenie->FieldByName("prof_zn")->AsFloat+
                      DM->qObnovlenie->FieldByName("lider")->AsFloat+
                      DM->qObnovlenie->FieldByName("otvetstv")->AsFloat+
                      DM->qObnovlenie->FieldByName("kom_rez")->AsFloat)/32*100;

              if (rezult>0) effekt = ((rezult*0.6)+(komp*0.4));
              else if (!DM->qObnovlenie->FieldByName("kpe_ocen")->AsString.IsEmpty()) effekt = ((DM->qObnovlenie->FieldByName("kpe_ocen")->AsString*0.6)+(komp*0.4));
              else effekt=0;

              Massiv.PutElement(i, 0);
              Massiv.PutElement(DM->qObnovlenie->FieldByName("fio")->AsString.c_str(), 1);
              Massiv.PutElement(DM->qObnovlenie->FieldByName("tn")->AsString.c_str(), 2);
              Massiv.PutElement(DM->qObnovlenie->FieldByName("dolg")->AsString.c_str(), 3);
              Massiv.PutElement(DM->qObnovlenie->FieldByName("funct")->AsString.c_str(), 4);
              Massiv.PutElement(DM->qObnovlenie->FieldByName("uu")->AsString.c_str(), 5);
              Massiv.PutElement(DM->qObnovlenie->FieldByName("funct_g")->AsString.c_str(), 6);

              Massiv.PutElement(DM->qObnovlenie->FieldByName("kat")->AsString.c_str(), 7);
              Massiv.PutElement(DM->qObnovlenie->FieldByName("naim_direkt")->AsString.c_str(), 8);
              Massiv.PutElement(DM->qObnovlenie->FieldByName("zex")->AsString.c_str(), 9);
              Massiv.PutElement(DM->qObnovlenie->FieldByName("naim_zex")->AsString.c_str(), 10);
              Massiv.PutElement(DM->qObnovlenie->FieldByName("nuch")->AsString.c_str(), 11);

              //Критерии результатов работы
              if (DM->qObnovlenie->FieldByName("realizac")->AsFloat==0) Massiv.PutElement(DM->qObnovlenie->FieldByName("realizac")->AsString.c_str(), 12);
              else Massiv.PutElement(DM->qObnovlenie->FieldByName("realizac")->AsFloat, 12);
              if (DM->qObnovlenie->FieldByName("kachestvo")->AsFloat==0) Massiv.PutElement(DM->qObnovlenie->FieldByName("kachestvo")->AsString.c_str(), 13);
              else Massiv.PutElement(DM->qObnovlenie->FieldByName("kachestvo")->AsFloat, 13);
              if (DM->qObnovlenie->FieldByName("resurs")->AsFloat==0) Massiv.PutElement(DM->qObnovlenie->FieldByName("resurs")->AsString.c_str(), 14);
              else Massiv.PutElement(DM->qObnovlenie->FieldByName("resurs")->AsFloat, 14);

              //Оценка результатов работы
              if ((DM->qObnovlenie->FieldByName("realizac")->AsFloat+DM->qObnovlenie->FieldByName("kachestvo")->AsFloat+DM->qObnovlenie->FieldByName("resurs")->AsFloat)/3==0) Massiv.PutElement("", 15);
              else Massiv.PutElement((DM->qObnovlenie->FieldByName("realizac")->AsFloat+DM->qObnovlenie->FieldByName("kachestvo")->AsFloat+DM->qObnovlenie->FieldByName("resurs")->AsFloat)/3, 15);
              Massiv.PutElement(rezult, 16);
              if (DM->qObnovlenie->FieldByName("kpe_ocen")->AsFloat==0) Massiv.PutElement(DM->qObnovlenie->FieldByName("kpe_ocen")->AsString.c_str(), 17);
              else Massiv.PutElement(DM->qObnovlenie->FieldByName("kpe_ocen")->AsFloat, 17);

              //Критерии компетенций
              if (DM->qObnovlenie->FieldByName("eff")->AsFloat==0) Massiv.PutElement(DM->qObnovlenie->FieldByName("eff")->AsString.c_str(), 18);
              else Massiv.PutElement(DM->qObnovlenie->FieldByName("eff")->AsFloat, 18);
              if (DM->qObnovlenie->FieldByName("prof_zn")->AsFloat==0) Massiv.PutElement(DM->qObnovlenie->FieldByName("prof_zn")->AsString.c_str(), 19);
              else Massiv.PutElement(DM->qObnovlenie->FieldByName("prof_zn")->AsFloat, 19);
              if (DM->qObnovlenie->FieldByName("lider")->AsFloat==0) Massiv.PutElement(DM->qObnovlenie->FieldByName("lider")->AsString.c_str(), 20);
              else Massiv.PutElement(DM->qObnovlenie->FieldByName("lider")->AsFloat, 20);
              if (DM->qObnovlenie->FieldByName("otvetstv")->AsFloat==0) Massiv.PutElement(DM->qObnovlenie->FieldByName("otvetstv")->AsString.c_str(), 21);
              else Massiv.PutElement(DM->qObnovlenie->FieldByName("otvetstv")->AsFloat, 21);
              if (DM->qObnovlenie->FieldByName("kom_rez")->AsFloat==0) Massiv.PutElement(DM->qObnovlenie->FieldByName("kom_rez")->AsString.c_str(), 22);
              else Massiv.PutElement(DM->qObnovlenie->FieldByName("kom_rez")->AsFloat, 22);
              if (DM->qObnovlenie->FieldByName("stand")->AsFloat==0) Massiv.PutElement(DM->qObnovlenie->FieldByName("stand")->AsString.c_str(), 23);
              else Massiv.PutElement(DM->qObnovlenie->FieldByName("stand")->AsFloat, 23);
              if (DM->qObnovlenie->FieldByName("potreb")->AsFloat==0) Massiv.PutElement(DM->qObnovlenie->FieldByName("potreb")->AsString.c_str(), 24);
              else Massiv.PutElement(DM->qObnovlenie->FieldByName("potreb")->AsFloat, 24);
              if (DM->qObnovlenie->FieldByName("kach")->AsFloat==0) Massiv.PutElement(DM->qObnovlenie->FieldByName("kach")->AsString.c_str(), 25);
              else Massiv.PutElement(DM->qObnovlenie->FieldByName("kach")->AsFloat, 25);


              //Оценка компетенций
              if (((DM->qObnovlenie->FieldByName("potreb")->AsFloat+
                      DM->qObnovlenie->FieldByName("stand")->AsFloat+
                      DM->qObnovlenie->FieldByName("kach")->AsFloat+
                      DM->qObnovlenie->FieldByName("eff")->AsFloat+
                      DM->qObnovlenie->FieldByName("prof_zn")->AsFloat+
                      DM->qObnovlenie->FieldByName("lider")->AsFloat+
                      DM->qObnovlenie->FieldByName("otvetstv")->AsFloat+
                      DM->qObnovlenie->FieldByName("kom_rez")->AsFloat))==0) Massiv.PutElement("", 26);
              else Massiv.PutElement(((DM->qObnovlenie->FieldByName("potreb")->AsFloat+
                                       DM->qObnovlenie->FieldByName("stand")->AsFloat+
                                       DM->qObnovlenie->FieldByName("kach")->AsFloat+
                                       DM->qObnovlenie->FieldByName("eff")->AsFloat+
                                       DM->qObnovlenie->FieldByName("prof_zn")->AsFloat+
                                       DM->qObnovlenie->FieldByName("lider")->AsFloat+
                                       DM->qObnovlenie->FieldByName("otvetstv")->AsFloat+
                                       DM->qObnovlenie->FieldByName("kom_rez")->AsFloat)), 26);
              Massiv.PutElement(komp, 27);
              Massiv.PutElement(effekt, 28);

              //Итоговая оценка
              Massiv.PutElement(DM->qObnovlenie->FieldByName("avt_reit")->AsString.c_str(), 29);
              Massiv.PutElement(DM->qObnovlenie->FieldByName("skor_reit")->AsString.c_str(), 30);
              Massiv.PutElement(DM->qObnovlenie->FieldByName("kom_reit")->AsString.c_str(), 31);

              if (DM->qObnovlenie->FieldByName("rezerv")->AsString==1) Massiv.PutElement("да", 32);
              else Massiv.PutElement("нет", 32);
              Massiv.PutElement(DM->qObnovlenie->FieldByName("dolg_rez")->AsString.c_str(), 33);

              Massiv.PutElement(DM->qObnovlenie->FieldByName("naim_zex_rez")->AsString.c_str(), 34);
              Massiv.PutElement(DM->qObnovlenie->FieldByName("naim_uch_rez")->AsString.c_str(), 35);
              Massiv.PutElement(DM->qObnovlenie->FieldByName("zex_rez")->AsString.c_str(), 36);
              Massiv.PutElement(DM->qObnovlenie->FieldByName("shifr_rez")->AsString.c_str(), 37);

              Massiv.PutElement(DM->qObnovlenie->FieldByName("fio_ocen")->AsString.c_str(), 38);
              Massiv.PutElement(DM->qObnovlenie->FieldByName("dolg_ocen")->AsString.c_str(), 39);
              Massiv.PutElement(DM->qObnovlenie->FieldByName("data_ocen")->AsString.c_str(), 40);

              Sh.OlePropertyGet("Range", ("A" + IntToStr(n) + ":AO" + IntToStr(n)).c_str()).OlePropertySet("Value", Massiv); //строка с данными с ячейки A по ячейку АВ

              i++;
              n++;
              DM->qObnovlenie->Next();
              zex1 =DM->qObnovlenie->FieldByName("direkt")->AsString;
              ProgressBar->Position++;
            }

          //окрашивание ячеек
          Sh.OlePropertyGet("Range",("Q18:Q"+IntToStr(n-1)).c_str()).OlePropertyGet("Interior").OlePropertySet("Color",14277081);
          Sh.OlePropertyGet("Range",("AB18:AD"+IntToStr(n-1)).c_str()).OlePropertyGet("Interior").OlePropertySet("Color",14277081);

          Sh.OlePropertyGet("Range",("B18:L"+IntToStr(n-1)).c_str()).OlePropertyGet("Interior").OlePropertySet("Color",14408946);
          Sh.OlePropertyGet("Range",("R18:R"+IntToStr(n-1)).c_str()).OlePropertyGet("Interior").OlePropertySet("Color",14408946);

          //рисуем сетку
          Sh.OlePropertyGet("Range",("A18:AP"+IntToStr(n-1)).c_str()).OlePropertyGet("Borders").OlePropertySet("LineStyle", xlContinuous);

          //Сохранить книгу в папке в файле по указанию
          if (zex=="")
            {
              zex="Не указан цех";
            }
          AnsiString vAsCurDir1=WorkPath+"\\Формирование списка по цехам "+IntToStr(god)+" год\\"+zex+".xlsx";

          Sh.OleProcedure("SaveAs",vAsCurDir1.c_str());
          //AppEx.OlePropertyGet("WorkBooks",1).OlePropertyGet("WorkSheets",1).OleProcedure("SaveAs",vAsCurDir1.c_str());


         //Закрыть книгу Excel с шаблоном для вывода информации
          AppEx.OlePropertyGet("WorkBooks",1).OleProcedure("Close");
          //Закрыть открытое приложение Excel
          //   AppEx.OleProcedure("Quit");
         // AppEx = Unassigned;
          //AppEx.OlePropertySet("AskToUpdateLinks",true);
          //AppEx.OlePropertySet("DisplayAlerts",true);

          StatusBar1->SimpleText= "Формирование отчета: по цеху "+ zex+" выполнено.";

          
        }
      catch (...)
        {
          AppEx.OleProcedure("Quit");
          AppEx = Unassigned;
          Cursor = crDefault;
          ProgressBar->Position=0;
          StatusBar1->SimpleText ="Отчетный период: "+IntToStr(god)+" год";
          ProgressBar->Visible=false;
          InsertLog("Возникла ошибка при формировании списка работников за "+IntToStr(god)+" год по цехам в Excel");
          Abort();
        }
    }

  AppEx.OleProcedure("Quit");
  AppEx = Unassigned;
  Cursor = crDefault;
  ProgressBar->Position=0;
  StatusBar1->SimpleText ="Отчетный период: "+IntToStr(god)+" год";
  ProgressBar->Visible=false;
  InsertLog("Формирование списка работников за "+IntToStr(god)+" год по цехам в Excel успешно завершено");
  DM->qLogs->Requery();
  Application->MessageBox("Формирование файлов успешно завершено","Формирование файлов",
                           MB_OK+MB_ICONINFORMATION);
}
//---------------------------------------------------------------------------

// Возвращает путь на папку "Мои документы"
bool __fastcall TMain::GetMyDocumentsDir(AnsiString &FolderPath)
{
  char f[MAX_PATH];

  if (SUCCEEDED(SHGetFolderPath(NULL, CSIDL_PERSONAL|CSIDL_FLAG_CREATE, NULL, SHGFP_TYPE_CURRENT, f)))
    {
      FolderPath = AnsiString(f);
      return(true);
    }

  return(false);
}
//---------------------------------------------------------------------------

// Возвращает путь на папку Temp
bool __fastcall TMain::GetTempDir(AnsiString &FolderPath)
{
  char f[MAX_PATH];

  if (GetTempPath(MAX_PATH, f))
    {
      FolderPath = AnsiString(f);
      FolderPath = FolderPath.SubString(1, FolderPath.Length()-1);
      return(true);
    }

  return(false);
}
//---------------------------------------------------------------------------

 // Определяет версию установленного Word
  AnsiString __fastcall TMain::FindWordPath()
{
  TRegistry *Reg = new TRegistry;
    try {
    Reg->RootKey = HKEY_LOCAL_MACHINE;

    for (int v=20; v>5; v--) {
      if (Reg->OpenKeyReadOnly("Software\\Microsoft\\Office\\"+IntToStr(v)+".0\\Word\\InstallRoot")) {
        if (Reg->ValueExists("Path")) {
          WordPath = Reg->ReadString("Path") + "winword.exe";
          Reg->CloseKey();
          break;
        }
        Reg->CloseKey();
      }
    }
  }
  __finally {
    delete Reg;
  }
  return(WordPath);
}
//---------------------------------------------------------------------------
void __fastcall TMain::SpeedButton1Click(TObject *Sender)
{
  N3Click(Sender);        
}
//---------------------------------------------------------------------------

void __fastcall TMain::SpeedButtonRedaktirovanieClick(TObject *Sender)
{
  Vvod->ShowModal();
}
//---------------------------------------------------------------------------

void __fastcall TMain::SpeedButton2Click(TObject *Sender)
{
  Reiting->ShowModal();        
}
//---------------------------------------------------------------------------

//Загрузка ФИО оценщика и дата оценки
void __fastcall TMain::N7Click(TObject *Sender)
{
  Zagruzka->RadioButtonDATAO->Checked = true;
  Zagruzka->SpeedButton1Click(Sender);
}
//---------------------------------------------------------------------------


void __fastcall TMain::SpeedButton4Click(TObject *Sender)
{
  Zagruzka->ShowModal();         
}
//---------------------------------------------------------------------------

//Отчет по процессу проведения ежегодной оценки персонала
void __fastcall TMain::N8Click(TObject *Sender)
{
  AnsiString Sql, row, direkt, direkt1, dir, plan_dir;
  int plan_d=0, fakt_d=0, oplan_d=0, ofakt_d=0, oplan_dir=0;
  Variant AppEx, Sh;

  StatusBar1->SimpleText=" Идет формирование отчета...";

  Sql="select distinct o.direkt as zex, kod_d, dir, naim_zex,                                  \
              nvl(count(*) over (partition by kod_d),0) plan_dir,                                              \
              nvl(count(*) over (partition by o.direkt),0) plan_zex,                                           \
              nvl(plan_d,0) as plan_d, nvl(fakt_d,0) as fakt_d                                                                              \
       from (                                                                                \
             (select * from ocenka where data_ocen is not null and god="+IntToStr(god)+") o                           \
               left join                                                                     \
             (select zex, naim_zex, kod_d, s.naim as dir, otchet from sp_direkt s left join sp_pdirekt p on kod=kod_d and otchet is not null and p.god="+IntToStr(god)+" and s.god="+IntToStr(god)+") d \
               on o.direkt=d.zex                                                                                                                 \
               left join                                                                                                                         \
             (select distinct f1.direkt, plan_d, fakt_d from (                                                                                               \
               (select direkt, count(*) over (partition by direkt) plan_d from ocenka                                                             \
                where data_ocen<=to_date(sysdate)-3  and god="+IntToStr(god)+" and direkt in (select zex from sp_pdirekt pdir where kod_d in (select kod from sp_direkt where otchet is not null and god="+IntToStr(god)+") and pdir.god="+IntToStr(god)+")) f1 \
                  left join                                                                                                        \
                (select direkt, count(*) over (partition by direkt) fakt_d from ocenka                                             \
                 where data_ocen<=to_date(sysdate)-3   and god="+IntToStr(god)+"                                                                           \
                 and direkt in (select zex from sp_pdirekt pdir1 where kod_d in (select kod from sp_direkt where otchet is not null and god="+IntToStr(god)+") and pdir1.god="+IntToStr(god)+")  \
                 and nvl(komp_ocen,0)>0 and (nvl(rezult_ocen,0)>0 or nvl(kpe_ocen,0)>0)) f2                                      \
                   on f1.direkt=f2.direkt)                                                                                          \
              )f3                                                                                                               \
               on o.direkt=f3.direkt                                                                                              \
            )                                                                                                                                  \
       where otchet is not null                                                                                                           \
       order by kod_d";

  DM->qObnovlenie->Close();
  DM->qObnovlenie->SQL->Clear();
  DM->qObnovlenie->SQL->Add(Sql);
  try
    {
      DM->qObnovlenie->Open();
    }
  catch(Exception &E)
    {
      Application->MessageBox(("Возникла ошибка при получении данных из таблицы Ocenka" + E.Message).c_str(),"Ошибка",
                              MB_OK+MB_ICONERROR);
      StatusBar1->SimpleText ="Отчетный период: "+IntToStr(god)+" год";
      Abort();
    }

  if (DM->qObnovlenie->RecordCount)
    {
      Cursor = crHourGlass;
      ProgressBar->Position = 0;
      ProgressBar->Visible = true;
      ProgressBar->Max=DM->qObnovlenie->RecordCount;

     //Открытие документа Excel
     // инициализируем Excel, открываем этот шаблон
     try
       {
         AppEx = CreateOleObject("Excel.Application");
       }
     catch (...)
       {
         Application->MessageBox("Невозможно открыть Microsoft Excel!"
                                 " Возможно это приложение на компьютере не установлено.","Ошибка",MB_OK+MB_ICONERROR);
         Abort();
       }

     //Если возникает ошибка во время формирования отчета
     try
       {
         try
           {
             AppEx.OlePropertySet("AskToUpdateLinks",false);
             AppEx.OlePropertySet("DisplayAlerts",false);
             AppEx.OlePropertyGet("WorkBooks").OleProcedure("Open",(Path +"\\RTF\\3dnya.xlt").c_str())    ;  //открываем книгу, указав её имя

             Sh=AppEx.OlePropertyGet("WorkSheets",1);                               //выбираем № активного листа книги
             //Sh=AppEx.OlePropertyGet("WorkSheets","Расчет");                      //выбираем лист по наименованию

           }
         catch(...)
           {
             Application->MessageBox("Ошибка открытия книги Microsoft Excel!","Ошибка",MB_OK + MB_ICONERROR);
           }

         row = DM->qObnovlenie->RecordCount+1;

         // выводим в шаблон данные
         // сначала заголовок
         // toExcel(AppEx,"data",String(Date()));

         // вставляем в шаблон нужное количество строк
         Variant C;
         Sh.OleProcedure("Select");
         C=Sh.OlePropertyGet("Range","zex");
         C=Sh.OlePropertyGet("Rows",(int) C.OlePropertyGet("Row")+1);
         for(int i=1;i<row;i++) C.OleProcedure("Insert");
         int i=1;
         int n=6;

         //вывод даты
         toExcel(Sh,"data",String(Date()));
         toExcel(Sh,"data2",String(Date()));

         while(!DM->qObnovlenie->Eof)
           {
             direkt = DM->qObnovlenie->FieldByName("kod_d")->AsString;
             direkt1 = DM->qObnovlenie->FieldByName("kod_d")->AsString;

             while(!DM->qObnovlenie->Eof && direkt==direkt1)
               {
                 //вывод данных
                 toExcel(Sh,"zex",i,i+1);
                 toExcel(Sh,"zex",i, DM->qObnovlenie->FieldByName("zex")->AsString.c_str());
                 toExcel(Sh,"naim",i, DM->qObnovlenie->FieldByName("naim_zex")->AsString.c_str());
                 toExcel(Sh,"plan",i,DM->qObnovlenie->FieldByName("plan_zex")->AsString.c_str());
                 toExcel(Sh,"plan_d",i, DM->qObnovlenie->FieldByName("plan_d")->AsString.c_str());
                 toExcel(Sh,"fakt_d",i, DM->qObnovlenie->FieldByName("fakt_d")->AsString.c_str());

                 //Вычисляемые поля
                 Sh.OlePropertyGet("Range", "proc").OlePropertyGet("Offset", i).OlePropertySet("Formula", ("=F"+IntToStr(n)+"/C"+IntToStr(n)+"*100").c_str());
                 Sh.OlePropertyGet("Range", "proc_d").OlePropertyGet("Offset", i).OlePropertySet("Formula", ("=ЕСЛИ(E"+IntToStr(n)+"=0;0;(F"+IntToStr(n)+"/E"+IntToStr(n)+")*100)").c_str());
                // Sh.OlePropertyGet("Range", "proc_d").OlePropertyGet("Offset", i).OlePropertySet("Formula", ("=F"+IntToStr(n)+"/E"+IntToStr(n)+"*100").c_str());
                 /* toExcel(AppEx,"dtuvol",i, DM->qObnovlenie->FieldByName("dtuvol")->AsString.c_str());
                 toExcel(AppEx,"sum",i,DM->qObnovlenie->FieldByName("spre")->AsFloat); */
                 i++;
                 n++;

                 dir = DM->qObnovlenie->FieldByName("dir")->AsString;
                 plan_dir = DM->qObnovlenie->FieldByName("plan_dir")->AsString;
                 plan_d+= DM->qObnovlenie->FieldByName("plan_d")->AsInteger;
                 fakt_d+= DM->qObnovlenie->FieldByName("fakt_d")->AsInteger;

                 DM->qObnovlenie->Next();
                 ProgressBar->Position++;
                 direkt = DM->qObnovlenie->FieldByName("kod_d")->AsString;
               }

             //вывод наименования дирекции
             toExcel(Sh,"naim",i, dir.c_str());
             toExcel(Sh,"plan",i, plan_dir.c_str());
             toExcel(Sh,"plan_d",i, IntToStr(plan_d).c_str());
             toExcel(Sh,"fakt_d",i, IntToStr(fakt_d).c_str());

             //вычисляемые поля
             Sh.OlePropertyGet("Range", "proc").OlePropertyGet("Offset", i).OlePropertySet("Formula", ("=F"+IntToStr(n)+"/C"+IntToStr(n)+"*100").c_str());
             Sh.OlePropertyGet("Range", "proc_d").OlePropertyGet("Offset", i).OlePropertySet("Formula", ("=ЕСЛИ(E"+IntToStr(n)+"=0;0;(F"+IntToStr(n)+"/E"+IntToStr(n)+")*100)").c_str());
             //Sh.OlePropertyGet("Range", "proc_d").OlePropertyGet("Offset", i).OlePropertySet("Formula", ("=F"+IntToStr(n)+"/E"+IntToStr(n)+"*100").c_str());
             //объединение ячеек
             Sh.OlePropertyGet("Range",("A"+IntToStr(n)+":B"+IntToStr(n)).c_str()).OleProcedure("Merge");
             //окрашивание ячеек
             Sh.OlePropertyGet("Range", ("A"+IntToStr(n)+":G"+IntToStr(n)).c_str()).OlePropertyGet("Interior").OlePropertySet("ColorIndex",40);
             //жирный шрифт
             Sh.OlePropertyGet("Range",("A"+IntToStr(n)+":G"+IntToStr(n)).c_str()).OlePropertyGet("Font").OlePropertySet("Bold",true);

             oplan_dir+=StrToInt(plan_dir);
             oplan_d+=plan_d;
             ofakt_d+=fakt_d;

             plan_d=0;
             fakt_d=0;
             i++;
             n++;
           }

         //вывод общих итогов
         toExcel(Sh,"naim",i, "Общий итог");
         toExcel(Sh,"plan",i, IntToStr(oplan_dir).c_str());
         toExcel(Sh,"plan_d",i, IntToStr(oplan_d).c_str());
         toExcel(Sh,"fakt_d",i, IntToStr(ofakt_d).c_str());

         //вычисляемые поля
         Sh.OlePropertyGet("Range", "proc").OlePropertyGet("Offset", i).OlePropertySet("Formula", ("=F"+IntToStr(n)+"/C"+IntToStr(n)+"*100").c_str());
                                                                                                       //=ЕСЛИ(E13=0;0;(F13/E13)*100)
         Sh.OlePropertyGet("Range", "proc_d").OlePropertyGet("Offset", i).OlePropertySet("Formula", ("=ЕСЛИ(E"+IntToStr(n)+"=0;0;(F"+IntToStr(n)+"/E"+IntToStr(n)+")*100)").c_str());
         //Sh.OlePropertyGet("Range", "proc_d").OlePropertyGet("Offset", i).OlePropertySet("Formula", ("=F"+IntToStr(n)+"/E"+IntToStr(n)+"*100").c_str());
         
         //объединение ячеек
         Sh.OlePropertyGet("Range",("A"+IntToStr(n)+":B"+IntToStr(n)).c_str()).OleProcedure("Merge");
         //окрашивание ячеек
         //Sh.OlePropertyGet("Range", ("A"+IntToStr(n)+":G"+IntToStr(n)).c_str()).OlePropertyGet("Interior").OlePropertySet("ColorIndex",40);
         //жирный шрифт
         Sh.OlePropertyGet("Range",("A"+IntToStr(n)+":G"+IntToStr(n)).c_str()).OlePropertyGet("Font").OlePropertySet("Bold",true);
         //увеличить шрифт
         Sh.OlePropertyGet("Range", ("A"+IntToStr(n)+":G"+IntToStr(n)).c_str()).OlePropertyGet("Font").OlePropertySet("Size",13);

         //рисуем сетку
         Sh.OlePropertyGet("Range",("A5:G"+IntToStr(n)).c_str()).OlePropertyGet("Borders").OlePropertySet("LineStyle", xlContinuous);

         //Отключить вывод сообщений с вопросами типа "Заменить файл..."
         AppEx.OlePropertySet("DisplayAlerts",false);

         //Создание папки, если ее не существует
         ForceDirectories(Main->WorkPath);

         //Сохранить книгу в папке в файле по указанию
         AnsiString vAsCurDir1=WorkPath+"\\Очет по ежегодной оценке персонала.xlsx";
         Sh.OleProcedure("SaveAs",vAsCurDir1.c_str());

         //Закрыть открытое приложение Excel
         // AppEx.OleProcedure("Quit");
         AppEx.OlePropertyGet("WorkBooks",1).OleProcedure("Close");
         Application->MessageBox("Отчет успешно сформирован!", "Формирование отчета",
                                  MB_OK+MB_ICONINFORMATION);
         AppEx.OlePropertyGet("WorkBooks").OleProcedure("Open",vAsCurDir1.c_str());
         AppEx.OlePropertySet("Visible",true);
         AppEx.OlePropertySet("AskToUpdateLinks",true);
         AppEx.OlePropertySet("DisplayAlerts",true);

         Cursor = crDefault;
         ProgressBar->Position = 0;
         ProgressBar->Visible = false;
         StatusBar1->SimpleText ="Отчетный период: "+IntToStr(god)+" год";
       }
     catch(...)
       {
         AppEx.OleProcedure("Quit");
         AppEx = Unassigned;
         Cursor = crDefault;
         ProgressBar->Position = 0;
         ProgressBar->Visible = false;
         StatusBar1->SimpleText ="Отчетный период: "+IntToStr(god)+" год";
       }
    }
  else
    {
      Application->MessageBox("Нет данных для формирования отчета!", "Формирование отчета",
                               MB_OK+MB_ICONINFORMATION);
      StatusBar1->SimpleText ="Отчетный период: "+IntToStr(god)+" год";
    }
}
//---------------------------------------------------------------------------
//---------------------------------------------------------------------------

AnsiString  __fastcall TMain::SetNull (AnsiString str, AnsiString r)
{
  if (str.Length()) return str;
  else return r;
}
//---------------------------------------------------------------------------
float __fastcall TMain::SetNullF (AnsiString str)
{
  if (str.Length()) return StrToFloat(str);
  else return 0;
}
//---------------------------------------------------------------------------
void __fastcall TMain::N9Click(TObject *Sender)
{
  AnsiString Answer, Answer2, Sql;
  TLocateOptions SearchOptions;

  if( InputQuery("Поиск по цеху","Введите цех", Answer) == True && InputQuery("Поиск по табельному номеру","Введите табельный номер", Answer2) == True)
    {
      if(Answer=="" && (StrToInt(Answer)))
        {
          Application->MessageBox("Не введен цех","Предупреждение",MB_OK +MB_ICONWARNING);
          Abort();
        }
      if(Answer2=="")
        {
          Application->MessageBox("Не введен табельный номер","Предупреждение",MB_OK +MB_ICONWARNING);
        }
      else
        {
           Variant locvalues[] = {Answer, Answer2};
           if (!DM->qOcenka->Locate("zex;tn", VarArrayOf(locvalues, 1),
                                               SearchOptions << loCaseInsensitive) )
            {
              Application->MessageBox("Введенный табельный номер в цехе не найден",
                                      "Резултаты поиска",
                                      MB_OK + MB_ICONINFORMATION);
            }
       }
   }
}
//---------------------------------------------------------------------------

void __fastcall TMain::DBGridEh1DrawColumnCell(TObject *Sender,
      const TRect &Rect, int DataCol, TColumnEh *Column,
      TGridDrawState State)
{
  // выделение цветом активной записи
 if (State.Contains(gdSelected))
    {
      ((TDBGridEh *) Sender)->Canvas->Brush->Color = TColor(0x00C8F7E3);//0x00A3F1D1);//clInfoBk;
      ((TDBGridEh *) Sender)->Canvas->Font->Color= clBlack;
      ((TDBGridEh *) Sender)->DefaultDrawColumnCell(Rect, DataCol, Column, State);
    }
}
//---------------------------------------------------------------------------

void __fastcall TMain::N6Click(TObject *Sender)
{
  SpeedButtonRedaktirovanieClick(Sender);        
}
//---------------------------------------------------------------------------

void __fastcall TMain::DBGridEh1KeyDown(TObject *Sender, WORD &Key,
      TShiftState Shift)
{
/* if (Key == VK_INSERT)
    {
      if (DM->qZakritieMes->FieldByName("bol")->AsString==2)
        {
          Abort();
        }
      else
        {
          //добавление
          NDobavClick(Sender);
        }
    }

  if (Key == VK_DELETE && DM->qFormat72->RecordCount!=0 && N2Delete->Enabled==true)
    {
      // удаление записи
      if (Prava!=1 || (Prava==1 && DM->qFormat72->FieldByName("status")->AsInteger==0))
        {
          N2DeleteClick(Sender);
        }
    }

  if ((Shift.Contains(ssCtrl)) && (Key == 80))
    {
      //поиск
      N5PoiskClick(Sender);
    } */       
}
//---------------------------------------------------------------------------

void __fastcall TMain::DBGridEh1KeyPress(TObject *Sender, char &Key)
{
  if (Key == VK_RETURN)
    {
      // редактирование
      SpeedButtonRedaktirovanieClick(Sender);
    }
}
//---------------------------------------------------------------------------

void __fastcall TMain::FormKeyDown(TObject *Sender, WORD &Key,
      TShiftState Shift)
{
  if (Key==VK_RETURN)
  FindNextControl((TWinControl *)Sender, true, true, false)->SetFocus();        
}
//---------------------------------------------------------------------------

void __fastcall TMain::DBGridEh1DblClick(TObject *Sender)
{
  SpeedButtonRedaktirovanieClick(Sender);         
}
//---------------------------------------------------------------------------

//Расчет рейтинга
void __fastcall TMain::N11Click(TObject *Sender)
{
  AnsiString Sql;

  StatusBar1->SimpleText=" Идет расчет рейтинга...";

  //update всех значений (обнуление автоматического рейтинга)
  Sql = "update ocenka set avt_reit=NULL                                                        \
         where (zex,tn) in (select zex,tn                                                           \
                            from ( select zex,                                                           \
                                          tn,                                                            \
                                          min(efect) over (partition by kat,funct_g,fio_ocen,kpe) a_min, \
                                          count(*) over (partition by kat, funct_g, fio_ocen,kpe) kol    \
                                   from ( select decode(nvl(kpe_ocen,0),0,0,1) as kpe,                   \
                                          o.*                                                            \
                                          from ocenka o                                            \
                                          where nvl(efect,0)!=0  and god="+IntToStr(god)+"                                         \
                                        )                                                                \
                                  ) d \
                            where kol>=5)";

  DM->qObnovlenie->Close();
  DM->qObnovlenie->SQL->Clear();
  DM->qObnovlenie->SQL->Add(Sql);
  try
    {
      DM->qObnovlenie->ExecSQL();
    }
  catch(Exception &E)
    {
      Application->MessageBox(("Возникла ошибка при обновлении автоматического рейтинга в таблице Ocenka" + E.Message).c_str(),"Ошибка",
                              MB_OK+MB_ICONERROR);
      StatusBar1->SimpleText ="Отчетный период: "+IntToStr(god)+" год";
      Abort();
    }

  //Запрос                                                                               
  Sql = "select distinct upper(fio_ocen) as fio_ocen,                                                       \
                upper(kat) as kat,                                                                     \
                upper(decode(funct_g,NULL,'1',funct_g)) as funct_g,                                                                 \
                a_min,                                                                   \
                a_max,                                                                   \
                kol*0.05 as A5,                                                          \
                kol*0.2 as A20,                                                          \
                kol*0.6 as B60,                                                          \
                kol*0.2 as C20,                                                          \
                kol*0.05 as C5,                                                          \
                kpe,                                                                     \
                god                                                                      \
          from (                                                                         \
                 select  min(efect) over (partition by kat,funct_g,fio_ocen, kpe) a_min, \
                         max(efect) over (partition by kat,funct_g,fio_ocen, kpe) a_max, \
                         count(*) over (partition by kat, funct_g, fio_ocen, kpe) kol,   \
                         d.*                                                             \
                 from (                                                                  \
                        select decode (nvl(kpe_ocen,0),0,0,1) as kpe,                    \
                        o.*                                                              \
                        from ocenka o where nvl(efect,0)!=0 and fio_ocen is not null and kat is not null  and god="+IntToStr(god)+"         \
                      ) d                                                                \
               )                                                                         \
          where kol>=5                                                                   \
          order by fio_ocen, kat, funct_g";

  DM->qObnovlenie->Close();
  DM->qObnovlenie->SQL->Clear();
  DM->qObnovlenie->SQL->Add(Sql);
  try
    {
      DM->qObnovlenie->Open();
    }
  catch(Exception &E)
    {
      Application->MessageBox(("Возникла ошибка при обновлении получении данных из таблицы Ocenka" + E.Message).c_str(),"Ошибка",
                              MB_OK+MB_ICONERROR);
      StatusBar1->SimpleText ="Отчетный период: "+IntToStr(god)+" год";
      Abort();
    }

  if (DM->qObnovlenie->RecordCount>0)
    {
      Cursor = crHourGlass;
      ProgressBar->Position = 0;
      ProgressBar->Visible = true;
      ProgressBar->Max=DM->qObnovlenie->RecordCount;

      while (!DM->qObnovlenie->Eof)
        {
          StatusBar1->SimpleText=" Идет расчет рейтинга, оценщик:  "+ DM->qObnovlenie->FieldByName("fio_ocen")->AsString;
          
          //Выполнение продцедуры по расчету рейтинга
          DM->spOcenka->Parameters->ParamByName("pFio_ocen")->Value = SetNull(DM->qObnovlenie->FieldByName("fio_ocen")->AsString);
          DM->spOcenka->Parameters->ParamByName("pKat")->Value = SetNull(DM->qObnovlenie->FieldByName("kat")->AsString);
          DM->spOcenka->Parameters->ParamByName("pFunct_g")->Value = SetNull(DM->qObnovlenie->FieldByName("funct_g")->AsString);
          DM->spOcenka->Parameters->ParamByName("pMin")->Value = DM->qObnovlenie->FieldByName("a_min")->AsFloat;
          DM->spOcenka->Parameters->ParamByName("pMax")->Value = DM->qObnovlenie->FieldByName("a_max")->AsFloat;
          DM->spOcenka->Parameters->ParamByName("PA5")->Value = DM->qObnovlenie->FieldByName("A5")->AsFloat;
          DM->spOcenka->Parameters->ParamByName("PA20")->Value = DM->qObnovlenie->FieldByName("A20")->AsFloat;
          DM->spOcenka->Parameters->ParamByName("PB60")->Value = DM->qObnovlenie->FieldByName("B60")->AsFloat;
          DM->spOcenka->Parameters->ParamByName("PC20")->Value = DM->qObnovlenie->FieldByName("C20")->AsFloat;
          DM->spOcenka->Parameters->ParamByName("PC5")->Value = DM->qObnovlenie->FieldByName("C5")->AsFloat;
          DM->spOcenka->Parameters->ParamByName("pKpe")->Value = DM->qObnovlenie->FieldByName("kpe")->AsFloat;
          DM->spOcenka->Parameters->ParamByName("pGod")->Value = DM->qObnovlenie->FieldByName("god")->AsInteger;

          try
            {
              DM->spOcenka->ExecProc();
            }
          catch(Exception &E)
            {
              Application->MessageBox(("Возникла ошибка при расчете автоматического рейтинга (продцедура CALC_OCENKA)" + E.Message).c_str(),"Ошибка",
                                       MB_OK+MB_ICONERROR);
              InsertLog("Расчет рейтинга за "+IntToStr(god)+" год: возникла ошибка при расчете рейтинга по всем работникам");
              DM->qLogs->Requery();
              StatusBar1->SimpleText ="Отчетный период: "+IntToStr(god)+" год";
              ProgressBar->Visible = false;
              Abort();
            }

          DM->qObnovlenie->Next();
          ProgressBar->Position++;
        }

      DM->qOcenka->Requery();

      Cursor = crDefault;
      ProgressBar->Position = 0;
      ProgressBar->Visible = false;
      StatusBar1->SimpleText ="Отчетный период: "+IntToStr(god)+" год";

      Application->MessageBox("Расчет рейтинга выполнен успешно =)","Расчет рейтинга",
                                       MB_OK+MB_ICONINFORMATION);
      InsertLog("Расчет рейтинга за "+IntToStr(god)+" год по ВСЕМ работникам выполнен успешно");
      DM->qLogs->Requery();

    }
  else
    {
      Application->MessageBox("Нет данных для расчета рейтинга. \nВозможно по данным с не проставленным рейтингом меннее 5 записей \nв пределах одного оценщика, категории и функциональной группы. ","Расчет рейтинга",
                                       MB_OK+MB_ICONINFORMATION);
      StatusBar1->SimpleText ="Отчетный период: "+IntToStr(god)+" год";
    }
}
//---------------------------------------------------------------------------


void __fastcall TMain::N14Click(TObject *Sender)
{
  Sprav->ShowModal();
}
//---------------------------------------------------------------------------


//Формирование списка по предприятию в Excel
void __fastcall TMain::Excel1Click(TObject *Sender)
{
  if (god<god_t) SpisokExcel(0);
  else SpisokExcel2017(0);

}
//---------------------------------------------------------------------------
void __fastcall TMain::N110Click(TObject *Sender)
{
  AnsiString Answer, Sql;

  if( InputQuery("Формирование списка по цеху","Введите номер цеха", Answer) == True)
    {
      if(Answer=="" && (StrToInt(Answer)))
        {
          Application->MessageBox("Не введен шифр цеха","Предупреждение",MB_OK +MB_ICONWARNING);
          Abort();
        }
      else
        {
          if (god<god_t) SpisokExcel(Answer);
          else SpisokExcel2017(Answer);
        }
   }
}
//---------------------------------------------------------------------------
 //Формирование списка по цеху или предприятию в Excel
void __fastcall TMain::SpisokExcel(AnsiString otchet_zex)
{
  AnsiString sFile, Sql;
  int n=18;
  double  rezult, komp, effekt;
  Variant AppEx, Sh;


  Sql="select initcap(fio) as fio, initcap(fio_ocen) as fio_ocen, direkt,                                                                                 \
              (select naim_zex from sp_pdirekt pdir where o.direkt=pdir.zex and pdir.god="+IntToStr(god)+") as naim_zex,                                      \
              (select naim from sp_direkt where god="+IntToStr(god)+" and kod = (select kod_d from sp_pdirekt pdir1 where pdir1.zex=o.direkt and pdir1.god="+IntToStr(god)+")) naim_direkt,\
              (select name_ur1 from sap_osn_sved where tn_sap=o.tn                                                   \
               union all                                                                                             \
               select name_ur1 from sap_sved_uvol where tn_sap=o.tn) as uch,                                        \
               (select distinct nazv_cex from ssap_cex where id_cex=substr(orez.zex_rez,1,2) and nazv_cex not like '%(устар.)')  naim_zex_rez,        \
               (select stext from p1000@sapmig_buffdb where otype='O' and langu='R' and short=orez.zex_rez and stext not like '%(уст%') as naim_uch_rez,  \
               orez.zex_rez as zex_rez, \
               orez.shifr_rez as shifr_rez, \
              o.*                                                                                                    \
       from ocenka  o left join ocenka_rez orez where o.god="+IntToStr(god)+" and orez.god="+IntToStr(god)+" and o.tn=orez.tn";
       // (select stext from p1000@sapmig_buffdb where otype='O' and langu='R' and objid=o.objid) as naim_uch_rez,





  if (otchet_zex==0)
    {
      Sql+=" order by zex";
      StatusBar1->SimpleText=" Идет формирование списка работников по предприятию в Excel...";
    }
  else
    {
      Sql+=" and direkt ="+otchet_zex+" order by zex, tn";
      StatusBar1->SimpleText=" Идет формирование списка работников по "+otchet_zex+" цеху в Excel...";                                                                                                                                                \
    }
                                                                                                                                                                        \
  DM->qObnovlenie->Close();
  DM->qObnovlenie->SQL->Clear();
  DM->qObnovlenie->SQL->Add(Sql);
  try
    {
      DM->qObnovlenie->Open();
    }
  catch(Exception &E)
    {
      Application->MessageBox(("Возникла ошибка при получении данных из таблицы Ocenka" + E.Message).c_str(),"Ошибка",
                              MB_OK+MB_ICONERROR);
      if (otchet_zex==0) InsertLog("Возникла ошибка при формировании списка работников за "+IntToStr(god)+" год по предприятию в Excel");
      else InsertLog("Возникла ошибка при формировании списка работников за "+IntToStr(god)+" год по "+otchet_zex+" цеху в Excel");
      DM->qLogs->Requery();
      StatusBar1->SimpleText ="Отчетный период: "+IntToStr(god)+" год";
      Abort();
    }

 // sFile = Path+"\\RTF\\ocenka.xlt";

  Cursor = crHourGlass;
  ProgressBar->Position = 0;
  ProgressBar->Visible = true;
  ProgressBar->Max=DM->qObnovlenie->RecordCount;

  // инициализируем Excel, открываем этот шаблон
  try
    {
      AppEx=CreateOleObject("Excel.Application");
    }
  catch (...)
    {
      Application->MessageBox("Невозможно открыть Microsoft Excel!"
                              " Возможно это приложение на компьютере не установлено.","Ошибка",MB_OK+MB_ICONERROR);
      StatusBar1->SimpleText ="Отчетный период: "+IntToStr(god)+" год";
      ProgressBar->Visible = false;
      Cursor = crDefault;
    }

  //Если возникает ошибка во время формирования отчета
  try
    {
      try
        {
          AppEx.OlePropertySet("AskToUpdateLinks",false);
          AppEx.OlePropertySet("DisplayAlerts",false);

          //Копируем шаблон файла в Мои документы
          if (otchet_zex==0)
            {
              //Создание папки, если ее не существует
              ForceDirectories(WorkPath);
              CopyFile((Path+"\\RTF\\ocenka.xlsx").c_str(), (WorkPath+"\\Формирование списка по предприятию.xlsx").c_str(), false);
              sFile = WorkPath+"\\Формирование списка по предприятию.xlsx";
            }
          else
            {
              DeleteFile(WorkPath+"\\Формирование списка по 1 цеху");
              //Создание папки, если ее не существует
              ForceDirectories(WorkPath+"\\Формирование списка по 1 цеху");
              CopyFile((Path+"\\RTF\\ocenka.xlsx").c_str(), (WorkPath+"\\Формирование списка по 1 цеху\\Формирование списка по "+otchet_zex+" цеху.xlsx").c_str(), false);
              sFile = WorkPath+"\\Формирование списка по 1 цеху\\Формирование списка по "+otchet_zex+" цеху.xlsx";
            }
          AppEx.OlePropertyGet("WorkBooks").OleProcedure("Open",sFile.c_str())    ;  //открываем книгу, указав её имя

          Sh=AppEx.OlePropertyGet("WorkSheets",1);                               //выбираем № активного листа книги
          //Sh=AppEx.OlePropertyGet("WorkSheets","Расчет");                      //выбираем лист по наименованию
        }
      catch(...)
        {
          Application->MessageBox("Ошибка открытия книги Microsoft Excel!","Ошибка",MB_OK+MB_ICONERROR);
          StatusBar1->SimpleText ="Отчетный период: "+IntToStr(god)+" год";
          ProgressBar->Visible = false;
          Cursor = crDefault;
        }

      int i=1;
      n=19;

      Variant Massiv;
      Massiv = VarArrayCreate(OPENARRAY(int,(0,30)),varVariant); //массив на 41 элементов

      while (!DM->qObnovlenie->Eof)
        {
          if (otchet_zex==0) StatusBar1->SimpleText=" Идет формирование списка работников по предприятию в Excel: цех "+ DM->qObnovlenie->FieldByName("zex")->AsString;
          else StatusBar1->SimpleText=" Идет формирование списка по "+DM->qObnovlenie->FieldByName("direkt")->AsString+" цеху в Excel";

          Massiv.PutElement(i, 0);
          Massiv.PutElement(DM->qObnovlenie->FieldByName("fio")->AsString.c_str(), 1);
          Massiv.PutElement(DM->qObnovlenie->FieldByName("tn")->AsString.c_str(), 2);
          Massiv.PutElement(DM->qObnovlenie->FieldByName("dolg")->AsString.c_str(), 3);
          Massiv.PutElement(DM->qObnovlenie->FieldByName("funct")->AsString.c_str(), 4);
          Massiv.PutElement(DM->qObnovlenie->FieldByName("uu")->AsString.c_str(), 5);
          Massiv.PutElement(DM->qObnovlenie->FieldByName("funct_g")->AsString.c_str(), 6);
          Massiv.PutElement(DM->qObnovlenie->FieldByName("kat")->AsString.c_str(), 7);
          Massiv.PutElement(DM->qObnovlenie->FieldByName("naim_direkt")->AsString.c_str(), 8);
          Massiv.PutElement(DM->qObnovlenie->FieldByName("zex")->AsString.c_str(), 9);
          Massiv.PutElement(DM->qObnovlenie->FieldByName("naim_zex")->AsString.c_str(), 10);
          Massiv.PutElement(DM->qObnovlenie->FieldByName("nuch")->AsString.c_str(), 11);
          if (DM->qObnovlenie->FieldByName("rezult_ocen")->AsFloat==0) Massiv.PutElement(DM->qObnovlenie->FieldByName("rezult_ocen")->AsString.c_str(), 12);
          else Massiv.PutElement(DM->qObnovlenie->FieldByName("rezult_ocen")->AsFloat, 12);
          Massiv.PutElement(DM->qObnovlenie->FieldByName("rezult_proc")->AsFloat, 13);
          if (DM->qObnovlenie->FieldByName("kpe_ocen")->AsFloat==0) Massiv.PutElement(DM->qObnovlenie->FieldByName("kpe_ocen")->AsString.c_str(), 14);
          else Massiv.PutElement(DM->qObnovlenie->FieldByName("kpe_ocen")->AsFloat, 14);
          if (DM->qObnovlenie->FieldByName("komp_ocen")->AsFloat==0) Massiv.PutElement(DM->qObnovlenie->FieldByName("komp_ocen")->AsString.c_str(), 15);
          else Massiv.PutElement(DM->qObnovlenie->FieldByName("komp_ocen")->AsFloat, 15);
          Massiv.PutElement(DM->qObnovlenie->FieldByName("komp_proc")->AsFloat, 16);
          Massiv.PutElement(DM->qObnovlenie->FieldByName("efect")->AsFloat, 17);
          Massiv.PutElement(DM->qObnovlenie->FieldByName("avt_reit")->AsString.c_str(), 18);
          Massiv.PutElement(DM->qObnovlenie->FieldByName("skor_reit")->AsString.c_str(), 19);
          Massiv.PutElement(DM->qObnovlenie->FieldByName("kom_reit")->AsString.c_str(), 20);
          if (DM->qObnovlenie->FieldByName("rezerv")->AsString==1) Massiv.PutElement("да", 21);
          else Massiv.PutElement("нет", 21);
          Massiv.PutElement(DM->qObnovlenie->FieldByName("dolg_rezerv")->AsString.c_str(), 22);

          Massiv.PutElement(DM->qObnovlenie->FieldByName("naim_zex_rez")->AsString.c_str(), 23);
          Massiv.PutElement(DM->qObnovlenie->FieldByName("naim_uch_rez")->AsString.c_str(), 24);
          Massiv.PutElement(DM->qObnovlenie->FieldByName("zex_rez")->AsString.c_str(), 25);
          Massiv.PutElement(DM->qObnovlenie->FieldByName("shifr_rez")->AsString.c_str(), 26);

          Massiv.PutElement(DM->qObnovlenie->FieldByName("fio_ocen")->AsString.c_str(), 27);
          Massiv.PutElement(DM->qObnovlenie->FieldByName("dolg_ocen")->AsString.c_str(), 28);
          Massiv.PutElement(DM->qObnovlenie->FieldByName("data_ocen")->AsString.c_str(), 29);

          Sh.OlePropertyGet("Range", ("A" + IntToStr(n) + ":AE" + IntToStr(n)).c_str()).OlePropertySet("Value", Massiv); //строка с данными с ячейки A по ячейку АВ

          i++;
          n++;
          DM->qObnovlenie->Next();
          ProgressBar->Position++;
        }

       // Sh.OlePropertyGet("Range", ("LQ18:LQ" + IntToStr(i-1)).c_str()).OlePropertySet("NumberFormat", "0.00");

      //окрашивание ячеек
      Sh.OlePropertyGet("Range",("N18:N"+IntToStr(n-1)).c_str()).OlePropertyGet("Interior").OlePropertySet("Color",14277081);
      Sh.OlePropertyGet("Range",("Q18:S"+IntToStr(n-1)).c_str()).OlePropertyGet("Interior").OlePropertySet("Color",14277081);

      Sh.OlePropertyGet("Range",("B18:L"+IntToStr(n-1)).c_str()).OlePropertyGet("Interior").OlePropertySet("Color",14408946);
      Sh.OlePropertyGet("Range",("O18:O"+IntToStr(n-1)).c_str()).OlePropertyGet("Interior").OlePropertySet("Color",14408946);

      //рисуем сетку
      Sh.OlePropertyGet("Range",("A18:AE"+IntToStr(n-1)).c_str()).OlePropertyGet("Borders").OlePropertySet("LineStyle", xlContinuous);

      //Сохранить книгу в папке в файле по указанию
     // AnsiString vAsCurDir1=WorkPath+"\\Формирование списка по предприятию";

     // Sh.OleProcedure("SaveAs",vAsCurDir1.c_str());
     AppEx.OlePropertyGet("WorkBooks",1).OleFunction("Save");

      /* //Закрыть открытое приложение Excel
      AppEx.OleProcedure("Quit");
      AppEx = Unassigned;  */

      //Закрыть книгу Excel с шаблоном для вывода информации
     // AppEx.OlePropertyGet("WorkBooks",1).OleProcedure("Close");
      Application->MessageBox("Отчет успешно сформирован!", "Формирование отчета",
                               MB_OK+MB_ICONINFORMATION);
      //AppEx.OlePropertyGet("WorkBooks").OleProcedure("Open",vAsCurDir1.c_str());
      AppEx.OlePropertySet("Visible",true);
      AppEx.OlePropertySet("AskToUpdateLinks",true);
      AppEx.OlePropertySet("DisplayAlerts",true);

      StatusBar1->SimpleText= "Формирование отчета выполнено.";

      Cursor = crDefault;
      ProgressBar->Position=0;
      ProgressBar->Visible = false;
      StatusBar1->SimpleText ="Отчетный период: "+IntToStr(god)+" год";
    }
  catch(...)
    {
      AppEx.OleProcedure("Quit");
      AppEx = Unassigned;
      Cursor = crDefault;
      ProgressBar->Position=0;
      ProgressBar->Visible = false;

      StatusBar1->SimpleText ="Отчетный период: "+IntToStr(god)+" год";
      if (otchet_zex==0) InsertLog("Возникла ошибка при формировании списка работников за "+IntToStr(god)+" год по предприятию в Excel");
      else InsertLog("Возникла ошибка при формировании списка работников за "+IntToStr(god)+" год по "+otchet_zex+" цеху в Excel");
      Abort();
    }

  if (otchet_zex==0) InsertLog("Формирование списка работников за "+IntToStr(god)+" год по предприятию в Excel успешно завершено");
  else InsertLog("Формирование списка работников за "+IntToStr(god)+" год по  "+otchet_zex+" цеху в Excel успешно завершено");

  DM->qLogs->Requery();

}
//---------------------------------------------------------------------------

//Формирование списка по цеху или предприятию в Excel
void __fastcall TMain::SpisokExcel2017(AnsiString otchet_zex)
{
  AnsiString sFile, Sql;
  int n=18;
  double  rezult, komp, effekt;
  Variant AppEx, Sh;


  Sql="select initcap(fio) as fio, initcap(fio_ocen) as fio_ocen, direkt,                                                                                 \
              (select naim_zex from sp_pdirekt pdir where o.direkt=pdir.zex and pdir.god="+IntToStr(god)+") as naim_zex,                                      \
              (select naim from sp_direkt where god="+IntToStr(god)+" and kod = (select kod_d from sp_pdirekt pdir1 where pdir1.zex=o.direkt and pdir1.god="+IntToStr(god)+")) naim_direkt,\
              (select name_ur1 from sap_osn_sved where tn_sap=o.tn                                                   \
               union all                                                                                             \
               select name_ur1 from sap_sved_uvol where tn_sap=o.tn) as uch,                                        \
               (select distinct nazv_cex from ssap_cex where id_cex=substr(orez.zex_rez,1,2) and nazv_cex not like '%(устар.)')  naim_zex_rez,        \
               (select stext from p1000@sapmig_buffdb where otype='O' and langu='R' and short=orez.zex_rez and stext not like '%(уст%') as naim_uch_rez,  \
               orez.zex_rez as zex_rez, \
              orez.shifr_rez as shifr_rez, \
               orez.dolg_rez, \
              o.*    \                                                                                                \
       from (     \
             (select * from ocenka where god="+IntToStr(god)+") o      \
             left join                                  \
             (select * from ocenka_rez where god="+IntToStr(god)+" and (tn,tn_sap_rez) in (select tn, min(tn_sap_rez) from ocenka_rez where god="+IntToStr(god)+" group by tn)) orez \
             on o.tn=orez.tn    \
            )";


       //from ocenka o left join ocenka_rez orez where o.god="+IntToStr(god)+" and orez.god="+IntToStr(god)+" and o.tn=orez.tn";
       // (select stext from p1000@sapmig_buffdb where otype='O' and langu='R' and objid=o.objid) as naim_uch_rez,

  if (otchet_zex==0)
    {
      Sql+=" order by zex, o.tn";
      StatusBar1->SimpleText=" Идет формирование списка работников по предприятию в Excel...";
    }
  else
    {
      Sql+=" where direkt ='"+otchet_zex+"' order by zex, o.tn";
      StatusBar1->SimpleText=" Идет формирование списка работников по "+otchet_zex+" цеху в Excel...";                                                                                                                                                \
    }
                                                                                                                                                                        \
  DM->qObnovlenie->Close();
  DM->qObnovlenie->SQL->Clear();
  DM->qObnovlenie->SQL->Add(Sql);
  try
    {
      DM->qObnovlenie->Open();
    }
  catch(Exception &E)
    {
      Application->MessageBox(("Возникла ошибка при получении данных из таблицы Ocenka" + E.Message).c_str(),"Ошибка",
                              MB_OK+MB_ICONERROR);
      if (otchet_zex==0) InsertLog("Возникла ошибка при формировании списка работников за "+IntToStr(god)+" год по предприятию в Excel");
      else InsertLog("Возникла ошибка при формировании списка работников за "+IntToStr(god)+" год по "+otchet_zex+" цеху в Excel");
      DM->qLogs->Requery();
      StatusBar1->SimpleText ="Отчетный период: "+IntToStr(god)+" год";
      Abort();
    }

 // sFile = Path+"\\RTF\\ocenka.xlt";

  Cursor = crHourGlass;
  ProgressBar->Position = 0;
  ProgressBar->Visible = true;
  ProgressBar->Max=DM->qObnovlenie->RecordCount;

  // инициализируем Excel, открываем этот шаблон
  try
    {
      AppEx=CreateOleObject("Excel.Application");
    }
  catch (...)
    {
      Application->MessageBox("Невозможно открыть Microsoft Excel!"
                              " Возможно это приложение на компьютере не установлено.","Ошибка",MB_OK+MB_ICONERROR);
      StatusBar1->SimpleText ="Отчетный период: "+IntToStr(god)+" год";
      ProgressBar->Visible = false;
      Cursor = crDefault;
    }

  //Если возникает ошибка во время формирования отчета
  try
    {
      try
        {
          AppEx.OlePropertySet("AskToUpdateLinks",false);
          AppEx.OlePropertySet("DisplayAlerts",false);

          //Копируем шаблон файла в Мои документы
          if (otchet_zex==0)
            {
              //Создание папки, если ее не существует
              ForceDirectories(WorkPath);
              CopyFile((Path+"\\RTF\\ocenka2017.xlsx").c_str(), (WorkPath+"\\Формирование списка по предприятию.xlsx").c_str(), false);
              sFile = WorkPath+"\\Формирование списка по предприятию.xlsx";
            }
          else
            {
              DeleteFile(WorkPath+"\\Формирование списка по 1 цеху");
              //Создание папки, если ее не существует
              ForceDirectories(WorkPath+"\\Формирование списка по 1 цеху");
              CopyFile((Path+"\\RTF\\ocenka2017.xlsx").c_str(), (WorkPath+"\\Формирование списка по 1 цеху\\Формирование списка по "+otchet_zex+" цеху.xlsx").c_str(), false);
              sFile = WorkPath+"\\Формирование списка по 1 цеху\\Формирование списка по "+otchet_zex+" цеху.xlsx";
            }
          AppEx.OlePropertyGet("WorkBooks").OleProcedure("Open",sFile.c_str())    ;  //открываем книгу, указав её имя

          Sh=AppEx.OlePropertyGet("WorkSheets",1);                               //выбираем № активного листа книги
          //Sh=AppEx.OlePropertyGet("WorkSheets","Расчет");                      //выбираем лист по наименованию
        }
      catch(...)
        {
          Application->MessageBox("Ошибка открытия книги Microsoft Excel!","Ошибка",MB_OK+MB_ICONERROR);
          StatusBar1->SimpleText ="Отчетный период: "+IntToStr(god)+" год";
          ProgressBar->Visible = false;
          Cursor = crDefault;
        }

      int i=1;
      n=19;

      Variant Massiv;
      Massiv = VarArrayCreate(OPENARRAY(int,(0,40)),varVariant); //массив на 31 элементов

      while (!DM->qObnovlenie->Eof)
        {
          if (otchet_zex==0) StatusBar1->SimpleText=" Идет формирование списка работников по предприятию в Excel: цех "+ DM->qObnovlenie->FieldByName("zex")->AsString;
          else StatusBar1->SimpleText=" Идет формирование списка по "+DM->qObnovlenie->FieldByName("direkt")->AsString+" цеху в Excel";

          rezult = (DM->qObnovlenie->FieldByName("realizac")->AsFloat+
                    DM->qObnovlenie->FieldByName("kachestvo")->AsFloat+
                    DM->qObnovlenie->FieldByName("resurs")->AsFloat)/12*100;
          komp = (DM->qObnovlenie->FieldByName("potreb")->AsFloat+
                  DM->qObnovlenie->FieldByName("stand")->AsFloat+
                  DM->qObnovlenie->FieldByName("kach")->AsFloat+
                  DM->qObnovlenie->FieldByName("eff")->AsFloat+
                  DM->qObnovlenie->FieldByName("prof_zn")->AsFloat+
                  DM->qObnovlenie->FieldByName("lider")->AsFloat+
                  DM->qObnovlenie->FieldByName("otvetstv")->AsFloat+
                  DM->qObnovlenie->FieldByName("kom_rez")->AsFloat)/32*100;

          if (rezult>0) effekt = ((rezult*0.6)+(komp*0.4));
          else if (!DM->qObnovlenie->FieldByName("kpe_ocen")->AsString.IsEmpty()) effekt = ((DM->qObnovlenie->FieldByName("kpe_ocen")->AsString*0.6)+(komp*0.4));
          else effekt=0;

          Massiv.PutElement(i, 0);
          Massiv.PutElement(DM->qObnovlenie->FieldByName("fio")->AsString.c_str(), 1);
          Massiv.PutElement(DM->qObnovlenie->FieldByName("tn")->AsString.c_str(), 2);
          Massiv.PutElement(DM->qObnovlenie->FieldByName("dolg")->AsString.c_str(), 3);
          Massiv.PutElement(DM->qObnovlenie->FieldByName("funct")->AsString.c_str(), 4);
          Massiv.PutElement(DM->qObnovlenie->FieldByName("uu")->AsString.c_str(), 5);
          Massiv.PutElement(DM->qObnovlenie->FieldByName("funct_g")->AsString.c_str(), 6);
          Massiv.PutElement(DM->qObnovlenie->FieldByName("kat")->AsString.c_str(), 7);
          Massiv.PutElement(DM->qObnovlenie->FieldByName("naim_direkt")->AsString.c_str(), 8);
          Massiv.PutElement(DM->qObnovlenie->FieldByName("zex")->AsString.c_str(), 9);
          Massiv.PutElement(DM->qObnovlenie->FieldByName("naim_zex")->AsString.c_str(), 10);
          Massiv.PutElement(DM->qObnovlenie->FieldByName("nuch")->AsString.c_str(), 11);

          //Критерии результатов работы
          if (DM->qObnovlenie->FieldByName("realizac")->AsFloat==0) Massiv.PutElement(DM->qObnovlenie->FieldByName("realizac")->AsString.c_str(), 12);
          else Massiv.PutElement(DM->qObnovlenie->FieldByName("realizac")->AsFloat, 12);
          if (DM->qObnovlenie->FieldByName("kachestvo")->AsFloat==0) Massiv.PutElement(DM->qObnovlenie->FieldByName("kachestvo")->AsString.c_str(), 13);
          else Massiv.PutElement(DM->qObnovlenie->FieldByName("kachestvo")->AsFloat, 13);
          if (DM->qObnovlenie->FieldByName("resurs")->AsFloat==0) Massiv.PutElement(DM->qObnovlenie->FieldByName("resurs")->AsString.c_str(), 14);
          else Massiv.PutElement(DM->qObnovlenie->FieldByName("resurs")->AsFloat, 14);

          //Оценка результатов работы
          if ((DM->qObnovlenie->FieldByName("realizac")->AsFloat+DM->qObnovlenie->FieldByName("kachestvo")->AsFloat+DM->qObnovlenie->FieldByName("resurs")->AsFloat)/3==0) Massiv.PutElement("", 15);
          else Massiv.PutElement((DM->qObnovlenie->FieldByName("realizac")->AsFloat+DM->qObnovlenie->FieldByName("kachestvo")->AsFloat+DM->qObnovlenie->FieldByName("resurs")->AsFloat)/3, 15);
          Massiv.PutElement(rezult, 16);
          if (DM->qObnovlenie->FieldByName("kpe_ocen")->AsFloat==0) Massiv.PutElement(DM->qObnovlenie->FieldByName("kpe_ocen")->AsString.c_str(), 17);
          else Massiv.PutElement(DM->qObnovlenie->FieldByName("kpe_ocen")->AsFloat, 17);

          //Критерии компетенций
          if (DM->qObnovlenie->FieldByName("eff")->AsFloat==0) Massiv.PutElement(DM->qObnovlenie->FieldByName("eff")->AsString.c_str(), 18);
          else Massiv.PutElement(DM->qObnovlenie->FieldByName("eff")->AsFloat, 18);
          if (DM->qObnovlenie->FieldByName("prof_zn")->AsFloat==0) Massiv.PutElement(DM->qObnovlenie->FieldByName("prof_zn")->AsString.c_str(), 19);
          else Massiv.PutElement(DM->qObnovlenie->FieldByName("prof_zn")->AsFloat, 19);
          if (DM->qObnovlenie->FieldByName("lider")->AsFloat==0) Massiv.PutElement(DM->qObnovlenie->FieldByName("lider")->AsString.c_str(), 20);
          else Massiv.PutElement(DM->qObnovlenie->FieldByName("lider")->AsFloat, 20);
          if (DM->qObnovlenie->FieldByName("otvetstv")->AsFloat==0) Massiv.PutElement(DM->qObnovlenie->FieldByName("otvetstv")->AsString.c_str(), 21);
          else Massiv.PutElement(DM->qObnovlenie->FieldByName("otvetstv")->AsFloat, 21);
          if (DM->qObnovlenie->FieldByName("kom_rez")->AsFloat==0) Massiv.PutElement(DM->qObnovlenie->FieldByName("kom_rez")->AsString.c_str(), 22);
          else Massiv.PutElement(DM->qObnovlenie->FieldByName("kom_rez")->AsFloat, 22);
          if (DM->qObnovlenie->FieldByName("stand")->AsFloat==0) Massiv.PutElement(DM->qObnovlenie->FieldByName("stand")->AsString.c_str(), 23);
          else Massiv.PutElement(DM->qObnovlenie->FieldByName("stand")->AsFloat, 23);
          if (DM->qObnovlenie->FieldByName("potreb")->AsFloat==0) Massiv.PutElement(DM->qObnovlenie->FieldByName("potreb")->AsString.c_str(), 24);
          else Massiv.PutElement(DM->qObnovlenie->FieldByName("potreb")->AsFloat, 24);
          if (DM->qObnovlenie->FieldByName("kach")->AsFloat==0) Massiv.PutElement(DM->qObnovlenie->FieldByName("kach")->AsString.c_str(), 25);
          else Massiv.PutElement(DM->qObnovlenie->FieldByName("kach")->AsFloat, 25);


          //Оценка компетенций
          if (((DM->qObnovlenie->FieldByName("potreb")->AsFloat+
                DM->qObnovlenie->FieldByName("stand")->AsFloat+
                DM->qObnovlenie->FieldByName("kach")->AsFloat+
                DM->qObnovlenie->FieldByName("eff")->AsFloat+
                DM->qObnovlenie->FieldByName("prof_zn")->AsFloat+
                DM->qObnovlenie->FieldByName("lider")->AsFloat+
                DM->qObnovlenie->FieldByName("otvetstv")->AsFloat+
                DM->qObnovlenie->FieldByName("kom_rez")->AsFloat))==0) Massiv.PutElement("", 26);
          else Massiv.PutElement(((DM->qObnovlenie->FieldByName("potreb")->AsFloat+
                                   DM->qObnovlenie->FieldByName("stand")->AsFloat+
                                   DM->qObnovlenie->FieldByName("kach")->AsFloat+
                                   DM->qObnovlenie->FieldByName("eff")->AsFloat+
                                   DM->qObnovlenie->FieldByName("prof_zn")->AsFloat+
                                   DM->qObnovlenie->FieldByName("lider")->AsFloat+
                                   DM->qObnovlenie->FieldByName("otvetstv")->AsFloat+
                                   DM->qObnovlenie->FieldByName("kom_rez")->AsFloat)), 26);
          Massiv.PutElement(komp, 27);
          Massiv.PutElement(effekt, 28);

          //Итоговая оценка
          Massiv.PutElement(DM->qObnovlenie->FieldByName("avt_reit")->AsString.c_str(), 29);
          Massiv.PutElement(DM->qObnovlenie->FieldByName("skor_reit")->AsString.c_str(), 30);
          Massiv.PutElement(DM->qObnovlenie->FieldByName("kom_reit")->AsString.c_str(), 31);

          if (DM->qObnovlenie->FieldByName("rezerv")->AsString==1) Massiv.PutElement("да", 32);
          else Massiv.PutElement("нет", 32);
          Massiv.PutElement(DM->qObnovlenie->FieldByName("dolg_rez")->AsString.c_str(), 33);

          Massiv.PutElement(DM->qObnovlenie->FieldByName("naim_zex_rez")->AsString.c_str(), 34);
          Massiv.PutElement(DM->qObnovlenie->FieldByName("naim_uch_rez")->AsString.c_str(), 35);
          Massiv.PutElement(DM->qObnovlenie->FieldByName("zex_rez")->AsString.c_str(), 36);
          Massiv.PutElement(DM->qObnovlenie->FieldByName("shifr_rez")->AsString.c_str(), 37);

          Massiv.PutElement(DM->qObnovlenie->FieldByName("fio_ocen")->AsString.c_str(), 38);
          Massiv.PutElement(DM->qObnovlenie->FieldByName("dolg_ocen")->AsString.c_str(), 39);
          Massiv.PutElement(DM->qObnovlenie->FieldByName("data_ocen")->AsString.c_str(), 40);


          Sh.OlePropertyGet("Range", ("A" + IntToStr(n) + ":AO" + IntToStr(n)).c_str()).OlePropertySet("Value", Massiv); //строка с данными с ячейки A по ячейку АВ

          i++;
          n++;
          DM->qObnovlenie->Next();
          ProgressBar->Position++;
        }

       // Sh.OlePropertyGet("Range", ("LQ18:LQ" + IntToStr(i-1)).c_str()).OlePropertySet("NumberFormat", "0.00");

      //окрашивание ячеек
      Sh.OlePropertyGet("Range",("Q18:Q"+IntToStr(n-1)).c_str()).OlePropertyGet("Interior").OlePropertySet("Color",14277081);
      Sh.OlePropertyGet("Range",("AB18:AD"+IntToStr(n-1)).c_str()).OlePropertyGet("Interior").OlePropertySet("Color",14277081);

      Sh.OlePropertyGet("Range",("B18:L"+IntToStr(n-1)).c_str()).OlePropertyGet("Interior").OlePropertySet("Color",14408946);
      Sh.OlePropertyGet("Range",("R18:R"+IntToStr(n-1)).c_str()).OlePropertyGet("Interior").OlePropertySet("Color",14408946);

      //рисуем сетку
      Sh.OlePropertyGet("Range",("A18:AP"+IntToStr(n-1)).c_str()).OlePropertyGet("Borders").OlePropertySet("LineStyle", xlContinuous);

      //Сохранить книгу в папке в файле по указанию
     // AnsiString vAsCurDir1=WorkPath+"\\Формирование списка по предприятию";

     // Sh.OleProcedure("SaveAs",vAsCurDir1.c_str());
     AppEx.OlePropertyGet("WorkBooks",1).OleFunction("Save");

      /* //Закрыть открытое приложение Excel
      AppEx.OleProcedure("Quit");
      AppEx = Unassigned;  */

      //Закрыть книгу Excel с шаблоном для вывода информации
     // AppEx.OlePropertyGet("WorkBooks",1).OleProcedure("Close");
      Application->MessageBox("Отчет успешно сформирован!", "Формирование отчета",
                               MB_OK+MB_ICONINFORMATION);
      //AppEx.OlePropertyGet("WorkBooks").OleProcedure("Open",vAsCurDir1.c_str());
      AppEx.OlePropertySet("Visible",true);
      AppEx.OlePropertySet("AskToUpdateLinks",true);
      AppEx.OlePropertySet("DisplayAlerts",true);

      StatusBar1->SimpleText= "Формирование отчета выполнено.";

      Cursor = crDefault;
      ProgressBar->Position=0;
      ProgressBar->Visible = false;
      StatusBar1->SimpleText ="Отчетный период: "+IntToStr(god)+" год";
    }
  catch(...)
    {
      AppEx.OleProcedure("Quit");
      AppEx = Unassigned;
      Cursor = crDefault;
      ProgressBar->Position=0;
      ProgressBar->Visible = false;

      StatusBar1->SimpleText ="Отчетный период: "+IntToStr(god)+" год";
      if (otchet_zex==0) InsertLog("Возникла ошибка при формировании списка работников за "+IntToStr(god)+" год по предприятию в Excel");
      else InsertLog("Возникла ошибка при формировании списка работников за "+IntToStr(god)+" год по "+otchet_zex+" цеху в Excel");
      Abort();
    }

  if (otchet_zex==0) InsertLog("Формирование списка работников за "+IntToStr(god)+" год по предприятию в Excel успешно завершено");
  else InsertLog("Формирование списка работников за "+IntToStr(god)+" год по  "+otchet_zex+" цеху в Excel успешно завершено");

  DM->qLogs->Requery();

}
//---------------------------------------------------------------------------

//Логи
void __fastcall TMain::InsertLog(AnsiString Msg)
{
  AnsiString Data;
  DateTimeToString(Data, "dd.mm.yyyy hh:nn:ss", Now());
  DM->qObnovlenie->SQL->Clear();
  DM->qObnovlenie->SQL->Add("insert into logs_ocenka (DT,DOMAIN,USEROK,PROG,USEROK_FIO,TEXT) values \
                            (to_date(" + QuotedStr(Data) + ", 'DD.MM.YYYY HH24:MI:SS'),\
                             " + QuotedStr(DomainName) + "," + QuotedStr(UserName) + ", 'Ocenka',\
                             " + QuotedStr(UserFullName)+",  \
                             " + QuotedStr(Msg)+")");
  try
    {
      DM->qObnovlenie->ExecSQL();
    }
  catch(...)
    {
      MessageBox(Handle,"Возникла ошибка при вставке данных в таблицу LOGS_OCENKA","Ошибка",8202);
    }

  DM->qObnovlenie->Close();
}
 //---------------------------------------------------------------------------

//Список работников с не предоставленными формами ЕОП
void __fastcall TMain::N15Click(TObject *Sender)
{
  AnsiString Sql, row;
  Variant AppEx, Sh;

  StatusBar1->SimpleText=" Идет формирование отчета...";

  Sql="select initcap(fio) as fio, \
              tn,                  \
              dolg,                \
              (select naim from sp_direkt where god="+IntToStr(god)+" and kod = (select kod_d from sp_pdirekt pdir where pdir.zex=o.direkt and pdir.god="+IntToStr(god)+")) direkt, \
              zex,                                                                                               \
              (select naim_zex from sp_pdirekt pdir1 where o.direkt=pdir1.zex and pdir1.god="+IntToStr(god)+") as naim_zex,                                  \
              fio_ocen,  \
              dolg_ocen, \
              data_ocen  \
       from ocenka o      \
       where data_ocen<=to_date(sysdate)-3 and data_ocen is not null and (nvl(komp_ocen,0)=0 or (nvl(rezult_ocen,0)=0 and nvl(kpe_ocen,0)=0)) \
       and god="+IntToStr(god)+" \
       order by zex, data_ocen";

  DM->qObnovlenie->Close();
  DM->qObnovlenie->SQL->Clear();
  DM->qObnovlenie->SQL->Add(Sql);
  try
    {
      DM->qObnovlenie->Open();
    }
  catch(Exception &E)
    {
      Application->MessageBox(("Возникла ошибка при получении данных из таблицы Ocenka" + E.Message).c_str(),"Ошибка",
                              MB_OK+MB_ICONERROR);
      StatusBar1->SimpleText ="Отчетный период: "+IntToStr(god)+" год";
      Abort();
    }

  if (DM->qObnovlenie->RecordCount>0)
    {
      Cursor = crHourGlass;
      ProgressBar->Position = 0;
      ProgressBar->Visible = true;
      ProgressBar->Max=DM->qObnovlenie->RecordCount;

      //Открытие документа Excel
      try
        {
          AppEx = CreateOleObject("Excel.Application");
        }
      catch (...)
        {
          Application->MessageBox("Невозможно открыть Microsoft Excel!"
                                  " Возможно это приложение на компьютере не установлено.","Ошибка",MB_OK+MB_ICONERROR);
          Abort();
        }

      //Если возникает ошибка во время формирования отчета
      try
        {
          try
            {
              AppEx.OlePropertySet("AskToUpdateLinks",false);
              AppEx.OlePropertySet("DisplayAlerts",false);
              AppEx.OlePropertyGet("WorkBooks").OleProcedure("Open",(Path +"\\RTF\\dolgniki.xlt").c_str())    ;  //открываем книгу, указав её имя

              Sh=AppEx.OlePropertyGet("WorkSheets",1);                               //выбираем № активного листа книги
            }
          catch(...)
            {
              Application->MessageBox("Ошибка открытия книги Microsoft Excel!","Ошибка",MB_OK + MB_ICONERROR);
            }

          row = DM->qObnovlenie->RecordCount+1;

          // выводим в шаблон данные
          // вставляем в шаблон нужное количество строк
          Variant C;
          Sh.OleProcedure("Select");
          C=Sh.OlePropertyGet("Range","zex");
          C=Sh.OlePropertyGet("Rows",(int) C.OlePropertyGet("Row")+1);
          for(int i=1;i<row;i++) C.OleProcedure("Insert");
          int i=1, n=8;

          //вывод даты
          toExcel(Sh,"data",String(Date()));

          while(!DM->qObnovlenie->Eof)
            {
              //вывод данных
              toExcel(Sh,"nn",i,i);
              toExcel(Sh,"fio",i, DM->qObnovlenie->FieldByName("fio")->AsString.c_str());
              toExcel(Sh,"tn",i, DM->qObnovlenie->FieldByName("tn")->AsInteger);
              toExcel(Sh,"dolg",i,DM->qObnovlenie->FieldByName("dolg")->AsString.c_str());
              toExcel(Sh,"direkt",i, DM->qObnovlenie->FieldByName("direkt")->AsString.c_str());
              toExcel(Sh,"zex",i, DM->qObnovlenie->FieldByName("zex")->AsString);
              toExcel(Sh,"naim_zex",i,DM->qObnovlenie->FieldByName("naim_zex")->AsString.c_str());
              toExcel(Sh,"fio_ocen",i, DM->qObnovlenie->FieldByName("fio_ocen")->AsString.c_str());
              toExcel(Sh,"dolg_ocen",i, DM->qObnovlenie->FieldByName("dolg_ocen")->AsString.c_str());
              toExcel(Sh,"data_ocen",i,DM->qObnovlenie->FieldByName("data_ocen")->AsString.c_str());
              i++;
              n++;

              DM->qObnovlenie->Next();
              ProgressBar->Position++;
            }

          //рисуем сетку
          Sh.OlePropertyGet("Range",("A7:J"+IntToStr(n-1)).c_str()).OlePropertyGet("Borders").OlePropertySet("LineStyle", xlContinuous);

          //Отключить вывод сообщений с вопросами типа "Заменить файл..."
          //AppEx.OlePropertySet("DisplayAlerts",false);

          //Создание папки, если ее не существует
          ForceDirectories(Main->WorkPath);

          //Сохранить книгу в папке в файле по указанию
          AnsiString vAsCurDir1=WorkPath+"\\Список работников предприятия с не предоставленными формами ЕОП.xls";
          Sh.OleProcedure("SaveAs",vAsCurDir1.c_str());

          //Закрыть книгу Excel с шаблоном для вывода информации
          AppEx.OlePropertyGet("WorkBooks",1).OleProcedure("Close");
          Application->MessageBox("Отчет успешно сформирован!", "Формирование отчета",
                                   MB_OK+MB_ICONINFORMATION);
          AppEx.OlePropertyGet("WorkBooks").OleProcedure("Open",vAsCurDir1.c_str());
          AppEx.OlePropertySet("Visible",true);
          AppEx.OlePropertySet("AskToUpdateLinks",true);
          AppEx.OlePropertySet("DisplayAlerts",true);

          Cursor = crDefault;
          ProgressBar->Position = 0;
          ProgressBar->Visible = false;
          StatusBar1->SimpleText ="Отчетный период: "+IntToStr(god)+" год";
        }
      catch (...)
        {
          AppEx.OleProcedure("Quit");
          AppEx = Unassigned;
        }
    }
  else
    {
      Application->MessageBox("Нет данных для формирования отчета!", "Формирование отчета",
                              MB_OK+MB_ICONINFORMATION);
      StatusBar1->SimpleText ="Отчетный период: "+IntToStr(god)+" год";
    }

}
//---------------------------------------------------------------------------

//Статус ЕОП на ММКИ
void __fastcall TMain::N16Click(TObject *Sender)
{
  AnsiString Sql, row;
  Variant AppEx, Sh;

  StatusBar1->SimpleText=" Идет формирование отчета...";

  Sql = "select sum(ob_kol1) ob_kol, sum(kpe_kol1) kpe_kol, sum(rr_kol1) rr_kol, sum(p_rukov1) p_rukov, sum(p_linmen1) p_linmen,                                     \
                sum(p_sotrud1) p_sotrud, sum(v_rukov1) v_rukov, sum(v_linmen1) v_linmen, sum(v_sotrud1) v_sotrud, sum(efect_kol1) efect_kol     \
         from (                                                                                                                                 \
                select count(*) ob_kol1, 0 kpe_kol1, 0 rr_kol1, 0 p_rukov1, 0 p_linmen1, 0 p_sotrud1, 0 v_rukov1, 0 v_linmen1, 0 v_sotrud1, 0 efect_kol1   \
                from ocenka where god="+IntToStr(god)+"                                                                                                                    \
                union all                                                                                                                       \
                select 0 ob_kol1, count(*) kpe_kol1, 0 rr_kol1, 0 p_rukov1, 0 p_linmen1, 0 p_sotrud1, 0 v_rukov1, 0 v_linmen1, 0 v_sotrud1, 0 efect_kol1   \
                from ocenka where nvl(kpe_ocen,0)>0 and god="+IntToStr(god)+"                                                                                            \
                union all                                                                                                                       \
                select 0 ob_kol1, 0 kpe_kol1, count(*) rr_kol1, 0 p_rukov1, 0 p_linmen1, 0 p_sotrud1, 0 v_rukov1, 0 v_linmen1, 0 v_sotrud1, 0 efect_kol1   \
                from ocenka where nvl(rezult_ocen,0)>0  and god="+IntToStr(god)+" \
                union all                                                                                                                       \
                select 0 ob_kol1, 0 kpe_kol1, 0 rr_kol1, count(*) p_rukov1, 0 p_linmen1, 0 p_sotrud1, 0 v_rukov1, 0 v_linmen1, 0 v_sotrud1, 0 efect_kol1   \
                from ocenka where funct_g='Производство' and kat='руководитель подразделения' and god="+IntToStr(god)+"                                                  \
                union all                                                                                                                       \
                select 0 ob_kol1, 0 kpe_kol1, 0 rr_kol1, 0 p_rukov1, count(*) p_linmen1, 0 p_sotrud1, 0 v_rukov1, 0 v_linmen1, 0 v_sotrud1, 0 efect_kol1   \
                from ocenka where funct_g='Производство' and kat='линейный менеджер' and god="+IntToStr(god)+"                                                           \
                union all                                                                                                                       \
                select 0 ob_kol1, 0 kpe_kol1, 0 rr_kol1, 0 p_rukov1, 0 p_linmen1, count(*) p_sotrud1, 0 v_rukov1, 0 v_linmen1, 0 v_sotrud1, 0 efect_kol1   \
                from ocenka where funct_g='Производство' and kat='сотрудник' and god="+IntToStr(god)+"                                                                   \
                union all                                                                                                                       \
                select 0 ob_kol1, 0 kpe_kol1, 0 rr_kol1, 0 p_rukov1, 0 p_linmen1, 0 p_sotrud1, count(*) v_rukov1, 0 v_linmen1, 0 v_sotrud1, 0 efect_kol1   \
                from ocenka where funct_g='Внутренний сервис. Продажи' and kat='руководитель подразделения' and god="+IntToStr(god)+"                                     \
                union all                                                                                                                       \
                select 0 ob_kol1, 0 kpe_kol1, 0 rr_kol1, 0 p_rukov1, 0 p_linmen1, 0 p_sotrud1, 0 v_rukov1, count(*) v_linmen1, 0 v_sotrud1, 0 efect_kol1   \
                from ocenka where funct_g='Внутренний сервис. Продажи' and kat='линейный менеджер' and god="+IntToStr(god)+"                                              \
                union all                                                                                                                       \
                select 0 ob_kol1, 0 kpe_kol1, 0 rr_kol1, 0 p_rukov1, 0 p_linmen1, 0 p_sotrud1, 0 v_rukov1, 0 v_linmen1, count(*) v_sotrud1, 0 efect_kol1   \
                from ocenka where funct_g='Внутренний сервис. Продажи' and kat='сотрудник' and god="+IntToStr(god)+"                                                     \
                union all                                                                                                                       \
                select 0 ob_kol1, 0 kpe_kol1, 0 rr_kol1, 0 p_rukov1, 0 p_linmen1, 0 p_sotrud1, 0 v_rukov1, 0 v_linmen1, 0 v_sotrud1, count(*) efect_kol1   \
                from ocenka where efect is not null and efect!=0 and god="+IntToStr(god)+"                                                                                           \
              )";
  DM->qObnovlenie->Close();
  DM->qObnovlenie->SQL->Clear();
  DM->qObnovlenie->SQL->Add(Sql);

  try
    {
      DM->qObnovlenie->Open();
    }
  catch(Exception &E)
    {
      Application->MessageBox(("Возникла ошибка при получении данных из таблицы Ocenka" + E.Message).c_str(),"Ошибка",
                              MB_OK+MB_ICONERROR);
      StatusBar1->SimpleText ="Отчетный период: "+IntToStr(god)+" год";
      Abort();
    }

  Cursor = crHourGlass;
  ProgressBar->Position = 0;
  ProgressBar->Visible = true;
  ProgressBar->Max=DM->qObnovlenie->RecordCount;

  //Открытие документа Excel
  try
    {
      AppEx = CreateOleObject("Excel.Application");
    }
  catch (...)
    {
      Application->MessageBox("Невозможно открыть Microsoft Excel!"
                              " Возможно это приложение на компьютере не установлено.","Ошибка",MB_OK+MB_ICONERROR);
      Abort();
    }

  //Если возникает ошибка во время формирования отчета
  try
    {
      try
        {
          AppEx.OlePropertySet("AskToUpdateLinks",false);
          AppEx.OlePropertySet("DisplayAlerts",false);
          AppEx.OlePropertyGet("WorkBooks").OleProcedure("Open",(Path +"\\RTF\\statusEOP.xlt").c_str())    ;  //открываем книгу, указав её имя

          Sh=AppEx.OlePropertyGet("WorkSheets",1);                               //выбираем № активного листа книги
        }
      catch(...)
        {
          Application->MessageBox("Ошибка открытия книги Microsoft Excel!","Ошибка",MB_OK + MB_ICONERROR);
        }

      row = DM->qObnovlenie->RecordCount+1;

      // выводим в шаблон данные
      // вставляем в шаблон нужное количество строк
      Variant C;
      Sh.OleProcedure("Select");
      // C=Sh.OlePropertyGet("Range","zex");
      // C=Sh.OlePropertyGet("Rows",(int) C.OlePropertyGet("Row")+1);
      //  for(int i=1;i<row;i++) C.OleProcedure("Insert");
      int i=0, n=7;

      //вывод даты
      toExcel(Sh,"data",String(Date()));

      while(!DM->qObnovlenie->Eof)
        {
          //вывод данных
          //toExcel(Sh,"ob_kol",i,i+1);
          toExcel(Sh,"ob_kol", DM->qObnovlenie->FieldByName("ob_kol")->AsInteger);
          toExcel(Sh,"kpe_kol", DM->qObnovlenie->FieldByName("kpe_kol")->AsInteger);
          toExcel(Sh,"rr_kol", DM->qObnovlenie->FieldByName("rr_kol")->AsInteger);
          toExcel(Sh,"p_rukov",DM->qObnovlenie->FieldByName("p_rukov")->AsInteger);
          toExcel(Sh,"p_linmen", DM->qObnovlenie->FieldByName("p_linmen")->AsInteger);
          toExcel(Sh,"p_sotrud", DM->qObnovlenie->FieldByName("p_sotrud")->AsInteger);
          toExcel(Sh,"v_rukov",DM->qObnovlenie->FieldByName("v_rukov")->AsInteger);
          toExcel(Sh,"v_linmen", DM->qObnovlenie->FieldByName("v_linmen")->AsInteger);
          toExcel(Sh,"v_sotrud", DM->qObnovlenie->FieldByName("v_sotrud")->AsInteger);
          toExcel(Sh,"efect_kol",DM->qObnovlenie->FieldByName("efect_kol")->AsInteger);
          i++;
          n++;

          DM->qObnovlenie->Next();
          ProgressBar->Position++;
        }

      //рисуем сетку
      Sh.OlePropertyGet("Range",("A7:J"+IntToStr(n-1)).c_str()).OlePropertyGet("Borders").OlePropertySet("LineStyle", xlContinuous);

      //Отключить вывод сообщений с вопросами типа "Заменить файл..."
      //AppEx.OlePropertySet("DisplayAlerts",false);

      //Создание папки, если ее не существует
      ForceDirectories(Main->WorkPath);

      //Сохранить книгу в папке в файле по указанию
      AnsiString vAsCurDir1=WorkPath+"\\Статус ЕОП на ММКИ.xls";
      Sh.OleProcedure("SaveAs",vAsCurDir1.c_str());

      //Закрыть книгу Excel с шаблоном для вывода информации
      AppEx.OlePropertyGet("WorkBooks",1).OleProcedure("Close");
      Application->MessageBox("Отчет успешно сформирован!", "Формирование отчета",
                             MB_OK+MB_ICONINFORMATION);
      AppEx.OlePropertyGet("WorkBooks").OleProcedure("Open",vAsCurDir1.c_str());
      AppEx.OlePropertySet("Visible",true);
      AppEx.OlePropertySet("AskToUpdateLinks",true);
      AppEx.OlePropertySet("DisplayAlerts",true);
    }
  catch(...)
    {
      AppEx.OleProcedure("Quit");
      AppEx = Unassigned;
    }
    
  Cursor = crDefault;
  ProgressBar->Position = 0;
  ProgressBar->Visible = false;
  StatusBar1->SimpleText ="Отчетный период: "+IntToStr(god)+" год";

}
//---------------------------------------------------------------------------

//Расчет рейтинга по 1 работнику
void __fastcall TMain::N18Click(TObject *Sender)
{
  AnsiString Sql;
  int rec;

  rec = DM->qOcenka->RecNo;
  StatusBar1->SimpleText=" Идет расчет рейтинга...";
  
  //Проверить кол-во записей >=5
  Sql = "select distinct upper(fio_ocen) as fio_ocen,                                                       \
                upper(kat) as kat,                                                                     \
                upper(decode(funct_g,NULL,'1',funct_g)) as funct_g,                                                                 \
                a_min,                                                                   \
                a_max,                                                                   \
                kol*0.05 as A5,                                                          \
                kol*0.2 as A20,                                                          \
                kol*0.6 as B60,                                                          \
                kol*0.2 as C20,                                                          \
                kol*0.05 as C5,                                                          \
                kpe,\
                god                                                                      \
          from (                                                                         \
                 select  min(efect) over (partition by kat,funct_g,fio_ocen, kpe) a_min, \
                         max(efect) over (partition by kat,funct_g,fio_ocen, kpe) a_max, \
                         count(*) over (partition by kat, funct_g, fio_ocen, kpe) kol,   \
                         d.*                                                             \
                 from (                                                                  \
                        select decode (nvl(kpe_ocen,0),0,0,1) as kpe,                    \
                        o.*                                                              \
                        from ocenka o where nvl(efect,0)!=0 and god="+IntToStr(god)+"                       \
                        and upper(fio_ocen)=upper("+QuotedStr(DM->qOcenka->FieldByName("fio_ocen")->AsString)+")    \
                        and upper(kat)=upper("+QuotedStr(DM->qOcenka->FieldByName("kat")->AsString)+")              \
                        and upper(decode(funct_g,NULL,'1',funct_g))=upper(decode("+QuotedStr(DM->qOcenka->FieldByName("funct_g")->AsString)+",NULL,'1',funct_g))      \
                        and decode (nvl(kpe_ocen,0),0,0,1) = decode (nvl("+SetNull(DM->qOcenka->FieldByName("kpe_ocen")->AsString)+",0),0,0,1)\
                      ) d                                                                \
               )                                                                         \
          where kol>=5                                                                   \
          order by fio_ocen, kat, funct_g";

  DM->qObnovlenie->Close();
  DM->qObnovlenie->SQL->Clear();
  DM->qObnovlenie->SQL->Add(Sql);
  try
    {
      DM->qObnovlenie->Open();
    }
  catch(Exception &E)
    {
      Application->MessageBox(("Возникла ошибка при обновлении получении данных из таблицы Ocenka" + E.Message).c_str(),"Ошибка",
                              MB_OK+MB_ICONERROR);
      StatusBar1->SimpleText ="Отчетный период: "+IntToStr(god)+" год";
      Abort();
    }

  if (DM->qObnovlenie->RecordCount>0)
    {
      Cursor = crHourGlass;
      ProgressBar->Position = 0;
      ProgressBar->Visible = true;
      ProgressBar->Max=DM->qObnovlenie->RecordCount;

      //update всех значений (обнуление автоматического рейтинга)
      Sql = "update ocenka set avt_reit=NULL                                                        \
             where (zex,tn) in (select zex,tn                                                           \
                                from ( select zex,                                                           \
                                              tn,                                                            \
                                              min(efect) over (partition by kat,funct_g,fio_ocen,kpe) a_min, \
                                              count(*) over (partition by kat, funct_g, fio_ocen,kpe) kol    \
                                       from ( select decode(nvl(kpe_ocen,0),0,0,1) as kpe,                   \
                                                     o.*                                                            \
                                              from ocenka o                                            \
                                              where nvl(efect,0)!=0 \
                                              and god="+IntToStr(god)+"                                                       \
                                              and upper(fio_ocen)=upper("+QuotedStr(DM->qOcenka->FieldByName("fio_ocen")->AsString)+")    \
                                              and upper(kat)=upper("+QuotedStr(DM->qOcenka->FieldByName("kat")->AsString)+")              \
                                              and upper(decode(funct_g,NULL,'1',funct_g))=upper(decode("+QuotedStr(DM->qOcenka->FieldByName("funct_g")->AsString)+",NULL,'1',funct_g))     \
                                              and decode (nvl(kpe_ocen,0),0,0,1) = decode (nvl("+SetNull(DM->qOcenka->FieldByName("kpe_ocen")->AsString)+",0),0,0,1)\                                         \
                                            )                                                                \
                                     ) d \
                                where kol>=5)";

      DM->qObnovlenie2->Close();
      DM->qObnovlenie2->SQL->Clear();
      DM->qObnovlenie2->SQL->Add(Sql);
      try
        {
          DM->qObnovlenie2->ExecSQL();
        }
      catch(Exception &E)
        {
          Application->MessageBox(("Возникла ошибка при обновлении автоматического рейтинга в таблице Ocenka" + E.Message).c_str(),"Ошибка",
                                   MB_OK+MB_ICONERROR);
          ProgressBar->Visible = false;
         StatusBar1->SimpleText ="Отчетный период: "+IntToStr(god)+" год";
          Abort();
        }

      while (!DM->qObnovlenie->Eof)
        {
          //Выполнение продцедуры по расчету рейтинга
          DM->spOcenka->Parameters->ParamByName("pFio_ocen")->Value = DM->qObnovlenie->FieldByName("fio_ocen")->AsString;
          DM->spOcenka->Parameters->ParamByName("pKat")->Value = DM->qObnovlenie->FieldByName("kat")->AsString;
          DM->spOcenka->Parameters->ParamByName("pFunct_g")->Value = DM->qObnovlenie->FieldByName("funct_g")->AsString;
          DM->spOcenka->Parameters->ParamByName("pMin")->Value = DM->qObnovlenie->FieldByName("a_min")->AsFloat;
          DM->spOcenka->Parameters->ParamByName("pMax")->Value = DM->qObnovlenie->FieldByName("a_max")->AsFloat;
          DM->spOcenka->Parameters->ParamByName("PA5")->Value = DM->qObnovlenie->FieldByName("A5")->AsFloat;
          DM->spOcenka->Parameters->ParamByName("PA20")->Value = DM->qObnovlenie->FieldByName("A20")->AsFloat;
          DM->spOcenka->Parameters->ParamByName("PB60")->Value = DM->qObnovlenie->FieldByName("B60")->AsFloat;
          DM->spOcenka->Parameters->ParamByName("PC20")->Value = DM->qObnovlenie->FieldByName("C20")->AsFloat;
          DM->spOcenka->Parameters->ParamByName("PC5")->Value = DM->qObnovlenie->FieldByName("C5")->AsFloat;
          DM->spOcenka->Parameters->ParamByName("pKpe")->Value = DM->qObnovlenie->FieldByName("kpe")->AsFloat;
          DM->spOcenka->Parameters->ParamByName("pGod")->Value = DM->qObnovlenie->FieldByName("god")->AsInteger;

          try
            {
              DM->spOcenka->ExecProc();
            }
          catch(Exception &E)
            {
              Application->MessageBox(("Возникла ошибка при расчете автоматического рейтинга (продцедура CALC_OCENKA)" + E.Message).c_str(),"Ошибка",
                                       MB_OK+MB_ICONERROR);
              InsertLog("Расчет рейтинга за "+IntToStr(god)+" год: возникла ошибка при расчете рейтинга по работнику "+DM->qOcenka->FieldByName("fio")->AsString+" и другим работникам из его группы по рейтингу");
              DM->qLogs->Requery();
              StatusBar1->SimpleText ="Отчетный период: "+IntToStr(god)+" год";
              ProgressBar->Visible = false;
              Abort();
            }

          DM->qObnovlenie->Next();
          ProgressBar->Position++;
        }

      Application->MessageBox(("Расчет рейтинга за "+IntToStr(god)+" год выполнен успешно по работнику "+DM->qOcenka->FieldByName("fio")->AsString+" и другим работникам из его группы по рейтингу (оценщик "+DM->qOcenka->FieldByName("fio_ocen")->AsString+")").c_str(),"Ошибка",
                                        MB_OK+MB_ICONINFORMATION);
      DM->qOcenka->Requery();
      InsertLog("Расчет рейтинга за "+IntToStr(god)+" год выполнен успешно по работнику "+DM->qOcenka->FieldByName("fio")->AsString+" и другим работникам из его группы по рейтингу");
      DM->qLogs->Requery();

      //Вернуть на запись
      DM->qOcenka->RecNo = rec;

      Cursor = crDefault;
      ProgressBar->Position = 0;
      ProgressBar->Visible = false;
      StatusBar1->SimpleText ="Отчетный период: "+IntToStr(god)+" год";
    }
  else
    {
      Application->MessageBox("Невозможно расчитать рейтинг по данному работнику, \nтак как количество записей менее 5 в пределах \nодного оценщика, категории и функциональной группы. ","Ошибка",
                               MB_OK+MB_ICONINFORMATION);
      StatusBar1->SimpleText ="Отчетный период: "+IntToStr(god)+" год";
    }

}
//---------------------------------------------------------------------------

void __fastcall TMain::FormShow(TObject *Sender)
{
  //Год для замещений
  DM->qZamesh->Parameters->ParamByName("pgod")->Value=IntToStr(Main->god);
  DM->qZamesh->Active = true;


  //Очистка полей для загрузки из Excel
  Zagruzka->EditDATA->Text = "";
  Zagruzka->EditFIO->Text = "";
  Zagruzka->EditDOLGO->Text = "";
  Zagruzka->EditOCENKA->Text = "";
  Zagruzka->EditREZERV->Text = "";
  Zagruzka->EditDOLG->Text = "";
  Zagruzka->EditZEX->Text = "";
  Zagruzka->EditTN->Text = "";
  Zagruzka->EditREZULT_OCEN->Text = "";
  Zagruzka->EditKPE_OCEN->Text = "";
  Zagruzka->EditKOMP_OCEN->Text = "";
  Zagruzka->EditFIOEOP->Text = "";
  Zagruzka->EditTNEOP->Text = "";
  Zagruzka->EditTN_KPE->Text = "";
  Zagruzka->EditKPE1->Text = "";
  Zagruzka->EditKPE2->Text = "";
  Zagruzka->EditKPE3->Text = "";
  Zagruzka->EditKPE4->Text = "";
  Zagruzka->EditTN_VZ->Text = "";
  Zagruzka->EditVZ->Text = "";

  Zagruzka->EditKR_ZEX->Text = "";
  Zagruzka->EditTN_KR->Text = "";
  Zagruzka->EditKR_FIO->Text = "";
  Zagruzka->EditKRSHIFR_DOLG->Text = "";
}
//---------------------------------------------------------------------------

//Загрузка скорректированной руководителем оценки
void __fastcall TMain::N19Click(TObject *Sender)
{
  Zagruzka->CheckBoxOCENKA->Checked = true;
  Zagruzka->SpeedButton1Click(Sender);
}
//---------------------------------------------------------------------------

//Загрузка рекомендаций в кадровый резерв
void __fastcall TMain::N20Click(TObject *Sender)
{
  Zagruzka->CheckBoxREZERV->Checked = true;
  Zagruzka->SpeedButton1Click(Sender);
}
//---------------------------------------------------------------------------

//Просмотр логов
void __fastcall TMain::N21Click(TObject *Sender)
{
  Logs->ShowModal();      
}
//---------------------------------------------------------------------------

void __fastcall TMain::FormResize(TObject *Sender)
{
  ProgressBar->Left = Main->Width-ProgressBar->Width-13;
}
//---------------------------------------------------------------------------

//Загрузка оценки с формы ЕОП
void __fastcall TMain::N22Click(TObject *Sender)
{
  Zagruzka->RadioButtonEOP->Checked = true;
  Zagruzka->SpeedButton1Click(Sender);
}
//---------------------------------------------------------------------------

void __fastcall TMain::Cghfdrf1Click(TObject *Sender)
{
  try
    {
      WinExec(("\""+ WordPath+"\"\""+ Path+"\\Инструкция пользователя.doc\"").c_str(),SW_MAXIMIZE);
    }
  catch(...)
    {
      Application->MessageBox("Не найден файл со справкой","Предупреждение",
                              MB_OK+MB_ICONINFORMATION);
    }
}
//---------------------------------------------------------------------------

void __fastcall TMain::DBGridEh1Columns13GetCellParams(TObject *Sender,
      bool EditMode, TColCellParamsEh *Params)
{
  //Params->Text = IntToStr(Params->Row);
   Params->Text = IntToStr(DBGridEh1->DataRowToRecNo(Params->Row - 2));
}
//---------------------------------------------------------------------------

void __fastcall TMain::SpeedButton3Click(TObject *Sender)
{
  Zameshenie->ShowModal();
}
//---------------------------------------------------------------------------

void __fastcall TMain::N26Click(TObject *Sender)
{
  SpeedButton3Click(Sender);
}
//---------------------------------------------------------------------------

void __fastcall TMain::StatusBar1DblClick(TObject *Sender)
{
  Data->ShowModal();
}
//---------------------------------------------------------------------------

void __fastcall TMain::N30Click(TObject *Sender)
{
  Data->ShowModal();
}
//---------------------------------------------------------------------------

//Отчет по переведенным работникам
void __fastcall TMain::Gjgthtdtltyysvhfjnybrfv1Click(TObject *Sender)
{
  AnsiString Sql, sFile;
  int i,n;
  Variant AppEx,Sh;

  StatusBar1->SimpleText ="Идет формирование отчета по переведенным работникам...";

  Sql="select o.tn as tn_sap,                                                  \
              o.zex as old_zex,                                                \
              (select distinct nazv_cex from ssap_cex where o.zex=id_cex) as old_nzex,  \
              initcap(o.fio) as fio,                                           \
              o.dolg as old_dolg,                                              \
              k.zex as zex,                                                    \
              k.nzex as nzex,                                                  \
              k.name_dolg_ru as dolg,                                          \
              k.name_ur1 as nuch,                                              \
              o.uch as old_uch,                                                  \
              ur as uch,                                                        \
              dat_job                                                          \
       from (                                                                  \
             (select tn, zex, uch, dolg, fio from ocenka where god="+IntToStr(god)+") o     \
             left join                                                         \
             (select tn_sap,                                                   \
                     zex,                                                      \
                     nzex,                                                     \
                     name_dolg_ru,                                             \
                     name_ur1,                                                 \
                    (select dat_job from sap_perevod p where p.tn_sap=s.tn_sap and dat_job in (select max(dat_job) from sap_perevod m where m.tn_sap=p.tn_sap) ) as dat_job, \
                     case when ur1 is null then zex                            \
                          when ur2 is null then ur1                            \
                          when ur3 is null then ur2                            \
                          when ur4 is null then ur3 end ur                     \
              from sap_osn_sved s) k                                           \
              on o.tn=k.tn_sap)                                                \
        where nvl(o.zex,0)!=nvl(k.zex,0) and k.zex is not null    \
        order by o.zex, o.tn";
                                        //or upper(nvl(o.dolg,0))!=upper(nvl(name_dolg_ru,0))  or nvl(o.uch,0)!=nvl(k.ur,0))


  DM->qObnovlenie->Close();
  DM->qObnovlenie->SQL->Clear();
  DM->qObnovlenie->SQL->Add(Sql);
  try
    {
      DM->qObnovlenie->Open();
    }
  catch (Exception &E)
    {
      Application->MessageBox(("Возникла ошибка при выборке данных из картотеки по оценке персонала и кадров\n(OCENKA, SAP_OSN_SVED, SAP_PEREVOD)"+E.Message).c_str(),"Ошибка",
                              MB_OK+MB_ICONERROR);

      StatusBar1->SimpleText ="Отчетный период: "+IntToStr(god)+" год";
      Abort();
    }

  if (DM->qObnovlenie->RecordCount==0)
    {
      Application->MessageBox(("Нет данных по переведенным работникам за "+IntToStr(god)+" год").c_str(),"Предупреждение",
                              MB_OK+MB_ICONINFORMATION);

      StatusBar1->SimpleText ="Отчетный период: "+IntToStr(god)+" год";
      Abort();
    }

  Cursor = crHourGlass;
  ProgressBar->Position = 0;
  ProgressBar->Visible = true;
  ProgressBar->Max=DM->qObnovlenie->RecordCount;

  // инициализируем Excel, открываем этот шаблон
  try
    {
      AppEx=CreateOleObject("Excel.Application");
    }
  catch (...)
    {
      Application->MessageBox("Невозможно открыть Microsoft Excel!"
                              " Возможно это приложение на компьютере не установлено.","Ошибка",MB_OK+MB_ICONERROR);
      StatusBar1->SimpleText ="Отчетный период: "+IntToStr(god)+" год";
      ProgressBar->Visible = false;
      Cursor = crDefault;
    }

  //Если возникает ошибка во время формирования отчета
  try
    {
      try
        {
          AppEx.OlePropertySet("AskToUpdateLinks",false);
          AppEx.OlePropertySet("DisplayAlerts",false);

          //Создание папки, если ее не существует
          ForceDirectories(WorkPath);

          //Копируем шаблон файла в Мои документы
          CopyFile((Path+"\\RTF\\perevod.xlsx").c_str(), (WorkPath+"\\Отчет по переведенным работникам.xlsx").c_str(), false);
          sFile = WorkPath+"\\Отчет по переведенным работникам.xlsx";

          AppEx.OlePropertyGet("WorkBooks").OleProcedure("Open",sFile.c_str())    ;  //открываем книгу, указав её имя
          Sh=AppEx.OlePropertyGet("WorkSheets",1);                               //выбираем № активного листа книги
          //Sh=AppEx.OlePropertyGet("WorkSheets","Расчет");                      //выбираем лист по наименованию
        }
      catch(...)
        {
          Application->MessageBox("Ошибка открытия книги Microsoft Excel!","Ошибка",MB_OK+MB_ICONERROR);
          StatusBar1->SimpleText ="Отчетный период: "+IntToStr(god)+" год";
          ProgressBar->Visible = false;
          Cursor = crDefault;
        }


      i=1;
      n=3;

      //Вывод данных в шаблон
      Variant Massiv;
      Massiv = VarArrayCreate(OPENARRAY(int,(0,12)),varVariant); //массив на 11 элементов

      while (!DM->qObnovlenie->Eof)
        {
          Massiv.PutElement(i, 0);
          Massiv.PutElement(DM->qObnovlenie->FieldByName("fio")->AsString.c_str(), 1);
          Massiv.PutElement(DM->qObnovlenie->FieldByName("tn_sap")->AsString.c_str(), 2);
          Massiv.PutElement(DM->qObnovlenie->FieldByName("zex")->AsString.c_str(), 3);
          Massiv.PutElement(DM->qObnovlenie->FieldByName("nzex")->AsString.c_str(), 4);
          Massiv.PutElement(DM->qObnovlenie->FieldByName("nuch")->AsString.c_str(), 5);
          Massiv.PutElement(DM->qObnovlenie->FieldByName("dolg")->AsString.c_str(), 6);
          Massiv.PutElement(DM->qObnovlenie->FieldByName("dat_job")->AsString.c_str(), 7);
          Massiv.PutElement(DM->qObnovlenie->FieldByName("old_nzex")->AsString.c_str(), 8);
          Massiv.PutElement(DM->qObnovlenie->FieldByName("old_zex")->AsString.c_str(), 9);
          Massiv.PutElement(DM->qObnovlenie->FieldByName("old_dolg")->AsString.c_str(), 10);

          Sh.OlePropertyGet("Range", ("A" + IntToStr(n) + ":K" + IntToStr(n)).c_str()).OlePropertySet("Value", Massiv); //строка с данными с ячейки A по ячейку АВ

          i++;
          n++;
          DM->qObnovlenie->Next();
          ProgressBar->Position++;
        }

      //рисуем сетку
      Sh.OlePropertyGet("Range",("A3:K"+IntToStr(n-1)).c_str()).OlePropertyGet("Borders").OlePropertySet("LineStyle", xlContinuous);

     // Sh.OleProcedure("SaveAs",vAsCurDir1.c_str());
     AppEx.OlePropertyGet("WorkBooks",1).OleFunction("Save");

      /* //Закрыть открытое приложение Excel
      AppEx.OleProcedure("Quit");
      AppEx = Unassigned;  */

      //Закрыть книгу Excel с шаблоном для вывода информации
     // AppEx.OlePropertyGet("WorkBooks",1).OleProcedure("Close");
      Application->MessageBox("Отчет успешно сформирован!", "Формирование отчета",
                               MB_OK+MB_ICONINFORMATION);
      //AppEx.OlePropertyGet("WorkBooks").OleProcedure("Open",vAsCurDir1.c_str());
      AppEx.OlePropertySet("Visible",true);
      AppEx.OlePropertySet("AskToUpdateLinks",true);
      AppEx.OlePropertySet("DisplayAlerts",true);

      StatusBar1->SimpleText= "Формирование отчета выполнено.";

      Cursor = crDefault;
      ProgressBar->Position=0;
      ProgressBar->Visible = false;
      StatusBar1->SimpleText ="Отчетный период: "+IntToStr(god)+" год";
    }
  catch(...)
    {
      AppEx.OleProcedure("Quit");
      AppEx = Unassigned;
      Cursor = crDefault;
      ProgressBar->Position=0;
      ProgressBar->Visible = false;

      StatusBar1->SimpleText ="Отчетный период: "+IntToStr(god)+" год";
      Abort();
    }
}
//---------------------------------------------------------------------------

//Отчет по рекомендациям в кадровый резерв (расширенный)

//Отчет по рекомендациям в кадровый резерв (расширенный) по подразделению
void __fastcall TMain::N43Click(TObject *Sender)
{
  AnsiString Answer, Sql;

  if( InputQuery("Формирование отчета по рекомендациям в КР по цеху","Введите шифр подразделения", Answer) == True)
    {
      if(Answer=="" && (StrToInt(Answer)))
        {
          Application->MessageBox("Не введен шифр подразделения","Предупреждение",MB_OK +MB_ICONWARNING);
          Abort();
        }
      else
        {
          OtchetKR(Answer);
        }
   }
}
//---------------------------------------------------------------------------

//Отчет по рекомендациям в кадровый резерв (расширенный) по предприятию
void __fastcall TMain::N41Click(TObject *Sender)
{
  OtchetKR(0);
}
//---------------------------------------------------------------------------
//Отчет по рекомендациям в кадровый резерв (расширенный)
void __fastcall TMain::OtchetKR(AnsiString otchet)
{
  AnsiString Sql, sFile, tn, tn1;
  int i,n;
  Variant AppEx,Sh;


  Sql="select rukov.tn as r_tn,           \
              rukov.fio as r_fio,          \
              case when dolgn.kod_zex in ('47','54','99')  then  (select naim from sp_direkt where god="+IntToStr(god)+" and kod=(select kod_d from sp_pdirekt pdir where pdir.zex=dolgn.kod_szex and pdir.god="+IntToStr(god)+"))  \
                    else (select naim from sp_direkt where god="+IntToStr(god)+" and kod=(select kod_d from sp_pdirekt pdir1 where pdir1.zex=dolgn.kod_zex and pdir1.god="+IntToStr(god)+")) end  as r_direkt,           \
              id_dolg,            \
              id_shtat,           \
              dolgn.dolg as dolg,         \
              shifr_zex,          \
              kod_zex,            \
              kod_szex,           \
              objid_p1000,        \
              n_shtat,            \
              short,              \
              nzex,               \
              dolgn.uch as uch,                \
              zamesh.tn as z_tn,          \
              zamesh.fio as z_fio,         \
              zamesh.zex as z_zex,         \
              zamesh.dolg as z_dolg,        \
              (select naim from sp_direkt where god="+IntToStr(god)+" and kod=(select kod_d from sp_pdirekt pdir2 where pdir2.zex=zamesh.direkt and pdir2.god="+IntToStr(god)+")) as z_direkt, \
              (select nazv_cexk from ssap_cex where id_cex=substr(zamesh.zex,1,2) and nazv_cexk not like '%(устар.)') as  z_nzex, \
              decode(length(zamesh.uch),2,NULL,(select stext from p1000@sapmig_buffdb where otype='O' and langu='R' and stext not like '%(уст%' and short=zamesh.uch)) as z_nuch \
       from (                             \
                      (select stext as dolg,       \
                              shifr_zex,           \
                              kod_zex,             \
                              kod_szex,            \
                              objid_p1000,         \
                              n_shtat,             \
                              short,               \
                              (select stext from p1000@sapmig_buffdb where otype='O' and langu='R' and stext not like '%(устар.)' and short=kod_zex) as nzex, \
                              uch                  \
                                                   \
                       from (                      \                                                                                                                          \
                             (select p1.otype, p1.objid as zvezda1,                                                                                                           \
                                     p1.begda, p1.endda, p1.sobid as sobid_p1001,                                                                                             \
                                     p2.stext, kod as zvezda3, p2.objid as n_shtat, p2.short                                                                                  \
                              from (select r.otype, r.objid, r.begda, r.endda, s.sobid as sobid, s.objid as kod                                                               \
                                    from p1013@sapmig_buffdb r left join p1001@sapmig_buffdb s on r.objid=s.objid and s.otype='S' and s.sclas='O' where r.otype='S' and       \
                                    (r.persk=10 or s.objid in (select objid from p1000@sapmig_buffdb where otype='S' and langu='R' and trim(stext) in ('Механик цеха',        \
                                    'Энергетик цеха','Электрик цеха','Механик участка','Энергетик участка','Электрик участка','Механик фабрики','Энергетик фабрики',          \
                                    'Электрик фабрики','Механик управления','Энергетик управления','Электрик управления','Сменный механик участка','Сменный энергетик участка', \
                                    'Сменный электрик участка')) )) p1,                                                                                                          \
                                    p1000@sapmig_buffdb p2                                                                                                                       \
                              where p2.otype='S' and p2.langu='R' and p1.objid=p2.objid  and upper(stext) not like '%МЕНЕДЖЕР%'                                                   \
                              ) obsh1                                                                            \
                             left join                                                                           \
                             (select objid as objid_p1000, short as shifr_zex, substr(short,1,2) as kod_zex,     \
                                     substr(short,1,5) as kod_szex, stext as uch                                 \
                              from p1000@sapmig_buffdb where otype='O' and langu='R'                             \
                              )obsh2                                                                             \
                             on  sobid_p1001=objid_p1000                                                         \
                             )                                                                                   \
                       where endda>sysdate   and substr(shifr_zex,1,2) not in ('10','11','49','50')                                                       \
                      ) dolgn                                                                                    \
                     left join                                                                                   \
                      (select case when sv.ur1 is null then sv.zex    \
                                   when sv.ur2 is null then sv.ur1    \
                                   when sv.ur3 is null then sv.ur2    \
                                   when sv.ur4 is null then sv.ur3 end as z, tn, fio, direkt, id_dolg, id_shtat                                                 \
                       from                                                                                      \
                           (select tn, initcap(fio) as fio, direkt from ocenka where tn not in (select tn_sap from sap_decr) and god="+IntToStr(god)+") o                                \
                           left join                                                                             \
                           (select tn_sap, zex, ur1,ur2,ur3,ur4, id_dolg, id_shtat from sap_osn_sved) sv             \
                           on tn=sv.tn_sap )  rukov                                                                   \
                     on id_shtat=dolgn.n_shtat and id_dolg=dolgn.short                                                  \
                     left join                                                                                          \
                      (select oc.tn, initcap(fio) as fio, zex, dolg, direkt, uch,                                        \
                              orez.id_shtat as shtat_zam,                                                                \
                              orez.tn_sap_rez as tn_sap_zam, orez.zex_rez as zex_rez,                                    \
                              orez.shifr_rez as shifr_rez                                                                \
                      from ocenka oc left join ocenka_rez orez on oc.tn=orez.tn where oc.god="+IntToStr(god)+" and nvl(oc.zam,0)=0 and orez.god="+IntToStr(god)+") zamesh \
                      on dolgn.n_shtat=zamesh.shtat_zam                                             \
                   )";                                                                                             \

    if (otchet==0)
    {
      Sql+=" order by shifr_zex, n_shtat, short, z_tn";
      StatusBar1->SimpleText=" Идет формирование отчета по рекомендациям в КР на руководящие должности по предприятию в Excel...";
    }
  else
    {
      Sql+=" where shifr_zex like '"+otchet+"%' order by shifr_zex, n_shtat, short, z_tn";
      StatusBar1->SimpleText=" Идет формирование отчета по рекомендациям в КР на руководящие должности по подразделению в Excel...";                                                                                                                                                \
    }

  DM->qObnovlenie->Close();
  DM->qObnovlenie->SQL->Clear();
  DM->qObnovlenie->SQL->Add(Sql);
  try
    {
      DM->qObnovlenie->Open();
    }
  catch (Exception &E)
    {
      Application->MessageBox(("Возникла ошибка при выборке данных из картотеки по оценке персонала и кадров\n(OCENKA, SAP_OSN_SVED, SAP_PEREVOD)"+E.Message).c_str(),"Ошибка",
                              MB_OK+MB_ICONERROR);

      StatusBar1->SimpleText ="Отчетный период: "+IntToStr(god)+" год";
      Abort();
    }

  if (DM->qObnovlenie->RecordCount==0)
    {
      Application->MessageBox(("Нет данных для формирования отчета за "+IntToStr(god)+" год").c_str(),"Предупреждение",
                              MB_OK+MB_ICONINFORMATION);

      StatusBar1->SimpleText ="Отчетный период: "+IntToStr(god)+" год";
      Abort();
    }

  Cursor = crHourGlass;
  ProgressBar->Position = 0;
  ProgressBar->Visible = true;
  ProgressBar->Max=DM->qObnovlenie->RecordCount;

  // инициализируем Excel, открываем этот шаблон
  try
    {
      AppEx=CreateOleObject("Excel.Application");
    }
  catch (...)
    {
      Application->MessageBox("Невозможно открыть Microsoft Excel!"
                              " Возможно это приложение на компьютере не установлено.","Ошибка",MB_OK+MB_ICONERROR);
      StatusBar1->SimpleText ="Отчетный период: "+IntToStr(god)+" год";
      ProgressBar->Visible = false;
      Cursor = crDefault;
    }

  //Если возникает ошибка во время формирования отчета
  try
    {
      try
        {
          AppEx.OlePropertySet("AskToUpdateLinks",false);
          AppEx.OlePropertySet("DisplayAlerts",false);

          //Создание папки, если ее не существует
          ForceDirectories(WorkPath);

          //Создание папки, если ее не существует
          ForceDirectories(WorkPath+"\\Формирование списка КР(расширенный)");

          if (otchet==0)
            {
              //Копируем шаблон файла в Мои документы
              CopyFile((Path+"\\RTF\\kr.xlsx").c_str(), (WorkPath+"\\Формирование списка КР(расширенный)\\По предприятию.xlsx").c_str(), false);
              sFile = WorkPath+"\\Формирование списка КР(расширенный)\\По предприятию.xlsx";
            }
          else
            {
              //Копируем шаблон файла в Мои документы
              CopyFile((Path+"\\RTF\\kr.xlsx").c_str(), (WorkPath+"\\Формирование списка КР(расширенный)\\"+otchet+".xlsx").c_str(), false);
              sFile = WorkPath+"\\Формирование списка КР(расширенный)\\"+otchet+".xlsx";
            }

          AppEx.OlePropertyGet("WorkBooks").OleProcedure("Open",sFile.c_str())    ;  //открываем книгу, указав её имя
          Sh=AppEx.OlePropertyGet("WorkSheets",1);                               //выбираем № активного листа книги
          //Sh=AppEx.OlePropertyGet("WorkSheets","Расчет");                      //выбираем лист по наименованию

        }
      catch(...)
        {
          Application->MessageBox("Ошибка открытия книги Microsoft Excel!","Ошибка",MB_OK+MB_ICONERROR);
          StatusBar1->SimpleText ="Отчетный период: "+IntToStr(god)+" год";
          ProgressBar->Visible = false;
          Cursor = crDefault;
        }
     // AppEx.OlePropertySet("Visible",true);

      i=1;
      n=4;
      int num=0;

      //Вывод данных в шаблон
      Variant Massiv;
      Massiv = VarArrayCreate(OPENARRAY(int,(0,17)),varVariant); //массив на 17 элементов


      while (!DM->qObnovlenie->Eof)
        {
          tn=DM->qObnovlenie->FieldByName("n_shtat")->AsString;
          tn1=DM->qObnovlenie->FieldByName("n_shtat")->AsString;
          num=1;

          while (!DM->qObnovlenie->Eof && tn==tn1)
            {
              if (num!=1 && !tn.IsEmpty())
                {
                  //Объединение ячеек
                  Sh.OlePropertyGet("Range",("A"+IntToStr(n)+":A"+IntToStr(n-1)).c_str()).OleProcedure("Merge");
                  Sh.OlePropertyGet("Range",("B"+IntToStr(n)+":B"+IntToStr(n-1)).c_str()).OleProcedure("Merge");
                  Sh.OlePropertyGet("Range",("C"+IntToStr(n)+":C"+IntToStr(n-1)).c_str()).OleProcedure("Merge");
                  Sh.OlePropertyGet("Range",("D"+IntToStr(n)+":D"+IntToStr(n-1)).c_str()).OleProcedure("Merge");
                  Sh.OlePropertyGet("Range",("E"+IntToStr(n)+":E"+IntToStr(n-1)).c_str()).OleProcedure("Merge");
                  Sh.OlePropertyGet("Range",("F"+IntToStr(n)+":F"+IntToStr(n-1)).c_str()).OleProcedure("Merge");
                  Sh.OlePropertyGet("Range",("G"+IntToStr(n)+":G"+IntToStr(n-1)).c_str()).OleProcedure("Merge");
                  Sh.OlePropertyGet("Range",("H"+IntToStr(n)+":H"+IntToStr(n-1)).c_str()).OleProcedure("Merge");
                  Sh.OlePropertyGet("Range",("I"+IntToStr(n)+":I"+IntToStr(n-1)).c_str()).OleProcedure("Merge");
                  Sh.OlePropertyGet("Range",("J"+IntToStr(n)+":I"+IntToStr(n-1)).c_str()).OleProcedure("Merge");
                }

              Massiv.PutElement(i, 0);
              Massiv.PutElement(DM->qObnovlenie->FieldByName("n_shtat")->AsString.c_str(), 1);
              Massiv.PutElement(DM->qObnovlenie->FieldByName("short")->AsString.c_str(), 2);
              Massiv.PutElement(DM->qObnovlenie->FieldByName("dolg")->AsString.c_str(), 3);
              Massiv.PutElement(DM->qObnovlenie->FieldByName("r_fio")->AsString.c_str(), 4);
              Massiv.PutElement(DM->qObnovlenie->FieldByName("r_tn")->AsString.c_str(), 5);
              Massiv.PutElement(DM->qObnovlenie->FieldByName("r_direkt")->AsString.c_str(), 6);
              Massiv.PutElement(DM->qObnovlenie->FieldByName("shifr_zex")->AsString.c_str(), 7);
              Massiv.PutElement(DM->qObnovlenie->FieldByName("nzex")->AsString.c_str(), 8);
              Massiv.PutElement(DM->qObnovlenie->FieldByName("uch")->AsString.c_str(), 9);
              if (DM->qObnovlenie->FieldByName("z_tn")->AsString.IsEmpty()) Massiv.PutElement("", 10);
              else Massiv.PutElement(num, 10);
              Massiv.PutElement(DM->qObnovlenie->FieldByName("z_fio")->AsString.c_str(), 11);
              Massiv.PutElement(DM->qObnovlenie->FieldByName("z_tn")->AsString.c_str(), 12);
              Massiv.PutElement(DM->qObnovlenie->FieldByName("z_dolg")->AsString.c_str(), 13);
              Massiv.PutElement(DM->qObnovlenie->FieldByName("z_direkt")->AsString.c_str(), 14);
              Massiv.PutElement(DM->qObnovlenie->FieldByName("z_zex")->AsString.c_str(), 15);
              Massiv.PutElement(DM->qObnovlenie->FieldByName("z_nzex")->AsString.c_str(), 16);
              Massiv.PutElement(DM->qObnovlenie->FieldByName("z_nuch")->AsString.c_str(), 17);

              Sh.OlePropertyGet("Range", ("A" + IntToStr(n) + ":R" + IntToStr(n)).c_str()).OlePropertySet("Value", Massiv); //строка с данными с ячейки A по ячейку АВ

              i++;
              n++;
              num++;
              DM->qObnovlenie->Next();

              tn=DM->qObnovlenie->FieldByName("n_shtat")->AsString;
              StatusBar1->SimpleText ="Идет формирование отчета по рекомендациям в КР на руководящие должности по предприятию... "+DM->qObnovlenie->FieldByName("shifr_zex")->AsString;
              ProgressBar->Position++;
            }
        }

      //рисуем сетку
      Sh.OlePropertyGet("Range",("A4:R"+IntToStr(n-1)).c_str()).OlePropertyGet("Borders").OlePropertySet("LineStyle", xlContinuous);

     // Sh.OleProcedure("SaveAs",vAsCurDir1.c_str());
     AppEx.OlePropertyGet("WorkBooks",1).OleFunction("Save");

      /* //Закрыть открытое приложение Excel
      AppEx.OleProcedure("Quit");
      AppEx = Unassigned;  */

      //Закрыть книгу Excel с шаблоном для вывода информации
     // AppEx.OlePropertyGet("WorkBooks",1).OleProcedure("Close");
      Application->MessageBox("Отчет успешно сформирован!", "Формирование отчета",
                               MB_OK+MB_ICONINFORMATION);
      //AppEx.OlePropertyGet("WorkBooks").OleProcedure("Open",vAsCurDir1.c_str());
      AppEx.OlePropertySet("Visible",true);
      AppEx.OlePropertySet("AskToUpdateLinks",true);
      AppEx.OlePropertySet("DisplayAlerts",true);

      StatusBar1->SimpleText= "Формирование отчета выполнено.";

      Cursor = crDefault;
      ProgressBar->Position=0;
      ProgressBar->Visible = false;
      StatusBar1->SimpleText ="Отчетный период: "+IntToStr(god)+" год";

       if (otchet==0) InsertLog("Формирование КР работников (расширенный) за "+IntToStr(god)+" год по предприятию в Excel успешно завершено");
       else InsertLog("Формирование КР работников (расширенный) за "+IntToStr(god)+" год по подразделению "+otchet+" в Excel успешно завершено");
       DM->qLogs->Requery();
    }
  catch(...)
    {
      AppEx.OleProcedure("Quit");
      AppEx = Unassigned;
      Cursor = crDefault;
      ProgressBar->Position=0;
      ProgressBar->Visible = false;
      if (otchet==0) InsertLog("Не выполнено формирование КР работников (расширенный) за "+IntToStr(god)+" год по предприятию в Excel");
      else InsertLog("Не выполнено формирование КР работников (расширенный) за "+IntToStr(god)+" год по подразделению "+otchet+" в Excel");

      StatusBar1->SimpleText ="Отчетный период: "+IntToStr(god)+" год";
      Abort();
    }
}
//---------------------------------------------------------------------------

//Отчет по рекомендациям в кадровый резерв (расширенный) по дирекциям
void __fastcall TMain::N42Click(TObject *Sender)
{
  AnsiString Sql, sFile, tn, tn1, zex, zex1;
  int i,n;
  Variant AppEx,Sh;

  StatusBar1->SimpleText ="Идет формирование отчета по рекомендациям в КР на руководящие должности по дирекциям...";

  Sql="select rukov.tn as r_tn,            \
              rukov.fio as r_fio,          \
              nvl(rukov.direkt,0) as direkt,      \
              case when dolgn.kod_zex in ('47','54','99')  then  (select naim from sp_direkt where god="+IntToStr(god)+" and kod=(select kod_d from sp_pdirekt pdir where pdir.zex=dolgn.kod_szex and pdir.god="+IntToStr(god)+"))  \
                    else (select naim from sp_direkt where god="+IntToStr(god)+" and kod=(select kod_d from sp_pdirekt pdir1 where pdir1.zex=dolgn.kod_zex and pdir1.god="+IntToStr(god)+")) end  as r_direkt,           \
              id_dolg,            \
              id_shtat,           \
              dolgn.dolg as dolg,         \
              shifr_zex,          \
              kod_zex,            \
              kod_szex,           \
              objid_p1000,        \
              n_shtat,            \
              short,              \
              nzex,               \
              dolgn.uch as uch,                \
              zamesh.tn as z_tn,          \
              zamesh.fio as z_fio,         \
              zamesh.zex as z_zex,         \
              zamesh.dolg as z_dolg,        \
              (select naim from sp_direkt where god="+IntToStr(god)+" and kod=(select kod_d from sp_pdirekt pdir2 where pdir2.zex=zamesh.direkt and pdir2.god="+IntToStr(god)+")) as z_direkt, \
              (select nazv_cexk from ssap_cex where id_cex=substr(zamesh.zex,1,2) and nazv_cexk not like '%(устар.)') as  z_nzex, \
              decode(length(zamesh.uch),2,NULL,(select stext from p1000@sapmig_buffdb where otype='O' and langu='R' and stext not like '%(уст%' and short=zamesh.uch)) as z_nuch \
       from (                             \
                      (select stext as dolg,       \
                              shifr_zex,           \
                              kod_zex,             \
                              kod_szex,            \
                              objid_p1000,         \
                              n_shtat,             \
                              short,               \
                              (select stext from p1000@sapmig_buffdb where otype='O' and langu='R' and stext not like '%(устар.)' and short=kod_zex) as nzex, \
                              uch                  \
                                                   \
                       from (                      \
                             (select p1.otype, p1.objid as zvezda1,                                                                                                           \
                                     p1.begda, p1.endda, p1.sobid as sobid_p1001,                                                                                             \
                                     p2.stext, kod as zvezda3, p2.objid as n_shtat, p2.short                                                                                  \
                              from (select r.otype, r.objid, r.begda, r.endda, s.sobid as sobid, s.objid as kod                                                               \
                                    from p1013@sapmig_buffdb r left join p1001@sapmig_buffdb s on r.objid=s.objid and s.otype='S' and s.sclas='O' where r.otype='S' and       \
                                    (r.persk=10 or s.objid in (select objid from p1000@sapmig_buffdb where otype='S' and langu='R' and trim(stext) in ('Механик цеха',        \
                                    'Энергетик цеха','Электрик цеха','Механик участка','Энергетик участка','Электрик участка','Механик фабрики','Энергетик фабрики',          \
                                    'Электрик фабрики','Механик управления','Энергетик управления','Электрик управления','Сменный механик участка','Сменный энергетик участка', \
                                    'Сменный электрик участка')) )) p1,                                                                                                          \
                                    p1000@sapmig_buffdb p2                                                                                                                       \
                              where p2.otype='S' and p2.langu='R' and p1.objid=p2.objid  and upper(stext) not like '%МЕНЕДЖЕР%'                                                   \
                              ) obsh1                                                                            \
                             left join                                                                           \
                             (select objid as objid_p1000, short as shifr_zex, substr(short,1,2) as kod_zex,     \
                                     substr(short,1,5) as kod_szex, stext as uch                                 \
                              from p1000@sapmig_buffdb where otype='O' and langu='R'                             \
                              )obsh2                                                                             \
                             on  sobid_p1001=objid_p1000                                                         \
                             )                                                                                   \
                       where endda>sysdate  and substr(shifr_zex,1,2) not in ('10','11','49','50')                                                        \
                      ) dolgn                                                                                    \
                     left join                                                                                   \
                      (select case when sv.ur1 is null then sv.zex    \
                                   when sv.ur2 is null then sv.ur1    \
                                   when sv.ur3 is null then sv.ur2    \
                                   when sv.ur4 is null then sv.ur3 end as z, tn, fio, direkt, id_dolg, id_shtat                                                 \
                       from                                                                                      \
                           (select tn, initcap(fio) as fio, direkt from ocenka where tn not in (select tn_sap from sap_decr) and god="+IntToStr(god)+") o                                \
                           left join                                                                             \
                           (select tn_sap, zex, ur1,ur2,ur3,ur4, id_dolg, id_shtat from sap_osn_sved) sv             \
                           on tn=sv.tn_sap )  rukov                                                                   \
                     on id_shtat=dolgn.n_shtat and id_dolg=dolgn.short                                                  \
                     left join                                                                                          \
                      (select oc.tn, initcap(fio) as fio, zex, dolg, direkt, uch,                                        \
                              orez.id_shtat as shtat_zam,                                                                \
                              orez.tn_sap_rez as tn_sap_zam, orez.zex_rez as zex_rez,                                    \
                              orez.shifr_rez as shifr_rez                                                                \
                      from ocenka oc left join ocenka_rez orez on oc.tn=orez.tn where oc.god="+IntToStr(god)+" and nvl(oc.zam,0)=0 and orez.god="+IntToStr(god)+") zamesh \
                      on dolgn.n_shtat=zamesh.shtat_zam                                           \
                   )                                                                    \
         order by nvl(rukov.direkt,0), shifr_zex, n_shtat, short, z_tn";
                                              //   rukov.tn=tn_sap_zam
  DM->qObnovlenie->Close();
  DM->qObnovlenie->SQL->Clear();
  DM->qObnovlenie->SQL->Add(Sql);
  try
    {
      DM->qObnovlenie->Open();
    }
  catch (Exception &E)
    {
      Application->MessageBox(("Возникла ошибка при выборке данных из картотеки по оценке персонала и кадров\n(OCENKA, SAP_OSN_SVED, SAP_PEREVOD)"+E.Message).c_str(),"Ошибка",
                              MB_OK+MB_ICONERROR);

      StatusBar1->SimpleText ="Отчетный период: "+IntToStr(god)+" год";
      Abort();
    }

  if (DM->qObnovlenie->RecordCount==0)
    {
      Application->MessageBox(("Нет данных по переведенным работникам за "+IntToStr(god)+" год").c_str(),"Предупреждение",
                              MB_OK+MB_ICONINFORMATION);

      StatusBar1->SimpleText ="Отчетный период: "+IntToStr(god)+" год";
      Abort();
    }

  Cursor = crHourGlass;
  ProgressBar->Position = 0;
  ProgressBar->Visible = true;
  ProgressBar->Max=DM->qObnovlenie->RecordCount;

  // инициализируем Excel, открываем этот шаблон
  try
    {
      AppEx=CreateOleObject("Excel.Application");
    }
  catch (...)
    {
      Application->MessageBox("Невозможно открыть Microsoft Excel!"
                              " Возможно это приложение на компьютере не установлено.","Ошибка",MB_OK+MB_ICONERROR);
      StatusBar1->SimpleText ="Отчетный период: "+IntToStr(god)+" год";
      ProgressBar->Visible = false;
      Cursor = crDefault;
    }

  while (DM->qObnovlenie->RecordCount>0 && !DM->qObnovlenie->Eof)
    {
      zex = DM->qObnovlenie->FieldByName("direkt")->AsString;

      //Если возникает ошибка во время формирования отчета
      try
        {
          try
            {
              AppEx.OlePropertySet("AskToUpdateLinks",false);
              AppEx.OlePropertySet("DisplayAlerts",false);

              //Создание папки, если ее не существует
              ForceDirectories(WorkPath+"\\Формирование списка КР(расширенный) по дирекциям");

              //Копируем шаблон файла в Мои документы
              CopyFile((Path+"\\RTF\\kr.xlsx").c_str(), (WorkPath+"\\Формирование списка КР(расширенный) по дирекциям\\"+zex+".xlsx").c_str(), false);
              sFile = WorkPath+"\\Формирование списка КР(расширенный) по дирекциям\\"+zex+".xlsx";

              AppEx.OlePropertyGet("WorkBooks").OleProcedure("Open",sFile.c_str())    ;  //открываем книгу, указав её имя
              Sh=AppEx.OlePropertyGet("WorkSheets",1);                               //выбираем № активного листа книги
              //Sh=AppEx.OlePropertyGet("WorkSheets","Расчет");                      //выбираем лист по наименованию
            }
          catch(...)
            {
              Application->MessageBox("Ошибка открытия книги Microsoft Excel!","Ошибка",MB_OK+MB_ICONERROR);
              StatusBar1->SimpleText ="Отчетный период: "+IntToStr(god)+" год";
              ProgressBar->Visible = false;
              Cursor = crDefault;
            }

          //AppEx.OlePropertySet("Visible",true);


          i=1;
          n=4;
          int num=0;

          //Вывод данных в шаблон
          Variant Massiv;
          Massiv = VarArrayCreate(OPENARRAY(int,(0,17)),varVariant); //массив на 17 элементов

          zex1 = DM->qObnovlenie->FieldByName("direkt")->AsString;
          zex = DM->qObnovlenie->FieldByName("direkt")->AsString;

          StatusBar1->SimpleText ="Идет формирование отчета по рекомендациям в КР на руководящие должности по дирекциям...  "+zex;

          while (!DM->qObnovlenie->Eof && zex==zex1)
            {
              tn=DM->qObnovlenie->FieldByName("n_shtat")->AsString;
              tn1=DM->qObnovlenie->FieldByName("n_shtat")->AsString;
              num=1;

              while (!DM->qObnovlenie->Eof && tn==tn1)
                {
                  if (num!=1 && !tn.IsEmpty())
                    {
                      //Объединение ячеек
                      Sh.OlePropertyGet("Range",("A"+IntToStr(n)+":A"+IntToStr(n-1)).c_str()).OleProcedure("Merge");
                      Sh.OlePropertyGet("Range",("B"+IntToStr(n)+":B"+IntToStr(n-1)).c_str()).OleProcedure("Merge");
                      Sh.OlePropertyGet("Range",("C"+IntToStr(n)+":C"+IntToStr(n-1)).c_str()).OleProcedure("Merge");
                      Sh.OlePropertyGet("Range",("D"+IntToStr(n)+":D"+IntToStr(n-1)).c_str()).OleProcedure("Merge");
                      Sh.OlePropertyGet("Range",("E"+IntToStr(n)+":E"+IntToStr(n-1)).c_str()).OleProcedure("Merge");
                      Sh.OlePropertyGet("Range",("F"+IntToStr(n)+":F"+IntToStr(n-1)).c_str()).OleProcedure("Merge");
                      Sh.OlePropertyGet("Range",("G"+IntToStr(n)+":G"+IntToStr(n-1)).c_str()).OleProcedure("Merge");
                      Sh.OlePropertyGet("Range",("H"+IntToStr(n)+":H"+IntToStr(n-1)).c_str()).OleProcedure("Merge");
                      Sh.OlePropertyGet("Range",("I"+IntToStr(n)+":I"+IntToStr(n-1)).c_str()).OleProcedure("Merge");
                      Sh.OlePropertyGet("Range",("J"+IntToStr(n)+":I"+IntToStr(n-1)).c_str()).OleProcedure("Merge");
                    }

                  Massiv.PutElement(i, 0);
                  Massiv.PutElement(DM->qObnovlenie->FieldByName("n_shtat")->AsString.c_str(), 1);
                  Massiv.PutElement(DM->qObnovlenie->FieldByName("short")->AsString.c_str(), 2);
                  Massiv.PutElement(DM->qObnovlenie->FieldByName("dolg")->AsString.c_str(), 3);
                  Massiv.PutElement(DM->qObnovlenie->FieldByName("r_fio")->AsString.c_str(), 4);
                  Massiv.PutElement(DM->qObnovlenie->FieldByName("r_tn")->AsString.c_str(), 5);
                  Massiv.PutElement(DM->qObnovlenie->FieldByName("r_direkt")->AsString.c_str(), 6);
                  Massiv.PutElement(DM->qObnovlenie->FieldByName("shifr_zex")->AsString.c_str(), 7);
                  Massiv.PutElement(DM->qObnovlenie->FieldByName("nzex")->AsString.c_str(), 8);
                  Massiv.PutElement(DM->qObnovlenie->FieldByName("uch")->AsString.c_str(), 9);
                  if (DM->qObnovlenie->FieldByName("z_tn")->AsString.IsEmpty()) Massiv.PutElement("", 10);
                  else Massiv.PutElement(num, 10);
                  Massiv.PutElement(DM->qObnovlenie->FieldByName("z_fio")->AsString.c_str(), 11);
                  Massiv.PutElement(DM->qObnovlenie->FieldByName("z_tn")->AsString.c_str(), 12);
                  Massiv.PutElement(DM->qObnovlenie->FieldByName("z_dolg")->AsString.c_str(), 13);
                  Massiv.PutElement(DM->qObnovlenie->FieldByName("z_direkt")->AsString.c_str(), 14);
                  Massiv.PutElement(DM->qObnovlenie->FieldByName("z_zex")->AsString.c_str(), 15);
                  Massiv.PutElement(DM->qObnovlenie->FieldByName("z_nzex")->AsString.c_str(), 16);
                  Massiv.PutElement(DM->qObnovlenie->FieldByName("z_nuch")->AsString.c_str(), 17);

                  Sh.OlePropertyGet("Range", ("A" + IntToStr(n) + ":R" + IntToStr(n)).c_str()).OlePropertySet("Value", Massiv); //строка с данными с ячейки A по ячейку АВ

                  i++;
                  n++;
                  num++;
                  DM->qObnovlenie->Next();

                  zex1 = DM->qObnovlenie->FieldByName("direkt")->AsString;
                  tn1=DM->qObnovlenie->FieldByName("n_shtat")->AsString;
                  ProgressBar->Position++;
                }
            }

          //рисуем сетку
          Sh.OlePropertyGet("Range",("A4:R"+IntToStr(n-1)).c_str()).OlePropertyGet("Borders").OlePropertySet("LineStyle", xlContinuous);

          //Сохранить книгу в папке в файле по указанию
          AppEx.OlePropertyGet("WorkBooks",1).OleFunction("Save");

          //Закрыть книгу Excel с шаблоном для вывода информации
           AppEx.OlePropertyGet("WorkBooks",1).OleProcedure("Close");

          //AppEx.OlePropertySet("Visible",true);
       }
     catch(...)
       {
         AppEx.OleProcedure("Quit");
         AppEx = Unassigned;
         Cursor = crDefault;
         ProgressBar->Position=0;
         ProgressBar->Visible = false;

         StatusBar1->SimpleText ="Отчетный период: "+IntToStr(god)+" год";
         Abort();
       }
    }

  //Закрыть открытое приложение Excel
  AppEx.OleProcedure("Quit");
  AppEx.OlePropertySet("AskToUpdateLinks",true);
  AppEx.OlePropertySet("DisplayAlerts",true);
  AppEx = Unassigned;

  Cursor = crDefault;
  ProgressBar->Position=0;
  ProgressBar->Visible = false;
  
  StatusBar1->SimpleText= "Формирование отчета выполнено.";
  StatusBar1->SimpleText ="Отчетный период: "+IntToStr(god)+" год";

  InsertLog("Формирование КР работников за "+IntToStr(god)+" год по дирекциям в Excel успешно завершено");
  DM->qLogs->Requery();
  Application->MessageBox("Формирование файлов успешно завершено","Формирование файлов",
                           MB_OK+MB_ICONINFORMATION);

}
//---------------------------------------------------------------------------

//Загрузка КПЭ
void __fastcall TMain::N34Click(TObject *Sender)
{
  Zagruzka->RadioButtonKPE->Checked = true;
  Zagruzka->SpeedButton1Click(Sender);
}
//---------------------------------------------------------------------------

//Загрузка возраста ухода на пенсию
void __fastcall TMain::N35Click(TObject *Sender)
{
  Zagruzka->RadioButtonVZ->Checked = true;
  Zagruzka->SpeedButton1Click(Sender);
}
//---------------------------------------------------------------------------

//Загрузка рекомендаций КР из отчета
void __fastcall TMain::N36Click(TObject *Sender)
{
  Zagruzka->RadioButtonKR->Checked = true;
  Zagruzka->SpeedButton1Click(Sender);
}
//---------------------------------------------------------------------------

//Отчет по рекомендациям в кадровый резерв (сокращенный) по предприятию
void __fastcall TMain::N23Click(TObject *Sender)
{
  OtchetKRSokr(0);
}
//---------------------------------------------------------------------------
//Отчет по рекомендациям в кадровый резерв (сокращенный) по дирекциям
void __fastcall TMain::N24Click(TObject *Sender)
{
  AnsiString Sql, sFile, zex, zex1;
  int i,n;
  Variant AppEx,Sh;

  StatusBar1->SimpleText ="Идет формирование отчета по рекомендациям в КР на руководящие должности предприятия...";

  Sql="select  distinct rukov.tn as r_tn,                                                                           \
               short,                       \
               n_shtat,\                                                                                             \
               dolgn.dolg as dolg,     \                                                                             \
               rukov.fio as r_fio,                                                                                  \
               nvl(rukov.direkt,0) as direkt,                                                                                        \
               case when dolgn.kod_zex in ('47','54','99')  then  (select naim from sp_direkt god="+IntToStr(god)+" and where kod=(select kod_d from sp_pdirekt pdir where pdir.zex=dolgn.kod_szex and pdir.god="+IntToStr(god)+"))  \
                    else (select naim from sp_direkt where god="+IntToStr(god)+" and kod=(select kod_d from sp_pdirekt pdir1 where pdir1.zex=dolgn.kod_zex and pdir1.god="+IntToStr(god)+")) end  as r_direkt,           \                                         \
               kod_zex as zex,                                                                                      \
               shifr_zex,                                                                                           \
               nzex,                                                                                                \
               uch as n_uch,                                                                                                 \
               case when rukov.tn is not null and zamesh.tn is not null then count(*) over (partition by rukov.tn)  \
               else 0 end as kol                                                                                    \
       from (                                                                                                       \
                      (select stext as dolg,                                                                        \
                              shifr_zex,                                                                            \
                              kod_zex,                                                                              \
                              kod_szex,                                                                             \
                              objid_p1000,                                                                          \
                              n_shtat,                                                                              \
                              short,                                                                                \
                              (select stext from p1000@sapmig_buffdb where otype='O' and langu='R' and stext not like '%(устар.)' and short=kod_zex) as nzex,  \
                              uch                                                                                   \
                                                                                                                    \
                       from (                                                                                       \
                             (select p1.otype, p1.objid as zvezda1,                                                                                                           \
                                     p1.begda, p1.endda, p1.sobid as sobid_p1001,                                                                                             \
                                     p2.stext, kod as zvezda3, p2.objid as n_shtat, p2.short                                                                                  \
                              from (select r.otype, r.objid, r.begda, r.endda, s.sobid as sobid, s.objid as kod                                                               \
                                    from p1013@sapmig_buffdb r left join p1001@sapmig_buffdb s on r.objid=s.objid and s.otype='S' and s.sclas='O' where r.otype='S' and       \
                                    (r.persk=10 or s.objid in (select objid from p1000@sapmig_buffdb where otype='S' and langu='R' and trim(stext) in ('Механик цеха',        \
                                    'Энергетик цеха','Электрик цеха','Механик участка','Энергетик участка','Электрик участка','Механик фабрики','Энергетик фабрики',          \
                                    'Электрик фабрики','Механик управления','Энергетик управления','Электрик управления','Сменный механик участка','Сменный энергетик участка', \
                                    'Сменный электрик участка')) )) p1,                                                                                                          \
                                    p1000@sapmig_buffdb p2                                                                                                                       \
                              where p2.otype='S' and p2.langu='R' and p1.objid=p2.objid  and upper(stext) not like '%МЕНЕДЖЕР%'                                                   \
                              ) obsh1                                                                               \
                             left join                                                                              \
                             (select objid as objid_p1000, short as shifr_zex, substr(short,1,2) as kod_zex,        \
                                     substr(short,1,5) as kod_szex, stext as uch                                    \
                              from p1000@sapmig_buffdb where otype='O' and langu='R'                                \
                              )obsh2                                                                                \
                             on  sobid_p1001=objid_p1000                                                            \
                             )                                                                                      \
                       where endda>sysdate and substr(shifr_zex,1,2) not in ('10','11','49','50')                                                        \
                      ) dolgn                                                                                       \
                     left join                                                                                      \
                      (select case when sv.ur1 is null then sv.zex     \
                                   when sv.ur2 is null then sv.ur1     \
                                   when sv.ur3 is null then sv.ur2     \
                                   when sv.ur4 is null then sv.ur3 end z, tn, fio, direkt, id_dolg, id_shtat                                                    \
                       from                                                                                         \
                           (select tn, initcap(fio) as fio, direkt from ocenka where tn not in (select tn_sap from sap_decr) and god="+IntToStr(god)+") o                                   \
                           left join                                                                                \
                           (select tn_sap, zex, ur1,ur2,ur3, ur4, id_dolg, id_shtat from sap_osn_sved) sv                \
                           on tn=sv.tn_sap ) rukov                                                                  \
                     on id_shtat=dolgn.n_shtat and id_dolg=dolgn.short                                              \
                    left join                                                                                       \
                      (select oc.tn, initcap(fio) as fio, zex, dolg, direkt,                                       \
                              orez.id_shtat as shtat_zam,                                                                \
                              orez.tn_sap_rez as tn_sap_zam, orez.zex_rez as zex_rez,                                    \
                              orez.shifr_rez as shifr_rez                                                                \
                      from ocenka oc left join ocenka_rez orez on oc.tn=orez.tn where oc.god="+IntToStr(god)+" and nvl(oc.zam,0)=0 and orez.god="+IntToStr(god)+") zamesh \
                      on dolgn.n_shtat=zamesh.shtat_zam                                                                             \
            ) order by nvl(rukov.direkt,0), shifr_zex, short";

  DM->qObnovlenie->Close();
  DM->qObnovlenie->SQL->Clear();
  DM->qObnovlenie->SQL->Add(Sql);
  try
    {
      DM->qObnovlenie->Open();
    }
  catch (Exception &E)
    {
      Application->MessageBox(("Возникла ошибка при выборке данных из картотеки по оценке персонала и кадров\n(OCENKA, SAP_OSN_SVED, SAP_PEREVOD)"+E.Message).c_str(),"Ошибка",
                              MB_OK+MB_ICONERROR);

      StatusBar1->SimpleText ="Отчетный период: "+IntToStr(god)+" год";
      Abort();
    }

  if (DM->qObnovlenie->RecordCount==0)
    {
      Application->MessageBox(("Нет данных для формирования отчета за "+IntToStr(god)+" год").c_str(),"Предупреждение",
                              MB_OK+MB_ICONINFORMATION);

      StatusBar1->SimpleText ="Отчетный период: "+IntToStr(god)+" год";
      Abort();
    }

  Cursor = crHourGlass;
  ProgressBar->Position = 0;
  ProgressBar->Visible = true;
  ProgressBar->Max=DM->qObnovlenie->RecordCount;

  // инициализируем Excel, открываем этот шаблон
  try
    {
      AppEx=CreateOleObject("Excel.Application");
    }
  catch (...)
    {
      Application->MessageBox("Невозможно открыть Microsoft Excel!"
                              " Возможно это приложение на компьютере не установлено.","Ошибка",MB_OK+MB_ICONERROR);
      StatusBar1->SimpleText ="Отчетный период: "+IntToStr(god)+" год";
      ProgressBar->Visible = false;
      Cursor = crDefault;
    }

  while (DM->qObnovlenie->RecordCount>0 && !DM->qObnovlenie->Eof)
    {
      zex = DM->qObnovlenie->FieldByName("direkt")->AsString;

      //Если возникает ошибка во время формирования отчета
      try
        {
          try
            {
              //Создание папки, если ее не существует
              ForceDirectories(WorkPath+"\\Формирование списка КР по дирекциям(сокращенный)");
             
              //Копируем шаблон файла в Мои документы
              CopyFile((Path+"\\RTF\\kr2.xlsx").c_str(), (WorkPath+"\\Формирование списка КР по дирекциям(сокращенный)\\"+zex+".xlsx").c_str(), false);
              sFile = WorkPath+"\\Формирование списка КР по дирекциям(сокращенный)\\"+zex+".xlsx";

              AppEx.OlePropertyGet("WorkBooks").OleProcedure("Open",sFile.c_str())    ;  //открываем книгу, указав её имя
              Sh=AppEx.OlePropertyGet("WorkSheets",1);                               //выбираем № активного листа книги
              //Sh=AppEx.OlePropertyGet("WorkSheets","Расчет");                      //выбираем лист по наименованию
            }
          catch(...)
            {
              Application->MessageBox("Ошибка открытия книги Microsoft Excel!","Ошибка",MB_OK+MB_ICONERROR);
              StatusBar1->SimpleText ="Отчетный период: "+IntToStr(god)+" год";
              ProgressBar->Visible = false;
              Cursor = crDefault;
            }

          i=1;
          n=3;

          zex1 = DM->qObnovlenie->FieldByName("direkt")->AsString;
          zex = DM->qObnovlenie->FieldByName("direkt")->AsString;

          StatusBar1->SimpleText ="Идет формирование отчета по рекомендациям в КР на руководящие должности по дирекциям...  "+zex;

          //Вывод данных в шаблон
          Variant Massiv;
          Massiv = VarArrayCreate(OPENARRAY(int,(0,10)),varVariant); //массив на 10 элементов

          while (!DM->qObnovlenie->Eof && zex==zex1)
            {
              Massiv.PutElement(i, 0);
              Massiv.PutElement(DM->qObnovlenie->FieldByName("n_shtat")->AsString.c_str(), 1);
              Massiv.PutElement(DM->qObnovlenie->FieldByName("short")->AsString.c_str(), 2);
              Massiv.PutElement(DM->qObnovlenie->FieldByName("dolg")->AsString.c_str(), 3);
              Massiv.PutElement(DM->qObnovlenie->FieldByName("r_fio")->AsString.c_str(), 4);
              Massiv.PutElement(DM->qObnovlenie->FieldByName("r_tn")->AsString.c_str(), 5);
              Massiv.PutElement(DM->qObnovlenie->FieldByName("r_direkt")->AsString.c_str(), 6);
              Massiv.PutElement(DM->qObnovlenie->FieldByName("zex")->AsString.c_str(), 7);
              Massiv.PutElement(DM->qObnovlenie->FieldByName("nzex")->AsString.c_str(), 8);
              Massiv.PutElement(DM->qObnovlenie->FieldByName("n_uch")->AsString.c_str(), 9);
              Massiv.PutElement(DM->qObnovlenie->FieldByName("kol")->AsString.c_str(), 10);

              Sh.OlePropertyGet("Range", ("A" + IntToStr(n) + ":K" + IntToStr(n)).c_str()).OlePropertySet("Value", Massiv); //строка с данными с ячейки A по ячейку АВ

              i++;
              n++;
              DM->qObnovlenie->Next();
              ProgressBar->Position++;
              zex1 = DM->qObnovlenie->FieldByName("direkt")->AsString;
            }

          //рисуем сетку
          Sh.OlePropertyGet("Range",("A3:K"+IntToStr(n-1)).c_str()).OlePropertyGet("Borders").OlePropertySet("LineStyle", xlContinuous);

          //Сохранить книгу в папке в файле по указанию
          AppEx.OlePropertyGet("WorkBooks",1).OleFunction("Save");

          //Закрыть книгу Excel с шаблоном для вывода информации
          AppEx.OlePropertyGet("WorkBooks",1).OleProcedure("Close");
        }
      catch(...)
        {
          AppEx.OleProcedure("Quit");
          AppEx = Unassigned;
          Cursor = crDefault;
          ProgressBar->Position=0;
          ProgressBar->Visible = false;

          StatusBar1->SimpleText ="Отчетный период: "+IntToStr(god)+" год";
          Abort();
        }
    }

  //Закрыть открытое приложение Excel
  AppEx.OleProcedure("Quit");
  AppEx.OlePropertySet("AskToUpdateLinks",true);
  AppEx.OlePropertySet("DisplayAlerts",true);
  AppEx = Unassigned;

  Cursor = crDefault;
  ProgressBar->Position=0;
  ProgressBar->Visible = false;

  StatusBar1->SimpleText= "Формирование отчета выполнено.";
  StatusBar1->SimpleText ="Отчетный период: "+IntToStr(god)+" год";

  InsertLog("Формирование КР работников (сокращенный) за "+IntToStr(god)+" год по дирекциям в Excel успешно завершено");
  DM->qLogs->Requery();
  Application->MessageBox("Формирование файлов успешно завершено","Формирование файлов",
                           MB_OK+MB_ICONINFORMATION);
}
//---------------------------------------------------------------------------
//Отчет по рекомендациям в кадровый резерв (сокращенный) по структурному подразделению
void __fastcall TMain::N27Click(TObject *Sender)
{
  AnsiString Answer, Sql;

  if( InputQuery("Формирование отчета по рекомендациям в КР по цеху","Введите шифр подразделения", Answer) == True)
    {
      if(Answer=="" && (StrToInt(Answer)))
        {
          Application->MessageBox("Не введен шифр подразделения","Предупреждение",MB_OK +MB_ICONWARNING);
          Abort();
        }
      else
        {
          OtchetKRSokr(Answer);
        }
   }
}
//---------------------------------------------------------------------------
//Отчет по рекомендациям в кадровый резерв (сокращенный)
void __fastcall TMain::OtchetKRSokr(AnsiString otchet)
{
  AnsiString Sql, sFile;
  int i,n;
  Variant AppEx,Sh;

  Sql="select  distinct rukov.tn as r_tn,                                                                           \
               short,  \
               n_shtat,                                                                                              \
               dolgn.dolg as dolg,                                                                                  \
               rukov.fio as r_fio,                                                                                  \
               case when dolgn.kod_zex in ('47','54','99')  then  (select naim from sp_direkt where god="+IntToStr(god)+" and kod=(select kod_d from sp_pdirekt pdir where pdir.zex=dolgn.kod_szex and pdir.god="+IntToStr(god)+"))  \
                    else (select naim from sp_direkt where god="+IntToStr(god)+" and kod=(select kod_d from sp_pdirekt pdir1 where pdir1.zex=dolgn.kod_zex and pdir1.god="+IntToStr(god)+")) end  as r_direkt,           \                                                         \
               kod_zex as zex,                                                                                      \
               shifr_zex,                                                                                           \
               nzex,                                                                                                \
               uch as n_uch,                                                                                                 \
               case when rukov.tn is not null and zamesh.tn is not null then count(*) over (partition by rukov.tn)  \
               else 0 end as kol                                                                                    \
       from (                                                                                                       \
                      (select stext as dolg,                                                                        \
                              shifr_zex,                                                                            \
                              kod_zex,                                                                              \
                              kod_szex,                                                                             \
                              objid_p1000,                                                                          \
                              n_shtat,                                                                              \
                              short,                                                                                \
                              (select stext from p1000@sapmig_buffdb where otype='O' and langu='R' and stext not like '%(устар.)' and short=kod_zex) as nzex,  \
                              uch                                                                                   \
                                                                                                                    \
                       from (                                                                                       \
                             (select p1.otype, p1.objid as zvezda1,                                                                                                           \
                                     p1.begda, p1.endda, p1.sobid as sobid_p1001,                                                                                             \
                                     p2.stext, kod as zvezda3, p2.objid as n_shtat, p2.short                                                                                  \
                              from (select r.otype, r.objid, r.begda, r.endda, s.sobid as sobid, s.objid as kod                                                               \
                                    from p1013@sapmig_buffdb r left join p1001@sapmig_buffdb s on r.objid=s.objid and s.otype='S' and s.sclas='O' where r.otype='S' and       \
                                    (r.persk=10 or s.objid in (select objid from p1000@sapmig_buffdb where otype='S' and langu='R' and trim(stext) in ('Механик цеха',        \
                                    'Энергетик цеха','Электрик цеха','Механик участка','Энергетик участка','Электрик участка','Механик фабрики','Энергетик фабрики',          \
                                    'Электрик фабрики','Механик управления','Энергетик управления','Электрик управления','Сменный механик участка','Сменный энергетик участка', \
                                    'Сменный электрик участка')) )) p1,                                                                                                          \
                                    p1000@sapmig_buffdb p2                                                                                                                       \
                              where p2.otype='S' and p2.langu='R' and p1.objid=p2.objid  and upper(stext) not like '%МЕНЕДЖЕР%'                                                   \
                              ) obsh1                                                                               \
                             left join                                                                              \
                             (select objid as objid_p1000, short as shifr_zex, substr(short,1,2) as kod_zex,        \
                                     substr(short,1,5) as kod_szex, stext as uch                                    \
                              from p1000@sapmig_buffdb where otype='O' and langu='R'                                \
                              )obsh2                                                                                \
                             on  sobid_p1001=objid_p1000                                                            \
                             )                                                                                      \
                       where endda>sysdate  and substr(shifr_zex,1,2) not in ('10','11','49','50')                                                       \
                      ) dolgn                                                                                       \
                     left join                                                                                      \
                      (select case when sv.ur1 is null then sv.zex    \
                                   when sv.ur2 is null then sv.ur1    \
                                   when sv.ur3 is null then sv.ur2    \
                                   when sv.ur4 is null then sv.ur3 end as z, tn, fio, direkt, id_dolg, id_shtat                                                    \
                       from                                                                                         \
                           (select tn, initcap(fio) as fio, direkt from ocenka where tn not in (select tn_sap from sap_decr) and god="+IntToStr(god)+") o                                   \
                           left join                                                                                \
                           (select tn_sap, zex, ur1,ur2,ur3,ur4, id_dolg, id_shtat from sap_osn_sved) sv                \
                           on tn=sv.tn_sap ) rukov                                                                  \
                     on id_shtat=dolgn.n_shtat and id_dolg=dolgn.short                                              \
                    left join                                                                                       \
                      (select oc.tn, initcap(fio) as fio, zex, dolg, direkt,                                         \
                              orez.id_shtat as shtat_zam,                                                                \
                              orez.tn_sap_rez as tn_sap_zam, orez.zex_rez as zex_rez,                                    \
                              orez.shifr_rez as shifr_rez                                                                \
                      from ocenka oc left join ocenka_rez orez on oc.tn=orez.tn where oc.god="+IntToStr(god)+" and nvl(oc.zam,0)=0 and orez.god="+IntToStr(god)+") zamesh \
                      on dolgn.n_shtat=zamesh.shtat_zam                                                                            \
            )";

  if (otchet==0)
    {
      Sql+=" order by shifr_zex, short";
      StatusBar1->SimpleText=" Идет формирование отчета по рекомендациям в КР на руководящие должности по предприятию в Excel...";
    }
  else
    {
      Sql+=" where shifr_zex like '"+otchet+"%' order by shifr_zex, short";
      StatusBar1->SimpleText=" Идет формирование отчета по рекомендациям в КР на руководящие должности по подразделению в Excel...";                                                                                                                                                \
    }


  DM->qObnovlenie->Close();
  DM->qObnovlenie->SQL->Clear();
  DM->qObnovlenie->SQL->Add(Sql);
  try
    {
      DM->qObnovlenie->Open();
    }
  catch (Exception &E)
    {
      Application->MessageBox(("Возникла ошибка при выборке данных из картотеки по оценке персонала и кадров\n(OCENKA, SAP_OSN_SVED, SAP_PEREVOD)"+E.Message).c_str(),"Ошибка",
                              MB_OK+MB_ICONERROR);

      StatusBar1->SimpleText ="Отчетный период: "+IntToStr(god)+" год";
      Abort();
    }

  if (DM->qObnovlenie->RecordCount==0)
    {
      Application->MessageBox(("Нет данных для формирования отчета за "+IntToStr(god)+" год").c_str(),"Предупреждение",
                              MB_OK+MB_ICONINFORMATION);

      StatusBar1->SimpleText ="Отчетный период: "+IntToStr(god)+" год";
      Abort();
    }

  Cursor = crHourGlass;
  ProgressBar->Position = 0;
  ProgressBar->Visible = true;
  ProgressBar->Max=DM->qObnovlenie->RecordCount;

  // инициализируем Excel, открываем этот шаблон
  try
    {
      AppEx=CreateOleObject("Excel.Application");
    }
  catch (...)
    {
      Application->MessageBox("Невозможно открыть Microsoft Excel!"
                              " Возможно это приложение на компьютере не установлено.","Ошибка",MB_OK+MB_ICONERROR);
      StatusBar1->SimpleText ="Отчетный период: "+IntToStr(god)+" год";
      ProgressBar->Visible = false;
      Cursor = crDefault;
    }

  //Если возникает ошибка во время формирования отчета
  try
    {
      try
        {
          AppEx.OlePropertySet("AskToUpdateLinks",false);
          AppEx.OlePropertySet("DisplayAlerts",false);

          //Создание папки, если ее не существует
          ForceDirectories(WorkPath);

          //Создание папки, если ее не существует
          ForceDirectories(WorkPath+"\\Формирование списка КР(сокращенный)");

          if (otchet==0)
            {
              //Копируем шаблон файла в Мои документы
              CopyFile((Path+"\\RTF\\kr2.xlsx").c_str(), (WorkPath+"\\Формирование списка КР(сокращенный)\\По предприятию.xlsx").c_str(), false);
              sFile = WorkPath+"\\Формирование списка КР(сокращенный)\\По предприятию.xlsx";
            }
          else
            {
              //Копируем шаблон файла в Мои документы
              CopyFile((Path+"\\RTF\\kr2.xlsx").c_str(), (WorkPath+"\\Формирование списка КР(сокращенный)\\"+otchet+".xlsx").c_str(), false);
              sFile = WorkPath+"\\Формирование списка КР(сокращенный)\\"+otchet+".xlsx";
            }

          AppEx.OlePropertyGet("WorkBooks").OleProcedure("Open",sFile.c_str())    ;  //открываем книгу, указав её имя
          Sh=AppEx.OlePropertyGet("WorkSheets",1);                               //выбираем № активного листа книги
          //Sh=AppEx.OlePropertyGet("WorkSheets","Расчет");                      //выбираем лист по наименованию
        }
      catch(...)
        {
          Application->MessageBox("Ошибка открытия книги Microsoft Excel!","Ошибка",MB_OK+MB_ICONERROR);
          StatusBar1->SimpleText ="Отчетный период: "+IntToStr(god)+" год";
          ProgressBar->Visible = false;
          Cursor = crDefault;
        }


      i=1;
      n=3;

      //Вывод данных в шаблон
      Variant Massiv;
      Massiv = VarArrayCreate(OPENARRAY(int,(0,10)),varVariant); //массив на 10 элементов

      while (!DM->qObnovlenie->Eof)
        {
          Massiv.PutElement(i, 0);
          Massiv.PutElement(DM->qObnovlenie->FieldByName("n_shtat")->AsString.c_str(), 1);
          Massiv.PutElement(DM->qObnovlenie->FieldByName("short")->AsString.c_str(), 2);
          Massiv.PutElement(DM->qObnovlenie->FieldByName("dolg")->AsString.c_str(), 3);
          Massiv.PutElement(DM->qObnovlenie->FieldByName("r_fio")->AsString.c_str(), 4);
          Massiv.PutElement(DM->qObnovlenie->FieldByName("r_tn")->AsString.c_str(), 5);
          Massiv.PutElement(DM->qObnovlenie->FieldByName("r_direkt")->AsString.c_str(), 6);
          Massiv.PutElement(DM->qObnovlenie->FieldByName("zex")->AsString.c_str(), 7);
          Massiv.PutElement(DM->qObnovlenie->FieldByName("nzex")->AsString.c_str(), 8);
          Massiv.PutElement(DM->qObnovlenie->FieldByName("n_uch")->AsString.c_str(), 9);
          Massiv.PutElement(DM->qObnovlenie->FieldByName("kol")->AsString.c_str(), 10);


          Sh.OlePropertyGet("Range", ("A" + IntToStr(n) + ":K" + IntToStr(n)).c_str()).OlePropertySet("Value", Massiv); //строка с данными с ячейки A по ячейку АВ

          i++;
          n++;
          DM->qObnovlenie->Next();
          ProgressBar->Position++;
        }

      //рисуем сетку
      Sh.OlePropertyGet("Range",("A3:K"+IntToStr(n-1)).c_str()).OlePropertyGet("Borders").OlePropertySet("LineStyle", xlContinuous);

     // Sh.OleProcedure("SaveAs",vAsCurDir1.c_str());
     AppEx.OlePropertyGet("WorkBooks",1).OleFunction("Save");

      /* //Закрыть открытое приложение Excel
      AppEx.OleProcedure("Quit");
      AppEx = Unassigned;  */

      //Закрыть книгу Excel с шаблоном для вывода информации
     // AppEx.OlePropertyGet("WorkBooks",1).OleProcedure("Close");
      Application->MessageBox("Отчет успешно сформирован!", "Формирование отчета",
                               MB_OK+MB_ICONINFORMATION);
      //AppEx.OlePropertyGet("WorkBooks").OleProcedure("Open",vAsCurDir1.c_str());
      AppEx.OlePropertySet("Visible",true);
      AppEx.OlePropertySet("AskToUpdateLinks",true);
      AppEx.OlePropertySet("DisplayAlerts",true);


      StatusBar1->SimpleText= "Формирование отчета выполнено.";

      Cursor = crDefault;
      ProgressBar->Position=0;
      ProgressBar->Visible = false;
      StatusBar1->SimpleText ="Отчетный период: "+IntToStr(god)+" год";

      if (otchet==0) InsertLog("Формирование КР работников (сокращенный) за "+IntToStr(god)+" год по предприятию в Excel успешно завершено");
      else InsertLog("Формирование КР работников (сокращенный) за "+IntToStr(god)+" год по подразделению "+otchet+" в Excel успешно завершено");
      DM->qLogs->Requery();
    }
  catch(...)
    {
      AppEx.OleProcedure("Quit");
      AppEx = Unassigned;
      Cursor = crDefault;
      ProgressBar->Position=0;
      ProgressBar->Visible = false;

      StatusBar1->SimpleText ="Отчетный период: "+IntToStr(god)+" год";
      Abort();
    }
}
//---------------------------------------------------------------------------

//Обновление ФИО оценщика на "Уволен"
void __fastcall TMain::NUvolClick(TObject *Sender)
{
  AnsiString Sql;

  if ((Application->MessageBox("Вы действительно хотите обновить информацию в поле \n'ФИО оценщика' по уволенным работникам на значение 'Уволен'?","",
                           MB_YESNO+MB_ICONINFORMATION))==ID_NO)
    {
      Abort();
    }

  Sql ="update ocenka set fio_ocen='Уволен', dolg_ocen=null, data_ocen=null \
        where tn in (select tn from ocenka \
                     where god="+IntToStr(god)+" and tn in (select tn_sap from sap_sved_uvol) \
                     and fio_ocen!='Уволен' \
                     and tn not in (select tn_sap from sap_osn_sved))";

  DM->qObnovlenie->Close();
  DM->qObnovlenie->SQL->Clear();
  DM->qObnovlenie->SQL->Add(Sql);
  try
    {
      DM->qObnovlenie->Open();
    }
  catch(Exception &E)
    {
      Application->MessageBox(("Возникла ошибка при попытке обновить данные по уволенным из таблицы"+E.Message).c_str(),"",
                              MB_OK+MB_ICONINFORMATION);
      InsertLog("Обновление по уволенным выполнено успешно");
      Abort();
    }

  InsertLog("Обновление по уволенным выполнено успешно");

  Application->MessageBox("Обновление по уволенным выполнено успешно","Обновление данных по уволенным",
                              MB_OK+MB_ICONINFORMATION);
}
//---------------------------------------------------------------------------

//План преемственности по предприятию
void __fastcall TMain::N44Click(TObject *Sender)
{
  OtchetPP(0);
}
//---------------------------------------------------------------------------

//План преемственности по структурному подразделению
void __fastcall TMain::N45Click(TObject *Sender)
{
   AnsiString Answer, Sql;

  if( InputQuery("Формирование 'Плана преемственности'","Введите шифр подразделения", Answer) == True)
    {
      if(Answer=="" && (StrToInt(Answer)))
        {
          Application->MessageBox("Не введен шифр подразделения","Предупреждение",MB_OK +MB_ICONWARNING);
          Abort();
        }
      else
        {
          OtchetPP(Answer);
        }
   }
}
//---------------------------------------------------------------------------

//Отчет "План преемственности"
void __fastcall TMain::OtchetPP(AnsiString otchet)
{
  AnsiString Sql, sFile, tn, tn1;
  int i,n, yanvar, fevral, mart, aprel, mai, iyn, iyl, avgust, sentyabr, oktyabr, noyabr, dekabr, pos_kol=0, kol_dolg=1;
  Variant AppEx,Sh;

  
  Sql=" select vse.*,                                                                                                   \
       max(kol_period1) over () as max_kol_kolonok1,                                                                    \
       max(kol_period2) over () as max_kol_kolonok2,                                                                    \
       max(kol_period3) over () as max_kol_kolonok3,                                                                    \
       max(kol_period4) over () as max_kol_kolonok4,                                                                    \
       max(kol_period5) over () as max_kol_kolonok5,                                                                    \
       max(kol_period6) over () as max_kol_kolonok6,                                                                    \
       max(kol_period7) over () as max_kol_kolonok7,                                                                    \
       max(kol_period8) over () as max_kol_kolonok8,                                                                    \
       max(kol_period9) over () as max_kol_kolonok9,                                                                    \
       max(kol_period10) over () as max_kol_kolonok10,                                                                  \
       max(kol_period11) over () as max_kol_kolonok11,                                                                  \
       max(kol_period12) over () as max_kol_kolonok12 \
       from (select                                                                                                           \
       case when dolgn.kod_zex in ('47','54','99')  then  (select naim from sp_direkt where god="+IntToStr(god)+" and kod=(select kod_d from sp_pdirekt pdir where pdir.zex=dolgn.kod_szex and pdir.god="+IntToStr(god)+"))\
            else (select naim from sp_direkt where god="+IntToStr(god)+" and kod=(select kod_d from sp_pdirekt pdir1 where pdir1.zex=dolgn.kod_zex and pdir1.god="+IntToStr(god)+")) end  as r_direkt, \
       case when dolgn.kod_zex in ('47','54','99')  then dolgn.kod_szex         \
       else dolgn.kod_zex  end as kod_zex, \                                                                                      \
       nzex,                                                                                                            \
       rukov.fio as r_fio,        \
       decode(length(zamesh.uch),2,NULL,(select stext from p1000@sapmig_buffdb where otype='O' and langu='R' and stext not like '%(уст%' and short=zamesh.uch)) as z_nuch, \
       nuch,                    \
       dolgn.n_shtat as n_shtat, \                                                                                            \
       dolgn.short as short,                                                                                            \
       dolgn.dolg as dolg,                                                                                              \
       rukov.tn as r_tn,                                                                                                \
       case when (select count(*) from ocenka_rez k where k.tn=rukov.tn)>0 then 'да'      \
       else NULL end as r_preem,                                                                                           \
       rukov.dat_r as r_vz,                                                                                             \
       decode(rukov.kpe1,0,'-',rukov.kpe1) as r_kpe1,     \
       decode(rukov.kpe2,0,'-',rukov.kpe2) as r_kpe2,     \
       decode(rukov.kpe3,0,'-',rukov.kpe3) as r_kpe3,     \
       decode(rukov.kpe4,0,'-',rukov.kpe4) as r_kpe4,     \                                                                                       \
       case when (decode(nvl(rukov.kpe1,0),0,0,1)+decode(nvl(rukov.kpe2,0),0,0,1)+decode(nvl(rukov.kpe3,0),0,0,1)+decode(nvl(rukov.kpe4,0),0,0,1))=0 then NULL      \
       else round(((nvl(rukov.kpe1,0)+nvl(rukov.kpe2,0)+nvl(rukov.kpe3,0)+nvl(rukov.kpe4,0))/(decode(nvl(rukov.kpe1,0),0,0,1)+decode(nvl(rukov.kpe2,0),0,0,1)+decode(nvl(rukov.kpe3,0),0,0,1)+decode(nvl(rukov.kpe4,0),0,0,1))),2) \
       end as r_kpe,                                                                                   \
       decode(rukov.tn,NULL,'высокий', decode(zamesh.risk,1,'высокий',2,'средний',3,'низкий','')) as r_risk,           \
       zamesh.risk_prich as r_risk_prich,                                                                               \
       rukov.kom_reit as r_ocenka,                                                                                      \
       rukov.rezerv as r_rezerv,                                                                                        \
       rukov.vz_pens as r_vz_pens,                                                                                      \
       case when (rukov.zex, dolgn.short) in (select zex, shifr_dolg from sp_ocenka_krd) then 'да'                       \
            else NULL end as r_krd,                                                                                     \
       decode((select max(gotov) from ocenka_rez where tn=rukov.tn and id_shtat!=n_shtat and god="+IntToStr(Main->god)+"),1,'низкая',2,'средняя',3,'высокая','') as r_gotov,   \
       zamesh.tn as z_tn,                                                                                               \
       zamesh.fio as z_fio,                                                                                             \
       zamesh.dolg as z_dolg,                                                                                           \
       zamesh.zex as z_zex,                                                                                             \
       case when substr(zamesh.zex,1,2)not in ('47','54','99') then (select nazv_cexk from ssap_cex where id_cex=substr(zamesh.zex,1,2) and nazv_cexk not like '%(устар.)') \
       else (select nazv_cexk from ssap_cex where id_cex=substr(zamesh.zex,1,5) and nazv_cexk not like '%(устар.)') end as z_nzex,                                          \
       z_vz,                \
       decode(zamesh.kpe1,0,'-',zamesh.kpe1) as z_kpe1,      \
       decode(zamesh.kpe2,0,'-',zamesh.kpe2) as z_kpe2,      \
       decode(zamesh.kpe3,0,'-',zamesh.kpe3) as z_kpe3,      \
       decode(zamesh.kpe4,0,'-',zamesh.kpe4) as z_kpe4,      \                                                                                       \
       case when (decode(nvl(zamesh.kpe1,0),0,0,1)+decode(nvl(zamesh.kpe2,0),0,0,1)+decode(nvl(zamesh.kpe3,0),0,0,1)+decode(nvl(zamesh.kpe4,0),0,0,1))=0 then NULL \
       else round(((nvl(zamesh.kpe1,0)+nvl(zamesh.kpe2,0)+nvl(zamesh.kpe3,0)+nvl(zamesh.kpe4,0))/(decode(nvl(zamesh.kpe1,0),0,0,1)+decode(nvl(zamesh.kpe2,0),0,0,1)+decode(nvl(zamesh.kpe3,0),0,0,1)+decode(nvl(zamesh.kpe4,0),0,0,1))),2) \
       end  as z_kpe,                                                                                      \                                                                                                                                \
       zamesh.rezerv as z_rezerv,                                                                                         \
       zamesh.kom_reit as z_ocenka,                                                                                       \
       zamesh.vz_pens as z_vz_pens,                                                                                       \
       decode(zamesh.gotov,1,'низкая',2,'средняя',3,'высокая','') as z_gotov,                                             \
       decode(m1,'01',s1+1,s1)+decode(m2,'01',s1+1,s1)+decode(m3,'01',s1+1,s1)+decode(m4,'01',s1+1,s1)+                               \
       decode(m5,'01',s1+1,s1)+decode(m6,'01',s1+1,s1)+decode(m7,'01',s1+1,s1)+decode(m8,'01',s1+1,s1)                                \
       +decode(m9,'01',s1+1,s1)+decode(m10,'01',s1+1,s1)+decode(m11,'01',s1+1,s1)+decode(m12,'01',s1+1,s1) as kol_period1,            \
       \
       decode(m1,'02',s2+1,s2)+decode(m2,'02',s2+1,s2)+decode(m3,'02',s2+1,s2)+decode(m4,'02',s2+1,s2)+                               \
       decode(m5,'02',s2+1,s2)+decode(m6,'02',s2+1,s2)+decode(m7,'02',s2+1,s2)+decode(m8,'02',s2+1,s2)                                \
       +decode(m9,'02',s2+1,s2)+decode(m10,'02',s2+1,s2)+decode(m11,'02',s2+1,s2)+decode(m12,'02',s2+1,s2) as kol_period2,            \
                                                                                                                          \
       decode(m1,'03',s3+1,s3)+decode(m2,'03',s3+1,s3)+decode(m3,'03',s3+1,s3)+decode(m4,'03',s3+1,s3)+                               \
       decode(m5,'03',s3+1,s3)+decode(m6,'03',s3+1,s3)+decode(m7,'03',s3+1,s3)+decode(m8,'03',s3+1,s3)                                \
       +decode(m9,'03',s3+1,s3)+decode(m10,'03',s3+1,s3)+decode(m11,'03',s3+1,s3)+decode(m12,'03',s3+1,s3) as kol_period3,            \
                                                                                                                          \
       decode(m1,'04',s4+1,s4)+decode(m2,'04',s4+1,s4)+decode(m3,'04',s4+1,s4)+decode(m4,'04',s4+1,s4)+                               \
       decode(m5,'04',s4+1,s4)+decode(m6,'04',s4+1,s4)+decode(m7,'04',s4+1,s4)+decode(m8,'04',s4+1,s4)                                \
       +decode(m9,'04',s4+1,s4)+decode(m10,'04',s4+1,s4)+decode(m11,'04',s4+1,s4)+decode(m12,'04',s4+1,s4) as kol_period4,            \
                                                                                                                          \
       decode(m1,'05',s5+1,s5)+decode(m2,'05',s5+1,s5)+decode(m3,'05',s5+1,s5)+decode(m4,'05',s5+1,s5)+                               \
       decode(m5,'05',s5+1,s5)+decode(m6,'05',s5+1,s5)+decode(m7,'05',s5+1,s5)+decode(m8,'05',s5+1,s5)                                \
       +decode(m9,'05',s5+1,s5)+decode(m10,'05',s5+1,s5)+decode(m11,'05',s5+1,s5)+decode(m12,'05',s5+1,s5) as kol_period5,            \
                                                                                                                          \
       decode(m1,'06',s6+1,s6)+decode(m2,'06',s6+1,s6)+decode(m3,'06',s6+1,s6)+decode(m4,'06',s6+1,s6)+                               \
       decode(m5,'06',s6+1,s6)+decode(m6,'06',s6+1,s6)+decode(m7,'06',s6+1,s6)+decode(m8,'06',s6+1,s6)                                \
       +decode(m9,'06',s6+1,s6)+decode(m10,'06',s6+1,s6)+decode(m11,'06',s6+1,s6)+decode(m12,'06',s6+1,s6) as kol_period6,            \
                                                                                                                          \
       decode(m1,'07',s7+1,s7)+decode(m2,'07',s7+1,s7)+decode(m3,'07',s7+1,s7)+decode(m4,'07',s7+1,s7)+                               \
       decode(m5,'07',s7+1,s7)+decode(m6,'07',s7+1,s7)+decode(m7,'07',s7+1,s7)+decode(m8,'07',s7+1,s7)                                \
       +decode(m9,'07',s7+1,s7)+decode(m10,'07',s7+1,s7)+decode(m11,'07',s7+1,s7)+decode(m12,'07',s7+1,s7) as kol_period7,            \
                                                                                                                          \
       decode(m1,'08',s8+1,s8)+decode(m2,'08',s8+1,s8)+decode(m3,'08',s8+1,s8)+decode(m4,'08',s8+1,s8)+                               \
       decode(m5,'08',s8+1,s8)+decode(m6,'08',s8+1,s8)+decode(m7,'08',s8+1,s8)+decode(m8,'08',s8+1,s8)                                \
       +decode(m9,'08',s8+1,s8)+decode(m10,'08',s8+1,s8)+decode(m11,'08',s8+1,s8)+decode(m12,'08',s8+1,s8) as kol_period8,            \
                                                                                                                          \
       decode(m1,'09',s9+1,s9)+decode(m2,'09',s9+1,s9)+decode(m3,'09',s9+1,s9)+decode(m4,'09',s9+1,s9)+                               \
       decode(m5,'09',s9+1,s9)+decode(m6,'09',s9+1,s9)+decode(m7,'09',s9+1,s9)+decode(m8,'09',s9+1,s9)                                \
       +decode(m9,'09',s9+1,s9)+decode(m10,'09',s9+1,s9)+decode(m11,'09',s9+1,s9)+decode(m12,'09',s9+1,s9) as kol_period9,            \
                                                                                                                          \
       decode(m1,'10',s10+1,s10)+decode(m2,'10',s10+1,s10)+decode(m3,'10',s10+1,s10)+decode(m4,'10',s10+1,s10)+                   \
       decode(m5,'10',s10+1,s10)+decode(m6,'10',s10+1,s10)+decode(m7,'10',s10+1,s10)+decode(m8,'10',s10+1,s10)                    \
       +decode(m9,'10',s10+1,s10)+decode(m10,'10',s10+1,s10)+decode(m11,'10',s10+1,s10)+decode(m12,'10',s10+1,s10) as kol_period10, \
                                                                                                                            \
       decode(m1,'11',s11+1,s11)+decode(m2,'11',s11+1,s11)+decode(m3,'11',s11+1,s11)+decode(m4,'11',s11+1,s11)+                     \
       decode(m5,'11',s11+1,s11)+decode(m6,'11',s11+1,s11)+decode(m7,'11',s11+1,s11)+decode(m8,'11',s11+1,s11)                      \
       +decode(m9,'11',s11+1,s11)+decode(m10,'11',s11+1,s11)+decode(m11,'11',s11+1,s11)+decode(m12,'11',s11+1,s11) as kol_period11, \
                                                                                                                            \
       decode(m1,'12',s12+1,s12)+decode(m2,'12',s12+1,s12)+decode(m3,'12',s12+1,s12)+decode(m4,'12',s12+1,s12)+                     \
       decode(m5,'12',s12+1,s12)+decode(m6,'12',s12+1,s12)+decode(m7,'12',s12+1,s12)+decode(m8,'12',s12+1,s12)                      \
       +decode(m9,'12',s12+1,s12)+decode(m10,'12',s12+1,s12)+decode(m11,'12',s12+1,s12)+decode(m12,'12',s12+1,s12) as kol_period12,  \
       datn1, datn2, datn3, datn4, datn5, datn6, datn7, datn8, datn9, datn10, datn11, datn12,     \
       datk1, datk2, datk3, datk4, datk5, datk6, datk7, datk8, datk9, datk10, datk11, datk12,     \
       decode((nvl((to_date(to_char(datk1, 'dd.mm.yyyy'), 'dd.mm.yyyy')-to_date(to_char(datn1, 'dd.mm.yyyy'), 'dd.mm.yyyy')+1),0)+nvl((to_date(to_char(datk2, 'dd.mm.yyyy'), 'dd.mm.yyyy')-to_date(to_char(datn2, 'dd.mm.yyyy'), 'dd.mm.yyyy')+1),0)+                        \
        nvl((to_date(to_char(datk3, 'dd.mm.yyyy'), 'dd.mm.yyyy')-to_date(to_char(datn3, 'dd.mm.yyyy'), 'dd.mm.yyyy')+1),0)+nvl((to_date(to_char(datk4, 'dd.mm.yyyy'), 'dd.mm.yyyy')-to_date(to_char(datn4, 'dd.mm.yyyy'), 'dd.mm.yyyy')+1),0)+                        \
        nvl((to_date(to_char(datk5, 'dd.mm.yyyy'), 'dd.mm.yyyy')-to_date(to_char(datn5, 'dd.mm.yyyy'), 'dd.mm.yyyy')+1),0)+nvl((to_date(to_char(datk6, 'dd.mm.yyyy'), 'dd.mm.yyyy')-to_date(to_char(datn6, 'dd.mm.yyyy'), 'dd.mm.yyyy')+1),0)+                        \
        nvl((to_date(to_char(datk7, 'dd.mm.yyyy'), 'dd.mm.yyyy')-to_date(to_char(datn7, 'dd.mm.yyyy'), 'dd.mm.yyyy')+1),0)+nvl((to_date(to_char(datk8, 'dd.mm.yyyy'), 'dd.mm.yyyy')-to_date(to_char(datn8, 'dd.mm.yyyy'), 'dd.mm.yyyy')+1),0)+                        \
        nvl((to_date(to_char(datk9, 'dd.mm.yyyy'), 'dd.mm.yyyy')-to_date(to_char(datn9, 'dd.mm.yyyy'), 'dd.mm.yyyy')+1),0)+nvl((to_date(to_char(datk10, 'dd.mm.yyyy'), 'dd.mm.yyyy')-to_date(to_char(datn10, 'dd.mm.yyyy'), 'dd.mm.yyyy')+1),0)+                      \
        nvl((to_date(to_char(datk11, 'dd.mm.yyyy'), 'dd.mm.yyyy')-to_date(to_char(datn11, 'dd.mm.yyyy'), 'dd.mm.yyyy')+1),0)+nvl((to_date(to_char(datk12, 'dd.mm.yyyy'), 'dd.mm.yyyy')-to_date(to_char(datn12, 'dd.mm.yyyy'), 'dd.mm.yyyy')+1),0)), '0',NULL,          \
        (nvl((to_date(to_char(datk1, 'dd.mm.yyyy'), 'dd.mm.yyyy')-to_date(to_char(datn1, 'dd.mm.yyyy'), 'dd.mm.yyyy')+1),0)+nvl((to_date(to_char(datk2, 'dd.mm.yyyy'), 'dd.mm.yyyy')-to_date(to_char(datn2, 'dd.mm.yyyy'), 'dd.mm.yyyy')+1),0)+                        \
        nvl((to_date(to_char(datk3, 'dd.mm.yyyy'), 'dd.mm.yyyy')-to_date(to_char(datn3, 'dd.mm.yyyy'), 'dd.mm.yyyy')+1),0)+nvl((to_date(to_char(datk4, 'dd.mm.yyyy'), 'dd.mm.yyyy')-to_date(to_char(datn4, 'dd.mm.yyyy'), 'dd.mm.yyyy')+1),0)+                        \
        nvl((to_date(to_char(datk5, 'dd.mm.yyyy'), 'dd.mm.yyyy')-to_date(to_char(datn5, 'dd.mm.yyyy'), 'dd.mm.yyyy')+1),0)+nvl((to_date(to_char(datk6, 'dd.mm.yyyy'), 'dd.mm.yyyy')-to_date(to_char(datn6, 'dd.mm.yyyy'), 'dd.mm.yyyy')+1),0)+                        \
        nvl((to_date(to_char(datk7, 'dd.mm.yyyy'), 'dd.mm.yyyy')-to_date(to_char(datn7, 'dd.mm.yyyy'), 'dd.mm.yyyy')+1),0)+nvl((to_date(to_char(datk8, 'dd.mm.yyyy'), 'dd.mm.yyyy')-to_date(to_char(datn8, 'dd.mm.yyyy'), 'dd.mm.yyyy')+1),0)+                        \
        nvl((to_date(to_char(datk9, 'dd.mm.yyyy'), 'dd.mm.yyyy')-to_date(to_char(datn9, 'dd.mm.yyyy'), 'dd.mm.yyyy')+1),0)+nvl((to_date(to_char(datk10, 'dd.mm.yyyy'), 'dd.mm.yyyy')-to_date(to_char(datn10, 'dd.mm.yyyy'), 'dd.mm.yyyy')+1),0)+                      \
        nvl((to_date(to_char(datk11, 'dd.mm.yyyy'), 'dd.mm.yyyy')-to_date(to_char(datn11, 'dd.mm.yyyy'), 'dd.mm.yyyy')+1),0)+nvl((to_date(to_char(datk12, 'dd.mm.yyyy'), 'dd.mm.yyyy')-to_date(to_char(datn12, 'dd.mm.yyyy'), 'dd.mm.yyyy')+1),0))) as kol_day_zam,    \
       m1,m2,m3,m4,m5,m6,m7,m8,m9,m10,m11,m12 \
       from (                                                                                                             \
                      (select stext as dolg,                                                                              \
                              shifr_zex,                                                                                  \
                              kod_zex,                                                                                    \
                              kod_szex,                                                                                   \
                              objid_p1000,                                                                                \
                              n_shtat,                                                                                    \
                              short,                                                                                      \
                              case when kod_zex in ('47', '99', '54') then (select stext from p1000@sapmig_buffdb where otype='O' and langu='R' and stext not like '%(устар.)' and short=kod_szex)  \
                                    else (select stext from p1000@sapmig_buffdb where otype='O' and langu='R' and stext not like '%(устар.)' and short=kod_zex) end as nzex,              \
                              decode(length(shifr_zex),2,NULL,(select stext from p1000@sapmig_buffdb where otype='O' and langu='R' and stext not like '%(уст%' and short=shifr_zex)) as nuch, \
                              uch                                                                                                                                       \
                                                                                                                                                                        \
                       from (                                                                                                                                           \
                             (select p1.otype, p1.objid as zvezda1,                                                                                                     \
                                     p1.begda, p1.endda, p1.sobid as sobid_p1001,                                                                                       \
                                     p2.stext, kod as zvezda3, p2.objid as n_shtat, p2.short                                                                            \
                              from (select r.otype, r.objid, r.begda, r.endda, s.sobid as sobid, s.objid as kod                                                         \
                                    from p1013@sapmig_buffdb r left join p1001@sapmig_buffdb s on r.objid=s.objid and s.otype='S' and s.sclas='O' where r.otype='S' and    \
                                    (r.persk=10 or s.objid in (select objid from p1000@sapmig_buffdb where otype='S' and langu='R' and trim(stext) in ('Механик цеха',     \
                                    'Энергетик цеха','Электрик цеха','Механик участка','Энергетик участка','Электрик участка','Механик фабрики','Энергетик фабрики',       \
                                    'Электрик фабрики','Механик управления','Энергетик управления','Электрик управления','Сменный механик участка','Сменный энергетик участка', \
                                    'Сменный электрик участка')) )) p1, \
                                    p1000@sapmig_buffdb p2                                                                                                              \
                              where p2.otype='S' and p2.langu='R' and p1.objid=p2.objid  and upper(stext) not like '%МЕНЕДЖЕР%'                                                                                \
                              ) obsh1                                                                                                                                   \
                             left join                                                                                                                                  \
                             (select objid as objid_p1000, short as shifr_zex, substr(short,1,2) as kod_zex,                                                            \
                                     substr(short,1,5) as kod_szex, stext as uch                                                                                        \
                              from p1000@sapmig_buffdb where otype='O' and langu='R'                                                                                    \
                              )obsh2                                                                                                                                    \
                             on  sobid_p1001=objid_p1000                                                                                                                \
                             )                                                                                                                                          \
                       where endda>sysdate  and substr(shifr_zex,1,2) not in ('10','11','49','50')                                                                                                     \
                      ) dolgn                                                                                                                                           \
                     left join                                                                                                                                          \
                      (select *                                                                                                                                         \
                       from                                                                                                                                             \
                           (select tn, initcap(fio) as fio, direkt,                                                                                                     \
                                   risk, risk_prich,                                                                                                                    \
                                   round(kpe1,2) as kpe1, round(kpe2,2) as kpe2, round(kpe3,2) as kpe3, round(kpe4,2) as kpe4,                                          \                                             \
                                   kom_reit,                                                                                                                            \
                                   decode(rezerv,1,'да','') as rezerv, vz_pens                                                                                          \
                            from ocenka where god="+IntToStr(god)+" and tn not in (select tn_sap from sap_decr)) o                                                                                                  \                            \
                           left join                                                                                                                                    \
                           (select tn_sap, zex, ur1,ur2,ur3, id_dolg, id_shtat,        \                                                                                 \
                                   case when ur1 is null then zex                     \
                                        when ur2 is null then ur1                     \
                                        when ur3 is null then ur2                     \
                                        when ur4 is null then ur3 end as z,           \
                                   trunc(months_between(sysdate, dat_r)/12) as dat_r  \                                                                                  \
                            from sap_osn_sved) sv                                                                                                                       \
                           on tn=sv.tn_sap )  rukov                                                                                                                     \
                     on id_shtat=dolgn.n_shtat and id_dolg=dolgn.short                                                                                                  \
                     left join                                                                                                                                          \
                      (select *                                                                                                                                         \
                       from (                                                                                                                                           \
                              (select oc.tn, orez.tn as zam_tn, initcap(fio) as fio, direkt,  uch,                                                                      \
                                      orez.gotov,                                                                                                                       \
                                      round(kpe1,2) as kpe1, round(kpe2,2) as kpe2, round(kpe3,2) as kpe3, round(kpe4,2) as kpe4, vz_pens, kom_reit,                    \
                                      decode(rezerv,1,'да','') as rezerv,                                                                                               \
                                      orez.datn1, orez.datn2, orez.datn3, orez.datn4, orez.datn5, orez.datn6, orez.datn7, orez.datn8, orez.datn9, orez.datn10, orez.datn11, orez.datn12,  \
                                      orez.datk1, orez.datk2, orez.datk3, orez.datk4, orez.datk5, orez.datk6, orez.datk7, orez.datk8, orez.datk9, orez.datk10, orez.datk11, orez.datk12,  \
                                      substr(to_char(orez.datn1,'dd.mm.yyyy'),4,2) as m1,                                                                                                 \
                                      substr(to_char(orez.datn2,'dd.mm.yyyy'),4,2) as m2,                                                                                                 \
                                      substr(to_char(orez.datn3,'dd.mm.yyyy'),4,2) as m3,                                                                                                 \
                                      substr(to_char(orez.datn4,'dd.mm.yyyy'),4,2) as m4,                                                                                                 \
                                      substr(to_char(orez.datn5,'dd.mm.yyyy'),4,2) as m5,                                                                                                 \
                                      substr(to_char(orez.datn6,'dd.mm.yyyy'),4,2) as m6,                                                                                                 \
                                      substr(to_char(orez.datn7,'dd.mm.yyyy'),4,2) as m7,                                                                                                 \
                                      substr(to_char(orez.datn8,'dd.mm.yyyy'),4,2) as m8,                                                                                                 \
                                      substr(to_char(orez.datn9,'dd.mm.yyyy'),4,2) as m9,                                                                                                 \
                                      substr(to_char(orez.datn10,'dd.mm.yyyy'),4,2) as m10,                                                                                               \
                                      substr(to_char(orez.datn11,'dd.mm.yyyy'),4,2) as m11,                                                                                               \
                                      substr(to_char(orez.datn12,'dd.mm.yyyy'),4,2) as m12,                                                                                               \
                                      0 as s1,                                                                                                                                            \
                                      0 as s2,                                                                                                                                            \
                                      0 as s3,                                                                                                                                            \
                                      0 as s4,                                                                                                                                            \
                                      0 as s5,                                                                                                                                            \
                                      0 as s6,                                                                                                                                            \
                                      0 as s7,                                                                                                                                            \
                                      0 as s8,                                                                                                                                            \
                                      0 as s9,                                                                                                                                            \
                                      0 as s10,                                                                                                                                           \
                                      0 as s11,                                                                                                                                           \
                                      0 as s12,                                                                                                                                           \
                                      orez.id_shtat as shtat_zam, orez.dolg_rez, orez.tn_sap_rez,                                                                                         \
                                      orez.fio_rez, orez.zex_rez, orez.shifr_rez,                                                                                                         \
                                      orez.risk, orez.risk_prich                                                                                                                          \
                               from ocenka oc                                                                                                                                             \
                               left join ocenka_rez orez                                                                                                                                  \
                               on oc.tn=orez.tn where oc.god="+IntToStr(god)+" and orez.god="+IntToStr(god)+") tab_oc                                                                     \
                              left join                                                                                                                                                   \
                              (select tn_sap, zex, nzex as zam_nzex, name_dolg_ru as dolg,                                                                                                \
                                      trunc(months_between(sysdate, dat_r)/12) as z_vz                                                                                                    \
                                from sap_osn_sved) tab_sap                                                                                                                                \
                              on tab_sap.tn_sap=tab_oc.tn)                                                                                                                                \
                     ) zamesh                  \                                                                                                                                          \
                       on dolgn.n_shtat=zamesh.shtat_zam  and zam_tn!=nvl(rukov.tn,0)                                                                                                                 \               \
                   ) order by shifr_zex, n_shtat, short, z_tn) vse ";

   if (otchet==0)
    {
      StatusBar1->SimpleText=" Идет формирование плана преемственности по предприятию в Excel...";
    }
  else
    {
      Sql+=" where kod_zex="+QuotedStr(otchet);
      StatusBar1->SimpleText=" Идет формирование плана преемственности по подразделению в Excel...";                                                                                                                                                \
    }

                 //    where datn1 is not null          where rukov.tn=49012708 or datn1 is not null
                 //decode(rukov.gotov,1,'низкая',2,'средняя',3,'высокая','') as r_gotov,                                            
  StatusBar1->SimpleText=" Идет формирование плана преемственности в Excel...";                                                                                                                                                \
  //ShowMessage(Sql);

  DM->qObnovlenie->Close();
  DM->qObnovlenie->SQL->Clear();
  DM->qObnovlenie->SQL->Add(Sql);
  try
    {
      DM->qObnovlenie->Open();
    }
  catch (Exception &E)
    {
      Application->MessageBox(("Возникла ошибка при выборке данных из картотеки по оценке персонала и кадров\n(OCENKA, SAP_OSN_SVED, SAP_PEREVOD)"+E.Message).c_str(),"Ошибка",
                              MB_OK+MB_ICONERROR);

      StatusBar1->SimpleText ="Отчетный период: "+IntToStr(god)+" год";
      Abort();
    }

  if (DM->qObnovlenie->RecordCount==0)
    {
      Application->MessageBox(("Нет данных для формирования отчета за "+IntToStr(god)+" год").c_str(),"Предупреждение",
                              MB_OK+MB_ICONINFORMATION);

      StatusBar1->SimpleText ="Отчетный период: "+IntToStr(god)+" год";
      Abort();
    }

  Cursor = crHourGlass;
  ProgressBar->Position = 0;
  ProgressBar->Visible = true;
  ProgressBar->Max=DM->qObnovlenie->RecordCount;

  // инициализируем Excel, открываем этот шаблон
  try
    {
      AppEx=CreateOleObject("Excel.Application");
    }
  catch (...)
    {
      Application->MessageBox("Невозможно открыть Microsoft Excel!"
                              " Возможно это приложение на компьютере не установлено.","Ошибка",MB_OK+MB_ICONERROR);
      StatusBar1->SimpleText ="Отчетный период: "+IntToStr(god)+" год";
      ProgressBar->Visible = false;
      Cursor = crDefault;
    }

  //Если возникает ошибка во время формирования отчета
  try
    {
      try
        {
          AppEx.OlePropertySet("AskToUpdateLinks",false);
          AppEx.OlePropertySet("DisplayAlerts",false);

          //Создание папки, если ее не существует
          ForceDirectories(WorkPath);

          if (otchet==0)
            {
              //Копируем шаблон файла в Мои документы
              CopyFile((Path+"\\RTF\\pp.xlsx").c_str(), (WorkPath+"\\План преемственности.xlsx").c_str(), false);
              sFile = WorkPath+"\\План преемственности.xlsx";
            }
          else
            {
              //Создание папки, если ее не существует
              ForceDirectories(WorkPath+"\\План преемственности");
              //Копируем шаблон файла в Мои документы
              CopyFile((Path+"\\RTF\\pp.xlsx").c_str(), (WorkPath+"\\План преемственности\\"+otchet+" цех.xlsx").c_str(), false);
              sFile = WorkPath+"\\План преемственности\\"+otchet+" цех.xlsx";
            }

          

          AppEx.OlePropertyGet("WorkBooks").OleProcedure("Open",sFile.c_str())    ;  //открываем книгу, указав её имя
          Sh=AppEx.OlePropertyGet("WorkSheets",1);                               //выбираем № активного листа книги
          //Sh=AppEx.OlePropertyGet("WorkSheets","Расчет");                      //выбираем лист по наименованию
        }
      catch(...)
        {
          Application->MessageBox("Ошибка открытия книги Microsoft Excel!","Ошибка",MB_OK+MB_ICONERROR);
          StatusBar1->SimpleText ="Отчетный период: "+IntToStr(god)+" год";
          ProgressBar->Visible = false;
          Cursor = crDefault;
        }
     // AppEx.OlePropertySet("Visible",true);

      i=1;
      n=6;
      int num=0;

      //Вывод данных в шаблон
      Variant Massiv;
      Massiv = VarArrayCreate(OPENARRAY(int,(0,99)),varVariant); //массив на 100 элементов


      //Определение колоноки с какой начинать заполнение для дат замещения
      //Январь
      yanvar=47;

      //Февраль
      if (DM->qObnovlenie->FieldByName("max_kol_kolonok1")->AsString>0) fevral = yanvar + DM->qObnovlenie->FieldByName("max_kol_kolonok1")->AsString*2;
      else fevral = yanvar+2;

      //Март
      if (DM->qObnovlenie->FieldByName("max_kol_kolonok2")->AsString>0) mart = fevral + DM->qObnovlenie->FieldByName("max_kol_kolonok2")->AsString*2;
      else mart = fevral+2;

      //Апрель
      if (DM->qObnovlenie->FieldByName("max_kol_kolonok3")->AsString>0) aprel = mart + DM->qObnovlenie->FieldByName("max_kol_kolonok3")->AsString*2;
      else aprel = mart+2;

      //Май
      if (DM->qObnovlenie->FieldByName("max_kol_kolonok4")->AsString>0) mai = aprel + DM->qObnovlenie->FieldByName("max_kol_kolonok4")->AsString*2;
      else mai = mart+2;

      //Июнь
      if (DM->qObnovlenie->FieldByName("max_kol_kolonok5")->AsString>0) iyn = mai + DM->qObnovlenie->FieldByName("max_kol_kolonok5")->AsString*2;
      iyn = mai+2;

      //Июль
      if (DM->qObnovlenie->FieldByName("max_kol_kolonok6")->AsString>0) iyl = iyn + DM->qObnovlenie->FieldByName("max_kol_kolonok6")->AsString*2;
      else iyl = iyn+2;

      //Август
      if (DM->qObnovlenie->FieldByName("max_kol_kolonok7")->AsString>0) avgust = iyl + DM->qObnovlenie->FieldByName("max_kol_kolonok7")->AsString*2;
      else avgust = iyl+2;

      //Сентябрь
      if (DM->qObnovlenie->FieldByName("max_kol_kolonok8")->AsString>0) sentyabr = avgust + DM->qObnovlenie->FieldByName("max_kol_kolonok8")->AsString*2;
      else sentyabr = avgust+2;

      //Октябрь
      if (DM->qObnovlenie->FieldByName("max_kol_kolonok9")->AsString>0) oktyabr = sentyabr + DM->qObnovlenie->FieldByName("max_kol_kolonok9")->AsString*2;
      else oktyabr = sentyabr+2;

      //Ноябрь
      if (DM->qObnovlenie->FieldByName("max_kol_kolonok10")->AsString>0) noyabr = oktyabr + DM->qObnovlenie->FieldByName("max_kol_kolonok10")->AsString*2;
      else noyabr = oktyabr+2;

      //Декабрь
      if (DM->qObnovlenie->FieldByName("max_kol_kolonok11")->AsString>0) dekabr = noyabr + DM->qObnovlenie->FieldByName("max_kol_kolonok11")->AsString*2;
      else dekabr = noyabr+2;


      //Заголовок**************************************************************


      //Строка 3 "Исполнение обязанностей (замещение)"
      //Определяем последнюю колонку
      for (int i=1; i<13; i++)
        {
          if (DM->qObnovlenie->FieldByName("max_kol_kolonok"+IntToStr(i))->AsInteger==0) pos_kol = pos_kol+2;
          else pos_kol = pos_kol+DM->qObnovlenie->FieldByName("max_kol_kolonok"+IntToStr(i))->AsInteger*2;
        }
      pos_kol = pos_kol+48;
      //Объединение
      Sh.OlePropertyGet("Range", "AV3",Sh.OlePropertyGet("Cells",3,pos_kol)).OleProcedure("Merge");
      //Значение
      Sh.OlePropertyGet("Range", "AV3",Sh.OlePropertyGet("Cells",3,pos_kol)).OlePropertySet("Value","Исполнение обязанностей (замещение)");
      //Выравнивание
      Sh.OlePropertyGet("Range","AV3",Sh.OlePropertyGet("Cells",3,pos_kol)).OlePropertySet("HorizontalAlignment", xlCenter); //выровнять по гор.
      Sh.OlePropertyGet("Range","AV3",Sh.OlePropertyGet("Cells",3,pos_kol)).OlePropertySet("VerticalAlignment", xlCenter); //выровнять по верт.
      pos_kol=48;


      //Строка 4 "Месяцы замещения"
      for (int i=1; i<13; i++)
        {
          //Определяем количество колонок в месяце
          if (DM->qObnovlenie->FieldByName("max_kol_kolonok"+IntToStr(i))->AsInteger==0)
             {
               //Объединение месяцев
               Sh.OlePropertyGet("Range", Sh.OlePropertyGet("Cells",4,pos_kol),Sh.OlePropertyGet("Cells",4,pos_kol+1)).OleProcedure("Merge");
               //Зачение
               Sh.OlePropertyGet("Range", Sh.OlePropertyGet("Cells",4,pos_kol),Sh.OlePropertyGet("Cells",4,pos_kol+1)).OlePropertySet("Value",(Mes[i]).c_str());
               //Выравнивание
               Sh.OlePropertyGet("Range", Sh.OlePropertyGet("Cells",4,pos_kol),Sh.OlePropertyGet("Cells",4,pos_kol+1)).OlePropertySet("HorizontalAlignment", xlCenter); //выровнять по гор.
               Sh.OlePropertyGet("Range", Sh.OlePropertyGet("Cells",4,pos_kol),Sh.OlePropertyGet("Cells",4,pos_kol+1)).OlePropertySet("VerticalAlignment", xlCenter); //выровнять по верт.
               pos_kol = pos_kol+2;
             }
           else
             {
               //Объединение
               Sh.OlePropertyGet("Range", Sh.OlePropertyGet("Cells",4,pos_kol),Sh.OlePropertyGet("Cells",4,(pos_kol+DM->qObnovlenie->FieldByName("max_kol_kolonok"+IntToStr(i))->AsInteger*2-1))).OleProcedure("Merge");
               //Зачение
               Sh.OlePropertyGet("Range", Sh.OlePropertyGet("Cells",4,pos_kol),Sh.OlePropertyGet("Cells",4,(pos_kol+DM->qObnovlenie->FieldByName("max_kol_kolonok"+IntToStr(i))->AsInteger*2-1))).OlePropertySet("Value",(Mes[i]).c_str());
               //Выравнивание
               Sh.OlePropertyGet("Range", Sh.OlePropertyGet("Cells",4,pos_kol),Sh.OlePropertyGet("Cells",4,(pos_kol+DM->qObnovlenie->FieldByName("max_kol_kolonok"+IntToStr(i))->AsInteger*2-1))).OlePropertySet("HorizontalAlignment", xlCenter); //выровнять по гор.
               Sh.OlePropertyGet("Range", Sh.OlePropertyGet("Cells",4,pos_kol),Sh.OlePropertyGet("Cells",4,(pos_kol+DM->qObnovlenie->FieldByName("max_kol_kolonok"+IntToStr(i))->AsInteger*2-1))).OlePropertySet("VerticalAlignment", xlCenter); //выровнять по верт.
               pos_kol = pos_kol+DM->qObnovlenie->FieldByName("max_kol_kolonok"+IntToStr(i))->AsInteger*2;
             }
        }


      //Строка 5 "с/по"
      int j=48;
      while  (j<pos_kol)
        {
          //Значение
          if (j%2==0)  Sh.OlePropertyGet("Range", Sh.OlePropertyGet("Cells",5,j),Sh.OlePropertyGet("Cells",5,j)).OlePropertySet("Value","по");
          else Sh.OlePropertyGet("Range", Sh.OlePropertyGet("Cells",5,j),Sh.OlePropertyGet("Cells",5,j)).OlePropertySet("Value","с");
          //Выравнивание
          Sh.OlePropertyGet("Range","AV5",Sh.OlePropertyGet("Cells",5,j)).OlePropertySet("HorizontalAlignment", xlCenter); //выровнять по гор.
          Sh.OlePropertyGet("Range","AV5",Sh.OlePropertyGet("Cells",5,j)).OlePropertySet("VerticalAlignment", xlCenter); //выровнять по верт.
          j++;
        }

      //Объединение и вывод данных в последнюю ячейку
      Sh.OlePropertyGet("Range", Sh.OlePropertyGet("Cells",4,pos_kol),Sh.OlePropertyGet("Cells",5,pos_kol)).OleProcedure("Merge");
      Sh.OlePropertyGet("Range", Sh.OlePropertyGet("Cells",4,pos_kol),Sh.OlePropertyGet("Cells",5,pos_kol)).OlePropertySet("Value","Кол-во дней замещения");
      //Выравнивание
      Sh.OlePropertyGet("Range", Sh.OlePropertyGet("Cells",4,pos_kol),Sh.OlePropertyGet("Cells",5,pos_kol)).OlePropertySet("HorizontalAlignment", xlCenter); //выровнять по гор.
      Sh.OlePropertyGet("Range", Sh.OlePropertyGet("Cells",4,pos_kol),Sh.OlePropertyGet("Cells",5,pos_kol)).OlePropertySet("VerticalAlignment", xlCenter); //выровнять по верт.
      //Перенос текста
      Sh.OlePropertyGet("Range", Sh.OlePropertyGet("Cells",4,pos_kol),Sh.OlePropertyGet("Cells",5,pos_kol)).OlePropertySet("WrapText",true);

      //Рисуем сетку
      Sh.OlePropertyGet("Range","AU3",Sh.OlePropertyGet("Cells",5,pos_kol)).OlePropertyGet("Borders").OlePropertySet("LineStyle", xlContinuous);
      //Закрасить цветом
      Sh.OlePropertyGet("Range","AU3",Sh.OlePropertyGet("Cells",5,pos_kol)).OlePropertyGet("Interior").OlePropertySet("Color",14277081);


      while (!DM->qObnovlenie->Eof)
        {
          tn=DM->qObnovlenie->FieldByName("n_shtat")->AsString;
          tn1=DM->qObnovlenie->FieldByName("n_shtat")->AsString;
          num=1;


         //Определение колоноки с какой начинать заполнение для дат замещения
      //Январь
      yanvar=47;

      //Февраль
      if (DM->qObnovlenie->FieldByName("max_kol_kolonok1")->AsString>0) fevral = yanvar + DM->qObnovlenie->FieldByName("max_kol_kolonok1")->AsString*2;
      else fevral = yanvar+2;

      //Март
      if (DM->qObnovlenie->FieldByName("max_kol_kolonok2")->AsString>0) mart = fevral + DM->qObnovlenie->FieldByName("max_kol_kolonok2")->AsString*2;
      else mart = fevral+2;

      //Апрель
      if (DM->qObnovlenie->FieldByName("max_kol_kolonok3")->AsString>0) aprel = mart + DM->qObnovlenie->FieldByName("max_kol_kolonok3")->AsString*2;
      else aprel = mart+2;

      //Май
      if (DM->qObnovlenie->FieldByName("max_kol_kolonok4")->AsString>0) mai = aprel + DM->qObnovlenie->FieldByName("max_kol_kolonok4")->AsString*2;
      else mai = mart+2;

      //Июнь
      if (DM->qObnovlenie->FieldByName("max_kol_kolonok5")->AsString>0) iyn = mai + DM->qObnovlenie->FieldByName("max_kol_kolonok5")->AsString*2;
      iyn = mai+2;

      //Июль
      if (DM->qObnovlenie->FieldByName("max_kol_kolonok6")->AsString>0) iyl = iyn + DM->qObnovlenie->FieldByName("max_kol_kolonok6")->AsString*2;
      else iyl = iyn+2;

      //Август
      if (DM->qObnovlenie->FieldByName("max_kol_kolonok7")->AsString>0) avgust = iyl + DM->qObnovlenie->FieldByName("max_kol_kolonok7")->AsString*2;
      else avgust = iyl+2;

      //Сентябрь
      if (DM->qObnovlenie->FieldByName("max_kol_kolonok8")->AsString>0) sentyabr = avgust + DM->qObnovlenie->FieldByName("max_kol_kolonok8")->AsString*2;
      else sentyabr = avgust+2;

      //Октябрь
      if (DM->qObnovlenie->FieldByName("max_kol_kolonok9")->AsString>0) oktyabr = sentyabr + DM->qObnovlenie->FieldByName("max_kol_kolonok9")->AsString*2;
      else oktyabr = sentyabr+2;

      //Ноябрь
      if (DM->qObnovlenie->FieldByName("max_kol_kolonok10")->AsString>0) noyabr = oktyabr + DM->qObnovlenie->FieldByName("max_kol_kolonok10")->AsString*2;
      else noyabr = oktyabr+2;

      //Декабрь
      if (DM->qObnovlenie->FieldByName("max_kol_kolonok11")->AsString>0) dekabr = noyabr + DM->qObnovlenie->FieldByName("max_kol_kolonok11")->AsString*2;
      else dekabr = noyabr+2;



          while (!DM->qObnovlenie->Eof && tn==tn1)
            {
              if (num!=1 && !tn.IsEmpty())
                {
                  //Объединение ячеек
                  //Sh.OlePropertyGet("Range",("G"+IntToStr(n)+":G"+IntToStr(n-1)).c_str()).OleProcedure("Merge");
                  Sh.OlePropertyGet("Range",("H"+IntToStr(n)+":H"+IntToStr(n-1)).c_str()).OleProcedure("Merge");
                  Sh.OlePropertyGet("Range",("I"+IntToStr(n)+":I"+IntToStr(n-1)).c_str()).OleProcedure("Merge");
                  Sh.OlePropertyGet("Range",("J"+IntToStr(n)+":J"+IntToStr(n-1)).c_str()).OleProcedure("Merge");
                  Sh.OlePropertyGet("Range",("K"+IntToStr(n)+":K"+IntToStr(n-1)).c_str()).OleProcedure("Merge");
                  Sh.OlePropertyGet("Range",("L"+IntToStr(n)+":L"+IntToStr(n-1)).c_str()).OleProcedure("Merge");
                  Sh.OlePropertyGet("Range",("M"+IntToStr(n)+":M"+IntToStr(n-1)).c_str()).OleProcedure("Merge");
                  Sh.OlePropertyGet("Range",("N"+IntToStr(n)+":N"+IntToStr(n-1)).c_str()).OleProcedure("Merge");
                  Sh.OlePropertyGet("Range",("O"+IntToStr(n)+":O"+IntToStr(n-1)).c_str()).OleProcedure("Merge");
                  Sh.OlePropertyGet("Range",("P"+IntToStr(n)+":P"+IntToStr(n-1)).c_str()).OleProcedure("Merge");
                  Sh.OlePropertyGet("Range",("Q"+IntToStr(n)+":Q"+IntToStr(n-1)).c_str()).OleProcedure("Merge");
                  Sh.OlePropertyGet("Range",("R"+IntToStr(n)+":R"+IntToStr(n-1)).c_str()).OleProcedure("Merge");
                  Sh.OlePropertyGet("Range",("S"+IntToStr(n)+":S"+IntToStr(n-1)).c_str()).OleProcedure("Merge");
                  Sh.OlePropertyGet("Range",("T"+IntToStr(n)+":T"+IntToStr(n-1)).c_str()).OleProcedure("Merge");
                  Sh.OlePropertyGet("Range",("U"+IntToStr(n)+":U"+IntToStr(n-1)).c_str()).OleProcedure("Merge");
                  Sh.OlePropertyGet("Range",("V"+IntToStr(n)+":V"+IntToStr(n-1)).c_str()).OleProcedure("Merge");
                  Sh.OlePropertyGet("Range",("W"+IntToStr(n)+":W"+IntToStr(n-1)).c_str()).OleProcedure("Merge");
                  Sh.OlePropertyGet("Range",("X"+IntToStr(n)+":X"+IntToStr(n-1)).c_str()).OleProcedure("Merge");
                  Sh.OlePropertyGet("Range",("Y"+IntToStr(n)+":Y"+IntToStr(n-1)).c_str()).OleProcedure("Merge");
                  Sh.OlePropertyGet("Range",("Z"+IntToStr(n)+":Z"+IntToStr(n-1)).c_str()).OleProcedure("Merge");
                  Sh.OlePropertyGet("Range",("AA"+IntToStr(n)+":AA"+IntToStr(n-1)).c_str()).OleProcedure("Merge");
                  Sh.OlePropertyGet("Range",("AB"+IntToStr(n)+":AB"+IntToStr(n-1)).c_str()).OleProcedure("Merge");
                  Sh.OlePropertyGet("Range",("AC"+IntToStr(n)+":AC"+IntToStr(n-1)).c_str()).OleProcedure("Merge");
                  Sh.OlePropertyGet("Range",("AD"+IntToStr(n)+":AD"+IntToStr(n-1)).c_str()).OleProcedure("Merge");
                  Sh.OlePropertyGet("Range",("AE"+IntToStr(n)+":AE"+IntToStr(n-1)).c_str()).OleProcedure("Merge");
                  kol_dolg--;
                }

              Massiv.PutElement(DM->qObnovlenie->FieldByName("r_direkt")->AsString.c_str(), 0);
              Massiv.PutElement(DM->qObnovlenie->FieldByName("n_shtat")->AsString.c_str(), 1);
              Massiv.PutElement(DM->qObnovlenie->FieldByName("kod_zex")->AsString.c_str(), 2);
              Massiv.PutElement(DM->qObnovlenie->FieldByName("nzex")->AsString.c_str(), 3);
              Massiv.PutElement(DM->qObnovlenie->FieldByName("nuch")->AsString.c_str(), 4);
              Massiv.PutElement(DM->qObnovlenie->FieldByName("r_tn")->AsString.c_str(), 5);
              Massiv.PutElement(DM->qObnovlenie->FieldByName("r_fio")->AsString.c_str(), 6);
              Massiv.PutElement(DM->qObnovlenie->FieldByName("r_direkt")->AsString.c_str(), 7);
              Massiv.PutElement(DM->qObnovlenie->FieldByName("kod_zex")->AsString.c_str(), 8);
              Massiv.PutElement(DM->qObnovlenie->FieldByName("nzex")->AsString.c_str(), 9);
              Massiv.PutElement(DM->qObnovlenie->FieldByName("nuch")->AsString.c_str(), 10);
              Massiv.PutElement(kol_dolg, 11);
              Massiv.PutElement(DM->qObnovlenie->FieldByName("n_shtat")->AsString.c_str(), 12);
              Massiv.PutElement(DM->qObnovlenie->FieldByName("short")->AsString.c_str(), 13);
              Massiv.PutElement(DM->qObnovlenie->FieldByName("dolg")->AsString.c_str(), 14);
              Massiv.PutElement(DM->qObnovlenie->FieldByName("r_tn")->AsString.c_str(), 15);
              Massiv.PutElement(DM->qObnovlenie->FieldByName("r_fio")->AsString.c_str(), 16);
              Massiv.PutElement(DM->qObnovlenie->FieldByName("r_preem")->AsString.c_str(), 17);
              Massiv.PutElement(DM->qObnovlenie->FieldByName("r_vz")->AsString.c_str(), 18);

              if (DM->qObnovlenie->FieldByName("r_kpe1")->AsString=="-" || DM->qObnovlenie->FieldByName("r_kpe1")->AsString=="") Massiv.PutElement(DM->qObnovlenie->FieldByName("r_kpe1")->AsString, 19);
              else Massiv.PutElement(DM->qObnovlenie->FieldByName("r_kpe1")->AsFloat, 19);
              if (DM->qObnovlenie->FieldByName("r_kpe2")->AsString=="-" || DM->qObnovlenie->FieldByName("r_kpe2")->AsString=="") Massiv.PutElement(DM->qObnovlenie->FieldByName("r_kpe2")->AsString, 20);
              else Massiv.PutElement(DM->qObnovlenie->FieldByName("r_kpe2")->AsFloat, 20);
              if (DM->qObnovlenie->FieldByName("r_kpe3")->AsString=="-" || DM->qObnovlenie->FieldByName("r_kpe3")->AsString=="") Massiv.PutElement(DM->qObnovlenie->FieldByName("r_kpe3")->AsString, 21);
              else Massiv.PutElement(DM->qObnovlenie->FieldByName("r_kpe3")->AsFloat, 21);
              if (DM->qObnovlenie->FieldByName("r_kpe4")->AsString=="-" || DM->qObnovlenie->FieldByName("r_kpe4")->AsString=="") Massiv.PutElement(DM->qObnovlenie->FieldByName("r_kpe4")->AsString, 22);
              else Massiv.PutElement(DM->qObnovlenie->FieldByName("r_kpe4")->AsFloat, 22);
              if (DM->qObnovlenie->FieldByName("r_kpe")->AsString=="-" || DM->qObnovlenie->FieldByName("r_kpe")->AsString=="") Massiv.PutElement(DM->qObnovlenie->FieldByName("r_kpe")->AsString, 23);
              else Massiv.PutElement(DM->qObnovlenie->FieldByName("r_kpe")->AsFloat, 23);


               //окрашивание ячеек
              if (DM->qObnovlenie->FieldByName("r_risk")->AsString=="низкий") Sh.OlePropertyGet("Range",("Y"+IntToStr(n)).c_str()).OlePropertyGet("Interior").OlePropertySet("Color",9437071);
              else if (DM->qObnovlenie->FieldByName("r_risk")->AsString=="средний") Sh.OlePropertyGet("Range",("Y"+IntToStr(n)).c_str()).OlePropertyGet("Interior").OlePropertySet("Color",10092543);
              else if (DM->qObnovlenie->FieldByName("r_risk")->AsString=="высокий") Sh.OlePropertyGet("Range",("Y"+IntToStr(n)).c_str()).OlePropertyGet("Interior").OlePropertySet("Color",7961087);

              Massiv.PutElement(DM->qObnovlenie->FieldByName("r_risk")->AsString.c_str(), 24);
              Massiv.PutElement(DM->qObnovlenie->FieldByName("r_risk_prich")->AsString.c_str(), 25);
              Massiv.PutElement(DM->qObnovlenie->FieldByName("r_ocenka")->AsString.c_str(), 26);
              Massiv.PutElement(DM->qObnovlenie->FieldByName("r_rezerv")->AsString.c_str(), 27);
              Massiv.PutElement(DM->qObnovlenie->FieldByName("r_vz_pens")->AsString.c_str(), 28);
              Massiv.PutElement(DM->qObnovlenie->FieldByName("r_krd")->AsString.c_str(), 29);

              //окрашивание ячеек
              if (DM->qObnovlenie->FieldByName("r_gotov")->AsString=="низкая") Sh.OlePropertyGet("Range",("AE"+IntToStr(n)).c_str()).OlePropertyGet("Interior").OlePropertySet("Color",7961087);
              else if (DM->qObnovlenie->FieldByName("r_gotov")->AsString=="средняя") Sh.OlePropertyGet("Range",("AE"+IntToStr(n)).c_str()).OlePropertyGet("Interior").OlePropertySet("Color",10092543);
              else if (DM->qObnovlenie->FieldByName("r_gotov")->AsString=="высокая") Sh.OlePropertyGet("Range",("AE"+IntToStr(n)).c_str()).OlePropertyGet("Interior").OlePropertySet("Color",9437071);

              Massiv.PutElement(DM->qObnovlenie->FieldByName("r_gotov")->AsString.c_str(), 30);

              if (DM->qObnovlenie->FieldByName("z_tn")->AsString.IsEmpty()) Massiv.PutElement("", 31);
              else Massiv.PutElement(num, 31);
              Massiv.PutElement(DM->qObnovlenie->FieldByName("z_tn")->AsString.c_str(), 32);
              Massiv.PutElement(DM->qObnovlenie->FieldByName("z_fio")->AsString.c_str(), 33);
              Massiv.PutElement(DM->qObnovlenie->FieldByName("z_dolg")->AsString.c_str(), 34);
              Massiv.PutElement(DM->qObnovlenie->FieldByName("z_zex")->AsString.c_str(), 35);
              Massiv.PutElement(DM->qObnovlenie->FieldByName("z_nzex")->AsString.c_str(), 36);
              Massiv.PutElement(DM->qObnovlenie->FieldByName("z_vz")->AsString.c_str(), 37);

              if (DM->qObnovlenie->FieldByName("z_kpe1")->AsString=="-" || DM->qObnovlenie->FieldByName("z_kpe1")->AsString=="") Massiv.PutElement(DM->qObnovlenie->FieldByName("z_kpe1")->AsString, 38);
              else Massiv.PutElement(DM->qObnovlenie->FieldByName("z_kpe1")->AsFloat, 38);
              if (DM->qObnovlenie->FieldByName("z_kpe2")->AsString=="-" || DM->qObnovlenie->FieldByName("z_kpe2")->AsString=="") Massiv.PutElement(DM->qObnovlenie->FieldByName("z_kpe2")->AsString, 39);
              else Massiv.PutElement(DM->qObnovlenie->FieldByName("z_kpe2")->AsFloat, 39);
              if (DM->qObnovlenie->FieldByName("z_kpe3")->AsString=="-" || DM->qObnovlenie->FieldByName("z_kpe3")->AsString=="") Massiv.PutElement(DM->qObnovlenie->FieldByName("z_kpe3")->AsString, 40);
              else Massiv.PutElement(DM->qObnovlenie->FieldByName("z_kpe3")->AsFloat, 40);
              if (DM->qObnovlenie->FieldByName("z_kpe4")->AsString=="-" || DM->qObnovlenie->FieldByName("z_kpe4")->AsString=="") Massiv.PutElement(DM->qObnovlenie->FieldByName("z_kpe4")->AsString, 41);
              else Massiv.PutElement(DM->qObnovlenie->FieldByName("z_kpe4")->AsFloat, 41);
              if (DM->qObnovlenie->FieldByName("z_kpe")->AsString=="-" || DM->qObnovlenie->FieldByName("z_kpe")->AsString=="") Massiv.PutElement(DM->qObnovlenie->FieldByName("z_kpe")->AsString, 42);
              else Massiv.PutElement(DM->qObnovlenie->FieldByName("z_kpe")->AsFloat, 42);

              Massiv.PutElement(DM->qObnovlenie->FieldByName("z_rezerv")->AsString.c_str(), 43);
              Massiv.PutElement(DM->qObnovlenie->FieldByName("z_ocenka")->AsString.c_str(), 44);
              Massiv.PutElement(DM->qObnovlenie->FieldByName("z_vz_pens")->AsString.c_str(), 45);

              //окрашивание ячеек
              if (DM->qObnovlenie->FieldByName("z_gotov")->AsString=="низкая") Sh.OlePropertyGet("Range",("AU"+IntToStr(n)).c_str()).OlePropertyGet("Interior").OlePropertySet("Color",7961087);
              else if (DM->qObnovlenie->FieldByName("z_gotov")->AsString=="средняя") Sh.OlePropertyGet("Range",("AU"+IntToStr(n)).c_str()).OlePropertyGet("Interior").OlePropertySet("Color",10092543);
              else if (DM->qObnovlenie->FieldByName("z_gotov")->AsString=="высокая") Sh.OlePropertyGet("Range",("AU"+IntToStr(n)).c_str()).OlePropertyGet("Interior").OlePropertySet("Color",9437071);

              Massiv.PutElement(DM->qObnovlenie->FieldByName("z_gotov")->AsString.c_str(), 46);

              //Вывод периодов замещения
              if (!DM->qObnovlenie->FieldByName("m1")->AsString.IsEmpty())
                {
                  for (int i=1; i<13; i++)
                    {
                      //Январь
                      if (DM->qObnovlenie->FieldByName("m"+IntToStr(i))->AsString=="01")
                        {
                          Massiv.PutElement(DM->qObnovlenie->FieldByName("datn"+IntToStr(i))->AsString.c_str(), yanvar++);
                          Massiv.PutElement(DM->qObnovlenie->FieldByName("datk"+IntToStr(i))->AsString.c_str(), yanvar++);
                        }

                      //Февраль
                      if (DM->qObnovlenie->FieldByName("m"+IntToStr(i))->AsString=="02")
                        {
                          Massiv.PutElement(DM->qObnovlenie->FieldByName("datn"+IntToStr(i))->AsString.c_str(), fevral++);
                          Massiv.PutElement(DM->qObnovlenie->FieldByName("datk"+IntToStr(i))->AsString.c_str(), fevral++);
                        }

                      //Март
                      if (DM->qObnovlenie->FieldByName("m"+IntToStr(i))->AsString=="03")
                        {
                          Massiv.PutElement(DM->qObnovlenie->FieldByName("datn"+IntToStr(i))->AsString.c_str(), mart++);
                          Massiv.PutElement(DM->qObnovlenie->FieldByName("datk"+IntToStr(i))->AsString.c_str(), mart++);
                        }

                      //Апрель
                      if (DM->qObnovlenie->FieldByName("m"+IntToStr(i))->AsString=="04")
                        {
                          Massiv.PutElement(DM->qObnovlenie->FieldByName("datn"+IntToStr(i))->AsString.c_str(), aprel++);
                          Massiv.PutElement(DM->qObnovlenie->FieldByName("datk"+IntToStr(i))->AsString.c_str(), aprel++);
                        }

                      //Май
                      if (DM->qObnovlenie->FieldByName("m"+IntToStr(i))->AsString=="05")
                        {
                          Massiv.PutElement(DM->qObnovlenie->FieldByName("datn"+IntToStr(i))->AsString.c_str(), mai++);
                          Massiv.PutElement(DM->qObnovlenie->FieldByName("datk"+IntToStr(i))->AsString.c_str(), mai++);
                        }

                      //Июнь
                      if (DM->qObnovlenie->FieldByName("m"+IntToStr(i))->AsString=="06")
                        {
                          Massiv.PutElement(DM->qObnovlenie->FieldByName("datn"+IntToStr(i))->AsString.c_str(), iyn++);
                          Massiv.PutElement(DM->qObnovlenie->FieldByName("datk"+IntToStr(i))->AsString.c_str(), iyn++);
                        }

                      //Июль
                      if (DM->qObnovlenie->FieldByName("m"+IntToStr(i))->AsString=="07")
                        {
                          Massiv.PutElement(DM->qObnovlenie->FieldByName("datn"+IntToStr(i))->AsString.c_str(), iyl++);
                          Massiv.PutElement(DM->qObnovlenie->FieldByName("datk"+IntToStr(i))->AsString.c_str(), iyl++);
                        }

                      //Август
                      if (DM->qObnovlenie->FieldByName("m"+IntToStr(i))->AsString=="08")
                        {
                          Massiv.PutElement(DM->qObnovlenie->FieldByName("datn"+IntToStr(i))->AsString.c_str(), avgust++);
                          Massiv.PutElement(DM->qObnovlenie->FieldByName("datk"+IntToStr(i))->AsString.c_str(), avgust++);
                        }

                      //Сентябрь
                      if (DM->qObnovlenie->FieldByName("m"+IntToStr(i))->AsString=="09")
                        {
                          Massiv.PutElement(DM->qObnovlenie->FieldByName("datn"+IntToStr(i))->AsString.c_str(), sentyabr++);
                          Massiv.PutElement(DM->qObnovlenie->FieldByName("datk"+IntToStr(i))->AsString.c_str(), sentyabr++);
                        }

                      //Октябрь
                      if (DM->qObnovlenie->FieldByName("m"+IntToStr(i))->AsString=="10")
                        {
                          Massiv.PutElement(DM->qObnovlenie->FieldByName("datn"+IntToStr(i))->AsString.c_str(), oktyabr++);
                          Massiv.PutElement(DM->qObnovlenie->FieldByName("datk"+IntToStr(i))->AsString.c_str(), oktyabr++);
                        }

                       //Ноябрь
                      if (DM->qObnovlenie->FieldByName("m"+IntToStr(i))->AsString=="11")
                        {
                          Massiv.PutElement(DM->qObnovlenie->FieldByName("datn"+IntToStr(i))->AsString.c_str(), noyabr++);
                          Massiv.PutElement(DM->qObnovlenie->FieldByName("datk"+IntToStr(i))->AsString.c_str(), noyabr++);
                        }

                       //Декабрь
                      if (DM->qObnovlenie->FieldByName("m"+IntToStr(i))->AsString=="12")
                        {
                          Massiv.PutElement(DM->qObnovlenie->FieldByName("datn"+IntToStr(i))->AsString.c_str(), dekabr++);
                          Massiv.PutElement(DM->qObnovlenie->FieldByName("datk"+IntToStr(i))->AsString.c_str(), dekabr++);
                        }
                     }
                 }
               else
                 {
                   //Если нет периодов по замещению
                   for (int i=48; i<pos_kol+1; i++)
                     {
                       Massiv.PutElement("", i);
                     }
                 }


               Massiv.PutElement(DM->qObnovlenie->FieldByName("kol_day_zam")->AsString.c_str(), --pos_kol);
              // Massiv.PutElement("=ЕСЛИ(СЧЁТЗ(AQ"+ IntToStr(n)+":BN"+ IntToStr(n)+");\"+\";\"\")", --dekabr);
              //Massiv.PutElement("=ЕСЛИ(СЧЁТЗ(AQ"+ IntToStr(n)+":"+ Sh.OlePropertyGet("Cells",n,dekabr)+");\"+\";\"\")", --dekabr);

              Sh.OlePropertyGet("Range", ("A" + IntToStr(n)).c_str(),Sh.OlePropertyGet("Cells",n,++pos_kol)).OlePropertySet("Value", Massiv); //строка с данными с ячейки A по ячейку АВ

              i++;
              n++;
              num++;
              kol_dolg++;
              DM->qObnovlenie->Next();

              tn=DM->qObnovlenie->FieldByName("n_shtat")->AsString;
              StatusBar1->SimpleText ="Идет формирование Плана преемственности... "+DM->qObnovlenie->FieldByName("kod_zex")->AsString;
              ProgressBar->Position++;
            }
        }

      //рисуем сетку
      Sh.OlePropertyGet("Range","A6",Sh.OlePropertyGet("Cells",n-1,pos_kol)).OlePropertyGet("Borders").OlePropertySet("LineStyle", xlContinuous);

      // Sh.OleProcedure("SaveAs",vAsCurDir1.c_str());
      AppEx.OlePropertyGet("WorkBooks",1).OleFunction("Save");

      /* //Закрыть открытое приложение Excel
      AppEx.OleProcedure("Quit");
      AppEx = Unassigned;  */

      //Закрыть книгу Excel с шаблоном для вывода информации
      // AppEx.OlePropertyGet("WorkBooks",1).OleProcedure("Close");
      Application->MessageBox("Отчет успешно сформирован!", "Формирование отчета",
                               MB_OK+MB_ICONINFORMATION);
      //AppEx.OlePropertyGet("WorkBooks").OleProcedure("Open",vAsCurDir1.c_str());
      AppEx.OlePropertySet("Visible",true);
      AppEx.OlePropertySet("AskToUpdateLinks",true);
      AppEx.OlePropertySet("DisplayAlerts",true);

      StatusBar1->SimpleText= "Формирование отчета выполнено.";

      Cursor = crDefault;
      ProgressBar->Position=0;
      ProgressBar->Visible = false;
      StatusBar1->SimpleText ="Отчетный период: "+IntToStr(god)+" год";


      if (otchet==0) InsertLog("Формирование 'Плана преемственности' по предприятию за "+IntToStr(god)+" год в Excel успешно завершено");
      else InsertLog("Формирование 'Плана преемственности' за "+IntToStr(god)+" год по подразделению '"+otchet+"' в Excel успешно завершено");

      DM->qLogs->Requery();
    }
  catch(...)
    {
      AppEx.OleProcedure("Quit");
      AppEx = Unassigned;
      Cursor = crDefault;
      ProgressBar->Position=0;
      ProgressBar->Visible = false;

      if (otchet==0) InsertLog("Не выполнено формирование 'Плана преемственности' по предприятию за "+IntToStr(god)+" год в Excel");
      else InsertLog("Не выполнено формирование 'Плана преемственности' за "+IntToStr(god)+" год по подразделению '"+otchet+"' в Excel");

      StatusBar1->SimpleText ="Отчетный период: "+IntToStr(god)+" год";
      Abort();
    }
}
//---------------------------------------------------------------------------



