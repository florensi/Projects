//---------------------------------------------------------------------------
#define NO_WIN32_LEAN_AND_MEAN
#include <stdio.h>

#include <vcl.h>
#pragma hdrstop

#include "uMain.h"
#include "uDM.h"
#include "FuncUser.h"
#include "RepoRTFM.h"
#include "RepoRTFO.h"
#include "uVvod.h"
#include "uSprav.h"
#include "uData.h"

//---------------------------------------------------------------------------
#pragma package(smart_init)
#pragma link "DBGridEh"
#pragma resource "*.dfm"
TMain *Main;
Variant AppEx, Sh;
AnsiString Mes[12]={"", "Январь", "Февраль", "Март", "Апрель", "Май", "Июнь", "Июль", "Август", "Сентябрь", "Октябрь", "Ноябрь"};
//---------------------------------------------------------------------------
__fastcall TMain::TMain(TComponent* Owner)
        : TForm(Owner)
{
}

Variant toExcel(Variant AppEx,const char *Exc, int off, double data)
{
  try {
    AppEx.OlePropertyGet("Range", Exc).OlePropertyGet("Offset", off).OlePropertySet("Value", data);
  } catch(...) { ; }
}
//---------------------------------------------------------------------------
Variant toExcel(Variant AppEx,const char *Exc, int off, String data)
{
  try {
    AppEx.OlePropertyGet("Range", Exc).OlePropertyGet("Offset", off).OlePropertySet("Value", data.c_str());
  } catch(...) { ; }
}
//---------------------------------------------------------------------------
Variant  toExcel(Variant AppEx,const char *Exc, double data)
{
  try {
    AppEx.OlePropertyGet("Range", Exc).OlePropertySet("Value", data);
  } catch(...) { ; }
}

//---------------------------------------------------------------------------
Variant  toExcel(Variant AppEx,const char *Exc, int data)
{
  try {
    AppEx.OlePropertyGet("Range", Exc).OlePropertySet("Value", data);
  } catch(...) { ; }
}

//---------------------------------------------------------------------------
Variant  toExcel(Variant AppEx,const char *Exc, AnsiString data)
{
  try {
    Variant  cur = AppEx.OlePropertyGet("Range", Exc);
    cur.OlePropertySet("Value", data.c_str());
  } catch(...) { ; }
}
//---------------------------------------------------------------------------

//Сортировка цехов и участков по возрастанию
int __fastcall MySort(TStringList* SL, int Index1, int Index2)
{
  AnsiString str1, str2;

  str1 = SL->Strings[Index1];
  str2 = SL->Strings[Index2];

//сравнение для графиков
if ((SL->Strings[Index1]).SubString(1,17)=="mmk-itsvc-hgrf-gr" && (SL->Strings[Index2]).SubString(1,17)=="mmk-itsvc-hgrf-gr")
   {
     if (SL->Strings[Index1].Length()>SL->Strings[Index2].Length())
       return 1;
     else if (SL->Strings[Index1].Length()<SL->Strings[Index2].Length())
       return -1;
     else
       return strcmp((SL->Strings[Index1]).c_str(),(SL->Strings[Index2]).c_str());
   }
//сравнение для других групп доступа   
 else
   {
     return strcmp((SL->Strings[Index1]).c_str(),(SL->Strings[Index2]).c_str());
   }
}
//---------------------------------------------------------------------------
void __fastcall TMain::FormCreate(TObject *Sender)
{
  AnsiString grafik, len, znak, Sql;
  kol_grafik=1;

  //очистка массива со списком графиков
  for (int i=0; i<149; i++)
    {
      //n_grafik[kol_grafik]=NULL;
      n_grafik[i]=NULL;
    }

  /*grafik - добовляемый график,
    len - длинна строки
    znak - символ
    n_grafik - список графиков
    kol_grafik - количество добавляемых графиков*/

  // Получение данных о пользователе из домена
  TStringList *SL_Groups = new TStringList();
  TStringList *SL_Groups2 = new TStringList();




  // Получение данных о пользователе из домена
  // Переменные UserName, DomainName, UserFullName должны быть объявлены как AnsiString
  if (!GetFullUserInfo(UserName, DomainName, UserFullName))
    {
      MessageBox(Handle,"Ошибка получения данных о пользователе","Ошибка",8208);
      Application->Terminate();
      Abort();
    }

 // UserName="nadezhda.iordanova";
//  DomainName="MMK";


  //получение групп доступа из АД
  if (!GetUserGroups(UserName, DomainName, SL_Groups))
    {
      MessageBox(Handle,"Ошибка получения данных о пользователе","Ошибка",8208);
      Application->Terminate();
      Abort();
    }

 // ShowMessage(UserName);
 // ShowMessage(SL_Groups->Text);


  //проверка на доступ к услуге
  if ((SL_Groups->IndexOf("mmk-itsvc-hgrf-admin")<=-1) && (SL_Groups->IndexOf("mmk-itsvc-hgrf")<=-1))
    {
      MessageBox(Handle,"У вас нет прав для работы с\n программой АРМ 'Графики работы'!!!","Права доступа",8208);
      Application->Terminate();
      Abort();
    }

  //Считывание отчетного года из grafr
/*  DM->qObnovlenie->Close();
  DM->qObnovlenie->SQL->Clear();
  DM->qObnovlenie->SQL->Add("select * from grafr");
  try
    {
      DM->qObnovlenie->Open();
    }
  catch(...)
    {
      Application->MessageBox("Невозможно считать отчетный период","Ошибка",
                              MB_OK + MB_ICONERROR);
      Application->Terminate();
      Abort();
    }  */

  Word year, month, day;
  DecodeDate(Date(), year, month, day);

  god = StrToInt(year);//DM->qObnovlenie->FieldByName("god")->AsInteger;
  grafr = StrToInt(year);//DM->qObnovlenie->FieldByName("god")->AsInteger;

  //god = 2015;
  //grafr = 2015;



  //проверка прав
  //права на просмотр графиков
  if (SL_Groups->IndexOf("mmk-itsvc-hgrf-01")>-1)
    {
      N2RaschetVsex->Visible = false;   //Расчет всех графиков
      N5RaschetOdin->Visible = false;   //Расчет одного графика
      N3Redaktirovat->Visible = false;  //Редактирование графика
      Prava=0;                          //Права для редактирования

      //Автоматически загружать все графики имеющиеся в таблице SPGRAFIKI
      // Вывод в ComboBox выбираемых графиков
      DM->qObnovlenie2->Close();
      DM->qObnovlenie2->SQL->Clear();
     // DM->qObnovlenie2->SQL->Add("select distinct ograf from spograf order by ograf");
      //Если у графика только 12 месяц, он в общий список графиков не выводится
      Sql="select distinct ograf \
           from spograf \
           where ograf not in (select ograf \
                               from (select ograf, mes  \
                                     from spgrafiki \
                                     where god="+IntToStr(god)+" group by ograf, mes) \
                               group  by ograf having count(*)=1) order by ograf ";
      DM->qObnovlenie2->SQL->Add(Sql);

      try
        {
          DM->qObnovlenie2->Open();
        }
      catch(...)
        {
          Application->MessageBox("Ошибка доступа к таблице графиков (SPOGRAF)","Ошибка доступа",
                                   MB_OK + MB_ICONERROR);
          Application->Terminate();
          Abort();
        }

      ComboBox1->Items->Clear();
      while(!DM->qObnovlenie2->Eof)
        {
          ComboBox1->Items->Add(DM->qObnovlenie2->FieldByName("ograf")->AsString);
          DM->qObnovlenie2->Next();
        }
      ComboBox1->ItemIndex = -1;
    }
  else
    {
      Application->MessageBox("У вас нет прав на просмотр графиков!!!","Права доступа",
                               MB_OK+MB_ICONERROR);
      Application->Terminate();
      Abort();
    }

  //права на редактирование графиков
  if (SL_Groups->IndexOf("mmk-itsvc-hgrf-02")>-1)
    {
      
      N5RaschetOdin->Visible = true;   //Расчет одного графика
      N3Redaktirovat->Visible = true;  //Редактирование графика
      redakt=1;                        //Доступ к редактированию
      Prava=1;                         //Права для редактирования
      int ind=1; //индекс найденой записи

      //поиск группы доступа со всеми цехами
      if (SL_Groups->IndexOf("mmk-itsvc-hgrf-gr-all")>-1)
        {
          N2RaschetVsex->Visible = true;   //Расчет всех графиков

          //Возвращает все существующие группы с цехами в АД
          GetGroups("OU=ITServices,OU=MMK", "mmk-itsvc-hgrf-gr", SL_Groups2);

          //сортировка групп доступа
          SL_Groups2->CustomSort(MySort);
          SL_Groups2->Find("mmk-itsvc-hgrf-gr%",ind);

           while (ind+1<=SL_Groups2->Count)
             {
                  // ShowMessage(SL_Groups2->Text);
               if (SL_Groups2->Strings[ind].Pos("mmk-itsvc-hgrf-gr-") && SL_Groups2->Strings[ind].SubString(20,255)!="all")
                 {
                   n_grafik[kol_grafik] = SL_Groups2->Strings[ind].SubString(19,255);
                   kol_grafik++;
                 }
              ind++;
            }
        }
      else
        {
          //Список некоторых графиков для редактирования (разрешены те группы с № графика, которые есть в АД у пользователя)

          N2RaschetVsex->Visible = false;   //Расчет всех графиков

          int ind=1; //индекс найденой записи

          //сортировка групп доступа
          SL_Groups->CustomSort(MySort);
          //поиск группы доступа с цехом
          SL_Groups->Find("mmk-itsvc-hgrf-gr%",ind);

          //если группа доступа с цехом найдена
          if (ind!=-1)
            {
              ind=ind-1;
              while (ind<SL_Groups->Count)
                {
                  if (SL_Groups->Strings[ind].Pos("mmk-itsvc-hgrf-gr-"))
                    {
                      n_grafik[kol_grafik] = StrToInt(SL_Groups->Strings[ind].SubString(19,255));
                      kol_grafik++;

                    }
                  ind++;
                }
            }
        }
    }
  else
    {
      redakt=0;   //Запрет на редактирование
      Prava=0;    //Права для редактирования
    }

  //права на выгрузку графиков в НСИ
  if (SL_Groups->IndexOf("mmk-itsvc-hgrf-03")>-1)
    {
      //N5V_UIT->Visible = true;        //Выгрузка данных в УИТ
      N5V_UIT->Visible = false;
    }
  else
    {
      N5V_UIT->Visible = false;       //Выгрузка данных в УИТ
    }

  delete SL_Groups;


  DBGridEh1->Enabled = false;
  StatusBar1->SimpleText="Отчетный период:  "+IntToStr(god)+" год";
  
  //Праздничные дни
  DM->qPrazdDni->Close();
  DM->qPrazdDni->Parameters->ParamByName("pgod")->Value = god;
  DM->qPrazdDni->Open();

  //Предпраздничные дни
  DM->qPrdPrazdDni->Close();
  DM->qPrdPrazdDni->Parameters->ParamByName("pgod")->Value = god;
  DM->qPrdPrazdDni->Open();


  //Определение даты перехода на летнее/зимнее время
  TDateTime data;
//  Word year, month, day;

  // дата в марте
  data = DateToStr(EncodeDateMonthWeek(god,3,4,6));
  DecodeDate(data, year, month, day);
  day_mart = day;
  //для 40 и 90 графика, первой смены, дата в марте
  if (day_mart==31)
    {
      mes_mart2=4;
      day_mart2=1;
    }
  else
    {
      mes_mart2=3;
      day_mart2=day_mart+1;
    }

  //дата в октябре
  data = DateToStr(EncodeDateMonthWeek(god,10,4,6));
  DecodeDate(data, year, month, day);
  day_oktyabr = day;
  //для 40 и 90 графика, первой смены, дата в октябре
  if (day_oktyabr==31)
    {
      mes_oktyabr2=11;
      day_oktyabr2=1;
    }
  else
    {
      mes_oktyabr2=10;
      day_oktyabr2=day_oktyabr+1;
    }

  //Получение пути к папке "Мои документы", "Temp"
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

  WorkPath = DocPath + "\\Расчет графиков";
  Path = GetCurrentDir();
  FindWordPath();

  // Создание CGauge на StatusBar
  ProgressBar = new TProgressBar (StatusBar1);
  ProgressBar->Parent = StatusBar1;
  ProgressBar->Position = 0;
  ProgressBar->Left = StatusBar1->Panels->Items[0]->Width*19.3 + 33;
  ProgressBar->Top = StatusBar1->Height/6;
  ProgressBar->Height = StatusBar1->Height-3;
  ProgressBar->Visible = false;

}
//---------------------------------------------------------------------------

// Возвращает путь на папку "Мои документы"
bool __fastcall TMain::GetMyDocumentsDir(AnsiString &FolderPath)
{
  char f[MAX_PATH];

  if (SUCCEEDED(SHGetFolderPath(NULL, CSIDL_PERSONAL|CSIDL_FLAG_CREATE, NULL, SHGFP_TYPE_CURRENT, f))) {
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

  if (GetTempPath(MAX_PATH, f)) {
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

//Определяет день недели 6 - суббота, 1- воскресенье, 2 - понедельник
int __fastcall TMain::DayWeek(int d, int m, int y)
{
  int week = DayOfWeek(FormatDateTime("dd.mm.yyyy",StrToDate(IntToStr(d)+"."+ IntToStr(m)+"."+IntToStr(y))));
  return(week);
}
//---------------------------------------------------------------------------

//Определение праздничного дня
bool __fastcall TMain::PrazdDni(int d, int m)
{
  Variant locvalues[] = {d,m};
  if (DM->qPrazdDni->Locate("den;mes", VarArrayOf(locvalues, 1),
                                           SearchOptions << loCaseInsensitive))
    {
      return(true);
    }
  else
    {
      return(false);
    }
}
//---------------------------------------------------------------------------

//Определение предпраздничного дня
bool __fastcall TMain::PrdPrazdDni(int d, int m)
{
  Variant locvalues[] = {(d<10 ? "0"+IntToStr(d)  : IntToStr(d)),(m<10? "0"+IntToStr(m) : IntToStr(m))};
  if (DM->qPrdPrazdDni->Locate("den;mes", VarArrayOf(locvalues, 1),
                                           SearchOptions << loCaseInsensitive))
    {
      return(true);
    }
  else
    {
      return(false);
    }
}
//---------------------------------------------------------------------------

//Определение праздничного дня на субботу и воскресенье
bool __fastcall TMain::PrazdDniVihodnue(int d, int m, int y)
{
  Variant locvalues[] = {(d<10 ? "0"+IntToStr(d)  : IntToStr(d)),(m<10? "0"+IntToStr(m) : IntToStr(m)),y};
  if (DM->qPrazdDniVihodnue->Locate("den;mes;god", VarArrayOf(locvalues, 2),
                                           SearchOptions << loCaseInsensitive))
    {
      return(true);
    }
  else
    {
      return(false);
    }
}

//---------------------------------------------------------------------------
//Выбор графика
void __fastcall TMain::ComboBox1Click(TObject *Sender)
{
  AnsiString Sql, vihod;

  if (ComboBox1->Text.IsEmpty())
    {
      Application->MessageBox("Выберите необходимый график!!!","Предупреждение",
                              MB_OK+MB_ICONINFORMATION);
      ComboBox1->SetFocus();
      DBGridEh1->Enabled = false;
      Abort();
    }
  else
    {
      DBGridEh1->Enabled = true;
    }

  DBGridEh1->DataSource= DM->dsGrafik;


  //Выбор общих данных по графику
  DM->qOgraf->Close();
  DM->qOgraf->Parameters->ParamByName("pograf")->Value = ComboBox1->Text;
  DM->qOgraf->Open();


  //Режим отображения графика (часы или выходы)
  if (DM->qOgraf->FieldByName("otchet")->AsInteger==1) vihod = "nsm";
  else if (DM->qOgraf->FieldByName("otchet")->AsInteger==2) vihod = "chf";
  else
    {
      Application->MessageBox("Не указан режим отображения отчета в таблице OGRAF","Предупреждение",
                               MB_OK + MB_ICONWARNING);
      Abort();
    }

  //Вывод графика
  DM->qGrafik->Close();
  DM->qGrafik->Parameters->ParamByName("pgod")->Value=god;
  DM->qGrafik->Parameters->ParamByName("pograf")->Value=ComboBox1->Text;
  DM->qGrafik->Open();


  //Видимость колонок:
  // без переработки,ночных, вечерних, праздничных
  if (ComboBox1->Text==11 || ComboBox1->Text==81 ||
      ComboBox1->Text==111 ||
      ComboBox1->Text==650 || ComboBox1->Text==660 ||
      ComboBox1->Text==655 || ComboBox1->Text==1655 ||
      ComboBox1->Text==820 || ComboBox1->Text==830 ||
      ComboBox1->Text==771 || ComboBox1->Text==800 ||
      ComboBox1->Text==480 || ComboBox1->Text==780 ||
      ComboBox1->Text==1011 || ComboBox1->Text==2011 ||
      ComboBox1->Text==3011 || ComboBox1->Text==18 ||
      ComboBox1->Text==1018 || ComboBox1->Text==2018 ||
      ComboBox1->Text==3018
      )
    {
      DBGridEh1->Columns->Items[35]->Visible = false;
      DBGridEh1->Columns->Items[36]->Visible = false;
      DBGridEh1->Columns->Items[37]->Visible = false;
      DBGridEh1->Columns->Items[38]->Visible = false;
      DBGridEh1->AutoFitColWidths = true;
    }
  //только праздничные
  else if (ComboBox1->Text==150 || ComboBox1->Text==30)
    {
      DBGridEh1->Columns->Items[35]->Visible = false;
      DBGridEh1->Columns->Items[36]->Visible = false;
      DBGridEh1->Columns->Items[37]->Visible = false;
      DBGridEh1->Columns->Items[38]->Visible = true;
      DBGridEh1->AutoFitColWidths = true;
    }
  //только вечерние
  else if (ComboBox1->Text==230 || ComboBox1->Text==410 ||
           ComboBox1->Text==315 || ComboBox1->Text==855 ||
           ComboBox1->Text==865 || ComboBox1->Text==880 ||
           ComboBox1->Text==280)
    {
      DBGridEh1->Columns->Items[35]->Visible = false;
      DBGridEh1->Columns->Items[36]->Visible = false;
      DBGridEh1->Columns->Items[38]->Visible = false;
      DBGridEh1->Columns->Items[37]->Visible = true;
      DBGridEh1->AutoFitColWidths = true;
    }
  //только вечерние и праздничные
  else if (ComboBox1->Text==190 || ComboBox1->Text==790||
           ComboBox1->Text==131)
    {
      DBGridEh1->Columns->Items[35]->Visible = false;
      DBGridEh1->Columns->Items[36]->Visible = false;
      DBGridEh1->Columns->Items[37]->Visible = true;
      DBGridEh1->Columns->Items[38]->Visible = true;
      DBGridEh1->AutoFitColWidths = true;
    }
  //без вечерних и праздничных
  else if (ComboBox1->Text==85)
    {
      DBGridEh1->Columns->Items[35]->Visible = true;
      DBGridEh1->Columns->Items[36]->Visible = true;
      DBGridEh1->Columns->Items[37]->Visible = false;
      DBGridEh1->Columns->Items[38]->Visible = false;
      DBGridEh1->AutoFitColWidths = true;
    }
  //только вечерние и ночные
  else if (ComboBox1->Text==20 || ComboBox1->Text==1020 ||
           ComboBox1->Text==2020 ||
           ComboBox1->Text==25 || ComboBox1->Text==470 ||
           ComboBox1->Text==775 || ComboBox1->Text==160 ||
           ComboBox1->Text==140)
    {
      DBGridEh1->Columns->Items[35]->Visible = false;
      DBGridEh1->Columns->Items[36]->Visible = true;
      DBGridEh1->Columns->Items[37]->Visible = true;
      DBGridEh1->Columns->Items[38]->Visible = false;
      DBGridEh1->AutoFitColWidths = true;
    }
    //только вечерние и переработка
  else if (ComboBox1->Text==690)
    {
      DBGridEh1->Columns->Items[35]->Visible = true;
      DBGridEh1->Columns->Items[36]->Visible = false;
      DBGridEh1->Columns->Items[38]->Visible = false;
      DBGridEh1->Columns->Items[37]->Visible = true;
      DBGridEh1->AutoFitColWidths = true;
    }
  //без вечерних и ночных
  else if (ComboBox1->Text==630 || ComboBox1->Text==1630)
    {
      DBGridEh1->Columns->Items[35]->Visible = true;
      DBGridEh1->Columns->Items[36]->Visible = false;
      DBGridEh1->Columns->Items[37]->Visible = false;
      DBGridEh1->Columns->Items[38]->Visible = true;
      DBGridEh1->AutoFitColWidths = true;
    }
  //без переработки
  else if (ComboBox1-> Text==120 || ComboBox1->Text==220 || ComboBox1->Text==260 ||
           ComboBox1->Text==90 || ComboBox1->Text==1090 ||
           ComboBox1->Text==133 || ComboBox1->Text==24 || ComboBox1->Text==23 ||
           ComboBox1->Text==170 || ComboBox1->Text==50 || ComboBox1->Text==270 ||
           ComboBox1->Text==250)
    {
      DBGridEh1->Columns->Items[35]->Visible = false;
      DBGridEh1->Columns->Items[36]->Visible = true;
      DBGridEh1->Columns->Items[37]->Visible = true;
      DBGridEh1->Columns->Items[38]->Visible = true;
      DBGridEh1->AutoFitColWidths = true;
    }
  else
    {
      //без ночных
      if (ComboBox1->Text==300 || ComboBox1->Text==1300 ||
          ComboBox1->Text==2300 || ComboBox1->Text==3300 ||
          ComboBox1->Text==4300 ||
          ComboBox1->Text==335 || ComboBox1->Text==400 ||
          ComboBox1->Text==670 || ComboBox1->Text==680 ||
          ComboBox1->Text==850 || ComboBox1->Text==935)
        {
          DBGridEh1->Columns->Items[36]->Visible = false;
          DBGridEh1->AutoFitColWidths = true;
        }
      else
        {
          DBGridEh1->Columns->Items[36]->Visible = true;
          DBGridEh1->AutoFitColWidths = false;
        }
      DBGridEh1->Columns->Items[35]->Visible = true;
      DBGridEh1->Columns->Items[37]->Visible = true;
      DBGridEh1->Columns->Items[38]->Visible = true;
    }

  //Проверка доступен ли график для редактирования
  if (Prava == 1 && god >= grafr)
    {
      for (int i=1; i<=kol_grafik; i++)
        {
          if (n_grafik[i]==StrToInt(ComboBox1->Text))
            {
              redakt = 1;
              break;
            }
          else
            {
              redakt = 0;
            }
        }
    }
  else if (Prava == 1 && god < grafr)
    {
      redakt = 0;
    }

  //Видимость пунктов контекстного меню
  if (DM->qGrafik->RecordCount==0)
    {
      N3Redaktirovat->Enabled=false;

      if (redakt==0)
        {
          N3Redaktirovat->Visible = false;  //Редактировать график
          N5RaschetOdin->Visible = false; //Расчитать график
        }

    }
  else if (redakt==0)
    {
      N3Redaktirovat->Visible = false;  //Редактировать график
      N5RaschetOdin->Visible = false; //Расчитать график
    }
  else
    {
      N3Redaktirovat->Visible = true;
      N5RaschetOdin->Visible = true;
      N3Redaktirovat->Enabled = true;
      N5RaschetOdin->Enabled = true;
    }

  if (!ComboBox1->Text.IsEmpty())
    {
      DBGridEh1->SetFocus();
    }        
}
//---------------------------------------------------------------------------

//Расчет всех графиков на текущий год
void __fastcall TMain::N2RaschetVsexClick(TObject *Sender)
{
  AnsiString Sql;
  int n=1;

  DBGridEh1->DataSource = NULL;

  //Существование праздничных дней за выбранный год
  DM->qObnovlenie2->Close();
  DM->qObnovlenie2->SQL->Clear();
  DM->qObnovlenie2->SQL->Add("select * from sp_prd where god="+IntToStr(god));
  try
    {
      DM->qObnovlenie2->Open();
    }
  catch(...)
    {
      Application->MessageBox("Ошибка выбора данных из справочника праздничных дней","Ошибка",
                               MB_OK+MB_ICONERROR);
    }

  if (DM->qObnovlenie2->RecordCount==0)
    {
      Application->MessageBox(("Нет данных о праздничных днях за "+IntToStr(god)+" год в справочнике праздничных дней(SP_PRD)").c_str(), "Предупреждение",
                               MB_OK+MB_ICONINFORMATION);
      Abort();
    }

  if (Application->MessageBox(("Будет выполнен расчет графиков на "+IntToStr(god)+ " год. \nВсе ранее расчитаные графики будут удалены. Продолжить?").c_str(),
                            "Расчет всех графиков", MB_YESNO+MB_ICONINFORMATION)==ID_NO)
    {
      Abort();
    }


  // все графики имеющиеся в таблице SPGRAFIKI
  Sql="select distinct ograf \
           from spograf order by ograf";
  DM->qObnovlenie->Close();
  DM->qObnovlenie->SQL->Clear();
  DM->qObnovlenie->SQL->Add(Sql);

  try
    {
      DM->qObnovlenie->Open();
    }
  catch(...)
    {
      Application->MessageBox("Ошибка доступа к таблице графиков (SPOGRAF)","Ошибка доступа",
                               MB_OK + MB_ICONERROR);
      Application->Terminate();
      Abort();
    }

  int kol_grafikv=0;
  //DM->qObnovlenie->First();

  //Заполнение массива данными
  for (int n=0; n<DM->qObnovlenie->RecordCount; n++)
    {
      n_grafikv[n]=DM->qObnovlenie->FieldByName("ograf")->AsString;
      DM->qObnovlenie->Next();
      kol_grafikv++;
    }


  ComboBox1->ItemIndex=-1;
  ProgressBar->Position = 0;
  ProgressBar->Visible = true;
  ProgressBar->Max=ComboBox1->Items->Count-1;
  StatusBar1->SimpleText = "Идет расчет графиков...";

  n=0;

  while (n<kol_grafikv)
    {
      status=0;

      //Номер рассчитываемого графика
      graf = StrToInt(n_grafikv[n]);

      //Выбор общих данных по графику
      DM->qOgraf->Close();
      DM->qOgraf->Parameters->ParamByName("pograf")->Value = graf;
      try
        {
          DM->qOgraf->Open();
        }
      catch (...)
        {
          Application->MessageBox(("Невозможно получить данные по графику "+IntToStr(graf)+" из справочника SPOGRAF").c_str(),"",
                                    MB_OK + MB_ICONERROR);
          StatusBar1->SimpleText = "Отчетный период: "+IntToStr(god)+" год";
          Abort();
        }

      //Вывод графика
      DM->qGrafik->Close();
      DM->qGrafik->Parameters->ParamByName("pgod")->Value=god;
      DM->qGrafik->Parameters->ParamByName("pograf")->Value=graf;
      try
        {
          DM->qGrafik->Open();
        }
      catch(...)
        {
          Application->MessageBox(("Невозможно получить данные по графику "+IntToStr(graf)+" из таблицы SPGRAFIKI").c_str(),"",
                                    MB_OK + MB_ICONERROR);
          StatusBar1->SimpleText = "Отчетный период: "+IntToStr(god)+" год";
          Abort();
        }

      RaschetGraf(graf, god);
      n++;
      ProgressBar->Position++;

      //если график рассчитан
      if (status==0)
        {
          StatusBar1->SimpleText = "Идет расчет графиков... Расчет по графику "+IntToStr(graf)+" выполнен";
        }
    }

  ProgressBar->Visible = false;
  StatusBar1->SimpleText = "Отчетный период: "+IntToStr(god)+" год";
  Application->MessageBox("Расчет по графикам выполнен!","Результат расчета",
                               MB_OK + MB_ICONINFORMATION);
}
//---------------------------------------------------------------------------

//Расчитать текущий график
void __fastcall TMain::N5RaschetOdinClick(TObject *Sender)
{
  status = 0;  //статус расчета графика (status = 1 - не рассчитан)

  //Существование праздничных дней за выбранный год
  DM->qObnovlenie2->Close();
  DM->qObnovlenie2->SQL->Clear();
  DM->qObnovlenie2->SQL->Add("select * from sp_prd where god="+IntToStr(god));
  try
    {
      DM->qObnovlenie2->Open();
    }
  catch(...)
    {
      Application->MessageBox("Ошибка выбора данных из справочника праздничных дней","Ошибка",
                               MB_OK+MB_ICONERROR);
      Abort();
    }

  if (DM->qObnovlenie2->RecordCount==0)
    {
      Application->MessageBox(("Нет данных о праздничных днях за "+IntToStr(god)+" год в справочнике праздничных дней (SP_PRD)").c_str(), "Предупреждение",
                               MB_OK+MB_ICONINFORMATION);
      Abort();
    }

  if (!ComboBox1->Text.IsEmpty())
    {
      if (Application->MessageBox(("Будет выполнен расчет "+ComboBox1->Text+" графика на "+IntToStr(god)+ " год. \nРанее расчитаные данные по этому графику будут удалены. Продолжить?").c_str(),
                                   "Расчет всех графиков", MB_YESNO+MB_ICONINFORMATION)==ID_NO)
        {
          Abort();
        }

      graf = StrToInt(ComboBox1->Text);
      RaschetGraf(graf, god);

      StatusBar1->SimpleText = "Отчетный период: "+IntToStr(god)+" год";


      //Если график был рассчитан
      if (status==0)
        {
          Application->MessageBox("Расчет выполнен успешно!","Результат расчета",
                                  MB_OK + MB_ICONINFORMATION);

          //Проверка доступен ли график для редактирования
          if (Prava == 1)
            {
              for (int i=1; i<=kol_grafik; i++)
                {
                  if (n_grafik[i]==StrToInt(ComboBox1->Text))
                    {
                      redakt = 1;
                      break;
                    }
                  else
                    {
                      redakt = 0;
                    }
                }
            }

          //Видимость пунктов контекстного меню
         if (DM->qGrafik->RecordCount==0)
           {
             N3Redaktirovat->Enabled=false;

             if (redakt==0)
               {
                 N3Redaktirovat->Visible = false;  //Редактировать график
                 N5RaschetOdin->Visible = false; //Расчитать график
               }

           }
         else if (redakt==0)
           {
             N3Redaktirovat->Visible = false;  //Редактировать график
             N5RaschetOdin->Visible = false; //Расчитать график
           }
         else
           {
             N3Redaktirovat->Visible = true;
             N5RaschetOdin->Visible = true;
             N3Redaktirovat->Enabled = true;
             N5RaschetOdin->Enabled = true;
           }

        }
    }
}
//---------------------------------------------------------------------------

//Расчет в зависимости от графика
void __fastcall TMain::RaschetGraf(int graf, int year)
{
  AnsiString Sql;
  int kol_br, day1, day2;

  /* br - текущий номер бригады,
     kol_br - количество бригад в графике*/


  //Часы по 11 графику по месяцам для расчета переработки по другим графикам
  if  (graf!=11 && graf!=81 && graf!=111 && graf!=650 && graf!=655 && graf!=1655 && graf!=660 && graf!=820 && graf!=830 &&
       graf!=1011 && graf!=2011 && graf!=3011 && graf!=18 && graf!=1018 && graf!=2018 && graf!=3018)
    {
      DM->qNorma11Graf->Close();
      DM->qNorma11Graf->Parameters->ParamByName("pgod")->Value = year;
      DM->qNorma11Graf->Parameters->ParamByName("pograf")->Value = DM->qOgraf->FieldByName("norma")->AsString;
      DM->qNorma11Graf->Open();

      //Если 11 график не рассчитан
      if (DM->qNorma11Graf->RecordCount==0)
        {
          Application->MessageBox(("Для расчета выбранного графика сначала \nнеобходимо выполнить расчет "+DM->qOgraf->FieldByName("norma")->AsString+" графика").c_str(),"Предупреждение",
                                   MB_OK + MB_ICONWARNING);
          StatusBar1->SimpleText = "Отчетный период: "+IntToStr(god)+" год";
          Abort();

        }
      else if (DM->qNorma11Graf->RecordCount < 12 || DM->qNorma11Graf->RecordCount > 12)
        {
          Application->MessageBox("Неверно выполнен расчет 11 графика. \nДля расчета выбранного графика сначала \nвыполните повторный расчет 11 графика (81 графика)","Ошибка",
                                   MB_OK + MB_ICONWARNING);
          StatusBar1->SimpleText = "Отчетный период: "+IntToStr(god)+" год";
          Abort();
        }
    }

  //Проверка на существование графика на текущий год
  if (DM->qGrafik->RecordCount>0)
    {
     /*  if (Application->MessageBox((IntToStr(graf)+" график на текущий год уже рассчитан. \nПроизвести повторный расчет?").c_str(),"Расчет графика",
                               MB_YESNO + MB_ICONINFORMATION)==ID_YES)       */
        {
          //удаление графика
          Sql = "delete from spgrafiki where ograf="+IntToStr(graf)+" and god="+year;

          DM->qObnovlenie->Close();
          DM->qObnovlenie->SQL->Clear();
          DM->qObnovlenie->SQL->Add(Sql);
          try
            {
              DM->qObnovlenie->ExecSQL();
            }
          catch(...)
            {
              Application->MessageBox("Возникла ошибка при удалении графика","Ошибка удаления",
                                      MB_OK+MB_ICONERROR);
              Abort();
            }
        }
    }

  //Считывание данных по бригаде с прошлого года
  //  Sql = "select * from spgrafiki where god="+IntToStr(year-1)+" and mes=12 and ograf="+graf+" order by graf ";
  /*  Sql = "select * from spgrafiki                                                     \
         where god="+IntToStr(year-1)+" and ograf="+graf+"                                                \
               and mes=(select max(mes) from spgrafiki where ograf="+graf+" and god="+IntToStr(year-1)+") \
         order by graf";      */


  Sql = "select * from spgrafiki                                                                                    \
         where god="+IntToStr(year-1)+" and ograf="+graf+"                                                                                \
         and (graf,mes) in (select graf, max(mes) as mes from spgrafiki  where ograf="+graf+" and god="+IntToStr(year-1)+" group by graf) \
         order by graf ";

  DM->qObnovlenie2->Close();
  DM->qObnovlenie2->SQL->Clear();
  DM->qObnovlenie2->SQL->Add(Sql);
  try
    {
      DM->qObnovlenie2->Open();
    }
  catch(...)
    {
      Application->MessageBox("Невозможно выбрать информацию по данному графику за прошлый год","Ошибка",
                              MB_OK+MB_ICONERROR);
      Abort();
    }

  kol_br = DM->qObnovlenie2->RecordCount;
  br=1;

  if (kol_br<1)
    {
      Application->MessageBox(("Невозможно расчитать "+IntToStr(graf)+" график.\nНет данных за прошлый год!!!").c_str(), "Прпедупреждение",
                              MB_OK+MB_ICONINFORMATION);
      StatusBar1->SimpleText = "Идет расчет графиков... Расчет по графику "+IntToStr(graf)+" не выполнен";
      status=1;
      return;
    }

  //Расчет графика по бригадам
  while (br <= kol_br)
    {
     /* //Очистка массива
      for(mes=1; mes<=12; mes++)
        {
          for(den=1; den<=31; den++)
            {
              chf[mes][den] = "NULL";
              vihod[mes][den] = "NULL";
              vchf[mes][den] = "NULL";
              pchf[mes][den] = "NULL";
              nchf[mes][den] = "NULL";
            }
          ochf[mes]=0;
          ovchf[mes]=0;
          onchf[mes]=0;
          opchf[mes]=0;
          pgraf[mes]=0;
          chf0[mes]=0;
          nchf0[mes]=0;
          pchf0[mes]=0;
       }   */

      //Очистка массива
      for(mes=0; mes<=12; mes++)
        {
          for(den=0; den<=31; den++)
            {
              chf[mes][den] = NULL;
              vihod[mes][den] = NULL;
              vchf[mes][den] = NULL;
              pchf[mes][den] = NULL;
              nchf[mes][den] = NULL;
            }
          ochf[mes]=0;
          ovchf[mes]=0;
          onchf[mes]=0;
          opchf[mes]=0;
          pgraf[mes]=0;
          chf0[mes]=0;
          nchf0[mes]=0;
          pchf0[mes]=0;
       }


      mes_n=1, mes_k=12; // количество месяцев в графике для расчета и отображения

      if (graf!=11 && graf!=81 && graf!=111 && graf!=650 && graf!=655 && graf!=1655 && graf!=660 && graf!=820 && graf!=830 &&
          graf!=1011 && graf!=2011 && graf!=3011 && graf!=18 && graf!=1018 && graf!=2018 && graf!=3018)
        {
          //Первый месяц 11 графика
          DM->qNorma11Graf->First();
        }
      //Расчет 11 графика
      if (graf==11 || graf==81 || graf==111 || graf==820 || graf==830 ||
          graf==1011 || graf==2011 || graf==3011 || graf==18 || graf==1018 || graf==2018 || graf==3018)    // || graf==655 || graf==1655
        {
          Graf11();
        }
      //Расчет 20 графика
      if ( graf==20 || graf==1020 || graf==2020)   // graf==25 ||
        {
         /* if (graf==25) {v=4; n=1; r=0.5;}   //вечерние и ночные часы, на сколько уменьшаются вечерние часы если это предпраздничный день
          else*/ if (graf==20) {v=4; n=1.5; r=0.5;}   //r - разница между вечерними часами в праздничный и обычный день
          else if (graf==1020) {v=4; n=1; r=1;}
          else if (graf==2020) {v=3.5; n=2; r=0;}

          Graf25(v, n, r);
        }
      //Расчет 23 графика
      if (graf==23)
        {
          d1=4;      //d1 - длительность смены
          v=3;      //v - вечерние часы 2-ой смены
          n=1;      //n - ночные часы 2-ой смены

          Graf23(d1,v,n);
        }
      //Расчет 24 графика
      if (graf==24)
        {
          d1=8;            //d1 - длительность смены
          v=3;             //v - вечерние часы в праздничные дни
          n=1;             //n - ночные часы в праздничные дни

          Graf24(d1,v,n);
        }
      //Расчет 40 графика
      else if(graf==40||graf==1040||graf==2040 ||graf==3040)
        {
          if (graf==40)        {d1=8.5, d2=8, d3=7.5, p1=1.5; p2=7; p=7.5; n1=1.5; n2=6; n=0.5;}    // p1/p2 - праздничные часы(окончание/начало суток), p - праздничные часы
          else if (graf==1040) {d1=d2=d3=8; p1=1; p2=7; p=8; n1=1; n2=6; n=1;}            // n1, n2 - ночные часы 1 смены (за 1-ю и 2-ю половину смены соответственно), n - ночные часы 3 смены
          else if (graf==2040) {d1=9, d2=8, d3=7; p1=2; p2=7; p=8; n1=2; n2=6; n=NULL;}            // d1 - длительность 1 смены, d2 - длительность 2 смены, d3 - длительность 3 смены
          else if (graf==3040) {d1=8, d2=8, d3=8; p1=0; p2=8; p=8; n1=0; n2=6; n=2;} 

          Graf40(d1,d2,d3,p1,p2,p,n1,n2,n);
        }
      //Расчет 60 и 70 графика
      else if (graf==60 || graf==1060 ||graf==2060||graf==3060 ||
               graf==70 || graf==1070 ||graf==2070||graf==3070 ||
               graf==170 || graf==50)
        {
          if (graf==60) {p1=4.5; p2=7.5; v1=1.5; v2=2.5; n1=2; n2=6;}          // p1/p2 - праздничные часы(окончание/начало суток)
          else if (graf==1060) {p1=5.5; p2=6.5; v1=0.5; v2=3.5; n1=2; n2=6;}   // v1, v2 - вечерние часы первой и второй смены
          else if (graf==2060) {p1=5; p2=7; v1=1; v2=3; n1=2; n2=6;}           // n1,n2 - ночные часы(окончание/начало суток)
          else if (graf==3060) {p1=4; p2=8; v1=2; v2=2; n1=2; n2=6;}
          else if (graf==70) {p1=4.5; p2=7; v1=1.5; v2=2.5; n1=2; n2=5.5;}
          else if (graf==1070) {p1=5.5; p2=6; v1=0.5; v2=3.5; n1=2; n2=5.5;}
          else if (graf==2070) {p1=5; p2=6.5; v1=1; v2=3; n1=2; n2=5.5;}
          else if (graf==3070) {p1=4; p2=7.5; v1=2; v2=2; n1=2; n2=5.5;}
          else if (graf==170) {p1=4; p2=7; v1=2; v2=2; n1=2; n2=5;}
          else if (graf==50) {p1=4; p2=8; v1=2; v2=2; n1=2; n2=6;}

          Graf60(p1, p2, v1, v2, n1, n2);

        }
      //Расчет 4060 графика
      else if (graf==4060)
        {
          p=10, p1=p2=7; v1=0; v2=4; n1=2; n2=6; // p - длительность 1 смены, p1/p2 - праздничные часы(окончание/начало суток) 2 смены
                                                // v1, v2 - вечерние часы первой и второй смены, n1,n2 - ночные часы(окончание/начало суток)
          Graf4060(p, p1, p2, v1, v2, n1, n2);
        }
      //Расчет 85 графика
      else if (graf==85)
        {
          n=1.5;           //ночные смены
          Graf85(n);
        }
      //Расчет 90 графика
      else if(graf==90)
        {
          if (graf==90) {d1=9; d2=8; d3=7; p1=2; p2=7; p=7; v=4; n1=2; n2=6; n=0;}     // p1/p2 - праздничные часы(окончание/начало суток), p - праздничные часы
                                                                                              // n1 - ночные часы 1 смены, n2 - ночные часы 3 смены
          Graf90(d1,d2,d3,p1,p2,p,v,n1,n2,n);
        }
      //Расчет 120 графика
      else if (graf==120)
        {
          d1=9, d2=8, d3=7; //длительность каждой смены
          p1=2, p2=7;         //окончание/начало суток
          v=4;                  //вечерние
          n1=2, n2=6;         //ночные часы окончание /начало 1 смены
          n=0;                //ночные часы 3 смены

          Graf120(d1, d2, d3, p1, p2, v, n1, n2, n);
        }
      //Расчет 133 графика
   /*   else if (graf==133)
        {
          v=4;
          n=0.5;
          Graf133(v, n);
        }*/
      //Расчет 140 графика
      if (graf==140)
        {
          d1=8;              //d1 - длительность смены
          v=2;               //v - вечерние часы второй смены
          n=2;               //n - ночные часы второй смены

          Graf140(d1,v,n);
        }
      //Расчет 150 графика
      else if (graf==150)
        {
          Graf150();
        }
      //Расчет 160 графика
      if (graf==160)
        {
          d1=8;              //d1 - длительность смены
          v=3.5;             //v - вечерние часы второй смены
          n=2;               //n - ночные часы второй смены

          Graf160(d1,v,n);
        }
      //Расчет 180 графика
      if(graf==180)
        {
          d1=8;                 // d1 - длительность смены
          p1=1; p2=7;           // p1/p2 - праздничные часы(окончание/начало суток)
          v=4;                  // v - вечерние часы
          n1=1; n2=6; n=1;  // n1, n2 - ночные часы 1 смены (за 1-ю и 2-ю половину смены соответственно), n - ночные часы 3 смены



          Graf180(d1,p1,p2,v,n1,n2,n);
        }
      //Расчет 190 графика
      else if (graf==190)
        {
          d1=11; v1=0.5; v2=3.5;     // d1 - длительность смены, v1 - вечерние часы первой и второй смены, v2 - вечерние часы третьей и четвертой
          Graf190(d1,v1,v2);
        }
      //Расчет 220 графика
      else if (graf==220)
        {
          v=3.5; n=0.75;     // v - вечерние часы, n - ночные часы
          Graf220(v,n);
        }
      //Расчет 225 графика
     /* else if (graf==225)
        {
          v1=3.5, v2=4;
          n=0.5;              // v - вечерние часы, n - ночные часы
          Graf225(v1, v2, n);
        }   */
      //Расчет 230 графика
      else if (graf==230)
        {
          v=1;
          Graf230(v);
        }
      //Расчет 240 графика
    /*  else if (graf==240)
        {
          v=3.5;
          n1=2, n2=1.7;     //n1 - ночные часы первой смены, n2 - ночные часы второй смены
          Graf240(v, n1, n2);
        }  */

      //Расчет 250 графика
      if (graf==250)
        {
          d1=8;      //d1 - длительность смены
          v1=2;      //v1 - вечерние часы 2-ой бригады вторник-пятница
          v2=3;      //v2 - вечерние часы 2-ой бригады пятница, праздничный день
          n=1;      //n - ночные часы 2-ой смены

          Graf250(d1,v1,v2,n);
        }
        /*
      //Расчет 260 графика
      else if (graf==260)
        {
          v=3.5;
          n=1;
          Graf260(v, n);
        }   */

      //Расчет 270 графика
      if (graf==270)
        {
          d1=8.25;      //d1 - длительность смены
          v=3.5;      //v - вечерние часы 2-ой смены
          n=2;      //n - ночные часы 2-ой смены

          Graf270(d1,v,n);
        }
      //Расчет 280 графика
      if (graf==280)
        {
          d1=8;      //d1 - длительность смены
          v=3;       //v - вечерние часы

          Graf280(d1,v);
        }
      //Расчет 300 и 131 графика
      else if (graf==300||graf==1300||graf==2300||graf==3300|| graf==335 || graf==131 || graf==4300)  //|| graf==935
        {
          if (graf==300 || graf==935) v=1;              //v - вечерние часы
          else if (graf==1300) v=0.5;
       //   else if (graf==2300) v=1.5;
          else if (graf==3300 || graf==335 || graf==4300) v=2;
       //   else if (graf==131) v=1.5;

          Graf300(v);
        }
      //Расчет 315 графика
      else if (graf==315)
        {
          v=0.9;
          Graf315(v);
        }
      //Расчет 320 графика
      else if (graf==320)
        {
          v=4;
          n1=2; n2=6;
          p1=16; p2=8;
          Graf320(v, n1, n2, p1, p2);
        }
      //Расчет 370 графика
      else if (graf==370)
        {
          v=4;
          n1=1.5; n2=6;
          p1=5.5; p2=6.5;
          Graf370(v, n1, n2, p1, p2);
        }

      //Расчет 390 и 950 график
      else if (graf==390||graf==1390||graf==950)
        {
          if (graf==390||graf==950) {p1=15; p2=8; n1=1.5; n2=6;} //p1/p2 - праздничные часы (начало суток/окончание суток)
          else if (graf==1390) {p1=16; p2=7; n1=1.5; n2=6;}      //n1,n2 - ночные часы (начало суток/окончание суток)

          Graf390(p1,p2,n1,n2);
        }
       //Расчет 400 графика
     /* if (graf==400)
        {
          v=1;         //вечерние часы
          Graf400(v);
        }
      //Расчет 410 графика
      if (graf==410)
        {
          v=0.9;         //вечерние часы
          Graf410(v);
        }   */
      //Расчет 450 графика
      if (graf==450)
        {
          v=3;               // вечерние часы
          n1=2, n2=5.5;      // n1 - ночные часы окончание суток, n2 - ночные часы начало суток
          p1=5, p2=6.5;      // р1 - праздничные окончание суток, р2 - праздничные начало суток
          Graf450(v, n1, n2, p1 ,p2);
        }
      //Расчет 470 графика
      if (graf==470)
        {
          v=3.5;
          n=1.5;
          
          Graf470(v, n);
        }
      //Расчет 480 графика
      else if (graf==480)
        {
          Graf480();
        }
      //Расчет 210(ХМФ) или 525 графика
      else if (graf==520 || graf==210) //  || graf==525
        {
          if (graf==520)
            {
              d1=15.5, d2=24, d3=16.5;        // длительность смены
              p1=8.5, p2=17, p3=9.5, p=7;     //р1, р2, р3 -  окончание суток / р - начало суток
              v=4;                            //вечерние часы
              n1=2, n2=6;                     //n1 - окончание суток, n2 - начало суток (ночные часы)
            }
          else if (graf==210)
            {
              d1=15, d2=24, d3=16;        // длительность смены
              p1=7, p2=16, p3=8, p=8;     //р1, р2, р3 -  окончание суток / р - начало суток
              v=4;                            //вечерние часы
              n1=2, n2=6;                     //n1 - окончание суток, n2 - начало суток (ночные часы)
            }

          Graf520(d1, d2, d3, p1, p2, p3, p, v, n1, n2);
        }
      //Расчет 630 графика
      if (graf==630 ||graf==1630 )
        {
          Graf630();
        }
      //Расчет 650 и 660 графика
    /*  if (graf==650 || graf==660)
        {
          Graf650();
        }  */
      //Расчет 670 и 790 графика
      else if (graf==670 || graf==790 )
        {
          if (graf==670) v=2;
          else if (graf==790) v=1.5;  //вечерние часы

          Graf670(v);
        }
      //Расчет 680 графика
    /*  else if (graf==680)
        {
          v=1;                        //вечерние часы
          Graf680(v);
        }       */
      //Расчет 690 графика
      else if (graf==690)
        {
          v=3;
          Graf690(v);
        }
      //Расчет 771 графика
    /*  else if (graf==771)
        {
          Graf771();
        }
      //Расчет 775 графика
      else if (graf==775)
        {
          v=4;
          n=0.5;
          Graf775(v, n);
        } */
      //Расчет 780 графика
      else if (graf==780)
        {
          Graf780();
        }
      //Расчет 800 графика
      else if (graf==800)
        {
          if (br==1) day1=6, day2=7;  //рабочие дни с воскресенья по четверг, выходные пятница=6 и суббота=7
          else day1=1, day2=2;        //рабочие дни с вторника по субботу, выходные воскресенье=1 и понедельник=2

          Graf800(day1, day2);       //day1 и day2 - выходные дни
        }
      //Расчет 850 графика
      else if (graf==850)
        {
          v=2.5;

          Graf850(v);
        }
      //Расчет 855 графика
      else if (graf==855 || graf==880)
        {
          if (graf==855) v=2;
          else if (graf==880) v=1.75;

          Graf855(v);
        }
      //Расчет 865 графика
    /*  else if (graf==865)
        {
          v=2;
          Graf865(v);
        }     */
      //Расчет 960 графика
      else if (graf==960)
        {
          d1=8, d2=7, d3=10, d4=6, d5=9;    //d1-d2 - длительность смены в зависимости от периода
          v=3;                              //v - вечерние часы
          n=0.5;                            //n - ночные часы
          Graf960(d1, d2, d3, d4, d5, v, n);
        }
      //Расчет 980 графика
     /* else if (graf==980)
        {
          p1=4, p2=8;
          v=2;                            //v - вечерние часы
          n1=2, n2=6;                     //n1,n2 - ночные часы(окончание суток/начало суток)
          Graf980(p1, p2, v, n1, n2);
        }  */

      //вставка графика рассчитанного на год
      for (mes=mes_n; mes<=mes_k; mes++)
        {

          Sql = "insert into spgrafiki (god,mes, ograf, graf, chf, vch, nch, pch, pgraf, nsm, dnism,\
                                        chf0, nch0, pch0, \
                                        nsm1, nsm2, nsm3, nsm4, nsm5, nsm6, \
                                        nsm7, nsm8, nsm9, nsm10, nsm11, nsm12, nsm13, nsm14, nsm15, nsm16, \
                                        nsm17, nsm18, nsm19, nsm20, nsm21, nsm22, nsm23, nsm24, nsm25, nsm26, \
                                        nsm27, nsm28, nsm29, nsm30, nsm31,\
                                        chf1, chf2, chf3, chf4, chf5, chf6, chf7, chf8, \
                                        chf9, chf10, chf11, chf12, chf13, chf14, chf15, chf16, chf17, \
                                        chf18, chf19, chf20, chf21, chf22, chf23, chf24, chf25, chf26, \
                                        chf27, chf28, chf29, chf30, chf31,\
                                        vch1, vch2, vch3, vch4, vch5, vch6, vch7, vch8, \
                                        vch9, vch10, vch11, vch12, vch13, vch14, vch15, vch16, vch17, vch18, vch19, \
                                        vch20, vch21, vch22, vch23, vch24, vch25, vch26, vch27, vch28, vch29, vch30, vch31,\
                                        nch1, nch2, nch3, nch4, nch5, nch6, nch7, nch8, nch9, nch10, nch11, \
                                        nch12, nch13, nch14, nch15, nch16, nch17, nch18, nch19, nch20, nch21, nch22, \
                                        nch23, nch24, nch25, nch26, nch27, nch28, nch29, nch30, nch31, \
                                        pch1, pch2, pch3, pch4, pch5, pch6, pch7, pch8, pch9, pch10, pch11, \
                                        pch12, pch13, pch14, pch15, pch16, pch17, pch18, pch19, pch20, pch21, pch22, \
                                        pch23, pch24, pch25, pch26, pch27, pch28, pch29, pch30, pch31 )  \
               values("+IntToStr(year)+"," + IntToStr(mes) +" ,"+ IntToStr(graf)+","+ \
                        QuotedStr(DM->qObnovlenie2->FieldByName("graf")->AsString)+","+ochf[mes]+","+ ovchf[mes] +","+ onchf[mes] +","+ opchf[mes]+","+ pgraf[mes];

               if (mes==12|| (graf==680 && mes==9)){Sql+=","+IntToStr(nsm)+","+IntToStr(dnism);} else {Sql+=",'',''";}
               Sql = Sql + "," + chf0[mes] +"," +nchf0[mes]+ "," +pchf0[mes];
               for(den=1; den<=31; den++) Sql = Sql + "," + (vihod[mes][den]);
               for(den=1; den<=31; den++) Sql = Sql + "," + (chf[mes][den]);
               for(den=1; den<=31; den++) Sql = Sql + "," + (vchf[mes][den]);
               for(den=1; den<=31; den++) Sql = Sql + "," + (nchf[mes][den]);
               for(den=1; den<=31; den++) Sql = Sql + "," + (pchf[mes][den]);
               Sql = Sql +")";

          DM->qObnovlenie->Close();
          DM->qObnovlenie->SQL->Clear();
          DM->qObnovlenie->SQL->Add(Sql);
          try
            {
              DM->qObnovlenie->ExecSQL();
            }
          catch(...)
            {
              Application->MessageBox("Ошибка вставки данных по графику","Ошибка записи",
                                      MB_OK + MB_ICONERROR);
              StatusBar1->SimpleText = "Отчетный период: "+IntToStr(god)+" год";
              Abort();
            }
        }

      DM->qObnovlenie2->Next();
      br++;
    }

  DM->qGrafik->Requery();

  StatusBar1->SimpleText = "Идет расчет графиков...  расчитан "+IntToStr(graf)+" график" ;


  InsertLog("Расчет "+IntToStr(graf)+" графика на "+year+" год выполнен успешно");

}

//---------------------------------------------------------------------------

//Расчет 11 графика
void __fastcall TMain::Graf11()
{
  int kol, prazd;

  /* chf[32] - рабочие часы по дням
     chf[den] = 8 - рабочий день
     chf[den] = 7 - предпраздничный день
     vihod[32] - выходы по дням (рабочий, отдых, праздничный)
     vihod[den] = 1 - рабочий день
     vihod[den] = 9 - праздник
     prazd - количество часов на которое сокращается предпраздничная смена
  */

  if (graf==11 || graf==1011 || graf==2011 || graf==3011 || graf==18 || graf==1018 || graf==2018 || graf==3018) prazd=1;
  else prazd=0;

  //по месяцам
  for (mes=1; mes<=12; mes++)
    {
      kol = DaysInAMonth(god, mes);

      //по дням
      for (den=1; den<=kol; den++)
        {
          //проверка дня недели
          if (DayWeek(den,mes,god)==1||DayWeek(den,mes,god)==7)
            {
              //проверка праздничного дня
              if (PrazdDni(den,mes)==true)
                {
                  vihod[mes][den]=9;
                  chf[mes][den]=30;
                }
              else
                {
                  vihod[mes][den]=0;
                  chf[mes][den]=0;
                }
            }
          else
            {
              //проверка праздничного дня
              if (PrazdDni(den,mes)==true)
                {
                  vihod[mes][den]=9;
                  chf[mes][den]=30;
                }
              //праздничный попадающий на субботу или воскресенье
              else if (PrazdDniVihodnue(den,mes,god)==true)
                {
                  vihod[mes][den]=0;
                  chf[mes][den]=0;
                }
              //проверка предпраздничного дня
              else if (PrdPrazdDni(den,mes)==true)
                {
                  vihod[mes][den]=1;
                  chf[mes][den]=DM->qOgraf->FieldByName("DLIT")->AsFloat-prazd;
                  ochf[mes]+= DM->qOgraf->FieldByName("DLIT")->AsFloat-prazd;
                }
              //рабочий день
              else
                {
                  vihod[mes][den]=1;
                  chf[mes][den]=DM->qOgraf->FieldByName("DLIT")->AsFloat;
                  ochf[mes]+= DM->qOgraf->FieldByName("DLIT")->AsFloat;
                }
            }

        }
      //отсутствующие дни в месяце
      if (den<32)
        {
          while (den<=32)
            {
              vihod[mes][den]="NULL";
              chf[mes][den]="NULL";
              den++;
            }
        }
    }
}
//---------------------------------------------------------------------------

//Расчет 23 графика
void __fastcall TMain::Graf23(double d1, double v, double n)
{
  int kol;

  /* chf[32] - рабочие часы по дням
     chf[den] = 8 - рабочий день
     chf[den] = 7 - предпраздничный день
     vihod[32] - выходы по дням (рабочий, отдых, праздничный)
     vihod[den] = 1 - рабочий день
     vihod[den] = 9 - праздник
     prazd - количество часов на которое сокращается предпраздничная смена
  */

  //по месяцам
  for (mes=1; mes<=12; mes++)
    {
      kol = DaysInAMonth(god, mes);

      //по дням
      for (den=1; den<=kol; den++)
        {
          //проверка дня недели (вторник - четверг)
          if (DayWeek(den,mes,god)==3||DayWeek(den,mes,god)==4||DayWeek(den,mes,god)==5)
            {
              //проверка праздничного дня
              if (PrazdDni(den,mes)==true)
                {
                  vihod[mes][den]=1;
                  chf[mes][den]=d1;
                  vchf[mes][den]=v;
                  nchf[mes][den]=n;
                  pchf[mes][den]=d1;

                  //сумма часов
                  ochf[mes]+=d1;
                  ovchf[mes]+=v;
                  onchf[mes]+=n;
                  opchf[mes]+=d1;
                }
              else
                {
                  vihod[mes][den]=1;
                  chf[mes][den]=d1;

                  //сумма часов
                  ochf[mes]+=d1;
                }
            }
          //проверка дня недели (пятница - суббота)
          else if (DayWeek(den,mes,god)==6||DayWeek(den,mes,god)==7)
            {
              vihod[mes][den]=1;
              chf[mes][den]=d1;
              vchf[mes][den]=v;
              nchf[mes][den]=n;

              //сумма часов
              ochf[mes]+=d1;
              ovchf[mes]+=v;
              onchf[mes]+=n;

              //проверка праздничного дня
              if (PrazdDni(den,mes)==true)
                {
                  pchf[mes][den]=d1;
                  opchf[mes]+=d1;
                }
            }
          //выходной
          else
            {
              vihod[mes][den]=0;
              chf[mes][den]=NULL;
              vchf[mes][den]=NULL;
              nchf[mes][den]=NULL;
              pchf[mes][den]=NULL;
            }
        }
      //отсутствующие дни в месяце
      if (den<32)
        {
          while (den<=32)
            {
              vihod[mes][den]="NULL";
              chf[mes][den]="NULL";
              den++;
            }
        }
    }
}
//---------------------------------------------------------------------------

//Расчет 24 графика
void __fastcall TMain::Graf24(double d1, double v, double n)
{
  int kol;

  //по месяцам
  for (mes=1; mes<=12; mes++)
    {
      kol = DaysInAMonth(god, mes);

      //по дням
      for (den=1; den<=kol; den++)
        {
          //проверка дня недели (понедельник - пятница)
          if (DayWeek(den,mes,god)!=7 && DayWeek(den,mes,god)!=1)
            {
              vihod[mes][den]=1;

              //проверка на 31 декабря
              if (den==31 && mes==12)
                {
                  chf[mes][den]=d1;
                  vchf[mes][den]=v;
                  nchf[mes][den]=n;

                  //сумма часов
                  ochf[mes]+=d1;
                  ovchf[mes]+=v;
                  onchf[mes]+=n;
                }
              //проверка праздничного дня
              else if (PrazdDni(den,mes)==true)
                {
                  chf[mes][den]=d1;
                  vchf[mes][den]=v;
                  nchf[mes][den]=n;
                  pchf[mes][den]=d1;

                  //сумма часов
                  ochf[mes]+=d1;
                  ovchf[mes]+=v;
                  onchf[mes]+=n;
                  opchf[mes]+=d1;
                }
              //рабочий день
              else
                {
                  //проверка предпраздничного дня
                  if (PrdPrazdDni(den,mes)==true)
                    {
                      chf[mes][den]=d1-1;
                      //сумма часов
                      ochf[mes]+=d1-1;
                    }
                  else
                    {
                      chf[mes][den]=d1;
                      //сумма часов
                      ochf[mes]+=d1;
                    }
                }
            }
          //выходной
          else
            {
              vihod[mes][den]=0;
              chf[mes][den]=NULL;
              vchf[mes][den]=NULL;
              nchf[mes][den]=NULL;
              pchf[mes][den]=NULL;
            }
        }
      //отсутствующие дни в месяце
      if (den<32)
        {
          while (den<=32)
            {
              vihod[mes][den]="NULL";
              chf[mes][den]="NULL";
              den++;
            }
        }
    }
}
//---------------------------------------------------------------------------

//Расчет 25 графика
void __fastcall TMain::Graf25(double v, double n, double r)
{
  AnsiString kol;

  /*nsm - номер смены последнего месяца,
   dnism - день смены последнего месяца*/

   nsm = DM->qObnovlenie2->FieldByName("nsm")->AsInteger;
   dnism = DM->qObnovlenie2->FieldByName("dnism")->AsInteger;

  //по месяцам
  for (mes=1; mes<=12; mes++)
    {
      kol = DaysInAMonth(god, mes);

      //по дням
      for (den=1; den<=kol; den++)
        {
           //проверка дня недели (выходной)
          if (DayWeek(den,mes,god)==1||DayWeek(den,mes,god)==7)
            {
              //проверка праздничного дня
              if (PrazdDni(den,mes)==true)
                {
                  vihod[mes][den]=9;
                  chf[mes][den]=30;
                }
              else
                {
                  vihod[mes][den]=0;
                  chf[mes][den]=0;
                }
            }
          //рабочий день
          else
            {
              //вторая смена (6.30-15.00)
              if (nsm==2)
                {
                  //проверка праздничного дня
                  if (PrazdDni(den,mes)==true)
                    {
                      vihod[mes][den]=9;
                      chf[mes][den]=30;
                    }
                  //праздничный попадающий на субботу или воскресенье
                  else if (PrazdDniVihodnue(den,mes,god)==true)
                    {
                      vihod[mes][den]=0;
                      chf[mes][den]=0;
                    }
                  //проверка предпраздничного дня
                  else if (PrdPrazdDni(den,mes)==true)
                    {
                      vihod[mes][den]=2;
                      chf[mes][den]=DM->qOgraf->FieldByName("DLIT")->AsFloat-1;
                      ochf[mes]+= DM->qOgraf->FieldByName("DLIT")->AsFloat-1;
                    }
                  //рабочий день
                  else
                    {
                      vihod[mes][den]=2;
                      chf[mes][den]=DM->qOgraf->FieldByName("DLIT")->AsFloat;
                      ochf[mes]+= DM->qOgraf->FieldByName("DLIT")->AsFloat;
                    }

                  if (dnism==5)
                    {
                      nsm=3;
                      dnism=1;
                    }
                  else
                    {
                      dnism++;
                    }
                }
              //третья смена (14.30-23.00)
              else if (nsm==3)
                {
                  //проверка праздничного дня
                  if (PrazdDni(den,mes)==true)
                    {
                      vihod[mes][den]=9;
                      chf[mes][den]=30;
                    }
                  //праздничный попадающий на субботу или воскресенье
                  else if (PrazdDniVihodnue(den,mes,god)==true)
                    {
                      vihod[mes][den]=0;
                      chf[mes][den]=0;
                    }
                  //проверка предпраздничного дня
                  else if (PrdPrazdDni(den,mes)==true)
                    {
                      vihod[mes][den]=3;
                      chf[mes][den]=DM->qOgraf->FieldByName("DLIT")->AsFloat-1;
                      vchf[mes][den]= v-r;

                      ochf[mes]+= DM->qOgraf->FieldByName("DLIT")->AsFloat-1;
                      ovchf[mes]+=v-r;
                    }
                  //рабочий день
                  else
                    {
                      vihod[mes][den]=3;
                      chf[mes][den]=DM->qOgraf->FieldByName("DLIT")->AsFloat;
                      vchf[mes][den]=v;
                      nchf[mes][den]=n;

                      ochf[mes]+= DM->qOgraf->FieldByName("DLIT")->AsFloat;
                      ovchf[mes]+=v;
                      onchf[mes]+=n;
                    }

                  if (dnism==5)
                    {
                      nsm=2;
                      dnism=1;
                    }
                  else
                    {
                      dnism++;
                    }
                }
            }
        }
      //отсутствующие дни в месяце
      if (den<32)
        {
          while (den<=32)
            {
              vihod[mes][den]="NULL";
              chf[mes][den]="NULL";
              nchf[mes][den]="NULL";
              vchf[mes][den]="NULL";
              pchf[mes][den]="NULL";
              den++;
            }
        }

      DM->qNorma11Graf->Next();
    }
}
//------------------------------------------------------------------------------

//Расчет 40 графика
void __fastcall TMain::Graf40(double d1, double d2, double d3, double p1, double p2, double p, double n1, double n2, double n)
{
   AnsiString kol;

  /*nsm - номер смены последнего месяца,
   dnism - день смены последнего месяца*/

  nsm = DM->qObnovlenie2->FieldByName("nsm")->AsInteger;
  dnism = DM->qObnovlenie2->FieldByName("dnism")->AsInteger;

  //по месяцам
  for (mes=1; mes<=12; mes++)
    {
      kol = DaysInAMonth(god, mes);

      //по дням
      for (den=1; den<=kol; den++)
        {
          //первая смена (22.30-7.00)
          //*************************
          if (nsm==1)
            {
              vchf[mes][den]="NULL";
              vihod[mes][den]=1;


              //переход на летнее время (март)
              if (mes==3 && den==day_mart2 && dnism!=1)
                {
                  if (dnism==4)
                    {
                      chf[mes][den]=p2-1;
                      nchf[mes][den]=n2-1;
                      //проверка праздничного дня
                      if (PrazdDni(den,mes)==true)
                        {
                          pchf[mes][den]=p2-1;
                          opchf[mes]+=p2-1;
                        }
                    }
                  else
                    {
                      chf[mes][den]=d1-1;
                      nchf[mes][den]=(n1+n2)-1;
                      //проверка праздничного дня
                      if (PrazdDni(den,mes)==true)
                        {
                          pchf[mes][den]=d1-1;
                          opchf[mes]+=d1-1;
                        }

                    }

                  //общие суммы по часам
                  onchf[mes]+=(n1+n2)-1;
                  ochf[mes]+= d1-1;

                  //проверка праздничного дня (праздничные проставляются в день когда пришел со смены)
                /*  if (PrazdDni(den,mes)==true)
                    {
                      if (dnism==4)
                        {
                          if (pchf[mes][den]!="NULL") pchf[mes][den]=FloatToStr(StrToFloat(pchf[mes][den])+p2);
                          else pchf[mes][den]=p2;
                          opchf[mes]+=p2;
                        }
                      else
                        {
                          if (pchf[mes][den]!="NULL") pchf[mes][den]=FloatToStr(StrToFloat(pchf[mes][den])+p2);
                          else pchf[mes][den]=p2;
                          opchf[mes]+=p1+p2;

                          //если не последняя переходящая смена
                          if (den!=kol)
                            {
                              pchf[mes][den+1]=p1;
                            }
                          else
                            {
                              // если последняя смена =(((((((
                              pchf[mes+1][1]=p1;
                            }
                        }

                    }*/
                }
              else if (mes==mes_mart2 && den==day_mart2 && dnism==1)
                {
                  if (dnism==4)
                    {
                      chf[mes][den]=p2-1;
                      nchf[mes][den]=n2-1;
                      //проверка праздничного дня
                      if (PrazdDni(den,mes)==true)
                        {
                          pchf[mes][den]=p2-1;
                          opchf[mes]+=p2-1;
                        }
                    }
                  else
                    {
                      chf[mes][den]=d1-1;
                      nchf[mes][den]=(n1+n2)-1;
                      //проверка праздничного дня
                      if (PrazdDni(den,mes)==true)
                        {
                          pchf[mes][den]=d1-1;
                          opchf[mes]+=d1-1;
                        }
                    }

                  //общие суммы по часам
                  onchf[mes]+=(n1+n2)-1;
                  ochf[mes]+= d1-1;

                  //проверка праздничного дня (праздничные проставляются в день когда пришел со смены)
                 /* if (PrazdDni(den,mes)==true )
                    {
                      if (pchf[mes][den]!="NULL") pchf[mes][den]=FloatToStr(StrToFloat(pchf[mes][den])+p2-1);
                      else pchf[mes][den]=p2-1;
                      opchf[mes]+=p1+p2-1;

                      //если не последняя переходящая смена
                      if (den!=kol)
                        {
                          pchf[mes][den+1]=p1;
                        }
                      else
                        {
                          // если последняя смена =(((((((
                          pchf[mes+1][1]=p1;
                        }
                    } */
                }
              //переход на зимнее время (октябрь)
              else if (mes==10 && den==day_oktyabr2 && dnism!=1)
                {
                  if (dnism==4)
                    {
                      chf[mes][den]=p2+1;
                      nchf[mes][den]=n2+1;
                      //проверка праздничного дня
                      if (PrazdDni(den,mes)==true)
                        {
                          pchf[mes][den]=p2+1;
                          opchf[mes]+=p2+1;
                        }
                    }
                  else
                    {
                      chf[mes][den]=d1+1;
                      nchf[mes][den]=(n1+n2)+1;
                      //проверка праздничного дня
                      if (PrazdDni(den,mes)==true)
                        {
                          pchf[mes][den]=d1+1;
                          opchf[mes]+=d1+1;
                        }
                    }

                  //общие суммы по часам
                  onchf[mes]+=(n1+n2)+1;
                  ochf[mes]+= d1+1;

                  //проверка праздничного дня (праздничные проставляются в день когда пришел со смены)
                 /* if (PrazdDni(den,mes)==true)
                    {
                      if (dnism==4)
                        {
                          if (pchf[mes][den]!="NULL") pchf[mes][den]=FloatToStr(StrToFloat(pchf[mes][den])+p2);
                          else pchf[mes][den]=p2;
                          opchf[mes]+=p2;
                        }
                      else
                        {
                          if (pchf[mes][den]!="NULL") pchf[mes][den]=FloatToStr(StrToFloat(pchf[mes][den])+p2);
                          else pchf[mes][den]=p2;
                          opchf[mes]+=p1+p2;

                          //если не последняя переходящая смена
                          if (den!=kol)
                            {
                              pchf[mes][den+1]=p1;
                            }
                          else
                            {
                              // если последняя смена =(((((((
                              pchf[mes+1][1]=p1;
                            }
                        }

                    } */
                }
              else if (mes==mes_oktyabr2 && den==day_oktyabr2 && dnism==1)
                {
                  if (dnism==4)
                    {
                      chf[mes][den]=p2+1;
                      nchf[mes][den]=n2+1;
                      //проверка праздничного дня
                      if (PrazdDni(den,mes)==true)
                        {
                          pchf[mes][den]=p2+1;
                          opchf[mes]+=p2+1;
                        }
                    }
                  else
                    {
                      chf[mes][den]=d1+1;
                      nchf[mes][den]=(n1+n2)+1;
                      //проверка праздничного дня
                      if (PrazdDni(den,mes)==true)
                        {
                          pchf[mes][den]=d1+1;
                          opchf[mes]+=d1+1;
                        }
                    }

                  //общие суммы по часам
                  onchf[mes]+=(n1+n2)+1;
                  ochf[mes]+= d1+1;

                  //проверка праздничного дня (праздничные проставляются в день когда пришел со смены)
                /*  if (PrazdDni(den,mes)==true)
                    {
                      if (pchf[mes][den]!="NULL") pchf[mes][den]=FloatToStr(StrToFloat(pchf[mes][den])+p2+1);
                      else pchf[mes][den]=p2+1;
                      opchf[mes]+=p1+p2+1;

                      //если не последняя переходящая смена
                      if (den!=kol)
                        {
                          pchf[mes][den+1]=p1;
                        }
                      else
                        {
                          // если последняя смена =(((((((
                          pchf[mes+1][1]=p1;
                        }
                    } */
                }
              else
                {
                  //если 4 день 1 смены сохранять только часы за вторую половину суток
                  if (dnism==4)
                    {
                      chf[mes][den]=p2;
                      nchf[mes][den]=n2;
                      //проверка праздничного дня
                      if (PrazdDni(den,mes)==true)
                        {
                          pchf[mes][den]=p2;
                          opchf[mes]+=p2;
                        }
                    }
                  else
                    {
                      chf[mes][den]=d1;
                      nchf[mes][den]=n1+n2;
                      //проверка праздничного дня
                      if (PrazdDni(den,mes)==true)
                        {
                          pchf[mes][den]=d1;
                          opchf[mes]+=d1;
                        }
                    }

                  //общие суммы по часам
                  onchf[mes]+=n1+n2;
                  ochf[mes]+=d1;

                  //проверка праздничного дня (праздничные проставляются в день когда пришел со смены)
                /*  if (PrazdDni(den,mes)==true )
                    {
                      if (dnism==4)
                        {
                          if (pchf[mes][den]!="NULL") pchf[mes][den]=FloatToStr(StrToFloat(pchf[mes][den])+p2);
                          else pchf[mes][den]=p2;
                          opchf[mes]+=p2;
                        }
                      else
                        {
                          if (pchf[mes][den]!="NULL") pchf[mes][den]=FloatToStr(StrToFloat(pchf[mes][den])+p2);
                          else pchf[mes][den]=p2;
                          opchf[mes]+=p1+p2;

                          //если не последняя переходящая смена
                          if (den!=kol)
                            {
                              pchf[mes][den+1]=p1;
                            }
                          else
                            {
                              // если последняя смена =(((((((
                              pchf[mes+1][1]=p1;
                            }
                        }

                    } */
                }

              if (den==1)
                {
                  ochf[mes]-=p1;
                  onchf[mes]-=p1;
                  //проверка праздничного дня
                 /* if (PrazdDni(den,mes)==true)
                    {
                      opchf[mes]-=p1;
                      pchf0[mes]=StrToFloat(-p1);
                    }
                /*  if (mes!=1)
                    { */
                      chf0[mes]=StrToFloat(-p1);
                      nchf0[mes]=StrToFloat(-p1);




                     /* chf0[mes-1]=p2;
                      nchf0[mes-1]=6;

                      //проверка праздничного дня
                      if (PrazdDni(den,mes)==true)
                        {
                          pchf0[mes-1]=p2;
                        } */
                  //  }
                }
              else if (den==kol && dnism!=4)
                {
                  ochf[mes]+=p1;
                  onchf[mes]+=p1;

                  //проверка праздничного дня
                /*  if (PrazdDni(den,mes)==true)
                    {
                      opchf[mes]+=p1;
                    }   */
                }

              //проверка дня в смене
              if (dnism==4)
                {
                  nsm=0;
                  dnism=1;
                }
              else
                {
                  dnism++;
                }
            }
          //вторая смена (7.00-15.00)
          //*************************
          else if (nsm==2)
            {
              vchf[mes][den]="NULL";
              nchf[mes][den]="NULL";
              chf[mes][den]=d2;
              vihod[mes][den]=2;

              //общие суммы по часам
              ochf[mes]+=d2;


              //проверка праздничного дня
              if (PrazdDni(den,mes)==true)
                {
                  pchf[mes][den]=d2;
                  opchf[mes]+=d2;
                }
              //рабочий день
              else
                {
                  pchf[mes][den]="NULL";
                }

              //проверка дня в смене
              if (dnism==4)
                {
                  nsm=0;
                  dnism=2;
                }
              else
                {
                  dnism++;
                }
            }
          //третья смена (15.00-22.30)
          //**************************
          else if (nsm==3)
            {
              vchf[mes][den]=4;
              nchf[mes][den]=n;
              chf[mes][den]=d3;
              vihod[mes][den]=3;

              //общие суммы по часам
              ovchf[mes]+=4;
              onchf[mes]+=n;
              ochf[mes]+=d3;


              //проверка праздничного дня
              if (PrazdDni(den,mes)==true)
                {
                  pchf[mes][den]=d3;
                  opchf[mes]+=d3;
                }
              //рабочий день
              else
                {
                  pchf[mes][den]="NULL";
                }

              //проверка дня в смене
              if (dnism==4)
                {
                  nsm=0;
                  dnism=3;
                }
              else
                {
                  dnism++;
                }
            }
          //выходной
          //************************
          else
            {
              chf[mes][den]=0;
              nchf[mes][den]="NULL";
              vchf[mes][den]="NULL";
              pchf[mes][den]="NULL";
              vihod[mes][den]=0;


              //выходной перед 1 ночной сменой в последний день месяца
              if (dnism==0 && den==kol)
                {
                  chf[mes][den]=p1;
                  nchf[mes][den]=n1;
                  ochf[mes]+=p1;
                  onchf[mes]+=p1;

                  //проверка праздничного дня
                  if (PrazdDni(den,mes)==true)
                    {
                      pchf[mes][den]=p1;
                      opchf[mes]+=p1;
                    }
                }
              else if (dnism==0) //выходной перед 1 ночной сменой
                {

                  chf[mes][den]=p1;
                  nchf[mes][den]=n1;

                  //проверка праздничного дня
                  if (PrazdDni(den,mes)==true)
                    {
                      pchf[mes][den]=p1;
                      opchf[mes]+=p1;
               //       opchf[mes]+=p1;
                    }
                }


             /* if (den==kol && dnism==0)
                {
                  chf0[mes-1]=p2;
                  nchf0[mes-1]=6;
                }
              //часы переходящие c предыдущего месяца
            /*  if (den==1 && dnism==1 && mes==1)
                {
                  ochf[mes]+= Vvod->SetN(DM->qObnovlenie2->FieldByName("chf0")->AsString);
                  onchf[mes]+= Vvod->SetN(DM->qObnovlenie2->FieldByName("nch0")->AsString);
                  opchf[mes]+= Vvod->SetN(DM->qObnovlenie2->FieldByName("pch0")->AsString);
                }
              else if (den==1 && dnism==1)
                {
                  chf0[mes-1]=p2;
                  ochf[mes]+=p2;
                  nchf0[mes-1]=6;
                  onchf[mes]+=6;

                  //проверка праздничного дня
                  if (PrazdDni(den,mes)==true)
                    {
                      pchf0[mes-1]=p2;
                      opchf[mes]+=p2;
                    }
                }   */

               //проверка дня в смене
               if (dnism==1)
                 {
                   nsm=2;
                   dnism=1;
                 }
               else if (dnism==2)
                 {
                   nsm=3;
                   dnism=1;
                 }
               else if (dnism==3)
                 {
                   nsm=0;
                   dnism=0;
                 }
               else
                 {
                   nsm=1;
                   dnism=1;
                 }
            }
        }

      //расчет переработки
      if ((ochf[mes] - DM->qNorma11Graf->FieldByName("chf")->AsFloat - opchf[mes])>0)
        {
          pgraf[mes] = ochf[mes] - DM->qNorma11Graf->FieldByName("chf")->AsFloat - opchf[mes];
        }

      // сохранение переходящих часов последнего дня в году
      if ((mes==12 && dnism==0 && nsm==0)||(mes==12 && nsm==1 && dnism!=4))
        {
        /*  chf0[mes]=p2;
          nchf0[mes]=6;
          pchf0[mes]=p2;

        /*  chf0[mes]=p2;
          nchf0[mes]=6;
          pchf0[mes]=p2;*/

        }

      //отсутствующие дни в месяце
      if (den<32)
        {
          while (den<=32)
            {
              vihod[mes][den]="NULL";
              chf[mes][den]="NULL";
              nchf[mes][den]="NULL";
              vchf[mes][den]="NULL";
              pchf[mes][den]="NULL";
              den++;
            }
        }

      DM->qNorma11Graf->Next();
    }
}
//------------------------------------------------------------------------------

//Расчет 60 и 70 графика
void __fastcall TMain::Graf60(double p1, double p2, double v1, double v2, double n1, double n2)
{
  AnsiString kol;

  /*nsm - номер смены последнего месяца,
   dnism - день смены последнего месяца*/

  nsm = DM->qObnovlenie2->FieldByName("nsm")->AsInteger;
  dnism = DM->qObnovlenie2->FieldByName("dnism")->AsInteger;

  //по месяцам
  for (mes=1; mes<=12; mes++)
    {
      kol = DaysInAMonth(god, mes);

      //по дням
      for (den=1; den<=kol; den++)
        {
          //первая смена (7.30-19.30)
          //*************************
          if (nsm==1)
            {
              vihod[mes][den]=1;
              vchf[mes][den]=v1;
              nchf[mes][den]="NULL";
              chf[mes][den]=DM->qOgraf->FieldByName("DLIT")->AsFloat;

              //общие суммы по часам
              ovchf[mes]+=v1;
              ochf[mes]+= DM->qOgraf->FieldByName("DLIT")->AsFloat;

              //проверка праздничного дня
              if (PrazdDni(den,mes)==true)
                {
                  pchf[mes][den]=p1+p2;
                  opchf[mes]+=p1+p2;
                }

              dnism=1;
              nsm=2;
            }
          //вторая смена (19.30-7.30)
          //*************************
          else if (nsm==2)
            {
              vihod[mes][den]=2;

              //переход на летнее время (март)
              if (mes==3 && den==day_mart)
                {
                  vchf[mes][den]=v2;
                  nchf[mes][den]=n1;
                  chf[mes][den]=p1;

                  //общие суммы по часам
                  ovchf[mes]+=v2;
                  onchf[mes]+=(n1+n2)-1;
                  ochf[mes]+= DM->qOgraf->FieldByName("DLIT")->AsFloat-1;

                  //проверка праздничного дня
                  if (PrazdDni(den,mes)==true)
                    {
                      pchf[mes][den]=p1;
                      opchf[mes]+=p1;
                    }
                  //проверка предпраздничного дня
                 /* else if (PrdPrazdDni(den,mes)==true)
                    {
                      pchf[mes][den]=p2-1;
                      opchf[mes]+=p2-1;
                    } */
                }
              //переход на зимнее время (октябрь)
              else if (mes==10 && den==day_oktyabr)
                {
                  vchf[mes][den]=v2;
                  nchf[mes][den]=n1;
                  chf[mes][den]=p1;

                  //общие суммы по часам
                  ovchf[mes]+=v2;
                  onchf[mes]+=(n1+n2)+1;
                  ochf[mes]+= DM->qOgraf->FieldByName("DLIT")->AsFloat+1;

                  //проверка праздничного дня
                  if (PrazdDni(den,mes)==true)
                    {
                      pchf[mes][den]=p1;
                      opchf[mes]+=p1;
                    }
                  //проверка предпраздничного дня
                  /*else if (PrdPrazdDni(den,mes)==true)
                    {
                      pchf[mes][den]=p2+1;
                      opchf[mes]+=p2+1;
                    }*/
                }
              else
                {
                  //если ночная смена попадает на последний день месяца
                  if (den==kol)
                    {
                      vchf[mes][den]=v2;
                      nchf[mes][den]=n1;
                      chf[mes][den]=p1;

                      //общие суммы по часам
                      onchf[mes]+=n1;
                      ovchf[mes]+=v2;
                      ochf[mes]+=p1;

                      //проверка праздничного дня
                      if (PrazdDni(den,mes)==true)
                        {
                          pchf[mes][den]=p1;
                          opchf[mes]+=p1;
                        }

                     /* //проверка праздничного и предпраздничного дня (1 мая)
                      if (PrazdDni(den,mes)==true && PrdPrazdDni(den,mes)==true)
                        {
                          pchf[mes][den]=p1;
                          opchf[mes]+=p1;
                        }
                      else
                        {
                          //проверка праздничного дня
                          if (PrazdDni(den,mes)==true)
                            {
                              pchf[mes][den]=p1;
                              opchf[mes]+=p1;
                            }
                          //проверка предпраздничного дня
                          else if (PrdPrazdDni(den,mes)==true)
                            {
                              pchf[mes][den]=p2;
                            }
                       }  */
                    }
                  else
                    {
                      vchf[mes][den]=v2;
                      nchf[mes][den]=n1;
                      chf[mes][den]=p1;

                      //общие суммы по часам
                      onchf[mes]+=(n1+n2);
                      ovchf[mes]+=v2;
                      ochf[mes]+=DM->qOgraf->FieldByName("DLIT")->AsFloat;

                      //проверка праздничного дня
                      if (PrazdDni(den,mes)==true)
                        {
                          pchf[mes][den]=p1;
                          opchf[mes]+=p1;
                        }
                      //проверка праздничного и предпраздничного дня
                     /* if (PrazdDni(den,mes)==true && PrdPrazdDni(den,mes)==true)
                        {
                          pchf[mes][den]=p1+p2;
                          opchf[mes]+=p1+p2;
                        }
                      else
                        {
                          //проверка праздничного дня
                          if (PrazdDni(den,mes)==true)
                            {
                              pchf[mes][den]=p1;
                              opchf[mes]+=p1;
                            }
                          //проверка предпраздничного дня
                          else if (PrdPrazdDni(den,mes)==true)
                            {
                              pchf[mes][den]=p2;
                              opchf[mes]+=p2;
                            }
                        } */
                    }
                }

              dnism=1;
              nsm=0;
            }

          //выходной
          //************************
          else
            {
              nchf[mes][den]="NULL";
              vchf[mes][den]="NULL";
              pchf[mes][den]="NULL";
              vihod[mes][den]=0;

              if (dnism==1)
                {
                  if (mes==3 && den==day_mart2)
                    {
                      chf[mes][den]=p2-1;
                      nchf[mes][den]=n2-1;
                      //проверка праздничного дня
                      if (PrazdDni(den,mes)==true)
                        {
                          pchf[mes][den]=p2-1;
                          opchf[mes]+=p2-1;
                        }
                    }
                  //переход на зимнее время (октябрь)
                  else if (mes==10 && den==day_oktyabr2)
                    {
                      chf[mes][den]=p2+1;
                      nchf[mes][den]=n2+1;
                      //проверка праздничного дня
                      if (PrazdDni(den,mes)==true)
                        {
                          pchf[mes][den]=p2+1;
                          opchf[mes]+=p2+1;
                        }
                    }
                  else
                    {
                      chf[mes][den]=p2;
                      nchf[mes][den]=n2;
                      //проверка праздничного дня
                      if (PrazdDni(den,mes)==true)
                        {
                          pchf[mes][den]=p2;
                          opchf[mes]+=p2;
                        }
                    }
                }
              else
                {
                  chf[mes][den]=0;
                  nchf[mes][den]="NULL";
                }

              //часы переходящие c предыдущего месяца
              if (den==1 && dnism==1 && mes==1)
                {
                  ochf[mes]+= Vvod->SetN(DM->qObnovlenie2->FieldByName("chf0")->AsString);
                  onchf[mes]+= Vvod->SetN(DM->qObnovlenie2->FieldByName("nch0")->AsString);
                //  opchf[mes]+= Vvod->SetN(DM->qObnovlenie2->FieldByName("pch0")->AsString);
                }
              else if (den==1 && dnism==1)
                {
                  chf0[mes-1]=p2;
                  ochf[mes]+=p2;
                  nchf0[mes-1]=n2;
                  onchf[mes]+=n2;

                  //проверка праздничного дня
                  if (PrazdDni(den,mes)==true)
                    {
                      pchf0[mes-1]=p2;
                    //  opchf[mes]+=p2;
                    }
                }

              //проверка дня в смене
              if (dnism==2)
                {
                  nsm=1;
                  dnism=1;
                }
              else
                {
                  dnism++;
                }
            }
        }



      //расчет переработки
      if ((ochf[mes] - DM->qNorma11Graf->FieldByName("chf")->AsFloat - opchf[mes])>0 && graf!=170 && graf!=50)
        {
          pgraf[mes] = ochf[mes] - DM->qNorma11Graf->FieldByName("chf")->AsFloat - opchf[mes];
        }

      // сохранение переходящих часов последнего дня в году
      if (mes==12 && dnism==1 && nsm==0)
        {
          chf0[mes]=p2;
          nchf0[mes]=n2;
          pchf0[mes]=p2;
        }

      //отсутствующие дни в месяце
      if (den<32)
        {
          while (den<=32)
            {
              vihod[mes][den]="NULL";
              chf[mes][den]="NULL";
              nchf[mes][den]="NULL";
              vchf[mes][den]="NULL";
              pchf[mes][den]="NULL";
              den++;
            }
        }

      DM->qNorma11Graf->Next();
    }
}
//------------------------------------------------------------------------------

//Расчет 4060
void __fastcall TMain::Graf4060(double p, double p1, double p2, double v1, double v2, double n1, double n2)
{
  AnsiString kol;

  /*nsm - номер смены последнего месяца,
   dnism - день смены последнего месяца*/

  nsm = DM->qObnovlenie2->FieldByName("nsm")->AsInteger;
  dnism = DM->qObnovlenie2->FieldByName("dnism")->AsInteger;

  //по месяцам
  for (mes=1; mes<=12; mes++)
    {
      kol = DaysInAMonth(god, mes);

      if (mes==4 || mes==5 || mes==6 || mes==7 || mes==8 || mes==9)
        {
          p=12, p1=5, p2=7;
          v1=1, v2=3;
          n1=2; n2=6;
        }
      else
        {
          p=10, p1=p2=7;
          v1=NULL, v2=4;
          n1=2; n2=6;
        }

      //по дням
      for (den=1; den<=kol; den++)
        {
          //первая смена (7.30-19.30)
          //*************************
          if (nsm==1)
            {
              vihod[mes][den]=1;
              vchf[mes][den]=v1;
              nchf[mes][den]="NULL";
              chf[mes][den]=p;

              //общие суммы по часам
              ovchf[mes]+=v1;
              ochf[mes]+=p;

              //проверка праздничного дня
              if (PrazdDni(den,mes)==true)
                {
                  pchf[mes][den]=p;
                  opchf[mes]+=p;
                }

              dnism=1;
              nsm=2;
            }
          //вторая смена (19.30-7.30)
          //*************************
          else if (nsm==2)
            {
              vihod[mes][den]=2;

              //переход на летнее время (март)
              if (mes==3 && den==day_mart)
                {
                  vchf[mes][den]=v2;
                  nchf[mes][den]=n1;
                  chf[mes][den]=p1;

                  //общие суммы по часам
                  ovchf[mes]+=v2;
                  onchf[mes]+=(n1+n2)-1;
                  ochf[mes]+= (p1+p2)-1;

                  //проверка праздничного дня
                  if (PrazdDni(den,mes)==true)
                    {
                      pchf[mes][den]=p1;
                      opchf[mes]+=p1;
                    }
                 /* //проверка предпраздничного дня
                  else if (PrdPrazdDni(den,mes)==true)
                    {
                      pchf[mes][den]=p2-1;
                      opchf[mes]+=p2-1;
                    }  */
                }
              //переход на зимнее время (октябрь)
              else if (mes==10 && den==day_oktyabr)
                {
                  vchf[mes][den]=v2;
                  nchf[mes][den]=n1;
                  chf[mes][den]=p1;

                  //общие суммы по часам
                  ovchf[mes]+=v2;
                  onchf[mes]+=(n1+n2)+1;
                  ochf[mes]+=(p1+p2)+1;

                  //проверка праздничного дня
                  if (PrazdDni(den,mes)==true)
                    {
                      pchf[mes][den]=p1;
                      opchf[mes]+=p1;
                    }
                  //проверка предпраздничного дня
                /*  else if (PrdPrazdDni(den,mes)==true)
                    {
                      pchf[mes][den]=p2+1;
                      opchf[mes]+=p2+1;
                    } */
                }
              else
                {
                  //если ночная смена попадает на последний день месяца
                  if (den==kol)
                    {
                      vchf[mes][den]=v2;
                      nchf[mes][den]=n1;
                      chf[mes][den]=p1;

                      //общие суммы по часам
                      onchf[mes]+=n1;
                      ovchf[mes]+=v2;
                      ochf[mes]+=p1;

                      //проверка праздничного дня
                      if (PrazdDni(den,mes)==true)
                        {
                          pchf[mes][den]=p1;
                          opchf[mes]+=p1;
                        }

                      //проверка праздничного и предпраздничного дня (1 мая)
                     /* if (PrazdDni(den,mes)==true && PrdPrazdDni(den,mes)==true)
                        {
                          pchf[mes][den]=p1;
                          opchf[mes]+=p1;
                        }
                      else
                        {
                          //проверка праздничного дня
                          if (PrazdDni(den,mes)==true)
                            {
                              pchf[mes][den]=p1;
                              opchf[mes]+=p1;
                            }
                          //проверка предпраздничного дня
                          else if (PrdPrazdDni(den,mes)==true)
                            {
                              pchf[mes][den]=p2;
                            }
                       } */
                    }
                  else
                    {
                      vchf[mes][den]=v2;
                      nchf[mes][den]=n1;
                      chf[mes][den]=p1;

                      //общие суммы по часам
                      onchf[mes]+=(n1+n2);
                      ovchf[mes]+=v2;
                      ochf[mes]+=(p1+p2);

                     //проверка праздничного дня
                     if (PrazdDni(den,mes)==true)
                       {
                         pchf[mes][den]=p1;
                         opchf[mes]+=p1;
                       }
                      //проверка праздничного и предпраздничного дня
                     /* if (PrazdDni(den,mes)==true && PrdPrazdDni(den,mes)==true)
                        {
                          pchf[mes][den]=p1+p2;
                          opchf[mes]+=p1+p2;
                        }
                      else
                        {
                          //проверка праздничного дня
                          if (PrazdDni(den,mes)==true)
                            {
                              pchf[mes][den]=p1;
                              opchf[mes]+=p1;
                            }
                          //проверка предпраздничного дня
                          else if (PrdPrazdDni(den,mes)==true)
                            {
                              pchf[mes][den]=p2;
                              opchf[mes]+=p2;
                            }
                        }  */
                    }
                }

              dnism=1;
              nsm=0;
            }

          //выходной
          //************************
          else
            {
              vchf[mes][den]="NULL";
              pchf[mes][den]="NULL";
              vihod[mes][den]=0;

              if(dnism==1)
                {
                  if(mes==3 && den==day_mart2)
                    {
                      chf[mes][den]=p2-1;
                      nchf[mes][den]=n2-1;
                      //проверка праздничного дня
                      if (PrazdDni(den,mes)==true)
                        {
                          pchf[mes][den]=p2-1;
                          opchf[mes]+=p2-1;
                        }
                    }
                  else if (mes==10 && den==day_oktyabr2)
                    {
                      chf[mes][den]=p2+1;
                      nchf[mes][den]=n2+1;
                      //проверка праздничного дня
                      if (PrazdDni(den,mes)==true)
                        {
                          pchf[mes][den]=p2+1;
                          opchf[mes]+=p2+1;
                        }
                    }
                  else
                    {
                      chf[mes][den]=p2;
                      nchf[mes][den]=n2;
                      //проверка праздничного дня
                      if (PrazdDni(den,mes)==true)
                        {
                          pchf[mes][den]=p2;
                          opchf[mes]+=p2;
                        }
                    }
                }
              else
                {
                  chf[mes][den]=0;
                  nchf[mes][den]="NULL";
                }


              //часы переходящие c предыдущего месяца
              if (den==1 && dnism==1 && mes==1)
                {
                  ochf[mes]+= Vvod->SetN(DM->qObnovlenie2->FieldByName("chf0")->AsString);
                  onchf[mes]+= Vvod->SetN(DM->qObnovlenie2->FieldByName("nch0")->AsString);
               //   opchf[mes]+= Vvod->SetN(DM->qObnovlenie2->FieldByName("pch0")->AsString);
                }
              else if (den==1 && dnism==1)
                {
                  chf0[mes-1]=p2;
                  ochf[mes]+=p2;
                  nchf0[mes-1]=n2;
                  onchf[mes]+=n2;

                  //проверка праздничного дня
                  if (PrazdDni(den,mes)==true)
                    {
                      pchf0[mes-1]=p2;
                 //     opchf[mes]+=p2;
                    }
                }

              //проверка дня в смене
              if (dnism==2)
                {
                  nsm=1;
                  dnism=1;
                }
              else
                {
                  dnism++;
                }
            }
        }

      //расчет переработки
      if ((ochf[mes] - DM->qNorma11Graf->FieldByName("chf")->AsFloat - opchf[mes])>0)
        {
          pgraf[mes] = ochf[mes] - DM->qNorma11Graf->FieldByName("chf")->AsFloat - opchf[mes];
        }

      // сохранение переходящих часов последнего дня в году
      if (mes==12 && dnism==1 && nsm==0)
        {
          chf0[mes]=p2;
          nchf0[mes]=n2;
          pchf0[mes]=p2;
        }

      //отсутствующие дни в месяце
      if (den<32)
        {
          while (den<=32)
            {
              vihod[mes][den]="NULL";
              chf[mes][den]="NULL";
              nchf[mes][den]="NULL";
              vchf[mes][den]="NULL";
              pchf[mes][den]="NULL";
              den++;
            }
        }

      DM->qNorma11Graf->Next();
    }
}
//------------------------------------------------------------------------------

//Расчет 85 графика
void __fastcall TMain::Graf85(double n)
{
  int kol;

  /* chf[32] - рабочие часы по дням
     chf[den] = 8 - рабочий день
     chf[den] = 7 - предпраздничный день
     vihod[32] - выходы по дням (рабочий, отдых, праздничный)
     vihod[den] = 1 - рабочий день
     vihod[den] = 9 - праздник
  */

  //по месяцам
  for (mes=1; mes<=12; mes++)
    {
      kol = DaysInAMonth(god, mes);

      //по дням
      for (den=1; den<=kol; den++)
        {
          //проверка дня недели
          if (DayWeek(den,mes,god)==1||DayWeek(den,mes,god)==2)
            {
              //проверка праздничного дня
              if (PrazdDni(den,mes)==true)
                {
                  vihod[mes][den]=9;
                  chf[mes][den]=30;
                }
              else
                {
                  vihod[mes][den]=0;
                  chf[mes][den]=0;
                }
            }
          else
            {
              //проверка праздничного дня
              if (PrazdDni(den,mes)==true)
                {
                  vihod[mes][den]=9;
                  chf[mes][den]=30;
                }

              //проверка предпраздничного дня
              else if (PrdPrazdDni(den,mes)==true)
                {
                  vihod[mes][den]=1;
                  chf[mes][den]=DM->qOgraf->FieldByName("DLIT")->AsFloat-1;
                  nchf[mes][den]=n;

                  ochf[mes]+= DM->qOgraf->FieldByName("DLIT")->AsFloat-1;
                  onchf[mes]+=n;
                }
              //рабочий день
              else
                {
                  vihod[mes][den]=1;
                  chf[mes][den]=DM->qOgraf->FieldByName("DLIT")->AsFloat;
                  nchf[mes][den]=n;

                  ochf[mes]+=DM->qOgraf->FieldByName("DLIT")->AsFloat;
                  onchf[mes]+=n;
                }
            }

        }
      //расчет переработки
      if ((ochf[mes] - DM->qNorma11Graf->FieldByName("chf")->AsFloat - opchf[mes])>0)
        {
          pgraf[mes] = ochf[mes] - DM->qNorma11Graf->FieldByName("chf")->AsFloat - opchf[mes];
        }


      //отсутствующие дни в месяце
      if (den<32)
        {
          while (den<=32)
            {
              vihod[mes][den]="NULL";
              chf[mes][den]="NULL";
              den++;
            }
        }
       DM->qNorma11Graf->Next();
    }
}
//---------------------------------------------------------------------------

//Расчет 90 графика
void __fastcall TMain::Graf90(double d1, double d2, double d3, double p1, double p2, double p, double v, double n1, double n2, double n)
{
  AnsiString kol;

  /*nsm - номер смены последнего месяца,
   dnism - день смены последнего месяца*/

  nsm = DM->qObnovlenie2->FieldByName("nsm")->AsInteger;
  dnism = DM->qObnovlenie2->FieldByName("dnism")->AsInteger;

  //по месяцам
  for (mes=1; mes<=12; mes++)
    {
      kol = DaysInAMonth(god, mes);

      //по дням
      for (den=1; den<=kol; den++)
        {
          //первая смена (23.00-7.00)
          //*************************
          if (nsm==1)
            {
              vihod[mes][den]=1;

              //переход на летнее время (март)
              if (mes==3 && den==day_mart && dnism!=3)
                {
                  chf[mes][den]=p1+p2;
                  nchf[mes][den]=n1+n2;

                  //общие суммы по часам
                  onchf[mes]+=(n1+n2)-1;
                  ochf[mes]+=d1-1;

                  //проверка праздничного дня
                  if (PrazdDni(den,mes)==true)
                    {
                      pchf[mes][den]=p1+p2;
                      opchf[mes]+=p1+p2;
                    }
                  //проверка праздничного дня (праздничные проставляются в день когда пришел со смены)
                 /* if (PrazdDni(den,mes)==true)
                    {
                      if (dnism==3)
                        {
                          if (pchf[mes][den]!="NULL") pchf[mes][den]=FloatToStr(StrToFloat(pchf[mes][den])+p2);
                          else pchf[mes][den]=p2;
                          opchf[mes]+=p2;
                        }
                      else
                        {
                          if (pchf[mes][den]!="NULL") pchf[mes][den]=FloatToStr(StrToFloat(pchf[mes][den])+p2);
                          else pchf[mes][den]=p2;
                          opchf[mes]+=p1+p2;

                          //если не последняя переходящая смена
                          if (den!=kol)
                            {
                              pchf[mes][den+1]=p1;
                            }
                          else
                            {
                              // если последняя смена =(((((((
                              pchf[mes+1][1]=p1;
                            }
                        }

                    } */
                }
              else if (mes==mes_mart2 && den==day_mart2 && dnism==1)
                {
                  chf[mes][den]=(p1+p2)-1;
                  nchf[mes][den]=(n1+n2)-1;

                  //общие суммы по часам
                  onchf[mes]+=(n1+n2)-1;
                  ochf[mes]+=d1-1;

                  //проверка праздничного дня
                  if (PrazdDni(den,mes)==true)
                    {
                      pchf[mes][den]=(p1+p2)-1;
                      opchf[mes]+=(p1+p2)-1;
                    }

                  //проверка праздничного дня (праздничные проставляются в день когда пришел со смены)
                  /*if (PrazdDni(den,mes)==true )
                    {
                      if (pchf[mes][den]!="NULL") pchf[mes][den]=FloatToStr(StrToFloat(pchf[mes][den])+p2-1);
                      else pchf[mes][den]=p2-1;
                      opchf[mes]+=p1+p2-1;

                      //если не последняя переходящая смена
                      if (den!=kol)
                        {
                          pchf[mes][den+1]=p1;
                        }
                      else
                        {
                          // если последняя смена =(((((((
                          pchf[mes+1][1]=p1;
                        }
                    }  */
                }
              //переход на зимнее время (октябрь)
              else if (mes==10 && den==day_oktyabr && dnism!=3)
                {
                  chf[mes][den]=p1+p2;
                  nchf[mes][den]=n1+n2;

                  //общие суммы по часам
                  onchf[mes]+=(n1+n2)+1;
                  ochf[mes]+=d1+1;

                  //проверка праздничного дня
                  if (PrazdDni(den,mes)==true)
                    {
                      pchf[mes][den]=p1+p2;
                      opchf[mes]+=p1+p2;
                    }

                  //проверка праздничного дня (праздничные проставляются в день когда пришел со смены)
                 /* if (PrazdDni(den,mes)==true)
                    {
                      if (dnism==3)
                        {
                          if (pchf[mes][den]!="NULL") pchf[mes][den]=FloatToStr(StrToFloat(pchf[mes][den])+p2);
                          else pchf[mes][den]=p2;
                          opchf[mes]+=p2;
                        }
                      else
                        {
                          if (pchf[mes][den]!="NULL") pchf[mes][den]=FloatToStr(StrToFloat(pchf[mes][den])+p2);
                          else pchf[mes][den]=p2;
                          opchf[mes]+=p1+p2;

                          //если не последняя переходящая смена
                          if (den!=kol)
                            {
                              pchf[mes][den+1]=p1;
                            }
                          else
                            {
                              // если последняя смена =(((((((
                              pchf[mes+1][1]=p1;
                            }
                        }

                    } */
                }
              else if (mes==mes_oktyabr2 && den==day_oktyabr2 && dnism==1)
                {
                  chf[mes][den]=(p1+p2)+1;
                  nchf[mes][den]=(n1+n2)+1;

                  //общие суммы по часам
                  onchf[mes]+=(n1+n2)+1;
                  ochf[mes]+=d1+1;

                  //проверка праздничного дня
                  if (PrazdDni(den,mes)==true)
                    {
                      pchf[mes][den]=(p1+p2)+1;
                      opchf[mes]+=(p1+p2)+1;
                    }

                  //проверка праздничного дня (праздничные проставляются в день когда пришел со смены)
                  /*if (PrazdDni(den,mes)==true)
                    {
                      if (pchf[mes][den]!="NULL") pchf[mes][den]=FloatToStr(StrToFloat(pchf[mes][den])+p2+1);
                      else pchf[mes][den]=p2+1;
                      opchf[mes]+=p1+p2+1;

                      //если не последняя переходящая смена
                      if (den!=kol)
                        {
                          pchf[mes][den+1]=p1;
                        }
                      else
                        {
                          // если последняя смена =(((((((
                          pchf[mes+1][1]=p1;
                        }
                    }*/
                }
              else
                {
                  if (mes==mes_mart2 && den==day_mart2 && dnism==3)
                    {
                      chf[mes][den]=p2-1;
                      nchf[mes][den]=n2-1;
                      //проверка праздничного дня
                      if (PrazdDni(den,mes)==true)
                        {
                          pchf[mes][den]=p2-1;
                          opchf[mes]+=p2-1;
                        }
                    }
                  else if (mes==mes_oktyabr2 && den==day_oktyabr2 && dnism==3)
                    {
                      chf[mes][den]=p2+1;
                      nchf[mes][den]=n2+1;
                      //проверка праздничного дня
                      if (PrazdDni(den,mes)==true)
                        {
                          pchf[mes][den]=p2+1;
                          opchf[mes]+=p2+1;
                        }
                    }
                  else if (dnism==3)
                    {
                      chf[mes][den]=p2;
                      nchf[mes][den]=n2;
                      //проверка праздничного дня
                      if (PrazdDni(den,mes)==true)
                        {
                          pchf[mes][den]=p2;
                          opchf[mes]+=p2;
                        }
                    }
                  else
                    {
                      chf[mes][den]=p1+p2;
                      nchf[mes][den]=n1+n2;
                      //проверка праздничного дня
                      if (PrazdDni(den,mes)==true)
                        {
                          pchf[mes][den]=p1+p2;
                          opchf[mes]+=p1+p2;
                        }
                    }

                  //общие суммы по часам
                  onchf[mes]+=(n1+n2);
                  ochf[mes]+=d1;

                  //проверка праздничного дня (праздничные проставляются в день когда пришел со смены)
                 /* if (PrazdDni(den,mes)==true )
                    {
                      if (dnism==3)
                        {
                          if (pchf[mes][den]!="NULL") pchf[mes][den]=FloatToStr(StrToFloat(pchf[mes][den])+p2);
                          else pchf[mes][den]=p2;
                          opchf[mes]+=p2;
                        }
                      else
                        {
                          if (pchf[mes][den]!="NULL") pchf[mes][den]=FloatToStr(StrToFloat(pchf[mes][den])+p2);
                          else pchf[mes][den]=p2;
                          opchf[mes]+=p1+p2;

                          //если не последняя переходящая смена
                          if (den!=kol)
                            {
                              pchf[mes][den+1]=p1;
                            }
                          else
                            {
                              // если последняя смена =(((((((
                              pchf[mes+1][1]=p1;
                            }
                        }
                    }*/
                }

              if (den==1)
                {
                  ochf[mes]-=p1;
                  onchf[mes]-=p1;

                  if (mes!=1)
                    {
                      chf0[mes-1]=p2;
                      nchf0[mes-1]=n2;

                      //проверка праздничного дня
                      if (PrazdDni(den,mes)==true)
                        {
                          pchf0[mes-1]=p2;
                        }
                    }
                }
              else if (den==kol && dnism!=3)
                {
                  ochf[mes]+=p1;
                  onchf[mes]+=p1;
                 
                }

              //проверка дня в смене
              if (dnism==3)
                {
                  nsm=0;
                  dnism=1;
                }
              else
                {
                  dnism++;
                }
            }

                  /*
                  //последний день не учитывается в общих суммах
                  if (den==kol)
                    {
                      chf[mes][den]=d1;
                      nchf[mes][den]=n1;

                     //общие суммы по часам
                     onchf[mes]+=p1;
                     ochf[mes]+=p1;
                    }
                  else
                    {
                      chf[mes][den]=d1;
                      nchf[mes][den]=n1;

                      //общие суммы по часам
                      onchf[mes]+=n1;
                      ochf[mes]+=d1;
                    }

                  //проверка и праздничного и предпраздничного дня (1 мая)
                  if (PrazdDni(den,mes)==true  && PrdPrazdDni(den,mes)==true)
                    {
                      pchf[mes][den]=p1+p2;
                      opchf[mes]+=p1+p2;
                    }
                  else
                    {
                      //проверка праздничного дня
                      if (PrazdDni(den,mes)==true)
                        {
                          pchf[mes][den]=p1;
                          opchf[mes]+=p1;
                        }
                      //проверка предпраздничного дня
                      else if (PrdPrazdDni(den,mes)==true)
                        {
                          //если последний день, в общие праздничные часы не учитываются
                          if (den==kol)
                            {
                              pchf[mes][den]=p2;
                            }
                          else
                            {
                              pchf[mes][den]=p2;
                              opchf[mes]+=p2;
                            }
                         }
                    }
                }

              //часы переходящие c предыдущего месяца
              if (den==1 && mes==1 && dnism!=1)
                {
                  ochf[mes]+= Vvod->SetN(DM->qObnovlenie2->FieldByName("chf0")->AsString);
                  onchf[mes]+= Vvod->SetN(DM->qObnovlenie2->FieldByName("nch0")->AsString);
                  opchf[mes]+= Vvod->SetN(DM->qObnovlenie2->FieldByName("pch0")->AsString);
                }
              else if (den==1 && dnism!=1)
                {
                  chf0[mes-1]=p2;
                  ochf[mes]+=p2;
                  nchf0[mes-1]=6;
                  onchf[mes]+=6;


                  //проверка праздничного дня
                  if (PrazdDni(den,mes)==true)
                    {
                      pchf0[mes-1]=p2;
                      opchf[mes]+=p2;
                    }
                }

              //проверка дня в смене
              if (dnism==3)
                {
                  nsm=0;
                  dnism=1;
                }
              else
                {
                  dnism++;
                }
            }    */
          //вторая смена (7.00-15.00)
          //*************************
          else if (nsm==2)
            {
              chf[mes][den]=d2;
              vihod[mes][den]=2;

              //общие суммы по часам
              ochf[mes]+=d2;


              //проверка праздничного дня
              if (PrazdDni(den,mes)==true)
                {
                  pchf[mes][den]=d2;
                  opchf[mes]+=d2;
                }

              //проверка дня в смене
              if (dnism==3)
                {
                  nsm=0;
                  dnism=3;
                }
              else
                {
                  dnism++;
                }
            }
          //третья смена (15.00-23.00)
          //**************************
          else if (nsm==3)
            {
              vihod[mes][den]=3;
              chf[mes][den]=d3;
              vchf[mes][den]=v;
              nchf[mes][den]=n;


              //общие суммы по часам
              ovchf[mes]+=v;
              onchf[mes]+=n;
              ochf[mes]+=d3;


              //проверка праздничного дня
              if (PrazdDni(den,mes)==true)
                {
                  pchf[mes][den]=p;
                  opchf[mes]+=p;
                }

              //проверка дня в смене
              if (dnism==3)
                {
                  nsm=0;
                  dnism=5;
                }
              else
                {
                  dnism++;
                }
            }
          //выходной
          //************************
          else
            {
              if (dnism==6)
                {
                  chf[mes][den]=p1;
                  nchf[mes][den]=n1;
                  //проверка праздничного дня
                  if (PrazdDni(den,mes)==true)
                    {
                      pchf[mes][den]=p1;
                      opchf[mes]+=p1;
                    }
                }
              else
                {
                  chf[mes][den]=0;
                  nchf[mes][den]="NULL";
                }

              vihod[mes][den]=0;


              //выходной перед 1 ночной сменой
              if (dnism==6 && den==kol)
                {
                  ochf[mes]+=p1;
                  onchf[mes]+=p1;

                  //проверка праздничного дня
                  if (PrazdDni(den,mes)==true)
                    {
                      pchf[mes][den]=p1;
                      opchf[mes]+=p1;
                    }
                }
             /* else if (dnism==6)
                {
                  //проверка праздничного дня
                  if (PrazdDni(den,mes)==true)
                    {
                      pchf[mes][den]=p1;
                      opchf[mes]+=p1;
                    }
                }*/

               //проверка дня в смене
               if (dnism==2)
                 {
                   nsm=2;
                   dnism=1;
                 }
               else if (dnism==4)
                 {
                   nsm=3;
                   dnism=1;
                 }
               else if (dnism==6)
                 {
                   nsm=1;
                   dnism=1;
                 }
               else
                 {
                   dnism++;
                 }
            }
        }

      // сохранение переходящих часов последнего дня в году
      if ((mes==12 && nsm==1))
        {
          chf0[mes]=p2;
          nchf0[mes]=n2;
          pchf0[mes]=p2;
        }

      //отсутствующие дни в месяце
      if (den<32)
        {
          while (den<=32)
            {
              vihod[mes][den]="NULL";
              chf[mes][den]="NULL";
              nchf[mes][den]="NULL";
              vchf[mes][den]="NULL";
              pchf[mes][den]="NULL";
              den++;
            }
        }

      DM->qNorma11Graf->Next();
    }

}
//------------------------------------------------------------------------------

//Расчет 120 графика
void __fastcall TMain::Graf120(double d1, double d2, double d3, double p1, double p2, double v, double n1, double n2, double n)
{
  AnsiString kol;

  /*nsm - номер смены последнего месяца,
   dnism - день смены последнего месяца*/

  nsm = DM->qObnovlenie2->FieldByName("nsm")->AsInteger;
  dnism = DM->qObnovlenie2->FieldByName("dnism")->AsInteger;

  //по месяцам
  for (mes=1; mes<=12; mes++)
    {
      kol = DaysInAMonth(god, mes);

      //по дням
      for (den=1; den<=kol; den++)
        {
          //первая смена (23.00-7.00)
          //*************************
          if (nsm==1)
            {
              vihod[mes][den]=1;

              //переход на летнее время (март)
              if (mes==mes_mart2 && den==day_mart2)
                {
                  if (dnism==2)
                    {
                      chf[mes][den]=p2-1;
                      nchf[mes][den]=n2-1;
                      //проверка праздничного дня
                      if (PrazdDni(den,mes)==true)
                        {
                          pchf[mes][den]=p2-1;
                          opchf[mes]+=p2-1;
                        }

                      //общие суммы по часам
                      ochf[mes]+=p2-1;
                      onchf[mes]+=n2-1;
                    }
                  else
                    {
                      chf[mes][den]=p1;
                      nchf[mes][den]=n1;
                      //проверка праздничного дня
                      if (PrazdDni(den,mes)==true)
                        {
                          pchf[mes][den]=p1;
                          opchf[mes]+=p1;
                        }

                      //общие суммы по часам
                      ochf[mes]+=p1;
                      onchf[mes]+=n1;

                      /*chf[mes][den]=d1-1;
                      nchf[mes][den]=(n1+n2)-1;
                      //проверка праздничного дня
                      if (PrazdDni(den,mes)==true)
                        {
                          pchf[mes][den]=d1-1;
                          opchf[mes]+=d1-1;
                        }
                      //общие суммы по часам
                      ochf[mes]+=d1-1;
                      onchf[mes]+=(n1+n2)-1;  */
                    }
                }
              //переход на зимнее время (октябрь)
              else if (mes==mes_oktyabr2 && den==day_oktyabr2)
                {
                  if (dnism==2)
                    {
                      chf[mes][den]=p2+1;
                      nchf[mes][den]=n2+1;
                      //проверка праздничного дня
                      if (PrazdDni(den,mes)==true)
                        {
                          pchf[mes][den]=p2+1;
                          opchf[mes]+=p2+1;
                        }
                      //общие суммы по часам
                      ochf[mes]+=p2+1;
                      onchf[mes]+=n2+1;
                    }
                  else
                    {
                      chf[mes][den]=p1;
                      nchf[mes][den]=n1;
                      //проверка праздничного дня
                      if (PrazdDni(den,mes)==true)
                        {
                          pchf[mes][den]=p1;
                          opchf[mes]+=p1;
                        }

                      //общие суммы по часам
                      ochf[mes]+=p1;
                      onchf[mes]+=n1;

                      /*chf[mes][den]=d1+1;
                      nchf[mes][den]=(n1+n2)+1;
                      //проверка праздничного дня
                      if (PrazdDni(den,mes)==true)
                        {
                          pchf[mes][den]=d1+1;
                          opchf[mes]+=d1+1;
                        }
                      //общие суммы по часам
                      ochf[mes]+=d1+1;
                      onchf[mes]+=(n1+n2)+1;*/
                    }

                }
              else
                {
                  if (dnism==1)
                        {
                          chf[mes][den]=p1;
                          nchf[mes][den]=n1;
                          //проверка праздничного дня
                          if (PrazdDni(den,mes)==true)
                            {
                              pchf[mes][den]=p1;
                              opchf[mes]+=p1;
                            }

                          //общие суммы по часам
                          ochf[mes]+=p1;
                          onchf[mes]+=n1;
                        }
                      else
                        {
                          chf[mes][den]=d1;
                          nchf[mes][den]=(n1+n2);
                          //проверка праздничного дня
                          if (PrazdDni(den,mes)==true)
                            {
                              pchf[mes][den]=d1;
                              opchf[mes]+=d1;
                            }

                          //общие суммы по часам
                          ochf[mes]+=d1;
                          onchf[mes]+=(n1+n2);
                        }
                }

              //часы переходящие c предыдущего месяца
              if (den==1 && mes==1 && dnism==2)
                {
              //    ochf[mes]+= Vvod->SetN(DM->qObnovlenie2->FieldByName("chf0")->AsString);
               //   onchf[mes]+= Vvod->SetN(DM->qObnovlenie2->FieldByName("nch0")->AsString);
                //  opchf[mes]+= Vvod->SetN(DM->qObnovlenie2->FieldByName("pch0")->AsString);
                }
              else if (den==1 && dnism==2)
                {
                  chf0[mes-1]=p2;
                 // ochf[mes]+=p2;
                  nchf0[mes-1]=n2;
                 // onchf[mes]+=n2;

                  //проверка праздничного дня
                  if (PrazdDni(den,mes)==true)
                    {
                      pchf0[mes-1]=p2;
                    //  opchf[mes]+=p2;
                    }
                }

              //проверка дня в смене
              if (dnism==2)
                {
                  nsm=0;
                  dnism=1;
                }
              else
                {
                  dnism++;
                }
            }
          //вторая смена (7.00-15.00)
          //*************************
          else if (nsm==2)
            {
              vihod[mes][den]=2;
              chf[mes][den]=d2;

              //общие суммы по часам
              ochf[mes]+=d2;


              //проверка праздничного дня
              if (PrazdDni(den,mes)==true)
                {
                  pchf[mes][den]=d2;
                  opchf[mes]+=d2;
                }

              //проверка дня в смене
              if (dnism==2)
                {
                  nsm=0;
                  dnism=3;
                }
              else
                {
                  dnism++;
                }
            }
          //третья смена (15.00-23.00)
          //**************************
          else if (nsm==3)
            {
              vihod[mes][den]=3;
              chf[mes][den]=d3;
              vchf[mes][den]=v;
              nchf[mes][den]=n;


              //общие суммы по часам
              ovchf[mes]+=v;
              onchf[mes]+=n;
              ochf[mes]+=d3;


              //проверка праздничного дня
              if (PrazdDni(den,mes)==true)
                {
                  pchf[mes][den]=d3;
                  opchf[mes]+=d3;
                }

              //проверка дня в смене
              if (dnism==2)
                {
                  nsm=0;
                  dnism=5;
                }
              else
                {
                  dnism++;
                }
            }
          //выходной
          //************************
          else
            {
              if (dnism==1)
                {
                  if (mes==mes_mart2 && den==day_mart2)
                    {
                      chf[mes][den]=p2-1;
                      nchf[mes][den]=n2-1;
                      //проверка праздничного дня
                      if (PrazdDni(den,mes)==true)
                        {
                          pchf[mes][den]=p2-1;
                          opchf[mes]+=p2-1;
                        }
                      ochf[mes]+=p2-1;
                      onchf[mes]+=n2-1;
                    }
                  else if (mes==mes_oktyabr2 && den==day_oktyabr2)
                    {
                      chf[mes][den]=p2+1;
                      nchf[mes][den]=n2+1;
                      //проверка праздничного дня
                      if (PrazdDni(den,mes)==true)
                        {
                          pchf[mes][den]=p2+1;
                          opchf[mes]+=p2+1;
                        }
                      ochf[mes]+=p2+1;
                      onchf[mes]+=n2+1;
                    }
                  else
                    {
                      chf[mes][den]=p2;
                      nchf[mes][den]=n2;
                      //проверка праздничного дня
                      if (PrazdDni(den,mes)==true)
                        {
                          pchf[mes][den]=p2;
                          opchf[mes]+=p2;
                        }
                      ochf[mes]+=p2;
                      onchf[mes]+=n2;
                    }
                }
              else
                {
                  chf[mes][den]=0;
                  nchf[mes][den]="NULL";
                }

              vihod[mes][den]=0;

              //часы переходящие c предыдущего месяца
              if (den==1 && dnism==1 && mes==1)
                {
             //     ochf[mes]+= Vvod->SetN(DM->qObnovlenie2->FieldByName("chf0")->AsString);
             //     onchf[mes]+= Vvod->SetN(DM->qObnovlenie2->FieldByName("nch0")->AsString);
               //   opchf[mes]+= Vvod->SetN(DM->qObnovlenie2->FieldByName("pch0")->AsString);
                }
              else if (den==1 && dnism==1)
                {
                  chf0[mes-1]=p2;
                  //ochf[mes]+=p2;
                  nchf0[mes-1]=n2;
                 // onchf[mes]+=n2;

                  //проверка праздничного дня
                  if (PrazdDni(den,mes)==true)
                    {
                      pchf0[mes-1]=p2;
                   //   opchf[mes]+=p2;
                    }
                }

               //проверка дня в смене
               if (dnism==2)
                 {
                   nsm=2;
                   dnism=1;
                 }
               else if (dnism==4)
                 {
                   nsm=3;
                   dnism=1;
                 }
               else if (dnism==6)
                 {
                   nsm=1;
                   dnism=1;
                 }
               else
                 {
                   dnism++;
                 }
            }
        }

      // сохранение переходящих часов последнего дня в году
      if ((mes==12 && dnism==1 && nsm==0)||(mes==12 && nsm==1 && dnism==2))
        {
          chf0[mes]=p2;
          nchf0[mes]=n2;
          pchf0[mes]=p2;
        }

      //отсутствующие дни в месяце
      if (den<32)
        {
          while (den<=32)
            {
              vihod[mes][den]="NULL";
              chf[mes][den]="NULL";
              nchf[mes][den]="NULL";
              vchf[mes][den]="NULL";
              pchf[mes][den]="NULL";
              den++;
            }
        }

      DM->qNorma11Graf->Next();
    }
}
//------------------------------------------------------------------------------

// Расчет 133 графика
void __fastcall TMain::Graf133(double v, double n)
{
  int kol;

  /*nsm - номер смены последнего месяца,
   dnism - день смены последнего месяца*/

  nsm = DM->qObnovlenie2->FieldByName("nsm")->AsInteger;
  dnism = DM->qObnovlenie2->FieldByName("dnism")->AsInteger;

  //по месяцам
  for (mes=1; mes<=12; mes++)
    {
      kol = DaysInAMonth(god, mes);

      //по дням
      for (den=1; den<=kol; den++)
        {
          //рабочий день
          if (nsm==1)
            {
              vihod[mes][den]=1;
              chf[mes][den]=DM->qOgraf->FieldByName("DLIT")->AsFloat;
              vchf[mes][den]=v;
              nchf[mes][den]=n;

              ochf[mes]+=DM->qOgraf->FieldByName("DLIT")->AsFloat;
              ovchf[mes]+=v;
              onchf[mes]+=n;

                 // праздничный день
                 if (PrazdDni(den,mes)==true)
                  {
                    pchf[mes][den]=DM->qOgraf->FieldByName("DLIT")->AsFloat;
                    opchf[mes]+=DM->qOgraf->FieldByName("DLIT")->AsFloat;
                  }

              if (dnism==2)
                {
                  nsm=0;
                  dnism=1;
                }
              else
                {
                  dnism++;
                }
            }
          //выходной
          else
            {
              vihod[mes][den]=0;
              chf[mes][den]=0;

              if (dnism==2)
                {
                  nsm=1;
                  dnism=1;
                }
              else
                {
                  dnism++;
                }
            }
        }

      //отсутствующие дни в месяце
      if (den<32)
        {
          while (den<=32)
            {
              vihod[mes][den]="NULL";
              chf[mes][den]="NULL";
              den++;
            }
        }
      DM->qNorma11Graf->Next();
    }
}
//------------------------------------------------------------------------------

//Расчет 140 графика
void __fastcall TMain::Graf140(double d1, double v, double n)
{
  int kol, vihodnoy1=0, vihodnoy2=0;

 /* if (br==1||br==3)
    {
      vihodnoy1=1;
      vihodnoy2=2;
    }
  else if (br==2||br==4)
    {
      vihodnoy1=6;
      vihodnoy2=7;
    }  */

  /*nsm - номер смены последнего месяца,
  dnism - день смены последнего месяца*/

  nsm = DM->qObnovlenie2->FieldByName("nsm")->AsInteger;
  dnism = DM->qObnovlenie2->FieldByName("dnism")->AsInteger;

  //по месяцам
  for (mes=1; mes<=12; mes++)
    {
      kol = DaysInAMonth(god, mes);

      if ((br==1||br==3) && mes%2==0)
        {
          vihodnoy1=6;
          vihodnoy2=7;
        }
      else if ((br==1||br==3) && mes%2!=0)
        {
          vihodnoy1=1;
          vihodnoy2=2;
        }
      else if ((br==2||br==4) && mes%2==0)
        {
          vihodnoy1=1;
          vihodnoy2=2;
        }
      else if ((br==2||br==4) && mes%2!=0)
        {
          vihodnoy1=6;
          vihodnoy2=7;
        }

      //по дням
      for (den=1; den<=kol; den++)
        {
          //проверка дня недели (понедельник - пятница)
          if (DayWeek(den,mes,god)==vihodnoy1 || DayWeek(den,mes,god)==vihodnoy2)
            {
              if (DayWeek(den,mes,god)==vihodnoy2)
                {
                  /*if (dnism!=1)
                    {
                      if (nsm==2) nsm=3;
                      else if (nsm==3) nsm=2;
                      dnism=1;
                    } */
                    if (nsm==2) nsm=3;
                      else if (nsm==3) nsm=2;
                      dnism=1;

                }

             //проверка праздничного дня
              if (PrazdDni(den,mes)==true)
                {
                  vihod[mes][den]=9;
                  chf[mes][den]=30;
                }
              else
                {
                  vihod[mes][den]=0;
                  chf[mes][den]=NULL;
                }

              vchf[mes][den]=NULL;
              nchf[mes][den]=NULL;
              pchf[mes][den]=NULL;


           }
         //вторая смена (8.00-16.00)
         //*************************
         else if (nsm==2)
           {
             //проверка праздничного дня
             if (PrazdDni(den,mes)==true)
               {
                 vihod[mes][den]=9;
                 chf[mes][den]=30;
               }
             //проверка предпраздничного дня
             else if (PrdPrazdDni(den,mes)==true)
               {
                 vihod[mes][den]=2;
                 chf[mes][den]=d1-1;
                 ochf[mes]+=d1-1;
               }
             //рабочий день
             else
               {
                 vihod[mes][den]=2;
                 chf[mes][den]=d1;
                 ochf[mes]+=d1;
               }

         /*   if (dnism==5)
               {
                 nsm=3;
                 dnism=1;
               }
             else
               {   */
                 dnism++;
              // }
           }
         //третья смена (16.00-24.00)
         //**************************
         else if (nsm==3)
           {
             //проверка праздничного дня
             if (PrazdDni(den,mes)==true)
               {
                 vihod[mes][den]=9;
                 chf[mes][den]=30;
               }
             //проверка предпраздничного дня
             else if (PrdPrazdDni(den,mes)==true)
               {
                 vihod[mes][den]=3;
                 chf[mes][den]=d1-1;
                 vchf[mes][den]=v;

                 ochf[mes]+=d1-1;
                 ovchf[mes]+=v;
               }
             //рабочий день
             else
               {
                 vihod[mes][den]=3;
                 chf[mes][den]=d1;
                 vchf[mes][den]=v;
                 nchf[mes][den]=n;

                 ochf[mes]+=d1;
                 ovchf[mes]+=v;
                 onchf[mes]+=n;
               }

             /*if (dnism==5)
               {
                 nsm=2;
                 dnism=1;
               }
             else
               { */
                 dnism++;
              // }
            }
        }
      //отсутствующие дни в месяце
      if (den<32)
        {
          while (den<=32)
            {
              vihod[mes][den]="NULL";
              chf[mes][den]="NULL";
              den++;
            }
        }
    }
}
//---------------------------------------------------------------------------

// Расчет 150 графика
void __fastcall TMain::Graf150()
{
  int vih1, vih2, peredvih, kol;

  /*vih1, vih2 - выходные для бригады*/

  switch(br)
    { //1 - воскресенье, 2 - понедельник, 3 - вторник, 4 - среда, 5 - четверг, 6 - пятница, 7 - суббота
      // Для бригады 1
      case 1:
        vih1 = 2;
        vih2 = 3;
        peredvih = 1;
      break;

      //Для бригады 2
      case 2:
        vih1 = 4;
        vih2 = 5;
        peredvih = 3;
      break;

      //Для бригады 3
      case 3:
        vih1 = 6;
        vih2 = 7;
        peredvih = 5;
      break;

      //Для бригады 4
      case 4:
        vih1 = 1;
        vih2 = 2;
        peredvih = 7;
      break;
    }

  for (mes=1; mes<=12; mes++)
    {
      kol = DaysInAMonth(god, mes);

      //по дням
      for (den=1; den<=kol; den++)
        {
          //проверка дня недели
          if (DayWeek(den,mes,god)==vih1||DayWeek(den,mes,god)==vih2)
            {
              vihod[mes][den]=0;
              chf[mes][den]=0;
            }
          else
            {
              //проверка дня перед выходным
              if (DayWeek(den,mes,god)==peredvih)
                {
                  vihod[mes][den]=1;
                  chf[mes][den]=6.5;

                   //проверка праздничного дня
                   if (PrazdDni(den,mes)==true)
                     {
                       pchf[mes][den]=6.5;
                       opchf[mes]+=6.5;
                     }

                   //общие суммы по часам
                   ochf[mes]+=6.5;
                }
              //рабочий день
              else
                {
                  vihod[mes][den]=1;
                  chf[mes][den]=DM->qOgraf->FieldByName("DLIT")->AsFloat;

                  //проверка праздничного дня
                  if (PrazdDni(den,mes)==true)
                    {
                      pchf[mes][den]=DM->qOgraf->FieldByName("DLIT")->AsFloat;
                      opchf[mes]+=DM->qOgraf->FieldByName("DLIT")->AsFloat;
                    }

                  //общие суммы по часам
                  ochf[mes]+= DM->qOgraf->FieldByName("DLIT")->AsFloat;
                }
            }
        }

     //отсутствующие дни в месяце
      if (den<32)
        {
          while (den<=32)
            {
              vihod[mes][den]="NULL";
              chf[mes][den]="NULL";
              nchf[mes][den]="NULL";
              vchf[mes][den]="NULL";
              pchf[mes][den]="NULL";
              den++;
            }
        }

      DM->qNorma11Graf->Next();
    }
}

//------------------------------------------------------------------------------

//Расчет 160 графика
void __fastcall TMain::Graf160(double d1, double v, double n)
{
  int kol;

  //по месяцам
  for (mes=1; mes<=12; mes++)
    {
      kol = DaysInAMonth(god, mes);

      //по дням
      for (den=1; den<=kol; den++)
        {
          //проверка дня недели (понедельник - пятница)
          if (DayWeek(den,mes,god)!=7 && DayWeek(den,mes,god)!=1)
            {
              //проверка праздничного дня
              if (PrazdDni(den,mes)==true)
                {
                  vihod[mes][den]=9;
                  chf[mes][den]=30;
                }
              //праздничный попадающий на субботу или воскресенье
              else if (PrazdDniVihodnue(den,mes,god)==true)
                {
                  vihod[mes][den]=0;
                  chf[mes][den]=NULL;
                }
              else
                {
                  //если четверг или пятница
                  if (DayWeek(den,mes,god)==5 || DayWeek(den,mes,god)==6)
                    {
                      vihod[mes][den]=3;

                      vchf[mes][den]=v;
                      ovchf[mes]+=v;

                      //проверка предпраздничного дня
                      if (PrdPrazdDni(den,mes)==true)
                        {
                          chf[mes][den]=d1-1;
                          nchf[mes][den]=n-1;

                          //сумма часов
                          ochf[mes]+=d1-1;
                          onchf[mes]+=n-1;
                        }
                      else
                        {
                          chf[mes][den]=d1;
                          nchf[mes][den]=n;

                          //сумма часов
                          ochf[mes]+=d1;
                          onchf[mes]+=n;
                        }
                    }
                  else
                    {
                      vihod[mes][den]=2;

                      //проверка предпраздничного дня
                      if (PrdPrazdDni(den,mes)==true)
                        {
                          chf[mes][den]=d1-1;
                          //сумма часов
                          ochf[mes]+=d1-1;
                        }
                      else
                        {
                          chf[mes][den]=d1;
                          //сумма часов
                          ochf[mes]+=d1;
                        }
                    }
                }
            }
          //выходной
          else
            {
              //проверка праздничного дня
              if (PrazdDni(den,mes)==true)
                {
                  vihod[mes][den]=9;
                  chf[mes][den]=30;
                }
              else
                {
                  vihod[mes][den]=0;
                  chf[mes][den]=NULL;
                }

              vchf[mes][den]=NULL;
              nchf[mes][den]=NULL;
              pchf[mes][den]=NULL;
            }
        }
      //отсутствующие дни в месяце
      if (den<32)
        {
          while (den<=32)
            {
              vihod[mes][den]="NULL";
              chf[mes][den]="NULL";
              den++;
            }
        }
    }
}
//---------------------------------------------------------------------------

//Расчет 180 графика
void __fastcall TMain::Graf180(double d1, double p1, double p2, double v, double n1, double n2, double n)
{
   AnsiString kol;

  /*nsm - номер смены последнего месяца,
   dnism - день смены последнего месяца*/

  nsm = DM->qObnovlenie2->FieldByName("nsm")->AsInteger;
  dnism = DM->qObnovlenie2->FieldByName("dnism")->AsInteger;

  //по месяцам
  for (mes=1; mes<=12; mes++)
    {
      kol = DaysInAMonth(god, mes);

      //по дням
      for (den=1; den<=kol; den++)
        {
          //проверка дня недели (понедельник - пятница)
          //выходной
          if (DayWeek(den,mes,god)==7 || DayWeek(den,mes,god)==1)
            {
              //проверка праздничного дня
              if (PrazdDni(den,mes)==true)
                {
                  vihod[mes][den]=9;
                  chf[mes][den]=30;
                }
              else
                {
                  vihod[mes][den]=0;

                  //первая ночная смена
                  if (dnism==2)
                    {
                      chf[mes][den]=p1;
                      nchf[mes][den]=n1;

                      //общие часы
                      ochf[mes]+=p1;
                      onchf[mes]+=n1;
                    }
                  else
                    {
                      chf[mes][den]=NULL;
                      nchf[mes][den]=NULL;
                    }

                  vchf[mes][den]=NULL;
                  pchf[mes][den]=NULL;
                }

              //проверка дня в смене
              if (dnism==2)
                {
                  nsm=1;
                  dnism=1;
                }
              else if (dnism==4)
                {
                  nsm=2;
                  dnism=1;
                }
              else if (dnism==6)
                {
                  nsm=3;
                  dnism=1;
                }
              else
                {
                  dnism++;
                }
            }
          //первая смена (23.00-7.00)
          //*************************
          else if (nsm==1)
            {
              //проверка праздничного дня
              if (PrazdDni(den,mes)==true)
                {
                  vihod[mes][den]=9;

                  if (dnism!=5)
                    {
                      chf[mes][den]=p1;
                      pchf[mes][den]=p1;

                      //общие часы
                      ochf[mes]+=p1;
                      opchf[mes]+=p1;
                    }
                  else
                    {
                      chf[mes][den]=30;
                    }
                }
              //праздничный попадающий на субботу или воскресенье
              else if (PrazdDniVihodnue(den,mes,god)==true)
                {
                  vihod[mes][den]=0;

                  if (dnism!=5)
                    {
                      chf[mes][den]=p1;
                      nchf[mes][den]=n1;

                      //общие часы
                      ochf[mes]+=p1;
                      onchf[mes]+=n1;
                    }
                  else
                    {
                      chf[mes][den]=NULL;
                    }
                }
              else
                {
                  //если смена последняя или предпраздничная
                  if (dnism==5 || PrdPrazdDni(den,mes)==true)
                    {
                      vihod[mes][den]=1;
                      nchf[mes][den]=n2;
                      onchf[mes]+=n2;

                      //проверка предпраздничного дня
                      if (PrdPrazdDni(den,mes)==true)
                        {
                          chf[mes][den]=p2-1;
                          ochf[mes]+=p2-1;
                        }
                      else
                        {
                          chf[mes][den]=p2;
                          ochf[mes]+=p2;
                        }
                    }
                  //если смена не последняя
                  else
                    {
                      vihod[mes][den]=1;
                      chf[mes][den]=d1;
                      nchf[mes][den]=n1+n2;

                      //общие часы
                      onchf[mes]+=n1+n2;
                      ochf[mes]+=d1;
                    }
                }

              //проверка дня в смене
              if (dnism==5)
                {
                  nsm=0;
                  dnism=3;
                }
              else
                {
                  dnism++;
                }
            }
          //вторая смена (7.00-15.00)
          else if (nsm==2)
            {
              //проверка праздничного дня
              if (PrazdDni(den,mes)==true)
                {
                  vihod[mes][den]=9;
                  chf[mes][den]=30;
                }
              //праздничный попадающий на субботу или воскресенье
              else if (PrazdDniVihodnue(den,mes,god)==true)
                {
                  vihod[mes][den]=0;
                  chf[mes][den]=NULL;
                }
              //проверка предпраздничного дня
              else if (PrdPrazdDni(den,mes)==true)
                {
                  vihod[mes][den]=2;
                  chf[mes][den]=d1-1;
                  //общие суммы по часам
                  ochf[mes]+=d1-1;
                }
              else
                {
                  vihod[mes][den]=2;
                  chf[mes][den]=d1;
                  //общие суммы по часам
                  ochf[mes]+=d1;
                }

              //проверка дня в смене
              if (dnism==5)
                {
                  nsm=0;
                  dnism=5;
                }
              else
                {
                  dnism++;
                }
            }
          //третья смена (15.00-23.00)
          else if (nsm==3)
            {
              //проверка праздничного дня
              if (PrazdDni(den,mes)==true)
                {
                  vihod[mes][den]=9;
                  chf[mes][den]=30;
                }
              //праздничный попадающий на субботу или воскресенье
              else if (PrazdDniVihodnue(den,mes,god)==true)
                {
                  vihod[mes][den]=0;
                  chf[mes][den]=NULL;
                }  
              //проверка предпраздничного дня
              else if (PrdPrazdDni(den,mes)==true)
                {
                  vihod[mes][den]=3;
                  chf[mes][den]=d1-1;
                  vchf[mes][den]=v;
                  nchf[mes][den]=n-1;

                  //общие суммы по часам
                  ochf[mes]+=d1-1;
                  ovchf[mes]+=v;
                  onchf[mes]+=n-1;
                }
              else
                {

                  vihod[mes][den]=3;
                  chf[mes][den]=d1;
                  vchf[mes][den]=v;
                  nchf[mes][den]=n;

                  //общие суммы по часам
                  ochf[mes]+=d1;
                  ovchf[mes]+=v;
                  onchf[mes]+=n;
                }

              //проверка дня в смене
              if (dnism==5)
                {
                  nsm=0;
                  dnism=1;
                }
              else
                {
                  dnism++;
                }
            }
        }

      //расчет переработки
      if ((ochf[mes] - DM->qNorma11Graf->FieldByName("chf")->AsFloat - opchf[mes])>0)
        {
          pgraf[mes] = ochf[mes] - DM->qNorma11Graf->FieldByName("chf")->AsFloat - opchf[mes];
        }

      //отсутствующие дни в месяце
      if (den<32)
        {
          while (den<=32)
            {
              vihod[mes][den]="NULL";
              chf[mes][den]="NULL";
              nchf[mes][den]="NULL";
              vchf[mes][den]="NULL";
              pchf[mes][den]="NULL";
              den++;
            }
        }

      DM->qNorma11Graf->Next();
    }
}
//------------------------------------------------------------------------------

// Расчет 190 графика
void __fastcall TMain::Graf190(double d1, double v1, double v2)
{
  int kol;
  double ochf_obsh; //общая сумма по факту по всем месяцам
  /*nsm - номер смены последнего месяца,
   dnism - день смены последнего месяца*/

  nsm = DM->qObnovlenie2->FieldByName("nsm")->AsInteger;
  dnism = DM->qObnovlenie2->FieldByName("dnism")->AsInteger;

  //по месяцам
  for (mes=1; mes<=12; mes++)
    {
      kol = DaysInAMonth(god, mes);

      //по дням
   for (den=1; den<=kol; den++)
        {
          //рабочий день
          if (nsm==1||nsm==2)
            {
              vihod[mes][den]=1;
              chf[mes][den]= d1;
              ochf[mes]+=d1;

              //вечерние если бригада 1 или 2
              if (br==1||br==2)
                {
                  vchf[mes][den]=v1;
                  ovchf[mes]+=v1;
                }
              //вечерние если бригада 3 или 4
              else
                {
                  vchf[mes][den]=v2;
                  ovchf[mes]+=v2;
                }

              // праздничный день
              if (PrazdDni(den,mes)==true)
                {
                  pchf[mes][den]=d1;
                  opchf[mes]+=d1;
                }

              if (dnism==2)
                {
                  nsm=0;
                  dnism=1;
                }
              else
                {
                  dnism++;
                }
            }
          //выходной
          else
            {
              vihod[mes][den]=0;
              chf[mes][den]=0;
              vchf[mes][den]="NULL";

              if (dnism==2)
                {
                  nsm=1;
                  dnism=1;
                }
              else
                {
                  dnism++;
                }
            }
        }

      //отсутствующие дни в месяце
      if (den<32)
        {
          while (den<=32)
            {
              vihod[mes][den]="NULL";
              chf[mes][den]="NULL";
              den++;
            }
        }
      DM->qNorma11Graf->Next();
      ochf_obsh+=ochf[mes];
    }

   //10-ти часовые смены у 2 и 4 бригады из-за переработки
   mes=1;
   DM->qNorma11Graf->First();
   int i=1;

   //2 бригада
   if (ochf_obsh > DM->qNorma11Graf->FieldByName("onorma")->AsFloat)
     {
       int dni = (ochf_obsh - DM->qNorma11Graf->FieldByName("onorma")->AsFloat);

       if (br==2)
         {
           for (den=1; den<=31 && i<=dni; den++)
             {
               if (vihod[1][den]==1)
                 {
                   chf[1][den] = d1-1;
                   ochf[1] = ochf[1]-1;
                   i++;

                   //праздничные
                   if (PrazdDni(den,mes)==true)
                     {
                       pchf[1][den] = pchf[1][den]-1;
                       opchf[1] = opchf[1]-1;
                     }

                   //вечерние
                   if (v1<1)
                     {
                       vchf[1][den] = NULL;
                       ovchf[1] = ovchf[1]-v1;
                     }
                   else
                     {
                       vchf[1][den] = v1-1;
                       ovchf[1] = ovchf[1]-1;
                     }
                 }
             }
         }
       //4 бригада
       else if (br==4)
         {
           for (den=1; den<=31 && i<=dni; den++)
             {
               if (vihod[1][den]==1)
                 {
                   chf[1][den] = d1-1;
                   ochf[1] = ochf[1]-1;
                   i++;

                   //праздничные
                   if (PrazdDni(den,mes)==true)
                     {
                       pchf[1][den] = pchf[1][den]-1;
                       opchf[1] = opchf[1]-1;
                     }

                   //вечерние
                   vchf[1][den] = v2-1;
                   ovchf[1] = ovchf[1]-1;

                 }
             }
         }
     }
}
//------------------------------------------------------------------------------

// Расчет 220 графика
void __fastcall TMain::Graf220(double v, double n)
{
  int kol;

  /*nsm - номер смены последнего месяца,
   dnism - день смены последнего месяца*/

  nsm = DM->qObnovlenie2->FieldByName("nsm")->AsInteger;
  dnism = DM->qObnovlenie2->FieldByName("dnism")->AsInteger;

  //по месяцам
  for (mes=1; mes<=12; mes++)
    {
      kol = DaysInAMonth(god, mes);

      //по дням
      for (den=1; den<=kol; den++)
        {
          //1 смена
          if (nsm==1)
            {
              vihod[mes][den]=1;
              chf[mes][den]= DM->qOgraf->FieldByName("DLIT")->AsFloat;
              ochf[mes]+=DM->qOgraf->FieldByName("DLIT")->AsFloat;

              // праздничный день
              if (PrazdDni(den,mes)==true)
                {
                  pchf[mes][den]=DM->qOgraf->FieldByName("DLIT")->AsFloat;
                  opchf[mes]+=DM->qOgraf->FieldByName("DLIT")->AsFloat;
                }

              nsm=2;
            }
          //2 смена
          else if (nsm==2)
            {
              vihod[mes][den]=2;
              chf[mes][den]= DM->qOgraf->FieldByName("DLIT")->AsFloat;
              vchf[mes][den]=v;
              nchf[mes][den]=n;

              //общие суммы по часам
              ochf[mes]+=DM->qOgraf->FieldByName("DLIT")->AsFloat;
              ovchf[mes]+=v;
              onchf[mes]+=n;

              // праздничный день
              if (PrazdDni(den,mes)==true)
                {
                  pchf[mes][den]=DM->qOgraf->FieldByName("DLIT")->AsFloat;
                  opchf[mes]+=DM->qOgraf->FieldByName("DLIT")->AsFloat;
                }

              nsm=0;
            }
          //выходной
          else
            {
              vihod[mes][den]=0;
              chf[mes][den]=0;

              nsm=1;
            }
        }

      //отсутствующие дни в месяце
      if (den<32)
        {
          while (den<=32)
            {
              vihod[mes][den]="NULL";
              chf[mes][den]="NULL";
              den++;
            }
        }
      DM->qNorma11Graf->Next();
    }
}
//------------------------------------------------------------------------------

// Расчет 225 графика
void __fastcall TMain::Graf225(double v1, double v2, double n)
{
  int kol;

  /*nsm - номер смены последнего месяца,
   dnism - день смены последнего месяца*/

  nsm = DM->qObnovlenie2->FieldByName("nsm")->AsInteger;
  dnism = DM->qObnovlenie2->FieldByName("dnism")->AsInteger;

  //по месяцам
  for (mes=1; mes<=12; mes++)
    {
      kol = DaysInAMonth(god, mes);

      //по дням
      for (den=1; den<=kol; den++)
        {
           //********************************************************************
          //выходной
          //проверка дня недели
          if (DayWeek(den,mes,god)==7||DayWeek(den,mes,god)==1)
            {
              //проверка праздничного дня
              if (PrazdDni(den,mes)==true)
                {
                  vihod[mes][den]=9;
                  chf[mes][den]=30;
                }
              else
                {
                  vihod[mes][den]=0;
                  chf[mes][den]=0;
                }
            }
          //рабочий день
          else
            {
              //********************************************************************
              //1 смена (8.00-16.30)
              if (nsm==1)
                {
                  //проверка праздничного дня
                  if (PrazdDni(den,mes)==true)
                    {
                      vihod[mes][den]=9;
                      chf[mes][den]=30;
                    }
                  //праздничный попадающий на субботу или воскресенье
                  else if (PrazdDniVihodnue(den,mes,god)==true)
                    {
                      vihod[mes][den]=0;
                      chf[mes][den]=0;
                    }
                  //рабочий день
                  else
                    {
                      vihod[mes][den]=1;
                      chf[mes][den]=DM->qOgraf->FieldByName("DLIT")->AsFloat;
                      ochf[mes]+= DM->qOgraf->FieldByName("DLIT")->AsFloat;
                    }

                  if (dnism==5)
                    {
                      nsm=2;
                      dnism=1;
                    }
                  else
                    {
                      dnism++;
                    }
                 }
              //********************************************************************
              //2 смена (13.00-21.30)
              else if (nsm==2)
                {
                  //проверка праздничного дня
                  if (PrazdDni(den,mes)==true)
                    {
                      vihod[mes][den]=9;
                      chf[mes][den]=30;
                    }
                  //праздничный попадающий на субботу или воскресенье
                  else if (PrazdDniVihodnue(den,mes,god)==true)
                    {
                      vihod[mes][den]=0;
                      chf[mes][den]=0;
                    }
                  //рабочий день
                  else
                    {
                      vihod[mes][den]=2;
                      chf[mes][den]=DM->qOgraf->FieldByName("DLIT")->AsFloat;
                      vchf[mes][den]=v1;

                      ochf[mes]+=DM->qOgraf->FieldByName("DLIT")->AsFloat;
                      ovchf[mes]+=v1;
                    }

                  nsm=3;
                }
              //********************************************************************
              //3 смена (9.00-17.00)
              else if (nsm==3)
                {
                  //проверка праздничного дня
                  if (PrazdDni(den,mes)==true)
                    {
                      vihod[mes][den]=9;
                      chf[mes][den]=30;
                    }
                  //праздничный попадающий на субботу или воскресенье
                  else if (PrazdDniVihodnue(den,mes,god)==true)
                    {
                      vihod[mes][den]=0;
                      chf[mes][den]=0;
                    }
                  //рабочий день
                  else
                    {
                      vihod[mes][den]=3;
                      chf[mes][den]=DM->qOgraf->FieldByName("DLIT")->AsFloat;
                      ochf[mes]+= DM->qOgraf->FieldByName("DLIT")->AsFloat;
                    }

                  if (dnism==1)
                    {
                      nsm=4;
                      dnism=1;
                    }
                  else
                    {
                      nsm=4;
                      dnism=2;
                    }
                }
              //********************************************************************
              //4 смена (14.00-22.30)
              else if (nsm==4)
                {
                  //проверка праздничного дня
                  if (PrazdDni(den,mes)==true)
                    {
                      vihod[mes][den]=9;
                      chf[mes][den]=30;
                    }
                  //праздничный попадающий на субботу или воскресенье
                  else if (PrazdDniVihodnue(den,mes,god)==true)
                    {
                      vihod[mes][den]=0;
                      chf[mes][den]=0;
                    }
                /*  //проверка предпраздничного дня
                  else if (PrdPrazdDni(den,mes)==true)
                    {
                      vihod[mes][den]=4;
                      chf[mes][den]=DM->qOgraf->FieldByName("DLIT")->AsFloat-1;
                      vchf[mes][den]=v2-1;

                      ochf[mes]+= DM->qOgraf->FieldByName("DLIT")->AsFloat-1;
                      ovchf[mes]+=v2-1;
                    }     */
                  //рабочий день
                  else
                    {
                      vihod[mes][den]=4;
                      chf[mes][den]=DM->qOgraf->FieldByName("DLIT")->AsFloat;
                      vchf[mes][den]=v2;
                      nchf[mes][den]=n;

                      ochf[mes]+= DM->qOgraf->FieldByName("DLIT")->AsFloat;
                      ovchf[mes]+=v2;
                      onchf[mes]+=n;
                    }

                  if (dnism==1)
                    {
                      nsm=3;
                      dnism=2;
                    }
                  else
                    {
                      nsm=1;
                      dnism=1;
                    }
                }
            }
        }
      //расчет переработки
      if ((ochf[mes] - DM->qNorma11Graf->FieldByName("chf")->AsFloat)>0)
        {
          pgraf[mes] = ochf[mes] - DM->qNorma11Graf->FieldByName("chf")->AsFloat;
        }

      //отсутствующие дни в месяце
      if (den<32)
        {
          while (den<=32)
            {
              vihod[mes][den]="NULL";
              chf[mes][den]="NULL";
              den++;
            }
        }
      DM->qNorma11Graf->Next();  
    }
}
//------------------------------------------------------------------------------

// Расчет 230 графика
void __fastcall TMain::Graf230(double v)
{
  int kol;

  /* chf[32] - рабочие часы по дням
     chf[den] = 8 - рабочий день
     chf[den] = 7 - предпраздничный день
     vihod[32] - выходы по дням (рабочий, отдых, праздничный)
     vihod[den] = 1 - рабочий день
     vihod[den] = 9 - праздник
  */

  //по месяцам
  for (mes=1; mes<=12; mes++)
    {
      kol = DaysInAMonth(god, mes);

      //Для периода с мая по сентябрь
      //*****************************
      if (mes==5 || mes==6 || mes==7 || mes==8 || mes==9)
        {
          //по дням
          for (den=1; den<=kol; den++)
            {
              //проверка дня недели
              if (DayWeek(den,mes,god)==3||DayWeek(den,mes,god)==4)
                {
                  //проверка праздничного дня
                  if (PrazdDni(den,mes)==true)
                    {
                      vihod[mes][den]=9;
                      chf[mes][den]=30;
                    }
                  else
                    {
                      vihod[mes][den]=0;
                      chf[mes][den]=0;
                    }
                }
              else
                {
                  //проверка праздничного дня
                  if (PrazdDni(den,mes)==true)
                    {
                      vihod[mes][den]=9;
                      chf[mes][den]=30;
                    }
                  //рабочий день
                  else
                    {
                      vihod[mes][den]=1;
                      chf[mes][den]=DM->qOgraf->FieldByName("DLIT")->AsFloat;
                      vchf[mes][den]=v;

                      //общие суммы по часам
                      ochf[mes]+= DM->qOgraf->FieldByName("DLIT")->AsFloat;
                      ovchf[mes]+=v;
                    }
                }
            }

        }
      //Для периода с января по апрель и с октября по декабрь
      else
        {
          //по дням
          for (den=1; den<=kol; den++)
            {
              //проверка дня недели
              if (DayWeek(den,mes,god)==1||DayWeek(den,mes,god)==7)
                {
                  //проверка праздничного дня
                  if (PrazdDni(den,mes)==true)
                    {
                      vihod[mes][den]=9;
                      chf[mes][den]=30;
                    }
                  else
                    {
                      vihod[mes][den]=0;
                      chf[mes][den]=0;
                    }
                }
              else
                {
                  //проверка праздничного дня
                  if (PrazdDni(den,mes)==true)
                    {
                      vihod[mes][den]=9;
                      chf[mes][den]=30;
                    }
                  //праздничный попадающий на субботу или воскресенье
                  else if (PrazdDniVihodnue(den,mes,god)==true)
                    {
                      vihod[mes][den]=9;
                      chf[mes][den]=30;
                    }
                  //проверка предпраздничного дня
                  else if (PrdPrazdDni(den,mes)==true)
                    {
                      vihod[mes][den]=1;
                      chf[mes][den]=DM->qOgraf->FieldByName("DLIT")->AsFloat-1;
                      ochf[mes]+= DM->qOgraf->FieldByName("DLIT")->AsFloat-1;
                    }
                  //рабочий день
                  else
                    {
                      vihod[mes][den]=1;
                      chf[mes][den]=DM->qOgraf->FieldByName("DLIT")->AsFloat;

                      //общие суммы по часам
                      ochf[mes]+= DM->qOgraf->FieldByName("DLIT")->AsFloat;
                    }
                }
            }
        }

      //отсутствующие дни в месяце
      if (den<32)
        {
          while (den<=32)
            {
              vihod[mes][den]="NULL";
              chf[mes][den]="NULL";
              den++;
            }
        }
    }

       


}

//------------------------------------------------------------------------------

//Расчет 240 графика
void __fastcall TMain::Graf240(double v, double n1, double n2)
{
  AnsiString kol;

  /*nsm - номер смены последнего месяца,
   dnism - день смены последнего месяца*/

  nsm = DM->qObnovlenie2->FieldByName("nsm")->AsInteger;
  dnism = DM->qObnovlenie2->FieldByName("dnism")->AsInteger;

  //по месяцам
  for (mes=1; mes<=12; mes++)
    {
      kol = DaysInAMonth(god, mes);

      //по дням
      for (den=1; den<=kol; den++)
        {
          //первая смена (4.00-12.42)
          //*************************
          if (nsm==1)
            {
              vihod[mes][den]=1;
              nchf[mes][den]=n1;
              chf[mes][den]=DM->qOgraf->FieldByName("DLIT")->AsFloat;

              //общие суммы по часам
              onchf[mes]+=n1;
              ochf[mes]+=DM->qOgraf->FieldByName("DLIT")->AsFloat;

              //проверка праздничного дня
              if (PrazdDni(den,mes)==true)
                {
                  pchf[mes][den]=DM->qOgraf->FieldByName("DLIT")->AsFloat;
                  opchf[mes]+=DM->qOgraf->FieldByName("DLIT")->AsFloat;
                }

              if (dnism==2)
                {
                  dnism=1;
                  nsm=2;
                }
              else
                {
                  dnism++;
                }
            }
          //вторая смена (15.00-23.42)
          //*************************
          else if (nsm==2)
            {
              vihod[mes][den]=2;
              vchf[mes][den]=v;
              nchf[mes][den]=n2;
              chf[mes][den]=DM->qOgraf->FieldByName("DLIT")->AsFloat;

              //общие суммы по часам
              onchf[mes]+=n2;
              ovchf[mes]+=v;
              ochf[mes]+=DM->qOgraf->FieldByName("DLIT")->AsFloat;

              //проверка праздничного дня
              if (PrazdDni(den,mes)==true)
                {
                  pchf[mes][den]=DM->qOgraf->FieldByName("DLIT")->AsFloat;
                  opchf[mes]+=DM->qOgraf->FieldByName("DLIT")->AsFloat;
                }

              if (dnism==2)
                {
                  dnism=1;
                  nsm=0;
                }
              else
                {
                  dnism++;
                }
            }

          //выходной
          //************************
          else
            {
              chf[mes][den]=0;
              vihod[mes][den]=0;

              //проверка дня в смене
              if (dnism==2)
                {
                  nsm=1;
                  dnism=1;
                }
              else
                {
                  dnism++;
                }
            }
        }

      //расчет переработки
      if ((ochf[mes] - DM->qNorma11Graf->FieldByName("chf")->AsFloat - opchf[mes])>0)
        {
          pgraf[mes] = ochf[mes] - DM->qNorma11Graf->FieldByName("chf")->AsFloat - opchf[mes];
        }

      //отсутствующие дни в месяце
      if (den<32)
        {
          while (den<=32)
            {
              vihod[mes][den]="NULL";
              chf[mes][den]="NULL";
              nchf[mes][den]="NULL";
              vchf[mes][den]="NULL";
              pchf[mes][den]="NULL";
              den++;
            }
        }

      DM->qNorma11Graf->Next();
    }
}
//------------------------------------------------------------------------------
//Расчет 250 графика
void __fastcall TMain::Graf250(double d1, double v1, double v2, double n)
{
  int kol, vihodnoy1, vihodnoy2;

  if (br==1)
    {
      vihodnoy1=6;
    }
  else if (br==2)
    {
      vihodnoy1=7;
    }

  //по месяцам
  for (mes=1; mes<=12; mes++)
    {
      kol = DaysInAMonth(god, mes);

      //по дням
      for (den=1; den<=kol; den++)
        {
          //выходной
          if (DayWeek(den,mes,god)==1||DayWeek(den,mes,god)==2)
            {
              vihod[mes][den]=0;
              chf[mes][den]=NULL;
              vchf[mes][den]=NULL;
              nchf[mes][den]=NULL;
              pchf[mes][den]=NULL;
            }
          //рабочий день
          else
            {
              //проверка праздничного дня и предпраздничного
              if (PrazdDni(den,mes)==true && PrdPrazdDni(den,mes)==true)
                {
                  vihod[mes][den]=9;
                  chf[mes][den]=d1-1;
                  vchf[mes][den]=v2;
                  pchf[mes][den]=d1-1;

                  //общие часы
                  ochf[mes]+=d1-1;
                  ovchf[mes]+=v2;
                  opchf[mes]+=d1-1;
                }
              //проверка праздничного дня 
              else if (PrazdDni(den,mes)==true)
                {
                  vihod[mes][den]=9;
                  chf[mes][den]=d1;
                  vchf[mes][den]=v2;
                  nchf[mes][den]=n;
                  pchf[mes][den]=d1;

                  //общие часы
                  ochf[mes]+=d1;
                  ovchf[mes]+=v2;
                  onchf[mes]+=n;
                  opchf[mes]+=d1;
                }
              //проверка 1 бригада - пятницы, 2 бригада - суббота
              else if (DayWeek(den,mes,god)==vihodnoy1)
                {
                  vihod[mes][den]=1;
                  vchf[mes][den]=v2;

                  //общие часы
                  ovchf[mes]+=v2;

                  //проверка предпраздничного дня
                  if (PrdPrazdDni(den,mes)==true)
                    {
                      chf[mes][den]=d1-1;
                      ochf[mes]+=d1-1;
                    }
                  else
                    {
                      chf[mes][den]=d1;
                      nchf[mes][den]=n;

                      ochf[mes]+=d1;
                      onchf[mes]+=n;
                    }
                }
              //проверка предпраздничного дня
              else if (PrdPrazdDni(den,mes)==true)
                {
                  vihod[mes][den]=1;
                  chf[mes][den]=d1-1;
                  if (br==2)
                    {
                      vchf[mes][den]=v1-1;
                      ovchf[mes]+=v1-1;
                    }

                  //общие часы
                  ochf[mes]+=d1-1;
                }
              //рабочий день
              else
                {
                  vihod[mes][den]=1;
                  chf[mes][den]=d1;
                  if (br==2)
                    {
                      vchf[mes][den]=v1;
                      ovchf[mes]+=v1;
                    }

                  //общие часы
                  ochf[mes]+=d1;
                }
            }
        }

      //отсутствующие дни в месяце
      if (den<32)
        {
          while (den<=32)
            {
              vihod[mes][den]="NULL";
              chf[mes][den]="NULL";
              nchf[mes][den]="NULL";
              vchf[mes][den]="NULL";
              pchf[mes][den]="NULL";
              den++;
            }
        }
      DM->qNorma11Graf->Next();
    }
}
//---------------------------------------------------------------------------

//Расчет 260 графика
void __fastcall TMain::Graf260(double v, double n)
{
  int vih1, vih2, peredvih, kol;

  /*vih1, vih2 - выходные для бригады*/

  switch(br)
    { //1 - воскресенье, 2 - понедельник, 3 - вторник, 4 - среда, 5 - четверг, 6 - пятница, 7 - суббота
      // Для бригады 1
      case 1:
        vih1 = 1;
        vih2 = 2;
        peredvih = 7;
      break;

      //Для бригады 2
      case 2:
        vih1 = 4;
        vih2 = 5;
        peredvih = 3;
      break;

      //Для бригады 3
      case 3:
        vih1 = 6;
        vih2 = 7;
        peredvih = 5;
      break;
    }

  /*nsm - номер смены последнего месяца,
   dnism - день смены последнего месяца*/

   nsm = DM->qObnovlenie2->FieldByName("nsm")->AsInteger;
   dnism = DM->qObnovlenie2->FieldByName("dnism")->AsInteger;

  //по месяцам
  for (mes=1; mes<=12; mes++)
    {
      kol = DaysInAMonth(god, mes);

      //по дням
      for (den=1; den<=kol; den++)
        {
          //проверка дня недели (выходной)
          if (DayWeek(den,mes,god)==vih1||DayWeek(den,mes,god)==vih2)
            {
              vihod[mes][den]=0;
              chf[mes][den]=0;
            }
          //рабочий день
          else
            {
              //первая смена (6.30-14.30)
              if (nsm==1)
                {
                  vihod[mes][den]=1;

                  //проверка дня перед выходным
                  if (DayWeek(den,mes,god)==peredvih)
                    {
                      chf[mes][den]=DM->qOgraf->FieldByName("DLIT")->AsFloat-1.5;
                      ochf[mes]+=DM->qOgraf->FieldByName("DLIT")->AsFloat-1.5;

                      //проверка праздничного дня
                      if (PrazdDni(den,mes)==true)
                        {
                          pchf[mes][den]=DM->qOgraf->FieldByName("DLIT")->AsFloat-1.5;
                          opchf[mes]+=DM->qOgraf->FieldByName("DLIT")->AsFloat-1.5;
                        }
                    }
                  //рабочий день
                  else
                    {
                      chf[mes][den]=DM->qOgraf->FieldByName("DLIT")->AsFloat;
                      ochf[mes]+= DM->qOgraf->FieldByName("DLIT")->AsFloat;

                      //проверка праздничного дня
                      if (PrazdDni(den,mes)==true)
                        {
                          pchf[mes][den]=DM->qOgraf->FieldByName("DLIT")->AsFloat;
                          opchf[mes]+=DM->qOgraf->FieldByName("DLIT")->AsFloat;
                        }
                    }

                  if (dnism==5)
                    {
                      nsm=2;
                      dnism=1;
                    }
                  else
                    {
                      dnism++;
                    }
                }
              //вторая смена (14.30-22.30)
              else if (nsm==2)
                {
                  vihod[mes][den]=2;

                  //проверка дня перед выходным
                  if (DayWeek(den,mes,god)==peredvih)
                    {
                      chf[mes][den]=DM->qOgraf->FieldByName("DLIT")->AsFloat-1.5;
                      vchf[mes][den]=v-0.5;

                      ochf[mes]+=DM->qOgraf->FieldByName("DLIT")->AsFloat-1.5;
                      ovchf[mes]+=v-0.5;

                      //проверка праздничного дня
                      if (PrazdDni(den,mes)==true)
                        {
                          pchf[mes][den]=DM->qOgraf->FieldByName("DLIT")->AsFloat-1.5;
                          opchf[mes]+=DM->qOgraf->FieldByName("DLIT")->AsFloat-1.5;
                        }
                    }

                  //рабочий день
                  else
                    {
                      chf[mes][den]=DM->qOgraf->FieldByName("DLIT")->AsFloat;
                      vchf[mes][den]=v;
                      nchf[mes][den]=n;

                      ochf[mes]+=DM->qOgraf->FieldByName("DLIT")->AsFloat;
                      ovchf[mes]+=v;
                      onchf[mes]+=n;

                      //проверка праздничного дня
                      if (PrazdDni(den,mes)==true)
                        {
                          pchf[mes][den]=DM->qOgraf->FieldByName("DLIT")->AsFloat;
                          opchf[mes]+=DM->qOgraf->FieldByName("DLIT")->AsFloat;
                        }
                    }

                  if (dnism==5)
                    {
                      nsm=1;
                      dnism=1;
                    }
                  else
                    {
                      dnism++;
                    }
                }
            }
        }
      //отсутствующие дни в месяце
      if (den<32)
        {
          while (den<=32)
            {
              vihod[mes][den]="NULL";
              chf[mes][den]="NULL";
              nchf[mes][den]="NULL";
              vchf[mes][den]="NULL";
              pchf[mes][den]="NULL";
              den++;
            }
        }
      DM->qNorma11Graf->Next();
    }
}
//------------------------------------------------------------------------------
//Расчет 270 графика
void __fastcall TMain::Graf270(double d1, double v, double n)
{
  AnsiString kol;

  /*nsm - номер смены последнего месяца,
   dnism - день смены последнего месяца*/

   nsm = DM->qObnovlenie2->FieldByName("nsm")->AsInteger;
   dnism = DM->qObnovlenie2->FieldByName("dnism")->AsInteger;

  //по месяцам
  for (mes=1; mes<=12; mes++)
    {
      kol = DaysInAMonth(god, mes);

      //по дням
      for (den=1; den<=kol; den++)
        {
          //вторая смена (6.30-15.15)
          if (nsm==2)
            {
              vihod[mes][den]=2;
              chf[mes][den]=d1;
              ochf[mes]+=d1;

              //проверка праздничного дня
              if (PrazdDni(den,mes)==true)
                {
                  pchf[mes][den]=d1;
                  opchf[mes]+=d1;
                }

              if (dnism==4)
                {
                  nsm=0;
                  dnism=3;
                }
              else
                {
                  dnism++;
                }
            }
          //третья смена (15.15-00.00)
          else if (nsm==3)
            {
              vihod[mes][den]=3;
              chf[mes][den]=d1;
              vchf[mes][den]=v;
              nchf[mes][den]=n;


              ochf[mes]+=d1;
              ovchf[mes]+=v;
              onchf[mes]+=n;

              //проверка праздничного дня
              if (PrazdDni(den,mes)==true)
                {
                  pchf[mes][den]=d1;
                  opchf[mes]+=d1;
                }

              if (dnism==4)
                {
                  nsm=0;
                  dnism=1;
                }
              else
                {
                  dnism++;
                }
            }
          //выходной
          else if (nsm==0)
            {
              vihod[mes][den]=0;
              chf[mes][den]=NULL;
              vchf[mes][den]=NULL;
              nchf[mes][den]=NULL;

              if (dnism==2)
                {
                  nsm=2;
                  dnism=1;
                }
              else if (dnism==4)
                {
                  nsm=3;
                  dnism=1;
                }
              else
                {
                  dnism++;
                }
            }
        }
      //отсутствующие дни в месяце
      if (den<32)
        {
          while (den<=32)
            {
              vihod[mes][den]="NULL";
              chf[mes][den]="NULL";
              nchf[mes][den]="NULL";
              vchf[mes][den]="NULL";
              pchf[mes][den]="NULL";
              den++;
            }
        }

      DM->qNorma11Graf->Next();
    }
}
//------------------------------------------------------------------------------
//Расчет 280 графика
void __fastcall TMain::Graf280(double d1, double v)
{
  AnsiString kol;

  /*nsm - номер смены последнего месяца,
   dnism - день смены последнего месяца*/

   nsm = DM->qObnovlenie2->FieldByName("nsm")->AsInteger;
   dnism = DM->qObnovlenie2->FieldByName("dnism")->AsInteger;

  //по месяцам
  for (mes=1; mes<=12; mes++)
    {
      kol = DaysInAMonth(god, mes);

      //c 1 января по 30 апреля и с 1 ноября по 31 декабря работы выполняются по 18 графику
      //***********************************************************************************
      if (mes==1 || mes==2 || mes==3 || mes==4 || mes==11 || mes==12)
        {
          kol = DaysInAMonth(god, mes);

          //по дням
          for (den=1; den<=kol; den++)
            {
              //проверка дня недели
              if (DayWeek(den,mes,god)==1||DayWeek(den,mes,god)==7)
                {
                  //проверка праздничного дня
                  if (PrazdDni(den,mes)==true)
                    {
                      vihod[mes][den]=9;
                      chf[mes][den]=30;
                    }
                  else
                    {
                      vihod[mes][den]=0;
                      chf[mes][den]=0;
                    }
                }
              else
                {
                  //проверка праздничного дня
                  if (PrazdDni(den,mes)==true)
                    {
                      vihod[mes][den]=9;
                      chf[mes][den]=30;
                    }
                  //праздничный попадающий на субботу или воскресенье
                  else if (PrazdDniVihodnue(den,mes,god)==true)
                    {
                      vihod[mes][den]=0;
                      chf[mes][den]=0;
                    }
                  //проверка предпраздничного дня
                  else if (PrdPrazdDni(den,mes)==true)
                    {
                      vihod[mes][den]=1;
                      chf[mes][den]=DM->qOgraf->FieldByName("DLIT")->AsFloat-1;
                      ochf[mes]+= DM->qOgraf->FieldByName("DLIT")->AsFloat-1;
                    }
                  //рабочий день
                  else
                    {
                      vihod[mes][den]=1;
                      chf[mes][den]=DM->qOgraf->FieldByName("DLIT")->AsFloat;
                      ochf[mes]+= DM->qOgraf->FieldByName("DLIT")->AsFloat;
                    }
                }
            }
          //отсутствующие дни в месяце
          if (den<32)
            {
              while (den<=32)
                {
                  vihod[mes][den]="NULL";
                  chf[mes][den]="NULL";
                  den++;
                }
            }

          DM->qNorma11Graf->Next();
         // mes++;

        }
      //с 1 мая по 31 октября
      //**********************************************************************
      else
        {
          kol = DaysInAMonth(god, mes);

          //по дням
          for (den=1; den<=kol; den++)
            {
              //проверка дня недели
              if (DayWeek(den,mes,god)==1||DayWeek(den,mes,god)==7)
                {
                  //проверка праздничного дня
                  if (PrazdDni(den,mes)==true)
                    {
                      vihod[mes][den]=9;
                      chf[mes][den]=30;
                    }
                  else
                    {
                      vihod[mes][den]=0;
                      chf[mes][den]=NULL;
                    }
                }
              //2-я смена (7.00-15.30)
              else if (nsm==2)
                {
                  if (DayWeek(den,mes,god)==2)
                    {
                      dnism=1;
                    }

                  //проверка праздничного дня
                  if (PrazdDni(den,mes)==true)
                    {
                      vihod[mes][den]=9;
                      chf[mes][den]=30;
                    }
                  //праздничный попадающий на субботу или воскресенье
                  else if (PrazdDniVihodnue(den,mes,god)==true)
                    {
                      vihod[mes][den]=0;
                      chf[mes][den]=0;
                    }
                  //проверка предпраздничного дня
                  else if (PrdPrazdDni(den,mes)==true)
                    {
                      vihod[mes][den]=2;
                      chf[mes][den]=d1-1;
                      ochf[mes]+=d1-1;
                    }
                  //рабочий день
                  else
                    {
                      vihod[mes][den]=2;
                      chf[mes][den]=d1;
                      ochf[mes]+=d1;
                    }

                  //проверка дня смены
                  if (dnism==5)
                    {
                      nsm=3;
                      dnism=1;
                    }
                  else
                    {
                      dnism++;
                    }
                }
              //3-я смена (12.30-21.00)
              else if (nsm==3)
                {
                  if (DayWeek(den,mes,god)==2)
                    {
                      dnism=1;
                    }
                  //проверка праздничного дня
                  if (PrazdDni(den,mes)==true)
                    {
                      vihod[mes][den]=9;
                      chf[mes][den]=30;
                    }
                  //праздничный попадающий на субботу или воскресенье
                  else if (PrazdDniVihodnue(den,mes,god)==true)
                    {
                      vihod[mes][den]=0;
                      chf[mes][den]=0;
                    }
                  //проверка предпраздничного дня
                  else if (PrdPrazdDni(den,mes)==true)
                    {
                      vihod[mes][den]=3;
                      chf[mes][den]=d1-1;
                      vchf[mes][den]=v-1;

                      ochf[mes]+=d1-1;
                      ovchf[mes]+=v-1;
                    }
                  //рабочий день
                  else
                    {
                      vihod[mes][den]=3;
                      chf[mes][den]=d1;
                      vchf[mes][den]=v;

                      ochf[mes]+=d1;
                      ovchf[mes]+=v;
                    }

                  //проверка дня смены
                  if (dnism==5)
                    {
                      nsm=2;
                      dnism=1;
                    }
                  else
                    {
                      dnism++;
                    }
                }
            }
          //отсутствующие дни в месяце
          if (den<32)
            {
              while (den<=32)
                {
                  vihod[mes][den]="NULL";
                  chf[mes][den]="NULL";
                  den++;
                }
            }

          DM->qNorma11Graf->Next();
          //mes++;
       }


    }
  if (br==1)
        {
          nsm=2;
          dnism=1;
        }
      else
        {
          nsm=3;
          dnism=1;
        }  
}
//------------------------------------------------------------------------------

// Расчет 300 и 131 графика
void __fastcall TMain::Graf300(double v)
{
  int kol;

  /*nsm - номер смены последнего месяца,
   dnism - день смены последнего месяца*/

  nsm = DM->qObnovlenie2->FieldByName("nsm")->AsInteger;
  dnism = DM->qObnovlenie2->FieldByName("dnism")->AsInteger;

  //по месяцам
  for (mes=1; mes<=12; mes++)
    {
      kol = DaysInAMonth(god, mes);

      //по дням
   for (den=1; den<=kol; den++)
        {
          //рабочий день
          if (nsm==1||nsm==2)
            {
              vihod[mes][den]=1;
              chf[mes][den]= DM->qOgraf->FieldByName("DLIT")->AsFloat;
              vchf[mes][den]=v;       //vchf[mes][den]=v;

              ochf[mes]+=DM->qOgraf->FieldByName("DLIT")->AsFloat;
              ovchf[mes]+=v;   //ovchf[mes]=v;

                 // праздничный день
                 if (PrazdDni(den,mes)==true)
                  {
                    pchf[mes][den]=DM->qOgraf->FieldByName("DLIT")->AsFloat;
                    opchf[mes]+=DM->qOgraf->FieldByName("DLIT")->AsFloat;
                  }

              if (dnism==2)
                {
                  nsm=0;
                  dnism=1;
                }
              else
                {
                  dnism++;
                }
            }
          //выходной
          else
            {
              vihod[mes][den]=0;
              chf[mes][den]=0;
              vchf[mes][den]="NULL";

              if (dnism==2)
                {
                  nsm=1;
                  dnism=1;
                }
              else
                {
                  dnism++;
                }
            }
        }
     if (graf!=131)
       {
       //расчет переработки
      if ((ochf[mes] - DM->qNorma11Graf->FieldByName("chf")->AsFloat - opchf[mes])>0)
        {
          pgraf[mes] = ochf[mes] - DM->qNorma11Graf->FieldByName("chf")->AsFloat - opchf[mes];
        }
      }

      //отсутствующие дни в месяце
      if (den<32)
        {
          while (den<=32)
            {
              vihod[mes][den]="NULL";
              chf[mes][den]="NULL";
              den++;
            }
        }
      DM->qNorma11Graf->Next();
    }
}
//------------------------------------------------------------------------------

// Расчет 315 графика
void __fastcall TMain::Graf315(double v)
{
  int kol;
  int ogod = god-1;
  int omes=11;

  /*nsm - номер смены последнего месяца,
   dnism - день смены последнего месяца*/

  nsm = DM->qObnovlenie2->FieldByName("nsm")->AsInteger;
  dnism = DM->qObnovlenie2->FieldByName("dnism")->AsInteger;

  //по месяцам
  for (mes=1; mes<=12; mes++)
    {
      NextMonth(omes, ogod);
      kol = DaysInAMonth(ogod, omes);

      //по дням с 26 числа прошлого месяца
      for (den=26; den<=kol; den++)
        {
          //рабочий день
          if (nsm==1)
            {
              // праздничный день
              if (PrazdDni(den,omes)==true)
                {
                  vihod[mes][den]=9;
                  chf[mes][den]=30;
                }
              else
                {
                  vihod[mes][den]=1;
                  chf[mes][den]= DM->qOgraf->FieldByName("DLIT")->AsFloat;
                  vchf[mes][den]=v;

                  ochf[mes]+=DM->qOgraf->FieldByName("DLIT")->AsFloat;
                  ovchf[mes]+=v;
                }

              if (dnism==2)
                {
                  nsm=0;
                  dnism=1;
                }
              else
                {
                  dnism++;
                }
            }
          //выходной
          else
            {
              // праздничный день
              if (PrazdDni(den,omes)==true)
                {
                  vihod[mes][den]=9;
                  chf[mes][den]=30;
                }
              else
                {
                  vihod[mes][den]=0;
                  chf[mes][den]=0;
                }

              if (dnism==2)
                {
                  nsm=1;
                  dnism=1;
                }
              else
                {
                  dnism++;
                }
            }
        }

      //отсутствующие дни в месяце
      if (den<32)
        {
          while (den<=32)
            {
              vihod[mes][den]="NULL";
              chf[mes][den]="NULL";
              den++;
            }
        }

      //по дням с 1 числа по 25 текущего месяца
      for (den=1; den<=25; den++)
        {
          //рабочий день
          if (nsm==1)
            {
              // праздничный день
              if (PrazdDni(den,mes)==true)
                {
                  vihod[mes][den]=9;
                  chf[mes][den]=30;
                }
              else
                {
                  vihod[mes][den]=1;
                  chf[mes][den]= DM->qOgraf->FieldByName("DLIT")->AsFloat;
                  vchf[mes][den]=v;

                  ochf[mes]+=DM->qOgraf->FieldByName("DLIT")->AsFloat;
                  ovchf[mes]+=v;
                }

              if (dnism==2)
                {
                  nsm=0;
                  dnism=1;
                }
              else
                {
                  dnism++;
                }
            }
          //выходной
          else
            {
              // праздничный день
              if (PrazdDni(den,mes)==true)
                {
                  vihod[mes][den]=9;
                  chf[mes][den]=30;
                }
              else
                {
                  vihod[mes][den]=0;
                  chf[mes][den]=0;
                }

              if (dnism==2)
                {
                  nsm=1;
                  dnism=1;
                }
              else
                {
                  dnism++;
                }
            }
        }

      DM->qNorma11Graf->Next();
    }
}
//------------------------------------------------------------------------------

//Расчет 320 графика
void __fastcall TMain::Graf320(double v, double n1, double n2, double p1, double p2)
{
   int kol;

  /*nsm - номер смены последнего месяца,
   dnism - день смены последнего месяца*/

  nsm = DM->qObnovlenie2->FieldByName("nsm")->AsInteger;
  dnism = DM->qObnovlenie2->FieldByName("dnism")->AsInteger;

  //по месяцам
  for (mes=1; mes<=12; mes++)
    {
      kol = DaysInAMonth(god, mes);

      //по дням
      for (den=1; den<=kol; den++)
        {
          //рабочий день
          if (nsm==1)
            {
              vihod[mes][den]=1;
              vchf[mes][den]=v;
              ovchf[mes]+=v;

              //если ночная смена попадает на последний день месяца
              if (den==kol)
                {
                  chf[mes][den]=p1;
                  nchf[mes][den]=n1;

                  //общие суммы
                  ochf[mes]+=p1;
                  onchf[mes]+=n1;

                  // праздничный день
                  if (PrazdDni(den,mes)==true)
                    {
                      pchf[mes][den]=p1;
                      opchf[mes]+=p1;
                    }

                  //если день и праздничный и предпраздничный
                  /*if (PrazdDni(den,mes)==true && PrdPrazdDni(den,mes)==true)
                    {
                      pchf[mes][den]=p1+p2;
                      opchf[mes]+=p1;
                    }
                  else
                    {
                      // праздничный день
                      if (PrazdDni(den,mes)==true)
                        {
                          pchf[mes][den]=p1;
                          opchf[mes]+=p1;
                        }
                      //проверка предпраздничного дня
                      else if (PrdPrazdDni(den,mes)==true)
                        {
                          pchf[mes][den]=p2;
                        }
                    } */
                }
              //если ночная смена не попадает на последний день месяца
              else
                {
                  //переход на летнее время (март) (часы+1, ночные+1)
                  if (mes==3 && den==day_mart)
                    {
                      chf[mes][den]=p1;
                      ochf[mes]+=DM->qOgraf->FieldByName("DLIT")->AsFloat-1;

                      nchf[mes][den]=n1;
                      onchf[mes]+=(n1+n2)-1;
                    }
                  //переход на зимнее время (октябрь) (часы-1, ночные-1)
                  else if (mes==10 && den==day_oktyabr)
                    {
                      chf[mes][den]=p1;
                      ochf[mes]+=DM->qOgraf->FieldByName("DLIT")->AsFloat+1;

                      nchf[mes][den]=n1;
                      onchf[mes]+=n1+n2+1;
                    }
                  else
                    {
                      chf[mes][den]=p1;
                      nchf[mes][den]=n1;

                      ochf[mes]+=DM->qOgraf->FieldByName("DLIT")->AsFloat;
                      onchf[mes]+=n1+n2;
                    }

                  // праздничный день
                  if (PrazdDni(den,mes)==true)
                    {
                      pchf[mes][den]=p1;
                      opchf[mes]+=p1;
                    }
                  //если день и праздничный и предпраздничный
                 /* if (PrazdDni(den,mes)==true && PrdPrazdDni(den,mes)==true)
                    {
                      pchf[mes][den]=p1+p2;
                      opchf[mes]+=p1+p2;
                    }
                  else
                    {
                      // праздничный день
                      if (PrazdDni(den,mes)==true)
                        {
                          pchf[mes][den]=p1;
                          opchf[mes]+=p1;
                        }
                      //проверка предпраздничного дня
                      else if (PrdPrazdDni(den,mes)==true)
                        {
                          pchf[mes][den]=p2;
                          opchf[mes]+=p2;
                        }
                    }   */
                 }

              nsm=0;
              dnism=1;
            }
          //выходной
          else
            {
              vihod[mes][den]=0;
              vchf[mes][den]="NULL";
              nchf[mes][den]="NULL";
              pchf[mes][den]="NULL";

              if (dnism==1)
                {
                  if (mes==3 && den==day_mart2)
                    {
                      chf[mes][den]=p2-1;
                      nchf[mes][den]=n2-1;
                      // праздничный день
                      if (PrazdDni(den,mes)==true)
                        {
                          pchf[mes][den]=p2-1;
                          opchf[mes]+=p2-1;
                        }
                    }
                  else if (mes==10 && den==day_oktyabr2)
                    {
                      chf[mes][den]=p2+1;
                      nchf[mes][den]=n2+1;
                      // праздничный день
                      if (PrazdDni(den,mes)==true)
                        {
                          pchf[mes][den]=p2+1;
                          opchf[mes]+=p2+1;
                        }
                    }
                  else
                    {
                      chf[mes][den]=p2;
                      nchf[mes][den]=n2;
                      // праздничный день
                      if (PrazdDni(den,mes)==true)
                        {
                          pchf[mes][den]=p2;
                          opchf[mes]+=p2;
                        }
                    }
                }
              else
                {
                  chf[mes][den]=0;
                  nchf[mes][den]="NULL";
                }
               //часы переходящие c предыдущего месяца
              if (den==1 && dnism==1 && mes==1)
                {
              //  nchf[mes][den]= Vvod->SetN(DM->qObnovlenie2->FieldByName("nch0")->AsString);   //для части часов переносящихся на следующий месяц
                  ochf[mes]+= Vvod->SetN(DM->qObnovlenie2->FieldByName("chf0")->AsString);
                  onchf[mes]+= Vvod->SetN(DM->qObnovlenie2->FieldByName("nch0")->AsString);
            //      opchf[mes]+= Vvod->SetN(DM->qObnovlenie2->FieldByName("pch0")->AsString);
                }
              else if (den==1 && dnism==1)
                {
                  chf0[mes-1]=p2;
                  ochf[mes]+=p2;
                  nchf0[mes-1]=n2;
                  onchf[mes]+=n2;

                  //nchf[mes][den]=n2;   //для части часов переносящихся на следующий месяц

                  //проверка праздничного дня
                  if (PrazdDni(den,mes)==true)
                    {
                      pchf0[mes-1]=p2;
              //        opchf[mes]+=p2;
                    }
                }

              if (dnism==3)
                {
                  nsm=1;
                  dnism=1;
                }
              else
                {
                  dnism++;
                }
            }

        }

      //расчет переработки
      if ((ochf[mes] - DM->qNorma11Graf->FieldByName("chf")->AsFloat - opchf[mes])>0)
        {
          pgraf[mes] = ochf[mes] - DM->qNorma11Graf->FieldByName("chf")->AsFloat - opchf[mes];
        }

      // сохранение переходящих часов последнего дня в году
      if (mes==12 && dnism==1 && nsm==0)
        {
          chf0[mes]=p2;
          nchf0[mes]=n2;
          pchf0[mes]=p2;
        }

      //отсутствующие дни в месяце
      if (den<32)
        {
          while (den<=32)
            {
              vihod[mes][den]="NULL";
              chf[mes][den]="NULL";
              den++;
            }
        }

       DM->qNorma11Graf->Next();
    }
}
//------------------------------------------------------------------------------

// Расчет 370 графика
void __fastcall TMain::Graf370(double v, double n1, double n2, double p1, double p2)
{
   int kol;

  /*nsm - номер смены последнего месяца,
   dnism - день смены последнего месяца*/

  nsm = DM->qObnovlenie2->FieldByName("nsm")->AsInteger;
  dnism = DM->qObnovlenie2->FieldByName("dnism")->AsInteger;

  //по месяцам
  for (mes=1; mes<=12; mes++)
    {
      kol = DaysInAMonth(god, mes);

      //по дням
      for (den=1; den<=kol; den++)
        {
          //рабочий день
          if (nsm==1)
            {
              vihod[mes][den]=1;
              vchf[mes][den]=v;
              ovchf[mes]+=v;

              //если ночная смена попадает на последний день месяца
              if (den==kol)
                {
                  if (dnism==1)
                    {
                      chf[mes][den]=p1;
                      nchf[mes][den]=n1;
                      // праздничный день
                      if (PrazdDni(den,mes)==true)
                        {
                          pchf[mes][den]=p1;
                          opchf[mes]+=p1;
                        }
                    }
                  else
                    {
                      chf[mes][den]=DM->qOgraf->FieldByName("DLIT")->AsFloat;
                      nchf[mes][den]=n1+n2;
                      // праздничный день
                      if (PrazdDni(den,mes)==true)
                        {
                          pchf[mes][den]=p1+p2;
                          opchf[mes]+=p1+p2;
                        }
                    }

                  //общие суммы
                  ochf[mes]+=p1;
                  onchf[mes]+=n1;


                  //если день и праздничный и предпраздничный
                /*  if (PrazdDni(den,mes)==true && PrdPrazdDni(den,mes)==true)
                    {
                      pchf[mes][den]=p1+p2;
                      opchf[mes]+=p1;
                    }
                  else
                    {
                      // праздничный день
                      if (PrazdDni(den,mes)==true)
                        {
                          pchf[mes][den]=p1;
                          opchf[mes]+=p1;
                        }
                      //проверка предпраздничного дня
                      else if (PrdPrazdDni(den,mes)==true)
                        {
                          pchf[mes][den]=p2;
                        }
                    }  */
                }
              //если ночная смена не попадает на последний день месяца
              else
                {
                  //переход на летнее время (март) (часы+1, ночные+1)
                  if (mes==3 && den==day_mart)
                    {
                      if (dnism==1)
                        {
                          chf[mes][den]=p1;
                          nchf[mes][den]=n1;
                          // праздничный день
                          if (PrazdDni(den,mes)==true)
                            {
                              pchf[mes][den]=p1;
                              opchf[mes]+=p1;
                            }
                        }
                      else
                        {
                          chf[mes][den]=DM->qOgraf->FieldByName("DLIT")->AsFloat;
                          nchf[mes][den]=n1+n2;
                          // праздничный день
                          if (PrazdDni(den,mes)==true)
                            {
                              pchf[mes][den]=p1+p2;
                              opchf[mes]+=p1+p2;
                            }
                        }
                      ochf[mes]+=DM->qOgraf->FieldByName("DLIT")->AsFloat-1;
                      onchf[mes]+=(n1+n2)-1;
                    }
                  //переход на зимнее время (октябрь) (часы-1, ночные-1)
                  else if (mes==10 && den==day_oktyabr)
                    {
                      if (dnism==1)
                        {
                          chf[mes][den]=p1;
                          nchf[mes][den]=n1;
                          // праздничный день
                          if (PrazdDni(den,mes)==true)
                            {
                              pchf[mes][den]=p1;
                              opchf[mes]+=p1;
                            }
                        }
                      else
                        {
                          chf[mes][den]=DM->qOgraf->FieldByName("DLIT")->AsFloat;
                          nchf[mes][den]=n1+n2;
                          // праздничный день
                          if (PrazdDni(den,mes)==true)
                            {
                              pchf[mes][den]=p1+p2;
                              opchf[mes]+=p1+p2;
                            }
                        }
                      ochf[mes]+=DM->qOgraf->FieldByName("DLIT")->AsFloat+1;
                      onchf[mes]+=(n1+n2)+1;
                    }
                  else
                    {
                      if (mes==3 && den==day_mart2)
                        {
                          if (dnism==2)
                            {
                              chf[mes][den]=DM->qOgraf->FieldByName("DLIT")->AsFloat-1;
                              nchf[mes][den]=(n1+n2)-1;
                              // праздничный день
                              if (PrazdDni(den,mes)==true)
                                {
                                  pchf[mes][den]=(p1+p2)-1;
                                  opchf[mes]+=(p1+p2)-1;
                                }
                            }
                          else
                            {
                              chf[mes][den]=p1;
                              nchf[mes][den]=n1;
                              // праздничный день
                              if (PrazdDni(den,mes)==true)
                                {
                                  pchf[mes][den]=p1;
                                  opchf[mes]+=p1;
                                }
                            }
                        }
                      else if (mes==10 && den==day_oktyabr2)
                        {
                          if (dnism==2)
                            {
                              chf[mes][den]=DM->qOgraf->FieldByName("DLIT")->AsFloat+1;
                              nchf[mes][den]=(n1+n2)+1;
                              // праздничный день
                              if (PrazdDni(den,mes)==true)
                                {
                                  pchf[mes][den]=p1+p2+1;
                                  opchf[mes]+=p1+p2+1;
                                }
                            }
                          else
                            {
                              chf[mes][den]=p1;
                              nchf[mes][den]=n1;
                              // праздничный день
                              if (PrazdDni(den,mes)==true)
                                {
                                  pchf[mes][den]=p1;
                                  opchf[mes]+=p1;
                                }
                            }
                        }
                      else
                        {
                          if (dnism==1)
                            {
                              chf[mes][den]=p1;
                              nchf[mes][den]=n1;
                              // праздничный день
                              if (PrazdDni(den,mes)==true)
                                {
                                  pchf[mes][den]=p1;
                                  opchf[mes]+=p1;
                                }
                            }
                          else
                            {
                              chf[mes][den]=DM->qOgraf->FieldByName("DLIT")->AsFloat;
                              nchf[mes][den]=n1+n2;
                              // праздничный день
                              if (PrazdDni(den,mes)==true)
                                {
                                  pchf[mes][den]=p1+p2;
                                  opchf[mes]+=p1+p2;
                                }
                            }
                        }

                      ochf[mes]+=DM->qOgraf->FieldByName("DLIT")->AsFloat;
                      onchf[mes]+=n1+n2;
                    }


                 /* //если день и праздничный и предпраздничный
                  if (PrazdDni(den,mes)==true && PrdPrazdDni(den,mes)==true)
                    {
                      pchf[mes][den]=p1+p2;
                      opchf[mes]+=p1+p2;
                    }
                  else
                    {
                      // праздничный день
                      if (PrazdDni(den,mes)==true)
                        {
                          pchf[mes][den]=p1;
                          opchf[mes]+=p1;
                        }
                      //проверка предпраздничного дня
                      else if (PrdPrazdDni(den,mes)==true)
                        {
                          pchf[mes][den]=p2;
                          opchf[mes]+=p2;
                        }
                    }  */
                 }

              //часы переходящие c предыдущего месяца
              if (den==1 && ((dnism==1 && nsm==0) || (dnism==2 && nsm==1)) && mes==1)
                {
                  //nchf[mes][den]= Vvod->SetN(DM->qObnovlenie2->FieldByName("nch0")->AsString);  //для части часов переносящихся на следующий месяц
                  ochf[mes]+= Vvod->SetN(DM->qObnovlenie2->FieldByName("chf0")->AsString);
                  onchf[mes]+= Vvod->SetN(DM->qObnovlenie2->FieldByName("nch0")->AsString);
      //            opchf[mes]+= Vvod->SetN(DM->qObnovlenie2->FieldByName("pch0")->AsString);
                }
              else if (den==1 && ((dnism==1 && nsm==0) || (dnism==2 && nsm==1)))
                {
                  chf0[mes-1]=p2;
                  ochf[mes]+=p2;
                  nchf0[mes-1]=n2;
                  onchf[mes]+=n2;

                  //  nchf[mes][den]=n2;   //для части часов переносящихся на следующий месяц

                  //проверка праздничного дня
                  if (PrazdDni(den,mes)==true)
                    {
                      pchf0[mes-1]=p2;
                   //   opchf[mes]+=p2;
                    }
                }

              if (dnism==2)
                {
                  nsm=0;
                  dnism=1;
                }
              else
                {
                  dnism++;
                }

            }
          //выходной
          else
            {
              vihod[mes][den]=0;
              pchf[mes][den]="NULL";

              if (dnism==1)
                {
                  if (mes==3 && den==day_mart2)
                    {
                      chf[mes][den]=p2-1;
                      nchf[mes][den]=n2-1;
                      // праздничный день
                      if (PrazdDni(den,mes)==true)
                        {
                          pchf[mes][den]=p2-1;
                          opchf[mes]+=p2-1;
                        }
                    }
                  else if (mes==10 && den==day_oktyabr2)
                    {
                      chf[mes][den]=p2+1;
                      nchf[mes][den]=n2+1;
                      // праздничный день
                      if (PrazdDni(den,mes)==true)
                        {
                          pchf[mes][den]=p2+1;
                          opchf[mes]+=p2+1;
                        }
                    }
                  else
                    {
                      chf[mes][den]=p2;
                      nchf[mes][den]=n2;
                      // праздничный день
                      if (PrazdDni(den,mes)==true)
                        {
                          pchf[mes][den]=p2;
                          opchf[mes]+=p2;
                        }
                    }
                }
              else
                {
                  chf[mes][den]=0;
                  nchf[mes][den]="NULL";
                }

              vchf[mes][den]="NULL";


              //часы переходящие c предыдущего месяца
              if (den==1 && ((dnism==1 && nsm==0) || (dnism==2 && nsm==1)) && mes==1)
                {
                  //nchf[mes][den]= Vvod->SetN(DM->qObnovlenie2->FieldByName("nch0")->AsString);  //для части часов переносящихся на следующий месяц
                  ochf[mes]+= Vvod->SetN(DM->qObnovlenie2->FieldByName("chf0")->AsString);
                  onchf[mes]+= Vvod->SetN(DM->qObnovlenie2->FieldByName("nch0")->AsString);
              //    opchf[mes]+= Vvod->SetN(DM->qObnovlenie2->FieldByName("pch0")->AsString);
                }
              else if (den==1 && ((dnism==1 && nsm==0) || (dnism==2 && nsm==1)))
                {
                  chf0[mes-1]=p2;
                  ochf[mes]+=p2;
                  nchf0[mes-1]=n2;
                  onchf[mes]+=n2;

              //  nchf[mes][den]=n2;   //для части часов переносящихся на следующий месяц


                  //проверка праздничного дня
                  if (PrazdDni(den,mes)==true)
                    {
                      pchf0[mes-1]=p2;
                  //    opchf[mes]+=p2;
                    }
                }

              if (dnism==2)
                {
                  nsm=1;
                  dnism=1;
                }
              else
                {
                  dnism++;
                }
            }
        }

      //расчет переработки
      if ((ochf[mes] - DM->qNorma11Graf->FieldByName("chf")->AsFloat - opchf[mes])>0)
        {
          pgraf[mes] = ochf[mes] - DM->qNorma11Graf->FieldByName("chf")->AsFloat - opchf[mes];
        }

      // сохранение переходящих часов последнего дня в году
      if ((mes==12 && dnism==1 && nsm==0) || mes==12 && dnism==2 && nsm==1)
        {
          chf0[mes]=p2;
          nchf0[mes]=n2;
          pchf0[mes]=p2;
        }

      //отсутствующие дни в месяце
      if (den<32)
        {
          while (den<=32)
            {
              vihod[mes][den]="NULL";
              chf[mes][den]="NULL";
              den++;
            }
        }

       DM->qNorma11Graf->Next();
    }
}

//------------------------------------------------------------------------------

// Расчет 390 графика
void __fastcall TMain::Graf390(double p1, double p2, double n1, double n2)
{
     int kol;

  /*nsm - номер смены последнего месяца,
   dnism - день смены последнего месяца*/

  nsm = DM->qObnovlenie2->FieldByName("nsm")->AsInteger;
  dnism = DM->qObnovlenie2->FieldByName("dnism")->AsInteger;

  //по месяцам
  for (mes=1; mes<=12; mes++)
    {
      kol = DaysInAMonth(god, mes);

      //по дням
      for (den=1; den<=kol; den++)
        {
          //рабочий день
          if (nsm==1)
            {
              vihod[mes][den]=1;
              vchf[mes][den]=4;
              ovchf[mes]+=4;

              //если ночная смена попадает на последний день месяца
              if (den==kol)
                {
                  chf[mes][den]=p1;
                  nchf[mes][den]=n1;

                  //общие суммы
                  ochf[mes]+=p1;
                  onchf[mes]+=n1;

                  // праздничный день
                  if (PrazdDni(den,mes)==true)
                    {
                      pchf[mes][den]=p1;
                      opchf[mes]+=p1;
                    }

                 /* //если день и праздничный и предпраздничный
                  if (PrazdDni(den,mes)==true && PrdPrazdDni(den,mes)==true)
                    {
                      pchf[mes][den]=p1+p2;
                      opchf[mes]+=p1;
                    }
                  else
                    {
                      // праздничный день
                      if (PrazdDni(den,mes)==true)
                        {
                          pchf[mes][den]=p1;
                          opchf[mes]+=p1;
                        }
                      //проверка предпраздничного дня
                      else if (PrdPrazdDni(den,mes)==true)
                        {
                          pchf[mes][den]=p2;
                        }
                    } */
                }
              //если ночная смена не попадает на последний день месяца
              else
                {
                  chf[mes][den]=p1;

                  //переход на летнее время (март) (часы+1, ночные+1)
                  if (mes==3 && den==day_mart)
                    {
                      chf[mes][den]=p1;
                      ochf[mes]+=DM->qOgraf->FieldByName("DLIT")->AsFloat-1;

                      nchf[mes][den]=n1;
                      onchf[mes]+=(n1+n2)-1;
                    }
                  //переход на зимнее время (октябрь) (часы-1, ночные-1)
                  else if (mes==10 && den==day_oktyabr)
                    {
                      chf[mes][den]=p1;
                      ochf[mes]+=DM->qOgraf->FieldByName("DLIT")->AsFloat+1;

                      nchf[mes][den]=n1;
                      onchf[mes]+=(n1+n2)+1;
                    }
                  else
                    {
                      chf[mes][den]=p1;
                      nchf[mes][den]=n1;

                      ochf[mes]+=DM->qOgraf->FieldByName("DLIT")->AsFloat;
                      onchf[mes]+=n1+n2;
                   }

                 // праздничный день
                 if (PrazdDni(den,mes)==true)
                   {
                     pchf[mes][den]=p1;
                     opchf[mes]+=p1;
                   }
                 
                  /*//если день и праздничный и предпраздничный
                  if (PrazdDni(den,mes)==true && PrdPrazdDni(den,mes)==true)
                    {
                      pchf[mes][den]=p1+p2;
                      opchf[mes]+=p1+p2;
                    }
                  else
                    {
                      // праздничный день
                      if (PrazdDni(den,mes)==true)
                        {
                          pchf[mes][den]=p1;
                          opchf[mes]+=p1;
                        }
                      //проверка предпраздничного дня
                      else if (PrdPrazdDni(den,mes)==true)
                        {
                          pchf[mes][den]=p2;
                          opchf[mes]+=p2;
                        }
                    } */
                 }

              nsm=0;
              dnism=1;
            }
          //выходной
          else
            {
              vihod[mes][den]=0;
              vchf[mes][den]="NULL";
              pchf[mes][den]="NULL";

              if (dnism==1)
                {
                  //переход на летнее время (март) (часы+1, ночные+1)
                  if (mes==3 && den==day_mart2)
                    {
                      chf[mes][den]=p2-1;
                      nchf[mes][den]=n2-1;
                      // праздничный день
                      if (PrazdDni(den,mes)==true)
                        {
                          pchf[mes][den]=p2-1;
                          opchf[mes]+=p2-1;
                        }  
                    }
                  //переход на зимнее время (октябрь) (часы-1, ночные-1)
                  else if (mes==10 && den==day_oktyabr2)
                    {
                      chf[mes][den]=p2+1;
                      nchf[mes][den]=n2+1;
                      // праздничный день
                      if (PrazdDni(den,mes)==true)
                        {
                          pchf[mes][den]=p2+1;
                          opchf[mes]+=p2+1;
                        }
                    }
                  else
                    {
                      chf[mes][den]=p2;
                      nchf[mes][den]=n2;
                      // праздничный день
                      if (PrazdDni(den,mes)==true)
                        {
                          pchf[mes][den]=p2;
                          opchf[mes]+=p2;
                        }
                    }
                }
              else
                {
                  chf[mes][den]=0;
                  nchf[mes][den]="NULL";
                }

              //часы переходящие c предыдущего месяца
              if (den==1 && dnism==1 && mes==1)
                {
                  ochf[mes]+= Vvod->SetN(DM->qObnovlenie2->FieldByName("chf0")->AsString);
                  onchf[mes]+= Vvod->SetN(DM->qObnovlenie2->FieldByName("nch0")->AsString);
             //     opchf[mes]+= Vvod->SetN(DM->qObnovlenie2->FieldByName("pch0")->AsString);
                }
              else if (den==1 && dnism==1)
                {
                  chf0[mes-1]=p2;
                  ochf[mes]+=p2;
                  nchf0[mes-1]=n2;
                  onchf[mes]+=n2;

                  //проверка праздничного дня
                  if (PrazdDni(den,mes)==true)
                    {
                      pchf0[mes-1]=p2;
            //          opchf[mes]+=p2;
                    }
                }

              if (dnism==3)
                {
                  nsm=1;
                  dnism=1;
                }
              else
                {
                  dnism++;
                }
            }
        }

      //расчет переработки
      if ((ochf[mes] - DM->qNorma11Graf->FieldByName("chf")->AsFloat - opchf[mes])>0)
        {
          pgraf[mes] = ochf[mes] - DM->qNorma11Graf->FieldByName("chf")->AsFloat - opchf[mes];
        }

      // сохранение переходящих часов последнего дня в году
      if (mes==12 && dnism==1 && nsm==0)
        {
          chf0[mes]=p2;
          nchf0[mes]=n2;
          pchf0[mes]=p2;
        }

      //отсутствующие дни в месяце
      if (den<32)
        {
          while (den<=32)
            {
              vihod[mes][den]="NULL";
              chf[mes][den]="NULL";
              den++;
            }
        }

       DM->qNorma11Graf->Next();
    }
}

//------------------------------------------------------------------------------

// Расчет 400 графика
void __fastcall TMain::Graf400(double v)
{
  int kol;
  int ogod = god-1;
  int omes=11;

  /*nsm - номер смены последнего месяца,
   dnism - день смены последнего месяца*/

  nsm = DM->qObnovlenie2->FieldByName("nsm")->AsInteger;
  dnism = DM->qObnovlenie2->FieldByName("dnism")->AsInteger;

  //по месяцам
  for (mes=1; mes<=12; mes++)
    {
      kol = DaysInAMonth(god, mes);

      NextMonth(omes, ogod);
     // PrevMonth(mes, ogod);
      kol = DaysInAMonth(ogod, omes);

      //по дням с 26 числа прошлого месяца
      for (den=26; den<=kol; den++)
        {
          //рабочий день
          if (nsm==1)
            {
              vihod[mes][den]=1;
              chf[mes][den]= DM->qOgraf->FieldByName("DLIT")->AsFloat;
              vchf[mes][den]=v;

              //общие суммы
              ochf[mes]+=DM->qOgraf->FieldByName("DLIT")->AsFloat;
              ovchf[mes]+=v;

              // праздничный день
              if (PrazdDni(den,omes)==true)
                {
                  pchf[mes][den]=DM->qOgraf->FieldByName("DLIT")->AsFloat;
                  opchf[mes]+=DM->qOgraf->FieldByName("DLIT")->AsFloat;
                }

              if (dnism==2)
                {
                  nsm=0;
                  dnism=1;
                }
              else
                {
                  dnism++;
                }
            }
          //выходной
          else
            {
              vihod[mes][den]=0;
              chf[mes][den]=0;

              if (dnism==2)
                {
                  nsm=1;
                  dnism=1;
                }
              else
                {
                  dnism++;
                }
            }
        }

      //отсутствующие дни в месяце
      if (den<32)
        {
          while (den<=32)
            {
              vihod[mes][den]="NULL";
              chf[mes][den]="NULL";
              den++;
            }
        }

      //по дням
      for (den=1; den<=25; den++)
        {
          //рабочий день
          if (nsm==1)
            {
              vihod[mes][den]=1;
              chf[mes][den]= DM->qOgraf->FieldByName("DLIT")->AsFloat;
              vchf[mes][den]=v;

              //общие суммы
              ochf[mes]+=DM->qOgraf->FieldByName("DLIT")->AsFloat;
              ovchf[mes]+=v;

              // праздничный день
              if (PrazdDni(den,mes)==true)
                {
                  pchf[mes][den]=DM->qOgraf->FieldByName("DLIT")->AsFloat;
                  opchf[mes]+=DM->qOgraf->FieldByName("DLIT")->AsFloat;
                }

              if (dnism==2)
                {
                  nsm=0;
                  dnism=1;
                }
              else
                {
                  dnism++;
                }
            }
          //выходной
          else
            {
              vihod[mes][den]=0;
              chf[mes][den]=0;

              if (dnism==2)
                {
                  nsm=1;
                  dnism=1;
                }
              else
                {
                  dnism++;
                }
            }
        }

       //расчет переработки
      if ((ochf[mes] - DM->qNorma11Graf->FieldByName("chf")->AsFloat - opchf[mes])>0)
        {
          pgraf[mes] = ochf[mes] - DM->qNorma11Graf->FieldByName("chf")->AsFloat - opchf[mes];
        }
          
      DM->qNorma11Graf->Next();
    }
}
//------------------------------------------------------------------------------

// Расчет 410 графика
void __fastcall TMain::Graf410(double v)
{
  int kol;

  /*nsm - номер смены последнего месяца,
   dnism - день смены последнего месяца*/

  nsm = DM->qObnovlenie2->FieldByName("nsm")->AsInteger;
  dnism = DM->qObnovlenie2->FieldByName("dnism")->AsInteger;

  //по месяцам
  for (mes=1; mes<=12; mes++)
    {
      kol = DaysInAMonth(god, mes);

      //по дням
   for (den=1; den<=kol; den++)
        {
          //рабочий день
          if (nsm==1)
            {
              // праздничный день
              if (PrazdDni(den,mes)==true)
                {
                  vihod[mes][den]=9;
                  chf[mes][den]=30;
                }
              else
                {
                  vihod[mes][den]=1;
                  chf[mes][den]= DM->qOgraf->FieldByName("DLIT")->AsFloat;
                  vchf[mes][den]=v;

                  ochf[mes]+=DM->qOgraf->FieldByName("DLIT")->AsFloat;
                  ovchf[mes]+=v;
                }

              if (dnism==2)
                {
                  nsm=0;
                  dnism=1;
                }
              else
                {
                  dnism++;
                }
            }
          //выходной
          else
            {
              // праздничный день
              if (PrazdDni(den,mes)==true)
                {
                  vihod[mes][den]=9;
                  chf[mes][den]=30;
                }
              else
                {
                  vihod[mes][den]=0;
                  chf[mes][den]=0;
                }  

              if (dnism==2)
                {
                  nsm=1;
                  dnism=1;
                }
              else
                {
                  dnism++;
                }
            }
        }

      //отсутствующие дни в месяце
      if (den<32)
        {
          while (den<=32)
            {
              vihod[mes][den]="NULL";
              chf[mes][den]="NULL";
              den++;
            }
        }
      DM->qNorma11Graf->Next();
    }
}
//------------------------------------------------------------------------------

// Расчет 450 графика
void __fastcall TMain::Graf450(double v, double n1, double n2, double p1, double p2)
{
  int kol;

  /*nsm - номер смены последнего месяца,
   dnism - день смены последнего месяца*/

  nsm = DM->qObnovlenie2->FieldByName("nsm")->AsInteger;
  dnism = DM->qObnovlenie2->FieldByName("dnism")->AsInteger;

  //по месяцам
  for (mes=1; mes<=12; mes++)
    {
      kol = DaysInAMonth(god, mes);

      //по дням
      for (den=1; den<=kol; den++)
        {
          //*************
          //рабочий день
          if (nsm==1)
            {
              vihod[mes][den]=1;
              vchf[mes][den]=v;
              ovchf[mes]+=v;

              //если ночная смена попадает на последний день месяца
              if (den==kol)
                {
                  /*//если день и праздничный и предпраздничный
                  if (PrazdDni(den,mes)==true && PrdPrazdDni(den,mes)==true)
                    {
                      if (dnism==1)
                        {
                          chf[mes][den]=p1;
                          nchf[mes][den]=n1;
                        }
                      else
                        {
                          chf[mes][den]=DM->qOgraf->FieldByName("DLIT")->AsFloat;
                          nchf[mes][den]=n1+n2;
                        }

                      pchf[mes][den]=p1+p2;

                      //общие суммы
                      ochf[mes]+=p1;
                      onchf[mes]+=n1;
                      opchf[mes]+=p1;
                    }
                  else
                    { */
                      // праздничный день
                      if (PrazdDni(den,mes)==true)
                        {
                          if (dnism==1)
                            {
                              chf[mes][den]=p1;
                              nchf[mes][den]=n1;
                              pchf[mes][den]=p1;
                              opchf[mes]+=p1;
                            }
                          else
                            {
                              chf[mes][den]=DM->qOgraf->FieldByName("DLIT")->AsFloat;
                              nchf[mes][den]=n1+n2;
                              pchf[mes][den]=p1+p2;
                              opchf[mes]+=p1+p2;
                            }

                           //общие суммы
                          ochf[mes]+=p1;
                          onchf[mes]+=n1;
                        }
                      //проверка предпраздничного дня
                     /* else if (PrdPrazdDni(den,mes)==true)
                        {
                          if (dnism==1)
                            {
                              chf[mes][den]=p1;
                              nchf[mes][den]=n1;
                            }
                          else
                            {
                              chf[mes][den]=DM->qOgraf->FieldByName("DLIT")->AsFloat;
                              nchf[mes][den]=n1+n2;
                            }

                          pchf[mes][den]=p2;

                          //общие суммы
                          ochf[mes]+=DM->qOgraf->FieldByName("DLIT")->AsFloat-p2;
                          onchf[mes]+=n1;  
                        }         */
                      else
                        {
                          if (dnism==1)
                            {
                              chf[mes][den]=p1;
                              nchf[mes][den]=n1;
                            }
                          else
                            {
                              chf[mes][den]=DM->qOgraf->FieldByName("DLIT")->AsFloat;
                              nchf[mes][den]=n1+n2;
                            }

                          //общие суммы
                          ochf[mes]+=DM->qOgraf->FieldByName("DLIT")->AsFloat-p2;
                          onchf[mes]+=n1;
                        }
                   // }
                }
              //если ночная смена не попадает на последний день месяца
              else
                {
                  //переход на летнее время (март) (часы+1, ночные+1)
                  if (mes==3 && den==day_mart)
                    {
                      // праздничный день
                      if (PrazdDni(den,mes)==true)
                        {
                          if (dnism==1)
                            {
                              chf[mes][den]=p1;
                              nchf[mes][den]=n1;
                              pchf[mes][den]=p1;
                              opchf[mes]+=p1;
                            }
                          else
                            {
                              chf[mes][den]=DM->qOgraf->FieldByName("DLIT")->AsFloat;
                              nchf[mes][den]=n1+n2;
                              pchf[mes][den]=p1+p2;
                              opchf[mes]+=p1+p2;
                            }

                          ochf[mes]+=DM->qOgraf->FieldByName("DLIT")->AsFloat-1;
                          onchf[mes]+=(n1+n2)-1;
                          
                        }
                      else
                        {
                          if (dnism==1)
                            {
                              chf[mes][den]=p1;
                              nchf[mes][den]=n1;
                            }
                          else
                            {
                              chf[mes][den]=DM->qOgraf->FieldByName("DLIT")->AsFloat;
                              nchf[mes][den]=n1+n2;
                            }

                          ochf[mes]+=DM->qOgraf->FieldByName("DLIT")->AsFloat-1;
                          onchf[mes]+=(n1+n2)-1;
                        }
                    }
                  //переход на зимнее время (октябрь) (часы-1, ночные-1)
                  else if (mes==10 && den==day_oktyabr)
                    {
                      // праздничный день
                      if (PrazdDni(den,mes)==true)
                        {
                          if (dnism==1)
                            {
                              chf[mes][den]=p1;
                              nchf[mes][den]=n1;
                              pchf[mes][den]=p1;
                              opchf[mes]+=p1;
                            }
                          else
                            {
                              chf[mes][den]=DM->qOgraf->FieldByName("DLIT")->AsFloat;
                              nchf[mes][den]=n1+n2;
                              pchf[mes][den]=p1+p2;
                              opchf[mes]+=p1+p2;
                            }

                          ochf[mes]+=DM->qOgraf->FieldByName("DLIT")->AsFloat+1;
                          onchf[mes]+=(n1+n2)+1;
                          
                        }
                      else
                        {
                          if (dnism==1)
                            {
                              chf[mes][den]=p1;
                              nchf[mes][den]=n1;
                            }
                          else
                            {
                              chf[mes][den]=DM->qOgraf->FieldByName("DLIT")->AsFloat;
                              nchf[mes][den]=n1+n2;
                            }

                          ochf[mes]+=DM->qOgraf->FieldByName("DLIT")->AsFloat+1;
                          onchf[mes]+=(n1+n2)+1;
                        }

                    }
                  else
                    {
                      if (mes==3 && den==day_mart2)
                        {
                          if (dnism==2)
                            {
                              chf[mes][den]=DM->qOgraf->FieldByName("DLIT")->AsFloat-1;
                              nchf[mes][den]=(n1+n2)-1;
                              // праздничный день
                              if (PrazdDni(den,mes)==true)
                                {
                                  pchf[mes][den]=(p1+p2)-1;
                                  opchf[mes]+=(p1+p2)-1;
                                }

                            }
                          else
                            {
                              chf[mes][den]=p1;
                              nchf[mes][den]=n1;
                              // праздничный день
                              if (PrazdDni(den,mes)==true)
                                {
                                  pchf[mes][den]=p1;
                                  opchf[mes]+=p1;
                                }
                            }
                        }
                      else if (mes==10 && den==day_oktyabr2)
                        {
                          if (dnism==2)
                            {
                              chf[mes][den]=DM->qOgraf->FieldByName("DLIT")->AsFloat+1;
                              nchf[mes][den]=(n1+n2)+1;
                              // праздничный день
                              if (PrazdDni(den,mes)==true)
                                {
                                  pchf[mes][den]=(p1+p2)+1;
                                  opchf[mes]+=(p1+p2)+1;
                                }
                            }
                          else
                            {
                              chf[mes][den]=p1;
                              nchf[mes][den]=n1;
                              // праздничный день
                              if (PrazdDni(den,mes)==true)
                                {
                                  pchf[mes][den]=p1;
                                  opchf[mes]+=p1;
                                }
                            }
                        }
                      else
                        {
                          if (dnism==1)
                            {
                              chf[mes][den]=p1;
                              nchf[mes][den]=n1;
                              // праздничный день
                              if (PrazdDni(den,mes)==true)
                                {
                                  pchf[mes][den]=p1;
                                  opchf[mes]+=p1;
                                }
                            }
                          else
                            {
                              chf[mes][den]=DM->qOgraf->FieldByName("DLIT")->AsFloat;
                              nchf[mes][den]=(n1+n2);
                              // праздничный день
                              if (PrazdDni(den,mes)==true)
                                {
                                  pchf[mes][den]=p1+p2;
                                  opchf[mes]+=p1+p2;
                                }
                            }
                        }

                      ochf[mes]+=DM->qOgraf->FieldByName("DLIT")->AsFloat;
                      onchf[mes]+=(n1+n2);


                      //если день и праздничный и предпраздничный
                      /*if (PrazdDni(den,mes)==true && PrdPrazdDni(den,mes)==true)
                        {
                          pchf[mes][den]=p1+p2;

                          ochf[mes]+=DM->qOgraf->FieldByName("DLIT")->AsFloat;
                          onchf[mes]+=(n1+n2);
                          opchf[mes]+=p1+p2;
                        }
                      else
                        {
                          // праздничный день
                          if (PrazdDni(den,mes)==true)
                            {
                              pchf[mes][den]=p1;

                              ochf[mes]+=DM->qOgraf->FieldByName("DLIT")->AsFloat;
                              onchf[mes]+=(n1+n2);
                              opchf[mes]+=p1;
                            }
                          //проверка предпраздничного дня
                          else if (PrdPrazdDni(den,mes)==true)
                            {
                              pchf[mes][den]=p2;

                              ochf[mes]+=DM->qOgraf->FieldByName("DLIT")->AsFloat;
                              onchf[mes]+=(n1+n2);
                              opchf[mes]+=p2;
                            }
                          else
                            {
                              ochf[mes]+=DM->qOgraf->FieldByName("DLIT")->AsFloat;
                              onchf[mes]+=(n1+n2);
                            }
                        }*/
                    } 
                 }

              //часы переходящие c предыдущего месяца
              if (den==1 && ((dnism==1 && nsm==0) || (dnism==2 && nsm==1)) && mes==1)
                {
                  ochf[mes]+= Vvod->SetN(DM->qObnovlenie2->FieldByName("chf0")->AsString);
                  onchf[mes]+= Vvod->SetN(DM->qObnovlenie2->FieldByName("nch0")->AsString);
                //  opchf[mes]+= Vvod->SetN(DM->qObnovlenie2->FieldByName("pch0")->AsString);
                }
              else if (den==1 && ((dnism==1 && nsm==0) || (dnism==2 && nsm==1)))
                {
                  chf0[mes-1]=p2;
                  ochf[mes]+=p2;
                  nchf0[mes-1]=n2;
                  onchf[mes]+=n2;

                  //проверка праздничного дня
                  if (PrazdDni(den,mes)==true)
                    {
                      pchf0[mes-1]=p2;
                    //  opchf[mes]+=p2;
                    }
                }

              if (dnism==2)
                {
                  nsm=0;
                  dnism=1;
                }
              else
                {
                  dnism++;
                }

            }
          //выходной
          else
            {
              vihod[mes][den]=0;

              if (dnism==1)
                {
                  if (mes==3 && den==day_mart2)
                    {
                      chf[mes][den]=p2-1;
                      nchf[mes][den]=n2-1;
                      // праздничный день
                      if (PrazdDni(den,mes)==true)
                        {
                          pchf[mes][den]=p2-1;
                          opchf[mes]+=p2-1;
                        }
                    }
                  else if (mes==10 && den==day_oktyabr2)
                    {
                      chf[mes][den]=p2+1;
                      nchf[mes][den]=n2+1;
                      // праздничный день
                      if (PrazdDni(den,mes)==true)
                        {
                          pchf[mes][den]=p2+1;
                          opchf[mes]+=p2+1;
                        }
                    }
                  else
                    {
                      chf[mes][den]=p2;
                      nchf[mes][den]=n2;
                      // праздничный день
                      if (PrazdDni(den,mes)==true)
                        {
                          pchf[mes][den]=p2;
                          opchf[mes]+=p2;
                        }
                    }
                }
              else
                {
                  chf[mes][den]=0;
                  nchf[mes][den]="NULL";
                }

              //часы переходящие c предыдущего месяца
              if (den==1 && ((dnism==1 && nsm==0) || (dnism==2 && nsm==1)) && mes==1)
                {
                  ochf[mes]+= Vvod->SetN(DM->qObnovlenie2->FieldByName("chf0")->AsString);
                  onchf[mes]+= Vvod->SetN(DM->qObnovlenie2->FieldByName("nch0")->AsString);
                //  opchf[mes]+= Vvod->SetN(DM->qObnovlenie2->FieldByName("pch0")->AsString);
                }
              else if (den==1 && ((dnism==1 && nsm==0) || (dnism==2 && nsm==1)))
                {
                  chf0[mes-1]=p2;
                  ochf[mes]+=p2;
                  nchf0[mes-1]=n2;
                  onchf[mes]+=n2;

                  //проверка праздничного дня
                  if (PrazdDni(den,mes)==true)
                    {
                      pchf0[mes-1]=p2;
                    //  opchf[mes]+=p2;
                    }
                }

              if (dnism==2)
                {
                  nsm=1;
                  dnism=1;
                }
              else
                {
                  dnism++;
                }
            }
        }

      //расчет переработки
      if ((ochf[mes] - DM->qNorma11Graf->FieldByName("chf")->AsFloat - opchf[mes])>0)
        {
          pgraf[mes] = ochf[mes] - DM->qNorma11Graf->FieldByName("chf")->AsFloat - opchf[mes];
        }

      // сохранение переходящих часов последнего дня в году
      if ((mes==12 && dnism==1 && nsm==0) || (mes==12 && dnism==2 && nsm==1))
        {
          chf0[mes]=p2;
          nchf0[mes]=n2;
          pchf0[mes]=p2;
        }

      //отсутствующие дни в месяце
      if (den<32)
        {
          while (den<=32)
            {
              vihod[mes][den]="NULL";
              chf[mes][den]="NULL";
              den++;
            }
        }

       DM->qNorma11Graf->Next();
    }
}
//------------------------------------------------------------------------------

//Расчет 470 графика
void __fastcall TMain::Graf470(double v, double n)
{
  int kol, prazd;

  /* chf[32] - рабочие часы по дням
     chf[den] = 8 - рабочий день
     chf[den] = 7 - предпраздничный день
     vihod[32] - выходы по дням (рабочий, отдых, праздничный)
     vihod[den] = 1 - рабочий день
     vihod[den] = 9 - праздник
     prazd - количество часов на которое сокращается предпраздничная смена
  */

  //по месяцам
  for (mes=1; mes<=12; mes++)
    {
      kol = DaysInAMonth(god, mes);

      //по дням
      for (den=1; den<=kol; den++)
        {
          //проверка дня недели
          if (DayWeek(den,mes,god)==1||DayWeek(den,mes,god)==7)
            {
              //проверка праздничного дня
              if (PrazdDni(den,mes)==true)
                {
                  vihod[mes][den]=9;
                  chf[mes][den]=30;
                }
              else
                {
                  vihod[mes][den]=0;
                  chf[mes][den]=0;
                }
            }
          else
            {
              //проверка праздничного дня
              if (PrazdDni(den,mes)==true)
                {
                  vihod[mes][den]=9;
                  chf[mes][den]=30;
                }
              //праздничный попадающий на субботу или воскресенье
              else if (PrazdDniVihodnue(den,mes,god)==true)
                {
                  vihod[mes][den]=0;
                  chf[mes][den]=0;
                }
              //проверка предпраздничного дня
              else if (PrdPrazdDni(den,mes)==true)
                {
                  vihod[mes][den]=1;
                  chf[mes][den]=DM->qOgraf->FieldByName("DLIT")->AsFloat-1;
                  vchf[mes][den]=v;
                  nchf[mes][den]=n-1;

                  ochf[mes]+= DM->qOgraf->FieldByName("DLIT")->AsFloat-1;
                  ovchf[mes]+=v;
                  onchf[mes]+=n-1;
                }
              //рабочий день
              else
                {
                  vihod[mes][den]=1;
                  chf[mes][den]=DM->qOgraf->FieldByName("DLIT")->AsFloat;
                  vchf[mes][den]=v;
                  nchf[mes][den]=n;

                  ochf[mes]+= DM->qOgraf->FieldByName("DLIT")->AsFloat;
                  ovchf[mes]+=v;
                  onchf[mes]+=n;
                }
            }

        }
      //отсутствующие дни в месяце
      if (den<32)
        {
          while (den<=32)
            {
              vihod[mes][den]="NULL";
              chf[mes][den]="NULL";
              den++;
            }
        }
    }
}
//---------------------------------------------------------------------------

// Расчет 480 графика
void __fastcall TMain::Graf480()
{
  int kol, prazdnik=0;

  /*prazdnik - признак праздника попадающего на воскресенье и
               переносящегося на понедельник*/

  for (mes=1; mes<=12; mes++)
    {
      kol = DaysInAMonth(god, mes);

      //по дням
      for (den=1; den<=kol; den++)
        {
          //выходной (воскресенье)
          if (DayWeek(den,mes,god)==1)
            {
              //проверка праздничного дня
              if (PrazdDni(den,mes)==true)
                {
                  vihod[mes][den]=9;
                  chf[mes][den]=30;
                  prazdnik = 1;
                }
              else
                {
                  vihod[mes][den]=0;
                  chf[mes][den]=0;
                }
            }
          //рабочий день
          else
            {
              if (PrazdDni(den,mes)==true)
                {
                  vihod[mes][den]=9;
                  chf[mes][den]=30;
                }
              //проверка на перенос праздника с воскресенья на понедельник
              else if (prazdnik==1 && (DayWeek(den,mes,god)==2 || DayWeek(den,mes,god)==3))
                {
                  vihod[mes][den]=0;
                  chf[mes][den]=0;
                  prazdnik = 0;
                }
              //проверка праздничного дня
              else
                {
                  //проверка дня перед выходным (суббота) или предпраздничного дня (длительность смены= 5 часов)
                  if (DayWeek(den,mes,god)==7 || PrdPrazdDni(den,mes)==true)
                    {
                      vihod[mes][den]=1;
                      chf[mes][den]=5;

                      //общие суммы по часам
                      ochf[mes]+=5;
                    }
                  //обычный день
                  else
                    {
                      vihod[mes][den]=1;
                      chf[mes][den]=DM->qOgraf->FieldByName("DLIT")->AsFloat;

                      //общие суммы по часам
                      ochf[mes]+=DM->qOgraf->FieldByName("DLIT")->AsFloat;

                    }
                }
            }
        }

     //отсутствующие дни в месяце
      if (den<32)
        {
          while (den<=32)
            {
              vihod[mes][den]="NULL";
              chf[mes][den]="NULL";
              nchf[mes][den]="NULL";
              vchf[mes][den]="NULL";
              pchf[mes][den]="NULL";
              den++;
            }
        }

      DM->qNorma11Graf->Next();
    }
}
//------------------------------------------------------------------------------
// Расчет 520 графика
void __fastcall TMain::Graf520(double d1, double d2, double d3, double p1, double p2,
                               double p3, double p, double v, double n1, double n2)
{
  int kol, k;

  /*nsm - номер смены последнего месяца,
   dnism - день смены последнего месяца
   k - количество выходных дней в графике*/

  nsm = DM->qObnovlenie2->FieldByName("nsm")->AsInteger;
  dnism = DM->qObnovlenie2->FieldByName("dnism")->AsInteger;

  if (graf==520|| graf==210) k=2;
  else k=1;


  //по месяцам
  for (mes=1; mes<=12; mes++)
    {
      kol = DaysInAMonth(god, mes);

      //по дням
      for (den=1; den<=kol; den++)
        {
          //рабочий день
          if (nsm==1)
            {
              vihod[mes][den]=1;

              //переход на летнее время (март) (часы+1, ночные+1)
              if (mes==3 && den==day_mart)
                {
                  vchf[mes][den]=v;

                  //общие суммы
                  ovchf[mes]+=v;

                  // праздничный день
                  if (PrazdDni(den,mes)==true)
                    {
                      chf[mes][den]=p2;
                      nchf[mes][den]=n1;
                      pchf[mes][den]=p2;

                      //общие суммы
                      ochf[mes]+=d2-1;
                      onchf[mes]+=(n1+n2)-1;
                      opchf[mes]+=p2;
                    }
                  // выходной день
                  else if (DayWeek(den,mes,god)==7 || DayWeek(den,mes,god)==1 || PrazdDniVihodnue(den,mes,god)==true)
                    {
                      chf[mes][den]=p2;
                      nchf[mes][den]=n1;

                      ochf[mes]+=d2-1;
                      onchf[mes]+=(n1+n2)-1;

                      /*//предпраздничный день
                      if (PrdPrazdDni(den,mes)==true)
                        {
                          pchf[mes][den]=p-1;
                          opchf[mes]+=p-1;
                        }  */
                    }
                  // предпраздничный день
                  else if (PrdPrazdDni(den,mes)==true)
                    {
                      chf[mes][den]=p3;
                     // pchf[mes][den]=p;
                      nchf[mes][den]=n1;

                      ochf[mes]+=d3-1;
                      onchf[mes]+=(n1+n2)-1;
                     // opchf[mes]+=p;
                    }
                  // будний день
                  else
                    {
                      chf[mes][den]=p1;
                      nchf[mes][den]=n1;

                      ochf[mes]+=d1-1;
                      onchf[mes]+=(n1+n2)-1;
                    }
                }
              //переход на зимнее время (октябрь) (часы-1, ночные-1)
              else if (mes==10 && den==day_oktyabr)
                {
                  vchf[mes][den]=v;

                  //общие суммы
                  ovchf[mes]+=v;

                  // праздничный день
                  if (PrazdDni(den,mes)==true)
                    {
                      chf[mes][den]=p2;
                      nchf[mes][den]=n1;
                      pchf[mes][den]=p2;

                      //общие суммы
                      ochf[mes]+=d2+1;
                      onchf[mes]+=(n1+n2)+1;
                      opchf[mes]+=p2;
                    }
                  // выходной день
                  else if (DayWeek(den,mes,god)==7 || DayWeek(den,mes,god)==1 || PrazdDniVihodnue(den,mes,god)==true)
                    {
                      chf[mes][den]=p2;
                      nchf[mes][den]=n1;

                      ochf[mes]+=d2+1;
                      onchf[mes]+=(n1+n2)+1;

                      //предпраздничный день
                     /* if (PrdPrazdDni(den,mes)==true)
                        {
                          pchf[mes][den]=p+1;
                          opchf[mes]+=p+1;
                        } */
                    }
                  // предпраздничный день
                  else if (PrdPrazdDni(den,mes)==true)
                    {
                      chf[mes][den]=p3;
                      nchf[mes][den]=n1;
                    //  pchf[mes][den]=p;

                      ochf[mes]+=d3+1;
                      onchf[mes]+=(n1+n2)+1;
                     // opchf[mes]+=p;
                    }
                  // будний день
                  else
                    {
                      chf[mes][den]=p1;
                      nchf[mes][den]=n1;

                      ochf[mes]+=d1+1;
                      onchf[mes]+=(n1+n2)+1;
                    }
                }
              else
                {
                  //если ночная смена попадает на последний день месяца
                  if (den==kol)
                    {
                      vchf[mes][den]=v;

                      //общие суммы
                      ovchf[mes]+=v;

                      // праздничный день и предпраздничный день
                     /* if(PrdPrazdDni(den,mes)==true && PrazdDni(den,mes)==true)
                        {
                          chf[mes][den]=p2;
                          nchf[mes][den]=n1;
                          pchf[mes][den]=p2+p;

                          //общие суммы
                          ochf[mes]+=p2;
                          onchf[mes]+=n1;
                          opchf[mes]+=p2;
                        }
                      // праздничный день
                      else*/ if (PrazdDni(den,mes)==true)
                        {
                          chf[mes][den]=p2;
                          nchf[mes][den]=n1;
                          pchf[mes][den]=p2;

                          //общие суммы
                          ochf[mes]+=p2;
                          onchf[mes]+=n1;
                          opchf[mes]+=p2;
                        }
                      // выходной день
                      else if (DayWeek(den,mes,god)==7 || DayWeek(den,mes,god)==1 || PrazdDniVihodnue(den,mes,god)==true)
                        {
                          chf[mes][den]=p2;
                          nchf[mes][den]=n1;

                          ochf[mes]+=p2;
                          onchf[mes]+=n1;

                          //предпраздничный день
                        /*  if (PrdPrazdDni(den,mes)==true)
                            {
                              pchf[mes][den]=p;
                    //          opchf[mes]+=p;
                            }  */
                        }
                      // предпраздничный день
                      else if (PrdPrazdDni(den,mes)==true)
                        {
                          chf[mes][den]=p3;
                          nchf[mes][den]=n1;
                         // pchf[mes][den]=p;

                          ochf[mes]+=p3;
                          onchf[mes]+=n1;
           //               opchf[mes]+=p;
                        }
                      // будний день
                      else
                        {
                          chf[mes][den]=p1;
                          nchf[mes][den]=n1;

                          ochf[mes]+=p1;
                          onchf[mes]+=n1;
                        }
                    }
                  //если ночная смена не попадает на последний день месяца
                  else
                    {
                      vchf[mes][den]=v;

                      //общие суммы
                      ovchf[mes]+=v;

                       // праздничный день и предпраздничный день
                    /*  if(PrdPrazdDni(den,mes)==true && PrazdDni(den,mes)==true)
                        {
                          chf[mes][den]=p2;
                          nchf[mes][den]=n1;
                          pchf[mes][den]=p2+p;

                          //общие суммы
                          ochf[mes]+=d2;
                          onchf[mes]+=(n1+n2);
                          opchf[mes]+=p2+p;
                        }
                      // праздничный день
                      else */if (PrazdDni(den,mes)==true)
                        {
                          chf[mes][den]=p2;
                          nchf[mes][den]=n1;
                          pchf[mes][den]=p2;

                          //общие суммы
                          ochf[mes]+=d2;
                          onchf[mes]+=(n1+n2);
                          opchf[mes]+=p2;
                        }
                      // выходной день
                      else if (DayWeek(den,mes,god)==7 || DayWeek(den,mes,god)==1 || PrazdDniVihodnue(den,mes,god)==true)
                        {
                          chf[mes][den]=p2;
                          nchf[mes][den]=n1;

                          ochf[mes]+=d2;
                          onchf[mes]+=(n1+n2);

                        /*  if (PrdPrazdDni(den,mes)==true)
                            {
                              pchf[mes][den]=p;
                              opchf[mes]+=p;
                            }   */
                        }
                      // предпраздничный день
                      else if (PrdPrazdDni(den,mes)==true)
                        {
                          chf[mes][den]=p3;
                          nchf[mes][den]=n1;
                       //   pchf[mes][den]=p;

                          ochf[mes]+=d3;
                          onchf[mes]+=(n1+n2);
                        //  opchf[mes]+=p;
                        }
                      // будний день
                      else
                        {
                          chf[mes][den]=p1;
                          nchf[mes][den]=n1;

                          ochf[mes]+=d1;
                          onchf[mes]+=(n1+n2);
                        }
                    }
                }

              //часы переходящие c предыдущего месяца
              if (den==1 && ((dnism==1 && nsm==0) || (dnism==2 && nsm==1)) && mes==1)
                {
                  ochf[mes]+= Vvod->SetN(DM->qObnovlenie2->FieldByName("chf0")->AsString);
                  onchf[mes]+= Vvod->SetN(DM->qObnovlenie2->FieldByName("nch0")->AsString);
                 // opchf[mes]+= Vvod->SetN(DM->qObnovlenie2->FieldByName("pch0")->AsString);
                }
              else if (den==1 && ((dnism==1 && nsm==0) || (dnism==2 && nsm==1)))
                {
                  chf0[mes-1]=p;
                  ochf[mes]+=p;
                  nchf0[mes-1]=n2;
                  onchf[mes]+=n2;

                  //проверка праздничного дня
                  if (PrazdDni(den,mes)==true)
                    {
                      pchf0[mes-1]=p;
                   //   opchf[mes]+=p;
                    }
                }

            /*  if (dnism==2)
                {    */
                  nsm=0;
                  dnism=1;
          /*      }
              else
                {
                  dnism++;
                }   */
            }
          //выходной
          else
            {
              vihod[mes][den]=0;
              if (dnism==1)
                {
                  if (mes==3 && den==day_mart2)
                    {
                      chf[mes][den]=p-1;
                      nchf[mes][den]=n2-1;
                      // праздничный день
                      if (PrazdDni(den,mes)==true)
                        {
                          pchf[mes][den]=p-1;
                          opchf[mes]+=p-1;
                        }
                    }
                  else if (mes==10 && den==day_oktyabr2)
                    {
                      chf[mes][den]=p+1;
                      nchf[mes][den]=n2+1;
                      // праздничный день
                      if (PrazdDni(den,mes)==true)
                        {
                          pchf[mes][den]=p+1;
                          opchf[mes]+=p+1;
                        }
                    }
                  else
                    {
                      chf[mes][den]=p;
                      nchf[mes][den]=n2;
                      // праздничный день
                      if (PrazdDni(den,mes)==true)
                        {
                          pchf[mes][den]=p;
                          opchf[mes]+=p;
                        }
                    }
                }
              else
                {
                  chf[mes][den]=0;
                  nchf[mes][den]=NULL;
                }

              //часы переходящие c предыдущего месяца
              if (den==1 && dnism==1 && mes==1)
                {
                  ochf[mes]+= Vvod->SetN(DM->qObnovlenie2->FieldByName("chf0")->AsString);
                  onchf[mes]+= Vvod->SetN(DM->qObnovlenie2->FieldByName("nch0")->AsString);
                //  opchf[mes]+= Vvod->SetN(DM->qObnovlenie2->FieldByName("pch0")->AsString);
                }
              else if (den==1 && dnism==1)
                {
                  chf0[mes-1]=p;
                  ochf[mes]+=p;
                  nchf0[mes-1]=n2;
                  onchf[mes]+=n2;

                  //проверка праздничного дня
                  if (PrazdDni(den,mes)==true)
                    {
                      pchf0[mes-1]=p;
                //      opchf[mes]+=p;
                    }
                }

              if (dnism==k)
                {
                  nsm=1;
                  dnism=1;
                }
              else
                {
                  dnism++;
                }

            }
        }

      // сохранение переходящих часов последнего дня в году
      if (mes==12 && dnism==1 && nsm==0)
        {
          chf0[mes]=p;
          nchf0[mes]=n2;
          pchf0[mes]=p;
        }

      //расчет переработки
      if ((ochf[mes] - DM->qNorma11Graf->FieldByName("chf")->AsFloat - opchf[mes])>0)
        {
          pgraf[mes] = ochf[mes] - DM->qNorma11Graf->FieldByName("chf")->AsFloat - opchf[mes];
        }

      //отсутствующие дни в месяце
      if (den<32)
        {
          while (den<=32)
            {
              vihod[mes][den]="NULL";
              chf[mes][den]="NULL";
              den++;
            }
        }
      DM->qNorma11Graf->Next();
    }
}
//------------------------------------------------------------------------------


// Расчет 630 графика
void __fastcall TMain::Graf630()
{
  int vih1, vih2, kol;

  /*vih1, vih2 - выходные для бригады*/

  switch(br)
    { //1 - воскресенье, 2 - понедельник, 3 - вторник, 4 - среда, 5 - четверг, 6 - пятница, 7 - суббота
      // Для бригады 1
      case 1:
        vih1 = 6;
        vih2 = 7;
      break;

      //Для бригады 2
      case 2:
        vih1 = 1;
        vih2 = 2;
      break;
    }

  for (mes=1; mes<=12; mes++)
    {
      kol = DaysInAMonth(god, mes);

      //по дням
      for (den=1; den<=kol; den++)
        {
          //проверка дня недели
          if (DayWeek(den,mes,god)==vih1||DayWeek(den,mes,god)==vih2)
            {
              vihod[mes][den]=0;
              chf[mes][den]=0;
            }
          else
            {
              vihod[mes][den]=1;
              chf[mes][den]=DM->qOgraf->FieldByName("DLIT")->AsFloat;

              //проверка праздничного дня
              if (PrazdDni(den,mes)==true)
                {
                  pchf[mes][den]=DM->qOgraf->FieldByName("DLIT")->AsFloat;
                  opchf[mes]+=DM->qOgraf->FieldByName("DLIT")->AsFloat;
                }

              //общие суммы по часам
              ochf[mes]+= DM->qOgraf->FieldByName("DLIT")->AsFloat;
            }
        }

     //расчет переработки
     if ((ochf[mes] - DM->qNorma11Graf->FieldByName("chf")->AsFloat - opchf[mes])>0)
       {
         pgraf[mes] = ochf[mes] - DM->qNorma11Graf->FieldByName("chf")->AsFloat - opchf[mes];
       }

     //отсутствующие дни в месяце
      if (den<32)
        {
          while (den<=32)
            {
              vihod[mes][den]="NULL";
              chf[mes][den]="NULL";
              nchf[mes][den]="NULL";
              vchf[mes][den]="NULL";
              pchf[mes][den]="NULL";
              den++;
            }
        }

      DM->qNorma11Graf->Next();
    }
}

//------------------------------------------------------------------------------

//Расчет 650 и 660 графика
void __fastcall TMain::Graf650()
{
  int kol, prazdnik=0;

  /* chf[32] - рабочие часы по дням
     chf[den] = 8 - рабочий день
     chf[den] = 7 - предпраздничный день
     vihod[32] - выходы по дням (рабочий, отдых, праздничный)
     vihod[den] = 1 - рабочий день
     vihod[den] = 9 - праздник
     prazd - количество часов на которое сокращается предпраздничная смена
  */

  //по месяцам
  for (mes=1; mes<=12; mes++)
    {
      kol = DaysInAMonth(god, mes);

      //по дням
      for (den=1; den<=kol; den++)
        {
          //проверка дня недели
          if (DayWeek(den,mes,god)==1)
            {
              //проверка праздничного дня
              if (PrazdDni(den,mes)==true)
                {
                  vihod[mes][den]=9;
                  chf[mes][den]=30;
                  prazdnik=1;
                }
              else
                {
                  vihod[mes][den]=0;
                  chf[mes][den]=0;
                }
            }
          else
            {
              //проверка праздничного дня
              if (PrazdDni(den,mes)==true)
                {
                  vihod[mes][den]=9;
                  chf[mes][den]=30;
                }
              //праздничный попадающий на субботу или воскресенье
              else if (prazdnik==1)
                {
                  vihod[mes][den]=0;
                  chf[mes][den]=0;
                  prazdnik=0;
                }
            /*  //проверка предпраздничного дня
              else if (PrdPrazdDni(den,mes)==true)
                {
                  vihod[mes][den]=1;
                  chf[mes][den]=DM->qOgraf->FieldByName("DLIT")->AsFloat-prazd;
                  ochf[mes]+= DM->qOgraf->FieldByName("DLIT")->AsFloat-prazd;
                }   */
              //рабочий день
              else
                {
                  vihod[mes][den]=1;
                  chf[mes][den]=DM->qOgraf->FieldByName("DLIT")->AsFloat;
                  ochf[mes]+= DM->qOgraf->FieldByName("DLIT")->AsFloat;
                }
            }

        }
      //отсутствующие дни в месяце
      if (den<32)
        {
          while (den<=32)
            {
              vihod[mes][den]="NULL";
              chf[mes][den]="NULL";
              den++;
            }
        }
    }
}
//---------------------------------------------------------------------------

// Расчет 670 и 790 графика
void __fastcall TMain::Graf670(double v)
{
  int kol;

  /*nsm - номер смены последнего месяца,
   dnism - день смены последнего месяца*/

  nsm = DM->qObnovlenie2->FieldByName("nsm")->AsInteger;
  dnism = DM->qObnovlenie2->FieldByName("dnism")->AsInteger;

  //по месяцам
  for (mes=1; mes<=12; mes++)
    {
      kol = DaysInAMonth(god, mes);

      //по дням
      for (den=1; den<=kol; den++)
        {
          //рабочий день
          if (nsm==1)
            {
              vihod[mes][den]=1;
              chf[mes][den]= DM->qOgraf->FieldByName("DLIT")->AsFloat;
              vchf[mes][den]=v;       //vchf[mes][den]=v;

              ochf[mes]+=DM->qOgraf->FieldByName("DLIT")->AsFloat;
              ovchf[mes]+=v;   //ovchf[mes]=v;

              // праздничный день
              if (PrazdDni(den,mes)==true)
                {
                  pchf[mes][den]=DM->qOgraf->FieldByName("DLIT")->AsFloat;
                  opchf[mes]+=DM->qOgraf->FieldByName("DLIT")->AsFloat;
                }
                
              nsm=0;
            }
          //выходной
          else
            {
              vihod[mes][den]=0;
              chf[mes][den]=0;

              nsm=1;
            }
        }

      if (graf!=790)
        {
          //расчет переработки
          if ((ochf[mes] - DM->qNorma11Graf->FieldByName("chf")->AsFloat - opchf[mes])>0)
            {
              pgraf[mes] = ochf[mes] - DM->qNorma11Graf->FieldByName("chf")->AsFloat - opchf[mes];
            }
        }

      //отсутствующие дни в месяце
      if (den<32)
        {
          while (den<=32)
            {
              vihod[mes][den]="NULL";
              chf[mes][den]="NULL";
              den++;
            }
        }
      DM->qNorma11Graf->Next();
    }
}
//------------------------------------------------------------------------------

// Расчет 680 графика
void __fastcall TMain::Graf680(double v)
{
  int kol;
  //
  DM->qNorma11Graf->MoveBy(5);

  nsm = DM->qObnovlenie2->FieldByName("nsm")->AsInteger;
  dnism = DM->qObnovlenie2->FieldByName("dnism")->AsInteger;

  mes_n=6;  //количество месяцев в графике для расчета и отображения
  mes_k=9;

  for (mes=mes_n; mes<=mes_k; mes++)
    {
      kol = DaysInAMonth(god, mes);

      //по дням
      for (den=1; den<=kol; den++)
        {
          // рабочий день
          if (nsm==1)
            {
              // после 17 сентября
              if (mes==9 && den>=17)
                {
                  vihod[mes][den]=1;
                  chf[mes][den]=8;
                  //vchf[mes][den]=v;

                  //проверка праздничного дня
                  if (PrazdDni(den,mes)==true)
                    {
                      pchf[mes][den]=8;
                      opchf[mes]+=8;
                    }

                  //общие суммы по часам
                  ochf[mes]+=8;
                  //   ovchf[mes]+=v;

                  if (dnism==5)
                    {
                      dnism=1;
                      nsm=0;
                    }
                  else
                    {
                      dnism++;
                    }
                }
              //до 17 сентября
              else
                {
                  vihod[mes][den]=1;
                  chf[mes][den]=DM->qOgraf->FieldByName("DLIT")->AsFloat;
                  vchf[mes][den]=v;

                  //проверка праздничного дня
                  if (PrazdDni(den,mes)==true)
                    {
                      pchf[mes][den]=DM->qOgraf->FieldByName("DLIT")->AsFloat;
                      opchf[mes]+=DM->qOgraf->FieldByName("DLIT")->AsFloat;
                    }

                  //общие суммы по часам
                  ochf[mes]+=DM->qOgraf->FieldByName("DLIT")->AsFloat;
                  ovchf[mes]+=v;

                  if (dnism==3)
                    {
                      dnism=1;
                      nsm=0;
                    }
                  else
                    {
                      dnism++;
                    }
                }


            }
          //выходной
          else
            {
              vihod[mes][den]=0;
              chf[mes][den]=0;

              //один выходной при переходе с 12 часов на 8 часов работы
              if (mes==9 && den==17)
                {
                  dnism=2;
                  nsm=1;
                }
              // 2 выходных при 8 часовом графике
              else if (mes==9 && den>17)
                {
                  if (dnism==2)
                    {
                      dnism=1;
                      nsm=1;
                    }
                  else
                    {
                      dnism++;
                    }
                }
              // 3 выходных при 12 часовом графике
              else
                {
                  if (dnism==3)
                    {
                      dnism=1;
                      nsm=1;
                    }
                  else
                    {
                      dnism++;
                    }
                }
            }
        }

      //расчет переработки
      if ((ochf[mes] - DM->qNorma11Graf->FieldByName("chf")->AsFloat - opchf[mes])>0)
        {
          pgraf[mes] = ochf[mes] - DM->qNorma11Graf->FieldByName("chf")->AsFloat - opchf[mes];
        }

     //отсутствующие дни в месяце
      if (den<32)
        {
          while (den<=32)
            {
              vihod[mes][den]="NULL";
              chf[mes][den]="NULL";
              nchf[mes][den]="NULL";
              vchf[mes][den]="NULL";
              pchf[mes][den]="NULL";
              den++;
            }
        }

      DM->qNorma11Graf->Next();
    }
  if (br==1) nsm=1;
  else nsm=0;
  dnism=1;
}
//------------------------------------------------------------------------------

// Расчет 690 графика
void __fastcall TMain::Graf690(double v)
{
  int kol;

  /*nsm - номер смены последнего месяца,
   dnism - день смены последнего месяца*/

  nsm = DM->qObnovlenie2->FieldByName("nsm")->AsInteger;
  dnism = DM->qObnovlenie2->FieldByName("dnism")->AsInteger;

  //по месяцам
  for (mes=1; mes<=12; mes++)
    {
      kol = DaysInAMonth(god, mes);

      //по дням
      for (den=1; den<=kol; den++)
        {
          //рабочий день
          if (nsm==1)
            {
              // праздничный день
              if (PrazdDni(den,mes)==true)
                {
                  vihod[mes][den]=9;
                  chf[mes][den]=30;
                }
              else
                {
                  vihod[mes][den]=1;

                  chf[mes][den]= DM->qOgraf->FieldByName("DLIT")->AsFloat;
                  vchf[mes][den]=v;

                  ochf[mes]+=DM->qOgraf->FieldByName("DLIT")->AsFloat;
                  ovchf[mes]+=v;
                }

                nsm=0;
            /*  if (dnism==6)
                {
                  dnism=1;
                  nsm=0;
                }
              else
                {
                  nsm=0;
                }  */
            }
          //выходной
          else
            {
              // праздничный день
              if (PrazdDni(den,mes)==true)
                {
                  vihod[mes][den]=9;
                  chf[mes][den]=30;
                }
              else
                {
                  vihod[mes][den]=0;
                  chf[mes][den]=0;
                }

              if (dnism==6 || dnism==7)
                {
                  nsm=0;
                  dnism++;
                }
              else if (dnism==8)
                {
                  nsm=1;
                  dnism=1;
                }
              else
                {
                  nsm=1;
                  dnism++;
                }
            }
        }

      //расчет переработки
      if ((ochf[mes] - DM->qNorma11Graf->FieldByName("chf")->AsFloat - opchf[mes])>0)
        {
          pgraf[mes] = ochf[mes] - DM->qNorma11Graf->FieldByName("chf")->AsFloat - opchf[mes];
        }
   
      //отсутствующие дни в месяце
      if (den<32)
        {
          while (den<=32)
            {
              vihod[mes][den]="NULL";
              chf[mes][den]="NULL";
              den++;
            }
        }
      DM->qNorma11Graf->Next();
    }
}
//------------------------------------------------------------------------------

//Расчет 771 графика
void __fastcall TMain::Graf771()
{
  int kol, prazdnik=0;

  /* chf[32] - рабочие часы по дням
     chf[den] = 8 - рабочий день
     chf[den] = 7 - предпраздничный день
     vihod[32] - выходы по дням (рабочий, отдых, праздничный)
     vihod[den] = 1 - рабочий день
     vihod[den] = 9 - праздник
  */

  //по месяцам
  for (mes=1; mes<=12; mes++)
    {
      kol = DaysInAMonth(god, mes);

      //по дням
      for (den=1; den<=kol; den++)
        {
          //проверка дня недели
          if (DayWeek(den,mes,god)==1||DayWeek(den,mes,god)==2)
            {

              //праздничный попадающий на субботу или воскресенье
              if (PrazdDniVihodnue(den,mes,god)==true )
                {
                  prazdnik=1;
                }
              
              //проверка праздничного дня
              if (PrazdDni(den,mes)==true)
                {
                  vihod[mes][den]=9;
                  chf[mes][den]=30;
                  prazdnik=1;
                }
              else
                {
                  vihod[mes][den]=0;
                  chf[mes][den]=0;
                }
            }
          else
            {
              //проверка праздничного дня
              if (PrazdDni(den,mes)==true)
                {
                  vihod[mes][den]=9;
                  chf[mes][den]=30;
                }
              //праздничный попадающий на субботу или воскресенье
              else if (PrazdDniVihodnue(den,mes,god)==true)
                {
                  vihod[mes][den]=0;
                  chf[mes][den]=0;
                }
              else if (prazdnik==1)
                {
                  vihod[mes][den]=0;
                  chf[mes][den]=0;
                  prazdnik=0;
                }
              //рабочий день
              else
                {
                  vihod[mes][den]=1;
                  chf[mes][den]=DM->qOgraf->FieldByName("DLIT")->AsFloat;

                  ochf[mes]+=DM->qOgraf->FieldByName("DLIT")->AsFloat;
                }
            }

        }

      //отсутствующие дни в месяце
      if (den<32)
        {
          while (den<=32)
            {
              vihod[mes][den]="NULL";
              chf[mes][den]="NULL";
              den++;
            }
        }
       DM->qNorma11Graf->Next();
    }
}
//---------------------------------------------------------------------------

//Расчет 775 графика
void __fastcall TMain::Graf775(double v, double n)
{
  AnsiString kol;
  int ogod = god-1;
  int omes=11;

  /*nsm - номер смены последнего месяца,
   dnism - день смены последнего месяца*/

   nsm = DM->qObnovlenie2->FieldByName("nsm")->AsInteger;
   dnism = DM->qObnovlenie2->FieldByName("dnism")->AsInteger;

  //по месяцам
  for (mes=1; mes<=12; mes++)
    {
      NextMonth(omes, ogod);
      kol = DaysInAMonth(ogod, omes);

      //по дням с 26 числа прошлого месяца
      for (den=26; den<=kol; den++)
        {
          //проверка дня недели (выходной)
          if (DayWeek(den,omes,ogod)==1||DayWeek(den,omes,ogod)==7)
            {
              //проверка праздничного дня
              if (PrazdDni(den,omes)==true)
                {
                  vihod[mes][den]=9;
                  chf[mes][den]=30;
                }
              else
                {
                  vihod[mes][den]=0;
                  chf[mes][den]=0;
                }
            }
          //рабочий день
          else
            {
              //первая смена (6.30-14.30)
              if (nsm==1)
                {
                  //проверка праздничного дня
                  if (PrazdDni(den,omes)==true)
                    {
                      vihod[mes][den]=9;
                      chf[mes][den]=30;
                    }
                  //праздничный попадающий на субботу или воскресенье
                  else if (PrazdDniVihodnue(den,omes,ogod)==true)
                    {
                      vihod[mes][den]=0;
                      chf[mes][den]=0;
                    }
                  //проверка предпраздничного дня
                  else if (PrdPrazdDni(den,omes)==true)
                    {
                      vihod[mes][den]=1;
                      chf[mes][den]=DM->qOgraf->FieldByName("DLIT")->AsFloat-1;
                      ochf[mes]+= DM->qOgraf->FieldByName("DLIT")->AsFloat-1;
                    }
                  //рабочий день
                  else
                    {
                      vihod[mes][den]=1;
                      chf[mes][den]=DM->qOgraf->FieldByName("DLIT")->AsFloat;
                      ochf[mes]+= DM->qOgraf->FieldByName("DLIT")->AsFloat;
                    }

                  if (dnism==5)
                    {
                      nsm=2;
                      dnism=1;
                    }
                  else
                    {
                      dnism++;
                    }
                }
              //вторая смена (14.30-22.30)
              else if (nsm==2)
                {
                  //проверка праздничного дня
                  if (PrazdDni(den,omes)==true)
                    {
                      vihod[mes][den]=9;
                      chf[mes][den]=30;
                    }
                  //праздничный попадающий на субботу или воскресенье
                  else if (PrazdDniVihodnue(den,omes,ogod)==true)
                    {
                      vihod[mes][den]=0;
                      chf[mes][den]=0;
                    }
                  //проверка предпраздничного дня
                  else if (PrdPrazdDni(den,omes)==true)
                    {
                      vihod[mes][den]=2;
                      chf[mes][den]=DM->qOgraf->FieldByName("DLIT")->AsFloat-1;
                      vchf[mes][den]= v-1.5;

                      ochf[mes]+= DM->qOgraf->FieldByName("DLIT")->AsFloat-1;
                      ovchf[mes]+=v-1.5;
                    }
                  //рабочий день
                  else
                    {
                      vihod[mes][den]=2;
                      chf[mes][den]=DM->qOgraf->FieldByName("DLIT")->AsFloat;
                      vchf[mes][den]=v;
                      nchf[mes][den]=n;

                      ochf[mes]+= DM->qOgraf->FieldByName("DLIT")->AsFloat;
                      ovchf[mes]+=v;
                      onchf[mes]+=n;
                    }

                  if (dnism==5)
                    {
                      nsm=1;
                      dnism=1;
                    }
                  else
                    {
                      dnism++;
                    }
                }
            }
        }
      //отсутствующие дни в месяце
      if (den<32)
        {
          while (den<=32)
            {
              vihod[mes][den]="NULL";
              chf[mes][den]="NULL";
              nchf[mes][den]="NULL";
              vchf[mes][den]="NULL";
              pchf[mes][den]="NULL";
              den++;
            }
        }

      //по дням с 1 числа текущего месяца
      for (den=1; den<=25; den++)
        {
          //проверка дня недели (выходной)
          if (DayWeek(den,mes,god)==1||DayWeek(den,mes,god)==7)
            {
              //проверка праздничного дня
              if (PrazdDni(den,mes)==true)
                {
                  vihod[mes][den]=9;
                  chf[mes][den]=30;
                }
              else
                {
                  vihod[mes][den]=0;
                  chf[mes][den]=0;
                }
            }
          //рабочий день
          else
            {
              //первая смена (6.30-14.30)
              if (nsm==1)
                {
                  //проверка праздничного дня
                  if (PrazdDni(den,mes)==true)
                    {
                      vihod[mes][den]=9;
                      chf[mes][den]=30;
                    }
                  //праздничный попадающий на субботу или воскресенье
                  else if (PrazdDniVihodnue(den,mes,god)==true)
                    {
                      vihod[mes][den]=0;
                      chf[mes][den]=0;
                    }
                  //проверка предпраздничного дня
                  else if (PrdPrazdDni(den,mes)==true)
                    {
                      vihod[mes][den]=1;
                      chf[mes][den]=DM->qOgraf->FieldByName("DLIT")->AsFloat-1;
                      ochf[mes]+= DM->qOgraf->FieldByName("DLIT")->AsFloat-1;
                    }
                  //рабочий день
                  else
                    {
                      vihod[mes][den]=1;
                      chf[mes][den]=DM->qOgraf->FieldByName("DLIT")->AsFloat;
                      ochf[mes]+= DM->qOgraf->FieldByName("DLIT")->AsFloat;
                    }

                  if (dnism==5)
                    {
                      nsm=2;
                      dnism=1;
                    }
                  else
                    {
                      dnism++;
                    }
                }
              //вторая смена (14.30-22.30)
              else if (nsm==2)
                {
                  //проверка праздничного дня
                  if (PrazdDni(den,mes)==true)
                    {
                      vihod[mes][den]=9;
                      chf[mes][den]=30;
                    }
                  //праздничный попадающий на субботу или воскресенье
                  else if (PrazdDniVihodnue(den,mes,god)==true)
                    {
                      vihod[mes][den]=0;
                      chf[mes][den]=0;
                    }
                  //проверка предпраздничного дня
                  else if (PrdPrazdDni(den,mes)==true)
                    {
                      vihod[mes][den]=2;
                      chf[mes][den]=DM->qOgraf->FieldByName("DLIT")->AsFloat-1;
                      vchf[mes][den]= v-1.5;

                      ochf[mes]+= DM->qOgraf->FieldByName("DLIT")->AsFloat-1;
                      ovchf[mes]+=v-1.5;
                    }
                  //рабочий день
                  else
                    {
                      vihod[mes][den]=2;
                      chf[mes][den]=DM->qOgraf->FieldByName("DLIT")->AsFloat;
                      vchf[mes][den]=v;
                      nchf[mes][den]=n;

                      ochf[mes]+= DM->qOgraf->FieldByName("DLIT")->AsFloat;
                      ovchf[mes]+=v;
                      onchf[mes]+=n;
                    }

                  if (dnism==5)
                    {
                      nsm=1;
                      dnism=1;
                    }
                  else
                    {
                      dnism++;
                    }
                }
            }
        }

      DM->qNorma11Graf->Next();
    }
}
//------------------------------------------------------------------------------

//Расчет 780 графика
void __fastcall TMain::Graf780()
{
  int kol;

  /* chf[32] - рабочие часы по дням
     chf[den] = 8 - рабочий день
     chf[den] = 7 - предпраздничный день
     vihod[32] - выходы по дням (рабочий, отдых, праздничный)
     vihod[den] = 1 - рабочий день
     vihod[den] = 9 - праздник
  */

  //по месяцам
  for (mes=1; mes<=12; mes++)
    {
      //kol = DaysInAMonth(god, mes);

      //по дням c 1 по 25 число
      for (den=1; den<=25; den++)
        {
          //проверка дня недели
          if (DayWeek(den,mes,god)==1||DayWeek(den,mes,god)==7)
            {
              //проверка праздничного дня
              if (PrazdDni(den,mes)==true)
                {
                  vihod[mes][den]=9;
                  chf[mes][den]=30;
                }
              else
                {
                  vihod[mes][den]=0;
                  chf[mes][den]=0;
                }
            }
          else
            {
              //проверка праздничного дня
              if (PrazdDni(den,mes)==true)
                {
                  vihod[mes][den]=9;
                  chf[mes][den]=30;
                }
              //праздничный попадающий на субботу или воскресенье
              else if (PrazdDniVihodnue(den,mes,god)==true)
                {
                  vihod[mes][den]=0;
                  chf[mes][den]=0;
                }
              //проверка предпраздничного дня
              else if (PrdPrazdDni(den,mes)==true)
                {
                  vihod[mes][den]=1;
                  chf[mes][den]=DM->qOgraf->FieldByName("DLIT")->AsFloat-1;
                  ochf[mes]+= DM->qOgraf->FieldByName("DLIT")->AsFloat-1;
                }
              //рабочий день
              else
                {
                  vihod[mes][den]=1;
                  chf[mes][den]=DM->qOgraf->FieldByName("DLIT")->AsFloat;
                  ochf[mes]+= DM->qOgraf->FieldByName("DLIT")->AsFloat;
                }
            }
        }
    }

  int ogod = god-1;
  int omes=11;

  //по месяцам
  for (mes=1; mes<=12; mes++)
    {
      NextMonth(omes, ogod);
     // PrevMonth(mes, ogod);
      kol = DaysInAMonth(ogod, omes);

      //по дням с 26 числа прошлого месяца
      for (den=26; den<=kol; den++)
        {
          //проверка дня недели
          if (DayWeek(den,omes,ogod)==1||DayWeek(den,omes,ogod)==7)
            {
              //проверка праздничного дня
              if (PrazdDni(den,mes)==true)
                {
                  vihod[mes][den]=9;
                  chf[mes][den]=30;
                }
              else
                {
                  vihod[mes][den]=0;
                  chf[mes][den]=0;
                }
            }
          else
            {
              //проверка праздничного дня
              if (PrazdDni(den,omes)==true)
                {
                  vihod[mes][den]=9;
                  chf[mes][den]=30;
                }
              //праздничный попадающий на субботу или воскресенье
              else if (PrazdDniVihodnue(den,omes,ogod)==true)
                {
                  vihod[mes][den]=0;
                  chf[mes][den]=0;
                }
              //проверка предпраздничного дня
              else if (PrdPrazdDni(den,omes)==true)
                {
                  vihod[mes][den]=1;
                  chf[mes][den]=DM->qOgraf->FieldByName("DLIT")->AsFloat-1;
                  ochf[mes]+= DM->qOgraf->FieldByName("DLIT")->AsFloat-1;
                }
              //рабочий день
              else
                {
                  vihod[mes][den]=1;
                  chf[mes][den]=DM->qOgraf->FieldByName("DLIT")->AsFloat;
                  ochf[mes]+= DM->qOgraf->FieldByName("DLIT")->AsFloat;
                }
            }

        }

      //отсутствующие дни в месяце
      if (den<32)
        {
          while (den<=32)
            {
              vihod[mes][den]="NULL";
              chf[mes][den]="NULL";
              den++;
            }
        }
    }
}
//---------------------------------------------------------------------------

//Расчет 800 графика
void __fastcall TMain::Graf800(int day1, int day2)
{

  int kol;

  /* chf[32] - рабочие часы по дням
     chf[den] = 8 - рабочий день
     chf[den] = 7 - предпраздничный день
     vihod[32] - выходы по дням (рабочий, отдых, праздничный)
     vihod[den] = 1 - рабочий день
     vihod[den] = 9 - праздник
  */

  //по месяцам
  for (mes=1; mes<=12; mes++)
    {
      kol = DaysInAMonth(god, mes);

      //по дням
      for (den=1; den<=kol; den++)
        {
          //проверка дня недели
          if (DayWeek(den,mes,god)==day1||DayWeek(den,mes,god)==day2)
            {
              //проверка праздничного дня
              if (PrazdDni(den,mes)==true)
                {
                  vihod[mes][den]=9;
                  chf[mes][den]=30;
                }
              else
                {
                  vihod[mes][den]=0;
                  chf[mes][den]=0;
                }
            }
          else
            {
              //проверка праздничного дня
              if (PrazdDni(den,mes)==true)
                {
                  vihod[mes][den]=9;
                  chf[mes][den]=30;
                }
              //проверка предпраздничного дня
              else if (PrdPrazdDni(den,mes)==true)
                {
                  vihod[mes][den]=1;
                  chf[mes][den]=DM->qOgraf->FieldByName("DLIT")->AsFloat-1;
                  ochf[mes]+= DM->qOgraf->FieldByName("DLIT")->AsFloat-1;
                }
              //рабочий день
              else
                {
                  vihod[mes][den]=1;
                  chf[mes][den]=DM->qOgraf->FieldByName("DLIT")->AsFloat;
                  ochf[mes]+= DM->qOgraf->FieldByName("DLIT")->AsFloat;
                }
            }

        }

      //расчет переработки
   /*   if ((ochf[mes] - DM->qNorma11Graf->FieldByName("chf")->AsFloat)>0)
        {
          pgraf[mes] = ochf[mes] - DM->qNorma11Graf->FieldByName("chf")->AsFloat;
        }*/

      //отсутствующие дни в месяце
      if (den<32)
        {
          while (den<=32)
            {
              vihod[mes][den]="NULL";
              chf[mes][den]="NULL";
              den++;
            }
        }
    }
}
//------------------------------------------------------------------------------


// Расчет 850 графика
void __fastcall TMain::Graf850(double v)
{
  int kol;

  /*nsm - номер смены последнего месяца,
   dnism - день смены последнего месяца*/

  nsm = DM->qObnovlenie2->FieldByName("nsm")->AsInteger;
  dnism = DM->qObnovlenie2->FieldByName("dnism")->AsInteger;

  //по месяцам
  for (mes=1; mes<=12; mes++)
    {
      kol = DaysInAMonth(god, mes);

      //по дням
   for (den=1; den<=kol; den++)
        {
          //рабочий день
          if (nsm==1)
            {
              vihod[mes][den]=1;

              //проверка предпраздничного дня 
              if (PrdPrazdDni(den,mes)==true)
                {
                  chf[mes][den]= DM->qOgraf->FieldByName("DLIT")->AsFloat-1;
                  vchf[mes][den]=v-1;

                  ochf[mes]+=DM->qOgraf->FieldByName("DLIT")->AsFloat-1;
                  ovchf[mes]+=v-1;

                  // праздничный день
                  if (PrazdDni(den,mes)==true)
                    {
                      pchf[mes][den]=DM->qOgraf->FieldByName("DLIT")->AsFloat-1;
                      opchf[mes]+=DM->qOgraf->FieldByName("DLIT")->AsFloat-1;
                    }

                }
              //проверка 1 мая
              else if(mes==5 && den==1)
                {
                  chf[mes][den]= DM->qOgraf->FieldByName("DLIT")->AsFloat-1;
                  pchf[mes][den]=DM->qOgraf->FieldByName("DLIT")->AsFloat-1;
                  vchf[mes][den]=v-1;
                  
                  ochf[mes]+=DM->qOgraf->FieldByName("DLIT")->AsFloat-1;
                  opchf[mes]+=DM->qOgraf->FieldByName("DLIT")->AsFloat-1;
                  ovchf[mes]+=v-1;
                }
              else
                {
                  chf[mes][den]= DM->qOgraf->FieldByName("DLIT")->AsFloat;
                  vchf[mes][den]=v;

                  ochf[mes]+=DM->qOgraf->FieldByName("DLIT")->AsFloat;
                  ovchf[mes]+=v;

                  // праздничный день
                  if (PrazdDni(den,mes)==true)
                    {
                      pchf[mes][den]=DM->qOgraf->FieldByName("DLIT")->AsFloat;
                      opchf[mes]+=DM->qOgraf->FieldByName("DLIT")->AsFloat;
                    }
                }

              if (dnism==2)
                {
                  nsm=0;
                  dnism=1;
                }
              else
                {
                  dnism++;
                }
            }
          //выходной
          else
            {
              vihod[mes][den]=0;
              chf[mes][den]=0;

              if (dnism==2)
                {
                  nsm=1;
                  dnism=1;
                }
              else
                {
                  dnism++;
                }
            }
        }

       //расчет переработки
      if ((ochf[mes] - DM->qNorma11Graf->FieldByName("chf")->AsFloat - opchf[mes])>0)
        {
          pgraf[mes] = ochf[mes] - DM->qNorma11Graf->FieldByName("chf")->AsFloat - opchf[mes];
        }


      //отсутствующие дни в месяце
      if (den<32)
        {
          while (den<=32)
            {
              vihod[mes][den]="NULL";
              chf[mes][den]="NULL";
              den++;
            }
        }
      DM->qNorma11Graf->Next();
    }
}
//------------------------------------------------------------------------------

//Расчет 855 графика
void __fastcall TMain::Graf855(double v)
{
  int kol, prazdnik=0;

  nsm = DM->qObnovlenie2->FieldByName("nsm")->AsInteger;
  dnism = DM->qObnovlenie2->FieldByName("dnism")->AsInteger;

  /* chf[32] - рабочие часы по дням
     chf[den] = 8 - рабочий день
     chf[den] = 7 - предпраздничный день
     vihod[32] - выходы по дням (рабочий, отдых, праздничный)
     vihod[den] = 1 - рабочий день
     vihod[den] = 9 - праздник
     prazd - количество часов на которое сокращается предпраздничная смена
     prazdnik - праздник переносящийся с воскресенья
     nsm - номер смены последнего месяца,
     dnism - день смены последнего месяца*/

  //по месяцам
  for (mes=1; mes<=12; mes++)
    {
      kol = DaysInAMonth(god, mes);

      //по дням
      for (den=1; den<=kol; den++)
        {
          //проверка дня недели
          if (DayWeek(den,mes,god)==1)
            {
              //проверка праздничного дня
              if (PrazdDni(den,mes)==true)
                {
                  vihod[mes][den]=9;
                  chf[mes][den]=30;
                  prazdnik=1;
                }
              else
                {
                  vihod[mes][den]=0;
                  chf[mes][den]=0;
                }

              if (nsm==1) nsm=2;
              else nsm=1;
            }
          else
            {
              //проверка праздничного дня
              if (PrazdDni(den,mes)==true)
                {
                  vihod[mes][den]=9;
                  chf[mes][den]=30;
                }
              //праздничный попадающий на воскресенье
              else if ((graf==855 && prazdnik==1) || (graf==880 && PrazdDniVihodnue(den,mes,god)==true))
                {
                  vihod[mes][den]=0;
                  chf[mes][den]=0;
                  prazdnik=0;
                }
              //проверка предпраздничного дня
              else if (PrdPrazdDni(den,mes)==true)
                {
                  if (nsm==1) vihod[mes][den]=1;
                  else
                    {
                      vihod[mes][den]=2;
                      vchf[mes][den]=v-1;
                      ovchf[mes]+=v-1;
                    }

                  chf[mes][den]=DM->qOgraf->FieldByName("DLIT")->AsFloat-1;
                  ochf[mes]+=DM->qOgraf->FieldByName("DLIT")->AsFloat-1;
                }
              //рабочий день
              else
                {
                  if (nsm==1) vihod[mes][den]=1;
                  else
                    {
                      vihod[mes][den]=2;
                      vchf[mes][den]=v;
                      ovchf[mes]+=v;
                    }

                  chf[mes][den]=DM->qOgraf->FieldByName("DLIT")->AsFloat;
                  ochf[mes]+= DM->qOgraf->FieldByName("DLIT")->AsFloat;
                }

              if (nsm==1) nsm=2;
              else nsm=1;
            }

        }
      //отсутствующие дни в месяце
      if (den<32)
        {
          while (den<=32)
            {
              vihod[mes][den]="NULL";
              chf[mes][den]="NULL";
              den++;
            }
        }
    }
}
//---------------------------------------------------------------------------

//Расчет 865 графика
void __fastcall TMain::Graf865(double v)
{
  AnsiString kol, prazdnik=0;

  /*nsm - номер смены последнего месяца,
   dnism - день смены последнего месяца*/

   nsm = DM->qObnovlenie2->FieldByName("nsm")->AsInteger;
   dnism = DM->qObnovlenie2->FieldByName("dnism")->AsInteger;

  //по месяцам
  for (mes=1; mes<=12; mes++)
    {
      kol = DaysInAMonth(god, mes);

      //по дням 
      for (den=1; den<=kol; den++)
        {
          //проверка дня недели (выходной)
          if (DayWeek(den,mes,god)==1)
            {
              //проверка праздничного дня
              if (PrazdDni(den,mes)==true)
                {
                  vihod[mes][den]=9;
                  chf[mes][den]=30;
                  prazdnik=1;
                }
              else
                {
                  vihod[mes][den]=0;
                  chf[mes][den]=0;
                }
              if (nsm==1) nsm=2;
              else nsm=1;
            }
          //рабочий день
          else
            {
              // проверка, если суббота - 5 часов
              if (DayWeek(den,mes,god)==7)
                {
                  //проверка праздничного дня
                  if (PrazdDni(den,mes)==true)
                    {
                      vihod[mes][den]=9;
                      chf[mes][den]=30;
                      if (nsm==1) nsm=2;
                      else nsm=1;

                    }
                  //проверка предпраздничного дня
                  else if (PrdPrazdDni(den,mes)==true)
                    {
                      chf[mes][den]=DM->qOgraf->FieldByName("DLIT")->AsFloat-2;
                      ochf[mes]+= DM->qOgraf->FieldByName("DLIT")->AsFloat-2;
                      if (nsm==1)
                        {
                          nsm=2;
                          vihod[mes][den]=1;
                        }
                      else if (nsm==2)
                        {
                          nsm=1;
                          vihod[mes][den]=2;
                        }
                    }
                  else
                    {
                      chf[mes][den]=DM->qOgraf->FieldByName("DLIT")->AsFloat-2;
                      ochf[mes]+= DM->qOgraf->FieldByName("DLIT")->AsFloat-2;
                      if (nsm==1)
                        {
                          nsm=2;
                          vihod[mes][den]=1;
                        }
                      else if (nsm==2)
                        {
                          nsm=1;
                          vihod[mes][den]=2;
                        }
                    }

                }
              //другой день недели
              else
                {
                  //первая смена (7.00-14.15)
                  if (nsm==1)
                    {
                      //проверка праздничного дня
                      if (PrazdDni(den,mes)==true)
                        {
                          vihod[mes][den]=9;
                          chf[mes][den]=30;
                        }
                      //проверка на перенос праздника с воскресенья на понедельник
                      else if (prazdnik==1 && DayWeek(den,mes,god)==2)
                        {
                          vihod[mes][den]=0;
                          chf[mes][den]=0;
                          prazdnik = 0;
                        }
                      //проверка предпраздничного дня
                      else if (PrdPrazdDni(den,mes)==true)
                        {
                          vihod[mes][den]=1;
                          chf[mes][den]=DM->qOgraf->FieldByName("DLIT")->AsFloat-1;
                          ochf[mes]+= DM->qOgraf->FieldByName("DLIT")->AsFloat-1;
                        }
                      //рабочий день
                      else
                        {
                          vihod[mes][den]=1;
                          chf[mes][den]=DM->qOgraf->FieldByName("DLIT")->AsFloat;
                          ochf[mes]+= DM->qOgraf->FieldByName("DLIT")->AsFloat;
                        }
                      nsm=2;
                    }
                  //вторая смена (14.30-22.30)
                  else if (nsm==2)
                    {
                      //проверка праздничного дня
                      if (PrazdDni(den,mes)==true)
                        {
                          vihod[mes][den]=9;
                          chf[mes][den]=30;
                        }
                      //проверка на перенос праздника с воскресенья на понедельник
                      else if (prazdnik==1 && DayWeek(den,mes,god)==2)
                        {
                          vihod[mes][den]=0;
                          chf[mes][den]=0;
                          prazdnik = 0;
                        }
                      //проверка предпраздничного дня
                     else if (PrdPrazdDni(den,mes)==true)
                       {
                         vihod[mes][den]=2;
                         chf[mes][den]=DM->qOgraf->FieldByName("DLIT")->AsFloat-1;
                         vchf[mes][den]= v-1;

                         ochf[mes]+= DM->qOgraf->FieldByName("DLIT")->AsFloat-1;
                         ovchf[mes]+=v-1;
                       }
                     //рабочий день
                     else
                       {
                         vihod[mes][den]=2;
                         chf[mes][den]=DM->qOgraf->FieldByName("DLIT")->AsFloat;
                         vchf[mes][den]=v;

                         ochf[mes]+= DM->qOgraf->FieldByName("DLIT")->AsFloat;
                         ovchf[mes]+=v;
                       }
                     nsm=1;
                   }
                }
            }
        }
      //отсутствующие дни в месяце
      if (den<32)
        {
          while (den<=32)
            {
              vihod[mes][den]="NULL";
              chf[mes][den]="NULL";
              nchf[mes][den]="NULL";
              vchf[mes][den]="NULL";
              pchf[mes][den]="NULL";
              den++;
            }
        }
      DM->qNorma11Graf->Next();
    }
}
//------------------------------------------------------------------------------

//Расчет 960 графика
void __fastcall TMain::Graf960(double d1, double d2, double d3, double d4, double d5,
                               double v, double n)
{
  int vih1, vih2, kol;   //vih1 - количество выходных, vih2 - выходные для бригад
                         //1 - воскресенье, 2 - понедельник, 3 - вторник, 4 - среда, 5 - четверг, 6 - пятница, 7 - суббота
  int ponedelnik=0;

  //по месяцам
  for (mes=1; mes<=12; mes++)
    {

      if (br==1 || (br==2 && mes>4 && mes<11) ||
                   (br==3 && mes>4 && mes<11) ||
                   (br==4 && mes>4 && mes<11) ||
                   (br==5 && mes>4 && mes<11))
        {
          kol = DaysInAMonth(god, mes);

          //для мая
          if (mes==5)
            {
              if (br==1)
                {
                  nsm=1;
                  dnism=1;
                }
              else if (br==2)
                {
                  nsm=0;
                  dnism=1;
                }
              else if (br==3)
                {
                  nsm=1;
                  dnism=5;
                }
              else if (br==4)
                {
                  nsm=1;
                  dnism=3;
                }
              else if (br==5)
                {
                  nsm=1;
                  dnism=1;
                }
            }
          else if (mes==4)
            {
              nsm=1;
              dnism=1;
            }   


          //по дням
          for (den=1; den<=kol; den++)
            {
//******************************************************************************
              //период с 1 января по 11 апреля или с 4 ноября по 31 декабря
              if ((mes<4 || (mes==4 && den<12)) || (mes>11 || mes==11 && den>3))
                {
                  //проверка дня недели
                  if (DayWeek(den,mes,god)==1||DayWeek(den,mes,god)==7)
                    {
                      //проверка праздничного дня
                      if (PrazdDni(den,mes)==true)
                        {
                          vihod[mes][den]=9;
                          chf[mes][den]=30;
                        //  pchf[mes][den]=d1;

                        //  opchf[mes]+=d1;
                        }
                      else
                        {
                          vihod[mes][den]=0;
                          chf[mes][den]=0;
                        }
                    }
                  else
                    {
                      //проверка праздничного дня
                      if (PrazdDni(den,mes)==true)
                        {
                          vihod[mes][den]=9;
                          chf[mes][den]=30;
                   //       pchf[mes][den]=d1;

                   //       opchf[mes]+=d1;
                        }
                      //праздничный попадающий на субботу или воскресенье
                      else if (PrazdDniVihodnue(den,mes,god)==true)
                        {
                          vihod[mes][den]=0;
                          chf[mes][den]=0;
                        }
                      //проверка предпраздничного дня
                      else if (PrdPrazdDni(den,mes)==true)
                        {
                          vihod[mes][den]=1;
                          chf[mes][den]=d1-1;
                          ochf[mes]+= d1-1;
                        }
                      //рабочий день
                      else
                        {
                          vihod[mes][den]=1;
                          chf[mes][den]=d1;
                          ochf[mes]+=d1;
                        }
                    }
                }
//******************************************************************************
              //12 апреля
              else if (mes==4 && den==12)
                {
                  vihod[mes][den]=0;
                  chf[mes][den]=0;
                }
//******************************************************************************
              //период с 1 мая по 31 августа
              else if ((mes>4 || mes==4 && den>12) && (mes<9))
                {
                  //для бригады 5
                  switch(br)
                    {
                      case 5:  vih1 = 1;
                               //vih2 = ;
                      break;

                      default: vih1 = 2;
                    }

                  //рабочий день
                  if (nsm==1)
                    {
                      vihod[mes][den]=1;
                      vchf[mes][den]=v;
                      nchf[mes][den]=n;

                      ovchf[mes]+=v;
                      onchf[mes]+=n;

                      //проверка дня недели //выходные
                      if (DayWeek(den,mes,god)==1||DayWeek(den,mes,god)==7)
                        {
                          chf[mes][den]=d3;
                          ochf[mes]+=d3;

                          //проверка праздничного дня
                          if (PrazdDni(den,mes)==true)
                            {
                              pchf[mes][den]=d3;
                              opchf[mes]+=d3;
                            }
                        }
                      else
                        {
                          //проверка праздничного дня
                          if (PrazdDni(den,mes)==true)
                            {
                              chf[mes][den]=d3;
                              ochf[mes]+=d3;

                              pchf[mes][den]=d3;
                              opchf[mes]+=d3;
                            }
                          //праздничный попадающий на субботу или воскресенье
                          else if (PrazdDniVihodnue(den,mes,god)==true)
                            {
                              chf[mes][den]=d3;
                              ochf[mes]+=d3;
                            }
                          else
                            {
                              chf[mes][den]=d2;
                              ochf[mes]+=d2;
                            }
                        }

                      if (dnism==6)
                        {
                          nsm=0;
                          dnism=1;
                        }
                      else
                        {
                          dnism++;
                        }

                    }
                  //выходной
                  else
                    {
                      vihod[mes][den]=0;
                      chf[mes][den]=0;

                      if (dnism==vih1)
                        {
                          nsm=1;
                          dnism=1;
                        }
                      else
                        {
                          dnism++;
                        }
                    }
                }

//******************************************************************************
             //период с 1 сентября по 30 сентября
              else if (mes==9)
                {
                  //1 сентября
                  if (den==1)
                    {
                      vihod[mes][den]=1;
                      chf[mes][den]=d3;
                      vchf[mes][den]=v;
                      nchf[mes][den]=n;

                      ochf[mes]+=d3;
                      ovchf[mes]+=v;
                      onchf[mes]+=n;


                      if (dnism==6 && nsm==1)
                        {
                          nsm=0;
                          dnism=1;
                        }
                      else if (dnism==2 && nsm==0)
                        {
                          nsm=1;
                          dnism=1;
                        }
                      else
                        {
                          dnism++;
                        }
                     }
                   //проверка дня недели (понедельник)
                   else if (DayWeek(den,mes,god)==2)
                     {
                       vihod[mes][den]=0;
                       chf[mes][den]=0;
                       ponedelnik =1;
                     }

                  //до 17 сентября
                  else if (den<17)
                    {
                    /*  //проверка дня недели понедельник
                      if (DayWeek(den,mes,god)==2)
                        {
                          
                        }      */

                      //до 1-го понедельника
                      if (ponedelnik==0)
                        {
                          //для бригады 5
                          switch(br)
                            {
                              case 5:  vih1 = 1;
                              break;

                              default: vih1 = 2;
                            }

                          //рабочий день
                          if (nsm==1)
                            {
                              vihod[mes][den]=1;
                              vchf[mes][den]=v;
                              nchf[mes][den]=n;

                              ovchf[mes]+=v;
                              onchf[mes]+=n;

                              //проверка дня недели //выходные
                              if (DayWeek(den,mes,god)==1||DayWeek(den,mes,god)==7)
                                {
                                  chf[mes][den]=d3;
                                  ochf[mes]+=d3;

                                  //проверка праздничного дня
                                  if (PrazdDni(den,mes)==true)
                                    {
                                      pchf[mes][den]=d3;
                                      opchf[mes]+=d3;
                                    }
                                }
                              else
                                {
                                  //проверка праздничного дня
                                  if (PrazdDni(den,mes)==true)
                                    {
                                      chf[mes][den]=d3;
                                      ochf[mes]+=d3;

                                      pchf[mes][den]=d3;
                                      opchf[mes]+=d3;
                                    }
                                  //праздничный попадающий на субботу или воскресенье
                                  else if (PrazdDniVihodnue(den,mes,god)==true)
                                    {
                                      chf[mes][den]=d3;
                                      ochf[mes]+=d3;
                                    }
                                  else
                                    {
                                      chf[mes][den]=d2;
                                      ochf[mes]+=d2;
                                    }
                                }

                              if (dnism==6)
                                {
                                  nsm=0;
                                  dnism=1;
                                }
                              else
                                {
                                  dnism++;
                                }
                            }
                          //выходной
                          else
                            {
                              vihod[mes][den]=0;
                              chf[mes][den]=0;

                              if (dnism==vih1)
                                {
                                  nsm=1;
                                  dnism=1;
                                }
                              else
                                {
                                  dnism++;
                                }
                            }

                        }
                      //после 1-го понедельника
                      else
                        {
                          if (br==1 && DayWeek(den,mes,god)==3)
                            {
                              vihod[mes][den]=0;
                              chf[mes][den]=0;
                            }
                          else if (br==2 && DayWeek(den,mes,god)==4)
                            {
                              vihod[mes][den]=0;
                              chf[mes][den]=0;
                            }
                          else if (br==3 && DayWeek(den,mes,god)==5)
                            {
                              vihod[mes][den]=0;
                              chf[mes][den]=0;
                            }
                          else if (br==4 && DayWeek(den,mes,god)==6)
                            {
                              vihod[mes][den]=0;
                              chf[mes][den]=0;
                            }
                          else
                            {
                               vihod[mes][den]=1;
                               vchf[mes][den]=v;
                               nchf[mes][den]=n;

                               ovchf[mes]+=v;
                               onchf[mes]+=n;

                              //проверка выходного, праздничный попадающий на субботу или воскресенье
                              if (DayWeek(den,mes,god)==7 || DayWeek(den,mes,god)==1 ||
                                  PrazdDniVihodnue(den,mes,god)==true)
                                {
                                  chf[mes][den]=d3;
                                  ochf[mes]+=d3;
                                }
                              //Проверка праздничного дня
                              else if (PrazdDni(den,mes)==true)
                                {
                                  chf[mes][den]=d3;
                                  pchf[mes][den]=d3;

                                  ochf[mes]+=d3;
                                  opchf[mes]+=d3;
                                }
                              //рабочий день
                              else
                                {
                                  chf[mes][den]=d2;
                                  ochf[mes]+=d2;
                                }
                            }
                        }
                    }
                  //после 17 сентября
                  else
                    {
                     //проверка дня недели (понедельник)
                      if (DayWeek(den,mes,god)==2)
                        {
                          vihod[mes][den]=0;
                          chf[mes][den]=0;
                        }
                      else
                        {
                          if (br==1 && DayWeek(den,mes,god)==3)
                            {
                              vihod[mes][den]=0;
                              chf[mes][den]=0;
                            }
                          else if (br==2 && DayWeek(den,mes,god)==4)
                            {
                              vihod[mes][den]=0;
                              chf[mes][den]=0;
                            }
                          else if (br==3 && DayWeek(den,mes,god)==5)
                            {
                              vihod[mes][den]=0;
                              chf[mes][den]=0;
                            }
                          else if (br==4 && DayWeek(den,mes,god)==6)
                            {
                              vihod[mes][den]=0;
                              chf[mes][den]=0;
                            }
                          else
                            {
                               vihod[mes][den]=1;
                               vchf[mes][den]=v;
                               //nchf[mes][den]=n;

                               ovchf[mes]+=v;
                               //onchf[mes]+=n;

                              //проверка выходного, праздничный попадающий на субботу или воскресенье
                              if (DayWeek(den,mes,god)==7 || DayWeek(den,mes,god)==1 ||
                                  PrazdDniVihodnue(den,mes,god)==true)
                                {
                                  chf[mes][den]=d5;
                                  ochf[mes]+=d5;
                                }
                              //Проверка праздничного дня
                              else if (PrazdDni(den,mes)==true)
                                {
                                  chf[mes][den]=d5;
                                  pchf[mes][den]=d5;

                                  ochf[mes]+=d5;
                                  opchf[mes]+=d5;
                                }
                              //рабочий день
                              else
                                {
                                  chf[mes][den]=d4;
                                  ochf[mes]+=d4;
                                }
                            }
                        }
                     }
                }

//******************************************************************************
              //период с 1 октября по 3 ноября
              else if ((mes==10) || (mes==11 && den<3))
                {
                  //выходные
                  if (DayWeek(den,mes,god)==2||DayWeek(den,mes,god)==3)
                    {
                      vihod[mes][den]=0;
                      chf[mes][den]=0;
                    }
                  //рабочий день
                  else
                    {
                      vihod[mes][den]=1;
                      chf[mes][den]=d2;
                      vchf[mes][den]=v;

                      ochf[mes]+=d2;
                      ovchf[mes]+=v;
                    }
                }
              // 3 ноября
              else if (mes==11 && den==3)
                {
                  vihod[mes][den]=0;
                  chf[mes][den]=0;
                }
            }

          //расчет переработки
          if ((ochf[mes] - DM->qNorma11Graf->FieldByName("chf")->AsFloat - opchf[mes])>0)
            {

              pgraf[mes] = ochf[mes] - DM->qNorma11Graf->FieldByName("chf")->AsFloat - opchf[mes];

           //   ShowMessage(FloatToStr(pgraf[mes])+" = "+FloatToStr(ochf[mes])+" - "+ DM->qNorma11Graf->FieldByName("chf")->AsString+ " - "+opchf[mes]);

            }

          //отсутствующие дни в месяце
          if (den<32)
            {
              while (den<=32)
                {
                  vihod[mes][den]="NULL";
                  chf[mes][den]="NULL";
                  den++;
                }
            }

          DM->qNorma11Graf->Next();

        }
      else
        {
          DM->qNorma11Graf->Next();
        }
    }
}
//------------------------------------------------------------------------------

//Расчет 980 графика
void __fastcall TMain::Graf980(double p1, double p2, double v, double n1, double n2)
{
  AnsiString kol;

  /*nsm - номер смены последнего месяца,
   dnism - день смены последнего месяца*/

  nsm = DM->qObnovlenie2->FieldByName("nsm")->AsInteger;
  dnism = DM->qObnovlenie2->FieldByName("dnism")->AsInteger;

  //по месяцам
  for (mes=1; mes<=12; mes++)
    {
      kol = DaysInAMonth(god, mes);

      //по дням
      for (den=1; den<=kol; den++)
        {
          //первая смена (8.00-20.00)
          //*************************
          if (nsm==1)
            {
              vihod[mes][den]=1;
              vchf[mes][den]=v;
              chf[mes][den]=p1+p2;

              //общие суммы по часам
              ovchf[mes]+=v;
              ochf[mes]+= DM->qOgraf->FieldByName("DLIT")->AsFloat;

              //проверка праздничного дня
              if (PrazdDni(den,mes)==true)
                {
                  pchf[mes][den]=p1+p2;
                  opchf[mes]+=p1+p2;
                }

              if (dnism==2)
                {
                  dnism=1;
                  nsm=2;
                }
              else
                {
                  dnism++;
                }
            }
          //вторая смена (20.00-8.00)
          //*************************
          else if (nsm==2)
            {
              vihod[mes][den]=2;

              //переход на летнее время (март)
              if (mes==3 && den==day_mart)
                {
                  vchf[mes][den]=v;
                  if (dnism==1)
                    {
                      chf[mes][den]=p1;
                      nchf[mes][den]=n1;
                    }
                  else
                    {
                      chf[mes][den]=DM->qOgraf->FieldByName("DLIT")->AsFloat;
                      nchf[mes][den]=n1+n2;
                    }

                  //общие суммы по часам
                  ovchf[mes]+=v;
                  onchf[mes]+=(n1+n2)-1;
                  ochf[mes]+= DM->qOgraf->FieldByName("DLIT")->AsFloat-1;

                  //проверка праздничного дня
                  if (PrazdDni(den,mes)==true)
                    {
                      pchf[mes][den]=p1;
                      opchf[mes]+=p1;
                    }
                  //проверка предпраздничного дня
                  else if (PrdPrazdDni(den,mes)==true)
                    {
                      pchf[mes][den]=p2-1;
                      opchf[mes]+=p2-1;
                    }
                }
              //переход на зимнее время (октябрь)
              else if (mes==10 && den==day_oktyabr)
                {
                  vchf[mes][den]=v;
                  if (dnism==1)
                    {
                      chf[mes][den]=p1;
                      nchf[mes][den]=n1;
                    }
                  else
                    {
                      chf[mes][den]=DM->qOgraf->FieldByName("DLIT")->AsFloat;
                      nchf[mes][den]=n1+n2;
                    }

                  //общие суммы по часам
                  ovchf[mes]+=v;
                  onchf[mes]+=(n1+n2)+1;
                  ochf[mes]+= DM->qOgraf->FieldByName("DLIT")->AsFloat+1;

                  //проверка праздничного дня
                  if (PrazdDni(den,mes)==true)
                    {
                      pchf[mes][den]=p1;
                      opchf[mes]+=p1;
                    }
                  //проверка предпраздничного дня
                  else if (PrdPrazdDni(den,mes)==true)
                    {
                      pchf[mes][den]=p2+1;
                      opchf[mes]+=p2+1;
                    }
                }
              else
                {
                  if (mes==3 && den==day_mart2)
                    {
                      if (dnism==1)
                        {
                          chf[mes][den]=p1;
                          nchf[mes][den]=n1;
                        }
                      else
                        {
                          chf[mes][den]=DM->qOgraf->FieldByName("DLIT")->AsFloat-1;
                          nchf[mes][den]=(n1+n2)-1;
                        }
                    }
                  else if (mes==10 && den==day_oktyabr2)
                    {
                      if (dnism==1)
                        {
                          chf[mes][den]=p1;
                          nchf[mes][den]=n1;
                        }
                      else
                        {
                          chf[mes][den]=DM->qOgraf->FieldByName("DLIT")->AsFloat+1;
                          nchf[mes][den]=(n1+n2)+1;
                        }
                    }
                  else
                    {
                      if (dnism==1)
                        {
                          chf[mes][den]=p1;
                          nchf[mes][den]=n1;
                        }
                      else
                        {
                          chf[mes][den]=DM->qOgraf->FieldByName("DLIT")->AsFloat;
                          nchf[mes][den]=(n1+n2);
                        }
                    }

                  //если ночная смена попадает на последний день месяца
                  if (den==kol)
                    {
                      vchf[mes][den]=v;

                      //общие суммы по часам
                      onchf[mes]+=n1;
                      ovchf[mes]+=v;
                      ochf[mes]+=p1;

                      //проверка праздничного и предпраздничного дня (1 мая)
                      if (PrazdDni(den,mes)==true && PrdPrazdDni(den,mes)==true)
                        {
                          pchf[mes][den]=p1;
                          opchf[mes]+=p1;
                        }
                      else
                        {
                          //проверка праздничного дня
                          if (PrazdDni(den,mes)==true)
                            {
                              pchf[mes][den]=p1;
                              opchf[mes]+=p1;
                            }
                          //проверка предпраздничного дня
                          else if (PrdPrazdDni(den,mes)==true)
                            {
                              pchf[mes][den]=p2;
                            }
                       }
                    }
                  else
                    {
                      vchf[mes][den]=v;

                      //общие суммы по часам
                      onchf[mes]+=(n1+n2);
                      ovchf[mes]+=v;
                      ochf[mes]+=DM->qOgraf->FieldByName("DLIT")->AsFloat;

                      //проверка праздничного и предпраздничного дня
                      if (PrazdDni(den,mes)==true && PrdPrazdDni(den,mes)==true)
                        {
                          pchf[mes][den]=p1+p2;
                          opchf[mes]+=p1+p2;
                        }
                      else
                        {
                          //проверка праздничного дня
                          if (PrazdDni(den,mes)==true)
                            {
                              pchf[mes][den]=p1;
                              opchf[mes]+=p1;
                            }
                          //проверка предпраздничного дня
                          else if (PrdPrazdDni(den,mes)==true)
                            {
                              pchf[mes][den]=p2;
                              opchf[mes]+=p2;
                            }
                        }
                    }
                }

              //часы переходящие c предыдущего месяца
              if (den==1 && dnism==2 && mes==1)
                {
                  ochf[mes]+= Vvod->SetN(DM->qObnovlenie2->FieldByName("chf0")->AsString);
                  onchf[mes]+= Vvod->SetN(DM->qObnovlenie2->FieldByName("nch0")->AsString);
                  opchf[mes]+= Vvod->SetN(DM->qObnovlenie2->FieldByName("pch0")->AsString);
                }
              else if (den==1 && dnism==2)
                {
                  chf0[mes-1]=p2;
                  ochf[mes]+=p2;
                  nchf0[mes-1]=n2;
                  onchf[mes]+=n2;

                  //проверка праздничного дня
                  if (PrazdDni(den,mes)==true)
                    {
                      pchf0[mes-1]=p2;
                      opchf[mes]+=p2;
                    }
                }



              if (dnism==2)
                {
                  dnism=1;
                  nsm=0;
                }
              else
                {
                  dnism++;
                }
            }

          //выходной
          //************************
          else
            {
              if (dnism==1)
                {
                  if (mes==3 && den==day_mart2)
                    {
                      chf[mes][den]=p2-1;
                      nchf[mes][den]=n2-1;
                    }
                  else if (mes==10 && den==day_oktyabr2)
                    {
                      chf[mes][den]=p2+1;
                      nchf[mes][den]=n2+1;
                    }
                  else
                    {
                      chf[mes][den]=p2;
                      nchf[mes][den]=n2;
                    }
                }
              else
                {
                  chf[mes][den]=0;
                  nchf[mes][den]=0;
                }

              vihod[mes][den]=0;

              //часы переходящие c предыдущего месяца
              if (den==1 && dnism==1 && mes==1)
                {
                  ochf[mes]+= Vvod->SetN(DM->qObnovlenie2->FieldByName("chf0")->AsString);
                  onchf[mes]+= Vvod->SetN(DM->qObnovlenie2->FieldByName("nch0")->AsString);
                  opchf[mes]+= Vvod->SetN(DM->qObnovlenie2->FieldByName("pch0")->AsString);
                }
              else if (den==1 && dnism==1)
                {
                  chf0[mes-1]=p2;
                  ochf[mes]+=p2;
                  nchf0[mes-1]=n2;
                  onchf[mes]+=n2;

                  //проверка праздничного дня
                  if (PrazdDni(den,mes)==true)
                    {
                      pchf0[mes-1]=p2;
                      opchf[mes]+=p2;
                    }
                }

              //проверка дня в смене
              if (dnism==4)
                {
                  nsm=1;
                  dnism=1;
                }
              else
                {
                  dnism++;
                }
            }
        }

      //расчет переработки
      if ((ochf[mes] - DM->qNorma11Graf->FieldByName("chf")->AsFloat - opchf[mes])>0)
        {
          pgraf[mes] = ochf[mes] - DM->qNorma11Graf->FieldByName("chf")->AsFloat - opchf[mes];
        }

      // сохранение переходящих часов последнего дня в году
      if ((mes==12 && dnism==1 && nsm==0) || (mes==12 && dnism==2 && nsm==2))
        {
          chf0[mes]=p2;
          nchf0[mes]=6;
          pchf0[mes]=p2;
        }

      //отсутствующие дни в месяце
      if (den<32)
        {
          while (den<=32)
            {
              vihod[mes][den]="NULL";
              chf[mes][den]="NULL";
              nchf[mes][den]="NULL";
              vchf[mes][den]="NULL";
              pchf[mes][den]="NULL";
              den++;
            }
        }

      DM->qNorma11Graf->Next();
    }
}
//------------------------------------------------------------------------------

//Прорисовка Gridа
void __fastcall TMain::DBGridEh1DrawColumnCell(TObject *Sender,
      const TRect &Rect, int DataCol, TColumnEh *Column,
      TGridDrawState State)
{
  TDBGridEh *Grid = (TDBGridEh *) Sender;
  TRect rect, rect1;


  // Разукрашивание грида, в зависимости от месяца
  if (DM->qGrafik->FieldByName("mes")->AsInteger%2==0)
    {
      ((TDBGridEh *) Sender)->Canvas->Brush->Color = (TColor)0x00F8D6E4;
    }
  else
    {
      ((TDBGridEh *) Sender)->Canvas->Brush->Color = (TColor)0x00DAFCE4;
    }


  // выделение серым цветом активной записи
  if (State.Contains(gdSelected) )
    {
      ((TDBGridEh *) Sender)->Canvas->Brush->Color = clSkyBlue;
      ((TDBGridEh *) Sender)->Canvas->Font->Color= clBlack;
    }

  //Вывод месяца в зависимости от бригад
  switch (DM->qOgraf->FieldByName("br")->AsInteger)
    {
      case 1:
        if (!DataCol)
          {
            rect = TRect(Rect.Left, Rect.Top, Rect.Right, Rect.Bottom);

            Grid->Canvas->Brush->Color = clGradientActiveCaption;
            Grid->Canvas->FillRect(rect);

            int X = (rect.Right - rect.Left - Grid->Canvas->TextWidth(DM->qGrafik->FieldByName("mes1")->AsString))/2,
                Y = (rect.Bottom - rect.Top - Grid->Canvas->TextHeight(DM->qGrafik->FieldByName("mes1")->AsString))/2 - 1;

            Grid->Canvas->Font->Color = clBlack;
            Grid->Canvas->Font->Style=TFontStyles()<<fsBold;
            Grid->Canvas->TextOutA(rect.Left + X, rect.Top + Y, DM->qGrafik->FieldByName("mes1")->AsString);

            return;
          }
      break;
      case 2:
        if (!DataCol)
          {
            switch (div(Grid->DataSource->DataSet->RecNo,2).rem)
              {
                case 1:
                  rect = TRect(Rect.Left, Rect.Top, Rect.Right, Rect.Bottom+(Rect.Bottom-Rect.Top+1));
                break;
                case 0:
                  rect = TRect(Rect.Left, Rect.Top-(Rect.Bottom-Rect.Top+1), Rect.Right, Rect.Bottom);
                break;
              }

             Grid->Canvas->MoveTo(Grid->Columns->Items[0]->Width+12, Rect.Bottom);
            Grid->Canvas->LineTo(Grid->Columns->Items[0]->Width+12, Rect.Bottom);

            Grid->Canvas->Brush->Color = clGradientActiveCaption;
            Grid->Canvas->FillRect(rect);

           int X = (rect.Right - rect.Left - Grid->Canvas->TextWidth(DM->qGrafik->FieldByName("mes1")->AsString))/2,
                Y = (rect.Bottom - rect.Top - Grid->Canvas->TextHeight(DM->qGrafik->FieldByName("mes1")->AsString))/2 - 1;

            Grid->Canvas->Font->Color = clBlack;
            Grid->Canvas->Font->Style=TFontStyles()<<fsBold;
            Grid->Canvas->TextOutA(rect.Left + X, rect.Top + Y, DM->qGrafik->FieldByName("mes1")->AsString);
            return;
         }
      break;
      case 3:
        //Объединение поля месяц для 3 бригад
        if (!DataCol)
          {
            switch (div(Grid->DataSource->DataSet->RecNo,3).rem)
              {
                case 1:
                  rect = TRect(Rect.Left, Rect.Top, Rect.Right, Rect.Bottom+2*(Rect.Bottom-Rect.Top+1));
                break;
                case 2:
                  rect = TRect(Rect.Left, Rect.Top-(Rect.Bottom-Rect.Top+1), Rect.Right, Rect.Bottom+(Rect.Bottom-Rect.Top+1));
                break;
                case 0:
                  rect = TRect(Rect.Left, Rect.Top-2*(Rect.Bottom-Rect.Top+1), Rect.Right, Rect.Bottom);
                break;  
              }

            Grid->Canvas->Brush->Color = clGradientActiveCaption;
            Grid->Canvas->FillRect(rect);

            int X = (rect.Right - rect.Left - Grid->Canvas->TextWidth(DM->qGrafik->FieldByName("mes1")->AsString))/2,
                Y = (rect.Bottom - rect.Top - Grid->Canvas->TextHeight(DM->qGrafik->FieldByName("mes1")->AsString))/2 - 1;

            Grid->Canvas->Font->Color = clBlack;
            Grid->Canvas->Font->Style=TFontStyles()<<fsBold;
            Grid->Canvas->TextOutA(rect.Left + X, rect.Top + Y, DM->qGrafik->FieldByName("mes1")->AsString);
            return;
         }
      break;
      case 4:
        //Объединение поля месяц для 4 бригад
        if (!DataCol)
          {
            switch (div(Grid->DataSource->DataSet->RecNo,4).rem)
              {
                case 1:
                  rect = TRect(Rect.Left, Rect.Top, Rect.Right, Rect.Bottom+3*(Rect.Bottom-Rect.Top+1));
                break;
                case 2:
                  rect = TRect(Rect.Left, Rect.Top-(Rect.Bottom-Rect.Top+1), Rect.Right, Rect.Bottom+2*(Rect.Bottom-Rect.Top+1));
                break;
                case 3:
                  rect = TRect(Rect.Left, Rect.Top-2*(Rect.Bottom-Rect.Top+1), Rect.Right, Rect.Bottom+(Rect.Bottom-Rect.Top+1));
                break;
                case 0:
                  rect = TRect(Rect.Left, Rect.Top-3*(Rect.Bottom-Rect.Top+1), Rect.Right, Rect.Bottom);
                break;
              }

            Grid->Canvas->Brush->Color = clGradientActiveCaption;
            Grid->Canvas->FillRect(rect);

            int X = (rect.Right - rect.Left - Grid->Canvas->TextWidth(DM->qGrafik->FieldByName("mes1")->AsString))/2,
                Y = (rect.Bottom - rect.Top - Grid->Canvas->TextHeight(DM->qGrafik->FieldByName("mes1")->AsString))/2 - 1;

            Grid->Canvas->Font->Color = clBlack;
            Grid->Canvas->Font->Style=TFontStyles()<<fsBold;
            Grid->Canvas->TextOutA(rect.Left + X, rect.Top + Y, DM->qGrafik->FieldByName("mes1")->AsString);
            return;
         }
      break;
      case 5:
        //Объединение поля месяц для 4 бригад
        if (!DataCol)
          {
            switch (div(Grid->DataSource->DataSet->RecNo,5).rem)
              {
                case 1:
                  rect = TRect(Rect.Left, Rect.Top, Rect.Right, Rect.Bottom+4*(Rect.Bottom-Rect.Top+1));
                break;
                case 2:
                  rect = TRect(Rect.Left, Rect.Top-(Rect.Bottom-Rect.Top+1), Rect.Right, Rect.Bottom+3*(Rect.Bottom-Rect.Top+1));
                break;
                case 3:
                  rect = TRect(Rect.Left, Rect.Top-2*(Rect.Bottom-Rect.Top+1), Rect.Right, Rect.Bottom+2*(Rect.Bottom-Rect.Top+1));
                break;
                case 4:
                  rect = TRect(Rect.Left, Rect.Top-3*(Rect.Bottom-Rect.Top+1), Rect.Right, Rect.Bottom+(Rect.Bottom-Rect.Top+1));
                break;
                case 0:
                  rect = TRect(Rect.Left, Rect.Top-4*(Rect.Bottom-Rect.Top+1), Rect.Right, Rect.Bottom);
                break;
              }

            Grid->Canvas->Brush->Color = clGradientActiveCaption;
            Grid->Canvas->FillRect(rect);

            int X = (rect.Right - rect.Left - Grid->Canvas->TextWidth(DM->qGrafik->FieldByName("mes1")->AsString))/2,
                Y = (rect.Bottom - rect.Top - Grid->Canvas->TextHeight(DM->qGrafik->FieldByName("mes1")->AsString))/2 - 1;

            Grid->Canvas->Font->Color = clBlack;
            Grid->Canvas->Font->Style=TFontStyles()<<fsBold;
            Grid->Canvas->TextOutA(rect.Left + X, rect.Top + Y, DM->qGrafik->FieldByName("mes1")->AsString);
            return;
         }
       case 6:
        //Объединение поля месяц для 4 бригад
        if (!DataCol)
          {
            switch (div(Grid->DataSource->DataSet->RecNo,6).rem)
              {
                case 1:
                  rect = TRect(Rect.Left, Rect.Top, Rect.Right, Rect.Bottom+4*(Rect.Bottom-Rect.Top+1));
                break;
                case 2:
                  rect = TRect(Rect.Left, Rect.Top-(Rect.Bottom-Rect.Top+1), Rect.Right, Rect.Bottom+4*(Rect.Bottom-Rect.Top+1));
                break;
                case 3:
                  rect = TRect(Rect.Left, Rect.Top-2*(Rect.Bottom-Rect.Top+1), Rect.Right, Rect.Bottom+3*(Rect.Bottom-Rect.Top+1));
                break;
                case 4:
                  rect = TRect(Rect.Left, Rect.Top-3*(Rect.Bottom-Rect.Top+1), Rect.Right, Rect.Bottom+2*(Rect.Bottom-Rect.Top+1));
                break;
                case 5:
                  rect = TRect(Rect.Left, Rect.Top-4*(Rect.Bottom-Rect.Top+1), Rect.Right, Rect.Bottom+(Rect.Bottom-Rect.Top+1));
                break;
                case 0:
                  rect = TRect(Rect.Left, Rect.Top-5*(Rect.Bottom-Rect.Top+1), Rect.Right, Rect.Bottom);
                break;
              }

            Grid->Canvas->Brush->Color = clGradientActiveCaption;
            Grid->Canvas->FillRect(rect);

            int X = (rect.Right - rect.Left - Grid->Canvas->TextWidth(DM->qGrafik->FieldByName("mes1")->AsString))/2,
                Y = (rect.Bottom - rect.Top - Grid->Canvas->TextHeight(DM->qGrafik->FieldByName("mes1")->AsString))/2 - 1;

            Grid->Canvas->Font->Color = clBlack;
            Grid->Canvas->Font->Style=TFontStyles()<<fsBold;
            Grid->Canvas->TextOutA(rect.Left + X, rect.Top + Y, DM->qGrafik->FieldByName("mes1")->AsString);
            return;
         }
      break;      

  //    default :
  //    Application->MessageBox("Не указано количество бригад графика в таблице SPOGRAF","Предупреждение", MB_OK+MB_ICONWARNING);
  //    Abort();
    }

  Grid->DefaultDrawColumnCell(Rect, DataCol, Column, State);

  //рисование горизонтальных полос на гриде
  if ((div(Grid->DataSource->DataSet->RecNo,DM->qOgraf->FieldByName("br")->AsInteger).rem)==0)
    {
      Grid->Canvas->Pen->Width=1;
      Grid->Canvas->Pen->Color=clBlack;
      Grid->Canvas->MoveTo(Rect.Left-150, Rect.Bottom);
      Grid->Canvas->LineTo(Rect.Right, Rect.Bottom);

      Grid->Canvas->MoveTo(Grid->Columns->Items[0]->Width+12, Rect.Bottom);
      Grid->Canvas->LineTo(Grid->Columns->Items[0]->Width+12, Rect.Bottom);
    }


 /*     Grid->Canvas->MoveTo(Grid->Columns->Items[0]->Width+12, Rect.Bottom);
      Grid->Canvas->LineTo(Grid->Columns->Items[0]->Width+12, Grid->Height);
      Grid->Canvas->MoveTo(Grid->Columns->Items[0]->Width+12+Grid->Columns->Items[1]->Width+33*Grid->Columns->Items[2]->Width, Rect.Bottom);
      Grid->Canvas->LineTo(Grid->Columns->Items[0]->Width+12+Grid->Columns->Items[1]->Width+33*Grid->Columns->Items[2]->Width, Grid->Height);
     */


/* DrawEdge(DBGridEh1->Canvas, FRect, BDR_RAISEDINNER, BF_BOTTOM);
   DrawEdge(Canvas->Handle, Rect, BDR_RAISEDINNER, BF_TOPLEFT);

//Объединение ячеек в одну строку
 TRect FRect;

  if (DM->qGrafik->RecNo == 1)
    {
      FRect.Left=11+1;
    //  FRect.Right=FRect.Left+DBGridEh1->Columns[0]->Width+1+DBGridEh1->Columns[1]->Width;
      FRect.Top=Rect.Top;
      FRect.Bottom=Rect.Bottom;
      DBGridEh1->Canvas->FillRect(FRect);
      DBGridEh1->Canvas->TextOut(16, Rect.Top+2, "Какая-то зверюга");
    }
*/

 }
//---------------------------------------------------------------------------

void __fastcall TMain::DBGridEh1DblClick(TObject *Sender)
{
  if (DM->qGrafik->RecordCount!=0 && redakt!=0)
    {
      numk = StringReplace(DBGridEh1->SelectedField->FieldName,"f","", TReplaceFlags()<<rfReplaceAll <<rfIgnoreCase);

      if (DBGridEh1->SelectedField->FieldName!="MES" &&
          DBGridEh1->SelectedField->FieldName!="GRAF" &&
          DBGridEh1->SelectedField->FieldName!="CHF" &&
          DBGridEh1->SelectedField->FieldName!="NCH" &&
          DBGridEh1->SelectedField->FieldName!="VCH" &&
          DBGridEh1->SelectedField->FieldName!="PCH" &&
          DBGridEh1->SelectedField->FieldName!="NORMA" &&
          DBGridEh1->SelectedField->FieldName!="PGRAF" &&
          DM->qGrafik->FieldByName("chf"+numk)->AsString!="-")
        {

          //Получение имени колонки
          numk =  StringReplace(DBGridEh1->Columns->Items[StrToInt(numk)+1]->Title->Caption, "Числа месяца|","", TReplaceFlags()<<rfReplaceAll <<rfIgnoreCase);
        
          N3RedaktirovatClick(Sender);
        }
    }
}
//---------------------------------------------------------------------------

//Редактирование смены
void __fastcall TMain::N3RedaktirovatClick(TObject *Sender)
{
  if (DM->qGrafik->RecordCount!=0)
    {
      numk = StringReplace(DBGridEh1->SelectedField->FieldName,"f","", TReplaceFlags()<<rfReplaceAll <<rfIgnoreCase);

      if (DBGridEh1->SelectedField->FieldName!="MES" &&
          DBGridEh1->SelectedField->FieldName!="GRAF" &&
          DBGridEh1->SelectedField->FieldName!="CHF" &&
          DBGridEh1->SelectedField->FieldName!="NCH" &&
          DBGridEh1->SelectedField->FieldName!="VCH" &&
          DBGridEh1->SelectedField->FieldName!="PCH" &&
          DBGridEh1->SelectedField->FieldName!="NORMA" &&
          DBGridEh1->SelectedField->FieldName!="PGRAF" &&
          DM->qGrafik->FieldByName("chf"+numk)->AsString!="-")
        {
          

          //Получение имени колонки
          numk =  StringReplace(DBGridEh1->Columns->Items[StrToInt(numk)+1]->Title->Caption, "Числа месяца|","", TReplaceFlags()<<rfReplaceAll <<rfIgnoreCase);
        
          SetInfoEdit();


          //Вывод сообщения, о том что смена переходящая
          if (((ComboBox1->Text==40||ComboBox1->Text==1040||ComboBox1->Text==2040||ComboBox1->Text==90||
                ComboBox1->Text==120||ComboBox1->Text==320||ComboBox1->Text==370||ComboBox1->Text==390||
                ComboBox1->Text==450||ComboBox1->Text==520||ComboBox1->Text==950||ComboBox1->Text==210|| ComboBox1->Text==180) && Vvod->EditNSM->Text==1) ||
              ((ComboBox1->Text==60||ComboBox1->Text==1060||ComboBox1->Text==2060||ComboBox1->Text==3060||
                ComboBox1->Text==70||ComboBox1->Text==1070||ComboBox1->Text==2070||ComboBox1->Text==3070||
                ComboBox1->Text==4060||ComboBox1->Text==50||ComboBox1->Text==170) && Vvod->EditNSM->Text==2))
            {
              Vvod->Label8->Visible=true;
            }
          else
            {
              Vvod->Label8->Visible=false;
            }

          Vvod->ShowModal();
        }
    }    
}
//---------------------------------------------------------------------------

// Заполнение Edit-ов
void __fastcall TMain::SetInfoEdit()
{
  AnsiString Sql;
  //определение выбранного поля

  //Получение имя поля
  numk = StringReplace(DBGridEh1->SelectedField->FieldName,"f","", TReplaceFlags()<<rfReplaceAll <<rfIgnoreCase);

  //Получение имени колонки
  numk =  StringReplace(DBGridEh1->Columns->Items[StrToInt(numk)+1]->Title->Caption, "Числа месяца|","", TReplaceFlags()<<rfReplaceAll <<rfIgnoreCase);

  // определение количества дней в месяце
  int kol = DaysInAMonth(god, DM->qGrafik->FieldByName("mes")->AsInteger);

  //существую ли дни в месяце
  if (kol<numk)
    {
      Abort();

    }

  //Вывод значений в Edit
  Vvod->EditNSM->Text = znsm = DM->qGrafik->FieldByName("nsm"+numk)->AsString;
  Vvod->EditCHF->Text = zchf = DM->qGrafik->FieldByName("chf"+numk)->AsString;
  Vvod->EditPCH->Text = zpch = DM->qGrafik->FieldByName("pch"+numk)->AsString;
  Vvod->EditVCH->Text = zvch = DM->qGrafik->FieldByName("vch"+numk)->AsString;
  Vvod->EditNCH->Text = znch = DM->qGrafik->FieldByName("nch"+numk)->AsString;
  Vvod->EditCHF0->Text = zchf0 = DM->qGrafik->FieldByName("chf0")->AsString;
  Vvod->EditNCH0->Text = znch0 = DM->qGrafik->FieldByName("nch0")->AsString;
  Vvod->EditPCH0->Text = zpch0 = DM->qGrafik->FieldByName("pch0")->AsString;

  Vvod->DostupRedaktEdit();
 /*
  //если редактируется последний день месяца в 60, 70, 390 графике и т.д.
  if (kol==numk && (DM->qGrafik->FieldByName("ograf")->AsInteger==60 ||
                    DM->qGrafik->FieldByName("ograf")->AsInteger==1060 ||
                    DM->qGrafik->FieldByName("ograf")->AsInteger==2060 ||
                    DM->qGrafik->FieldByName("ograf")->AsInteger==3060 ||
                    DM->qGrafik->FieldByName("ograf")->AsInteger==70 ||
                    DM->qGrafik->FieldByName("ograf")->AsInteger==1070 ||
                    DM->qGrafik->FieldByName("ograf")->AsInteger==2070 ||
                    DM->qGrafik->FieldByName("ograf")->AsInteger==3070 ||
                    DM->qGrafik->FieldByName("ograf")->AsInteger==100 ||
                    DM->qGrafik->FieldByName("ograf")->AsInteger==120 ||
                    DM->qGrafik->FieldByName("ograf")->AsInteger==320 ||
                    DM->qGrafik->FieldByName("ograf")->AsInteger==370 ||
                    DM->qGrafik->FieldByName("ograf")->AsInteger==390 ||
                    DM->qGrafik->FieldByName("ograf")->AsInteger==1390||
                    DM->qGrafik->FieldByName("ograf")->AsInteger==450 ||
                    DM->qGrafik->FieldByName("ograf")->AsInteger==520 ||
                    DM->qGrafik->FieldByName("ograf")->AsInteger==525 ||
                    DM->qGrafik->FieldByName("ograf")->AsInteger==950 ||
                    DM->qGrafik->FieldByName("ograf")->AsInteger==980))
   {
     //Вывод значений в Edit
     Vvod->EditCHF0->Text = DM->qGrafik->FieldByName("chf0")->AsString;
     Vvod->EditPCH0->Text = DM->qGrafik->FieldByName("pch0")->AsString;
     Vvod->EditNCH0->Text = DM->qGrafik->FieldByName("nch0")->AsString;

     Vvod->Label7->Visible = true;
     Vvod->EditCHF0->Visible = true;
     Vvod->EditPCH0->Visible = true;
     Vvod->EditNCH0->Visible = true;

     Vvod->Label2->Left = 41;
     Vvod->Label3->Left = 41;
     Vvod->Label4->Left = 41;
     Vvod->Label5->Left = 41;
     Vvod->Label6->Left = 41;

     Vvod->EditNSM->Left = 157;
     Vvod->EditCHF->Left = 157;
     Vvod->EditPCH->Left = 157;
     Vvod->EditVCH->Left = 157;
     Vvod->EditNCH->Left = 157;

   }
  else
   {     */
     Vvod->Label7->Visible = false;
     Vvod->EditCHF0->Visible = false;
     Vvod->EditPCH0->Visible = false;
     Vvod->EditNCH0->Visible = false;

     Vvod->Label2->Left = 86;
     Vvod->Label3->Left = 86;
     Vvod->Label4->Left = 86;
     Vvod->Label5->Left = 86;
     Vvod->Label6->Left = 86;


     Vvod->EditNSM->Left = 202;
     Vvod->EditCHF->Left = 202;
     Vvod->EditPCH->Left = 202;
     Vvod->EditVCH->Left = 202;
     Vvod->EditNCH->Left = 202;
  // }




//s = DBGridEh1->Columns->Items[StrToInt(DBGridEh1->SelectedField->Index)]->Title->Caption;
//Vvod->EditCHF->Text = DBGridEh1->Columns[0][StrToInt(DBGridEh1->SelectedField->FieldName)]->Title->Caption;

}
//---------------------------------------------------------------------------

void __fastcall TMain::DBGridEh1KeyPress(TObject *Sender, char &Key)
{
  // редактирование по Enter
  if (Key == VK_RETURN && DM->qGrafik->RecordCount!=0 && redakt!=0)
    {
     if (DBGridEh1->SelectedField->FieldName!="MES1" &&
          DBGridEh1->SelectedField->FieldName!="GRAF" &&
          DBGridEh1->SelectedField->FieldName!="CHF" &&
          DBGridEh1->SelectedField->FieldName!="NCH" &&
          DBGridEh1->SelectedField->FieldName!="VCH" &&
          DBGridEh1->SelectedField->FieldName!="PCH" &&
          DBGridEh1->SelectedField->FieldName!="NORMA" &&
          DBGridEh1->SelectedField->FieldName!="PGRAF")
        {
          N3RedaktirovatClick(Sender);
        }
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

void __fastcall TMain::DBGrid1DrawColumnCell(TObject *Sender,
      const TRect &Rect, int DataCol, TColumn *Column,
      TGridDrawState State)
{
 if (State.Contains(gdSelected))
   {
     ((TDBGrid *) Sender)->Canvas->Brush->Color = clInfoBk;
     ((TDBGrid *) Sender)->Canvas->Font->Color= clRed;

     ((TDBGrid *) Sender)->Canvas->Pen->Color=clGradientActiveCaption;
     ((TDBGrid *) Sender)->Canvas->Pen->Width=5;
     ((TDBGrid *) Sender)->Canvas->Font->Style<<fsBold;

     ((TDBGrid *) Sender)->DefaultDrawColumnCell(Rect, DataCol, Column, State);
   }
}
//---------------------------------------------------------------------------

//Логи
void __fastcall TMain::InsertLog(AnsiString Msg)
{
  AnsiString Data;
  DateTimeToString(Data, "dd.mm.yyyy hh:nn:ss", Now());
  DM->qObnovlenie->SQL->Clear();
  DM->qObnovlenie->SQL->Add("insert into logs (DT,DOMAIN,USERSZPD, PROG, TEXT) values \
                            (to_date(" + QuotedStr(Data) + ", 'DD.MM.YYYY HH24:MI:SS'),\
                             " + QuotedStr(DomainName) + "," + QuotedStr(UserName) + ", 'Grafiki',\
                             " + QuotedStr(Msg) + ")");
  try
    {
      DM->qObnovlenie->ExecSQL();
    }
  catch(...)
    {
      MessageBox(Handle,"Возникла ошибка при вставке данных в таблицу LOGS","Ошибка",8202);
    }

  DM->qObnovlenie->Close();
}
//---------------------------------------------------------------------------

//Передача данных в УИТ
void __fastcall TMain::N5V_UITClick(TObject *Sender)
{
  AnsiString Sql,Str;

  if (Application->MessageBox(("Вы действительно хотите выполнить передачу данных в УИТ \nпо графикам за "+IntToStr(god)+" год? ").c_str(),"Предупреждение",
                              MB_YESNO+ MB_ICONINFORMATION)==ID_NO)
    {
      Abort();
    }
  // spgraf
  // Проверка на наличие записей в таблице SPGRAF
  Sql= "select god, ograf, mes, fakt, graf from spgraf \
        where god="+IntToStr(god)+" \
        and ograf in (select distinct ograf from spgrafiki where god="+IntToStr(god)+")";

  DM->qObnovlenie->Close();
  DM->qObnovlenie->SQL->Clear();
  DM->qObnovlenie->SQL->Add(Sql);

  try
    {
      DM->qObnovlenie->Open();
    }
  catch(...)
    {
      Application->MessageBox("Ошибка получения данных из таблицы SPGRAF","Ошибка", MB_OK + MB_ICONERROR);
      Abort();
    }

  if(DM->qObnovlenie->RecordCount>0)
    {
      if (Application->MessageBox(("В таблице уже существуют данные за "+IntToStr(god)+" год \nпо передаваемым графикам. Перезаписать? ").c_str(),"Предупреждение",
                                   MB_YESNO+ MB_ICONINFORMATION)==ID_NO)
        {
          Abort();
        }
      // spgraf
      // Удаление имеющихся записей из SPGRAF
      Sql = "delete from spgraf \
             where god="+IntToStr(god)+" \
             and ograf in (select distinct ograf from spgrafiki where god="+IntToStr(god)+")";

      DM->qObnovlenie->Close();
      DM->qObnovlenie->SQL->Clear();
      DM->qObnovlenie->SQL->Add(Sql);

      try
        {
          DM->qObnovlenie->ExecSQL();
        }
      catch(...)
        {
          Application->MessageBox("Невозможно удалить предыдущие записи","Ошибка",
                                   MB_OK + MB_ICONERROR);
          Abort();
        }

      Str = "Повторная вставка данных";
    }
  else
    {
      Str = "Вставка данных";
    }
  // spgraf
  //Вставка данных
  Sql = "insert into spgraf (god, ograf, mes, fakt, graf, dlit)                               \
         (select god, ograf, mes, chf, graf,            \
                 (round(nvl(chf,0)/                                                                                           \
                  decode(nvl(decode(nsm1,9,0,1,1,2,1,3,1,4,1,5,1,nsm1),0)+nvl(decode(nsm2,9,0,1,1,2,1,3,1,4,1,5,1,nsm2),0)+   \
                  nvl(decode(nsm3,9,0,1,1,2,1,3,1,4,1,5,1,nsm3),0)+nvl(decode(nsm4,9,0,1,1,2,1,3,1,4,1,5,1,nsm4),0)+          \
                  nvl(decode(nsm5,9,0,1,1,2,1,3,1,4,1,5,1,nsm5),0)+nvl(decode(nsm6,9,0,1,1,2,1,3,1,4,1,5,1,nsm6),0)+          \
                  nvl(decode(nsm7,9,0,1,1,2,1,3,1,4,1,5,1,nsm7),0)+nvl(decode(nsm8,9,0,1,1,2,1,3,1,4,1,5,1,nsm8),0)+          \
                  nvl(decode(nsm9,9,0,1,1,2,1,3,1,4,1,5,1,nsm9),0)+nvl(decode(nsm10,9,0,1,1,2,1,3,1,4,1,5,1,nsm10),0)+        \
                  nvl(decode(nsm11,9,0,1,1,2,1,3,1,4,1,5,1,nsm11),0)+nvl(decode(nsm12,9,0,1,1,2,1,3,1,4,1,5,1,nsm12),0)+      \
                  nvl(decode(nsm13,9,0,1,1,2,1,3,1,4,1,5,1,nsm13),0)+nvl(decode(nsm14,9,0,1,1,2,1,3,1,4,1,5,1,nsm14),0)+      \
                  nvl(decode(nsm15,9,0,1,1,2,1,3,1,4,1,5,1,nsm15),0)+nvl(decode(nsm16,9,0,1,1,2,1,3,1,4,1,5,1,nsm16),0)+      \
                  nvl(decode(nsm17,9,0,1,1,2,1,3,1,4,1,5,1,nsm17),0)+nvl(decode(nsm18,9,0,1,1,2,1,3,1,4,1,5,1,nsm18),0)+      \
                  nvl(decode(nsm19,9,0,1,1,2,1,3,1,4,1,5,1,nsm19),0)+nvl(decode(nsm20,9,0,1,1,2,1,3,1,4,1,5,1,nsm20),0)+      \
                  nvl(decode(nsm21,9,0,1,1,2,1,3,1,4,1,5,1,nsm21),0)+nvl(decode(nsm22,9,0,1,1,2,1,3,1,4,1,5,1,nsm22),0)+      \
                  nvl(decode(nsm23,9,0,1,1,2,1,3,1,4,1,5,1,nsm23),0)+nvl(decode(nsm24,9,0,1,1,2,1,3,1,4,1,5,1,nsm24),0)+      \
                  nvl(decode(nsm25,9,0,1,1,2,1,3,1,4,1,5,1,nsm25),0)+nvl(decode(nsm26,9,0,1,1,2,1,3,1,4,1,5,1,nsm26),0)+      \
                  nvl(decode(nsm27,9,0,1,1,2,1,3,1,4,1,5,1,nsm27),0)+nvl(decode(nsm28,9,0,1,1,2,1,3,1,4,1,5,1,nsm28),0)+      \
                  nvl(decode(nsm29,9,0,1,1,2,1,3,1,4,1,5,1,nsm29),0)+nvl(decode(nsm30,9,0,1,1,2,1,3,1,4,1,5,1,nsm30),0)+      \
                  nvl(decode(nsm31,9,0,1,1,2,1,3,1,4,1,5,1,nsm31),0),0,1,                                                     \
                  (nvl(decode(nsm1,9,0,1,1,2,1,3,1,4,1,5,1,nsm1),0)+nvl(decode(nsm2,9,0,1,1,2,1,3,1,4,1,5,1,nsm2),0)+         \
                  nvl(decode(nsm3,9,0,1,1,2,1,3,1,4,1,5,1,nsm3),0)+nvl(decode(nsm4,9,0,1,1,2,1,3,1,4,1,5,1,nsm4),0)+          \
                  nvl(decode(nsm5,9,0,1,1,2,1,3,1,4,1,5,1,nsm5),0)+nvl(decode(nsm6,9,0,1,1,2,1,3,1,4,1,5,1,nsm6),0)+          \
                  nvl(decode(nsm7,9,0,1,1,2,1,3,1,4,1,5,1,nsm7),0)+nvl(decode(nsm8,9,0,1,1,2,1,3,1,4,1,5,1,nsm8),0)+          \
                  nvl(decode(nsm9,9,0,1,1,2,1,3,1,4,1,5,1,nsm9),0)+nvl(decode(nsm10,9,0,1,1,2,1,3,1,4,1,5,1,nsm10),0)+        \
                  nvl(decode(nsm11,9,0,1,1,2,1,3,1,4,1,5,1,nsm11),0)+nvl(decode(nsm12,9,0,1,1,2,1,3,1,4,1,5,1,nsm12),0)+      \
                  nvl(decode(nsm13,9,0,1,1,2,1,3,1,4,1,5,1,nsm13),0)+nvl(decode(nsm14,9,0,1,1,2,1,3,1,4,1,5,1,nsm14),0)+      \
                  nvl(decode(nsm15,9,0,1,1,2,1,3,1,4,1,5,1,nsm15),0)+nvl(decode(nsm16,9,0,1,1,2,1,3,1,4,1,5,1,nsm16),0)+      \
                  nvl(decode(nsm17,9,0,1,1,2,1,3,1,4,1,5,1,nsm17),0)+nvl(decode(nsm18,9,0,1,1,2,1,3,1,4,1,5,1,nsm18),0)+      \
                  nvl(decode(nsm19,9,0,1,1,2,1,3,1,4,1,5,1,nsm19),0)+nvl(decode(nsm20,9,0,1,1,2,1,3,1,4,1,5,1,nsm20),0)+      \
                  nvl(decode(nsm21,9,0,1,1,2,1,3,1,4,1,5,1,nsm21),0)+nvl(decode(nsm22,9,0,1,1,2,1,3,1,4,1,5,1,nsm22),0)+      \
                  nvl(decode(nsm23,9,0,1,1,2,1,3,1,4,1,5,1,nsm23),0)+nvl(decode(nsm24,9,0,1,1,2,1,3,1,4,1,5,1,nsm24),0)+      \
                  nvl(decode(nsm25,9,0,1,1,2,1,3,1,4,1,5,1,nsm25),0)+nvl(decode(nsm26,9,0,1,1,2,1,3,1,4,1,5,1,nsm26),0)+      \
                  nvl(decode(nsm27,9,0,1,1,2,1,3,1,4,1,5,1,nsm27),0)+nvl(decode(nsm28,9,0,1,1,2,1,3,1,4,1,5,1,nsm28),0)+      \
                  nvl(decode(nsm29,9,0,1,1,2,1,3,1,4,1,5,1,nsm29),0)+nvl(decode(nsm30,9,0,1,1,2,1,3,1,4,1,5,1,nsm30),0)+      \
                  nvl(decode(nsm31,9,0,1,1,2,1,3,1,4,1,5,1,nsm31),0))),2)) as dlit                                            \
           from spgrafiki k where god="+IntToStr(god)+")";

  DM->qObnovlenie->Close();
  DM->qObnovlenie->SQL->Clear();
  DM->qObnovlenie->SQL->Add(Sql);
  try
    {
      DM->qObnovlenie->ExecSQL();
    }
  catch(...)
    {
      Application->MessageBox("Возникла ошибка при вставке данных по графикам в таблицу SPGRAF","Ошибка",
                              MB_OK + MB_ICONERROR);

      InsertLog("Возникла ошибка при вставке данных в таблицу SPGRAF за" +IntToStr(god)+" год");
      Abort();
    }

  InsertLog(Str +" по графикам в таблицу SPGRAF за "+IntToStr(god)+" год выполнена успешно");
  Application->MessageBox("Передача данных в УИТ по графикам выполнена успешно =)","Вставка данных",MB_OK+MB_ICONINFORMATION);

}
//---------------------------------------------------------------------------

//Справка
void __fastcall TMain::N1Click(TObject *Sender)
{
  WinExec(("\""+ WordPath+"\"\""+ Path+"\\Инструкция пользователя.doc\"").c_str(),SW_MAXIMIZE);
}
//---------------------------------------------------------------------------


// Печать графика (Word)
void __fastcall TMain::Word1Click(TObject *Sender)
{
  AnsiString Sql, mm, mm1, graf, vihod;
  double norma=0, chf=0, vch=0, nch=0, pch=0, pererab=0;
  int br;

  /*norma - сумма фактически отработанного месяца за год по 11 графику,
    chf - сумма фактически отработанного месяца за год
    mm - месяц предыдущей строки
    mm1 - месяц последующей строки
    graf - имя шаблона отчета для формирования отчета по вібранному графику
    */


  if (ComboBox1->Text.IsEmpty())
    {
      Application->MessageBox("Выберите необходимый график!!!","Предупреждение",
                              MB_OK+MB_ICONINFORMATION);
      ComboBox1->SetFocus();
      Abort();
    }

/*  if ((ComboBox1->Text!=DM->qGrafik->FieldByName("ograf")->AsString)&&(!DM->qGrafik->FieldByName("ograf")->AsString.IsEmpty()))
    {
      Application->MessageBox("График не выбран!!! \nВыберите необходимый номер графика \nв выпадающем списке и нажмите кнопку 'Выбрать'","Предупреждение",
                              MB_OK+MB_ICONINFORMATION);
      ComboBox1->SetFocus();
      Abort();
    }*/

   //Очистка массива
   norma=0;
   for(br=0; br<=6; br++)
     {
       ochf[br]=0;
       ovchf[br]=0;
       onchf[br]=0;
       opchf[br]=0;
       pgraf[br]=0;
     }

  //Выбор шаблона отчета для выбранного графика
  //if (ComboBox1->Text.Length()==4) graf=StrToInt((ComboBox1->Text).SubString(2,3));
  //else graf = ComboBox1->Text;
  graf = ComboBox1->Text;


  //Режим отображения отчета
  if (DM->qGrafik->FieldByName("otchet")->AsInteger==1) vihod = "nsm";
  else if (DM->qOgraf->FieldByName("otchet")->AsInteger==2) vihod = "chf";
  else
    {
      Application->MessageBox("Не указан режим отображения отчета в таблице OGRAF","Предупреждение",
                               MB_OK + MB_ICONWARNING);
      Abort();
    }

  Sql ="select s.ograf,graf, mes,                                                        \
               dlit, otchet, br, name,    \
               decode(mes, 1, 'Январь',2, 'Февраль',3, 'Март',4, 'Апрель',5, 'Май',6, 'Июнь',7, 'Июль',8, 'Август',9, 'Сентябрь',10, 'Октябрь',11, 'Ноябрь',12, 'Декабрь') as mes1, \
               chf,   \
               decode(pgraf,0,to_number(NULL),pgraf) as pgraf, \
               decode(nch,0,to_number(NULL),nch) as nch,       \
               decode(vch,0,to_number(NULL),vch) as vch,       \
               decode(pch,0,to_number(NULL),pch) as pch,      \
               (select distinct chf from spgrafiki k where k.ograf=f.norma and k.god=s.god and k.mes=s.mes) as norma,                                                          \
               decode(nsm1,9,'П','','-',"+vihod+"1) as nsm1,         \
               decode(nsm2,9,'П','','-',"+vihod+"2) as nsm2,         \
               decode(nsm3,9,'П','','-',"+vihod+"3) as nsm3,         \
               decode(nsm4,9,'П','','-',"+vihod+"4) as nsm4,         \
               decode(nsm5,9,'П','','-',"+vihod+"5) as nsm5,         \
               decode(nsm6,9,'П','','-',"+vihod+"6) as nsm6,         \
               decode(nsm7,9,'П','','-',"+vihod+"7) as nsm7,         \
               decode(nsm8,9,'П','','-',"+vihod+"8) as nsm8,         \
               decode(nsm9,9,'П','','-',"+vihod+"9) as nsm9,         \
               decode(nsm10,9,'П','','-',"+vihod+"10) as nsm10,      \
               decode(nsm11,9,'П','','-',"+vihod+"11) as nsm11,      \
               decode(nsm12,9,'П','','-',"+vihod+"12) as nsm12,      \
               decode(nsm13,9,'П','','-',"+vihod+"13) as nsm13,      \
               decode(nsm14,9,'П','','-',"+vihod+"14) as nsm14,      \
               decode(nsm15,9,'П','','-',"+vihod+"15) as nsm15,      \
               decode(nsm16,9,'П','','-',"+vihod+"16) as nsm16,      \
               decode(nsm17,9,'П','','-',"+vihod+"17) as nsm17,      \
               decode(nsm18,9,'П','','-',"+vihod+"18) as nsm18,      \
               decode(nsm19,9,'П','','-',"+vihod+"19) as nsm19,      \
               decode(nsm20,9,'П','','-',"+vihod+"20) as nsm20,      \
               decode(nsm21,9,'П','','-',"+vihod+"21) as nsm21,      \
               decode(nsm22,9,'П','','-',"+vihod+"22) as nsm22,      \
               decode(nsm23,9,'П','','-',"+vihod+"23) as nsm23,      \
               decode(nsm24,9,'П','','-',"+vihod+"24) as nsm24,      \
               decode(nsm25,9,'П','','-',"+vihod+"25) as nsm25,      \
               decode(nsm26,9,'П','','-',"+vihod+"26) as nsm26,      \
               decode(nsm27,9,'П','','-',"+vihod+"27) as nsm27,      \
               decode(nsm28,9,'П','','-',"+vihod+"28) as nsm28,      \
               decode(nsm29,9,'П','','-',"+vihod+"29) as nsm29,      \
               decode(nsm30,9,'П','','-',"+vihod+"30) as nsm30,      \
               decode(nsm31,9,'П','','-',"+vihod+"31) as nsm31       \
        from spgrafiki s left join spograf f on s.ograf=f.ograf      \
        where god="+IntToStr(god)+" and s.ograf="+ComboBox1->Text+"    \
        order by br, mes, graf";

  DM->qObnovlenie->Close();
  DM->qObnovlenie->SQL->Clear();
  DM->qObnovlenie->SQL->Add(Sql);
  try
    {
      DM->qObnovlenie->Open();
    }
  catch(...)
    {
      Application->MessageBox("Ошибка доступа к таблице с графиками (SPGRAFIKI)","Ошибка доступа",
                              MB_OK + MB_ICONERROR);
      Abort();
    }

  if (DM->qObnovlenie->RecordCount>0)
    {
      Main->Cursor = crHourGlass;
      StatusBar1->SimpleText=" Идет формирование отчета...";

      ProgressBar->Position = 0;
      ProgressBar->Visible = true;
      ProgressBar->Max=DM->qObnovlenie->RecordCount;
     /*
      WaitForm->Visible=true;
      WaitForm->Image1->Visible=true;
      WaitForm->FormStyle=fsStayOnTop;
      WaitForm->CGauge1->MinValue=0;
      WaitForm->CGauge1->MaxValue=DM->qObnovlenie->RecordCount;
      WaitForm->CGauge1->Progress=0; */

      //Создание папки, если ее не существует
      ForceDirectories(WorkPath);

      //Формирование файла данных для печати
      if (!rtf_Open((TempPath + "\\"+graf+".txt").c_str()))
        {
          MessageBox(Handle,"Ошибка открытия файла данных\nВозможно не найден ФАЙЛ шаблона для формирования отчета!!",
                            "Ошибка",8192);
        }
      else
        {
          rtf_Out("ograf", ComboBox1->Text, 0);
          rtf_Out("name", DM->qObnovlenie->FieldByName("name")->AsString,0);
          rtf_Out("god", god,0);


      AnsiString d;
          while (!DM->qObnovlenie->Eof)
            {
              //месяц
              rtf_Out("mes", DM->qObnovlenie->FieldByName("mes1")->AsString,1);

              br=0;

              while (!DM->qObnovlenie->Eof && br!=DM->qObnovlenie->FieldByName("br")->AsInteger)
               {   br++ ;
              //дни

              if (br==1) {d="o"; norma+= DM->qObnovlenie->FieldByName("norma")->AsFloat;}
              else if (br==2) d="d";
              else if (br==3) d="t";
              else if (br==4) d="c";
              else if (br==5) d="p";
              else if (br==6) d="h";
              else
                {
                  Application->MessageBox("Не указано количество бригад в таблице SPOGRAF","Ошибка",
                                          MB_OK + MB_ICONERROR);
                  Abort();
                }

              rtf_Out(d+"br", br,1);
              rtf_Out(d+"1", DM->qObnovlenie->FieldByName("nsm1")->AsString,1);
              rtf_Out(d+"2", DM->qObnovlenie->FieldByName("nsm2")->AsString,1);
              rtf_Out(d+"3", DM->qObnovlenie->FieldByName("nsm3")->AsString,1);
              rtf_Out(d+"4", DM->qObnovlenie->FieldByName("nsm4")->AsString,1);
              rtf_Out(d+"5", DM->qObnovlenie->FieldByName("nsm5")->AsString,1);
              rtf_Out(d+"6", DM->qObnovlenie->FieldByName("nsm6")->AsString,1);
              rtf_Out(d+"7", DM->qObnovlenie->FieldByName("nsm7")->AsString,1);
              rtf_Out(d+"8", DM->qObnovlenie->FieldByName("nsm8")->AsString,1);
              rtf_Out(d+"9", DM->qObnovlenie->FieldByName("nsm9")->AsString,1);
              rtf_Out(d+"10", DM->qObnovlenie->FieldByName("nsm10")->AsString,1);
              rtf_Out(d+"11", DM->qObnovlenie->FieldByName("nsm11")->AsString,1);
              rtf_Out(d+"12", DM->qObnovlenie->FieldByName("nsm12")->AsString,1);
              rtf_Out(d+"13", DM->qObnovlenie->FieldByName("nsm13")->AsString,1);
              rtf_Out(d+"14", DM->qObnovlenie->FieldByName("nsm14")->AsString,1);
              rtf_Out(d+"15", DM->qObnovlenie->FieldByName("nsm15")->AsString,1);
              rtf_Out(d+"16", DM->qObnovlenie->FieldByName("nsm16")->AsString,1);
              rtf_Out(d+"17", DM->qObnovlenie->FieldByName("nsm17")->AsString,1);
              rtf_Out(d+"18", DM->qObnovlenie->FieldByName("nsm18")->AsString,1);
              rtf_Out(d+"19", DM->qObnovlenie->FieldByName("nsm19")->AsString,1);
              rtf_Out(d+"20", DM->qObnovlenie->FieldByName("nsm20")->AsString,1);
              rtf_Out(d+"21", DM->qObnovlenie->FieldByName("nsm21")->AsString,1);
              rtf_Out(d+"22", DM->qObnovlenie->FieldByName("nsm22")->AsString,1);
              rtf_Out(d+"23", DM->qObnovlenie->FieldByName("nsm23")->AsString,1);
              rtf_Out(d+"24", DM->qObnovlenie->FieldByName("nsm24")->AsString,1);
              rtf_Out(d+"25", DM->qObnovlenie->FieldByName("nsm25")->AsString,1);
              rtf_Out(d+"26", DM->qObnovlenie->FieldByName("nsm26")->AsString,1);
              rtf_Out(d+"27", DM->qObnovlenie->FieldByName("nsm27")->AsString,1);
              rtf_Out(d+"28", DM->qObnovlenie->FieldByName("nsm28")->AsString,1);
              rtf_Out(d+"29", DM->qObnovlenie->FieldByName("nsm29")->AsString,1);
              rtf_Out(d+"30", DM->qObnovlenie->FieldByName("nsm30")->AsString,1);
              rtf_Out(d+"31", DM->qObnovlenie->FieldByName("nsm31")->AsString,1);

              //сумма факт, норма
              rtf_Out(d+"graf",DM->qObnovlenie->FieldByName("graf")->AsString,1);
              rtf_Out(d+"chf", DM->qObnovlenie->FieldByName("chf")->AsString,1);
              rtf_Out("norma", DM->qObnovlenie->FieldByName("norma")->AsString,1);
              

              //сумма вечерние ночные
              rtf_Out(d+"pch",DM->qObnovlenie->FieldByName("pch")->AsString,1);
              rtf_Out(d+"vch", DM->qObnovlenie->FieldByName("vch")->AsString,1);
              rtf_Out(d+"nch", DM->qObnovlenie->FieldByName("nch")->AsString,1);
              rtf_Out(d+"pgraf", DM->qObnovlenie->FieldByName("pgraf")->AsString,1);

              //для расчета среднего значания сумм по всем бригадам
             // norma+= DM->qObnovlenie->FieldByName("norma")->AsFloat;
              chf+= DM->qObnovlenie->FieldByName("chf")->AsFloat;
              pch+= DM->qObnovlenie->FieldByName("pch")->AsFloat;
              vch+= DM->qObnovlenie->FieldByName("vch")->AsFloat;
              nch+= DM->qObnovlenie->FieldByName("nch")->AsFloat;
              pererab+= DM->qObnovlenie->FieldByName("pgraf")->AsFloat;

              //для расчета сумм отдельно по каждой бригаде
              ochf[br]+= DM->qObnovlenie->FieldByName("chf")->AsFloat;
              ovchf[br]+= DM->qObnovlenie->FieldByName("vch")->AsFloat;
              onchf[br]+= DM->qObnovlenie->FieldByName("nch")->AsFloat;
              opchf[br]+= DM->qObnovlenie->FieldByName("pch")->AsFloat;
              pgraf[br]+= DM->qObnovlenie->FieldByName("pgraf")->AsFloat;


              //для 800 графика
              if (DM->qObnovlenie->FieldByName("mes")->AsInteger==12)
                {
                  rtf_Out(d+"raz_pgraf", FloatToStr(ochf[br] - (norma)),1);
                }


              DM->qObnovlenie->Next();
              ProgressBar->Position++;
              //WaitForm->CGauge1->Progress++;

              }

                if(!rtf_LineFeed())
                {
                  MessageBox(Handle,"Ошибка записи в файл данных","Ошибка",8192);
                  if (!rtf_Close()) MessageBox(Handle,"Ошибка закрытия файла данных","Ошибка",8192);
                  return;
                } 

            }

           //итоговая сумма отдельно по бригадам
           rtf_Out("norma_sum", FloatToStr(norma),2);
           for (br=1; br<=DM->qObnovlenie->FieldByName("br")->AsInteger; br++)
             {
               if (br==1) d="o";
               else if (br==2) d="d";
               else if (br==3) d="t";
               else if (br==4) d="c";
               else if (br==5) d="p";
               else if (br==6) d="h";
               else
                 {
                   Application->MessageBox("Не указано количество бригад в таблице SPOGRAF","Ошибка",
                                            MB_OK + MB_ICONERROR);
                   Abort();
                 }

               rtf_Out(d+"chf_sum", FloatToStr(ochf[br]),2);
               rtf_Out(d+"vch_sum", FloatToStr(ovchf[br]),2);
               rtf_Out(d+"nch_sum", FloatToStr(onchf[br]),2);
               rtf_Out(d+"pch_sum", FloatToStr(opchf[br]),2);
               rtf_Out(d+"pgraf_sum", FloatToStr(pgraf[br]),2);

               //для 800 графика
               rtf_Out(d+"raz_pgraf", FloatToStr(ochf[br] - (norma)),2);
             }

           // среднегодовые часы
           rtf_Out("norma_sred", FloatToStrF(norma/12, ffFixed, 10,1),2);
           rtf_Out("chf_sred", FloatToStrF(chf/(DM->qObnovlenie->FieldByName("br")->AsInteger*12), ffFixed, 10,1),2);
           rtf_Out("vch_sred", FloatToStrF(vch/(DM->qObnovlenie->FieldByName("br")->AsInteger*12), ffFixed, 10,1),2);
           rtf_Out("nch_sred", FloatToStrF(nch/(DM->qObnovlenie->FieldByName("br")->AsInteger*12), ffFixed, 10,1),2);
           rtf_Out("pch_sred", FloatToStrF(pch/(DM->qObnovlenie->FieldByName("br")->AsInteger*12), ffFixed, 10,1),2);
           rtf_Out("pgraf_sred", FloatToStrF(pererab/(DM->qObnovlenie->FieldByName("br")->AsInteger*12), ffFixed, 10,1),2);
           rtf_Out("raz_pgraf_sred", FloatToStrF(pererab/(DM->qObnovlenie->FieldByName("br")->AsInteger*12), ffFixed, 10,1),2);

           //для 800 графика
        //   rtf_Out("pgraf_sred", FloatToStrF(pererab/(DM->qObnovlenie->FieldByName("br")->AsInteger*12), ffFixed, 10,1),2);


          if(!rtf_LineFeed())
            {
              MessageBox(Handle,"Ошибка записи в файл данных","Ошибка",8192);
              if (!rtf_Close()) MessageBox(Handle,"Ошибка закрытия файла данных","Ошибка",8192);
              return;
            }

          ProgressBar->Visible = false;
          StatusBar1->SimpleText = "Отчет сформирован. Выполняется открытие...";
          Main->Cursor = crDefault;


          if(!rtf_Close())
            {
              MessageBox(Handle,"Ошибка закрытия файла данных", "Ошибка", 8192);
              return;
            }

          int istrd;
          try
            {
              rtf_CreateReport(TempPath +"\\"+graf+".txt",
                               Path+"\\RTF\\"+graf+".rtf",
                         WorkPath+"\\График "+graf+".doc",NULL,&istrd);
              DeleteFile(TempPath+"\\"+graf+".txt");

              WinExec(("\""+ WordPath+"\"\""+ WorkPath+"\\График "+graf+".doc\"").c_str(),SW_MAXIMIZE);
            }
          catch(RepoRTF_Error E)
            {
              MessageBox(Handle,("Ошибка формирования отчета:"+ AnsiString(E.Err)+
                                 "\nСтрока файла данных:"+IntToStr(istrd)).c_str(),"Ошибка",8192);
              StatusBar1->SimpleText = "Отчетный период: "+ IntToStr(god);
              Abort();
            }
        }

         StatusBar1->SimpleText = "Отчетный период: "+ IntToStr(god);

    }
  else
    {
      Application->MessageBox("Нет данных по выбранному графику","Предупреждение",
                               MB_OK + MB_ICONINFORMATION);
    }
}
//---------------------------------------------------------------------------

// Печать графика (Excel) включающий часы и выходы
void __fastcall TMain::Excel1Click(TObject *Sender)
{
   AnsiString Sql, vihod;
  int RowCount, ColCount;

  if (ComboBox1->Text.IsEmpty())
    {
      Application->MessageBox("Выберите необходимый график!!!","Предупреждение",
                              MB_OK+MB_ICONINFORMATION);
      ComboBox1->SetFocus();
      Abort();
    }

   //Выбор шаблона отчета для выбранного графика
  /*if (ComboBox1->Text==11 || ComboBox1->Text==81 || ComboBox1->Text==111 || ComboBox1->Text==480 || ComboBox1->Text==650 || ComboBox1->Text==655 || ComboBox1->Text==1655 ||ComboBox1->Text==660 || ComboBox1->Text==771 || ComboBox1->Text==800 || ComboBox1->Text==820 || ComboBox1->Text==830) graf=11;
  else*/ if (ComboBox1->Text==100 || ComboBox1->Text==315 || ComboBox1->Text==400 || ComboBox1->Text==775 || ComboBox1->Text==780) graf=26;
  else graf=NULL;


  Sql ="select s.ograf, f.smena as smena, graf, mes,                                                        \
               dlit, otchet, br, name, \
               decode(mes, 1, 'Январь',2, 'Февраль',3, 'Март',4, 'Апрель',5, 'Май',6, 'Июнь',7, 'Июль',8, 'Август',9, 'Сентябрь',10, 'Октябрь',11, 'Ноябрь',12, 'Декабрь') as mes1, \
               chf,   \
               decode(pgraf,0,to_number(NULL),pgraf) as pgraf, \
               decode(nch,0,to_number(NULL),nch) as nch,       \
               decode(vch,0,to_number(NULL),vch) as vch,       \
               decode(pch,0,to_number(NULL),pch) as pch,      \
               case when (s.ograf=23) then  (select distinct round(chf/2,0) from spgrafiki k where k.ograf=f.norma and k.god=s.god and k.mes=s.mes) \
                    else (select distinct chf from spgrafiki k where k.ograf=f.norma and k.god=s.god and k.mes=s.mes) end as norma,                 \
               decode(nsm1,9,'П','','-',nsm1) as nsm1,         \
               decode(nsm2,9,'П','','-',nsm2) as nsm2,         \
               decode(nsm3,9,'П','','-',nsm3) as nsm3,         \
               decode(nsm4,9,'П','','-',nsm4) as nsm4,         \
               decode(nsm5,9,'П','','-',nsm5) as nsm5,         \
               decode(nsm6,9,'П','','-',nsm6) as nsm6,         \
               decode(nsm7,9,'П','','-',nsm7) as nsm7,         \
               decode(nsm8,9,'П','','-',nsm8) as nsm8,         \
               decode(nsm9,9,'П','','-',nsm9) as nsm9,         \
               decode(nsm10,9,'П','','-',nsm10) as nsm10,      \
               decode(nsm11,9,'П','','-',nsm11) as nsm11,      \
               decode(nsm12,9,'П','','-',nsm12) as nsm12,      \
               decode(nsm13,9,'П','','-',nsm13) as nsm13,      \
               decode(nsm14,9,'П','','-',nsm14) as nsm14,      \
               decode(nsm15,9,'П','','-',nsm15) as nsm15,      \
               decode(nsm16,9,'П','','-',nsm16) as nsm16,      \
               decode(nsm17,9,'П','','-',nsm17) as nsm17,      \
               decode(nsm18,9,'П','','-',nsm18) as nsm18,      \
               decode(nsm19,9,'П','','-',nsm19) as nsm19,      \
               decode(nsm20,9,'П','','-',nsm20) as nsm20,      \
               decode(nsm21,9,'П','','-',nsm21) as nsm21,      \
               decode(nsm22,9,'П','','-',nsm22) as nsm22,      \
               decode(nsm23,9,'П','','-',nsm23) as nsm23,      \
               decode(nsm24,9,'П','','-',nsm24) as nsm24,      \
               decode(nsm25,9,'П','','-',nsm25) as nsm25,      \
               decode(nsm26,9,'П','','-',nsm26) as nsm26,      \
               decode(nsm27,9,'П','','-',nsm27) as nsm27,      \
               decode(nsm28,9,'П','','-',nsm28) as nsm28,      \
               decode(nsm29,9,'П','','-',nsm29) as nsm29,      \
               decode(nsm30,9,'П','','-',nsm30) as nsm30,      \
               decode(nsm31,9,'П','','-',nsm31) as nsm31,       \
               decode(chf1,30,'П','','-',chf1) as chf1,         \
               decode(chf2,30,'П','','-',chf2) as chf2,         \
               decode(chf3,30,'П','','-',chf3) as chf3,         \
               decode(chf4,30,'П','','-',chf4) as chf4,         \
               decode(chf5,30,'П','','-',chf5) as chf5,         \
               decode(chf6,30,'П','','-',chf6) as chf6,         \
               decode(chf7,30,'П','','-',chf7) as chf7,         \
               decode(chf8,30,'П','','-',chf8) as chf8,         \
               decode(chf9,30,'П','','-',chf9) as chf9,         \
               decode(chf10,30,'П','','-',chf10) as chf10,      \
               decode(chf11,30,'П','','-',chf11) as chf11,      \
               decode(chf12,30,'П','','-',chf12) as chf12,      \
               decode(chf13,30,'П','','-',chf13) as chf13,      \
               decode(chf14,30,'П','','-',chf14) as chf14,      \
               decode(chf15,30,'П','','-',chf15) as chf15,      \
               decode(chf16,30,'П','','-',chf16) as chf16,      \
               decode(chf17,30,'П','','-',chf17) as chf17,      \
               decode(chf18,30,'П','','-',chf18) as chf18,      \
               decode(chf19,30,'П','','-',chf19) as chf19,      \
               decode(chf20,30,'П','','-',chf20) as chf20,      \
               decode(chf21,30,'П','','-',chf21) as chf21,      \
               decode(chf22,30,'П','','-',chf22) as chf22,      \
               decode(chf23,30,'П','','-',chf23) as chf23,      \
               decode(chf24,30,'П','','-',chf24) as chf24,      \
               decode(chf25,30,'П','','-',chf25) as chf25,      \
               decode(chf26,30,'П','','-',chf26) as chf26,      \
               decode(chf27,30,'П','','-',chf27) as chf27,      \
               decode(chf28,30,'П','','-',chf28) as chf28,      \
               decode(chf29,30,'П','','-',chf29) as chf29,      \
               decode(chf30,30,'П','','-',chf30) as chf30,      \
               decode(chf31,30,'П','','-',chf31) as chf31       \
        from spgrafiki s left join spograf f on s.ograf=f.ograf                                          \
        where god="+IntToStr(god)+" and s.ograf="+ComboBox1->Text+"    \
        order by br, mes, graf";

  DM->qObnovlenie->Close();
  DM->qObnovlenie->SQL->Clear();
  DM->qObnovlenie->SQL->Add(Sql);
  try
    {
      DM->qObnovlenie->Open();
    }
  catch(...)
    {
      Application->MessageBox("Ошибка доступа к таблице с графиками (SPGRAFIKI)","Ошибка доступа",
                              MB_OK + MB_ICONERROR);
      Abort();
    }

  // количество записей
  int row = DM->qObnovlenie->RecordCount+1;

  if (DM->qObnovlenie->RecordCount>0)
    {
      Main->Cursor = crHourGlass;
      StatusBar1->SimpleText=" Идет формирование отчета...";

      ProgressBar->Position = 0;
      ProgressBar->Visible = true;
      ProgressBar->Max=DM->qObnovlenie->RecordCount;

      // устанавливаем путь к файлу шаблона
      AnsiString sFile = Path+"\\RTF\\Grafik"+graf+".xlt";

      // инициализируем Excel, открываем этот шаблон
      try
        {
          //проверяем, нет ли запущенного Excel
          AppEx=GetActiveOleObject("Excel.Application");
        }
      catch(...)
        {
          try
            {
              AppEx=CreateOleObject("Excel.Application");
            }
          catch (...)
            {
              Application->MessageBox("Невозможно открыть Microsoft Excel!"
              " Возможно это приложение на компьютере не установлено.","Ошибка",MB_OK+MB_ICONERROR);
            }
        }

      try
        {
          AppEx.OlePropertyGet("WorkBooks").OleProcedure("Open",sFile.c_str());  //открываем книгу, указав её имя
          Sh=AppEx.OlePropertyGet("WorkSheets",1);                               //выбираем № активного листа книги
        }
      catch(...)
        {
          Application->MessageBox("Ошибка открытия книги Microsoft Excel!","Ошибка",MB_OK+MB_ICONERROR);
        }


   /* AnsiString f = Excel->DecimalSeparator;
      ExcelApp.UseSystemSeparators = 0;
      SHEET.CELLS[3, 2].NumberFormat := Format("000%s00", [Excel.DecimalSeparator]);

      Макрос для смены:
      With Application
       .DecimalSeparator = ":"
       .ThousandsSeparator = "."
       . = False
      End With  */


      AnsiString razdelitel = AppEx.OlePropertyGet("DecimalSeparator");        //разделитель дробного числа в Excel
      AnsiString f = AppEx.OlePropertyGet("UseSystemSeparators");

      // выводим в шаблон данные
      // сначала заголовок
      toExcel(Sh,"god",(IntToStr(god)+" год").c_str());
      toExcel(Sh,"graf", (DM->qObnovlenie->FieldByName("ograf")->AsString).c_str());
      toExcel(Sh,"name", (DM->qObnovlenie->FieldByName("name")->AsString).c_str());

      // вставляем в шаблон нужное количество строк
      Variant C;
      Sh.OleProcedure("Select");
      C=Sh.OlePropertyGet("Range","br");
      C=Sh.OlePropertyGet("Rows",(int) C.OlePropertyGet("Row")+1);
      //2 строки по каждой бригаде //+ 11 строк с заголовком месяца
      for(int i=1;i<2*row;i++) C.OleProcedure("Insert");

      int i=0, mes, mes1;
      mes = 0;
      mes1 = DM->qObnovlenie->FieldByName("mes")->AsInteger;

      // количество записей
   //   int row = DM->qObnovlenie->RecordCount+1;

      //определение ячейки с которой начинается заполнение
      // int num = AppEx.OlePropertyGet("ActiveCell").OlePropertyGet("Row");
      int num=8;
      int n=7;
      int br=1;

     // AppEx.OlePropertySet("Visible",true);


       //Месяц
              toExcel(Sh,"mes",i, "Месяц");

              //toExcel(Sh,"mes",i, DM->qObnovlenie->FieldByName("mes1")->AsString.c_str());
              Sh.OlePropertyGet("Range",("A"+IntToStr(n+1)+":A"+IntToStr(n+2)).c_str()).OlePropertySet("HorizontalAlignment", xlCenter); //выровнять по гор.
              Sh.OlePropertyGet("Range",("A"+IntToStr(n+1)+":A"+IntToStr(n+2)).c_str()).OlePropertySet("VerticalAlignment", xlCenter); //выровнять по верт.
              mes1 = DM->qObnovlenie->FieldByName("mes")->AsInteger;

              //Вывод шапки
              toExcel(Sh,"br",i, "№ бриг.");
              Sh.OlePropertyGet("Range",("B"+IntToStr(n+1)+":B"+IntToStr(n+2)).c_str()).OlePropertySet("HorizontalAlignment", xlCenter); //выровнять по гор.
              Sh.OlePropertyGet("Range",("B"+IntToStr(n+1)+":B"+IntToStr(n+2)).c_str()).OlePropertySet("VerticalAlignment", xlCenter); //выровнять по верт.
              toExcel(Sh,"d_1",i, "1");
              toExcel(Sh,"d_2",i, "2");
              toExcel(Sh,"d_3",i, "3");
              toExcel(Sh,"d_4",i, "4");
              toExcel(Sh,"d_5",i, "5");
              toExcel(Sh,"d_6",i, "6");
              toExcel(Sh,"d_7",i, "7");
              toExcel(Sh,"d_8",i, "8");
              toExcel(Sh,"d_9",i, "9");
              toExcel(Sh,"d_10",i, "10");
              toExcel(Sh,"d_11",i, "11");
              toExcel(Sh,"d_12",i, "12");
              toExcel(Sh,"d_13",i, "13");
              toExcel(Sh,"d_14",i, "14");
              toExcel(Sh,"d_15",i, "15");
              toExcel(Sh,"d_16",i, "16");
              toExcel(Sh,"d_17",i, "17");
              toExcel(Sh,"d_18",i, "18");
              toExcel(Sh,"d_19",i, "19");
              toExcel(Sh,"d_20",i, "20");
              toExcel(Sh,"d_21",i, "21");
              toExcel(Sh,"d_22",i, "22");
              toExcel(Sh,"d_23",i, "23");
              toExcel(Sh,"d_24",i, "24");
              toExcel(Sh,"d_25",i, "25");
              toExcel(Sh,"d_26",i, "26");
              toExcel(Sh,"d_27",i, "27");
              toExcel(Sh,"d_28",i, "28");
              toExcel(Sh,"d_29",i, "29");
              toExcel(Sh,"d_30",i, "30");
              toExcel(Sh,"d_31",i, "31");
              toExcel(Sh,"sm1",i, "рабочие смены");
              toExcel(Sh,"chf",i, "факт. время");
              toExcel(Sh,"norma",i, "норма");
              toExcel(Sh,"vch",i, "вечерн. часы");
              toExcel(Sh,"nch",i, "ночные часы");
              toExcel(Sh,"pch",i, "праздн. часы");
              toExcel(Sh,"pgraf",i, "перераб. графика");



              //жирный шрифт
              Sh.OlePropertyGet("Range",("A"+IntToStr(n+1)).c_str()).OlePropertyGet("Font").OlePropertySet("Bold",true);
              Sh.OlePropertyGet("Range",("B"+IntToStr(n+1)).c_str()).OlePropertyGet("Font").OlePropertySet("Bold",true);
              Sh.OlePropertyGet("Range",("C"+IntToStr(n+1)).c_str()).OlePropertyGet("Font").OlePropertySet("Bold",true);
              Sh.OlePropertyGet("Range",("D"+IntToStr(n+1)).c_str()).OlePropertyGet("Font").OlePropertySet("Bold",true);
              Sh.OlePropertyGet("Range",("E"+IntToStr(n+1)).c_str()).OlePropertyGet("Font").OlePropertySet("Bold",true);
              Sh.OlePropertyGet("Range",("F"+IntToStr(n+1)).c_str()).OlePropertyGet("Font").OlePropertySet("Bold",true);
              Sh.OlePropertyGet("Range",("G"+IntToStr(n+1)).c_str()).OlePropertyGet("Font").OlePropertySet("Bold",true);
              Sh.OlePropertyGet("Range",("H"+IntToStr(n+1)).c_str()).OlePropertyGet("Font").OlePropertySet("Bold",true);
              Sh.OlePropertyGet("Range",("I"+IntToStr(n+1)).c_str()).OlePropertyGet("Font").OlePropertySet("Bold",true);
              Sh.OlePropertyGet("Range",("J"+IntToStr(n+1)).c_str()).OlePropertyGet("Font").OlePropertySet("Bold",true);
              Sh.OlePropertyGet("Range",("K"+IntToStr(n+1)).c_str()).OlePropertyGet("Font").OlePropertySet("Bold",true);
              Sh.OlePropertyGet("Range",("L"+IntToStr(n+1)).c_str()).OlePropertyGet("Font").OlePropertySet("Bold",true);
              Sh.OlePropertyGet("Range",("M"+IntToStr(n+1)).c_str()).OlePropertyGet("Font").OlePropertySet("Bold",true);
              Sh.OlePropertyGet("Range",("N"+IntToStr(n+1)).c_str()).OlePropertyGet("Font").OlePropertySet("Bold",true);
              Sh.OlePropertyGet("Range",("O"+IntToStr(n+1)).c_str()).OlePropertyGet("Font").OlePropertySet("Bold",true);
              Sh.OlePropertyGet("Range",("P"+IntToStr(n+1)).c_str()).OlePropertyGet("Font").OlePropertySet("Bold",true);
              Sh.OlePropertyGet("Range",("Q"+IntToStr(n+1)).c_str()).OlePropertyGet("Font").OlePropertySet("Bold",true);
              Sh.OlePropertyGet("Range",("R"+IntToStr(n+1)).c_str()).OlePropertyGet("Font").OlePropertySet("Bold",true);
              Sh.OlePropertyGet("Range",("S"+IntToStr(n+1)).c_str()).OlePropertyGet("Font").OlePropertySet("Bold",true);
              Sh.OlePropertyGet("Range",("T"+IntToStr(n+1)).c_str()).OlePropertyGet("Font").OlePropertySet("Bold",true);
              Sh.OlePropertyGet("Range",("U"+IntToStr(n+1)).c_str()).OlePropertyGet("Font").OlePropertySet("Bold",true);
              Sh.OlePropertyGet("Range",("V"+IntToStr(n+1)).c_str()).OlePropertyGet("Font").OlePropertySet("Bold",true);
              Sh.OlePropertyGet("Range",("W"+IntToStr(n+1)).c_str()).OlePropertyGet("Font").OlePropertySet("Bold",true);
              Sh.OlePropertyGet("Range",("X"+IntToStr(n+1)).c_str()).OlePropertyGet("Font").OlePropertySet("Bold",true);
              Sh.OlePropertyGet("Range",("Y"+IntToStr(n+1)).c_str()).OlePropertyGet("Font").OlePropertySet("Bold",true);
              Sh.OlePropertyGet("Range",("Z"+IntToStr(n+1)).c_str()).OlePropertyGet("Font").OlePropertySet("Bold",true);
              Sh.OlePropertyGet("Range",("AA"+IntToStr(n+1)).c_str()).OlePropertyGet("Font").OlePropertySet("Bold",true);
              Sh.OlePropertyGet("Range",("AB"+IntToStr(n+1)).c_str()).OlePropertyGet("Font").OlePropertySet("Bold",true);
              Sh.OlePropertyGet("Range",("AC"+IntToStr(n+1)).c_str()).OlePropertyGet("Font").OlePropertySet("Bold",true);
              Sh.OlePropertyGet("Range",("AD"+IntToStr(n+1)).c_str()).OlePropertyGet("Font").OlePropertySet("Bold",true);
              Sh.OlePropertyGet("Range",("AE"+IntToStr(n+1)).c_str()).OlePropertyGet("Font").OlePropertySet("Bold",true);
              Sh.OlePropertyGet("Range",("AF"+IntToStr(n+1)).c_str()).OlePropertyGet("Font").OlePropertySet("Bold",true);
              Sh.OlePropertyGet("Range",("AG"+IntToStr(n+1)).c_str()).OlePropertyGet("Font").OlePropertySet("Bold",true);
              Sh.OlePropertyGet("Range",("AH"+IntToStr(n+1)).c_str()).OlePropertyGet("Font").OlePropertySet("Bold",true);
              Sh.OlePropertyGet("Range",("AI"+IntToStr(n+1)).c_str()).OlePropertyGet("Font").OlePropertySet("Bold",true);
              Sh.OlePropertyGet("Range",("AJ"+IntToStr(n+1)).c_str()).OlePropertyGet("Font").OlePropertySet("Bold",true);
              Sh.OlePropertyGet("Range",("AK"+IntToStr(n+1)).c_str()).OlePropertyGet("Font").OlePropertySet("Bold",true);
              Sh.OlePropertyGet("Range",("AL"+IntToStr(n+1)).c_str()).OlePropertyGet("Font").OlePropertySet("Bold",true);
              Sh.OlePropertyGet("Range",("AM"+IntToStr(n+1)).c_str()).OlePropertyGet("Font").OlePropertySet("Bold",true);
              Sh.OlePropertyGet("Range",("AN"+IntToStr(n+1)).c_str()).OlePropertyGet("Font").OlePropertySet("Bold",true);
              //Sh.OlePropertyGet("Range", "d_1").OlePropertyGet("Font").OlePropertySet("Bold",true);

              //вертикальный текст в заголовке
              Sh.OlePropertyGet("Range",("B"+IntToStr(n+1)).c_str()).OlePropertySet("Orientation",90);
              Sh.OlePropertyGet("Range",("AH"+IntToStr(n+1)).c_str()).OlePropertySet("Orientation",90);
              Sh.OlePropertyGet("Range",("AI"+IntToStr(n+1)).c_str()).OlePropertySet("Orientation",90);
              Sh.OlePropertyGet("Range",("AJ"+IntToStr(n+1)).c_str()).OlePropertySet("Orientation",90);
              Sh.OlePropertyGet("Range",("AK"+IntToStr(n+1)).c_str()).OlePropertySet("Orientation",90);
              Sh.OlePropertyGet("Range",("AL"+IntToStr(n+1)).c_str()).OlePropertySet("Orientation",90);
              Sh.OlePropertyGet("Range",("AM"+IntToStr(n+1)).c_str()).OlePropertySet("Orientation",90);
              Sh.OlePropertyGet("Range",("AN"+IntToStr(n+1)).c_str()).OlePropertySet("Orientation",90);
              //Sh.OlePropertyGet("Range", "chf").OlePropertySet("Orientation",90);

              //переносить текст
              Sh.OlePropertyGet("Range",("AG"+IntToStr(n+1)).c_str()).OlePropertySet("WrapText",true);
              Sh.OlePropertyGet("Range",("AH"+IntToStr(n+1)).c_str()).OlePropertySet("WrapText",true);
              Sh.OlePropertyGet("Range",("AI"+IntToStr(n+1)).c_str()).OlePropertySet("WrapText",true);
              Sh.OlePropertyGet("Range",("AJ"+IntToStr(n+1)).c_str()).OlePropertySet("WrapText",true);
              Sh.OlePropertyGet("Range",("AK"+IntToStr(n+1)).c_str()).OlePropertySet("WrapText",true);
              Sh.OlePropertyGet("Range",("AL"+IntToStr(n+1)).c_str()).OlePropertySet("WrapText",true);
              Sh.OlePropertyGet("Range",("AM"+IntToStr(n+1)).c_str()).OlePropertySet("WrapText",true);
              Sh.OlePropertyGet("Range",("AN"+IntToStr(n+1)).c_str()).OlePropertySet("WrapText",true);
              //Sh.OlePropertyGet("Range", "chf").OlePropertySet("WrapText",true);

              //задать высоту ячейки
              Sh.OlePropertyGet("Range",("AH"+IntToStr(n+1)).c_str()).OlePropertySet("RowHeight", 48);
              Sh.OlePropertyGet("Range",("AI"+IntToStr(n+1)).c_str()).OlePropertySet("RowHeight", 48);
              Sh.OlePropertyGet("Range",("AJ"+IntToStr(n+1)).c_str()).OlePropertySet("RowHeight", 48);
              Sh.OlePropertyGet("Range",("AK"+IntToStr(n+1)).c_str()).OlePropertySet("RowHeight", 48);
              Sh.OlePropertyGet("Range",("AL"+IntToStr(n+1)).c_str()).OlePropertySet("RowHeight", 48);
              Sh.OlePropertyGet("Range",("AM"+IntToStr(n+1)).c_str()).OlePropertySet("RowHeight", 48);
              Sh.OlePropertyGet("Range",("AN"+IntToStr(n+1)).c_str()).OlePropertySet("RowHeight", 48);

              //выравнивание по горизонтали
              Sh.OlePropertyGet("Range",("AG"+IntToStr(n+1)).c_str()).OlePropertySet("HorizontalAlignment", xlCenter); //выровнять по гор.
              Sh.OlePropertyGet("Range",("AH"+IntToStr(n+1)).c_str()).OlePropertySet("HorizontalAlignment", xlCenter); //выровнять по гор.
              Sh.OlePropertyGet("Range",("AI"+IntToStr(n+1)).c_str()).OlePropertySet("HorizontalAlignment", xlCenter); //выровнять по гор.
              Sh.OlePropertyGet("Range",("AJ"+IntToStr(n+1)).c_str()).OlePropertySet("HorizontalAlignment", xlCenter); //выровнять по гор.
              Sh.OlePropertyGet("Range",("AK"+IntToStr(n+1)).c_str()).OlePropertySet("HorizontalAlignment", xlCenter); //выровнять по гор.
              Sh.OlePropertyGet("Range",("AL"+IntToStr(n+1)).c_str()).OlePropertySet("HorizontalAlignment", xlCenter); //выровнять по гор.
              Sh.OlePropertyGet("Range",("AM"+IntToStr(n+1)).c_str()).OlePropertySet("HorizontalAlignment", xlCenter); //выровнять по гор.
              Sh.OlePropertyGet("Range",("AN"+IntToStr(n+1)).c_str()).OlePropertySet("HorizontalAlignment", xlCenter); //выровнять по гор.

              //выравнивание по вертикали
              Sh.OlePropertyGet("Range",("AG"+IntToStr(n+1)).c_str()).OlePropertySet("VerticalAlignment", xlCenter);
              Sh.OlePropertyGet("Range",("AH"+IntToStr(n+1)).c_str()).OlePropertySet("VerticalAlignment", xlCenter);
              Sh.OlePropertyGet("Range",("AI"+IntToStr(n+1)).c_str()).OlePropertySet("VerticalAlignment", xlCenter);
              Sh.OlePropertyGet("Range",("AJ"+IntToStr(n+1)).c_str()).OlePropertySet("VerticalAlignment", xlCenter);
              Sh.OlePropertyGet("Range",("AK"+IntToStr(n+1)).c_str()).OlePropertySet("VerticalAlignment", xlCenter);
              Sh.OlePropertyGet("Range",("AL"+IntToStr(n+1)).c_str()).OlePropertySet("VerticalAlignment", xlCenter);
              Sh.OlePropertyGet("Range",("AM"+IntToStr(n+1)).c_str()).OlePropertySet("VerticalAlignment", xlCenter);
              Sh.OlePropertyGet("Range",("AN"+IntToStr(n+1)).c_str()).OlePropertySet("VerticalAlignment", xlCenter);

              //жирная сетка заголовка

              //рисуем сетку
              //Sh.OlePropertyGet("Range",("A"+IntToStr(n+1)+":AM"+IntToStr(n+DM->qGrafik->FieldByName("br")->AsInteger-2)).c_str()).OlePropertyGet("Borders").OlePropertySet("LineStyle", xlContinuous);
              //Sh.OlePropertyGet("Range",("A"+IntToStr(n)+":AM"+IntToStr(n+DM->qGrafik->FieldByName("br")->AsInteger-2)).c_str()).OlePropertyGet("Borders").OlePropertySet("Weight",4);

      i++;
      num++;
      n++;

      while(! DM->qObnovlenie->Eof)
        {
        /* if (mes!=mes1)
            {




               i++;
              num++;
              n++;
            }   */





          //рисуем сетку
          Sh.OlePropertyGet("Range",("A"+IntToStr(n)+":AN"+IntToStr(n)).c_str()).OlePropertyGet("Borders").OlePropertySet("LineStyle", 1);


          if (br==1)
            {
              //Вывод месяца
              toExcel(Sh,"mes",i, DM->qObnovlenie->FieldByName("mes1")->AsString);
              Sh.OlePropertyGet("Range",("A"+IntToStr(n+1)+":A"+IntToStr(n+DM->qGrafik->FieldByName("br")->AsInteger*2)).c_str()).OleProcedure("Merge");
              Sh.OlePropertyGet("Range",("A"+IntToStr(n+1)+":A"+IntToStr(n+DM->qGrafik->FieldByName("br")->AsInteger*2)).c_str()).OlePropertyGet("Font").OlePropertySet("Bold",true);
              if (DM->qGrafik->FieldByName("br")->AsInteger>1)
                {
                  Sh.OlePropertyGet("Range",("A"+IntToStr(n+1)+":A"+IntToStr(n+DM->qGrafik->FieldByName("br")->AsInteger*2)).c_str()).OlePropertySet("Orientation",90);
                }
                 //рисуем сетку
              // Sh.OlePropertyGet("Range",("A"+IntToStr(n)+":AN"+IntToStr(n+DM->qGrafik->FieldByName("br")->AsInteger-br)).c_str()).OlePropertyGet("Borders").OlePropertySet("LineStyle", 1);
            }


          //вывод бригады
          toExcel(Sh,"br",i, br);
          Sh.OlePropertyGet("Range",("B"+IntToStr(n+1)+":B"+IntToStr(n+2)).c_str()).OleProcedure("Merge");
          Sh.OlePropertyGet("Range",("B"+IntToStr(n+1)+":B"+IntToStr(n+2)).c_str()).OlePropertySet("HorizontalAlignment", xlCenter); //выровнять по гор.
          Sh.OlePropertyGet("Range",("B"+IntToStr(n+1)+":B"+IntToStr(n+2)).c_str()).OlePropertySet("VerticalAlignment", xlCenter); //выровнять по верт.
          Sh.OlePropertyGet("Range",("B"+IntToStr(n+1)+":B"+IntToStr(n+2)).c_str()).OlePropertyGet("Font").OlePropertySet("Bold",true);

          //вывод смен
          //1
          if (DM->qObnovlenie->FieldByName("nsm1")->AsString=="П" || DM->qObnovlenie->FieldByName("nsm1")->AsString=="-") toExcel(Sh,"d_1",i,DM->qObnovlenie->FieldByName("nsm1")->AsString.c_str());
          else if (DM->qObnovlenie->FieldByName("nsm1")->AsString==0) toExcel(Sh,"d_1",i,"в");
          else toExcel(Sh,"d_1",i,DM->qObnovlenie->FieldByName("nsm1")->AsFloat);
          //2
          if (DM->qObnovlenie->FieldByName("nsm2")->AsString=="П" || DM->qObnovlenie->FieldByName("nsm2")->AsString=="-") toExcel(Sh,"d_2",i, DM->qObnovlenie->FieldByName("nsm2")->AsString.c_str());
          else if (DM->qObnovlenie->FieldByName("nsm2")->AsString==0) toExcel(Sh,"d_2",i,"в");
          else toExcel(Sh,"d_2",i, DM->qObnovlenie->FieldByName("nsm2")->AsFloat);
          //3
          if (DM->qObnovlenie->FieldByName("nsm3")->AsString=="П" || DM->qObnovlenie->FieldByName("nsm3")->AsString=="-") toExcel(Sh,"d_3",i,DM->qObnovlenie->FieldByName("nsm3")->AsString.c_str());
          else if (DM->qObnovlenie->FieldByName("nsm3")->AsString==0) toExcel(Sh,"d_3",i,"в");
          else toExcel(Sh,"d_3",i,DM->qObnovlenie->FieldByName("nsm3")->AsFloat);
          //4
          if (DM->qObnovlenie->FieldByName("nsm4")->AsString=="П" || DM->qObnovlenie->FieldByName("nsm4")->AsString=="-") toExcel(Sh,"d_4",i,DM->qObnovlenie->FieldByName("nsm4")->AsString.c_str());
          else if (DM->qObnovlenie->FieldByName("nsm4")->AsString==0) toExcel(Sh,"d_4",i,"в");
          else toExcel(Sh,"d_4",i,DM->qObnovlenie->FieldByName("nsm4")->AsFloat);
          //5
          if (DM->qObnovlenie->FieldByName("nsm5")->AsString=="П" || DM->qObnovlenie->FieldByName("nsm5")->AsString=="-") toExcel(Sh,"d_5",i,DM->qObnovlenie->FieldByName("nsm5")->AsString.c_str());
          else if (DM->qObnovlenie->FieldByName("nsm5")->AsString==0) toExcel(Sh,"d_5",i,"в");
          else toExcel(Sh,"d_5",i,DM->qObnovlenie->FieldByName("nsm5")->AsFloat);
          //6
          if (DM->qObnovlenie->FieldByName("nsm6")->AsString=="П" || DM->qObnovlenie->FieldByName("nsm6")->AsString=="-") toExcel(Sh,"d_6",i,DM->qObnovlenie->FieldByName("nsm6")->AsString.c_str());
          else if (DM->qObnovlenie->FieldByName("nsm6")->AsString==0) toExcel(Sh,"d_6",i,"в");
          else toExcel(Sh,"d_6",i,DM->qObnovlenie->FieldByName("nsm6")->AsFloat);
          //7
          if (DM->qObnovlenie->FieldByName("nsm7")->AsString=="П" || DM->qObnovlenie->FieldByName("nsm7")->AsString=="-") toExcel(Sh,"d_7",i,DM->qObnovlenie->FieldByName("nsm7")->AsString.c_str());
          else if (DM->qObnovlenie->FieldByName("nsm7")->AsString==0) toExcel(Sh,"d_7",i,"в");
          else toExcel(Sh,"d_7",i,DM->qObnovlenie->FieldByName("nsm7")->AsFloat);
          //8
          if (DM->qObnovlenie->FieldByName("nsm8")->AsString=="П" || DM->qObnovlenie->FieldByName("nsm8")->AsString=="-") toExcel(Sh,"d_8",i,DM->qObnovlenie->FieldByName("nsm8")->AsString.c_str());
          else if (DM->qObnovlenie->FieldByName("nsm8")->AsString==0) toExcel(Sh,"d_8",i,"в");
          else toExcel(Sh,"d_8",i,DM->qObnovlenie->FieldByName("nsm8")->AsFloat);
          //9
          if (DM->qObnovlenie->FieldByName("nsm9")->AsString=="П" || DM->qObnovlenie->FieldByName("nsm9")->AsString=="-") toExcel(Sh,"d_9",i,DM->qObnovlenie->FieldByName("nsm9")->AsString.c_str());
          else if (DM->qObnovlenie->FieldByName("nsm9")->AsString==0) toExcel(Sh,"d_9",i,"в");
          else toExcel(Sh,"d_9",i,DM->qObnovlenie->FieldByName("nsm9")->AsFloat);
          //10
          if (DM->qObnovlenie->FieldByName("nsm10")->AsString=="П" || DM->qObnovlenie->FieldByName("nsm10")->AsString=="-") toExcel(Sh,"d_10",i,DM->qObnovlenie->FieldByName("nsm10")->AsString.c_str());
          else if (DM->qObnovlenie->FieldByName("nsm10")->AsString==0) toExcel(Sh,"d_10",i,"в");
          else toExcel(Sh,"d_10",i,DM->qObnovlenie->FieldByName("nsm10")->AsFloat);
          //11
          if (DM->qObnovlenie->FieldByName("nsm11")->AsString=="П" || DM->qObnovlenie->FieldByName("nsm11")->AsString=="-") toExcel(Sh,"d_11",i,DM->qObnovlenie->FieldByName("nsm11")->AsString.c_str());
          else if (DM->qObnovlenie->FieldByName("nsm11")->AsString==0) toExcel(Sh,"d_11",i,"в");
          else toExcel(Sh,"d_11",i,DM->qObnovlenie->FieldByName("nsm11")->AsFloat);
          //12
          if (DM->qObnovlenie->FieldByName("nsm12")->AsString=="П" || DM->qObnovlenie->FieldByName("nsm12")->AsString=="-") toExcel(Sh,"d_12",i,DM->qObnovlenie->FieldByName("nsm12")->AsString.c_str());
          else if (DM->qObnovlenie->FieldByName("nsm12")->AsString==0) toExcel(Sh,"d_12",i,"в");
          else toExcel(Sh,"d_12",i,DM->qObnovlenie->FieldByName("nsm12")->AsFloat);
          //13
          if (DM->qObnovlenie->FieldByName("nsm13")->AsString=="П" || DM->qObnovlenie->FieldByName("nsm13")->AsString=="-") toExcel(Sh,"d_13",i,DM->qObnovlenie->FieldByName("nsm13")->AsString.c_str());
          else if (DM->qObnovlenie->FieldByName("nsm13")->AsString==0) toExcel(Sh,"d_13",i,"в");
          else toExcel(Sh,"d_13",i,DM->qObnovlenie->FieldByName("nsm13")->AsFloat);
          //14
          if (DM->qObnovlenie->FieldByName("nsm14")->AsString=="П" || DM->qObnovlenie->FieldByName("nsm14")->AsString=="-") toExcel(Sh,"d_14",i,DM->qObnovlenie->FieldByName("nsm14")->AsString.c_str());
          else if (DM->qObnovlenie->FieldByName("nsm14")->AsString==0) toExcel(Sh,"d_14",i,"в");
          else toExcel(Sh,"d_14",i,DM->qObnovlenie->FieldByName("nsm14")->AsFloat);
          //15
          if (DM->qObnovlenie->FieldByName("nsm15")->AsString=="П" || DM->qObnovlenie->FieldByName("nsm15")->AsString=="-") toExcel(Sh,"d_15",i,DM->qObnovlenie->FieldByName("nsm15")->AsString.c_str());
          else if (DM->qObnovlenie->FieldByName("nsm15")->AsString==0) toExcel(Sh,"d_15",i,"в");
          else toExcel(Sh,"d_15",i,DM->qObnovlenie->FieldByName("nsm15")->AsFloat);
          //16
          if (DM->qObnovlenie->FieldByName("nsm16")->AsString=="П" || DM->qObnovlenie->FieldByName("nsm16")->AsString=="-") toExcel(Sh,"d_16",i,DM->qObnovlenie->FieldByName("nsm16")->AsString.c_str());
          else if (DM->qObnovlenie->FieldByName("nsm16")->AsString==0) toExcel(Sh,"d_16",i,"в");
          else toExcel(Sh,"d_16",i,DM->qObnovlenie->FieldByName("nsm16")->AsFloat);
          //17
          if (DM->qObnovlenie->FieldByName("nsm17")->AsString=="П" || DM->qObnovlenie->FieldByName("nsm17")->AsString=="-") toExcel(Sh,"d_17",i,DM->qObnovlenie->FieldByName("nsm17")->AsString.c_str());
          else if (DM->qObnovlenie->FieldByName("nsm17")->AsString==0) toExcel(Sh,"d_17",i,"в");
          else toExcel(Sh,"d_17",i,DM->qObnovlenie->FieldByName("nsm17")->AsFloat);
          //18
          if (DM->qObnovlenie->FieldByName("nsm18")->AsString=="П" || DM->qObnovlenie->FieldByName("nsm18")->AsString=="-") toExcel(Sh,"d_18",i,DM->qObnovlenie->FieldByName("nsm18")->AsString.c_str());
          else if (DM->qObnovlenie->FieldByName("nsm18")->AsString==0) toExcel(Sh,"d_18",i,"в");
          else toExcel(Sh,"d_18",i,DM->qObnovlenie->FieldByName("nsm18")->AsFloat);
          //19
          if (DM->qObnovlenie->FieldByName("nsm19")->AsString=="П" || DM->qObnovlenie->FieldByName("nsm19")->AsString=="-") toExcel(Sh,"d_19",i,DM->qObnovlenie->FieldByName("nsm19")->AsString.c_str());
          else if (DM->qObnovlenie->FieldByName("nsm19")->AsString==0) toExcel(Sh,"d_19",i,"в");
          else toExcel(Sh,"d_19",i,DM->qObnovlenie->FieldByName("nsm19")->AsFloat);
          //20
          if (DM->qObnovlenie->FieldByName("nsm20")->AsString=="П" || DM->qObnovlenie->FieldByName("nsm20")->AsString=="-") toExcel(Sh,"d_20",i,DM->qObnovlenie->FieldByName("nsm20")->AsString.c_str());
          else if (DM->qObnovlenie->FieldByName("nsm20")->AsString==0) toExcel(Sh,"d_20",i,"в");
          else toExcel(Sh,"d_20",i,DM->qObnovlenie->FieldByName("nsm20")->AsFloat);
          //21
          if (DM->qObnovlenie->FieldByName("nsm21")->AsString=="П" || DM->qObnovlenie->FieldByName("nsm21")->AsString=="-") toExcel(Sh,"d_21",i,DM->qObnovlenie->FieldByName("nsm21")->AsString.c_str());
          else if (DM->qObnovlenie->FieldByName("nsm21")->AsString==0) toExcel(Sh,"d_21",i,"в");
          else toExcel(Sh,"d_21",i,DM->qObnovlenie->FieldByName("nsm21")->AsFloat);
          //22
          if (DM->qObnovlenie->FieldByName("nsm22")->AsString=="П" || DM->qObnovlenie->FieldByName("nsm22")->AsString=="-") toExcel(Sh,"d_22",i,DM->qObnovlenie->FieldByName("nsm22")->AsString.c_str());
          else if (DM->qObnovlenie->FieldByName("nsm22")->AsString==0) toExcel(Sh,"d_22",i,"в");
          else toExcel(Sh,"d_22",i,DM->qObnovlenie->FieldByName("nsm22")->AsFloat);
          //23
          if (DM->qObnovlenie->FieldByName("nsm23")->AsString=="П" || DM->qObnovlenie->FieldByName("nsm23")->AsString=="-") toExcel(Sh,"d_23",i,DM->qObnovlenie->FieldByName("nsm23")->AsString.c_str());
          else if (DM->qObnovlenie->FieldByName("nsm23")->AsString==0) toExcel(Sh,"d_23",i,"в");
          else toExcel(Sh,"d_23",i,DM->qObnovlenie->FieldByName("nsm23")->AsFloat);
          //24
          if (DM->qObnovlenie->FieldByName("nsm24")->AsString=="П" || DM->qObnovlenie->FieldByName("nsm24")->AsString=="-") toExcel(Sh,"d_24",i,DM->qObnovlenie->FieldByName("nsm24")->AsString.c_str());
          else if (DM->qObnovlenie->FieldByName("nsm24")->AsString==0) toExcel(Sh,"d_24",i,"в");
          else toExcel(Sh,"d_24",i,DM->qObnovlenie->FieldByName("nsm24")->AsFloat);
          //25
          if (DM->qObnovlenie->FieldByName("nsm25")->AsString=="П" || DM->qObnovlenie->FieldByName("nsm25")->AsString=="-") toExcel(Sh,"d_25",i,DM->qObnovlenie->FieldByName("nsm25")->AsString.c_str());
          else if (DM->qObnovlenie->FieldByName("nsm25")->AsString==0) toExcel(Sh,"d_25",i,"в");
          else toExcel(Sh,"d_25",i,DM->qObnovlenie->FieldByName("nsm25")->AsFloat);
          //26
          if (DM->qObnovlenie->FieldByName("nsm26")->AsString=="П" || DM->qObnovlenie->FieldByName("nsm26")->AsString=="-") toExcel(Sh,"d_26",i,DM->qObnovlenie->FieldByName("nsm26")->AsString.c_str());
          else if (DM->qObnovlenie->FieldByName("nsm26")->AsString==0) toExcel(Sh,"d_26",i,"в");
          else toExcel(Sh,"d_26",i,DM->qObnovlenie->FieldByName("nsm26")->AsFloat);
          //27
          if (DM->qObnovlenie->FieldByName("nsm27")->AsString=="П" || DM->qObnovlenie->FieldByName("nsm27")->AsString=="-") toExcel(Sh,"d_27",i,DM->qObnovlenie->FieldByName("nsm27")->AsString.c_str());
          else if (DM->qObnovlenie->FieldByName("nsm27")->AsString==0) toExcel(Sh,"d_27",i,"в");
          else toExcel(Sh,"d_27",i,DM->qObnovlenie->FieldByName("nsm27")->AsFloat);
          //28
          if (DM->qObnovlenie->FieldByName("nsm28")->AsString=="П" || DM->qObnovlenie->FieldByName("nsm28")->AsString=="-") toExcel(Sh,"d_28",i,DM->qObnovlenie->FieldByName("nsm28")->AsString.c_str());
          else if (DM->qObnovlenie->FieldByName("nsm28")->AsString==0) toExcel(Sh,"d_28",i,"в");
          else toExcel(Sh,"d_28",i,DM->qObnovlenie->FieldByName("nsm28")->AsFloat);
          //29
          if (DM->qObnovlenie->FieldByName("nsm29")->AsString=="П" || DM->qObnovlenie->FieldByName("nsm29")->AsString=="-") toExcel(Sh,"d_29",i,DM->qObnovlenie->FieldByName("nsm29")->AsString.c_str());
          else if (DM->qObnovlenie->FieldByName("nsm29")->AsString==0) toExcel(Sh,"d_29",i,"в");
          else toExcel(Sh,"d_29",i,DM->qObnovlenie->FieldByName("nsm29")->AsFloat);
          //30
          if (DM->qObnovlenie->FieldByName("nsm30")->AsString=="П" || DM->qObnovlenie->FieldByName("nsm30")->AsString=="-") toExcel(Sh,"d_30",i,DM->qObnovlenie->FieldByName("nsm30")->AsString.c_str());
          else if (DM->qObnovlenie->FieldByName("nsm30")->AsString==0) toExcel(Sh,"d_30",i,"в");
          else toExcel(Sh,"d_30",i,DM->qObnovlenie->FieldByName("nsm30")->AsFloat);
          //31
          if (DM->qObnovlenie->FieldByName("nsm31")->AsString=="П" || DM->qObnovlenie->FieldByName("nsm31")->AsString=="-") toExcel(Sh,"d_31",i,DM->qObnovlenie->FieldByName("nsm31")->AsString.c_str());
          else if (DM->qObnovlenie->FieldByName("nsm31")->AsString==0) toExcel(Sh,"d_31",i,"в");
          else toExcel(Sh,"d_31",i,DM->qObnovlenie->FieldByName("nsm31")->AsFloat);

          //закрашивание смен
          if (graf!=26)
            {
              if (DM->qObnovlenie->FieldByName("nsm1")->AsString=="П" || DM->qObnovlenie->FieldByName("nsm1")->AsString==0) Sh.OlePropertyGet("Range", ("C"+IntToStr(n+1)).c_str()).OlePropertyGet("Interior").OlePropertySet("ColorIndex",35);
              if (DM->qObnovlenie->FieldByName("nsm2")->AsString=="П" || DM->qObnovlenie->FieldByName("nsm2")->AsString==0) Sh.OlePropertyGet("Range", ("D"+IntToStr(n+1)).c_str()).OlePropertyGet("Interior").OlePropertySet("ColorIndex",35);
              if (DM->qObnovlenie->FieldByName("nsm3")->AsString=="П" || DM->qObnovlenie->FieldByName("nsm3")->AsString==0) Sh.OlePropertyGet("Range", ("E"+IntToStr(n+1)).c_str()).OlePropertyGet("Interior").OlePropertySet("ColorIndex",35);
              if (DM->qObnovlenie->FieldByName("nsm4")->AsString=="П" || DM->qObnovlenie->FieldByName("nsm4")->AsString==0) Sh.OlePropertyGet("Range", ("F"+IntToStr(n+1)).c_str()).OlePropertyGet("Interior").OlePropertySet("ColorIndex",35);
              if (DM->qObnovlenie->FieldByName("nsm5")->AsString=="П" || DM->qObnovlenie->FieldByName("nsm5")->AsString==0) Sh.OlePropertyGet("Range", ("G"+IntToStr(n+1)).c_str()).OlePropertyGet("Interior").OlePropertySet("ColorIndex",35);
              if (DM->qObnovlenie->FieldByName("nsm6")->AsString=="П" || DM->qObnovlenie->FieldByName("nsm6")->AsString==0) Sh.OlePropertyGet("Range", ("H"+IntToStr(n+1)).c_str()).OlePropertyGet("Interior").OlePropertySet("ColorIndex",35);
              if (DM->qObnovlenie->FieldByName("nsm7")->AsString=="П" || DM->qObnovlenie->FieldByName("nsm7")->AsString==0) Sh.OlePropertyGet("Range", ("I"+IntToStr(n+1)).c_str()).OlePropertyGet("Interior").OlePropertySet("ColorIndex",35);
              if (DM->qObnovlenie->FieldByName("nsm8")->AsString=="П" || DM->qObnovlenie->FieldByName("nsm8")->AsString==0) Sh.OlePropertyGet("Range", ("J"+IntToStr(n+1)).c_str()).OlePropertyGet("Interior").OlePropertySet("ColorIndex",35);
              if (DM->qObnovlenie->FieldByName("nsm9")->AsString=="П" || DM->qObnovlenie->FieldByName("nsm9")->AsString==0) Sh.OlePropertyGet("Range", ("K"+IntToStr(n+1)).c_str()).OlePropertyGet("Interior").OlePropertySet("ColorIndex",35);
              if (DM->qObnovlenie->FieldByName("nsm10")->AsString=="П" || DM->qObnovlenie->FieldByName("nsm10")->AsString==0) Sh.OlePropertyGet("Range", ("L"+IntToStr(n+1)).c_str()).OlePropertyGet("Interior").OlePropertySet("ColorIndex",35);
              if (DM->qObnovlenie->FieldByName("nsm11")->AsString=="П" || DM->qObnovlenie->FieldByName("nsm11")->AsString==0) Sh.OlePropertyGet("Range", ("M"+IntToStr(n+1)).c_str()).OlePropertyGet("Interior").OlePropertySet("ColorIndex",35);
              if (DM->qObnovlenie->FieldByName("nsm12")->AsString=="П" || DM->qObnovlenie->FieldByName("nsm12")->AsString==0) Sh.OlePropertyGet("Range", ("N"+IntToStr(n+1)).c_str()).OlePropertyGet("Interior").OlePropertySet("ColorIndex",35);
              if (DM->qObnovlenie->FieldByName("nsm13")->AsString=="П" || DM->qObnovlenie->FieldByName("nsm13")->AsString==0) Sh.OlePropertyGet("Range", ("O"+IntToStr(n+1)).c_str()).OlePropertyGet("Interior").OlePropertySet("ColorIndex",35);
              if (DM->qObnovlenie->FieldByName("nsm14")->AsString=="П" || DM->qObnovlenie->FieldByName("nsm14")->AsString==0) Sh.OlePropertyGet("Range", ("P"+IntToStr(n+1)).c_str()).OlePropertyGet("Interior").OlePropertySet("ColorIndex",35);
              if (DM->qObnovlenie->FieldByName("nsm15")->AsString=="П" || DM->qObnovlenie->FieldByName("nsm15")->AsString==0) Sh.OlePropertyGet("Range", ("Q"+IntToStr(n+1)).c_str()).OlePropertyGet("Interior").OlePropertySet("ColorIndex",35);
              if (DM->qObnovlenie->FieldByName("nsm16")->AsString=="П" || DM->qObnovlenie->FieldByName("nsm16")->AsString==0) Sh.OlePropertyGet("Range", ("R"+IntToStr(n+1)).c_str()).OlePropertyGet("Interior").OlePropertySet("ColorIndex",35);
              if (DM->qObnovlenie->FieldByName("nsm17")->AsString=="П" || DM->qObnovlenie->FieldByName("nsm17")->AsString==0) Sh.OlePropertyGet("Range", ("S"+IntToStr(n+1)).c_str()).OlePropertyGet("Interior").OlePropertySet("ColorIndex",35);
              if (DM->qObnovlenie->FieldByName("nsm18")->AsString=="П" || DM->qObnovlenie->FieldByName("nsm18")->AsString==0) Sh.OlePropertyGet("Range", ("T"+IntToStr(n+1)).c_str()).OlePropertyGet("Interior").OlePropertySet("ColorIndex",35);
              if (DM->qObnovlenie->FieldByName("nsm19")->AsString=="П" || DM->qObnovlenie->FieldByName("nsm19")->AsString==0) Sh.OlePropertyGet("Range", ("U"+IntToStr(n+1)).c_str()).OlePropertyGet("Interior").OlePropertySet("ColorIndex",35);
              if (DM->qObnovlenie->FieldByName("nsm20")->AsString=="П" || DM->qObnovlenie->FieldByName("nsm20")->AsString==0) Sh.OlePropertyGet("Range", ("V"+IntToStr(n+1)).c_str()).OlePropertyGet("Interior").OlePropertySet("ColorIndex",35);
              if (DM->qObnovlenie->FieldByName("nsm21")->AsString=="П" || DM->qObnovlenie->FieldByName("nsm21")->AsString==0) Sh.OlePropertyGet("Range", ("W"+IntToStr(n+1)).c_str()).OlePropertyGet("Interior").OlePropertySet("ColorIndex",35);
              if (DM->qObnovlenie->FieldByName("nsm22")->AsString=="П" || DM->qObnovlenie->FieldByName("nsm22")->AsString==0) Sh.OlePropertyGet("Range", ("X"+IntToStr(n+1)).c_str()).OlePropertyGet("Interior").OlePropertySet("ColorIndex",35);
              if (DM->qObnovlenie->FieldByName("nsm23")->AsString=="П" || DM->qObnovlenie->FieldByName("nsm23")->AsString==0) Sh.OlePropertyGet("Range", ("Y"+IntToStr(n+1)).c_str()).OlePropertyGet("Interior").OlePropertySet("ColorIndex",35);
              if (DM->qObnovlenie->FieldByName("nsm24")->AsString=="П" || DM->qObnovlenie->FieldByName("nsm24")->AsString==0) Sh.OlePropertyGet("Range", ("Z"+IntToStr(n+1)).c_str()).OlePropertyGet("Interior").OlePropertySet("ColorIndex",35);
              if (DM->qObnovlenie->FieldByName("nsm25")->AsString=="П" || DM->qObnovlenie->FieldByName("nsm25")->AsString==0) Sh.OlePropertyGet("Range", ("AA"+IntToStr(n+1)).c_str()).OlePropertyGet("Interior").OlePropertySet("ColorIndex",35);
              if (DM->qObnovlenie->FieldByName("nsm26")->AsString=="П" || DM->qObnovlenie->FieldByName("nsm26")->AsString==0) Sh.OlePropertyGet("Range", ("AB"+IntToStr(n+1)).c_str()).OlePropertyGet("Interior").OlePropertySet("ColorIndex",35);
              if (DM->qObnovlenie->FieldByName("nsm27")->AsString=="П" || DM->qObnovlenie->FieldByName("nsm27")->AsString==0) Sh.OlePropertyGet("Range", ("AC"+IntToStr(n+1)).c_str()).OlePropertyGet("Interior").OlePropertySet("ColorIndex",35);
              if (DM->qObnovlenie->FieldByName("nsm28")->AsString=="П" || DM->qObnovlenie->FieldByName("nsm28")->AsString==0) Sh.OlePropertyGet("Range", ("AD"+IntToStr(n+1)).c_str()).OlePropertyGet("Interior").OlePropertySet("ColorIndex",35);
              if (DM->qObnovlenie->FieldByName("nsm29")->AsString=="П" || DM->qObnovlenie->FieldByName("nsm29")->AsString==0) Sh.OlePropertyGet("Range", ("AE"+IntToStr(n+1)).c_str()).OlePropertyGet("Interior").OlePropertySet("ColorIndex",35);
              if (DM->qObnovlenie->FieldByName("nsm30")->AsString=="П" || DM->qObnovlenie->FieldByName("nsm30")->AsString==0) Sh.OlePropertyGet("Range", ("AF"+IntToStr(n+1)).c_str()).OlePropertyGet("Interior").OlePropertySet("ColorIndex",35);
              if (DM->qObnovlenie->FieldByName("nsm31")->AsString=="П" || DM->qObnovlenie->FieldByName("nsm31")->AsString==0) Sh.OlePropertyGet("Range", ("AG"+IntToStr(n+1)).c_str()).OlePropertyGet("Interior").OlePropertySet("ColorIndex",35);
            }
          else //для графиков начинающихся с 26 числа
            {
              if (DM->qObnovlenie->FieldByName("nsm1")->AsString=="П" || DM->qObnovlenie->FieldByName("nsm1")->AsString==0) Sh.OlePropertyGet("Range", ("I"+IntToStr(n+1)).c_str()).OlePropertyGet("Interior").OlePropertySet("ColorIndex",35);
              if (DM->qObnovlenie->FieldByName("nsm2")->AsString=="П" || DM->qObnovlenie->FieldByName("nsm2")->AsString==0) Sh.OlePropertyGet("Range", ("J"+IntToStr(n+1)).c_str()).OlePropertyGet("Interior").OlePropertySet("ColorIndex",35);
              if (DM->qObnovlenie->FieldByName("nsm3")->AsString=="П" || DM->qObnovlenie->FieldByName("nsm3")->AsString==0) Sh.OlePropertyGet("Range", ("K"+IntToStr(n+1)).c_str()).OlePropertyGet("Interior").OlePropertySet("ColorIndex",35);
              if (DM->qObnovlenie->FieldByName("nsm4")->AsString=="П" || DM->qObnovlenie->FieldByName("nsm4")->AsString==0) Sh.OlePropertyGet("Range", ("L"+IntToStr(n+1)).c_str()).OlePropertyGet("Interior").OlePropertySet("ColorIndex",35);
              if (DM->qObnovlenie->FieldByName("nsm5")->AsString=="П" || DM->qObnovlenie->FieldByName("nsm5")->AsString==0) Sh.OlePropertyGet("Range", ("M"+IntToStr(n+1)).c_str()).OlePropertyGet("Interior").OlePropertySet("ColorIndex",35);
              if (DM->qObnovlenie->FieldByName("nsm6")->AsString=="П" || DM->qObnovlenie->FieldByName("nsm6")->AsString==0) Sh.OlePropertyGet("Range", ("N"+IntToStr(n+1)).c_str()).OlePropertyGet("Interior").OlePropertySet("ColorIndex",35);
              if (DM->qObnovlenie->FieldByName("nsm7")->AsString=="П" || DM->qObnovlenie->FieldByName("nsm7")->AsString==0) Sh.OlePropertyGet("Range", ("O"+IntToStr(n+1)).c_str()).OlePropertyGet("Interior").OlePropertySet("ColorIndex",35);
              if (DM->qObnovlenie->FieldByName("nsm8")->AsString=="П" || DM->qObnovlenie->FieldByName("nsm8")->AsString==0) Sh.OlePropertyGet("Range", ("P"+IntToStr(n+1)).c_str()).OlePropertyGet("Interior").OlePropertySet("ColorIndex",35);
              if (DM->qObnovlenie->FieldByName("nsm9")->AsString=="П" || DM->qObnovlenie->FieldByName("nsm9")->AsString==0) Sh.OlePropertyGet("Range", ("Q"+IntToStr(n+1)).c_str()).OlePropertyGet("Interior").OlePropertySet("ColorIndex",35);
              if (DM->qObnovlenie->FieldByName("nsm10")->AsString=="П" || DM->qObnovlenie->FieldByName("nsm10")->AsString==0) Sh.OlePropertyGet("Range", ("R"+IntToStr(n+1)).c_str()).OlePropertyGet("Interior").OlePropertySet("ColorIndex",35);
              if (DM->qObnovlenie->FieldByName("nsm11")->AsString=="П" || DM->qObnovlenie->FieldByName("nsm11")->AsString==0) Sh.OlePropertyGet("Range", ("S"+IntToStr(n+1)).c_str()).OlePropertyGet("Interior").OlePropertySet("ColorIndex",35);
              if (DM->qObnovlenie->FieldByName("nsm12")->AsString=="П" || DM->qObnovlenie->FieldByName("nsm12")->AsString==0) Sh.OlePropertyGet("Range", ("T"+IntToStr(n+1)).c_str()).OlePropertyGet("Interior").OlePropertySet("ColorIndex",35);
              if (DM->qObnovlenie->FieldByName("nsm13")->AsString=="П" || DM->qObnovlenie->FieldByName("nsm13")->AsString==0) Sh.OlePropertyGet("Range", ("U"+IntToStr(n+1)).c_str()).OlePropertyGet("Interior").OlePropertySet("ColorIndex",35);
              if (DM->qObnovlenie->FieldByName("nsm14")->AsString=="П" || DM->qObnovlenie->FieldByName("nsm14")->AsString==0) Sh.OlePropertyGet("Range", ("V"+IntToStr(n+1)).c_str()).OlePropertyGet("Interior").OlePropertySet("ColorIndex",35);
              if (DM->qObnovlenie->FieldByName("nsm15")->AsString=="П" || DM->qObnovlenie->FieldByName("nsm15")->AsString==0) Sh.OlePropertyGet("Range", ("W"+IntToStr(n+1)).c_str()).OlePropertyGet("Interior").OlePropertySet("ColorIndex",35);
              if (DM->qObnovlenie->FieldByName("nsm16")->AsString=="П" || DM->qObnovlenie->FieldByName("nsm16")->AsString==0) Sh.OlePropertyGet("Range", ("X"+IntToStr(n+1)).c_str()).OlePropertyGet("Interior").OlePropertySet("ColorIndex",35);
              if (DM->qObnovlenie->FieldByName("nsm17")->AsString=="П" || DM->qObnovlenie->FieldByName("nsm17")->AsString==0) Sh.OlePropertyGet("Range", ("Y"+IntToStr(n+1)).c_str()).OlePropertyGet("Interior").OlePropertySet("ColorIndex",35);
              if (DM->qObnovlenie->FieldByName("nsm18")->AsString=="П" || DM->qObnovlenie->FieldByName("nsm18")->AsString==0) Sh.OlePropertyGet("Range", ("Z"+IntToStr(n+1)).c_str()).OlePropertyGet("Interior").OlePropertySet("ColorIndex",35);
              if (DM->qObnovlenie->FieldByName("nsm19")->AsString=="П" || DM->qObnovlenie->FieldByName("nsm19")->AsString==0) Sh.OlePropertyGet("Range", ("AA"+IntToStr(n+1)).c_str()).OlePropertyGet("Interior").OlePropertySet("ColorIndex",35);
              if (DM->qObnovlenie->FieldByName("nsm20")->AsString=="П" || DM->qObnovlenie->FieldByName("nsm20")->AsString==0) Sh.OlePropertyGet("Range", ("AB"+IntToStr(n+1)).c_str()).OlePropertyGet("Interior").OlePropertySet("ColorIndex",35);
              if (DM->qObnovlenie->FieldByName("nsm21")->AsString=="П" || DM->qObnovlenie->FieldByName("nsm21")->AsString==0) Sh.OlePropertyGet("Range", ("AC"+IntToStr(n+1)).c_str()).OlePropertyGet("Interior").OlePropertySet("ColorIndex",35);
              if (DM->qObnovlenie->FieldByName("nsm22")->AsString=="П" || DM->qObnovlenie->FieldByName("nsm22")->AsString==0) Sh.OlePropertyGet("Range", ("AD"+IntToStr(n+1)).c_str()).OlePropertyGet("Interior").OlePropertySet("ColorIndex",35);
              if (DM->qObnovlenie->FieldByName("nsm23")->AsString=="П" || DM->qObnovlenie->FieldByName("nsm23")->AsString==0) Sh.OlePropertyGet("Range", ("AE"+IntToStr(n+1)).c_str()).OlePropertyGet("Interior").OlePropertySet("ColorIndex",35);
              if (DM->qObnovlenie->FieldByName("nsm24")->AsString=="П" || DM->qObnovlenie->FieldByName("nsm24")->AsString==0) Sh.OlePropertyGet("Range", ("AF"+IntToStr(n+1)).c_str()).OlePropertyGet("Interior").OlePropertySet("ColorIndex",35);
              if (DM->qObnovlenie->FieldByName("nsm25")->AsString=="П" || DM->qObnovlenie->FieldByName("nsm25")->AsString==0) Sh.OlePropertyGet("Range", ("AG"+IntToStr(n+1)).c_str()).OlePropertyGet("Interior").OlePropertySet("ColorIndex",35);
              if (DM->qObnovlenie->FieldByName("nsm26")->AsString=="П" || DM->qObnovlenie->FieldByName("nsm26")->AsString==0) Sh.OlePropertyGet("Range", ("C"+IntToStr(n+1)).c_str()).OlePropertyGet("Interior").OlePropertySet("ColorIndex",35);
              if (DM->qObnovlenie->FieldByName("nsm27")->AsString=="П" || DM->qObnovlenie->FieldByName("nsm27")->AsString==0) Sh.OlePropertyGet("Range", ("D"+IntToStr(n+1)).c_str()).OlePropertyGet("Interior").OlePropertySet("ColorIndex",35);
              if (DM->qObnovlenie->FieldByName("nsm28")->AsString=="П" || DM->qObnovlenie->FieldByName("nsm28")->AsString==0) Sh.OlePropertyGet("Range", ("E"+IntToStr(n+1)).c_str()).OlePropertyGet("Interior").OlePropertySet("ColorIndex",35);
              if (DM->qObnovlenie->FieldByName("nsm29")->AsString=="П" || DM->qObnovlenie->FieldByName("nsm29")->AsString==0) Sh.OlePropertyGet("Range", ("F"+IntToStr(n+1)).c_str()).OlePropertyGet("Interior").OlePropertySet("ColorIndex",35);
              if (DM->qObnovlenie->FieldByName("nsm30")->AsString=="П" || DM->qObnovlenie->FieldByName("nsm30")->AsString==0) Sh.OlePropertyGet("Range", ("G"+IntToStr(n+1)).c_str()).OlePropertyGet("Interior").OlePropertySet("ColorIndex",35);
              if (DM->qObnovlenie->FieldByName("nsm31")->AsString=="П" || DM->qObnovlenie->FieldByName("nsm31")->AsString==0) Sh.OlePropertyGet("Range", ("H"+IntToStr(n+1)).c_str()).OlePropertyGet("Interior").OlePropertySet("ColorIndex",35);
            }

          //Месяц
          Sh.OlePropertyGet("Range",("A"+IntToStr(n+1)+":A"+IntToStr(n+2)).c_str()).OleProcedure("Merge"); //объединение ячеек
          Sh.OlePropertyGet("Range",("A"+IntToStr(n+1)+":A"+IntToStr(n+2)).c_str()).OlePropertySet("HorizontalAlignment", xlCenter); //выровнять по гор.
          Sh.OlePropertyGet("Range",("A"+IntToStr(n+1)+":A"+IntToStr(n+2)).c_str()).OlePropertySet("VerticalAlignment", xlCenter); //выровнять по верт.


          //Бригада
          Sh.OlePropertyGet("Range",("B"+IntToStr(n+1)+":B"+IntToStr(n+2)).c_str()).OleProcedure("Merge"); //объединение ячеек
          Sh.OlePropertyGet("Range",("B"+IntToStr(n+1)+":B"+IntToStr(n+2)).c_str()).OlePropertySet("HorizontalAlignment", xlCenter); //выровнять по гор.
          Sh.OlePropertyGet("Range",("B"+IntToStr(n+1)+":B"+IntToStr(n+2)).c_str()).OlePropertySet("VerticalAlignment", xlCenter); //выровнять по верт.

          //подсчет количества смен
          Sh.OlePropertyGet("Range", "sm1").OlePropertyGet("Offset", i).OlePropertySet("Formula", ("=СЧЁТЕСЛИ(C"+IntToStr(num)+":AG"+IntToStr(num)+";\">0\")").c_str());
          Sh.OlePropertyGet("Range",("AH"+IntToStr(n+1)+":AH"+IntToStr(n+2)).c_str()).OleProcedure("Merge"); //объединение ячеек
          Sh.OlePropertyGet("Range",("AH"+IntToStr(n+1)+":AH"+IntToStr(n+2)).c_str()).OlePropertySet("HorizontalAlignment", xlCenter); //выровнять по гор.
          Sh.OlePropertyGet("Range",("AH"+IntToStr(n+1)+":AH"+IntToStr(n+2)).c_str()).OlePropertySet("VerticalAlignment", xlCenter); //выровнять по верт.

          //суммы по часам
          //общие
          toExcel(Sh,"chf",i,DM->qObnovlenie->FieldByName("chf")->AsFloat);
          Sh.OlePropertyGet("Range",("AI"+IntToStr(n+1)+":AI"+IntToStr(n+2)).c_str()).OleProcedure("Merge");
          Sh.OlePropertyGet("Range",("AI"+IntToStr(n+1)+":AI"+IntToStr(n+2)).c_str()).OlePropertySet("HorizontalAlignment", xlCenter); //выровнять по гор.
          Sh.OlePropertyGet("Range",("AI"+IntToStr(n+1)+":AI"+IntToStr(n+2)).c_str()).OlePropertySet("VerticalAlignment", xlCenter); //выровнять по верт.

          //норма
          toExcel(Sh,"norma",i,DM->qObnovlenie->FieldByName("norma")->AsFloat);
          Sh.OlePropertyGet("Range",("AJ"+IntToStr(n+1)+":AJ"+IntToStr(n+2)).c_str()).OleProcedure("Merge");
          Sh.OlePropertyGet("Range",("AJ"+IntToStr(n+1)+":AJ"+IntToStr(n+2)).c_str()).OlePropertySet("HorizontalAlignment", xlCenter); //выровнять по гор.
          Sh.OlePropertyGet("Range",("AJ"+IntToStr(n+1)+":AJ"+IntToStr(n+2)).c_str()).OlePropertySet("VerticalAlignment", xlCenter); //выровнять по верт.

          //вечерние
          toExcel(Sh,"vch",i,DM->qObnovlenie->FieldByName("vch")->AsFloat);
          Sh.OlePropertyGet("Range",("AK"+IntToStr(n+1)+":AK"+IntToStr(n+2)).c_str()).OleProcedure("Merge");
          Sh.OlePropertyGet("Range",("AK"+IntToStr(n+1)+":AK"+IntToStr(n+2)).c_str()).OlePropertySet("HorizontalAlignment", xlCenter); //выровнять по гор.
          Sh.OlePropertyGet("Range",("AK"+IntToStr(n+1)+":AK"+IntToStr(n+2)).c_str()).OlePropertySet("VerticalAlignment", xlCenter); //выровнять по верт.

          //ночные
          toExcel(Sh,"nch",i,DM->qObnovlenie->FieldByName("nch")->AsFloat);
          Sh.OlePropertyGet("Range",("AL"+IntToStr(n+1)+":AL"+IntToStr(n+2)).c_str()).OleProcedure("Merge");
          Sh.OlePropertyGet("Range",("AL"+IntToStr(n+1)+":AL"+IntToStr(n+2)).c_str()).OlePropertySet("HorizontalAlignment", xlCenter); //выровнять по гор.
          Sh.OlePropertyGet("Range",("AL"+IntToStr(n+1)+":AL"+IntToStr(n+2)).c_str()).OlePropertySet("VerticalAlignment", xlCenter); //выровнять по верт.

          //праздничные
          toExcel(Sh,"pch",i,DM->qObnovlenie->FieldByName("pch")->AsFloat);
          Sh.OlePropertyGet("Range",("AM"+IntToStr(n+1)+":AM"+IntToStr(n+2)).c_str()).OleProcedure("Merge");
          Sh.OlePropertyGet("Range",("AM"+IntToStr(n+1)+":AM"+IntToStr(n+2)).c_str()).OlePropertySet("HorizontalAlignment", xlCenter); //выровнять по гор.
          Sh.OlePropertyGet("Range",("AM"+IntToStr(n+1)+":AM"+IntToStr(n+2)).c_str()).OlePropertySet("VerticalAlignment", xlCenter); //выровнять по верт.

          //переработка
          toExcel(Sh,"pgraf",i,DM->qObnovlenie->FieldByName("pgraf")->AsFloat);
          Sh.OlePropertyGet("Range",("AN"+IntToStr(n+1)+":AN"+IntToStr(n+2)).c_str()).OleProcedure("Merge");
          Sh.OlePropertyGet("Range",("AN"+IntToStr(n+1)+":AN"+IntToStr(n+2)).c_str()).OlePropertySet("HorizontalAlignment", xlCenter); //выровнять по гор.
          Sh.OlePropertyGet("Range",("AN"+IntToStr(n+1)+":AN"+IntToStr(n+2)).c_str()).OlePropertySet("VerticalAlignment", xlCenter); //выровнять по верт.

          i++;
          num++;
          n++;

           //рисуем сетку
          Sh.OlePropertyGet("Range",("A"+IntToStr(n)+":AN"+IntToStr(n)).c_str()).OlePropertyGet("Borders").OlePropertySet("LineStyle", 1);

          //вывод часов
          if (DM->qObnovlenie->FieldByName("chf1")->AsString=="П" || DM->qObnovlenie->FieldByName("chf1")->AsString=="-") toExcel(Sh,"d_1",i,DM->qObnovlenie->FieldByName("chf1")->AsString.c_str());
          else toExcel(Sh,"d_1",i,DM->qObnovlenie->FieldByName("chf1")->AsFloat);
          if (DM->qObnovlenie->FieldByName("chf2")->AsString=="П" || DM->qObnovlenie->FieldByName("chf2")->AsString=="-") toExcel(Sh,"d_2",i, DM->qObnovlenie->FieldByName("chf2")->AsString.c_str());
          else toExcel(Sh,"d_2",i, DM->qObnovlenie->FieldByName("chf2")->AsFloat);
          if (DM->qObnovlenie->FieldByName("chf3")->AsString=="П" || DM->qObnovlenie->FieldByName("chf3")->AsString=="-") toExcel(Sh,"d_3",i,DM->qObnovlenie->FieldByName("chf3")->AsString.c_str());
          else toExcel(Sh,"d_3",i,DM->qObnovlenie->FieldByName("chf3")->AsFloat);
          if (DM->qObnovlenie->FieldByName("chf4")->AsString=="П" || DM->qObnovlenie->FieldByName("chf4")->AsString=="-") toExcel(Sh,"d_4",i,DM->qObnovlenie->FieldByName("chf4")->AsString.c_str());
          else toExcel(Sh,"d_4",i,DM->qObnovlenie->FieldByName("chf4")->AsFloat);
          if (DM->qObnovlenie->FieldByName("chf5")->AsString=="П" || DM->qObnovlenie->FieldByName("chf5")->AsString=="-") toExcel(Sh,"d_5",i,DM->qObnovlenie->FieldByName("chf5")->AsString.c_str());
          else toExcel(Sh,"d_5",i,DM->qObnovlenie->FieldByName("chf5")->AsFloat);
          if (DM->qObnovlenie->FieldByName("chf6")->AsString=="П" || DM->qObnovlenie->FieldByName("chf6")->AsString=="-") toExcel(Sh,"d_6",i,DM->qObnovlenie->FieldByName("chf6")->AsString.c_str());
          else toExcel(Sh,"d_6",i,DM->qObnovlenie->FieldByName("chf6")->AsFloat);
          if (DM->qObnovlenie->FieldByName("chf7")->AsString=="П" || DM->qObnovlenie->FieldByName("chf7")->AsString=="-") toExcel(Sh,"d_7",i,DM->qObnovlenie->FieldByName("chf7")->AsString.c_str());
          else toExcel(Sh,"d_7",i,DM->qObnovlenie->FieldByName("chf7")->AsFloat);
          if (DM->qObnovlenie->FieldByName("chf8")->AsString=="П" || DM->qObnovlenie->FieldByName("chf8")->AsString=="-") toExcel(Sh,"d_8",i,DM->qObnovlenie->FieldByName("chf8")->AsString.c_str());
          else toExcel(Sh,"d_8",i,DM->qObnovlenie->FieldByName("chf8")->AsFloat);
          if (DM->qObnovlenie->FieldByName("chf9")->AsString=="П" || DM->qObnovlenie->FieldByName("chf9")->AsString=="-") toExcel(Sh,"d_9",i,DM->qObnovlenie->FieldByName("chf9")->AsString.c_str());
          else toExcel(Sh,"d_9",i,DM->qObnovlenie->FieldByName("chf9")->AsFloat);
          if (DM->qObnovlenie->FieldByName("chf10")->AsString=="П" || DM->qObnovlenie->FieldByName("chf10")->AsString=="-") toExcel(Sh,"d_10",i,DM->qObnovlenie->FieldByName("chf10")->AsString.c_str());
          else toExcel(Sh,"d_10",i,DM->qObnovlenie->FieldByName("chf10")->AsFloat);
          if (DM->qObnovlenie->FieldByName("chf11")->AsString=="П" || DM->qObnovlenie->FieldByName("chf11")->AsString=="-") toExcel(Sh,"d_11",i,DM->qObnovlenie->FieldByName("chf11")->AsString.c_str());
          else toExcel(Sh,"d_11",i,DM->qObnovlenie->FieldByName("chf11")->AsFloat);
          if (DM->qObnovlenie->FieldByName("chf12")->AsString=="П" || DM->qObnovlenie->FieldByName("chf12")->AsString=="-") toExcel(Sh,"d_12",i,DM->qObnovlenie->FieldByName("chf12")->AsString.c_str());
          else toExcel(Sh,"d_12",i,DM->qObnovlenie->FieldByName("chf12")->AsFloat);
          if (DM->qObnovlenie->FieldByName("chf13")->AsString=="П" || DM->qObnovlenie->FieldByName("chf13")->AsString=="-") toExcel(Sh,"d_13",i,DM->qObnovlenie->FieldByName("chf13")->AsString.c_str());
          else toExcel(Sh,"d_13",i,DM->qObnovlenie->FieldByName("chf13")->AsFloat);
          if (DM->qObnovlenie->FieldByName("chf14")->AsString=="П" || DM->qObnovlenie->FieldByName("chf14")->AsString=="-") toExcel(Sh,"d_14",i,DM->qObnovlenie->FieldByName("chf14")->AsString.c_str());
          else toExcel(Sh,"d_14",i,DM->qObnovlenie->FieldByName("chf14")->AsFloat);
          if (DM->qObnovlenie->FieldByName("chf15")->AsString=="П" || DM->qObnovlenie->FieldByName("chf15")->AsString=="-") toExcel(Sh,"d_15",i,DM->qObnovlenie->FieldByName("chf15")->AsString.c_str());
          else toExcel(Sh,"d_15",i,DM->qObnovlenie->FieldByName("chf15")->AsFloat);
          if (DM->qObnovlenie->FieldByName("chf16")->AsString=="П" || DM->qObnovlenie->FieldByName("chf16")->AsString=="-") toExcel(Sh,"d_16",i,DM->qObnovlenie->FieldByName("chf16")->AsString.c_str());
          else toExcel(Sh,"d_16",i,DM->qObnovlenie->FieldByName("chf16")->AsFloat);
          if (DM->qObnovlenie->FieldByName("chf17")->AsString=="П" || DM->qObnovlenie->FieldByName("chf17")->AsString=="-") toExcel(Sh,"d_17",i,DM->qObnovlenie->FieldByName("chf17")->AsString.c_str());
          else toExcel(Sh,"d_17",i,DM->qObnovlenie->FieldByName("chf17")->AsFloat);
          if (DM->qObnovlenie->FieldByName("chf18")->AsString=="П" || DM->qObnovlenie->FieldByName("chf18")->AsString=="-") toExcel(Sh,"d_18",i,DM->qObnovlenie->FieldByName("chf18")->AsString.c_str());
          else toExcel(Sh,"d_18",i,DM->qObnovlenie->FieldByName("chf18")->AsFloat);
          if (DM->qObnovlenie->FieldByName("chf19")->AsString=="П" || DM->qObnovlenie->FieldByName("chf19")->AsString=="-") toExcel(Sh,"d_19",i,DM->qObnovlenie->FieldByName("chf19")->AsString.c_str());
          else toExcel(Sh,"d_19",i,DM->qObnovlenie->FieldByName("chf19")->AsFloat);
          if (DM->qObnovlenie->FieldByName("chf20")->AsString=="П" || DM->qObnovlenie->FieldByName("chf20")->AsString=="-") toExcel(Sh,"d_20",i,DM->qObnovlenie->FieldByName("chf20")->AsString.c_str());
          else toExcel(Sh,"d_20",i,DM->qObnovlenie->FieldByName("chf20")->AsFloat);
          if (DM->qObnovlenie->FieldByName("chf21")->AsString=="П" || DM->qObnovlenie->FieldByName("chf21")->AsString=="-") toExcel(Sh,"d_21",i,DM->qObnovlenie->FieldByName("chf21")->AsString.c_str());
          else toExcel(Sh,"d_21",i,DM->qObnovlenie->FieldByName("chf21")->AsFloat);
          if (DM->qObnovlenie->FieldByName("chf22")->AsString=="П" || DM->qObnovlenie->FieldByName("chf22")->AsString=="-") toExcel(Sh,"d_22",i,DM->qObnovlenie->FieldByName("chf22")->AsString.c_str());
          else toExcel(Sh,"d_22",i,DM->qObnovlenie->FieldByName("chf22")->AsFloat);
          if (DM->qObnovlenie->FieldByName("chf23")->AsString=="П" || DM->qObnovlenie->FieldByName("chf23")->AsString=="-") toExcel(Sh,"d_23",i,DM->qObnovlenie->FieldByName("chf23")->AsString.c_str());
          else toExcel(Sh,"d_23",i,DM->qObnovlenie->FieldByName("chf23")->AsFloat);
          if (DM->qObnovlenie->FieldByName("chf24")->AsString=="П" || DM->qObnovlenie->FieldByName("chf24")->AsString=="-") toExcel(Sh,"d_24",i,DM->qObnovlenie->FieldByName("chf24")->AsString.c_str());
          else toExcel(Sh,"d_24",i,DM->qObnovlenie->FieldByName("chf24")->AsFloat);
          if (DM->qObnovlenie->FieldByName("chf25")->AsString=="П" || DM->qObnovlenie->FieldByName("chf25")->AsString=="-") toExcel(Sh,"d_25",i,DM->qObnovlenie->FieldByName("chf25")->AsString.c_str());
          else toExcel(Sh,"d_25",i,DM->qObnovlenie->FieldByName("chf25")->AsFloat);
          if (DM->qObnovlenie->FieldByName("chf26")->AsString=="П" || DM->qObnovlenie->FieldByName("chf26")->AsString=="-") toExcel(Sh,"d_26",i,DM->qObnovlenie->FieldByName("chf26")->AsString.c_str());
          else toExcel(Sh,"d_26",i,DM->qObnovlenie->FieldByName("chf26")->AsFloat);
          if (DM->qObnovlenie->FieldByName("chf27")->AsString=="П" || DM->qObnovlenie->FieldByName("chf27")->AsString=="-") toExcel(Sh,"d_27",i,DM->qObnovlenie->FieldByName("chf27")->AsString.c_str());
          else toExcel(Sh,"d_27",i,DM->qObnovlenie->FieldByName("chf27")->AsFloat);
          if (DM->qObnovlenie->FieldByName("chf28")->AsString=="П" || DM->qObnovlenie->FieldByName("chf28")->AsString=="-") toExcel(Sh,"d_28",i,DM->qObnovlenie->FieldByName("chf28")->AsString.c_str());
          else toExcel(Sh,"d_28",i,DM->qObnovlenie->FieldByName("chf28")->AsFloat);
          if (DM->qObnovlenie->FieldByName("chf29")->AsString=="П" || DM->qObnovlenie->FieldByName("chf29")->AsString=="-") toExcel(Sh,"d_29",i,DM->qObnovlenie->FieldByName("chf29")->AsString.c_str());
          else toExcel(Sh,"d_29",i,DM->qObnovlenie->FieldByName("chf29")->AsFloat);
          if (DM->qObnovlenie->FieldByName("chf30")->AsString=="П" || DM->qObnovlenie->FieldByName("chf30")->AsString=="-") toExcel(Sh,"d_30",i,DM->qObnovlenie->FieldByName("chf30")->AsString.c_str());
          else toExcel(Sh,"d_30",i,DM->qObnovlenie->FieldByName("chf30")->AsFloat);
          if (DM->qObnovlenie->FieldByName("chf31")->AsString=="П" || DM->qObnovlenie->FieldByName("chf31")->AsString=="-") toExcel(Sh,"d_31",i,DM->qObnovlenie->FieldByName("chf31")->AsString.c_str());
          else toExcel(Sh,"d_31",i,DM->qObnovlenie->FieldByName("chf31")->AsFloat);



          //Окрашивание праздничных
              //для графиков начинающихся с 1 числа
              if (graf!=26)
                {
                  Variant locvalues[] = {god,DM->qObnovlenie->FieldByName("mes")->AsInteger,1};
                  if (DM->qPrazdDni->Locate("god;mes;den", VarArrayOf(locvalues, 2), SearchOptions << loCaseInsensitive)) Sh.OlePropertyGet("Range", ("C"+IntToStr(n+1)).c_str()).OlePropertyGet("Interior").OlePropertySet("ColorIndex",40);
                  Variant locvalues1[] = {god,DM->qObnovlenie->FieldByName("mes")->AsInteger,2};
                  if (DM->qPrazdDni->Locate("god;mes;den", VarArrayOf(locvalues1, 2), SearchOptions << loCaseInsensitive)) Sh.OlePropertyGet("Range", ("D"+IntToStr(n+1)).c_str()).OlePropertyGet("Interior").OlePropertySet("ColorIndex",40);
                  Variant locvalues2[] = {god,DM->qObnovlenie->FieldByName("mes")->AsInteger,3};
                  if (DM->qPrazdDni->Locate("god;mes;den", VarArrayOf(locvalues2, 2), SearchOptions << loCaseInsensitive)) Sh.OlePropertyGet("Range", ("E"+IntToStr(n+1)).c_str()).OlePropertyGet("Interior").OlePropertySet("ColorIndex",40);
                  Variant locvalues3[] = {god,DM->qObnovlenie->FieldByName("mes")->AsInteger,4};
                  if (DM->qPrazdDni->Locate("god;mes;den", VarArrayOf(locvalues3, 2), SearchOptions << loCaseInsensitive)) Sh.OlePropertyGet("Range", ("F"+IntToStr(n+1)).c_str()).OlePropertyGet("Interior").OlePropertySet("ColorIndex",40);
                  Variant locvalues4[] = {god,DM->qObnovlenie->FieldByName("mes")->AsInteger,5};
                  if (DM->qPrazdDni->Locate("god;mes;den", VarArrayOf(locvalues4, 2), SearchOptions << loCaseInsensitive)) Sh.OlePropertyGet("Range", ("G"+IntToStr(n+1)).c_str()).OlePropertyGet("Interior").OlePropertySet("ColorIndex",40);
                  Variant locvalues5[] = {god,DM->qObnovlenie->FieldByName("mes")->AsInteger,6};
                  if (DM->qPrazdDni->Locate("god;mes;den", VarArrayOf(locvalues5, 2), SearchOptions << loCaseInsensitive)) Sh.OlePropertyGet("Range", ("H"+IntToStr(n+1)).c_str()).OlePropertyGet("Interior").OlePropertySet("ColorIndex",40);
                  Variant locvalues6[] = {god,DM->qObnovlenie->FieldByName("mes")->AsInteger,7};
                  if (DM->qPrazdDni->Locate("god;mes;den", VarArrayOf(locvalues6, 2), SearchOptions << loCaseInsensitive)) Sh.OlePropertyGet("Range", ("I"+IntToStr(n+1)).c_str()).OlePropertyGet("Interior").OlePropertySet("ColorIndex",40);
                  Variant locvalues7[] = {god,DM->qObnovlenie->FieldByName("mes")->AsInteger,8};
                  if (DM->qPrazdDni->Locate("god;mes;den", VarArrayOf(locvalues7, 2), SearchOptions << loCaseInsensitive)) Sh.OlePropertyGet("Range", ("J"+IntToStr(n+1)).c_str()).OlePropertyGet("Interior").OlePropertySet("ColorIndex",40);
                  Variant locvalues8[] = {god,DM->qObnovlenie->FieldByName("mes")->AsInteger,9};
                  if (DM->qPrazdDni->Locate("god;mes;den", VarArrayOf(locvalues8, 2), SearchOptions << loCaseInsensitive)) Sh.OlePropertyGet("Range", ("K"+IntToStr(n+1)).c_str()).OlePropertyGet("Interior").OlePropertySet("ColorIndex",40);
                  Variant locvalues9[] = {god,DM->qObnovlenie->FieldByName("mes")->AsInteger,10};
                  if (DM->qPrazdDni->Locate("god;mes;den", VarArrayOf(locvalues9, 2), SearchOptions << loCaseInsensitive)) Sh.OlePropertyGet("Range", ("L"+IntToStr(n+1)).c_str()).OlePropertyGet("Interior").OlePropertySet("ColorIndex",40);
                  Variant locvalues10[] = {god,DM->qObnovlenie->FieldByName("mes")->AsInteger,11};
                  if (DM->qPrazdDni->Locate("god;mes;den", VarArrayOf(locvalues10, 2), SearchOptions << loCaseInsensitive)) Sh.OlePropertyGet("Range", ("M"+IntToStr(n+1)).c_str()).OlePropertyGet("Interior").OlePropertySet("ColorIndex",40);
                  Variant locvalues11[] = {god,DM->qObnovlenie->FieldByName("mes")->AsInteger,12};
                  if (DM->qPrazdDni->Locate("god;mes;den", VarArrayOf(locvalues11, 2), SearchOptions << loCaseInsensitive)) Sh.OlePropertyGet("Range", ("N"+IntToStr(n+1)).c_str()).OlePropertyGet("Interior").OlePropertySet("ColorIndex",40);
                  Variant locvalues12[] = {god,DM->qObnovlenie->FieldByName("mes")->AsInteger,13};
                  if (DM->qPrazdDni->Locate("god;mes;den", VarArrayOf(locvalues12, 2), SearchOptions << loCaseInsensitive)) Sh.OlePropertyGet("Range", ("O"+IntToStr(n+1)).c_str()).OlePropertyGet("Interior").OlePropertySet("ColorIndex",40);
                  Variant locvalues13[] = {god,DM->qObnovlenie->FieldByName("mes")->AsInteger,14};
                  if (DM->qPrazdDni->Locate("god;mes;den", VarArrayOf(locvalues13, 2), SearchOptions << loCaseInsensitive)) Sh.OlePropertyGet("Range", ("P"+IntToStr(n+1)).c_str()).OlePropertyGet("Interior").OlePropertySet("ColorIndex",40);
                  Variant locvalues14[] = {god,DM->qObnovlenie->FieldByName("mes")->AsInteger,15};
                  if (DM->qPrazdDni->Locate("god;mes;den", VarArrayOf(locvalues14, 2), SearchOptions << loCaseInsensitive)) Sh.OlePropertyGet("Range", ("Q"+IntToStr(n+1)).c_str()).OlePropertyGet("Interior").OlePropertySet("ColorIndex",40);
                  Variant locvalues15[] = {god,DM->qObnovlenie->FieldByName("mes")->AsInteger,16};
                  if (DM->qPrazdDni->Locate("god;mes;den", VarArrayOf(locvalues15, 2), SearchOptions << loCaseInsensitive)) Sh.OlePropertyGet("Range", ("R"+IntToStr(n+1)).c_str()).OlePropertyGet("Interior").OlePropertySet("ColorIndex",40);
                  Variant locvalues16[] = {god,DM->qObnovlenie->FieldByName("mes")->AsInteger,17};
                  if (DM->qPrazdDni->Locate("god;mes;den", VarArrayOf(locvalues16, 2), SearchOptions << loCaseInsensitive)) Sh.OlePropertyGet("Range", ("S"+IntToStr(n+1)).c_str()).OlePropertyGet("Interior").OlePropertySet("ColorIndex",40);
                  Variant locvalues17[] = {god,DM->qObnovlenie->FieldByName("mes")->AsInteger,18};
                  if (DM->qPrazdDni->Locate("god;mes;den", VarArrayOf(locvalues17, 2), SearchOptions << loCaseInsensitive)) Sh.OlePropertyGet("Range", ("T"+IntToStr(n+1)).c_str()).OlePropertyGet("Interior").OlePropertySet("ColorIndex",40);
                  Variant locvalues18[] = {god,DM->qObnovlenie->FieldByName("mes")->AsInteger,19};
                  if (DM->qPrazdDni->Locate("god;mes;den", VarArrayOf(locvalues18, 2), SearchOptions << loCaseInsensitive)) Sh.OlePropertyGet("Range", ("U"+IntToStr(n+1)).c_str()).OlePropertyGet("Interior").OlePropertySet("ColorIndex",40);
                  Variant locvalues19[] = {god,DM->qObnovlenie->FieldByName("mes")->AsInteger,20};
                  if (DM->qPrazdDni->Locate("god;mes;den", VarArrayOf(locvalues19, 2), SearchOptions << loCaseInsensitive)) Sh.OlePropertyGet("Range", ("V"+IntToStr(n+1)).c_str()).OlePropertyGet("Interior").OlePropertySet("ColorIndex",40);
                  Variant locvalues20[] = {god,DM->qObnovlenie->FieldByName("mes")->AsInteger,21};
                  if (DM->qPrazdDni->Locate("god;mes;den", VarArrayOf(locvalues20, 2), SearchOptions << loCaseInsensitive)) Sh.OlePropertyGet("Range", ("W"+IntToStr(n+1)).c_str()).OlePropertyGet("Interior").OlePropertySet("ColorIndex",40);
                  Variant locvalues21[] = {god,DM->qObnovlenie->FieldByName("mes")->AsInteger,22};
                  if (DM->qPrazdDni->Locate("god;mes;den", VarArrayOf(locvalues21, 2), SearchOptions << loCaseInsensitive)) Sh.OlePropertyGet("Range", ("X"+IntToStr(n+1)).c_str()).OlePropertyGet("Interior").OlePropertySet("ColorIndex",40);
                  Variant locvalues22[] = {god,DM->qObnovlenie->FieldByName("mes")->AsInteger,23};
                  if (DM->qPrazdDni->Locate("god;mes;den", VarArrayOf(locvalues22, 2), SearchOptions << loCaseInsensitive)) Sh.OlePropertyGet("Range", ("Y"+IntToStr(n+1)).c_str()).OlePropertyGet("Interior").OlePropertySet("ColorIndex",40);
                  Variant locvalues23[] = {god,DM->qObnovlenie->FieldByName("mes")->AsInteger,24};
                  if (DM->qPrazdDni->Locate("god;mes;den", VarArrayOf(locvalues23, 2), SearchOptions << loCaseInsensitive)) Sh.OlePropertyGet("Range", ("Z"+IntToStr(n+1)).c_str()).OlePropertyGet("Interior").OlePropertySet("ColorIndex",40);
                  Variant locvalues24[] = {god,DM->qObnovlenie->FieldByName("mes")->AsInteger,25};
                  if (DM->qPrazdDni->Locate("god;mes;den", VarArrayOf(locvalues24, 2), SearchOptions << loCaseInsensitive)) Sh.OlePropertyGet("Range", ("AA"+IntToStr(n+1)).c_str()).OlePropertyGet("Interior").OlePropertySet("ColorIndex",40);
                  Variant locvalues25[] = {god,DM->qObnovlenie->FieldByName("mes")->AsInteger,26};
                  if (DM->qPrazdDni->Locate("god;mes;den", VarArrayOf(locvalues25, 2), SearchOptions << loCaseInsensitive)) Sh.OlePropertyGet("Range", ("AB"+IntToStr(n+1)).c_str()).OlePropertyGet("Interior").OlePropertySet("ColorIndex",40);
                  Variant locvalues26[] = {god,DM->qObnovlenie->FieldByName("mes")->AsInteger,27};
                  if (DM->qPrazdDni->Locate("god;mes;den", VarArrayOf(locvalues26, 2), SearchOptions << loCaseInsensitive)) Sh.OlePropertyGet("Range", ("AC"+IntToStr(n+1)).c_str()).OlePropertyGet("Interior").OlePropertySet("ColorIndex",40);
                  Variant locvalues27[] = {god,DM->qObnovlenie->FieldByName("mes")->AsInteger,28};
                  if (DM->qPrazdDni->Locate("god;mes;den", VarArrayOf(locvalues27, 2), SearchOptions << loCaseInsensitive)) Sh.OlePropertyGet("Range", ("AD"+IntToStr(n+1)).c_str()).OlePropertyGet("Interior").OlePropertySet("ColorIndex",40);
                  Variant locvalues28[] = {god,DM->qObnovlenie->FieldByName("mes")->AsInteger,29};
                  if (DM->qPrazdDni->Locate("god;mes;den", VarArrayOf(locvalues28, 2), SearchOptions << loCaseInsensitive)) Sh.OlePropertyGet("Range", ("AE"+IntToStr(n+1)).c_str()).OlePropertyGet("Interior").OlePropertySet("ColorIndex",40);
                  Variant locvalues29[] = {god,DM->qObnovlenie->FieldByName("mes")->AsInteger,30};
                  if (DM->qPrazdDni->Locate("god;mes;den", VarArrayOf(locvalues29, 2), SearchOptions << loCaseInsensitive)) Sh.OlePropertyGet("Range", ("AF"+IntToStr(n+1)).c_str()).OlePropertyGet("Interior").OlePropertySet("ColorIndex",40);
                  Variant locvalues30[] = {god,DM->qObnovlenie->FieldByName("mes")->AsInteger,31};
                  if (DM->qPrazdDni->Locate("god;mes;den", VarArrayOf(locvalues30, 2), SearchOptions << loCaseInsensitive)) Sh.OlePropertyGet("Range", ("AG"+IntToStr(n+1)).c_str()).OlePropertyGet("Interior").OlePropertySet("ColorIndex",40);

                }
              else //для графиков начинающихся с 26 числа
                {
                  Variant locvalues[] = {god,DM->qObnovlenie->FieldByName("mes")->AsInteger,1};
                  if (DM->qPrazdDni->Locate("god;mes;den", VarArrayOf(locvalues, 2), SearchOptions << loCaseInsensitive)) Sh.OlePropertyGet("Range", ("I"+IntToStr(n+1)).c_str()).OlePropertyGet("Interior").OlePropertySet("ColorIndex",40);
                  Variant locvalues1[] = {god,DM->qObnovlenie->FieldByName("mes")->AsInteger,2};
                  if (DM->qPrazdDni->Locate("god;mes;den", VarArrayOf(locvalues1, 2), SearchOptions << loCaseInsensitive)) Sh.OlePropertyGet("Range", ("J"+IntToStr(n+1)).c_str()).OlePropertyGet("Interior").OlePropertySet("ColorIndex",40);
                  Variant locvalues2[] = {god,DM->qObnovlenie->FieldByName("mes")->AsInteger,3};
                  if (DM->qPrazdDni->Locate("god;mes;den", VarArrayOf(locvalues2, 2), SearchOptions << loCaseInsensitive)) Sh.OlePropertyGet("Range", ("K"+IntToStr(n+1)).c_str()).OlePropertyGet("Interior").OlePropertySet("ColorIndex",40);
                  Variant locvalues3[] = {god,DM->qObnovlenie->FieldByName("mes")->AsInteger,4};
                  if (DM->qPrazdDni->Locate("god;mes;den", VarArrayOf(locvalues3, 2), SearchOptions << loCaseInsensitive)) Sh.OlePropertyGet("Range", ("L"+IntToStr(n+1)).c_str()).OlePropertyGet("Interior").OlePropertySet("ColorIndex",40);
                  Variant locvalues4[] = {god,DM->qObnovlenie->FieldByName("mes")->AsInteger,5};
                  if (DM->qPrazdDni->Locate("god;mes;den", VarArrayOf(locvalues4, 2), SearchOptions << loCaseInsensitive)) Sh.OlePropertyGet("Range", ("M"+IntToStr(n+1)).c_str()).OlePropertyGet("Interior").OlePropertySet("ColorIndex",40);
                  Variant locvalues5[] = {god,DM->qObnovlenie->FieldByName("mes")->AsInteger,6};
                  if (DM->qPrazdDni->Locate("god;mes;den", VarArrayOf(locvalues5, 2), SearchOptions << loCaseInsensitive)) Sh.OlePropertyGet("Range", ("N"+IntToStr(n+1)).c_str()).OlePropertyGet("Interior").OlePropertySet("ColorIndex",40);
                  Variant locvalues6[] = {god,DM->qObnovlenie->FieldByName("mes")->AsInteger,7};
                  if (DM->qPrazdDni->Locate("god;mes;den", VarArrayOf(locvalues6, 2), SearchOptions << loCaseInsensitive)) Sh.OlePropertyGet("Range", ("O"+IntToStr(n+1)).c_str()).OlePropertyGet("Interior").OlePropertySet("ColorIndex",40);
                  Variant locvalues7[] = {god,DM->qObnovlenie->FieldByName("mes")->AsInteger,8};
                  if (DM->qPrazdDni->Locate("god;mes;den", VarArrayOf(locvalues7, 2), SearchOptions << loCaseInsensitive)) Sh.OlePropertyGet("Range", ("P"+IntToStr(n+1)).c_str()).OlePropertyGet("Interior").OlePropertySet("ColorIndex",40);
                  Variant locvalues8[] = {god,DM->qObnovlenie->FieldByName("mes")->AsInteger,9};
                  if (DM->qPrazdDni->Locate("god;mes;den", VarArrayOf(locvalues8, 2), SearchOptions << loCaseInsensitive)) Sh.OlePropertyGet("Range", ("Q"+IntToStr(n+1)).c_str()).OlePropertyGet("Interior").OlePropertySet("ColorIndex",40);
                  Variant locvalues9[] = {god,DM->qObnovlenie->FieldByName("mes")->AsInteger,10};
                  if (DM->qPrazdDni->Locate("god;mes;den", VarArrayOf(locvalues9, 2), SearchOptions << loCaseInsensitive)) Sh.OlePropertyGet("Range", ("R"+IntToStr(n+1)).c_str()).OlePropertyGet("Interior").OlePropertySet("ColorIndex",40);
                  Variant locvalues10[] = {god,DM->qObnovlenie->FieldByName("mes")->AsInteger,11};
                  if (DM->qPrazdDni->Locate("god;mes;den", VarArrayOf(locvalues10, 2), SearchOptions << loCaseInsensitive)) Sh.OlePropertyGet("Range", ("S"+IntToStr(n+1)).c_str()).OlePropertyGet("Interior").OlePropertySet("ColorIndex",40);
                  Variant locvalues11[] = {god,DM->qObnovlenie->FieldByName("mes")->AsInteger,12};
                  if (DM->qPrazdDni->Locate("god;mes;den", VarArrayOf(locvalues11, 2), SearchOptions << loCaseInsensitive)) Sh.OlePropertyGet("Range", ("T"+IntToStr(n+1)).c_str()).OlePropertyGet("Interior").OlePropertySet("ColorIndex",40);
                  Variant locvalues12[] = {god,DM->qObnovlenie->FieldByName("mes")->AsInteger,13};
                  if (DM->qPrazdDni->Locate("god;mes;den", VarArrayOf(locvalues12, 2), SearchOptions << loCaseInsensitive)) Sh.OlePropertyGet("Range", ("U"+IntToStr(n+1)).c_str()).OlePropertyGet("Interior").OlePropertySet("ColorIndex",40);
                  Variant locvalues13[] = {god,DM->qObnovlenie->FieldByName("mes")->AsInteger,14};
                  if (DM->qPrazdDni->Locate("god;mes;den", VarArrayOf(locvalues13, 2), SearchOptions << loCaseInsensitive)) Sh.OlePropertyGet("Range", ("V"+IntToStr(n+1)).c_str()).OlePropertyGet("Interior").OlePropertySet("ColorIndex",40);
                  Variant locvalues14[] = {god,DM->qObnovlenie->FieldByName("mes")->AsInteger,15};
                  if (DM->qPrazdDni->Locate("god;mes;den", VarArrayOf(locvalues14, 2), SearchOptions << loCaseInsensitive)) Sh.OlePropertyGet("Range", ("W"+IntToStr(n+1)).c_str()).OlePropertyGet("Interior").OlePropertySet("ColorIndex",40);
                  Variant locvalues15[] = {god,DM->qObnovlenie->FieldByName("mes")->AsInteger,16};
                  if (DM->qPrazdDni->Locate("god;mes;den", VarArrayOf(locvalues15, 2), SearchOptions << loCaseInsensitive)) Sh.OlePropertyGet("Range", ("X"+IntToStr(n+1)).c_str()).OlePropertyGet("Interior").OlePropertySet("ColorIndex",40);
                  Variant locvalues16[] = {god,DM->qObnovlenie->FieldByName("mes")->AsInteger,17};
                  if (DM->qPrazdDni->Locate("god;mes;den", VarArrayOf(locvalues16, 2), SearchOptions << loCaseInsensitive)) Sh.OlePropertyGet("Range", ("Y"+IntToStr(n+1)).c_str()).OlePropertyGet("Interior").OlePropertySet("ColorIndex",40);
                  Variant locvalues17[] = {god,DM->qObnovlenie->FieldByName("mes")->AsInteger,18};
                  if (DM->qPrazdDni->Locate("god;mes;den", VarArrayOf(locvalues17, 2), SearchOptions << loCaseInsensitive)) Sh.OlePropertyGet("Range", ("Z"+IntToStr(n+1)).c_str()).OlePropertyGet("Interior").OlePropertySet("ColorIndex",40);
                  Variant locvalues18[] = {god,DM->qObnovlenie->FieldByName("mes")->AsInteger,19};
                  if (DM->qPrazdDni->Locate("god;mes;den", VarArrayOf(locvalues18, 2), SearchOptions << loCaseInsensitive)) Sh.OlePropertyGet("Range", ("AA"+IntToStr(n+1)).c_str()).OlePropertyGet("Interior").OlePropertySet("ColorIndex",40);
                  Variant locvalues19[] = {god,DM->qObnovlenie->FieldByName("mes")->AsInteger,20};
                  if (DM->qPrazdDni->Locate("god;mes;den", VarArrayOf(locvalues19, 2), SearchOptions << loCaseInsensitive)) Sh.OlePropertyGet("Range", ("AB"+IntToStr(n+1)).c_str()).OlePropertyGet("Interior").OlePropertySet("ColorIndex",40);
                  Variant locvalues20[] = {god,DM->qObnovlenie->FieldByName("mes")->AsInteger,21};
                  if (DM->qPrazdDni->Locate("god;mes;den", VarArrayOf(locvalues20, 2), SearchOptions << loCaseInsensitive)) Sh.OlePropertyGet("Range", ("AC"+IntToStr(n+1)).c_str()).OlePropertyGet("Interior").OlePropertySet("ColorIndex",40);
                  Variant locvalues21[] = {god,DM->qObnovlenie->FieldByName("mes")->AsInteger,22};
                  if (DM->qPrazdDni->Locate("god;mes;den", VarArrayOf(locvalues21, 2), SearchOptions << loCaseInsensitive)) Sh.OlePropertyGet("Range", ("AD"+IntToStr(n+1)).c_str()).OlePropertyGet("Interior").OlePropertySet("ColorIndex",40);
                  Variant locvalues22[] = {god,DM->qObnovlenie->FieldByName("mes")->AsInteger,23};
                  if (DM->qPrazdDni->Locate("god;mes;den", VarArrayOf(locvalues22, 2), SearchOptions << loCaseInsensitive)) Sh.OlePropertyGet("Range", ("AE"+IntToStr(n+1)).c_str()).OlePropertyGet("Interior").OlePropertySet("ColorIndex",40);
                  Variant locvalues23[] = {god,DM->qObnovlenie->FieldByName("mes")->AsInteger,24};
                  if (DM->qPrazdDni->Locate("god;mes;den", VarArrayOf(locvalues23, 2), SearchOptions << loCaseInsensitive)) Sh.OlePropertyGet("Range", ("AF"+IntToStr(n+1)).c_str()).OlePropertyGet("Interior").OlePropertySet("ColorIndex",40);
                  Variant locvalues24[] = {god,DM->qObnovlenie->FieldByName("mes")->AsInteger,25};
                  if (DM->qPrazdDni->Locate("god;mes;den", VarArrayOf(locvalues24, 2), SearchOptions << loCaseInsensitive)) Sh.OlePropertyGet("Range", ("AG"+IntToStr(n+1)).c_str()).OlePropertyGet("Interior").OlePropertySet("ColorIndex",40);
                  Variant locvalues25[] = {god,DM->qObnovlenie->FieldByName("mes")->AsInteger,26};
                  if (DM->qPrazdDni->Locate("god;mes;den", VarArrayOf(locvalues25, 2), SearchOptions << loCaseInsensitive)) Sh.OlePropertyGet("Range", ("C"+IntToStr(n+1)).c_str()).OlePropertyGet("Interior").OlePropertySet("ColorIndex",40);
                  Variant locvalues26[] = {god,DM->qObnovlenie->FieldByName("mes")->AsInteger,27};
                  if (DM->qPrazdDni->Locate("god;mes;den", VarArrayOf(locvalues26, 2), SearchOptions << loCaseInsensitive)) Sh.OlePropertyGet("Range", ("D"+IntToStr(n+1)).c_str()).OlePropertyGet("Interior").OlePropertySet("ColorIndex",40);
                  Variant locvalues27[] = {god,DM->qObnovlenie->FieldByName("mes")->AsInteger,28};
                  if (DM->qPrazdDni->Locate("god;mes;den", VarArrayOf(locvalues27, 2), SearchOptions << loCaseInsensitive)) Sh.OlePropertyGet("Range", ("E"+IntToStr(n+1)).c_str()).OlePropertyGet("Interior").OlePropertySet("ColorIndex",40);
                  Variant locvalues28[] = {god,DM->qObnovlenie->FieldByName("mes")->AsInteger,29};
                  if (DM->qPrazdDni->Locate("god;mes;den", VarArrayOf(locvalues28, 2), SearchOptions << loCaseInsensitive)) Sh.OlePropertyGet("Range", ("F"+IntToStr(n+1)).c_str()).OlePropertyGet("Interior").OlePropertySet("ColorIndex",40);
                  Variant locvalues29[] = {god,DM->qObnovlenie->FieldByName("mes")->AsInteger,30};
                  if (DM->qPrazdDni->Locate("god;mes;den", VarArrayOf(locvalues29, 2), SearchOptions << loCaseInsensitive)) Sh.OlePropertyGet("Range", ("G"+IntToStr(n+1)).c_str()).OlePropertyGet("Interior").OlePropertySet("ColorIndex",40);
                  Variant locvalues30[] = {god,DM->qObnovlenie->FieldByName("mes")->AsInteger,31};
                  if (DM->qPrazdDni->Locate("god;mes;den", VarArrayOf(locvalues30, 2), SearchOptions << loCaseInsensitive)) Sh.OlePropertyGet("Range", ("H"+IntToStr(n+1)).c_str()).OlePropertyGet("Interior").OlePropertySet("ColorIndex",40);
                }





           if (br>=DM->qGrafik->FieldByName("br")->AsInteger)
            {
              br=1;

            }
          else br++;

          i++;
          num++;
          n++;
         
          DM->qObnovlenie->Next();
          ProgressBar->Position++;
          mes = DM->qObnovlenie->FieldByName("mes")->AsInteger;
           //рисуем сетку
          Sh.OlePropertyGet("Range",("A"+IntToStr(n)+":AN"+IntToStr(n)).c_str()).OlePropertyGet("Borders").OlePropertySet("LineStyle", 1);

        }

      i=0;
      int j=0;
      num=9;
      int z;


      //Вывод описания начала и окончания смен
      AnsiString smena = DM->qObnovlenie->FieldByName("smena")->AsString;
      while (smena.Pos(',')>0)
        {
          toExcel(Sh,"smena",i, smena.SubString(1, smena.Pos(',')-1));
          smena = Trim(smena.Delete(1,smena.Pos(',')));
          i++;
        }
      toExcel(Sh,"smena",i, smena);

      i=0;
 
      //итоговые суммы
      switch(DM->qOgraf->FieldByName("br")->AsInteger)
        {
          case 1:
                  z=1*2;  //по 2 строки на 1 бригаду + 1 строка с заголовком месяца
          break;

          case 2:
                  z=2*2;
          break;

          case 3:
                  z=3*2;
          break;

          case 4:
                  z=4*2;
          break;

          case 5:
                  z=5*2;
          break;

          case 6:
                  z=6*2;
          break;
        }

      while (j<DM->qGrafik->FieldByName("br")->AsInteger)
        {



          Sh.OlePropertyGet("Range", "ksm1").OlePropertyGet("Offset", j).OlePropertySet("Formula", ("=СУММ(AH"+IntToStr(i+9)+"+AH"+IntToStr(num+1*z+i)+"+AH"+IntToStr(num+2*z+i)+"+AH"+IntToStr(num+3*z+i)+"\
                                                                                                          +AH"+IntToStr(num+4*z+i)+"+AH"+IntToStr(num+5*z+i)+"+AH"+IntToStr(num+6*z+i)+" \
                                                                                                          +AH"+IntToStr(num+7*z+i)+"+AH"+IntToStr(num+8*z+i)+"+AH"+IntToStr(num+9*z+i)+"\
                                                                                                          +Ah"+IntToStr(num+10*z+i)+"+AH"+IntToStr(num+11*z+i)+")").c_str());

          Sh.OlePropertyGet("Range", "ochf").OlePropertyGet("Offset", j).OlePropertySet("Formula", ("=СУММ(AI"+IntToStr(i+9)+"+AI"+IntToStr(num+1*z+i)+"+AI"+IntToStr(num+2*z+i)+"+AI"+IntToStr(num+3*z+i)+"\
                                                                                                          +AI"+IntToStr(num+4*z+i)+"+AI"+IntToStr(num+5*z+i)+"+AI"+IntToStr(num+6*z+i)+" \
                                                                                                          +AI"+IntToStr(num+7*z+i)+"+AI"+IntToStr(num+8*z+i)+"+AI"+IntToStr(num+9*z+i)+"\
                                                                                                          +AI"+IntToStr(num+10*z+i)+"+AI"+IntToStr(num+11*z+i)+")").c_str());
          Sh.OlePropertyGet("Range", "onorma").OlePropertyGet("Offset", j).OlePropertySet("Formula", ("=СУММ(AJ"+IntToStr(i+9)+"+AJ"+IntToStr(num+z*1+i)+"+AJ"+IntToStr(num+z*2+i)+"+AJ"+IntToStr(num+z*3+i)+"\
                                                                                                          +AJ"+IntToStr(num+z*4+i)+"+AJ"+IntToStr(num+z*5+i)+"+AJ"+IntToStr(num+z*6+i)+" \
                                                                                                          +AJ"+IntToStr(num+z*7+i)+"+AJ"+IntToStr(num+z*8+i)+"+AJ"+IntToStr(num+z*9+i)+"\
                                                                                                          +AJ"+IntToStr(num+z*10+i)+"+AJ"+IntToStr(num+z*11+i)+")").c_str());
          Sh.OlePropertyGet("Range", "ovch").OlePropertyGet("Offset", j).OlePropertySet("Formula", ("=СУММ(AK"+IntToStr(i+9)+"+AK"+IntToStr(num+z*1+i)+"+AK"+IntToStr(num+z*2+i)+"+AK"+IntToStr(num+z*3+i)+"\
                                                                                                          +AK"+IntToStr(num+z*4+i)+"+AK"+IntToStr(num+z*5+i)+"+AK"+IntToStr(num+z*6+i)+" \
                                                                                                          +AK"+IntToStr(num+z*7+i)+"+AK"+IntToStr(num+z*8+i)+"+AK"+IntToStr(num+z*9+i)+"\
                                                                                                          +AK"+IntToStr(num+z*10+i)+"+AK"+IntToStr(num+z*11+i)+")").c_str());
          Sh.OlePropertyGet("Range", "onch").OlePropertyGet("Offset", j).OlePropertySet("Formula", ("=СУММ(AL"+IntToStr(i+9)+"+AL"+IntToStr(num+z*1+i)+"+AL"+IntToStr(num+z*2+i)+"+AL"+IntToStr(num+z*3+i)+"\
                                                                                                          +AL"+IntToStr(num+z*4+i)+"+AL"+IntToStr(num+z*5+i)+"+AL"+IntToStr(num+z*6+i)+" \
                                                                                                          +AL"+IntToStr(num+z*7+i)+"+AL"+IntToStr(num+z*8+i)+"+AL"+IntToStr(num+z*9+i)+"\
                                                                                                          +AL"+IntToStr(num+z*10+i)+"+AL"+IntToStr(num+z*11+i)+")").c_str());
          Sh.OlePropertyGet("Range", "opch").OlePropertyGet("Offset", j).OlePropertySet("Formula", ("=СУММ(AM"+IntToStr(i+9)+"+AM"+IntToStr(num+z*1+i)+"+AM"+IntToStr(num+z*2+i)+"+AM"+IntToStr(num+z*3+i)+"\
                                                                                                          +AM"+IntToStr(num+z*4+i)+"+AM"+IntToStr(num+z*5+i)+"+AM"+IntToStr(num+z*6+i)+" \
                                                                                                          +AM"+IntToStr(num+z*7+i)+"+AM"+IntToStr(num+z*8+i)+"+AM"+IntToStr(num+z*9+i)+"\
                                                                                                          +AM"+IntToStr(num+z*10+i)+"+AM"+IntToStr(num+z*11+i)+")").c_str());
          Sh.OlePropertyGet("Range", "opgraf").OlePropertyGet("Offset", j).OlePropertySet("Formula", ("=СУММ(AN"+IntToStr(i+9)+"+AN"+IntToStr(num+z*1+i)+"+AN"+IntToStr(num+z*2+i)+"+AN"+IntToStr(num+z*3+i)+"\
                                                                                                          +AN"+IntToStr(num+z*4+i)+"+AN"+IntToStr(num+z*5+i)+"+AN"+IntToStr(num+z*6+i)+" \
                                                                                                          +AN"+IntToStr(num+z*7+i)+"+AN"+IntToStr(num+z*8+i)+"+AN"+IntToStr(num+z*9+i)+"\
                                                                                                          +AN"+IntToStr(num+z*10+i)+"+AN"+IntToStr(num+z*11+i)+")").c_str());
           i=i+2;
           j++;

        }



      //определение ячейки с которой начинается заполнение
      num = Sh.OlePropertyGet("Range", "ksm1").OlePropertyGet("Offset", j).OlePropertyGet("Row");

      // среднее значение
      toExcel(Sh,"sr_chf",j, "Среднегодовые часы:");
      Sh.OlePropertyGet("Range", "sr_chf").OlePropertyGet("Offset", j).OlePropertyGet("Font").OlePropertySet("Bold",true);
      Sh.OlePropertyGet("Range", "sr_chf").OlePropertyGet("Offset", j).OlePropertyGet("Font").OlePropertySet("Size",13);
      Sh.OlePropertyGet("Range", "sr_chf").OlePropertyGet("Offset", j).OlePropertySet("HorizontalAlignment", xlHAlignRight);

      Sh.OlePropertyGet("Range", "schf").OlePropertyGet("Offset", j).OlePropertyGet("Font").OlePropertySet("Bold",true);
      Sh.OlePropertyGet("Range", "snorma").OlePropertyGet("Offset", j).OlePropertyGet("Font").OlePropertySet("Bold",true);
      Sh.OlePropertyGet("Range", "svch").OlePropertyGet("Offset", j).OlePropertyGet("Font").OlePropertySet("Bold",true);
      Sh.OlePropertyGet("Range", "snch").OlePropertyGet("Offset", j).OlePropertyGet("Font").OlePropertySet("Bold",true);
      Sh.OlePropertyGet("Range", "spch").OlePropertyGet("Offset", j).OlePropertyGet("Font").OlePropertySet("Bold",true);
      Sh.OlePropertyGet("Range", "spgraf").OlePropertyGet("Offset", j).OlePropertyGet("Font").OlePropertySet("Bold",true);

      Sh.OlePropertyGet("Range", "schf").OlePropertyGet("Offset", j).OlePropertySet("Formula", ("=СУММ(AI"+IntToStr(num-DM->qGrafik->FieldByName("br")->AsInteger)+":AI"+IntToStr(num-1)+")/12/"+DM->qGrafik->FieldByName("br")->AsInteger).c_str());
      Sh.OlePropertyGet("Range", "snorma").OlePropertyGet("Offset", j).OlePropertySet("Formula", ("=СУММ(AJ"+IntToStr(num-DM->qGrafik->FieldByName("br")->AsInteger)+":AJ"+IntToStr(num-1)+")/12/"+DM->qGrafik->FieldByName("br")->AsInteger).c_str());
      Sh.OlePropertyGet("Range", "svch").OlePropertyGet("Offset", j).OlePropertySet("Formula", ("=СУММ(AK"+IntToStr(num-DM->qGrafik->FieldByName("br")->AsInteger)+":AK"+IntToStr(num-1)+")/12/"+DM->qGrafik->FieldByName("br")->AsInteger).c_str());
      Sh.OlePropertyGet("Range", "snch").OlePropertyGet("Offset", j).OlePropertySet("Formula", ("=СУММ(AL"+IntToStr(num-DM->qGrafik->FieldByName("br")->AsInteger)+":AL"+IntToStr(num-1)+")/12/"+DM->qGrafik->FieldByName("br")->AsInteger).c_str());
      Sh.OlePropertyGet("Range", "spch").OlePropertyGet("Offset", j).OlePropertySet("Formula", ("=СУММ(AM"+IntToStr(num-DM->qGrafik->FieldByName("br")->AsInteger)+":AM"+IntToStr(num-1)+")/12/"+DM->qGrafik->FieldByName("br")->AsInteger).c_str());
      Sh.OlePropertyGet("Range", "spgraf").OlePropertyGet("Offset", j).OlePropertySet("Formula", ("=СУММ(AN"+IntToStr(num-DM->qGrafik->FieldByName("br")->AsInteger)+":AN"+IntToStr(num-1)+")/12/"+DM->qGrafik->FieldByName("br")->AsInteger).c_str());





      //Отключить вывод сообщений с вопросами типа "Заменить файл..."
      AppEx.OlePropertySet("DisplayAlerts",false);


      //Создание папки, если ее не существует
      ForceDirectories(Main->WorkPath);

      //Сохранить книгу в папке в файле по указанию
      AnsiString vAsCurDir1=WorkPath+"\\"+ComboBox1->Text+" График.xls";
      //AppEx.OlePropertyGet("Workbooks").OlePropertyGet("Item",1).OleProcedure("SaveAs",vAsCurDir1.c_str());
      Sh.OleProcedure("SaveAs",vAsCurDir1.c_str());

      //Закрыть открытое приложение Excel
     // AppEx.OleProcedure("Quit");
      //AppEx.OlePropertySet("Visible",true);
     // AppEx.OlePropertyGet("WorkBooks").OleProcedure("Open",vAsCurDir1.c_str());
    // AppEx.OlePropertyGet("WorkBooks").OleProcedure("Open",sFile.c_str());
      AppEx.OlePropertySet("Visible",true);
      Main->Cursor=crDefault;
      ProgressBar->Visible = false;
      StatusBar1->SimpleText = "Отчетный период: "+ IntToStr(god);
    }
  else
    {
      Application->MessageBox("Нет данных по графику!!!", "", MB_OK + MB_ICONINFORMATION);
      Abort();
    }
}
//---------------------------------------------------------------------------


void __fastcall TMain::N7Click(TObject *Sender)
{
  Sprav->Panel2->Visible = false;


  DM->qSprav->Close();
  DM->qSprav->Parameters->ParamByName("pgod")->Value = god;
  DM->qSprav->Parameters->ParamByName("pgod1")->Value = god+1;
  try
    {
      DM->qSprav->Open();
    }
  catch(...)
    {
      Application->MessageBox("Невозможно получить данные из справочника праздничных дней (SP_PRD)","Ошибка",
                              MB_OK + MB_ICONERROR);
      Abort();
    }
  Sprav->ShowModal();
}
//---------------------------------------------------------------------------

//Просмотр графиков на текущий год
void __fastcall TMain::N8Click(TObject *Sender)
{
  AnsiString Sql;
  Word year, month, day;

  //Считывание отчетного года из grafr
 /* DM->qObnovlenie->Close();
  DM->qObnovlenie->SQL->Clear();
  DM->qObnovlenie->SQL->Add("select * from grafr");
  try
    {
      DM->qObnovlenie->Open();
    }
  catch(...)
    {
      Application->MessageBox("Невозможно считать отчетный период","Ошибка",
                              MB_OK + MB_ICONERROR);
      Application->Terminate();
      Abort();
    }*/

  
  DecodeDate(Date(), year, month, day);

  god = StrToInt(year);
  //god = DM->qObnovlenie->FieldByName("god")->AsInteger;
  //DM->qGrafik->Close();
  DBGridEh1->Enabled = false;
  ComboBox1->ItemIndex = -1;
  StatusBar1->SimpleText="Отчетный период:  "+IntToStr(god)+" год";

  //Праздничные дни
  DM->qPrazdDni->Close();
  DM->qPrazdDni->Parameters->ParamByName("pgod")->Value = god;
  try
    {
      DM->qPrazdDni->Open();
    }
  catch(...)
    {
      Application->MessageBox("Возникла ошибка при обращении к справочнику праздничных дней","Ошибка",
                              MB_OK+MB_ICONERROR);
      Abort();
    }

  //Предпраздничные дни
  DM->qPrdPrazdDni->Close();
  DM->qPrdPrazdDni->Parameters->ParamByName("pgod")->Value = god;
  try
    {
      DM->qPrdPrazdDni->Open();
    }
  catch(...)
    {
      Application->MessageBox("Возникла ошибка при обращении к справочнику праздничных дней","Ошибка",
                              MB_OK+MB_ICONERROR);
      Abort();
    }


  //Определение даты перехода на летнее/зимнее время
  TDateTime data;
 // Word year, month, day;

  // дата в марте
  data = DateToStr(EncodeDateMonthWeek(god,3,4,6));
  DecodeDate(data, year, month, day);
  day_mart = day;
  //для 40 и 90 графика, первой смены, дата в марте
  if (day_mart==31)
    {
      mes_mart2=4;
      day_mart2=1;
    }
  else
    {
      mes_mart2=3;
      day_mart2=day_mart+1;
    }

  //дата в октябре
  data = DateToStr(EncodeDateMonthWeek(god,10,4,6));
  DecodeDate(data, year, month, day);
  day_oktyabr = day;
  //для 40 и 90 графика, первой смены, дата в октябре
  if (day_oktyabr==31)
    {
      mes_oktyabr2=11;
      day_oktyabr2=1;
    }
  else
    {
      mes_oktyabr2=10;
      day_oktyabr2=day_oktyabr+1;
    }

   // Вывод в ComboBox выбираемых графиков
   DM->qObnovlenie2->Close();
   DM->qObnovlenie2->SQL->Clear();
   Sql="select distinct ograf \
        from spograf \
        where ograf not in (select ograf \
                            from (select ograf, mes  \
                                  from spgrafiki \
                                  where god="+IntToStr(god)+" group by ograf, mes) \
                            group  by ograf having count(*)=1) order by ograf";
   DM->qObnovlenie2->SQL->Add(Sql);

   try
     {
       DM->qObnovlenie2->Open();
     }
   catch(...)
     {
       Application->MessageBox("Ошибка доступа к таблице графиков (SPOGRAF)","Ошибка доступа",
                                MB_OK + MB_ICONERROR);
       Application->Terminate();
       Abort();
     }

   ComboBox1->Items->Clear();
   while(!DM->qObnovlenie2->Eof)
     {
       ComboBox1->Items->Add(DM->qObnovlenie2->FieldByName("ograf")->AsString);
       DM->qObnovlenie2->Next();
     }
   ComboBox1->ItemIndex = -1;


  Application->MessageBox(("Отчетный период изменен!!!\nОтображение графиков на "+IntToStr(god)+" год").c_str(),"Графики работы", MB_OK+MB_ICONINFORMATION);

}
//---------------------------------------------------------------------------

//Просмотр графиков на следующий год
void __fastcall TMain::N9Click(TObject *Sender)
{
  AnsiString Sql;
  Word year, month, day;

  //Считывание отчетного года из grafr
/*  DM->qObnovlenie->Close();
  DM->qObnovlenie->SQL->Clear();
  DM->qObnovlenie->SQL->Add("select * from grafr");
  try
    {
      DM->qObnovlenie->Open();
    }
  catch(...)
    {
      Application->MessageBox("Невозможно считать отчетный период","Ошибка",
                              MB_OK + MB_ICONERROR);
      Application->Terminate();
      Abort();
    } */

 
  DecodeDate(Date(), year, month, day);
  god = StrToInt(year)+1;


  //god = DM->qObnovlenie->FieldByName("god")->AsInteger+1;
  //DM->qGrafik->Close();
  DBGridEh1->Enabled = false;
  ComboBox1->ItemIndex = -1;
  StatusBar1->SimpleText="Отчетный период:  "+IntToStr(god)+" год";

  //Праздничные дни
  DM->qPrazdDni->Close();
  DM->qPrazdDni->Parameters->ParamByName("pgod")->Value = god;
  try
    {
      DM->qPrazdDni->Open();
    }
  catch(...)
    {
      Application->MessageBox("Возникла ошибка при обращении к справочнику праздничных дней","Ошибка",
                              MB_OK+MB_ICONERROR);
      Abort();
    }

  //Предпраздничные дни
  DM->qPrdPrazdDni->Close();
  DM->qPrdPrazdDni->Parameters->ParamByName("pgod")->Value = god;
  try
    {
      DM->qPrdPrazdDni->Open();
    }
  catch(...)
    {
      Application->MessageBox("Возникла ошибка при обращении к справочнику праздничных дней","Ошибка",
                              MB_OK+MB_ICONERROR);
      Abort();
    }

  //Определение даты перехода на летнее/зимнее время
  TDateTime data;
//  Word year, month, day;

  // дата в марте
  data = DateToStr(EncodeDateMonthWeek(god,3,4,6));
  DecodeDate(data, year, month, day);
  day_mart = day;
  //для 40 и 90 графика, первой смены, дата в марте
  if (day_mart==31)
    {
      mes_mart2=4;
      day_mart2=1;
    }
  else
    {
      mes_mart2=3;
      day_mart2=day_mart+1;
    }

  //дата в октябре
  data = DateToStr(EncodeDateMonthWeek(god,10,4,6));
  DecodeDate(data, year, month, day);
  day_oktyabr = day;
  //для 40 и 90 графика, первой смены, дата в октябре
  if (day_oktyabr==31)
    {
      mes_oktyabr2=11;
      day_oktyabr2=1;
    }
  else
    {
      mes_oktyabr2=10;
      day_oktyabr2=day_oktyabr+1;
    }

  // Вывод в ComboBox выбираемых графиков
   DM->qObnovlenie2->Close();
   DM->qObnovlenie2->SQL->Clear();
   Sql="select distinct ograf \
        from spograf \
        where ograf not in (select ograf \
                            from (select ograf, mes  \
                                  from spgrafiki \
                                  where god="+IntToStr(god)+" group by ograf, mes) \
                            group  by ograf having count(*)=1) order by ograf";
   DM->qObnovlenie2->SQL->Add(Sql);
   
   try
     {
       DM->qObnovlenie2->Open();
     }
   catch(...)
     {
       Application->MessageBox("Ошибка доступа к таблице графиков (SPOGRAF)","Ошибка доступа",
                                MB_OK + MB_ICONERROR);
       Application->Terminate();
       Abort();
     }

   ComboBox1->Items->Clear();

   while(!DM->qObnovlenie2->Eof)
     {
       ComboBox1->Items->Add(DM->qObnovlenie2->FieldByName("ograf")->AsString);
       DM->qObnovlenie2->Next();
     }
   ComboBox1->ItemIndex = -1;

  Application->MessageBox(("Отчетный период изменен!!!\nОтображение графиков на "+IntToStr(god)+" год").c_str(),"Графики работы", MB_OK+MB_ICONINFORMATION);
}
//---------------------------------------------------------------------------


void __fastcall TMain::N3Click(TObject *Sender)
{
  Main->Close();        
}
//---------------------------------------------------------------------------

// Следующий месяц
void __fastcall TMain::NextMonth(int &Month, int &Year, int k)
{
  for (int i=1; i<=k; i++) {
    if (Month>11) { Month = 1; Year++; }
    else Month++;
  }
}
//---------------------------------------------------------------------------

// Предыдущий месяц
void __fastcall TMain::PrevMonth(int &Month, int &Year, int k)
{
  for (int i=1; i<=k; i++) {
    if (Month==1) { Month = 12; Year--; }
    else Month--;
  }
}
//---------------------------------------------------------------------------



void __fastcall TMain::StatusBar1DblClick(TObject *Sender)
{
  Data->ShowModal();
}
//---------------------------------------------------------------------------

void __fastcall TMain::N5Click(TObject *Sender)
{
  Data->ShowModal();
}
//---------------------------------------------------------------------------

void __fastcall TMain::FormClose(TObject *Sender, TCloseAction &Action)
{
 //Очистка массива
      for(mes=0; mes<=12; mes++)
        {
          for(den=0; den<=31; den++)
            {
              chf[mes][den] = NULL;
              vihod[mes][den] = NULL;
              vchf[mes][den] = NULL;
              pchf[mes][den] = NULL;
              nchf[mes][den] = NULL;
            }
          ochf[mes]=0;
          ovchf[mes]=0;
          onchf[mes]=0;
          opchf[mes]=0;
          pgraf[mes]=0;
          chf0[mes]=0;
          nchf0[mes]=0;
          pchf0[mes]=0;
       }

  //очистка массива со списком графиков
  for (int i=0; i<149; i++)
    {
      //n_grafik[kol_grafik]=NULL;
      n_grafik[i]=NULL;
    }
}
//---------------------------------------------------------------------------

