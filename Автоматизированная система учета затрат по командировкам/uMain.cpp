//---------------------------------------------------------------------------
#define NO_WIN32_LEAN_AND_MEAN
//#include <stdio.h>

#include <vcl.h>
#pragma hdrstop

#include "uMain.h"
#include "uDM.h"
#include "uVvod.h"
#include "uGostinica.h"
#include "uSprav.h"
#include "FuncUser.h"
//---------------------------------------------------------------------------
#pragma package(smart_init)
#pragma link "DBGridEh"
#pragma resource "*.dfm"
TMain *Main;

const AnsiString Mes[] = {"Январь","Февраль","Март","Апрель","Май","Июнь","Июль","Август","Сентябрь","Октябрь","Ноябрь","Декабрь",""};
//---------------------------------------------------------------------------
__fastcall TMain::TMain(TComponent* Owner)
        : TForm(Owner)
{
}
//---------------------------------------------------------------------------
void __fastcall TMain::FormCreate(TObject *Sender)
{
  // Получение данных о пользователе из домена
  TStringList *SL_Groups = new TStringList();

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
  if ((SL_Groups->IndexOf("mmk-itsvc-hkom-admin")<=-1) && (SL_Groups->IndexOf("mmk-itsvc-hkom")<=-1))
    {
      MessageBox(Handle,"У вас нет прав для работы с\nпрограммой 'Учет командировок'!!!","Права доступа",8208);
      Application->Terminate();
      Abort();
    }

 /* //проверка прав
  //если права группы БОТ
  if (SL_Groups->IndexOf("mmk-itsvc-hrsr-ukil")>-1)
    {
      PageControl1->Visible = false;   //Панель с суммами
      N10RaschetMenu->Visible = false; //Пункт главного меню расчет
      NRschet->Visible = false;        //Предварительный расчет
      NPrint->Visible = true;          //Признак печати
      NReestr->Visible = true;         //Отчет "Реестр листков нетрудоспособности"
      NProtokol->Visible = true;       //Отчет "Выписка из заседания комиссии соц. страхования УИТ"
      NOtchetRaschet1->Visible = false; //Отчет "Расчет по больничным листам"
      NOtchetRaschet2->Visible = false; //Отчет "Расчет по больничным листам по травмам"
      N10Doplata->Visible = false;     //Отображение данных -> только доплаты
      N10Doplati->Visible = false;     //Отчет по доплатам
      NPerenosNaSledMes->Visible = false; //Перенос больничного листа на следующий месяц
      Nepodtvergd_zapN10->Visible = false; //Неподтвержденные к оплате записи
      N18->Visible = false;            //Порядок ввода данных
      Prava=1;
    }
  else
    {
      Application->MessageBox("Не установлены права доступа(УКИЛ, ОУЗП) для работы с программой АСПД 'Средний заработок'!!!","Права доступа",
                              MB_OK+MB_ICONERROR);
      Application->Terminate();
      Abort();

 DecimalSeparator = '.';

  //получение года из grafr
  DM->qObnovlenie->Close();
  DM->qObnovlenie->SQL->Clear();
  DM->qObnovlenie->SQL->Add("select * from grafr");
  try
    {
      DM->qObnovlenie->Open();
    }
  catch(...)
    {
      Application->MessageBox("Нет доступа к таблице GRAFR.\nНевозможно получить отчетный период","Отчетный период",
                              MB_OK+ MB_ICONERROR);
      //Application->Terminate();
    }


  god = DM->qObnovlenie->FieldByName("god")->AsInteger+1;

     //Фильтрация автоматическая без нажатия Enter
  DBGridEh1->Style->FilterEditCloseUpApplyFilter =true;
    */


  StatusBar1->SimplePanel = true;
  //StatusBar1->SimpleText = "Отчетный период: "+ Mes[mm-1] +" "+yyyy;
  StatusBar1->SimpleText ="";


 // Label7->Caption= Mes[StrToInt(DateToStr(Date()).SubString(4,2))-1]+ " "+DateToStr(Date()).SubString(7,4);

  //Вывод отчетного периода
  DM->qObnovlenie->Close();
  DM->qObnovlenie->SQL->Clear();
  DM->qObnovlenie->SQL->Add("select * from grafr");
  try
    {
      DM->qObnovlenie->Open();
    }
  catch(Exception &E)
    {
      Application->MessageBox(("Невозможно получить отчетный период из таблицы GRAFR"+ E.Message).c_str(),"Ошибка",
                               MB_OK+MB_ICONERROR);
      Application->Terminate();
      Abort();

    }

  mm=DM->qObnovlenie->FieldByName("mes")->AsInteger;
  yyyy=DM->qObnovlenie->FieldByName("god")->AsInteger;

  Label7->Caption= Mes[mm-1]+ " "+IntToStr(yyyy);

 /* //Вывод определенного периода
  DM->qKomandirovki->Filtered=false;
  DM->qKomandirovki->Filter="data_n>="+QuotedStr("01."+DateToStr(Date()).SubString(4,255));
  DM->qKomandirovki->Filtered=true;*/

  //Вывод определенного периода
  DM->qKomandirovki->Filtered=false;
  DM->qKomandirovki->Filter="data_n>="+QuotedStr("01."+(mm<10 ? "0"+IntToStr(mm) : IntToStr(mm))+"."+IntToStr(yyyy));
  DM->qKomandirovki->Filtered=true;



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

  WorkPath = DocPath + "\\Учет затрат по командировкам";
  Path = GetCurrentDir();
  FindWordPath();

  // Создание ProgressBar на StatusBar
  ProgressBar = new TProgressBar ( StatusBar1 );
  ProgressBar->Parent = StatusBar1;
  ProgressBar->Position = 0;
  ProgressBar->Left = Main->Width-ProgressBar->Width-28;//StatusBar1->Width-ProgressBar->Width-10;//StatusBar1->Panels->Items[0]->Width+StatusBar1->Panels->Items[1]->Width - ProgressBar->Width;//Width*18 + 81;
  //ProgressBar->Anchors = ProgressBar->Anchors << akRight << akTop << akLeft << akBottom;
  ProgressBar->Top = StatusBar1->Height/6;
  ProgressBar->Height = StatusBar1->Height-3;
  PostMessage(ProgressBar->Handle,0x0409,0,clBlue);
  ProgressBar->Visible = false;
}


//---------------------------------------------------------------------------

void __fastcall TMain::DBGridEh1DblClick(TObject *Sender)
{
  N2RedaktirClick(Sender);
}
//---------------------------------------------------------------------------

void __fastcall TMain::N2RedaktirClick(TObject *Sender)
{
  int kol;

  fl_redakt=1;
  Vvod->Label1->Caption="Редактирование записи";

   //Справочник городов
  Vvod->ComboBoxGOROD->Items->Clear();
  DM->qSP_city->First();
  kol = DM->qSP_city->RecordCount;
  for (int i=1; i<=kol; i++)
    {
      Vvod->ComboBoxGOROD->Items->Add(DM->qSP_city->FieldByName("city")->AsString);
      DM->qSP_city->Next();
    }
  DM->qSP_city->First();

  //Справочник объектов
  Vvod->ComboBoxOBEKT->Items->Clear();
  DM->qSP_obekt->First();
  kol = DM->qSP_obekt->RecordCount;
  for (int i=1; i<=kol; i++)
    {
      Vvod->ComboBoxOBEKT->Items->Add(DM->qSP_obekt->FieldByName("obekt")->AsString);
      DM->qSP_obekt->Next();
    }
  DM->qSP_obekt->First();

  //Справочник гостиниц
  Vvod->ComboBoxGOSTINICA->Items->Clear();
  DM->qSP_gostinica->First();
  kol = DM->qSP_gostinica->RecordCount;
  for (int i=1; i<=kol; i++)
    {
      Vvod->ComboBoxGOSTINICA->Items->Add(DM->qSP_gostinica->FieldByName("gostinica")->AsString);
      DM->qSP_gostinica->Next();
    }
  DM->qSP_gostinica->First();

  //Заполнение Edit-ов
  SetInfoEdit();

  Vvod->ShowModal();

}
//---------------------------------------------------------------------------
//Заполнение Edit-ов
void __fastcall TMain::SetInfoEdit()
{
  Vvod->EditKOD_KOM->Text=DM->qKomandirovki->FieldByName("kod_kom")->AsString;
  Vvod->EditZEX->Text=DM->qKomandirovki->FieldByName("zex")->AsString;
  Vvod->EditTN->Text=DM->qKomandirovki->FieldByName("tn")->AsString;
  Vvod->EditFIO->Text=DM->qKomandirovki->FieldByName("fio")->AsString;
  Vvod->EditPROF->Text=DM->qKomandirovki->FieldByName("prof")->AsString;
  Vvod->EditGRADE->Text=DM->qKomandirovki->FieldByName("grade")->AsString;
  //Vvod->EditG_KIEV->Text=DM->qKomandirovki->FieldByName("g_kiev")->AsString;
 // Vvod->EditG_UKR->Text=DM->qKomandirovki->FieldByName("g_kiev")->AsString;
 // Vvod->EditG_ZAGRAN->Text=DM->qKomandirovki->FieldByName("g_zagran")->AsString;
  Vvod->EditDATA_N->Text=DM->qKomandirovki->FieldByName("data_n")->AsString;
  Vvod->EditDATA_K->Text=DM->qKomandirovki->FieldByName("data_k")->AsString;
  Vvod->EditSROK->Text=DM->qKomandirovki->FieldByName("srok")->AsString;
  Vvod->EditDATA_ZAK->Text=DM->qKomandirovki->FieldByName("data_zak")->AsString;
  Vvod->EditADRESS->Text=DM->qKomandirovki->FieldByName("adress")->AsString;
  Vvod->EditGOST_ADR->Text=DM->qKomandirovki->FieldByName("gost_adr")->AsString;
  Vvod->EditNAPRAVL->Text=DM->qKomandirovki->FieldByName("napravl")->AsString;
  Vvod->EditDATA_GOST_N->Text=DM->qKomandirovki->FieldByName("data_gost_n")->AsString;
  Vvod->EditDATA_GOST_K->Text=DM->qKomandirovki->FieldByName("data_gost_k")->AsString;
  Vvod->EditSTOIM->Text=DM->qKomandirovki->FieldByName("stoim")->AsString;
  Vvod->EditN_DOCUM->Text=DM->qKomandirovki->FieldByName("n_docum")->AsString;
  Vvod->EditSUM_SUT->Text=DM->qKomandirovki->FieldByName("sum_sut")->AsString;
  Vvod->EditSUM_PROGIV->Text=DM->qKomandirovki->FieldByName("sum_progiv")->AsString;
  Vvod->EditSUM_TRANSP->Text=DM->qKomandirovki->FieldByName("sum_transp")->AsString;
  Vvod->EditSUM_AVIA->Text=DM->qKomandirovki->FieldByName("sum_avia")->AsString;
  Vvod->EditSUM_GD->Text=DM->qKomandirovki->FieldByName("sum_gd")->AsString;
  Vvod->EditSUM_PROCH->Text=DM->qKomandirovki->FieldByName("sum_proch")->AsString;
  Vvod->MemoPRIMECH->Text=DM->qKomandirovki->FieldByName("primech")->AsString;

  Vvod->LabelZEX->Caption=DM->qKomandirovki->FieldByName("zex_naim")->AsString;

  if (DM->qKomandirovki->FieldByName("avia")->AsString==1) Vvod->CheckBoxAVIA->Checked=true;
  else Vvod->CheckBoxAVIA->Checked=false;
  if (DM->qKomandirovki->FieldByName("gd")->AsString==1) Vvod->CheckBoxGD->Checked=true;
  else Vvod->CheckBoxGD->Checked=false;
  if (DM->qKomandirovki->FieldByName("bus")->AsString==1) Vvod->CheckBoxBUS->Checked=true;
  else Vvod->CheckBoxBUS->Checked=false;
  if (DM->qKomandirovki->FieldByName("avto")->AsString==1) Vvod->CheckBoxAVTO->Checked=true;
  else Vvod->CheckBoxAVTO->Checked=false;
  if (DM->qKomandirovki->FieldByName("proezd")->AsString==1) Vvod->CheckBoxPROEZD->Checked=true;
  else Vvod->CheckBoxPROEZD->Checked=false;

  Vvod->ComboBoxCHEL->ItemIndex=Vvod->ComboBoxCHEL->Items->IndexOf(DM->qKomandirovki->FieldByName("chel")->AsString);
  Vvod->ComboBoxSTRANA->ItemIndex=Vvod->ComboBoxSTRANA->Items->IndexOf(DM->qKomandirovki->FieldByName("strana")->AsString);
  Vvod->ComboBoxGOROD->ItemIndex=Vvod->ComboBoxGOROD->Items->IndexOf(DM->qKomandirovki->FieldByName("gorod")->AsString);
  Vvod->ComboBoxOBEKT->ItemIndex=Vvod->ComboBoxOBEKT->Items->IndexOf(DM->qKomandirovki->FieldByName("obekt")->AsString);
  Vvod->ComboBoxGOSTINICA->ItemIndex=Vvod->ComboBoxGOSTINICA->Items->IndexOf(DM->qKomandirovki->FieldByName("gostinica")->AsString);
}
//---------------------------------------------------------------------------

void __fastcall TMain::N1Click(TObject *Sender)
{
  fl_redakt=0;
  Vvod->Label1->Caption="Добавление записи";

  //Очистка ComboBox город, гостиница, объект
  Vvod->ComboBoxGOROD->Items->Clear();
  Vvod->ComboBoxOBEKT->Items->Clear();
  Vvod->ComboBoxGOSTINICA->Items->Clear();

  Vvod->EditKOD_KOM->Text="";
  Vvod->EditZEX->Text="";
  Vvod->EditTN->Text="";
  Vvod->EditFIO->Text="";
  Vvod->EditPROF->Text="";
  Vvod->EditGRADE->Text="";
  Vvod->EditG_KIEV->Text="";
  Vvod->EditG_UKR->Text="";
  Vvod->EditG_ZAGRAN->Text="";
  Vvod->EditDATA_N->Text="";
  Vvod->EditDATA_K->Text="";
  Vvod->EditSROK->Text="";
  Vvod->EditDATA_ZAK->Text="";
  Vvod->EditNAPRAVL->Text="";
  Vvod->EditDATA_GOST_N->Text="";
  Vvod->EditDATA_GOST_K->Text="";
  Vvod->EditSTOIM->Text="";
  Vvod->EditN_DOCUM->Text="";
  Vvod->EditSUM_SUT->Text="";
  Vvod->EditSUM_PROGIV->Text="";
  Vvod->EditSUM_TRANSP->Text="";
  Vvod->EditSUM_AVIA->Text="";
  Vvod->EditSUM_GD->Text="";
  Vvod->EditSUM_PROCH->Text="";
  Vvod->MemoPRIMECH->Text="";
  Vvod->EditADRESS->Text="";
  Vvod->EditGOST_ADR->Text="";

  Vvod->LabelZEX->Caption="";

  Vvod->CheckBoxAVIA->Checked=false;
  Vvod->CheckBoxGD->Checked=false;
  Vvod->CheckBoxBUS->Checked=false;
  Vvod->CheckBoxAVTO->Checked=false;
  Vvod->CheckBoxPROEZD->Checked=false;

  Vvod->ComboBoxCHEL->ItemIndex=-1;
  Vvod->ComboBoxSTRANA->ItemIndex=-1;
  Vvod->ComboBoxGOROD->ItemIndex=-1;
  Vvod->ComboBoxOBEKT->ItemIndex=-1;
  Vvod->ComboBoxGOSTINICA->ItemIndex=-1;

  Vvod->ComboBoxCHEL->Text="";
  Vvod->ComboBoxSTRANA->Text="";
  Vvod->ComboBoxGOROD->Text="";
  Vvod->ComboBoxOBEKT->Text="";
  Vvod->ComboBoxGOSTINICA->Text="";


  Vvod->ShowModal();
}
//---------------------------------------------------------------------------

void __fastcall TMain::N5OBRAT_SVClick(TObject *Sender)
{
  Gostinica->Label1->Caption="Город: "+DM->qKomandirovki->FieldByName("gorod")->AsString;
  Gostinica->Label2->Caption="Гостиница: "+DM->qKomandirovki->FieldByName("gostinica")->AsString;


  Gostinica->EditCOMFORT->Text=DM->qKomandirovki->FieldByName("comfort")->AsString;
  Gostinica->EditCLEAR->Text=DM->qKomandirovki->FieldByName("clear")->AsString;
  Gostinica->EditPERSONAL->Text=DM->qKomandirovki->FieldByName("personal")->AsString;
  Gostinica->EditPITANIE->Text=DM->qKomandirovki->FieldByName("pitanie")->AsString;
  Gostinica->EditSERVIS->Text=DM->qKomandirovki->FieldByName("servis")->AsString;
  Gostinica->EditUSLUGI->Text=DM->qKomandirovki->FieldByName("uslugi")->AsString;
  Gostinica->EditRASPOLOG->Text=DM->qKomandirovki->FieldByName("raspolog")->AsString;
  Gostinica->EditVPECHAT->Text=DM->qKomandirovki->FieldByName("vpechat")->AsString;
  Gostinica->EditORGANIZ->Text=DM->qKomandirovki->FieldByName("organiz")->AsString;




  Gostinica->ShowModal();
}
//---------------------------------------------------------------------------

void __fastcall TMain::FormKeyDown(TObject *Sender, WORD &Key,
      TShiftState Shift)
{
 if (Key==VK_RETURN)
  FindNextControl((TWinControl *)Sender, true, true, false)->SetFocus();
}
//---------------------------------------------------------------------------

void __fastcall TMain::DBGridEh1DrawColumnCell(TObject *Sender,
      const TRect &Rect, int DataCol, TColumnEh *Column,
      TGridDrawState State)
{
  if (!DM->qKomandirovki->FieldByName("data_zak")->AsString.IsEmpty())
    {
      ((TDBGridEh *) Sender)->Canvas->Brush->Color=(TColor)0x0092E9FE;
      //((TDBGridEh *) Sender)->Canvas->Brush->Color=clRed;
      ((TDBGridEh *) Sender)->DefaultDrawColumnCell(Rect, DataCol, Column, State);
    }
  else
    {
      ((TDBGridEh *) Sender)->Canvas->Brush->Color=clCream;
      ((TDBGridEh *) Sender)->DefaultDrawColumnCell(Rect, DataCol, Column, State);
    }

 // выделение цветом активной записи
  if (State.Contains(gdSelected) )
    {
      if(!DM->qKomandirovki->FieldByName("data_zak")->AsString.IsEmpty())
        {
          ((TDBGridEh *) Sender)->Canvas->Brush->Color = clScrollBar;
        }
      else
        {
          ((TDBGridEh *) Sender)->Canvas->Brush->Color = clSkyBlue;//(TColor)0x00DEF5F4;//clInfoBk;
        }
      ((TDBGridEh *) Sender)->Canvas->Font->Color = clBlack;
    }
  ((TDBGridEh *) Sender)->DefaultDrawColumnCell(Rect, DataCol, Column, State);
}
//---------------------------------------------------------------------------

void __fastcall TMain::FormShow(TObject *Sender)
{
  DBGridEh1->SetFocus();

  EditZEX->Text="";
  EditTN->Text="";
  EditFAM->Text="";
  EditS->Text="";
  EditPO->Text="";

 // Label7->Caption="";

}
//---------------------------------------------------------------------------


//Справочник цели командировки
void __fastcall TMain::N9Click(TObject *Sender)
{
  // red_spr=0;

  Sprav->PageControl1->ActivePage = Sprav->TabSheet1; //Активная страница
  Sprav->TabSheet1->Caption = "Цели командировки";
  Sprav->TabSheet1->TabVisible = true;
  Sprav->TabSheet2->TabVisible = true;
  Sprav->TabSheet3->TabVisible = true;
  Sprav->TabSheet4->TabVisible = true;
  Sprav->TabSheet5->TabVisible = true;
  Sprav->TabSheet6->TabVisible = true;


  //Sprav->Panel1->Visible = true;

  Sprav->ShowModal();
}
//---------------------------------------------------------------------------

//Справочник грейдов
void __fastcall TMain::N7Click(TObject *Sender)
{
 // red_spr=0;

  Sprav->PageControl1->ActivePage = Sprav->TabSheet2; //Активная страница
  Sprav->TabSheet2->Caption = "Грейды";
  Sprav->TabSheet1->TabVisible = true;
  Sprav->TabSheet2->TabVisible = true;
  Sprav->TabSheet3->TabVisible = true;
  Sprav->TabSheet4->TabVisible = true;
  Sprav->TabSheet5->TabVisible = true;
  Sprav->TabSheet6->TabVisible = true;


  //Sprav->Panel1->Visible = true;

  Sprav->ShowModal();
}
//---------------------------------------------------------------------------

//Справочник гостиниц
void __fastcall TMain::N5Click(TObject *Sender)
{
  // red_spr=0;

  Sprav->PageControl1->ActivePage = Sprav->TabSheet3; //Активная страница
  Sprav->TabSheet3->Caption = "Гостиницы";
  Sprav->TabSheet1->TabVisible = true;
  Sprav->TabSheet2->TabVisible = true;
  Sprav->TabSheet3->TabVisible = true;
  Sprav->TabSheet4->TabVisible = true;
  Sprav->TabSheet5->TabVisible = true;
  Sprav->TabSheet6->TabVisible = true;


  //Sprav->Panel1->Visible = true;

  Sprav->ShowModal();
}
//---------------------------------------------------------------------------

//Справочник объектов
void __fastcall TMain::N12Click(TObject *Sender)
{
  // red_spr=0;

  Sprav->PageControl1->ActivePage = Sprav->TabSheet4; //Активная страница
  Sprav->TabSheet4->Caption = "Объекты";
  Sprav->TabSheet1->TabVisible = true;
  Sprav->TabSheet2->TabVisible = true;
  Sprav->TabSheet3->TabVisible = true;
  Sprav->TabSheet4->TabVisible = true;
  Sprav->TabSheet5->TabVisible = true;
  Sprav->TabSheet6->TabVisible = true;


  //Sprav->Panel1->Visible = true;

  Sprav->ShowModal();
}
//---------------------------------------------------------------------------

//Справочник стран
void __fastcall TMain::N10Click(TObject *Sender)
{
  // red_spr=0;

  Sprav->PageControl1->ActivePage = Sprav->TabSheet5; //Активная страница
  Sprav->TabSheet5->Caption = "Страны";
  Sprav->TabSheet1->TabVisible = true;
  Sprav->TabSheet2->TabVisible = true;
  Sprav->TabSheet3->TabVisible = true;
  Sprav->TabSheet4->TabVisible = true;
  Sprav->TabSheet5->TabVisible = true;
  Sprav->TabSheet6->TabVisible = true;


  //Sprav->Panel1->Visible = true;

  Sprav->ShowModal();
}
//---------------------------------------------------------------------------

//Справочник городов
void __fastcall TMain::N11Click(TObject *Sender)
{
  // red_spr=0;

  Sprav->PageControl1->ActivePage = Sprav->TabSheet6; //Активная страница
  Sprav->TabSheet6->Caption = "Города";
  Sprav->TabSheet1->TabVisible = true;
  Sprav->TabSheet2->TabVisible = true;
  Sprav->TabSheet3->TabVisible = true;
  Sprav->TabSheet4->TabVisible = true;
  Sprav->TabSheet5->TabVisible = true;
  Sprav->TabSheet6->TabVisible = true;


  //Sprav->Panel1->Visible = true;

  Sprav->ShowModal();
}
//---------------------------------------------------------------------------

void __fastcall TMain::N3Click(TObject *Sender)
{
  if (Application->MessageBox("Вы действительно хотите удалить выбранную запись?","Удаление записи",
                               MB_YESNO+MB_ICONINFORMATION)==ID_NO)
    {
      Abort();
    }

  //Удаление записи
  DM->qObnovlenie->Close();
  DM->qObnovlenie->SQL->Clear();
  DM->qObnovlenie->SQL->Add("delete from komandirovki where rowid=chartorowid("+QuotedStr(DM->qKomandirovki->FieldByName("rw")->AsString)+")");
  try
    {
      DM->qObnovlenie->ExecSQL();
    }
  catch(Exception &E)
    {
      Application->MessageBox(("Возникла ошибка при удалении записи \nиз таблицы по командировкам (KOMANDIROVKI) "+E.Message).c_str(),"Удаление записи",
                              MB_OK+MB_ICONERROR);
      //Логи
      InsertLog("Удаление записи не выполнено: цех "+DM->qKomandirovki->FieldByName("zex")->AsString+" таб.№ "+DM->qKomandirovki->FieldByName("tn")->AsString);
      Abort();
    }

  //Логи
  InsertLog("Удаление записи выполнено: цех "+DM->qKomandirovki->FieldByName("zex")->AsString+" таб.№ "+DM->qKomandirovki->FieldByName("tn")->AsString);
  
  DM->qKomandirovki->Requery();

  Application->MessageBox("Запись успешно удалена =)","Удаление записи",
                          MB_OK+MB_ICONINFORMATION);
}
//---------------------------------------------------------------------------

//Реестр командировок
void __fastcall TMain::N13Click(TObject *Sender)
{
  AnsiString sFile, Sql;
  int n=7;
  Variant AppEx, Sh;


  StatusBar1->SimpleText=" Идет формирование реестра командировок в Excel...";

  Sql="select rownum, z.*                                                                    \
       from (                                                                                \
              select                                                                         \
                     initcap(fio) as fio,                                                    \
                    (select naim from sp_komandir where kod=chel) as chel,                   \
                     zex,                                                                    \
                     tn,                                                                     \
                     grade,                                                                  \
                     kod_kom,                                                                \
                     data_n||' - '||data_k as dat,                                           \
                    (select city from sp_city where kod=gorod) as gorod,                     \
                    (select gostinica from sp_gostinica where kod=k.gostinica) as gostinica, \
                     data_gost_n||' - '||data_gost_k as dat_gost,                            \
                     stoim,                                                                  \
                     sum_progiv,                                                             \
                     case when nvl(avia,0)=1 then 'Авиа' else NULL end transp1,                        \
                     case when nvl(gd,0)=1 then 'Ж/д' else NULL end transp2,                           \
                     case when nvl(bus,0)=1 then 'Автобус' else NULL  end transp3,                      \
                     case when nvl(Avto,0)=1 then 'Авто' else NULL  end transp4,                        \
                     case when nvl(proezd,0)=1 then 'Прочие' else NULL end transp5,                    \
                     napravl,                                                                \
                     sum_avia as sum_transp1,                                                \
                     sum_gd as sum_transp2,                                                  \
                     sum_transp as sum_transp3                                               \
               from komandirovki k                                                           \
               where (data_n between '01."+(mm<10? "0"+IntToStr(mm): IntToStr(mm))+"."+IntToStr(yyyy)+"' \
                             and '"+DateToStr(EndOfTheMonth(StrToDate("01."+(mm<10? "0"+IntToStr(mm): IntToStr(mm))+"."+IntToStr(yyyy)))).SubString(1,2)+"."+(mm<10? "0"+IntToStr(mm): IntToStr(mm))+"."+IntToStr(yyyy)+"'  \
                      or data_k between '01."+(mm<10? "0"+IntToStr(mm): IntToStr(mm))+"."+IntToStr(yyyy)+"' \
                and '"+DateToStr(EndOfTheMonth(StrToDate("01."+(mm<10? "0"+IntToStr(mm): IntToStr(mm))+"."+IntToStr(yyyy)))).SubString(1,2)+"."+(mm<10? "0"+IntToStr(mm): IntToStr(mm))+"."+IntToStr(yyyy)+"'  )\
               order by zex, tn                                                              \
             ) z";                                                                           \


                /* where (data_n between '01."+DateToStr(Date()).SubString(4,2)+"."+DateToStr(Date()).SubString(7,4)+"' \
                             and '"+DateToStr(EndOfTheMonth(Date())).SubString(1,2)+"."+DateToStr(Date()).SubString(4,2)+"."+DateToStr(Date()).SubString(7,4)+"'  \
                      or data_k between '01."+DateToStr(Date()).SubString(4,2)+"."+DateToStr(Date()).SubString(7,4)+"' \
                and '"+DateToStr(EndOfTheMonth(Date())).SubString(1,2)+"."+DateToStr(Date()).SubString(4,2)+"."+DateToStr(Date()).SubString(7,4)+"'  )\
               order by zex, tn                                                              \
             ) z";                                                                           \
*/

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
      Application->MessageBox(("Возникла ошибка при получении данных из таблицы по командировкам (KOMANDIROVKI)" + E.Message).c_str(),"Ошибка",
                              MB_OK+MB_ICONERROR);
     // InsertLog("Возникла ошибка при формировании списка работников по цехам в Excel");
     // DM->qLogs->Requery();
      StatusBar1->SimpleText="";
      Abort();
    }

  if (DM->qObnovlenie->RecordCount==0)
    {
      Application->MessageBox("Нет данных за отчетный период!!!","Предупреждение",
                               MB_OK+MB_ICONINFORMATION);
      StatusBar1->SimpleText="";
      Abort();
    }

  sFile = Path+"\\RTF\\reestr_komandir.xlsx";

  //Создание папки, если ее не существует
  ForceDirectories(WorkPath);


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
      StatusBar1->SimpleText="";
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
              StatusBar1->SimpleText="";
              ProgressBar->Visible = false;
              Cursor = crDefault;
            }

          //AppEx.OlePropertySet("Visible",true);

          int i=1;
          n=7;

          Variant Massiv, Massiv2;
          Massiv = VarArrayCreate(OPENARRAY(int,(0,17)),varVariant); //массив на 16 элементов
          Massiv2 = VarArrayCreate(OPENARRAY(int,(0,4)),varVariant); //массив на 3 элементов


          Sh.OlePropertyGet("Range", "I3").OlePropertySet("Value",MonthOf(Date()));
          Sh.OlePropertyGet("Range", "J3").OlePropertySet("Value",YearOf(Date()));

          while (!DM->qObnovlenie->Eof)
            {
              Massiv.PutElement(DM->qObnovlenie->FieldByName("rownum")->AsString.c_str(), 0);
              Massiv.PutElement(DM->qObnovlenie->FieldByName("fio")->AsString.c_str(), 1);
              Massiv.PutElement(DM->qObnovlenie->FieldByName("chel")->AsString.c_str(), 2);
              Massiv.PutElement(DM->qObnovlenie->FieldByName("zex")->AsString.c_str(), 4);
              Massiv.PutElement(DM->qObnovlenie->FieldByName("tn")->AsString.c_str(), 5);
              Massiv.PutElement(DM->qObnovlenie->FieldByName("grade")->AsString.c_str(), 6);
              Massiv.PutElement(DM->qObnovlenie->FieldByName("kod_kom")->AsString.c_str(), 7);
              Massiv.PutElement(DM->qObnovlenie->FieldByName("dat")->AsString.c_str(), 8);
              Massiv.PutElement(DM->qObnovlenie->FieldByName("gorod")->AsString.c_str(), 9);
              Massiv.PutElement(DM->qObnovlenie->FieldByName("gostinica")->AsString.c_str(), 10);
              Massiv.PutElement(DM->qObnovlenie->FieldByName("dat_gost")->AsString.c_str(), 11);
              Massiv.PutElement(DM->qObnovlenie->FieldByName("stoim")->AsString.c_str(), 12);
              Massiv.PutElement(DM->qObnovlenie->FieldByName("sum_progiv")->AsString.c_str(), 13);

              Sh.OlePropertyGet("Range", ("A" + IntToStr(n) + ":Q" + IntToStr(n)).c_str()).OlePropertySet("Value", Massiv); //строка с данными с ячейки A по ячейку АВ

              if (!DM->qObnovlenie->FieldByName("transp1")->AsString.IsEmpty())
                {

                  Massiv2.PutElement(DM->qObnovlenie->FieldByName("transp1")->AsString.c_str(), 0);
                  Massiv2.PutElement(DM->qObnovlenie->FieldByName("napravl")->AsString.c_str(), 1);
                  Massiv2.PutElement(DM->qObnovlenie->FieldByName("sum_transp1")->AsString.c_str(), 2);

                  Sh.OlePropertyGet("Range", ("O" + IntToStr(n) + ":Q" + IntToStr(n)).c_str()).OlePropertySet("Value", Massiv2); //строка с данными с ячейки A по ячейку АВ

                  i++;
                  n++;
                }
              if (!DM->qObnovlenie->FieldByName("transp2")->AsString.IsEmpty())
                {
                  Massiv2.PutElement(DM->qObnovlenie->FieldByName("transp2")->AsString.c_str(), 0);
                  Massiv2.PutElement(DM->qObnovlenie->FieldByName("napravl")->AsString.c_str(), 1);
                  Massiv2.PutElement(DM->qObnovlenie->FieldByName("sum_transp2")->AsString.c_str(), 2);

                  Sh.OlePropertyGet("Range", ("O" + IntToStr(n) + ":Q" + IntToStr(n)).c_str()).OlePropertySet("Value", Massiv2); //строка с данными с ячейки A по ячейку АВ

                  i++;
                  n++;
                }
              if (!DM->qObnovlenie->FieldByName("transp3")->AsString.IsEmpty())
                {
                  Massiv2.PutElement(DM->qObnovlenie->FieldByName("transp3")->AsString.c_str(), 0);
                  Massiv2.PutElement(DM->qObnovlenie->FieldByName("napravl")->AsString.c_str(), 1);
                  Massiv2.PutElement(DM->qObnovlenie->FieldByName("sum_transp3")->AsString.c_str(), 2);

                  Sh.OlePropertyGet("Range", ("O" + IntToStr(n) + ":Q" + IntToStr(n)).c_str()).OlePropertySet("Value", Massiv2); //строка с данными с ячейки A по ячейку АВ

                  i++;
                  n++;
                }
              if (!DM->qObnovlenie->FieldByName("transp4")->AsString.IsEmpty())
                {
                  Massiv2.PutElement(DM->qObnovlenie->FieldByName("transp4")->AsString.c_str(), 0);
                  Massiv2.PutElement(DM->qObnovlenie->FieldByName("napravl")->AsString.c_str(), 1);
                  Massiv2.PutElement(DM->qObnovlenie->FieldByName("sum_transp3")->AsString.c_str(), 2);

                  Sh.OlePropertyGet("Range", ("O" + IntToStr(n) + ":Q" + IntToStr(n)).c_str()).OlePropertySet("Value", Massiv2); //строка с данными с ячейки A по ячейку АВ

                  i++;
                  n++;
                }
              if (!DM->qObnovlenie->FieldByName("transp5")->AsString.IsEmpty())
                {
                  Massiv2.PutElement(DM->qObnovlenie->FieldByName("transp5")->AsString.c_str(), 0);
                  Massiv2.PutElement(DM->qObnovlenie->FieldByName("napravl")->AsString.c_str(), 1);
                  Massiv2.PutElement(DM->qObnovlenie->FieldByName("sum_transp3")->AsString.c_str(), 2);

                  Sh.OlePropertyGet("Range", ("O" + IntToStr(n) + ":Q" + IntToStr(n)).c_str()).OlePropertySet("Value", Massiv2); //строка с данными с ячейки A по ячейку АВ

                  i++;
                  n++;
                }

              DM->qObnovlenie->Next();
              ProgressBar->Position++;
            }

          // вставляем в шаблон нужное количество строк

          //окрашивание ячеек
     /*     Sh.OlePropertyGet("Range",("M18:M"+IntToStr(n-1)).c_str()).OlePropertyGet("Interior").OlePropertySet("Color",14277081);
          Sh.OlePropertyGet("Range",("P18:R"+IntToStr(n-1)).c_str()).OlePropertyGet("Interior").OlePropertySet("Color",14277081);

          Sh.OlePropertyGet("Range",("B18:K"+IntToStr(n-1)).c_str()).OlePropertyGet("Interior").OlePropertySet("Color",14408946);
          Sh.OlePropertyGet("Range",("N18:N"+IntToStr(n-1)).c_str()).OlePropertyGet("Interior").OlePropertySet("Color",14408946);
       */
          //рисуем сетку
          Sh.OlePropertyGet("Range",("A7:Q"+IntToStr(n-1)).c_str()).OlePropertyGet("Borders").OlePropertySet("LineStyle", xlContinuous);


          //Отключить вывод сообщений с вопросами типа "Заменить файл..."
          AppEx.OlePropertySet("DisplayAlerts",false);


          //Сохранить книгу в папке в файле по указанию
          AnsiString vAsCurDir1=WorkPath+"\\Реестр командировок.xlsx";
          Sh.OleProcedure("SaveAs",vAsCurDir1.c_str());

          //Закрыть открытое приложение Excel
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
          StatusBar1->SimpleText= "";
        }
      catch (...)
        {
          AppEx.OleProcedure("Quit");
          AppEx = Unassigned;
          Cursor = crDefault;
          ProgressBar->Position=0;
          StatusBar1->SimpleText= "";
          ProgressBar->Visible=false;
         // InsertLog("Возникла ошибка при формировании списка работников по цехам в Excel");
          Abort();
        }
    }

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


void __fastcall TMain::N17Click(TObject *Sender)
{
  AnsiString sFile, Sql;
  int n=7;
  Variant AppEx, Sh;


  StatusBar1->SimpleText=" Идет формирование рейтинга гостиниц по командируемым работникам в Excel...";

  Sql="select fio,                                                                     \
              to_char(data_n,'dd.mm.yyyy')||' - '||to_char(data_k,'dd.mm.yyyy') as data,                                           \
             (select city from sp_city where kod=k.gorod) as gorod,                    \
             (select gostinica from sp_gostinica where kod=k.gostinica) as gostinica,  \
              comfort,                                                                 \
              clear,                                                                   \
              personal,                                                                \
              pitanie,                                                                 \
              servis,                                                                  \
              uslugi,                                                                  \
              raspolog,                                                                \
              vpechat,                                                                 \
              round((sum(nvl(comfort,0))/count(*)+ sum(nvl(clear,0))/count(*)+ sum(nvl(personal,0))/count(*)+ \
              sum(nvl(pitanie,0))/count(*)+sum(nvl(servis,0))/count(*)+sum(nvl(uslugi,0))/count(*)+           \
              sum(nvl(raspolog,0))/count(*)+sum(nvl(vpechat,0))/count(*))/40*100,2) as reit                   \
       from komandirovki k                                                                                    \
       group by fio,data_n,data_k,gorod,gostinica,comfort,clear,personal,                                     \
                pitanie,servis,uslugi,raspolog,vpechat                                                        \
       order by fio";
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
      Application->MessageBox(("Возникла ошибка при получении данных из таблицы по командировкам (KOMANDIROVKI)" + E.Message).c_str(),"Ошибка",
                              MB_OK+MB_ICONERROR);
     // InsertLog("Возникла ошибка при формировании списка работников по цехам в Excel");
     // DM->qLogs->Requery();
      StatusBar1->SimpleText="";
      Abort();
    }

  sFile = Path+"\\RTF\\reit_gost_rab.xlsx";

  //Создание папки, если ее не существует
  ForceDirectories(WorkPath);


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
      StatusBar1->SimpleText="";
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
              StatusBar1->SimpleText="";
              ProgressBar->Visible = false;
              Cursor = crDefault;
            }

          int i=1;
          n=7;

          Variant Massiv;
          Massiv = VarArrayCreate(OPENARRAY(int,(0,14)),varVariant); //массив на 18 элементов

          Sh.OlePropertyGet("Range", "E3").OlePropertySet("Value",YearOf(Date()));

          while (!DM->qObnovlenie->Eof)
            {
             // Massiv.PutElement(i, 0);
              Massiv.PutElement(DM->qObnovlenie->FieldByName("fio")->AsString.c_str(), 0);
              Massiv.PutElement(DM->qObnovlenie->FieldByName("data")->AsString.c_str(), 1);
              Massiv.PutElement(DM->qObnovlenie->FieldByName("gorod")->AsString.c_str(), 2);
              Massiv.PutElement(DM->qObnovlenie->FieldByName("gostinica")->AsString.c_str(), 3);
              Massiv.PutElement(DM->qObnovlenie->FieldByName("comfort")->AsString.c_str(), 4);
              Massiv.PutElement(DM->qObnovlenie->FieldByName("clear")->AsString.c_str(), 5);
              Massiv.PutElement(DM->qObnovlenie->FieldByName("personal")->AsString.c_str(), 6);
              Massiv.PutElement(DM->qObnovlenie->FieldByName("pitanie")->AsString.c_str(), 7);
              Massiv.PutElement(DM->qObnovlenie->FieldByName("servis")->AsString.c_str(), 8);
              Massiv.PutElement(DM->qObnovlenie->FieldByName("uslugi")->AsString.c_str(), 9);
              Massiv.PutElement(DM->qObnovlenie->FieldByName("raspolog")->AsString.c_str(), 10);
              Massiv.PutElement(DM->qObnovlenie->FieldByName("vpechat")->AsString.c_str(), 11);
              Massiv.PutElement(DM->qObnovlenie->FieldByName("reit")->AsString.c_str(), 12);


              Sh.OlePropertyGet("Range", ("A" + IntToStr(n) + ":M" + IntToStr(n)).c_str()).OlePropertySet("Value", Massiv); //строка с данными с ячейки A по ячейку АВ

              i++;
              n++;
              DM->qObnovlenie->Next();
              ProgressBar->Position++;
            }

          // вставляем в шаблон нужное количество строк

          //окрашивание ячеек
          Sh.OlePropertyGet("Range",("A7:D"+IntToStr(n-1)).c_str()).OlePropertyGet("Interior").OlePropertySet("Color",14277081);
          Sh.OlePropertyGet("Range",("M7:M"+IntToStr(n-1)).c_str()).OlePropertyGet("Interior").OlePropertySet("Color",14277081);

          //рисуем сетку
          Sh.OlePropertyGet("Range",("A7:M"+IntToStr(n-1)).c_str()).OlePropertyGet("Borders").OlePropertySet("LineStyle", xlContinuous);
     
          //Отключить вывод сообщений с вопросами типа "Заменить файл..."
          AppEx.OlePropertySet("DisplayAlerts",false);


          //Сохранить книгу в папке в файле по указанию
          AnsiString vAsCurDir1=WorkPath+"\\Реестр командировок.xlsx";
          Sh.OleProcedure("SaveAs",vAsCurDir1.c_str());

          //Закрыть открытое приложение Excel
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
          StatusBar1->SimpleText= "";
        }
      catch (...)
        {
          AppEx.OleProcedure("Quit");
          AppEx = Unassigned;
          Cursor = crDefault;
          ProgressBar->Position=0;
          StatusBar1->SimpleText= "";
          ProgressBar->Visible=false;
         // InsertLog("Возникла ошибка при формировании списка работников по цехам в Excel");
          Abort();
        }
    }
}
//---------------------------------------------------------------------------

void __fastcall TMain::N16Click(TObject *Sender)
{
 AnsiString sFile, Sql;
  int n=7;
  Variant AppEx, Sh;


  StatusBar1->SimpleText=" Идет формирование общего рейтинга гостиниц в Excel...";

  Sql="select (select city from sp_city where kod=(select kod_city from sp_gostinica where kod=k.gostinica)) as city,  \
              (select gostinica from sp_gostinica where kod=k.gostinica) as gostinica,  \
              sum(nvl(comfort,0))/count(*) as comfort,                                  \
              sum(nvl(clear,0))/count(*) as clear,                                      \
              sum(nvl(personal,0))/count(*) as personal,                                \
              sum(nvl(pitanie,0))/count(*) as pitanie,                                  \
              sum(nvl(servis,0))/count(*) as servis,                                    \
              sum(nvl(uslugi,0))/count(*) as uslugi,                                    \
              sum(nvl(raspolog,0))/count(*) as raspolog,                                \
              sum(nvl(vpechat,0))/count(*) as vpechat,                                  \
              round((sum(nvl(comfort,0))/count(*)+                                      \
                     sum(nvl(clear,0))/count(*)+                                        \
                     sum(nvl(personal,0))/count(*)+                                     \
                     sum(nvl(pitanie,0))/count(*)+                                      \
                     sum(nvl(servis,0))/count(*)+                                       \
                     sum(nvl(uslugi,0))/count(*)+                                       \
                     sum(nvl(raspolog,0))/count(*)+                                     \
                     sum(nvl(vpechat,0))/count(*)                                       \
                    )/count(*)/40*100,2) as reit                                        \
      from komandirovki  k                                                              \
      group by gostinica                                                                \
      order by reit desc";                                                              \
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
      Application->MessageBox(("Возникла ошибка при получении данных из таблицы по командировкам (KOMANDIROVKI)" + E.Message).c_str(),"Ошибка",
                              MB_OK+MB_ICONERROR);
     // InsertLog("Возникла ошибка при формировании списка работников по цехам в Excel");
     // DM->qLogs->Requery();
      StatusBar1->SimpleText="";
      Abort();
    }

  sFile = Path+"\\RTF\\reit_gost.xlsx";

  //Создание папки, если ее не существует
  ForceDirectories(WorkPath);


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
      StatusBar1->SimpleText="";
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
              StatusBar1->SimpleText="";
              ProgressBar->Visible = false;
              Cursor = crDefault;
            }

          int i=1;
          n=7;

          Variant Massiv;
          Massiv = VarArrayCreate(OPENARRAY(int,(0,14)),varVariant); //массив на 18 элементов

          Sh.OlePropertyGet("Range", "F3").OlePropertySet("Value",YearOf(Date()));

          while (!DM->qObnovlenie->Eof)
            {
              Massiv.PutElement(i, 0);
              Massiv.PutElement(DM->qObnovlenie->FieldByName("city")->AsString.c_str(), 1);
              Massiv.PutElement(DM->qObnovlenie->FieldByName("gostinica")->AsString.c_str(), 2);
              Massiv.PutElement(DM->qObnovlenie->FieldByName("comfort")->AsString.c_str(), 3);
              Massiv.PutElement(DM->qObnovlenie->FieldByName("clear")->AsString.c_str(), 4);
              Massiv.PutElement(DM->qObnovlenie->FieldByName("personal")->AsString.c_str(), 5);
              Massiv.PutElement(DM->qObnovlenie->FieldByName("pitanie")->AsString.c_str(), 6);
              Massiv.PutElement(DM->qObnovlenie->FieldByName("servis")->AsString.c_str(), 7);
              Massiv.PutElement(DM->qObnovlenie->FieldByName("uslugi")->AsString.c_str(), 8);
              Massiv.PutElement(DM->qObnovlenie->FieldByName("raspolog")->AsString.c_str(), 9);
              Massiv.PutElement(DM->qObnovlenie->FieldByName("vpechat")->AsString.c_str(), 10);
              Massiv.PutElement(DM->qObnovlenie->FieldByName("reit")->AsString.c_str(), 11);


              Sh.OlePropertyGet("Range", ("A" + IntToStr(n) + ":L" + IntToStr(n)).c_str()).OlePropertySet("Value", Massiv); //строка с данными с ячейки A по ячейку АВ

              i++;
              n++;
              DM->qObnovlenie->Next();
              ProgressBar->Position++;
            }

          // вставляем в шаблон нужное количество строк

          //окрашивание ячеек
          Sh.OlePropertyGet("Range",("B7:B"+IntToStr(n-1)).c_str()).OlePropertyGet("Interior").OlePropertySet("Color",14277081);
          Sh.OlePropertyGet("Range",("L7:L"+IntToStr(n-1)).c_str()).OlePropertyGet("Interior").OlePropertySet("Color",14277081);

          //рисуем сетку
          Sh.OlePropertyGet("Range",("A7:L"+IntToStr(n-1)).c_str()).OlePropertyGet("Borders").OlePropertySet("LineStyle", xlContinuous);

          //Отключить вывод сообщений с вопросами типа "Заменить файл..."
          AppEx.OlePropertySet("DisplayAlerts",false);


          //Сохранить книгу в папке в файле по указанию
          AnsiString vAsCurDir1=WorkPath+"\\Реестр командировок.xlsx";
          Sh.OleProcedure("SaveAs",vAsCurDir1.c_str());

          //Закрыть открытое приложение Excel
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
          StatusBar1->SimpleText= "";
        }
      catch (...)
        {
          AppEx.OleProcedure("Quit");
          AppEx = Unassigned;
          Cursor = crDefault;
          ProgressBar->Position=0;
          StatusBar1->SimpleText= "";
          ProgressBar->Visible=false;
         // InsertLog("Возникла ошибка при формировании списка работников по цехам в Excel");
          Abort();
        }
    }        
}
//---------------------------------------------------------------------------

//Командировочные расходы
void __fastcall TMain::N14Click(TObject *Sender)
{
 AnsiString sFile, Sql;
  int n=6;
  Variant AppEx, Sh;


  StatusBar1->SimpleText=" Идет формирование отчета по командировочным расходам в Excel...";

  Sql="select s.*,                                                                                                  \
              (nvl(u_sum_sut,0)+nvl(u_sum_progiv,0)+nvl(u_sum_avia,0)+nvl(u_sum_gd,0)+nvl(u_sum_transp,0)) u_itogo, \
              (nvl(z_sum_sut,0)+nvl(z_sum_progiv,0)+nvl(z_sum_avia,0)+nvl(z_sum_gd,0)+nvl(z_sum_transp,0)) z_itogo  \
       from                                                           \
            (select initcap(fio) as fio,                                     \
                    n_docum, data_zak,                                         \
                   (select naim from sp_komandir where kod=k.chel) as chel,   \
                    case when strana='UA' then nvl(sum_sut,0)         \
                         else NULL end as u_sum_sut,                  \
                    case when strana!='UA' then nvl(sum_sut,0)        \
                         else NULL end as z_sum_sut,                  \
                    case when strana='UA' then nvl(sum_progiv,0)      \
                         else NULL end as u_sum_progiv,               \
                    case when strana!='UA' then nvl(sum_progiv,0)     \
                         else NULL end as z_sum_progiv,               \
                    case when strana='UA' then nvl(sum_avia,0)        \
                         else NULL end as u_sum_avia,                 \
                    case when strana!='UA' then nvl(sum_avia,0)       \
                         else NULL end as z_sum_avia,                 \
                    case when strana='UA' then nvl(sum_gd,0)          \
                         else NULL end as u_sum_gd,                   \
                    case when strana!='UA' then nvl(sum_gd,0)         \
                         else NULL end as z_sum_gd,                   \
                    case when strana='UA' then nvl(sum_transp,0)+nvl(sum_proch,0)      \
                         else NULL end as u_sum_transp,               \
                    case when strana!='UA' then nvl(sum_transp,0)+nvl(sum_proch,0)     \
                         else NULL end as z_sum_transp                \
              from komandirovki k) s                                  \
              where data_zak is not null order by fio";
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
      Application->MessageBox(("Возникла ошибка при получении данных из таблицы по командировкам (KOMANDIROVKI)" + E.Message).c_str(),"Ошибка",
                              MB_OK+MB_ICONERROR);
     // InsertLog("Возникла ошибка при формировании списка работников по цехам в Excel");
      StatusBar1->SimpleText="";
      Abort();
    }

  if (DM->qObnovlenie->RecordCount==0)
    {
      Application->MessageBox("Нет данных за отчетный период!!!","Предупреждение",
                               MB_OK+MB_ICONINFORMATION);
      StatusBar1->SimpleText="";
      Abort();
    }

  sFile = Path+"\\RTF\\komandir_rashod.xlsx";

  //Создание папки, если ее не существует
  ForceDirectories(WorkPath);


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
      StatusBar1->SimpleText="";
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
              StatusBar1->SimpleText="";
              ProgressBar->Visible = false;
              Cursor = crDefault;
            }

          int i=1;
          n=6;

          Variant Massiv;
          Massiv = VarArrayCreate(OPENARRAY(int,(0,18)),varVariant); //массив на 18 элементов

          Sh.OlePropertyGet("Range", "J2").OlePropertySet("Value",Date());

          while (!DM->qObnovlenie->Eof)
            {
              Massiv.PutElement(i, 0);
              Massiv.PutElement(DM->qObnovlenie->FieldByName("fio")->AsString.c_str(), 1);
              Massiv.PutElement(DM->qObnovlenie->FieldByName("n_docum")->AsString.c_str(), 2);
              Massiv.PutElement(DM->qObnovlenie->FieldByName("chel")->AsString.c_str(), 3);
              Massiv.PutElement(DM->qObnovlenie->FieldByName("u_sum_sut")->AsString.c_str(), 4);
              Massiv.PutElement(DM->qObnovlenie->FieldByName("z_sum_sut")->AsString.c_str(), 5);
              Massiv.PutElement(DM->qObnovlenie->FieldByName("u_sum_progiv")->AsString.c_str(), 6);
              Massiv.PutElement(DM->qObnovlenie->FieldByName("z_sum_progiv")->AsString.c_str(), 7);
              Massiv.PutElement(DM->qObnovlenie->FieldByName("u_sum_avia")->AsString.c_str(), 8);
              Massiv.PutElement(DM->qObnovlenie->FieldByName("z_sum_avia")->AsString.c_str(), 9);
              Massiv.PutElement(DM->qObnovlenie->FieldByName("u_sum_gd")->AsString.c_str(), 10);
              Massiv.PutElement(DM->qObnovlenie->FieldByName("z_sum_gd")->AsString.c_str(), 11);
              Massiv.PutElement(DM->qObnovlenie->FieldByName("u_sum_transp")->AsString.c_str(), 12);
              Massiv.PutElement(DM->qObnovlenie->FieldByName("z_sum_transp")->AsString.c_str(), 13);
              Massiv.PutElement(DM->qObnovlenie->FieldByName("u_itogo")->AsString.c_str(), 14);
              Massiv.PutElement(DM->qObnovlenie->FieldByName("z_itogo")->AsString.c_str(), 15);
              Massiv.PutElement(DM->qObnovlenie->FieldByName("u_itogo")->AsFloat+DM->qObnovlenie->FieldByName("z_itogo")->AsFloat, 16);


              Sh.OlePropertyGet("Range", ("A" + IntToStr(n) + ":Q" + IntToStr(n)).c_str()).OlePropertySet("Value", Massiv); //строка с данными с ячейки A по ячейку АВ

              i++;
              n++;
              DM->qObnovlenie->Next();
              ProgressBar->Position++;
            }

          // вставляем в шаблон нужное количество строк

          //окрашивание ячеек
         // Sh.OlePropertyGet("Range",("B7:B"+IntToStr(n-1)).c_str()).OlePropertyGet("Interior").OlePropertySet("Color",14277081);
         // Sh.OlePropertyGet("Range",("L7:L"+IntToStr(n-1)).c_str()).OlePropertyGet("Interior").OlePropertySet("Color",14277081);

          //рисуем сетку
          Sh.OlePropertyGet("Range",("A6:Q"+IntToStr(n-1)).c_str()).OlePropertyGet("Borders").OlePropertySet("LineStyle", xlContinuous);

          //Отключить вывод сообщений с вопросами типа "Заменить файл..."
          AppEx.OlePropertySet("DisplayAlerts",false);


          //Сохранить книгу в папке в файле по указанию
          AnsiString vAsCurDir1=WorkPath+"\\Реестр командировок.xlsx";
          Sh.OleProcedure("SaveAs",vAsCurDir1.c_str());

          //Закрыть открытое приложение Excel
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
          StatusBar1->SimpleText= "";
        }
      catch (...)
        {
          AppEx.OleProcedure("Quit");
          AppEx = Unassigned;
          Cursor = crDefault;
          ProgressBar->Position=0;
          StatusBar1->SimpleText= "";
          ProgressBar->Visible=false;
         // InsertLog("Возникла ошибка при формировании списка работников по цехам в Excel");
          Abort();
        }
    }
}
//---------------------------------------------------------------------------

//Поиск по цеху и таб.№
void __fastcall TMain::BitBtn1Click(TObject *Sender)
{
  TLocateOptions SearchOptions;

  if (EditZEX->Text.IsEmpty() && EditFAM->Text.IsEmpty())
    {
      Application->MessageBox("Не указан цех!!!","Предупреждение",
                              MB_OK+MB_ICONINFORMATION);
      EditZEX->SetFocus();
      Abort();
    }

  if (!EditZEX->Text.IsEmpty() && EditTN->Text.IsEmpty())
    {
      Application->MessageBox("Не указан табельный №!!!","Предупреждение",
                              MB_OK+MB_ICONINFORMATION);
      EditTN->SetFocus();
      Abort();
    }


  EditTN->Text = EditTN->Text.Length() ==5? EditTN->Text.SubString(2,255) : EditTN->Text;

  if (!EditZEX->Text.IsEmpty())
    {
      Variant locvalues[] = {EditZEX->Text, EditTN->Text};
      if (!DM->qKomandirovki->Locate("zex;tn", VarArrayOf(locvalues,2), SearchOptions))
        {
          Application->MessageBox("Работник не найден!!!","Предупреждение",
                              MB_OK+MB_ICONINFORMATION);
        }
    }
  else
    {
       if (!DM->qKomandirovki->Locate("fam", EditFAM->Text, SearchOptions <<loPartialKey))
        {
          Application->MessageBox("Работник с введенной фамилией не найден!!!","Предупреждение",
                              MB_OK+MB_ICONINFORMATION);
        }
    }


  DBGridEh1->SetFocus();

  EditZEX->Text="";
  EditTN->Text="";
  EditFAM->Text="";
}
//---------------------------------------------------------------------------


void __fastcall TMain::N18Click(TObject *Sender)
{
  WinExec(("\""+ WordPath+"\"\""+ Path+"\\Инструкция пользователя.docx\"").c_str(),SW_MAXIMIZE);
}
//---------------------------------------------------------------------------

void __fastcall TMain::BitBtn2Click(TObject *Sender)
{
  if (EditS->Text.IsEmpty())
    {
      Application->MessageBox("Не указана начальная дата периода","Предупреждение",
                              MB_OK+MB_ICONINFORMATION);
      EditS->SetFocus();
      Abort();
    }

  if (EditPO->Text.IsEmpty())
    {
      Application->MessageBox("Не указана конечная дата периода","Предупреждение",
                              MB_OK+MB_ICONINFORMATION);
      EditPO->SetFocus();
      Abort();
    }


      //Вывод определенного периода
  DM->qKomandirovki->Filtered=false;
  DM->qKomandirovki->Filter="data_n>="+QuotedStr(EditS->Text)+" and data_n<="+QuotedStr(EditPO->Text);
  DM->qKomandirovki->Filtered=true;

  DBGridEh1->SetFocus();

  EditS->Text="";
  EditPO->Text="";
}
//---------------------------------------------------------------------------

void __fastcall TMain::EditSExit(TObject *Sender)
{
  TDateTime d;

  if (!EditS->Text.IsEmpty())
    {
      // Добавление к дате отчетного месяца и года
      if (EditS->Text.Length()<3)
        {
          if(EditS->Text.Pos("."))
            {
              Application->MessageBox("Неверный формат даты","Ошибка", MB_OK+MB_ICONINFORMATION);
              EditS->Font->Color = clRed;
              EditS->SetFocus();
              Abort();
            }
          else
            {
              EditS->Text = EditS->Text+ "."+ DateToStr(Date()).SubString(4,2) +"."+ DateToStr(Date()).SubString(7,5);
              EditS->Font->Color = clBlack;
            }
        }

      // Проверка на правильность ввода даты
      if(!TryStrToDate(EditS->Text,d))
        {
          Application->MessageBox("Неверный формат даты","Ошибка", MB_OK);
          EditS->Font->Color = clRed;
          EditS->SetFocus();
        }
      else
        {
          EditS->Text=FormatDateTime("dd.mm.yyyy",d);
          EditS->Font->Color = clBlack;
        }

    }
}
//---------------------------------------------------------------------------

void __fastcall TMain::EditPOExit(TObject *Sender)
{
  TDateTime d;

  if (!EditPO->Text.IsEmpty())
    {
      // Добавление к дате отчетного месяца и года
      if (EditPO->Text.Length()<3)
        {
          if(EditPO->Text.Pos("."))
            {
              Application->MessageBox("Неверный формат даты","Ошибка", MB_OK+MB_ICONINFORMATION);
              EditPO->Font->Color = clRed;
              EditPO->SetFocus();
              Abort();
            }
          else
            {
              EditPO->Text = EditPO->Text+ "."+ DateToStr(Date()).SubString(4,2) +"."+ DateToStr(Date()).SubString(7,5);
              EditPO->Font->Color = clBlack;
            }
        }

      // Проверка на правильность ввода даты
      if(!TryStrToDate(EditPO->Text,d))
        {
          Application->MessageBox("Неверный формат даты","Ошибка", MB_OK);
          EditPO->Font->Color = clRed;
          EditPO->SetFocus();
        }
      else
        {
          EditPO->Text=FormatDateTime("dd.mm.yyyy",d);
          EditPO->Font->Color = clBlack;
        }

    }
}
//---------------------------------------------------------------------------
//Логи
void __fastcall TMain::InsertLog(AnsiString Msg)
{
  AnsiString Data;
  DomainName="MMK";
  UserName="Лена";

  DateTimeToString(Data, "dd.mm.yyyy hh:nn:ss", Now());
  DM->qObnovlenie->SQL->Clear();
  DM->qObnovlenie->SQL->Add("insert into logs (DT,DOMAIN,USERA, PROG, TEXT) values \
                            (to_date(" + QuotedStr(Data) + ", 'DD.MM.YYYY HH24:MI:SS'),\
                             " + QuotedStr(DomainName) + "," + QuotedStr(UserName) + ", 'Komandirovki',\
                             " + QuotedStr(Msg) +")");
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
//Командировочные расходы по тем работникам, которые проходят только через бухгалтерию
void __fastcall TMain::N19Click(TObject *Sender)
{
 AnsiString sFile, Sql;
  int n=6;
  Variant AppEx, Sh;

  StatusBar1->SimpleText=" Идет формирование отчета по командировочным расходам в Excel...";

  Sql="select s.zex as zex, s.tab as tn, fio, s.n_doc, punkt, datnkom||' - '||datkkom as dati, d.data as dato,                                                                \
              sut, kvart, avia, gd, proch, itgsum                                                                                                      \
       from  (                                                                                                                                         \
              (select n_doc,                                                                                                                           \
                      decode(sum(nvl(sut,0))+sum(nvl(sut_bez,0)),0,NULL,sum(nvl(sut,0))+sum(nvl(sut_bez,0))) as sut,                                                                           \
                      decode(sum(nvl(kvart,0))+sum(nvl(kvart_bez,0)),0,NULL,sum(nvl(kvart,0))+sum(nvl(kvart_bez,0))) as kvart,                                                                     \
                      decode(sum(nvl(sum,0))+sum(nvl(proez,0))+sum(nvl(stop,0))+sum(nvl(viza_bez,0)),0,NULL,(sum(nvl(sum,0))+sum(nvl(proez,0))+sum(nvl(stop,0))+sum(nvl(viza_bez,0)))) as proch, \
                      decode(sum(nvl(gd,0))+sum(nvl(gd_bez,0)),0,NULL,sum(nvl(gd,0))+sum(nvl(gd_bez,0))) as gd,                                                                              \
                      decode(sum(nvl(avia,0))+sum(nvl(avia_bez,0)),0,NULL,sum(nvl(avia,0))+sum(nvl(avia_bez,0))) as avia,                                                                        \
                      data                                                                                                                             \
               from k_avans2@F                                                                                                                         \
               group by n_doc, data) d                                                                                                                 \
              left join                                                                                                                                \
               (select n_doc, zex, tab, punkt, initcap(fio) fio, datnkom, datkkom, n_order, data, itgsum from k_avans1@F) s                            \
              on d.n_doc=s.n_doc and d.data=s.data                                                                                                     \
             )                                                                                                                                         \
       where (s.zex, s.tab, s.datnkom, s.datkkom) not in (select zex, tn, data_n, data_k from komandirovki)                                            \
       and d.data between '01."+(mm<10? "0"+IntToStr(mm): IntToStr(mm))+"."+IntToStr(yyyy)+"'  \
                and '"+DateToStr(EndOfTheMonth(StrToDate("01."+(mm<10? "0"+IntToStr(mm): IntToStr(mm))+"."+IntToStr(yyyy)))).SubString(1,2)+"."+(mm<10? "0"+IntToStr(mm): IntToStr(mm))+"."+IntToStr(yyyy)+"'                                                                                                  \
       order by zex,tab,n_doc";


      
       /*and d.data between '01."+DateToStr(Date()).SubString(4,2)+"."+DateToStr(Date()).SubString(7,4)+"' \
                and '"+DateToStr(EndOfTheMonth(Date())).SubString(1,2)+"."+DateToStr(Date()).SubString(4,2)+"."+DateToStr(Date()).SubString(7,4)+"'                                                                                                  \
       order by zex,tab,n_doc";*/



  DM->qObnovlenie->Close();
  DM->qObnovlenie->SQL->Clear();
  DM->qObnovlenie->SQL->Add(Sql);
  try
    {
      DM->qObnovlenie->Open();
    }
  catch(Exception &E)
    {
      Application->MessageBox(("Возникла ошибка при получении данных из таблицы по командировкам (KOMANDIROVKI)" + E.Message).c_str(),"Ошибка",
                              MB_OK+MB_ICONERROR);
     // InsertLog("Возникла ошибка при формировании списка работников по цехам в Excel");
     // DM->qLogs->Requery();
      StatusBar1->SimpleText="";
      Abort();
    }

  
  if (DM->qObnovlenie->RecordCount==0)
    {
      Application->MessageBox("Нет данных за отчетный период!!!","Предупреждение",
                               MB_OK+MB_ICONINFORMATION);
      StatusBar1->SimpleText="";
      Abort();
    }


  sFile = Path+"\\RTF\\komandir_rashod2.xlsx";

  //Создание папки, если ее не существует
  ForceDirectories(WorkPath);


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
      StatusBar1->SimpleText="";
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
              StatusBar1->SimpleText="";
              ProgressBar->Visible = false;
              Cursor = crDefault;
            }

          int i=1;
          n=5;

          Variant Massiv;
          Massiv = VarArrayCreate(OPENARRAY(int,(0,14)),varVariant); //массив на 14 элементов

          Sh.OlePropertyGet("Range", "I2").OlePropertySet("Value",Date());

          while (!DM->qObnovlenie->Eof)
            {
              Massiv.PutElement(i, 0);
              Massiv.PutElement(DM->qObnovlenie->FieldByName("zex")->AsString.c_str(), 1);
              Massiv.PutElement(DM->qObnovlenie->FieldByName("tn")->AsString.c_str(), 2);
              Massiv.PutElement(DM->qObnovlenie->FieldByName("fio")->AsString.c_str(), 3);
              Massiv.PutElement(DM->qObnovlenie->FieldByName("n_doc")->AsString.c_str(), 4);
              Massiv.PutElement(DM->qObnovlenie->FieldByName("punkt")->AsString.c_str(), 5);
              Massiv.PutElement(DM->qObnovlenie->FieldByName("dati")->AsString.c_str(), 6);
              Massiv.PutElement(DM->qObnovlenie->FieldByName("dato")->AsString.c_str(), 7);
              Massiv.PutElement(DM->qObnovlenie->FieldByName("sut")->AsString.c_str(), 8);
              Massiv.PutElement(DM->qObnovlenie->FieldByName("kvart")->AsString.c_str(), 9);
              Massiv.PutElement(DM->qObnovlenie->FieldByName("avia")->AsString.c_str(), 10);
              Massiv.PutElement(DM->qObnovlenie->FieldByName("gd")->AsString.c_str(), 11);
              Massiv.PutElement(DM->qObnovlenie->FieldByName("proch")->AsString.c_str(), 12);
              Massiv.PutElement(DM->qObnovlenie->FieldByName("itgsum")->AsString.c_str(), 13);

              Sh.OlePropertyGet("Range", ("A" + IntToStr(n) + ":N" + IntToStr(n)).c_str()).OlePropertySet("Value", Massiv); //строка с данными с ячейки A по ячейку АВ

              i++;
              n++;
              DM->qObnovlenie->Next();
              ProgressBar->Position++;
            }

          // вставляем в шаблон нужное количество строк

          //окрашивание ячеек
         // Sh.OlePropertyGet("Range",("B7:B"+IntToStr(n-1)).c_str()).OlePropertyGet("Interior").OlePropertySet("Color",14277081);
         // Sh.OlePropertyGet("Range",("L7:L"+IntToStr(n-1)).c_str()).OlePropertyGet("Interior").OlePropertySet("Color",14277081);

          //рисуем сетку
          Sh.OlePropertyGet("Range",("A5:N"+IntToStr(n-1)).c_str()).OlePropertyGet("Borders").OlePropertySet("LineStyle", xlContinuous);

          //Отключить вывод сообщений с вопросами типа "Заменить файл..."
          AppEx.OlePropertySet("DisplayAlerts",false);


          //Сохранить книгу в папке в файле по указанию
          AnsiString vAsCurDir1=WorkPath+"\\Реестр командировок для работников оформляющихся только через бухгалтерию.xlsx";
          Sh.OleProcedure("SaveAs",vAsCurDir1.c_str());

          //Закрыть открытое приложение Excel
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
          StatusBar1->SimpleText= "";
        }
      catch (...)
        {
          AppEx.OleProcedure("Quit");
          AppEx = Unassigned;
          Cursor = crDefault;
          ProgressBar->Position=0;
          StatusBar1->SimpleText= "";
          ProgressBar->Visible=false;
          InsertLog("Возникла ошибка при формировании отчета по коммандировочным расходам работников проходящих через бухгалтерию в Excel");
          Abort();
        }
    }
}
//---------------------------------------------------------------------------

//Реестр командировок по работникам, проходящим только через бухгалтерию
void __fastcall TMain::N20Click(TObject *Sender)
{
  AnsiString sFile, Sql;
  int n=7;
  Variant AppEx, Sh;


  StatusBar1->SimpleText=" Идет формирование реестра командировок в Excel...";


  if (Main->EditS->Text.IsEmpty() || Main->EditPO->Text.IsEmpty())
    {
  Sql="select * from                                                                                                            \
          ((select initcap(fio) as fio,                                                                                         \
                   zex,                                                                                                         \
                   tab,                                                                                                         \
                   (select grade from avans where ncex=zex and substr(lpad(tn,5,'0'),2,4)=substr(lpad(tab,5,'0'),2,4)) as grade,\
                   n_doc,                                                                                                       \
                   datnkom||' - '||datkkom as dati,                                                                             \
                   punkt, data,                                                                                                 \
                   datnkom, datkkom                                                                                             \
            from k_avans1@F) n                                                                                                  \
           left join                                                                                                            \
           (select n_doc, data,                                                                                                 \
                   decode(sum(nvl(avia,0))+sum(nvl(avia_bez,0)),0,NULL,sum(nvl(avia,0))+sum(nvl(avia_bez,0))) as avia,                                                    \
                   decode(sum(nvl(gd,0))+sum(nvl(gd_bez,0)),0,NULL,sum(nvl(gd,0))+sum(nvl(gd_bez,0))) as gd,                                                          \
                   decode(sum(nvl(kvart,0))+sum(nvl(kvart_bez,0)),0,NULL,sum(nvl(kvart,0))+sum(nvl(kvart_bez,0))) as kvart                                                  \
            from k_avans2@F                                                                                                     \
            group by n_doc, data) d                                                                                             \
           on n.n_doc=d.n_doc and n.data=d.data)                                                                                \
       where (n.zex, n.tab, n.datnkom, n.datkkom) not in (select zex, tn, data_n, data_k from komandirovki)                     \
       and (kvart is not null or avia is not null or gd is not null)                                                            \
       and (datnkom between '01."+(mm<10? "0"+IntToStr(mm): IntToStr(mm))+"."+IntToStr(yyyy)+"' \
                and '"+DateToStr(EndOfTheMonth(StrToDate("01."+(mm<10? "0"+IntToStr(mm): IntToStr(mm))+"."+IntToStr(yyyy)))).SubString(1,2)+"."+(mm<10? "0"+IntToStr(mm): IntToStr(mm))+"."+IntToStr(yyyy)+"'  \
       or datkkom between '01."+(mm<10? "0"+IntToStr(mm): IntToStr(mm))+"."+IntToStr(yyyy)+"' \
                and '"+DateToStr(EndOfTheMonth(StrToDate("01."+(mm<10? "0"+IntToStr(mm): IntToStr(mm))+"."+IntToStr(yyyy)))).SubString(1,2)+"."+(mm<10? "0"+IntToStr(mm): IntToStr(mm))+"."+IntToStr(yyyy)+"'  )\
       order by n.zex,n.tab,n.n_doc,avia";
     }
   else
     {
         Sql="select * from                                                                                                            \
          ((select initcap(fio) as fio,                                                                                         \
                   zex,                                                                                                         \
                   tab,                                                                                                         \
                   (select grade from avans where ncex=zex and substr(lpad(tn,5,'0'),2,4)=substr(lpad(tab,5,'0'),2,4)) as grade,\
                   n_doc,                                                                                                       \
                   datnkom||' - '||datkkom as dati,                                                                             \
                   punkt, data,                                                                                                 \
                   datnkom, datkkom                                                                                             \
            from k_avans1@F) n                                                                                                  \
           left join                                                                                                            \
           (select n_doc, data,                                                                                                 \
                   decode(sum(nvl(avia,0))+sum(nvl(avia_bez,0)),0,NULL,sum(nvl(avia,0))+sum(nvl(avia_bez,0))) as avia,                                                    \
                   decode(sum(nvl(gd,0))+sum(nvl(gd_bez,0)),0,NULL,sum(nvl(gd,0))+sum(nvl(gd_bez,0))) as gd,                                                          \
                   decode(sum(nvl(kvart,0))+sum(nvl(kvart_bez,0)),0,NULL,sum(nvl(kvart,0))+sum(nvl(kvart_bez,0))) as kvart                                                  \
            from k_avans2@F                                                                                                     \
            group by n_doc, data) d                                                                                             \
           on n.n_doc=d.n_doc and n.data=d.data)                                                                                \
       where (n.zex, n.tab, n.datnkom, n.datkkom) not in (select zex, tn, data_n, data_k from komandirovki)                     \
       and (kvart is not null or avia is not null or gd is not null)                                                            \
       and (datnkom between '01."+(mm<10? "0"+IntToStr(mm): IntToStr(mm))+"."+IntToStr(yyyy)+"' \
                and '"+DateToStr(EndOfTheMonth(StrToDate("01."+(mm<10? "0"+IntToStr(mm): IntToStr(mm))+"."+IntToStr(yyyy)))).SubString(1,2)+"."+(mm<10? "0"+IntToStr(mm): IntToStr(mm))+"."+IntToStr(yyyy)+"'  \
       or datkkom between '01."+(mm<10? "0"+IntToStr(mm): IntToStr(mm))+"."+IntToStr(yyyy)+"' \
                and '"+DateToStr(EndOfTheMonth(StrToDate("01."+(mm<10? "0"+IntToStr(mm): IntToStr(mm))+"."+IntToStr(yyyy)))).SubString(1,2)+"."+(mm<10? "0"+IntToStr(mm): IntToStr(mm))+"."+IntToStr(yyyy)+"'  )\
       order by n.zex,n.tab,n.n_doc,avia";
     }
     



       /*   and (datnkom between '01."+DateToStr(Date()).SubString(4,2)+"."+DateToStr(Date()).SubString(7,4)+"' \
                and '"+DateToStr(EndOfTheMonth(Date())).SubString(1,2)+"."+DateToStr(Date()).SubString(4,2)+"."+DateToStr(Date()).SubString(7,4)+"'  \
       or datkkom between '01."+DateToStr(Date()).SubString(4,2)+"."+DateToStr(Date()).SubString(7,4)+"' \
                and '"+DateToStr(EndOfTheMonth(Date())).SubString(1,2)+"."+DateToStr(Date()).SubString(4,2)+"."+DateToStr(Date()).SubString(7,4)+"'  )\
       order by n.zex,n.tab,n.n_doc,avia";   */


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
      Application->MessageBox(("Возникла ошибка при получении данных из таблицы по командировкам (KOMANDIROVKI)" + E.Message).c_str(),"Ошибка",
                              MB_OK+MB_ICONERROR);
     // InsertLog("Возникла ошибка при формировании списка работников по цехам в Excel");
     // DM->qLogs->Requery();
      StatusBar1->SimpleText="";
      Abort();
    }

  if (DM->qObnovlenie->RecordCount==0)
    {
      Application->MessageBox("Нет данных за отчетный период!!!","Предупреждение",
                               MB_OK+MB_ICONINFORMATION);
      StatusBar1->SimpleText="";
      Abort();
    }

  sFile = Path+"\\RTF\\reestr_komandir.xlsx";

  //Создание папки, если ее не существует
  ForceDirectories(WorkPath);


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
      StatusBar1->SimpleText="";
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
              StatusBar1->SimpleText="";
              ProgressBar->Visible = false;
              Cursor = crDefault;
            }

          //AppEx.OlePropertySet("Visible",true);

          int i=1;
          n=7;

          Variant Massiv, Massiv2;
          Massiv = VarArrayCreate(OPENARRAY(int,(0,17)),varVariant); //массив на 16 элементов
          Massiv2 = VarArrayCreate(OPENARRAY(int,(0,4)),varVariant); //массив на 3 элементов


         //char z=Mes[MonthOf(Date())-1];

          /*char f = YearOf(Date());
          Sh.OlePropertyGet("Range", "I3").OlePropertySet("Value",f);
            */

          Sh.OlePropertyGet("Range", "I3").OlePropertySet("Value", MonthOf(Date()));
          Sh.OlePropertyGet("Range", "J3").OlePropertySet("Value", YearOf(Date()));
          Sh.OlePropertyGet("Range", "I4").OlePropertySet("Value","по работникам, оформляющимся только через бухгалтерию");

          while (!DM->qObnovlenie->Eof)
            {
              Massiv.PutElement(i, 0);
              Massiv.PutElement(DM->qObnovlenie->FieldByName("fio")->AsString.c_str(), 1);
              Massiv.PutElement("", 2);
              Massiv.PutElement(DM->qObnovlenie->FieldByName("zex")->AsString.c_str(), 4);
              Massiv.PutElement(DM->qObnovlenie->FieldByName("tab")->AsString.c_str(), 5);
              Massiv.PutElement(DM->qObnovlenie->FieldByName("grade")->AsString.c_str(), 6);
              Massiv.PutElement(DM->qObnovlenie->FieldByName("n_doc")->AsString.c_str(), 7);
              Massiv.PutElement(DM->qObnovlenie->FieldByName("dati")->AsString.c_str(), 8);
              Massiv.PutElement(DM->qObnovlenie->FieldByName("punkt")->AsString.c_str(), 9);
              Massiv.PutElement("", 10);
              Massiv.PutElement("", 11);
              Massiv.PutElement("", 12);
              Massiv.PutElement(DM->qObnovlenie->FieldByName("kvart")->AsString.c_str(), 13);

              Sh.OlePropertyGet("Range", ("A" + IntToStr(n) + ":Q" + IntToStr(n)).c_str()).OlePropertySet("Value", Massiv); //строка с данными с ячейки A по ячейку АВ

              if (!DM->qObnovlenie->FieldByName("avia")->AsString.IsEmpty())
                {

                  Massiv2.PutElement("Авиа", 0);
                  Massiv2.PutElement("", 1);
                  Massiv2.PutElement(DM->qObnovlenie->FieldByName("avia")->AsString.c_str(), 2);

                  Sh.OlePropertyGet("Range", ("O" + IntToStr(n) + ":Q" + IntToStr(n)).c_str()).OlePropertySet("Value", Massiv2); //строка с данными с ячейки A по ячейку АВ

                  i++;
                  n++;
                }
              if (!DM->qObnovlenie->FieldByName("gd")->AsString.IsEmpty())
                {
                  Massiv2.PutElement("Ж/д", 0);
                  Massiv2.PutElement("", 1);
                  Massiv2.PutElement(DM->qObnovlenie->FieldByName("gd")->AsString.c_str(), 2);

                  Sh.OlePropertyGet("Range", ("O" + IntToStr(n) + ":Q" + IntToStr(n)).c_str()).OlePropertySet("Value", Massiv2); //строка с данными с ячейки A по ячейку АВ

                  i++;
                  n++;
                }
             

              DM->qObnovlenie->Next();
              ProgressBar->Position++;
            }

          // вставляем в шаблон нужное количество строк

          //окрашивание ячеек
     /*     Sh.OlePropertyGet("Range",("M18:M"+IntToStr(n-1)).c_str()).OlePropertyGet("Interior").OlePropertySet("Color",14277081);
          Sh.OlePropertyGet("Range",("P18:R"+IntToStr(n-1)).c_str()).OlePropertyGet("Interior").OlePropertySet("Color",14277081);

          Sh.OlePropertyGet("Range",("B18:K"+IntToStr(n-1)).c_str()).OlePropertyGet("Interior").OlePropertySet("Color",14408946);
          Sh.OlePropertyGet("Range",("N18:N"+IntToStr(n-1)).c_str()).OlePropertyGet("Interior").OlePropertySet("Color",14408946);
       */
          //рисуем сетку
          Sh.OlePropertyGet("Range",("A7:Q"+IntToStr(n-1)).c_str()).OlePropertyGet("Borders").OlePropertySet("LineStyle", xlContinuous);


          //Отключить вывод сообщений с вопросами типа "Заменить файл..."
          AppEx.OlePropertySet("DisplayAlerts",false);


          //Сохранить книгу в папке в файле по указанию
          AnsiString vAsCurDir1=WorkPath+"\\Реестр командировок.xlsx";
          Sh.OleProcedure("SaveAs",vAsCurDir1.c_str());

          //Закрыть открытое приложение Excel
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
          StatusBar1->SimpleText= "";
        }
      catch (...)
        {
          AppEx.OleProcedure("Quit");
          AppEx = Unassigned;
          Cursor = crDefault;
          ProgressBar->Position=0;
          StatusBar1->SimpleText= "";
          ProgressBar->Visible=false;
         // InsertLog("Возникла ошибка при формировании списка работников по цехам в Excel");
          Abort();
        }
    }
}
//---------------------------------------------------------------------------

