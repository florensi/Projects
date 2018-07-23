//---------------------------------------------------------------------------

#include <vcl.h>
#pragma hdrstop

#include "uSprav.h"
#include "uDM.h"
#include "uVvod.h"
#include "uMain.h"
//---------------------------------------------------------------------------
#pragma package(smart_init)
#pragma link "DBGridEh"
#pragma resource "*.dfm"
TSprav *Sprav;
//---------------------------------------------------------------------------
__fastcall TSprav::TSprav(TComponent* Owner)
        : TForm(Owner)
{
}
//---------------------------------------------------------------------------
void __fastcall TSprav::Ljfdbnmpfgbcm1Click(TObject *Sender)
{
  Panel1->Visible=true;
  fl_sp_red=0;

  EditCHEL->Text="";
  Vvod->Label14->Caption="Добавление записи";
  DBGridEh1->Top=139;
  DBGridEh1->Height=323;

  EditCHEL->SetFocus();
}
//---------------------------------------------------------------------------

void __fastcall TSprav::N1RedaktClick(TObject *Sender)
{
  Panel1->Visible=true;
  fl_sp_red=1;

  EditCHEL->Text=DM->qSP_chel->FieldByName("naim")->AsString;
  Vvod->Label14->Caption="Редактирование записи";
  DBGridEh1->Top=139;
  DBGridEh1->Height=323;

  EditCHEL->SetFocus();
}
//---------------------------------------------------------------------------


void __fastcall TSprav::N4Click(TObject *Sender)
{
  Panel4->Visible=true;

  fl_sp_red=0;

  EditGRADE->Text="";
  EditKAT->Text="";
  EditVAGON->Text="";
  EditG_MIN_KIEV->Text="";
  EditG_KIEV->Text="";
  EditG_MIN_UKR->Text="";
  EditG_UKR->Text="";
  EditG_ZAGRAN->Text="";
  Label2->Caption="Добавление записи";
  DBGridEh2->Top=116;
  DBGridEh2->Height=350;

  EditGRADE->SetFocus();
 }
//---------------------------------------------------------------------------

void __fastcall TSprav::N5Click(TObject *Sender)
{
  Panel4->Visible=true;
  fl_sp_red=1;


  EditGRADE->Text=DM->qSP_grade->FieldByName("grade")->AsString;
  EditKAT->Text=DM->qSP_grade->FieldByName("kat")->AsString;
  EditVAGON->Text=DM->qSP_grade->FieldByName("vagon")->AsString;
  EditG_MIN_KIEV->Text=DM->qSP_grade->FieldByName("g_min_kiev")->AsString;
  EditG_KIEV->Text=DM->qSP_grade->FieldByName("g_kiev")->AsString;
  EditG_MIN_UKR->Text=DM->qSP_grade->FieldByName("g_min_ukr")->AsString;
  EditG_UKR->Text=DM->qSP_grade->FieldByName("g_ukr")->AsString;
  EditG_ZAGRAN->Text=DM->qSP_grade->FieldByName("g_zagran")->AsString;


  DBGridEh2->Top=143;
  DBGridEh2->Height=458;

  EditGRADE->SetFocus();
}
//---------------------------------------------------------------------------

void __fastcall TSprav::N8Click(TObject *Sender)
{
  Panel7->Visible=true;
  fl_sp_red=0;

  EditGOROD->Text="";
  EditGOSTINICA->Text="";
  EditGOST_ADR->Text="";

  DBGridEh3->Top=134;
  DBGridEh3->Height=335;

  EditGOROD->SetFocus();
}
//---------------------------------------------------------------------------

void __fastcall TSprav::N9Click(TObject *Sender)
{
  Panel7->Visible=true;
  fl_sp_red=1;

  EditGOROD->Text=DM->qSP_gostinica->FieldByName("city")->AsString;
  EditGOSTINICA->Text=DM->qSP_gostinica->FieldByName("gostinica")->AsString;
  EditGOST_ADR->Text=DM->qSP_gostinica->FieldByName("adress")->AsString;

  DBGridEh3->Top=134;
  DBGridEh3->Height=335;

  EditGOROD->SetFocus();
}
//---------------------------------------------------------------------------

void __fastcall TSprav::N12Click(TObject *Sender)
{
  Panel10->Visible=true;
  fl_sp_red=0;

  EditGOROD1->Text="";
  EditOBEKT->Text="";
  EditADRESS->Text="";

  DBGridEh4->Top=131;
  DBGridEh4->Height=344;

  EditGOROD1->SetFocus();
}
//---------------------------------------------------------------------------

void __fastcall TSprav::N13Click(TObject *Sender)
{
  Panel10->Visible=true;
  fl_sp_red=1;


  EditGOROD1->Text=DM->qSP_obekt->FieldByName("city")->AsString;
  EditOBEKT->Text=DM->qSP_obekt->FieldByName("obekt")->AsString;
  EditADRESS->Text=DM->qSP_obekt->FieldByName("adress")->AsString;

  DBGridEh4->Top=131;
  DBGridEh4->Height=344;

  EditGOROD1->SetFocus();
}
//---------------------------------------------------------------------------

void __fastcall TSprav::DBGridEh2DrawColumnCell(TObject *Sender,
      const TRect &Rect, int DataCol, TColumnEh *Column,
      TGridDrawState State)
{
  // выделение цветом активной записи
  if (State.Contains(gdSelected) )
    {
      ((TDBGridEh *) Sender)->Canvas->Brush->Color = clSkyBlue;//(TColor)0x00DEF5F4;//clInfoBk;
    }
  ((TDBGridEh *) Sender)->Canvas->Font->Color = clBlack;
  ((TDBGridEh *) Sender)->DefaultDrawColumnCell(Rect, DataCol, Column, State);
}
//---------------------------------------------------------------------------

void __fastcall TSprav::FormShow(TObject *Sender)
{
  Sprav->TabSheet1->Caption = "Цели командировки";
  Sprav->TabSheet2->Caption = "Грейды";
  Sprav->TabSheet3->Caption = "Гостиницы";
  Sprav->TabSheet4->Caption = "Объекты";
  Sprav->TabSheet5->Caption = "Страны";
  Sprav->TabSheet6->Caption = "Города";

  Panel16->Visible=false;
  Panel1->Visible=false;
  Panel4->Visible=false;
 // DBGridEh2->Height=457;
 // DBGridEh2->Top=143;
  Panel7->Visible=false;
  Panel10->Visible=false;
  Panel13->Visible=false;
  

  PageControl1->OwnerDraw = true;
}
//---------------------------------------------------------------------------

void __fastcall TSprav::BitBtn2Click(TObject *Sender)
{
  Panel4->Visible=false;
  DBGridEh2->Top=143;
  DBGridEh2->Height=458;
}
//---------------------------------------------------------------------------

void __fastcall TSprav::FormKeyDown(TObject *Sender, WORD &Key,
      TShiftState Shift)
{
  if (Key==VK_RETURN)
  FindNextControl((TWinControl *)Sender, true, true, false)->SetFocus();
}
//---------------------------------------------------------------------------

//Сохранение грейда
void __fastcall TSprav::BitBtn1Click(TObject *Sender)
{
  TLocateOptions SearchOptions;
  AnsiString Sql;
  int rec;

  //Проверка на введенный грейд
  if (EditGRADE->Text.IsEmpty())
    {
      Application->MessageBox("Не указан грейд!!!","Предупреждение",
                               MB_OK+MB_ICONINFORMATION);
      EditGRADE->SetFocus();
    }

  //Проверка на наличие записи с таким грейдом
  if (DM->qSP_grade->Locate("grade",EditGRADE->Text,SearchOptions << loCaseInsensitive))
    {
      Application->MessageBox("Введенный грейд уже есть в справочнике","Предупреждение",
                              MB_OK+MB_ICONINFORMATION);
      EditGRADE->SetFocus();
    }

  //Проверка на введеную категорию
  if (EditKAT->Text.IsEmpty())
    {
      Application->MessageBox("Не указана категория грейда!!!","Предупреждение",
                               MB_OK+MB_ICONINFORMATION);
      EditKAT->SetFocus();
    }

  //проверка на введенный вагон
  if (EditVAGON->Text.IsEmpty())
    {
      Application->MessageBox("Не указана разновидность вагона!!!","Предупреждение",
                               MB_OK+MB_ICONINFORMATION);
      EditVAGON->SetFocus();
    }

 //Проверка на ввод min суммы по Киеву
  if (EditG_MIN_KIEV->Text.IsEmpty())
    {
      Application->MessageBox("Не указана минимальная сумма по Киеву!!!","Предупреждение",
                               MB_OK+MB_ICONINFORMATION);
      EditG_MIN_KIEV->SetFocus();
    }

 //Проверка на ввод max суммы по Киеву
  if (EditG_KIEV->Text.IsEmpty())
    {
      Application->MessageBox("Не указана максимальная сумма по Киеву!!!","Предупреждение",
                               MB_OK+MB_ICONINFORMATION);
      EditG_KIEV->SetFocus();
    }
  //Проверка, чтоб min сумма не была больше max по Киеву
  if (StrToFloat(EditG_MIN_KIEV->Text)>StrToFloat(EditG_KIEV->Text))
    {
      Application->MessageBox("Мнимальная сумма по Киеву превышает максимальную!!!","Предупреждение",
                               MB_OK+MB_ICONINFORMATION);
      EditG_MIN_KIEV->SetFocus();
    }

  //Проверка на ввод min суммы по Украине
  if (EditG_MIN_UKR->Text.IsEmpty())
    {
      Application->MessageBox("Не указана минимальная сумма по Украине!!!","Предупреждение",
                               MB_OK+MB_ICONINFORMATION);
      EditG_MIN_UKR->SetFocus();
    }

  //Проверка на ввод max суммы по Украине
  if (EditG_UKR->Text.IsEmpty())
    {
      Application->MessageBox("Не указана максимальная сумма по Украине!!!","Предупреждение",
                               MB_OK+MB_ICONINFORMATION);
      EditG_UKR->SetFocus();
    }

  //Проверка, чтоб min сумма не была больше max по Киеву
   if (StrToFloat(EditG_MIN_UKR->Text)>StrToFloat(EditG_UKR->Text))
    {
      Application->MessageBox("Мнимальная сумма по Украине превышает максимальную!!!","Предупреждение",
                               MB_OK+MB_ICONINFORMATION);
      EditG_MIN_UKR->SetFocus();
    }

  //Проверка на ввод суммы по поездкам за рубеж
  if (EditG_ZAGRAN->Text.IsEmpty())
    {
      Application->MessageBox("Не указана сумма для заграницы!!!","Предупреждение",
                               MB_OK+MB_ICONINFORMATION);
      EditG_ZAGRAN->SetFocus();
    }


  //Добавление записи
  if (fl_sp_red==0)
    {
       Sql="insert into sp_grade (grade, kat, g_min_kiev, g_kiev, \
                                  g_min_ukr, g_ukr, g_zagran, vagon)\
           values ("\
                   +EditGRADE->Text+","
                   +QuotedStr(EditKAT->Text)+","
                   +EditG_MIN_KIEV->Text+","
                   +EditG_KIEV->Text+","
                   +EditG_MIN_UKR->Text+","
                   +EditG_UKR->Text+","
                   +EditG_ZAGRAN->Text+","
                   +QuotedStr(EditVAGON->Text)+")";
    }
  //Обновление записи
  else if (fl_sp_red==1)
    {
      Sql="update sp_grade set\
                               grade = "+EditGRADE->Text+",  \
                               kat="+QuotedStr(EditKAT->Text)+",        \
                               g_min_kiev="+EditG_MIN_KIEV->Text+", \
                               g_kiev="+EditG_KIEV->Text+",         \
                               g_min_ukr="+EditG_MIN_UKR->Text+",   \
                               g_ukr="+EditG_UKR->Text+",           \
                               g_zagran="+EditG_ZAGRAN->Text+",     \
                               vagon="+QuotedStr(EditVAGON->Text)+"            \
           where rowid=chartorowid("+QuotedStr(DM->qSP_grade->FieldByName("rw")->AsString)+")";

      rec = DM->qSP_grade->RecNo;
    }

  DM->qObnovlenie->Close();
  DM->qObnovlenie->SQL->Clear();
  DM->qObnovlenie->SQL->Add(Sql);
  try
    {
      DM->qObnovlenie->ExecSQL();
    }
  catch(Exception &E)
    {
      Application->MessageBox(("Невозможно добавить/обновить запись в справочнике грейдов (SP_grade)"+E.Message).c_str(),"Ошибка",
                              MB_OK+MB_ICONERROR);
      Panel4->Visible=false;
      DBGridEh2->Top=143;
      DBGridEh2->Height=458;
      Abort();
    }

  //Логи

  DM->qSP_grade->Requery();

  //Возвращение курсора на строку
  if (fl_sp_red==0)
    {
      // При добавлении записи возвращать на нее курсор
      DM->qSP_grade->Locate("grade",EditGRADE->Text,SearchOptions << loCaseInsensitive);
    }
  else
    {
      DM->qSP_grade->RecNo = rec;
    }


  Panel4->Visible=false;
  DBGridEh2->Top=143;
  DBGridEh2->Height=458;

}
//---------------------------------------------------------------------------

void __fastcall TSprav::BitBtn4Click(TObject *Sender)
{
  Panel1->Visible=false;
  DBGridEh1->Top=175;
  DBGridEh1->Height=425;
}
//---------------------------------------------------------------------------


void __fastcall TSprav::DBGridEh1DrawColumnCell(TObject *Sender,
      const TRect &Rect, int DataCol, TColumnEh *Column,
      TGridDrawState State)
{
  // выделение цветом активной записи
  if (State.Contains(gdSelected) )
    {
      ((TDBGridEh *) Sender)->Canvas->Brush->Color = clSkyBlue;//(TColor)0x00DEF5F4;//clInfoBk;
    }
  ((TDBGridEh *) Sender)->Canvas->Font->Color = clBlack;
  ((TDBGridEh *) Sender)->DefaultDrawColumnCell(Rect, DataCol, Column, State);
}
//---------------------------------------------------------------------------

void __fastcall TSprav::DBGridEh3DrawColumnCell(TObject *Sender,
      const TRect &Rect, int DataCol, TColumnEh *Column,
      TGridDrawState State)
{
  // выделение цветом активной записи
  if (State.Contains(gdSelected) )
    {
      ((TDBGridEh *) Sender)->Canvas->Brush->Color = clSkyBlue;//(TColor)0x00DEF5F4;//clInfoBk;
    }
  ((TDBGridEh *) Sender)->Canvas->Font->Color = clBlack;
  ((TDBGridEh *) Sender)->DefaultDrawColumnCell(Rect, DataCol, Column, State);        
}
//---------------------------------------------------------------------------

void __fastcall TSprav::DBGridEh4DrawColumnCell(TObject *Sender,
      const TRect &Rect, int DataCol, TColumnEh *Column,
      TGridDrawState State)
{
  // выделение цветом активной записи
  if (State.Contains(gdSelected) )
    {
      ((TDBGridEh *) Sender)->Canvas->Brush->Color = clSkyBlue;//(TColor)0x00DEF5F4;//clInfoBk;
    }
  ((TDBGridEh *) Sender)->Canvas->Font->Color = clBlack;
  ((TDBGridEh *) Sender)->DefaultDrawColumnCell(Rect, DataCol, Column, State);        
}
//---------------------------------------------------------------------------

void __fastcall TSprav::DBGridEh5DrawColumnCell(TObject *Sender,
      const TRect &Rect, int DataCol, TColumnEh *Column,
      TGridDrawState State)
{
  // выделение цветом активной записи
  if (State.Contains(gdSelected) )
    {
      ((TDBGridEh *) Sender)->Canvas->Brush->Color = clSkyBlue;//(TColor)0x00DEF5F4;//clInfoBk;
    }
  ((TDBGridEh *) Sender)->Canvas->Font->Color = clBlack;
  ((TDBGridEh *) Sender)->DefaultDrawColumnCell(Rect, DataCol, Column, State);        
}
//---------------------------------------------------------------------------

void __fastcall TSprav::DBGridEh6DrawColumnCell(TObject *Sender,
      const TRect &Rect, int DataCol, TColumnEh *Column,
      TGridDrawState State)
{
  // выделение цветом активной записи
  if (State.Contains(gdSelected) )
    {
      ((TDBGridEh *) Sender)->Canvas->Brush->Color = clSkyBlue;//(TColor)0x00DEF5F4;//clInfoBk;
    }
  ((TDBGridEh *) Sender)->Canvas->Font->Color = clBlack;
  ((TDBGridEh *) Sender)->DefaultDrawColumnCell(Rect, DataCol, Column, State);
}
//---------------------------------------------------------------------------

void __fastcall TSprav::BitBtn3Click(TObject *Sender)
{
  TLocateOptions SearchOptions;
  AnsiString Sql;
  int rec;

  //Проверка на заполнение поля
  if (EditCHEL->Text.IsEmpty())
    {
      Application->MessageBox("Не заполнено наименование цели командировки","Предупреждение",
                               MB_OK+MB_ICONINFORMATION);
      EditCHEL->SetFocus();
    }

  //Проверка на существование поля
  if (DM->qSP_chel->Locate("naim",EditCHEL->Text,SearchOptions << loCaseInsensitive))
    {
      Application->MessageBox("Введенная цель командировки уже есть в справочнике","Предупреждение",
                              MB_OK+MB_ICONINFORMATION);
      EditCHEL->SetFocus();
    }

  //Добавление записи
  if (fl_sp_red==0)
    {
      Sql="insert into sp_komandir (kod,naim)\
           values (\
                    (select max(kod)+1 from sp_komandir), "
                    +QuotedStr(EditCHEL->Text)+")";
    }
  //Редактирование записи
  else if (fl_sp_red==1)
    {
      Sql="update sp_komandir set\
                                   naim="+QuotedStr(EditCHEL->Text)+" \
           where kod="+DM->qSP_chel->FieldByName("kod")->AsString;

      rec = DM->qSP_chel->RecNo;
    }

  DM->qObnovlenie->Close();
  DM->qObnovlenie->SQL->Clear();
  DM->qObnovlenie->SQL->Add(Sql);
  try
    {
      DM->qObnovlenie->ExecSQL();
    }
  catch(Exception &E)
    {
      Application->MessageBox(("Невозможно добавить/обновить запись в справочнике целей командировки (SP_chel)"+E.Message).c_str(),"Ошибка",
                              MB_OK+MB_ICONERROR);
      Panel1->Visible=false;
      DBGridEh1->Top=175;
      DBGridEh1->Height=425;
      Abort();
    }

  //Логи

  DM->qSP_chel->Requery();

  //Возвращение курсора на строку
  if (fl_sp_red==0)
    {
      // При добавлении записи возвращать на нее курсор
      DM->qSP_chel->Locate("naim",EditCHEL->Text,SearchOptions << loCaseInsensitive);
    }
  else
    {
      DM->qSP_chel->RecNo = rec;
    }

  
  Panel1->Visible=false;
  DBGridEh1->Top=175;
  DBGridEh1->Height=425;
}
//---------------------------------------------------------------------------

void __fastcall TSprav::N3Click(TObject *Sender)
{
  if (Application->MessageBox("Вы действительно хотите удалить выбранную запись?","Удаление записи",
                          MB_YESNO+MB_ICONINFORMATION)==ID_NO)
    {
      Abort();
    }

  //Удаление записи
  DM->qObnovlenie->Close();
  DM->qObnovlenie->SQL->Clear();
  DM->qObnovlenie->SQL->Add("delete from sp_komandir where kod="+DM->qSP_chel->FieldByName("kod")->AsString);
  try
    {
      DM->qObnovlenie->ExecSQL();
    }
  catch(Exception &E)
    {
      Application->MessageBox(("Невозможно удалить запись из справочника целей командировки (SP_chel) "+E.Message).c_str(),"Ошибка",
                              MB_OK+MB_ICONERROR);
      Main->InsertLog("Не выполнено удаление записи из справочника целей: цель "+DM->qSP_chel->FieldByName("naim")->AsString);
      Abort();
    }


  //Логи
  Main->InsertLog("Выполнено удаление записи из справочника целей: цель "+DM->qSP_chel->FieldByName("naim")->AsString);

  DM->qSP_chel->Requery();

  Application->MessageBox("Запись успешно удалена","Удаление записи",
                          MB_OK+MB_ICONINFORMATION);

}
//---------------------------------------------------------------------------

void __fastcall TSprav::DBGridEh1DblClick(TObject *Sender)
{
  N1RedaktClick(Sender);        
}
//---------------------------------------------------------------------------

void __fastcall TSprav::EditGORODKeyPress(TObject *Sender, char &Key)
{
  if (IsNumeric(Key)) Key=0;         
}
//---------------------------------------------------------------------------

void __fastcall TSprav::BitBtn6Click(TObject *Sender)
{
  Panel7->Visible=false;
  DBGridEh3->Top=167;
  DBGridEh3->Height=435;
}
//---------------------------------------------------------------------------

//Удаление гостиницы
void __fastcall TSprav::N11Click(TObject *Sender)
{
  if (Application->MessageBox("Вы действительно хотите удалить выбранную запись?","Удаление записи",
                          MB_YESNO+MB_ICONINFORMATION)==ID_NO)
    {
      Abort();
    }

  //Удаление записи
  DM->qObnovlenie->Close();
  DM->qObnovlenie->SQL->Clear();
  DM->qObnovlenie->SQL->Add("delete from sp_gostinica where rowid=chartorowid("+QuotedStr(DM->qSP_gostinica->FieldByName("rw")->AsString)+")");
  try
    {
      DM->qObnovlenie->ExecSQL();
    }
  catch(Exception &E)
    {
      Application->MessageBox(("Невозможно удалить запись из справочника гостиниц (SP_GOSTINICA) "+E.Message).c_str(),"Ошибка",
                              MB_OK+MB_ICONERROR);
      Main->InsertLog("Не выполнено удаление записи из справочника гостиниц: гостиница "+DM->qSP_gostinica->FieldByName("gostinica")->AsString);
      Abort();
    }


  //Логи
  Main->InsertLog("Выполнено удаление записи из справочника гостиниц: гостиница "+DM->qSP_gostinica->FieldByName("gostinica")->AsString);

  DM->qSP_gostinica->Requery();

  Application->MessageBox("Запись успешно удалена","Удаление записи",
                          MB_OK+MB_ICONINFORMATION);
}
//---------------------------------------------------------------------------

//Добавление/изменение гостиницы
void __fastcall TSprav::BitBtn5Click(TObject *Sender)
{
  TLocateOptions SearchOptions;
  AnsiString Sql;
  int rec;

  //Проверка на введенный город
  if (EditGOROD->Text.IsEmpty())
    {
      Application->MessageBox("Не указан город!!!","Предупреждение",
                               MB_OK+MB_ICONINFORMATION);
      EditGOROD->SetFocus();
    }

  //Проверка на введеную гостиницу
  if (EditGOSTINICA->Text.IsEmpty())
    {
      Application->MessageBox("Не указана гостиница!!!","Предупреждение",
                               MB_OK+MB_ICONINFORMATION);
      EditGOSTINICA->SetFocus();
    }
  //Проверка на уже существующую запись
  if (DM->qSP_gostinica->Locate("gostinica",EditGOSTINICA->Text,SearchOptions << loCaseInsensitive))
    {
      Application->MessageBox("Введенная гостиница уже есть в справочнике","Предупреждение",
                              MB_OK+MB_ICONINFORMATION);
      EditGOSTINICA->SetFocus();
    }

  //проверка на введенный адрес
  if (EditGOST_ADR->Text.IsEmpty())
    {
      Application->MessageBox("Не указанадрес гостиницы!!!","Предупреждение",
                               MB_OK+MB_ICONINFORMATION);
      EditGOST_ADR->SetFocus();
    }

  //Добавление записи
  if (fl_sp_red==0)
    {
       Sql="insert into sp_gostinica (kod, gostinica, adress, kod_city, reit)\
           values ( \
                   (select max(kod)+1 from sp_gostinica),"                                        \
                   +QuotedStr(EditGOSTINICA->Text)+","                                                       \
                   +QuotedStr(EditGOST_ADR->Text)+",                                                           \
                   (select kod from sp_city where upper(city)=upper(trim("+QuotedStr(EditGOROD->Text)+"))),  \
                   0)";
    }
  //Обновление записи
  else if (fl_sp_red==1)
    {
      Sql="update sp_gostinica set\
                                gostinica="+QuotedStr(EditGOSTINICA->Text)+",         \
                                adress="+QuotedStr(EditGOST_ADR->Text)+",                                                   \
                                kod_city=(select kod from sp_city where upper(city)=upper(trim("+QuotedStr(EditGOROD->Text)+")))   \
           where rowid=chartorowid("+QuotedStr(DM->qSP_gostinica->FieldByName("rw")->AsString)+")";

      rec = DM->qSP_gostinica->RecNo;
    }

  DM->qObnovlenie->Close();
  DM->qObnovlenie->SQL->Clear();
  DM->qObnovlenie->SQL->Add(Sql);
  try
    {
      DM->qObnovlenie->ExecSQL();
    }
  catch(Exception &E)
    {
      Application->MessageBox(("Невозможно добавить/обновить запись в справочнике гостиниц (SP_gostinica)"+E.Message).c_str(),"Ошибка",
                              MB_OK+MB_ICONERROR);
      Panel7->Visible=false;
      DBGridEh3->Top=167;
      DBGridEh3->Height=435;
      Abort();
    }

  //Логи

  DM->qSP_gostinica->Requery();

  //Возвращение курсора на строку
  if (fl_sp_red==0)
    {
      // При добавлении записи возвращать на нее курсор
      DM->qSP_gostinica->Locate("gostinica",EditGOSTINICA->Text,SearchOptions << loCaseInsensitive);
    }
  else
    {
      DM->qSP_gostinica->RecNo = rec;
    }


  Panel7->Visible=false;
  DBGridEh3->Top=167;
  DBGridEh3->Height=435;

}
//---------------------------------------------------------------------------

void __fastcall TSprav::BitBtn8Click(TObject *Sender)
{
  Panel10->Visible=false;
  DBGridEh4->Top=163;
  DBGridEh4->Height=435;
}
//---------------------------------------------------------------------------

void __fastcall TSprav::BitBtn10Click(TObject *Sender)
{
  Panel16->Visible=false;
  DBGridEh5->Top=166;
  DBGridEh5->Height=437;
}
//---------------------------------------------------------------------------

void __fastcall TSprav::BitBtn12Click(TObject *Sender)
{
  Panel13->Visible=false;
  DBGridEh6->Top=160;
  DBGridEh6->Height=437;
}
//---------------------------------------------------------------------------

void __fastcall TSprav::N1Click(TObject *Sender)
{
  Panel16->Visible=true;
  fl_sp_red=0;

  EditKOD->Text="";
  EditCOUNTRY->Text="";
  EditCOUNTRY_K->Text="";

  DBGridEh5->Top=132;
  DBGridEh5->Height=332;

  EditKOD->SetFocus();
}
//---------------------------------------------------------------------------

void __fastcall TSprav::N18Click(TObject *Sender)
{
  Panel13->Visible=true;
  fl_sp_red=0;

  EditCOUNTRY2->Text="";
  EditGOROD2->Text="";

  DBGridEh6->Top=132;
  DBGridEh6->Height=332;

  EditCOUNTRY2->SetFocus();
}
//---------------------------------------------------------------------------

void __fastcall TSprav::N2Click(TObject *Sender)
{
  Panel16->Visible=true;
  fl_sp_red=1;


  EditKOD->Text=DM->qSP_country->FieldByName("kod")->AsString;
  EditCOUNTRY->Text=DM->qSP_country->FieldByName("country")->AsString;
  EditCOUNTRY_K->Text=DM->qSP_country->FieldByName("country_k")->AsString;

  DBGridEh5->Top=132;
  DBGridEh5->Height=332;

  EditKOD->SetFocus();
}
//---------------------------------------------------------------------------

void __fastcall TSprav::N19Click(TObject *Sender)
{
  Panel13->Visible=true;
  fl_sp_red=1;


  EditCOUNTRY2->Text=DM->qSP_city->FieldByName("country")->AsString;
  EditGOROD2->Text=DM->qSP_city->FieldByName("city")->AsString;

  DBGridEh6->Top=128;
  DBGridEh6->Height=338;

  EditCOUNTRY2->SetFocus();
}
//---------------------------------------------------------------------------


//Удаление грейда
void __fastcall TSprav::N7Click(TObject *Sender)
{
  if (Application->MessageBox("Вы действительно хотите удалить выбранную запись?","Удаление записи",
                          MB_YESNO+MB_ICONINFORMATION)==ID_NO)
    {
      Abort();
    }

  //Удаление записи
  DM->qObnovlenie->Close();
  DM->qObnovlenie->SQL->Clear();
  DM->qObnovlenie->SQL->Add("delete from sp_grade where rowid=chartorowid("+QuotedStr(DM->qSP_grade->FieldByName("rw")->AsString)+")");
  try
    {
      DM->qObnovlenie->ExecSQL();
    }
  catch(Exception &E)
    {
      Application->MessageBox(("Невозможно удалить запись из справочника грейдов (SP_GRADE) "+E.Message).c_str(),"Ошибка",
                              MB_OK+MB_ICONERROR);
      Main->InsertLog("Не выполнено удаление записи из справочника грейдов: грейд "+DM->qSP_grade->FieldByName("grade")->AsString);
      Abort();
    }


  //Логи
  Main->InsertLog("Выполнено удаление записи из справочника грейдов: грейд "+DM->qSP_grade->FieldByName("grade")->AsString);

  DM->qSP_grade->Requery();

  Application->MessageBox("Запись успешно удалена","Удаление записи",
                          MB_OK+MB_ICONINFORMATION);
}
//---------------------------------------------------------------------------


//Удаление объекта
void __fastcall TSprav::N15Click(TObject *Sender)
{
  if (Application->MessageBox("Вы действительно хотите удалить выбранную запись?","Удаление записи",
                          MB_YESNO+MB_ICONINFORMATION)==ID_NO)
    {
      Abort();
    }

  //Удаление записи
  DM->qObnovlenie->Close();
  DM->qObnovlenie->SQL->Clear();
  DM->qObnovlenie->SQL->Add("delete from sp_obekt where rowid=chartorowid("+QuotedStr(DM->qSP_obekt->FieldByName("rw")->AsString)+")");
  try
    {
      DM->qObnovlenie->ExecSQL();
    }
  catch(Exception &E)
    {
      Application->MessageBox(("Невозможно удалить запись из справочника объектов (SP_OBEKT) "+E.Message).c_str(),"Ошибка",
                              MB_OK+MB_ICONERROR);
      Main->InsertLog("Не выполнено удаление записи из справочника объектов: объектов "+DM->qSP_obekt->FieldByName("obekt")->AsString);
      Abort();
    }


  //Логи
  Main->InsertLog("Выполнено удаление записи из справочника объектов: объектов "+DM->qSP_obekt->FieldByName("obekt")->AsString);

  DM->qSP_obekt->Requery();

  Application->MessageBox("Запись успешно удалена","Удаление записи",
                          MB_OK+MB_ICONINFORMATION);
}
//---------------------------------------------------------------------------

//Удаление страны
void __fastcall TSprav::N17Click(TObject *Sender)
{
  if (Application->MessageBox("Вы действительно хотите удалить выбранную запись?","Удаление записи",
                          MB_YESNO+MB_ICONINFORMATION)==ID_NO)
    {
      Abort();
    }

  //Удаление записи
  DM->qObnovlenie->Close();
  DM->qObnovlenie->SQL->Clear();
  DM->qObnovlenie->SQL->Add("delete from sp_country where rowid=chartorowid("+QuotedStr(DM->qSP_country->FieldByName("rw")->AsString)+")");
  try
    {
      DM->qObnovlenie->ExecSQL();
    }
  catch(Exception &E)
    {
      Application->MessageBox(("Невозможно удалить запись из справочника стран (SP_COUNTRY) "+E.Message).c_str(),"Ошибка",
                              MB_OK+MB_ICONERROR);
      Main->InsertLog("Не выполнено удаление записи из справочника стран: страна "+DM->qSP_country->FieldByName("country")->AsString);
      Abort();
    }


  //Логи
  Main->InsertLog("Выполнено удаление записи из справочника стран: страна "+DM->qSP_country->FieldByName("country")->AsString);

  DM->qSP_country->Requery();

  Application->MessageBox("Запись успешно удалена","Удаление записи",
                          MB_OK+MB_ICONINFORMATION);
}
//---------------------------------------------------------------------------

//Удаление города
void __fastcall TSprav::N21Click(TObject *Sender)
{
  if (Application->MessageBox("Вы действительно хотите удалить выбранную запись?","Удаление записи",
                          MB_YESNO+MB_ICONINFORMATION)==ID_NO)
    {
      Abort();
    }

  //Удаление записи
  DM->qObnovlenie->Close();
  DM->qObnovlenie->SQL->Clear();
  DM->qObnovlenie->SQL->Add("delete from sp_city where rowid=chartorowid("+QuotedStr(DM->qSP_city->FieldByName("rw")->AsString)+")");
  try
    {
      DM->qObnovlenie->ExecSQL();
    }
  catch(Exception &E)
    {
      Application->MessageBox(("Невозможно удалить запись из справочника городов (SP_CITY) "+E.Message).c_str(),"Ошибка",
                              MB_OK+MB_ICONERROR);
      Main->InsertLog("Не выполнено удаление записи из справочника городов: город "+DM->qSP_city->FieldByName("city")->AsString);
      Abort();
    }


  //Логи
  Main->InsertLog("Выполнено удаление записи из справочника городов: город "+DM->qSP_city->FieldByName("city")->AsString);

  DM->qSP_city->Requery();

  Application->MessageBox("Запись успешно удалена","Удаление записи",
                          MB_OK+MB_ICONINFORMATION);
}
//---------------------------------------------------------------------------

//Изменение/добавление объекта
void __fastcall TSprav::BitBtn7Click(TObject *Sender)
{
  TLocateOptions SearchOptions;
  AnsiString Sql;
  int rec;

  //Проверка на введенный город
  if (EditGOROD1->Text.IsEmpty())
    {
      Application->MessageBox("Не указан город!!!","Предупреждение",
                               MB_OK+MB_ICONINFORMATION);
      EditGOROD1->SetFocus();
    }

  //Проверка на введеный объект
  if (EditOBEKT->Text.IsEmpty())
    {
      Application->MessageBox("Не указан объект!!!","Предупреждение",
                               MB_OK+MB_ICONINFORMATION);
      EditOBEKT->SetFocus();
    }
  //Проверка на уже существующую запись
  if (DM->qSP_obekt->Locate("obekt",EditOBEKT->Text,SearchOptions << loCaseInsensitive))
    {
      Application->MessageBox("Введенный объект уже есть в справочнике","Предупреждение",
                              MB_OK+MB_ICONINFORMATION);
      EditOBEKT->SetFocus();
    }

  //проверка на введенный адрес
  if (EditADRESS->Text.IsEmpty())
    {
      Application->MessageBox("Не указан адрес объекта!!!","Предупреждение",
                               MB_OK+MB_ICONINFORMATION);
      EditADRESS->SetFocus();
    }

  //Добавление записи
  if (fl_sp_red==0)
    {
       Sql="insert into sp_obekt (kod, obekt, adress, kod_city)\
           values ( \
                   (select max(kod)+1 from sp_obekt),"                                        \
                   +QuotedStr(EditOBEKT->Text)+","                                                       \
                   +QuotedStr(EditADRESS->Text)+",                                                           \
                   (select kod from sp_city where upper(city)=upper(trim("+QuotedStr(EditGOROD1->Text)+")))  \
                   )";
    }
  //Обновление записи
  else if (fl_sp_red==1)
    {
      Sql="update sp_obekt set\
                                obekt="+QuotedStr(EditOBEKT->Text)+",         \
                                adress="+QuotedStr(EditADRESS->Text)+",                                                   \
                                kod_city=(select kod from sp_city where upper(city)=upper(trim("+QuotedStr(EditGOROD1->Text)+")))  \
           where rowid=chartorowid("+QuotedStr(DM->qSP_obekt->FieldByName("rw")->AsString)+")";

      rec = DM->qSP_obekt->RecNo;
    }

  DM->qObnovlenie->Close();
  DM->qObnovlenie->SQL->Clear();
  DM->qObnovlenie->SQL->Add(Sql);
  try
    {
      DM->qObnovlenie->ExecSQL();
    }
  catch(Exception &E)
    {
      Application->MessageBox(("Невозможно добавить/обновить запись в справочнике объектов (SP_OBEKT)"+E.Message).c_str(),"Ошибка",
                              MB_OK+MB_ICONERROR);
      Panel10->Visible=false;
      DBGridEh4->Top=163;
      DBGridEh4->Height=435;
      Abort();
    }

  //Логи

  DM->qSP_obekt->Requery();

  //Возвращение курсора на строку
  if (fl_sp_red==0)
    {
      // При добавлении записи возвращать на нее курсор
      DM->qSP_obekt->Locate("obekt",EditOBEKT->Text,SearchOptions << loCaseInsensitive);
    }
  else
    {
      DM->qSP_obekt->RecNo = rec;
    }


  Panel10->Visible=false;
  DBGridEh4->Top=163;
  DBGridEh4->Height=435;
}
//---------------------------------------------------------------------------

//Добавление/изменение страны
void __fastcall TSprav::BitBtn9Click(TObject *Sender)
{
  TLocateOptions SearchOptions;
  AnsiString Sql;
  int rec;

  //Проверка на введенный код
  if (EditKOD->Text.IsEmpty())
    {
      Application->MessageBox("Не указан код страны!!!","Предупреждение",
                               MB_OK+MB_ICONINFORMATION);
      EditKOD->SetFocus();
    }

  //Проверка на введенную страну
  if (EditCOUNTRY->Text.IsEmpty())
    {
      Application->MessageBox("Не указана страна!!!","Предупреждение",
                               MB_OK+MB_ICONINFORMATION);
      EditCOUNTRY->SetFocus();
    }

  //Проверка на уже существующую страну
  if (DM->qSP_country->Locate("country",EditCOUNTRY->Text,SearchOptions << loCaseInsensitive))
    {
      Application->MessageBox("Введенная страна уже есть в справочнике","Предупреждение",
                              MB_OK+MB_ICONINFORMATION);
      EditCOUNTRY->SetFocus();
    }

  //проверка на введенное сокращенное название страны
  if (EditCOUNTRY_K->Text.IsEmpty())
    {
      Application->MessageBox("Не указано краткое название страны!!!","Предупреждение",
                               MB_OK+MB_ICONINFORMATION);
      EditCOUNTRY_K->SetFocus();
    }

  //Добавление записи
  if (fl_sp_red==0)
    {
      //Проверка на уже существующий код
      if (DM->qSP_country->Locate("kod",EditKOD->Text,SearchOptions << loCaseInsensitive))
        {
          Application->MessageBox("Введенный код страны уже есть в справочнике","Предупреждение",
                                   MB_OK+MB_ICONINFORMATION);
          EditKOD->SetFocus();
        }


       Sql="insert into sp_country (kod, country, country_k)\
           values ("\
                   +QuotedStr(EditKOD->Text)+","                                        \
                   +QuotedStr(EditCOUNTRY->Text)+","                                    \
                   +QuotedStr(EditCOUNTRY_K->Text)+")";
    }
  //Обновление записи
  else if (fl_sp_red==1)
    {
      Sql="update sp_country set\                                            \
                               kod="+QuotedStr(EditKOD->Text)+",                        \
                               country="+QuotedStr(EditCOUNTRY->Text)+",                \
                               country_k="+QuotedStr(EditCOUNTRY_K->Text)+"             \
           where rowid=chartorowid("+QuotedStr(DM->qSP_country->FieldByName("rw")->AsString)+")";

      rec = DM->qSP_country->RecNo;
    }

  DM->qObnovlenie->Close();
  DM->qObnovlenie->SQL->Clear();
  DM->qObnovlenie->SQL->Add(Sql);
  try
    {
      DM->qObnovlenie->ExecSQL();
    }
  catch(Exception &E)
    {
      Application->MessageBox(("Невозможно добавить/обновить запись в справочнике стран (SP_COUNTRY)"+E.Message).c_str(),"Ошибка",
                              MB_OK+MB_ICONERROR);
      Panel16->Visible=false;
      DBGridEh5->Top=166;
      DBGridEh5->Height=437;
      Abort();
    }

  //Логи

  DM->qSP_country->Requery();

  //Возвращение курсора на строку
  if (fl_sp_red==0)
    {
      // При добавлении записи возвращать на нее курсор
      DM->qSP_country->Locate("country",EditCOUNTRY->Text,SearchOptions << loCaseInsensitive);
    }
  else
    {
      DM->qSP_country->RecNo = rec;
    }


  Panel16->Visible=false;
  DBGridEh5->Top=166;
  DBGridEh5->Height=437;
}
//---------------------------------------------------------------------------


void __fastcall TSprav::BitBtn11Click(TObject *Sender)
{
  TLocateOptions SearchOptions;
  AnsiString Sql;
  int rec;


  //Проверка на введенную страну
  if (EditCOUNTRY2->Text.IsEmpty())
    {
      Application->MessageBox("Не указана страна!!!","Предупреждение",
                               MB_OK+MB_ICONINFORMATION);
      EditCOUNTRY2->SetFocus();
    }

  //Проверка на введенный город
  if (EditGOROD2->Text.IsEmpty())
    {
      Application->MessageBox("Не указан город!!!","Предупреждение",
                               MB_OK+MB_ICONINFORMATION);
      EditGOROD2->SetFocus();
    }

  //Проверка на уже существующую запись
  if (DM->qSP_city->Locate("city",EditGOROD2->Text,SearchOptions << loCaseInsensitive))
    {
      Application->MessageBox("Введенный город уже есть в справочнике","Предупреждение",
                              MB_OK+MB_ICONINFORMATION);
      EditGOROD2->SetFocus();
    }


  //Добавление записи
  if (fl_sp_red==0)
    {
       Sql="insert into sp_city (kod, kod_country, city)\
           values ( \
                   (select max(kod)+1 from sp_city),                                        \
                   (select kod from sp_country where upper(country)=upper(trim("+QuotedStr(EditCOUNTRY2->Text)+"))),"  \
                   +QuotedStr(EditGOROD2->Text)+")";
    }
  //Обновление записи
  else if (fl_sp_red==1)
    {
      Sql="update sp_city set\
                               kod_country=(select kod from sp_country where upper(country)=upper(trim("+QuotedStr(EditCOUNTRY2->Text)+"))),  \
                               city="+QuotedStr(EditGOROD2->Text)+"                                                   \
           where rowid=chartorowid("+QuotedStr(DM->qSP_city->FieldByName("rw")->AsString)+")";

      rec = DM->qSP_city->RecNo;
    }

  DM->qObnovlenie->Close();
  DM->qObnovlenie->SQL->Clear();
  DM->qObnovlenie->SQL->Add(Sql);
  try
    {
      DM->qObnovlenie->ExecSQL();
    }
  catch(Exception &E)
    {
      Application->MessageBox(("Невозможно добавить/обновить запись в справочнике городов (SP_CITY)"+E.Message).c_str(),"Ошибка",
                              MB_OK+MB_ICONERROR);
      Panel13->Visible=false;
      DBGridEh6->Top=160;
      DBGridEh6->Height=437;
      Abort();
    }

  //Логи

  DM->qSP_city->Requery();

  //Возвращение курсора на строку
  if (fl_sp_red==0)
    {
      // При добавлении записи возвращать на нее курсор
      DM->qSP_city->Locate("city",EditGOROD2->Text,SearchOptions << loCaseInsensitive);
    }
  else
    {
      DM->qSP_city->RecNo = rec;
    }


  Panel13->Visible=false;
  DBGridEh6->Top=160;
  DBGridEh6->Height=437;
}
//---------------------------------------------------------------------------

void __fastcall TSprav::PageControl1DrawTab(TCustomTabControl *Control,
      int TabIndex, const TRect &Rect, bool Active)
{
  AnsiString S;
  int x, y;

  S = PageControl1->Pages[TabIndex]->Caption;
  Control->Canvas->FillRect(Rect);

  if (Active)
    {
      Control->Canvas->Brush->Color = (TColor)0x00DEF5F4;
      Control->Canvas->Font->Color = clBlack;
      Control->Canvas->FillRect(Rect);
    }
  else
    {
      Control->Canvas->Font->Color = clBlack;
    }

  x = CenterPoint(Rect).x - div(Control->Canvas->TextWidth(S),2).quot;
  y = CenterPoint(Rect).y - div(Control->Canvas->TextHeight(S),2).quot;
  Control->Canvas->TextOut(x,y,S);               
}
//---------------------------------------------------------------------------


