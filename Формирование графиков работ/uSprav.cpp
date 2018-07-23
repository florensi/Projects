//---------------------------------------------------------------------------

#include <vcl.h>
#pragma hdrstop

#include "uSprav.h"
#include "uDM.h"
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

void __fastcall TSprav::N1DobavClick(TObject *Sender)
{
  fl_sp=0;
  Label2->Caption = "Добавление записи";
  DBGridEh1->Enabled = false;

  EditDEN->Text = "";
  EditMES->Text = "";
  EditGOD->Text = "";
  EditDEN->Font->Color = clBlack;

  Panel2->Visible = true;
  EditDEN->SetFocus();
}                              
//---------------------------------------------------------------------------

void __fastcall TSprav::N2RedaktClick(TObject *Sender)
{
  fl_sp =1;

  EditDEN->Text = DM->qSprav->FieldByName("den")->AsString;
  EditMES->Text = DM->qSprav->FieldByName("mes")->AsString;
  EditGOD->Text = DM->qSprav->FieldByName("god")->AsString;
  EditDEN->Font->Color = clBlack;
  DBGridEh1->Enabled = false;
  
  Label2->Caption = "Редактирование записи";
  Panel2->Visible = true;
  EditDEN->SetFocus();
}
//---------------------------------------------------------------------------

void __fastcall TSprav::BitBtn2Click(TObject *Sender)
{
  Panel2->Visible = false;
  DBGridEh1->Enabled = true;
  DBGridEh1->SetFocus();        
}
//---------------------------------------------------------------------------

void __fastcall TSprav::BitBtn1Click(TObject *Sender)
{
  TDateTime d;
  AnsiString Str, Sql, prazd;
  int rec;
  TLocateOptions SearchOpt;

  /*prazd - добавляемая праздничная дата*/

  //Проверка на заполнение дня
  if (EditDEN->Text.IsEmpty())
    {
      Application->MessageBox("Введите день!!!","Предупреждение",
                              MB_OK + MB_ICONINFORMATION);
      EditDEN->SetFocus();
      Abort();
    }

  //Проверка на заполнение месяца
  if (EditMES->Text.IsEmpty())
    {
      Application->MessageBox("Введите месяц!!!","Предупреждение",
                              MB_OK + MB_ICONINFORMATION);
      EditMES->SetFocus();
      Abort();
    }

  //Проверка на заполнение года
  if (EditGOD->Text.IsEmpty())
    {
      Application->MessageBox("Введите год!!!","Предупреждение",
                              MB_OK + MB_ICONINFORMATION);
      EditGOD->SetFocus();
      Abort();
    }
  else
    {
      if (StrToInt(EditGOD->Text)<Main->god)
        {
          Application->MessageBox("Год не может быть меньше текущего!!!","Предупреждение",
                                  MB_OK+MB_ICONINFORMATION);
          EditGOD->SetFocus();
          Abort();
        }
    }

  //Проверка на существование даты
  if(!TryStrToDate((EditDEN->Text+"."+EditMES->Text+"."+EditGOD->Text),d))
    {
      Application->MessageBox("Неверный формат даты","Ошибка", MB_OK);
      EditDEN->Font->Color = clRed;
      EditDEN->SetFocus();
    }
  else
    {
      EditDEN->Font->Color = clBlack;
    }

  //добовляемая праздничная дата
  prazd = FormatDateTime("dd.mm.yyyy", StrToDate(EditDEN->Text+"."+EditMES->Text+"."+EditGOD->Text));

  if (fl_sp==0)
    {
      //Проверка на существование такой записи в таблице
      Variant locvalues[] = {EditDEN->Text, EditMES->Text, EditGOD->Text};
      if (DM->qSprav->Locate("den;mes;god", VarArrayOf(locvalues,2), SearchOpt <<loCaseInsensitive))
        {
          if (Application->MessageBox("Данная запись в таблице уже существует!!!\nВы действительно хотите добавить еще одну?","Предупреждение",
                                      MB_YESNO + MB_ICONINFORMATION)==ID_NO)
            {
              Abort();
            }
        }

      //Добавление записи
      Sql = "insert into sp_prd (den,mes,god) \
             values (\
             "+ EditDEN->Text+", \
             "+ EditMES->Text+", \
             "+ EditGOD->Text+")";

      //Логи
      Str = "Добавление праздничного дня '"+prazd+"' ";
    }
  else
    {
      //Обновление данных
      Sql = "update sp_prd set \
                           den = "+EditDEN->Text+", \
                           mes = "+EditMES->Text+", \
                           god = "+EditGOD->Text+"  \
             where rowid = chartorowid("+QuotedStr(DM->qSprav->FieldByName("rw")->AsString)+")";

      //Логи
      Str = "Обновление праздничного дня с '"+(DM->qSprav->FieldByName("den")->AsInteger < 10 ? "0"+ DM->qSprav->FieldByName("den")->AsString : DM->qSprav->FieldByName("den")->AsString)+"."+
                                              (DM->qSprav->FieldByName("mes")->AsInteger < 10 ? "0"+DM->qSprav->FieldByName("mes")->AsString : DM->qSprav->FieldByName("mes")->AsString)+"."+
                                               DM->qSprav->FieldByName("god")->AsString+"' на дату '"+prazd+"' ";
    }

  rec = DM->qSprav->RecNo;
  DM->qObnovlenie->Close();
  DM->qObnovlenie->SQL->Clear();
  DM->qObnovlenie->SQL->Add(Sql);
  try
    {
      DM->qObnovlenie->ExecSQL();
    }
  catch(...)
    {
      Application->MessageBox("Возникла ошибка при попытке добавления(обновления) записи в таблице (SP_PRD)",
                              "Обновление данных",
                              MB_OK + MB_ICONERROR);
      Main->InsertLog(Str+"не выполнено");
      Abort();
    }

  //Логи
  Main->InsertLog(Str+"выполнено успешно");

  //Возвращение курсора на редактируемую запись
  if (fl_sp==0)
    {
      Variant locvalues [] = {EditDEN->Text, EditMES->Text, EditGOD->Text};
      DM->qSprav->Locate("den;mes;god", VarArrayOf(locvalues,2), SearchOpt << loCaseInsensitive);
    }
  else
    {
      DM->qSprav->RecNo = rec;
    }

  DM->qSprav->Requery();

  DBGridEh1->Enabled = true;
  DBGridEh1->SetFocus();
  Panel2->Visible = false;

  Sprav->N2Redakt->Enabled = true;
  Sprav->N3Delet->Enabled = true;
}
//---------------------------------------------------------------------------

void __fastcall TSprav::EditDENKeyPress(TObject *Sender, char &Key)
{
  if (!(IsNumeric(Key)|| Key=='\b')) Key=0;
}
//---------------------------------------------------------------------------

void __fastcall TSprav::N3DeletClick(TObject *Sender)
{
  AnsiString Sql, prazd;
  int rec;

  /*prazd - удаляемій праздничный день*/

  if (Application->MessageBox("Вы действительно хотите удалить выбранную запись?","Удаление записи",
                              MB_YESNO + MB_ICONINFORMATION)==ID_NO)
    {
      Abort();
    }

  //Удаление записи
  Sql = "delete from sp_prd \
         where rowid = chartorowid("+QuotedStr(DM->qSprav->FieldByName("rw")->AsString)+")";

  rec = DM->qSprav->RecNo;
  prazd = FormatDateTime("dd.mm.yyyy", StrToDate(DM->qSprav->FieldByName("den")->AsString+"."+
                                                 DM->qSprav->FieldByName("mes")->AsString+"."+
                                                 DM->qSprav->FieldByName("god")->AsString));
  DM->qObnovlenie->Close();
  DM->qObnovlenie->SQL->Clear();
  DM->qObnovlenie->SQL->Add(Sql);
  try
    {
      DM->qObnovlenie->ExecSQL();
    }
  catch(...)
    {
      Application->MessageBox("Возникла ошибка при удалении записи","Ошибка",
                              MB_OK + MB_ICONERROR);
      Main->InsertLog("Удаление праздничной даты "+prazd+" не выполнено =(");
      Abort();
    }

  //Логи
  Main->InsertLog("Удаление праздничной даты "+prazd+" выполнено успешно =)");

  DM->qSprav->Requery();

  //Возвращение курсора на строку
  if (DM->qSprav->RecordCount>0)
    {
      if (rec==1)
        {
           DM->qSprav->RecNo = rec;
        }
      else
        {
          DM->qSprav->RecNo = rec-1;
        }
    }

  //Видимость пунктов контекстного меню
  if (DM->qSprav->RecordCount>0)
    {
      Sprav->N2Redakt->Enabled = true;
      Sprav->N3Delet->Enabled = true;
    }
  else
    {
      Sprav->N2Redakt->Enabled = false;
      Sprav->N3Delet->Enabled = false;
    }
}
//---------------------------------------------------------------------------

void __fastcall TSprav::FormKeyDown(TObject *Sender, WORD &Key,
      TShiftState Shift)
{
  if (Key==VK_RETURN)
  FindNextControl((TWinControl *)Sender, true, true, false)->SetFocus();           
}
//---------------------------------------------------------------------------

void __fastcall TSprav::FormShow(TObject *Sender)
{
  DBGridEh1->SetFocus();

  //Видимость пунктов контекстного меню
  if (DM->qSprav->RecordCount>0)
    {
      Sprav->N2Redakt->Enabled = true;
      Sprav->N3Delet->Enabled = true;
    }
  else
    {
      Sprav->N2Redakt->Enabled = false;
      Sprav->N3Delet->Enabled = false;
    }
}
//---------------------------------------------------------------------------

void __fastcall TSprav::DBGridEh1DrawColumnCell(TObject *Sender,
      const TRect &Rect, int DataCol, TColumnEh *Column,
      TGridDrawState State)
{
 // выделение серым цветом активной записи
  if (State.Contains(gdSelected) )
    {
      ((TDBGridEh *) Sender)->Canvas->Brush->Color = clSkyBlue;
      ((TDBGridEh *) Sender)->Canvas->Font->Color = clBlack;
    }

  ((TDBGridEh *) Sender)->DefaultDrawColumnCell(Rect, DataCol, Column, State);
}
//---------------------------------------------------------------------------

void __fastcall TSprav::DBGridEh1KeyDown(TObject *Sender, WORD &Key,
      TShiftState Shift)
{
  //Добавление праздничной даты
  if (Key == VK_INSERT)
    {
      N1DobavClick(Sender);
    }

  //Удаление записи
  if (Key == VK_DELETE && DM->qSprav->RecordCount!=0)
    {
      N3DeletClick(Sender);
    }

   //Редактирование
   if (Key == VK_RETURN && DM->qSprav->RecordCount!=0 )
    {
      N2RedaktClick(Sender);
    }
}
//---------------------------------------------------------------------------

void __fastcall TSprav::DBGridEh1DblClick(TObject *Sender)
{
  //Редактирование
  if (DM->qSprav->RecordCount!=0)
    {
      N2RedaktClick(Sender);
    }
}
//---------------------------------------------------------------------------

