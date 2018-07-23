//---------------------------------------------------------------------------

#include <vcl.h>
#pragma hdrstop

#include "uGostinica.h"
#include "uDM.h"
#include "uMain.h"
//---------------------------------------------------------------------------
#pragma package(smart_init)
#pragma resource "*.dfm"
TGostinica *Gostinica;
//---------------------------------------------------------------------------
__fastcall TGostinica::TGostinica(TComponent* Owner)
        : TForm(Owner)
{
}
//---------------------------------------------------------------------------
void __fastcall TGostinica::CanselClick(TObject *Sender)
{
  Close();
}
//---------------------------------------------------------------------------
void __fastcall TGostinica::FormKeyDown(TObject *Sender, WORD &Key,
      TShiftState Shift)
{
  if (Key==VK_RETURN)
  FindNextControl((TWinControl *)Sender, true, true, false)->SetFocus();
}
//---------------------------------------------------------------------------
void __fastcall TGostinica::BitBtn1Click(TObject *Sender)
{
  AnsiString Sql;
  int rec;

  //Проверки

  //Комфорт
  if (EditCOMFORT->Text.IsEmpty())
    {
      Application->MessageBox("Не указана оценка в графе 'КОМФОРТ'","Предупреждение",
                              MB_ICONINFORMATION+MB_OK);
      EditCOMFORT->SetFocus();
      Abort();
    }

  //Чистота
  if (EditCLEAR->Text.IsEmpty())
    {
      Application->MessageBox("Не указана оценка в графе 'ЧИСТОТА'","Предупреждение",
                              MB_ICONINFORMATION+MB_OK);
      EditCLEAR->SetFocus();
      Abort();
    }

  //Персонал
  if (EditPERSONAL->Text.IsEmpty())
    {
      Application->MessageBox("Не указана оценка в графе 'ПЕРСОНАЛ'","Предупреждение",
                              MB_ICONINFORMATION+MB_OK);
      EditPERSONAL->SetFocus();
      Abort();
    }

  //Питание
  if (EditPITANIE->Text.IsEmpty())
    {
      Application->MessageBox("Не указана оценка в графе 'ПИТАНИЕ'","Предупреждение",
                              MB_ICONINFORMATION+MB_OK);
      EditPITANIE->SetFocus();
      Abort();
    }

  //Сервис
  if (EditSERVIS->Text.IsEmpty())
    {
      Application->MessageBox("Не указана оценка в графе 'СЕРВИС'","Предупреждение",
                              MB_ICONINFORMATION+MB_OK);
      EditSERVIS->SetFocus();
      Abort();
    }

  //Услуги
  if (EditUSLUGI->Text.IsEmpty())
    {
      Application->MessageBox("Не указана оценка в графе 'УСЛУГИ'","Предупреждение",
                              MB_ICONINFORMATION+MB_OK);
      EditUSLUGI->SetFocus();
      Abort();
    }

  //Расположение
  if (EditRASPOLOG->Text.IsEmpty())
    {
      Application->MessageBox("Не указана оценка в графе 'РАСПОЛОЖЕНИЕ'","Предупреждение",
                              MB_ICONINFORMATION+MB_OK);
      EditRASPOLOG->SetFocus();
      Abort();
    }

  //Впечатление
  if (EditVPECHAT->Text.IsEmpty())
    {
      Application->MessageBox("Не указана оценка в графе 'ВПЕЧАТЛЕНИЕ'","Предупреждение",
                              MB_ICONINFORMATION+MB_OK);
      EditVPECHAT->SetFocus();
      Abort();
    }

  //Организация
  if (EditORGANIZ->Text.IsEmpty())
    {
      Application->MessageBox("Не указана оценка в графе 'ОРГАНИЗАЦИЯ КОМАНДИРОВКИ'","Предупреждение",
                              MB_ICONINFORMATION+MB_OK);
      EditORGANIZ->SetFocus();
      Abort();
    }

  //Редактирование записи
  Sql="update komandirovki set     \
                                COMFORT="+EditCOMFORT->Text+",                 \
                                CLEAR="+EditCLEAR->Text+",                     \
                                PERSONAL="+EditPERSONAL->Text+",               \
                                PITANIE="+EditPITANIE->Text+",                 \
                                SERVIS="+EditSERVIS->Text+",                   \
                                USLUGI="+EditUSLUGI->Text+",                   \
                                RASPOLOG="+EditRASPOLOG->Text+",               \
                                VPECHAT="+EditVPECHAT->Text+",                 \
                                ORGANIZ="+EditORGANIZ->Text+"                  \
         where rowid=chartorowid("+QuotedStr(DM->qKomandirovki->FieldByName("rw")->AsString)+")";

  rec = DM->qKomandirovki->RecNo;

  DM->qObnovlenie->Close();
  DM->qObnovlenie->SQL->Clear();
  DM->qObnovlenie->SQL->Add(Sql);
  try
    {
      DM->qObnovlenie->ExecSQL();
    }
  catch(Exception &E)
    {
      Application->MessageBox(("Невозможно добавить/изменить запись о гостинице в таблице по командировкам (KOMANDIROVKI) "+ E.Message).c_str(),"Ошибка",
                              MB_ICONINFORMATION+MB_OK);
      Abort();
    }


  //Обновление рейтинга в справочнике гостиниц
  Sql="update sp_gostinica set reit= (select round((sum(nvl(comfort,0))/count(*)+  \
                                             sum(nvl(clear,0))/count(*)+           \
                                             sum(nvl(personal,0))/count(*)+        \
                                             sum(nvl(pitanie,0))/count(*)+         \
                                             sum(nvl(servis,0))/count(*)+          \
                                             sum(nvl(uslugi,0))/count(*)+          \
                                             sum(nvl(raspolog,0))/count(*)+        \
                                             sum(nvl(vpechat,0))/count(*)          \
                                             )/40*100,2) as reit                   \
                                      from komandirovki                            \
                                      where gostinica="+DM->qKomandirovki->FieldByName("kod_gostinica")->AsString+"    \
                                      group by gostinica)                          \
      where kod="+DM->qKomandirovki->FieldByName("kod_gostinica")->AsString;

  DM->qObnovlenie->Close();
  DM->qObnovlenie->SQL->Clear();
  DM->qObnovlenie->SQL->Add(Sql);
  try
    {
      DM->qObnovlenie->ExecSQL();
    }
  catch(Exception &E)
    {
      Application->MessageBox(("Невозможно добавить/изменить запись о рейтинге гостиницы в справочник (SP_GOSTINICA) "+ E.Message).c_str(),"Ошибка",
                              MB_ICONINFORMATION+MB_OK);
      Main->InsertLog("Не выполнено обновление рейтинга в справочнике гостиниц: город "+DM->qKomandirovki->FieldByName("gorod")->AsString+", гостиница "+DM->qKomandirovki->FieldByName("gostinica")->AsString);
      Abort();
    }

  DM->qKomandirovki->Requery();
  DM->qSP_gostinica->Requery();

  //Логи
  Main->InsertLog("Выполнено обновление рейтинга в справочнике гостиниц: город "+DM->qKomandirovki->FieldByName("gorod")->AsString+", гостиница "+DM->qKomandirovki->FieldByName("gostinica")->AsString);

  //Возвращение курсора на строку
  DM->qKomandirovki->RecNo = rec;

  Gostinica->Close();
}
//---------------------------------------------------------------------------

void __fastcall TGostinica::EditCOMFORTExit(TObject *Sender)
{
  if (ActiveControl == Cansel)
    {
      Gostinica->Close();
    }
  else
    {

      if (!EditCOMFORT->Text.IsEmpty() && StrToInt(EditCOMFORT->Text)>5)
        {
          Application->MessageBox("Оценка может быть в пределах от 1 до 5","Предупреждение",
                                   MB_OK+MB_ICONINFORMATION);
          EditCOMFORT->SetFocus();
          Abort();
        }
    }
}
//---------------------------------------------------------------------------

void __fastcall TGostinica::EditCOMFORTKeyPress(TObject *Sender, char &Key)
{
  if (! (IsNumeric(Key) || Key=='\b') ) Key=0;
}
//---------------------------------------------------------------------------

void __fastcall TGostinica::EditCLEARExit(TObject *Sender)
{
  if (ActiveControl == Cansel)
    {
      Gostinica->Close();
    }
  else
    {
      if (!EditCOMFORT->Text.IsEmpty() && StrToInt(EditCLEAR->Text)>5)
        {
          Application->MessageBox("Оценка может быть в пределах от 1 до 5","Предупреждение",
                                   MB_OK+MB_ICONINFORMATION);
          EditCLEAR->SetFocus();
          Abort();
        }
    }
}
//---------------------------------------------------------------------------

void __fastcall TGostinica::EditPERSONALExit(TObject *Sender)
{
  if (ActiveControl == Cansel)
    {
      Gostinica->Close();
    }
  else
    {
      if (!EditCOMFORT->Text.IsEmpty() && StrToInt(EditPERSONAL->Text)>5)
        {
          Application->MessageBox("Оценка может быть в пределах от 1 до 5","Предупреждение",
                                   MB_OK+MB_ICONINFORMATION);
          EditPERSONAL->SetFocus();
          Abort();
        }
    }
}
//---------------------------------------------------------------------------

void __fastcall TGostinica::EditPITANIEExit(TObject *Sender)
{
  if (ActiveControl == Cansel)
    {
      Gostinica->Close();
    }
  else
    {
      if (!EditCOMFORT->Text.IsEmpty() && StrToInt(EditPITANIE->Text)>5)
        {
          Application->MessageBox("Оценка может быть в пределах от 1 до 5","Предупреждение",
                                   MB_OK+MB_ICONINFORMATION);
          EditPITANIE->SetFocus();
          Abort();
        }
    }
}
//---------------------------------------------------------------------------

void __fastcall TGostinica::EditSERVISExit(TObject *Sender)
{
  if (ActiveControl == Cansel)
    {
      Gostinica->Close();
    }
  else
    {
      if (!EditCOMFORT->Text.IsEmpty() && StrToInt(EditSERVIS->Text)>5)
        {
          Application->MessageBox("Оценка может быть в пределах от 1 до 5","Предупреждение",
                                   MB_OK+MB_ICONINFORMATION);
          EditSERVIS->SetFocus();
          Abort();
        }
    }
}
//---------------------------------------------------------------------------

void __fastcall TGostinica::EditUSLUGIExit(TObject *Sender)
{
  if (ActiveControl == Cansel)
    {
      Gostinica->Close();
    }
  else
    {
      if (!EditCOMFORT->Text.IsEmpty() && StrToInt(EditUSLUGI->Text)>5)
        {
          Application->MessageBox("Оценка может быть в пределах от 1 до 5","Предупреждение",
                                   MB_OK+MB_ICONINFORMATION);
          EditUSLUGI->SetFocus();
          Abort();
        }
    }
}
//---------------------------------------------------------------------------

void __fastcall TGostinica::EditRASPOLOGExit(TObject *Sender)
{
  if (ActiveControl == Cansel)
    {
      Gostinica->Close();
    }
  else
    {
      if (!EditCOMFORT->Text.IsEmpty() && StrToInt(EditRASPOLOG->Text)>5)
        {
          Application->MessageBox("Оценка может быть в пределах от 1 до 5","Предупреждение",
                                   MB_OK+MB_ICONINFORMATION);
          EditRASPOLOG->SetFocus();
          Abort();
        }
    }
}
//---------------------------------------------------------------------------

void __fastcall TGostinica::EditVPECHATExit(TObject *Sender)
{
  if (ActiveControl == Cansel)
    {
      Gostinica->Close();
    }
  else
    {
      if (!EditCOMFORT->Text.IsEmpty() && StrToInt(EditVPECHAT->Text)>5)
        {
          Application->MessageBox("Оценка может быть в пределах от 1 до 5","Предупреждение",
                                   MB_OK+MB_ICONINFORMATION);
          EditVPECHAT->SetFocus();
          Abort();
        }
    }
}
//---------------------------------------------------------------------------

void __fastcall TGostinica::EditORGANIZExit(TObject *Sender)
{
  if (ActiveControl == Cansel)
    {
      Gostinica->Close();
    }
  else
    {
      if (!EditCOMFORT->Text.IsEmpty() && StrToInt(EditORGANIZ->Text)>5)
        {
          Application->MessageBox("Оценка может быть в пределах от 1 до 5","Предупреждение",
                                   MB_OK+MB_ICONINFORMATION);
          EditORGANIZ->SetFocus();
          Abort();
        }
    }    
}
//---------------------------------------------------------------------------

void __fastcall TGostinica::FormShow(TObject *Sender)
{
  EditCOMFORT->SetFocus();        
}
//---------------------------------------------------------------------------

