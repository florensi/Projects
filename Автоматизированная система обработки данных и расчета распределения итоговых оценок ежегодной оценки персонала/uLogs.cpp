//---------------------------------------------------------------------------
#pragma link "EhLibADO"

#include <vcl.h>
#pragma hdrstop

#include "uLogs.h"
#include "uDM.h"
//---------------------------------------------------------------------------
#pragma package(smart_init)
#pragma link "DBGridEh"
#pragma resource "*.dfm"
TLogs *Logs;
//---------------------------------------------------------------------------
__fastcall TLogs::TLogs(TComponent* Owner)
        : TForm(Owner)
{
}
//---------------------------------------------------------------------------
void __fastcall TLogs::RadioGroup1Click(TObject *Sender)
{
  //все записи
  if (RadioGroup1->ItemIndex == 0)
    {
      DM->qLogs->Filtered = false;
    }
  //редактирование записи
  else if (RadioGroup1->ItemIndex == 1)
    {
      DM->qLogs->Filtered = false;
      DM->qLogs->Filter = " text like 'Обновление записи%'";
      DM->qLogs->Filtered = true;
    }
  //редактирование рейтинга
  else if (RadioGroup1->ItemIndex == 2)
    {
      DM->qLogs->Filtered = false;
      DM->qLogs->Filter = " text like 'Обновление рейтинга%'";
      DM->qLogs->Filtered = true;
    }
  //расчет рейтинга
  else if (RadioGroup1->ItemIndex == 3)
    {
      DM->qLogs->Filtered = false;
      DM->qLogs->Filter = " text like 'Расчет рейтинга%'";
      DM->qLogs->Filtered = true;
    }
  //загрузка данных из Excel
  else if (RadioGroup1->ItemIndex == 4)
    {
      DM->qLogs->Filtered = false;
      DM->qLogs->Filter = " text like 'Загрузка%'";
      DM->qLogs->Filtered = true;
    }
}
//---------------------------------------------------------------------------

void __fastcall TLogs::FormShow(TObject *Sender)
{
  Logs->RadioGroup1->SetFocus();
  RadioGroup1->ItemIndex = 0;
  DM->qLogs->Requery();
}
//---------------------------------------------------------------------------

void __fastcall TLogs::DBGridEh1DrawColumnCell(TObject *Sender,
      const TRect &Rect, int DataCol, TColumnEh *Column,
      TGridDrawState State)
{
  // выделение цветом активной записи
 if (State.Contains(gdSelected))
    {
      ((TDBGridEh *) Sender)->Canvas->Brush->Color = TColor(0x00C8F7E3);//0x00A3F1D1);//clInfoBk;
      ((TDBGridEh *) Sender)->Canvas->Font->Color= clBlack;
    }
  ((TDBGridEh *) Sender)->DefaultDrawColumnCell(Rect, DataCol, Column, State);        
}
//---------------------------------------------------------------------------

void __fastcall TLogs::FormKeyPress(TObject *Sender, char &Key)
{
  /*if (Key == VK_ESCAPE)
    {
      Logs->Close();
    } */
}
//---------------------------------------------------------------------------

