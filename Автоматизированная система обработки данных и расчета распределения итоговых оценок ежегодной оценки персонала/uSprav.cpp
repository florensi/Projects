//---------------------------------------------------------------------------

#include <vcl.h>
#pragma hdrstop

#include "uSprav.h"
#include "uDM.h"
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
void __fastcall TSprav::DBGridEh1DrawColumnCell(TObject *Sender,
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
