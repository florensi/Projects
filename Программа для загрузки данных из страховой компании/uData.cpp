//---------------------------------------------------------------------------

#include <vcl.h>
#pragma hdrstop

#include "uData.h"
#include "uMain.h"
#include "uDM.h"
//---------------------------------------------------------------------------
#pragma package(smart_init)
#pragma resource "*.dfm"
TData *Data;


const AnsiString Mes[]={"������","�������","����","������","���","����","����",
                        "������","��������","�������","������","�������"};
//---------------------------------------------------------------------------
__fastcall TData::TData(TComponent* Owner)
        : TForm(Owner)
{
}
//---------------------------------------------------------------------------

void __fastcall TData::btnViborKeyDown(TObject *Sender, WORD &Key,
      TShiftState Shift)
{
    if (Key == VK_RETURN)
  FindNextControl((TWinControl *)Sender, true, true,
                   false)->SetFocus();         
}
//---------------------------------------------------------------------------


void __fastcall TData::FormShow(TObject *Sender)
{

   //����� ���� �� grafr � DateTimePicker
  dt = TDateTime( "01." + IntToStr(DM->mm) + "." + IntToStr(DM->yyyy));
  Data->DateTimePicker1->Date = dt;
}
//---------------------------------------------------------------------------



