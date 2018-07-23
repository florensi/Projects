//---------------------------------------------------------------------------

#include <vcl.h>
#pragma hdrstop

#include "uWaitForm.h"
//---------------------------------------------------------------------------
#pragma package(smart_init)
#pragma link "CGAUGES"
#pragma resource "*.dfm"
TWaitForm *WaitForm;
//---------------------------------------------------------------------------
__fastcall TWaitForm::TWaitForm(TComponent* Owner)
        : TForm(Owner)
{
}
//---------------------------------------------------------------------------
void __fastcall TWaitForm::FormMouseDown(TObject *Sender,
      TMouseButton Button, TShiftState Shift, int X, int Y)
{
/*  long SC_DRAGMOVE = 0xF012;
  if(Button == mbLeft)
    {
      ReleaseCapture();
      SendMessage(Handle, WM_SYSCOMMAND, SC_DRAGMOVE, 0);
    }    */
}
//---------------------------------------------------------------------------
void __fastcall TWaitForm::Image1MouseDown(TObject *Sender,
      TMouseButton Button, TShiftState Shift, int X, int Y)
{
/*  long SC_DRAGMOVE = 0xF012;
  if(Button == mbLeft)
    {
      ReleaseCapture();
      SendMessage(Handle, WM_SYSCOMMAND, SC_DRAGMOVE, 0);
    }  */
}
//---------------------------------------------------------------------------
