//---------------------------------------------------------------------------

#ifndef uWaitFormH
#define uWaitFormH
//---------------------------------------------------------------------------
#include <Classes.hpp>
#include <Controls.hpp>
#include <StdCtrls.hpp>
#include <Forms.hpp>
#include "CGAUGES.h"
#include <ExtCtrls.hpp>
#include <jpeg.hpp>
//---------------------------------------------------------------------------
class TWaitForm : public TForm
{
__published:	// IDE-managed Components
        TCGauge *CGauge1;
        TImage *Image1;
        TLabel *Label1;
        void __fastcall FormMouseDown(TObject *Sender, TMouseButton Button,
          TShiftState Shift, int X, int Y);
        void __fastcall Image1MouseDown(TObject *Sender,
          TMouseButton Button, TShiftState Shift, int X, int Y);
private:	// User declarations
public:		// User declarations
        __fastcall TWaitForm(TComponent* Owner);
};
//---------------------------------------------------------------------------
extern PACKAGE TWaitForm *WaitForm;
//---------------------------------------------------------------------------
#endif
