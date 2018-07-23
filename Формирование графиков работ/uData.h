//---------------------------------------------------------------------------

#ifndef uDataH
#define uDataH
//---------------------------------------------------------------------------
#include <Classes.hpp>
#include <Controls.hpp>
#include <StdCtrls.hpp>
#include <Forms.hpp>
#include <Buttons.hpp>
#include <ComCtrls.hpp>
#include <ExtCtrls.hpp>
//---------------------------------------------------------------------------
class TData : public TForm
{
__published:	// IDE-managed Components
        TPanel *Panel1;
        TBevel *Bevel1;
        TBitBtn *btnVibor;
        TBitBtn *BitBtn2;
        TDateTimePicker *DateTimePicker1;
        TLabel *Label1;
        void __fastcall btnViborKeyDown(TObject *Sender, WORD &Key,
          TShiftState Shift);
        void __fastcall FormShow(TObject *Sender);
        void __fastcall btnViborClick(TObject *Sender);
private:	// User declarations
public:		// User declarations
        __fastcall TData(TComponent* Owner);
};
//---------------------------------------------------------------------------
extern PACKAGE TData *Data;
//---------------------------------------------------------------------------
#endif
