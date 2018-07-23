//---------------------------------------------------------------------------

#ifndef uLogsH
#define uLogsH
//---------------------------------------------------------------------------
#include <Classes.hpp>
#include <Controls.hpp>
#include <StdCtrls.hpp>
#include <Forms.hpp>
#include "DBGridEh.hpp"
#include <ExtCtrls.hpp>
#include <Grids.hpp>
#include <jpeg.hpp>
#include <ComCtrls.hpp>
//---------------------------------------------------------------------------
class TLogs : public TForm
{
__published:	// IDE-managed Components
        TPanel *Panel1;
        TPanel *Panel2;
        TImage *Image1;
        TDBGridEh *DBGridEh1;
        TRadioGroup *RadioGroup1;
        TStatusBar *StatusBar1;
        void __fastcall RadioGroup1Click(TObject *Sender);
        void __fastcall FormShow(TObject *Sender);
        void __fastcall DBGridEh1DrawColumnCell(TObject *Sender,
          const TRect &Rect, int DataCol, TColumnEh *Column,
          TGridDrawState State);
        void __fastcall FormKeyPress(TObject *Sender, char &Key);
private:	// User declarations
public:		// User declarations
        __fastcall TLogs(TComponent* Owner);
};
//---------------------------------------------------------------------------
extern PACKAGE TLogs *Logs;
//---------------------------------------------------------------------------
#endif
