//---------------------------------------------------------------------------

#ifndef uSpravH
#define uSpravH
//---------------------------------------------------------------------------
#include <Classes.hpp>
#include <Controls.hpp>
#include <StdCtrls.hpp>
#include <Forms.hpp>
#include <ExtCtrls.hpp>
#include "DBGridEh.hpp"
#include <Grids.hpp>
#include <jpeg.hpp>
//---------------------------------------------------------------------------
class TSprav : public TForm
{
__published:	// IDE-managed Components
        TPanel *Panel1;
        TPanel *Panel2;
        TDBGridEh *DBGridEh1;
        TBevel *Bevel1;
        TBevel *Bevel2;
        TImage *Image1;
        TLabel *Label1;
        void __fastcall DBGridEh1DrawColumnCell(TObject *Sender,
          const TRect &Rect, int DataCol, TColumnEh *Column,
          TGridDrawState State);
private:	// User declarations
public:		// User declarations
        __fastcall TSprav(TComponent* Owner);
};
//---------------------------------------------------------------------------
extern PACKAGE TSprav *Sprav;
//---------------------------------------------------------------------------
#endif
