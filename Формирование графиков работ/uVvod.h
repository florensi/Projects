//---------------------------------------------------------------------------

#ifndef uVvodH
#define uVvodH
//---------------------------------------------------------------------------
#include <Classes.hpp>
#include <Controls.hpp>
#include <StdCtrls.hpp>
#include <Forms.hpp>
#include <ExtCtrls.hpp>
#include <Graphics.hpp>
#include <Buttons.hpp>
#include <IdGlobal.hpp>
 
//---------------------------------------------------------------------------
class TVvod : public TForm
{
__published:	// IDE-managed Components
        TPanel *Panel1;
        TBevel *Bevel1;
        TBevel *Bevel2;
        TImage *Image1;
        TLabel *Label1;
        TEdit *EditCHF;
        TEdit *EditPCH;
        TEdit *EditVCH;
        TEdit *EditNCH;
        TLabel *Label2;
        TLabel *Label3;
        TLabel *Label4;
        TLabel *Label5;
        TBitBtn *BitBtn1;
        TBitBtn *BitBtn2;
        TLabel *Label6;
        TEdit *EditNSM;
        TLabel *Label7;
        TEdit *EditCHF0;
        TEdit *EditPCH0;
        TEdit *EditNCH0;
        TLabel *Label8;
        void __fastcall BitBtn2Click(TObject *Sender);
        void __fastcall FormKeyDown(TObject *Sender, WORD &Key,
          TShiftState Shift);
        void __fastcall EditCHFKeyPress(TObject *Sender, char &Key);
        void __fastcall BitBtn1Click(TObject *Sender);
        void __fastcall EditPCHKeyPress(TObject *Sender, char &Key);
        void __fastcall FormShow(TObject *Sender);
        void __fastcall EditNSMExit(TObject *Sender);
        void __fastcall EditCHFExit(TObject *Sender);


private:	// User declarations
public:		// User declarations

        AnsiString __fastcall SetNull(AnsiString str, AnsiString r="NULL");
        double __fastcall SetN (AnsiString str, double r=0);
        void __fastcall DostupRedaktEdit();                                     //¬озможность редактировани€ вечерних, ночных и праздничных часов в зависимости от графика

        __fastcall TVvod(TComponent* Owner);
};
//---------------------------------------------------------------------------
extern PACKAGE TVvod *Vvod;
//---------------------------------------------------------------------------
#endif
