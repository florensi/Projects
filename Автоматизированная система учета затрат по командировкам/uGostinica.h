//---------------------------------------------------------------------------

#ifndef uGostinicaH
#define uGostinicaH
//---------------------------------------------------------------------------
#include <Classes.hpp>
#include <Controls.hpp>
#include <StdCtrls.hpp>
#include <Forms.hpp>
#include <Buttons.hpp>
#include <ExtCtrls.hpp>
#include <jpeg.hpp>
#include <IdGlobal.hpp>
//---------------------------------------------------------------------------
class TGostinica : public TForm
{
__published:	// IDE-managed Components
        TImage *Image1;
        TPanel *Panel1;
        TBevel *Bevel1;
        TBitBtn *BitBtn1;
        TBitBtn *Cansel;
        TLabel *Label1;
        TLabel *Label3;
        TEdit *EditCOMFORT;
        TLabel *Label4;
        TEdit *EditCLEAR;
        TLabel *Label5;
        TEdit *EditPERSONAL;
        TLabel *Label6;
        TEdit *EditPITANIE;
        TLabel *Label7;
        TEdit *EditSERVIS;
        TLabel *Label8;
        TEdit *EditUSLUGI;
        TLabel *Label9;
        TEdit *EditRASPOLOG;
        TLabel *Label10;
        TEdit *EditVPECHAT;
        TLabel *Label11;
        TEdit *EditORGANIZ;
        TLabel *Label2;
        TBevel *Bevel2;
        TBevel *Bevel3;
        TBevel *Bevel4;
        TBevel *Bevel5;
        TBevel *Bevel6;
        TBevel *Bevel7;
        TBevel *Bevel8;
        TBevel *Bevel9;
        void __fastcall CanselClick(TObject *Sender);
        void __fastcall FormKeyDown(TObject *Sender, WORD &Key,
          TShiftState Shift);
        void __fastcall BitBtn1Click(TObject *Sender);
        void __fastcall EditCOMFORTExit(TObject *Sender);
        void __fastcall EditCOMFORTKeyPress(TObject *Sender, char &Key);
        void __fastcall EditCLEARExit(TObject *Sender);
        void __fastcall EditPERSONALExit(TObject *Sender);
        void __fastcall EditPITANIEExit(TObject *Sender);
        void __fastcall EditSERVISExit(TObject *Sender);
        void __fastcall EditUSLUGIExit(TObject *Sender);
        void __fastcall EditRASPOLOGExit(TObject *Sender);
        void __fastcall EditVPECHATExit(TObject *Sender);
        void __fastcall EditORGANIZExit(TObject *Sender);
        void __fastcall FormShow(TObject *Sender);
private:	// User declarations
public:		// User declarations
        __fastcall TGostinica(TComponent* Owner);
};
//---------------------------------------------------------------------------
extern PACKAGE TGostinica *Gostinica;
//---------------------------------------------------------------------------
#endif
