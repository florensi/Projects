//---------------------------------------------------------------------------

#ifndef uVvodH
#define uVvodH
//---------------------------------------------------------------------------
#include <Classes.hpp>
#include <Controls.hpp>
#include <StdCtrls.hpp>
#include <Forms.hpp>
#include <ExtCtrls.hpp>
#include <jpeg.hpp>
#include <Buttons.hpp>
#include <DateUtils.hpp>
#include <IdGlobal.hpp>
#include "PropStorageEh.hpp"

//---------------------------------------------------------------------------
class TVvod : public TForm
{
__published:	// IDE-managed Components
        TImage *Image1;
        TLabel *Label1;
        TLabel *Label2;
        TLabel *Label4;
        TLabel *Label5;
        TLabel *LabelZEX;
        TLabel *Label7;
        TLabel *Label9;
        TLabel *Label10;
        TLabel *Label11;
        TLabel *Label12;
        TPanel *Panel1;
        TLabel *Label14;
        TLabel *Label15;
        TLabel *Label16;
        TBitBtn *BitBtn1;
        TBitBtn *Cansel;
        TGroupBox *GroupBox1;
        TGroupBox *GroupBox2;
        TPanel *Panel2;
        TLabel *Label13;
        TLabel *Label17;
        TLabel *Label18;
        TLabel *Label19;
        TLabel *Label20;
        TLabel *Label23;
        TLabel *Label24;
        TLabel *Label25;
        TLabel *Label27;
        TLabel *Label28;
        TLabel *Label8;
        TLabel *Label29;
        TGroupBox *GroupBox3;
        TEdit *EditFIO;
        TEdit *EditTN;
        TEdit *EditZEX;
        TEdit *EditPROF;
        TEdit *EditGRADE;
        TEdit *EditG_KIEV;
        TEdit *EditG_UKR;
        TEdit *EditDATA_N;
        TEdit *EditDATA_K;
        TComboBox *ComboBoxCHEL;
        TEdit *EditDATA_GOST_N;
        TEdit *EditDATA_GOST_K;
        TEdit *EditSUM_PROGIV;
        TEdit *EditSUM_SUT;
        TEdit *EditSUM_AVIA;
        TMemo *MemoPRIMECH;
        TEdit *EditSROK;
        TEdit *EditDATA_ZAK;
        TBevel *Bevel1;
        TEdit *EditKOD_KOM;
        TBevel *Bevel2;
        TBevel *Bevel3;
        TComboBox *ComboBoxGOSTINICA;
        TLabel *Label21;
        TComboBox *ComboBoxSTRANA;
        TComboBox *ComboBoxGOROD;
        TComboBox *ComboBoxOBEKT;
        TEdit *EditSUM_GD;
        TEdit *EditSUM_PROCH;
        TLabel *Label22;
        TLabel *Label26;
        TLabel *Label30;
        TLabel *Label31;
        TEdit *EditSUM_TRANSP;
        TButton *ButtonGOSTINICA;
        TLabel *Label32;
        TCheckBox *CheckBoxAVIA;
        TCheckBox *CheckBoxGD;
        TCheckBox *CheckBoxBUS;
        TCheckBox *CheckBoxAVTO;
        TCheckBox *CheckBoxPROEZD;
        TLabel *Label3;
        TLabel *Label6;
        TEdit *EditNAPRAVL;
        TBevel *Bevel4;
        TEdit *EditSTOIM;
        TLabel *Label33;
        TLabel *Label34;
        TEdit *EditN_DOCUM;
        TLabel *Label35;
        TEdit *EditG_ZAGRAN;
        TEdit *EditADRESS;
        TEdit *EditGOST_ADR;
        void __fastcall ButtonGOSTINICAClick(TObject *Sender);
        void __fastcall CanselClick(TObject *Sender);
        void __fastcall FormKeyDown(TObject *Sender, WORD &Key,
          TShiftState Shift);
        void __fastcall FormShow(TObject *Sender);
        void __fastcall BitBtn1Click(TObject *Sender);
        void __fastcall FormCreate(TObject *Sender);
        void __fastcall EditTNChange(TObject *Sender);
        void __fastcall EditDATA_NExit(TObject *Sender);
        void __fastcall EditDATA_KExit(TObject *Sender);
        void __fastcall EditDATA_GOST_NExit(TObject *Sender);
        void __fastcall EditDATA_GOST_KExit(TObject *Sender);
        void __fastcall EditDATA_NKeyPress(TObject *Sender, char &Key);
        void __fastcall EditZEXKeyPress(TObject *Sender, char &Key);
        void __fastcall EditG_KIEVKeyPress(TObject *Sender, char &Key);
        void __fastcall EditFIOKeyPress(TObject *Sender, char &Key);
        void __fastcall ComboBoxSTRANAExit(TObject *Sender);
        void __fastcall ComboBoxGORODExit(TObject *Sender);
        void __fastcall ComboBoxCHELExit(TObject *Sender);
        void __fastcall ComboBoxOBEKTChange(TObject *Sender);
        void __fastcall ComboBoxGOSTINICAChange(TObject *Sender);
        void __fastcall ComboBoxOBEKTExit(TObject *Sender);
        void __fastcall ComboBoxGOSTINICAExit(TObject *Sender);
        void __fastcall EditZEXChange(TObject *Sender);
        void __fastcall EditZEXExit(TObject *Sender);
        void __fastcall EditGRADEExit(TObject *Sender);
        void __fastcall EditGRADEChange(TObject *Sender);
private:	// User declarations
public:		// User declarations
       AnsiString  __fastcall SetNull (AnsiString str, AnsiString r="NULL");
       
        __fastcall TVvod(TComponent* Owner);
};
//---------------------------------------------------------------------------
extern PACKAGE TVvod *Vvod;
//---------------------------------------------------------------------------
#endif
