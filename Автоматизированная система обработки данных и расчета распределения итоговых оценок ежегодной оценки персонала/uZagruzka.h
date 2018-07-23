//---------------------------------------------------------------------------

#ifndef uZagruzkaH
#define uZagruzkaH
//---------------------------------------------------------------------------
#include <Classes.hpp>
#include <Controls.hpp>
#include <StdCtrls.hpp>
#include <Forms.hpp>
#include <Buttons.hpp>
#include <ExtCtrls.hpp>
#include <jpeg.hpp>
#include <FileCtrl.hpp>
#include <ComObj.hpp>
#include <IdGlobal.hpp>
 
#include <SysUtils.hpp>
//---------------------------------------------------------------------------
class TZagruzka : public TForm
{
__published:	// IDE-managed Components
        TPanel *Panel1;
        TPanel *Panel2;
        TBevel *Bevel1;
        TBevel *Bevel2;
        TBevel *Bevel3;
        TBevel *Bevel4;
        TSpeedButton *SpeedButton1;
        TSpeedButton *SpeedButton2;
        TImage *Image1;
        TLabel *Label1;
        TRadioButton *RadioButtonDATAO;
        TCheckBox *CheckBox1;
        TEdit *EditDATA;
        TEdit *EditFIO;
        TEdit *EditOCENKA;
        TEdit *EditREZERV;
        TEdit *EditDOLG;
        TLabel *LabelDATA;
        TLabel *LabelFIO;
        TLabel *LabelOCENKA;
        TLabel *LabelREZERV;
        TLabel *LabelDOLG;
        TLabel *LabelDOLGO;
        TEdit *EditDOLGO;
        TFileListBox *FileListBox1;
        TLabel *LabelZEX;
        TLabel *LabelTN;
        TEdit *EditZEX;
        TEdit *EditTN;
        TRadioButton *RadioButtonEOP;
        TLabel *LabelREZULT_OCEN;
        TLabel *LabelKPE_OCEN;
        TLabel *LabelKOMP_OCEN;
        TEdit *EditREZULT_OCEN;
        TEdit *EditKPE_OCEN;
        TEdit *EditKOMP_OCEN;
        TLabel *LabelFIOEOP;
        TLabel *LabelTNEOP;
        TEdit *EditFIOEOP;
        TEdit *EditTNEOP;
        TCheckBox *CheckBoxREZERV;
        TCheckBox *CheckBoxOCENKA;
        TBevel *Bevel5;
        TLabel *Label2;
        TLabel *Label3;
        TBevel *Bevel6;
        TLabel *Label4;
        TRadioButton *RadioButtonKPE;
        TRadioButton *RadioButtonVZ;
        TRadioButton *RadioButtonKR;
        TEdit *EditTN_KPE;
        TEdit *EditKPE1;
        TEdit *EditKPE2;
        TEdit *EditKPE3;
        TEdit *EditKPE4;
        TEdit *EditTN_VZ;
        TEdit *EditVZ;
        TLabel *LabelTN_KPE;
        TLabel *LabelKPE1;
        TLabel *LabelKPE2;
        TLabel *LabelKPE3;
        TLabel *LabelKPE4;
        TLabel *LabelTN_VZ;
        TLabel *LabelVZ;
        TEdit *EditTN_KR;
        TEdit *EditKRSHIFR_DOLG;
        TEdit *EditKR_FIO;
        TEdit *EditKR_ZEX;
        TLabel *LabelTN_KR;
        TLabel *LabelKR_ZEX;
        TLabel *LabelKRSHIFR_DOLG;
        TLabel *LabelKR_FIO;
        void __fastcall CheckBox1Click(TObject *Sender);
        void __fastcall SpeedButton2Click(TObject *Sender);
        void __fastcall SpeedButton1Click(TObject *Sender);
        void __fastcall FormShow(TObject *Sender);
        void __fastcall EditDATAKeyPress(TObject *Sender, char &Key);
        void __fastcall CheckBoxREZERVClick(TObject *Sender);
        void __fastcall CheckBoxOCENKAClick(TObject *Sender);
        void __fastcall RadioButtonDATAOClick(TObject *Sender);
        void __fastcall RadioButtonEOPClick(TObject *Sender);
private:	// User declarations
public:		// User declarations

        bool  __fastcall Proverka(AnsiString zex);
        __fastcall TZagruzka(TComponent* Owner);
};
//---------------------------------------------------------------------------
extern PACKAGE TZagruzka *Zagruzka;
//---------------------------------------------------------------------------
#endif
