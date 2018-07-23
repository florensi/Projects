//---------------------------------------------------------------------------

#ifndef uVvodH
#define uVvodH
//---------------------------------------------------------------------------
#include <System.Classes.hpp>
#include <Vcl.Controls.hpp>
#include <Vcl.StdCtrls.hpp>
#include <Vcl.Forms.hpp>
#include <Vcl.ExtCtrls.hpp>
#include <Vcl.Buttons.hpp>
#include <Vcl.Imaging.jpeg.hpp>
#include <IdGlobal.hpp>
//---------------------------------------------------------------------------
class TVvod : public TForm
{
__published:	// IDE-managed Components
	TPanel *Panel1;
	TGroupBox *GroupBoxREZULT_PP;
	TGroupBox *GroupBoxREZULT_KPE;
	TGroupBox *GroupBoxSTAND;
	TLabel *LabelFIO;
	TLabel *Label2;
	TLabel *LabelNZEX;
	TEdit *EditZEX;
	TLabel *Label4;
	TEdit *EditTN;
	TLabel *Label5;
	TEdit *EditID_DOLG;
	TEdit *EditDOLG;
	TBevel *Bevel1;
	TBevel *BevelImage;
	TImage *Image1;
	TLabel *Label7;
	TLabel *Label8;
	TEdit *EditPROIZ;
	TEdit *EditPROIZ_BALL;
	TLabel *Label9;
	TLabel *Label10;
	TEdit *EditKPE;
	TEdit *EditKPE_BALL;
	TLabel *Label11;
	TLabel *Label12;
	TEdit *EditOTKL;
	TEdit *EditOTKL_BALL;
	TGroupBox *GroupBoxRPERSONAL_PREEM;
	TGroupBox *GroupBoxRPERSONAL_INFO;
	TGroupBox *GroupBoxVPERSONAL_5C;
	TLabel *Label13;
	TLabel *Label14;
	TEdit *EditPRIEM;
	TEdit *EditPRIEM_BALL;
	TLabel *Label15;
	TLabel *Label16;
	TEdit *EditINFO;
	TEdit *EditINFO_BALL;
	TLabel *Label17;
	TLabel *Label18;
	TEdit *EditC5;
	TEdit *EditC5_BALL;
	TGroupBox *GroupBoxVPERSONAL_KNS;
	TGroupBox *GroupBoxVPERSONAL_SPP;
	TGroupBox *GroupBoxOT;
	TLabel *Label19;
	TLabel *Label20;
	TEdit *EditKNS;
	TEdit *EditKNS_BALL;
	TLabel *Label21;
	TLabel *Label22;
	TEdit *EditSPP_KOL;
	TEdit *EditSPP_BALL;
	TLabel *Label23;
	TLabel *Label24;
	TLabel *Label25;
	TEdit *EditOT_UPR;
	TEdit *EditOT_TREB;
	TEdit *EditTRUD_D;
	TBitBtn *Save;
	TBitBtn *Cansel;
	TBevel *BevelREZULT;
	TBevel *BevelRPERSONAL;
	TBevel *BevelSTAND;
	TBevel *BevelVPERSONAL;
	TBevel *BevelOT;
	TLabel *LabelREZULT;
	TLabel *LabelRPERSONAL;
	TLabel *LabelSTAND;
	TLabel *LabelVPERSONAL;
	TCheckBox *CheckBoxPODCH;
	TGroupBox *GroupBoxTD;
	TGroupBox *GroupBoxOCENKA;
	TLabel *Label1;
	TEdit *EditOCENKA;
	TLabel *Label3;
	TComboBox *ComboBoxREIT;
	void __fastcall CanselClick(TObject *Sender);
	void __fastcall FormKeyDown(TObject *Sender, WORD &Key, TShiftState Shift);
	void __fastcall FormShow(TObject *Sender);
	void __fastcall EditZEXKeyPress(TObject *Sender, System::WideChar &Key);
	void __fastcall EditTNKeyPress(TObject *Sender, System::WideChar &Key);
	void __fastcall EditPROIZ_BALLExit(TObject *Sender);
	void __fastcall EditKPE_BALLExit(TObject *Sender);
	void __fastcall EditPRIEM_BALLExit(TObject *Sender);
	void __fastcall EditINFO_BALLExit(TObject *Sender);
	void __fastcall EditOTKL_BALLExit(TObject *Sender);
	void __fastcall EditC5_BALLExit(TObject *Sender);
	void __fastcall EditKNS_BALLExit(TObject *Sender);
	void __fastcall EditSPP_BALLExit(TObject *Sender);
	void __fastcall EditOT_UPRExit(TObject *Sender);
	void __fastcall EditOT_TREBExit(TObject *Sender);
	void __fastcall EditTRUD_DExit(TObject *Sender);
	void __fastcall SaveClick(TObject *Sender);
	void __fastcall EditTNChange(TObject *Sender);
	void __fastcall EditPROIZExit(TObject *Sender);
	void __fastcall EditKPEExit(TObject *Sender);
	void __fastcall EditOTKLExit(TObject *Sender);
	void __fastcall EditPRIEMExit(TObject *Sender);
	void __fastcall EditINFOExit(TObject *Sender);
	void __fastcall EditC5Exit(TObject *Sender);
	void __fastcall EditKNSExit(TObject *Sender);
	void __fastcall EditSPP_KOLExit(TObject *Sender);
	void __fastcall FormCreate(TObject *Sender);
	void __fastcall EditPROIZ_BALLChange(TObject *Sender);
	void __fastcall EditKPE_BALLChange(TObject *Sender);
	void __fastcall EditPRIEM_BALLChange(TObject *Sender);
	void __fastcall EditINFO_BALLChange(TObject *Sender);
	void __fastcall EditOTKL_BALLChange(TObject *Sender);
	void __fastcall EditC5_BALLChange(TObject *Sender);
	void __fastcall EditKNS_BALLChange(TObject *Sender);
	void __fastcall EditSPP_BALLChange(TObject *Sender);
	void __fastcall EditOT_UPRChange(TObject *Sender);
	void __fastcall EditOT_TREBChange(TObject *Sender);
	void __fastcall EditTRUD_DChange(TObject *Sender);
private:	// User declarations
public:		// User declarations
	AnsiString zfio, zzex, ztn, zid_dolg, zdolg, zpodch, zproiz, zproiz_ball,
			   zkpe, zkpe_ball, zpriem, zpriem_ball, zinfo, zinfo_ball, zotkl,
			   zotkl_ball, zc5, zc5_ball, zkns, zkns_ball, zspp_kol, zspp_ball,
			   zot_upr, zot_treb, ztrud_d, zocenka, zreit;
	void __fastcall SetDataEdit();  //Заполнение Edit-ов
	void __fastcall SetNullEdit(); 	//Очищение Edit-ов
	AnsiString  __fastcall SetNull(AnsiString str, AnsiString r="NULL");
	float  __fastcall SetN(String str, float r=0);

    void __fastcall Ocenka();   //Расчет оценки при редактировании записи
	__fastcall TVvod(TComponent* Owner);
};
//---------------------------------------------------------------------------
extern PACKAGE TVvod *Vvod;
//---------------------------------------------------------------------------
#endif
