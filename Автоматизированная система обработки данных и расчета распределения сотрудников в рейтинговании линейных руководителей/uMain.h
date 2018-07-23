//---------------------------------------------------------------------------

#ifndef uMainH
#define uMainH
//---------------------------------------------------------------------------
#include <System.Classes.hpp>
#include <Vcl.Controls.hpp>
#include <Vcl.StdCtrls.hpp>
#include <Vcl.Forms.hpp>
#include "DBAccess.hpp"
#include "MemDS.hpp"
#include "OracleUniProvider.hpp"
#include "Uni.hpp"
#include "UniProvider.hpp"
#include <Data.DB.hpp>
#include "DBAxisGridsEh.hpp"
#include "DBGridEh.hpp"
#include "DBGridEhGrouping.hpp"
#include "DBGridEhToolCtrls.hpp"
#include "DynVarsEh.hpp"
#include "EhLibVCL.hpp"
#include "GridsEh.hpp"
#include "ToolCtrlsEh.hpp"
#include <Vcl.ComCtrls.hpp>
#include <Vcl.ExtCtrls.hpp>
#include <Vcl.Menus.hpp>
#include <Vcl.ToolWin.hpp>
#include <Vcl.Buttons.hpp>
#include <Vcl.Imaging.jpeg.hpp>
#include <Vcl.FileCtrl.hpp>
#include <Vcl.Dialogs.hpp>
#include <ComObj.hpp>
#include <Vcl.Imaging.pngimage.hpp>
//#include <System.SysUtils.hpp>


//---------------------------------------------------------------------------
class TMain : public TForm
{
__published:	// IDE-managed Components
	TPanel *Panel1;
	TStatusBar *StatusBar1;
	TDBGridEh *DBGridEh1;
	TMainMenu *MainMenu1;
	TPopupMenu *PopupMenu1;
	TMenuItem *N1;
	TMenuItem *N2;
	TMenuItem *N3Otchet;
	TMenuItem *N4;
	TMenuItem *N5;
	TMenuItem *N6Spisok;
	TMenuItem *N7PP;
	TMenuItem *N8KPE;
	TMenuItem *NSprav;
	TMenuItem *N9;
	TMenuItem *N51C5;
	TMenuItem *N10PREEM;
	TMenuItem *N11SPP;
	TMenuItem *N12OT;
	TMenuItem *N13NOT;
	TMenuItem *N14TD;
	TMenuItem *N15;
	TSpeedButton *SpeedButton1;
	TSpeedButton *SpeedButton2;
	TSpeedButton *SpeedButton3;
	TSpeedButton *SpeedButton4;
	TOpenDialog *OpenDialog1;
	TPanel *Panel2;
	TMenuItem *N16Dobav;
	TMenuItem *N17Redact;
	TMenuItem *N6Reiting;
	TPanel *Panel3;
	TImage *Image1;
	TMenuItem *jjjj1;
	TMenuItem *N3;
	void __fastcall N6SpisokClick(TObject *Sender);
	void __fastcall FormCreate(TObject *Sender);
	void __fastcall SpeedButton2Click(TObject *Sender);
	void __fastcall N16DobavClick(TObject *Sender);
	void __fastcall N17RedactClick(TObject *Sender);
	void __fastcall DBGridEh1DblClick(TObject *Sender);
	void __fastcall N7PPClick(TObject *Sender);
	void __fastcall N8KPEClick(TObject *Sender);
	void __fastcall N10PREEMClick(TObject *Sender);
	void __fastcall N11SPPClick(TObject *Sender);
	void __fastcall N12OTClick(TObject *Sender);
	void __fastcall N13NOTClick(TObject *Sender);
	void __fastcall N14TDClick(TObject *Sender);
	void __fastcall N51C5Click(TObject *Sender);
	void __fastcall SpeedButton1Click(TObject *Sender);
	void __fastcall SpeedButton3Click(TObject *Sender);
	void __fastcall SpeedButton4Click(TObject *Sender);
	void __fastcall N15Click(TObject *Sender);
	void __fastcall N6ReitingClick(TObject *Sender);
	void __fastcall DBGridEh1DrawColumnCell(TObject *Sender, const TRect &Rect, int DataCol,
          TColumnEh *Column, TGridDrawState State);
	void __fastcall FormResize(TObject *Sender);
	void __fastcall N5Click(TObject *Sender);
	void __fastcall jjjj1Click(TObject *Sender);
	void __fastcall N9Click(TObject *Sender);

private:	// User declarations
public:		// User declarations

	TProgressBar *ProgressBar;

	int god, kvartal;  //Отчетный период
	int update;  //Признак обновления уже существующего списка работников
	int redakt;  //Признак добавления записи =0, признак редактирования =1
	AnsiString DocPath, TempPath, WorkPath, Path, WordPath;
	String DomainName, UserName, UserFullName, Prava;



	bool __fastcall GetMyDocumentsDir(AnsiString &FolderPath);
	bool __fastcall GetTempDir(AnsiString &FolderPath);
	AnsiString __fastcall FindWordPath();
	void __fastcall InsertLog(String Msg);
	void __fastcall Zagruzka(int tn, int otchet); //Загрузка данных из Excel
	void __fastcall RaschetReit(int pr, String zex, int podch); //Расчет рейтинга
	void __fastcall RaschetOcen(int pr); //Расчет оценки
	void __fastcall OtchetExcelItog(int otchet); //Формирование итогового отчета

	bool  __fastcall Proverka(String tn);

	//void RebuildWindowRgn(TPanel *Panel);
	__fastcall TMain(TComponent* Owner);
};
//---------------------------------------------------------------------------
extern PACKAGE TMain *Main;
//---------------------------------------------------------------------------
#endif
