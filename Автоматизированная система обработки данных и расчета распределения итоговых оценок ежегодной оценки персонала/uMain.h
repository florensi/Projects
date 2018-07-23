//---------------------------------------------------------------------------

#ifndef uMainH
#define uMainH
//---------------------------------------------------------------------------
#include <Classes.hpp>
#include <Controls.hpp>
#include <StdCtrls.hpp>
#include <Forms.hpp>
#include <ComCtrls.hpp>
#include <Menus.hpp>
#include "DBGridEh.hpp"
#include <ExtCtrls.hpp>
#include <Grids.hpp>
#include <Buttons.hpp>

#include <ComObj.hpp>
#include <jpeg.hpp>
#include <FileCtrl.hpp>
#include "Excel_2K_SRVR.h"

//---------------------------------------------------------------------------
class TMain : public TForm
{
__published:	// IDE-managed Components
        TMainMenu *MainMenu1;
        TPopupMenu *PopupMenu1;
        TStatusBar *StatusBar1;
        TMenuItem *N1;
        TMenuItem *N2;
        TMenuItem *N3;
        TMenuItem *N4;
        TMenuItem *N5;
        TPanel *Panel1;
        TDBGridEh *DBGridEh1;
        TPanel *Panel2;
        TSpeedButton *SpeedButton1;
        TSpeedButton *SpeedButton2;
        TSpeedButton *SpeedButtonRedaktirovanie;
        TMenuItem *N6;
        TImage *Image1;
        TImage *Image3;
        TMenuItem *N7;
        TBevel *Bevel1;
        TSpeedButton *SpeedButton4;
        TMenuItem *N8;
        TMenuItem *N9;
        TMenuItem *N10;
        TMenuItem *N11;
        TMenuItem *N12;
        TMenuItem *N13;
        TMenuItem *N14;
        TMenuItem *Excel1;
        TMenuItem *N15;
        TMenuItem *N16;
        TMenuItem *N17;
        TMenuItem *N18;
        TMenuItem *N19;
        TMenuItem *N20;
        TMenuItem *Afqk1;
        TMenuItem *N21;
        TMenuItem *N22;
        TMenuItem *Cghfdrf1;
        TMenuItem *N110;
        TMenuItem *N25;
        TSpeedButton *SpeedButton3;
        TMenuItem *N26;
        TFileListBox *FileListBox1;
        TMenuItem *N29;
        TMenuItem *N30;
        TMenuItem *N31;
        TMenuItem *Gjgthtdtltyysvhfjnybrfv1;
        TMenuItem *N32;
        TMenuItem *N33;
        TMenuItem *N34;
        TMenuItem *N35;
        TMenuItem *N36;
        TMenuItem *N37;
        TMenuItem *N38;
        TMenuItem *N39;
        TMenuItem *N41;
        TMenuItem *N42;
        TMenuItem *N43;
        TMenuItem *N40;
        TMenuItem *N23;
        TMenuItem *N24;
        TMenuItem *N27;
        TMenuItem *N28;
        TMenuItem *NUvol;
        TMenuItem *N44;
        TMenuItem *N45;
        TMenuItem *N46;

        void __fastcall N5Click(TObject *Sender);
        void __fastcall FormCreate(TObject *Sender);
        void __fastcall N3Click(TObject *Sender);
        void __fastcall SpeedButton1Click(TObject *Sender);
        void __fastcall SpeedButtonRedaktirovanieClick(TObject *Sender);
        void __fastcall SpeedButton2Click(TObject *Sender);
        void __fastcall N7Click(TObject *Sender);
        void __fastcall SpeedButton4Click(TObject *Sender);
        void __fastcall N8Click(TObject *Sender);
        void __fastcall N9Click(TObject *Sender);
        void __fastcall DBGridEh1DrawColumnCell(TObject *Sender,
          const TRect &Rect, int DataCol, TColumnEh *Column,
          TGridDrawState State);
        void __fastcall N6Click(TObject *Sender);
        void __fastcall DBGridEh1KeyDown(TObject *Sender, WORD &Key,
          TShiftState Shift);
        void __fastcall DBGridEh1KeyPress(TObject *Sender, char &Key);
        void __fastcall FormKeyDown(TObject *Sender, WORD &Key,
          TShiftState Shift);
        void __fastcall DBGridEh1DblClick(TObject *Sender);
        void __fastcall N11Click(TObject *Sender);
        void __fastcall N14Click(TObject *Sender);
        void __fastcall Excel1Click(TObject *Sender);
        void __fastcall N15Click(TObject *Sender);
        void __fastcall N16Click(TObject *Sender);
        void __fastcall N18Click(TObject *Sender);
        void __fastcall FormShow(TObject *Sender);
        void __fastcall N19Click(TObject *Sender);
        void __fastcall N20Click(TObject *Sender);
        void __fastcall N21Click(TObject *Sender);
        void __fastcall FormResize(TObject *Sender);
        void __fastcall N22Click(TObject *Sender);
        void __fastcall Cghfdrf1Click(TObject *Sender);
        void __fastcall DBGridEh1Columns13GetCellParams(TObject *Sender,
          bool EditMode, TColCellParamsEh *Params);
        void __fastcall N110Click(TObject *Sender);
        void __fastcall SpeedButton3Click(TObject *Sender);
        void __fastcall N26Click(TObject *Sender);
        void __fastcall StatusBar1DblClick(TObject *Sender);
        void __fastcall N30Click(TObject *Sender);
        void __fastcall Gjgthtdtltyysvhfjnybrfv1Click(TObject *Sender);
        void __fastcall N43Click(TObject *Sender);
        void __fastcall N41Click(TObject *Sender);
        void __fastcall N42Click(TObject *Sender);
        void __fastcall N34Click(TObject *Sender);
        void __fastcall N35Click(TObject *Sender);
        void __fastcall N36Click(TObject *Sender);
        void __fastcall N23Click(TObject *Sender);
        void __fastcall N24Click(TObject *Sender);
        void __fastcall N27Click(TObject *Sender);
        void __fastcall NUvolClick(TObject *Sender);
        void __fastcall N44Click(TObject *Sender);
        void __fastcall N45Click(TObject *Sender);
private:	// User declarations

public:		// User declarations
        TProgressBar *ProgressBar;

        int god,   //Отчетный год (год за который выполняется оценка персонала) или просмотр оценки за предыдущий период
            god_t; //Первоначальный текущий год, определяемый программой
        AnsiString DocPath, TempPath, WorkPath, Path, WordPath,
                   UserName, DomainName, UserFullName;

        



        bool __fastcall GetMyDocumentsDir(AnsiString &FolderPath);
        bool __fastcall GetTempDir(AnsiString &FolderPath);

        void __fastcall InsertLog(AnsiString Msg);
        AnsiString __fastcall FindWordPath();
        AnsiString  __fastcall SetNull(AnsiString str, AnsiString r="NULL");
        float __fastcall SetNullF (AnsiString str);
        void __fastcall SpisokExcel(AnsiString otchet_zex);
        void __fastcall OtchetKR(AnsiString otchet);
        void __fastcall OtchetKRSokr(AnsiString otchet);
        void __fastcall TMain::OtchetPP(AnsiString otchet);

        void __fastcall Spisok_po_zex2016(); //Формирование списка по цехам в Excel до 2017 года
        void __fastcall Spisok_po_zex2017(); //Формирование списка по цехам в Excel с 2017 года
        void __fastcall SpisokExcel2017(AnsiString otchet_zex); //Формирование списка по предприятию в Excel с 2017 года

        __fastcall TMain(TComponent* Owner);
};
//---------------------------------------------------------------------------
extern PACKAGE TMain *Main;
//---------------------------------------------------------------------------
#endif
