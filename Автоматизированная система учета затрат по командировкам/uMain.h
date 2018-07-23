//---------------------------------------------------------------------------

#ifndef uMainH
#define uMainH
//---------------------------------------------------------------------------
#include <Classes.hpp>
#include <Controls.hpp>
#include <StdCtrls.hpp>
#include <Forms.hpp>
#include "DBGridEh.hpp"
#include <ExtCtrls.hpp>
#include <Grids.hpp>
#include <jpeg.hpp>
#include <Menus.hpp>
#include <ComCtrls.hpp>
#include "Excel_2K_SRVR.h"
#include <Buttons.hpp>
//---------------------------------------------------------------------------
class TMain : public TForm
{
__published:	// IDE-managed Components
        TImage *Image1;
        TDBGridEh *DBGridEh1;
        TLabel *Label1;
        TLabel *Label3;
        TEdit *EditZEX;
        TEdit *EditS;
        TEdit *EditPO;
        TLabel *Label2;
        TLabel *Label4;
        TPopupMenu *PopupMenu1;
        TMainMenu *MainMenu1;
        TMenuItem *N1;
        TMenuItem *N2Redaktir;
        TMenuItem *N3;
        TMenuItem *N4;
        TMenuItem *N5OBRAT_SV;
        TMenuItem *N6;
        TEdit *EditTN;
        TLabel *Label5;
        TLabel *Label6;
        TLabel *Label7;
        TMenuItem *N5;
        TMenuItem *N7;
        TMenuItem *N8;
        TMenuItem *N9;
        TMenuItem *N10;
        TMenuItem *N11;
        TMenuItem *N12;
        TMenuItem *N2;
        TMenuItem *N13;
        TMenuItem *N14;
        TMenuItem *N15;
        TMenuItem *N16;
        TMenuItem *N17;
        TStatusBar *StatusBar1;
        TBitBtn *BitBtn1;
        TBitBtn *BitBtn2;
        TMenuItem *N18;
        TLabel *Label8;
        TEdit *EditFAM;
        TMenuItem *N19;
        TMenuItem *N20;
        void __fastcall DBGridEh1DblClick(TObject *Sender);
        void __fastcall N2RedaktirClick(TObject *Sender);
        void __fastcall N1Click(TObject *Sender);
        void __fastcall N5OBRAT_SVClick(TObject *Sender);
        void __fastcall FormKeyDown(TObject *Sender, WORD &Key,
          TShiftState Shift);
        void __fastcall DBGridEh1DrawColumnCell(TObject *Sender,
          const TRect &Rect, int DataCol, TColumnEh *Column,
          TGridDrawState State);
        void __fastcall FormShow(TObject *Sender);
        void __fastcall N9Click(TObject *Sender);
        void __fastcall N7Click(TObject *Sender);
        void __fastcall N5Click(TObject *Sender);
        void __fastcall N12Click(TObject *Sender);
        void __fastcall N10Click(TObject *Sender);
        void __fastcall N11Click(TObject *Sender);
        void __fastcall N3Click(TObject *Sender);
        void __fastcall N13Click(TObject *Sender);
        void __fastcall FormCreate(TObject *Sender);
        void __fastcall N17Click(TObject *Sender);
        void __fastcall N16Click(TObject *Sender);
        void __fastcall N14Click(TObject *Sender);
        void __fastcall BitBtn1Click(TObject *Sender);
        void __fastcall N18Click(TObject *Sender);
        void __fastcall BitBtn2Click(TObject *Sender);
        void __fastcall EditSExit(TObject *Sender);
        void __fastcall EditPOExit(TObject *Sender);
        void __fastcall N19Click(TObject *Sender);
        void __fastcall N20Click(TObject *Sender);
private:	// User declarations
public:		// User declarations
        TProgressBar *ProgressBar;
        int mm, yyyy; //ќтчетный мес€ц и год
        int fl_redakt; //fl_redakt=0 - добавление записи, fl_redakt=1 -  редактирование записи
        AnsiString DocPath, TempPath, WorkPath, Path, WordPath,
                   UserName, DomainName, UserFullName;

        bool __fastcall GetMyDocumentsDir(AnsiString &FolderPath);
        bool __fastcall GetTempDir(AnsiString &FolderPath);
        //void __fastcall InsertLog(AnsiString Msg);
        AnsiString __fastcall FindWordPath();
        void __fastcall SetInfoEdit();
        void __fastcall InsertLog(AnsiString Msg);

        __fastcall TMain(TComponent* Owner);
};
//---------------------------------------------------------------------------
extern PACKAGE TMain *Main;
//---------------------------------------------------------------------------
#endif
