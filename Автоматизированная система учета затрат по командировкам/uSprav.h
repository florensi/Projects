//---------------------------------------------------------------------------

#ifndef uSpravH
#define uSpravH
//---------------------------------------------------------------------------
#include <Classes.hpp>
#include <Controls.hpp>
#include <StdCtrls.hpp>
#include <Forms.hpp>
#include <ComCtrls.hpp>
#include <ExtCtrls.hpp>
#include "DBGridEh.hpp"
#include <Grids.hpp>
#include <Menus.hpp>
#include <jpeg.hpp>
#include <Buttons.hpp>
#include <IdGlobal.hpp>
//---------------------------------------------------------------------------
class TSprav : public TForm
{
__published:	// IDE-managed Components
        TPageControl *PageControl1;
        TTabSheet *TabSheet1;
        TTabSheet *TabSheet2;
        TTabSheet *TabSheet3;
        TTabSheet *TabSheet4;
        TTabSheet *TabSheet5;
        TTabSheet *TabSheet6;
        TPanel *Panel1;
        TPanel *Panel3;
        TPanel *Panel4;
        TPanel *Panel6;
        TPanel *Panel7;
        TPanel *Panel9;
        TPanel *Panel10;
        TPanel *Panel12;
        TPanel *Panel13;
        TPanel *Panel15;
        TPanel *Panel16;
        TPanel *Panel18;
        TDBGridEh *DBGridEh1;
        TDBGridEh *DBGridEh2;
        TDBGridEh *DBGridEh3;
        TDBGridEh *DBGridEh4;
        TDBGridEh *DBGridEh5;
        TDBGridEh *DBGridEh6;
        TPopupMenu *PopupMenu1;
        TPopupMenu *PopupMenu2;
        TPopupMenu *PopupMenu3;
        TPopupMenu *PopupMenu4;
        TMenuItem *Ljfdbnmpfgbcm1;
        TMenuItem *N1Redakt;
        TMenuItem *N2Redakt;
        TMenuItem *N3;
        TMenuItem *N4;
        TMenuItem *N5;
        TMenuItem *N6;
        TMenuItem *N7;
        TMenuItem *N8;
        TMenuItem *N9;
        TMenuItem *N10;
        TMenuItem *N11;
        TMenuItem *N12;
        TMenuItem *N13;
        TMenuItem *N14;
        TMenuItem *N15;
        TImage *Image1;
        TLabel *Label1;
        TBitBtn *BitBtn1;
        TBitBtn *BitBtn2;
        TLabel *Label2;
        TBevel *Bevel1;
        TLabel *Label3;
        TLabel *Label4;
        TLabel *Label5;
        TLabel *Label6;
        TLabel *Label7;
        TLabel *Label8;
        TLabel *Label9;
        TLabel *Label10;
        TEdit *EditGRADE;
        TEdit *EditKAT;
        TEdit *EditG_MIN_KIEV;
        TEdit *EditG_KIEV;
        TLabel *Label11;
        TEdit *EditG_MIN_UKR;
        TEdit *EditG_UKR;
        TLabel *Label12;
        TEdit *EditG_ZAGRAN;
        TEdit *EditVAGON;
        TBevel *Bevel2;
        TImage *Image2;
        TLabel *Label13;
        TBitBtn *BitBtn3;
        TBitBtn *BitBtn4;
        TBevel *Bevel3;
        TLabel *Label14;
        TEdit *EditCHEL;
        TImage *Image3;
        TLabel *Label15;
        TBitBtn *BitBtn5;
        TBitBtn *BitBtn6;
        TBevel *Bevel4;
        TLabel *Label16;
        TEdit *EditGOROD;
        TLabel *Label17;
        TEdit *EditGOSTINICA;
        TLabel *Label18;
        TEdit *EditGOST_ADR;
        TBevel *Bevel5;
        TLabel *Label19;
        TEdit *EditGOROD1;
        TLabel *Label20;
        TEdit *EditOBEKT;
        TLabel *Label21;
        TEdit *EditADRESS;
        TBitBtn *BitBtn7;
        TBitBtn *BitBtn8;
        TBevel *Bevel6;
        TLabel *Label22;
        TEdit *EditKOD;
        TLabel *Label23;
        TEdit *EditCOUNTRY;
        TLabel *Label24;
        TEdit *EditCOUNTRY_K;
        TBitBtn *BitBtn9;
        TBitBtn *BitBtn10;
        TBevel *Bevel7;
        TBitBtn *BitBtn11;
        TBitBtn *BitBtn12;
        TLabel *Label25;
        TEdit *EditCOUNTRY2;
        TLabel *Label26;
        TEdit *EditGOROD2;
        TPopupMenu *PopupMenu5;
        TPopupMenu *PopupMenu6;
        TImage *Image4;
        TImage *Image5;
        TImage *Image6;
        TLabel *Label27;
        TLabel *Label28;
        TLabel *Label29;
        TMenuItem *N1;
        TMenuItem *N2;
        TMenuItem *N16;
        TMenuItem *N17;
        TMenuItem *N18;
        TMenuItem *N19;
        TMenuItem *N20;
        TMenuItem *N21;
        void __fastcall Ljfdbnmpfgbcm1Click(TObject *Sender);
        void __fastcall N1RedaktClick(TObject *Sender);
        void __fastcall N4Click(TObject *Sender);
        void __fastcall N5Click(TObject *Sender);
        void __fastcall N8Click(TObject *Sender);
        void __fastcall N9Click(TObject *Sender);
        void __fastcall N12Click(TObject *Sender);
        void __fastcall N13Click(TObject *Sender);
        void __fastcall DBGridEh2DrawColumnCell(TObject *Sender,
          const TRect &Rect, int DataCol, TColumnEh *Column,
          TGridDrawState State);
        void __fastcall FormShow(TObject *Sender);
        void __fastcall BitBtn2Click(TObject *Sender);
        void __fastcall FormKeyDown(TObject *Sender, WORD &Key,
          TShiftState Shift);
        void __fastcall BitBtn1Click(TObject *Sender);
        void __fastcall BitBtn4Click(TObject *Sender);
        void __fastcall DBGridEh1DrawColumnCell(TObject *Sender,
          const TRect &Rect, int DataCol, TColumnEh *Column,
          TGridDrawState State);
        void __fastcall DBGridEh3DrawColumnCell(TObject *Sender,
          const TRect &Rect, int DataCol, TColumnEh *Column,
          TGridDrawState State);
        void __fastcall DBGridEh4DrawColumnCell(TObject *Sender,
          const TRect &Rect, int DataCol, TColumnEh *Column,
          TGridDrawState State);
        void __fastcall DBGridEh5DrawColumnCell(TObject *Sender,
          const TRect &Rect, int DataCol, TColumnEh *Column,
          TGridDrawState State);
        void __fastcall DBGridEh6DrawColumnCell(TObject *Sender,
          const TRect &Rect, int DataCol, TColumnEh *Column,
          TGridDrawState State);
        void __fastcall BitBtn3Click(TObject *Sender);
        void __fastcall N3Click(TObject *Sender);
        void __fastcall DBGridEh1DblClick(TObject *Sender);
        void __fastcall EditGORODKeyPress(TObject *Sender, char &Key);
        void __fastcall BitBtn6Click(TObject *Sender);
        void __fastcall N11Click(TObject *Sender);
        void __fastcall BitBtn5Click(TObject *Sender);
        void __fastcall BitBtn8Click(TObject *Sender);
        void __fastcall BitBtn10Click(TObject *Sender);
        void __fastcall BitBtn12Click(TObject *Sender);
        void __fastcall N1Click(TObject *Sender);
        void __fastcall N18Click(TObject *Sender);
        void __fastcall N2Click(TObject *Sender);
        void __fastcall N19Click(TObject *Sender);
        void __fastcall N7Click(TObject *Sender);
        void __fastcall N15Click(TObject *Sender);
        void __fastcall N17Click(TObject *Sender);
        void __fastcall N21Click(TObject *Sender);
        void __fastcall BitBtn7Click(TObject *Sender);
        void __fastcall BitBtn9Click(TObject *Sender);
        void __fastcall BitBtn11Click(TObject *Sender);
        void __fastcall PageControl1DrawTab(TCustomTabControl *Control,
          int TabIndex, const TRect &Rect, bool Active);
private:	// User declarations
public:		// User declarations
        int fl_sp_red;  //fl_sp_red=0 - добавление записи, fl_sp_red=1 - редактирование
        __fastcall TSprav(TComponent* Owner);
};
//---------------------------------------------------------------------------
extern PACKAGE TSprav *Sprav;
//---------------------------------------------------------------------------
#endif
