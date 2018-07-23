//---------------------------------------------------------------------------

#ifndef uSpravH
#define uSpravH
//---------------------------------------------------------------------------
#include <Classes.hpp>
#include <Controls.hpp>
#include <StdCtrls.hpp>
#include <Forms.hpp>
#include <ExtCtrls.hpp>
#include <Graphics.hpp>
#include "DBGridEh.hpp"
#include <Grids.hpp>
#include <Buttons.hpp>
#include <Menus.hpp>
#include <IdGlobal.hpp>

//---------------------------------------------------------------------------
class TSprav : public TForm
{
__published:	// IDE-managed Components
        TPanel *Panel1;
        TPanel *Panel2;
        TBevel *Bevel1;
        TBevel *Bevel2;
        TLabel *Label1;
        TImage *Image1;
        TDBGridEh *DBGridEh1;
        TLabel *Label2;
        TLabel *Label3;
        TLabel *Label4;
        TLabel *Label5;
        TEdit *EditDEN;
        TEdit *EditMES;
        TEdit *EditGOD;
        TBitBtn *BitBtn1;
        TBitBtn *BitBtn2;
        TBevel *Bevel3;
        TPopupMenu *PopupMenu1;
        TMenuItem *N1Dobav;
        TMenuItem *N2Redakt;
        TMenuItem *N3Delet;
        void __fastcall N1DobavClick(TObject *Sender);
        void __fastcall N2RedaktClick(TObject *Sender);
        void __fastcall BitBtn2Click(TObject *Sender);
        void __fastcall BitBtn1Click(TObject *Sender);
        void __fastcall EditDENKeyPress(TObject *Sender, char &Key);
        void __fastcall N3DeletClick(TObject *Sender);
        void __fastcall FormKeyDown(TObject *Sender, WORD &Key,
          TShiftState Shift);
        void __fastcall FormShow(TObject *Sender);
        void __fastcall DBGridEh1DrawColumnCell(TObject *Sender,
          const TRect &Rect, int DataCol, TColumnEh *Column,
          TGridDrawState State);
        void __fastcall DBGridEh1KeyDown(TObject *Sender, WORD &Key,
          TShiftState Shift);
        void __fastcall DBGridEh1DblClick(TObject *Sender);
private:	// User declarations
public:		// User declarations

       int fl_sp; //Редактирование справочника празднечных дней fl_sp=1, Добавление записи fl_sp=0
        __fastcall TSprav(TComponent* Owner);
};
//---------------------------------------------------------------------------
extern PACKAGE TSprav *Sprav;
//---------------------------------------------------------------------------
#endif
