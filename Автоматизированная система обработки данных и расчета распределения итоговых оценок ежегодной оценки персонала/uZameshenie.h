//---------------------------------------------------------------------------

#ifndef uZameshenieH
#define uZameshenieH
//---------------------------------------------------------------------------
#include <Classes.hpp>
#include <Controls.hpp>
#include <StdCtrls.hpp>
#include <Forms.hpp>
#include <ExtCtrls.hpp>
#include <jpeg.hpp>
#include <Buttons.hpp>
#include <IdGlobal.hpp>
#include <Grids.hpp>
#include "DBGridEh.hpp"
#include <Menus.hpp>

//---------------------------------------------------------------------------
class TZameshenie : public TForm
{
__published:	// IDE-managed Components
        TPanel *Panel1;
        TPanel *Panel2;
        TPanel *Panel3;
        TEdit *EditZEX;
        TEdit *EditTN;
        TLabel *LabelFIO;
        TBevel *Bevel1;
        TGroupBox *GroupBox1;
        TGroupBox *GroupBox2;
        TEdit *EditFIO_R;
        TLabel *LabelTN_R;
        TEdit *EditKPE1;
        TEdit *EditKPE2;
        TEdit *EditKPE3;
        TEdit *EditKPE4;
        TLabel *Label3;
        TLabel *Label4;
        TLabel *Label5;
        TLabel *Label6;
        TLabel *Label7;
        TLabel *Label8;
        TBevel *Bevel3;
        TBevel *Bevel4;
        TLabel *Label9;
        TImage *Image1;
        TBitBtn *BitBtn1;
        TBitBtn *BitBtn2;
        TLabel *Label17;
        TEdit *EditVZ_PENS;
        TBevel *Bevel6;
        TComboBox *ComboBoxRISK;
        TLabel *Label34;
        TLabel *Label36;
        TLabel *Label37;
        TEdit *EditRISK_PRICH;
        TBevel *Bevel7;
        TLabel *Label38;
        TLabel *LabelKPE;
        TLabel *LabelKRD;
        TLabel *Label40;
        TEdit *EditZEX_ZAM;
        TLabel *LabelZEX_ZAM;
        TLabel *Label42;
        TEdit *EditSHIFR_ZAM;
        TLabel *LabelDOLG_ZAM;
        TStringGrid *StringGrid1;
        TGroupBox *GroupBox3;
        TGroupBox *GroupBoxZAM;
        TCheckBox *CheckBoxREZERV;
        TDBGridEh *DBGridEh1;
        TCheckBox *CheckBoxZAM;
        TCheckBox *CheckBoxPREEM;
        TLabel *Label1;
        TBitBtn *BitBtn3;
        TBitBtn *BitBtn4;
        TBevel *Bevel5;
        TPopupMenu *PopupMenu1;
        TMenuItem *N1Dobav;
        TMenuItem *N2Redak;
        TMenuItem *N3;
        TMenuItem *N3Delete;
        TLabel *Label35;
        TComboBox *ComboBoxGOTOV;
        void __fastcall BitBtn2Click(TObject *Sender);
        void __fastcall FormShow(TObject *Sender);
        void __fastcall CheckBoxZAMClick(TObject *Sender);
        void __fastcall EditZEXKeyPress(TObject *Sender, char &Key);
        void __fastcall EditFIO_RKeyPress(TObject *Sender, char &Key);
        void __fastcall EditKPE1KeyPress(TObject *Sender, char &Key);
        void __fastcall BitBtn1Click(TObject *Sender);
        void __fastcall CheckBox1Click(TObject *Sender);
        void __fastcall CheckBox2Click(TObject *Sender);
        void __fastcall CheckBox3Click(TObject *Sender);
        void __fastcall CheckBox4Click(TObject *Sender);
        void __fastcall CheckBox5Click(TObject *Sender);
        void __fastcall CheckBox6Click(TObject *Sender);
        void __fastcall CheckBox7Click(TObject *Sender);
        void __fastcall CheckBox8Click(TObject *Sender);
        void __fastcall CheckBox9Click(TObject *Sender);
        void __fastcall CheckBox10Click(TObject *Sender);
        void __fastcall CheckBox11Click(TObject *Sender);
        void __fastcall CheckBox12Click(TObject *Sender);
        void __fastcall EditTNChange(TObject *Sender);
        void __fastcall EditZEXChange(TObject *Sender);
        void __fastcall FormCreate(TObject *Sender);
        void __fastcall ComboBoxGOTOVChange(TObject *Sender);
        void __fastcall ComboBoxRISKChange(TObject *Sender);
        void __fastcall EditZEX_ZAMChange(TObject *Sender);
        void __fastcall EditSHIFR_ZAMChange(TObject *Sender);
        void __fastcall EditFIO_RChange(TObject *Sender);
        void __fastcall EditZEXKeyDown(TObject *Sender, WORD &Key,
          TShiftState Shift);
        void __fastcall StringGrid1DrawCell(TObject *Sender, int ACol,
          int ARow, TRect &Rect, TGridDrawState State);
        void __fastcall StringGrid1SetEditText(TObject *Sender, int ACol,
          int ARow, const AnsiString Value);
        void __fastcall StringGrid2SetEditText(TObject *Sender, int ACol,
          int ARow, const AnsiString Value);
        void __fastcall StringGrid1SelectCell(TObject *Sender, int ACol,
          int ARow, bool &CanSelect);
        void __fastcall StringGrid1Enter(TObject *Sender);
        void __fastcall StringGrid1Exit(TObject *Sender);
        void __fastcall StringGrid1KeyPress(TObject *Sender, char &Key);
        void __fastcall StringGrid1KeyDown(TObject *Sender, WORD &Key,
          TShiftState Shift);
        void __fastcall CheckBoxREZERVClick(TObject *Sender);
        void __fastcall FormClose(TObject *Sender, TCloseAction &Action);
        void __fastcall N1DobavClick(TObject *Sender);
        void __fastcall N2RedakClick(TObject *Sender);
        void __fastcall BitBtn4Click(TObject *Sender);
        void __fastcall BitBtn3Click(TObject *Sender);
        void __fastcall FormKeyDown(TObject *Sender, WORD &Key,
          TShiftState Shift);
        void __fastcall DBGridEh1DblClick(TObject *Sender);
        void __fastcall N3DeleteClick(TObject *Sender);
private:	// User declarations
public:		// User declarations

       AnsiString zzex, ztn, zfio, zkpe1, zkpe2, zkpe3, zkpe4,
                  zfio_r, ztn_r, zdatn1, zdatn2, zdatn3,
                  zdatn4, zdatn5, zdatn6, zdatn7, zdatn8,
                  zdatn9, zdatn10, zdatn11, zdatn12,
                  zdatk1, zdatk2, zdatk3, zdatk4, zdatk5,
                  zdatk6, zdatk7, zdatk8, zdatk9, zdatk10,
                  zdatk11, zdatk12, zvz_pens, zrisk_prich,
                  zgotov, zrisk, zzam, zpreem, zzex_zam, zshifr_zam,
                  zrezerv,
                  id_dolg;

       int kol_str1;  // оличество заполненных строк в StringGrid
       int fl_red; //ѕризнак редактировани€ должности 0-добавление, 1-редактирование
       void __fastcall ZapolnenieInfo();

        __fastcall TZameshenie(TComponent* Owner);
};
//---------------------------------------------------------------------------
extern PACKAGE TZameshenie *Zameshenie;
//---------------------------------------------------------------------------
#endif
