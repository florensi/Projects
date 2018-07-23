//---------------------------------------------------------------------------

#ifndef uSpravH
#define uSpravH
//---------------------------------------------------------------------------
#include <System.Classes.hpp>
#include <Vcl.Controls.hpp>
#include <Vcl.StdCtrls.hpp>
#include <Vcl.Forms.hpp>
#include <Vcl.ExtCtrls.hpp>
#include "DBAxisGridsEh.hpp"
#include "DBGridEh.hpp"
#include "DBGridEhGrouping.hpp"
#include "DBGridEhToolCtrls.hpp"
#include "DynVarsEh.hpp"
#include "EhLibVCL.hpp"
#include "GridsEh.hpp"
#include "ToolCtrlsEh.hpp"
#include <Vcl.Buttons.hpp>
#include <Vcl.Menus.hpp>
#include <Vcl.Imaging.jpeg.hpp>
#include <IdGlobal.hpp>
//---------------------------------------------------------------------------
class TSprav : public TForm
{
__published:	// IDE-managed Components
	TPanel *Panel1;
	TBevel *Bevel1;
	TBevel *Bevel2;
	TImage *Image1;
	TBitBtn *BitBtn2;
	TDBGridEh *DBGridEh1;
	TPopupMenu *PopupMenu1;
	TMenuItem *N1;
	TMenuItem *N2;
	TMenuItem *N3;
	TMenuItem *N4;
	TLabel *Label1;
	TGroupBox *GroupBoxDOBAV;
	TBitBtn *BitBtn1;
	TBitBtn *BitBtn3;
	TLabel *Label2;
	TEdit *EditZEX;
	TLabel *Label3;
	TLabel *LabelNZEX;
	TComboBox *ComboBoxPZ;
	void __fastcall BitBtn3Click(TObject *Sender);
	void __fastcall BitBtn1Click(TObject *Sender);
	void __fastcall BitBtn2Click(TObject *Sender);
	void __fastcall N1Click(TObject *Sender);
	void __fastcall N2Click(TObject *Sender);
	void __fastcall DBGridEh1DblClick(TObject *Sender);
	void __fastcall FormKeyDown(TObject *Sender, WORD &Key, TShiftState Shift);
	void __fastcall EditZEXKeyPress(TObject *Sender, System::WideChar &Key);
	void __fastcall EditZEXChange(TObject *Sender);
	void __fastcall N4Click(TObject *Sender);
private:	// User declarations
public:		// User declarations

	  int pr_in; //Признак вставки/редактирования
	  String szex, spz, snzex;

	__fastcall TSprav(TComponent* Owner);
};
//---------------------------------------------------------------------------
extern PACKAGE TSprav *Sprav;
//---------------------------------------------------------------------------
#endif
