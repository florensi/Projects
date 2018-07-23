//---------------------------------------------------------------------------

#ifndef uMainH
#define uMainH
//---------------------------------------------------------------------------
#include <Classes.hpp>
#include <Controls.hpp>
#include <StdCtrls.hpp>
#include <Forms.hpp>
#include <Buttons.hpp>
#include <comobj.hpp>
#include <ComCtrls.hpp>
#include <Menus.hpp>
#include <DateUtils.hpp>
#include <IdGlobal.hpp>


#include <Registry.hpp>
#include <ExtCtrls.hpp>
#include <jpeg.hpp>
#include <Dialogs.hpp>


#include <Variants.hpp>


#include <FileCtrl.hpp>
#include "DBGridEh.hpp"
#include <Grids.hpp>  // Для диалогового окна выбора папки

//---------------------------------------------------------------------------
class TMain : public TForm
{
__published:	// IDE-managed Components
        TStatusBar *StatusBar1;
        TMainMenu *MainMenu1;
        TMenuItem *N1;
        TMenuItem *N2;
        TMenuItem *N3;
        TMenuItem *NewDogov;
        TImage *Image1;
        TMenuItem *N6;
        TMenuItem *izm_grn;
        TMenuItem *izm_val;
        TMenuItem *kurs_pereschet;
        TMenuItem *N10;
        TOpenDialog *OpenDialog1;
        TMenuItem *N4;
        TMenuItem *N7;
        TMenuItem *N9;
        TPanel *Panel1;
        TBitBtn *BitBtn1;
        TBitBtn *BitBtn2;
        TImage *Image2;
        TBevel *Bevel1;
        TEdit *EditZEX;
        TEdit *EditTN;
        TEdit *EditZEX2;
        TEdit *EditTN2;
        TBitBtn *BitBtn3;
        TLabel *Label1;
        TBevel *Bevel2;
        TEdit *EditSum;
        TEdit *EditVal;
        TEdit *EditData_s;
        TEdit *EditData_po;
        TLabel *Label2;
        TLabel *Label3;
        TLabel *Label4;
        TLabel *Label5;
        TLabel *Label6;
        TLabel *Label7;
        TDBGridEh *DBGridEh1;
        TLabel *Label8;
        TLabel *Label9;
        TBevel *Bevel3;
        TLabel *Label10;
        TSpeedButton *SpeedButton1;
        TMenuItem *N11;
        TMenuItem *N12;
        TMenuItem *N13;
        TMenuItem *N14;
        TMenuItem *N15;
        TMenuItem *N16;
        TMenuItem *N17;
        TMenuItem *N18;
        TLabel *Label11;
        TLabel *Label12;
        TMenuItem *N19;
        TMenuItem *N20;
        TMenuItem *N21;
        TLabel *LabelNDOG;
        TLabel *Label14;
        TEdit *EditNDOG;
        TMenuItem *N22;
        TMenuItem *N23;
        TMenuItem *N24;
        TMenuItem *N25;
        TEdit *EditPRIZNAK;
        TMenuItem *Excel1;
        TMenuItem *N00151;
        TMenuItem *N26;
        TMenuItem *N5;
        TMenuItem *N8;
        TMenuItem *N27;
        TMenuItem *N28;
        TMenuItem *N29;
        TMenuItem *N30;
        TMenuItem *N31;
        TMenuItem *N32;
        TMenuItem *N33;
        TMenuItem *N34;
        TMenuItem *N35;
        TMenuItem *N36;

        void __fastcall N2Click(TObject *Sender);
        void __fastcall NewDogovClick(TObject *Sender);
        void __fastcall FormCreate(TObject *Sender);
        void __fastcall izm_valClick(TObject *Sender);
        void __fastcall izm_grnClick(TObject *Sender);
        void __fastcall kurs_pereschetClick(TObject *Sender);
        void __fastcall FormShow(TObject *Sender);
        void __fastcall BitBtn1Click(TObject *Sender);
        void __fastcall BitBtn2Click(TObject *Sender);
        void __fastcall BitBtn3Click(TObject *Sender);
        void __fastcall FormKeyDown(TObject *Sender, WORD &Key,
          TShiftState Shift);
        void __fastcall EditZEXKeyPress(TObject *Sender, char &Key);
        void __fastcall EditSumKeyPress(TObject *Sender, char &Key);
        void __fastcall EditData_sExit(TObject *Sender);
        void __fastcall EditData_poExit(TObject *Sender);
        void __fastcall SpeedButton1Click(TObject *Sender);
        void __fastcall N7Click(TObject *Sender);
        void __fastcall DBGridEh1DrawColumnCell(TObject *Sender,
          const TRect &Rect, int DataCol, TColumnEh *Column,
          TGridDrawState State);
        void __fastcall N11Click(TObject *Sender);
        void __fastcall N13Click(TObject *Sender);
        void __fastcall N15Click(TObject *Sender);
        void __fastcall N17Click(TObject *Sender);
        void __fastcall N14Click(TObject *Sender);
        void __fastcall N16Click(TObject *Sender);
        void __fastcall N18Click(TObject *Sender);
        void __fastcall N19Click(TObject *Sender);
        void __fastcall N20Click(TObject *Sender);
        void __fastcall N21Click(TObject *Sender);
        void __fastcall EditNDOGKeyPress(TObject *Sender, char &Key);
        void __fastcall N23Click(TObject *Sender);
        void __fastcall N24Click(TObject *Sender);
        void __fastcall N25Click(TObject *Sender);
        void __fastcall EditPRIZNAKKeyPress(TObject *Sender, char &Key);
        void __fastcall N5Click(TObject *Sender);
        void __fastcall N8Click(TObject *Sender);
        void __fastcall N27Click(TObject *Sender);
        void __fastcall N28Click(TObject *Sender);
        void __fastcall N29Click(TObject *Sender);
        void __fastcall N30Click(TObject *Sender);
        void __fastcall N31Click(TObject *Sender);
        void __fastcall N32Click(TObject *Sender);
        void __fastcall N34Click(TObject *Sender);
        void __fastcall N36Click(TObject *Sender);
private:	// User declarations
        TProgressBar *ProgressBar;
public:		// User declarations
        __fastcall TMain(TComponent* Owner);
        AnsiString __fastcall SetNull (AnsiString str, AnsiString r="NULL");
        AnsiString zex;

        AnsiString Path, WordPath, DocPath, TempPath, WorkPath, Dir2, FileName,
                   UserName, DomainName, UserFullName, TN;

        AnsiString zzex, ztn, zsum, zdata_s, zdata_po, zval, zpriznak; //Для проверки было ли редактирование
        AnsiString ob_kol, obnov_kol; // Для логов. Общее количество записей и обновленное количество записей

        Variant Excel, Book, Sheet;
        int Row; //Количество занятых строк в Excel
        int im_fl; //Заголовок диалогового окна при выборе  файла
        int rec;
        int fl_r; //fl=1 - редактирование, fl=0 - вставка
        Double prozhitMin; // Сумма прожиточного минимума

        Word dtp_day, dtp_month, dtp_year; // дата из DateTimePicker
        AnsiString dtp_mm; // дата из DateTimePicker c "0"
        int ana;  //ПАО ММК им.Ильича ana=1, РМЗ ana=4 

        AnsiString __fastcall FindWordPath();

        bool __fastcall Proverka(AnsiString zex);
        bool __fastcall GetMyDocumentsDir(AnsiString &FolderPath);  // Возвращает путь на папку Мои документы"
        bool __fastcall GetTempDir(AnsiString &FolderPath);         // Возвращает путь на папку Temp


        void __fastcall ProverkaInfoExcel();     // Проверка правильности цеха, тн и суммы страхования на превышение 15%
        void __fastcall ProverkaInfoExcelIzmeneniya();   // Проверка на существование инн в таблице avans
        void __fastcall UpdateValuta_I_Grivna();  // Обновление изменений по валюте и гривне

        void __fastcall SetEditData();
        void __fastcall SetEditNull();
        void __fastcall ProverkaProzhitMin();   // Получение суммы прожиточного минимума
        void __fastcall EventsMessage(tagMSG &Msg, bool &Handled);
        void __fastcall InsertLog(AnsiString Msg);
        void __fastcall OtchetStrahovaya(int valuta);

        void __fastcall ExcelSAP(int valuta);

};
//---------------------------------------------------------------------------
extern PACKAGE TMain *Main;
//---------------------------------------------------------------------------
#endif
