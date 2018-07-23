//---------------------------------------------------------------------------

#ifndef uReitingH
#define uReitingH
//---------------------------------------------------------------------------
#include <Classes.hpp>
#include <Controls.hpp>
#include <StdCtrls.hpp>
#include <Forms.hpp>
#include <ExtCtrls.hpp>
#include <Buttons.hpp>
#include <jpeg.hpp>
#include <IdGlobal.hpp>
//---------------------------------------------------------------------------
class TReiting : public TForm
{
__published:	// IDE-managed Components
        TPanel *Panel1;
        TPanel *Panel2;
        TBevel *Bevel1;
        TBevel *Bevel2;
        TBevel *Bevel4;
        TImage *Image1;
        TBevel *Bevel5;
        TLabel *Label1;
        TEdit *EditTN;
        TLabel *LabelZEX;
        TLabel *Label3;
        TLabel *Label4;
        TLabel *Label5;
        TLabel *Label6;
        TEdit *EditREZ;
        TEdit *EditKE;
        TEdit *EditKOM;
        TGroupBox *GroupBox1;
        TButton *Button1;
        TButton *Cansel;
        TLabel *Label7;
        TLabel *Label8;
        TLabel *LabelREZ;
        TLabel *LabelKOMP;
        TLabel *LabelEFFEKT;
        TLabel *Label9;
        TLabel *LabelFIO_OCEN;
        TLabel *LabelNZEX;
        TLabel *Label2;
        TGroupBox *GroupBox2;
        TLabel *Label10;
        TLabel *Label11;
        TLabel *Label12;
        TLabel *Label13;
        TLabel *Label14;
        TLabel *Label15;
        TLabel *Label16;
        TLabel *Label17;
        TBevel *Bevel3;
        TGroupBox *GroupBox3;
        TLabel *Label18;
        TLabel *Label19;
        TLabel *Label20;
        TEdit *EditREALIZAC;
        TEdit *EditKACHESTVO;
        TEdit *EditRESURS;
        TEdit *EditSTAND;
        TEdit *EditPOTREB;
        TEdit *EditKACH;
        TEdit *EditEFF;
        TEdit *EditPROF_ZN;
        TEdit *EditLIDER;
        TEdit *EditOTVETSTV;
        TEdit *EditKOM_REZ;
        TBevel *Bevel6;
        void __fastcall SpeedButton2Click(TObject *Sender);
        void __fastcall SpeedButton1Click(TObject *Sender);
        void __fastcall FormKeyDown(TObject *Sender, WORD &Key,
          TShiftState Shift);
        void __fastcall Button1Click(TObject *Sender);
        void __fastcall CanselClick(TObject *Sender);
        void __fastcall FormShow(TObject *Sender);
        void __fastcall EditREZExit(TObject *Sender);
        void __fastcall EditKEKeyPress(TObject *Sender, char &Key);
        void __fastcall EditZEXKeyPress(TObject *Sender, char &Key);
        void __fastcall EditTNChange(TObject *Sender);
        void __fastcall EditKEExit(TObject *Sender);
        void __fastcall EditREZChange(TObject *Sender);
        void __fastcall EditKOMChange(TObject *Sender);
        void __fastcall EditKEChange(TObject *Sender);
        void __fastcall EditKOMExit(TObject *Sender);
        void __fastcall EditZEXKeyDown(TObject *Sender, WORD &Key,
          TShiftState Shift);
        void __fastcall EditREALIZACExit(TObject *Sender);
        void __fastcall EditKACHESTVOExit(TObject *Sender);
        void __fastcall EditRESURSExit(TObject *Sender);
        void __fastcall EditREALIZACChange(TObject *Sender);
        void __fastcall EditSTANDChange(TObject *Sender);
        void __fastcall EditSTANDExit(TObject *Sender);
        void __fastcall EditPOTREBExit(TObject *Sender);
        void __fastcall EditKACHExit(TObject *Sender);
        void __fastcall EditEFFExit(TObject *Sender);
        void __fastcall EditPROF_ZNExit(TObject *Sender);
        void __fastcall EditLIDERExit(TObject *Sender);
        void __fastcall EditOTVETSTVExit(TObject *Sender);
        void __fastcall EditKOM_REZExit(TObject *Sender);
        void __fastcall EditKACHESTVOChange(TObject *Sender);
private:	// User declarations
public:		// User declarations

        double rezult, komp, effekt;
        AnsiString zrez, zke, zkom,
                   zrealizac, zkachestvo, zresurs,
                   zstand, zpotreb, zkach, zeff, zprof_zn, zlider,
                   zotvetstv, zkom_rez;

        void __fastcall IzmenRezRab();    //Изменение значений по результатам работы
        void __fastcall IzmenKomp();      //Изменение значений по компетенциям
        __fastcall TReiting(TComponent* Owner);

};
//---------------------------------------------------------------------------
extern PACKAGE TReiting *Reiting;
//---------------------------------------------------------------------------
#endif
