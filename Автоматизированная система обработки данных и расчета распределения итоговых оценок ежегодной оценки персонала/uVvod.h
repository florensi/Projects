//---------------------------------------------------------------------------

#ifndef uVvodH
#define uVvodH
//---------------------------------------------------------------------------
#include <Classes.hpp>
#include <Controls.hpp>
#include <StdCtrls.hpp>
#include <Forms.hpp>
#include <ExtCtrls.hpp>
#include <jpeg.hpp>
#include <IdGlobal.hpp>
//---------------------------------------------------------------------------
class TVvod : public TForm
{
__published:	// IDE-managed Components
        TPanel *Panel1;
        TPanel *Panel2;
        TBevel *Bevel1;
        TBevel *Bevel2;
        TBevel *Bevel3;
        TBevel *Bevel4;
        TButton *Button1;
        TButton *Cansel;
        TImage *Image1;
        TLabel *Label1;
        TLabel *Label3;
        TLabel *Label4;
        TGroupBox *GroupBox1;
        TGroupBox *GroupBox2;
        TGroupBox *GroupBox3;
        TLabel *LabelZEX_NAIM;
        TLabel *Label6;
        TLabel *Label7;
        TLabel *Label8;
        TLabel *Label9;
        TEdit *EditZEX;
        TEdit *EditTN;
        TComboBox *ComboBoxFUNCT_G;
        TComboBox *ComboBoxKAT;
        TLabel *Label10;
        TLabel *Label11;
        TLabel *Label12;
        TLabel *LabelREZULT_PROC;
        TLabel *LabelKOMP_PROC;
        TLabel *Label16;
        TLabel *LabelEFFEKT;
        TLabel *Label18;
        TLabel *Label19;
        TLabel *Label21;
        TEdit *EditREZULT_OCEN;
        TEdit *EditKPE_OCEN;
        TEdit *EditKOMP_OCEN;
        TEdit *EditFIO_OCEN;
        TEdit *EditDOLGO;
        TEdit *EditDATA_OCEN;
        TBevel *Bevel5;
        TLabel *Label20;
        TLabel *Label23;
        TLabel *Label24;
        TEdit *EditAVT_REIT;
        TEdit *EditSKOR_REIT;
        TEdit *EditKOM_REIT;
        TComboBox *ComboBoxUU;
        TComboBox *ComboBoxFUNCT;
        TBevel *Bevel6;
        TLabel *Label2;
        TEdit *EditDIREKT;
        TLabel *LabelDIREKT;
        TLabel *LabelNAME_DOLG;
        TEdit *EditFIO;
        TLabel *LabelDAT_JOB;
        TGroupBox *GroupBox4;
        TLabel *Label5;
        TLabel *Label13;
        TLabel *Label14;
        TLabel *Label15;
        TLabel *Label17;
        TLabel *Label22;
        TLabel *Label25;
        TLabel *Label26;
        TEdit *EditSTAND;
        TEdit *EditPOTREB;
        TEdit *EditKACH;
        TEdit *EditEFF;
        TEdit *EditPROF_ZN;
        TEdit *EditLIDER;
        TEdit *EditOTVETSTV;
        TEdit *EditKOM_REZ;
        TGroupBox *GroupBox5;
        TLabel *Label27;
        TLabel *Label28;
        TLabel *Label29;
        TEdit *EditREALIZAC;
        TEdit *EditKACHESTVO;
        TEdit *EditRESURS;
        TBevel *Bevel8;
        TBevel *Bevel7;
        void __fastcall CanselClick(TObject *Sender);
        void __fastcall FormShow(TObject *Sender);
        void __fastcall FormKeyDown(TObject *Sender, WORD &Key,
          TShiftState Shift);
        void __fastcall EditREZULT_OCENExit(TObject *Sender);
        void __fastcall EditZEXKeyPress(TObject *Sender, char &Key);
        void __fastcall Button1Click(TObject *Sender);
        void __fastcall EditDATA_OCENExit(TObject *Sender);
        void __fastcall EditREZULT_OCENChange(TObject *Sender);
        void __fastcall EditKOMP_OCENChange(TObject *Sender);
        void __fastcall EditKPE_OCENChange(TObject *Sender);
        void __fastcall EditKPE_OCENExit(TObject *Sender);
        void __fastcall EditREZULT_OCENKeyPress(TObject *Sender,
          char &Key);
        void __fastcall EditFIO_OCENKeyPress(TObject *Sender, char &Key);
        void __fastcall EditDIREKTChange(TObject *Sender);
        void __fastcall EditKOMP_OCENExit(TObject *Sender);
        void __fastcall EditZEXChange(TObject *Sender);
        void __fastcall EditTNChange(TObject *Sender);
        void __fastcall EditDIREKTExit(TObject *Sender);
        void __fastcall EditZEX_REZKeyPress(TObject *Sender, char &Key);
        void __fastcall EditSHIFR_REZKeyPress(TObject *Sender, char &Key);
        void __fastcall EditZEXKeyDown(TObject *Sender, WORD &Key,
          TShiftState Shift);
        void __fastcall EditREALIZACExit(TObject *Sender);
        void __fastcall EditKACHESTVOExit(TObject *Sender);
        void __fastcall EditRESURSExit(TObject *Sender);
        void __fastcall EditSTANDExit(TObject *Sender);
        void __fastcall EditPOTREBExit(TObject *Sender);
        void __fastcall EditKACHExit(TObject *Sender);
        void __fastcall EditEFFExit(TObject *Sender);
        void __fastcall EditPROF_ZNExit(TObject *Sender);
        void __fastcall EditLIDERExit(TObject *Sender);
        void __fastcall EditOTVETSTVExit(TObject *Sender);
        void __fastcall EditKOM_REZExit(TObject *Sender);
        void __fastcall EditREALIZACChange(TObject *Sender);
        void __fastcall EditSTANDChange(TObject *Sender);
private:	// User declarations
public:		// User declarations
        double rezult, komp, effekt;
        AnsiString zzex, ztn, zdirekt, zrezult_ocen, zkpe_ocen, zkomp_ocen,
                   zdata_ocen, zfio_ocen, zdolgo, zavt_reit, zskor_ocen,
                   zkom_reit,  zdolg_rezerv, zkat, zfunct_g, zfunct,
                   znaim_dolg, zuu, zfio, zzex_rez, zshifr_rez, zdat_job,
                   zrealizac, zkachestvo, zresurs,
                   zstand, zpotreb, zkach, zeff, zprof_zn, zlider,
                   zotvetstv, zkom_rez;

        AnsiString uch, nuch;
        void __fastcall SetDataEdit();
        void __fastcall IzmenRezRab();
        void __fastcall IzmenKomp();
        __fastcall TVvod(TComponent* Owner);
};
//---------------------------------------------------------------------------
extern PACKAGE TVvod *Vvod;
//---------------------------------------------------------------------------
#endif
