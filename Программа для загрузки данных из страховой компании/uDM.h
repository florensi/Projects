//---------------------------------------------------------------------------

#ifndef uDMH
#define uDMH
//---------------------------------------------------------------------------
#include <Classes.hpp>
#include <Controls.hpp>
#include <StdCtrls.hpp>
#include <Forms.hpp>
#include <ADODB.hpp>
#include <DB.hpp>
#include <DBXpress.hpp>
#include <FMTBcd.hpp>
#include <SqlExpr.hpp>
//---------------------------------------------------------------------------
class TDM : public TDataModule
{
__published:	// IDE-managed Components
        TADOConnection *ADOConnection1;
        TADOQuery *qZagruzka;
        TADOQuery *qObnovlenie;
        TADOQuery *qKorrektirovka;
        TDataSource *dsKorrektirovka;
        void __fastcall DataModuleCreate(TObject *Sender);
        void __fastcall DataModuleDestroy(TObject *Sender);
        void __fastcall dsKorrektirovkaDataChange(TObject *Sender,
          TField *Field);
private:	// User declarations
public:		// User declarations

        int mm, yyyy;   //Отчетный период из grafr
        int mm2, yyyy2;  // Отчетный период с предыдущим месяцем от grafr
        Word year, month, day; // Текущая дата
        
        void __fastcall PrevMonth(int &Month, int &Year, int k=1);
        __fastcall TDM(TComponent* Owner);
};
//---------------------------------------------------------------------------
extern PACKAGE TDM *DM;
//---------------------------------------------------------------------------
#endif
