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
//---------------------------------------------------------------------------
class TDM : public TDataModule
{
__published:	// IDE-managed Components
        TADOQuery *qObnovlenie;
        TDataSource *dsKomandirovki;
        TADOQuery *qKomandirovki;
        TDataSource *dsSP_chel;
        TDataSource *dsSP_grade;
        TDataSource *dsSP_gostinica;
        TDataSource *dsSP_obekt;
        TDataSource *dsSP_country;
        TDataSource *dsSP_city;
        TADOQuery *qSP_chel;
        TADOQuery *qSP_grade;
        TADOQuery *qSP_gostinica;
        TADOQuery *qSP_obekt;
        TADOQuery *qSP_country;
        TADOQuery *qSP_city;
        TADOQuery *qObnovlenie1;
        TADOConnection *ADOConnection2;
        TADOConnection *ADOConnection1;
        void __fastcall DataModuleCreate(TObject *Sender);
        void __fastcall DataModuleDestroy(TObject *Sender);
private:	// User declarations
public:		// User declarations
        __fastcall TDM(TComponent* Owner);
        HANDLE mutex1;
};
//---------------------------------------------------------------------------
extern PACKAGE TDM *DM;
//---------------------------------------------------------------------------
#endif
