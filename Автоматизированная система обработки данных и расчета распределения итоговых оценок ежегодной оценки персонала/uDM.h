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
        TADOConnection *ADOConnection1;
        TDataSource *dsOcenka;
        TADOQuery *qOcenka;
        TADOQuery *qObnovlenie;
        TADOQuery *qSprav;
        TDataSource *dsSprav;
        TADOQuery *qObnovlenie2;
        TADOStoredProc *spOcenka;
        TADOQuery *qLogs;
        TDataSource *dsLogs;
        TADOQuery *qDolg;
        TDataSource *dsDolg;
        TADOQuery *qRezerv;
        TDataSource *dsRezerv;
        TADOQuery *qProverka;
        TDataSource *dsZamesh;
        TADOQuery *qZamesh;
        void __fastcall DataModuleDestroy(TObject *Sender);
        void __fastcall DataModuleCreate(TObject *Sender);
        void __fastcall dsZameshDataChange(TObject *Sender, TField *Field);
private:	// User declarations
public:		// User declarations
        
        __fastcall TDM(TComponent* Owner);
};
//---------------------------------------------------------------------------
extern PACKAGE TDM *DM;
//---------------------------------------------------------------------------
#endif
