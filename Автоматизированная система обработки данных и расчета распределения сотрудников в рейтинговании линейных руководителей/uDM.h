//---------------------------------------------------------------------------

#ifndef uDMH
#define uDMH
//---------------------------------------------------------------------------
#include <System.Classes.hpp>
#include "DBAccess.hpp"
#include "MemDS.hpp"
#include "OracleUniProvider.hpp"
#include "Uni.hpp"
#include "UniProvider.hpp"
#include <Data.DB.hpp>
#include <SHDocVw.hpp>

//---------------------------------------------------------------------------
class TDM : public TDataModule
{
__published:	// IDE-managed Components
	TUniConnection *UniConnection1;
	TUniQuery *qObnovlenie;
	TUniQuery *qReiting;
	TOracleUniProvider *OracleUniProvider1;
	TDataSource *dsReiting;
	TUniQuery *qProverka;
	TUniQuery *qObnovlenie2;
	TUniQuery *qRaschet;
	TUniQuery *qSprav;
	TDataSource *dsSprav;
	void __fastcall DataModuleCreate(TObject *Sender);
	void __fastcall DataModuleDestroy(TObject *Sender);
private:	// User declarations
public:		// User declarations
	__fastcall TDM(TComponent* Owner);
};
//---------------------------------------------------------------------------
extern PACKAGE TDM *DM;
//---------------------------------------------------------------------------
#endif
