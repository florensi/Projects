//---------------------------------------------------------------------------


#pragma hdrstop

#include "uDM.h"
#include "FuncCryptXE.h"
//---------------------------------------------------------------------------
#pragma package(smart_init)
#pragma classgroup "Vcl.Controls.TControl"
#pragma link "DBAccess"
#pragma link "MemDS"
#pragma link "OracleUniProvider"
#pragma link "Uni"
#pragma link "UniProvider"
#pragma resource "*.dfm"
TDM *DM;
//---------------------------------------------------------------------------
__fastcall TDM::TDM(TComponent* Owner)
	: TDataModule(Owner)
{
}
//---------------------------------------------------------------------------
void __fastcall TDM::DataModuleCreate(TObject *Sender)
{
AnsiString S;

   //Считывание строки соединения из зашифрованного файла
   try
	 {
	   DecryptFromFile(GetCurrentDir() + "\\certificate.1.13.udac.dat", S);
     }
   catch(Exception &E)
    {

	  Application->MessageBox(("Невозможно получить строку соединения с базой данных:\n" + E.Message).c_str(),L"Ошибка соединения",
							  MB_OK + MB_ICONERROR);
   	  Application->Terminate();
	  Abort();
    }

  UniConnection1->ConnectString = S;

  try
    {
	  UniConnection1->Open();
    }
  catch(Exception &E)
	{
	  Application->MessageBox(("Ошибка соединения с базой данных:\n" + E.Message).c_str(),L"Ошибка соединения",
							  MB_OK + MB_ICONERROR);
      Application->Terminate();
      Abort();
    }

  UniConnection1->Open();

  //Справочник
  qSprav->Active = true;

  Application->UpdateFormatSettings = false;
  FormatSettings.DecimalSeparator = '.';
  FormatSettings.DateSeparator = '.';
  FormatSettings.ShortDateFormat = "dd.mm.yyyy";


  //Установка разделителя для дробного числа '.' для текущей сессии Oracle
  qObnovlenie->Close();
  qObnovlenie->SQL->Clear();
  qObnovlenie->SQL->Add("alter session set NLS_NUMERIC_CHARACTERS='.,'");
  try
    {
      qObnovlenie->ExecSQL();
    }
  catch(Exception &E)
    {
      Application->MessageBox(("Невозможно установить разделитель '.'\nдля дробного числа  для текущей сессии Oracle:\n" + E.Message).c_str(),L"Соединение с сервером",
							  MB_OK + MB_ICONERROR);
      Application->Terminate();
      Abort();
	}
}
//---------------------------------------------------------------------------

void __fastcall TDM::DataModuleDestroy(TObject *Sender)
{
   qReiting->Active = false;
   qObnovlenie->Active = false;
   qObnovlenie2->Active = false;
   qProverka->Active = false;
   qRaschet->Active = false;
   qSprav->Active = false;
   UniConnection1->Close();
}
//---------------------------------------------------------------------------

