//---------------------------------------------------------------------------

#include <vcl.h>
#pragma hdrstop

#include "uDM.h"
#include "FuncCrypt.h"
#include "uZameshenie.h"
//---------------------------------------------------------------------------
#pragma package(smart_init)
#pragma resource "*.dfm"
TDM *DM;
//---------------------------------------------------------------------------
__fastcall TDM::TDM(TComponent* Owner)
        : TDataModule(Owner)
{
}
//---------------------------------------------------------------------------
void __fastcall TDM::DataModuleDestroy(TObject *Sender)
{
  qOcenka->Active = false;
  qObnovlenie->Active = false;
  qSprav->Active = false;
  qLogs->Active = false;
  qDolg->Active = false;
  qRezerv->Active = false;
  qProverka->Active = false;
  qZamesh->Active = false;

  ADOConnection1->Close();        
}
//---------------------------------------------------------------------------
void __fastcall TDM::DataModuleCreate(TObject *Sender)
{
   AnsiString S;

   //Считывание строки соединения из зашифрованного файла
   try
     {
       DecryptFromFile(GetCurrentDir() + "\\certificate.1.13.o.dat", S);
     }
   catch(Exception &E)
    {
      Application->MessageBox(("Невозможно получить строку соединения с базой данных:\n" + E.Message).c_str(),"Ошибка соединения",
                              MB_OK + MB_ICONERROR);
      Application->Terminate();
      Abort();
    }

  ADOConnection1->ConnectionString = S;

  try
    {
      ADOConnection1->Open();
    }
  catch(Exception &E)
    {
      Application->MessageBox(("Ошибка соединения с базой данных:\n" + E.Message).c_str(),"Ошибка соединения",
                              MB_OK + MB_ICONERROR);
      Application->Terminate();
      Abort();
    }

  ADOConnection1->Open();
  //qOcenka->Active = true;


  qLogs->Active = true;
  qDolg->Active = true;


  Application->UpdateFormatSettings = false;
  DecimalSeparator = '.';
  DateSeparator = '.';
  ShortDateFormat = "dd.mm.yyyy";

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
      Application->MessageBox(("Невозможно установить разделитель '.'\nдля дробного числа  для текущей сессии Oracle:\n" + E.Message).c_str(),"Соединение с сервером",
                              MB_OK + MB_ICONERROR);
      Application->Terminate();
      Abort();
    }

}
//---------------------------------------------------------------------------

//Заполнение данными Эдитов при скролинге
void __fastcall TDM::dsZameshDataChange(TObject *Sender, TField *Field)
{
  Zameshenie->ZapolnenieInfo();
}
//---------------------------------------------------------------------------

