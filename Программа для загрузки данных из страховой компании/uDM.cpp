//---------------------------------------------------------------------------

#include <vcl.h>
#pragma hdrstop

#include "uDM.h"
#include "uData.h"
#include "uMain.h"
#include "FuncCrypt.h"
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
void __fastcall TDM::DataModuleCreate(TObject *Sender)
{
  int mm1, yyyy1;
  
   AnsiString S;

  //Считывание строки соединения из зашифрованного файла
  try
     {
       DecryptFromFile(GetCurrentDir() + "\\certificate.1.13.m.dat", S);
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


  Application->UpdateFormatSettings = false;
  DateSeparator = '.';
  DecimalSeparator = '.';


//Отчетный месяц

  // Текущая дата
  DecodeDate(Date(), year, month, day);
  mm1 = month;
  yyyy1 = year;

  //Текущий месяц
  mm = mm1;
  yyyy = yyyy1;

  PrevMonth(mm1, yyyy1);

  // Предыдущий месяц
  mm2 = mm1;
  yyyy2 = yyyy1;


  //Установка разделителя для дробного числа ',' для текущей сессии Oracle
  qObnovlenie->Close();
  qObnovlenie->SQL->Clear();
  qObnovlenie->SQL->Add("alter session set NLS_NUMERIC_CHARACTERS=',.'");
  try
    {
      qObnovlenie->ExecSQL();
    }
  catch(Exception &E)
    {
      Application->MessageBox(("Невозможно установить разделитель '.'\nдля дробного числа  для текущей сессии Oracle:\n" + E.Message).c_str(),"Соединение с сервером",
                              MB_OK+MB_ICONERROR);
      Application->Terminate();
      Abort();
    }

}
//---------------------------------------------------------------------------

void __fastcall TDM::DataModuleDestroy(TObject *Sender)
{
  ADOConnection1->Close();
  qZagruzka->Active = false;
  qObnovlenie->Active = false;
  qKorrektirovka->Active = false;
}
//---------------------------------------------------------------------------

//---------------------------------------------------------------------------
// Предыдущий месяц
void __fastcall TDM::PrevMonth(int &Month, int &Year, int k)
{
  for (int i=1; i<=k; i++) {
    if (Month==1) { Month = 12; Year--; }
    else Month--;
  }
}
//---------------------------------------------------------------------------
void __fastcall TDM::dsKorrektirovkaDataChange(TObject *Sender,
      TField *Field)
{
  Main->SetEditData();        
}
//---------------------------------------------------------------------------

