//---------------------------------------------------------------------------

#include <vcl.h>
#pragma hdrstop

#include "uDM.h"
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
 AnsiString S;

  //Проверка на запущенное на компьютере приложение
  try
    {
      mutex1 = CreateMutex(NULL,false,"Komandirovki");
      int Res = GetLastError();
      if (Res == ERROR_ALREADY_EXISTS)
        {
          Application->MessageBox("Данное приложение уже запущено на этом компьютере!!!","Предупреждение",
                                  MB_OK+MB_ICONWARNING);
          Application->Terminate();
          Abort();
        }
      if (Res == ERROR_INVALID_HANDLE)
        {
          Application->MessageBox("Ошибка в имени mutex","Предупреждение",
                                  MB_OK+MB_ICONWARNING);
          Application->Terminate();
          Abort();
        }
    }
  catch(...)
    {
      Application->MessageBox("Возникла ошибка при проверке на запущенный второй экземпляр программы!!!","Предупреждение",
                                  MB_OK+MB_ICONWARNING);
      Application->Terminate();
      Abort();
    }

  //Считывание строки соединения из зашифрованного файла
  if (!DecryptFromFile(GetCurrentDir() + "\\certificate.1.1.m.dat", S))
    {
      Application->MessageBox("Невозможно получить строку соединения с базой данных","Ошибка соединения",
                               MB_OK + MB_ICONERROR);
      Application->Terminate();
      Abort();
    }

  ADOConnection1->ConnectionString = S;

  try
    {
      ADOConnection1->Open();
    }
  catch(...)
    {
      Application->MessageBox("Ошибка соединения с базой данных","Ошибка соединения",
                              MB_OK + MB_ICONERROR);
      Application->Terminate();
      Abort();
    }

  //Считывание строки соединения из зашифрованного файла
  if (!DecryptFromFile(GetCurrentDir() + "\\certificate.1.13.m.dat", S))
    {
      Application->MessageBox("Невозможно получить строку соединения с базой данных","Ошибка соединения",
                               MB_OK + MB_ICONERROR);
      Application->Terminate();
      Abort();
    }

  ADOConnection2->ConnectionString = S;

  try
    {
      ADOConnection1->Open();
    }
  catch(...)
    {
      Application->MessageBox("Ошибка соединения с базой данных","Ошибка соединения",
                              MB_OK + MB_ICONERROR);
      Application->Terminate();
      Abort();
    }


  Application->UpdateFormatSettings = false;
  DateSeparator = '.';
  DecimalSeparator = '.';

  ShortDateFormat = "dd.mm.yyyy";

  //Установка разделителя для дробного числа '.' для текущей сессии Oracle
  qObnovlenie->Close();
  qObnovlenie->SQL->Clear();
  qObnovlenie->SQL->Add("alter session set NLS_NUMERIC_CHARACTERS='.,'");
  try
    {
      qObnovlenie->ExecSQL();
    }
  catch(...)
    {
      Application->MessageBox("Ошибка зой данных","Соединение с сервером",
                              MB_OK+MB_ICONERROR);
      Application->Terminate();
      Abort();
    }

  DM->qKomandirovki->Active = true;
  DM->qSP_chel->Active = true;
  DM->qSP_country->Active = true;
  DM->qSP_grade->Active = true;
  DM->qSP_city->Active = true;
  DM->qSP_gostinica->Active = true;
  DM->qSP_obekt->Active = true;

}
//---------------------------------------------------------------------------

void __fastcall TDM::DataModuleDestroy(TObject *Sender)
{

  DM->qObnovlenie1->Active = false;
  DM->qKomandirovki->Active = false;
  DM->qObnovlenie->Active = false;
  DM->qSP_chel->Active = false;
  DM->qSP_grade->Active = false;
  DM->qSP_country->Active = false;
  DM->qSP_city->Active = false;
  DM->qSP_gostinica->Active = false;
  DM->qSP_obekt->Active = false;

  ADOConnection1->Close();
  ADOConnection2->Close();

  ReleaseMutex(mutex1);

}
//---------------------------------------------------------------------------

