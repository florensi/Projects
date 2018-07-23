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

  //�������� �� ���������� �� ���������� ����������
  try
    {
      mutex1 = CreateMutex(NULL,false,"Komandirovki");
      int Res = GetLastError();
      if (Res == ERROR_ALREADY_EXISTS)
        {
          Application->MessageBox("������ ���������� ��� �������� �� ���� ����������!!!","��������������",
                                  MB_OK+MB_ICONWARNING);
          Application->Terminate();
          Abort();
        }
      if (Res == ERROR_INVALID_HANDLE)
        {
          Application->MessageBox("������ � ����� mutex","��������������",
                                  MB_OK+MB_ICONWARNING);
          Application->Terminate();
          Abort();
        }
    }
  catch(...)
    {
      Application->MessageBox("�������� ������ ��� �������� �� ���������� ������ ��������� ���������!!!","��������������",
                                  MB_OK+MB_ICONWARNING);
      Application->Terminate();
      Abort();
    }

  //���������� ������ ���������� �� �������������� �����
  if (!DecryptFromFile(GetCurrentDir() + "\\certificate.1.1.m.dat", S))
    {
      Application->MessageBox("���������� �������� ������ ���������� � ����� ������","������ ����������",
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
      Application->MessageBox("������ ���������� � ����� ������","������ ����������",
                              MB_OK + MB_ICONERROR);
      Application->Terminate();
      Abort();
    }

  //���������� ������ ���������� �� �������������� �����
  if (!DecryptFromFile(GetCurrentDir() + "\\certificate.1.13.m.dat", S))
    {
      Application->MessageBox("���������� �������� ������ ���������� � ����� ������","������ ����������",
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
      Application->MessageBox("������ ���������� � ����� ������","������ ����������",
                              MB_OK + MB_ICONERROR);
      Application->Terminate();
      Abort();
    }


  Application->UpdateFormatSettings = false;
  DateSeparator = '.';
  DecimalSeparator = '.';

  ShortDateFormat = "dd.mm.yyyy";

  //��������� ����������� ��� �������� ����� '.' ��� ������� ������ Oracle
  qObnovlenie->Close();
  qObnovlenie->SQL->Clear();
  qObnovlenie->SQL->Add("alter session set NLS_NUMERIC_CHARACTERS='.,'");
  try
    {
      qObnovlenie->ExecSQL();
    }
  catch(...)
    {
      Application->MessageBox("������ ��� ������","���������� � ��������",
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

