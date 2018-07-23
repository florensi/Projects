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

   //���������� ������ ���������� �� �������������� �����
   try
     {
       DecryptFromFile(GetCurrentDir() + "\\certificate.1.13.o.dat", S);
     }
   catch(Exception &E)
    {
      Application->MessageBox(("���������� �������� ������ ���������� � ����� ������:\n" + E.Message).c_str(),"������ ����������",
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
      Application->MessageBox(("������ ���������� � ����� ������:\n" + E.Message).c_str(),"������ ����������",
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

  //��������� ����������� ��� �������� ����� '.' ��� ������� ������ Oracle
  qObnovlenie->Close();
  qObnovlenie->SQL->Clear();
  qObnovlenie->SQL->Add("alter session set NLS_NUMERIC_CHARACTERS='.,'");
  try
    {
      qObnovlenie->ExecSQL();
    }
  catch(Exception &E)
    {
      Application->MessageBox(("���������� ���������� ����������� '.'\n��� �������� �����  ��� ������� ������ Oracle:\n" + E.Message).c_str(),"���������� � ��������",
                              MB_OK + MB_ICONERROR);
      Application->Terminate();
      Abort();
    }

}
//---------------------------------------------------------------------------

//���������� ������� ������ ��� ���������
void __fastcall TDM::dsZameshDataChange(TObject *Sender, TField *Field)
{
  Zameshenie->ZapolnenieInfo();
}
//---------------------------------------------------------------------------

