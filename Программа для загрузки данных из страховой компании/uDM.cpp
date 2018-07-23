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

  //���������� ������ ���������� �� �������������� �����
  try
     {
       DecryptFromFile(GetCurrentDir() + "\\certificate.1.13.m.dat", S);
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


  Application->UpdateFormatSettings = false;
  DateSeparator = '.';
  DecimalSeparator = '.';


//�������� �����

  // ������� ����
  DecodeDate(Date(), year, month, day);
  mm1 = month;
  yyyy1 = year;

  //������� �����
  mm = mm1;
  yyyy = yyyy1;

  PrevMonth(mm1, yyyy1);

  // ���������� �����
  mm2 = mm1;
  yyyy2 = yyyy1;


  //��������� ����������� ��� �������� ����� ',' ��� ������� ������ Oracle
  qObnovlenie->Close();
  qObnovlenie->SQL->Clear();
  qObnovlenie->SQL->Add("alter session set NLS_NUMERIC_CHARACTERS=',.'");
  try
    {
      qObnovlenie->ExecSQL();
    }
  catch(Exception &E)
    {
      Application->MessageBox(("���������� ���������� ����������� '.'\n��� �������� �����  ��� ������� ������ Oracle:\n" + E.Message).c_str(),"���������� � ��������",
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
// ���������� �����
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

