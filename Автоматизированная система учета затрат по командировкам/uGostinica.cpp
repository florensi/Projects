//---------------------------------------------------------------------------

#include <vcl.h>
#pragma hdrstop

#include "uGostinica.h"
#include "uDM.h"
#include "uMain.h"
//---------------------------------------------------------------------------
#pragma package(smart_init)
#pragma resource "*.dfm"
TGostinica *Gostinica;
//---------------------------------------------------------------------------
__fastcall TGostinica::TGostinica(TComponent* Owner)
        : TForm(Owner)
{
}
//---------------------------------------------------------------------------
void __fastcall TGostinica::CanselClick(TObject *Sender)
{
  Close();
}
//---------------------------------------------------------------------------
void __fastcall TGostinica::FormKeyDown(TObject *Sender, WORD &Key,
      TShiftState Shift)
{
  if (Key==VK_RETURN)
  FindNextControl((TWinControl *)Sender, true, true, false)->SetFocus();
}
//---------------------------------------------------------------------------
void __fastcall TGostinica::BitBtn1Click(TObject *Sender)
{
  AnsiString Sql;
  int rec;

  //��������

  //�������
  if (EditCOMFORT->Text.IsEmpty())
    {
      Application->MessageBox("�� ������� ������ � ����� '�������'","��������������",
                              MB_ICONINFORMATION+MB_OK);
      EditCOMFORT->SetFocus();
      Abort();
    }

  //�������
  if (EditCLEAR->Text.IsEmpty())
    {
      Application->MessageBox("�� ������� ������ � ����� '�������'","��������������",
                              MB_ICONINFORMATION+MB_OK);
      EditCLEAR->SetFocus();
      Abort();
    }

  //��������
  if (EditPERSONAL->Text.IsEmpty())
    {
      Application->MessageBox("�� ������� ������ � ����� '��������'","��������������",
                              MB_ICONINFORMATION+MB_OK);
      EditPERSONAL->SetFocus();
      Abort();
    }

  //�������
  if (EditPITANIE->Text.IsEmpty())
    {
      Application->MessageBox("�� ������� ������ � ����� '�������'","��������������",
                              MB_ICONINFORMATION+MB_OK);
      EditPITANIE->SetFocus();
      Abort();
    }

  //������
  if (EditSERVIS->Text.IsEmpty())
    {
      Application->MessageBox("�� ������� ������ � ����� '������'","��������������",
                              MB_ICONINFORMATION+MB_OK);
      EditSERVIS->SetFocus();
      Abort();
    }

  //������
  if (EditUSLUGI->Text.IsEmpty())
    {
      Application->MessageBox("�� ������� ������ � ����� '������'","��������������",
                              MB_ICONINFORMATION+MB_OK);
      EditUSLUGI->SetFocus();
      Abort();
    }

  //������������
  if (EditRASPOLOG->Text.IsEmpty())
    {
      Application->MessageBox("�� ������� ������ � ����� '������������'","��������������",
                              MB_ICONINFORMATION+MB_OK);
      EditRASPOLOG->SetFocus();
      Abort();
    }

  //�����������
  if (EditVPECHAT->Text.IsEmpty())
    {
      Application->MessageBox("�� ������� ������ � ����� '�����������'","��������������",
                              MB_ICONINFORMATION+MB_OK);
      EditVPECHAT->SetFocus();
      Abort();
    }

  //�����������
  if (EditORGANIZ->Text.IsEmpty())
    {
      Application->MessageBox("�� ������� ������ � ����� '����������� ������������'","��������������",
                              MB_ICONINFORMATION+MB_OK);
      EditORGANIZ->SetFocus();
      Abort();
    }

  //�������������� ������
  Sql="update komandirovki set     \
                                COMFORT="+EditCOMFORT->Text+",                 \
                                CLEAR="+EditCLEAR->Text+",                     \
                                PERSONAL="+EditPERSONAL->Text+",               \
                                PITANIE="+EditPITANIE->Text+",                 \
                                SERVIS="+EditSERVIS->Text+",                   \
                                USLUGI="+EditUSLUGI->Text+",                   \
                                RASPOLOG="+EditRASPOLOG->Text+",               \
                                VPECHAT="+EditVPECHAT->Text+",                 \
                                ORGANIZ="+EditORGANIZ->Text+"                  \
         where rowid=chartorowid("+QuotedStr(DM->qKomandirovki->FieldByName("rw")->AsString)+")";

  rec = DM->qKomandirovki->RecNo;

  DM->qObnovlenie->Close();
  DM->qObnovlenie->SQL->Clear();
  DM->qObnovlenie->SQL->Add(Sql);
  try
    {
      DM->qObnovlenie->ExecSQL();
    }
  catch(Exception &E)
    {
      Application->MessageBox(("���������� ��������/�������� ������ � ��������� � ������� �� ������������� (KOMANDIROVKI) "+ E.Message).c_str(),"������",
                              MB_ICONINFORMATION+MB_OK);
      Abort();
    }


  //���������� �������� � ����������� ��������
  Sql="update sp_gostinica set reit= (select round((sum(nvl(comfort,0))/count(*)+  \
                                             sum(nvl(clear,0))/count(*)+           \
                                             sum(nvl(personal,0))/count(*)+        \
                                             sum(nvl(pitanie,0))/count(*)+         \
                                             sum(nvl(servis,0))/count(*)+          \
                                             sum(nvl(uslugi,0))/count(*)+          \
                                             sum(nvl(raspolog,0))/count(*)+        \
                                             sum(nvl(vpechat,0))/count(*)          \
                                             )/40*100,2) as reit                   \
                                      from komandirovki                            \
                                      where gostinica="+DM->qKomandirovki->FieldByName("kod_gostinica")->AsString+"    \
                                      group by gostinica)                          \
      where kod="+DM->qKomandirovki->FieldByName("kod_gostinica")->AsString;

  DM->qObnovlenie->Close();
  DM->qObnovlenie->SQL->Clear();
  DM->qObnovlenie->SQL->Add(Sql);
  try
    {
      DM->qObnovlenie->ExecSQL();
    }
  catch(Exception &E)
    {
      Application->MessageBox(("���������� ��������/�������� ������ � �������� ��������� � ���������� (SP_GOSTINICA) "+ E.Message).c_str(),"������",
                              MB_ICONINFORMATION+MB_OK);
      Main->InsertLog("�� ��������� ���������� �������� � ����������� ��������: ����� "+DM->qKomandirovki->FieldByName("gorod")->AsString+", ��������� "+DM->qKomandirovki->FieldByName("gostinica")->AsString);
      Abort();
    }

  DM->qKomandirovki->Requery();
  DM->qSP_gostinica->Requery();

  //����
  Main->InsertLog("��������� ���������� �������� � ����������� ��������: ����� "+DM->qKomandirovki->FieldByName("gorod")->AsString+", ��������� "+DM->qKomandirovki->FieldByName("gostinica")->AsString);

  //����������� ������� �� ������
  DM->qKomandirovki->RecNo = rec;

  Gostinica->Close();
}
//---------------------------------------------------------------------------

void __fastcall TGostinica::EditCOMFORTExit(TObject *Sender)
{
  if (ActiveControl == Cansel)
    {
      Gostinica->Close();
    }
  else
    {

      if (!EditCOMFORT->Text.IsEmpty() && StrToInt(EditCOMFORT->Text)>5)
        {
          Application->MessageBox("������ ����� ���� � �������� �� 1 �� 5","��������������",
                                   MB_OK+MB_ICONINFORMATION);
          EditCOMFORT->SetFocus();
          Abort();
        }
    }
}
//---------------------------------------------------------------------------

void __fastcall TGostinica::EditCOMFORTKeyPress(TObject *Sender, char &Key)
{
  if (! (IsNumeric(Key) || Key=='\b') ) Key=0;
}
//---------------------------------------------------------------------------

void __fastcall TGostinica::EditCLEARExit(TObject *Sender)
{
  if (ActiveControl == Cansel)
    {
      Gostinica->Close();
    }
  else
    {
      if (!EditCOMFORT->Text.IsEmpty() && StrToInt(EditCLEAR->Text)>5)
        {
          Application->MessageBox("������ ����� ���� � �������� �� 1 �� 5","��������������",
                                   MB_OK+MB_ICONINFORMATION);
          EditCLEAR->SetFocus();
          Abort();
        }
    }
}
//---------------------------------------------------------------------------

void __fastcall TGostinica::EditPERSONALExit(TObject *Sender)
{
  if (ActiveControl == Cansel)
    {
      Gostinica->Close();
    }
  else
    {
      if (!EditCOMFORT->Text.IsEmpty() && StrToInt(EditPERSONAL->Text)>5)
        {
          Application->MessageBox("������ ����� ���� � �������� �� 1 �� 5","��������������",
                                   MB_OK+MB_ICONINFORMATION);
          EditPERSONAL->SetFocus();
          Abort();
        }
    }
}
//---------------------------------------------------------------------------

void __fastcall TGostinica::EditPITANIEExit(TObject *Sender)
{
  if (ActiveControl == Cansel)
    {
      Gostinica->Close();
    }
  else
    {
      if (!EditCOMFORT->Text.IsEmpty() && StrToInt(EditPITANIE->Text)>5)
        {
          Application->MessageBox("������ ����� ���� � �������� �� 1 �� 5","��������������",
                                   MB_OK+MB_ICONINFORMATION);
          EditPITANIE->SetFocus();
          Abort();
        }
    }
}
//---------------------------------------------------------------------------

void __fastcall TGostinica::EditSERVISExit(TObject *Sender)
{
  if (ActiveControl == Cansel)
    {
      Gostinica->Close();
    }
  else
    {
      if (!EditCOMFORT->Text.IsEmpty() && StrToInt(EditSERVIS->Text)>5)
        {
          Application->MessageBox("������ ����� ���� � �������� �� 1 �� 5","��������������",
                                   MB_OK+MB_ICONINFORMATION);
          EditSERVIS->SetFocus();
          Abort();
        }
    }
}
//---------------------------------------------------------------------------

void __fastcall TGostinica::EditUSLUGIExit(TObject *Sender)
{
  if (ActiveControl == Cansel)
    {
      Gostinica->Close();
    }
  else
    {
      if (!EditCOMFORT->Text.IsEmpty() && StrToInt(EditUSLUGI->Text)>5)
        {
          Application->MessageBox("������ ����� ���� � �������� �� 1 �� 5","��������������",
                                   MB_OK+MB_ICONINFORMATION);
          EditUSLUGI->SetFocus();
          Abort();
        }
    }
}
//---------------------------------------------------------------------------

void __fastcall TGostinica::EditRASPOLOGExit(TObject *Sender)
{
  if (ActiveControl == Cansel)
    {
      Gostinica->Close();
    }
  else
    {
      if (!EditCOMFORT->Text.IsEmpty() && StrToInt(EditRASPOLOG->Text)>5)
        {
          Application->MessageBox("������ ����� ���� � �������� �� 1 �� 5","��������������",
                                   MB_OK+MB_ICONINFORMATION);
          EditRASPOLOG->SetFocus();
          Abort();
        }
    }
}
//---------------------------------------------------------------------------

void __fastcall TGostinica::EditVPECHATExit(TObject *Sender)
{
  if (ActiveControl == Cansel)
    {
      Gostinica->Close();
    }
  else
    {
      if (!EditCOMFORT->Text.IsEmpty() && StrToInt(EditVPECHAT->Text)>5)
        {
          Application->MessageBox("������ ����� ���� � �������� �� 1 �� 5","��������������",
                                   MB_OK+MB_ICONINFORMATION);
          EditVPECHAT->SetFocus();
          Abort();
        }
    }
}
//---------------------------------------------------------------------------

void __fastcall TGostinica::EditORGANIZExit(TObject *Sender)
{
  if (ActiveControl == Cansel)
    {
      Gostinica->Close();
    }
  else
    {
      if (!EditCOMFORT->Text.IsEmpty() && StrToInt(EditORGANIZ->Text)>5)
        {
          Application->MessageBox("������ ����� ���� � �������� �� 1 �� 5","��������������",
                                   MB_OK+MB_ICONINFORMATION);
          EditORGANIZ->SetFocus();
          Abort();
        }
    }    
}
//---------------------------------------------------------------------------

void __fastcall TGostinica::FormShow(TObject *Sender)
{
  EditCOMFORT->SetFocus();        
}
//---------------------------------------------------------------------------

