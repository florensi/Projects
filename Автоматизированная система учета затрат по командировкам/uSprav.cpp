//---------------------------------------------------------------------------

#include <vcl.h>
#pragma hdrstop

#include "uSprav.h"
#include "uDM.h"
#include "uVvod.h"
#include "uMain.h"
//---------------------------------------------------------------------------
#pragma package(smart_init)
#pragma link "DBGridEh"
#pragma resource "*.dfm"
TSprav *Sprav;
//---------------------------------------------------------------------------
__fastcall TSprav::TSprav(TComponent* Owner)
        : TForm(Owner)
{
}
//---------------------------------------------------------------------------
void __fastcall TSprav::Ljfdbnmpfgbcm1Click(TObject *Sender)
{
  Panel1->Visible=true;
  fl_sp_red=0;

  EditCHEL->Text="";
  Vvod->Label14->Caption="���������� ������";
  DBGridEh1->Top=139;
  DBGridEh1->Height=323;

  EditCHEL->SetFocus();
}
//---------------------------------------------------------------------------

void __fastcall TSprav::N1RedaktClick(TObject *Sender)
{
  Panel1->Visible=true;
  fl_sp_red=1;

  EditCHEL->Text=DM->qSP_chel->FieldByName("naim")->AsString;
  Vvod->Label14->Caption="�������������� ������";
  DBGridEh1->Top=139;
  DBGridEh1->Height=323;

  EditCHEL->SetFocus();
}
//---------------------------------------------------------------------------


void __fastcall TSprav::N4Click(TObject *Sender)
{
  Panel4->Visible=true;

  fl_sp_red=0;

  EditGRADE->Text="";
  EditKAT->Text="";
  EditVAGON->Text="";
  EditG_MIN_KIEV->Text="";
  EditG_KIEV->Text="";
  EditG_MIN_UKR->Text="";
  EditG_UKR->Text="";
  EditG_ZAGRAN->Text="";
  Label2->Caption="���������� ������";
  DBGridEh2->Top=116;
  DBGridEh2->Height=350;

  EditGRADE->SetFocus();
 }
//---------------------------------------------------------------------------

void __fastcall TSprav::N5Click(TObject *Sender)
{
  Panel4->Visible=true;
  fl_sp_red=1;


  EditGRADE->Text=DM->qSP_grade->FieldByName("grade")->AsString;
  EditKAT->Text=DM->qSP_grade->FieldByName("kat")->AsString;
  EditVAGON->Text=DM->qSP_grade->FieldByName("vagon")->AsString;
  EditG_MIN_KIEV->Text=DM->qSP_grade->FieldByName("g_min_kiev")->AsString;
  EditG_KIEV->Text=DM->qSP_grade->FieldByName("g_kiev")->AsString;
  EditG_MIN_UKR->Text=DM->qSP_grade->FieldByName("g_min_ukr")->AsString;
  EditG_UKR->Text=DM->qSP_grade->FieldByName("g_ukr")->AsString;
  EditG_ZAGRAN->Text=DM->qSP_grade->FieldByName("g_zagran")->AsString;


  DBGridEh2->Top=143;
  DBGridEh2->Height=458;

  EditGRADE->SetFocus();
}
//---------------------------------------------------------------------------

void __fastcall TSprav::N8Click(TObject *Sender)
{
  Panel7->Visible=true;
  fl_sp_red=0;

  EditGOROD->Text="";
  EditGOSTINICA->Text="";
  EditGOST_ADR->Text="";

  DBGridEh3->Top=134;
  DBGridEh3->Height=335;

  EditGOROD->SetFocus();
}
//---------------------------------------------------------------------------

void __fastcall TSprav::N9Click(TObject *Sender)
{
  Panel7->Visible=true;
  fl_sp_red=1;

  EditGOROD->Text=DM->qSP_gostinica->FieldByName("city")->AsString;
  EditGOSTINICA->Text=DM->qSP_gostinica->FieldByName("gostinica")->AsString;
  EditGOST_ADR->Text=DM->qSP_gostinica->FieldByName("adress")->AsString;

  DBGridEh3->Top=134;
  DBGridEh3->Height=335;

  EditGOROD->SetFocus();
}
//---------------------------------------------------------------------------

void __fastcall TSprav::N12Click(TObject *Sender)
{
  Panel10->Visible=true;
  fl_sp_red=0;

  EditGOROD1->Text="";
  EditOBEKT->Text="";
  EditADRESS->Text="";

  DBGridEh4->Top=131;
  DBGridEh4->Height=344;

  EditGOROD1->SetFocus();
}
//---------------------------------------------------------------------------

void __fastcall TSprav::N13Click(TObject *Sender)
{
  Panel10->Visible=true;
  fl_sp_red=1;


  EditGOROD1->Text=DM->qSP_obekt->FieldByName("city")->AsString;
  EditOBEKT->Text=DM->qSP_obekt->FieldByName("obekt")->AsString;
  EditADRESS->Text=DM->qSP_obekt->FieldByName("adress")->AsString;

  DBGridEh4->Top=131;
  DBGridEh4->Height=344;

  EditGOROD1->SetFocus();
}
//---------------------------------------------------------------------------

void __fastcall TSprav::DBGridEh2DrawColumnCell(TObject *Sender,
      const TRect &Rect, int DataCol, TColumnEh *Column,
      TGridDrawState State)
{
  // ��������� ������ �������� ������
  if (State.Contains(gdSelected) )
    {
      ((TDBGridEh *) Sender)->Canvas->Brush->Color = clSkyBlue;//(TColor)0x00DEF5F4;//clInfoBk;
    }
  ((TDBGridEh *) Sender)->Canvas->Font->Color = clBlack;
  ((TDBGridEh *) Sender)->DefaultDrawColumnCell(Rect, DataCol, Column, State);
}
//---------------------------------------------------------------------------

void __fastcall TSprav::FormShow(TObject *Sender)
{
  Sprav->TabSheet1->Caption = "���� ������������";
  Sprav->TabSheet2->Caption = "������";
  Sprav->TabSheet3->Caption = "���������";
  Sprav->TabSheet4->Caption = "�������";
  Sprav->TabSheet5->Caption = "������";
  Sprav->TabSheet6->Caption = "������";

  Panel16->Visible=false;
  Panel1->Visible=false;
  Panel4->Visible=false;
 // DBGridEh2->Height=457;
 // DBGridEh2->Top=143;
  Panel7->Visible=false;
  Panel10->Visible=false;
  Panel13->Visible=false;
  

  PageControl1->OwnerDraw = true;
}
//---------------------------------------------------------------------------

void __fastcall TSprav::BitBtn2Click(TObject *Sender)
{
  Panel4->Visible=false;
  DBGridEh2->Top=143;
  DBGridEh2->Height=458;
}
//---------------------------------------------------------------------------

void __fastcall TSprav::FormKeyDown(TObject *Sender, WORD &Key,
      TShiftState Shift)
{
  if (Key==VK_RETURN)
  FindNextControl((TWinControl *)Sender, true, true, false)->SetFocus();
}
//---------------------------------------------------------------------------

//���������� ������
void __fastcall TSprav::BitBtn1Click(TObject *Sender)
{
  TLocateOptions SearchOptions;
  AnsiString Sql;
  int rec;

  //�������� �� ��������� �����
  if (EditGRADE->Text.IsEmpty())
    {
      Application->MessageBox("�� ������ �����!!!","��������������",
                               MB_OK+MB_ICONINFORMATION);
      EditGRADE->SetFocus();
    }

  //�������� �� ������� ������ � ����� �������
  if (DM->qSP_grade->Locate("grade",EditGRADE->Text,SearchOptions << loCaseInsensitive))
    {
      Application->MessageBox("��������� ����� ��� ���� � �����������","��������������",
                              MB_OK+MB_ICONINFORMATION);
      EditGRADE->SetFocus();
    }

  //�������� �� �������� ���������
  if (EditKAT->Text.IsEmpty())
    {
      Application->MessageBox("�� ������� ��������� ������!!!","��������������",
                               MB_OK+MB_ICONINFORMATION);
      EditKAT->SetFocus();
    }

  //�������� �� ��������� �����
  if (EditVAGON->Text.IsEmpty())
    {
      Application->MessageBox("�� ������� ������������� ������!!!","��������������",
                               MB_OK+MB_ICONINFORMATION);
      EditVAGON->SetFocus();
    }

 //�������� �� ���� min ����� �� �����
  if (EditG_MIN_KIEV->Text.IsEmpty())
    {
      Application->MessageBox("�� ������� ����������� ����� �� �����!!!","��������������",
                               MB_OK+MB_ICONINFORMATION);
      EditG_MIN_KIEV->SetFocus();
    }

 //�������� �� ���� max ����� �� �����
  if (EditG_KIEV->Text.IsEmpty())
    {
      Application->MessageBox("�� ������� ������������ ����� �� �����!!!","��������������",
                               MB_OK+MB_ICONINFORMATION);
      EditG_KIEV->SetFocus();
    }
  //��������, ���� min ����� �� ���� ������ max �� �����
  if (StrToFloat(EditG_MIN_KIEV->Text)>StrToFloat(EditG_KIEV->Text))
    {
      Application->MessageBox("���������� ����� �� ����� ��������� ������������!!!","��������������",
                               MB_OK+MB_ICONINFORMATION);
      EditG_MIN_KIEV->SetFocus();
    }

  //�������� �� ���� min ����� �� �������
  if (EditG_MIN_UKR->Text.IsEmpty())
    {
      Application->MessageBox("�� ������� ����������� ����� �� �������!!!","��������������",
                               MB_OK+MB_ICONINFORMATION);
      EditG_MIN_UKR->SetFocus();
    }

  //�������� �� ���� max ����� �� �������
  if (EditG_UKR->Text.IsEmpty())
    {
      Application->MessageBox("�� ������� ������������ ����� �� �������!!!","��������������",
                               MB_OK+MB_ICONINFORMATION);
      EditG_UKR->SetFocus();
    }

  //��������, ���� min ����� �� ���� ������ max �� �����
   if (StrToFloat(EditG_MIN_UKR->Text)>StrToFloat(EditG_UKR->Text))
    {
      Application->MessageBox("���������� ����� �� ������� ��������� ������������!!!","��������������",
                               MB_OK+MB_ICONINFORMATION);
      EditG_MIN_UKR->SetFocus();
    }

  //�������� �� ���� ����� �� �������� �� �����
  if (EditG_ZAGRAN->Text.IsEmpty())
    {
      Application->MessageBox("�� ������� ����� ��� ���������!!!","��������������",
                               MB_OK+MB_ICONINFORMATION);
      EditG_ZAGRAN->SetFocus();
    }


  //���������� ������
  if (fl_sp_red==0)
    {
       Sql="insert into sp_grade (grade, kat, g_min_kiev, g_kiev, \
                                  g_min_ukr, g_ukr, g_zagran, vagon)\
           values ("\
                   +EditGRADE->Text+","
                   +QuotedStr(EditKAT->Text)+","
                   +EditG_MIN_KIEV->Text+","
                   +EditG_KIEV->Text+","
                   +EditG_MIN_UKR->Text+","
                   +EditG_UKR->Text+","
                   +EditG_ZAGRAN->Text+","
                   +QuotedStr(EditVAGON->Text)+")";
    }
  //���������� ������
  else if (fl_sp_red==1)
    {
      Sql="update sp_grade set\
                               grade = "+EditGRADE->Text+",  \
                               kat="+QuotedStr(EditKAT->Text)+",        \
                               g_min_kiev="+EditG_MIN_KIEV->Text+", \
                               g_kiev="+EditG_KIEV->Text+",         \
                               g_min_ukr="+EditG_MIN_UKR->Text+",   \
                               g_ukr="+EditG_UKR->Text+",           \
                               g_zagran="+EditG_ZAGRAN->Text+",     \
                               vagon="+QuotedStr(EditVAGON->Text)+"            \
           where rowid=chartorowid("+QuotedStr(DM->qSP_grade->FieldByName("rw")->AsString)+")";

      rec = DM->qSP_grade->RecNo;
    }

  DM->qObnovlenie->Close();
  DM->qObnovlenie->SQL->Clear();
  DM->qObnovlenie->SQL->Add(Sql);
  try
    {
      DM->qObnovlenie->ExecSQL();
    }
  catch(Exception &E)
    {
      Application->MessageBox(("���������� ��������/�������� ������ � ����������� ������� (SP_grade)"+E.Message).c_str(),"������",
                              MB_OK+MB_ICONERROR);
      Panel4->Visible=false;
      DBGridEh2->Top=143;
      DBGridEh2->Height=458;
      Abort();
    }

  //����

  DM->qSP_grade->Requery();

  //����������� ������� �� ������
  if (fl_sp_red==0)
    {
      // ��� ���������� ������ ���������� �� ��� ������
      DM->qSP_grade->Locate("grade",EditGRADE->Text,SearchOptions << loCaseInsensitive);
    }
  else
    {
      DM->qSP_grade->RecNo = rec;
    }


  Panel4->Visible=false;
  DBGridEh2->Top=143;
  DBGridEh2->Height=458;

}
//---------------------------------------------------------------------------

void __fastcall TSprav::BitBtn4Click(TObject *Sender)
{
  Panel1->Visible=false;
  DBGridEh1->Top=175;
  DBGridEh1->Height=425;
}
//---------------------------------------------------------------------------


void __fastcall TSprav::DBGridEh1DrawColumnCell(TObject *Sender,
      const TRect &Rect, int DataCol, TColumnEh *Column,
      TGridDrawState State)
{
  // ��������� ������ �������� ������
  if (State.Contains(gdSelected) )
    {
      ((TDBGridEh *) Sender)->Canvas->Brush->Color = clSkyBlue;//(TColor)0x00DEF5F4;//clInfoBk;
    }
  ((TDBGridEh *) Sender)->Canvas->Font->Color = clBlack;
  ((TDBGridEh *) Sender)->DefaultDrawColumnCell(Rect, DataCol, Column, State);
}
//---------------------------------------------------------------------------

void __fastcall TSprav::DBGridEh3DrawColumnCell(TObject *Sender,
      const TRect &Rect, int DataCol, TColumnEh *Column,
      TGridDrawState State)
{
  // ��������� ������ �������� ������
  if (State.Contains(gdSelected) )
    {
      ((TDBGridEh *) Sender)->Canvas->Brush->Color = clSkyBlue;//(TColor)0x00DEF5F4;//clInfoBk;
    }
  ((TDBGridEh *) Sender)->Canvas->Font->Color = clBlack;
  ((TDBGridEh *) Sender)->DefaultDrawColumnCell(Rect, DataCol, Column, State);        
}
//---------------------------------------------------------------------------

void __fastcall TSprav::DBGridEh4DrawColumnCell(TObject *Sender,
      const TRect &Rect, int DataCol, TColumnEh *Column,
      TGridDrawState State)
{
  // ��������� ������ �������� ������
  if (State.Contains(gdSelected) )
    {
      ((TDBGridEh *) Sender)->Canvas->Brush->Color = clSkyBlue;//(TColor)0x00DEF5F4;//clInfoBk;
    }
  ((TDBGridEh *) Sender)->Canvas->Font->Color = clBlack;
  ((TDBGridEh *) Sender)->DefaultDrawColumnCell(Rect, DataCol, Column, State);        
}
//---------------------------------------------------------------------------

void __fastcall TSprav::DBGridEh5DrawColumnCell(TObject *Sender,
      const TRect &Rect, int DataCol, TColumnEh *Column,
      TGridDrawState State)
{
  // ��������� ������ �������� ������
  if (State.Contains(gdSelected) )
    {
      ((TDBGridEh *) Sender)->Canvas->Brush->Color = clSkyBlue;//(TColor)0x00DEF5F4;//clInfoBk;
    }
  ((TDBGridEh *) Sender)->Canvas->Font->Color = clBlack;
  ((TDBGridEh *) Sender)->DefaultDrawColumnCell(Rect, DataCol, Column, State);        
}
//---------------------------------------------------------------------------

void __fastcall TSprav::DBGridEh6DrawColumnCell(TObject *Sender,
      const TRect &Rect, int DataCol, TColumnEh *Column,
      TGridDrawState State)
{
  // ��������� ������ �������� ������
  if (State.Contains(gdSelected) )
    {
      ((TDBGridEh *) Sender)->Canvas->Brush->Color = clSkyBlue;//(TColor)0x00DEF5F4;//clInfoBk;
    }
  ((TDBGridEh *) Sender)->Canvas->Font->Color = clBlack;
  ((TDBGridEh *) Sender)->DefaultDrawColumnCell(Rect, DataCol, Column, State);
}
//---------------------------------------------------------------------------

void __fastcall TSprav::BitBtn3Click(TObject *Sender)
{
  TLocateOptions SearchOptions;
  AnsiString Sql;
  int rec;

  //�������� �� ���������� ����
  if (EditCHEL->Text.IsEmpty())
    {
      Application->MessageBox("�� ��������� ������������ ���� ������������","��������������",
                               MB_OK+MB_ICONINFORMATION);
      EditCHEL->SetFocus();
    }

  //�������� �� ������������� ����
  if (DM->qSP_chel->Locate("naim",EditCHEL->Text,SearchOptions << loCaseInsensitive))
    {
      Application->MessageBox("��������� ���� ������������ ��� ���� � �����������","��������������",
                              MB_OK+MB_ICONINFORMATION);
      EditCHEL->SetFocus();
    }

  //���������� ������
  if (fl_sp_red==0)
    {
      Sql="insert into sp_komandir (kod,naim)\
           values (\
                    (select max(kod)+1 from sp_komandir), "
                    +QuotedStr(EditCHEL->Text)+")";
    }
  //�������������� ������
  else if (fl_sp_red==1)
    {
      Sql="update sp_komandir set\
                                   naim="+QuotedStr(EditCHEL->Text)+" \
           where kod="+DM->qSP_chel->FieldByName("kod")->AsString;

      rec = DM->qSP_chel->RecNo;
    }

  DM->qObnovlenie->Close();
  DM->qObnovlenie->SQL->Clear();
  DM->qObnovlenie->SQL->Add(Sql);
  try
    {
      DM->qObnovlenie->ExecSQL();
    }
  catch(Exception &E)
    {
      Application->MessageBox(("���������� ��������/�������� ������ � ����������� ����� ������������ (SP_chel)"+E.Message).c_str(),"������",
                              MB_OK+MB_ICONERROR);
      Panel1->Visible=false;
      DBGridEh1->Top=175;
      DBGridEh1->Height=425;
      Abort();
    }

  //����

  DM->qSP_chel->Requery();

  //����������� ������� �� ������
  if (fl_sp_red==0)
    {
      // ��� ���������� ������ ���������� �� ��� ������
      DM->qSP_chel->Locate("naim",EditCHEL->Text,SearchOptions << loCaseInsensitive);
    }
  else
    {
      DM->qSP_chel->RecNo = rec;
    }

  
  Panel1->Visible=false;
  DBGridEh1->Top=175;
  DBGridEh1->Height=425;
}
//---------------------------------------------------------------------------

void __fastcall TSprav::N3Click(TObject *Sender)
{
  if (Application->MessageBox("�� ������������� ������ ������� ��������� ������?","�������� ������",
                          MB_YESNO+MB_ICONINFORMATION)==ID_NO)
    {
      Abort();
    }

  //�������� ������
  DM->qObnovlenie->Close();
  DM->qObnovlenie->SQL->Clear();
  DM->qObnovlenie->SQL->Add("delete from sp_komandir where kod="+DM->qSP_chel->FieldByName("kod")->AsString);
  try
    {
      DM->qObnovlenie->ExecSQL();
    }
  catch(Exception &E)
    {
      Application->MessageBox(("���������� ������� ������ �� ����������� ����� ������������ (SP_chel) "+E.Message).c_str(),"������",
                              MB_OK+MB_ICONERROR);
      Main->InsertLog("�� ��������� �������� ������ �� ����������� �����: ���� "+DM->qSP_chel->FieldByName("naim")->AsString);
      Abort();
    }


  //����
  Main->InsertLog("��������� �������� ������ �� ����������� �����: ���� "+DM->qSP_chel->FieldByName("naim")->AsString);

  DM->qSP_chel->Requery();

  Application->MessageBox("������ ������� �������","�������� ������",
                          MB_OK+MB_ICONINFORMATION);

}
//---------------------------------------------------------------------------

void __fastcall TSprav::DBGridEh1DblClick(TObject *Sender)
{
  N1RedaktClick(Sender);        
}
//---------------------------------------------------------------------------

void __fastcall TSprav::EditGORODKeyPress(TObject *Sender, char &Key)
{
  if (IsNumeric(Key)) Key=0;         
}
//---------------------------------------------------------------------------

void __fastcall TSprav::BitBtn6Click(TObject *Sender)
{
  Panel7->Visible=false;
  DBGridEh3->Top=167;
  DBGridEh3->Height=435;
}
//---------------------------------------------------------------------------

//�������� ���������
void __fastcall TSprav::N11Click(TObject *Sender)
{
  if (Application->MessageBox("�� ������������� ������ ������� ��������� ������?","�������� ������",
                          MB_YESNO+MB_ICONINFORMATION)==ID_NO)
    {
      Abort();
    }

  //�������� ������
  DM->qObnovlenie->Close();
  DM->qObnovlenie->SQL->Clear();
  DM->qObnovlenie->SQL->Add("delete from sp_gostinica where rowid=chartorowid("+QuotedStr(DM->qSP_gostinica->FieldByName("rw")->AsString)+")");
  try
    {
      DM->qObnovlenie->ExecSQL();
    }
  catch(Exception &E)
    {
      Application->MessageBox(("���������� ������� ������ �� ����������� �������� (SP_GOSTINICA) "+E.Message).c_str(),"������",
                              MB_OK+MB_ICONERROR);
      Main->InsertLog("�� ��������� �������� ������ �� ����������� ��������: ��������� "+DM->qSP_gostinica->FieldByName("gostinica")->AsString);
      Abort();
    }


  //����
  Main->InsertLog("��������� �������� ������ �� ����������� ��������: ��������� "+DM->qSP_gostinica->FieldByName("gostinica")->AsString);

  DM->qSP_gostinica->Requery();

  Application->MessageBox("������ ������� �������","�������� ������",
                          MB_OK+MB_ICONINFORMATION);
}
//---------------------------------------------------------------------------

//����������/��������� ���������
void __fastcall TSprav::BitBtn5Click(TObject *Sender)
{
  TLocateOptions SearchOptions;
  AnsiString Sql;
  int rec;

  //�������� �� ��������� �����
  if (EditGOROD->Text.IsEmpty())
    {
      Application->MessageBox("�� ������ �����!!!","��������������",
                               MB_OK+MB_ICONINFORMATION);
      EditGOROD->SetFocus();
    }

  //�������� �� �������� ���������
  if (EditGOSTINICA->Text.IsEmpty())
    {
      Application->MessageBox("�� ������� ���������!!!","��������������",
                               MB_OK+MB_ICONINFORMATION);
      EditGOSTINICA->SetFocus();
    }
  //�������� �� ��� ������������ ������
  if (DM->qSP_gostinica->Locate("gostinica",EditGOSTINICA->Text,SearchOptions << loCaseInsensitive))
    {
      Application->MessageBox("��������� ��������� ��� ���� � �����������","��������������",
                              MB_OK+MB_ICONINFORMATION);
      EditGOSTINICA->SetFocus();
    }

  //�������� �� ��������� �����
  if (EditGOST_ADR->Text.IsEmpty())
    {
      Application->MessageBox("�� ����������� ���������!!!","��������������",
                               MB_OK+MB_ICONINFORMATION);
      EditGOST_ADR->SetFocus();
    }

  //���������� ������
  if (fl_sp_red==0)
    {
       Sql="insert into sp_gostinica (kod, gostinica, adress, kod_city, reit)\
           values ( \
                   (select max(kod)+1 from sp_gostinica),"                                        \
                   +QuotedStr(EditGOSTINICA->Text)+","                                                       \
                   +QuotedStr(EditGOST_ADR->Text)+",                                                           \
                   (select kod from sp_city where upper(city)=upper(trim("+QuotedStr(EditGOROD->Text)+"))),  \
                   0)";
    }
  //���������� ������
  else if (fl_sp_red==1)
    {
      Sql="update sp_gostinica set\
                                gostinica="+QuotedStr(EditGOSTINICA->Text)+",         \
                                adress="+QuotedStr(EditGOST_ADR->Text)+",                                                   \
                                kod_city=(select kod from sp_city where upper(city)=upper(trim("+QuotedStr(EditGOROD->Text)+")))   \
           where rowid=chartorowid("+QuotedStr(DM->qSP_gostinica->FieldByName("rw")->AsString)+")";

      rec = DM->qSP_gostinica->RecNo;
    }

  DM->qObnovlenie->Close();
  DM->qObnovlenie->SQL->Clear();
  DM->qObnovlenie->SQL->Add(Sql);
  try
    {
      DM->qObnovlenie->ExecSQL();
    }
  catch(Exception &E)
    {
      Application->MessageBox(("���������� ��������/�������� ������ � ����������� �������� (SP_gostinica)"+E.Message).c_str(),"������",
                              MB_OK+MB_ICONERROR);
      Panel7->Visible=false;
      DBGridEh3->Top=167;
      DBGridEh3->Height=435;
      Abort();
    }

  //����

  DM->qSP_gostinica->Requery();

  //����������� ������� �� ������
  if (fl_sp_red==0)
    {
      // ��� ���������� ������ ���������� �� ��� ������
      DM->qSP_gostinica->Locate("gostinica",EditGOSTINICA->Text,SearchOptions << loCaseInsensitive);
    }
  else
    {
      DM->qSP_gostinica->RecNo = rec;
    }


  Panel7->Visible=false;
  DBGridEh3->Top=167;
  DBGridEh3->Height=435;

}
//---------------------------------------------------------------------------

void __fastcall TSprav::BitBtn8Click(TObject *Sender)
{
  Panel10->Visible=false;
  DBGridEh4->Top=163;
  DBGridEh4->Height=435;
}
//---------------------------------------------------------------------------

void __fastcall TSprav::BitBtn10Click(TObject *Sender)
{
  Panel16->Visible=false;
  DBGridEh5->Top=166;
  DBGridEh5->Height=437;
}
//---------------------------------------------------------------------------

void __fastcall TSprav::BitBtn12Click(TObject *Sender)
{
  Panel13->Visible=false;
  DBGridEh6->Top=160;
  DBGridEh6->Height=437;
}
//---------------------------------------------------------------------------

void __fastcall TSprav::N1Click(TObject *Sender)
{
  Panel16->Visible=true;
  fl_sp_red=0;

  EditKOD->Text="";
  EditCOUNTRY->Text="";
  EditCOUNTRY_K->Text="";

  DBGridEh5->Top=132;
  DBGridEh5->Height=332;

  EditKOD->SetFocus();
}
//---------------------------------------------------------------------------

void __fastcall TSprav::N18Click(TObject *Sender)
{
  Panel13->Visible=true;
  fl_sp_red=0;

  EditCOUNTRY2->Text="";
  EditGOROD2->Text="";

  DBGridEh6->Top=132;
  DBGridEh6->Height=332;

  EditCOUNTRY2->SetFocus();
}
//---------------------------------------------------------------------------

void __fastcall TSprav::N2Click(TObject *Sender)
{
  Panel16->Visible=true;
  fl_sp_red=1;


  EditKOD->Text=DM->qSP_country->FieldByName("kod")->AsString;
  EditCOUNTRY->Text=DM->qSP_country->FieldByName("country")->AsString;
  EditCOUNTRY_K->Text=DM->qSP_country->FieldByName("country_k")->AsString;

  DBGridEh5->Top=132;
  DBGridEh5->Height=332;

  EditKOD->SetFocus();
}
//---------------------------------------------------------------------------

void __fastcall TSprav::N19Click(TObject *Sender)
{
  Panel13->Visible=true;
  fl_sp_red=1;


  EditCOUNTRY2->Text=DM->qSP_city->FieldByName("country")->AsString;
  EditGOROD2->Text=DM->qSP_city->FieldByName("city")->AsString;

  DBGridEh6->Top=128;
  DBGridEh6->Height=338;

  EditCOUNTRY2->SetFocus();
}
//---------------------------------------------------------------------------


//�������� ������
void __fastcall TSprav::N7Click(TObject *Sender)
{
  if (Application->MessageBox("�� ������������� ������ ������� ��������� ������?","�������� ������",
                          MB_YESNO+MB_ICONINFORMATION)==ID_NO)
    {
      Abort();
    }

  //�������� ������
  DM->qObnovlenie->Close();
  DM->qObnovlenie->SQL->Clear();
  DM->qObnovlenie->SQL->Add("delete from sp_grade where rowid=chartorowid("+QuotedStr(DM->qSP_grade->FieldByName("rw")->AsString)+")");
  try
    {
      DM->qObnovlenie->ExecSQL();
    }
  catch(Exception &E)
    {
      Application->MessageBox(("���������� ������� ������ �� ����������� ������� (SP_GRADE) "+E.Message).c_str(),"������",
                              MB_OK+MB_ICONERROR);
      Main->InsertLog("�� ��������� �������� ������ �� ����������� �������: ����� "+DM->qSP_grade->FieldByName("grade")->AsString);
      Abort();
    }


  //����
  Main->InsertLog("��������� �������� ������ �� ����������� �������: ����� "+DM->qSP_grade->FieldByName("grade")->AsString);

  DM->qSP_grade->Requery();

  Application->MessageBox("������ ������� �������","�������� ������",
                          MB_OK+MB_ICONINFORMATION);
}
//---------------------------------------------------------------------------


//�������� �������
void __fastcall TSprav::N15Click(TObject *Sender)
{
  if (Application->MessageBox("�� ������������� ������ ������� ��������� ������?","�������� ������",
                          MB_YESNO+MB_ICONINFORMATION)==ID_NO)
    {
      Abort();
    }

  //�������� ������
  DM->qObnovlenie->Close();
  DM->qObnovlenie->SQL->Clear();
  DM->qObnovlenie->SQL->Add("delete from sp_obekt where rowid=chartorowid("+QuotedStr(DM->qSP_obekt->FieldByName("rw")->AsString)+")");
  try
    {
      DM->qObnovlenie->ExecSQL();
    }
  catch(Exception &E)
    {
      Application->MessageBox(("���������� ������� ������ �� ����������� �������� (SP_OBEKT) "+E.Message).c_str(),"������",
                              MB_OK+MB_ICONERROR);
      Main->InsertLog("�� ��������� �������� ������ �� ����������� ��������: �������� "+DM->qSP_obekt->FieldByName("obekt")->AsString);
      Abort();
    }


  //����
  Main->InsertLog("��������� �������� ������ �� ����������� ��������: �������� "+DM->qSP_obekt->FieldByName("obekt")->AsString);

  DM->qSP_obekt->Requery();

  Application->MessageBox("������ ������� �������","�������� ������",
                          MB_OK+MB_ICONINFORMATION);
}
//---------------------------------------------------------------------------

//�������� ������
void __fastcall TSprav::N17Click(TObject *Sender)
{
  if (Application->MessageBox("�� ������������� ������ ������� ��������� ������?","�������� ������",
                          MB_YESNO+MB_ICONINFORMATION)==ID_NO)
    {
      Abort();
    }

  //�������� ������
  DM->qObnovlenie->Close();
  DM->qObnovlenie->SQL->Clear();
  DM->qObnovlenie->SQL->Add("delete from sp_country where rowid=chartorowid("+QuotedStr(DM->qSP_country->FieldByName("rw")->AsString)+")");
  try
    {
      DM->qObnovlenie->ExecSQL();
    }
  catch(Exception &E)
    {
      Application->MessageBox(("���������� ������� ������ �� ����������� ����� (SP_COUNTRY) "+E.Message).c_str(),"������",
                              MB_OK+MB_ICONERROR);
      Main->InsertLog("�� ��������� �������� ������ �� ����������� �����: ������ "+DM->qSP_country->FieldByName("country")->AsString);
      Abort();
    }


  //����
  Main->InsertLog("��������� �������� ������ �� ����������� �����: ������ "+DM->qSP_country->FieldByName("country")->AsString);

  DM->qSP_country->Requery();

  Application->MessageBox("������ ������� �������","�������� ������",
                          MB_OK+MB_ICONINFORMATION);
}
//---------------------------------------------------------------------------

//�������� ������
void __fastcall TSprav::N21Click(TObject *Sender)
{
  if (Application->MessageBox("�� ������������� ������ ������� ��������� ������?","�������� ������",
                          MB_YESNO+MB_ICONINFORMATION)==ID_NO)
    {
      Abort();
    }

  //�������� ������
  DM->qObnovlenie->Close();
  DM->qObnovlenie->SQL->Clear();
  DM->qObnovlenie->SQL->Add("delete from sp_city where rowid=chartorowid("+QuotedStr(DM->qSP_city->FieldByName("rw")->AsString)+")");
  try
    {
      DM->qObnovlenie->ExecSQL();
    }
  catch(Exception &E)
    {
      Application->MessageBox(("���������� ������� ������ �� ����������� ������� (SP_CITY) "+E.Message).c_str(),"������",
                              MB_OK+MB_ICONERROR);
      Main->InsertLog("�� ��������� �������� ������ �� ����������� �������: ����� "+DM->qSP_city->FieldByName("city")->AsString);
      Abort();
    }


  //����
  Main->InsertLog("��������� �������� ������ �� ����������� �������: ����� "+DM->qSP_city->FieldByName("city")->AsString);

  DM->qSP_city->Requery();

  Application->MessageBox("������ ������� �������","�������� ������",
                          MB_OK+MB_ICONINFORMATION);
}
//---------------------------------------------------------------------------

//���������/���������� �������
void __fastcall TSprav::BitBtn7Click(TObject *Sender)
{
  TLocateOptions SearchOptions;
  AnsiString Sql;
  int rec;

  //�������� �� ��������� �����
  if (EditGOROD1->Text.IsEmpty())
    {
      Application->MessageBox("�� ������ �����!!!","��������������",
                               MB_OK+MB_ICONINFORMATION);
      EditGOROD1->SetFocus();
    }

  //�������� �� �������� ������
  if (EditOBEKT->Text.IsEmpty())
    {
      Application->MessageBox("�� ������ ������!!!","��������������",
                               MB_OK+MB_ICONINFORMATION);
      EditOBEKT->SetFocus();
    }
  //�������� �� ��� ������������ ������
  if (DM->qSP_obekt->Locate("obekt",EditOBEKT->Text,SearchOptions << loCaseInsensitive))
    {
      Application->MessageBox("��������� ������ ��� ���� � �����������","��������������",
                              MB_OK+MB_ICONINFORMATION);
      EditOBEKT->SetFocus();
    }

  //�������� �� ��������� �����
  if (EditADRESS->Text.IsEmpty())
    {
      Application->MessageBox("�� ������ ����� �������!!!","��������������",
                               MB_OK+MB_ICONINFORMATION);
      EditADRESS->SetFocus();
    }

  //���������� ������
  if (fl_sp_red==0)
    {
       Sql="insert into sp_obekt (kod, obekt, adress, kod_city)\
           values ( \
                   (select max(kod)+1 from sp_obekt),"                                        \
                   +QuotedStr(EditOBEKT->Text)+","                                                       \
                   +QuotedStr(EditADRESS->Text)+",                                                           \
                   (select kod from sp_city where upper(city)=upper(trim("+QuotedStr(EditGOROD1->Text)+")))  \
                   )";
    }
  //���������� ������
  else if (fl_sp_red==1)
    {
      Sql="update sp_obekt set\
                                obekt="+QuotedStr(EditOBEKT->Text)+",         \
                                adress="+QuotedStr(EditADRESS->Text)+",                                                   \
                                kod_city=(select kod from sp_city where upper(city)=upper(trim("+QuotedStr(EditGOROD1->Text)+")))  \
           where rowid=chartorowid("+QuotedStr(DM->qSP_obekt->FieldByName("rw")->AsString)+")";

      rec = DM->qSP_obekt->RecNo;
    }

  DM->qObnovlenie->Close();
  DM->qObnovlenie->SQL->Clear();
  DM->qObnovlenie->SQL->Add(Sql);
  try
    {
      DM->qObnovlenie->ExecSQL();
    }
  catch(Exception &E)
    {
      Application->MessageBox(("���������� ��������/�������� ������ � ����������� �������� (SP_OBEKT)"+E.Message).c_str(),"������",
                              MB_OK+MB_ICONERROR);
      Panel10->Visible=false;
      DBGridEh4->Top=163;
      DBGridEh4->Height=435;
      Abort();
    }

  //����

  DM->qSP_obekt->Requery();

  //����������� ������� �� ������
  if (fl_sp_red==0)
    {
      // ��� ���������� ������ ���������� �� ��� ������
      DM->qSP_obekt->Locate("obekt",EditOBEKT->Text,SearchOptions << loCaseInsensitive);
    }
  else
    {
      DM->qSP_obekt->RecNo = rec;
    }


  Panel10->Visible=false;
  DBGridEh4->Top=163;
  DBGridEh4->Height=435;
}
//---------------------------------------------------------------------------

//����������/��������� ������
void __fastcall TSprav::BitBtn9Click(TObject *Sender)
{
  TLocateOptions SearchOptions;
  AnsiString Sql;
  int rec;

  //�������� �� ��������� ���
  if (EditKOD->Text.IsEmpty())
    {
      Application->MessageBox("�� ������ ��� ������!!!","��������������",
                               MB_OK+MB_ICONINFORMATION);
      EditKOD->SetFocus();
    }

  //�������� �� ��������� ������
  if (EditCOUNTRY->Text.IsEmpty())
    {
      Application->MessageBox("�� ������� ������!!!","��������������",
                               MB_OK+MB_ICONINFORMATION);
      EditCOUNTRY->SetFocus();
    }

  //�������� �� ��� ������������ ������
  if (DM->qSP_country->Locate("country",EditCOUNTRY->Text,SearchOptions << loCaseInsensitive))
    {
      Application->MessageBox("��������� ������ ��� ���� � �����������","��������������",
                              MB_OK+MB_ICONINFORMATION);
      EditCOUNTRY->SetFocus();
    }

  //�������� �� ��������� ����������� �������� ������
  if (EditCOUNTRY_K->Text.IsEmpty())
    {
      Application->MessageBox("�� ������� ������� �������� ������!!!","��������������",
                               MB_OK+MB_ICONINFORMATION);
      EditCOUNTRY_K->SetFocus();
    }

  //���������� ������
  if (fl_sp_red==0)
    {
      //�������� �� ��� ������������ ���
      if (DM->qSP_country->Locate("kod",EditKOD->Text,SearchOptions << loCaseInsensitive))
        {
          Application->MessageBox("��������� ��� ������ ��� ���� � �����������","��������������",
                                   MB_OK+MB_ICONINFORMATION);
          EditKOD->SetFocus();
        }


       Sql="insert into sp_country (kod, country, country_k)\
           values ("\
                   +QuotedStr(EditKOD->Text)+","                                        \
                   +QuotedStr(EditCOUNTRY->Text)+","                                    \
                   +QuotedStr(EditCOUNTRY_K->Text)+")";
    }
  //���������� ������
  else if (fl_sp_red==1)
    {
      Sql="update sp_country set\                                            \
                               kod="+QuotedStr(EditKOD->Text)+",                        \
                               country="+QuotedStr(EditCOUNTRY->Text)+",                \
                               country_k="+QuotedStr(EditCOUNTRY_K->Text)+"             \
           where rowid=chartorowid("+QuotedStr(DM->qSP_country->FieldByName("rw")->AsString)+")";

      rec = DM->qSP_country->RecNo;
    }

  DM->qObnovlenie->Close();
  DM->qObnovlenie->SQL->Clear();
  DM->qObnovlenie->SQL->Add(Sql);
  try
    {
      DM->qObnovlenie->ExecSQL();
    }
  catch(Exception &E)
    {
      Application->MessageBox(("���������� ��������/�������� ������ � ����������� ����� (SP_COUNTRY)"+E.Message).c_str(),"������",
                              MB_OK+MB_ICONERROR);
      Panel16->Visible=false;
      DBGridEh5->Top=166;
      DBGridEh5->Height=437;
      Abort();
    }

  //����

  DM->qSP_country->Requery();

  //����������� ������� �� ������
  if (fl_sp_red==0)
    {
      // ��� ���������� ������ ���������� �� ��� ������
      DM->qSP_country->Locate("country",EditCOUNTRY->Text,SearchOptions << loCaseInsensitive);
    }
  else
    {
      DM->qSP_country->RecNo = rec;
    }


  Panel16->Visible=false;
  DBGridEh5->Top=166;
  DBGridEh5->Height=437;
}
//---------------------------------------------------------------------------


void __fastcall TSprav::BitBtn11Click(TObject *Sender)
{
  TLocateOptions SearchOptions;
  AnsiString Sql;
  int rec;


  //�������� �� ��������� ������
  if (EditCOUNTRY2->Text.IsEmpty())
    {
      Application->MessageBox("�� ������� ������!!!","��������������",
                               MB_OK+MB_ICONINFORMATION);
      EditCOUNTRY2->SetFocus();
    }

  //�������� �� ��������� �����
  if (EditGOROD2->Text.IsEmpty())
    {
      Application->MessageBox("�� ������ �����!!!","��������������",
                               MB_OK+MB_ICONINFORMATION);
      EditGOROD2->SetFocus();
    }

  //�������� �� ��� ������������ ������
  if (DM->qSP_city->Locate("city",EditGOROD2->Text,SearchOptions << loCaseInsensitive))
    {
      Application->MessageBox("��������� ����� ��� ���� � �����������","��������������",
                              MB_OK+MB_ICONINFORMATION);
      EditGOROD2->SetFocus();
    }


  //���������� ������
  if (fl_sp_red==0)
    {
       Sql="insert into sp_city (kod, kod_country, city)\
           values ( \
                   (select max(kod)+1 from sp_city),                                        \
                   (select kod from sp_country where upper(country)=upper(trim("+QuotedStr(EditCOUNTRY2->Text)+"))),"  \
                   +QuotedStr(EditGOROD2->Text)+")";
    }
  //���������� ������
  else if (fl_sp_red==1)
    {
      Sql="update sp_city set\
                               kod_country=(select kod from sp_country where upper(country)=upper(trim("+QuotedStr(EditCOUNTRY2->Text)+"))),  \
                               city="+QuotedStr(EditGOROD2->Text)+"                                                   \
           where rowid=chartorowid("+QuotedStr(DM->qSP_city->FieldByName("rw")->AsString)+")";

      rec = DM->qSP_city->RecNo;
    }

  DM->qObnovlenie->Close();
  DM->qObnovlenie->SQL->Clear();
  DM->qObnovlenie->SQL->Add(Sql);
  try
    {
      DM->qObnovlenie->ExecSQL();
    }
  catch(Exception &E)
    {
      Application->MessageBox(("���������� ��������/�������� ������ � ����������� ������� (SP_CITY)"+E.Message).c_str(),"������",
                              MB_OK+MB_ICONERROR);
      Panel13->Visible=false;
      DBGridEh6->Top=160;
      DBGridEh6->Height=437;
      Abort();
    }

  //����

  DM->qSP_city->Requery();

  //����������� ������� �� ������
  if (fl_sp_red==0)
    {
      // ��� ���������� ������ ���������� �� ��� ������
      DM->qSP_city->Locate("city",EditGOROD2->Text,SearchOptions << loCaseInsensitive);
    }
  else
    {
      DM->qSP_city->RecNo = rec;
    }


  Panel13->Visible=false;
  DBGridEh6->Top=160;
  DBGridEh6->Height=437;
}
//---------------------------------------------------------------------------

void __fastcall TSprav::PageControl1DrawTab(TCustomTabControl *Control,
      int TabIndex, const TRect &Rect, bool Active)
{
  AnsiString S;
  int x, y;

  S = PageControl1->Pages[TabIndex]->Caption;
  Control->Canvas->FillRect(Rect);

  if (Active)
    {
      Control->Canvas->Brush->Color = (TColor)0x00DEF5F4;
      Control->Canvas->Font->Color = clBlack;
      Control->Canvas->FillRect(Rect);
    }
  else
    {
      Control->Canvas->Font->Color = clBlack;
    }

  x = CenterPoint(Rect).x - div(Control->Canvas->TextWidth(S),2).quot;
  y = CenterPoint(Rect).y - div(Control->Canvas->TextHeight(S),2).quot;
  Control->Canvas->TextOut(x,y,S);               
}
//---------------------------------------------------------------------------


