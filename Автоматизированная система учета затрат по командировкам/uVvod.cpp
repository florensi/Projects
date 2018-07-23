//---------------------------------------------------------------------------

#include <vcl.h>
#pragma hdrstop

#include "uVvod.h"
#include "uDM.h"
#include "uGostinica.h"
#include "uMain.h"
#include "uSprav.h"
//---------------------------------------------------------------------------
#pragma package(smart_init)
#pragma link "PropStorageEh"
#pragma resource "*.dfm"
TVvod *Vvod;
//---------------------------------------------------------------------------
__fastcall TVvod::TVvod(TComponent* Owner)
        : TForm(Owner)
{
}
//---------------------------------------------------------------------------
void __fastcall TVvod::ButtonGOSTINICAClick(TObject *Sender)
{
  Main->N5OBRAT_SVClick(Sender);
}
//---------------------------------------------------------------------------

void __fastcall TVvod::CanselClick(TObject *Sender)
{
  Close();        
}
//---------------------------------------------------------------------------

void __fastcall TVvod::FormKeyDown(TObject *Sender, WORD &Key,
      TShiftState Shift)
{
  if (Key==VK_RETURN)
  FindNextControl((TWinControl *)Sender, true, true, false)->SetFocus();
}
//---------------------------------------------------------------------------

void __fastcall TVvod::FormShow(TObject *Sender)
{
  EditZEX->SetFocus();
}
//---------------------------------------------------------------------------

void __fastcall TVvod::BitBtn1Click(TObject *Sender)
{
  TLocateOptions SearchOptions;
  AnsiString Sql, tn_sap, avia, gd, bus, avto, proezd;
  int rec;

  if (CheckBoxAVIA->Checked==true) avia=1;
  else avia=NULL;

  if (CheckBoxGD->Checked==true) gd=1;
  else gd=NULL;

  if (CheckBoxBUS->Checked==true) bus=1;
  else bus=NULL;

  if (CheckBoxAVTO->Checked==true) avto=1;
  else avto=NULL;

  if (CheckBoxPROEZD->Checked==true) proezd=1;
  else proezd=NULL;


  //��������

  //���
  if (EditZEX->Text.IsEmpty())
    {
      Application->MessageBox("�� ������ ���!!!","��������������",
                              MB_ICONINFORMATION+MB_OK);
      EditZEX->SetFocus();
      Abort();
    }

  //���.�
  if (EditTN->Text.IsEmpty())
    {
      Application->MessageBox("�� ������ ��������� ����� ���������!!!","��������������",
                              MB_ICONINFORMATION+MB_OK);
      EditTN->SetFocus();
      Abort();
    }

  //��������� ���������� ������ ���
  DM->qObnovlenie1->Close();
  DM->qObnovlenie1->SQL->Clear();
  DM->qObnovlenie1->SQL->Add("select id_sap from p_work where zex="+EditZEX->Text+" and tn like '%"+EditTN->Text+"'");

  try
    {
      DM->qObnovlenie1->Open();
    }
  catch(Exception &E)
    {
      Application->MessageBox(("������ ��������� ������ �� ������� ������ (P_WORK)"+ E.Message).c_str(),"������",
                               MB_OK + MB_ICONERROR);
      Abort();
    }


  tn_sap=DM->qObnovlenie1->FieldByName("id_sap")->AsString;

  EditTN->Text = EditTN->Text.Length() ==5? EditTN->Text.SubString(2,255) : EditTN->Text;

  //�����
  if (EditGRADE->Text.IsEmpty())
    {
      Application->MessageBox("�� ������ ����� ���������!!!","��������������",
                              MB_ICONINFORMATION+MB_OK);
      EditGRADE->SetFocus();
      Abort();
    }

  //���
  if (EditFIO->Text.IsEmpty())
    {
      Application->MessageBox("�� ������� ���!!!","��������������",
                              MB_ICONINFORMATION+MB_OK);
      EditFIO->SetFocus();
      Abort();
    }

  //���������
  if (EditPROF->Text.IsEmpty())
    {
      Application->MessageBox("�� ������� ��������� ���������!!!","��������������",
                              MB_ICONINFORMATION+MB_OK);
      EditPROF->SetFocus();
      Abort();
    }

  //����
  if (ComboBoxCHEL->Text.IsEmpty())
    {
      Application->MessageBox("�� ������� ���� ������������!!!","��������������",
                              MB_ICONINFORMATION+MB_OK);
      ComboBoxCHEL->SetFocus();
      Abort();
    }

    //���� �
  if (EditDATA_N->Text.IsEmpty())
    {
      Application->MessageBox("�� ������� ���� ������ ������������!!!","��������������",
                              MB_ICONINFORMATION+MB_OK);
      EditDATA_N->SetFocus();
      Abort();
    }

    //���� ��
  if (EditDATA_K->Text.IsEmpty())
    {
      Application->MessageBox("�� ������� ���� ��������� ������������!!!","��������������",
                              MB_ICONINFORMATION+MB_OK);
      EditDATA_K->SetFocus();
      Abort();
    }

  //���� �� �� ������ ���� �
  if (StrToDate(EditDATA_K->Text)<StrToDate(EditDATA_N->Text))
    {
      Application->MessageBox("�������� ���� ������������ ������ ��������� ����!!!","��������������",
                              MB_ICONINFORMATION+MB_OK);
      EditDATA_N->SetFocus();
      Abort();
    }


  //����
  if (EditSROK->Text.IsEmpty())
    {
      Application->MessageBox("�� ������ ���� ������������!!!","��������������",
                              MB_ICONINFORMATION+MB_OK);
      EditSROK->SetFocus();
      Abort();
    }

  //���������
  if (CheckBoxAVIA->Checked==false &&
      CheckBoxGD->Checked==false &&
      CheckBoxAVTO->Checked==false &&
      CheckBoxBUS->Checked==false &&
      CheckBoxPROEZD->Checked==false)
    {
      Application->MessageBox("�� ������ ��� ����������!!!","��������������",
                              MB_ICONINFORMATION+MB_OK);
      CheckBoxAVIA->SetFocus();
      Abort();
    }

  //�����������
  if (EditNAPRAVL->Text.IsEmpty())
    {
      Application->MessageBox("�� ������� �����������!!!","��������������",
                              MB_ICONINFORMATION+MB_OK);
      EditNAPRAVL->SetFocus();
      Abort();
    }

  //������
  if (ComboBoxSTRANA->Text.IsEmpty())
    {
      Application->MessageBox("�� ������� ������!!!","��������������",
                              MB_ICONINFORMATION+MB_OK);
      ComboBoxSTRANA->SetFocus();
      Abort();
    }

  //�����
  if (ComboBoxGOROD->Text.IsEmpty())
    {
      Application->MessageBox("�� ������ �����!!!","��������������",
                              MB_ICONINFORMATION+MB_OK);
      ComboBoxGOROD->SetFocus();
      Abort();
    }

  //������
  if (ComboBoxOBEKT->Text.IsEmpty())
    {
      Application->MessageBox("�� ������ ������!!!","��������������",
                              MB_ICONINFORMATION+MB_OK);
      ComboBoxOBEKT->SetFocus();
      Abort();
    }

  //������ �������
  if (EditADRESS->Text.IsEmpty())
    {
      Application->MessageBox("�� ������ ����� �������!!!","��������������",
                              MB_ICONINFORMATION+MB_OK);
      EditADRESS->SetFocus();
      Abort();
    }

  //���� ���������� �
  if (EditDATA_GOST_N->Text.IsEmpty())
    {
      Application->MessageBox("�� ������� ��������� ���� ���������� � ���������!!!","��������������",
                              MB_ICONINFORMATION+MB_OK);
      EditDATA_GOST_N->SetFocus();
      Abort();
    }

  //���� ���������� ��
  if (EditDATA_GOST_K->Text.IsEmpty())
    {
      Application->MessageBox("�� ������� �������� ���� ���������� � ���������!!!","��������������",
                              MB_ICONINFORMATION+MB_OK);
      EditDATA_GOST_K->SetFocus();
      Abort();
    }

  //���� ���������� �� �� ������ ���� ���������� �
  if (StrToDate(EditDATA_GOST_K->Text)<StrToDate(EditDATA_GOST_N->Text))
    {
      Application->MessageBox("�������� ���� ���������� ������ ��������� ����!!!","��������������",
                              MB_ICONINFORMATION+MB_OK);
      EditDATA_GOST_N->SetFocus();
      Abort();
    }

  //���� ���������� � ������ ���� � � ������ ���� ��
  if ((StrToDate(EditDATA_GOST_N->Text)<StrToDate(EditDATA_N->Text))||
      (StrToDate(EditDATA_GOST_N->Text)>StrToDate(EditDATA_K->Text)))
    {
      Application->MessageBox("��������� ���� ���������� �� �������� � ������ ������������!!!","��������������",
                              MB_ICONINFORMATION+MB_OK);
      EditDATA_GOST_N->SetFocus();
      Abort();
    }

  //���� ���������� �� ������ ���� � � ������ ���� ��
  if ((StrToDate(EditDATA_GOST_K->Text)<StrToDate(EditDATA_N->Text))||
     (StrToDate(EditDATA_GOST_K->Text)>StrToDate(EditDATA_K->Text)))
    {
      Application->MessageBox("�������� ���� ���������� �� �������� � ������ ������������!!!","��������������",
                              MB_ICONINFORMATION+MB_OK);
      EditDATA_GOST_K->SetFocus();
      Abort();
    }

  //���������
  if (ComboBoxGOSTINICA->Text.IsEmpty())
    {
      Application->MessageBox("�� ������� ���������!!!","��������������",
                              MB_ICONINFORMATION+MB_OK);
      ComboBoxGOSTINICA->SetFocus();
      Abort();
    }

  //����� ���������
  if (EditGOST_ADR->Text.IsEmpty())
    {
      Application->MessageBox("�� ������ ����� ���������!!!","��������������",
                              MB_ICONINFORMATION+MB_OK);
      EditGOST_ADR->SetFocus();
      Abort();
    }

  //��������� ���������
  if (EditSTOIM->Text.IsEmpty())
    {
      Application->MessageBox("�� ������� ��������� ���������!!!","��������������",
                              MB_ICONINFORMATION+MB_OK);
      EditSTOIM->SetFocus();
      Abort();
    }


//��������, ���� �� ����������� ������ ������������

 Sql = "select * from komandirovki where zex="+EditZEX->Text+" and tn="+EditTN->Text+" \                                                         \
            and ((data_n < to_date(" + QuotedStr(EditDATA_N->Text) + ", 'dd.mm.yyyy') and data_k > to_date(" + QuotedStr(EditDATA_N->Text) + ", 'dd.mm.yyyy'))                                                    \
            or (data_n > to_date(" + QuotedStr(EditDATA_N->Text) + ", 'dd.mm.yyyy') and data_k > to_date(" + QuotedStr(EditDATA_N->Text) + ", 'dd.mm.yyyy') and (data_n < to_date(" + QuotedStr(EditDATA_K->Text) + ", 'dd.mm.yyyy') or data_n = to_date(" + QuotedStr(EditDATA_K->Text) + ", 'dd.mm.yyyy')))    \
            or  (data_n = to_date(" + QuotedStr(EditDATA_N->Text) + ", 'dd.mm.yyyy') or data_k = to_date(" + QuotedStr(EditDATA_N->Text) + ", 'dd.mm.yyyy')))";

  if (Main->fl_redakt==1)  Sql+= "and rowid!=chartorowid("+QuotedStr(DM->qKomandirovki->FieldByName("rw")->AsString)+")";


  DM->qObnovlenie->Close();
  DM->qObnovlenie->SQL->Clear();
  DM->qObnovlenie->SQL->Add(Sql);

  try
    {
      DM->qObnovlenie->Open();
    }
  catch(Exception &E)
    {
      Application->MessageBox(("������ ������� � ������� �� ������������� (KOMANDIROVKI)"+ E.Message).c_str(),"������",
                               MB_OK + MB_ICONERROR);
      Abort();
    }

  if (DM->qObnovlenie->RecordCount>0)
    {
      Application->MessageBox("�������� ������������ �� ������� ��������� ������������ \n� ��� ������������ � ����","������",
                               MB_OK + MB_ICONERROR);
      EditDATA_N->SetFocus();
      Abort();
    }


  //���������� ������
  if (Main->fl_redakt==0)
    {
      //�������� �� ��� ������������ � ������������
      if (!EditKOD_KOM->Text.IsEmpty())
        {
          if (DM->qKomandirovki->Locate("kod_kom",EditKOD_KOM->Text,SearchOptions << loCaseInsensitive))
            {
              Application->MessageBox("��������� ����� ������������ ��� ���������� � ��������� �� �������������","��������������",
                                       MB_OK+MB_ICONINFORMATION);
              EditKOD_KOM->SetFocus();
              Abort();
            }
        }


      Sql = "insert into komandirovki (KOD_KOM, ZEX, ZEX_NAIM, TN, TN_SAP, FIO,                       \
                                          PROF, GRADE, CHEL, DATA_N, DATA_K, DATA_ZAK,                 \
                                          SROK, AVIA, GD, BUS, AVTO, PROEZD, STRANA,                   \
                                          GOROD, OBEKT, ADRESS, NAPRAVL, DATA_GOST_N, DATA_GOST_K,             \
                                          GOSTINICA, GOST_ADR, STOIM, N_DOCUM, SUM_SUT,                          \
                                          SUM_PROGIV, SUM_TRANSP, SUM_AVIA, SUM_GD,SUM_PROCH,                    \
                                          PRIMECH)                                                     \
                values("                                                                               \
                +SetNull(EditKOD_KOM->Text)+","                                                        \
                +SetNull(EditZEX->Text)+","                                                            \
                +QuotedStr(LabelZEX->Caption)+","                                                      \
                +SetNull(EditTN->Text)+","                                                             \
                +SetNull(tn_sap)+","                                                                   \
                +QuotedStr(EditFIO->Text)+","                                                          \
                +QuotedStr(EditPROF->Text)+","                                                         \
                +SetNull(EditGRADE->Text)+",                                                           \
                (select kod from sp_komandir where naim="+QuotedStr(ComboBoxCHEL->Text)+"),"           \
                +QuotedStr(EditDATA_N->Text)+","                                                       \
                +QuotedStr(EditDATA_K->Text)+","                                                       \
                +QuotedStr(EditDATA_ZAK->Text)+","                                                     \
                +SetNull(EditSROK->Text)+","                                                           \
                +avia+","                                                                    \
                +gd+","                                                                      \
                +bus+","                                                                     \
                +avto+","                                                                    \
                +proezd+",                                                                   \
                (select kod from sp_country where country="+QuotedStr(ComboBoxSTRANA->Text)+"),        \
                (select kod from sp_city where city="+QuotedStr(ComboBoxGOROD->Text)+"),               \
                (select kod from sp_obekt where obekt="+QuotedStr(ComboBoxOBEKT->Text)+"),"            \
                +QuotedStr(EditADRESS->Text)+","                                                   \
                +QuotedStr(EditNAPRAVL->Text)+","                                                      \
                +QuotedStr(EditDATA_GOST_N->Text)+","                                                  \
                +QuotedStr(EditDATA_GOST_K->Text)+",                                                   \
                (select kod from sp_gostinica where gostinica="+QuotedStr(ComboBoxGOSTINICA->Text)+"),"   \
                +QuotedStr(EditGOST_ADR->Text)+","                                                              \
                +SetNull(EditSTOIM->Text)+","                                                                       \
                +SetNull(EditN_DOCUM->Text)+","                                                                     \
                +SetNull(EditSUM_SUT->Text)+","                                                                     \
                +SetNull(EditSUM_PROGIV->Text)+","                                                                  \
                +SetNull(EditSUM_TRANSP->Text)+","                                                                  \
                +SetNull(EditSUM_AVIA->Text)+","                                                                    \
                +SetNull(EditSUM_GD->Text)+","                                                                      \
                +SetNull(EditSUM_PROCH->Text)+","
                +QuotedStr(MemoPRIMECH->Text)+")";

    }
  //���������� ������
  else if (Main->fl_redakt==1)
    {
      //�������� �� �������� ���������


      Sql = "update komandirovki set                                                                     \
                         KOD_KOM="+SetNull(EditKOD_KOM->Text)+",                                         \
                         ZEX="+SetNull(EditZEX->Text)+",                                                 \
                         ZEX_NAIM="+QuotedStr(LabelZEX->Caption)+",                                      \
                         TN="+SetNull(EditTN->Text)+",                                                   \
                         TN_SAP="+SetNull(tn_sap)+",                                                     \
                         FIO="+QuotedStr(EditFIO->Text)+",                                               \
                         PROF="+QuotedStr(EditPROF->Text)+",                                             \
                         GRADE="+SetNull(EditGRADE->Text)+",                                             \
                         CHEL=(select kod from sp_komandir where naim="+QuotedStr(ComboBoxCHEL->Text)+"),\
                         DATA_N="+QuotedStr(EditDATA_N->Text)+",                                         \
                         DATA_K="+QuotedStr(EditDATA_K->Text)+",                                         \
                         DATA_ZAK="+QuotedStr(EditDATA_ZAK->Text)+",                                     \
                         SROK="+SetNull(EditSROK->Text)+",                                               \
                         AVIA="+avia+",                                                                  \
                         GD="+gd+",                                                                      \
                         BUS="+bus+",                                                                    \
                         AVTO="+avto+",                                                                  \
                         PROEZD="+proezd+",                                                              \
                         STRANA=(select kod from sp_country where country="+QuotedStr(ComboBoxSTRANA->Text)+"),   \
                         GOROD=(select kod from sp_city where city="+QuotedStr(ComboBoxGOROD->Text)+"),           \
                         OBEKT=(select kod from sp_obekt where obekt="+QuotedStr(ComboBoxOBEKT->Text)+"),        \
                         ADRESS="+QuotedStr(EditADRESS->Text)+",                                                  \
                         NAPRAVL="+QuotedStr(EditNAPRAVL->Text)+",                                                \
                         DATA_GOST_N="+QuotedStr(EditDATA_GOST_N->Text)+",                                        \
                         DATA_GOST_K="+QuotedStr(EditDATA_GOST_K->Text)+",                                        \
                         GOSTINICA=(select kod from sp_gostinica where gostinica="+QuotedStr(ComboBoxGOSTINICA->Text)+"), \
                         GOST_ADR="+QuotedStr(EditGOST_ADR->Text)+",                                                      \
                         STOIM="+SetNull(EditSTOIM->Text)+",                                                              \
                         N_DOCUM="+SetNull(EditN_DOCUM->Text)+",                                                          \
                         SUM_SUT="+SetNull(EditSUM_SUT->Text)+",                                                          \
                         SUM_PROGIV="+SetNull(EditSUM_PROGIV->Text)+",                                                    \
                         SUM_TRANSP="+SetNull(EditSUM_TRANSP->Text)+",                                                    \
                         SUM_AVIA="+SetNull(EditSUM_AVIA->Text)+",                                                        \
                         SUM_GD="+SetNull(EditSUM_GD->Text)+",   \
                         SUM_PROCH="+SetNull(EditSUM_PROCH->Text)+",                                                          \
                         PRIMECH="+QuotedStr(MemoPRIMECH->Text)+"                                                         \
                where rowid=chartorowid("+QuotedStr(DM->qKomandirovki->FieldByName("rw")->AsString)+")";

      rec = DM->qKomandirovki->RecNo;
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
      Application->MessageBox(("���������� ��������/�������� ������ � ������� �� ������������� (KOMANDIROVKI) "+ E.Message).c_str(),"������",
                              MB_ICONINFORMATION+MB_OK);
      Abort();
    }


  DM->qKomandirovki->Requery();

  //����

  //����������� ������� �� ������
  if (Main->fl_redakt==0)
    {
      // ��� ���������� ������ ���������� �� ��� ������
   //   DM->qKomandirovki->Locate("naim",EditCHEL->Text,SearchOptions << loCaseInsensitive);
    }
  else
    {
      DM->qKomandirovki->RecNo = rec;
    }

  Close();
}
//---------------------------------------------------------------------------

void __fastcall TVvod::FormCreate(TObject *Sender)
{
  AnsiString Sql;
  int kol;

  //���������� ComboBox �� ����� ��������������

  //���������� �����
  ComboBoxCHEL->Items->Clear();
  DM->qSP_chel->First();
  kol = DM->qSP_chel->RecordCount;
  for (int i=1; i<=kol; i++)
    {
      ComboBoxCHEL->Items->Add(DM->qSP_chel->FieldByName("naim")->AsString);
      DM->qSP_chel->Next();
    }
  DM->qSP_chel->First();


  //���������� �����
  ComboBoxSTRANA->Items->Clear();
  DM->qSP_country->First();
  kol = DM->qSP_country->RecordCount;
  for (int i=1; i<=kol; i++)
    {
      ComboBoxSTRANA->Items->Add(DM->qSP_country->FieldByName("country")->AsString);
      DM->qSP_country->Next();
    }
  DM->qSP_country->First();

  //���������� �������
  ComboBoxGOROD->Items->Clear();
  DM->qSP_city->First();
  kol = DM->qSP_city->RecordCount;
  for (int i=1; i<=kol; i++)
    {
      ComboBoxGOROD->Items->Add(DM->qSP_city->FieldByName("city")->AsString);
      DM->qSP_city->Next();
    }
  DM->qSP_city->First();

  //���������� ��������
  ComboBoxOBEKT->Items->Clear();
  DM->qSP_obekt->First();
  kol = DM->qSP_obekt->RecordCount;
  for (int i=1; i<=kol; i++)
    {
      ComboBoxOBEKT->Items->Add(DM->qSP_obekt->FieldByName("obekt")->AsString);
      DM->qSP_obekt->Next();
    }
  DM->qSP_obekt->First();

  //���������� ��������
  ComboBoxGOSTINICA->Items->Clear();
  DM->qSP_gostinica->First();
  kol = DM->qSP_gostinica->RecordCount;
  for (int i=1; i<=kol; i++)
    {
      ComboBoxGOSTINICA->Items->Add(DM->qSP_gostinica->FieldByName("gostinica")->AsString);
      DM->qSP_gostinica->Next();
    }
  DM->qSP_gostinica->First();

}
//---------------------------------------------------------------------------

void __fastcall TVvod::EditTNChange(TObject *Sender)
{
  AnsiString Sql, grade;


  if (!EditZEX->Text.IsEmpty() && !EditTN->Text.IsEmpty())
    {
      //������� ������ �� ��������� �� AVANS
      Sql="select initcap(fam||' '||im||' '||ot) as fio,                            \
                  decode(a.kat, 4, (select prf from spprf where nprf=a.ppof),       \
                        (select ndolg from spdolg1 where dolg=a.ppof)||decode(nvl(a.dprof,0),0,'',' - '||(select ndolg from spdolg2 where dolg=a.ppof and dolg1=a.dprof))) as dolg, \
                  grade                                                             \
           from avans a where ncex="+EditZEX->Text+" and tn="+EditTN->Text;           \

      DM->qObnovlenie->Close();
      DM->qObnovlenie->SQL->Clear();
      DM->qObnovlenie->SQL->Add(Sql);

      try
        {
          DM->qObnovlenie->Open();
        }
      catch(Exception &E)
        {
          Application->MessageBox(("���������� �������� ������ �� ������� AVANS"+ E.Message).c_str(),"��������������",
                                  MB_OK+MB_ICONERROR);
        }

      EditFIO->Text=DM->qObnovlenie->FieldByName("fio")->AsString;
      EditPROF->Text=DM->qObnovlenie->FieldByName("dolg")->AsString;
      EditGRADE->Text=DM->qObnovlenie->FieldByName("grade")->AsString;




      if (EditGRADE->Text.IsEmpty())
        {
          EditG_KIEV->Text="";
          EditG_UKR->Text="";
          EditG_ZAGRAN->Text="";
        }
      else
        {
          if (!EditGRADE->Text.IsEmpty())
            {
              if (StrToInt(EditGRADE->Text)<=12) grade=12;
              else if (StrToInt(EditGRADE->Text)>=18) grade=18;
              else  grade=EditGRADE->Text;

              //������� ���� �� ������ �� �����������
              Sql="select g_kiev,g_ukr,g_zagran from sp_grade where grade="+grade;

              DM->qObnovlenie->Close();
              DM->qObnovlenie->SQL->Clear();
              DM->qObnovlenie->SQL->Add(Sql);

              try
                {
                  DM->qObnovlenie->Open();
                }
              catch(Exception &E)
                {
                  Application->MessageBox(("���������� �������� ������ �� ����������� ������� (SP_GRADE)"+E.Message).c_str(), "��������������",
                                            MB_OK+MB_ICONERROR);
                }

              EditG_KIEV->Text=DM->qObnovlenie->FieldByName("g_kiev")->AsString;
              EditG_UKR->Text=DM->qObnovlenie->FieldByName("g_ukr")->AsString;
              EditG_ZAGRAN->Text=DM->qObnovlenie->FieldByName("g_zagran")->AsString;
            }
          else
            {
              grade=EditGRADE->Text;
            }
        }       
    }
}
//---------------------------------------------------------------------------

AnsiString  __fastcall TVvod::SetNull (AnsiString str, AnsiString r)
{
  if (str.Length()) return str;
  else return r;
}
//---------------------------------------------------------------------------

void __fastcall TVvod::EditDATA_NExit(TObject *Sender)
{
  TDateTime d;

  if (ActiveControl == Cansel)
    {
      Vvod->Close();
    }
  else
    {
      if (!EditDATA_N->Text.IsEmpty())
        {
          // ���������� � ���� ��������� ������ � ����
          if (EditDATA_N->Text.Length()<3)
            {
              if(EditDATA_N->Text.Pos("."))
                {
                  Application->MessageBox("�������� ������ ����","������", MB_OK+MB_ICONINFORMATION);
                  EditDATA_N->Font->Color = clRed;
                  EditDATA_N->SetFocus();
                  Abort();
                }
              else
                {
                  EditDATA_N->Text = EditDATA_N->Text+ "."+ DateToStr(Date()).SubString(4,2) +"."+ DateToStr(Date()).SubString(7,5);
                  EditDATA_N->Font->Color = clBlack;
                }
            }

          // �������� �� ������������ ����� ����
          if(!TryStrToDate(EditDATA_N->Text,d))
            {
              Application->MessageBox("�������� ������ ����","������", MB_OK);
              EditDATA_N->Font->Color = clRed;
              EditDATA_N->SetFocus();
            }
          else
            {
              EditDATA_N->Text=FormatDateTime("dd.mm.yyyy",d);
              EditDATA_N->Font->Color = clBlack;
            }

          if (!EditDATA_K->Text.IsEmpty())
            {
              EditSROK->Text=DaysBetween(StrToDate(EditDATA_K->Text),StrToDate(EditDATA_N->Text))+1;
            }
        }
    }
}
//---------------------------------------------------------------------------

void __fastcall TVvod::EditDATA_KExit(TObject *Sender)
{
  TDateTime d;

  if (ActiveControl == Cansel)
    {
      Vvod->Close();
    }
  else
    {
      if (!EditDATA_K->Text.IsEmpty())
        {
          // ���������� � ���� ��������� ������ � ����
          if (EditDATA_K->Text.Length()<3)
            {
              if(EditDATA_K->Text.Pos("."))
                {
                  Application->MessageBox("�������� ������ ����","������", MB_OK+MB_ICONINFORMATION);
                  EditDATA_K->Font->Color = clRed;
                  EditDATA_K->SetFocus();
                  Abort();
                }
              else
                {
                  EditDATA_K->Text = EditDATA_K->Text+ "."+ DateToStr(Date()).SubString(4,2) +"."+ DateToStr(Date()).SubString(7,5);
                  EditDATA_K->Font->Color = clBlack;
                }
            }

          // �������� �� ������������ ����� ����
          if(!TryStrToDate(EditDATA_K->Text,d))
            {
              Application->MessageBox("�������� ������ ����","������", MB_OK);
              EditDATA_K->Font->Color = clRed;
              EditDATA_K->SetFocus();
            }
          else
            {
              EditDATA_K->Text=FormatDateTime("dd.mm.yyyy",d);
              EditDATA_K->Font->Color = clBlack;
            }

          if (!EditDATA_N->Text.IsEmpty())
            {
              EditSROK->Text=DaysBetween(StrToDate(EditDATA_K->Text),StrToDate(EditDATA_N->Text))+1;


              //������� ������ �� ������ �� ������ �����������
              AnsiString Sql;
              Sql="select * from                                                          \
                   ((select n_doc, data as dato, zex, tab, datnkom, datkkom, n_order from k_avans1@F) n    \
                   left join                                                              \
                   (select n_doc, data,                                                   \
                           decode(sum(nvl(sut,0))+sum(nvl(sut_bez,0)),0,NULL,to_number(sum(nvl(sut,0))+sum(nvl(sut_bez,0)))) as sut,         \
                           decode(sum(nvl(kvart,0))+sum(nvl(kvart_bez,0)),0,NULL,to_number(sum(nvl(kvart,0))+sum(nvl(kvart_bez,0)))) as kvart,   \
                           decode(sum(nvl(avia,0))+sum(nvl(avia_bez,0)),0,NULL,to_number(sum(nvl(avia,0))+sum(nvl(avia_bez,0)))) as avia,      \
                           decode(sum(nvl(gd,0))+sum(nvl(gd_bez,0)),0,NULL,to_number(sum(nvl(gd,0))+sum(nvl(gd_bez,0)))) as gd,            \
                           decode((sum(nvl(stop,0))+sum(nvl(proez,0))),0,NULL,to_number((sum(nvl(stop,0))+sum(nvl(proez,0))))) as proez, \
                           decode(sum(nvl(sum,0))+sum(nvl(viza_bez,0)),0,NULL,to_number(sum(nvl(sum,0))+sum(nvl(viza_bez,0)))) as proch        \
                    from k_avans2@F                                                       \
                    group by n_doc, data) d                                               \
                   on n.n_doc=d.n_doc and dato=d.data)                                  \
                   where n.zex="+EditZEX->Text+"  and substr(lpad(n.tab,5,'0'),2,4)=substr(lpad("+EditTN->Text+",5,'0'),2,4) and n.datnkom="+QuotedStr(EditDATA_N->Text)+" \
                   and n.datkkom="+QuotedStr(EditDATA_K->Text);

              DM->qObnovlenie->Close();
              DM->qObnovlenie->SQL->Clear();
              DM->qObnovlenie->SQL->Add(Sql);
              try
                 {
                   DM->qObnovlenie->Open();
                 }
               catch (Exception &E)
                 {
                   Application->MessageBox(("�������� ������ ��� ��������� ���� �� ������������� �� ������ ��� �����������"+E.Message).c_str(),"������",
                                           MB_OK+MB_ICONERROR);
                   Abort();
                 }

               EditSUM_SUT->Text=DM->qObnovlenie->FieldByName("sut")->AsString;
               EditSUM_PROGIV->Text=DM->qObnovlenie->FieldByName("kvart")->AsString;
               EditSUM_TRANSP->Text=DM->qObnovlenie->FieldByName("proez")->AsString;
               EditSUM_AVIA->Text=DM->qObnovlenie->FieldByName("avia")->AsString;
               EditSUM_GD->Text=DM->qObnovlenie->FieldByName("gd")->AsString;
               EditSUM_PROCH->Text=DM->qObnovlenie->FieldByName("proch")->AsString;
               EditN_DOCUM->Text=DM->qObnovlenie->FieldByName("n_doc")->AsString;
               EditDATA_ZAK->Text=DM->qObnovlenie->FieldByName("dato")->AsString;

            }
        }
    }
}
//---------------------------------------------------------------------------

void __fastcall TVvod::EditDATA_GOST_NExit(TObject *Sender)
{
  TDateTime d;

  if (ActiveControl == Cansel)
    {
      Vvod->Close();
    }
  else
    {
      if (!EditDATA_GOST_N->Text.IsEmpty())
        {
          // ���������� � ���� ��������� ������ � ����
          if (EditDATA_GOST_N->Text.Length()<3)
            {
              if(EditDATA_GOST_N->Text.Pos("."))
                {
                  Application->MessageBox("�������� ������ ����","������", MB_OK+MB_ICONINFORMATION);
                  EditDATA_GOST_N->Font->Color = clRed;
                  EditDATA_GOST_N->SetFocus();
                  Abort();
                }
              else
                {
                  EditDATA_GOST_N->Text = EditDATA_GOST_N->Text+ "."+ DateToStr(Date()).SubString(4,2) +"."+ DateToStr(Date()).SubString(7,5);
                  EditDATA_GOST_N->Font->Color = clBlack;
                }
            }

          // �������� �� ������������ ����� ����
          if(!TryStrToDate(EditDATA_GOST_N->Text,d))
            {
              Application->MessageBox("�������� ������ ����","������", MB_OK);
              EditDATA_GOST_N->Font->Color = clRed;
              EditDATA_GOST_N->SetFocus();
            }
          else
            {
              EditDATA_GOST_N->Text=FormatDateTime("dd.mm.yyyy",d);
              EditDATA_GOST_N->Font->Color = clBlack;
            }

        }
    }
}
//---------------------------------------------------------------------------

void __fastcall TVvod::EditDATA_GOST_KExit(TObject *Sender)
{
  TDateTime d;

  if (ActiveControl == Cansel)
    {
      Vvod->Close();
    }
  else
    {
      if (!EditDATA_GOST_K->Text.IsEmpty())
        {
          // ���������� � ���� ��������� ������ � ����
          if (EditDATA_GOST_K->Text.Length()<3)
            {
              if(EditDATA_GOST_K->Text.Pos("."))
                {
                  Application->MessageBox("�������� ������ ����","������", MB_OK+MB_ICONINFORMATION);
                  EditDATA_GOST_K->Font->Color = clRed;
                  EditDATA_GOST_K->SetFocus();
                  Abort();
                }
              else
                {
                  EditDATA_GOST_K->Text = EditDATA_GOST_K->Text+ "."+ DateToStr(Date()).SubString(4,2) +"."+ DateToStr(Date()).SubString(7,5);
                  EditDATA_GOST_K->Font->Color = clBlack;
                }
            }

          // �������� �� ������������ ����� ����
          if(!TryStrToDate(EditDATA_GOST_K->Text,d))
            {
              Application->MessageBox("�������� ������ ����","������", MB_OK);
              EditDATA_GOST_K->Font->Color = clRed;
              EditDATA_GOST_K->SetFocus();
            }
          else
            {
              EditDATA_GOST_K->Text=FormatDateTime("dd.mm.yyyy",d);
              EditDATA_GOST_K->Font->Color = clBlack;
            }

        }
    }
}
//---------------------------------------------------------------------------

void __fastcall TVvod::EditDATA_NKeyPress(TObject *Sender, char &Key)
{
  if (! (IsNumeric(Key) || Key=='\b' ||Key=='/' || Key==','|| Key=='.') ) Key=0;
  if (Key=='/' || Key==',') Key='.';        
}
//---------------------------------------------------------------------------

void __fastcall TVvod::EditZEXKeyPress(TObject *Sender, char &Key)
{
  if (! (IsNumeric(Key) || Key=='\b') ) Key=0;        
}
//---------------------------------------------------------------------------

void __fastcall TVvod::EditG_KIEVKeyPress(TObject *Sender, char &Key)
{
  if (! (IsNumeric(Key) || Key=='.' || Key==',' || Key=='/' || Key=='\b') ) Key=0;
  if (Key==',' || Key=='/') Key='.';        
}
//---------------------------------------------------------------------------

void __fastcall TVvod::EditFIOKeyPress(TObject *Sender, char &Key)
{
  if (IsNumeric(Key)) Key=0;        
}
//---------------------------------------------------------------------------

void __fastcall TVvod::ComboBoxSTRANAExit(TObject *Sender)
{
  TLocateOptions SearchOptions;
  AnsiString Sql;
  int kol;

  //���������� ComboBox �� ����� ��������������

  if (!ComboBoxSTRANA->Text.IsEmpty())
    {
      if (DM->qSP_country->Locate("country",ComboBoxSTRANA->Text,SearchOptions << loCaseInsensitive))
        {
         /* ComboBoxGOROD->ItemIndex=-1;
          ComboBoxOBEKT->ItemIndex=-1;
          ComboBoxGOSTINICA->ItemIndex=-1;
          EditADRESS->Text="";
          EditGOST_ADR->Text=""; */

          DM->qSP_city->Filtered=false;
          DM->qSP_city->Filter="kod_country="+QuotedStr(DM->qSP_country->FieldByName("kod")->AsString);
          DM->qSP_city->Filtered=true;

          //���������� ����������� ������� � ����������� �� ��������� ������
          ComboBoxGOROD->Items->Clear();
          DM->qSP_city->First();
          kol = DM->qSP_city->RecordCount;
          for (int i=1; i<=kol; i++)
            {
              ComboBoxGOROD->Items->Add(DM->qSP_city->FieldByName("city")->AsString);
              DM->qSP_city->Next();
            }

          DM->qSP_city->Filtered=false;
          DM->qSP_city->First();
        }
      else
        {
          if (Application->MessageBox("�������� ������ ��� � �����������. \n�������� ������ ������ � ����������?","��������������",
                                      MB_YESNO+MB_ICONWARNING)==IDYES)
            {
              //���������� ������ � ���������� �����
              Sprav->PageControl1->ActivePage = Sprav->TabSheet5; //�������� ��������

              Sprav->TabSheet5->Caption = "������";
             /* Sprav->TabSheet5->TabVisible = true;

              Sprav->TabSheet1->TabVisible = false;
              Sprav->TabSheet2->TabVisible = false;
              Sprav->TabSheet3->TabVisible = false;
              Sprav->TabSheet4->TabVisible = false;
              Sprav->TabSheet6->TabVisible = false;  */

              Sprav->ShowModal();

              ComboBoxSTRANA->SetFocus();
            }
          else
            {
              ComboBoxSTRANA->SetFocus();
              Abort();
            }



        }

      //�������� ������ � ������� �����
      ComboBoxSTRANA->Text= AnsiUpperCase((ComboBoxSTRANA->Text).SubString(1,1))+(ComboBoxSTRANA->Text).SubString(2,255);
    }
  else
    {
      ComboBoxGOROD->Text="";
      ComboBoxOBEKT->Text="";
      ComboBoxGOSTINICA->Text="";

      EditADRESS->Text="";
      EditGOST_ADR->Text="";

      ComboBoxGOROD->Clear();
      ComboBoxOBEKT->Clear();
      ComboBoxGOSTINICA->Clear();

    }
}
//---------------------------------------------------------------------------


void __fastcall TVvod::ComboBoxGORODExit(TObject *Sender)
{
  TLocateOptions SearchOptions;
  AnsiString Sql;
  int kol;

  //���������� ComboBox �� ����� ��������������

  if (!ComboBoxGOROD->Text.IsEmpty())
    {
      if (DM->qSP_city->Locate("city",ComboBoxGOROD->Text,SearchOptions << loCaseInsensitive))
        {
          if (ComboBoxSTRANA->Text.IsEmpty())
            {
              ComboBoxGOSTINICA->ItemIndex=-1;
              ComboBoxOBEKT->ItemIndex=-1;
              EditADRESS->Text="";
              EditGOST_ADR->Text="";

              ComboBoxGOROD->Clear();
              ComboBoxOBEKT->Clear();
              ComboBoxGOSTINICA->Clear();
            }  

          DM->qSP_obekt->Filtered=false;
          DM->qSP_obekt->Filter="kod_city="+QuotedStr(DM->qSP_city->FieldByName("kod")->AsString);
          DM->qSP_obekt->Filtered=true;

          //���������� ����������� �������� � ����������� �� ������
          ComboBoxOBEKT->Items->Clear();
          DM->qSP_obekt->First();
          kol = DM->qSP_obekt->RecordCount;
          for (int i=1; i<=kol; i++)
            {
              ComboBoxOBEKT->Items->Add(DM->qSP_obekt->FieldByName("obekt")->AsString);
              DM->qSP_obekt->Next();
            }

          DM->qSP_obekt->Filtered=false;
          DM->qSP_obekt->First();
        }
      else
        {
          if (Application->MessageBox("��������� ������ ��� � �����������. \n�������� ������ ������ � ����������?","��������������",
                                      MB_YESNO+MB_ICONWARNING)==IDYES)
            {
              //���������� ������ � ���������� �������
              Sprav->PageControl1->ActivePage = Sprav->TabSheet6; //�������� ��������
              Sprav->ShowModal();

              ComboBoxGOROD->SetFocus();
            }
          else
            {
              ComboBoxGOROD->SetFocus();
              Abort();
            }
        }

      DM->qSP_city->Locate("city",ComboBoxGOROD->Text,SearchOptions << loCaseInsensitive);
      //ComboBoxGOSTINICA->ItemIndex=-1;

      DM->qSP_gostinica->Filtered=false;
      DM->qSP_gostinica->Filter="kod_city="+QuotedStr(DM->qSP_city->FieldByName("kod")->AsString);
      DM->qSP_gostinica->Filtered=true;

      //���������� ����������� �������� � ����������� �� ������
      ComboBoxGOSTINICA->Items->Clear();
      DM->qSP_gostinica->First();
      kol = DM->qSP_gostinica->RecordCount;
      for (int i=1; i<=kol; i++)
        {
          ComboBoxGOSTINICA->Items->Add(DM->qSP_gostinica->FieldByName("gostinica")->AsString);
          DM->qSP_gostinica->Next();
        }

      DM->qSP_gostinica->Filtered=false;
      DM->qSP_gostinica->First();

      //�������� ������ � ������� �����
      ComboBoxGOROD->Text= AnsiUpperCase((ComboBoxGOROD->Text).SubString(1,1))+(ComboBoxGOROD->Text).SubString(2,255);
    }
  else
    {
      ComboBoxGOSTINICA->ItemIndex=-1;
      ComboBoxOBEKT->ItemIndex=-1;
      EditADRESS->Text="";
      EditGOST_ADR->Text="";

      ComboBoxOBEKT->Clear();
      ComboBoxGOSTINICA->Clear();
    }

}
//---------------------------------------------------------------------------

void __fastcall TVvod::ComboBoxCHELExit(TObject *Sender)
{
  //�������� ���� � ������� �����
  ComboBoxCHEL->Text= AnsiUpperCase((ComboBoxCHEL->Text).SubString(1,1))+(ComboBoxCHEL->Text).SubString(2,255);
}
//---------------------------------------------------------------------------

void __fastcall TVvod::ComboBoxOBEKTChange(TObject *Sender)
{
  TLocateOptions SearchOptions;
  AnsiString Sql;
  int kol;

  if (!ComboBoxOBEKT->Text.IsEmpty())
    {
      if (DM->qSP_obekt->Locate("obekt",ComboBoxOBEKT->Text,SearchOptions << loCaseInsensitive))
        {
          //����� ������ ���������
          EditADRESS->Text=DM->qSP_obekt->FieldByName("adress")->AsString;
        }
    }
}
//---------------------------------------------------------------------------

void __fastcall TVvod::ComboBoxGOSTINICAChange(TObject *Sender)
{
  TLocateOptions SearchOptions;
  AnsiString Sql;
  int kol;

  if (!ComboBoxGOSTINICA->Text.IsEmpty())
    {
      if (DM->qSP_gostinica->Locate("gostinica",ComboBoxGOSTINICA->Text,SearchOptions << loCaseInsensitive))
        {
          //����� ������ ���������
          EditGOST_ADR->Text=DM->qSP_gostinica->FieldByName("adress")->AsString;
        }
    }
}
//---------------------------------------------------------------------------

void __fastcall TVvod::ComboBoxOBEKTExit(TObject *Sender)
{
  TLocateOptions SearchOptions;
  AnsiString Sql;
  int kol;

  if (!ComboBoxOBEKT->Text.IsEmpty())
    {
      if (!DM->qSP_obekt->Locate("obekt",ComboBoxOBEKT->Text,SearchOptions << loCaseInsensitive))
        {
          if (Application->MessageBox("��������� ������� ��� � �����������. \n�������� ������ ������ � ����������?","��������������",
                                      MB_YESNO+MB_ICONWARNING)==IDYES)
            {
              //���������� ������� � ����������
              Sprav->PageControl1->ActivePage = Sprav->TabSheet4; //�������� ��������
              Sprav->ShowModal();

              ComboBoxOBEKT->SetFocus();
            }
          else
            {
              ComboBoxOBEKT->SetFocus();
              Abort();
            }

          
    /*      DM->qSP_gostinica->Filtered=false;
          DM->qSP_gostinica->Filter="kod_city="+QuotedStr(DM->qSP_city->FieldByName("kod")->AsString);
          DM->qSP_gostinica->Filtered=true;

          //���������� ����������� �������� � ����������� �� �������
          ComboBoxGOSTINICA->Items->Clear();
          DM->qSP_gostinica->First();
          kol = DM->qSP_gostinica->RecordCount;
          for (int i=1; i<=kol; i++)
            {
              ComboBoxGOSTINICA->Items->Add(DM->qSP_gostinica->FieldByName("gostinica")->AsString);
              DM->qSP_gostinica->Next();
            }

          DM->qSP_gostinica->Filtered=false;
          DM->qSP_gostinica->First(); */


          if (ComboBoxSTRANA->Text.IsEmpty() || ComboBoxGOROD->Text.IsEmpty())
            {
              ComboBoxGOSTINICA->ItemIndex=-1;
              ComboBoxOBEKT->ItemIndex=-1;
              EditADRESS->Text="";
              EditGOST_ADR->Text="";
              EditGOST_ADR->Text="";
            }


        }

      //�������� ������ � ������� �����
      ComboBoxOBEKT->Text= AnsiUpperCase((ComboBoxOBEKT->Text).SubString(1,1))+(ComboBoxOBEKT->Text).SubString(2,255);
    }
  else
    {
      EditADRESS->Text="";
    }
}
//---------------------------------------------------------------------------

void __fastcall TVvod::ComboBoxGOSTINICAExit(TObject *Sender)
{
  TLocateOptions SearchOptions;
  AnsiString Sql;
  int kol;
  
 if (!ComboBoxGOSTINICA->Text.IsEmpty())
    {
      if (!DM->qSP_gostinica->Locate("gostinica",ComboBoxGOSTINICA->Text,SearchOptions << loCaseInsensitive))
        {
          if (Application->MessageBox("�������� ��������� ��� � �����������. \n�������� ������ ������ � ����������?","��������������",
                                      MB_YESNO+MB_ICONWARNING)==IDYES)
            {
              //���������� ������� � ����������
              Sprav->PageControl1->ActivePage = Sprav->TabSheet3; //�������� ��������
              Sprav->ShowModal();

              ComboBoxGOSTINICA->SetFocus();
            }
          else
            {
              ComboBoxGOSTINICA->SetFocus();
              Abort();
            }
        }

      //�������� ��������� � ������� �����
      ComboBoxGOSTINICA->Text= AnsiUpperCase((ComboBoxGOSTINICA->Text).SubString(1,1))+(ComboBoxGOSTINICA->Text).SubString(2,255);
    }
  else
    {
      EditGOST_ADR->Text="";
    }
}
//---------------------------------------------------------------------------

void __fastcall TVvod::EditZEXChange(TObject *Sender)
{
 AnsiString Sql;


  if (!EditZEX->Text.IsEmpty())
    {
      Sql="select imck from spnc where nc="+EditZEX->Text;

      DM->qObnovlenie->Close();
      DM->qObnovlenie->SQL->Clear();
      DM->qObnovlenie->SQL->Add(Sql);
      try
        {
          DM->qObnovlenie->Open();
        }
      catch(Exception &E)
        {
          Application->MessageBox(("�� �������� ������� ������������ ���� �� ����������� ����� "+E.Message).c_str(),"������",
                               MB_OK+MB_ICONERROR);
        }
      LabelZEX->Caption=DM->qObnovlenie->FieldByName("imck")->AsString;
    }
  else
    {
      LabelZEX->Caption="";
    }
}
//---------------------------------------------------------------------------

void __fastcall TVvod::EditZEXExit(TObject *Sender)
{
  AnsiString Sql;


  if (!EditZEX->Text.IsEmpty())
    {
      Sql="select imck from spnc where nc="+EditZEX->Text;

      DM->qObnovlenie->Close();
      DM->qObnovlenie->SQL->Clear();
      DM->qObnovlenie->SQL->Add(Sql);
      try
        {
          DM->qObnovlenie->Open();
        }
      catch(Exception &E)
        {
          Application->MessageBox(("�� �������� ������� ������������ ���� �� ����������� ����� "+E.Message).c_str(),"������",
                               MB_OK+MB_ICONERROR);
        }
      LabelZEX->Caption=DM->qObnovlenie->FieldByName("imck")->AsString;
    }
  else
    {
      LabelZEX->Caption="";
    } 
}
//---------------------------------------------------------------------------


void __fastcall TVvod::EditGRADEExit(TObject *Sender)
{
  /*AnsiString grade, Sql;

  if (!EditGRADE->Text.IsEmpty())
    {
      if (StrToInt(EditGRADE->Text)<=12) grade=12;
      else if (StrToInt(EditGRADE->Text)>=18) grade=18;
      else  grade=EditGRADE->Text;

      //������� ���� �� ������ �� �����������
      Sql="select g_kiev,g_ukr,g_zagran from sp_grade where grade="+grade;

      DM->qObnovlenie->Close();
      DM->qObnovlenie->SQL->Clear();
      DM->qObnovlenie->SQL->Add(Sql);

      try
        {
          DM->qObnovlenie->Open();
        }
      catch(Exception &E)
        {
          Application->MessageBox(("���������� �������� ������ �� ����������� ������� (SP_GRADE)"+E.Message).c_str(), "��������������",
                                    MB_OK+MB_ICONERROR);
        }

      EditG_KIEV->Text=DM->qObnovlenie->FieldByName("g_kiev")->AsString;
      EditG_UKR->Text=DM->qObnovlenie->FieldByName("g_ukr")->AsString;
      EditG_ZAGRAN->Text=DM->qObnovlenie->FieldByName("g_zagran")->AsString;
    }  */
}
//---------------------------------------------------------------------------

void __fastcall TVvod::EditGRADEChange(TObject *Sender)
{
  AnsiString grade, Sql;

  if (!EditGRADE->Text.IsEmpty())
    {
      if (StrToInt(EditGRADE->Text)<=12) grade=12;
      else if (StrToInt(EditGRADE->Text)>=18) grade=18;
      else  grade=EditGRADE->Text;

      //������� ���� �� ������ �� �����������
      Sql="select g_kiev,g_ukr,g_zagran from sp_grade where grade="+grade;

      DM->qObnovlenie->Close();
      DM->qObnovlenie->SQL->Clear();
      DM->qObnovlenie->SQL->Add(Sql);

      try
        {
          DM->qObnovlenie->Open();
        }
      catch(Exception &E)
        {
          Application->MessageBox(("���������� �������� ������ �� ����������� ������� (SP_GRADE)"+E.Message).c_str(), "��������������",
                                    MB_OK+MB_ICONERROR);
        }

      EditG_KIEV->Text=DM->qObnovlenie->FieldByName("g_kiev")->AsString;
      EditG_UKR->Text=DM->qObnovlenie->FieldByName("g_ukr")->AsString;
      EditG_ZAGRAN->Text=DM->qObnovlenie->FieldByName("g_zagran")->AsString;
    }
}
//---------------------------------------------------------------------------

