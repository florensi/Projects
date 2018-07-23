//---------------------------------------------------------------------------

#include <vcl.h>
#pragma hdrstop

#include "uVvod.h"
#include "uMain.h"
#include "uDM.h"
//---------------------------------------------------------------------------
#pragma package(smart_init)
#pragma resource "*.dfm"
TVvod *Vvod;
//---------------------------------------------------------------------------
__fastcall TVvod::TVvod(TComponent* Owner)
        : TForm(Owner)
{
}
//---------------------------------------------------------------------------
void __fastcall TVvod::BitBtn2Click(TObject *Sender)
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

void __fastcall TVvod::EditCHFKeyPress(TObject *Sender, char &Key)
{
  if (! (IsNumeric(Key) || Key=='.' || Key==',' || Key=='/' || Key=='\b' || Key=='�'|| Key=='�' || Key=='-') ) Key=0;
  if (Key==',' || Key=='/') Key='.';
}
//---------------------------------------------------------------------------

//�������������� �����
void __fastcall TVvod::BitBtn1Click(TObject *Sender)
{
  AnsiString Sql, chf, nsm, pgraf;
  int row, kol;
  double ochf;

  //�������� ���� �� ��������� ��������������
 /* if (Main->znsm == EditNSM->Text &&
      Main->zchf == EditCHF->Text &&
      Main->zpch == EditPCH->Text &&
      Main->zvch == EditVCH->Text &&
      Main->znch == EditNCH->Text &&
      Main->zchf0 == EditCHF0->Text &&
      Main->znch0 == EditNCH0->Text &&
      Main->zpch0 == EditPCH0->Text )
    {
      Vvod->Close();
      Abort();
    }  */

  if (AnsiUpperCase(EditCHF->Text)=="�" || AnsiUpperCase(EditNSM->Text)=="�" )

    {
     

     if (DM->qGrafik->FieldByName("ograf")->AsInteger==180 && AnsiUpperCase(EditCHF->Text)!="�")
       {
         chf = SetN(EditCHF->Text);
         nsm = 9;
         ochf = 1;
       }
     else
       {
         chf = 30;
         nsm = 9;
         ochf = 0;
       }

    /* chf = ochf = SetN(EditCHF->Text);
     ShowMessage(chf);
     ShowMessage(ochf);
     ShowMessage(SetN(EditCHF->Text));  */

      
    }
  else
    {
      chf = ochf = SetN(EditCHF->Text);
      nsm = SetN(EditNSM->Text);

    }

  //�������� ���� �������� ����� ����� ����� �� ��������� ������������ �����
  if (ochf > SetN(DM->qGrafik->FieldByName("dlit")->AsString)+0.5)
    {
      Application->MessageBox("��������� ����� ���� ��������� ������������ �����","���������� ������������",
                               MB_OK + MB_ICONWARNING);
      EditCHF->SetFocus();
      Abort();
    }

   //��������, ���� ������ ��������, ���� �� ���� �����
  /* if ((AnsiUpperCase(EditCHF->Text)=="�" || AnsiUpperCase(EditNSM->Text)=="�") &&
                                                                               (SetN(EditPCH->Text)>0 ||
                                                                                SetN(EditVCH->Text)>0 ||
                                                                                SetN(EditNCH->Text)>0 ||
                                                                                SetN(EditPCH0->Text)>0 ||
                                                                                SetN(EditCHF0->Text)>0 ||
                                                                                SetN(EditNCH0->Text)>0))
     {
       Application->MessageBox("��� ����������� �������� ��� ��� ������� �����!!!","��������������",
                               MB_OK+MB_ICONWARNING);
       EditNSM->SetFocus();
       Abort();
     }  */

   //�������� ���� �� ���� ����� = NULL,� ���� ����
   if ( EditNSM->Text.IsEmpty() && (ochf!=0 ||
                   SetN(EditPCH->Text)!=0 ||
                   SetN(EditVCH->Text)!=0 ||
                   SetN(EditNCH->Text)!=0))
     {
       Application->MessageBox("������� ����, �� �� ������ ����� �����","��������������",
                               MB_OK+MB_ICONWARNING);
       EditNSM->SetFocus();
       Abort();
     }
   else if (nsm!=0 && ochf==0 && chf!=30 &&
                      SetN(EditPCH->Text)==0 &&
                      SetN(EditVCH->Text)==0 &&
                      SetN(EditNCH->Text)==0)
     {
       Application->MessageBox("������ ����� �����, �� �� ������� ����","��������������",
                               MB_OK+MB_ICONWARNING);
       EditCHF->SetFocus();
       Abort();
     }

  // ����������� ���������� ���� � ������
 // kol = DaysInAMonth(DM->qGrafik->FieldByName("god")->AsInteger, DM->qGrafik->FieldByName("mes")->AsInteger);

  //���� ������ �������� ����� "�"
  if (AnsiUpperCase(Main->zchf)=="�" || AnsiUpperCase(Main->zchf)=="-") Main->zchf=0;

  if (DM->qGrafik->FieldByName("ograf")->AsInteger!=11 &&
         DM->qGrafik->FieldByName("ograf")->AsInteger!=1011 &&
         DM->qGrafik->FieldByName("ograf")->AsInteger!=2011 &&
         DM->qGrafik->FieldByName("ograf")->AsInteger!=3011 &&
         DM->qGrafik->FieldByName("ograf")->AsInteger!=18 &&
         DM->qGrafik->FieldByName("ograf")->AsInteger!=1018 &&
         DM->qGrafik->FieldByName("ograf")->AsInteger!=2018 &&
         DM->qGrafik->FieldByName("ograf")->AsInteger!=3018 &&
         DM->qGrafik->FieldByName("ograf")->AsInteger!=20 &&
         DM->qGrafik->FieldByName("ograf")->AsInteger!=1020 &&
         DM->qGrafik->FieldByName("ograf")->AsInteger!=2020 &&
         DM->qGrafik->FieldByName("ograf")->AsInteger!=23 &&
         DM->qGrafik->FieldByName("ograf")->AsInteger!=24 &&
         DM->qGrafik->FieldByName("ograf")->AsInteger!=50 &&
         DM->qGrafik->FieldByName("ograf")->AsInteger!=81 &&
         DM->qGrafik->FieldByName("ograf")->AsInteger!=90 &&
         DM->qGrafik->FieldByName("ograf")->AsInteger!=111 &&
         DM->qGrafik->FieldByName("ograf")->AsInteger!=120 &&
         DM->qGrafik->FieldByName("ograf")->AsInteger!=150 &&
         DM->qGrafik->FieldByName("ograf")->AsInteger!=160 &&
         DM->qGrafik->FieldByName("ograf")->AsInteger!=170 &&
         DM->qGrafik->FieldByName("ograf")->AsInteger!=180 &&
         DM->qGrafik->FieldByName("ograf")->AsInteger!=190 &&
         DM->qGrafik->FieldByName("ograf")->AsInteger!=220 &&
         DM->qGrafik->FieldByName("ograf")->AsInteger!=230 &&
         DM->qGrafik->FieldByName("ograf")->AsInteger!=250 &&
         DM->qGrafik->FieldByName("ograf")->AsInteger!=270 &&
         DM->qGrafik->FieldByName("ograf")->AsInteger!=315 &&
         DM->qGrafik->FieldByName("ograf")->AsInteger!=470 &&
         DM->qGrafik->FieldByName("ograf")->AsInteger!=480 &&
         DM->qGrafik->FieldByName("ograf")->AsInteger!=660 &&
         DM->qGrafik->FieldByName("ograf")->AsInteger!=780 &&
         DM->qGrafik->FieldByName("ograf")->AsInteger!=790 &&
         DM->qGrafik->FieldByName("ograf")->AsInteger!=820 &&
         DM->qGrafik->FieldByName("ograf")->AsInteger!=830 &&
         DM->qGrafik->FieldByName("ograf")->AsInteger!=855 &&
         DM->qGrafik->FieldByName("ograf")->AsInteger!=880)
      {
        //����������� ����� �����������
        if ((SetN(DM->qGrafik->FieldByName("chf")->AsString) - SetN(Main->zchf) + ochf) -
             SetN(DM->qGrafik->FieldByName("norma")->AsString) -
            (SetN(DM->qGrafik->FieldByName("pch")->AsString) - SetN(Main->zpch) + SetN(EditPCH->Text))<0)
          {
            pgraf = NULL;
          }
        else
          {
            pgraf = SetNull((SetN(DM->qGrafik->FieldByName("chf")->AsString) - SetN(Main->zchf) + ochf) -
                             SetN(DM->qGrafik->FieldByName("norma")->AsString) -
                            (SetN(DM->qGrafik->FieldByName("pch")->AsString) - SetN(Main->zpch)+ SetN(EditPCH->Text)));


          /*  ShowMessage(("����������� = "+pgraf).c_str());
            ShowMessage(("chf = "+DM->qGrafik->FieldByName("chf")->AsString).c_str());
            ShowMessage(("zchf = "+Main->zchf).c_str());
            ShowMessage(("ochf = "+FloatToStr(ochf)).c_str());
            ShowMessage(("norma = "+DM->qGrafik->FieldByName("norma")->AsString).c_str());
            ShowMessage(("pch = "+DM->qGrafik->FieldByName("pch")->AsString).c_str());
            ShowMessage(("zpch = "+Main->zpch).c_str());
            ShowMessage(("opch = "+EditPCH->Text).c_str());   */


          }
      }
    else
      {
        pgraf = NULL;
      }


     /*
      ShowMessage(SetNull(SetN(DM->qGrafik->FieldByName("chf")->AsString) - SetN(Main->zchf) + ochf));
      ShowMessage(SetN(DM->qGrafik->FieldByName("chf")->AsString));
      ShowMessage(SetN(Main->zchf));
      ShowMessage(ochf);  */


      //���������� �������� �� ���                                  `
      Sql = "update spgrafiki set chf"+Main->numk+"="+chf+",\
                                  pch"+Main->numk+"="+SetNull(EditPCH->Text)+",\
                                  vch"+Main->numk+"="+SetNull(EditVCH->Text)+",\
                                  nch"+Main->numk+"="+SetNull(EditNCH->Text)+", \
                                  nsm"+Main->numk+"="+nsm+", \
                                  chf="+SetNull(SetN(DM->qGrafik->FieldByName("chf")->AsString) - SetN(Main->zchf) + ochf)+",    \
                                  vch="+SetNull(SetN(DM->qGrafik->FieldByName("vch")->AsString) - SetN(Main->zvch) + SetN(EditVCH->Text))+",                                                  \
                                  nch="+SetNull(SetN(DM->qGrafik->FieldByName("nch")->AsString) - SetN(DM->qGrafik->FieldByName("nch"+Main->numk)->AsString) + SetN(EditNCH->Text))+", \
                                  pch="+SetNull(SetN(DM->qGrafik->FieldByName("pch")->AsString) - SetN(DM->qGrafik->FieldByName("pch"+Main->numk)->AsString) + SetN(EditPCH->Text))+", \
                                  pgraf="+pgraf+" \
             where ograf="+DM->qGrafik->FieldByName("ograf")->AsString+" and graf= "+SetNull(DM->qGrafik->FieldByName("graf")->AsString)+" \
             and mes="+DM->qGrafik->FieldByName("mes")->AsString+ "and god="+DM->qGrafik->FieldByName("god")->AsString;

      row=DM->qGrafik->RecNo;

      DM->qObnovlenie->Close();
      DM->qObnovlenie->SQL->Clear();
      DM->qObnovlenie->SQL->Add(Sql);
      try
        {
          DM->qObnovlenie->ExecSQL();
        }
      catch(...)
        {
          Application->MessageBox("�������� ������ �� ����� ���������� ������","������",
                                   MB_OK+MB_ICONERROR);
          Abort();
        }

      DM->qGrafik->Requery();
  // }


  DM->qGrafik->RecNo=row;

  Main->InsertLog("��������� �������������� "+DM->qGrafik->FieldByName("ograf")->AsString+" ������� �� "+
                   (StrToInt(Main->numk)<10 ? "0"+Main->numk : Main->numk)+"."+
                   (DM->qGrafik->FieldByName("mes")->AsInteger<10 ? "0"+DM->qGrafik->FieldByName("mes")->AsString : DM->qGrafik->FieldByName("mes")->AsString)+"."+
                   IntToStr(Main->god)+". ����� � "+
                   QuotedStr(Main->znsm)+" �� "+QuotedStr(Vvod->EditNSM->Text)+", ����� ���� � "+
                   QuotedStr(Main->zchf)+" �� "+QuotedStr(Vvod->EditCHF->Text)+", ����������� � "+
                   QuotedStr(Main->zpch)+" �� "+QuotedStr(Vvod->EditPCH->Text)+", �������� � "+
                   QuotedStr(Main->zvch)+" �� "+QuotedStr(Vvod->EditVCH->Text)+", ������ � "+
                   QuotedStr(Main->znch)+" �� "+QuotedStr(Vvod->EditNCH->Text)+". �����������: ����� � "+
                   QuotedStr(Main->zchf0)+" �� "+QuotedStr(Vvod->EditCHF0->Text)+", ������ � "+
                   QuotedStr(Main->znch0)+" �� "+QuotedStr(Vvod->EditNCH0->Text)+", ����������� � "+
                   QuotedStr(Main->zpch0)+" �� "+QuotedStr(Vvod->EditPCH0->Text));

  Vvod->Close();
}
//---------------------------------------------------------------------------


AnsiString __fastcall TVvod::SetNull(AnsiString str, AnsiString r)
{
  if (str.Length()) return str;
  else return r;
}
//---------------------------------------------------------------------------

double __fastcall TVvod::SetN(AnsiString str, double r)
{
  if (str.Length()) return StrToFloat(str);
  else return r;
}
//---------------------------------------------------------------------------

void __fastcall TVvod::EditPCHKeyPress(TObject *Sender, char &Key)
{
  if (! (IsNumeric(Key) || Key=='.' || Key==',' || Key=='/' || Key=='\b') ) Key=0;
  if (Key==',' || Key=='/') Key='.';        
}
//---------------------------------------------------------------------------

void __fastcall TVvod::FormShow(TObject *Sender)
{
  Vvod->EditNSM->SetFocus();        
}
//---------------------------------------------------------------------------

void __fastcall TVvod::EditNSMExit(TObject *Sender)
{
  //�������� Edit-��, ���� ��������
  if (AnsiUpperCase(EditNSM->Text)=="�" && DM->qGrafik->FieldByName("ograf")->AsInteger!=180)
    {
      EditNSM->Text = "�";
      EditCHF->Text = "�";
      EditPCH->Text = "";
      EditVCH->Text = "";
      EditNCH->Text = "";
      EditCHF0->Text = "";
      EditPCH0->Text = "";
      EditNCH0->Text = "";

      EditPCH->Enabled = false;
      EditVCH->Enabled = false;
      EditNCH->Enabled = false;
      EditCHF0->Enabled = false;
      EditPCH0->Enabled = false;
      EditNCH0->Enabled = false;
    }
  else
    {
      DostupRedaktEdit();
    }    
}
//---------------------------------------------------------------------------

void __fastcall TVvod::EditCHFExit(TObject *Sender)
{
  //�������� Edit-��, ���� ��������
  if (AnsiUpperCase(EditCHF->Text)=="�")
    {
      EditCHF->Text = "�";
      EditNSM->Text = "�";
      EditPCH->Text = "";
      EditVCH->Text = "";
      EditNCH->Text = "";
      EditCHF0->Text = "";
      EditPCH0->Text = "";
      EditNCH0->Text = "";
    }
}
//---------------------------------------------------------------------------

//������ � �������������� �����
void __fastcall TVvod::DostupRedaktEdit()
{
  //��� ������, ����������� � ��������
  if (DM->qGrafik->FieldByName("ograf")->AsInteger==11 ||
      DM->qGrafik->FieldByName("ograf")->AsInteger==18 ||
      DM->qGrafik->FieldByName("ograf")->AsInteger==81 ||
      DM->qGrafik->FieldByName("ograf")->AsInteger==111 ||
      DM->qGrafik->FieldByName("ograf")->AsInteger==480 ||
      DM->qGrafik->FieldByName("ograf")->AsInteger==650 ||
      DM->qGrafik->FieldByName("ograf")->AsInteger==655 ||
      DM->qGrafik->FieldByName("ograf")->AsInteger==660 ||
      DM->qGrafik->FieldByName("ograf")->AsInteger==771 ||
      DM->qGrafik->FieldByName("ograf")->AsInteger==780 ||
      DM->qGrafik->FieldByName("ograf")->AsInteger==800 ||
      DM->qGrafik->FieldByName("ograf")->AsInteger==820 ||
      DM->qGrafik->FieldByName("ograf")->AsInteger==830 ||
      DM->qGrafik->FieldByName("ograf")->AsInteger==1011 ||
      DM->qGrafik->FieldByName("ograf")->AsInteger==1018 ||
      DM->qGrafik->FieldByName("ograf")->AsInteger==1655 ||
      DM->qGrafik->FieldByName("ograf")->AsInteger==2011 ||
      DM->qGrafik->FieldByName("ograf")->AsInteger==2018 ||
      DM->qGrafik->FieldByName("ograf")->AsInteger==3011 ||
      DM->qGrafik->FieldByName("ograf")->AsInteger==3018 )
    {
      Vvod->EditPCH->Enabled = false;
      Vvod->EditVCH->Enabled = false;
      Vvod->EditNCH->Enabled = false;
    }
  //������ ��������
  else if (DM->qGrafik->FieldByName("ograf")->AsInteger==230 ||
           DM->qGrafik->FieldByName("ograf")->AsInteger==280 ||
           DM->qGrafik->FieldByName("ograf")->AsInteger==315 ||
           DM->qGrafik->FieldByName("ograf")->AsInteger==410 ||
           DM->qGrafik->FieldByName("ograf")->AsInteger==690 ||
           DM->qGrafik->FieldByName("ograf")->AsInteger==855 ||
           DM->qGrafik->FieldByName("ograf")->AsInteger==865 ||
           DM->qGrafik->FieldByName("ograf")->AsInteger==880)
    {
      Vvod->EditPCH->Enabled = false;
      Vvod->EditVCH->Enabled = true;
      Vvod->EditNCH->Enabled = false;
    }
  //������ ������
  else if (DM->qGrafik->FieldByName("ograf")->AsInteger==85)
    {
      Vvod->EditPCH->Enabled = false;
      Vvod->EditVCH->Enabled = false;
      Vvod->EditNCH->Enabled = true;
    }
  //������ �����������
  else if (DM->qGrafik->FieldByName("ograf")->AsInteger==30 ||
           DM->qGrafik->FieldByName("ograf")->AsInteger==150 ||
           DM->qGrafik->FieldByName("ograf")->AsInteger==630)
    {
      Vvod->EditPCH->Enabled = true;
      Vvod->EditVCH->Enabled = false;
      Vvod->EditNCH->Enabled = false;
    }
  //������ �������� � ������
  else if (DM->qGrafik->FieldByName("ograf")->AsInteger==25 ||
           DM->qGrafik->FieldByName("ograf")->AsInteger==140 ||
           DM->qGrafik->FieldByName("ograf")->AsInteger==160 ||
           DM->qGrafik->FieldByName("ograf")->AsInteger==470 ||
           DM->qGrafik->FieldByName("ograf")->AsInteger==775 ||
           DM->qGrafik->FieldByName("ograf")->AsInteger==2020)
    {
      Vvod->EditPCH->Enabled = false;
      Vvod->EditVCH->Enabled = true;
      Vvod->EditNCH->Enabled = true;
    }
  //������ �������� � �����������
  else if (DM->qGrafik->FieldByName("ograf")->AsInteger==131 ||
           DM->qGrafik->FieldByName("ograf")->AsInteger==190 ||
           DM->qGrafik->FieldByName("ograf")->AsInteger==300 ||
           DM->qGrafik->FieldByName("ograf")->AsInteger==1030 ||
           DM->qGrafik->FieldByName("ograf")->AsInteger==2030 ||
           DM->qGrafik->FieldByName("ograf")->AsInteger==3030 ||
           DM->qGrafik->FieldByName("ograf")->AsInteger==335 ||
           DM->qGrafik->FieldByName("ograf")->AsInteger==400 ||
           DM->qGrafik->FieldByName("ograf")->AsInteger==670 ||
           DM->qGrafik->FieldByName("ograf")->AsInteger==680 ||
           DM->qGrafik->FieldByName("ograf")->AsInteger==790 ||
           DM->qGrafik->FieldByName("ograf")->AsInteger==850 ||
           DM->qGrafik->FieldByName("ograf")->AsInteger==935)
    {
      Vvod->EditVCH->Enabled = true;
      Vvod->EditPCH->Enabled = true;
      Vvod->EditNCH->Enabled = false;
    }
  //���
  else
    {
      Vvod->EditVCH->Enabled = true;
      Vvod->EditNCH->Enabled = true;
      Vvod->EditPCH->Enabled = true;
    }
}
//---------------------------------------------------------------------------

