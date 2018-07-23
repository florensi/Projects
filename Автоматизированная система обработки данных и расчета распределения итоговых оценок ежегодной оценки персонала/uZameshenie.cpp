//---------------------------------------------------------------------------

#include <vcl.h>
#pragma hdrstop

#include "uZameshenie.h"
#include "uDM.h"
#include "uMain.h"
//---------------------------------------------------------------------------
#pragma package(smart_init)
#pragma link "DBGridEh"
#pragma resource "*.dfm"
TZameshenie *Zameshenie;
//---------------------------------------------------------------------------
__fastcall TZameshenie::TZameshenie(TComponent* Owner)
        : TForm(Owner)
{
}
//---------------------------------------------------------------------------
void __fastcall TZameshenie::BitBtn2Click(TObject *Sender)
{
  Zameshenie->Close();
}
//---------------------------------------------------------------------------

void __fastcall TZameshenie::FormShow(TObject *Sender)
{
  //���������� Edit-��

  //��, ���, ���
  EditTN->Text="490";
  EditZEX->Text="";
  EditZEX->Text=zzex=DM->qOcenka->FieldByName("zex")->AsString;
  EditTN->Text=ztn=DM->qOcenka->FieldByName("tn")->AsString;
  LabelFIO->Caption=zfio=DM->qOcenka->FieldByName("fio")->AsString;


  //����. �������
  EditVZ_PENS->Text=zvz_pens=DM->qOcenka->FieldByName("vz_pens")->AsString;

   //���
  EditKPE1->Text=zkpe1=DM->qOcenka->FieldByName("kpe1")->AsString;
  EditKPE2->Text=zkpe2=DM->qOcenka->FieldByName("kpe2")->AsString;
  EditKPE3->Text=zkpe3=DM->qOcenka->FieldByName("kpe3")->AsString;
  EditKPE4->Text=zkpe4=DM->qOcenka->FieldByName("kpe4")->AsString;

  if (DM->qOcenka->FieldByName("kpe1")->AsFloat+DM->qOcenka->FieldByName("kpe2")->AsFloat+
      DM->qOcenka->FieldByName("kpe3")->AsFloat+DM->qOcenka->FieldByName("kpe4")->AsFloat==0)
    {
      LabelKPE->Caption="";
    }
  else
    {
      LabelKPE->Caption=FloatToStrF((DM->qOcenka->FieldByName("kpe1")->AsFloat+DM->qOcenka->FieldByName("kpe2")->AsFloat+
      DM->qOcenka->FieldByName("kpe3")->AsFloat+DM->qOcenka->FieldByName("kpe4")->AsFloat)/4, ffFixed, 2,2)+ " %";
    }

      
 //��������� ���������

  //���������
  if (DM->qOcenka->FieldByName("rezerv")->AsString=="1")
    {
      CheckBoxREZERV->Checked=true;
      zrezerv=1;
      CheckBoxREZERV->Enabled = true;
    }
  else
    {
      CheckBoxREZERV->Checked=false;
      zrezerv="NULL";
    }
  //����������
  if (DM->qOcenka->FieldByName("zam")->AsString=="1")
    {
      CheckBoxZAM->Checked=true;
      zzam=1;
      CheckBoxZAM->Enabled = true;
    }
  else
    {
      CheckBoxZAM->Checked=false;
      zzam="NULL";
    }



  if (DM->qOcenka->FieldByName("rezerv")->AsString!="1" && DM->qOcenka->FieldByName("zam")->AsString!="1")
    {
      CheckBoxREZERV->Enabled = true;
      CheckBoxZAM->Enabled = true;
      DM->qZamesh->Filtered = false;
      DM->qZamesh->Active = false;
      DBGridEh1->Enabled = false;
    }

  //������� �����
  Panel3->Visible = false;
  Panel1->Align=alClient;
  Zameshenie->Height=440;
  BitBtn1->Top=316;
  BitBtn2->Top=355;
  Bevel3->Height=385;

 /*
  //����� ���������� �� ���������� � ����������� �� ��������
  DM->qZamesh->Close();
  DM->qZamesh->Parameters->ParamByName("pgod")->Value=IntToStr(Main->god);
  DM->qZamesh->Parameters->ParamByName("ptype")->Value=IntToStr(type);
  DM->qZamesh->Parameters->ParamByName("ptn")->Value=DM->qOcenka->FieldByName("tn")->AsString;

  try
    {
      DM->qZamesh->Open();
    }
  catch(Exception &E)
    {
      Application->MessageBox(("���������� �������� ������ �� ��������� ���������� ��������� "+E.Message).c_str(),"������",
                               MB_OK+MB_ICONERROR);
      Abort();
    } */


  EditZEX->SetFocus();
}
//---------------------------------------------------------------------------

void __fastcall TZameshenie::CheckBoxZAMClick(TObject *Sender)
{
  int rec;

  if (CheckBoxZAM->Checked==true)
    {
      //�������� �� �������� �� �������� �����������
      if (DM->qOcenka->FieldByName("rezerv")->AsString=="1")
        {
          Application->MessageBox("��� ��� ������ �������� �������� �����������, \n�� �� ����� ������������ ���� � ����������!!!","��������������",
                                  MB_OK+MB_ICONINFORMATION);

          CheckBoxZAM->Checked = false;
          CheckBoxREZERV->Checked = true;
          CheckBoxREZERV->SetFocus();
          Abort();
        }


      CheckBoxREZERV->Enabled = false;
      CheckBoxZAM->Enabled = true;
      DBGridEh1->Enabled = true;

      //����� ���������� �� ���������� � ����������� �� ��������
      DM->qZamesh->Filtered = false;
      DM->qZamesh->Filter = " type=2 and tn="+DM->qOcenka->FieldByName("tn")->AsString;
      if (DM->qZamesh->Active==false) DM->qZamesh->Active=true;
      DM->qZamesh->Filtered = true;
    }
  else
    {
      if (!EditTN->Text.IsEmpty() && EditTN->Text!="490")
        {
          //����� ���������� �� ���������� � ����������� �� ��������
          DM->qZamesh->Filtered = false;
          DM->qZamesh->Filter = " type=2 and tn="+DM->qOcenka->FieldByName("tn")->AsString;
          if (DM->qZamesh->Active==false) DM->qZamesh->Active=true;
          DM->qZamesh->Filtered = true;

          if (DM->qZamesh->RecordCount>0)
            {
              if(Application->MessageBox("�� ������� ��������� ���������� ���������, ������� �� ��������. ������� ��� ��������� �� ���������?","��������������",
                                      MB_YESNO+MB_ICONWARNING)==ID_NO)
                {
                  CheckBoxZAM->Checked=true;
                  CheckBoxZAM->Enabled=true;
                  CheckBoxREZERV->Enabled = false;
                  DBGridEh1->Enabled = true;
                  Abort();
                }
              else
                {
                  rec=DM->qOcenka->RecNo;
                  //�������� ���� ���������� ����������
                  DM->qZamesh->First();
                  while (!DM->qZamesh->Eof)
                    {
                      DM->qObnovlenie->Close();
                      DM->qObnovlenie->SQL->Clear();
                      DM->qObnovlenie->SQL->Add("delete from ocenka_rez where rowid = chartorowid("+ QuotedStr(DM->qZamesh->FieldByName("rw")->AsString)+")");
                      try
                        {
                          DM->qObnovlenie->ExecSQL();
                        }
                      catch (Exception &E)
                        {
                          Application->MessageBox(("�������� ������ ��� �������� ���������� ��������� � ������� OCENKA_REZ "+E.Message).c_str(),"������",
                                                   MB_OK+MB_ICONERROR);
                          DM->qZamesh->Requery();
                          Main->InsertLog("�������� ������ ��� �������� ���������� ���������("+DM->qZamesh->FieldByName("id_shtat")->AsString+") �� ���������: ���.�='"+EditTN->Text+"' ���='"+LabelFIO->Caption+"'");
                          Abort();
                        }

                      DM->qZamesh->Next();
                    }
                 if (DM->qObnovlenie->RowsAffected>0)
                   {
                     //���������� �������� � ������� ���������
                     DM->qObnovlenie->Close();
                      DM->qObnovlenie->SQL->Clear();
                      DM->qObnovlenie->SQL->Add("update ocenka set zam=NULL where god="+IntToStr(Main->god)+" and tn="+EditTN->Text);
                      try
                        {
                          DM->qObnovlenie->ExecSQL();
                        }
                      catch (Exception &E)
                        {
                          Application->MessageBox(("�������� ������ ��� ���������� �������� ����������� � ������� OCENKA "+E.Message).c_str(),"������",
                                                   MB_OK+MB_ICONERROR);
                          DM->qZamesh->Requery();
                          //����
                          Main->InsertLog("�������� ������ ��� ���������� �������� ����������� �� ���������: ���.�='"+EditTN->Text+"' ���='"+LabelFIO->Caption+"'");
                          Abort();
                        }

                     DM->qOcenka->Requery();
                     DM->qOcenka->RecNo=rec;

                     if (CheckBoxZAM->Checked==true)
                       {
                         CheckBoxZAM->Checked=false;
                         CheckBoxREZERV->Enabled=true;
                       }

                   }
                 
                 //����
                 Main->InsertLog("�������� ���������� ���������� ��������� ������� �� ���������: ���.�='"+EditTN->Text+"' ���='"+LabelFIO->Caption+"'");
                 
                }
            }
        }

      CheckBoxREZERV->Enabled = true;
      DM->qZamesh->Filtered = false;
      DM->qZamesh->Active=false;
      DBGridEh1->Enabled = false;
    }
}
//---------------------------------------------------------------------------



void __fastcall TZameshenie::EditZEXKeyPress(TObject *Sender, char &Key)
{
  if (!(IsNumeric(Key)||Key=='\b')) Key=0;
}
//---------------------------------------------------------------------------

void __fastcall TZameshenie::EditFIO_RKeyPress(TObject *Sender, char &Key)
{
  if (IsNumeric(Key)) Key=0;        
}
//---------------------------------------------------------------------------


void __fastcall TZameshenie::EditKPE1KeyPress(TObject *Sender, char &Key)
{
  if (! (IsNumeric(Key) || Key=='.' || Key==',' || Key=='/' || Key=='\b') ) Key=0;
  if (Key==',' || Key=='/') Key='.';        
}
//---------------------------------------------------------------------------

//���������� ���������
void __fastcall TZameshenie::BitBtn1Click(TObject *Sender)
{
  AnsiString Sql, zam, preem, rezerv, Str;
  int rec;

  if (CheckBoxZAM->Checked) zam=1;
  else zam="NULL";
 // if (CheckBoxPREEM->Checked) preem=1;
 // else preem=0;
  if (CheckBoxREZERV->Checked) rezerv=1;
  else rezerv="NULL";



  //��������
  //���
  if (EditZEX->Text.IsEmpty())
    {
      Application->MessageBox("�� ������ ��� ���������!!!","��������������",
                               MB_OK+MB_ICONINFORMATION);
      EditZEX->SetFocus();
      Abort();
    }
  //���.�
  if (EditTN->Text.IsEmpty() || EditTN->Text=="490")
    {
      Application->MessageBox("�� ������ ��������� ����� ���������!!!","��������������",
                               MB_OK+MB_ICONINFORMATION);
      EditTN->SetFocus();
      Abort();
    }

  //���������� �� �������� � ����� ����� � ���������
  DM->qObnovlenie->Close();
  DM->qObnovlenie->SQL->Clear();
  DM->qObnovlenie->SQL->Add("select * from ocenka where zex="+EditZEX->Text+" and tn="+EditTN->Text);
  try
    {
      DM->qObnovlenie->Open();
    }
  catch(Exception &E)
    {
      Application->MessageBox(("�������� ������ ��� ��������� ������ �� ������� Ocenka" + E.Message).c_str(),"������",
                              MB_OK+MB_ICONERROR);
      Abort();
    }

  if (DM->qObnovlenie->RecordCount<=0)
    {
      Application->MessageBox("� ��������� �� ���������� ��������� � ��������� ����� � ��������� �������!!!","��������������",
                               MB_OK+MB_ICONINFORMATION);
      EditZEX->SetFocus();
      Abort();
    }

  //�������� ������� "���������"**********************************************
  //�������� ��������� � �� ������� ���������� ���������
  if (CheckBoxZAM->Checked==true && DM->qZamesh->RecordCount==0)
    {
      Application->MessageBox("�������� ���������, �� �� ������� �� ���� ���������!!!","��������������",
                               MB_OK+MB_ICONINFORMATION);
      CheckBoxZAM->SetFocus();
      Abort();
    }

  //�� �������� ��������� � ������� ���������� ���������
  if (CheckBoxZAM->Checked==false && CheckBoxREZERV->Checked==false)
    {
      //���� �� � ������� ��������� �� ����� ���������
      DM->qObnovlenie->Close();
      DM->qObnovlenie->SQL->Clear();
      DM->qObnovlenie->SQL->Add("select * from ocenka_rez where god="+IntToStr(Main->god)+" and tn="+EditTN->Text);
      try
        {
          DM->qObnovlenie->Open();
        }
      catch(Exception &E)
        {
          Application->MessageBox(("�������� ������ ��� ��������� ������ �� ������� Ocenka" + E.Message).c_str(),"������",
                                    MB_OK+MB_ICONERROR);
          Abort();
        }

      if (DM->qObnovlenie->RecordCount>0)
        {
          Application->MessageBox("�� �������� ���������, �� ���� ��������� �� ���������!!!","��������������",
                                   MB_OK+MB_ICONINFORMATION);
          CheckBoxREZERV->SetFocus();
          Abort();
        }
    }

  //����������
  if (EditZEX->Text!=zzex ||
      EditTN->Text!=ztn ||
      LabelFIO->Caption!=zfio ||
      EditVZ_PENS->Text!=zvz_pens ||
      EditKPE1->Text!=zkpe1 ||
      EditKPE2->Text!=zkpe2 ||
      EditKPE3->Text!=zkpe3 ||
      EditKPE4->Text!=zkpe4 ||
      zrezerv!=rezerv ||
      zzam!=zam
       )
    {
      Sql="update ocenka set  \
                             vz_pens="+Main->SetNull(EditVZ_PENS->Text)+",\
                             kpe1="+Main->SetNull(EditKPE1->Text)+",\
                             kpe2="+Main->SetNull(EditKPE2->Text)+",\
                             kpe3="+Main->SetNull(EditKPE3->Text)+",\
                             kpe4="+Main->SetNull(EditKPE4->Text)+",\
                             zam="+Main->SetNull(zam)+",\
                             rezerv="+Main->SetNull(rezerv)+"\
            where rowid = chartorowid("+ QuotedStr(DM->qOcenka->FieldByName("rw")->AsString)+")";

      DM->qObnovlenie->Close();
      DM->qObnovlenie->SQL->Clear();
      DM->qObnovlenie->SQL->Add(Sql);
      rec=DM->qOcenka->RecNo;
      try
        {
          DM->qObnovlenie->ExecSQL();
        }
      catch(Exception &E)
        {
          Application->MessageBox(("�������� ������ ��� ������� ���������� ������ � ������� �� ������ ��������� (OCENKA) "+E.Message).c_str(),"������",
                                  MB_OK+MB_ICONERROR);
          Abort();
        }

      //����������� �������
      DM->qOcenka->RecNo = rec;

      Str="���������� ��������� �� ��������� ���="+EditZEX->Text+" ���.�="+EditTN->Text+":";

      if (Main->SetNull(EditVZ_PENS->Text)!=Main->SetNull(zvz_pens)) Str+=" ����.������� � '"+Main->SetNull(zvz_pens)+"' �� '"+Main->SetNull(EditVZ_PENS->Text)+"',";
      if (Main->SetNull(EditKPE1->Text)!=Main->SetNull(zkpe1)) Str+=", ��� �� 1��. � '"+Main->SetNull(zkpe1)+"' �� '"+Main->SetNull(EditKPE1->Text)+"'";
      if (Main->SetNull(EditKPE2->Text)!=Main->SetNull(zkpe2)) Str+=", ��� �� 2��. � '"+Main->SetNull(zkpe2)+"' �� '"+Main->SetNull(EditKPE2->Text)+"'";
      if (Main->SetNull(EditKPE3->Text)!=Main->SetNull(zkpe3)) Str+=", ��� �� 3��. � '"+Main->SetNull(zkpe3)+"' �� '"+Main->SetNull(EditKPE3->Text)+"'";
      if (Main->SetNull(EditKPE4->Text)!=Main->SetNull(zkpe4)) Str+=", ��� �� 1��. � '"+Main->SetNull(zkpe4)+"' �� '"+Main->SetNull(EditKPE4->Text)+"'";
      if (Main->SetNull(zzam)!=Main->SetNull(zam)) Str+=", ��������� � '"+Main->SetNull(zzam)+"' �� '"+Main->SetNull(zam)+"'";
      if (zrezerv!=rezerv) Str+=", ������ � '"+zrezerv+"' �� '"+rezerv+"'";
      Str+=" ���������";
      Main->InsertLog(Str);
      DM->qLogs->Requery();

    }

  Application->MessageBox("������ ������� ���������!","��������������",
                               MB_OK+MB_ICONINFORMATION);
  DM->qOcenka->Requery();

//  Zameshenie->Close();

  //�������� Edit-��
  EditTN->Text="490";
  EditZEX->Text="";

  LabelFIO->Caption="";
  EditVZ_PENS->Text="";
  ComboBoxGOTOV->ItemIndex=-1;

  EditKPE1->Text="";
  EditKPE2->Text="";
  EditKPE3->Text="";
  EditKPE4->Text="";
  LabelKPE->Caption="";

  CheckBoxZAM->Checked=false;
  CheckBoxREZERV->Checked=false;
  CheckBoxZAM->Enabled=true;
  CheckBoxREZERV->Enabled=true;
  ComboBoxGOTOV->Color=clWindow;
  ComboBoxRISK->Color=clWindow;

  //������ ������ ��������������
  Panel3->Visible = false;
  Panel1->Align=alClient;
  Zameshenie->Height=440;
  BitBtn1->Top=316;
  BitBtn2->Top=355;
  Bevel3->Height=385;


  EditZEX->SetFocus();

}
//---------------------------------------------------------------------------

void __fastcall TZameshenie::CheckBox1Click(TObject *Sender)
{
 /* if (CheckBox1->Checked==true)
    {
      if (EditDATN1->Text.IsEmpty())
        {
          EditDATN1->Text=DM->qOcenka->FieldByName("datn1")->AsString;
          EditDATK1->Text=DM->qOcenka->FieldByName("datk1")->AsString;
        }
    }
  else
    {
      EditDATN1->Text="";
      EditDATK1->Text="";
    } */
}
//---------------------------------------------------------------------------

void __fastcall TZameshenie::CheckBox2Click(TObject *Sender)
{
  /*if (CheckBox2->Checked==true)
    {
      if (EditDATN2->Text.IsEmpty())
        {
          EditDATN2->Text=DM->qOcenka->FieldByName("datn2")->AsString;
          EditDATK2->Text=DM->qOcenka->FieldByName("datk2")->AsString;
        }  
    }
  else
    {
      EditDATN2->Text="";
      EditDATK2->Text="";
    } */
}
//---------------------------------------------------------------------------

void __fastcall TZameshenie::CheckBox3Click(TObject *Sender)
{
  /*if (CheckBox3->Checked==true)
    {
      if (EditDATN3->Text.IsEmpty())
        {
          EditDATN3->Text=DM->qOcenka->FieldByName("datn3")->AsString;
          EditDATK3->Text=DM->qOcenka->FieldByName("datk3")->AsString;
        }
    }
  else
    {
      EditDATN3->Text="";
      EditDATK3->Text="";
    } */
}
//---------------------------------------------------------------------------

void __fastcall TZameshenie::CheckBox4Click(TObject *Sender)
{
  /*if (CheckBox4->Checked==true)
    {
      if (EditDATN4->Text.IsEmpty())
        {
          EditDATN4->Text=DM->qOcenka->FieldByName("datn4")->AsString;
          EditDATK4->Text=DM->qOcenka->FieldByName("datk4")->AsString;
        }
    }
  else
    {
      EditDATN4->Text="";
      EditDATK4->Text="";
    } */
}
//---------------------------------------------------------------------------

void __fastcall TZameshenie::CheckBox5Click(TObject *Sender)
{
  /*if (CheckBox5->Checked==true)
    {
      if (EditDATN5->Text.IsEmpty())
        {
          EditDATN5->Text=DM->qOcenka->FieldByName("datn5")->AsString;
          EditDATK5->Text=DM->qOcenka->FieldByName("datk5")->AsString;
        }
    }
  else
    {
      EditDATN5->Text="";
      EditDATK5->Text="";
    }*/
}
//---------------------------------------------------------------------------

void __fastcall TZameshenie::CheckBox6Click(TObject *Sender)
{
 /* if (CheckBox6->Checked==true)
    {
      if (EditDATN6->Text.IsEmpty())
        {
          EditDATN6->Text=DM->qOcenka->FieldByName("datn6")->AsString;
          EditDATK6->Text=DM->qOcenka->FieldByName("datk6")->AsString;
        }
    }
  else
    {
      EditDATN6->Text="";
      EditDATK6->Text="";
    } */
}
//---------------------------------------------------------------------------

void __fastcall TZameshenie::CheckBox7Click(TObject *Sender)
{
 /* if (CheckBox7->Checked==true)
    {
      if (EditDATN7->Text.IsEmpty())
        {
          EditDATN7->Text=DM->qOcenka->FieldByName("datn7")->AsString;
          EditDATK7->Text=DM->qOcenka->FieldByName("datk7")->AsString;
        }
    }
  else
    {
      EditDATN7->Text="";
      EditDATK7->Text="";
    }  */
}
//---------------------------------------------------------------------------

void __fastcall TZameshenie::CheckBox8Click(TObject *Sender)
{
  /*if (CheckBox8->Checked==true)
    {
      if (EditDATN8->Text.IsEmpty())
        {
          EditDATN8->Text=DM->qOcenka->FieldByName("datn8")->AsString;
          EditDATK8->Text=DM->qOcenka->FieldByName("datk8")->AsString;
        }
    }
  else
    {
      EditDATN8->Text="";
      EditDATK8->Text="";
    }*/
}
//---------------------------------------------------------------------------

void __fastcall TZameshenie::CheckBox9Click(TObject *Sender)
{
  /*if (CheckBox9->Checked==true)
    {
      if (EditDATN9->Text.IsEmpty())
        {
          EditDATN9->Text=DM->qOcenka->FieldByName("datn9")->AsString;
          EditDATK9->Text=DM->qOcenka->FieldByName("datk9")->AsString;
        }
    }
  else
    {
      EditDATN9->Text="";
      EditDATK9->Text="";
    } */
}
//---------------------------------------------------------------------------

void __fastcall TZameshenie::CheckBox10Click(TObject *Sender)
{
  /*if (CheckBox10->Checked==true)
    {
      if (EditDATN10->Text.IsEmpty())
        {
          EditDATN10->Text=DM->qOcenka->FieldByName("datn10")->AsString;
          EditDATK10->Text=DM->qOcenka->FieldByName("datk10")->AsString;
        }
    }
  else
    {
      EditDATN10->Text="";
      EditDATK10->Text="";
    } */
}
//---------------------------------------------------------------------------

void __fastcall TZameshenie::CheckBox11Click(TObject *Sender)
{
  /*if (CheckBox11->Checked==true)
    {
      if (EditDATN11->Text.IsEmpty())
        {
          EditDATN11->Text=DM->qOcenka->FieldByName("datn11")->AsString;
          EditDATK11->Text=DM->qOcenka->FieldByName("datk11")->AsString;
        }
    }
  else
    {
      EditDATN11->Text="";
      EditDATK11->Text="";
    } */
}
//---------------------------------------------------------------------------

void __fastcall TZameshenie::CheckBox12Click(TObject *Sender)
{
  /*if (CheckBox12->Checked==true)
    {
      if (EditDATN12->Text.IsEmpty())
        {
          EditDATN12->Text=DM->qOcenka->FieldByName("datn12")->AsString;
          EditDATK12->Text=DM->qOcenka->FieldByName("datk12")->AsString;
        }
    }
  else
    {
      EditDATN12->Text="";
      EditDATK12->Text="";
    }*/
}
//---------------------------------------------------------------------------

//��������� ���.� ���������
void __fastcall TZameshenie::EditTNChange(TObject *Sender)
{
  //���� �������� ���������
  TLocateOptions SearchOptions;

  if (!EditZEX->Text.IsEmpty())
    {
      Variant locvalues[] = {Main->SetNull(EditZEX->Text), Main->SetNull(EditTN->Text)};

      if (!DM->qOcenka->Locate("zex;tn", VarArrayOf(locvalues, 1),
                                         SearchOptions << loCaseInsensitive) )
        {
          //�� ������� ��������

          //�������� Edit-��
          LabelFIO->Caption="";
          EditVZ_PENS->Text="";
          ComboBoxGOTOV->ItemIndex=-1;
          ComboBoxGOTOV->Color=clWindow;

          EditKPE1->Text="";
          EditKPE2->Text="";
          EditKPE3->Text="";
          EditKPE4->Text="";

          CheckBoxZAM->Checked=false;
          CheckBoxREZERV->Checked=false;
          CheckBoxREZERV->Enabled = false;
          CheckBoxZAM->Enabled = false;

          DM->qZamesh->Filtered = false;
          DM->qZamesh->Active = false;
          DBGridEh1->Enabled = false;

          //������� �����
          Panel3->Visible = false;
          Panel1->Align=alClient;
          Zameshenie->Height=440;
          BitBtn1->Top=316;
          BitBtn2->Top=355;
          Bevel3->Height=385;

          Abort();
        }
      else
        {
          //���������� Edit-��
          //��, ���, ���
          EditZEX->Text=zzex=DM->qOcenka->FieldByName("zex")->AsString;
          EditTN->Text=ztn=DM->qOcenka->FieldByName("tn")->AsString;
          LabelFIO->Caption=zfio=DM->qOcenka->FieldByName("fio")->AsString;


          //����. ������� � ����������
          EditVZ_PENS->Text=zvz_pens=DM->qOcenka->FieldByName("vz_pens")->AsString;
          if (DM->qOcenka->FieldByName("gotov")->AsString==1)
            {
              ComboBoxGOTOV->ItemIndex=ComboBoxGOTOV->Items->IndexOf("������");
              ComboBoxGOTOV->Color=(TColor)0x008080FF;
              zgotov=1;
            }
          else if (DM->qOcenka->FieldByName("gotov")->AsString==2)
            {
              ComboBoxGOTOV->ItemIndex=ComboBoxGOTOV->Items->IndexOf("�������");
              ComboBoxGOTOV->Color=(TColor)0x0080FFFF;
              zgotov=2;
            }
          else if (DM->qOcenka->FieldByName("gotov")->AsString==3)
            {
              ComboBoxGOTOV->ItemIndex=ComboBoxGOTOV->Items->IndexOf("�������");
              ComboBoxGOTOV->Color=clMoneyGreen;
              zgotov=3;
            }
          else
            {
              ComboBoxGOTOV->ItemIndex=-1;
              ComboBoxGOTOV->Color=clWindow;
              zgotov="NULL";
            }


          //���
          EditKPE1->Text=zkpe1=DM->qOcenka->FieldByName("kpe1")->AsString;
          EditKPE2->Text=zkpe2=DM->qOcenka->FieldByName("kpe2")->AsString;
          EditKPE3->Text=zkpe3=DM->qOcenka->FieldByName("kpe3")->AsString;
          EditKPE4->Text=zkpe4=DM->qOcenka->FieldByName("kpe4")->AsString;

          if (DM->qOcenka->FieldByName("kpe1")->AsFloat+DM->qOcenka->FieldByName("kpe2")->AsFloat+
              DM->qOcenka->FieldByName("kpe3")->AsFloat+DM->qOcenka->FieldByName("kpe4")->AsFloat==0)
            {
              LabelKPE->Caption="";
            }
          else
            {
              LabelKPE->Caption=FloatToStrF((DM->qOcenka->FieldByName("kpe1")->AsFloat+DM->qOcenka->FieldByName("kpe2")->AsFloat+
              DM->qOcenka->FieldByName("kpe3")->AsFloat+DM->qOcenka->FieldByName("kpe4")->AsFloat)/4, ffFixed, 2,2)+ " %";
            }


          //��������� ���������
          //����������
          if (DM->qOcenka->FieldByName("zam")->AsString=="1")
            {
              CheckBoxZAM->Checked=true;
              zzam=1;
            }
          else
            {
              CheckBoxZAM->Checked=false;
              zzam="NULL";
            }

          //���������
          if (DM->qOcenka->FieldByName("rezerv")->AsString=="1")
            {
              CheckBoxREZERV->Checked=true;
              zrezerv=1;
            }
          else
            {
              CheckBoxREZERV->Checked=false;
              zrezerv="NULL";
            }

          if (DM->qOcenka->FieldByName("rezerv")->AsString!="1" && DM->qOcenka->FieldByName("zam")->AsString!="1")
            {
              CheckBoxREZERV->Enabled = true;
              CheckBoxZAM->Enabled = true;
              DM->qZamesh->Filtered = false;
              DM->qZamesh->Active = false;
              DBGridEh1->Enabled = false;
            }

          //������� �����
          Panel3->Visible = false;
          Panel1->Align=alClient;
          Zameshenie->Height=440;
          BitBtn1->Top=316;
          BitBtn2->Top=355;
          Bevel3->Height=385;

        }
    }
}
//---------------------------------------------------------------------------

//��������� ���� ��� ��������� ������
void __fastcall TZameshenie::EditZEXChange(TObject *Sender)
{

  if (!EditTN->Text.IsEmpty() && EditTN->Text!="490")
    {
      TLocateOptions SearchOptions;

      Variant locvalues[] = {Main->SetNull(EditZEX->Text), Main->SetNull(EditTN->Text)};

      if (!DM->qOcenka->Locate("zex;tn", VarArrayOf(locvalues, 1),
                                         SearchOptions << loCaseInsensitive) )
        {
          //�� ������� ��������

          //�������� Edit-��
          LabelFIO->Caption="";
          EditVZ_PENS->Text="";
          ComboBoxGOTOV->ItemIndex=-1;
          ComboBoxGOTOV->Color=clWindow;

          EditKPE1->Text="";
          EditKPE2->Text="";
          EditKPE3->Text="";
          EditKPE4->Text="";

          CheckBoxZAM->Checked=false;
          CheckBoxREZERV->Checked=false;
          CheckBoxREZERV->Enabled = false;
          CheckBoxZAM->Enabled = false;

          DM->qZamesh->Filtered = false;
          DM->qZamesh->Active = false;
          DBGridEh1->Enabled = false;

          //������� �����
          Panel3->Visible = false;
          Panel1->Align=alClient;
          Zameshenie->Height=440;
          BitBtn1->Top=316;
          BitBtn2->Top=355;
          Bevel3->Height=385;

          Abort();
        }
      else
        {
          //���������� Edit-��
          //��, ���, ���
          EditZEX->Text=zzex=DM->qOcenka->FieldByName("zex")->AsString;
          EditTN->Text=ztn=DM->qOcenka->FieldByName("tn")->AsString;
          LabelFIO->Caption=zfio=DM->qOcenka->FieldByName("fio")->AsString;


          //����. ������� � ����������
          EditVZ_PENS->Text=zvz_pens=DM->qOcenka->FieldByName("vz_pens")->AsString;
          if (DM->qOcenka->FieldByName("gotov")->AsString==1)
            {
              ComboBoxGOTOV->ItemIndex=ComboBoxGOTOV->Items->IndexOf("������");
              ComboBoxGOTOV->Color=(TColor)0x008080FF;
              zgotov=1;
            }
          else if (DM->qOcenka->FieldByName("gotov")->AsString==2)
            {
              ComboBoxGOTOV->ItemIndex=ComboBoxGOTOV->Items->IndexOf("�������");
              ComboBoxGOTOV->Color=(TColor)0x0080FFFF;
              zgotov=2;
            }
          else if (DM->qOcenka->FieldByName("gotov")->AsString==3)
            {
              ComboBoxGOTOV->ItemIndex=ComboBoxGOTOV->Items->IndexOf("�������");
              ComboBoxGOTOV->Color=clMoneyGreen;
              zgotov=3;
            }
          else
            {
              ComboBoxGOTOV->ItemIndex=-1;
              ComboBoxGOTOV->Color=clWindow;
              zgotov="NULL";
            }


          //���
          EditKPE1->Text=zkpe1=DM->qOcenka->FieldByName("kpe1")->AsString;
          EditKPE2->Text=zkpe2=DM->qOcenka->FieldByName("kpe2")->AsString;
          EditKPE3->Text=zkpe3=DM->qOcenka->FieldByName("kpe3")->AsString;
          EditKPE4->Text=zkpe4=DM->qOcenka->FieldByName("kpe4")->AsString;

          if (DM->qOcenka->FieldByName("kpe1")->AsFloat+DM->qOcenka->FieldByName("kpe2")->AsFloat+
              DM->qOcenka->FieldByName("kpe3")->AsFloat+DM->qOcenka->FieldByName("kpe4")->AsFloat==0)
            {
              LabelKPE->Caption="";
            }
          else
            {
              LabelKPE->Caption=FloatToStrF((DM->qOcenka->FieldByName("kpe1")->AsFloat+DM->qOcenka->FieldByName("kpe2")->AsFloat+
              DM->qOcenka->FieldByName("kpe3")->AsFloat+DM->qOcenka->FieldByName("kpe4")->AsFloat)/4, ffFixed, 2,2)+ " %";
            }


          //��������� ���������
          //����������
          if (DM->qOcenka->FieldByName("zam")->AsString=="1")
            {
              CheckBoxZAM->Checked=true;
              zzam=1;
            }
          else
            {
              CheckBoxZAM->Checked=false;
              zzam="NULL";
            }

          //���������
          if (DM->qOcenka->FieldByName("rezerv")->AsString=="1")
            {
              CheckBoxREZERV->Checked=true;
              zrezerv=1;
            }
          else
            {
              CheckBoxREZERV->Checked=false;
              zrezerv="NULL";
            }

          if (DM->qOcenka->FieldByName("rezerv")->AsString!="1" && DM->qOcenka->FieldByName("zam")->AsString!="1")
            {
              CheckBoxREZERV->Enabled = true;
              CheckBoxZAM->Enabled = true;
              DM->qZamesh->Filtered = false;
              DM->qZamesh->Active = false;
              DBGridEh1->Enabled = false;
            }

          //������� �����
          Panel3->Visible = false;
          Panel1->Align=alClient;
          Zameshenie->Height=440;
          BitBtn1->Top=316;
          BitBtn2->Top=355;
          Bevel3->Height=385;

        }
    }
}
//---------------------------------------------------------------------------

void __fastcall TZameshenie::FormCreate(TObject *Sender)
{
  EditZEX->Text="";
  EditTN->Text="";
  StringGrid1->Cells[0][0]="c";
  StringGrid1->Cells[1][0]="��";
}
//---------------------------------------------------------------------------

void __fastcall TZameshenie::ComboBoxGOTOVChange(TObject *Sender)
{
  if (ComboBoxGOTOV->ItemIndex==0) ComboBoxGOTOV->Color=(TColor)0x008080FF;
  else if (ComboBoxGOTOV->ItemIndex==1) ComboBoxGOTOV->Color=(TColor)0x0080FFFF;
  else if (ComboBoxGOTOV->ItemIndex==2) ComboBoxGOTOV->Color=clMoneyGreen;
  else ComboBoxGOTOV->Color=clWindow;
}
//---------------------------------------------------------------------------

void __fastcall TZameshenie::ComboBoxRISKChange(TObject *Sender)
{
  if (ComboBoxRISK->ItemIndex==0) ComboBoxRISK->Color=clMoneyGreen;
  else if (ComboBoxRISK->ItemIndex==1) ComboBoxRISK->Color=(TColor)0x0080FFFF;
  else if (ComboBoxRISK->ItemIndex==2) ComboBoxRISK->Color=(TColor)0x008080FF;
  else ComboBoxRISK->Color=clWindow;
}
//---------------------------------------------------------------------------

void __fastcall TZameshenie::EditZEX_ZAMChange(TObject *Sender)
{
  //����� ���� � �������
  if (!EditZEX_ZAM->Text.IsEmpty())
    {
      DM->qDolg->Filtered = false;
      TLocateOptions SearchOptions;
      DM->qDolg->Locate("shifr_zex", EditZEX_ZAM->Text,
                            SearchOptions << loCaseInsensitive);

      if (EditZEX_ZAM->Text.Length()>2) LabelZEX_ZAM->Caption = DM->qDolg->FieldByName("nzex")->AsString +"\n"+ DM->qDolg->FieldByName("uch")->AsString;
      else LabelZEX_ZAM->Caption = DM->qDolg->FieldByName("nzex")->AsString;

      if (!EditSHIFR_ZAM->Text.IsEmpty())
        {
          //��� ���������
          //�������� �������� �� ��������� ���
          DM->qObnovlenie->Close();
          DM->qObnovlenie->SQL->Clear();
          DM->qObnovlenie->SQL->Add("select * from sp_ocenka_krd \
                                     where zex="+QuotedStr(EditZEX_ZAM->Text)+"\
                                     and shifr_dolg=(select short from p1000@sapmig_buffdb where otype='S' and langu='R' and objid="+QuotedStr(EditSHIFR_ZAM->Text)+")");
          try
            {
              DM->qObnovlenie->Open();
            }
          catch(Exception &E)
            {
              Application->MessageBox(("���������� �������� ������ �� ����������� ��� (SP_OCENKA_KRD)"+E.Message).c_str(),"������",
                                        MB_OK+MB_ICONERROR);
            }

          if (DM->qObnovlenie->RecordCount>0) LabelKRD->Caption="�������� ��������� ���������";
          else LabelKRD->Caption="";
        }
    }
  else LabelZEX_ZAM->Caption =""; 
}
//---------------------------------------------------------------------------

void __fastcall TZameshenie::EditSHIFR_ZAMChange(TObject *Sender)
{
  AnsiString Sql;

  if (!EditSHIFR_ZAM->Text.IsEmpty())
    {
      Sql=" select vse.zec, vse.nzec as nzec, vse.dolg as dolg, vse.id_dolg as id_dolg, \
                   ruk.tn_sap as tn_sap, (ruk.fam||' '||ruk.im||' '||ruk.ot) as fio                                 \
            from                                                                                                 \
                (select prof.*, cex.short as zec, cex.stext as nzec                                              \
                 from                                                                                            \
                     (select pr.objid as shtat, pr.short as id_dolg, pr.stext as dolg, s.sobid as sobid          \
                      from p1000@sapmig_buffdb pr                                                                \
                      left join                                                                                  \
                      p1001@sapmig_buffdb s on pr.objid=s.objid                                                  \
                      where pr.otype='S' and pr.langu='R' and s.otype='S' and s.sclas='O' and pr.objid="+QuotedStr(EditSHIFR_ZAM->Text)+"\
                      ) prof                                                                                     \
                     left join                                                                                   \
                     p1000@sapmig_buffdb cex                                                                     \
                     on prof.sobid=cex.objid and cex.otype='O' and cex.langu='R'                                 \
                     ) vse                                                                                       \
                 left join sap_osn_sved ruk                                                                      \
                 on vse.shtat=ruk.id_shtat";

      //Sql="select stext from p1000@sapmig_buffdb where otype='S' and langu='R' and objid=:pkod_prof";


      DM->qObnovlenie->Close();
      DM->qObnovlenie->SQL->Clear();
      DM->qObnovlenie->SQL->Add(Sql);
      try
        {
          DM->qObnovlenie->Open();
        }
      catch(Exception &E)
        {
          Application->MessageBox(("�������� ������ ��� ������� ������� � ����������� ���������� (P1000)"+ E.Message).c_str(),"������",
                                  MB_OK+MB_ICONERROR);
        }


     // LabelZEX_ZAM->Caption=DM->qObnovlenie->FieldByName("nzec")->AsString;
      LabelDOLG_ZAM->Caption=DM->qObnovlenie->FieldByName("dolg")->AsString;
      LabelTN_R->Caption=DM->qObnovlenie->FieldByName("tn_sap")->AsString;
      EditFIO_R->Text=DM->qObnovlenie->FieldByName("fio")->AsString;
      id_dolg=DM->qObnovlenie->FieldByName("id_dolg")->AsString;
      EditZEX_ZAM->Text=DM->qObnovlenie->FieldByName("zec")->AsString;


  if (!LabelTN_R->Caption.IsEmpty() && ComboBoxRISK->ItemIndex==-1)
    {
      Sql="select risk, risk_prich from ocenka where tn="+LabelTN_R->Caption+" and god="+IntToStr(Main->god);

      DM->qObnovlenie->Close();
      DM->qObnovlenie->SQL->Clear();
      DM->qObnovlenie->SQL->Add(Sql);
      try
        {
          DM->qObnovlenie->Open();
        }
      catch(Exception &E)
        {
          Application->MessageBox(("�������� ������ ��� ������� ������ �� ��������� �� ������ ��������� (Ocenka)"+ E.Message).c_str(),"������",
                                    MB_OK+MB_ICONERROR);
        }

      //���� � �������
      if (DM->qObnovlenie->FieldByName("risk")->AsString==1)
        {
          ComboBoxRISK->ItemIndex=ComboBoxRISK->Items->IndexOf("������");
          ComboBoxRISK->Color=clMoneyGreen;
        }
      else if (DM->qObnovlenie->FieldByName("risk")->AsString==2)
        {
          ComboBoxRISK->ItemIndex=ComboBoxRISK->Items->IndexOf("�������");
          ComboBoxRISK->Color=(TColor)0x0080FFFF;
        }
      else if (DM->qObnovlenie->FieldByName("risk")->AsString==3)
        {
          ComboBoxRISK->ItemIndex=ComboBoxRISK->Items->IndexOf("�������");
          ComboBoxRISK->Color=(TColor)0x008080FF;
        }
      else
        {
          ComboBoxRISK->ItemIndex=-1;
          ComboBoxRISK->Color=clWindow;
        }
      EditRISK_PRICH->Text=DM->qObnovlenie->FieldByName("risk_prich")->AsString;
    }

    if (!EditZEX_ZAM->Text.IsEmpty())
      {
        //��� ���������
        //�������� �������� �� ��������� ���
        DM->qObnovlenie->Close();
        DM->qObnovlenie->SQL->Clear();
        DM->qObnovlenie->SQL->Add("select * from sp_ocenka_krd \
                                   where zex="+QuotedStr(EditZEX_ZAM->Text)+"\
                                   and shifr_dolg=(select short from p1000@sapmig_buffdb where otype='S' and langu='R' and objid="+QuotedStr(EditSHIFR_ZAM->Text)+")");
        try
          {
            DM->qObnovlenie->Open();
          }
        catch(Exception &E)
          {
            Application->MessageBox(("���������� �������� ������ �� ����������� ��� (SP_OCENKA_KRD)"+E.Message).c_str(),"������",
                                     MB_OK+MB_ICONERROR);
          }

        if (DM->qObnovlenie->RecordCount>0) LabelKRD->Caption="�������� ��������� ���������";
        else LabelKRD->Caption="";
      }
    }
  else
    {
      LabelDOLG_ZAM->Caption="";
    }
}
//---------------------------------------------------------------------------

void __fastcall TZameshenie::EditFIO_RChange(TObject *Sender)
{
/*  if (!EditFIO_R->Text.IsEmpty())
    {
      DM->qObnovlenie->Close();
      DM->qObnovlenie->SQL->Clear();
      DM->qObnovlenie->SQL->Add("select tn_sap from sap_work \
                                 where upper(regexp_replace("+QuotedStr(EditFIO_R->Text)+", ' (.*)'))=upper(fam) \
                                 and upper(regexp_replace("+QuotedStr(EditFIO_R->Text)+", ' (.*)|^[^ ]* '))=upper(im) \
                                 and upper(regexp_replace("+QuotedStr(EditFIO_R->Text)+", '(.*) '))=upper(ot)");
      DM->qObnovlenie->Open();


      LabelTN_R->Caption=DM->qObnovlenie->FieldByName("tn_sap")->AsString;
    }*/        
}
//---------------------------------------------------------------------------

void __fastcall TZameshenie::EditZEXKeyDown(TObject *Sender, WORD &Key,
      TShiftState Shift)
{
  if (Key==VK_RETURN)
  FindNextControl((TWinControl *)Sender, true, true, false)->SetFocus();
  EditTN->SelStart=EditTN->Text.Length();
}
//---------------------------------------------------------------------------

void __fastcall TZameshenie::StringGrid1DrawCell(TObject *Sender, int ACol,
      int ARow, TRect &Rect, TGridDrawState State)
{
  int x,y;
  TDateTime d;


  x = Rect.left+(Rect.Width() - StringGrid1->Canvas->TextWidth(StringGrid1->Cells[ACol][ARow]))/2;
  y = Rect.top+(Rect.Height() - StringGrid1->Canvas->TextHeight(StringGrid1->Cells[ACol][ARow]))/2;


  // ��������� ������ �������� ������
  if (State.Contains(gdSelected))
    {
      StringGrid1->Canvas->Brush->Color =TColor(0x00C8F7E3);// clInfoBk;
      StringGrid1->Canvas->Font->Color = clBlack;
      StringGrid1->Canvas->FillRect(Rect);
      StringGrid1->Canvas->TextOut(x,y,StringGrid1->Cells[ACol][ARow]);
    }

  // ������ ��������� �������� ������ ��� ������ �� StringGrid1
  if (ActiveControl != StringGrid1)
    {
      if (State.Contains(gdSelected))
        {
          StringGrid1->Canvas->Brush->Color = clWhite;
          StringGrid1->Canvas->Font->Color = clBlack;
          StringGrid1->Canvas->FillRect(Rect);
          StringGrid1->Canvas->TextOut(x,y,StringGrid1->Cells[ACol][ARow]);
        }
    }

  if (ARow==0)
    {
      StringGrid1->Canvas->Font->Color=clBlack;
      StringGrid1->Canvas->Font->Style=TFontStyles()<<fsBold;
      //DrawText(StringGrid1->Canvas->Handle, StringGrid1->Cells[ACol][ARow].c_str(), strlen(StringGrid1->Cells[ACol][ARow].c_str()), &Rect, DT_WORDBREAK); // ������� ����� � ������ ��������� �-��� WinAPI
    }

  //��������� �����, ���� �� ����� ������� ����
  // �������� �� ������������ ����� ����

/*      if  (ARow!=0) {
          if(!TryStrToDate(StringGrid1->Cells[ARow][ACol],d) )
            {

              StringGrid1->Canvas->Brush->Color =TColor(0x00C8F7E3);// clInfoBk;
      StringGrid1->Canvas->Font->Color = clRed;
      StringGrid1->Canvas->FillRect(Rect);
      StringGrid1->Canvas->TextOut(x,y,StringGrid1->Cells[ACol][ARow]);



          //    StringGrid1->Canvas->Brush->Color = clRed;
              //StringGrid1->Canvas->Font->Color=clRed;//   TFontStyles()<<fsBold;

  // StringGrid1.Canvas.Font.Color := clWhite;


        //       StringGrid1->Canvas->TextOut(Rect.Left, Rect.Top, StringGrid1->Cells[ACol][ ARow]);

            }
          else
            {
              StringGrid1->Cells[ARow][ACol]=FormatDateTime("dd.mm.yyyy",d);
           //   StringGrid1->Canvas->Brush->Color = clBlack;
             // StringGrid1->Canvas->Font->Color=clBlack;


             StringGrid1->Canvas->Brush->Color =TColor(0x00C8F7E3);// clInfoBk;
      StringGrid1->Canvas->Font->Color = clBlack;
      StringGrid1->Canvas->FillRect(Rect);
      StringGrid1->Canvas->TextOut(x,y,StringGrid1->Cells[ACol][ARow]);
            }

          }   */


 /*  int i=1;

*/


  if(!TryStrToDate(StringGrid1->Cells[StringGrid1->Col][StringGrid1->Row],d) && ARow!=0 && StringGrid1->Cells[StringGrid1->Col][StringGrid1->Row]!="" )
    {
      StringGrid1->Canvas->Font->Color = clRed;
      StringGrid1->Canvas->FillRect(Rect);
      StringGrid1->Canvas->TextOut(x,y,StringGrid1->Cells[StringGrid1->Col][StringGrid1->Row]);
    }
  else
    {
      StringGrid1->Canvas->Font->Color = clBlack;
      StringGrid1->Canvas->FillRect(Rect);
      StringGrid1->Canvas->TextOut(x,y,StringGrid1->Cells[ACol][ARow]);
    }

  StringGrid1->Canvas->Brush->Color = clGreen;


 /* StringGrid1->Canvas->FillRect(Rect);
  StringGrid1->Canvas->TextOut(x,y,StringGrid1->Cells[ACol][ARow]);  */


  // ����� ���������� ������
  /*  StringGrid1->Canvas->FillRect(Rect);
      DrawText(StringGrid1->Canvas->Handle, StringGrid1->Cells[ACol][ARow].c_str(), strlen(StringGrid1->Cells[ACol][ARow].c_str()),
      &Rect, DT_WORDBREAK); // ������� ����� � ������ ��������� �-��� WinAPI
  */


}
//---------------------------------------------------------------------------

void __fastcall TZameshenie::StringGrid1SetEditText(TObject *Sender,
      int ACol, int ARow, const AnsiString Value)
{
/* TDateTime d;

  if (ActiveControl == BitBtn2)
    {
      Zameshenie->Close();
    }
  else
    {
      //int i= StringGrid1->Col;
      //int j= StringGrid1->Row;
      if (StringGrid1->Cells[ACol][ARow]!="")
        {

          // ���������� � ���� ��������� ������ � ����
          if (StringGrid1->Cells[ACol][ARow].Length()<3)
            {
              if(StringGrid1->Cells[ACol][ARow].Pos("."))
                {
                  Application->MessageBox("�������� ������ ����","������", MB_OK+MB_ICONINFORMATION);
         //         EditDATN1->Font->Color = clRed;
                   StringGrid1->Canvas->Brush->Color =clRed;

                  StringGrid1->Row=ARow;
                  StringGrid1->Col=ACol;
                  Abort();
                }
              else
                {
                  StringGrid1->Cells[ACol][ARow] = StringGrid1->Cells[ACol][ARow]+ "."+ DateToStr(Date()).SubString(4,2) +"."+ DateToStr(Date()).SubString(7,5);
                  //EditDATN1->Font->Color = clBlack;
                }
            }

          // �������� �� ������������ ����� ����
          if(!TryStrToDate(StringGrid1->Cells[ACol][ARow],d))
            {
              Application->MessageBox("�������� ������ ����","������", MB_OK);
          //    EditDATN1->Font->Color = clRed;
              StringGrid1->Row=ARow;
              StringGrid1->Col=ACol;
            }
          else
            {
              StringGrid1->Cells[ACol][ARow]=FormatDateTime("dd.mm.yyyy",d);
            //  EditDATN1->Font->Color = clBlack;
            }

        }
    }   */
}
//---------------------------------------------------------------------------



void __fastcall TZameshenie::StringGrid2SetEditText(TObject *Sender,
      int ACol, int ARow, const AnsiString Value)
{
 TDateTime d;

  if (ActiveControl == BitBtn2)
    {
      Zameshenie->Close();
    }
  else
    {
      //int i= StringGrid1->Col;
      //int j= StringGrid1->Row;
      if (StringGrid1->Cells[ACol][ARow]!="")
        {

          // ���������� � ���� ��������� ������ � ����
          if (StringGrid1->Cells[ACol][ARow].Length()<3)
            {
              if(StringGrid1->Cells[ACol][ARow].Pos("."))
                {
                  Application->MessageBox("�������� ������ ����","������", MB_OK+MB_ICONINFORMATION);
         //         EditDATN1->Font->Color = clRed;
                   StringGrid1->Canvas->Brush->Color =clRed;

                  StringGrid1->Row=ARow;
                  StringGrid1->Col=ACol;
                  Abort();
                }
              else
                {
                  StringGrid1->Cells[ACol][ARow] = StringGrid1->Cells[ACol][ARow]+ "."+ DateToStr(Date()).SubString(4,2) +"."+ DateToStr(Date()).SubString(7,5);
                  //EditDATN1->Font->Color = clBlack;
                }
            }

          // �������� �� ������������ ����� ����
          if(!TryStrToDate(StringGrid1->Cells[ACol][ARow],d))
            {
              Application->MessageBox("�������� ������ ����","������", MB_OK);
          //    EditDATN1->Font->Color = clRed;
              StringGrid1->Row=ARow;
              StringGrid1->Col=ACol;
            }
          else
            {
              StringGrid1->Cells[ACol][ARow]=FormatDateTime("dd.mm.yyyy",d);
            //  EditDATN1->Font->Color = clBlack;
            }

        }
    }
}
//---------------------------------------------------------------------------

void __fastcall TZameshenie::StringGrid1SelectCell(TObject *Sender,
      int ACol, int ARow, bool &CanSelect)
{
   TDateTime d;

if (ActiveControl == BitBtn2)
    {
      Zameshenie->Close();
    }
  else
    {
      int i= StringGrid1->Col;
      int j= StringGrid1->Row;
      if (StringGrid1->Cells[i][j]!="")
        {

          // ���������� � ���� ��������� ������ � ����
          if (StringGrid1->Cells[i][j].Length()<3)
            {
              if(StringGrid1->Cells[i][j].Pos("."))
                {
                  Application->MessageBox("�������� ������ ����","������", MB_OK+MB_ICONINFORMATION);
         //         EditDATN1->Font->Color = clRed;
                   StringGrid1->Canvas->Brush->Color =clRed;

                  StringGrid1->Row=StringGrid1->Row;
                  StringGrid1->Col=StringGrid1->Col;
               //   StringGrid1->Options << goEditing;
             //     CanSelect = true;
                  Abort();
                }
              else
                {
                  StringGrid1->Cells[i][j] = StringGrid1->Cells[i][j]+ "."+ DateToStr(Date()).SubString(4,2) +"."+ DateToStr(Date()).SubString(7,5);
                  //EditDATN1->Font->Color = clBlack;
                }
            }

          // �������� �� ������������ ����� ����
          if(!TryStrToDate(StringGrid1->Cells[i][j],d))
            {
             Application->MessageBox("�������� ������ ����","������", MB_OK+MB_ICONINFORMATION);
             // Application->MessageBox("�������� ������ ����","������", MB_OK);
          //    EditDATN1->Font->Color = clRed;

              //StringGrid1->Cells[i][j];

              StringGrid1->Row=StringGrid1->Row;
              StringGrid1->Col=StringGrid1->Col;
              Abort();
          //    StringGrid1->Options << goEditing;
            //  CanSelect = true;



            }
          else
            {
              StringGrid1->Cells[i][j]=FormatDateTime("dd.mm.yyyy",d);
            //  EditDATN1->Font->Color = clBlack;
            }

        }
    }

}
//---------------------------------------------------------------------------

void __fastcall TZameshenie::StringGrid1Enter(TObject *Sender)
{
  StringGrid1->Col = 0;
  StringGrid1->Row = 1;
}
//---------------------------------------------------------------------------




void __fastcall TZameshenie::StringGrid1Exit(TObject *Sender)
{
   TDateTime d;

if (ActiveControl == BitBtn2)
    {
      Zameshenie->Close();
    }
  else
    {
      int i= StringGrid1->Col;
      int j= StringGrid1->Row;
      if (StringGrid1->Cells[i][j]!="")
        {

          // ���������� � ���� ��������� ������ � ����
          if (StringGrid1->Cells[i][j].Length()<3)
            {
              if(StringGrid1->Cells[i][j].Pos("."))
                {
                  Application->MessageBox("�������� ������ ����","������", MB_OK+MB_ICONINFORMATION);
         //         EditDATN1->Font->Color = clRed;
                   StringGrid1->Canvas->Brush->Color =clRed;

                  StringGrid1->Row=StringGrid1->Row;
                  StringGrid1->Col=StringGrid1->Col;
                  StringGrid1->SetFocus();
               //   StringGrid1->Options << goEditing;
             //     CanSelect = true;
                  Abort();
                }
              else
                {
                  StringGrid1->Cells[i][j] = StringGrid1->Cells[i][j]+ "."+ DateToStr(Date()).SubString(4,2) +"."+ DateToStr(Date()).SubString(7,5);
                  //EditDATN1->Font->Color = clBlack;
                }
            }

          // �������� �� ������������ ����� ����
          if(!TryStrToDate(StringGrid1->Cells[i][j],d))
            {
           //   Application->MessageBox("�������� ������ ����","������", MB_OK);
          //    EditDATN1->Font->Color = clRed;

              //StringGrid1->Cells[i][j];

              StringGrid1->Row=StringGrid1->Row;
              StringGrid1->Col=StringGrid1->Col;

              StringGrid1->SetFocus();
              Abort();
          //    StringGrid1->Options << goEditing;
            //  CanSelect = true;



            }
          else
            {
              StringGrid1->Cells[i][j]=FormatDateTime("dd.mm.yyyy",d);
            //  EditDATN1->Font->Color = clBlack;
            }

        }
    }
      StringGrid1->Invalidate();
}
//---------------------------------------------------------------------------


void __fastcall TZameshenie::StringGrid1KeyPress(TObject *Sender,
      char &Key)
{
  if (!(IsNumeric(Key) || Key=='\b' || Key=='.' || Key==',' || Key=='/')) Key=0;
  if (Key==',' || Key=='/') Key='.';
}
//---------------------------------------------------------------------------

void __fastcall TZameshenie::StringGrid1KeyDown(TObject *Sender, WORD &Key,
      TShiftState Shift)
{

  //��� ������� Enter
  if (Key==VK_RETURN)
    {
      if (StringGrid1->Col==0)
        {
          StringGrid1->Col=1;
        }
      else if (StringGrid1->Col==1 && StringGrid1->Row < kol_str1)
        {
          StringGrid1->Col=0;
          StringGrid1->Row = StringGrid1->Row +1;
        }
      else
        {
          //if (StringGrid2->Cells[0][1]=="") BitBtn1->SetFocus();
          FindNextControl((TWinControl *)Sender, true, true, false)->SetFocus();
        }
    }


    /*
      if (StringGrid1->Row < Main->kol_str1)
        {
          StringGrid1->Row = StringGrid1->Row +1;

        }
      else if (StringGrid1->Col < StringGrid1->ColCount-1)
        {
          StringGrid1->Col = StringGrid1->Col+1;
          StringGrid1->Row = 1;
        }
      else
        {
          FindNextControl((TWinControl *)Sender, true, true, false)->SetFocus();
        }
    }   */



 /* //���, ���, ������, ������� �����, ������� ������, ������� �����, ������� ����, Enter                                              ������� ������      ������� �����   ������� ����
if (Key==VK_LBUTTON || Key==VK_RBUTTON || Key==VK_SPACE || Key==VK_LEFT ||
     Key==VK_RIGHT || Key==VK_UP || Key==VK_DOWN)// || Key==VK_RETURN)

   {

     t_r = StringGrid1->Row;
     t_k = StringGrid1->Col;

     // �������� �� ������������ ����� ����
          if(!TryStrToDate(StringGrid1->Cells[StringGrid1->Col][StringGrid1->Row],d))
            {
              StringGrid1->Row=StringGrid1->Row-1;
              StringGrid1->Col=StringGrid1->Col;

              StringGrid1->Options << goEditing;
            //  CanSelect = true;
            }


   }



 /*  //��� ������� Enter
  if (Key==VK_RETURN)
    {




 if (ActiveControl == BitBtn2)
    {
      Zameshenie->Close();
    }
  else
    {
      int i= StringGrid1->Col;
      int j= StringGrid1->Row;
      if (StringGrid1->Cells[i][j]!="")
        {
          // ���������� � ���� ��������� ������ � ����
          if (StringGrid1->Cells[i][j].Length()<3)
            {
              if(StringGrid1->Cells[i][j].Pos("."))
                {
                  Application->MessageBox("�������� ������ ����","������", MB_OK+MB_ICONINFORMATION);

                  StringGrid1->Row=j;
                  StringGrid1->Col=i;
                  StringGrid1->Options << goEditing;
                  //CanSelect = true;
                  Abort();
                }
              else
                {
                  StringGrid1->Cells[i][j] = StringGrid1->Cells[i][j]+ "."+ DateToStr(Date()).SubString(4,2) +"."+ DateToStr(Date()).SubString(7,5);
                  //EditDATN1->Font->Color = clBlack;
                }
            }

          // �������� �� ������������ ����� ����
          if(!TryStrToDate(StringGrid1->Cells[i][j],d))
            {
              Application->MessageBox("�������� ������ ����","������", MB_OK);
              //CanSelect = true;
              StringGrid1->Row=j;
              StringGrid1->Col=i;

              StringGrid1->Options << goEditing;
            TGridRect g;

 g.Bottom=StringGrid1->Row;
 g.Top=StringGrid1->Row;
 g.Left=StringGrid1->Col;
 g.Right=StringGrid1->Col;
 StringGrid1->Selection=g;

             // StringGrid1->TopRow = j;
             // StringGrid1->Selection = TGridRect(Rect(i, j, i, j));

              
            }
          else
            {
              StringGrid1->Cells[i][j]=FormatDateTime("dd.mm.yyyy",d);
            }

        }
    }

     










      if (StringGrid1->Row < Main->kol_str1)
        {
          StringGrid1->Row = StringGrid1->Row +1;

        }
      else if (StringGrid1->Col < StringGrid1->ColCount-1)
        {
          StringGrid1->Col = StringGrid1->Col+1;
          StringGrid1->Row = 1;
        }
      else
        {
          FindNextControl((TWinControl *)Sender, true, true, false)->SetFocus();
        }
    }   */
}
//---------------------------------------------------------------------------

void __fastcall TZameshenie::CheckBoxREZERVClick(TObject *Sender)
{
  int rec;

  if (CheckBoxREZERV->Checked==true)
    {
      CheckBoxZAM->Enabled = false;
      CheckBoxREZERV->Enabled = true;
      DBGridEh1->Enabled = true;

      //����� ���������� �� ���������� � ����������� �� ��������
      DM->qZamesh->Filtered = false;
      DM->qZamesh->Filter = " type=1 and tn="+DM->qOcenka->FieldByName("tn")->AsString;
      if (DM->qZamesh->Active==false) DM->qZamesh->Active=true;
      DM->qZamesh->Filtered = true;

    }
  else
    {
      if (!EditTN->Text.IsEmpty() && EditTN->Text!="490")
        {
          //����� ���������� �� ���������� � ����������� �� ��������
          DM->qZamesh->Filtered = false;
          DM->qZamesh->Filter = " type=1 and tn="+DM->qOcenka->FieldByName("tn")->AsString;
          if (DM->qZamesh->Active==false) DM->qZamesh->Active=true;
          DM->qZamesh->Filtered = true;

          if (DM->qZamesh->RecordCount>0)
            {
              if(Application->MessageBox("�� ������� ��������� ���������� ���������, ������� �� ��������. ������� ��� ��������� �� ���������?","��������������",
                                      MB_YESNO+MB_ICONWARNING)==ID_NO)
                {
                  CheckBoxREZERV->Checked=true;
                  Abort();
                }
              else
                {
                  rec=DM->qOcenka->RecNo;
                  //�������� ���� ���������� ����������
                  DM->qZamesh->First();
                  while (!DM->qZamesh->Eof)
                    {
                      DM->qObnovlenie->Close();
                      DM->qObnovlenie->SQL->Clear();
                      DM->qObnovlenie->SQL->Add("delete from ocenka_rez where rowid = chartorowid("+ QuotedStr(DM->qZamesh->FieldByName("rw")->AsString)+")");
                      try
                        {
                          DM->qObnovlenie->ExecSQL();
                        }
                      catch (Exception &E)
                        {
                          Application->MessageBox(("�������� ������ ��� �������� ���������� ��������� � ������� OCENKA_REZ "+E.Message).c_str(),"������",
                                                   MB_OK+MB_ICONERROR);
                          DM->qZamesh->Requery();
                          //����
                          Main->InsertLog("�������� ������ ��� �������� ��������� ���������("+DM->qZamesh->FieldByName("id_shtat")->AsString+") �� ���������: ���.�='"+EditTN->Text+"' ���='"+LabelFIO->Caption+"'");
                          Abort();
                        }

                      DM->qZamesh->Next();
                    }
                 if (DM->qObnovlenie->RowsAffected>0)
                   {
                     //���������� �������� � ������� ���������
                     DM->qObnovlenie->Close();
                      DM->qObnovlenie->SQL->Clear();
                      DM->qObnovlenie->SQL->Add("update ocenka set rezerv=NULL where god="+IntToStr(Main->god)+" and tn="+EditTN->Text);
                      try
                        {
                          DM->qObnovlenie->ExecSQL();
                        }
                      catch (Exception &E)
                        {
                          Application->MessageBox(("�������� ������ ��� ���������� �������� ���������� � ������� OCENKA "+E.Message).c_str(),"������",
                                                   MB_OK+MB_ICONERROR);
                          DM->qZamesh->Requery();
                          //����
                          Main->InsertLog("�������� ������ ��� ���������� �������� ���������� �� ���������: ���.�='"+EditTN->Text+"' ���='"+LabelFIO->Caption+"'");
                          Abort();
                        }
                   }
                  DM->qOcenka->Requery();
                  DM->qOcenka->RecNo=rec;

                  //����
                  Main->InsertLog("�������� ���������� ���������� ��������� ������� �� ���������: ���.�='"+EditTN->Text+"' ���='"+LabelFIO->Caption+"'");
                }
            }
        }

      CheckBoxZAM->Enabled = true;
      DM->qZamesh->Filtered = false;
      DM->qZamesh->Active=false;
      DBGridEh1->Enabled = false;
    }
}
//---------------------------------------------------------------------------

void __fastcall TZameshenie::FormClose(TObject *Sender,
      TCloseAction &Action)
{
  DM->qZamesh->Filtered = false;
}
//---------------------------------------------------------------------------

//���������� ��������� ��� ��������� (���������� - ���������, ����������� - ����������)
void __fastcall TZameshenie::N1DobavClick(TObject *Sender)
{
  fl_red=0;
  Panel3->Visible = true;
  Panel1->Align=alClient;
  Zameshenie->Height=705;
  BitBtn1->Top=579;
  BitBtn2->Top=618;
  Bevel3->Height=651;
  if (DM->qOcenka->FieldByName("rezerv")->AsString=="1") GroupBoxZAM->Caption="���������� ��������� ���������";
  else GroupBoxZAM->Caption="���������� ���������";
  BitBtn3->Caption = "��������";


  //������� ������
  //���
  EditZEX_ZAM->Text="";
  LabelZEX_ZAM->Caption ="";

  //����
  EditSHIFR_ZAM->Text="";
  LabelDOLG_ZAM->Caption="";

  //�� �����������
  LabelTN_R->Caption="";

  //��� ���������
  LabelKRD->Caption="";

  //��� �����������
   EditFIO_R->Text="";

  //���� � �������
  ComboBoxRISK->ItemIndex=-1;
  ComboBoxRISK->Color=clWindow;
  EditRISK_PRICH->Text="";


  //����������
  ComboBoxGOTOV->ItemIndex=-1;
  ComboBoxGOTOV->Color=clWindow;


  //������� ���������
  //���� ������ ���������
  StringGrid1->Cells[0][1]="";
  StringGrid1->Cells[0][2]="";
  StringGrid1->Cells[0][3]="";
  StringGrid1->Cells[0][4]="";
  StringGrid1->Cells[0][5]="";
  StringGrid1->Cells[0][6]="";
  //���� ����� ���������
  StringGrid1->Cells[1][1]="";
  StringGrid1->Cells[1][2]="";
  StringGrid1->Cells[1][3]="";
  StringGrid1->Cells[1][4]="";
  StringGrid1->Cells[1][5]="";
  StringGrid1->Cells[1][6]="";
}
//---------------------------------------------------------------------------

//�������������� ���������
void __fastcall TZameshenie::N2RedakClick(TObject *Sender)
{
  if (DM->qZamesh->RecordCount==0)
    {
      fl_red=0;
      N1DobavClick(Sender);
    }
  else
    {
      fl_red=1;

      Panel1->Align=alClient;
      Zameshenie->Height=705;
      BitBtn1->Top=579;
      BitBtn2->Top=618;
      Bevel3->Height=651;
      GroupBoxZAM->Caption="�������������� ���������";
      if (DM->qOcenka->FieldByName("rezerv")->AsString=="1") GroupBoxZAM->Caption="�������������� ��������� ���������";
      else GroupBoxZAM->Caption="�������������� ���������";
      BitBtn3->Caption = "�������������";

      ZapolnenieInfo();

      Panel3->Visible = true;
   }
 /* if (DM->qOcenka->FieldByName("preem")->AsString=="1")
    {
      CheckBoxPREEM->Checked=true;
      zpreem=1;
    }
  else
    {
      CheckBoxPREEM->Checked=false;
      zpreem="NULL";
    } */

}
//---------------------------------------------------------------------------

//���������� ������
void __fastcall TZameshenie::ZapolnenieInfo()
{
  //���
  EditZEX_ZAM->Text=zzex_zam=DM->qZamesh->FieldByName("zex_rez")->AsString;
  //����� ���� � �������
  if (!EditZEX_ZAM->Text.IsEmpty())
    {
      DM->qDolg->Filtered = false;
      TLocateOptions SearchOptions;
      DM->qDolg->Locate("shifr_zex", EditZEX_ZAM->Text,
                            SearchOptions << loCaseInsensitive);

      if (EditZEX_ZAM->Text.Length()>2) LabelZEX_ZAM->Caption = DM->qDolg->FieldByName("nzex")->AsString +"\n"+ DM->qDolg->FieldByName("uch")->AsString;
      else LabelZEX_ZAM->Caption = DM->qDolg->FieldByName("nzex")->AsString;
    }
  else LabelZEX_ZAM->Caption ="";


  //����
  EditSHIFR_ZAM->Text=zshifr_zam=DM->qZamesh->FieldByName("id_shtat")->AsString;


  //�� �����������
  LabelTN_R->Caption=ztn_r=DM->qZamesh->FieldByName("tn_sap_rez")->AsString;


  //��� ���������
  //�������� �������� �� ��������� ���
  DM->qObnovlenie->Close();
  DM->qObnovlenie->SQL->Clear();
  DM->qObnovlenie->SQL->Add("select * from sp_ocenka_krd \
                             where zex="+QuotedStr(EditZEX_ZAM->Text)+"\
                             and shifr_dolg=(select short from p1000@sapmig_buffdb where otype='S' and langu='R' and objid="+QuotedStr(EditSHIFR_ZAM->Text)+")");
  try
    {
      DM->qObnovlenie->Open();
    }
  catch(Exception &E)
    {
      Application->MessageBox(("���������� �������� ������ �� ����������� ��� (SP_OCENKA_KRD)"+E.Message).c_str(),"������",
                               MB_OK+MB_ICONERROR);
    }

  if (DM->qObnovlenie->RecordCount>0) LabelKRD->Caption="�������� ��������� ���������";
  else LabelKRD->Caption="";


  //��� �����������
   EditFIO_R->Text=zfio_r=DM->qZamesh->FieldByName("fio_rez")->AsString;


  //���� � �������
  if (DM->qZamesh->FieldByName("risk")->AsString==1)
    {
      ComboBoxRISK->ItemIndex=ComboBoxRISK->Items->IndexOf("������");
      ComboBoxRISK->Color=clMoneyGreen;
      zrisk=1;
    }
  else if (DM->qZamesh->FieldByName("risk")->AsString==2)
    {
      ComboBoxRISK->ItemIndex=ComboBoxRISK->Items->IndexOf("�������");
      ComboBoxRISK->Color=(TColor)0x0080FFFF;
      zrisk=2;
    }
  else if (DM->qZamesh->FieldByName("risk")->AsString==3)
    {
      ComboBoxRISK->ItemIndex=ComboBoxRISK->Items->IndexOf("�������");
      ComboBoxRISK->Color=(TColor)0x008080FF;
      zrisk=3;
    }
  else
    {
      ComboBoxRISK->ItemIndex=-1;
      ComboBoxRISK->Color=clWindow;
      zrisk="NULL";
    }
  EditRISK_PRICH->Text=zrisk_prich=DM->qZamesh->FieldByName("risk_prich")->AsString;


  //����������
  if (DM->qZamesh->FieldByName("gotov")->AsString==1)
    {
      ComboBoxGOTOV->ItemIndex=ComboBoxGOTOV->Items->IndexOf("������");
      ComboBoxGOTOV->Color=(TColor)0x008080FF;
      zgotov=1;
    }
  else if (DM->qZamesh->FieldByName("gotov")->AsString==2)
    {
      ComboBoxGOTOV->ItemIndex=ComboBoxGOTOV->Items->IndexOf("�������");
      ComboBoxGOTOV->Color=(TColor)0x0080FFFF;
      zgotov=2;
    }
  else if (DM->qZamesh->FieldByName("gotov")->AsString==3)
    {
      ComboBoxGOTOV->ItemIndex=ComboBoxGOTOV->Items->IndexOf("�������");
      ComboBoxGOTOV->Color=clMoneyGreen;
      zgotov=3;
    }
  else
    {
      ComboBoxGOTOV->ItemIndex=-1;
      ComboBoxGOTOV->Color=clWindow;
      zgotov="NULL";
    }


  //������� ���������
  //���� ������ ���������
  StringGrid1->Cells[0][1]=zdatn1=DM->qZamesh->FieldByName("datn1")->AsString;
  StringGrid1->Cells[0][2]=zdatn2=DM->qZamesh->FieldByName("datn2")->AsString;
  StringGrid1->Cells[0][3]=zdatn3=DM->qZamesh->FieldByName("datn3")->AsString;
  StringGrid1->Cells[0][4]=zdatn4=DM->qZamesh->FieldByName("datn4")->AsString;
  StringGrid1->Cells[0][5]=zdatn5=DM->qZamesh->FieldByName("datn5")->AsString;
  StringGrid1->Cells[0][6]=zdatn6=DM->qZamesh->FieldByName("datn6")->AsString;
  //���� ����� ���������
  StringGrid1->Cells[1][1]=zdatk1=DM->qZamesh->FieldByName("datk1")->AsString;
  StringGrid1->Cells[1][2]=zdatk2=DM->qZamesh->FieldByName("datk2")->AsString;
  StringGrid1->Cells[1][3]=zdatk3=DM->qZamesh->FieldByName("datk3")->AsString;
  StringGrid1->Cells[1][4]=zdatk4=DM->qZamesh->FieldByName("datk4")->AsString;
  StringGrid1->Cells[1][5]=zdatk5=DM->qZamesh->FieldByName("datk5")->AsString;
  StringGrid1->Cells[1][6]=zdatk6=DM->qZamesh->FieldByName("datk6")->AsString;

  kol_str1=0;
  //���������� ����������� ����� � StringGrid1
  for (int i=1; i<7; i++)
    {
      if (!DM->qZamesh->FieldByName("datn"+IntToStr(i))->AsString.IsEmpty()) kol_str1++;
    }
}
//---------------------------------------------------------------------------
void __fastcall TZameshenie::BitBtn4Click(TObject *Sender)
{
  //������ ������ ��������������
  Panel3->Visible = false;
  Panel1->Align=alClient;
  Zameshenie->Height=440;
  BitBtn1->Top=316;
  BitBtn2->Top=355;
  Bevel3->Height=385;
}
//---------------------------------------------------------------------------

//����������/�������������� ���������
void __fastcall TZameshenie::BitBtn3Click(TObject *Sender)
{
  AnsiString risk, Sql, Str, zam,  gotov, rez;
  int type=0, rec1, rec2;

  //����
  if (ComboBoxRISK->Text=="������") risk=1;
  else if (ComboBoxRISK->Text=="�������") risk=2;
  else if (ComboBoxRISK->Text=="�������") risk=3;
  else risk="NULL";

    //����������
  if (ComboBoxGOTOV->Text=="������") gotov=1;
  else if (ComboBoxGOTOV->Text=="�������") gotov=2;
  else if (ComboBoxGOTOV->Text=="�������") gotov=3;
  else gotov="NULL";

  //���������/����������
  if (CheckBoxREZERV->Checked==true)
    {
      type=1;
      zam="NULL";
      rez=1;
    }
  else if (CheckBoxZAM->Checked==true)
    {
      type=2;
      zam=1;
      rez="NULL";
    }
  else
    {
      type=0;
      zam="NULL";
      rez="NULL";
    }

  //�������� ������� "���������"
  //���.�
  if (EditTN->Text.IsEmpty())
    {
      Application->MessageBox("�� ������ ���.� ���������!!!","��������������",
                               MB_OK+MB_ICONINFORMATION);
      EditTN->SetFocus();
      Abort();
    }

  //���
  if (EditZEX_ZAM->Text.IsEmpty())
    {
      Application->MessageBox("�� ������ ��� ����������� ���������!!!","��������������",
                               MB_OK+MB_ICONINFORMATION);
      EditZEX_ZAM->SetFocus();
      Abort();
    }

  //����
  if (EditSHIFR_ZAM->Text.IsEmpty())
    {
      Application->MessageBox("�� ������ ���� ���������� ���������!!!","��������������",
                               MB_OK+MB_ICONINFORMATION);
      EditSHIFR_ZAM->SetFocus();
      Abort();
    }

  //���������� ���������� ������ ��� ����������� ��������� ��� ������� ���
  if (!EditFIO_R->Text.IsEmpty() && LabelTN_R->Caption=="")
    {
      Application->MessageBox("���������� �������� ��������� ����� ����������� ���������.\n�������� ������� ������� ��� ����������� ��������� \n���� ���������� �������� ������.","��������������",
                                   MB_OK+MB_ICONWARNING);
      EditFIO_R->SetFocus();
      Abort();
    }

  //���������� ��� ����������� ��������� ��� ������� ���������� ������ ���
  if (EditFIO_R->Text.IsEmpty() && LabelTN_R->Caption!="")
    {
      Application->MessageBox("�� ������� ��� ����������� ���������","��������������",
                                   MB_OK+MB_ICONWARNING);
      EditFIO_R->SetFocus();
      Abort();
    }

  //������������� ���������
  if (!EditSHIFR_ZAM->Text.IsEmpty())
    {
      Sql="select stext from p1000@sapmig_buffdb where otype='S' and langu='R' and objid=:pkod_prof";


      DM->qObnovlenie->Close();
      DM->qObnovlenie->SQL->Clear();
      DM->qObnovlenie->SQL->Add(Sql);
      DM->qObnovlenie->Parameters->ParamByName("pkod_prof")->Value = Main->SetNull(EditSHIFR_ZAM->Text);
      try
        {
          DM->qObnovlenie->Open();
        }
      catch(Exception &E)
        {
          Application->MessageBox(("�������� ������ ��� ������� ������� � ����������� ���������� (P1000)"+ E.Message).c_str(),"������",
                                  MB_OK+MB_ICONERROR);
        }

      if (DM->qObnovlenie->RecordCount==0)
        {
          Application->MessageBox("��� ���������� ����� ��������� � ����������� ����������!!!","��������������",
                                   MB_OK+MB_ICONINFORMATION);
          EditSHIFR_ZAM->SetFocus();
          Abort();
        }
    }

  //������������� ����������� ��������� �� ����� ���������
  if (!EditFIO_R->Text.IsEmpty())
    {
      Sql="select tn_sap from \
                             (select case when ur1 is null then zex          \
                                          when ur2 is null then ur1          \
                                          when ur3 is null then ur2          \
                                          when ur4 is null then ur3 end ur,  \
                                     tn_sap,                                 \
                                     id_shtat                                 \
                              from sap_osn_sved)                             \
           where id_shtat=:pid_dolg and ur=:pur and tn_sap=:ptn_sap";


      DM->qObnovlenie->Close();
      DM->qObnovlenie->SQL->Clear();
      DM->qObnovlenie->SQL->Add(Sql);
      DM->qObnovlenie->Parameters->ParamByName("pid_dolg")->Value = Main->SetNull(EditSHIFR_ZAM->Text);
      DM->qObnovlenie->Parameters->ParamByName("pur")->Value = Main->SetNull(EditZEX_ZAM->Text);
      DM->qObnovlenie->Parameters->ParamByName("ptn_sap")->Value = Main->SetNull(LabelTN_R->Caption);
      try
        {
          DM->qObnovlenie->Open();
        }
      catch(Exception &E)
        {
          Application->MessageBox(("�������� ������ ��� ������� ������� � ��������� ���������� (SAP_OSN_SVED)"+ E.Message).c_str(),"������",
                                  MB_OK+MB_ICONERROR);
        }

      if (DM->qObnovlenie->RecordCount==0)
        {
          Application->MessageBox("�������������� ����� ���������� ��������� � ����������� ���������!!!\n�������� ������� ������ ���� �������������, \n���� ��������� ��� ��� ����������� ���������","��������������",
                                   MB_OK+MB_ICONINFORMATION);
          EditZEX_ZAM->SetFocus();
          Abort();
        }

    }


//�������� ������� "������� ���������"
//***************************************
  //���� ������ �� ���������, � ���� ����� ���������
  //StringGrid1
  for (int i=1; i<7; i++)
    {
      if (StringGrid1->Cells[0][i]=="" && StringGrid1->Cells[1][i]!="")
        {
          Application->MessageBox("������� ���� ����� ���������, �� �� ������� ���� ������ ���������!!!","��������������",
                              MB_OK+MB_ICONINFORMATION);
          StringGrid1->SetFocus();
          StringGrid1->Row=i;
          StringGrid1->Col=0;
          StringGrid1->EditorMode = true;

          Abort();
        }
    }

  //���� ������ ���������, � ���� ����� �� ���������
  //StringGrid1
  for (int i=1; i<7; i++)
    {
      if (StringGrid1->Cells[0][i]!="" && StringGrid1->Cells[1][i]=="")
        {
          Application->MessageBox("������� ���� ������ ���������, �� �� ������� ���� ����� ���������!!!","��������������",
                              MB_OK+MB_ICONINFORMATION);

          StringGrid1->SetFocus();
          StringGrid1->Row=i;
          StringGrid1->Col=1;
          StringGrid1->EditorMode = true;

          Abort();
        }
    }

  //���� ����� ������ ���� ������
  //StringGrid1
  for (int i=1; i<7; i++)
    {
      if (StringGrid1->Cells[0][i]!="" && StrToDate(StringGrid1->Cells[0][i])>StrToDate(StringGrid1->Cells[1][i]))
        {
          Application->MessageBox("���� ������ ��������� ������, ��� ���� ����� ���������!!!","��������������",
                                   MB_OK+MB_ICONINFORMATION);
          StringGrid1->SetFocus();
          StringGrid1->Row=i;
          StringGrid1->Col=0;
          StringGrid1->EditorMode = true;

          Abort();
        }
    }

  // �������� �� ���������� ������ ������
  //���� � StringGrid1
  for (int j=1; j<7; j++)
    {
      if (StringGrid1->Cells[0][j]!="")
        {
          for (int i=1; i<7; i++)
            {
              //�������� �������������� ��� �� StringGrid1 � StringGrid1
              if (i!=j && StringGrid1->Cells[0][i]!="")
                {
                  if (((StrToDate(StringGrid1->Cells[0][i]) < StrToDate(StringGrid1->Cells[0][j]) && StrToDate(StringGrid1->Cells[1][i]) > StrToDate(StringGrid1->Cells[0][j]))
                        || (StrToDate(StringGrid1->Cells[0][i]) > StrToDate(StringGrid1->Cells[0][j]) && StrToDate(StringGrid1->Cells[1][i]) > StrToDate(StringGrid1->Cells[0][j])
                            && (StrToDate(StringGrid1->Cells[0][i]) < StrToDate(StringGrid1->Cells[1][j]) || StrToDate(StringGrid1->Cells[0][i]) == StrToDate(StringGrid1->Cells[1][j])))
                        ||  (StrToDate(StringGrid1->Cells[0][i]) == StrToDate(StringGrid1->Cells[0][j]) || StrToDate(StringGrid1->Cells[1][i]) == StrToDate(StringGrid1->Cells[0][j])))
                      )
                    {
                      Application->MessageBox("�������� ������ ��������� ������������\n� ��� ������������","������",
                                               MB_OK + MB_ICONERROR);

                      StringGrid1->SetFocus();
                      StringGrid1->Row=j;
                      StringGrid1->Col=0;
                      StringGrid1->EditorMode = true;

                      Abort();
                    }
                }

              //�������� �������������� ��� �� StringGrid1 � StringGrid2
            /*  if (StringGrid2->Cells[0][i]!="")
                {
                    if (((StrToDate(StringGrid2->Cells[0][i]) < StrToDate(StringGrid1->Cells[0][j]) && StrToDate(StringGrid2->Cells[1][i]) > StrToDate(StringGrid1->Cells[0][j]))
                        || (StrToDate(StringGrid2->Cells[0][i]) > StrToDate(StringGrid1->Cells[0][j]) && StrToDate(StringGrid2->Cells[1][i]) > StrToDate(StringGrid1->Cells[0][j])
                            && (StrToDate(StringGrid2->Cells[0][i]) < StrToDate(StringGrid1->Cells[1][j]) || StrToDate(StringGrid2->Cells[0][i]) == StrToDate(StringGrid1->Cells[1][j])))
                        ||  (StrToDate(StringGrid2->Cells[0][i]) == StrToDate(StringGrid1->Cells[0][j]) || StrToDate(StringGrid2->Cells[1][i]) == StrToDate(StringGrid1->Cells[0][j])))
                      )
                    {
                      Application->MessageBox("�������� ������ ��������� ������������\n� ��� ������������","������",
                                               MB_OK + MB_ICONERROR);

                      StringGrid1->SetFocus();
                      StringGrid1->Row=j;
                      StringGrid1->Col=0;
                      StringGrid1->EditorMode = true;

                      Abort();
                    }
                }  */
            }
        }
    }

  //�������
  if (fl_red==0)
    {
      //�������� �� ��� ������� ����� �������� � ������� ���������
      DM->qObnovlenie->Close();
      DM->qObnovlenie->SQL->Clear();
      DM->qObnovlenie->SQL->Add("select * from ocenka_rez where god="+IntToStr(Main->god)+" and tn="+EditTN->Text+" and id_shtat="+QuotedStr(EditSHIFR_ZAM->Text));
      try
        {
          DM->qObnovlenie->Open();
        }
      catch(Exception &E)
        {
          Application->MessageBox(("�������� ������ ��� ��������� ������ �� ������� Ocenka_rez" + E.Message).c_str(),"������",
                                  MB_OK+MB_ICONERROR);
          Abort();
        }

      if (DM->qObnovlenie->RecordCount>0)
        {
          Application->MessageBox("� ������� ��������� ��� ���� ��������� �� ��������� ���������!!!","��������������",
                                   MB_OK+MB_ICONINFORMATION);
          EditZEX->SetFocus();
          Abort();
        }


      Sql="insert into ocenka_rez (god, tn, id_shtat, dolg_rez, tn_sap_rez, fio_rez, zex_rez, shifr_rez, type, risk, risk_prich, gotov, \
                                   datn1, datn2, datn3, datn4, datn5, datn6, datk1, datk2, datk3, datk4, datk5, datk6) \
           values ("+IntToStr(Main->god)+",                                                                                         \
                    "+EditTN->Text+",                                                                                                    \
                    "+QuotedStr(EditSHIFR_ZAM->Text)+",                                                                                             \
                    "+QuotedStr(LabelDOLG_ZAM->Caption)+",                                                                                          \
                    "+Main->SetNull(LabelTN_R->Caption)+",                                                                                              \
                    "+QuotedStr(EditFIO_R->Text)+",                                                                                                 \
                    "+QuotedStr(EditZEX_ZAM->Text)+",                                                                                               \
                    "+QuotedStr(id_dolg)+",  \
                    "+type+",                                                                                                        \
                    "+risk+",                                                                                                        \
                    "+QuotedStr(EditRISK_PRICH->Text)+",                                                                                         \
                    "+gotov+",                                           \
                     to_date("+QuotedStr(StringGrid1->Cells[0][1])+",'dd.mm.yyyy'),\
                     to_date("+QuotedStr(StringGrid1->Cells[0][2])+",'dd.mm.yyyy'),\
                     to_date("+QuotedStr(StringGrid1->Cells[0][3])+",'dd.mm.yyyy'),\
                     to_date("+QuotedStr(StringGrid1->Cells[0][4])+",'dd.mm.yyyy'),\
                     to_date("+QuotedStr(StringGrid1->Cells[0][5])+",'dd.mm.yyyy'),\
                     to_date("+QuotedStr(StringGrid1->Cells[0][6])+",'dd.mm.yyyy'),\
                     to_date("+QuotedStr(StringGrid1->Cells[1][1])+",'dd.mm.yyyy'),\
                     to_date("+QuotedStr(StringGrid1->Cells[1][2])+",'dd.mm.yyyy'),\
                     to_date("+QuotedStr(StringGrid1->Cells[1][3])+",'dd.mm.yyyy'),\
                     to_date("+QuotedStr(StringGrid1->Cells[1][4])+",'dd.mm.yyyy'),\
                     to_date("+QuotedStr(StringGrid1->Cells[1][5])+",'dd.mm.yyyy'),\
                     to_date("+QuotedStr(StringGrid1->Cells[1][6])+",'dd.mm.yyyy')\
                  )";
      rec1=DM->qOcenka->RecNo;
      rec2=DM->qZamesh->RecNo;

    }
  //����������
  else if (fl_red==1)
    {
      //�������� �� ������� ����������
     if (EditTN->Text!=ztn ||
         EditZEX_ZAM->Text!=zzex_zam ||
         EditSHIFR_ZAM->Text!=zshifr_zam ||
         EditFIO_R->Text!=zfio_r ||
         LabelTN_R->Caption!=ztn_r ||
         DM->qOcenka->FieldByName("rezerv")->AsString!=rez ||
         DM->qOcenka->FieldByName("zam")->AsString!=zam ||
         gotov!=zgotov ||
         risk!=zrisk ||
         EditRISK_PRICH->Text!=zrisk_prich ||
         StringGrid1->Cells[0][1]!=zdatn1 ||
         StringGrid1->Cells[0][2]!=zdatn2 ||
         StringGrid1->Cells[0][3]!=zdatn3 ||
         StringGrid1->Cells[0][4]!=zdatn4 ||
         StringGrid1->Cells[0][5]!=zdatn5 ||
         StringGrid1->Cells[0][6]!=zdatn6 ||
         StringGrid1->Cells[1][1]!=zdatk1 ||
         StringGrid1->Cells[1][2]!=zdatk2 ||
         StringGrid1->Cells[1][3]!=zdatk3 ||
         StringGrid1->Cells[1][4]!=zdatk4 ||
         StringGrid1->Cells[1][5]!=zdatk5 ||
         StringGrid1->Cells[1][6]!=zdatk6 )
       {
         Sql="update ocenka_rez set  \                                                                                                                  \
                             dolg_rez="+QuotedStr(LabelDOLG_ZAM->Caption)+",                                                                             \
                             tn_sap_rez="+Main->SetNull(LabelTN_R->Caption)+",\                                                                         \
                             fio_rez="+QuotedStr(EditFIO_R->Text)+",\                                                                                   \
                             zex_rez="+QuotedStr(EditZEX_ZAM->Text)+",\                                                                                 \
                             shifr_rez="+QuotedStr(id_dolg)+",  \
                             type="+type+",                                                                                                             \
                             risk="+risk+",                                                                                                             \
                             risk_prich="+QuotedStr(EditRISK_PRICH->Text)+",\
                             gotov="+gotov+",                               \
                             datn1=to_date("+QuotedStr(StringGrid1->Cells[0][1])+",'dd.mm.yyyy'),\
                             datn2=to_date("+QuotedStr(StringGrid1->Cells[0][2])+",'dd.mm.yyyy'),\
                             datn3=to_date("+QuotedStr(StringGrid1->Cells[0][3])+",'dd.mm.yyyy'),\
                             datn4=to_date("+QuotedStr(StringGrid1->Cells[0][4])+",'dd.mm.yyyy'),\
                             datn5=to_date("+QuotedStr(StringGrid1->Cells[0][5])+",'dd.mm.yyyy'),\
                             datn6=to_date("+QuotedStr(StringGrid1->Cells[0][6])+",'dd.mm.yyyy'),\
                             datk1=to_date("+QuotedStr(StringGrid1->Cells[1][1])+",'dd.mm.yyyy'),\
                             datk2=to_date("+QuotedStr(StringGrid1->Cells[1][2])+",'dd.mm.yyyy'),\
                             datk3=to_date("+QuotedStr(StringGrid1->Cells[1][3])+",'dd.mm.yyyy'),\
                             datk4=to_date("+QuotedStr(StringGrid1->Cells[1][4])+",'dd.mm.yyyy'),\
                             datk5=to_date("+QuotedStr(StringGrid1->Cells[1][5])+",'dd.mm.yyyy'),\
                             datk6=to_date("+QuotedStr(StringGrid1->Cells[1][6])+",'dd.mm.yyyy')\
            where rowid = chartorowid("+ QuotedStr(DM->qZamesh->FieldByName("rw")->AsString)+") and god="+IntToStr(Main->god)+" \
            and tn="+EditTN->Text+" and id_shtat="+QuotedStr(EditSHIFR_ZAM->Text);

          rec1=DM->qOcenka->RecNo;
          rec2=DM->qZamesh->RecNo;
        }
       else
         {
           //������� ������
           Panel3->Visible = false;
           Panel1->Align=alClient;
           Zameshenie->Height=440;
           BitBtn1->Top=316;
           BitBtn2->Top=355;
           Bevel3->Height=385;
           Abort();
         }
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
      Application->MessageBox(("�������� ������ ��� ���������� ������ � ������� Ocenka_rez" + E.Message).c_str(),"������",
                               MB_OK+MB_ICONERROR);
      Main->InsertLog("�������� ������ ��� ���������� ���.��������� � ������� OCENKA_REZ �� ��������� "+LabelTN_R->Caption);
      Abort();
    }


  //���� ��� ����������, �� �������� ���� � ������� ����� � ������������ � ������� Ocenka
   if (DM->qObnovlenie->RowsAffected>0 && !EditFIO_R->Text.IsEmpty()
       && (risk!=zrisk || EditRISK_PRICH->Text!=zrisk_prich ))
     {
       Sql="update ocenka set                                                \
                          risk="+risk+",                                     \
                          risk_prich="+QuotedStr(EditRISK_PRICH->Text)+"               \                                      \
            where god="+IntToStr(Main->god)+" and tn="+LabelTN_R->Caption;

       DM->qObnovlenie->Close();
       DM->qObnovlenie->SQL->Clear();
       DM->qObnovlenie->SQL->Add(Sql);
       try
         {
           DM->qObnovlenie->ExecSQL();
         }
       catch(Exception &E)
         {
           Application->MessageBox(("�������� ������ ��� ���������� ������ � ������� Ocenka" + E.Message).c_str(),"������",
                                   MB_OK+MB_ICONERROR);
           Main->InsertLog("�������� ������ ��� ���������� ����� � ������� ����� �� ����� ������������� ���.��������� � ������� OCENKA �� ��������� "+LabelTN_R->Caption);
           Abort();
         }
     }

   //���� ��� ����������, �� �������� ���� ���������/���������� � ��������� � ������� Ocenka
   if ((DM->qOcenka->FieldByName("rezerv")->AsString!=rez || DM->qOcenka->FieldByName("zam")->AsString!=zam)
        && DM->qObnovlenie->RowsAffected>0)
     {
       Sql = "update ocenka set \
                               rezerv="+rez+", \
                               zam="+zam+"     \
              where tn="+EditTN->Text+" and god="+IntToStr(Main->god);

       DM->qObnovlenie->Close();
       DM->qObnovlenie->SQL->Clear();
       DM->qObnovlenie->SQL->Add(Sql);
       try
         {
           DM->qObnovlenie->ExecSQL();
         }
       catch(Exception &E)
         {
           Application->MessageBox(("�������� ������ ��� ������� �������� ������ � ������� Ocenka" + E.Message).c_str(),"������",
                                    MB_OK+MB_ICONERROR);

           Main->InsertLog("�������� ������ ��� ���������� ������� ��������� �� ����� ������������� ���.��������� � ������� OCENKA �� ��������� "+EditTN->Text);
           DM->qLogs->Requery();
           DM->qOcenka->Requery();
           Main->StatusBar1->SimpleText ="�������� ������: "+IntToStr(Main->god)+" ���";
           Abort();
         }
     }
     

  //����
  if (DM->qObnovlenie->RowsAffected>0)
    {

      if (fl_red==0)
        {
          Str ="���������� ���������� ��������� �� "+IntToStr(Main->god)+" ��� �� ��������� '"+EditTN->Text+"': ";
          Str+="���.� �����������="+EditTN->Text+" ���="+LabelFIO->Caption+" ��� ����������.����.="+EditZEX_ZAM->Text;
          Str+="����="+EditSHIFR_ZAM->Text+", ��� �����������="+EditFIO_R->Text+", ���="+LabelTN_R->Caption+", ����="+risk+", �������="+EditRISK_PRICH->Text+", ����������="+gotov;
          if (StringGrid1->Cells[0][1]!="") Str+=", �����. c"+StringGrid1->Cells[0][1]+" �� "+StringGrid1->Cells[1][1];
          if (StringGrid1->Cells[0][2]!="") Str+=", �����. c"+StringGrid1->Cells[0][2]+" �� "+StringGrid1->Cells[1][2];
          if (StringGrid1->Cells[0][3]!="") Str+=", �����. c"+StringGrid1->Cells[0][3]+" �� "+StringGrid1->Cells[1][3];
          if (StringGrid1->Cells[0][4]!="") Str+=", �����. c"+StringGrid1->Cells[0][4]+" �� "+StringGrid1->Cells[1][4];
          if (StringGrid1->Cells[0][5]!="") Str+=", �����. c"+StringGrid1->Cells[0][5]+" �� "+StringGrid1->Cells[1][5];
          if (StringGrid1->Cells[0][6]!="") Str+=", �����. c"+StringGrid1->Cells[0][6]+" �� "+StringGrid1->Cells[1][6];
        }
      else if (fl_red==1)
        {
          Str ="�������������� ���������� ���������  �� "+IntToStr(Main->god)+" ��� �� ��������� '"+EditTN->Text+"': ";
          if (Main->SetNull(EditZEX_ZAM->Text)!=Main->SetNull(zzex_zam)) Str+=" ��� ���.���. � '"+QuotedStr(zzex_zam)+"' �� '"+QuotedStr(EditZEX_ZAM->Text)+"',";
          if (Main->SetNull(EditSHIFR_ZAM->Text)!=Main->SetNull(zshifr_zam)) Str+=" ���� ����.���.���. � '"+QuotedStr(zshifr_zam)+"' �� '"+QuotedStr(EditSHIFR_ZAM->Text)+"',";
          if (Main->SetNull(LabelTN_R->Caption)!=Main->SetNull(ztn_r)) Str+=" ��� ���.���. � '"+Main->SetNull(ztn_r)+"' �� '"+Main->SetNull(LabelTN_R->Caption)+"',";
          if (Main->SetNull(EditFIO_R->Text)!=Main->SetNull(zfio_r)) Str+=" ��� ���.���. � '"+QuotedStr(zfio_r)+"' �� '"+QuotedStr(EditFIO_R->Text)+"',";
          if (risk!=zrisk) Str+=" ���� � '"+QuotedStr(zrisk)+"' �� '"+QuotedStr(risk)+"',";
          if (Main->SetNull(EditRISK_PRICH->Text)!=Main->SetNull(zrisk_prich)) Str+=" ������� ����� � '"+Main->SetNull(zrisk_prich)+"' �� '"+Main->SetNull(EditRISK_PRICH->Text)+"',";
          if (gotov!=zgotov) Str+=" ���������� � '"+QuotedStr(zgotov)+"' �� '"+QuotedStr(gotov)+"',";
          if (Main->SetNull(StringGrid1->Cells[0][1])!=Main->SetNull(zdatn1)) Str+=" ���� ���.���. � '"+QuotedStr(zdatn1)+"' �� '"+QuotedStr(StringGrid1->Cells[0][1])+"',";
          if (Main->SetNull(StringGrid1->Cells[0][2])!=Main->SetNull(zdatn2)) Str+=" ���� ���.���. � '"+QuotedStr(zdatn2)+"' �� '"+QuotedStr(StringGrid1->Cells[0][2])+"',";
          if (Main->SetNull(StringGrid1->Cells[0][3])!=Main->SetNull(zdatn3)) Str+=" ���� ���.���. � '"+QuotedStr(zdatn3)+"' �� '"+QuotedStr(StringGrid1->Cells[0][3])+"',";
          if (Main->SetNull(StringGrid1->Cells[0][4])!=Main->SetNull(zdatn4)) Str+=" ���� ���.���. � '"+QuotedStr(zdatn4)+"' �� '"+QuotedStr(StringGrid1->Cells[0][4])+"',";
          if (Main->SetNull(StringGrid1->Cells[0][5])!=Main->SetNull(zdatn5)) Str+=" ���� ���.���. � '"+QuotedStr(zdatn5)+"' �� '"+QuotedStr(StringGrid1->Cells[0][5])+"',";
          if (Main->SetNull(StringGrid1->Cells[0][6])!=Main->SetNull(zdatn6)) Str+=" ���� ���.���. � '"+QuotedStr(zdatn6)+"' �� '"+QuotedStr(StringGrid1->Cells[0][6])+"',";
          if (Main->SetNull(StringGrid1->Cells[1][1])!=Main->SetNull(zdatk1)) Str+=" ���� ��.���. � '"+QuotedStr(zdatk1)+"' �� '"+QuotedStr(StringGrid1->Cells[1][1])+"',";
          if (Main->SetNull(StringGrid1->Cells[1][2])!=Main->SetNull(zdatk2)) Str+=" ���� ��.���. � '"+QuotedStr(zdatk2)+"' �� '"+QuotedStr(StringGrid1->Cells[1][2])+"',";
          if (Main->SetNull(StringGrid1->Cells[1][3])!=Main->SetNull(zdatk3)) Str+=" ���� ��.���. � '"+QuotedStr(zdatk3)+"' �� '"+QuotedStr(StringGrid1->Cells[1][3])+"',";
          if (Main->SetNull(StringGrid1->Cells[1][4])!=Main->SetNull(zdatk4)) Str+=" ���� ��.���. � '"+QuotedStr(zdatk4)+"' �� '"+QuotedStr(StringGrid1->Cells[1][4])+"',";
          if (Main->SetNull(StringGrid1->Cells[1][5])!=Main->SetNull(zdatk5)) Str+=" ���� ��.���. � '"+QuotedStr(zdatk5)+"' �� '"+QuotedStr(StringGrid1->Cells[1][5])+"',";
          if (Main->SetNull(StringGrid1->Cells[1][6])!=Main->SetNull(zdatk6)) Str+=" ���� ��.���. � '"+QuotedStr(zdatk6)+"' �� '"+QuotedStr(StringGrid1->Cells[1][6])+"',";
        }

      Main->InsertLog(Str);
      DM->qLogs->Requery();
    }
  else
    {
      Main->InsertLog("���������� ������ �� "+IntToStr(Main->god)+" ��� �� ���������: ���="+EditZEX->Text+" ���.�="+EditTN->Text+" �� ���������");
      DM->qLogs->Requery();
    }

  //���������� �������
  DM->qOcenka->Requery();
  DM->qZamesh->Requery();

  //����������� �� ��������� ������
  if (rec2==-1) rec2=1;
  DM->qZamesh->RecNo=rec2;
  DM->qOcenka->RecNo=rec1;
  
  Application->MessageBox("������ ������� ��������","��������������",
                           MB_OK+MB_ICONINFORMATION);

  //������� ������
  Panel3->Visible = false;
  Panel1->Align=alClient;
  Zameshenie->Height=440;
  BitBtn1->Top=316;
  BitBtn2->Top=355;
  Bevel3->Height=385;
}
//---------------------------------------------------------------------------

void __fastcall TZameshenie::FormKeyDown(TObject *Sender, WORD &Key,
      TShiftState Shift)
{
  if (Key==VK_RETURN)
  FindNextControl((TWinControl *)Sender, true, true, false)->SetFocus();        
}
//---------------------------------------------------------------------------

void __fastcall TZameshenie::DBGridEh1DblClick(TObject *Sender)
{
  if (DM->qZamesh->RecordCount==0)
    {
      //�������� ������
      N1DobavClick(Sender);
    }
  else
    {
      //������������� ������
      N2RedakClick(Sender);
    }
}
//---------------------------------------------------------------------------


//�������� ���������
void __fastcall TZameshenie::N3DeleteClick(TObject *Sender)
{
  int rec;

  if (Application->MessageBox("�� ������������� ������ ������������ ������� \n��������� �� ������� ���������?","��������������",
                              MB_YESNO+MB_ICONWARNING)==ID_NO)
    {
      Abort();
    }

  rec=DM->qOcenka->RecNo;
  //�������� ���������
  DM->qObnovlenie->Close();
  DM->qObnovlenie->SQL->Clear();
  DM->qObnovlenie->SQL->Add("delete from ocenka_rez where god="+IntToStr(Main->god)+" and tn="+EditTN->Text+"  \
                             and id_shtat="+DM->qZamesh->FieldByName("id_shtat")->AsString+" and rowid = chartorowid("+ QuotedStr(DM->qZamesh->FieldByName("rw")->AsString)+")");
  try
    {
      DM->qObnovlenie->ExecSQL();
    }
  catch(Exception &E)
    {
      Application->MessageBox(("�������� ������ ��� ��������� ������ �� ������� Ocenka" + E.Message).c_str(),"������",
                               MB_OK+MB_ICONERROR);
      Main->InsertLog("�������� ������ ��� �������� ����������/��������� ���������("+DM->qZamesh->FieldByName("id_shtat")->AsString+") �� ���������: ���.�='"+EditTN->Text+"' ���='"+LabelFIO->Caption+"'");
      Abort();
    }

  DM->qZamesh->Requery();
  //�������� �������� ���������, ���� ��� ���� ������������ ������
  if (DM->qZamesh->RecordCount==0)
    {
      AnsiString zap;
      if (CheckBoxZAM->Checked==true) zap="zam";
      else if (CheckBoxREZERV->Checked==true) zap="rezerv";


      //���������� �������� � ������� ���������
      DM->qObnovlenie->Close();
      DM->qObnovlenie->SQL->Clear();
      DM->qObnovlenie->SQL->Add("update ocenka set "+zap+"=NULL where god="+IntToStr(Main->god)+" and tn="+EditTN->Text);
      try
        {
          DM->qObnovlenie->ExecSQL();
        }
      catch (Exception &E)
        {
          Application->MessageBox(("�������� ������ ��� ���������� �������� ���������� � ������� OCENKA "+E.Message).c_str(),"������",
                                   MB_OK+MB_ICONERROR);
          DM->qZamesh->Requery();
          Main->InsertLog("�������� ������ ��� ���������� �������� ���������/���������� �� ���������: ���.�='"+EditTN->Text+"' ���='"+LabelFIO->Caption+"'");
          Abort();
        }

       //���������� ������
       DM->qZamesh->Requery();
       DM->qOcenka->Requery();

      if (CheckBoxZAM->Checked==true)
        {
          CheckBoxZAM->Checked=false;
          CheckBoxREZERV->Enabled=true;
        }

    }

  
  //����
  Main->InsertLog("�������� ����������/��������� ��������� ��������� ������� �� ���������: ���.�='"+EditTN->Text+"' ���='"+LabelFIO->Caption+"'");

  DM->qOcenka->RecNo=rec;

}
//---------------------------------------------------------------------------


