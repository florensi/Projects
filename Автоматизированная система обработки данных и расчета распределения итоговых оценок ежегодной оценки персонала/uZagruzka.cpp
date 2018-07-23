//---------------------------------------------------------------------------

#include <vcl.h>
#pragma hdrstop

#include "uZagruzka.h"
#include "uDM.h"
#include "RepoRTFM.h"
#include "RepoRTFO.h"
#include "uMain.h"
//---------------------------------------------------------------------------
#pragma package(smart_init)
#pragma resource "*.dfm"
TZagruzka *Zagruzka;
//---------------------------------------------------------------------------
__fastcall TZagruzka::TZagruzka(TComponent* Owner)
        : TForm(Owner)
{
}
//---------------------------------------------------------------------------
void __fastcall TZagruzka::CheckBox1Click(TObject *Sender)
{
  //������� ����� ��� �������� �� Excel
  EditDATA->Text = "";
  EditFIO->Text = "";
  EditDOLGO->Text = "";
  EditOCENKA->Text = "";
  EditREZERV->Text = "";
  EditREZULT_OCEN->Text = "";
  EditKPE_OCEN->Text = "";
  EditKOMP_OCEN->Text = "";
  EditDOLG->Text = "";
  EditZEX->Text = "";
  EditTN->Text = "";
  EditFIOEOP->Text = "";
  EditTNEOP->Text = "";
  EditTN_KPE->Text = "";
  EditKPE1->Text = "";
  EditKPE2->Text = "";
  EditKPE3->Text = "";
  EditKPE4->Text = "";
  EditTN_VZ->Text = "";
  EditVZ->Text = "";
  EditKR_ZEX->Text = "";
  EditTN_KR->Text = "";
  EditKR_FIO->Text = "";
  EditKRSHIFR_DOLG->Text = "";

  //��������� ����� ��� �������� �� Excel
  if (CheckBox1->Checked)
    {
      EditDATA->Visible = true;
      EditFIO->Visible = true;
      EditDOLGO->Visible = true;
      EditOCENKA->Visible = true;
      EditREZERV->Visible = true;
      EditDOLG->Visible = true;
      EditZEX->Visible = true;
      EditTN->Visible = true;
      EditREZULT_OCEN->Visible = true;
      EditKPE_OCEN->Visible = true;
      EditKOMP_OCEN->Visible = true;
      EditFIOEOP->Visible = true;
      EditTNEOP->Visible = true;
      EditFIO->Visible = true;
      EditTN_KPE->Visible = true;
      EditKPE1->Visible = true;
      EditKPE2->Visible = true;
      EditKPE3->Visible = true;
      EditKPE4->Visible = true;
      EditTN_VZ->Visible = true;
      EditVZ->Visible = true;

      EditKR_ZEX->Visible = true;
      EditTN_KR->Visible = true;
      EditKR_FIO->Visible = true;
      EditKRSHIFR_DOLG->Visible = true;

      LabelDATA->Visible = true;
      LabelDOLGO->Visible = true;
      LabelOCENKA->Visible = true;
      LabelREZERV->Visible = true;
      LabelDOLG->Visible = true;
      LabelZEX->Visible = true;
      LabelTN->Visible = true;
      LabelREZULT_OCEN->Visible = true;
      LabelKPE_OCEN->Visible = true;
      LabelKOMP_OCEN->Visible = true;
      LabelFIOEOP->Visible = true;
      LabelTNEOP->Visible = true;
      LabelTN_KPE->Visible = true;
      LabelKPE1->Visible = true;
      LabelKPE2->Visible = true;
      LabelKPE3->Visible = true;
      LabelKPE4->Visible = true;
      LabelTN_VZ->Visible = true;
      LabelVZ->Visible = true;

      LabelKR_ZEX->Visible = true;
      LabelTN_KR->Visible = true;
      LabelKR_FIO->Visible = true;
      LabelKRSHIFR_DOLG->Visible = true;


      Application->MessageBox("��� ��������� ����� ��� �������� \n������ � ����� ��� \n������� ��� ������ (�������� E28), \n��� �������� ������ ������ ������ \n���������� ����� �������","������ ������",MB_OK+MB_ICONINFORMATION);
    }
  else
    {
      EditDATA->Visible = false;
      EditFIO->Visible = false;
      EditDOLGO->Visible = false;
      EditOCENKA->Visible = false;
      EditREZERV->Visible = false;
      EditDOLG->Visible = false;
      EditZEX->Visible = false;
      EditTN->Visible = false;
      EditREZULT_OCEN->Visible = false;
      EditKPE_OCEN->Visible = false;
      EditKOMP_OCEN->Visible = false;
      EditFIOEOP->Visible = false;
      EditTNEOP->Visible = false;
      EditFIO->Visible = false;
      EditTN_KPE->Visible = false;
      EditKPE1->Visible = false;
      EditKPE2->Visible = false;
      EditKPE3->Visible = false;
      EditKPE4->Visible = false;
      EditTN_VZ->Visible = false;
      EditVZ->Visible = false;

      EditKR_ZEX->Visible = false;
      EditTN_KR->Visible = false;
      EditKR_FIO->Visible = false;
      EditKRSHIFR_DOLG->Visible = false;

      LabelDATA->Visible = false;
      LabelDOLGO->Visible = false;
      LabelOCENKA->Visible = false;
      LabelREZERV->Visible = false;
      LabelDOLG->Visible = false;
      LabelZEX->Visible = false;
      LabelTN->Visible = false;
      LabelREZULT_OCEN->Visible = false;
      LabelKPE_OCEN->Visible = false;
      LabelKOMP_OCEN->Visible = false;
      LabelFIOEOP->Visible = false;
      LabelTNEOP->Visible = false;
      LabelTN_KPE->Visible = false;
      LabelKPE1->Visible = false;
      LabelKPE2->Visible = false;
      LabelKPE3->Visible = false;
      LabelKPE4->Visible = false;
      LabelTN_VZ->Visible = false;
      LabelVZ->Visible = false;

      LabelKR_ZEX->Visible = false;
      LabelTN_KR->Visible = false;
      LabelKR_FIO->Visible = false;
      LabelKRSHIFR_DOLG->Visible = false;
    }
}
//---------------------------------------------------------------------------
void __fastcall TZagruzka::SpeedButton2Click(TObject *Sender)
{
  Zagruzka->Close();
}
//---------------------------------------------------------------------------

// �������� �� �������� ���� � Excel-�����
bool  __fastcall TZagruzka::Proverka(AnsiString zex)
{
   try {
    StrToInt(zex);
  }
  catch (...) {
    return false;
  }
  return true;

}
//---------------------------------------------------------------------------

//�������� ����� � ����� Excel
void __fastcall TZagruzka::SpeedButton1Click(TObject *Sender)
{
  int doc=0, pole_data, pole_fio, pole_dolgo, n,
             pole_ocenka,pole_ocenka2, pole_rezerv, pole_dolg,
             pole_zex, pole_tn,
             pole_tn_kpe, pole_kpe1, pole_kpe2, pole_kpe3,
             pole_kpe4, pole_tn_vz, pole_vz,
             pole_kr_zex, pole_tn_kr, pole_kr_fio, pole_krshifr_dolg,
             pole_rezerv_dolg_kr,
             pole_id_shtat,
             otchet=0, kol=0, rec=0,
             ob_kol=0, obnov_kol=0, kol_zam=0;
  AnsiString zex, Sql, Dir, rezerv, logi,
             pole_rez, pole_kpe, pole_komp, pole_tneop, pole_fioeop;
  TDateTime d;
  Variant AppEx, Sh;
  TLocateOptions SearchOptions;

  /*Dir - ���� � ��������� �����
    logi - ���� � ����������� �� ��������� ��� �������� ������*/


  Main->StatusBar1->SimpleText=" ���� �������� ������...";

  // ������������ ����� ��� ��������
  if (RadioButtonDATAO->Checked)
    {
      //���� ������
      if (EditDATA->Text.IsEmpty()) pole_data=41; //"AO"
      else  pole_data = StrToInt(EditDATA->Text);
      //��� ��������
      if (EditFIO->Text.IsEmpty()) pole_fio=39; //"AM"
      else pole_fio = StrToInt(EditFIO->Text);
      //��������� ��������
      if (EditDOLGO->Text.IsEmpty()) pole_dolgo=40; //"AN"
      else pole_dolgo = StrToInt(EditDOLGO->Text);

      logi = "�������� ������ � ��� �������� �� "+IntToStr(Main->god)+" ���";
    }
  else if (RadioButtonEOP->Checked)
    {
      //���������� ������
      if (EditREZULT_OCEN->Text.IsEmpty()) pole_rez="H41";//"8:42";
      else  pole_rez = StrToInt(EditREZULT_OCEN->Text);
      //���������� ������ �� ���
      if (EditKPE_OCEN->Text.IsEmpty()) pole_kpe="E28";//"5:28";
      else  pole_kpe = StrToInt(EditKPE_OCEN->Text);
      //���.�
      if (EditTNEOP->Text.IsEmpty()) pole_tneop="E10";//"5:10";
      else  pole_tneop = StrToInt(EditTNEOP->Text);
      //���
      if (EditFIOEOP->Text.IsEmpty()) pole_fioeop="E9";//"5:9";
      else  pole_fioeop = StrToInt(EditFIOEOP->Text);
      Main->ProgressBar->Position = 0;
      logi = "�������� ������ � ����� ��� �� "+IntToStr(Main->god)+" ���";
    }
  else if (CheckBoxOCENKA->Checked && CheckBoxREZERV->Checked)
    {
      //����������������� ������
      if (EditOCENKA->Text.IsEmpty()) pole_ocenka=31;//"AE";
      else  pole_ocenka = StrToInt(EditOCENKA->Text);
      //������������ ���������
      pole_ocenka2=32;

      //������
      if (EditREZERV->Text.IsEmpty()) pole_rezerv=33;//"AG";
      else  pole_rezerv = StrToInt(EditREZERV->Text);
      //��������� ����������
      if (EditDOLG->Text.IsEmpty()) pole_dolg=34;//"AH";
      else  pole_dolg = StrToInt(EditDOLG->Text);

      logi = "�������� ������ ����������������� ������������� � ������������ � �������� ������ �� "+IntToStr(Main->god)+" ���";
    }
  else if (CheckBoxOCENKA->Checked)
    {
      //����������������� ������
      if (EditOCENKA->Text.IsEmpty()) pole_ocenka=31;//"AE";
      else  pole_ocenka = StrToInt(EditOCENKA->Text);

      //������������ ���������
      pole_ocenka2=32;


      logi = "�������� ������ ����������������� ������������� �� "+IntToStr(Main->god)+" ���";
    }
  else if (CheckBoxREZERV->Checked)
    {
      //������
      if (EditREZERV->Text.IsEmpty()) pole_rezerv=33;//"AG";
      else  pole_rezerv = StrToInt(EditREZERV->Text);
      //��������� ����������
      if (EditDOLG->Text.IsEmpty()) pole_dolg=34;//"AH";
      else  pole_dolg = StrToInt(EditDOLG->Text);

      logi = "�������� ������������ � �������� ������ �� "+IntToStr(Main->god)+" ���";
    }
  else if (RadioButtonKPE->Checked)
    {
      //���.�
      if (EditTN_KPE->Text.IsEmpty()) pole_tn_kpe=2;//"B";
      else  pole_tn_kpe = StrToInt(EditTN_KPE->Text);
      //���1
      if (EditKPE1->Text.IsEmpty()) pole_kpe1=7;//"G";
      else  pole_kpe1 = StrToInt(EditKPE1->Text);
      //���2
      if (EditKPE2->Text.IsEmpty()) pole_kpe2=8;//"H";
      else  pole_kpe2 = StrToInt(EditKPE2->Text);
      //���3
      if (EditKPE3->Text.IsEmpty()) pole_kpe3=9;//"I";
      else  pole_kpe3 = StrToInt(EditKPE3->Text);
      //���4
      if (EditKPE4->Text.IsEmpty()) pole_kpe4=10;//"J";
      else  pole_kpe4 = StrToInt(EditKPE4->Text);

      logi = "�������� ��� ����������� ����������� �� "+IntToStr(Main->god)+" ���";
    }
  else if (RadioButtonVZ->Checked)
    {
      //���.�
      if (EditTN_VZ->Text.IsEmpty()) pole_tn_vz=3;//"�";
      else  pole_tn_vz = StrToInt(EditTN_VZ->Text);
      //�������
      if (EditVZ->Text.IsEmpty()) pole_vz=7;//"G";
      else  pole_vz = StrToInt(EditVZ->Text);

      logi = "�������� �������� ����� �� ������ ����������� �� "+IntToStr(Main->god)+" ���";
    }
  else if (RadioButtonKR->Checked)
    {
      //���.�
      if (EditTN_KR->Text.IsEmpty()) pole_tn_kr=6;//"F";
      else  pole_tn_kr = StrToInt(EditTN_KR->Text);
      //���
      if (EditKR_ZEX->Text.IsEmpty()) pole_kr_zex=8;//"H";
      else  pole_kr_zex = StrToInt(EditKR_ZEX->Text);
      //���
      if (EditKR_FIO->Text.IsEmpty()) pole_kr_fio=5;//"E";
      else  pole_kr_fio = StrToInt(EditKR_FIO->Text);
      //���������
      if (EditKRSHIFR_DOLG->Text.IsEmpty()) pole_krshifr_dolg=3;//"C";
      else  pole_krshifr_dolg = StrToInt(EditKRSHIFR_DOLG->Text);
      //������������ ���������
      pole_rezerv_dolg_kr=4; //"C"
      //���� ��������
      pole_id_shtat=2;//B


      logi = "�������� ������������ � �� �� ������ �� �� �� "+IntToStr(Main->god)+" ���";

    }
  else
    {
      Application->MessageBox("�� ������ ��� �������� ������","��������������",
                                MB_OK+MB_ICONINFORMATION);
      Abort();
    }

  //���� ��� � ���.�
  if (EditZEX->Text.IsEmpty()) pole_zex=10;//"J";
  else pole_zex = StrToInt(EditZEX->Text);
  if (EditTN->Text.IsEmpty()) pole_tn=3;//"C";
  else pole_tn = StrToInt(EditTN->Text);

  Main->StatusBar1->SimpleText=" ����� ����� � �����������...";

  //����� ����� � �����������
  if (!SelectDirectory("Select directory",WideString(""),Dir))
    {
      Main->StatusBar1->SimpleText ="�������� ������: "+IntToStr(Main->god)+" ���";
      Abort();
    }

  Main->StatusBar1->SimpleText=" ��������� ������ ���� ������ � �����...";

  //��������� ������ ���� ������ � �����
  FileListBox1->Directory = Dir;
  //FindClose(SearchRecord);   //����������� �������, ������ ��������� ������


  //�������� �� ������� ������ Excel � �����
  if (FileListBox1->Count==0)
    {
      Application->MessageBox("�������� ����� �� �������� ������ Excel!!!","��������������",
                              MB_OK+MB_ICONINFORMATION);

      Main->StatusBar1->SimpleText ="�������� ������: "+IntToStr(Main->god)+" ���";
      Abort();
    }

  //�������� ����� ������ ��� ������ �� ����������� ������
  if (!rtf_Open((Main->TempPath + "\\zagruzka.txt").c_str()))
    {
      MessageBox(Handle,"������ �������� ����� ������","������",8192);
      Abort();
    }
  rtf_Out("data", Now(),0);

  //���� �� ���� ���������� � �����
  while (doc<FileListBox1->Count)
    {
      Main->StatusBar1->SimpleText = " �������� ������ �� ����� "+FileListBox1->Items->Strings[doc];

      //�������� ��������� Excel
      try
        {
          AppEx = CreateOleObject("Excel.Application");
        }
      catch (...)
        {
          Application->MessageBox("���������� ������� Microsoft Excel!"
                                  " �������� ��� ���������� �� ���������� �� �����������.","������",MB_OK+MB_ICONERROR);
          Main->StatusBar1->SimpleText ="�������� ������: "+IntToStr(Main->god)+" ���";
          Abort();
        }

      //���� ��������� ������ �� ����� ������������ ������
      try
        {
          try
            {
              AppEx.OlePropertyGet("Workbooks").OlePropertyGet("Open", (Dir +"\\"+ FileListBox1->Items->Strings[doc]).c_str());
              AppEx.OlePropertySet("Visible",false);
              Sh = AppEx.OlePropertyGet("Worksheets", 1);
              // MsExcel.ActiveSheet.Names.Item('_FilterDatabase').Delete;
            }
          catch(...)
            {
              Application->MessageBox("������ �������� ����� Microsoft Excel!","������",MB_OK + MB_ICONERROR);
              Main->StatusBar1->SimpleText ="�������� ������: "+IntToStr(Main->god)+" ���";
            }

          rec=0;
          kol=0;

          //���������� ���������� ������� ����� � ���������
          AnsiString Row = Sh.OlePropertyGet("UsedRange").OlePropertyGet("Rows").OlePropertyGet("Count");

          Cursor = crHourGlass;

          if (RadioButtonEOP->Checked == false) Main->ProgressBar->Position = 0;
          Main->ProgressBar->Visible = true;
          Main->ProgressBar->Max=StrToInt(Row)+1;

//�������� ������ c ���
//******************************************************************************
          if (RadioButtonEOP->Checked == true)
            {
              Main->ProgressBar->Max=FileListBox1->Count;

              //�������� ���� � �������� ����������� �������� �����������
              if (EditKOMP_OCEN->Text.IsEmpty())
                {
                  if (String(Sh.OlePropertyGet("Range","H3"))=="������������")
                    { //������������
                      if (String(Sh.OlePropertyGet("Range","H4"))=="������������ �������������") pole_komp = "H100";
                      else if (String(Sh.OlePropertyGet("Range","H4"))=="�������� ��������") pole_komp = "H97";
                      else if (String(Sh.OlePropertyGet("Range","H4"))=="���������") pole_komp = "H87";
                    }
                  else if (String(Sh.OlePropertyGet("Range","H3"))=="���������� ������ � �������")
                    { //���������� ������ � �������
                      if (String(Sh.OlePropertyGet("Range","H4"))=="������������ �������������") pole_komp = "H98";
                      else if (String(Sh.OlePropertyGet("Range","H4"))=="�������� ��������") pole_komp = "H96";
                      else if (String(Sh.OlePropertyGet("Range","H4"))=="���������") pole_komp = "H90";
                    }
                  else
                    {
                      Application->MessageBox("�� ������� �������������� ������ � �����!!!","��������������",
                                               MB_OK+MB_ICONINFORMATION);
                      Abort();
                    }
                }
              else  pole_komp = StrToInt(EditKOMP_OCEN->Text);

              //�������� �� ������ ����
              if (String(Sh.OlePropertyGet("Range",pole_komp.c_str())).IsEmpty() || String(Sh.OlePropertyGet("Range",pole_komp.c_str()))=="0" ||
                 (String(Sh.OlePropertyGet("Range",pole_rez.c_str())).IsEmpty() && String(Sh.OlePropertyGet("Range",pole_kpe.c_str())).IsEmpty()) ||
                 (String(Sh.OlePropertyGet("Range",pole_rez.c_str()))=="0" && String(Sh.OlePropertyGet("Range",pole_kpe.c_str()))=="0" ))
                {
                  Application->MessageBox(("�� ������� ���������� ������ ��� ����������� \n� ����� '"+FileListBox1->Items->Strings[doc]+"' \n�� ��������� '"+String(Sh.OlePropertyGet("Range","E9"))+"'. \n������� ��������� � ��������� �������� ������� ����� \n��� ��������������� ������ �������").c_str(), "��������������",
                                            MB_OK+MB_ICONWARNING);

                  //������������ ������ �� ������������� �������
                  rtf_Out("zex", String(Sh.OlePropertyGet("Range","E12")),1);
                  rtf_Out("tn", String(Sh.OlePropertyGet("Range","E10")),1);
                  rtf_Out("fio", String(Sh.OlePropertyGet("Range","E9"))+ " (��� ������ �� ����������� ������ ��� ������������, ���� "+FileListBox1->Items->Strings[doc]+" )",1);

                  if(!rtf_LineFeed())
                    {
                      MessageBox(Handle,"������ ������ � ���� ������","������",8192);
                      if (!rtf_Close()) MessageBox(Handle,"������ �������� ����� ������","������",8192);
                      return;
                    }

                  otchet=1;  //������� ������������ ������ �� ������������� �������
                }
              //���������� ������ �� ����� 4
              else if (!String(Sh.OlePropertyGet("Range",pole_rez.c_str())).IsEmpty() && Double(Sh.OlePropertyGet("Range",pole_rez.c_str()))>4)
                {
                  Application->MessageBox(("���������� ������ ��������� 4 \n� ����� '"+FileListBox1->Items->Strings[doc]+"' \n�� ��������� '"+String(Sh.OlePropertyGet("Range","E9"))+"'. \n������� ��������� � ��������� �������� ������� ����� \n��� ��������������� ������ �������").c_str(),"��������������",
                                            MB_OK+MB_ICONINFORMATION);

                  //������������ ������ �� ������������� �������
                  rtf_Out("zex", String(Sh.OlePropertyGet("Range","E12")),1);
                  rtf_Out("tn", String(Sh.OlePropertyGet("Range","E10")),1);
                  rtf_Out("fio", String(Sh.OlePropertyGet("Range","E9"))+ " (���������� ������ ��������� 4, ���� "+FileListBox1->Items->Strings[doc]+")",1);

                  if(!rtf_LineFeed())
                    {
                      MessageBox(Handle,"������ ������ � ���� ������","������",8192);
                      if (!rtf_Close()) MessageBox(Handle,"������ �������� ����� ������","������",8192);
                      return;
                    }
                  otchet=1;  //������� ������������ ������ �� ������������� �������
                }
              //����������� �� ����� 32
              else if (Double(Sh.OlePropertyGet("Range",pole_komp.c_str()))>32)
                {
                  Application->MessageBox(("����������� ��������� 32 \n� ����� '"+FileListBox1->Items->Strings[doc]+"' \n�� ��������� '"+String(Sh.OlePropertyGet("Range","E9"))+"'. \n������� ��������� � ��������� �������� ������� ����� \n��� ��������������� ������ �������").c_str(),"��������������",
                                            MB_OK+MB_ICONINFORMATION);

                  //������������ ������ �� ������������� �������
                  rtf_Out("zex", String(Sh.OlePropertyGet("Range","E12")),1);
                  rtf_Out("tn", String(Sh.OlePropertyGet("Range","E10")),1);
                  rtf_Out("fio", String(Sh.OlePropertyGet("Range","E9"))+ " (����������� ��������� 32, ���� "+FileListBox1->Items->Strings[doc]+")",1);

                  if(!rtf_LineFeed())
                    {
                      MessageBox(Handle,"������ ������ � ���� ������","������",8192);
                      if (!rtf_Close()) MessageBox(Handle,"������ �������� ����� ������","������",8192);
                      return;
                    }
                  otchet=1;  //������� ������������ ������ �� ������������� �������
                }
              else
                {
                  //������ ������������� � ��.
                  AnsiString rez_r, komp, kpe, efect;

                  //���������� ������
                  if (String(Sh.OlePropertyGet("Range",pole_rez.c_str())).IsEmpty() || String(Sh.OlePropertyGet("Range",pole_rez.c_str()))=="0") rez_r="NULL";
                  else rez_r= FloatToStrF(Double(Sh.OlePropertyGet("Range",pole_rez.c_str())), ffFixed, 10,2);

                  //�����������
                  komp= FloatToStrF(Double(Sh.OlePropertyGet("Range",pole_komp.c_str())), ffFixed, 10,2);

                  //�������������
                  if ((String(Sh.OlePropertyGet("Range",pole_rez.c_str())).IsEmpty() ||
                       String(Sh.OlePropertyGet("Range",pole_rez.c_str()))=="0") &&
                       Double(Sh.OlePropertyGet("Range",pole_kpe.c_str()))>0)
                    {
                      efect = FloatToStrF(Double(Sh.OlePropertyGet("Range",pole_kpe.c_str()))*0.6+((StrToFloat(komp)*100/32)*0.4), ffFixed, 10,2);
                    }
                  else if (Double(Sh.OlePropertyGet("Range",pole_rez.c_str()))>0)
                    {
                      efect = FloatToStrF((StrToFloat(rez_r)*100/4)*0.6+(StrToFloat(komp)*100/32)*0.4, ffFixed, 10,2);
                    }

                  Sql = "update ocenka set \
                                           rezult_ocen = "+ rez_r +", \
                                           komp_ocen = "+ komp+",   \
                                           kpe_ocen = decode("+ Main->SetNull(Double(Sh.OlePropertyGet("Range",pole_kpe.c_str())))+",0,NULL,"+ Main->SetNull(Double(Sh.OlePropertyGet("Range",pole_kpe.c_str())))+"), \
                                           efect = "+efect+" \
                         where upper(substr(trim(fio),1,instr(trim(fio),' ')-1))=substr(upper(trim("+QuotedStr(Sh.OlePropertyGet("Range",pole_fioeop.c_str()))+")),1,instr("+QuotedStr(Sh.OlePropertyGet("Range",pole_fioeop.c_str()))+" ,' ')-1) \
                         and tn="+ Sh.OlePropertyGet("Range",pole_tneop.c_str()) +"  and god="+IntToStr(Main->god);

                  DM->qObnovlenie->Close();
                  DM->qObnovlenie->SQL->Clear();
                  DM->qObnovlenie->SQL->Add(Sql);
                  try
                    {
                      DM->qObnovlenie->ExecSQL();
                    }
                  catch(Exception &E)
                    {
                      Application->MessageBox(("�������� ������ ��� ������� �������� ������ � ������� ocenka_nadya" + E.Message).c_str(),"������",
                                                MB_OK+MB_ICONERROR);

                      Main->InsertLog(logi+". �������� ������ ��� ������� ���������� ������ � ������� Ocenka �� ����� '"+FileListBox1->Items->Strings[doc]+"'");
                      DM->qLogs->Requery();
                      Main->StatusBar1->SimpleText ="�������� ������: "+IntToStr(Main->god)+" ���";
                      Abort();
                    }

                  rec++;
                  kol+=DM->qObnovlenie->RowsAffected;

                  // ���������� ����������� �������
                  if (DM->qObnovlenie->RowsAffected == 0)
                    {
                      //������������ ������ �� ������������� �������
                      rtf_Out("zex", String(Sh.OlePropertyGet("Range","E12")),1);
                      rtf_Out("tn", String(Sh.OlePropertyGet("Range","E10")),1);
                      rtf_Out("fio", String(Sh.OlePropertyGet("Range","E9")),1);

                      if(!rtf_LineFeed())
                        {
                          MessageBox(Handle,"������ ������ � ���� ������","������",8192);
                          if (!rtf_Close()) MessageBox(Handle,"������ �������� ����� ������","������",8192);
                          return;
                        }
                      otchet=1;  //������� ������������ ������ �� ������������� �������
                    }
                  Main->ProgressBar->Position++;
              }
            }
//�������� ��������, ����������������� ������, �������
//******************************************************************************
          else if (RadioButtonDATAO->Checked==true ||
                    CheckBoxOCENKA->Checked==true ||
                    CheckBoxREZERV->Checked==true)
            {
              for (int i=1; i<Row+1; i++)
                {
                  zex = Sh.OlePropertyGet("Cells",i,pole_tn);

                  if (zex.IsEmpty() || !Proverka(zex))  continue;
                    {
                      //�������� ������ � ����
                      //�������� ���� ���������� � ��� ��������
                      if (RadioButtonDATAO->Checked)
                        {
                          //�������� ������������ ����
                          if (!String(Sh.OlePropertyGet("Cells",i,pole_data)).IsEmpty())
                            {
                              if(!TryStrToDate(Sh.OlePropertyGet("Cells",i,pole_data),d))
                                {
                                  Application->MessageBox(("������� �������������� ���� ������ \n�� ��������� � ���="+String(Sh.OlePropertyGet("Cells",i,pole_tn))+" � ���������="+String(Sh.OlePropertyGet("Cells",i,pole_zex))+"\n���������� ��������� ���� ������ � ����� Excel \n� ��������� �������� ��� ��������������� ������ �������").c_str(),
                                                           "������", MB_OK+MB_ICONWARNING);
                                  //������������ ������ �� ������������� �������
                                  rtf_Out("zex", String(Sh.OlePropertyGet("Cells",i,pole_zex)),1);
                                  rtf_Out("tn", String(Sh.OlePropertyGet("Cells",i,pole_tn)),1);
                                  rtf_Out("fio", String(Sh.OlePropertyGet("Cells",i,2))+" (�������������� ���� ������)" ,1);

                                  if(!rtf_LineFeed())
                                    {
                                      MessageBox(Handle,"������ ������ � ���� ������","������",8192);
                                      if (!rtf_Close()) MessageBox(Handle,"������ �������� ����� ������","������",8192);
                                      return;
                                    }
                                  otchet=1;  //������� ������������ ������ �� ������������� �������
                                }
                              else
                                {
                                  Sql = "update ocenka set \
                                                           data_ocen=to_date("+ QuotedStr(Sh.OlePropertyGet("Cells",i,pole_data)) +", 'dd.mm.yyyy'), \
                                                           fio_ocen=initcap(trim("+ QuotedStr(Sh.OlePropertyGet("Cells",i,pole_fio)) +")),  \
                                                           dolg_ocen=trim("+ QuotedStr(Sh.OlePropertyGet("Cells",i,pole_dolgo)) +")  \
                                          where tn="+ Sh.OlePropertyGet("Cells",i,pole_tn)+" and god="+IntToStr(Main->god);     //direkt="+Sh.OlePropertyGet("Cells",i,pole_zex) +" and
                                }
                            }
                          else
                            {
                              Sql = "update ocenka set \
                                                      data_ocen=to_date("+ QuotedStr(Sh.OlePropertyGet("Cells",i,pole_data)) +", 'dd.mm.yyyy'), \
                                                      fio_ocen=initcap(trim("+ QuotedStr(Sh.OlePropertyGet("Cells",i,pole_fio)) +")),  \
                                                      dolg_ocen=trim("+ QuotedStr(Sh.OlePropertyGet("Cells",i,pole_dolgo)) +")  \
                                     where tn="+ Sh.OlePropertyGet("Cells",i,pole_tn)+" and god="+IntToStr(Main->god);         //direkt="+Sh.OlePropertyGet("Cells",i,pole_zex) +" and
                            }
                        }
                      //�������� ����������������� ������ � �������
                      else if (CheckBoxOCENKA->Checked && CheckBoxREZERV->Checked)
                        {
                          if (AnsiUpperCase(Trim(String(Sh.OlePropertyGet("Cells",i,pole_rezerv))))==AnsiUpperCase("��")) rezerv=1;
                          else if (AnsiUpperCase(Trim(String(Sh.OlePropertyGet("Cells",i,pole_rezerv))))==AnsiUpperCase("���")) rezerv=0;
                          else  rezerv="NULL";

                          Sql = "update ocenka set \
                                                   kom_reit="+ QuotedStr(Sh.OlePropertyGet("Cells",i,pole_ocenka2)) +", \
                                                   skor_reit="+ QuotedStr(Sh.OlePropertyGet("Cells",i,pole_ocenka)) +", \
                                                   rezerv="+ rezerv +", \
                                                   dolg_rezerv=trim("+ QuotedStr(Sh.OlePropertyGet("Cells",i,pole_dolg)) +")  \
                                 where tn="+ Sh.OlePropertyGet("Cells",i,pole_tn)+"  and god="+IntToStr(Main->god);  //direkt="+Sh.OlePropertyGet("Cells",i,pole_zex) +" and
                        }
                      //�������� ����������������� ������
                      else if (CheckBoxOCENKA->Checked)
                        {
                          Sql = "update ocenka set \
                                                  kom_reit="+ QuotedStr(Sh.OlePropertyGet("Cells",i,pole_ocenka2)) +", \
                                                  skor_reit="+ QuotedStr(Sh.OlePropertyGet("Cells",i,pole_ocenka)) +" \
                                 where tn="+ Sh.OlePropertyGet("Cells",i,pole_tn)+"  and god="+IntToStr(Main->god);  //direkt="+Sh.OlePropertyGet("Cells",i,pole_zex) +" and
                        }
                      //�������� �������
                      else if (CheckBoxREZERV->Checked)
                        {
                          if (AnsiUpperCase(Trim(String(Sh.OlePropertyGet("Cells",i,pole_rezerv))))==AnsiUpperCase("��")) rezerv=1;
                          else if (AnsiUpperCase(Trim(String(Sh.OlePropertyGet("Cells",i,pole_rezerv))))==AnsiUpperCase("���")) rezerv=0;
                          else  rezerv="NULL";

                          Sql = "update ocenka set \
                                                   rezerv="+ rezerv +", \
                                                   dolg_rezerv=trim("+ QuotedStr(Sh.OlePropertyGet("Cells",i,pole_dolg)) +")  \
                                 where tn="+ Sh.OlePropertyGet("Cells",i,pole_tn)+"  and god="+IntToStr(Main->god);  //direkt="+Sh.OlePropertyGet("Cells",i,pole_zex) +" and 
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
                          Application->MessageBox(("�������� ������ ��� ������� �������� ������ � ������� Ocenka" + E.Message).c_str(),"������",
                                                    MB_OK+MB_ICONERROR);

                          Main->InsertLog(logi+". �������� ������ ��� ������� ���������� ������ � ������� OCENKA �� ����� '"+FileListBox1->Items->Strings[doc]+"'");
                          DM->qLogs->Requery();
                          DM->qOcenka->Requery();
                          Main->StatusBar1->SimpleText ="�������� ������: "+IntToStr(Main->god)+" ���";
                          Abort();
                        }

                      rec++;
                      kol+=DM->qObnovlenie->RowsAffected;

                      // ���������� ����������� �������
                      if (DM->qObnovlenie->RowsAffected == 0)
                        {
                          //������������ ������ �� ������������� �������
                          rtf_Out("zex", String(Sh.OlePropertyGet("Cells",i,pole_zex)),1);
                          rtf_Out("tn", String(Sh.OlePropertyGet("Cells",i,pole_tn)),1);
                          rtf_Out("fio", String(Sh.OlePropertyGet("Cells",i,2)),1);

                          if(!rtf_LineFeed())
                            {
                              MessageBox(Handle,"������ ������ � ���� ������","������",8192);
                              if (!rtf_Close()) MessageBox(Handle,"������ �������� ����� ������","������",8192);
                              return;
                            }
                          otchet=1;  //������� ������������ ������ �� ������������� �������
                        }
                    }
                  Main->ProgressBar->Position++;
                }
          }
//�������� ��� ���������
//******************************************************************************
         else if (RadioButtonKPE->Checked == true)
           {
             for (int i=1; i<Row+1; i++)
               {
                 zex = Sh.OlePropertyGet("Cells",i,pole_tn_kpe);

                 if (zex.IsEmpty() || !Proverka(zex))  continue;
                   {
                     //�������� ������ � ����
                     //�������� ������������ ���������� �����
                     if (String(Sh.OlePropertyGet("Cells",i,pole_kpe1)).IsEmpty() &&
                         String(Sh.OlePropertyGet("Cells",i,pole_kpe2)).IsEmpty() &&
                         String(Sh.OlePropertyGet("Cells",i,pole_kpe3)).IsEmpty() &&
                         String(Sh.OlePropertyGet("Cells",i,pole_kpe4)).IsEmpty() )
                       {
                         Application->MessageBox(("�� ��������� ���� �� ��������� ���\n�� ��������� � ���="+String(Sh.OlePropertyGet("Cells",i,pole_tn_kpe))+"\n���������� ��������� ������ � ����� Excel \n� ��������� �������� ��� ��������������� ������ �������").c_str(),
                                                   "������", MB_OK+MB_ICONWARNING);
                         //������������ ������ � ��������, � ������� �� ����� ���������� ���
                         rtf_Out("zex", String(Sh.OlePropertyGet("Cells",i,6)),1);
                         rtf_Out("tn", String(Sh.OlePropertyGet("Cells",i,pole_tn_kpe)),1);
                         rtf_Out("fio", String(Sh.OlePropertyGet("Cells",i,3))+" (�� ��������� ���� ���)" ,1);

                         if(!rtf_LineFeed())
                           {
                             MessageBox(Handle,"������ ������ � ���� ������","������",8192);
                             if (!rtf_Close()) MessageBox(Handle,"������ �������� ����� ������","������",8192);
                             return;
                           }

                         otchet=1;  //������� ������������ ������ �� ������������� �������
                       }

                     Sql = "update ocenka set \
                                              kpe1="+Main->SetNull(Double(Sh.OlePropertyGet("Cells",i,pole_kpe1))) +", \
                                              kpe2="+Main->SetNull(Double(Sh.OlePropertyGet("Cells",i,pole_kpe2))) +", \
                                              kpe3="+Main->SetNull(Double(Sh.OlePropertyGet("Cells",i,pole_kpe3))) +", \
                                              kpe4="+Main->SetNull(Double(Sh.OlePropertyGet("Cells",i,pole_kpe4))) +" \
                            where tn="+ Sh.OlePropertyGet("Cells",i,pole_tn_kpe)+" and god="+IntToStr(Main->god);

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

                         Main->InsertLog(logi+". �������� ������ ��� ������� ���������� ������ � ������� OCENKA �� ����� '"+FileListBox1->Items->Strings[doc]+"'");
                         DM->qLogs->Requery();
                         DM->qOcenka->Requery();
                         Main->StatusBar1->SimpleText ="�������� ������: "+IntToStr(Main->god)+" ���";
                         Abort();
                       }

                     rec++;
                     kol+=DM->qObnovlenie->RowsAffected;

                     // ���������� ����������� �������
                     if (DM->qObnovlenie->RowsAffected == 0)
                       {
                         //������������ ������ �� ������������� �������
                         rtf_Out("zex", String(Sh.OlePropertyGet("Cells",i,6)),1);
                         rtf_Out("tn", String(Sh.OlePropertyGet("Cells",i,pole_tn_kpe)),1);
                         rtf_Out("fio", String(Sh.OlePropertyGet("Cells",i,3))+" (�� ���������, �������� �������� ���.�)" ,1);

                         if(!rtf_LineFeed())
                           {
                             MessageBox(Handle,"������ ������ � ���� ������","������",8192);
                             if (!rtf_Close()) MessageBox(Handle,"������ �������� ����� ������","������",8192);
                             return;
                           }
                         otchet=1;  //������� ������������ ������ �� ������������� �������
                       }
                   }
                 Main->ProgressBar->Position++;
               }
           }
//�������� �������� ����� �� ������
//******************************************************************************
         else if (RadioButtonVZ->Checked == true)
           {
             for (int i=1; i<Row+1; i++)
               {
                 zex = Sh.OlePropertyGet("Cells",i,pole_tn_vz);

                 if (zex.IsEmpty() || !Proverka(zex))  continue;
                   {
                     //�������� ������ � ����
                     //�������� ������������ ���������� �����
                     if (String(Sh.OlePropertyGet("Cells",i,pole_vz)).IsEmpty())
                       {
                         Application->MessageBox(("�� ������ ������� �� ��������� � ���="+String(Sh.OlePropertyGet("Cells",i,pole_tn_vz))+"\n���������� ��������� ������ � ����� Excel \n� ��������� �������� ��� ��������������� ������ �������").c_str(),
                                                   "������", MB_OK+MB_ICONWARNING);
                         //������������ ������ � ��������, � ������� �� ���������� �������
                         rtf_Out("zex", String(Sh.OlePropertyGet("Cells",i,5)),1);
                         rtf_Out("tn", String(Sh.OlePropertyGet("Cells",i,pole_tn_vz)),1);
                         rtf_Out("fio", String(Sh.OlePropertyGet("Cells",i,2))+" (�� ������ ������� ����� �� ������)" ,1);

                         if(!rtf_LineFeed())
                           {
                             MessageBox(Handle,"������ ������ � ���� ������","������",8192);
                             if (!rtf_Close()) MessageBox(Handle,"������ �������� ����� ������","������",8192);
                             return;
                           }

                         otchet=1;  //������� ������������ ������ �� ������������� �������
                         rec++;
                       }
                     else
                       {
                         Sql = "update ocenka set \
                                                 vz_pens="+QuotedStr(Sh.OlePropertyGet("Cells",i,pole_vz)) +"\
                                where tn="+ Sh.OlePropertyGet("Cells",i,pole_tn_vz)+" and god="+IntToStr(Main->god);

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

                             Main->InsertLog(logi+". �������� ������ ��� ������� ���������� ������ � ������� OCENKA �� ����� '"+FileListBox1->Items->Strings[doc]+"'");
                             DM->qLogs->Requery();
                             DM->qOcenka->Requery();
                             Main->StatusBar1->SimpleText ="�������� ������: "+IntToStr(Main->god)+" ���";
                             Abort();
                           }

                         rec++;
                         kol+=DM->qObnovlenie->RowsAffected;

                         // ���������� ����������� �������
                         if (DM->qObnovlenie->RowsAffected == 0)
                           {
                             //������������ ������ �� ������������� �������
                             rtf_Out("zex", String(Sh.OlePropertyGet("Cells",i,5)),1);
                             rtf_Out("tn", String(Sh.OlePropertyGet("Cells",i,pole_tn_vz)),1);
                             rtf_Out("fio", String(Sh.OlePropertyGet("Cells",i,2))+" (�� ���������� ������, �������� �������� ���.�)" ,1);

                             if(!rtf_LineFeed())
                               {
                                 MessageBox(Handle,"������ ������ � ���� ������","������",8192);
                                 if (!rtf_Close()) MessageBox(Handle,"������ �������� ����� ������","������",8192);
                                 return;
                               }
                             otchet=1;  //������� ������������ ������ �� ������������� �������
                           }
                       }
                   }
                 Main->ProgressBar->Position++;
               }
           }
//�������� ������������ � �� �� ������ �� �������������
//******************************************************************************
         else if (RadioButtonKR->Checked == true)
           {
             for (int i=1; i<Row+1; i++)
               {
                 zex = Sh.OlePropertyGet("Cells",i,13);

                 if (zex.IsEmpty() || !Proverka(zex))  continue;
                   {
                     //�������� ������ � ����
                     //�������� ������������ ���������� �����
                     DM->qProverka->Close();
                     
                     DM->qProverka->Parameters->ParamByName("pgod")->Value=IntToStr(Main->god);
                     DM->qProverka->Parameters->ParamByName("ptn")->Value=String(Sh.OlePropertyGet("Cells",i,13));

                     try
                       {
                         DM->qProverka->Open();
                       }
                     catch(Exception &E)
                       {
                         Application->MessageBox(("�������� ������ ��� ������� �������� ������ � ������� Ocenka" + E.Message).c_str(),"������",
                                                   MB_OK+MB_ICONERROR);

                         Main->InsertLog(logi+". �������� ������ ��� ������� ���������� ������ � ������� OCENKA �� ����� '"+FileListBox1->Items->Strings[doc]+"'");
                         DM->qLogs->Requery();
                         DM->qOcenka->Requery();
                         Main->StatusBar1->SimpleText ="�������� ������: "+IntToStr(Main->god)+" ���";
                         Abort();
                       }

                     //ShowMessage(String(Sh.OlePropertyGet("Cells",i,pole_id_shtat))+"  "+String(Sh.OlePropertyGet("Cells",i,11)));
                     //ShowMessage(String(Sh.OlePropertyGet("Cells",i,pole_kr_zex)));


                     //��� ��-�������� ��������
                     if ((String(Sh.OlePropertyGet("Cells",i,pole_id_shtat)).IsEmpty() && String(Sh.OlePropertyGet("Cells",i,11))=="1") ||
                         (String(Sh.OlePropertyGet("Cells",i,pole_id_shtat)).IsEmpty() && String(Sh.OlePropertyGet("Cells",i,11)).IsEmpty() && !String(Sh.OlePropertyGet("Cells",i,13)).IsEmpty()))
                       {
                         Application->MessageBox(("�� ������ ���� ������� ��������� ����������� ��������� ���.�="+String(Sh.OlePropertyGet("Cells",i,13))+", ���="+String(Sh.OlePropertyGet("Cells",i,16))+"\n���������� ��������� ������ � ����� Excel \n� ��������� �������� ��� ��������������� ������ �������").c_str(),
                                                   "������", MB_OK+MB_ICONWARNING);
                         //������������ ������
                         rtf_Out("zex", String(Sh.OlePropertyGet("Cells",i,16)),1);
                         rtf_Out("tn", String(Sh.OlePropertyGet("Cells",i,13)),1);
                         rtf_Out("fio", String(Sh.OlePropertyGet("Cells",i,12))+" (�� ������ ���.� ����������� ���������)" ,1);

                         if(!rtf_LineFeed())
                           {
                             MessageBox(Handle,"������ ������ � ���� ������","������",8192);
                             if (!rtf_Close()) MessageBox(Handle,"������ �������� ����� ������","������",8192);
                             return;
                           }

                         otchet=1;  //������� ������������ ������ �� ������������� �������
                         rec++;

                       }
                     else if ((String(Sh.OlePropertyGet("Cells",i,pole_kr_zex)).IsEmpty() && String(Sh.OlePropertyGet("Cells",i,11))=="1") ||
                              (String(Sh.OlePropertyGet("Cells",i,pole_kr_zex)).IsEmpty() && String(Sh.OlePropertyGet("Cells",i,11)).IsEmpty() && !String(Sh.OlePropertyGet("Cells",i,13)).IsEmpty()))
                       {
                         //��� ���� ����������� ���������
                         Application->MessageBox(("�� ������ ��� ����������� ��������� �� ��������� ���.�="+String(Sh.OlePropertyGet("Cells",i,13))+", ���="+String(Sh.OlePropertyGet("Cells",i,16))+"\n���������� ��������� ������ � ����� Excel \n� ��������� �������� ��� ��������������� ������ �������").c_str(),
                                                  "������", MB_OK+MB_ICONWARNING);
                         //������������ ������ � ��������, � ������� �� ���������� �������
                         rtf_Out("zex", String(Sh.OlePropertyGet("Cells",i,16)),1);
                         rtf_Out("tn", String(Sh.OlePropertyGet("Cells",i,13)),1);
                         rtf_Out("fio", String(Sh.OlePropertyGet("Cells",i,12))+" (�� ������ ��� ����������� ���������)" ,1);
                         if(!rtf_LineFeed())
                           {
                             MessageBox(Handle,"������ ������ � ���� ������","������",8192);
                             if (!rtf_Close()) MessageBox(Handle,"������ �������� ����� ������","������",8192);
                             return;
                           }

                         otchet=1;  //������� ������������ ������ �� ������������� �������
                         rec++;
                       }

                     else if ((String(Sh.OlePropertyGet("Cells",i,pole_krshifr_dolg)).IsEmpty() && String(Sh.OlePropertyGet("Cells",i,11))=="1") ||
                              (String(Sh.OlePropertyGet("Cells",i,pole_krshifr_dolg)).IsEmpty() && String(Sh.OlePropertyGet("Cells",i,11)).IsEmpty() && !String(Sh.OlePropertyGet("Cells",i,13)).IsEmpty()))
                       {
                         //��� ����� ����������� ���������
                         Application->MessageBox(("�� ������ ���� ���������� ��������� ��������� ���.�="+String(Sh.OlePropertyGet("Cells",i,13))+", ���="+String(Sh.OlePropertyGet("Cells",i,16))+"\n���������� ��������� ������ � ����� Excel \n� ��������� �������� ��� ��������������� ������ �������").c_str(),
                                                  "������", MB_OK+MB_ICONWARNING);
                         //������������ ������ � ��������, � ������� �� ���������� �������
                         rtf_Out("zex", String(Sh.OlePropertyGet("Cells",i,16)),1);
                         rtf_Out("tn", String(Sh.OlePropertyGet("Cells",i,13)),1);
                         rtf_Out("fio", String(Sh.OlePropertyGet("Cells",i,12))+" (�� ������ ���� ��������� ����������� ���������)" ,1);

                         if(!rtf_LineFeed())
                           {
                             MessageBox(Handle,"������ ������ � ���� ������","������",8192);
                             if (!rtf_Close()) MessageBox(Handle,"������ �������� ����� ������","������",8192);
                             return;
                           }

                         otchet=1;  //������� ������������ ������ �� ������������� �������
                         rec++;
                       }
                     //�������� �� ������� ���������� ����������� ��������� � ������� Ocenka
                     else if (DM->qProverka->FieldByName("tn")->AsString.IsEmpty() )
                       {
                         //ShowMessage(String(Sh.OlePropertyGet("Cells",i,13)));
                         Application->MessageBox(("���������� �������� � ��������� ���.�="+String(Sh.OlePropertyGet("Cells",i,13))+" �� ������ � ��������� �� ������ ��������� (OCENKA)").c_str(), "������",
                                                  MB_OK + MB_ICONWARNING);

                         //������������ ������
                         rtf_Out("zex", String(Sh.OlePropertyGet("Cells",i,16)),1);
                         rtf_Out("tn", String(Sh.OlePropertyGet("Cells",i,13)),1);
                         rtf_Out("fio", String(Sh.OlePropertyGet("Cells",i,12))+" (��� ���.� � ��������� �� ������ ���������)" ,1);

                         if(!rtf_LineFeed())
                           {
                             MessageBox(Handle,"������ ������ � ���� ������","������",8192);
                             if (!rtf_Close()) MessageBox(Handle,"������ �������� ����� ������","������",8192);
                             return;
                           }

                         otchet=1;  //������� ������������ ������ �� ������������� �������
                         rec++;
                       }
                     //�������� �� ������� ������ �- ��� � � ����������� ��������� � ������� Ocenka
                     else if (DM->qProverka->FieldByName("kom_reit")->AsString=="B-" || DM->qOcenka->FieldByName("kom_reit")->AsString=="C")
                       {

                         Application->MessageBox(("���������� �������� � ��������� ���.�="+String(Sh.OlePropertyGet("Cells",i,13))+" ����� ������� "+DM->qProverka->FieldByName("kom_reit")->AsString+" �� ������ ���������").c_str(), "������",
                                                  MB_OK + MB_ICONWARNING);
                         //ShowMessage(DM->qObnovlenie->FieldByName("tn")->AsString+" "+DM->qObnovlenie->FieldByName("fio")->AsString+" "+DM->qObnovlenie->FieldByName("kom_reit")->AsString);
                         //������������ ������
                         rtf_Out("zex", String(Sh.OlePropertyGet("Cells",i,16)),1);
                         rtf_Out("tn", String(Sh.OlePropertyGet("Cells",i,13)),1);
                         rtf_Out("fio", String(Sh.OlePropertyGet("Cells",i,12))+(" (������� "+DM->qProverka->FieldByName("kom_reit")->AsString+")").c_str() ,1);

                         if(!rtf_LineFeed())
                           {
                             MessageBox(Handle,"������ ������ � ���� ������","������",8192);
                             if (!rtf_Close()) MessageBox(Handle,"������ �������� ����� ������","������",8192);
                             return;
                           }

                         otchet=1;  //������� ������������ ������ �� ������������� �������
                         rec++;
                       }
                     else
                       {
                         if (String(Sh.OlePropertyGet("Cells",i,11)).IsEmpty()) kol_zam=0;
                         else kol_zam=StrToInt(String(Sh.OlePropertyGet("Cells",i,11)));

                         if (kol_zam>1) n=i-(kol_zam-1);
                         else n=i;

                         //�������� �� ������� ������� � ������� Ocenka
                         if (DM->qProverka->FieldByName("rezerv")->AsInteger!=1)
                           {
                             //���������� ������� � ������� Ocenka
                             Sql = "update ocenka set \
                                                     rezerv=1 \
                                    where tn="+ Sh.OlePropertyGet("Cells",i,13)+" and god="+IntToStr(Main->god);

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

                                 Main->InsertLog(logi+". �������� ������ ��� ������� ���������� ������ � ������� OCENKA �� ����� '"+FileListBox1->Items->Strings[doc]+"'");
                                 DM->qLogs->Requery();
                                 DM->qOcenka->Requery();
                                 Main->StatusBar1->SimpleText ="�������� ������: "+IntToStr(Main->god)+" ���";
                                 Abort();
                               }
                           }

                        /* AnsiString shtat;

                         if (String(Sh.OlePropertyGet("Cells",i,pole_id_shtat)).Length()==7) shtat= "0"+String(Sh.OlePropertyGet("Cells",i,pole_id_shtat));
                         else (String(Sh.OlePropertyGet("Cells",i,pole_id_shtat)).Length()==7) shtat= "0"+String(Sh.OlePropertyGet("Cells",i,pole_id_shtat));
                         else String(Sh.OlePropertyGet("Cells",i,pole_id_shtat));

                         Variant locvalues[] = {String(Sh.OlePropertyGet("Cells",i,13)),String(Sh.OlePropertyGet("Cells",i,pole_id_shtat)) };*/

                         //���� � ���� ��� ���� � ��������� ����� ��_��������, ����� ��������� �� ���� ������
                         Sql = "select * from ocenka_rez where god=:pgod and tn=:ptn and id_shtat=:lpad(pshtat,8,'0')";

                         DM->qRezerv->Close();
                         //DM->qRezerv->Parameters->ParseSQL(DM->qRezerv->SQL->Text, true);
                         DM->qRezerv->Parameters->ParamByName("pgod")->Value=IntToStr(Main->god);
                         DM->qRezerv->Parameters->ParamByName("ptn")->Value=String(Sh.OlePropertyGet("Cells",i,13));
                         DM->qRezerv->Parameters->ParamByName("pshtat")->Value=String(Sh.OlePropertyGet("Cells",n,pole_id_shtat));
                         try
                           {
                             DM->qRezerv->Open();
                           }
                         catch(Exception &E)
                           {
                             Application->MessageBox(("�������� ������ ��� ������� �������� ������ � ������� Ocenka" + E.Message).c_str(),"������",
                                                           MB_OK+MB_ICONERROR);

                             Main->InsertLog(logi+". �������� ������ ��� ������� ���������� ������ � ������� OCENKA �� ����� '"+FileListBox1->Items->Strings[doc]+"'");
                             DM->qLogs->Requery();
                             DM->qOcenka->Requery();
                             DM->qZamesh->Requery();
                             DM->qRezerv->Requery();
                             Main->StatusBar1->SimpleText ="�������� ������: "+IntToStr(Main->god)+" ���";
                             Abort();
                           }

                         if (DM->qRezerv->RecordCount>0)
                           {
                             //���������� ������� � ������� Ocenka
                             Sql = "update ocenka_rez set \
                                                     god = "+IntToStr(Main->god)+",                                                               \
                                                     tn = "+ Sh.OlePropertyGet("Cells",i,13)+",                                                  \
                                                     id_shtat = lpad("+QuotedStr(Sh.OlePropertyGet("Cells",n,pole_id_shtat)) +",8,'0'),\
                                                     dolg_rez = "+QuotedStr(Sh.OlePropertyGet("Cells",n,pole_rezerv_dolg_kr)) +",\
                                                     tn_sap_rez = "+QuotedStr(Sh.OlePropertyGet("Cells",n,pole_tn_kr)) +",\
                                                     fio_rez = "+QuotedStr(Sh.OlePropertyGet("Cells",n,pole_kr_fio)) +",\
                                                     zex_rez = "+QuotedStr(Sh.OlePropertyGet("Cells",n,pole_kr_zex)) +",\
                                                     type=1,\
                                                     shifr_rez = "+QuotedStr(Sh.OlePropertyGet("Cells",n,pole_krshifr_dolg)) +"\
                                    where tn="+ Sh.OlePropertyGet("Cells",i,13)+" and id_shtat=lpad("+QuotedStr(Sh.OlePropertyGet("Cells",n,pole_id_shtat))+",8,'0') and god="+IntToStr(Main->god);
                           }
                         else
                           {
                             //���������� ������ � ������� Ocenka_rez
                             Sql = "insert into ocenka_rez  (god, tn, id_shtat, dolg_rez, tn_sap_rez, fio_rez, zex_rez, type, shifr_rez) \
                                values ("+IntToStr(Main->god)+",                                                               \
                                        "+ Sh.OlePropertyGet("Cells",i,13)+",                                                  \
                                        lpad("+QuotedStr(Sh.OlePropertyGet("Cells",n,pole_id_shtat)) +",8,'0'),\
                                        "+QuotedStr(Sh.OlePropertyGet("Cells",n,pole_rezerv_dolg_kr)) +",\
                                        "+QuotedStr(Sh.OlePropertyGet("Cells",n,pole_tn_kr)) +",\
                                        "+QuotedStr(Sh.OlePropertyGet("Cells",n,pole_kr_fio)) +",\
                                        "+QuotedStr(Sh.OlePropertyGet("Cells",n,pole_kr_zex)) +",\
                                        1,\
                                        "+QuotedStr(Sh.OlePropertyGet("Cells",n,pole_krshifr_dolg)) +"\
                                       )";
                           }

                        /* Sql = "update ocenka set \
                                              rezerv=1, \
                                              shtat_zam="+QuotedStr(Sh.OlePropertyGet("Cells",n,pole_id_shtat)) +",\
                                              dolg_rezerv="+QuotedStr(Sh.OlePropertyGet("Cells",n,pole_rezerv_dolg_kr)) +",\
                                              zex_rez="+QuotedStr(Sh.OlePropertyGet("Cells",n,pole_kr_zex)) +",\
                                              shifr_rez="+QuotedStr(Sh.OlePropertyGet("Cells",n,pole_krshifr_dolg)) +",\
                                              zex_zam="+QuotedStr(Sh.OlePropertyGet("Cells",n,pole_kr_zex)) +",\
                                              shifr_zam="+QuotedStr(Sh.OlePropertyGet("Cells",n,pole_krshifr_dolg)) +",\
                                              tn_sap_zam="+QuotedStr(Sh.OlePropertyGet("Cells",n,pole_tn_kr)) +",\
                                              fio_zam="+QuotedStr(Sh.OlePropertyGet("Cells",n,pole_kr_fio)) +"\
                                where tn="+ Sh.OlePropertyGet("Cells",i,13)+" and god="+IntToStr(Main->god);  */


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

                             Main->InsertLog(logi+". �������� ������ ��� ������� ���������� ������ � ������� OCENKA �� ����� '"+FileListBox1->Items->Strings[doc]+"'");
                             DM->qLogs->Requery();
                             DM->qOcenka->Requery();
                             Main->StatusBar1->SimpleText ="�������� ������: "+IntToStr(Main->god)+" ���";
                             Abort();
                           }

                         rec++;
                         kol+=DM->qObnovlenie->RowsAffected;

                         // ���������� ����������� �������
                         if (DM->qObnovlenie->RowsAffected == 0)
                           {
                             //������������ ������ �� ������������� �������
                             rtf_Out("zex", String(Sh.OlePropertyGet("Cells",i,16)),1);
                             rtf_Out("tn", String(Sh.OlePropertyGet("Cells",i,13)),1);
                             rtf_Out("fio", String(Sh.OlePropertyGet("Cells",i,12))+" (�� ���������� ������, �������� �������� ���.�)" ,1);

                             if(!rtf_LineFeed())
                               {
                                 MessageBox(Handle,"������ ������ � ���� ������","������",8192);
                                 if (!rtf_Close()) MessageBox(Handle,"������ �������� ����� ������","������",8192);
                                 return;
                               }
                             otchet=1;  //������� ������������ ������ �� ������������� �������
                           }
                       }
                   }
                 Main->ProgressBar->Position++;
               }
           }


          //�������� Excel
          AppEx.OleProcedure("Quit");

          Main->StatusBar1->SimpleText="�������� ������... �� "+zex+" ���� ��������� " + IntToStr(kol) + " �� " + IntToStr(rec) + " �������";
          ob_kol+= kol;
          obnov_kol+= rec;
          doc++;

          DM->qOcenka->Requery();
          DM->qZamesh->Requery();
          DM->qRezerv->Requery();
        }
      catch (...)
        {
          AppEx.OleProcedure("Quit");
          AppEx = Unassigned;
          

          if(!rtf_Close())
            {
              MessageBox(Handle,"������ �������� ����� ������", "������", 8192);
              return;
            }

          DM->qZamesh->Requery();
          DM->qRezerv->Requery();

          Cursor = crDefault;
          Main->ProgressBar->Position = 0;
          Main->ProgressBar->Visible = false;
          Main->StatusBar1->SimpleText ="�������� ������: "+IntToStr(Main->god)+" ���";
          Abort();
        }
    }

  if(!rtf_Close())
    {
      MessageBox(Handle,"������ �������� ����� ������", "������", 8192);
      return;
    }

  // ������������ ������, ���� ���� �� ����������� ������
  if (otchet==1)
    {
      Main->StatusBar1->SimpleText = " ������������ ������ � �������������� ��������...";
      //�������� �����, ���� �� �� ����������
      ForceDirectories(Main->WorkPath);

      int istrd;
      try
        {
          rtf_CreateReport(Main->TempPath +"\\zagruzka.txt", Main->Path+"\\RTF\\zagruzka.rtf",
                           Main->WorkPath+"\\�� ����������� ������.doc",NULL,&istrd);

          WinExec(("\""+ Main->WordPath+"\"\""+Main->WorkPath+"\\�� ����������� ������.doc\"").c_str(),SW_MAXIMIZE);
          DeleteFile(Main->TempPath+"\\zagruzka.txt");
        }
      catch(RepoRTF_Error E)
        {
          MessageBox(Handle,("������ ������������ ������:"+ AnsiString(E.Err)+
                             "\n������ ����� ������:"+IntToStr(istrd)).c_str(),"������",8192);
        }

      doc=doc-1;
      Main->InsertLog(logi+" ��������� �� ����� '"+FileListBox1->Items->Strings[doc]+"'. ���� �� ����������� ������!");
      DM->qLogs->Requery();

    }
  else
    {
      doc=doc-1;
      Main->InsertLog(logi+" ��������� ������� �� ����� '"+FileListBox1->Items->Strings[doc]+"'");
      DM->qLogs->Requery();
    }

  Cursor = crDefault;
  Main->ProgressBar->Position = 0;
  Main->ProgressBar->Visible = false;
  Main->StatusBar1->SimpleText ="�������� ������: "+IntToStr(Main->god)+" ���";

 /* //������� �����
  if (RadioButtonDATAO->Checked) InsertLog("����������� ����� �� ����������(������): ��������� " + IntToStr(kol) + " �� " + IntToStr(rec) + " �������.");
  else if (RadioButtonOCENKA->Checked) InsertLog("����������� ����� �� ����������(������): ��������� " + IntToStr(kol) + " �� " + IntToStr(rec) + " �������.");
  else if (RadioButtonREZERV->Checked  InsertLog("����������� ����� �� ����������(������): ��������� " + IntToStr(kol) + " �� " + IntToStr(rec) + " �������.");
  */

  Application->MessageBox(("�������� ������ ��������� ������� =) \n��������� " + IntToStr(ob_kol) + " �� " + IntToStr(obnov_kol)+" �������").c_str(),
                           "���������� ������ �� ������ ���������",
                           MB_OK + MB_ICONINFORMATION);
}
//---------------------------------------------------------------------------

void __fastcall TZagruzka::FormShow(TObject *Sender)
{
  //������� ����� ��� �������� �� Excel
  EditDATA->Text = "";
  EditFIO->Text = "";
  EditDOLGO->Text = "";
  EditOCENKA->Text = "";
  EditREZERV->Text = "";
  EditDOLG->Text = "";
  EditZEX->Text = "";
  EditTN->Text = "";
  EditREZULT_OCEN->Text = "";
  EditKPE_OCEN->Text = "";
  EditKOMP_OCEN->Text = "";
  EditFIOEOP->Text = "";
  EditTNEOP->Text = "";
  EditTN_KPE->Text = "";
  EditKPE1->Text = "";
  EditKPE2->Text = "";
  EditKPE3->Text = "";
  EditKPE4->Text = "";
  EditTN_VZ->Text = "";
  EditVZ->Text = "";

  EditKR_ZEX->Text = "";
  EditTN_KR->Text = "";
  EditKR_FIO->Text = "";
  EditKRSHIFR_DOLG->Text = "";

  RadioButtonDATAO->Checked = true;

}
//---------------------------------------------------------------------------

void __fastcall TZagruzka::EditDATAKeyPress(TObject *Sender, char &Key)
{
  if (!(IsNumeric(Key)||Key=='\b')) Key=0;
}
//---------------------------------------------------------------------------

void __fastcall TZagruzka::CheckBoxREZERVClick(TObject *Sender)
{
  RadioButtonDATAO->Checked = false;
  RadioButtonEOP->Checked = false;
}
//---------------------------------------------------------------------------

void __fastcall TZagruzka::CheckBoxOCENKAClick(TObject *Sender)
{
  RadioButtonDATAO->Checked = false;
  RadioButtonEOP->Checked = false;        
}
//---------------------------------------------------------------------------

void __fastcall TZagruzka::RadioButtonDATAOClick(TObject *Sender)
{
  CheckBoxOCENKA->Checked = false;
  CheckBoxREZERV->Checked = false;
}
//---------------------------------------------------------------------------

void __fastcall TZagruzka::RadioButtonEOPClick(TObject *Sender)
{
  CheckBoxOCENKA->Checked = false;
  CheckBoxREZERV->Checked = false;
}
//---------------------------------------------------------------------------

