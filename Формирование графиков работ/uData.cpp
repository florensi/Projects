//---------------------------------------------------------------------------

#include <vcl.h>
#pragma hdrstop

#include "uData.h"
#include "uMain.h"
#include "uDM.h"
//---------------------------------------------------------------------------
#pragma package(smart_init)
#pragma resource "*.dfm"
TData *Data;
//---------------------------------------------------------------------------
__fastcall TData::TData(TComponent* Owner)
        : TForm(Owner)
{
}
//---------------------------------------------------------------------------

void __fastcall TData::btnViborKeyDown(TObject *Sender, WORD &Key,
      TShiftState Shift)
{
  if (Key == VK_RETURN)
  FindNextControl((TWinControl *)Sender, true, true,
                   false)->SetFocus();         
}
//---------------------------------------------------------------------------

void __fastcall TData::FormShow(TObject *Sender)
{
  TDateTime dt;

 //����� ��������� ���� � DateTimePicker
 dt = TDateTime( "01.01." + IntToStr(Main->god));
 Data->DateTimePicker1->Date = dt;
}
//---------------------------------------------------------------------------

void __fastcall TData::btnViborClick(TObject *Sender)
{
  Word Year, Month, Day;

  //���������� ��������� ���� �� DateTimePicker
  DecodeDate(Data->DateTimePicker1->Date,Year, Month, Day);
  Main->god = Year;

  //�������� �� ������� ������ �� ��������� ��� � ������� SPGRAFIKI
  AnsiString Sql ="select distinct ograf \
                   from spograf \
                   where ograf not in (select ograf \
                                       from (select ograf, mes  \
                                             from spgrafiki \
                                             where god="+IntToStr(Main->god)+" group by ograf, mes) \
                                       group  by ograf having count(*)=1) order by ograf";
  DM->qObnovlenie->Close();
  DM->qObnovlenie->SQL->Clear();
  DM->qObnovlenie->SQL->Add(Sql);
  try
    {
      DM->qObnovlenie->Open();
    }
  catch(...)
    {
      Application->MessageBox("�� �������� ������� ������ �� ������� SPGRAFIKI","������",
                              MB_OK + MB_ICONERROR);
      Abort();
    }

   Main->ComboBox1->Items->Clear();
   while(!DM->qObnovlenie->Eof)
     {
       Main->ComboBox1->Items->Add(DM->qObnovlenie->FieldByName("ograf")->AsString);
       DM->qObnovlenie->Next();
     }
   Main->ComboBox1->ItemIndex = -1;
                             
  if (Main->god>=2013 && DM->qObnovlenie->RecordCount>0)
    {
      DM->qGrafik->Close();
      Main->DBGridEh1->Enabled = false;
      Main->ComboBox1->ItemIndex = -1;
      Main->StatusBar1->SimpleText="�������� ������:  "+IntToStr(Main->god)+" ���";

      //����������� ���
      DM->qPrazdDni->Close();
      DM->qPrazdDni->Parameters->ParamByName("pgod")->Value = Main->god;
      try
        {
          DM->qPrazdDni->Open();
        }
      catch(...)
        {
          Application->MessageBox("�������� ������ ��� ��������� � ����������� ����������� ����","������",
                                   MB_OK+MB_ICONERROR);
          Abort();
        }

      //��������������� ���
      DM->qPrdPrazdDni->Close();
      DM->qPrdPrazdDni->Parameters->ParamByName("pgod")->Value = Main->god;
      try
        {
          DM->qPrdPrazdDni->Open();
        }
      catch(...)
        {
          Application->MessageBox("�������� ������ ��� ��������� � ����������� ����������� ����","������",
                                   MB_OK+MB_ICONERROR);
          Abort();
        }

      //����������� ���� �������� �� ������/������ �����
      TDateTime data;
      Word year, month, day;

      // ���� � �����
      data = DateToStr(EncodeDateMonthWeek(Main->god,3,4,6));
      DecodeDate(data, year, month, day);
      Main->day_mart = day;

      //��� 40 � 90 �������, ������ �����, ���� � �����
      if (Main->day_mart==31)
        {
          Main->mes_mart2=4;
          Main->day_mart2=1;
        }
      else
        {
          Main->mes_mart2=3;
          Main->day_mart2=Main->day_mart+1;
        }

      //���� � �������
      data = DateToStr(EncodeDateMonthWeek(Main->god,10,4,6));
      DecodeDate(data, year, month, day);
      Main->day_oktyabr = day;

      //��� 40 � 90 �������, ������ �����, ���� � �������
      if (Main->day_oktyabr==31)
        {
          Main->mes_oktyabr2=11;
          Main->day_oktyabr2=1;
        }
      else
        {
          Main->mes_oktyabr2=10;
          Main->day_oktyabr2=Main->day_oktyabr+1;
        }

      Application->MessageBox(("�������� ������ �������!!!\n����������� �������� �� "+IntToStr(Main->god)+" ���").c_str(),"������� ������", MB_OK+MB_ICONINFORMATION);


      //�������������� �������, ���� ��� ���������
      if (Main->god < Main->grafr) Main->redakt=0;
      else Main->redakt=1;
    }
  else
    {
      Application->MessageBox("��� ������ �� ��������� ���","��������������",
                              MB_OK + MB_ICONINFORMATION);
    }
}
//---------------------------------------------------------------------------

