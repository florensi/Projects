//---------------------------------------------------------------------------

#include <vcl.h>
#pragma hdrstop

#include "uDM.h"
#include "uMain.h"
#include "uSprav.h"
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

  //���������� ������ ���������� �� �������������� �����
  if (!DecryptFromFile(GetCurrentDir() + "\\certificate.1.2.m.dat", S))
    {
      Application->MessageBox("���������� �������� ������ ���������� � ����� ������","������ ����������",
                               MB_OK + MB_ICONERROR);
      Application->Terminate();
      Abort();
    }

  ADOConnection1->ConnectionString = S;

  // ���������� � ����� ������
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

  qPrazdDniVihodnue->Active = true;

  Application->UpdateFormatSettings = false;

  DecimalSeparator='.';
  DateSeparator='.';

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
}
//---------------------------------------------------------------------------

void __fastcall TDM::DataModuleDestroy(TObject *Sender)
{
  ADOConnection1->Close();
  qObnovlenie->Active = false;
  qObnovlenie2->Active = false;
  qGrafik->Active = false;
  qOgraf->Active = false;
  qPrazdDni->Active = false;
  qPrdPrazdDni->Active = false;
  qPrazdDniVihodnue->Active = false;
  qNorma11Graf->Active = false;
  qSprav->Active = false;

 // delete Main->ProgressBar;
}
//---------------------------------------------------------------------------

void __fastcall TDM::qGrafikCalcFields(TDataSet *DataSet)
{
  //����� �������� ������������ � 26 ����� �������� ������
  if (DM->qOgraf->FieldByName("ograf")->AsString==100 ||
      DM->qOgraf->FieldByName("ograf")->AsString==315 ||
      DM->qOgraf->FieldByName("ograf")->AsString==400 ||
      DM->qOgraf->FieldByName("ograf")->AsString==775 ||
      DM->qOgraf->FieldByName("ograf")->AsString==780)
    {
      Main->DBGridEh1->Columns->Items[2]->Title->Caption = "����� ������|26";
      Main->DBGridEh1->Columns->Items[3]->Title->Caption = "����� ������|27";
      Main->DBGridEh1->Columns->Items[4]->Title->Caption = "����� ������|28";
      Main->DBGridEh1->Columns->Items[5]->Title->Caption = "����� ������|29";
      Main->DBGridEh1->Columns->Items[6]->Title->Caption = "����� ������|30";
      Main->DBGridEh1->Columns->Items[7]->Title->Caption = "����� ������|31";
      Main->DBGridEh1->Columns->Items[8]->Title->Caption = "����� ������|1";
      Main->DBGridEh1->Columns->Items[9]->Title->Caption = "����� ������|2";
      Main->DBGridEh1->Columns->Items[10]->Title->Caption = "����� ������|3";
      Main->DBGridEh1->Columns->Items[11]->Title->Caption = "����� ������|4";
      Main->DBGridEh1->Columns->Items[12]->Title->Caption = "����� ������|5";
      Main->DBGridEh1->Columns->Items[13]->Title->Caption = "����� ������|6";
      Main->DBGridEh1->Columns->Items[14]->Title->Caption = "����� ������|7";
      Main->DBGridEh1->Columns->Items[15]->Title->Caption = "����� ������|8";
      Main->DBGridEh1->Columns->Items[16]->Title->Caption = "����� ������|9";
      Main->DBGridEh1->Columns->Items[17]->Title->Caption = "����� ������|10";
      Main->DBGridEh1->Columns->Items[18]->Title->Caption = "����� ������|11";
      Main->DBGridEh1->Columns->Items[19]->Title->Caption = "����� ������|12";
      Main->DBGridEh1->Columns->Items[20]->Title->Caption = "����� ������|13";
      Main->DBGridEh1->Columns->Items[21]->Title->Caption = "����� ������|14";
      Main->DBGridEh1->Columns->Items[22]->Title->Caption = "����� ������|15";
      Main->DBGridEh1->Columns->Items[23]->Title->Caption = "����� ������|16";
      Main->DBGridEh1->Columns->Items[24]->Title->Caption = "����� ������|17";
      Main->DBGridEh1->Columns->Items[25]->Title->Caption = "����� ������|18";
      Main->DBGridEh1->Columns->Items[26]->Title->Caption = "����� ������|19";
      Main->DBGridEh1->Columns->Items[27]->Title->Caption = "����� ������|20";
      Main->DBGridEh1->Columns->Items[28]->Title->Caption = "����� ������|21";
      Main->DBGridEh1->Columns->Items[29]->Title->Caption = "����� ������|22";
      Main->DBGridEh1->Columns->Items[30]->Title->Caption = "����� ������|23";
      Main->DBGridEh1->Columns->Items[31]->Title->Caption = "����� ������|24";
      Main->DBGridEh1->Columns->Items[32]->Title->Caption = "����� ������|25";

      if (DM->qOgraf->FieldByName("otchet")->AsInteger==1)
        {
          qGrafikf1->Value = DM->qGrafik->FieldByName("nsm26")->AsString;
          qGrafikf2->Value = DM->qGrafik->FieldByName("nsm27")->AsString;
          qGrafikf3->Value = DM->qGrafik->FieldByName("nsm28")->AsString;
          qGrafikf4->Value = DM->qGrafik->FieldByName("nsm29")->AsString;
          qGrafikf5->Value = DM->qGrafik->FieldByName("nsm30")->AsString;
          qGrafikf6->Value = DM->qGrafik->FieldByName("nsm31")->AsString;
          qGrafikf7->Value = DM->qGrafik->FieldByName("nsm1")->AsString;
          qGrafikf8->Value = DM->qGrafik->FieldByName("nsm2")->AsString;
          qGrafikf9->Value = DM->qGrafik->FieldByName("nsm3")->AsString;
          qGrafikf10->Value = DM->qGrafik->FieldByName("nsm4")->AsString;
          qGrafikf11->Value = DM->qGrafik->FieldByName("nsm5")->AsString;
          qGrafikf12->Value = DM->qGrafik->FieldByName("nsm6")->AsString;
          qGrafikf13->Value = DM->qGrafik->FieldByName("nsm7")->AsString;
          qGrafikf14->Value = DM->qGrafik->FieldByName("nsm8")->AsString;
          qGrafikf15->Value = DM->qGrafik->FieldByName("nsm9")->AsString;
          qGrafikf16->Value = DM->qGrafik->FieldByName("nsm10")->AsString;
          qGrafikf17->Value = DM->qGrafik->FieldByName("nsm11")->AsString;
          qGrafikf18->Value = DM->qGrafik->FieldByName("nsm12")->AsString;
          qGrafikf19->Value = DM->qGrafik->FieldByName("nsm13")->AsString;
          qGrafikf20->Value = DM->qGrafik->FieldByName("nsm14")->AsString;
          qGrafikf21->Value = DM->qGrafik->FieldByName("nsm15")->AsString;
          qGrafikf22->Value = DM->qGrafik->FieldByName("nsm16")->AsString;
          qGrafikf23->Value = DM->qGrafik->FieldByName("nsm17")->AsString;
          qGrafikf24->Value = DM->qGrafik->FieldByName("nsm18")->AsString;
          qGrafikf25->Value = DM->qGrafik->FieldByName("nsm19")->AsString;
          qGrafikf26->Value = DM->qGrafik->FieldByName("nsm20")->AsString;
          qGrafikf27->Value = DM->qGrafik->FieldByName("nsm21")->AsString;
          qGrafikf28->Value = DM->qGrafik->FieldByName("nsm22")->AsString;
          qGrafikf29->Value = DM->qGrafik->FieldByName("nsm23")->AsString;
          qGrafikf30->Value = DM->qGrafik->FieldByName("nsm24")->AsString;
          qGrafikf31->Value = DM->qGrafik->FieldByName("nsm25")->AsString;
        }
      else
        {
          qGrafikf1->Value = DM->qGrafik->FieldByName("chf26")->AsString;
          qGrafikf2->Value = DM->qGrafik->FieldByName("chf27")->AsString;
          qGrafikf3->Value = DM->qGrafik->FieldByName("chf28")->AsString;
          qGrafikf4->Value = DM->qGrafik->FieldByName("chf29")->AsString;
          qGrafikf5->Value = DM->qGrafik->FieldByName("chf30")->AsString;
          qGrafikf6->Value = DM->qGrafik->FieldByName("chf31")->AsString;
          qGrafikf7->Value = DM->qGrafik->FieldByName("chf1")->AsString;
          qGrafikf8->Value = DM->qGrafik->FieldByName("chf2")->AsString;
          qGrafikf9->Value = DM->qGrafik->FieldByName("chf3")->AsString;
          qGrafikf10->Value = DM->qGrafik->FieldByName("chf4")->AsString;
          qGrafikf11->Value = DM->qGrafik->FieldByName("chf5")->AsString;
          qGrafikf12->Value = DM->qGrafik->FieldByName("chf6")->AsString;
          qGrafikf13->Value = DM->qGrafik->FieldByName("chf7")->AsString;
          qGrafikf14->Value = DM->qGrafik->FieldByName("chf8")->AsString;
          qGrafikf15->Value = DM->qGrafik->FieldByName("chf9")->AsString;
          qGrafikf16->Value = DM->qGrafik->FieldByName("chf10")->AsString;
          qGrafikf17->Value = DM->qGrafik->FieldByName("chf11")->AsString;
          qGrafikf18->Value = DM->qGrafik->FieldByName("chf12")->AsString;
          qGrafikf19->Value = DM->qGrafik->FieldByName("chf13")->AsString;
          qGrafikf20->Value = DM->qGrafik->FieldByName("chf14")->AsString;
          qGrafikf21->Value = DM->qGrafik->FieldByName("chf15")->AsString;
          qGrafikf22->Value = DM->qGrafik->FieldByName("chf16")->AsString;
          qGrafikf23->Value = DM->qGrafik->FieldByName("chf17")->AsString;
          qGrafikf24->Value = DM->qGrafik->FieldByName("chf18")->AsString;
          qGrafikf25->Value = DM->qGrafik->FieldByName("chf19")->AsString;
          qGrafikf26->Value = DM->qGrafik->FieldByName("chf20")->AsString;
          qGrafikf27->Value = DM->qGrafik->FieldByName("chf21")->AsString;
          qGrafikf28->Value = DM->qGrafik->FieldByName("chf22")->AsString;
          qGrafikf29->Value = DM->qGrafik->FieldByName("chf23")->AsString;
          qGrafikf30->Value = DM->qGrafik->FieldByName("chf24")->AsString;
          qGrafikf31->Value = DM->qGrafik->FieldByName("chf25")->AsString;
        }
    }
  // ����� �������� ������������ � 1 ����� ������
  else
    {
      Main->DBGridEh1->Columns->Items[2]->Title->Caption = "����� ������|1";
      Main->DBGridEh1->Columns->Items[3]->Title->Caption = "����� ������|2";
      Main->DBGridEh1->Columns->Items[4]->Title->Caption = "����� ������|3";
      Main->DBGridEh1->Columns->Items[5]->Title->Caption = "����� ������|4";
      Main->DBGridEh1->Columns->Items[6]->Title->Caption = "����� ������|5";
      Main->DBGridEh1->Columns->Items[7]->Title->Caption = "����� ������|6";
      Main->DBGridEh1->Columns->Items[8]->Title->Caption = "����� ������|7";
      Main->DBGridEh1->Columns->Items[9]->Title->Caption = "����� ������|8";
      Main->DBGridEh1->Columns->Items[10]->Title->Caption = "����� ������|9";
      Main->DBGridEh1->Columns->Items[11]->Title->Caption = "����� ������|10";
      Main->DBGridEh1->Columns->Items[12]->Title->Caption = "����� ������|11";
      Main->DBGridEh1->Columns->Items[13]->Title->Caption = "����� ������|12";
      Main->DBGridEh1->Columns->Items[14]->Title->Caption = "����� ������|13";
      Main->DBGridEh1->Columns->Items[15]->Title->Caption = "����� ������|14";
      Main->DBGridEh1->Columns->Items[16]->Title->Caption = "����� ������|15";
      Main->DBGridEh1->Columns->Items[17]->Title->Caption = "����� ������|16";
      Main->DBGridEh1->Columns->Items[18]->Title->Caption = "����� ������|17";
      Main->DBGridEh1->Columns->Items[19]->Title->Caption = "����� ������|18";
      Main->DBGridEh1->Columns->Items[20]->Title->Caption = "����� ������|19";
      Main->DBGridEh1->Columns->Items[21]->Title->Caption = "����� ������|20";
      Main->DBGridEh1->Columns->Items[22]->Title->Caption = "����� ������|21";
      Main->DBGridEh1->Columns->Items[23]->Title->Caption = "����� ������|22";
      Main->DBGridEh1->Columns->Items[24]->Title->Caption = "����� ������|23";
      Main->DBGridEh1->Columns->Items[25]->Title->Caption = "����� ������|24";
      Main->DBGridEh1->Columns->Items[26]->Title->Caption = "����� ������|25";
      Main->DBGridEh1->Columns->Items[27]->Title->Caption = "����� ������|26";
      Main->DBGridEh1->Columns->Items[28]->Title->Caption = "����� ������|27";
      Main->DBGridEh1->Columns->Items[29]->Title->Caption = "����� ������|28";
      Main->DBGridEh1->Columns->Items[30]->Title->Caption = "����� ������|29";
      Main->DBGridEh1->Columns->Items[31]->Title->Caption = "����� ������|30";
      Main->DBGridEh1->Columns->Items[32]->Title->Caption = "����� ������|31";

      if (DM->qOgraf->FieldByName("otchet")->AsInteger==1)
        {
          qGrafikf1->Value = DM->qGrafik->FieldByName("nsm1")->AsString;
          qGrafikf2->Value = DM->qGrafik->FieldByName("nsm2")->AsString;
          qGrafikf3->Value = DM->qGrafik->FieldByName("nsm3")->AsString;
          qGrafikf4->Value = DM->qGrafik->FieldByName("nsm4")->AsString;
          qGrafikf5->Value = DM->qGrafik->FieldByName("nsm5")->AsString;
          qGrafikf6->Value = DM->qGrafik->FieldByName("nsm6")->AsString;
          qGrafikf7->Value = DM->qGrafik->FieldByName("nsm7")->AsString;
          qGrafikf8->Value = DM->qGrafik->FieldByName("nsm8")->AsString;
          qGrafikf9->Value = DM->qGrafik->FieldByName("nsm9")->AsString;
          qGrafikf10->Value = DM->qGrafik->FieldByName("nsm10")->AsString;
          qGrafikf11->Value = DM->qGrafik->FieldByName("nsm11")->AsString;
          qGrafikf12->Value = DM->qGrafik->FieldByName("nsm12")->AsString;
          qGrafikf13->Value = DM->qGrafik->FieldByName("nsm13")->AsString;
          qGrafikf14->Value = DM->qGrafik->FieldByName("nsm14")->AsString;
          qGrafikf15->Value = DM->qGrafik->FieldByName("nsm15")->AsString;
          qGrafikf16->Value = DM->qGrafik->FieldByName("nsm16")->AsString;
          qGrafikf17->Value = DM->qGrafik->FieldByName("nsm17")->AsString;
          qGrafikf18->Value = DM->qGrafik->FieldByName("nsm18")->AsString;
          qGrafikf19->Value = DM->qGrafik->FieldByName("nsm19")->AsString;
          qGrafikf20->Value = DM->qGrafik->FieldByName("nsm20")->AsString;
          qGrafikf21->Value = DM->qGrafik->FieldByName("nsm21")->AsString;
          qGrafikf22->Value = DM->qGrafik->FieldByName("nsm22")->AsString;
          qGrafikf23->Value = DM->qGrafik->FieldByName("nsm23")->AsString;
          qGrafikf24->Value = DM->qGrafik->FieldByName("nsm24")->AsString;
          qGrafikf25->Value = DM->qGrafik->FieldByName("nsm25")->AsString;
          qGrafikf26->Value = DM->qGrafik->FieldByName("nsm26")->AsString;
          qGrafikf27->Value = DM->qGrafik->FieldByName("nsm27")->AsString;
          qGrafikf28->Value = DM->qGrafik->FieldByName("nsm28")->AsString;
          qGrafikf29->Value = DM->qGrafik->FieldByName("nsm29")->AsString;
          qGrafikf30->Value = DM->qGrafik->FieldByName("nsm30")->AsString;
          qGrafikf31->Value = DM->qGrafik->FieldByName("nsm31")->AsString;
        }
      else
        {
          qGrafikf1->Value = DM->qGrafik->FieldByName("chf1")->AsString;
          qGrafikf2->Value = DM->qGrafik->FieldByName("chf2")->AsString;
          qGrafikf3->Value = DM->qGrafik->FieldByName("chf3")->AsString;
          qGrafikf4->Value = DM->qGrafik->FieldByName("chf4")->AsString;
          qGrafikf5->Value = DM->qGrafik->FieldByName("chf5")->AsString;
          qGrafikf6->Value = DM->qGrafik->FieldByName("chf6")->AsString;
          qGrafikf7->Value = DM->qGrafik->FieldByName("chf7")->AsString;
          qGrafikf8->Value = DM->qGrafik->FieldByName("chf8")->AsString;
          qGrafikf9->Value = DM->qGrafik->FieldByName("chf9")->AsString;
          qGrafikf10->Value = DM->qGrafik->FieldByName("chf10")->AsString;
          qGrafikf11->Value = DM->qGrafik->FieldByName("chf11")->AsString;
          qGrafikf12->Value = DM->qGrafik->FieldByName("chf12")->AsString;
          qGrafikf13->Value = DM->qGrafik->FieldByName("chf13")->AsString;
          qGrafikf14->Value = DM->qGrafik->FieldByName("chf14")->AsString;
          qGrafikf15->Value = DM->qGrafik->FieldByName("chf15")->AsString;
          qGrafikf16->Value = DM->qGrafik->FieldByName("chf16")->AsString;
          qGrafikf17->Value = DM->qGrafik->FieldByName("chf17")->AsString;
          qGrafikf18->Value = DM->qGrafik->FieldByName("chf18")->AsString;
          qGrafikf19->Value = DM->qGrafik->FieldByName("chf19")->AsString;
          qGrafikf20->Value = DM->qGrafik->FieldByName("chf20")->AsString;
          qGrafikf21->Value = DM->qGrafik->FieldByName("chf21")->AsString;
          qGrafikf22->Value = DM->qGrafik->FieldByName("chf22")->AsString;
          qGrafikf23->Value = DM->qGrafik->FieldByName("chf23")->AsString;
          qGrafikf24->Value = DM->qGrafik->FieldByName("chf24")->AsString;
          qGrafikf25->Value = DM->qGrafik->FieldByName("chf25")->AsString;
          qGrafikf26->Value = DM->qGrafik->FieldByName("chf26")->AsString;
          qGrafikf27->Value = DM->qGrafik->FieldByName("chf27")->AsString;
          qGrafikf28->Value = DM->qGrafik->FieldByName("chf28")->AsString;
          qGrafikf29->Value = DM->qGrafik->FieldByName("chf29")->AsString;
          qGrafikf30->Value = DM->qGrafik->FieldByName("chf30")->AsString;
          qGrafikf31->Value = DM->qGrafik->FieldByName("chf31")->AsString;
        }
    }
}
//---------------------------------------------------------------------------



