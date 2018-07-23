//---------------------------------------------------------------------------
#define NO_WIN32_LEAN_AND_MEAN
#include <vcl.h>
#pragma hdrstop


#include "uMain.h"
#include "uDM.h"
#include "RepoRTFM.h"
#include "RepoRTFO.h"
#include "uVvod.h"
#include "FuncUserXE.h"
#include "uSprav.h"

#include "EhLibDAC.hpp"
//---------------------------------------------------------------------------
#pragma package(smart_init)
#pragma link "DBAccess"
#pragma link "MemDS"
#pragma link "OracleUniProvider"
#pragma link "Uni"
#pragma link "UniProvider"
#pragma link "DBAxisGridsEh"
#pragma link "DBGridEh"
#pragma link "DBGridEhGrouping"
#pragma link "DBGridEhToolCtrls"
#pragma link "DynVarsEh"
#pragma link "EhLibVCL"
#pragma link "GridsEh"
#pragma link "ToolCtrlsEh"

#pragma link "EhLibDAC"
#pragma resource "*.dfm"
TMain *Main;
//---------------------------------------------------------------------------
__fastcall TMain::TMain(TComponent* Owner)
	: TForm(Owner)
{
}
//---------------------------------------------------------------------------
//MultiByteToWideChar
//WideCharToMultiByte

//---------------------------------------------------------------------------
void __fastcall TMain::FormCreate(TObject *Sender)
{
  // ��������� ������ � ������������ �� ������
  TStringList *SL_Groups = new TStringList();
  // TStringList *SL_Groups2 = new TStringList();

  // ��������� ������ � ������������ �� ������
  // ���������� UserName, DomainName, UserFullName ������ ���� ��������� ��� AnsiString
  if (!GetFullUserInfo(UserName, DomainName, UserFullName))
	{
	  MessageBox(Handle,L"������ ��������� ������ � ������������",L"������",8208);
	  Application->Terminate();
	  Abort();
	}

  //��������� ����� ������� �� ��
  if (!GetUserGroups(UserName, DomainName, SL_Groups))
	{
	  MessageBox(Handle,L"������ ��������� ������ � ������������",L"������",8208);
	  Application->Terminate();
	  Abort();
	}

  //�������� �� ������ � ������
  if ((SL_Groups->IndexOf("mmk-itsvc-hocn-admin")<=-1) && (SL_Groups->IndexOf("mmk-itsvc-hocn")<=-1))
	{
	  MessageBox(Handle,L"� ��� ��� ���� ��� ������ �\n���������� '������������� �������������'!!!",L"����� �������",8208);
	  Application->Terminate();
	  Abort();
	}
  Prava = "ocen";

/*  //�������� ����
  //���� ����� ������ ������ ���������
  if (SL_Groups->IndexOf("mmk-itsvc-hrru-ocen")>-1)
	{
	  DBGridEh1->Columns->Items[6]->Visible = true;     //������� �����������
	  DBGridEh1->Columns->Items[7]->Visible = true;     //������
	  DBGridEh1->Columns->Items[9]->Visible = true;     //���������������� �������, ����
	  DBGridEh1->Columns->Items[10]->Visible = true;    //���������������� �������, ����
	  DBGridEh1->Columns->Items[11]->Visible = true;    //���, ����
	  DBGridEh1->Columns->Items[12]->Visible = true;    //���, ����
	  DBGridEh1->Columns->Items[13]->Visible = true;    //����������, ����������
	  DBGridEh1->Columns->Items[14]->Visible = true;    //����������, ����
	  DBGridEh1->Columns->Items[15]->Visible = true;    //���������, ����������
	  DBGridEh1->Columns->Items[16]->Visible = true;    //���������, ����
	  DBGridEh1->Columns->Items[17]->Visible = true;    //��������������, �������
	  DBGridEh1->Columns->Items[18]->Visible = true;    //��������������, ����
	  DBGridEh1->Columns->Items[19]->Visible = true;    //5�, �������
	  DBGridEh1->Columns->Items[20]->Visible = true;    //5�, ����
	  DBGridEh1->Columns->Items[21]->Visible = true;    //���, ����������
	  DBGridEh1->Columns->Items[22]->Visible = true;    //���, ����
	  DBGridEh1->Columns->Items[23]->Visible = true;    //���, ���-�� �������
	  DBGridEh1->Columns->Items[24]->Visible = true;    //���, ����
	  DBGridEh1->Columns->Items[25]->Visible = true;    //��, �� � �
	  DBGridEh1->Columns->Items[26]->Visible = true;    //��������� ��
	  DBGridEh1->Columns->Items[27]->Visible = true;    //�������� ����������

	  //������� ����
	  N6Spisok->Visible = true;     //�������� ������ ���������� ��� �������������
	  N7PP->Visible = true;         //�������� ����������������� �������
	  N8KPE->Visible = true;        //�������� ���
	  N51C5->Visible = true;        //�������� 5 ������
	  N10PREEM->Visible = true;     //�������� ����������
	  N11SPP->Visible = true;       //�������� ���
	  N12OT->Visible = true;        //�������� ��
	  N13NOT->Visible = true;       //�������� ��������� �� ��
	  N14TD->Visible = true;        //�������� �������� ����������
	  N3Otchet->Visible = true;     //������������ ��������� ������
	  NSprav->Visible = true;        //���������� �� ����������������� �������

	  //����������� ����
	  N16Dobav->Visible = true;     //���������� ������
	  N17Redact->Visible = true;    //�������������� ������
	  N6Reiting->Visible = true;    //������ ��������

	  //������
	  SpeedButton1->Visible = true; //�������� ������ ���������� ��� �������������
	  SpeedButton2->Visible = true; //�������������� ������
	  SpeedButton3->Visible = true; //������ ��������


	  Prava = "ocen";
	}
  else
	{

	  DBGridEh1->Columns->Items[6]->Visible = false;     //������� �����������
	  DBGridEh1->Columns->Items[7]->Visible = false;     //������
	  DBGridEh1->Columns->Items[9]->Visible = false;     //���������������� �������, ����
	  DBGridEh1->Columns->Items[10]->Visible = false;    //���������������� �������, ����
	  DBGridEh1->Columns->Items[11]->Visible = false;    //���, ����
	  DBGridEh1->Columns->Items[12]->Visible = false;    //���, ����
	  DBGridEh1->Columns->Items[13]->Visible = false;    //����������, ����������
	  DBGridEh1->Columns->Items[14]->Visible = false;    //����������, ����
	  DBGridEh1->Columns->Items[15]->Visible = false;    //���������, ����������
	  DBGridEh1->Columns->Items[16]->Visible = false;    //���������, ����
	  DBGridEh1->Columns->Items[17]->Visible = false;    //��������������, �������
	  DBGridEh1->Columns->Items[18]->Visible = false;    //��������������, ����
	  DBGridEh1->Columns->Items[19]->Visible = false;    //5�, �������
	  DBGridEh1->Columns->Items[20]->Visible = false;    //5�, ����
	  DBGridEh1->Columns->Items[21]->Visible = false;    //���, ����������
	  DBGridEh1->Columns->Items[22]->Visible = false;    //���, ����
	  DBGridEh1->Columns->Items[23]->Visible = false;    //���, ���-�� �������
	  DBGridEh1->Columns->Items[24]->Visible = false;    //���, ����
	  DBGridEh1->Columns->Items[25]->Visible = false;    //��, �� � �
	  DBGridEh1->Columns->Items[26]->Visible = false;    //��������� ��
	  DBGridEh1->Columns->Items[27]->Visible = false;    //�������� ����������

	  //������� ����
	  N6Spisok->Visible = false;     //�������� ������ ���������� ��� �������������
	  N7PP->Visible = false;         //�������� ����������������� �������
	  N8KPE->Visible = false;        //�������� ���
	  N51C5->Visible = false;        //�������� 5 ������
	  N10PREEM->Visible = false;     //�������� ����������
	  N11SPP->Visible = false;       //�������� ���
	  N12OT->Visible = false;        //�������� ��
	  N13NOT->Visible = false;       //�������� ��������� �� ��
	  N14TD->Visible = false;        //�������� �������� ����������
	  N3Otchet->Visible = false;     //������������ ��������� ������
	  NSprav->Visible = false;        //���������� �� ����������������� �������

	  //����������� ����
	  N16Dobav->Visible = false;     //���������� ������
	  N6Reiting->Visible = false;    //������ ��������

	  //������
	  SpeedButton1->Visible = false; //�������� ������ ���������� ��� �������������
	  SpeedButton3->Visible = false; //������ ��������

	  //���� ����� ������ ����
	  if (SL_Groups->IndexOf("mmk-itsvc-hrru-unou")>-1)
		{
		  DBGridEh1->Columns->Items[6]->Visible = true;  //������� �����������
		  DBGridEh1->Columns->Items[13]->Visible = true;    //����������, ����������
		  DBGridEh1->Columns->Items[14]->Visible = true;    //����������, ����
		  DBGridEh1->Columns->Items[17]->Visible = true;    //��������������, �������
		  DBGridEh1->Columns->Items[18]->Visible = true;    //��������������, ����
		  DBGridEh1->Columns->Items[19]->Visible = true;    //5�, �������
		  DBGridEh1->Columns->Items[20]->Visible = true;    //5�, ����
		  DBGridEh1->Columns->Items[21]->Visible = true;    //���, ����������
		  DBGridEh1->Columns->Items[22]->Visible = true;    //���, ����

		  //������� ����
		  N6Spisok->Visible = true;     //�������� ������ ���������� ��� �������������
		  N51C5->Visible = true;        //�������� 5 ������

		  //����������� ����
		  N16Dobav->Visible = true;     //���������� ������

		  //������
		  SpeedButton1->Visible = true; //�������� ������ ���������� ��� �������������

		  Prava = "unou";
		}

	  //���� ����� ������ �������� ����������������� �������
	  else if (SL_Groups->IndexOf("mmk-itsvc-hrru-pp")>-1)
		{
		  DBGridEh1->Columns->Items[9]->Visible = true;     //���������������� �������, ����
		  DBGridEh1->Columns->Items[10]->Visible = true;    //���������������� �������, ����
		  N7PP->Visible = true;         //�������� ����������������� �������

		  Prava = "pp";
		}

	  //���� ����� ������ �������� ���
	  else if (SL_Groups->IndexOf("mmk-itsvc-hrru-kpe")>-1)
		{
		  DBGridEh1->Columns->Items[11]->Visible = true;    //���, ����
		  DBGridEh1->Columns->Items[12]->Visible = true;    //���, ����
		  N8KPE->Visible = true;        //�������� ���

		  Prava = "kpe";
		}

	  //���� ����� ������ �������� ���
	  else if (SL_Groups->IndexOf("mmk-itsvc-hrru-spp")>-1)
		{
		  DBGridEh1->Columns->Items[23]->Visible = true;    //���, ���-�� �������
		  DBGridEh1->Columns->Items[24]->Visible = true;    //���, ����
		  N11SPP->Visible = true;       //�������� ���

		  Prava = "spp";
		}

	  //���� ����� ������ �������� ��
	  else if (SL_Groups->IndexOf("mmk-itsvc-hrru-ot")>-1)
		{
		  DBGridEh1->Columns->Items[25]->Visible = true;    //��, �� � �
		  DBGridEh1->Columns->Items[26]->Visible = true;    //��������� ��
		  N12OT->Visible = true;        //�������� ��
		  N13NOT->Visible = true;       //�������� ��������� �� ��

		  Prava = "ot";
		}
	  //���� ����� ������ �������� �������� ����������
	  else if (SL_Groups->IndexOf("mmk-itsvc-hrru-td")>-1)
		{
		  DBGridEh1->Columns->Items[27]->Visible = true;    //�������� ����������
		  N14TD->Visible = true;        //�������� �������� ����������

		  Prava = "td";
		}
	  else
		{
		  Application->MessageBox(L"�� ����������� ����� ������� ��� ������ � ���������� '������������� �������������'!!!",L"����� �������",
								  MB_OK+MB_ICONERROR);
		  Application->Terminate();
		  Abort();

		}
	} */


  //���������� �� ���� ����� ���� ������� �����
  Main->WindowState = wsMaximized;

  //����������� ���������� ������
  AnsiString width = Screen->Width;     //������
  AnsiString height = Screen->Height;   //������

  //��������� ������������� ����������� ���� � ����������� �� ����������
  if (width >= 1280 && height >= 1024 ||
	  width >=1600 && height >= 900)
	{
	  DBGridEh1->AutoFitColWidths = true;
	}
  else
	{
	  DBGridEh1->AutoFitColWidths = false;
	}

  //���������� �������������� ��� ������� Enter
  //DBGridEh1->Style->FilterEditCloseUpApplyFilter =true;
  //DBGridEhCenter()->FilterEditCloseUpApplyFilter = true;
	// RebuildWindowRgn(Panel3);
		// SetWindowLong( this->Handle, GWL_EXSTYLE, this->GetExStyle() | WS_EX_TRANSPARENT );

 /*
��� ��������:
����� ���� �� ����������� =
			������� ����� �� ����������� * ������� ���� �� �����������+ (1-������� ����� �� �����������)*�������������� ���� �������
		   /*     int Transparency = 75;
long ExtStyle = GetWindowLong(Handle, GWL_EXSTYLE);
SetWindowLong(Handle, GWL_EXSTYLE, ExtStyle | WS_EX_LAYERED);
SetLayeredWindowAttributes(Handle, 0 , (255 * Transparency) / 100, LWA_ALPHA);



   /*	 SetWindowLong(Panel3->Handle, GWL_EXSTYLE,
		GetWindowLong(Panel3->Handle, GWL_EXSTYLE) & ~WS_EX_LAYERED);
	SetWindowlong(Panel3->Handle, GWL_EXSTYLE,
		GetWindowLong(Panel3->Handle, GWL_EXSTYLE) | WS_EX_LAYERED);
	SetLayeredWindowAttributes(Panel3->Handle, 0, 125, LWA_ALPHA);   */

   //	Panel3->Al
  //	ParentBackground = false;
  //	Panel3->Color = clBlack;
   //	Panel3->Canvas->Transparent = 50;
   //Panel3->C
   //�������� - Transparent. ��� ������� Image - Image1->Transparent = 50; Panel1->Canvas->Transparent = 50;

   //���������� �������������� ��� ������� Enter
/*  //DBGridEh1->Style->FilterEditCloseUpApplyFilter =true;

  //����������� ���������� ������
  AnsiString width = Screen->Width;     //������
  AnsiString height = Screen->Height;   //������

 //��������� ������� ������ � ����������� �� ���������� ������
  if ( width >=1600 && height >= 900)
	{
	  DBGridEh1->Font->Size = 11;
	}
  else
	{
	  DBGridEh1->Font->Size = 10;
	}
 */


  //SpeedButton1->Glyph->TransparentMode=tmFixed;
  //SpeedButton1->Glyph->Transparent = false;

  //������� ������������ �� �������
  SpeedButton1->Glyph->TransparentColor = clBlue;
  SpeedButton2->Glyph->TransparentColor = clBlue;
  SpeedButton3->Glyph->TransparentColor = clBlue;
  SpeedButton4->Glyph->TransparentColor = clBlue;


  //������������ ��������� �������
   Word Year, Month, Day;

  DecodeDate(Date(),Year, Month, Day);

  //�������� ���
  god=Year;

  //�������� �������
  if (Month==1 || Month==2 || Month==3) kvartal=1;
  else if (Month==4 || Month==5 || Month==6) kvartal=2;
  else if (Month==7 || Month==8 || Month==9) kvartal=3;
  else if (Month==10 || Month==11 || Month==12) kvartal=4;
  else{
	Application->MessageBox(L"���������� ���������� ������� �������",L"������",MB_OK+MB_ICONERROR);
	Application->Terminate();
	Abort();
  }

  //������ �� �������������� ���� ����� ������ ��  ������ ��������� ����� 25 �����
  if (Day>24 && Prava!="ocen") {

	 //������� ����
	  N6Spisok->Enabled = false;     //�������� ������ ���������� ��� �������������
	  N7PP->Enabled = false;         //�������� ����������������� �������
	  N8KPE->Enabled = false;        //�������� ���
	  N51C5->Enabled = false;        //�������� 5 ������
	  N10PREEM->Enabled = false;     //�������� ����������
	  N11SPP->Enabled = false;       //�������� ���
	  N12OT->Enabled = false;        //�������� ��
	  N13NOT->Enabled = false;       //�������� ��������� �� ��
	  N14TD->Enabled = false;        //�������� �������� ����������

	  //����������� ����
	  N16Dobav->Enabled = false;     //���������� ������
	  N17Redact->Enabled = false;    //�������������� ������

	  //������
	  SpeedButton1->Enabled = false; //�������� ������ ���������� ��� �������������
	  SpeedButton2->Enabled = false; //�������������� ������
  }
  else {
	 //������� ����
	  N6Spisok->Enabled = true;     //�������� ������ ���������� ��� �������������
	  N7PP->Enabled = true;         //�������� ����������������� �������
	  N8KPE->Enabled = true;        //�������� ���
	  N51C5->Enabled = true;        //�������� 5 ������
	  N10PREEM->Enabled = true;     //�������� ����������
	  N11SPP->Enabled = true;       //�������� ���
	  N12OT->Enabled = true;        //�������� ��
	  N13NOT->Enabled = true;       //�������� ��������� �� ��
	  N14TD->Enabled = true;        //�������� �������� ����������

	  //����������� ����
	  N16Dobav->Enabled = true;     //���������� ������
	  N17Redact->Enabled = true;    //�������������� ������

	  //������
	  SpeedButton1->Enabled = true; //�������� ������ ���������� ��� �������������
	  SpeedButton2->Enabled = true; //�������������� ������
  }


  //������� ������ � ������ ��������� �������
  //DM->qReiting->Close();
  DM->qReiting->ParamByName("pgod")->Value= god;
  DM->qReiting->ParamByName("pkvartal")->Value = kvartal;
  try
	{
	  DM->qReiting->Active=true;
	}
  catch(Exception &E)
	{
	  Application->MessageBox(("�������� ������ ��� ������� ��������� ������ �� ������� REIT_RUK "+E.Message).c_str(),L"������",MB_OK+MB_ICONERROR);
	  Application->Terminate();
	  Abort();
	}


  if (!GetMyDocumentsDir(DocPath))
    {
	  MessageBox(Handle,L"������ ������� � ����� ����������",L"������",8208);
	  Application->Terminate();
	  Abort();
	}

  if (!GetTempDir(TempPath))
	{
      MessageBox(Handle,L"������ ������� � ��������� �����",L"������",8208);
	  Application->Terminate();
      Abort();
    }

  WorkPath = DocPath + "\\������������� �������������";
  Path = GetCurrentDir();
  FindWordPath();

  Application->UpdateFormatSettings = false;
  FormatSettings.DecimalSeparator = '.';
  FormatSettings.DateSeparator = '.';
  FormatSettings.ShortDateFormat = "dd.mm.yyyy";

  StatusBar1->SimpleText ="  �������� ������: "+IntToStr(god)+" ���, "+IntToStr(kvartal)+" �������";

   // �������� ProgressBar �� StatusBar
  ProgressBar = new TProgressBar ( StatusBar1 );
  ProgressBar->Parent = StatusBar1;
  ProgressBar->Position = 0;
  ProgressBar->Left = Main->Width-ProgressBar->Width-40;//StatusBar1->Width-ProgressBar->Width-10;//StatusBar1->Panels->Items[0]->Width+StatusBar1->Panels->Items[1]->Width - ProgressBar->Width;//Width*18 + 81;
  //ProgressBar->Anchors = ProgressBar->Anchors << akRight << akTop << akLeft << akBottom;
  ProgressBar->Top = StatusBar1->Height/11;
  ProgressBar->Height = StatusBar1->Height-3;
  PostMessage(ProgressBar->Handle,0x0409,0,clRed);
  ProgressBar->Visible = false;
}
//---------------------------------------------------------------------------
//���������� ������
void __fastcall TMain::N16DobavClick(TObject *Sender)
{
   redakt = 0;
   Vvod->ShowModal();
}
//---------------------------------------------------------------------------
//�������������� ������
void __fastcall TMain::N17RedactClick(TObject *Sender)
{
  redakt = 1;
  Vvod->ShowModal();
}
//---------------------------------------------------------------------------
//�������� ������ ������ ����������
void __fastcall TMain::N6SpisokClick(TObject *Sender)
{
   Variant AppEx, Sh;
   AnsiString  Dir, Sql, tn_proverka="NULL";

   int otchet=0, kol=0, rec=0, ob_kol=0, obnov_kol=0,
   pr=0,
   zex, tn, fio, id_dolg, dolg, uch, podch;



  StatusBar1->SimpleText="  ���� �������� ������...";

  // ������������ ����� ��� ��������
  zex=2;     //B
  tn=4;      //D
  fio=5;     //E
  id_dolg=7; //G
  dolg=6;    //F
  uch=8;     //H
  podch=9;   //I
  update=0;

  StatusBar1->SimpleText="  ����� ��������� ��� ��������...";

  OpenDialog1->Filter = "Excel files (*.xls, *.xlsx)|*.xls; *.xlsx";
  // DefaultExt

  //����� ����� ��� ��������
  if (!OpenDialog1->Execute()){
	  StatusBar1->SimpleText ="  �������� ������: "+IntToStr(god)+" ���, "+IntToStr(kvartal)+" �������";
	  Abort();
  }

  StatusBar1->SimpleText = "  �������� ������ �� ����� "+OpenDialog1->FileName;


   //�������� ����� ������ ��� ������ �� ����������� ������
  if (!rtf_Open((TempPath + "\\zagruzka.txt").c_str()))
	{
	  MessageBox(Handle,L"������ �������� ����� ������",L"������",8192);
	  Abort();
	}

  rtf_Out("data", DateTimeToStr(Now()),0);


  //�������� ��������� Excel
  try
	{
	  AppEx = CreateOleObject("Excel.Application");
	}
  catch (...)
	{
	  Application->MessageBox(L"���������� ������� Microsoft Excel!\n �������� ��� ���������� �� ���������� �� �����������.",
							  L"������", MB_OK+MB_ICONERROR);
	  StatusBar1->SimpleText ="  �������� ������: "+IntToStr(god)+" ���, "+IntToStr(kvartal)+" �������";
	  Abort();
	}

  //���� ��������� ������ �� ����� ������������ ������
  try
	{
	  try
		{
		  AppEx.OlePropertyGet("Workbooks").OlePropertyGet("Open", WideString(OpenDialog1->FileName));
		  AppEx.OlePropertySet("Visible",false);
		  Sh = AppEx.OlePropertyGet("Worksheets", 1);
		}
	  catch(...)
		{
		  Application->MessageBox(L"������ �������� ����� Microsoft Excel!", L"������",MB_OK + MB_ICONERROR);
		  StatusBar1->SimpleText ="  �������� ������: "+IntToStr(god)+" ���, "+IntToStr(kvartal)+" �������";
		  Abort();
		}


	  //���������� ���������� ������� ����� � ���������
	  AnsiString Row = Sh.OlePropertyGet("UsedRange").OlePropertyGet("Rows").OlePropertyGet("Count");


	  //�������� �� ������� ������ � �������
	  Sql = "select count(*) as kol from reit_ruk \
			 where god="+IntToStr(god) +" and kvart="+IntToStr(kvartal);


	  DM->qObnovlenie->Close();
	  DM->qObnovlenie->SQL->Clear();
	  DM->qObnovlenie->SQL->Add(Sql);
	  try
		{
		  DM->qObnovlenie->Open();
		}
	  catch(Exception &E)
		{
		  Application->MessageBox(("�������� ������ ��� ������� ������� ������ �� ������� REIT_RUK: " + E.Message).c_str(),L"������",
									MB_OK+MB_ICONERROR);

		  InsertLog("�������� ������ ��� �������� ������ ���������� ��� ������������� �� ����� '"+OpenDialog1->FileName+"' �� "+IntToStr(god)+" ���, "+IntToStr(kvartal)+" �������");
		  DM->qReiting->Refresh();
		  StatusBar1->SimpleText ="  �������� ������: "+IntToStr(god)+" ���, "+IntToStr(kvartal)+" �������";
		  Abort();
		}

	  if (DM->qObnovlenie->FieldByName("kol")->AsInteger>0)
		{
		  if (Application->MessageBox(("� ������� ��� ����������� ������ �� "+IntToStr(kvartal)+" ������� "+IntToStr(god)+" ���\n��� ���������� ������ �� ���� ������ ����� ������������\n�� ������������� ������ �������� ������?").c_str(),
										L"��������� ������",MB_YESNO+MB_ICONWARNING)==ID_NO)
			 {
			   update=0;
			   Abort();
			 }

		   //�������� ���������� ������� � �������� ������
		   Sql = "delete from reit_ruk \
				  where god="+IntToStr(god) +" and kvart="+IntToStr(kvartal);

		   DM->qObnovlenie->Close();
		   DM->qObnovlenie->SQL->Clear();
		   DM->qObnovlenie->SQL->Add(Sql);
		   try
			 {
			   DM->qObnovlenie->ExecSQL();
			 }
		   catch(Exception &E)
			 {
			   Application->MessageBox(("�������� ������ ��� ������� ������� ������ �� ������� REIT_RUK: " + E.Message).c_str(),L"������",
										MB_OK+MB_ICONERROR);

			   InsertLog("�������� ������ ��� �������� ������ ���������� ��� ������������� �� ����� '"+OpenDialog1->FileName+"' �� "+IntToStr(god)+" ���, "+IntToStr(kvartal)+" �������");
			   DM->qReiting->Refresh();
			   StatusBar1->SimpleText ="  �������� ������: "+IntToStr(god)+" ���, "+IntToStr(kvartal)+" �������";
			   Abort();
			 }

		   update=0;
		}
	 // else update=0;

	  StatusBar1->SimpleText ="  ����������� �������� ������ ����������...";


	  Cursor = crHourGlass;
	  ProgressBar->Position = 0;
	  ProgressBar->Visible = true;
	  ProgressBar->Max=StrToInt(Row);

	  //�������� ������
	  for (int i=1; i<Row+1; i++)
		{
		  tn_proverka = Sh.OlePropertyGet("Cells",i,tn);//.OlePropertyGet("Value");


		  //�������� �� ������� ���.� � ����� ������ � ������� ����������� ����
		  if (tn_proverka.IsEmpty() || !Proverka(tn_proverka))  continue;
			{
//******************************************************************************
			  //�������� �� ���������� ������ � ������ � � ����������� ������ ����������
			  DM->qProverka->Close();
			  DM->qProverka->ParamByName("ptn_sap")->Value=tn_proverka;

			  try
				{
				  DM->qProverka->Active = true;
				}
			  catch (Exception &E)
				{
				  Application->MessageBox(("�������� ������ ��� ������� ������� ������ �� ������� SAP_OSN_SVED: " + E.Message).c_str(),L"������",
										   MB_OK+MB_ICONERROR);

				  InsertLog("�������� ������ ��� �������� ������ ���������� ��� ������������� �� ����� '"+OpenDialog1->FileName+"' �� "+IntToStr(god)+" ���, "+IntToStr(kvartal)+" �������");
				  DM->qReiting->Refresh();
				  StatusBar1->SimpleText ="  �������� ������: "+IntToStr(god)+" ���, "+IntToStr(kvartal)+" �������";
				  Cursor = crDefault;
				  ProgressBar->Visible = false;
				  Abort();
				}


			  //�������������� ����� ����
			  if (DM->qProverka->FieldByName("zex")->AsString!=VarToStr(Sh.OlePropertyGet("Cells",i,zex)))
				{
				  //������������ ��������� � ����� �������
				  if (pr!=1)
					{
					   rtf_Out("z", " ",1);
					   if(!rtf_LineFeed())
						 {
						   MessageBox(Handle,L"������ ������ � ���� ������",L"������",8192);
						   if (!rtf_Close()) MessageBox(Handle,L"������ �������� ����� ������",L"������",8192);
						   return;
						 }
					 }
						 //AnsiString aarr =  Sh.OlePropertyGet("Cells",i,tn);



				   //������������ ������ �� ������������� �������
				   rtf_Out("tn", VarToStr(Sh.OlePropertyGet("Cells",i,tn)),2);
				   rtf_Out("zex_f", VarToStr(Sh.OlePropertyGet("Cells",i,zex)),2);
				   rtf_Out("zex", DM->qProverka->FieldByName("zex")->AsString,2);
				   rtf_Out("fio", VarToStr(Sh.OlePropertyGet("Cells",i,fio)),2);

				   if(!rtf_LineFeed())
					 {
					   MessageBox(Handle,L"������ ������ � ���� ������",L"������",8192);
					   if (!rtf_Close()) MessageBox(Handle,L"������ �������� ����� ������",L"������",8192);
					   return;
					 }
				   pr=1;      //������� ������������ ����� ������
				   otchet=1;  //������� ������������ ������ �� ������������� �������
				}

			  //�������������� ���
			  if (DM->qProverka->FieldByName("fio")->AsString!=VarToStr(Sh.OlePropertyGet("Cells",i,fio)))
				{
				  //������������ ��������� � ����� �������
				  if (pr!=3)
					{
					   rtf_Out("z", " ",3);
					   if(!rtf_LineFeed())
						 {
						   MessageBox(Handle,L"������ ������ � ���� ������",L"������",8192);
						   if (!rtf_Close()) MessageBox(Handle,L"������ �������� ����� ������",L"������",8192);
						   return;
						 }
					}

				   //������������ ������ �� ������������� �������
				   rtf_Out("tn", VarToStr(Sh.OlePropertyGet("Cells",i,tn)),4);
				   rtf_Out("zex", VarToStr(Sh.OlePropertyGet("Cells",i,zex)),4);
				   rtf_Out("fio_f", VarToStr(Sh.OlePropertyGet("Cells",i,fio)),4);
				   rtf_Out("fio", DM->qProverka->FieldByName("fio")->AsString,4);

				   if(!rtf_LineFeed())
					 {
					   MessageBox(Handle,L"������ ������ � ���� ������",L"������",8192);
					   if (!rtf_Close()) MessageBox(Handle,L"������ �������� ����� ������",L"������",8192);
					   return;
					 }
				   pr=3;      //������� ������������ ����� ������
				   otchet=1;  //������� ������������ ������ �� ������������� �������
				}
			  //�������������� �� ���������
			  if (DM->qProverka->FieldByName("id_shtat")->AsString!=VarToStr(Sh.OlePropertyGet("Cells",i,id_dolg)))
				{
				  //������������ ��������� � ����� �������
				   if (pr!=5)
					 {
					   rtf_Out("z", " ",5);
					   if(!rtf_LineFeed())
						 {
						   MessageBox(Handle,L"������ ������ � ���� ������",L"������",8192);
						   if (!rtf_Close()) MessageBox(Handle,L"������ �������� ����� ������",L"������",8192);
						   return;
						 }
					 }

				   //������������ ������ �� ������������� �������
				   rtf_Out("tn", VarToStr(Sh.OlePropertyGet("Cells",i,tn)),6);
				   rtf_Out("zex", VarToStr(Sh.OlePropertyGet("Cells",i,zex)),6);
				   rtf_Out("id_dolg_f", VarToStr(Sh.OlePropertyGet("Cells",i,id_dolg)),6);
				   rtf_Out("id_dolg", DM->qProverka->FieldByName("id_shtat")->AsString,6);
				   rtf_Out("fio", VarToStr(Sh.OlePropertyGet("Cells",i,fio)),6);

				   if(!rtf_LineFeed())
					 {
					   MessageBox(Handle,L"������ ������ � ���� ������",L"������",8192);
					   if (!rtf_Close()) MessageBox(Handle,L"������ �������� ����� ������",L"������",8192);
					   return;
					 }
				   pr=5;      //������� ������������ ����� ������
				   otchet=1;  //������� ������������ ������ �� ������������� �������
				}
			  //�������������� ������������ ���������
			  if (DM->qProverka->FieldByName("name_dolg_ru")->AsString!=VarToStr(Sh.OlePropertyGet("Cells",i,dolg)))
				{
					//������������ ��������� � ����� �������
				   if (pr!=7)
					 {
					   rtf_Out("z", " ",7);
					   if(!rtf_LineFeed())
						 {
						   MessageBox(Handle,L"������ ������ � ���� ������",L"������",8192);
						   if (!rtf_Close()) MessageBox(Handle,L"������ �������� ����� ������",L"������",8192);
						   return;
						 }
					 }

				   //������������ ������ �� ������������� �������
				   rtf_Out("tn", VarToStr(Sh.OlePropertyGet("Cells",i,tn)),8);
				   rtf_Out("zex", VarToStr(Sh.OlePropertyGet("Cells",i,zex)),8);
				   rtf_Out("dolg_f", VarToStr(Sh.OlePropertyGet("Cells",i,dolg)),8);
				   rtf_Out("dolg", DM->qProverka->FieldByName("name_dolg_ru")->AsString,8);

				   if(!rtf_LineFeed())
					 {
					   MessageBox(Handle,L"������ ������ � ���� ������",L"������",8192);
					   if (!rtf_Close()) MessageBox(Handle,L"������ �������� ����� ������",L"������",8192);
					   return;
					 }
				   pr=7;      //������� ������������ ����� ������
				   otchet=1;  //������� ������������ ������ �� ������������� �������
				}


//******************************************************************************
			  //�������� �� ������� ������ � �������

			  //�������� ������ � ����
			  if (update==1)
				{
				  Sql = "update reit_ruk set \
										 zex=trim('"+ Sh.OlePropertyGet("Cells",i,zex) +"'), \
										 tn=trim('"+ Sh.OlePropertyGet("Cells",i,tn) +"'), \
										 fio=initcap(trim('"+ Sh.OlePropertyGet("Cells",i,fio) +"')),  \
										 id_dolg=lpad(trim('"+ Sh.OlePropertyGet("Cells",i,id_dolg) +"'),'0',8), \
										 dolg=trim('"+ Sh.OlePropertyGet("Cells",i,dolg) +"'),  \
										 uch=trim('"+ Sh.OlePropertyGet("Cells",i,uch) +"'),  \
										 podch=decode(trim('"+ Sh.OlePropertyGet("Cells",i,podch) +"'),'��','1',0) \
						  where tn="+ Sh.OlePropertyGet("Cells",i,tn)+" and god="+IntToStr(god) +" and kvart="+IntToStr(kvartal);
				}
			  else
				{
				  Sql = "insert into reit_ruk (god, kvart, zex, tn, fio, id_dolg, dolg, uch, podch) \
						 values ( "+IntToStr(god) +",    \
								  "+IntToStr(kvartal)+", \
								   trim('"+ Sh.OlePropertyGet("Cells",i,zex) +"'), \
								   trim('"+ Sh.OlePropertyGet("Cells",i,tn) +"'), \
								   initcap(trim('"+ Sh.OlePropertyGet("Cells",i,fio) +"')),  \
								   lpad(trim('"+ Sh.OlePropertyGet("Cells",i,id_dolg) +"'),'0',8), \
								   trim('"+ Sh.OlePropertyGet("Cells",i,dolg) +"'),  \
								   trim('"+ Sh.OlePropertyGet("Cells",i,uch) +"'),  \
								   decode(trim('"+ Sh.OlePropertyGet("Cells",i,podch) +"'),'��','1',0)) ";
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
				  Application->MessageBox(("�������� ������ ��� ������� �������� ������ � ������� REIT_RUK: " + E.Message).c_str(),L"������",
											MB_OK+MB_ICONERROR);

				  InsertLog("�������� ������ ��� �������� ������ ���������� ��� ������������� �� ����� '"+OpenDialog1->FileName+"' �� "+IntToStr(god)+" ���, "+IntToStr(kvartal)+" �������");
				  DM->qReiting->Refresh();
				  StatusBar1->SimpleText ="  �������� ������: "+IntToStr(god)+" ���, "+IntToStr(kvartal)+" �������";
				  Cursor = crDefault;
				  ProgressBar->Visible = false;
				  Abort();
				}

			  rec++;
			  kol+=DM->qObnovlenie->RowsAffected;

			  // ���������� ����������� �������
			  if (DM->qObnovlenie->RowsAffected == 0)
				{
				  //������������ ��������� � ����� �������
				   if (pr!=9)
					 {
					   rtf_Out("z", " ",9);
					   if(!rtf_LineFeed())
						 {
						   MessageBox(Handle,L"������ ������ � ���� ������",L"������",8192);
						   if (!rtf_Close()) MessageBox(Handle,L"������ �������� ����� ������",L"������",8192);
						   return;
						 }
					 }

				   //������������ ������ �� ������������� �������
				   rtf_Out("tn", VarToStr(Sh.OlePropertyGet("Cells",i,tn)),10);
				   rtf_Out("zex", VarToStr(Sh.OlePropertyGet("Cells",i,zex)),10);
				   rtf_Out("fio", VarToStr(Sh.OlePropertyGet("Cells",i,fio)),10);

				   if(!rtf_LineFeed())
					 {
					   MessageBox(Handle,L"������ ������ � ���� ������",L"������",8192);
					   if (!rtf_Close()) MessageBox(Handle,L"������ �������� ����� ������",L"������",8192);
					   return;
					 }
				   pr=9;      //������� ������������ ����� ������
				   otchet=1;  //������� ������������ ������ �� ������������� �������
				 }
			   else obnov_kol++;
		  }

		  ProgressBar->Position++;
		  ob_kol++;
		}


	  StatusBar1->SimpleText = "  �������� ������ ���������.";

	  DM->qReiting->Refresh();

	  //�������� Excel
	  AppEx.OleProcedure("Quit");
	  AppEx = Unassigned;


	  if(!rtf_Close())
		{
		  MessageBox(Handle,L"������ �������� ����� ������", L"������", 8192);
		  return;
		}

	  //������������ ������ � Word
	  if (otchet==1)
		{
		  StatusBar1->SimpleText = "  ������������ ������ � ��������...";

		  //�������� �����, ���� �� �� ����������
		  ForceDirectories(WorkPath);

		  int istrd;
		  try
			{
			  rtf_CreateReport(TempPath + "\\zagruzka.txt", Path+"\\RTF\\zagruzka.rtf",
							   WorkPath+"\\�����.doc",NULL,&istrd);


			  WinExec(("\""+ WordPath+"\"\""+WorkPath+"\\�����.doc\"").c_str(),SW_MAXIMIZE);

			}
		  catch(RepoRTF_Error E)
			{
			  Application->MessageBox(("������ ������������ ������:"+ String(E.Err)+
								 "\n������ ����� ������:"+IntToStr(istrd)).c_str(),
								 L"������",
								 MB_OK+MB_ICONERROR);
			}

		  Application->MessageBox(("���������� �������������� ���������� � ����������� ����� � ���� ������ �� ���������.\n��������� ������������� ���������� � ����� \n "+OpenDialog1->FileName+" � ��������� ��������� ��������").c_str() ,L" �������� ������ ����������",
								  MB_OK + MB_ICONINFORMATION);

		}

	  DeleteFile(TempPath+"\\otchet.txt");
	  InsertLog("�������� ������ ���������� ��� ������������� �� ����� '"+OpenDialog1->FileName+"' �� "+IntToStr(god)+" ���, "+IntToStr(kvartal)+" ������� ��������� �������. ��������� " + IntToStr(obnov_kol) + " �� " + IntToStr(ob_kol)+" �������");


       Application->MessageBox(("�������� ������ ���������� ��� ������������� ��������� �������. =) \n��������� " + IntToStr(obnov_kol) + " �� " + IntToStr(ob_kol)+" �������").c_str(),
						   L"���������� ������ �� ������ ���������",
						   MB_OK + MB_ICONINFORMATION);
	}
  catch(...)
	{
	  AppEx.OleProcedure("Quit");
	  //AppEx.Clear();
	  //VarClear(AppEx);
	  AppEx=Unassigned;
	  InsertLog("�������� ������ ��� �������� ������ ���������� ��� ������������� �� ����� '"+OpenDialog1->FileName+"' �� "+IntToStr(god)+" ���, "+IntToStr(kvartal)+" �������");
	}


  Cursor = crDefault;
  ProgressBar->Position = 0;
  ProgressBar->Visible = false;

  StatusBar1->SimplePanel = false;
  StatusBar1->SimpleText ="  �������� ������: "+IntToStr(god)+" ���, "+IntToStr(kvartal)+" �������";

}
//---------------------------------------------------------------------------


// �������� �� �������� ���.� � Excel-�����
bool  __fastcall TMain::Proverka(String tn)
{
   try {
	StrToInt(tn);
  }
  catch (...) {
	return false;
  }
  return true;

}
//---------------------------------------------------------------------------

// ���������� ���� �� ����� "��� ���������"
bool __fastcall TMain::GetMyDocumentsDir(AnsiString &FolderPath)
{
  wchar_t f[MAX_PATH];

  if (SUCCEEDED(SHGetFolderPath(NULL, CSIDL_PERSONAL|CSIDL_FLAG_CREATE, NULL, SHGFP_TYPE_CURRENT, f)))
	{
	  FolderPath = AnsiString(f);
	  return(true);
	}

  return(false);
}
//---------------------------------------------------------------------------

// ���������� ���� �� ����� Temp
bool __fastcall TMain::GetTempDir(AnsiString &FolderPath)
{
  wchar_t f[MAX_PATH];

  if (GetTempPath(MAX_PATH, f))
	{
	  FolderPath = AnsiString(f);
	  FolderPath = FolderPath.SubString(1, FolderPath.Length()-1);
	  return(true);
	}

  return(false);
}
//---------------------------------------------------------------------------

 // ���������� ������ �������������� Word
  AnsiString __fastcall TMain::FindWordPath()
{
  TRegistry *Reg = new TRegistry;
    try {
    Reg->RootKey = HKEY_LOCAL_MACHINE;

    for (int v=20; v>5; v--) {
      if (Reg->OpenKeyReadOnly("Software\\Microsoft\\Office\\"+IntToStr(v)+".0\\Word\\InstallRoot")) {
        if (Reg->ValueExists("Path")) {
          WordPath = Reg->ReadString("Path") + "winword.exe";
          Reg->CloseKey();
          break;
        }
        Reg->CloseKey();
      }
    }
  }
  __finally {
    delete Reg;
  }
  return(WordPath);
}
//---------------------------------------------------------------------------
//�������������� ������
void __fastcall TMain::SpeedButton2Click(TObject *Sender)
{
  redakt=1;
  Vvod->ShowModal();
}
//---------------------------------------------------------------------------


void __fastcall TMain::DBGridEh1DblClick(TObject *Sender)
{
   SpeedButton2Click(Sender);
}
//---------------------------------------------------------------------------
//---------------------------------------------------------------------------

//����
void __fastcall TMain::InsertLog(String Msg)
{
  String Data;
  //DateTimeToStr(Data, "dd.mm.yyyy hh:nn:ss", Now());
  Data = DateTimeToStr(Now());
  DM->qObnovlenie->SQL->Clear();
  DM->qObnovlenie->SQL->Add("insert into logs_reit (DT,DOMAIN,USEROK,PROG,USEROK_FIO,TEXT) values \
                            (to_date(" + QuotedStr(Data) + ", 'DD.MM.YYYY HH24:MI:SS'),\
							 " + QuotedStr(DomainName) + "," + QuotedStr(UserName) + ", 'Reit_ruk',\
							 " + QuotedStr(UserFullName)+",  \
                             " + QuotedStr(Msg)+")");
  try
	{
	  DM->qObnovlenie->ExecSQL();
	}
  catch(...)
	{
	  Application->MessageBoxW(L"�������� ������ ��� ������� ������ � ������� LOGS_REIT",L"������",
							   MB_ICONERROR);
	}

  DM->qObnovlenie->Close();
}
 //---------------------------------------------------------------------------   */
//���������������� �������
void __fastcall TMain::N7PPClick(TObject *Sender)
{
   Zagruzka(3,1);
}
//---------------------------------------------------------------------------
//���
void __fastcall TMain::N8KPEClick(TObject *Sender)
{
   Zagruzka(3,2);
}
//---------------------------------------------------------------------------
//---------------------------------------------------------------------------
//5 ������
void __fastcall TMain::N51C5Click(TObject *Sender)
{
  Zagruzka(2,3);
}
//---------------------------------------------------------------------------
//���������
void __fastcall TMain::N10PREEMClick(TObject *Sender)
{
  Zagruzka(3,4);
}
//---------------------------------------------------------------------------
//���
void __fastcall TMain::N11SPPClick(TObject *Sender)
{
  Zagruzka(3,5);
}
//---------------------------------------------------------------------------
//��, �� � �
void __fastcall TMain::N12OTClick(TObject *Sender)
{
  Zagruzka(3,6);
}
//---------------------------------------------------------------------------
//��������� ��
void __fastcall TMain::N13NOTClick(TObject *Sender)
{
  Zagruzka(3,7);
}
//---------------------------------------------------------------------------
//�������� ����������
void __fastcall TMain::N14TDClick(TObject *Sender)
{
  Zagruzka(3,8);
}
//-------------------------------------------------------------------------
void __fastcall TMain::Zagruzka(int tn, int otchet)
{
   Variant AppEx, Sh;
   AnsiString  Dir, Sql, tn_proverka="NULL", str;
   String stroka;
   double stroka1;

   int fotchet=0, kol=0, rec=0, pr=0,
	   proiz, proiz_ball, kpe, kpe_ball, otkl, otkl_ball,
	   info, info_ball, c5, c5_ball, kns, kns_ball, priem,
	   priem_ball, spp_kol, spp_ball, ot_upr, ot_treb, trud_d, fio,
	   ob_kol=0, obnov_kol=0;


  StatusBar1->SimpleText="  ���� �������� ������...";

  // ������������ ����� ��� ��������
  proiz =5;       //E
  proiz_ball =6;  //F
  kpe = 5;        //E
  kpe_ball = 6;   //F
  otkl =7;        //H
  otkl_ball =8;   //I
  info =9;       //J
  info_ball = 10; //K
  c5 = 11;        //L
  c5_ball =12 ;   //M
  kns = 13;       //N
  kns_ball =14;   //O
  priem = 5;      //E
  priem_ball =6;  //F
  spp_kol = 6;    //F
  spp_ball = 8;   //H
  ot_upr =5;      //E
  ot_treb = 5;    //E
  trud_d = 5;     //E
  if (otchet!=3) fio =4; //D
  else fio = 3; //C


  StatusBar1->SimpleText="  ����� ��������� ��� ��������...";

  OpenDialog1->Filter = "Excel files (*.xls, *.xlsx)|*.xls; *.xlsx";
  // DefaultExt

  //����� ����� ��� ��������
  if (!OpenDialog1->Execute()){
	  Abort();
  }

  StatusBar1->SimpleText = "  �������� ������ �� ����� "+OpenDialog1->FileName;


   //�������� ����� ������ ��� ������ �� ����������� ������
  if (!rtf_Open((TempPath + "\\zagruzka.txt").c_str()))
	{
	  MessageBox(Handle,L"������ �������� ����� ������",L"������",8192);
	  Abort();
	}

  rtf_Out("data", DateTimeToStr(Now()),0);


  //�������� ��������� Excel
  try
	{
	  AppEx = CreateOleObject("Excel.Application");
	}
  catch (...)
	{
	  Application->MessageBox(L"���������� ������� Microsoft Excel!\n �������� ��� ���������� �� ���������� �� �����������.",
							  L"������", MB_OK+MB_ICONERROR);
	  StatusBar1->SimpleText ="  �������� ������: "+IntToStr(god)+" ���, "+IntToStr(kvartal)+" �������";
	  Abort();
	}

  //���� ��������� ������ �� ����� ������������ ������
  try
	{
	  try
		{
		  AppEx.OlePropertyGet("Workbooks").OlePropertyGet("Open", WideString(OpenDialog1->FileName));
		  AppEx.OlePropertySet("Visible",false);
		  Sh = AppEx.OlePropertyGet("Worksheets", 1);
		}
	  catch(...)
		{
		  Application->MessageBox(L"������ �������� ����� Microsoft Excel!", L"������",MB_OK + MB_ICONERROR);
		  StatusBar1->SimpleText ="  �������� ������: "+IntToStr(god)+" ���, "+IntToStr(kvartal)+" �������";
		  Abort();
		}


	  //���������� ���������� ������� ����� � ���������
	  AnsiString Row = Sh.OlePropertyGet("UsedRange").OlePropertyGet("Rows").OlePropertyGet("Count");


	  switch (otchet) {
		  case 1: str = "����������������� �������...";
		  break;
		  case 2: str = "���...";
		  break;
		  case 3: str = "5 ������...";
		  break;
		  case 4: str = "������� ����������...";
		  break;
		  case 5: str = "���...";
		  break;
		  case 6: str = "������ �����...";
		  break;
		  case 7: str = "���������� ���������� ��...";
		  break;
		  default: str = "�������� ����������...";
	  }


	  StatusBar1->SimpleText ="  ����������� �������� ������ �� "+str;

      Cursor = crHourGlass;
	  ProgressBar->Position = 0;
	  ProgressBar->Visible = true;
	  ProgressBar->Max=StrToInt(Row);

	  //�������� ������
	  for (int i=1; i<Row+1; i++)
		{
		  tn_proverka = Sh.OlePropertyGet("Cells",i,tn);//.OlePropertyGet("Value");


		  //�������� �� ������� ���.� � ����� ������ � ������� ����������� ����
		  if (tn_proverka.IsEmpty() || !Proverka(tn_proverka))  continue;
			{
//******************************************************************************
			  //�������� �� ������� ��������� � ������
			  DM->qProverka->Close();
			  DM->qProverka->ParamByName("ptn_sap")->Value=tn_proverka;

			  try
				{
				  DM->qProverka->Active = true;
				}
			  catch (Exception &E)
				{
				  Application->MessageBox(("�������� ������ ��� ������� ������� ������ �� ������� SAP_OSN_SVED: " + E.Message).c_str(),L"������",
										   MB_OK+MB_ICONERROR);

				  DM->qReiting->Refresh();
				  StatusBar1->SimpleText ="  �������� ������: "+IntToStr(god)+" ���, "+IntToStr(kvartal)+" �������";
				  InsertLog("�������� ������ ��� �������� ������ �� "+str+" �� ����� '"+OpenDialog1->FileName+"' �� "+IntToStr(god)+" ���, "+IntToStr(kvartal)+" �������");
				  Cursor = crDefault;
				  ProgressBar->Visible = false;
				  Abort();
				}


			  //������������ ������
			  if (DM->qProverka->RecordCount==0)
				{
                   //������������ ��������� � ����� �������
				   if (pr!=1)
					 {
					   rtf_Out("z", " ",1);
					   if(!rtf_LineFeed())
						 {
						   MessageBox(Handle,L"������ ������ � ���� ������",L"������",8192);
						   if (!rtf_Close()) MessageBox(Handle,L"������ �������� ����� ������",L"������",8192);
						   return;
						 }
					 }


				   rtf_Out("tn", VarToStr(Sh.OlePropertyGet("Cells",i,tn)),2);
				   rtf_Out("fio", VarToStr(Sh.OlePropertyGet("Cells",i,fio)),2);

				   if(!rtf_LineFeed())
					 {
					   MessageBox(Handle,L"������ ������ � ���� ������",L"������",8192);
					   if (!rtf_Close()) MessageBox(Handle,L"������ �������� ����� ������",L"������",8192);
					   return;
					 }
				   pr=1;      //������� ������������ ����� ������
				   fotchet=1;  //������� ������������ ������ �� ������������� �������
				}

			   float treb;
//******************************************************************************
			  //�������� ������ � ����
			  switch (otchet)
			  {

				//���������������� �������
				case 1:
					//���� ������ ���������� ������
					if (Sh.OlePropertyGet("Cells",i,proiz).OlePropertyGet("NumberFormat")=="0,00%" ||
						Sh.OlePropertyGet("Cells",i,proiz).OlePropertyGet("NumberFormat")=="0,00%" ||
						Sh.OlePropertyGet("Cells",i,proiz).OlePropertyGet("NumberFormat")=="0.0%" ||
						Sh.OlePropertyGet("Cells",i,proiz).OlePropertyGet("NumberFormat")=="0,0%" ||
						Sh.OlePropertyGet("Cells",i,proiz).OlePropertyGet("NumberFormat")=="0%")
					  {
						Sh.OlePropertyGet("Cells",i,proiz).OlePropertySet("NumberFormat",L"General");
						stroka = Sh.OlePropertyGet("Cells",i,proiz).OlePropertyGet("Value")*100;
					  }
					else stroka = Sh.OlePropertyGet("Cells",i,proiz).OlePropertyGet("Value");

				   Sql = "update reit_ruk set \
										 proiz=trim('"+stroka+"'),  \
										 proiz_ball=trim('"+ Sh.OlePropertyGet("Cells",i,proiz_ball) +"')  \
						  where tn="+ Sh.OlePropertyGet("Cells",i,tn)+" and god="+IntToStr(god) +" and kvart="+IntToStr(kvartal);
								 //.OlePropertyGet("Value")
				break;

				//���
				case 2:
					//���� ������ ���������� ������
					if (Sh.OlePropertyGet("Cells",i,kpe).OlePropertyGet("NumberFormat")=="0,00%" ||
						Sh.OlePropertyGet("Cells",i,kpe).OlePropertyGet("NumberFormat")=="0,00%" ||
						Sh.OlePropertyGet("Cells",i,kpe).OlePropertyGet("NumberFormat")=="0.0%" ||
						Sh.OlePropertyGet("Cells",i,kpe).OlePropertyGet("NumberFormat")=="0,0%" ||
						Sh.OlePropertyGet("Cells",i,kpe).OlePropertyGet("NumberFormat")=="0%")
					  {
						Sh.OlePropertyGet("Cells",i,kpe).OlePropertySet("NumberFormat",L"General");
						stroka = Sh.OlePropertyGet("Cells",i,kpe).OlePropertyGet("Value")*100;
					  }
					else stroka = Sh.OlePropertyGet("Cells",i,kpe).OlePropertyGet("Value");

					Sql = "update reit_ruk set \
										 kpe=trim('"+stroka+"'),  \
										 kpe_ball=trim('"+ Sh.OlePropertyGet("Cells",i,kpe_ball) +"')  \
						  where tn="+ Sh.OlePropertyGet("Cells",i,tn)+" and god="+IntToStr(god) +" and kvart="+IntToStr(kvartal);
				break;

				//5 ������
				case 3:


					//���� ������ ���������� ������
					if (Sh.OlePropertyGet("Cells",i,otkl).OlePropertyGet("NumberFormat")=="0,00%" ||
						Sh.OlePropertyGet("Cells",i,otkl).OlePropertyGet("NumberFormat")=="0,00%" ||
						Sh.OlePropertyGet("Cells",i,otkl).OlePropertyGet("NumberFormat")=="0.0%" ||
						Sh.OlePropertyGet("Cells",i,otkl).OlePropertyGet("NumberFormat")=="0,0%" ||
						Sh.OlePropertyGet("Cells",i,otkl).OlePropertyGet("NumberFormat")=="0%")
					  {
						Sh.OlePropertyGet("Cells",i,otkl).OlePropertySet("NumberFormat",L"General");

						stroka = Sh.OlePropertyGet("Cells",i,otkl).OlePropertyGet("Value")*100;
						//stroka =  FloatToStrF(String(stroka),ffFixed,20,1);
						stroka1=stroka.ToDouble();
				   //		stroka=FloatToStrF(StrToFloat(stroka),ffFixed,20,1);
					  }
					else stroka = Sh.OlePropertyGet("Cells",i,otkl).OlePropertyGet("Value");

				 /*	Sql = "update reit_ruk set \
										 otkl='"+ stroka +"',  \
										 otkl_ball=trim('"+ Sh.OlePropertyGet("Cells",i,otkl_ball) +"'),  \
										 info=trim('"+ Sh.OlePropertyGet("Cells",i,info) +"'),  \
										 info_ball=trim('"+ Sh.OlePropertyGet("Cells",i,info_ball) +"'),  \
										 c5=round(trim('"+ Sh.OlePropertyGet("Cells",i,c5) +"'),2),  \
										 c5_ball=trim('"+ Sh.OlePropertyGet("Cells",i,c5_ball) +"'),  \
										 kns=trim('"+ Sh.OlePropertyGet("Cells",i,kns) +"'),  \
										 kns_ball=trim('"+ Sh.OlePropertyGet("Cells",i,kns_ball) +"')  \
						  where tn="+ Sh.OlePropertyGet("Cells",i,tn)+" and god="+IntToStr(god) +" and kvart="+IntToStr(kvartal);
				  */
							Sql = "update reit_ruk set \
										 otkl=round(trim('"+ stroka +"'),2),  \
										 otkl_ball=trim('"+ Sh.OlePropertyGet("Cells",i,otkl_ball) +"'),  \
										 info=trim('"+ Sh.OlePropertyGet("Cells",i,info) +"'),  \
										 info_ball=trim('"+ Sh.OlePropertyGet("Cells",i,info_ball) +"'),  \
										 c5=trunc(trim('"+ Sh.OlePropertyGet("Cells",i,c5) +"')),  \
										 c5_ball=trim('"+ Sh.OlePropertyGet("Cells",i,c5_ball) +"'),  \
										 kns=trim('"+ Sh.OlePropertyGet("Cells",i,kns) +"'),  \
										 kns_ball=trim('"+ Sh.OlePropertyGet("Cells",i,kns_ball) +"')  \
						  where tn="+ Sh.OlePropertyGet("Cells",i,tn)+" and god="+IntToStr(god) +" and kvart="+IntToStr(kvartal);


				break;

				//������� ����������
				case 4:
					Sql = "update reit_ruk set \
										 priem=trim('"+ Sh.OlePropertyGet("Cells",i,priem) +"'),  \
										 priem_ball=trim('"+ Sh.OlePropertyGet("Cells",i,priem_ball) +"')  \
						  where tn="+ Sh.OlePropertyGet("Cells",i,tn)+" and god="+IntToStr(god) +" and kvart="+IntToStr(kvartal);
				break;

				//���
				case 5:
					Sql = "update reit_ruk set \
										 spp_kol=trim('"+ Sh.OlePropertyGet("Cells",i,spp_kol) +"'),  \
										 spp_ball=trim('"+ Sh.OlePropertyGet("Cells",i,spp_ball) +"')  \
						  where tn="+ Sh.OlePropertyGet("Cells",i,tn)+" and god="+IntToStr(god) +" and kvart="+IntToStr(kvartal);
				break;

				//������ �����
				case 6:
					Sql = "update reit_ruk set \
										 ot_upr=trim('"+ Sh.OlePropertyGet("Cells",i,ot_upr) +"')  \
						  where tn="+ Sh.OlePropertyGet("Cells",i,tn)+" and god="+IntToStr(god) +" and kvart="+IntToStr(kvartal);
				break;

				//���������� ���������� ��
				case 7:
					Sql = "update reit_ruk set \
										 ot_treb=trim('"+ Sh.OlePropertyGet("Cells",i,ot_treb).OlePropertyGet("Value") +"')  \
						  where tn="+ Sh.OlePropertyGet("Cells",i,tn)+" and god="+IntToStr(god) +" and kvart="+IntToStr(kvartal);
				break;

				//�������� ����������
				case 8:
					Sql = "update reit_ruk set \
										 trud_d=trim('"+ Sh.OlePropertyGet("Cells",i,trud_d).OlePropertyGet("Value") +"')  \
						  where tn="+ Sh.OlePropertyGet("Cells",i,tn)+" and god="+IntToStr(god) +" and kvart="+IntToStr(kvartal);
				break;
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
				  Application->MessageBox(("�������� ������ ��� ������� �������� ������ � ������� REIT_RUK: " + E.Message).c_str(),L"������",
											MB_OK+MB_ICONERROR);

				  DM->qReiting->Refresh();
				  StatusBar1->SimpleText ="  �������� ������: "+IntToStr(god)+" ���, "+IntToStr(kvartal)+" �������";
				  InsertLog("�������� ������ ��� �������� ������ �� "+str+" �� ����� '"+OpenDialog1->FileName+"' �� "+IntToStr(god)+" ���, "+IntToStr(kvartal)+" �������");
				  Cursor = crDefault;
				  ProgressBar->Visible = false;
				  Abort();
				}

			  rec++;
			  kol+=DM->qObnovlenie->RowsAffected;

			  // ���������� ����������� �������
			  if (DM->qObnovlenie->RowsAffected == 0)
				{
				  //������������ ��������� � ����� �������
				   if (pr!=2)
					 {
					   rtf_Out("z", " ",3);
					   if(!rtf_LineFeed())
						 {
						   MessageBox(Handle,L"������ ������ � ���� ������",L"������",8192);
						   if (!rtf_Close()) MessageBox(Handle,L"������ �������� ����� ������",L"������",8192);
						   return;
						 }
					 }

				   //������������ ������ �� ������������� �������
				   rtf_Out("tn", VarToStr(Sh.OlePropertyGet("Cells",i,tn)),4);
				   rtf_Out("fio", VarToStr(Sh.OlePropertyGet("Cells",i,fio)),4);

				   if(!rtf_LineFeed())
					 {
					   MessageBox(Handle,L"������ ������ � ���� ������",L"������",8192);
					   if (!rtf_Close()) MessageBox(Handle,L"������ �������� ����� ������",L"������",8192);
					   return;
					 }
				   pr=2;      //������� ������������ ����� ������
				   fotchet=1;  //������� ������������ ������ �� ������������� �������
				 }
			   else obnov_kol++;

			}

          ProgressBar->Position++;
		  ob_kol++;
		}


      DM->qReiting->Refresh();

	  //�������� Excel
	  AppEx.OleProcedure("Quit");
	  AppEx = Unassigned;


	  if(!rtf_Close())
		{
		  MessageBox(Handle,L"������ �������� ����� ������", L"������", 8192);
		  return;
		}

	  //������������ ������ � Word
	  if (fotchet==1)
		{
		  StatusBar1->SimpleText = "������������ ������ � ��������...";

		  //�������� �����, ���� �� �� ����������
		  ForceDirectories(WorkPath);

		  int istrd;
		  try
			{
			  rtf_CreateReport(TempPath + "\\zagruzka.txt", Path+"\\RTF\\zagruzka2.rtf",
							   WorkPath+"\\����� �� �������� ������.doc",NULL,&istrd);


			  WinExec(("\""+ WordPath+"\"\""+WorkPath+"\\����� �� �������� ������.doc\"").c_str(),SW_MAXIMIZE);

			}
		  catch(RepoRTF_Error E)
			{
			  Application->MessageBox(("������ ������������ ������:"+ String(E.Err)+
								 "\n������ ����� ������:"+IntToStr(istrd)).c_str(),
								 L"������",
								 MB_OK+MB_ICONERROR);
			}

		  Application->MessageBox(("���������� �������������� ���������� � ����������� ����� � ���� ������ �� ���������.\n��������� ������������� ���������� � ����� \n "+OpenDialog1->FileName+" � ��������� ��������� ��������").c_str() ,L" �������� ������",
								  MB_OK + MB_ICONINFORMATION);
		}

	  DeleteFile(TempPath+"\\otchet.txt");

	  InsertLog("�������� ������ �� "+str+" �� ����� '"+OpenDialog1->FileName+"' �� "+IntToStr(god)+" ���, "+IntToStr(kvartal)+" ������� ��������� �������. ��������� " + IntToStr(obnov_kol) + " �� " + IntToStr(ob_kol)+" �������");

	  Application->MessageBox(("�������� ������ �� "+WideString(str)+" ��������� �������. =) \n��������� " + IntToStr(obnov_kol) + " �� " + IntToStr(ob_kol)+" �������").c_str(),
							   L"���������� ������ �� ������ ���������",
						       MB_OK + MB_ICONINFORMATION);
	}
  catch(...)
	{
	  AppEx.OleProcedure("Quit");
	  //AppEx.Clear();
	  //VarClear(AppEx);
	  AppEx=Unassigned;
	  InsertLog("�������� ������ ��� �������� ������ �� "+str+" �� ����� '"+OpenDialog1->FileName+"' �� "+IntToStr(god)+" ���, "+IntToStr(kvartal)+" �������");
	}


  Cursor = crDefault;
  ProgressBar->Position = 0;
  ProgressBar->Visible = false;

  StatusBar1->SimplePanel = false;
  StatusBar1->SimpleText ="  �������� ������: "+IntToStr(god)+" ���, "+IntToStr(kvartal)+" �������";
}
//---------------------------------------------------------------------------




void __fastcall TMain::SpeedButton1Click(TObject *Sender)
{
   N6SpisokClick(Sender);
}
//---------------------------------------------------------------------------
//�������������� ������� ������ �� ���� ����������
void __fastcall TMain::SpeedButton3Click(TObject *Sender)
{
  if (Application->MessageBox(("����� �������� ������ ������ � �������� �� ���� ���������� \n�� "+IntToStr(god)+" ���, "+IntToStr(kvartal)+" �������. ����������?").c_str(), L"������ ��������",
							  MB_YESNO+MB_ICONINFORMATION )==ID_YES)
	{
      //������� ������
	  RaschetOcen(0);

	  //������� ��������
      RaschetReit(0, NULL, NULL);
	}

}
//---------------------------------------------------------------------------
//�������������� ������� ������ �� ������ ���������
void __fastcall TMain::N6ReitingClick(TObject *Sender)
{
  RaschetOcen(1);

  //������� ��������
  RaschetReit(1, DM->qReiting->FieldByName("zex")->AsString, DM->qReiting->FieldByName("podch")->AsInteger);
}
//---------------------------------------------------------------------------
//�������������� ������� ������
void __fastcall TMain::RaschetOcen(int pr)
{
  AnsiString Sql;

  StatusBar1->SimpleText ="���� ������ �������� �� ���� ���������� ��: "+IntToStr(god)+" ���, "+IntToStr(kvartal)+" �������";


  Sql = " update reit_ruk s set avt_ocenka =(case when (select distinct(pz) from sp_reit_proizv sp where sp.zex=s.zex)=1                                                                                                                       \
			then (nvl(proiz_ball,0)*0.1+nvl(kpe_ball,0)*0.15+nvl(otkl_ball,0)*0.15+nvl(priem_ball,0)*0.1+nvl(info_ball,0)*0.1+nvl(kns_ball,0)*0.1+nvl(c5_ball,0)*0.1+nvl(spp_ball,0)*0.1+nvl(ot_upr,0)*0.1) - nvl(ot_treb,0)-nvl(trud_d,0)                                         \
			else (nvl(kpe_ball,0)*0.2+nvl(otkl_ball,0)*0.2+nvl(priem_ball,0)*0.1+nvl(info_ball,0)*0.1+nvl(kns_ball,0)*0.1+nvl(c5_ball,0)*0.1+nvl(spp_ball,0)*0.1+nvl(ot_upr,0)*0.1) - nvl(ot_treb,0)-nvl(trud_d,0) end),                                                      \
			ocenka = (case when (select distinct(pz) from sp_reit_proizv sp where sp.zex=s.zex)=1                                                                                                                                              \
			then (nvl(proiz_ball,0)*0.1+nvl(kpe_ball,0)*0.15+nvl(otkl_ball,0)*0.15+nvl(priem_ball,0)*0.1+nvl(info_ball,0)*0.1+nvl(kns_ball,0)*0.1+nvl(c5_ball,0)*0.1+nvl(spp_ball,0)*0.1+nvl(ot_upr,0)*0.1) - nvl(ot_treb,0)-nvl(trud_d,0)                                         \
			else (nvl(kpe_ball,0)*0.2+nvl(otkl_ball,0)*0.2+nvl(priem_ball,0)*0.1+nvl(info_ball,0)*0.1+nvl(kns_ball,0)*0.1+nvl(c5_ball,0)*0.1+nvl(spp_ball,0)*0.1+nvl(ot_upr,0)*0.1) - nvl(ot_treb,0)-nvl(trud_d,0) end )                                                      \
		 where god="+IntToStr(god) +" and kvart="+IntToStr(kvartal)+" and ocenka is null";                                                                                                                                                                           \

  if (pr==1) //Sql+=" and tn="+DM->qReiting->FieldByName("tn")->AsString;
  Sql+= " and zex="+DM->qReiting->FieldByName("zex")->AsString+" and podch="+DM->qReiting->FieldByName("podch")->AsString;
																																																											   \
  DM->qObnovlenie->Close();
  DM->qObnovlenie->SQL->Clear();
  DM->qObnovlenie->SQL->Add(Sql);
  try
	{
	  DM->qObnovlenie->ExecSQL();
	}
  catch (Exception &E)
	{
	  Application->MessageBox(("�������� ������ ��� �������� �������� (������� REIT_RUK) "+E.Message).c_str(),L"������",
							  MB_OK+MB_ICONERROR);

	  StatusBar1->SimpleText ="�������� ������: "+IntToStr(god)+" ���, "+IntToStr(kvartal)+" �������";
	  if (pr==1) InsertLog("������ �������� �� ��������� � ���.� = "+DM->qReiting->FieldByName("tn")->AsString+" �� ��������");
	  else InsertLog("������ �������� �� ���� ���������� �� ��������");
	  Abort();
	}

   if (pr==1) InsertLog("������ ������ �� ��������� � ���.� = "+DM->qReiting->FieldByName("tn")->AsString+" �������� �������");
   else InsertLog("������ ������ �� ���� ���������� �������� �������");
  // Application->MessageBox(L"������ ������ �� ���� ���������� �������� �������!!!" ,L"������ ��������",
  //								  MB_OK + MB_ICONINFORMATION);

    DM->qReiting->Refresh();

}
//---------------------------------------------------------------------------

//�������������� ������� ��������
void __fastcall TMain::RaschetReit(int pr, String zex, int podch)
{
  AnsiString Sql;
  int kol_kr_zona=0, kol_zl_zona=0;


  //����� ������ ������ ����� ��� �������������
  Sql = "select * from (                                                       \
						select distinct zex,                                   \
							   nvl(podch,0) as podch,                                   \
							   count(*) over (partition by zex, nvl(podch,0)) as kol_zex,    \
							   min(ocenka) over (partition by zex, nvl(podch,0)) as zn_min,  \
							   max(ocenka) over (partition by zex, nvl(podch,0)) as zn_max,  \
							   (count(*) over (partition by zex, nvl(podch,0)))*0.2 as zona  \
						from reit_ruk s where god="+IntToStr(god) +" and kvart="+IntToStr(kvartal);

  if (pr==1) Sql+=" and zex="+DM->qReiting->FieldByName("zex")->AsString+" and nvl(podch,0)="+DM->qReiting->FieldByName("podch")->AsString;

  Sql+=" ) where kol_zex>4 order by zex, nvl(podch,0)";

  DM->qObnovlenie->Close();
  DM->qObnovlenie->SQL->Clear();
  DM->qObnovlenie->SQL->Add(Sql);
  try
	{
	  DM->qObnovlenie->Open();
	}
  catch (Exception &E)
	{
	  Application->MessageBox(("�������� ������ ��� �������� �������� (������� REIT_RUK) "+E.Message).c_str(),L"������",
							  MB_OK+MB_ICONERROR);

	  StatusBar1->SimpleText ="�������� ������: "+IntToStr(god)+" ���, "+IntToStr(kvartal)+" �������";
	  if (pr==1) InsertLog("������ �������� �� ��������� � ���.� = "+DM->qReiting->FieldByName("tn")->AsString+" �� ��������");
	  else InsertLog("������ �������� �� ���� ���������� �� ��������");
	  Abort();
	}

  if (DM->qObnovlenie->RecordCount==0)
  {
	 Application->MessageBox(L"��� ���������� ��� ������������� ��������������� ������� ������������� (����������� ���������� ��� ������������� 5 ������� � �������� ������ ������������ ������������� � ����� �� ��������� ���������� (� ������������ � ��� �����������))",L"��������������",
							 MB_OK+MB_ICONINFORMATION);
	  StatusBar1->SimpleText ="�������� ������: "+IntToStr(god)+" ���, "+IntToStr(kvartal)+" �������";
	  if (pr==1) InsertLog("������ �������� �� ��������� � ���.� = "+DM->qReiting->FieldByName("tn")->AsString+" �� ��������, � ����� � ����������� ����������, ��������������� ������� �������������");
	  else InsertLog("������ �������� �� ���� ���������� �� ��������, � ����� � ����������� ����������, ��������������� ������� �������������");
	  Abort();
  }

  while (!DM->qObnovlenie->Eof)
	{
	  //������� �������� �� �������� ��� ���������
	  //����� ������ ������ ����� ��� �������������
	  Sql = "update reit_ruk set reit = NULL                                      \
			 where god="+IntToStr(god) +" and kvart="+IntToStr(kvartal)+"           \
			 and  zex ="+DM->qObnovlenie->FieldByName("zex")->AsString+"          \
			 and nvl(podch,0)= "+DM->qObnovlenie->FieldByName("podch")->AsString;

	  DM->qObnovlenie2->Close();
	  DM->qObnovlenie2->SQL->Clear();
	  DM->qObnovlenie2->SQL->Add(Sql);
	  try
		{
		  DM->qObnovlenie2->ExecSQL();
		}
	  catch (Exception &E)
		{
		  Application->MessageBox(("�������� ������ ��� �������� �������� (������� REIT_RUK) "+E.Message).c_str(),L"������",
								  MB_OK+MB_ICONERROR);

		  StatusBar1->SimpleText ="�������� ������: "+IntToStr(god)+" ���, "+IntToStr(kvartal)+" �������";
		  if (pr==1) InsertLog("������ �������� �� ��������� � ���.� = "+DM->qReiting->FieldByName("tn")->AsString+" �� ��������");
		  else InsertLog("������ �������� �� ���� ���������� �� ��������");
		  Abort();
		}

	   kol_kr_zona=0;

	  //����� ������ ���������� �� ������������ ����
	  DM->qRaschet->Active = false;
	  DM->qRaschet->ParamByName("pgod")->Value = IntToStr(god);
	  DM->qRaschet->ParamByName("pkvart")->Value = IntToStr(kvartal);
	  DM->qRaschet->ParamByName("pzex")->Value = DM->qObnovlenie->FieldByName("zex")->AsString;
	  DM->qRaschet->ParamByName("ppodch")->Value = DM->qObnovlenie->FieldByName("podch")->AsString;

	  try
		{
		  DM->qRaschet->Open();
		}
	  catch (Exception &E)
		{
		  Application->MessageBox(("�������� ������ ��� �������� �������� (������� REIT_RUK) "+E.Message).c_str(),L"������",
									  MB_OK+MB_ICONERROR);

		  StatusBar1->SimpleText ="�������� ������: "+IntToStr(god)+" ���, "+IntToStr(kvartal)+" �������";
		  if (pr==1) InsertLog("������ �������� �� ��������� � ���.� = "+DM->qReiting->FieldByName("tn")->AsString+" �� ��������");
		  else InsertLog("������ �������� �� ���� ���������� �� ��������");
		  Abort();
		}

	  if (DM->qRaschet->RecordCount>0)
		{
		  //********************************************************************
		  //���������� ������� ���� ���� �� ����� 20% �� �����
		  while (kol_kr_zona<DM->qObnovlenie->FieldByName("zona")->AsInteger && DM->qRaschet->FieldByName("kol_zex")->AsInteger>0)
			{

			  Sql = "update reit_ruk                                            \
								set reit = 3                                    \
					 where zex = "+QuotedStr(DM->qRaschet->FieldByName("zex")->AsString)+" \
					 and ocenka="+DM->qRaschet->FieldByName("zn_min")->AsString+"\
					 and nvl(podch,0) = "+DM->qRaschet->FieldByName("podch")->AsString+"\
					 and god="+IntToStr(god) +" and kvart="+IntToStr(kvartal);

			  DM->qObnovlenie2->Close();
			  DM->qObnovlenie2->SQL->Clear();
			  DM->qObnovlenie2->SQL->Add(Sql);
			  try
				{
				  DM->qObnovlenie2->ExecSQL();
				}
			  catch (Exception &E)
				{
				  Application->MessageBox(("�������� ������ ��� �������� �������� (������� REIT_RUK) "+E.Message).c_str(),L"������",
										  MB_OK+MB_ICONERROR);

				  StatusBar1->SimpleText ="�������� ������: "+IntToStr(god)+" ���, "+IntToStr(kvartal)+" �������";
				  if (pr==1) InsertLog("������ �������� �� ��������� � ���.� = "+DM->qReiting->FieldByName("tn")->AsString+" �� ��������");
				  else InsertLog("������ �������� �� ���� ���������� �� ��������");
				  Abort();
				}

			  //�������� �� ������� 20% �� ���������� � ������� ����
			  kol_kr_zona+= DM->qObnovlenie2->RowsAffected;

			  DM->qRaschet->Refresh();

			}
		  //********************************************************************
		  //���������� ������� ���� ���� �� ����� 20% �� �����
		  kol_zl_zona = 0;
		  while (kol_zl_zona<DM->qObnovlenie->FieldByName("zona")->AsInteger && DM->qRaschet->FieldByName("kol_zex")->AsInteger>0)
			{

			  //�������� ��������� �� ���������� ���������� � ������������ ��������� 20%
			  Sql = "select count(*) as kol from reit_ruk                                    \
					 where ocenka = "+DM->qRaschet->FieldByName("zn_max")->AsString+"     \
					 and zex = "+DM->qRaschet->FieldByName("zex")->AsString+"             \
					 and nvl(podch,0) = "+DM->qRaschet->FieldByName("podch")->AsString+" \
					 and god="+IntToStr(god) +" and kvart="+IntToStr(kvartal);

			  DM->qObnovlenie2->Close();
			  DM->qObnovlenie2->SQL->Clear();
			  DM->qObnovlenie2->SQL->Add(Sql);
			  try
				{
				  DM->qObnovlenie2->Open();
				}
			  catch (Exception &E)
				{
				  Application->MessageBox(("�������� ������ ��� �������� �������� (������� REIT_RUK) "+E.Message).c_str(),L"������",
										  MB_OK+MB_ICONERROR);

				  StatusBar1->SimpleText ="�������� ������: "+IntToStr(god)+" ���, "+IntToStr(kvartal)+" �������";
				  if (pr==1) InsertLog("������ �������� �� ��������� � ���.� = "+DM->qReiting->FieldByName("tn")->AsString+" �� ��������");
				  else InsertLog("������ �������� �� ���� ���������� �� ��������");
				  Abort();
				}

			  if (DM->qObnovlenie2->FieldByName("kol")->AsInteger+kol_zl_zona<=DM->qObnovlenie->FieldByName("zona")->AsInteger)
				{
				  //�������� ������� ����
				  Sql = "update reit_ruk                                                \
								set reit = 1                                    \
						 where zex = "+DM->qRaschet->FieldByName("zex")->AsString+"  \
						 and ocenka="+DM->qRaschet->FieldByName("zn_max")->AsString+"   \
						 and nvl(podch,0) = "+DM->qRaschet->FieldByName("podch")->AsString+"  \
						 and god="+IntToStr(god) +" and kvart="+IntToStr(kvartal);

				  DM->qObnovlenie2->Close();
				  DM->qObnovlenie2->SQL->Clear();
				  DM->qObnovlenie2->SQL->Add(Sql);
				  try
					{
					  DM->qObnovlenie2->ExecSQL();
					}
				  catch (Exception &E)
					{
					  Application->MessageBox(("�������� ������ ��� �������� �������� (������� REIT_RUK) "+E.Message).c_str(),L"������",
										  MB_OK+MB_ICONERROR);

					  StatusBar1->SimpleText ="�������� ������: "+IntToStr(god)+" ���, "+IntToStr(kvartal)+" �������";
					  if (pr==1) InsertLog("������ �������� �� ��������� � ���.� = "+DM->qReiting->FieldByName("tn")->AsString+" �� ��������");
					  else InsertLog("������ �������� �� ���� ���������� �� ��������");
					  Abort();
					}

				 //�������� �� ������� 20% �� ���������� � ������� ����
				 kol_zl_zona+= DM->qObnovlenie2->RowsAffected;

				 DM->qRaschet->Refresh();

				}
			  else
				{
				  kol_zl_zona+= DM->qObnovlenie2->FieldByName("kol")->AsInteger;
				}

			}
		  //********************************************************************
		  //���������� ������ ����
		  if (DM->qRaschet->FieldByName("kol_zex")->AsInteger>0)
			{
			  Sql = "update reit_ruk                                                \
								set reit = 2                                    \
					 where zex = "+DM->qRaschet->FieldByName("zex")->AsString+"  \
					 and reit is null \
					 and nvl(podch,0) = "+DM->qRaschet->FieldByName("podch")->AsString+"  \
					 and god="+IntToStr(god) +" and kvart="+IntToStr(kvartal);

			  DM->qObnovlenie2->Close();
			  DM->qObnovlenie2->SQL->Clear();
			  DM->qObnovlenie2->SQL->Add(Sql);
			  try
				{
				  DM->qObnovlenie2->ExecSQL();
				}
			  catch (Exception &E)
				{
				  Application->MessageBox(("�������� ������ ��� �������� �������� (������� REIT_RUK) "+E.Message).c_str(),L"������",
											  MB_OK+MB_ICONERROR);

				  StatusBar1->SimpleText ="�������� ������: "+IntToStr(god)+" ���, "+IntToStr(kvartal)+" �������";
				  if (pr==1) InsertLog("������ �������� �� ��������� � ���.� = "+DM->qReiting->FieldByName("tn")->AsString+" �� ��������");
				  else InsertLog("������ �������� �� ���� ���������� �� ��������");
				  Abort();
				}
			}

			Application->MessageBox(L"������ �������� �������� �������! ",L"������ ��������",
											  MB_OK+MB_ICONINFORMATION);

		}
		DM->qObnovlenie->Next();
	}

   DM->qReiting->Refresh();
}
//---------------------------------------------------------------------------
void __fastcall TMain::SpeedButton4Click(TObject *Sender)
{
//Vvod->ShowModal();
}
//---------------------------------------------------------------------------
//������������ ��������� ������ � ������������
void __fastcall TMain::N15Click(TObject *Sender)
{
  OtchetExcelItog(0);
}
//---------------------------------------------------------------------------
//������������ ��������� ������ ��� �����������
void __fastcall TMain::jjjj1Click(TObject *Sender)
{
  OtchetExcelItog(1);
}
//---------------------------------------------------------------------------
//������������ ��������� ������
void __fastcall TMain::OtchetExcelItog(int otchet)
{
  AnsiString Sql, sFile;
  int i,n;
  Variant AppEx,Sh;

  StatusBar1->SimpleText ="  ���� ������������ ��������� ������...";

  Sql="select * from reit_ruk where god="+IntToStr(god) +" and kvart="+IntToStr(kvartal)+" and nvl(reit,0)>0";

  if (otchet==1) Sql+= " and nvl(podch,0)=0 ";
  else Sql+= " and nvl(podch,0)>0 ";

  Sql+=" order by zex, ocenka desc, reit desc, tn ";

  DM->qObnovlenie->Close();
  DM->qObnovlenie->SQL->Clear();
  DM->qObnovlenie->SQL->Add(Sql);
  try
	{
	  DM->qObnovlenie->Open();
	}
  catch (Exception &E)
	{
	  Application->MessageBox(("�������� ������ ��� ������� ������ �� ������� �� ������������� REIT_RUK "+E.Message).c_str(),L"������",
							  MB_OK+MB_ICONERROR);

	  StatusBar1->SimpleText ="  �������� ������: "+IntToStr(god)+" ���, "+IntToStr(kvartal)+" �������";
	  Abort();
	}

  if (DM->qObnovlenie->RecordCount!=0)
	{


  Cursor = crHourGlass;
  ProgressBar->Position = 0;
  ProgressBar->Visible = true;
  ProgressBar->Max=DM->qObnovlenie->RecordCount;

  // �������������� Excel, ��������� ���� ������
  try
	{
	  AppEx=CreateOleObject("Excel.Application");
	}
  catch (...)
    {
      Application->MessageBox(L"���������� ������� Microsoft Excel!"
							  " �������� ��� ���������� �� ���������� �� �����������.",L"������",MB_OK+MB_ICONERROR);
	  StatusBar1->SimpleText ="  �������� ������: "+IntToStr(god)+" ���, "+IntToStr(kvartal)+" �������";
	  Cursor = crDefault;
	  ProgressBar->Visible = false;
	  Abort();
    }

  //���� ��������� ������ �� ����� ������������ ������
  try
    {
      try
        {
		  AppEx.OlePropertySet("AskToUpdateLinks",false);
          AppEx.OlePropertySet("DisplayAlerts",false);

          //�������� �����, ���� �� �� ����������
          ForceDirectories(WorkPath);

		  //�������� ������ ����� � ��� ���������
		  CopyFile(WideString(Path+"\\RTF\\itogovaya_tablica.xlsx").c_bstr(), WideString(WorkPath+"\\�������� ������� ���������� �������� ������.xlsx").c_bstr(), false);
		  //sFile = WorkPath+"\\�������� ������� ���������� �������� ������.xlsx";

		  AppEx.OlePropertyGet("WorkBooks").OleProcedure("Open",WideString(WorkPath+"\\�������� ������� ���������� �������� ������.xlsx").c_bstr());  //��������� �����, ������ � ���
		  Sh=AppEx.OlePropertyGet("WorkSheets",1);                               //�������� � ��������� ����� �����
		  //Sh=AppEx.OlePropertyGet("WorkSheets","������");                      //�������� ���� �� ������������
		}
	  catch(...)
		{
		  Application->MessageBox(L"������ �������� ����� Microsoft Excel!",L"������",MB_OK+MB_ICONERROR);
		  StatusBar1->SimpleText ="  �������� ������: "+IntToStr(god)+" ���, "+IntToStr(kvartal)+" �������";
		  Cursor = crDefault;
		  ProgressBar->Visible = false;
		  Abort();
        }


	  i=1;
	  n=8;
      //����� ����
	  Sh.OlePropertyGet("Cells",2,1).OlePropertySet("Value", WideString(IntToStr(kvartal)+" �������, "+IntToStr(god) +" ���" ));

	  //����� ������ � ������
      Variant Massiv;
	  Massiv = VarArrayCreate(OPENARRAY(int,(0,27)),varVariant); //������ �� 26 ���������

	  while (!DM->qObnovlenie->Eof)
		{
		  Massiv.PutElement(i, 0);
		  Massiv.PutElement(DM->qObnovlenie->FieldByName("zex")->AsString, 1);
		  Massiv.PutElement(DM->qObnovlenie->FieldByName("tn")->AsString, 2);
		  Massiv.PutElement(DM->qObnovlenie->FieldByName("fio")->AsString, 3);
		  Massiv.PutElement(DM->qObnovlenie->FieldByName("dolg")->AsString, 4);
		  Massiv.PutElement(DM->qObnovlenie->FieldByName("id_dolg")->AsString, 5);


		  if (!DM->qObnovlenie->FieldByName("proiz")->AsString.IsEmpty()) Massiv.PutElement(DM->qObnovlenie->FieldByName("proiz")->AsFloat/100, 6);
		  else  Massiv.PutElement("", 6);
		  if (!DM->qObnovlenie->FieldByName("proiz_ball")->AsString.IsEmpty()) Massiv.PutElement(DM->qObnovlenie->FieldByName("proiz_ball")->AsFloat, 7);
		  else  Massiv.PutElement("", 7);
		  if (!DM->qObnovlenie->FieldByName("kpe")->AsString.IsEmpty()) Massiv.PutElement(DM->qObnovlenie->FieldByName("kpe")->AsFloat/100, 8);
		  else  Massiv.PutElement("", 8);
		  if (!DM->qObnovlenie->FieldByName("kpe_ball")->AsString.IsEmpty()) Massiv.PutElement(DM->qObnovlenie->FieldByName("kpe_ball")->AsFloat, 9);
		  else  Massiv.PutElement("", 9);
		  if (!DM->qObnovlenie->FieldByName("otkl")->AsString.IsEmpty()) Massiv.PutElement(DM->qObnovlenie->FieldByName("otkl")->AsFloat/100, 10);
		  else  Massiv.PutElement("", 10);
		  if (!DM->qObnovlenie->FieldByName("otkl_ball")->AsString.IsEmpty()) Massiv.PutElement(DM->qObnovlenie->FieldByName("otkl_ball")->AsFloat, 11);
		  else  Massiv.PutElement("", 11);
		  if (!DM->qObnovlenie->FieldByName("priem")->AsString.IsEmpty()) Massiv.PutElement(DM->qObnovlenie->FieldByName("priem")->AsFloat, 12);
		  else  Massiv.PutElement("", 12);
		  if (!DM->qObnovlenie->FieldByName("priem_ball")->AsString.IsEmpty()) Massiv.PutElement(DM->qObnovlenie->FieldByName("priem_ball")->AsFloat, 13);
		  else  Massiv.PutElement("", 13);
		  if (!DM->qObnovlenie->FieldByName("info")->AsString.IsEmpty()) Massiv.PutElement(DM->qObnovlenie->FieldByName("info")->AsFloat, 14);
		  else  Massiv.PutElement("", 14);
		  if (!DM->qObnovlenie->FieldByName("info_ball")->AsString.IsEmpty()) Massiv.PutElement(DM->qObnovlenie->FieldByName("info_ball")->AsFloat, 15);
		  else  Massiv.PutElement("", 15);
		  if (!DM->qObnovlenie->FieldByName("c5")->AsString.IsEmpty()) Massiv.PutElement(DM->qObnovlenie->FieldByName("c5")->AsFloat, 16);
		  else  Massiv.PutElement("", 16);
		  if (!DM->qObnovlenie->FieldByName("c5_ball")->AsString.IsEmpty()) Massiv.PutElement(DM->qObnovlenie->FieldByName("c5_ball")->AsFloat, 17);
		  else  Massiv.PutElement("", 17);
		  if (!DM->qObnovlenie->FieldByName("kns")->AsString.IsEmpty()) Massiv.PutElement(DM->qObnovlenie->FieldByName("kns")->AsFloat, 18);
		  else  Massiv.PutElement("", 18);
		  if (!DM->qObnovlenie->FieldByName("kns_ball")->AsString.IsEmpty()) Massiv.PutElement(DM->qObnovlenie->FieldByName("kns_ball")->AsFloat, 19);
		  else  Massiv.PutElement("", 19);
		  if (!DM->qObnovlenie->FieldByName("spp_kol")->AsString.IsEmpty()) Massiv.PutElement(DM->qObnovlenie->FieldByName("spp_kol")->AsFloat, 20);
		  else  Massiv.PutElement("", 20);
		  if (!DM->qObnovlenie->FieldByName("spp_ball")->AsString.IsEmpty()) Massiv.PutElement(DM->qObnovlenie->FieldByName("spp_ball")->AsFloat, 21);
		  else  Massiv.PutElement("", 21);
		  if (!DM->qObnovlenie->FieldByName("ot_upr")->AsString.IsEmpty()) Massiv.PutElement(DM->qObnovlenie->FieldByName("ot_upr")->AsFloat, 22);
		  else  Massiv.PutElement("", 22);
		  if (!DM->qObnovlenie->FieldByName("ot_treb")->AsString.IsEmpty()) Massiv.PutElement(DM->qObnovlenie->FieldByName("ot_treb")->AsFloat, 23);
		  else  Massiv.PutElement("", 23);
		  if (!DM->qObnovlenie->FieldByName("trud_d")->AsString.IsEmpty()) Massiv.PutElement(DM->qObnovlenie->FieldByName("trud_d")->AsFloat, 24);
		  else  Massiv.PutElement("", 24);
		  if (!DM->qObnovlenie->FieldByName("ocenka")->AsString.IsEmpty()) Massiv.PutElement(DM->qObnovlenie->FieldByName("ocenka")->AsFloat, 25);
		  else  Massiv.PutElement("", 25);


		  Sh.OlePropertyGet("Range", WideString("A" + IntToStr(n) + ":Z" + IntToStr(n))).OlePropertySet("Value", Massiv); //������ � ������� � ������ A �� ������ Z
		  //	Sh.OlePropertyGet("Range", WideString("A8:Z30")).OlePropertySet("Value", Massiv); //������ � ������� � ������ A �� ������ Z


          //����������� ������� ����
		  if (DM->qObnovlenie->FieldByName("reit")->AsInteger==1) Sh.OlePropertyGet("Cells",n,26).OlePropertyGet("Interior").OlePropertySet("Color",0x00D6EFE4);
		  else if (DM->qObnovlenie->FieldByName("reit")->AsInteger==2) Sh.OlePropertyGet("Cells",n,26).OlePropertyGet("Interior").OlePropertySet("Color",0x00C2EAF5);
		  else if (DM->qObnovlenie->FieldByName("reit")->AsInteger==3) Sh.OlePropertyGet("Cells",n,26).OlePropertyGet("Interior").OlePropertySet("Color",0x00ECEEFF);
		  else Sh.OlePropertyGet("Cells",n,26).OlePropertyGet("Interior").OlePropertySet("Color",clWhite);


		  i++;
		  n++;
		  DM->qObnovlenie->Next();
          ProgressBar->Position++;
		}

      //������ �����
	  Sh.OlePropertyGet("Range",WideString("A8:Z"+IntToStr(n-1))).OlePropertyGet("Borders").OlePropertySet("LineStyle", 1);
																													   //xlContinuous
     // Sh.OleProcedure("SaveAs",vAsCurDir1.c_str());
     AppEx.OlePropertyGet("WorkBooks",1).OleFunction("Save");


      //������� ����� Excel � �������� ��� ������ ����������
     // AppEx.OlePropertyGet("WorkBooks",1).OleProcedure("Close");
	  Application->MessageBox(L"����� � Excel ������� �����������!", L"������������ ������",
							   MB_OK+MB_ICONINFORMATION);
	  //AppEx.OlePropertyGet("WorkBooks").OleProcedure("Open",vAsCurDir1.c_str());
	  AppEx.OlePropertySet("Visible",true);
	  AppEx.OlePropertySet("AskToUpdateLinks",true);
	  AppEx.OlePropertySet("DisplayAlerts",true);

	  StatusBar1->SimpleText= "  ������������ ������ ���������.";
	}
  catch(...)
	{
	  //������� �������� ���������� Excel
	  AppEx.OleProcedure("Quit");
	  AppEx = Unassigned;
	}


  ProgressBar->Position=0;
  ProgressBar->Visible = false;

	}
  else
	{
	  Application->MessageBox(("��� ���������� ��������� ������������� �� "+IntToStr(god)+" ���, "+IntToStr(kvartal)+" �������").c_str(),L"��������������",
							  MB_OK+MB_ICONINFORMATION);

	  StatusBar1->SimpleText ="  �������� ������: "+IntToStr(god)+" ���, "+IntToStr(kvartal)+" �������";
	}

  //***************************************************************************
  //����� �� �� �������������� ����������
  Sql="select * from reit_ruk \
	   where god="+IntToStr(god) +" and kvart="+IntToStr(kvartal)+" and nvl(reit,0)=0";

  if (otchet==1) Sql+= " and nvl(podch,0)=0 ";
  else Sql+= " and nvl(podch,0)>0 ";

  Sql+=" order by zex, ocenka desc, reit desc, tn ";


  DM->qObnovlenie->Close();
  DM->qObnovlenie->SQL->Clear();
  DM->qObnovlenie->SQL->Add(Sql);
  try
	{
	  DM->qObnovlenie->Open();
	}
  catch (Exception &E)
	{
	  Application->MessageBox(("�������� ������ ��� ������� ������ �� ������� �� ������������� REIT_RUK "+E.Message).c_str(),L"������",
							  MB_OK+MB_ICONERROR);

	  StatusBar1->SimpleText ="  �������� ������: "+IntToStr(god)+" ���, "+IntToStr(kvartal)+" �������";
	  Abort();
	}

  if (DM->qObnovlenie->RecordCount==0)
	{
	  Application->MessageBox(("��� ��������� ������ ������������� �� "+IntToStr(god)+" ���, "+IntToStr(kvartal)+" �������").c_str(),L"��������������",
							  MB_OK+MB_ICONINFORMATION);

	  StatusBar1->SimpleText ="  �������� ������: "+IntToStr(god)+" ���, "+IntToStr(kvartal)+" �������";
	  Abort();
	}

  StatusBar1->SimpleText ="  ���� ������������ ������ �� �� �������������� ����������...";
  Cursor = crHourGlass;
  ProgressBar->Position = 0;
  ProgressBar->Visible = true;
  ProgressBar->Max=DM->qObnovlenie->RecordCount;

    //�������� ����� ������ ��� ������ �� ����������� ������
  if (!rtf_Open((TempPath + "\\ne_reit.txt").c_str()))
	{
	  MessageBox(Handle,L"������ �������� ����� ������",L"������",8192);
	  Abort();
	}

  rtf_Out("data", DateTimeToStr(Now()),0);

  while (!DM->qObnovlenie->Eof)
	{
	  //������������ ������ �� ������������� �������
	  rtf_Out("zex", DM->qObnovlenie->FieldByName("zex")->AsString,1);
	  rtf_Out("tn", DM->qObnovlenie->FieldByName("tn")->AsString,1);
	  rtf_Out("dolg", DM->qObnovlenie->FieldByName("dolg")->AsString,1);
	  rtf_Out("fio", DM->qObnovlenie->FieldByName("fio")->AsString,1);
	  rtf_Out("ocen", DM->qObnovlenie->FieldByName("ocenka")->AsString,1);

	  if(!rtf_LineFeed())
		{
		  MessageBox(Handle,L"������ ������ � ���� ������",L"������",8192);
		  if (!rtf_Close()) MessageBox(Handle,L"������ �������� ����� ������",L"������",8192);
		  return;
		}

		DM->qObnovlenie->Next();
		ProgressBar->Position++;
	 }

  if(!rtf_Close())
	{
	  MessageBox(Handle,L"������ �������� ����� ������", L"������", 8192);
	  return;
	}

  //������������ ������ � Word
  StatusBar1->SimpleText = "  ������������ ������ �� �� �������������� ����������...";

  //�������� �����, ���� �� �� ����������
  ForceDirectories(WorkPath);

  int istrd;
  try
	{
	  rtf_CreateReport(TempPath + "\\ne_reit.txt", Path+"\\RTF\\ne_reit.rtf",
					   WorkPath+"\\����� �� �� �������������� ����������.doc",NULL,&istrd);

	  WinExec(("\""+ WordPath+"\"\""+WorkPath+"\\����� �� �� �������������� ����������.doc\"").c_str(),SW_MAXIMIZE);
	}
  catch(RepoRTF_Error E)
	{
	  Application->MessageBox(("������ ������������ ������:"+ String(E.Err)+
								 "\n������ ����� ������:"+IntToStr(istrd)).c_str(),
								 L"������",
								 MB_OK+MB_ICONERROR);


	}

  DeleteFile(TempPath+"\\ne_reit.txt");

  Cursor = crDefault;
  ProgressBar->Position=0;
  ProgressBar->Visible = false;
  StatusBar1->SimpleText ="  �������� ������: "+IntToStr(god)+" ���, "+IntToStr(kvartal)+" �������";


}
//---------------------------------------------------------------------------


void __fastcall TMain::DBGridEh1DrawColumnCell(TObject *Sender, const TRect &Rect,
		  int DataCol, TColumnEh *Column, TGridDrawState State)
{
  if (Prava=="ocen") {
	switch  (DM->qReiting->FieldByName("reit")->AsInteger)
	  {
		case 1: //�������
				((TDBGridEh *) Sender)->Canvas->Brush->Color = TColor(0x00D6EFE4);//0x00A3F1D1);//clInfoBk;
				((TDBGridEh *) Sender)->Canvas->Font->Color= clBlack;
				((TDBGridEh *) Sender)->DefaultDrawColumnCell(Rect, DataCol, Column, State);
		break;

		case 2: //������
				((TDBGridEh *) Sender)->Canvas->Brush->Color = TColor(0x00C2EAF5);//0x00A3F1D1);//clInfoBk;
				((TDBGridEh *) Sender)->Canvas->Font->Color= clBlack;
				((TDBGridEh *) Sender)->DefaultDrawColumnCell(Rect, DataCol, Column, State);
		break;

		case 3: //�������
				((TDBGridEh *) Sender)->Canvas->Brush->Color = TColor(0x00ECEEFF);//0x00A3F1D1);//clInfoBk;
				((TDBGridEh *) Sender)->Canvas->Font->Color= clBlack;
				((TDBGridEh *) Sender)->DefaultDrawColumnCell(Rect, DataCol, Column, State);
		break;

		default:
				((TDBGridEh *) Sender)->Canvas->Brush->Color = clWhite;//0x00A3F1D1);//clInfoBk;
				((TDBGridEh *) Sender)->Canvas->Font->Color= clBlack;
				((TDBGridEh *) Sender)->DefaultDrawColumnCell(Rect, DataCol, Column, State);
	  }

	// ��������� ������ �������� ������
	if (State.Contains(gdSelected))
	  {
		((TDBGridEh *) Sender)->Canvas->Brush->Color = TColor(0x008FDCEF);//0x00A3F1D1);//clInfoBk;
		((TDBGridEh *) Sender)->Canvas->Font->Color= clBlack;
		((TDBGridEh *) Sender)->DefaultDrawColumnCell(Rect, DataCol, Column, State);
	  }
  }
}
//---------------------------------------------------------------------------
 /*
void TMain::RebuildWindowRgn(TPanel *Panel)
{

  HRGN FullRgn, Rgn;
  int ClientX, ClientY;
  TControl *ChildControl;

  ClientX = (Panel->Width - Panel->ClientWidth) / 2;
  ClientY = Panel->Height - Panel->ClientHeight - ClientX;

  FullRgn = CreateRectRgn(0, 0, Panel->Width, Panel->Height);

  Rgn = CreateRectRgn(ClientX, ClientY, ClientX + Panel->ClientWidth, ClientY +
                      Panel->ClientHeight);
  CombineRgn(FullRgn, FullRgn, Rgn, RGN_DIFF);

  for(int i=0; i<Panel->ControlCount; i++)
  {
     ChildControl=Panel->Controls[i];
     Rgn=CreateRectRgn(ClientX + ChildControl->Left, ClientY + ChildControl->Top,
                         ClientX + ChildControl->Left + ChildControl->Width,
                         ClientY + ChildControl->Top + ChildControl->Height);
      CombineRgn(FullRgn, FullRgn, Rgn, RGN_OR);
  }

  SetWindowRgn(Panel->Handle, FullRgn, true);
}
//------------------------------------------------------------------------------
 */
void __fastcall TMain::FormResize(TObject *Sender)
{
  //������������ ������ �� ������

  if (Prava=="ocen")
	{
	  SpeedButton1->Left = Main->Width/2 - (SpeedButton1->Width)*2-3;
	  SpeedButton2->Left = Main->Width/2 - SpeedButton2->Width-1;
	  SpeedButton3->Left = Main->Width/2 +1;
	  SpeedButton4->Left = Main->Width/2 + SpeedButton4->Width +3;
	}
  else if (Prava=="unou")
	{
	  SpeedButton2->Left = Main->Width/2 - SpeedButton2->Width/2;
	  SpeedButton1->Left = SpeedButton2->Left - SpeedButton1->Width-2;
	  SpeedButton4->Left = SpeedButton2->Left + SpeedButton2->Width + 2;
	}
  else
	{
	  SpeedButton2->Left = Main->Width/2 - SpeedButton2->Width-1;
	  SpeedButton4->Left = Main->Width/2 +1;
	}



	if (Prava!="unou" && Prava!="ocen") DBGridEh1->AutoFitColWidths = true;
	else
	  {
        if (Main->Width<1500) DBGridEh1->AutoFitColWidths = false;
		else DBGridEh1->AutoFitColWidths = true;
	  }

	ProgressBar->Left = Main->Width-ProgressBar->Width-40;

}
//---------------------------------------------------------------------------

void __fastcall TMain::N5Click(TObject *Sender)
{
   Main->Close();
}
//---------------------------------------------------------------------------

void __fastcall TMain::N9Click(TObject *Sender)
{
  Sprav->ShowModal();
}
//---------------------------------------------------------------------------





