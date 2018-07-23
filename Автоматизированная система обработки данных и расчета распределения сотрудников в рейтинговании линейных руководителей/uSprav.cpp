//---------------------------------------------------------------------------

#include <vcl.h>
#pragma hdrstop

#include "uSprav.h"
#include "uDM.h"
#include "uMain.h"
//---------------------------------------------------------------------------
#pragma package(smart_init)
#pragma link "DBAxisGridsEh"
#pragma link "DBGridEh"
#pragma link "DBGridEhGrouping"
#pragma link "DBGridEhToolCtrls"
#pragma link "DynVarsEh"
#pragma link "EhLibVCL"
#pragma link "GridsEh"
#pragma link "ToolCtrlsEh"
#pragma resource "*.dfm"
TSprav *Sprav;
//---------------------------------------------------------------------------
__fastcall TSprav::TSprav(TComponent* Owner)
	: TForm(Owner)
{
}
//---------------------------------------------------------------------------
void __fastcall TSprav::BitBtn3Click(TObject *Sender)
{
   GroupBoxDOBAV->Visible = false;
}
//---------------------------------------------------------------------------

void __fastcall TSprav::BitBtn1Click(TObject *Sender)
{
  //����������/���������� ������
  AnsiString Sql, pz, rec, logi;

  if (ComboBoxPZ->Text=="��") pz=1;
  else pz=0;

  //��������
  if (EditZEX->Text.IsEmpty()) {
	Application->MessageBox(L"�� ������ ���!!!",L"��������������",
							MB_OK+MB_ICONWARNING);
	EditZEX->SetFocus();
  }

  if (ComboBoxPZ->Text.IsEmpty()) {
	Application->MessageBox(L"�� ������ ���������������� ����!!!",L"��������������",
							MB_OK+MB_ICONWARNING);
	ComboBoxPZ->SetFocus();
  }

  if (LabelNZEX->Caption.IsEmpty()) {
	Application->MessageBox(L"��� ������������ ����, �������� ���� ���� ������ �� �����!!!",L"��������������",
							MB_OK+MB_ICONWARNING);
	EditZEX->SetFocus();
  }

  //����������
  if (pr_in==0) {
	//�������� �� ������������� � ���� ����� ������
	 Sql ="select * from sp_reit_proizv \
		   where zex="+QuotedStr(EditZEX->Text);

	 DM->qObnovlenie->Close();
	 DM->qObnovlenie->SQL->Clear();
	 DM->qObnovlenie->SQL->Add(Sql);

	 try
	   {
		 DM->qObnovlenie->Open();
	   }
	 catch(...)
	   {
		 Application->MessageBox(L"������ ������� � ������� SP_REIT_PROIZV",
								 L"������",MB_OK + MB_ICONERROR);
		 Abort();
	   }

	 if (DM->qObnovlenie->RecordCount!=0)
	   {
		 Application->MessageBox(L"������ � ��������� ����� ��� ���������� � �������.\n��������� ������������ ����� ����� ����",
								 L"��������������",MB_OK+MB_ICONINFORMATION);
		 EditZEX->SetFocus();
		 Abort();
	   }

	 //���������� ������
	 Sql ="insert into sp_reit_proizv (zex, naim, pz) \
		   values ( \
					"+ QuotedStr(EditZEX->Text) +",\
					"+ QuotedStr(LabelNZEX->Caption) +",\
					"+ pz +")";
	 logi ="���������� ";
  }
  //����������
  else if (pr_in==1) {
	if (szex!=EditZEX->Text ||
		 spz!=pz ||
		 snzex!=LabelNZEX->Caption) {

		//���������� ������
		Sql = "update sp_reit_proizv set \
					   zex = "+ QuotedStr(EditZEX->Text) +",\
					   naim = "+ QuotedStr(LabelNZEX->Caption)+",\
					   pz = "+ pz+"\
				 where rowid = chartorowid("+ QuotedStr(DM->qSprav->FieldByName("rw")->AsString)+")";
		logi ="�������������� ";
	}
	else {
       DM->qSprav->Refresh();
	   GroupBoxDOBAV->Visible = false;
	   Abort();
	}
  }

  rec = EditZEX->Text;
  DM->qObnovlenie->Close();
  DM->qObnovlenie->SQL->Clear();
  DM->qObnovlenie->SQL->Add(Sql);

  try
	{
	  DM->qObnovlenie->ExecSQL();
	}
  catch(...)
	{
	  Application->MessageBox(L"�������� ������ ��� ����������/�������������� ������ � �����������",
								 L"������ ",MB_OK + MB_ICONERROR);
	  Main->InsertLog(logi + "��������� � ������� � ����������� �� ����������������� ����� �� ���� "+EditZEX->Text+" "+LabelNZEX->Caption);
	  Abort();
	}

  //������� ������
  DM->qSprav->Refresh();

  Main->InsertLog(logi+"��������� ������� � ����������� �� ����������������� ����� �� ���� "+ EditZEX->Text +" "+LabelNZEX->Caption);

  TLocateOptions SearchOptions;
  DM->qSprav->Locate("zex",rec,SearchOptions<<loPartialKey<<loCaseInsensitive);

  Application->MessageBox(L"����������/�������������� ������ ��������� ������� :)",L"��������������",
						   MB_OK+MB_ICONINFORMATION);

  GroupBoxDOBAV->Visible = false;
}
//---------------------------------------------------------------------------

void __fastcall TSprav::BitBtn2Click(TObject *Sender)
{
   Sprav->Close();
}
//---------------------------------------------------------------------------

void __fastcall TSprav::N1Click(TObject *Sender)
{
  //�������� ������ � ���������
  EditZEX->Text = "";
  ComboBoxPZ->ItemIndex = -1;
  LabelNZEX->Caption = "";
  pr_in = 0;

  GroupBoxDOBAV->Caption = "���������� ������ ����";
  BitBtn1->Caption = "��������";

  GroupBoxDOBAV->Visible = true;
  EditZEX->SetFocus();
}
//---------------------------------------------------------------------------

void __fastcall TSprav::N2Click(TObject *Sender)
{
 //������������� ������
  EditZEX->Text = szex = DM->qSprav->FieldByName("zex")->AsString;

  if (DM->qSprav->FieldByName("pz")->AsInteger==1) {
	ComboBoxPZ->ItemIndex = 0;
	spz = 1;
  }
  else{
	ComboBoxPZ->ItemIndex = 1;
	spz = 0;
  }

  LabelNZEX->Caption = snzex = DM->qSprav->FieldByName("naim")->AsString;
  pr_in = 1;

  GroupBoxDOBAV->Caption = "�������������� ����";
  BitBtn1->Caption = "�������������";

  GroupBoxDOBAV->Visible = true;
  EditZEX->SetFocus();

}
//---------------------------------------------------------------------------

void __fastcall TSprav::DBGridEh1DblClick(TObject *Sender)
{
  N2Click(Sender);
}
//---------------------------------------------------------------------------


void __fastcall TSprav::FormKeyDown(TObject *Sender, WORD &Key, TShiftState Shift)

{
   if (Key==VK_RETURN)
   FindNextControl((TWinControl *)Sender, true, true, false)->SetFocus();
}
//---------------------------------------------------------------------------

void __fastcall TSprav::EditZEXKeyPress(TObject *Sender, System::WideChar &Key)
{
  if (! (IsNumeric(Key) || Key=='.' || Key==',' || Key=='/' || Key=='\b') ) Key=0;
  if (Key==',' || Key=='/') Key='.';
}
//---------------------------------------------------------------------------

void __fastcall TSprav::EditZEXChange(TObject *Sender)
{
  AnsiString Sql;

  if (!EditZEX->Text.IsEmpty() && pr_in==0) {
	 //�������� �� ������������� � ����������� ����
	 Sql ="select nazv_cexk from ssap_cex \
		   where id_cex="+QuotedStr(EditZEX->Text)+" and nazv_cexk not like '%(�����.)%'";

	 DM->qObnovlenie->Close();
	 DM->qObnovlenie->SQL->Clear();
	 DM->qObnovlenie->SQL->Add(Sql);

	 try
	   {
		 DM->qObnovlenie->Open();
	   }
	 catch(...)
	   {
		 Application->MessageBox(L"������ ������� � ������� SP_REIT_PROIZV",
								 L"������",MB_OK + MB_ICONERROR);
		 Abort();
	   }

	LabelNZEX->Caption = DM->qObnovlenie->FieldByName("nazv_cexk")->AsString;
  }
}
//---------------------------------------------------------------------------

void __fastcall TSprav::N4Click(TObject *Sender)
{
  //�������� ������
  if (Application->MessageBox(L"�� ������������� ������ ������� ������\n �� ���������� ���� �� �����������?",L"��������������",
								MB_YESNO + MB_ICONWARNING)== IDNO)
	{
	  Abort();
	}

  AnsiString Sql = " delete from sp_reit_proizv \
					 where rowid = chartorowid("+ QuotedStr(DM->qSprav->FieldByName("rw")->AsString)+")";

  DM->qObnovlenie->Close();
  DM->qObnovlenie->SQL->Clear();
  DM->qObnovlenie->SQL->Add(Sql);
  try
	{
	  DM->qObnovlenie->ExecSQL();
	}
  catch(...)
	{
	  Application->MessageBox(L"�������� ������ ��� �������� ������ �� ����������� �� ����������������� �����",L"������",
							   MB_OK + MB_ICONERROR);
	}

  ShowMessage("���������� �� ����������������� ����� ��� ���������� \n���� ���� ������� ������� �� �����������");
  Main->InsertLog("�������� ������ �� ����������� �� ����������������� �������: ��� "+ DM->qSprav->FieldByName("zex")->AsString );


  DM->qSprav->Refresh();
}
//---------------------------------------------------------------------------

