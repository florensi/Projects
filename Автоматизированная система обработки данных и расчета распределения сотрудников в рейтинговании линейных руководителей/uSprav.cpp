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
  //Сохранение/обновление данных
  AnsiString Sql, pz, rec, logi;

  if (ComboBoxPZ->Text=="да") pz=1;
  else pz=0;

  //Проверки
  if (EditZEX->Text.IsEmpty()) {
	Application->MessageBox(L"Не указан цех!!!",L"Предупреждение",
							MB_OK+MB_ICONWARNING);
	EditZEX->SetFocus();
  }

  if (ComboBoxPZ->Text.IsEmpty()) {
	Application->MessageBox(L"Не указан производственный план!!!",L"Предупреждение",
							MB_OK+MB_ICONWARNING);
	ComboBoxPZ->SetFocus();
  }

  if (LabelNZEX->Caption.IsEmpty()) {
	Application->MessageBox(L"Нет наименования цеха, возможно шифр цеха указан не верно!!!",L"Предупреждение",
							MB_OK+MB_ICONWARNING);
	EditZEX->SetFocus();
  }

  //Сохранение
  if (pr_in==0) {
	//Проверка на существование в базе такой записи
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
		 Application->MessageBox(L"Ошибка доступа к таблице SP_REIT_PROIZV",
								 L"Ошибка",MB_OK + MB_ICONERROR);
		 Abort();
	   }

	 if (DM->qObnovlenie->RecordCount!=0)
	   {
		 Application->MessageBox(L"Запись с указанным цехом уже существует в таблице.\nПроверьте правильность ввода шифра цеха",
								 L"Предупреждение",MB_OK+MB_ICONINFORMATION);
		 EditZEX->SetFocus();
		 Abort();
	   }

	 //Сохранение данных
	 Sql ="insert into sp_reit_proizv (zex, naim, pz) \
		   values ( \
					"+ QuotedStr(EditZEX->Text) +",\
					"+ QuotedStr(LabelNZEX->Caption) +",\
					"+ pz +")";
	 logi ="Добавление ";
  }
  //Обновление
  else if (pr_in==1) {
	if (szex!=EditZEX->Text ||
		 spz!=pz ||
		 snzex!=LabelNZEX->Caption) {

		//Обновление записи
		Sql = "update sp_reit_proizv set \
					   zex = "+ QuotedStr(EditZEX->Text) +",\
					   naim = "+ QuotedStr(LabelNZEX->Caption)+",\
					   pz = "+ pz+"\
				 where rowid = chartorowid("+ QuotedStr(DM->qSprav->FieldByName("rw")->AsString)+")";
		logi ="Редактирование ";
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
	  Application->MessageBox(L"Возникла ошибка при добавлении/редактировании записи в справочнике",
								 L"Ошибка ",MB_OK + MB_ICONERROR);
	  Main->InsertLog(logi + "выполнено с ошибкой в справочнике по производственному плану по цеху "+EditZEX->Text+" "+LabelNZEX->Caption);
	  Abort();
	}

  //Вернуть курсор
  DM->qSprav->Refresh();

  Main->InsertLog(logi+"выполнено успешно в справочнике по производственному плану по цеху "+ EditZEX->Text +" "+LabelNZEX->Caption);

  TLocateOptions SearchOptions;
  DM->qSprav->Locate("zex",rec,SearchOptions<<loPartialKey<<loCaseInsensitive);

  Application->MessageBox(L"Добавление/редактирование записи выполнено успешно :)",L"Предупреждение",
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
  //Добавить запись в спраочник
  EditZEX->Text = "";
  ComboBoxPZ->ItemIndex = -1;
  LabelNZEX->Caption = "";
  pr_in = 0;

  GroupBoxDOBAV->Caption = "Добавление нового цеха";
  BitBtn1->Caption = "Добавить";

  GroupBoxDOBAV->Visible = true;
  EditZEX->SetFocus();
}
//---------------------------------------------------------------------------

void __fastcall TSprav::N2Click(TObject *Sender)
{
 //Редактировать запись
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

  GroupBoxDOBAV->Caption = "Редактирование цеха";
  BitBtn1->Caption = "Редактировать";

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
	 //Проверка на существование в справочнике цеха
	 Sql ="select nazv_cexk from ssap_cex \
		   where id_cex="+QuotedStr(EditZEX->Text)+" and nazv_cexk not like '%(устар.)%'";

	 DM->qObnovlenie->Close();
	 DM->qObnovlenie->SQL->Clear();
	 DM->qObnovlenie->SQL->Add(Sql);

	 try
	   {
		 DM->qObnovlenie->Open();
	   }
	 catch(...)
	   {
		 Application->MessageBox(L"Ошибка доступа к таблице SP_REIT_PROIZV",
								 L"Ошибка",MB_OK + MB_ICONERROR);
		 Abort();
	   }

	LabelNZEX->Caption = DM->qObnovlenie->FieldByName("nazv_cexk")->AsString;
  }
}
//---------------------------------------------------------------------------

void __fastcall TSprav::N4Click(TObject *Sender)
{
  //Удаление записи
  if (Application->MessageBox(L"Вы действительно хотите удалить запись\n по выбранному цеху из справочника?",L"Предупреждение",
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
	  Application->MessageBox(L"Возникла ошибка при удалении записи из справочника по производственному плану",L"Ошибка",
							   MB_OK + MB_ICONERROR);
	}

  ShowMessage("Информация по производственному плану для выбранного \nцеха была успешно удалена из справочника");
  Main->InsertLog("Удаление записи из справочника по производственному заданию: цех "+ DM->qSprav->FieldByName("zex")->AsString );


  DM->qSprav->Refresh();
}
//---------------------------------------------------------------------------

