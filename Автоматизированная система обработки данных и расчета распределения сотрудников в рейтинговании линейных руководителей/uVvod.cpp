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
void __fastcall TVvod::CanselClick(TObject *Sender)
{
   Vvod->Close();
}
//---------------------------------------------------------------------------

void __fastcall TVvod::FormKeyDown(TObject *Sender, WORD &Key, TShiftState Shift)

{
  if (Key==VK_RETURN)
    FindNextControl((TWinControl *)Sender, true, true, false)->SetFocus();
}
//---------------------------------------------------------------------------

void __fastcall TVvod::FormShow(TObject *Sender)
{
   if (Main->redakt==1) {
      SetNullEdit();
	  SetDataEdit();
	  EditZEX->SetFocus();
   }
   else{
	  SetNullEdit();
	  EditTN->SetFocus();
   }
}
//---------------------------------------------------------------------------

//Очищение Edit-ов
void __fastcall TVvod::SetNullEdit()
{
   LabelFIO->Caption = "";
   EditZEX->Text = "";
   EditTN->Text = "";

   LabelNZEX->Caption = "";
   EditID_DOLG->Text = "";
   EditDOLG->Text = "";

   CheckBoxPODCH->Checked = false;

   //производственное задание
   EditPROIZ->Text = "";
   EditPROIZ_BALL->Text = "";

   //КПЭ
   EditKPE->Text = "";
   EditKPE_BALL->Text = "";

   //наличие приемников
   EditPRIEM->Text = "";
   EditPRIEM_BALL->Text = "";

   //эффективное информирование
   EditINFO->Text = "";
   EditINFO_BALL->Text = "";

   //разбор отклонений
   EditOTKL->Text = "";
   EditOTKL_BALL->Text = "";

   //уровень 5С
   EditC5->Text = "";
   EditC5_BALL->Text = "";

   //КНС
   EditKNS->Text = "";
   EditKNS_BALL->Text = "";

   //СПП
   EditSPP_KOL->Text = "";
   EditSPP_BALL->Text = "";

   EditOT_UPR->Text = "";
   EditOT_TREB->Text = "";
   EditTRUD_D->Text = "";

   EditOCENKA->Text = "";

   ComboBoxREIT->ItemIndex = -1;
}
//---------------------------------------------------------------------------
//Заполнение Edit-ов данными
void __fastcall TVvod::SetDataEdit()
{
   LabelFIO->Caption = zfio = DM->qReiting->FieldByName("fio")->AsString;
   EditZEX->Text = zzex = DM->qReiting->FieldByName("zex")->AsString;
   EditTN->Text = ztn = DM->qReiting->FieldByName("tn")->AsString;

   LabelNZEX->Caption = DM->qReiting->FieldByName("zex_naim")->AsString;
   EditID_DOLG->Text = zid_dolg = DM->qReiting->FieldByName("id_dolg")->AsString;
   EditDOLG->Text = zdolg = DM->qReiting->FieldByName("dolg")->AsString;
   if ( DM->qReiting->FieldByName("podch")->AsInteger==1) {
	  zpodch = 1;
	  CheckBoxPODCH->Checked = true;
   }
   else{
	  zpodch = 0;
	  CheckBoxPODCH->Checked = false;
   }

   //производственное задание
   EditPROIZ->Text = zproiz = DM->qReiting->FieldByName("proiz")->AsString;
   EditPROIZ_BALL->Text = zproiz_ball = DM->qReiting->FieldByName("proiz_ball")->AsString;

   //КПЭ
   EditKPE->Text = zkpe = DM->qReiting->FieldByName("kpe")->AsString;
   EditKPE_BALL->Text = zkpe_ball = DM->qReiting->FieldByName("kpe_ball")->AsString;

   //наличие приемников
   EditPRIEM->Text = zpriem = DM->qReiting->FieldByName("priem")->AsString;
   EditPRIEM_BALL->Text = zpriem_ball = DM->qReiting->FieldByName("priem_ball")->AsString;

   //эффективное информирование
   EditINFO->Text = zinfo = DM->qReiting->FieldByName("info")->AsString;
   EditINFO_BALL->Text = zinfo_ball = DM->qReiting->FieldByName("info_ball")->AsString;

   //разбор отклонений
   EditOTKL->Text = zotkl = DM->qReiting->FieldByName("otkl")->AsString;
   EditOTKL_BALL->Text = zotkl_ball = DM->qReiting->FieldByName("otkl_ball")->AsString;

   //уровень 5С
   EditC5->Text = zc5 = DM->qReiting->FieldByName("c5")->AsString;
   EditC5_BALL->Text = zc5_ball = DM->qReiting->FieldByName("c5_ball")->AsString;

   //КНС
   EditKNS->Text = zkns = DM->qReiting->FieldByName("kns")->AsString;
   EditKNS_BALL->Text = zkns_ball = DM->qReiting->FieldByName("kns_ball")->AsString;

   //СПП
   EditSPP_KOL->Text = zspp_kol = DM->qReiting->FieldByName("spp_kol")->AsString;
   EditSPP_BALL->Text = zspp_ball = DM->qReiting->FieldByName("spp_ball")->AsString;

   EditOT_UPR->Text = zot_upr = DM->qReiting->FieldByName("ot_upr")->AsString;
   EditOT_TREB->Text = zot_treb = DM->qReiting->FieldByName("ot_treb")->AsString;
   EditTRUD_D->Text = ztrud_d = DM->qReiting->FieldByName("trud_d")->AsString;

   EditOCENKA->Text = zocenka = DM->qReiting->FieldByName("ocenka")->AsString;

   if (DM->qReiting->FieldByName("reit")->AsInteger==0) {
	  zreit = 0;
	  ComboBoxREIT->ItemIndex = -1;
	 // ComboBoxREIT->Color = clWhite;
   }
   else if (DM->qReiting->FieldByName("reit")->AsInteger==1) {
	  zreit = 1;
	  ComboBoxREIT->ItemIndex = 0;
	 // ComboBoxREIT->Color = clGreen;
   }
   else if (DM->qReiting->FieldByName("reit")->AsInteger==2) {
	  zreit = 2;
	  ComboBoxREIT->ItemIndex = 1;
	 // ComboBoxREIT->Color = clYellow;
   }
   else if (DM->qReiting->FieldByName("reit")->AsInteger==3) {
	  zreit = 3;
	  ComboBoxREIT->ItemIndex = 2;
	 // ComboBoxREIT->Color = clRed;
   }
}
//---------------------------------------------------------------------------

void __fastcall TVvod::EditZEXKeyPress(TObject *Sender, System::WideChar &Key)
{
  if (! (IsNumeric(Key) || Key=='.' || Key==',' || Key=='/' || Key=='\b') ) Key=0;
  if (Key==',' || Key=='/') Key='.';
}
//---------------------------------------------------------------------------

void __fastcall TVvod::EditTNKeyPress(TObject *Sender, System::WideChar &Key)
{
   if (!(IsNumeric(Key)||Key=='\b')) Key=0;
}
//---------------------------------------------------------------------------

void __fastcall TVvod::EditPROIZ_BALLExit(TObject *Sender)
{
  if (ActiveControl == Cansel)
	{
	  Vvod->Close();
	}
  else
	{
	  if (!EditPROIZ_BALL->Text.IsEmpty() && (EditPROIZ_BALL->Text>5 || EditPROIZ_BALL->Text<2)) {
		Application->MessageBoxW(L"Балл может быть только в пределах от 2 до 5",L"Предупреждение",
								 MB_ICONWARNING);
		EditPROIZ_BALL->SetFocus();
	  }
	}
}
//---------------------------------------------------------------------------

void __fastcall TVvod::EditKPE_BALLExit(TObject *Sender)
{
  if (ActiveControl == Cansel)
	{
	  Vvod->Close();
	}
  else
	{
	  if (!EditKPE_BALL->Text.IsEmpty() && (EditKPE_BALL->Text>5 || EditKPE_BALL->Text<2)) {
		Application->MessageBoxW(L"Балл может быть только в пределах от 2 до 5",L"Предупреждение",
								 MB_ICONWARNING);
		EditKPE_BALL->SetFocus();
	  }
	}
}
//---------------------------------------------------------------------------

void __fastcall TVvod::EditPRIEM_BALLExit(TObject *Sender)
{
  if (ActiveControl == Cansel)
	{
	  Vvod->Close();
	}
  else
	{
	  if (!EditPRIEM_BALL->Text.IsEmpty() && (EditPRIEM_BALL->Text>5 || EditPRIEM_BALL->Text<2)) {
		Application->MessageBoxW(L"Балл может быть только в пределах от 2 до 5",L"Предупреждение",
								 MB_ICONWARNING);
		EditPRIEM_BALL->SetFocus();
	  }
	}
}
//---------------------------------------------------------------------------

void __fastcall TVvod::EditINFO_BALLExit(TObject *Sender)
{
  if (ActiveControl == Cansel)
	{
	  Vvod->Close();
	}
  else
	{
	  if (!EditINFO_BALL->Text.IsEmpty() && (EditINFO_BALL->Text>5 || EditINFO_BALL->Text<2)) {
		Application->MessageBoxW(L"Балл может быть только в пределах от 2 до 5",L"Предупреждение",
								 MB_ICONWARNING);
		EditINFO_BALL->SetFocus();
	  }
	}
}
//---------------------------------------------------------------------------

void __fastcall TVvod::EditOTKL_BALLExit(TObject *Sender)
{
  if (ActiveControl == Cansel)
	{
	  Vvod->Close();
	}
  else
	{
	  if (!EditOTKL_BALL->Text.IsEmpty() && (EditOTKL_BALL->Text>5 || EditOTKL_BALL->Text<2)) {
		Application->MessageBoxW(L"Балл может быть только в пределах от 2 до 5",L"Предупреждение",
								 MB_ICONWARNING);
		EditOTKL_BALL->SetFocus();
	  }
	}
}
//---------------------------------------------------------------------------

void __fastcall TVvod::EditC5_BALLExit(TObject *Sender)
{
  if (ActiveControl == Cansel)
	{
	  Vvod->Close();
	}
  else
	{
	  if (!EditC5_BALL->Text.IsEmpty() && (EditC5_BALL->Text>5 || EditC5_BALL->Text<2)) {
		Application->MessageBoxW(L"Балл может быть только в пределах от 2 до 5",L"Предупреждение",
								 MB_ICONWARNING);
		EditC5_BALL->SetFocus();
	  }
	}
}
//---------------------------------------------------------------------------

void __fastcall TVvod::EditKNS_BALLExit(TObject *Sender)
{
  if (ActiveControl == Cansel)
	{
	  Vvod->Close();
	}
  else
	{
	  if (!EditKNS_BALL->Text.IsEmpty() && (EditKNS_BALL->Text>5 || EditKNS_BALL->Text<2)) {
		Application->MessageBoxW(L"Балл может быть только в пределах от 2 до 5",L"Предупреждение",
								 MB_ICONWARNING);
		EditKNS_BALL->SetFocus();
	  }
	}
}
//---------------------------------------------------------------------------

void __fastcall TVvod::EditSPP_BALLExit(TObject *Sender)
{
  if (ActiveControl == Cansel)
	{
	  Vvod->Close();
	}
  else
	{
	  if (!EditSPP_BALL->Text.IsEmpty() && (EditSPP_BALL->Text>5 || EditSPP_BALL->Text<2)) {
		Application->MessageBoxW(L"Балл может быть только в пределах от 2 до 5",L"Предупреждение",
								 MB_ICONWARNING);
		EditSPP_BALL->SetFocus();
	  }
	}
}
//---------------------------------------------------------------------------

void __fastcall TVvod::EditOT_UPRExit(TObject *Sender)
{
  if (ActiveControl == Cansel)
	{
	  Vvod->Close();
	}
  else
	{
	  if (!EditOT_UPR->Text.IsEmpty() && (EditOT_UPR->Text>5 || EditOT_UPR->Text<2)) {
		Application->MessageBoxW(L"Балл может быть только в пределах от 2 до 5",L"Предупреждение",
								 MB_ICONWARNING);
		EditOT_UPR->SetFocus();
	  }
	}
}
//---------------------------------------------------------------------------

void __fastcall TVvod::EditOT_TREBExit(TObject *Sender)
{
  if (ActiveControl == Cansel)
	{
	  Vvod->Close();
	}
  else
	{
	  if (!EditOT_TREB->Text.IsEmpty() && (EditOT_TREB->Text>5 || EditOT_TREB->Text<2)) {
		Application->MessageBoxW(L"Балл может быть только в пределах от 2 до 5",L"Предупреждение",
								 MB_ICONWARNING);
		EditOT_TREB->SetFocus();
	  }
	}

}
//---------------------------------------------------------------------------

void __fastcall TVvod::EditTRUD_DExit(TObject *Sender)
{
  if (ActiveControl == Cansel)
	{
	  Vvod->Close();
	}
  else
	{
	  if (!EditTRUD_D->Text.IsEmpty() && (EditTRUD_D->Text>5 || EditTRUD_D->Text<2)) {
		Application->MessageBoxW(L"Балл может быть только в пределах от 2 до 5",L"Предупреждение",
								 MB_ICONWARNING);
		EditTRUD_D->SetFocus();
	  }
	}
}
//---------------------------------------------------------------------------

void __fastcall TVvod::SaveClick(TObject *Sender)
{
  AnsiString Sql, podch, reit;
  int rec;

  //Наличие подчиненных
  if (CheckBoxPODCH->Checked == true) podch = 1;
  else podch = 0;

  //Рейтинг
  if (ComboBoxREIT->ItemIndex == -1) reit = 0;
   else if (ComboBoxREIT->ItemIndex == 0) reit = 1;
   else if (ComboBoxREIT->ItemIndex == 1) reit = 2;
   else if (ComboBoxREIT->ItemIndex == 2) reit = 3;

  //Проверки
  //цех
  if (EditZEX->Text.IsEmpty()) {
	  Application->MessageBoxW(L"Не указан цех работника",L"Предупреждение",
							   MB_ICONWARNING);
	  EditZEX->SetFocus();
	  Abort();
  }

  //тн
	if (EditTN->Text.IsEmpty()) {
	  Application->MessageBoxW(L"Не указан таб.№ работника",L"Предупреждение",
							   MB_ICONWARNING);
	  EditTN->SetFocus();
	  Abort();
  }

  //фио
	if (LabelFIO->Caption.IsEmpty()) {
	  Application->MessageBoxW(L"Не указана ФИО работника",L"Предупреждение",
							   MB_ICONWARNING);
	  Abort();
  }

  //шифр должности
	if (EditID_DOLG->Text.IsEmpty()) {
	  Application->MessageBoxW(L"Не указан шифр должности работника",L"Предупреждение",
							   MB_ICONWARNING);
	  EditID_DOLG->SetFocus();
	  Abort();
  }

  //должность
	if (EditDOLG->Text.IsEmpty()) {
	  Application->MessageBoxW(L"Не указана должность работника",L"Предупреждение",
							   MB_ICONWARNING);
	  EditDOLG->SetFocus();
	  Abort();
  }

  //Расчет оценки
  if (EditOCENKA->Text.IsEmpty())
	{
	  Ocenka();
	}


  //Добавление записи
  if (Main->redakt==0) {

	 //Проверка на наличие записи по работнику по тн
	 Sql ="select * from reit_ruk \
		   where tn="+EditTN->Text+" and god="+IntToStr(Main->god) +" and kvart="+IntToStr(Main->kvartal);

	 DM->qObnovlenie->Close();
	 DM->qObnovlenie->SQL->Clear();
	 DM->qObnovlenie->SQL->Add(Sql);

	 try
	   {
		 DM->qObnovlenie->Open();
	   }
	 catch(...)
	   {
		 Application->MessageBox(L"Ошибка доступа к таблице SAP_OSN_SVED",
								 L"Ошибка доступа",MB_OK + MB_ICONERROR);
		 Abort();
	   }

	 if (DM->qObnovlenie->RecordCount!=0)
	   {
		 Application->MessageBox(L"Запись по работнику с указанным таб.№ за текущий квартал уже существует в таблице.\nПроверьте правильность ввода табельного номера",
								 L"Предупреждение",MB_OK+MB_ICONINFORMATION);
		 EditTN->SetFocus();
		 Abort();
	   }



     //Сохранение данных
	 Sql ="insert into reit_ruk (god, kvart, zex, tn, fio, id_dolg, dolg, uch, podch, \
								 kpe, kpe_ball, proiz, proiz_ball, otkl, otkl_ball,   \
								 info, info_ball, c5, c5_ball, kns, kns_ball, priem,  \
								 priem_ball, spp_kol, spp_ball, ot_upr, ot_treb, ocenka, reit, trud_d) \
		   values ( \
					"+ IntToStr(Main->god) +",\
					"+ IntToStr(Main->kvartal)+",\
					"+ QuotedStr(EditZEX->Text) +",\
					"+ EditTN->Text +",\
					"+ QuotedStr(LabelFIO->Caption) +",\
					"+ QuotedStr(EditID_DOLG->Text) +",\
					"+ QuotedStr(EditDOLG->Text) +",\
					  (select name_ur1 from sap_osn_sved where tn_sap="+EditTN->Text+"), \
					"+podch+", \
					"+ SetNull(EditKPE->Text) +",\
					"+ SetNull(EditKPE_BALL->Text) +",\
					"+ SetNull(EditPROIZ->Text) +",\
					"+ SetNull(EditPROIZ_BALL->Text) +",\
					"+ SetNull(EditOTKL->Text) +",\
					"+ SetNull(EditOTKL_BALL->Text) +",\
					"+ SetNull(EditINFO->Text) +",\
					"+ SetNull(EditINFO_BALL->Text) +",\
					"+ SetNull(EditC5->Text) +",\
					"+ SetNull(EditC5_BALL->Text) +",\
					"+ SetNull(EditKNS->Text) +",\
					"+ SetNull(EditKNS_BALL->Text) +",\
					"+ SetNull(EditPRIEM->Text) +",\
					"+ SetNull(EditPRIEM_BALL->Text) +",\
					"+ SetNull(EditSPP_KOL->Text) +",\
					"+ SetNull(EditSPP_BALL->Text) +",\
					"+ SetNull(EditOT_UPR->Text) +",\
					"+ SetNull(EditOT_TREB->Text) +",\
					"+ SetNull(EditOCENKA->Text) +",\
					"+ SetNull(reit) +",\
					"+ SetNull(EditTRUD_D->Text) +")";


	 DM->qObnovlenie->Close();
	 DM->qObnovlenie->SQL->Clear();
	 DM->qObnovlenie->SQL->Add(Sql);

	 rec = StrToInt(EditTN->Text);
	 try
	   {
		 DM->qObnovlenie->ExecSQL();
	   }
	 catch(...)
	   {
		 Application->MessageBox(L"Возникла ошибка при добавлении данных",
								 L"Ошибка добавления новой записи",MB_OK + MB_ICONERROR);
		 Main->InsertLog("Возникла ошибка при добавлении записи за "+IntToStr(Main->god)+" год, "+IntToStr(Main->kvartal)+" квартал по работнику: цех="+EditZEX->Text+" таб.№="+EditTN->Text);
		 Abort();
	   }

	/* if (!EditOCENKA->Text.IsEmpty())
	   {
		 //Пересчитать рейтинг
		 Main->RaschetReit(1, EditZEX->Text, StrToInt(podch));
	   }  */

	 DM->qReiting->Refresh();

	 Main->InsertLog("Выполнено добавление записи: цех ="+ EditZEX->Text +", таб.№ ="+ EditTN->Text +", ФИО = "+LabelFIO->Caption);

	 TLocateOptions SearchOptions;
	 DM->qReiting->Locate("tn",rec,SearchOptions<<loPartialKey<<loCaseInsensitive);
  }

  //Редактирование записи
  else {
	 //Проверка были ли вообще изменения
	 if (zzex!=EditZEX->Text ||
		 ztn!=EditTN->Text ||
		 zid_dolg!=EditID_DOLG->Text ||
		 zdolg!=EditDOLG->Text ||
		 zpodch!=podch ||
		 zproiz!=EditPROIZ->Text ||
		 zproiz_ball!=EditPROIZ_BALL->Text ||
		 zkpe!=EditKPE->Text ||
		 zkpe_ball!=EditKPE_BALL->Text ||
		 zpriem!=EditPRIEM->Text ||
		 zpriem_ball!=EditPRIEM_BALL->Text ||
		 zinfo!=EditINFO->Text ||
		 zinfo_ball!=EditINFO_BALL->Text ||
		 zotkl!=EditOTKL->Text ||
		 zotkl_ball!=EditOTKL_BALL->Text ||
		 zc5!=EditC5->Text ||
		 zc5_ball!=EditC5_BALL->Text ||
		 zkns!=EditKNS->Text ||
		 zkns_ball!=EditKNS_BALL->Text ||
		 zspp_kol!=EditSPP_KOL->Text ||
		 zspp_ball!=EditSPP_BALL->Text ||
		 zot_upr!=EditOT_UPR->Text ||
		 zot_treb!=EditOT_TREB->Text ||
		 ztrud_d!=EditTRUD_D->Text ||
		 zocenka!=EditOCENKA->Text ||
		 zreit!=reit
		 )
	   {
		  //Обновление записи
		  if (Main->Prava=="ocen")
			{
			  Sql = "update reit_ruk set \
					   zex = "+QuotedStr(EditZEX->Text)+",\
					   tn= "+EditTN->Text+",\
					   id_dolg = "+ QuotedStr(EditID_DOLG->Text)+",\
					   dolg = "+ QuotedStr(EditDOLG->Text)+",\
					   uch = (select name_ur1 from sap_osn_sved where tn_sap="+EditTN->Text+"), \
					   podch = "+ podch+",\
					   kpe = "+ SetNull(EditKPE->Text)+",\
					   kpe_ball = "+ SetNull(EditKPE_BALL->Text)+",\
					   proiz = "+ SetNull(EditPROIZ->Text)+",\
					   proiz_ball = "+ SetNull(EditPROIZ_BALL->Text)+",\
					   otkl = "+ SetNull(EditOTKL->Text)+",\
					   otkl_ball = "+ SetNull(EditOTKL_BALL->Text)+",\
					   info = "+ SetNull(EditINFO->Text)+",\
					   info_ball = "+ SetNull(EditINFO_BALL->Text)+",\
					   c5 = "+ SetNull(EditC5->Text)+",\
					   c5_ball = "+ SetNull(EditC5_BALL->Text)+",\
					   kns = "+ SetNull(EditKNS->Text)+",\
					   kns_ball = "+ SetNull(EditKNS_BALL->Text)+",\
					   priem  = "+ SetNull(EditPRIEM->Text)+",\
					   priem_ball = "+ SetNull(EditPRIEM_BALL->Text)+",\
					   spp_kol = "+ SetNull(EditSPP_KOL->Text)+",\
					   spp_ball = "+ SetNull(EditSPP_BALL->Text)+",\
					   ot_upr = "+ SetNull(EditOT_UPR->Text)+",\
					   ot_treb = "+ SetNull(EditOT_TREB->Text)+",\
					   trud_d = "+ SetNull(EditTRUD_D->Text)+",\
					   ocenka = "+ SetNull(EditOCENKA->Text)+",\
					   reit = "+ SetNull(reit)+"\
				 where rowid = chartorowid("+ QuotedStr(DM->qReiting->FieldByName("rw")->AsString)+")";
			}
		  else if (Main->Prava=="unou")
			{
			  Sql = "update reit_ruk set \
					   zex = "+QuotedStr(EditZEX->Text)+",\
					   tn= "+EditTN->Text+",\
					   id_dolg = "+ QuotedStr(EditID_DOLG->Text)+",\
					   dolg = "+ QuotedStr(EditDOLG->Text)+",\
					   uch = (select name_ur1 from sap_osn_sved where tn_sap="+EditTN->Text+"), \
					   podch = "+ podch+",\
					   otkl = "+ SetNull(EditOTKL->Text)+",\
					   otkl_ball = "+ SetNull(EditOTKL_BALL->Text)+",\
					   info = "+ SetNull(EditINFO->Text)+",\
					   info_ball = "+ SetNull(EditINFO_BALL->Text)+",\
					   c5 = "+ SetNull(EditC5->Text)+",\
					   c5_ball = "+ SetNull(EditC5_BALL->Text)+",\
					   kns = "+ SetNull(EditKNS->Text)+",\
					   kns_ball = "+ SetNull(EditKNS_BALL->Text)+"\
				 where rowid = chartorowid("+ QuotedStr(DM->qReiting->FieldByName("rw")->AsString)+")";
			}
		  else if (Main->Prava=="pp")
			{
              Sql = "update reit_ruk set \
					   proiz = "+ SetNull(EditPROIZ->Text)+",\
					   proiz_ball = "+ SetNull(EditPROIZ_BALL->Text)+"\
				 where rowid = chartorowid("+ QuotedStr(DM->qReiting->FieldByName("rw")->AsString)+")";
			}
		  else if (Main->Prava=="kpe")
			{
			  Sql = "update reit_ruk set \
					   kpe = "+ SetNull(EditKPE->Text)+",\
					   kpe_ball = "+ SetNull(EditKPE_BALL->Text)+"\
				 where rowid = chartorowid("+ QuotedStr(DM->qReiting->FieldByName("rw")->AsString)+")";
			}
		  else if (Main->Prava=="spp")
			{
			  Sql = "update reit_ruk set \
					   spp_kol = "+ SetNull(EditSPP_KOL->Text)+",\
					   spp_ball = "+ SetNull(EditSPP_BALL->Text)+"\
				 where rowid = chartorowid("+ QuotedStr(DM->qReiting->FieldByName("rw")->AsString)+")";
			}
		  else if (Main->Prava=="ot")
			{
			  Sql = "update reit_ruk set \
					   ot_upr = "+ SetNull(EditOT_UPR->Text)+",\
					   ot_treb = "+ SetNull(EditOT_TREB->Text)+"\
				 where rowid = chartorowid("+ QuotedStr(DM->qReiting->FieldByName("rw")->AsString)+")";
			}
		  else if (Main->Prava=="td")
			{
               Sql = "update reit_ruk set \
					   trud_d = "+ SetNull(EditTRUD_D->Text)+"\
				 where rowid = chartorowid("+ QuotedStr(DM->qReiting->FieldByName("rw")->AsString)+")";
			}

		  rec = StrToInt(EditTN->Text);
		  DM->qObnovlenie->Close();
		  DM->qObnovlenie->SQL->Clear();
		  DM->qObnovlenie->SQL->Add(Sql);
		  try
			{
			  DM->qObnovlenie->ExecSQL();
			}
		  catch(Exception &E)
			{
			  Application->MessageBox(("Возникла ошибка при получении данных из таблицы REIT_RUK" + E.Message).c_str(),L"Ошибка",
										MB_OK+MB_ICONERROR);
			  Main->InsertLog("Возникла ошибка при обновлении данных за "+IntToStr(Main->god)+" год, "+IntToStr(Main->kvartal)+" квартал по работнику: цех="+EditZEX->Text+" таб.№="+EditTN->Text);
			  Abort();
			}


		  //Логи
		if (DM->qObnovlenie->RowsAffected>0)
			{
			  String Str ="Обновление записи за "+IntToStr(Main->god)+" год, "+IntToStr(Main->kvartal)+" квартал: ";

			  if (zzex!=EditZEX->Text || ztn!=EditTN->Text)  Str+="цех с '"+zzex+"' на '"+EditZEX->Text+"' таб.№ с '"+ztn+"' на '"+EditTN->Text+"'";
			  else Str+="цех='"+zzex+"' таб.№ ='"+ztn+"'";

			  if (zid_dolg!=EditID_DOLG->Text) Str+=", шифр долж. с '"+zid_dolg+"' на '"+EditID_DOLG->Text+"'";
			  if (zdolg!=EditDOLG->Text) Str+=", должность с '"+zdolg+"' на '"+EditDOLG->Text+"'";
			  if (zpodch!=podch) Str+=", налич. подчиненных с '"+zpodch+"' на '"+podch+"'";
			  if (zproiz!=EditPROIZ->Text) Str+=", произв. задан. с '"+zproiz+"' на '"+EditPROIZ->Text+"'";
			  if (zproiz_ball!=EditPROIZ_BALL->Text) Str+=", балл произв. задан. с '"+zproiz_ball+"' на '"+EditPROIZ_BALL->Text+"'";
			  if (zkpe!=EditKPE->Text) Str+=", КПЭ с '"+zkpe+"' на '"+EditKPE->Text+"'";
			  if (zkpe_ball!=EditKPE_BALL->Text) Str+=", балл КПЭ с '"+zkpe_ball+"' на '"+EditKPE_BALL->Text+"'";
			  if (zpriem!=EditPRIEM->Text) Str+=", налич. преем. с '"+zpriem+"' на '"+EditPRIEM->Text+"'";
			  if (zpriem_ball!=EditPRIEM_BALL->Text) Str+=", балл преем. с '"+zpriem_ball+"' на '"+EditPRIEM_BALL->Text+"'";
			  if (zinfo!=EditINFO->Text) Str+=", эф. информ. с '"+zinfo+"' на '"+EditINFO->Text+"'";
			  if (zinfo_ball!=EditINFO_BALL->Text) Str+=", балл эф. информ. с '"+zinfo_ball+"' на '"+EditINFO_BALL->Text+"'";
			  if (zotkl!=EditOTKL->Text) Str+=", отклон. с '"+zotkl+"' на '"+EditOTKL->Text+"'";
			  if (zotkl_ball!=EditOTKL_BALL->Text) Str+=", балл отклон. с '"+zotkl_ball+"' на '"+EditOTKL_BALL->Text+"'";
			  if (zc5!=EditC5->Text) Str+=", 5С с '"+zc5+"' на '"+EditC5->Text+"'";
			  if (zc5_ball!=EditC5_BALL->Text) Str+=", балл 5С с '"+zc5_ball+"' на '"+EditC5_BALL->Text+"'";
			  if (zkns!=EditKNS->Text) Str+=", КНС с '"+zkns+"' на '"+EditKNS->Text+"'";
			  if (zkns_ball!=EditKNS_BALL->Text) Str+=", балл КНС с '"+zkns_ball+"' на '"+EditKNS_BALL->Text+"'";
			  if (zspp_kol!=EditSPP_KOL->Text) Str+=", кол. СПП с '"+zspp_kol+"' на '"+EditSPP_KOL->Text+"'";
			  if (zspp_ball!=EditSPP_BALL->Text) Str+=", балл СПП с '"+zspp_ball+"' на '"+EditSPP_BALL->Text+"'";
			  if (zot_upr!=EditOT_UPR->Text) Str+=", ОТ,ПБ и Э с '"+zot_upr+"' на '"+EditOT_UPR->Text+"'";
			  if (zot_treb!=EditOT_TREB->Text) Str+=", Наруш. ОТ с '"+zot_treb+"' на '"+EditOT_TREB->Text+"'";
			  if (ztrud_d!=EditTRUD_D->Text) Str+=", ТД с '"+ztrud_d+"' на '"+EditTRUD_D->Text+"'";
			  if (zocenka!=EditOCENKA->Text) Str+=", Оценка с '"+zocenka+"' на '"+EditOCENKA->Text+"'";
			  if (zreit!=reit) Str+=", Оценка с '"+zreit+"' на '"+reit+"'";


			  Main->InsertLog(Str);
			}
		  else
			{
			  Main->InsertLog("Обновление данных за "+IntToStr(Main->god)+" год, "+IntToStr(Main->kvartal)+" квартал по работнику: цех="+EditZEX->Text+" таб.№="+EditTN->Text+" не выполнено");
			}


		  /*if (EditOCENKA->Text!=zocenka)
			{
			  //Пересчитать рейтинг
			  Main->RaschetReit(1, EditZEX->Text, StrToInt(podch));
			} */


		  DM->qReiting->Refresh();
		  //Возвращение на выбранную строку
		  TLocateOptions SearchOptions;
		  DM->qReiting->Locate("tn",rec,SearchOptions<<loPartialKey<<loCaseInsensitive);

		  Application->MessageBox(L"Запись успешно изменена",L"Предупреждение",
								   MB_OK+MB_ICONINFORMATION);

	   }
	}
  Vvod->Close();
}
//---------------------------------------------------------------------------

AnsiString  __fastcall TVvod::SetNull (AnsiString str, AnsiString r)
{
  if (str.Length()) return str;
  else return r;
}
//---------------------------------------------------------------------------
//---------------------------------------------------------------------------

float  __fastcall TVvod::SetN(String str, float r)
{
  if (str.Length()) return StrToFloat(str);
  else return r;
}
//---------------------------------------------------------------------------


void __fastcall TVvod::EditTNChange(TObject *Sender)
{
  //Заполнение Edit-ов данными по введенному таб№ при добавлении новой записи
  AnsiString Sql;

  if (!EditTN->Text.IsEmpty() && Main->redakt==0) {

	 Sql ="select zex,                          \
				  fam||' '||im||' '||ot as fio, \
				  nzex,                         \
				  id_shtat,                     \
				  name_dolg_ru                 \
		   from sap_osn_sved \                  \
		   where tn_sap="+EditTN->Text;

	 DM->qObnovlenie->Close();
	 DM->qObnovlenie->SQL->Clear();
	 DM->qObnovlenie->SQL->Add(Sql);

	 try
	   {
		 DM->qObnovlenie->Open();
	   }
	 catch(...)
	   {
		 Application->MessageBox(L"Ошибка доступа к таблице SAP_OSN_SVED",
								 L"Ошибка доступа",MB_OK + MB_ICONERROR);
		 Abort();
	   }

   LabelFIO->Caption = DM->qObnovlenie->FieldByName("fio")->AsString;
   EditZEX->Text = DM->qObnovlenie->FieldByName("zex")->AsString;
   LabelNZEX->Caption = DM->qObnovlenie->FieldByName("nzex")->AsString;
   EditID_DOLG->Text = DM->qObnovlenie->FieldByName("id_shtat")->AsString;
   EditDOLG->Text = DM->qObnovlenie->FieldByName("name_dolg_ru")->AsString;

  }
}
//---------------------------------------------------------------------------
void __fastcall TVvod::EditPROIZExit(TObject *Sender)
{
 if (ActiveControl == Cansel)
	{
	  Vvod->Close();
	}
  else
	{
	  if (!EditPROIZ->Text.IsEmpty() && StrToFloat(EditPROIZ->Text)>200)
		{
		  Application->MessageBox(L"Максимальное значение по производственному заданию \nне может превышать 200%",L"Предупреждение",
								  MB_OK+MB_ICONWARNING);
		  EditPROIZ->SetFocus();
		}
	}
}
//---------------------------------------------------------------------------

void __fastcall TVvod::EditKPEExit(TObject *Sender)
{
 if (ActiveControl == Cansel)
	{
	  Vvod->Close();
	}
  else
	{
	  if (!EditKPE->Text.IsEmpty() && StrToFloat(EditKPE->Text)>100)
		{
		  Application->MessageBox(L"Максимальное значение по КПЭ не может превышать 100%",L"Предупреждение",
								  MB_OK+MB_ICONWARNING);
		  EditKPE->SetFocus();
		}
	}
}
//---------------------------------------------------------------------------

void __fastcall TVvod::EditOTKLExit(TObject *Sender)
{
 if (ActiveControl == Cansel)
	{
	  Vvod->Close();
	}
  else
	{
	  if (!EditOTKL->Text.IsEmpty() && StrToFloat(EditOTKL->Text)>100)
		{
		  Application->MessageBox(L"Максимальное значение по разбору отклонений \nне может превышать 100%",L"Предупреждение",
								  MB_OK+MB_ICONWARNING);
		  EditOTKL->SetFocus();
		}
	}
}
//---------------------------------------------------------------------------

void __fastcall TVvod::EditPRIEMExit(TObject *Sender)
{
 if (ActiveControl == Cansel)
	{
	  Vvod->Close();
	}
  else
	{
	  if (!EditPRIEM->Text.IsEmpty() && StrToFloat(EditPRIEM->Text)>10)
		{
		  Application->MessageBox(L"Максимальное значение по наличию приемников \nне может превышать 10",L"Предупреждение",
								  MB_OK+MB_ICONWARNING);
		  EditPRIEM->SetFocus();
		}
	}
}
//---------------------------------------------------------------------------

void __fastcall TVvod::EditINFOExit(TObject *Sender)
{
  if (ActiveControl == Cansel)
	{
	  Vvod->Close();
	}
  else
	{
	  if (!EditINFO->Text.IsEmpty() && StrToFloat(EditINFO->Text)>32)
		{
		  Application->MessageBox(L"Максимальное значение уровня по эффективному информированию\nне может превышать 32",L"Предупреждение",
								  MB_OK+MB_ICONWARNING);
		  EditINFO->SetFocus();
		}
	}
}
//---------------------------------------------------------------------------

void __fastcall TVvod::EditC5Exit(TObject *Sender)
{
 if (ActiveControl == Cansel)
	{
	  Vvod->Close();
	}
  else
	{
	  if (!EditC5->Text.IsEmpty() && StrToFloat(EditC5->Text)>12)
		{
		  Application->MessageBox(L"Максимальное значение по уровню 5С\nне может превышать 12",L"Предупреждение",
								  MB_OK+MB_ICONWARNING);
		  EditC5->SetFocus();
		}
	}
}
//---------------------------------------------------------------------------

void __fastcall TVvod::EditKNSExit(TObject *Sender)
{
 if (ActiveControl == Cansel)
	{
	  Vvod->Close();
	}
  else
	{
	  if (!EditKNS->Text.IsEmpty() && StrToFloat(EditKNS->Text)>100)
		{
		  Application->MessageBox(L"Максимальное значение по КНС\nне может превышать 100",L"Предупреждение",
								  MB_OK+MB_ICONWARNING);
		  EditKNS->SetFocus();
		}
	}
}
//---------------------------------------------------------------------------

void __fastcall TVvod::EditSPP_KOLExit(TObject *Sender)
{
 if (ActiveControl == Cansel)
	{
	  Vvod->Close();
	}
  else
	{
	  if (!EditSPP_KOL->Text.IsEmpty() && StrToFloat(EditSPP_KOL->Text)>2)
		{
		  Application->MessageBox(L"Максимальное значение по СПП\nне может превышать 2",L"Предупреждение",
								  MB_OK+MB_ICONWARNING);
		  EditSPP_KOL->SetFocus();
		}
	}
}
//---------------------------------------------------------------------------
void __fastcall TVvod::FormCreate(TObject *Sender)
{
  //Видимость полей на форме редактирования в зависимости от группы доступа
  //если права группы Оценки персонала
  if (Main->Prava=="ocen")
	{
	  //Результативность
	  //Произ. показатель и КПЭ
	  BevelREZULT->Visible = true;
	  LabelREZULT->Visible = true;
	  GroupBoxREZULT_PP->Visible = true;
	  GroupBoxREZULT_KPE->Visible = true;

	  BevelREZULT->Top = 162;
	  BevelREZULT->Left = 15;
	  BevelREZULT->Width = 291;

	  LabelREZULT->Top = 165;
	  LabelREZULT->Left = 24;

	  GroupBoxREZULT_PP->Top =198;
	  GroupBoxREZULT_PP->Left = 24;

	  GroupBoxREZULT_KPE->Top = 198;
	  GroupBoxREZULT_KPE->Left = 161;


	  //Развитие персонала
	  //Преемники и информирование
	  //************************************************************************
	  BevelRPERSONAL->Visible = true;
	  LabelRPERSONAL->Visible = true;
	  GroupBoxRPERSONAL_PREEM->Visible = true;
	  GroupBoxRPERSONAL_INFO->Visible = true;

	  BevelRPERSONAL->Top = 162;
	  BevelRPERSONAL->Left = 310;
	  BevelRPERSONAL->Width = 354;

	  LabelRPERSONAL->Top = 165;
	  LabelRPERSONAL->Left = 322;

	  GroupBoxRPERSONAL_PREEM->Top =198;
	  GroupBoxRPERSONAL_PREEM->Left = 322;

	  GroupBoxRPERSONAL_INFO->Top = 198;
	  GroupBoxRPERSONAL_INFO->Left = 488;


	  //Стандартизация
	  //Отклонения
	  //************************************************************************
	  BevelSTAND->Visible = true;
	  LabelSTAND->Visible = true;
	  GroupBoxSTAND->Visible = true;

	  //Вовлечение персонала
	  //5С, КНС и СПП
	  //************************************************************************
	  BevelVPERSONAL->Visible = true;
	  LabelVPERSONAL->Visible = true;
	  GroupBoxVPERSONAL_5C->Visible = true;
	  GroupBoxVPERSONAL_KNS->Visible = true;
	  GroupBoxVPERSONAL_SPP->Visible = true;

	  BevelVPERSONAL->Top = 309;
	  BevelVPERSONAL->Left = 200;
	  BevelVPERSONAL->Width = 464;

	  LabelVPERSONAL->Top = 317;
	  LabelVPERSONAL->Left = 209;

	  GroupBoxVPERSONAL_5C->Top =352;
	  GroupBoxVPERSONAL_5C->Left = 206;

	  GroupBoxVPERSONAL_KNS->Top = 352;
	  GroupBoxVPERSONAL_KNS->Left = 353;

	  GroupBoxVPERSONAL_SPP->Top = 352;
	  GroupBoxVPERSONAL_SPP->Left = 521;

	  //ОТ, ТД, Оценка
	  //************************************************************************
	  BevelOT->Visible = true;
	  GroupBoxTD->Visible = true;
	  GroupBoxOT->Visible = true;
	  GroupBoxOCENKA->Visible = true;

	  BevelOT->Top = 467;
	  BevelOT->Left = 15;
	  BevelOT->Width = 649;
	  BevelOT->Height = 82;

	  GroupBoxOT->Top =481;
	  GroupBoxOT->Left = 24;

	  GroupBoxTD->Top = 481;
	  GroupBoxTD->Left = 353;

	  GroupBoxOCENKA->Top = 481;
	  GroupBoxOCENKA->Left = 476;


	  BevelImage->Height = 517;
	  Image1->Height= 244;

	  Save->Top = 453;
	  Cansel->Top = 493;

	  Vvod->Height = 604;

	  EditZEX->Enabled = true;
	  EditTN->Enabled = true;
	  EditID_DOLG->Enabled = true;
	  EditDOLG->Enabled = true;
	  CheckBoxPODCH->Enabled = true;

	}
  else
	{

	  //Результативность
	  //Произ. показатель и КПЭ
	  BevelREZULT->Visible = false;
	  LabelREZULT->Visible = false;
	  GroupBoxREZULT_PP->Visible = false;
	  GroupBoxREZULT_KPE->Visible = false;


	  //Развитие персонала
	  //Преемники и информирование
	  BevelRPERSONAL->Visible = false;
	  LabelRPERSONAL->Visible = false;
	  GroupBoxRPERSONAL_PREEM->Visible = false;
	  GroupBoxRPERSONAL_INFO->Visible = false;

	  //Стандартизация
	  //Отклонения
	  BevelSTAND->Visible = false;
	  LabelSTAND->Visible = false;
	  GroupBoxSTAND->Visible = false;

	  //Вовлечение персонала
	  //5С, КНС и СПП
	  BevelVPERSONAL->Visible = false;
	  LabelVPERSONAL->Visible = false;
	  GroupBoxVPERSONAL_5C->Visible = false;
	  GroupBoxVPERSONAL_KNS->Visible = false;
	  GroupBoxVPERSONAL_SPP->Visible = false;

	  //ОТ
	  BevelOT->Visible = false;
	  GroupBoxOT->Visible = false;

	  //ТД
	  GroupBoxTD->Visible = false;

	  //Оценка
	  GroupBoxOCENKA->Visible = false;

	  Image1->Height = 163;

	  Save->Top = 214;
	  Cansel->Top = 254;

	  BevelImage->Height = 274;
	  Vvod->Height = 348;

	  EditZEX->Enabled = false;
	  EditTN->Enabled = false;
	  EditID_DOLG->Enabled = false;
	  EditDOLG->Enabled = false;
	  CheckBoxPODCH->Enabled = false;

	  //если права группы УНОУ
	  if (Main->Prava=="unou")
		{
		  BevelRPERSONAL->Visible = true;
		  LabelRPERSONAL->Visible = true;
		  GroupBoxRPERSONAL_INFO->Visible = true;

		  BevelRPERSONAL->Top = 162;
		  BevelRPERSONAL->Left = 15;
		  BevelRPERSONAL->Width = 649;

		  LabelRPERSONAL->Top = 165;
		  LabelRPERSONAL->Left = 24;

		  GroupBoxRPERSONAL_INFO->Top = 198;
		  GroupBoxRPERSONAL_INFO->Left = 24;


		  BevelSTAND->Visible = true;
		  LabelSTAND->Visible = true;
		  GroupBoxSTAND->Visible = true;

		  BevelVPERSONAL->Visible = true;
		  LabelVPERSONAL->Visible = true;
		  GroupBoxVPERSONAL_5C->Visible = true;
		  GroupBoxVPERSONAL_KNS->Visible = true;

		  BevelVPERSONAL->Top = 309;
		  BevelVPERSONAL->Left = 200;
		  BevelVPERSONAL->Width = 464;

		  LabelVPERSONAL->Top = 317;
		  LabelVPERSONAL->Left = 209;

		  GroupBoxVPERSONAL_5C->Top =352;
		  GroupBoxVPERSONAL_5C->Left = 206;

		  GroupBoxVPERSONAL_KNS->Top = 352;
		  GroupBoxVPERSONAL_KNS->Left = 353;

		  GroupBoxVPERSONAL_SPP->Top = 352;
		  GroupBoxVPERSONAL_SPP->Left = 521;

		  BevelImage->Height = 431;
		  Image1->Height= 244;

		  Save->Top = 373;
		  Cansel->Top = 413;

		  Vvod->Height = 505;

		  EditZEX->Enabled = true;
		  EditTN->Enabled = true;
		  EditID_DOLG->Enabled = true;
		  EditDOLG->Enabled = true;
		  CheckBoxPODCH->Enabled = true;
		}

	  //если права группы загрузки производственного задания
	  else if (Main->Prava=="pp")
		{
		  //Результативность
		  //Произ. показатель и КПЭ
		  BevelREZULT->Visible = true;
		  LabelREZULT->Visible = true;
		  GroupBoxREZULT_PP->Visible = true;

		  BevelREZULT->Top = 162;
		  BevelREZULT->Left = 15;
		  BevelREZULT->Width = 649;

		  LabelREZULT->Top = 165;
		  LabelREZULT->Left = 24;

		  GroupBoxREZULT_PP->Top =198;
		  GroupBoxREZULT_PP->Left = 24;

		  GroupBoxREZULT_KPE->Top = 198;
		  GroupBoxREZULT_KPE->Left = 161;
		}

	  //если права группы загрузки КПЭ
	  else if (Main->Prava=="kpe")
		{
		  //Результативность
		  //КПЭ
		  BevelREZULT->Visible = true;
		  LabelREZULT->Visible = true;
		  GroupBoxREZULT_KPE->Visible = true;

		  BevelREZULT->Top = 162;
		  BevelREZULT->Left = 15;
		  BevelREZULT->Width = 649;

		  LabelREZULT->Top = 165;
		  LabelREZULT->Left = 24;

		  GroupBoxREZULT_KPE->Top = 198;
		  GroupBoxREZULT_KPE->Left = 24;
		}

	  //если права группы загрузки СПП
	  else if (Main->Prava=="spp")
		{
		  BevelVPERSONAL->Visible = true;
		  LabelVPERSONAL->Visible = true;
		  GroupBoxVPERSONAL_SPP->Visible = true;

		  BevelVPERSONAL->Top = 162;
		  BevelVPERSONAL->Left = 15;
		  BevelVPERSONAL->Width = 649;

		  LabelVPERSONAL->Top = 165;
		  LabelVPERSONAL->Left = 24;

		  GroupBoxVPERSONAL_SPP->Top = 198;
		  GroupBoxVPERSONAL_SPP->Left = 24;
		}

	  //если права группы загрузки ОТ
	  else if (Main->Prava=="ot")
		{
		  BevelOT->Visible = true;
		  GroupBoxOT->Visible = true;

		  BevelOT->Top = 162;
		  BevelOT->Left = 15;
		  BevelOT->Width = 649;
		  BevelOT->Height = 144;

		  GroupBoxOT->Top =178;
		  GroupBoxOT->Left = 24;
		}
	  //если права группы загрузки трудовой дисциплины
	  else if (Main->Prava=="td")
		{
		  BevelOT->Visible = true;
		  GroupBoxTD->Visible = true;

		  BevelOT->Top = 162;
		  BevelOT->Left = 15;
		  BevelOT->Width = 649;
		  BevelOT->Height = 144;

		  GroupBoxTD->Top = 178;
		  GroupBoxTD->Left = 24;
		}
	}
}
//---------------------------------------------------------------------------

//---------------------------------------------------------------------------
//Расчет оценки при редактировании записи
void __fastcall TVvod::Ocenka()
{
  if (DM->qReiting->FieldByName("pz")->AsInteger==1)
	{
	 EditOCENKA->Text = (SetN(EditPROIZ_BALL->Text)*0.1 +
						  SetN(EditKPE_BALL->Text)*0.15 +
						  SetN(EditOTKL_BALL->Text)*0.15 +
						  SetN(EditPRIEM_BALL->Text)*0.1 +
						  SetN(EditINFO_BALL->Text)*0.1 +
						  SetN(EditKNS_BALL->Text)*0.1 +
						  SetN(EditC5_BALL->Text)*0.1 +
						  SetN(EditSPP_BALL->Text)*0.1 +
						  SetN(EditOT_UPR->Text)*0.1 ) -
						  SetN(EditOT_TREB->Text)-SetN(EditTRUD_D->Text);
	}
  else
	{
	 EditOCENKA->Text = (SetN(EditKPE_BALL->Text)*0.2 +
						 SetN(EditOTKL_BALL->Text)*0.2 +
						 SetN(EditPRIEM_BALL->Text)*0.1 +
						 SetN(EditINFO_BALL->Text)*0.1 +
						 SetN(EditKNS_BALL->Text)*0.1 +
						 SetN(EditC5_BALL->Text)*0.1 +
						 SetN(EditSPP_BALL->Text)*0.1 +
						 SetN(EditOT_UPR->Text)*0.1 ) -
						 SetN(EditOT_TREB->Text)-SetN(EditTRUD_D->Text);
	}

}
//---------------------------------------------------------------------------

void __fastcall TVvod::EditPROIZ_BALLChange(TObject *Sender)
{
  if (!EditPROIZ_BALL->Text.IsEmpty()) {
	Ocenka();
  }
  else if (EditKPE_BALL->Text.IsEmpty() && EditPRIEM_BALL->Text.IsEmpty() &&
		   EditINFO_BALL->Text.IsEmpty() && EditOTKL_BALL->Text.IsEmpty() &&
		   EditC5_BALL->Text.IsEmpty() && EditKNS_BALL->Text.IsEmpty() &&
		   EditSPP_BALL->Text.IsEmpty() && EditOT_UPR->Text.IsEmpty() &&
		   EditOT_TREB->Text.IsEmpty() && EditTRUD_D->Text.IsEmpty()) {
		 EditOCENKA->Text = "";
	   }
}
//---------------------------------------------------------------------------

void __fastcall TVvod::EditKPE_BALLChange(TObject *Sender)
{
  if (!EditKPE_BALL->Text.IsEmpty()) {
	Ocenka();
  }
  else if (EditKPE_BALL->Text.IsEmpty() && EditPRIEM_BALL->Text.IsEmpty() &&
		   EditINFO_BALL->Text.IsEmpty() && EditOTKL_BALL->Text.IsEmpty() &&
		   EditC5_BALL->Text.IsEmpty() && EditKNS_BALL->Text.IsEmpty() &&
		   EditSPP_BALL->Text.IsEmpty() && EditOT_UPR->Text.IsEmpty() &&
		   EditOT_TREB->Text.IsEmpty() && EditTRUD_D->Text.IsEmpty()) {
		 EditOCENKA->Text = "";
	   }
}
//---------------------------------------------------------------------------

void __fastcall TVvod::EditPRIEM_BALLChange(TObject *Sender)
{
  if (!EditPRIEM_BALL->Text.IsEmpty()) {
	Ocenka();
  }
  else if (EditKPE_BALL->Text.IsEmpty() && EditPRIEM_BALL->Text.IsEmpty() &&
		   EditINFO_BALL->Text.IsEmpty() && EditOTKL_BALL->Text.IsEmpty() &&
		   EditC5_BALL->Text.IsEmpty() && EditKNS_BALL->Text.IsEmpty() &&
		   EditSPP_BALL->Text.IsEmpty() && EditOT_UPR->Text.IsEmpty() &&
		   EditOT_TREB->Text.IsEmpty() && EditTRUD_D->Text.IsEmpty()) {
		 EditOCENKA->Text = "";
	   }
}
//---------------------------------------------------------------------------

void __fastcall TVvod::EditINFO_BALLChange(TObject *Sender)
{
  if (!EditINFO_BALL->Text.IsEmpty()) {
	Ocenka();
  }
  else if (EditKPE_BALL->Text.IsEmpty() && EditPRIEM_BALL->Text.IsEmpty() &&
		   EditINFO_BALL->Text.IsEmpty() && EditOTKL_BALL->Text.IsEmpty() &&
		   EditC5_BALL->Text.IsEmpty() && EditKNS_BALL->Text.IsEmpty() &&
		   EditSPP_BALL->Text.IsEmpty() && EditOT_UPR->Text.IsEmpty() &&
		   EditOT_TREB->Text.IsEmpty() && EditTRUD_D->Text.IsEmpty()) {
		 EditOCENKA->Text = "";
	   }
}
//---------------------------------------------------------------------------

void __fastcall TVvod::EditOTKL_BALLChange(TObject *Sender)
{
  if (!EditOTKL_BALL->Text.IsEmpty()) {
	Ocenka();
  }
  else if (EditKPE_BALL->Text.IsEmpty() && EditPRIEM_BALL->Text.IsEmpty() &&
		   EditINFO_BALL->Text.IsEmpty() && EditOTKL_BALL->Text.IsEmpty() &&
		   EditC5_BALL->Text.IsEmpty() && EditKNS_BALL->Text.IsEmpty() &&
		   EditSPP_BALL->Text.IsEmpty() && EditOT_UPR->Text.IsEmpty() &&
		   EditOT_TREB->Text.IsEmpty() && EditTRUD_D->Text.IsEmpty()) {
		 EditOCENKA->Text = "";
	   }
}
//---------------------------------------------------------------------------

void __fastcall TVvod::EditC5_BALLChange(TObject *Sender)
{
  if (!EditC5_BALL->Text.IsEmpty()) {
	Ocenka();
  }
  else if (EditKPE_BALL->Text.IsEmpty() && EditPRIEM_BALL->Text.IsEmpty() &&
		   EditINFO_BALL->Text.IsEmpty() && EditOTKL_BALL->Text.IsEmpty() &&
		   EditC5_BALL->Text.IsEmpty() && EditKNS_BALL->Text.IsEmpty() &&
		   EditSPP_BALL->Text.IsEmpty() && EditOT_UPR->Text.IsEmpty() &&
		   EditOT_TREB->Text.IsEmpty() && EditTRUD_D->Text.IsEmpty()) {
		 EditOCENKA->Text = "";
	   }
}
//---------------------------------------------------------------------------

void __fastcall TVvod::EditKNS_BALLChange(TObject *Sender)
{
  if (!EditKNS_BALL->Text.IsEmpty()) {
	Ocenka();
  }
  else if (EditKPE_BALL->Text.IsEmpty() && EditPRIEM_BALL->Text.IsEmpty() &&
		   EditINFO_BALL->Text.IsEmpty() && EditOTKL_BALL->Text.IsEmpty() &&
		   EditC5_BALL->Text.IsEmpty() && EditKNS_BALL->Text.IsEmpty() &&
		   EditSPP_BALL->Text.IsEmpty() && EditOT_UPR->Text.IsEmpty() &&
		   EditOT_TREB->Text.IsEmpty() && EditTRUD_D->Text.IsEmpty()) {
		 EditOCENKA->Text = "";
	   }
}
//---------------------------------------------------------------------------

void __fastcall TVvod::EditSPP_BALLChange(TObject *Sender)
{
  if (!EditSPP_BALL->Text.IsEmpty()) {
	Ocenka();
  }
  else if (EditKPE_BALL->Text.IsEmpty() && EditPRIEM_BALL->Text.IsEmpty() &&
		   EditINFO_BALL->Text.IsEmpty() && EditOTKL_BALL->Text.IsEmpty() &&
		   EditC5_BALL->Text.IsEmpty() && EditKNS_BALL->Text.IsEmpty() &&
		   EditSPP_BALL->Text.IsEmpty() && EditOT_UPR->Text.IsEmpty() &&
		   EditOT_TREB->Text.IsEmpty() && EditTRUD_D->Text.IsEmpty()) {
		 EditOCENKA->Text = "";
	   }
}
//---------------------------------------------------------------------------

void __fastcall TVvod::EditOT_UPRChange(TObject *Sender)
{
  if (!EditOT_UPR->Text.IsEmpty()) {
	Ocenka();
  }
  else if (EditKPE_BALL->Text.IsEmpty() && EditPRIEM_BALL->Text.IsEmpty() &&
		   EditINFO_BALL->Text.IsEmpty() && EditOTKL_BALL->Text.IsEmpty() &&
		   EditC5_BALL->Text.IsEmpty() && EditKNS_BALL->Text.IsEmpty() &&
		   EditSPP_BALL->Text.IsEmpty() && EditOT_UPR->Text.IsEmpty() &&
		   EditOT_TREB->Text.IsEmpty() && EditTRUD_D->Text.IsEmpty()) {
		 EditOCENKA->Text = "";
	   }
}
//---------------------------------------------------------------------------

void __fastcall TVvod::EditOT_TREBChange(TObject *Sender)
{
  if (!EditOT_TREB->Text.IsEmpty()) {
	Ocenka();
  }
  else if (EditKPE_BALL->Text.IsEmpty() && EditPRIEM_BALL->Text.IsEmpty() &&
		   EditINFO_BALL->Text.IsEmpty() && EditOTKL_BALL->Text.IsEmpty() &&
		   EditC5_BALL->Text.IsEmpty() && EditKNS_BALL->Text.IsEmpty() &&
		   EditSPP_BALL->Text.IsEmpty() && EditOT_UPR->Text.IsEmpty() &&
		   EditOT_TREB->Text.IsEmpty() && EditTRUD_D->Text.IsEmpty()) {
		 EditOCENKA->Text = "";
	   }
}
//---------------------------------------------------------------------------

void __fastcall TVvod::EditTRUD_DChange(TObject *Sender)
{
  if (!EditTRUD_D->Text.IsEmpty()) {
	Ocenka();
  }
  else if (EditKPE_BALL->Text.IsEmpty() && EditPRIEM_BALL->Text.IsEmpty() &&
		   EditINFO_BALL->Text.IsEmpty() && EditOTKL_BALL->Text.IsEmpty() &&
		   EditC5_BALL->Text.IsEmpty() && EditKNS_BALL->Text.IsEmpty() &&
		   EditSPP_BALL->Text.IsEmpty() && EditOT_UPR->Text.IsEmpty() &&
		   EditOT_TREB->Text.IsEmpty() && EditTRUD_D->Text.IsEmpty()) {
		 EditOCENKA->Text = "";
	   }
}
//---------------------------------------------------------------------------

