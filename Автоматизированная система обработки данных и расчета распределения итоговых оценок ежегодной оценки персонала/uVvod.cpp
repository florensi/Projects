//---------------------------------------------------------------------------

#include <vcl.h>
#pragma hdrstop

#include "uVvod.h"
#include "uDM.h"
#include "uMain.h"
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
void __fastcall TVvod::FormShow(TObject *Sender)
{
  //Очистка Edit-ов
  EditZEX->Text = "";
  EditTN->Text = "490";
  EditDIREKT->Text = "";


  EditREALIZAC->Text = "";
  EditKACHESTVO->Text = "";
  EditRESURS->Text = "";
  EditSTAND->Text = "";
  EditPOTREB->Text = "";
  EditKACH->Text = "";
  EditEFF->Text = "";
  EditPROF_ZN->Text = "";
  EditLIDER->Text = "";
  EditOTVETSTV->Text = "";
  EditKOM_REZ->Text = "";

  EditKPE_OCEN->Text = "";
  EditREZULT_OCEN->Text = "";
  EditKOMP_OCEN->Text = "";

  EditDATA_OCEN->Text = "";
  EditFIO_OCEN->Text = "";
  EditDOLGO->Text = "";
  EditAVT_REIT->Text = "";

  EditSKOR_REIT->Text = "";
  EditKOM_REIT->Text = "";

  SetDataEdit();

  EditREZULT_OCEN->Enabled = false;
  EditKOMP_OCEN->Enabled = false;

  if (Main->god<Main->god_t)
    {
      EditKPE_OCEN->Enabled = false;

      EditREALIZAC->Enabled = false;
      EditKACHESTVO->Enabled = false;
      EditRESURS->Enabled = false;
      EditSTAND->Enabled = false;
      EditPOTREB->Enabled = false;
      EditKACH->Enabled = false;
      EditEFF->Enabled = false;
      EditPROF_ZN->Enabled = false;
      EditLIDER->Enabled = false;
      EditOTVETSTV->Enabled = false;
      EditKOM_REZ->Enabled = false;
    }
  else
    {
      EditKPE_OCEN->Enabled = true;

      EditREALIZAC->Enabled = true;
      EditKACHESTVO->Enabled = true;
      EditRESURS->Enabled = true;
      EditSTAND->Enabled = true;
      EditPOTREB->Enabled = true;
      EditKACH->Enabled = true;
      EditEFF->Enabled = true;
      EditPROF_ZN->Enabled = true;
      EditLIDER->Enabled = true;
      EditOTVETSTV->Enabled = true;
      EditKOM_REZ->Enabled = true;
   }

  EditZEX->SetFocus();
}
//---------------------------------------------------------------------------

//Заполнение Edit-oв
void __fastcall TVvod::SetDataEdit()
{
  AnsiString Sql;
  double rezult_n, komp_n;

  EditDATA_OCEN->Font->Color = clBlack;

  LabelZEX_NAIM->Caption = DM->qOcenka->FieldByName("naim_zex")->AsString;
  LabelDAT_JOB->Caption = zdat_job = DM->qOcenka->FieldByName("dat_job")->AsString;
  LabelDIREKT->Caption = DM->qOcenka->FieldByName("naim_direkt")->AsString;
  LabelNAME_DOLG->Caption = znaim_dolg = DM->qOcenka->FieldByName("dolg")->AsString;

  EditFIO->Text = zfio = DM->qOcenka->FieldByName("fio")->AsString;
  EditZEX->Text = zzex = DM->qOcenka->FieldByName("zex")->AsString;
  EditTN->Text = ztn = DM->qOcenka->FieldByName("tn")->AsString;

 /* if (!DM->qOcenka->FieldByName("dolg")->AsString.IsEmpty()) LabelNAME_DOLG->Caption = znaim_dolg = DM->qOcenka->FieldByName("dolg")->AsString;
  else
    {
      Sql = "select name_dolg_ru, dat_job, nzex,                                   \
                    case when ur1 is null then zex                            \
                         when ur2 is null then ur1                            \
                         when ur3 is null then ur2                            \
                         when ur4 is null then ur3 end ur,                    \
                    name_ur1                                                  \
             from sap_osn_sved where zex="+QuotedStr(EditZEX->Text)+" and tn_sap= "+EditTN->Text;
      DM->qObnovlenie->Close();
      DM->qObnovlenie->SQL->Clear();
      DM->qObnovlenie->SQL->Add(Sql);
      try
        {
          DM->qObnovlenie->Open();
        }
      catch (...)
        {
          Application->MessageBox("Возникла ошибка при получении данных с таблицы SAP_OSN_SVED","Ошибка", MB_OK+MB_ICONERROR);
          Abort();
        }

      if (DM->qObnovlenie->RecordCount==0)
        {
          //При необходимости оставить старый цех, выборка из SAP_PEREVOD
          Sql = "select name_dolg_ru, dat_job, nzex,                                   \
                        case when ur1 is null then zex                            \
                             when ur2 is null then ur1                            \
                             when ur3 is null then ur2                            \
                             when ur4 is null then ur3 end ur,                    \
                        name_ur1                                                  \
                 from sap_perevod where zex="+QuotedStr(EditZEX->Text)+" and tn_sap= "+EditTN->Text;
          DM->qObnovlenie->Close();
          DM->qObnovlenie->SQL->Clear();
          DM->qObnovlenie->SQL->Add(Sql);
          try
            {
              DM->qObnovlenie->Open();
            }
          catch (...)
            {
              Application->MessageBox("Возникла ошибка при получении данных с таблицы SAP_OSN_SVED","Ошибка", MB_OK+MB_ICONERROR);
              Abort();
            }

          if (DM->qObnovlenie->RecordCount>0)
            {
              LabelNAME_DOLG->Caption = DM->qObnovlenie->FieldByName("name_dolg_ru")->AsString;
              LabelDAT_JOB->Caption = DM->qObnovlenie->FieldByName("dat_job")->AsString;
              uch = DM->qObnovlenie->FieldByName("ur")->AsString;
              nuch = DM->qObnovlenie->FieldByName("name_ur1")->AsString;
            }
          else
            {
              //При необходимости оставить старый цех, выборка из P_PEREVOD
              Sql = "select name_dolg, dat_job, nzex,                                   \
                            null as ur,                    \
                            name_uch                                                  \
                            from p_perevod where zex="+QuotedStr(EditZEX->Text)+" and id_sap= "+EditTN->Text;
              DM->qObnovlenie->Close();
              DM->qObnovlenie->SQL->Clear();
              DM->qObnovlenie->SQL->Add(Sql);
              try
                {
                  DM->qObnovlenie->Open();
                }
              catch (...)
                {
                  Application->MessageBox("Возникла ошибка при получении данных с таблицы SAP_OSN_SVED","Ошибка", MB_OK+MB_ICONERROR);
                  Abort();
                }

              LabelNAME_DOLG->Caption = DM->qObnovlenie->FieldByName("name_dolg")->AsString;
              LabelDAT_JOB->Caption = DM->qObnovlenie->FieldByName("dat_job")->AsString;
              uch = DM->qObnovlenie->FieldByName("ur")->AsString;
              nuch = DM->qObnovlenie->FieldByName("name_uch")->AsString;
            }
          }
        else
         {
           LabelNAME_DOLG->Caption = DM->qObnovlenie->FieldByName("name_dolg_ru")->AsString;
           LabelDAT_JOB->Caption = DM->qObnovlenie->FieldByName("dat_job")->AsString;
           uch = DM->qObnovlenie->FieldByName("ur")->AsString;
           nuch = DM->qObnovlenie->FieldByName("name_ur1")->AsString;
         }
    }  */

  EditDIREKT->Text = zdirekt = DM->qOcenka->FieldByName("direkt")->AsString;
  EditDATA_OCEN->Text = zdata_ocen = DM->qOcenka->FieldByName("data_ocen")->AsString;
  EditFIO_OCEN->Text = zfio_ocen = DM->qOcenka->FieldByName("fio_ocen")->AsString;
  EditDOLGO->Text = zdolgo = DM->qOcenka->FieldByName("dolg_ocen")->AsString;
  EditAVT_REIT->Text = zavt_reit = DM->qOcenka->FieldByName("avt_reit")->AsString;
  EditSKOR_REIT->Text = zskor_ocen = DM->qOcenka->FieldByName("skor_reit")->AsString;
  EditKOM_REIT->Text = zkom_reit = DM->qOcenka->FieldByName("kom_reit")->AsString;

  EditREALIZAC->Text = zrealizac = DM->qOcenka->FieldByName("realizac")->AsString;
  EditKACHESTVO->Text = zkachestvo = DM->qOcenka->FieldByName("kachestvo")->AsString;
  EditRESURS->Text = zresurs = DM->qOcenka->FieldByName("resurs")->AsString;
  EditSTAND->Text = zstand = DM->qOcenka->FieldByName("stand")->AsString;
  EditPOTREB->Text = zpotreb = DM->qOcenka->FieldByName("potreb")->AsString;
  EditKACH->Text = zkach = DM->qOcenka->FieldByName("kach")->AsString;
  EditEFF->Text = zeff = DM->qOcenka->FieldByName("eff")->AsString;
  EditPROF_ZN->Text = zprof_zn = DM->qOcenka->FieldByName("prof_zn")->AsString;
  EditLIDER->Text = zlider = DM->qOcenka->FieldByName("lider")->AsString;
  EditOTVETSTV->Text = zotvetstv = DM->qOcenka->FieldByName("otvetstv")->AsString;
  EditKOM_REZ->Text = zkom_rez = DM->qOcenka->FieldByName("kom_rez")->AsString;

  EditKPE_OCEN->Text = zkpe_ocen = DM->qOcenka->FieldByName("kpe_ocen")->AsString;


  if (Main->god<Main->god_t)
    {
      EditREZULT_OCEN->Text = zrezult_ocen = DM->qOcenka->FieldByName("rezult_ocen")->AsString;
      EditKOMP_OCEN->Text = zkomp_ocen = DM->qOcenka->FieldByName("komp_ocen")->AsString;

      //% по результатам работы
      if (!EditREZULT_OCEN->Text.IsEmpty()) LabelREZULT_PROC->Caption = FloatToStrF(DM->qOcenka->FieldByName("rezult_proc")->AsFloat, ffFixed, 10,2) + " %";
      else LabelREZULT_PROC->Caption ="";

      //% по компетенции
      if (!EditKOMP_OCEN->Text.IsEmpty()) LabelKOMP_PROC->Caption = FloatToStrF(DM->qOcenka->FieldByName("komp_proc")->AsFloat, ffFixed,10,2) + " %";
      else LabelKOMP_PROC->Caption = "";

      //эффективность
      if ((!EditKPE_OCEN->Text.IsEmpty() && !EditKOMP_OCEN->Text.IsEmpty()) || (!EditREZULT_OCEN->Text.IsEmpty() && !EditKOMP_OCEN->Text.IsEmpty()))
        {
          effekt = DM->qOcenka->FieldByName("efect")->AsFloat;
        }
      else  effekt = 0;

      LabelEFFEKT->Caption = FloatToStrF(effekt, ffFixed,10,2) + " %";

    }
  else
    {
      rezult_n = (DM->qOcenka->FieldByName("realizac")->AsFloat+
                  DM->qOcenka->FieldByName("kachestvo")->AsFloat+
                  DM->qOcenka->FieldByName("resurs")->AsFloat);

      komp_n = (DM->qOcenka->FieldByName("stand")->AsFloat+
                DM->qOcenka->FieldByName("potreb")->AsFloat+
                DM->qOcenka->FieldByName("kach")->AsFloat+
                DM->qOcenka->FieldByName("eff")->AsFloat+
                DM->qOcenka->FieldByName("prof_zn")->AsFloat+
                DM->qOcenka->FieldByName("lider")->AsFloat+
                DM->qOcenka->FieldByName("otvetstv")->AsFloat+
                DM->qOcenka->FieldByName("kom_rez")->AsFloat);


      //расчет эффективности изменился с 2017 года
      //результаты работы
       if (rezult_n/3==0)
         {
           EditREZULT_OCEN->Text = zrezult_ocen = "";
           LabelREZULT_PROC->Caption ="";
         }
       else
         {
           EditREZULT_OCEN->Text = zrezult_ocen = FloatToStrF(rezult_n/3, ffFixed, 10,2);

           //% по результатам работы
           rezult = rezult_n/12*100;
           LabelREZULT_PROC->Caption = FloatToStrF(rezult, ffFixed, 10,2) + " %";
         }

       if (komp_n/8==0)
         {
           EditKOMP_OCEN->Text = zkomp_ocen = "";
           LabelKOMP_PROC->Caption = "";
         }
       else
         {
           EditKOMP_OCEN->Text = zkomp_ocen = komp_n;//FloatToStrF(komp_n, ffFixed,10,2);

           //% по компетенции
           komp = komp_n/32*100;
           LabelKOMP_PROC->Caption = FloatToStrF(komp, ffFixed,10,2) + " %";
         }


      //эффективность
      if (EditKPE_OCEN->Text.IsEmpty() && !EditREZULT_OCEN->Text.IsEmpty() && !EditKOMP_OCEN->Text.IsEmpty())
        {
          effekt = (((rezult_n/12*100)*0.6)+((komp_n/32*100)*0.4));
        }
      else if (!EditKPE_OCEN->Text.IsEmpty() && !EditKOMP_OCEN->Text.IsEmpty())
        {
          effekt = ((StrToFloat(EditKPE_OCEN->Text)*0.6)+((komp_n*0.4)/32*100));
        }
      else  effekt = 0;
        LabelEFFEKT->Caption = FloatToStrF(effekt, ffFixed,10,2) + " %";



      if (!EditREZULT_OCEN->Text.IsEmpty()) EditKPE_OCEN->Enabled = false;
      else EditKPE_OCEN->Enabled = true;

      if (!EditKPE_OCEN->Text.IsEmpty())
        {
          EditREALIZAC->Enabled = false;
          EditKACHESTVO->Enabled = false;
          EditRESURS->Enabled = false;
        }
      else
        {
          EditREALIZAC->Enabled = true;
          EditKACHESTVO->Enabled = true;
          EditRESURS->Enabled = true;
        }

    }

  ComboBoxKAT->ItemIndex = ComboBoxKAT->Items->IndexOf(DM->qOcenka->FieldByName("kat")->AsString);
  ComboBoxFUNCT_G->ItemIndex = ComboBoxFUNCT_G->Items->IndexOf(DM->qOcenka->FieldByName("funct_g")->AsString);
  ComboBoxFUNCT->ItemIndex = ComboBoxFUNCT->Items->IndexOf(DM->qOcenka->FieldByName("funct")->AsString);
  ComboBoxUU->ItemIndex = ComboBoxUU->Items->IndexOf(DM->qOcenka->FieldByName("uu")->AsString);

  zkat = DM->qOcenka->FieldByName("kat")->AsString;
  zfunct_g = DM->qOcenka->FieldByName("funct_g")->AsString;
  zfunct = DM->qOcenka->FieldByName("funct")->AsString;
  zuu = DM->qOcenka->FieldByName("uu")->AsString;
}
//---------------------------------------------------------------------------
void __fastcall TVvod::FormKeyDown(TObject *Sender, WORD &Key,
      TShiftState Shift)
{
  if (Key==VK_RETURN)
  FindNextControl((TWinControl *)Sender, true, true, false)->SetFocus();
}
//---------------------------------------------------------------------------

void __fastcall TVvod::EditREZULT_OCENExit(TObject *Sender)
{
/* if (ActiveControl == Cansel)
    {
      Vvod->Close();
    }
  else
    {
      if (!EditREZULT_OCEN->Text.IsEmpty())
        {
          EditKPE_OCEN->Enabled = false;

          if (StrToFloat(EditREZULT_OCEN->Text)>4)
            {
              Application->MessageBox("Максимальное значение по результатам \nработы не может превышать 4","Предупреждение", MB_OK+MB_ICONWARNING);

              EditREZULT_OCEN->SetFocus();
            }
        }
      else
        {
          EditKPE_OCEN->Enabled = true;
        }
    } */
}
//---------------------------------------------------------------------------

void __fastcall TVvod::EditZEXKeyPress(TObject *Sender, char &Key)
{
  if (!(IsNumeric(Key)||Key=='\b')) Key=0;         
}
//---------------------------------------------------------------------------

void __fastcall TVvod::Button1Click(TObject *Sender)
{
  AnsiString Sql, seffekt, srezult, skomp;
  int rec;
  TLocateOptions SearchOptions;

  if(effekt==0) seffekt="NULL";
  else seffekt=effekt;

  if(rezult==0) srezult="NULL";
  else srezult=rezult;

  if(komp==0) skomp="NULL";
  else skomp=komp;

  //проверка на заполнение цеха
  if (EditZEX->Text.IsEmpty())
    {
      Application->MessageBox("Не указан цех работника!!!","Предупреждение",
                              MB_OK+MB_ICONINFORMATION);
      EditZEX->SetFocus();
    }

  //проверка на 0 перед цехом
  if (EditZEX->Text.Length()==1) EditZEX->Text = "0"+EditZEX->Text;
  //проверка на 0 перед дирекцией
  if (EditDIREKT->Text.Length()==1 && !EditDIREKT->Text.IsEmpty()) EditDIREKT->Text = "0"+EditDIREKT->Text;

  //проверка на заполнение таб.№
  if (EditTN->Text.IsEmpty() || EditTN->Text=="490")
    {
      Application->MessageBox("Не указан таб.№ работника!!!","Предупреждение",
                              MB_OK+MB_ICONINFORMATION);
      EditTN->SetFocus();
    }

  //проверка на заполнение ФИО
  if (EditFIO->Text.IsEmpty())
    {
      Application->MessageBox("Не указана фамилия работника!!!","Предупреждение",
                              MB_OK+MB_ICONINFORMATION);
      EditTN->SetFocus();
    }

     //Введено все по результатам работ
  if (EditKPE_OCEN->Text.IsEmpty())
    {
      if ((!EditREALIZAC->Text.IsEmpty() && (EditKACHESTVO->Text.IsEmpty() || EditRESURS->Text.IsEmpty())) ||
          (!EditKACHESTVO->Text.IsEmpty() && (EditREALIZAC->Text.IsEmpty() || EditRESURS->Text.IsEmpty())) ||
          (!EditRESURS->Text.IsEmpty() && (EditKACHESTVO->Text.IsEmpty() || EditREALIZAC->Text.IsEmpty())))
        {
          Application->MessageBox("Указаны не все данные по результатам работ","Предупреждение",
                              MB_OK+MB_ICONINFORMATION);
          EditREALIZAC->SetFocus();
          Abort();
        }
    }

  //Введено все по компетенциям
  if ((!EditSTAND->Text.IsEmpty() && (EditPOTREB->Text.IsEmpty() || EditKACH->Text.IsEmpty() || EditEFF->Text.IsEmpty() || EditPROF_ZN->Text.IsEmpty() || EditLIDER->Text.IsEmpty() || EditOTVETSTV->Text.IsEmpty() || EditKOM_REZ->Text.IsEmpty())) ||
      (!EditPOTREB->Text.IsEmpty() && (EditSTAND->Text.IsEmpty() || EditKACH->Text.IsEmpty() || EditEFF->Text.IsEmpty() || EditPROF_ZN->Text.IsEmpty() || EditLIDER->Text.IsEmpty() || EditOTVETSTV->Text.IsEmpty() || EditKOM_REZ->Text.IsEmpty())) ||
      (!EditKACH->Text.IsEmpty() && (EditPOTREB->Text.IsEmpty() || EditSTAND->Text.IsEmpty() || EditEFF->Text.IsEmpty() || EditPROF_ZN->Text.IsEmpty() || EditLIDER->Text.IsEmpty() || EditOTVETSTV->Text.IsEmpty() || EditKOM_REZ->Text.IsEmpty())) ||
      (!EditEFF->Text.IsEmpty() && (EditPOTREB->Text.IsEmpty() || EditKACH->Text.IsEmpty() || EditSTAND->Text.IsEmpty() || EditPROF_ZN->Text.IsEmpty() || EditLIDER->Text.IsEmpty() || EditOTVETSTV->Text.IsEmpty() || EditKOM_REZ->Text.IsEmpty())) ||
      (!EditPROF_ZN->Text.IsEmpty() && (EditPOTREB->Text.IsEmpty() || EditKACH->Text.IsEmpty() || EditEFF->Text.IsEmpty() || EditSTAND->Text.IsEmpty() || EditLIDER->Text.IsEmpty() || EditOTVETSTV->Text.IsEmpty() || EditKOM_REZ->Text.IsEmpty())) ||
      (!EditLIDER->Text.IsEmpty() && (EditPOTREB->Text.IsEmpty() || EditKACH->Text.IsEmpty() || EditEFF->Text.IsEmpty() || EditPROF_ZN->Text.IsEmpty() || EditSTAND->Text.IsEmpty() || EditOTVETSTV->Text.IsEmpty() || EditKOM_REZ->Text.IsEmpty())) ||
      (!EditOTVETSTV->Text.IsEmpty() && (EditPOTREB->Text.IsEmpty() || EditKACH->Text.IsEmpty() || EditEFF->Text.IsEmpty() || EditPROF_ZN->Text.IsEmpty() || EditLIDER->Text.IsEmpty() || EditSTAND->Text.IsEmpty() || EditKOM_REZ->Text.IsEmpty())) ||
      (!EditKOM_REZ->Text.IsEmpty() && (EditPOTREB->Text.IsEmpty() || EditKACH->Text.IsEmpty() || EditEFF->Text.IsEmpty() || EditPROF_ZN->Text.IsEmpty() || EditLIDER->Text.IsEmpty() || EditOTVETSTV->Text.IsEmpty() || EditSTAND->Text.IsEmpty())))
    {
      Application->MessageBox("Указаны не все данные по компетенциям","Предупреждение",
                               MB_OK+MB_ICONINFORMATION);
      EditEFF->SetFocus();
      Abort();
    }

  if (zfio!=EditFIO->Text ||
      zzex!=EditZEX->Text ||
      ztn!=EditTN->Text ||
      znaim_dolg!=LabelNAME_DOLG->Caption ||
      zdat_job!=LabelDAT_JOB->Caption ||
      zdirekt!=EditDIREKT->Text ||
      zrezult_ocen!=EditREZULT_OCEN->Text ||
      zkpe_ocen!=EditKPE_OCEN->Text ||
      zkomp_ocen!=EditKOMP_OCEN->Text ||
      zdata_ocen!=EditDATA_OCEN->Text ||
      zfio_ocen!=EditFIO_OCEN->Text ||
      zdolgo!=EditDOLGO->Text ||
      zavt_reit!=EditAVT_REIT->Text ||
      zskor_ocen!=EditSKOR_REIT->Text ||
      zkom_reit!=EditKOM_REIT->Text ||
      zkat!=ComboBoxKAT->Text ||
      zfunct_g!=ComboBoxFUNCT_G->Text ||
      zfunct!=ComboBoxFUNCT->Text ||
      zuu!=ComboBoxUU->Text ||
      zrealizac!=EditREALIZAC->Text ||
      zkachestvo!=EditKACHESTVO->Text ||
      zresurs!=EditRESURS->Text ||
      zstand!=EditSTAND->Text ||
      zpotreb!=EditPOTREB->Text ||
      zkach!=EditKACH->Text ||
      zeff!=EditEFF->Text ||
      zprof_zn!=EditPROF_ZN->Text ||
      zlider!=EditLIDER->Text ||
      zotvetstv!=EditOTVETSTV->Text ||
      zkom_rez!=EditKOM_REZ->Text)
    {
      //Обновление записи
      Sql = "update ocenka set \
                       fio = initcap("+QuotedStr(EditFIO->Text)+"),\
                       zex = "+QuotedStr(EditZEX->Text)+",\
                       tn= "+EditTN->Text+",\
                       direkt = "+QuotedStr(EditDIREKT->Text)+",\
                       funct = "+QuotedStr(ComboBoxFUNCT->Text)+",\
                       uu = "+QuotedStr(ComboBoxUU->Text)+",\
                       funct_g = "+QuotedStr(ComboBoxFUNCT_G->Text)+",\
                       kat = "+QuotedStr(ComboBoxKAT->Text)+",\
                       rezult_ocen = "+Main->SetNull(EditREZULT_OCEN->Text)+",\
                       komp_ocen = "+Main->SetNull(EditKOMP_OCEN->Text)+",\
                       kpe_ocen = "+Main->SetNull(EditKPE_OCEN->Text)+",\
                       realizac = "+Main->SetNull(EditREALIZAC->Text)+",\
                       kachestvo = "+Main->SetNull(EditKACHESTVO->Text)+",\
                       resurs = "+Main->SetNull(EditRESURS->Text)+",\
                       stand = "+Main->SetNull(EditSTAND->Text)+",\
                       potreb = "+Main->SetNull(EditPOTREB->Text)+",\
                       kach = "+Main->SetNull(EditKACH->Text)+",\
                       eff = "+Main->SetNull(EditEFF->Text)+",\
                       prof_zn = "+Main->SetNull(EditPROF_ZN->Text)+",\
                       lider = "+Main->SetNull(EditLIDER->Text)+",\
                       otvetstv = "+Main->SetNull(EditOTVETSTV->Text)+",\
                       kom_rez = "+Main->SetNull(EditKOM_REZ->Text)+",\
                       rezult_proc = "+srezult+", \
                       komp_proc = "+skomp+", \
                       efect="+seffekt+",\
                       data_ocen = to_date("+QuotedStr(EditDATA_OCEN->Text)+",'dd.mm.yyyy'),\
                       fio_ocen = initcap("+QuotedStr(EditFIO_OCEN->Text)+"),\
                       dolg_ocen = "+QuotedStr(EditDOLGO->Text)+",\
                       avt_reit = "+QuotedStr(EditAVT_REIT->Text)+",\
                       skor_reit = "+QuotedStr(EditSKOR_REIT->Text)+",\
                       kom_reit = "+QuotedStr(EditKOM_REIT->Text);
      if (zzex!=EditZEX->Text || ztn!=EditTN->Text)
        {
          Sql+= ", dolg ="+ QuotedStr(LabelNAME_DOLG->Caption)+", uch="+QuotedStr(uch)+", nuch="+QuotedStr(nuch)+", dat_job=to_date("+ QuotedStr(LabelDAT_JOB->Caption)+",'dd.mm.yyyy')";
        }

      Sql+=" where rowid = chartorowid("+ QuotedStr(DM->qOcenka->FieldByName("rw")->AsString)+")";


      rec = DM->qOcenka->RecNo;
      DM->qObnovlenie->Close();
      DM->qObnovlenie->SQL->Clear();
      DM->qObnovlenie->SQL->Add(Sql);
      try
        {
          DM->qObnovlenie->ExecSQL();
        }
      catch(Exception &E)
        {
          Application->MessageBox(("Возникла ошибка при получении данных из таблицы Ocenka" + E.Message).c_str(),"Ошибка",
                                    MB_OK+MB_ICONERROR);
          Main->InsertLog("Возникла ошибка при обновлении данных за "+IntToStr(Main->god)+" год по работнику: цех="+EditZEX->Text+" таб.№="+EditTN->Text);
          DM->qLogs->Requery();
          Abort();
        }


      //Логи
      if (DM->qObnovlenie->RowsAffected>0)
        {
          AnsiString Str ="Обновление записи за "+IntToStr(Main->god)+" год: ";

          if (zzex!=EditZEX->Text || ztn!=EditTN->Text)  Str+="цех с '"+zzex+"' на '"+EditZEX->Text+"' таб.№ с '"+ztn+"' на '"+EditTN->Text+"'";
          else Str+="цех='"+zzex+"' таб.№ ='"+ztn+"'";

          if (zfio!=EditFIO->Text) Str+=", ФИО с '"+zfio+"' на '"+EditFIO->Text+"'";
          if (znaim_dolg!=LabelNAME_DOLG->Caption) Str+=", должность с '"+znaim_dolg+"' на '"+LabelNAME_DOLG->Caption+"'";
          if (zdat_job!=LabelDAT_JOB->Caption) Str+=", прием на долж. с '"+zdat_job+"' на '"+LabelDAT_JOB->Caption+"'";
          if (zdirekt!=EditDIREKT->Text) Str+=", шифр дирекции с '"+zdirekt+"' на '"+EditDIREKT->Text+"'";
          if (zrezult_ocen!=EditREZULT_OCEN->Text) Str+=", рез.раб. с '"+zrezult_ocen+"' на '"+EditREZULT_OCEN->Text+"'";
          if (zkpe_ocen!=EditKPE_OCEN->Text) Str+=", рез. по КЕ с '"+zkpe_ocen+"' на '"+EditKPE_OCEN->Text+"'";
          if (zkomp_ocen!=EditKOMP_OCEN->Text) Str+=", компетенции с '"+zkomp_ocen+"' на '"+EditKOMP_OCEN->Text+"'";
          if (zdata_ocen!=EditDATA_OCEN->Text) Str+=", дата оценки с '"+zdata_ocen+"' на '"+EditDATA_OCEN->Text+"'";
          if (zfio_ocen!=EditFIO_OCEN->Text) Str+=", ФИО оценщика с '"+zfio_ocen+"' на '"+EditFIO_OCEN->Text+"'";
          if (zdolgo!=EditDOLGO->Text) Str+=", долж.оценщика с '"+zdolgo+"' на '"+EditDOLGO->Text+"'";
          if (zavt_reit!=EditAVT_REIT->Text) Str+=", авт.рейт. с '"+zavt_reit+"' на '"+EditAVT_REIT->Text+"'";
          if (zskor_ocen!=EditSKOR_REIT->Text) Str+=", скор.рейт. с '"+zskor_ocen+"' на '"+EditSKOR_REIT->Text+"'";
          if (zkom_reit!=EditKOM_REIT->Text) Str+=", ком.рейт. с '"+zkom_reit+"' на '"+EditKOM_REIT->Text+"'";
          if (zkat!=ComboBoxKAT->Text) Str+=", категория с '"+zkat+"' на '"+ComboBoxKAT->Text+"'";
          if (zfunct_g!=ComboBoxFUNCT_G->Text) Str+=", функц.группа с '"+zfunct_g+"' на '"+ComboBoxFUNCT_G->Text+"'";
          if (zfunct!=ComboBoxFUNCT->Text) Str+=", функция с '"+zfunct+"' на '"+ComboBoxFUNCT->Text+"'";
          if (zuu!=ComboBoxUU->Text) Str+=", ур.управ. с '"+zuu+"' на '"+ComboBoxUU->Text+"'";
          if (zrealizac!=EditREALIZAC->Text) Str+= ", степень реализ. задач с'"+zrealizac+"' на '"+EditREALIZAC->Text+"'";
          if (zkachestvo!=EditKACHESTVO->Text) Str+=", качество с '"+zkachestvo+"' на '"+EditKACHESTVO->Text+"'";
          if (zresurs!=EditRESURS->Text) Str+=", эконом. ресурсов с '"+zresurs+"' на '"+EditRESURS->Text+"'";
          if (zstand!=EditSTAND->Text) Str+=", стандарты ОТ, ПБ с '"+zstand+"' на '"+EditSTAND->Text+"'";
          if (zpotreb!=EditPOTREB->Text) Str+=", ориент. на потреб. с '"+zpotreb+"' на '"+EditPOTREB->Text+"'";
          if (zkach!=EditKACH->Text) Str+=", обеспеч. качество с '"+zkach+"' на '"+EditKACH->Text+"'";
          if (zeff!=EditEFF->Text) Str+=", дейст. эффективно с '"+zeff+"' на '"+EditEFF->Text+"'";
          if (zprof_zn!=EditPROF_ZN->Text) Str+=", проф. знания с '"+zprof_zn+"' на '"+EditPROF_ZN->Text+"'";
          if (zlider!=EditLIDER->Text) Str+=", быть лидером с '"+zlider+"' на '"+EditLIDER->Text+"'";
          if (zotvetstv!=EditOTVETSTV->Text) Str+=", быть ответственным с '"+zotvetstv+"' на '"+EditOTVETSTV->Text+"'";
          if (zkom_rez!=EditKOM_REZ->Text) Str+=", командный результат с '"+zkom_rez+"' на '"+EditKOM_REZ->Text+"'";

          Main->InsertLog(Str);
          DM->qLogs->Requery();

        }
      else
        {
          Main->InsertLog("Обновление данных за "+IntToStr(Main->god)+" год по работнику: цех="+EditZEX->Text+" таб.№="+EditTN->Text+" не выполнено");
          DM->qLogs->Requery();
        }


      DM->qOcenka->Requery();
      //Возвращение на выбранную строку
      DM->qOcenka->RecNo = rec;

      Application->MessageBox("Запись успешно изменена","Предупреждение",
                               MB_OK+MB_ICONINFORMATION);

    }

  Vvod->Close();
}
//---------------------------------------------------------------------------

void __fastcall TVvod::EditDATA_OCENExit(TObject *Sender)
{
  TDateTime d;

  if (ActiveControl == Cansel)
    {
      Vvod->Close();
    }
  else
    {
      if (!EditDATA_OCEN->Text.IsEmpty())
        {
          // Добавление к дате отчетного месяца и года
          if (EditDATA_OCEN->Text.Length()<3)
            {
              if(EditDATA_OCEN->Text.Pos("."))
                {
                  Application->MessageBox("Неверный формат даты","Ошибка", MB_OK+MB_ICONINFORMATION);
                  EditDATA_OCEN->Font->Color = clRed;
                  EditDATA_OCEN->SetFocus();
                  Abort();
                }
              else
                {
                  EditDATA_OCEN->Text = EditDATA_OCEN->Text+ "."+ DateToStr(Date()).SubString(4,2) +"."+ DateToStr(Date()).SubString(7,5);
                  EditDATA_OCEN->Font->Color = clBlack;
                }
            }

          // Проверка на правильность ввода даты
          if(!TryStrToDate(EditDATA_OCEN->Text,d))
            {
              Application->MessageBox("Неверный формат даты","Ошибка", MB_OK);
              EditDATA_OCEN->Font->Color = clRed;
              EditDATA_OCEN->SetFocus();
            }
          else
            {
              EditDATA_OCEN->Text=FormatDateTime("dd.mm.yyyy",d);
              EditDATA_OCEN->Font->Color = clBlack;
            }

        }
    }
}
//---------------------------------------------------------------------------

//Изменение % результатов работы
void __fastcall TVvod::EditREZULT_OCENChange(TObject *Sender)
{
  if (Main->god<2017)
    {
      if (!EditREZULT_OCEN->Text.IsEmpty())
        {
          EditKPE_OCEN->Enabled = false;

          //% по результатам работы
          rezult = StrToFloat(EditREZULT_OCEN->Text)/4*100;
          LabelREZULT_PROC->Caption = FloatToStrF(rezult, ffFixed, 10,2) + " %";

          if (!EditKOMP_OCEN->Text.IsEmpty())
            {
              effekt = (((StrToFloat(EditREZULT_OCEN->Text)*0.6)/4*100)+((StrToFloat(EditKOMP_OCEN->Text)*0.4)/32*100));
              LabelEFFEKT->Caption = FloatToStrF(effekt, ffFixed,10,2) + " %";
            }
          else
            {
              effekt=0;
              LabelEFFEKT->Caption = FloatToStrF(effekt, ffFixed,10,2) + " %";
            }
        }
      else
        {
          EditKPE_OCEN->Enabled = true;
          rezult = 0;
          LabelREZULT_PROC->Caption = FloatToStrF(rezult, ffFixed, 10,2) + " %";
          effekt = 0;
          LabelEFFEKT->Caption = FloatToStrF(effekt, ffFixed,10,2) + " %";
        }
    }
}
//---------------------------------------------------------------------------

//Изменение % компетенции
void __fastcall TVvod::EditKOMP_OCENChange(TObject *Sender)
{
  if (Main->god<2017)
    {
      if (!EditKOMP_OCEN->Text.IsEmpty())
        {
          //% по компетенции
          komp = StrToFloat(EditKOMP_OCEN->Text)/32*100;
          LabelKOMP_PROC->Caption = FloatToStrF(komp, ffFixed,10,2) + " %";

          //эффективность
          if (EditKPE_OCEN->Text.IsEmpty() && !EditREZULT_OCEN->Text.IsEmpty() && !EditKOMP_OCEN->Text.IsEmpty())
            {
              effekt = (((StrToFloat(EditREZULT_OCEN->Text)*0.6)/4*100)+((StrToFloat(EditKOMP_OCEN->Text)*0.4)/32*100));
              LabelEFFEKT->Caption = FloatToStrF(effekt, ffFixed,10,2) + " %";
            }
          else if (!EditKPE_OCEN->Text.IsEmpty() && !EditKOMP_OCEN->Text.IsEmpty())
            {
              effekt = ((StrToFloat(EditKPE_OCEN->Text)*0.6)+((StrToFloat(EditKOMP_OCEN->Text)*0.4)/32*100));
              LabelEFFEKT->Caption = FloatToStrF(effekt, ffFixed,10,2) + " %";
            }
          else
            {
              effekt = 0;
              LabelEFFEKT->Caption = FloatToStrF(effekt, ffFixed,10,2) + " %";
            }
        }
      else
        {
          komp = 0;
          LabelKOMP_PROC->Caption = FloatToStrF(komp, ffFixed,10,2) + " %";
          effekt = 0;
          LabelEFFEKT->Caption = FloatToStrF(effekt, ffFixed,10,2) + " %";
        }
    }            
}
//---------------------------------------------------------------------------

//Изменение % по КПЕ
void __fastcall TVvod::EditKPE_OCENChange(TObject *Sender)
{
  if (Main->god<2017)
    {
      if (!EditKPE_OCEN->Text.IsEmpty())
        {
          if (!EditKOMP_OCEN->Text.IsEmpty())
            {
              effekt = ((StrToFloat(EditKPE_OCEN->Text)*0.6)+((StrToFloat(EditKOMP_OCEN->Text)*0.4)/32*100));
              LabelEFFEKT->Caption = FloatToStrF(effekt, ffFixed,10,2) + " %";
            }
          else
            {
              effekt = 0;
              LabelEFFEKT->Caption = FloatToStrF(effekt, ffFixed,10,2) + " %";
            }
        }
      else
        {
          if (!EditKPE_OCEN->Text.IsEmpty() && !EditKPE_OCEN->Text.IsEmpty()) effekt = (((StrToFloat(EditREZULT_OCEN->Text)*0.6)/4*100)+((StrToFloat(EditKOMP_OCEN->Text)*0.4)/32*100));
          else effekt = 0;
          LabelEFFEKT->Caption = FloatToStrF(effekt, ffFixed,10,2) + " %";
        }
    }
  else
    {
      if (!EditKPE_OCEN->Text.IsEmpty())
        {
          if (!EditKOMP_OCEN->Text.IsEmpty())
            {
              effekt = ((StrToFloat(EditKPE_OCEN->Text)*0.6)+((Main->SetNullF(EditSTAND->Text)+
                                                               Main->SetNullF(EditPOTREB->Text)+
                                                               Main->SetNullF(EditKACH->Text)+
                                                               Main->SetNullF(EditEFF->Text)+
                                                               Main->SetNullF(EditPROF_ZN->Text)+
                                                               Main->SetNullF(EditLIDER->Text)+
                                                               Main->SetNullF(EditOTVETSTV->Text)+
                                                               Main->SetNullF(EditKOM_REZ->Text))/32*100)*0.4);
              LabelEFFEKT->Caption = FloatToStrF(effekt, ffFixed,10,2) + " %";
              EditREZULT_OCEN->Text="";
            }
          else
            {
              effekt = 0;
              LabelEFFEKT->Caption = FloatToStrF(effekt, ffFixed,10,2) + " %";
            }

          EditREALIZAC->Enabled = false;
          EditKACHESTVO->Enabled = false;
          EditRESURS->Enabled = false;
        }
      else
        {
          effekt =  ((((Main->SetNullF(EditREALIZAC->Text)+
                        Main->SetNullF(EditKACHESTVO->Text)+
                        Main->SetNullF(EditRESURS->Text))/12*100)*0.6)+
                    (((Main->SetNullF(EditSTAND->Text)+
                       Main->SetNullF(EditPOTREB->Text)+
                       Main->SetNullF(EditKACH->Text)+
                       Main->SetNullF(EditEFF->Text)+
                       Main->SetNullF(EditPROF_ZN->Text)+
                       Main->SetNullF(EditLIDER->Text)+
                       Main->SetNullF(EditOTVETSTV->Text)+
                       Main->SetNullF(EditKOM_REZ->Text))/32*100)*0.4)
                     );
          LabelEFFEKT->Caption = FloatToStrF(effekt, ffFixed,10,2) + " %";

          EditREALIZAC->Enabled = true;
          EditKACHESTVO->Enabled = true;
          EditRESURS->Enabled = true;

          
        }
    }
}
//---------------------------------------------------------------------------

void __fastcall TVvod::EditKPE_OCENExit(TObject *Sender)
{
    if (!EditKPE_OCEN->Text.IsEmpty())
    {
      EditREALIZAC->Enabled = false;
      EditKACHESTVO->Enabled = false;
      EditRESURS->Enabled = false;


      EditREALIZAC->Text = "";
      EditKACHESTVO->Text = "";
      EditRESURS->Text = "";
    }
  else
    {
      EditREALIZAC->Enabled = true;
      EditKACHESTVO->Enabled = true;
      EditRESURS->Enabled = true;
      
    }
}
//---------------------------------------------------------------------------

void __fastcall TVvod::EditREZULT_OCENKeyPress(TObject *Sender, char &Key)
{
  if (! (IsNumeric(Key) || Key=='.' || Key==',' || Key=='/' || Key=='\b') ) Key=0;
  if (Key==',' || Key=='/') Key='.';

}
//---------------------------------------------------------------------------

void __fastcall TVvod::EditFIO_OCENKeyPress(TObject *Sender, char &Key)
{
  if (IsNumeric(Key)) Key=0;
}
//---------------------------------------------------------------------------


void __fastcall TVvod::EditDIREKTChange(TObject *Sender)
{
  TLocateOptions SearchOptions;
  if (!EditDIREKT->Text.IsEmpty())
    {
      //Вывод наименования дерекции
      if (DM->qSprav->Locate("zex", EditDIREKT->Text,
                          SearchOptions << loCaseInsensitive))
        {
          LabelDIREKT->Caption = DM->qSprav->FieldByName("naim_direkt")->AsString;
        }
      else
        {
          LabelDIREKT->Caption = "";
        }
    }
}
//---------------------------------------------------------------------------


void __fastcall TVvod::EditKOMP_OCENExit(TObject *Sender)
{
/* if (ActiveControl == Cansel)
    {
      Vvod->Close();
    }
  else
    {
      if (!EditKOMP_OCEN->Text.IsEmpty())
        {
          if (StrToFloat(EditKOMP_OCEN->Text)>32)
            {
              Application->MessageBox("Максимальное значение по компетенции \nне может превышать 32","Предупреждение", MB_OK+MB_ICONWARNING);

              EditKOMP_OCEN->SetFocus();
            }
        }
    } */
}
//---------------------------------------------------------------------------



void __fastcall TVvod::EditZEXChange(TObject *Sender)
{
 AnsiString Sql;
 TLocateOptions SearchOptions;

 //Вывод должности
  if (!EditZEX->Text.IsEmpty())
    {
      if (zzex!=EditZEX->Text)
        {
          if (EditTN->Text.IsEmpty() || EditTN->Text=="490")
            {
              Abort();
            }
          else
            {
              Sql = "select name_dolg_ru, dat_job, nzex,                                   \
                            case when ur1 is null then zex                            \
                                 when ur2 is null then ur1                            \
                                 when ur3 is null then ur2                            \
                                 when ur4 is null then ur3 end ur,                    \
                            name_ur1                                                  \
                     from sap_osn_sved where zex="+QuotedStr(EditZEX->Text)+" and tn_sap= "+EditTN->Text;
              DM->qObnovlenie->Close();
              DM->qObnovlenie->SQL->Clear();
              DM->qObnovlenie->SQL->Add(Sql);
              try
                {
                  DM->qObnovlenie->Open();
                }
              catch (...)
                {
                  Application->MessageBox("Возникла ошибка при получении данных с таблицы SAP_OSN_SVED","Ошибка", MB_OK+MB_ICONERROR);
                  Abort();
                }

              if (DM->qObnovlenie->RecordCount==0)
                {
                  //При необходимости оставить старый цех, выборка из SAP_PEREVOD
                  Sql = "select name_dolg_ru, dat_job, nzex,                                   \
                            case when ur1 is null then zex                            \
                                 when ur2 is null then ur1                            \
                                 when ur3 is null then ur2                            \
                                 when ur4 is null then ur3 end ur,                    \
                            name_ur1                                                  \
                     from sap_perevod where zex="+QuotedStr(EditZEX->Text)+" and tn_sap= "+EditTN->Text;
                  DM->qObnovlenie->Close();
                  DM->qObnovlenie->SQL->Clear();
                  DM->qObnovlenie->SQL->Add(Sql);
                  try
                    {
                      DM->qObnovlenie->Open();
                    }
                  catch (...)
                    {
                      Application->MessageBox("Возникла ошибка при получении данных с таблицы SAP_OSN_SVED","Ошибка", MB_OK+MB_ICONERROR);
                      Abort();
                    }

                  if (DM->qObnovlenie->RecordCount>0)
                    {
                      LabelZEX_NAIM->Caption = DM->qObnovlenie->FieldByName("nzex")->AsString;
                      LabelNAME_DOLG->Caption = DM->qObnovlenie->FieldByName("name_dolg_ru")->AsString;
                      LabelDAT_JOB->Caption = DM->qObnovlenie->FieldByName("dat_job")->AsString;
                      uch = DM->qObnovlenie->FieldByName("ur")->AsString;
                      nuch = DM->qObnovlenie->FieldByName("name_ur1")->AsString;
                    }
                  else
                    {
                      //При необходимости оставить старый цех, выборка из P_PEREVOD
                      Sql = "select name_dolg, dat_job, nzex,                                   \
                            null as ur,                    \
                            name_uch                                                  \
                            from p_perevod where zex="+QuotedStr(EditZEX->Text)+" and id_sap= "+EditTN->Text;
                      DM->qObnovlenie->Close();
                      DM->qObnovlenie->SQL->Clear();
                      DM->qObnovlenie->SQL->Add(Sql);
                      try
                        {
                          DM->qObnovlenie->Open();
                        }
                      catch (...)
                        {
                          Application->MessageBox("Возникла ошибка при получении данных с таблицы SAP_OSN_SVED","Ошибка", MB_OK+MB_ICONERROR);
                          Abort();
                        }

                      LabelZEX_NAIM->Caption = DM->qObnovlenie->FieldByName("nzex")->AsString;
                      LabelNAME_DOLG->Caption = DM->qObnovlenie->FieldByName("name_dolg")->AsString;
                      LabelDAT_JOB->Caption = DM->qObnovlenie->FieldByName("dat_job")->AsString;
                      uch = DM->qObnovlenie->FieldByName("ur")->AsString;
                      nuch = DM->qObnovlenie->FieldByName("name_uch")->AsString;
                    }
                }
              else
                {
                  LabelZEX_NAIM->Caption = DM->qObnovlenie->FieldByName("nzex")->AsString;
                  LabelNAME_DOLG->Caption = DM->qObnovlenie->FieldByName("name_dolg_ru")->AsString;
                  LabelDAT_JOB->Caption = DM->qObnovlenie->FieldByName("dat_job")->AsString;
                  uch = DM->qObnovlenie->FieldByName("ur")->AsString;
                  nuch = DM->qObnovlenie->FieldByName("name_ur1")->AsString;
                }
            }
        }
    }
}
//---------------------------------------------------------------------------

void __fastcall TVvod::EditTNChange(TObject *Sender)
{
 AnsiString Sql;

  //Вывод должности
  if (!EditTN->Text.IsEmpty() && EditTN->Text!="490")
    {
      if (ztn!=EditTN->Text)
        {
          if (EditZEX->Text.IsEmpty())
            {
              Application->MessageBox("Не указан цех", "Предупреждение",
                                       MB_OK+MB_ICONINFORMATION);
              EditZEX->SetFocus();
              Abort();
            }
          else
            {
              Sql = "select name_dolg_ru, dat_job from sap_osn_sved where zex="+QuotedStr(EditZEX->Text)+" and tn_sap= "+EditTN->Text;
              DM->qObnovlenie->Close();
              DM->qObnovlenie->SQL->Clear();
              DM->qObnovlenie->SQL->Add(Sql);
              try
                {
                  DM->qObnovlenie->Open();
                }
              catch (...)
                {
                  Application->MessageBox("Возникла ошибка при получении данных с таблицы SAP_OSN_SVED","Ошибка", MB_OK+MB_ICONERROR);
                  Abort();
                }

              LabelNAME_DOLG->Caption = DM->qObnovlenie->FieldByName("name_dolg_ru")->AsString;
              LabelDAT_JOB->Caption = DM->qObnovlenie->FieldByName("dat_job")->AsString;
            }
        }
    }
}
//---------------------------------------------------------------------------

void __fastcall TVvod::EditDIREKTExit(TObject *Sender)
{
  AnsiString Sql;
  TLocateOptions SearchOptions;

  if (ActiveControl == Cansel)
    {
      Vvod->Close();
    }
  else
    {
      if (!EditDIREKT->Text.IsEmpty())
        {
          if (!DM->qSprav->Locate("zex", EditDIREKT->Text,
                                  SearchOptions << loCaseInsensitive))
            {
              Application->MessageBox("Введенной дирекции нет в справочнике!","Предупреждение",
                                       MB_OK+MB_ICONWARNING);
              EditDIREKT->SetFocus();
              Abort();
            }
        }    
    }
}
//---------------------------------------------------------------------------




void __fastcall TVvod::EditZEX_REZKeyPress(TObject *Sender, char &Key)
{
  if (! (IsNumeric(Key) || Key=='.' || Key==',' || Key=='/' || Key=='\b') ) Key=0;
  if (Key==',' || Key=='/') Key='.';        
}
//---------------------------------------------------------------------------

void __fastcall TVvod::EditSHIFR_REZKeyPress(TObject *Sender, char &Key)
{
  if (! (IsNumeric(Key) || Key=='.' || Key==',' || Key=='/' || Key=='\b') ) Key=0;
  if (Key==',' || Key=='/') Key='.';        
}
//---------------------------------------------------------------------------





void __fastcall TVvod::EditZEXKeyDown(TObject *Sender, WORD &Key,
      TShiftState Shift)
{
  if (Key==VK_RETURN)
  FindNextControl((TWinControl *)Sender, true, true, false)->SetFocus();
  EditTN->SelStart=EditTN->Text.Length();
}
//---------------------------------------------------------------------------


void __fastcall TVvod::EditREALIZACExit(TObject *Sender)
{
  if (ActiveControl == Cansel)
    {
      Vvod->Close();
    }
  else
    {
      if (!EditREALIZAC->Text.IsEmpty())
        {
          EditKPE_OCEN->Enabled = false;

          if (StrToFloat(EditREALIZAC->Text)>4)
            {
              Application->MessageBox("Максимальное значение по результатам \nработы не может превышать 4","Предупреждение", MB_OK+MB_ICONWARNING);

              EditREALIZAC->SetFocus();
            }
        }
      else if (!EditKACHESTVO->Text.IsEmpty() || !EditRESURS->Text.IsEmpty())
        {
          EditKPE_OCEN->Enabled = false;
        }
      else
        {
          EditKPE_OCEN->Enabled = true;
        }
    }        
}
//---------------------------------------------------------------------------

void __fastcall TVvod::EditKACHESTVOExit(TObject *Sender)
{
 if (ActiveControl == Cansel)
    {
      Vvod->Close();
    }
  else
    {
      if (!EditREALIZAC->Text.IsEmpty())
        {
          EditKPE_OCEN->Enabled = false;

          if (StrToFloat(EditREALIZAC->Text)>4)
            {
              Application->MessageBox("Максимальное значение по результатам \nработы не может превышать 4","Предупреждение", MB_OK+MB_ICONWARNING);

              EditREALIZAC->SetFocus();
            }
        }
      else if (!EditKACHESTVO->Text.IsEmpty() || !EditRESURS->Text.IsEmpty())
        {
          EditKPE_OCEN->Enabled = false;
        }
      else
        {
          EditKPE_OCEN->Enabled = true;
        }
    }        
}
//---------------------------------------------------------------------------

void __fastcall TVvod::EditRESURSExit(TObject *Sender)
{
 if (ActiveControl == Cansel)
    {
      Vvod->Close();
    }
  else
    {
      if (!EditREALIZAC->Text.IsEmpty())
        {
          EditKPE_OCEN->Enabled = false;

          if (StrToFloat(EditREALIZAC->Text)>4)
            {
              Application->MessageBox("Максимальное значение по результатам \nработы не может превышать 4","Предупреждение", MB_OK+MB_ICONWARNING);

              EditREALIZAC->SetFocus();
            }
        }
      else if (!EditKACHESTVO->Text.IsEmpty() || !EditRESURS->Text.IsEmpty())
        {
          EditKPE_OCEN->Enabled = false;
        }
      else
        {
          EditKPE_OCEN->Enabled = true;
        }
    }
}
//---------------------------------------------------------------------------

void __fastcall TVvod::EditSTANDExit(TObject *Sender)
{
 if (ActiveControl == Cansel)
    {
      Vvod->Close();
    }
  else
    {
      if (!EditSTAND->Text.IsEmpty() && StrToFloat(EditSTAND->Text)>4)
        {
          Application->MessageBox("Максимальное значение по компетенции \n не может превышать 4","Предупреждение", MB_OK+MB_ICONWARNING);
          EditSTAND->SetFocus();
        }
    } 
}
//---------------------------------------------------------------------------

void __fastcall TVvod::EditPOTREBExit(TObject *Sender)
{
  if (ActiveControl == Cansel)
    {
      Vvod->Close();
    }
  else
    {
      if (!EditPOTREB->Text.IsEmpty() && StrToFloat(EditPOTREB->Text)>4)
        {
          Application->MessageBox("Максимальное значение по компетенции \n не может превышать 4","Предупреждение", MB_OK+MB_ICONWARNING);
          EditPOTREB->SetFocus();
        }
    }
}
//---------------------------------------------------------------------------

void __fastcall TVvod::EditKACHExit(TObject *Sender)
{
 if (ActiveControl == Cansel)
    {
      Vvod->Close();
    }
  else
    {
      if (!EditKACH->Text.IsEmpty() && StrToFloat(EditKACH->Text)>4)
        {
          Application->MessageBox("Максимальное значение по компетенции \n не может превышать 4","Предупреждение", MB_OK+MB_ICONWARNING);
          EditKACH->SetFocus();
        }
    }
}
//---------------------------------------------------------------------------

void __fastcall TVvod::EditEFFExit(TObject *Sender)
{
 if (ActiveControl == Cansel)
    {
      Vvod->Close();
    }
  else
    {
      if (!EditEFF->Text.IsEmpty() && StrToFloat(EditEFF->Text)>4)
        {
          Application->MessageBox("Максимальное значение по компетенции \n не может превышать 4","Предупреждение", MB_OK+MB_ICONWARNING);
          EditEFF->SetFocus();
        }
    }
}
//---------------------------------------------------------------------------

void __fastcall TVvod::EditPROF_ZNExit(TObject *Sender)
{
 if (ActiveControl == Cansel)
    {
      Vvod->Close();
    }
  else
    {
      if (!EditPROF_ZN->Text.IsEmpty() && StrToFloat(EditPROF_ZN->Text)>4)
        {
          Application->MessageBox("Максимальное значение по компетенции \n не может превышать 4","Предупреждение", MB_OK+MB_ICONWARNING);
          EditPROF_ZN->SetFocus();
        }
    }
}
//---------------------------------------------------------------------------

void __fastcall TVvod::EditLIDERExit(TObject *Sender)
{
  if (ActiveControl == Cansel)
    {
      Vvod->Close();
    }
  else
    {
      if (!EditLIDER->Text.IsEmpty() && StrToFloat(EditLIDER->Text)>4)
        {
          Application->MessageBox("Максимальное значение по компетенции \n не может превышать 4","Предупреждение", MB_OK+MB_ICONWARNING);
          EditLIDER->SetFocus();
        }
    }
}
//---------------------------------------------------------------------------

void __fastcall TVvod::EditOTVETSTVExit(TObject *Sender)
{
  if (ActiveControl == Cansel)
    {
      Vvod->Close();
    }
  else
    {
      if (!EditOTVETSTV->Text.IsEmpty() && StrToFloat(EditOTVETSTV->Text)>4)
        {
          Application->MessageBox("Максимальное значение по компетенции \n не может превышать 4","Предупреждение", MB_OK+MB_ICONWARNING);
          EditOTVETSTV->SetFocus();
        }
    }
}
//---------------------------------------------------------------------------

void __fastcall TVvod::EditKOM_REZExit(TObject *Sender)
{
   if (ActiveControl == Cansel)
    {
      Vvod->Close();
    }
  else
    {
      if (!EditKOM_REZ->Text.IsEmpty() && StrToFloat(EditKOM_REZ->Text)>4)
        {
          Application->MessageBox("Максимальное значение по компетенции \n не может превышать 4","Предупреждение", MB_OK+MB_ICONWARNING);
          EditKOM_REZ->SetFocus();
        }
    }
}
//---------------------------------------------------------------------------

void __fastcall TVvod::EditREALIZACChange(TObject *Sender)
{
  IzmenRezRab();
}
//---------------------------------------------------------------------------
void __fastcall TVvod::IzmenRezRab()
{
  double rezult_n;
  
  if (!EditREALIZAC->Text.IsEmpty() || !EditKACHESTVO->Text.IsEmpty() || !EditRESURS->Text.IsEmpty())
    {
      EditKPE_OCEN->Enabled = false;

      //% по результатам работы
      rezult_n = (Main->SetNullF(EditREALIZAC->Text)+
                  Main->SetNullF(EditKACHESTVO->Text)+
                  Main->SetNullF(EditRESURS->Text));

      rezult = StrToFloat(FloatToStrF(rezult_n/12*100, ffFixed, 10,2));


      EditREZULT_OCEN->Text =  FloatToStrF(rezult_n/3, ffFixed, 10,2);
      LabelREZULT_PROC->Caption = FloatToStrF(rezult_n/12*100, ffFixed, 10,2) + " %";

      if (!EditKOMP_OCEN->Text.IsEmpty())
        {

          effekt =((((Main->SetNullF(EditREALIZAC->Text)+
                      Main->SetNullF(EditKACHESTVO->Text)+
                      Main->SetNullF(EditRESURS->Text))/12*100)*0.6)+
                   (((Main->SetNullF(EditSTAND->Text)+
                      Main->SetNullF(EditPOTREB->Text)+
                      Main->SetNullF(EditKACH->Text)+
                      Main->SetNullF(EditEFF->Text)+
                      Main->SetNullF(EditPROF_ZN->Text)+
                      Main->SetNullF(EditLIDER->Text)+
                      Main->SetNullF(EditOTVETSTV->Text)+
                      Main->SetNullF(EditKOM_REZ->Text))/32*100)*0.4)
                    );

          LabelEFFEKT->Caption = FloatToStrF(effekt, ffFixed,10,2) + " %";
        }
      else
        {
          effekt=0;
          LabelEFFEKT->Caption = FloatToStrF(effekt, ffFixed,10,2) + " %";
        }
    }
  else
    {
      EditKPE_OCEN->Enabled = true;
      rezult = 0;
      EditREZULT_OCEN->Text ="";
      LabelREZULT_PROC->Caption = FloatToStrF(rezult, ffFixed, 10,2) + " %";
      effekt = 0;
      LabelEFFEKT->Caption = FloatToStrF(effekt, ffFixed,10,2) + " %";
    }
  }

//---------------------------------------------------------------------------

void __fastcall TVvod::EditSTANDChange(TObject *Sender)
{
  IzmenKomp();
}
//---------------------------------------------------------------------------
void __fastcall TVvod::IzmenKomp()
{
  double komp_n;

  if (!EditSTAND->Text.IsEmpty() || !EditPOTREB->Text.IsEmpty() ||
      !EditKACH->Text.IsEmpty() || !EditEFF->Text.IsEmpty() ||
      !EditPROF_ZN->Text.IsEmpty() || !EditLIDER->Text.IsEmpty() ||
      !EditOTVETSTV->Text.IsEmpty() || !EditKOM_REZ->Text.IsEmpty())
    {
      //% по компетенции
      komp_n = (Main->SetNullF(EditSTAND->Text)+
                Main->SetNullF(EditPOTREB->Text)+
                Main->SetNullF(EditKACH->Text)+
                Main->SetNullF(EditEFF->Text)+
                Main->SetNullF(EditPROF_ZN->Text)+
                Main->SetNullF(EditLIDER->Text)+
                Main->SetNullF(EditOTVETSTV->Text)+
                Main->SetNullF(EditKOM_REZ->Text));

      komp = StrToFloat(FloatToStrF(komp_n/32*100, ffFixed, 10,2));

      EditKOMP_OCEN->Text = komp_n;//FloatToStrF(komp_n, ffFixed, 10,2);
      LabelKOMP_PROC->Caption = FloatToStrF(komp_n/32*100, ffFixed,10,2) + " %";

       //эффективность
      if (EditKPE_OCEN->Text.IsEmpty() && !EditREZULT_OCEN->Text.IsEmpty() && !EditKOMP_OCEN->Text.IsEmpty())
        {
          effekt =  ((((Main->SetNullF(EditREALIZAC->Text)+
                        Main->SetNullF(EditKACHESTVO->Text)+
                        Main->SetNullF(EditRESURS->Text))/12*100)*0.6)+
                    (((Main->SetNullF(EditSTAND->Text)+
                       Main->SetNullF(EditPOTREB->Text)+
                       Main->SetNullF(EditKACH->Text)+
                       Main->SetNullF(EditEFF->Text)+
                       Main->SetNullF(EditPROF_ZN->Text)+
                       Main->SetNullF(EditLIDER->Text)+
                       Main->SetNullF(EditOTVETSTV->Text)+
                       Main->SetNullF(EditKOM_REZ->Text))/32*100)*0.4)
                     );
          LabelEFFEKT->Caption = FloatToStrF(effekt, ffFixed,10,2) + " %";
        }
      else if (!EditKPE_OCEN->Text.IsEmpty() && !EditKOMP_OCEN->Text.IsEmpty())
        {
          effekt = ((StrToFloat(EditKPE_OCEN->Text)*0.6)+((Main->SetNullF(EditSTAND->Text)+
                                                     Main->SetNullF(EditPOTREB->Text)+
                                                     Main->SetNullF(EditKACH->Text)+
                                                     Main->SetNullF(EditEFF->Text)+
                                                     Main->SetNullF(EditPROF_ZN->Text)+
                                                     Main->SetNullF(EditLIDER->Text)+
                                                     Main->SetNullF(EditOTVETSTV->Text)+
                                                     Main->SetNullF(EditKOM_REZ->Text))/32*100)*0.4);
          LabelEFFEKT->Caption = FloatToStrF(effekt, ffFixed,10,2) + " %";
        }
      else
        {
          effekt = 0;
          LabelEFFEKT->Caption = FloatToStrF(effekt, ffFixed,10,2) + " %";
        }
    }
  else
    {
      komp = 0;
      EditKOMP_OCEN->Text ="";
      LabelKOMP_PROC->Caption = FloatToStrF(komp, ffFixed,10,2) + " %";
      effekt = 0;
      LabelEFFEKT->Caption = FloatToStrF(effekt, ffFixed,10,2) + " %";
    }
}
//---------------------------------------------------------------------------

