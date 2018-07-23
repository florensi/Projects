//---------------------------------------------------------------------------

#include <vcl.h>
#pragma hdrstop

#include "uReiting.h"
#include "uDM.h"
#include "uMain.h"
//---------------------------------------------------------------------------
#pragma package(smart_init)
#pragma resource "*.dfm"
TReiting *Reiting;
//---------------------------------------------------------------------------
__fastcall TReiting::TReiting(TComponent* Owner)
        : TForm(Owner)
{
}
//---------------------------------------------------------------------------
void __fastcall TReiting::SpeedButton2Click(TObject *Sender)
{
  Reiting->Close();        
}
//---------------------------------------------------------------------------

void __fastcall TReiting::SpeedButton1Click(TObject *Sender)
{
//проверка полей        
}
//---------------------------------------------------------------------------
void __fastcall TReiting::FormKeyDown(TObject *Sender, WORD &Key,
      TShiftState Shift)
{
  if (Key==VK_RETURN)
  FindNextControl((TWinControl *)Sender, true, true, false)->SetFocus();
}
//---------------------------------------------------------------------------


//Сохранить
void __fastcall TReiting::Button1Click(TObject *Sender)
{
  AnsiString Sql, seffekt, srezult, skomp;

  if(effekt==0) seffekt="NULL";
  else seffekt=effekt;

  if(rezult==0) srezult="NULL";
  else srezult=rezult;

  if(komp==0) skomp="NULL";
  else skomp=komp;

  //Проверки
  if (LabelZEX->Caption=="" || (EditTN->Text.IsEmpty() || EditTN->Text=="490"))
    {
      Application->MessageBox("Не указан цех или таб.№","Предупреждение",
                              MB_OK+MB_ICONINFORMATION);
      EditTN->SetFocus();
      Abort();
    }


  //Введено все по результатам работ
  if (EditKE->Text.IsEmpty())
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


  //Проверка на изменение данных
  if (zrez!=EditREZ->Text ||
      zkom!=EditKOM->Text ||
      zke!=EditKE->Text ||
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
      zkom_rez!=EditKOM_REZ->Text
      )
    {
      //Обновление
      Sql = "update ocenka set\
                           rezult_ocen = "+Main->SetNull(EditREZ->Text)+",  \
                           kpe_ocen = "+Main->SetNull(EditKE->Text)+",  \
                           komp_ocen = "+Main->SetNull(EditKOM->Text)+", \
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
                           efect = "+seffekt+" \
             where zex="+LabelZEX->Caption+" and tn="+EditTN->Text+" and god="+IntToStr(Main->god);

      DM->qObnovlenie->Close();
      DM->qObnovlenie->SQL->Clear();
      DM->qObnovlenie->SQL->Add(Sql);
      try
        {
          DM->qObnovlenie->ExecSQL();
        }
      catch(Exception &E)
        {
          Application->MessageBox(("Возникла ошибка при обновлении рейтинга в таблице OCENKA" + E.Message).c_str(),"Ошибка",
                                   MB_OK+MB_ICONERROR);
          Abort();
        }
        
      //Логи
      if (DM->qObnovlenie->RowsAffected>0)
        {
          AnsiString Str ="Обновление рейтинга по работнику: цех="+LabelZEX->Caption+" таб.№="+EditTN->Text;

          if (zrez!=EditREZ->Text) Str+=" результ.работы с '"+zrez+"' на '"+EditREZ->Text+"'";
          if (zkom!=EditKOM->Text) Str+=", компетенции с '"+zkom+"' на '"+EditKOM->Text+"'";
          if (zke!=EditKE->Text) Str+=", результ. по КЕ с '"+zke+"' на '"+EditKE->Text+"'";
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
          Main->InsertLog("Обновление рейтинга за "+IntToStr(Main->god)+" год по работнику: цех="+LabelZEX->Caption+" таб.№="+EditTN->Text+" не выполнено");
          DM->qLogs->Requery();
        }

      DM->qOcenka->Requery();
    }

  //Очищение Edit-ов
  LabelZEX->Caption="";
  LabelNZEX->Caption="";
  EditTN->Text = "490";


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

  EditKE->Text = "";
  EditREZ->Text = "";
  EditKOM->Text = "";

  Label1->Caption = "";

  LabelREZ->Caption = "";
  LabelKOMP->Caption = "";
  LabelEFFEKT->Caption = "";
  LabelFIO_OCEN->Caption = "";

  EditTN->SetFocus();
  EditTN->SelStart=EditTN->Text.Length();

}
//---------------------------------------------------------------------------

void __fastcall TReiting::CanselClick(TObject *Sender)
{
  Reiting->Close();
}
//---------------------------------------------------------------------------


void __fastcall TReiting::FormShow(TObject *Sender)
{
  LabelZEX->Caption = "";
  LabelNZEX->Caption = "";
  EditTN->Text = "490";


  Label1->Caption = "";

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


  EditREZ->Text = "";
  EditKE->Text = "";
  EditKOM->Text = "";

  LabelREZ->Caption = "";
  LabelKOMP->Caption = "";
  LabelEFFEKT->Caption = "";
  LabelFIO_OCEN->Caption = "";

  EditTN->SetFocus();
  EditTN->SelStart=EditTN->Text.Length();
}
//---------------------------------------------------------------------------

void __fastcall TReiting::EditREZExit(TObject *Sender)
{
 /* if (ActiveControl == Cansel)
    {
      Reiting->Close();
    }
  else
    {
      if (!EditREZ->Text.IsEmpty())
        {
          EditKE->Enabled = false;

          if (StrToFloat(EditREZ->Text)>4)
            {
              Application->MessageBox("Максимальное значение по результатам \nработы не может превышать 4","Предупреждение", MB_OK+MB_ICONWARNING);

              EditREZ->SetFocus();
            }
        }
      else
        {
          EditKE->Enabled = true;
        }
    }  */
}
//---------------------------------------------------------------------------

void __fastcall TReiting::EditKEKeyPress(TObject *Sender, char &Key)
{
  if (! (IsNumeric(Key) || Key=='.' || Key==',' || Key=='/' || Key=='\b') ) Key=0;
  if (Key==',' || Key=='/') Key='.';
}
//---------------------------------------------------------------------------

void __fastcall TReiting::EditZEXKeyPress(TObject *Sender, char &Key)
{
  if (!(IsNumeric(Key)||Key=='\b')) Key=0;        
}
//---------------------------------------------------------------------------

void __fastcall TReiting::EditTNChange(TObject *Sender)
{
  double rezult_n, komp_n;
  
  if (!EditTN->Text.IsEmpty() && EditTN->Text!="490")
    {
      EditREZ->Enabled = false;
      EditKOM->Enabled = false;

      if (Main->god<Main->god_t)
        {
          
          EditKE->Enabled = false;

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
          EditKE->Enabled = true;

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

      if (!EditTN->Text.IsEmpty())
        {
          AnsiString Sql = "select initcap(fio) as fio, initcap(fio_ocen) as fio_ocen, rezult_ocen, kpe_ocen, komp_ocen,\
                                   realizac, kachestvo, resurs, stand, potreb, kach, eff, prof_zn, lider, otvetstv, kom_rez, \
                                   zex,(select distinct nazv_cexk from ssap_cex where id_cex=zex and nazv_cexk not like '%устар.%') knaim_zex   \
                            from ocenka \
                            where tn="+EditTN->Text+"  and god="+IntToStr(Main->god);

          DM->qObnovlenie->Close();
          DM->qObnovlenie->SQL->Clear();
          DM->qObnovlenie->SQL->Add(Sql);
          DM->qObnovlenie->Open();

          LabelZEX->Caption=DM->qObnovlenie->FieldByName("zex")->AsString;
          LabelNZEX->Caption=DM->qObnovlenie->FieldByName("knaim_zex")->AsString;
          Label1->Caption = DM->qObnovlenie->FieldByName("fio")->AsString;
          LabelFIO_OCEN->Caption = DM->qObnovlenie->FieldByName("fio_ocen")->AsString;
          EditKE->Text = zke = DM->qObnovlenie->FieldByName("kpe_ocen")->AsString;

          EditREALIZAC->Text = zrealizac = DM->qObnovlenie->FieldByName("realizac")->AsString;
          EditKACHESTVO->Text = zkachestvo = DM->qObnovlenie->FieldByName("kachestvo")->AsString;
          EditRESURS->Text = zresurs = DM->qObnovlenie->FieldByName("resurs")->AsString;
          EditSTAND->Text = zstand = DM->qObnovlenie->FieldByName("stand")->AsString;
          EditPOTREB->Text = zpotreb = DM->qObnovlenie->FieldByName("potreb")->AsString;
          EditKACH->Text = zkach = DM->qObnovlenie->FieldByName("kach")->AsString;
          EditEFF->Text = zeff = DM->qObnovlenie->FieldByName("eff")->AsString;
          EditPROF_ZN->Text = zprof_zn = DM->qObnovlenie->FieldByName("prof_zn")->AsString;
          EditLIDER->Text = zlider = DM->qObnovlenie->FieldByName("lider")->AsString;
          EditOTVETSTV->Text = zotvetstv = DM->qObnovlenie->FieldByName("otvetstv")->AsString;
          EditKOM_REZ->Text = zkom_rez = DM->qObnovlenie->FieldByName("kom_rez")->AsString;

          if (Main->god<Main->god_t)
            {
              EditREZ->Text = zrez = DM->qObnovlenie->FieldByName("rezult_ocen")->AsString;
              EditKOM->Text = zkom = DM->qObnovlenie->FieldByName("komp_ocen")->AsString;

              //% по результатам работы
              if (!EditREZ->Text.IsEmpty()) LabelREZ->Caption = FloatToStrF(DM->qOcenka->FieldByName("rezult_proc")->AsFloat, ffFixed, 10,2) + " %";
              else LabelREZ->Caption ="";

              //% по компетенции
              if (!EditKOM->Text.IsEmpty()) LabelKOMP->Caption = FloatToStrF(DM->qOcenka->FieldByName("komp_proc")->AsFloat, ffFixed,10,2) + " %";
              else LabelKOMP->Caption = "";

              //эффективность
              if ((!EditKE->Text.IsEmpty() && !EditKOM->Text.IsEmpty()) || (!EditREZ->Text.IsEmpty() && !EditKOM->Text.IsEmpty()))
                {
                  effekt = DM->qOcenka->FieldByName("efect")->AsFloat;
                }
              else  effekt = 0;

              LabelEFFEKT->Caption = FloatToStrF(effekt, ffFixed,10,2) + " %";
            }
          else
            {
               rezult_n = (DM->qObnovlenie->FieldByName("realizac")->AsFloat+
                           DM->qObnovlenie->FieldByName("kachestvo")->AsFloat+
                           DM->qObnovlenie->FieldByName("resurs")->AsFloat);

               komp_n = (DM->qObnovlenie->FieldByName("stand")->AsFloat+
                         DM->qObnovlenie->FieldByName("potreb")->AsFloat+
                         DM->qObnovlenie->FieldByName("kach")->AsFloat+
                         DM->qObnovlenie->FieldByName("eff")->AsFloat+
                         DM->qObnovlenie->FieldByName("prof_zn")->AsFloat+
                         DM->qObnovlenie->FieldByName("lider")->AsFloat+
                         DM->qObnovlenie->FieldByName("otvetstv")->AsFloat+
                         DM->qObnovlenie->FieldByName("kom_rez")->AsFloat);


              //расчет эффективности изменился с 2017 года
              //результаты работы
              if (rezult_n/3==0)
                {
                  EditREZ->Text = zrez = "";
                  LabelREZ->Caption ="";
                }
              else
                {
                  EditREZ->Text =  FloatToStrF(rezult_n/3, ffFixed, 10,2);

                  //% по результатам работы
                  rezult = rezult_n/12*100;
                  LabelREZ->Caption = FloatToStrF(rezult, ffFixed, 10,2) + " %";
                }

              if (komp_n/8==0)
                {
                  EditKOM->Text = zkom = "";
                  LabelKOMP->Caption = "";
                }
              else
                {
                  EditKOM->Text = zkom = komp_n;


                  //% по компетенции
                  komp = komp_n/32*100;
                  LabelKOMP->Caption = FloatToStrF(komp, ffFixed,10,2) + " %";
                }

              //эффективность
              if (EditKE->Text.IsEmpty() && !EditREZ->Text.IsEmpty() && !EditKOM->Text.IsEmpty())
                {
                  effekt = (((rezult_n*0.6)/12*100)+((komp_n*0.4)/32*100));
                }
              else if (!EditKE->Text.IsEmpty() && !EditKOM->Text.IsEmpty())
                {
                  effekt = ((DM->qObnovlenie->FieldByName("kpe_ocen")->AsFloat*0.6)+((komp_n*0.4)/32*100));
                }
              else  effekt = 0;
              LabelEFFEKT->Caption = FloatToStrF(effekt, ffFixed,10,2) + " %";
            }

        }
    }
}
//---------------------------------------------------------------------------

void __fastcall TReiting::EditKEExit(TObject *Sender)
{
  if (!EditKE->Text.IsEmpty())
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


void __fastcall TReiting::EditREZChange(TObject *Sender)
{
  if (Main->god<2017)
    {
      if (!EditREZ->Text.IsEmpty())
        {
          EditKE->Enabled = false;

          //% по результатам работы
          rezult = StrToFloat(EditREZ->Text)/4*100;
          LabelREZ->Caption = FloatToStrF(rezult, ffFixed, 10,2) + " %";

          if (!EditKOM->Text.IsEmpty())
            {
              effekt = (((StrToFloat(EditREZ->Text)*0.6)/4*100)+((StrToFloat(EditKOM->Text)*0.4)/32*100));
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
          EditKE->Enabled = true;
          rezult = 0;
          LabelREZ->Caption = FloatToStrF(rezult, ffFixed, 10,2) + " %";
          effekt = 0;
          LabelEFFEKT->Caption = FloatToStrF(effekt, ffFixed,10,2) + " %";
        }
    }
}
//---------------------------------------------------------------------------

void __fastcall TReiting::EditKOMChange(TObject *Sender)
{
  if (Main->god<2017)
    {
      if (!EditKOM->Text.IsEmpty())
        {
          //% по компетенции
          komp = StrToFloat(EditKOM->Text)/32*100;
          LabelKOMP->Caption = FloatToStrF(komp, ffFixed,10,2) + " %";

          //эффективность
          if (EditKE->Text.IsEmpty() && !EditREZ->Text.IsEmpty() && !EditKOM->Text.IsEmpty())
            {
              effekt = (((StrToFloat(EditREZ->Text)*0.6)/4*100)+((StrToFloat(EditKOM->Text)*0.4)/32*100));
              LabelEFFEKT->Caption = FloatToStrF(effekt, ffFixed,10,2) + " %";
            }
          else if (!EditKE->Text.IsEmpty() && !EditKOM->Text.IsEmpty())
            {
              effekt = ((StrToFloat(EditKE->Text)*0.6)+((StrToFloat(EditKOM->Text)*0.4)/32*100));
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
          LabelKOMP->Caption = FloatToStrF(komp, ffFixed,10,2) + " %";
          effekt = 0;
          LabelEFFEKT->Caption = FloatToStrF(effekt, ffFixed,10,2) + " %";
        }
     }
}
//---------------------------------------------------------------------------

void __fastcall TReiting::EditKEChange(TObject *Sender)
{
 if (!EditKE->Text.IsEmpty())
   {
      if (!EditKOM->Text.IsEmpty())
        {
          effekt = ((StrToFloat(EditKE->Text)*0.6)+((Main->SetNullF(EditSTAND->Text)+
                                                     Main->SetNullF(EditPOTREB->Text)+
                                                     Main->SetNullF(EditKACH->Text)+
                                                     Main->SetNullF(EditEFF->Text)+
                                                     Main->SetNullF(EditPROF_ZN->Text)+
                                                     Main->SetNullF(EditLIDER->Text)+
                                                     Main->SetNullF(EditOTVETSTV->Text)+
                                                     Main->SetNullF(EditKOM_REZ->Text))/32*100)*0.4);
          LabelEFFEKT->Caption = FloatToStrF(effekt, ffFixed,10,2) + " %";
          EditREZ->Text="";
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
//---------------------------------------------------------------------------

void __fastcall TReiting::EditKOMExit(TObject *Sender)
{
/* if (ActiveControl == Cansel)
    {
      Reiting->Close();
    }
  else
    {
      if (!EditKOM->Text.IsEmpty())
        {
          if (StrToFloat(EditKOM->Text)>32)
            {
              Application->MessageBox("Максимальное значение по компетенции \nне может превышать 32","Предупреждение", MB_OK+MB_ICONWARNING);

              EditKOM->SetFocus();
            }
        }
    }  */
}
//---------------------------------------------------------------------------


void __fastcall TReiting::EditZEXKeyDown(TObject *Sender, WORD &Key,
      TShiftState Shift)
{
  if (Key==VK_RETURN)
  FindNextControl((TWinControl *)Sender, true, true, false)->SetFocus();
  EditTN->SelStart=EditTN->Text.Length();        
}
//---------------------------------------------------------------------------





void __fastcall TReiting::EditREALIZACExit(TObject *Sender)
{
 if (ActiveControl == Cansel)
    {
      Reiting->Close();
    }
  else
    {
      if (!EditREALIZAC->Text.IsEmpty())
        {
          EditKE->Enabled = false;

          if (StrToFloat(EditREALIZAC->Text)>4)
            {
              Application->MessageBox("Максимальное значение по результатам \nработы не может превышать 4","Предупреждение", MB_OK+MB_ICONWARNING);

              EditREALIZAC->SetFocus();
            }
        }
      else if (!EditKACHESTVO->Text.IsEmpty() || !EditRESURS->Text.IsEmpty())
        {
          EditKE->Enabled = false;
        }
      else
        {
          EditKE->Enabled = true;
        }
    }
}
//---------------------------------------------------------------------------

void __fastcall TReiting::EditKACHESTVOExit(TObject *Sender)
{
 if (ActiveControl == Cansel)
    {
      Reiting->Close();
    }
  else
    {
      if (!EditKACHESTVO->Text.IsEmpty())
        {
          EditKE->Enabled = false;

          if (StrToFloat(EditKACHESTVO->Text)>4)
            {
              Application->MessageBox("Максимальное значение по результатам \nработы не может превышать 4","Предупреждение", MB_OK+MB_ICONWARNING);

              EditKACHESTVO->SetFocus();
            }
        }
      else if (!EditREALIZAC->Text.IsEmpty() || !EditRESURS->Text.IsEmpty())
        {
          EditKE->Enabled = false;
        }
      else
        {
          EditKE->Enabled = true;
        }
    }
}
//---------------------------------------------------------------------------

void __fastcall TReiting::EditRESURSExit(TObject *Sender)
{
 if (ActiveControl == Cansel)
    {
      Reiting->Close();
    }
  else
    {
      if (!EditRESURS->Text.IsEmpty())
        {
          EditKE->Enabled = false;

          if (StrToFloat(EditRESURS->Text)>4)
            {
              Application->MessageBox("Максимальное значение по результатам \nработы не может превышать 4","Предупреждение", MB_OK+MB_ICONWARNING);

              EditRESURS->SetFocus();
            }
        }
      else if (!EditKACHESTVO->Text.IsEmpty() || !EditREALIZAC->Text.IsEmpty())
        {
          EditKE->Enabled = false;
        }
      else
        {
          EditKE->Enabled = true;
        }
    }
}
//---------------------------------------------------------------------------

void __fastcall TReiting::EditREALIZACChange(TObject *Sender)
{
  IzmenRezRab();

    if (EditREALIZAC->Text.Length()<1) EditREALIZAC->SetFocus();
    else if (EditKACHESTVO->Text.Length()<1) EditKACHESTVO->SetFocus();
    else if (EditRESURS->Text.Length()<1) EditRESURS->SetFocus();
    else if (EditEFF->Text.Length()<1) EditEFF->SetFocus();
    else if (EditPROF_ZN->Text.Length()<1) EditPROF_ZN->SetFocus();
    else if (EditLIDER->Text.Length()<1) EditLIDER->SetFocus();
    else if (EditOTVETSTV->Text.Length()<1) EditOTVETSTV->SetFocus();
    else if (EditKOM_REZ->Text.Length()<1) EditKOM_REZ->SetFocus();
    else if (EditSTAND->Text.Length()<1) EditSTAND->SetFocus();
    else if (EditPOTREB->Text.Length()<1) EditPOTREB->SetFocus();
    else if (EditKACH->Text.Length()<1) EditKACH->SetFocus();
    else FindNextControl((TWinControl *)Sender, true, true, false)->SetFocus();
}
//---------------------------------------------------------------------------

void __fastcall TReiting::IzmenRezRab()
{
  double rezult_n;

  if (!EditREALIZAC->Text.IsEmpty() || !EditKACHESTVO->Text.IsEmpty() || !EditRESURS->Text.IsEmpty())
    {
      EditKE->Enabled = false;

      //% по результатам работы
      rezult_n = (Main->SetNullF(EditREALIZAC->Text)+
                  Main->SetNullF(EditKACHESTVO->Text)+
                  Main->SetNullF(EditRESURS->Text));

      rezult = StrToFloat(FloatToStrF(rezult_n/12*100, ffFixed, 10,2));

      EditREZ->Text =  FloatToStrF(rezult_n/3, ffFixed, 10,2);
      LabelREZ->Caption = FloatToStrF(rezult_n/12*100, ffFixed, 10,2) + " %";


      if (!EditKOM->Text.IsEmpty())
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
      EditKE->Enabled = true;
      rezult = 0;
      EditREZ->Text ="";
      LabelREZ->Caption = FloatToStrF(rezult, ffFixed, 10,2) + " %";
      effekt = 0;
      LabelEFFEKT->Caption = FloatToStrF(effekt, ffFixed,10,2) + " %";
    }
  }


//---------------------------------------------------------------------------


void __fastcall TReiting::EditSTANDChange(TObject *Sender)
{
  IzmenKomp();

  if (EditEFF->Text.Length()<1) EditEFF->SetFocus();
  else if (EditPROF_ZN->Text.Length()<1) EditPROF_ZN->SetFocus();
  else if (EditLIDER->Text.Length()<1) EditLIDER->SetFocus();
  else if (EditOTVETSTV->Text.Length()<1) EditOTVETSTV->SetFocus();
  else if (EditKOM_REZ->Text.Length()<1) EditKOM_REZ->SetFocus();
  else if (EditSTAND->Text.Length()<1) EditSTAND->SetFocus();
  else if (EditPOTREB->Text.Length()<1) EditPOTREB->SetFocus();
  else if (EditKACH->Text.Length()<1) EditKACH->SetFocus();
  else FindNextControl((TWinControl *)Sender, true, true, false)->SetFocus();
}
//---------------------------------------------------------------------------

void __fastcall TReiting::IzmenKomp()
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

      EditKOM->Text = komp_n;//FloatToStrF(komp_n, ffFixed, 10,2);
      LabelKOMP->Caption = FloatToStrF(komp_n/32*100, ffFixed,10,2) + " %";

       //эффективность
      if (EditKE->Text.IsEmpty() && !EditREZ->Text.IsEmpty() && !EditKOM->Text.IsEmpty())
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
      else if (!EditKE->Text.IsEmpty() && !EditKOM->Text.IsEmpty())
        {
          effekt = ((StrToFloat(EditKE->Text)*0.6)+((Main->SetNullF(EditSTAND->Text)+
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
      EditKOM->Text ="";
      LabelKOMP->Caption = FloatToStrF(komp, ffFixed,10,2) + " %";
      effekt = 0;
      LabelEFFEKT->Caption = FloatToStrF(effekt, ffFixed,10,2) + " %";
    }
}
//---------------------------------------------------------------------------





void __fastcall TReiting::EditSTANDExit(TObject *Sender)
{
 if (ActiveControl == Cansel)
    {
      Reiting->Close();
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

void __fastcall TReiting::EditPOTREBExit(TObject *Sender)
{
 if (ActiveControl == Cansel)
    {
      Reiting->Close();
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

void __fastcall TReiting::EditKACHExit(TObject *Sender)
{
 if (ActiveControl == Cansel)
    {
      Reiting->Close();
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

void __fastcall TReiting::EditEFFExit(TObject *Sender)
{
 if (ActiveControl == Cansel)
    {
      Reiting->Close();
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

void __fastcall TReiting::EditPROF_ZNExit(TObject *Sender)
{
 if (ActiveControl == Cansel)
    {
      Reiting->Close();
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

void __fastcall TReiting::EditLIDERExit(TObject *Sender)
{
 if (ActiveControl == Cansel)
    {
      Reiting->Close();
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

void __fastcall TReiting::EditOTVETSTVExit(TObject *Sender)
{
  if (ActiveControl == Cansel)
    {
      Reiting->Close();
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

void __fastcall TReiting::EditKOM_REZExit(TObject *Sender)
{
  if (ActiveControl == Cansel)
    {
      Reiting->Close();
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



void __fastcall TReiting::EditKACHESTVOChange(TObject *Sender)
{
 /* IzmenRezRab();

  if (EditKACHESTVO->Text.Length()>0){
    if (EditREALIZAC->Text.Length()<1) EditREALIZAC->SetFocus();
    else if (EditRESURS->Text.Length()<1) EditRESURS->SetFocus();
    else if (EditEFF->Text.Length()<1) EditEFF->SetFocus();
    else if (EditPROF_ZN->Text.Length()<1) EditPROF_ZN->SetFocus();
    else if (EditLIDER->Text.Length()<1) EditLIDER->SetFocus();
    else if (EditOTVETSTV->Text.Length()<1) EditOTVETSTV->SetFocus();
    else if (EditKOM_REZ->Text.Length()<1) EditKOM_REZ->SetFocus();
    else if (EditSTAND->Text.Length()<1) EditSTAND->SetFocus();
    else if (EditPOTREB->Text.Length()<1) EditPOTREB->SetFocus();
    else if (EditKACH->Text.Length()<1) EditKACH->SetFocus();
    else FindNextControl((TWinControl *)Sender, true, true, false)->SetFocus();
  }  */      
}
//---------------------------------------------------------------------------

