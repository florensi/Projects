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

 //Вывод отчетного года в DateTimePicker
 dt = TDateTime( "01.01." + IntToStr(Main->god));
 Data->DateTimePicker1->Date = dt;
}
//---------------------------------------------------------------------------

void __fastcall TData::btnViborClick(TObject *Sender)
{
  Word Year, Month, Day;

  //Считывание отчетного года из DateTimePicker
  DecodeDate(Data->DateTimePicker1->Date,Year, Month, Day);
  Main->god = Year;

  //Проверка на наличие данных за выбранный год в таблице SPGRAFIKI
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
      Application->MessageBox("Не возможно выбрать данные из таблицы SPGRAFIKI","Ошибка",
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
      Main->StatusBar1->SimpleText="Отчетный период:  "+IntToStr(Main->god)+" год";

      //Праздничные дни
      DM->qPrazdDni->Close();
      DM->qPrazdDni->Parameters->ParamByName("pgod")->Value = Main->god;
      try
        {
          DM->qPrazdDni->Open();
        }
      catch(...)
        {
          Application->MessageBox("Возникла ошибка при обращении к справочнику праздничных дней","Ошибка",
                                   MB_OK+MB_ICONERROR);
          Abort();
        }

      //Предпраздничные дни
      DM->qPrdPrazdDni->Close();
      DM->qPrdPrazdDni->Parameters->ParamByName("pgod")->Value = Main->god;
      try
        {
          DM->qPrdPrazdDni->Open();
        }
      catch(...)
        {
          Application->MessageBox("Возникла ошибка при обращении к справочнику праздничных дней","Ошибка",
                                   MB_OK+MB_ICONERROR);
          Abort();
        }

      //Определение даты перехода на летнее/зимнее время
      TDateTime data;
      Word year, month, day;

      // дата в марте
      data = DateToStr(EncodeDateMonthWeek(Main->god,3,4,6));
      DecodeDate(data, year, month, day);
      Main->day_mart = day;

      //для 40 и 90 графика, первой смены, дата в марте
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

      //дата в октябре
      data = DateToStr(EncodeDateMonthWeek(Main->god,10,4,6));
      DecodeDate(data, year, month, day);
      Main->day_oktyabr = day;

      //для 40 и 90 графика, первой смены, дата в октябре
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

      Application->MessageBox(("Отчетный период изменен!!!\nОтображение графиков на "+IntToStr(Main->god)+" год").c_str(),"Графики работы", MB_OK+MB_ICONINFORMATION);


      //Редактирование закрыто, если год прошедший
      if (Main->god < Main->grafr) Main->redakt=0;
      else Main->redakt=1;
    }
  else
    {
      Application->MessageBox("Нет данных за выбранный год","Предупреждение",
                              MB_OK + MB_ICONINFORMATION);
    }
}
//---------------------------------------------------------------------------

