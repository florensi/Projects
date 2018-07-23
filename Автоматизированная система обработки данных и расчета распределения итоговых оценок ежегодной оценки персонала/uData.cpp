//---------------------------------------------------------------------------

#include <vcl.h>
#pragma hdrstop

#include "uData.h"
#include "uMain.h"
#include "uDM.h"
#include "uReiting.h"
#include "uVvod.h"
#include "uZameshenie.h"
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


  Main->DBGridEh1->DataSource = NULL;
  Main->DBGridEh1->ClearFilter();
  Main->DBGridEh1->DataSource = DM->dsOcenka;
  DM->qOcenka->Close();
  DM->qOcenka->Parameters->ParamByName("pgod")->Value = Main->god;
  DM->qOcenka->Active=true;
  DM->qOcenka->Filtered = false;

  Main->StatusBar1->SimplePanel = true;
  Main->StatusBar1->SimpleText ="Отчетный период: "+IntToStr(Main->god)+" год";

  //Проверка на наличие данных за выбранный год в таблице
  if (DM->qOcenka->RecordCount==0)
    {
      Application->MessageBox("Нет данных за выбранный год","Предупреждение",
                              MB_OK + MB_ICONINFORMATION);
    }
  else
    {
      Application->MessageBox(("Отчетный период изменен на "+IntToStr(Main->god)+" год").c_str(),"Предупреждение",
                              MB_OK + MB_ICONINFORMATION);
    }


  //Год для замещений
  DM->qZamesh->Close();
  DM->qZamesh->Parameters->ParamByName("pgod")->Value=IntToStr(Main->god);
  DM->qZamesh->Active = true;


  //Блокирование загрузки и редактирования даннных, если выполняется просмотр данных за предыдущие года
  if (Main->god<Main->god_t)
    {
      //Главное меню
      Main->N1->Visible = false; //Скрытие пункта меню "Загрузка данных"
      Main->N10->Visible = false; //Скрытие пункта меню "Расчет рейтинга"
      Main->N5->Visible = false; //Скрытие пункта меню "Ежегодная единоразовая загрузка списка работников"

      //Главная форма
      Main->SpeedButton4->Enabled = false;  //Видимость кнопки "Загрузка данных" на главной форме

      //Контекстное меню
      Main->N18->Visible = false;  //Видимость пункта контекстного меню "Рассчитать рейтинг"

      //Другие формы
      Vvod->Button1->Enabled = false;    //Кнопка "Сохранить" на форме редактирования данных по работнику
      Reiting->Button1->Enabled = false; //Кнопка "Сохранить" на форме редактирования рейтинга
      Zameshenie->BitBtn1->Enabled = false;       //Кнопка "Сохранить" на форме редактирования замещения

      //Видимость новых полей на формах редактирования
      Vvod->EditREZULT_OCEN->Enabled = false;
      Vvod->EditREALIZAC->Enabled = false;
      Vvod->EditKACHESTVO->Enabled = false;
      Vvod->EditRESURS->Enabled = false;
      Vvod->EditSTAND->Enabled = false;
      Vvod->EditPOTREB->Enabled = false;
      Vvod->EditKACH->Enabled = false;
      Vvod->EditEFF->Enabled = false;
      Vvod->EditPROF_ZN->Enabled = false;
      Vvod->EditLIDER->Enabled = false;
      Vvod->EditOTVETSTV->Enabled = false;
      Vvod->EditKOM_REZ->Enabled = false;

      Reiting->EditREALIZAC->Enabled = false;
      Reiting->EditKACHESTVO->Enabled = false;
      Reiting->EditRESURS->Enabled = false;
      Reiting->EditSTAND->Enabled = false;
      Reiting->EditPOTREB->Enabled = false;
      Reiting->EditKACH->Enabled = false;
      Reiting->EditEFF->Enabled = false;
      Reiting->EditPROF_ZN->Enabled = false;
      Reiting->EditLIDER->Enabled = false;
      Reiting->EditOTVETSTV->Enabled = false;
      Reiting->EditKOM_REZ->Enabled = false;

    }
  else
    {
      //Главное меню
      Main->N1->Visible = true; //Скрытие пункта меню "Загрузка данных"
      Main->N10->Visible = true; //Скрытие пункта меню "Расчет рейтинга"
      Main->N5->Visible = true; //Скрытие пункта меню "Ежегодная единоразовая загрузка списка работников"

      //Главная форма
      Main->SpeedButton4->Enabled = true;  //Видимость кнопки "Загрузка данных" на главной форме

      //Контекстное меню
      Main->N18->Visible = true;  //Видимость пункта контекстного меню "Рассчитать рейтинг"

      //Другие формы
      Vvod->Button1->Enabled = true;    //Кнопка "Сохранить" на форме редактирования данных по работнику
      Reiting->Button1->Enabled = true; //Кнопка "Сохранить" на форме редактирования рейтинга
      Zameshenie->BitBtn1->Enabled = true;       //Кнопка "Сохранить" на форме редактирования замещения

      //Видимость новых полей на формах редактирования
      Vvod->EditREZULT_OCEN->Enabled = true;
      Vvod->EditREALIZAC->Enabled = true;
      Vvod->EditKACHESTVO->Enabled = true;
      Vvod->EditRESURS->Enabled = true;
      Vvod->EditSTAND->Enabled = true;
      Vvod->EditPOTREB->Enabled = true;
      Vvod->EditKACH->Enabled = true;
      Vvod->EditEFF->Enabled = true;
      Vvod->EditPROF_ZN->Enabled = true;
      Vvod->EditLIDER->Enabled = true;
      Vvod->EditOTVETSTV->Enabled = true;
      Vvod->EditKOM_REZ->Enabled = true;

      Reiting->EditREALIZAC->Enabled = true;
      Reiting->EditKACHESTVO->Enabled = true;
      Reiting->EditRESURS->Enabled = true;
      Reiting->EditSTAND->Enabled = true;
      Reiting->EditPOTREB->Enabled = true;
      Reiting->EditKACH->Enabled = true;
      Reiting->EditEFF->Enabled = true;
      Reiting->EditPROF_ZN->Enabled = true;
      Reiting->EditLIDER->Enabled = true;
      Reiting->EditOTVETSTV->Enabled = true;
      Reiting->EditKOM_REZ->Enabled = true;

    }
}
//---------------------------------------------------------------------------

