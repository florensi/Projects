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
  // Получение данных о пользователе из домена
  TStringList *SL_Groups = new TStringList();
  // TStringList *SL_Groups2 = new TStringList();

  // Получение данных о пользователе из домена
  // Переменные UserName, DomainName, UserFullName должны быть объявлены как AnsiString
  if (!GetFullUserInfo(UserName, DomainName, UserFullName))
	{
	  MessageBox(Handle,L"Ошибка получения данных о пользователе",L"Ошибка",8208);
	  Application->Terminate();
	  Abort();
	}

  //получение групп доступа из АД
  if (!GetUserGroups(UserName, DomainName, SL_Groups))
	{
	  MessageBox(Handle,L"Ошибка получения данных о пользователе",L"Ошибка",8208);
	  Application->Terminate();
	  Abort();
	}

  //проверка на доступ к услуге
  if ((SL_Groups->IndexOf("mmk-itsvc-hocn-admin")<=-1) && (SL_Groups->IndexOf("mmk-itsvc-hocn")<=-1))
	{
	  MessageBox(Handle,L"У вас нет прав для работы с\nпрограммой 'Рейтингование руководителей'!!!",L"Права доступа",8208);
	  Application->Terminate();
	  Abort();
	}
  Prava = "ocen";

/*  //проверка прав
  //если права группы Оценки персонала
  if (SL_Groups->IndexOf("mmk-itsvc-hrru-ocen")>-1)
	{
	  DBGridEh1->Columns->Items[6]->Visible = true;     //Наличие подчиненных
	  DBGridEh1->Columns->Items[7]->Visible = true;     //Оценка
	  DBGridEh1->Columns->Items[9]->Visible = true;     //Производственное задание, факт
	  DBGridEh1->Columns->Items[10]->Visible = true;    //Производственное задание, балл
	  DBGridEh1->Columns->Items[11]->Visible = true;    //КПЭ, факт
	  DBGridEh1->Columns->Items[12]->Visible = true;    //КПЭ, балл
	  DBGridEh1->Columns->Items[13]->Visible = true;    //Отклонения, достижение
	  DBGridEh1->Columns->Items[14]->Visible = true;    //Отклонения, балл
	  DBGridEh1->Columns->Items[15]->Visible = true;    //Преемники, достижение
	  DBGridEh1->Columns->Items[16]->Visible = true;    //Преемники, балл
	  DBGridEh1->Columns->Items[17]->Visible = true;    //Информирование, уровень
	  DBGridEh1->Columns->Items[18]->Visible = true;    //Информирование, балл
	  DBGridEh1->Columns->Items[19]->Visible = true;    //5С, уровень
	  DBGridEh1->Columns->Items[20]->Visible = true;    //5С, балл
	  DBGridEh1->Columns->Items[21]->Visible = true;    //КНС, достижение
	  DBGridEh1->Columns->Items[22]->Visible = true;    //КНС, балл
	  DBGridEh1->Columns->Items[23]->Visible = true;    //СПП, кол-во поданых
	  DBGridEh1->Columns->Items[24]->Visible = true;    //СПП, балл
	  DBGridEh1->Columns->Items[25]->Visible = true;    //ОТ, ПБ и Э
	  DBGridEh1->Columns->Items[26]->Visible = true;    //Нарушения ОТ
	  DBGridEh1->Columns->Items[27]->Visible = true;    //Трудовая дисциплина

	  //Главное меню
	  N6Spisok->Visible = true;     //Загрузка списка работников для рейтингования
	  N7PP->Visible = true;         //Загрузка производственного задания
	  N8KPE->Visible = true;        //Загрузка КПЭ
	  N51C5->Visible = true;        //Загрузка 5 Почему
	  N10PREEM->Visible = true;     //Загрузка преемников
	  N11SPP->Visible = true;       //Загрузка СПП
	  N12OT->Visible = true;        //Загрузка ОТ
	  N13NOT->Visible = true;       //Загрузка нарушений по ОТ
	  N14TD->Visible = true;        //Загрузка трудовой дисциплины
	  N3Otchet->Visible = true;     //Формирование итогового отчета
	  NSprav->Visible = true;        //Справочник по производственному заданию

	  //Контекстное меню
	  N16Dobav->Visible = true;     //Добавление записи
	  N17Redact->Visible = true;    //Редактирование записи
	  N6Reiting->Visible = true;    //Расчет рейтинга

	  //Кнопки
	  SpeedButton1->Visible = true; //Загрузка списка работников для рейтингования
	  SpeedButton2->Visible = true; //Редактирование записи
	  SpeedButton3->Visible = true; //Расчет рейтинга


	  Prava = "ocen";
	}
  else
	{

	  DBGridEh1->Columns->Items[6]->Visible = false;     //Наличие подчиненных
	  DBGridEh1->Columns->Items[7]->Visible = false;     //Оценка
	  DBGridEh1->Columns->Items[9]->Visible = false;     //Производственное задание, факт
	  DBGridEh1->Columns->Items[10]->Visible = false;    //Производственное задание, балл
	  DBGridEh1->Columns->Items[11]->Visible = false;    //КПЭ, факт
	  DBGridEh1->Columns->Items[12]->Visible = false;    //КПЭ, балл
	  DBGridEh1->Columns->Items[13]->Visible = false;    //Отклонения, достижение
	  DBGridEh1->Columns->Items[14]->Visible = false;    //Отклонения, балл
	  DBGridEh1->Columns->Items[15]->Visible = false;    //Преемники, достижение
	  DBGridEh1->Columns->Items[16]->Visible = false;    //Преемники, балл
	  DBGridEh1->Columns->Items[17]->Visible = false;    //Информирование, уровень
	  DBGridEh1->Columns->Items[18]->Visible = false;    //Информирование, балл
	  DBGridEh1->Columns->Items[19]->Visible = false;    //5С, уровень
	  DBGridEh1->Columns->Items[20]->Visible = false;    //5С, балл
	  DBGridEh1->Columns->Items[21]->Visible = false;    //КНС, достижение
	  DBGridEh1->Columns->Items[22]->Visible = false;    //КНС, балл
	  DBGridEh1->Columns->Items[23]->Visible = false;    //СПП, кол-во поданых
	  DBGridEh1->Columns->Items[24]->Visible = false;    //СПП, балл
	  DBGridEh1->Columns->Items[25]->Visible = false;    //ОТ, ПБ и Э
	  DBGridEh1->Columns->Items[26]->Visible = false;    //Нарушения ОТ
	  DBGridEh1->Columns->Items[27]->Visible = false;    //Трудовая дисциплина

	  //Главное меню
	  N6Spisok->Visible = false;     //Загрузка списка работников для рейтингования
	  N7PP->Visible = false;         //Загрузка производственного задания
	  N8KPE->Visible = false;        //Загрузка КПЭ
	  N51C5->Visible = false;        //Загрузка 5 Почему
	  N10PREEM->Visible = false;     //Загрузка преемников
	  N11SPP->Visible = false;       //Загрузка СПП
	  N12OT->Visible = false;        //Загрузка ОТ
	  N13NOT->Visible = false;       //Загрузка нарушений по ОТ
	  N14TD->Visible = false;        //Загрузка трудовой дисциплины
	  N3Otchet->Visible = false;     //Формирование итогового отчета
	  NSprav->Visible = false;        //Справочник по производственному заданию

	  //Контекстное меню
	  N16Dobav->Visible = false;     //Добавление записи
	  N6Reiting->Visible = false;    //Расчет рейтинга

	  //Кнопки
	  SpeedButton1->Visible = false; //Загрузка списка работников для рейтингования
	  SpeedButton3->Visible = false; //Расчет рейтинга

	  //если права группы УНОУ
	  if (SL_Groups->IndexOf("mmk-itsvc-hrru-unou")>-1)
		{
		  DBGridEh1->Columns->Items[6]->Visible = true;  //Наличие подчиненных
		  DBGridEh1->Columns->Items[13]->Visible = true;    //Отклонения, достижение
		  DBGridEh1->Columns->Items[14]->Visible = true;    //Отклонения, балл
		  DBGridEh1->Columns->Items[17]->Visible = true;    //Информирование, уровень
		  DBGridEh1->Columns->Items[18]->Visible = true;    //Информирование, балл
		  DBGridEh1->Columns->Items[19]->Visible = true;    //5С, уровень
		  DBGridEh1->Columns->Items[20]->Visible = true;    //5С, балл
		  DBGridEh1->Columns->Items[21]->Visible = true;    //КНС, достижение
		  DBGridEh1->Columns->Items[22]->Visible = true;    //КНС, балл

		  //Главное меню
		  N6Spisok->Visible = true;     //Загрузка списка работников для рейтингования
		  N51C5->Visible = true;        //Загрузка 5 Почему

		  //Контекстное меню
		  N16Dobav->Visible = true;     //Добавление записи

		  //Кнопки
		  SpeedButton1->Visible = true; //Загрузка списка работников для рейтингования

		  Prava = "unou";
		}

	  //если права группы загрузки производственного задания
	  else if (SL_Groups->IndexOf("mmk-itsvc-hrru-pp")>-1)
		{
		  DBGridEh1->Columns->Items[9]->Visible = true;     //Производственное задание, факт
		  DBGridEh1->Columns->Items[10]->Visible = true;    //Производственное задание, балл
		  N7PP->Visible = true;         //Загрузка производственного задания

		  Prava = "pp";
		}

	  //если права группы загрузки КПЭ
	  else if (SL_Groups->IndexOf("mmk-itsvc-hrru-kpe")>-1)
		{
		  DBGridEh1->Columns->Items[11]->Visible = true;    //КПЭ, факт
		  DBGridEh1->Columns->Items[12]->Visible = true;    //КПЭ, балл
		  N8KPE->Visible = true;        //Загрузка КПЭ

		  Prava = "kpe";
		}

	  //если права группы загрузки СПП
	  else if (SL_Groups->IndexOf("mmk-itsvc-hrru-spp")>-1)
		{
		  DBGridEh1->Columns->Items[23]->Visible = true;    //СПП, кол-во поданых
		  DBGridEh1->Columns->Items[24]->Visible = true;    //СПП, балл
		  N11SPP->Visible = true;       //Загрузка СПП

		  Prava = "spp";
		}

	  //если права группы загрузки ОТ
	  else if (SL_Groups->IndexOf("mmk-itsvc-hrru-ot")>-1)
		{
		  DBGridEh1->Columns->Items[25]->Visible = true;    //ОТ, ПБ и Э
		  DBGridEh1->Columns->Items[26]->Visible = true;    //Нарушения ОТ
		  N12OT->Visible = true;        //Загрузка ОТ
		  N13NOT->Visible = true;       //Загрузка нарушений по ОТ

		  Prava = "ot";
		}
	  //если права группы загрузки трудовой дисциплины
	  else if (SL_Groups->IndexOf("mmk-itsvc-hrru-td")>-1)
		{
		  DBGridEh1->Columns->Items[27]->Visible = true;    //Трудовая дисциплина
		  N14TD->Visible = true;        //Загрузка трудовой дисциплины

		  Prava = "td";
		}
	  else
		{
		  Application->MessageBox(L"Не установлены права доступа для работы с программой 'Рейтингование руководителей'!!!",L"Права доступа",
								  MB_OK+MB_ICONERROR);
		  Application->Terminate();
		  Abort();

		}
	} */


  //Развернуть на весь экран окно главной формы
  Main->WindowState = wsMaximized;

  //Определение разрешения экрана
  AnsiString width = Screen->Width;     //ширина
  AnsiString height = Screen->Height;   //высота

  //Установка автоматически растягивать грид в зависимости от разрешения
  if (width >= 1280 && height >= 1024 ||
	  width >=1600 && height >= 900)
	{
	  DBGridEh1->AutoFitColWidths = true;
	}
  else
	{
	  DBGridEh1->AutoFitColWidths = false;
	}

  //Фильтрация автоматическая без нажатия Enter
  //DBGridEh1->Style->FilterEditCloseUpApplyFilter =true;
  //DBGridEhCenter()->FilterEditCloseUpApplyFilter = true;
	// RebuildWindowRgn(Panel3);
		// SetWindowLong( this->Handle, GWL_EXSTYLE, this->GetExStyle() | WS_EX_TRANSPARENT );

 /*
что означает:
Новый цвет во фреймбуфере =
			Текущий альфа во фреймбуфере * текущий цвет во фреймбуфере+ (1-текущая альфа во фреймбуфере)*результирующий цвет шейдера
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
   //Свойство - Transparent. Для объекта Image - Image1->Transparent = 50; Panel1->Canvas->Transparent = 50;

   //Фильтрация автоматическая без нажатия Enter
/*  //DBGridEh1->Style->FilterEditCloseUpApplyFilter =true;

  //Определение разрешения экрана
  AnsiString width = Screen->Width;     //ширина
  AnsiString height = Screen->Height;   //высота

 //Установка размера шрифта в зависимости от разрешения экрана
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

  //Убирает прозрачность на кнопках
  SpeedButton1->Glyph->TransparentColor = clBlue;
  SpeedButton2->Glyph->TransparentColor = clBlue;
  SpeedButton3->Glyph->TransparentColor = clBlue;
  SpeedButton4->Glyph->TransparentColor = clBlue;


  //Формирование отчетного периода
   Word Year, Month, Day;

  DecodeDate(Date(),Year, Month, Day);

  //Отчетный год
  god=Year;

  //Отчетный квартал
  if (Month==1 || Month==2 || Month==3) kvartal=1;
  else if (Month==4 || Month==5 || Month==6) kvartal=2;
  else if (Month==7 || Month==8 || Month==9) kvartal=3;
  else if (Month==10 || Month==11 || Month==12) kvartal=4;
  else{
	Application->MessageBox(L"Невозможно определить текущий квартал",L"Ошибка",MB_OK+MB_ICONERROR);
	Application->Terminate();
	Abort();
  }

  //Запрет на редактирование всем кроме отдела по  Оценке персонала после 25 числа
  if (Day>24 && Prava!="ocen") {

	 //Главное меню
	  N6Spisok->Enabled = false;     //Загрузка списка работников для рейтингования
	  N7PP->Enabled = false;         //Загрузка производственного задания
	  N8KPE->Enabled = false;        //Загрузка КПЭ
	  N51C5->Enabled = false;        //Загрузка 5 Почему
	  N10PREEM->Enabled = false;     //Загрузка преемников
	  N11SPP->Enabled = false;       //Загрузка СПП
	  N12OT->Enabled = false;        //Загрузка ОТ
	  N13NOT->Enabled = false;       //Загрузка нарушений по ОТ
	  N14TD->Enabled = false;        //Загрузка трудовой дисциплины

	  //Контекстное меню
	  N16Dobav->Enabled = false;     //Добавление записи
	  N17Redact->Enabled = false;    //Редактирование записи

	  //Кнопки
	  SpeedButton1->Enabled = false; //Загрузка списка работников для рейтингования
	  SpeedButton2->Enabled = false; //Редактирование записи
  }
  else {
	 //Главное меню
	  N6Spisok->Enabled = true;     //Загрузка списка работников для рейтингования
	  N7PP->Enabled = true;         //Загрузка производственного задания
	  N8KPE->Enabled = true;        //Загрузка КПЭ
	  N51C5->Enabled = true;        //Загрузка 5 Почему
	  N10PREEM->Enabled = true;     //Загрузка преемников
	  N11SPP->Enabled = true;       //Загрузка СПП
	  N12OT->Enabled = true;        //Загрузка ОТ
	  N13NOT->Enabled = true;       //Загрузка нарушений по ОТ
	  N14TD->Enabled = true;        //Загрузка трудовой дисциплины

	  //Контекстное меню
	  N16Dobav->Enabled = true;     //Добавление записи
	  N17Redact->Enabled = true;    //Редактирование записи

	  //Кнопки
	  SpeedButton1->Enabled = true; //Загрузка списка работников для рейтингования
	  SpeedButton2->Enabled = true; //Редактирование записи
  }


  //Выборка данных с учетом отчетного периода
  //DM->qReiting->Close();
  DM->qReiting->ParamByName("pgod")->Value= god;
  DM->qReiting->ParamByName("pkvartal")->Value = kvartal;
  try
	{
	  DM->qReiting->Active=true;
	}
  catch(Exception &E)
	{
	  Application->MessageBox(("Возникла ошибка при попытке получения данных из таблицы REIT_RUK "+E.Message).c_str(),L"Ошибка",MB_OK+MB_ICONERROR);
	  Application->Terminate();
	  Abort();
	}


  if (!GetMyDocumentsDir(DocPath))
    {
	  MessageBox(Handle,L"Ошибка доступа к папке документов",L"Ошибка",8208);
	  Application->Terminate();
	  Abort();
	}

  if (!GetTempDir(TempPath))
	{
      MessageBox(Handle,L"Ошибка доступа к временной папке",L"Ошибка",8208);
	  Application->Terminate();
      Abort();
    }

  WorkPath = DocPath + "\\Рейтингование руководителей";
  Path = GetCurrentDir();
  FindWordPath();

  Application->UpdateFormatSettings = false;
  FormatSettings.DecimalSeparator = '.';
  FormatSettings.DateSeparator = '.';
  FormatSettings.ShortDateFormat = "dd.mm.yyyy";

  StatusBar1->SimpleText ="  Отчетный период: "+IntToStr(god)+" год, "+IntToStr(kvartal)+" квартал";

   // Создание ProgressBar на StatusBar
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
//Добавление записи
void __fastcall TMain::N16DobavClick(TObject *Sender)
{
   redakt = 0;
   Vvod->ShowModal();
}
//---------------------------------------------------------------------------
//Редактирование записи
void __fastcall TMain::N17RedactClick(TObject *Sender)
{
  redakt = 1;
  Vvod->ShowModal();
}
//---------------------------------------------------------------------------
//Загрузка общего списка работников
void __fastcall TMain::N6SpisokClick(TObject *Sender)
{
   Variant AppEx, Sh;
   AnsiString  Dir, Sql, tn_proverka="NULL";

   int otchet=0, kol=0, rec=0, ob_kol=0, obnov_kol=0,
   pr=0,
   zex, tn, fio, id_dolg, dolg, uch, podch;



  StatusBar1->SimpleText="  Идет загрузка данных...";

  // Проставление полей для загрузки
  zex=2;     //B
  tn=4;      //D
  fio=5;     //E
  id_dolg=7; //G
  dolg=6;    //F
  uch=8;     //H
  podch=9;   //I
  update=0;

  StatusBar1->SimpleText="  Выбор документа для загрузки...";

  OpenDialog1->Filter = "Excel files (*.xls, *.xlsx)|*.xls; *.xlsx";
  // DefaultExt

  //Выбор файла для загрузки
  if (!OpenDialog1->Execute()){
	  StatusBar1->SimpleText ="  Отчетный период: "+IntToStr(god)+" год, "+IntToStr(kvartal)+" квартал";
	  Abort();
  }

  StatusBar1->SimpleText = "  Загрузка данных из файла "+OpenDialog1->FileName;


   //Открытие файла данных для записи не обновленных данных
  if (!rtf_Open((TempPath + "\\zagruzka.txt").c_str()))
	{
	  MessageBox(Handle,L"Ошибка открытия файла данных",L"Ошибка",8192);
	  Abort();
	}

  rtf_Out("data", DateTimeToStr(Now()),0);


  //Открытие документа Excel
  try
	{
	  AppEx = CreateOleObject("Excel.Application");
	}
  catch (...)
	{
	  Application->MessageBox(L"Невозможно открыть Microsoft Excel!\n Возможно это приложение на компьютере не установлено.",
							  L"Ошибка", MB_OK+MB_ICONERROR);
	  StatusBar1->SimpleText ="  Отчетный период: "+IntToStr(god)+" год, "+IntToStr(kvartal)+" квартал";
	  Abort();
	}

  //Если возникает ошибка во время формирования отчета
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
		  Application->MessageBox(L"Ошибка открытия книги Microsoft Excel!", L"Ошибка",MB_OK + MB_ICONERROR);
		  StatusBar1->SimpleText ="  Отчетный период: "+IntToStr(god)+" год, "+IntToStr(kvartal)+" квартал";
		  Abort();
		}


	  //Определяет количество занятых строк в документе
	  AnsiString Row = Sh.OlePropertyGet("UsedRange").OlePropertyGet("Rows").OlePropertyGet("Count");


	  //Проверка на наличие данных в таблице
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
		  Application->MessageBox(("Возникла ошибка при попытке выбрать данные из таблицы REIT_RUK: " + E.Message).c_str(),L"Ошибка",
									MB_OK+MB_ICONERROR);

		  InsertLog("Возникла ошибка при загрузке списка работников для рейтингования из файла '"+OpenDialog1->FileName+"' за "+IntToStr(god)+" год, "+IntToStr(kvartal)+" квартал");
		  DM->qReiting->Refresh();
		  StatusBar1->SimpleText ="  Отчетный период: "+IntToStr(god)+" год, "+IntToStr(kvartal)+" квартал";
		  Abort();
		}

	  if (DM->qObnovlenie->FieldByName("kol")->AsInteger>0)
		{
		  if (Application->MessageBox(("В таблице уже содержаться записи за "+IntToStr(kvartal)+" квартал "+IntToStr(god)+" год\nВсе предыдущие данные за этот период будут перезаписаны\nВы действительно хотите обновить данные?").c_str(),
										L"Изменение данных",MB_YESNO+MB_ICONWARNING)==ID_NO)
			 {
			   update=0;
			   Abort();
			 }

		   //Удаление предыдущих записей и загрузка заново
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
			   Application->MessageBox(("Возникла ошибка при попытке выбрать данные из таблицы REIT_RUK: " + E.Message).c_str(),L"Ошибка",
										MB_OK+MB_ICONERROR);

			   InsertLog("Возникла ошибка при загрузке списка работников для рейтингования из файла '"+OpenDialog1->FileName+"' за "+IntToStr(god)+" год, "+IntToStr(kvartal)+" квартал");
			   DM->qReiting->Refresh();
			   StatusBar1->SimpleText ="  Отчетный период: "+IntToStr(god)+" год, "+IntToStr(kvartal)+" квартал";
			   Abort();
			 }

		   update=0;
		}
	 // else update=0;

	  StatusBar1->SimpleText ="  Выполняется загрузка списка работников...";


	  Cursor = crHourGlass;
	  ProgressBar->Position = 0;
	  ProgressBar->Visible = true;
	  ProgressBar->Max=StrToInt(Row);

	  //Загрузка данных
	  for (int i=1; i<Row+1; i++)
		{
		  tn_proverka = Sh.OlePropertyGet("Cells",i,tn);//.OlePropertyGet("Value");


		  //Проверка на наличие таб.№ и поиск строки с которой загружается файл
		  if (tn_proverka.IsEmpty() || !Proverka(tn_proverka))  continue;
			{
//******************************************************************************
			  //Проверка на совпадение данных в кадрах и в загружаемом списке работников
			  DM->qProverka->Close();
			  DM->qProverka->ParamByName("ptn_sap")->Value=tn_proverka;

			  try
				{
				  DM->qProverka->Active = true;
				}
			  catch (Exception &E)
				{
				  Application->MessageBox(("Возникла ошибка при попытке выбрать данные из таблицы SAP_OSN_SVED: " + E.Message).c_str(),L"Ошибка",
										   MB_OK+MB_ICONERROR);

				  InsertLog("Возникла ошибка при загрузке списка работников для рейтингования из файла '"+OpenDialog1->FileName+"' за "+IntToStr(god)+" год, "+IntToStr(kvartal)+" квартал");
				  DM->qReiting->Refresh();
				  StatusBar1->SimpleText ="  Отчетный период: "+IntToStr(god)+" год, "+IntToStr(kvartal)+" квартал";
				  Cursor = crDefault;
				  ProgressBar->Visible = false;
				  Abort();
				}


			  //Несоответствие шифра цеха
			  if (DM->qProverka->FieldByName("zex")->AsString!=VarToStr(Sh.OlePropertyGet("Cells",i,zex)))
				{
				  //Формирование заголовка и шапки таблицы
				  if (pr!=1)
					{
					   rtf_Out("z", " ",1);
					   if(!rtf_LineFeed())
						 {
						   MessageBox(Handle,L"Ошибка записи в файл данных",L"Ошибка",8192);
						   if (!rtf_Close()) MessageBox(Handle,L"Ошибка закрытия файла данных",L"Ошибка",8192);
						   return;
						 }
					 }
						 //AnsiString aarr =  Sh.OlePropertyGet("Cells",i,tn);



				   //Формирование отчета по необновленным записям
				   rtf_Out("tn", VarToStr(Sh.OlePropertyGet("Cells",i,tn)),2);
				   rtf_Out("zex_f", VarToStr(Sh.OlePropertyGet("Cells",i,zex)),2);
				   rtf_Out("zex", DM->qProverka->FieldByName("zex")->AsString,2);
				   rtf_Out("fio", VarToStr(Sh.OlePropertyGet("Cells",i,fio)),2);

				   if(!rtf_LineFeed())
					 {
					   MessageBox(Handle,L"Ошибка записи в файл данных",L"Ошибка",8192);
					   if (!rtf_Close()) MessageBox(Handle,L"Ошибка закрытия файла данных",L"Ошибка",8192);
					   return;
					 }
				   pr=1;      //Признак формирования шапки отчета
				   otchet=1;  //Признак формирования отчета по необновленным записям
				}

			  //Несоответствие ФИО
			  if (DM->qProverka->FieldByName("fio")->AsString!=VarToStr(Sh.OlePropertyGet("Cells",i,fio)))
				{
				  //Формирование заголовка и шапки таблицы
				  if (pr!=3)
					{
					   rtf_Out("z", " ",3);
					   if(!rtf_LineFeed())
						 {
						   MessageBox(Handle,L"Ошибка записи в файл данных",L"Ошибка",8192);
						   if (!rtf_Close()) MessageBox(Handle,L"Ошибка закрытия файла данных",L"Ошибка",8192);
						   return;
						 }
					}

				   //Формирование отчета по необновленным записям
				   rtf_Out("tn", VarToStr(Sh.OlePropertyGet("Cells",i,tn)),4);
				   rtf_Out("zex", VarToStr(Sh.OlePropertyGet("Cells",i,zex)),4);
				   rtf_Out("fio_f", VarToStr(Sh.OlePropertyGet("Cells",i,fio)),4);
				   rtf_Out("fio", DM->qProverka->FieldByName("fio")->AsString,4);

				   if(!rtf_LineFeed())
					 {
					   MessageBox(Handle,L"Ошибка записи в файл данных",L"Ошибка",8192);
					   if (!rtf_Close()) MessageBox(Handle,L"Ошибка закрытия файла данных",L"Ошибка",8192);
					   return;
					 }
				   pr=3;      //Признак формирования шапки отчета
				   otchet=1;  //Признак формирования отчета по необновленным записям
				}
			  //Несоответствие ИД должности
			  if (DM->qProverka->FieldByName("id_shtat")->AsString!=VarToStr(Sh.OlePropertyGet("Cells",i,id_dolg)))
				{
				  //Формирование заголовка и шапки таблицы
				   if (pr!=5)
					 {
					   rtf_Out("z", " ",5);
					   if(!rtf_LineFeed())
						 {
						   MessageBox(Handle,L"Ошибка записи в файл данных",L"Ошибка",8192);
						   if (!rtf_Close()) MessageBox(Handle,L"Ошибка закрытия файла данных",L"Ошибка",8192);
						   return;
						 }
					 }

				   //Формирование отчета по необновленным записям
				   rtf_Out("tn", VarToStr(Sh.OlePropertyGet("Cells",i,tn)),6);
				   rtf_Out("zex", VarToStr(Sh.OlePropertyGet("Cells",i,zex)),6);
				   rtf_Out("id_dolg_f", VarToStr(Sh.OlePropertyGet("Cells",i,id_dolg)),6);
				   rtf_Out("id_dolg", DM->qProverka->FieldByName("id_shtat")->AsString,6);
				   rtf_Out("fio", VarToStr(Sh.OlePropertyGet("Cells",i,fio)),6);

				   if(!rtf_LineFeed())
					 {
					   MessageBox(Handle,L"Ошибка записи в файл данных",L"Ошибка",8192);
					   if (!rtf_Close()) MessageBox(Handle,L"Ошибка закрытия файла данных",L"Ошибка",8192);
					   return;
					 }
				   pr=5;      //Признак формирования шапки отчета
				   otchet=1;  //Признак формирования отчета по необновленным записям
				}
			  //Несоответствие наименования должности
			  if (DM->qProverka->FieldByName("name_dolg_ru")->AsString!=VarToStr(Sh.OlePropertyGet("Cells",i,dolg)))
				{
					//Формирование заголовка и шапки таблицы
				   if (pr!=7)
					 {
					   rtf_Out("z", " ",7);
					   if(!rtf_LineFeed())
						 {
						   MessageBox(Handle,L"Ошибка записи в файл данных",L"Ошибка",8192);
						   if (!rtf_Close()) MessageBox(Handle,L"Ошибка закрытия файла данных",L"Ошибка",8192);
						   return;
						 }
					 }

				   //Формирование отчета по необновленным записям
				   rtf_Out("tn", VarToStr(Sh.OlePropertyGet("Cells",i,tn)),8);
				   rtf_Out("zex", VarToStr(Sh.OlePropertyGet("Cells",i,zex)),8);
				   rtf_Out("dolg_f", VarToStr(Sh.OlePropertyGet("Cells",i,dolg)),8);
				   rtf_Out("dolg", DM->qProverka->FieldByName("name_dolg_ru")->AsString,8);

				   if(!rtf_LineFeed())
					 {
					   MessageBox(Handle,L"Ошибка записи в файл данных",L"Ошибка",8192);
					   if (!rtf_Close()) MessageBox(Handle,L"Ошибка закрытия файла данных",L"Ошибка",8192);
					   return;
					 }
				   pr=7;      //Признак формирования шапки отчета
				   otchet=1;  //Признак формирования отчета по необновленным записям
				}


//******************************************************************************
			  //Проверка на наличие данных в таблице

			  //Загрузка данных в базу
			  if (update==1)
				{
				  Sql = "update reit_ruk set \
										 zex=trim('"+ Sh.OlePropertyGet("Cells",i,zex) +"'), \
										 tn=trim('"+ Sh.OlePropertyGet("Cells",i,tn) +"'), \
										 fio=initcap(trim('"+ Sh.OlePropertyGet("Cells",i,fio) +"')),  \
										 id_dolg=lpad(trim('"+ Sh.OlePropertyGet("Cells",i,id_dolg) +"'),'0',8), \
										 dolg=trim('"+ Sh.OlePropertyGet("Cells",i,dolg) +"'),  \
										 uch=trim('"+ Sh.OlePropertyGet("Cells",i,uch) +"'),  \
										 podch=decode(trim('"+ Sh.OlePropertyGet("Cells",i,podch) +"'),'да','1',0) \
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
								   decode(trim('"+ Sh.OlePropertyGet("Cells",i,podch) +"'),'да','1',0)) ";
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
				  Application->MessageBox(("Возникла ошибка при попытке обновить данные в таблице REIT_RUK: " + E.Message).c_str(),L"Ошибка",
											MB_OK+MB_ICONERROR);

				  InsertLog("Возникла ошибка при загрузке списка работников для рейтингования из файла '"+OpenDialog1->FileName+"' за "+IntToStr(god)+" год, "+IntToStr(kvartal)+" квартал");
				  DM->qReiting->Refresh();
				  StatusBar1->SimpleText ="  Отчетный период: "+IntToStr(god)+" год, "+IntToStr(kvartal)+" квартал";
				  Cursor = crDefault;
				  ProgressBar->Visible = false;
				  Abort();
				}

			  rec++;
			  kol+=DM->qObnovlenie->RowsAffected;

			  // Количество обновленных записей
			  if (DM->qObnovlenie->RowsAffected == 0)
				{
				  //Формирование заголовка и шапки таблицы
				   if (pr!=9)
					 {
					   rtf_Out("z", " ",9);
					   if(!rtf_LineFeed())
						 {
						   MessageBox(Handle,L"Ошибка записи в файл данных",L"Ошибка",8192);
						   if (!rtf_Close()) MessageBox(Handle,L"Ошибка закрытия файла данных",L"Ошибка",8192);
						   return;
						 }
					 }

				   //Формирование отчета по необновленным записям
				   rtf_Out("tn", VarToStr(Sh.OlePropertyGet("Cells",i,tn)),10);
				   rtf_Out("zex", VarToStr(Sh.OlePropertyGet("Cells",i,zex)),10);
				   rtf_Out("fio", VarToStr(Sh.OlePropertyGet("Cells",i,fio)),10);

				   if(!rtf_LineFeed())
					 {
					   MessageBox(Handle,L"Ошибка записи в файл данных",L"Ошибка",8192);
					   if (!rtf_Close()) MessageBox(Handle,L"Ошибка закрытия файла данных",L"Ошибка",8192);
					   return;
					 }
				   pr=9;      //Признак формирования шапки отчета
				   otchet=1;  //Признак формирования отчета по необновленным записям
				 }
			   else obnov_kol++;
		  }

		  ProgressBar->Position++;
		  ob_kol++;
		}


	  StatusBar1->SimpleText = "  Загрузка данных выполнена.";

	  DM->qReiting->Refresh();

	  //Закрытие Excel
	  AppEx.OleProcedure("Quit");
	  AppEx = Unassigned;


	  if(!rtf_Close())
		{
		  MessageBox(Handle,L"Ошибка закрытия файла данных", L"Ошибка", 8192);
		  return;
		}

	  //Формирование отчета в Word
	  if (otchet==1)
		{
		  StatusBar1->SimpleText = "  Формирование отчета с ошибками...";

		  //Создание папки, если ее не существует
		  ForceDirectories(WorkPath);

		  int istrd;
		  try
			{
			  rtf_CreateReport(TempPath + "\\zagruzka.txt", Path+"\\RTF\\zagruzka.rtf",
							   WorkPath+"\\Отчет.doc",NULL,&istrd);


			  WinExec(("\""+ WordPath+"\"\""+WorkPath+"\\Отчет.doc\"").c_str(),SW_MAXIMIZE);

			}
		  catch(RepoRTF_Error E)
			{
			  Application->MessageBox(("Ошибка формирования отчета:"+ String(E.Err)+
								 "\nСтрока файла данных:"+IntToStr(istrd)).c_str(),
								 L"Ошибка",
								 MB_OK+MB_ICONERROR);
			}

		  Application->MessageBox(("Существует несоответствие информации в загружаемом файле и базе данных по Персоналу.\nПроверьте достоверность информации в файле \n "+OpenDialog1->FileName+" и выполните повторную загрузку").c_str() ,L" Загрузка списка работников",
								  MB_OK + MB_ICONINFORMATION);

		}

	  DeleteFile(TempPath+"\\otchet.txt");
	  InsertLog("Загрузка списка работников для рейтингования из файла '"+OpenDialog1->FileName+"' за "+IntToStr(god)+" год, "+IntToStr(kvartal)+" квартал выполнена успешно. Загружено " + IntToStr(obnov_kol) + " из " + IntToStr(ob_kol)+" записей");


       Application->MessageBox(("Загрузка списка работников для рейтингования выполнена успешно. =) \nЗагружено " + IntToStr(obnov_kol) + " из " + IntToStr(ob_kol)+" записей").c_str(),
						   L"Обновление данных по оценке персонала",
						   MB_OK + MB_ICONINFORMATION);
	}
  catch(...)
	{
	  AppEx.OleProcedure("Quit");
	  //AppEx.Clear();
	  //VarClear(AppEx);
	  AppEx=Unassigned;
	  InsertLog("Возникла ошибка при загрузке списка работников для рейтингования из файла '"+OpenDialog1->FileName+"' за "+IntToStr(god)+" год, "+IntToStr(kvartal)+" квартал");
	}


  Cursor = crDefault;
  ProgressBar->Position = 0;
  ProgressBar->Visible = false;

  StatusBar1->SimplePanel = false;
  StatusBar1->SimpleText ="  Отчетный период: "+IntToStr(god)+" год, "+IntToStr(kvartal)+" квартал";

}
//---------------------------------------------------------------------------


// Проверка на значение таб.№ в Excel-файле
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

// Возвращает путь на папку "Мои документы"
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

// Возвращает путь на папку Temp
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

 // Определяет версию установленного Word
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
//Редактирование записи
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

//Логи
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
	  Application->MessageBoxW(L"Возникла ошибка при вставке данных в таблицу LOGS_REIT",L"Ошибка",
							   MB_ICONERROR);
	}

  DM->qObnovlenie->Close();
}
 //---------------------------------------------------------------------------   */
//Производственное задание
void __fastcall TMain::N7PPClick(TObject *Sender)
{
   Zagruzka(3,1);
}
//---------------------------------------------------------------------------
//КПЭ
void __fastcall TMain::N8KPEClick(TObject *Sender)
{
   Zagruzka(3,2);
}
//---------------------------------------------------------------------------
//---------------------------------------------------------------------------
//5 Почему
void __fastcall TMain::N51C5Click(TObject *Sender)
{
  Zagruzka(2,3);
}
//---------------------------------------------------------------------------
//Преемники
void __fastcall TMain::N10PREEMClick(TObject *Sender)
{
  Zagruzka(3,4);
}
//---------------------------------------------------------------------------
//СПП
void __fastcall TMain::N11SPPClick(TObject *Sender)
{
  Zagruzka(3,5);
}
//---------------------------------------------------------------------------
//ОТ, ПБ и Э
void __fastcall TMain::N12OTClick(TObject *Sender)
{
  Zagruzka(3,6);
}
//---------------------------------------------------------------------------
//Нарушения ОТ
void __fastcall TMain::N13NOTClick(TObject *Sender)
{
  Zagruzka(3,7);
}
//---------------------------------------------------------------------------
//Трудовая дисциплина
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


  StatusBar1->SimpleText="  Идет загрузка данных...";

  // Проставление полей для загрузки
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


  StatusBar1->SimpleText="  Выбор документа для загрузки...";

  OpenDialog1->Filter = "Excel files (*.xls, *.xlsx)|*.xls; *.xlsx";
  // DefaultExt

  //Выбор файла для загрузки
  if (!OpenDialog1->Execute()){
	  Abort();
  }

  StatusBar1->SimpleText = "  Загрузка данных из файла "+OpenDialog1->FileName;


   //Открытие файла данных для записи не обновленных данных
  if (!rtf_Open((TempPath + "\\zagruzka.txt").c_str()))
	{
	  MessageBox(Handle,L"Ошибка открытия файла данных",L"Ошибка",8192);
	  Abort();
	}

  rtf_Out("data", DateTimeToStr(Now()),0);


  //Открытие документа Excel
  try
	{
	  AppEx = CreateOleObject("Excel.Application");
	}
  catch (...)
	{
	  Application->MessageBox(L"Невозможно открыть Microsoft Excel!\n Возможно это приложение на компьютере не установлено.",
							  L"Ошибка", MB_OK+MB_ICONERROR);
	  StatusBar1->SimpleText ="  Отчетный период: "+IntToStr(god)+" год, "+IntToStr(kvartal)+" квартал";
	  Abort();
	}

  //Если возникает ошибка во время формирования отчета
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
		  Application->MessageBox(L"Ошибка открытия книги Microsoft Excel!", L"Ошибка",MB_OK + MB_ICONERROR);
		  StatusBar1->SimpleText ="  Отчетный период: "+IntToStr(god)+" год, "+IntToStr(kvartal)+" квартал";
		  Abort();
		}


	  //Определяет количество занятых строк в документе
	  AnsiString Row = Sh.OlePropertyGet("UsedRange").OlePropertyGet("Rows").OlePropertyGet("Count");


	  switch (otchet) {
		  case 1: str = "производственному заданию...";
		  break;
		  case 2: str = "КПЭ...";
		  break;
		  case 3: str = "5 Почему...";
		  break;
		  case 4: str = "наличию преемников...";
		  break;
		  case 5: str = "СПП...";
		  break;
		  case 6: str = "охране труда...";
		  break;
		  case 7: str = "соблюдению требований ОТ...";
		  break;
		  default: str = "трудовой дисциплине...";
	  }


	  StatusBar1->SimpleText ="  Выполняется загрузка данных по "+str;

      Cursor = crHourGlass;
	  ProgressBar->Position = 0;
	  ProgressBar->Visible = true;
	  ProgressBar->Max=StrToInt(Row);

	  //Загрузка данных
	  for (int i=1; i<Row+1; i++)
		{
		  tn_proverka = Sh.OlePropertyGet("Cells",i,tn);//.OlePropertyGet("Value");


		  //Проверка на наличие таб.№ и поиск строки с которой загружается файл
		  if (tn_proverka.IsEmpty() || !Proverka(tn_proverka))  continue;
			{
//******************************************************************************
			  //Проверка на наличие работника в кадрах
			  DM->qProverka->Close();
			  DM->qProverka->ParamByName("ptn_sap")->Value=tn_proverka;

			  try
				{
				  DM->qProverka->Active = true;
				}
			  catch (Exception &E)
				{
				  Application->MessageBox(("Возникла ошибка при попытке выбрать данные из таблицы SAP_OSN_SVED: " + E.Message).c_str(),L"Ошибка",
										   MB_OK+MB_ICONERROR);

				  DM->qReiting->Refresh();
				  StatusBar1->SimpleText ="  Отчетный период: "+IntToStr(god)+" год, "+IntToStr(kvartal)+" квартал";
				  InsertLog("Возникла ошибка при загрузке данных по "+str+" из файла '"+OpenDialog1->FileName+"' за "+IntToStr(god)+" год, "+IntToStr(kvartal)+" квартал");
				  Cursor = crDefault;
				  ProgressBar->Visible = false;
				  Abort();
				}


			  //Формирование отчета
			  if (DM->qProverka->RecordCount==0)
				{
                   //Формирование заголовка и шапки таблицы
				   if (pr!=1)
					 {
					   rtf_Out("z", " ",1);
					   if(!rtf_LineFeed())
						 {
						   MessageBox(Handle,L"Ошибка записи в файл данных",L"Ошибка",8192);
						   if (!rtf_Close()) MessageBox(Handle,L"Ошибка закрытия файла данных",L"Ошибка",8192);
						   return;
						 }
					 }


				   rtf_Out("tn", VarToStr(Sh.OlePropertyGet("Cells",i,tn)),2);
				   rtf_Out("fio", VarToStr(Sh.OlePropertyGet("Cells",i,fio)),2);

				   if(!rtf_LineFeed())
					 {
					   MessageBox(Handle,L"Ошибка записи в файл данных",L"Ошибка",8192);
					   if (!rtf_Close()) MessageBox(Handle,L"Ошибка закрытия файла данных",L"Ошибка",8192);
					   return;
					 }
				   pr=1;      //Признак формирования шапки отчета
				   fotchet=1;  //Признак формирования отчета по необновленным записям
				}

			   float treb;
//******************************************************************************
			  //Загрузка данных в базу
			  switch (otchet)
			  {

				//Производственное задание
				case 1:
					//Если указан процентный формат
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

				//КПЭ
				case 2:
					//Если указан процентный формат
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

				//5 Почему
				case 3:


					//Если указан процентный формат
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

				//Наличие преемников
				case 4:
					Sql = "update reit_ruk set \
										 priem=trim('"+ Sh.OlePropertyGet("Cells",i,priem) +"'),  \
										 priem_ball=trim('"+ Sh.OlePropertyGet("Cells",i,priem_ball) +"')  \
						  where tn="+ Sh.OlePropertyGet("Cells",i,tn)+" and god="+IntToStr(god) +" and kvart="+IntToStr(kvartal);
				break;

				//СПП
				case 5:
					Sql = "update reit_ruk set \
										 spp_kol=trim('"+ Sh.OlePropertyGet("Cells",i,spp_kol) +"'),  \
										 spp_ball=trim('"+ Sh.OlePropertyGet("Cells",i,spp_ball) +"')  \
						  where tn="+ Sh.OlePropertyGet("Cells",i,tn)+" and god="+IntToStr(god) +" and kvart="+IntToStr(kvartal);
				break;

				//Охрана труда
				case 6:
					Sql = "update reit_ruk set \
										 ot_upr=trim('"+ Sh.OlePropertyGet("Cells",i,ot_upr) +"')  \
						  where tn="+ Sh.OlePropertyGet("Cells",i,tn)+" and god="+IntToStr(god) +" and kvart="+IntToStr(kvartal);
				break;

				//Соблюдение требований ОТ
				case 7:
					Sql = "update reit_ruk set \
										 ot_treb=trim('"+ Sh.OlePropertyGet("Cells",i,ot_treb).OlePropertyGet("Value") +"')  \
						  where tn="+ Sh.OlePropertyGet("Cells",i,tn)+" and god="+IntToStr(god) +" and kvart="+IntToStr(kvartal);
				break;

				//Трудовая дисциплина
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
				  Application->MessageBox(("Возникла ошибка при попытке обновить данные в таблице REIT_RUK: " + E.Message).c_str(),L"Ошибка",
											MB_OK+MB_ICONERROR);

				  DM->qReiting->Refresh();
				  StatusBar1->SimpleText ="  Отчетный период: "+IntToStr(god)+" год, "+IntToStr(kvartal)+" квартал";
				  InsertLog("Возникла ошибка при загрузке данных по "+str+" из файла '"+OpenDialog1->FileName+"' за "+IntToStr(god)+" год, "+IntToStr(kvartal)+" квартал");
				  Cursor = crDefault;
				  ProgressBar->Visible = false;
				  Abort();
				}

			  rec++;
			  kol+=DM->qObnovlenie->RowsAffected;

			  // Количество обновленных записей
			  if (DM->qObnovlenie->RowsAffected == 0)
				{
				  //Формирование заголовка и шапки таблицы
				   if (pr!=2)
					 {
					   rtf_Out("z", " ",3);
					   if(!rtf_LineFeed())
						 {
						   MessageBox(Handle,L"Ошибка записи в файл данных",L"Ошибка",8192);
						   if (!rtf_Close()) MessageBox(Handle,L"Ошибка закрытия файла данных",L"Ошибка",8192);
						   return;
						 }
					 }

				   //Формирование отчета по необновленным записям
				   rtf_Out("tn", VarToStr(Sh.OlePropertyGet("Cells",i,tn)),4);
				   rtf_Out("fio", VarToStr(Sh.OlePropertyGet("Cells",i,fio)),4);

				   if(!rtf_LineFeed())
					 {
					   MessageBox(Handle,L"Ошибка записи в файл данных",L"Ошибка",8192);
					   if (!rtf_Close()) MessageBox(Handle,L"Ошибка закрытия файла данных",L"Ошибка",8192);
					   return;
					 }
				   pr=2;      //Признак формирования шапки отчета
				   fotchet=1;  //Признак формирования отчета по необновленным записям
				 }
			   else obnov_kol++;

			}

          ProgressBar->Position++;
		  ob_kol++;
		}


      DM->qReiting->Refresh();

	  //Закрытие Excel
	  AppEx.OleProcedure("Quit");
	  AppEx = Unassigned;


	  if(!rtf_Close())
		{
		  MessageBox(Handle,L"Ошибка закрытия файла данных", L"Ошибка", 8192);
		  return;
		}

	  //Формирование отчета в Word
	  if (fotchet==1)
		{
		  StatusBar1->SimpleText = "Формирование отчета с ошибками...";

		  //Создание папки, если ее не существует
		  ForceDirectories(WorkPath);

		  int istrd;
		  try
			{
			  rtf_CreateReport(TempPath + "\\zagruzka.txt", Path+"\\RTF\\zagruzka2.rtf",
							   WorkPath+"\\Отчет по загрузке данных.doc",NULL,&istrd);


			  WinExec(("\""+ WordPath+"\"\""+WorkPath+"\\Отчет по загрузке данных.doc\"").c_str(),SW_MAXIMIZE);

			}
		  catch(RepoRTF_Error E)
			{
			  Application->MessageBox(("Ошибка формирования отчета:"+ String(E.Err)+
								 "\nСтрока файла данных:"+IntToStr(istrd)).c_str(),
								 L"Ошибка",
								 MB_OK+MB_ICONERROR);
			}

		  Application->MessageBox(("Существует несоответствие информации в загружаемом файле и базе данных по Персоналу.\nПроверьте достоверность информации в файле \n "+OpenDialog1->FileName+" и выполните повторную загрузку").c_str() ,L" Загрузка данных",
								  MB_OK + MB_ICONINFORMATION);
		}

	  DeleteFile(TempPath+"\\otchet.txt");

	  InsertLog("Загрузка данных по "+str+" из файла '"+OpenDialog1->FileName+"' за "+IntToStr(god)+" год, "+IntToStr(kvartal)+" квартал выполнена успешно. Загружено " + IntToStr(obnov_kol) + " из " + IntToStr(ob_kol)+" записей");

	  Application->MessageBox(("Загрузка данных по "+WideString(str)+" выполнена успешно. =) \nЗагружено " + IntToStr(obnov_kol) + " из " + IntToStr(ob_kol)+" записей").c_str(),
							   L"Обновление данных по оценке персонала",
						       MB_OK + MB_ICONINFORMATION);
	}
  catch(...)
	{
	  AppEx.OleProcedure("Quit");
	  //AppEx.Clear();
	  //VarClear(AppEx);
	  AppEx=Unassigned;
	  InsertLog("Возникла ошибка при загрузке данных по "+str+" из файла '"+OpenDialog1->FileName+"' за "+IntToStr(god)+" год, "+IntToStr(kvartal)+" квартал");
	}


  Cursor = crDefault;
  ProgressBar->Position = 0;
  ProgressBar->Visible = false;

  StatusBar1->SimplePanel = false;
  StatusBar1->SimpleText ="  Отчетный период: "+IntToStr(god)+" год, "+IntToStr(kvartal)+" квартал";
}
//---------------------------------------------------------------------------




void __fastcall TMain::SpeedButton1Click(TObject *Sender)
{
   N6SpisokClick(Sender);
}
//---------------------------------------------------------------------------
//Автоматический рассчет оценки по всем работникам
void __fastcall TMain::SpeedButton3Click(TObject *Sender)
{
  if (Application->MessageBox(("Будет выполнен расчет оценок и рейтинга по всем работникам \nза "+IntToStr(god)+" год, "+IntToStr(kvartal)+" квартал. Продолжить?").c_str(), L"Расчет рейтинга",
							  MB_YESNO+MB_ICONINFORMATION )==ID_YES)
	{
      //Рассчет оценки
	  RaschetOcen(0);

	  //Рсссчкт рейтинга
      RaschetReit(0, NULL, NULL);
	}

}
//---------------------------------------------------------------------------
//Автоматический рассчет оценки по одному работнику
void __fastcall TMain::N6ReitingClick(TObject *Sender)
{
  RaschetOcen(1);

  //Рсссчкт рейтинга
  RaschetReit(1, DM->qReiting->FieldByName("zex")->AsString, DM->qReiting->FieldByName("podch")->AsInteger);
}
//---------------------------------------------------------------------------
//Автоматический рассчет оценки
void __fastcall TMain::RaschetOcen(int pr)
{
  AnsiString Sql;

  StatusBar1->SimpleText ="Идет расчет рейтинга по всем работникам за: "+IntToStr(god)+" год, "+IntToStr(kvartal)+" квартал";


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
	  Application->MessageBox(("Возникла ошибка при рассчете рейтинга (таблица REIT_RUK) "+E.Message).c_str(),L"Ошибка",
							  MB_OK+MB_ICONERROR);

	  StatusBar1->SimpleText ="Отчетный период: "+IntToStr(god)+" год, "+IntToStr(kvartal)+" квартал";
	  if (pr==1) InsertLog("Расчет рейтинга по работнику с таб.№ = "+DM->qReiting->FieldByName("tn")->AsString+" не выполнен");
	  else InsertLog("Расчет рейтинга по всем работникам не выполнен");
	  Abort();
	}

   if (pr==1) InsertLog("Расчет оценки по работнику с таб.№ = "+DM->qReiting->FieldByName("tn")->AsString+" выполнен успешно");
   else InsertLog("Расчет оценки по всем работникам выполнен успешно");
  // Application->MessageBox(L"Расчет оценки по всем работникам выполнен успешно!!!" ,L"Расчет рейтинга",
  //								  MB_OK + MB_ICONINFORMATION);

    DM->qReiting->Refresh();

}
//---------------------------------------------------------------------------

//Автоматический рассчет рейтинга
void __fastcall TMain::RaschetReit(int pr, String zex, int podch)
{
  AnsiString Sql;
  int kol_kr_zona=0, kol_zl_zona=0;


  //Выбор общего списка цехов для рейтингования
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
	  Application->MessageBox(("Возникла ошибка при рассчете рейтинга (таблица REIT_RUK) "+E.Message).c_str(),L"Ошибка",
							  MB_OK+MB_ICONERROR);

	  StatusBar1->SimpleText ="Отчетный период: "+IntToStr(god)+" год, "+IntToStr(kvartal)+" квартал";
	  if (pr==1) InsertLog("Расчет рейтинга по работнику с таб.№ = "+DM->qReiting->FieldByName("tn")->AsString+" не выполнен");
	  else InsertLog("Расчет рейтинга по всем работникам не выполнен");
	  Abort();
	}

  if (DM->qObnovlenie->RecordCount==0)
  {
	 Application->MessageBox(L"Нет работников для рейтингования удовлетворяющих условию рейтингования (минимальное количество для рейтингования 5 человек в пределах одного структурного подразделения и одной из категории должностей (с подчиненными и без подчиненных))",L"Предупреждение",
							 MB_OK+MB_ICONINFORMATION);
	  StatusBar1->SimpleText ="Отчетный период: "+IntToStr(god)+" год, "+IntToStr(kvartal)+" квартал";
	  if (pr==1) InsertLog("Расчет рейтинга по работнику с таб.№ = "+DM->qReiting->FieldByName("tn")->AsString+" не выполнен, в связи с отсутствием работников, удовлетворяющих условию рейтингования");
	  else InsertLog("Расчет рейтинга по всем работникам не выполнен, в связи с отсутствием работников, удовлетворяющих условию рейтингования");
	  Abort();
  }

  while (!DM->qObnovlenie->Eof)
	{
	  //Очистка значений по рейтингу для пересчета
	  //Выбор общего списка цехов для рейтингования
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
		  Application->MessageBox(("Возникла ошибка при рассчете рейтинга (таблица REIT_RUK) "+E.Message).c_str(),L"Ошибка",
								  MB_OK+MB_ICONERROR);

		  StatusBar1->SimpleText ="Отчетный период: "+IntToStr(god)+" год, "+IntToStr(kvartal)+" квартал";
		  if (pr==1) InsertLog("Расчет рейтинга по работнику с таб.№ = "+DM->qReiting->FieldByName("tn")->AsString+" не выполнен");
		  else InsertLog("Расчет рейтинга по всем работникам не выполнен");
		  Abort();
		}

	   kol_kr_zona=0;

	  //Выбор списка работников по оцениваемому цеху
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
		  Application->MessageBox(("Возникла ошибка при рассчете рейтинга (таблица REIT_RUK) "+E.Message).c_str(),L"Ошибка",
									  MB_OK+MB_ICONERROR);

		  StatusBar1->SimpleText ="Отчетный период: "+IntToStr(god)+" год, "+IntToStr(kvartal)+" квартал";
		  if (pr==1) InsertLog("Расчет рейтинга по работнику с таб.№ = "+DM->qReiting->FieldByName("tn")->AsString+" не выполнен");
		  else InsertLog("Расчет рейтинга по всем работникам не выполнен");
		  Abort();
		}

	  if (DM->qRaschet->RecordCount>0)
		{
		  //********************************************************************
		  //Обновление красной зоны пока не будет 20% от числа
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
				  Application->MessageBox(("Возникла ошибка при рассчете рейтинга (таблица REIT_RUK) "+E.Message).c_str(),L"Ошибка",
										  MB_OK+MB_ICONERROR);

				  StatusBar1->SimpleText ="Отчетный период: "+IntToStr(god)+" год, "+IntToStr(kvartal)+" квартал";
				  if (pr==1) InsertLog("Расчет рейтинга по работнику с таб.№ = "+DM->qReiting->FieldByName("tn")->AsString+" не выполнен");
				  else InsertLog("Расчет рейтинга по всем работникам не выполнен");
				  Abort();
				}

			  //Проверка на наличие 20% от количества в красной зоне
			  kol_kr_zona+= DM->qObnovlenie2->RowsAffected;

			  DM->qRaschet->Refresh();

			}
		  //********************************************************************
		  //Обновление зеленой зоны пока не будет 20% от числа
		  kol_zl_zona = 0;
		  while (kol_zl_zona<DM->qObnovlenie->FieldByName("zona")->AsInteger && DM->qRaschet->FieldByName("kol_zex")->AsInteger>0)
			{

			  //Проверка превышает ли количество работников с максимальным значением 20%
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
				  Application->MessageBox(("Возникла ошибка при рассчете рейтинга (таблица REIT_RUK) "+E.Message).c_str(),L"Ошибка",
										  MB_OK+MB_ICONERROR);

				  StatusBar1->SimpleText ="Отчетный период: "+IntToStr(god)+" год, "+IntToStr(kvartal)+" квартал";
				  if (pr==1) InsertLog("Расчет рейтинга по работнику с таб.№ = "+DM->qReiting->FieldByName("tn")->AsString+" не выполнен");
				  else InsertLog("Расчет рейтинга по всем работникам не выполнен");
				  Abort();
				}

			  if (DM->qObnovlenie2->FieldByName("kol")->AsInteger+kol_zl_zona<=DM->qObnovlenie->FieldByName("zona")->AsInteger)
				{
				  //Обновить зеленую зону
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
					  Application->MessageBox(("Возникла ошибка при рассчете рейтинга (таблица REIT_RUK) "+E.Message).c_str(),L"Ошибка",
										  MB_OK+MB_ICONERROR);

					  StatusBar1->SimpleText ="Отчетный период: "+IntToStr(god)+" год, "+IntToStr(kvartal)+" квартал";
					  if (pr==1) InsertLog("Расчет рейтинга по работнику с таб.№ = "+DM->qReiting->FieldByName("tn")->AsString+" не выполнен");
					  else InsertLog("Расчет рейтинга по всем работникам не выполнен");
					  Abort();
					}

				 //Проверка на наличие 20% от количества в красной зоне
				 kol_zl_zona+= DM->qObnovlenie2->RowsAffected;

				 DM->qRaschet->Refresh();

				}
			  else
				{
				  kol_zl_zona+= DM->qObnovlenie2->FieldByName("kol")->AsInteger;
				}

			}
		  //********************************************************************
		  //Обновление желтой зоны
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
				  Application->MessageBox(("Возникла ошибка при рассчете рейтинга (таблица REIT_RUK) "+E.Message).c_str(),L"Ошибка",
											  MB_OK+MB_ICONERROR);

				  StatusBar1->SimpleText ="Отчетный период: "+IntToStr(god)+" год, "+IntToStr(kvartal)+" квартал";
				  if (pr==1) InsertLog("Расчет рейтинга по работнику с таб.№ = "+DM->qReiting->FieldByName("tn")->AsString+" не выполнен");
				  else InsertLog("Расчет рейтинга по всем работникам не выполнен");
				  Abort();
				}
			}

			Application->MessageBox(L"Расчет рейтинга выполнен успешно! ",L"Расчет рейтинга",
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
//Формирование итогового отчета с подчиненными
void __fastcall TMain::N15Click(TObject *Sender)
{
  OtchetExcelItog(0);
}
//---------------------------------------------------------------------------
//Формирование итогового отчета без подчиненных
void __fastcall TMain::jjjj1Click(TObject *Sender)
{
  OtchetExcelItog(1);
}
//---------------------------------------------------------------------------
//Формирование итогового отчета
void __fastcall TMain::OtchetExcelItog(int otchet)
{
  AnsiString Sql, sFile;
  int i,n;
  Variant AppEx,Sh;

  StatusBar1->SimpleText ="  Идет формирование итогового отчета...";

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
	  Application->MessageBox(("Возникла ошибка при выборке данных из таблицы по рейтингованию REIT_RUK "+E.Message).c_str(),L"Ошибка",
							  MB_OK+MB_ICONERROR);

	  StatusBar1->SimpleText ="  Отчетный период: "+IntToStr(god)+" год, "+IntToStr(kvartal)+" квартал";
	  Abort();
	}

  if (DM->qObnovlenie->RecordCount!=0)
	{


  Cursor = crHourGlass;
  ProgressBar->Position = 0;
  ProgressBar->Visible = true;
  ProgressBar->Max=DM->qObnovlenie->RecordCount;

  // инициализируем Excel, открываем этот шаблон
  try
	{
	  AppEx=CreateOleObject("Excel.Application");
	}
  catch (...)
    {
      Application->MessageBox(L"Невозможно открыть Microsoft Excel!"
							  " Возможно это приложение на компьютере не установлено.",L"Ошибка",MB_OK+MB_ICONERROR);
	  StatusBar1->SimpleText ="  Отчетный период: "+IntToStr(god)+" год, "+IntToStr(kvartal)+" квартал";
	  Cursor = crDefault;
	  ProgressBar->Visible = false;
	  Abort();
    }

  //Если возникает ошибка во время формирования отчета
  try
    {
      try
        {
		  AppEx.OlePropertySet("AskToUpdateLinks",false);
          AppEx.OlePropertySet("DisplayAlerts",false);

          //Создание папки, если ее не существует
          ForceDirectories(WorkPath);

		  //Копируем шаблон файла в Мои документы
		  CopyFile(WideString(Path+"\\RTF\\itogovaya_tablica.xlsx").c_bstr(), WideString(WorkPath+"\\Итоговая таблица построения рейтинга РПСиТС.xlsx").c_bstr(), false);
		  //sFile = WorkPath+"\\Итоговая таблица построения рейтинга РПСиТС.xlsx";

		  AppEx.OlePropertyGet("WorkBooks").OleProcedure("Open",WideString(WorkPath+"\\Итоговая таблица построения рейтинга РПСиТС.xlsx").c_bstr());  //открываем книгу, указав её имя
		  Sh=AppEx.OlePropertyGet("WorkSheets",1);                               //выбираем № активного листа книги
		  //Sh=AppEx.OlePropertyGet("WorkSheets","Расчет");                      //выбираем лист по наименованию
		}
	  catch(...)
		{
		  Application->MessageBox(L"Ошибка открытия книги Microsoft Excel!",L"Ошибка",MB_OK+MB_ICONERROR);
		  StatusBar1->SimpleText ="  Отчетный период: "+IntToStr(god)+" год, "+IntToStr(kvartal)+" квартал";
		  Cursor = crDefault;
		  ProgressBar->Visible = false;
		  Abort();
        }


	  i=1;
	  n=8;
      //вывод даты
	  Sh.OlePropertyGet("Cells",2,1).OlePropertySet("Value", WideString(IntToStr(kvartal)+" квартал, "+IntToStr(god) +" год" ));

	  //Вывод данных в шаблон
      Variant Massiv;
	  Massiv = VarArrayCreate(OPENARRAY(int,(0,27)),varVariant); //массив на 26 элементов

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


		  Sh.OlePropertyGet("Range", WideString("A" + IntToStr(n) + ":Z" + IntToStr(n))).OlePropertySet("Value", Massiv); //строка с данными с ячейки A по ячейку Z
		  //	Sh.OlePropertyGet("Range", WideString("A8:Z30")).OlePropertySet("Value", Massiv); //строка с данными с ячейки A по ячейку Z


          //Разукрасить колонку итог
		  if (DM->qObnovlenie->FieldByName("reit")->AsInteger==1) Sh.OlePropertyGet("Cells",n,26).OlePropertyGet("Interior").OlePropertySet("Color",0x00D6EFE4);
		  else if (DM->qObnovlenie->FieldByName("reit")->AsInteger==2) Sh.OlePropertyGet("Cells",n,26).OlePropertyGet("Interior").OlePropertySet("Color",0x00C2EAF5);
		  else if (DM->qObnovlenie->FieldByName("reit")->AsInteger==3) Sh.OlePropertyGet("Cells",n,26).OlePropertyGet("Interior").OlePropertySet("Color",0x00ECEEFF);
		  else Sh.OlePropertyGet("Cells",n,26).OlePropertyGet("Interior").OlePropertySet("Color",clWhite);


		  i++;
		  n++;
		  DM->qObnovlenie->Next();
          ProgressBar->Position++;
		}

      //рисуем сетку
	  Sh.OlePropertyGet("Range",WideString("A8:Z"+IntToStr(n-1))).OlePropertyGet("Borders").OlePropertySet("LineStyle", 1);
																													   //xlContinuous
     // Sh.OleProcedure("SaveAs",vAsCurDir1.c_str());
     AppEx.OlePropertyGet("WorkBooks",1).OleFunction("Save");


      //Закрыть книгу Excel с шаблоном для вывода информации
     // AppEx.OlePropertyGet("WorkBooks",1).OleProcedure("Close");
	  Application->MessageBox(L"Отчет в Excel успешно сформирован!", L"Формирование отчета",
							   MB_OK+MB_ICONINFORMATION);
	  //AppEx.OlePropertyGet("WorkBooks").OleProcedure("Open",vAsCurDir1.c_str());
	  AppEx.OlePropertySet("Visible",true);
	  AppEx.OlePropertySet("AskToUpdateLinks",true);
	  AppEx.OlePropertySet("DisplayAlerts",true);

	  StatusBar1->SimpleText= "  Формирование отчета выполнено.";
	}
  catch(...)
	{
	  //Закрыть открытое приложение Excel
	  AppEx.OleProcedure("Quit");
	  AppEx = Unassigned;
	}


  ProgressBar->Position=0;
  ProgressBar->Visible = false;

	}
  else
	{
	  Application->MessageBox(("Нет работников прошедших рейтингование за "+IntToStr(god)+" год, "+IntToStr(kvartal)+" квартал").c_str(),L"Предупреждение",
							  MB_OK+MB_ICONINFORMATION);

	  StatusBar1->SimpleText ="  Отчетный период: "+IntToStr(god)+" год, "+IntToStr(kvartal)+" квартал";
	}

  //***************************************************************************
  //Отчет по не рейтингованным работникам
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
	  Application->MessageBox(("Возникла ошибка при выборке данных из таблицы по рейтингованию REIT_RUK "+E.Message).c_str(),L"Ошибка",
							  MB_OK+MB_ICONERROR);

	  StatusBar1->SimpleText ="  Отчетный период: "+IntToStr(god)+" год, "+IntToStr(kvartal)+" квартал";
	  Abort();
	}

  if (DM->qObnovlenie->RecordCount==0)
	{
	  Application->MessageBox(("Все работники прошли рейтингование за "+IntToStr(god)+" год, "+IntToStr(kvartal)+" квартал").c_str(),L"Предупреждение",
							  MB_OK+MB_ICONINFORMATION);

	  StatusBar1->SimpleText ="  Отчетный период: "+IntToStr(god)+" год, "+IntToStr(kvartal)+" квартал";
	  Abort();
	}

  StatusBar1->SimpleText ="  Идет формирование отчета по не рейтингованным работникам...";
  Cursor = crHourGlass;
  ProgressBar->Position = 0;
  ProgressBar->Visible = true;
  ProgressBar->Max=DM->qObnovlenie->RecordCount;

    //Открытие файла данных для записи не обновленных данных
  if (!rtf_Open((TempPath + "\\ne_reit.txt").c_str()))
	{
	  MessageBox(Handle,L"Ошибка открытия файла данных",L"Ошибка",8192);
	  Abort();
	}

  rtf_Out("data", DateTimeToStr(Now()),0);

  while (!DM->qObnovlenie->Eof)
	{
	  //Формирование отчета по необновленным записям
	  rtf_Out("zex", DM->qObnovlenie->FieldByName("zex")->AsString,1);
	  rtf_Out("tn", DM->qObnovlenie->FieldByName("tn")->AsString,1);
	  rtf_Out("dolg", DM->qObnovlenie->FieldByName("dolg")->AsString,1);
	  rtf_Out("fio", DM->qObnovlenie->FieldByName("fio")->AsString,1);
	  rtf_Out("ocen", DM->qObnovlenie->FieldByName("ocenka")->AsString,1);

	  if(!rtf_LineFeed())
		{
		  MessageBox(Handle,L"Ошибка записи в файл данных",L"Ошибка",8192);
		  if (!rtf_Close()) MessageBox(Handle,L"Ошибка закрытия файла данных",L"Ошибка",8192);
		  return;
		}

		DM->qObnovlenie->Next();
		ProgressBar->Position++;
	 }

  if(!rtf_Close())
	{
	  MessageBox(Handle,L"Ошибка закрытия файла данных", L"Ошибка", 8192);
	  return;
	}

  //Формирование отчета в Word
  StatusBar1->SimpleText = "  Формирование отчета по не рейтингованным работникам...";

  //Создание папки, если ее не существует
  ForceDirectories(WorkPath);

  int istrd;
  try
	{
	  rtf_CreateReport(TempPath + "\\ne_reit.txt", Path+"\\RTF\\ne_reit.rtf",
					   WorkPath+"\\Отчет по не рейтингованным работникам.doc",NULL,&istrd);

	  WinExec(("\""+ WordPath+"\"\""+WorkPath+"\\Отчет по не рейтингованным работникам.doc\"").c_str(),SW_MAXIMIZE);
	}
  catch(RepoRTF_Error E)
	{
	  Application->MessageBox(("Ошибка формирования отчета:"+ String(E.Err)+
								 "\nСтрока файла данных:"+IntToStr(istrd)).c_str(),
								 L"Ошибка",
								 MB_OK+MB_ICONERROR);


	}

  DeleteFile(TempPath+"\\ne_reit.txt");

  Cursor = crDefault;
  ProgressBar->Position=0;
  ProgressBar->Visible = false;
  StatusBar1->SimpleText ="  Отчетный период: "+IntToStr(god)+" год, "+IntToStr(kvartal)+" квартал";


}
//---------------------------------------------------------------------------


void __fastcall TMain::DBGridEh1DrawColumnCell(TObject *Sender, const TRect &Rect,
		  int DataCol, TColumnEh *Column, TGridDrawState State)
{
  if (Prava=="ocen") {
	switch  (DM->qReiting->FieldByName("reit")->AsInteger)
	  {
		case 1: //зеленый
				((TDBGridEh *) Sender)->Canvas->Brush->Color = TColor(0x00D6EFE4);//0x00A3F1D1);//clInfoBk;
				((TDBGridEh *) Sender)->Canvas->Font->Color= clBlack;
				((TDBGridEh *) Sender)->DefaultDrawColumnCell(Rect, DataCol, Column, State);
		break;

		case 2: //желтый
				((TDBGridEh *) Sender)->Canvas->Brush->Color = TColor(0x00C2EAF5);//0x00A3F1D1);//clInfoBk;
				((TDBGridEh *) Sender)->Canvas->Font->Color= clBlack;
				((TDBGridEh *) Sender)->DefaultDrawColumnCell(Rect, DataCol, Column, State);
		break;

		case 3: //красный
				((TDBGridEh *) Sender)->Canvas->Brush->Color = TColor(0x00ECEEFF);//0x00A3F1D1);//clInfoBk;
				((TDBGridEh *) Sender)->Canvas->Font->Color= clBlack;
				((TDBGridEh *) Sender)->DefaultDrawColumnCell(Rect, DataCol, Column, State);
		break;

		default:
				((TDBGridEh *) Sender)->Canvas->Brush->Color = clWhite;//0x00A3F1D1);//clInfoBk;
				((TDBGridEh *) Sender)->Canvas->Font->Color= clBlack;
				((TDBGridEh *) Sender)->DefaultDrawColumnCell(Rect, DataCol, Column, State);
	  }

	// выделение цветом активной записи
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
  //Расположение кнопок по центру

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





