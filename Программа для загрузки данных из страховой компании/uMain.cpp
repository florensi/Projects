//---------------------------------------------------------------------------
#define NO_WIN32_LEAN_AND_MEAN
#include <stdio.h>
#include <vcl.h>
//#include <utilcls.h>
#pragma hdrstop

#include "uMain.h"
#include "uDM.h"
#include "RepoRTFM.h"
#include "RepoRTFO.h"
#include "FuncUser.h"
#include "uData.h"

//#include "dstring.h"

//---------------------------------------------------------------------------
#pragma package(smart_init)
#pragma link "DBGridEh"
#pragma resource "*.dfm"
TMain *Main;

Variant AppEx, Sh, AppEx1, Sh1;

const AnsiString Mes[]={"январь","февраль","март","апрель","май","июнь","июль",
                        "август","сентябрь","октябрь","ноябрь","декабрь"};
//---------------------------------------------------------------------------
__fastcall TMain::TMain(TComponent* Owner)
        : TForm(Owner)
{
}

Variant toExcel(Variant AppEx,const char *Exc, int off, double data)
{
  try {
    AppEx.OlePropertyGet("Range", Exc).OlePropertyGet("Offset", off).OlePropertySet("Value", data);
  } catch(...) { ; }
}/* toExcel() */
//---------------------------------------------------------------------------
Variant toExcel(Variant AppEx,const char *Exc, int off, String data)
{
  try {
    AppEx.OlePropertyGet("Range", Exc).OlePropertyGet("Offset", off).OlePropertySet("Value", data.c_str());
  } catch(...) { ; }
}/* toExcel() */
//---------------------------------------------------------------------------
Variant  toExcel(Variant AppEx,const char *Exc, double data)
{
  try {
    AppEx.OlePropertyGet("Range", Exc).OlePropertySet("Value", data);
  } catch(...) { ; }
}/* toExcel() */

//---------------------------------------------------------------------------
Variant  toExcel(Variant AppEx,const char *Exc, int data)
{
  try {
    AppEx.OlePropertyGet("Range", Exc).OlePropertySet("Value", data);
  } catch(...) { ; }
}/* toExcel() */

//---------------------------------------------------------------------------
Variant  toExcel(Variant AppEx,const char *Exc, AnsiString data)
{
  try {
    Variant  cur = AppEx.OlePropertyGet("Range", Exc);
    cur.OlePropertySet("Value", data.c_str());
  } catch(...) { ; }
}/* toExcel() */
//---------------------------------------------------------------------------
//---------------------------------------------------------------------------
 
void __fastcall TMain::FormCreate(TObject *Sender)
{
  int Prava;

  Path = GetCurrentDir();
  FindWordPath();

  // Получение данных о пользователе из домена
  TStringList *SL_Groups = new TStringList();


  // Переменные UserName, DomainName, UserFullName должны быть объявлены как AnsiString
  if (!GetFullUserInfo(UserName, DomainName, UserFullName))
    {
      MessageBox(Handle,"Ошибка получения данных о пользователе","Ошибка",8208);
      Application->Terminate();
      Abort();
    }

  //получение групп доступа из АД
  if (!GetUserGroups(UserName, DomainName, SL_Groups))
    {
      MessageBox(Handle,"Ошибка получения данных о пользователе","Ошибка",8208);
      Application->Terminate();
      Abort();
    }

  //проверка на доступ к услуге
  if ((SL_Groups->IndexOf("mmk-itsvc-hstr-admin")<=-1) && (SL_Groups->IndexOf("mmk-itsvc-hstr")<=-1))
    {
      MessageBox(Handle,"У вас нет прав для работы с\n с программой 'Страхование жизни и пенсий'!!!","Права доступа",8208);
      Application->Terminate();
      Abort();
    }

  if (UserFullName.SubString(1,3)=="rmz")
    {
      ana = 4; //МРМЗ
    }
  else
    {
      ana = 1; //ММК им.Ильича
    }

  //проверка прав
  //1- полный доступ
 // if (SL_Groups->IndexOf("mmk-itsvc-hstr-01")>-1)
 //   {
      //Полный доступ
      N9->Visible = true;       //Редактирование
      N3->Visible = true;       //Загрузка данных по договорам
      N4->Visible = true;       //Подготовка для УИТ
      N22->Visible = true;      //Справка
    //  N15->Visible = false;     //Отчет "Управление с/х"

      if (ana==4)
        {
          N13->Visible = false;  //Отчет "Для РО"
          N14->Visible = false;  //Отчет "Для ОК"
          N16->Visible = false;  //Отчет "Для Страховой"
          N17->Visible = true;   //Отчет "Для РМЗ"
        }
      else
        {
          N13->Visible = true;   //Отчет "Для РО"
          N14->Visible = true;   //Отчет "Для ОК"
          N16->Visible = true;   //Отчет "Для Страховой"
          N17->Visible = false;  //Отчет "Для РМЗ"
        }
/*    }
  //2- для операторов (отчет по УСХ)
  else if (SL_Groups->IndexOf("mmk-itsvc-hstr-02")>-1)
    {
      //Для операторов
      N9->Visible=false;
      N3->Visible=false;
      N4->Visible=false;
      N22->Visible=false;
      N13->Visible=false;
      N14->Visible=false;
      N16->Visible=false;
      N17->Visible=false;
      //N15->Visible=true;
    }
  else
    {
      Application->MessageBox("Не установлены права доступа(УКИЛ, ОУЗП) для работы с программой АСПД 'Средний заработок'!!!","Права доступа",
                              MB_OK+MB_ICONERROR);
      Application->Terminate();
      Abort();

    }  */
 /*
  // Получение данных о пользователе из домена
  // Переменные UserName, DomainName, UserFullName должны быть объявлены как AnsiString
  if (!GetUserInfo(UserName, DomainName, UserFullName))
    {
      MessageBox(Handle,"Ошибка получения данных о пользователе","Ошибка",8208);
      Application->Terminate();
      Abort();
    }

  // Получение прав доступа из таблицы users_ro
  DM->qRO_user->Close();
  DM->qRO_user->SQL->Clear();
  DM->qRO_user->SQL->Add("select VU_859, tn, factory from USERS_RO@SLST5 where domain=" + QuotedStr(DomainName) + " and userro=" + QuotedStr(UserName));
  DM->qRO_user->Open();

  if (!DM->qRO_user->RecordCount)
    {
      MessageBox(Handle,("Нет данных о пользователе " + UserName).c_str(),"Ошибка",8208);
      Application->Terminate();
      Abort();
    }

  if (DM->qRO_user->FieldByName("VU_859")->AsString.IsEmpty())
    {
      MessageBox(Handle,("Нет данных о пользователе " + UserName).c_str(),"Ошибка",8208);
      Application->Terminate();
      Abort();
    }

  if (DM->qRO_user->FieldByName("factory")->AsString.IsEmpty()||
      (DM->qRO_user->FieldByName("factory")->AsInteger !=1 &&
       DM->qRO_user->FieldByName("factory")->AsInteger !=4 ))
    {
      MessageBox(Handle,"Не указано предприятие(поле FACTORY) в таблице USERS_RO","Ошибка",8208);
      Application->Terminate();
      Abort();
    }

  Prava = DM->qRO_user->FieldByName("VU_859")->AsInteger;
  TN = DM->qRO_user->FieldByName("tn")->AsString;
  ana = DM->qRO_user->FieldByName("factory")->AsInteger;
  DM->qRO_user->Close();

  switch (Prava)
    {
      case 1:  //Полный доступ
              N9->Visible = true;       //Редактирование
              N3->Visible = true;       //Загрузка данных по договорам
              N4->Visible = true;       //Подготовка для УИТ
              N22->Visible = true;      //Справка
              N15->Visible = false;     //Отчет "Управление с/х"

             if (ana==4)
               {
                 N13->Visible = false;  //Отчет "Для РО"
                 N14->Visible = false;  //Отчет "Для ОК"
                 N16->Visible = false;  //Отчет "Для Страховой"
                 N17->Visible = true;   //Отчет "Для РМЗ"
               }
             else
               {
                 N13->Visible = true;   //Отчет "Для РО"
                 N14->Visible = true;   //Отчет "Для ОК"
                 N16->Visible = true;   //Отчет "Для Страховой"
                 N17->Visible = false;  //Отчет "Для РМЗ"
               }

      break;

      case 2:  //Для операторов
              N9->Visible=false;
              N3->Visible=false;
              N4->Visible=false;
              N22->Visible=false;
              N13->Visible=false;
              N14->Visible=false;
              N16->Visible=false;
              N17->Visible=false;
              N15->Visible=true;
      break;


      default:
        Application->MessageBox("Несуществующие права доступа", "Ошибка",
                                     MB_OK + MB_ICONERROR);
        Application->Terminate();
        Abort();
    }

  */


  if (!GetMyDocumentsDir(DocPath))
    {
      MessageBox(Handle,"Ошибка доступа к папке документов","Ошибка",8208);
      Application->Terminate();
      Abort();
    }

  if (!GetTempDir(TempPath))
    {
      MessageBox(Handle,"Ошибка доступа к временной папке","Ошибка",8208);
      Application->Terminate();
      Abort();
    }


 WorkPath = DocPath + "\\Загрузка данных из страховой компании";


 // Создание ProgressBar на StatusBar
      ProgressBar = new TProgressBar ( StatusBar1 );
      ProgressBar->Parent = StatusBar1;
      ProgressBar->Position = 0;
      ProgressBar->Left = StatusBar1->Panels->Items[0]->Width + 3;
      ProgressBar->Top = StatusBar1->Height/6;
      ProgressBar->Height = StatusBar1->Height-3;
      ProgressBar->Visible = false;
}
//---------------------------------------------------------------------------

// Проверка на значение цеха в Excel-файле
bool  __fastcall TMain::Proverka(AnsiString zex)
{
   try {
    StrToInt(zex);
  }
  catch (...) {
    return false;
  }
  return true;

}
//---------------------------------------------------------------------------

// Проверка правильности цеха, тн и суммы страхования на превышение 15%
 void __fastcall TMain::ProverkaInfoExcel()
{
  AnsiString Sql, Sql1, Sql2, inn, nn, fam, n_doc, nnom, name;
  int fl=0, pr_sum=0, pr_inn=0, pr_dv=0;
  int i=1;
  double sum;
  TSearchRec SearchRecord;  //Для поиска файла


  /*tn - таб.№,
    Row - общее количество занятых строк в документе
    sum - сумма удержаний на страховку
    fl - признак формирования отчета (fl=1 - формировать)
    name - имя Excel файла
    FileName - полный путь к файлу с его именем
    Dir2 - путь к выбранной папке
    pr_sum - признак вывода шапки таблицы и заголовка по превыш. суммы (pr_sum = 0 выводить)
    pr_inn - признак вывода шапки таблицы и заголовка по отсут. индификац.№ (pr_inn = 0 выводить)
    pr_dv - признак вывода шапки таблицы и заголовка по двойным записям (pr_dv = 0 выводить)*/


   //Окно выбора директории к папке
  if (!SelectDirectory("Select directory",WideString(""),Dir2))
    {
      Abort();
    }

   //Поиск файла Excel
   switch(im_fl)
     {
       case 1 : if (FindFirst(Dir2 + LowerCase("\\Новые ММК.xls"), faAnyFile, SearchRecord)==0 )
                  {
                    name = LowerCase("\\Новые ММК.xls");
                  }
                else if (FindFirst (Dir2 + LowerCase("\\Новые ММК.xlsx"), faAnyFile, SearchRecord)==0)
                  {
                    name = LowerCase("\\Новые ММК.xlsx");
                  }
                else
                  {
                    Application->MessageBox("Не найден файл для загрузки данных. \nВозможно указано НЕВЕРНОЕ ИМЯ файла \nили файл не найден в данной папке. ",
                                           "Ошибка загрузки данных", MB_OK + MB_ICONERROR);
                    Abort();
                  }

       break;
       case 2 :  if (FindFirst(Dir2 + LowerCase("\\Изменения(гривна).xls"), faAnyFile, SearchRecord)==0 )
                   {
                     name = LowerCase("\\Изменения(гривна).xls");
                   }
                 else if (FindFirst (Dir2 + LowerCase("\\Изменения(гривна).xlsx"), faAnyFile, SearchRecord)==0)
                   {
                     name = LowerCase("\\Изменения(гривна).xlsx");
                   }
                 else
                   {
                     Application->MessageBox("Не найден файл для загрузки данных. \nВозможно указано НЕВЕРНОЕ ИМЯ файла \nили файл не найден в данной папке. ",
                                           "Ошибка загрузки данных", MB_OK + MB_ICONERROR);
                     Abort();
                   }
       break;
       case 3 :  if (FindFirst(Dir2 + LowerCase("\\Изменения(валюта).xls"), faAnyFile, SearchRecord)==0 )
                   {
                     name = LowerCase("\\Изменения(валюта).xls");
                   }
                 else if (FindFirst (Dir2 + LowerCase("\\Изменения(валюта).xlsx"), faAnyFile, SearchRecord)==0)
                   {
                     name = LowerCase("\\Изменения(валюта).xlsx");
                   }
                 else
                   {
                     Application->MessageBox("Не найден файл для загрузки данных. \nВозможно указано НЕВЕРНОЕ ИМЯ файла \nили файл не найден в данной папке. ",
                                            "Ошибка загрузки данных", MB_OK + MB_ICONERROR);
                     Abort();
                   }
       break;
       case 4 :  if (FindFirst(Dir2 + LowerCase("\\Изменения(курс).xls"), faAnyFile, SearchRecord)==0 )
                   {
                     name = LowerCase("\\Изменения(курс).xls");
                   }
                 else if (FindFirst (Dir2 + LowerCase("\\Изменения(курс).xlsx"), faAnyFile, SearchRecord)==0)
                   {
                     name = LowerCase("\\Изменения(курс).xlsx");
                   }
                 else
                   {
                     Application->MessageBox("Не найден файл для загрузки данных. \nВозможно указано НЕВЕРНОЕ ИМЯ файла \nили файл не найден в данной папке. ",
                                           "Ошибка загрузки данных", MB_OK + MB_ICONERROR);
                     Abort();
                   }
       break;
       case 5 :  if (FindFirst(Dir2 + LowerCase("\\Новые ВР.xls"), faAnyFile, SearchRecord)==0 )
                   {
                     name = LowerCase("\\Новые ВР.xls");
                   }
                 else if (FindFirst (Dir2 + LowerCase("\\Новые ВР.xlsx"), faAnyFile, SearchRecord)==0)
                   {
                     name = LowerCase("\\Новые ВР.xlsx");
                   }
                 else
                   {
                     Application->MessageBox("Не найден файл для загрузки данных. \nВозможно указано НЕВЕРНОЕ ИМЯ файла \nили файл не найден в данной папке. ",
                                           "Ошибка загрузки данных", MB_OK + MB_ICONERROR);
                     Abort();
                   }
       break;
       case 7 : if (FindFirst(Dir2 + LowerCase("\\Новые пенсионное.xls"), faAnyFile, SearchRecord)==0 )
                  {
                    name = LowerCase("\\Новые пенсионное.xls");
                  }
                else if (FindFirst (Dir2 + LowerCase("\\Новые пенсионное.xlsx"), faAnyFile, SearchRecord)==0)
                  {
                    name = LowerCase("\\Новые пенсионное.xlsx");
                  }
                else
                  {
                    Application->MessageBox("Не найден файл для загрузки данных. Возможно указано НЕВЕРНОЕ ИМЯ файла (должно быть 'Новые пенсионное.xls' или 'Новые пенсионное.xlsx') или файл не найден в данной папке.",
                                           "Ошибка загрузки данных", MB_OK + MB_ICONERROR);
                    Abort();
                  }
       break;

     }

  FileName = Dir2 + name;  //Путь к файлу Excel
  FindClose(SearchRecord);   //освобождает ресурсы, взятые процессом поиска
     
  StatusBar1->SimpleText = "";

   // инициализируем Excel, открываем этот шаблон
  try
    {
      //проверяем, нет ли запущенного Excel
      Excel = GetActiveOleObject("Excel.Application");
    }
  catch(...)
    {
      try
        {
          Excel = CreateOleObject("Excel.Application");
        }
      catch (...)
        {
          Application->MessageBox("Невозможно открыть Microsoft Excel!"
          " Возможно это приложение на компьютере не установлено.","Ошибка",MB_OK+MB_ICONERROR);
          Abort();
        }
    }

  try
    {
      Book = Excel.OlePropertyGet("Workbooks").OlePropertyGet("Open", FileName.c_str());
      Sheet = Book.OlePropertyGet("Worksheets", 1);
    }
  catch(...)
    {
      Application->MessageBox("Ошибка открытия книги Microsoft Excel!","Ошибка",MB_OK + MB_ICONERROR);
    }


  //Excel.OlePropertySet("Visible",true);


  //Определяет количество занятых строк в документе
  Row = Sheet.OlePropertyGet("UsedRange").OlePropertyGet("Rows").OlePropertyGet("Count");


  // Открываем файл данных для формирования отчета по несуществующим цех и тн и превышения
  if (!rtf_Open((TempPath + "\\otchet.txt").c_str()))
    {
      MessageBox(Handle,"Ошибка открытия файла данных","Ошибка",8192);
    }
  else
    {
      Main->Cursor = crHourGlass;
      StatusBar1->SimplePanel = true;    // 2 панели на StatusBar1
      StatusBar1->SimpleText=" Выполняется проверка данных...";
      ProgressBar->Visible = true;
      ProgressBar->Position = 0;
      ProgressBar->Max = Row;


      for ( i ; i<Row+1; i++)
        {                                                      
          nn = Excel.OlePropertyGet("Cells",i,1);
          inn = Excel.OlePropertyGet("Cells",i,5);
          ProgressBar->Position++;


          // Выбор строк необходимых для загрузки из Excel
          if (nn.IsEmpty() || !Proverka(nn) || inn.IsEmpty())  continue;

            sum = Excel.OlePropertyGet("Cells",i,8);
            fam = TrimRight(""+Excel.OlePropertyGet("Cells",i,2)+" "+Excel.OlePropertyGet("Cells",i,3)+" "+Excel.OlePropertyGet("Cells",i,4));
            n_doc = Excel.OlePropertyGet("Cells",i,9);


//Проверка на несколько записей в sap_osn_sved и sap_sved_uvol и наличие инд.№
//******************************************************************************
            Sql1 = "select tn_sap, numident from sap_osn_sved where numident=:pnumident                \
                    union all                                                                          \
                    select tn_sap, numident from sap_sved_uvol                                         \
                    where substr(to_char(dat_job,'dd.mm.yyyy'),4,7)='"+(DM->mm<10 ? "0"+IntToStr(DM->mm) : IntToStr(DM->mm))+"."+DM->yyyy+"' and numident=:pnumi";

            try
              {
                DM->qObnovlenie->Close();
                DM->qObnovlenie->SQL->Clear();
                DM->qObnovlenie->SQL->Add(Sql1);
                DM->qObnovlenie->Parameters->ParamByName("pnumident")->Value =inn;
                DM->qObnovlenie->Parameters->ParamByName("pnumi")->Value =inn;
                DM->qObnovlenie->Open();
              }
            catch(...)
              {
                Application->MessageBox("Невозможно получить данные из картотеки работников(SAP_OSN_SVED, SAP_SVED_UVOL)","Ошибка",MB_OK + MB_ICONERROR);
                Abort();
              }

            if (DM->qObnovlenie->RecordCount>1)
              {
                 pr_sum=0;
                 pr_inn=0;
                //Вывод в отчет двойных записей
//******************************************************************************
                //Вывод наименования и шапки таблицы
                if (DM->qObnovlenie->RecordCount>1 && pr_dv==0)
                  {
                    rtf_Out("z", " ",3);
                    if(!rtf_LineFeed())
                      {
                        MessageBox(Handle,"Ошибка записи в файл данных","Ошибка",8192);
                        if (!rtf_Close()) MessageBox(Handle,"Ошибка закрытия файла данных","Ошибка",8192);
                        return;
                      }
                   }
                //Вывод записей в отчет
                rtf_Out("inn", inn,4);
                rtf_Out("fio",fam,4);

                if(!rtf_LineFeed())
                  {
                    MessageBox(Handle,"Ошибка записи в файл данных","Ошибка",8192);
                    if (!rtf_Close()) MessageBox(Handle,"Ошибка закрытия файла данных","Ошибка",8192);
                    return;
                  }
                fl=1;
                pr_dv=1;

              }
            else
              {

                // Вывод несуществующего инд.№ в отчет
//******************************************************************************
                if (DM->qObnovlenie->RecordCount==0)
                  {
                     pr_sum=0;
                     pr_dv=0;
                   //Вывод наименования и шапки таблицы
                    if (DM->qObnovlenie->RecordCount==0 && pr_inn==0)
                      {
                        rtf_Out("z", " ",1);
                        if(!rtf_LineFeed())
                          {
                            MessageBox(Handle,"Ошибка записи в файл данных","Ошибка",8192);
                            if (!rtf_Close()) MessageBox(Handle,"Ошибка закрытия файла данных","Ошибка",8192);
                            return;
                          }
                      }

                    rtf_Out("inn",inn,2);
                    rtf_Out("fio",fam,2);

                    if(!rtf_LineFeed())
                      {
                        MessageBox(Handle,"Ошибка записи в файл данных","Ошибка",8192);
                        if (!rtf_Close()) MessageBox(Handle,"Ошибка закрытия файла данных","Ошибка",8192);
                        return;
                      }

                    fl=1;

                    pr_inn=1;

                  }
                else
                  {
                    /*//Проверка на превышение удерживаемой суммы свыше 15%
//******************************************************************************
                    Sql2 = "select (sum(decode(typs,3,sum*-1,sum))*15/100) sum from slst"+(DM->mm2 < 10 ? "0" + IntToStr(DM->mm2) : IntToStr(DM->mm2))+ DM->yyyy2 + " \
                            where klus="+nnom+" \
                            and typs in (1,3,5) \
                            and vo<800";


                    DM->qObnovlenie->Close();
                    DM->qObnovlenie->SQL->Clear();
                    DM->qObnovlenie->SQL->Add(Sql2);
                    DM->qObnovlenie->Open();

                    if (DM->qObnovlenie->FieldByName("sum")->AsString.IsEmpty())
                      {
                        if (Application->MessageBox(("Нет суммы за прошлый месяц\nцех="+zex+" таб.№="+tn+" \nФИО="+fam+" сумма="+FloatToStrF(sum,ffFixed,20,2)+" \nЗагрузить запись в таблицу?").c_str(),
                                                    "Превышение",MB_YESNO + MB_ICONINFORMATION)==IDNO)
                          {
                            pr_inn=0;
                            pr_dv=0;
                            // Вывод в отчет если нет суммы за прошлый месяц
//******************************************************************************
                            //Вывод наименования и шапки таблицы
                            if ((sum >= DM->qObnovlenie->FieldByName("sum")->AsFloat) && pr_sum==0)
                              {
                                rtf_Out("zz", " ",3);

                                if(!rtf_LineFeed())
                                  {
                                    MessageBox(Handle,"Ошибка записи в файл данных","Ошибка",8192);
                                    if (!rtf_Close()) MessageBox(Handle,"Ошибка закрытия файла данных","Ошибка",8192);
                                    return;
                                  }
                              }

                            rtf_Out("zex", zex,4);
                            rtf_Out("tn", tn,4);
                            rtf_Out("fio",fam,4);
                            rtf_Out("n_doc",n_doc ,4);
                            rtf_Out("sum","нет суммы прошлого месяца",4);

                            if(!rtf_LineFeed())
                              {
                                MessageBox(Handle,"Ошибка записи в файл данных","Ошибка",8192);
                                if (!rtf_Close()) MessageBox(Handle,"Ошибка закрытия файла данных","Ошибка",8192);
                                return;
                              }

                            fl=1;
                            pr_sum=1;

                        }
                      }
                    else if (sum > DM->qObnovlenie->FieldByName("sum")->AsFloat)
                      {
                        if (Application->MessageBox(("Сумма превышает 15%\nцех="+zex+" таб.№="+tn+" ФИО="+fam+" сумма="+FloatToStrF(sum,ffFixed,20,2)+" \nЗагрузить запись в таблицу?").c_str(),
                                                    "Превышение",MB_YESNO + MB_ICONINFORMATION)==IDNO)
                          {
                            pr_inn=0;
                            pr_dv=0;
                            // Вывод в отчет превышающей 15% суммы
//******************************************************************************
                            //Вывод наименования и шапки таблицы
                            if ((sum > DM->qObnovlenie->FieldByName("sum")->AsFloat) && pr_sum==0)
                              {
                                rtf_Out("zz", " ",3);

                                if(!rtf_LineFeed())
                                  {
                                    MessageBox(Handle,"Ошибка записи в файл данных","Ошибка",8192);
                                    if (!rtf_Close()) MessageBox(Handle,"Ошибка закрытия файла данных","Ошибка",8192);
                                    return;
                                  }
                              }

                            rtf_Out("zex", zex,4);
                            rtf_Out("tn", tn,4);
                            rtf_Out("fio",fam,4);
                            rtf_Out("n_doc",n_doc ,4);
                            rtf_Out("sum",FloatToStrF(sum,ffFixed,20,2),4);

                            if(!rtf_LineFeed())
                              {
                                MessageBox(Handle,"Ошибка записи в файл данных","Ошибка",8192);
                                if (!rtf_Close()) MessageBox(Handle,"Ошибка закрытия файла данных","Ошибка",8192);
                                return;
                              }

                            fl=1;
                            pr_sum=1;
                          }
                      } */
                  } 
              }
        }

      StatusBar1->SimplePanel = false;
      ProgressBar->Visible = false;
      StatusBar1->SimpleText = "Проверка данных выполнена.";
      Main->Cursor = crDefault;

      if(!rtf_Close())
        {
          MessageBox(Handle,"Ошибка закрытия файла данных", "Ошибка", 8192);
          return;
        }


      if (fl==1)
        {
          Excel.OleProcedure("Quit");
          StatusBar1->SimpleText = "Формирование отчета...";
          //Создание папки, если ее не существует
          ForceDirectories(WorkPath);

          int istrd;
          try
            {
              rtf_CreateReport(TempPath +"\\otchet.txt", Path+"\\RTF\\otchet.rtf",
                         WorkPath+"\\Отчет.doc",NULL,&istrd);


              WinExec(("\""+ WordPath+"\"\""+WorkPath+"\\Отчет.doc\"").c_str(),SW_MAXIMIZE);

            }
          catch(RepoRTF_Error E)
            {
              MessageBox(Handle,("Ошибка формирования отчета:"+ AnsiString(E.Err)+
                                 "\nСтрока файла данных:"+IntToStr(istrd)).c_str(),"Ошибка",8192);
            }

          Application->MessageBox(("Проверьте достоверность информации в файле \n \""+FileName+"\" и выполните повторную загрузку").c_str() ," Загрузка новых договоров по ММК",
                                  MB_OK + MB_ICONINFORMATION);
          StatusBar1->SimpleText = "";

          switch (im_fl)
            {
              case 1: InsertLog("Сформирован отчет по Новым договорам ММК: нет данных по ИНН");
              break;
              case 5: InsertLog("Сформирован отчет по Новым внешним договорам: нет данных по ИНН");
              break;
              case 7: InsertLog("Сформирован отчет по Новым договорам по пенсионному страхованию: нет данных по ИНН");
              break;
            }

          Abort();
        }

         DeleteFile(TempPath+"\\otchet.txt");        
    }                                            

   
}
//---------------------------------------------------------------------------

// Обновление изменений по валюте или гривне или курсу
 void __fastcall TMain::UpdateValuta_I_Grivna()
{
  AnsiString Sql, inn, nn, data_s, data_po, fam, n_dogovora, Sql1, sum, kod_dog,
             prich, name_otchet;
  int rec=0, kol=0;
  bool fl=0;

 //Открытие файла данных для записи не обновленных данных
   if (!rtf_Open((TempPath + "\\izmeneniya.txt").c_str()))
     {
       MessageBox(Handle,"Ошибка открытия файла данных","Ошибка",8192);
     }
   else
     {
       //   Sheet.OleProcedure("Activate");
       Sheet = Book.OlePropertyGet("Worksheets", 1);

       int i=1;

       Main->Cursor = crHourGlass;
       StatusBar1->SimplePanel = true;    // 2 панели на StatusBar1
       StatusBar1->SimpleText=" Идет загрузка данных...";

       ProgressBar->Visible = true;
       ProgressBar->Position = 0;
       ProgressBar->Max = Row;

        /*  //Проверка на наличие одинаковых записей цех + тн + МФ
          Sql ="SELECT [Лист1$].* From [Лист1$] ";

          DM->qZagruzka->Close();
                DM->qZagruzka->SQL->Clear();;
                DM->qZagruzka->SQL->Add(Sql);
                DM->qZagruzka->ExecSQL();
          if (DM->qZagruzka->RecordCount>0)
          {ShowMessage("=)");}

              */

   
       for ( i ; i<Row+1; i++)
         {
           nn= Excel.OlePropertyGet("Cells",i,1);
           inn = Excel.OlePropertyGet("Cells",i,5);

           ProgressBar->Position++;

           // Выбор строк необходимых для загрузки из Excel
           if (nn.IsEmpty() || !Proverka(nn) || inn.IsEmpty())  continue;

           
           if (im_fl==7)
             {
               data_s = Excel.OlePropertyGet("Cells",i,6);
               n_dogovora = Excel.OlePropertyGet("Cells",i,10);
               fam = TrimRight(""+Excel.OlePropertyGet("Cells",i,2)+" "+Excel.OlePropertyGet("Cells",i,3)+" "+Excel.OlePropertyGet("Cells",i,4));
               sum = Excel.OlePropertyGet("Cells",i,9);
               kod_dog = Excel.OlePropertyGet("Cells",i,7);
               prich = Excel.OlePropertyGet("Cells",i,12);
             }
           else
             {
               data_s = Excel.OlePropertyGet("Cells",i,6);
               data_po = Excel.OlePropertyGet("Cells",i,7);
               n_dogovora = Excel.OlePropertyGet("Cells",i,11);
               fam = TrimRight(""+Excel.OlePropertyGet("Cells",i,2)+" "+Excel.OlePropertyGet("Cells",i,3)+" "+Excel.OlePropertyGet("Cells",i,4));
               sum = Excel.OlePropertyGet("Cells",i,10);
               kod_dog = Excel.OlePropertyGet("Cells",i,8);
               prich = Excel.OlePropertyGet("Cells",i,12);
             }


           //Добавление цех+тн из sap_osn_sved
           Sql1="select zex, tn_sap, numident from sap_osn_sved where trim(numident)=trim("+ QuotedStr(Excel.OlePropertyGet("Cells",i,5)) +" )   \
                 union all                                                                                            \
                 select zex, tn_sap, numident from sap_sved_uvol                                                           \
                 where substr(to_char(dat_job,'dd.mm.yyyy'),4,7)='"+(DM->mm<10 ? "0"+IntToStr(DM->mm) : IntToStr(DM->mm))+"."+DM->yyyy+"'  \
                 and trim(numident)=trim("+QuotedStr(Excel.OlePropertyGet("Cells",i,5))+")";

           try
             {
               DM->qObnovlenie->Close();
               DM->qObnovlenie->SQL->Clear();
               DM->qObnovlenie->SQL->Add(Sql1);
               DM->qObnovlenie->Open();
             }
           catch(...)
             {
               Application->MessageBox("Ошибка получения данных из картотеки по работникам (SAP_OSN_SVED, SAP_SVED_UVOL)","Ошибка",MB_OK+ MB_ICONERROR);
               Excel.OleProcedure("Quit");
               StatusBar1->SimpleText="";
               Main->Cursor = crDefault;
               Abort();
             }


           Sql = "update VU_859_N set zex="+QuotedStr(DM->qObnovlenie->FieldByName("zex")->AsString)+", \
                                      tn="+DM->qObnovlenie->FieldByName("tn_sap")->AsString;


          /*      if (sum==0 || sum.IsEmpty())
                  {
                    Sql+=", sum="+ QuotedStr(Excel.OlePropertyGet("Cells",i,14))+" , priznak=6";
                  }
                else
                  {
                    Sql+=", sum="+ QuotedStr(Excel.OlePropertyGet("Cells",i,14))+" , priznak=0";
                  }    */

           // Проверка на наличие суммы
           if (LowerCase(prich)=="ВС")
             {
               Sql+=", sum="+ QuotedStr(sum)+" , priznak=4";
             }
           else if (LowerCase(prich)=="Р")
             {
               Sql+=", sum="+ QuotedStr(sum)+" , priznak=6";
             }
           else if (LowerCase(prich).IsEmpty() &&(sum==0 || sum.IsEmpty()))
             {
               Sql+=", sum="+ QuotedStr(sum)+" , priznak=6";
             }
           else
             {
               Sql+=", sum="+ QuotedStr(sum)+" , priznak=0";
             }


           //по гривневым договорам
           if (im_fl==2)
             {
               Sql+=", kod_dogovora=0";
             }
           // по валютным договорам
           if(im_fl==3)
             {
               if (kod_dog =="дол"||kod_dog =="дол.")
                 {
                   Sql+=", kod_dogovora=1";
                 }
               else
                 {
                   Sql+=", kod_dogovora=2";
                 }
             }
           //по внешним договорам
           if (im_fl==6)
             {
               Sql+=", kod_dogovora=3";
             }
           //по пенсионному страхованию
           if (im_fl==7)
             {
               Sql+=", kod_dogovora=4";
             }

           if (!data_po.IsEmpty() && im_fl!=7)
             {
               Sql+= ", data_po="+QuotedStr(data_po);
             }


           Sql+= " where trim(inn) = trim("+ QuotedStr(Excel.OlePropertyGet("Cells",i,5))+ ")\
                   and trim(n_dogovora)=trim(" + QuotedStr(n_dogovora)+")";

           DM->qZagruzka->Close();
           DM->qZagruzka->SQL->Clear();;
           DM->qZagruzka->SQL->Add(Sql);
           try
             {
               DM->qZagruzka->ExecSQL();
             }
           catch(...)
             {
               Application->MessageBox("Ошибка обновления данных по договорам страхования","Ошибка",MB_OK+ MB_ICONERROR);
               Excel.OleProcedure("Quit");
               StatusBar1->SimpleText="";
               ProgressBar->Visible = false;
               Main->Cursor = crDefault;
               Abort();
             }
           rec++;
           kol+=DM->qZagruzka->RowsAffected;

           // Количество обновленных записей
           if (DM->qZagruzka->RowsAffected == 0)
             {
               //Формирование отчета по необновленным записям
               rtf_Out("inn", inn,1);
               rtf_Out("fio",fam,1);
               rtf_Out("n_dogovora",n_dogovora,1);

               if(!rtf_LineFeed())
                 {
                   MessageBox(Handle,"Ошибка записи в файл данных","Ошибка",8192);
                   if (!rtf_Close()) MessageBox(Handle,"Ошибка закрытия файла данных","Ошибка",8192);
                   return;
                 }
               fl=1;  //Признак формирования отчета по необновленным записям
             }
         }

       StatusBar1->SimplePanel = false;
       ProgressBar->Visible = false;
       Main->Cursor = crDefault;

       if(!rtf_Close())
         {
           MessageBox(Handle,"Ошибка закрытия файла данных", "Ошибка", 8192);
           return;
         }

        // Формирование отчета, если есть не обновленные записи
       if (fl==1)
         {
           //Создание папки, если ее не существует
           ForceDirectories(WorkPath);


           switch (im_fl)
             {
               case 2: name_otchet = "Изменения по гривневым договорам.doc";
               break;
               case 3: name_otchet = "Изменения по валютным договорам.doc";
               break;
               case 4: name_otchet = "Изменения по курсу.doc";
               break;
               case 6: name_otchet = "Изменения по ВР.doc";
               break;
               case 7: name_otchet = "Изменения по пенсионным договорам.doc";
               break;
             }


           int istrd;
           try
             {
               rtf_CreateReport(TempPath +"\\izmeneniya.txt", Path+"\\RTF\\izmeneniya.rtf",
                                WorkPath+"\\"+name_otchet,NULL,&istrd);

               WinExec(("\""+ WordPath+"\"\""+WorkPath+"\\"+name_otchet+"\"").c_str(),SW_MAXIMIZE);
             }
           catch(RepoRTF_Error E)
             {
               MessageBox(Handle,("Ошибка формирования отчета:"+ AnsiString(E.Err)+
                                  "\nСтрока файла данных:"+IntToStr(istrd)).c_str(),"Ошибка",8192);
             }

           Application->MessageBox(("Обновлено " + IntToStr(kol) + " из " + IntToStr(rec) + " записей.\nПроверьте достоверность информации в файле \n \""+FileName+"\" и выполните повторную загрузку").c_str() ," Загрузка новых договоров по ММК",
                                        MB_OK + MB_ICONINFORMATION);
           StatusBar1->SimpleText = "";

           switch (im_fl)
             {
               case 2: InsertLog("Сформирован отчет по Изменениям(гривна): обновлено " + IntToStr(kol) + " из " + IntToStr(rec) + " записей.");
               break;
               case 3: InsertLog("Сформирован отчет по Изменениям(валюта): обновлено " + IntToStr(kol) + " из " + IntToStr(rec) + " записей.");
               break;
               case 4: InsertLog("Сформирован отчет по Изменениям курса: обновлено " + IntToStr(kol) + " из " + IntToStr(rec) + " записей.");
               break;
               case 6: InsertLog("Сформирован отчет по Изменениям ВР: обновлено " + IntToStr(kol) + " из " + IntToStr(rec) + " записей.");
               break;
               case 7: InsertLog("Сформирован отчет по Изменениям по пенсионным договорам: обновлено " + IntToStr(kol) + " из " + IntToStr(rec) + " записей.");
               break;
             }

           Abort();
         }

       ob_kol = rec;
       obnov_kol = kol;

       DeleteFile(TempPath+"\\izmeneniya.txt");
       StatusBar1->SimpleText = "Данные обновлены";
       Application->MessageBox(("Обновление данных выполнено успешно =) \nОбновлено " + IntToStr(kol) + " из " + IntToStr(rec)+" записей").c_str(),
                                   "Обновление записей по страхованию жизни",
                                   MB_OK + MB_ICONINFORMATION);
     }

   StatusBar1->SimpleText = "";
   Excel.OleProcedure("Quit");

}
//---------------------------------------------------------------------------
void __fastcall TMain::N2Click(TObject *Sender)
{
  Close();
}
//---------------------------------------------------------------------------

//Загрузка данных по новым договорам по ММК
void __fastcall TMain::NewDogovClick(TObject *Sender)
{
  AnsiString Sql, Sql1, inn, nn, data_po, data_po1;
  int i=1, rec=0;
  im_fl=1;

  /*rec - количество вставленных в таблицу записей*/


  if (Application->MessageBox(("Вы действительно хотите загрузить данные \n по новым договорам за " + Mes[DM->mm-1] + " " + DM->yyyy + " года?").c_str(),
                               "Загрузка данных по новым договорам",
                               MB_YESNO + MB_ICONINFORMATION) == IDNO)
    {
      Abort();
    }

  // Проверка правильности ИНН и наличие двойных записей в картотеке
  ProverkaInfoExcel();

  StatusBar1->SimpleText = "";


  try
    {
      Sheet.OleProcedure("Activate");

      Main->Cursor = crHourGlass;
      StatusBar1->SimplePanel = true;    // 2 панели на StatusBar1
      StatusBar1->SimpleText=" Идет загрузка данных...";

      ProgressBar->Visible = true;
      ProgressBar->Position = 0;
      ProgressBar->Max = Row;

      for ( i ; i<Row+1; i++)
        {
          nn = Excel.OlePropertyGet("Cells",i,1);
          inn = Excel.OlePropertyGet("Cells",i,5);

          ProgressBar->Position++;

          // Выбор строк необходимых для загрузки из Excel
          if (nn.IsEmpty() || !Proverka(nn) || inn.IsEmpty())  continue;

             //Проверка на наличие уже существующих записей в таблице VU_859_N
            Sql1 = "select * from VU_859_N where trim(inn)=trim("+ QuotedStr(Excel.OlePropertyGet("Cells",i,5)) +") \
                                           and trim(n_dogovora) = trim("+ QuotedStr(Excel.OlePropertyGet("Cells",i,9))+")" ;

            try
              {
                DM->qObnovlenie->Close();
                DM->qObnovlenie->SQL->Clear();
                DM->qObnovlenie->SQL->Add(Sql1);
                DM->qObnovlenie->Open();
              }
            catch(...)
              {
                Application->MessageBox("Ошибка получения данных из таблицы по страхованию 859 в/у","Ошибка",MB_OK+ MB_ICONERROR);
                Abort();
              }

            if (DM->qObnovlenie->RecordCount>0)
              {
                 if (Application->MessageBox(("Запись: цех = "+ DM->qObnovlenie->FieldByName("zex")->AsString +
                                               ", таб.№ = "+ DM->qObnovlenie->FieldByName("tn")->AsString +
                                               ", ИНН = "+ DM->qObnovlenie->FieldByName("inn")->AsString +
                                               " и № договора = "+DM->qObnovlenie->FieldByName("n_dogovora")->AsString +
                                              " уже существует. Записать ее еще раз?").c_str(),"Предупреждение",
                                              MB_YESNO + MB_ICONINFORMATION) ==ID_NO)
                    {
                       continue;
                    }
              }

            //Добавление цех+тн из sap_osn_sved
            Sql1="select zex, tn_sap, numident from sap_osn_sved where trim(numident)=trim("+ QuotedStr(Excel.OlePropertyGet("Cells",i,5)) +" )   \
                  union all                                                                                            \
                  select zex, tn_sap, numident from sap_sved_uvol                                                           \
                  where substr(to_char(dat_job,'dd.mm.yyyy'),4,7)='"+(DM->mm<10 ? "0"+IntToStr(DM->mm) : IntToStr(DM->mm))+"."+DM->yyyy+"'  \
                  and trim(numident)=trim("+QuotedStr(Excel.OlePropertyGet("Cells",i,5))+")";

           //  decode(translate('   123455','-0123456789 ','-'),null, '=p','=)')

            try
              {
                DM->qObnovlenie->Close();
                DM->qObnovlenie->SQL->Clear();
                DM->qObnovlenie->SQL->Add(Sql1);
                DM->qObnovlenie->Open();
              }
            catch(...)
              {
                Application->MessageBox("Ошибка получения данных из из картотеки по работникам (SAP_OSN_SVED, SAP_SVED_UVOL)","Ошибка",MB_OK+ MB_ICONERROR);
                Abort();
              }

            //Проверка на конечную дату
            data_po = Excel.OlePropertyGet("Cells",i,7);
            data_po1 = Excel.OlePropertyGet("Cells",i,7);

            if ((data_po.SubString(1,2)=="31" && data_po.SubString(4,2)=="04")||
                (data_po.SubString(1,2)=="31" && data_po.SubString(4,2)=="06")||
                (data_po.SubString(1,2)=="31" && data_po.SubString(4,2)=="09")||
                (data_po.SubString(1,2)=="31" && data_po.SubString(4,2)=="11"))
              {
                data_po = "30"+ data_po1.SubString(3,255);
              }

            //Запись данных в таблицу VU_859_N
            Sql = "insert into vu_859_N (zex, tn, fio, n_dogovora, kod_dogovora, data_s, data_po, sum, inn, priznak) \
                   values("+ QuotedStr(DM->qObnovlenie->FieldByName("zex")->AsString)+", \
                          "+ SetNull(DM->qObnovlenie->FieldByName("tn_sap")->AsString)+", \
                          initcap("+ QuotedStr(Excel.OlePropertyGet("Cells",i,2))+"||' '||"+QuotedStr(Excel.OlePropertyGet("Cells",i,3))+"||' '||"+QuotedStr(Excel.OlePropertyGet("Cells",i,4))+"), \
                          trim("+ QuotedStr(Excel.OlePropertyGet("Cells",i,9))+"), \
                             0, \
                          "+ QuotedStr(Excel.OlePropertyGet("Cells",i,6))+", \
                          "+ QuotedStr(data_po)+", \
                          "+ QuotedStr(Excel.OlePropertyGet("Cells",i,8))+", \
                          trim("+ QuotedStr(Excel.OlePropertyGet("Cells",i,5))+"),\
                             0 ) ";
            try
              {
                DM->qZagruzka->Close();
                DM->qZagruzka->SQL->Clear();
                DM->qZagruzka->SQL->Add(Sql);
                DM->qZagruzka->ExecSQL();
                rec++;
              }
            catch(...)
              {
                Application->MessageBox("Ошибка вставки данных в таблицу по страхованию 859 в/у","Ошибка",MB_OK+ MB_ICONERROR);
                Application->MessageBox("Данные не были загружены. Повторите загрузку","Ошибка",MB_OK+ MB_ICONERROR);
                StatusBar1->SimpleText = "";

                Excel.OleProcedure("Quit");
                Abort();
             }
        }


      Application->MessageBox(("Загрузка данных выполнена успешно =) \n Добавлено " + IntToStr(rec) + " записей").c_str(),
                               "Загрузка новых договоров по ММК",MB_OK+ MB_ICONINFORMATION);
      InsertLog("Выполнена загрузка данных по новым договорам по ММК. Загружено "+IntToStr(rec)+" записей");

      Excel.OleProcedure("Quit");
      Excel = Unassigned;

      StatusBar1->SimplePanel = false;
      ProgressBar->Visible = false;
      StatusBar1->SimpleText = "Обновление выполнено.";
      Main->Cursor = crDefault;
      StatusBar1->SimpleText = "";
    }
  catch(...)
    {
      Application->MessageBox("Ошибка загрузки данных по новым договорам по ММК","Ошибка",MB_OK+ MB_ICONERROR);
      Excel.OleProcedure("Quit");

      StatusBar1->SimplePanel = false;
      ProgressBar->Visible = false;
      StatusBar1->SimpleText = "";
      Main->Cursor = crDefault;
    }
}
//---------------------------------------------------------------------------

//---------------------------------------------------------------------------

// Возвращает путь на папку Мои документы"
bool __fastcall TMain::GetMyDocumentsDir(AnsiString &FolderPath)
{
  char f[MAX_PATH];

  if (SUCCEEDED(SHGetFolderPath(NULL, CSIDL_PERSONAL|CSIDL_FLAG_CREATE, NULL, SHGFP_TYPE_CURRENT, f))) {
    FolderPath = AnsiString(f);
    return(true);
  }

  return(false);
}
//---------------------------------------------------------------------------

// Возвращает путь на папку Temp
bool __fastcall TMain::GetTempDir(AnsiString &FolderPath)
{
  char f[MAX_PATH];

  if (GetTempPath(MAX_PATH, f)) {
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
AnsiString  __fastcall TMain::SetNull (AnsiString str, AnsiString r)
{
  if (str.Length()) return str;
  else return r;
}
//---------------------------------------------------------------------------

//Загрузка изменений по валютным договорам
void __fastcall TMain::izm_valClick(TObject *Sender)
{
  im_fl=3;
  if (Application->MessageBox(("Вы действительно хотите загрузить изменения \n по валютным договорам за " + Mes[DM->mm-1] + " " + DM->yyyy + " года?").c_str(),
                               "Загрузка изменений по валютным договорам",
                               MB_YESNO + MB_ICONINFORMATION) == IDNO)
    {
      Abort();
    }

  // Проверка на существование инн в таблице Avans
  ProverkaInfoExcelIzmeneniya();

  StatusBar1->SimpleText = "";

  //Обновление изменений по внешней страховке
  UpdateValuta_I_Grivna();

  InsertLog("Выполнена загрузка изменений по валютным договорам. Обновлено "+obnov_kol+" из "+ob_kol+" записей");

  StatusBar1->SimpleText = "";

}
//---------------------------------------------------------------------------

//Загрузка изменений по гривневым договорам
void __fastcall TMain::izm_grnClick(TObject *Sender)
{
  im_fl=2;
  
  if (Application->MessageBox(("Вы действительно хотите загрузить изменения \n по гривневым договорам за " + Mes[DM->mm-1] + " " + DM->yyyy + " года?").c_str(),
                               "Загрузка изменений по гривневым договорам",
                               MB_YESNO + MB_ICONINFORMATION) == IDNO)
    {
      Abort();
    }

  // Проверка на существование инн в таблице
  ProverkaInfoExcelIzmeneniya();

  StatusBar1->SimpleText = "";

  //Обновление изменений по гривневой страховке
  UpdateValuta_I_Grivna();

  InsertLog("Выполнена загрузка изменений по гривневым договорам. Обновлено "+obnov_kol+" из "+ob_kol+" записей");

  StatusBar1->SimpleText = "";
}
//---------------------------------------------------------------------------

//Загрузка изменений пересчета по курсу
void __fastcall TMain::kurs_pereschetClick(TObject *Sender)
{ /*int i=0, rec=0;
  AnsiString tn, fam, data_s,data_po, n_dogovora, Sql;
  bool fl=0;   */

  im_fl=4;

  if (Application->MessageBox(("Вы действительно хотите загрузить корректировки платежей \n в пересчете по курсу НБУ за " + Mes[DM->mm-1] + " " + DM->yyyy + " года?").c_str(),
                               "Загрузка изменений пересчета по курсу",
                               MB_YESNO + MB_ICONINFORMATION) == IDNO)
    {
      Abort();
    }

  // Проверка на существование инн в таблице Avans
  ProverkaInfoExcelIzmeneniya();

  UpdateValuta_I_Grivna();

  InsertLog("Выполнена загрузка изменений пересчета по курсу. Обновлено "+obnov_kol+" из "+ob_kol+" записей");

  StatusBar1->SimpleText = "";

}
//---------------------------------------------------------------------------




//---------------------------------------------------------------------------

void __fastcall TMain::InsertLog(AnsiString Msg)
{
  AnsiString Sql;
  AnsiString Data;
  DateTimeToString(Data, "dd.mm.yyyy hh:nn:ss", Now());
  
  Sql= "insert into logs_strax (DT, DOMAIN, USEROK, PROG, TEXT, USEROK_FIO) values \
                     (to_date(" + QuotedStr(Data) + ", 'DD.MM.YYYY HH24:MI:SS'),\
                      "+ QuotedStr(DomainName) +", " + QuotedStr(UserName) + ", \
                      'Strahovka', replace(" + QuotedStr(Msg) + ",',','.')," + QuotedStr(UserFullName)+")";

  DM->qObnovlenie->SQL->Clear();
  DM->qObnovlenie->SQL->Add(Sql);

  DM->qObnovlenie->ExecSQL();
  DM->qObnovlenie->Close();
}
//---------------------------------------------------------------------------


void __fastcall TMain::FormShow(TObject *Sender)
{
 Panel1->Visible = false;        
}
//---------------------------------------------------------------------------


void __fastcall TMain::BitBtn1Click(TObject *Sender)
{
  AnsiString Sql, fio, inn, n_dog;
   
   if (EditZEX2->Text.IsEmpty() ||
      EditTN2->Text.IsEmpty() ||
      EditSum->Text.IsEmpty()||
      EditVal->Text.IsEmpty()||
      EditData_s->Text.IsEmpty()||
      EditData_po->Text.IsEmpty())
    {
      Application->MessageBox("Введены не все данные для изменения","Обновление записи",
                              MB_OK + MB_ICONINFORMATION);
      EditZEX2->SetFocus();
      Abort();
    }

  if (fl_r==0)
    {
      // Добавление записи

      //Проверка на ввод номера договора
      if (EditNDOG->Text.IsEmpty() || EditNDOG->Text.Length()<12)
        {
          Application->MessageBox("Введите № Договора","Добавление записи",
                                   MB_OK + MB_ICONINFORMATION);
          EditNDOG->SetFocus();
          Abort();
        }

        //Проверка на существование работника в таблице
       Sql ="select distinct numident, fam||' '||im||' '||ot as fio from sap_osn_sved \
             where zex="+EditZEX2->Text+" and tn_sap="+EditTN2->Text;

       DM->qObnovlenie->Close();
       DM->qObnovlenie->SQL->Clear();
       DM->qObnovlenie->SQL->Add(Sql);

       try
         {
           DM->qObnovlenie->Open();
         }
       catch(...)
         {
           Application->MessageBox("Ошибка доступа к таблице SAP_OSN_SVED",
                                   "Ошибка доступа",MB_OK + MB_ICONERROR);
           Abort();
         }

       if (DM->qObnovlenie->RecordCount==0)
         {
           Application->MessageBox("Нет данных по этому работнику.\nПроверьте правильность ввода цеха и табельного номера",
                                   "Предупреждение",MB_OK+MB_ICONINFORMATION);
           Abort();
         }
       else if (DM->qObnovlenie->RecordCount>1 || DM->qObnovlenie->FieldByName("vnvi")->AsString.IsEmpty())
         {
           Application->MessageBox("Более двух записей или нет инд.№ в таблице SAP_OSN_SVED по данному работнику.\nНевозможно сохранить информацию.",
                                   "Ошибка",MB_OK+MB_ICONINFORMATION);
         }
       else
         {
           fio = DM->qObnovlenie->FieldByName("fio")->AsString;
           inn = DM->qObnovlenie->FieldByName("numident")->AsString;

           //Проверка на существование записи в таблице vu_859_n
           Sql = "select * from vu_859_n where trim(n_dogovora)=trim("+ QuotedStr(EditNDOG->Text)+") and \
                                               trim(inn) = trim(" + QuotedStr(inn)+")";

           DM->qObnovlenie->Close();
           DM->qObnovlenie->SQL->Clear();
           DM->qObnovlenie->SQL->Add(Sql);

           try
             {
               DM->qObnovlenie->Open();
             }
           catch(...)
             {
               Application->MessageBox("Возникла ошибка при проверке данных в таблице VU_859_N",
                                       "Ошибка",MB_OK + MB_ICONERROR);
               Abort();
            }

          if (DM->qObnovlenie->RecordCount>0)
            {
              Application->MessageBox("Работник с таким номером договора уже существует",
                                       "Ошибка",MB_OK + MB_ICONERROR);
              EditNDOG->SetFocus();
              Abort();
            }

          Sql ="insert into vu_859_n (zex, tn, fio, n_dogovora, kod_dogovora, data_s, data_po, sum, priznak, inn) \
                values ("+ EditZEX2->Text +",\
                        "+ EditTN2->Text +",\
                        "+ QuotedStr(fio) +",\
                        "+ QuotedStr(EditNDOG->Text) +",\
                        "+ EditVal->Text +",\
                        "+ QuotedStr(EditData_s->Text) +",\
                        "+ QuotedStr(EditData_po->Text) +",\
                        "+ EditSum->Text +",\
                        0, \
                        "+ QuotedStr(inn) +")";
          DM->qObnovlenie->Close();
          DM->qObnovlenie->SQL->Clear();
          DM->qObnovlenie->SQL->Add(Sql);

          n_dog = EditNDOG->Text;
          try
            {
              DM->qObnovlenie->ExecSQL();
            }
          catch(...)
            {
              Application->MessageBox("Возникла ошибка при добавлении данных",
                                      "Ошибка добавления новой записи",MB_OK + MB_ICONERROR);
              Abort();
            }

           DM->qKorrektirovka->Close();
           DM->qKorrektirovka->Parameters->ParamByName("pzex")->Value = EditZEX2->Text;
           DM->qKorrektirovka->Parameters->ParamByName("ptn")->Value = EditTN2->Text;

           try
             {
               DM->qKorrektirovka->Open();
             }
           catch(...)
             {
               Application->MessageBox("Ошибка получения данных из таблицы","Ошибка",MB_OK + MB_ICONERROR);
               Abort();
             }

           InsertLog("Выполнено добавление записи: цех ="+ EditZEX2->Text +", таб.№ ="+ EditTN2->Text +", № договора = "+EditNDOG->Text+", сумма = "+EditSum->Text);

           TLocateOptions SearchOptions;
           DM->qKorrektirovka->Locate("n_dogovora",n_dog,SearchOptions<<loPartialKey<<loCaseInsensitive);

        }
    }
  else
    {
      // Редактирование записи

      //Проверка на ввод признака платежа
      if (EditPRIZNAK->Text.IsEmpty())
        {
          Application->MessageBox("Введите признак платежа\n   0 - платит\n   1 - уволен\n   2 - окончание срока договора\n   3 - декрет\n   4 - прекращение договора\n   6 - приостановление договора\n   7 - договор не вступил в силу",
                                  "Добавление записи",
                                  MB_OK + MB_ICONINFORMATION);
          EditPRIZNAK->SetFocus();
          Abort();
        }

      Sql= "update vu_859_n set \
                            zex="+EditZEX2->Text+", \
                            tn="+EditTN2->Text+",\
                            priznak="+EditPRIZNAK->Text+",\
                            sum= "+EditSum->Text+",\
                            kod_dogovora= "+EditVal->Text+",\
                            data_s="+QuotedStr(EditData_s->Text)+",\
                            data_po="+QuotedStr(EditData_po->Text)+  " \
            where zex="+DM->qKorrektirovka->FieldByName("zex")->AsString+" and \
                  tn= "+DM->qKorrektirovka->FieldByName("tn")->AsString+" and \
                  rowid=chartorowid("+QuotedStr(DM->qKorrektirovka->FieldByName("rw")->AsString)+")";
        //     n_dogovora="+SetNull(DM->qKorrektirovka->FieldByName("n_dogovora")->AsString);

      DM->qObnovlenie->Close();
      DM->qObnovlenie->SQL->Clear();
      DM->qObnovlenie->SQL->Add(Sql);

      try
        {
          DM->qObnovlenie->ExecSQL();
        }
      catch(...)
        {
          Application->MessageBox("Возникла ошибка при обновлении данных",
                                  "Ошибка обновления данных",MB_OK + MB_ICONERROR);
          Abort();
        }

      if (zzex != EditZEX2->Text ||
          ztn != EditTN2->Text ||
          zsum != EditSum->Text ||
          zdata_s != EditData_s->Text ||
          zdata_po != EditData_po->Text ||
          zval != EditVal->Text ||
          zpriznak != EditPRIZNAK->Text)
        {
          InsertLog("Редактирование записи по договору №: "+DM->qKorrektirovka->FieldByName("n_dogovora")->AsString+" c цех = "+ zzex +" на "+EditZEX2->Text+" c таб.№ ="+ztn+" на "+EditTN2->Text+" c суммы = "+zsum+" на "+EditSum->Text+" с валюты "+zval+" на "+EditVal->Text+" с признака "+zpriznak+" на "+EditPRIZNAK->Text+" с даты "+zdata_s+"-"+zdata_po+" на "+EditData_s->Text+"-"+EditData_po->Text);
        }
      rec = DM->qKorrektirovka->RecNo;

      DM->qKorrektirovka->Requery();

      //Возврат на обновляемую запись
      if (!(EditTN2->Text.IsEmpty() && EditZEX2->Text.IsEmpty()))
        {
          DM->qKorrektirovka->RecNo = rec;
        }
  }

 /*
  // select v.*, rowidtochar(rowid) rw from vu_859_n v where zex=:pzex and tn=:ptn


//  TLocateOptions SearchOptions;
  // Variant locvalues[] = {DM->qKorrektirovka->FieldByName("rw")->AsString, TABPEdit->Text};

 // DM->qKorrektirovka->Locate("rw",QuotedStr(rw),SearchOptions<<loPartialKey<<loCaseInsensitive);


        */
}
//---------------------------------------------------------------------------

void __fastcall TMain::BitBtn2Click(TObject *Sender)
{
  Panel1->Visible = false;
}
//---------------------------------------------------------------------------

//Поиск записи для редактирования
void __fastcall TMain::BitBtn3Click(TObject *Sender)
{
  if (EditZEX->Text.IsEmpty() || EditTN->Text.IsEmpty())
    {
      Application->MessageBox("Не введен цех или табельный номер работника","Введите необходимую информацию", MB_OK + MB_ICONINFORMATION);
      EditZEX->SetFocus();
      EditZEX->SelectAll();
      Abort();
    }

  DM->qKorrektirovka->Close();
  DM->qKorrektirovka->Parameters->ParamByName("pzex")->Value = EditZEX->Text;
  DM->qKorrektirovka->Parameters->ParamByName("ptn")->Value = EditTN->Text;

  try
    {
      DM->qKorrektirovka->Open();
    }
  catch(...)
    {
      Application->MessageBox("Ошибка получения данных из таблицы","Ошибка",MB_OK + MB_ICONERROR);
      Abort();
    }

  if (DM->qKorrektirovka->RecordCount==0)
    {
       Application->MessageBox("Работник с таким цехом и табельным номером не найден","Поиск записи",MB_OK + MB_ICONINFORMATION);
       EditZEX->SetFocus();
       EditZEX->SelectAll();
       Label11->Visible=false;
       Abort();

    }

  //вывод записей в DBGrid
  Label11->Visible = true;
  Label12->Visible = true;


  Label9->Caption="Редактирование данных:";

  DBGridEh1->Visible=true;
  EditZEX2->Visible=true;
  EditTN2->Visible=true;
  EditSum->Visible=true;
  EditData_s->Visible=true;
  EditData_po->Visible=true;
  EditVal->Visible=true;
  BitBtn1->Visible=true;
  BitBtn2->Visible=true;
  Label2->Visible=true;
  Label3->Visible=true;
  Label4->Visible=true;
  Label5->Visible=true;
  Label6->Visible=true;
  Label7->Visible=true;
  Bevel1->Visible=true;
  Bevel3->Visible=true;
  Label8->Visible=true;
  Label9->Visible=true;
  Label10->Visible=true;
  EditPRIZNAK->Visible = true;
  DBGridEh1->SetFocus();
  SetEditData();


}
//---------------------------------------------------------------------------
void __fastcall TMain::SetEditData()
{
  EditData_po->Font->Color = clBlack;
  EditData_s->Font->Color = clBlack;
  EditZEX2->Text = zzex = DM->qKorrektirovka->FieldByName("ZEX")->AsString;
  EditTN2->Text = ztn = DM->qKorrektirovka->FieldByName("TN")->AsString;
  EditSum->Text = zsum = DM->qKorrektirovka->FieldByName("sum")->AsString;
  EditData_s->Text = zdata_s = DM->qKorrektirovka->FieldByName("data_s")->AsString;
  EditData_po->Text = zdata_po = DM->qKorrektirovka->FieldByName("data_po")->AsString;
  EditVal->Text = zval = DM->qKorrektirovka->FieldByName("kod_dogovora")->AsString;
  EditPRIZNAK->Text = zpriznak = DM->qKorrektirovka->FieldByName("priznak")->AsString;
  
    switch (DM->qKorrektirovka->FieldByName("priznak")->AsInteger)
    {
      case 0:  Label11->Caption="платит";
      break;
      case 1:  Label11->Caption="уволен";
      break;
      case 2:  Label11->Caption="закончился срок договора";
      break;
      case 3:  Label11->Caption="декрет";
      break;
      case 4:  Label11->Caption="расторжение";
      break;
      case 5:  Label11->Caption="приостановлен";
      break;
      case 6:  Label11->Caption="приостановлен";
      break;
      case 7:  Label11->Caption=" не вступил в силу";
      break;
      default:
          Label11->Caption=" ";
    }

  EditNDOG->Text = DM->qKorrektirovka->FieldByName("n_dogovora")->AsString;
}
//---------------------------------------------------------------------------


void __fastcall TMain::FormKeyDown(TObject *Sender, WORD &Key,
      TShiftState Shift)
{
  if (Key == VK_RETURN)
  FindNextControl((TWinControl *)Sender, true, true,
                   false)->SetFocus();

}
//---------------------------------------------------------------------------

void __fastcall TMain::EditZEXKeyPress(TObject *Sender, char &Key)
{
  if (Key==','||Key=='/') Key='.';
  if (!(IsNumeric(Key)||Key=='\b'||Key==','||Key=='.'||Key=='/')) Key=0;
}
//---------------------------------------------------------------------------

void __fastcall TMain::EditSumKeyPress(TObject *Sender, char &Key)
{
  if (Key==','||Key=='/') Key='.';
  if (!(IsNumeric(Key)||Key=='\b'||Key==','||Key=='.'||Key=='/')) Key=0;

}
//---------------------------------------------------------------------------

void __fastcall TMain::EditData_sExit(TObject *Sender)
{
    // проверка даты

  TDateTime d;

  if (ActiveControl == BitBtn2)
    {
      Panel1->Visible = false;
    }
  else
    {
      if (!EditData_s->Text.IsEmpty())
        {
          if(!TryStrToDate(EditData_s->Text,d))
            {
              Application->MessageBox("Неверный формат даты","Ошибка", MB_OK);
              EditData_s->Font->Color = clRed;
              EditData_s->SetFocus();
            }
          else
            {
              EditData_s->Text=FormatDateTime("dd.mm.yyyy",d);
              EditData_s->Font->Color = clBlack;
            }

        }
    }
}
//---------------------------------------------------------------------------

void __fastcall TMain::EditData_poExit(TObject *Sender)
{
  TDateTime d;

  if (ActiveControl == BitBtn2)
    {
      Panel1->Visible = false;
    }
  else
    {
      if (!EditData_po->Text.IsEmpty())
        {
          if(!TryStrToDate(EditData_po->Text,d))
            {
              Application->MessageBox("Неверный формат даты","Ошибка", MB_OK);
              EditData_po->Font->Color = clRed;
              EditData_po->SetFocus();
            }
          else
            {
              EditData_po->Text=FormatDateTime("dd.mm.yyyy",d);
              EditData_po->Font->Color = clBlack;
            }

        }
    }
}
//---------------------------------------------------------------------------

void __fastcall TMain::SpeedButton1Click(TObject *Sender)
{
   Panel1->Visible = false;

}
//---------------------------------------------------------------------------

// Сверка данных
void __fastcall TMain::N7Click(TObject *Sender)
{
  AnsiString Sql;
//  Word year,month,day;

 // DecodeDate(Date(), year, month,day);


//******************************************************************************
   //Обновление цех+тн по переводам

   StatusBar1->SimpleText = "Обновление данных по переведенным работникам...";

   Sql = "select v.zex as s_zex, v.tn as s_tn, v.fio, v.inn as inn, s.zex as zex, s.tn_sap as tn                  \
          from (                                                                                                  \
                (select * from vu_859_n) v                                                                        \
                 left join                                                                                        \
                (select zex, tn_sap, numident, initcap(fam||' '||im||' '||ot) as fio  from sap_osn_sved           \
                 union all                                                                                        \
                 select zex, tn_sap, numident, initcap(fam||' '||im||' '||ot) as fio  from sap_sved_uvol) s       \
                 on v.inn=s.numident                                                                              \
                )                                                                                                 \
          where nvl(priznak,0)!=1                                                                                 \
          and (inn in (select numident from sap_osn_sved) or inn in (select numident from sap_sved_uvol where substr(to_char(dat_job,'dd.mm.yyyy'),4,7)='"+(DM->mm<10 ? "0"+IntToStr(DM->mm) : IntToStr(DM->mm))+"."+DM->yyyy+"')) \
          and (to_char(s.zex)!=to_char(v.zex) or s.tn_sap!=v.tn or (to_char(s.zex)!=to_char(v.zex) and s.tn_sap!=v.tn))";


   DM->qObnovlenie->Close();
   DM->qObnovlenie->SQL->Clear();
   DM->qObnovlenie->SQL->Add(Sql);
   try
     {
       DM->qObnovlenie->Open();
     }
   catch (...)
     {
       Application->MessageBox("Возникла ошибка при выборке данных по переводам",
                               "Обновление данных",MB_OK + MB_ICONERROR);
       StatusBar1->SimpleText = "";
       Abort();
     }

   while (!DM->qObnovlenie->Eof)
     {
       Sql = " update vu_859_n set zex = "+QuotedStr(DM->qObnovlenie->FieldByName("zex")->AsString)+", \
                                   tn = "+DM->qObnovlenie->FieldByName("tn")->AsString+", \
                                   mes = "+IntToStr(DM->mm)+",  \
                                   god = "+IntToStr(DM->yyyy)+"   \
               where inn="+QuotedStr(DM->qObnovlenie->FieldByName("inn")->AsString)+" and priznak!=1";

       DM->qZagruzka->Close();
       DM->qZagruzka->SQL->Clear();
       DM->qZagruzka->SQL->Add(Sql);
       try
         {
           DM->qZagruzka->ExecSQL();
         }
       catch (...)
         {
           Application->MessageBox("Возникла ошибка при обновлении данных по переводам",
                                   "Обновление данных",MB_OK + MB_ICONERROR);

           StatusBar1->SimpleText = "";
           InsertLog("Сверка данных не выполнена: ошибка обновления цех+таб по переводам");
           Abort();
         }

       DM->qObnovlenie->Next();

     }
//********************************************************************************
  // Обновление поля priznak: уволенные = 1
  StatusBar1->SimpleText = "Обновление данных по уволеным";

  Sql = "update vu_859_n set priznak = 1, mes="+IntToStr(DM->mm)+", \
                             god="+IntToStr(DM->yyyy)+" \
         where tn in (select tn_sap from sap_sved_uvol where substr(to_char(dat_job,'dd.mm.yyyy'),4,7)<'"+(DM->mm<10 ? "0"+IntToStr(DM->mm) : IntToStr(DM->mm))+"."+DM->yyyy+"')     \
         and priznak in (0,3,7,5)";

  DM->qObnovlenie->Close();
  DM->qObnovlenie->SQL->Clear();
  DM->qObnovlenie->SQL->Add(Sql);
  try
    {
      DM->qObnovlenie->ExecSQL();
    }
  catch (...)
    {
      Application->MessageBox("Возникла ошибка при обновлении данных по уволеным",
                              "Обновление данных",MB_OK + MB_ICONERROR);
      StatusBar1->SimpleText = "";
      InsertLog("Сверка данных не выполнена: ошибка обновления данных по уволеным");
      Abort();
    }

  // Обновление поля priznak: окончен срок выплат = 2, кроме пенсионного страхования
  StatusBar1->SimpleText = "Обновление данных по окончанию срока договора";

  Sql = " update vu_859_n set priznak = 2, mes="+IntToStr(DM->mm)+", \
                              god="+IntToStr(DM->yyyy)+" \
          where to_char(data_po,'yyyymm')< " \
                + IntToStr(DM->yyyy) + "||lpad("+IntToStr(DM->mm)+",2,'0') and priznak in (0,3,7,5) \
          and  (tn in (select tn_sap from sap_osn_sved) or tn in (select tn_sap from sap_sved_uvol)) \
          and kod_dogovora!=4";

  DM->qObnovlenie->Close();
  DM->qObnovlenie->SQL->Clear();
  DM->qObnovlenie->SQL->Add(Sql);
  try
    {
      DM->qObnovlenie->ExecSQL();
    }
  catch (...)
    {
      Application->MessageBox("Возникла ошибка при обновлении данных по договорам, \n у которых окончился срок",
                              "Обновление данных",MB_OK + MB_ICONERROR);

      StatusBar1->SimpleText = "";
      InsertLog("Сверка данных не выполнена: ошибка обновления данных по окончанию срока договора");
      Abort();
    }


 //*****************************************************************************


  // Формирование отчета содержащего уволенных, окончанием выплат и измененным цехом и таб.№
  StatusBar1->SimpleText = "Идет формирование отчета...";

  //Открытие файла данных содержащего уволенных, окончанием выплат и измененным цехом и таб.№
  if (!rtf_Open((TempPath + "\\sverka.txt").c_str()))
    {
      MessageBox(Handle,"Ошибка открытия файла данных","Ошибка",8192);
    }
  else
    {

// Выборка данных по уволенным
      Sql = "select distinct zex, tn, fio, (select dat_job from sap_sved_uvol s where s.tn_sap=v.tn) as dtuvol \
             from vu_859_n v                                                                                   \
             where priznak=1                                                                                   \
             and mes="+IntToStr(DM->mm)+" and god="+IntToStr(DM->yyyy)+"                                       \
             and tn in (select tn_sap from sap_sved_uvol)                                                   \
             order by zex,tn";

      DM->qObnovlenie->Close();
      DM->qObnovlenie->SQL->Clear();
      DM->qObnovlenie->SQL->Add(Sql);
      try
        {
          DM->qObnovlenie->Open();
        }
      catch (...)
        {
          Application->MessageBox("Возникла ошибка при выборке данных из таблицы по уволеным",
                                  "Формирование отчета",MB_OK + MB_ICONERROR);
          StatusBar1->SimpleText = "";
          Abort();
        }

// Вывод в отчет уволенных
      if (DM->qObnovlenie->RecordCount>0)
        {
          // Вывод заголовка и шапки таблицы
          rtf_Out("z", " ", 1);
          if(!rtf_LineFeed())
            {
              MessageBox(Handle,"Ошибка записи в файл данных","Ошибка",8192);
              if (!rtf_Close()) MessageBox(Handle,"Ошибка закрытия файла данных","Ошибка",8192);
              return;
            }
        }
      while (!DM->qObnovlenie->Eof)
        {
          rtf_Out("zex", DM->qObnovlenie->FieldByName("zex")->AsString,2);
          rtf_Out("tn", DM->qObnovlenie->FieldByName("tn")->AsString,2);
          rtf_Out("fio",DM->qObnovlenie->FieldByName("fio")->AsString,2);
          rtf_Out("dtuvol",DM->qObnovlenie->FieldByName("dtuvol")->AsString,2);

          if(!rtf_LineFeed())
            {
              MessageBox(Handle,"Ошибка записи в файл данных","Ошибка",8192);
              if (!rtf_Close()) MessageBox(Handle,"Ошибка закрытия файла данных","Ошибка",8192);
              return;
            }

          DM->qObnovlenie->Next();
        }


// Выборка данных по договорам, у которых окончился срок выплат
      Sql = " select  zex, tn, fio, n_dogovora, \
                      data_po from vu_859_n where priznak=2 \
                      and mes="+IntToStr(DM->mm)+" and god="+IntToStr(DM->yyyy)+"  \
                      and  (tn in (select tn_sap from sap_osn_sved) or tn in (select tn_sap from sap_sved_uvol)) \
                      and kod_dogovora!=4 \
                      order by zex,tn";

      DM->qObnovlenie->Close();
      DM->qObnovlenie->SQL->Clear();
      DM->qObnovlenie->SQL->Add(Sql);
      try
        {
          DM->qObnovlenie->Open();
        }
      catch (...)
        {
          Application->MessageBox("Возникла ошибка при выборке данных по договорам, \n у которых окончился срок",
                                  "Формирование отчета",MB_OK + MB_ICONERROR);

          StatusBar1->SimpleText = "";
          Abort();
        }

      //Вывод в отчет договора с окончанием срока
      if (DM->qObnovlenie->RecordCount>0)
        {
          // Вывод заголовка и шапки таблицы
          rtf_Out("zz", " ", 3);
          if(!rtf_LineFeed())
            {
              MessageBox(Handle,"Ошибка записи в файл данных","Ошибка",8192);
              if (!rtf_Close()) MessageBox(Handle,"Ошибка закрытия файла данных","Ошибка",8192);
              return;
            }
        }
      while (!DM->qObnovlenie->Eof)
        {
          rtf_Out("zex", DM->qObnovlenie->FieldByName("zex")->AsString,4);
          rtf_Out("tn", DM->qObnovlenie->FieldByName("tn")->AsString,4);
          rtf_Out("fio",DM->qObnovlenie->FieldByName("fio")->AsString,4);
          rtf_Out("n_dogovora",DM->qObnovlenie->FieldByName("n_dogovora")->AsString,4);
          rtf_Out("data_po",DM->qObnovlenie->FieldByName("data_po")->AsString,4);

          if(!rtf_LineFeed())
            {
              MessageBox(Handle,"Ошибка записи в файл данных","Ошибка",8192);
              if (!rtf_Close()) MessageBox(Handle,"Ошибка закрытия файла данных","Ошибка",8192);
              return;
            }

          DM->qObnovlenie->Next();
        }
  //  }

//Выборка данных по переводам по цех и тн
       Sql="select v.zex as zex, tn, initcap(fio) as fio, s.zex as zexp, dat_job                                          \
            from vu_859_n v, sap_perevod s                                                                 \
            where priznak not in (1,2)                                                                     \
            and tn=tn_sap                                                                                  \
            and mes=10 and god=2015                                                                        \
            and (tn in (select tn_sap from sap_osn_sved) or tn in (select tn_sap from sap_sved_uvol))      \
            and dat_job in ((select max(dat_job)                                                           \
                             from sap_perevod s2                                                           \
                             where dat_job not in  (select max(dat_job) from sap_perevod s1 where s1.tn_sap=v.tn) and s2.tn_sap=v.tn))";


      DM->qObnovlenie->Close();
      DM->qObnovlenie->SQL->Clear();
      DM->qObnovlenie->SQL->Add(Sql);
      try
        {
          DM->qObnovlenie->Open();
        }
      catch (...)
        {
          Application->MessageBox("Возникла ошибка при выборке данных по переводам",
                                  "Обновление данных",MB_OK + MB_ICONERROR);

          StatusBar1->SimpleText = "";
          Abort();
        }

//Вывод в отчет переводов по цех и тн
      if (DM->qObnovlenie->RecordCount>0)
        {
          // Вывод заголовка и шапки таблицы
          rtf_Out("zzz", " ", 5);
          if(!rtf_LineFeed())
            {
              MessageBox(Handle,"Ошибка записи в файл данных","Ошибка",8192);
              if (!rtf_Close()) MessageBox(Handle,"Ошибка закрытия файла данных","Ошибка",8192);
              return;
            }
        }

      while (!DM->qObnovlenie->Eof)
        {
          rtf_Out("tn", DM->qObnovlenie->FieldByName("tn")->AsString,6);
          rtf_Out("zexp", DM->qObnovlenie->FieldByName("zexp")->AsString,6);
          rtf_Out("zex", DM->qObnovlenie->FieldByName("zex")->AsString,6);
          rtf_Out("dat_job", DM->qObnovlenie->FieldByName("dat_job")->AsString,6);
          rtf_Out("fio",DM->qObnovlenie->FieldByName("fio")->AsString,6);

          
        /*  AnsiString dtuvol= DM->qObnovlenie->FieldByName("dtuvol")->AsString;
          rtf_Out("dtuvol",(dtuvol.SubString(7,2)+"."+
                            dtuvol.SubString(5,2)+"."+
                            dtuvol.SubString(1,4)), 6);
                                                             */

         // rtf_Out("dtuvol",DM->qObnovlenie->FieldByName("av.dtuvol")->AsString,6);

          if(!rtf_LineFeed())
            {
              MessageBox(Handle,"Ошибка записи в файл данных","Ошибка",8192);
              if (!rtf_Close()) MessageBox(Handle,"Ошибка закрытия файла данных","Ошибка",8192);
              return;
            }

          DM->qObnovlenie->Next();
        }
      
      if(!rtf_Close())
        {
          MessageBox(Handle,"Ошибка закрытия файла данных", "Ошибка", 8192);
          return;
        }
      int istrd;
      try
        {
          rtf_CreateReport(TempPath +"\\sverka.txt", Path+"\\RTF\\sverka.rtf",
                           WorkPath+"\\Сверка данных.doc",NULL,&istrd);


              WinExec(("\""+ WordPath+"\"\""+WorkPath+"\\Сверка данных.doc\"").c_str(),SW_MAXIMIZE);

        }
      catch(RepoRTF_Error E)
        {
          MessageBox(Handle,("Ошибка формирования отчета:"+ AnsiString(E.Err)+
                             "\nСтрока файла данных:"+IntToStr(istrd)).c_str(),"Ошибка",8192);
        }

    }
    
    InsertLog("Сверка данных выполнена успешно");
    StatusBar1->SimpleText = "";
}
//---------------------------------------------------------------------------

void __fastcall TMain::DBGridEh1DrawColumnCell(TObject *Sender,
      const TRect &Rect, int DataCol, TColumnEh *Column,
      TGridDrawState State)
{

/*     if TDBGrideh(Sender).DataSource.DataSet.FieldByName('color').AsString =FormatDateTime('DDDD', Now) then
  begin
    TDBGrideh(Sender).Canvas.Font.Color:=clBlue;
    TDBGrideh(Sender).DefaultDrawColumnCell(Rect,DataCol,Column,State);
    TDBGrideh(Sender).Canvas.Font.Style:=[fsBold];
    TDBGrideh(Sender).Canvas.Brush.Color:=clYellow;
    TDBGrideh(Sender).DefaultDrawColumnCell(Rect,DataCol,Column,State);
  end;

               */
   /*

      TDBGridEh *  pObj = (TDBGridEh *)Sender;
   if( !State.Contains(gdFixed) && (pObj->rec== ) ) {
      TCanvas * pCanvas = pObj->Canvas;
      TColor OldColor = pCanvas->Brush->Color;
      pCanvas->Brush->Color = clGray;
      pCanvas->FillRect(Rect);
      pCanvas->TextRect(Rect,Rect.Left+2,Rect.Top+2,Column->Field->Text);
      pCanvas->Brush->Color = OldColor;
   }          */
    /*

     Выделение выбранной строки зеленым цветом
      if( State.Contains(gdSelected) ) {
      TColor oldColor = DBGridEh1->Canvas->Brush->Color;
      DBGridEh1->Canvas->Brush->Color = TColor(0x001FE0B0);
      DBGridEh1->Canvas->FillRect(Rect);
      DBGridEh1->Canvas->TextOut(Rect.Left+2,Rect.Top+2,Column->Field->Text);
      DBGridEh1->Canvas->Brush->Color = oldColor;
   }                                                    */

 /*  if(!State.Contains(gdFixed) && DM->qKorrektirovka->RecNo==rec ) {
      TColor oldColor = DBGridEh1->Canvas->Brush->Color;
      DBGridEh1->Canvas->Brush->Color = TColor(0x001FE0B0);
      DBGridEh1->Canvas->FillRect(Rect);
      DBGridEh1->Canvas->TextOut(Rect.Left+2,Rect.Top+2,Column->Field->Text);
      DBGridEh1->Canvas->Brush->Color = oldColor;
   }      */

      // выделение цветом всех записей,
      if(!State.Contains(gdFixed) && DM->qKorrektirovka->RecNo==rec)
        {
          ((TDBGridEh *) Sender)->Canvas->Brush->Color = TColor(0x001FE0B0);           //clGradientActiveCaption;
          ((TDBGridEh *) Sender)->DefaultDrawColumnCell(Rect, DataCol, Column, State);
        }

       // выделение цветом активной записи
      if (State.Contains(gdSelected) )
        {
          if( DM->qKorrektirovka->RecNo!=rec )
            {
              ((TDBGridEh *) Sender)->Canvas->Brush->Color = clInactiveCaption;
            }
          else
            {
              ((TDBGridEh *) Sender)->Canvas->Brush->Color = cl3DLight;
            }
          ((TDBGridEh *) Sender)->Canvas->Font->Color= clBlack;
        }
      ((TDBGridEh *) Sender)->DefaultDrawColumnCell(Rect, DataCol, Column, State);

    

}
//---------------------------------------------------------------------------
// Загрузка внешних договоров
void __fastcall TMain::N11Click(TObject *Sender)
{
  AnsiString Sql, Sql1, nn, inn, vnvi, data_po, data_po1;
  int i=1, rec=0;

  im_fl=5;  // Для выбора имени загружаемого файла

  if (Application->MessageBox(("Вы действительно хотите загрузить данные \n по внешним договорам за " + Mes[DM->mm-1] + " " + DM->yyyy + " года?").c_str(),
                               "Загрузка данных по внешним договорам",
                               MB_YESNO + MB_ICONINFORMATION) == IDNO)
    {
      Abort();
    }

  // Проверка правильности цеха, тн и суммы страхования на превышение 15%
  ProverkaInfoExcel();

  StatusBar1->SimpleText = "";

  try
    {
      Sheet.OleProcedure("Activate");

      Main->Cursor = crHourGlass;
      StatusBar1->SimplePanel = true;    // 2 панели на StatusBar1
      StatusBar1->SimpleText=" Идет загрузка внешних договоров...";

      ProgressBar->Visible = true;
      ProgressBar->Position = 0;
      ProgressBar->Max = Row;

      for ( i ; i<Row+1; i++)
        {
          nn = Excel.OlePropertyGet("Cells",i,1);
          inn = Excel.OlePropertyGet("Cells",i,5);
          ProgressBar->Position++;


          // Выбор строк необходимых для загрузки из Excel
          if (nn.IsEmpty() || !Proverka(nn) || inn.IsEmpty())  continue;

            //Проверка на наличие уже существующих записей в таблице VU_859_N
            Sql1 = "select * from VU_859_N where trim(inn)=trim("+ QuotedStr(Excel.OlePropertyGet("Cells",i,5)) +") \
                                           and trim(n_dogovora) = trim("+ QuotedStr(Excel.OlePropertyGet("Cells",i,9))+")" ;

            try
              {
                DM->qObnovlenie->Close();
                DM->qObnovlenie->SQL->Clear();
                DM->qObnovlenie->SQL->Add(Sql1);
                DM->qObnovlenie->Open();
              }
            catch(...)
              {
                Application->MessageBox("Ошибка получения данных из таблицы по страхованию 859 в/у","Ошибка",MB_OK+ MB_ICONERROR);
                Abort();
              }

            if (DM->qObnovlenie->RecordCount>0)
              {
                 if (Application->MessageBox(("Запись: цех = "+ DM->qObnovlenie->FieldByName("zex")->AsString +
                                               ", таб.№ = "+ DM->qObnovlenie->FieldByName("tn")->AsString +
                                               ", ИНН = "+ DM->qObnovlenie->FieldByName("inn")->AsString +
                                               " и № договора = "+DM->qObnovlenie->FieldByName("n_dogovora")->AsString +
                                              " уже существует. Записать ее еще раз?").c_str(),"Предупреждение",
                                              MB_YESNO + MB_ICONINFORMATION) ==ID_NO)
                    {
                       continue;
                    }
              }

            //Добавление цех+тн из sap_osn_sved
            Sql1="select zex, tn_sap, numident from sap_osn_sved where numident=trim("+ QuotedStr(Excel.OlePropertyGet("Cells",i,5)) +")   \
                  union all                                                                                            \
                  select zex, tn_sap, numident from sap_sved_uvol                                                           \
                  where substr(to_char(dat_job,'dd.mm.yyyy'),4,7)='"+(DM->mm<10 ? "0"+ IntToStr(DM->mm) : IntToStr(DM->mm))+"."+DM->yyyy+"'                                            \
                  and numident=trim("+ QuotedStr(Excel.OlePropertyGet("Cells",i,5)) +")";

            try
              {
                DM->qObnovlenie->Close();
                DM->qObnovlenie->SQL->Clear();
                DM->qObnovlenie->SQL->Add(Sql1);
                DM->qObnovlenie->Open();
              }
            catch(...)
              {
                Application->MessageBox("Ошибка получения данных из таблицы avans","Ошибка",MB_OK+ MB_ICONERROR);
                Abort();
              }

             //Проверка на конечную дату

            data_po = Excel.OlePropertyGet("Cells",i,7);
            data_po1 = Excel.OlePropertyGet("Cells",i,7);

            if ((data_po.SubString(1,2)=="31" && data_po.SubString(4,2)=="04")||
                (data_po.SubString(1,2)=="31" && data_po.SubString(4,2)=="06")||
                (data_po.SubString(1,2)=="31" && data_po.SubString(4,2)=="09")||
                (data_po.SubString(1,2)=="31" && data_po.SubString(4,2)=="11"))
              {
                data_po = "30"+ data_po1.SubString(3,255);
              }

            //Запись данных в таблицу VU_859_N
            Sql = "insert into vu_859_N (zex, tn, fio, n_dogovora, kod_dogovora, data_s, data_po, sum, inn, priznak) \
                   values("+ QuotedStr(DM->qObnovlenie->FieldByName("zex")->AsString)+", \
                          "+ SetNull(DM->qObnovlenie->FieldByName("tn_sap")->AsString)+", \
                          initcap("+ QuotedStr(Excel.OlePropertyGet("Cells",i,2))+"||' '||"+QuotedStr(Excel.OlePropertyGet("Cells",i,3))+"||' '||"+QuotedStr(Excel.OlePropertyGet("Cells",i,4))+"), \
                          trim("+ QuotedStr(Excel.OlePropertyGet("Cells",i,9))+"), \
                             3, \
                          "+ QuotedStr(Excel.OlePropertyGet("Cells",i,6))+", \
                          "+ QuotedStr(data_po)+", \
                          "+ QuotedStr(Excel.OlePropertyGet("Cells",i,8))+", \
                          trim("+ QuotedStr(Excel.OlePropertyGet("Cells",i,5))+"),\
                             0 ) ";
            try
              {
                DM->qZagruzka->Close();
                DM->qZagruzka->SQL->Clear();
                DM->qZagruzka->SQL->Add(Sql);
                DM->qZagruzka->ExecSQL();
                rec++;
              }
            catch(...)
              {
                Application->MessageBox("Ошибка вставки данных в таблицу по страхованию 859 в/у","Ошибка",MB_OK+ MB_ICONERROR);
                Application->MessageBox("Данные не были загружены. Повторите загрузку","Ошибка",MB_OK+ MB_ICONERROR);
                StatusBar1->SimpleText = "";

                Excel.OleProcedure("Quit");
                Abort();
             }
        }



      Application->MessageBox(("Загрузка данных выполнена успешно =) \n Добавлено " + IntToStr(rec) + " записей").c_str(),
                               "Загрузка внешних договоров",MB_OK+ MB_ICONINFORMATION);
      InsertLog("Выполнена загрузка данных по внешним договорам. Загружено "+IntToStr(rec)+" записей");

      Excel.OleProcedure("Quit");
      Excel = Unassigned;

      StatusBar1->SimplePanel = false;
      ProgressBar->Visible = false;
      StatusBar1->SimpleText = "Запись внешних договоров выполнена успешно";
      Main->Cursor = crDefault;
      StatusBar1->SimpleText = "";
    }
  catch(...)
    {
      Application->MessageBox("Ошибка загрузки данных по внешним договорам","Ошибка",MB_OK+ MB_ICONERROR);
      Excel.OleProcedure("Quit");

      StatusBar1->SimplePanel = false;
      ProgressBar->Visible = false;
      StatusBar1->SimpleText = "";
      Main->Cursor = crDefault;
    }

}

//---------------------------------------------------------------------------

// Получение суммы прожиточного минимума
void __fastcall TMain::ProverkaProzhitMin()
{
  AnsiString Sql = " select summn from spiud where mes="+IntToStr(dtp_month)+" \
                                             and god="+dtp_year+"";

  DM->qObnovlenie->Close();
  DM->qObnovlenie->SQL->Clear();
  DM->qObnovlenie->SQL->Add(Sql);
  DM->qObnovlenie->Open();

  prozhitMin = DM->qObnovlenie->FieldByName("summn")->AsFloat;
}
//---------------------------------------------------------------------------

// Для РО
void __fastcall TMain::N13Click(TObject *Sender)
{
  Data->ShowModal();
  if (Data->ModalResult == mrCancel) {Abort();}

  AnsiString Sql,Sql1,sum_pr, tn_pr,tn_pr1;
  int zex1,zex,kust1,kust;
  Double sum_zex=0, sum_kust=0, rl_sum=0, obsh_sum=0;

      /* dtp_month, dtp_year - месяц и год из DateTimePicker
         dtp_mm - месяц из DateTimePicker c "0"
         sum_pr - сумма превышений
         zex1 - предыдущий цех
         zex - текущий цех
         kust1 - предыдущий куст
         kust - текущий куст
         sum_kust - сумма по кусту
         rl_sum - начисленно по расчетке
         obsh_sum - общая сумма по всем договорам вид 859
         tn_pr - текущий таб.№
         tn_pr1 - предыдущий таб.№*/

  //Считывание данных из DateTimePicker
  DecodeDate(Data->DateTimePicker1->Date, dtp_year, dtp_month, dtp_day );

  if (StrToInt(dtp_month)<10)
        {
          dtp_mm ="0"+ IntToStr(dtp_month);
        }
      else
        {
          dtp_mm = IntToStr(dtp_month);
        }


  ProverkaProzhitMin();

  if (!rtf_Open((TempPath + "\\dlya_ro.txt").c_str()))
    {
      MessageBox(Handle,"Ошибка открытия файла данных","Ошибка",8192);
    }
  else
    {
      StatusBar1->SimplePanel = true;    // 2 панели на StatusBar1
      StatusBar1->SimpleText = " Идет формирование гривневой страховки по пром. цехам... ";

      // Формирование гривневой страховки по пром. цехам
      Sql = "select (select fam||' '||im||' ' ||ot from avans  where ncex=sl.zex and tn=sl.tn) as fio ,\
                     sl.zex, sl.tn, sl.sum, sp.kust, sum(sum) over (partition by sl.zex) sum_po_zex,     \
                     sum(sum) over (partition by sp.kust)  sum_po_kust,     \
                     sum(sum) over() sum_po_kombinat                        \
             from slst"+ dtp_mm + dtp_year+" sl, spnc sp            \
             where vo = 859                                         \
             and nvl(nist,0)=0                                      \
             and sl.zex=sp.nc                                       \
             and sp.ana="+ana+"                                           \
             and nvl(sum,0)>0                                       \
             order by kust, zex, tn,sl.sum";

      DM->qZagruzka->Close();
      DM->qZagruzka->SQL->Clear();
      DM->qZagruzka->SQL->Add(Sql);
      try
        {
          DM->qZagruzka->Open();
        }
      catch(...)
       {
         Application->MessageBox("Ошибка получения данных из таблицы SLST. \n ВОЗМОЖНО неверно выбран период.","Ошибка",MB_OK);
         StatusBar1->SimplePanel = false;
         ProgressBar->Visible = false;
         StatusBar1->SimpleText = "";
         Main->Cursor = crDefault;

         Abort();
       }

      Main->Cursor = crHourGlass;
      ProgressBar->Visible = true;
      ProgressBar->Position = 0;
      ProgressBar->Max = DM->qZagruzka->RecordCount;

      zex = DM->qZagruzka->FieldByName("zex")->AsInteger;
      kust = DM->qZagruzka->FieldByName("kust")->AsInteger;
      tn_pr = DM->qZagruzka->FieldByName("tn")->AsInteger;

      while (!DM->qZagruzka->Eof)
        {

          kust1 = DM->qZagruzka->FieldByName("kust")->AsInteger;

          while (!DM->qZagruzka->Eof && kust==kust1)
            {
              
              zex1 = DM->qZagruzka->FieldByName("zex")->AsInteger;


              while (!DM->qZagruzka->Eof && zex==zex1)
                {
                  tn_pr1 = DM->qZagruzka->FieldByName("tn")->AsInteger;
                  int tnn=DM->qZagruzka->FieldByName("tn")->AsInteger;
                  rtf_Out("kust", DM->qZagruzka->FieldByName("kust")->AsString, 1);
                  rtf_Out("zex", DM->qZagruzka->FieldByName("zex")->AsString, 1);
                  rtf_Out("tn", DM->qZagruzka->FieldByName("tn")->AsString, 1);
                  rtf_Out("fio", DM->qZagruzka->FieldByName("fio")->AsString, 1);
                  rtf_Out("sum", DM->qZagruzka->FieldByName("sum")->AsFloat,20,2, 1);

                  //проверка превышения
                  Sql1 = "select (nvl(sum,0) - nvl( (select sum(sum)                                               \
                                                     from slst"+dtp_mm + dtp_year+"                                \
                                                     where zex="+ DM->qZagruzka->FieldByName("zex")->AsString +"   \
                                                     and tn="+ DM->qZagruzka->FieldByName("tn")->AsString +" and vo=576), 0))*0.15 as rl_sum,  \
                                  nvl((select sum(sum)                                                                                                  \
                                       from slst"+dtp_mm + dtp_year+"                                                                                    \
                                       where zex="+  DM->qZagruzka->FieldByName("zex")->AsString +" and tn="+ DM->qZagruzka->FieldByName("tn")->AsString +"   \
                                       and vo=859), 0) as obsh_sum                     \
                                  from slst"+dtp_mm + dtp_year+"                                                           \
                                  where typs=9 and zex="+ DM->qZagruzka->FieldByName("zex")->AsString +"    \
                                  and tn="+ DM->qZagruzka->FieldByName("tn")->AsString;

                  DM->qObnovlenie->Close();
                  DM->qObnovlenie->SQL->Clear();
                  DM->qObnovlenie->SQL->Add(Sql1);
                  try
                    {
                      DM->qObnovlenie->Open();
                    }
                  catch(...)
                    {
                      Application->MessageBox("Ошибка получения данных из таблицы SLST.","Ошибка",MB_OK);
                      StatusBar1->SimplePanel = false;
                      ProgressBar->Visible = false;
                      StatusBar1->SimpleText = "";
                      Main->Cursor = crDefault;

                      Abort();
                    }

                  rl_sum = DM->qObnovlenie->FieldByName("rl_sum")->AsFloat;
                  obsh_sum = DM->qObnovlenie->FieldByName("obsh_sum")->AsFloat;

                  if (rl_sum > prozhitMin || rl_sum == prozhitMin)
                    {
                      if (obsh_sum > prozhitMin)
                        {
                          sum_pr=FloatToStrF(obsh_sum - prozhitMin,ffFixed,20,2);
                        }
                      else
                        {
                          sum_pr = "";
                        }
                    }
                  else
                    {
                      if (obsh_sum > rl_sum)
                        {
                          sum_pr=FloatToStrF(obsh_sum - rl_sum,ffFixed,20,2);
                        }
                      else
                        {
                          sum_pr = "";
                        }

                    }

                  //вывод превышения
                  if (zex==zex1 && tn_pr==tn_pr1)
                    {
                      rtf_Out("prev"," ", 1);

                      if(!rtf_LineFeed())
                        {
                          MessageBox(Handle,"Ошибка записи в файл данных","Ошибка",8192);
                          if (!rtf_Close()) MessageBox(Handle,"Ошибка закрытия файла данных","Ошибка",8192);
                          return;
                        }
                    }
                  else
                    {
                      rtf_Out("prev",sum_pr, 1);

                      if(!rtf_LineFeed())
                        {
                          MessageBox(Handle,"Ошибка записи в файл данных","Ошибка",8192);
                          if (!rtf_Close()) MessageBox(Handle,"Ошибка закрытия файла данных","Ошибка",8192);
                          return;
                        }
                   }

                  sum_zex += DM->qZagruzka->FieldByName("sum")->AsFloat,20,2;
                  sum_kust = DM->qZagruzka->FieldByName("sum_po_kust")->AsFloat;

                  tn_pr = DM->qZagruzka->FieldByName("tn")->AsInteger;
                  DM->qZagruzka->Next();
                  ProgressBar->Position++;

                  zex = DM->qZagruzka->FieldByName("zex")->AsInteger;
                  kust = DM->qZagruzka->FieldByName("kust")->AsInteger;
                  
                }

              //вывод суммы по цеху
              rtf_Out("sum_po_zex", FloatToStrF(sum_zex,ffFixed,20,2),2);

              if(!rtf_LineFeed())
                {
                  MessageBox(Handle,"Ошибка записи в файл данных","Ошибка",8192);
                  if (!rtf_Close()) MessageBox(Handle,"Ошибка закрытия файла данных","Ошибка",8192);
                  return;
                }

              sum_zex=0;
            }

          //вывод суммы по кусту
          rtf_Out("sum_po_kust", FloatToStrF(sum_kust, ffFixed,20,2),3);
          if(!rtf_LineFeed())
            {
              MessageBox(Handle,"Ошибка записи в файл данных","Ошибка",8192);
              if (!rtf_Close()) MessageBox(Handle,"Ошибка закрытия файла данных","Ошибка",8192);
              return;
            }
        }

      // по комбинату
      rtf_Out("sum_po_kombinat",FloatToStrF(DM->qZagruzka->FieldByName("sum_po_kombinat")->AsFloat,ffFixed,20,2), 0);

      StatusBar1->SimpleText = " Идет формирование валютной страховки по пром. цехам... ";

 //*****************************************************************************
      // Формирование валютной страховки по пром. цехам

      Sql = "select (select fam||' '||im||' ' ||ot  from avans where ncex=sl.zex and tn=sl.tn) as fio, \
             sl.zex, sl.tn, sl.sum, sp.kust, sum(sum) over (partition by sl.zex) sum_po_zex,     \
             sum(sum) over (partition by sp.kust)  sum_po_kust,     \
             sum(sum) over() sum_po_kombinat                        \
             from slst"+ dtp_mm + dtp_year+" sl, spnc sp            \
             where vo = 859                                         \
             and nvl(nist,0) in (1,2)                               \
             and sl.zex=sp.nc                                       \
             and sp.ana="+ana+"                                           \
             and nvl(sum,0)>0                                       \
             order by kust, zex, tn,sl.sum";

      DM->qZagruzka->Close();
      DM->qZagruzka->SQL->Clear();
      DM->qZagruzka->SQL->Add(Sql);
      try
        {
          DM->qZagruzka->Open();
        }
      catch(...)
        {
          Application->MessageBox("Ошибка получения данных из таблицы SLST. \n ВОЗМОЖНО неверно выбран период.","Ошибка",MB_OK);
          StatusBar1->SimplePanel = false;
          ProgressBar->Visible = false;
          StatusBar1->SimpleText = "";
          Main->Cursor = crDefault;

          Abort();
        }

      ProgressBar->Position = 0;
      ProgressBar->Max = DM->qZagruzka->RecordCount;

      zex = DM->qZagruzka->FieldByName("zex")->AsInteger;
      kust = DM->qZagruzka->FieldByName("kust")->AsInteger;

      while (!DM->qZagruzka->Eof)
        {

          kust1 = DM->qZagruzka->FieldByName("kust")->AsInteger;

          while (!DM->qZagruzka->Eof && kust==kust1)
            {

              zex1 = DM->qZagruzka->FieldByName("zex")->AsInteger;

              while (!DM->qZagruzka->Eof && zex==zex1)
                {
                  tn_pr1 = DM->qZagruzka->FieldByName("tn")->AsInteger;
                  int tnn=DM->qZagruzka->FieldByName("tn")->AsInteger;
                  rtf_Out("kust", DM->qZagruzka->FieldByName("kust")->AsString, 4);
                  rtf_Out("zex", DM->qZagruzka->FieldByName("zex")->AsString, 4);
                  rtf_Out("tn", DM->qZagruzka->FieldByName("tn")->AsString,4);
                  rtf_Out("fio", DM->qZagruzka->FieldByName("fio")->AsString, 4);
                  rtf_Out("sum", DM->qZagruzka->FieldByName("sum")->AsFloat,20,2, 4);

                  //проверка превышения
                   Sql1 = "select (nvl(sum,0) - nvl( (select sum(sum)                                               \
                                                     from slst"+dtp_mm + dtp_year+"                                \
                                                     where zex="+ DM->qZagruzka->FieldByName("zex")->AsString +"   \
                                                     and tn="+ DM->qZagruzka->FieldByName("tn")->AsString +" and vo=576), 0))*0.15 as rl_sum,  \
                                  nvl((select sum(sum)                                                                                                  \
                                       from slst"+dtp_mm + dtp_year+"                                                                                    \
                                       where zex="+  DM->qZagruzka->FieldByName("zex")->AsString +" and tn="+ DM->qZagruzka->FieldByName("tn")->AsString +"   \
                                       and vo=859), 0) as obsh_sum                     \
                                  from slst"+dtp_mm + dtp_year+"                                                          \
                                  where typs=9 and zex="+ DM->qZagruzka->FieldByName("zex")->AsString +"    \
                                  and tn="+ DM->qZagruzka->FieldByName("tn")->AsString;

                  DM->qObnovlenie->Close();
                  DM->qObnovlenie->SQL->Clear();
                  DM->qObnovlenie->SQL->Add(Sql1);
                  try
                    {
                      DM->qObnovlenie->Open();
                    }
                  catch(...)
                    {
                      Application->MessageBox("Ошибка получения данных из таблицы SLST.","Ошибка",MB_OK);
                      StatusBar1->SimplePanel = false;
                      ProgressBar->Visible = false;
                      StatusBar1->SimpleText = "";
                      Main->Cursor = crDefault;

                      Abort();
                    }

                  rl_sum = DM->qObnovlenie->FieldByName("rl_sum")->AsFloat;
                  obsh_sum = DM->qObnovlenie->FieldByName("obsh_sum")->AsFloat;

                  if (rl_sum > prozhitMin || rl_sum == prozhitMin)
                    {
                      if (obsh_sum > prozhitMin)
                        {
                          sum_pr=FloatToStrF(obsh_sum - prozhitMin,ffFixed,20,2);
                        }
                      else
                        {
                          sum_pr = "";
                        }
                    }
                  else
                    {
                      if (obsh_sum > rl_sum)
                        {
                          sum_pr=FloatToStrF(obsh_sum - rl_sum,ffFixed,20,2);
                        }
                      else
                        {
                          sum_pr = "";
                        }

                    }

                  //вывод превышения
                  if (zex==zex1 && tn_pr==tn_pr1)
                    {
                      rtf_Out("prev"," ", 4);

                      if(!rtf_LineFeed())
                        {
                          MessageBox(Handle,"Ошибка записи в файл данных","Ошибка",8192);
                          if (!rtf_Close()) MessageBox(Handle,"Ошибка закрытия файла данных","Ошибка",8192);
                          return;
                        }
                    }
                  else
                    {
                      rtf_Out("prev",sum_pr, 4);

                      if(!rtf_LineFeed())
                        {
                          MessageBox(Handle,"Ошибка записи в файл данных","Ошибка",8192);
                          if (!rtf_Close()) MessageBox(Handle,"Ошибка закрытия файла данных","Ошибка",8192);
                          return;
                        }
                    }

                  sum_zex += DM->qZagruzka->FieldByName("sum")->AsFloat,20,2;
                  sum_kust = DM->qZagruzka->FieldByName("sum_po_kust")->AsFloat;

                  tn_pr = DM->qZagruzka->FieldByName("tn")->AsInteger;
                  DM->qZagruzka->Next();
                  ProgressBar->Position++;

                  zex = DM->qZagruzka->FieldByName("zex")->AsInteger;
                  kust = DM->qZagruzka->FieldByName("kust")->AsInteger;
                }

              //вывод суммы по цеху
              rtf_Out("sum_po_zex", FloatToStrF(sum_zex,ffFixed,20,2),5);

              if(!rtf_LineFeed())
                {
                  MessageBox(Handle,"Ошибка записи в файл данных","Ошибка",8192);
                  if (!rtf_Close()) MessageBox(Handle,"Ошибка закрытия файла данных","Ошибка",8192);
                  return;
                }

              sum_zex=0;
            }

           //вывод суммы по кусту
           rtf_Out("sum_po_kust", FloatToStrF(sum_kust, ffFixed,20,2),6);
           if(!rtf_LineFeed())
             {
               MessageBox(Handle,"Ошибка записи в файл данных","Ошибка",8192);
               if (!rtf_Close()) MessageBox(Handle,"Ошибка закрытия файла данных","Ошибка",8192);
               return;
             }

        }

      // по комбинату
      rtf_Out("sum_po_kombinat",FloatToStrF(DM->qZagruzka->FieldByName("sum_po_kombinat")->AsFloat,ffFixed,20,2), 0);

      StatusBar1->SimpleText = " Идет формирование внешних договоров по пром. цехам... ";

//*****************************************************************************
// Внешние договора по пром. цехам

      Sql = "select (select fam||' '||im||' ' ||ot  from avans where ncex=sl.zex and tn=sl.tn) as fio, \
                     sl.zex, sl.tn, sl.sum, sp.kust, sum(sum) over (partition by sl.zex) sum_po_zex,     \
                     sum(sum) over (partition by sp.kust)  sum_po_kust,     \
                     sum(sum) over() sum_po_kombinat                        \
             from slst"+ dtp_mm + dtp_year+" sl, spnc sp            \
             where vo = 859                                         \
             and nvl(nist,0) = 3                               \
             and sl.zex=sp.nc                                       \
             and sp.ana="+ana+" \
             and nvl(sum,0)>0                                           \
             order by kust, zex, tn,sl.sum";

      DM->qZagruzka->Close();
      DM->qZagruzka->SQL->Clear();
      DM->qZagruzka->SQL->Add(Sql);
      try
        {
          DM->qZagruzka->Open();
        }
      catch(...)
        {
          Application->MessageBox("Ошибка получения данных из таблицы SLST.","Ошибка",MB_OK);

          StatusBar1->SimplePanel = false;
          ProgressBar->Visible = false;
          StatusBar1->SimpleText = "";
          Main->Cursor = crDefault;

          Abort();
        }

      ProgressBar->Position = 0;
      ProgressBar->Max = DM->qZagruzka->RecordCount;

      zex = DM->qZagruzka->FieldByName("zex")->AsInteger;
      kust = DM->qZagruzka->FieldByName("kust")->AsInteger;

      while (!DM->qZagruzka->Eof)
        {

          kust1 = DM->qZagruzka->FieldByName("kust")->AsInteger;

          while (!DM->qZagruzka->Eof && kust==kust1)
            {

              zex1 = DM->qZagruzka->FieldByName("zex")->AsInteger;

              while (!DM->qZagruzka->Eof && zex==zex1)
                {
                  tn_pr1 = DM->qZagruzka->FieldByName("tn")->AsInteger;
                  int tnn=DM->qZagruzka->FieldByName("tn")->AsInteger;
                  rtf_Out("kust", DM->qZagruzka->FieldByName("kust")->AsString, 7);
                  rtf_Out("zex", DM->qZagruzka->FieldByName("zex")->AsString, 7);
                  rtf_Out("tn", DM->qZagruzka->FieldByName("tn")->AsString,7);
                  rtf_Out("fio", DM->qZagruzka->FieldByName("fio")->AsString, 7);
                  rtf_Out("sum", DM->qZagruzka->FieldByName("sum")->AsFloat,20,2, 7);

                  //проверка превышения
                   Sql1 = "select (nvl(sum,0) - nvl( (select sum(sum)                                               \
                                                     from slst"+dtp_mm + dtp_year+"                                \
                                                     where zex="+ DM->qZagruzka->FieldByName("zex")->AsString +"   \
                                                     and tn="+ DM->qZagruzka->FieldByName("tn")->AsString +" and vo=576), 0))*0.15 as rl_sum,  \
                                  nvl((select sum(sum)                                                                                                  \
                                       from slst"+dtp_mm + dtp_year+"                                                                                    \
                                       where zex="+  DM->qZagruzka->FieldByName("zex")->AsString +" and tn="+ DM->qZagruzka->FieldByName("tn")->AsString +"   \
                                       and vo=859), 0) as obsh_sum                     \
                                  from slst"+dtp_mm + dtp_year+"                                                          \
                                  where typs=9 and zex="+ DM->qZagruzka->FieldByName("zex")->AsString +"    \
                                  and tn="+ DM->qZagruzka->FieldByName("tn")->AsString;

                  DM->qObnovlenie->Close();
                  DM->qObnovlenie->SQL->Clear();
                  DM->qObnovlenie->SQL->Add(Sql1);
                  try
                    {
                      DM->qObnovlenie->Open();
                    }
                  catch(...)
                    {
                      Application->MessageBox("Ошибка получения данных из таблицы SLST.","Ошибка",MB_OK);

                      StatusBar1->SimplePanel = false;
                      ProgressBar->Visible = false;
                      StatusBar1->SimpleText = "";
                      Main->Cursor = crDefault;

                      Abort();
                    }

                  rl_sum = DM->qObnovlenie->FieldByName("rl_sum")->AsFloat;
                  obsh_sum = DM->qObnovlenie->FieldByName("obsh_sum")->AsFloat;

                  if (rl_sum > prozhitMin || rl_sum == prozhitMin)
                    {
                      if (obsh_sum > prozhitMin)
                        {
                          sum_pr=FloatToStrF(obsh_sum - prozhitMin,ffFixed,20,2);
                        }
                      else
                        {
                          sum_pr = "";
                        }
                    }
                  else
                    {
                      if (obsh_sum > rl_sum)
                        {
                          sum_pr=FloatToStrF(obsh_sum - rl_sum,ffFixed,20,2);
                        }
                      else
                        {
                          sum_pr = "";
                        }

                    }

                  //вывод превышения
                  if (zex==zex1 && tn_pr==tn_pr1)
                    {
                      rtf_Out("prev"," ", 7);

                      if(!rtf_LineFeed())
                        {
                          MessageBox(Handle,"Ошибка записи в файл данных","Ошибка",8192);
                          if (!rtf_Close()) MessageBox(Handle,"Ошибка закрытия файла данных","Ошибка",8192);
                          return;
                        }
                    }
                  else
                    {
                      rtf_Out("prev",sum_pr, 7);

                      if(!rtf_LineFeed())
                        {
                          MessageBox(Handle,"Ошибка записи в файл данных","Ошибка",8192);
                          if (!rtf_Close()) MessageBox(Handle,"Ошибка закрытия файла данных","Ошибка",8192);
                          return;
                        }
                    }

                  sum_zex += DM->qZagruzka->FieldByName("sum")->AsFloat,20,2;
                  sum_kust = DM->qZagruzka->FieldByName("sum_po_kust")->AsFloat;

                  tn_pr = DM->qZagruzka->FieldByName("tn")->AsInteger;
                  DM->qZagruzka->Next();
                  ProgressBar->Position++;

                  zex = DM->qZagruzka->FieldByName("zex")->AsInteger;
                  kust = DM->qZagruzka->FieldByName("kust")->AsInteger;
                }

              //вывод суммы по цеху
              rtf_Out("sum_po_zex", FloatToStrF(sum_zex,ffFixed,20,2),8);

              if(!rtf_LineFeed())
                {
                  MessageBox(Handle,"Ошибка записи в файл данных","Ошибка",8192);
                  if (!rtf_Close()) MessageBox(Handle,"Ошибка закрытия файла данных","Ошибка",8192);
                  return;
                }

              sum_zex=0;
            }

           //вывод суммы по кусту
           rtf_Out("sum_po_kust", FloatToStrF(sum_kust, ffFixed,20,2),9);
           if(!rtf_LineFeed())
             {
               MessageBox(Handle,"Ошибка записи в файл данных","Ошибка",8192);
               if (!rtf_Close()) MessageBox(Handle,"Ошибка закрытия файла данных","Ошибка",8192);
               return;
             }

        }

  // по комбинату
  rtf_Out("sum_po_kombinat",FloatToStrF(DM->qZagruzka->FieldByName("sum_po_kombinat")->AsFloat,ffFixed,20,2), 0);

  StatusBar1->SimplePanel = false;
  ProgressBar->Visible = false;
  StatusBar1->SimpleText = "";
  Main->Cursor = crDefault;


  if(!rtf_Close())
    {
      MessageBox(Handle,"Ошибка закрытия файла данных", "Ошибка", 8192);
      return;
    }

  //Создание папки, если ее не существует
  ForceDirectories(WorkPath);

  int istrd;
  try
    {
      rtf_CreateReport(TempPath +"\\dlya_ro.txt", Path+"\\RTF\\dlya_ro.rtf",
                       WorkPath+"\\Для РО(пром.цеха).doc",NULL,&istrd);


      WinExec(("\""+ WordPath+"\"\""+WorkPath+"\\Для РО(пром.цеха).doc\"").c_str(),SW_MAXIMIZE);
    }
  catch(RepoRTF_Error E)
    {
      MessageBox(Handle,("Ошибка формирования отчета:"+ AnsiString(E.Err)+
                         "\nСтрока файла данных:"+IntToStr(istrd)).c_str(),"Ошибка",8192);
    }
    StatusBar1->SimpleText = " ";
  }

}

//---------------------------------------------------------------------------

// Управление С\Х
void __fastcall TMain::N15Click(TObject *Sender)
{
  Data->ShowModal();
  if (Data->ModalResult == mrCancel) {Abort();}

  AnsiString Sql,Sql1, sum_pr;
  int zex1,zex,kust1,kust;
  Double sum_zex=0, sum_kust=0, rl_sum=0, obsh_sum=0;

      /* dtp_month, dtp_year - месяц и год из DateTimePicker
         dtp_mm - месяц из DateTimePicker c "0"
         sum_pr - сумма превышений*/

  //Считывание данных из DateTimePicker
  DecodeDate(Data->DateTimePicker1->Date, dtp_year, dtp_month, dtp_day );

  if (StrToInt(dtp_month)<10)
        {
          dtp_mm ="0"+ IntToStr(dtp_month);
        }
      else
        {
          dtp_mm = IntToStr(dtp_month);
        }


  ProverkaProzhitMin();

   // Формирование гривневой страховки по управлению с\х
   Sql = "select (select fam||' '||im||' ' ||ot from avans  where ncex=sl.zex and tn=sl.tn) as fio ,\
                  sl.zex, sl.tn, sl.sum, sp.kust, sum(sum) over (partition by sl.zex) sum_po_zex,     \
                  sum(sum) over (partition by sp.kust)  sum_po_kust,     \
                  sum(sum) over() sum_po_kombinat                        \
          from slst"+ dtp_mm + dtp_year+" sl, spnc sp            \
          where vo = 859                                         \
          and nvl(nist,0)=0                                      \
          and sl.zex=sp.nc                                       \
          and sp.ana=6                                            \
          and nvl(sum,0)>0                                         \
          order by kust, zex, tn,sl.sum";

  DM->qZagruzka->Close();
  DM->qZagruzka->SQL->Clear();
  DM->qZagruzka->SQL->Add(Sql);
  try
    {
      DM->qZagruzka->Open();
    }
  catch(...)
    {
      Application->MessageBox("Ошибка получения данных из таблицы SLST. \n ВОЗМОЖНО неверно выбран период.","Ошибка",MB_OK);
      Abort();
    }
  //Открытие файла данных содержащего уволенных, окончанием выплат и измененным цехом и таб.№
  if (!rtf_Open((TempPath + "\\dlya_sh.txt").c_str()))
    {
      MessageBox(Handle,"Ошибка открытия файла данных","Ошибка",8192);
    }
  else
    {
      Main->Cursor = crHourGlass;
      ProgressBar->Visible = true;
      ProgressBar->Position = 0;
      ProgressBar->Max = DM->qZagruzka->RecordCount;

      zex = DM->qZagruzka->FieldByName("zex")->AsInteger;
      kust = DM->qZagruzka->FieldByName("kust")->AsInteger;

      while (!DM->qZagruzka->Eof)
        {

          kust1 = DM->qZagruzka->FieldByName("kust")->AsInteger;

          while (!DM->qZagruzka->Eof && kust==kust1)
            {

              zex1 = DM->qZagruzka->FieldByName("zex")->AsInteger;

              while (!DM->qZagruzka->Eof && zex==zex1)
                {
                  rtf_Out("kust", DM->qZagruzka->FieldByName("kust")->AsString, 1);
                  rtf_Out("zex", DM->qZagruzka->FieldByName("zex")->AsString, 1);
                  rtf_Out("tn", DM->qZagruzka->FieldByName("tn")->AsString, 1);
                  rtf_Out("fio", DM->qZagruzka->FieldByName("fio")->AsString, 1);
                  rtf_Out("sum", DM->qZagruzka->FieldByName("sum")->AsFloat,20,2, 1);

                  //проверка превышения
                  Sql1 = "select (nvl(sum,0) - nvl( (select sum(sum)                                               \
                                                     from slst"+dtp_mm + dtp_year+"                                \
                                                     where zex="+ DM->qZagruzka->FieldByName("zex")->AsString +"   \
                                                     and tn="+ DM->qZagruzka->FieldByName("tn")->AsString +" and vo=576), 0))*0.15 as rl_sum,  \
                                  nvl((select sum(sum)                                                                                                  \
                                       from slst"+dtp_mm + dtp_year+"                                                                                    \
                                       where zex="+  DM->qZagruzka->FieldByName("zex")->AsString +" and tn="+ DM->qZagruzka->FieldByName("tn")->AsString +"   \
                                       and vo=859), 0) as obsh_sum                     \
                                  from slst"+dtp_mm + dtp_year+"                                                          \
                                  where typs=9 and zex="+ DM->qZagruzka->FieldByName("zex")->AsString +"    \
                                  and tn="+ DM->qZagruzka->FieldByName("tn")->AsString;

                  DM->qObnovlenie->Close();
                  DM->qObnovlenie->SQL->Clear();
                  DM->qObnovlenie->SQL->Add(Sql1);
                  try
                    {
                      DM->qObnovlenie->Open();
                    }
                  catch(...)
                    {
                      Application->MessageBox("Ошибка получения данных из таблицы SLST.","Ошибка",MB_OK);
                      StatusBar1->SimplePanel = false;
                      ProgressBar->Visible = false;
                      StatusBar1->SimpleText = "";
                      Main->Cursor = crDefault;

                      Abort();
                    }

                  rl_sum = DM->qObnovlenie->FieldByName("rl_sum")->AsFloat;
                  obsh_sum = DM->qObnovlenie->FieldByName("obsh_sum")->AsFloat;

                  if (rl_sum > prozhitMin || rl_sum == prozhitMin)
                    {
                      if (obsh_sum > prozhitMin)
                        {
                          sum_pr=FloatToStrF(obsh_sum - prozhitMin,ffFixed,20,2);
                        }
                      else
                        {
                          sum_pr = "";
                        }
                    }
                  else
                    {
                      if (obsh_sum > rl_sum)
                        {
                          sum_pr=FloatToStrF(obsh_sum - rl_sum,ffFixed,20,2);
                        }
                      else
                        {
                          sum_pr = "";
                        }

                    }

                  //вывод превышения
                  rtf_Out("prev",sum_pr, 1);

                  if(!rtf_LineFeed())
                    {
                      MessageBox(Handle,"Ошибка записи в файл данных","Ошибка",8192);
                      if (!rtf_Close()) MessageBox(Handle,"Ошибка закрытия файла данных","Ошибка",8192);
                      return;
                    }

                  sum_zex += DM->qZagruzka->FieldByName("sum")->AsFloat,20,2;
                  sum_kust = DM->qZagruzka->FieldByName("sum_po_kust")->AsFloat;


                  DM->qZagruzka->Next();
                  ProgressBar->Position++;

                  zex = DM->qZagruzka->FieldByName("zex")->AsInteger;
                  kust = DM->qZagruzka->FieldByName("kust")->AsInteger;
                }

              //вывод суммы по цеху
              rtf_Out("sum_po_zex", FloatToStrF(sum_zex,ffFixed,20,2),2);

              if(!rtf_LineFeed())
                {
                  MessageBox(Handle,"Ошибка записи в файл данных","Ошибка",8192);
                  if (!rtf_Close()) MessageBox(Handle,"Ошибка закрытия файла данных","Ошибка",8192);
                  return;
                }

              sum_zex=0;
            }

          //вывод суммы по кусту
          rtf_Out("sum_po_kust", FloatToStrF(sum_kust, ffFixed,20,2),3);
          if(!rtf_LineFeed())
            {
              MessageBox(Handle,"Ошибка записи в файл данных","Ошибка",8192);
              if (!rtf_Close()) MessageBox(Handle,"Ошибка закрытия файла данных","Ошибка",8192);
              return;
            }
        }

      // по комбинату
      rtf_Out("sum_po_kombinat",FloatToStrF(DM->qZagruzka->FieldByName("sum_po_kombinat")->AsFloat,ffFixed,20,2), 0);

       StatusBar1->SimpleText = " Идет формирование валютной страховки по управлению с\х ";

 //*****************************************************************************
// Формирование валютной страховки по управлению с/х
       Sql = "select (select fam||' '||im||' ' ||ot from avans  where ncex=sl.zex and tn=sl.tn) as fio ,\
                  sl.zex, sl.tn, sl.sum, sp.kust, sum(sum) over (partition by sl.zex) sum_po_zex,     \
                  sum(sum) over (partition by sp.kust)  sum_po_kust,     \
                  sum(sum) over() sum_po_kombinat                        \
              from slst"+ dtp_mm + dtp_year+" sl, spnc sp            \
              where vo = 859                                         \
              and nvl(nist,0) in (1,2)                                      \
              and sl.zex=sp.nc                                       \
              and sp.ana=6     \
              and nvl(sum, 0)>0                                           \
              order by kust, zex, tn,sl.sum";

       DM->qZagruzka->Close();
       DM->qZagruzka->SQL->Clear();
       DM->qZagruzka->SQL->Add(Sql);
       try
         {
           DM->qZagruzka->Open();
         }
       catch(...)
         {
           Application->MessageBox("Ошибка получения данных из таблицы SLST. \n ВОЗМОЖНО неверно выбран период.","Ошибка",MB_OK);
           StatusBar1->SimplePanel = false;
           ProgressBar->Visible = false;
           StatusBar1->SimpleText = "";
           Main->Cursor = crDefault;

           Abort();
         }

      ProgressBar->Position = 0;
      ProgressBar->Max = DM->qZagruzka->RecordCount;
     
      zex = DM->qZagruzka->FieldByName("zex")->AsInteger;
      kust = DM->qZagruzka->FieldByName("kust")->AsInteger;

      while (!DM->qZagruzka->Eof)
        {

          kust1 = DM->qZagruzka->FieldByName("kust")->AsInteger;

          while (!DM->qZagruzka->Eof && kust==kust1)
            {

              zex1 = DM->qZagruzka->FieldByName("zex")->AsInteger;

              while (!DM->qZagruzka->Eof && zex==zex1)
                {
                  rtf_Out("kust", DM->qZagruzka->FieldByName("kust")->AsString, 4);
                  rtf_Out("zex", DM->qZagruzka->FieldByName("zex")->AsString, 4);
                  rtf_Out("tn", DM->qZagruzka->FieldByName("tn")->AsString, 4);
                  rtf_Out("fio", DM->qZagruzka->FieldByName("fio")->AsString,4);
                  rtf_Out("sum", DM->qZagruzka->FieldByName("sum")->AsFloat,20,2, 4);

                  //проверка превышения
                  Sql1 = "select (nvl(sum,0) - nvl( (select sum(sum)                                               \
                                                     from slst"+dtp_mm + dtp_year+"                                \
                                                     where zex="+ DM->qZagruzka->FieldByName("zex")->AsString +"   \
                                                     and tn="+ DM->qZagruzka->FieldByName("tn")->AsString +" and vo=576), 0))*0.15 as rl_sum,  \
                                  nvl((select sum(sum)                                                                                                  \
                                       from slst"+dtp_mm + dtp_year+"                                                                                    \
                                       where zex="+  DM->qZagruzka->FieldByName("zex")->AsString +" and tn="+ DM->qZagruzka->FieldByName("tn")->AsString +"   \
                                       and vo=859), 0) as obsh_sum                     \
                                  from slst"+dtp_mm + dtp_year+"                                                          \
                                  where typs=9 and zex="+ DM->qZagruzka->FieldByName("zex")->AsString +"    \
                                  and tn="+ DM->qZagruzka->FieldByName("tn")->AsString;

                  DM->qObnovlenie->Close();
                  DM->qObnovlenie->SQL->Clear();
                  DM->qObnovlenie->SQL->Add(Sql1);
                  try
                    {
                      DM->qObnovlenie->Open();
                    }
                  catch(...)
                    {
                      Application->MessageBox("Ошибка получения данных из таблицы SLST.","Ошибка",MB_OK);
                      StatusBar1->SimplePanel = false;
                      ProgressBar->Visible = false;
                      StatusBar1->SimpleText = "";
                      Main->Cursor = crDefault;

                      Abort();
                    }

                  rl_sum = DM->qObnovlenie->FieldByName("rl_sum")->AsFloat;
                  obsh_sum = DM->qObnovlenie->FieldByName("obsh_sum")->AsFloat;

                  if (rl_sum > prozhitMin || rl_sum == prozhitMin)
                    {
                      if (obsh_sum > prozhitMin)
                        {
                          sum_pr=FloatToStrF(obsh_sum - prozhitMin,ffFixed,20,2);
                        }
                      else
                        {
                          sum_pr = "";
                        }
                    }
                  else
                    {
                      if (obsh_sum > rl_sum)
                        {
                          sum_pr=FloatToStrF(obsh_sum - rl_sum,ffFixed,20,2);
                        }
                      else
                        {
                          sum_pr = "";
                        }

                    }

                  //вывод превышения
                  rtf_Out("prev",sum_pr, 4);

                  if(!rtf_LineFeed())
                    {
                      MessageBox(Handle,"Ошибка записи в файл данных","Ошибка",8192);
                      if (!rtf_Close()) MessageBox(Handle,"Ошибка закрытия файла данных","Ошибка",8192);
                      return;
                    }

                  sum_zex += DM->qZagruzka->FieldByName("sum")->AsFloat,20,2;
                  sum_kust = DM->qZagruzka->FieldByName("sum_po_kust")->AsFloat;


                  DM->qZagruzka->Next();
                  ProgressBar->Position++;

                  zex = DM->qZagruzka->FieldByName("zex")->AsInteger;
                  kust = DM->qZagruzka->FieldByName("kust")->AsInteger;
                }

              //вывод суммы по цеху
              rtf_Out("sum_po_zex", FloatToStrF(sum_zex,ffFixed,20,2),5);

              if(!rtf_LineFeed())
                {
                  MessageBox(Handle,"Ошибка записи в файл данных","Ошибка",8192);
                  if (!rtf_Close()) MessageBox(Handle,"Ошибка закрытия файла данных","Ошибка",8192);
                  return;
                }

              sum_zex=0;
            }

          //вывод суммы по кусту
          rtf_Out("sum_po_kust", FloatToStrF(sum_kust, ffFixed,20,2),6);
          if(!rtf_LineFeed())
            {
              MessageBox(Handle,"Ошибка записи в файл данных","Ошибка",8192);
              if (!rtf_Close()) MessageBox(Handle,"Ошибка закрытия файла данных","Ошибка",8192);
              return;
            }
        }

      // по комбинату
      rtf_Out("sum_po_kombinat",FloatToStrF(DM->qZagruzka->FieldByName("sum_po_kombinat")->AsFloat,ffFixed,20,2), 0);

       StatusBar1->SimpleText = " Идет формирование валютной страховки по управлению с\х";

 //*****************************************************************************
   // Формирование внешних договоров по управлению с\х
      Sql = "select (select fam||' '||im||' ' ||ot from avans  where ncex=sl.zex and tn=sl.tn) as fio ,\
                  sl.zex, sl.tn, sl.sum, sp.kust, sum(sum) over (partition by sl.zex) sum_po_zex,     \
                  sum(sum) over (partition by sp.kust)  sum_po_kust,     \
                  sum(sum) over() sum_po_kombinat                        \
             from slst"+ dtp_mm + dtp_year+" sl, spnc sp            \
             where vo = 859                                         \
             and nvl(nist,0)=3                                      \
             and sl.zex=sp.nc                                       \
             and sp.ana=6  \
             and nvl(sum,0)>0                                         \
             order by kust, zex, tn,sl.sum";

      DM->qZagruzka->Close();
      DM->qZagruzka->SQL->Clear();
      DM->qZagruzka->SQL->Add(Sql);
      try
        {
          DM->qZagruzka->Open();
        }
      catch(...)
        {
          Application->MessageBox("Ошибка получения данных из таблицы SLST. \n ВОЗМОЖНО неверно выбран период.","Ошибка",MB_OK);
          StatusBar1->SimplePanel = false;
          ProgressBar->Visible = false;
          StatusBar1->SimpleText = "";
          Main->Cursor = crDefault;

          Abort();
        }

      ProgressBar->Position = 0;
      ProgressBar->Max = DM->qZagruzka->RecordCount;

      zex = DM->qZagruzka->FieldByName("zex")->AsInteger;
      kust = DM->qZagruzka->FieldByName("kust")->AsInteger;

      while (!DM->qZagruzka->Eof)
        {

          kust1 = DM->qZagruzka->FieldByName("kust")->AsInteger;

          while (!DM->qZagruzka->Eof && kust==kust1)
            {

              zex1 = DM->qZagruzka->FieldByName("zex")->AsInteger;

              while (!DM->qZagruzka->Eof && zex==zex1)
                {
                  rtf_Out("kust", DM->qZagruzka->FieldByName("kust")->AsString, 7);
                  rtf_Out("zex", DM->qZagruzka->FieldByName("zex")->AsString, 7);
                  rtf_Out("tn", DM->qZagruzka->FieldByName("tn")->AsString, 7);
                  rtf_Out("fio", DM->qZagruzka->FieldByName("fio")->AsString, 7);
                  rtf_Out("sum", DM->qZagruzka->FieldByName("sum")->AsFloat,20,2, 7);

                  //проверка превышения
                  Sql1 = "select (nvl(sum,0) - nvl( (select sum(sum)                                               \
                                                     from slst"+dtp_mm + dtp_year+"                                \
                                                     where zex="+ DM->qZagruzka->FieldByName("zex")->AsString +"   \
                                                     and tn="+ DM->qZagruzka->FieldByName("tn")->AsString +" and vo=576), 0))*0.15 as rl_sum,  \
                                  nvl((select sum(sum)                                                                                                  \
                                       from slst"+dtp_mm + dtp_year+"                                                                                    \
                                       where zex="+  DM->qZagruzka->FieldByName("zex")->AsString +" and tn="+ DM->qZagruzka->FieldByName("tn")->AsString +"   \
                                       and vo=859), 0) as obsh_sum                     \
                                  from slst"+dtp_mm + dtp_year+"                                                         \
                                  where typs=9 and zex="+ DM->qZagruzka->FieldByName("zex")->AsString +"    \
                                  and tn="+ DM->qZagruzka->FieldByName("tn")->AsString;

                  DM->qObnovlenie->Close();
                  DM->qObnovlenie->SQL->Clear();
                  DM->qObnovlenie->SQL->Add(Sql1);
                  try
                    {
                      DM->qObnovlenie->Open();
                    }
                  catch(...)
                    {
                      Application->MessageBox("Ошибка получения данных из таблицы SLST.","Ошибка",MB_OK);
                      StatusBar1->SimplePanel = false;
                      ProgressBar->Visible = false;
                      StatusBar1->SimpleText = "";
                      Main->Cursor = crDefault;

                      Abort();
                    }

                  rl_sum = DM->qObnovlenie->FieldByName("rl_sum")->AsFloat;
                  obsh_sum = DM->qObnovlenie->FieldByName("obsh_sum")->AsFloat;

                  if (rl_sum > prozhitMin || rl_sum == prozhitMin)
                    {
                      if (obsh_sum > prozhitMin)
                        {
                          sum_pr=FloatToStrF(obsh_sum - prozhitMin,ffFixed,20,2);
                        }
                      else
                        {
                          sum_pr = "";
                        }
                    }
                  else
                    {
                      if (obsh_sum > rl_sum)
                        {
                          sum_pr=FloatToStrF(obsh_sum - rl_sum,ffFixed,20,2);
                        }
                      else
                        {
                          sum_pr = "";
                        }

                    }

                  //вывод превышения
                  rtf_Out("prev",sum_pr, 7);

                  if(!rtf_LineFeed())
                    {
                      MessageBox(Handle,"Ошибка записи в файл данных","Ошибка",8192);
                      if (!rtf_Close()) MessageBox(Handle,"Ошибка закрытия файла данных","Ошибка",8192);
                      return;
                    }

                  sum_zex += DM->qZagruzka->FieldByName("sum")->AsFloat,20,2;
                  sum_kust = DM->qZagruzka->FieldByName("sum_po_kust")->AsFloat;


                  DM->qZagruzka->Next();
                  ProgressBar->Position++;

                  zex = DM->qZagruzka->FieldByName("zex")->AsInteger;
                  kust = DM->qZagruzka->FieldByName("kust")->AsInteger;
                }

              //вывод суммы по цеху
              rtf_Out("sum_po_zex", FloatToStrF(sum_zex,ffFixed,20,2),8);

              if(!rtf_LineFeed())
                {
                  MessageBox(Handle,"Ошибка записи в файл данных","Ошибка",8192);
                  if (!rtf_Close()) MessageBox(Handle,"Ошибка закрытия файла данных","Ошибка",8192);
                  return;
                }

              sum_zex=0;
            }

          //вывод суммы по кусту
          rtf_Out("sum_po_kust", FloatToStrF(sum_kust, ffFixed,20,2),9);
          if(!rtf_LineFeed())
            {
              MessageBox(Handle,"Ошибка записи в файл данных","Ошибка",8192);
              if (!rtf_Close()) MessageBox(Handle,"Ошибка закрытия файла данных","Ошибка",8192);
              return;
            }
        }

      // по комбинату
      rtf_Out("sum_po_kombinat",FloatToStrF(DM->qZagruzka->FieldByName("sum_po_kombinat")->AsFloat,ffFixed,20,2), 0);

  StatusBar1->SimplePanel = false;
  ProgressBar->Visible = false;
  StatusBar1->SimpleText = "";
  Main->Cursor = crDefault;


  if(!rtf_Close())
    {
      MessageBox(Handle,"Ошибка закрытия файла данных", "Ошибка", 8192);
      return;
    }

  //Создание папки, если ее не существует
  ForceDirectories(WorkPath);

  int istrd;
  try
    {
      rtf_CreateReport(TempPath +"\\dlya_sh.txt", Path+"\\RTF\\dlya_sh.rtf",
                       WorkPath+"\\Для Управление С\Х.doc",NULL,&istrd);


      WinExec(("\""+ WordPath+"\"\""+WorkPath+"\\Для Управление С\Х.doc\"").c_str(),SW_MAXIMIZE);
    }
  catch(RepoRTF_Error E)
    {
      MessageBox(Handle,("Ошибка формирования отчета:"+ AnsiString(E.Err)+
                         "\nСтрока файла данных:"+IntToStr(istrd)).c_str(),"Ошибка",8192);
    }

    }
   
   StatusBar1->SimpleText = " ";
}
//---------------------------------------------------------------------------

//Агрофирмы
void __fastcall TMain::N17Click(TObject *Sender)
{
  Data->ShowModal();
  if (Data->ModalResult == mrCancel) {Abort();}

  AnsiString Sql,Sql1, sum_pr,firma, firma1, tn_pr, tn_pr1;
  int zex1,zex, ana, ana1;
  Double sum_zex=0, sum_kust=0, rl_sum=0, obsh_sum=0;

      /* dtp_month, dtp_year - месяц и год из DateTimePicker
         dtp_mm - месяц из DateTimePicker c "0"
         sum_pr - сумма превышений*/

  //Считывание данных из DateTimePicker
  DecodeDate(Data->DateTimePicker1->Date, dtp_year, dtp_month, dtp_day );

  if (StrToInt(dtp_month)<10)
        {
          dtp_mm ="0"+ IntToStr(dtp_month);
        }
      else
        {
          dtp_mm = IntToStr(dtp_month);
        }


  ProverkaProzhitMin();

   StatusBar1->SimpleText = " Идет формирование гривневой страховки по РМЗ...";

     // Формирование гривневой страховки по агрофирмам
   Sql = "select (select fam||' '||im||' ' ||ot from avans  where ncex=sl.zex and tn=sl.tn) as fio ,    \
                  sl.zex, sl.tn, sl.sum, sp.kust, nvl(sp.firma, 0)as firma, sp.ana,                                    \
                  sum(sum) over (partition by sl.zex) sum_po_zex,                                      \
                  sum(sum) over (partition by sp.firma)  sum_po_kust,                                   \
                  sum(sum) over() sum_po_kombinat                                                      \
          from slst"+dtp_mm + dtp_year+" sl, spnc sp                                                   \
          where vo = 859                                                                              \
          and nvl(nist,0)=0                                                                            \
          and sl.zex=sp.nc                                                                             \
          and sp.ana between 2 and 10 and ana not in (6,7)                                                               \
          and nvl(sum,0)>0                                                                             \
          group by sp.firma, sp.kust, sl.zex, sl.tn, sl.sum, sp.ana                                            \
          order by firma, kust, zex, tn";

  DM->qZagruzka->Close();
  DM->qZagruzka->SQL->Clear();
  DM->qZagruzka->SQL->Add(Sql);
  try
    {
      DM->qZagruzka->Open();
    }
  catch(...)
    {
      Application->MessageBox("Ошибка получения данных из таблицы SLST. \n ВОЗМОЖНО неверно выбран период.","Ошибка",MB_OK);
      StatusBar1->SimpleText = "";
      Abort();
    }
  //Открытие файла данных
  if (!rtf_Open((TempPath + "\\dlya_agro.txt").c_str()))
    {
      MessageBox(Handle,"Ошибка открытия файла данных","Ошибка",8192);
    }
  else
    {
      Main->Cursor = crHourGlass;
      ProgressBar->Visible = true;
      ProgressBar->Position = 0;
      ProgressBar->Max = DM->qZagruzka->RecordCount;

      zex = DM->qZagruzka->FieldByName("zex")->AsInteger;
      firma = DM->qZagruzka->FieldByName("firma")->AsString;


      while (!DM->qZagruzka->Eof)
        {

          firma1 = DM->qZagruzka->FieldByName("firma")->AsString;

          while (!DM->qZagruzka->Eof && firma==firma1)
            {

              zex1 = DM->qZagruzka->FieldByName("zex")->AsInteger;
                  if (!DM->qZagruzka->Eof && ana!=ana1 &&firma==firma1)
                  {
                    rtf_Out("firma", DM->qZagruzka->FieldByName("firma")->AsString, 1);
                    rtf_Out("naim", "ПО ЗАЯВЛЕНИЮ", 1);

                    if(!rtf_LineFeed())
                      {
                        MessageBox(Handle,"Ошибка записи в файл данных","Ошибка",8192);
                        if (!rtf_Close()) MessageBox(Handle,"Ошибка закрытия файла данных","Ошибка",8192);
                        return;
                      }

                  }
              while (!DM->qZagruzka->Eof && zex==zex1)
                {

                  ana1 = DM->qZagruzka->FieldByName("ana")->AsInteger;
                  tn_pr1 = DM->qZagruzka->FieldByName("tn")->AsInteger;

                  rtf_Out("kust", DM->qZagruzka->FieldByName("kust")->AsString, 2);
                  rtf_Out("zex", DM->qZagruzka->FieldByName("zex")->AsString, 2);
                  rtf_Out("tn", DM->qZagruzka->FieldByName("tn")->AsString, 2);
                  rtf_Out("fio", DM->qZagruzka->FieldByName("fio")->AsString, 2);
                  rtf_Out("sum", DM->qZagruzka->FieldByName("sum")->AsFloat,20,2, 2);

                  //проверка превышения
                  Sql1 = "select (nvl(sum,0) - nvl( (select sum(sum)                                               \
                                                     from slst"+dtp_mm + dtp_year+"                                \
                                                     where zex="+ DM->qZagruzka->FieldByName("zex")->AsString +"   \
                                                     and tn="+ DM->qZagruzka->FieldByName("tn")->AsString +" and vo=576), 0))*0.15 as rl_sum,  \
                                  nvl((select sum(sum)                                                                                                  \
                                       from slst"+dtp_mm + dtp_year+"                                                                                    \
                                       where zex="+  DM->qZagruzka->FieldByName("zex")->AsString +" and tn="+ DM->qZagruzka->FieldByName("tn")->AsString +"   \
                                       and vo=859), 0) as obsh_sum                     \
                                  from slst"+dtp_mm + dtp_year+"                                                           \
                                  where typs=9 and zex="+ DM->qZagruzka->FieldByName("zex")->AsString +"    \
                                  and tn="+ DM->qZagruzka->FieldByName("tn")->AsString;

                  DM->qObnovlenie->Close();
                  DM->qObnovlenie->SQL->Clear();
                  DM->qObnovlenie->SQL->Add(Sql1);
                  try
                    {
                      DM->qObnovlenie->Open();
                    }
                  catch(...)
                    {
                      Application->MessageBox("Ошибка получения данных из таблицы SLST.","Ошибка",MB_OK);
                      StatusBar1->SimplePanel = false;
                      ProgressBar->Visible = false;
                      StatusBar1->SimpleText = "";
                      Main->Cursor = crDefault;

                      Abort();
                    }

                  rl_sum = DM->qObnovlenie->FieldByName("rl_sum")->AsFloat;
                  obsh_sum = DM->qObnovlenie->FieldByName("obsh_sum")->AsFloat;

                  if (rl_sum > prozhitMin || rl_sum == prozhitMin)
                    {
                      if (obsh_sum > prozhitMin)
                        {
                          sum_pr=FloatToStrF(obsh_sum - prozhitMin,ffFixed,20,2);
                        }
                      else
                        {
                          sum_pr = "";
                        }
                    }
                  else
                    {
                      if (obsh_sum > rl_sum)
                        {
                          sum_pr=FloatToStrF(obsh_sum - rl_sum,ffFixed,20,2);
                        }
                      else
                        {
                          sum_pr = "";
                        }

                    }

                  //вывод превышения
                  if (zex==zex1 && tn_pr==tn_pr1)
                    {
                      rtf_Out("prev"," ", 2);

                      if(!rtf_LineFeed())
                        {
                          MessageBox(Handle,"Ошибка записи в файл данных","Ошибка",8192);
                          if (!rtf_Close()) MessageBox(Handle,"Ошибка закрытия файла данных","Ошибка",8192);
                          return;
                        }
                    }
                  else
                    {
                      rtf_Out("prev",sum_pr, 2);

                      if(!rtf_LineFeed())
                        {
                          MessageBox(Handle,"Ошибка записи в файл данных","Ошибка",8192);
                          if (!rtf_Close()) MessageBox(Handle,"Ошибка закрытия файла данных","Ошибка",8192);
                          return;
                        }
                    }

                  sum_zex += DM->qZagruzka->FieldByName("sum")->AsFloat,20,2;
                  sum_kust = DM->qZagruzka->FieldByName("sum_po_kust")->AsFloat;

                  tn_pr = DM->qZagruzka->FieldByName("tn")->AsInteger;
                  DM->qZagruzka->Next();
                  ProgressBar->Position++;

                  zex = DM->qZagruzka->FieldByName("zex")->AsInteger;
                  firma = DM->qZagruzka->FieldByName("firma")->AsString;
                  ana = DM->qZagruzka->FieldByName("ana")->AsInteger;
                }

              //вывод суммы по цеху
              rtf_Out("sum_po_zex", FloatToStrF(sum_zex,ffFixed,20,2),3);

              if(!rtf_LineFeed())
                {
                  MessageBox(Handle,"Ошибка записи в файл данных","Ошибка",8192);
                  if (!rtf_Close()) MessageBox(Handle,"Ошибка закрытия файла данных","Ошибка",8192);
                  return;
                }

              sum_zex=0;
            }

          //вывод суммы по кусту
          rtf_Out("sum_po_kust", FloatToStrF(sum_kust, ffFixed,20,2),4);
          rtf_Out("firma2", firma1,4);

          if(!rtf_LineFeed())
            {
              MessageBox(Handle,"Ошибка записи в файл данных","Ошибка",8192);
              if (!rtf_Close()) MessageBox(Handle,"Ошибка закрытия файла данных","Ошибка",8192);
              return;
            }
        }
    /*
      // по комбинату
      rtf_Out("sum_po_kombinat",FloatToStrF(DM->qZagruzka->FieldByName("sum_po_kombinat")->AsFloat,ffFixed,20,2), 5);
      if(!rtf_LineFeed())
        {
          MessageBox(Handle,"Ошибка записи в файл данных","Ошибка",8192);
          if (!rtf_Close()) MessageBox(Handle,"Ошибка закрытия файла данных","Ошибка",8192);
          return;
        }
    */
       StatusBar1->SimpleText = " Идет формирование валютной страховки по РМЗ... ";

 //*****************************************************************************
    // Формирование валютной страховки по агрофирмам
   Sql = "select (select fam||' '||im||' ' ||ot from avans  where ncex=sl.zex and tn=sl.tn) as fio ,    \
                  sl.zex, sl.tn, sl.sum, sp.kust, nvl(sp.firma, 0) as firma, sp.ana,nist,                   \
                  sum(sum) over (partition by sl.zex) sum_po_zex,                                      \
                  sum(sum) over (partition by sp.firma)  sum_po_kust,                                   \
                  sum(sum) over() sum_po_kombinat                                                      \
          from slst"+dtp_mm + dtp_year+" sl, spnc sp                                                   \
          where vo = 859                                                                             \
          and nvl(nist,0) in (1,2)                                                                     \
          and sl.zex=sp.nc                                                                             \
          and sp.ana between 2 and 10 and ana not in (6,7)                                                                 \
          and nvl(sum,0)>0                                                                             \
          order by firma, kust, zex, tn,nist";

  DM->qZagruzka->Close();
  DM->qZagruzka->SQL->Clear();
  DM->qZagruzka->SQL->Add(Sql);
  try
    {
      DM->qZagruzka->Open();
    }
  catch(...)
    {
      Application->MessageBox("Ошибка получения данных из таблицы SLST. \n ВОЗМОЖНО неверно выбран период.","Ошибка",MB_OK);
      StatusBar1->SimplePanel = false;
      ProgressBar->Visible = false;
      StatusBar1->SimpleText = "";
      Main->Cursor = crDefault;

      Abort();
    }

      ProgressBar->Position = 0;
      ProgressBar->Max = DM->qZagruzka->RecordCount;

      ana=8888;
      zex = DM->qZagruzka->FieldByName("zex")->AsInteger;
      firma = DM->qZagruzka->FieldByName("firma")->AsString;


      while (!DM->qZagruzka->Eof)
        {

          firma1 = DM->qZagruzka->FieldByName("firma")->AsString;

          while (!DM->qZagruzka->Eof && firma==firma1)
            {

              zex1 = DM->qZagruzka->FieldByName("zex")->AsInteger;
                  if (!DM->qZagruzka->Eof && ana!=ana1 &&firma==firma1)
                  {
                    rtf_Out("firma", DM->qZagruzka->FieldByName("firma")->AsString, 1);
                    rtf_Out("naim", "В ВАЛЮТЕ", 1);

                    if(!rtf_LineFeed())
                      {
                        MessageBox(Handle,"Ошибка записи в файл данных","Ошибка",8192);
                        if (!rtf_Close()) MessageBox(Handle,"Ошибка закрытия файла данных","Ошибка",8192);
                        return;
                      }

                  }
              while (!DM->qZagruzka->Eof && zex==zex1)
                {

                  ana1 = DM->qZagruzka->FieldByName("ana")->AsInteger;
                  tn_pr1 = DM->qZagruzka->FieldByName("tn")->AsInteger;

                  rtf_Out("kust", DM->qZagruzka->FieldByName("kust")->AsString, 2);
                  rtf_Out("zex", DM->qZagruzka->FieldByName("zex")->AsString, 2);
                  rtf_Out("tn", DM->qZagruzka->FieldByName("tn")->AsString, 2);
                  rtf_Out("fio", DM->qZagruzka->FieldByName("fio")->AsString, 2);
                  rtf_Out("sum", DM->qZagruzka->FieldByName("sum")->AsFloat,20,2, 2);

                  //проверка превышения
                  Sql1 = "select (nvl(sum,0) - nvl( (select sum(sum)                                               \
                                                     from slst"+dtp_mm + dtp_year+"                                \
                                                     where zex="+ DM->qZagruzka->FieldByName("zex")->AsString +"   \
                                                     and tn="+ DM->qZagruzka->FieldByName("tn")->AsString +" and vo=576), 0))*0.15 as rl_sum,  \
                                  nvl((select sum(sum)                                                                                                  \
                                       from slst"+dtp_mm + dtp_year+"                                                                                    \
                                       where zex="+  DM->qZagruzka->FieldByName("zex")->AsString +" and tn="+ DM->qZagruzka->FieldByName("tn")->AsString +"   \
                                       and vo=859), 0) as obsh_sum                     \
                                  from slst"+dtp_mm + dtp_year+"                                                          \
                                  where typs=9 and zex="+ DM->qZagruzka->FieldByName("zex")->AsString +"    \
                                  and tn="+ DM->qZagruzka->FieldByName("tn")->AsString;

                  DM->qObnovlenie->Close();
                  DM->qObnovlenie->SQL->Clear();
                  DM->qObnovlenie->SQL->Add(Sql1);
                  try
                    {
                      DM->qObnovlenie->Open();
                    }
                  catch(...)
                    {
                      Application->MessageBox("Ошибка получения данных из таблицы SLST.","Ошибка",MB_OK);
                      StatusBar1->SimplePanel = false;
                      ProgressBar->Visible = false;
                      StatusBar1->SimpleText = "";
                      Main->Cursor = crDefault;

                      Abort();
                    }

                  rl_sum = DM->qObnovlenie->FieldByName("rl_sum")->AsFloat;
                  obsh_sum = DM->qObnovlenie->FieldByName("obsh_sum")->AsFloat;

                  if (rl_sum > prozhitMin || rl_sum == prozhitMin)
                    {
                      if (obsh_sum > prozhitMin)
                        {
                          sum_pr=FloatToStrF(obsh_sum - prozhitMin,ffFixed,20,2);
                        }
                      else
                        {
                          sum_pr = "";
                        }
                    }
                  else
                    {
                      if (obsh_sum > rl_sum)
                        {
                          sum_pr=FloatToStrF(obsh_sum - rl_sum,ffFixed,20,2);
                        }
                      else
                        {
                          sum_pr = "";
                        }

                    }

                  //вывод превышения
                  if (zex==zex1 && tn_pr==tn_pr1)
                    {
                      rtf_Out("prev"," ", 2);

                      if(!rtf_LineFeed())
                        {
                          MessageBox(Handle,"Ошибка записи в файл данных","Ошибка",8192);
                          if (!rtf_Close()) MessageBox(Handle,"Ошибка закрытия файла данных","Ошибка",8192);
                          return;
                        }
                    }
                  else
                    {
                      rtf_Out("prev",sum_pr, 2);

                      if(!rtf_LineFeed())
                        {
                          MessageBox(Handle,"Ошибка записи в файл данных","Ошибка",8192);
                          if (!rtf_Close()) MessageBox(Handle,"Ошибка закрытия файла данных","Ошибка",8192);
                          return;
                        }
                    }

                  sum_zex += DM->qZagruzka->FieldByName("sum")->AsFloat,20,2;
                  sum_kust = DM->qZagruzka->FieldByName("sum_po_kust")->AsFloat;

                  tn_pr = DM->qZagruzka->FieldByName("tn")->AsInteger;
                  DM->qZagruzka->Next();
                  ProgressBar->Position++;

                  zex = DM->qZagruzka->FieldByName("zex")->AsInteger;
                  firma = DM->qZagruzka->FieldByName("firma")->AsString;
                  ana = DM->qZagruzka->FieldByName("ana")->AsInteger;
                }

              //вывод суммы по цеху
              rtf_Out("sum_po_zex", FloatToStrF(sum_zex,ffFixed,20,2),3);

              if(!rtf_LineFeed())
                {
                  MessageBox(Handle,"Ошибка записи в файл данных","Ошибка",8192);
                  if (!rtf_Close()) MessageBox(Handle,"Ошибка закрытия файла данных","Ошибка",8192);
                  return;
                }

              sum_zex=0;
            }

          //вывод суммы по кусту
          rtf_Out("sum_po_kust", FloatToStrF(sum_kust, ffFixed,20,2),4);
          rtf_Out("firma2", firma1,4);

          if(!rtf_LineFeed())
            {
              MessageBox(Handle,"Ошибка записи в файл данных","Ошибка",8192);
              if (!rtf_Close()) MessageBox(Handle,"Ошибка закрытия файла данных","Ошибка",8192);
              return;
            }
        }
   /*
      // по комбинату
      rtf_Out("sum_po_kombinat",FloatToStrF(DM->qZagruzka->FieldByName("sum_po_kombinat")->AsFloat,ffFixed,20,2), 5);
      if(!rtf_LineFeed())
        {
          MessageBox(Handle,"Ошибка записи в файл данных","Ошибка",8192);
          if (!rtf_Close()) MessageBox(Handle,"Ошибка закрытия файла данных","Ошибка",8192);
          return;
        }
            */
       StatusBar1->SimpleText = " Идет формирование внешних договоров по РМЗ...";

 //*****************************************************************************
   // Формирование внешних договоров по агрофирмам
   Sql = "select (select fam||' '||im||' ' ||ot from avans  where ncex=sl.zex and tn=sl.tn) as fio ,    \
                  sl.zex, sl.tn, sl.sum, sp.kust, nvl(sp.firma, 0) as firma, sp.ana, nist,                                   \
                  sum(sum) over (partition by sl.zex) sum_po_zex,                                      \
                  sum(sum) over (partition by sp.firma)  sum_po_kust,                                   \
                  sum(sum) over() sum_po_kombinat                                                      \
          from slst"+dtp_mm + dtp_year+" sl, spnc sp                                                                  \
          where vo = 859                                                                               \
          and nvl(nist,0)=3                                                                            \
          and sl.zex=sp.nc                                                                             \
          and sp.ana between 2 and 10 and ana not in (6,7)                                                              \
          and nvl(sum,0)>0                                                                             \                                           \
          order by firma, kust, zex, tn,nist";
  DM->qZagruzka->Close();
  DM->qZagruzka->SQL->Clear();
  DM->qZagruzka->SQL->Add(Sql);
  try
    {
      DM->qZagruzka->Open();
    }
  catch(...)
    {
      Application->MessageBox("Ошибка получения данных из таблицы SLST. \n ВОЗМОЖНО неверно выбран период.","Ошибка",MB_OK);
      StatusBar1->SimplePanel = false;
      ProgressBar->Visible = false;
      StatusBar1->SimpleText = "";
      Main->Cursor = crDefault;

      Abort();
    }

      ProgressBar->Position = 0;
      ProgressBar->Max = DM->qZagruzka->RecordCount;
    
      ana=8888;
      zex = DM->qZagruzka->FieldByName("zex")->AsInteger;
      firma = DM->qZagruzka->FieldByName("firma")->AsString;


      while (!DM->qZagruzka->Eof)
        {

          firma1 = DM->qZagruzka->FieldByName("firma")->AsString;
       
          while (!DM->qZagruzka->Eof && firma==firma1)
            {

              zex1 = DM->qZagruzka->FieldByName("zex")->AsInteger;
                  if (!DM->qZagruzka->Eof && ana!=ana1 && firma==firma1)
                  {
                    rtf_Out("firma", DM->qZagruzka->FieldByName("firma")->AsString, 1);
                    rtf_Out("naim", "ПО ВНЕШНИМ ДОГОВОРАМ", 1);

                    if(!rtf_LineFeed())
                      {
                        MessageBox(Handle,"Ошибка записи в файл данных","Ошибка",8192);
                        if (!rtf_Close()) MessageBox(Handle,"Ошибка закрытия файла данных","Ошибка",8192);
                        return;
                      }

                  }
              while (!DM->qZagruzka->Eof && zex==zex1)
                {

                  ana1 = DM->qZagruzka->FieldByName("ana")->AsInteger;
                  tn_pr1 = DM->qZagruzka->FieldByName("tn")->AsInteger;

                  rtf_Out("kust", DM->qZagruzka->FieldByName("kust")->AsString, 2);
                  rtf_Out("zex", DM->qZagruzka->FieldByName("zex")->AsString, 2);
                  rtf_Out("tn", DM->qZagruzka->FieldByName("tn")->AsString, 2);
                  rtf_Out("fio", DM->qZagruzka->FieldByName("fio")->AsString, 2);
                  rtf_Out("sum", DM->qZagruzka->FieldByName("sum")->AsFloat,20,2, 2);

                  //проверка превышения
                  Sql1 = "select (nvl(sum,0) - nvl( (select sum(sum)                                               \
                                                     from slst"+dtp_mm + dtp_year+"                                \
                                                     where zex="+ DM->qZagruzka->FieldByName("zex")->AsString +"   \
                                                     and tn="+ DM->qZagruzka->FieldByName("tn")->AsString +" and vo=576), 0))*0.15 as rl_sum,  \
                                  nvl((select sum(sum)                                                                                                  \
                                       from slst"+dtp_mm + dtp_year+"                                                                                    \
                                       where zex="+  DM->qZagruzka->FieldByName("zex")->AsString +" and tn="+ DM->qZagruzka->FieldByName("tn")->AsString +"   \
                                       and vo=859), 0) as obsh_sum                     \
                                  from slst"+dtp_mm + dtp_year+"                                                          \
                                  where typs=9 and zex="+ DM->qZagruzka->FieldByName("zex")->AsString +"    \
                                  and tn="+ DM->qZagruzka->FieldByName("tn")->AsString;

                  DM->qObnovlenie->Close();
                  DM->qObnovlenie->SQL->Clear();
                  DM->qObnovlenie->SQL->Add(Sql1);
                  try
                    {
                      DM->qObnovlenie->Open();
                    }
                  catch(...)
                    {
                      Application->MessageBox("Ошибка получения данных из таблицы SLST.","Ошибка",MB_OK);
                      StatusBar1->SimplePanel = false;
                      ProgressBar->Visible = false;
                      StatusBar1->SimpleText = "";
                      Main->Cursor = crDefault;

                      Abort();
                    }

                  rl_sum = DM->qObnovlenie->FieldByName("rl_sum")->AsFloat;
                  obsh_sum = DM->qObnovlenie->FieldByName("obsh_sum")->AsFloat;

                  if (rl_sum > prozhitMin || rl_sum == prozhitMin)
                    {
                      if (obsh_sum > prozhitMin)
                        {
                          sum_pr=FloatToStrF(obsh_sum - prozhitMin,ffFixed,20,2);
                        }
                      else
                        {
                          sum_pr = "";
                        }
                    }
                  else
                    {
                      if (obsh_sum > rl_sum)
                        {
                          sum_pr=FloatToStrF(obsh_sum - rl_sum,ffFixed,20,2);
                        }
                      else
                        {
                          sum_pr = "";
                        }

                    }

                  //вывод превышения
                  if (zex==zex1 && tn_pr==tn_pr1)
                    {
                      rtf_Out("prev"," ", 2);

                      if(!rtf_LineFeed())
                        {
                          MessageBox(Handle,"Ошибка записи в файл данных","Ошибка",8192);
                          if (!rtf_Close()) MessageBox(Handle,"Ошибка закрытия файла данных","Ошибка",8192);
                          return;
                        }
                    }
                  else
                    {
                      rtf_Out("prev",sum_pr, 2);

                      if(!rtf_LineFeed())
                        {
                          MessageBox(Handle,"Ошибка записи в файл данных","Ошибка",8192);
                          if (!rtf_Close()) MessageBox(Handle,"Ошибка закрытия файла данных","Ошибка",8192);
                          return;
                        }
                    }

                  sum_zex += DM->qZagruzka->FieldByName("sum")->AsFloat,20,2;
                  sum_kust = DM->qZagruzka->FieldByName("sum_po_kust")->AsFloat;

                  tn_pr = DM->qZagruzka->FieldByName("tn")->AsInteger;
                  DM->qZagruzka->Next();
                  ProgressBar->Position++;

                  zex = DM->qZagruzka->FieldByName("zex")->AsInteger;
                  firma = DM->qZagruzka->FieldByName("firma")->AsString;
                  ana = DM->qZagruzka->FieldByName("ana")->AsInteger;
                }

              //вывод суммы по цеху
              rtf_Out("sum_po_zex", FloatToStrF(sum_zex,ffFixed,20,2),3);

              if(!rtf_LineFeed())
                {
                  MessageBox(Handle,"Ошибка записи в файл данных","Ошибка",8192);
                  if (!rtf_Close()) MessageBox(Handle,"Ошибка закрытия файла данных","Ошибка",8192);
                  return;
                }

              sum_zex=0;
            }

          //вывод суммы по кусту
          rtf_Out("sum_po_kust", FloatToStrF(sum_kust, ffFixed,20,2),4);
          rtf_Out("firma2", firma1,4);

          if(!rtf_LineFeed())
            {
              MessageBox(Handle,"Ошибка записи в файл данных","Ошибка",8192);
              if (!rtf_Close()) MessageBox(Handle,"Ошибка закрытия файла данных","Ошибка",8192);
              return;
            }
        }
   /*
      // по комбинату
      rtf_Out("sum_po_kombinat",FloatToStrF(DM->qZagruzka->FieldByName("sum_po_kombinat")->AsFloat,ffFixed,20,2), 5);
      if(!rtf_LineFeed())
        {
          MessageBox(Handle,"Ошибка записи в файл данных","Ошибка",8192);
          if (!rtf_Close()) MessageBox(Handle,"Ошибка закрытия файла данных","Ошибка",8192);
          return;
        }
         */
  StatusBar1->SimplePanel = false;
  ProgressBar->Visible = false;
  StatusBar1->SimpleText = "";
  Main->Cursor = crDefault;

  if(!rtf_Close())
    {
      MessageBox(Handle,"Ошибка закрытия файла данных", "Ошибка", 8192);
      return;
    }

  //Создание папки, если ее не существует
  ForceDirectories(WorkPath);

  int istrd;
  try
    {
      rtf_CreateReport(TempPath +"\\dlya_agro.txt", Path+"\\RTF\\dlya_agro.rtf",
                       WorkPath+"\\Для Агрофирм.doc",NULL,&istrd);


      WinExec(("\""+ WordPath+"\"\""+WorkPath+"\\Для Агрофирм.doc\"").c_str(),SW_MAXIMIZE);
    }
  catch(RepoRTF_Error E)
    {
      MessageBox(Handle,("Ошибка формирования отчета:"+ AnsiString(E.Err)+
                         "\nСтрока файла данных:"+IntToStr(istrd)).c_str(),"Ошибка",8192);
    }

    }
   
   StatusBar1->SimpleText = " ";
}
//---------------------------------------------------------------------------
//Формирование файла для ОК
void __fastcall TMain::N14Click(TObject *Sender)
{
  AnsiString Sql;

  Data->ShowModal();
  if (Data->ModalResult == mrCancel) {Abort();}

  //Считывание данных из DateTimePicker
  DecodeDate(Data->DateTimePicker1->Date, dtp_year, dtp_month, dtp_day );

  if (StrToInt(dtp_month)<10)
        {
          dtp_mm ="0"+ IntToStr(dtp_month);
        }
      else
        {
          dtp_mm = IntToStr(dtp_month);
        }

  Sql = "select (select fam||' '||im||' ' ||ot from avans  where ncex=sl.zex and tn=sl.tn) as fio,\
                (select fam from avans where ncex=sl.zex and tn=sl.tn) as fam,\
                 sl.zex , sl.tn, sl.sum as sum, sp.kust, sp.firma, decode(nvl(sl.nist,0),0,'ГРН',1,'ДОЛ',2,'ЕВРО') as nist,             \
                 sum(sum) over (partition by sl.zex) sum_po_zex,     \
                 sum(sum) over (partition by sp.kust) sum_po_kust,     \
                 sum(sum) over() sum_po_kombinat                        \
          from slst"+ dtp_mm + dtp_year+" sl, spnc sp            \
          where vo = 859                                         \
          and nvl(nist,0) in (0,1,2)                                      \
          and sl.zex=sp.nc                                       \
          and sp.ana in (1,6)                                         \
          and nvl(sum,0)>0                                         \
          order by nist,ana, kust,  zex, tn,sl.sum";


      //    (select n_dogovora from vu_859_n where zex=sl.zex and tn=sl.tn) as mf,

  DM->qZagruzka->Close();
  DM->qZagruzka->SQL->Clear();
  DM->qZagruzka->SQL->Add(Sql);
  try
    {
      DM->qZagruzka->Open();
    }
  catch(...)
    {
      Application->MessageBox("Ошибка получения данных из таблицы SLST. \n ВОЗМОЖНО неверно выбран период.","Ошибка",MB_OK);
      Abort();
    }

   // Количество записей
  int row = DM->qZagruzka->RecordCount;

  // устанавливаем путь к файлу шаблона
  AnsiString sFile = Path+"\\RTF\\dlya_ok.xlt";


   // инициализируем Excel, открываем этот шаблон
  try
    {
      AppEx=GetActiveOleObject("Excel.Application");
    }
  catch(...)
    {
     //проверяем, нет ли запущенного Excel
      try
        {
          AppEx=CreateOleObject("Excel.Application");
        }
      catch (...)
        {
          Application->MessageBox("Невозможно открыть Microsoft Excel!"
              " Возможно это приложение на компьютере не установлено.","Ошибка",MB_OK+MB_ICONERROR);
        }
    }

  try
    {
      AppEx.OlePropertyGet("WorkBooks").OleProcedure("Open",sFile.c_str());     //открываем книгу, указав её имя
      Sh=AppEx.OlePropertyGet("WorkSheets",1);                                  //выбираем № активного листа книги
    }
  catch(...)
    {
      Application->MessageBox("Ошибка открытия книги Microsoft Excel!","Ошибка",MB_OK+MB_ICONERROR);
    }

  StatusBar1->SimpleText = " Идет формирование страховки по пром. цехам и управлению с\х... ";

  Main->Cursor = crHourGlass;
  ProgressBar->Visible = true;
  ProgressBar->Position = 0;
  ProgressBar->Max = DM->qZagruzka->RecordCount;

   // Вставляем в шаблон нужное количество строк

  Variant C;
  AppEx.OlePropertyGet("WorkSheets",1).OleProcedure("Select");
  C=AppEx.OlePropertyGet("Range","zex");
  C=AppEx.OlePropertyGet("Rows",(int) C.OlePropertyGet("Row")+1);
  for(int i=1;i<row;i++) C.OleProcedure("Insert");

  int i=0;

  while (!DM->qZagruzka->Eof)
    {
      toExcel(AppEx,"zex",i,i+1);
      toExcel(AppEx,"zex",i, DM->qZagruzka->FieldByName("zex")->AsString.c_str());
      toExcel(AppEx,"tn",i, DM->qZagruzka->FieldByName("tn")->AsString.c_str());
      toExcel(AppEx,"fio",i, DM->qZagruzka->FieldByName("fio")->AsString.c_str());
      toExcel(AppEx,"sum",i, DM->qZagruzka->FieldByName("sum")->AsFloat);
      toExcel(AppEx,"inn",i, DM->qZagruzka->FieldByName("fam")->AsString.c_str());
 //    toExcel(AppEx,"mf", DM->qZagruzka->FieldByName("mf")->AsString.c_str());
      toExcel(AppEx,"nist",i, DM->qZagruzka->FieldByName("nist")->AsString.c_str());
      i++;

      DM->qZagruzka->Next();
      ProgressBar->Position++;
       
    }

   //Отключить вывод сообщений с вопросами типа "Заменить файл"..."
   AppEx.OlePropertySet("DisplayAlerts",false);

   //Создание папки если ее не существуют
   ForceDirectories(WorkPath);

   //Сохранить книгу в папке в файле по указанию
   AnsiString vAsCurDir1=WorkPath+"\\Для ОК.xls";
   AppEx.OlePropertyGet("Workbooks").OlePropertyGet("Item",1).
   OleProcedure("SaveAs",vAsCurDir1.c_str());

   //Закрыть Excel
   AppEx.OleProcedure("Quit");
   AppEx.OlePropertySet("Visible",true);

   StatusBar1->SimplePanel = false;
   ProgressBar->Visible = false;
   StatusBar1->SimpleText = "";
   Main->Cursor = crDefault;

   AppEx.OlePropertyGet("WorkBooks").OleProcedure("Open",vAsCurDir1.c_str());
  //AppEx.OlePropertySet("DisplayAlerts",true);
}
//---------------------------------------------------------------------------

void __fastcall TMain::N16Click(TObject *Sender)
{
AnsiString sql;

  //Формирование файла по гривневой страховке Пром+Агро
  FILE *grn;


  Data->ShowModal();
  if (Data->ModalResult == mrCancel) {Abort();}

  //Считывание данных из DateTimePicker
  DecodeDate(Data->DateTimePicker1->Date, dtp_year, dtp_month, dtp_day );

  //Создание папки если ее не существуют
  ForceDirectories(WorkPath);

  if ((grn=fopen((WorkPath+"\\fkiev.txt").c_str(),"wt"))==NULL)
    {
      ShowMessage("Файл не удается открыть");
      return;
    }

   StatusBar1->SimpleText = "Формирование файла по гривневой страховке Пром+Агро для страховой компании... ";

  if (StrToInt(dtp_month)<10)
        {
          dtp_mm ="0"+ IntToStr(dtp_month);
        }
      else
        {
          dtp_mm = IntToStr(dtp_month);
        }

   //запрос
  sql = "select * from (select (select fio from slst"+ dtp_mm + dtp_year+" where zex=sl.zex and tn=sl.tn and typs=9) as fio ,   \
                        sl.zex, sl.tn, sl.sum, sp.kust, sp.ana,                                                     \
                        (select vnvi from slst"+ dtp_mm + dtp_year+" where zex=sl.zex and tn=sl.tn and typs=9) as vnvi          \                                              \
               from slst"+ dtp_mm + dtp_year+" sl, spnc sp                                                                          \
               where vo = 859                                                                                      \
               and nvl(nist,0) =0                                                                                  \
               and sp.ana=1                                                                                        \
               and sl.zex=sp.nc                                                                                    \
               and nvl(sum,0)>0                                                                                    \
          union all                                                                                                   \
             (select (select fio from slst"+ dtp_mm + dtp_year+" where zex=sl.zex and tn=sl.tn and typs=9) as fio ,           \
                        sl.zex, sl.tn, sum(sl.sum) as sum, sp.kust, sp.ana,                                                    \
                          (select vnvi from slst"+ dtp_mm + dtp_year+" where zex=sl.zex and tn=sl.tn and typs=9) as vnvi         \
              from slst"+ dtp_mm + dtp_year+" sl, spnc sp                                                                          \
              where vo = 859                                                                                       \
              and nvl(nist,0)=0                                                                                    \
              and sl.zex=sp.nc                                                                                     \
              and sp.ana between 2 and 9                                                                           \
              and nvl(sum,0)>0                                                                                     \
              group by sp.ana, sp.kust, sl.zex, sl.tn)                                                     \
          )                                                                                                        \
          order by ana, kust, zex, tn";

                               
  DM->qObnovlenie->Close();
  DM->qObnovlenie->SQL->Clear();
  DM->qObnovlenie->SQL->Add(sql);
  try
    {
      DM->qObnovlenie->Open();
    }
  catch(...)
    {
      Application->MessageBox("Ошибка получения данных из таблицы SLST. \n ВОЗМОЖНО неверно выбран период.","Ошибка",MB_OK);
      StatusBar1->SimplePanel = false;
      ProgressBar->Visible = false;
      StatusBar1->SimpleText = "";
      Main->Cursor = crDefault;
      Abort();
    }

  Main->Cursor = crHourGlass;
  ProgressBar->Visible = true;
  ProgressBar->Position = 0;
  ProgressBar->Max = DM->qObnovlenie->RecordCount;

  //вывод в файл
  while (!DM->qObnovlenie->Eof)
    {
      fprintf(grn,"\n%3s|%4s|%-37s|%8s|%10s|",DM->qObnovlenie->FieldByName("zex")->AsString,
                                              DM->qObnovlenie->FieldByName("tn")->AsString,
                                              DM->qObnovlenie->FieldByName("fio")->AsString,
                                              FloatToStrF(DM->qObnovlenie->FieldByName("sum")->AsFloat, ffFixed, 20,2),
                                              DM->qObnovlenie->FieldByName("vnvi")->AsString);

      DM->qObnovlenie->Next();
      ProgressBar->Position++;
    }
  fclose(grn);


  StatusBar1->SimpleText = "Формирование файла по валютной страховке и внешним договорам Пром+Агро...";

  //Формирование файла по валютной страховке и внешним договорам Пром+Агро
  FILE *val;
  if ((val=fopen((WorkPath+"\\fkievv.txt").c_str(),"wt"))==NULL)
    {
      ShowMessage("Файл не удается открыть");
      return;
    }
        // валюта по пром.
       //запрос
      sql=" select (select fio from slst"+ dtp_mm + dtp_year+" where zex=sl.zex and tn=sl.tn and typs=9) as fio ,             \
                                    sl.zex, sl.tn, sl.sum, sp.kust, sp.ana,                                              \
                                   (select vnvi from slst"+ dtp_mm + dtp_year+" where zex=sl.zex and tn=sl.tn and typs=9) as vnvi,nist   \
                            from slst"+ dtp_mm + dtp_year+" sl, spnc sp                                                                  \
                            where vo = 859                                                                               \
                            and nvl(nist,0) in (1,2)                                                                     \
                            and sp.ana=1                                                                                 \
                            and sl.zex=sp.nc                                                                             \
                            and nvl(sum,0)>0                                                                             \
                            order by  kust,  zex,   tn ";

    /*
  sql = "select * from (select (select fio from slst"+ dtp_mm + dtp_year+" where zex=sl.zex and tn=sl.tn and typs=9) as fio ,   \
                        sl.zex, sl.tn, sl.sum, sp.kust, sp.ana,                                                       \
                        (select vnvi from slst"+ dtp_mm + dtp_year+" where zex=sl.zex and tn=sl.tn and typs=9) as vnvi,nist         \
               from slst"+ dtp_mm + dtp_year+" sl, spnc sp                                                                          \
               where vo = 859                                                                                      \
               and nvl(nist,0) in (1,2,3)                                                                                 \
               and sp.ana=1                                                                                        \
               and sl.zex=sp.nc                                                                                    \
               and nvl(sum,0)>0                                                                                    \
          union all                                                                                                   \
             (select (select fio from slst"+ dtp_mm + dtp_year+" where zex=sl.zex and tn=sl.tn and typs=9) as fio ,           \
                        sl.zex, sl.tn, sl.sum, sp.kust, sp.ana,                                                    \
                        (select vnvi from slst"+ dtp_mm + dtp_year+" where zex=sl.zex and tn=sl.tn and typs=9) as vnvi,nist         \
              from slst"+ dtp_mm + dtp_year+" sl, spnc sp                                                                          \
              where vo = 859                                                                                       \
              and nvl(nist,0) in (1,2,3)                                                                                    \
              and sl.zex=sp.nc                                                                                     \
              and sp.ana between 2 and 9                                                                           \
              and nvl(sum,0)>0                                                                                     \
             )                                                     \
          )                                                                                                        \
          order by  ana, kust, zex, tn, nist ";     */

  DM->qObnovlenie->Close();
  DM->qObnovlenie->SQL->Clear();
  DM->qObnovlenie->SQL->Add(sql);
  try
    {
      DM->qObnovlenie->Open();
    }
  catch(...)
    {
      Application->MessageBox("Ошибка получения данных из таблицы SLST. \n ВОЗМОЖНО неверно выбран период.","Ошибка",MB_OK);
      StatusBar1->SimplePanel = false;
      ProgressBar->Visible = false;
      StatusBar1->SimpleText = "";
      Main->Cursor = crDefault;
    }

  ProgressBar->Position = 0;
  ProgressBar->Max = DM->qObnovlenie->RecordCount;

  //вывод в файл
  while (!DM->qObnovlenie->Eof)
    {
      fprintf(grn,"\n%3s|%4s|%-37s|%8s|%10s|",DM->qObnovlenie->FieldByName("zex")->AsString,
                                              DM->qObnovlenie->FieldByName("tn")->AsString,
                                              DM->qObnovlenie->FieldByName("fio")->AsString,
                                              FloatToStrF(DM->qObnovlenie->FieldByName("sum")->AsFloat, ffFixed,20,2),
                                              DM->qObnovlenie->FieldByName("vnvi")->AsString);

      DM->qObnovlenie->Next();
      ProgressBar->Position++;
    }

    // внешние пром
   sql =" select (select fio from slst"+ dtp_mm + dtp_year+" where zex=sl.zex and tn=sl.tn and typs=9) as fio ,                    \
                                    sl.zex, sl.tn, sl.sum, sp.kust, sp.ana,                                              \
                                   (select vnvi from slst"+ dtp_mm + dtp_year+" where zex=sl.zex and tn=sl.tn and typs=9) as vnvi,nist   \
                            from slst"+ dtp_mm + dtp_year+" sl, spnc sp                                                                  \
                            where vo = 859                                                                               \
                            and nvl(nist,0) in (3)                                                                       \
                            and sp.ana=1                                                                                 \
                            and sl.zex=sp.nc                                                                             \
                            and nvl(sum,0)>0                                                                             \
                            order by  kust,  zex,   tn ";

  DM->qObnovlenie->Close();
  DM->qObnovlenie->SQL->Clear();
  DM->qObnovlenie->SQL->Add(sql);
  try
    {
      DM->qObnovlenie->Open();
    }
  catch(...)
    {
      Application->MessageBox("Ошибка получения данных из таблицы SLST. \n ВОЗМОЖНО неверно выбран период.","Ошибка",MB_OK);
      StatusBar1->SimplePanel = false;
      ProgressBar->Visible = false;
      StatusBar1->SimpleText = "";
      Main->Cursor = crDefault;
    }

  ProgressBar->Position = 0;
  ProgressBar->Max = DM->qObnovlenie->RecordCount;

  //вывод в файл
  while (!DM->qObnovlenie->Eof)
    {
      fprintf(grn,"\n%3s|%4s|%-37s|%8s|%10s|",DM->qObnovlenie->FieldByName("zex")->AsString,
                                              DM->qObnovlenie->FieldByName("tn")->AsString,
                                              DM->qObnovlenie->FieldByName("fio")->AsString,
                                              FloatToStrF(DM->qObnovlenie->FieldByName("sum")->AsFloat, ffFixed,20,2),
                                              DM->qObnovlenie->FieldByName("vnvi")->AsString);

      DM->qObnovlenie->Next();
      ProgressBar->Position++;
    }

  // агро валюта+пром
  sql="select (select fio from slst"+ dtp_mm + dtp_year+" where zex=sl.zex and tn=sl.tn and typs=9) as fio ,  \
                                    sl.zex, sl.tn, sl.sum, sp.kust, sp.ana,                                                       \
                                   (select vnvi from slst"+ dtp_mm + dtp_year+" where zex=sl.zex and tn=sl.tn and typs=9) as vnvi,nist            \
                            from slst"+ dtp_mm + dtp_year+" sl, spnc sp                                                                           \
                            where vo = 859                                                                                        \
                            and nvl(nist,0) in (1,2,3)                                                                            \
                            and sp.ana between 2 and 9                                                                            \
                            and sl.zex=sp.nc                                                                                      \
                            and nvl(sum,0)>0                                                                                      \
                            order by ana,  kust,  zex, tn ";

  DM->qObnovlenie->Close();
  DM->qObnovlenie->SQL->Clear();
  DM->qObnovlenie->SQL->Add(sql);
  try
    {
      DM->qObnovlenie->Open();
    }
  catch(...)
    {
      Application->MessageBox("Ошибка получения данных из таблицы SLST. \n ВОЗМОЖНО неверно выбран период.","Ошибка",MB_OK);
      StatusBar1->SimplePanel = false;
      ProgressBar->Visible = false;
      StatusBar1->SimpleText = "";
      Main->Cursor = crDefault;
      Abort();
    }

  ProgressBar->Position = 0;
  ProgressBar->Max = DM->qObnovlenie->RecordCount;

  //вывод в файл
  while (!DM->qObnovlenie->Eof)
    {
      fprintf(grn,"\n%3s|%4s|%-37s|%8s|%10s|",DM->qObnovlenie->FieldByName("zex")->AsString,
                                              DM->qObnovlenie->FieldByName("tn")->AsString,
                                              DM->qObnovlenie->FieldByName("fio")->AsString,
                                              FloatToStrF(DM->qObnovlenie->FieldByName("sum")->AsFloat, ffFixed,20,2),
                                              DM->qObnovlenie->FieldByName("vnvi")->AsString);

      DM->qObnovlenie->Next();
      ProgressBar->Position++;
    }


  fclose(val);

  StatusBar1->SimplePanel = false;
  ProgressBar->Visible = false;
  StatusBar1->SimpleText = "";
  Main->Cursor = crDefault;
  ShowMessage("Формирование файлов успешно завершено");
}
//---------------------------------------------------------------------------

void __fastcall TMain::EventsMessage(tagMSG &Msg, bool &Handled)
{
 if(Msg.message == WM_MOUSEWHEEL){
    Msg.message = WM_KEYDOWN;
    Msg.lParam = 0;
    short int i = HIWORD(Msg.wParam);
    Msg.wParam =(i > 0)?VK_UP:VK_DOWN;
    Handled = false;
  }
}
//---------------------------------------------------------------------------

// Проверка на существование инн для изменений по договорам
 void __fastcall TMain::ProverkaInfoExcelIzmeneniya()
{
  AnsiString Sql, Sql1, Sql2, tn, fam, n_doc, nn, inn, nnom, name, data_po;
  int fl=0, pr_inn=0, pr_dv=0, pr_sum=0;
  int i=1;
  double sum;
  TSearchRec SearchRecord;  //Для поиска файла


  /*tn - таб.№,
    Row - общее количество занятых строк в документе
    sum - сумма удержаний на страховку
    fl - признак формирования отчета (fl=1 - формировать)
    name - имя Excel файла
    FileName - полный путь к файлу с его именем
    Dir2 - путь к выбранной папке
    pr_sum - признак вывода шапки таблицы и заголовка по превыш. суммы (pr_sum = 0 выводить)
    pr_inn - признак вывода шапки таблицы и заголовка по отсут. индификац.№ (pr_inn = 0 выводить)
    pr_dv - признак вывода шапки таблицы и заголовка по двойным записям (pr_dv = 0 выводить)*/


   //Окно выбора директории к папке
  if (!SelectDirectory("Select directory",WideString(""),Dir2))
    {
      Abort();
    }

   //Поиск файла Excel
   switch(im_fl)
     {
       case 2 :  if (FindFirst(Dir2 + LowerCase("\\Изменения(гривна).xls"), faAnyFile, SearchRecord)==0 )
                   {
                     name = LowerCase("\\Изменения(гривна).xls");
                   }
                 else if (FindFirst (Dir2 + LowerCase("\\Изменения(гривна).xlsx"), faAnyFile, SearchRecord)==0)
                   {
                     name = LowerCase("\\Изменения(гривна).xlsx");
                   }
                 else
                   {
                     Application->MessageBox("Не найден файл для загрузки данных. \nВозможно указано НЕВЕРНОЕ ИМЯ файла \nили файл не найден в данной папке. ",
                                           "Ошибка загрузки данных", MB_OK + MB_ICONERROR);
                     Abort();
                   }
       break;
       case 3 :  if (FindFirst(Dir2 + LowerCase("\\Изменения(валюта).xls"), faAnyFile, SearchRecord)==0 )
                   {
                     name = LowerCase("\\Изменения(валюта).xls");
                   }
                 else if (FindFirst (Dir2 + LowerCase("\\Изменения(валюта).xlsx"), faAnyFile, SearchRecord)==0)
                   {
                     name = LowerCase("\\Изменения(валюта).xlsx");
                   }
                 else
                   {
                     Application->MessageBox("Не найден файл для загрузки данных. \nВозможно указано НЕВЕРНОЕ ИМЯ файла \nили файл не найден в данной папке. ",
                                            "Ошибка загрузки данных", MB_OK + MB_ICONERROR);
                     Abort();
                   }
       break;
       case 4 :  if (FindFirst(Dir2 + LowerCase("\\Изменения(курс).xls"), faAnyFile, SearchRecord)==0 )
                   {
                     name = LowerCase("\\Изменения(курс).xls");
                   }
                 else if (FindFirst (Dir2 + LowerCase("\\Изменения(курс).xlsx"), faAnyFile, SearchRecord)==0)
                   {
                     name = LowerCase("\\Изменения(курс).xlsx");
                   }
                 else
                   {
                     Application->MessageBox("Не найден файл для загрузки данных. \nВозможно указано НЕВЕРНОЕ ИМЯ файла \nили файл не найден в данной папке. ",
                                           "Ошибка загрузки данных", MB_OK + MB_ICONERROR);
                     Abort();
                   }
       break;
       case 6 :  if (FindFirst(Dir2 + LowerCase("\\Изменения(вр).xls"), faAnyFile, SearchRecord)==0 )
                   {
                     name = LowerCase("\\Изменения(вр).xls");
                   }
                 else if (FindFirst (Dir2 + LowerCase("\\Изменения(вр).xlsx"), faAnyFile, SearchRecord)==0)
                   {
                     name = LowerCase("\\Изменения(вр).xlsx");
                   }
                 else
                   {
                     Application->MessageBox("Не найден файл для загрузки данных. \nВозможно указано НЕВЕРНОЕ ИМЯ файла \nили файл не найден в данной папке. ",
                                           "Ошибка загрузки данных", MB_OK + MB_ICONERROR);
                     Abort();
                   }
       break;
       case 7 : if (FindFirst(Dir2 + LowerCase("\\Изменения(пенсионное).xls"), faAnyFile, SearchRecord)==0 )
                  {
                    name = LowerCase("\\Изменения(пенсионное).xls");
                  }
                else if (FindFirst (Dir2 + LowerCase("\\Изменения(пенсионное).xlsx"), faAnyFile, SearchRecord)==0)
                  {
                    name = LowerCase("\\Изменения(пенсионное).xlsx");
                  }
                else
                  {
                    Application->MessageBox("Не найден файл для загрузки данных. Возможно указано НЕВЕРНОЕ ИМЯ файла (должно быть 'Изменения(пенсионное).xls' или 'Изменения(пенсионное).xlsx') или файл не найден в данной папке.",
                                           "Ошибка загрузки данных", MB_OK + MB_ICONERROR);
                    Abort();
                  }
       break;


     }

  FileName = Dir2 + name;  //Путь к файлу Excel
  FindClose(SearchRecord);   //освобождает ресурсы, взятые процессом поиска
     
  StatusBar1->SimpleText = "";

   // инициализируем Excel, открываем этот шаблон
  try
    {
      //проверяем, нет ли запущенного Excel
      Excel = GetActiveOleObject("Excel.Application");
    }
  catch(...)
    {
      try
        {
          Excel = CreateOleObject("Excel.Application");
        }
      catch (...)
        {
          Application->MessageBox("Невозможно открыть Microsoft Excel!"
          " Возможно это приложение на компьютере не установлено.","Ошибка",MB_OK+MB_ICONERROR);
          Abort();
        }
    }

  try
    {
      Book = Excel.OlePropertyGet("Workbooks").OlePropertyGet("Open", FileName.c_str());
      Sheet = Book.OlePropertyGet("Worksheets", 1);
    }
  catch(...)
    {
      Application->MessageBox("Ошибка открытия книги Microsoft Excel!","Ошибка",MB_OK + MB_ICONERROR);
    }


  //Excel.OlePropertySet("Visible",true);

  //Определяет количество занятых строк в документе
  Row = Sheet.OlePropertyGet("UsedRange").OlePropertyGet("Rows").OlePropertyGet("Count");
//Row=100;
  // Открываем файл данных для формирования отчета по несуществующим ИНН
  if (!rtf_Open((TempPath + "\\otchet.txt").c_str()))
    {
      MessageBox(Handle,"Ошибка открытия файла данных","Ошибка",8192);
    }
  else
    {

      Main->Cursor = crHourGlass;
      StatusBar1->SimplePanel = true;    // 2 панели на StatusBar1
      StatusBar1->SimpleText=" Выполняется проверка данных...";
      ProgressBar->Visible = true;
      ProgressBar->Position = 0;
      ProgressBar->Max = Row;


      for ( i ; i<Row+1; i++)
        {
          nn = Excel.OlePropertyGet("Cells",i,1);
          inn = Excel.OlePropertyGet("Cells",i,5);
          ProgressBar->Position++;


          // Выбор строк необходимых для загрузки из Excel
          if (nn.IsEmpty() || !Proverka(nn) || inn.IsEmpty())  continue;

            inn = Excel.OlePropertyGet("Cells",i,5);
            if (im_fl==7) sum = Excel.OlePropertyGet("Cells",i,9);
            else sum = Excel.OlePropertyGet("Cells",i,10);
            fam = TrimRight(""+Excel.OlePropertyGet("Cells",i,2)+" "+Excel.OlePropertyGet("Cells",i,3)+" "+Excel.OlePropertyGet("Cells",i,4));
            n_doc = Excel.OlePropertyGet("Cells",i,11);
            data_po = Excel.OlePropertyGet("Cells",i,7);



//Проверка на несколько записей в sap_osn_sved и sap_sved_uvol и наличие инд.№
//******************************************************************************
            Sql1 = "select tn_sap, numident, 1 as priznak from sap_osn_sved where numident=:pnumident                \
                    union all                                                                          \
                    select tn_sap, numident, 2 as priznak from sap_sved_uvol                                         \
                    where numident=:pinn and substr(to_char(dat_job,'dd.mm.yyyy'),4,7)<='"+(DM->mm<10 ? "0"+IntToStr(DM->mm) : IntToStr(DM->mm))+"."+DM->yyyy+"'";

            try
              {
                DM->qObnovlenie->Close();
                DM->qObnovlenie->SQL->Clear();
                DM->qObnovlenie->SQL->Add(Sql1);
                DM->qObnovlenie->Parameters->ParamByName("pnumident")->Value=inn;
                DM->qObnovlenie->Parameters->ParamByName("pinn")->Value=inn;
                DM->qObnovlenie->Open();
              }
            catch(...)
              {
                Application->MessageBox("Невозможно получить данные из картотеки работников(SAP_OSN_SVED, SAP_SVED_UVOL)","Ошибка",MB_OK + MB_ICONERROR);
                Abort();
              }

            if (DM->qObnovlenie->RecordCount>1)
              {
                 pr_inn=0;
                 pr_sum=0;
                //Вывод в отчет двойных записей
//******************************************************************************
                //Вывод наименования и шапки таблицы
                if (DM->qObnovlenie->RecordCount>1 && pr_dv==0)
                  {
                    rtf_Out("z", " ",3);
                    if(!rtf_LineFeed())
                      {
                        MessageBox(Handle,"Ошибка записи в файл данных","Ошибка",8192);
                        if (!rtf_Close()) MessageBox(Handle,"Ошибка закрытия файла данных","Ошибка",8192);
                        return;
                      }
                   }
                //Вывод записей в отчет
                rtf_Out("inn", inn,4);
                rtf_Out("fio",fam,4);

                if(!rtf_LineFeed())
                  {
                    MessageBox(Handle,"Ошибка записи в файл данных","Ошибка",8192);
                    if (!rtf_Close()) MessageBox(Handle,"Ошибка закрытия файла данных","Ошибка",8192);
                    return;
                  }
                fl=1;
                pr_dv=1;

              }
            else if (DM->qObnovlenie->RecordCount==0)
              {
               //нет записей ни в таблице sap_osn_sved, ни в sap_sved_uvol
                // Вывод несуществующего инд.№ в отчет
//******************************************************************************
                if (DM->qObnovlenie->RecordCount==0)
                  {
                     pr_dv=0;
                     pr_sum=0;
                   //Вывод наименования и шапки таблицы
                    if (DM->qObnovlenie->RecordCount==0 && pr_inn==0)
                      {
                        rtf_Out("z", " ",1);
                        if(!rtf_LineFeed())
                          {
                            MessageBox(Handle,"Ошибка записи в файл данных","Ошибка",8192);
                            if (!rtf_Close()) MessageBox(Handle,"Ошибка закрытия файла данных","Ошибка",8192);
                            return;
                          }
                      }


                    rtf_Out("inn",inn,2);
                    rtf_Out("fio",fam,2);

                    if(!rtf_LineFeed())
                      {
                        MessageBox(Handle,"Ошибка записи в файл данных","Ошибка",8192);
                        if (!rtf_Close()) MessageBox(Handle,"Ошибка закрытия файла данных","Ошибка",8192);
                        return;
                      }

                    fl=1;

                    pr_inn=1;

                  }
              }
           //Работник уволен текущим месяцем
           else if (DM->qObnovlenie->RecordCount==1 && DM->qObnovlenie->FieldByName("priznak")->AsInteger==2 )
              {
               //нет записей ни в таблице sap_osn_sved, ни в sap_sved_uvol
                // Вывод несуществующего инд.№ в отчет
//******************************************************************************
                if (DM->qObnovlenie->RecordCount==1 && DM->qObnovlenie->FieldByName("priznak")->AsInteger==2 )
                  {
                     pr_dv=0;
                     pr_inn=0;
                   //Вывод наименования и шапки таблицы
                    if (DM->qObnovlenie->RecordCount==1 && pr_sum==0)
                      {
                        rtf_Out("z", " ",5);
                        if(!rtf_LineFeed())
                          {
                            MessageBox(Handle,"Ошибка записи в файл данных","Ошибка",8192);
                            if (!rtf_Close()) MessageBox(Handle,"Ошибка закрытия файла данных","Ошибка",8192);
                            return;
                          }
                      }


                    rtf_Out("inn",inn,6);
                    rtf_Out("fio",fam,6);

                    if(!rtf_LineFeed())
                      {
                        MessageBox(Handle,"Ошибка записи в файл данных","Ошибка",8192);
                        if (!rtf_Close()) MessageBox(Handle,"Ошибка закрытия файла данных","Ошибка",8192);
                        return;
                      }

                    fl=1;

                    pr_sum=1;

                  }
              }


            else
              {
           //Проверка на превышение удерживаемой суммы свыше 15%,
           // если есть возобновление (дата не пустая) по изменению по гривне
           // и изменению по валюте
//******************************************************************************
/*
                if (!data_po.IsEmpty()&&(im_fl==2||im_fl==3||im_fl==6))
                  {
                    Sql2 = "select (sum(decode(typs,3,sum*-1,sum))*15/100) sum from slst"+(DM->mm2 < 10 ? "0" + IntToStr(DM->mm2) : IntToStr(DM->mm2))+ DM->yyyy2 + " \
                            where klus="+nnom+" \
                            and typs in (1,3,5) and vo<800";

                    DM->qObnovlenie->Close();
                    DM->qObnovlenie->SQL->Clear();
                    DM->qObnovlenie->SQL->Add(Sql2);
                    DM->qObnovlenie->Open();

                    if (DM->qObnovlenie->FieldByName("sum")->AsString.IsEmpty())
                      {
                        if (Application->MessageBox(("Нет суммы за прошлый месяц\nцех="+zex+" таб.№="+tn+" ФИО="+fam+" сумма="+FloatToStrF(sum,ffFixed,20,2)+" \nЗагрузить запись в таблицу?").c_str(),
                                                    "Превышение",MB_YESNO + MB_ICONINFORMATION)==IDNO)
                          {
                            pr_inn=0;
                            pr_dv=0;
                            // Вывод в отчет если нет суммы за прошлый месяц
//******************************************************************************
                            //Вывод наименования и шапки таблицы
                            if ((sum >= DM->qObnovlenie->FieldByName("sum")->AsFloat) && pr_sum==0)
                              {
                                rtf_Out("zz", " ",3);

                                if(!rtf_LineFeed())
                                  {
                                    MessageBox(Handle,"Ошибка записи в файл данных","Ошибка",8192);
                                    if (!rtf_Close()) MessageBox(Handle,"Ошибка закрытия файла данных","Ошибка",8192);
                                    return;
                                  }
                              }

                            rtf_Out("zex", zex,4);
                            rtf_Out("tn", tn,4);
                            rtf_Out("fio",fam,4);
                            rtf_Out("n_doc",n_doc ,4);
                            rtf_Out("sum","нет суммы прошлого месяца",4);

                            if(!rtf_LineFeed())
                              {
                                MessageBox(Handle,"Ошибка записи в файл данных","Ошибка",8192);
                                if (!rtf_Close()) MessageBox(Handle,"Ошибка закрытия файла данных","Ошибка",8192);
                                return;
                              }

                            fl=1;
                            pr_sum=1;
                          }
                      }
                   else if (sum > DM->qObnovlenie->FieldByName("sum")->AsFloat)
                     {
                       if (Application->MessageBox(("Сумма превышает 15%\nцех="+zex+" таб.№="+tn+" ФИО="+fam+" сумма="+FloatToStrF(sum,ffFixed,20,2)+" \nЗагрузить запись в таблицу?").c_str(),
                                                    "Превышение",MB_YESNO + MB_ICONINFORMATION)==IDNO)
                         {
                           pr_inn=0;
                           pr_dv=0;
                           // Вывод в отчет превышающей 15% суммы
//******************************************************************************
                           //Вывод наименования и шапки таблицы
                           if ((sum > DM->qObnovlenie->FieldByName("sum")->AsFloat) && pr_sum==0)
                             {
                               rtf_Out("zz", " ",3);

                               if(!rtf_LineFeed())
                                 {
                                   MessageBox(Handle,"Ошибка записи в файл данных","Ошибка",8192);
                                   if (!rtf_Close()) MessageBox(Handle,"Ошибка закрытия файла данных","Ошибка",8192);
                                   return;
                                 }
                             }

                           rtf_Out("zex", zex,4);
                           rtf_Out("tn", tn,4);
                           rtf_Out("fio",fam,4);
                           rtf_Out("n_doc",n_doc ,4);
                           rtf_Out("sum",FloatToStrF(sum,ffFixed,20,2),4);

                           if(!rtf_LineFeed())
                             {
                               MessageBox(Handle,"Ошибка записи в файл данных","Ошибка",8192);
                               if (!rtf_Close()) MessageBox(Handle,"Ошибка закрытия файла данных","Ошибка",8192);
                               return;
                             }

                           fl=1;
                           pr_sum=1;
                         }
                     }
                  }*/
              }
        }

      StatusBar1->SimplePanel = false;
      ProgressBar->Visible = false;
      StatusBar1->SimpleText = "Проверка данных выполнена.";
      Main->Cursor = crDefault;

      if(!rtf_Close())
        {
          MessageBox(Handle,"Ошибка закрытия файла данных", "Ошибка", 8192);
          return;
        }


      if (fl==1)
        {
          Excel.OleProcedure("Quit");
          StatusBar1->SimpleText = "Формирование отчета...";
          //Создание папки, если ее не существует
          ForceDirectories(WorkPath);

          int istrd;
          try
            {
              rtf_CreateReport(TempPath +"\\otchet.txt", Path+"\\RTF\\otchet.rtf",
                         WorkPath+"\\Отчет.doc",NULL,&istrd);


              WinExec(("\""+ WordPath+"\"\""+WorkPath+"\\Отчет.doc\"").c_str(),SW_MAXIMIZE);

            }
          catch(RepoRTF_Error E)
            {
              MessageBox(Handle,("Ошибка формирования отчета:"+ AnsiString(E.Err)+
                                 "\nСтрока файла данных:"+IntToStr(istrd)).c_str(),"Ошибка",8192);
            }

          Application->MessageBox(("Проверьте достоверность информации в файле \n \""+FileName+"\" и выполните повторную загрузку").c_str() ," Загрузка новых договоров по ММК",
                                  MB_OK + MB_ICONINFORMATION);
          StatusBar1->SimpleText = "";

          switch (im_fl)
             {
               case 2: InsertLog("Сформирован отчет по Изменениям(гривна): нет данных по ИНН");
               break;
               case 3: InsertLog("Сформирован отчет по Изменениям(валюта): нет данных по ИНН");
               break;
               case 4: InsertLog("Сформирован отчет по Изменениям курса: нет данных по ИНН");
               break;
               case 6: InsertLog("Сформирован отчет по Изменениям ВР: нет данных по ИНН");
               break;
               case 7: InsertLog("Сформирован отчет по Изменениям по пенсионному страхованию: нет данных по ИНН");
               break;
             }

          Abort();
        }

         DeleteFile(TempPath+"\\otchet.txt");
    }


}
//---------------------------------------------------------------------------
void __fastcall TMain::N18Click(TObject *Sender)
{
 AnsiString Sql;
 double sum_dog=0;
 int val=0,val1=0, kol_dog=0;


  /*kust - номер текущего куста, kust1 - номер следующего куста;
  double sum_kust - сумма по кусту */

  Sql = "select distinct zex,priznak, kod_dogovora,                     \
                sum(sum) over (partition by  kod_dogovora,zex) sumzex,  \
                sum(sum) over (partition by kod_dogovora) sumdog,       \
                count(*) over (partition by kod_dogovora) koldog,       \
                sum(sum) over () sumobsh,                               \
                count(*) over () kolobsh                                \
         from vu_859_n                                                  \
         where priznak=0 and nvl(sum,0)!=0                              \
         order by kod_dogovora, zex";

  DM->qObnovlenie->Close();
  DM->qObnovlenie->SQL->Clear();
  DM->qObnovlenie->SQL->Add(Sql);
  try
    {
      DM->qObnovlenie->Open();
    }
  catch(...)
    {
      MessageBox(Handle,"Возникла ошибка при выборке данных по отчету Сводная ведомость удержаний для УИТ","Ошибка",8202);
      Abort();
    }


  if (DM->qObnovlenie->RecordCount>0)
    {
      //Создание папки, если ее не существует
      ForceDirectories(WorkPath);

      if (!rtf_Open ((TempPath + "\\v_yit.txt").c_str()))
        {
          MessageBox(Handle,"Ошибка открытия файла данных","Ошибка",8192);
        }
      else
        {
          StatusBar1->SimpleText = "Формирование сводной ведомости удержаний для УИТ...";
          rtf_Out("mes", Mes[DM->mm-1], 0);
          rtf_Out("god", DM->yyyy, 0);

          val = DM->qObnovlenie->FieldByName("kod_dogovora")->AsInteger;
          val1 = DM->qObnovlenie->FieldByName("kod_dogovora")->AsInteger;

          while (!DM->qObnovlenie->Eof)
            {
              while(!DM->qObnovlenie->Eof && val==val1)
                {
                  rtf_Out("zex", DM->qObnovlenie->FieldByName("zex")->AsString, 1);
                  rtf_Out("sum", DM->qObnovlenie->FieldByName("sumzex")->AsFloat,10,2, 1);
                  sum_dog = DM->qObnovlenie->FieldByName("sumdog")->AsFloat;
                  kol_dog = DM->qObnovlenie->FieldByName("koldog")->AsInteger;

                  if(!rtf_LineFeed())
                    {
                      MessageBox(Handle,"Ошибка записи в файл данных","Ошибка",8192);
                      if (!rtf_Close()) MessageBox(Handle,"Ошибка закрытия файла данных","Ошибка",8192);
                      return;
                    }

                  DM->qObnovlenie->Next();
                  val1 = DM->qObnovlenie->FieldByName("kod_dogovora")->AsInteger;
                }

              // вывод суммы по договору
              switch (val)
                {
                  case 0: rtf_Out("naim", "гривневым", 2);
                  break;
                  case 1: rtf_Out("naim", "валютным(доллар)", 2);
                  break;
                  case 2: rtf_Out("naim", "валютным(евро)", 2);
                  break;
                  case 3: rtf_Out("naim", "внешним", 2);
                  break;

                }

              rtf_Out("sumdog", sum_dog,10,2, 2);
              rtf_Out("koldog", kol_dog, 2);

              val = DM->qObnovlenie->FieldByName("kod_dogovora")->AsInteger;
              if(!rtf_LineFeed())
                {
                   MessageBox(Handle,"Ошибка записи в файл данных","Ошибка",8192);
                   if (!rtf_Close()) MessageBox(Handle,"Ошибка закрытия файла данных","Ошибка",8192);
                   return;
                }
            }

          // вывод суммы по предприятию
          rtf_Out("sumobsh", DM->qObnovlenie->FieldByName("sumobsh")->AsString, 0);
          rtf_Out("kolobsh", DM->qObnovlenie->FieldByName("kolobsh")->AsString, 0);

          if(!rtf_Close())
            {
              MessageBox(Handle,"Ошибка закрытия файла данных", "Ошибка", 8192);
              return;
            }

          int istrd;
          try
            {
              rtf_CreateReport(TempPath +"\\v_yit.txt", Path+"\\RTF\\v_yit.rtf",
                               WorkPath+"\\Сводная ведомость удержаний для УИТ.doc",NULL,&istrd);
              DeleteFile(TempPath+"\\v_yit.txt");

              WinExec(("\""+ WordPath+"\"\""+WorkPath+"\\Сводная ведомость удержаний для УИТ.doc\"").c_str(),SW_MAXIMIZE);
            }
          catch(RepoRTF_Error E)
            {
              MessageBox(Handle,("Ошибка формирования отчета:"+ AnsiString(E.Err)+
                                 "\nСтрока файла данных:"+IntToStr(istrd)).c_str(),"Ошибка",8192);
            }
         StatusBar1->SimpleText = "";
        }
    }
  else
   {
     Application->MessageBox("Нет данных за текущий месяц", "Предупреждение",
                              MB_OK + MB_ICONWARNING);
   }

}
//---------------------------------------------------------------------------

void __fastcall TMain::N19Click(TObject *Sender)
{
  im_fl=6;
  if (Application->MessageBox(("Вы действительно хотите загрузить изменения \n по валютным договорам за " + Mes[DM->mm-1] + " " + DM->yyyy + " года?").c_str(),
                               "Загрузка изменений по валютным договорам",
                               MB_YESNO + MB_ICONINFORMATION) == IDNO)
    {
      Abort();
    }

  // Проверка на существование инн в таблице Avans
  ProverkaInfoExcelIzmeneniya();

  StatusBar1->SimpleText = "";

  //Обновление изменений по валютной страховке
  UpdateValuta_I_Grivna();

  InsertLog("Выполнена загрузка изменений по внешним договорам. Обновлено "+obnov_kol+" из "+ob_kol+" записей");

  StatusBar1->SimpleText = "";
}
//---------------------------------------------------------------------------

//Редактирование записи
void __fastcall TMain::N20Click(TObject *Sender)
{
  Panel1->Visible = true;
  EditZEX->Visible=true;
  EditTN->Visible=true;
  BitBtn3->Visible=true;
  EditZEX->SetFocus();
  EditZEX->Text="";
  EditTN->Text="";
  Label1->Visible=true;

  DBGridEh1->Visible=false;
  
  EditZEX2->Visible=false;
  EditTN2->Visible=false;
  EditSum->Visible=false;
  EditData_s->Visible=false;
  EditData_po->Visible=false;
  EditVal->Visible=false;
  BitBtn1->Visible=false;
  BitBtn2->Visible=false;
  Label2->Visible=false;
  Label3->Visible=false;
  Label4->Visible=false;
  Label5->Visible=false;
  Label6->Visible=false;
  Label7->Visible=false;
  Bevel1->Visible=false;
  Bevel3->Visible=false;
  Label8->Visible=false;
  Label9->Visible=false;
  Label10->Visible=false;

  Label11->Visible=false;
  Label12->Visible = false;

  Label14->Visible = false;
  LabelNDOG->Visible = false;
  EditNDOG->Visible = false;
  EditPRIZNAK->Visible = false;
  
  fl_r =1;
}
//---------------------------------------------------------------------------

//Добавление записи
void __fastcall TMain::N21Click(TObject *Sender)
{
  Panel1->Visible = true;
  EditNDOG->Visible=true;
  EditNDOG->SetFocus();
 // EditZEX->Text="";
 // EditTN->Text="";

  EditZEX->Visible=false;
  EditTN->Visible=false;
  BitBtn3->Visible=false;
  Label1->Visible=false;
  Label9->Caption="Введите данные:";
  Label11->Visible=false;
  Label12->Visible=false;
  Label14->Visible = true;
  LabelNDOG->Visible = true;
  EditNDOG->Visible = true;

  
  DBGridEh1->Visible=true;

  EditZEX2->Visible=true;
  EditTN2->Visible=true;
  EditSum->Visible=true;
  EditData_s->Visible=true;
  EditData_po->Visible=true;
  EditVal->Visible=true;
  BitBtn1->Visible=true;
  BitBtn2->Visible=true;
  Label2->Visible=true;
  Label3->Visible=true;
  Label4->Visible=true;
  Label5->Visible=true;
  Label6->Visible=true;
  Label7->Visible=true;
  Bevel1->Visible=true;
  Bevel3->Visible=true;
  Label8->Visible=true;
  Label9->Visible=true;
  Label10->Visible=false;

  EditPRIZNAK->Visible = false;
  SetEditNull();
  fl_r=0;

}
//---------------------------------------------------------------------------
 void __fastcall TMain::SetEditNull()
{
  EditNDOG->Text="";
  EditZEX2->Text="";
  EditTN2->Text="";
  EditSum->Text="";
  EditData_s->Text="";
  EditData_po->Text="";
  EditVal->Text="";
  EditVal->Text="";
  EditPRIZNAK->Text ="";
  DM->qKorrektirovka->Close();
}
//---------------------------------------------------------------------------
void __fastcall TMain::EditNDOGKeyPress(TObject *Sender, char &Key)
{
  if (Key=='/'||Key==',') Key='.';

}
//---------------------------------------------------------------------------

void __fastcall TMain::N23Click(TObject *Sender)
{
  WinExec(("\""+ WordPath+"\"\""+ Path+"\\Инструкция пользователя.doc\"").c_str(),SW_MAXIMIZE);
}
//---------------------------------------------------------------------------

void __fastcall TMain::N24Click(TObject *Sender)
{
  WinExec(("\""+ WordPath+"\"\""+ Path+"\\Пошаговая инструкция.doc\"").c_str(),SW_MAXIMIZE);
}
//---------------------------------------------------------------------------

void __fastcall TMain::N25Click(TObject *Sender)
{
 AnsiString Sql;
 double sum_dog=0, sum_obsh=0;
 int val=0,val1=0, kol_dog=0, kol_obsh, kust=0, kust1=0;

  StatusBar1->SimpleText = "Формирование сводной ведомости удержаний...";

  Sql="select distinct zex, priznak,kod_dogovora, 1 as ana,                \
                sum(sum) over (partition by kod_dogovora, zex) sumzex,     \
                count(*) over (partition by kod_dogovora) koldog,          \
                sum(sum) over (partition by kod_dogovora) sumdog,          \
                sum(sum) over () sumobsh,                                  \
                count(*) over () kolobsh                                   \
         from vu_859_n                                                     \
         where priznak=0 and nvl(sum,0)!=0                                 \
         and (inn in (select numident from sap_osn_sved)                   \
         or  inn in (select numident from sap_sved_uvol where substr(to_char(dat_job,'dd.mm.yyyy'),4,7)='"+(DM->mm<10 ? "0"+IntToStr(DM->mm) : IntToStr(DM->mm))+"."+DM->yyyy+"')) \
         order by ana, kod_dogovora,zex";


  DM->qObnovlenie->Close();
  DM->qObnovlenie->SQL->Clear();
  DM->qObnovlenie->SQL->Add(Sql);
  try
    {
      DM->qObnovlenie->Open();
    }
  catch(...)
    {
      MessageBox(Handle,"Возникла ошибка при выборке данных по отчету Сводная ведомость удержаний для УИТ","Ошибка",8202);
      StatusBar1->SimpleText = "";
      Abort();
    }


  if (DM->qObnovlenie->RecordCount>0)
    {
      //Создание папки, если ее не существует
      ForceDirectories(WorkPath);

      if (!rtf_Open ((TempPath + "\\v_yit2.txt").c_str()))
        {
          MessageBox(Handle,"Ошибка открытия файла данных","Ошибка",8192);
        }
      else
        {
          rtf_Out("data", Now(),0);
          rtf_Out("mes", Mes[DM->mm-1], 0);
          rtf_Out("god", DM->yyyy, 0);

          val = DM->qObnovlenie->FieldByName("kod_dogovora")->AsInteger;
          val1 = DM->qObnovlenie->FieldByName("kod_dogovora")->AsInteger;
          kust = DM->qObnovlenie->FieldByName("ana")->AsInteger;
          kust1 = DM->qObnovlenie->FieldByName("ana")->AsInteger;

          while (!DM->qObnovlenie->Eof)
            {
              kust = DM->qObnovlenie->FieldByName("ana")->AsInteger;

              while(!DM->qObnovlenie->Eof && kust==kust1)
                {
                  kust = DM->qObnovlenie->FieldByName("ana")->AsInteger;

                  while(!DM->qObnovlenie->Eof && val==val1)
                    {
                      rtf_Out("zex", DM->qObnovlenie->FieldByName("zex")->AsString, 1);
                      rtf_Out("sum", DM->qObnovlenie->FieldByName("sumzex")->AsFloat,10,2, 1);
                      sum_dog = DM->qObnovlenie->FieldByName("sumdog")->AsFloat;
                      kol_dog = DM->qObnovlenie->FieldByName("koldog")->AsInteger;
                      sum_obsh = DM->qObnovlenie->FieldByName("sumobsh")->AsFloat;
                      kol_obsh = DM->qObnovlenie->FieldByName("kolobsh")->AsInteger;

                      if(!rtf_LineFeed())
                        {
                          MessageBox(Handle,"Ошибка записи в файл данных","Ошибка",8192);
                          if (!rtf_Close()) MessageBox(Handle,"Ошибка закрытия файла данных","Ошибка",8192);
                          return;
                        }

                       DM->qObnovlenie->Next();
                       val1 = DM->qObnovlenie->FieldByName("kod_dogovora")->AsInteger;
                       kust1 = DM->qObnovlenie->FieldByName("ana")->AsInteger;
                    }

                  // вывод суммы по договору
                  switch (val)
                    {
                      case 0: rtf_Out("naim", "гривневым", 2);
                      break;
                      case 1: rtf_Out("naim", "валютным(доллар)", 2);
                      break;
                      case 2: rtf_Out("naim", "валютным(евро)", 2);
                      break;
                      case 3: rtf_Out("naim", "внешним", 2);
                      break;
                    }

                   rtf_Out("sumdog", sum_dog,10,2, 2);
                   rtf_Out("koldog", kol_dog, 2);

                   val = DM->qObnovlenie->FieldByName("kod_dogovora")->AsInteger;
 //                  kust = DM->qObnovlenie->FieldByName("ana")->AsInteger;
                   if(!rtf_LineFeed())
                     {
                       MessageBox(Handle,"Ошибка записи в файл данных","Ошибка",8192);
                       if (!rtf_Close()) MessageBox(Handle,"Ошибка закрытия файла данных","Ошибка",8192);
                       return;
                     }
                }

               // вывод суммы по кусту
               switch (kust)
                 {
                   case 1: rtf_Out("naim2", "промкомплексу", 3);
                   break;
                   default: rtf_Out("naim2", "агрофирмам", 3);
                   break;

                 }

               rtf_Out("sumobsh", sum_obsh,10,2, 3);
               rtf_Out("kolobsh", kol_obsh, 3);
               kust = DM->qObnovlenie->FieldByName("ana")->AsInteger;
               if(!rtf_LineFeed())
                 {
                   MessageBox(Handle,"Ошибка записи в файл данных","Ошибка",8192);
                   if (!rtf_Close()) MessageBox(Handle,"Ошибка закрытия файла данных","Ошибка",8192);
                   return;
                 }

            }
          if(!rtf_Close())
            {
              MessageBox(Handle,"Ошибка закрытия файла данных", "Ошибка", 8192);
              return;
            }

          int istrd;
          try
            {
              rtf_CreateReport(TempPath +"\\v_yit2.txt", Path+"\\RTF\\v_yit2.rtf",
                               WorkPath+"\\Сводная ведомость удержаний для УИТ.doc",NULL,&istrd);
              DeleteFile(TempPath+"\\v_yit2.txt");

              WinExec(("\""+ WordPath+"\"\""+WorkPath+"\\Сводная ведомость удержаний для УИТ.doc\"").c_str(),SW_MAXIMIZE);
            }
          catch(RepoRTF_Error E)
            {
              MessageBox(Handle,("Ошибка формирования отчета:"+ AnsiString(E.Err)+
                                 "\nСтрока файла данных:"+IntToStr(istrd)).c_str(),"Ошибка",8192);
            }
         StatusBar1->SimpleText = "";
        }
    }
  else
   {
     Application->MessageBox("Нет данных за текущий месяц", "Предупреждение",
                              MB_OK + MB_ICONWARNING);
   }

}
//---------------------------------------------------------------------------

void __fastcall TMain::EditPRIZNAKKeyPress(TObject *Sender, char &Key)
{
  if (!(IsNumeric(Key)||Key=='\b')) Key=0;
  if (Key=='5'|| Key=='8' || Key=='9') Key=0;

  if (Key != '\0'){switch (Key)
    {
      case '0':  Label11->Caption="платит";
      break;
      case '1':  Label11->Caption="уволен";
      break;
      case '2':  Label11->Caption="закончился срок договора";
      break;
      case '3':  Label11->Caption="декрет";
      break;
      case '4':  Label11->Caption="расторжение";
      break;
      case '5':  Label11->Caption="приостановлен";
      break;
      case '6':  Label11->Caption="приостановлен";
      break;
      case '7':  Label11->Caption=" не вступил в силу";
      break;
      default :  Label11->Caption="";
    }  }

}
//---------------------------------------------------------------------------

// Выгрузка в Excel для САП
void __fastcall TMain::ExcelSAP(int valuta)
{

  AnsiString sFile, Sql, tn, tn1;
  int n=2;
  Variant AppEx, Sh;

  if (Application->MessageBox(("Будет выполнено формирование Excel-файла за "+Mes[DM->mm-1]+" "+DM->yyyy+" года.\nПродолжить?").c_str(),"Предупреждение",
                              MB_YESNO+MB_ICONINFORMATION)==ID_NO)
    {
      Abort();
    }

  StatusBar1->SimpleText=" Идет формирование файла в Excel...";

  DecimalSeparator=',';

  Sql="select case when inn in (select numident from sap_osn_sved)                              \
                   then (select tn_sap from sap_osn_sved where numident = inn)                  \
                   else (select tn_sap from sap_sved_uvol where numident = inn) end as tn_sap,  \
              decode(nvl(kod_dogovora,0),'0','7428','1','7433','2','7433','3','7434','4','7428') as vo,                                                                                                                                                       \
              '01.'||substr(to_char(sysdate,'dd.mm.yyyy'),4,7) as datn,                                                                                                           \
              last_day(to_date(to_char(sysdate,'dd.mm.yyyy'), 'dd.mm.yyyy')) as datk,                                                                                             \
               sum,                                                                                                                                                               \
              'UAH'  as valuta,                                                                                                                                                   \
              '100101103' as num,    \
              n_dogovora                                                                                                                                                    \
       from vu_859_n where nvl(priznak,0)=0 and nvl(sum,0)>0                                                                                                                      \
       and (inn in (select numident from sap_osn_sved)                                          \
       or inn in (select numident from sap_sved_uvol where substr(to_char(dat_job,'dd.mm.yyyy'),4,7)='"+(DM->mm <10? "0"+IntToStr(DM->mm):IntToStr(DM->mm))+"."+IntToStr(DM->yyyy)+ "') )";

  if (valuta==1) Sql+=" and nvl(kod_dogovora,0)=0 order by tn_sap, nvl(kod_dogovora,0), sum";
  else if (valuta==2) Sql+=" and nvl(kod_dogovora,0) in (1,2) order by tn_sap, nvl(kod_dogovora,0), sum";
  else if (valuta==3) Sql+=" and nvl(kod_dogovora,0)=3 order by tn_sap, nvl(kod_dogovora,0), sum";
  else if (valuta==4) Sql+=" and nvl(kod_dogovora,0)=4 order by tn_sap, nvl(kod_dogovora,0), sum";


  //decode(translate(inn,'-0123456789 ','-'),null, inn,substr(inn,1,2)||substr(inn,5,10)) in (select numident from sap_osn_sved)

  DM->qObnovlenie->Close();
  DM->qObnovlenie->SQL->Clear();
  DM->qObnovlenie->SQL->Add(Sql);
  try
    {
      DM->qObnovlenie->Open();
    }
  catch(Exception &E)
    {
      Application->MessageBox(("Возникла ошибка при получении данных из таблицы по страховке или картотеке по работникам (VU_859_N, SAP_SVED_UVOL, SAP_OSN_SVED)" + E.Message).c_str(),"Ошибка",
                              MB_OK+MB_ICONERROR);
      if (valuta==1) InsertLog("Возникла ошибка при формировании файла по гривневой страховке для САП в Excel");
      else if (valuta==2) InsertLog("Возникла ошибка при формировании файла по валютной страховке для САП в Excel");
      else if (valuta==3) InsertLog("Возникла ошибка при формировании файла по внешним договорам для САП в Excel");
      else if (valuta==4) InsertLog("Возникла ошибка при формировании файла по пенсионным договорам для САП в Excel");

      StatusBar1->SimpleText="";
      Abort();
    }


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
      Application->MessageBox("Невозможно открыть Microsoft Excel!"
                              " Возможно это приложение на компьютере не установлено.","Ошибка",MB_OK+MB_ICONERROR);
      StatusBar1->SimpleText="";
      ProgressBar->Visible = false;
      Cursor = crDefault;
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

          if (valuta==1)
            {
              DeleteFile(WorkPath+"\\SAP 0015ИТ(гривна).xlsx");
              CopyFile((Path+"\\RTF\\sap.xlsx").c_str(), (WorkPath+"\\SAP 0015ИТ(гривна).xlsx").c_str(), false);
              sFile = WorkPath+"\\SAP 0015ИТ(гривна).xlsx";
            }
          else if (valuta==2)
            {
              DeleteFile(WorkPath+"\\SAP 0015ИТ(валюта).xlsx");
              CopyFile((Path+"\\RTF\\sap.xlsx").c_str(), (WorkPath+"\\SAP 0015ИТ(валюта).xlsx").c_str(), false);
              sFile = WorkPath+"\\SAP 0015ИТ(валюта).xlsx";
            }
          else if (valuta==3)
            {
              DeleteFile(WorkPath+"\\SAP 0015ИТ(внешние договора).xlsx");
              CopyFile((Path+"\\RTF\\sap.xlsx").c_str(), (WorkPath+"\\SAP 0015ИТ(внешние договора).xlsx").c_str(), false);
              sFile = WorkPath+"\\SAP 0015ИТ(внешние договора).xlsx";
            }
          else if (valuta==4)
            {
              DeleteFile(WorkPath+"\\SAP 0015ИТ(пенсионные договора).xlsx");
              CopyFile((Path+"\\RTF\\sap.xlsx").c_str(), (WorkPath+"\\SAP 0015ИТ(пенсионные договора).xlsx").c_str(), false);
              sFile = WorkPath+"\\SAP 0015ИТ(пенсионные договора).xlsx";
            }

          AppEx.OlePropertyGet("WorkBooks").OleProcedure("Open",sFile.c_str())    ;  //открываем книгу, указав её имя

          Sh=AppEx.OlePropertyGet("WorkSheets",1);                               //выбираем № активного листа книги
          //Sh=AppEx.OlePropertyGet("WorkSheets","Расчет");                      //выбираем лист по наименованию
        }
      catch(...)
        {
          Application->MessageBox("Ошибка открытия книги Microsoft Excel!","Ошибка",MB_OK+MB_ICONERROR);
          StatusBar1->SimpleText="";
          ProgressBar->Visible = false;
          Cursor = crDefault;
          DecimalSeparator='.';
          if (valuta==1) InsertLog("Возникла ошибка при формировании файла по гривневой страховке для САП в Excel");
          else if (valuta==2) InsertLog("Возникла ошибка при формировании файла по валютной страховке для САП в Excel");
          else if (valuta==3) InsertLog("Возникла ошибка при формировании файла по внешним договорам для САП в Excel");
          else if (valuta==4) InsertLog("Возникла ошибка при формировании файла по пенсионным договорам для САП в Excel");
        }

      int i=1;
      n=2;
      int d=-1;
      
      Variant Massiv;
      Massiv = VarArrayCreate(OPENARRAY(int,(0,13)),varVariant); //массив на 11 элементов

      tn=DM->qObnovlenie->FieldByName("tn_sap")->AsString;
      tn1=DM->qObnovlenie->FieldByName("tn_sap")->AsString;

      while (!DM->qObnovlenie->Eof)
        {
          if (tn==tn1) d++;
          else d=0;

          Massiv.PutElement(DM->qObnovlenie->FieldByName("tn_sap")->AsString.c_str(), 0);
          Massiv.PutElement(DM->qObnovlenie->FieldByName("vo")->AsString.c_str(), 1);
          
          Massiv.PutElement(DM->qObnovlenie->FieldByName("datn")->AsDateTime+d, 2);
          Massiv.PutElement(DM->qObnovlenie->FieldByName("datn")->AsDateTime+d, 3);
          //Massiv.PutElement(DM->qObnovlenie->FieldByName("datk")->AsString.c_str(), 3);
          Massiv.PutElement(FloatToStrF(DM->qObnovlenie->FieldByName("sum")->AsFloat, ffFixed,10,2), 4);
          Massiv.PutElement(DM->qObnovlenie->FieldByName("valuta")->AsString.c_str(), 7);

          if (valuta==4) Massiv.PutElement("200000650", 8);
          else Massiv.PutElement(DM->qObnovlenie->FieldByName("num")->AsString.c_str(), 8);
          Massiv.PutElement(DM->qObnovlenie->FieldByName("n_dogovora")->AsString.c_str(), 9);


          Sh.OlePropertyGet("Range", ("A" + IntToStr(n) + ":J" + IntToStr(n)).c_str()).OlePropertySet("Value", Massiv); //строка с данными с ячейки A по ячейку АВ

          i++;
          n++;
          tn=DM->qObnovlenie->FieldByName("tn_sap")->AsString;

          DM->qObnovlenie->Next();
          ProgressBar->Position++;
          tn1=DM->qObnovlenie->FieldByName("tn_sap")->AsString;
        }

       // Sh.OlePropertyGet("Range", ("LQ18:LQ" + IntToStr(i-1)).c_str()).OlePropertySet("NumberFormat", "0.00");

      //окрашивание ячеек
   /*   Sh.OlePropertyGet("Range",("M18:M"+IntToStr(n-1)).c_str()).OlePropertyGet("Interior").OlePropertySet("Color",14277081);
      Sh.OlePropertyGet("Range",("P18:R"+IntToStr(n-1)).c_str()).OlePropertyGet("Interior").OlePropertySet("Color",14277081);

      Sh.OlePropertyGet("Range",("B18:K"+IntToStr(n-1)).c_str()).OlePropertyGet("Interior").OlePropertySet("Color",14408946);
      Sh.OlePropertyGet("Range",("N18:N"+IntToStr(n-1)).c_str()).OlePropertyGet("Interior").OlePropertySet("Color",14408946);
   */
      //рисуем сетку
      Sh.OlePropertyGet("Range",("A2:J"+IntToStr(n-1)).c_str()).OlePropertyGet("Borders").OlePropertySet("LineStyle",1);

      //Сохранить книгу в папке в файле по указанию
     // AnsiString vAsCurDir1=WorkPath+"\\Формирование списка по предприятию";

     // Sh.OleProcedure("SaveAs",vAsCurDir1.c_str());
     AppEx.OlePropertyGet("WorkBooks",1).OleFunction("Save");

      /* //Закрыть открытое приложение Excel
      AppEx.OleProcedure("Quit");
      AppEx = Unassigned;  */

      //Закрыть книгу Excel с шаблоном для вывода информации
     // AppEx.OlePropertyGet("WorkBooks",1).OleProcedure("Close");
      Application->MessageBox("Отчет успешно сформирован!", "Формирование отчета",
                               MB_OK+MB_ICONINFORMATION);
      //AppEx.OlePropertyGet("WorkBooks").OleProcedure("Open",vAsCurDir1.c_str());
      AppEx.OlePropertySet("Visible",true);
      AppEx.OlePropertySet("AskToUpdateLinks",true);
      AppEx.OlePropertySet("DisplayAlerts",true);

      StatusBar1->SimpleText= "Формирование отчета выполнено.";

      Cursor = crDefault;
      ProgressBar->Position=0;
      ProgressBar->Visible = false;
      StatusBar1->SimpleText= "";

      if (valuta==1) InsertLog("Выполнено формирование файла по гривневой страховке для САП в Excel. Количество записей = "+ IntToStr(DM->qObnovlenie->RecordCount));
      else if (valuta==2) InsertLog("Выполнено формирование файла по валютной страховке для САП в Excel. Количество записей = "+ IntToStr(DM->qObnovlenie->RecordCount));
      else if (valuta==3) InsertLog("Выполнено формирование файла по внешним договорам для САП в Excel. Количество записей = "+ IntToStr(DM->qObnovlenie->RecordCount));
      else if (valuta==4) InsertLog("Выполнено формирование файла по пенсионным договорам для САП в Excel. Количество записей = "+ IntToStr(DM->qObnovlenie->RecordCount));

    }
  catch(...)
    {
      AppEx.OleProcedure("Quit");
      AppEx = Unassigned;
      Cursor = crDefault;
      ProgressBar->Position=0;
      ProgressBar->Visible = false;

      StatusBar1->SimpleText= "";
      DecimalSeparator='.';
      if (valuta==1) InsertLog("Возникла ошибка при формировании файла по гривневой страховке для САП в Excel");
      else if (valuta==2) InsertLog("Возникла ошибка при формировании файла по валютной страховке для САП в Excel");
      else if (valuta==3) InsertLog("Возникла ошибка при формировании файла по внешним договорам для САП в Excel");
      else if (valuta==4) InsertLog("Возникла ошибка при формировании файла по пенсионным договорам для САП в Excel");

      Abort();
    }

 // if (otchet_zex==0) InsertLog("Формирование списка работников по предприятию в Excel успешно завершено");
//  else InsertLog("Формирование списка работников по  "+otchet_zex+" цеху в Excel успешно завершено");
  DecimalSeparator='.';

}
//---------------------------------------------------------------------------



//Формирование шаблона для SAP по гривне
void __fastcall TMain::N5Click(TObject *Sender)
{
  ExcelSAP(1);
}
//---------------------------------------------------------------------------

//Формирование шаблона для SAP по валюте
void __fastcall TMain::N8Click(TObject *Sender)
{
  ExcelSAP(2);
}
//---------------------------------------------------------------------------

//Формирование шаблона для SAP по внешним договорам
void __fastcall TMain::N27Click(TObject *Sender)
{
  ExcelSAP(3);
}
//---------------------------------------------------------------------------

void __fastcall TMain::N28Click(TObject *Sender)
{
  OtchetStrahovaya(1);
}
//---------------------------------------------------------------------------

void __fastcall TMain::N29Click(TObject *Sender)
{
  OtchetStrahovaya(2);
}
//---------------------------------------------------------------------------

void __fastcall TMain::N30Click(TObject *Sender)
{
  OtchetStrahovaya(3);
}
//---------------------------------------------------------------------------

//Формирование файлов для страховой из SAP
void __fastcall TMain::OtchetStrahovaya(int valuta)
{
  AnsiString sFile, Sql, str,s;
  int otchet=0, kolzap=0, kolnzap=0;
  FILE *grn;

  //Выбор файла
  if(OpenDialog1->Execute())
    {
      // устанавливаем путь к файлу шаблона
      sFile = OpenDialog1->FileName;
    }
  else
    {
      Abort();
    }

// инициализируем Excel, открываем этот шаблон
  try
    {
      AppEx=CreateOleObject("Excel.Application");
    }
  catch (...)
    {
      Application->MessageBox("Невозможно открыть Microsoft Excel!"
                              " Возможно это приложение на компьютере не установлено.","Ошибка",MB_OK+MB_ICONERROR);
      StatusBar1->SimpleText="";
      InsertLog("Возникла ошибка при формировании текстового файла для страховой компании");
      ProgressBar->Visible = false;
      Cursor = crDefault;
    }

  Cursor = crHourGlass;
  ProgressBar->Position = 0;
  ProgressBar->Visible = true;
  StatusBar1->SimpleText = "Идет формирование файлов для страховой компании...";

  //Если возникает ошибка во время формирования отчета
  try
    {
      try
        {
          AppEx.OlePropertySet("AskToUpdateLinks",false);
          AppEx.OlePropertySet("DisplayAlerts",false);

          //Создание папки, если ее не существует
          ForceDirectories(WorkPath+"\\Для страховой компании");
          AppEx.OlePropertyGet("WorkBooks").OleProcedure("Open",sFile.c_str())    ;  //открываем книгу, указав её имя

          Sh=AppEx.OlePropertyGet("WorkSheets",1);                               //выбираем № активного листа книги
          //Sh=AppEx.OlePropertyGet("WorkSheets","Расчет");                      //выбираем лист по наименованию
        }
      catch(...)
        {
          Application->MessageBox("Ошибка открытия книги Microsoft Excel!","Ошибка",MB_OK+MB_ICONERROR);
          StatusBar1->SimpleText="";
          InsertLog("Возникла ошибка при формировании текстового файла для страховой компании");
          ProgressBar->Visible = false;
          Cursor = crDefault;
        }


      //Количество строк в документе
      int row = Sh.OlePropertyGet("UsedRange").OlePropertyGet("Rows").OlePropertyGet("Count");

      ProgressBar->Max=row;


      int i=1;

      if (valuta==1)
        {
          if ((grn=fopen((WorkPath+"\\Для страховой компании\\grivna.txt").c_str(),"wt"))==NULL)
            {
              ShowMessage("Файл не удается открыть");
              return;
            }
        }
      else if (valuta==2)
        {
          if ((grn=fopen((WorkPath+"\\Для страховой компании\\valuta.txt").c_str(),"wt"))==NULL)
            {
              ShowMessage("Файл не удается открыть");
              return;
            }
        }
      else if (valuta==3)
        {
          if ((grn=fopen((WorkPath+"\\Для страховой компании\\vneshnie.txt").c_str(),"wt"))==NULL)
            {
              ShowMessage("Файл не удается открыть");
              return;
            }
        }

      //Номер строки, с которой начинать вывод
      while (String(AppEx.OlePropertyGet("Range", ("A"+IntToStr(i)).c_str())).IsEmpty() || !Proverka(String(AppEx.OlePropertyGet("Range", ("A"+IntToStr(i)).c_str()))))
        {
          i++;
          ProgressBar->Position++;

          if (i==row)
            {
              Application->MessageBox("Таб.№ работника должен находиться в колонке А файла Excel \nна первой странице документа.\nВнесите изменения в Excel файл и повторите загрузку.\nЕсли ошибка будет возникать и в дальнейшем \nобратитесь к разработчику","Предупреждение",
                                      MB_OK+MB_ICONWARNING);
              Abort();
            }
        }

      //вывод в файл
      while(!String(AppEx.OlePropertyGet("Range", ("A"+IntToStr(i)).c_str())).IsEmpty() && Proverka(String(AppEx.OlePropertyGet("Range", ("A"+IntToStr(i)).c_str()))))
        {

          //Проверка на соответствие формируемого файла и вида оплат в документе
          if ((valuta==1 && String(AppEx.OlePropertyGet("Range", ("D"+IntToStr(i)).c_str()))!="7428")||
               (valuta==2 && String(AppEx.OlePropertyGet("Range", ("D"+IntToStr(i)).c_str()))!="7433") ||
               (valuta==3 && String(AppEx.OlePropertyGet("Range", ("D"+IntToStr(i)).c_str()))!="7434"))
            {
              if (valuta==1) s="гривневого договора";
              else if (valuta==2) s="валютного договора";
              else if (valuta==3) s="внешнего договора";

              Application->MessageBox(("В загружаемом документе Excel значение вида оплаты в колонке 'D'="+String(AppEx.OlePropertyGet("Range", ("D"+IntToStr(i)).c_str()))+" \nне является видом оплаты "+s+".\nВозможно неверно выбран файл Excel для загрузки.\nВыберите другой файл и повторите формирование отчета.").c_str(), "Ошибка",
                                      MB_OK+MB_ICONWARNING);
              //Закрыть открытое приложение Excel
              AppEx.OleProcedure("Quit");
              StatusBar1->SimpleText="";
              ProgressBar->Visible = false;
              Cursor = crDefault;
              Abort();

            }

          //Выборка данных по таб.№ из работающих
          Sql="select tn_sap, zex, initcap(fam||' '||im||' '||ot) as fio, numident \
               from sap_osn_sved where tn_sap="+String(AppEx.OlePropertyGet("Range", ("A"+IntToStr(i)).c_str()));

          DM->qObnovlenie->Close();
          DM->qObnovlenie->SQL->Clear();
          DM->qObnovlenie->SQL->Add(Sql);
          try
            {
              DM->qObnovlenie->Open();
            }
          catch(Exception &E)
            {
              Application->MessageBox(("Возникла ошибка при получении информации по работнику из кадров (SAP_OSN_SVED).\nНеобходимо повторить формирование! "+E.Message).c_str(), "Ошибка",
                                      MB_OK+MB_ICONERROR);
              //Закрыть открытое приложение Excel
              AppEx.OleProcedure("Quit");
              InsertLog("Возникла ошибка при формировании текстового файла для страховой компании");
              StatusBar1->SimpleText="";
              ProgressBar->Visible = false;
              Cursor = crDefault;
              Abort();
            }

          if (DM->qObnovlenie->RecordCount==0)
            {
              //Выборка данных из уволенных
              Sql="select tn_sap, zex, initcap(fam||' '||im||' '||ot) as fio, numident \
                   from sap_sved_uvol where tn_sap="+String(AppEx.OlePropertyGet("Range", ("A"+IntToStr(i)).c_str()));

              DM->qObnovlenie->Close();
              DM->qObnovlenie->SQL->Clear();
              DM->qObnovlenie->SQL->Add(Sql);
              try
                {
                  DM->qObnovlenie->Open();
                }
              catch(Exception &E)
                {
                  Application->MessageBox(("Возникла ошибка при получении информации по работнику из кадров (SAP_OSN_SVED).\nНеобходимо повторить формирование! "+E.Message).c_str(), "Ошибка",
                                           MB_OK+MB_ICONERROR);

                  //Закрыть открытое приложение Excel
                  AppEx.OleProcedure("Quit");
                  InsertLog("Возникла ошибка при формировании текстового файла для страховой компании");
                  StatusBar1->SimpleText="";
                  ProgressBar->Visible = false;
                  Cursor = crDefault;
                  Abort();
                }

              float sum=0;
              if (DM->qObnovlenie->RecordCount==0)
                {
                  //Нет записей, формирование отчета
                  if (otchet==0)
                    {
                      //Открытие файла данных содержащего уволенных, окончанием выплат и измененным цехом и таб.№
                      if (!rtf_Open((TempPath + "\\otchet2.txt").c_str()))
                        {
                          MessageBox(Handle,"Ошибка открытия файла данных","Ошибка",8192);
                        }
                      // Вывод заголовка и шапки таблицы
                      rtf_Out("data", Now(), 0);
                    }

                  rtf_Out("tn", String(AppEx.OlePropertyGet("Range", ("A"+IntToStr(i)).c_str())),1);
                  rtf_Out("sum", FloatToStrF(-1*Double(AppEx.OlePropertyGet("Range", ("E"+IntToStr(i)).c_str())),ffFixed,10,2) ,1);

                  if(!rtf_LineFeed())
                    {
                      MessageBox(Handle,"Ошибка записи в файл данных","Ошибка",8192);
                      if (!rtf_Close()) MessageBox(Handle,"Ошибка закрытия файла данных","Ошибка",8192);
                      return;
                    }

                  otchet=1;
                  i++;
                  kolnzap++;
                  ProgressBar->Position++;
                  continue;
                }
            }

          //вывод строки
          fprintf(grn,"\n%5s|%8s|%-37s|%8s|%10s|",DM->qObnovlenie->FieldByName("zex")->AsString,
                                                  DM->qObnovlenie->FieldByName("tn_sap")->AsString,
                                                  DM->qObnovlenie->FieldByName("fio")->AsString,
                                                  FloatToStrF(-1*Double(AppEx.OlePropertyGet("Range", ("E"+IntToStr(i)).c_str())), ffFixed, 20,2),
                                                  DM->qObnovlenie->FieldByName("numident")->AsString);

          i++;
          kolzap++;
          ProgressBar->Position++;
        }

      fclose(grn);


      if (otchet==1)
        {
          StatusBar1->SimpleText = "Идет формирование отчета с ошибками...";
          if(!rtf_Close())
            {
              MessageBox(Handle,"Ошибка закрытия файла данных", "Ошибка", 8192);
              return;
            }
          int istrd;
          try
            {
              rtf_CreateReport(TempPath +"\\otchet2.txt", Path+"\\RTF\\otchet2.rtf",
                           WorkPath+"\\Выгрузка данных для страховой.doc",NULL,&istrd);


              WinExec(("\""+ WordPath+"\"\""+WorkPath+"\\Выгрузка данных для страховой.doc\"").c_str(),SW_MAXIMIZE);

            }
          catch(RepoRTF_Error E)
            {
              MessageBox(Handle,("Ошибка формирования отчета:"+ AnsiString(E.Err)+
                                 "\nСтрока файла данных:"+IntToStr(istrd)).c_str(),"Ошибка",8192);
            }

           StatusBar1->SimpleText = "";
        }
    }
  catch(Exception &E)
    {
      Application->MessageBox(("Возникла ошибка при формировании файла для страховой компании"+E.Message).c_str(),"Ошибка",
                              MB_OK+MB_ICONERROR);

      //Отключить вывод сообщений с вопросами типа "Заменить файл..."
      AppEx.OlePropertySet("DisplayAlerts",false);
      //Закрыть открытое приложение Excel
      AppEx.OleProcedure("Quit");
      InsertLog("Возникла ошибка при формировании текстового файла для страховой компании");

      StatusBar1->SimpleText="";
      ProgressBar->Visible = false;
      Cursor = crDefault;
      Abort();
    }

   //Отключить вывод сообщений с вопросами типа "Заменить файл..."
  AppEx.OlePropertySet("DisplayAlerts",false);

  //Закрыть открытое приложение Excel
  AppEx.OleProcedure("Quit");

  if (kolnzap>0) str="Не загружено записей - "+IntToStr(kolnzap);
  else str="";

  Application->MessageBox(("Формирование выполнено успешно =)\nФайл для страховой компании находится в папке: "+WorkPath+"\\fkiev.txt.\nДобавлено записей в файл - "+kolzap+". "+str).c_str(),"Формирование распоряжений",
                          MB_OK+MB_ICONINFORMATION);

  InsertLog("Выполнено формирование текстового файла для страховой компании.\n Добавлено записей в файл - "+IntToStr(kolzap)+". "+str);
  StatusBar1->SimpleText="";
  ProgressBar->Visible = false;
  Cursor = crDefault;
}
//---------------------------------------------------------------------------

//Загрузка данных по новым договорам по пенсионному страхованию
void __fastcall TMain::N31Click(TObject *Sender)
{
  AnsiString Sql, Sql1, inn, nn;
  int i=1, rec=0;

  /*rec - количество вставленных в таблицу записей*/


  im_fl=7;

  if (Application->MessageBox(("Вы действительно хотите загрузить данные \n по новым договорам по пенсионному страхованию за " + Mes[DM->mm-1] + " " + DM->yyyy + " года?").c_str(),
                               "Загрузка данных по новым договорам",
                               MB_YESNO + MB_ICONINFORMATION) == IDNO)
    {
      Abort();
    }


  // Проверка правильности ИНН и наличие двойных записей в картотеке
  ProverkaInfoExcel();

  StatusBar1->SimpleText = "";

  try
    {
      Sheet.OleProcedure("Activate");

      Main->Cursor = crHourGlass;
      StatusBar1->SimplePanel = true;    // 2 панели на StatusBar1
      StatusBar1->SimpleText=" Идет загрузка данных...";

      ProgressBar->Visible = true;
      ProgressBar->Position = 0;
      ProgressBar->Max = Row;

      for ( i ; i<Row+1; i++)
        {
          nn = Excel.OlePropertyGet("Cells",i,1);
          inn = Excel.OlePropertyGet("Cells",i,5);

          ProgressBar->Position++;

          // Выбор строк необходимых для загрузки из Excel
          if (nn.IsEmpty() || !Proverka(nn) || inn.IsEmpty())  continue;

             //Проверка на наличие уже существующих записей в таблице VU_859_N
            Sql1 = "select * from VU_859_N where trim(inn)=trim("+ QuotedStr(Excel.OlePropertyGet("Cells",i,5)) +") \
                                           and trim(n_dogovora) = trim("+ QuotedStr(Excel.OlePropertyGet("Cells",i,9))+")" ;

            try
              {
                DM->qObnovlenie->Close();
                DM->qObnovlenie->SQL->Clear();
                DM->qObnovlenie->SQL->Add(Sql1);
                DM->qObnovlenie->Open();
              }
            catch(...)
              {
                Application->MessageBox("Ошибка получения данных из таблицы по страхованию 859 в/у","Ошибка",MB_OK+ MB_ICONERROR);
                Abort();
              }

            if (DM->qObnovlenie->RecordCount>0)
              {
                 if (Application->MessageBox(("Запись: цех = "+ DM->qObnovlenie->FieldByName("zex")->AsString +
                                               ", таб.№ = "+ DM->qObnovlenie->FieldByName("tn")->AsString +
                                               ", ИНН = "+ DM->qObnovlenie->FieldByName("inn")->AsString +
                                               " и № договора = "+DM->qObnovlenie->FieldByName("n_dogovora")->AsString +
                                              " уже существует. Записать ее еще раз?").c_str(),"Предупреждение",
                                              MB_YESNO + MB_ICONINFORMATION) ==ID_NO)
                    {
                       continue;
                    }
              }

            //Добавление цех+тн из sap_osn_sved
            Sql1="select zex, tn_sap, numident from sap_osn_sved where trim(numident)=trim("+ QuotedStr(Excel.OlePropertyGet("Cells",i,5)) +" )   \
                  union all                                                                                            \
                  select zex, tn_sap, numident from sap_sved_uvol                                                           \
                  where substr(to_char(dat_job,'dd.mm.yyyy'),4,7)='"+(DM->mm<10 ? "0"+IntToStr(DM->mm) : IntToStr(DM->mm))+"."+DM->yyyy+"'  \
                  and trim(numident)=trim("+QuotedStr(Excel.OlePropertyGet("Cells",i,5))+")";

           //  decode(translate('   123455','-0123456789 ','-'),null, '=p','=)')

            try
              {
                DM->qObnovlenie->Close();
                DM->qObnovlenie->SQL->Clear();
                DM->qObnovlenie->SQL->Add(Sql1);
                DM->qObnovlenie->Open();
              }
            catch(...)
              {
                Application->MessageBox("Ошибка получения данных из из картотеки по работникам (SAP_OSN_SVED, SAP_SVED_UVOL)","Ошибка",MB_OK+ MB_ICONERROR);
                Abort();
              }


            //Запись данных в таблицу VU_859_N
            Sql = "insert into vu_859_N (zex, tn, fio, n_dogovora, kod_dogovora, data_s, n_ind_schet, sum, inn, priznak) \
                   values("+ QuotedStr(DM->qObnovlenie->FieldByName("zex")->AsString)+", \
                          "+ SetNull(DM->qObnovlenie->FieldByName("tn_sap")->AsString)+", \
                          initcap("+ QuotedStr(Excel.OlePropertyGet("Cells",i,2))+"||' '||"+QuotedStr(Excel.OlePropertyGet("Cells",i,3))+"||' '||"+QuotedStr(Excel.OlePropertyGet("Cells",i,4))+"), \
                          trim("+ QuotedStr(Excel.OlePropertyGet("Cells",i,9))+"), \
                             4, \
                          "+ QuotedStr(Excel.OlePropertyGet("Cells",i,6))+", \
                          trim("+ QuotedStr(Excel.OlePropertyGet("Cells",i,10))+"), \
                          "+ QuotedStr(Excel.OlePropertyGet("Cells",i,8))+", \
                          trim("+ QuotedStr(Excel.OlePropertyGet("Cells",i,5))+"),\
                             0 ) ";
            try
              {
                DM->qZagruzka->Close();
                DM->qZagruzka->SQL->Clear();
                DM->qZagruzka->SQL->Add(Sql);
                DM->qZagruzka->ExecSQL();
                rec++;
              }
            catch(...)
              {
                Application->MessageBox("Ошибка вставки данных в таблицу по страхованию 859 в/у","Ошибка",MB_OK+ MB_ICONERROR);
                Application->MessageBox("Данные не были загружены. Повторите загрузку","Ошибка",MB_OK+ MB_ICONERROR);
                StatusBar1->SimpleText = "";

                Excel.OleProcedure("Quit");
                Abort();
             }
        }


      Application->MessageBox(("Загрузка данных выполнена успешно =) \n Добавлено " + IntToStr(rec) + " записей").c_str(),
                               "Загрузка новых договоров по пенсионному страхованию",MB_OK+ MB_ICONINFORMATION);
      InsertLog("Выполнена загрузка данных по новым договорам по пенсионному страхованию. Загружено "+IntToStr(rec)+" записей");

      Excel.OleProcedure("Quit");
      Excel = Unassigned;

      StatusBar1->SimplePanel = false;
      ProgressBar->Visible = false;
      StatusBar1->SimpleText = "Обновление выполнено.";
      Main->Cursor = crDefault;
      StatusBar1->SimpleText = "";
    }
  catch(...)
    {
      Application->MessageBox("Ошибка загрузки данных по новым договорам по пенсионному страхованию","Ошибка",MB_OK+ MB_ICONERROR);
      Excel.OleProcedure("Quit");

      StatusBar1->SimplePanel = false;
      ProgressBar->Visible = false;
      StatusBar1->SimpleText = "";
      Main->Cursor = crDefault;
    }
}
//---------------------------------------------------------------------------

//Загрузка изменений по пенсионному страхованию
void __fastcall TMain::N32Click(TObject *Sender)
{
  im_fl=7;
  
  if (Application->MessageBox(("Вы действительно хотите загрузить изменения \n по договорам пенсионного страхования за " + Mes[DM->mm-1] + " " + DM->yyyy + " года?").c_str(),
                               "Загрузка изменений по гривневым договорам",
                               MB_YESNO + MB_ICONINFORMATION) == IDNO)
    {
      Abort();
    }

  // Проверка на существование инн в таблице
  ProverkaInfoExcelIzmeneniya();

  StatusBar1->SimpleText = "";

  //Обновление изменений по гривневой страховке
  UpdateValuta_I_Grivna();

  InsertLog("Выполнена загрузка изменений по пенсионному страхованию. Обновлено "+obnov_kol+" из "+ob_kol+" записей");

  StatusBar1->SimpleText = "";
}
//---------------------------------------------------------------------------



//Формирование отчета для SAP по пенсионному страхованию
void __fastcall TMain::N34Click(TObject *Sender)
{
  ExcelSAP(4);
}
//---------------------------------------------------------------------------

//Формирование итогового отчета по страховой компании
void __fastcall TMain::N36Click(TObject *Sender)
{

  AnsiString sFile, sFile1, Sql, str,s;
  int otchet=0, kolzap=0, kolnzap=0, n=5, num=1;

  //Выбор файла
  if(OpenDialog1->Execute())
    {
      // устанавливаем путь к файлу шаблона
      sFile = OpenDialog1->FileName;
    }
  else
    {
      Abort();
    }

// инициализируем Excel, открываем файл с данными из файла Excel из SAP
  try
    {
      AppEx=CreateOleObject("Excel.Application");
    }
  catch (...)
    {
      Application->MessageBox("Невозможно открыть Microsoft Excel!"
                              " Возможно это приложение на компьютере не установлено.","Ошибка",MB_OK+MB_ICONERROR);
      StatusBar1->SimpleText="";
      InsertLog("Возникла ошибка при формировании текстового файла для страховой компании");
      ProgressBar->Visible = false;
      Cursor = crDefault;
    }

  Cursor = crHourGlass;
  ProgressBar->Position = 0;
  ProgressBar->Visible = true;
  StatusBar1->SimpleText = "Идет формирование файлов для страховой компании...";

  //Если возникает ошибка во время формирования отчета
  try
    {
      try
        {
          AppEx.OlePropertySet("AskToUpdateLinks",false);
          AppEx.OlePropertySet("DisplayAlerts",false);

          //Создание папки, если ее не существует
          ForceDirectories(WorkPath+"\\Для страховой компании");
          AppEx.OlePropertyGet("WorkBooks").OleProcedure("Open",sFile.c_str())    ;  //открываем книгу, указав её имя

          Sh=AppEx.OlePropertyGet("WorkSheets",1);                               //выбираем № активного листа книги
          //Sh=AppEx.OlePropertyGet("WorkSheets","Расчет");                      //выбираем лист по наименованию
        }
      catch(...)
        {
          Application->MessageBox("Ошибка открытия книги Microsoft Excel!","Ошибка",MB_OK+MB_ICONERROR);
          StatusBar1->SimpleText="";
          InsertLog("Возникла ошибка при формировании текстового файла для страховой компании");
          ProgressBar->Visible = false;
          Cursor = crDefault;
        }


      //Количество строк в документе
      int row = Sh.OlePropertyGet("UsedRange").OlePropertyGet("Rows").OlePropertyGet("Count");

      ProgressBar->Max=row;

      int i=1;

      // инициализируем Excel, открываем файл итогового отчета, куда будем копировать данные из файла Excel из SAP
      try
         {
           AppEx1=CreateOleObject("Excel.Application");
         }
      catch (...)
         {
           Application->MessageBox("Невозможно открыть Microsoft Excel!"
                                   " Возможно это приложение на компьютере не установлено.","Ошибка",MB_OK+MB_ICONERROR);
           StatusBar1->SimpleText="";
           ProgressBar->Visible = false;
           Cursor = crDefault;
         }

      //Если возникает ошибка во время формирования отчета
      try
        {
          try
            {
              AppEx1.OlePropertySet("AskToUpdateLinks",false);
              AppEx1.OlePropertySet("DisplayAlerts",false);

              //Создание папки, если ее не существует
              ForceDirectories(WorkPath);

              DeleteFile(WorkPath+"\\Итоговый отчет по пенсионным договорам.xlsx");
              CopyFile((Path+"\\RTF\\itog_pens.xlsx").c_str(), (WorkPath+"\\Итоговый отчет по пенсионным договорам.xlsx").c_str(), false);
              sFile1 = WorkPath+"\\Итоговый отчет по пенсионным договорам.xlsx";


              AppEx1.OlePropertyGet("WorkBooks").OleProcedure("Open",sFile1.c_str())    ;  //открываем книгу, указав её имя

              Sh=AppEx1.OlePropertyGet("WorkSheets",1);                               //выбираем № активного листа книги
            }
          catch(...)
            {
              Application->MessageBox("Ошибка открытия книги Microsoft Excel!","Ошибка",MB_OK+MB_ICONERROR);
              StatusBar1->SimpleText="";
              ProgressBar->Visible = false;
              Cursor = crDefault;
              DecimalSeparator='.';
              InsertLog("Возникла ошибка при формировании итогового отчета по пенсионной страховке в Excel");
            }


      //Номер строки, с которой начинать вывод данных в файле SAP
      while (String(AppEx.OlePropertyGet("Range", ("A"+IntToStr(i)).c_str())).IsEmpty() || !Proverka(String(AppEx.OlePropertyGet("Range", ("A"+IntToStr(i)).c_str()))))
        {
          i++;
          ProgressBar->Position++;

          if (i==row+1)
            {
              Application->MessageBox("Таб.№ работника должен находиться в колонке А файла Excel \nна первой странице документа.\nВнесите изменения в Excel файл и повторите загрузку.\nЕсли ошибка будет возникать и в дальнейшем \nобратитесь к разработчику","Предупреждение",
                                      MB_OK+MB_ICONWARNING);
              Abort();
            }
        }



      Variant Massiv;
      Massiv = VarArrayCreate(OPENARRAY(int,(0,8)),varVariant); //массив на 11 элементов


     // AppEx1.OlePropertySet("Visible",true);

      //вывод в файл
      while(!String(AppEx.OlePropertyGet("Range", ("A"+IntToStr(i)).c_str())).IsEmpty() && Proverka(String(AppEx.OlePropertyGet("Range", ("A"+IntToStr(i)).c_str()))))
        {

          //Проверка на соответствие формируемого файла и вида оплат в документе
          if (String(AppEx.OlePropertyGet("Range", ("D"+IntToStr(i)).c_str()))!="7428")
            {
              s="пенсионного договора";


              Application->MessageBox(("В загружаемом документе Excel значение вида оплаты в колонке 'D'="+String(AppEx.OlePropertyGet("Range", ("D"+IntToStr(i)).c_str()))+" \nне является видом оплаты "+s+".\nВозможно неверно выбран файл Excel для загрузки.\nВыберите другой файл и повторите формирование отчета.").c_str(), "Ошибка",
                                      MB_OK+MB_ICONWARNING);
              //Закрыть открытое приложение Excel
              AppEx.OleProcedure("Quit");
              StatusBar1->SimpleText="";
              ProgressBar->Visible = false;
              Cursor = crDefault;
              Abort();

            }

          //Выборка данных по таб.№ из работающих и уволенных и добавление номера договора из таблицы по страхованию
          Sql = "select k.numident,        \
                        v.n_dogovora,      \
                        v.n_ind_schet,   \
                        k.fam_ukr,             \
                        k.im_ukr,              \
                        k.ot_ukr               \
                 from                      \
                     (select tn_sap, zex, fam_ukr, im_ukr, ot_ukr, numident     \
                      from sap_osn_sved where tn_sap="+String(AppEx.OlePropertyGet("Range", ("A"+IntToStr(i)).c_str()))+"   \
                      union all                                                               \
                      select tn_sap, zex, fam_ukr, im_ukr, ot_ukr, numident     \
                      from sap_sved_uvol where tn_sap="+String(AppEx.OlePropertyGet("Range", ("A"+IntToStr(i)).c_str()))+"  \
                     ) k                                                                      \
                 left join vu_859_n v                                                         \
                 on k.numident=inn                                                            \
                 and v.kod_dogovora=4                                                         \
                 and nvl(v.priznak,0)=0";

          DM->qObnovlenie->Close();
          DM->qObnovlenie->SQL->Clear();
          DM->qObnovlenie->SQL->Add(Sql);
          try
            {
              DM->qObnovlenie->Open();
            }
          catch(Exception &E)
            {
              Application->MessageBox(("Возникла ошибка при получении информации по работнику из кадров (SAP_OSN_SVED).\nНеобходимо повторить формирование! "+E.Message).c_str(), "Ошибка",
                                      MB_OK+MB_ICONERROR);
              //Закрыть открытое приложение Excel
              AppEx.OleProcedure("Quit");
              InsertLog("Возникла ошибка при формировании отчета по пенсионному страхованию");
              StatusBar1->SimpleText="";
              ProgressBar->Visible = false;
              Cursor = crDefault;
              Abort();
            }

          if (DM->qObnovlenie->RecordCount==0)
            {
              //Нет записей, формирование отчета
              if (otchet==0)
                {
                  //Открытие файла данных содержащего уволенных, окончанием выплат и измененным цехом и таб.№
                  if (!rtf_Open((TempPath + "\\otchet2.txt").c_str()))
                    {
                      MessageBox(Handle,"Ошибка открытия файла данных","Ошибка",8192);
                    }
                  // Вывод заголовка и шапки таблицы
                  rtf_Out("data", Now(), 0);
                }

              rtf_Out("tn", String(AppEx.OlePropertyGet("Range", ("A"+IntToStr(i)).c_str())),1);
              rtf_Out("sum", FloatToStrF(-1*Double(AppEx.OlePropertyGet("Range", ("E"+IntToStr(i)).c_str())),ffFixed,10,2) ,1);

              if(!rtf_LineFeed())
                {
                  MessageBox(Handle,"Ошибка записи в файл данных","Ошибка",8192);
                  if (!rtf_Close()) MessageBox(Handle,"Ошибка закрытия файла данных","Ошибка",8192);
                  return;
                }

              otchet=1;
              i++;
              kolnzap++;
              ProgressBar->Position++;
              continue;
            }

          //вывод даты формирования отчета
          Sh.OlePropertyGet("Range", "E2").OlePropertySet("Value", ("за "+Mes[DM->mm-1] + " " + DM->yyyy+" года").c_str());


          //вывод строки
          Massiv.PutElement(num, 0);
          Massiv.PutElement(DM->qObnovlenie->FieldByName("numident")->AsString.c_str(), 1);
          Massiv.PutElement(DM->qObnovlenie->FieldByName("n_dogovora")->AsString.c_str(), 2);
          Massiv.PutElement(DM->qObnovlenie->FieldByName("n_ind_schet")->AsString.c_str(), 3);
          Massiv.PutElement(DM->qObnovlenie->FieldByName("fam_ukr")->AsString.c_str(), 4);
          Massiv.PutElement(DM->qObnovlenie->FieldByName("im_ukr")->AsString.c_str(), 5);
          Massiv.PutElement(DM->qObnovlenie->FieldByName("ot_ukr")->AsString.c_str(), 6);
          //Massiv.PutElement(FloatToStrF(Double(AppEx.OlePropertyGet("Range", ("E"+IntToStr(i)).c_str())), ffFixed, 20,2), 7);
          Massiv.PutElement(-1*Double(AppEx.OlePropertyGet("Range", ("E"+IntToStr(i)).c_str())), 7);

          Sh.OlePropertyGet("Range", ("A" + IntToStr(n) + ":H" + IntToStr(n)).c_str()).OlePropertySet("Value", Massiv); //строка с данными с ячейки A по ячейку АВ

          num++;
          i++;
          n++;
          kolzap++;
          ProgressBar->Position++;
        }


       //вывот строки с итогами
       Sh.OlePropertyGet("Range", ("A" + IntToStr(n)).c_str()).OlePropertySet("Value", "Загальна сума по відомості (грн):");
       Sh.OlePropertyGet("Range", ("H" + IntToStr(n)).c_str()).OlePropertySet("Formula", ("=СУММ(H5:H"+IntToStr(n-1)).c_str());


       //.OlePropertyGet("Offset", n)
       //жирный шрифт
       Sh.OlePropertyGet("Range",("A"+IntToStr(n)).c_str()).OlePropertyGet("Font").OlePropertySet("Bold",true);
       Sh.OlePropertyGet("Range",("H"+IntToStr(n)).c_str()).OlePropertyGet("Font").OlePropertySet("Bold",true);

       //переносить текст
       // Sh.OlePropertyGet("Range",("A"+IntToStr(n)).c_str()).OlePropertySet("WrapText",true);

       //выравнивание по горизонтали
       Sh.OlePropertyGet("Range",("A"+IntToStr(n)).c_str()).OlePropertySet("HorizontalAlignment", 1); //выровнять по гор.

       //сетка
       Sh.OlePropertyGet("Range",("A5:H"+IntToStr(n)).c_str()).OlePropertyGet("Borders").OlePropertySet("LineStyle", 1);


      if (otchet==1)
        {
          StatusBar1->SimpleText = "Идет формирование отчета с ошибками...";
          if(!rtf_Close())
            {
              MessageBox(Handle,"Ошибка закрытия файла данных", "Ошибка", 8192);
              return;
            }
          int istrd;
          try
            {
              rtf_CreateReport(TempPath +"\\otchet2.txt", Path+"\\RTF\\otchet2.rtf",
                           WorkPath+"\\Ошибки при формировании итогового отчета по пенсионным договорам.doc",NULL,&istrd);


              WinExec(("\""+ WordPath+"\"\""+WorkPath+"\\Ошибки при формировании итогового отчета по пенсионным договорам.doc\"").c_str(),SW_MAXIMIZE);

            }
          catch(RepoRTF_Error E)
            {
              MessageBox(Handle,("Ошибка формирования отчета:"+ AnsiString(E.Err)+
                                 "\nСтрока файла данных:"+IntToStr(istrd)).c_str(),"Ошибка",8192);
            }

           StatusBar1->SimpleText = "";
        }

        }
  catch(Exception &E)
    {
      Application->MessageBox(("Возникла ошибка при формировании итогового отчета по пенсионным договорам страхования"+E.Message).c_str(),"Ошибка",
                              MB_OK+MB_ICONERROR);

      //Отключить вывод сообщений с вопросами типа "Заменить файл..."
      AppEx1.OlePropertySet("DisplayAlerts",false);
      //Закрыть открытое приложение Excel
      AppEx1.OleProcedure("Quit");
      InsertLog("Возникла ошибка при формировании итогового отчета по пенсионным договорам страхования");

      StatusBar1->SimpleText="";
      ProgressBar->Visible = false;
      Cursor = crDefault;
      Abort();
    }



    }
  catch(Exception &E)
    {
      Application->MessageBox(("Возникла ошибка при формировании итогового отчета по пенсионным договорам страхования"+E.Message).c_str(),"Ошибка",
                              MB_OK+MB_ICONERROR);

      //Отключить вывод сообщений с вопросами типа "Заменить файл..."
      AppEx.OlePropertySet("DisplayAlerts",false);
      //Закрыть открытое приложение Excel
      AppEx.OleProcedure("Quit");
      InsertLog("Возникла ошибка при формировании итогового отчета по пенсионным договорам страхования");

      StatusBar1->SimpleText="";
      ProgressBar->Visible = false;
      Cursor = crDefault;
      Abort();
    }

   //Отключить вывод сообщений с вопросами типа "Заменить файл..."
  AppEx.OlePropertySet("DisplayAlerts",false);
  //Закрыть открытое приложение Excel с файлом из SAP
  AppEx.OleProcedure("Quit");


   AppEx1.OlePropertyGet("WorkBooks",1).OleFunction("Save");
   AppEx1.OlePropertySet("Visible",true);
   AppEx1.OlePropertySet("AskToUpdateLinks",true);
   AppEx1.OlePropertySet("DisplayAlerts",true);



  if (kolnzap>0) str="Не загружено записей - "+IntToStr(kolnzap);
  else str="";

  Application->MessageBox("Формирование выполнено успешно =)","Формирование итогового отчета",
                          MB_OK+MB_ICONINFORMATION);

  InsertLog("Выполнено формирование итогового отчета по пенсионным договорам по страхованию.\n Количество записей - "+IntToStr(kolzap)+". "+str);
  StatusBar1->SimpleText="";
  ProgressBar->Visible = false;
  Cursor = crDefault;
}
//---------------------------------------------------------------------------

