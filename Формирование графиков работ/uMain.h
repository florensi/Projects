//---------------------------------------------------------------------------

#ifndef uMainH
#define uMainH
//---------------------------------------------------------------------------
#include <Classes.hpp>
#include <Controls.hpp>
#include <StdCtrls.hpp>
#include <Forms.hpp>
#include <Menus.hpp>
#include "DBGridEh.hpp"
#include <ExtCtrls.hpp>
#include <Grids.hpp>
#include <Buttons.hpp>


#include <DateUtils.hpp>
#include <Graphics.hpp>
#include <DBGrids.hpp>
#include <ComCtrls.hpp>
#include "Excel_2K_SRVR.h"

//---------------------------------------------------------------------------
class TMain : public TForm
{
__published:	// IDE-managed Components
        TMainMenu *MainMenu1;
        TPanel *Panel1;
        TPanel *Panel2;
        TComboBox *ComboBox1;
        TDBGridEh *DBGridEh1;
        TPopupMenu *PopupMenu1;
        TMenuItem *OtchetN1;
        TImage *Image1;
        TBevel *Bevel1;
        TMenuItem *N3Redaktirovat;
        TStatusBar *StatusBar1;
        TMenuItem *N5V_UIT;
        TMenuItem *N1;
        TMenuItem *N5RaschetOdin;
        TMenuItem *N6;
        TMenuItem *Word1;
        TMenuItem *Excel1;
        TMenuItem *Cghfdjxybrb1;
        TMenuItem *N7;
        TMenuItem *Cthdbc1;
        TMenuItem *N8;
        TMenuItem *N9;
        TMenuItem *N2;
        TMenuItem *N2RaschetVsex;
        TMenuItem *N10;
        TMenuItem *N3;
        TMenuItem *N4;
        TMenuItem *N5;
        void __fastcall FormCreate(TObject *Sender);
        void __fastcall DBGridEh1DrawColumnCell(TObject *Sender,
          const TRect &Rect, int DataCol, TColumnEh *Column,
          TGridDrawState State);
        void __fastcall DBGridEh1DblClick(TObject *Sender);
        void __fastcall N3RedaktirovatClick(TObject *Sender);
        void __fastcall DBGridEh1KeyPress(TObject *Sender, char &Key);
        void __fastcall FormKeyDown(TObject *Sender, WORD &Key,
          TShiftState Shift);
        void __fastcall DBGrid1DrawColumnCell(TObject *Sender,
          const TRect &Rect, int DataCol, TColumn *Column,
          TGridDrawState State);
        void __fastcall N5V_UITClick(TObject *Sender);
        void __fastcall N1Click(TObject *Sender);
        void __fastcall N5RaschetOdinClick(TObject *Sender);
        void __fastcall ComboBox1Click(TObject *Sender);
        void __fastcall Word1Click(TObject *Sender);
        void __fastcall Excel1Click(TObject *Sender);
        void __fastcall N7Click(TObject *Sender);
        void __fastcall N8Click(TObject *Sender);
        void __fastcall N9Click(TObject *Sender);
        void __fastcall N2RaschetVsexClick(TObject *Sender);
        void __fastcall N3Click(TObject *Sender);
        void __fastcall StatusBar1DblClick(TObject *Sender);
        void __fastcall N5Click(TObject *Sender);
        void __fastcall FormClose(TObject *Sender, TCloseAction &Action);
private:	// User declarations
        TProgressBar *ProgressBar;
public:		// User declarations

        TLocateOptions SearchOptions;

        AnsiString UserName, DomainName, UserFullName,
                    WordPath, Path, DocPath, TempPath, WorkPath;

        AnsiString numk;  //номер колонки(поля) для редактирования (например chf11, nch11...)

        int Prava;        //права для редактирования
        int nsm, dnism;   //nsm - номер смены последнего месяца,  dnism - день смены последнего месяца
        int day_mart, day_oktyabr;                             //день перехода на зимнее/летнее время
        int day_mart2, day_oktyabr2, mes_mart2, mes_oktyabr2;  //следующий день после перехода на зимнее/летнее время для первого дня первой смены 40 и 90 графика
        int god, den, mes;                                     //год из grafr
        int grafr;                                             //год из grafr, который не меняется
        int mes_n, mes_k;                                      // количество месяцев в графике для расчета и отображения
        int graf;                                              //расчитываемый график
        AnsiString n_grafik[150], n_grafikv[150];
        int kol_grafik;                         //количество и список редактируемых графиков
        int redakt;                                            //права на редактирование выбранного в ComboBox1 графика
        int br;                                                //текущий номер бригады
        int status;                                            //рассчитался ли график, если нет status=1

        AnsiString chf[13][32], vihod[13][32];                 //часы по дням и выходы по дням
        AnsiString vchf[13][32], pchf[13][32], nchf[13][32];   //вечерние, праздничные, ночные часы по дням
        AnsiString zchf, znsm, zpch, zvch, znch, zchf0, znch0, zpch0;  //для редактирования, сохранение предыдущих значений смены и часов по дате
        AnsiString n_graf;                                     //Список доступных графиков

        double ochf[13], ovchf[13], opchf[13], onchf[13], pgraf[13];  //общие суммы по часам факт, вечерние, праздничные, ночные, переработка.
        double d1, d2, d3, d4, d5, v, v1, v2, n, n1, n2, p1, p2, p3, p, r ;//длительность смены, вечерние, ночные, праздничные
        double chf0[13], pchf0[13], nchf0[13];  //переходящие часы с прошлого месяца

        bool __fastcall GetMyDocumentsDir(AnsiString &FolderPath);  //Путь к папке "Мои документы"
        bool __fastcall GetTempDir(AnsiString &FolderPath);         //Путь к временной папке
        AnsiString __fastcall FindWordPath();

         void __fastcall RaschetGraf(int graf, int year);       //функция расчета графика
         void __fastcall Graf11();                              //расчет 11 и 81 и 820 и 830 графика
         void __fastcall Graf23(double d1, double v, double n);  //расчет 23 графика
         void __fastcall Graf24(double d1, double v, double n);  //расчет 24 графика
         void __fastcall Graf25(double v, double n, double r);  //расчет 25 графика
         void __fastcall Graf40(double d1, double d2, double d3, double p1, double p2, double p, double n1, double n2, double n);  //расчет 40 графика
         void __fastcall Graf60(double p1, double p2, double v1, double v2, double n1, double n2);  //расчет 60 графика
         void __fastcall Graf4060(double p, double p1, double p2, double v1, double v2, double n1, double n2);  //расчет 4060 графика
         void __fastcall Graf85(double n);                      //расчет 85 графика
         void __fastcall Graf90(double d1, double d2, double d3, double p1, double p2, double p, double v, double n1, double n2, double n);  //расчет 90 графика
         void __fastcall Graf120(double d1, double d2, double d3, double p1, double p2, double v, double n1, double n2, double n);  //расчет 120 графика
         void __fastcall Graf133(double v, double n);           //расчет 133 графика
         void __fastcall Graf140(double d1, double v, double n); //расчет 140 графика
         void __fastcall Graf150();                             //расчет 150 графика
         void __fastcall Graf160(double d1, double v, double n); //расчет 160 графика
         void __fastcall Graf180(double d1, double p1, double p2, double v, double n1, double n2, double n);  //расчет 180 графика
         void __fastcall Graf190(double d1, double v1, double v2);         //расчет 190 графика
         void __fastcall Graf220(double v, double n);           //расчет 220 графика
         void __fastcall Graf225(double v1, double v2, double n);           //расчет 225 графика
         void __fastcall Graf230(double v);                     //расчет 230 графика
         void __fastcall Graf240(double v, double n1, double n2); //расчет 240 графика
         void __fastcall Graf250(double d1, double v1, double v2, double n);  //расчет 250 графика
         void __fastcall Graf260(double v, double n);           //расчет 260 графика
         void __fastcall Graf270(double d1, double v, double n);  //расчет 270 графика
         void __fastcall Graf280(double d1, double v);  //расчет 280 графика
         void __fastcall Graf300(double v);                     //расчет 300 и 131 графика / v - вечерние часы
         void __fastcall Graf315(double v);                     //расчет 315 графика / v - вечерние часы
         void __fastcall Graf320(double v, double n1, double n2, double p1, double p2);  //расчет 320 графика
         void __fastcall Graf370(double v, double n1, double n2, double p1, double p2);  //расчет 370 графика
         void __fastcall Graf390(double p1, double p2, double n1, double n2);         //расчет 390 и 950 графика /p1,p2 - праздничные дни
         void __fastcall Graf400(double v);                     //расчет 400 графика
         void __fastcall Graf410(double v);                     //расчет 410 графика
         void __fastcall Graf450(double v, double n1, double n2, double p1, double p2);  //расчет 450 графика
         void __fastcall Graf470(double v, double n);           //расчет 470 графика
         void __fastcall Graf480();                             //расчет 480 графика
         void __fastcall Graf520(double d1, double d2, double d3, double p1, double p2,        // Расчет 210(ХМФ) или 520 графика
                               double p3, double p, double v, double n1, double n2);
         void __fastcall Graf630();                             //расчет 630 графика
         void __fastcall Graf650();                             //расчет 650 и 660 графика
         void __fastcall Graf670(double v);                     //расчет 670 и 790 графика
         void __fastcall Graf680(double v);                     //расчет 680 графика
         void __fastcall Graf690(double v);                     //расчет 690 графика
         void __fastcall Graf771();                    //расчет 711 графика
         void __fastcall Graf775(double v, double n);           //расчет 775 графика
         void __fastcall Graf780();                             //расчет 780 графика
         void __fastcall Graf800(int day1, int day2);           //расчет 800 графика
         void __fastcall Graf850(double v);                     //расчет 850 графика / v - вечерние часы
         void __fastcall Graf855(double v);                     //расчет 855 графика 
         void __fastcall Graf865(double v);                     //расчет 865 графика
         void __fastcall Graf960(double d1, double d2, double d3, double d4,
                                 double d5, double v, double n);//расчет 960 графика /d1-d6 - длительность смены в зависимости от периода работы
         void __fastcall Graf980(double p1, double p2, double v, double n1, double n2); //расчет 980 графика
         



         int __fastcall DayWeek(int d, int m, int y);           //определение дня недели
         bool __fastcall PrazdDni(int d, int m);                //определение праздничного дня
         bool __fastcall PrdPrazdDni(int d, int m);             //предпраздничные дни
         bool __fastcall PrazdDniVihodnue(int d, int m, int y); //праздничные приходящиеся на выходные
         void __fastcall SetInfoEdit();                         //Заполнение Edit-ов
         void __fastcall NextMonth(int &Month, int &Year, int k=1);
         void __fastcall PrevMonth(int &Month, int &Year, int k=1);

         void __fastcall InsertLog(AnsiString Msg); //вставка записей в таблицу logs_ro


        __fastcall TMain(TComponent* Owner);
};
//---------------------------------------------------------------------------
extern PACKAGE TMain *Main;
//---------------------------------------------------------------------------
#endif
