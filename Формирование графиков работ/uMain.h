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

        AnsiString numk;  //����� �������(����) ��� �������������� (�������� chf11, nch11...)

        int Prava;        //����� ��� ��������������
        int nsm, dnism;   //nsm - ����� ����� ���������� ������,  dnism - ���� ����� ���������� ������
        int day_mart, day_oktyabr;                             //���� �������� �� ������/������ �����
        int day_mart2, day_oktyabr2, mes_mart2, mes_oktyabr2;  //��������� ���� ����� �������� �� ������/������ ����� ��� ������� ��� ������ ����� 40 � 90 �������
        int god, den, mes;                                     //��� �� grafr
        int grafr;                                             //��� �� grafr, ������� �� ��������
        int mes_n, mes_k;                                      // ���������� ������� � ������� ��� ������� � �����������
        int graf;                                              //������������� ������
        AnsiString n_grafik[150], n_grafikv[150];
        int kol_grafik;                         //���������� � ������ ������������� ��������
        int redakt;                                            //����� �� �������������� ���������� � ComboBox1 �������
        int br;                                                //������� ����� �������
        int status;                                            //����������� �� ������, ���� ��� status=1

        AnsiString chf[13][32], vihod[13][32];                 //���� �� ���� � ������ �� ����
        AnsiString vchf[13][32], pchf[13][32], nchf[13][32];   //��������, �����������, ������ ���� �� ����
        AnsiString zchf, znsm, zpch, zvch, znch, zchf0, znch0, zpch0;  //��� ��������������, ���������� ���������� �������� ����� � ����� �� ����
        AnsiString n_graf;                                     //������ ��������� ��������

        double ochf[13], ovchf[13], opchf[13], onchf[13], pgraf[13];  //����� ����� �� ����� ����, ��������, �����������, ������, �����������.
        double d1, d2, d3, d4, d5, v, v1, v2, n, n1, n2, p1, p2, p3, p, r ;//������������ �����, ��������, ������, �����������
        double chf0[13], pchf0[13], nchf0[13];  //����������� ���� � �������� ������

        bool __fastcall GetMyDocumentsDir(AnsiString &FolderPath);  //���� � ����� "��� ���������"
        bool __fastcall GetTempDir(AnsiString &FolderPath);         //���� � ��������� �����
        AnsiString __fastcall FindWordPath();

         void __fastcall RaschetGraf(int graf, int year);       //������� ������� �������
         void __fastcall Graf11();                              //������ 11 � 81 � 820 � 830 �������
         void __fastcall Graf23(double d1, double v, double n);  //������ 23 �������
         void __fastcall Graf24(double d1, double v, double n);  //������ 24 �������
         void __fastcall Graf25(double v, double n, double r);  //������ 25 �������
         void __fastcall Graf40(double d1, double d2, double d3, double p1, double p2, double p, double n1, double n2, double n);  //������ 40 �������
         void __fastcall Graf60(double p1, double p2, double v1, double v2, double n1, double n2);  //������ 60 �������
         void __fastcall Graf4060(double p, double p1, double p2, double v1, double v2, double n1, double n2);  //������ 4060 �������
         void __fastcall Graf85(double n);                      //������ 85 �������
         void __fastcall Graf90(double d1, double d2, double d3, double p1, double p2, double p, double v, double n1, double n2, double n);  //������ 90 �������
         void __fastcall Graf120(double d1, double d2, double d3, double p1, double p2, double v, double n1, double n2, double n);  //������ 120 �������
         void __fastcall Graf133(double v, double n);           //������ 133 �������
         void __fastcall Graf140(double d1, double v, double n); //������ 140 �������
         void __fastcall Graf150();                             //������ 150 �������
         void __fastcall Graf160(double d1, double v, double n); //������ 160 �������
         void __fastcall Graf180(double d1, double p1, double p2, double v, double n1, double n2, double n);  //������ 180 �������
         void __fastcall Graf190(double d1, double v1, double v2);         //������ 190 �������
         void __fastcall Graf220(double v, double n);           //������ 220 �������
         void __fastcall Graf225(double v1, double v2, double n);           //������ 225 �������
         void __fastcall Graf230(double v);                     //������ 230 �������
         void __fastcall Graf240(double v, double n1, double n2); //������ 240 �������
         void __fastcall Graf250(double d1, double v1, double v2, double n);  //������ 250 �������
         void __fastcall Graf260(double v, double n);           //������ 260 �������
         void __fastcall Graf270(double d1, double v, double n);  //������ 270 �������
         void __fastcall Graf280(double d1, double v);  //������ 280 �������
         void __fastcall Graf300(double v);                     //������ 300 � 131 ������� / v - �������� ����
         void __fastcall Graf315(double v);                     //������ 315 ������� / v - �������� ����
         void __fastcall Graf320(double v, double n1, double n2, double p1, double p2);  //������ 320 �������
         void __fastcall Graf370(double v, double n1, double n2, double p1, double p2);  //������ 370 �������
         void __fastcall Graf390(double p1, double p2, double n1, double n2);         //������ 390 � 950 ������� /p1,p2 - ����������� ���
         void __fastcall Graf400(double v);                     //������ 400 �������
         void __fastcall Graf410(double v);                     //������ 410 �������
         void __fastcall Graf450(double v, double n1, double n2, double p1, double p2);  //������ 450 �������
         void __fastcall Graf470(double v, double n);           //������ 470 �������
         void __fastcall Graf480();                             //������ 480 �������
         void __fastcall Graf520(double d1, double d2, double d3, double p1, double p2,        // ������ 210(���) ��� 520 �������
                               double p3, double p, double v, double n1, double n2);
         void __fastcall Graf630();                             //������ 630 �������
         void __fastcall Graf650();                             //������ 650 � 660 �������
         void __fastcall Graf670(double v);                     //������ 670 � 790 �������
         void __fastcall Graf680(double v);                     //������ 680 �������
         void __fastcall Graf690(double v);                     //������ 690 �������
         void __fastcall Graf771();                    //������ 711 �������
         void __fastcall Graf775(double v, double n);           //������ 775 �������
         void __fastcall Graf780();                             //������ 780 �������
         void __fastcall Graf800(int day1, int day2);           //������ 800 �������
         void __fastcall Graf850(double v);                     //������ 850 ������� / v - �������� ����
         void __fastcall Graf855(double v);                     //������ 855 ������� 
         void __fastcall Graf865(double v);                     //������ 865 �������
         void __fastcall Graf960(double d1, double d2, double d3, double d4,
                                 double d5, double v, double n);//������ 960 ������� /d1-d6 - ������������ ����� � ����������� �� ������� ������
         void __fastcall Graf980(double p1, double p2, double v, double n1, double n2); //������ 980 �������
         



         int __fastcall DayWeek(int d, int m, int y);           //����������� ��� ������
         bool __fastcall PrazdDni(int d, int m);                //����������� ������������ ���
         bool __fastcall PrdPrazdDni(int d, int m);             //��������������� ���
         bool __fastcall PrazdDniVihodnue(int d, int m, int y); //����������� ������������ �� ��������
         void __fastcall SetInfoEdit();                         //���������� Edit-��
         void __fastcall NextMonth(int &Month, int &Year, int k=1);
         void __fastcall PrevMonth(int &Month, int &Year, int k=1);

         void __fastcall InsertLog(AnsiString Msg); //������� ������� � ������� logs_ro


        __fastcall TMain(TComponent* Owner);
};
//---------------------------------------------------------------------------
extern PACKAGE TMain *Main;
//---------------------------------------------------------------------------
#endif
