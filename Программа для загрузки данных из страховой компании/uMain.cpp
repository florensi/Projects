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

const AnsiString Mes[]={"������","�������","����","������","���","����","����",
                        "������","��������","�������","������","�������"};
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

  // ��������� ������ � ������������ �� ������
  TStringList *SL_Groups = new TStringList();


  // ���������� UserName, DomainName, UserFullName ������ ���� ��������� ��� AnsiString
  if (!GetFullUserInfo(UserName, DomainName, UserFullName))
    {
      MessageBox(Handle,"������ ��������� ������ � ������������","������",8208);
      Application->Terminate();
      Abort();
    }

  //��������� ����� ������� �� ��
  if (!GetUserGroups(UserName, DomainName, SL_Groups))
    {
      MessageBox(Handle,"������ ��������� ������ � ������������","������",8208);
      Application->Terminate();
      Abort();
    }

  //�������� �� ������ � ������
  if ((SL_Groups->IndexOf("mmk-itsvc-hstr-admin")<=-1) && (SL_Groups->IndexOf("mmk-itsvc-hstr")<=-1))
    {
      MessageBox(Handle,"� ��� ��� ���� ��� ������ �\n � ���������� '����������� ����� � ������'!!!","����� �������",8208);
      Application->Terminate();
      Abort();
    }

  if (UserFullName.SubString(1,3)=="rmz")
    {
      ana = 4; //����
    }
  else
    {
      ana = 1; //��� ��.������
    }

  //�������� ����
  //1- ������ ������
 // if (SL_Groups->IndexOf("mmk-itsvc-hstr-01")>-1)
 //   {
      //������ ������
      N9->Visible = true;       //��������������
      N3->Visible = true;       //�������� ������ �� ���������
      N4->Visible = true;       //���������� ��� ���
      N22->Visible = true;      //�������
    //  N15->Visible = false;     //����� "���������� �/�"

      if (ana==4)
        {
          N13->Visible = false;  //����� "��� ��"
          N14->Visible = false;  //����� "��� ��"
          N16->Visible = false;  //����� "��� ���������"
          N17->Visible = true;   //����� "��� ���"
        }
      else
        {
          N13->Visible = true;   //����� "��� ��"
          N14->Visible = true;   //����� "��� ��"
          N16->Visible = true;   //����� "��� ���������"
          N17->Visible = false;  //����� "��� ���"
        }
/*    }
  //2- ��� ���������� (����� �� ���)
  else if (SL_Groups->IndexOf("mmk-itsvc-hstr-02")>-1)
    {
      //��� ����������
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
      Application->MessageBox("�� ����������� ����� �������(����, ����) ��� ������ � ���������� ���� '������� ���������'!!!","����� �������",
                              MB_OK+MB_ICONERROR);
      Application->Terminate();
      Abort();

    }  */
 /*
  // ��������� ������ � ������������ �� ������
  // ���������� UserName, DomainName, UserFullName ������ ���� ��������� ��� AnsiString
  if (!GetUserInfo(UserName, DomainName, UserFullName))
    {
      MessageBox(Handle,"������ ��������� ������ � ������������","������",8208);
      Application->Terminate();
      Abort();
    }

  // ��������� ���� ������� �� ������� users_ro
  DM->qRO_user->Close();
  DM->qRO_user->SQL->Clear();
  DM->qRO_user->SQL->Add("select VU_859, tn, factory from USERS_RO@SLST5 where domain=" + QuotedStr(DomainName) + " and userro=" + QuotedStr(UserName));
  DM->qRO_user->Open();

  if (!DM->qRO_user->RecordCount)
    {
      MessageBox(Handle,("��� ������ � ������������ " + UserName).c_str(),"������",8208);
      Application->Terminate();
      Abort();
    }

  if (DM->qRO_user->FieldByName("VU_859")->AsString.IsEmpty())
    {
      MessageBox(Handle,("��� ������ � ������������ " + UserName).c_str(),"������",8208);
      Application->Terminate();
      Abort();
    }

  if (DM->qRO_user->FieldByName("factory")->AsString.IsEmpty()||
      (DM->qRO_user->FieldByName("factory")->AsInteger !=1 &&
       DM->qRO_user->FieldByName("factory")->AsInteger !=4 ))
    {
      MessageBox(Handle,"�� ������� �����������(���� FACTORY) � ������� USERS_RO","������",8208);
      Application->Terminate();
      Abort();
    }

  Prava = DM->qRO_user->FieldByName("VU_859")->AsInteger;
  TN = DM->qRO_user->FieldByName("tn")->AsString;
  ana = DM->qRO_user->FieldByName("factory")->AsInteger;
  DM->qRO_user->Close();

  switch (Prava)
    {
      case 1:  //������ ������
              N9->Visible = true;       //��������������
              N3->Visible = true;       //�������� ������ �� ���������
              N4->Visible = true;       //���������� ��� ���
              N22->Visible = true;      //�������
              N15->Visible = false;     //����� "���������� �/�"

             if (ana==4)
               {
                 N13->Visible = false;  //����� "��� ��"
                 N14->Visible = false;  //����� "��� ��"
                 N16->Visible = false;  //����� "��� ���������"
                 N17->Visible = true;   //����� "��� ���"
               }
             else
               {
                 N13->Visible = true;   //����� "��� ��"
                 N14->Visible = true;   //����� "��� ��"
                 N16->Visible = true;   //����� "��� ���������"
                 N17->Visible = false;  //����� "��� ���"
               }

      break;

      case 2:  //��� ����������
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
        Application->MessageBox("�������������� ����� �������", "������",
                                     MB_OK + MB_ICONERROR);
        Application->Terminate();
        Abort();
    }

  */


  if (!GetMyDocumentsDir(DocPath))
    {
      MessageBox(Handle,"������ ������� � ����� ����������","������",8208);
      Application->Terminate();
      Abort();
    }

  if (!GetTempDir(TempPath))
    {
      MessageBox(Handle,"������ ������� � ��������� �����","������",8208);
      Application->Terminate();
      Abort();
    }


 WorkPath = DocPath + "\\�������� ������ �� ��������� ��������";


 // �������� ProgressBar �� StatusBar
      ProgressBar = new TProgressBar ( StatusBar1 );
      ProgressBar->Parent = StatusBar1;
      ProgressBar->Position = 0;
      ProgressBar->Left = StatusBar1->Panels->Items[0]->Width + 3;
      ProgressBar->Top = StatusBar1->Height/6;
      ProgressBar->Height = StatusBar1->Height-3;
      ProgressBar->Visible = false;
}
//---------------------------------------------------------------------------

// �������� �� �������� ���� � Excel-�����
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

// �������� ������������ ����, �� � ����� ����������� �� ���������� 15%
 void __fastcall TMain::ProverkaInfoExcel()
{
  AnsiString Sql, Sql1, Sql2, inn, nn, fam, n_doc, nnom, name;
  int fl=0, pr_sum=0, pr_inn=0, pr_dv=0;
  int i=1;
  double sum;
  TSearchRec SearchRecord;  //��� ������ �����


  /*tn - ���.�,
    Row - ����� ���������� ������� ����� � ���������
    sum - ����� ��������� �� ���������
    fl - ������� ������������ ������ (fl=1 - �����������)
    name - ��� Excel �����
    FileName - ������ ���� � ����� � ��� ������
    Dir2 - ���� � ��������� �����
    pr_sum - ������� ������ ����� ������� � ��������� �� ������. ����� (pr_sum = 0 ��������)
    pr_inn - ������� ������ ����� ������� � ��������� �� �����. ���������.� (pr_inn = 0 ��������)
    pr_dv - ������� ������ ����� ������� � ��������� �� ������� ������� (pr_dv = 0 ��������)*/


   //���� ������ ���������� � �����
  if (!SelectDirectory("Select directory",WideString(""),Dir2))
    {
      Abort();
    }

   //����� ����� Excel
   switch(im_fl)
     {
       case 1 : if (FindFirst(Dir2 + LowerCase("\\����� ���.xls"), faAnyFile, SearchRecord)==0 )
                  {
                    name = LowerCase("\\����� ���.xls");
                  }
                else if (FindFirst (Dir2 + LowerCase("\\����� ���.xlsx"), faAnyFile, SearchRecord)==0)
                  {
                    name = LowerCase("\\����� ���.xlsx");
                  }
                else
                  {
                    Application->MessageBox("�� ������ ���� ��� �������� ������. \n�������� ������� �������� ��� ����� \n��� ���� �� ������ � ������ �����. ",
                                           "������ �������� ������", MB_OK + MB_ICONERROR);
                    Abort();
                  }

       break;
       case 2 :  if (FindFirst(Dir2 + LowerCase("\\���������(������).xls"), faAnyFile, SearchRecord)==0 )
                   {
                     name = LowerCase("\\���������(������).xls");
                   }
                 else if (FindFirst (Dir2 + LowerCase("\\���������(������).xlsx"), faAnyFile, SearchRecord)==0)
                   {
                     name = LowerCase("\\���������(������).xlsx");
                   }
                 else
                   {
                     Application->MessageBox("�� ������ ���� ��� �������� ������. \n�������� ������� �������� ��� ����� \n��� ���� �� ������ � ������ �����. ",
                                           "������ �������� ������", MB_OK + MB_ICONERROR);
                     Abort();
                   }
       break;
       case 3 :  if (FindFirst(Dir2 + LowerCase("\\���������(������).xls"), faAnyFile, SearchRecord)==0 )
                   {
                     name = LowerCase("\\���������(������).xls");
                   }
                 else if (FindFirst (Dir2 + LowerCase("\\���������(������).xlsx"), faAnyFile, SearchRecord)==0)
                   {
                     name = LowerCase("\\���������(������).xlsx");
                   }
                 else
                   {
                     Application->MessageBox("�� ������ ���� ��� �������� ������. \n�������� ������� �������� ��� ����� \n��� ���� �� ������ � ������ �����. ",
                                            "������ �������� ������", MB_OK + MB_ICONERROR);
                     Abort();
                   }
       break;
       case 4 :  if (FindFirst(Dir2 + LowerCase("\\���������(����).xls"), faAnyFile, SearchRecord)==0 )
                   {
                     name = LowerCase("\\���������(����).xls");
                   }
                 else if (FindFirst (Dir2 + LowerCase("\\���������(����).xlsx"), faAnyFile, SearchRecord)==0)
                   {
                     name = LowerCase("\\���������(����).xlsx");
                   }
                 else
                   {
                     Application->MessageBox("�� ������ ���� ��� �������� ������. \n�������� ������� �������� ��� ����� \n��� ���� �� ������ � ������ �����. ",
                                           "������ �������� ������", MB_OK + MB_ICONERROR);
                     Abort();
                   }
       break;
       case 5 :  if (FindFirst(Dir2 + LowerCase("\\����� ��.xls"), faAnyFile, SearchRecord)==0 )
                   {
                     name = LowerCase("\\����� ��.xls");
                   }
                 else if (FindFirst (Dir2 + LowerCase("\\����� ��.xlsx"), faAnyFile, SearchRecord)==0)
                   {
                     name = LowerCase("\\����� ��.xlsx");
                   }
                 else
                   {
                     Application->MessageBox("�� ������ ���� ��� �������� ������. \n�������� ������� �������� ��� ����� \n��� ���� �� ������ � ������ �����. ",
                                           "������ �������� ������", MB_OK + MB_ICONERROR);
                     Abort();
                   }
       break;
       case 7 : if (FindFirst(Dir2 + LowerCase("\\����� ����������.xls"), faAnyFile, SearchRecord)==0 )
                  {
                    name = LowerCase("\\����� ����������.xls");
                  }
                else if (FindFirst (Dir2 + LowerCase("\\����� ����������.xlsx"), faAnyFile, SearchRecord)==0)
                  {
                    name = LowerCase("\\����� ����������.xlsx");
                  }
                else
                  {
                    Application->MessageBox("�� ������ ���� ��� �������� ������. �������� ������� �������� ��� ����� (������ ���� '����� ����������.xls' ��� '����� ����������.xlsx') ��� ���� �� ������ � ������ �����.",
                                           "������ �������� ������", MB_OK + MB_ICONERROR);
                    Abort();
                  }
       break;

     }

  FileName = Dir2 + name;  //���� � ����� Excel
  FindClose(SearchRecord);   //����������� �������, ������ ��������� ������
     
  StatusBar1->SimpleText = "";

   // �������������� Excel, ��������� ���� ������
  try
    {
      //���������, ��� �� ����������� Excel
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
          Application->MessageBox("���������� ������� Microsoft Excel!"
          " �������� ��� ���������� �� ���������� �� �����������.","������",MB_OK+MB_ICONERROR);
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
      Application->MessageBox("������ �������� ����� Microsoft Excel!","������",MB_OK + MB_ICONERROR);
    }


  //Excel.OlePropertySet("Visible",true);


  //���������� ���������� ������� ����� � ���������
  Row = Sheet.OlePropertyGet("UsedRange").OlePropertyGet("Rows").OlePropertyGet("Count");


  // ��������� ���� ������ ��� ������������ ������ �� �������������� ��� � �� � ����������
  if (!rtf_Open((TempPath + "\\otchet.txt").c_str()))
    {
      MessageBox(Handle,"������ �������� ����� ������","������",8192);
    }
  else
    {
      Main->Cursor = crHourGlass;
      StatusBar1->SimplePanel = true;    // 2 ������ �� StatusBar1
      StatusBar1->SimpleText=" ����������� �������� ������...";
      ProgressBar->Visible = true;
      ProgressBar->Position = 0;
      ProgressBar->Max = Row;


      for ( i ; i<Row+1; i++)
        {                                                      
          nn = Excel.OlePropertyGet("Cells",i,1);
          inn = Excel.OlePropertyGet("Cells",i,5);
          ProgressBar->Position++;


          // ����� ����� ����������� ��� �������� �� Excel
          if (nn.IsEmpty() || !Proverka(nn) || inn.IsEmpty())  continue;

            sum = Excel.OlePropertyGet("Cells",i,8);
            fam = TrimRight(""+Excel.OlePropertyGet("Cells",i,2)+" "+Excel.OlePropertyGet("Cells",i,3)+" "+Excel.OlePropertyGet("Cells",i,4));
            n_doc = Excel.OlePropertyGet("Cells",i,9);


//�������� �� ��������� ������� � sap_osn_sved � sap_sved_uvol � ������� ���.�
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
                Application->MessageBox("���������� �������� ������ �� ��������� ����������(SAP_OSN_SVED, SAP_SVED_UVOL)","������",MB_OK + MB_ICONERROR);
                Abort();
              }

            if (DM->qObnovlenie->RecordCount>1)
              {
                 pr_sum=0;
                 pr_inn=0;
                //����� � ����� ������� �������
//******************************************************************************
                //����� ������������ � ����� �������
                if (DM->qObnovlenie->RecordCount>1 && pr_dv==0)
                  {
                    rtf_Out("z", " ",3);
                    if(!rtf_LineFeed())
                      {
                        MessageBox(Handle,"������ ������ � ���� ������","������",8192);
                        if (!rtf_Close()) MessageBox(Handle,"������ �������� ����� ������","������",8192);
                        return;
                      }
                   }
                //����� ������� � �����
                rtf_Out("inn", inn,4);
                rtf_Out("fio",fam,4);

                if(!rtf_LineFeed())
                  {
                    MessageBox(Handle,"������ ������ � ���� ������","������",8192);
                    if (!rtf_Close()) MessageBox(Handle,"������ �������� ����� ������","������",8192);
                    return;
                  }
                fl=1;
                pr_dv=1;

              }
            else
              {

                // ����� ��������������� ���.� � �����
//******************************************************************************
                if (DM->qObnovlenie->RecordCount==0)
                  {
                     pr_sum=0;
                     pr_dv=0;
                   //����� ������������ � ����� �������
                    if (DM->qObnovlenie->RecordCount==0 && pr_inn==0)
                      {
                        rtf_Out("z", " ",1);
                        if(!rtf_LineFeed())
                          {
                            MessageBox(Handle,"������ ������ � ���� ������","������",8192);
                            if (!rtf_Close()) MessageBox(Handle,"������ �������� ����� ������","������",8192);
                            return;
                          }
                      }

                    rtf_Out("inn",inn,2);
                    rtf_Out("fio",fam,2);

                    if(!rtf_LineFeed())
                      {
                        MessageBox(Handle,"������ ������ � ���� ������","������",8192);
                        if (!rtf_Close()) MessageBox(Handle,"������ �������� ����� ������","������",8192);
                        return;
                      }

                    fl=1;

                    pr_inn=1;

                  }
                else
                  {
                    /*//�������� �� ���������� ������������ ����� ����� 15%
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
                        if (Application->MessageBox(("��� ����� �� ������� �����\n���="+zex+" ���.�="+tn+" \n���="+fam+" �����="+FloatToStrF(sum,ffFixed,20,2)+" \n��������� ������ � �������?").c_str(),
                                                    "����������",MB_YESNO + MB_ICONINFORMATION)==IDNO)
                          {
                            pr_inn=0;
                            pr_dv=0;
                            // ����� � ����� ���� ��� ����� �� ������� �����
//******************************************************************************
                            //����� ������������ � ����� �������
                            if ((sum >= DM->qObnovlenie->FieldByName("sum")->AsFloat) && pr_sum==0)
                              {
                                rtf_Out("zz", " ",3);

                                if(!rtf_LineFeed())
                                  {
                                    MessageBox(Handle,"������ ������ � ���� ������","������",8192);
                                    if (!rtf_Close()) MessageBox(Handle,"������ �������� ����� ������","������",8192);
                                    return;
                                  }
                              }

                            rtf_Out("zex", zex,4);
                            rtf_Out("tn", tn,4);
                            rtf_Out("fio",fam,4);
                            rtf_Out("n_doc",n_doc ,4);
                            rtf_Out("sum","��� ����� �������� ������",4);

                            if(!rtf_LineFeed())
                              {
                                MessageBox(Handle,"������ ������ � ���� ������","������",8192);
                                if (!rtf_Close()) MessageBox(Handle,"������ �������� ����� ������","������",8192);
                                return;
                              }

                            fl=1;
                            pr_sum=1;

                        }
                      }
                    else if (sum > DM->qObnovlenie->FieldByName("sum")->AsFloat)
                      {
                        if (Application->MessageBox(("����� ��������� 15%\n���="+zex+" ���.�="+tn+" ���="+fam+" �����="+FloatToStrF(sum,ffFixed,20,2)+" \n��������� ������ � �������?").c_str(),
                                                    "����������",MB_YESNO + MB_ICONINFORMATION)==IDNO)
                          {
                            pr_inn=0;
                            pr_dv=0;
                            // ����� � ����� ����������� 15% �����
//******************************************************************************
                            //����� ������������ � ����� �������
                            if ((sum > DM->qObnovlenie->FieldByName("sum")->AsFloat) && pr_sum==0)
                              {
                                rtf_Out("zz", " ",3);

                                if(!rtf_LineFeed())
                                  {
                                    MessageBox(Handle,"������ ������ � ���� ������","������",8192);
                                    if (!rtf_Close()) MessageBox(Handle,"������ �������� ����� ������","������",8192);
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
                                MessageBox(Handle,"������ ������ � ���� ������","������",8192);
                                if (!rtf_Close()) MessageBox(Handle,"������ �������� ����� ������","������",8192);
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
      StatusBar1->SimpleText = "�������� ������ ���������.";
      Main->Cursor = crDefault;

      if(!rtf_Close())
        {
          MessageBox(Handle,"������ �������� ����� ������", "������", 8192);
          return;
        }


      if (fl==1)
        {
          Excel.OleProcedure("Quit");
          StatusBar1->SimpleText = "������������ ������...";
          //�������� �����, ���� �� �� ����������
          ForceDirectories(WorkPath);

          int istrd;
          try
            {
              rtf_CreateReport(TempPath +"\\otchet.txt", Path+"\\RTF\\otchet.rtf",
                         WorkPath+"\\�����.doc",NULL,&istrd);


              WinExec(("\""+ WordPath+"\"\""+WorkPath+"\\�����.doc\"").c_str(),SW_MAXIMIZE);

            }
          catch(RepoRTF_Error E)
            {
              MessageBox(Handle,("������ ������������ ������:"+ AnsiString(E.Err)+
                                 "\n������ ����� ������:"+IntToStr(istrd)).c_str(),"������",8192);
            }

          Application->MessageBox(("��������� ������������� ���������� � ����� \n \""+FileName+"\" � ��������� ��������� ��������").c_str() ," �������� ����� ��������� �� ���",
                                  MB_OK + MB_ICONINFORMATION);
          StatusBar1->SimpleText = "";

          switch (im_fl)
            {
              case 1: InsertLog("����������� ����� �� ����� ��������� ���: ��� ������ �� ���");
              break;
              case 5: InsertLog("����������� ����� �� ����� ������� ���������: ��� ������ �� ���");
              break;
              case 7: InsertLog("����������� ����� �� ����� ��������� �� ����������� �����������: ��� ������ �� ���");
              break;
            }

          Abort();
        }

         DeleteFile(TempPath+"\\otchet.txt");        
    }                                            

   
}
//---------------------------------------------------------------------------

// ���������� ��������� �� ������ ��� ������ ��� �����
 void __fastcall TMain::UpdateValuta_I_Grivna()
{
  AnsiString Sql, inn, nn, data_s, data_po, fam, n_dogovora, Sql1, sum, kod_dog,
             prich, name_otchet;
  int rec=0, kol=0;
  bool fl=0;

 //�������� ����� ������ ��� ������ �� ����������� ������
   if (!rtf_Open((TempPath + "\\izmeneniya.txt").c_str()))
     {
       MessageBox(Handle,"������ �������� ����� ������","������",8192);
     }
   else
     {
       //   Sheet.OleProcedure("Activate");
       Sheet = Book.OlePropertyGet("Worksheets", 1);

       int i=1;

       Main->Cursor = crHourGlass;
       StatusBar1->SimplePanel = true;    // 2 ������ �� StatusBar1
       StatusBar1->SimpleText=" ���� �������� ������...";

       ProgressBar->Visible = true;
       ProgressBar->Position = 0;
       ProgressBar->Max = Row;

        /*  //�������� �� ������� ���������� ������� ��� + �� + ��
          Sql ="SELECT [����1$].* From [����1$] ";

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

           // ����� ����� ����������� ��� �������� �� Excel
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


           //���������� ���+�� �� sap_osn_sved
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
               Application->MessageBox("������ ��������� ������ �� ��������� �� ���������� (SAP_OSN_SVED, SAP_SVED_UVOL)","������",MB_OK+ MB_ICONERROR);
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

           // �������� �� ������� �����
           if (LowerCase(prich)=="��")
             {
               Sql+=", sum="+ QuotedStr(sum)+" , priznak=4";
             }
           else if (LowerCase(prich)=="�")
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


           //�� ��������� ���������
           if (im_fl==2)
             {
               Sql+=", kod_dogovora=0";
             }
           // �� �������� ���������
           if(im_fl==3)
             {
               if (kod_dog =="���"||kod_dog =="���.")
                 {
                   Sql+=", kod_dogovora=1";
                 }
               else
                 {
                   Sql+=", kod_dogovora=2";
                 }
             }
           //�� ������� ���������
           if (im_fl==6)
             {
               Sql+=", kod_dogovora=3";
             }
           //�� ����������� �����������
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
               Application->MessageBox("������ ���������� ������ �� ��������� �����������","������",MB_OK+ MB_ICONERROR);
               Excel.OleProcedure("Quit");
               StatusBar1->SimpleText="";
               ProgressBar->Visible = false;
               Main->Cursor = crDefault;
               Abort();
             }
           rec++;
           kol+=DM->qZagruzka->RowsAffected;

           // ���������� ����������� �������
           if (DM->qZagruzka->RowsAffected == 0)
             {
               //������������ ������ �� ������������� �������
               rtf_Out("inn", inn,1);
               rtf_Out("fio",fam,1);
               rtf_Out("n_dogovora",n_dogovora,1);

               if(!rtf_LineFeed())
                 {
                   MessageBox(Handle,"������ ������ � ���� ������","������",8192);
                   if (!rtf_Close()) MessageBox(Handle,"������ �������� ����� ������","������",8192);
                   return;
                 }
               fl=1;  //������� ������������ ������ �� ������������� �������
             }
         }

       StatusBar1->SimplePanel = false;
       ProgressBar->Visible = false;
       Main->Cursor = crDefault;

       if(!rtf_Close())
         {
           MessageBox(Handle,"������ �������� ����� ������", "������", 8192);
           return;
         }

        // ������������ ������, ���� ���� �� ����������� ������
       if (fl==1)
         {
           //�������� �����, ���� �� �� ����������
           ForceDirectories(WorkPath);


           switch (im_fl)
             {
               case 2: name_otchet = "��������� �� ��������� ���������.doc";
               break;
               case 3: name_otchet = "��������� �� �������� ���������.doc";
               break;
               case 4: name_otchet = "��������� �� �����.doc";
               break;
               case 6: name_otchet = "��������� �� ��.doc";
               break;
               case 7: name_otchet = "��������� �� ���������� ���������.doc";
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
               MessageBox(Handle,("������ ������������ ������:"+ AnsiString(E.Err)+
                                  "\n������ ����� ������:"+IntToStr(istrd)).c_str(),"������",8192);
             }

           Application->MessageBox(("��������� " + IntToStr(kol) + " �� " + IntToStr(rec) + " �������.\n��������� ������������� ���������� � ����� \n \""+FileName+"\" � ��������� ��������� ��������").c_str() ," �������� ����� ��������� �� ���",
                                        MB_OK + MB_ICONINFORMATION);
           StatusBar1->SimpleText = "";

           switch (im_fl)
             {
               case 2: InsertLog("����������� ����� �� ����������(������): ��������� " + IntToStr(kol) + " �� " + IntToStr(rec) + " �������.");
               break;
               case 3: InsertLog("����������� ����� �� ����������(������): ��������� " + IntToStr(kol) + " �� " + IntToStr(rec) + " �������.");
               break;
               case 4: InsertLog("����������� ����� �� ���������� �����: ��������� " + IntToStr(kol) + " �� " + IntToStr(rec) + " �������.");
               break;
               case 6: InsertLog("����������� ����� �� ���������� ��: ��������� " + IntToStr(kol) + " �� " + IntToStr(rec) + " �������.");
               break;
               case 7: InsertLog("����������� ����� �� ���������� �� ���������� ���������: ��������� " + IntToStr(kol) + " �� " + IntToStr(rec) + " �������.");
               break;
             }

           Abort();
         }

       ob_kol = rec;
       obnov_kol = kol;

       DeleteFile(TempPath+"\\izmeneniya.txt");
       StatusBar1->SimpleText = "������ ���������";
       Application->MessageBox(("���������� ������ ��������� ������� =) \n��������� " + IntToStr(kol) + " �� " + IntToStr(rec)+" �������").c_str(),
                                   "���������� ������� �� ����������� �����",
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

//�������� ������ �� ����� ��������� �� ���
void __fastcall TMain::NewDogovClick(TObject *Sender)
{
  AnsiString Sql, Sql1, inn, nn, data_po, data_po1;
  int i=1, rec=0;
  im_fl=1;

  /*rec - ���������� ����������� � ������� �������*/


  if (Application->MessageBox(("�� ������������� ������ ��������� ������ \n �� ����� ��������� �� " + Mes[DM->mm-1] + " " + DM->yyyy + " ����?").c_str(),
                               "�������� ������ �� ����� ���������",
                               MB_YESNO + MB_ICONINFORMATION) == IDNO)
    {
      Abort();
    }

  // �������� ������������ ��� � ������� ������� ������� � ���������
  ProverkaInfoExcel();

  StatusBar1->SimpleText = "";


  try
    {
      Sheet.OleProcedure("Activate");

      Main->Cursor = crHourGlass;
      StatusBar1->SimplePanel = true;    // 2 ������ �� StatusBar1
      StatusBar1->SimpleText=" ���� �������� ������...";

      ProgressBar->Visible = true;
      ProgressBar->Position = 0;
      ProgressBar->Max = Row;

      for ( i ; i<Row+1; i++)
        {
          nn = Excel.OlePropertyGet("Cells",i,1);
          inn = Excel.OlePropertyGet("Cells",i,5);

          ProgressBar->Position++;

          // ����� ����� ����������� ��� �������� �� Excel
          if (nn.IsEmpty() || !Proverka(nn) || inn.IsEmpty())  continue;

             //�������� �� ������� ��� ������������ ������� � ������� VU_859_N
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
                Application->MessageBox("������ ��������� ������ �� ������� �� ����������� 859 �/�","������",MB_OK+ MB_ICONERROR);
                Abort();
              }

            if (DM->qObnovlenie->RecordCount>0)
              {
                 if (Application->MessageBox(("������: ��� = "+ DM->qObnovlenie->FieldByName("zex")->AsString +
                                               ", ���.� = "+ DM->qObnovlenie->FieldByName("tn")->AsString +
                                               ", ��� = "+ DM->qObnovlenie->FieldByName("inn")->AsString +
                                               " � � �������� = "+DM->qObnovlenie->FieldByName("n_dogovora")->AsString +
                                              " ��� ����������. �������� �� ��� ���?").c_str(),"��������������",
                                              MB_YESNO + MB_ICONINFORMATION) ==ID_NO)
                    {
                       continue;
                    }
              }

            //���������� ���+�� �� sap_osn_sved
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
                Application->MessageBox("������ ��������� ������ �� �� ��������� �� ���������� (SAP_OSN_SVED, SAP_SVED_UVOL)","������",MB_OK+ MB_ICONERROR);
                Abort();
              }

            //�������� �� �������� ����
            data_po = Excel.OlePropertyGet("Cells",i,7);
            data_po1 = Excel.OlePropertyGet("Cells",i,7);

            if ((data_po.SubString(1,2)=="31" && data_po.SubString(4,2)=="04")||
                (data_po.SubString(1,2)=="31" && data_po.SubString(4,2)=="06")||
                (data_po.SubString(1,2)=="31" && data_po.SubString(4,2)=="09")||
                (data_po.SubString(1,2)=="31" && data_po.SubString(4,2)=="11"))
              {
                data_po = "30"+ data_po1.SubString(3,255);
              }

            //������ ������ � ������� VU_859_N
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
                Application->MessageBox("������ ������� ������ � ������� �� ����������� 859 �/�","������",MB_OK+ MB_ICONERROR);
                Application->MessageBox("������ �� ���� ���������. ��������� ��������","������",MB_OK+ MB_ICONERROR);
                StatusBar1->SimpleText = "";

                Excel.OleProcedure("Quit");
                Abort();
             }
        }


      Application->MessageBox(("�������� ������ ��������� ������� =) \n ��������� " + IntToStr(rec) + " �������").c_str(),
                               "�������� ����� ��������� �� ���",MB_OK+ MB_ICONINFORMATION);
      InsertLog("��������� �������� ������ �� ����� ��������� �� ���. ��������� "+IntToStr(rec)+" �������");

      Excel.OleProcedure("Quit");
      Excel = Unassigned;

      StatusBar1->SimplePanel = false;
      ProgressBar->Visible = false;
      StatusBar1->SimpleText = "���������� ���������.";
      Main->Cursor = crDefault;
      StatusBar1->SimpleText = "";
    }
  catch(...)
    {
      Application->MessageBox("������ �������� ������ �� ����� ��������� �� ���","������",MB_OK+ MB_ICONERROR);
      Excel.OleProcedure("Quit");

      StatusBar1->SimplePanel = false;
      ProgressBar->Visible = false;
      StatusBar1->SimpleText = "";
      Main->Cursor = crDefault;
    }
}
//---------------------------------------------------------------------------

//---------------------------------------------------------------------------

// ���������� ���� �� ����� ��� ���������"
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

// ���������� ���� �� ����� Temp
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


 // ���������� ������ �������������� Word
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

//�������� ��������� �� �������� ���������
void __fastcall TMain::izm_valClick(TObject *Sender)
{
  im_fl=3;
  if (Application->MessageBox(("�� ������������� ������ ��������� ��������� \n �� �������� ��������� �� " + Mes[DM->mm-1] + " " + DM->yyyy + " ����?").c_str(),
                               "�������� ��������� �� �������� ���������",
                               MB_YESNO + MB_ICONINFORMATION) == IDNO)
    {
      Abort();
    }

  // �������� �� ������������� ��� � ������� Avans
  ProverkaInfoExcelIzmeneniya();

  StatusBar1->SimpleText = "";

  //���������� ��������� �� ������� ���������
  UpdateValuta_I_Grivna();

  InsertLog("��������� �������� ��������� �� �������� ���������. ��������� "+obnov_kol+" �� "+ob_kol+" �������");

  StatusBar1->SimpleText = "";

}
//---------------------------------------------------------------------------

//�������� ��������� �� ��������� ���������
void __fastcall TMain::izm_grnClick(TObject *Sender)
{
  im_fl=2;
  
  if (Application->MessageBox(("�� ������������� ������ ��������� ��������� \n �� ��������� ��������� �� " + Mes[DM->mm-1] + " " + DM->yyyy + " ����?").c_str(),
                               "�������� ��������� �� ��������� ���������",
                               MB_YESNO + MB_ICONINFORMATION) == IDNO)
    {
      Abort();
    }

  // �������� �� ������������� ��� � �������
  ProverkaInfoExcelIzmeneniya();

  StatusBar1->SimpleText = "";

  //���������� ��������� �� ��������� ���������
  UpdateValuta_I_Grivna();

  InsertLog("��������� �������� ��������� �� ��������� ���������. ��������� "+obnov_kol+" �� "+ob_kol+" �������");

  StatusBar1->SimpleText = "";
}
//---------------------------------------------------------------------------

//�������� ��������� ��������� �� �����
void __fastcall TMain::kurs_pereschetClick(TObject *Sender)
{ /*int i=0, rec=0;
  AnsiString tn, fam, data_s,data_po, n_dogovora, Sql;
  bool fl=0;   */

  im_fl=4;

  if (Application->MessageBox(("�� ������������� ������ ��������� ������������� �������� \n � ��������� �� ����� ��� �� " + Mes[DM->mm-1] + " " + DM->yyyy + " ����?").c_str(),
                               "�������� ��������� ��������� �� �����",
                               MB_YESNO + MB_ICONINFORMATION) == IDNO)
    {
      Abort();
    }

  // �������� �� ������������� ��� � ������� Avans
  ProverkaInfoExcelIzmeneniya();

  UpdateValuta_I_Grivna();

  InsertLog("��������� �������� ��������� ��������� �� �����. ��������� "+obnov_kol+" �� "+ob_kol+" �������");

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
      Application->MessageBox("������� �� ��� ������ ��� ���������","���������� ������",
                              MB_OK + MB_ICONINFORMATION);
      EditZEX2->SetFocus();
      Abort();
    }

  if (fl_r==0)
    {
      // ���������� ������

      //�������� �� ���� ������ ��������
      if (EditNDOG->Text.IsEmpty() || EditNDOG->Text.Length()<12)
        {
          Application->MessageBox("������� � ��������","���������� ������",
                                   MB_OK + MB_ICONINFORMATION);
          EditNDOG->SetFocus();
          Abort();
        }

        //�������� �� ������������� ��������� � �������
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
           Application->MessageBox("������ ������� � ������� SAP_OSN_SVED",
                                   "������ �������",MB_OK + MB_ICONERROR);
           Abort();
         }

       if (DM->qObnovlenie->RecordCount==0)
         {
           Application->MessageBox("��� ������ �� ����� ���������.\n��������� ������������ ����� ���� � ���������� ������",
                                   "��������������",MB_OK+MB_ICONINFORMATION);
           Abort();
         }
       else if (DM->qObnovlenie->RecordCount>1 || DM->qObnovlenie->FieldByName("vnvi")->AsString.IsEmpty())
         {
           Application->MessageBox("����� ���� ������� ��� ��� ���.� � ������� SAP_OSN_SVED �� ������� ���������.\n���������� ��������� ����������.",
                                   "������",MB_OK+MB_ICONINFORMATION);
         }
       else
         {
           fio = DM->qObnovlenie->FieldByName("fio")->AsString;
           inn = DM->qObnovlenie->FieldByName("numident")->AsString;

           //�������� �� ������������� ������ � ������� vu_859_n
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
               Application->MessageBox("�������� ������ ��� �������� ������ � ������� VU_859_N",
                                       "������",MB_OK + MB_ICONERROR);
               Abort();
            }

          if (DM->qObnovlenie->RecordCount>0)
            {
              Application->MessageBox("�������� � ����� ������� �������� ��� ����������",
                                       "������",MB_OK + MB_ICONERROR);
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
              Application->MessageBox("�������� ������ ��� ���������� ������",
                                      "������ ���������� ����� ������",MB_OK + MB_ICONERROR);
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
               Application->MessageBox("������ ��������� ������ �� �������","������",MB_OK + MB_ICONERROR);
               Abort();
             }

           InsertLog("��������� ���������� ������: ��� ="+ EditZEX2->Text +", ���.� ="+ EditTN2->Text +", � �������� = "+EditNDOG->Text+", ����� = "+EditSum->Text);

           TLocateOptions SearchOptions;
           DM->qKorrektirovka->Locate("n_dogovora",n_dog,SearchOptions<<loPartialKey<<loCaseInsensitive);

        }
    }
  else
    {
      // �������������� ������

      //�������� �� ���� �������� �������
      if (EditPRIZNAK->Text.IsEmpty())
        {
          Application->MessageBox("������� ������� �������\n   0 - ������\n   1 - ������\n   2 - ��������� ����� ��������\n   3 - ������\n   4 - ����������� ��������\n   6 - ��������������� ��������\n   7 - ������� �� ������� � ����",
                                  "���������� ������",
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
          Application->MessageBox("�������� ������ ��� ���������� ������",
                                  "������ ���������� ������",MB_OK + MB_ICONERROR);
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
          InsertLog("�������������� ������ �� �������� �: "+DM->qKorrektirovka->FieldByName("n_dogovora")->AsString+" c ��� = "+ zzex +" �� "+EditZEX2->Text+" c ���.� ="+ztn+" �� "+EditTN2->Text+" c ����� = "+zsum+" �� "+EditSum->Text+" � ������ "+zval+" �� "+EditVal->Text+" � �������� "+zpriznak+" �� "+EditPRIZNAK->Text+" � ���� "+zdata_s+"-"+zdata_po+" �� "+EditData_s->Text+"-"+EditData_po->Text);
        }
      rec = DM->qKorrektirovka->RecNo;

      DM->qKorrektirovka->Requery();

      //������� �� ����������� ������
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

//����� ������ ��� ��������������
void __fastcall TMain::BitBtn3Click(TObject *Sender)
{
  if (EditZEX->Text.IsEmpty() || EditTN->Text.IsEmpty())
    {
      Application->MessageBox("�� ������ ��� ��� ��������� ����� ���������","������� ����������� ����������", MB_OK + MB_ICONINFORMATION);
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
      Application->MessageBox("������ ��������� ������ �� �������","������",MB_OK + MB_ICONERROR);
      Abort();
    }

  if (DM->qKorrektirovka->RecordCount==0)
    {
       Application->MessageBox("�������� � ����� ����� � ��������� ������� �� ������","����� ������",MB_OK + MB_ICONINFORMATION);
       EditZEX->SetFocus();
       EditZEX->SelectAll();
       Label11->Visible=false;
       Abort();

    }

  //����� ������� � DBGrid
  Label11->Visible = true;
  Label12->Visible = true;


  Label9->Caption="�������������� ������:";

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
      case 0:  Label11->Caption="������";
      break;
      case 1:  Label11->Caption="������";
      break;
      case 2:  Label11->Caption="���������� ���� ��������";
      break;
      case 3:  Label11->Caption="������";
      break;
      case 4:  Label11->Caption="�����������";
      break;
      case 5:  Label11->Caption="�������������";
      break;
      case 6:  Label11->Caption="�������������";
      break;
      case 7:  Label11->Caption=" �� ������� � ����";
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
    // �������� ����

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
              Application->MessageBox("�������� ������ ����","������", MB_OK);
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
              Application->MessageBox("�������� ������ ����","������", MB_OK);
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

// ������ ������
void __fastcall TMain::N7Click(TObject *Sender)
{
  AnsiString Sql;
//  Word year,month,day;

 // DecodeDate(Date(), year, month,day);


//******************************************************************************
   //���������� ���+�� �� ���������

   StatusBar1->SimpleText = "���������� ������ �� ������������ ����������...";

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
       Application->MessageBox("�������� ������ ��� ������� ������ �� ���������",
                               "���������� ������",MB_OK + MB_ICONERROR);
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
           Application->MessageBox("�������� ������ ��� ���������� ������ �� ���������",
                                   "���������� ������",MB_OK + MB_ICONERROR);

           StatusBar1->SimpleText = "";
           InsertLog("������ ������ �� ���������: ������ ���������� ���+��� �� ���������");
           Abort();
         }

       DM->qObnovlenie->Next();

     }
//********************************************************************************
  // ���������� ���� priznak: ��������� = 1
  StatusBar1->SimpleText = "���������� ������ �� ��������";

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
      Application->MessageBox("�������� ������ ��� ���������� ������ �� ��������",
                              "���������� ������",MB_OK + MB_ICONERROR);
      StatusBar1->SimpleText = "";
      InsertLog("������ ������ �� ���������: ������ ���������� ������ �� ��������");
      Abort();
    }

  // ���������� ���� priznak: ������� ���� ������ = 2, ����� ����������� �����������
  StatusBar1->SimpleText = "���������� ������ �� ��������� ����� ��������";

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
      Application->MessageBox("�������� ������ ��� ���������� ������ �� ���������, \n � ������� ��������� ����",
                              "���������� ������",MB_OK + MB_ICONERROR);

      StatusBar1->SimpleText = "";
      InsertLog("������ ������ �� ���������: ������ ���������� ������ �� ��������� ����� ��������");
      Abort();
    }


 //*****************************************************************************


  // ������������ ������ ����������� ���������, ���������� ������ � ���������� ����� � ���.�
  StatusBar1->SimpleText = "���� ������������ ������...";

  //�������� ����� ������ ����������� ���������, ���������� ������ � ���������� ����� � ���.�
  if (!rtf_Open((TempPath + "\\sverka.txt").c_str()))
    {
      MessageBox(Handle,"������ �������� ����� ������","������",8192);
    }
  else
    {

// ������� ������ �� ���������
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
          Application->MessageBox("�������� ������ ��� ������� ������ �� ������� �� ��������",
                                  "������������ ������",MB_OK + MB_ICONERROR);
          StatusBar1->SimpleText = "";
          Abort();
        }

// ����� � ����� ���������
      if (DM->qObnovlenie->RecordCount>0)
        {
          // ����� ��������� � ����� �������
          rtf_Out("z", " ", 1);
          if(!rtf_LineFeed())
            {
              MessageBox(Handle,"������ ������ � ���� ������","������",8192);
              if (!rtf_Close()) MessageBox(Handle,"������ �������� ����� ������","������",8192);
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
              MessageBox(Handle,"������ ������ � ���� ������","������",8192);
              if (!rtf_Close()) MessageBox(Handle,"������ �������� ����� ������","������",8192);
              return;
            }

          DM->qObnovlenie->Next();
        }


// ������� ������ �� ���������, � ������� ��������� ���� ������
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
          Application->MessageBox("�������� ������ ��� ������� ������ �� ���������, \n � ������� ��������� ����",
                                  "������������ ������",MB_OK + MB_ICONERROR);

          StatusBar1->SimpleText = "";
          Abort();
        }

      //����� � ����� �������� � ���������� �����
      if (DM->qObnovlenie->RecordCount>0)
        {
          // ����� ��������� � ����� �������
          rtf_Out("zz", " ", 3);
          if(!rtf_LineFeed())
            {
              MessageBox(Handle,"������ ������ � ���� ������","������",8192);
              if (!rtf_Close()) MessageBox(Handle,"������ �������� ����� ������","������",8192);
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
              MessageBox(Handle,"������ ������ � ���� ������","������",8192);
              if (!rtf_Close()) MessageBox(Handle,"������ �������� ����� ������","������",8192);
              return;
            }

          DM->qObnovlenie->Next();
        }
  //  }

//������� ������ �� ��������� �� ��� � ��
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
          Application->MessageBox("�������� ������ ��� ������� ������ �� ���������",
                                  "���������� ������",MB_OK + MB_ICONERROR);

          StatusBar1->SimpleText = "";
          Abort();
        }

//����� � ����� ��������� �� ��� � ��
      if (DM->qObnovlenie->RecordCount>0)
        {
          // ����� ��������� � ����� �������
          rtf_Out("zzz", " ", 5);
          if(!rtf_LineFeed())
            {
              MessageBox(Handle,"������ ������ � ���� ������","������",8192);
              if (!rtf_Close()) MessageBox(Handle,"������ �������� ����� ������","������",8192);
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
              MessageBox(Handle,"������ ������ � ���� ������","������",8192);
              if (!rtf_Close()) MessageBox(Handle,"������ �������� ����� ������","������",8192);
              return;
            }

          DM->qObnovlenie->Next();
        }
      
      if(!rtf_Close())
        {
          MessageBox(Handle,"������ �������� ����� ������", "������", 8192);
          return;
        }
      int istrd;
      try
        {
          rtf_CreateReport(TempPath +"\\sverka.txt", Path+"\\RTF\\sverka.rtf",
                           WorkPath+"\\������ ������.doc",NULL,&istrd);


              WinExec(("\""+ WordPath+"\"\""+WorkPath+"\\������ ������.doc\"").c_str(),SW_MAXIMIZE);

        }
      catch(RepoRTF_Error E)
        {
          MessageBox(Handle,("������ ������������ ������:"+ AnsiString(E.Err)+
                             "\n������ ����� ������:"+IntToStr(istrd)).c_str(),"������",8192);
        }

    }
    
    InsertLog("������ ������ ��������� �������");
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

     ��������� ��������� ������ ������� ������
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

      // ��������� ������ ���� �������,
      if(!State.Contains(gdFixed) && DM->qKorrektirovka->RecNo==rec)
        {
          ((TDBGridEh *) Sender)->Canvas->Brush->Color = TColor(0x001FE0B0);           //clGradientActiveCaption;
          ((TDBGridEh *) Sender)->DefaultDrawColumnCell(Rect, DataCol, Column, State);
        }

       // ��������� ������ �������� ������
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
// �������� ������� ���������
void __fastcall TMain::N11Click(TObject *Sender)
{
  AnsiString Sql, Sql1, nn, inn, vnvi, data_po, data_po1;
  int i=1, rec=0;

  im_fl=5;  // ��� ������ ����� ������������ �����

  if (Application->MessageBox(("�� ������������� ������ ��������� ������ \n �� ������� ��������� �� " + Mes[DM->mm-1] + " " + DM->yyyy + " ����?").c_str(),
                               "�������� ������ �� ������� ���������",
                               MB_YESNO + MB_ICONINFORMATION) == IDNO)
    {
      Abort();
    }

  // �������� ������������ ����, �� � ����� ����������� �� ���������� 15%
  ProverkaInfoExcel();

  StatusBar1->SimpleText = "";

  try
    {
      Sheet.OleProcedure("Activate");

      Main->Cursor = crHourGlass;
      StatusBar1->SimplePanel = true;    // 2 ������ �� StatusBar1
      StatusBar1->SimpleText=" ���� �������� ������� ���������...";

      ProgressBar->Visible = true;
      ProgressBar->Position = 0;
      ProgressBar->Max = Row;

      for ( i ; i<Row+1; i++)
        {
          nn = Excel.OlePropertyGet("Cells",i,1);
          inn = Excel.OlePropertyGet("Cells",i,5);
          ProgressBar->Position++;


          // ����� ����� ����������� ��� �������� �� Excel
          if (nn.IsEmpty() || !Proverka(nn) || inn.IsEmpty())  continue;

            //�������� �� ������� ��� ������������ ������� � ������� VU_859_N
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
                Application->MessageBox("������ ��������� ������ �� ������� �� ����������� 859 �/�","������",MB_OK+ MB_ICONERROR);
                Abort();
              }

            if (DM->qObnovlenie->RecordCount>0)
              {
                 if (Application->MessageBox(("������: ��� = "+ DM->qObnovlenie->FieldByName("zex")->AsString +
                                               ", ���.� = "+ DM->qObnovlenie->FieldByName("tn")->AsString +
                                               ", ��� = "+ DM->qObnovlenie->FieldByName("inn")->AsString +
                                               " � � �������� = "+DM->qObnovlenie->FieldByName("n_dogovora")->AsString +
                                              " ��� ����������. �������� �� ��� ���?").c_str(),"��������������",
                                              MB_YESNO + MB_ICONINFORMATION) ==ID_NO)
                    {
                       continue;
                    }
              }

            //���������� ���+�� �� sap_osn_sved
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
                Application->MessageBox("������ ��������� ������ �� ������� avans","������",MB_OK+ MB_ICONERROR);
                Abort();
              }

             //�������� �� �������� ����

            data_po = Excel.OlePropertyGet("Cells",i,7);
            data_po1 = Excel.OlePropertyGet("Cells",i,7);

            if ((data_po.SubString(1,2)=="31" && data_po.SubString(4,2)=="04")||
                (data_po.SubString(1,2)=="31" && data_po.SubString(4,2)=="06")||
                (data_po.SubString(1,2)=="31" && data_po.SubString(4,2)=="09")||
                (data_po.SubString(1,2)=="31" && data_po.SubString(4,2)=="11"))
              {
                data_po = "30"+ data_po1.SubString(3,255);
              }

            //������ ������ � ������� VU_859_N
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
                Application->MessageBox("������ ������� ������ � ������� �� ����������� 859 �/�","������",MB_OK+ MB_ICONERROR);
                Application->MessageBox("������ �� ���� ���������. ��������� ��������","������",MB_OK+ MB_ICONERROR);
                StatusBar1->SimpleText = "";

                Excel.OleProcedure("Quit");
                Abort();
             }
        }



      Application->MessageBox(("�������� ������ ��������� ������� =) \n ��������� " + IntToStr(rec) + " �������").c_str(),
                               "�������� ������� ���������",MB_OK+ MB_ICONINFORMATION);
      InsertLog("��������� �������� ������ �� ������� ���������. ��������� "+IntToStr(rec)+" �������");

      Excel.OleProcedure("Quit");
      Excel = Unassigned;

      StatusBar1->SimplePanel = false;
      ProgressBar->Visible = false;
      StatusBar1->SimpleText = "������ ������� ��������� ��������� �������";
      Main->Cursor = crDefault;
      StatusBar1->SimpleText = "";
    }
  catch(...)
    {
      Application->MessageBox("������ �������� ������ �� ������� ���������","������",MB_OK+ MB_ICONERROR);
      Excel.OleProcedure("Quit");

      StatusBar1->SimplePanel = false;
      ProgressBar->Visible = false;
      StatusBar1->SimpleText = "";
      Main->Cursor = crDefault;
    }

}

//---------------------------------------------------------------------------

// ��������� ����� ������������ ��������
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

// ��� ��
void __fastcall TMain::N13Click(TObject *Sender)
{
  Data->ShowModal();
  if (Data->ModalResult == mrCancel) {Abort();}

  AnsiString Sql,Sql1,sum_pr, tn_pr,tn_pr1;
  int zex1,zex,kust1,kust;
  Double sum_zex=0, sum_kust=0, rl_sum=0, obsh_sum=0;

      /* dtp_month, dtp_year - ����� � ��� �� DateTimePicker
         dtp_mm - ����� �� DateTimePicker c "0"
         sum_pr - ����� ����������
         zex1 - ���������� ���
         zex - ������� ���
         kust1 - ���������� ����
         kust - ������� ����
         sum_kust - ����� �� �����
         rl_sum - ���������� �� ��������
         obsh_sum - ����� ����� �� ���� ��������� ��� 859
         tn_pr - ������� ���.�
         tn_pr1 - ���������� ���.�*/

  //���������� ������ �� DateTimePicker
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
      MessageBox(Handle,"������ �������� ����� ������","������",8192);
    }
  else
    {
      StatusBar1->SimplePanel = true;    // 2 ������ �� StatusBar1
      StatusBar1->SimpleText = " ���� ������������ ��������� ��������� �� ����. �����... ";

      // ������������ ��������� ��������� �� ����. �����
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
         Application->MessageBox("������ ��������� ������ �� ������� SLST. \n �������� ������� ������ ������.","������",MB_OK);
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

                  //�������� ����������
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
                      Application->MessageBox("������ ��������� ������ �� ������� SLST.","������",MB_OK);
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

                  //����� ����������
                  if (zex==zex1 && tn_pr==tn_pr1)
                    {
                      rtf_Out("prev"," ", 1);

                      if(!rtf_LineFeed())
                        {
                          MessageBox(Handle,"������ ������ � ���� ������","������",8192);
                          if (!rtf_Close()) MessageBox(Handle,"������ �������� ����� ������","������",8192);
                          return;
                        }
                    }
                  else
                    {
                      rtf_Out("prev",sum_pr, 1);

                      if(!rtf_LineFeed())
                        {
                          MessageBox(Handle,"������ ������ � ���� ������","������",8192);
                          if (!rtf_Close()) MessageBox(Handle,"������ �������� ����� ������","������",8192);
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

              //����� ����� �� ����
              rtf_Out("sum_po_zex", FloatToStrF(sum_zex,ffFixed,20,2),2);

              if(!rtf_LineFeed())
                {
                  MessageBox(Handle,"������ ������ � ���� ������","������",8192);
                  if (!rtf_Close()) MessageBox(Handle,"������ �������� ����� ������","������",8192);
                  return;
                }

              sum_zex=0;
            }

          //����� ����� �� �����
          rtf_Out("sum_po_kust", FloatToStrF(sum_kust, ffFixed,20,2),3);
          if(!rtf_LineFeed())
            {
              MessageBox(Handle,"������ ������ � ���� ������","������",8192);
              if (!rtf_Close()) MessageBox(Handle,"������ �������� ����� ������","������",8192);
              return;
            }
        }

      // �� ���������
      rtf_Out("sum_po_kombinat",FloatToStrF(DM->qZagruzka->FieldByName("sum_po_kombinat")->AsFloat,ffFixed,20,2), 0);

      StatusBar1->SimpleText = " ���� ������������ �������� ��������� �� ����. �����... ";

 //*****************************************************************************
      // ������������ �������� ��������� �� ����. �����

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
          Application->MessageBox("������ ��������� ������ �� ������� SLST. \n �������� ������� ������ ������.","������",MB_OK);
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

                  //�������� ����������
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
                      Application->MessageBox("������ ��������� ������ �� ������� SLST.","������",MB_OK);
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

                  //����� ����������
                  if (zex==zex1 && tn_pr==tn_pr1)
                    {
                      rtf_Out("prev"," ", 4);

                      if(!rtf_LineFeed())
                        {
                          MessageBox(Handle,"������ ������ � ���� ������","������",8192);
                          if (!rtf_Close()) MessageBox(Handle,"������ �������� ����� ������","������",8192);
                          return;
                        }
                    }
                  else
                    {
                      rtf_Out("prev",sum_pr, 4);

                      if(!rtf_LineFeed())
                        {
                          MessageBox(Handle,"������ ������ � ���� ������","������",8192);
                          if (!rtf_Close()) MessageBox(Handle,"������ �������� ����� ������","������",8192);
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

              //����� ����� �� ����
              rtf_Out("sum_po_zex", FloatToStrF(sum_zex,ffFixed,20,2),5);

              if(!rtf_LineFeed())
                {
                  MessageBox(Handle,"������ ������ � ���� ������","������",8192);
                  if (!rtf_Close()) MessageBox(Handle,"������ �������� ����� ������","������",8192);
                  return;
                }

              sum_zex=0;
            }

           //����� ����� �� �����
           rtf_Out("sum_po_kust", FloatToStrF(sum_kust, ffFixed,20,2),6);
           if(!rtf_LineFeed())
             {
               MessageBox(Handle,"������ ������ � ���� ������","������",8192);
               if (!rtf_Close()) MessageBox(Handle,"������ �������� ����� ������","������",8192);
               return;
             }

        }

      // �� ���������
      rtf_Out("sum_po_kombinat",FloatToStrF(DM->qZagruzka->FieldByName("sum_po_kombinat")->AsFloat,ffFixed,20,2), 0);

      StatusBar1->SimpleText = " ���� ������������ ������� ��������� �� ����. �����... ";

//*****************************************************************************
// ������� �������� �� ����. �����

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
          Application->MessageBox("������ ��������� ������ �� ������� SLST.","������",MB_OK);

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

                  //�������� ����������
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
                      Application->MessageBox("������ ��������� ������ �� ������� SLST.","������",MB_OK);

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

                  //����� ����������
                  if (zex==zex1 && tn_pr==tn_pr1)
                    {
                      rtf_Out("prev"," ", 7);

                      if(!rtf_LineFeed())
                        {
                          MessageBox(Handle,"������ ������ � ���� ������","������",8192);
                          if (!rtf_Close()) MessageBox(Handle,"������ �������� ����� ������","������",8192);
                          return;
                        }
                    }
                  else
                    {
                      rtf_Out("prev",sum_pr, 7);

                      if(!rtf_LineFeed())
                        {
                          MessageBox(Handle,"������ ������ � ���� ������","������",8192);
                          if (!rtf_Close()) MessageBox(Handle,"������ �������� ����� ������","������",8192);
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

              //����� ����� �� ����
              rtf_Out("sum_po_zex", FloatToStrF(sum_zex,ffFixed,20,2),8);

              if(!rtf_LineFeed())
                {
                  MessageBox(Handle,"������ ������ � ���� ������","������",8192);
                  if (!rtf_Close()) MessageBox(Handle,"������ �������� ����� ������","������",8192);
                  return;
                }

              sum_zex=0;
            }

           //����� ����� �� �����
           rtf_Out("sum_po_kust", FloatToStrF(sum_kust, ffFixed,20,2),9);
           if(!rtf_LineFeed())
             {
               MessageBox(Handle,"������ ������ � ���� ������","������",8192);
               if (!rtf_Close()) MessageBox(Handle,"������ �������� ����� ������","������",8192);
               return;
             }

        }

  // �� ���������
  rtf_Out("sum_po_kombinat",FloatToStrF(DM->qZagruzka->FieldByName("sum_po_kombinat")->AsFloat,ffFixed,20,2), 0);

  StatusBar1->SimplePanel = false;
  ProgressBar->Visible = false;
  StatusBar1->SimpleText = "";
  Main->Cursor = crDefault;


  if(!rtf_Close())
    {
      MessageBox(Handle,"������ �������� ����� ������", "������", 8192);
      return;
    }

  //�������� �����, ���� �� �� ����������
  ForceDirectories(WorkPath);

  int istrd;
  try
    {
      rtf_CreateReport(TempPath +"\\dlya_ro.txt", Path+"\\RTF\\dlya_ro.rtf",
                       WorkPath+"\\��� ��(����.����).doc",NULL,&istrd);


      WinExec(("\""+ WordPath+"\"\""+WorkPath+"\\��� ��(����.����).doc\"").c_str(),SW_MAXIMIZE);
    }
  catch(RepoRTF_Error E)
    {
      MessageBox(Handle,("������ ������������ ������:"+ AnsiString(E.Err)+
                         "\n������ ����� ������:"+IntToStr(istrd)).c_str(),"������",8192);
    }
    StatusBar1->SimpleText = " ";
  }

}

//---------------------------------------------------------------------------

// ���������� �\�
void __fastcall TMain::N15Click(TObject *Sender)
{
  Data->ShowModal();
  if (Data->ModalResult == mrCancel) {Abort();}

  AnsiString Sql,Sql1, sum_pr;
  int zex1,zex,kust1,kust;
  Double sum_zex=0, sum_kust=0, rl_sum=0, obsh_sum=0;

      /* dtp_month, dtp_year - ����� � ��� �� DateTimePicker
         dtp_mm - ����� �� DateTimePicker c "0"
         sum_pr - ����� ����������*/

  //���������� ������ �� DateTimePicker
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

   // ������������ ��������� ��������� �� ���������� �\�
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
      Application->MessageBox("������ ��������� ������ �� ������� SLST. \n �������� ������� ������ ������.","������",MB_OK);
      Abort();
    }
  //�������� ����� ������ ����������� ���������, ���������� ������ � ���������� ����� � ���.�
  if (!rtf_Open((TempPath + "\\dlya_sh.txt").c_str()))
    {
      MessageBox(Handle,"������ �������� ����� ������","������",8192);
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

                  //�������� ����������
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
                      Application->MessageBox("������ ��������� ������ �� ������� SLST.","������",MB_OK);
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

                  //����� ����������
                  rtf_Out("prev",sum_pr, 1);

                  if(!rtf_LineFeed())
                    {
                      MessageBox(Handle,"������ ������ � ���� ������","������",8192);
                      if (!rtf_Close()) MessageBox(Handle,"������ �������� ����� ������","������",8192);
                      return;
                    }

                  sum_zex += DM->qZagruzka->FieldByName("sum")->AsFloat,20,2;
                  sum_kust = DM->qZagruzka->FieldByName("sum_po_kust")->AsFloat;


                  DM->qZagruzka->Next();
                  ProgressBar->Position++;

                  zex = DM->qZagruzka->FieldByName("zex")->AsInteger;
                  kust = DM->qZagruzka->FieldByName("kust")->AsInteger;
                }

              //����� ����� �� ����
              rtf_Out("sum_po_zex", FloatToStrF(sum_zex,ffFixed,20,2),2);

              if(!rtf_LineFeed())
                {
                  MessageBox(Handle,"������ ������ � ���� ������","������",8192);
                  if (!rtf_Close()) MessageBox(Handle,"������ �������� ����� ������","������",8192);
                  return;
                }

              sum_zex=0;
            }

          //����� ����� �� �����
          rtf_Out("sum_po_kust", FloatToStrF(sum_kust, ffFixed,20,2),3);
          if(!rtf_LineFeed())
            {
              MessageBox(Handle,"������ ������ � ���� ������","������",8192);
              if (!rtf_Close()) MessageBox(Handle,"������ �������� ����� ������","������",8192);
              return;
            }
        }

      // �� ���������
      rtf_Out("sum_po_kombinat",FloatToStrF(DM->qZagruzka->FieldByName("sum_po_kombinat")->AsFloat,ffFixed,20,2), 0);

       StatusBar1->SimpleText = " ���� ������������ �������� ��������� �� ���������� �\� ";

 //*****************************************************************************
// ������������ �������� ��������� �� ���������� �/�
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
           Application->MessageBox("������ ��������� ������ �� ������� SLST. \n �������� ������� ������ ������.","������",MB_OK);
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

                  //�������� ����������
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
                      Application->MessageBox("������ ��������� ������ �� ������� SLST.","������",MB_OK);
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

                  //����� ����������
                  rtf_Out("prev",sum_pr, 4);

                  if(!rtf_LineFeed())
                    {
                      MessageBox(Handle,"������ ������ � ���� ������","������",8192);
                      if (!rtf_Close()) MessageBox(Handle,"������ �������� ����� ������","������",8192);
                      return;
                    }

                  sum_zex += DM->qZagruzka->FieldByName("sum")->AsFloat,20,2;
                  sum_kust = DM->qZagruzka->FieldByName("sum_po_kust")->AsFloat;


                  DM->qZagruzka->Next();
                  ProgressBar->Position++;

                  zex = DM->qZagruzka->FieldByName("zex")->AsInteger;
                  kust = DM->qZagruzka->FieldByName("kust")->AsInteger;
                }

              //����� ����� �� ����
              rtf_Out("sum_po_zex", FloatToStrF(sum_zex,ffFixed,20,2),5);

              if(!rtf_LineFeed())
                {
                  MessageBox(Handle,"������ ������ � ���� ������","������",8192);
                  if (!rtf_Close()) MessageBox(Handle,"������ �������� ����� ������","������",8192);
                  return;
                }

              sum_zex=0;
            }

          //����� ����� �� �����
          rtf_Out("sum_po_kust", FloatToStrF(sum_kust, ffFixed,20,2),6);
          if(!rtf_LineFeed())
            {
              MessageBox(Handle,"������ ������ � ���� ������","������",8192);
              if (!rtf_Close()) MessageBox(Handle,"������ �������� ����� ������","������",8192);
              return;
            }
        }

      // �� ���������
      rtf_Out("sum_po_kombinat",FloatToStrF(DM->qZagruzka->FieldByName("sum_po_kombinat")->AsFloat,ffFixed,20,2), 0);

       StatusBar1->SimpleText = " ���� ������������ �������� ��������� �� ���������� �\�";

 //*****************************************************************************
   // ������������ ������� ��������� �� ���������� �\�
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
          Application->MessageBox("������ ��������� ������ �� ������� SLST. \n �������� ������� ������ ������.","������",MB_OK);
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

                  //�������� ����������
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
                      Application->MessageBox("������ ��������� ������ �� ������� SLST.","������",MB_OK);
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

                  //����� ����������
                  rtf_Out("prev",sum_pr, 7);

                  if(!rtf_LineFeed())
                    {
                      MessageBox(Handle,"������ ������ � ���� ������","������",8192);
                      if (!rtf_Close()) MessageBox(Handle,"������ �������� ����� ������","������",8192);
                      return;
                    }

                  sum_zex += DM->qZagruzka->FieldByName("sum")->AsFloat,20,2;
                  sum_kust = DM->qZagruzka->FieldByName("sum_po_kust")->AsFloat;


                  DM->qZagruzka->Next();
                  ProgressBar->Position++;

                  zex = DM->qZagruzka->FieldByName("zex")->AsInteger;
                  kust = DM->qZagruzka->FieldByName("kust")->AsInteger;
                }

              //����� ����� �� ����
              rtf_Out("sum_po_zex", FloatToStrF(sum_zex,ffFixed,20,2),8);

              if(!rtf_LineFeed())
                {
                  MessageBox(Handle,"������ ������ � ���� ������","������",8192);
                  if (!rtf_Close()) MessageBox(Handle,"������ �������� ����� ������","������",8192);
                  return;
                }

              sum_zex=0;
            }

          //����� ����� �� �����
          rtf_Out("sum_po_kust", FloatToStrF(sum_kust, ffFixed,20,2),9);
          if(!rtf_LineFeed())
            {
              MessageBox(Handle,"������ ������ � ���� ������","������",8192);
              if (!rtf_Close()) MessageBox(Handle,"������ �������� ����� ������","������",8192);
              return;
            }
        }

      // �� ���������
      rtf_Out("sum_po_kombinat",FloatToStrF(DM->qZagruzka->FieldByName("sum_po_kombinat")->AsFloat,ffFixed,20,2), 0);

  StatusBar1->SimplePanel = false;
  ProgressBar->Visible = false;
  StatusBar1->SimpleText = "";
  Main->Cursor = crDefault;


  if(!rtf_Close())
    {
      MessageBox(Handle,"������ �������� ����� ������", "������", 8192);
      return;
    }

  //�������� �����, ���� �� �� ����������
  ForceDirectories(WorkPath);

  int istrd;
  try
    {
      rtf_CreateReport(TempPath +"\\dlya_sh.txt", Path+"\\RTF\\dlya_sh.rtf",
                       WorkPath+"\\��� ���������� �\�.doc",NULL,&istrd);


      WinExec(("\""+ WordPath+"\"\""+WorkPath+"\\��� ���������� �\�.doc\"").c_str(),SW_MAXIMIZE);
    }
  catch(RepoRTF_Error E)
    {
      MessageBox(Handle,("������ ������������ ������:"+ AnsiString(E.Err)+
                         "\n������ ����� ������:"+IntToStr(istrd)).c_str(),"������",8192);
    }

    }
   
   StatusBar1->SimpleText = " ";
}
//---------------------------------------------------------------------------

//���������
void __fastcall TMain::N17Click(TObject *Sender)
{
  Data->ShowModal();
  if (Data->ModalResult == mrCancel) {Abort();}

  AnsiString Sql,Sql1, sum_pr,firma, firma1, tn_pr, tn_pr1;
  int zex1,zex, ana, ana1;
  Double sum_zex=0, sum_kust=0, rl_sum=0, obsh_sum=0;

      /* dtp_month, dtp_year - ����� � ��� �� DateTimePicker
         dtp_mm - ����� �� DateTimePicker c "0"
         sum_pr - ����� ����������*/

  //���������� ������ �� DateTimePicker
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

   StatusBar1->SimpleText = " ���� ������������ ��������� ��������� �� ���...";

     // ������������ ��������� ��������� �� ����������
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
      Application->MessageBox("������ ��������� ������ �� ������� SLST. \n �������� ������� ������ ������.","������",MB_OK);
      StatusBar1->SimpleText = "";
      Abort();
    }
  //�������� ����� ������
  if (!rtf_Open((TempPath + "\\dlya_agro.txt").c_str()))
    {
      MessageBox(Handle,"������ �������� ����� ������","������",8192);
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
                    rtf_Out("naim", "�� ���������", 1);

                    if(!rtf_LineFeed())
                      {
                        MessageBox(Handle,"������ ������ � ���� ������","������",8192);
                        if (!rtf_Close()) MessageBox(Handle,"������ �������� ����� ������","������",8192);
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

                  //�������� ����������
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
                      Application->MessageBox("������ ��������� ������ �� ������� SLST.","������",MB_OK);
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

                  //����� ����������
                  if (zex==zex1 && tn_pr==tn_pr1)
                    {
                      rtf_Out("prev"," ", 2);

                      if(!rtf_LineFeed())
                        {
                          MessageBox(Handle,"������ ������ � ���� ������","������",8192);
                          if (!rtf_Close()) MessageBox(Handle,"������ �������� ����� ������","������",8192);
                          return;
                        }
                    }
                  else
                    {
                      rtf_Out("prev",sum_pr, 2);

                      if(!rtf_LineFeed())
                        {
                          MessageBox(Handle,"������ ������ � ���� ������","������",8192);
                          if (!rtf_Close()) MessageBox(Handle,"������ �������� ����� ������","������",8192);
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

              //����� ����� �� ����
              rtf_Out("sum_po_zex", FloatToStrF(sum_zex,ffFixed,20,2),3);

              if(!rtf_LineFeed())
                {
                  MessageBox(Handle,"������ ������ � ���� ������","������",8192);
                  if (!rtf_Close()) MessageBox(Handle,"������ �������� ����� ������","������",8192);
                  return;
                }

              sum_zex=0;
            }

          //����� ����� �� �����
          rtf_Out("sum_po_kust", FloatToStrF(sum_kust, ffFixed,20,2),4);
          rtf_Out("firma2", firma1,4);

          if(!rtf_LineFeed())
            {
              MessageBox(Handle,"������ ������ � ���� ������","������",8192);
              if (!rtf_Close()) MessageBox(Handle,"������ �������� ����� ������","������",8192);
              return;
            }
        }
    /*
      // �� ���������
      rtf_Out("sum_po_kombinat",FloatToStrF(DM->qZagruzka->FieldByName("sum_po_kombinat")->AsFloat,ffFixed,20,2), 5);
      if(!rtf_LineFeed())
        {
          MessageBox(Handle,"������ ������ � ���� ������","������",8192);
          if (!rtf_Close()) MessageBox(Handle,"������ �������� ����� ������","������",8192);
          return;
        }
    */
       StatusBar1->SimpleText = " ���� ������������ �������� ��������� �� ���... ";

 //*****************************************************************************
    // ������������ �������� ��������� �� ����������
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
      Application->MessageBox("������ ��������� ������ �� ������� SLST. \n �������� ������� ������ ������.","������",MB_OK);
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
                    rtf_Out("naim", "� ������", 1);

                    if(!rtf_LineFeed())
                      {
                        MessageBox(Handle,"������ ������ � ���� ������","������",8192);
                        if (!rtf_Close()) MessageBox(Handle,"������ �������� ����� ������","������",8192);
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

                  //�������� ����������
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
                      Application->MessageBox("������ ��������� ������ �� ������� SLST.","������",MB_OK);
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

                  //����� ����������
                  if (zex==zex1 && tn_pr==tn_pr1)
                    {
                      rtf_Out("prev"," ", 2);

                      if(!rtf_LineFeed())
                        {
                          MessageBox(Handle,"������ ������ � ���� ������","������",8192);
                          if (!rtf_Close()) MessageBox(Handle,"������ �������� ����� ������","������",8192);
                          return;
                        }
                    }
                  else
                    {
                      rtf_Out("prev",sum_pr, 2);

                      if(!rtf_LineFeed())
                        {
                          MessageBox(Handle,"������ ������ � ���� ������","������",8192);
                          if (!rtf_Close()) MessageBox(Handle,"������ �������� ����� ������","������",8192);
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

              //����� ����� �� ����
              rtf_Out("sum_po_zex", FloatToStrF(sum_zex,ffFixed,20,2),3);

              if(!rtf_LineFeed())
                {
                  MessageBox(Handle,"������ ������ � ���� ������","������",8192);
                  if (!rtf_Close()) MessageBox(Handle,"������ �������� ����� ������","������",8192);
                  return;
                }

              sum_zex=0;
            }

          //����� ����� �� �����
          rtf_Out("sum_po_kust", FloatToStrF(sum_kust, ffFixed,20,2),4);
          rtf_Out("firma2", firma1,4);

          if(!rtf_LineFeed())
            {
              MessageBox(Handle,"������ ������ � ���� ������","������",8192);
              if (!rtf_Close()) MessageBox(Handle,"������ �������� ����� ������","������",8192);
              return;
            }
        }
   /*
      // �� ���������
      rtf_Out("sum_po_kombinat",FloatToStrF(DM->qZagruzka->FieldByName("sum_po_kombinat")->AsFloat,ffFixed,20,2), 5);
      if(!rtf_LineFeed())
        {
          MessageBox(Handle,"������ ������ � ���� ������","������",8192);
          if (!rtf_Close()) MessageBox(Handle,"������ �������� ����� ������","������",8192);
          return;
        }
            */
       StatusBar1->SimpleText = " ���� ������������ ������� ��������� �� ���...";

 //*****************************************************************************
   // ������������ ������� ��������� �� ����������
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
      Application->MessageBox("������ ��������� ������ �� ������� SLST. \n �������� ������� ������ ������.","������",MB_OK);
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
                    rtf_Out("naim", "�� ������� ���������", 1);

                    if(!rtf_LineFeed())
                      {
                        MessageBox(Handle,"������ ������ � ���� ������","������",8192);
                        if (!rtf_Close()) MessageBox(Handle,"������ �������� ����� ������","������",8192);
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

                  //�������� ����������
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
                      Application->MessageBox("������ ��������� ������ �� ������� SLST.","������",MB_OK);
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

                  //����� ����������
                  if (zex==zex1 && tn_pr==tn_pr1)
                    {
                      rtf_Out("prev"," ", 2);

                      if(!rtf_LineFeed())
                        {
                          MessageBox(Handle,"������ ������ � ���� ������","������",8192);
                          if (!rtf_Close()) MessageBox(Handle,"������ �������� ����� ������","������",8192);
                          return;
                        }
                    }
                  else
                    {
                      rtf_Out("prev",sum_pr, 2);

                      if(!rtf_LineFeed())
                        {
                          MessageBox(Handle,"������ ������ � ���� ������","������",8192);
                          if (!rtf_Close()) MessageBox(Handle,"������ �������� ����� ������","������",8192);
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

              //����� ����� �� ����
              rtf_Out("sum_po_zex", FloatToStrF(sum_zex,ffFixed,20,2),3);

              if(!rtf_LineFeed())
                {
                  MessageBox(Handle,"������ ������ � ���� ������","������",8192);
                  if (!rtf_Close()) MessageBox(Handle,"������ �������� ����� ������","������",8192);
                  return;
                }

              sum_zex=0;
            }

          //����� ����� �� �����
          rtf_Out("sum_po_kust", FloatToStrF(sum_kust, ffFixed,20,2),4);
          rtf_Out("firma2", firma1,4);

          if(!rtf_LineFeed())
            {
              MessageBox(Handle,"������ ������ � ���� ������","������",8192);
              if (!rtf_Close()) MessageBox(Handle,"������ �������� ����� ������","������",8192);
              return;
            }
        }
   /*
      // �� ���������
      rtf_Out("sum_po_kombinat",FloatToStrF(DM->qZagruzka->FieldByName("sum_po_kombinat")->AsFloat,ffFixed,20,2), 5);
      if(!rtf_LineFeed())
        {
          MessageBox(Handle,"������ ������ � ���� ������","������",8192);
          if (!rtf_Close()) MessageBox(Handle,"������ �������� ����� ������","������",8192);
          return;
        }
         */
  StatusBar1->SimplePanel = false;
  ProgressBar->Visible = false;
  StatusBar1->SimpleText = "";
  Main->Cursor = crDefault;

  if(!rtf_Close())
    {
      MessageBox(Handle,"������ �������� ����� ������", "������", 8192);
      return;
    }

  //�������� �����, ���� �� �� ����������
  ForceDirectories(WorkPath);

  int istrd;
  try
    {
      rtf_CreateReport(TempPath +"\\dlya_agro.txt", Path+"\\RTF\\dlya_agro.rtf",
                       WorkPath+"\\��� ��������.doc",NULL,&istrd);


      WinExec(("\""+ WordPath+"\"\""+WorkPath+"\\��� ��������.doc\"").c_str(),SW_MAXIMIZE);
    }
  catch(RepoRTF_Error E)
    {
      MessageBox(Handle,("������ ������������ ������:"+ AnsiString(E.Err)+
                         "\n������ ����� ������:"+IntToStr(istrd)).c_str(),"������",8192);
    }

    }
   
   StatusBar1->SimpleText = " ";
}
//---------------------------------------------------------------------------
//������������ ����� ��� ��
void __fastcall TMain::N14Click(TObject *Sender)
{
  AnsiString Sql;

  Data->ShowModal();
  if (Data->ModalResult == mrCancel) {Abort();}

  //���������� ������ �� DateTimePicker
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
                 sl.zex , sl.tn, sl.sum as sum, sp.kust, sp.firma, decode(nvl(sl.nist,0),0,'���',1,'���',2,'����') as nist,             \
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
      Application->MessageBox("������ ��������� ������ �� ������� SLST. \n �������� ������� ������ ������.","������",MB_OK);
      Abort();
    }

   // ���������� �������
  int row = DM->qZagruzka->RecordCount;

  // ������������� ���� � ����� �������
  AnsiString sFile = Path+"\\RTF\\dlya_ok.xlt";


   // �������������� Excel, ��������� ���� ������
  try
    {
      AppEx=GetActiveOleObject("Excel.Application");
    }
  catch(...)
    {
     //���������, ��� �� ����������� Excel
      try
        {
          AppEx=CreateOleObject("Excel.Application");
        }
      catch (...)
        {
          Application->MessageBox("���������� ������� Microsoft Excel!"
              " �������� ��� ���������� �� ���������� �� �����������.","������",MB_OK+MB_ICONERROR);
        }
    }

  try
    {
      AppEx.OlePropertyGet("WorkBooks").OleProcedure("Open",sFile.c_str());     //��������� �����, ������ � ���
      Sh=AppEx.OlePropertyGet("WorkSheets",1);                                  //�������� � ��������� ����� �����
    }
  catch(...)
    {
      Application->MessageBox("������ �������� ����� Microsoft Excel!","������",MB_OK+MB_ICONERROR);
    }

  StatusBar1->SimpleText = " ���� ������������ ��������� �� ����. ����� � ���������� �\�... ";

  Main->Cursor = crHourGlass;
  ProgressBar->Visible = true;
  ProgressBar->Position = 0;
  ProgressBar->Max = DM->qZagruzka->RecordCount;

   // ��������� � ������ ������ ���������� �����

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

   //��������� ����� ��������� � ��������� ���� "�������� ����"..."
   AppEx.OlePropertySet("DisplayAlerts",false);

   //�������� ����� ���� �� �� ����������
   ForceDirectories(WorkPath);

   //��������� ����� � ����� � ����� �� ��������
   AnsiString vAsCurDir1=WorkPath+"\\��� ��.xls";
   AppEx.OlePropertyGet("Workbooks").OlePropertyGet("Item",1).
   OleProcedure("SaveAs",vAsCurDir1.c_str());

   //������� Excel
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

  //������������ ����� �� ��������� ��������� ����+����
  FILE *grn;


  Data->ShowModal();
  if (Data->ModalResult == mrCancel) {Abort();}

  //���������� ������ �� DateTimePicker
  DecodeDate(Data->DateTimePicker1->Date, dtp_year, dtp_month, dtp_day );

  //�������� ����� ���� �� �� ����������
  ForceDirectories(WorkPath);

  if ((grn=fopen((WorkPath+"\\fkiev.txt").c_str(),"wt"))==NULL)
    {
      ShowMessage("���� �� ������� �������");
      return;
    }

   StatusBar1->SimpleText = "������������ ����� �� ��������� ��������� ����+���� ��� ��������� ��������... ";

  if (StrToInt(dtp_month)<10)
        {
          dtp_mm ="0"+ IntToStr(dtp_month);
        }
      else
        {
          dtp_mm = IntToStr(dtp_month);
        }

   //������
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
      Application->MessageBox("������ ��������� ������ �� ������� SLST. \n �������� ������� ������ ������.","������",MB_OK);
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

  //����� � ����
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


  StatusBar1->SimpleText = "������������ ����� �� �������� ��������� � ������� ��������� ����+����...";

  //������������ ����� �� �������� ��������� � ������� ��������� ����+����
  FILE *val;
  if ((val=fopen((WorkPath+"\\fkievv.txt").c_str(),"wt"))==NULL)
    {
      ShowMessage("���� �� ������� �������");
      return;
    }
        // ������ �� ����.
       //������
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
      Application->MessageBox("������ ��������� ������ �� ������� SLST. \n �������� ������� ������ ������.","������",MB_OK);
      StatusBar1->SimplePanel = false;
      ProgressBar->Visible = false;
      StatusBar1->SimpleText = "";
      Main->Cursor = crDefault;
    }

  ProgressBar->Position = 0;
  ProgressBar->Max = DM->qObnovlenie->RecordCount;

  //����� � ����
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

    // ������� ����
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
      Application->MessageBox("������ ��������� ������ �� ������� SLST. \n �������� ������� ������ ������.","������",MB_OK);
      StatusBar1->SimplePanel = false;
      ProgressBar->Visible = false;
      StatusBar1->SimpleText = "";
      Main->Cursor = crDefault;
    }

  ProgressBar->Position = 0;
  ProgressBar->Max = DM->qObnovlenie->RecordCount;

  //����� � ����
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

  // ���� ������+����
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
      Application->MessageBox("������ ��������� ������ �� ������� SLST. \n �������� ������� ������ ������.","������",MB_OK);
      StatusBar1->SimplePanel = false;
      ProgressBar->Visible = false;
      StatusBar1->SimpleText = "";
      Main->Cursor = crDefault;
      Abort();
    }

  ProgressBar->Position = 0;
  ProgressBar->Max = DM->qObnovlenie->RecordCount;

  //����� � ����
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
  ShowMessage("������������ ������ ������� ���������");
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

// �������� �� ������������� ��� ��� ��������� �� ���������
 void __fastcall TMain::ProverkaInfoExcelIzmeneniya()
{
  AnsiString Sql, Sql1, Sql2, tn, fam, n_doc, nn, inn, nnom, name, data_po;
  int fl=0, pr_inn=0, pr_dv=0, pr_sum=0;
  int i=1;
  double sum;
  TSearchRec SearchRecord;  //��� ������ �����


  /*tn - ���.�,
    Row - ����� ���������� ������� ����� � ���������
    sum - ����� ��������� �� ���������
    fl - ������� ������������ ������ (fl=1 - �����������)
    name - ��� Excel �����
    FileName - ������ ���� � ����� � ��� ������
    Dir2 - ���� � ��������� �����
    pr_sum - ������� ������ ����� ������� � ��������� �� ������. ����� (pr_sum = 0 ��������)
    pr_inn - ������� ������ ����� ������� � ��������� �� �����. ���������.� (pr_inn = 0 ��������)
    pr_dv - ������� ������ ����� ������� � ��������� �� ������� ������� (pr_dv = 0 ��������)*/


   //���� ������ ���������� � �����
  if (!SelectDirectory("Select directory",WideString(""),Dir2))
    {
      Abort();
    }

   //����� ����� Excel
   switch(im_fl)
     {
       case 2 :  if (FindFirst(Dir2 + LowerCase("\\���������(������).xls"), faAnyFile, SearchRecord)==0 )
                   {
                     name = LowerCase("\\���������(������).xls");
                   }
                 else if (FindFirst (Dir2 + LowerCase("\\���������(������).xlsx"), faAnyFile, SearchRecord)==0)
                   {
                     name = LowerCase("\\���������(������).xlsx");
                   }
                 else
                   {
                     Application->MessageBox("�� ������ ���� ��� �������� ������. \n�������� ������� �������� ��� ����� \n��� ���� �� ������ � ������ �����. ",
                                           "������ �������� ������", MB_OK + MB_ICONERROR);
                     Abort();
                   }
       break;
       case 3 :  if (FindFirst(Dir2 + LowerCase("\\���������(������).xls"), faAnyFile, SearchRecord)==0 )
                   {
                     name = LowerCase("\\���������(������).xls");
                   }
                 else if (FindFirst (Dir2 + LowerCase("\\���������(������).xlsx"), faAnyFile, SearchRecord)==0)
                   {
                     name = LowerCase("\\���������(������).xlsx");
                   }
                 else
                   {
                     Application->MessageBox("�� ������ ���� ��� �������� ������. \n�������� ������� �������� ��� ����� \n��� ���� �� ������ � ������ �����. ",
                                            "������ �������� ������", MB_OK + MB_ICONERROR);
                     Abort();
                   }
       break;
       case 4 :  if (FindFirst(Dir2 + LowerCase("\\���������(����).xls"), faAnyFile, SearchRecord)==0 )
                   {
                     name = LowerCase("\\���������(����).xls");
                   }
                 else if (FindFirst (Dir2 + LowerCase("\\���������(����).xlsx"), faAnyFile, SearchRecord)==0)
                   {
                     name = LowerCase("\\���������(����).xlsx");
                   }
                 else
                   {
                     Application->MessageBox("�� ������ ���� ��� �������� ������. \n�������� ������� �������� ��� ����� \n��� ���� �� ������ � ������ �����. ",
                                           "������ �������� ������", MB_OK + MB_ICONERROR);
                     Abort();
                   }
       break;
       case 6 :  if (FindFirst(Dir2 + LowerCase("\\���������(��).xls"), faAnyFile, SearchRecord)==0 )
                   {
                     name = LowerCase("\\���������(��).xls");
                   }
                 else if (FindFirst (Dir2 + LowerCase("\\���������(��).xlsx"), faAnyFile, SearchRecord)==0)
                   {
                     name = LowerCase("\\���������(��).xlsx");
                   }
                 else
                   {
                     Application->MessageBox("�� ������ ���� ��� �������� ������. \n�������� ������� �������� ��� ����� \n��� ���� �� ������ � ������ �����. ",
                                           "������ �������� ������", MB_OK + MB_ICONERROR);
                     Abort();
                   }
       break;
       case 7 : if (FindFirst(Dir2 + LowerCase("\\���������(����������).xls"), faAnyFile, SearchRecord)==0 )
                  {
                    name = LowerCase("\\���������(����������).xls");
                  }
                else if (FindFirst (Dir2 + LowerCase("\\���������(����������).xlsx"), faAnyFile, SearchRecord)==0)
                  {
                    name = LowerCase("\\���������(����������).xlsx");
                  }
                else
                  {
                    Application->MessageBox("�� ������ ���� ��� �������� ������. �������� ������� �������� ��� ����� (������ ���� '���������(����������).xls' ��� '���������(����������).xlsx') ��� ���� �� ������ � ������ �����.",
                                           "������ �������� ������", MB_OK + MB_ICONERROR);
                    Abort();
                  }
       break;


     }

  FileName = Dir2 + name;  //���� � ����� Excel
  FindClose(SearchRecord);   //����������� �������, ������ ��������� ������
     
  StatusBar1->SimpleText = "";

   // �������������� Excel, ��������� ���� ������
  try
    {
      //���������, ��� �� ����������� Excel
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
          Application->MessageBox("���������� ������� Microsoft Excel!"
          " �������� ��� ���������� �� ���������� �� �����������.","������",MB_OK+MB_ICONERROR);
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
      Application->MessageBox("������ �������� ����� Microsoft Excel!","������",MB_OK + MB_ICONERROR);
    }


  //Excel.OlePropertySet("Visible",true);

  //���������� ���������� ������� ����� � ���������
  Row = Sheet.OlePropertyGet("UsedRange").OlePropertyGet("Rows").OlePropertyGet("Count");
//Row=100;
  // ��������� ���� ������ ��� ������������ ������ �� �������������� ���
  if (!rtf_Open((TempPath + "\\otchet.txt").c_str()))
    {
      MessageBox(Handle,"������ �������� ����� ������","������",8192);
    }
  else
    {

      Main->Cursor = crHourGlass;
      StatusBar1->SimplePanel = true;    // 2 ������ �� StatusBar1
      StatusBar1->SimpleText=" ����������� �������� ������...";
      ProgressBar->Visible = true;
      ProgressBar->Position = 0;
      ProgressBar->Max = Row;


      for ( i ; i<Row+1; i++)
        {
          nn = Excel.OlePropertyGet("Cells",i,1);
          inn = Excel.OlePropertyGet("Cells",i,5);
          ProgressBar->Position++;


          // ����� ����� ����������� ��� �������� �� Excel
          if (nn.IsEmpty() || !Proverka(nn) || inn.IsEmpty())  continue;

            inn = Excel.OlePropertyGet("Cells",i,5);
            if (im_fl==7) sum = Excel.OlePropertyGet("Cells",i,9);
            else sum = Excel.OlePropertyGet("Cells",i,10);
            fam = TrimRight(""+Excel.OlePropertyGet("Cells",i,2)+" "+Excel.OlePropertyGet("Cells",i,3)+" "+Excel.OlePropertyGet("Cells",i,4));
            n_doc = Excel.OlePropertyGet("Cells",i,11);
            data_po = Excel.OlePropertyGet("Cells",i,7);



//�������� �� ��������� ������� � sap_osn_sved � sap_sved_uvol � ������� ���.�
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
                Application->MessageBox("���������� �������� ������ �� ��������� ����������(SAP_OSN_SVED, SAP_SVED_UVOL)","������",MB_OK + MB_ICONERROR);
                Abort();
              }

            if (DM->qObnovlenie->RecordCount>1)
              {
                 pr_inn=0;
                 pr_sum=0;
                //����� � ����� ������� �������
//******************************************************************************
                //����� ������������ � ����� �������
                if (DM->qObnovlenie->RecordCount>1 && pr_dv==0)
                  {
                    rtf_Out("z", " ",3);
                    if(!rtf_LineFeed())
                      {
                        MessageBox(Handle,"������ ������ � ���� ������","������",8192);
                        if (!rtf_Close()) MessageBox(Handle,"������ �������� ����� ������","������",8192);
                        return;
                      }
                   }
                //����� ������� � �����
                rtf_Out("inn", inn,4);
                rtf_Out("fio",fam,4);

                if(!rtf_LineFeed())
                  {
                    MessageBox(Handle,"������ ������ � ���� ������","������",8192);
                    if (!rtf_Close()) MessageBox(Handle,"������ �������� ����� ������","������",8192);
                    return;
                  }
                fl=1;
                pr_dv=1;

              }
            else if (DM->qObnovlenie->RecordCount==0)
              {
               //��� ������� �� � ������� sap_osn_sved, �� � sap_sved_uvol
                // ����� ��������������� ���.� � �����
//******************************************************************************
                if (DM->qObnovlenie->RecordCount==0)
                  {
                     pr_dv=0;
                     pr_sum=0;
                   //����� ������������ � ����� �������
                    if (DM->qObnovlenie->RecordCount==0 && pr_inn==0)
                      {
                        rtf_Out("z", " ",1);
                        if(!rtf_LineFeed())
                          {
                            MessageBox(Handle,"������ ������ � ���� ������","������",8192);
                            if (!rtf_Close()) MessageBox(Handle,"������ �������� ����� ������","������",8192);
                            return;
                          }
                      }


                    rtf_Out("inn",inn,2);
                    rtf_Out("fio",fam,2);

                    if(!rtf_LineFeed())
                      {
                        MessageBox(Handle,"������ ������ � ���� ������","������",8192);
                        if (!rtf_Close()) MessageBox(Handle,"������ �������� ����� ������","������",8192);
                        return;
                      }

                    fl=1;

                    pr_inn=1;

                  }
              }
           //�������� ������ ������� �������
           else if (DM->qObnovlenie->RecordCount==1 && DM->qObnovlenie->FieldByName("priznak")->AsInteger==2 )
              {
               //��� ������� �� � ������� sap_osn_sved, �� � sap_sved_uvol
                // ����� ��������������� ���.� � �����
//******************************************************************************
                if (DM->qObnovlenie->RecordCount==1 && DM->qObnovlenie->FieldByName("priznak")->AsInteger==2 )
                  {
                     pr_dv=0;
                     pr_inn=0;
                   //����� ������������ � ����� �������
                    if (DM->qObnovlenie->RecordCount==1 && pr_sum==0)
                      {
                        rtf_Out("z", " ",5);
                        if(!rtf_LineFeed())
                          {
                            MessageBox(Handle,"������ ������ � ���� ������","������",8192);
                            if (!rtf_Close()) MessageBox(Handle,"������ �������� ����� ������","������",8192);
                            return;
                          }
                      }


                    rtf_Out("inn",inn,6);
                    rtf_Out("fio",fam,6);

                    if(!rtf_LineFeed())
                      {
                        MessageBox(Handle,"������ ������ � ���� ������","������",8192);
                        if (!rtf_Close()) MessageBox(Handle,"������ �������� ����� ������","������",8192);
                        return;
                      }

                    fl=1;

                    pr_sum=1;

                  }
              }


            else
              {
           //�������� �� ���������� ������������ ����� ����� 15%,
           // ���� ���� ������������� (���� �� ������) �� ��������� �� ������
           // � ��������� �� ������
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
                        if (Application->MessageBox(("��� ����� �� ������� �����\n���="+zex+" ���.�="+tn+" ���="+fam+" �����="+FloatToStrF(sum,ffFixed,20,2)+" \n��������� ������ � �������?").c_str(),
                                                    "����������",MB_YESNO + MB_ICONINFORMATION)==IDNO)
                          {
                            pr_inn=0;
                            pr_dv=0;
                            // ����� � ����� ���� ��� ����� �� ������� �����
//******************************************************************************
                            //����� ������������ � ����� �������
                            if ((sum >= DM->qObnovlenie->FieldByName("sum")->AsFloat) && pr_sum==0)
                              {
                                rtf_Out("zz", " ",3);

                                if(!rtf_LineFeed())
                                  {
                                    MessageBox(Handle,"������ ������ � ���� ������","������",8192);
                                    if (!rtf_Close()) MessageBox(Handle,"������ �������� ����� ������","������",8192);
                                    return;
                                  }
                              }

                            rtf_Out("zex", zex,4);
                            rtf_Out("tn", tn,4);
                            rtf_Out("fio",fam,4);
                            rtf_Out("n_doc",n_doc ,4);
                            rtf_Out("sum","��� ����� �������� ������",4);

                            if(!rtf_LineFeed())
                              {
                                MessageBox(Handle,"������ ������ � ���� ������","������",8192);
                                if (!rtf_Close()) MessageBox(Handle,"������ �������� ����� ������","������",8192);
                                return;
                              }

                            fl=1;
                            pr_sum=1;
                          }
                      }
                   else if (sum > DM->qObnovlenie->FieldByName("sum")->AsFloat)
                     {
                       if (Application->MessageBox(("����� ��������� 15%\n���="+zex+" ���.�="+tn+" ���="+fam+" �����="+FloatToStrF(sum,ffFixed,20,2)+" \n��������� ������ � �������?").c_str(),
                                                    "����������",MB_YESNO + MB_ICONINFORMATION)==IDNO)
                         {
                           pr_inn=0;
                           pr_dv=0;
                           // ����� � ����� ����������� 15% �����
//******************************************************************************
                           //����� ������������ � ����� �������
                           if ((sum > DM->qObnovlenie->FieldByName("sum")->AsFloat) && pr_sum==0)
                             {
                               rtf_Out("zz", " ",3);

                               if(!rtf_LineFeed())
                                 {
                                   MessageBox(Handle,"������ ������ � ���� ������","������",8192);
                                   if (!rtf_Close()) MessageBox(Handle,"������ �������� ����� ������","������",8192);
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
                               MessageBox(Handle,"������ ������ � ���� ������","������",8192);
                               if (!rtf_Close()) MessageBox(Handle,"������ �������� ����� ������","������",8192);
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
      StatusBar1->SimpleText = "�������� ������ ���������.";
      Main->Cursor = crDefault;

      if(!rtf_Close())
        {
          MessageBox(Handle,"������ �������� ����� ������", "������", 8192);
          return;
        }


      if (fl==1)
        {
          Excel.OleProcedure("Quit");
          StatusBar1->SimpleText = "������������ ������...";
          //�������� �����, ���� �� �� ����������
          ForceDirectories(WorkPath);

          int istrd;
          try
            {
              rtf_CreateReport(TempPath +"\\otchet.txt", Path+"\\RTF\\otchet.rtf",
                         WorkPath+"\\�����.doc",NULL,&istrd);


              WinExec(("\""+ WordPath+"\"\""+WorkPath+"\\�����.doc\"").c_str(),SW_MAXIMIZE);

            }
          catch(RepoRTF_Error E)
            {
              MessageBox(Handle,("������ ������������ ������:"+ AnsiString(E.Err)+
                                 "\n������ ����� ������:"+IntToStr(istrd)).c_str(),"������",8192);
            }

          Application->MessageBox(("��������� ������������� ���������� � ����� \n \""+FileName+"\" � ��������� ��������� ��������").c_str() ," �������� ����� ��������� �� ���",
                                  MB_OK + MB_ICONINFORMATION);
          StatusBar1->SimpleText = "";

          switch (im_fl)
             {
               case 2: InsertLog("����������� ����� �� ����������(������): ��� ������ �� ���");
               break;
               case 3: InsertLog("����������� ����� �� ����������(������): ��� ������ �� ���");
               break;
               case 4: InsertLog("����������� ����� �� ���������� �����: ��� ������ �� ���");
               break;
               case 6: InsertLog("����������� ����� �� ���������� ��: ��� ������ �� ���");
               break;
               case 7: InsertLog("����������� ����� �� ���������� �� ����������� �����������: ��� ������ �� ���");
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


  /*kust - ����� �������� �����, kust1 - ����� ���������� �����;
  double sum_kust - ����� �� ����� */

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
      MessageBox(Handle,"�������� ������ ��� ������� ������ �� ������ ������� ��������� ��������� ��� ���","������",8202);
      Abort();
    }


  if (DM->qObnovlenie->RecordCount>0)
    {
      //�������� �����, ���� �� �� ����������
      ForceDirectories(WorkPath);

      if (!rtf_Open ((TempPath + "\\v_yit.txt").c_str()))
        {
          MessageBox(Handle,"������ �������� ����� ������","������",8192);
        }
      else
        {
          StatusBar1->SimpleText = "������������ ������� ��������� ��������� ��� ���...";
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
                      MessageBox(Handle,"������ ������ � ���� ������","������",8192);
                      if (!rtf_Close()) MessageBox(Handle,"������ �������� ����� ������","������",8192);
                      return;
                    }

                  DM->qObnovlenie->Next();
                  val1 = DM->qObnovlenie->FieldByName("kod_dogovora")->AsInteger;
                }

              // ����� ����� �� ��������
              switch (val)
                {
                  case 0: rtf_Out("naim", "���������", 2);
                  break;
                  case 1: rtf_Out("naim", "��������(������)", 2);
                  break;
                  case 2: rtf_Out("naim", "��������(����)", 2);
                  break;
                  case 3: rtf_Out("naim", "�������", 2);
                  break;

                }

              rtf_Out("sumdog", sum_dog,10,2, 2);
              rtf_Out("koldog", kol_dog, 2);

              val = DM->qObnovlenie->FieldByName("kod_dogovora")->AsInteger;
              if(!rtf_LineFeed())
                {
                   MessageBox(Handle,"������ ������ � ���� ������","������",8192);
                   if (!rtf_Close()) MessageBox(Handle,"������ �������� ����� ������","������",8192);
                   return;
                }
            }

          // ����� ����� �� �����������
          rtf_Out("sumobsh", DM->qObnovlenie->FieldByName("sumobsh")->AsString, 0);
          rtf_Out("kolobsh", DM->qObnovlenie->FieldByName("kolobsh")->AsString, 0);

          if(!rtf_Close())
            {
              MessageBox(Handle,"������ �������� ����� ������", "������", 8192);
              return;
            }

          int istrd;
          try
            {
              rtf_CreateReport(TempPath +"\\v_yit.txt", Path+"\\RTF\\v_yit.rtf",
                               WorkPath+"\\������� ��������� ��������� ��� ���.doc",NULL,&istrd);
              DeleteFile(TempPath+"\\v_yit.txt");

              WinExec(("\""+ WordPath+"\"\""+WorkPath+"\\������� ��������� ��������� ��� ���.doc\"").c_str(),SW_MAXIMIZE);
            }
          catch(RepoRTF_Error E)
            {
              MessageBox(Handle,("������ ������������ ������:"+ AnsiString(E.Err)+
                                 "\n������ ����� ������:"+IntToStr(istrd)).c_str(),"������",8192);
            }
         StatusBar1->SimpleText = "";
        }
    }
  else
   {
     Application->MessageBox("��� ������ �� ������� �����", "��������������",
                              MB_OK + MB_ICONWARNING);
   }

}
//---------------------------------------------------------------------------

void __fastcall TMain::N19Click(TObject *Sender)
{
  im_fl=6;
  if (Application->MessageBox(("�� ������������� ������ ��������� ��������� \n �� �������� ��������� �� " + Mes[DM->mm-1] + " " + DM->yyyy + " ����?").c_str(),
                               "�������� ��������� �� �������� ���������",
                               MB_YESNO + MB_ICONINFORMATION) == IDNO)
    {
      Abort();
    }

  // �������� �� ������������� ��� � ������� Avans
  ProverkaInfoExcelIzmeneniya();

  StatusBar1->SimpleText = "";

  //���������� ��������� �� �������� ���������
  UpdateValuta_I_Grivna();

  InsertLog("��������� �������� ��������� �� ������� ���������. ��������� "+obnov_kol+" �� "+ob_kol+" �������");

  StatusBar1->SimpleText = "";
}
//---------------------------------------------------------------------------

//�������������� ������
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

//���������� ������
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
  Label9->Caption="������� ������:";
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
  WinExec(("\""+ WordPath+"\"\""+ Path+"\\���������� ������������.doc\"").c_str(),SW_MAXIMIZE);
}
//---------------------------------------------------------------------------

void __fastcall TMain::N24Click(TObject *Sender)
{
  WinExec(("\""+ WordPath+"\"\""+ Path+"\\��������� ����������.doc\"").c_str(),SW_MAXIMIZE);
}
//---------------------------------------------------------------------------

void __fastcall TMain::N25Click(TObject *Sender)
{
 AnsiString Sql;
 double sum_dog=0, sum_obsh=0;
 int val=0,val1=0, kol_dog=0, kol_obsh, kust=0, kust1=0;

  StatusBar1->SimpleText = "������������ ������� ��������� ���������...";

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
      MessageBox(Handle,"�������� ������ ��� ������� ������ �� ������ ������� ��������� ��������� ��� ���","������",8202);
      StatusBar1->SimpleText = "";
      Abort();
    }


  if (DM->qObnovlenie->RecordCount>0)
    {
      //�������� �����, ���� �� �� ����������
      ForceDirectories(WorkPath);

      if (!rtf_Open ((TempPath + "\\v_yit2.txt").c_str()))
        {
          MessageBox(Handle,"������ �������� ����� ������","������",8192);
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
                          MessageBox(Handle,"������ ������ � ���� ������","������",8192);
                          if (!rtf_Close()) MessageBox(Handle,"������ �������� ����� ������","������",8192);
                          return;
                        }

                       DM->qObnovlenie->Next();
                       val1 = DM->qObnovlenie->FieldByName("kod_dogovora")->AsInteger;
                       kust1 = DM->qObnovlenie->FieldByName("ana")->AsInteger;
                    }

                  // ����� ����� �� ��������
                  switch (val)
                    {
                      case 0: rtf_Out("naim", "���������", 2);
                      break;
                      case 1: rtf_Out("naim", "��������(������)", 2);
                      break;
                      case 2: rtf_Out("naim", "��������(����)", 2);
                      break;
                      case 3: rtf_Out("naim", "�������", 2);
                      break;
                    }

                   rtf_Out("sumdog", sum_dog,10,2, 2);
                   rtf_Out("koldog", kol_dog, 2);

                   val = DM->qObnovlenie->FieldByName("kod_dogovora")->AsInteger;
 //                  kust = DM->qObnovlenie->FieldByName("ana")->AsInteger;
                   if(!rtf_LineFeed())
                     {
                       MessageBox(Handle,"������ ������ � ���� ������","������",8192);
                       if (!rtf_Close()) MessageBox(Handle,"������ �������� ����� ������","������",8192);
                       return;
                     }
                }

               // ����� ����� �� �����
               switch (kust)
                 {
                   case 1: rtf_Out("naim2", "�������������", 3);
                   break;
                   default: rtf_Out("naim2", "����������", 3);
                   break;

                 }

               rtf_Out("sumobsh", sum_obsh,10,2, 3);
               rtf_Out("kolobsh", kol_obsh, 3);
               kust = DM->qObnovlenie->FieldByName("ana")->AsInteger;
               if(!rtf_LineFeed())
                 {
                   MessageBox(Handle,"������ ������ � ���� ������","������",8192);
                   if (!rtf_Close()) MessageBox(Handle,"������ �������� ����� ������","������",8192);
                   return;
                 }

            }
          if(!rtf_Close())
            {
              MessageBox(Handle,"������ �������� ����� ������", "������", 8192);
              return;
            }

          int istrd;
          try
            {
              rtf_CreateReport(TempPath +"\\v_yit2.txt", Path+"\\RTF\\v_yit2.rtf",
                               WorkPath+"\\������� ��������� ��������� ��� ���.doc",NULL,&istrd);
              DeleteFile(TempPath+"\\v_yit2.txt");

              WinExec(("\""+ WordPath+"\"\""+WorkPath+"\\������� ��������� ��������� ��� ���.doc\"").c_str(),SW_MAXIMIZE);
            }
          catch(RepoRTF_Error E)
            {
              MessageBox(Handle,("������ ������������ ������:"+ AnsiString(E.Err)+
                                 "\n������ ����� ������:"+IntToStr(istrd)).c_str(),"������",8192);
            }
         StatusBar1->SimpleText = "";
        }
    }
  else
   {
     Application->MessageBox("��� ������ �� ������� �����", "��������������",
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
      case '0':  Label11->Caption="������";
      break;
      case '1':  Label11->Caption="������";
      break;
      case '2':  Label11->Caption="���������� ���� ��������";
      break;
      case '3':  Label11->Caption="������";
      break;
      case '4':  Label11->Caption="�����������";
      break;
      case '5':  Label11->Caption="�������������";
      break;
      case '6':  Label11->Caption="�������������";
      break;
      case '7':  Label11->Caption=" �� ������� � ����";
      break;
      default :  Label11->Caption="";
    }  }

}
//---------------------------------------------------------------------------

// �������� � Excel ��� ���
void __fastcall TMain::ExcelSAP(int valuta)
{

  AnsiString sFile, Sql, tn, tn1;
  int n=2;
  Variant AppEx, Sh;

  if (Application->MessageBox(("����� ��������� ������������ Excel-����� �� "+Mes[DM->mm-1]+" "+DM->yyyy+" ����.\n����������?").c_str(),"��������������",
                              MB_YESNO+MB_ICONINFORMATION)==ID_NO)
    {
      Abort();
    }

  StatusBar1->SimpleText=" ���� ������������ ����� � Excel...";

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
      Application->MessageBox(("�������� ������ ��� ��������� ������ �� ������� �� ��������� ��� ��������� �� ���������� (VU_859_N, SAP_SVED_UVOL, SAP_OSN_SVED)" + E.Message).c_str(),"������",
                              MB_OK+MB_ICONERROR);
      if (valuta==1) InsertLog("�������� ������ ��� ������������ ����� �� ��������� ��������� ��� ��� � Excel");
      else if (valuta==2) InsertLog("�������� ������ ��� ������������ ����� �� �������� ��������� ��� ��� � Excel");
      else if (valuta==3) InsertLog("�������� ������ ��� ������������ ����� �� ������� ��������� ��� ��� � Excel");
      else if (valuta==4) InsertLog("�������� ������ ��� ������������ ����� �� ���������� ��������� ��� ��� � Excel");

      StatusBar1->SimpleText="";
      Abort();
    }


  Cursor = crHourGlass;
  ProgressBar->Position = 0;
  ProgressBar->Visible = true;
  ProgressBar->Max=DM->qObnovlenie->RecordCount;

  // �������������� Excel, ��������� ���� ������
  try
    {
      AppEx=CreateOleObject("Excel.Application");
    }
  catch (...)
    {
      Application->MessageBox("���������� ������� Microsoft Excel!"
                              " �������� ��� ���������� �� ���������� �� �����������.","������",MB_OK+MB_ICONERROR);
      StatusBar1->SimpleText="";
      ProgressBar->Visible = false;
      Cursor = crDefault;
    }

  //���� ��������� ������ �� ����� ������������ ������
  try
    {
      try
        {
          AppEx.OlePropertySet("AskToUpdateLinks",false);
          AppEx.OlePropertySet("DisplayAlerts",false);

          //�������� �����, ���� �� �� ����������
          ForceDirectories(WorkPath);

          if (valuta==1)
            {
              DeleteFile(WorkPath+"\\SAP 0015��(������).xlsx");
              CopyFile((Path+"\\RTF\\sap.xlsx").c_str(), (WorkPath+"\\SAP 0015��(������).xlsx").c_str(), false);
              sFile = WorkPath+"\\SAP 0015��(������).xlsx";
            }
          else if (valuta==2)
            {
              DeleteFile(WorkPath+"\\SAP 0015��(������).xlsx");
              CopyFile((Path+"\\RTF\\sap.xlsx").c_str(), (WorkPath+"\\SAP 0015��(������).xlsx").c_str(), false);
              sFile = WorkPath+"\\SAP 0015��(������).xlsx";
            }
          else if (valuta==3)
            {
              DeleteFile(WorkPath+"\\SAP 0015��(������� ��������).xlsx");
              CopyFile((Path+"\\RTF\\sap.xlsx").c_str(), (WorkPath+"\\SAP 0015��(������� ��������).xlsx").c_str(), false);
              sFile = WorkPath+"\\SAP 0015��(������� ��������).xlsx";
            }
          else if (valuta==4)
            {
              DeleteFile(WorkPath+"\\SAP 0015��(���������� ��������).xlsx");
              CopyFile((Path+"\\RTF\\sap.xlsx").c_str(), (WorkPath+"\\SAP 0015��(���������� ��������).xlsx").c_str(), false);
              sFile = WorkPath+"\\SAP 0015��(���������� ��������).xlsx";
            }

          AppEx.OlePropertyGet("WorkBooks").OleProcedure("Open",sFile.c_str())    ;  //��������� �����, ������ � ���

          Sh=AppEx.OlePropertyGet("WorkSheets",1);                               //�������� � ��������� ����� �����
          //Sh=AppEx.OlePropertyGet("WorkSheets","������");                      //�������� ���� �� ������������
        }
      catch(...)
        {
          Application->MessageBox("������ �������� ����� Microsoft Excel!","������",MB_OK+MB_ICONERROR);
          StatusBar1->SimpleText="";
          ProgressBar->Visible = false;
          Cursor = crDefault;
          DecimalSeparator='.';
          if (valuta==1) InsertLog("�������� ������ ��� ������������ ����� �� ��������� ��������� ��� ��� � Excel");
          else if (valuta==2) InsertLog("�������� ������ ��� ������������ ����� �� �������� ��������� ��� ��� � Excel");
          else if (valuta==3) InsertLog("�������� ������ ��� ������������ ����� �� ������� ��������� ��� ��� � Excel");
          else if (valuta==4) InsertLog("�������� ������ ��� ������������ ����� �� ���������� ��������� ��� ��� � Excel");
        }

      int i=1;
      n=2;
      int d=-1;
      
      Variant Massiv;
      Massiv = VarArrayCreate(OPENARRAY(int,(0,13)),varVariant); //������ �� 11 ���������

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


          Sh.OlePropertyGet("Range", ("A" + IntToStr(n) + ":J" + IntToStr(n)).c_str()).OlePropertySet("Value", Massiv); //������ � ������� � ������ A �� ������ ��

          i++;
          n++;
          tn=DM->qObnovlenie->FieldByName("tn_sap")->AsString;

          DM->qObnovlenie->Next();
          ProgressBar->Position++;
          tn1=DM->qObnovlenie->FieldByName("tn_sap")->AsString;
        }

       // Sh.OlePropertyGet("Range", ("LQ18:LQ" + IntToStr(i-1)).c_str()).OlePropertySet("NumberFormat", "0.00");

      //����������� �����
   /*   Sh.OlePropertyGet("Range",("M18:M"+IntToStr(n-1)).c_str()).OlePropertyGet("Interior").OlePropertySet("Color",14277081);
      Sh.OlePropertyGet("Range",("P18:R"+IntToStr(n-1)).c_str()).OlePropertyGet("Interior").OlePropertySet("Color",14277081);

      Sh.OlePropertyGet("Range",("B18:K"+IntToStr(n-1)).c_str()).OlePropertyGet("Interior").OlePropertySet("Color",14408946);
      Sh.OlePropertyGet("Range",("N18:N"+IntToStr(n-1)).c_str()).OlePropertyGet("Interior").OlePropertySet("Color",14408946);
   */
      //������ �����
      Sh.OlePropertyGet("Range",("A2:J"+IntToStr(n-1)).c_str()).OlePropertyGet("Borders").OlePropertySet("LineStyle",1);

      //��������� ����� � ����� � ����� �� ��������
     // AnsiString vAsCurDir1=WorkPath+"\\������������ ������ �� �����������";

     // Sh.OleProcedure("SaveAs",vAsCurDir1.c_str());
     AppEx.OlePropertyGet("WorkBooks",1).OleFunction("Save");

      /* //������� �������� ���������� Excel
      AppEx.OleProcedure("Quit");
      AppEx = Unassigned;  */

      //������� ����� Excel � �������� ��� ������ ����������
     // AppEx.OlePropertyGet("WorkBooks",1).OleProcedure("Close");
      Application->MessageBox("����� ������� �����������!", "������������ ������",
                               MB_OK+MB_ICONINFORMATION);
      //AppEx.OlePropertyGet("WorkBooks").OleProcedure("Open",vAsCurDir1.c_str());
      AppEx.OlePropertySet("Visible",true);
      AppEx.OlePropertySet("AskToUpdateLinks",true);
      AppEx.OlePropertySet("DisplayAlerts",true);

      StatusBar1->SimpleText= "������������ ������ ���������.";

      Cursor = crDefault;
      ProgressBar->Position=0;
      ProgressBar->Visible = false;
      StatusBar1->SimpleText= "";

      if (valuta==1) InsertLog("��������� ������������ ����� �� ��������� ��������� ��� ��� � Excel. ���������� ������� = "+ IntToStr(DM->qObnovlenie->RecordCount));
      else if (valuta==2) InsertLog("��������� ������������ ����� �� �������� ��������� ��� ��� � Excel. ���������� ������� = "+ IntToStr(DM->qObnovlenie->RecordCount));
      else if (valuta==3) InsertLog("��������� ������������ ����� �� ������� ��������� ��� ��� � Excel. ���������� ������� = "+ IntToStr(DM->qObnovlenie->RecordCount));
      else if (valuta==4) InsertLog("��������� ������������ ����� �� ���������� ��������� ��� ��� � Excel. ���������� ������� = "+ IntToStr(DM->qObnovlenie->RecordCount));

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
      if (valuta==1) InsertLog("�������� ������ ��� ������������ ����� �� ��������� ��������� ��� ��� � Excel");
      else if (valuta==2) InsertLog("�������� ������ ��� ������������ ����� �� �������� ��������� ��� ��� � Excel");
      else if (valuta==3) InsertLog("�������� ������ ��� ������������ ����� �� ������� ��������� ��� ��� � Excel");
      else if (valuta==4) InsertLog("�������� ������ ��� ������������ ����� �� ���������� ��������� ��� ��� � Excel");

      Abort();
    }

 // if (otchet_zex==0) InsertLog("������������ ������ ���������� �� ����������� � Excel ������� ���������");
//  else InsertLog("������������ ������ ���������� ��  "+otchet_zex+" ���� � Excel ������� ���������");
  DecimalSeparator='.';

}
//---------------------------------------------------------------------------



//������������ ������� ��� SAP �� ������
void __fastcall TMain::N5Click(TObject *Sender)
{
  ExcelSAP(1);
}
//---------------------------------------------------------------------------

//������������ ������� ��� SAP �� ������
void __fastcall TMain::N8Click(TObject *Sender)
{
  ExcelSAP(2);
}
//---------------------------------------------------------------------------

//������������ ������� ��� SAP �� ������� ���������
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

//������������ ������ ��� ��������� �� SAP
void __fastcall TMain::OtchetStrahovaya(int valuta)
{
  AnsiString sFile, Sql, str,s;
  int otchet=0, kolzap=0, kolnzap=0;
  FILE *grn;

  //����� �����
  if(OpenDialog1->Execute())
    {
      // ������������� ���� � ����� �������
      sFile = OpenDialog1->FileName;
    }
  else
    {
      Abort();
    }

// �������������� Excel, ��������� ���� ������
  try
    {
      AppEx=CreateOleObject("Excel.Application");
    }
  catch (...)
    {
      Application->MessageBox("���������� ������� Microsoft Excel!"
                              " �������� ��� ���������� �� ���������� �� �����������.","������",MB_OK+MB_ICONERROR);
      StatusBar1->SimpleText="";
      InsertLog("�������� ������ ��� ������������ ���������� ����� ��� ��������� ��������");
      ProgressBar->Visible = false;
      Cursor = crDefault;
    }

  Cursor = crHourGlass;
  ProgressBar->Position = 0;
  ProgressBar->Visible = true;
  StatusBar1->SimpleText = "���� ������������ ������ ��� ��������� ��������...";

  //���� ��������� ������ �� ����� ������������ ������
  try
    {
      try
        {
          AppEx.OlePropertySet("AskToUpdateLinks",false);
          AppEx.OlePropertySet("DisplayAlerts",false);

          //�������� �����, ���� �� �� ����������
          ForceDirectories(WorkPath+"\\��� ��������� ��������");
          AppEx.OlePropertyGet("WorkBooks").OleProcedure("Open",sFile.c_str())    ;  //��������� �����, ������ � ���

          Sh=AppEx.OlePropertyGet("WorkSheets",1);                               //�������� � ��������� ����� �����
          //Sh=AppEx.OlePropertyGet("WorkSheets","������");                      //�������� ���� �� ������������
        }
      catch(...)
        {
          Application->MessageBox("������ �������� ����� Microsoft Excel!","������",MB_OK+MB_ICONERROR);
          StatusBar1->SimpleText="";
          InsertLog("�������� ������ ��� ������������ ���������� ����� ��� ��������� ��������");
          ProgressBar->Visible = false;
          Cursor = crDefault;
        }


      //���������� ����� � ���������
      int row = Sh.OlePropertyGet("UsedRange").OlePropertyGet("Rows").OlePropertyGet("Count");

      ProgressBar->Max=row;


      int i=1;

      if (valuta==1)
        {
          if ((grn=fopen((WorkPath+"\\��� ��������� ��������\\grivna.txt").c_str(),"wt"))==NULL)
            {
              ShowMessage("���� �� ������� �������");
              return;
            }
        }
      else if (valuta==2)
        {
          if ((grn=fopen((WorkPath+"\\��� ��������� ��������\\valuta.txt").c_str(),"wt"))==NULL)
            {
              ShowMessage("���� �� ������� �������");
              return;
            }
        }
      else if (valuta==3)
        {
          if ((grn=fopen((WorkPath+"\\��� ��������� ��������\\vneshnie.txt").c_str(),"wt"))==NULL)
            {
              ShowMessage("���� �� ������� �������");
              return;
            }
        }

      //����� ������, � ������� �������� �����
      while (String(AppEx.OlePropertyGet("Range", ("A"+IntToStr(i)).c_str())).IsEmpty() || !Proverka(String(AppEx.OlePropertyGet("Range", ("A"+IntToStr(i)).c_str()))))
        {
          i++;
          ProgressBar->Position++;

          if (i==row)
            {
              Application->MessageBox("���.� ��������� ������ ���������� � ������� � ����� Excel \n�� ������ �������� ���������.\n������� ��������� � Excel ���� � ��������� ��������.\n���� ������ ����� ��������� � � ���������� \n���������� � ������������","��������������",
                                      MB_OK+MB_ICONWARNING);
              Abort();
            }
        }

      //����� � ����
      while(!String(AppEx.OlePropertyGet("Range", ("A"+IntToStr(i)).c_str())).IsEmpty() && Proverka(String(AppEx.OlePropertyGet("Range", ("A"+IntToStr(i)).c_str()))))
        {

          //�������� �� ������������ ������������ ����� � ���� ����� � ���������
          if ((valuta==1 && String(AppEx.OlePropertyGet("Range", ("D"+IntToStr(i)).c_str()))!="7428")||
               (valuta==2 && String(AppEx.OlePropertyGet("Range", ("D"+IntToStr(i)).c_str()))!="7433") ||
               (valuta==3 && String(AppEx.OlePropertyGet("Range", ("D"+IntToStr(i)).c_str()))!="7434"))
            {
              if (valuta==1) s="���������� ��������";
              else if (valuta==2) s="��������� ��������";
              else if (valuta==3) s="�������� ��������";

              Application->MessageBox(("� ����������� ��������� Excel �������� ���� ������ � ������� 'D'="+String(AppEx.OlePropertyGet("Range", ("D"+IntToStr(i)).c_str()))+" \n�� �������� ����� ������ "+s+".\n�������� ������� ������ ���� Excel ��� ��������.\n�������� ������ ���� � ��������� ������������ ������.").c_str(), "������",
                                      MB_OK+MB_ICONWARNING);
              //������� �������� ���������� Excel
              AppEx.OleProcedure("Quit");
              StatusBar1->SimpleText="";
              ProgressBar->Visible = false;
              Cursor = crDefault;
              Abort();

            }

          //������� ������ �� ���.� �� ����������
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
              Application->MessageBox(("�������� ������ ��� ��������� ���������� �� ��������� �� ������ (SAP_OSN_SVED).\n���������� ��������� ������������! "+E.Message).c_str(), "������",
                                      MB_OK+MB_ICONERROR);
              //������� �������� ���������� Excel
              AppEx.OleProcedure("Quit");
              InsertLog("�������� ������ ��� ������������ ���������� ����� ��� ��������� ��������");
              StatusBar1->SimpleText="";
              ProgressBar->Visible = false;
              Cursor = crDefault;
              Abort();
            }

          if (DM->qObnovlenie->RecordCount==0)
            {
              //������� ������ �� ���������
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
                  Application->MessageBox(("�������� ������ ��� ��������� ���������� �� ��������� �� ������ (SAP_OSN_SVED).\n���������� ��������� ������������! "+E.Message).c_str(), "������",
                                           MB_OK+MB_ICONERROR);

                  //������� �������� ���������� Excel
                  AppEx.OleProcedure("Quit");
                  InsertLog("�������� ������ ��� ������������ ���������� ����� ��� ��������� ��������");
                  StatusBar1->SimpleText="";
                  ProgressBar->Visible = false;
                  Cursor = crDefault;
                  Abort();
                }

              float sum=0;
              if (DM->qObnovlenie->RecordCount==0)
                {
                  //��� �������, ������������ ������
                  if (otchet==0)
                    {
                      //�������� ����� ������ ����������� ���������, ���������� ������ � ���������� ����� � ���.�
                      if (!rtf_Open((TempPath + "\\otchet2.txt").c_str()))
                        {
                          MessageBox(Handle,"������ �������� ����� ������","������",8192);
                        }
                      // ����� ��������� � ����� �������
                      rtf_Out("data", Now(), 0);
                    }

                  rtf_Out("tn", String(AppEx.OlePropertyGet("Range", ("A"+IntToStr(i)).c_str())),1);
                  rtf_Out("sum", FloatToStrF(-1*Double(AppEx.OlePropertyGet("Range", ("E"+IntToStr(i)).c_str())),ffFixed,10,2) ,1);

                  if(!rtf_LineFeed())
                    {
                      MessageBox(Handle,"������ ������ � ���� ������","������",8192);
                      if (!rtf_Close()) MessageBox(Handle,"������ �������� ����� ������","������",8192);
                      return;
                    }

                  otchet=1;
                  i++;
                  kolnzap++;
                  ProgressBar->Position++;
                  continue;
                }
            }

          //����� ������
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
          StatusBar1->SimpleText = "���� ������������ ������ � ��������...";
          if(!rtf_Close())
            {
              MessageBox(Handle,"������ �������� ����� ������", "������", 8192);
              return;
            }
          int istrd;
          try
            {
              rtf_CreateReport(TempPath +"\\otchet2.txt", Path+"\\RTF\\otchet2.rtf",
                           WorkPath+"\\�������� ������ ��� ���������.doc",NULL,&istrd);


              WinExec(("\""+ WordPath+"\"\""+WorkPath+"\\�������� ������ ��� ���������.doc\"").c_str(),SW_MAXIMIZE);

            }
          catch(RepoRTF_Error E)
            {
              MessageBox(Handle,("������ ������������ ������:"+ AnsiString(E.Err)+
                                 "\n������ ����� ������:"+IntToStr(istrd)).c_str(),"������",8192);
            }

           StatusBar1->SimpleText = "";
        }
    }
  catch(Exception &E)
    {
      Application->MessageBox(("�������� ������ ��� ������������ ����� ��� ��������� ��������"+E.Message).c_str(),"������",
                              MB_OK+MB_ICONERROR);

      //��������� ����� ��������� � ��������� ���� "�������� ����..."
      AppEx.OlePropertySet("DisplayAlerts",false);
      //������� �������� ���������� Excel
      AppEx.OleProcedure("Quit");
      InsertLog("�������� ������ ��� ������������ ���������� ����� ��� ��������� ��������");

      StatusBar1->SimpleText="";
      ProgressBar->Visible = false;
      Cursor = crDefault;
      Abort();
    }

   //��������� ����� ��������� � ��������� ���� "�������� ����..."
  AppEx.OlePropertySet("DisplayAlerts",false);

  //������� �������� ���������� Excel
  AppEx.OleProcedure("Quit");

  if (kolnzap>0) str="�� ��������� ������� - "+IntToStr(kolnzap);
  else str="";

  Application->MessageBox(("������������ ��������� ������� =)\n���� ��� ��������� �������� ��������� � �����: "+WorkPath+"\\fkiev.txt.\n��������� ������� � ���� - "+kolzap+". "+str).c_str(),"������������ ������������",
                          MB_OK+MB_ICONINFORMATION);

  InsertLog("��������� ������������ ���������� ����� ��� ��������� ��������.\n ��������� ������� � ���� - "+IntToStr(kolzap)+". "+str);
  StatusBar1->SimpleText="";
  ProgressBar->Visible = false;
  Cursor = crDefault;
}
//---------------------------------------------------------------------------

//�������� ������ �� ����� ��������� �� ����������� �����������
void __fastcall TMain::N31Click(TObject *Sender)
{
  AnsiString Sql, Sql1, inn, nn;
  int i=1, rec=0;

  /*rec - ���������� ����������� � ������� �������*/


  im_fl=7;

  if (Application->MessageBox(("�� ������������� ������ ��������� ������ \n �� ����� ��������� �� ����������� ����������� �� " + Mes[DM->mm-1] + " " + DM->yyyy + " ����?").c_str(),
                               "�������� ������ �� ����� ���������",
                               MB_YESNO + MB_ICONINFORMATION) == IDNO)
    {
      Abort();
    }


  // �������� ������������ ��� � ������� ������� ������� � ���������
  ProverkaInfoExcel();

  StatusBar1->SimpleText = "";

  try
    {
      Sheet.OleProcedure("Activate");

      Main->Cursor = crHourGlass;
      StatusBar1->SimplePanel = true;    // 2 ������ �� StatusBar1
      StatusBar1->SimpleText=" ���� �������� ������...";

      ProgressBar->Visible = true;
      ProgressBar->Position = 0;
      ProgressBar->Max = Row;

      for ( i ; i<Row+1; i++)
        {
          nn = Excel.OlePropertyGet("Cells",i,1);
          inn = Excel.OlePropertyGet("Cells",i,5);

          ProgressBar->Position++;

          // ����� ����� ����������� ��� �������� �� Excel
          if (nn.IsEmpty() || !Proverka(nn) || inn.IsEmpty())  continue;

             //�������� �� ������� ��� ������������ ������� � ������� VU_859_N
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
                Application->MessageBox("������ ��������� ������ �� ������� �� ����������� 859 �/�","������",MB_OK+ MB_ICONERROR);
                Abort();
              }

            if (DM->qObnovlenie->RecordCount>0)
              {
                 if (Application->MessageBox(("������: ��� = "+ DM->qObnovlenie->FieldByName("zex")->AsString +
                                               ", ���.� = "+ DM->qObnovlenie->FieldByName("tn")->AsString +
                                               ", ��� = "+ DM->qObnovlenie->FieldByName("inn")->AsString +
                                               " � � �������� = "+DM->qObnovlenie->FieldByName("n_dogovora")->AsString +
                                              " ��� ����������. �������� �� ��� ���?").c_str(),"��������������",
                                              MB_YESNO + MB_ICONINFORMATION) ==ID_NO)
                    {
                       continue;
                    }
              }

            //���������� ���+�� �� sap_osn_sved
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
                Application->MessageBox("������ ��������� ������ �� �� ��������� �� ���������� (SAP_OSN_SVED, SAP_SVED_UVOL)","������",MB_OK+ MB_ICONERROR);
                Abort();
              }


            //������ ������ � ������� VU_859_N
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
                Application->MessageBox("������ ������� ������ � ������� �� ����������� 859 �/�","������",MB_OK+ MB_ICONERROR);
                Application->MessageBox("������ �� ���� ���������. ��������� ��������","������",MB_OK+ MB_ICONERROR);
                StatusBar1->SimpleText = "";

                Excel.OleProcedure("Quit");
                Abort();
             }
        }


      Application->MessageBox(("�������� ������ ��������� ������� =) \n ��������� " + IntToStr(rec) + " �������").c_str(),
                               "�������� ����� ��������� �� ����������� �����������",MB_OK+ MB_ICONINFORMATION);
      InsertLog("��������� �������� ������ �� ����� ��������� �� ����������� �����������. ��������� "+IntToStr(rec)+" �������");

      Excel.OleProcedure("Quit");
      Excel = Unassigned;

      StatusBar1->SimplePanel = false;
      ProgressBar->Visible = false;
      StatusBar1->SimpleText = "���������� ���������.";
      Main->Cursor = crDefault;
      StatusBar1->SimpleText = "";
    }
  catch(...)
    {
      Application->MessageBox("������ �������� ������ �� ����� ��������� �� ����������� �����������","������",MB_OK+ MB_ICONERROR);
      Excel.OleProcedure("Quit");

      StatusBar1->SimplePanel = false;
      ProgressBar->Visible = false;
      StatusBar1->SimpleText = "";
      Main->Cursor = crDefault;
    }
}
//---------------------------------------------------------------------------

//�������� ��������� �� ����������� �����������
void __fastcall TMain::N32Click(TObject *Sender)
{
  im_fl=7;
  
  if (Application->MessageBox(("�� ������������� ������ ��������� ��������� \n �� ��������� ����������� ����������� �� " + Mes[DM->mm-1] + " " + DM->yyyy + " ����?").c_str(),
                               "�������� ��������� �� ��������� ���������",
                               MB_YESNO + MB_ICONINFORMATION) == IDNO)
    {
      Abort();
    }

  // �������� �� ������������� ��� � �������
  ProverkaInfoExcelIzmeneniya();

  StatusBar1->SimpleText = "";

  //���������� ��������� �� ��������� ���������
  UpdateValuta_I_Grivna();

  InsertLog("��������� �������� ��������� �� ����������� �����������. ��������� "+obnov_kol+" �� "+ob_kol+" �������");

  StatusBar1->SimpleText = "";
}
//---------------------------------------------------------------------------



//������������ ������ ��� SAP �� ����������� �����������
void __fastcall TMain::N34Click(TObject *Sender)
{
  ExcelSAP(4);
}
//---------------------------------------------------------------------------

//������������ ��������� ������ �� ��������� ��������
void __fastcall TMain::N36Click(TObject *Sender)
{

  AnsiString sFile, sFile1, Sql, str,s;
  int otchet=0, kolzap=0, kolnzap=0, n=5, num=1;

  //����� �����
  if(OpenDialog1->Execute())
    {
      // ������������� ���� � ����� �������
      sFile = OpenDialog1->FileName;
    }
  else
    {
      Abort();
    }

// �������������� Excel, ��������� ���� � ������� �� ����� Excel �� SAP
  try
    {
      AppEx=CreateOleObject("Excel.Application");
    }
  catch (...)
    {
      Application->MessageBox("���������� ������� Microsoft Excel!"
                              " �������� ��� ���������� �� ���������� �� �����������.","������",MB_OK+MB_ICONERROR);
      StatusBar1->SimpleText="";
      InsertLog("�������� ������ ��� ������������ ���������� ����� ��� ��������� ��������");
      ProgressBar->Visible = false;
      Cursor = crDefault;
    }

  Cursor = crHourGlass;
  ProgressBar->Position = 0;
  ProgressBar->Visible = true;
  StatusBar1->SimpleText = "���� ������������ ������ ��� ��������� ��������...";

  //���� ��������� ������ �� ����� ������������ ������
  try
    {
      try
        {
          AppEx.OlePropertySet("AskToUpdateLinks",false);
          AppEx.OlePropertySet("DisplayAlerts",false);

          //�������� �����, ���� �� �� ����������
          ForceDirectories(WorkPath+"\\��� ��������� ��������");
          AppEx.OlePropertyGet("WorkBooks").OleProcedure("Open",sFile.c_str())    ;  //��������� �����, ������ � ���

          Sh=AppEx.OlePropertyGet("WorkSheets",1);                               //�������� � ��������� ����� �����
          //Sh=AppEx.OlePropertyGet("WorkSheets","������");                      //�������� ���� �� ������������
        }
      catch(...)
        {
          Application->MessageBox("������ �������� ����� Microsoft Excel!","������",MB_OK+MB_ICONERROR);
          StatusBar1->SimpleText="";
          InsertLog("�������� ������ ��� ������������ ���������� ����� ��� ��������� ��������");
          ProgressBar->Visible = false;
          Cursor = crDefault;
        }


      //���������� ����� � ���������
      int row = Sh.OlePropertyGet("UsedRange").OlePropertyGet("Rows").OlePropertyGet("Count");

      ProgressBar->Max=row;

      int i=1;

      // �������������� Excel, ��������� ���� ��������� ������, ���� ����� ���������� ������ �� ����� Excel �� SAP
      try
         {
           AppEx1=CreateOleObject("Excel.Application");
         }
      catch (...)
         {
           Application->MessageBox("���������� ������� Microsoft Excel!"
                                   " �������� ��� ���������� �� ���������� �� �����������.","������",MB_OK+MB_ICONERROR);
           StatusBar1->SimpleText="";
           ProgressBar->Visible = false;
           Cursor = crDefault;
         }

      //���� ��������� ������ �� ����� ������������ ������
      try
        {
          try
            {
              AppEx1.OlePropertySet("AskToUpdateLinks",false);
              AppEx1.OlePropertySet("DisplayAlerts",false);

              //�������� �����, ���� �� �� ����������
              ForceDirectories(WorkPath);

              DeleteFile(WorkPath+"\\�������� ����� �� ���������� ���������.xlsx");
              CopyFile((Path+"\\RTF\\itog_pens.xlsx").c_str(), (WorkPath+"\\�������� ����� �� ���������� ���������.xlsx").c_str(), false);
              sFile1 = WorkPath+"\\�������� ����� �� ���������� ���������.xlsx";


              AppEx1.OlePropertyGet("WorkBooks").OleProcedure("Open",sFile1.c_str())    ;  //��������� �����, ������ � ���

              Sh=AppEx1.OlePropertyGet("WorkSheets",1);                               //�������� � ��������� ����� �����
            }
          catch(...)
            {
              Application->MessageBox("������ �������� ����� Microsoft Excel!","������",MB_OK+MB_ICONERROR);
              StatusBar1->SimpleText="";
              ProgressBar->Visible = false;
              Cursor = crDefault;
              DecimalSeparator='.';
              InsertLog("�������� ������ ��� ������������ ��������� ������ �� ���������� ��������� � Excel");
            }


      //����� ������, � ������� �������� ����� ������ � ����� SAP
      while (String(AppEx.OlePropertyGet("Range", ("A"+IntToStr(i)).c_str())).IsEmpty() || !Proverka(String(AppEx.OlePropertyGet("Range", ("A"+IntToStr(i)).c_str()))))
        {
          i++;
          ProgressBar->Position++;

          if (i==row+1)
            {
              Application->MessageBox("���.� ��������� ������ ���������� � ������� � ����� Excel \n�� ������ �������� ���������.\n������� ��������� � Excel ���� � ��������� ��������.\n���� ������ ����� ��������� � � ���������� \n���������� � ������������","��������������",
                                      MB_OK+MB_ICONWARNING);
              Abort();
            }
        }



      Variant Massiv;
      Massiv = VarArrayCreate(OPENARRAY(int,(0,8)),varVariant); //������ �� 11 ���������


     // AppEx1.OlePropertySet("Visible",true);

      //����� � ����
      while(!String(AppEx.OlePropertyGet("Range", ("A"+IntToStr(i)).c_str())).IsEmpty() && Proverka(String(AppEx.OlePropertyGet("Range", ("A"+IntToStr(i)).c_str()))))
        {

          //�������� �� ������������ ������������ ����� � ���� ����� � ���������
          if (String(AppEx.OlePropertyGet("Range", ("D"+IntToStr(i)).c_str()))!="7428")
            {
              s="����������� ��������";


              Application->MessageBox(("� ����������� ��������� Excel �������� ���� ������ � ������� 'D'="+String(AppEx.OlePropertyGet("Range", ("D"+IntToStr(i)).c_str()))+" \n�� �������� ����� ������ "+s+".\n�������� ������� ������ ���� Excel ��� ��������.\n�������� ������ ���� � ��������� ������������ ������.").c_str(), "������",
                                      MB_OK+MB_ICONWARNING);
              //������� �������� ���������� Excel
              AppEx.OleProcedure("Quit");
              StatusBar1->SimpleText="";
              ProgressBar->Visible = false;
              Cursor = crDefault;
              Abort();

            }

          //������� ������ �� ���.� �� ���������� � ��������� � ���������� ������ �������� �� ������� �� �����������
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
              Application->MessageBox(("�������� ������ ��� ��������� ���������� �� ��������� �� ������ (SAP_OSN_SVED).\n���������� ��������� ������������! "+E.Message).c_str(), "������",
                                      MB_OK+MB_ICONERROR);
              //������� �������� ���������� Excel
              AppEx.OleProcedure("Quit");
              InsertLog("�������� ������ ��� ������������ ������ �� ����������� �����������");
              StatusBar1->SimpleText="";
              ProgressBar->Visible = false;
              Cursor = crDefault;
              Abort();
            }

          if (DM->qObnovlenie->RecordCount==0)
            {
              //��� �������, ������������ ������
              if (otchet==0)
                {
                  //�������� ����� ������ ����������� ���������, ���������� ������ � ���������� ����� � ���.�
                  if (!rtf_Open((TempPath + "\\otchet2.txt").c_str()))
                    {
                      MessageBox(Handle,"������ �������� ����� ������","������",8192);
                    }
                  // ����� ��������� � ����� �������
                  rtf_Out("data", Now(), 0);
                }

              rtf_Out("tn", String(AppEx.OlePropertyGet("Range", ("A"+IntToStr(i)).c_str())),1);
              rtf_Out("sum", FloatToStrF(-1*Double(AppEx.OlePropertyGet("Range", ("E"+IntToStr(i)).c_str())),ffFixed,10,2) ,1);

              if(!rtf_LineFeed())
                {
                  MessageBox(Handle,"������ ������ � ���� ������","������",8192);
                  if (!rtf_Close()) MessageBox(Handle,"������ �������� ����� ������","������",8192);
                  return;
                }

              otchet=1;
              i++;
              kolnzap++;
              ProgressBar->Position++;
              continue;
            }

          //����� ���� ������������ ������
          Sh.OlePropertyGet("Range", "E2").OlePropertySet("Value", ("�� "+Mes[DM->mm-1] + " " + DM->yyyy+" ����").c_str());


          //����� ������
          Massiv.PutElement(num, 0);
          Massiv.PutElement(DM->qObnovlenie->FieldByName("numident")->AsString.c_str(), 1);
          Massiv.PutElement(DM->qObnovlenie->FieldByName("n_dogovora")->AsString.c_str(), 2);
          Massiv.PutElement(DM->qObnovlenie->FieldByName("n_ind_schet")->AsString.c_str(), 3);
          Massiv.PutElement(DM->qObnovlenie->FieldByName("fam_ukr")->AsString.c_str(), 4);
          Massiv.PutElement(DM->qObnovlenie->FieldByName("im_ukr")->AsString.c_str(), 5);
          Massiv.PutElement(DM->qObnovlenie->FieldByName("ot_ukr")->AsString.c_str(), 6);
          //Massiv.PutElement(FloatToStrF(Double(AppEx.OlePropertyGet("Range", ("E"+IntToStr(i)).c_str())), ffFixed, 20,2), 7);
          Massiv.PutElement(-1*Double(AppEx.OlePropertyGet("Range", ("E"+IntToStr(i)).c_str())), 7);

          Sh.OlePropertyGet("Range", ("A" + IntToStr(n) + ":H" + IntToStr(n)).c_str()).OlePropertySet("Value", Massiv); //������ � ������� � ������ A �� ������ ��

          num++;
          i++;
          n++;
          kolzap++;
          ProgressBar->Position++;
        }


       //����� ������ � �������
       Sh.OlePropertyGet("Range", ("A" + IntToStr(n)).c_str()).OlePropertySet("Value", "�������� ���� �� ������� (���):");
       Sh.OlePropertyGet("Range", ("H" + IntToStr(n)).c_str()).OlePropertySet("Formula", ("=����(H5:H"+IntToStr(n-1)).c_str());


       //.OlePropertyGet("Offset", n)
       //������ �����
       Sh.OlePropertyGet("Range",("A"+IntToStr(n)).c_str()).OlePropertyGet("Font").OlePropertySet("Bold",true);
       Sh.OlePropertyGet("Range",("H"+IntToStr(n)).c_str()).OlePropertyGet("Font").OlePropertySet("Bold",true);

       //���������� �����
       // Sh.OlePropertyGet("Range",("A"+IntToStr(n)).c_str()).OlePropertySet("WrapText",true);

       //������������ �� �����������
       Sh.OlePropertyGet("Range",("A"+IntToStr(n)).c_str()).OlePropertySet("HorizontalAlignment", 1); //��������� �� ���.

       //�����
       Sh.OlePropertyGet("Range",("A5:H"+IntToStr(n)).c_str()).OlePropertyGet("Borders").OlePropertySet("LineStyle", 1);


      if (otchet==1)
        {
          StatusBar1->SimpleText = "���� ������������ ������ � ��������...";
          if(!rtf_Close())
            {
              MessageBox(Handle,"������ �������� ����� ������", "������", 8192);
              return;
            }
          int istrd;
          try
            {
              rtf_CreateReport(TempPath +"\\otchet2.txt", Path+"\\RTF\\otchet2.rtf",
                           WorkPath+"\\������ ��� ������������ ��������� ������ �� ���������� ���������.doc",NULL,&istrd);


              WinExec(("\""+ WordPath+"\"\""+WorkPath+"\\������ ��� ������������ ��������� ������ �� ���������� ���������.doc\"").c_str(),SW_MAXIMIZE);

            }
          catch(RepoRTF_Error E)
            {
              MessageBox(Handle,("������ ������������ ������:"+ AnsiString(E.Err)+
                                 "\n������ ����� ������:"+IntToStr(istrd)).c_str(),"������",8192);
            }

           StatusBar1->SimpleText = "";
        }

        }
  catch(Exception &E)
    {
      Application->MessageBox(("�������� ������ ��� ������������ ��������� ������ �� ���������� ��������� �����������"+E.Message).c_str(),"������",
                              MB_OK+MB_ICONERROR);

      //��������� ����� ��������� � ��������� ���� "�������� ����..."
      AppEx1.OlePropertySet("DisplayAlerts",false);
      //������� �������� ���������� Excel
      AppEx1.OleProcedure("Quit");
      InsertLog("�������� ������ ��� ������������ ��������� ������ �� ���������� ��������� �����������");

      StatusBar1->SimpleText="";
      ProgressBar->Visible = false;
      Cursor = crDefault;
      Abort();
    }



    }
  catch(Exception &E)
    {
      Application->MessageBox(("�������� ������ ��� ������������ ��������� ������ �� ���������� ��������� �����������"+E.Message).c_str(),"������",
                              MB_OK+MB_ICONERROR);

      //��������� ����� ��������� � ��������� ���� "�������� ����..."
      AppEx.OlePropertySet("DisplayAlerts",false);
      //������� �������� ���������� Excel
      AppEx.OleProcedure("Quit");
      InsertLog("�������� ������ ��� ������������ ��������� ������ �� ���������� ��������� �����������");

      StatusBar1->SimpleText="";
      ProgressBar->Visible = false;
      Cursor = crDefault;
      Abort();
    }

   //��������� ����� ��������� � ��������� ���� "�������� ����..."
  AppEx.OlePropertySet("DisplayAlerts",false);
  //������� �������� ���������� Excel � ������ �� SAP
  AppEx.OleProcedure("Quit");


   AppEx1.OlePropertyGet("WorkBooks",1).OleFunction("Save");
   AppEx1.OlePropertySet("Visible",true);
   AppEx1.OlePropertySet("AskToUpdateLinks",true);
   AppEx1.OlePropertySet("DisplayAlerts",true);



  if (kolnzap>0) str="�� ��������� ������� - "+IntToStr(kolnzap);
  else str="";

  Application->MessageBox("������������ ��������� ������� =)","������������ ��������� ������",
                          MB_OK+MB_ICONINFORMATION);

  InsertLog("��������� ������������ ��������� ������ �� ���������� ��������� �� �����������.\n ���������� ������� - "+IntToStr(kolzap)+". "+str);
  StatusBar1->SimpleText="";
  ProgressBar->Visible = false;
  Cursor = crDefault;
}
//---------------------------------------------------------------------------

