/*****************************************************************

RepoRTFM : ������ ���������� ������� � ������� RTF.

�����  : A.PL. simple
E-mail : aplmail@box.vsi.ru
WWW    : http://www.vsi.ru/~apl
         http://aplsimple.boom.ru   (����� ����� ������ ����� ������ ������)

��������� �������� ��. � _usergid.rtf � � _read_me.txt

*****************************************************************/

#include <vcl.h>
#include <stdlib.h>
#include <stdio.h>
#include <algorithm>
using std::min;
using std::max;
#include "RepoRTFM.h"
#pragma hdrstop


#pragma package(smart_init)

namespace RepoRTF_innerface {

//����������
AnsiString rtf_asSpec;
int rtf_lenSpec; //� ��� �����

bool rtf_RTFpic;

//������.���-�� �������
const rtf_maxLev=999;

//������.����� ������ ��� ������/������
const rtf_iSSize=4096;

//���� �����
const rtf_LevMarker=2;
const rtf_ParMarker=3;
AnsiString rtf_asParMarker, rtf_asLevMarker;

//������� ����
FILE *rtf_Instream;
//�������� ����
FILE *rtf_Outstream;

//���������� ������� ����������� ���������� RepoRTF_innerface

void __fastcall rtf_SpecSymbols(
  AnsiChar acSpec, AnsiString asRepl, AnsiString* asParvalue);

void __fastcall rtf_GetParameter(
  AnsiString* asDat, AnsiString* asParname,
  AnsiString* asParvalue, int* curlev);

void __fastcall rtf_StrIntoReport(
  AnsiString* asStr, int* kz);

void __fastcall rtf_ParIntoReport(AnsiString* rtf_StrOut,
  AnsiString asParname, AnsiString asParvalue);

void __fastcall rtf_CloseFiles(char* szErr);

AnsiString __fastcall rtf_SymbCode(AnsiString asSymb);

AnsiString __fastcall rtf_Marker(int lev);

char rtf_ErrStruc[] =
  "������ � ��������� ����� ������� �/��� ������";
}

using namespace RepoRTF_innerface;

AnsiString rtf_SubstOld="\\fcharset0", rtf_SubstNew="\\fcharset204";

//========================================================
//��������� ��������� ����.�������� ������� rtf �
//�������� ��������� (��� ������ ���� �������� ������)

//������� ���������:
//acSpec - ����.������
//asRepl - ��� ������

//�������/�������� ���������:
//asParvalue- �������� ���������

void __fastcall RepoRTF_innerface::rtf_SpecSymbols(
  AnsiChar acSpec, AnsiString asRepl, AnsiString* asParvalue)
{
  for (int ib=1; ib<=(*asParvalue).Length(); ib++)
    if ((*asParvalue)[ib]==acSpec)
      *asParvalue = (*asParvalue).SubString(1,ib-1)+asRepl+
        (*asParvalue).SubString(ib+1,(*asParvalue).Length());
}

//========================================================
//��������� ��������� �������� ��������� �� ������ ������

//������� ���������:
//asDat - ������ ������

//�������� ���������:
//asParname - ��� ���������
//asParvalue - �������� ���������
//curlev - ������� ������ ������

void __fastcall RepoRTF_innerface::rtf_GetParameter(
  AnsiString* asDat, AnsiString* asParname,
  AnsiString* asParvalue, int* curlev)
{
int ib, ilen=(*asDat).Length();
  *curlev = 0;
  *asParname = "";
  ib = (*asDat).AnsiPos(":");
  if (ib<2)  return;
AnsiString asTmp=(*asDat).SubString(1,ib);
  if (asTmp[1]!='|') return;
   *curlev = atoi(asTmp.SubString(2,ib).c_str());
//����������� �� ������ ����.��������-������������ ����������,
//��������,"|1:"
 *asDat = (*asDat).SubString(ib+1,ilen);
//���� ����� ����� ���������
 ib = (*asDat).AnsiPos(":");
 if (ib>1)
   {
//�������� ��� ���������
   *asParname = (*asDat).SubString(1,ib-1);
   bool boolSpec = (*asParname).AnsiPos("_SYM")==1;
   *asParname = rtf_asParMarker+(*asParname)+rtf_asParMarker;
//���� ���������
   *asDat = (*asDat).SubString(ib+1,ilen);
   ib=(*asDat).AnsiPos(asTmp);
//�������� �������� ���������
   if (ib)
     {
     *asParvalue = (*asDat).SubString(1,ib-1);
     *asDat = (*asDat).SubString(ib,ilen);
     }
   else
     *asParvalue = *asDat;
//��� ����� ���� "����������" �������� ����� �������� (������ Symbol)
   if (boolSpec)
     {
     asTmp = "{\\f3\\lang1033\\langfe1049\\langnp1033 ";
     for (int i=1;i<=(*asParvalue).Length();i++)
       asTmp = asTmp+
         rtf_SymbCode((*asParvalue).SubString(i,1));
     *asParvalue = asTmp+" }";
     return;  //��������� �������� �� ���������
     }
   else
//���� ������ ����.������� ����������� ���� ������� rtf
     if (rtf_RTFpic)
       {
       rtf_SpecSymbols('\\',"\\'5C",asParvalue);
       rtf_SpecSymbols('{', "\\'7B",asParvalue);
       rtf_SpecSymbols('}', "\\'7D",asParvalue);
       }
//���� �������� �������� ���������� |EOL:,
// ������� � ��� ������� �����,
//���� �������� �������� ���������� |EOP:,
// ������� � ��� ������� �������
   for (ib=1; ib;)
     {
     ib = (*asParvalue).AnsiPos("|EOL:");
     if (ib==0)
      {
      ib = (*asParvalue).AnsiPos("|EOP:");
      if (ib)
        *asParvalue = (*asParvalue).SubString(1,ib-1)+
          (rtf_RTFpic?"{\\page}":"\014")+
          (*asParvalue).SubString(ib+5,ilen);
       }
     else
       *asParvalue = (*asParvalue).SubString(1,ib-1)+
         (rtf_RTFpic?"{\\par}":"\n")+
         (*asParvalue).SubString(ib+5,ilen);
     }
   }
 else
   *asParname = "";
}

//========================================================
//��������� �������� �������� ������ � ���� ������

//������� ���������:
//asStr - �������������� ������

//�������� ���������:
//kz - ��� ����������

void __fastcall RepoRTF_innerface::rtf_StrIntoReport(
  AnsiString* asStr, int* kz)
{
int ib,ib1,ib2,ib3;
unsigned char ch;
//������� �� �������� ������ ��� ������������� ��������� � ������
//(��� ����� ���� � �����������)
  for (ib=1; ib<=(ib1=(*asStr).Length()); )
    {
    ch = (*asStr)[ib];
    if (((ch==rtf_LevMarker))||(ch==rtf_ParMarker))
      {
      if (ch==rtf_LevMarker)
        ib3=4;
      else
        ib3=1;
      ib2 = ((*asStr).SubString(ib+1,ib1)).AnsiPos(
        (*asStr).SubString(ib,ib3));
      (*asStr) = (*asStr).SubString(1,ib-1)+
        (*asStr).SubString(ib+ib2+ib3,ib1);
      }
    else
      ib++;
    }
//����� �������� ������ � ���� ������
 for (ib=1; (*kz==0)&&ib;)
   {
   ib = (*asStr).Length();
   if (ib)
     {
     ib1 = min(rtf_iSSize,ib);
     ib2 = fwrite((*asStr).SubString(1,rtf_iSSize).c_str(),
           ib1,1,rtf_Outstream);
     if (ib2==0) *kz=6;
     (*asStr) = (*asStr).SubString(rtf_iSSize+1,ib);
     }
   }
 (*asStr) = "";
}

//========================================================
//��������� ������ �������� ��������� � �������� ������

//������� ���������:
//rtf_StrOut - �������������� ������
//asParname  - �������� ���������
//asParvalue - �������� ���������

//�������� ���������:
//rtf_StrOut - �������������� ������

void __fastcall RepoRTF_innerface::rtf_ParIntoReport(
  AnsiString* rtf_StrOut, AnsiString asParname,
  AnsiString asParvalue)
{
//����������� ��������� ������ ����������� � ������� -
//���� �������� � ������ (��� � test8)
   for (int ib=1; ib;)
     if ((ib=(*rtf_StrOut).AnsiPos(asParname))>0)
       (*rtf_StrOut) = (*rtf_StrOut).SubString(1,ib-1)+
        asParvalue+(*rtf_StrOut).SubString(ib+
        asParname.Length(),(*rtf_StrOut).Length());
}

//========================================================
// �������� ������

void __fastcall RepoRTF_innerface::rtf_CloseFiles(char* szErr)
{
 fclose(rtf_Instream);
 fclose(rtf_Outstream);
 if (strlen(szErr)) throw RepoRTF_Error(szErr);
}

//========================================================
// ������ ������

AnsiString __fastcall RepoRTF_innerface::rtf_Marker(int lev)
{
  return rtf_asLevMarker+(AnsiString(lev)+"**").SubString(1,3);
}

//========================================================
// ��������� ���� ������� � 16-������ ����

AnsiString __fastcall RepoRTF_innerface::rtf_SymbCode
 (AnsiString asSymb)
{
int ic=WideChar(asSymb[1])%256;
  return "\\'"+AnsiString::IntToHex(ic,2);
}

//========================================================
//��������� �������� ������ �� ������ ������ � �������

void __fastcall RepoRTF_interface::rtf_CreateReport
(
const AnsiString asDatFileName, //��� ����� ������
const AnsiString asPicFileName, //��� ����� �������
const AnsiString asRepFileName, //��� ����� ������
callingProc cProc, //��� ��������� ����������
int* istrd         //����� ������� ������ � ����� ������
)
{
AnsiString rtf_StrInp;   //�������� ������
AnsiString rtf_StrOut;   //�������� ������
char szIO[rtf_iSSize+1]; //������ ��� ��/���
char *pos;               //����� ������� ������� � ������
AnsiString asDat;        //������ ����� ������
AnsiString asParname;    //��� ���������
AnsiString asParvalue;   //�������� ���������
int curlev;            //����� �������� ������ ������
int prevlev=0;             //����� ����������� ������ ������
int kz=0;                //��� ����������
bool boolDOSdata;        //���� "������ � ��������� DOS"
bool boolLev;            //���� "������ �������"

AnsiString asTmp;        //������� ����������
int ib, ib0, ib1, ib2, ilen, ilen2;
unsigned char ch;

  if (asPicFileName==asRepFileName)
   throw RepoRTF_Error("����� ������ ������� � ������ ���������.");

  *istrd=0;
  rtf_asParMarker = AnsiString(char(rtf_ParMarker));
  rtf_asLevMarker = AnsiString(char(rtf_LevMarker));
  rtf_asSpec="#";
  rtf_lenSpec=1;

//�������� ������
//�������� ������� ���� ������� � ��������� ���
 rtf_Instream = fopen(asPicFileName.c_str(),"r");
 if (rtf_Instream==NULL)
   throw RepoRTF_Error("�� ������ ���� �������");
 rtf_StrInp = "";

//������ ����� ������� � �������� ������
 while (!feof(rtf_Instream))
   {
   ib = fread(szIO,1,rtf_iSSize,rtf_Instream);
   szIO[ib]=0;
   rtf_StrInp = rtf_StrInp+AnsiString(szIO);
   }

//�������� ������� ���� ������
 rtf_Outstream = fopen(asRepFileName.c_str(),"w");
 if (rtf_Outstream==NULL)
   {
   fclose(rtf_Instream);
   throw RepoRTF_Error("�� ������ ���� ������.\n\
�������� �� ������ ������ �����.");
   }
 fclose(rtf_Instream);

//�������� ������� ���� ������
 rtf_Instream = fopen(asDatFileName.c_str(),"r");
 if (rtf_Instream==NULL)
   throw RepoRTF_Error("�� ������ ���� ������");

//�������� ������� �� RTF - �����, ���� ����� �����������
//������ � ��������� �������
 rtf_RTFpic = (rtf_StrInp.AnsiPos("{\\rtf")==1);

//������� 16.11.2002 :
//�������� ��������� ������ �� ����� (��������, ����� �������� -
//�� ���� �� ����?)
 asParname  = rtf_SubstOld;
 asParvalue = rtf_SubstNew;
 for (;asParname.Length();)
   {
   ib = asParname.AnsiPos(rtf_asSpec);
   if (!ib)
    ib = ib0 = rtf_iSSize;
   else
    ib0 = asParvalue.AnsiPos(rtf_asSpec);
   asTmp = asParname.SubString(1,ib-1);
   asParname = asParname.SubString(ib+rtf_lenSpec,rtf_iSSize);
   asDat = asParvalue.SubString(1,ib0-1);
   asParvalue = asParvalue.SubString(ib0+rtf_lenSpec,rtf_iSSize);
   if (asDat!=asTmp)
    while ((ib=rtf_StrInp.AnsiPos(asTmp))>0)
     rtf_StrInp = rtf_StrInp.SubString(1,ib-1)+asDat+
       rtf_StrInp.SubString(ib+asTmp.Length(),rtf_StrInp.Length());
   }

//���������� �������� ������
 for (ib=ib0=1; ib;)
   {
   ilen = rtf_StrInp.Length();
   ib = rtf_StrInp.SubString(ib0,ilen).AnsiPos(rtf_asSpec);
   if (ib)
     {
     ib=ib0=ib0+ib-1;
//����� # - ������ # ������ ����, ����� ������
     ib1 = rtf_StrInp.SubString(ib+rtf_lenSpec,ilen).AnsiPos(rtf_asSpec);
     if (!ib1)
       rtf_CloseFiles(rtf_ErrStruc);
//�������� ��� ���������
     asTmp = rtf_StrInp.SubString(ib+rtf_lenSpec,ib1-1);
     if (asTmp.AnsiPos("FONT=")==1) //�������� �����
        {
        asTmp = asTmp.SubString(6,rtf_iSSize);
        AddFontResource(asTmp.c_str());
        }
     ilen2 = asTmp.Length();
//����� �������� : � ��� ����� ������ ���.����������� �����,
//������� ����� � �.�. - �� ��� ���� ������ (� �����������
//���������� �����)
     ib2 = ib1+ib+rtf_lenSpec*2-1;
     if (ilen2)
      {
      if (asTmp[1]=='@') // ������ �����������
        {
        rtf_asSpec = asTmp.SubString(2,ilen2);
        rtf_lenSpec = rtf_asSpec.Length();
        rtf_StrInp = rtf_StrInp.SubString(1,ib-1)+
         rtf_StrInp.SubString(ib1+ib+ilen2-rtf_lenSpec,ilen);
        continue;
        }
//������� \' ��� ������� ���� (� ������������) �������� �����
      if (asTmp.AnsiPos("\\'")==0)
       for (ib1=1; ib1<=ilen2;)
        {
        ch = asTmp[ib1];
        if (!(isalnum(ch)||(ch=='_')))  //���������-�������� ������?
         {                         //�� �� :
         if ((ch=='{')||(ch=='\\'))
           {
           ilen2= asTmp.AnsiPos(" ");
           if (ilen2==0)
             ilen2= asTmp.AnsiPos("}");//������ ���������
           if (ilen2==0) ilen2=ib1; //���� ���� (27.01.03)
           }
         else
           ilen2 = ib1;             //��� ���� ������
         asTmp = asTmp.SubString(1,ib1-1)+
           asTmp.SubString(ilen2+1,rtf_iSSize);
         ilen2 = asTmp.Length();
         }
        else
         ib1++;
        }
//������ #D1#,#D2#...,#8D#,#9D# � �������� ������ ������� ������� �
//������ rtf_LevMarker+1,2,... ����� ������� �����������)
      ib1=asTmp.Length();
      boolLev=ib1;
      for (int c=0,i=1; i<=ib1; i++)
        {
        asParvalue=asTmp.SubString(i,1);
        if (asParvalue=="D")
          {c++; if ((c>1)||((i>1)&&(i<ib1))) {boolLev=false; break;}}
        else
          if ((asParvalue<"0")||(asParvalue>"9")) {boolLev=false; break;}
        }
      if (boolLev)
        {
//���������� �� D1 � 1D, D2 � 2D � �.�. �� ����������� (�����
//���� �� �������� � ������ 1D-1D,2D-2D,...) � ���� � ����
//����� ����������� �� ������������� ���������, �� �������
//��� ����������� � �������� �������� ����� �����������
        if (asTmp[1]=='D')
//������ �����
         curlev = atoi(asTmp.SubString(2,7).c_str());
        else
//����� �����
         curlev = atoi(asTmp.c_str());
        asTmp = rtf_StrInp.SubString(ib2,ilen);
//������� ����� ����� � ������ ������ '}', ���������� �� \par
        if (rtf_RTFpic)
          {
          ib2 = asTmp.AnsiPos("}");
          if (ib2==1)
            {
            asTmp = asTmp.SubString(2,ilen);
            ib2 = asTmp.AnsiPos("}");
            }
          ib1 = asTmp.AnsiPos(rtf_asSpec );
          if (ib1&&(ib1<ib2)) ib2 = ib1;
          }
        else
          ib2 = 1;
        if (curlev>0)
          asParvalue = rtf_Marker(curlev);
        else
          asParvalue = "";
        rtf_StrInp = rtf_StrInp.SubString(1,ib-1)+asParvalue+
          asTmp.SubString(ib2,ilen);
        continue;
        }
      else
//������ # � �������� ������ ������� ������� � �����
//����� ����������
        {
//��������: �������� �� ����������� ��
        if (asTmp=="DATE")// ������� ����
         asParvalue = FormatDateTime("dd.mm.yyyy",Date());
        else
        if (asTmp=="TIME")// ������� �����
         asParvalue = FormatDateTime("hh:mm",Time());
        else
        if (asTmp=="EOP") // ����� ��������
         asParvalue = (rtf_RTFpic?"{\\page}":"\014");
        else
         asParvalue = rtf_asParMarker+asTmp+rtf_asParMarker;
        rtf_StrInp = rtf_StrInp.SubString(1,ib-1)+
         asParvalue+rtf_StrInp.SubString(ib2,ilen);
        }
      }
     else
//���������� �� ## - ������� ������ ���� #
      rtf_StrInp = rtf_StrInp.SubString(1,ib-1)+
         rtf_SymbCode(rtf_asSpec)+rtf_StrInp.SubString(ib2,ilen);
     }
   }

//���� ��������� ����� ������
 rtf_StrOut = "";
 boolDOSdata = false;
 while (!feof(rtf_Instream))
   {
//������ ������ ����� ������
   pos = NULL;
   for (asDat=""; pos==NULL;)
     {
     szIO[0]=0;
//������ ����� ���� � �������, ��� rtf_iSSize
     fgets(szIO,rtf_iSSize,rtf_Instream);
     if (strlen(szIO)<1) break;
     pos = strstr(szIO,"\n");
     if (pos!=NULL) *pos = 0;
//���� � ����� ������ ������ ��������� DOS, ��������� ������ �
//��������� Windows
     if (boolDOSdata) OemToAnsi(szIO,szIO);
     asDat = asDat+AnsiString(szIO);
     }

//����� ���.������ ����� ������ ��� ����������
//(� ���������� ������� ������)
   (*istrd)++;
//� ��� ���������
   if (cProc!=NULL)
     {
     kz = cProc(*istrd);      //���� ��������� ������� ��������
     if (kz)                  //���������, ���������
       rtf_CloseFiles("�������� ������ ��������");
     }

//������ ������ � ����������� �� ������������
   if ((asDat.Length()<1)||(asDat[1]=='#'))
     continue;

//����� �������� ���������
   if (asDat.AnsiPos("|DOS:")==1)
     {
     boolDOSdata = true;
     continue;
     }
   if (asDat.AnsiPos("|WIN:")==1)
     {
     boolDOSdata = false;
     continue;
     }

//��������� ��������� �� ������ ������
   rtf_GetParameter(&asDat,&asParname,&asParvalue,&curlev);
   if ((curlev<0)||(curlev>rtf_maxLev))
     rtf_CloseFiles(rtf_ErrStruc);

   ilen = rtf_StrInp.Length();
//���� ������� ������ ����� ������ ����� 0, ��
   if (curlev==0)
     {
     //���� �������� � �������� ������
     ib1 = rtf_StrOut.AnsiPos(asParname);
     if (!ib1)
       {
       //���� �� �������, ���� � ��������
       ib1 = rtf_StrInp.AnsiPos(asParname);
       if (ib1)
         {
         //���� �������, ���� ������� ������� � ����
         //���������� ������� �� �������� ������ � ��������
         ib2 = ilen+1;
         for (ib=ib1; ib<ib2; ib++)
           {
           ch = rtf_StrInp[ib];
           if (ch==rtf_LevMarker)
             {         //���� ��������� ������ �� ��������,
             ib2=ib;   //����� �� ���� ����������� -
             break;    //���������� �� 0-� ����������
             }         //1-�� ������
           }
         rtf_StrOut = rtf_StrOut+rtf_StrInp.SubString(1,ib2-1);
         rtf_StrInp = rtf_StrInp.SubString(ib2,ilen);
         }
       }
     //��� ����� ��������� � ���.������ �������� �� �� ��������
     rtf_ParIntoReport(&rtf_StrOut,asParname,asParvalue);
     }

// ���� ������� ������ ����� ������ > 0, ��
   else //if (curlev>0)
     {
     //���� ������� � �������
     asTmp = rtf_Marker(curlev);
     ib1 = rtf_StrInp.AnsiPos(asTmp);
     ib2 = rtf_StrInp.SubString(ib1+4,ilen).AnsiPos(asTmp);
//������� 23.04.02 :
//������� �� ������ � ������� - ���������� ������ ������
     if (ib2==0) continue;
//��������� ������� �� �������
     asTmp = rtf_StrInp.SubString(ib1+4,ib2-1);
//������ ������ � �������� ������ (��� �����)
     if (prevlev==0)
//���� ��� ������� �������, ������� � �������� ������
//��, ��� ���� � ��� �� ������� ������: �������, ��� ������
//����� ������ �������
       {
       if (curlev>1)
         ib1 = rtf_StrInp.AnsiPos(rtf_Marker(1));
       if (ib1>1)
         {
         rtf_StrOut = rtf_StrOut+rtf_StrInp.SubString(1,ib1-1);
         rtf_StrInp = rtf_StrInp.SubString(ib1,ilen);
         }
       }
     rtf_StrOut = rtf_StrOut+asTmp;
//������ �������� ���������� �� ������ ������ � ���.������
     while (asParname.Length())
       {
       rtf_ParIntoReport(&rtf_StrOut,asParname,asParvalue);
       rtf_GetParameter(&asDat,&asParname,&asParvalue,&ib1);
       }
//������ ������ ������ � ������� >0 ���������� :
//��������� �������� ������ � ���� ������
     rtf_StrIntoReport(&rtf_StrOut,&kz);
     if (kz)
       rtf_CloseFiles("������ ������ � ���� ������");
     }
//����� �������� ������ ��������, ���� ��������� 0-� �������
//������������
   prevlev = curlev;
   }

//������� �������� ������ ���� �������� � �������� (������� � �.�.)
 rtf_StrOut = rtf_StrOut+rtf_StrInp;

//������ �������� ������ � ���� ������
 rtf_StrIntoReport(&rtf_StrOut,&kz);
 rtf_CloseFiles(kz?"������ ������ � ���� ������":"");
}
