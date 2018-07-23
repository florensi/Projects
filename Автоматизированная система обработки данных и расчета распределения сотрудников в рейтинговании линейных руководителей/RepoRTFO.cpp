//---------------------------------------------------------------------------
/*
  ������ RepoRTFO.cpp �������� ������� ��� ������������
  ������ ������, �������������� ���������� RepoRTF:
    �������� ����� ������ (rtf_Open)
    ������ � ���� ������ (rtf_Out, rtf_LineFeed)
    �������� ����� ������ (rtf_Close)
  ������� rtf_Out - �������������, ��� ������ ������ ������ �����.
*/
//---------------------------------------------------------------------------
#include <vcl.h>
#include <stdlib.h>
#include <stdio.h>
#include <math.h>
#include <algorithm>
using std::min;
using std::max;
#pragma hdrstop

#include "RepoRTFO.h"

//---------------------------------------------------------------------------
#pragma package(smart_init)

using namespace RepoRTF_outerface;

//�������� ����
FILE *rtf_Outerstream;

//�������� ����� ������
bool __fastcall RepoRTF_outerface::rtf_Open(char* fname)
{
  rtf_Outerstream = fopen(fname,"w");
  return (rtf_Outerstream!=NULL);
}

//�������� ����� ������
bool __fastcall RepoRTF_outerface::rtf_Close()
{
  return (fclose(rtf_Outerstream)==0);
}

//������� � ����.������ � ����� ������
bool __fastcall RepoRTF_outerface::rtf_LineFeed()
{
  return (fputs("\n",rtf_Outerstream)!=EOF);
}

//������ � ���� ������ ������
bool __fastcall RepoRTF_outerface::rtf_Out
(AnsiString nam, AnsiString val, int nlev)
{
  int rez = fputs(("|"+AnsiString(nlev)+":"+nam+":"+val).c_str(),
    rtf_Outerstream);
  if (nlev==0) return rtf_LineFeed();
  return (rez!=EOF);
}

//��������� ������ ���� 1123456789 � ���� 1,123,456,789
AnsiString __fastcall GetDivNumber(AnsiString inpst)
{
AnsiString stret="", st = inpst.Trim();
int i=st.Length();
  while (i>3)
    {
    i -=3;
    stret = ","+st.SubString(i+1,3)+stret;
    st = st.SubString(1,i);
    }
  return st+stret;
}

//������ � ���� ������ �������.�����
bool __fastcall RepoRTF_outerface::rtf_Out
(AnsiString nam, double val, int zn1, int zn2, int nlev)
{
char st[30];
  sprintf(st,("%"+AnsiString(zn1)+"."+
    AnsiString(max(0,zn2))+"lf").c_str(),
      val+(val<0?-1:1)*0.000000000001);
  if (zn2<0)
    return rtf_Out(nam,GetDivNumber(st),nlev);
  else
    return rtf_Out(nam,AnsiString(st).Trim(),nlev);
}

//������ � ���� ������ ������ �����
bool __fastcall RepoRTF_outerface::rtf_Out
(AnsiString nam, long val, int nlev)
{
  char st[12];
  ltoa(val,st,10);
  return rtf_Out(nam,AnsiString(st).Trim(),nlev);
}

//������ � ���� ������ ���� (form="dd.mm.yyyy")
//� ������� (form="hh:mm")
bool __fastcall RepoRTF_outerface::rtf_Out
(AnsiString nam, AnsiString form, TDateTime val, int nlev)
{
  return rtf_Out(nam,FormatDateTime(form,val),nlev);
}

