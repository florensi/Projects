//---------------------------------------------------------------------------
/*
  Модуль RepoRTFO.cpp содержит функции для формирования
  файлов данных, обрабатываемых печатником RepoRTF:
    открытие файла данных (rtf_Open)
    запись в файл данных (rtf_Out, rtf_LineFeed)
    закрытие файла данных (rtf_Close)
  Функции rtf_Out - перегруженные, для записи данных разных типов.
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

//выходной файл
FILE *rtf_Outerstream;

//открытие файла данных
bool __fastcall RepoRTF_outerface::rtf_Open(char* fname)
{
  rtf_Outerstream = fopen(fname,"w");
  return (rtf_Outerstream!=NULL);
}

//закрытие файла данных
bool __fastcall RepoRTF_outerface::rtf_Close()
{
  return (fclose(rtf_Outerstream)==0);
}

//переход к след.строке в файле данных
bool __fastcall RepoRTF_outerface::rtf_LineFeed()
{
  return (fputs("\n",rtf_Outerstream)!=EOF);
}

//Запись в файл данных строки
bool __fastcall RepoRTF_outerface::rtf_Out
(AnsiString nam, AnsiString val, int nlev)
{
  int rez = fputs(("|"+AnsiString(nlev)+":"+nam+":"+val).c_str(),
    rtf_Outerstream);
  if (nlev==0) return rtf_LineFeed();
  return (rez!=EOF);
}

//Получение строки типа 1123456789 в виде 1,123,456,789
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

//Запись в файл данных веществ.числа
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

//Запись в файл данных целого числа
bool __fastcall RepoRTF_outerface::rtf_Out
(AnsiString nam, long val, int nlev)
{
  char st[12];
  ltoa(val,st,10);
  return rtf_Out(nam,AnsiString(st).Trim(),nlev);
}

//Запись в файл данных даты (form="dd.mm.yyyy")
//и времени (form="hh:mm")
bool __fastcall RepoRTF_outerface::rtf_Out
(AnsiString nam, AnsiString form, TDateTime val, int nlev)
{
  return rtf_Out(nam,FormatDateTime(form,val),nlev);
}

