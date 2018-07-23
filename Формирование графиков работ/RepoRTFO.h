//---------------------------------------------------------------------------
#ifndef RepoRTFOH
#define RepoRTFOH
//---------------------------------------------------------------------------

namespace RepoRTF_outerface {

//открытие файла данных
bool __fastcall rtf_Open(char* fname);

//закрытие файла данных
bool __fastcall rtf_Close();

//переход к след.строке в файле данных (конец строки уровня)
bool __fastcall rtf_LineFeed();

//Запись в файл данных строки
bool __fastcall rtf_Out(AnsiString nam, AnsiString val, int nlev);

//Запись в файл данных веществ.числа
bool __fastcall rtf_Out(AnsiString nam, double val,
 int zn1, int zn2, int nlev);

//Запись в файл данных целого числа
bool __fastcall rtf_Out(AnsiString nam, long val, int nlev);

//Запись в файл данных даты (form="dd.mm.yyyy")
//и времени (form="hh:mm")
bool __fastcall RepoRTF_outerface::rtf_Out
(AnsiString nam, AnsiString form, TDateTime val, int nlev);
}

using namespace RepoRTF_outerface;
#endif
