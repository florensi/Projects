//---------------------------------------------------------------------------
#ifndef RepoRTFOH
#define RepoRTFOH
//---------------------------------------------------------------------------

namespace RepoRTF_outerface {

//�������� ����� ������
bool __fastcall rtf_Open(char* fname);

//�������� ����� ������
bool __fastcall rtf_Close();

//������� � ����.������ � ����� ������ (����� ������ ������)
bool __fastcall rtf_LineFeed();

//������ � ���� ������ ������
bool __fastcall rtf_Out(AnsiString nam, AnsiString val, int nlev);

//������ � ���� ������ �������.�����
bool __fastcall rtf_Out(AnsiString nam, double val,
 int zn1, int zn2, int nlev);

//������ � ���� ������ ������ �����
bool __fastcall rtf_Out(AnsiString nam, long val, int nlev);

//������ � ���� ������ ���� (form="dd.mm.yyyy")
//� ������� (form="hh:mm")
bool __fastcall RepoRTF_outerface::rtf_Out
(AnsiString nam, AnsiString form, TDateTime val, int nlev);
}

using namespace RepoRTF_outerface;
#endif
