//---------------------------------------------------------------------------
#ifndef RepoRTFMH
#define RepoRTFMH
//---------------------------------------------------------------------------

namespace RepoRTF_interface {

//Внешний интерфейс RepoRTF_interface

typedef int __fastcall (*callingProc)(int);

void __fastcall rtf_CreateReport(const AnsiString asDatFileName,
const AnsiString asPicFileName, const AnsiString asRepFileName,
callingProc cProc, int* istrd);

struct RepoRTF_Error
  {
  const char* Err;
  RepoRTF_Error(const char* szErr)
    {Err=szErr;}
  };

}

using namespace RepoRTF_interface;

#endif
