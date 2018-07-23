//---------------------------------------------------------------------------
#ifndef FuncCryptH
#define FuncCryptH
//---------------------------------------------------------------------------

namespace FuncCrypt {

  struct StructKeyBlob {
    BLOBHEADER Header;
    DWORD Length;
    BYTE Key[16];
  };

  bool __fastcall DecryptString(AnsiString InStr, AnsiString &OutStr, AnsiString Key = "@3y7kfyvu9h#?xc6");
  bool __fastcall DecryptFromFile(AnsiString FileName, AnsiString &OutStr, AnsiString Key = "@3y7kfyvu9h#?xc6");
}

using namespace FuncCrypt;

#endif
