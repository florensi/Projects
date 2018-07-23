//---------------------------------------------------------------------------
#ifndef FuncCryptXEH
#define FuncCryptXEH
//---------------------------------------------------------------------------

namespace FuncCryptXE {

  struct StructKeyBlob {
	BLOBHEADER Header;
	DWORD Length;
	BYTE Key[16];
  };

  bool __fastcall DecryptString(AnsiString InStr, AnsiString &OutStr, AnsiString Key = "@3y7kfyvu9h#?xc6");
  bool __fastcall DecryptFromFile(String FileName, AnsiString &OutStr, AnsiString Key = "@3y7kfyvu9h#?xc6");
}

using namespace FuncCryptXE;

#endif
