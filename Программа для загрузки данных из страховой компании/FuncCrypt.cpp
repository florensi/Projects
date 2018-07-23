//---------------------------------------------------------------------------
#include <vcl.h>
#include <wincrypt.h>
#include <fstream.h>
#pragma hdrstop
#include "FuncCrypt.h"
//---------------------------------------------------------------------------
#pragma package(smart_init)

using namespace FuncCrypt;

//---------------------------------------------------------------------------
bool __fastcall FuncCrypt::DecryptString(AnsiString InStr, AnsiString &OutStr, AnsiString Key)
{
  HCRYPTPROV hProv = NULL;
  HCRYPTKEY hKey = NULL;
  StructKeyBlob KeyBlob;
  DWORD DataLen, Mode;

  if (InStr.IsEmpty()) return(false);
  if (Key.IsEmpty()) return(false);
  if (Key.Length()>16) return(false);

  if (!CryptAcquireContext(&hProv, NULL, NULL, PROV_RSA_AES, CRYPT_VERIFYCONTEXT)) return(false);

  try {
    KeyBlob.Header.bType = PLAINTEXTKEYBLOB;
    KeyBlob.Header.bVersion = CUR_BLOB_VERSION;
    KeyBlob.Header.reserved = 0;
    KeyBlob.Header.aiKeyAlg = CALG_AES_128;
    KeyBlob.Length = 16;

    for (int i=0; i<16; i++) KeyBlob.Key[i] = 0;
    CopyMemory(KeyBlob.Key, Key.c_str(), Key.Length());

    if (!CryptImportKey(hProv, (BYTE*)&KeyBlob, sizeof(KeyBlob), 0, 0, &hKey)) return(false);

    try {
      Mode = CRYPT_MODE_CBC;
      if (!CryptSetKeyParam(hKey, KP_MODE, (BYTE*)&Mode, 0)) return(false);

      Mode = PKCS5_PADDING;
      if (!CryptSetKeyParam(hKey, KP_PADDING, (BYTE*)&Mode, 0)) return(false);

      OutStr = InStr;
      DataLen = OutStr.Length();

      if (!CryptDecrypt(hKey, 0, true, 0, (BYTE*)OutStr.c_str(), &DataLen)) return(false);

      OutStr.SetLength(DataLen);
    }
    __finally {
      CryptDestroyKey(hKey);
    }
  }
  __finally {
    CryptReleaseContext(hProv, 0);
  }

  return(true);
}
//---------------------------------------------------------------------------
bool __fastcall FuncCrypt::DecryptFromFile(AnsiString FileName, AnsiString &OutStr, AnsiString Key)
{
  HANDLE FileHandle;
  DWORD Size, R;

  FileHandle = CreateFile(FileName.c_str(), GENERIC_READ, FILE_SHARE_READ, NULL, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, NULL);
  Size = GetFileSize(FileHandle, &R);
  char *Buf = (char *) malloc(Size);
  ReadFile(FileHandle, Buf, Size, &R, NULL);
  CloseHandle(FileHandle);
  bool Result = DecryptString(AnsiString(Buf, R), OutStr, Key);
  free(Buf);
  return(Result);
}
//---------------------------------------------------------------------------

