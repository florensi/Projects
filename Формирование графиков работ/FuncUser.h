//---------------------------------------------------------------------------
#ifndef FuncUserH
#define FuncUserH
//---------------------------------------------------------------------------

namespace FuncUser {
  bool __fastcall GetFullUserInfo(AnsiString &UserName, AnsiString &DomainName, AnsiString &UserFullName);
  bool __fastcall GetUserInfo(AnsiString &UserName, AnsiString &DomainName);
  bool __fastcall GetCurrentUserAndDomain(PTSTR szUser, PDWORD pcchUser, PTSTR szDomain, PDWORD pcchDomain);
  bool __fastcall GetFullName(char *sUser, char *sDomain, char *sFullName);
  bool __fastcall GetUserGroups(AnsiString UserName, AnsiString DomainName, TStringList *SL);
  bool __fastcall GetUserTelephone(AnsiString UserName, AnsiString DomainName, AnsiString &UserTelephone, bool Flg=true);
  bool __fastcall GetGroups(AnsiString OrganizationUnitsName, AnsiString GroupName, TStringList *SL);
  bool __fastcall GetUserEmail(AnsiString UserName, AnsiString DomainName, AnsiString &UserEmail, bool Flg=true);
  bool __fastcall GetUserEmailID(int ID, AnsiString &UserEmail, bool Flg=true);
  bool __fastcall GetUserNameID(int ID, AnsiString &UserName, bool Flg=true);
  bool __fastcall GetUserFullName(AnsiString UserName, AnsiString DomainName, AnsiString &UserFullName, bool Flg=true);
}

using namespace FuncUser;

#endif
