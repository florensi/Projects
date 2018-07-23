//---------------------------------------------------------------------------
#ifndef FuncUserXEH
#define FuncUserXEH
//---------------------------------------------------------------------------

namespace FuncUserXE {
  struct TUserInfo {
	String Domain;
	String User; // samaccountname
	String ID; // extensionattribute6 (табельный SAP ММКИ) или extensionattribute1 (уникальный номер МРМЗ)
	String Name; // name
	String Telephone; // telephonenumber
    String Email; // mail
  };

  bool __fastcall GetUserInfo(String &UserName, String &DomainName);
  bool __fastcall GetFullUserInfo(String &UserName, String &DomainName, String &UserFullName);
  bool __fastcall GetCurrentUserAndDomain(PTSTR szUser, PDWORD pcchUser, PTSTR szDomain, PDWORD pcchDomain);
  bool __fastcall GetFullName(wchar_t *sUser, wchar_t *sDomain, wchar_t *sFullName);
  bool __fastcall GetUserGroups(String UserName, String DomainName, TStringList *SL, String Filter="");
  bool __fastcall GetGroupList(String OrganizationUnitsName, String GroupName, TStringList *SL);
  bool __fastcall GetUserData(TUserInfo &UI, bool Flg=true);
  bool __fastcall GetUserDataID(TUserInfo &UI, bool Flg=true);
}

using namespace FuncUserXE;

#endif
