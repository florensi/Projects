//---------------------------------------------------------------------------
#include <vcl.h>
#include <lm.h>
#include <ComCtrls.hpp>
#pragma hdrstop

#pragma comment(lib,"netapi32.lib")

#include "FuncUser.h"
//---------------------------------------------------------------------------
#pragma package(smart_init)

using namespace FuncUser;

//---------------------------------------------------------------------------
bool __fastcall FuncUser::GetCurrentUserAndDomain(PTSTR szUser, PDWORD pcchUser, PTSTR szDomain, PDWORD pcchDomain)
{
  bool fSuccess = false;
  HANDLE hToken = NULL;
  PTOKEN_USER ptiUser = NULL;
  DWORD cbti = 0;
  SID_NAME_USE snu;

  __try {
    // �������� ����� ������� ����������� ������
    if (!OpenThreadToken(GetCurrentThread(), TOKEN_QUERY, true, &hToken)) {
      if (GetLastError() != ERROR_NO_TOKEN) throw Exception("");

      // ���� ������ ������ �� ����������, �� ����������� ����� ��������
      if (!OpenProcessToken(GetCurrentProcess(), TOKEN_QUERY, &hToken)) throw Exception("");
    }

    // �������� ������ ���������� � ������������ � ������
    if (GetTokenInformation(hToken, TokenUser, NULL, 0, &cbti)) {
      // ���� ����� ������ ����� ����, �� ������
      throw Exception("");
    }
    else {
      // ���� ����� ������ ����� ����, �� ������
      if (GetLastError() != ERROR_INSUFFICIENT_BUFFER) throw Exception("");
    }

    // ������������ ����� ��� ���������� � ������������ � ������
    ptiUser = (PTOKEN_USER) HeapAlloc(GetProcessHeap(), 0, cbti);
    if (!ptiUser) throw Exception("");

    // �������� ���������� � ������������ �� ������
    if (!GetTokenInformation(hToken, TokenUser, ptiUser, cbti, &cbti)) throw Exception("");

    // �������� ��� ������������ � ��� ������ �� ����������������� SID
    if (!LookupAccountSid(NULL, ptiUser->User.Sid, szUser, pcchUser, szDomain, pcchDomain, &snu)) throw Exception("");

    fSuccess = true;
  }
  __finally {
    // ����������� �������
    if (hToken) CloseHandle(hToken);
    if (ptiUser) HeapFree(GetProcessHeap(), 0, ptiUser);
  }

  return fSuccess;
}
//---------------------------------------------------------------------------
/*bool __fastcall FuncUser::GetFullName(char *sUser, char *sDomain, char *sFullName)
{
  WCHAR wszUserName[255]; // ��� ������������ � Unicode
  WCHAR wszDomain[255];
  LPBYTE DCName;

  struct _SERVER_INFO_100 *si100; // ��������� ��� �������
  struct _USER_INFO_2 *ui; // ��������� ��� ������������

  // ������������ ��� ������������ � ������ �� ASCII � Unicode
  MultiByteToWideChar(CP_ACP, 0, sUser, strlen(sUser)+1, wszUserName, sizeof(wszUserName));
  MultiByteToWideChar(CP_ACP, 0, sDomain, strlen(sDomain)+1, wszDomain, sizeof(wszDomain));

  // �������� ��� ����������, ������� �������� ������������ ������ (DC) ��� ���������� ������
LPWKSTA_USER_INFO_1 pBuf = NULL;
if (NetWkstaUserGetInfo(NULL, 1, (LPBYTE *)&pBuf) != NERR_Success) {
  if (pBuf) NetApiBufferFree(pBuf);
  return(false);
}

  // ���� ������������ � ����������� ������
if (NetUserGetInfo(pBuf->wkui1_logon_server, pBuf->wkui1_username, 2, (LPBYTE *) &ui)) return(false);
if (pBuf) NetApiBufferFree(pBuf);

  // ����������� ������ ��� �� Unicode � ASCII
  WideCharToMultiByte(CP_ACP, 0, ui->usri2_full_name, -1, sFullName, 255, NULL, NULL);

  return(true);
}
//---------------------------------------------------------------------------*/
bool __fastcall FuncUser::GetFullName(char *sUser, char *sDomain, char *sFullName)
{
  LPWKSTA_USER_INFO_1 pBuf = NULL;
  struct _USER_INFO_2 *ui; // ��������� ��� ������������
  NET_API_STATUS st;

  // �������� ��� ����������, ������� �������� ������������ ������ ��� ���������� ������
  if (NetWkstaUserGetInfo(NULL, 1, (LPBYTE *)&pBuf) != NERR_Success) {
    if (pBuf) NetApiBufferFree(pBuf);
    return(false);
  }

  // ���� ������������ � ����������� ������
  if (AnsiString(sDomain)=="METINVEST") st = NetUserGetInfo((WideString(pBuf->wkui1_logon_server) + ".metinvest.ua").c_bstr(), pBuf->wkui1_username, 2, (LPBYTE *) &ui);
  else st = NetUserGetInfo(pBuf->wkui1_logon_server, pBuf->wkui1_username, 2, (LPBYTE *) &ui);
  if (st) return(false);
  if (pBuf) NetApiBufferFree(pBuf);

  // ����������� ������ ��� �� Unicode � ASCII
  WideCharToMultiByte(CP_ACP, 0, ui->usri2_full_name, -1, sFullName, 255, NULL, NULL);

  return(true);
}
//---------------------------------------------------------------------------

bool __fastcall FuncUser::GetFullUserInfo(AnsiString &UserName, AnsiString &DomainName, AnsiString &UserFullName)
{
  char cUser[255], cDomain[255], cFullName[255];
  DWORD cchUser = 255, cchDomain = 255;

  if (GetCurrentUserAndDomain(cUser, &cchUser, cDomain, &cchDomain)) {
    UserName = AnsiString(cUser);
    DomainName = AnsiString(cDomain);

    if (GetFullName(cUser, cDomain, cFullName)) {
      UserFullName = AnsiString(cFullName);
      UserName = UserName.LowerCase();
      return(true);
    }
  }

  return(false);
}
//---------------------------------------------------------------------------
bool __fastcall FuncUser::GetUserInfo(AnsiString &UserName, AnsiString &DomainName)
{
  char cUser[255], cDomain[255];
  DWORD cchUser = 255, cchDomain = 255;

  if (GetCurrentUserAndDomain(cUser, &cchUser, cDomain, &cchDomain)) {
    UserName = AnsiString(cUser).LowerCase();
    DomainName = AnsiString(cDomain);
    return(true);
  }

  return(false);
}
//---------------------------------------------------------------------------
/*bool __fastcall FuncUser::GetUserGroups(AnsiString UserName, AnsiString DomainName, TStringList *SL)
{
  bool fSuccess = false;
  WCHAR wszUserName[255]; // ��� ������������ � Unicode
  WCHAR wszDomain[255];
  LPBYTE DCName;

  LPGROUP_USERS_INFO_0 pBuf = NULL;
  DWORD dwEntriesRead = 0;
  DWORD dwTotalEntries = 0;

  // ������������ ��� ������������ � ������ �� ASCII � Unicode
  MultiByteToWideChar(CP_ACP, 0, UserName.c_str(), UserName.Length()+1, wszUserName, sizeof(wszUserName));
  MultiByteToWideChar(CP_ACP, 0, DomainName.c_str(), DomainName.Length()+1, wszDomain, sizeof(wszDomain));

  // �������� ��� ����������, ������� �������� ������������ ������ (DC) ��� ���������� ������
LPWKSTA_USER_INFO_1 pBuf2 = NULL;
if (NetWkstaUserGetInfo(NULL, 1, (LPBYTE *)&pBuf2) != NERR_Success) {
  if (pBuf2) NetApiBufferFree(pBuf2);
  return(false);
}

  // �������� �������� �������� �����, � ������� ������� ������������
if (NetUserGetGroups(pBuf2->wkui1_logon_server, pBuf2->wkui1_username, 0, (LPBYTE*)&pBuf, MAX_PREFERRED_LENGTH, &dwEntriesRead, &dwTotalEntries) == NERR_Success) {
    LPGROUP_USERS_INFO_0 pTmpBuf;

    if ((pTmpBuf = pBuf) != NULL) {
      for (DWORD i=0; i<dwEntriesRead; i++) {
        SL->Add(AnsiString(pTmpBuf->grui0_name));
        pTmpBuf++;
      }

      fSuccess = true;
    }
  }

  if (pBuf) NetApiBufferFree(pBuf);
if (pBuf2) NetApiBufferFree(pBuf2);
  return fSuccess;
}
//---------------------------------------------------------------------------   */
bool __fastcall FuncUser::GetUserGroups(AnsiString UserName, AnsiString DomainName, TStringList *SL)
{
  bool fSuccess = false;
  LPGROUP_USERS_INFO_0 pBuf = NULL;
  DWORD dwEntriesRead = 0;
  DWORD dwTotalEntries = 0;
  LPWKSTA_USER_INFO_1 pBuf2 = NULL;
  NET_API_STATUS st;

  // �������� ��� ����������, ������� �������� ������������ ������ ��� ���������� ������
  if (NetWkstaUserGetInfo(NULL, 1, (LPBYTE *)&pBuf2) != NERR_Success) {
    if (pBuf2) NetApiBufferFree(pBuf2);
    return(false);
  }

  // �������� �������� �������� �����, � ������� ������� ������������
  if (DomainName=="METINVEST") st = NetUserGetGroups((WideString(pBuf2->wkui1_logon_server) + ".metinvest.ua").c_bstr(), pBuf2->wkui1_username, 0, (LPBYTE*)&pBuf, MAX_PREFERRED_LENGTH, &dwEntriesRead, &dwTotalEntries);
  else st = NetUserGetGroups(pBuf2->wkui1_logon_server, pBuf2->wkui1_username, 0, (LPBYTE*)&pBuf, MAX_PREFERRED_LENGTH, &dwEntriesRead, &dwTotalEntries);
  if (st == NERR_Success) {
    LPGROUP_USERS_INFO_0 pTmpBuf;

    if ((pTmpBuf = pBuf) != NULL) {
      for (DWORD i=0; i<dwEntriesRead; i++) {
        SL->Add(AnsiString(pTmpBuf->grui0_name));
        pTmpBuf++;
      }

      fSuccess = true;
    }
  }

  if (pBuf) NetApiBufferFree(pBuf);
  if (pBuf2) NetApiBufferFree(pBuf2);
  
  return fSuccess;
}
//---------------------------------------------------------------------------

bool __fastcall FuncUser::GetUserTelephone(AnsiString UserName, AnsiString DomainName, AnsiString &UserTelephone, bool Flg)
{
  bool fSuccess = false;
  Variant objConnection, objQuery;

  __try {
    objConnection = Variant::CreateObject("ADODB.Connection");
    objConnection.OlePropertySet("CommandTimeout", 120);
    objConnection.OlePropertySet("Provider", "ADsDSOObject");
    objConnection.OleFunction("Open");

    if (Flg) objQuery = objConnection.OleFunction("Execute", ("select telephoneNumber from 'LDAP://OU=Active,OU=MMK,DC=MMK,DC=Local' where objectClass='user' and userPrincipalName=" + QuotedStr(UserName + "@" + DomainName + ".Local")).c_str());
    else objQuery = objConnection.OleFunction("Execute", ("select telephoneNumber from 'LDAP://OU=������� ������������,OU=MMK,DC=MMK,DC=Local' where objectClass='user' and userPrincipalName=" + QuotedStr(UserName + "@" + DomainName + ".Local")).c_str());
    if (objQuery.OlePropertyGet("RecordCount")==1 && !objQuery.OlePropertyGet("Fields", "telephoneNumber").OlePropertyGet("Value").IsNull()) UserTelephone = objQuery.OlePropertyGet("Fields", "telephoneNumber").OlePropertyGet("Value");
    else UserTelephone = "";

    objConnection.OleFunction("Close");
    objConnection = Unassigned;
    fSuccess = true;
  }
  catch(...) {
    objConnection = Unassigned;
  }

  return fSuccess;
}
//---------------------------------------------------------------------------
bool __fastcall FuncUser::GetGroups(AnsiString OrganizationUnitsName, AnsiString GroupName, TStringList *SL)
{
  bool fSuccess = false;
  Variant objConnection, objQuery;

  __try {
    objConnection = Variant::CreateObject("ADODB.Connection");
    objConnection.OlePropertySet("CommandTimeout", 120);
    objConnection.OlePropertySet("Provider", "ADsDSOObject");
    objConnection.OleFunction("Open");

    objQuery = objConnection.OleFunction("Execute", ("select name from 'LDAP://" + (OrganizationUnitsName.IsEmpty() ? AnsiString("") : OrganizationUnitsName+",") + "DC=MMK,DC=Local' where objectClass='group'" + (GroupName.IsEmpty() ? AnsiString("") : " and Name='"+GroupName+"*'") + " order by name").c_str());
    for (int i=1; i<=objQuery.OlePropertyGet("RecordCount"); i++) {
      SL->Add(AnsiString(objQuery.OlePropertyGet("Fields", "name").OlePropertyGet("Value")));
      objQuery.OleFunction("MoveNext");
    }

    objConnection.OleFunction("Close");
    objConnection = Unassigned;
    fSuccess = true;
  }
  catch(...) {
    objConnection = Unassigned;
  }

  return fSuccess;
}
//---------------------------------------------------------------------------
bool __fastcall FuncUser::GetUserEmail(AnsiString UserName, AnsiString DomainName, AnsiString &UserEmail, bool Flg)
{
  bool fSuccess = false;
  Variant objConnection, objQuery;

  __try {
    objConnection = Variant::CreateObject("ADODB.Connection");
    objConnection.OlePropertySet("CommandTimeout", 120);
    objConnection.OlePropertySet("Provider", "ADsDSOObject");
    objConnection.OleFunction("Open");

    if (Flg) objQuery = objConnection.OleFunction("Execute", ("select mail from 'LDAP://OU=Active,OU=MMK,DC=MMK,DC=Local' where objectClass='user' and userPrincipalName=" + QuotedStr(UserName + "@" + DomainName + ".Local")).c_str());
    else objQuery = objConnection.OleFunction("Execute", ("select mail from 'LDAP://OU=������� ������������,OU=MMK,DC=MMK,DC=Local' where objectClass='user' and userPrincipalName=" + QuotedStr(UserName + "@" + DomainName + ".Local")).c_str());
    if (objQuery.OlePropertyGet("RecordCount")==1 && !objQuery.OlePropertyGet("Fields", "mail").OlePropertyGet("Value").IsNull()) UserEmail = objQuery.OlePropertyGet("Fields", "mail").OlePropertyGet("Value");
    else UserEmail = "";

    objConnection.OleFunction("Close");
    objConnection = Unassigned;
    fSuccess = true;
  }
  catch(...) {
    objConnection = Unassigned;
  }

  return fSuccess;
}
//---------------------------------------------------------------------------
bool __fastcall FuncUser::GetUserEmailID(int ID, AnsiString &UserEmail, bool Flg)
{
  bool fSuccess = false;
  Variant objConnection, objQuery;

  __try {
    objConnection = Variant::CreateObject("ADODB.Connection");
    objConnection.OlePropertySet("CommandTimeout", 120);
    objConnection.OlePropertySet("Provider", "ADsDSOObject");
    objConnection.OleFunction("Open");

    if (Flg) objQuery = objConnection.OleFunction("Execute", ("select mail from 'LDAP://OU=Active,OU=MMK,DC=MMK,DC=Local' where objectClass='user' and extensionAttribute1=" + IntToStr(ID)).c_str());
    else objQuery = objConnection.OleFunction("Execute", ("select mail from 'LDAP://OU=������� ������������,OU=MMK,DC=MMK,DC=Local' where objectClass='user' and extensionAttribute1=" + IntToStr(ID)).c_str());
    if (objQuery.OlePropertyGet("RecordCount")==1 && !objQuery.OlePropertyGet("Fields", "mail").OlePropertyGet("Value").IsNull()) UserEmail = objQuery.OlePropertyGet("Fields", "mail").OlePropertyGet("Value");
    else UserEmail = "";

    objConnection.OleFunction("Close");
    objConnection = Unassigned;
    fSuccess = true;
  }
  catch(...) {
    objConnection = Unassigned;
  }

  return fSuccess;
}
//---------------------------------------------------------------------------
bool __fastcall FuncUser::GetUserNameID(int ID, AnsiString &UserName, bool Flg)
{
  bool fSuccess = false;
  Variant objConnection, objQuery;

  __try {
    objConnection = Variant::CreateObject("ADODB.Connection");
    objConnection.OlePropertySet("CommandTimeout", 120);
    objConnection.OlePropertySet("Provider", "ADsDSOObject");
    objConnection.OleFunction("Open");

    if (Flg) objQuery = objConnection.OleFunction("Execute", ("select samaccountname from 'LDAP://OU=Active,OU=MMK,DC=MMK,DC=Local' where objectClass='user' and extensionAttribute1=" + IntToStr(ID)).c_str());
    else objQuery = objConnection.OleFunction("Execute", ("select samaccountname from 'LDAP://OU=������� ������������,OU=MMK,DC=MMK,DC=Local' where objectClass='user' and extensionAttribute1=" + IntToStr(ID)).c_str());
    if (objQuery.OlePropertyGet("RecordCount")==1 && !objQuery.OlePropertyGet("Fields", "samaccountname").OlePropertyGet("Value").IsNull()) UserName = objQuery.OlePropertyGet("Fields", "samaccountname").OlePropertyGet("Value");
    else UserName = "";

    objConnection.OleFunction("Close");
    objConnection = Unassigned;
    fSuccess = true;
  }
  catch(...) {
    objConnection = Unassigned;
  }

  return fSuccess;
}
//---------------------------------------------------------------------------
bool __fastcall FuncUser::GetUserFullName(AnsiString UserName, AnsiString DomainName, AnsiString &UserFullName, bool Flg)
{
  bool fSuccess = false;
  Variant objConnection, objQuery;

  __try {
    objConnection = Variant::CreateObject("ADODB.Connection");
    objConnection.OlePropertySet("CommandTimeout", 120);
    objConnection.OlePropertySet("Provider", "ADsDSOObject");
    objConnection.OleFunction("Open");

    if (Flg) objQuery = objConnection.OleFunction("Execute", ("select name from 'LDAP://OU=Active,OU=MMK,DC=MMK,DC=Local' where objectClass='user' and userPrincipalName=" + QuotedStr(UserName + "@" + DomainName + ".Local")).c_str());
    else objQuery = objConnection.OleFunction("Execute", ("select name from 'LDAP://OU=������� ������������,OU=MMK,DC=MMK,DC=Local' where objectClass='user' and userPrincipalName=" + QuotedStr(UserName + "@" + DomainName + ".Local")).c_str());
    if (objQuery.OlePropertyGet("RecordCount")==1 && !objQuery.OlePropertyGet("Fields", "name").OlePropertyGet("Value").IsNull()) UserFullName = objQuery.OlePropertyGet("Fields", "name").OlePropertyGet("Value");
    else UserFullName = "";

    objConnection.OleFunction("Close");
    objConnection = Unassigned;
    fSuccess = true;
  }
  catch(...) {
    objConnection = Unassigned;
  }

  return fSuccess;
}
//---------------------------------------------------------------------------

