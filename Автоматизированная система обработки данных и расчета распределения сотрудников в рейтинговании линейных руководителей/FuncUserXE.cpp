//---------------------------------------------------------------------------
#include <vcl.h>
#include <lm.h>
#pragma hdrstop

#include "FuncUserXE.h"
//---------------------------------------------------------------------------
#pragma package(smart_init)

using namespace FuncUserXE;

//---------------------------------------------------------------------------
bool __fastcall FuncUserXE::GetUserInfo(String &UserName, String &DomainName)
{
  wchar_t cUser[255], cDomain[255];
  DWORD cchUser = 255, cchDomain = 255;

  if (GetCurrentUserAndDomain(cUser, &cchUser, cDomain, &cchDomain)) {
	UserName = String(cUser).LowerCase();
	DomainName = String(cDomain);
	return(true);
  }

  return(false);
}
//---------------------------------------------------------------------------
bool __fastcall FuncUserXE::GetFullUserInfo(String &UserName, String &DomainName, String &UserFullName)
{
  wchar_t cUser[255], cDomain[255], cFullName[255];
  DWORD cchUser = 255, cchDomain = 255;

  if (GetCurrentUserAndDomain(cUser, &cchUser, cDomain, &cchDomain)) {
	UserName = String(cUser);
	DomainName = String(cDomain);

	if (GetFullName(cUser, cDomain, cFullName)) {
	  UserFullName = String(cFullName);
	  UserName = UserName.LowerCase();
	  return(true);
	}
  }

  return(false);
}
//---------------------------------------------------------------------------
bool __fastcall FuncUserXE::GetCurrentUserAndDomain(PTSTR szUser, PDWORD pcchUser, PTSTR szDomain, PDWORD pcchDomain)
{
  bool fSuccess = false;
  HANDLE hToken = NULL;
  PTOKEN_USER ptiUser = NULL;
  DWORD cbti = 0;
  SID_NAME_USE snu;

  __try {
	// Получаем маркёр доступа вызывающего потока
	if (!OpenThreadToken(GetCurrentThread(), TOKEN_QUERY, true, &hToken)) {
	  if (GetLastError() != ERROR_NO_TOKEN) throw Exception("");

	  // Если маркёра потока не существует, то запрашиваем маркёр процесса
	  if (!OpenProcessToken(GetCurrentProcess(), TOKEN_QUERY, &hToken)) throw Exception("");
	}

	// Получаем размер информации о пользователе в маркёре
	if (GetTokenInformation(hToken, TokenUser, NULL, 0, &cbti)) {
	  // Если длина буфера равна нулю, то ошибка
	  throw Exception("");
	}
	else {
	  // Если длина буфера равна нулю, то ошибка
	  if (GetLastError() != ERROR_INSUFFICIENT_BUFFER) throw Exception("");
	}

	// Распределяем буфер для информации о пользователе в маркёре
	ptiUser = (PTOKEN_USER) HeapAlloc(GetProcessHeap(), 0, cbti);
	if (!ptiUser) throw Exception("");

	// Получаем информацию о пользователе из маркёра
	if (!GetTokenInformation(hToken, TokenUser, ptiUser, cbti, &cbti)) throw Exception("");

	// Получаем имя пользователя и имя домена по пользовательскому SID
	if (!LookupAccountSid(NULL, ptiUser->User.Sid, szUser, pcchUser, szDomain, pcchDomain, &snu)) throw Exception("");

	fSuccess = true;
  }
  __finally {
	// Освобождаем ресурсы
	if (hToken) CloseHandle(hToken);
	if (ptiUser) HeapFree(GetProcessHeap(), 0, ptiUser);
  }

  return fSuccess;
}
//---------------------------------------------------------------------------
bool __fastcall FuncUserXE::GetFullName(wchar_t *sUser, wchar_t *sDomain, wchar_t *sFullName)
{
  LPWKSTA_USER_INFO_1 pBuf = NULL;
  struct _USER_INFO_2 *ui; // Структура для пользователя
  NET_API_STATUS st;

  // Получаем имя компьютера, который является контроллером домена для указанного домена
  if (NetWkstaUserGetInfo(NULL, 1, (LPBYTE *)&pBuf) != NERR_Success) {
	if (pBuf) NetApiBufferFree(pBuf);
	return(false);
  }

  // Ищем пользователя в контроллере домена
  if (String(sDomain)==String("METINVEST")) st = NetUserGetInfo((String(pBuf->wkui1_logon_server) + ".metinvest.ua").c_str(), pBuf->wkui1_username, 2, (LPBYTE *) &ui);
  else st = NetUserGetInfo(pBuf->wkui1_logon_server, pBuf->wkui1_username, 2, (LPBYTE *) &ui);
  if (st) return(false);
  if (pBuf) NetApiBufferFree(pBuf);

  wcscpy(sFullName, ui->usri2_full_name);

  return(true);
}
//---------------------------------------------------------------------------
bool __fastcall FuncUserXE::GetUserGroups(String UserName, String DomainName, TStringList *SL, String Filter)
{
  bool fSuccess = false;
  LPGROUP_USERS_INFO_0 pBuf = NULL;
  DWORD dwEntriesRead = 0;
  DWORD dwTotalEntries = 0;
  LPWKSTA_USER_INFO_1 pBuf2 = NULL;
  NET_API_STATUS st;

  // Получаем имя компьютера, который является контроллером домена для указанного домена
  if (NetWkstaUserGetInfo(NULL, 1, (LPBYTE *)&pBuf2) != NERR_Success) {
	if (pBuf2) NetApiBufferFree(pBuf2);
	return(false);
  }

  // Получаем перечень доменных групп, в которые включен пользователь
  if (DomainName==String("METINVEST")) st = NetUserGetGroups((String(pBuf2->wkui1_logon_server) + ".metinvest.ua").c_str(), pBuf2->wkui1_username, 0, (LPBYTE*)&pBuf, MAX_PREFERRED_LENGTH, &dwEntriesRead, &dwTotalEntries);
  else st = NetUserGetGroups(pBuf2->wkui1_logon_server, pBuf2->wkui1_username, 0, (LPBYTE*)&pBuf, MAX_PREFERRED_LENGTH, &dwEntriesRead, &dwTotalEntries);
  if (st == NERR_Success) {
	LPGROUP_USERS_INFO_0 pTmpBuf;

	if ((pTmpBuf = pBuf) != NULL) {
	  for (DWORD i=0; i<dwEntriesRead; i++) {
		if (wcsstr(pTmpBuf->grui0_name, Filter.c_str())) SL->Add(String(pTmpBuf->grui0_name));
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
bool __fastcall FuncUserXE::GetGroupList(String OrganizationUnitsName, String GroupName, TStringList *SL)
{
  bool fSuccess = false;
  Variant objConnection, objQuery;

  __try {
	objConnection = Variant::CreateObject("ADODB.Connection");
	objConnection.OlePropertySet("CommandTimeout", 120);
	objConnection.OlePropertySet("Provider", WideString("ADsDSOObject").c_bstr());
	objConnection.OleFunction("Open");

	objQuery = objConnection.OleFunction("Execute", WideString("select name from 'LDAP://" + OrganizationUnitsName + "' where objectClass='group'" + (GroupName.IsEmpty() ? String("") : String(" and name='"+GroupName+"*'")) + " order by name").c_bstr());
	for (int i=1; i<=objQuery.OlePropertyGet("RecordCount"); i++) {
	  SL->Add(objQuery.OlePropertyGet("Fields", WideString("name").c_bstr()).OlePropertyGet("Value"));
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
bool __fastcall FuncUserXE::GetUserData(TUserInfo &UI, bool Flg)
{
  bool fSuccess = false;
  Variant objConnection, objQuery;
  String LDAP;

  if (Flg) {
	if (UI.Domain == String("METINVEST")) LDAP = "LDAP://OU=MMK,DC=metinvest,DC=ua";
	else LDAP = "LDAP://OU=MMK,DC=MMK,DC=Local";
  }
  else {
    if (UI.Domain == String("METINVEST")) LDAP = "LDAP://DC=metinvest,DC=ua";
    else LDAP = "LDAP://OU=Внешние пользователи,OU=MMK,DC=MMK,DC=Local";
  }

  __try {
	objConnection = Variant::CreateObject("ADODB.Connection");
	objConnection.OlePropertySet("CommandTimeout", 120);
	objConnection.OlePropertySet("Provider", WideString("ADsDSOObject").c_bstr());
	objConnection.OleFunction("Open");

	if (Flg) objQuery = objConnection.OleFunction("Execute", WideString("select extensionattribute6, name, telephonenumber, mail from " + QuotedStr(LDAP) + " where objectClass='user' and samaccountname=" + QuotedStr(UI.User)).c_bstr());
	else objQuery = objConnection.OleFunction("Execute", WideString("select extensionattribute1, name, telephonenumber, mail from " + QuotedStr(LDAP) + " where objectClass='user' and samaccountname=" + QuotedStr(UI.User)).c_bstr());
	if (objQuery.OlePropertyGet("RecordCount")==1) {
	  if (Flg) {
		if (!objQuery.OlePropertyGet("Fields", WideString("extensionattribute6").c_bstr()).OlePropertyGet("Value").IsNull()) UI.ID = objQuery.OlePropertyGet("Fields", WideString("extensionattribute6").c_bstr()).OlePropertyGet("Value");
		else UI.ID = "";
	  }
	  else {
		if (!objQuery.OlePropertyGet("Fields", WideString("extensionattribute1").c_bstr()).OlePropertyGet("Value").IsNull()) UI.ID = objQuery.OlePropertyGet("Fields", WideString("extensionattribute1").c_bstr()).OlePropertyGet("Value");
		else UI.ID = "";
	  }

	  if (!objQuery.OlePropertyGet("Fields", WideString("name").c_bstr()).OlePropertyGet("Value").IsNull()) UI.Name = objQuery.OlePropertyGet("Fields", WideString("name").c_bstr()).OlePropertyGet("Value");
	  else UI.Name = "";

	  if (!objQuery.OlePropertyGet("Fields", WideString("telephonenumber").c_bstr()).OlePropertyGet("Value").IsNull()) UI.Telephone = objQuery.OlePropertyGet("Fields", WideString("telephonenumber").c_bstr()).OlePropertyGet("Value");
	  else UI.Telephone = "";

	  if (!objQuery.OlePropertyGet("Fields", WideString("mail").c_bstr()).OlePropertyGet("Value").IsNull()) UI.Email = objQuery.OlePropertyGet("Fields", WideString("mail").c_bstr()).OlePropertyGet("Value");
	  else UI.Email = "";
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
bool __fastcall FuncUserXE::GetUserDataID(TUserInfo &UI, bool Flg)
{
  bool fSuccess = false;
  Variant objConnection, objQuery;
  String LDAP;

  if (Flg) {
	if (UI.Domain == String("METINVEST")) LDAP = "LDAP://OU=MMK,DC=metinvest,DC=ua";
	else LDAP = "LDAP://OU=MMK,DC=MMK,DC=Local";
  }
  else {
    if (UI.Domain == String("METINVEST")) LDAP = "LDAP://DC=metinvest,DC=ua";
    else LDAP = "LDAP://OU=Внешние пользователи,OU=MMK,DC=MMK,DC=Local";
  }

  __try {
	objConnection = Variant::CreateObject("ADODB.Connection");
	objConnection.OlePropertySet("CommandTimeout", 120);
	objConnection.OlePropertySet("Provider", WideString("ADsDSOObject").c_bstr());
	objConnection.OleFunction("Open");

	if (Flg) objQuery = objConnection.OleFunction("Execute", WideString("select samaccountname, name, telephonenumber, mail from " + QuotedStr(LDAP) + " where objectClass='user' and extensionattribute6=" + UI.ID).c_bstr());
	else objQuery = objConnection.OleFunction("Execute", WideString("select samaccountname, name, telephonenumber, mail from " + QuotedStr(LDAP) + " where objectClass='user' and extensionattribute1=" + UI.ID).c_bstr());
	if (objQuery.OlePropertyGet("RecordCount")==1) {
	  if (!objQuery.OlePropertyGet("Fields", WideString("samaccountname").c_bstr()).OlePropertyGet("Value").IsNull()) UI.User = objQuery.OlePropertyGet("Fields", WideString("samaccountname").c_bstr()).OlePropertyGet("Value");
	  else UI.User = "";

	  if (!objQuery.OlePropertyGet("Fields", WideString("name").c_bstr()).OlePropertyGet("Value").IsNull()) UI.Name = objQuery.OlePropertyGet("Fields", WideString("name").c_bstr()).OlePropertyGet("Value");
	  else UI.Name = "";

	  if (!objQuery.OlePropertyGet("Fields", WideString("telephonenumber").c_bstr()).OlePropertyGet("Value").IsNull()) UI.Telephone = objQuery.OlePropertyGet("Fields", WideString("telephonenumber").c_bstr()).OlePropertyGet("Value");
	  else UI.Telephone = "";

	  if (!objQuery.OlePropertyGet("Fields", WideString("mail").c_bstr()).OlePropertyGet("Value").IsNull()) UI.Email = objQuery.OlePropertyGet("Fields", WideString("mail").c_bstr()).OlePropertyGet("Value");
	  else UI.Email = "";
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

