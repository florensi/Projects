//---------------------------------------------------------------------------

#include <vcl.h>
#pragma hdrstop
#include <tchar.h>
//---------------------------------------------------------------------------
#include <Vcl.Styles.hpp>
#include <Vcl.Themes.hpp>
USEFORM("uDM.cpp", DM); /* TDataModule: File Type */
USEFORM("uMain.cpp", Main);
USEFORM("uVvod.cpp", Vvod);
USEFORM("uSprav.cpp", Sprav);
//---------------------------------------------------------------------------
int WINAPI _tWinMain(HINSTANCE, HINSTANCE, LPTSTR, int)
{
	try
	{
		Application->Initialize();
		Application->MainFormOnTaskBar = true;
		TStyleManager::TrySetStyle("Silver");
		Application->Title = "Рейтингование линейных руководителей";
		Application->CreateForm(__classid(TDM), &DM);
		Application->CreateForm(__classid(TMain), &Main);
		Application->CreateForm(__classid(TVvod), &Vvod);
		Application->CreateForm(__classid(TSprav), &Sprav);
		Application->Run();
	}
	catch (Exception &exception)
	{
		Application->ShowException(&exception);
	}
	catch (...)
	{
		try
		{
			throw Exception("");
		}
		catch (Exception &exception)
		{
			Application->ShowException(&exception);
		}
	}
	return 0;
}
//---------------------------------------------------------------------------
