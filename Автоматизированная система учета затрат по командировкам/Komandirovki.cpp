//---------------------------------------------------------------------------

#include <vcl.h>
#pragma hdrstop
//---------------------------------------------------------------------------
USEFORM("uMain.cpp", Main);
USEFORM("uDM.cpp", DM); /* TDataModule: File Type */
USEFORM("uVvod.cpp", Vvod);
USEFORM("uGostinica.cpp", Gostinica);
USEFORM("uSprav.cpp", Sprav);
//---------------------------------------------------------------------------
WINAPI WinMain(HINSTANCE, HINSTANCE, LPSTR, int)
{
        try
        {
                 Application->Initialize();
                 Application->Title = "Учет затрат по командировкам";
                 Application->CreateForm(__classid(TDM), &DM);
                 Application->CreateForm(__classid(TMain), &Main);
                 Application->CreateForm(__classid(TVvod), &Vvod);
                 Application->CreateForm(__classid(TGostinica), &Gostinica);
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
