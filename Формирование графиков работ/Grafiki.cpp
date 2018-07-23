//---------------------------------------------------------------------------

#include <vcl.h>
#pragma hdrstop
//---------------------------------------------------------------------------
USEFORM("uMain.cpp", Main);
USEFORM("uDM.cpp", DM); /* TDataModule: File Type */
USEFORM("uVvod.cpp", Vvod);
USEFORM("uSprav.cpp", Sprav);
USEFORM("uData.cpp", Data);
//---------------------------------------------------------------------------
WINAPI WinMain(HINSTANCE, HINSTANCE, LPSTR, int)
{
        try
        {
                 Application->Initialize();
                 Application->Title = "Графики работы";
                 Application->CreateForm(__classid(TDM), &DM);
                 Application->CreateForm(__classid(TMain), &Main);
                 Application->CreateForm(__classid(TVvod), &Vvod);
                 Application->CreateForm(__classid(TSprav), &Sprav);
                 Application->CreateForm(__classid(TData), &Data);
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
