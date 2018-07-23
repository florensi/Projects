//---------------------------------------------------------------------------

#include <vcl.h>
#pragma hdrstop
//---------------------------------------------------------------------------
USEFORM("uMain.cpp", Main);
USEFORM("uDM.cpp", DM); /* TDataModule: File Type */
USEFORM("uData.cpp", Data);
//---------------------------------------------------------------------------
WINAPI WinMain(HINSTANCE, HINSTANCE, LPSTR, int)
{
        try
        {
                 Application->Initialize();

                 Application->Title = "Страхование жизни";
                 Application->CreateForm(__classid(TDM), &DM);
                 Application->CreateForm(__classid(TMain), &Main);
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
