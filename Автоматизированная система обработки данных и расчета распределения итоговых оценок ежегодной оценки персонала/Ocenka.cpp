//---------------------------------------------------------------------------

#include <vcl.h>
#pragma hdrstop
//---------------------------------------------------------------------------
USEFORM("uMain.cpp", Main);
USEFORM("uDM.cpp", DM); /* TDataModule: File Type */
USEFORM("uReiting.cpp", Reiting);
USEFORM("uZagruzka.cpp", Zagruzka);
USEFORM("uVvod.cpp", Vvod);
USEFORM("uSprav.cpp", Sprav);
USEFORM("uLogs.cpp", Logs);
USEFORM("uZameshenie.cpp", Zameshenie);
USEFORM("uData.cpp", Data);
//---------------------------------------------------------------------------
WINAPI WinMain(HINSTANCE, HINSTANCE, LPSTR, int)
{
        try
        {
                 Application->Initialize();
                 Application->Title = "Обработка данных  по ежегодной оценке персонала";
                 Application->CreateForm(__classid(TDM), &DM);
                 Application->CreateForm(__classid(TMain), &Main);
                 Application->CreateForm(__classid(TReiting), &Reiting);
                 Application->CreateForm(__classid(TZagruzka), &Zagruzka);
                 Application->CreateForm(__classid(TVvod), &Vvod);
                 Application->CreateForm(__classid(TSprav), &Sprav);
                 Application->CreateForm(__classid(TZameshenie), &Zameshenie);
                 Application->CreateForm(__classid(TLogs), &Logs);
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
