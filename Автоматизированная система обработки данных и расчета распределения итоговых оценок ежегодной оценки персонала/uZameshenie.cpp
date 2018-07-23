//---------------------------------------------------------------------------

#include <vcl.h>
#pragma hdrstop

#include "uZameshenie.h"
#include "uDM.h"
#include "uMain.h"
//---------------------------------------------------------------------------
#pragma package(smart_init)
#pragma link "DBGridEh"
#pragma resource "*.dfm"
TZameshenie *Zameshenie;
//---------------------------------------------------------------------------
__fastcall TZameshenie::TZameshenie(TComponent* Owner)
        : TForm(Owner)
{
}
//---------------------------------------------------------------------------
void __fastcall TZameshenie::BitBtn2Click(TObject *Sender)
{
  Zameshenie->Close();
}
//---------------------------------------------------------------------------

void __fastcall TZameshenie::FormShow(TObject *Sender)
{
  //Заполнение Edit-ов

  //Тн, цех, ФИО
  EditTN->Text="490";
  EditZEX->Text="";
  EditZEX->Text=zzex=DM->qOcenka->FieldByName("zex")->AsString;
  EditTN->Text=ztn=DM->qOcenka->FieldByName("tn")->AsString;
  LabelFIO->Caption=zfio=DM->qOcenka->FieldByName("fio")->AsString;


  //Пенс. возраст
  EditVZ_PENS->Text=zvz_pens=DM->qOcenka->FieldByName("vz_pens")->AsString;

   //КПЭ
  EditKPE1->Text=zkpe1=DM->qOcenka->FieldByName("kpe1")->AsString;
  EditKPE2->Text=zkpe2=DM->qOcenka->FieldByName("kpe2")->AsString;
  EditKPE3->Text=zkpe3=DM->qOcenka->FieldByName("kpe3")->AsString;
  EditKPE4->Text=zkpe4=DM->qOcenka->FieldByName("kpe4")->AsString;

  if (DM->qOcenka->FieldByName("kpe1")->AsFloat+DM->qOcenka->FieldByName("kpe2")->AsFloat+
      DM->qOcenka->FieldByName("kpe3")->AsFloat+DM->qOcenka->FieldByName("kpe4")->AsFloat==0)
    {
      LabelKPE->Caption="";
    }
  else
    {
      LabelKPE->Caption=FloatToStrF((DM->qOcenka->FieldByName("kpe1")->AsFloat+DM->qOcenka->FieldByName("kpe2")->AsFloat+
      DM->qOcenka->FieldByName("kpe3")->AsFloat+DM->qOcenka->FieldByName("kpe4")->AsFloat)/4, ffFixed, 2,2)+ " %";
    }

      
 //Замещение работника

  //Резервист
  if (DM->qOcenka->FieldByName("rezerv")->AsString=="1")
    {
      CheckBoxREZERV->Checked=true;
      zrezerv=1;
      CheckBoxREZERV->Enabled = true;
    }
  else
    {
      CheckBoxREZERV->Checked=false;
      zrezerv="NULL";
    }
  //Замещающий
  if (DM->qOcenka->FieldByName("zam")->AsString=="1")
    {
      CheckBoxZAM->Checked=true;
      zzam=1;
      CheckBoxZAM->Enabled = true;
    }
  else
    {
      CheckBoxZAM->Checked=false;
      zzam="NULL";
    }



  if (DM->qOcenka->FieldByName("rezerv")->AsString!="1" && DM->qOcenka->FieldByName("zam")->AsString!="1")
    {
      CheckBoxREZERV->Enabled = true;
      CheckBoxZAM->Enabled = true;
      DM->qZamesh->Filtered = false;
      DM->qZamesh->Active = false;
      DBGridEh1->Enabled = false;
    }

  //Размеры формы
  Panel3->Visible = false;
  Panel1->Align=alClient;
  Zameshenie->Height=440;
  BitBtn1->Top=316;
  BitBtn2->Top=355;
  Bevel3->Height=385;

 /*
  //Вывод информации по замещениям в зависимости от фильтров
  DM->qZamesh->Close();
  DM->qZamesh->Parameters->ParamByName("pgod")->Value=IntToStr(Main->god);
  DM->qZamesh->Parameters->ParamByName("ptype")->Value=IntToStr(type);
  DM->qZamesh->Parameters->ParamByName("ptn")->Value=DM->qOcenka->FieldByName("tn")->AsString;

  try
    {
      DM->qZamesh->Open();
    }
  catch(Exception &E)
    {
      Application->MessageBox(("Невозможно получить данные по замещению выбранного работника "+E.Message).c_str(),"Ошибка",
                               MB_OK+MB_ICONERROR);
      Abort();
    } */


  EditZEX->SetFocus();
}
//---------------------------------------------------------------------------

void __fastcall TZameshenie::CheckBoxZAMClick(TObject *Sender)
{
  int rec;

  if (CheckBoxZAM->Checked==true)
    {
      //Проверка не является ли работник резервистом
      if (DM->qOcenka->FieldByName("rezerv")->AsString=="1")
        {
          Application->MessageBox("Так как данный работник является резервистом, \nон не может одновременно быть и замещающим!!!","Предупреждение",
                                  MB_OK+MB_ICONINFORMATION);

          CheckBoxZAM->Checked = false;
          CheckBoxREZERV->Checked = true;
          CheckBoxREZERV->SetFocus();
          Abort();
        }


      CheckBoxREZERV->Enabled = false;
      CheckBoxZAM->Enabled = true;
      DBGridEh1->Enabled = true;

      //Вывод информации по замещениям в зависимости от фильтров
      DM->qZamesh->Filtered = false;
      DM->qZamesh->Filter = " type=2 and tn="+DM->qOcenka->FieldByName("tn")->AsString;
      if (DM->qZamesh->Active==false) DM->qZamesh->Active=true;
      DM->qZamesh->Filtered = true;
    }
  else
    {
      if (!EditTN->Text.IsEmpty() && EditTN->Text!="490")
        {
          //Вывод информации по замещениям в зависимости от фильтров
          DM->qZamesh->Filtered = false;
          DM->qZamesh->Filter = " type=2 and tn="+DM->qOcenka->FieldByName("tn")->AsString;
          if (DM->qZamesh->Active==false) DM->qZamesh->Active=true;
          DM->qZamesh->Filtered = true;

          if (DM->qZamesh->RecordCount>0)
            {
              if(Application->MessageBox("По данному работнику существуют должности, которые он замещает. Удалить все должности на замещение?","Предупреждение",
                                      MB_YESNO+MB_ICONWARNING)==ID_NO)
                {
                  CheckBoxZAM->Checked=true;
                  CheckBoxZAM->Enabled=true;
                  CheckBoxREZERV->Enabled = false;
                  DBGridEh1->Enabled = true;
                  Abort();
                }
              else
                {
                  rec=DM->qOcenka->RecNo;
                  //Удаление всех замещаемых должностей
                  DM->qZamesh->First();
                  while (!DM->qZamesh->Eof)
                    {
                      DM->qObnovlenie->Close();
                      DM->qObnovlenie->SQL->Clear();
                      DM->qObnovlenie->SQL->Add("delete from ocenka_rez where rowid = chartorowid("+ QuotedStr(DM->qZamesh->FieldByName("rw")->AsString)+")");
                      try
                        {
                          DM->qObnovlenie->ExecSQL();
                        }
                      catch (Exception &E)
                        {
                          Application->MessageBox(("Возникла ошибка при удалении замещаемой должности в таблице OCENKA_REZ "+E.Message).c_str(),"Ошибка",
                                                   MB_OK+MB_ICONERROR);
                          DM->qZamesh->Requery();
                          Main->InsertLog("Возникла ошибка при удалении замещаемой должности("+DM->qZamesh->FieldByName("id_shtat")->AsString+") по работнику: таб.№='"+EditTN->Text+"' ФИО='"+LabelFIO->Caption+"'");
                          Abort();
                        }

                      DM->qZamesh->Next();
                    }
                 if (DM->qObnovlenie->RowsAffected>0)
                   {
                     //Обновление признака в таблице замещения
                     DM->qObnovlenie->Close();
                      DM->qObnovlenie->SQL->Clear();
                      DM->qObnovlenie->SQL->Add("update ocenka set zam=NULL where god="+IntToStr(Main->god)+" and tn="+EditTN->Text);
                      try
                        {
                          DM->qObnovlenie->ExecSQL();
                        }
                      catch (Exception &E)
                        {
                          Application->MessageBox(("Возникла ошибка при обновлении признака замещающего в таблице OCENKA "+E.Message).c_str(),"Ошибка",
                                                   MB_OK+MB_ICONERROR);
                          DM->qZamesh->Requery();
                          //Логи
                          Main->InsertLog("Возникла ошибка при обновлении признака замещающего по работнику: таб.№='"+EditTN->Text+"' ФИО='"+LabelFIO->Caption+"'");
                          Abort();
                        }

                     DM->qOcenka->Requery();
                     DM->qOcenka->RecNo=rec;

                     if (CheckBoxZAM->Checked==true)
                       {
                         CheckBoxZAM->Checked=false;
                         CheckBoxREZERV->Enabled=true;
                       }

                   }
                 
                 //Логи
                 Main->InsertLog("Удаление замещаемых должностей выполнено успешно по работнику: таб.№='"+EditTN->Text+"' ФИО='"+LabelFIO->Caption+"'");
                 
                }
            }
        }

      CheckBoxREZERV->Enabled = true;
      DM->qZamesh->Filtered = false;
      DM->qZamesh->Active=false;
      DBGridEh1->Enabled = false;
    }
}
//---------------------------------------------------------------------------



void __fastcall TZameshenie::EditZEXKeyPress(TObject *Sender, char &Key)
{
  if (!(IsNumeric(Key)||Key=='\b')) Key=0;
}
//---------------------------------------------------------------------------

void __fastcall TZameshenie::EditFIO_RKeyPress(TObject *Sender, char &Key)
{
  if (IsNumeric(Key)) Key=0;        
}
//---------------------------------------------------------------------------


void __fastcall TZameshenie::EditKPE1KeyPress(TObject *Sender, char &Key)
{
  if (! (IsNumeric(Key) || Key=='.' || Key==',' || Key=='/' || Key=='\b') ) Key=0;
  if (Key==',' || Key=='/') Key='.';        
}
//---------------------------------------------------------------------------

//Сохранение замещения
void __fastcall TZameshenie::BitBtn1Click(TObject *Sender)
{
  AnsiString Sql, zam, preem, rezerv, Str;
  int rec;

  if (CheckBoxZAM->Checked) zam=1;
  else zam="NULL";
 // if (CheckBoxPREEM->Checked) preem=1;
 // else preem=0;
  if (CheckBoxREZERV->Checked) rezerv=1;
  else rezerv="NULL";



  //Проверки
  //Цех
  if (EditZEX->Text.IsEmpty())
    {
      Application->MessageBox("Не указан цех работника!!!","Предупреждение",
                               MB_OK+MB_ICONINFORMATION);
      EditZEX->SetFocus();
      Abort();
    }
  //Таб.№
  if (EditTN->Text.IsEmpty() || EditTN->Text=="490")
    {
      Application->MessageBox("Не указан табельный номер работника!!!","Предупреждение",
                               MB_OK+MB_ICONINFORMATION);
      EditTN->SetFocus();
      Abort();
    }

  //Существует ли работник с таким цехом и табельным
  DM->qObnovlenie->Close();
  DM->qObnovlenie->SQL->Clear();
  DM->qObnovlenie->SQL->Add("select * from ocenka where zex="+EditZEX->Text+" and tn="+EditTN->Text);
  try
    {
      DM->qObnovlenie->Open();
    }
  catch(Exception &E)
    {
      Application->MessageBox(("Возникла ошибка при получении данных из таблицы Ocenka" + E.Message).c_str(),"Ошибка",
                              MB_OK+MB_ICONERROR);
      Abort();
    }

  if (DM->qObnovlenie->RecordCount<=0)
    {
      Application->MessageBox("В картотеке не существует работника с указанным цехом и табельным номером!!!","Предупреждение",
                               MB_OK+MB_ICONINFORMATION);
      EditZEX->SetFocus();
      Abort();
    }

  //Проверки раздела "Замещение"**********************************************
  //Отмечено замещение и не указана замещаемая должность
  if (CheckBoxZAM->Checked==true && DM->qZamesh->RecordCount==0)
    {
      Application->MessageBox("Отмечено замещение, но не указана ни одна должность!!!","Предупреждение",
                               MB_OK+MB_ICONINFORMATION);
      CheckBoxZAM->SetFocus();
      Abort();
    }

  //Не отмечено замещение и указана замещаемая должность
  if (CheckBoxZAM->Checked==false && CheckBoxREZERV->Checked==false)
    {
      //Есть ли в таблице замещения по этому работнику
      DM->qObnovlenie->Close();
      DM->qObnovlenie->SQL->Clear();
      DM->qObnovlenie->SQL->Add("select * from ocenka_rez where god="+IntToStr(Main->god)+" and tn="+EditTN->Text);
      try
        {
          DM->qObnovlenie->Open();
        }
      catch(Exception &E)
        {
          Application->MessageBox(("Возникла ошибка при получении данных из таблицы Ocenka" + E.Message).c_str(),"Ошибка",
                                    MB_OK+MB_ICONERROR);
          Abort();
        }

      if (DM->qObnovlenie->RecordCount>0)
        {
          Application->MessageBox("Не отмечено замещение, но есть должность на замещение!!!","Предупреждение",
                                   MB_OK+MB_ICONINFORMATION);
          CheckBoxREZERV->SetFocus();
          Abort();
        }
    }

  //Сохранение
  if (EditZEX->Text!=zzex ||
      EditTN->Text!=ztn ||
      LabelFIO->Caption!=zfio ||
      EditVZ_PENS->Text!=zvz_pens ||
      EditKPE1->Text!=zkpe1 ||
      EditKPE2->Text!=zkpe2 ||
      EditKPE3->Text!=zkpe3 ||
      EditKPE4->Text!=zkpe4 ||
      zrezerv!=rezerv ||
      zzam!=zam
       )
    {
      Sql="update ocenka set  \
                             vz_pens="+Main->SetNull(EditVZ_PENS->Text)+",\
                             kpe1="+Main->SetNull(EditKPE1->Text)+",\
                             kpe2="+Main->SetNull(EditKPE2->Text)+",\
                             kpe3="+Main->SetNull(EditKPE3->Text)+",\
                             kpe4="+Main->SetNull(EditKPE4->Text)+",\
                             zam="+Main->SetNull(zam)+",\
                             rezerv="+Main->SetNull(rezerv)+"\
            where rowid = chartorowid("+ QuotedStr(DM->qOcenka->FieldByName("rw")->AsString)+")";

      DM->qObnovlenie->Close();
      DM->qObnovlenie->SQL->Clear();
      DM->qObnovlenie->SQL->Add(Sql);
      rec=DM->qOcenka->RecNo;
      try
        {
          DM->qObnovlenie->ExecSQL();
        }
      catch(Exception &E)
        {
          Application->MessageBox(("Возникла ошибка при попытке обновления данных в таблице по оценке персонала (OCENKA) "+E.Message).c_str(),"Ошибка",
                                  MB_OK+MB_ICONERROR);
          Abort();
        }

      //Возвращение курсора
      DM->qOcenka->RecNo = rec;

      Str="Обновление замещения по работнику цех="+EditZEX->Text+" таб.№="+EditTN->Text+":";

      if (Main->SetNull(EditVZ_PENS->Text)!=Main->SetNull(zvz_pens)) Str+=" пенс.возраст с '"+Main->SetNull(zvz_pens)+"' на '"+Main->SetNull(EditVZ_PENS->Text)+"',";
      if (Main->SetNull(EditKPE1->Text)!=Main->SetNull(zkpe1)) Str+=", КПЭ за 1кв. с '"+Main->SetNull(zkpe1)+"' на '"+Main->SetNull(EditKPE1->Text)+"'";
      if (Main->SetNull(EditKPE2->Text)!=Main->SetNull(zkpe2)) Str+=", КПЭ за 2кв. с '"+Main->SetNull(zkpe2)+"' на '"+Main->SetNull(EditKPE2->Text)+"'";
      if (Main->SetNull(EditKPE3->Text)!=Main->SetNull(zkpe3)) Str+=", КПЭ за 3кв. с '"+Main->SetNull(zkpe3)+"' на '"+Main->SetNull(EditKPE3->Text)+"'";
      if (Main->SetNull(EditKPE4->Text)!=Main->SetNull(zkpe4)) Str+=", КПЭ за 1кв. с '"+Main->SetNull(zkpe4)+"' на '"+Main->SetNull(EditKPE4->Text)+"'";
      if (Main->SetNull(zzam)!=Main->SetNull(zam)) Str+=", замещение с '"+Main->SetNull(zzam)+"' на '"+Main->SetNull(zam)+"'";
      if (zrezerv!=rezerv) Str+=", резерв с '"+zrezerv+"' на '"+rezerv+"'";
      Str+=" выполнено";
      Main->InsertLog(Str);
      DM->qLogs->Requery();

    }

  Application->MessageBox("Запись успешно сохранена!","Предупреждение",
                               MB_OK+MB_ICONINFORMATION);
  DM->qOcenka->Requery();

//  Zameshenie->Close();

  //Очищение Edit-ов
  EditTN->Text="490";
  EditZEX->Text="";

  LabelFIO->Caption="";
  EditVZ_PENS->Text="";
  ComboBoxGOTOV->ItemIndex=-1;

  EditKPE1->Text="";
  EditKPE2->Text="";
  EditKPE3->Text="";
  EditKPE4->Text="";
  LabelKPE->Caption="";

  CheckBoxZAM->Checked=false;
  CheckBoxREZERV->Checked=false;
  CheckBoxZAM->Enabled=true;
  CheckBoxREZERV->Enabled=true;
  ComboBoxGOTOV->Color=clWindow;
  ComboBoxRISK->Color=clWindow;

  //Скрыть панель редактирования
  Panel3->Visible = false;
  Panel1->Align=alClient;
  Zameshenie->Height=440;
  BitBtn1->Top=316;
  BitBtn2->Top=355;
  Bevel3->Height=385;


  EditZEX->SetFocus();

}
//---------------------------------------------------------------------------

void __fastcall TZameshenie::CheckBox1Click(TObject *Sender)
{
 /* if (CheckBox1->Checked==true)
    {
      if (EditDATN1->Text.IsEmpty())
        {
          EditDATN1->Text=DM->qOcenka->FieldByName("datn1")->AsString;
          EditDATK1->Text=DM->qOcenka->FieldByName("datk1")->AsString;
        }
    }
  else
    {
      EditDATN1->Text="";
      EditDATK1->Text="";
    } */
}
//---------------------------------------------------------------------------

void __fastcall TZameshenie::CheckBox2Click(TObject *Sender)
{
  /*if (CheckBox2->Checked==true)
    {
      if (EditDATN2->Text.IsEmpty())
        {
          EditDATN2->Text=DM->qOcenka->FieldByName("datn2")->AsString;
          EditDATK2->Text=DM->qOcenka->FieldByName("datk2")->AsString;
        }  
    }
  else
    {
      EditDATN2->Text="";
      EditDATK2->Text="";
    } */
}
//---------------------------------------------------------------------------

void __fastcall TZameshenie::CheckBox3Click(TObject *Sender)
{
  /*if (CheckBox3->Checked==true)
    {
      if (EditDATN3->Text.IsEmpty())
        {
          EditDATN3->Text=DM->qOcenka->FieldByName("datn3")->AsString;
          EditDATK3->Text=DM->qOcenka->FieldByName("datk3")->AsString;
        }
    }
  else
    {
      EditDATN3->Text="";
      EditDATK3->Text="";
    } */
}
//---------------------------------------------------------------------------

void __fastcall TZameshenie::CheckBox4Click(TObject *Sender)
{
  /*if (CheckBox4->Checked==true)
    {
      if (EditDATN4->Text.IsEmpty())
        {
          EditDATN4->Text=DM->qOcenka->FieldByName("datn4")->AsString;
          EditDATK4->Text=DM->qOcenka->FieldByName("datk4")->AsString;
        }
    }
  else
    {
      EditDATN4->Text="";
      EditDATK4->Text="";
    } */
}
//---------------------------------------------------------------------------

void __fastcall TZameshenie::CheckBox5Click(TObject *Sender)
{
  /*if (CheckBox5->Checked==true)
    {
      if (EditDATN5->Text.IsEmpty())
        {
          EditDATN5->Text=DM->qOcenka->FieldByName("datn5")->AsString;
          EditDATK5->Text=DM->qOcenka->FieldByName("datk5")->AsString;
        }
    }
  else
    {
      EditDATN5->Text="";
      EditDATK5->Text="";
    }*/
}
//---------------------------------------------------------------------------

void __fastcall TZameshenie::CheckBox6Click(TObject *Sender)
{
 /* if (CheckBox6->Checked==true)
    {
      if (EditDATN6->Text.IsEmpty())
        {
          EditDATN6->Text=DM->qOcenka->FieldByName("datn6")->AsString;
          EditDATK6->Text=DM->qOcenka->FieldByName("datk6")->AsString;
        }
    }
  else
    {
      EditDATN6->Text="";
      EditDATK6->Text="";
    } */
}
//---------------------------------------------------------------------------

void __fastcall TZameshenie::CheckBox7Click(TObject *Sender)
{
 /* if (CheckBox7->Checked==true)
    {
      if (EditDATN7->Text.IsEmpty())
        {
          EditDATN7->Text=DM->qOcenka->FieldByName("datn7")->AsString;
          EditDATK7->Text=DM->qOcenka->FieldByName("datk7")->AsString;
        }
    }
  else
    {
      EditDATN7->Text="";
      EditDATK7->Text="";
    }  */
}
//---------------------------------------------------------------------------

void __fastcall TZameshenie::CheckBox8Click(TObject *Sender)
{
  /*if (CheckBox8->Checked==true)
    {
      if (EditDATN8->Text.IsEmpty())
        {
          EditDATN8->Text=DM->qOcenka->FieldByName("datn8")->AsString;
          EditDATK8->Text=DM->qOcenka->FieldByName("datk8")->AsString;
        }
    }
  else
    {
      EditDATN8->Text="";
      EditDATK8->Text="";
    }*/
}
//---------------------------------------------------------------------------

void __fastcall TZameshenie::CheckBox9Click(TObject *Sender)
{
  /*if (CheckBox9->Checked==true)
    {
      if (EditDATN9->Text.IsEmpty())
        {
          EditDATN9->Text=DM->qOcenka->FieldByName("datn9")->AsString;
          EditDATK9->Text=DM->qOcenka->FieldByName("datk9")->AsString;
        }
    }
  else
    {
      EditDATN9->Text="";
      EditDATK9->Text="";
    } */
}
//---------------------------------------------------------------------------

void __fastcall TZameshenie::CheckBox10Click(TObject *Sender)
{
  /*if (CheckBox10->Checked==true)
    {
      if (EditDATN10->Text.IsEmpty())
        {
          EditDATN10->Text=DM->qOcenka->FieldByName("datn10")->AsString;
          EditDATK10->Text=DM->qOcenka->FieldByName("datk10")->AsString;
        }
    }
  else
    {
      EditDATN10->Text="";
      EditDATK10->Text="";
    } */
}
//---------------------------------------------------------------------------

void __fastcall TZameshenie::CheckBox11Click(TObject *Sender)
{
  /*if (CheckBox11->Checked==true)
    {
      if (EditDATN11->Text.IsEmpty())
        {
          EditDATN11->Text=DM->qOcenka->FieldByName("datn11")->AsString;
          EditDATK11->Text=DM->qOcenka->FieldByName("datk11")->AsString;
        }
    }
  else
    {
      EditDATN11->Text="";
      EditDATK11->Text="";
    } */
}
//---------------------------------------------------------------------------

void __fastcall TZameshenie::CheckBox12Click(TObject *Sender)
{
  /*if (CheckBox12->Checked==true)
    {
      if (EditDATN12->Text.IsEmpty())
        {
          EditDATN12->Text=DM->qOcenka->FieldByName("datn12")->AsString;
          EditDATK12->Text=DM->qOcenka->FieldByName("datk12")->AsString;
        }
    }
  else
    {
      EditDATN12->Text="";
      EditDATK12->Text="";
    }*/
}
//---------------------------------------------------------------------------

//Изменение таб.№ работника
void __fastcall TZameshenie::EditTNChange(TObject *Sender)
{
  //Если меняется табельный
  TLocateOptions SearchOptions;

  if (!EditZEX->Text.IsEmpty())
    {
      Variant locvalues[] = {Main->SetNull(EditZEX->Text), Main->SetNull(EditTN->Text)};

      if (!DM->qOcenka->Locate("zex;tn", VarArrayOf(locvalues, 1),
                                         SearchOptions << loCaseInsensitive) )
        {
          //Не найдеен работник

          //Очищение Edit-ов
          LabelFIO->Caption="";
          EditVZ_PENS->Text="";
          ComboBoxGOTOV->ItemIndex=-1;
          ComboBoxGOTOV->Color=clWindow;

          EditKPE1->Text="";
          EditKPE2->Text="";
          EditKPE3->Text="";
          EditKPE4->Text="";

          CheckBoxZAM->Checked=false;
          CheckBoxREZERV->Checked=false;
          CheckBoxREZERV->Enabled = false;
          CheckBoxZAM->Enabled = false;

          DM->qZamesh->Filtered = false;
          DM->qZamesh->Active = false;
          DBGridEh1->Enabled = false;

          //Размеры формы
          Panel3->Visible = false;
          Panel1->Align=alClient;
          Zameshenie->Height=440;
          BitBtn1->Top=316;
          BitBtn2->Top=355;
          Bevel3->Height=385;

          Abort();
        }
      else
        {
          //Заполнение Edit-ов
          //Тн, цех, ФИО
          EditZEX->Text=zzex=DM->qOcenka->FieldByName("zex")->AsString;
          EditTN->Text=ztn=DM->qOcenka->FieldByName("tn")->AsString;
          LabelFIO->Caption=zfio=DM->qOcenka->FieldByName("fio")->AsString;


          //Пенс. возраст и готовность
          EditVZ_PENS->Text=zvz_pens=DM->qOcenka->FieldByName("vz_pens")->AsString;
          if (DM->qOcenka->FieldByName("gotov")->AsString==1)
            {
              ComboBoxGOTOV->ItemIndex=ComboBoxGOTOV->Items->IndexOf("низкая");
              ComboBoxGOTOV->Color=(TColor)0x008080FF;
              zgotov=1;
            }
          else if (DM->qOcenka->FieldByName("gotov")->AsString==2)
            {
              ComboBoxGOTOV->ItemIndex=ComboBoxGOTOV->Items->IndexOf("средняя");
              ComboBoxGOTOV->Color=(TColor)0x0080FFFF;
              zgotov=2;
            }
          else if (DM->qOcenka->FieldByName("gotov")->AsString==3)
            {
              ComboBoxGOTOV->ItemIndex=ComboBoxGOTOV->Items->IndexOf("высокая");
              ComboBoxGOTOV->Color=clMoneyGreen;
              zgotov=3;
            }
          else
            {
              ComboBoxGOTOV->ItemIndex=-1;
              ComboBoxGOTOV->Color=clWindow;
              zgotov="NULL";
            }


          //КПЭ
          EditKPE1->Text=zkpe1=DM->qOcenka->FieldByName("kpe1")->AsString;
          EditKPE2->Text=zkpe2=DM->qOcenka->FieldByName("kpe2")->AsString;
          EditKPE3->Text=zkpe3=DM->qOcenka->FieldByName("kpe3")->AsString;
          EditKPE4->Text=zkpe4=DM->qOcenka->FieldByName("kpe4")->AsString;

          if (DM->qOcenka->FieldByName("kpe1")->AsFloat+DM->qOcenka->FieldByName("kpe2")->AsFloat+
              DM->qOcenka->FieldByName("kpe3")->AsFloat+DM->qOcenka->FieldByName("kpe4")->AsFloat==0)
            {
              LabelKPE->Caption="";
            }
          else
            {
              LabelKPE->Caption=FloatToStrF((DM->qOcenka->FieldByName("kpe1")->AsFloat+DM->qOcenka->FieldByName("kpe2")->AsFloat+
              DM->qOcenka->FieldByName("kpe3")->AsFloat+DM->qOcenka->FieldByName("kpe4")->AsFloat)/4, ffFixed, 2,2)+ " %";
            }


          //Замещение работника
          //Замещающий
          if (DM->qOcenka->FieldByName("zam")->AsString=="1")
            {
              CheckBoxZAM->Checked=true;
              zzam=1;
            }
          else
            {
              CheckBoxZAM->Checked=false;
              zzam="NULL";
            }

          //Резервист
          if (DM->qOcenka->FieldByName("rezerv")->AsString=="1")
            {
              CheckBoxREZERV->Checked=true;
              zrezerv=1;
            }
          else
            {
              CheckBoxREZERV->Checked=false;
              zrezerv="NULL";
            }

          if (DM->qOcenka->FieldByName("rezerv")->AsString!="1" && DM->qOcenka->FieldByName("zam")->AsString!="1")
            {
              CheckBoxREZERV->Enabled = true;
              CheckBoxZAM->Enabled = true;
              DM->qZamesh->Filtered = false;
              DM->qZamesh->Active = false;
              DBGridEh1->Enabled = false;
            }

          //Размеры формы
          Panel3->Visible = false;
          Panel1->Align=alClient;
          Zameshenie->Height=440;
          BitBtn1->Top=316;
          BitBtn2->Top=355;
          Bevel3->Height=385;

        }
    }
}
//---------------------------------------------------------------------------

//Изменение цеха при введенной записи
void __fastcall TZameshenie::EditZEXChange(TObject *Sender)
{

  if (!EditTN->Text.IsEmpty() && EditTN->Text!="490")
    {
      TLocateOptions SearchOptions;

      Variant locvalues[] = {Main->SetNull(EditZEX->Text), Main->SetNull(EditTN->Text)};

      if (!DM->qOcenka->Locate("zex;tn", VarArrayOf(locvalues, 1),
                                         SearchOptions << loCaseInsensitive) )
        {
          //Не найдеен работник

          //Очищение Edit-ов
          LabelFIO->Caption="";
          EditVZ_PENS->Text="";
          ComboBoxGOTOV->ItemIndex=-1;
          ComboBoxGOTOV->Color=clWindow;

          EditKPE1->Text="";
          EditKPE2->Text="";
          EditKPE3->Text="";
          EditKPE4->Text="";

          CheckBoxZAM->Checked=false;
          CheckBoxREZERV->Checked=false;
          CheckBoxREZERV->Enabled = false;
          CheckBoxZAM->Enabled = false;

          DM->qZamesh->Filtered = false;
          DM->qZamesh->Active = false;
          DBGridEh1->Enabled = false;

          //Размеры формы
          Panel3->Visible = false;
          Panel1->Align=alClient;
          Zameshenie->Height=440;
          BitBtn1->Top=316;
          BitBtn2->Top=355;
          Bevel3->Height=385;

          Abort();
        }
      else
        {
          //Заполнение Edit-ов
          //Тн, цех, ФИО
          EditZEX->Text=zzex=DM->qOcenka->FieldByName("zex")->AsString;
          EditTN->Text=ztn=DM->qOcenka->FieldByName("tn")->AsString;
          LabelFIO->Caption=zfio=DM->qOcenka->FieldByName("fio")->AsString;


          //Пенс. возраст и готовность
          EditVZ_PENS->Text=zvz_pens=DM->qOcenka->FieldByName("vz_pens")->AsString;
          if (DM->qOcenka->FieldByName("gotov")->AsString==1)
            {
              ComboBoxGOTOV->ItemIndex=ComboBoxGOTOV->Items->IndexOf("низкая");
              ComboBoxGOTOV->Color=(TColor)0x008080FF;
              zgotov=1;
            }
          else if (DM->qOcenka->FieldByName("gotov")->AsString==2)
            {
              ComboBoxGOTOV->ItemIndex=ComboBoxGOTOV->Items->IndexOf("средняя");
              ComboBoxGOTOV->Color=(TColor)0x0080FFFF;
              zgotov=2;
            }
          else if (DM->qOcenka->FieldByName("gotov")->AsString==3)
            {
              ComboBoxGOTOV->ItemIndex=ComboBoxGOTOV->Items->IndexOf("высокая");
              ComboBoxGOTOV->Color=clMoneyGreen;
              zgotov=3;
            }
          else
            {
              ComboBoxGOTOV->ItemIndex=-1;
              ComboBoxGOTOV->Color=clWindow;
              zgotov="NULL";
            }


          //КПЭ
          EditKPE1->Text=zkpe1=DM->qOcenka->FieldByName("kpe1")->AsString;
          EditKPE2->Text=zkpe2=DM->qOcenka->FieldByName("kpe2")->AsString;
          EditKPE3->Text=zkpe3=DM->qOcenka->FieldByName("kpe3")->AsString;
          EditKPE4->Text=zkpe4=DM->qOcenka->FieldByName("kpe4")->AsString;

          if (DM->qOcenka->FieldByName("kpe1")->AsFloat+DM->qOcenka->FieldByName("kpe2")->AsFloat+
              DM->qOcenka->FieldByName("kpe3")->AsFloat+DM->qOcenka->FieldByName("kpe4")->AsFloat==0)
            {
              LabelKPE->Caption="";
            }
          else
            {
              LabelKPE->Caption=FloatToStrF((DM->qOcenka->FieldByName("kpe1")->AsFloat+DM->qOcenka->FieldByName("kpe2")->AsFloat+
              DM->qOcenka->FieldByName("kpe3")->AsFloat+DM->qOcenka->FieldByName("kpe4")->AsFloat)/4, ffFixed, 2,2)+ " %";
            }


          //Замещение работника
          //Замещающий
          if (DM->qOcenka->FieldByName("zam")->AsString=="1")
            {
              CheckBoxZAM->Checked=true;
              zzam=1;
            }
          else
            {
              CheckBoxZAM->Checked=false;
              zzam="NULL";
            }

          //Резервист
          if (DM->qOcenka->FieldByName("rezerv")->AsString=="1")
            {
              CheckBoxREZERV->Checked=true;
              zrezerv=1;
            }
          else
            {
              CheckBoxREZERV->Checked=false;
              zrezerv="NULL";
            }

          if (DM->qOcenka->FieldByName("rezerv")->AsString!="1" && DM->qOcenka->FieldByName("zam")->AsString!="1")
            {
              CheckBoxREZERV->Enabled = true;
              CheckBoxZAM->Enabled = true;
              DM->qZamesh->Filtered = false;
              DM->qZamesh->Active = false;
              DBGridEh1->Enabled = false;
            }

          //Размеры формы
          Panel3->Visible = false;
          Panel1->Align=alClient;
          Zameshenie->Height=440;
          BitBtn1->Top=316;
          BitBtn2->Top=355;
          Bevel3->Height=385;

        }
    }
}
//---------------------------------------------------------------------------

void __fastcall TZameshenie::FormCreate(TObject *Sender)
{
  EditZEX->Text="";
  EditTN->Text="";
  StringGrid1->Cells[0][0]="c";
  StringGrid1->Cells[1][0]="по";
}
//---------------------------------------------------------------------------

void __fastcall TZameshenie::ComboBoxGOTOVChange(TObject *Sender)
{
  if (ComboBoxGOTOV->ItemIndex==0) ComboBoxGOTOV->Color=(TColor)0x008080FF;
  else if (ComboBoxGOTOV->ItemIndex==1) ComboBoxGOTOV->Color=(TColor)0x0080FFFF;
  else if (ComboBoxGOTOV->ItemIndex==2) ComboBoxGOTOV->Color=clMoneyGreen;
  else ComboBoxGOTOV->Color=clWindow;
}
//---------------------------------------------------------------------------

void __fastcall TZameshenie::ComboBoxRISKChange(TObject *Sender)
{
  if (ComboBoxRISK->ItemIndex==0) ComboBoxRISK->Color=clMoneyGreen;
  else if (ComboBoxRISK->ItemIndex==1) ComboBoxRISK->Color=(TColor)0x0080FFFF;
  else if (ComboBoxRISK->ItemIndex==2) ComboBoxRISK->Color=(TColor)0x008080FF;
  else ComboBoxRISK->Color=clWindow;
}
//---------------------------------------------------------------------------

void __fastcall TZameshenie::EditZEX_ZAMChange(TObject *Sender)
{
  //ВЫвод цеха и участка
  if (!EditZEX_ZAM->Text.IsEmpty())
    {
      DM->qDolg->Filtered = false;
      TLocateOptions SearchOptions;
      DM->qDolg->Locate("shifr_zex", EditZEX_ZAM->Text,
                            SearchOptions << loCaseInsensitive);

      if (EditZEX_ZAM->Text.Length()>2) LabelZEX_ZAM->Caption = DM->qDolg->FieldByName("nzex")->AsString +"\n"+ DM->qDolg->FieldByName("uch")->AsString;
      else LabelZEX_ZAM->Caption = DM->qDolg->FieldByName("nzex")->AsString;

      if (!EditSHIFR_ZAM->Text.IsEmpty())
        {
          //КРД должность
          //Проверка является ли должность КРД
          DM->qObnovlenie->Close();
          DM->qObnovlenie->SQL->Clear();
          DM->qObnovlenie->SQL->Add("select * from sp_ocenka_krd \
                                     where zex="+QuotedStr(EditZEX_ZAM->Text)+"\
                                     and shifr_dolg=(select short from p1000@sapmig_buffdb where otype='S' and langu='R' and objid="+QuotedStr(EditSHIFR_ZAM->Text)+")");
          try
            {
              DM->qObnovlenie->Open();
            }
          catch(Exception &E)
            {
              Application->MessageBox(("Невозможно получить данные из справочника КРД (SP_OCENKA_KRD)"+E.Message).c_str(),"Ошибка",
                                        MB_OK+MB_ICONERROR);
            }

          if (DM->qObnovlenie->RecordCount>0) LabelKRD->Caption="Ключевая Резервная Должность";
          else LabelKRD->Caption="";
        }
    }
  else LabelZEX_ZAM->Caption =""; 
}
//---------------------------------------------------------------------------

void __fastcall TZameshenie::EditSHIFR_ZAMChange(TObject *Sender)
{
  AnsiString Sql;

  if (!EditSHIFR_ZAM->Text.IsEmpty())
    {
      Sql=" select vse.zec, vse.nzec as nzec, vse.dolg as dolg, vse.id_dolg as id_dolg, \
                   ruk.tn_sap as tn_sap, (ruk.fam||' '||ruk.im||' '||ruk.ot) as fio                                 \
            from                                                                                                 \
                (select prof.*, cex.short as zec, cex.stext as nzec                                              \
                 from                                                                                            \
                     (select pr.objid as shtat, pr.short as id_dolg, pr.stext as dolg, s.sobid as sobid          \
                      from p1000@sapmig_buffdb pr                                                                \
                      left join                                                                                  \
                      p1001@sapmig_buffdb s on pr.objid=s.objid                                                  \
                      where pr.otype='S' and pr.langu='R' and s.otype='S' and s.sclas='O' and pr.objid="+QuotedStr(EditSHIFR_ZAM->Text)+"\
                      ) prof                                                                                     \
                     left join                                                                                   \
                     p1000@sapmig_buffdb cex                                                                     \
                     on prof.sobid=cex.objid and cex.otype='O' and cex.langu='R'                                 \
                     ) vse                                                                                       \
                 left join sap_osn_sved ruk                                                                      \
                 on vse.shtat=ruk.id_shtat";

      //Sql="select stext from p1000@sapmig_buffdb where otype='S' and langu='R' and objid=:pkod_prof";


      DM->qObnovlenie->Close();
      DM->qObnovlenie->SQL->Clear();
      DM->qObnovlenie->SQL->Add(Sql);
      try
        {
          DM->qObnovlenie->Open();
        }
      catch(Exception &E)
        {
          Application->MessageBox(("Возникла ошибка при попытке доступа к справочнику должностей (P1000)"+ E.Message).c_str(),"Ошибка",
                                  MB_OK+MB_ICONERROR);
        }


     // LabelZEX_ZAM->Caption=DM->qObnovlenie->FieldByName("nzec")->AsString;
      LabelDOLG_ZAM->Caption=DM->qObnovlenie->FieldByName("dolg")->AsString;
      LabelTN_R->Caption=DM->qObnovlenie->FieldByName("tn_sap")->AsString;
      EditFIO_R->Text=DM->qObnovlenie->FieldByName("fio")->AsString;
      id_dolg=DM->qObnovlenie->FieldByName("id_dolg")->AsString;
      EditZEX_ZAM->Text=DM->qObnovlenie->FieldByName("zec")->AsString;


  if (!LabelTN_R->Caption.IsEmpty() && ComboBoxRISK->ItemIndex==-1)
    {
      Sql="select risk, risk_prich from ocenka where tn="+LabelTN_R->Caption+" and god="+IntToStr(Main->god);

      DM->qObnovlenie->Close();
      DM->qObnovlenie->SQL->Clear();
      DM->qObnovlenie->SQL->Add(Sql);
      try
        {
          DM->qObnovlenie->Open();
        }
      catch(Exception &E)
        {
          Application->MessageBox(("Возникла ошибка при выборке данных из картотеки по оценке персонала (Ocenka)"+ E.Message).c_str(),"Ошибка",
                                    MB_OK+MB_ICONERROR);
        }

      //Риск и причина
      if (DM->qObnovlenie->FieldByName("risk")->AsString==1)
        {
          ComboBoxRISK->ItemIndex=ComboBoxRISK->Items->IndexOf("низкий");
          ComboBoxRISK->Color=clMoneyGreen;
        }
      else if (DM->qObnovlenie->FieldByName("risk")->AsString==2)
        {
          ComboBoxRISK->ItemIndex=ComboBoxRISK->Items->IndexOf("средний");
          ComboBoxRISK->Color=(TColor)0x0080FFFF;
        }
      else if (DM->qObnovlenie->FieldByName("risk")->AsString==3)
        {
          ComboBoxRISK->ItemIndex=ComboBoxRISK->Items->IndexOf("высокий");
          ComboBoxRISK->Color=(TColor)0x008080FF;
        }
      else
        {
          ComboBoxRISK->ItemIndex=-1;
          ComboBoxRISK->Color=clWindow;
        }
      EditRISK_PRICH->Text=DM->qObnovlenie->FieldByName("risk_prich")->AsString;
    }

    if (!EditZEX_ZAM->Text.IsEmpty())
      {
        //КРД должность
        //Проверка является ли должность КРД
        DM->qObnovlenie->Close();
        DM->qObnovlenie->SQL->Clear();
        DM->qObnovlenie->SQL->Add("select * from sp_ocenka_krd \
                                   where zex="+QuotedStr(EditZEX_ZAM->Text)+"\
                                   and shifr_dolg=(select short from p1000@sapmig_buffdb where otype='S' and langu='R' and objid="+QuotedStr(EditSHIFR_ZAM->Text)+")");
        try
          {
            DM->qObnovlenie->Open();
          }
        catch(Exception &E)
          {
            Application->MessageBox(("Невозможно получить данные из справочника КРД (SP_OCENKA_KRD)"+E.Message).c_str(),"Ошибка",
                                     MB_OK+MB_ICONERROR);
          }

        if (DM->qObnovlenie->RecordCount>0) LabelKRD->Caption="Ключевая Резервная Должность";
        else LabelKRD->Caption="";
      }
    }
  else
    {
      LabelDOLG_ZAM->Caption="";
    }
}
//---------------------------------------------------------------------------

void __fastcall TZameshenie::EditFIO_RChange(TObject *Sender)
{
/*  if (!EditFIO_R->Text.IsEmpty())
    {
      DM->qObnovlenie->Close();
      DM->qObnovlenie->SQL->Clear();
      DM->qObnovlenie->SQL->Add("select tn_sap from sap_work \
                                 where upper(regexp_replace("+QuotedStr(EditFIO_R->Text)+", ' (.*)'))=upper(fam) \
                                 and upper(regexp_replace("+QuotedStr(EditFIO_R->Text)+", ' (.*)|^[^ ]* '))=upper(im) \
                                 and upper(regexp_replace("+QuotedStr(EditFIO_R->Text)+", '(.*) '))=upper(ot)");
      DM->qObnovlenie->Open();


      LabelTN_R->Caption=DM->qObnovlenie->FieldByName("tn_sap")->AsString;
    }*/        
}
//---------------------------------------------------------------------------

void __fastcall TZameshenie::EditZEXKeyDown(TObject *Sender, WORD &Key,
      TShiftState Shift)
{
  if (Key==VK_RETURN)
  FindNextControl((TWinControl *)Sender, true, true, false)->SetFocus();
  EditTN->SelStart=EditTN->Text.Length();
}
//---------------------------------------------------------------------------

void __fastcall TZameshenie::StringGrid1DrawCell(TObject *Sender, int ACol,
      int ARow, TRect &Rect, TGridDrawState State)
{
  int x,y;
  TDateTime d;


  x = Rect.left+(Rect.Width() - StringGrid1->Canvas->TextWidth(StringGrid1->Cells[ACol][ARow]))/2;
  y = Rect.top+(Rect.Height() - StringGrid1->Canvas->TextHeight(StringGrid1->Cells[ACol][ARow]))/2;


  // выделение цветом активной записи
  if (State.Contains(gdSelected))
    {
      StringGrid1->Canvas->Brush->Color =TColor(0x00C8F7E3);// clInfoBk;
      StringGrid1->Canvas->Font->Color = clBlack;
      StringGrid1->Canvas->FillRect(Rect);
      StringGrid1->Canvas->TextOut(x,y,StringGrid1->Cells[ACol][ARow]);
    }

  // убрать выделение активной ячейки при выходе из StringGrid1
  if (ActiveControl != StringGrid1)
    {
      if (State.Contains(gdSelected))
        {
          StringGrid1->Canvas->Brush->Color = clWhite;
          StringGrid1->Canvas->Font->Color = clBlack;
          StringGrid1->Canvas->FillRect(Rect);
          StringGrid1->Canvas->TextOut(x,y,StringGrid1->Cells[ACol][ARow]);
        }
    }

  if (ARow==0)
    {
      StringGrid1->Canvas->Font->Color=clBlack;
      StringGrid1->Canvas->Font->Style=TFontStyles()<<fsBold;
      //DrawText(StringGrid1->Canvas->Handle, StringGrid1->Cells[ACol][ARow].c_str(), strlen(StringGrid1->Cells[ACol][ARow].c_str()), &Rect, DT_WORDBREAK); // Выводим текст в ячейку используя ф-цию WinAPI
    }

  //Изменение цвета, если не верно введена дата
  // Проверка на правильность ввода даты

/*      if  (ARow!=0) {
          if(!TryStrToDate(StringGrid1->Cells[ARow][ACol],d) )
            {

              StringGrid1->Canvas->Brush->Color =TColor(0x00C8F7E3);// clInfoBk;
      StringGrid1->Canvas->Font->Color = clRed;
      StringGrid1->Canvas->FillRect(Rect);
      StringGrid1->Canvas->TextOut(x,y,StringGrid1->Cells[ACol][ARow]);



          //    StringGrid1->Canvas->Brush->Color = clRed;
              //StringGrid1->Canvas->Font->Color=clRed;//   TFontStyles()<<fsBold;

  // StringGrid1.Canvas.Font.Color := clWhite;


        //       StringGrid1->Canvas->TextOut(Rect.Left, Rect.Top, StringGrid1->Cells[ACol][ ARow]);

            }
          else
            {
              StringGrid1->Cells[ARow][ACol]=FormatDateTime("dd.mm.yyyy",d);
           //   StringGrid1->Canvas->Brush->Color = clBlack;
             // StringGrid1->Canvas->Font->Color=clBlack;


             StringGrid1->Canvas->Brush->Color =TColor(0x00C8F7E3);// clInfoBk;
      StringGrid1->Canvas->Font->Color = clBlack;
      StringGrid1->Canvas->FillRect(Rect);
      StringGrid1->Canvas->TextOut(x,y,StringGrid1->Cells[ACol][ARow]);
            }

          }   */


 /*  int i=1;

*/


  if(!TryStrToDate(StringGrid1->Cells[StringGrid1->Col][StringGrid1->Row],d) && ARow!=0 && StringGrid1->Cells[StringGrid1->Col][StringGrid1->Row]!="" )
    {
      StringGrid1->Canvas->Font->Color = clRed;
      StringGrid1->Canvas->FillRect(Rect);
      StringGrid1->Canvas->TextOut(x,y,StringGrid1->Cells[StringGrid1->Col][StringGrid1->Row]);
    }
  else
    {
      StringGrid1->Canvas->Font->Color = clBlack;
      StringGrid1->Canvas->FillRect(Rect);
      StringGrid1->Canvas->TextOut(x,y,StringGrid1->Cells[ACol][ARow]);
    }

  StringGrid1->Canvas->Brush->Color = clGreen;


 /* StringGrid1->Canvas->FillRect(Rect);
  StringGrid1->Canvas->TextOut(x,y,StringGrid1->Cells[ACol][ARow]);  */


  // Здесь прорисовка текста
  /*  StringGrid1->Canvas->FillRect(Rect);
      DrawText(StringGrid1->Canvas->Handle, StringGrid1->Cells[ACol][ARow].c_str(), strlen(StringGrid1->Cells[ACol][ARow].c_str()),
      &Rect, DT_WORDBREAK); // Выводим текст в ячейку используя ф-цию WinAPI
  */


}
//---------------------------------------------------------------------------

void __fastcall TZameshenie::StringGrid1SetEditText(TObject *Sender,
      int ACol, int ARow, const AnsiString Value)
{
/* TDateTime d;

  if (ActiveControl == BitBtn2)
    {
      Zameshenie->Close();
    }
  else
    {
      //int i= StringGrid1->Col;
      //int j= StringGrid1->Row;
      if (StringGrid1->Cells[ACol][ARow]!="")
        {

          // Добавление к дате отчетного месяца и года
          if (StringGrid1->Cells[ACol][ARow].Length()<3)
            {
              if(StringGrid1->Cells[ACol][ARow].Pos("."))
                {
                  Application->MessageBox("Неверный формат даты","Ошибка", MB_OK+MB_ICONINFORMATION);
         //         EditDATN1->Font->Color = clRed;
                   StringGrid1->Canvas->Brush->Color =clRed;

                  StringGrid1->Row=ARow;
                  StringGrid1->Col=ACol;
                  Abort();
                }
              else
                {
                  StringGrid1->Cells[ACol][ARow] = StringGrid1->Cells[ACol][ARow]+ "."+ DateToStr(Date()).SubString(4,2) +"."+ DateToStr(Date()).SubString(7,5);
                  //EditDATN1->Font->Color = clBlack;
                }
            }

          // Проверка на правильность ввода даты
          if(!TryStrToDate(StringGrid1->Cells[ACol][ARow],d))
            {
              Application->MessageBox("Неверный формат даты","Ошибка", MB_OK);
          //    EditDATN1->Font->Color = clRed;
              StringGrid1->Row=ARow;
              StringGrid1->Col=ACol;
            }
          else
            {
              StringGrid1->Cells[ACol][ARow]=FormatDateTime("dd.mm.yyyy",d);
            //  EditDATN1->Font->Color = clBlack;
            }

        }
    }   */
}
//---------------------------------------------------------------------------



void __fastcall TZameshenie::StringGrid2SetEditText(TObject *Sender,
      int ACol, int ARow, const AnsiString Value)
{
 TDateTime d;

  if (ActiveControl == BitBtn2)
    {
      Zameshenie->Close();
    }
  else
    {
      //int i= StringGrid1->Col;
      //int j= StringGrid1->Row;
      if (StringGrid1->Cells[ACol][ARow]!="")
        {

          // Добавление к дате отчетного месяца и года
          if (StringGrid1->Cells[ACol][ARow].Length()<3)
            {
              if(StringGrid1->Cells[ACol][ARow].Pos("."))
                {
                  Application->MessageBox("Неверный формат даты","Ошибка", MB_OK+MB_ICONINFORMATION);
         //         EditDATN1->Font->Color = clRed;
                   StringGrid1->Canvas->Brush->Color =clRed;

                  StringGrid1->Row=ARow;
                  StringGrid1->Col=ACol;
                  Abort();
                }
              else
                {
                  StringGrid1->Cells[ACol][ARow] = StringGrid1->Cells[ACol][ARow]+ "."+ DateToStr(Date()).SubString(4,2) +"."+ DateToStr(Date()).SubString(7,5);
                  //EditDATN1->Font->Color = clBlack;
                }
            }

          // Проверка на правильность ввода даты
          if(!TryStrToDate(StringGrid1->Cells[ACol][ARow],d))
            {
              Application->MessageBox("Неверный формат даты","Ошибка", MB_OK);
          //    EditDATN1->Font->Color = clRed;
              StringGrid1->Row=ARow;
              StringGrid1->Col=ACol;
            }
          else
            {
              StringGrid1->Cells[ACol][ARow]=FormatDateTime("dd.mm.yyyy",d);
            //  EditDATN1->Font->Color = clBlack;
            }

        }
    }
}
//---------------------------------------------------------------------------

void __fastcall TZameshenie::StringGrid1SelectCell(TObject *Sender,
      int ACol, int ARow, bool &CanSelect)
{
   TDateTime d;

if (ActiveControl == BitBtn2)
    {
      Zameshenie->Close();
    }
  else
    {
      int i= StringGrid1->Col;
      int j= StringGrid1->Row;
      if (StringGrid1->Cells[i][j]!="")
        {

          // Добавление к дате отчетного месяца и года
          if (StringGrid1->Cells[i][j].Length()<3)
            {
              if(StringGrid1->Cells[i][j].Pos("."))
                {
                  Application->MessageBox("Неверный формат даты","Ошибка", MB_OK+MB_ICONINFORMATION);
         //         EditDATN1->Font->Color = clRed;
                   StringGrid1->Canvas->Brush->Color =clRed;

                  StringGrid1->Row=StringGrid1->Row;
                  StringGrid1->Col=StringGrid1->Col;
               //   StringGrid1->Options << goEditing;
             //     CanSelect = true;
                  Abort();
                }
              else
                {
                  StringGrid1->Cells[i][j] = StringGrid1->Cells[i][j]+ "."+ DateToStr(Date()).SubString(4,2) +"."+ DateToStr(Date()).SubString(7,5);
                  //EditDATN1->Font->Color = clBlack;
                }
            }

          // Проверка на правильность ввода даты
          if(!TryStrToDate(StringGrid1->Cells[i][j],d))
            {
             Application->MessageBox("Неверный формат даты","Ошибка", MB_OK+MB_ICONINFORMATION);
             // Application->MessageBox("Неверный формат даты","Ошибка", MB_OK);
          //    EditDATN1->Font->Color = clRed;

              //StringGrid1->Cells[i][j];

              StringGrid1->Row=StringGrid1->Row;
              StringGrid1->Col=StringGrid1->Col;
              Abort();
          //    StringGrid1->Options << goEditing;
            //  CanSelect = true;



            }
          else
            {
              StringGrid1->Cells[i][j]=FormatDateTime("dd.mm.yyyy",d);
            //  EditDATN1->Font->Color = clBlack;
            }

        }
    }

}
//---------------------------------------------------------------------------

void __fastcall TZameshenie::StringGrid1Enter(TObject *Sender)
{
  StringGrid1->Col = 0;
  StringGrid1->Row = 1;
}
//---------------------------------------------------------------------------




void __fastcall TZameshenie::StringGrid1Exit(TObject *Sender)
{
   TDateTime d;

if (ActiveControl == BitBtn2)
    {
      Zameshenie->Close();
    }
  else
    {
      int i= StringGrid1->Col;
      int j= StringGrid1->Row;
      if (StringGrid1->Cells[i][j]!="")
        {

          // Добавление к дате отчетного месяца и года
          if (StringGrid1->Cells[i][j].Length()<3)
            {
              if(StringGrid1->Cells[i][j].Pos("."))
                {
                  Application->MessageBox("Неверный формат даты","Ошибка", MB_OK+MB_ICONINFORMATION);
         //         EditDATN1->Font->Color = clRed;
                   StringGrid1->Canvas->Brush->Color =clRed;

                  StringGrid1->Row=StringGrid1->Row;
                  StringGrid1->Col=StringGrid1->Col;
                  StringGrid1->SetFocus();
               //   StringGrid1->Options << goEditing;
             //     CanSelect = true;
                  Abort();
                }
              else
                {
                  StringGrid1->Cells[i][j] = StringGrid1->Cells[i][j]+ "."+ DateToStr(Date()).SubString(4,2) +"."+ DateToStr(Date()).SubString(7,5);
                  //EditDATN1->Font->Color = clBlack;
                }
            }

          // Проверка на правильность ввода даты
          if(!TryStrToDate(StringGrid1->Cells[i][j],d))
            {
           //   Application->MessageBox("Неверный формат даты","Ошибка", MB_OK);
          //    EditDATN1->Font->Color = clRed;

              //StringGrid1->Cells[i][j];

              StringGrid1->Row=StringGrid1->Row;
              StringGrid1->Col=StringGrid1->Col;

              StringGrid1->SetFocus();
              Abort();
          //    StringGrid1->Options << goEditing;
            //  CanSelect = true;



            }
          else
            {
              StringGrid1->Cells[i][j]=FormatDateTime("dd.mm.yyyy",d);
            //  EditDATN1->Font->Color = clBlack;
            }

        }
    }
      StringGrid1->Invalidate();
}
//---------------------------------------------------------------------------


void __fastcall TZameshenie::StringGrid1KeyPress(TObject *Sender,
      char &Key)
{
  if (!(IsNumeric(Key) || Key=='\b' || Key=='.' || Key==',' || Key=='/')) Key=0;
  if (Key==',' || Key=='/') Key='.';
}
//---------------------------------------------------------------------------

void __fastcall TZameshenie::StringGrid1KeyDown(TObject *Sender, WORD &Key,
      TShiftState Shift)
{

  //При нажатии Enter
  if (Key==VK_RETURN)
    {
      if (StringGrid1->Col==0)
        {
          StringGrid1->Col=1;
        }
      else if (StringGrid1->Col==1 && StringGrid1->Row < kol_str1)
        {
          StringGrid1->Col=0;
          StringGrid1->Row = StringGrid1->Row +1;
        }
      else
        {
          //if (StringGrid2->Cells[0][1]=="") BitBtn1->SetFocus();
          FindNextControl((TWinControl *)Sender, true, true, false)->SetFocus();
        }
    }


    /*
      if (StringGrid1->Row < Main->kol_str1)
        {
          StringGrid1->Row = StringGrid1->Row +1;

        }
      else if (StringGrid1->Col < StringGrid1->ColCount-1)
        {
          StringGrid1->Col = StringGrid1->Col+1;
          StringGrid1->Row = 1;
        }
      else
        {
          FindNextControl((TWinControl *)Sender, true, true, false)->SetFocus();
        }
    }   */



 /* //ЛКМ, ПКМ, Пробел, Стрелка влево, Стрелка вправо, Стрелка вверх, Стрелка вниз, Enter                                              Стрелка вправо      Стрелка вверх   Стрелка вниз
if (Key==VK_LBUTTON || Key==VK_RBUTTON || Key==VK_SPACE || Key==VK_LEFT ||
     Key==VK_RIGHT || Key==VK_UP || Key==VK_DOWN)// || Key==VK_RETURN)

   {

     t_r = StringGrid1->Row;
     t_k = StringGrid1->Col;

     // Проверка на правильность ввода даты
          if(!TryStrToDate(StringGrid1->Cells[StringGrid1->Col][StringGrid1->Row],d))
            {
              StringGrid1->Row=StringGrid1->Row-1;
              StringGrid1->Col=StringGrid1->Col;

              StringGrid1->Options << goEditing;
            //  CanSelect = true;
            }


   }



 /*  //При нажатии Enter
  if (Key==VK_RETURN)
    {




 if (ActiveControl == BitBtn2)
    {
      Zameshenie->Close();
    }
  else
    {
      int i= StringGrid1->Col;
      int j= StringGrid1->Row;
      if (StringGrid1->Cells[i][j]!="")
        {
          // Добавление к дате отчетного месяца и года
          if (StringGrid1->Cells[i][j].Length()<3)
            {
              if(StringGrid1->Cells[i][j].Pos("."))
                {
                  Application->MessageBox("Неверный формат даты","Ошибка", MB_OK+MB_ICONINFORMATION);

                  StringGrid1->Row=j;
                  StringGrid1->Col=i;
                  StringGrid1->Options << goEditing;
                  //CanSelect = true;
                  Abort();
                }
              else
                {
                  StringGrid1->Cells[i][j] = StringGrid1->Cells[i][j]+ "."+ DateToStr(Date()).SubString(4,2) +"."+ DateToStr(Date()).SubString(7,5);
                  //EditDATN1->Font->Color = clBlack;
                }
            }

          // Проверка на правильность ввода даты
          if(!TryStrToDate(StringGrid1->Cells[i][j],d))
            {
              Application->MessageBox("Неверный формат даты","Ошибка", MB_OK);
              //CanSelect = true;
              StringGrid1->Row=j;
              StringGrid1->Col=i;

              StringGrid1->Options << goEditing;
            TGridRect g;

 g.Bottom=StringGrid1->Row;
 g.Top=StringGrid1->Row;
 g.Left=StringGrid1->Col;
 g.Right=StringGrid1->Col;
 StringGrid1->Selection=g;

             // StringGrid1->TopRow = j;
             // StringGrid1->Selection = TGridRect(Rect(i, j, i, j));

              
            }
          else
            {
              StringGrid1->Cells[i][j]=FormatDateTime("dd.mm.yyyy",d);
            }

        }
    }

     










      if (StringGrid1->Row < Main->kol_str1)
        {
          StringGrid1->Row = StringGrid1->Row +1;

        }
      else if (StringGrid1->Col < StringGrid1->ColCount-1)
        {
          StringGrid1->Col = StringGrid1->Col+1;
          StringGrid1->Row = 1;
        }
      else
        {
          FindNextControl((TWinControl *)Sender, true, true, false)->SetFocus();
        }
    }   */
}
//---------------------------------------------------------------------------

void __fastcall TZameshenie::CheckBoxREZERVClick(TObject *Sender)
{
  int rec;

  if (CheckBoxREZERV->Checked==true)
    {
      CheckBoxZAM->Enabled = false;
      CheckBoxREZERV->Enabled = true;
      DBGridEh1->Enabled = true;

      //Вывод информации по замещениям в зависимости от фильтров
      DM->qZamesh->Filtered = false;
      DM->qZamesh->Filter = " type=1 and tn="+DM->qOcenka->FieldByName("tn")->AsString;
      if (DM->qZamesh->Active==false) DM->qZamesh->Active=true;
      DM->qZamesh->Filtered = true;

    }
  else
    {
      if (!EditTN->Text.IsEmpty() && EditTN->Text!="490")
        {
          //Вывод информации по замещениям в зависимости от фильтров
          DM->qZamesh->Filtered = false;
          DM->qZamesh->Filter = " type=1 and tn="+DM->qOcenka->FieldByName("tn")->AsString;
          if (DM->qZamesh->Active==false) DM->qZamesh->Active=true;
          DM->qZamesh->Filtered = true;

          if (DM->qZamesh->RecordCount>0)
            {
              if(Application->MessageBox("По данному работнику существуют должности, которые он замещает. Удалить все должности на замещение?","Предупреждение",
                                      MB_YESNO+MB_ICONWARNING)==ID_NO)
                {
                  CheckBoxREZERV->Checked=true;
                  Abort();
                }
              else
                {
                  rec=DM->qOcenka->RecNo;
                  //Удаление всех замещаемых должностей
                  DM->qZamesh->First();
                  while (!DM->qZamesh->Eof)
                    {
                      DM->qObnovlenie->Close();
                      DM->qObnovlenie->SQL->Clear();
                      DM->qObnovlenie->SQL->Add("delete from ocenka_rez where rowid = chartorowid("+ QuotedStr(DM->qZamesh->FieldByName("rw")->AsString)+")");
                      try
                        {
                          DM->qObnovlenie->ExecSQL();
                        }
                      catch (Exception &E)
                        {
                          Application->MessageBox(("Возникла ошибка при удалении замещаемой должности в таблице OCENKA_REZ "+E.Message).c_str(),"Ошибка",
                                                   MB_OK+MB_ICONERROR);
                          DM->qZamesh->Requery();
                          //Логи
                          Main->InsertLog("Возникла ошибка при удалении резервной должности("+DM->qZamesh->FieldByName("id_shtat")->AsString+") по работнику: таб.№='"+EditTN->Text+"' ФИО='"+LabelFIO->Caption+"'");
                          Abort();
                        }

                      DM->qZamesh->Next();
                    }
                 if (DM->qObnovlenie->RowsAffected>0)
                   {
                     //Обновление признака в таблице замещения
                     DM->qObnovlenie->Close();
                      DM->qObnovlenie->SQL->Clear();
                      DM->qObnovlenie->SQL->Add("update ocenka set rezerv=NULL where god="+IntToStr(Main->god)+" and tn="+EditTN->Text);
                      try
                        {
                          DM->qObnovlenie->ExecSQL();
                        }
                      catch (Exception &E)
                        {
                          Application->MessageBox(("Возникла ошибка при обновлении признака резервиста в таблице OCENKA "+E.Message).c_str(),"Ошибка",
                                                   MB_OK+MB_ICONERROR);
                          DM->qZamesh->Requery();
                          //Логи
                          Main->InsertLog("Возникла ошибка при обновлении признака резервиста по работнику: таб.№='"+EditTN->Text+"' ФИО='"+LabelFIO->Caption+"'");
                          Abort();
                        }
                   }
                  DM->qOcenka->Requery();
                  DM->qOcenka->RecNo=rec;

                  //Логи
                  Main->InsertLog("Удаление замещаемых должностей выполнено успешно по работнику: таб.№='"+EditTN->Text+"' ФИО='"+LabelFIO->Caption+"'");
                }
            }
        }

      CheckBoxZAM->Enabled = true;
      DM->qZamesh->Filtered = false;
      DM->qZamesh->Active=false;
      DBGridEh1->Enabled = false;
    }
}
//---------------------------------------------------------------------------

void __fastcall TZameshenie::FormClose(TObject *Sender,
      TCloseAction &Action)
{
  DM->qZamesh->Filtered = false;
}
//---------------------------------------------------------------------------

//Добавление должности для замещения (резервисту - резервную, замещающему - замещаемую)
void __fastcall TZameshenie::N1DobavClick(TObject *Sender)
{
  fl_red=0;
  Panel3->Visible = true;
  Panel1->Align=alClient;
  Zameshenie->Height=705;
  BitBtn1->Top=579;
  BitBtn2->Top=618;
  Bevel3->Height=651;
  if (DM->qOcenka->FieldByName("rezerv")->AsString=="1") GroupBoxZAM->Caption="Добавление резервной должности";
  else GroupBoxZAM->Caption="Добавление замещения";
  BitBtn3->Caption = "Добавить";


  //Очистка эдитов
  //Цех
  EditZEX_ZAM->Text="";
  LabelZEX_ZAM->Caption ="";

  //Шифр
  EditSHIFR_ZAM->Text="";
  LabelDOLG_ZAM->Caption="";

  //ТН замещаемого
  LabelTN_R->Caption="";

  //КРД должность
  LabelKRD->Caption="";

  //ФИО замещаемого
   EditFIO_R->Text="";

  //Риск и причина
  ComboBoxRISK->ItemIndex=-1;
  ComboBoxRISK->Color=clWindow;
  EditRISK_PRICH->Text="";


  //Готовность
  ComboBoxGOTOV->ItemIndex=-1;
  ComboBoxGOTOV->Color=clWindow;


  //Периоды замещения
  //дата начала замещения
  StringGrid1->Cells[0][1]="";
  StringGrid1->Cells[0][2]="";
  StringGrid1->Cells[0][3]="";
  StringGrid1->Cells[0][4]="";
  StringGrid1->Cells[0][5]="";
  StringGrid1->Cells[0][6]="";
  //дата конца замещения
  StringGrid1->Cells[1][1]="";
  StringGrid1->Cells[1][2]="";
  StringGrid1->Cells[1][3]="";
  StringGrid1->Cells[1][4]="";
  StringGrid1->Cells[1][5]="";
  StringGrid1->Cells[1][6]="";
}
//---------------------------------------------------------------------------

//Редактирование должности
void __fastcall TZameshenie::N2RedakClick(TObject *Sender)
{
  if (DM->qZamesh->RecordCount==0)
    {
      fl_red=0;
      N1DobavClick(Sender);
    }
  else
    {
      fl_red=1;

      Panel1->Align=alClient;
      Zameshenie->Height=705;
      BitBtn1->Top=579;
      BitBtn2->Top=618;
      Bevel3->Height=651;
      GroupBoxZAM->Caption="Редактирование замещения";
      if (DM->qOcenka->FieldByName("rezerv")->AsString=="1") GroupBoxZAM->Caption="Редактирование резервной должности";
      else GroupBoxZAM->Caption="Редактирование замещения";
      BitBtn3->Caption = "Редактировать";

      ZapolnenieInfo();

      Panel3->Visible = true;
   }
 /* if (DM->qOcenka->FieldByName("preem")->AsString=="1")
    {
      CheckBoxPREEM->Checked=true;
      zpreem=1;
    }
  else
    {
      CheckBoxPREEM->Checked=false;
      zpreem="NULL";
    } */

}
//---------------------------------------------------------------------------

//Заполнение Эдитов
void __fastcall TZameshenie::ZapolnenieInfo()
{
  //Цех
  EditZEX_ZAM->Text=zzex_zam=DM->qZamesh->FieldByName("zex_rez")->AsString;
  //ВЫвод цеха и участка
  if (!EditZEX_ZAM->Text.IsEmpty())
    {
      DM->qDolg->Filtered = false;
      TLocateOptions SearchOptions;
      DM->qDolg->Locate("shifr_zex", EditZEX_ZAM->Text,
                            SearchOptions << loCaseInsensitive);

      if (EditZEX_ZAM->Text.Length()>2) LabelZEX_ZAM->Caption = DM->qDolg->FieldByName("nzex")->AsString +"\n"+ DM->qDolg->FieldByName("uch")->AsString;
      else LabelZEX_ZAM->Caption = DM->qDolg->FieldByName("nzex")->AsString;
    }
  else LabelZEX_ZAM->Caption ="";


  //Шифр
  EditSHIFR_ZAM->Text=zshifr_zam=DM->qZamesh->FieldByName("id_shtat")->AsString;


  //ТН замещаемого
  LabelTN_R->Caption=ztn_r=DM->qZamesh->FieldByName("tn_sap_rez")->AsString;


  //КРД должность
  //Проверка является ли должность КРД
  DM->qObnovlenie->Close();
  DM->qObnovlenie->SQL->Clear();
  DM->qObnovlenie->SQL->Add("select * from sp_ocenka_krd \
                             where zex="+QuotedStr(EditZEX_ZAM->Text)+"\
                             and shifr_dolg=(select short from p1000@sapmig_buffdb where otype='S' and langu='R' and objid="+QuotedStr(EditSHIFR_ZAM->Text)+")");
  try
    {
      DM->qObnovlenie->Open();
    }
  catch(Exception &E)
    {
      Application->MessageBox(("Невозможно получить данные из справочника КРД (SP_OCENKA_KRD)"+E.Message).c_str(),"Ошибка",
                               MB_OK+MB_ICONERROR);
    }

  if (DM->qObnovlenie->RecordCount>0) LabelKRD->Caption="Ключевая Резервная Должность";
  else LabelKRD->Caption="";


  //ФИО замещаемого
   EditFIO_R->Text=zfio_r=DM->qZamesh->FieldByName("fio_rez")->AsString;


  //Риск и причина
  if (DM->qZamesh->FieldByName("risk")->AsString==1)
    {
      ComboBoxRISK->ItemIndex=ComboBoxRISK->Items->IndexOf("низкий");
      ComboBoxRISK->Color=clMoneyGreen;
      zrisk=1;
    }
  else if (DM->qZamesh->FieldByName("risk")->AsString==2)
    {
      ComboBoxRISK->ItemIndex=ComboBoxRISK->Items->IndexOf("средний");
      ComboBoxRISK->Color=(TColor)0x0080FFFF;
      zrisk=2;
    }
  else if (DM->qZamesh->FieldByName("risk")->AsString==3)
    {
      ComboBoxRISK->ItemIndex=ComboBoxRISK->Items->IndexOf("высокий");
      ComboBoxRISK->Color=(TColor)0x008080FF;
      zrisk=3;
    }
  else
    {
      ComboBoxRISK->ItemIndex=-1;
      ComboBoxRISK->Color=clWindow;
      zrisk="NULL";
    }
  EditRISK_PRICH->Text=zrisk_prich=DM->qZamesh->FieldByName("risk_prich")->AsString;


  //Готовность
  if (DM->qZamesh->FieldByName("gotov")->AsString==1)
    {
      ComboBoxGOTOV->ItemIndex=ComboBoxGOTOV->Items->IndexOf("низкая");
      ComboBoxGOTOV->Color=(TColor)0x008080FF;
      zgotov=1;
    }
  else if (DM->qZamesh->FieldByName("gotov")->AsString==2)
    {
      ComboBoxGOTOV->ItemIndex=ComboBoxGOTOV->Items->IndexOf("средняя");
      ComboBoxGOTOV->Color=(TColor)0x0080FFFF;
      zgotov=2;
    }
  else if (DM->qZamesh->FieldByName("gotov")->AsString==3)
    {
      ComboBoxGOTOV->ItemIndex=ComboBoxGOTOV->Items->IndexOf("высокая");
      ComboBoxGOTOV->Color=clMoneyGreen;
      zgotov=3;
    }
  else
    {
      ComboBoxGOTOV->ItemIndex=-1;
      ComboBoxGOTOV->Color=clWindow;
      zgotov="NULL";
    }


  //Периоды замещения
  //дата начала замещения
  StringGrid1->Cells[0][1]=zdatn1=DM->qZamesh->FieldByName("datn1")->AsString;
  StringGrid1->Cells[0][2]=zdatn2=DM->qZamesh->FieldByName("datn2")->AsString;
  StringGrid1->Cells[0][3]=zdatn3=DM->qZamesh->FieldByName("datn3")->AsString;
  StringGrid1->Cells[0][4]=zdatn4=DM->qZamesh->FieldByName("datn4")->AsString;
  StringGrid1->Cells[0][5]=zdatn5=DM->qZamesh->FieldByName("datn5")->AsString;
  StringGrid1->Cells[0][6]=zdatn6=DM->qZamesh->FieldByName("datn6")->AsString;
  //дата конца замещения
  StringGrid1->Cells[1][1]=zdatk1=DM->qZamesh->FieldByName("datk1")->AsString;
  StringGrid1->Cells[1][2]=zdatk2=DM->qZamesh->FieldByName("datk2")->AsString;
  StringGrid1->Cells[1][3]=zdatk3=DM->qZamesh->FieldByName("datk3")->AsString;
  StringGrid1->Cells[1][4]=zdatk4=DM->qZamesh->FieldByName("datk4")->AsString;
  StringGrid1->Cells[1][5]=zdatk5=DM->qZamesh->FieldByName("datk5")->AsString;
  StringGrid1->Cells[1][6]=zdatk6=DM->qZamesh->FieldByName("datk6")->AsString;

  kol_str1=0;
  //Количество заполненных строк в StringGrid1
  for (int i=1; i<7; i++)
    {
      if (!DM->qZamesh->FieldByName("datn"+IntToStr(i))->AsString.IsEmpty()) kol_str1++;
    }
}
//---------------------------------------------------------------------------
void __fastcall TZameshenie::BitBtn4Click(TObject *Sender)
{
  //Скрыть панель редактирования
  Panel3->Visible = false;
  Panel1->Align=alClient;
  Zameshenie->Height=440;
  BitBtn1->Top=316;
  BitBtn2->Top=355;
  Bevel3->Height=385;
}
//---------------------------------------------------------------------------

//Добавление/редактирование должности
void __fastcall TZameshenie::BitBtn3Click(TObject *Sender)
{
  AnsiString risk, Sql, Str, zam,  gotov, rez;
  int type=0, rec1, rec2;

  //Риск
  if (ComboBoxRISK->Text=="низкий") risk=1;
  else if (ComboBoxRISK->Text=="средний") risk=2;
  else if (ComboBoxRISK->Text=="высокий") risk=3;
  else risk="NULL";

    //Готовность
  if (ComboBoxGOTOV->Text=="низкая") gotov=1;
  else if (ComboBoxGOTOV->Text=="средняя") gotov=2;
  else if (ComboBoxGOTOV->Text=="высокая") gotov=3;
  else gotov="NULL";

  //Резервист/замещающий
  if (CheckBoxREZERV->Checked==true)
    {
      type=1;
      zam="NULL";
      rez=1;
    }
  else if (CheckBoxZAM->Checked==true)
    {
      type=2;
      zam=1;
      rez="NULL";
    }
  else
    {
      type=0;
      zam="NULL";
      rez="NULL";
    }

  //Проверки раздела "Замещение"
  //Таб.№
  if (EditTN->Text.IsEmpty())
    {
      Application->MessageBox("Не указан таб.№ работника!!!","Предупреждение",
                               MB_OK+MB_ICONINFORMATION);
      EditTN->SetFocus();
      Abort();
    }

  //Цех
  if (EditZEX_ZAM->Text.IsEmpty())
    {
      Application->MessageBox("Не указан цех замещаемого работника!!!","Предупреждение",
                               MB_OK+MB_ICONINFORMATION);
      EditZEX_ZAM->SetFocus();
      Abort();
    }

  //Шифр
  if (EditSHIFR_ZAM->Text.IsEmpty())
    {
      Application->MessageBox("Не указан шифр замещаемой должности!!!","Предупреждение",
                               MB_OK+MB_ICONINFORMATION);
      EditSHIFR_ZAM->SetFocus();
      Abort();
    }

  //Отсутствие табельного номера сап замещаемого работника при наличии ФИО
  if (!EditFIO_R->Text.IsEmpty() && LabelTN_R->Caption=="")
    {
      Application->MessageBox("Невозможно получить табельный номер замещаемого работника.\nВозможно неверно введено ФИО замещаемого работника \nлибо замещаемый работник уволен.","Предупреждение",
                                   MB_OK+MB_ICONWARNING);
      EditFIO_R->SetFocus();
      Abort();
    }

  //Отсутствие ФИО замещаемого работника при наличии табельного номера сап
  if (EditFIO_R->Text.IsEmpty() && LabelTN_R->Caption!="")
    {
      Application->MessageBox("Не указано ФИО замещаемого работника","Предупреждение",
                                   MB_OK+MB_ICONWARNING);
      EditFIO_R->SetFocus();
      Abort();
    }

  //Существование должности
  if (!EditSHIFR_ZAM->Text.IsEmpty())
    {
      Sql="select stext from p1000@sapmig_buffdb where otype='S' and langu='R' and objid=:pkod_prof";


      DM->qObnovlenie->Close();
      DM->qObnovlenie->SQL->Clear();
      DM->qObnovlenie->SQL->Add(Sql);
      DM->qObnovlenie->Parameters->ParamByName("pkod_prof")->Value = Main->SetNull(EditSHIFR_ZAM->Text);
      try
        {
          DM->qObnovlenie->Open();
        }
      catch(Exception &E)
        {
          Application->MessageBox(("Возникла ошибка при попытке доступа к справочнику должностей (P1000)"+ E.Message).c_str(),"Ошибка",
                                  MB_OK+MB_ICONERROR);
        }

      if (DM->qObnovlenie->RecordCount==0)
        {
          Application->MessageBox("Нет указанного шифра должности в справочнике должностей!!!","Предупреждение",
                                   MB_OK+MB_ICONINFORMATION);
          EditSHIFR_ZAM->SetFocus();
          Abort();
        }
    }

  //Существование замещаемого работника на такой должности
  if (!EditFIO_R->Text.IsEmpty())
    {
      Sql="select tn_sap from \
                             (select case when ur1 is null then zex          \
                                          when ur2 is null then ur1          \
                                          when ur3 is null then ur2          \
                                          when ur4 is null then ur3 end ur,  \
                                     tn_sap,                                 \
                                     id_shtat                                 \
                              from sap_osn_sved)                             \
           where id_shtat=:pid_dolg and ur=:pur and tn_sap=:ptn_sap";


      DM->qObnovlenie->Close();
      DM->qObnovlenie->SQL->Clear();
      DM->qObnovlenie->SQL->Add(Sql);
      DM->qObnovlenie->Parameters->ParamByName("pid_dolg")->Value = Main->SetNull(EditSHIFR_ZAM->Text);
      DM->qObnovlenie->Parameters->ParamByName("pur")->Value = Main->SetNull(EditZEX_ZAM->Text);
      DM->qObnovlenie->Parameters->ParamByName("ptn_sap")->Value = Main->SetNull(LabelTN_R->Caption);
      try
        {
          DM->qObnovlenie->Open();
        }
      catch(Exception &E)
        {
          Application->MessageBox(("Возникла ошибка при попытке доступа к картотеке работников (SAP_OSN_SVED)"+ E.Message).c_str(),"Ошибка",
                                  MB_OK+MB_ICONERROR);
        }

      if (DM->qObnovlenie->RecordCount==0)
        {
          Application->MessageBox("Несоответствие шифра замещаемой должности и замещаемого работника!!!\nВозможно неверно указан шифр подразделения, \nшифр должности или ФИО замещаемого работника","Предупреждение",
                                   MB_OK+MB_ICONINFORMATION);
          EditZEX_ZAM->SetFocus();
          Abort();
        }

    }


//Проверки раздела "Периоды замещения"
//***************************************
  //Дата начала не заполнена, а дата конца заполнена
  //StringGrid1
  for (int i=1; i<7; i++)
    {
      if (StringGrid1->Cells[0][i]=="" && StringGrid1->Cells[1][i]!="")
        {
          Application->MessageBox("Указана дата конца замещения, но не указана дата начала замещения!!!","Предупреждение",
                              MB_OK+MB_ICONINFORMATION);
          StringGrid1->SetFocus();
          StringGrid1->Row=i;
          StringGrid1->Col=0;
          StringGrid1->EditorMode = true;

          Abort();
        }
    }

  //Дата начала заполнена, а дата конца не заполнена
  //StringGrid1
  for (int i=1; i<7; i++)
    {
      if (StringGrid1->Cells[0][i]!="" && StringGrid1->Cells[1][i]=="")
        {
          Application->MessageBox("Указана дата начала замещения, но не указана дата конца замещения!!!","Предупреждение",
                              MB_OK+MB_ICONINFORMATION);

          StringGrid1->SetFocus();
          StringGrid1->Row=i;
          StringGrid1->Col=1;
          StringGrid1->EditorMode = true;

          Abort();
        }
    }

  //Дата конца больше даты начала
  //StringGrid1
  for (int i=1; i<7; i++)
    {
      if (StringGrid1->Cells[0][i]!="" && StrToDate(StringGrid1->Cells[0][i])>StrToDate(StringGrid1->Cells[1][i]))
        {
          Application->MessageBox("Дата начала замещения больше, чем дата конца замещения!!!","Предупреждение",
                                   MB_OK+MB_ICONINFORMATION);
          StringGrid1->SetFocus();
          StringGrid1->Row=i;
          StringGrid1->Col=0;
          StringGrid1->EditorMode = true;

          Abort();
        }
    }

  // Проверка на совпадение сроков неявки
  //Даты в StringGrid1
  for (int j=1; j<7; j++)
    {
      if (StringGrid1->Cells[0][j]!="")
        {
          for (int i=1; i<7; i++)
            {
              //Проверка пересекающихся дат из StringGrid1 в StringGrid1
              if (i!=j && StringGrid1->Cells[0][i]!="")
                {
                  if (((StrToDate(StringGrid1->Cells[0][i]) < StrToDate(StringGrid1->Cells[0][j]) && StrToDate(StringGrid1->Cells[1][i]) > StrToDate(StringGrid1->Cells[0][j]))
                        || (StrToDate(StringGrid1->Cells[0][i]) > StrToDate(StringGrid1->Cells[0][j]) && StrToDate(StringGrid1->Cells[1][i]) > StrToDate(StringGrid1->Cells[0][j])
                            && (StrToDate(StringGrid1->Cells[0][i]) < StrToDate(StringGrid1->Cells[1][j]) || StrToDate(StringGrid1->Cells[0][i]) == StrToDate(StringGrid1->Cells[1][j])))
                        ||  (StrToDate(StringGrid1->Cells[0][i]) == StrToDate(StringGrid1->Cells[0][j]) || StrToDate(StringGrid1->Cells[1][i]) == StrToDate(StringGrid1->Cells[0][j])))
                      )
                    {
                      Application->MessageBox("Вводимый период замещения пересекается\nс уже существующим","Ошибка",
                                               MB_OK + MB_ICONERROR);

                      StringGrid1->SetFocus();
                      StringGrid1->Row=j;
                      StringGrid1->Col=0;
                      StringGrid1->EditorMode = true;

                      Abort();
                    }
                }

              //Проверка пересекающихся дат из StringGrid1 в StringGrid2
            /*  if (StringGrid2->Cells[0][i]!="")
                {
                    if (((StrToDate(StringGrid2->Cells[0][i]) < StrToDate(StringGrid1->Cells[0][j]) && StrToDate(StringGrid2->Cells[1][i]) > StrToDate(StringGrid1->Cells[0][j]))
                        || (StrToDate(StringGrid2->Cells[0][i]) > StrToDate(StringGrid1->Cells[0][j]) && StrToDate(StringGrid2->Cells[1][i]) > StrToDate(StringGrid1->Cells[0][j])
                            && (StrToDate(StringGrid2->Cells[0][i]) < StrToDate(StringGrid1->Cells[1][j]) || StrToDate(StringGrid2->Cells[0][i]) == StrToDate(StringGrid1->Cells[1][j])))
                        ||  (StrToDate(StringGrid2->Cells[0][i]) == StrToDate(StringGrid1->Cells[0][j]) || StrToDate(StringGrid2->Cells[1][i]) == StrToDate(StringGrid1->Cells[0][j])))
                      )
                    {
                      Application->MessageBox("Вводимый период замещения пересекается\nс уже существующим","Ошибка",
                                               MB_OK + MB_ICONERROR);

                      StringGrid1->SetFocus();
                      StringGrid1->Row=j;
                      StringGrid1->Col=0;
                      StringGrid1->EditorMode = true;

                      Abort();
                    }
                }  */
            }
        }
    }

  //Вставка
  if (fl_red==0)
    {
      //Проверка на уже наличие такой вакансии у данного работника
      DM->qObnovlenie->Close();
      DM->qObnovlenie->SQL->Clear();
      DM->qObnovlenie->SQL->Add("select * from ocenka_rez where god="+IntToStr(Main->god)+" and tn="+EditTN->Text+" and id_shtat="+QuotedStr(EditSHIFR_ZAM->Text));
      try
        {
          DM->qObnovlenie->Open();
        }
      catch(Exception &E)
        {
          Application->MessageBox(("Возникла ошибка при получении данных из таблицы Ocenka_rez" + E.Message).c_str(),"Ошибка",
                                  MB_OK+MB_ICONERROR);
          Abort();
        }

      if (DM->qObnovlenie->RecordCount>0)
        {
          Application->MessageBox("У данного работника уже есть замещение на указанную должность!!!","Предупреждение",
                                   MB_OK+MB_ICONINFORMATION);
          EditZEX->SetFocus();
          Abort();
        }


      Sql="insert into ocenka_rez (god, tn, id_shtat, dolg_rez, tn_sap_rez, fio_rez, zex_rez, shifr_rez, type, risk, risk_prich, gotov, \
                                   datn1, datn2, datn3, datn4, datn5, datn6, datk1, datk2, datk3, datk4, datk5, datk6) \
           values ("+IntToStr(Main->god)+",                                                                                         \
                    "+EditTN->Text+",                                                                                                    \
                    "+QuotedStr(EditSHIFR_ZAM->Text)+",                                                                                             \
                    "+QuotedStr(LabelDOLG_ZAM->Caption)+",                                                                                          \
                    "+Main->SetNull(LabelTN_R->Caption)+",                                                                                              \
                    "+QuotedStr(EditFIO_R->Text)+",                                                                                                 \
                    "+QuotedStr(EditZEX_ZAM->Text)+",                                                                                               \
                    "+QuotedStr(id_dolg)+",  \
                    "+type+",                                                                                                        \
                    "+risk+",                                                                                                        \
                    "+QuotedStr(EditRISK_PRICH->Text)+",                                                                                         \
                    "+gotov+",                                           \
                     to_date("+QuotedStr(StringGrid1->Cells[0][1])+",'dd.mm.yyyy'),\
                     to_date("+QuotedStr(StringGrid1->Cells[0][2])+",'dd.mm.yyyy'),\
                     to_date("+QuotedStr(StringGrid1->Cells[0][3])+",'dd.mm.yyyy'),\
                     to_date("+QuotedStr(StringGrid1->Cells[0][4])+",'dd.mm.yyyy'),\
                     to_date("+QuotedStr(StringGrid1->Cells[0][5])+",'dd.mm.yyyy'),\
                     to_date("+QuotedStr(StringGrid1->Cells[0][6])+",'dd.mm.yyyy'),\
                     to_date("+QuotedStr(StringGrid1->Cells[1][1])+",'dd.mm.yyyy'),\
                     to_date("+QuotedStr(StringGrid1->Cells[1][2])+",'dd.mm.yyyy'),\
                     to_date("+QuotedStr(StringGrid1->Cells[1][3])+",'dd.mm.yyyy'),\
                     to_date("+QuotedStr(StringGrid1->Cells[1][4])+",'dd.mm.yyyy'),\
                     to_date("+QuotedStr(StringGrid1->Cells[1][5])+",'dd.mm.yyyy'),\
                     to_date("+QuotedStr(StringGrid1->Cells[1][6])+",'dd.mm.yyyy')\
                  )";
      rec1=DM->qOcenka->RecNo;
      rec2=DM->qZamesh->RecNo;

    }
  //Обновление
  else if (fl_red==1)
    {
      //Проверка на наличие обновлений
     if (EditTN->Text!=ztn ||
         EditZEX_ZAM->Text!=zzex_zam ||
         EditSHIFR_ZAM->Text!=zshifr_zam ||
         EditFIO_R->Text!=zfio_r ||
         LabelTN_R->Caption!=ztn_r ||
         DM->qOcenka->FieldByName("rezerv")->AsString!=rez ||
         DM->qOcenka->FieldByName("zam")->AsString!=zam ||
         gotov!=zgotov ||
         risk!=zrisk ||
         EditRISK_PRICH->Text!=zrisk_prich ||
         StringGrid1->Cells[0][1]!=zdatn1 ||
         StringGrid1->Cells[0][2]!=zdatn2 ||
         StringGrid1->Cells[0][3]!=zdatn3 ||
         StringGrid1->Cells[0][4]!=zdatn4 ||
         StringGrid1->Cells[0][5]!=zdatn5 ||
         StringGrid1->Cells[0][6]!=zdatn6 ||
         StringGrid1->Cells[1][1]!=zdatk1 ||
         StringGrid1->Cells[1][2]!=zdatk2 ||
         StringGrid1->Cells[1][3]!=zdatk3 ||
         StringGrid1->Cells[1][4]!=zdatk4 ||
         StringGrid1->Cells[1][5]!=zdatk5 ||
         StringGrid1->Cells[1][6]!=zdatk6 )
       {
         Sql="update ocenka_rez set  \                                                                                                                  \
                             dolg_rez="+QuotedStr(LabelDOLG_ZAM->Caption)+",                                                                             \
                             tn_sap_rez="+Main->SetNull(LabelTN_R->Caption)+",\                                                                         \
                             fio_rez="+QuotedStr(EditFIO_R->Text)+",\                                                                                   \
                             zex_rez="+QuotedStr(EditZEX_ZAM->Text)+",\                                                                                 \
                             shifr_rez="+QuotedStr(id_dolg)+",  \
                             type="+type+",                                                                                                             \
                             risk="+risk+",                                                                                                             \
                             risk_prich="+QuotedStr(EditRISK_PRICH->Text)+",\
                             gotov="+gotov+",                               \
                             datn1=to_date("+QuotedStr(StringGrid1->Cells[0][1])+",'dd.mm.yyyy'),\
                             datn2=to_date("+QuotedStr(StringGrid1->Cells[0][2])+",'dd.mm.yyyy'),\
                             datn3=to_date("+QuotedStr(StringGrid1->Cells[0][3])+",'dd.mm.yyyy'),\
                             datn4=to_date("+QuotedStr(StringGrid1->Cells[0][4])+",'dd.mm.yyyy'),\
                             datn5=to_date("+QuotedStr(StringGrid1->Cells[0][5])+",'dd.mm.yyyy'),\
                             datn6=to_date("+QuotedStr(StringGrid1->Cells[0][6])+",'dd.mm.yyyy'),\
                             datk1=to_date("+QuotedStr(StringGrid1->Cells[1][1])+",'dd.mm.yyyy'),\
                             datk2=to_date("+QuotedStr(StringGrid1->Cells[1][2])+",'dd.mm.yyyy'),\
                             datk3=to_date("+QuotedStr(StringGrid1->Cells[1][3])+",'dd.mm.yyyy'),\
                             datk4=to_date("+QuotedStr(StringGrid1->Cells[1][4])+",'dd.mm.yyyy'),\
                             datk5=to_date("+QuotedStr(StringGrid1->Cells[1][5])+",'dd.mm.yyyy'),\
                             datk6=to_date("+QuotedStr(StringGrid1->Cells[1][6])+",'dd.mm.yyyy')\
            where rowid = chartorowid("+ QuotedStr(DM->qZamesh->FieldByName("rw")->AsString)+") and god="+IntToStr(Main->god)+" \
            and tn="+EditTN->Text+" and id_shtat="+QuotedStr(EditSHIFR_ZAM->Text);

          rec1=DM->qOcenka->RecNo;
          rec2=DM->qZamesh->RecNo;
        }
       else
         {
           //Скрытие панели
           Panel3->Visible = false;
           Panel1->Align=alClient;
           Zameshenie->Height=440;
           BitBtn1->Top=316;
           BitBtn2->Top=355;
           Bevel3->Height=385;
           Abort();
         }
    }


  DM->qObnovlenie->Close();
  DM->qObnovlenie->SQL->Clear();
  DM->qObnovlenie->SQL->Add(Sql);

  try
    {
      DM->qObnovlenie->ExecSQL();
    }
  catch(Exception &E)
    {
      Application->MessageBox(("Возникла ошибка при обновлении данных в таблице Ocenka_rez" + E.Message).c_str(),"Ошибка",
                               MB_OK+MB_ICONERROR);
      Main->InsertLog("Возникла ошибка при обновлении зам.должности в таблице OCENKA_REZ по работнику "+LabelTN_R->Caption);
      Abort();
    }


  //Если все обновилось, то обновить риск и причину риска у руководителя в таблице Ocenka
   if (DM->qObnovlenie->RowsAffected>0 && !EditFIO_R->Text.IsEmpty()
       && (risk!=zrisk || EditRISK_PRICH->Text!=zrisk_prich ))
     {
       Sql="update ocenka set                                                \
                          risk="+risk+",                                     \
                          risk_prich="+QuotedStr(EditRISK_PRICH->Text)+"               \                                      \
            where god="+IntToStr(Main->god)+" and tn="+LabelTN_R->Caption;

       DM->qObnovlenie->Close();
       DM->qObnovlenie->SQL->Clear();
       DM->qObnovlenie->SQL->Add(Sql);
       try
         {
           DM->qObnovlenie->ExecSQL();
         }
       catch(Exception &E)
         {
           Application->MessageBox(("Возникла ошибка при обновлении данных в таблице Ocenka" + E.Message).c_str(),"Ошибка",
                                   MB_OK+MB_ICONERROR);
           Main->InsertLog("Возникла ошибка при обновлении риска и причины риска во время редактрования зам.должности в таблице OCENKA по работнику "+LabelTN_R->Caption);
           Abort();
         }
     }

   //Если все обновилось, то обновить поле резервист/замещающий у преемника в таблице Ocenka
   if ((DM->qOcenka->FieldByName("rezerv")->AsString!=rez || DM->qOcenka->FieldByName("zam")->AsString!=zam)
        && DM->qObnovlenie->RowsAffected>0)
     {
       Sql = "update ocenka set \
                               rezerv="+rez+", \
                               zam="+zam+"     \
              where tn="+EditTN->Text+" and god="+IntToStr(Main->god);

       DM->qObnovlenie->Close();
       DM->qObnovlenie->SQL->Clear();
       DM->qObnovlenie->SQL->Add(Sql);
       try
         {
           DM->qObnovlenie->ExecSQL();
         }
       catch(Exception &E)
         {
           Application->MessageBox(("Возникла ошибка при попытке обновить данные в таблице Ocenka" + E.Message).c_str(),"Ошибка",
                                    MB_OK+MB_ICONERROR);

           Main->InsertLog("Возникла ошибка при обновлении статуса замещения во время редактрования зам.должности в таблице OCENKA по работнику "+EditTN->Text);
           DM->qLogs->Requery();
           DM->qOcenka->Requery();
           Main->StatusBar1->SimpleText ="Отчетный период: "+IntToStr(Main->god)+" год";
           Abort();
         }
     }
     

  //Логи
  if (DM->qObnovlenie->RowsAffected>0)
    {

      if (fl_red==0)
        {
          Str ="Добавление замещаемой должности за "+IntToStr(Main->god)+" год по работнику '"+EditTN->Text+"': ";
          Str+="таб.№ замещающего="+EditTN->Text+" ФИО="+LabelFIO->Caption+" цех замещаемой.долж.="+EditZEX_ZAM->Text;
          Str+="шифр="+EditSHIFR_ZAM->Text+", ФИО замещаемого="+EditFIO_R->Text+", таб№="+LabelTN_R->Caption+", риск="+risk+", причина="+EditRISK_PRICH->Text+", готовность="+gotov;
          if (StringGrid1->Cells[0][1]!="") Str+=", замещ. c"+StringGrid1->Cells[0][1]+" по "+StringGrid1->Cells[1][1];
          if (StringGrid1->Cells[0][2]!="") Str+=", замещ. c"+StringGrid1->Cells[0][2]+" по "+StringGrid1->Cells[1][2];
          if (StringGrid1->Cells[0][3]!="") Str+=", замещ. c"+StringGrid1->Cells[0][3]+" по "+StringGrid1->Cells[1][3];
          if (StringGrid1->Cells[0][4]!="") Str+=", замещ. c"+StringGrid1->Cells[0][4]+" по "+StringGrid1->Cells[1][4];
          if (StringGrid1->Cells[0][5]!="") Str+=", замещ. c"+StringGrid1->Cells[0][5]+" по "+StringGrid1->Cells[1][5];
          if (StringGrid1->Cells[0][6]!="") Str+=", замещ. c"+StringGrid1->Cells[0][6]+" по "+StringGrid1->Cells[1][6];
        }
      else if (fl_red==1)
        {
          Str ="Редактирование замещаемой должности  за "+IntToStr(Main->god)+" год по работнику '"+EditTN->Text+"': ";
          if (Main->SetNull(EditZEX_ZAM->Text)!=Main->SetNull(zzex_zam)) Str+=" Цех зам.раб. с '"+QuotedStr(zzex_zam)+"' на '"+QuotedStr(EditZEX_ZAM->Text)+"',";
          if (Main->SetNull(EditSHIFR_ZAM->Text)!=Main->SetNull(zshifr_zam)) Str+=" Шифр долг.зам.раб. с '"+QuotedStr(zshifr_zam)+"' на '"+QuotedStr(EditSHIFR_ZAM->Text)+"',";
          if (Main->SetNull(LabelTN_R->Caption)!=Main->SetNull(ztn_r)) Str+=" Таб№ зам.раб. с '"+Main->SetNull(ztn_r)+"' на '"+Main->SetNull(LabelTN_R->Caption)+"',";
          if (Main->SetNull(EditFIO_R->Text)!=Main->SetNull(zfio_r)) Str+=" ФИО зам.раб. с '"+QuotedStr(zfio_r)+"' на '"+QuotedStr(EditFIO_R->Text)+"',";
          if (risk!=zrisk) Str+=" риск с '"+QuotedStr(zrisk)+"' на '"+QuotedStr(risk)+"',";
          if (Main->SetNull(EditRISK_PRICH->Text)!=Main->SetNull(zrisk_prich)) Str+=" причина риска с '"+Main->SetNull(zrisk_prich)+"' на '"+Main->SetNull(EditRISK_PRICH->Text)+"',";
          if (gotov!=zgotov) Str+=" готовность с '"+QuotedStr(zgotov)+"' на '"+QuotedStr(gotov)+"',";
          if (Main->SetNull(StringGrid1->Cells[0][1])!=Main->SetNull(zdatn1)) Str+=" дата нач.зам. с '"+QuotedStr(zdatn1)+"' на '"+QuotedStr(StringGrid1->Cells[0][1])+"',";
          if (Main->SetNull(StringGrid1->Cells[0][2])!=Main->SetNull(zdatn2)) Str+=" дата нач.зам. с '"+QuotedStr(zdatn2)+"' на '"+QuotedStr(StringGrid1->Cells[0][2])+"',";
          if (Main->SetNull(StringGrid1->Cells[0][3])!=Main->SetNull(zdatn3)) Str+=" дата нач.зам. с '"+QuotedStr(zdatn3)+"' на '"+QuotedStr(StringGrid1->Cells[0][3])+"',";
          if (Main->SetNull(StringGrid1->Cells[0][4])!=Main->SetNull(zdatn4)) Str+=" дата нач.зам. с '"+QuotedStr(zdatn4)+"' на '"+QuotedStr(StringGrid1->Cells[0][4])+"',";
          if (Main->SetNull(StringGrid1->Cells[0][5])!=Main->SetNull(zdatn5)) Str+=" дата нач.зам. с '"+QuotedStr(zdatn5)+"' на '"+QuotedStr(StringGrid1->Cells[0][5])+"',";
          if (Main->SetNull(StringGrid1->Cells[0][6])!=Main->SetNull(zdatn6)) Str+=" дата нач.зам. с '"+QuotedStr(zdatn6)+"' на '"+QuotedStr(StringGrid1->Cells[0][6])+"',";
          if (Main->SetNull(StringGrid1->Cells[1][1])!=Main->SetNull(zdatk1)) Str+=" дата ок.зам. с '"+QuotedStr(zdatk1)+"' на '"+QuotedStr(StringGrid1->Cells[1][1])+"',";
          if (Main->SetNull(StringGrid1->Cells[1][2])!=Main->SetNull(zdatk2)) Str+=" дата ок.зам. с '"+QuotedStr(zdatk2)+"' на '"+QuotedStr(StringGrid1->Cells[1][2])+"',";
          if (Main->SetNull(StringGrid1->Cells[1][3])!=Main->SetNull(zdatk3)) Str+=" дата ок.зам. с '"+QuotedStr(zdatk3)+"' на '"+QuotedStr(StringGrid1->Cells[1][3])+"',";
          if (Main->SetNull(StringGrid1->Cells[1][4])!=Main->SetNull(zdatk4)) Str+=" дата ок.зам. с '"+QuotedStr(zdatk4)+"' на '"+QuotedStr(StringGrid1->Cells[1][4])+"',";
          if (Main->SetNull(StringGrid1->Cells[1][5])!=Main->SetNull(zdatk5)) Str+=" дата ок.зам. с '"+QuotedStr(zdatk5)+"' на '"+QuotedStr(StringGrid1->Cells[1][5])+"',";
          if (Main->SetNull(StringGrid1->Cells[1][6])!=Main->SetNull(zdatk6)) Str+=" дата ок.зам. с '"+QuotedStr(zdatk6)+"' на '"+QuotedStr(StringGrid1->Cells[1][6])+"',";
        }

      Main->InsertLog(Str);
      DM->qLogs->Requery();
    }
  else
    {
      Main->InsertLog("Обновление данных за "+IntToStr(Main->god)+" год по работнику: цех="+EditZEX->Text+" таб.№="+EditTN->Text+" не выполнено");
      DM->qLogs->Requery();
    }

  //Обновление таблицы
  DM->qOcenka->Requery();
  DM->qZamesh->Requery();

  //Возвращение на выбранную строку
  if (rec2==-1) rec2=1;
  DM->qZamesh->RecNo=rec2;
  DM->qOcenka->RecNo=rec1;
  
  Application->MessageBox("Запись успешно изменена","Предупреждение",
                           MB_OK+MB_ICONINFORMATION);

  //Скрытие панели
  Panel3->Visible = false;
  Panel1->Align=alClient;
  Zameshenie->Height=440;
  BitBtn1->Top=316;
  BitBtn2->Top=355;
  Bevel3->Height=385;
}
//---------------------------------------------------------------------------

void __fastcall TZameshenie::FormKeyDown(TObject *Sender, WORD &Key,
      TShiftState Shift)
{
  if (Key==VK_RETURN)
  FindNextControl((TWinControl *)Sender, true, true, false)->SetFocus();        
}
//---------------------------------------------------------------------------

void __fastcall TZameshenie::DBGridEh1DblClick(TObject *Sender)
{
  if (DM->qZamesh->RecordCount==0)
    {
      //Добавить запись
      N1DobavClick(Sender);
    }
  else
    {
      //Редактировать запись
      N2RedakClick(Sender);
    }
}
//---------------------------------------------------------------------------


//Удаление должности
void __fastcall TZameshenie::N3DeleteClick(TObject *Sender)
{
  int rec;

  if (Application->MessageBox("Вы действительно хотите безвозвратно удалить \nзамещение по данному работнику?","Предупреждение",
                              MB_YESNO+MB_ICONWARNING)==ID_NO)
    {
      Abort();
    }

  rec=DM->qOcenka->RecNo;
  //Удаление замещения
  DM->qObnovlenie->Close();
  DM->qObnovlenie->SQL->Clear();
  DM->qObnovlenie->SQL->Add("delete from ocenka_rez where god="+IntToStr(Main->god)+" and tn="+EditTN->Text+"  \
                             and id_shtat="+DM->qZamesh->FieldByName("id_shtat")->AsString+" and rowid = chartorowid("+ QuotedStr(DM->qZamesh->FieldByName("rw")->AsString)+")");
  try
    {
      DM->qObnovlenie->ExecSQL();
    }
  catch(Exception &E)
    {
      Application->MessageBox(("Возникла ошибка при получении данных из таблицы Ocenka" + E.Message).c_str(),"Ошибка",
                               MB_OK+MB_ICONERROR);
      Main->InsertLog("Возникла ошибка при удалении замещаемой/резервной должности("+DM->qZamesh->FieldByName("id_shtat")->AsString+") по работнику: таб.№='"+EditTN->Text+"' ФИО='"+LabelFIO->Caption+"'");
      Abort();
    }

  DM->qZamesh->Requery();
  //Удаление признака замещения, если это была единственная запись
  if (DM->qZamesh->RecordCount==0)
    {
      AnsiString zap;
      if (CheckBoxZAM->Checked==true) zap="zam";
      else if (CheckBoxREZERV->Checked==true) zap="rezerv";


      //Обновление признака в таблице замещения
      DM->qObnovlenie->Close();
      DM->qObnovlenie->SQL->Clear();
      DM->qObnovlenie->SQL->Add("update ocenka set "+zap+"=NULL where god="+IntToStr(Main->god)+" and tn="+EditTN->Text);
      try
        {
          DM->qObnovlenie->ExecSQL();
        }
      catch (Exception &E)
        {
          Application->MessageBox(("Возникла ошибка при обновлении признака резервиста в таблице OCENKA "+E.Message).c_str(),"Ошибка",
                                   MB_OK+MB_ICONERROR);
          DM->qZamesh->Requery();
          Main->InsertLog("Возникла ошибка при обновлении признака замещения/резервиста по работнику: таб.№='"+EditTN->Text+"' ФИО='"+LabelFIO->Caption+"'");
          Abort();
        }

       //Обновление таблиц
       DM->qZamesh->Requery();
       DM->qOcenka->Requery();

      if (CheckBoxZAM->Checked==true)
        {
          CheckBoxZAM->Checked=false;
          CheckBoxREZERV->Enabled=true;
        }

    }

  
  //Логи
  Main->InsertLog("Удаление замещаемой/резервной должности выполнено успешно по работнику: таб.№='"+EditTN->Text+"' ФИО='"+LabelFIO->Caption+"'");

  DM->qOcenka->RecNo=rec;

}
//---------------------------------------------------------------------------


