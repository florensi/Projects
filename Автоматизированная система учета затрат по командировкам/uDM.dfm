object DM: TDM
  OldCreateOrder = False
  OnCreate = DataModuleCreate
  OnDestroy = DataModuleDestroy
  Left = 290
  Top = 145
  Height = 709
  Width = 220
  object qObnovlenie: TADOQuery
    Connection = ADOConnection1
    Parameters = <>
    Left = 40
    Top = 96
  end
  object dsKomandirovki: TDataSource
    DataSet = qKomandirovki
    Left = 120
    Top = 160
  end
  object qKomandirovki: TADOQuery
    Connection = ADOConnection1
    CursorType = ctStatic
    Parameters = <>
    SQL.Strings = (
      'select rowidtochar(rowid) rw,'
      
        '           (select naim from sp_komandir where kod=chel) as chel' +
        ','
      
        '           (select country from sp_country where kod=strana) as ' +
        'strana,'
      '           (select city from sp_city where kod=gorod) as gorod,'
      
        '           (select obekt from sp_obekt where kod=k.obekt) as obe' +
        'kt,'
      
        '           (select gostinica from sp_gostinica where kod=k.gosti' +
        'nica) as gostinica,'
      '           chel as kod_chel,'
      '           strana as kod_strana,'
      '           gorod as kod_gorod,'
      '           gostinica as kod_gostinica,'
      '           obekt as kod_obekt,'
      '           k.*,'
      '          regexp_replace(fio, '#39' (.*)'#39') as fam '
      'from komandirovki k')
    Left = 40
    Top = 160
  end
  object dsSP_chel: TDataSource
    DataSet = qSP_chel
    Left = 120
    Top = 272
  end
  object dsSP_grade: TDataSource
    DataSet = qSP_grade
    Left = 120
    Top = 328
  end
  object dsSP_gostinica: TDataSource
    DataSet = qSP_gostinica
    Left = 120
    Top = 384
  end
  object dsSP_obekt: TDataSource
    DataSet = qSP_obekt
    Left = 120
    Top = 440
  end
  object dsSP_country: TDataSource
    DataSet = qSP_country
    Left = 120
    Top = 496
  end
  object dsSP_city: TDataSource
    DataSet = qSP_city
    Left = 120
    Top = 552
  end
  object qSP_chel: TADOQuery
    Connection = ADOConnection1
    CursorType = ctStatic
    Parameters = <>
    SQL.Strings = (
      'select * from sp_komandir'
      'order by naim')
    Left = 40
    Top = 272
  end
  object qSP_grade: TADOQuery
    Connection = ADOConnection1
    CursorType = ctStatic
    Parameters = <>
    SQL.Strings = (
      'select rowidtochar(rowid) rw,'
      '           s.* '
      'from sp_grade s'
      'order by grade')
    Left = 40
    Top = 328
  end
  object qSP_gostinica: TADOQuery
    Connection = ADOConnection1
    CursorType = ctStatic
    Parameters = <>
    SQL.Strings = (
      'select rowidtochar(rowid) rw,'
      '          (select city from sp_city where kod=kod_city) city,'
      '           s.* '
      'from sp_gostinica s '
      'order by city, nvl(reit,0), gostinica')
    Left = 40
    Top = 384
  end
  object qSP_obekt: TADOQuery
    Connection = ADOConnection1
    CursorType = ctStatic
    Parameters = <>
    SQL.Strings = (
      'select  rowidtochar(rowid) rw,'
      '          (select city from sp_city where kod=kod_city) city,'
      '          s.* '
      'from sp_obekt s'
      'order by city, obekt')
    Left = 40
    Top = 440
  end
  object qSP_country: TADOQuery
    Connection = ADOConnection1
    CursorType = ctStatic
    Parameters = <>
    SQL.Strings = (
      'select rowidtochar(rowid) rw,'
      '           s.* '
      'from sp_country s '
      'order by country')
    Left = 40
    Top = 496
  end
  object qSP_city: TADOQuery
    Connection = ADOConnection1
    CursorType = ctStatic
    Parameters = <>
    SQL.Strings = (
      'select rowidtochar(rowid) rw,'
      
        '          (select country from sp_country where kod=kod_country)' +
        ' country,'
      '          s. * '
      'from sp_city  s'
      'order by city')
    Left = 40
    Top = 552
  end
  object qObnovlenie1: TADOQuery
    Connection = ADOConnection2
    Parameters = <>
    Left = 120
    Top = 96
  end
  object ADOConnection2: TADOConnection
    LoginPrompt = False
    Provider = 'MSDAORA.1'
    Left = 120
    Top = 24
  end
  object ADOConnection1: TADOConnection
    LoginPrompt = False
    Provider = 'MSDAORA.1'
    Left = 40
    Top = 24
  end
end
