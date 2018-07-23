object DM: TDM
  OldCreateOrder = False
  OnCreate = DataModuleCreate
  OnDestroy = DataModuleDestroy
  Left = 250
  Top = 133
  Height = 781
  Width = 213
  object ADOConnection1: TADOConnection
    KeepConnection = False
    LoginPrompt = False
    Provider = 'OraOLEDB.Oracle.1'
    Left = 80
    Top = 32
  end
  object dsOcenka: TDataSource
    DataSet = qOcenka
    Left = 112
    Top = 96
  end
  object qOcenka: TADOQuery
    Connection = ADOConnection1
    CursorType = ctStatic
    Parameters = <
      item
        Name = 'pgod1'
        Size = -1
        Value = Null
      end
      item
        Name = 'pgod3'
        Size = -1
        Value = Null
      end
      item
        Name = 'pgod2'
        Size = -1
        Value = Null
      end
      item
        Name = 'pgod'
        Size = -1
        Value = Null
      end>
    SQL.Strings = (
      'select rowidtochar(rowid) rw, count(*) over() nn,'
      
        '       (select distinct nazv_cexk from ssap_cex where id_cex=o.z' +
        'ex and nazv_cexk not like '#39'%'#1091#1089#1090#1072#1088'.%'#39') knaim_zex,'
      '       initcap(fio) as fio,'
      '      initcap(fio_ocen) as fio_ocen,'
      
        '       decode(trim(upper(funct_g)),'#39#1055#1056#1054#1048#1047#1042#1054#1044#1057#1058#1042#1054#39','#39#1055#39','#39#1042#1053#1059#1058#1056#1045#1053#1053#1048 +
        #1049' '#1057#1045#1056#1042#1048#1057'. '#1055#1056#1054#1044#1040#1046#1048#39','#39#1042#1057#39',funct_g) as funct_g_name,'
      
        '       decode(trim(upper(kat)),'#39#1057#1054#1058#1056#1059#1044#1053#1048#1050#39','#39#1057#39','#39#1051#1048#1053#1045#1049#1053#1067#1049' '#1052#1045#1053#1045#1044#1046#1045 +
        #1056#39','#39#1051#1052#39','#39#1056#1059#1050#1054#1042#1054#1044#1048#1058#1045#1051#1068' '#1055#1054#1044#1056#1040#1047#1044#1045#1051#1045#1053#1048#1071#39','#39#1056#1055#39',kat) as kat_name,'
      '       decode(rezerv,1,'#39#1076#1072#39',0,'#39#1085#1077#1090#39') as rezerv_n,'
      
        '       (select naim_zex from sp_pdirekt k where o.zex=k.zex and ' +
        'k.zex is not null and k.god=:pgod1) as naim_zex,'
      
        '       (select naim from sp_direkt where god=:pgod3 and kod = (s' +
        'elect kod_d from sp_pdirekt z where z.zex=o.direkt and z.god=:pg' +
        'od2)) naim_direkt,'
      '       o.*'
      '       from ocenka o where '
      '/*FILTER*/ 1=1 and god=:pgod'
      '    order by zex, o.fio_ocen, o.funct_g, o.kat, o.efect')
    Left = 48
    Top = 96
  end
  object qObnovlenie: TADOQuery
    Connection = ADOConnection1
    Parameters = <>
    Left = 48
    Top = 416
  end
  object qSprav: TADOQuery
    Connection = ADOConnection1
    CursorType = ctStatic
    Parameters = <
      item
        Name = 'pgod1'
        Size = -1
        Value = Null
      end
      item
        Name = 'pgod'
        Size = -1
        Value = Null
      end>
    SQL.Strings = (
      
        'select zex, naim_zex, (select naim from sp_direkt where god=:pgo' +
        'd1 and kod = s.kod_d) naim_direkt'
      'from sp_pdirekt s where zex is not null and god=:pgod')
    Left = 48
    Top = 152
  end
  object dsSprav: TDataSource
    DataSet = qSprav
    Left = 112
    Top = 152
  end
  object qObnovlenie2: TADOQuery
    Connection = ADOConnection1
    Parameters = <>
    Left = 48
    Top = 472
  end
  object spOcenka: TADOStoredProc
    Connection = ADOConnection1
    ProcedureName = 'CALC_OCENKA'
    Parameters = <
      item
        Name = 'PFIO_OCEN'
        Attributes = [paNullable]
        DataType = ftString
        Size = 4000
        Value = Null
      end
      item
        Name = 'PKAT'
        Attributes = [paNullable]
        DataType = ftString
        Size = 4000
        Value = Null
      end
      item
        Name = 'PFUNCT_G'
        Attributes = [paNullable]
        DataType = ftString
        Size = 4000
        Value = Null
      end
      item
        Name = 'PMIN'
        Attributes = [paNullable]
        DataType = ftFloat
        Value = Null
      end
      item
        Name = 'PMAX'
        Attributes = [paNullable]
        DataType = ftFloat
        Value = Null
      end
      item
        Name = 'PA5'
        Attributes = [paNullable]
        DataType = ftFloat
        Value = Null
      end
      item
        Name = 'PA20'
        Attributes = [paNullable]
        DataType = ftFloat
        Value = Null
      end
      item
        Name = 'PB60'
        Attributes = [paNullable]
        DataType = ftFloat
        Value = Null
      end
      item
        Name = 'PC20'
        Attributes = [paNullable]
        DataType = ftFloat
        Value = Null
      end
      item
        Name = 'PC5'
        Attributes = [paNullable]
        DataType = ftFloat
        Value = Null
      end
      item
        Name = 'PKPE'
        Attributes = [paNullable]
        DataType = ftFloat
        Value = Null
      end
      item
        Name = 'PGOD'
        Attributes = [paNullable]
        DataType = ftInteger
        Value = Null
      end>
    Left = 48
    Top = 664
  end
  object qLogs: TADOQuery
    Connection = ADOConnection1
    CursorType = ctStatic
    Parameters = <>
    SQL.Strings = (
      'select '
      's.* from logs_ocenka s order by dt desc')
    Left = 48
    Top = 536
  end
  object dsLogs: TDataSource
    DataSet = qLogs
    Left = 112
    Top = 536
  end
  object qDolg: TADOQuery
    Connection = ADOConnection1
    Parameters = <>
    SQL.Strings = (
      'select stext as dolg,'
      '       shifr_zex,     '
      '       kod_zex,'
      '       kod_szex,'
      '       objid_p1000,'
      '       short,'
      
        '       (select stext from p1000@sapmig_buffdb where otype='#39'O'#39' an' +
        'd langu='#39'R'#39' and stext not like '#39'%('#1091#1089#1090#1072#1088'.)'#39' and short=kod_zex) as' +
        ' nzex,'
      '       uch'
      ' '
      'from ('
      
        '             (select p1.otype, p1.objid as zvezda1, p1.begda, p1' +
        '.endda, p1.sobid as sobid_p1001, p2.stext, kod as zvezda3, p2.ob' +
        'jid, p2.short'
      
        '              from (select r.otype, r.objid, r.begda, r.endda, s' +
        '.sobid as sobid, s.objid as kod from p1013@sapmig_buffdb r left ' +
        'join p1001@sapmig_buffdb s on r.objid=s.objid and s.otype='#39'S'#39' wh' +
        'ere r.otype='#39'S'#39' and r.persk=10) p1,'
      '                    p1000@sapmig_buffdb p2'
      
        '              where p2.otype='#39'S'#39' and p2.langu='#39'R'#39' and p1.objid=p' +
        '2.objid) obsh1'
      '            left join'
      
        '              (select objid as objid_p1000, short as shifr_zex, ' +
        'substr(short,1,2) as kod_zex, substr(short,1,5) as kod_szex, ste' +
        'xt as uch from p1000@sapmig_buffdb where otype='#39'O'#39' and langu='#39'R'#39 +
        ')obsh2'
      '               on  sobid_p1001=objid_p1000'
      '     ) '
      'where endda>sysdate '
      
        'group by stext, shifr_zex, kod_zex, objid_p1000, short, uch, kod' +
        '_szex'
      'order by shifr_zex, dolg')
    Left = 48
    Top = 600
  end
  object dsDolg: TDataSource
    DataSet = qDolg
    Left = 112
    Top = 600
  end
  object qRezerv: TADOQuery
    Connection = ADOConnection1
    CursorType = ctStatic
    Parameters = <
      item
        Name = 'pgod'
        DataType = ftString
        Size = 4
        Value = '2016'
      end
      item
        Name = 'ptn'
        Size = -1
        Value = Null
      end
      item
        Name = 'pshtat'
        Size = -1
        Value = Null
      end>
    SQL.Strings = (
      
        'select * from ocenka_rez where god=:pgod and tn=:ptn and id_shta' +
        't=:pshtat')
    Left = 48
    Top = 208
  end
  object dsRezerv: TDataSource
    DataSet = qRezerv
    Left = 112
    Top = 208
  end
  object qProverka: TADOQuery
    Connection = ADOConnection1
    Parameters = <
      item
        Name = 'ptn'
        Size = -1
        Value = Null
      end
      item
        Name = 'pgod'
        Size = -1
        Value = Null
      end>
    SQL.Strings = (
      'select * from ocenka where tn=:ptn and god=:pgod')
    Left = 48
    Top = 368
  end
  object dsZamesh: TDataSource
    DataSet = qZamesh
    OnDataChange = dsZameshDataChange
    Left = 112
    Top = 272
  end
  object qZamesh: TADOQuery
    Connection = ADOConnection1
    Parameters = <
      item
        Name = 'pgod'
        Size = -1
        Value = Null
      end>
    SQL.Strings = (
      'select rowidtochar(rowid) rw, r.* from ocenka_rez r'
      'where god=:pgod')
    Left = 48
    Top = 272
  end
end
