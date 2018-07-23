object DM: TDM
  OldCreateOrder = False
  OnCreate = DataModuleCreate
  OnDestroy = DataModuleDestroy
  Height = 467
  Width = 218
  object UniConnection1: TUniConnection
    SpecificOptions.Strings = (
      'Oracle.Direct=True')
    Left = 48
    Top = 104
  end
  object qObnovlenie: TUniQuery
    Connection = UniConnection1
    Left = 48
    Top = 384
  end
  object qReiting: TUniQuery
    Connection = UniConnection1
    SQL.Strings = (
      'select rowidtochar(rowid) rw, '
      '       (select nazv_cexk from ssap_cex '
      
        '        where nazv_cexk not like '#39'%('#1091#1089#1090#1072#1088'.)%'#39' and id_cex=r.zex) ' +
        'as zex_naim,'
      '       decode(podch,1,'#39#1076#1072#39',0,'#39#1085#1077#1090#39',null) as npodch,'
      
        '       (select pz from sp_reit_proizv p where p.zex=r.zex) as pz' +
        ','
      '       r.*'
      'from reit_ruk r'
      'where god=:pgod and kvart=:pkvartal'
      ''
      'order by zex, nvl(podch,0), nvl(reit,0), ocenka desc, tn'
      ''
      ''
      ''
      '')
    SpecificOptions.Strings = (
      'Oracle.FetchAll=True')
    Left = 48
    Top = 184
    ParamData = <
      item
        DataType = ftUnknown
        Name = 'pgod'
        Value = nil
      end
      item
        DataType = ftUnknown
        Name = 'pkvartal'
        Value = nil
      end>
  end
  object OracleUniProvider1: TOracleUniProvider
    Left = 48
    Top = 33
  end
  object dsReiting: TDataSource
    DataSet = qReiting
    Left = 96
    Top = 184
  end
  object qProverka: TUniQuery
    Connection = UniConnection1
    SQL.Strings = (
      'select zex, nzex, tn_sap,                      '
      #9'initcap(fam||'#39' '#39'||im||'#39' '#39'||ot) as fio,   '
      #9'id_shtat, name_dolg_ru                    '
      #9'from sap_osn_sved '
      'where tn_sap=:ptn_sap')
    Left = 48
    Top = 336
    ParamData = <
      item
        DataType = ftUnknown
        Name = 'ptn_sap'
        Value = nil
      end>
  end
  object qObnovlenie2: TUniQuery
    Connection = UniConnection1
    Left = 136
    Top = 384
  end
  object qRaschet: TUniQuery
    Connection = UniConnection1
    SQL.Strings = (
      ' select '
      '        zex,'
      '        tn, '
      '        fio,'
      '        ocenka,'
      '        nvl(podch,0) as podch, '
      
        '        count(*) over (partition by zex, nvl(podch,0)) as kol_ze' +
        'x,'
      
        '        min(ocenka) over (partition by zex, nvl(podch,0)) as zn_' +
        'min,'
      
        '        max(ocenka) over (partition by zex, nvl(podch,0)) as zn_' +
        'max,'
      
        '        count(*) over (partition by zex, ocenka, nvl(podch,0)) a' +
        's kol_min,'
      
        '        (count(*) over (partition by zex, nvl(podch,0)))*0.2 as ' +
        'zona     '
      
        'from reit_ruk s where god=:pgod and kvart=:pkvart and reit is nu' +
        'll'
      'and zex=:pzex and nvl(podch,0)=:ppodch '
      'order by ocenka')
    Left = 136
    Top = 336
    ParamData = <
      item
        DataType = ftUnknown
        Name = 'pgod'
        Value = nil
      end
      item
        DataType = ftUnknown
        Name = 'pkvart'
        Value = nil
      end
      item
        DataType = ftUnknown
        Name = 'pzex'
        Value = nil
      end
      item
        DataType = ftUnknown
        Name = 'ppodch'
        Value = nil
      end>
  end
  object qSprav: TUniQuery
    Connection = UniConnection1
    SQL.Strings = (
      'select sp.*, rowidtochar(rowid) rw,'
      
        'decode(pz,'#39'1'#39','#39#1076#1072#39','#39#1085#1077#1090#39') as pz1 from sp_reit_proizv sp order by' +
        ' zex')
    Left = 48
    Top = 240
  end
  object dsSprav: TDataSource
    DataSet = qSprav
    Left = 96
    Top = 240
  end
end
