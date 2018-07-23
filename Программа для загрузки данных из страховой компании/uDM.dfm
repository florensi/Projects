object DM: TDM
  OldCreateOrder = False
  OnCreate = DataModuleCreate
  OnDestroy = DataModuleDestroy
  Left = 379
  Top = 212
  Height = 298
  Width = 239
  object ADOConnection1: TADOConnection
    LoginPrompt = False
    Provider = 'MSDAORA.1'
    Left = 96
    Top = 40
  end
  object qZagruzka: TADOQuery
    Connection = ADOConnection1
    Parameters = <>
    Left = 128
    Top = 105
  end
  object qObnovlenie: TADOQuery
    Connection = ADOConnection1
    Parameters = <>
    Left = 56
    Top = 104
  end
  object qKorrektirovka: TADOQuery
    Connection = ADOConnection1
    CursorType = ctStatic
    Parameters = <
      item
        Name = 'pzex'
        Size = -1
        Value = Null
      end
      item
        Name = 'ptn'
        Size = -1
        Value = Null
      end>
    SQL.Strings = (
      'select v.*, rowidtochar(rowid) rw from vu_859_n v '
      'where zex=:pzex and tn=:ptn'
      
        'and (inn in (select numident from sap_osn_sved) or inn in (selec' +
        't numident from sap_sved_uvol))')
    Left = 56
    Top = 168
  end
  object dsKorrektirovka: TDataSource
    DataSet = qKorrektirovka
    OnDataChange = dsKorrektirovkaDataChange
    Left = 136
    Top = 168
  end
end
