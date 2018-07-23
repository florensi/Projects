object DM: TDM
  OldCreateOrder = False
  OnCreate = DataModuleCreate
  OnDestroy = DataModuleDestroy
  Left = 325
  Top = 175
  Height = 662
  Width = 213
  object ADOConnection1: TADOConnection
    LoginPrompt = False
    Provider = 'MSDAORA.1'
    Left = 78
    Top = 16
  end
  object qObnovlenie: TADOQuery
    Connection = ADOConnection1
    Parameters = <>
    Left = 38
    Top = 72
  end
  object qGrafik: TADOQuery
    Connection = ADOConnection1
    CursorType = ctStatic
    OnCalcFields = qGrafikCalcFields
    Parameters = <
      item
        Name = 'pgod'
        Size = -1
        Value = Null
      end
      item
        Name = 'pograf'
        Size = -1
        Value = Null
      end>
    SQL.Strings = (
      'select s.ograf,graf,mes, god,'
      '       dlit, otchet, br, name,       '
      
        'decode(mes, 1, '#39#1071#1085#1074#1072#1088#1100#39',2, '#39#1060#1077#1074#1088#1072#1083#1100#39',3, '#39#1052#1072#1088#1090#39',4, '#39#1040#1087#1088#1077#1083#1100#39',5, '#39#1052 +
        #1072#1081#39',6, '#39#1048#1102#1085#1100#39',7, '#39#1048#1102#1083#1100#39',8, '#39#1040#1074#1075#1091#1089#1090#39',9, '#39#1057#1077#1085#1090#1103#1073#1088#1100#39',10, '#39#1054#1082#1090#1103#1073#1088#1100#39',' +
        '11, '#39#1053#1086#1103#1073#1088#1100#39',12, '#39#1044#1077#1082#1072#1073#1088#1100#39') as mes1,'
      'chf0,'
      'nch0,'
      'pch0,'
      'chf,'
      'decode(pgraf,0,to_number(null),pgraf) as pgraf,                '
      'decode(nch,0,to_number(null),nch) as nch,'
      'decode(vch,0,to_number(null),vch) as vch,'
      'decode(pch,0,to_number(null),pch) as pch, '
      
        'case when (s.ograf=23) then  (select distinct round(chf/2,0) fro' +
        'm spgrafiki k where k.ograf=f.norma and k.god=s.god and k.mes=s.' +
        'mes)'
      
        '     else (select distinct chf from spgrafiki k where k.ograf=f.' +
        'norma and k.god=s.god and k.mes=s.mes) end as norma,            ' +
        '   '
      
        'decode(nsm1,9,'#39#1055#39',30,'#39#1055#39','#39#39','#39'-'#39',nsm1) as nsm1,                  ' +
        '      '
      
        'decode(nsm2,9,'#39#1055#39',30,'#39#1055#39','#39#39','#39'-'#39',nsm2) as nsm2,                  ' +
        '      '
      
        'decode(nsm3,9,'#39#1055#39',30,'#39#1055#39','#39#39','#39'-'#39',nsm3) as nsm3,                  ' +
        '      '
      
        'decode(nsm4,9,'#39#1055#39',30,'#39#1055#39','#39#39','#39'-'#39',nsm4) as nsm4,                  ' +
        '      '
      
        'decode(nsm5,9,'#39#1055#39',30,'#39#1055#39','#39#39','#39'-'#39',nsm5) as nsm5,                  ' +
        '      '
      
        'decode(nsm6,9,'#39#1055#39',30,'#39#1055#39','#39#39','#39'-'#39',nsm6) as nsm6,                  ' +
        '      '
      'decode(nsm7,9,'#39#1055#39',30,'#39#1055#39','#39#39','#39'-'#39',nsm7) as nsm7,'
      
        'decode(nsm8,9,'#39#1055#39',30,'#39#1055#39','#39#39','#39'-'#39',nsm8) as nsm8,                  ' +
        '      '
      
        'decode(nsm9,9,'#39#1055#39',30,'#39#1055#39','#39#39','#39'-'#39',nsm9) as nsm9,                  ' +
        '      '
      
        'decode(nsm10,9,'#39#1055#39',30,'#39#1055#39','#39#39','#39'-'#39',nsm10) as nsm10,               ' +
        '      '
      
        'decode(nsm11,9,'#39#1055#39',30,'#39#1055#39','#39#39','#39'-'#39',nsm11) as nsm11,               ' +
        '      '
      
        'decode(nsm12,9,'#39#1055#39',30,'#39#1055#39','#39#39','#39'-'#39',nsm12) as nsm12,               ' +
        '      '
      
        'decode(nsm13,9,'#39#1055#39',30,'#39#1055#39','#39#39','#39'-'#39',nsm13) as nsm13,               ' +
        '      '
      
        'decode(nsm14,9,'#39#1055#39',30,'#39#1055#39','#39#39','#39'-'#39',nsm14) as nsm14,               ' +
        '      '
      
        'decode(nsm15,9,'#39#1055#39',30,'#39#1055#39','#39#39','#39'-'#39',nsm15) as nsm15,               ' +
        '      '
      
        'decode(nsm16,9,'#39#1055#39',30,'#39#1055#39','#39#39','#39'-'#39',nsm16) as nsm16,               ' +
        '      '
      
        'decode(nsm17,9,'#39#1055#39',30,'#39#1055#39','#39#39','#39'-'#39',nsm17) as nsm17,               ' +
        '      '
      
        'decode(nsm18,9,'#39#1055#39',30,'#39#1055#39','#39#39','#39'-'#39',nsm18) as nsm18,               ' +
        '      '
      
        'decode(nsm19,9,'#39#1055#39',30,'#39#1055#39','#39#39','#39'-'#39',nsm19) as nsm19,               ' +
        '      '
      
        'decode(nsm20,9,'#39#1055#39',30,'#39#1055#39','#39#39','#39'-'#39',nsm20) as nsm20,               ' +
        '      '
      
        'decode(nsm21,9,'#39#1055#39',30,'#39#1055#39','#39#39','#39'-'#39',nsm21) as nsm21,               ' +
        '      '
      
        'decode(nsm22,9,'#39#1055#39',30,'#39#1055#39','#39#39','#39'-'#39',nsm22) as nsm22,               ' +
        '      '
      
        'decode(nsm23,9,'#39#1055#39',30,'#39#1055#39','#39#39','#39'-'#39',nsm23) as nsm23,               ' +
        '      '
      
        'decode(nsm24,9,'#39#1055#39',30,'#39#1055#39','#39#39','#39'-'#39',nsm24) as nsm24,               ' +
        '      '
      
        'decode(nsm25,9,'#39#1055#39',30,'#39#1055#39','#39#39','#39'-'#39',nsm25) as nsm25,               ' +
        '      '
      
        'decode(nsm26,9,'#39#1055#39',30,'#39#1055#39','#39#39','#39'-'#39',nsm26) as nsm26,               ' +
        '      '
      
        'decode(nsm27,9,'#39#1055#39',30,'#39#1055#39','#39#39','#39'-'#39',nsm27) as nsm27,               ' +
        '      '
      
        'decode(nsm28,9,'#39#1055#39',30,'#39#1055#39','#39#39','#39'-'#39',nsm28) as nsm28,               ' +
        '      '
      
        'decode(nsm29,9,'#39#1055#39',30,'#39#1055#39','#39#39','#39'-'#39',nsm29) as nsm29,               ' +
        '      '
      
        'decode(nsm30,9,'#39#1055#39',30,'#39#1055#39','#39#39','#39'-'#39',nsm30) as nsm30,               ' +
        '      '
      
        'decode(nsm31,9,'#39#1055#39',30,'#39#1055#39','#39#39','#39'-'#39',nsm31) as nsm31,               ' +
        '      '
      
        'nch1, nch2, nch3, nch4, nch5, nch6, nch7, nch8,                 ' +
        '      '
      
        'nch9, nch10, nch11, nch12, nch13, nch14, nch15,                 ' +
        '      '
      
        'nch16, nch17, nch18, nch19, nch20, nch21, nch22,                ' +
        '      '
      
        'nch23, nch24, nch25,nch26, nch27, nch28, nch29,                 ' +
        '      '
      
        'nch30, nch31,                                                   ' +
        '      '
      
        'vch1, vch2, vch3, vch4, vch5, vch6, vch7, vch8,                 ' +
        '      '
      
        'vch9, vch10, vch11, vch12, vch13, vch14, vch15,                 ' +
        '      '
      
        'vch16, vch17, vch18, vch19, vch20, vch21, vch22,                ' +
        '      '
      
        'vch23, vch24, vch25, vch26, vch27, vch28, vch29,                ' +
        '      '
      
        'vch30, vch31,                                                   ' +
        '      '
      
        'pch1, pch2, pch3, pch4, pch5, pch6, pch7, pch8,                 ' +
        '      '
      
        'pch9, pch10, pch11, pch12, pch13, pch14, pch15,                 ' +
        '      '
      
        'pch16, pch17, pch18, pch19, pch20, pch21, pch22,                ' +
        '      '
      
        'pch23, pch24, pch25, pch26, pch27, pch28, pch29,                ' +
        '      '
      
        'pch30, pch31,                                                   ' +
        '      '
      
        'decode(chf1,30,'#39#1055#39',to_number(null),'#39'-'#39',chf1)as chf1,            ' +
        '     '
      
        'decode(chf2,30,'#39#1055#39',to_number(null),'#39'-'#39',chf2)as chf2,            ' +
        '                  '
      
        'decode(chf3,30,'#39#1055#39',to_number(null),'#39'-'#39',chf3)as chf3,            ' +
        '                  '
      
        'decode(chf4,30,'#39#1055#39',to_number(null),'#39'-'#39',chf4)as chf4,            ' +
        '                  '
      
        'decode(chf5,30,'#39#1055#39',to_number(null),'#39'-'#39',chf5)as chf5,            ' +
        '                  '
      
        'decode(chf6,30,'#39#1055#39',to_number(null),'#39'-'#39',chf6)as chf6,            ' +
        '                  '
      
        'decode(chf7,30,'#39#1055#39',to_number(null),'#39'-'#39',chf7)as chf7,            ' +
        '                  '
      
        'decode(chf8,30,'#39#1055#39',to_number(null),'#39'-'#39',chf8)as chf8,            ' +
        '                  '
      
        'decode(chf9,30,'#39#1055#39',to_number(null),'#39'-'#39',chf9)as chf9,            ' +
        '                  '
      
        'decode(chf10,30,'#39#1055#39',to_number(null),'#39'-'#39',chf10)as chf10,         ' +
        '                  '
      
        'decode(chf11,30,'#39#1055#39',to_number(null),'#39'-'#39',chf11)as chf11,         ' +
        '                  '
      
        'decode(chf12,30,'#39#1055#39',to_number(null),'#39'-'#39',chf12)as chf12,         ' +
        '                  '
      
        'decode(chf13,30,'#39#1055#39',to_number(null),'#39'-'#39',chf13)as chf13,         ' +
        '                  '
      
        'decode(chf14,30,'#39#1055#39',to_number(null),'#39'-'#39',chf14)as chf14,         ' +
        '                  '
      
        'decode(chf15,30,'#39#1055#39',to_number(null),'#39'-'#39',chf15)as chf15,         ' +
        '                  '
      
        'decode(chf16,30,'#39#1055#39',to_number(null),'#39'-'#39',chf16)as chf16,         ' +
        '                  '
      
        'decode(chf17,30,'#39#1055#39',to_number(null),'#39'-'#39',chf17)as chf17,         ' +
        '                  '
      
        'decode(chf18,30,'#39#1055#39',to_number(null),'#39'-'#39',chf18)as chf18,         ' +
        '                  '
      
        'decode(chf19,30,'#39#1055#39',to_number(null),'#39'-'#39',chf19)as chf19,         ' +
        '                  '
      
        'decode(chf20,30,'#39#1055#39',to_number(null),'#39'-'#39',chf20)as chf20,         ' +
        '                  '
      
        'decode(chf21,30,'#39#1055#39',to_number(null),'#39'-'#39',chf21)as chf21,         ' +
        '                  '
      
        'decode(chf22,30,'#39#1055#39',to_number(null),'#39'-'#39',chf22)as chf22,         ' +
        '                  '
      
        'decode(chf23,30,'#39#1055#39',to_number(null),'#39'-'#39',chf23)as chf23,         ' +
        '                  '
      
        'decode(chf24,30,'#39#1055#39',to_number(null),'#39'-'#39',chf24)as chf24,         ' +
        '                  '
      
        'decode(chf25,30,'#39#1055#39',to_number(null),'#39'-'#39',chf25)as chf25,         ' +
        '                  '
      
        'decode(chf26,30,'#39#1055#39',to_number(null),'#39'-'#39',chf26)as chf26,         ' +
        '                  '
      
        'decode(chf27,30,'#39#1055#39',to_number(null),'#39'-'#39',chf27)as chf27,         ' +
        '                  '
      
        'decode(chf28,30,'#39#1055#39',to_number(null),'#39'-'#39',chf28)as chf28,         ' +
        '                  '
      
        'decode(chf29,30,'#39#1055#39',to_number(null),'#39'-'#39',chf29)as chf29,         ' +
        '                  '
      
        'decode(chf30,30,'#39#1055#39',to_number(null),'#39'-'#39',chf30)as chf30,         ' +
        '                  '
      
        'decode(chf31,30,'#39#1055#39',to_number(null),'#39'-'#39',chf31)as chf31          ' +
        '           '
      'from spgrafiki s left join spograf f on s.ograf=f.ograf'
      'where god=:pgod and s.ograf=:pograf'
      'order by mes, graf')
    Left = 38
    Top = 128
    object qGrafikf1: TStringField
      FieldKind = fkCalculated
      FieldName = 'f1'
      Size = 5
      Calculated = True
    end
    object qGrafikf2: TStringField
      FieldKind = fkCalculated
      FieldName = 'f2'
      Size = 5
      Calculated = True
    end
    object qGrafikf3: TStringField
      FieldKind = fkCalculated
      FieldName = 'f3'
      Size = 5
      Calculated = True
    end
    object qGrafikf4: TStringField
      FieldKind = fkCalculated
      FieldName = 'f4'
      Size = 5
      Calculated = True
    end
    object qGrafikf5: TStringField
      FieldKind = fkCalculated
      FieldName = 'f5'
      Size = 5
      Calculated = True
    end
    object qGrafikf6: TStringField
      FieldKind = fkCalculated
      FieldName = 'f6'
      Size = 5
      Calculated = True
    end
    object qGrafikf7: TStringField
      FieldKind = fkCalculated
      FieldName = 'f7'
      Size = 5
      Calculated = True
    end
    object qGrafikf8: TStringField
      FieldKind = fkCalculated
      FieldName = 'f8'
      Size = 5
      Calculated = True
    end
    object qGrafikf9: TStringField
      FieldKind = fkCalculated
      FieldName = 'f9'
      Size = 5
      Calculated = True
    end
    object qGrafikf10: TStringField
      FieldKind = fkCalculated
      FieldName = 'f10'
      Size = 5
      Calculated = True
    end
    object qGrafikf11: TStringField
      FieldKind = fkCalculated
      FieldName = 'f11'
      Size = 5
      Calculated = True
    end
    object qGrafikf12: TStringField
      FieldKind = fkCalculated
      FieldName = 'f12'
      Size = 5
      Calculated = True
    end
    object qGrafikf13: TStringField
      FieldKind = fkCalculated
      FieldName = 'f13'
      Size = 5
      Calculated = True
    end
    object qGrafikf14: TStringField
      FieldKind = fkCalculated
      FieldName = 'f14'
      Size = 5
      Calculated = True
    end
    object qGrafikf15: TStringField
      FieldKind = fkCalculated
      FieldName = 'f15'
      Size = 5
      Calculated = True
    end
    object qGrafikf16: TStringField
      FieldKind = fkCalculated
      FieldName = 'f16'
      Size = 5
      Calculated = True
    end
    object qGrafikf17: TStringField
      FieldKind = fkCalculated
      FieldName = 'f17'
      Size = 5
      Calculated = True
    end
    object qGrafikf18: TStringField
      FieldKind = fkCalculated
      FieldName = 'f18'
      Size = 5
      Calculated = True
    end
    object qGrafikf19: TStringField
      FieldKind = fkCalculated
      FieldName = 'f19'
      Size = 5
      Calculated = True
    end
    object qGrafikf20: TStringField
      FieldKind = fkCalculated
      FieldName = 'f20'
      Size = 5
      Calculated = True
    end
    object qGrafikf21: TStringField
      FieldKind = fkCalculated
      FieldName = 'f21'
      Size = 5
      Calculated = True
    end
    object qGrafikf22: TStringField
      FieldKind = fkCalculated
      FieldName = 'f22'
      Size = 5
      Calculated = True
    end
    object qGrafikf23: TStringField
      FieldKind = fkCalculated
      FieldName = 'f23'
      Size = 5
      Calculated = True
    end
    object qGrafikf24: TStringField
      FieldKind = fkCalculated
      FieldName = 'f24'
      Size = 5
      Calculated = True
    end
    object qGrafikf25: TStringField
      FieldKind = fkCalculated
      FieldName = 'f25'
      Size = 5
      Calculated = True
    end
    object qGrafikf26: TStringField
      FieldKind = fkCalculated
      FieldName = 'f26'
      Size = 5
      Calculated = True
    end
    object qGrafikf27: TStringField
      FieldKind = fkCalculated
      FieldName = 'f27'
      Size = 5
      Calculated = True
    end
    object qGrafikf28: TStringField
      FieldKind = fkCalculated
      FieldName = 'f28'
      Size = 5
      Calculated = True
    end
    object qGrafikf29: TStringField
      FieldKind = fkCalculated
      FieldName = 'f29'
      Size = 5
      Calculated = True
    end
    object qGrafikf30: TStringField
      FieldKind = fkCalculated
      FieldName = 'f30'
      Size = 5
      Calculated = True
    end
    object qGrafikf31: TStringField
      FieldKind = fkCalculated
      FieldName = 'f31'
      Size = 5
      Calculated = True
    end
    object qGrafikOGRAF: TIntegerField
      FieldName = 'OGRAF'
    end
    object qGrafikGRAF: TIntegerField
      FieldName = 'GRAF'
    end
    object qGrafikMES: TIntegerField
      FieldName = 'MES'
    end
    object qGrafikMES1: TStringField
      FieldName = 'MES1'
      Size = 8
    end
    object qGrafikCHF: TBCDField
      FieldName = 'CHF'
      Precision = 5
      Size = 2
    end
    object qGrafikPGRAF: TBCDField
      FieldName = 'PGRAF'
      Precision = 32
      Size = 0
    end
    object qGrafikNCH: TBCDField
      FieldName = 'NCH'
      Precision = 32
      Size = 0
    end
    object qGrafikVCH: TBCDField
      FieldName = 'VCH'
      Precision = 32
      Size = 0
    end
    object qGrafikPCH: TBCDField
      FieldName = 'PCH'
      Precision = 32
      Size = 0
    end
    object qGrafikNORMA: TBCDField
      FieldName = 'NORMA'
      Precision = 32
      Size = 0
    end
    object qGrafikNAME: TStringField
      FieldName = 'NAME'
      Size = 120
    end
    object qGrafikNSM1: TStringField
      FieldName = 'NSM1'
      Size = 40
    end
    object qGrafikNSM2: TStringField
      FieldName = 'NSM2'
      Size = 40
    end
    object qGrafikNSM3: TStringField
      FieldName = 'NSM3'
      Size = 40
    end
    object qGrafikNSM4: TStringField
      FieldName = 'NSM4'
      Size = 40
    end
    object qGrafikNSM5: TStringField
      FieldName = 'NSM5'
      Size = 40
    end
    object qGrafikNSM6: TStringField
      FieldName = 'NSM6'
      Size = 40
    end
    object qGrafikNSM7: TStringField
      FieldName = 'NSM7'
      Size = 40
    end
    object qGrafikNSM8: TStringField
      FieldName = 'NSM8'
      Size = 40
    end
    object qGrafikNSM9: TStringField
      FieldName = 'NSM9'
      Size = 40
    end
    object qGrafikNSM10: TStringField
      FieldName = 'NSM10'
      Size = 40
    end
    object qGrafikNSM11: TStringField
      FieldName = 'NSM11'
      Size = 40
    end
    object qGrafikNSM12: TStringField
      FieldName = 'NSM12'
      Size = 40
    end
    object qGrafikNSM13: TStringField
      FieldName = 'NSM13'
      Size = 40
    end
    object qGrafikNSM14: TStringField
      FieldName = 'NSM14'
      Size = 40
    end
    object qGrafikNSM15: TStringField
      FieldName = 'NSM15'
      Size = 40
    end
    object qGrafikNSM16: TStringField
      FieldName = 'NSM16'
      Size = 40
    end
    object qGrafikNSM17: TStringField
      FieldName = 'NSM17'
      Size = 40
    end
    object qGrafikNSM18: TStringField
      FieldName = 'NSM18'
      Size = 40
    end
    object qGrafikNSM19: TStringField
      FieldName = 'NSM19'
      Size = 40
    end
    object qGrafikNSM20: TStringField
      FieldName = 'NSM20'
      Size = 40
    end
    object qGrafikNSM21: TStringField
      FieldName = 'NSM21'
      Size = 40
    end
    object qGrafikNSM22: TStringField
      FieldName = 'NSM22'
      Size = 40
    end
    object qGrafikNSM23: TStringField
      FieldName = 'NSM23'
      Size = 40
    end
    object qGrafikNSM24: TStringField
      FieldName = 'NSM24'
      Size = 40
    end
    object qGrafikNSM25: TStringField
      FieldName = 'NSM25'
      Size = 40
    end
    object qGrafikNSM26: TStringField
      FieldName = 'NSM26'
      Size = 40
    end
    object qGrafikNSM27: TStringField
      FieldName = 'NSM27'
      Size = 40
    end
    object qGrafikNSM28: TStringField
      FieldName = 'NSM28'
      Size = 40
    end
    object qGrafikNSM29: TStringField
      FieldName = 'NSM29'
      Size = 40
    end
    object qGrafikNSM30: TStringField
      FieldName = 'NSM30'
      Size = 40
    end
    object qGrafikNSM31: TStringField
      FieldName = 'NSM31'
      Size = 40
    end
    object qGrafikNCH1: TBCDField
      FieldName = 'NCH1'
      Precision = 4
      Size = 2
    end
    object qGrafikNCH2: TBCDField
      FieldName = 'NCH2'
      Precision = 4
      Size = 2
    end
    object qGrafikNCH3: TBCDField
      FieldName = 'NCH3'
      Precision = 4
      Size = 2
    end
    object qGrafikNCH4: TBCDField
      FieldName = 'NCH4'
      Precision = 4
      Size = 2
    end
    object qGrafikNCH5: TBCDField
      FieldName = 'NCH5'
      Precision = 4
      Size = 2
    end
    object qGrafikNCH6: TBCDField
      FieldName = 'NCH6'
      Precision = 4
      Size = 2
    end
    object qGrafikNCH7: TBCDField
      FieldName = 'NCH7'
      Precision = 4
      Size = 2
    end
    object qGrafikNCH8: TBCDField
      FieldName = 'NCH8'
      Precision = 4
      Size = 2
    end
    object qGrafikNCH9: TBCDField
      FieldName = 'NCH9'
      Precision = 4
      Size = 2
    end
    object qGrafikNCH10: TBCDField
      FieldName = 'NCH10'
      Precision = 4
      Size = 2
    end
    object qGrafikNCH11: TBCDField
      FieldName = 'NCH11'
      Precision = 4
      Size = 2
    end
    object qGrafikNCH12: TBCDField
      FieldName = 'NCH12'
      Precision = 4
      Size = 2
    end
    object qGrafikNCH13: TBCDField
      FieldName = 'NCH13'
      Precision = 4
      Size = 2
    end
    object qGrafikNCH14: TBCDField
      FieldName = 'NCH14'
      Precision = 4
      Size = 2
    end
    object qGrafikNCH15: TBCDField
      FieldName = 'NCH15'
      Precision = 4
      Size = 2
    end
    object qGrafikNCH16: TBCDField
      FieldName = 'NCH16'
      Precision = 4
      Size = 2
    end
    object qGrafikNCH17: TBCDField
      FieldName = 'NCH17'
      Precision = 4
      Size = 2
    end
    object qGrafikNCH18: TBCDField
      FieldName = 'NCH18'
      Precision = 4
      Size = 2
    end
    object qGrafikNCH19: TBCDField
      FieldName = 'NCH19'
      Precision = 4
      Size = 2
    end
    object qGrafikNCH20: TBCDField
      FieldName = 'NCH20'
      Precision = 4
      Size = 2
    end
    object qGrafikNCH21: TBCDField
      FieldName = 'NCH21'
      Precision = 4
      Size = 2
    end
    object qGrafikNCH22: TBCDField
      FieldName = 'NCH22'
      Precision = 4
      Size = 2
    end
    object qGrafikNCH23: TBCDField
      FieldName = 'NCH23'
      Precision = 4
      Size = 2
    end
    object qGrafikNCH24: TBCDField
      FieldName = 'NCH24'
      Precision = 4
      Size = 2
    end
    object qGrafikNCH25: TBCDField
      FieldName = 'NCH25'
      Precision = 4
      Size = 2
    end
    object qGrafikNCH26: TBCDField
      FieldName = 'NCH26'
      Precision = 4
      Size = 2
    end
    object qGrafikNCH27: TBCDField
      FieldName = 'NCH27'
      Precision = 4
      Size = 2
    end
    object qGrafikNCH28: TBCDField
      FieldName = 'NCH28'
      Precision = 4
      Size = 2
    end
    object qGrafikNCH29: TBCDField
      FieldName = 'NCH29'
      Precision = 4
      Size = 2
    end
    object qGrafikNCH30: TBCDField
      FieldName = 'NCH30'
      Precision = 4
      Size = 2
    end
    object qGrafikNCH31: TBCDField
      FieldName = 'NCH31'
      Precision = 4
      Size = 2
    end
    object qGrafikVCH1: TBCDField
      FieldName = 'VCH1'
      Precision = 4
      Size = 2
    end
    object qGrafikVCH2: TBCDField
      FieldName = 'VCH2'
      Precision = 4
      Size = 2
    end
    object qGrafikVCH3: TBCDField
      FieldName = 'VCH3'
      Precision = 4
      Size = 2
    end
    object qGrafikVCH4: TBCDField
      FieldName = 'VCH4'
      Precision = 4
      Size = 2
    end
    object qGrafikVCH5: TBCDField
      FieldName = 'VCH5'
      Precision = 4
      Size = 2
    end
    object qGrafikVCH6: TBCDField
      FieldName = 'VCH6'
      Precision = 4
      Size = 2
    end
    object qGrafikVCH7: TBCDField
      FieldName = 'VCH7'
      Precision = 4
      Size = 2
    end
    object qGrafikVCH8: TBCDField
      FieldName = 'VCH8'
      Precision = 4
      Size = 2
    end
    object qGrafikVCH9: TBCDField
      FieldName = 'VCH9'
      Precision = 4
      Size = 2
    end
    object qGrafikVCH10: TBCDField
      FieldName = 'VCH10'
      Precision = 4
      Size = 2
    end
    object qGrafikVCH11: TBCDField
      FieldName = 'VCH11'
      Precision = 4
      Size = 2
    end
    object qGrafikVCH12: TBCDField
      FieldName = 'VCH12'
      Precision = 4
      Size = 2
    end
    object qGrafikVCH13: TBCDField
      FieldName = 'VCH13'
      Precision = 4
      Size = 2
    end
    object qGrafikVCH14: TBCDField
      FieldName = 'VCH14'
      Precision = 4
      Size = 2
    end
    object qGrafikVCH15: TBCDField
      FieldName = 'VCH15'
      Precision = 4
      Size = 2
    end
    object qGrafikVCH16: TBCDField
      FieldName = 'VCH16'
      Precision = 4
      Size = 2
    end
    object qGrafikVCH17: TBCDField
      FieldName = 'VCH17'
      Precision = 4
      Size = 2
    end
    object qGrafikVCH18: TBCDField
      FieldName = 'VCH18'
      Precision = 4
      Size = 2
    end
    object qGrafikVCH19: TBCDField
      FieldName = 'VCH19'
      Precision = 4
      Size = 2
    end
    object qGrafikVCH20: TBCDField
      FieldName = 'VCH20'
      Precision = 4
      Size = 2
    end
    object qGrafikVCH21: TBCDField
      FieldName = 'VCH21'
      Precision = 4
      Size = 2
    end
    object qGrafikVCH22: TBCDField
      FieldName = 'VCH22'
      Precision = 4
      Size = 2
    end
    object qGrafikVCH23: TBCDField
      FieldName = 'VCH23'
      Precision = 4
      Size = 2
    end
    object qGrafikVCH24: TBCDField
      FieldName = 'VCH24'
      Precision = 4
      Size = 2
    end
    object qGrafikVCH25: TBCDField
      FieldName = 'VCH25'
      Precision = 4
      Size = 2
    end
    object qGrafikVCH26: TBCDField
      FieldName = 'VCH26'
      Precision = 4
      Size = 2
    end
    object qGrafikVCH27: TBCDField
      FieldName = 'VCH27'
      Precision = 4
      Size = 2
    end
    object qGrafikVCH28: TBCDField
      FieldName = 'VCH28'
      Precision = 4
      Size = 2
    end
    object qGrafikVCH29: TBCDField
      FieldName = 'VCH29'
      Precision = 4
      Size = 2
    end
    object qGrafikVCH30: TBCDField
      FieldName = 'VCH30'
      Precision = 4
      Size = 2
    end
    object qGrafikVCH31: TBCDField
      FieldName = 'VCH31'
      Precision = 4
      Size = 2
    end
    object qGrafikPCH1: TBCDField
      FieldName = 'PCH1'
      Precision = 4
      Size = 2
    end
    object qGrafikPCH2: TBCDField
      FieldName = 'PCH2'
      Precision = 4
      Size = 2
    end
    object qGrafikPCH3: TBCDField
      FieldName = 'PCH3'
      Precision = 4
      Size = 2
    end
    object qGrafikPCH4: TBCDField
      FieldName = 'PCH4'
      Precision = 4
      Size = 2
    end
    object qGrafikPCH5: TBCDField
      FieldName = 'PCH5'
      Precision = 4
      Size = 2
    end
    object qGrafikPCH6: TBCDField
      FieldName = 'PCH6'
      Precision = 4
      Size = 2
    end
    object qGrafikPCH7: TBCDField
      FieldName = 'PCH7'
      Precision = 4
      Size = 2
    end
    object qGrafikPCH8: TBCDField
      FieldName = 'PCH8'
      Precision = 4
      Size = 2
    end
    object qGrafikPCH9: TBCDField
      FieldName = 'PCH9'
      Precision = 4
      Size = 2
    end
    object qGrafikPCH10: TBCDField
      FieldName = 'PCH10'
      Precision = 4
      Size = 2
    end
    object qGrafikPCH11: TBCDField
      FieldName = 'PCH11'
      Precision = 4
      Size = 2
    end
    object qGrafikPCH12: TBCDField
      FieldName = 'PCH12'
      Precision = 4
      Size = 2
    end
    object qGrafikPCH13: TBCDField
      FieldName = 'PCH13'
      Precision = 4
      Size = 2
    end
    object qGrafikPCH14: TBCDField
      FieldName = 'PCH14'
      Precision = 4
      Size = 2
    end
    object qGrafikPCH15: TBCDField
      FieldName = 'PCH15'
      Precision = 4
      Size = 2
    end
    object qGrafikPCH16: TBCDField
      FieldName = 'PCH16'
      Precision = 4
      Size = 2
    end
    object qGrafikPCH17: TBCDField
      FieldName = 'PCH17'
      Precision = 4
      Size = 2
    end
    object qGrafikPCH18: TBCDField
      FieldName = 'PCH18'
      Precision = 4
      Size = 2
    end
    object qGrafikPCH19: TBCDField
      FieldName = 'PCH19'
      Precision = 4
      Size = 2
    end
    object qGrafikPCH20: TBCDField
      FieldName = 'PCH20'
      Precision = 4
      Size = 2
    end
    object qGrafikPCH21: TBCDField
      FieldName = 'PCH21'
      Precision = 4
      Size = 2
    end
    object qGrafikPCH22: TBCDField
      FieldName = 'PCH22'
      Precision = 4
      Size = 2
    end
    object qGrafikPCH23: TBCDField
      FieldName = 'PCH23'
      Precision = 4
      Size = 2
    end
    object qGrafikPCH24: TBCDField
      FieldName = 'PCH24'
      Precision = 4
      Size = 2
    end
    object qGrafikPCH25: TBCDField
      FieldName = 'PCH25'
      Precision = 4
      Size = 2
    end
    object qGrafikPCH26: TBCDField
      FieldName = 'PCH26'
      Precision = 4
      Size = 2
    end
    object qGrafikPCH27: TBCDField
      FieldName = 'PCH27'
      Precision = 4
      Size = 2
    end
    object qGrafikPCH28: TBCDField
      FieldName = 'PCH28'
      Precision = 4
      Size = 2
    end
    object qGrafikPCH29: TBCDField
      FieldName = 'PCH29'
      Precision = 4
      Size = 2
    end
    object qGrafikPCH30: TBCDField
      FieldName = 'PCH30'
      Precision = 4
      Size = 2
    end
    object qGrafikPCH31: TBCDField
      FieldName = 'PCH31'
      Precision = 4
      Size = 2
    end
    object qGrafikCHF1: TStringField
      FieldName = 'CHF1'
      Size = 40
    end
    object qGrafikCHF2: TStringField
      FieldName = 'CHF2'
      Size = 40
    end
    object qGrafikCHF3: TStringField
      FieldName = 'CHF3'
      Size = 40
    end
    object qGrafikCHF4: TStringField
      FieldName = 'CHF4'
      Size = 40
    end
    object qGrafikCHF5: TStringField
      FieldName = 'CHF5'
      Size = 40
    end
    object qGrafikCHF6: TStringField
      FieldName = 'CHF6'
      Size = 40
    end
    object qGrafikCHF7: TStringField
      FieldName = 'CHF7'
      Size = 40
    end
    object qGrafikCHF8: TStringField
      FieldName = 'CHF8'
      Size = 40
    end
    object qGrafikCHF9: TStringField
      FieldName = 'CHF9'
      Size = 40
    end
    object qGrafikCHF10: TStringField
      FieldName = 'CHF10'
      Size = 40
    end
    object qGrafikCHF11: TStringField
      FieldName = 'CHF11'
      Size = 40
    end
    object qGrafikCHF12: TStringField
      FieldName = 'CHF12'
      Size = 40
    end
    object qGrafikCHF13: TStringField
      FieldName = 'CHF13'
      Size = 40
    end
    object qGrafikCHF14: TStringField
      FieldName = 'CHF14'
      Size = 40
    end
    object qGrafikCHF15: TStringField
      FieldName = 'CHF15'
      Size = 40
    end
    object qGrafikCHF16: TStringField
      FieldName = 'CHF16'
      Size = 40
    end
    object qGrafikCHF17: TStringField
      FieldName = 'CHF17'
      Size = 40
    end
    object qGrafikCHF18: TStringField
      FieldName = 'CHF18'
      Size = 40
    end
    object qGrafikCHF19: TStringField
      FieldName = 'CHF19'
      Size = 40
    end
    object qGrafikCHF20: TStringField
      FieldName = 'CHF20'
      Size = 40
    end
    object qGrafikCHF21: TStringField
      FieldName = 'CHF21'
      Size = 40
    end
    object qGrafikCHF22: TStringField
      FieldName = 'CHF22'
      Size = 40
    end
    object qGrafikCHF23: TStringField
      FieldName = 'CHF23'
      Size = 40
    end
    object qGrafikCHF24: TStringField
      FieldName = 'CHF24'
      Size = 40
    end
    object qGrafikCHF25: TStringField
      FieldName = 'CHF25'
      Size = 40
    end
    object qGrafikCHF26: TStringField
      FieldName = 'CHF26'
      Size = 40
    end
    object qGrafikCHF27: TStringField
      FieldName = 'CHF27'
      Size = 40
    end
    object qGrafikCHF28: TStringField
      FieldName = 'CHF28'
      Size = 40
    end
    object qGrafikCHF29: TStringField
      FieldName = 'CHF29'
      Size = 40
    end
    object qGrafikCHF30: TStringField
      FieldName = 'CHF30'
      Size = 40
    end
    object qGrafikCHF31: TStringField
      FieldName = 'CHF31'
      Size = 40
    end
    object qGrafikCHF0: TBCDField
      FieldName = 'CHF0'
      Precision = 4
      Size = 2
    end
    object qGrafikNCH0: TBCDField
      FieldName = 'NCH0'
      Precision = 4
      Size = 2
    end
    object qGrafikPCH0: TBCDField
      FieldName = 'PCH0'
      Precision = 4
      Size = 2
    end
    object qGrafikGOD: TIntegerField
      FieldName = 'GOD'
    end
    object qGrafikDLIT: TBCDField
      FieldName = 'DLIT'
      Precision = 4
      Size = 2
    end
    object qGrafikOTCHET: TIntegerField
      FieldName = 'OTCHET'
    end
    object qGrafikBR: TIntegerField
      FieldName = 'BR'
    end
  end
  object dsGrafik: TDataSource
    DataSet = qGrafik
    Left = 110
    Top = 128
  end
  object qPrazdDni: TADOQuery
    Connection = ADOConnection1
    Parameters = <
      item
        Name = 'pgod'
        Size = -1
        Value = Null
      end>
    SQL.Strings = (
      'select god,mes,den from sp_prd where god=:pgod')
    Left = 38
    Top = 248
  end
  object qPrdPrazdDni: TADOQuery
    Connection = ADOConnection1
    Parameters = <
      item
        Name = 'pgod'
        Size = -1
        Value = Null
      end>
    SQL.Strings = (
      
        'select to_date(den||'#39'.'#39'||mes||'#39'.'#39'||god, '#39'dd.mm.yyyy'#39')-1 as data,' +
        ' '
      
        '       substr(to_char((to_date(den||'#39'.'#39'||mes||'#39'.'#39'||god, '#39'dd.mm.y' +
        'yyy'#39')-1),'#39'dd.mm.yyyy'#39'),1,2) as den,'
      
        '       substr(to_char((to_date(den||'#39'.'#39'||mes||'#39'.'#39'||god, '#39'dd.mm.y' +
        'yyy'#39')-1),'#39'dd.mm.yyyy'#39'),4,2) as mes,'
      
        '       substr(to_char((to_date(den||'#39'.'#39'||mes||'#39'.'#39'||god, '#39'dd.mm.y' +
        'yyy'#39')-1),'#39'dd.mm.yyyy'#39'),7,4) as god'
      'from sp_prd where god=:pgod'
      'order by mes')
    Left = 38
    Top = 320
  end
  object qPrazdDniVihodnue: TADOQuery
    Connection = ADOConnection1
    Parameters = <>
    SQL.Strings = (
      
        'select substr(dat, 1,2) as den, substr(dat,4,2) as mes, substr(t' +
        'o_char(dat, '#39'dd.mm.yyyy'#39'),7,4) as god from ('
      
        'select data+1 as dat from (select  to_date(den||'#39'.'#39'||mes||'#39'.'#39'||g' +
        'od, '#39'dd.mm.yyyy'#39') as data'
      '               from sp_prd) '
      'where to_char(data,'#39'D'#39')=7 '
      'union'
      
        'select data+2 as dat from (select  to_date(den||'#39'.'#39'||mes||'#39'.'#39'||g' +
        'od, '#39'dd.mm.yyyy'#39') as data'
      '               from sp_prd) '
      'where to_char(data,'#39'D'#39')=6'
      'union all'
      
        'select data+1 as dat from (select  to_date(den||'#39'.'#39'||mes||'#39'.'#39'||g' +
        'od, '#39'dd.mm.yyyy'#39') as data'
      '               from sp_prd) '
      
        'where (to_char(data-1,'#39'D'#39')=7 and data-1 in (select  to_date(den|' +
        '|'#39'.'#39'||mes||'#39'.'#39'||god, '#39'dd.mm.yyyy'#39') as data'
      '               from sp_prd) )and to_char(data,'#39'D'#39')=1 '
      ''
      ')')
    Left = 38
    Top = 384
  end
  object qOgraf: TADOQuery
    Connection = ADOConnection1
    Parameters = <
      item
        Name = 'pograf'
        Size = -1
        Value = Null
      end>
    SQL.Strings = (
      'select * from spograf where ograf=:pograf')
    Left = 40
    Top = 192
  end
  object qObnovlenie2: TADOQuery
    Connection = ADOConnection1
    Parameters = <>
    Left = 112
    Top = 72
  end
  object qNorma11Graf: TADOQuery
    Connection = ADOConnection1
    Parameters = <
      item
        Name = 'pograf'
        Size = -1
        Value = Null
      end
      item
        Name = 'pgod'
        Size = -1
        Value = Null
      end>
    SQL.Strings = (
      'select mes,'
      '       chf,'
      '       sum(chf) over () as onorma'
      'from spgrafiki where ograf=:pograf and god=:pgod'
      'order by mes')
    Left = 40
    Top = 440
  end
  object qSprav: TADOQuery
    Connection = ADOConnection1
    CursorType = ctStatic
    Parameters = <
      item
        Name = 'pgod'
        Size = -1
        Value = Null
      end
      item
        Name = 'pgod1'
        Size = -1
        Value = Null
      end>
    SQL.Strings = (
      
        'select rowidtochar(rowid) rw,god,mes,den from sp_prd where god=:' +
        'pgod or god=:pgod1'
      'order by god desc, mes, den')
    Left = 40
    Top = 547
  end
  object dsSprav: TDataSource
    DataSet = qSprav
    Left = 112
    Top = 547
  end
end
