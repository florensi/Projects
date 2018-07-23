object Data: TData
  Left = 315
  Top = 120
  BorderStyle = bsSingle
  Caption = #1042#1099#1073#1086#1088' '#1086#1090#1095#1077#1090#1085#1086#1075#1086' '#1087#1077#1088#1080#1086#1076#1072
  ClientHeight = 181
  ClientWidth = 331
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  KeyPreview = True
  OldCreateOrder = False
  Position = poScreenCenter
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object Panel1: TPanel
    Left = 0
    Top = 0
    Width = 331
    Height = 181
    Align = alClient
    Color = clSilver
    TabOrder = 0
    object Bevel1: TBevel
      Left = 32
      Top = 16
      Width = 281
      Height = 105
    end
    object Label1: TLabel
      Left = 42
      Top = 18
      Width = 253
      Height = 48
      Alignment = taCenter
      AutoSize = False
      Caption = #1055#1088#1086#1089#1084#1086#1090#1088'  '#1080#1085#1092#1086#1088#1084#1072#1094#1080#1080' '#1087#1086' '#1086#1094#1077#1085#1082#1077' '#1087#1077#1088#1089#1086#1085#1072#1083#1072' '#1079#1072' '
      Font.Charset = RUSSIAN_CHARSET
      Font.Color = clRed
      Font.Height = -19
      Font.Name = 'Calibri'
      Font.Style = [fsBold]
      ParentFont = False
      WordWrap = True
    end
    object btnVibor: TBitBtn
      Left = 70
      Top = 134
      Width = 97
      Height = 33
      Caption = #1042#1074#1086#1076
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -16
      Font.Name = 'Calibri'
      Font.Style = [fsBold]
      ModalResult = 1
      ParentFont = False
      TabOrder = 0
      OnClick = btnViborClick
      OnKeyDown = btnViborKeyDown
    end
    object BitBtn2: TBitBtn
      Left = 174
      Top = 134
      Width = 97
      Height = 33
      Caption = #1054#1090#1084#1077#1085#1072
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -16
      Font.Name = 'Calibri'
      Font.Style = [fsBold]
      ModalResult = 2
      ParentFont = False
      TabOrder = 1
    end
    object DateTimePicker1: TDateTimePicker
      Left = 96
      Top = 72
      Width = 145
      Height = 37
      BevelOuter = bvRaised
      CalAlignment = dtaLeft
      Date = 0.375245844901656
      Format = '        yyyy'
      Time = 0.375245844901656
      DateFormat = dfLong
      DateMode = dmUpDown
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -24
      Font.Name = 'Calibri'
      Font.Style = [fsBold]
      Kind = dtkDate
      ParseInput = False
      ParentFont = False
      TabOrder = 2
    end
  end
end
