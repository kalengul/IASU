object FMain: TFMain
  Left = 0
  Top = 0
  Caption = #1055#1088#1086#1075#1088#1072#1084#1084#1072' '#1088#1072#1073#1086#1090#1099' '#1089' XLS '#1092#1072#1081#1083#1086#1084' '#1048#1040#1057#1059
  ClientHeight = 675
  ClientWidth = 1280
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  OnActivate = FormActivate
  OnClose = FormClose
  PixelsPerInch = 96
  TextHeight = 13
  object Panel1: TPanel
    Left = 0
    Top = 0
    Width = 1280
    Height = 66
    Align = alTop
    TabOrder = 0
    object Label1: TLabel
      Left = 143
      Top = 8
      Width = 163
      Height = 16
      Caption = #1042#1099#1073#1088#1072#1085#1085#1099#1081' '#1092#1072#1081#1083' XLS '#1048#1040#1057#1059
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -13
      Font.Name = 'Tahoma'
      Font.Style = []
      ParentFont = False
    end
    object LaNameFile: TLabel
      Left = 309
      Top = 8
      Width = 4
      Height = 16
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -13
      Font.Name = 'Tahoma'
      Font.Style = []
      ParentFont = False
    end
    object BtIASY: TButton
      Left = 8
      Top = 5
      Width = 129
      Height = 25
      Caption = #1042#1099#1073#1088#1072#1090#1100' '#1085#1086#1074#1099#1081' '#1092#1072#1081#1083
      TabOrder = 0
      OnClick = BtIASYClick
    end
    object BtProverka: TButton
      Left = 8
      Top = 35
      Width = 97
      Height = 25
      Caption = #1055#1088#1086#1074#1077#1088#1080#1090#1100
      TabOrder = 1
      OnClick = BtProverkaClick
    end
    object BtCreateResult: TButton
      Left = 374
      Top = 35
      Width = 75
      Height = 25
      Caption = #1048#1085#1076'. '#1087#1083#1072#1085#1099
      TabOrder = 2
      OnClick = BtCreateResultClick
    end
    object CbAutoProverkaClose: TCheckBox
      Left = 109
      Top = 39
      Width = 242
      Height = 17
      Caption = #1040#1074#1090#1086#1084#1072#1090#1080#1095#1077#1089#1082#1072#1103' '#1087#1088#1086#1074#1077#1088#1082#1072' '#1087#1088#1080' '#1074#1099#1093#1086#1076#1077
      TabOrder = 3
    end
    object BtExzToExcel: TButton
      Left = 455
      Top = 36
      Width = 72
      Height = 25
      Caption = #1069#1082#1079#1072#1084#1077#1085#1099
      TabOrder = 4
      OnClick = BtExzToExcelClick
    end
    object BtRaspisanie: TButton
      Left = 533
      Top = 36
      Width = 75
      Height = 25
      Caption = #1056#1072#1089#1087#1080#1089#1072#1085#1080#1077
      TabOrder = 5
      OnClick = BtRaspisanieClick
    end
    object BtMTO: TButton
      Left = 688
      Top = 35
      Width = 89
      Height = 25
      Caption = #1052#1058#1054
      TabOrder = 6
      OnClick = BtMTOClick
    end
    object BtGroup: TButton
      Left = 688
      Top = 4
      Width = 89
      Height = 25
      Caption = #1057#1086#1079#1076#1072#1090#1100#1075#1088#1091#1087#1087#1099' '
      TabOrder = 7
      OnClick = BtGroupClick
    end
    object BtVivodPrepVGr: TButton
      Left = 870
      Top = 4
      Width = 81
      Height = 25
      Caption = #1055#1088#1077#1087'. '#1074' '#1075#1088'.'
      TabOrder = 8
      OnClick = BtVivodPrepVGrClick
    end
    object BtPrepodInOOP: TButton
      Left = 870
      Top = 36
      Width = 75
      Height = 25
      Caption = #1055#1088#1077#1087' '#1054#1054#1055
      TabOrder = 9
      OnClick = BtPrepodInOOPClick
    end
    object Button2: TButton
      Left = 951
      Top = 35
      Width = 75
      Height = 25
      Caption = #1047#1072#1075#1088#1091#1079#1082#1072' '#1059#1055
      TabOrder = 10
      OnClick = Button2Click
    end
    object BtMTOSemestrov: TButton
      Left = 783
      Top = 35
      Width = 66
      Height = 25
      Caption = #1052#1058#1054' '#1057#1077#1084
      TabOrder = 11
      OnClick = BtMTOSemestrovClick
    end
    object BtAllPOAud: TButton
      Left = 1032
      Top = 36
      Width = 75
      Height = 25
      Caption = 'BtAllPOAud'
      TabOrder = 12
      OnClick = BtAllPOAudClick
    end
    object Button1: TButton
      Left = 784
      Top = 4
      Width = 65
      Height = 25
      Caption = #1052#1058#1054' '#1087#1086' '#1076#1080#1089
      TabOrder = 13
      OnClick = Button1Click
    end
    object BtCreateYMK: TButton
      Left = 957
      Top = 5
      Width = 75
      Height = 25
      Caption = #1059#1052#1050
      TabOrder = 14
      OnClick = BtCreateYMKClick
    end
    object BtTablePredmet: TButton
      Left = 1038
      Top = 5
      Width = 75
      Height = 25
      Caption = #1058#1072#1073' '#1087#1088#1077#1076#1084#1077#1090
      TabOrder = 15
      OnClick = BtTablePredmetClick
    end
    object BtExzGr: TButton
      Left = 592
      Top = 4
      Width = 72
      Height = 25
      Caption = #1069#1050#1047#1040#1052#1045#1053#1067' '#1043#1088
      TabOrder = 16
      OnClick = BtExzGrClick
    end
  end
  object Panel2: TPanel
    Left = 113
    Top = 66
    Width = 393
    Height = 364
    Align = alLeft
    TabOrder = 1
    object Panel5: TPanel
      Left = 1
      Top = 1
      Width = 391
      Height = 24
      Align = alTop
      TabOrder = 0
      object LaKaf: TLabel
        Left = 5
        Top = 5
        Width = 27
        Height = 13
        Caption = 'LaKaf'
      end
      object LaZavKaf: TLabel
        Left = 88
        Top = 5
        Width = 45
        Height = 13
        Caption = 'LaZavKaf'
      end
    end
    object Panel6: TPanel
      Left = 1
      Top = 326
      Width = 391
      Height = 37
      Align = alBottom
      TabOrder = 1
    end
    object SGPrepod: TStringGrid
      Left = 1
      Top = 25
      Width = 391
      Height = 301
      Align = alClient
      Color = clHighlight
      ColCount = 3
      TabOrder = 2
    end
  end
  object Panel3: TPanel
    Left = 1240
    Top = 66
    Width = 40
    Height = 364
    Align = alRight
    TabOrder = 2
  end
  object Panel4: TPanel
    Left = 506
    Top = 66
    Width = 734
    Height = 364
    Align = alClient
    TabOrder = 3
    object SgNagryzka: TStringGrid
      Left = 1
      Top = 113
      Width = 732
      Height = 250
      Align = alClient
      FixedCols = 0
      Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRangeSelect, goEditing]
      TabOrder = 0
      OnSelectCell = SgNagryzkaSelectCell
      OnSetEditText = SgNagryzkaSetEditText
    end
    object Panel10: TPanel
      Left = 1
      Top = 1
      Width = 732
      Height = 112
      Align = alTop
      TabOrder = 1
      object BtSaveSgNagryzkaXLSX: TButton
        Left = 5
        Top = 6
        Width = 132
        Height = 25
        Caption = #1057#1086#1093#1088#1072#1085#1080#1090#1100' '#1090#1072#1073#1083#1080#1094#1091' XLSX'
        TabOrder = 0
        OnClick = BtSaveSgNagryzkaXLSXClick
      end
      object BtRaspredelenieNagryzki: TButton
        Left = 152
        Top = 6
        Width = 89
        Height = 25
        Caption = #1056#1072#1089#1087#1088#1077#1076#1077#1083#1080#1090#1100
        TabOrder = 1
        OnClick = BtRaspredelenieNagryzkiClick
      end
      object SgNagryzkaSearth: TStringGrid
        Left = 1
        Top = 37
        Width = 730
        Height = 74
        Align = alBottom
        FixedCols = 0
        RowCount = 2
        Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRangeSelect, goEditing]
        TabOrder = 2
        OnSetEditText = SgNagryzkaSearthSetEditText
        ColWidths = (
          64
          64
          64
          64
          64)
      end
    end
  end
  object Panel7: TPanel
    Left = 0
    Top = 430
    Width = 1280
    Height = 245
    Align = alBottom
    TabOrder = 4
    object MeProtocol: TMemo
      Left = 1
      Top = 1
      Width = 1278
      Height = 243
      Align = alClient
      ScrollBars = ssBoth
      TabOrder = 0
    end
  end
  object Panel11: TPanel
    Left = 0
    Top = 66
    Width = 113
    Height = 364
    Align = alLeft
    TabOrder = 5
    object PNagrO: TPanel
      Left = 8
      Top = 6
      Width = 97
      Height = 26
      Caption = #1053#1072#1075#1088#1091#1079#1082#1072' '#1086#1089#1077#1085#1100
      Color = 1841306
      ParentBackground = False
      TabOrder = 0
    end
    object PGroup: TPanel
      Left = 8
      Top = 107
      Width = 97
      Height = 26
      Caption = #1043#1088#1091#1087#1087#1099
      Color = 1841306
      ParentBackground = False
      TabOrder = 1
    end
    object POsn: TPanel
      Left = 8
      Top = 269
      Width = 97
      Height = 26
      Caption = #1054#1089#1085#1072#1097#1077#1085#1080#1077' '#1072#1091#1076'.'
      Color = 1841306
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'Tahoma'
      Font.Style = []
      ParentBackground = False
      ParentFont = False
      TabOrder = 2
    end
    object PPrepod: TPanel
      Left = 8
      Top = 139
      Width = 97
      Height = 26
      Caption = #1055#1088#1077#1087#1086#1076#1072#1074#1072#1090#1077#1083#1080
      Color = 1841306
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'Tahoma'
      Font.Style = []
      ParentBackground = False
      ParentFont = False
      TabOrder = 3
    end
    object PEkz: TPanel
      Left = 8
      Top = 171
      Width = 97
      Height = 26
      Caption = #1069#1082#1079#1072#1084#1077#1085#1099
      Color = 1841306
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'Tahoma'
      Font.Style = []
      ParentBackground = False
      ParentFont = False
      TabOrder = 4
    end
    object PNagrV: TPanel
      Left = 8
      Top = 38
      Width = 97
      Height = 26
      Caption = #1053#1072#1075#1088#1091#1079#1082#1072' '#1074#1077#1089#1085#1072
      Color = 1841306
      ParentBackground = False
      TabOrder = 5
    end
    object PRaspPrepod: TPanel
      Left = 8
      Top = 205
      Width = 97
      Height = 26
      Caption = #1056#1072#1089#1087'. '#1087#1088#1077#1087#1086#1076'.'
      Color = 1841306
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'Tahoma'
      Font.Style = []
      ParentBackground = False
      ParentFont = False
      TabOrder = 6
    end
    object PPO: TPanel
      Left = 8
      Top = 302
      Width = 97
      Height = 26
      Caption = #1055#1054
      Color = 1841306
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'Tahoma'
      Font.Style = []
      ParentBackground = False
      ParentFont = False
      TabOrder = 7
    end
    object PRaspGroup: TPanel
      Left = 8
      Top = 237
      Width = 97
      Height = 26
      Caption = #1056#1072#1089#1087'. '#1075#1088#1091#1087#1087
      Color = 1841306
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'Tahoma'
      Font.Style = []
      ParentBackground = False
      ParentFont = False
      TabOrder = 8
    end
    object PRPD: TPanel
      Left = 8
      Top = 334
      Width = 97
      Height = 26
      Caption = #1055#1054' '#1074' '#1056#1055#1044
      Color = 1841306
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'Tahoma'
      Font.Style = []
      ParentBackground = False
      ParentFont = False
      TabOrder = 9
    end
    object PMergeDis: TPanel
      Left = 8
      Top = 72
      Width = 97
      Height = 26
      Caption = #1054#1073#1098#1077#1076#1080#1085#1077#1085#1080#1077' '#1076#1080#1089
      Color = 1841306
      ParentBackground = False
      TabOrder = 10
    end
  end
  object ODIASY: TOpenDialog
    Left = 328
  end
end
