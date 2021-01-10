object FChangeNagryzka: TFChangeNagryzka
  Left = 0
  Top = 0
  Caption = #1048#1079#1084#1077#1085#1077#1085#1080#1077' '#1085#1072#1075#1088#1091#1079#1082#1080
  ClientHeight = 679
  ClientWidth = 1109
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  OnActivate = FormActivate
  PixelsPerInch = 96
  TextHeight = 13
  object Panel1: TPanel
    Left = 0
    Top = 0
    Width = 185
    Height = 679
    Align = alLeft
    TabOrder = 0
    object Label1: TLabel
      Left = 8
      Top = 8
      Width = 61
      Height = 13
      Caption = #1044#1080#1089#1094#1080#1087#1083#1080#1085#1072
    end
    object LADis: TLabel
      Left = 8
      Top = 24
      Width = 3
      Height = 13
    end
    object Label3: TLabel
      Left = 8
      Top = 40
      Width = 68
      Height = 13
      Caption = #1042#1080#1076' '#1085#1072#1075#1088#1091#1079#1082#1080
    end
    object LaVid: TLabel
      Left = 8
      Top = 59
      Width = 3
      Height = 13
    end
    object Label2: TLabel
      Left = 8
      Top = 78
      Width = 36
      Height = 13
      Caption = #1043#1088#1091#1087#1087#1072
    end
    object Label4: TLabel
      Left = 8
      Top = 118
      Width = 26
      Height = 13
      Caption = #1063#1072#1089#1099
    end
    object Label5: TLabel
      Left = 8
      Top = 158
      Width = 80
      Height = 13
      Caption = #1055#1088#1077#1087#1086#1076#1072#1074#1072#1090#1077#1083#1100
    end
    object Label6: TLabel
      Left = 8
      Top = 198
      Width = 61
      Height = 13
      Caption = #1050#1086#1084#1077#1085#1090#1072#1088#1080#1080
    end
    object LaGroup: TLabel
      Left = 8
      Top = 97
      Width = 3
      Height = 13
    end
    object LaHour: TLabel
      Left = 8
      Top = 137
      Width = 3
      Height = 13
    end
    object LaFioPrep: TLabel
      Left = 8
      Top = 179
      Width = 3
      Height = 13
    end
    object LaComment: TLabel
      Left = 8
      Top = 217
      Width = 3
      Height = 13
    end
    object Label7: TLabel
      Left = 8
      Top = 236
      Width = 80
      Height = 13
      Caption = #1053#1086#1084#1077#1088' '#1085#1072#1075#1088#1091#1079#1082#1080
    end
    object LaNomNagryzka: TLabel
      Left = 8
      Top = 255
      Width = 3
      Height = 13
    end
    object Label8: TLabel
      Left = 88
      Top = 118
      Width = 90
      Height = 13
      Caption = #1063#1077#1083#1086#1074#1077#1082' '#1074' '#1075#1088#1091#1087#1087#1077
    end
    object LaStudent: TLabel
      Left = 88
      Top = 137
      Width = 3
      Height = 13
    end
  end
  object Panel2: TPanel
    Left = 185
    Top = 0
    Width = 240
    Height = 679
    Align = alLeft
    TabOrder = 1
    object LbPrep: TListBox
      Left = 1
      Top = 37
      Width = 238
      Height = 641
      Align = alClient
      ItemHeight = 13
      Sorted = True
      TabOrder = 0
    end
    object Panel4: TPanel
      Left = 1
      Top = 1
      Width = 238
      Height = 36
      Align = alTop
      Caption = #1055#1088#1077#1087#1086#1076#1072#1074#1072#1090#1077#1083#1080
      TabOrder = 1
    end
  end
  object Panel3: TPanel
    Left = 425
    Top = 0
    Width = 684
    Height = 679
    Align = alClient
    TabOrder = 2
    object Panel5: TPanel
      Left = 1
      Top = 37
      Width = 24
      Height = 641
      Align = alLeft
      TabOrder = 0
      object BtGoPrepodTable: TButton
        Left = -3
        Top = 83
        Width = 28
        Height = 25
        Caption = '>>>'
        TabOrder = 0
        OnClick = BtGoPrepodTableClick
      end
    end
    object Panel6: TPanel
      Left = 1
      Top = 1
      Width = 682
      Height = 36
      Align = alTop
      TabOrder = 1
      object BtNazn: TButton
        Left = 30
        Top = 3
        Width = 89
        Height = 29
        Caption = #1053#1072#1079#1085#1072#1095#1080#1090#1100
        TabOrder = 0
        OnClick = BtNaznClick
      end
    end
    object SgNewNagryzkaPrepod: TStringGrid
      Left = 25
      Top = 37
      Width = 658
      Height = 641
      Align = alClient
      ColCount = 3
      FixedRows = 4
      Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRangeSelect, goEditing]
      TabOrder = 2
      OnSetEditText = SgNewNagryzkaPrepodSetEditText
    end
  end
end
