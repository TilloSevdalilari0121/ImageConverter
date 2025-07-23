object SettingsForm: TSettingsForm
  Left = 0
  Top = 0
  Caption = 'Ayarlar'
  ClientHeight = 250
  ClientWidth = 350
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  Position = poScreenCenter
  OnCreate = FormCreate
  TextHeight = 13
  object GroupBox1: TGroupBox
    Left = 16
    Top = 16
    Width = 318
    Height = 137
    Caption = 'Dosya Adland'#305'rma'
    TabOrder = 0
    object RadioButton1: TRadioButton
      Left = 16
      Top = 24
      Width = 289
      Height = 17
      Caption = 'eski_isim.yeni_format'
      Checked = True
      TabOrder = 0
      TabStop = True
    end
    object RadioButton2: TRadioButton
      Left = 16
      Top = 56
      Width = 289
      Height = 17
      Caption = 'eski_isim_converted.yeni_format'
      TabOrder = 1
    end
    object RadioButton3: TRadioButton
      Left = 16
      Top = 88
      Width = 289
      Height = 17
      Caption = 'eski_isim_format.yeni_format'
      TabOrder = 2
    end
  end
  object Button1: TButton
    Left = 96
    Top = 176
    Width = 75
    Height = 25
    Caption = 'Kaydet'
    TabOrder = 1
    OnClick = Button1Click
  end
  object Button2: TButton
    Left = 192
    Top = 176
    Width = 75
    Height = 25
    Caption = #304'ptal'
    TabOrder = 2
    OnClick = Button2Click
  end
  object Button3: TButton
    Left = 96
    Top = 217
    Width = 171
    Height = 25
    Caption = 'Program'#305' Kald'#305'r'
    TabOrder = 3
    OnClick = Button3Click
  end
end
