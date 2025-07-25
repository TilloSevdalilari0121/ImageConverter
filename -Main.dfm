object MainForm: TMainForm
  Left = 0
  Top = 0
  Caption = 'Image Converter'
  ClientHeight = 100
  ClientWidth = 200
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OnCreate = FormCreate
  TextHeight = 13
  object PopupMenu1: TPopupMenu
    Left = 96
    Top = 48
    object ConvertMenu: TMenuItem
      Caption = 'D'#246'n'#252#351't'#252'r'
      object ToPNG: TMenuItem
        Tag = 1
        Caption = 'PNG'
        OnClick = ConvertToFormat
      end
      object ToJPEG: TMenuItem
        Tag = 2
        Caption = 'JPEG'
        OnClick = ConvertToFormat
      end
      object ToBMP: TMenuItem
        Tag = 3
        Caption = 'BMP'
        OnClick = ConvertToFormat
      end
      object ToGIF: TMenuItem
        Tag = 4
        Caption = 'GIF'
        OnClick = ConvertToFormat
      end
      object ToTIFF: TMenuItem
        Tag = 5
        Caption = 'TIFF'
        OnClick = ConvertToFormat
      end
    end
    object N2: TMenuItem
      Caption = '-'
    end
    object OCRtoWord: TMenuItem
      Caption = 'OCR -> Word'
      OnClick = OCRtoWordClick
    end
    object N1: TMenuItem
      Caption = '-'
    end
    object SettingsMenu: TMenuItem
      Caption = 'Ayarlar'
      OnClick = SettingsMenuClick
    end
    object ExitMenu: TMenuItem
      Caption = #199#305'k'#305#351
      OnClick = ExitMenuClick
    end
  end
end
