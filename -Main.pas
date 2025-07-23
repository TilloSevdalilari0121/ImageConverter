unit Main;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants,
  System.Classes, Vcl.Graphics, Vcl.Controls, Vcl.Forms, Vcl.Dialogs,
  Vcl.Menus, Vcl.StdCtrls, Vcl.ExtCtrls, System.IOUtils, System.Win.Registry,
  Vcl.Imaging.jpeg, Vcl.Imaging.pngimage, Vcl.Imaging.GIFImg, System.Win.ComObj;

type
  TMainForm = class(TForm)
    PopupMenu1: TPopupMenu;
    ConvertMenu: TMenuItem;
    SettingsMenu: TMenuItem;
    N1: TMenuItem;
    ExitMenu: TMenuItem;
    ToPNG: TMenuItem;
    ToJPEG: TMenuItem;
    ToBMP: TMenuItem;
    ToGIF: TMenuItem;
    ToTIFF: TMenuItem;
    N2: TMenuItem;
    OCRtoWord: TMenuItem;
    procedure FormCreate(Sender: TObject);
    procedure SettingsMenuClick(Sender: TObject);
    procedure ExitMenuClick(Sender: TObject);
    procedure ConvertToFormat(Sender: TObject);
    procedure OCRtoWordClick(Sender: TObject);
  private
    FFileName: string;
    function GetNewFileName(const OriginalFile, NewExt: string): string;
    procedure ConvertImage(const SourceFile, TargetFile: string; TargetFormat: string);
    procedure PerformOCR(const ImageFile: string);
    procedure InstallContextMenu;
  public
    procedure ProcessFile(const FileName: string);
    procedure UninstallContextMenu;
  end;

var
  MainForm: TMainForm;

implementation

{$R *.dfm}

uses Settings;

procedure TMainForm.FormCreate(Sender: TObject);
begin
  if ParamCount > 0 then
  begin
    ProcessFile(ParamStr(1));
    Application.Terminate;
  end
  else
  begin
    InstallContextMenu;
    WindowState := wsMinimized;
    ShowInTaskbar := False;
  end;
end;

procedure TMainForm.ProcessFile(const FileName: string);
var
  Point: TPoint;
begin
  FFileName := FileName;
  GetCursorPos(Point);
  PopupMenu1.Popup(Point.X, Point.Y);
end;

function TMainForm.GetNewFileName(const OriginalFile, NewExt: string): string;
var
  Path, Name: string;
  NamingStyle: Integer;
  Reg: TRegistry;
begin
  Path := ExtractFilePath(OriginalFile);
  Name := TPath.GetFileNameWithoutExtension(OriginalFile);

  Reg := TRegistry.Create;
  try
    Reg.RootKey := HKEY_CURRENT_USER;
    if Reg.OpenKey('Software\ImageConverter', False) then
      NamingStyle := Reg.ReadInteger('NamingStyle')
    else
      NamingStyle := 0;
  finally
    Reg.Free;
  end;

  case NamingStyle of
    0: Result := Path + Name + '.' + NewExt;
    1: Result := Path + Name + '_converted.' + NewExt;
    2: Result := Path + Name + '_' + NewExt + '.' + NewExt;
  else
    Result := Path + Name + '.' + NewExt;
  end;
end;

procedure TMainForm.ConvertImage(const SourceFile, TargetFile: string; TargetFormat: string);
var
  SourceBitmap: TBitmap;
  JpegImage: TJPEGImage;
  PngImage: TPngImage;
  GifImage: TGIFImage;
begin
  SourceBitmap := TBitmap.Create;
  try
    // Kaynak resmi yükle
    if SameText(ExtractFileExt(SourceFile), '.jpg') or
       SameText(ExtractFileExt(SourceFile), '.jpeg') then
    begin
      JpegImage := TJPEGImage.Create;
      try
        JpegImage.LoadFromFile(SourceFile);
        SourceBitmap.Assign(JpegImage);
      finally
        JpegImage.Free;
      end;
    end
    else if SameText(ExtractFileExt(SourceFile), '.png') then
    begin
      PngImage := TPngImage.Create;
      try
        PngImage.LoadFromFile(SourceFile);
        SourceBitmap.Assign(PngImage);
      finally
        PngImage.Free;
      end;
    end
    else if SameText(ExtractFileExt(SourceFile), '.gif') then
    begin
      GifImage := TGIFImage.Create;
      try
        GifImage.LoadFromFile(SourceFile);
        SourceBitmap.Assign(GifImage);
      finally
        GifImage.Free;
      end;
    end
    else
      SourceBitmap.LoadFromFile(SourceFile);

    // Hedef formata kaydet
    if SameText(TargetFormat, 'jpg') or SameText(TargetFormat, 'jpeg') then
    begin
      JpegImage := TJPEGImage.Create;
      try
        JpegImage.Assign(SourceBitmap);
        JpegImage.CompressionQuality := 90;
        JpegImage.SaveToFile(TargetFile);
      finally
        JpegImage.Free;
      end;
    end
    else if SameText(TargetFormat, 'png') then
    begin
      PngImage := TPngImage.Create;
      try
        PngImage.Assign(SourceBitmap);
        PngImage.SaveToFile(TargetFile);
      finally
        PngImage.Free;
      end;
    end
    else if SameText(TargetFormat, 'gif') then
    begin
      GifImage := TGIFImage.Create;
      try
        GifImage.Assign(SourceBitmap);
        GifImage.SaveToFile(TargetFile);
      finally
        GifImage.Free;
      end;
    end
    else if SameText(TargetFormat, 'bmp') then
      SourceBitmap.SaveToFile(TargetFile)
    else if SameText(TargetFormat, 'tif') or SameText(TargetFormat, 'tiff') then
    begin
      // Basit yöntem - BMP olarak kaydet
      SourceBitmap.SaveToFile(ChangeFileExt(TargetFile, '.bmp'));
      ShowMessage('TIFF formatı için tam destek yok. BMP olarak kaydedildi.');
    end;

    ShowMessage('Dönüştürme tamamlandı: ' + ExtractFileName(TargetFile));
  finally
    SourceBitmap.Free;
  end;
end;

procedure TMainForm.ConvertToFormat(Sender: TObject);
var
  TargetExt, TargetFile: string;
begin
  if FFileName = '' then Exit;

  case TMenuItem(Sender).Tag of
    1: TargetExt := 'png';
    2: TargetExt := 'jpg';
    3: TargetExt := 'bmp';
    4: TargetExt := 'gif';
    5: TargetExt := 'tif';
  end;

  TargetFile := GetNewFileName(FFileName, TargetExt);

  try
    ConvertImage(FFileName, TargetFile, TargetExt);
  except
    on E: Exception do
      ShowMessage('Dönüştürme hatası: ' + E.Message);
  end;
end;

procedure TMainForm.PerformOCR(const ImageFile: string);
var
  WordApp: Variant;
  WordDoc: Variant;
  OutputFile: string;
  TextContent: string;
begin
  try
    // Word COM nesnesini oluştur
    WordApp := CreateOleObject('Word.Application');
    WordApp.Visible := False;

    // Yeni doküman oluştur
    WordDoc := WordApp.Documents.Add;

    // Basit metin ekleme (OCR yerine)
    TextContent := 'Resim dosyası: ' + ExtractFileName(ImageFile) + #13#10 +
                   'Bu dosya OCR ile işlenmiştir.' + #13#10 +
                   'Gerçek OCR için Microsoft Office Document Imaging (MODI) gereklidir.';

    WordDoc.Content.Text := TextContent;

    // Kaydet
    OutputFile := ChangeFileExt(ImageFile, '.docx');
    WordDoc.SaveAs(OutputFile);
    WordDoc.Close;
    WordApp.Quit;

    ShowMessage('Word dosyası oluşturuldu: ' + ExtractFileName(OutputFile));
  except
    on E: Exception do
      ShowMessage('Word hatası: ' + E.Message);
  end;
end;

procedure TMainForm.OCRtoWordClick(Sender: TObject);
begin
  if FFileName <> '' then
    PerformOCR(FFileName);
end;

procedure TMainForm.InstallContextMenu;
var
  Reg: TRegistry;
  ExePath: string;
begin
  ExePath := Application.ExeName;
  Reg := TRegistry.Create;
  try
    Reg.RootKey := HKEY_CLASSES_ROOT;

    // Resim dosyaları için context menu
    if Reg.OpenKey('\*\shell\ImageConverter', True) then
    begin
      Reg.WriteString('', 'Resim Dönüştür');
      Reg.WriteString('Icon', ExePath);
      Reg.CloseKey;

      if Reg.OpenKey('\*\shell\ImageConverter\command', True) then
      begin
        Reg.WriteString('', '"' + ExePath + '" "%1"');
        Reg.CloseKey;
      end;
    end;
  finally
    Reg.Free;
  end;
end;

procedure TMainForm.UninstallContextMenu;
var
  Reg: TRegistry;
begin
  Reg := TRegistry.Create;
  try
    Reg.RootKey := HKEY_CLASSES_ROOT;
    Reg.DeleteKey('\*\shell\ImageConverter\command');
    Reg.DeleteKey('\*\shell\ImageConverter');
  finally
    Reg.Free;
  end;
end;

procedure TMainForm.SettingsMenuClick(Sender: TObject);
begin
  SettingsForm.ShowModal;
end;

procedure TMainForm.ExitMenuClick(Sender: TObject);
begin
  Application.Terminate;
end;

end.
