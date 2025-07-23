unit Main;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants,
  System.Classes, Vcl.Graphics, Vcl.Controls, Vcl.Forms, Vcl.Dialogs,
  Vcl.Menus, Vcl.StdCtrls, Vcl.ExtCtrls, System.IOUtils, System.Win.Registry,
  Vcl.Imaging.jpeg, Vcl.Imaging.pngimage, Vcl.Imaging.GIFImg;

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
    procedure FormCreate(Sender: TObject);
    procedure SettingsMenuClick(Sender: TObject);
    procedure ExitMenuClick(Sender: TObject);
    procedure ConvertToFormat(Sender: TObject);
  private
    FFileName: string;
    FConvertMode: Boolean;
    function IsImageFile(const FileName: string): Boolean;
    function GetNewFileName(const OriginalFile, NewExt: string): string;
    procedure ConvertImage(const SourceFile, TargetFile: string; TargetFormat: string);
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

function TMainForm.IsImageFile(const FileName: string): Boolean;
var
  Ext: string;
begin
  Ext := LowerCase(ExtractFileExt(FileName));
  Result := (Ext = '.jpg') or (Ext = '.jpeg') or
            (Ext = '.png') or (Ext = '.bmp') or
            (Ext = '.gif') or (Ext = '.tif') or
            (Ext = '.tiff');
end;

procedure TMainForm.FormCreate(Sender: TObject);
begin
 { FConvertMode := False;

  if ParamCount > 0 then
  begin
    FConvertMode := True;
    ProcessFile(ParamStr(1));
  end
  else
  begin
    InstallContextMenu;
    WindowState := wsNormal;
    ShowInTaskbar := True;
  end;  }

  ProcessFile(ParamStr(1));
  Application.Terminate;

  {    // Eğer program parametre ile açıldıysa (sağ tık menüsünden)
  if ParamCount > 0 then
  begin
    FConvertMode := True;

    // Hiçbir pencere gösterme
    Application.ShowMainForm := False;
    WindowState := wsMinimized;
    ShowInTaskbar := False;
    Visible := False;

    // Dosyayı işle
    ProcessFile(ParamStr(1));
  end
  else
  begin
    // Normal açılış - ayarlar penceresini göster
    FConvertMode := False;
    SettingsForm.Show;
  end; }
end;

procedure TMainForm.ProcessFile(const FileName: string);
var
  Point: TPoint;
begin
  if not IsImageFile(FileName) then
  begin
    ShowMessage('Bu program sadece resim dosyaları için çalışır!');
    Application.Terminate;
    Exit;
  end;

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
      SourceBitmap.SaveToFile(TargetFile);
    end;

    ShowMessage('Dönüştürme tamamlandı: ' + ExtractFileName(TargetFile));

    if FConvertMode then
      Application.Terminate;
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

procedure TMainForm.InstallContextMenu;
var
  Reg: TRegistry;
  ExePath: string;
begin
  ExePath := Application.ExeName;
  Reg := TRegistry.Create;
  try
    Reg.RootKey := HKEY_CLASSES_ROOT;

    if Reg.OpenKey('SystemFileAssociations\image\shell\ImageConverter', True) then
    begin
      Reg.WriteString('', 'Resim Dönüştür');
      Reg.WriteString('Icon', ExePath);
      Reg.CloseKey;

      if Reg.OpenKey('SystemFileAssociations\image\shell\ImageConverter\command', True) then
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

    if Reg.KeyExists('SystemFileAssociations\image\shell\ImageConverter\command') then
      Reg.DeleteKey('SystemFileAssociations\image\shell\ImageConverter\command');

    if Reg.KeyExists('SystemFileAssociations\image\shell\ImageConverter') then
      Reg.DeleteKey('SystemFileAssociations\image\shell\ImageConverter');
  finally
    Reg.Free;
  end;
end;

procedure TMainForm.SettingsMenuClick(Sender: TObject);
begin
  PopupMenu1.CloseMenu;
  Application.ProcessMessages;
  Sleep(100);
  SettingsForm.ShowModal;
end;

procedure TMainForm.ExitMenuClick(Sender: TObject);
begin
  Application.Terminate;
end;

end.
