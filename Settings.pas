unit Settings;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants,
  System.Classes, Vcl.Graphics, Vcl.Controls, Vcl.Forms, Vcl.Dialogs,
  Vcl.StdCtrls, Vcl.ExtCtrls, System.Win.Registry,Winapi.ShellAPI,System.IOUtils,main;

type
  TSettingsForm = class(TForm)
    GroupBox1: TGroupBox;
    RadioButton1: TRadioButton;
    RadioButton2: TRadioButton;
    RadioButton3: TRadioButton;
    Button1: TButton;
    Button2: TButton;
    Button3: TButton;
    procedure FormCreate(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure Button3Click(Sender: TObject);
  private
    procedure LoadSettings;
    procedure SaveSettings;
    procedure UninstallProgram;
    procedure ExplorerRefresh;
  public
  end;

var
  SettingsForm: TSettingsForm;

implementation

{$R *.dfm}

procedure TSettingsForm.FormCreate(Sender: TObject);
begin
  LoadSettings;
end;

procedure TSettingsForm.LoadSettings;
var
  Reg: TRegistry;
  NamingStyle: Integer;
begin
  Reg := TRegistry.Create;
  try
    Reg.RootKey := HKEY_CURRENT_USER;
    if Reg.OpenKey('Software\ImageConverter', False) then
      NamingStyle := Reg.ReadInteger('NamingStyle')
    else
      NamingStyle := 0;

    case NamingStyle of
      0: RadioButton1.Checked := True;
      1: RadioButton2.Checked := True;
      2: RadioButton3.Checked := True;
    end;
  finally
    Reg.Free;
  end;
end;

procedure TSettingsForm.SaveSettings;
var
  Reg: TRegistry;
  NamingStyle: Integer;
begin
  if RadioButton1.Checked then NamingStyle := 0
  else if RadioButton2.Checked then NamingStyle := 1
  else if RadioButton3.Checked then NamingStyle := 2
  else NamingStyle := 0;

  Reg := TRegistry.Create;
  try
    Reg.RootKey := HKEY_CURRENT_USER;
    if Reg.OpenKey('Software\ImageConverter', True) then
      Reg.WriteInteger('NamingStyle', NamingStyle);
  finally
    Reg.Free;
  end;
end;

procedure TSettingsForm.ExplorerRefresh;
var
  Res: Cardinal;
begin
  // Explorer'ı yeniden başlat
  Res := WinExec('taskkill /f /im explorer.exe', SW_HIDE);
  Sleep(1000);
  Res := WinExec('explorer.exe', SW_SHOW);
end;

procedure TSettingsForm.UninstallProgram;
var
  Reg: TRegistry;
begin
  if MessageDlg('Program sağ tık menüsünden kaldırılsın mı?', mtConfirmation, [mbYes, mbNo], 0) = mrYes then
  begin
    try
      // Ana formdaki kaldırma prosedürünü çağır
      MainForm.UninstallContextMenu;

      // Ayarları sil
      Reg := TRegistry.Create;
      try
        Reg.RootKey := HKEY_CURRENT_USER;
        if Reg.KeyExists('Software\ImageConverter') then
        begin
          Reg.DeleteKey('ImageConverter');
        end;
      finally
        Reg.Free;
      end;

      // Windows Shell'e (Explorer) değişiklikleri bildir
      //SHChangeNotify(SHCNE_ASSOCCHANGED, SHCNF_IDLIST, nil, nil);

      ShowMessage('Program sağ tık menüsünden başarıyla kaldırıldı.');
    except
      on E: Exception do
        ShowMessage('Kaldırma sırasında bir hata oluştu: ' + E.Message + #13#10 +
                    'Lütfen programı "Yönetici olarak çalıştır" seçeneği ile yeniden deneyin.');
    end;
    Close;
  end;
end;

procedure TSettingsForm.Button1Click(Sender: TObject);
begin
  SaveSettings;
  ShowMessage('Ayarlar kaydedildi.');
  Close;
end;



procedure TSettingsForm.Button2Click(Sender: TObject);
begin
  Close;
end;

procedure TSettingsForm.Button3Click(Sender: TObject);
var
  BatPath: string;
  ExePath: string;
begin
  if MessageDlg('Program kaldırılacak. Onaylıyor musunuz?', mtConfirmation, [mbYes, mbNo], 0) = mrYes then
  begin
    ExePath := ExtractFilePath(Application.ExeName);
    BatPath := ExePath + 'sil.bat';

    if FileExists(BatPath) then
    begin
      // Bat dosyasını sessizce çalıştır
      WinExec(PAnsiChar(AnsiString(BatPath)), SW_HIDE);

      // Programı kapat
      Application.Terminate;
    end
    else
    begin
      ShowMessage('sil.bat dosyası bulunamadı!');
    end;
  end;

end;

end.
