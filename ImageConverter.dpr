program ImageConverter;

uses
  Vcl.Forms,
  System.SysUtils,
  Main in 'Main.pas' {MainForm},
  Settings in 'Settings.pas' {SettingsForm};

{$R *.res}

begin
   Application.Initialize;
  Application.MainFormOnTaskbar := True;

  // Programın nasıl başlatıldığını kontrol et
  if ParamCount > 0 then
  begin
    // SAĞ TIK MENÜSÜNDEN AÇILDI (SESSİZ MOD)
    // Ana formun görünmesini engelle
    Application.ShowMainForm := False;
    Application.CreateForm(TMainForm, MainForm);
  end
  else
  begin
    // DOĞRUDAN ÇALIŞTIRILDI (AYARLAR EKRANI)
    // Ayarlar formunu ana form olarak oluştur
    Application.CreateForm(TSettingsForm, SettingsForm);
  end;

  // Uygulamayı çalıştır
  Application.Run;
//end;

end.