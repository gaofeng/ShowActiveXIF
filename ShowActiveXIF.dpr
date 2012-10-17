program ShowActiveXIF;

uses
  Forms,
  mainForm in 'mainForm.pas' {Form1},
  untComLib in 'untComLib.pas',
  untComTypeLibrary in 'untComTypeLibrary.pas';

{$R *.res}

begin
  Application.Initialize;
  Application.Title := 'ActiveX控件接口查看工具';
  Application.CreateForm(TForm1, Form1);
  Application.Run;
end.
