program ShowActiveXIF;

uses
  Forms,
  mainForm in 'mainForm.pas' {Form1},
  untComLib in 'untComLib.pas',
  untComTypeLibrary in 'untComTypeLibrary.pas';

{$R *.res}

begin
  Application.Initialize;
  Application.CreateForm(TForm1, Form1);
  Application.Run;
end.
