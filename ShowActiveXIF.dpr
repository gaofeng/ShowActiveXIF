program ShowActiveXIF;

uses
  Forms,
  mainForm in 'mainForm.pas' {Form1},
  untComLib in 'untComLib.pas',
  untComTypeLibrary in 'untComTypeLibrary.pas';

{$R *.res}
 var
  para:string;
begin
  Application.Initialize;
  Application.Title := 'ActiveX�ؼ��ӿڲ鿴����';
  Application.CreateForm(TForm1, Form1);
  Form1.FilePath := ParamStr(1);
  Application.Run;
end.
