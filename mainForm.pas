unit mainForm;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls,untComLib, ComCtrls,untComTypeLibrary, ShellApi;

type
  TForm1 = class(TForm)
    Button1: TButton;
    Edit1: TEdit;
    TreeView1: TTreeView;
    Button2: TButton;
    OpenDialog1: TOpenDialog;
    LabelCOMInfo: TLabel;
    Memo1: TMemo;
    Label1: TLabel;
    procedure Button1Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
  private
    { Private declarations }
    procedure AddToTreeView(list:TComInfoList);
    procedure ComLibToTreeView(com1: TTypeLibrary);
    procedure ComLibToMemo(com1: TTypeLibrary);
  public
    { Public declarations }
    // declare our DROPFILES message handler
    procedure AcceptFiles( var msg : TMessage );
      message WM_DROPFILES;
  end;

var
  Form1: TForm1;

implementation

{$R *.dfm}

procedure TForm1.AcceptFiles( var msg : TMessage );
const
  cnMaxFileNameLen = 255;
var
  i,
  nCount     : integer;
  acFileName : array [0..cnMaxFileNameLen] of char;
  FileExt    : string;
begin
  // find out how many files we're accepting
  nCount := DragQueryFile( msg.WParam,
                           $FFFFFFFF,
                           acFileName,
                           cnMaxFileNameLen );

  // query Windows one at a time for the file name
  for i := 0 to nCount-1 do
  begin
    DragQueryFile( msg.WParam, i,
                   acFileName, cnMaxFileNameLen );

    // do your thing with the acFileName
    //MessageBox( Handle, acFileName, '', MB_OK );
    FileExt := ExtractFileExt(acFileName);

    Edit1.Text := acFileName;
  end;

  // let Windows know that you're done
  DragFinish( msg.WParam );
end;

procedure TForm1.AddToTreeView(list: TComInfoList);
var
  i,j:Integer;
  MyTreeNode1: TTreeNode;
begin
  for i:=0 to List.Count -1 do
  begin
    MyTreeNode1:=TreeView1.Items.Add(nil, List.Items[i].CoClassName +'\'+ List.Items[i].Guid +'['+IntToHex(List.Items[i].TypeFlags,4 )+']' );
      for j:=0 to List.Items[i].Functions.Count -1 do
      begin
        TreeView1.Items.AddChild(MyTreeNode1,Format('%-30.30s [%4d][ %s ] %8.8x',[List.Items[i].Functions.Items[j].Name, List.Items[i].Functions.Items[j].oVft ,List.Items[i].Functions.Items[j].FuncKind,List.Items[i].Functions.Items[j].Offset] ));
      end;
  end;
end;

procedure TForm1.ComLibToTreeView(com1: TTypeLibrary);
var
  i,j:Integer;
  MyTreeNode1: TTreeNode;
begin
  TreeView1.Items.Clear;

  for i:=0 to com1.CoClassCount-1 do begin
    with com1.CoClasses[i] do begin
      MyTreeNode1:=TreeView1.Items.Add(nil,'[ CoClass ] '+Name +'\'+GuidToString(Guid ));
      for j:=0 to Interfacecount-1 do begin
        TreeView1.Items.AddChild(MyTreeNode1,Format('%s\%s',[Interfaces[j].Name,
               GuidToString(Interfaces[j].Guid ) ]));
      end;
      for j:=0 to functionCount-1 do begin
        TreeView1.Items.AddChild(MyTreeNode1,Format('%-30.30s  [%s] %8.8x',
              [Functions[j].Name,Functions[j].InvokeKind ,Functions[j].Offset ]));
      end;
    end;
  end;

  for i:=0 to com1.InterfaceCount -1 do begin
    with com1.Interfaces[i] do begin
      MyTreeNode1:=TreeView1.Items.Add(nil,'[ Interface ] '+Name );
      for j:=0 to FunctionCount-1 do begin
        TreeView1.Items.AddChild(MyTreeNode1,Format('[%d]%s',[Functions[j].id,
              Functions[j].Name]));
      end;
    end;
  end;
end;

procedure TForm1.ComLibToMemo(com1: TTypeLibrary);
var
  i, j, k:Integer;
  MyTreeNode1: TTreeNode;
  str: string;
begin
  Memo1.Clear;

  for i:=0 to com1.CoClassCount-1 do
  begin
    with com1.CoClasses[i] do
    begin
      str := Format('[CoClass %d/%d]', [i + 1, com1.CoClassCount]) + Name + '\'+ GuidToString(Guid );
      Memo1.Lines.Add(str);
      str := Format('    Interface count: %d', [Interfacecount]);
      Memo1.Lines.Add(str);
      for j:=0 to Interfacecount-1 do
      begin
        str := Format('        %s\%s',
              [Interfaces[j].Name,
              GuidToString(Interfaces[j].Guid ) ]);
        Memo1.Lines.Add(str);
      end;
      str := Format('    Function count: %d', [functionCount]);
      Memo1.Lines.Add(str);
      for j:=0 to functionCount-1 do
      begin
        str := Format('    %-30.30s  [%s] %8.8x',
              [Functions[j].Name,
              Functions[j].InvokeKind ,
              Functions[j].Offset ]);
        Memo1.Lines.Add(str);
      end;
    end;
  end;
  //显示所有接口的函数
  for i:=0 to com1.InterfaceCount -1 do
  begin
    with com1.Interfaces[i] do
    begin
      str := Format('[Interface %s(%d/%d)] %d Functions',
              [com1.Interfaces[i].Name, i + 1, com1.InterfaceCount,
              FunctionCount]);
      Memo1.Lines.Add(str);
      for j:=0 to FunctionCount-1 do
      begin
        str := Format('    [%-3d]%s %s',
              [Functions[j].ID,
              Functions[j].Value.DataTypeName,
              Functions[j].Name]);
        //Memo1.Lines.Add(str);
        str := str + '(';
        for k := 0 to Functions[j].ParamCount - 1 do
        begin
          str := str + Format('%s %s', [Functions[j].Params[k].DataTypeName,
                Functions[j].Params[k].Name]);
          if k < Functions[j].ParamCount - 1 then
          begin
            str := str + ', ';
          end;
        end;
        str := str + ')';
        Memo1.Lines.Add(str);
      end;
    end;
  end;
  //显示所有接口的属性
  for i:=0 to com1.InterfaceCount -1 do
  begin
    str := Format('[Interface %s(%d/%d)] %d Properties',
              [com1.Interfaces[i].Name, i + 1, com1.InterfaceCount,
              com1.Interfaces[i].PropertyCount]);
    Memo1.Lines.Add(str);
    for j:=0 to com1.Interfaces[i].PropertyCount - 1 do
    begin
      str := Format('    [%-3d]%s %s',
            [com1.Interfaces[i].Properties[j].ID,
             com1.Interfaces[i].Properties[j].Value.DataTypeName,
             com1.Interfaces[i].Properties[j].Name
             ]);
      Memo1.Lines.Add(str);
    end;
  end;
end;

procedure TForm1.Button1Click(Sender: TObject);
var
  com1:TTypeLibrary;

  i,j:Integer;
  MyTreeNode1: TTreeNode;
  path :string;
begin
  if Edit1.Text = '' then
  begin
    exit;
  end;
  try
    com1:=TTypeLibrary.Create(Edit1.Text);
    LabelCOMInfo.Caption := Com1.Name +'\'+GuidTostring(com1.Guid) +
          com1.Description;

    ComLibToTreeView(com1);
    ComLibToMemo(com1);
    com1.Free ;
  except
    on E : Exception do
      ShowMessage(E.Message);
  end;
end;

procedure TForm1.Button2Click(Sender: TObject);
begin
  if OpenDialog1.Execute then
    Edit1.Text := OpenDialog1.FileName ;
end;

procedure TForm1.FormCreate(Sender: TObject);
begin
// tell Windows that you're
  // accepting drag and drop files
  //
  DragAcceptFiles( Handle, True );
end;

end.
