object Form1: TForm1
  Left = 192
  Top = 115
  Caption = 'ActiveX'#25509#21475#26597#30475
  ClientHeight = 479
  ClientWidth = 777
  Color = clBtnFace
  Font.Charset = ANSI_CHARSET
  Font.Color = clWindowText
  Font.Height = -12
  Font.Name = #23435#20307
  Font.Style = []
  OldCreateOrder = False
  OnCreate = FormCreate
  DesignSize = (
    777
    479)
  PixelsPerInch = 96
  TextHeight = 12
  object LabelCOMInfo: TLabel
    Left = 8
    Top = 62
    Width = 30
    Height = 12
    Caption = #20449#24687':'
  end
  object Label1: TLabel
    AlignWithMargins = True
    Left = 8
    Top = 463
    Width = 168
    Height = 12
    Anchors = [akLeft, akBottom]
    Caption = '2012/12/23 gf@dareway.com.cn'
  end
  object lbl1: TLabel
    Left = 96
    Top = 31
    Width = 72
    Height = 12
    Caption = #25903#25345#25991#20214#25302#20837
    Font.Charset = ANSI_CHARSET
    Font.Color = clWindowText
    Font.Height = -12
    Font.Name = #23435#20307
    Font.Style = [fsItalic]
    ParentFont = False
  end
  object Button1: TButton
    Left = 8
    Top = 31
    Width = 82
    Height = 28
    Caption = #26174#31034
    TabOrder = 0
    OnClick = Button1Click
  end
  object Edit1: TEdit
    Left = 96
    Top = 5
    Width = 673
    Height = 20
    HelpType = htKeyword
    Anchors = [akLeft, akTop, akRight]
    TabOrder = 1
  end
  object TreeView1: TTreeView
    Left = 543
    Top = 31
    Width = 74
    Height = 33
    Indent = 19
    TabOrder = 2
    Visible = False
  end
  object Button2: TButton
    Left = 8
    Top = 3
    Width = 82
    Height = 28
    Hint = 'sdfdfe'
    Caption = #25171#24320#25991#20214'...'
    TabOrder = 3
    OnClick = Button2Click
  end
  object Memo1: TMemo
    AlignWithMargins = True
    Left = 8
    Top = 80
    Width = 761
    Height = 377
    Anchors = [akLeft, akTop, akRight, akBottom]
    Lines.Strings = (
      '')
    ReadOnly = True
    ScrollBars = ssBoth
    TabOrder = 4
    WordWrap = False
  end
  object OpenDialog1: TOpenDialog
    Filter = 'ActiveX'#25511#20214'(*.ocx;*.dll)|*.ocx;*.dll'
    InitialDir = 'D:\'
    Left = 736
    Top = 40
  end
end
