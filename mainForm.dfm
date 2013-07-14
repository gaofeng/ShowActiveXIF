object Form1: TForm1
  Left = 360
  Top = 225
  Caption = 'ActiveX'#25509#21475#26597#30475
  ClientHeight = 479
  ClientWidth = 777
  Color = clBtnFace
  DoubleBuffered = True
  Font.Charset = ANSI_CHARSET
  Font.Color = clWindowText
  Font.Height = -12
  Font.Name = #23435#20307
  Font.Style = []
  KeyPreview = True
  OldCreateOrder = False
  Position = poDesigned
  OnCreate = FormCreate
  OnKeyDown = FormKeyDown
  OnShow = FormShow
  DesignSize = (
    777
    479)
  PixelsPerInch = 96
  TextHeight = 12
  object LabelCOMInfo: TLabel
    Left = 101
    Top = 31
    Width = 36
    Height = 12
    Caption = 'CLSID:'
  end
  object Label1: TLabel
    AlignWithMargins = True
    Left = 8
    Top = 463
    Width = 168
    Height = 12
    Anchors = [akLeft, akBottom]
    Caption = '2013/07/14 gf@dareway.com.cn'
  end
  object lbl1: TLabel
    Left = 411
    Top = 31
    Width = 36
    Height = 12
    Caption = 'ProgID'
  end
  object Button1: TButton
    Left = 8
    Top = 34
    Width = 82
    Height = 27
    Caption = #20998#26512
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
    Left = 695
    Top = 8
    Width = 74
    Height = 33
    Indent = 19
    TabOrder = 2
    Visible = False
  end
  object Button2: TButton
    Left = 8
    Top = 0
    Width = 82
    Height = 28
    Hint = #25903#25345#25991#20214#25302#20837#35813#31383#21475
    Caption = #36873#25321#25991#20214'...'
    ParentShowHint = False
    ShowHint = True
    TabOrder = 3
    OnClick = Button2Click
  end
  object Memo1: TMemo
    AlignWithMargins = True
    Left = 8
    Top = 65
    Width = 761
    Height = 392
    Anchors = [akLeft, akTop, akRight, akBottom]
    Font.Charset = ANSI_CHARSET
    Font.Color = clWindowText
    Font.Height = -12
    Font.Name = #23435#20307
    Font.Style = []
    Lines.Strings = (
      '')
    ParentFont = False
    ReadOnly = True
    ScrollBars = ssBoth
    TabOrder = 4
    WordWrap = False
    OnKeyDown = Memo1KeyDown
  end
  object edtClsId: TEdit
    Left = 137
    Top = 28
    Width = 249
    Height = 18
    Ctl3D = False
    ParentCtl3D = False
    TabOrder = 5
    OnDblClick = edtClsIdDblClick
    OnKeyPress = edtClsIdKeyPress
  end
  object edtProgID: TEdit
    Left = 453
    Top = 28
    Width = 165
    Height = 20
    AutoSize = False
    Ctl3D = False
    ParentCtl3D = False
    TabOrder = 6
  end
  object OpenDialog1: TOpenDialog
    Filter = 'ActiveX'#25511#20214'(*.ocx;*.dll)|*.ocx;*.dll'
    InitialDir = 'D:\'
    Left = 736
    Top = 40
  end
end
