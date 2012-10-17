object Form1: TForm1
  Left = 192
  Top = 115
  Caption = 'Show ActiveX Interface'
  ClientHeight = 481
  ClientWidth = 776
  Color = clBtnFace
  Font.Charset = ANSI_CHARSET
  Font.Color = clWindowText
  Font.Height = -12
  Font.Name = #23435#20307
  Font.Style = []
  OldCreateOrder = False
  PixelsPerInch = 96
  TextHeight = 12
  object LabelCOMInfo: TLabel
    Left = 113
    Top = 44
    Width = 6
    Height = 12
  end
  object Button1: TButton
    Left = 8
    Top = 31
    Width = 75
    Height = 25
    Caption = 'Show'
    TabOrder = 0
    OnClick = Button1Click
  end
  object Edit1: TEdit
    Left = 89
    Top = 5
    Width = 640
    Height = 20
    TabOrder = 1
    Text = 'c:\test.ocx'
  end
  object TreeView1: TTreeView
    Left = 8
    Top = 62
    Width = 760
    Height = 411
    Indent = 19
    TabOrder = 2
  end
  object Button2: TButton
    Left = 8
    Top = 3
    Width = 73
    Height = 28
    Caption = #25171#24320#25991#20214
    TabOrder = 3
    OnClick = Button2Click
  end
  object OpenDialog1: TOpenDialog
    Left = 744
  end
end
