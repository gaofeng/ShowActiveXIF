object Form1: TForm1
  Left = 192
  Top = 115
  Caption = 'Show ActiveX Interface'
  ClientHeight = 466
  ClientWidth = 971
  Color = clBtnFace
  Font.Charset = ANSI_CHARSET
  Font.Color = clWindowText
  Font.Height = -12
  Font.Name = #23435#20307
  Font.Style = []
  OldCreateOrder = False
  PixelsPerInch = 96
  TextHeight = 12
  object Label1: TLabel
    Left = 264
    Top = 40
    Width = 36
    Height = 12
    Caption = 'Label1'
  end
  object Button1: TButton
    Left = 32
    Top = 168
    Width = 75
    Height = 25
    Caption = 'Show'
    TabOrder = 0
    OnClick = Button1Click
  end
  object Edit1: TEdit
    Left = 32
    Top = 128
    Width = 161
    Height = 20
    TabOrder = 1
    Text = 'c:\test.ocx'
  end
  object TreeView1: TTreeView
    Left = 256
    Top = 64
    Width = 665
    Height = 361
    Indent = 19
    TabOrder = 2
  end
  object Button2: TButton
    Left = 194
    Top = 129
    Width = 21
    Height = 20
    Caption = '...'
    TabOrder = 3
    OnClick = Button2Click
  end
  object OpenDialog1: TOpenDialog
    Left = 160
    Top = 24
  end
end
