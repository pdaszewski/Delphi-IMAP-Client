object AOknoGl: TAOknoGl
  Left = 0
  Top = 0
  BorderIcons = [biSystemMenu, biMinimize]
  BorderStyle = bsSingle
  Caption = 'Mailer'
  ClientHeight = 474
  ClientWidth = 790
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
  OnClose = FormClose
  OnCreate = FormCreate
  OnShow = FormShow
  DesignSize = (
    790
    474)
  PixelsPerInch = 96
  TextHeight = 13
  object btn_pobierz: TButton
    Left = 8
    Top = 8
    Width = 281
    Height = 25
    Caption = 'pobierz dane ze skrzynki pocztowej'
    TabOrder = 0
    OnClick = btn_pobierzClick
  end
  object ListBox1: TListBox
    Left = 8
    Top = 39
    Width = 774
    Height = 404
    Anchors = [akLeft, akTop, akRight, akBottom]
    ItemHeight = 13
    TabOrder = 1
  end
  object ProgressBar: TProgressBar
    Left = 8
    Top = 449
    Width = 774
    Height = 17
    Anchors = [akLeft, akRight, akBottom]
    TabOrder = 2
  end
  object btn_stop: TButton
    Left = 295
    Top = 8
    Width = 75
    Height = 25
    Caption = 'STOP'
    TabOrder = 3
    OnClick = btn_stopClick
  end
  object ProgressBarAutoStart: TProgressBar
    Left = 376
    Top = 8
    Width = 406
    Height = 25
    TabOrder = 4
  end
  object Start: TTimer
    Enabled = False
    Interval = 100
    OnTimer = StartTimer
    Left = 704
    Top = 72
  end
  object AutoRun: TTimer
    Enabled = False
    OnTimer = AutoRunTimer
    Left = 704
    Top = 128
  end
end
