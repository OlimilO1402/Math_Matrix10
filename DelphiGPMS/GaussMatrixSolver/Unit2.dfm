object FrmSelectMethod: TFrmSelectMethod
  Left = 267
  Top = 187
  Width = 201
  Height = 155
  Caption = 'Select Mehod'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  PixelsPerInch = 96
  TextHeight = 13
  object RGSelectCalcMethod: TRadioGroup
    Left = 8
    Top = 8
    Width = 177
    Height = 73
    Caption = 'vorhandene L'#246'sungsmethoden:'
    Items.Strings = (
      'PyramidMatrixSolver'
      'SparseMatrixSolver')
    TabOrder = 0
  end
  object BtnOK: TButton
    Left = 16
    Top = 88
    Width = 75
    Height = 25
    Caption = 'OK'
    TabOrder = 1
    OnClick = BtnOKClick
  end
  object Button1: TButton
    Left = 104
    Top = 88
    Width = 75
    Height = 25
    Caption = 'Abbrechen'
    TabOrder = 2
  end
end
