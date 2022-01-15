object Form1: TForm1
  Left = 196
  Top = 114
  Width = 667
  Height = 250
  Caption = 'GPMXS Gauss Pyramid Matrix Solver'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  Menu = MainMenu1
  OldCreateOrder = False
  OnCreate = FormCreate
  PixelsPerInch = 96
  TextHeight = 13
  object Panel2: TPanel
    Left = 0
    Top = 0
    Width = 659
    Height = 17
    Align = alTop
    BevelOuter = bvNone
    TabOrder = 0
    object Panel1: TPanel
      Left = 0
      Top = 0
      Width = 481
      Height = 17
      Align = alClient
      BevelOuter = bvLowered
      Caption = 'A *        '
      TabOrder = 0
    end
    object Panel3: TPanel
      Left = 481
      Top = 0
      Width = 89
      Height = 17
      Align = alRight
      BevelOuter = bvLowered
      Caption = 'x =        '
      TabOrder = 1
    end
    object Panel4: TPanel
      Left = 570
      Top = 0
      Width = 89
      Height = 17
      Align = alRight
      BevelOuter = bvLowered
      Caption = 'b        '
      TabOrder = 2
    end
  end
  object StatusBar1: TStatusBar
    Left = 0
    Top = 177
    Width = 659
    Height = 19
    Panels = <>
  end
  object Panel5: TPanel
    Left = 0
    Top = 17
    Width = 659
    Height = 160
    Align = alClient
    BevelOuter = bvNone
    TabOrder = 2
    object Splitter1: TSplitter
      Left = 569
      Top = 0
      Width = 2
      Height = 160
      Align = alRight
      ResizeStyle = rsUpdate
      OnMoved = Splitter1Moved
    end
    object Splitter2: TSplitter
      Left = 479
      Top = 0
      Width = 2
      Height = 160
      Align = alRight
      ResizeStyle = rsUpdate
      OnMoved = Splitter2Moved
    end
    object Panel6: TPanel
      Left = 0
      Top = 0
      Width = 479
      Height = 160
      Align = alClient
      BevelOuter = bvNone
      Caption = 'A *'
      TabOrder = 0
      object StG_A: TStringGrid
        Left = 0
        Top = 0
        Width = 479
        Height = 160
        Align = alClient
        ColCount = 256
        DefaultRowHeight = 16
        FixedCols = 0
        RowCount = 256
        FixedRows = 0
        Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRangeSelect, goDrawFocusSelected, goEditing]
        TabOrder = 0
        OnKeyDown = StG_AKeyDown
        OnKeyUp = StG_AKeyUp
        OnSelectCell = StG_ASelectCell
      end
    end
    object Panel7: TPanel
      Left = 481
      Top = 0
      Width = 88
      Height = 160
      Align = alRight
      BevelOuter = bvNone
      Caption = 'x ='
      TabOrder = 1
      object StG_x: TStringGrid
        Left = 0
        Top = 0
        Width = 88
        Height = 160
        Align = alClient
        ColCount = 1
        DefaultRowHeight = 16
        FixedCols = 0
        RowCount = 256
        FixedRows = 0
        Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRangeSelect, goDrawFocusSelected, goEditing]
        TabOrder = 0
        OnSelectCell = StG_ASelectCell
      end
    end
    object Panel8: TPanel
      Left = 571
      Top = 0
      Width = 88
      Height = 160
      Align = alRight
      BevelOuter = bvNone
      Caption = 'b'
      TabOrder = 2
      object StG_b: TStringGrid
        Left = 0
        Top = 0
        Width = 88
        Height = 160
        Align = alClient
        ColCount = 1
        DefaultRowHeight = 16
        FixedCols = 0
        RowCount = 256
        FixedRows = 0
        Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRangeSelect, goDrawFocusSelected, goEditing]
        TabOrder = 0
        OnKeyDown = StG_AKeyDown
        OnKeyUp = StG_AKeyUp
        OnSelectCell = StG_ASelectCell
      end
    end
  end
  object MainMenu1: TMainMenu
    object Berechnen1: TMenuItem
      Caption = '&Datei'
      object Beispielladen1: TMenuItem
        Caption = 'Beispiel laden'
        OnClick = Beispielladen_Click
      end
      object mnuSelectMethod: TMenuItem
        Caption = 'Methode w'#228'hlen...'
        OnClick = mnuSelectMethodClick
      end
      object Berechnen2: TMenuItem
        Caption = 'Be&rechnen!'
        OnClick = Berechnen_Click
      end
      object Onlinemitrechnen1: TMenuItem
        AutoCheck = True
        Caption = '&Online mitrechnen'
      end
      object ClearallGrids1: TMenuItem
        Caption = 'Grids l&eeren'
        OnClick = ClearallGrids_Click
      end
      object N1: TMenuItem
        Caption = '-'
      end
      object Beenden1: TMenuItem
        Caption = 'Be&enden'
        OnClick = Beenden_Click
      end
    end
    object Bearbeiten1: TMenuItem
      Caption = 'B&earbeiten'
      object Kopiermodus1: TMenuItem
        AutoCheck = True
        Caption = 'Auswahlmodus'
        OnClick = Kopiermodus1Click
      end
      object N2: TMenuItem
        Caption = '-'
      end
      object Kopieren: TMenuItem
        Caption = 'Kopieren'
      end
      object Einfgen: TMenuItem
        Caption = 'Einf'#252'gen'
        OnClick = EinfgenClick
      end
    end
    object Extras1: TMenuItem
      Caption = 'E&xtras'
      object MatrixSpaltenbreite1: TMenuItem
        Caption = 'Tabelle Spaltenbreite'
        OnClick = MatrixSpaltenbreite1Click
      end
      object Einheitsmatrixgenerieren1: TMenuItem
        Caption = 'Einheitsmatrix generieren'
        OnClick = Einheitsmatrixgenerieren1Click
      end
      object Matrixvergrern1: TMenuItem
        Caption = 'Tabelle vergr'#246#223'ern'
        OnClick = Matrixvergrern1Click
      end
    end
    object N3: TMenuItem
      Caption = ' &? '
      object Info1: TMenuItem
        Caption = '&Info'
        OnClick = Info1Click
      end
    end
  end
end
