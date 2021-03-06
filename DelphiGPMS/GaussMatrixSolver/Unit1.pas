unit Unit1;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ExtCtrls, Menus, Grids, ComCtrls, Math, UClsSparseSolv;

type

  TForm1 = class(TForm)
    MainMenu1: TMainMenu;
    Panel2: TPanel;
    Berechnen1: TMenuItem;
    StatusBar1: TStatusBar;
    Panel1: TPanel;
    Panel3: TPanel;
    Panel4: TPanel;
    Panel5: TPanel;
    Panel6: TPanel;
    Panel7: TPanel;
    Panel8: TPanel;
    StG_A: TStringGrid;
    StG_x: TStringGrid;
    StG_b: TStringGrid;
    Berechnen2: TMenuItem;
    Beispielladen1: TMenuItem;
    Beenden1: TMenuItem;
    ClearallGrids1: TMenuItem;
    N1: TMenuItem;
    Onlinemitrechnen1: TMenuItem;
    Splitter1: TSplitter;
    Splitter2: TSplitter;
    Bearbeiten1: TMenuItem;
    Kopiermodus1: TMenuItem;
    Kopieren: TMenuItem;
    Einfgen: TMenuItem;
    N2: TMenuItem;
    MatrixSpaltenbreite1: TMenuItem;
    Extras1: TMenuItem;
    Einheitsmatrixgenerieren1: TMenuItem;
    N3: TMenuItem;
    Info1: TMenuItem;
    Matrixvergrern1: TMenuItem;
    mnuSelectMethod: TMenuItem;
    procedure FormCreate(Sender: TObject);
    procedure Beispielladen_Click(Sender: TObject);
    procedure Berechnen_Click(Sender: TObject);
    procedure ClearallGrids_Click(Sender: TObject);
    procedure Beenden_Click(Sender: TObject);

    procedure StG_AKeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure StG_AKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);

    procedure Splitter1Moved(Sender: TObject);
    procedure Splitter2Moved(Sender: TObject);
    procedure Kopiermodus1Click(Sender: TObject);
    procedure EinfgenClick(Sender: TObject);

    procedure StG_ASelectCell(Sender: TObject; ACol, ARow: Integer;
      var CanSelect: Boolean);
    procedure MatrixSpaltenbreite1Click(Sender: TObject);
    procedure Info1Click(Sender: TObject);
    procedure Einheitsmatrixgenerieren1Click(Sender: TObject);
    procedure Matrixvergrern1Click(Sender: TObject);
    procedure mnuSelectMethodClick(Sender: TObject);
  private
    { Private-Deklarationen }
    function GetStringGridExtent(StG: TStringGrid; Mode: Char): Integer;
    procedure FillZeroCells(Ext: Integer; StG: TStringGrid);
    function  GetHowMuchMemorYouNeed(m: Integer): String;
  public
    { Public-Deklarationen }
  end;
  TOMPyrMxSolv = class
  private
    //die Gr??e, Ausdehnung der Matrix (extent)
    iEx: Integer;
    procedure SetExtent(Ext: Integer);
    function  GetExtent: Integer;
    procedure Clear;
  public
    //Gesucht ist der x-Vektor der Glchg: Amat * x = b
    Amat: array of array of array of Double; //Pyramidenmatrix
    b:    array of array of Double; //Ergebnisvektor als Dreiecksmatrix
    x:    array of Double; //der gesuchte zu berechnende Vektor

    a2: array of array of Double;
    b2: array of Double;
    //Solver: TSelectSolver;
    property Ext: Integer read GetExtent write SetExtent;
    constructor Create(Ext: Integer);
    procedure SolvePyramidMx;
    procedure Solve2;
    //procedure SolveSinglValDecomp(var a: glmpbynp; m,n,mp,np: integer; var w: glnparray; var v: glnpbynp);
  end;

var
  Form1: TForm1;
  Solvr: TOMPyrMxSolv;
  Spars: TSparseSolv;

implementation

uses Unit2;

{$R *.dfm}
//var

  
procedure TForm1.FormCreate(Sender: TObject);
begin
  Solvr:=TOMPyrMxSolv.Create(1);
end;

procedure TForm1.StG_ASelectCell(Sender: TObject; ACol, ARow: Integer;
  var CanSelect: Boolean);
var ST, Bez, Cont: String;
begin
  Bez:=(Sender as TStringGrid).Name;
  Cont:=(Sender as TStringGrid).Cells[ACol, ARow];
  ST:=Bez + '[' + IntToStr(ACol) + ', ' +  IntToStr(ARow) + ']= ' + Cont;
  StatusBar1.SimpleText:=ST;
end;

procedure TForm1.Beispielladen_Click(Sender: TObject);
var
  A: array of array of Double;
  b: array of Double;
  Ex: Integer;
  i,j: Integer;
begin
  Ex:=5;
  SetLength(A, Ex, Ex);
  SetLength(b, Ex);

  A[0,0]:= 1.35; A[1,0]:=-0.35; A[2,0]:=-1.00; A[3,0]:= 0.00; A[4,0]:=-0.35;
  A[0,1]:=-0.35; A[1,1]:= 1.35; A[2,1]:= 0.00; A[3,1]:= 0.00; A[4,1]:= 0.35;
  A[0,2]:=-1.00; A[1,2]:= 0.00; A[2,2]:= 1.35; A[3,2]:= 0.35; A[4,2]:= 0.00;
  A[0,3]:= 0.00; A[1,3]:= 0.00; A[2,3]:= 0.35; A[3,3]:= 1.35; A[4,3]:= 0.00;
  A[0,4]:=-0.35; A[1,4]:= 0.35; A[2,4]:= 0.00; A[3,4]:= 0.00; A[4,4]:= 1.35;

  b[0]:=  0.00;
  b[1]:=  0.00;
  b[2]:= 10.00;
  b[3]:=-10.00;
  b[4]:=  0.00;

//Ergebnis sollte sein:
//  x[0]:= 24.29
//  x[1]:=  5.00
//  x[2]:= 29.29
//  x[3]:=-15.00
//  x[4]:=  5.00

  for j:=0 to Ex-1 do
  begin
    for i:= 0 to Ex-1 do
    begin
      StG_A.Cells[i, j]:=FloatToStr(A[i, j]);
    end;
    StG_b.Cells[0,j]:=FloatToStr(b[j]);
  end;
end;

procedure TForm1.Berechnen_Click(Sender: TObject);
var
  maxA, maxb, m: Integer;
  i, j: Integer;
  Mem: String;
  mess: String;
  RetV: Integer;
begin
  maxA:=GetStringGridExtent(StG_A, 'm'); //f?r die Matrix
  maxb:=GetStringGridExtent(StG_b, 'v'); //f?r den Vektor
  m:=max(maxA, maxb);
  FillZeroCells(m, stG_A);
  FillZeroCells(m, stG_b);

  Mem:=GetHowMuchMemoryouNeed(m);
  mess:='Sie ben?tigen ca.: ' + Mem + ' Speicher.';
  RetV:=MessageDlg(mess, mtInformation, [mbOK, mbCancel], 0);
  if RetV=mrCancel then exit;

  Case WhichS

  end;

  Solvr.Ext:=m;

  for i:=0 to m-1 do
  begin
    for j:=0 to m-1 do
    begin
      Solvr.Amat[0, i, j]:= StrToFloat(StG_A.Cells[i, j]);
      //Solvr.a2[i, j]:= StrToFloat(StG_A.Cells[i, j]);
    end;
  end;
  for i:=0 to m-1 do
  begin
    Solvr.b[0, i]:= StrToFloat(StG_b.Cells[0, i]);
    //Solvr.b2[i]:= StrToFloat(StG_b.Cells[0, i]);
  end;
  Solvr.SolvePyramidMx;
  //Solvr.Solve2;
  for i:=0 to m-1 do
  begin
    StG_x.Cells[0, i]:= FloatToStr(Solvr.x[i]);
  end;
end;

function TForm1.GetStringGridExtent(StG: TStringGrid; Mode: Char): Integer;
var i, j, maxi, maxj : Integer;
begin
  maxi:=0;
  maxj:=0;
//gleiche Spalten und Zeilen Zahl: von unten,
//alles was dem ersten Nuller kommt wird abgeschnitten
  case Mode of
  'm':
    for i:= 0 to StG.RowCount do
    begin
      if StG.Cells[i, i] <> '' then
      begin
        if StrToFloat(StG.Cells[i, i]) <> 0 then
        begin
          maxi:=max(maxi, i+1);
        end;
      end
      else
      begin
        break;
      end;
    end;
  'v':
    for i:= 0 to StG.RowCount do
    begin
      if StG.Cells[0, i] <> '' then
      begin
        if StrToFloat(StG.Cells[0, i]) <> 0 then
        begin
          maxi:=max(maxi, i+1);
        end;
      end
      else
      begin
        break;
      end;
    end;
  end;

//gleiche Spalten und Zeilen Zahl: von oben
//ausgeklammert, da sehr zeitintensiv bei gro?er Tabelle
 {
  for i:= StG.RowCount-1 downto 0 do
  begin
    if StG.Cells[i, i] <> '' then
    begin
      if StrToFloat(StG.Cells[i, i]) <> 0 then
      begin
        maxi:=max(maxi, i+1);
      end;
    end;
  end;
 }

//unterschiedliche Spalten und Zeilenanzahl:
//ausgeklammert, da ohnehin nicht relevant, aber sehr zeitintensiv!
 {
  for j:= StG.RowCount-1 downto 0 do
  begin
    for i:= StG.ColCount-1 downto 0 do
    begin
      if StG.Cells[i, j] <> '' then
      begin
        if StrToFloat(StG.Cells[i, j]) <> 0 then
        begin
          maxi:=max(maxi, i+1);
          maxj:=max(maxj, j+1);
        end
        else
        begin
          StG.Cells[i, j]:='';
        end;
      end;
    end;
  end;
 }
  result:= max(maxi, maxj);
end;

procedure TForm1.FillZeroCells(Ext: Integer; StG: TStringGrid);
var i, j: Integer;
begin
  for i:= 0 to Ext-1 do
  begin
    for j:= 0 to Ext-1 do
    begin
      if StG.Cells[i, j]= '' then
      begin
        StG.Cells[i, j]:= '0';
      end;
    end;
  end;
end;

function TForm1.GetHowMuchMemorYouNeed(m: Integer): String;
var
  i, BByte, MemPm, MemBb, MemXx, MemGs: Integer;
  MemKB, MemMB: Single;
  MemStr: String;
begin
  MemPm:= 0;  MemBb:= 0;  MemXx:= 0;
  MemGs:= 0;  MemMB:= 0;
  //if Double then
  BByte:=4;
  //else if Single then
  //BBYte:=2;
  for i:=1 to m do
  begin
    MemPm:=MemPm+i*i;
  end;
  //die Pyramidenmatrix ben?tigt MemPm Speicher in Byte
  MemPm:= MemPm * BByte;
  //der Ergebnisvektor b ben?tigt MemBb Speicher in Byte
  MemBb:= m * m * BByte div 2;
  //der X-Vektor s ben?tigt MemXx Speicher in Byte
  MemXx:= m * BByte;
  //Gesamt ben?tigt das System MemGs Speicher in Byte
  MemGs:= MemPm + MemBb + MemXx;
  MemStr:= IntToStr(MemGs) + ' Byte';
  if MemGs > 1024 then
  begin
    MemKB:=MemGs/1024;
    MemStr:=FloatToStr(MemKB) + ' kb';
  end;
  if MemGs > 1024*1024 then
  begin
    MemMB:=MemGs/1024/1024;
    MemStr:=FloatToStr(MemMB) + ' Mb';
  end;
  Result:=MemStr;
end;

procedure TForm1.StG_AKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
//
end;

procedure TForm1.StG_AKeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
var c: Char;
begin
  c:=Chr(Key);
  case c of
  '0'..'9': ;
  ',': Key:=VK_DECIMAL;
  '.': Key:=VK_DECIMAL;
    else
    begin
      case Key of
      vk_return, vk_tab, vk_back, vk_insert, vk_clear,
      vk_escape, vk_end, vk_next, vk_home, vk_delete,
      VK_SUBTRACT, vk_Up, vk_Down, vk_Left, vk_Right: ;
      190: Key:=vk_numpad0;
      else
        Key:=vk_numpad0;
      end;
    end;
  end;

  if Onlinemitrechnen1.Checked then
  begin
    Berechnen_Click(Sender);
  end;
end;

procedure TForm1.Splitter2Moved(Sender: TObject);
begin
  Stg_x.DefaultColWidth:=Stg_x.Width-25;
  Panel3.Width:=Panel7.Width + Splitter2.Width div 2;
end;

procedure TForm1.Splitter1Moved(Sender: TObject);
begin
  Stg_b.DefaultColWidth:=Stg_b.Width-25;
  Panel4.Width:=Panel8.Width + Splitter1.Width div 2;
end;

procedure TForm1.ClearallGrids_Click(Sender: TObject);
var i, j: Integer;
  maxA, maxb: Integer;
begin
//soll man das ganze StingGrid durchsuchen, oder nur den n?tigsten Teil?
  maxA:=GetStringGridExtent(StG_A, 'm'); //f?r die Matrix
  maxb:=GetStringGridExtent(StG_b, 'v'); //f?r den Vektor

  for i:= 0 to maxA do
  begin
    for j:= 0 to maxA do
    begin
      StG_A.Cells[i, j]:= '';
    end;
    StG_x.Cells[0,i]:= '';
    StG_b.Cells[0,i]:= '';
  end;
end;

procedure TForm1.Beenden_Click(Sender: TObject);
begin
  Close;
end;

procedure TForm1.Kopiermodus1Click(Sender: TObject);
begin
  if Kopiermodus1.Checked then begin
    StG_A.Options:=[goVertLine,goHorzLine,goRangeSelect,goDrawFocusSelected];
    StG_x.Options:=[goVertLine,goHorzLine,goRangeSelect,goDrawFocusSelected];
    StG_b.Options:=[goVertLine,goHorzLine,goRangeSelect,goDrawFocusSelected];
  end else begin
    StG_A.Options:=[goVertLine,goHorzLine,goRangeSelect,goDrawFocusSelected,goEditing];
    StG_x.Options:=[goVertLine,goHorzLine,goRangeSelect,goDrawFocusSelected,goEditing];
    StG_b.Options:=[goVertLine,goHorzLine,goRangeSelect,goDrawFocusSelected,goEditing];
  end;
end;

procedure TForm1.EinfgenClick(Sender: TObject);
begin
  //
end;

procedure TForm1.MatrixSpaltenbreite1Click(Sender: TObject);
var
  Input, Title, Prmpt, Deflt: string;
  NeWid: Integer;
  ClickedOK: Boolean;
begin
  Title:= 'Spaltenbreite der Matrix A';
  Prmpt:= 'Geben Sie f?r die Spaltenbreite der Matrix A einen g?ltigen Integerwert ein. Geben Sie d f?r Defaultwert ein, oder o f?r die optimale Spaltenbreite:';
  Input:= IntToStr(StG_A.DefaultColWidth);
  ClickedOK:= InputQuery(Title, Prmpt, Input);
  if not ClickedOk then exit; 
  if Input = 'd' then NeWid:=64 else
  if Input = 'o' then NeWid:=20 {GetOptimColWidth}
    else NeWid:= StrToInt(Input);
  StG_A.DefaultColWidth:=NeWid;
end;

procedure TForm1.Einheitsmatrixgenerieren1Click(Sender: TObject);
var
  Title, Prmpt, Input: string;
  Extnt, i: Integer;
  ClickedOK: Boolean;
begin
  Title:= 'Einheitsmatrix generieren';
  Prmpt:= 'Geben Sie f?r die Gr??e der Quadratischen Einheitsmatrix einen g?ltigen Integerwert ein:';
  Input:= IntToStr(50);
  ClickedOK:= InputQuery(Title, Prmpt, Input);
  if ClickedOK then
    Extnt:= StrToInt(Input)
  else exit;  

  for i:= 0 to Extnt-1 do
  begin
    StG_A.Cells[i, i]:= IntToStr(1);
    StG_b.Cells[0, i]:= IntToStr(i+1);
  end;
end;

procedure TForm1.Info1Click(Sender: TObject);
var mesag, Versn: String;
begin
  Versn:='2004.1.1.1';
  mesag:='GPMXS Gauss Pyramid Matrix Solver' + #13#10 + 'Version: ' + Versn + #13#10 + 'Viel Spa? mit dem Programm w?nscht der Autor Oliver Meyer';
  MessageDlg(mesag, mtInformation, [mbOK], 0);
end;

procedure TForm1.Matrixvergrern1Click(Sender: TObject);
var
  Title, Prmpt, Input: string;
  Extnt, i: Integer;
  ClickedOK: Boolean;
begin
  Title:= 'Tabelle vergr??ern';
  Prmpt:= 'Geben Sie f?r die Tabelle die maximale Anzahl an Zeilen und Spalten ein:';
  Input:= IntToStr(StG_A.ColCount);
  ClickedOK:= InputQuery(Title, Prmpt, Input);
  if ClickedOK then
  begin
    StG_A.ColCount:= StrToInt(Input);
    StG_A.RowCount:= StrToInt(Input);
    StG_b.RowCount:= StrToInt(Input);
    StG_x.RowCount:= StrToInt(Input);
  end;
end;


//_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-

constructor TOMPyrMxSolv.Create(Ext: Integer);
begin
  SetExtent(Ext);
end;

procedure TOMPyrMxSolv.SetExtent(Ext: Integer);
var i: Integer;
begin
  iEx:=Ext;
  //f?r die Pyramidenmatrix Amat und den Ergebnisvektor b Speicher reservieren:
  //Pyramidenmatrix deshalb, weil von unten nach oben immer eine Zeile und Spalte
  //weggelassen wird, die sich ohnehin durch den Algorithmus zu Null ergeben soll.

  SetLength(Amat, 0);
  SetLength(b, 0);

  {
  SetLength(a2, 0);
  SetLength(a2, iEx, iEx);
  SetLength(b2, 0);
  SetLength(b2, iEx);
  }
  
  SetLength(x, 0);
  SetLength(x, iEx); //zu berechnender Vektor (Vektor der Unbekannten)


  SetLength(Amat, iEx);
  SetLength(b, iEx);

  for i:=0 to iEx-1 do
  begin
    SetLength(Amat[i], iEx-i, iEx-i);
    SetLength(b[i], iEx-i);
  end;

end;

function TOMPyrMxSolv.GetExtent: Integer;
begin
  result:= iEx;
end;

procedure TOMPyrMxSolv.Clear;
begin
  //
end;

procedure TOMPyrMxSolv.SolvePyramidMx;
var i, j, k: Integer;
    Sum: Double;
    mess: String;
    Ret: Integer;
begin
  //die Pyramidenmatrix aufstellen
  //Pyramidenmatrix deshalb, weil von unten nach oben immer die erste Zeile und Spalte
  //weggelassen wird, die sich ohnehin durch den Algorithmus zu Null ergeben soll.
  for i:= 1 to iEx-1 do //i=0 ist die ausgangsmatrix
  begin
    //?berpr?fen ob das Pivotalelement (NordWest, links oben) null ist:
    if Amat[i-1,0,0] <> 0 then
    begin
      for j:= 0 to iEx-i-1 do
      begin
        for k:= 0 to iEx-i-1 do
        begin
          Amat[i, j, k]:=Amat[i-1, j+1, k+1] - (Amat[i-1, 0, k+1] * Amat[i-1, j+1, 0]) /
                                             //----------------------------------------
                                                          Amat[i-1, 0, 0];
        end;
        //die Modifikation des Ergebnisvektors b
        b[i, j]:=b[i-1, j+1] - (Amat[i-1, 0, j+1] * b[i-1, 0]) /
                             //--------------------------------
                                        Amat[i-1, 0, 0];
      end;
    end
    else
    begin
      mess:='LGS kann nicht gel?st werden, Amat[' + IntToStr(i-1) + ', 0, 0] ist null.';
      Ret:=MessageDlg(mess, mtinformation, [mbOk, mbCancel], 0);
      if Ret=mrCancel then break;
    end;
  end;

  //den X-vektor von oben her aufstellen:
  //jetzt durchwandern wir die Matrix Amat,
  //und den Ergebnisvektor b von oben nach unten
  for i:= iEx-1 downto 0 do
  begin
    //?berpr?fen ob das Pivotelement (NordWest, links oben) null ist:
    if Amat[i,0,0] <> 0 then
    begin
      Sum:=0;
      for j:= 0 to iEx-i-1 do
      begin
        //die Summe aufstellen:
        Sum:= Sum + (-1) * Amat[i, j, 0] * x[i+j];
      end;
      x[i]:= (Sum + b[i, 0]) / Amat[i, 0, 0];
    end
    else
    begin
      mess:='LGS kann nicht gel?st werden, Amat[' + IntToStr(i) + ', 0, 0] ist null.';
      Ret:=MessageDlg(mess, mtinformation, [mbOk, mbCancel], 0);
      if Ret=mrCancel then break;
    end;
  end;
end;

//woher kommt der gleich wieder?
procedure TOMPyrMxSolv.Solve2;
var
  i, j, k: Integer; // Schleifen-Variablen
  faktor: Double; // Aufl?sungsfaktor
  summe: Double; // Summe beim Aufrollen auf der linken Seite
begin
  // Dreieck bilden
  for i:= 1 to iEx-1 do // Stufen
  begin
    for j:= iEx-1 downto i do // Gleichungen von unten nach oben
    begin
      // Pr?fen, ob Koeffizent Null ist
      if a2[i-1, j] = 0.0 then // Division durch Null vermeiden
      begin
        Continue; // Gl?ck gehabt, Koeffizent ist bereits Null
      end;
      // Letzt-vorhandene Variable durch Vorg?nger aufl?sen
      faktor:= - a2[i-1, j-1] / a2[i-1, j];
      // Gleichungen addieren und in aktueller Gleichung speichern
      for k:= i to iEx-1 do
      begin
        a2[k, j]:= faktor * a2[k, j] + a2[k, j-1];
      end;
      b2[j]     := faktor * b2[j]    + b2[j-1];
      // muss nicht addiert werden, da sowieso Null
      a2[i-1, j]:= 0.0;
    end;
  end;
  // Aufrollen
  for i:= iEx-1 downto 0 do // von unten nach oben
  begin
    // Summe der bekannten Gr??en auf der linken Seite bilden
    summe:= 0.0;
    for j:= iEx-1 downto i+1 do // Variablen einsetzen
    begin
      summe:= summe + a2[j, i] * x[j];
    end;
    // Summe auf die rechte Seite bringen
    summe:= b2[i] - summe;
    // Variable ermitteln
    x[i]:= summe / a2[i, i];
  end;
end;

{
procedure TOMPyrMxSolv.SolveSinglValDecomp(var a: glmpbynp; m,n,mp,np: integer; var w: glnparray; var v: glnpbynp);
(* Programs using routine SVDCMP must define the types
type
   glnparray = array [1..np] of Double;
   glmpbynp = array [1..mp,1..np] of Double;
   glnpbynp = array [1..np,1..np] of Double;
in the main routine. *)
label 1,2,3;
const
   nmax=100;
var
   nm,l,k,j,jj,its,i: integer;
   z,y,x,scale,s,h,g,f,c,anorm: Double;
   rv1: array [1..nmax] of Double;
function sign(a,b: Double): Double;
   begin
      if (b >= 0.0) then sign := abs(a) else sign := -abs(a)
   end;
function max(a,b: Double): Double;
   begin
      if (a > b) then max := a else max := b
   end;
begin
   g := 0.0;
   scale := 0.0;
   anorm := 0.0;
   for i := 1 to n do begin
      l := i+1;
      rv1[i] := scale*g;
      g := 0.0;
      s := 0.0;
      scale := 0.0;
      if (i <= m) then begin
         for k := i to m do begin
            scale := scale+abs(a[k,i])
         end;
         if (scale <> 0.0) then begin
            for k := i to m do begin
               a[k,i] := a[k,i]/scale;
               s := s+a[k,i]*a[k,i]
            end;
            f := a[i,i];
            g := -sign(sqrt(s),f);
            h := f*g-s;
            a[i,i] := f-g;
            if (i <> n) then begin
               for j := l to n do begin
                  s := 0.0;
                  for k := i to m do begin
                     s := s+a[k,i]*a[k,j]
                  end;
                  f := s/h;
                  for k := i to m do begin
                     a[k,j] := a[k,j]+
                        f*a[k,i]
                  end
               end
            end;
            for k := i to m do begin
               a[k,i] := scale*a[k,i]
            end
         end
      end;
      w[i] := scale*g;
      g := 0.0;
      s := 0.0;
      scale := 0.0;
      if ((i <= m) AND (i <> n)) then begin
         for k := l to n do begin
            scale := scale+abs(a[i,k])
         end;
         if (scale <> 0.0) then begin
            for k := l to n do begin
               a[i,k] := a[i,k]/scale;
               s := s+a[i,k]*a[i,k]
            end;
            f := a[i,l];
            g := -sign(sqrt(s),f);
            h := f*g-s;
            a[i,l] := f-g;
            for k := l to n do begin
               rv1[k] := a[i,k]/h
            end;
            if (i <> m) then begin
               for j := l to m do begin
                  s := 0.0;
                  for k := l to n do begin
                     s := s+a[j,k]*a[i,k]
                  end;
                  for k := l to n do begin
                     a[j,k] := a[j,k]
                        +s*rv1[k]
                  end
               end
            end;
            for k := l to n do begin
               a[i,k] := scale*a[i,k]
            end
         end
      end;
      anorm := max(anorm,(abs(w[i])+abs(rv1[i])))
   end;
   for i := n doWNto 1 do begin
      if (i < n) then begin
         if (g <> 0.0) then begin
            for j := l to n do begin
               v[j,i] := (a[i,j]/a[i,l])/g
            end;
            for j := l to n do begin
               s := 0.0;
               for k := l to n do begin
                  s := s+a[i,k]*v[k,j]
               end;
               for k := l to n do begin
                  v[k,j] := v[k,j]+s*v[k,i]
               end
            end
         end;
         for j := l to n do begin
            v[i,j] := 0.0;
            v[j,i] := 0.0
         end
      end;
      v[i,i] := 1.0;
      g := rv1[i];
      l := i
   end;
   for i := n doWNto 1 do begin
      l := i+1;
      g := w[i];
      if (i < n) then begin
         for j := l to n do begin
            a[i,j] := 0.0
         end
      end;
      if (g <> 0.0) then begin
         g := 1.0/g;
         if (i <> n) then begin
            for j := l to n do begin
               s := 0.0;
               for k := l to m do begin
                  s := s+a[k,i]*a[k,j]
               end;
               f := (s/a[i,i])*g;
               for k := i to m do begin
                  a[k,j] := a[k,j]+f*a[k,i]
               end
            end
         end;
         for j := i to m do begin
            a[j,i] := a[j,i]*g
         end
      end else begin
         for j := i to m do begin
            a[j,i] := 0.0
         end
      end;
      a[i,i] := a[i,i]+1.0
   end;
   for k := n doWNto 1 do begin
      for its := 1 to 30 do begin
         for l := k doWNto 1 do begin
            nm := l-1;
            if ((abs(rv1[l])+anorm) = anorm) then goto 2;
            if ((abs(w[nm])+anorm) = anorm) then goto 1
         end;
1:         c := 0.0;
         s := 1.0;
         for i := l to k do begin
            f := s*rv1[i];
            if ((abs(f)+anorm) <> anorm) then begin
               g := w[i];
               h := sqrt(f*f+g*g);
               w[i] := h;
               h := 1.0/h;
               c := (g*h);
               s := -(f*h);
               for j := 1 to m do begin
                  y := a[j,nm];
                  z := a[j,i];
                  a[j,nm] := (y*c)+(z*s);
                  a[j,i] := -(y*s)+(z*c)
               end
            end
         end;
2:         z := w[k];
         if (l = k) then begin
            if (z < 0.0) then begin
               w[k] := -z;
               for j := 1 to n do begin
               v[j,k] := -v[j,k]
            end
         end;
         goto 3
         end;
         if (its = 30) then begin
            writeln ('no convergence in 30 SVDCMP iterations'); readln
         end;
         x := w[l];
         nm := k-1;
         y := w[nm];
         g := rv1[nm];
         h := rv1[k];
         f := ((y-z)*(y+z)+(g-h)*(g+h))/(2.0*h*y);
         g := sqrt(f*f+1.0);
         f := ((x-z)*(x+z)+h*((y/(f+sign(g,f)))-h))/x;
         c := 1.0;
         s := 1.0;
         for j := l to nm do begin
            i := j+1;
            g := rv1[i];
            y := w[i];
            h := s*g;
            g := c*g;
            z := sqrt(f*f+h*h);
            rv1[j] := z;
            c := f/z;
            s := h/z;
            f := (x*c)+(g*s);
            g := -(x*s)+(g*c);
            h := y*s;
            y := y*c;
            for jj := 1 to n do begin
               x := v[jj,j];
               z := v[jj,i];
               v[jj,j] := (x*c)+(z*s);
               v[jj,i] := -(x*s)+(z*c)
            end;
            z := sqrt(f*f+h*h);
            w[j] := z;
            if (z <> 0.0) then begin
               z := 1.0/z;
               c := f*z;
               s := h*z
            end;
            f := (c*g)+(s*y);
            x := -(s*g)+(c*y);
            for jj := 1 to m do begin
               y := a[jj,j];
               z := a[jj,i];
               a[jj,j] := (y*c)+(z*s);
               a[jj,i] := -(y*s)+(z*c)
            end
         end;
         rv1[l] := 0.0;
         rv1[k] := f;
         w[k] := x
      end;
3:   end
end;
}


procedure TForm1.mnuSelectMethodClick(Sender: TObject);
begin
  FrmSelectMethod.ShowModal;
  if WhichS = TsPyramidMxSolve then
  begin

  end;
  if  WhichS = TsSparseSolve then
  begin

  end;
end;

end.
