program Test1;
 {a SparSolv example to test performance and memory use}

{$N+,G+}                          {enable 87 and 286 instructions}

{Next section is needed to make sure that a text window is opened when
either test program is run from Windows.  For 16-bit Windows, you need
the WinCrt unit.  For Delphi 2.0 and above, you need the APPTYPE CONSOLE
 directive}

{$IFDEF VER80}{$DEFINE DELPHI16}{$ENDIF} {Delphi 1 only}
{$IFDEF VER90}{$DEFINE DELPHI32}{$ENDIF} {Delphi 2 only}
{$IFDEF VER100}{$DEFINE DELPHI32}{$ENDIF} {Delphi 3 only}
{$IFDEF VER120}{$DEFINE DELPHI32}{$ENDIF} {Delphi 4 only}
{$IFDEF VER130}{$DEFINE DELPHI32}{$ENDIF} {Delphi 5 only}
{$IFDEF VER140}{$DEFINE DELPHI32}{$ENDIF} {Delphi 6 only}

{$IFDEF DELPHI32}{$DEFINE BIT32}{$ELSE}{$DEFINE BIT16}{$ENDIF}
{$IFDEF DELPHI32}{$DEFINE DELPHI}{$ENDIF}
{$IFDEF DELPHI16}{$DEFINE DELPHI}{$ENDIF}

{$IFDEF DELPHI32}{$APPTYPE CONSOLE}{$ENDIF} {Delphi 2,3,4,5 only}

{$IFDEF DELPHI}
uses SparSolv,SysUtils;
 {$ELSE}   {TPW BPW or Turbo Pascal}
    {$IFDEF Windows}
uses SparSolv,WinCrt;  {TPW BPW}
    {$ELSE}
uses SparSolv;             {DOS}
    {$ENDIF}
 {$ENDIF}


{$IFDEF DELPHI}
function Seconds: Double;
var Hour, Min, Sec, MSec: Word;
begin
  DecodeTime(Now, Hour, Min, Sec, MSec);
  Seconds := 3600 * Hour + 60 * Min + sec + MSec / 1000;
end;
{$ELSE}
function Seconds: Double;
begin
  Seconds := 0.0; {timer only works under Delphi}
end;
{$ENDIF}



const EntriesPerRow = 6;
var
 Reason   : String;
 ErrNo1, ErrNo2, ErrNo3 : Integer;
 Criterion, Density : Single;
 Col, Row, N : Integer;
 Count    : LongInt;
 InitMemUsed : LongInt;
 Total, Value : Double;
 Start, Stop : Double;
 ErrSum   : Double;
label Fail, EndProg;
begin



    {$IFNDEF DELPHI32}
 WriteLn('Initial Heap Size':30, MemAvail:12);
    {$ENDIF}
 WriteLn('Setting up matrix...');

 N := 1000;                        {size of matrix - set larger if you want}
 Count := 0;
 if not InitStruc(N) then goto Fail;

 {construct a matrix with random elements}
 {RHS is set so that variable V will have value V}
 {diagonal elements are all filled to prevent singularity}
 Start := Seconds;
 for Row := 1 to N do begin
  Total := 0.0;
  for Col := N downto 1 {1 to N} do begin
   Criterion := (EntriesPerRow - 1) * (Row + Col) / (LongInt(N) * N);
   {this causes some rows and cols to be fuller than others}
   if ((Row = Col) or (Random < Criterion)) then begin
    Inc(Count);
    Value := N * Random + Row + Col;
    if not AddLHS(Row, Col, Value) then goto Fail;
    Total := Total + Col * Value;
   end;
  end;
  if not AddRHS(Row, Total) then goto Fail;
 end;                             {for Row:= 1 TO N}
 Stop := Seconds;
 Density := 100.0 * (Count / N) / N;
 InitMemUsed := SparMemUsed;

 WriteLn('N':30, N:12);
 WriteLn('Count':30, Count:12);
 WriteLn('Density':30, Density:12:3, '%');
 WriteLn('Vars per Row':30, (Count / N):12:1);
 WriteLn('Initial Memory Used':30, InitMemUsed:12);
 WriteLn('Load time':30, (Stop - Start):12:1, ' seconds');
 {pause to read information so far}
 WriteLn('hit return to go on'); ReadLn;

 WriteLn('Solving...Please Wait');
 Start := Seconds;
 if not Solve1 then goto Fail;
 Stop := Seconds;
 ErrSum := 0.0;
 for Row := 1 to N do begin
  if not GetAnswer(Row, Value) then goto Fail;
  ErrSum := ErrSum + Abs(Value - Row);
  if ((Row = 1) or (Row = N) or ((Row mod (N div 10)) = 0)) then
   WriteLn(Row:5, Value:15:5);
 end;
 WriteLn('Solve time':30, (Stop - Start):12:1, ' seconds');
 WriteLn('Maximum Memory Used':30, MaxMemUsed:12);
 WriteLn('Memory Growth Factor':30, (MaxMemUsed / InitMemUsed):12:1);
 WriteLn('Sum of Errors:':30, ErrSum:12:5);
 goto EndProg;

Fail:
 GetErrorMsg(Reason, ErrNo1, ErrNo2, ErrNo3);
 WriteLn('Failed:  Error', ErrNo1:0, ' ', Reason, ' ', ' ', ErrNo2:3, ' ', ErrNo3:3);
EndProg:
 ReleaseStruc;
 WriteLn('Memory Used after ReleaseStruc':30, SparMemUsed:12);
    {$IFNDEF DELPHI32}
 WriteLn('Final Heap Size':30, MemAvail:12);
    {$ENDIF}
 {prevent DOS box closing before you can read it}
 WriteLn('hit return to finish'); ReadLn;
end.

