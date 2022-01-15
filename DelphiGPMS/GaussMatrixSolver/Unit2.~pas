unit Unit2;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ExtCtrls, Unit1;

type
  TSelectSolver = (TsPyramidMxSolve, TsSparseSolve);
  TFrmSelectMethod = class(TForm)
    RGSelectCalcMethod: TRadioGroup;
    BtnOK: TButton;
    Button1: TButton;
    procedure BtnOKClick(Sender: TObject);
  private
    { Private-Deklarationen }
  public
    { Public-Deklarationen }
  end;

var
  FrmSelectMethod: TFrmSelectMethod;
  WhichS: TSelectSolver;

implementation

{$R *.dfm}

procedure TFrmSelectMethod.BtnOKClick(Sender: TObject);
begin
  if RGSelectCalcMethod.ItemIndex = 0 then
  begin
    WhichS:= TsPyramidMxSolve;
  end
  else
  begin
    WhichS:= TsSparseSolve;
  end;
end;

end.
