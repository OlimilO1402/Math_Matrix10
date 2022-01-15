program Project1;

uses
  Forms,
  Unit1 in 'Unit1.pas' {Form1},
  UClsSparseSolv in 'UClsSparseSolv.pas',
  Unit2 in 'Unit2.pas' {FrmSelectMethod};

{$R *.res}

begin
  Application.Initialize;
  Application.CreateForm(TForm1, Form1);
  Application.CreateForm(TFrmSelectMethod, FrmSelectMethod);
  Application.Run;
end.
