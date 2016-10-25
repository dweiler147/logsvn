program Boletec;

uses
  Forms,
  uBoletec in 'uBoletec.pas' {FrPrincipal};

{$R *.res}

begin
  Application.Initialize;
  Application.CreateForm(TFrPrincipal, FrPrincipal);
  Application.Run;
end.
