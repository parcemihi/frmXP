program MFBoids;

uses
  Forms,
  ufrmBoids in 'ufrmBoids.pas' {frmBoids},
  uBoids in 'uBoids.pas',
  uBoidEngine in 'uBoidEngine.pas',
  ufrmAbout in 'ufrmAbout.pas' {frmAbout},
  StrFunctions in 'StrFunctions.pas',
  uTMovable in 'uTMovable.pas',
  uTMovableEngine in 'uTMovableEngine.pas',
  ufrmSettings in 'ufrmSettings.pas' {frmSettings};

{$R *.RES}

begin
  Application.Initialize;
  Application.Title := 'MatFa''s Boids 0.3';
  Application.CreateForm(TfrmBoids, frmBoids);
  Application.CreateForm(TfrmAbout, frmAbout);
  Application.CreateForm(TfrmSettings, frmSettings);
  Application.Run;
end.
