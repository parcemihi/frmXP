unit ufrmSettings;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls;

type
  TfrmSettings = class(TForm)
    cmdOK: TButton;
    GroupBox2: TGroupBox;
    chkExecuteObstacles: TCheckBox;
    chkExecuteSeparation: TCheckBox;
    chkExecuteAlignment: TCheckBox;
    chkExecuteCohesion: TCheckBox;
    chkObstacleAvoidance: TCheckBox;
    chkBoidAvoidance: TCheckBox;
    chkShowAlignment: TCheckBox;
    chkShowCohesion: TCheckBox;
    Label1: TLabel;
    txtCheckMap: TEdit;
    Label2: TLabel;
    procedure cmdOKClick(Sender: TObject);
    procedure chkObstacleAvoidanceClick(Sender: TObject);
    procedure chkBoidAvoidanceClick(Sender: TObject);
    procedure chkShowAlignmentClick(Sender: TObject);
    procedure chkShowCohesionClick(Sender: TObject);
    procedure chkExecuteObstaclesClick(Sender: TObject);
    procedure chkExecuteSeparationClick(Sender: TObject);
    procedure chkExecuteAlignmentClick(Sender: TObject);
    procedure chkExecuteCohesionClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure txtCheckMapExit(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmSettings: TfrmSettings;

implementation

uses ufrmBoids;

{$R *.DFM}

//******************************************************************************
procedure TfrmSettings.cmdOKClick(Sender: TObject);
begin
  Close;
end;

//******************************************************************************
procedure TfrmSettings.chkObstacleAvoidanceClick(Sender: TObject);
begin
  frmBoids.BoidEngine.bObstacleAvoidance :=
    chkObstacleAvoidance.Checked;

  frmBoids.BoidEngine.CopySettingsToAllBoids;
end;

//******************************************************************************
procedure TfrmSettings.chkBoidAvoidanceClick(Sender: TObject);
begin
  frmBoids.BoidEngine.bBoidAvoidance :=
    chkBoidAvoidance.Checked;

  frmBoids.BoidEngine.CopySettingsToAllBoids;
end;

//******************************************************************************
procedure TfrmSettings.chkShowAlignmentClick(Sender: TObject);
begin
  frmBoids.BoidEngine.bShowAlignment :=
    chkShowAlignment.Checked;

  frmBoids.BoidEngine.CopySettingsToAllBoids;
end;

//******************************************************************************
procedure TfrmSettings.chkShowCohesionClick(Sender: TObject);
begin
  frmBoids.BoidEngine.bShowCohesion :=
    chkShowCohesion.Checked;

  frmBoids.BoidEngine.CopySettingsToAllBoids;
end;

//******************************************************************************
procedure TfrmSettings.chkExecuteObstaclesClick(Sender: TObject);
begin
  frmBoids.BoidEngine.bExecuteObstacles :=
    chkExecuteObstacles.Checked;

  frmBoids.BoidEngine.CopySettingsToAllBoids;
end;

//******************************************************************************
procedure TfrmSettings.chkExecuteSeparationClick(Sender: TObject);
begin
  frmBoids.BoidEngine.bExecuteSeparation :=
    chkExecuteSeparation.Checked;

  frmBoids.BoidEngine.CopySettingsToAllBoids;
end;

//******************************************************************************
procedure TfrmSettings.chkExecuteAlignmentClick(Sender: TObject);
begin
  frmBoids.BoidEngine.bExecuteAlignment :=
    chkExecuteAlignment.Checked;

  frmBoids.BoidEngine.CopySettingsToAllBoids;
end;

//******************************************************************************
procedure TfrmSettings.chkExecuteCohesionClick(Sender: TObject);
begin
  frmBoids.BoidEngine.bExecuteCohesion :=
    chkExecuteCohesion.Checked;

  frmBoids.BoidEngine.CopySettingsToAllBoids;
end;

//******************************************************************************
procedure TfrmSettings.FormShow(Sender: TObject);
begin
  chkExecuteCohesion.Checked := frmBoids.BoidEngine.bExecuteCohesion;
  chkExecuteObstacles.Checked := frmBoids.BoidEngine.bExecuteObstacles;
  chkExecuteAlignment.Checked := frmBoids.BoidEngine.bExecuteAlignment;
  chkExecuteSeparation.Checked := frmBoids.BoidEngine.bExecuteSeparation;

  chkExecuteObstacles.Checked := frmBoids.BoidEngine.bExecuteObstacles;
  chkShowAlignment.Checked := frmBoids.BoidEngine.bShowAlignment;
  chkShowCohesion.Checked := frmBoids.BoidEngine.bShowCohesion;
  chkBoidAvoidance.Checked := frmBoids.BoidEngine.bBoidAvoidance;
end;

//******************************************************************************
procedure TfrmSettings.txtCheckMapExit(Sender: TObject);
begin
  frmBoids.BoidEngine.CheckChancePerFrame := StrToInt(txtCheckMap.Text);
end;

//******************************************************************************
end.
