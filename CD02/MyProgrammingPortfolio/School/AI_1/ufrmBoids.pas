unit ufrmBoids;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ExtCtrls, StdCtrls, Menus, ComCtrls, IniFiles,StrFunctions,
  uBoids, uBoidEngine, uTMovable;

type
  TfrmBoids = class(TForm)
    Panel1: TPanel;
    lblBoidsPerSecond: TLabel;
    tmrBoidCount: TTimer;
    Label1: TLabel;
    Label2: TLabel;
    txtNumberOfBoids: TEdit;
    MainMenu1: TMainMenu;
    File1: TMenuItem;
    Exit1: TMenuItem;
    Startboids1: TMenuItem;
    Stopboids1: TMenuItem;
    Help1: TMenuItem;
    About1: TMenuItem;
    Causesometurmoil1: TMenuItem;
    N2: TMenuItem;
    Obstacles1: TMenuItem;
    Clearallobstacles1: TMenuItem;
    chkObstacle: TCheckBox;
    Hideobstacles1: TMenuItem;
    Panel2: TPanel;
    Label3: TLabel;
    tbMaxSpeed: TTrackBar;
    Label4: TLabel;
    tbSpeedChange: TTrackBar;
    Label5: TLabel;
    tbSensorRange: TTrackBar;
    Label6: TLabel;
    tbOptimalDistance: TTrackBar;
    Label7: TLabel;
    tbTooClose: TTrackBar;
    Label9: TLabel;
    tbTurnRate: TTrackBar;
    tbLineLength: TTrackBar;
    Label10: TLabel;
    cmdAdjustBoidCount: TButton;
    cmdSaveSettings: TButton;
    cboBoidStyles: TComboBox;
    cmdEraseStyle: TButton;
    chkBlanking: TCheckBox;
    Image1: TImage;
    Label8: TLabel;
    tbStayInCenter: TTrackBar;
    chkDraw: TCheckBox;
    Label11: TLabel;
    lblFrameCount: TLabel;
    cmdTurmoil: TButton;
    cmdSettings: TButton;
    procedure cmdStartClick(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure FormShow(Sender: TObject);
    procedure cmdStopClick(Sender: TObject);
    procedure tmrBoidCountTimer(Sender: TObject);
    procedure Image1MouseMove(Sender: TObject; Shift: TShiftState; X,
      Y: Integer);
    procedure Image1Click(Sender: TObject);
    procedure Image1DblClick(Sender: TObject);
    procedure chkObstacleClick(Sender: TObject);
    procedure About1Click(Sender: TObject);
    procedure Exit1Click(Sender: TObject);
    procedure Clearallobstacles1Click(Sender: TObject);
    procedure Causesometurmoil1Click(Sender: TObject);
    procedure Hideobstacles1Click(Sender: TObject);
    procedure tbSpeedChangeChange(Sender: TObject);
    procedure tbMaxSpeedChange(Sender: TObject);
    procedure tbSensorRangeChange(Sender: TObject);
    procedure tbOptimalDistanceChange(Sender: TObject);
    procedure tbTooCloseChange(Sender: TObject);
    procedure tbTurnRateChange(Sender: TObject);
    procedure tbLineLengthChange(Sender: TObject);
    procedure cmdAdjustBoidCountClick(Sender: TObject);
    procedure cmdSaveSettingsClick(Sender: TObject);
    procedure cboBoidStylesChange(Sender: TObject);
    procedure cmdEraseStyleClick(Sender: TObject);
    procedure chkBlankingClick(Sender: TObject);
    procedure tbStayInCenterChange(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure chkDrawClick(Sender: TObject);
    procedure Panel1Click(Sender: TObject);
    procedure cmdSettingsClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
    bRun                : boolean;
    BoidEngine          : TBoidEngine;
    procedure ReadSettingsFromIniToTrackBars;
    procedure ReadSettingsFromTrackBars;
    procedure SaveSettingsFromTrackBarsToIni;
    procedure StartBoidEngine(bDirect : boolean);
  end;

  TCursorBoid = class(TBoid)
    function PrepareToMove(ClosestBoids : TList;ClosestObstacle : TMovable; Canvas : TCanvas) : boolean;override;
    procedure Move(Canvas : TCanvas);override;
    procedure TurnLeft(delta : single);override;
    procedure TurnRight(delta : single);override;
  end;

  procedure ClearArena(Canvas : TCanvas);


var
  frmBoids            : TfrmBoids;
  iBoidPerSecondCount : integer;
  iGenerationCounter	: integer;
  bPlaceObstacle			: boolean;
  MouseX,MouseY       : integer;
  bStarted						: boolean;
  CursorBoid          : TCursorBoid;
  IniFile							: TIniFile;
  OutputCanvas	 			: TCanvas;
  bDirectDraw					: boolean;

implementation

uses ufrmAbout, ufrmSettings;

{$R *.DFM}
//******************************************************************************
function TCursorBoid.PrepareToMove(ClosestBoids : TList;ClosestObstacle : TMovable; Canvas : TCanvas):boolean;
begin
end;

//******************************************************************************
procedure TCursorBoid.Move(Canvas : TCanvas);
begin
  x := MouseX;
  y := MouseY;
end;

//******************************************************************************
procedure TCursorBoid.TurnLeft(delta : single);
begin
  Dir := Dir - delta;
end;

//******************************************************************************
procedure TCursorBoid.TurnRight(delta : single);
begin
  Dir := Dir + delta;
end;

//******************************************************************************
procedure ClearArena(Canvas : TCanvas);
begin
	Canvas.Pen.Color := clBlack;
  Canvas.Brush.Color := clBlack;
	Canvas.FillRect(Canvas.ClipRect);
end;

//******************************************************************************
procedure TfrmBoids.ReadSettingsFromTrackBars;
var
	i : integer;
begin
  for i := 0 to ComponentCount -1 do
     if (Components[I] is TTrackBar) then
				TTrackBar(Components[i]).OnChange(nil);
end;

//******************************************************************************
procedure TfrmBoids.ReadSettingsFromIniToTrackBars;
var
	i : integer;
  iObsCount 			: integer;
  sSectionName 		: string;
  sObsString			: string;
  x,y,size				: single;
  iBCount					: integer;
begin
  sSectionName := cboBoidStyles.Text;

  for i := 0 to ComponentCount -1 do
  	if (Components[I] is TTrackBar) then
	    TTrackBar(Components[I]).Position :=
      	IniFile.ReadInteger(sSectionName,TTrackBar(Components[I]).Name,TTrackBar(Components[I]).Position);

  ReadSettingsFromTrackBars;

  with BoidEngine do
  begin
    bBoidAvoidance      := IniFile.ReadBool(sSectionName,'bBoidAvoidance',false);
    bObstacleAvoidance  := IniFile.ReadBool(sSectionName,'bObstacleAvoidance',false);
    bShowAlignment     := IniFile.ReadBool(sSectionName,'bShowAlignment',false);
    bShowCohesion     := IniFile.ReadBool(sSectionName,'bShowCohesion',false);

    bExecuteObstacles   := IniFile.ReadBool(sSectionName,'bExecuteObstacles',true);
    bExecuteSeparation  := IniFile.ReadBool(sSectionName,'bExecuteSeparation',true);
    bExecuteAlignment   := IniFile.ReadBool(sSectionName,'bExecuteAlignment',true);
    bExecuteCohesion    := IniFile.ReadBool(sSectionName,'bExecuteCohesion',true);
  end;

  BoidEngine.CopySettingsToAllBoids;

  iBCount := IniFile.ReadInteger(sSectionName,'boidcount',25);
  BoidEngine.AdjustMovableCount(iBCount);

  txtNumberOfBoids.Text := inttostr(iBCount);

  iObsCount := IniFile.ReadInteger(cboBoidStyles.Text,
  	'obstacle count', 0);

  if bRun and (iObsCount > 0) then with BoidEngine do
  begin
		while ObstacleList.Count <> 0 do
  	begin
  		TObstacle(ObstacleList[ObstacleList.Count-1]).Destroy;
      ObstacleList.Delete(ObstacleList.Count-1);
	  end;

    for i := 1 to iObsCount do
    begin
      sObsString := IniFile.ReadString(cboBoidStyles.Text,
		  	'obstacle ' + IntToStr(i-1),'1,1');

      x := StrToFloat(GetBefore(',',sObsString));
      sObsString := GetAfter(',',sObsString);
			y := StrToFloat(GetBefore(',',sObsString));
      size := StrToFloat(GetAfter(',',sObsString));

      ObstacleList.Add(TObstacle.Create(trunc(x),trunc(y),trunc(size),clYellow,Canvas));
    end;
  end;


end;

//******************************************************************************
procedure TfrmBoids.SaveSettingsFromTrackBarsToIni;
var
	i : integer;
  sSectionName : string;
begin
  sSectionName := cboBoidStyles.Text;

  for i := 0 to ComponentCount -1 do
	 	if (Components[I] is TTrackBar) then
     	IniFile.WriteInteger(sSectionName,TTrackBar(Components[I]).Name,TTrackBar(Components[I]).Position);

  IniFile.WriteInteger(sSectionName,'boidcount',BoidEngine.MovableList.Count-1);

  with BoidEngine do
  begin
    IniFile.WriteBool(sSectionName,'bBoidAvoidance',bBoidAvoidance);
    IniFile.WriteBool(sSectionName,'bObstacleAvoidance',bObstacleAvoidance);
    IniFile.WriteBool(sSectionName,'bShowAlignment',bShowAlignment);
    IniFile.WriteBool(sSectionName,'bShowCohesion',bShowCohesion);

    IniFile.WriteBool(sSectionName,'bExecuteObstacles',bExecuteObstacles);
    IniFile.WriteBool(sSectionName,'bExecuteSeparation',bExecuteSeparation);
    IniFile.WriteBool(sSectionName,'bExecuteAlignment',bExecuteAlignment);
    IniFile.WriteBool(sSectionName,'bExecuteCohesion',bExecuteCohesion);
  end;

  if bRun then
  begin
	  for i := 0 to BoidEngine.ObstacleList.Count - 1 do
    begin
      IniFile.WriteString(cboBoidStyles.Text,
		  	'obstacle ' + IntToStr(i),
	  		FloatToStr(TObstacle(BoidEngine.ObstacleList[i]).X)+','+
        FloatToStr(TObstacle(BoidEngine.ObstacleList[i]).Y) + ',' +
        FloatToStr(TObstacle(BoidEngine.ObstacleList[i]).Size) );
    end;
  end;
  IniFile.WriteInteger(cboBoidStyles.Text,
  	'obstacle count',
	  BoidEngine.ObstacleList.Count);
end;

//******************************************************************************
procedure TfrmBoids.StartBoidEngine(bDirect : boolean);
begin
	if frmBoids.bRun then exit;
  bDirectDraw := bDirect;

{	if bDirectDraw then
	  OutputCanvas := frmDirectDraw.DGCScreen1.Back.Canvas
  else //}
	  OutputCanvas := frmBoids.Image1.Canvas;

  BoidEngine := TBoidEngine.Create(StrToInt(frmBoids.txtNumberOfBoids.Text),OutputCanvas);
  BoidEngine.Application := Application;

  frmBoids.chkObstacleClick(nil);

  frmBoids.bRun := true;

  frmBoids.ReadSettingsFromIniToTrackBars;

  iBoidPerSecondCount := 0;
  Canvas.Pen.Width := 1;

  while frmBoids.bRun do
  begin
{  	if bDirectDraw then
	  	OutputCanvas := frmDirectDraw.DGCScreen1.Back.Canvas
    else}
	  	OutputCanvas := frmBoids.Image1.Canvas;

	  inc(iGenerationCounter);
    iBoidPerSecondCount := iBoidPerSecondCount + BoidEngine.MovableList.Count;
    BoidEngine.RunStep;
{    if bDirectDraw then
    begin
    	//This must be called to release the device context.
	    frmDirectDraw.DGCScreen1.Back.Canvas.Release;
  	  frmDirectDraw.DGCScreen1.Flip;
    end;}
    Application.ProcessMessages;
  end;

  BoidEngine.Destroy;
  CursorBoid := TCursorBoid.Create(nil);
end;

//******************************************************************************
procedure TfrmBoids.cmdStartClick(Sender: TObject);
begin
	StartBoidEngine(false);
end;

//******************************************************************************
procedure TfrmBoids.cmdStopClick(Sender: TObject);
begin
  bRun := false;
end;

//******************************************************************************
procedure TfrmBoids.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
begin
	Halt;
end;

//******************************************************************************
procedure TfrmBoids.FormShow(Sender: TObject);
begin
  CursorBoid := TCursorBoid.Create(nil);
  IniFile := TIniFile.Create(ExtractFilePath(Application.ExeName)+'boids.ini');

  IniFile.ReadSections(cboBoidStyles.Items);
  cboBoidStyles.ItemIndex := 0;
end;

//******************************************************************************
procedure TfrmBoids.tmrBoidCountTimer(Sender: TObject);
begin
	if not bStarted then
  begin
		// Set off the boids!
    StartBoidEngine(false);
  	bStarted := true;
  end;

  if not bDirectDraw then
  begin
	  lblBoidsPerSecond.Caption := IntToStr(iBoidPerSecondCount div 2);
    lblFrameCount.Caption := IntToStr(iGenerationCounter div 2);
  end;

  iBoidPerSecondCount := 0;

  if (BoidEngine <> nil) then iGenerationCounter := 0;
end;

//******************************************************************************
procedure TfrmBoids.Image1MouseMove(Sender: TObject; Shift: TShiftState; X,
  Y: Integer);
begin
  MouseX := x;
  MouseY := y;
end;

//******************************************************************************
procedure TfrmBoids.Image1Click(Sender: TObject);
var
	NewBoid : TBoid;
  NewObstacle : TObstacle;
begin
  CursorBoid.TurnLeft(1);
  if BoidEngine = nil then exit;

	NewObstacle := TObstacle.Create(MouseX,MouseY,20,clYellow,Canvas);
  BoidEngine.ObstacleList.Add(NewObstacle);

  bPlaceObstacle := false;
end;

//******************************************************************************
procedure TfrmBoids.Image1DblClick(Sender: TObject);
begin
  CursorBoid.TurnRight(1);
  CursorBoid.TurnRight(1);
end;

//******************************************************************************
procedure TfrmBoids.cmdAdjustBoidCountClick(Sender: TObject);
var
  iNewCount : integer;
begin
	if not bRun then exit;
  iNewCount := StrToInt(txtNumberOfBoids.Text);
	BoidEngine.AdjustMovableCount(iNewCount);
end;

//******************************************************************************
procedure TfrmBoids.chkObstacleClick(Sender: TObject);
var
	i : integer;
begin
	if (BoidEngine <> nil) then
  	for i := 0 to BoidEngine.ObstacleList.Count - 1 do
	  TObstacle(BoidEngine.ObstacleList[i]).bActive := (chkObstacle.State = cbChecked);
end;

//******************************************************************************
procedure TfrmBoids.About1Click(Sender: TObject);
begin
	frmAbout.Show;
end;

//******************************************************************************
procedure TfrmBoids.Exit1Click(Sender: TObject);
begin
	Halt;
end;

//******************************************************************************
procedure TfrmBoids.Clearallobstacles1Click(Sender: TObject);
var
	i : integer;
begin
	if (BoidEngine <> nil) then
  begin
  	for i := 0 to BoidEngine.ObstacleList.Count - 1 do
	  	TObstacle(BoidEngine.ObstacleList[i]).Destroy;

		BoidEngine.ObstacleList.Clear;
  end;
end;

//******************************************************************************
procedure TfrmBoids.Causesometurmoil1Click(Sender: TObject);
var
	i : integer;
begin
	if not bRun then exit;

	for i := 0 to BoidEngine.MovableList.Count - 1 do
		TBoid(BoidEngine.MovableList[i]).Dir := random(627) / 100.0 - pi;
end;

//******************************************************************************
procedure TfrmBoids.Hideobstacles1Click(Sender: TObject);
begin
	if chkObstacle.State = cbChecked then
  	chkObstacle.State := cbUnChecked
  else
  	chkObstacle.State := cbChecked;
end;

//******************************************************************************
procedure TfrmBoids.tbSpeedChangeChange(Sender: TObject);
begin
	if BoidEngine = nil then exit;
	BoidEngine.MaxSpeedChange := tbSpeedChange.Position/100;
  BoidEngine.CopySettingsToAllBoids;
end;

//******************************************************************************
procedure TfrmBoids.tbMaxSpeedChange(Sender: TObject);
begin
	if BoidEngine = nil then exit;
	BoidEngine.MaxSpeed := tbMaxSpeed.Position/2;
  BoidEngine.CopySettingsToAllBoids;
end;

//******************************************************************************
procedure TfrmBoids.tbSensorRangeChange(Sender: TObject);
begin
	if BoidEngine = nil then exit;
	BoidEngine.SensorDistance := tbSensorRange.Position;
  BoidEngine.CopySettingsToAllBoids;
end;

//******************************************************************************
procedure TfrmBoids.tbOptimalDistanceChange(Sender: TObject);
begin
	if BoidEngine = nil then exit;
	BoidEngine.OptimalDistance := tbOptimalDistance.Position;
  BoidEngine.CopySettingsToAllBoids;
end;

procedure TfrmBoids.tbStayInCenterChange(Sender: TObject);
begin
	if BoidEngine = nil then exit;
	BoidEngine.StayInCenter := tbStayInCenter.Position/10;
  BoidEngine.CopySettingsToAllBoids;
end;

//******************************************************************************
procedure TfrmBoids.tbTooCloseChange(Sender: TObject);
begin
	if BoidEngine = nil then exit;
	BoidEngine.TooClose := tbTooClose.Position;
  BoidEngine.ReallyClose := BoidEngine.TooClose - 2;
  BoidEngine.CopySettingsToAllBoids;
end;

{procedure TfrmBoids.tbReallyCloseChange(Sender: TObject);
begin
	if BoidEngine = nil then exit;
	BoidEngine.ReallyClose := tbReallyClose.Position;
  BoidEngine.CopySettingsToAllBoids;
end;}

//******************************************************************************
procedure TfrmBoids.tbTurnRateChange(Sender: TObject);
begin
	if BoidEngine = nil then exit;
	BoidEngine.MaxTurnSpeed := tbTurnRate.Position/100;
  BoidEngine.CopySettingsToAllBoids;
end;

//******************************************************************************
procedure TfrmBoids.tbLineLengthChange(Sender: TObject);
begin
	if BoidEngine = nil then exit;
	BoidEngine.LineLength := tbLineLength.Position;
  BoidEngine.CopySettingsToAllBoids;
end;

//******************************************************************************
procedure TfrmBoids.cmdSaveSettingsClick(Sender: TObject);
var
	i : integer;
begin
	SaveSettingsFromTrackBarsToIni;
  IniFile.ReadSections(cboBoidStyles.Items);
end;

//******************************************************************************
procedure TfrmBoids.cboBoidStylesChange(Sender: TObject);
begin
	ReadSettingsFromIniToTrackBars;
end;

//******************************************************************************
procedure TfrmBoids.cmdEraseStyleClick(Sender: TObject);
begin
	IniFile.EraseSection(cboBoidStyles.Text);
	//IniFile.ReadSections(cboBoidStyles.Items);
  cboBoidStyles.Items.Delete(cboBoidStyles.ItemIndex);
end;

//******************************************************************************
procedure TfrmBoids.chkBlankingClick(Sender: TObject);
begin
	if BoidEngine <> nil then
  	BoidEngine.bBlanking := chkBlanking.Checked;
end;

procedure TfrmBoids.Button1Click(Sender: TObject);
begin
{	bRun := false;
  Application.ProcessMessages;
  Application.ProcessMessages;
  Application.ProcessMessages;

	frmDirectDraw.Show;
  StartBoidEngine(true);}
end;

//******************************************************************************
procedure TfrmBoids.chkDrawClick(Sender: TObject);
begin
  BoidEngine.bDrawGraphics := chkDraw.Checked;
end;

procedure TfrmBoids.Panel1Click(Sender: TObject);
begin

end;

//******************************************************************************
procedure TfrmBoids.cmdSettingsClick(Sender: TObject);
begin
  frmSettings.Show;
end;

//******************************************************************************
end.

