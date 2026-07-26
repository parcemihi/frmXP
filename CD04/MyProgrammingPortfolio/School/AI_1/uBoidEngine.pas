unit uBoidEngine;

interface

uses
	Graphics, Classes, uTMovableEngine, uTMovable, uBoids;

type
  TBoidEngine = class(TMovableEngine)
    LineLength			: single;

	  OptimalDistance : single;
    StayInCenter		: single;
  	TooClose 				: single;
	  ReallyClose 		: single;
  	MaxTurnSpeed 		: single;

    bBoidAvoidance      : boolean;
    bObstacleAvoidance  : boolean;
    bShowAlignment      : boolean;
    bShowCohesion       : boolean;

    bExecuteObstacles   : boolean;
    bExecuteSeparation   : boolean;
    bExecuteAlignment   : boolean;
    bExecuteCohesion    : boolean;

//    procedure RunStep;
    procedure CopySettingsToBoid(Boid : TBoid);
    procedure CopySettingsToAllBoids;

    constructor Create(iBoidCount : integer; OutputCanvas : TCanvas);
    procedure AdjustMovableCount(iMovableCount : integer);
  end;

implementation

//******************************************************************************
constructor TBoidEngine.Create(iBoidCount : integer; OutputCanvas : TCanvas);
var
  i 				: integer;
  NewBoid 	: TBoid;
begin
	inherited Create(OutputCanvas);

  MaxSpeed := 5;
  MaxSpeedChange := 0.09;
  SensorDistance := 60;

  OptimalDistance := 20;
  TooClose := 19;
  ReallyClose := 12;
  MaxTurnSpeed := 0.05;

  LineLength := 3;

  bExecuteObstacles   := true;
  bExecuteSeparation  := true;
  bExecuteAlignment   := true;
  bExecuteCohesion    := true;

  for i := 1 to iBoidCount do
  begin
  	NewBoid := TBoid.Create(Canvas);
    MovableList.Add(NewBoid);

		CopySettingsToBoid(NewBoid);
  end;
end;

//******************************************************************************
procedure TBoidEngine.CopySettingsToAllBoids;
var
	i : integer;
begin
	for i := 0 to MovableList.Count - 1 do
  	CopySettingsToBoid(TBoid(MovableList[i]));
end;

//******************************************************************************
procedure TBoidEngine.CopySettingsToBoid(Boid : TBoid);
begin
  Boid.MaxSpeed := MaxSpeed;
  Boid.MaxSpeedChange := MaxSpeedChange;
 	Boid.SensorDistance := SensorDistance;

  Boid.OptimalDistance := OptimalDistance;
  Boid.StayInCenter := StayInCenter;
 	Boid.TooClose := TooClose;
  Boid.ReallyClose := ReallyClose;
 	Boid.MaxTurnSpeed := MaxTurnSpeed;
  Boid.LineLength := LineLength;

  Boid.bBoidAvoidance := bBoidAvoidance;
  Boid.bObstacleAvoidance := bObstacleAvoidance;
  Boid.bShowCohesion := bShowCohesion;
  Boid.bShowAlignment := bShowAlignment;

  Boid.bExecuteObstacles := bExecuteObstacles;
  Boid.bExecuteSeparation := bExecuteSeparation;
  Boid.bExecuteAlignment := bExecuteAlignment;
  Boid.bExecuteCohesion := bExecuteCohesion;
end;

//******************************************************************************
procedure TBoidEngine.AdjustMovableCount(iMovableCount : integer);
var
	i 					: integer;
	NewBoid 		: TBoid;
begin
  if iMovableCount > MovableList.Count then
  begin
	  for i := MovableList.Count to iMovableCount-1 do
  	begin
	  	NewBoid := TBoid.Create(Canvas);
    	MovableList.Add(NewBoid);

			CopySettingsToBoid(NewBoid);
  	end;
  end else
  while iMovableCount < MovableList.Count do
  begin
    TMovable(MovableList[MovableList.Count-1]).Destroy;
    MovableList.Delete(MovableList.Count-1);
    MovableList.Pack;
  end;
end;
end.
