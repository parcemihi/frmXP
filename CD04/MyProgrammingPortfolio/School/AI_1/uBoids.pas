unit uBoids;

interface

uses
	Graphics, Classes,uTMovable;

const
  RecommendedSpeed = 0.45;
  LowestSpeedAllowed = 0.5;

type
	TObstacle = class;

  TBoid = class(TMovable)
    bSmashed								: boolean;
    iTeamNumber							: integer;

    DeltaDirToClosest 			: single;
    AbsDirToClosest					: single;
    SQRDist									: single;
    sDx,sDy             		: single;
    AbsDirToAvg							: single;
    AvgSpeed								: single;
    AvgDir									: single;

	  // Unique to every boid if they are to be extended,
    // but usually the same.
    MaxSpeedChange 	: single;
  	SensorDistance 	: single;

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
    bExecuteSeparation  : boolean;
    bExecuteAlignment   : boolean;
    bExecuteCohesion    : boolean;

    procedure AvoidObstacle(Canvas : TCanvas;ClosestObstacle : TObstacle);
    function AvoidBoid(ClosestBoid : TBoid; Canvas : TCanvas) : boolean;
    procedure StayCentered(Canvas : TCanvas);
    function PrepareToMove(ClosestBoids : TList; ClosestObstacle : TMovable; Canvas : TCanvas) : boolean ; override;

    procedure IncreaseSpeed(delta : single);override;
    procedure DecreaseSpeed(delta : single);override;

    procedure TurnLeft(delta : single);override;
    procedure TurnRight(delta : single);override;
  private
  	TurnFraction				: single;
  end;

  TObstacle = class(TMovable)
  	Size 					: integer;
    AvoidSphere		: single;
    procedure Draw(Canvas : TCanvas);override;
    function PrepareToMove(ClosestBoids : TList;ClosestObstacle : TMovable; Canvas : TCanvas) : boolean;override;
   // procedure Move(Canvas : TCanvas);override;
    constructor Create(inX, inY,inSize : Integer; inColor : TColor; Canvas : TCanvas);
	end;

implementation

//******************************************************************************
procedure TBoid.TurnLeft(delta : single);
begin
  Dir := Dir - Min(MaxTurnSpeed,abs(delta));
end;

//******************************************************************************
procedure TBoid.TurnRight(delta : single);
begin
  Dir := Dir + Min(MaxTurnSpeed,abs(delta));
end;

//******************************************************************************
procedure TBoid.IncreaseSpeed(delta : single);
begin
  // Increase is slower than decrease
  Speed := Speed + Min(MaxSpeedChange,delta) + Deviation(0.05);
  if Speed > 1 then speed := 1;
  if Color <> clGreen then
	  Color := clBlue;
  bSpeedHasChanged := true;
end;

//******************************************************************************
procedure TBoid.DecreaseSpeed(delta : single);
begin
  // Increase is slower than decrease
  Speed := Speed - Min(MaxSpeedChange,delta) + Deviation(0.05);
  if Speed < LowestSpeedAllowed then Speed := LowestSpeedAllowed;
  if Color <> clGreen then
	  Color := clRed;
  bSpeedHasChanged := true;
end;

//******************************************************************************
procedure TBoid.AvoidObstacle(Canvas : TCanvas;ClosestObstacle : TObstacle);
var
	ODistSQR						: single;
  ODir								: single;
  ClosestDistance			: single;
  AvoidObstacleDistSQR : single;
  fPanicLevel					: single;
  fTurnDir						: single;
begin
 if (ClosestObstacle <> nil) then
  begin
 	  AvoidObstacleDistSQR := ClosestObstacle.AvoidSphere;

    ODistSQR := sqr(x-ClosestObstacle.X) + sqr(y-ClosestObstacle.y);

		if (ODistSQR < AvoidObstacleDistSQR) then
    begin
	  	ODir := ConfineDirection(Dir-MyArcTan(ClosestObstacle.X-x,ClosestObstacle.y-y));

			ClosestDistance := abs(sin(ODir) * Sqrt(ODistSQR));

			if (abs(ODir) < pi/2) and (ClosestDistance < ClosestObstacle.Size) then
			begin
        //fPanicLevel := 1-ODistSQR/AvoidObstacleDistSQR;
        fPanicLevel := max(1-ODistSQR/AvoidObstacleDistSQR,
        1 - ClosestDistance/(ClosestObstacle.Size/2));

        if bObstacleAvoidance then
        begin
          Canvas.Pen.Color := clYellow;
          Canvas.MoveTo(trunc(x),trunc(y));
          Canvas.LineTo(trunc(x + cos(Dir-ODir)*ClosestDistance),
                      trunc(y + sin(Dir-ODir)*ClosestDistance));

          Canvas.Pen.Color := clBlue;
          Canvas.MoveTo(trunc(x),trunc(y));
          Canvas.LineTo(trunc(x + cos(Dir)*ClosestDistance),
                        trunc(y + sin(Dir)*ClosestDistance));
        end;

				fTurnDir :=(pi-abs(ODir))*fPanicLevel;

        if ODir > 0 then
          TurnRight(fTurnDir)
        else
          TurnLeft(fTurnDir);

        Color := clGreen;

//        TurnFraction := 1 - fPanicLevel;
        TurnFraction := 0.1;

        if fPanicLevel > 0.4 then
        begin
        	DecreaseSpeed(fPanicLevel/2);
	        TurnFraction := 0.1;
        end;

 				if fPanicLevel > 0.6 then
          TurnFraction := 0.0;
      end;
    end;
  end; //}
end;

function TBoid.AvoidBoid(ClosestBoid : TBoid; Canvas : TCanvas) : boolean;
var
  iLineLen : integer;
begin
	AvoidBoid := false;

  if (ClosestBoid.DistanceSquared < sqr(TooClose)) then
  begin
    // Don't fly directly behind someone!
    //if abs(DeltaDir) < 0.02 then DeltaDir := 0.07;
    if abs(DeltaDirToClosest) < 0.02 then
      if DeltaDirToClosest < 0 then
        DeltaDirToClosest := -0.1
      else
        DeltaDirToClosest := 0.1;//}

	  // Allow boids going the same direction to be closer
    if (DeltaDirToClosest < 0.09) and
       (ClosestBoid.DistanceSquared > sqr(ReallyClose)) then
    	DeltaDirToClosest := 0;

    if DeltaDirToClosest <> 0 then
    begin
    	bSmashed := true;
      AvoidBoid := true;

      if bBoidAvoidance then
      begin
        iLineLen := Trunc(Sqrt(ClosestBoid.DistanceSquared));
        Canvas.Pen.Color := clYellow;
	      Canvas.MoveTo(trunc(x),trunc(y));
		  	Canvas.LineTo(trunc(x + cos(Dir+DeltaDirToClosest)*iLineLen),
      					trunc(y + sin(Dir+DeltaDirToClosest)*iLineLen));//}
      end;

      if DeltaDirToClosest < 0 then
        TurnRight(DeltaDirToClosest*TurnFraction)
      else
        TurnLeft(DeltaDirToClosest*TurnFraction);//}

      // Adjust speed!
      if Within(AbsDirToClosest,-Pi/2,0) and
         Within(DeltaDirToClosest,0,Pi) then  // Beta2
        DecreaseSpeed(abs(ClosestBoid.Speed-Speed)+0.09);

      if Within(AbsDirToClosest,0,Pi/2) and
         Within(DeltaDirToClosest,-Pi,0) then
        DecreaseSpeed(abs(ClosestBoid.Speed-Speed)+0.09);

      // Overrules the next rule
      if (ClosestBoid.DistanceSquared < sqr(ReallyClose)) then
      	TurnFraction := 0.01*TurnFraction;
    end;
  end;
  //}
end;

procedure TBoid.StayCentered(Canvas : TCanvas);
var
  NewAvgDir						: single;
  TestDir             : single;
  LocalTurnFraction		: single;
  iLineLen            : integer;
  DeltaToAverageDir   : single;
begin

  NewAvgDir := AvgDir;

  if bExecuteCohesion and
    (SqrDist > sqr(OptimalDistance-StayInCenter)) then
  begin
    //3.Cohesion: steer to move toward the average position of local flockmates. / Craig Reynolds
    // We're too far off, stear not only in the same direciton,
    // but towards the center of gravity!
    //   The influence of this rule is very small!

    TestDir := ConfineDirection((
      ConfineDirection(AbsDirToAvg)+
      ConfineDirection(AvgDir)*9)/10);//}

    // If the Alignment rule is turned off, then
    // cohesion decides all the way, otherwise they both decide!
    if bExecuteAlignment then
      NewAvgDir := TestDir
    else
      NewAvgDir := AbsDirToAvg;

    if bShowCohesion then
    begin
     iLineLen := Trunc(sqrt(SqrDist));
     Canvas.Pen.Color := clYellow;
     Canvas.MoveTo(trunc(x),trunc(y));
     Canvas.LineTo(trunc(x + cos(AbsDirToAvg)*iLineLen),
       trunc(y + sin(AbsDirToAvg)*iLineLen));//}

     iLineLen := Trunc(sqrt(SqrDist));
     Canvas.Pen.Color := clWhite;
     Canvas.MoveTo(trunc(x),trunc(y));
     Canvas.LineTo(trunc(x + cos(AvgDir)*iLineLen),
       trunc(y + sin(AvgDir)*iLineLen));//}

     iLineLen := Trunc(sqrt(SqrDist));
     Canvas.Pen.Color := clBlue;
     Canvas.MoveTo(trunc(x),trunc(y));
     Canvas.LineTo(trunc(x + cos(TestDir)*iLineLen),
       trunc(y + sin(TestDir)*iLineLen));//} *)
    end;

    // Break if you're in front of the crowd,
    // speed up if you're behind it.
    if Abs(Dir-AbsDirToAvg) < pi/2 then
      IncreaseSpeed(abs(AvgSpeed-Speed)+0.02)
    else
      DecreaseSpeed(abs(AvgSpeed-Speed)+0.02);
  end; //}

  DeltaToAverageDir := ConfineDirection(Dir - NewAvgDir);
  //DeltaToAverageDir := Dir - NewAvgDir;

  // Only show boids that differ in direction with XX
  // radians!
  if bShowAlignment and (abs(DeltaToAverageDir)>0.07) then
  begin
    iLineLen := Trunc(sqrt(SqrDist));
    Canvas.Pen.Color := clWhite;
    Canvas.MoveTo(trunc(x),trunc(y));
    Canvas.LineTo(trunc(x + cos(AvgDir)*iLineLen),
      trunc(y + sin(AvgDir)*iLineLen));//}

    iLineLen := Trunc(sqrt(SqrDist));
    Canvas.Pen.Color := clBlue;
    Canvas.MoveTo(trunc(x),trunc(y));
    Canvas.LineTo(trunc(x + cos(Dir)*iLineLen),
      trunc(y + sin(Dir)*iLineLen));//} *)
  end;

  if DeltaToAverageDir < 0 then
    TurnRight(DeltaToAverageDir*TurnFraction)
  else
    TurnLeft(DeltaToAverageDir*TurnFraction);
  //}       *)
end;

//******************************************************************************
function TBoid.PrepareToMove(ClosestBoids : TList; ClosestObstacle : TMovable;
  Canvas : TCanvas) : boolean;
var
  i                   : integer;
  ClosestBoid         : TBoid;
  RelativeDirection   : single;
  dx,dy               : single;

  bNeighboursFound    : boolean;

  RelDirToAvg        	: single;
  iTeamCount					: integer;
  LastMoveXSum				: single;
  LastMoveYSum				: single;
  TestBoid						: TBoid;
begin
  // First, collect data
  Color := clWhite;
  bSpeedHasChanged := false;
  bSmashed := false;
  iTeamCount := 0;

  PrepareToMove := false;

  sDx := 0;
  sDy := 0;
  AvgSpeed := 0;
  AvgDir := 0;
  LastMoveXSum := 0;
  LastMoveYSum := 0;

  bNeighboursFound := ClosestBoids.Count <> 0;

  if bNeighboursFound then
    ClosestBoid := ClosestBoids[0]
  else
    ClosestBoid := nil;

  for i := 0 to ClosestBoids.Count - 1 do
  begin
  	TestBoid := ClosestBoids[i];
	  AvgSpeed := AvgSpeed + TestBoid.Speed;
  	AvgDir := AvgDir + TestBoid.Dir;
    sDx := sDx + TestBoid.X;
    sDy := sDy + TestBoid.Y;
    LastMoveXSum := LastMoveXSum + TestBoid.LastMoveX;
    LastMoveYSum := LastMoveYSum + TestBoid.LastMoveY;

    if TestBoid.DistanceSquared < ClosestBoid.DistanceSquared then
      ClosestBoid := ClosestBoids[i];
  end;

  if bNeighboursFound then
  begin
    AvgSpeed := AvgSpeed / ClosestBoids.Count;

    // Iclude this boids direction!
    AvgDir := (AvgDir+Dir) / (ClosestBoids.Count + 1);


		// Calculate the average heading of the surrounding flock, including
    // the boid itself
    {LastMoveXSum := LastMoveXSum + LastMoveX;
    LastMoveYSum := LastMoveYSum + LastMoveY;

    AvgDir := MyArcTan(LastMoveYSum, LastMoveXSum);//}

    sDx := sDx / ClosestBoids.Count;
    sDy := sDy / ClosestBoids.Count;

    dx := ClosestBoid.X-x;
    dy := ClosestBoid.Y-y;

    AbsDirToClosest := MyArcTan(dx,dy);
    AbsDirToAvg := MyArcTan(sDx-x,sDy-y);
	  //DeltaDirToClosest := ConfineDirection(abs(ClosestBoid.Dir-Dir));
    DeltaDirToClosest := -ConfineDirection(Dir-MyArcTan(ClosestBoid.X-x,ClosestBoid.y-y));
    //ODir := ConfineDirection(Dir-MyArcTan(ClosestObstacle.X-x,ClosestObstacle.y-y));

  end
  else
  begin
    AvgSpeed := RecommendedSpeed;
    AvgDir := Dir;
    AbsDirToClosest := dir;
    sDx := 0;
    sDy := 0;
    DeltaDirToClosest := 0;
  end;

  SQRDist := sqr(sDx-x)+sqr(sDy-y);
  TurnFraction := 1.0;

  // Craig Reynold's three rules of flocking are;
  // 1.Separation: steer to avoid crowding local flockmates.
  // 2.Alignment: steer towards the average heading of local flockmates.
  // 3.Cohesion: steer to move toward the average position of local flockmates.

  // Rule zero is just to spice it up!

  // * Rule zero; don't hit the obstacle!
  if bExecuteObstacles then
   	AvoidObstacle(Canvas,TObstacle(ClosestObstacle));

  // 1.Separation: steer to avoid crowding local flockmates. / Craig Reynolds
  if bExecuteSeparation and bNeighboursFound then
		AvoidBoid(ClosestBoid, Canvas);//}

  //2.Alignment: steer towards the average heading of local flockmates. / Craig Reynolds
	// and
  //3.Cohesion: steer to move toward the average position of local flockmates.
  if (bExecuteAlignment or bExecuteCohesion) and bNeighboursFound then
		StayCentered(Canvas);//}
end;

//******************************************************************************
procedure TObstacle.Draw(Canvas : TCanvas);
var
  cX, cy : integer;
  hSize	 : integer;
begin
	if not bActive then exit;
	cx := trunc(x);
  cy := trunc(y);
  hSize := Size div 2;
  Canvas.Brush.Color := Color;
  Canvas.Pen.Color := Color;
  Canvas.Ellipse(Cx - hSize, Cy - hSize,Cx + hSize, Cy + hSize);
  Canvas.Brush.Color := clWhite;
end;

//******************************************************************************
function TObstacle.PrepareToMove(ClosestBoids : TList;ClosestObstacle : TMovable; Canvas : TCanvas):boolean;
begin
end;

//******************************************************************************
constructor TObstacle.Create(inX, inY,inSize : Integer; inColor : TColor; Canvas : TCanvas);
begin
	inherited Create(Canvas);
	X := inX;
  Y := inY;
  Speed := 0;
  Size := inSize;
  AvoidSphere := sqr(inSize*3);
  Color := inColor;
  bActive := true;
end;
end.
