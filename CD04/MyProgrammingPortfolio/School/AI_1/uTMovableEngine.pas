unit uTMovableEngine;

interface

uses
	Graphics, Classes, Forms, Windows, uTMovable;

type
  TWithinResultTypes = (wrtBefore,wrtWithin,wrtWithinOutside,wrtAfter);

  TMovableEngine = class
    MovableList			: Tlist;
    ObstacleList		: TList;
    Canvas					: TCanvas;

    bBlanking				: boolean;
    bDrawGraphics   : boolean; // For speed testing

    MaxSpeed 				: single;
	  MaxSpeedChange 	: single;
  	SensorDistance 	: single;
    Application			: TApplication;

    LineLength			: single;

    // 75% chance of checking the neighbourhood!
    // The rest of the time the old info is re-used.
    // Saves on cycles!
    CheckChancePerFrame : integer;

    procedure RunStep;virtual;
    procedure GetClosestMovables(MyMovable : TMovable; CloseList : TList);virtual;
    function GetClosestObstacle(MyMovable : TMovable) : TMovable;virtual;
    constructor Create(OutputCanvas : TCanvas);virtual;
    procedure SetNewCanvas(OutputCanvas : TCanvas);
    destructor Destroy;override;
  end;

  procedure ClearArena(Canvas : TCanvas);

implementation

//******************************************************************************
procedure ClearArena(Canvas : TCanvas);
begin
	Canvas.Pen.Color := clBlack;
  Canvas.Brush.Color := clBlack;
	Canvas.FillRect(Canvas.ClipRect);
end;

//******************************************************************************
function CompareMovable(Item1, Item2: Pointer): Integer;
var
  Movable1, Movable2  : TMovable;
  iTemp               : integer;
begin
  Movable1 := Item1;
  Movable2 := Item2;

  iTemp := trunc(Movable1.X-Movable2.X);

  if iTemp = 0 then
    iTemp := trunc(Movable1.Y-Movable2.Y);

  CompareMovable := iTemp;
end;

//******************************************************************************
function BinSearch(TestList : TList; FindObj : Pointer; CompFun : TListSortCompare;iStartBottom,iStartTop : integer) : integer;
var
  Top,Bottom,Test     : integer;
  TestResult          : integer;
  bItsAtTheBottom     : boolean;
begin
  Bottom := iStartBottom;

  if iStartTop <> -1 then
    Top := iStartTop
  else
    Top := TestList.Count-1;
  TestResult := -1;
  bItsAtTheBottom := false;

  while not bItsAtTheBottom and (Top-Bottom > 1) {and (TestResult <> 0)} do
  begin
    Test := (Bottom+Top) div 2;
    TestResult := CompFun(FindObj, TestList[Test]);
    if TestResult <= 0 then
      Top := Test
    else
    begin
      Bottom := Test;

      bItsAtTheBottom := (TestResult = 0);
    end;
  end;

  if bItsAtTheBottom then
    BinSearch := Bottom
  else
    BinSearch := Top;
end;

//******************************************************************************
function IsWithinRange(Movable : TMovable; Range : TRect) : TWithinResultTypes;
begin

  //TWithinResultTypes = (wrtBefore,wrtWithin,wrtWithinOutside,wrtAfter);
  if (Movable.X < Range.Left) then
    IsWithinRange := wrtBefore
  else
  if (Movable.X > Range.Right) then
    IsWithinRange := wrtAfter
  else
  if (Movable.X >= Range.Left) and (Movable.X <= Range.Right) and
    (Movable.Y <= Range.Bottom) and (Movable.Y >= Range.Top) then
    IsWithinRange := wrtWithin
  else
    IsWithinRange := wrtWithinOutside;
end;

//******************************************************************************
constructor TMovableEngine.Create(OutputCanvas : TCanvas);
var
  i 				: integer;
begin
  MovableList := TList.Create;
  ObstacleList := TList.Create;

  Canvas := OutputCanvas;

  bBlanking := true;

  MaxSpeed := 5;
  MaxSpeedChange := 0.09;
  SensorDistance := 60;
  bDrawGraphics := true;

  LineLength := 3;

  CheckChancePerFrame := 100;
end;

//******************************************************************************
procedure TMovableEngine.SetNewCanvas(OutputCanvas : TCanvas);
var
	i : integer;
begin
	Canvas := OutputCanvas;
{  for i := 0 to MovableList.Count - 1 do
    TMovable(MovableList[i]).Canvas := OutputCanvas;

 	for i := 0 to MovableList.Count - 1 do
    TMovable(ObstacleList[i]).Canvas := OutputCanvas;}
end;

//******************************************************************************
destructor TMovableEngine.Destroy;
var
  i : integer;
begin
  for i := 0 to MovableList.Count - 1 do
    TMovable(MovableList[i]).Destroy;

 	for i := 0 to ObstacleList.Count - 1 do
    TMovable(ObstacleList[i]).Destroy;

 	MovableList.Destroy;
  ObstacleList.Destroy;
end;

//******************************************************************************
function TMovableEngine.GetClosestObstacle(MyMovable : TMovable) : TMovable;
var
  i           	: integer;
  x,y         	: single;
  dx,dy					: single;
  Range        	: single;
  DistSQR      	: single;
  TestObstacle 	: TMovable;
  HDist         : single;
  fClosestSoFar : single;
begin
  X := MyMovable.X;
  Y := MyMovable.Y;
  fClosestSoFar := 0;

  // Assume none will be found!
  GetClosestObstacle := nil;

  for i := 0 to ObstacleList.Count - 1 do
  begin
    TestObstacle := ObstacleList[i];
    dx := Abs(TestObstacle.x-x);
    dy := abs(TestObstacle.y-y);

	  DistSQR := sqr(dx)+sqr(dy);
 	  if (DistSQR < fClosestSoFar) or (fClosestSoFar=0) then
    begin
    	fClosestSoFar := DistSQR;
      GetClosestObstacle := TestObstacle;
 	  end;
  end;
end;

//******************************************************************************
procedure TMovableEngine.RunStep;
var
  i                 : integer;
  Obstacle					: TMovable;
begin

  if ObstacleList.Count <> 0 then
	  Obstacle := ObstacleList[0];

	MovableList.Sort(CompareMovable);

  for i := 0 to MovableList.Count - 1 do
  begin

    if Random(100) < CheckChancePerFrame then
    begin
      TMovable(MovableList[i]).ClosestMovables.Clear;

      GetClosestMovables(TMovable(MovableList[i]),TMovable(MovableList[i]).ClosestMovables);
    end;

    Obstacle := GetClosestObstacle(TMovable(MovableList[i]));

    if (Obstacle <> nil) and (not Obstacle.bActive) then
      Obstacle := nil;

		TMovable(MovableList[i]).PrepareToMove(TMovable(MovableList[i]).ClosestMovables, Obstacle,Canvas);
  end;

  Application.ProcessMessages;

  for i := 0 to ObstacleList.Count - 1 do
    TMovable(ObstacleList[i]).Move(Canvas);

  for i := 0 to MovableList.Count - 1 do
    TMovable(MovableList[i]).Move(Canvas);

  if bDrawGraphics then
  begin
    if bBlanking then
	  	ClearArena(Canvas);

    for i := 0 to ObstacleList.Count - 1 do
      TMovable(ObstacleList[i]).Draw(Canvas);//}

    for i := 0 to MovableList.Count - 1 do
      TMovable(MovableList[i]).Draw(Canvas); //}*)
  end;
end;

//******************************************************************************
procedure TMovableEngine.GetClosestMovables(MyMovable : TMovable; CloseList : TList);
// This function is fairly complicated, if you need to alter it, be very careful,
// or use one of the three other versions included at the end of this file. They're
// all slower than this version, but that's mostly noticable when the number of
// movables are in the hundreds.
//
//function GetMovablesWithinRangeSmarter(MovableList : TList; Range : TRect) : integer;
var
  i             : integer;
  LastPosition  : integer;
  Test          : TMovable;
  TestMovable   : TMovable;
  WithinResult  : TWithinResultTypes;
  iFound        : integer;
  LastX         : single;
  Range					: TRect;
  SensHalf			: single;
begin
  Test := TMovable.Create(nil);

  SensHalf := SensorDistance/2;

  with MyMovable do
	  Range := Rect(trunc(X-SensHalf),trunc(Y-SensHalf),trunc(X+SensHalf),trunc(Y+SensHalf));

  if Range.Left < 0 then Range.Left := 0;
	if Range.Top < 0 then Range.Top := 0;

  Test.X := Range.Left;
  Test.Y := Range.Top;

  TestMovable := Test;
  iFound := 0;

  i := BinSearch(MovableList,Test,CompareMovable,0,-1);

  Test.X := Range.Right+1;
  LastPosition := BinSearch(MovableList,Test,CompareMovable,0,-1);

  LastX := Range.Left;
  while (i < MovableList.Count) and (WithinResult <> wrtAfter) and
        (i < LastPosition)  do
  begin
  	TestMovable := MovableList[i];
    WithinResult := IsWithinRange(TestMovable,Range);
    if WithinResult = wrtWithin then
    begin
      inc(iFound);
      LastX := TestMovable.X;

			if TestMovable <> MyMovable then
			begin
	      TestMovable.DistanceSquared := sqr(TestMovable.X-MyMovable.X) +
  	      sqr(TestMovable.Y-MyMovable.Y);

    	  CloseList.Add(TestMovable);
      end;

      inc(i);
    end else
    if WithinResult = wrtWithinOutside then
    begin
      LastX := LastX+1;
      Test.X := LastX;
      i := BinSearch(MovableList,Test,CompareMovable,i,LastPosition);
    end else inc(i);
  end;

  Test.Destroy;
//  result := iFound;
end;//}

//******************************************************************************
//Original version
{procedure TMovableEngine.GetClosestMovables(MyMovable : TMovable; CloseList : TList);
var
  i           	: integer;
  x,y         	: single;
  dx,dy					: single;
  Range        	: single;
  DistSQR      	: single;
  TestMovable  	: TMovable;
  HDist         : single;
begin
  Range := sqr(SensorDistance);
  HDist := SensorDistance/2;

  X := MyMovable.X;
  Y := MyMovable.Y;

  for i := 0 to MovableList.Count - 1 do
  begin
    TestMovable := MovableList[i];
    dx := Abs(TestMovable.x-x);
    dy := abs(TestMovable.y-y);

    //if (dx + dy)*10 < Range then
    if Within(dx, -HDist, +HDist) and
       Within(dy, -HDist, +HDist) then
    begin
	    DistSQR := (sqr(dx)+sqr(dy));
	 	  if (TestMovable <> MyMovable) and
    	   (DistSQR < Range) then
	    begin
  	    TestMovable.DistanceSquared := DistSQR;
	   	  CloseList.Add(TestMovable);
  	  end;
    end;
  end;
end;//}

//******************************************************************************
//Original version remade
{function GetMovablesWithinRange(MovableList : TList; Range : TRect) : integer;
var
  i             : integer;
  Test          : TMovable;
  WithinResult  : TWithinResultTypes;
  iFound        : integer;
begin
  Test := TMovable.Create(nil);
  Test.X := Range.Left;
  Test.Y := Range.Top;

  iFound := 0;

  i := BinSearch(MovableList,Test,CompareMovable,0,-1);

  while (i < MovableList.Count) and (WithinResult <> wrtAfter)  do
  begin
    WithinResult := IsWithinRange(TMovable(MovableList[i]),Range);
    if WithinResult = wrtWithin then
      inc(iFound);

    inc(i);
  end;

  result := iFound;

  Test.Destroy;
end;//}

//******************************************************************************
{Slightly smarter version}
{procedure TMovableEngine.GetClosestMovables(MyMovable : TMovable; CloseList : TList);
//function GetMovablesWithinRangeOld(MovableList : TList; Range : TRect) : integer;
var
  i             : integer;
  WithinResult  : TWithinResultTypes;
  iFound        : integer;
  Range					: TRect;
  SensHalf			: single;
begin
  SensHalf := SensorDistance/2;

  with MyMovable do
	  Range := Rect(trunc(X-SensHalf),trunc(Y-SensHalf),trunc(X+SensHalf),trunc(Y+SensHalf));

  if Range.Left < 0 then Range.Left := 0;
	if Range.Top < 0 then Range.Top := 0;

  iFound := 0;
  for i := 0 to MovableList.Count - 1 do
  	if MyMovable <> TMovable(MovableList[i]) then
    begin
      WithinResult := IsWithinRange(TMovable(MovableList[i]),Range);
      if WithinResult = wrtWithin then
      begin
        TMovable(MovableList[i]).DistanceSquared :=
          sqr(TMovable(MovableList[i]).X-MyMovable.X) +
          sqr(TMovable(MovableList[i]).Y-MyMovable.Y);
        CloseList.Add(MovableList[i]);
      end;
    end;
end;// }

end.
