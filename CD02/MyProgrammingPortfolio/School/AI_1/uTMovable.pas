unit uTMovable;

interface

uses
	Graphics, Classes;

type
  TMovable = class
  	X,Y   									: single;
    LastMoveX,LastMoveY 		: single;
    Speed 									: single;
    MaxSpeed 								: single;
    Dir   									: single;
    Color 									: TColor;
    bSpeedHasChanged 				: boolean;
    LineLength							: single;
    DistanceSquared 				: single;

    bActive 								: boolean;

    ClosestMovables         : TList;

    procedure Draw(Canvas : TCanvas);virtual;
    function PrepareToMove(ClosestMovables : TList; ClosestObstacle : TMovable; Canvas : TCanvas) : boolean ; virtual;

    procedure Move(Canvas : TCanvas);virtual;
    procedure TurnLeft(delta : single);virtual;
    procedure TurnRight(delta : single);virtual;
    procedure IncreaseSpeed(delta : single);virtual;
    procedure DecreaseSpeed(delta : single);virtual;
    constructor Create(Canvas : TCanvas);
    destructor Destroy;override;
  end;

  function Within(Val, Min, Max : single) : boolean;
  procedure Confine(var A : single; Min, Max : single);
  function Max(a,b : single) : single;
  function Min(a,b : single) : single;
  function Deviation(range : single) : single;
  function ConfineDirection(dir : single) : single;
  function MyArcTan(dx,dy : single) : single;

implementation

//******************************************************************************
procedure Confine(var A : single; Min, Max : single);
begin
  if A < Min then A := Max;
  if A > Max then A := Min;
end;

//******************************************************************************
function Max(a,b : single) : single;
begin
  if a > b then
    Max := a
  else
    Max := b;
end;

//******************************************************************************
function Min(a,b : single) : single;
begin
  if a < b then
    Min := a
  else
    Min := b;
end;

//******************************************************************************
function Deviation(range : single) : single;
begin
  //Deviation := (random(50)-100) * range / 100.0;
end;

//******************************************************************************
function ConfineDirection(dir : single) : single;
begin
  while dir < -pi do
    dir := dir + 2*pi;
  while dir > pi do
    dir := dir - 2*pi;
  ConfineDirection := dir;
end;

//******************************************************************************
function Within(Val, Min, Max : single) : boolean;
begin
	Within := (Val >= Min) and (Val <= Max);
end;

//******************************************************************************
function MyArcTan(dx,dy : single) : single;
var
  RelativeDirection : single;
begin
  if dx = 0 then dx := 0.0001;

  RelativeDirection := ArcTan(dy/dx);

  if dy >= 0 then
    RelativeDirection := Abs(RelativeDirection)
  else
    RelativeDirection := -Abs(RelativeDirection);

	if dx < 0 then
  	RelativeDirection := pi - RelativeDirection;

	RelativeDirection := ConfineDirection(RelativeDirection);

  MyArcTan := RelativeDirection;
end;

//******************************************************************************
constructor TMovable.Create(Canvas : TCanvas);
begin
  Speed := random(50) / 100.0 + 0.5;
  Dir :=  random(627) / 100.0 - pi;//}
  if Canvas <> nil then
  begin
	  X := Random(Canvas.ClipRect.Right);{100 +random(120);//};
  	Y := Random(Canvas.ClipRect.Bottom);{100 +random(120);//};
  end;

  ClosestMovables := TList.Create;
end;

destructor TMovable.Destroy;
begin
  ClosestMovables.Destroy;
end;

//******************************************************************************
procedure TMovable.IncreaseSpeed(delta : single);
begin
  // Increase is slower than decrease
  Speed := Speed + delta;
  if Color <> clGreen then
	  Color := clBlue;
  bSpeedHasChanged := true;
end;

//******************************************************************************
procedure TMovable.DecreaseSpeed(delta : single);
begin
  // Increase is slower than decrease
  Speed := Speed - delta;
  if Color <> clGreen then
	  Color := clRed;
  bSpeedHasChanged := true;
end;

//******************************************************************************
procedure TMovable.TurnLeft(delta : single);
begin
  Dir := Dir - abs(delta);
end;

//******************************************************************************
procedure TMovable.TurnRight(delta : single);
begin
  Dir := Dir + abs(delta);
end;

//******************************************************************************
procedure TMovable.Move(Canvas : TCanvas);
begin
	LastMoveX := Speed * cos(Dir) * MaxSpeed;
  LastMoveY := Speed * sin(Dir) * MaxSpeed;
  X := X + LastMoveX;
  Y := Y + LastMoveY;

  Confine(X,0,Canvas.ClipRect.Right);
  Confine(Y,0,Canvas.ClipRect.Bottom);
end;

//******************************************************************************
procedure TMovable.Draw(Canvas : TCanvas);
var
  dX, dy : single;
begin
  Canvas.Pen.Color := Color;
  Canvas.Brush.Color := Color;

  //Canvas.Pixels[trunc(x),trunc(y)] := Color;
  dX := LineLength * cos(Dir) / 2;
  dy := LineLength * sin(Dir) / 2;
  Canvas.MoveTo(trunc(X-dx),trunc(y - dy));
  Canvas.LineTo(trunc(X + dx), trunc(y + dy));//}

  {if bSmashed then
  	Canvas.Ellipse(trunc(X-3),trunc(y - 3),trunc(X + 3), trunc(y + 3));//}
end;

//******************************************************************************
function TMovable.PrepareToMove(ClosestMovables : TList; ClosestObstacle : TMovable; Canvas : TCanvas) : boolean;
begin
end;
end.
