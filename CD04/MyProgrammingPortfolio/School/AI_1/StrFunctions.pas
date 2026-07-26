unit StrFunctions;

interface

const
	TDigitsSet = ['0','1','2','3','4','5','6','7','8','9',',','-'];

  function GetBefore(sSubString, sLongString : string) : string;
  function GetAfter(sSubString, sLongString : string) : string;
  function GetBetween(sSubString, sLongString : string) : string;
  function GetBetweenDifferent(sSubStringFirst,sSubStringSecond, sLongString : string) : string;
  function GetNTh(sSubString, sLongString : string; iNumber : integer) : string;
  function IsNumeric(sToken : string) : boolean;

implementation

function GetBefore(sSubString, sLongString : string) : string;
var
  i : integer;
begin
  i := Pos(sSubString, sLongString);
  if i <> 0 then
    GetBefore := Copy(sLongString,0, i-1)
  else
    GetBefore := '';
end;

function GetAfter(sSubString, sLongString : string) : string;
var
  i : integer;
begin
  i := Pos(sSubString, sLongString);
  if i <> 0 then
    GetAfter := Copy(sLongString,i+1, Length(sLongString)-i)
  else
    GetAfter := '';
end;

function GetBetween(sSubString, sLongString : string) : string;
begin
  GetBetween := GetBefore(sSubString,GetAfter(sSubString,sLongString));
end;

function GetBetweenDifferent(sSubStringFirst,sSubStringSecond, sLongString : string) : string;
begin
  GetBetweenDifferent := GetBefore(sSubStringSecond,GetAfter(sSubStringFirst,sLongString));
end;

function GetNTh(sSubString, sLongString : string; iNumber : integer) : string;
var
  i           : integer;
  sLongLeft   : string;
  sTempResult : string;
begin
  sLongLeft := sLongString;
  sLongLeft := sLongLeft + ',';

  for i := 0 to iNumber do
  begin
    sTempResult := GetBefore(sSubString, sLongLeft);
    sLongLeft := GetAfter(sSubString, sLongLeft);
  end;

  GetNth := sTempResult;
end;

function IsNumeric(sToken : string) : boolean;
var
  i       : integer;
  bResult : boolean;
begin
  if Length(sToken) > 0 then
  begin
    bResult := true;
    for i := 1 to Length(sToken) do
    begin
      if not (sToken[i] in TDigitsSet) then
      begin
        bResult := false;
        break;
      end;
    end;
  end else bResult := false;

  if (sToken = '.') or (sToken = '-') or  (sToken = '+') then
  	bResult := false; 

  IsNumeric := bResult;
end;
end.
