unit ufrmAbout;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, ComCtrls, Menus;

type
  TfrmAbout = class(TForm)
    txtAbout: TRichEdit;
    MainMenu1: TMainMenu;
    File1: TMenuItem;
    Exit1: TMenuItem;
    Save1: TMenuItem;
    procedure Exit1Click(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure Save1Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmAbout: TfrmAbout;

implementation

{$R *.DFM}

//******************************************************************************
procedure TfrmAbout.Exit1Click(Sender: TObject);
begin
	Hide;
end;

//******************************************************************************
procedure TfrmAbout.FormShow(Sender: TObject);
begin
	txtAbout.Lines.LoadFromFile('about.rtf');
end;

//******************************************************************************
procedure TfrmAbout.Save1Click(Sender: TObject);
begin
	txtAbout.Lines.SaveToFile('about.rtf');
end;
end.
