unit Loading;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ExtCtrls, ComCtrls, StdCtrls, jpeg;

type
  TForm2 = class(TForm)
    Image1: TImage;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    ProgressBar1: TProgressBar;
    Timer1: TTimer;
    procedure Timer1Timer(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form2: TForm2;

implementation

uses Main;

{$R *.dfm}

procedure TForm2.Timer1Timer(Sender: TObject);
begin
  if ProgressBar1.Position = 100 then
  begin
    Timer1.Enabled := false;
    Form2.close;
  end

  else
  begin
    ProgressBar1.Position := ProgressBar1.Position + 5;
  end;
end;

end.procedure TForm2.Timer1Timer(Sender: TObject);
begin

end;


