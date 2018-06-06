unit Unit1;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, jpeg, ExtCtrls;

type
  TForm1 = class(TForm)
    Image1: TImage;
    Label1: TLabel;
    Memo1: TMemo;
    Label3: TLabel;
    Label2: TLabel;
    procedure FormShow(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form1: TForm1;

implementation

{$R *.dfm}

procedure TForm1.FormShow(Sender: TObject);
begin
  Memo1.Lines.Clear;
  Memo1.Lines.Add('Warning. All rights reserved intergroup copyright treaties.');
  Memo1.Lines.Add('Full or partial copying and distribution product by any');
  Memo1.Lines.Add('means is prohibited without the special agreement of the');
  Memo1.Lines.Add('author and will be prosecuted by all rigor of human laws!');
end;

end.
