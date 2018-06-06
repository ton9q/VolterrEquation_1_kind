unit Unit1;

{$WARN SYMBOL_PLATFORM OFF}

interface

uses
  Windows, ActiveX, Classes, ComObj, StdVcl,
  SysUtils, Variants, Math, Chart, Grids,
  WordXP, ExcelXP,
  Unit1_TLB;

type
  TCOMServer = class(TTypedComObject, ICOMServer)

  private
    a,b:extended; // limits of integration
    h:extended; // step of integration
    n:integer; // number of results
    //x:extended;
    myList:TStringList;

    function GetA():extended;
    function GetB():extended;
    function GetH():extended;
    function GetN():integer;

    procedure SetA(aa:extended);
    procedure SetB(bb:extended);
    procedure SetH(hh:extended);

  public

    procedure Create(aa,bb,hh:extended);stdcall;
//    procedure Destroy();stdcall;

    property Low:extended read GetA write SetA ;
    property Top:extended read GetB write SetB ;
    property Step:extended read GetH write SetH ;
    property NumberResults:integer read GetN;

    function k(x,s:extended):extended;
    function f(x:extended):extended;
    function y(x:extended):extended;
    function yy(i:integer):extended;

    procedure updateN();
    procedure displayOnGraph(str:TStringGrid;chr:TChart);
    procedure FillInTheTable(str:TStringGrid);

    procedure TextWord(word:TWordDocument;appl:TWordApplication);overload;
    procedure TextWord(word:TWordDocument;appl:TWordApplication;str:TStringGrid; chr:TChart);overload;
    procedure TextExcel();overload;
    procedure TextExcel(str:TStringGrid;chr:TChart);overload;
    procedure writeTextFile(filename:string);overload;
    procedure writeTextFile(filename:string;str:TStringGrid);overload;
    procedure readTextFile(filename:string);
end;

implementation

uses ComServ;

procedure TCOMServer.updateN();
begin
  n:=trunc((b-a)/h+1);
end;

procedure TCOMServer.displayOnGraph(str:TStringGrid;chr:TChart);
var
  int:integer;
begin
  chr.Series[0].Clear;
  chr.Series[1].Clear;

  for int:=1 to str.RowCount-1 do
  begin
    chr.Series[0].AddXY(StrToFloat(str.Cells[0,int]), StrToFloat(str.Cells[1,int]));
    chr.Series[1].AddXY(StrToFloat(str.Cells[0,int]), StrToFloat(str.Cells[2,int]));
  end;
end;

procedure TCOMServer.FillInTheTable(str:TStringGrid);
var
  i:integer;
  x:extended;
  buf:extended;
begin

  n:=trunc((b-a)/h) + 1;
  str.RowCount:=n + 1;

  for i:=1 to n do
  begin
    x:=a+(i-1)*h;
    buf:=yy(i);
    myList.Add(FloatToStr(buf));

    str.Cells[0,i]:=FloatToStr(x);
    str.Cells[1,i]:=FloatToStr(SimpleRoundTo(buf, -5));
    str.Cells[2,i]:=FloatToStr(SimpleRoundTo(y(x), -5));
  end;
end;

// solution of the equation by the method of quadrature formulas
// Yi = 2/K(Xi,Xi) * (F(Xi)/H - sum(from 1 to i-1)(Aj*K(Xi,Xj)*Yj))
// Aj=0,5 if j=1 else Aj=1
function TCOMServer.yy(i:integer):extended;
var
  g:extended;
  j:integer;
  x,k1:extended;
  x1:extended;
  sum:extended;
begin
  if i=1 then
  begin
    result:=0;
  end

  else
  begin
      x:=a+(i-1)*h;
      g:=f(x)/h;
      sum:=0;

      for j:=1 to i-1 do
      begin
        x1:=a+(j-1)*h;
        k1:=k(x,x1);

        if j=1 then
        begin
          k1:=k1/2;
        end;

        sum:=sum + k1*StrToFloat(myList[j-1]);
      end;

      g:=g-sum;
      result:=2*g/k(x,x);
  end;
end;

// K(x, s) = 2 + x^2 - s^2
function TCOMServer.k(x,s:extended):extended;
begin
  result:=2+(x-s)*(x+s);
end;

// F(x) = x^2
function TCOMServer.f(x:extended):extended;
begin
  result:=power(x,2);
end;

// y = x * e^(-x^2/2)
function TCOMServer.y(x:extended):extended;
begin
  result:=x*exp(-power(x,2)/2);
end;

procedure TCOMServer.Create(aa,bb,hh:extended);
begin
  a:=aa;
  b:=bb;
  h:=hh;
  myList:= TStringList.Create;
end;

function TCOMServer.GetA():extended;
begin
  result:=a;
end;

function TCOMServer.GetB():extended;
begin
  result:=b;
end;

function TCOMServer.GetH():extended;
begin
  result:=h;
end;

function TCOMServer.GetN():integer;
begin
  result:=n;
end;

procedure TCOMServer.SetA(aa:extended);
begin
  a:=aa;
  updateN();
end;

procedure TCOMServer.SetB(bb:extended);
begin
  b:=bb;
  updateN();
end;

procedure TCOMServer.SetH(hh:extended);
begin
  h:=hh;
  updateN();
end;

procedure TCOMServer.TextWord(word:TWordDocument;appl:TWordApplication);
begin
  //appl.Connect; // connect server word
  appl:= TWordApplication.Create(nil);
  word:= TWordDocument.Create(nil);

  appl.Documents.Add(EmptyParam, EmptyParam, EmptyParam, EmptyParam);// add document
  word.ConnectTo(appl.ActiveDocument);// connection WordDocument with WordApplication
  // turn off spell check for faster work
  appl.Options.CheckSpellingAsYouType := False;
  appl.Options.CheckGrammarAsYouType := False;
  // font
  appl.Selection.Font.Size:=16;
  // record date to word
  appl.Selection.TypeText('First limits of integration: ' + FloatToStr(a) + #13#10);
  appl.Selection.TypeText('Second limits of integration: ' + FloatToStr(b) + #13#10);
  appl.Selection.TypeText('Step of integration: ' + FloatToStr(h) + #13#10);

  appl.Visible:=true; // doing word visible
  //appl.Disconnect; // disconnect server word
end;

procedure TCOMServer.TextWord(word:TWordDocument;appl:TWordApplication;str:TStringGrid; chr:TChart);
var
  int:integer;
  end1:OleVariant;
begin
  //appl.Connect; // activate server word
  appl:= TWordApplication.Create(nil);
  word:= TWordDocument.Create(nil);

  appl.Documents.Add(EmptyParam, EmptyParam, EmptyParam, EmptyParam);// add document
  word.ConnectTo(appl.ActiveDocument);// connection WordDocument with WordApplication
  // turn off spell check for faster work
  appl.Options.CheckSpellingAsYouType := False;
  appl.Options.CheckGrammarAsYouType := False;
  // font
  appl.Selection.Font.Size:=16;
  // record date to word
  appl.Selection.TypeText('First limits of integration: ' + FloatToStr(a) + #13#10);
  appl.Selection.TypeText('Second limits of integration: ' + FloatToStr(b) + #13#10);
  appl.Selection.TypeText('Step of integration: ' + FloatToStr(h) + #13#10);
  appl.Selection.Select;
  // created 3 columns  of N lines
  word.Tables.Add(appl.Selection.Range, str.RowCount, 3, EmptyParam, EmptyParam);
  // set the width of columns
  word.Tables.Item(1).Columns.Item(1).Width:=100;
  word.Tables.Item(1).Columns.Item(2).Width:=200;
  word.Tables.Item(1).Columns.Item(3).Width:=200;
  // create column headers
  word.Tables.Item(1).Cell(1,1).select;
  appl.Selection.TypeText('X');
  word.Tables.Item(1).Cell(1,2).select;
  appl.Selection.TypeText('Y1');
  word.Tables.Item(1).Cell(1,3).select;
  appl.Selection.TypeText('Y2');
  // write data to columns
  for int:=1 to str.RowCount do
  begin
    word.Tables.Item(1).Cell(int+1,1).select;
    appl.Selection.TypeText(str.Cells[0,int]);
    word.Tables.Item(1).Cell(int+1,2).select;
    appl.Selection.TypeText(str.Cells[1,int]);
    word.Tables.Item(1).Cell(int+1,3).select;
    appl.Selection.TypeText(str.Cells[2,int]);
  end;

  chr.CopyToClipboardBitmap; // copy chart to clipboard
  end1:=word.Range.End_-1; // end of document
  word.Range(end1).Paste; // past graph in document

  appl.Visible:=true;// doing word visible
  //appl.Disconnect; // disconnect server word
end;

procedure TCOMServer.TextExcel();
var
  Excel,Work:variant;
begin
  Excel:=CreateOleObject('Excel.Application'); // create COM-object
  Excel.Application.EnableEvents := false; // Disable response to external events-to speed up the data transfer process
  Work:= Excel.WorkBooks.Add; // Add workbook
  // write data to columns
  Work.WorkSheets[1].Cells[1, 1]:='First limits of integration';
  Work.WorkSheets[1].Cells[1, 4]:=a;
  Work.WorkSheets[1].Cells[2, 1]:='Second limits of integration';
  Work.WorkSheets[1].Cells[2, 4]:=b;
  Work.WorkSheets[1].Cells[3, 1]:='Step of integration';
  Work.WorkSheets[1].Cells[3, 4]:=h;
  Excel.Visible:=true; // show Excel
end;

procedure TCOMServer.TextExcel(str:TStringGrid;chr:TChart);
var
  int:integer;
  Excel,Work:variant;
begin
  Excel:= CreateOleObject('Excel.Application'); // create COM-object
  Excel.Application.EnableEvents := false; // Disable response to external events-to speed up the data transfer process
  Work:= Excel.WorkBooks.Add; // add workbook
  // write data to columns
  Work.WorkSheets[1].Cells[1, 1]:='First limits of integration';
  Work.WorkSheets[1].Cells[1, 4]:=a;
  Work.WorkSheets[1].Cells[2, 1]:='Second limits of integration';
  Work.WorkSheets[1].Cells[2, 4]:=b;
  Work.WorkSheets[1].Cells[3, 1]:='Step of integration';
  Work.WorkSheets[1].Cells[3, 4]:=h;

  Work.WorkSheets[1].Cells[6, 1]:='X';
  Work.WorkSheets[1].Cells[6, 2]:='Y1';
  Work.WorkSheets[1].Cells[6, 3]:='Y2';

  for int:= 1 to str.RowCount do
  begin
    Work.WorkSheets[1].Cells[int+7, 1]:= str.Cells[0,int];
    Work.WorkSheets[1].Cells[int+7, 2]:= str.Cells[1,int];
    Work.WorkSheets[1].Cells[int+7, 3]:= str.Cells[2,int];
  end;
  Work.WorkSheets[1].Cells[5, 10].Select;
  chr.SaveToBitmapFile(GetCurrentDir + '\Excel.bmp');// save graph to file
  Work.WorkSheets[1].Shapes.AddPicture(GetCurrentDir + '\Excel.bmp',True, True, 300, 100,chr.Width * 0.8,chr.Height * 0.8); // inserrt picture graph in excel
  Excel.Visible := true;// show Excel
end;

procedure TCOMServer.writeTextFile(filename:string);
var
  f:TextFile;
begin
  AssignFile(f,filename); // associate a file variable with the file name
  Rewrite(f); // open for overwrite
  // write in file
  writeln(f,a:5:2);
  writeln(f,b:5:2);
  writeln(f,h:5:2);
  writeln(f,n:5);
  CloseFile(f); // close file
end;


procedure TCOMServer.writeTextFile(filename:string;str:TStringGrid);
var
  f:TextFile;
  int:integer;
begin
  AssignFile(f,filename);// associate a file variable with the file name
  Rewrite(f);// open for overwrite
  // write in file
  writeln(f,'First limits of integration: ', a:5:2);
  writeln(f,'Second limits of integration: ', b:5:2);
  writeln(f,'Step of integration: ', h:5:2);
  writeln(f,'Number of results: ', n:5);

  writeln(f,'Grid: ');
  // write if file data like a table
  writeln(f,str.Cells[0,0],'  ',str.Cells[1,0],';    ', str.Cells[2,0]);
  for int:=1 to str.RowCount do
  begin
    writeln(f,str.Cells[0,int],': ',str.Cells[1,int],';    ', str.Cells[2,int]);
  end;
  CloseFile(f); // close file
end;

procedure TCOMServer.readTextFile(filename:string);
var
  f:TextFile;
begin
  AssignFile(f,filename);// associate a file variable with the file name
  Reset(f);  // open for reading
  // read from file
  readln(f,a);
  readln(f,b);
  readln(f,h);
  readln(f,n);
  CloseFile(f); // close file
end;

initialization
  TComObjectFactory.Create(ComServer, TCOMServer, Class_COMServer,
    '', '', ciMultiInstance, tmApartment);
end.
