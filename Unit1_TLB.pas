unit Unit1_TLB;

interface
uses StdCtrls, ExtCtrls, Graphics, Math, SysUtils, Forms, Grids, WordXP,
     OleServer, ExcelXP, Types, Chart, ComObj;

const
  Class_COMServer: TGUID = '{59CC5BC2-49AC-4477-83EA-D378F37DDBE8}';

type
  ICOMServer = interface ['{C3F5B148-C9F3-4001-8770-CAA3A5A7AF27}']

    function GetA():extended;
    function GetB():extended;
    function GetH():extended;
    function GetN():integer;

    procedure SetA(aa:extended);
    procedure SetB(bb:extended);
    procedure SetH(hh:extended);


    procedure Create(aa,bb,hh:extended);stdcall;
//    procedure Destroy();stdcall;

    property Low:extended read GetA write SetA;
    property Top:extended read GetB write SetB;
    property Step:extended read GetH write SetH;
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

end.
