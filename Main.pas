unit Main;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ToolWin, ComCtrls, ImgList, WordXP, OleServer, Menus, ExtCtrls,
  StdCtrls,

  ShellApi, Grids, TeeProcs, TeEngine, Chart, IniFiles, Series, Buttons,
  jpeg, ComObj,

  TIntegralUnit, Loading, Unit1_TLB;

type
  TShow = procedure();

  TForm1 = class(TForm)
    StatusBar1: TStatusBar;
    OpenDialog1: TOpenDialog;
    SaveDialog1: TSaveDialog;
    Timer1: TTimer;
    WordDocument1: TWordDocument;
    WordApplication1: TWordApplication;
    ImageList1: TImageList;

    MainMenu1: TMainMenu;
      File1: TMenuItem;
        Save1: TMenuItem;
          Experimentaldata1: TMenuItem;
          Resultsofanexperiment1: TMenuItem;
          Projectsettings1: TMenuItem;
        Load1: TMenuItem;
          Data1: TMenuItem;
          Projectsettings2: TMenuItem;
        Exit1: TMenuItem;
      Integration1: TMenuItem;
        InWord1: TMenuItem;
        InExcel1: TMenuItem;
      Selectinganobject1: TMenuItem;
        Classic1: TMenuItem;
        COMObject1: TMenuItem;
      Reference1: TMenuItem;
        Help1: TMenuItem;
        About1: TMenuItem;
        Presentation1: TMenuItem;
        Calculate1: TMenuItem;
      Language1: TMenuItem;
        English1: TMenuItem;
        Russian1: TMenuItem;

    PageControl1: TPageControl;
      TabSheet1: TTabSheet;
      TabSheet2: TTabSheet;
      TabSheet3: TTabSheet;

    ToolBar1: TToolBar;
      ToolButton1: TToolButton;
      ToolButton1_2: TToolButton;
      ToolButton2: TToolButton;
      ToolButton2_1: TToolButton;
      ToolButton3: TToolButton;
      ToolButton3_1: TToolButton;
      ToolButton4: TToolButton;
      ToolButton4_1: TToolButton;
      ToolButton5: TToolButton;
      ToolButton5_1: TToolButton;
      ToolButton6: TToolButton;
      ToolButton6_1: TToolButton;
      ToolButton7: TToolButton;
      ToolButton7_1: TToolButton;
      ToolButton8: TToolButton;

    Chart1: TChart;
    StringGrid1: TStringGrid;
    Edit1: TEdit;
    Edit2: TEdit;
    Edit3: TEdit;
    Edit4: TEdit;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    Label6: TLabel;
    Label7: TLabel;
    Label8: TLabel;
    Label9: TLabel;
    Label5: TLabel;
    BitBtn1: TBitBtn;
    BitBtn3: TBitBtn;
    BitBtn5: TBitBtn;
    BitBtn2: TBitBtn;
    BitBtn4: TBitBtn;
    Series1: TLineSeries;
    Series2: TLineSeries;
    Image1: TImage;
    Image2: TImage;
    Image3: TImage;
    Image4: TImage;
    Image5: TImage;

    function isNumber(Key: Char):Boolean; // Is it the number?
    function Empty(s:TStringGrid):Boolean; // Is the StringGrid empty?
                    
    procedure FormShow(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure Timer1Timer(Sender: TObject); // timer for StatusBar
    procedure EditKeyPress(Sender: TObject; var Key: Char); // for limit input

    procedure BitBtn1Click(Sender: TObject); // fill class fields
    procedure BitBtn2Click(Sender: TObject); // fill in the table
    procedure BitBtn3Click(Sender: TObject); // clear StringGrid
    procedure BitBtn4Click(Sender: TObject); // draw a graph
    procedure BitBtn5Click(Sender: TObject); // clear Graph

    procedure Classic1Click(Sender: TObject);
    procedure COMObject1Click(Sender: TObject);
    procedure Help1Click(Sender: TObject);
    procedure Exit1Click(Sender: TObject);
    procedure Calculate1Click(Sender: TObject);
    procedure Presentation1Click(Sender: TObject);
    procedure InExcel1Click(Sender: TObject);
    procedure InWord1Click(Sender: TObject);
    procedure Experimentaldata1Click(Sender: TObject);
    procedure Resultsofanexperiment1Click(Sender: TObject);
    procedure Projectsettings1Click(Sender: TObject);
    procedure Data1Click(Sender: TObject);
    procedure Projectsettings2Click(Sender: TObject);
    procedure Russian1Click(Sender: TObject);
    procedure English1Click(Sender: TObject);
    procedure About1Click(Sender: TObject);

  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form1: TForm1;
  Server:IComServer; //interface COM
  ComObjUse:boolean;//flag for COM
  Integral:TIntegral;
  n:integer;
  rr:extended;

implementation

{$R *.dfm}

function TForm1.isNumber(Key: Char):Boolean;
begin
  result:= key in['0'..'9', #8, ',', '.', '-'];
end;

function TForm1.Empty(s:TStringGrid):Boolean;
var i, count:integer;
begin
  count:=-1;
  for i:=1 to s.RowCount do
  begin
    if(s.Cells[1,i] <> '') then inc(count);
  end;
  if(count > -1) then result:=true
  else result:=false;
end;

procedure TForm1.FormShow(Sender: TObject);
begin
  Form2.ShowModal;
end;

procedure TForm1.FormCreate(Sender: TObject);
begin
  Integral := TIntegral.Create(StrToFloat(Edit1.Text),StrToFloat(Edit2.Text),
    StrToFloat(Edit3.Text));
  n:=trunc((Integral.Top-Integral.Low)/Integral.Step + 1);
  Edit4.Text:=IntToStr(n);
  StringGrid1.RowCount:=5;
  StringGrid1.Cells[0, 0] := 'Y(x)';
  StringGrid1.Cells[1, 0] := 'Result 1';
  StringGrid1.Cells[2, 0] := 'Result 2';

  if(Russian1.Checked) then
    StatusBar1.Panels[0].Text := 'Классический объект'
  else
    StatusBar1.Panels[0].Text := 'The Classic object';
end;

procedure TForm1.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  Server := nil;
  Integral.Destroy;
  WordApplication1.Quit;
end;

procedure TForm1.Timer1Timer(Sender: TObject);
var
  day:TDateTime;
begin
  day:=Now;
  StatusBar1.Panels[3].Text:=TimeToStr(day);
  StatusBar1.Panels[2].Text:=FormatDateTime('dddd dd.mm.yyyy',day);
end;

procedure TForm1.EditKeyPress(Sender: TObject; var Key: Char);
begin
  if (not isNumber(key)) then key:=#0;
end;

procedure TForm1.BitBtn1Click(Sender: TObject);
begin
  if((Edit1.Text <> '') and  (Edit2.Text <> '') and (Edit3.Text <> '')) then
  begin
    if(ComObjUse) then
    begin
      Server.Low:=StrToFloat(Edit1.Text);
      Server.Top:=StrToFloat(Edit2.Text);
      Server.Step:=StrToFloat(Edit3.Text);
      Edit4.Text:=IntToStr(Server.NumberResults);
    end
    else
    begin
      if(Integral <> nil) then
      begin
        Integral.Low:=StrToFloat(Edit1.Text);
        Integral.Top:=StrToFloat(Edit2.Text);
        Integral.Step:=StrToFloat(Edit3.Text);
        Edit4.Text:=IntToStr(Integral.NumberResults);
      end;
    end;
  end
  else ShowMessage('Fill in all the fields');
end;

procedure TForm1.BitBtn2Click(Sender: TObject);
begin
  if(ComObjUse) then Server.FillInTheTable(StringGrid1)
  else
  begin
    if(Integral <> nil) then Integral.FillInTheTable(StringGrid1)
    else ShowMessage('Create object');
  end;
end;

procedure TForm1.BitBtn3Click(Sender: TObject);
var
  i:integer;
begin
  for i:=1 to StringGrid1.RowCount do
  begin
    StringGrid1.Rows[i].Clear;
  end;
  StringGrid1.RowCount:=5;
end;

procedure TForm1.BitBtn4Click(Sender: TObject);
begin
  Series1.Clear;
  Series2.Clear;

  if(ComObjUse) then
  begin
    if(Server <> nil) and (Empty(StringGrid1)) then
    begin
      Server.displayOnGraph(StringGrid1,Chart1);
    end
    else ShowMessage('Fill in the table');
  end
  else
  begin
    if(Integral <> nil) and (Empty(StringGrid1)) then
    begin
      Integral.displayOnGraph(StringGrid1,Chart1);
    end
    else ShowMessage('Fill in the table');
  end
end;
procedure TForm1.BitBtn5Click(Sender: TObject);
begin
  Series1.Clear;
  Series2.Clear;
end;

procedure TForm1.Classic1Click(Sender: TObject);
begin
  COMObject1.Checked:=false;
  Classic1.Checked:=true;
  StatusBar1.Panels[0].Text := 'The Classic object';
end;


procedure TForm1.COMObject1Click(Sender: TObject);
begin
  COMObject1.Checked:=true;
  Classic1.Checked:=false;

  ComObjUse:=not ComObjUse;

  if ComObjUse then
  begin
    StatusBar1.Panels[0].Text := 'COM-Object';

    if Server = nil then
    begin
      Server:=CreateComObject(Class_COMServer) as ICOMServer;
    end;

    if Server <> nil then
    begin
      Server.Create(StrToFloat(Edit1.Text),StrToFloat(Edit2.Text),StrToFloat(Edit3.Text));
    end;
  end;
end;

procedure TForm1.Help1Click(Sender: TObject);
begin
  ShellExecute(Handle, 'open', 'help\help.chm', nil, nil, SW_SHOWNORMAL);
end;

procedure TForm1.Exit1Click(Sender: TObject);
begin
  WordDocument1.Destroy;
  WordApplication1.Quit;
  close;
end;

procedure TForm1.Calculate1Click(Sender: TObject);
begin
  ShellExecute(Handle, 'open', 'calc.exe', nil, nil, SW_SHOWNORMAL);
end;

procedure TForm1.Presentation1Click(Sender: TObject);
begin
  ShellExecute(Handle, 'open', 'presentation\presentation.ppsx', nil, nil, SW_SHOWNORMAL);
end;

procedure TForm1.InExcel1Click(Sender: TObject);
begin
  if(ComObjUse) then
  begin
    if(Empty(StringGrid1)) then
    begin
      Server.TextExcel(StringGrid1,Chart1);
    end
    else Server.TextExcel();
  end
  else
  begin
    if(Integral <> nil) then
    begin
      if(Empty(StringGrid1)) then
      begin
        Integral.TextExcel(StringGrid1, Chart1);
      end
      else Integral.TextExcel();
    end
    else ShowMessage('Create object');
  end;
end;

procedure TForm1.InWord1Click(Sender: TObject);
begin
  if(ComObjUse) then
  begin
    if(Empty(StringGrid1)) then
    begin
      Server.TextWord(WordDocument1,WordApplication1,StringGrid1,Chart1);
    end
    else Server.TextWord(WordDocument1,WordApplication1);
  end
  else
  begin
    if(Integral <> nil) then
    begin
      if(Empty(StringGrid1)) then
      begin
        Integral.TextWord(WordDocument1,WordApplication1,StringGrid1, Chart1);
      end
      else Integral.TextWord(WordDocument1,WordApplication1);
    end
    else ShowMessage('Create object');
  end;
end;

procedure TForm1.Experimentaldata1Click(Sender: TObject);
begin
  if(ComObjUse) then
    Server.writeTextFile('Data.txt')
  else
  begin
    if(Integral <> nil) then
    begin
      Integral.writeTextFile('Data.txt')
    end
    else ShowMessage('Create object');
  end;
end;

procedure TForm1.Resultsofanexperiment1Click(Sender: TObject);
begin
  if(ComObjUse) then
  begin
    if(Empty(StringGrid1)) then
    begin
      if(SaveDialog1.Execute) then
        Server.writeTextFile(SaveDialog1.FileName,StringGrid1)
      else  ShowMessage('Choose file');
    end
    else ShowMessage('Filling in the table');
  end
  else
  begin
    if(Integral <> nil) then
    begin
      if(Empty(StringGrid1)) then
      begin
        if(SaveDialog1.Execute) then
          Integral.writeTextFile(SaveDialog1.FileName,StringGrid1)
        else  ShowMessage('Choose file');
      end
      else ShowMessage('Filling in the table');
    end
    else ShowMessage('Create object');
  end;
end;

procedure TForm1.Projectsettings1Click(Sender: TObject);
var
  ini:TIniFile;
begin
  if(SaveDialog1.Execute) then
  begin
    ini:=TIniFile.Create(Form1.SaveDialog1.FileName);

    ini.WriteString('Forms', 'Caption', Form1.Caption);
    ini.WriteInteger('Forms','Height', Form1.Height);
    ini.WriteInteger('Forms','Width', Form1.Width);
    ini.WriteInteger('Forms','Left', Form1.Left);
    ini.WriteInteger('Forms','Top', Form1.Top);

    ini.WriteString('Edits', 'Edit1Text', form1.Edit1.Text);
    ini.WriteString('Edits', 'Edit2Text', form1.Edit2.Text);
    ini.WriteString('Edits', 'Edit3Text', form1.Edit3.Text);
    ini.WriteString('Edits', 'Edit4Text', form1.Edit4.Text);

    ini.Free;
  end
  else  ShowMessage('Choose file');
end;

procedure TForm1.Data1Click(Sender: TObject);
begin
  if(ComObjUse) then
  begin
    if(FileExists('Data.txt')) then
    begin
      Server.readTextFile('Data.txt');
      Edit1.Text:=FloatToStr(Server.Low);
      Edit2.Text:=FloatToStr(Server.Top);
      Edit3.Text:=FloatToStr(Server.Step);
      Edit4.Text:=IntToStr(Server.NumberResults);
    end
    else  ShowMessage('File does not exist');
  end
  else
  begin
    if(FileExists('Data.txt')) then
    begin
      n:=trunc((Integral.Top-Integral.Low)/Integral.Step+1);
      Integral.readTextFile('Data.txt');
      Edit1.Text:=FloatToStr(Integral.Low);
      Edit2.Text:=FloatToStr(Integral.Top);
      Edit3.Text:=FloatToStr(Integral.Step);
      Edit4.Text:=IntToStr(Integral.NumberResults);
    end
    else  ShowMessage('File does not exist');
  end;
end;

procedure TForm1.Projectsettings2Click(Sender: TObject);
var
  ini:TIniFile;
begin
  if(OpenDialog1.Execute) then
  begin
    ini:=TIniFile.Create(Form1.OpenDialog1.FileName);

    Form1.Caption:=ini.ReadString('Forms', 'Caption', 'Error');
    Form1.Height:=ini.ReadInteger('Forms','Height', 500);
    Form1.Width:=ini.ReadInteger('Forms','Width', 1000);
    Form1.Left:=ini.ReadInteger('Forms','Left', 100);
    Form1.Top:= ini.ReadInteger('Forms','Top', 100);

    Form1.Edit1.Text:=ini.ReadString('Edits', 'Edit1Text', 'Error');
    Form1.Edit2.Text:=ini.ReadString('Edits', 'Edit2Text', 'Error');
    Form1.Edit3.Text:=ini.ReadString('Edits', 'Edit3Text', 'Error');
    Form1.Edit4.Text:=ini.ReadString('Edits', 'Edit4Text', 'Error');

    ini.Free;
  end
  else  ShowMessage('Choose file');
end;

procedure TForm1.Russian1Click(Sender: TObject);
begin
  English1.Checked:=false;
  Russian1.Checked:=true;

  Form1.Caption := 'Курсовая работа Кучма А.П. Решение линейн. интегр. уравнения Вольтерра I рода';

  File1.Caption := 'Файл';
    Save1.Caption := 'Сохранить';
      Experimentaldata1.Caption := 'Экспериментальные данные';
      Resultsofanexperiment1.Caption := 'Результаты эксперимента';
      Projectsettings1.Caption := 'Настройки проекта';
    Load1.Caption := 'Загрузить';
      Data1.Caption := 'Данные';
      Projectsettings2.Caption := 'Настройки проекта';
    Exit1.Caption := 'Выход';

  Integration1.Caption := 'Интеграция';
    InWord1.Caption := 'В Ворд';
    InExcel1.Caption := 'В Эксель';

  Selectinganobject1.Caption := 'Выбор обьекта';
    Classic1.Caption := 'Классический';
    COMObject1.Caption := 'COM-Обьект';

  Reference1.Caption := 'Ссылка';
    Help1.Caption := 'Помощь';
    About1.Caption := 'О программе';
    Presentation1.Caption := 'Презентация';
    Calculate1.Caption := 'Калькулятор';

  Language1.Caption := 'Язык';
    English1.Caption := 'Английский';
    Russian1.Caption := 'Русский';

  TabSheet1.Caption := 'Информация';
    Label5.Caption := 'Введите информацию:';
    Label6.Caption := 'Уравнение Вольтерра первого рода:';
    Label7.Caption := 'Точное решение уравнения:';
    Label8.Caption := 'Как вычисляются значения методом квадратурных формул:';
    Label9.Caption := 'Если k(x,s) и f(x):';
    BitBtn1.Caption := 'Занести данные';

  TabSheet2.Caption := 'Таблица результатов';
    BitBtn2.Caption := 'Заполнить таблицу';
    BitBtn3.Caption := 'Очистить';

  TabSheet3.Caption := 'Графики';
    BitBtn4.Caption := 'Нарисовать график';
    BitBtn5.Caption := 'Очистить';
end;

procedure TForm1.English1Click(Sender: TObject);
begin
  English1.Checked:=true;
  Russian1.Checked:=false;

  Form1.Caption := 'Course work Kuchma A.P. Solution of the linear integral Volterra equation of the first kind';

  File1.Caption := 'File';
    Save1.Caption := 'Save';
      Experimentaldata1.Caption := 'Experimental data';
      Resultsofanexperiment1.Caption := 'Result of an experiment';
      Projectsettings1.Caption := 'Project settings';
    Load1.Caption := 'Load';
      Data1.Caption := 'Data';
      Projectsettings2.Caption := 'Project settings';
    Exit1.Caption := 'Exit';

  Integration1.Caption := 'Integration';
    InWord1.Caption := 'In Word';
    InExcel1.Caption := 'In Excel';

  Selectinganobject1.Caption := 'Selecting an object';
    Classic1.Caption := 'Classic';
    COMObject1.Caption := 'COM-Object';

  Reference1.Caption := 'Reference';
    Help1.Caption := 'Help';
    About1.Caption := 'About';
    Presentation1.Caption := 'Presentaion';
    Calculate1.Caption := 'Calculate';

  Language1.Caption := 'Language';
    English1.Caption := 'English';
    Russian1.Caption := 'Russian';

  TabSheet1.Caption := 'Information';
    Label5.Caption := 'Enter the information:';
    Label6.Caption := 'Volterra equation of the first kind:';
    Label7.Caption := 'The exact solution of equation:';
    Label8.Caption := 'How the values are calculated by the method of quadrature formulas:';
    Label9.Caption := 'If k(x,s) and f(x):';
    BitBtn1.Caption := 'Enter';

  TabSheet2.Caption := 'Table of results';
    BitBtn2.Caption := 'Fill in the table';
    BitBtn3.Caption := 'Clear';

  TabSheet3.Caption := 'Graphics';
    BitBtn4.Caption := 'Draw a graph';
    BitBtn5.Caption := 'Clear';
end;

procedure TForm1.About1Click(Sender: TObject);
var
  H:THandle;
  Show:TShow;
begin
  H := loadLibrary('dll\about\loading.dll' );

  if (H <> 0) then
  begin
    @Show := getProcAddress (H, 'ShowLoading');
    
    if Assigned(Show) then
    begin
      Show;
    end
  end
  else ShowMessage('Dll not found');

  freeLibrary(H);
end;

end.
