library Project;

uses
  ComServ,
  Unit1 in 'Unit1.pas',
  Unit1_TLB in 'Unit1_TLB.pas';

exports
  DllGetClassObject,
  DllCanUnloadNow,
  DllRegisterServer,
  DllUnregisterServer;

{$R *.TLB}

{$R *.RES}

begin
end.
