unit Project_TLB;

// ************************************************************************ //
// WARNING                                                                    
// -------                                                                    
// The types declared in this file were generated from data read from a       
// Type Library. If this type library is explicitly or indirectly (via        
// another type library referring to this type library) re-imported, or the   
// 'Refresh' command of the Type Library Editor activated while editing the   
// Type Library, the contents of this file will be regenerated and all        
// manual modifications will be lost.                                         
// ************************************************************************ //

// PASTLWTR : 1.2
// File generated on 02/06/2018 14:37:57 from Type Library described below.

// ************************************************************************  //
// Type Lib: C:\Users\ton9\Desktop\course work delphi\com\Project.tlb (1)
// LIBID: {350DB911-C8FC-42EF-A75B-8EC4586558C2}
// LCID: 0
// Helpfile: 
// HelpString: Project Library
// DepndLst: 
//   (1) v2.0 stdole, (C:\Windows\SysWOW64\stdole2.tlb)
// ************************************************************************ //
{$TYPEDADDRESS OFF} // Unit must be compiled without type-checked pointers. 
{$WARN SYMBOL_PLATFORM OFF}
{$WRITEABLECONST ON}
{$VARPROPSETTER ON}
interface

uses Windows, ActiveX, Classes, Graphics, StdVCL, Variants;
  

// *********************************************************************//
// GUIDS declared in the TypeLibrary. Following prefixes are used:        
//   Type Libraries     : LIBID_xxxx                                      
//   CoClasses          : CLASS_xxxx                                      
//   DISPInterfaces     : DIID_xxxx                                       
//   Non-DISP interfaces: IID_xxxx                                        
// *********************************************************************//
const
  // TypeLibrary Major and minor versions
  ProjectMajorVersion = 1;
  ProjectMinorVersion = 0;

  LIBID_Project: TGUID = '{350DB911-C8FC-42EF-A75B-8EC4586558C2}';

  IID_ICOMServer: TGUID = '{C3F5B148-C9F3-4001-8770-CAA3A5A7AF27}';
  CLASS_COMServer: TGUID = '{59CC5BC2-49AC-4477-83EA-D378F37DDBE8}';
type

// *********************************************************************//
// Forward declaration of types defined in TypeLibrary                    
// *********************************************************************//
  ICOMServer = interface;

// *********************************************************************//
// Declaration of CoClasses defined in Type Library                       
// (NOTE: Here we map each CoClass to its Default Interface)              
// *********************************************************************//
  COMServer = ICOMServer;


// *********************************************************************//
// Interface: ICOMServer
// Flags:     (256) OleAutomation
// GUID:      {C3F5B148-C9F3-4001-8770-CAA3A5A7AF27}
// *********************************************************************//
  ICOMServer = interface(IUnknown)
    ['{C3F5B148-C9F3-4001-8770-CAA3A5A7AF27}']
  end;

// *********************************************************************//
// The Class CoCOMServer provides a Create and CreateRemote method to          
// create instances of the default interface ICOMServer exposed by              
// the CoClass COMServer. The functions are intended to be used by             
// clients wishing to automate the CoClass objects exposed by the         
// server of this typelibrary.                                            
// *********************************************************************//
  CoCOMServer = class
    class function Create: ICOMServer;
    class function CreateRemote(const MachineName: string): ICOMServer;
  end;

implementation

uses ComObj;

class function CoCOMServer.Create: ICOMServer;
begin
  Result := CreateComObject(CLASS_COMServer) as ICOMServer;
end;

class function CoCOMServer.CreateRemote(const MachineName: string): ICOMServer;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_COMServer) as ICOMServer;
end;

end.
