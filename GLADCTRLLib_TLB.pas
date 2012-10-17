unit GLADCTRLLib_TLB;

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
// File generated on 2006-7-1 23:00:02 from Type Library described below.

// ************************************************************************  //
// Type Lib: D:\Delphi\TLBDBG\Debug\gladctrl.ocx (1)
// LIBID: {8F8741A9-CDCB-4527-B807-C9688DAF5823}
// LCID: 0
// Helpfile: C:\Program Files\Globallink\Game\share\GLAdCtrl.hlp
// HelpString: GLAdCtrl ActiveX Control module
// DepndLst: 
//   (1) v2.0 stdole, (C:\WINDOWS\system32\stdole2.tlb)
// ************************************************************************ //
{$TYPEDADDRESS OFF} // Unit must be compiled without type-checked pointers. 
{$WARN SYMBOL_PLATFORM OFF}
{$WRITEABLECONST ON}
{$VARPROPSETTER ON}
interface

uses Windows, ActiveX, Classes, Graphics, OleCtrls, OleServer, StdVCL, Variants;
  


// *********************************************************************//
// GUIDS declared in the TypeLibrary. Following prefixes are used:        
//   Type Libraries     : LIBID_xxxx                                      
//   CoClasses          : CLASS_xxxx                                      
//   DISPInterfaces     : DIID_xxxx                                       
//   Non-DISP interfaces: IID_xxxx                                        
// *********************************************************************//
const
  // TypeLibrary Major and minor versions
  GLADCTRLLibMajorVersion = 1;
  GLADCTRLLibMinorVersion = 0;

  LIBID_GLADCTRLLib: TGUID = '{8F8741A9-CDCB-4527-B807-C9688DAF5823}';

  DIID__DGLAdCtrl: TGUID = '{C49CE1AB-5A36-4D2B-8EEB-67DFE9B6EE8D}';
  DIID__DGLAdCtrlEvents: TGUID = '{63D47059-7FB6-4776-8D39-38BECBF809A1}';
  CLASS_GLAdCtrl: TGUID = '{11DC869E-8872-401D-BE9C-8B139E766743}';
type

// *********************************************************************//
// Forward declaration of types defined in TypeLibrary                    
// *********************************************************************//
  _DGLAdCtrl = dispinterface;
  _DGLAdCtrlEvents = dispinterface;

// *********************************************************************//
// Declaration of CoClasses defined in Type Library                       
// (NOTE: Here we map each CoClass to its Default Interface)              
// *********************************************************************//
  GLAdCtrl = _DGLAdCtrl;


// *********************************************************************//
// DispIntf:  _DGLAdCtrl
// Flags:     (4112) Hidden Dispatchable
// GUID:      {C49CE1AB-5A36-4D2B-8EEB-67DFE9B6EE8D}
// *********************************************************************//
  _DGLAdCtrl = dispinterface
    ['{C49CE1AB-5A36-4D2B-8EEB-67DFE9B6EE8D}']
    procedure InitCtrl(const pszIniFile: WideString; iChannelID: Integer); dispid 1;
    function GetChannelID: Integer; dispid 2;
    procedure SetChannelID(iChannelID: Integer); dispid 3;
    function GetCtrlOwner: Integer; dispid 4;
    procedure SetCtrlOwner(hOwner: Integer); dispid 5;
    procedure Stop; dispid 6;
    procedure SetCtrlOwner2(hOwner: Integer); dispid 7;
    function Monica(clt: Integer): Integer; dispid 8;
    procedure PlayAdvert(ADID: Integer; const pszPrimaryURL: WideString; 
                         const pszPopupURL: WideString; const TooltipText: WideString; 
                         PopupWindowHeight: Integer; PopupWindowWidth: Integer); dispid 9;
    procedure PLayFSAdvert(ADID: Integer; const pszURL: WideString; hScreen: Integer; 
                           iSpeed: Integer; iLife: Integer); dispid 10;
    procedure PlayMBAdvert(ADID: Integer; const pszURL: WideString; iSpeed: Integer; 
                           hScreen: Integer; iWidth: Integer; iHeight: Integer; iLife: Integer; 
                           bSysmenu: Integer); dispid 11;
    procedure PlayFSAdvert2(ADID: Integer; const pszURL: WideString; iWidth: Integer; 
                            iHeight: Integer; iSpeed: Integer; iLife: Integer; bSysmenu: Integer); dispid 12;
    procedure AboutBox; dispid -552;
  end;

// *********************************************************************//
// DispIntf:  _DGLAdCtrlEvents
// Flags:     (4096) Dispatchable
// GUID:      {63D47059-7FB6-4776-8D39-38BECBF809A1}
// *********************************************************************//
  _DGLAdCtrlEvents = dispinterface
    ['{63D47059-7FB6-4776-8D39-38BECBF809A1}']
  end;


// *********************************************************************//
// OLE Control Proxy class declaration
// Control Name     : TGLAdCtrl
// Help String      : GLAdCtrl Control
// Default Interface: _DGLAdCtrl
// Def. Intf. DISP? : Yes
// Event   Interface: _DGLAdCtrlEvents
// TypeFlags        : (34) CanCreate Control
// *********************************************************************//
  TGLAdCtrl = class(TOleControl)
  private
    FIntf: _DGLAdCtrl;
    function  GetControlInterface: _DGLAdCtrl;
  protected
    procedure CreateControl;
    procedure InitControlData; override;
  public
    procedure InitCtrl(const pszIniFile: WideString; iChannelID: Integer);
    function GetChannelID: Integer;
    procedure SetChannelID(iChannelID: Integer);
    function GetCtrlOwner: Integer;
    procedure SetCtrlOwner(hOwner: Integer);
    procedure Stop;
    procedure SetCtrlOwner2(hOwner: Integer);
    function Monica(clt: Integer): Integer;
    procedure PlayAdvert(ADID: Integer; const pszPrimaryURL: WideString; 
                         const pszPopupURL: WideString; const TooltipText: WideString; 
                         PopupWindowHeight: Integer; PopupWindowWidth: Integer);
    procedure PLayFSAdvert(ADID: Integer; const pszURL: WideString; hScreen: Integer; 
                           iSpeed: Integer; iLife: Integer);
    procedure PlayMBAdvert(ADID: Integer; const pszURL: WideString; iSpeed: Integer; 
                           hScreen: Integer; iWidth: Integer; iHeight: Integer; iLife: Integer; 
                           bSysmenu: Integer);
    procedure PlayFSAdvert2(ADID: Integer; const pszURL: WideString; iWidth: Integer; 
                            iHeight: Integer; iSpeed: Integer; iLife: Integer; bSysmenu: Integer);
    procedure AboutBox;
    property  ControlInterface: _DGLAdCtrl read GetControlInterface;
    property  DefaultInterface: _DGLAdCtrl read GetControlInterface;
  published
    property Anchors;
    property  TabStop;
    property  Align;
    property  DragCursor;
    property  DragMode;
    property  ParentShowHint;
    property  PopupMenu;
    property  ShowHint;
    property  TabOrder;
    property  Visible;
    property  OnDragDrop;
    property  OnDragOver;
    property  OnEndDrag;
    property  OnEnter;
    property  OnExit;
    property  OnStartDrag;
  end;

procedure Register;

resourcestring
  dtlServerPage = 'ActiveX';

  dtlOcxPage = 'ActiveX';

implementation

uses ComObj;

procedure TGLAdCtrl.InitControlData;
const
  CControlData: TControlData2 = (
    ClassID: '{11DC869E-8872-401D-BE9C-8B139E766743}';
    EventIID: '';
    EventCount: 0;
    EventDispIDs: nil;
    LicenseKey: nil (*HR:$80004005*);
    Flags: $00000000;
    Version: 401);
begin
  ControlData := @CControlData;
end;

procedure TGLAdCtrl.CreateControl;

  procedure DoCreate;
  begin
    FIntf := IUnknown(OleObject) as _DGLAdCtrl;
  end;

begin
  if FIntf = nil then DoCreate;
end;

function TGLAdCtrl.GetControlInterface: _DGLAdCtrl;
begin
  CreateControl;
  Result := FIntf;
end;

procedure TGLAdCtrl.InitCtrl(const pszIniFile: WideString; iChannelID: Integer);
begin
  DefaultInterface.InitCtrl(pszIniFile, iChannelID);
end;

function TGLAdCtrl.GetChannelID: Integer;
begin
  Result := DefaultInterface.GetChannelID;
end;

procedure TGLAdCtrl.SetChannelID(iChannelID: Integer);
begin
  DefaultInterface.SetChannelID(iChannelID);
end;

function TGLAdCtrl.GetCtrlOwner: Integer;
begin
  Result := DefaultInterface.GetCtrlOwner;
end;

procedure TGLAdCtrl.SetCtrlOwner(hOwner: Integer);
begin
  DefaultInterface.SetCtrlOwner(hOwner);
end;

procedure TGLAdCtrl.Stop;
begin
  DefaultInterface.Stop;
end;

procedure TGLAdCtrl.SetCtrlOwner2(hOwner: Integer);
begin
  DefaultInterface.SetCtrlOwner2(hOwner);
end;

function TGLAdCtrl.Monica(clt: Integer): Integer;
begin
  Result := DefaultInterface.Monica(clt);
end;

procedure TGLAdCtrl.PlayAdvert(ADID: Integer; const pszPrimaryURL: WideString; 
                               const pszPopupURL: WideString; const TooltipText: WideString; 
                               PopupWindowHeight: Integer; PopupWindowWidth: Integer);
begin
  DefaultInterface.PlayAdvert(ADID, pszPrimaryURL, pszPopupURL, TooltipText, PopupWindowHeight, 
                              PopupWindowWidth);
end;

procedure TGLAdCtrl.PLayFSAdvert(ADID: Integer; const pszURL: WideString; hScreen: Integer; 
                                 iSpeed: Integer; iLife: Integer);
begin
  DefaultInterface.PLayFSAdvert(ADID, pszURL, hScreen, iSpeed, iLife);
end;

procedure TGLAdCtrl.PlayMBAdvert(ADID: Integer; const pszURL: WideString; iSpeed: Integer; 
                                 hScreen: Integer; iWidth: Integer; iHeight: Integer; 
                                 iLife: Integer; bSysmenu: Integer);
begin
  DefaultInterface.PlayMBAdvert(ADID, pszURL, iSpeed, hScreen, iWidth, iHeight, iLife, bSysmenu);
end;

procedure TGLAdCtrl.PlayFSAdvert2(ADID: Integer; const pszURL: WideString; iWidth: Integer; 
                                  iHeight: Integer; iSpeed: Integer; iLife: Integer; 
                                  bSysmenu: Integer);
begin
  DefaultInterface.PlayFSAdvert2(ADID, pszURL, iWidth, iHeight, iSpeed, iLife, bSysmenu);
end;

procedure TGLAdCtrl.AboutBox;
begin
  DefaultInterface.AboutBox;
end;

procedure Register;
begin
  RegisterComponents(dtlOcxPage, [TGLAdCtrl]);
end;

end.
