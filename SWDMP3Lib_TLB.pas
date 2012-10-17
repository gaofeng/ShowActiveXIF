unit SWDMP3Lib_TLB;

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
// File generated on 2006-7-3 9:24:32 from Type Library described below.

// ************************************************************************  //
// Type Lib: C:\Documents and Settings\a\桌面\TLBDBG\新建文件夹\Debug\test.ocx (1)
// LIBID: {830A4C67-4C87-481F-B670-66A07C83CBF4}
// LCID: 0
// Helpfile: 
// HelpString: MP3 Encoder & Decoder Controls 1.0
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
  SWDMP3LibMajorVersion = 1;
  SWDMP3LibMinorVersion = 0;

  LIBID_SWDMP3Lib: TGUID = '{830A4C67-4C87-481F-B670-66A07C83CBF4}';

  DIID__ISwdAudioEncoderEvents: TGUID = '{F9ADEB91-8A20-4CB8-9D2D-3B7E96A09E8A}';
  DIID__ISwdAudioDecoderEvents: TGUID = '{AD6A6218-BEF3-481E-8E4A-95B312CCF418}';
  IID_ISwdAudioEncoder: TGUID = '{9A45E4AB-68A2-4C85-B317-394351CC3330}';
  CLASS_SwdMP3Encoder: TGUID = '{E60F17EE-1FE2-4C07-9307-7D04EAA68C76}';
  IID_IAudioTag: TGUID = '{FEE7660B-D849-4106-A401-63CFB8735FBB}';
  IID_ISwdAudioDecoder: TGUID = '{F2E3A291-FEB6-4D01-BF4B-E69EC9DCAB96}';
  CLASS_SwdMP3Decoder: TGUID = '{3EC32CF4-BEC1-4030-81DC-034525429C14}';
  CLASS_AudioTag: TGUID = '{2F62CF6C-CC77-41DB-AE9A-679F9986D736}';
type

// *********************************************************************//
// Forward declaration of types defined in TypeLibrary                    
// *********************************************************************//
  _ISwdAudioEncoderEvents = dispinterface;
  _ISwdAudioDecoderEvents = dispinterface;
  ISwdAudioEncoder = interface;
  ISwdAudioEncoderDisp = dispinterface;
  IAudioTag = interface;
  IAudioTagDisp = dispinterface;
  ISwdAudioDecoder = interface;
  ISwdAudioDecoderDisp = dispinterface;

// *********************************************************************//
// Declaration of CoClasses defined in Type Library                       
// (NOTE: Here we map each CoClass to its Default Interface)              
// *********************************************************************//
  SwdMP3Encoder = ISwdAudioEncoder;
  SwdMP3Decoder = ISwdAudioDecoder;
  AudioTag = IAudioTag;


// *********************************************************************//
// Declaration of structures, unions and aliases.                         
// *********************************************************************//
  PUserType1 = ^WAVEFORMATEX; {*}
  PByte1 = ^Byte; {*}

  tWAVEFORMATEX = packed record
    wFormatTag: Word;
    nChannels: Word;
    nSamplesPerSec: LongWord;
    nAvgBytesPerSec: LongWord;
    nBlockAlign: Word;
    wBitsPerSample: Word;
    cbSize: Word;
  end;

  WAVEFORMATEX = tWAVEFORMATEX; 

// *********************************************************************//
// DispIntf:  _ISwdAudioEncoderEvents
// Flags:     (4096) Dispatchable
// GUID:      {F9ADEB91-8A20-4CB8-9D2D-3B7E96A09E8A}
// *********************************************************************//
  _ISwdAudioEncoderEvents = dispinterface
    ['{F9ADEB91-8A20-4CB8-9D2D-3B7E96A09E8A}']
    procedure OnError(const ErrorMessage: WideString); dispid 1;
    procedure OnWaveEncodeStart; dispid 2;
    procedure OnWaveEncodeProgress(PercentDone: Integer); dispid 3;
    procedure OnWaveEncodeFinish; dispid 4;
  end;

// *********************************************************************//
// DispIntf:  _ISwdAudioDecoderEvents
// Flags:     (4096) Dispatchable
// GUID:      {AD6A6218-BEF3-481E-8E4A-95B312CCF418}
// *********************************************************************//
  _ISwdAudioDecoderEvents = dispinterface
    ['{AD6A6218-BEF3-481E-8E4A-95B312CCF418}']
    procedure OnError(const ErrorMessage: WideString); dispid 1;
  end;

// *********************************************************************//
// Interface: ISwdAudioEncoder
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {9A45E4AB-68A2-4C85-B317-394351CC3330}
// *********************************************************************//
  ISwdAudioEncoder = interface(IDispatch)
    ['{9A45E4AB-68A2-4C85-B317-394351CC3330}']
    procedure Open(const szFilename: WideString); safecall;
    procedure OpenHandle(var hSource: Pointer); safecall;
    procedure SetInputFormat(nChannels: Integer; nSamplesPerSec: Integer; nBitsPerSample: Integer); safecall;
    procedure SetInputFormatWFX(var pWfx: WAVEFORMATEX); safecall;
    procedure SetOutputFormatVBR(fQuality: Single); safecall;
    procedure SetOutputFormatABR(nAvgBitrate: Integer); safecall;
    procedure SetOutputFormatCBR(nBitrate: Integer); safecall;
    procedure Write(var data: Byte; count: Integer); safecall;
    procedure Close; safecall;
    function Get_Filename: WideString; safecall;
    function Get_AddExtension: WordBool; safecall;
    procedure Set_AddExtension(pVal: WordBool); safecall;
    function Get_TagData: IAudioTag; safecall;
    function Get_OverwritePrompt: WordBool; safecall;
    procedure Set_OverwritePrompt(pVal: WordBool); safecall;
    function Get_LicenseNumber: WideString; safecall;
    procedure Set_LicenseNumber(const pVal: WideString); safecall;
    procedure EncodeFromWave(const waveFile: WideString); safecall;
    procedure EncodeFromWaveSync(const waveFile: WideString); safecall;
    property Filename: WideString read Get_Filename;
    property AddExtension: WordBool read Get_AddExtension write Set_AddExtension;
    property TagData: IAudioTag read Get_TagData;
    property OverwritePrompt: WordBool read Get_OverwritePrompt write Set_OverwritePrompt;
    property LicenseNumber: WideString read Get_LicenseNumber write Set_LicenseNumber;
  end;

// *********************************************************************//
// DispIntf:  ISwdAudioEncoderDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {9A45E4AB-68A2-4C85-B317-394351CC3330}
// *********************************************************************//
  ISwdAudioEncoderDisp = dispinterface
    ['{9A45E4AB-68A2-4C85-B317-394351CC3330}']
    procedure Open(const szFilename: WideString); dispid 1;
    procedure OpenHandle(var hSource: {??Pointer}OleVariant); dispid 2;
    procedure SetInputFormat(nChannels: Integer; nSamplesPerSec: Integer; nBitsPerSample: Integer); dispid 3;
    procedure SetInputFormatWFX(var pWfx: {??WAVEFORMATEX}OleVariant); dispid 4;
    procedure SetOutputFormatVBR(fQuality: Single); dispid 5;
    procedure SetOutputFormatABR(nAvgBitrate: Integer); dispid 6;
    procedure SetOutputFormatCBR(nBitrate: Integer); dispid 7;
    procedure Write(var data: Byte; count: Integer); dispid 8;
    procedure Close; dispid 9;
    property Filename: WideString readonly dispid 10;
    property AddExtension: WordBool dispid 11;
    property TagData: IAudioTag readonly dispid 12;
    property OverwritePrompt: WordBool dispid 13;
    property LicenseNumber: WideString dispid 14;
    procedure EncodeFromWave(const waveFile: WideString); dispid 15;
    procedure EncodeFromWaveSync(const waveFile: WideString); dispid 16;
  end;

// *********************************************************************//
// Interface: IAudioTag
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {FEE7660B-D849-4106-A401-63CFB8735FBB}
// *********************************************************************//
  IAudioTag = interface(IDispatch)
    ['{FEE7660B-D849-4106-A401-63CFB8735FBB}']
    function Get_Title: WideString; safecall;
    procedure Set_Title(const pVal: WideString); safecall;
    function Get_Artist: WideString; safecall;
    procedure Set_Artist(const pVal: WideString); safecall;
    function Get_Album: WideString; safecall;
    procedure Set_Album(const pVal: WideString); safecall;
    function Get_Year: WideString; safecall;
    procedure Set_Year(const pVal: WideString); safecall;
    function Get_Genre: WideString; safecall;
    procedure Set_Genre(const pVal: WideString); safecall;
    function Get_TrackNo: WideString; safecall;
    procedure Set_TrackNo(const pVal: WideString); safecall;
    function Get_Comment: WideString; safecall;
    procedure Set_Comment(const pVal: WideString); safecall;
    property Title: WideString read Get_Title write Set_Title;
    property Artist: WideString read Get_Artist write Set_Artist;
    property Album: WideString read Get_Album write Set_Album;
    property Year: WideString read Get_Year write Set_Year;
    property Genre: WideString read Get_Genre write Set_Genre;
    property TrackNo: WideString read Get_TrackNo write Set_TrackNo;
    property Comment: WideString read Get_Comment write Set_Comment;
  end;

// *********************************************************************//
// DispIntf:  IAudioTagDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {FEE7660B-D849-4106-A401-63CFB8735FBB}
// *********************************************************************//
  IAudioTagDisp = dispinterface
    ['{FEE7660B-D849-4106-A401-63CFB8735FBB}']
    property Title: WideString dispid 1;
    property Artist: WideString dispid 2;
    property Album: WideString dispid 3;
    property Year: WideString dispid 4;
    property Genre: WideString dispid 5;
    property TrackNo: WideString dispid 6;
    property Comment: WideString dispid 7;
  end;

// *********************************************************************//
// Interface: ISwdAudioDecoder
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {F2E3A291-FEB6-4D01-BF4B-E69EC9DCAB96}
// *********************************************************************//
  ISwdAudioDecoder = interface(IDispatch)
    ['{F2E3A291-FEB6-4D01-BF4B-E69EC9DCAB96}']
    procedure Open(const szFilename: WideString); safecall;
    procedure OpenHandle(var hSource: Pointer); safecall;
    function Get_Position: Integer; safecall;
    procedure Set_Position(pVal: Integer); safecall;
    function Get_Format: PUserType1; safecall;
    function Get_Bitrate: Integer; safecall;
    function Get_Duration: Integer; safecall;
    function Get_TagData: IAudioTag; safecall;
    function Get_LicenseNumber: WideString; safecall;
    procedure Set_LicenseNumber(const pVal: WideString); safecall;
    function Get_Filename: WideString; safecall;
    procedure Seek(Position: Integer); safecall;
    procedure Read(var data: Byte; count: Integer; out Read: Integer); safecall;
    procedure Close; safecall;
    procedure SetOutputFormat(nChannels: Integer; nSamplesPerSec: Integer; nBitsPerSample: Integer); safecall;
    procedure SetOutputFormatWFX(var pWfx: WAVEFORMATEX); safecall;
    property Position: Integer read Get_Position write Set_Position;
    property Format: PUserType1 read Get_Format;
    property Bitrate: Integer read Get_Bitrate;
    property Duration: Integer read Get_Duration;
    property TagData: IAudioTag read Get_TagData;
    property LicenseNumber: WideString read Get_LicenseNumber write Set_LicenseNumber;
    property Filename: WideString read Get_Filename;
  end;

// *********************************************************************//
// DispIntf:  ISwdAudioDecoderDisp
// Flags:     (4544) Dual NonExtensible OleAutomation Dispatchable
// GUID:      {F2E3A291-FEB6-4D01-BF4B-E69EC9DCAB96}
// *********************************************************************//
  ISwdAudioDecoderDisp = dispinterface
    ['{F2E3A291-FEB6-4D01-BF4B-E69EC9DCAB96}']
    procedure Open(const szFilename: WideString); dispid 1;
    procedure OpenHandle(var hSource: {??Pointer}OleVariant); dispid 2;
    property Position: Integer dispid 3;
    property Format: {??PUserType1}OleVariant readonly dispid 4;
    property Bitrate: Integer readonly dispid 5;
    property Duration: Integer readonly dispid 6;
    property TagData: IAudioTag readonly dispid 7;
    property LicenseNumber: WideString dispid 8;
    property Filename: WideString readonly dispid 9;
    procedure Seek(Position: Integer); dispid 10;
    procedure Read(var data: Byte; count: Integer; out Read: Integer); dispid 11;
    procedure Close; dispid 12;
    procedure SetOutputFormat(nChannels: Integer; nSamplesPerSec: Integer; nBitsPerSample: Integer); dispid 13;
    procedure SetOutputFormatWFX(var pWfx: {??WAVEFORMATEX}OleVariant); dispid 14;
  end;


// *********************************************************************//
// OLE Control Proxy class declaration
// Control Name     : TSwdMP3Encoder
// Help String      : This control is intented to make your application capable of creating MP3 files.
// Default Interface: ISwdAudioEncoder
// Def. Intf. DISP? : No
// Event   Interface: _ISwdAudioEncoderEvents
// TypeFlags        : (2) CanCreate
// *********************************************************************//
  TSwdMP3EncoderOnError = procedure(ASender: TObject; const ErrorMessage: WideString) of object;
  TSwdMP3EncoderOnWaveEncodeProgress = procedure(ASender: TObject; PercentDone: Integer) of object;

  TSwdMP3Encoder = class(TOleControl)
  private
    FOnError: TSwdMP3EncoderOnError;
    FOnWaveEncodeStart: TNotifyEvent;
    FOnWaveEncodeProgress: TSwdMP3EncoderOnWaveEncodeProgress;
    FOnWaveEncodeFinish: TNotifyEvent;
    FIntf: ISwdAudioEncoder;
    function  GetControlInterface: ISwdAudioEncoder;
  protected
    procedure CreateControl;
    procedure InitControlData; override;
    function Get_TagData: IAudioTag;
  public
    procedure Open(const szFilename: WideString);
    procedure OpenHandle(var hSource: Pointer);
    procedure SetInputFormat(nChannels: Integer; nSamplesPerSec: Integer; nBitsPerSample: Integer);
    procedure SetInputFormatWFX(var pWfx: WAVEFORMATEX);
    procedure SetOutputFormatVBR(fQuality: Single);
    procedure SetOutputFormatABR(nAvgBitrate: Integer);
    procedure SetOutputFormatCBR(nBitrate: Integer);
    procedure Write(var data: Byte; count: Integer);
    procedure Close;
    procedure EncodeFromWave(const waveFile: WideString);
    procedure EncodeFromWaveSync(const waveFile: WideString);
    property  ControlInterface: ISwdAudioEncoder read GetControlInterface;
    property  DefaultInterface: ISwdAudioEncoder read GetControlInterface;
    property Filename: WideString index 10 read GetWideStringProp;
    property TagData: IAudioTag read Get_TagData;
  published
    property Anchors;
    property AddExtension: WordBool index 11 read GetWordBoolProp write SetWordBoolProp stored False;
    property OverwritePrompt: WordBool index 13 read GetWordBoolProp write SetWordBoolProp stored False;
    property LicenseNumber: WideString index 14 read GetWideStringProp write SetWideStringProp stored False;
    property OnError: TSwdMP3EncoderOnError read FOnError write FOnError;
    property OnWaveEncodeStart: TNotifyEvent read FOnWaveEncodeStart write FOnWaveEncodeStart;
    property OnWaveEncodeProgress: TSwdMP3EncoderOnWaveEncodeProgress read FOnWaveEncodeProgress write FOnWaveEncodeProgress;
    property OnWaveEncodeFinish: TNotifyEvent read FOnWaveEncodeFinish write FOnWaveEncodeFinish;
  end;


// *********************************************************************//
// OLE Control Proxy class declaration
// Control Name     : TSwdMP3Decoder
// Help String      : This control is intended to make your application capable of reading MP3 files.
// Default Interface: ISwdAudioDecoder
// Def. Intf. DISP? : No
// Event   Interface: _ISwdAudioDecoderEvents
// TypeFlags        : (2) CanCreate
// *********************************************************************//
  TSwdMP3DecoderOnError = procedure(ASender: TObject; const ErrorMessage: WideString) of object;

  TSwdMP3Decoder = class(TOleControl)
  private
    FOnError: TSwdMP3DecoderOnError;
    FIntf: ISwdAudioDecoder;
    function  GetControlInterface: ISwdAudioDecoder;
  protected
    procedure CreateControl;
    procedure InitControlData; override;
    function Get_Format: PUserType1;
    function Get_TagData: IAudioTag;
  public
    procedure Open(const szFilename: WideString);
    procedure OpenHandle(var hSource: Pointer);
    procedure Seek(Position: Integer);
    procedure Read(var data: Byte; count: Integer; out Read: Integer);
    procedure Close;
    procedure SetOutputFormat(nChannels: Integer; nSamplesPerSec: Integer; nBitsPerSample: Integer);
    procedure SetOutputFormatWFX(var pWfx: WAVEFORMATEX);
    property  ControlInterface: ISwdAudioDecoder read GetControlInterface;
    property  DefaultInterface: ISwdAudioDecoder read GetControlInterface;
    property Format: PUserType1 read Get_Format;
    property Bitrate: Integer index 5 read GetIntegerProp;
    property Duration: Integer index 6 read GetIntegerProp;
    property TagData: IAudioTag read Get_TagData;
    property Filename: WideString index 9 read GetWideStringProp;
  published
    property Anchors;
    property Position: Integer index 3 read GetIntegerProp write SetIntegerProp stored False;
    property LicenseNumber: WideString index 8 read GetWideStringProp write SetWideStringProp stored False;
    property OnError: TSwdMP3DecoderOnError read FOnError write FOnError;
  end;

// *********************************************************************//
// The Class CoAudioTag provides a Create and CreateRemote method to          
// create instances of the default interface IAudioTag exposed by              
// the CoClass AudioTag. The functions are intended to be used by             
// clients wishing to automate the CoClass objects exposed by the         
// server of this typelibrary.                                            
// *********************************************************************//
  CoAudioTag = class
    class function Create: IAudioTag;
    class function CreateRemote(const MachineName: string): IAudioTag;
  end;

procedure Register;

resourcestring
  dtlServerPage = 'ActiveX';

  dtlOcxPage = 'ActiveX';

implementation

uses ComObj;

procedure TSwdMP3Encoder.InitControlData;
const
  CEventDispIDs: array [0..3] of DWORD = (
    $00000001, $00000002, $00000003, $00000004);
  CControlData: TControlData2 = (
    ClassID: '{E60F17EE-1FE2-4C07-9307-7D04EAA68C76}';
    EventIID: '{F9ADEB91-8A20-4CB8-9D2D-3B7E96A09E8A}';
    EventCount: 4;
    EventDispIDs: @CEventDispIDs;
    LicenseKey: nil (*HR:$80004002*);
    Flags: $00000000;
    Version: 401);
begin
  ControlData := @CControlData;
  TControlData2(CControlData).FirstEventOfs := Cardinal(@@FOnError) - Cardinal(Self);
end;

procedure TSwdMP3Encoder.CreateControl;

  procedure DoCreate;
  begin
    FIntf := IUnknown(OleObject) as ISwdAudioEncoder;
  end;

begin
  if FIntf = nil then DoCreate;
end;

function TSwdMP3Encoder.GetControlInterface: ISwdAudioEncoder;
begin
  CreateControl;
  Result := FIntf;
end;

function TSwdMP3Encoder.Get_TagData: IAudioTag;
begin
    Result := DefaultInterface.TagData;
end;

procedure TSwdMP3Encoder.Open(const szFilename: WideString);
begin
  DefaultInterface.Open(szFilename);
end;

procedure TSwdMP3Encoder.OpenHandle(var hSource: Pointer);
begin
  DefaultInterface.OpenHandle(hSource);
end;

procedure TSwdMP3Encoder.SetInputFormat(nChannels: Integer; nSamplesPerSec: Integer; 
                                        nBitsPerSample: Integer);
begin
  DefaultInterface.SetInputFormat(nChannels, nSamplesPerSec, nBitsPerSample);
end;

procedure TSwdMP3Encoder.SetInputFormatWFX(var pWfx: WAVEFORMATEX);
begin
  DefaultInterface.SetInputFormatWFX(pWfx);
end;

procedure TSwdMP3Encoder.SetOutputFormatVBR(fQuality: Single);
begin
  DefaultInterface.SetOutputFormatVBR(fQuality);
end;

procedure TSwdMP3Encoder.SetOutputFormatABR(nAvgBitrate: Integer);
begin
  DefaultInterface.SetOutputFormatABR(nAvgBitrate);
end;

procedure TSwdMP3Encoder.SetOutputFormatCBR(nBitrate: Integer);
begin
  DefaultInterface.SetOutputFormatCBR(nBitrate);
end;

procedure TSwdMP3Encoder.Write(var data: Byte; count: Integer);
begin
  DefaultInterface.Write(data, count);
end;

procedure TSwdMP3Encoder.Close;
begin
  DefaultInterface.Close;
end;

procedure TSwdMP3Encoder.EncodeFromWave(const waveFile: WideString);
begin
  DefaultInterface.EncodeFromWave(waveFile);
end;

procedure TSwdMP3Encoder.EncodeFromWaveSync(const waveFile: WideString);
begin
  DefaultInterface.EncodeFromWaveSync(waveFile);
end;

procedure TSwdMP3Decoder.InitControlData;
const
  CEventDispIDs: array [0..0] of DWORD = (
    $00000001);
  CControlData: TControlData2 = (
    ClassID: '{3EC32CF4-BEC1-4030-81DC-034525429C14}';
    EventIID: '{AD6A6218-BEF3-481E-8E4A-95B312CCF418}';
    EventCount: 1;
    EventDispIDs: @CEventDispIDs;
    LicenseKey: nil (*HR:$80004002*);
    Flags: $00000000;
    Version: 401);
begin
  ControlData := @CControlData;
  TControlData2(CControlData).FirstEventOfs := Cardinal(@@FOnError) - Cardinal(Self);
end;

procedure TSwdMP3Decoder.CreateControl;

  procedure DoCreate;
  begin
    FIntf := IUnknown(OleObject) as ISwdAudioDecoder;
  end;

begin
  if FIntf = nil then DoCreate;
end;

function TSwdMP3Decoder.GetControlInterface: ISwdAudioDecoder;
begin
  CreateControl;
  Result := FIntf;
end;

function TSwdMP3Decoder.Get_Format: PUserType1;
begin
    Result := DefaultInterface.Format;
end;

function TSwdMP3Decoder.Get_TagData: IAudioTag;
begin
    Result := DefaultInterface.TagData;
end;

procedure TSwdMP3Decoder.Open(const szFilename: WideString);
begin
  DefaultInterface.Open(szFilename);
end;

procedure TSwdMP3Decoder.OpenHandle(var hSource: Pointer);
begin
  DefaultInterface.OpenHandle(hSource);
end;

procedure TSwdMP3Decoder.Seek(Position: Integer);
begin
  DefaultInterface.Seek(Position);
end;

procedure TSwdMP3Decoder.Read(var data: Byte; count: Integer; out Read: Integer);
begin
  DefaultInterface.Read(data, count, Read);
end;

procedure TSwdMP3Decoder.Close;
begin
  DefaultInterface.Close;
end;

procedure TSwdMP3Decoder.SetOutputFormat(nChannels: Integer; nSamplesPerSec: Integer; 
                                         nBitsPerSample: Integer);
begin
  DefaultInterface.SetOutputFormat(nChannels, nSamplesPerSec, nBitsPerSample);
end;

procedure TSwdMP3Decoder.SetOutputFormatWFX(var pWfx: WAVEFORMATEX);
begin
  DefaultInterface.SetOutputFormatWFX(pWfx);
end;

class function CoAudioTag.Create: IAudioTag;
begin
  Result := CreateComObject(CLASS_AudioTag) as IAudioTag;
end;

class function CoAudioTag.CreateRemote(const MachineName: string): IAudioTag;
begin
  Result := CreateRemoteComObject(MachineName, CLASS_AudioTag) as IAudioTag;
end;

procedure Register;
begin
  RegisterComponents(dtlOcxPage, [TSwdMP3Encoder, TSwdMP3Decoder]);
end;

end.
