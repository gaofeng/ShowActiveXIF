unit untComTypeLibrary;

////////////////////////////////////////////////////////////////////////////////
//
//   TTypeLibrary
//      |
//      |_ TCoClass (TTypeClass)
//      |  |_ TCoInterface (Interfaces)
//      |  |_ TCoMember (Properties and Functions)
//      |     |_ TComDataType (Params and result Value)
//      |
//      |_ TInterface (TTypeClass)
//      |  |_ TMember (Properties and Functions)
//      |     |_ TComDataType (Params and result Value)
//      |
//      |_ TEnum (TTypeClass)
//      |  |_ TComDataType (Values)
//      |
//      |_ TRecord (TTypeClass)
//      |  |_ TComDataType (Fields)
//      |
//      |_ TAlias (TTypeClass)
//      |  |_ TComDataType (AliasType)
//      |
//      |_ TModules (TTypeClass)
//         |_ TMember (Functions)
//            |_ TComDataType (Params and result Value)
//
////////////////////////////////////////////////////////////////////////////////
interface

uses
  Windows, SysUtils, Classes, Contnrs, ActiveX, ComObj;

// Data type names
const
  DATA_TYPE_NAMES:  Array [0..31] of String =
                    ('',        '',         'SmallInt', 'Integer',  'Single',   'Double',   'Currency', 'Date',
                     'String',  'IDispatch','Integer',  'Boolean',  'Variant',  'IUnknown', 'Decimal',   '',
                     'Char',    'Byte',     'Word',     'LongWord', 'Int64',    'Int64',    'Integer',  'LongWord',
                     'Void',    'HResult',  'Pointer',  'Array',    'Array',    'Type',     'PChar',    'PWideChar');

// Type utils exception type
type
  ETypeUtilException=  class(Exception);                     

// Array bound type
type
  PArrayBound       =  ^TArrayBound;
  TArrayBound       =  packed record
     lBound:        Integer;
     uBound:        Integer;
  end;

// Com data type class
type
  TComDataType      =  class(TObject)
  private
     // Private declarations
     FVT:           Integer;
     FIsOptional:   Boolean;
     FConstValue:   Integer;
     FName:         String;
     FIsUserDefined:Boolean;
     FIsArray:      Boolean;
     FBounds:       TList;
     FGuid:         TGUID;
  protected
     // Protected declarations
     function       GetDataTypeName: String;
     function       GetBoundsCount: Integer;
     function       GetBounds(Index: Integer): TArrayBound;
     procedure      SetName(Value: String);
     procedure      SetVT(Value: Integer);
     procedure      SetIsUserDefined(Value: Boolean);
     procedure      SetIsArray(Value: Boolean);
     procedure      AddBound(LowBound, HiBound: Integer);
  public
     // Public declarations
     constructor    Create;
     destructor     Destroy; override;
     property       Name: String read FName;
     property       DataType: Integer read FVT;
     property       DataTypeName: String read GetDataTypeName;
     property       IsArray: Boolean read FIsArray;
     property       IsUserDefined: Boolean read FIsUserDefined;
     property       IsOptional: Boolean read FIsOptional;
     property       BoundsCount: Integer read GetBoundsCount;
     property       Bounds[Index: Integer]: TArrayBound read GetBounds;
     property       UserDefinedRef: TGUID read FGuid;
     property       ConstValue: Integer read FConstValue;
  end;

// Base class for all type information based objects
type
  TTypeClass        =  class(TObject)
  private
     // Private declarations
     FLoaded:       Boolean;
     FGuid:         TGUID;
     FName:         String;
     FTypeInfo:     ITypeInfo;
  protected
     // Protected declarations
     procedure      LoadBaseInfo;
  public
     // Public declarations
     constructor    Create(TypeInfo: ITypeInfo);
     destructor     Destroy; override;
     procedure      Load; virtual;
     property       Loaded: Boolean read FLoaded;
     property       Guid: TGUID read FGuid;
     property       Name: String read FName;
  end;

// Alias type info object
type
  TAlias            =  class(TTypeClass)
  private
     // Private declarations
     FAliasType:   TComDataType;
  protected
     // Protected declarations
  public
     // Public declarations
     constructor    Create(TypeInfo: ITypeInfo);
     destructor     Destroy; override;
     procedure      Load; override;
     property       AliasType: TComDataType read FAliasType;
  end;

// Record type info object
type
  TRecord           =  class(TTypeClass)
  private
     // Private declarations
     FFields:       TObjectList;
  protected
     // Protected declarations
     function       GetFields(Index: Integer): TComDataType;
     function       GetFieldCount: Integer;
  public
     // Public declarations
     constructor    Create(TypeInfo: ITypeInfo);
     destructor     Destroy; override;
     procedure      Load; override;
     property       Fields[Index: Integer]: TComDataType read GetFields;
     property       FieldCount: Integer read GetFieldCount;
  end;

// Enumeration type info object
type
  TEnum             =  class(TTypeClass)
  private
     // Private declarations
     FValues:       TObjectList;
  protected
     // Protected declarations
     function       GetValues(Index: Integer): TComDataType;
     function       GetValueCount: Integer;
  public
     // Public declarations
     constructor    Create(TypeInfo: ITypeInfo);
     destructor     Destroy; override;
     procedure      Load; override;
     property       Values[Index: Integer]: TComDataType read GetValues;
     property       ValueCount: Integer read GetValueCount;
  end;

// Member type class for interface functions and properties
type
  TMember           =  class(TObject)
  private
     // Private declarations
     FID:           Integer;
     FIsDispatch:   Boolean;
     FIsHidden:     Boolean;
     FName:         String;
     FParams:       TObjectList;
     FValue:        TComDataType;
     FCanRead:      Boolean;
     FCanWrite:     Boolean;
  protected
     // Protected declarations
     procedure      Load(TypeInfo: ITypeInfo; VarDesc: PVarDesc; Index: Integer); overload;
     procedure      Load(TypeInfo: ITypeInfo; FuncDesc: PFuncDesc; Index: Integer); overload;
     function       GetParam(Index: Integer): TComDataType;
     function       GetParamCount: Integer;
  public
     // Public declarations
     constructor    Create(TypeInfo: ITypeInfo; FuncDesc: PFuncDesc; Index: Integer); overload;
     constructor    Create(TypeInfo: ITypeInfo; VarDesc: PVarDesc; Index: Integer); overload;
     destructor     Destroy; override;
     property       ID: Integer read FID;
     property       IsDispatch: Boolean read FIsDispatch;
     property       IsHidden: Boolean read FIsHidden;
     property       Name: String read FName;
     property       Params[Index: Integer]: TComDataType read GetParam;
     property       ParamCount: Integer read GetParamCount;
     property       CanRead: Boolean read FCanRead;
     property       CanWrite: Boolean read FCanWrite;
     property       Value: TComDataType read FValue;
  end;

// Interface type info object
type
  TInterface        =  class(TTypeClass)
  private
     // Private declarations
     FProperties:   TObjectList;
     FFunctions:    TObjectList;
  protected
     // Protected declarations
     procedure      LoadVariables;
     procedure      LoadMembers;
     function       GetProperty(Index: Integer): TMember;
     function       GetFunction(Index: Integer): TMember;
     function       GetPropertyCount: Integer;
     function       GetFunctionCount: Integer;
  public
     // Public declarations
     constructor    Create(TypeInfo: ITypeInfo);
     destructor     Destroy; override;
     procedure      Load; override;
     property       Properties[Index: Integer]: TMember read GetProperty;
     property       PropertyCount: Integer read GetPropertyCount;
     property       Functions[Index: Integer]: TMember read GetFunction;
     property       FunctionCount: Integer read GetFunctionCount;
  end;

// CoClass member type class
  TCoMember = class(TObject)
  private
    fFuncName:    String;
    fFuncSection: Word;
    fFuncOffset:  Cardinal;
    foVft:        Word;
    fFuncKind:    String;
    fInvokeKind:  String;
    function      SetFuncName( rtinfo: ITypeInfo; memid: TMemberID ):HResult;
    procedure     SetFuncKindName(fk: TFuncKind);
    procedure     SetFuncOffset(address:Cardinal);
    procedure     SetFuncInvokeKind(invkind: TInvokeKind);
  protected
    //procedure SetFuncOffset(address:Cardinal);
  public
    //function SetFuncName( rtinfo: ITypeInfo; memid: TMemberID ):HResult;
    constructor    Create;
    destructor     Destroy; override;
    property       Name :String Read fFuncName ;
    property       Section:Word read fFuncSection ;
    property       Offset:Cardinal Read fFuncOffset ;
    property       oVft:Word Read  foVft ;
    property       FuncKind:string read fFuncKind ;
    property       InvokeKind:String read fInvokeKInd;
    //procedure SetFuncKindName(fk: TFuncKind);
  end;

// CoClass Interface type class
type
  TCoInterface           =  class(TObject)
  private
     // Private declarations
     FGuid:         TGUID;
     FName:         String;
     FIsDispatch:   Boolean;
     FIsDefault:    Boolean;
     FIsSource:     Boolean;
     FCanCreate:    Boolean;
  protected
     // Protected declarations
  public
     // Public declarations
     constructor    Create;
     destructor     Destroy; override;
     property       Guid: TGUID read FGuid;
     property       IsDispatch: Boolean read FIsDispatch;
     property       IsDefault: Boolean read FIsDefault;
     property       IsSource: Boolean read FIsSource;
     property       CanCreate: Boolean read FCanCreate;
     property       Name: String read FName;
  end;

// CoClass type info object
type
  TCoClass          =  class(TTypeClass)
  private
     // Private declarations
     FInterfaces:   TObjectList;
     FFunctions:    TObjectList;
  protected
     // Protected declarations
     function       GetInterface(Index: Integer): TCoInterface;
     function       GetInterfaceCount: Integer;
     function       GetFunction(Index: Integer): TCoMember;
     function       GetFunctionCount: Integer;
  public
     // Public declarations
     constructor    Create(TypeInfo: ITypeInfo);
     destructor     Destroy; override;
     procedure      Load; override;
     property       Interfaces[Index: Integer]: TCoInterface read GetInterface;
     property       InterfaceCount: Integer read GetInterfaceCount;
     property       Functions[Index: Integer]: TCoMember read GetFunction;
     property       FunctionCount: Integer read GetFunctionCount;
  end;

// TModule type info object
type
  TModule           =  class(TTypeClass)
  private
     // Private declarations
     FFunctions:    TObjectList;
  protected
     // Protected declarations
     function       GetFunction(Index: Integer): TMember;
     function       GetFunctionCount: Integer;
  public
     // Public declarations
     constructor    Create(TypeInfo: ITypeInfo);
     destructor     Destroy; override;
     procedure      Load; override;
  end;

// Type library wrapper object
type
  TTypeLibrary      =  class(TObject)
  private
     // Private declarations
     FTypeLib:      ITypeLib;
     FGuid:         TGUID;
     FName:         String;
     FDescription:  String;
     FModules:      TObjectList;
     FCoClasses:    TObjectList;
     FInterfaces:   TObjectList;
     FRecords:      TObjectList;
     FAliases:      TObjectList;
     FEnums:        TObjectList;
  protected
     // Protected declarations
     function       GetModule(Index: Integer): TModule;
     function       GetCoClass(Index: Integer): TCoClass;
     function       GetInterface(Index: Integer): TInterface;
     function       GetRecord(Index: Integer): TRecord;
     function       GetEnum(Index: Integer): TEnum;
     function       GetAlias(Index: Integer): TAlias;
     function       GetCoClassCount: Integer;
     function       GetInterfaceCount: Integer;
     function       GetRecordCount: Integer;
     function       GetEnumCount: Integer;
     function       GetAliasCount: Integer;
     function       GetModuleCount: Integer;
     procedure      Load(TypeLibrary: String);
  public
     // Public declarations
     constructor    Create(TypeLibrary: String);
     destructor     Destroy; override;
     property       Name: String read FName;
     property       Description: String read FDescription;
     property       Guid: TGUID read FGuid;
     property       CoClasses[Index: Integer]: TCoClass read GetCoClass;
     property       CoClassCount: Integer read GetCoClassCount;
     property       Interfaces[Index: Integer]: TInterface read GetInterface;
     property       InterfaceCount: Integer read GetInterfaceCount;
     property       Records[Index: Integer]: TRecord read GetRecord;
     property       RecordCount: Integer read GetRecordCount;
     property       Enums[Index: Integer]: TEnum read GetEnum;
     property       EnumCount: Integer read GetEnumCount;
     property       Aliases[Index: Integer]: TAlias read GetAlias;
     property       AliasCount: Integer read GetAliasCount;
     property       Modules[Index: Integer]: TModule read GetModule;
     property       ModuleCount: Integer read GetModuleCount;
  end;


// Resource strings
resourcestring
  resTypeInfoNil    =  'TypeInfo must be a non-nil interface pointer';

// Utility functions
procedure  LoadDataType(TypeInfo: ITypeInfo; Description: TTypeDesc; DataType: TComDataType); overload;
procedure  LoadDataType(TypeInfo: ITypeInfo; Description: TElemDesc; DataType: TComDataType); overload;

////////////////////////////////////////////////////////////////////////////////
implementation
////////////////////////////////////////////////////////////////////////////////

// TModule
procedure TModule.Load;
var  ptAttr:     PTypeAttr;
     pfDesc:     PFuncDesc;
     dwCount:    Integer;
begin

  // Check loaded state
  if FLoaded then exit;

  // Get type info attributes
  if (FTypeInfo.GetTypeAttr(ptAttr) = S_OK) then
  begin
     // Add properties and methods
     for dwCount:=0 to Pred(ptAttr^.cFuncs) do
     begin
        // Get the function description
        if (FTypeInfo.GetFuncDesc(dwCount, pfDesc) = S_OK) then
        begin
           // The member will load itself
           FFunctions.Add(TMember.Create(FTypeInfo, pfDesc, dwCount));
           // Release the function description
           FTypeInfo.ReleaseFuncDesc(pfDesc);
        end;
     end;
     // Release the type attributes
     FTypeInfo.ReleaseTypeAttr(ptAttr);
  end;

  // Perform inherited (sets the loaded state)
  inherited Load;

end;

function TModule.GetFunction(Index: Integer): TMember;
begin

  // Return the object
  result:=TMember(FFunctions[Index]);

end;

function TModule.GetFunctionCount: Integer;
begin

  // Return the count
  result:=FFunctions.Count;

end;

constructor TModule.Create(TypeInfo: ITypeInfo);
begin

  // Perform inherited
  inherited Create(TypeInfo);

  // Set starting defaults
  FFunctions:=TObjectList.Create;
  FFunctions.OwnsObjects:=True;

end;

destructor TModule.Destroy;
begin

  // Free the function list
  FFunctions.Free;

  // Perform inherited
  inherited Destroy;

end;

// TCoClass
procedure TCoClass.Load;
var  ptAttr1:    PTypeAttr;
     ptAttr2:    PTypeAttr;
     ptInfo:     ITypeInfo;
     pwName:     PWideChar;
     cmbrInt:    TCoInterface;
     hrType:     HRefType;
     dwFlags:    Integer;
     dwCount,i:  Integer;
     aFunc:      TCoMember;
     HR:         HResult;

     //pv:IUnknown;
     pVTable:PCardinal;
     funcDesc:PFuncDesc;
begin

  // Check loaded state
  if FLoaded then exit;

  // Load the CoClass default and source interfaces
  if (FTypeInfo.GetTypeAttr(ptAttr1) = S_OK) then
  begin
     // Iterate the vars of the enumeration
     for dwCount:=0 to Pred(ptAttr1^.cImplTypes) do
     begin
        // Get the reference type via the index
        if (FTypeInfo.GetRefTypeOfImplType(dwCount, hrType) = S_OK) then
        begin
          // Get the type info so we can get the name
          if (FTypeInfo.GetRefTypeInfo(hrType, ptInfo) = S_OK) then
          begin
             // Get the attributes
             if (ptInfo.GetTypeAttr(ptAttr2) = S_OK) then
             begin
                //***** Create the coclass Interface
                cmbrInt:=TCoInterface.Create;
                cmbrInt.FGuid:=ptAttr2^.guid;
                // Get the name
                if (ptInfo.GetDocumentation(MEMBERID_NIL, @pwName, nil, nil, nil) = S_OK) then
                begin
                   // Name
                   if Assigned(pwName) then
                   begin
                      cmbrInt.FName:=OleStrToString(pwName);
                      SysFreeString(pwName);
                   end;
                end;
                // Get implemented type for this interface
                if (FTypeInfo.GetImplTypeFlags(dwCount, dwFlags) = S_OK) then
                begin
                  // Get the kind
                  cmbrInt.FIsDispatch:=(ptAttr2^.typekind = TKIND_DISPATCH);
                  // Get default and source flags
                  cmbrInt.FIsDefault:=((dwFlags and IMPLTYPEFLAG_FDEFAULT) > 0);
                  cmbrInt.FIsSource:=((dwFlags and IMPLTYPEFLAG_FSOURCE) > 0);
                  cmbrInt.FCanCreate:=((ptAttr2^.wTypeFlags and TYPEFLAG_FCANCREATE) > 0);
                  // Add to interface member list
                end;
                FInterfaces.Add(cmbrInt);

                //***** Create the coclass Function

                hr:=CoCreateInstance( ptAttr1.guid,
                                      nil,					// pUnkOuter
                                      CLSCTX_INPROC_SERVER ,   //or CLSCTX_INPROC_HANDLER
                                      ptAttr2.guid,
                                      pVTable );
                if ( (S_OK = hr) and Assigned(pVTable) )then
                begin
                  //============================================================================
                  // Enumerate through each member of an ITypeInfo.  Send the method name and
                  // address to the CoClassSymsAddSymbol function.
                  //=============================================================================
                  //EnumTypeInfoMembers( rtinfo, rpta, Pointer(pv) );
                  for i := 0 to Pred(ptAttr2.cFuncs) do
                  begin
                    ptInfo.GetFuncDesc( i, funcDesc );
                    aFunc:=TCoMember.Create ;
                    aFunc.SetFuncName ( ptInfo,funcDesc.memid  );
                    aFunc.FoVft := funcDesc.oVft;
                    aFunc.SetFuncKindName ( funcDesc.funckind );
                    aFunc.SetFuncInvokeKind(funcDesc.invkind );
                    // Index into the vtable to retrieve the method's virtual address
                    aFunc.SetFuncOffset(pVTable^ + Cardinal(funcDesc.oVft));

                    fFunctions.Add(aFunc);
                    ptInfo.ReleaseFuncDesc( funcDesc );
                  end;
                end;
                // Release the type info attributes
                ptInfo.ReleaseTypeAttr(ptAttr2);
             end;
             // Release the type info
             ptInfo:=nil;
          end;
        end;
     end;
     // Release the type attr
     FTypeInfo.ReleaseTypeAttr(ptAttr1);
  end;

  // Perform inherited (sets the loaded state)
  inherited Load;

end;

function TCoClass.GetInterface(Index: Integer): TCoInterface;
begin

  // Return the object
  result:=TCoInterface(FInterfaces[Index]);

end;

function TCoClass.GetInterfaceCount: Integer;
begin

  // Return the count
  result:=FInterfaces.Count;

end;

function TCoClass.GetFunction(Index: Integer): TCoMember;
begin
  Result := TCoMember(fFunctions[Index]);
end;

function TCoClass.GetFunctionCount: Integer;
begin
  Result := fFunctions.Count ;
end;

constructor TCoClass.Create(TypeInfo: ITypeInfo);
begin

  // Perform inherited
  inherited Create(TypeInfo);

  // Create member list
  FInterfaces:=TObjectList.Create;
  FInterfaces.OwnsObjects:=True;

  FFunctions := TObjectList.Create ;
  FFunctions.OwnsObjects := True;

end;

destructor TCoClass.Destroy;
begin

  // Free the member list
  FInterfaces.Free;
  FFunctions.Free ;

  // Perform inherited
  inherited Destroy;

end;

// TCoInterface
constructor TCoInterface.Create;
begin

  // Perform inherited
  inherited Create;

  // Set starting defaults
  FGuid:=GUID_NULL;
  FName:='';
  FIsDispatch:=False;
  FIsDefault:=False;
  FIsSource:=False;
  FCanCreate:=False;

end;

destructor TCoInterface.Destroy;
begin

  // Perform inherited
  inherited Destroy;

end;

// TMember
procedure TMember.Load(TypeInfo: ITypeInfo; VarDesc: PVarDesc; Index: Integer);
var  pwName:     PWideChar;
begin

  // Set the ID (for dispatch based items)
  FIsDispatch:=(VarDesc^.varkind = VAR_DISPATCH);
  if FIsDispatch then
     FID:=VarDesc^.memid
  else
     FID:=0;

  // Get the name
  if (TypeInfo.GetDocumentation(VarDesc^.memid, @pwName, nil, nil, nil) = S_OK) then
  begin
     // Name
     if Assigned(pwName) then
     begin
        FName:=OleStrToString(pwName);
        SysFreeString(pwName);
     end;
  end;

  // Is this read only, or read write
  FCanWrite:=((VarDesc^.wVarFlags and VARFLAG_FREADONLY) = 0);

  // Load the data type info
  LoadDataType(TypeInfo, VarDesc^.elemdescVar, FValue);

  // Determine if member is hidden
  FIsHidden:=((VarDesc^.wVarFlags and VARFLAG_FHIDDEN) > 0);

end;

procedure TMember.Load(TypeInfo: ITypeInfo; FuncDesc: PFuncDesc; Index: Integer);
var  pwName:     PWideChar;
     pwNames:    PBStrList;
     cdtParam:   TComDataType;
     dwNames:    Integer;
     dwParams:   Integer;
     dwCount:    Integer;
begin

  // Set the ID (for dispatch based items)
  FIsDispatch:=(FuncDesc^.funckind = FUNC_DISPATCH);
  if FIsDispatch then
     FID:=FuncDesc^.memid
  else
     FID:=0;

  // Get the name
  if (TypeInfo.GetDocumentation(FuncDesc^.memid, @pwName, nil, nil, nil) = S_OK) then
  begin
     // Name
     if Assigned(pwName) then
     begin
        FName:=OleStrToString(pwName);
        SysFreeString(pwName);
     end;
  end;

  // Is this read only, or read write
  case FuncDesc^.invkind  of
     INVOKE_FUNC           :
     begin
        // Doesnt make sense for functions
        FCanRead:=False;
        FCanWrite:=False;
     end;
     INVOKE_PROPERTYGET    : FCanWrite:=False;
     INVOKE_PROPERTYPUT    : FCanRead:=False;
     INVOKE_PROPERTYPUTREF : FCanRead:=False;
  end;

  // Load the params for the functions
  dwParams:=FuncDesc^.cParams;
  Inc(dwParams);
  if (dwParams > 1) and (FuncDesc^.invkind in [INVOKE_PROPERTYPUT, INVOKE_PROPERTYPUTREF]) then Dec(dwParams);

  // Allocate string array large enough for the names
  pwNames:=AllocMem(dwParams * SizeOf(POleStr));

  // Get the names
  dwNames:=0;
  if (TypeInfo.GetNames(FuncDesc^.memid, pwNames, Succ(FuncDesc^.cParams), dwNames) = S_OK) then
  begin
     // Make sure all name entries are allocated (even if we have to do it ourselves)
     for dwCount:=dwNames to Pred(dwParams) do pwNames[dwCount]:=StringToOleStr(Format('Param%d', [dwCount]));
     // Need to decrease the cParams by 2. One is to account for the fact that we
     // got the member name in the string array, the other is to account for the
     // fact that the list is zero based
     Dec(dwParams, 2);
     // Build up the parameter list
     for dwCount:=0 to dwParams do
     begin
        // Create the parameter
        cdtParam:=TComDataType.Create;
        // Set the name
        cdtParam.FName:=OleStrToString(pwNames[Succ(dwCount)]);
        // Load the data type
        LoadDataType(TypeInfo, FuncDesc^.lprgelemdescParam[dwCount], cdtParam);
        // Determine if optional
        cdtParam.FIsOptional:=((FuncDesc^.lprgelemdescParam[dwCount].paramdesc.wParamFlags and PARAMFLAG_FOPT) > 0);
        // Add to the parameter list
        FParams.Add(cdtParam);
     end;
  end;

  // Free the strings
  for dwCount:=0 to Succ(dwParams) do SysFreeString(pwNames[dwCount]);

  // Free the string array
  FreeMem(pwNames);

  // Set the return type for the function/property
  case FuncDesc^.invkind  of
     INVOKE_FUNC          :  LoadDataType(TypeInfo, FuncDesc^.elemdescFunc, FValue);
     INVOKE_PROPERTYGET   :
     begin
        if not(FIsDispatch) and (FuncDesc^.cParams > 0) then
           LoadDataType(TypeInfo, FuncDesc^.lprgelemdescParam[Pred(FuncDesc^.cParams)], FValue)
        else
           LoadDataType(TypeInfo, FuncDesc^.elemdescFunc, FValue);
     end;
     INVOKE_PROPERTYPUT,
     INVOKE_PROPERTYPUTREF:
     begin
        // cParams MUST be at least one
        LoadDataType(TypeInfo, FuncDesc^.lprgelemdescParam[Pred(FuncDesc^.cParams)], FValue);
     end
  end;

end;

function TMember.GetParam(Index: Integer): TComDataType;
begin

  // Return the object
  result:=TComDataType(FParams[Index]);

end;

function TMember.GetParamCount: Integer;
begin

  // Return the count
  result:=FParams.Count;

end;

constructor TMember.Create(TypeInfo: ITypeInfo; VarDesc: PVarDesc; Index: Integer);
begin

  // Perform inherited
  inherited Create;

  // Create the object list
  FID:=0;
  FName:='';
  FParams:=TObjectList.Create;
  FParams.OwnsObjects:=true;
  FValue:=TComDataType.Create;
  FCanRead:=True;
  FCanWrite:=True;
  FIsDispatch:=False;
  FIsHidden:=False;

  // Load the member information
  Load(TypeInfo, VarDesc, Index);

end;

constructor TMember.Create(TypeInfo: ITypeInfo; FuncDesc: PFuncDesc; Index: Integer);
begin

  // Perform inherited
  inherited Create;

  // Create the object list
  FID:=0;
  FName:='';
  FParams:=TObjectList.Create;
  FParams.OwnsObjects:=true;
  FValue:=TComDataType.Create;
  FCanRead:=True;
  FCanWrite:=True;
  FIsDispatch:=False;
  FIsHidden:=False;

  // Load the member information
  Load(TypeInfo, FuncDesc, Index);

end;

destructor TMember.Destroy;
begin

  // Free the parameter list
  FParams.Free;

  // Free the value type
  FValue.Free;

  // Perform inherited
  inherited Destroy;

end;

// TInterface
procedure TInterface.LoadMembers;
var  ptAttr:     PTypeAttr;
     pfDesc:     PFuncDesc;
     dwCount:    Integer;
begin

  // Get type info attributes
  if (FTypeInfo.GetTypeAttr(ptAttr) = S_OK) then
  begin
     // Add properties and methods
     for dwCount:=0 to Pred(ptAttr^.cFuncs) do
     begin
        // Get the function description
        if (FTypeInfo.GetFuncDesc(dwCount, pfDesc) = S_OK) then
        begin
           // Check the invokation kind
           if (pfDesc^.invkind = INVOKE_FUNC) then
              // The member will load itself
              FFunctions.Add(TMember.Create(FTypeInfo, pfDesc, dwCount))
           else
              // The member will load itself
              FProperties.Add(TMember.Create(FTypeInfo, pfDesc, dwCount));
           // Release the function description
           FTypeInfo.ReleaseFuncDesc(pfDesc);
        end;
     end;
     // Release the type attributes
     FTypeInfo.ReleaseTypeAttr(ptAttr);
  end;

end;

procedure TInterface.LoadVariables;
var  ptAttr:     PTypeAttr;
     pvDesc:     PVarDesc;
     dwCount:    Integer;
begin

  // Get type info attributes
  if (FTypeInfo.GetTypeAttr(ptAttr) = S_OK) then
  begin
     // Add variables (which are properties)
     for dwCount:=0 to Pred(ptAttr^.cVars) do
     begin
        // Get the var description
        if (FTypeInfo.GetVarDesc(dwCount, pvDesc) = S_OK) then
        begin
           // Create the property member
           FProperties.Add(TMember.Create(FTypeInfo, pvDesc, dwCount));
           // Release the var desc
           FTypeInfo.ReleaseVarDesc(pvDesc);
        end;
     end;
     // Release the type attributes
     FTypeInfo.ReleaseTypeAttr(ptAttr);
  end;

end;

procedure TInterface.Load;
begin

  // Check loaded state
  if FLoaded then exit;

  // Load the variables
  LoadVariables;

  // Load the members (functions/properties)
  LoadMembers;

  // Perform inherited (sets the loaded state)
  inherited Load;

end;

function TInterface.GetProperty(Index: Integer): TMember;
begin

  // Return object
  result:=TMember(FProperties[Index]);

end;

function TInterface.GetFunction(Index: Integer): TMember;
begin

  // Return the object
  result:=TMember(FFunctions[Index]);

end;

function TInterface.GetPropertyCount: Integer;
begin

  // Return the count
  result:=FProperties.Count;

end;

function TInterface.GetFunctionCount: Integer;
begin

  // Return the count
  result:=FFunctions.Count;

end;

constructor TInterface.Create(TypeInfo: ITypeInfo);
begin

  // Perform inherited
  inherited Create(TypeInfo);

  // Set starting defaults
  FProperties:=TObjectList.Create;
  FProperties.OwnsObjects:=True;
  FFunctions:=TObjectList.Create;
  FFunctions.OwnsObjects:=True;

end;

destructor TInterface.Destroy;
begin

  // Free the object lists
  FProperties.Free;
  FFunctions.Free;

  // Perform inherited
  inherited Destroy;

end;

// TEnum
procedure TEnum.Load;
var  ptAttr:     PTypeAttr;
     pvDesc:     PVarDesc;
     pwName:     PWideChar;
     cdtValue:   TComDataType;
     dwCount:    Integer;
begin

  // Bail if loaded
  if FLoaded then exit;

  // Load the constant values
  if (FTypeInfo.GetTypeAttr(ptAttr) = S_OK) then
  begin
     // Iterate the vars of the record
     for dwCount:=0 to Pred(ptAttr^.cVars) do
     begin
        // Get the var description
        if (FTypeInfo.GetVarDesc(dwCount, pvDesc) = S_OK) then
        begin
           // Create the value data object
           cdtValue:=TComDataType.Create;
           // Get documentation
           if (FTypeInfo.GetDocumentation(pvDesc^.memid, @pwName, nil, nil, nil) = S_OK) then
           begin
              // Name
              if Assigned(pwName) then
              begin
                 cdtValue.FName:=OleStrToString(pwName);
                 SysFreeString(pwName);
              end;
           end;
           // Constant value
           cdtValue.FVT:=pvDesc^.varkind;
           cdtValue.FConstValue:=pvDesc^.lpvarValue^;
           // Add to the value list
           FValues.Add(cdtValue);
           // Release the var desc
           FTypeInfo.ReleaseVarDesc(pvDesc);
        end;
     end;
     // Release the type attr
     FTypeInfo.ReleaseTypeAttr(ptAttr);
  end;

  // Perform inherited (sets the loaded state)
  inherited Load;

end;

function TEnum.GetValues(Index: Integer): TComDataType;
begin

  // Return the value field
  result:=TComDataType(FValues[Index]);

end;

function TEnum.GetValueCount: Integer;
begin

  // Return count
  result:=FValues.Count;

end;

constructor TEnum.Create(TypeInfo: ITypeInfo);
begin

  // Perform inherited
  inherited Create(TypeInfo);

  // Set starting defaults
  FValues:=TObjectList.Create;
  FValues.OwnsObjects:=True;
  
end;

destructor TEnum.Destroy;
begin

  // Free the value list
  FValues.Free;

  // Perform inherited
  inherited Destroy;

end;

// TRecord
procedure TRecord.Load;
var  ptAttr:     PTypeAttr;
     pvDesc:     PVarDesc;
     pwName:     PWideChar;
     cdtField:   TComDataType;
     dwCount:    Integer;
begin

  // Bail if loaded
  if FLoaded then exit;

  // Load the fields
  if (FTypeInfo.GetTypeAttr(ptAttr) = S_OK) then
  begin
     // Iterate the vars of the record
     for dwCount:=0 to Pred(ptAttr^.cVars) do
     begin
        // Get the var description
        if (FTypeInfo.GetVarDesc(dwCount, pvDesc) = S_OK) then
        begin
           // Create the field data object
           cdtField:=TComDataType.Create;
           // Get documentation
           if (FTypeInfo.GetDocumentation(pvDesc^.memid, @pwName, nil, nil, nil) = S_OK) then
           begin
              // Name
              if Assigned(pwName) then
              begin
                 cdtField.FName:=OleStrToString(pwName);
                 SysFreeString(pwName);
              end;
           end;
           // Load the data type info
           LoadDataType(FTypeInfo, pvDesc^.elemdescVar, cdtField);
           // Add to the field list
           FFields.Add(cdtField);
           // Release the var desc
           FTypeInfo.ReleaseVarDesc(pvDesc);
        end;
     end;
     // Release the type attr
     FTypeInfo.ReleaseTypeAttr(ptAttr);
  end;

  // Perform inherited (sets the loaded state)
  inherited Load;

end;

function TRecord.GetFields(Index: Integer): TComDataType;
begin

  // Return the data field
  result:=TComDataType(FFields[Index]);

end;

function TRecord.GetFieldCount: Integer;
begin

  // Return the count
  result:=FFields.Count;

end;

constructor TRecord.Create(TypeInfo: ITypeInfo);
begin

  // Perform inherited
  inherited Create(TypeInfo);

  // Set starting defaults
  FFields:=TObjectList.Create;
  FFields.OwnsObjects:=True;

end;

destructor TRecord.Destroy;
begin

  // Free the field list
  FFields.Free;

  // Perform inherited
  inherited Destroy;

end;

// TAlias
procedure TAlias.Load;
var  ptAttr:     PTypeAttr;
begin

  // Bail if already loaded
  if FLoaded then exit;

  // Get the type info attributes
  if (FTypeInfo.GetTypeAttr(ptAttr) = S_OK) then
  begin
     // Get the data type that this describes
     LoadDataType(FTypeInfo, ptAttr^.tdescAlias, FAliasType);
     // Release the type info attr
     FTypeInfo.ReleaseTypeAttr(ptAttr);
  end;

  // Perform inherited (sets the loaded state)
  inherited Load;

end;

constructor TAlias.Create(TypeInfo: ITypeInfo);
begin

  // Perform inherited
  inherited Create(TypeInfo);

  // Set starting defaults
  FAliasType:=TComDataType.Create;

end;

destructor TAlias.Destroy;
begin

  // Free the alias data type
  FAliasType.Free;

  // Perform inherited
  inherited Destroy;

end;

// TTypeClass
procedure TTypeClass.LoadBaseInfo;
var  ptAttr:     PTypeAttr;
     pwName:     PWideChar;
begin

  // Get the class name
  if (FTypeInfo.GetDocumentation(MEMBERID_NIL, @pwName, nil, nil, nil) = S_OK) then
  begin
     // Name
     if Assigned(pwName) then
     begin
        FName:=OleStrToString(pwName);
        SysFreeString(pwName);
     end;
  end;

  // Get the type info attributes
  if (FTypeInfo.GetTypeAttr(ptAttr) = S_OK) then
  begin
     // Set the guid
     FGuid:=ptAttr^.guid;
     // Release the pointer to the attributes
     FTypeInfo.ReleaseTypeAttr(ptAttr);
  end;

end;

procedure TTypeClass.Load;
begin

  // Set the loaded flag
  FLoaded:=True;

end;

constructor TTypeClass.Create(TypeInfo: ITypeInfo);
begin

  // Perform inherited
  inherited Create;

  // Check for nil being passed
  if Assigned(TypeInfo) then
     FTypeInfo:=TypeInfo
  else
     raise ETypeUtilException.CreateRes(@resTypeInfoNil);

  // Set starting defaults
  FLoaded:=False;
  FGuid:=GUID_NULL;
  FName:='';

  // Load the base type information
  LoadBaseInfo;

end;

destructor TTypeClass.Destroy;
begin

  // Release the type info
  FTypeInfo:=nil;

  // Perform inherited
  inherited Destroy;

end;

// TTypeLibrary
function TTypeLibrary.GetModule(Index: Integer): TModule;
begin

  // Return object
  result:=TModule(FModules[Index]);

  // Load the object
  result.Load;

end;

function TTypeLibrary.GetCoClass(Index: Integer): TCoClass;
begin

  // Return object
  result:=TCoClass(FCoClasses[Index]);

  // Load the object
  result.Load;

end;

function TTypeLibrary.GetInterface(Index: Integer): TInterface;
begin

  // Return object
  result:=TInterface(FInterfaces[Index]);

  // Load the object
  result.Load;

end;

function TTypeLibrary.GetRecord(Index: Integer): TRecord;
begin

  // Return object
  result:=TRecord(FRecords[Index]);

  // Load the object
  result.Load;

end;

function TTypeLibrary.GetEnum(Index: Integer): TEnum;
begin

  // Return object
  result:=TEnum(FEnums[Index]);

  // Load the object
  result.Load;

end;

function TTypeLibrary.GetAlias(Index: Integer): TAlias;
begin

  // Return object
  result:=TAlias(FAliases[Index]);

  // Load the object
  result.Load;

end;

function TTypeLibrary.GetCoClassCount: Integer;
begin

  // Return count
  result:=FCoClasses.Count;

end;

function TTypeLibrary.GetModuleCount: Integer;
begin

  // Return count
  result:=FModules.Count;

end;

function TTypeLibrary.GetInterfaceCount: Integer;
begin

  // Return count
  result:=FInterfaces.Count;

end;

function TTypeLibrary.GetRecordCount: Integer;
begin

  // Return count
  result:=FRecords.Count;

end;

function TTypeLibrary.GetEnumCount: Integer;
begin

  // Return count
  result:=FEnums.Count;

end;

function TTypeLibrary.GetAliasCount: Integer;
begin

  // Return count
  result:=FAliases.Count;

end;

procedure TTypeLibrary.Load(TypeLibrary: String);
var  pwFileName: PWideChar;
     pwName:     PWideChar;
     pwDesc:     PWideChar;
     ptlAttr:    PTLibAttr;
     ptAttr:     PTypeAttr;
     ptInfo:     ITypeInfo;
     dwCount:    Integer;
     hrStatus:   HResult;
begin

  // Attempt to load the type library
  pwFileName:=StringToOleStr(TypeLibrary);
  hrStatus:=LoadTypeLib(pwFileName, FTypeLib);
  SysFreeString(pwFileName);

  // Check result
  if (hrStatus <> S_OK) then raise Exception.Create(SysErrorMessage(hrStatus));

  // Get the type library name
  if (FTypeLib.GetDocumentation(-1, @pwName, @pwDesc, nil, nil) = S_OK) then
  begin
     // Name
     if Assigned(pwName) then
     begin
        FName:=OleStrToString(pwName);
        SysFreeString(pwName);
     end;
     // Description
     if Assigned(pwDesc) then
     begin
        FDescription:=OleStrToString(pwDesc);
        SysFreeString(pwDesc);
     end;
  end;

  // Get the guid
  if (FTypeLib.GetLibAttr(ptlAttr) = S_OK) then
  begin
     FGuid:=ptlAttr^.Guid;
     // Release the library attr
     FTypeLib.ReleaseTLibAttr(ptlAttr);
  end;

  // Get the type info count in the library
  for dwCount:=0 to Pred(FTypeLib.GetTypeInfoCount) do
  begin;
     // Get the type info at the given index
     if (FTypeLib.GetTypeInfo(dwCount, ptInfo) = S_OK) then
     begin
        // Get the type info attribute
        if (ptInfo.GetTypeAttr(ptAttr) = S_OK) then
        begin
           // Create the desired sub object
           case ptAttr.typekind of
              TKIND_ENUM     :  FEnums.Add(TEnum.Create(ptInfo));
              TKIND_RECORD   :  FRecords.Add(TRecord.Create(ptInfo));
              TKIND_MODULE   :  FModules.Add(TModule.Create(ptInfo));
              TKIND_INTERFACE,
              TKIND_DISPATCH :  FInterfaces.Add(TInterface.Create(ptInfo));
              TKIND_COCLASS  :  FCoClasses.Add(TCoClass.Create(ptInfo));
              TKIND_ALIAS    :  FAliases.Add(TAlias.Create(ptInfo));
           end;
           // Release the type attribute
           ptInfo.ReleaseTypeAttr(ptAttr);
        end;
        // Release the type info
        ptInfo:=nil;
     end;
  end;

end;

constructor TTypeLibrary.Create(TypeLibrary: String);
begin

  // Perform inherited
  inherited Create;

  // Set starting defaults
  FTypeLib:=nil;
  FGuid:=GUID_NULL;
  FName:='';
  FDescription:='';
  FCoClasses:=TObjectList.Create;
  FCoClasses.OwnsObjects:=True;
  FInterfaces:=TObjectList.Create;
  FInterfaces.OwnsObjects:=True;
  FRecords:=TObjectList.Create;
  FRecords.OwnsObjects:=True;
  FAliases:=TObjectList.Create;
  FAliases.OwnsObjects:=True;
  FEnums:=TObjectList.Create;
  FEnums.OwnsObjects:=True;
  FModules:=TObjectList.Create;
  FModules.OwnsObjects:=True;

  // Load the classes
  Load(TypeLibrary);

end;

destructor TTypeLibrary.Destroy;
begin

  // Clear interfaces
  FTypeLib:=nil;

  // Free object lists
  FModules.Free;
  FCoClasses.Free;
  FInterfaces.Free;
  FRecords.Free;
  FAliases.Free;
  FEnums.Free;

  // Perform inherited
  inherited Destroy;

end;

// TComDataType
function TComDataType.GetBoundsCount: Integer;
begin

  // Return count
  result:=FBounds.Count;

end;

function TComDataType.GetDataTypeName;
begin

  // Return the delphi data type name
  result:=DATA_TYPE_NAMES[FVT];

end;

function TComDataType.GetBounds(Index: Integer): TArrayBound;
begin

  // Return the requested bounds
  result:=PArrayBound(FBounds[Index])^;

end;

procedure TComDataType.SetName(Value: String);
begin

  // Callable by "friends"
  FName:=Value;

end;

procedure TComDataType.SetVT(Value: Integer);
begin

  // Callable by "friends"
  FVT:=Value;

end;

procedure TComDataType.SetIsUserDefined(Value: Boolean);
begin

  // Callable by "friends"
  FIsUserDefined:=True;

end;

procedure TComDataType.SetIsArray(Value: Boolean);
begin

  // Callable by "friends"
  FIsArray:=True;

end;

procedure TComDataType.AddBound(LowBound, HiBound: Integer);
var  paBound:    PArrayBound;
begin

  // Allocate memory
  paBound:=AllocMem(SizeOf(TArrayBound));

  // Set bounds
  paBound^.lBound:=LowBound;
  paBound^.uBound:=HiBound;

  // Add to the list
  FBounds.Add(paBound);

end;

constructor TComDataType.Create;
begin

  // Perform inherited
  inherited Create;

  // Set starting values
  FVT:=0;
  FName:='';
  FIsOptional:=False;
  FIsUserDefined:=False;
  FIsArray:=False;
  FConstValue:=0;
  FBounds:=TList.Create;
  FGuid:=GUID_NULL;

end;

destructor TComDataType.Destroy;
var  dwBounds:   Integer;
begin

  // Clear all array bounds
  for dwBounds:=0 to Pred(FBounds.Count) do
  begin
     FreeMem(PArrayBound(FBounds[dwBounds]));
  end;

  // Free the list
  FBounds.Free;

  // Perform inherited
  inherited Destroy;

end;

procedure LoadDataType(TypeInfo: ITypeInfo; Description: TElemDesc; DataType: TComDataType);
var  ptInfo:     ITypeInfo;
     ptAttr:     PTypeAttr;
     udType:     Cardinal;
     dwBounds:   Integer;
begin

  // Get the data type from the type desc
  DataType.FVT:=Description.tdesc.vt;

  // Need to figure out what the actual type is
  udType:=0;
  case DataType.FVT of
     VT_PTR,
     VT_SAFEARRAY   :
     begin
        // Get the data type from the pointer type
        DataType.FVT:=Description.tdesc.ptdesc.vt;
        // If user defined, set the udt reftype
        if (DataType.FVT = VT_USERDEFINED) then udType:=Description.tdesc.padesc.tdescElem.hreftype;
     end;
     // Get the type from the array description
     VT_CARRAY      :
     begin
        // Set array flag
        DataType.FIsArray:=True;
        // Add the bounds
        for dwBounds:=0 to Pred(Description.tdesc.padesc^.cDims) do
        begin
           with Description.tdesc.padesc^ do
           begin
              DataType.AddBound(rgbounds[dwBounds].lLbound, Pred(rgbounds[dwBounds].cElements));
           end;
        end;
        // Set the data type
        DataType.FVT:=Description.tdesc.padesc.tdescElem.vt;
     end;
     // Get the user defined reftype
     VT_USERDEFINED :  udType:=Description.tdesc.hreftype;
  end;

  // Check for user defined type
  if (DataType.FVT = VT_USERDEFINED) then
  begin
     // User defined
     DataType.FIsUserDefined:=True;
     // Try to get the referenced type info
     if (TypeInfo.GetRefTypeInfo(udType, ptInfo) = S_OK) then
     begin
        // Get the attributes so we can get the guid
        if (ptInfo.GetTypeAttr(ptAttr) = S_OK) then
        begin
           DataType.FGuid:=ptAttr^.guid;
           // Release the type attributes
           ptInfo.ReleaseTypeAttr(ptAttr);
        end;
        // Release the type info
        ptInfo:=nil;
     end;
  end;

end;

procedure LoadDataType(TypeInfo: ITypeInfo; Description: TTypeDesc; DataType: TComDataType);
var  ptInfo:     ITypeInfo;
     ptAttr:     PTypeAttr;
     udType:     Cardinal;
     dwBounds:   Integer;
begin

  // Get the data type from the type desc
  DataType.FVT:=Description.vt;

  // Need to figure out what the actual type is
  udType:=0;
  case DataType.FVT of
     VT_PTR,
     VT_SAFEARRAY   :
     begin
        // Get the data type from the pointer type
        DataType.FVT:=Description.ptdesc.vt;
        // If user defined, set the udt reftype
        if (DataType.FVT = VT_USERDEFINED) then udType:=Description.padesc.tdescElem.hreftype;
     end;
     // Get the type from the array description
     VT_CARRAY      :
     begin
        // Set array flag
        DataType.FIsArray:=True;
        // Add the bounds
        for dwBounds:=0 to Pred(Description.padesc^.cDims) do
        begin
           with Description.padesc^ do
           begin
              DataType.AddBound(rgbounds[dwBounds].lLbound, Pred(rgbounds[dwBounds].cElements));
           end;
        end;
        // Set the array data type
        DataType.FVT:=Description.padesc.tdescElem.vt;
     end;
     // Get the user defined reftype
     VT_USERDEFINED :  udType:=Description.hreftype;
  end;

  // Check for user defined type
  if (DataType.FVT = VT_USERDEFINED) then
  begin
     // User defined
     DataType.FIsUserDefined:=True;
     // Try to get the referenced type info
     if (TypeInfo.GetRefTypeInfo(udType, ptInfo) = S_OK) then
     begin
        // Get the attributes so we can get the guid
        if (ptInfo.GetTypeAttr(ptAttr) = S_OK) then
        begin
           DataType.FGuid:=ptAttr^.guid;
           // Release the type attributes
           ptInfo.ReleaseTypeAttr(ptAttr);
        end;
        // Release the type info
        ptInfo:=nil;
     end;
  end;

end;

{ TCoMember }

constructor TCoMember.Create;
begin
  inherited;
  fFuncName:='';
  fFuncKind:='';
  fFuncOffset:=0;
  fFuncSection:=0;
  foVft:=0;
end;

destructor TCoMember.Destroy;
begin

  inherited;
end;

procedure TCoMember.SetFuncInvokeKind(invkind: TInvokeKind);
begin
  case invkind of
  1: fInvokeKind:= 'INVOKE_FUNC';

  2: fInvokeKind:= 'INVOKE_PROPERTYGET';

  4: fInvokeKind:= 'INVOKE_PROPERTYPUT';
  8:fInvokeKind:= 'INVOKE_PROPERTYPUTREF';
  end;
end;

procedure TCoMember.SetFuncKindName(fk: TFuncKind);
begin
  case fk of
  0: fFuncKind := 'FUNC_VIRTUAL';
  1: fFuncKind := 'FUNC_PUREVIRTUAL';
  2: fFuncKind := 'FUNC_NONVIRTUAL';
  3: fFuncKind := 'FUNC_STATIC';
  4: fFuncKind := 'FUNC_DISPATCH';
  end;
end;

function TCoMember.SetFuncName(rtinfo: ITypeInfo;
  memid: TMemberID): HResult;
var
  ws:WideString;
begin
	Result := rtinfo.GetDocumentation( memid, @ws, nil, nil, nil );
  fFuncName:= ws;
end;

procedure TCoMember.SetFuncOffset(address: Cardinal);
var
	mbi:TMemoryBasicInformation ;
  hModule :Pointer;
begin
	// Tricky way to get the containing module from a linear address
	VirtualQuery( Pointer(address), mbi, sizeof(mbi) );

	// "AllocationBase" is the same as an HMODULE
	hModule := mbi.AllocationBase;

  fFuncOffset := Cardinal(Pointer(address)^) - Cardinal(hModule);
end;

end.

