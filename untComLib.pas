unit untComLib;

interface

uses
  Windows,Classes,ActiveX,ComObj,SysUtils,Imagehlp;


type

  {* TFunction *}
  TFunction = class
  private
    fFuncName:String;
    fFuncSection:Word;
    fFuncOffset:Cardinal;
    foVft:Word;
    fFuncKind:String;
  protected
    procedure SetFuncOffset(address:Cardinal);
  public
    function SetFuncName( rtinfo: ITypeInfo; memid: TMemberID ):HResult;
    property Name :String Read fFuncName ;
    property Section:Word read fFuncSection write fFuncSection;
    property Offset:Cardinal Read fFuncOffset write SetFuncOffset;
    property oVft:Word Read  foVft write foVft;
    property FuncKind:string read fFuncKind ;
    procedure SetFuncKindName(fk: TFuncKind);
  end;
  {* TFunctionList *}
  TFunctionList = class (TList)
  private
    function GetItem(AIndex: Integer): TFunction ;
    function Load(tinfo: ITypeInfo; typeAttr: PTypeAttr):HResult;
  public
    constructor create;
    destructor Destroy; override;
    function add(aFunction:TFunction): integer;
    property Items[AIndex: Integer] : TFunction read getitem;
    procedure Clear; override;
  end;

  {* TComInfo *}
  TComInfo = Class
  private
    fClass:String;
    fGuid:String;
    fTypeName:String;
    fSizeVft: Word;
    fTypeFlags :Word;
    fFunctions:TFunctionList;
  public
    constructor create;
    destructor Destroy; override;

    property CoClassName:String Read fClass write fClass;
    property Guid:String Read fGuid write fGuid;
    property TypeName:String Read fTypeName write fTypeName;
    property Functions:TFunctionList Read fFunctions;
    property SizeVft: Word Read fSizeVft Write fSizeVft;
    property TypeFlags :Word read fTypeFlags write fTypeflags;
  end;
  {* TComInfoList *}
  TComInfoList = class (TList)
  private
    function GetItem(AIndex: Integer): TComInfo;
  public
    constructor create; overload;
    constructor create(FileName:String);overLoad;
    destructor Destroy; override;
    function add(AComInfo:TComInfo): integer;
    property Items[AIndex: Integer] : TComInfo read getitem;
    procedure Clear; override;
    function Load(FileName:String):HResult;
  end;

implementation

{ TComInfo }

constructor TComInfo.create;
begin
  fFunctions:=TFunctionList.create ;
end;

destructor TComInfo.Destroy;
begin
  try
    FreeAndNil(fFunctions);
  finally
    inherited;
  end;
end;

{ TComInfoList }

constructor TComInfoList.create;
begin
  inherited Create;
  CoInitialize( nil );
end;

constructor TComInfoList.create(FileName: String);
begin
  inherited Create;
  CoInitialize( nil );
  Load(FileName);
end;

destructor TComInfoList.Destroy;
begin
  try
    //Clear;
    CoUninitialize();
  Finally
    inherited Destroy;
  end;
end;

function TComInfoList.add(AComInfo: TComInfo): integer;
begin
  result := inherited Add(AComInfo);
end;

function TComInfoList.GetItem(AIndex: Integer): TComInfo;
begin
  result := TComInfo( inherited Items[ AIndex ] );
end;

procedure TComInfoList.Clear;
var
  i:Integer;
begin
  for i:=0 to Count -1 do
    Items[i].Functions.Clear ;;
  inherited Clear;
end;


function TComInfoList.Load(FileName: String): HResult;
var
    tlib: ITypeLib;
    swFile: WideString;
    i:integer;
    tinfo:ITypeInfo;

    typeAttr: PTypeAttr;
    sCoClass: WideString;
    ci: TComInfo;
    reftype: HRefType;
begin

  Clear;
  
  try
    //============================================================================
    // Given a filename for a typelib, attempt to get an ITypeLib for it.  Send
    // the resultant ITypeLib instance to EnumTypeLib.
    //============================================================================
    swFile := WideString(FileName);
    Result := LoadTypeLib(@swFile[1], tlib);
    if Result = S_OK then
      //============================================================================
      // Enumerate through all the ITypeInfo instances in an ITypeLib.  Pass each
      // instance to ProcessTypeInfo.
      //============================================================================
      //EnumTypeInfos(tlib);
      for i := 0 to tlib.GetTypeInfoCount()-1 do
      begin
        Result := tlib.GetTypeInfo( i, tinfo );
        if Result = S_OK then begin
          //============================================================================
          // Top level handling code for a single ITypeInfo extracted from a typelib
          //============================================================================
          //ProcessTypeInfo( tinfo );
          try
            Result := tinfo.GetTypeAttr(typeAttr);
            if Result = S_OK then
              if ((typeAttr.wTypeFlags and TYPEFLAG_FHIDDEN) = 0) and
                  (typeAttr.cImplTypes > 0) then begin
                Result:= tinfo.GetDocumentation(-1, @sCoClass, nil, nil, nil);
                if Result = S_OK then begin
                  ci:=TComInfo.Create ;
                  ci.CoClassName :=sCoClass;
                  ci.Guid:= GUIDToString(typeAttr.guid);
                  //ci.TypeName := GetTypeKindName( typeAttr.typekind );
                  ci.SizeVft := typeAttr.cbSizeVft;
                  ci.TypeFlags := typeAttr.wTypeFlags;
                  ci.Functions.Load(tinfo,typeAttr);
                  Add(ci);
                end;
              end;
            finally
              tinfo.ReleaseTypeAttr(typeAttr);
            end;
        end;
      end;
  finally
  end;
end;

{ TFunction }

//============================================================================
// Given an ITypeInfo instance, retrieve the name.
//=============================================================================
function TFunction.SetFuncName( rtinfo: ITypeInfo; memid: TMemberID  ):HResult;
var
  ws:WideString;
begin
	Result := rtinfo.GetDocumentation( memid, @ws, nil, nil, nil );
  fFuncName:= ws;
end;

procedure TFunction.SetFuncKindName(fk: TFuncKind);
begin
  case fk of
  0: fFuncKind := 'FUNC_VIRTUAL';
  1: fFuncKind := 'FUNC_PUREVIRTUAL';
  2: fFuncKind := 'FUNC_NONVIRTUAL';
  3: fFuncKind := 'FUNC_STATIC';
  4: fFuncKind := 'FUNC_DISPATCH';
  end;
end;

procedure TFunction.SetFuncOffset(address:Cardinal);
var
	mbi:TMemoryBasicInformation ;
  hModule :Pointer;
begin
	// Tricky way to get the containing module from a linear address
	VirtualQuery( Pointer(address), mbi, sizeof(mbi) );

	// "AllocationBase" is the same as an HMODULE
	hModule := mbi.AllocationBase;

  fFuncOffset := address - Cardinal(hModule);
end;

{ TFunctionList }

constructor TFunctionList.create;
begin
  inherited Create;
end;

destructor TFunctionList.Destroy;
begin
  try
    //Clear;
  Finally
    inherited Destroy;
  end;
end;

function TFunctionList.add(aFunction: TFunction): integer;
begin
  result := inherited Add(aFunction);
end;

function TFunctionList.GetItem(AIndex: Integer): TFunction;
begin
  result := TFunction( inherited Items[ AIndex ] );
end;

procedure TFunctionList.Clear;
begin

  inherited Clear;
end;

function TFunctionList.Load(tinfo: ITypeInfo; typeAttr: PTypeAttr): HResult;
var
  i,j:integer;
  reftype: HRefType;
	rtinfo: ITypeInfo;
  rtypeAttr: PTypeAttr;
  pv:IUnknown;

  pVTable:Pointer;
  funcDesc:PFuncDesc;
  aFunc:TFunction;
begin
  Result := S_OK;

  if ( TKIND_COCLASS = typeAttr.typekind ) then
    for i:=0 to typeAttr.cImplTypes -1 do  begin
      Result := tinfo.GetRefTypeOfImplType(i,reftype);
      if Result = S_OK then begin
          //============================================================================
          // Given a TKIND_COCLASS ITypeInfo, get the ITypeInfo that describes the
          // referenced (HREFTYPE) TKIND_DISPATCH or TKIND_INTERFACE.  Pass that
          // ITypeInfo to EnumTypeInfoMembers.
          //============================================================================
          //ProcessReferencedTypeInfo(tinfo,pta,reftype);
        	Result := tinfo.GetRefTypeInfo(reftype, rtinfo);
          if Result = S_OK then begin
            rtinfo.GetTypeAttr( rtypeAttr );
            try
              Result := CoCreateInstance( typeAttr.guid,
                                      nil,					// pUnkOuter
                                      CLSCTX_INPROC_SERVER ,   //or CLSCTX_INPROC_HANDLER
                                      rtypeAttr.guid,
                                      pv );
              if ( (S_OK = Result) and Assigned(pv) )then
              begin
                //============================================================================
                // Enumerate through each member of an ITypeInfo.  Send the method name and
                // address to the CoClassSymsAddSymbol function.
                //=============================================================================
                //EnumTypeInfoMembers( rtinfo, rpta, Pointer(pv) );
                pVTable := pointer(pointer(pv)^);
                for j := 0 to rtypeAttr.cFuncs-1 do
                begin
                  rtinfo.GetFuncDesc( j, funcDesc );
                  aFunc:=TFunction.Create ;
                  aFunc.SetFuncName ( rtinfo,funcDesc.memid  );
                  aFunc.oVft := funcDesc.oVft;
                  aFunc.SetFuncKindName ( funcDesc.funckind );
                  // Index into the vtable to retrieve the method's virtual address
                  aFunc.Offset:=Cardinal(Pointer(Cardinal(pVTable) + Cardinal(funcDesc.oVft))^);

                  Add(aFunc);
                  rtinfo.ReleaseFuncDesc( funcDesc );
                end;
              end;
            finally
              rtinfo.ReleaseTypeAttr( rtypeAttr );
            end;
          end;
      end;
    end;
end;



end.
