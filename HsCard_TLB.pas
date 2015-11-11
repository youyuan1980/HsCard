unit HsCard_TLB;

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
// File generated on 2015-11-09 17:25:14 from Type Library described below.

// ************************************************************************  //
// Type Lib: D:\Program Files\Borland\Delphi7\Projects\HsCard\HsCard.tlb (1)
// LIBID: {EC8C6755-B7CC-44FD-879E-AEF049FE5D0C}
// LCID: 0
// Helpfile: 
// HelpString: HsCard Library
// DepndLst: 
//   (1) v2.0 stdole, (C:\WINDOWS\system32\stdole2.tlb)
// ************************************************************************ //
{$TYPEDADDRESS OFF} // Unit must be compiled without type-checked pointers. 
{$WARN SYMBOL_PLATFORM OFF}
{$WRITEABLECONST ON}
{$VARPROPSETTER ON}
interface

uses Windows, ActiveX, Classes, Graphics, OleCtrls, StdVCL, Variants;
  

// *********************************************************************//
// GUIDS declared in the TypeLibrary. Following prefixes are used:        
//   Type Libraries     : LIBID_xxxx                                      
//   CoClasses          : CLASS_xxxx                                      
//   DISPInterfaces     : DIID_xxxx                                       
//   Non-DISP interfaces: IID_xxxx                                        
// *********************************************************************//
const
  // TypeLibrary Major and minor versions
  HsCardMajorVersion = 1;
  HsCardMinorVersion = 0;

  LIBID_HsCard: TGUID = '{EC8C6755-B7CC-44FD-879E-AEF049FE5D0C}';

  IID_IActiveFormX: TGUID = '{A28FDA56-A15D-4984-9005-AD200BD4634F}';
  DIID_IActiveFormXEvents: TGUID = '{B0CE2BFF-952E-40DD-84E2-2CD65650E603}';
  CLASS_ActiveFormX: TGUID = '{F4B1D277-33BE-4476-A13E-C3D47745B06B}';

// *********************************************************************//
// Declaration of Enumerations defined in Type Library                    
// *********************************************************************//
// Constants for enum TxActiveFormBorderStyle
type
  TxActiveFormBorderStyle = TOleEnum;
const
  afbNone = $00000000;
  afbSingle = $00000001;
  afbSunken = $00000002;
  afbRaised = $00000003;

// Constants for enum TxPrintScale
type
  TxPrintScale = TOleEnum;
const
  poNone = $00000000;
  poProportional = $00000001;
  poPrintToFit = $00000002;

// Constants for enum TxMouseButton
type
  TxMouseButton = TOleEnum;
const
  mbLeft = $00000000;
  mbRight = $00000001;
  mbMiddle = $00000002;

type

// *********************************************************************//
// Forward declaration of types defined in TypeLibrary                    
// *********************************************************************//
  IActiveFormX = interface;
  IActiveFormXDisp = dispinterface;
  IActiveFormXEvents = dispinterface;

// *********************************************************************//
// Declaration of CoClasses defined in Type Library                       
// (NOTE: Here we map each CoClass to its Default Interface)              
// *********************************************************************//
  ActiveFormX = IActiveFormX;


// *********************************************************************//
// Declaration of structures, unions and aliases.                         
// *********************************************************************//
  PPUserType1 = ^IFontDisp; {*}


// *********************************************************************//
// Interface: IActiveFormX
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {A28FDA56-A15D-4984-9005-AD200BD4634F}
// *********************************************************************//
  IActiveFormX = interface(IDispatch)
    ['{A28FDA56-A15D-4984-9005-AD200BD4634F}']
    function Get_Visible: WordBool; safecall;
    procedure Set_Visible(Value: WordBool); safecall;
    function Get_AutoScroll: WordBool; safecall;
    procedure Set_AutoScroll(Value: WordBool); safecall;
    function Get_AutoSize: WordBool; safecall;
    procedure Set_AutoSize(Value: WordBool); safecall;
    function Get_AxBorderStyle: TxActiveFormBorderStyle; safecall;
    procedure Set_AxBorderStyle(Value: TxActiveFormBorderStyle); safecall;
    function Get_Caption: WideString; safecall;
    procedure Set_Caption(const Value: WideString); safecall;
    function Get_Color: OLE_COLOR; safecall;
    procedure Set_Color(Value: OLE_COLOR); safecall;
    function Get_Font: IFontDisp; safecall;
    procedure Set_Font(const Value: IFontDisp); safecall;
    procedure _Set_Font(var Value: IFontDisp); safecall;
    function Get_KeyPreview: WordBool; safecall;
    procedure Set_KeyPreview(Value: WordBool); safecall;
    function Get_PixelsPerInch: Integer; safecall;
    procedure Set_PixelsPerInch(Value: Integer); safecall;
    function Get_PrintScale: TxPrintScale; safecall;
    procedure Set_PrintScale(Value: TxPrintScale); safecall;
    function Get_Scaled: WordBool; safecall;
    procedure Set_Scaled(Value: WordBool); safecall;
    function Get_Active: WordBool; safecall;
    function Get_DropTarget: WordBool; safecall;
    procedure Set_DropTarget(Value: WordBool); safecall;
    function Get_HelpFile: WideString; safecall;
    procedure Set_HelpFile(const Value: WideString); safecall;
    function Get_ScreenSnap: WordBool; safecall;
    procedure Set_ScreenSnap(Value: WordBool); safecall;
    function Get_SnapBuffer: Integer; safecall;
    procedure Set_SnapBuffer(Value: Integer); safecall;
    function Get_DoubleBuffered: WordBool; safecall;
    procedure Set_DoubleBuffered(Value: WordBool); safecall;
    function Get_AlignDisabled: WordBool; safecall;
    function Get_VisibleDockClientCount: Integer; safecall;
    function Get_Enabled: WordBool; safecall;
    procedure Set_Enabled(Value: WordBool); safecall;
    procedure GetCardInfo; safecall;
    function Get_NAME: WideString; safecall;
    procedure Set_NAME(const Value: WideString); safecall;
    function Get_ADDRESS: WideString; safecall;
    procedure Set_ADDRESS(const Value: WideString); safecall;
    function Get_BIRTHDAY: WideString; safecall;
    procedure Set_BIRTHDAY(const Value: WideString); safecall;
    function Get_DEPARTMENT: WideString; safecall;
    procedure Set_DEPARTMENT(const Value: WideString); safecall;
    function Get_IDCODE: WideString; safecall;
    procedure Set_IDCODE(const Value: WideString); safecall;
    function Get_MSG: WideString; safecall;
    procedure Set_MSG(const Value: WideString); safecall;
    function Get_NATION: WideString; safecall;
    procedure Set_NATION(const Value: WideString); safecall;
    function Get_SEX: WideString; safecall;
    procedure Set_SEX(const Value: WideString); safecall;
    function Get_STARTTOENDDATE: WideString; safecall;
    procedure Set_STARTTOENDDATE(const Value: WideString); safecall;
    function Get_ReadSysDirectory: WideString; safecall;
    procedure Set_ReadSysDirectory(const Value: WideString); safecall;
    function Get_PhotoImageBytes: WideString; safecall;
    procedure Set_PhotoImageBytes(const Value: WideString); safecall;
    property Visible: WordBool read Get_Visible write Set_Visible;
    property AutoScroll: WordBool read Get_AutoScroll write Set_AutoScroll;
    property AutoSize: WordBool read Get_AutoSize write Set_AutoSize;
    property AxBorderStyle: TxActiveFormBorderStyle read Get_AxBorderStyle write Set_AxBorderStyle;
    property Caption: WideString read Get_Caption write Set_Caption;
    property Color: OLE_COLOR read Get_Color write Set_Color;
    property Font: IFontDisp read Get_Font write Set_Font;
    property KeyPreview: WordBool read Get_KeyPreview write Set_KeyPreview;
    property PixelsPerInch: Integer read Get_PixelsPerInch write Set_PixelsPerInch;
    property PrintScale: TxPrintScale read Get_PrintScale write Set_PrintScale;
    property Scaled: WordBool read Get_Scaled write Set_Scaled;
    property Active: WordBool read Get_Active;
    property DropTarget: WordBool read Get_DropTarget write Set_DropTarget;
    property HelpFile: WideString read Get_HelpFile write Set_HelpFile;
    property ScreenSnap: WordBool read Get_ScreenSnap write Set_ScreenSnap;
    property SnapBuffer: Integer read Get_SnapBuffer write Set_SnapBuffer;
    property DoubleBuffered: WordBool read Get_DoubleBuffered write Set_DoubleBuffered;
    property AlignDisabled: WordBool read Get_AlignDisabled;
    property VisibleDockClientCount: Integer read Get_VisibleDockClientCount;
    property Enabled: WordBool read Get_Enabled write Set_Enabled;
    property NAME: WideString read Get_NAME write Set_NAME;
    property ADDRESS: WideString read Get_ADDRESS write Set_ADDRESS;
    property BIRTHDAY: WideString read Get_BIRTHDAY write Set_BIRTHDAY;
    property DEPARTMENT: WideString read Get_DEPARTMENT write Set_DEPARTMENT;
    property IDCODE: WideString read Get_IDCODE write Set_IDCODE;
    property MSG: WideString read Get_MSG write Set_MSG;
    property NATION: WideString read Get_NATION write Set_NATION;
    property SEX: WideString read Get_SEX write Set_SEX;
    property STARTTOENDDATE: WideString read Get_STARTTOENDDATE write Set_STARTTOENDDATE;
    property ReadSysDirectory: WideString read Get_ReadSysDirectory write Set_ReadSysDirectory;
    property PhotoImageBytes: WideString read Get_PhotoImageBytes write Set_PhotoImageBytes;
  end;

// *********************************************************************//
// DispIntf:  IActiveFormXDisp
// Flags:     (4416) Dual OleAutomation Dispatchable
// GUID:      {A28FDA56-A15D-4984-9005-AD200BD4634F}
// *********************************************************************//
  IActiveFormXDisp = dispinterface
    ['{A28FDA56-A15D-4984-9005-AD200BD4634F}']
    property Visible: WordBool dispid 201;
    property AutoScroll: WordBool dispid 202;
    property AutoSize: WordBool dispid 203;
    property AxBorderStyle: TxActiveFormBorderStyle dispid 204;
    property Caption: WideString dispid -518;
    property Color: OLE_COLOR dispid -501;
    property Font: IFontDisp dispid -512;
    property KeyPreview: WordBool dispid 205;
    property PixelsPerInch: Integer dispid 206;
    property PrintScale: TxPrintScale dispid 207;
    property Scaled: WordBool dispid 208;
    property Active: WordBool readonly dispid 209;
    property DropTarget: WordBool dispid 210;
    property HelpFile: WideString dispid 211;
    property ScreenSnap: WordBool dispid 212;
    property SnapBuffer: Integer dispid 213;
    property DoubleBuffered: WordBool dispid 214;
    property AlignDisabled: WordBool readonly dispid 215;
    property VisibleDockClientCount: Integer readonly dispid 216;
    property Enabled: WordBool dispid -514;
    procedure GetCardInfo; dispid 227;
    property NAME: WideString dispid 217;
    property ADDRESS: WideString dispid 218;
    property BIRTHDAY: WideString dispid 219;
    property DEPARTMENT: WideString dispid 220;
    property IDCODE: WideString dispid 222;
    property MSG: WideString dispid 223;
    property NATION: WideString dispid 224;
    property SEX: WideString dispid 225;
    property STARTTOENDDATE: WideString dispid 226;
    property ReadSysDirectory: WideString dispid 228;
    property PhotoImageBytes: WideString dispid 221;
  end;

// *********************************************************************//
// DispIntf:  IActiveFormXEvents
// Flags:     (4096) Dispatchable
// GUID:      {B0CE2BFF-952E-40DD-84E2-2CD65650E603}
// *********************************************************************//
  IActiveFormXEvents = dispinterface
    ['{B0CE2BFF-952E-40DD-84E2-2CD65650E603}']
    procedure OnActivate; dispid 201;
    procedure OnClick; dispid 202;
    procedure OnCreate; dispid 203;
    procedure OnDblClick; dispid 204;
    procedure OnDestroy; dispid 205;
    procedure OnDeactivate; dispid 206;
    procedure OnKeyPress(var Key: Smallint); dispid 207;
    procedure OnPaint; dispid 208;
  end;


// *********************************************************************//
// OLE Control Proxy class declaration
// Control Name     : TActiveFormX
// Help String      : ActiveFormX Control
// Default Interface: IActiveFormX
// Def. Intf. DISP? : No
// Event   Interface: IActiveFormXEvents
// TypeFlags        : (34) CanCreate Control
// *********************************************************************//
  TActiveFormXOnKeyPress = procedure(ASender: TObject; var Key: Smallint) of object;

  TActiveFormX = class(TOleControl)
  private
    FOnActivate: TNotifyEvent;
    FOnClick: TNotifyEvent;
    FOnCreate: TNotifyEvent;
    FOnDblClick: TNotifyEvent;
    FOnDestroy: TNotifyEvent;
    FOnDeactivate: TNotifyEvent;
    FOnKeyPress: TActiveFormXOnKeyPress;
    FOnPaint: TNotifyEvent;
    FIntf: IActiveFormX;
    function  GetControlInterface: IActiveFormX;
  protected
    procedure CreateControl;
    procedure InitControlData; override;
  public
    procedure GetCardInfo;
    property  ControlInterface: IActiveFormX read GetControlInterface;
    property  DefaultInterface: IActiveFormX read GetControlInterface;
    property Visible: WordBool index 201 read GetWordBoolProp write SetWordBoolProp;
    property Active: WordBool index 209 read GetWordBoolProp;
    property DropTarget: WordBool index 210 read GetWordBoolProp write SetWordBoolProp;
    property HelpFile: WideString index 211 read GetWideStringProp write SetWideStringProp;
    property ScreenSnap: WordBool index 212 read GetWordBoolProp write SetWordBoolProp;
    property SnapBuffer: Integer index 213 read GetIntegerProp write SetIntegerProp;
    property DoubleBuffered: WordBool index 214 read GetWordBoolProp write SetWordBoolProp;
    property AlignDisabled: WordBool index 215 read GetWordBoolProp;
    property VisibleDockClientCount: Integer index 216 read GetIntegerProp;
    property Enabled: WordBool index -514 read GetWordBoolProp write SetWordBoolProp;
  published
    property Anchors;
    property  ParentColor;
    property  ParentFont;
    property  Align;
    property  DragCursor;
    property  DragMode;
    property  ParentShowHint;
    property  PopupMenu;
    property  ShowHint;
    property  TabOrder;
    property  OnDragDrop;
    property  OnDragOver;
    property  OnEndDrag;
    property  OnEnter;
    property  OnExit;
    property  OnStartDrag;
    property AutoScroll: WordBool index 202 read GetWordBoolProp write SetWordBoolProp stored False;
    property AutoSize: WordBool index 203 read GetWordBoolProp write SetWordBoolProp stored False;
    property AxBorderStyle: TOleEnum index 204 read GetTOleEnumProp write SetTOleEnumProp stored False;
    property Caption: WideString index -518 read GetWideStringProp write SetWideStringProp stored False;
    property Color: TColor index -501 read GetTColorProp write SetTColorProp stored False;
    property Font: TFont index -512 read GetTFontProp write SetTFontProp stored False;
    property KeyPreview: WordBool index 205 read GetWordBoolProp write SetWordBoolProp stored False;
    property PixelsPerInch: Integer index 206 read GetIntegerProp write SetIntegerProp stored False;
    property PrintScale: TOleEnum index 207 read GetTOleEnumProp write SetTOleEnumProp stored False;
    property Scaled: WordBool index 208 read GetWordBoolProp write SetWordBoolProp stored False;
    property NAME: WideString index 217 read GetWideStringProp write SetWideStringProp stored False;
    property ADDRESS: WideString index 218 read GetWideStringProp write SetWideStringProp stored False;
    property BIRTHDAY: WideString index 219 read GetWideStringProp write SetWideStringProp stored False;
    property DEPARTMENT: WideString index 220 read GetWideStringProp write SetWideStringProp stored False;
    property IDCODE: WideString index 222 read GetWideStringProp write SetWideStringProp stored False;
    property MSG: WideString index 223 read GetWideStringProp write SetWideStringProp stored False;
    property NATION: WideString index 224 read GetWideStringProp write SetWideStringProp stored False;
    property SEX: WideString index 225 read GetWideStringProp write SetWideStringProp stored False;
    property STARTTOENDDATE: WideString index 226 read GetWideStringProp write SetWideStringProp stored False;
    property ReadSysDirectory: WideString index 228 read GetWideStringProp write SetWideStringProp stored False;
    property PhotoImageBytes: WideString index 221 read GetWideStringProp write SetWideStringProp stored False;
    property OnActivate: TNotifyEvent read FOnActivate write FOnActivate;
    property OnClick: TNotifyEvent read FOnClick write FOnClick;
    property OnCreate: TNotifyEvent read FOnCreate write FOnCreate;
    property OnDblClick: TNotifyEvent read FOnDblClick write FOnDblClick;
    property OnDestroy: TNotifyEvent read FOnDestroy write FOnDestroy;
    property OnDeactivate: TNotifyEvent read FOnDeactivate write FOnDeactivate;
    property OnKeyPress: TActiveFormXOnKeyPress read FOnKeyPress write FOnKeyPress;
    property OnPaint: TNotifyEvent read FOnPaint write FOnPaint;
  end;

procedure Register;

resourcestring
  dtlServerPage = 'Servers';

  dtlOcxPage = 'ActiveX';

implementation

uses ComObj;

procedure TActiveFormX.InitControlData;
const
  CEventDispIDs: array [0..7] of DWORD = (
    $000000C9, $000000CA, $000000CB, $000000CC, $000000CD, $000000CE,
    $000000CF, $000000D0);
  CTFontIDs: array [0..0] of DWORD = (
    $FFFFFE00);
  CControlData: TControlData2 = (
    ClassID: '{F4B1D277-33BE-4476-A13E-C3D47745B06B}';
    EventIID: '{B0CE2BFF-952E-40DD-84E2-2CD65650E603}';
    EventCount: 8;
    EventDispIDs: @CEventDispIDs;
    LicenseKey: nil (*HR:$00000000*);
    Flags: $0000001D;
    Version: 401;
    FontCount: 1;
    FontIDs: @CTFontIDs);
begin
  ControlData := @CControlData;
  TControlData2(CControlData).FirstEventOfs := Cardinal(@@FOnActivate) - Cardinal(Self);
end;

procedure TActiveFormX.CreateControl;

  procedure DoCreate;
  begin
    FIntf := IUnknown(OleObject) as IActiveFormX;
  end;

begin
  if FIntf = nil then DoCreate;
end;

function TActiveFormX.GetControlInterface: IActiveFormX;
begin
  CreateControl;
  Result := FIntf;
end;

procedure TActiveFormX.GetCardInfo;
begin
  DefaultInterface.GetCardInfo;
end;

procedure Register;
begin
  RegisterComponents(dtlOcxPage, [TActiveFormX]);
end;

end.
