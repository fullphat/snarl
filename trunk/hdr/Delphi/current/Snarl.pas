unit Snarl;

{$ifdef FPC}
  {$mode delphi}
{$endif}

interface

uses
  Windows, Messages, SysUtils;

(*
 * Registered window message and event identifiers (passed in wParam when either SNARL_GLOBAL_MSG or ReplyMsg is received)
 *)
const
  SNARL_GLOBAL_MSG              = 'SnarlGlobalEvent';
  SNARLAPP_MSG                  = 'SnarlAppMessage';

  SNARL_NOTIFICATION_CANCELLED  = 0;
  SNARL_LAUNCHED                = 1;
  SNARL_QUIT                    = 2;

  SNARL_NOTIFICATION_CLICKED    = 32;            // notification was right-clicked by user
  SNARL_NOTIFICATION_TIMED_OUT  = 33;
  SNARL_NOTIFICATION_ACK        = 34;            // notification was left-clicked by user


(*
 * Snarl Helper Functions
 *)

(*
function snShowMessage(ATitle, AText: String; ATimeout: Integer = 0; AIconPath: String = ''; AhwndReply: Integer = 0; AReplyMsg: Integer = 0): Integer;
function snUpdateMessage(AId: Integer; ATitle, AText: String): Boolean;
function snRevokeConfig(AHandle: HWND): Integer;
*)

function snDoRequest(Request: String): Integer;
function snarl_msg(): Integer;
function app_msg(): Integer;

function snarl_register(Signature: String; Name: String; Icon: String; Password: String = ''; ReplyTo: Integer = 0; ReplyWidth: Integer = 0): Integer;
function snarl_unregister(Signature: String; Password: String = ''): Integer;
function snarl_version(): Integer;

implementation

(*
 * snDoRequest() -- Primary V42 access function
 *
 * Locates the Snarl message handling window and sends <Request> to it.
 *
 *)
function snDoRequest(Request: String): Integer;
var
  hwnd: THandle;
  pcd:  TCopyDataStruct;

begin
  hwnd := FindWindow('w>Snarl', 'Snarl');

  //ShowMessage(IntToStr(hwnd));

  if not IsWindow(hwnd) then
    Result := -201                      // fix: should be constant

  else
    begin
      pcd.dwData := $534E4C03;            // "SNL",3
      pcd.cbData := StrLen(PChar(Request));
      pcd.lpData := PChar(Request);
      Result := Integer(SendMessage(hwnd, WM_COPYDATA, GetCurrentProcessId(), Integer(@pcd)));
  end;

end;


(*
 * snarl_msg() -- Returns global broadcast message identifier
 *
 *)
function snarl_msg(): Integer;
begin
  Result := RegisterWindowMessage(SNARL_GLOBAL_MSG);
end;


(*
 * app_msg() -- Returns app broadcast message identifier
 *
 *)
function app_msg(): Integer;
begin
  Result := RegisterWindowMessage(SNARLAPP_MSG);
end;



(************************************************************
 * Helper Functions
 ************************************************************)

(**
Public Function snarl_register(ByVal Signature As String, ByVal Name As String, ByVal Icon As String, Optional ByVal Password As String, Optional ByVal ReplyTo As Long, Optional ByVal Reply As Long, Optional ByVal Flags As SNARLAPP_FLAGS) As Long

    snarl_register = snDoRequest("register?app-sig=" & Signature & "&title=" & Name & "&icon=" & Icon & _
                                 IIf(Password <> "", "&password=" & Password, "") & _
                                 IIf(ReplyTo <> 0, "&reply-to=" & CStr(ReplyTo), "") & _
                                 IIf(Reply <> 0, "&reply=" & CStr(Reply), "") & _
                                 IIf(Flags <> 0, "&flags=" & CStr(Flags), ""))

End Function
**)


(*
 * snarl_register() -- Register an application with Snarl
 *
 * .
 *
 *)
function snarl_register(Signature: String; Name: String; Icon: String; Password: String = ''; ReplyTo: Integer = 0; ReplyWidth: Integer = 0): Integer;

begin

  Result := Integer(snDoRequest('register?app-sig=' + Signature + '&title=' + Name + '&icon=' + Icon));

end;



(*
 * snarl_unregister() -- Unregister an application with Snarl
 *
 * .
 *
 *)
function snarl_unregister(Signature: String; Password: String = ''): Integer;

begin

  Result := Integer(snDoRequest('unregister?app-sig=' + Signature + '&password=' + Password));

end;



(*
 * snarl_version() -- Returns running Snarl system version number
 *
 * .
 *
 *)
function snarl_version(): Integer;

begin

  Result := Integer(snDoRequest('version'));

end;



(*
function snGetVersion(var Major, Minor: Word): Boolean;
var
  pss: TSnarlStruct;
  hr: Integer;
begin
  _Clear(pss);
  pss.Cmd := SNARL_GET_VERSION;
  //hr := Integer(_Send(pss));
  Result := hr <> 0;
  if Result then
  begin
    Major := HiWord(hr);
    Minor := LoWord(hr);
  end;
end;
*)


end.
