EnableExplicit

XIncludeFile "MSWORD.pbi"

Define._Application wdApp

OleInitialize_(0)

If CoCreateInstance_(?CLSID_Application, 0, #CLSCTX_LOCAL_SERVER, ?IID__Application, @wdApp) <> #S_OK
  MessageRequester("Warning:","Couldn't init oWord",0)
  End
EndIf

wdApp\Activate()
wdApp\put_WindowState(#wdWindowStateNormal)
wdApp\put_Width(400)
wdApp\put_Height(300)
wdApp\put_Left(10)
wdApp\put_Top(10)
wdApp\put_Caption("Word from PureBasic")
wdApp\put_Visible(#True)
wdApp\Release()

OleUninitialize_()

