
Option Explicit
Option Compare Text
'MsgBox VN
' __   _____   _ ?
' \ \ / / _ | / \
'  \ \ /| _ \/ / \
'   \_/ |___/_/ \_\
'

Private Const DelayDoubleClick As Double = 0.95
Private Const GroupCenter = "GroupCenter"
Private CallerAutoScroll$
Private TimeAutoScroll%
Private AllowAutoScroll As Boolean
Private AllowTurnSheet As Boolean
Private ScrollDown%, ScrollToRight%
#If VBA7 Then
Private Declare PtrSafe Function getTickCount Lib "kernel32.dll" Alias "GetTickCount" () As Long
Private Declare PtrSafe Function SHGetFolderPath Lib "shell32.dll" Alias "SHGetFolderPathA" (ByVal hwndOwner As Long, ByVal nFolder As Long, ByVal hToken As Long, ByVal dwFlags As Long, ByVal lpszPath As String) As Long
#Else
Private Declare Function getTickCount Lib "kernel32.dll" Alias "GetTickCount" () As Long
Private Declare Function SHGetFolderPath Lib "shell32.dll" Alias "SHGetFolderPathA" (ByVal hwndOwner As Long, ByVal nFolder As Long, ByVal hToken As Long, ByVal dwFlags As Long, ByVal lpszPath As String) As Long
#End If


Public Selen As Object

Public Const Si_Chrome$ = "https://www.google.com/chrome/"
Public Const Si_SeleniumBasic$ = "https://github.com/florentbr/SeleniumBasic/releases/tag/v2.0.9.0"
Public Const Si_ChromeDriver$ = "http://chromedriver.chromium.org/downloads"



Public Const ZaloAppTitle = "Zalo Web"
Public Const ZaloAppSite = "https://chat.zalo.me/"
Public Const ZaloAppSite2 = "https://id.zalo.me/"

Public Const shapeStartZaloApp = "btnStartZaloApp"
Public Const shapeLogin = "btnLogin"

Public Const procZaloAppPlaying = "ZaloAppPlaying"


Public Const n_ = vbNullString

#If Mac Then

#Else
  #If VBA7 Then
    Private Declare PtrSafe Function GlobalUnlock Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
    Private Declare PtrSafe Function GlobalLock Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
    Private Declare PtrSafe Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As LongPtr) As LongPtr
    Private Declare PtrSafe Function CloseClipboard Lib "user32" () As Long
    Private Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hwnd As LongPtr) As LongPtr
    Public Declare PtrSafe Function EmptyClipboard Lib "user32" () As Long
    Private Declare PtrSafe Function lstrcpy Lib "kernel32" (ByVal lpString1 As Any, ByVal lpString2 As Any) As LongPtr
    Private Declare PtrSafe Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As LongPtr) As LongPtr
    Private Declare PtrSafe Sub CopyMem Lib "kernel32" Alias "RtlMoveMemory" (destination As Any, Source As Any, ByVal Length As LongPtr)
  #Else
    Private Declare Sub CopyMem Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)
    Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
    Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
    Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
    Private Declare Function CloseClipboard Lib "user32" () As Long
    Private Declare Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As Long
    Public Declare Function EmptyClipboard Lib "user32" () As Long
    Private Declare Function lstrcpy Lib "kernel32" (ByVal lpString1 As Any, ByVal lpString2 As Any) As Long
    Private Declare Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As Long) As Long
  #End If
#End If

#If VBA7 Then
  Private Declare PtrSafe Function GetClipboardData Lib "user32" (ByVal wFormat As Long) As Long
  Private Declare PtrSafe Function IsClipboardFormatAvailable Lib "user32" (ByVal wFormat As Long) As Long
  Private Declare PtrSafe Function DragQueryFile Lib "shell32.dll" Alias "DragQueryFileA" (ByVal hDrop As Long, ByVal UINT As Long, ByVal lpStr As String, ByVal ch As Long) As Long
  Private Declare PtrSafe Function DragQueryPoint Lib "shell32.dll" (ByVal hDrop As Long, lpPoint As POINTAPI) As Long
  Private Declare PtrSafe Function GlobalFree Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
#Else
  Private Declare Function GetClipboardData Lib "user32" (ByVal wFormat As Long) As Long
  Private Declare Function IsClipboardFormatAvailable Lib "user32" (ByVal wFormat As Long) As Long
  Private Declare Function DragQueryFile Lib "shell32.dll" Alias "DragQueryFileA" (ByVal hDrop As Long, ByVal UINT As Long, ByVal lpStr As String, ByVal ch As Long) As Long
  Private Declare Function DragQueryPoint Lib "shell32.dll" (ByVal hDrop As Long, lpPoint As POINTAPI) As Long
  Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
#End If
Private Const CF_TEXT = 1
Private Const CF_BITMAP = 2
Private Const CF_METAFILEPICT = 3
Private Const CF_SYLK = 4
Private Const CF_DIF = 5
Private Const CF_TIFF = 6
Private Const CF_OEMTEXT = 7
Private Const CF_DIB = 8
Private Const CF_PALETTE = 9
Private Const CF_PENDATA = 10
Private Const CF_RIFF = 11
Private Const CF_WAVE = 12
Private Const CF_UNICODETEXT = 13
Private Const CF_ENHMETAFILE = 14
Private Const CF_HDROP = 15
Private Const CF_LOCALE = 16
Private Const CF_MAX = 17
Private Const CFSTR_SHELLIDLIST As String = "Shell IDList Array"
Private Const CFSTR_SHELLIDLISTOFFSET As String = "Shell Object Offsets"
Private Const CFSTR_NETRESOURCES As String = "Net Resource"
Private Const CFSTR_FILEDESCRIPTOR As String = "FileGroupDescriptor"
Private Const CFSTR_FILECONTENTS As String = "FileContents"
Private Const CFSTR_FILENAME As String = "FileName"
Private Const CFSTR_PRINTERGROUP As String = "PrinterFriendlyName"
Private Const CFSTR_FILENAMEMAP As String = "FileNameMap"

Private Const GMEM_FIXED = &H0
Private Const GMEM_MOVEABLE = &H2
Private Const GMEM_NOCOMPACT = &H10
Private Const GMEM_NODISCARD = &H20
Private Const GMEM_ZEROINIT = &H40
Private Const GMEM_MODIFY = &H80
Private Const GMEM_DISCARDABLE = &H100
Private Const GMEM_NOT_BANKED = &H1000
Private Const GMEM_SHARE = &H2000
Private Const GMEM_DDESHARE = &H2000
Private Const GMEM_NOTIFY = &H4000
Private Const GMEM_LOWER = GMEM_NOT_BANKED
Private Const GMEM_VALID_FLAGS = &H7F72
Private Const GMEM_INVALID_HANDLE = &H8000
Private Const GHND = (GMEM_MOVEABLE Or GMEM_ZEROINIT)
Private Const GPTR = (GMEM_FIXED Or GMEM_ZEROINIT)
Private Type DROPFILES
  pFiles As Long
  pt As POINTAPI
  fNC As Long
  fWide As Long
End Type

#If VBA7 Then
Public Declare PtrSafe Function GetForegroundWindow Lib "user32" () As LongPtr
Private Declare PtrSafe Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As LongPtr, lpdwProcessId As Long) As Long
Private Declare PtrSafe Function AttachThreadInput Lib "user32" (ByVal idAttach As Long, ByVal idAttachTo As Long, ByVal fAttach As Long) As Long
Private Declare PtrSafe Function SetForegroundWindow Lib "user32" (ByVal hwnd As LongPtr) As Long
Private Declare PtrSafe Function IsIconic Lib "user32" (ByVal hwnd As LongPtr) As Long
Private Declare PtrSafe Function ShowWindow Lib "user32.dll" (ByVal hwnd As LongPtr, ByVal lCmdShow As Long) As Boolean
Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
Private Declare PtrSafe Function GetParent Lib "user32" (ByVal hwnd As LongPtr) As Long
Private Declare PtrSafe Function GetWindow Lib "user32" (ByVal hwnd As LongPtr, ByVal wCmd As Long) As Long
#Else
Public Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function GetWindowThreadProcessId Lib "USER32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Private Declare Function AttachThreadInput Lib "USER32" (ByVal idAttach As Long, ByVal idAttachTo As Long, ByVal fAttach As Long) As Long
Private Declare Function SetForegroundWindow Lib "USER32" (ByVal hwnd As Long) As Long
Private Declare Function IsIconic Lib "USER32" (ByVal hwnd As Long) As Long
Private Declare Function ShowWindow Lib "user32.dll" (ByVal hwnd As Long, ByVal lCmdShow As Long) As Boolean
Private Declare Function FindWindow Lib "USER32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function GetParent Lib "USER32" (ByVal hwnd As Long) As Long
Private Declare Function GetWindow Lib "USER32" (ByVal hwnd As LongPtr, ByVal wCmd As Long) As Long
#End If

Private GoNextParent As Boolean
Public Const MAINPORT9515 = 9515

Private Function txtSend()
  txtSend = "G" & ChrW(7917) & "i"
End Function
Private Function txtSendAll()
  txtSendAll = "G" & ChrW(7917) & "i t" & ChrW(7845) & "t c" & ChrW(7843)
End Function
Private Function txtSending()
  txtSending = ChrW(272) & "ang g" & ChrW(7917) & "i"
End Function

Private Function chxReSend() As Boolean
  On Error Resume Next
  chxReSend = shZaloExcel.CheckBoxes("chxReSend").Value = 1
End Function
Private Function chxSleepWindow() As Boolean
  On Error Resume Next
  chxSleepWindow = shZaloExcel.CheckBoxes("chxSleepWindow").Value = 1
End Function


Private Sub btnSend(Optional action As Boolean)
  On Error Resume Next
  shZaloExcel.Shapes("btnSend").TextFrame2.TextRange.Characters.Text = IIf(action, txtSending, txtSend)
End Sub
Private Sub btnSendAll(Optional action As Boolean)
  On Error Resume Next
  shZaloExcel.Shapes("btnSendAll").TextFrame2.TextRange.Characters.Text = IIf(action, txtSending, txtSendAll)
End Sub



Sub Login_click_()
  If ZaloAppLogin Then
    Alert ChrW(272) & ChrW(227) & " " & ChrW(273) & ChrW(259) & "ng nh" & ChrW(7853) & "p"
  End If
End Sub

Sub ZaloContact_click_()
  On Error Resume Next
  If Not ZaloAppGotoContact Then
    Exit Sub
  End If
  Dim o, D, cs, k%, n$, n2$, s$, url, rg
  Set D = VBA.CreateObject("Scripting.Dictionary")
  D.CompareMode = 1
  Set o = Selen.FindElementByClass("ReactVirtualized__Grid__innerScrollContainer", 200, False)
  If o Is Nothing Then
    Exit Sub
  End If
R1:
  Set cs = o.FindElementsByXPath("//*[contains(@id, 'friend-item-')]", , 200)
  n = cs(1).FindElementByXPath(".//*[contains(@class, 'conv-item-title__name')]", 200, False).Attribute("innerText")
  If n <> n2 And n <> "My Cloud" And n <> "Cloud c" & ChrW(7911) & "a t" & ChrW(244) & "i" Then
    n2 = n
    cs(1).ScrollIntoView False
    GoTo R1
  End If

r:
  Set cs = o.FindElementsByXPath("//*[contains(@id, 'friend-item-')]", , 200)
  For k = 1 To cs.Count
    Set o = cs(k)
    s = "": n = ""
    n = o.FindElementByXPath(".//*[contains(@class, 'conv-item-title__name')]", 200, False).Attribute("innerText")
    D(n) = ""
  Next
  If n <> n2 Then
    n2 = n
    o.ScrollIntoView False
    GoTo r
  End If
  If D.Count Then
    ReDim a(1 To D.Count, 1 To 1)
    For k = 0 To D.Count - 1
      a(k + 1, 1) = D.keys()(k)
    Next
    Set rg = shZaloExcel.Range("B4")
    Set rg = rg(10000, 1).End(3).Offset(1, 0)
    rg.Resize(UBound(a)).Value = a
    Alert "Danh b" & ChrW(7841) & " Zalo c" & ChrW(7911) & "a b" & ChrW(7841) & "n c" & ChrW(243) & " (" & D.Count & ") ng" & ChrW(432) & ChrW(7901) & "i b" & ChrW(7841) & "n"
  Else
    Alert "Danh b" & ChrW(7841) & " r" & ChrW(7895) & "ng"
  End If
End Sub

Sub ZaloSend_click_()
  Dim w, rg, rg2 As Object, r
  Set rg = shZaloExcel.Range("B4")
  Set rg2 = Selection
  If TypeName(rg2) <> "Range" Then
    Alert "Ch" & ChrW(7885) & "n t" & ChrW(234) & "n/phone " & ChrW(273) & ChrW(7875) & " g" & ChrW(7917) & "i"
    btnSend 0
    Exit Sub
  End If
  On Error Resume Next
  Set rg2 = Intersect(rg.Resize(10000), rg2)
  On Error GoTo 0
  If Not rg2 Is Nothing Then
    ZaloSendAll rg2, False
  End If
End Sub

Sub ZaloSendAll_click_()
  Dim w, rg, lr&
  Set rg = shZaloExcel.Range("B4")
  lr = rg(10000, 1).End(3).row - rg.row + 1
  If lr < 1 Then
    Exit Sub
  End If
  ZaloSendAll rg.Resize(lr), True
End Sub

Private Sub ZaloSendAll(ByVal Cells As Range, Optional SendAll As Boolean)
  If Not ZaloAppGotoContact Then
    Alert "Vui l" & ChrW(242) & "ng " & ChrW(273) & ChrW(259) & "ng nh" & ChrW(7853) & "p Zalo tr" & ChrW(432) & ChrW(7899) & "c khi th" & ChrW(7921) & "c hi" & ChrW(7879) & "n"
    Exit Sub
  End If
  Call BringWindow
  Dim w, s, r, b, recipient$, status$, time As Date, Y As Boolean, bnt

  If SendAll Then
    btnSendAll 1
  Else
    btnSend 1
  End If
  bnt = ClipboardTitle
  Y = chxReSend
  For Each r In Cells
    If r(1, 1).Value <> Empty And r(1, 3).Value = bnt Then
      b = ZaloAppSearchAndSend(r(1, 1).Value, r(1, 2).Value, bnt, Y, recipient, status, time)
      r(1, 4).Value = status
      If time > 0 Then
        r(1, 5).Value = time
      End If
      r(1, 6).Value = recipient
    End If
  Next
  For Each r In Cells
    If r(1, 1).Value <> Empty And (r(1, 2).Value <> Empty Or r(1, 3).Value <> Empty) Then
      If r(1, 3).Value <> bnt Then
        s = r(1, 3).Value
        If s <> Empty Then
          If Not s Like "[[]*[]]" Then
            s = Split(Mid(s, 2, Len(s) - 2), """,""")
          End If
        End If
        b = ZaloAppSearchAndSend(r(1, 1).Value, r(1, 2).Value, s, Y, recipient, status, time)
        r(1, 4).Value = status
        If time > 0 Then
          r(1, 5).Value = time
        End If
        r(1, 6).Value = recipient
      End If
    End If
  Next
  If SendAll Then
    btnSendAll 0
  Else
    btnSend 0
  End If
  Alert "H" & ChrW(242) & "an th" & ChrW(224) & "nh"
  If SendAll Then
    If chxSleepWindow Then
      ThisWorkbook.Save
      VBA.CreateObject("WScript.Shell").Run "Rundll32.exe Powrprof.dll,SetSuspendState Sleep"
    End If
  End If
Exit Sub
End Sub
Private Sub BringWindow()

  BringWindowToFront GetChromeHandleByProcessID
End Sub
Function shZaloExcel() As Worksheet
  On Error Resume Next
  Set shZaloExcel = ThisWorkbook.Worksheets("Zalo Excel")
End Function

Public Function ZaloAppFriendSentMessenge(Id$, Messenge$) As Boolean
  Dim oj, o, cs
  Set oj = Selen.FindElementsByXPath("//*[@class='ReactVirtualized__Grid__innerScrollContainer']/child::*", 200, False)
  If oj Is Nothing Then
    Exit Function
  End If

End Function
Private Function ClipboardTitle()
  ClipboardTitle = "[B" & ChrW(7897) & " nh" & ChrW(7899) & " t" & ChrW(7841) & "m]"
End Function
Private Sub ZaloAppSearchAndSend_test()
  ZaloAppSearchAndSend "0934847608"
  ZaloSendAll_click_
  Exit Sub
  If Not ZaloAppGotoContact Then
    Exit Sub
  End If
  Debug.Print ZaloAppSearchAndSend(". M" & ChrW(7865), "hello")
  Debug.Print ZaloAppSearchAndSend("Cloud c" & ChrW(7911) & "a t" & ChrW(244) & "i", "hello")
End Sub
Public Function ZaloAppSearchAndSend(Search$, Optional message$, Optional files, Optional oversend As Boolean, Optional recipient$, Optional status$, Optional time As Date) As Boolean
  time = 0: status = "": recipient = ""
  On Error Resume Next
  Dim t As Double: t = Timer
  status = "null"
  Dim ct%, ct2%, o, o2, o3, o4, o5, o6, sp, cs, k%, n$, s$, url, i, a() As String

  GoSub first


  GoSub Search



  k = 0
  Do
    Set o2 = o.FindElementByXPath(".//*[contains(@class, 'loadingMediaText')]", 200, False)
    If o2 Is Nothing Then
      Exit Do
    End If
    k = k + 1: If k > 20 Then Exit Do
    DoEvents
  Loop
  Set o2 = o.FindElementByXPath(".//*[contains(@class, 'global-search-no-result')]", 200, False)
  If Not o2 Is Nothing Then
    Exit Function
  End If

  Set cs = o.FindElementsByXPath("//*[contains(@id, 'friend-item-')]", , 1500)
  For Each o In cs
    s = "": n = ""
    s = o.FindElementByXPath(".//*[contains(@class, 'friend_online_status')]/*[@class='txt-highlight']", 1500, False).Attribute("innerText")
    n = o.FindElementByXPath(".//*[contains(@class, 'conv-item-title__name')]", 1500, False).Attribute("innerText")
    n = Replace(n, Chr(160), " ")
    If s = Search Or "0" & s = Search Or "84" & s = Search Or n = Search Then
      url = o.FindElementByClass("zl-avatar__photo", 1500, False).Attribute("src")
      recipient = n
      o.Click
      GoSub header
      GoSub sent
    End If
    DoEvents
  Next
Exit Function
header:
  k = 0
  Do
    Set o2 = Selen.FindElementById("chatViewContainer", 200, False)
    If Not o2 Is Nothing Then
      Exit Do
    End If
    k = k + 1: If k > 20 Then Exit Do
    DoEvents
  Loop
  If o2 Is Nothing Then
    Exit Function
  End If
Return

sent:
  Set o3 = Selen.FindElementById("chatInput", 1500, False)
  If Not oversend Then
    Set cs = o3.FindElementsByXPath("//*[contains(@id, 'mtc-')]", , 1500)
    If Not cs Is Nothing Then
      For Each o2 In cs
        If o2.Attribute("innerText") = message Then
          status = ChrW(272) & ChrW(227) & " g" & ChrW(7917) & "i"
          ZaloAppSearchAndSend = True
          Exit Function
        End If
        DoEvents
      Next
    End If
  End If

  If IsArray(files) Then
    If files(LBound(files)) Like "*//'[[]*" Then
      ct2 = 0
      For Each i In files
        Debug.Print "Objects: "; i
        If i Like "*//'[[]*" Then
          VBA.Err.clear
          sp = Split(i, "//'[")
          Set o4 = nameToObject("'[" & sp(2))
          Set o6 = o4.DrawingObjects(sp(0))
          If VBA.Err = 0 Then
            For Each o5 In o4.DrawingObjects
              If o5.Name = o6.Name And (CStr(o5.ShapeRange.ZOrderPosition) = sp(1) Or sp(1) = "0") Then
                ct2 = ct2 + 1
                GoSub copyPaste
                GoSub mutilple

                Exit For
              End If
            Next
          End If
        End If
      Next

    ElseIf files(LBound(files)) Like "*!*$*" Then
      ct2 = 0
      For Each i In files
        If i Like "*!*$*" Then
          VBA.Err.clear
          Set o6 = Application.Evaluate(i)
          ct2 = ct2 + 1
          GoSub copyPaste
          GoSub mutilple
        End If
      Next
    Else
      a = files
      If ClipboardCopyFiles(a) Then
        ct2 = 1
        GoSub Paste
        GoSub mutilple
        GoSub release
      End If
    End If
  Else
    If files Like "[[]*[]]" Then
      ct2 = 1
      GoSub Paste
      GoSub mutilple
      GoSub release
    End If

  End If

  If message <> Empty Then
    o3.FindElementById("richInput", 1500, False).SendKeys message
  End If

  o3.FindElementById("richInput", 1500, False).SendKeys SelenKeys.return
  status = "ok"
  time = VBA.Now
  ZaloAppSearchAndSend = True
  recipient = ChrW(272) & ChrW(7871) & "n: " & n & " (" & Round(Timer - t, 3) & "s)"
  Exit Function
Return
first:
  Set o2 = Selen.FindElementById("chatViewContainer", 200, False)
  If Not o2 Is Nothing Then
    Set o3 = o2.FindElementByXPath("//*[contains(@class, 'title header-title')]", 200, False)
    n = Replace(o3.Attribute("innerText"), Chr(160), " ")
    If n = Search Then
      GoSub sent
    End If
  End If
Return
Search:
  Set cs = Selen.FindElementById("contact-search-input", 200, False)
  If cs Is Nothing Then
    Exit Function
  End If
  cs.clear
  cs.SendKeys Search
  k = 0
  Do
    Set o = Selen.FindElementByClass("ReactVirtualized__Grid__innerScrollContainer", 200, False)
    If Not o Is Nothing Then
      Exit Do
    End If
    k = k + 1: If k > 20 Then Exit Do
    DoEvents
  Loop
  If o Is Nothing Then
    Exit Function
  End If
Return
mutilple:
  k = 0: ct = 0
  Do
    Set o2 = o3.FindElementByXPath(".//*[contains(@class, 'mutilple-select-count')]", 200, False)
    If Not o2 Is Nothing Then
      ct = o2.Attribute("innerText")
      If ct >= ct2 Then
        Exit Do
      End If
    End If
    k = k + 1: If k > 20 Then Exit Do
    DoEvents
  Loop

Return
release:
  Call EmptyClipboard
Return
copyPaste:
  VBA.Err.clear
  GoSub release
  o6.CopyPicture 1, 2
  If VBA.Err = 0 Then
    Delay 300
    GoSub Paste
    GoSub release
  End If
Return
Paste:
  o3.FindElementById("richInput", 1500, False).SendKeys SelenKeys.Control + "v"
Return
End Function

Public Function SelenKeys()
  Set SelenKeys = VBA.CreateObject("Selenium.Keys")
End Function

Public Function chromedriver()
  chromedriver = Environ("LocalAppData") & "\SeleniumBasic\chromedriver.exe"
End Function

Function UserDataDirForGame()
  UserDataDirForGame = IIf(Environ$("tmp") <> n_, Environ$("tmp"), Environ$("temp")) & "\remote-Zalo"
End Function

Function SEConnectChrome(Optional ByRef driver As Object, Optional ByRef boolStart As Boolean, Optional ByVal url As String, Optional ByVal newWindow As Boolean, Optional ByVal headless As Boolean, Optional ByVal indexPage As Integer = 1, Optional ByRef boolWait As Boolean, Optional ByRef checkUrl As Boolean, Optional ByVal startIfExists As Boolean, Optional ByVal refreshIfExists As Boolean, Optional ByVal startNewTab As Boolean, Optional ByVal boolApp As Boolean, Optional ByVal popupBlocking As Boolean, Optional ByVal visible = vbNormalFocus, Optional ByVal maximize As Boolean, Optional ByVal position As String = "", Optional ByVal screen As String = "", Optional ByVal disableGPU As Boolean = True, Optional ByVal boolClose As Boolean, Optional ByVal chromePath As String, Optional ByVal Port As Long = MAINPORT9515, Optional ByVal UserDataDir As String, Optional ByVal PrivateMode As Boolean) As Boolean
  DoEvents 'Open
  If chromePath = "" Then
    chromePath = getChromePath
  End If
  Dim Win, process, isBrowserOpen As Boolean, isOpen As Boolean, s$, k%, i%
  If Port <= 0 Then Port = MAINPORT9515
  If UserDataDir = n_ Then
    UserDataDir = UserDataDirForGame
  End If
  GoSub CheckCR: isOpen = isBrowserOpen
  If Not isOpen Then
    If Not boolStart And Not startIfExists Then GoTo Ends
    Dim CmdLn$
    CmdLn = IIf(Port > 0, "--remote-debugging-port=" & Port, "") & IIf(headless, " --headless", "") & IIf(UserDataDir <> n_, " --user-data-dir=""" & UserDataDir & """", "") & " --lang=vi" & IIf(url = "", "", IIf(boolApp, " --app=", " ")) & """" & url & """" & IIf(maximize And visible <> 0, " --start-maximized", "") & IIf(position <> n_ And Not maximize, " --window-position=" & position, "") & IIf(screen <> n_ And Not maximize, " --window-size=" & screen, "") & IIf(disableGPU, " --disable-gpu", "") & IIf(PrivateMode, " --incognito", "") & IIf(popupBlocking, "", " --disable-popup-blocking")
    Shell "cmd.exe /s /k start """" """ & chromePath & """" & CmdLn, vbHide
    Shell """" & chromePath & """" & CmdLn, vbHide
    Do Until isBrowserOpen:
      GoSub CheckCR:
      DoEvents
      Delay 200
      k = k + 1: If k > 12 Then GoTo Ends
    Loop
  End If
  If driver Is Nothing Then
    If isBrowserOpen Then
      Set driver = VBA.CreateObject("selenium.ChromeDriver")
      driver.SetCapability "debuggerAddress", "127.0.0.1:" & Port
      driver.Start "chrome"
      driver.Timeouts.ImplicitWait = 5000
      driver.Timeouts.PageLoad = 5000
      driver.Timeouts.Server = 10000
    End If
  End If
  GoSub checkUrl
Ends:
SEConnectChrome = isBrowserOpen
Set process = Nothing: Set Win = Nothing
Exit Function
checkUrl:
  If driver Is Nothing Then Return
  If Not checkUrl Then Return
  k = 0
  On Error Resume Next
  checkUrl = False
  GoSub compareUrl
  For Each Win In driver.Windows
    DoEvents
    If Err.Number <> 0 Then Return
    Win.Activate
    GoSub compareUrl
  Next
nextCheckUrl:
  If Not checkUrl Or (startIfExists And checkUrl) Then
    If startNewTab Then
      driver.ExecuteScript "window.open(arguments[0], '_blank');", url
    Else
      driver.Get url
    End If
  End If
  On Error GoTo 0
Return
compareUrl:
  If LCase$(driver.url) Like "*" & LCase$(url) & "*" Then
    k = k + 1: If indexPage = k Then checkUrl = True: GoTo nextCheckUrl
  End If
Return

CheckCR:
  On Error Resume Next
  isBrowserOpen = False
  For Each process In VBA.GetObject("winmgmts:\\.\root\CIMV2") _
    .ExecQuery("SELECT * FROM Win32_Process WHERE Name = ""chrome.exe""", , 48)
    DoEvents
    s = LCase$(process.commandline)
    If s Like LCase$("*chrome*--remote-debugging-port=" & Port & " *") Then
      isBrowserOpen = True: If boolClose Then process.Terminate: Set driver = Nothing: GoTo Ends
      Exit For
    End If
  Next
Return
End Function




Function glbShellA()
  Set glbShellA = VBA.CreateObject("Shell.Application")
End Function
Function glbHTMLFile()
 Set glbHTMLFile = VBA.CreateObject("HTMLFile")
End Function
Function glbXHR()
 Set glbXHR = VBA.CreateObject("Microsoft.XMLHTTP")
End Function
Private Sub getChromeLastVesion_test()
  Debug.Print getChromeLastVesion()
End Sub

Function getChromeLastVesion()
  Dim h, t, s$, w$
  #If VBA7 And Win64 Then
    w = "win64"
  #Else
    w = "win"
  #End If
  GoSub http
  If t = n_ Then
    Exit Function
  End If
  Set h = glbHTMLFile
  s = "function jsGetChromeLastVersion(os){" & vbLf
  s = s & "  var a = " & t & ";" & vbLf
  s = s & "  for (var i =0; i < a.length; i++) {" & vbLf
  s = s & "    if ( a[i].os == os ) {" & vbLf
  s = s & "      var b = a[i].versions;" & vbLf
  s = s & "      for (var j=0; j < b.length; j++) {" & vbLf
  s = s & "        if ( b[j].channel == 'stable' ) {" & vbLf
  s = s & "          return b[j].previous_version + '\n' + b[j].current_version;" & vbLf
  s = s & "        }" & vbLf
  s = s & "      }" & vbLf
  s = s & "      break;" & vbLf
  s = s & "    }" & vbLf
  s = s & "  }" & vbLf
  s = s & "}"

  h.parentWindow.execScript s, "javascript"

  getChromeLastVesion = h.parentWindow.jsGetChromeLastVersion(w)

Exit Function
http:
  Set h = glbXHR
  With h
    .Open "GET", "https://omahaproxy.appspot.com/all.json", False
    .setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/83.0.4103.97 Safari/537.36"
    .setRequestHeader "Content-type", "application/x-www-form-urlencoded"
    .Send
    t = .responseText
  End With
  Return
End Function

Function isChromeLastVesion() As Boolean
  Dim c$, f
  c = getChromePath
  Set f = glbFSO
  If Not f.FileExists(c) Then Exit Function
  c = f.GetFileVersion(c)
  If getChromeLastVesion Like "*" & VBA.vbLf & c Then
    isChromeLastVesion = True
  End If
End Function



Private Function UpdateChromedriver(Optional ByVal chromePath As String = "", Optional ByVal ReUpdate As Boolean) As Boolean
  If chromePath = "" Then
    chromePath = getChromePath
  End If
  On Error Resume Next
  Dim LastedUpdate As String
  Dim FSO As Object
  Set FSO = glbFSO
  Dim XMLHTTP As Object
  Dim a, Tmp1$, tmp$, eURL$, temp$, info$, sb$, sb2$
  sb = Environ("LOCALAPPDATA") & "\SeleniumBasic"
  sb2 = "C:\DevPrograms\Selenium"
  Const LATEST_RELEASE = "https://chromedriver.storage.googleapis.com/LATEST_RELEASE"
  Const url$ = "https://chromedriver.storage.googleapis.com/"

  Const EXE$ = "\chromedriver.exe"
  Const ZIP$ = "\chromedriver_win32.zip"
  temp = Environ("TEMP"): GoSub DelTemp
  If Not FSO.FileExists(chromePath) Then Exit Function
  info = FSO.GetFileVersion(chromePath)
  eURL = "https://chromedriver.storage.googleapis.com/LATEST_RELEASE_" & Split(info, ".")(0)
  GoSub http
  LastedUpdate = VBA.GetSetting("Chromedriver", "Update", "Last")

  If LastedUpdate < tmp Or ReUpdate Then
    GoSub Download
  Else
    UpdateChromedriver = True
  End If
Ends: Set FSO = Nothing
Exit Function
Download:
On Error Resume Next
  ChromeDriverCloseAll
  eURL = url$ & tmp & "/chromedriver_win32.zip"
  If ToFile(eURL, temp & ZIP) Then
    GoSub Extract
    Call VBA.SaveSetting("Chromedriver", "Update", "Last", tmp)
  End If
On Error GoTo 0
Return
Extract:
On Error Resume Next
With glbShellA
  .Namespace(temp & "\").CopyHere .Namespace(temp & ZIP).items
End With
With FSO
  If .FileExists(temp & EXE) Then
    If .folderexists(sb2) Then FSO.CopyFile temp & EXE, sb2 & EXE, True
    If .folderexists(sb) Then FSO.CopyFile temp & EXE, sb & EXE, True
  End If
  UpdateChromedriver = Err.Number = 0
End With
On Error GoTo 0
GoSub DelTemp
Return

DelTemp:
On Error Resume Next
  FSO.DeleteFile temp & ZIP
  FSO.DeleteFile temp & EXE
On Error GoTo 0
Return
http:
Set XMLHTTP = glbXHR
With XMLHTTP
  .Open "GET", eURL, False
  .setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/83.0.4103.97 Safari/537.36"
  .setRequestHeader "Content-type", "application/x-www-form-urlencoded"
  .Send
  tmp = VBA.Trim(Application.Clean(.responseText))
End With
Return
End Function
Sub ChromeDriverCloseAll()
  On Error Resume Next
  Dim p
  For Each p In VBA.GetObject("winmgmts:\\.\root\CIMV2") _
        .ExecQuery("SELECT * FROM Win32_Process WHERE Name = ""chromedriver.exe""", , 48)
    p.Terminate
  Next
End Sub
Function ToFile(ByVal szURL As String, ByVal szFileName As String) As Boolean
  Dim h
  Set h = glbXHR
  h.Open "GET", szURL, False: h.Send
  If h.status <> 200 Then Exit Function
  With VBA.CreateObject("ADODB.Stream")
    .Open
    .Type = 1
    .Write h.responseBody
    .SaveToFile szFileName, 2
    .Close
  End With
  ToFile = True
End Function

Function SwitchIFrame(ByVal IFrame, ByVal driver As Object, Optional ByVal Timeout% = -1, Optional ByVal hRaise As Boolean = True) As Boolean
  On Error Resume Next
  driver.SwitchToFrame IFrame, Timeout, hRaise
  SwitchIFrame = Err.Number = 0
  On Error GoTo 0
End Function


Sub SEChromeClose(Optional ByRef driver As Object)
  Dim i, p
  On Error Resume Next
  For i = driver.Windows.Count To 1 Step -1
    driver.Windows(i).Close
  Next:
  Set driver = Nothing
  For Each p In VBA.GetObject("winmgmts:\\.\root\CIMV2") _
  .ExecQuery("SELECT * FROM Win32_Process WHERE (Name = ""chromedriver.exe"" Or Name = ""chrome.exe"")", , 48)
    Select Case p.Name
    Case "chrome.exe"
      If p.commandline Like "*chrome.exe*--remote-debugging-port=*--user-data-dir=*" Then
        p.Terminate
      End If
    Case "chromedriver.exe": p.Terminate
    End Select
  Next

  On Error GoTo 0
End Sub




Function HTMLNode(ByVal obj As Object, ParamArray ArrayNodes()) As Object
  DoEvents 'Open
  Dim aNode
  On Error Resume Next
  For Each aNode In ArrayNodes
    If obj.ChildNodes(aNode) Is Nothing Then Exit Function
    Set obj = obj.ChildNodes(aNode)
    If Err.Number <> 0 Then Exit Function
  Next
  Set HTMLNode = obj
End Function





Sub toggleChromeVisible()

End Sub





#If VBA7 And Win64 Then
Function BringWindowToFront(ByVal hwnd As LongPtr) As Boolean
#Else
Function BringWindowToFront(ByVal hwnd As Long) As Boolean
#End If

  Dim ThreadID1 As Long, ThreadID2 As Long, nRet As Long
  On Error Resume Next
  If hwnd = GetForegroundWindow() Then
    BringWindowToFront = True
  Else
    ThreadID1 = GetWindowThreadProcessId(GetForegroundWindow, ByVal 0&)
    ThreadID2 = GetWindowThreadProcessId(hwnd, ByVal 0&)
    Call AttachThreadInput(ThreadID1, ThreadID2, True)
    nRet = SetForegroundWindow(hwnd)
    If IsIconic(hwnd) Then
      Call ShowWindow(hwnd, 9) ' SW_RESTORE)
      Call ShowWindow(hwnd, 5) 'SW_SHOW)
    Else
      Call ShowWindow(hwnd, 1) 'SW_SHOW 5)
    End If
    BringWindowToFront = CBool(nRet)
    Call AttachThreadInput(ThreadID1, ThreadID2, False)
  End If
End Function
#If VBA7 And Win64 Then
Public Function GetChromeHandleByProcessID(Optional ByVal Port$ = MAINPORT9515, Optional ByVal UserDataDir$) As LongPtr
#Else
Public Function GetChromeHandleByProcessID(Optional ByVal Port$ = MAINPORT9515, Optional ByVal UserDataDir$) As Long
#End If
  Dim p
  For Each p In VBA.GetObject("winmgmts:\\.\root\CIMV2") _
  .ExecQuery("SELECT * FROM Win32_Process WHERE Name = ""chrome.exe""", , 48)
    If LCase$(p.commandline) Like LCase$("*chrome*--remote-debugging-port=" & Port & "*--user-data-dir=*" & LCase$(UserDataDir) & "*") Then
      GetChromeHandleByProcessID = InstanceToWnd(p.Processid, "Chrome_WidgetWin_1")
      Exit For
    End If
  Next
End Function
#If VBA7 And Win64 Then
Public Function GetZaloHandle() As LongPtr
#Else
Public Function GetZaloHandle() As Long
#End If
  Dim p
  GetZaloHandle = FindWindowEx(FindWindow("Chrome_WidgetWin_1", "Zalo"), 0&, "Chrome_RenderWidgetHostHWND", "Chrome Legacy Window")
End Function
#If VBA7 And Win64 Then
Function InstanceToWnd(ByVal target_pid As Long, Optional ByVal title$, Optional ByVal class$) As LongPtr
#Else
Function InstanceToWnd(ByVal target_pid As Long, Optional ByVal title$, Optional ByVal class$) As Long
#End If
  #If VBA7 And Win64 Then
    Dim hwnd As LongPtr
  #Else
    Dim hwnd As Long
  #End If
  Dim pid As Long, thread_id As Long
  hwnd = FindWindow(class, title)
  Do While hwnd <> 0
    If GetParent(hwnd) = 0 Then
      thread_id = GetWindowThreadProcessId(hwnd, pid)
      If pid = target_pid Then
        InstanceToWnd = hwnd
        Exit Do
      End If
    End If
    hwnd = GetWindow(hwnd, 2)
  Loop
End Function
Public Function ZaloAppGotoContact() As Boolean
  If Not ZaloAppLogin Then
    Exit Function
  End If
  Dim o, cs
  Set cs = Selen.FindElementsByXPath("//*[@class='nav__tabs__top']/child::*", , 200)
  If cs Is Nothing Then
    Exit Function
  End If
  For Each o In cs
    If o.Attribute("data-id") = "div_Main_TabCT" Then
      o.Click
      ZaloAppGotoContact = True
      Exit For
    End If
  Next
End Function


Public Function ZaloAppLogin() As Boolean
  If Not Selen Is Nothing Then
    ZaloAppLogin = True
    Exit Function
  End If
  On Error Resume Next
  Dim k&, o As Object, e As Boolean: e = True
  SEConnectChrome driver:=Selen, boolStart:=True, boolApp:=True, url:=ZaloAppSite, checkUrl:=e
  If e Then
    Set o = Selen.FindElementByClass("zl-avatar__photo", 200, False)
    If o Is Nothing Then
      SEConnectChrome driver:=Selen, boolStart:=True, boolApp:=True, url:=ZaloAppSite2, refreshIfExists:=True
    Else
      ZaloAppLogin = True
    End If
  End If
    Selen.Window.SetPosition 1700, 0
    Selen.Window.SetSize 300, 1060
End Function

Sub ClipboardText()
 AppendObjectsToList ClipboardTitle, 4, True
End Sub

Sub AttachObjects()
  PickObjects 1
End Sub
Sub AttachFiles()
  PickObjects 3
End Sub

Sub PickObjects_test()
  PickObjects 1
End Sub
Sub PickObjects(Optional ByVal style% = 0)
  On Error GoTo e
  Dim k&, a, ws, ws2, sp1, sp2, o, o2, r, rg, rg1, rg2, rg3, rg4, s$, cap1$, cap2$
  Dim os As Boolean, nlist$
  Static list$, style_%
  Set ws = shZaloExcel

  If style = 3 Then
    a = DialogExplorer(MultiSelect:=True)
    If IsArray(a) Then
      list = """" & Join(a, """,""") & """"
      Call AppendObjectsToList(list, style, True)
    End If
  Else
    Set a = Selection

    Select Case TypeName(a)
    Case "Nothing": Exit Sub
    Case "DrawingObjects": list = "": style = 2: style_ = 0: For Each o In a: GoSub r: Next: GoSub ch
    Case "Range":
      If style_ = 2 Then
        GoSub ch
      End If
      list = ""
      On Error Resume Next
      Set a = Application.InputBox("?", Type:=8)
      On Error GoTo 0
      If TypeName(a) = "Range" Then
        For Each rg4 In a.Areas
           list = list & IIf(list = Empty, "", ",") & """" & rg4.Address(1, 0, external:=1) & """"
           nlist = nlist & IIf(nlist = Empty, "", ",") & rg4.Address(0, 0)
        Next
        os = AppendObjectsToList(list, style, False)
        If Not os Then
          GoSub ch
        End If
      End If

    Case Else '"Rectangle", "GroupObject", "Picture", "ChartArea"
      list = "": style = 2: style_ = 0: Set o = a: GoSub r: GoSub ch
    End Select
  End If
e:
Exit Sub

ch:
  Set sp1 = ws.Shapes("btnObjects")
  cap1 = ChrW(212) & "/" & ChrW(272) & "T"
  If sp1.TextFrame2.TextRange.Text = "D?n" And (style_ = 1 Or style_ = 2) Then
    style = style_
    Call AppendObjectsToList(list, style, True)
    sp1.TextFrame2.TextRange.Text = cap1
    Exit Sub
  Else
    If style = 1 Or style = 2 Then
      style_ = style
      sp1.TextFrame2.TextRange.Text = "D?n"
      Alert Timeout:=0, title:=ChrW(272) & ChrW(227) & " ch" & ChrW(233) & "p t" & ChrW(234) & "n v" & ChrW(224) & "o b" & ChrW(7897) & " nh" & ChrW(7899) & " t" & ChrW(7841) & "m th" & ChrW(7901) & "i", Prompt:= _
           ChrW(272) & ChrW(7889) & "i t" & ChrW(432) & ChrW(7907) & "ng: " & nlist & vbLf & vbLf & _
           "(" & ChrW(272) & ChrW(7871) & "n Trang [Zalo Excel] " & "ch" & ChrW(7885) & "n " & ChrW(244) & " c" & ChrW(7897) & "t D " & ChrW(273) & ChrW(7875) & " d" & ChrW(225) & "n)" & vbLf & _
           "Nh" & ChrW(7845) & "n [D" & ChrW(225) & "n] ho" & ChrW(7863) & "c ph" & ChrW(237) & "m t" & ChrW(7855) & "t Ctrl+Shift+X"
    End If
  End If
Return
r:
  k = 0
  Select Case TypeName(o.Parent)
  Case "Worksheet": k = o.ShapeRange.ZOrderPosition
      s = o.Parent.Range("A1").Address(external:=1)
      s = Split(s, "!$")(0)
  Case "Chart":
    s = "'[" & ActiveSheet.Parent.Name & "]" & ActiveSheet.Name & "'"
  Case Else
    Return
  End Select
  nlist = nlist & IIf(nlist = Empty, "", ",") & o.Name
  list = list & IIf(list = Empty, "", ",") & """" & o.Name & "//'[" & k & "//" & s & """"
Return
End Sub

Private Function AppendObjectsToList(list$, Optional ByVal style% = 0, Optional ByVal outRange As Boolean) As Boolean
  If list = Empty Then
    Exit Function
  End If
  Dim k&, ws, rg, rg1, rg2, rg3, s$
  Dim os As Boolean
  Set ws = shZaloExcel
  Set rg = ws.Range("D4")
  Set rg2 = Selection

  If TypeName(rg2) <> "Range" Then
    If Not outRange Then Exit Function
    Set rg2 = rg(10000, 1).End(3).Offset(1)
  End If
  If Not rg2.Parent Is ws Then
    Exit Function
  End If

  For Each rg3 In rg2.Areas
    For Each rg1 In rg3
      If Not Intersect(rg.Resize(10000), rg1) Is Nothing Then
        s = rg1.Value
        If s = Empty _
        Or (style = 1 And s Like "*$*") _
        Or (style = 2 And s Like "*//'[[]*") _
        Or (style = 3 And Not s Like "*//'[[]*" And Not s Like "*$*") _
        Or (style = 4 And s = Empty) Then
          rg1.Value = s & IIf(s = Empty, "", ",") & list
          Select Case style
          Case 1: rg1.Font.Color = vbWhite
          Case 2: rg1.Font.Color = vbYellow
          Case 3: rg1.Font.Color = vbGreen
          Case 4: rg1.Font.Color = vbCyan
          End Select
        End If
        os = True
      End If
    Next
  Next
  list = ""
  AppendObjectsToList = True
End Function



Private Function nameToObject(SheetName$) As Object
  Dim s, b, ws As Object
  On Error Resume Next
  s = SheetName
  If s Like "'*'" Then
    s = Mid(s, 3, Len(s) - 3)
    b = Split(s, "]")
    Set ws = Workbooks(b(0))
    If Not ws Is Nothing Then
      Set nameToObject = ws.Worksheets(b(1))
    End If
  Else
    Set nameToObject = ActiveWorkbook.Worksheets(SheetName)
  End If
  On Error GoTo 0
End Function

Sub SetOnKey(ByVal OnAction$, ByVal keys$, Optional ByVal ClearKey As Boolean)
  On Error Resume Next
  Application.OnKey keys
  If Not ClearKey Then
    If OnAction Like "*)" Then
      Application.OnKey keys, "'" & ThisWorkbook.Name & "'!" & OnAction
    Else
      Application.OnKey keys, "'" & ThisWorkbook.Name & "'!'" & OnAction & "'"
    End If
  End If
End Sub

Sub projectOnKey(Optional ClearKey As Boolean)
  On Error Resume Next
  SetOnKey "AttachObjects", "^+x", ClearKey
End Sub



Function TextToClipBoard(ByVal Text As String) As String
  #If Mac Then
    With New MSForms.DataObject
      .SetText Text: .PutInClipboard
    End With
  #Else
    #If VBA7 Then
      Dim hGlobalMemory     As LongPtr
      Dim hClipMemory       As LongPtr
      Dim lpGlobalMemory    As LongPtr
    #Else
      Dim hGlobalMemory     As Long
      Dim hClipMemory       As Long
      Dim lpGlobalMemory    As Long
    #End If
    Dim X                     As Long
    hGlobalMemory = GlobalAlloc(&H42, Len(Text) + 1)
    lpGlobalMemory = GlobalLock(hGlobalMemory)
    lpGlobalMemory = lstrcpy(lpGlobalMemory, Text)
    If GlobalUnlock(hGlobalMemory) <> 0 Then
      TextToClipBoard = "Could not unlock memory location. Copy aborted."
      GoTo PrepareToClose
    End If
    If OpenClipboard(0&) = 0 Then
      TextToClipBoard = "Could not open the Clipboard. Copy aborted."
      Exit Function
    End If
    X = EmptyClipboard()
    hClipMemory = SetClipboardData(1, hGlobalMemory)
PrepareToClose:
    If CloseClipboard() = 0 Then
      TextToClipBoard = "Could not close Clipboard."
    End If
  #End If
End Function
Public Function ClipboardCopyFiles(files() As String) As Boolean
  Dim data As String
  Dim df As DROPFILES
  #If VBA7 And Win64 Then
  Dim hGlobal As LongPtr
  Dim lpGlobal As LongPtr
  #Else
  Dim hGlobal As Long
  Dim lpGlobal As Long
  #End If
  Dim i As Long

  If OpenClipboard(0&) Then
    Call EmptyClipboard
    For i = LBound(files) To UBound(files)
      data = data & files(i) & vbNullChar
    Next
    data = data & vbNullChar
    hGlobal = GlobalAlloc(GHND, Len(df) + Len(data))
    If hGlobal Then
      lpGlobal = GlobalLock(hGlobal)
      df.pFiles = Len(df)
      Call CopyMem(ByVal lpGlobal, df, Len(df))
      Call CopyMem(ByVal (lpGlobal + Len(df)), ByVal data, Len(data))
      Call GlobalUnlock(hGlobal)
      If SetClipboardData(CF_HDROP, hGlobal) Then
        ClipboardCopyFiles = True
      End If
    End If
    Call CloseClipboard
  End If
End Function
Public Function ClipboardPasteFiles(files() As String) As Long
  Dim hDrop As Long
  Dim nFiles As Long
  Dim i As Long
  Dim desc As String
  Dim filename As String
  Dim pt As POINTAPI
  Const MAX_PATH As Long = 260
  If IsClipboardFormatAvailable(CF_HDROP) Then
    If OpenClipboard(0&) Then
      hDrop = GetClipboardData(CF_HDROP)
      nFiles = DragQueryFile(hDrop, -1&, "", 0)
      ReDim files(0 To nFiles - 1) As String
      filename = Space(MAX_PATH)
      For i = 0 To nFiles - 1
        Call DragQueryFile(hDrop, i, filename, Len(filename))
        files(i) = TrimNull(filename)
      Next
      Call CloseClipboard
    End If
    ClipboardPasteFiles = nFiles
  End If
End Function


Public Function ClipboardPasteReletang(files() As String) As Long
Const CF_TEXT = 1
Const CF_BITMAP = 2
Const CF_METAFILEPICT = 3
Const CF_SYLK = 4
Const CF_DIF = 5
Const CF_TIFF = 6
Const CF_OEMTEXT = 7
Const CF_DIB = 8
Const CF_PALETTE = 9
Const CF_PENDATA = 10
Const CF_RIFF = 11
Const CF_WAVE = 12
Const CF_UNICODETEXT = 13
Const CF_ENHMETAFILE = 14
Const CF_HDROP = 15
Const CF_LOCALE = 16
Const CF_MAX = 17
  Dim hDrop As Long
  Dim nFiles As Long
  Dim i As Long
  Dim desc As String
  Dim filename As String
  Dim pt As POINTAPI
  Const MAX_PATH As Long = 260
  If IsClipboardFormatAvailable(CF_HDROP) Then
    If OpenClipboard(0&) Then
      hDrop = GetClipboardData(CF_HDROP)
      nFiles = DragQueryFile(hDrop, -1&, "", 0)
      ReDim files(0 To nFiles - 1) As String
      filename = Space(MAX_PATH)
      For i = 0 To nFiles - 1
        Call DragQueryFile(hDrop, i, filename, Len(filename))
        files(i) = TrimNull(filename)
      Next
      Call CloseClipboard
    End If
    ClipboardPasteReletang = nFiles
  End If
End Function

Private Function TrimNull(ByVal sTmp As String) As String
  Dim nNul As Long
  nNul = InStr(sTmp, vbNullChar)
  Select Case nNul
  Case Is > 1
    TrimNull = Left(sTmp, nNul - 1)
  Case 1
    TrimNull = ""
  Case 0
    TrimNull = Trim(sTmp)
  End Select
End Function

Private Sub clickError()
  Alert "B" & ChrW(7841) & "n ch" & ChrW(432) & "a t" & ChrW(7843) & "i m" & ChrW(227) & " th" & ChrW(7921) & "c thi!" & vbLf & _
  "Click n" & ChrW(250) & "t T" & ChrW(7843) & "i m" & ChrW(227) & ", sao ch" & ChrW(233) & "p m" & ChrW(227) & " v" & ChrW(224) & "o Module zzzzzzzZaloController!"
End Sub

Private Sub Login_click()
  On Error Resume Next
  Application.Run "Login_click_"
  If VBA.Err Then clickError
End Sub
Private Sub ZaloContact_click()
  On Error Resume Next
  Application.Run "ZaloContact_click_"
  If VBA.Err Then clickError
End Sub

Sub ZaloSend_click()
  On Error Resume Next
  Application.Run "ZaloSend_click_"
  If VBA.Err Then clickError
End Sub
Private Sub ZaloSendAll_click()
  On Error Resume Next
  Application.Run "ZaloSendAll_click_"
  If VBA.Err Then clickError
End Sub


Function glbFSO()
 Set glbFSO = VBA.CreateObject("Scripting.FileSystemObject")
End Function
Sub btnUpdateChromedriver()
  Dim b As Boolean
  On Error Resume Next
  b = Application.Run("UpdateChromedriver", "", True)
  If VBA.Err Then
    Alert "B" & ChrW(7841) & "n ch" & ChrW(432) & "a sao ch" & ChrW(233) & "p m" & ChrW(227) & " th" & ChrW(7921) & "c thi v" & ChrW(224) & "o Module!"
  Else
    Alert "UpdateChromedriver: " & IIf(b, ChrW(272) & ChrW(227) & " c" & ChrW(7853) & "p nh" & ChrW(7853) & "t", "Failed"), Timeout:=5
  End If
End Sub
Private Sub downloadSeleniumBasic()
  Dim p
  p = Environ("LocalAppData") & "\SeleniumBasic"
  If Not glbFSO.folderexists(p) Then
    ThisWorkbook.FollowHyperlink Si_SeleniumBasic, , True
  Else
    Alert "Selenium: " & ChrW(272) & ChrW(227) & " c" & ChrW(224) & "i " & ChrW(273) & ChrW(7863) & "t", Timeout:=5
  End If
End Sub
Private Sub ChromeClose()
  On Error Resume Next
  Application.Run "SEChromeClose", Selen
  If VBA.Err Then
    Alert "B" & ChrW(7841) & "n ch" & ChrW(432) & "a th" & ChrW(234) & "m m" & ChrW(227) & " v" & ChrW(224) & "o Module!"
  Else
    Alert ChrW(272) & ChrW(227) & " " & ChrW(273) & ChrW(243) & "ng Chrome"
  End If
End Sub
Private Sub downloadChrome()
  Dim p, b As Boolean
  p = getChromePath
  If Not glbFSO.folderexists(p) Then
    ThisWorkbook.FollowHyperlink Si_Chrome, , True
  Else
    On Error Resume Next
    b = Application.Run("isChromeLastVesion")
    If VBA.Err Then
      Alert "B" & ChrW(7841) & "n ch" & ChrW(432) & "a th" & ChrW(234) & "m m" & ChrW(227) & " v" & ChrW(224) & "o Module!"
    Else
      If b Then
        btnUpdateChromedriver
      Else
        OpenURL getChromePath
      End If
    End If
  End If
End Sub

Sub Delay(Optional ByVal MiliSecond% = 1000)
  Dim Start&, check&
  Start = getTickCount&()
  Do Until check >= Start + MiliSecond
    DoEvents
    check = getTickCount&()
  Loop
End Sub



Function getChromePath()
  getChromePath = "C:\Program Files\Google\Chrome\Application\chrome.exe"
  If Len(Dir(getChromePath)) = 0 Then
    getChromePath = "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
    If Len(Dir(getChromePath)) = 0 Then
      getChromePath = GetFolder(&H1C) & "\Google\Chrome\Application\chrome.exe"
      If Len(Dir(getChromePath)) = 0 Then
        getChromePath = n_
      End If
    End If
  End If
End Function


Function GetFolder(ByVal lngFolder&)
 Dim strPath$, strBuffer As String * 1000
 If SHGetFolderPath(0&, lngFolder, 0&, 0, strBuffer) = 0 Then
   strPath = Left$(strBuffer, InStr(strBuffer, Chr$(0)) - 1)
 Else
   strPath = n_
 End If
 GetFolder = strPath
End Function



