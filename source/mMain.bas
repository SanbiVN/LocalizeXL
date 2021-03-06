Attribute VB_Name = "mMain"
Option Explicit
Option Compare Text


Public Type Var64
#If Win64 Then
  Long As LongPtr
#Else
  Long As Long
#End If
End Type

Public Type POINTAPI
X As Long
Y As Long
End Type

Public Type RECT6
  WindowState As Long
  Areas(1 To 4) As String
  XY(3) As POINTAPI
  Left As Long
  Top As Long
  Width As Long
  Height As Long
  marginLeft As Long
  marginTop As Long
  marginRight As Long
  marginBottom As Long
End Type

Public Type RECT
Left As Long
Top As Long
Right As Long
Bottom As Long
End Type

' API '//'//'//'//'//'//'//'//'//'//'//'//'//'//'//'//'//'//'//'//'//'//'//'//'//'//'//'//'//


#If VBA7 And Win64 Then    ' 64 bit Excel under 64-bit windows
                           ' Use LongLong and LongPtr
Declare PtrSafe Function SetTimer Lib "USER32" (ByVal hwnd As LongPtr, ByVal nIDEvent As LongPtr, ByVal uElapse As LongLong, ByVal lpTimerFunc As LongPtr) As LongPtr
Declare PtrSafe Function KillTimer Lib "USER32" (ByVal hwnd As LongPtr, ByVal nIDEvent As LongPtr) As LongPtr
#ElseIf VBA7 Then     ' 64 bit Excel in all environments
                      ' Use LongPtr only, LongLong is not available
Declare PtrSafe Function SetTimer Lib "USER32" (ByVal hwnd As LongPtr, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As LongPtr) As Long
Declare PtrSafe Function KillTimer Lib "USER32" (ByVal hwnd As LongPtr, ByVal nIDEvent As Long) As Long
#Else    ' 32 bit Excel
Declare Function SetTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Declare Function KillTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long
#End If


#If VBA7 Then
#If Win64 Then
  Declare PtrSafe Function FindWindowEx Lib "USER32" Alias "FindWindowExA" (ByVal hWnd1 As LongLong, ByVal hWnd2 As LongLong, ByVal lpsz1 As String, ByVal lpsz2 As String) As LongLong
  Declare PtrSafe Function SetWindowLong Lib "USER32" Alias "SetWindowLongPtrA" (ByVal hwnd As LongLong, ByVal nIndex As Long, ByVal dwNewLong As LongLong) As LongLong
  Declare PtrSafe Function GetWindowLong Lib "USER32" Alias "GetWindowLongPtrA" (ByVal hwnd As LongLong, ByVal nIndex As Long) As LongLong
#Else
  Declare PtrSafe Function SetWindowLong Lib "USER32" Alias "SetWindowLongA" (ByVal hwnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As Long
  Declare PtrSafe Function FindWindowEx Lib "USER32" Alias "FindWindowExA" (ByVal hWnd1 As LongPtr, ByVal hWnd2 As LongPtr, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
  Declare PtrSafe Function GetWindowLong Lib "USER32" Alias "GetWindowLongA" (ByVal hwnd As LongPtr, ByVal nIndex As Long) As Long
#End If
Declare PtrSafe Function ShowWindow Lib "USER32" (ByVal hwnd As LongPtr, ByVal nCmdShow As Long) As Long
Declare PtrSafe Function GetParent Lib "user32.dll" (ByVal hwnd As LongPtr) As LongPtr
Declare PtrSafe Function GetNextWindow Lib "USER32" Alias "GetWindow" (ByVal hwnd As LongPtr, ByVal wFlag As Long) As LongPtr
Declare PtrSafe Function GetWindow Lib "USER32" (ByVal hwnd As LongPtr, ByVal wCmd As Long) As LongPtr
Declare PtrSafe Function SetParent Lib "USER32" (ByVal hWndChild As LongPtr, ByVal hWndParent As LongPtr) As Long
Declare PtrSafe Function CreatePolygonRgn Lib "gdi32" (lpPoint As POINTAPI, ByVal nCount As Long, ByVal nPolyFillMode As Long) As LongPtr
Declare PtrSafe Function MoveWindow Lib "USER32" (ByVal hwnd As LongPtr, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Declare PtrSafe Function DwmSetWindowAttribute Lib "dwmapi" (ByVal hwnd As LongPtr, ByVal attr As Integer, ByRef attrValue As Integer, ByVal attrSize As Integer) As Long
Declare PtrSafe Function DwmExtendFrameIntoClientArea Lib "dwmapi" (ByVal hwnd As LongPtr, ByRef NEWMARGINS As RECT) As Long
Declare PtrSafe Function FindWindow Lib "USER32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Declare PtrSafe Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As LongPtr
Declare PtrSafe Function SetWindowRgn Lib "USER32" (ByVal hwnd As LongPtr, ByVal hRgn As LongPtr, ByVal bRedraw As Long) As Long
Declare PtrSafe Function DrawMenuBar Lib "USER32" (ByVal hwnd As LongPtr) As Long
Declare PtrSafe Function GetDC Lib "USER32" (ByVal hwnd As LongPtr) As Long
Declare PtrSafe Function GetDeviceCaps Lib "gdi32" (ByVal hDC As LongPtr, ByVal nIndex As Long) As Long
Declare PtrSafe Function ReleaseDC Lib "USER32" (ByVal hwnd As LongPtr, ByVal hDC As LongPtr) As Long
Declare PtrSafe Function SetWindowsHookEx Lib "USER32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As LongPtr, ByVal hmod As LongPtr, ByVal dwThreadId As Long) As Long
Declare PtrSafe Function CallNextHookEx Lib "USER32" (ByVal hHook As LongPtr, ByVal ncode As Long, ByVal wParam As LongPtr, lParam As Any) As Long
Declare PtrSafe Function UnhookWindowsHookEx Lib "USER32" (ByVal hhk As LongPtr) As Long
Declare PtrSafe Function SetLayeredWindowAttributes Lib "USER32" (ByVal hwnd As LongPtr, ByVal crKey As LongPtr, ByVal bAlpha As Byte, ByVal dwFlags As LongPtr) As Long
Declare PtrSafe Sub SetWindowPos Lib "USER32" (ByVal hwnd As LongPtr, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Declare PtrSafe Function ReleaseCapture Lib "USER32" () As Long
Declare PtrSafe Function GetWindowRect Lib "USER32" (ByVal hwnd As LongPtr, lpRect As RECT) As Long
Declare PtrSafe Function DeleteObject Lib "gdi32" (ByVal hObject As LongPtr) As Long
Declare PtrSafe Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As LongPtr
Declare PtrSafe Function timeGetTime Lib "winmm.dll" () As Long
Declare PtrSafe Function ScreenToClient Lib "USER32" (ByVal hwnd As LongPtr, lpPoint As POINTAPI) As Long
Declare PtrSafe Function MapWindowPoints Lib "USER32" (ByVal hwndFrom As LongPtr, ByVal hwndTo As LongPtr, lppt As Any, ByVal cPoints As Long) As Long
Declare PtrSafe Function GetWindowText Lib "USER32" Alias "GetWindowTextA" (ByVal hwnd As LongPtr, ByVal lpString As String, ByVal cch As Long) As Long
Declare PtrSafe Function IsWindowVisible Lib "USER32" (ByVal hwnd As LongPtr) As Long
Declare PtrSafe Function LockWindowUpdate Lib "USER32" (ByVal hwndLock As LongPtr) As Long
#Else
Declare Function ShowWindow Lib "USER32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Declare Function LockWindowUpdate Lib "USER32" (ByVal hwndLock As Long) As Long
Declare Function MapWindowPoints Lib "USER32" (ByVal hwndFrom As Long, ByVal hwndTo As Long, lppt As Any, ByVal cPoints As Long) As Long
Declare Function GetParent Lib "user32.dll" (ByVal hwnd As Long) As Long
Declare Function GetNextWindow Lib "USER32" Alias "GetWindow" (ByVal hwnd As Long, ByVal wFlag As Long) As Long
Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndParent As Long) As Long
Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Declare Function CreatePolygonRgn Lib "gdi32" (lpPoint As POINTAPI, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
Declare Function AccessibleObjectFromWindow Lib "oleacc" (ByVal hWnd As Long, ByVal dwId As Long, riid As GUID, xlWB As Object) As Long
Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Declare Function DwmSetWindowAttribute Lib "dwmapi" (ByVal hWnd As Long, ByVal attr As Integer, ByRef attrValue As Integer, ByVal attrSize As Integer) As Long
Declare Function DwmExtendFrameIntoClientArea Lib "dwmapi" (ByVal hWnd As Long, ByRef NEWMARGINS As RECT) As Long
Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Long) As Long
Declare Function DrawMenuBar Lib "user32" (ByVal hWnd As Long) As Long
Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal ncode As Long, ByVal wParam As Long, lParam As Any) As Long
Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hhk As Long) As Long
Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Declare Sub SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Declare Function ReleaseCapture Lib "user32" () As Long
Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Declare Function timeGetTime Lib "winmm.dll" () As Long
Declare Function ScreenToClient Lib "USER32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Declare Function GetWindowText Lib "USER32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Declare Function IsWindowVisible Lib "user32" (ByVal hWnd As Long) As Long
#End If

'//'//'//'//'//'//'//'//'//'//'//'//'//'//'//'//'//'//'//'//'//'//'//'//'//'//'//'//'//

Public Const ROOTLocalizeActived = "LocalizeCellActived"
Public Const MaxTimeout = 15000
Public Enum EnumLocalizeActived
  EHA_on = 1
  EHA_off
  EHA_set
  EHA_Opacity
  EHA_Fading
  EHA_Color
  EHA_Spin
  EHA_Reset
  EHA_quit
  EHA_Uninstall
End Enum
Private HCSTimeID As Var64
Private HCSTimeID2 As Var64
Private HCSTimeout As Long
Private HCSDirective As EnumLocalizeActived
Private HCSCaller As Object



Public Function LocalizeSpin()
  SetTimerHAA EHA_Spin
End Function
Public Function LocalizeReset()
  SetTimerHAA EHA_Reset
End Function
Public Function LocalizeQuit()
  SetTimerHAA EHA_quit
End Function

Public Function LocalizeOn()
  SetTimerHAA EHA_on
End Function
Public Function LocalizeOff()
  SetTimerHAA EHA_off
End Function
Public Function LocalizeSetFading(Optional ByVal miliseconds As Integer = 3000)
  SetTimerHAA EHA_Fading
  hcaSetFading miliseconds
End Function
Public Function LocalizeSetColor(ByVal color As String)
  SetTimerHAA EHA_Color
  Dim v As Long
  Select Case True
  Case color Like "*[a-fA-F]*"
    If color Like "[#]*" Then
      color = Mid(color, 2)
    End If
    color = Mid(color, 5, 2) & Mid(color, 3, 2) & Mid(color, 1, 2)
    v = CLng(IIf(color Like "&H*", "", "&H") & color)
  Case IsNumeric(color): v = CLng(color)
  Case Else: v = vbBlue
  End Select
  Call SaveSetting(ROOTLocalizeActived, "Settings", "BackColor", CStr(v))
  LocalizeSetColor = v
End Function
Public Function LocalizeSetOpacity(ByVal opacity As Byte)
  Call SaveSetting(ROOTLocalizeActived, "Settings", "Opacity", CStr(opacity))
  Call LocalizeCellActivedStart
  LocalizeSetOpacity = opacity
End Function

Public Function LocalizeSet( _
      Optional color As Long, _
      Optional opacity As Byte, _
      Optional Fading As Boolean)
  SetTimerHAA EHA_set
  LocalizeSetOpacity opacity
  LocalizeSetColor color
  hcaSetFading Fading
End Function
Private Sub test_()
  Debug.Print "Opacity:"; hcaGetOpacity
  Debug.Print "BackColor:"; hcaGetBackColor
  Debug.Print "Fading:"; hcaGetFading
End Sub
Function hcaGetOpacity() As Byte
  hcaGetOpacity = GetSetting(ROOTLocalizeActived, "Settings", "Opacity", "40")
End Function
Function hcaGetBackColor() As Long
  hcaGetBackColor = GetSetting(ROOTLocalizeActived, "Settings", "BackColor", CStr(vbBlue))
End Function
Function hcaGetFading() As Integer
  hcaGetFading = GetSetting(ROOTLocalizeActived, "Settings", "Fading", "4000")
End Function
Function hcaSetFading(Optional ByVal miliseconds As Integer = 4000)
  If miliseconds > MaxTimeout Then miliseconds = MaxTimeout
  If miliseconds < 0 Then miliseconds = 0
  Call SaveSetting(ROOTLocalizeActived, "Settings", "Fading", CStr(miliseconds))
End Function


Private Sub LocalizeActivedAction()
  On Error Resume Next
  KillTimer 0&, HCSTimeID2.Long
  HCSCaller.Value = vbNullString
  Set HCSCaller = Nothing
  Select Case HCSDirective
  Case EHA_on, EHA_Fading, EHA_Color, EHA_Opacity
    Call LocalizeCellActivedStart
    formLocalizeCellActived.SetNewPosition
  Case EHA_off: Call LocalizeCellActivedStop
  Case EHA_Spin: ThisWorkbook.IsAddin = Not ThisWorkbook.IsAddin
  Case EHA_Reset:
    DeleteSetting ROOTLocalizeActived
  Case EHA_quit: ThisWorkbook.Close False
  Case EHA_Uninstall:
    Dim A, b As Boolean
    For Each A In Application.AddIns
      b = False: b = A.FullName = ThisWorkbook.FullName
      If b Then
        A.Installed = False
        Exit For
      End If
    Next
  End Select
  HCSDirective = 0
End Sub


Private Sub SetTimerHAA(Optional ByVal Directive As EnumLocalizeActived)
  On Error Resume Next
  HCSDirective = Directive
  KillTimer 0&, HCSTimeID2.Long
  Set HCSCaller = Application.Caller
  HCSTimeID2.Long = SetTimer(0&, 0&, 1, AddressOf LocalizeActivedAction)
End Sub

Public Sub LocalizeCellActivedStart()
  Dim u
  For Each u In VBA.UserForms
    If u.Name = "formLocalizeCellActived" Then
      Exit Sub
    End If
  Next
  VBA.Load formLocalizeCellActived
End Sub

Public Sub LocalizeCellActivedStop()
  On Error Resume Next
  EndtimeLocalizeCellActivedControl
  VBA.Unload formLocalizeCellActived
End Sub



'Buttons Test '//'// '//'//'//'//'//'//'//'//'//'//'//'//
Public Sub ChangeBackColor()
  If Cells.Interior.color = &H1E000A Then
    Cells.Interior.Pattern = xlNone
    Cells.Font.color = vbBlack
  Else
    Cells.Interior.color = &H1E000A
    Cells.Font.color = &HD9D9D9
  End If
  formLocalizeCellActived.SetNewPosition
End Sub
Sub toggleFullScreen()
  If Application.DisplayFormulaBar Then
    Application.DisplayFormulaBar = False
    Application.ExecuteExcel4Macro "Show.ToolBar(""Ribbon"", False)"
    ActiveWindow.DisplayHeadings = False
  Else
    Application.DisplayFormulaBar = True
    ActiveWindow.DisplayHeadings = True
    Application.ExecuteExcel4Macro "Show.ToolBar(""Ribbon"", True)"
  End If
End Sub
Public Sub ChangeFreeze()
  If ActiveWindow.View = xlPageLayoutView Then
    ActiveWindow.View = xlNormalView
  End If

  With ActiveWindow
    If .FreezePanes Then
      .FreezePanes = False
    Else
      [D8].Select
      .FreezePanes = True
    End If
  End With
  formLocalizeCellActived.SetNewPosition
End Sub

Public Sub ChangeFreeze2()
  If ActiveWindow.View = xlPageLayoutView Then
    ActiveWindow.View = xlNormalView
  End If
  With ActiveWindow
    If .FreezePanes Then
      .FreezePanes = False
      ActiveWindow.ScrollColumn = 1
      ActiveWindow.ScrollRow = 1
    Else
      ActiveWindow.ScrollColumn = 13
      ActiveWindow.ScrollRow = 12
      [O15].Select
      
      .FreezePanes = True
    End If
  End With
  ActiveSheet.DisplayPageBreaks = False
  formLocalizeCellActived.SetNewPosition
End Sub
Public Sub ChangeSplit()
  With ActiveWindow
    If .SplitColumn > 0 Or .SplitRow > 0 Then
      .SplitColumn = 0
      .SplitRow = 0
    Else
      .SplitColumn = 4
      .SplitRow = 7
    End If
  End With
  formLocalizeCellActived.SetNewPosition
End Sub
Public Sub FormulaBar()
  Application.DisplayFormulaBar = Not Application.DisplayFormulaBar
  formLocalizeCellActived.SetNewPosition
End Sub
Public Sub Header()
  ActiveWindow.DisplayHeadings = Not ActiveWindow.DisplayHeadings
  formLocalizeCellActived.SetNewPosition
End Sub
Public Sub DisplayRightToLeftSpin()
  ActiveWindow.DisplayRightToLeft = Not ActiveWindow.DisplayRightToLeft
  formLocalizeCellActived.SetNewPosition
End Sub



Sub setGroupV()
  With Range("D7:F17")
    If .Rows.OutlineLevel > 1 Then
      .Rows.Ungroup
    Else
      .Rows.Group
    End If
  End With
  formLocalizeCellActived.SetNewPosition
End Sub
Sub setGroupH()
  With Range("D7:F17")
    If .Columns.OutlineLevel > 1 Then
      .Columns.Ungroup
    Else
      .Columns.Group
    End If
  End With
  formLocalizeCellActived.SetNewPosition
End Sub
Sub ScrollBarV()
  ActiveWindow.DisplayVerticalScrollBar = Not ActiveWindow.DisplayVerticalScrollBar
  formLocalizeCellActived.SetNewPosition
End Sub
Sub ScrollBarH()
  ActiveWindow.DisplayHorizontalScrollBar = Not ActiveWindow.DisplayHorizontalScrollBar
  formLocalizeCellActived.SetNewPosition
End Sub
Sub WorkbookTabs()
  ActiveWindow.DisplayWorkbookTabs = Not ActiveWindow.DisplayWorkbookTabs
  formLocalizeCellActived.SetNewPosition
End Sub



Public Sub LocalizeCellActivedControl()
  HCSTimeout = 150
  LocalizeCellActived
End Sub

Public Sub LocalizeCellActived()
  On Error Resume Next
  EndtimeLocalizeCellActivedControl
  HCSTimeID.Long = SetTimer(0&, 0&, HCSTimeout, AddressOf LocalizeCellActivedControlCallback)
End Sub

Private Sub LocalizeCellActivedControlCallback()
  With formLocalizeCellActived
    If Not .Blurred(HCSTimeout) Or Not .SetPosition _
    Or HCSTimeout > MaxTimeout Then
      EndtimeLocalizeCellActivedControl
      .HideNow
      HCSTimeout = 0
    End If
  End With
  HCSTimeout = HCSTimeout + 150
End Sub

Public Sub EndtimeLocalizeCellActivedControl()
  On Error Resume Next
  KillTimer 0&, HCSTimeID.Long
End Sub


