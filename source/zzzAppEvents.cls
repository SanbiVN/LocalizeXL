VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "zzzAppEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Public WithEvents App As Excel.Application
Attribute App.VB_VarHelpID = -1
Public Wnd As Object
Public Contain As VBA.Collection

Private Enum EnumDirective
  EDUnload = -1
  EDHide
  EDWait
  EDShow
End Enum

Private Sub App_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)
  If Target.Rows.Count > 20 Or Target.Columns.Count > 10 Then
    LocalizeController ActiveWindow.Caption, Sh.Parent, Wnd, EDHide
  Else
    LocalizeController ActiveWindow.Caption, Sh.Parent, Wnd, EDShow
  End If
End Sub

Private Sub App_WindowDeactivate(ByVal Wb As Workbook, ByVal Wn As Window)
  LocalizeEndtime
End Sub

Private Sub App_WindowActivate(ByVal Wb As Workbook, ByVal Wn As Window)
  Debug.Print "App_WindowActivate"
  LocalizeController Wn.Caption, Wb, Wnd, EDShow
End Sub

Private Sub App_WindowResize(ByVal Wb As Workbook, ByVal Wn As Window)
  LocalizeController Wn.Caption, Wb, Wnd, EDWait
  If Wn.WindowState <> xlMinimized Then
    LocalizeController Wn.Caption, Wb, Wnd, EDShow
  End If
End Sub

Private Sub App_WorkbookBeforeClose(ByVal Wb As Workbook, Cancel As Boolean)
  Debug.Print "App_WorkbookBeforeClose:"; Wb.Name
  LocalizeController ActiveWindow.Caption, ActiveWorkbook, Wnd, EDUnload
End Sub

Private Sub Class_Initialize()
  Set App = Application
  Set Contain = New VBA.Collection
End Sub

Sub init()
  LocalizeController ActiveWindow.Caption, ActiveWorkbook, Wnd, EDShow
End Sub

Private Sub Class_Terminate()
  Debug.Print "Class_Terminate"
  On Error Resume Next
  Dim a, o As Object, i%
   i = Contain.Count
  For a = i To 1 Step -1
    Set o = Contain(a)
    o.HideNow
    VBA.Unload o
    Contain.Remove a
  Next
  Set Contain = Nothing
  Set App = Nothing
End Sub


Sub ShowLocalize()
  On Error Resume Next
  Wnd.SetNewPosition
  On Error GoTo 0
End Sub
Function Finally(timeout As Long) As Boolean
  If Not Wnd Is Nothing Then
    If Not Wnd.Blurred(timeout) Or Not Wnd.SetPosition _
    Or timeout > MaxTimeout _
    Or App.hwnd <> GetForegroundWindow() Then
      Wnd.HideNow
      Finally = True
    End If
  Else
    Finally = True
  End If
End Function
Private Sub LocalizeController(Key$, Wb As Object, Wnd As Object, Directive%)

  On Error Resume Next
  Set Wnd = Nothing
  Set Wnd = Contain(Key)
  'On Error GoTo 0
  If Wnd Is Nothing Then
    Select Case Directive
    Case EDShow:
      Contain.Add New formLocalizeCellActived, Key
      Set Wnd = Contain(Key)
      Contain(Key).SetNewPosition
    End Select
  Else
    Select Case Directive
    Case EDUnload:
          Wnd.Terminate
          Contain.Remove Key
    Case EDHide: Wnd.HideNow
    Case EDShow, 2: Wnd.SetNewPosition
    End Select
  End If
End Sub

