VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} formLocalizeCellActived 
   Caption         =   "LocalizeCellActived"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "formLocalizeCellActived.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "formLocalizeCellActived"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Private WithEvents App  As Excel.Application
Attribute App.VB_VarHelpID = -1
Private AW As Excel.Window
Private hPolygon As Var64
Private hMain As Var64
Private hXL7 As Var64
Private hXLD As Var64
Private hWnd_  As Var64
Private coors(10) As POINTAPI
Private pt As POINTAPI
Private RA, nVer As Boolean
Private X&, Y&, w#, h#, VS#
Private opacity As Byte
Private Text100 As String * 100
Private Ret100 As Long
Private PA As RECT6, LA As RECT6
Private RXL7 As RECT
Private RO As Excel.Range
Private RS As Excel.Range
Private RV As Excel.Range
Private RH As Excel.Range
Private LRSC As Long
Private LCSC As Long
Private i As Byte
Private idx As Byte
Private c As Byte
Private T As Byte
Private sr As Boolean
Private sc As Boolean
Private bV As Boolean
Private bH As Boolean
Private DRTL As Boolean
Private DHSB As Boolean
Private DVSB As Boolean
Private DWB As Boolean
Private SU As Boolean

Sub HideNow()
  ShowWindow hWnd_.Long, 0
End Sub


Private Sub UserForm_Initialize()
  Me.Caption = "LocalizeCellActived"
  Dim b As Boolean
  Set App = Application
  nVer = Val(App.Version) > 14

  hWnd_ = FormHandle(Me.Caption)
  b = NewHandle()
  hPolygon.Long = GetWindowLong(hWnd_.Long, (-20)) Or &H80000
  Call SetWindowLong(hWnd_.Long, (-20), hPolygon.Long Or &H100000 Or &H8000000 Or &H20& Or &H1)

  SetWindowLong hWnd_.Long, -16, GetWindowLong(hWnd_.Long, -16) And Not &HC00000
  SetWindowLong hWnd_.Long, -20, GetWindowLong(hWnd_.Long, -20) And Not &H1
  SetWindowLong hWnd_.Long, -20, (GetWindowLong(hWnd_.Long, -20) Or &H80000) Or &H100000

  If b Then
    SetParent hWnd_.Long, hXL7.Long
  End If
End Sub

Function NewHandle() As Boolean
  Dim th As Var64
  If nVer Then
    th.Long = FindWindow("XLMAIN", App.Caption)
    If th.Long <> hMain.Long Then
      Set AW = ActiveWindow
      hMain.Long = th.Long
      hXLD.Long = FindWindowEx(th.Long, 0&, "XLDESK", vbNullString)
      hXL7.Long = FindWindowEx(hXLD.Long, 0&, "EXCEL7", AW.Caption)
      NewHandle = True
    End If
  Else
    hMain.Long = App.hwnd
    hXLD.Long = FindWindowEx(hMain.Long, 0&, "XLDESK", vbNullString)
    th.Long = FindWindowEx(hXLD.Long, 0&, "EXCEL7", AW.Caption)
    If th.Long <> hXL7.Long Then
      Set AW = ActiveWindow
      hXL7.Long = th.Long
      NewHandle = True
    End If
  End If

  Call GetWindowRect(hXL7.Long, RXL7)
End Function

Sub SetNewPosition()
  EndtimeLocalizeCellActivedControl
  Dim color As Long
  color = hcaGetBackColor()
  If Me.BackColor <> color Then
    Me.BackColor = color
  End If
  opacity = hcaGetOpacity()
  Transparent opacity
  ShowWindow hWnd_.Long, 5
  If NewHandle() Then
    SetParent hWnd_.Long, hXL7.Long
  End If
  If SetPosition Then
    LocalizeCellActivedControl
  End If
End Sub

Function SetPosition() As Boolean
  If opacity <= 2 Then
    Call HideNow
    Exit Function
  End If
  If Not Me.Visible Then
    Exit Function
  End If
  On Error Resume Next
  Set RA = Nothing
  Set RA = Selection
  If TypeName(RA) <> "Range" _
  Or AW.View = xlPageLayoutView Then
    Call HideNow
    Exit Function
  End If
  On Error GoTo 0
  
  '//'//'//'//'//'//'//'//'//'//
  Call Excel7Positions(RA, PA)
  For i = 0 To 3
    If PA.XY(i).X <> LA.XY(i).X Or PA.XY(i).Y <> LA.XY(i).Y Then
      GoTo n:
    End If
  Next
  If i > 0 Then
    SetPosition = True
    Exit Function
  End If
  If PA.XY(3).X <= 0 Or PA.XY(3).Y <= 0 Or VBA.Err Then
    Call HideNow
    Exit Function
  End If
n:
  LRSC = AW.ScrollRow
  LCSC = AW.ScrollColumn
  MoveWindow hWnd_.Long, 0, 0, PA.XY(3).X, PA.XY(3).Y, False
  '//'//'//'//'//'//'//'//'//'//
  coors(0).X = PA.XY(1).X: coors(0).Y = PA.XY(0).Y
  coors(1).X = PA.XY(2).X: coors(1).Y = PA.XY(0).Y
  coors(2).X = PA.XY(2).X: coors(2).Y = PA.XY(3).Y
  coors(3).X = PA.XY(1).X: coors(3).Y = PA.XY(3).Y
  coors(4).X = PA.XY(1).X: coors(4).Y = PA.XY(2).Y
  coors(5).X = PA.XY(3).X: coors(5).Y = PA.XY(2).Y
  coors(6).X = PA.XY(3).X: coors(6).Y = PA.XY(1).Y
  coors(7).X = PA.XY(0).X: coors(7).Y = PA.XY(1).Y
  coors(8).X = PA.XY(0).X: coors(8).Y = PA.XY(2).Y
  coors(9).X = PA.XY(1).X: coors(9).Y = PA.XY(2).Y
                         
  hPolygon.Long = CreatePolygonRgn(coors(0), UBound(coors), 1)
  Call SetWindowRgn(hWnd_.Long, hPolygon.Long, True)
  Call DeleteObject(hPolygon.Long)
  AppActivate AW.Caption
  SetPosition = True
  LA = PA
End Function


Private Sub Excel7Positions(ByVal Target As Range, PaneStatistics As RECT6)
' Last Edit: 19/02/2021 13:24
'   |--x0-----x1----x2-----x3
'   |
'   y0        0----->1
'   |         ^''''''|
'   |         |''''''|
'   y1 7------|------|----->6
'   |  ^''''''|      |''''''|
'   |  |''''''|      |''''''v
'   y2 8<----4&9-----|------5
'   |         |''''''|
'   |         |''''''v
'   y3        3<-----2
  On Error GoTo E
  Dim p As RECT6
  p.WindowState = AW.WindowState
  DRTL = AW.DisplayRightToLeft
  DVSB = AW.DisplayVerticalScrollBar
  DHSB = AW.DisplayHorizontalScrollBar
  DWB = AW.DisplayWorkbookTabs
  sr = AW.SplitRow
  sc = AW.SplitColumn
  c = AW.Panes.Count
  idx = AW.ActivePane.Index
  '//'//'//'//'//'//'//'//'//'//'//'//
  VS = IIf(DVSB, IIf(nVer, 20, 26), IIf(nVer, 0, 4))
  
  '//'//'//'//'//'//'//'//'//'//'//'//
  Set RS = AW.Panes(1).VisibleRange
  X = AW.Panes(1).PointsToScreenPixelsX(RS.Left)
  Y = AW.Panes(1).PointsToScreenPixelsY(RS.Top)
  
  If DRTL Then
    p.XY(0).X = RXL7.Left + VS
    p.XY(3).X = RXL7.Right - X + RXL7.Left
  Else
    p.XY(0).X = X
    p.XY(3).X = RXL7.Right - IIf(DVSB, VS, 0)
  End If
  p.XY(0).Y = Y
  p.XY(3).Y = RXL7.Bottom - IIf(DHSB Or DWB, 26, 0)
  
  '//'//'//'//'//'//'//'//'//'//'//'//
  T = 0: bV = False: bH = False
  For i = 1 To c
    Set RS = AW.Panes(i).VisibleRange
    p.Areas(i) = RS.Address(0, 0)
    If AW.FreezePanes Then
s:
      Set RO = App.Intersect(Target, RS)
      Set RV = App.Intersect(Target.EntireColumn, RS)
      Set RH = App.Intersect(Target.EntireRow, RS)
      If Not RV Is Nothing Then
        If Not bV Then
          p.XY(1 - DRTL).X = AW.Panes(i).PointsToScreenPixelsX(RV.Left)
          bV = True
        End If
        If T = 0 _
        Or (sc And T = 1 And i = 2) _
        Or (sc And T = 3 And i = 4) Then
          p.XY(2 + DRTL).X = AW.Panes(i).PointsToScreenPixelsX(RV(1, RV.Columns.Count + 1).Left)
        End If
      End If
      If Not RH Is Nothing Then
        If Not bH Then
          p.XY(1).Y = AW.Panes(i).PointsToScreenPixelsY(RH.Top)
          bH = True
        End If
        If T = 0 _
        Or (sr And c = 2 And T = 1 And i = 2) _
        Or (c = 4 And T = 1 And i = 3) _
        Or (c = 4 And T = 2 And i = 4) Then
          p.XY(2).Y = AW.Panes(i).PointsToScreenPixelsY(RH(RH.Rows.Count + 1, 1).Top)
        End If
      End If
      If T = 0 Then
        T = IIf(Not RO Is Nothing, i, 0)
      End If
    Else
      If Not AW.Split Or idx = i Then
        GoTo s
      End If
    End If
  Next i
  For i = 0 To 3
    ScreenToClient hXL7.Long, p.XY(i)
  Next
  If DRTL Then
    p.XY(2).X = p.XY(3).X - p.XY(2).X + X - RXL7.Left - IIf(nVer, 0, 9)
    p.XY(1).X = p.XY(3).X - p.XY(1).X + X - RXL7.Left - IIf(nVer, 0, 9)
    If p.XY(1).X < p.XY(0).X Then
      p.XY(1).X = p.XY(0).X
    End If
  End If
  If p.XY(2).X > p.XY(3).X Then
    p.XY(2).X = p.XY(3).X
  End If
  If p.XY(2).Y > p.XY(3).Y Then
    p.XY(2).Y = p.XY(3).Y
  End If
  
  If p.XY(2).X - p.XY(1).X > p.XY(3).X * 2 / 3 Then
    p.XY(3).X = 0
  End If
  If p.XY(2).Y - p.XY(1).Y > p.XY(3).Y * 2 / 3 Then
    p.XY(3).Y = 0
  End If
  
  PaneStatistics = p
E:
  Set RS = Nothing: Set RO = Nothing: Set RV = Nothing: Set RH = Nothing
End Sub


Private Sub App_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)
  If Target.Rows.Count > 20 Or Target.Columns.Count > 10 Then
    EndtimeLocalizeCellActivedControl
    Call HideNow
  Else
    Call SetNewPosition
  End If
End Sub


Private Sub App_WindowResize(ByVal Wb As Workbook, ByVal Wn As Window)
  If Wn.WindowState <> xlMinimized Then
    SetNewPosition
  Else
    EndtimeLocalizeCellActivedControl
    Call HideNow
  End If
End Sub
Private Sub App_WorkbookBeforeClose(ByVal Wb As Workbook, Cancel As Boolean)
  EndtimeLocalizeCellActivedControl
  Application.OnTime VBA.Now, "'" & ThisWorkbook.Name & "'!LocalizeCellActivedStart"
  Unload Me
End Sub
Function Blurred(Optional ByVal order As Integer) As Boolean
  Dim f As Integer
  f = hcaGetFading()
  Blurred = f <= 0
  If f > 0 And order <= f Then
    Blurred = True
    Transparent opacity * ((f - order) / f)
  End If
End Function

Private Function ScreenDPI(Optional bVertical As Boolean) As Long
  Dim h As Var64
  Static L&(1)
  If L(0) = 0 Then
    h.Long = GetDC(0&)
    L(0) = GetDeviceCaps(h.Long, 88&)
    L(1) = GetDeviceCaps(h.Long, 90&)
    Call ReleaseDC(0&, h.Long)
  End If
  ScreenDPI = L(-bVertical)
End Function

Private Function Points2Pixels(ByVal POINTS!, Optional bVertical As Boolean)
  Points2Pixels = POINTS * ScreenDPI(bVertical) / 72
End Function
Private Function Pixels2Points(ByVal PIXELS&, Optional bVertical As Boolean)
  Pixels2Points = PIXELS * 72 / ScreenDPI(bVertical)
End Function


Private Function FormHandle(Optional ByVal Caption$) As Var64
  If Val(App.Version) < 9 Then
    FormHandle.Long = FindWindow("ThunderXFrame", Caption)
  Else
    FormHandle.Long = FindWindow("ThunderDFrame", Caption)
  End If
End Function

Private Sub Transparent(Optional ByVal opacity As Byte = 255)
  SetLayeredWindowAttributes hWnd_.Long, 0&, opacity, 2
End Sub

Private Sub UserForm_Terminate()
  EndtimeLocalizeCellActivedControl
  Set App = Nothing
  Set AW = Nothing
  Set RO = Nothing
  Set RS = Nothing
  Set RV = Nothing
  Set RH = Nothing
End Sub





