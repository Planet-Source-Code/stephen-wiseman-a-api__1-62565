Attribute VB_Name = "AAPI"
Option Explicit
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As CPoint) As Long
Private Declare Function GetWindowText Lib "user32.dll" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetWindowTextLength Lib "user32.dll" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long

Private Declare Function IsWindowVisible Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function IsWindowEnabled Lib "user32" (ByVal hWnd As Long) As Long

Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

Public Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long

Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Global Const conHwndTopmost = -1
Global Const conSwpNoActivate = &H10
Global Const conSwpShowWindow = &H40
    
Public Type CPoint
    x As Long
    y As Long
End Type

Public Const SW_SHOW = 5
Public Const SW_HIDE = 0

Dim pSpace As Long

Public Sub Auto()

    Dim MouseLocation As CPoint
    Dim hWnd As Long, hWndLength As Long, PhWndLength As Long, lClass As Long, PhWnd As Long
    Dim sWTextA As String, sWTextB As String, sClass As String, sPClass As String, lPClass As String, sPWTextA As String, sPWTextB As String
    
    GetCursorPos MouseLocation
    hWnd = WindowFromPoint(MouseLocation.x, MouseLocation.y)
    frmMain.txtHandle.Text = hWnd
    If hWnd = 0 Then frmMain.txtHandle.Text = "[null]"
    
    
    hWndLength = GetWindowTextLength(hWnd)
    sWTextA = String$(hWndLength, 0)
    sWTextB = GetWindowText(hWnd, sWTextA, (hWndLength + 1))
    frmMain.txtText.Text = sWTextA
    If frmMain.txtText.Text = "" Then frmMain.txtText.Text = "[null]"
    
    
    sClass = String(250, 0)
    lClass = GetClassName(hWnd, sClass, 250)
    frmMain.txtClass.Text = sClass
    If sClass = "" Then sClass = "[null]"
    
    PhWnd = GetParent(hWnd)
    frmMain.txtPHandle.Text = PhWnd
    If PhWnd = 0 Then frmMain.txtPHandle.Text = "[null]"
    
    PhWndLength = GetWindowTextLength(PhWnd)
    sPWTextA = String$(PhWndLength, 0)
    sPWTextB = GetWindowText(PhWnd, sPWTextA, (PhWndLength + 1))
    frmMain.txtPText.Text = sPWTextA
    If frmMain.txtPText.Text = "" Then frmMain.txtPText.Text = "[null]"
    
    sPClass = String(250, 0)
    lPClass = GetClassName(PhWnd, sPClass, 250)
    frmMain.txtPClass.Text = sPClass
    If frmMain.txtPClass.Text = "" Then frmMain.txtPClass.Text = "[null]"
    
End Sub

Public Sub SpaceBarA()
    
    pSpace = GetAsyncKeyState(32)
    
    If pSpace <> 0 Then
        frmMain.lblScanStatus.Caption = "* Paused - Hit Space Bar To Start *"
        frmMain.tmrSpace.Enabled = True
        frmMain.tmrRun.Enabled = False
    End If

End Sub

Public Sub SpaceBarB()
    
    pSpace = GetAsyncKeyState(32)
    
    If pSpace <> 0 Then
        frmMain.lblScanStatus.Caption = "* Running - Hit Space Bar To Pause *"
        frmMain.tmrSpace.Enabled = False
        frmMain.tmrRun.Enabled = True
    End If

End Sub

Public Function EnumALLWindows() As Boolean
    
    Dim hWnd As Long
    EnumWindows AddressOf CallBack, hWnd

End Function
Public Function CallBack(ByVal hWnd As Long, ByVal lpData As Long) As Long
    
    Dim lResult As Long
    Dim sWndName As String, sClassName As String, nWindowHandle As String
    Dim IsVisible As String, IsEnabled As String
    Dim lItem As ListItem
    Dim Image As Integer
    
    CallBack = 1
    sClassName = Space(260)
    sWndName = Space(260)
    lResult = GetClassName(hWnd, sClassName, 260)
    sClassName = Left$(sClassName, lResult)
    lResult = GetWindowText(hWnd, sWndName, 260)
    sWndName = Left$(sWndName, lResult)
    
    If sWndName = "" Then sWndName = " [null] "
    
    If IsWindowVisible(hWnd) = 0 Then
        IsVisible = "Not Visible"
        Image = 1
    Else
        IsVisible = "Visible"
        Image = 2
    End If
    
    If IsWindowEnabled(hWnd) = 0 Then IsEnabled = "Disabled" Else IsEnabled = "Enabled"
    
    Set lItem = frmWindows.lvWindows.ListItems.Add(, , hWnd)
    lItem.ListSubItems.Add , , sWndName
        lItem.ListSubItems.Add , , sClassName
        lItem.ListSubItems.Add , , IsVisible
        lItem.ListSubItems.Add , , IsEnabled

End Function
