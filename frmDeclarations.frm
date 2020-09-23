VERSION 5.00
Begin VB.Form frmDeclarations 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " A+ API - Frequently Used Declarations"
   ClientHeight    =   2655
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   9495
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2655
   ScaleWidth      =   9495
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdDone 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Done"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8040
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   " Exit "
      Top             =   2160
      Width           =   975
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1920
      ItemData        =   "frmDeclarations.frx":0000
      Left            =   120
      List            =   "frmDeclarations.frx":0002
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   9255
   End
   Begin VB.Label Quote 
      AutoSize        =   -1  'True
      Caption         =   """"
      Height          =   195
      Left            =   4200
      TabIndex        =   4
      Top             =   2280
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " Click Any Item to Copy the Text to Clip Board."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   405
      TabIndex        =   2
      Top             =   2280
      Width           =   3285
   End
   Begin VB.Label CopyCode 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "¤"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   15.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   360
      Index           =   0
      Left            =   120
      TabIndex        =   1
      ToolTipText     =   " Copty Details to Clip Board "
      Top             =   2160
      Width           =   315
   End
End
Attribute VB_Name = "frmDeclarations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdDone_Click()

    Unload Me
    
End Sub

Private Sub Form_Load()
    
    SetWindowPos hWnd, conHwndTopmost, 100, 100, 640, 200, conSwpNoActivate Or conSwpShowWindow
 
    List1.AddItem "Private Declare Function GetCursorPos Lib " & Quote.Caption & "user32.dll" & Quote.Caption & " (lpPoint As CPoint) As Long"
    List1.AddItem "Private Declare Function GetWindowText Lib  " & Quote.Caption & "user32.dll" & Quote.Caption & " Alias  " & Quote.Caption & "GetWindowTextA" & Quote.Caption & " (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long"
    List1.AddItem "Private Declare Function GetWindowTextLength Lib  " & Quote.Caption & "user32.dll" & Quote.Caption & " Alias  " & Quote.Caption & "GetWindowTextLengthA" & Quote.Caption & " (ByVal hWnd As Long) As Long"
    List1.AddItem "Private Declare Function GetClassName Lib  " & Quote.Caption & "user32.dll" & Quote.Caption & " Alias  " & Quote.Caption & "GetClassNameA" & Quote.Caption & " (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long"
    List1.AddItem "Private Declare Function GetParent Lib  " & Quote.Caption & "user32.dll" & Quote.Caption & " (ByVal hWnd As Long) As Long"
    List1.AddItem "Private Declare Function ShowWindow Lib  " & Quote.Caption & "user32.dll" & Quote.Caption & " (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long"
    List1.AddItem "Private Declare Function IsWindowVisible Lib  " & Quote.Caption & "user32.dll" & Quote.Caption & " (ByVal hWnd As Long) As Long"
    List1.AddItem "Private Declare Function IsWindowEnabled Lib  " & Quote.Caption & "user32.dll" & Quote.Caption & " (ByVal hWnd As Long) As Long"
    List1.AddItem "Private Declare Function GetAsyncKeyState Lib  " & Quote.Caption & "user32.dll" & Quote.Caption & " (ByVal vKey As Long) As Integer"
    List1.AddItem "Private Declare Function SetWindowPos Lib  " & Quote.Caption & "user32.dll" & Quote.Caption & " (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long"
    List1.AddItem "Private Declare Function WindowFromPoint Lib  " & Quote.Caption & "user32.dll" & Quote.Caption & " (ByVal xPoint As Long, ByVal yPoint As Long) As Long"
    List1.AddItem "Private Declare Function ShellExecute Lib  " & Quote.Caption & "shell32.dll" & Quote.Caption & " Alias  " & Quote.Caption & "ShellExecuteA" & Quote.Caption & " (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long"
    List1.AddItem "Private Declare Function SendMessage Lib  " & Quote.Caption & "user32.dll" & Quote.Caption & " Alias  " & Quote.Caption & "SendMessageA" & Quote.Caption & " (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long"
    List1.AddItem "Private Declare Function SendMessageByString Lib  " & Quote.Caption & "user32.dll" & Quote.Caption & " Alias  " & Quote.Caption & "SendMessageA" & Quote.Caption & " (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long"
    List1.AddItem "Private Declare Function ReleaseCapture Lib  " & Quote.Caption & "user32.dll" & Quote.Caption & " () As Long"
    List1.AddItem "Private Declare Function EnumChildWindows Lib  " & Quote.Caption & "user32.dll" & Quote.Caption & " (ByVal hWndParent As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long"
    List1.AddItem "Private Declare Function EnumWindows Lib  " & Quote.Caption & "user32.dll" & Quote.Caption & " (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long"
    List1.AddItem "Private Declare Function FindWindow Lib  " & Quote.Caption & "user32.dll" & Quote.Caption & " Alias  " & Quote.Caption & "FindWindowA" & Quote.Caption & " (ByVal lpClassName As String, ByVal lpWindowName As String) As Long"
    List1.AddItem "Private Declare Function FindWindowEx Lib  " & Quote.Caption & "user32.dll" & Quote.Caption & " Alias  " & Quote.Caption & "FindWindowExA" & Quote.Caption & " (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long"

End Sub

Private Sub List1_Click()

    Clipboard.SetText List1.List(List1.ListIndex)


End Sub
