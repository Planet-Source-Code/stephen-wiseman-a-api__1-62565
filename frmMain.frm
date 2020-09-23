VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " A+ API - Main"
   ClientHeight    =   4575
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5175
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4575
   ScaleWidth      =   5175
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdWindows 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&hWnd's"
      Height          =   375
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   28
      ToolTipText     =   " hWnd List "
      Top             =   4080
      Width           =   975
   End
   Begin VB.CommandButton cmdPDF 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&PDF's"
      Height          =   375
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   29
      ToolTipText     =   " Private Declare Functions "
      Top             =   4080
      Width           =   975
   End
   Begin VB.CommandButton cmdColors 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Colors"
      Height          =   375
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   30
      ToolTipText     =   " Color Codes "
      Top             =   4080
      Width           =   975
   End
   Begin VB.CommandButton cmdASCII 
      BackColor       =   &H00E0E0E0&
      Caption         =   "A&SCII"
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   " ASCII Characters "
      Top             =   4080
      Width           =   975
   End
   Begin VB.Timer tmrCB 
      Interval        =   250
      Left            =   720
      Top             =   5160
   End
   Begin VB.Timer tmrSpace 
      Interval        =   150
      Left            =   1200
      Top             =   5160
   End
   Begin VB.Frame fraCB 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   " Clip Board Viewer "
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   120
      TabIndex        =   22
      Top             =   3360
      Width           =   4935
      Begin VB.TextBox txtCB 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Height          =   255
         Left            =   480
         Locked          =   -1  'True
         TabIndex        =   23
         ToolTipText     =   " Clip Board Display "
         Top             =   240
         Width           =   4335
      End
      Begin VB.Label CBPic 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "§"
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
         Index           =   6
         Left            =   120
         TabIndex        =   24
         ToolTipText     =   " Clear the Clip Board "
         Top             =   160
         Width           =   315
      End
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00E0E0E0&
      Caption         =   "E&xit"
      Height          =   375
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   " Exit "
      Top             =   4080
      Width           =   975
   End
   Begin VB.Frame fraDetails 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   " Window Details "
      ForeColor       =   &H80000008&
      Height          =   2775
      Left            =   120
      TabIndex        =   9
      ToolTipText     =   " Window Details View "
      Top             =   120
      Width           =   4935
      Begin VB.TextBox txtHandle 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   1
         Text            =   "[null]"
         ToolTipText     =   " Click and Drag to Highlight Text "
         Top             =   240
         Width           =   2535
      End
      Begin VB.TextBox txtPHandle 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   4
         Text            =   "[null]"
         ToolTipText     =   " Click and Drag to Highlight Text "
         Top             =   1920
         Width           =   2535
      End
      Begin VB.TextBox txtPText 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   5
         Text            =   "[null]"
         ToolTipText     =   " Click and Drag to Highlight Text "
         Top             =   2280
         Width           =   2535
      End
      Begin VB.TextBox txtPClass 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   6
         Text            =   "[null]"
         ToolTipText     =   " Click and Drag to Highlight Text "
         Top             =   1440
         Width           =   2535
      End
      Begin VB.TextBox txtText 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   3
         Text            =   "[null]"
         ToolTipText     =   " Click and Drag to Highlight Text "
         Top             =   600
         Width           =   2535
      End
      Begin VB.TextBox txtClass 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   2
         Text            =   "[null]"
         ToolTipText     =   " Click and Drag to Highlight Text "
         Top             =   1080
         Width           =   2535
      End
      Begin VB.Label CreateCode 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "ù"
         BeginProperty Font 
            Name            =   "Webdings"
            Size            =   15.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   360
         Index           =   1
         Left            =   4560
         TabIndex        =   26
         ToolTipText     =   " Click to Generate Code "
         Top             =   1320
         Width           =   315
      End
      Begin VB.Label CreateCode 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "ù"
         BeginProperty Font 
            Name            =   "Webdings"
            Size            =   15.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   360
         Index           =   0
         Left            =   4560
         TabIndex        =   25
         ToolTipText     =   " Click to Generate Code "
         Top             =   960
         Width           =   315
      End
      Begin VB.Label CopyToCB 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "§"
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
         Index           =   5
         Left            =   4200
         TabIndex        =   21
         ToolTipText     =   " Copty Details to Clip Board "
         Top             =   2160
         Width           =   315
      End
      Begin VB.Label CopyToCB 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "§"
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
         Index           =   4
         Left            =   4200
         TabIndex        =   20
         ToolTipText     =   " Copty Details to Clip Board "
         Top             =   1800
         Width           =   315
      End
      Begin VB.Label CopyToCB 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "§"
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
         Index           =   3
         Left            =   4200
         TabIndex        =   19
         ToolTipText     =   " Copty Details to Clip Board "
         Top             =   1320
         Width           =   315
      End
      Begin VB.Label CopyToCB 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "§"
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
         Index           =   2
         Left            =   4200
         TabIndex        =   18
         ToolTipText     =   " Copty Details to Clip Board "
         Top             =   480
         Width           =   315
      End
      Begin VB.Label CopyToCB 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "§"
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
         Index           =   1
         Left            =   4200
         TabIndex        =   17
         ToolTipText     =   " Copty Details to Clip Board "
         Top             =   960
         Width           =   315
      End
      Begin VB.Label CopyToCB 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "§"
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
         Left            =   4200
         TabIndex        =   16
         ToolTipText     =   " Copty Details to Clip Board "
         Top             =   120
         Width           =   315
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "+  Parent hWnd -"
         Height          =   210
         Left            =   120
         TabIndex        =   15
         Top             =   1920
         Width           =   1215
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "+  Parent Text -"
         Height          =   210
         Left            =   120
         TabIndex        =   14
         Top             =   2280
         Width           =   1110
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "+  Parent Class -"
         Height          =   210
         Left            =   120
         TabIndex        =   13
         Top             =   1440
         Width           =   1200
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "+  Text -"
         Height          =   210
         Left            =   120
         TabIndex        =   12
         Top             =   600
         Width           =   600
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "+  Class -"
         Height          =   210
         Left            =   120
         TabIndex        =   11
         Top             =   1080
         Width           =   690
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "+  Handle (hWnd) -"
         Height          =   210
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   1365
      End
   End
   Begin VB.Timer tmrRun 
      Interval        =   150
      Left            =   240
      Top             =   5160
   End
   Begin VB.Label lblScanStatus 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Paused - Hit Space Bar to Start"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   " Status "
      Top             =   3000
      Width           =   4935
   End
   Begin VB.Label Quotes 
      Caption         =   """"
      Height          =   135
      Left            =   0
      TabIndex        =   27
      Top             =   3240
      Visible         =   0   'False
      Width           =   255
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CBPic_Click(Index As Integer)

    Clipboard.Clear
    
End Sub

Private Sub cmdAbout_Click()

    frmWindows.Show
    
End Sub

Private Sub cmdASCII_Click()

    frmASCII.Show
    
End Sub

Private Sub cmdExit_Click()

    tmrRun.Enabled = False
    Unload Me
    End
    
End Sub

Private Sub cmdPDF_Click()

    frmDeclarations.Show
    
End Sub

Private Sub cmdWindows_Click()

    frmWindows.Show
    
End Sub

Private Sub cmdColors_Click()

    frmColors.Show
    
End Sub

Private Sub CopyToCB_Click(Index As Integer)

    Select Case Index
        
        Case 0
            Clipboard.SetText txtHandle.Text
        Case 1
            Clipboard.SetText txtClass.Text
        Case 2
            Clipboard.SetText txtText.Text
        Case 3
            Clipboard.SetText txtPHandle.Text
        Case 4
            Clipboard.SetText txtPClass.Text
        Case 5
            Clipboard.SetText txtPText.Text
    
    End Select

End Sub

Private Sub CreateCode_Click(Index As Integer)
    
    Dim pText As String
    
    Select Case Index
        
        Case 0
            frmCode.txtCode.Text = "Dim pWin as Long" & vbCrLf & vbCrLf & _
            "pWin = FindWindow(" & Quotes.Caption & txtClass.Text & Quotes.Caption & ", vbNullString)"
            frmCode.Show
        Case 1
            If txtPText.Text = "[null]" Then
                pText = " vbNullString"
            Else
                pText = Quotes.Caption & txtPText.Text & Quotes.Caption
            End If
            frmCode.txtCode.Text = "Dim pWin as Long" & vbCrLf & "Dim cWin as Long" & vbCrLf & _
            vbCrLf & "pWin = FindWindow(" & Quotes.Caption & txtPClass.Text & Quotes.Caption & ", vbNullString)" & vbCrLf & _
            "cWin = FindWindowEx(pWin, 0&," & pText & ")"
            frmCode.Show
    
    End Select

End Sub

Private Sub Form_Load()
  
    tmrSpace.Enabled = False
    tmrRun.Enabled = True
    Clipboard.Clear
    lblScanStatus.Caption = "* Running - Hit Space Bar To Pause *"
    SetWindowPos hWnd, conHwndTopmost, 100, 100, 350, 335, conSwpNoActivate Or conSwpShowWindow
    frmMain.SetFocus
    
End Sub

Private Sub Form_Terminate()

    tmrRun.Enabled = False
    tmrSpace.Enabled = False
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    tmrRun.Enabled = False
    tmrSpace.Enabled = False
    
End Sub

Private Sub tmrCB_Timer()

    txtCB.Text = Clipboard.GetText

End Sub

Private Sub tmrRun_Timer()

    Auto
    SpaceBarA

End Sub

Private Sub tmrSpace_Timer()

    SpaceBarB

End Sub
