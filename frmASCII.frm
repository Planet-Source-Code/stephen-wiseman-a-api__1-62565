VERSION 5.00
Begin VB.Form frmASCII 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " A+ API -ASCII"
   ClientHeight    =   1815
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4095
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1815
   ScaleWidth      =   4095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraASCIII 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   " ASCII Viewer "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3855
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Hit any key to see its' Information.  Click on a label to copy it to the Clip Board."
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   1080
         Width           =   3615
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Character"
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
         Left            =   240
         TabIndex        =   6
         Top             =   240
         Width           =   720
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Key ASCII"
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
         Left            =   2880
         TabIndex        =   5
         Top             =   240
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Key Code"
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
         Left            =   1680
         TabIndex        =   4
         Top             =   240
         Width           =   705
      End
      Begin VB.Label lblChar 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   3
         ToolTipText     =   " Regular Character "
         Top             =   480
         Width           =   975
      End
      Begin VB.Label lblKASCII 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2760
         TabIndex        =   2
         ToolTipText     =   " Character ASCII Code "
         Top             =   480
         Width           =   975
      End
      Begin VB.Label lblKC 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1560
         TabIndex        =   1
         ToolTipText     =   " Character Key Code "
         Top             =   480
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmASCII"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdDone_Click()

    frmASCII.Hide

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    lblKC.Caption = KeyCode
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
   
   lblKASCII = KeyAscii
   
   lblChar = Chr$(KeyAscii)

End Sub

Private Sub Form_Load()
    
    SetWindowPos hWnd, conHwndTopmost, 100, 100, 280, 155, conSwpNoActivate Or conSwpShowWindow
        
End Sub

Private Sub Form_LostFocus()

    Unload Me
    
End Sub

Private Sub lblChar_Click()

    Clipboard.SetText lblChar.Caption
    
End Sub

Private Sub lblKASCII_Click()

    Clipboard.SetText lblKASCII.Caption
    
End Sub

Private Sub lblKC_Click()

    Clipboard.SetText lblKC.Caption
    
End Sub
