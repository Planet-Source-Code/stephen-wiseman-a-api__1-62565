VERSION 5.00
Begin VB.Form frmColors 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " A+ API - Colors"
   ClientHeight    =   3990
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4350
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3990
   ScaleWidth      =   4350
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Done"
      Height          =   375
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   54
      ToolTipText     =   " Exit "
      Top             =   3480
      Width           =   975
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   3165
      Left            =   120
      ScaleHeight     =   3135
      ScaleWidth      =   4095
      TabIndex        =   0
      Top             =   120
      Width           =   4120
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   120
         ScaleHeight     =   585
         ScaleWidth      =   3825
         TabIndex        =   49
         Top             =   2400
         Width           =   3855
         Begin VB.PictureBox PicColor 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   120
            ScaleHeight     =   345
            ScaleWidth      =   465
            TabIndex        =   50
            Top             =   120
            Width           =   495
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
            Left            =   2040
            TabIndex        =   53
            ToolTipText     =   " Copty Details to Clip Board "
            Top             =   105
            Width           =   315
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Copy to Clip Board."
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
            Left            =   2400
            TabIndex        =   52
            Top             =   195
            Width           =   1380
         End
         Begin VB.Label lblColor 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "&&H00FFFFFF"
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
            Left            =   720
            TabIndex        =   51
            Top             =   200
            Width           =   930
         End
      End
      Begin VB.Timer tmrColor 
         Interval        =   150
         Left            =   120
         Top             =   3240
      End
      Begin VB.PictureBox Color 
         Appearance      =   0  'Flat
         BackColor       =   &H00400040&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   47
         Left            =   3480
         ScaleHeight     =   345
         ScaleWidth      =   465
         TabIndex        =   48
         Top             =   1920
         Width           =   495
      End
      Begin VB.PictureBox Color 
         Appearance      =   0  'Flat
         BackColor       =   &H00800080&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   46
         Left            =   3480
         ScaleHeight     =   345
         ScaleWidth      =   465
         TabIndex        =   47
         Top             =   1560
         Width           =   495
      End
      Begin VB.PictureBox Color 
         Appearance      =   0  'Flat
         BackColor       =   &H00C000C0&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   45
         Left            =   3480
         ScaleHeight     =   345
         ScaleWidth      =   465
         TabIndex        =   46
         Top             =   1200
         Width           =   495
      End
      Begin VB.PictureBox Color 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF00FF&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   44
         Left            =   3480
         ScaleHeight     =   345
         ScaleWidth      =   465
         TabIndex        =   45
         Top             =   840
         Width           =   495
      End
      Begin VB.PictureBox Color 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF80FF&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   43
         Left            =   3480
         ScaleHeight     =   345
         ScaleWidth      =   465
         TabIndex        =   44
         Top             =   480
         Width           =   495
      End
      Begin VB.PictureBox Color 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0FF&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   42
         Left            =   3480
         ScaleHeight     =   345
         ScaleWidth      =   465
         TabIndex        =   43
         Top             =   120
         Width           =   495
      End
      Begin VB.PictureBox Color 
         Appearance      =   0  'Flat
         BackColor       =   &H00400000&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   41
         Left            =   3000
         ScaleHeight     =   345
         ScaleWidth      =   465
         TabIndex        =   42
         Top             =   1920
         Width           =   495
      End
      Begin VB.PictureBox Color 
         Appearance      =   0  'Flat
         BackColor       =   &H00800000&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   40
         Left            =   3000
         ScaleHeight     =   345
         ScaleWidth      =   465
         TabIndex        =   41
         Top             =   1560
         Width           =   495
      End
      Begin VB.PictureBox Color 
         Appearance      =   0  'Flat
         BackColor       =   &H00C00000&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   39
         Left            =   3000
         ScaleHeight     =   345
         ScaleWidth      =   465
         TabIndex        =   40
         Top             =   1200
         Width           =   495
      End
      Begin VB.PictureBox Color 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   38
         Left            =   3000
         ScaleHeight     =   345
         ScaleWidth      =   465
         TabIndex        =   39
         Top             =   840
         Width           =   495
      End
      Begin VB.PictureBox Color 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   37
         Left            =   3000
         ScaleHeight     =   345
         ScaleWidth      =   465
         TabIndex        =   38
         Top             =   480
         Width           =   495
      End
      Begin VB.PictureBox Color 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   36
         Left            =   3000
         ScaleHeight     =   345
         ScaleWidth      =   465
         TabIndex        =   37
         Top             =   120
         Width           =   495
      End
      Begin VB.PictureBox Color 
         Appearance      =   0  'Flat
         BackColor       =   &H00404000&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   35
         Left            =   2520
         ScaleHeight     =   345
         ScaleWidth      =   465
         TabIndex        =   36
         Top             =   1920
         Width           =   495
      End
      Begin VB.PictureBox Color 
         Appearance      =   0  'Flat
         BackColor       =   &H00808000&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   34
         Left            =   2520
         ScaleHeight     =   345
         ScaleWidth      =   465
         TabIndex        =   35
         Top             =   1560
         Width           =   495
      End
      Begin VB.PictureBox Color 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C000&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   33
         Left            =   2520
         ScaleHeight     =   345
         ScaleWidth      =   465
         TabIndex        =   34
         Top             =   1200
         Width           =   495
      End
      Begin VB.PictureBox Color 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   32
         Left            =   2520
         ScaleHeight     =   345
         ScaleWidth      =   465
         TabIndex        =   33
         Top             =   840
         Width           =   495
      End
      Begin VB.PictureBox Color 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF80&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   31
         Left            =   2520
         ScaleHeight     =   345
         ScaleWidth      =   465
         TabIndex        =   32
         Top             =   480
         Width           =   495
      End
      Begin VB.PictureBox Color 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   30
         Left            =   2520
         ScaleHeight     =   345
         ScaleWidth      =   465
         TabIndex        =   31
         Top             =   120
         Width           =   495
      End
      Begin VB.PictureBox Color 
         Appearance      =   0  'Flat
         BackColor       =   &H00004000&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   29
         Left            =   2040
         ScaleHeight     =   345
         ScaleWidth      =   465
         TabIndex        =   30
         Top             =   1920
         Width           =   495
      End
      Begin VB.PictureBox Color 
         Appearance      =   0  'Flat
         BackColor       =   &H00008000&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   28
         Left            =   2040
         ScaleHeight     =   345
         ScaleWidth      =   465
         TabIndex        =   29
         Top             =   1560
         Width           =   495
      End
      Begin VB.PictureBox Color 
         Appearance      =   0  'Flat
         BackColor       =   &H0000C000&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   27
         Left            =   2040
         ScaleHeight     =   345
         ScaleWidth      =   465
         TabIndex        =   28
         Top             =   1200
         Width           =   495
      End
      Begin VB.PictureBox Color 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FF00&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   26
         Left            =   2040
         ScaleHeight     =   345
         ScaleWidth      =   465
         TabIndex        =   27
         Top             =   840
         Width           =   495
      End
      Begin VB.PictureBox Color 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FF80&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   25
         Left            =   2040
         ScaleHeight     =   345
         ScaleWidth      =   465
         TabIndex        =   26
         Top             =   480
         Width           =   495
      End
      Begin VB.PictureBox Color 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   24
         Left            =   2040
         ScaleHeight     =   345
         ScaleWidth      =   465
         TabIndex        =   25
         Top             =   120
         Width           =   495
      End
      Begin VB.PictureBox Color 
         Appearance      =   0  'Flat
         BackColor       =   &H00004040&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   23
         Left            =   1560
         ScaleHeight     =   345
         ScaleWidth      =   465
         TabIndex        =   24
         Top             =   1920
         Width           =   495
      End
      Begin VB.PictureBox Color 
         Appearance      =   0  'Flat
         BackColor       =   &H00008080&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   22
         Left            =   1560
         ScaleHeight     =   345
         ScaleWidth      =   465
         TabIndex        =   23
         Top             =   1560
         Width           =   495
      End
      Begin VB.PictureBox Color 
         Appearance      =   0  'Flat
         BackColor       =   &H0000C0C0&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   21
         Left            =   1560
         ScaleHeight     =   345
         ScaleWidth      =   465
         TabIndex        =   22
         Top             =   1200
         Width           =   495
      End
      Begin VB.PictureBox Color 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   20
         Left            =   1560
         ScaleHeight     =   345
         ScaleWidth      =   465
         TabIndex        =   21
         Top             =   840
         Width           =   495
      End
      Begin VB.PictureBox Color 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   19
         Left            =   1560
         ScaleHeight     =   345
         ScaleWidth      =   465
         TabIndex        =   20
         Top             =   480
         Width           =   495
      End
      Begin VB.PictureBox Color 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   18
         Left            =   1560
         ScaleHeight     =   345
         ScaleWidth      =   465
         TabIndex        =   19
         Top             =   120
         Width           =   495
      End
      Begin VB.PictureBox Color 
         Appearance      =   0  'Flat
         BackColor       =   &H00404080&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   17
         Left            =   1080
         ScaleHeight     =   345
         ScaleWidth      =   465
         TabIndex        =   18
         Top             =   1920
         Width           =   495
      End
      Begin VB.PictureBox Color 
         Appearance      =   0  'Flat
         BackColor       =   &H00004080&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   16
         Left            =   1080
         ScaleHeight     =   345
         ScaleWidth      =   465
         TabIndex        =   17
         Top             =   1560
         Width           =   495
      End
      Begin VB.PictureBox Color 
         Appearance      =   0  'Flat
         BackColor       =   &H000040C0&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   15
         Left            =   1080
         ScaleHeight     =   345
         ScaleWidth      =   465
         TabIndex        =   16
         Top             =   1200
         Width           =   495
      End
      Begin VB.PictureBox Color 
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   14
         Left            =   1080
         ScaleHeight     =   345
         ScaleWidth      =   465
         TabIndex        =   15
         Top             =   840
         Width           =   495
      End
      Begin VB.PictureBox Color 
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   13
         Left            =   1080
         ScaleHeight     =   345
         ScaleWidth      =   465
         TabIndex        =   14
         Top             =   480
         Width           =   495
      End
      Begin VB.PictureBox Color 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   12
         Left            =   1080
         ScaleHeight     =   345
         ScaleWidth      =   465
         TabIndex        =   13
         Top             =   120
         Width           =   495
      End
      Begin VB.PictureBox Color 
         Appearance      =   0  'Flat
         BackColor       =   &H00000040&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   11
         Left            =   600
         ScaleHeight     =   345
         ScaleWidth      =   465
         TabIndex        =   12
         Top             =   1920
         Width           =   495
      End
      Begin VB.PictureBox Color 
         Appearance      =   0  'Flat
         BackColor       =   &H00000080&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   10
         Left            =   600
         ScaleHeight     =   345
         ScaleWidth      =   465
         TabIndex        =   11
         Top             =   1560
         Width           =   495
      End
      Begin VB.PictureBox Color 
         Appearance      =   0  'Flat
         BackColor       =   &H000000C0&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   9
         Left            =   600
         ScaleHeight     =   345
         ScaleWidth      =   465
         TabIndex        =   10
         Top             =   1200
         Width           =   495
      End
      Begin VB.PictureBox Color 
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   8
         Left            =   600
         ScaleHeight     =   345
         ScaleWidth      =   465
         TabIndex        =   9
         Top             =   840
         Width           =   495
      End
      Begin VB.PictureBox Color 
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   7
         Left            =   600
         ScaleHeight     =   345
         ScaleWidth      =   465
         TabIndex        =   8
         Top             =   480
         Width           =   495
      End
      Begin VB.PictureBox Color 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   6
         Left            =   600
         ScaleHeight     =   345
         ScaleWidth      =   465
         TabIndex        =   7
         Top             =   120
         Width           =   495
      End
      Begin VB.PictureBox Color 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   5
         Left            =   120
         ScaleHeight     =   345
         ScaleWidth      =   465
         TabIndex        =   6
         Top             =   1920
         Width           =   495
      End
      Begin VB.PictureBox Color 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   4
         Left            =   120
         ScaleHeight     =   345
         ScaleWidth      =   465
         TabIndex        =   5
         Top             =   1560
         Width           =   495
      End
      Begin VB.PictureBox Color 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   3
         Left            =   120
         ScaleHeight     =   345
         ScaleWidth      =   465
         TabIndex        =   4
         Top             =   1200
         Width           =   495
      End
      Begin VB.PictureBox Color 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   2
         Left            =   120
         ScaleHeight     =   345
         ScaleWidth      =   465
         TabIndex        =   3
         Top             =   840
         Width           =   495
      End
      Begin VB.PictureBox Color 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   1
         Left            =   120
         ScaleHeight     =   345
         ScaleWidth      =   465
         TabIndex        =   2
         Top             =   480
         Width           =   495
      End
      Begin VB.PictureBox Color 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   0
         Left            =   120
         ScaleHeight     =   345
         ScaleWidth      =   465
         TabIndex        =   1
         Top             =   120
         Width           =   495
      End
   End
End
Attribute VB_Name = "frmColors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExit_Click()

    Unload Me
    
End Sub

Private Sub Color_Click(Index As Integer)

    Select Case Index
        
        Case 0
            lblColor.Caption = "&&" & "H00FFFFFF"
        Case 1
            lblColor.Caption = "&&" & "H00E0E0E0"
        Case 2
            lblColor.Caption = "&&" & "H00C0C0C0"
        Case 3
            lblColor.Caption = "&&" & "H00808080"
        Case 4
            lblColor.Caption = "&&" & "H00404040"
        Case 5
            lblColor.Caption = "&&" & "H00000000"
        Case 6
            lblColor.Caption = "&&" & "H00C0C0FF"
        Case 7
            lblColor.Caption = "&&" & "H008080FF"
        Case 8
            lblColor.Caption = "&&" & "H000000FF"
        Case 9
            lblColor.Caption = "&&" & "H000000C0"
        Case 10
            lblColor.Caption = "&&" & "H00000080"
        Case 11
            lblColor.Caption = "&&" & "H00000040"
        Case 12
            lblColor.Caption = "&&" & "H00C0E0FF"
        Case 13
            lblColor.Caption = "&&" & "H0080C0FF"
        Case 14
            lblColor.Caption = "&&" & "H000080FF"
        Case 15
            lblColor.Caption = "&&" & "H000040C0"
        Case 16
            lblColor.Caption = "&&" & "H00004080"
        Case 17
            lblColor.Caption = "&&" & "H00404080"
        Case 18
            lblColor.Caption = "&&" & "H00C0FFFF"
        Case 19
            lblColor.Caption = "&&" & "H0080FFFF"
        Case 20
            lblColor.Caption = "&&" & "H0000FFFF"
        Case 21
            lblColor.Caption = "&&" & "H0000C0C0"
        Case 22
            lblColor.Caption = "&&" & "H00008080"
        Case 23
            lblColor.Caption = "&&" & "H00004040"
        Case 24
            lblColor.Caption = "&&" & "H00C0FFC0"
        Case 25
            lblColor.Caption = "&&" & "H0080FF80"
        Case 26
            lblColor.Caption = "&&" & "H0000FF00"
        Case 27
            lblColor.Caption = "&&" & "H0000C000"
        Case 28
            lblColor.Caption = "&&" & "H00008000"
        Case 29
            lblColor.Caption = "&&" & "H00004000"
        Case 30
            lblColor.Caption = "&&" & "H00FFFFC0"
        Case 31
            lblColor.Caption = "&&" & "H00FFFF80"
        Case 32
            lblColor.Caption = "&&" & "H00FFFF00"
        Case 33
            lblColor.Caption = "&&" & "H00C0C000"
        Case 34
            lblColor.Caption = "&&" & "H00808000"
        Case 35
            lblColor.Caption = "&&" & "H00404000"
        Case 36
            lblColor.Caption = "&&" & "H00FFC0C0"
        Case 37
            lblColor.Caption = "&&" & "H00FF8080"
        Case 38
            lblColor.Caption = "&&" & "H00FF0000"
        Case 39
            lblColor.Caption = "&&" & "H00C00000"
        Case 40
            lblColor.Caption = "&&" & "H00800000"
        Case 41
            lblColor.Caption = "&&" & "H00400000"
        Case 42
            lblColor.Caption = "&&" & "H00FFC0FF"
        Case 43
            lblColor.Caption = "&&" & "H00FF80FF"
        Case 44
            lblColor.Caption = "&&" & "H00FF00FF"
        Case 45
            lblColor.Caption = "&&" & "H00C000C0"
        Case 46
            lblColor.Caption = "&&" & "H00800080"
        Case 47
            lblColor.Caption = "&&" & "H00400040"

    End Select

End Sub

Private Sub CopyCode_Click(Index As Integer)

    Clipboard.SetText "&H" & Right(lblColor.Caption, 7)
    
End Sub

Private Sub Form_Load()

    SetWindowPos hWnd, conHwndTopmost, 100, 100, 295, 289, conSwpNoActivate Or conSwpShowWindow
 
End Sub

Private Sub tmrColor_Timer()

    PicColor.BackColor = "&H" & Right(lblColor.Caption, 7)
    
End Sub
