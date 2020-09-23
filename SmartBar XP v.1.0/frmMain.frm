VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   0  'None
   ClientHeight    =   8370
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2055
   LinkTopic       =   "Form1"
   ScaleHeight     =   8370
   ScaleWidth      =   2055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Project1.chameleonButton Command1 
      Height          =   255
      Left            =   720
      TabIndex        =   16
      Top             =   7200
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   450
      BTYPE           =   3
      TX              =   "Hide"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmMain.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Project1.Clock Clock1 
      Height          =   1575
      Left            =   240
      TabIndex        =   18
      Top             =   6360
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   2778
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   1440
      Top             =   3840
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   4
      Top             =   5520
      Width           =   1815
   End
   Begin Project1.chameleonButton Command3 
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   4800
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   450
      BTYPE           =   3
      TX              =   "Close"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmMain.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Project1.chameleonButton Command2 
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   4440
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   450
      BTYPE           =   3
      TX              =   "Open"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmMain.frx":0038
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Project1.chameleonButton chameleonButton1 
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   5880
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   450
      BTYPE           =   3
      TX              =   "Search!"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmMain.frx":0054
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Project1.chameleonButton chameleonButton2 
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   480
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   450
      BTYPE           =   3
      TX              =   "Lock Workstation"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmMain.frx":0070
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Project1.chameleonButton chameleonButton3 
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   840
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   450
      BTYPE           =   3
      TX              =   "Log Off"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmMain.frx":008C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Project1.chameleonButton chameleonButton4 
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   1560
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   450
      BTYPE           =   3
      TX              =   "Shut Down"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmMain.frx":00A8
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Project1.chameleonButton chameleonButton5 
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   2280
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   450
      BTYPE           =   3
      TX              =   "Cascade"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmMain.frx":00C4
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Project1.chameleonButton chameleonButton6 
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   2640
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   450
      BTYPE           =   3
      TX              =   "Minimize All"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmMain.frx":00E0
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Project1.chameleonButton chameleonButton7 
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   3360
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   450
      BTYPE           =   3
      TX              =   "Tile Horizontally"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmMain.frx":00FC
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Project1.chameleonButton chameleonButton8 
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   3720
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   450
      BTYPE           =   3
      TX              =   "Tile Vertically"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmMain.frx":0118
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Project1.chameleonButton chameleonButton9 
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   3000
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   450
      BTYPE           =   3
      TX              =   "Undo Minimize All"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmMain.frx":0134
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Project1.chameleonButton chameleonButton10 
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   1200
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   450
      BTYPE           =   3
      TX              =   "Suspend"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmMain.frx":0150
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   600
      TabIndex        =   19
      Top             =   7920
      Width           =   1815
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Windows"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   720
      TabIndex        =   10
      Top             =   1920
      Width           =   735
   End
   Begin VB.Line Line7 
      BorderColor     =   &H00800000&
      X1              =   600
      X2              =   1440
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00800000&
      X1              =   600
      X2              =   1320
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Search 1vbstreet"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   480
      TabIndex        =   3
      Top             =   5160
      Width           =   1215
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00800000&
      X1              =   480
      X2              =   1680
      Y1              =   5400
      Y2              =   5400
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00800000&
      X1              =   600
      X2              =   1440
      Y1              =   4320
      Y2              =   4320
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      X1              =   600
      X2              =   1440
      Y1              =   4320
      Y2              =   4320
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "CD Drive"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   720
      TabIndex        =   0
      Top             =   4080
      Width           =   735
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      X1              =   480
      X2              =   1680
      Y1              =   5400
      Y2              =   5400
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "System"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   720
      TabIndex        =   6
      Top             =   120
      Width           =   735
   End
   Begin VB.Line Line6 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      X1              =   600
      X2              =   1320
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Line Line8 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      X1              =   600
      X2              =   1440
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Image Image1 
      Height          =   45
      Left            =   0
      Picture         =   "frmMain.frx":016C
      Stretch         =   -1  'True
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpszCommand As String, ByVal lpszReturnString As String, ByVal cchReturnLength As Long, ByVal hwndCallBack As Long) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
Private Const HWND_TOPMOST = -1
Dim H, m, s As Integer
Const toRad = 0.01745

Private Sub chameleonButton1_Click()
ShellExecute 0&, "Open", "http://www.1vbstreet.com/vb/scripts/BrowseCategoryOrSearchResults.asp?blnWorldDropDownUsed=TRUE&txtMaxNumberOfEntriesPerPage=10&blnResetAllVariables=TRUE&optSort=Alphabetical&txtCriteria=" & Text1.Text & "&lngWId=1", "", vbNullString, 1
End Sub

Private Sub chameleonButton10_Click()
SH.Suspend
End Sub

Private Sub chameleonButton2_Click()
LockWS
End Sub

Private Sub chameleonButton3_Click()
ExitWindowsEx EWX_LOGOFF, 0
End Sub

Private Sub chameleonButton4_Click()
SH.ShutdownWindows
End Sub

Private Sub chameleonButton5_Click()
SH.CascadeWindows
End Sub

Private Sub chameleonButton6_Click()
SH.MinimizeAll
End Sub

Private Sub chameleonButton7_Click()
SH.TileHorizontally
End Sub

Private Sub chameleonButton8_Click()
SH.TileVertically
End Sub

Private Sub chameleonButton9_Click()
SH.UndoMinimizeALL
End Sub

Private Sub Command1_Click()
With nid
.cbSize = Len(nid)
.hwnd = Me.hwnd
.uId = vbNull
.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
.uCallBackMessage = WM_MOUSEMOVE
.hIcon = Me.Icon
.szTip = "SmartBarXP" & vbNullChar
End With
Shell_NotifyIcon NIM_ADD, nid
Me.Hide
End Sub

Private Sub Command2_Click()
mciSendString "Set CDAudio Door Open Wait", 0&, 0&, 0&
End Sub

Private Sub Command3_Click()
mciSendString "Set CDAudio Door Closed Wait", 0&, 0&, 0&
End Sub

Private Sub Form_Load()
Dim bill As String
bill = GetTaskbarHeight
Me.Height = Screen.Height - bill
Me.Left = Screen.Width - Me.Width
Me.Top = 0
Command1.Top = Me.Height - Command1.Height - 100
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim Result, Action As Long
If Me.ScaleMode = vbPixels Then
Action = X
Else
Action = X / Screen.TwipsPerPixelX
End If
Select Case Action
Case WM_LBUTTONUP
Me.Show
Shell_NotifyIcon NIM_DELETE, nid
End Select
End Sub

Private Sub Form_Resize()
Image1.Height = Me.Height
Image1.Top = 0
SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
End Sub

Private Sub Form_Unload(Cancel As Integer)
Shell_NotifyIcon NIM_DELETE, nid
End
End Sub

Private Sub Timer1_Timer()
Label5.Caption = Date
End Sub
