VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H80000004&
   Caption         =   "Squares"
   ClientHeight    =   5805
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   7380
   DrawWidth       =   3
   FillStyle       =   0  'Solid
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   387
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   492
   StartUpPosition =   2  'CenterScreen
   Begin MSWinsockLib.Winsock sckConnection 
      Index           =   0
      Left            =   1560
      Top             =   5160
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock sckConnect 
      Left            =   840
      Top             =   5160
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   120
      Top             =   5160
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   18
      ImageHeight     =   18
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":058A
            Key             =   "new"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":070C
            Key             =   "options"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar sbStatusBar 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   0
      Top             =   5535
      Width           =   7380
      _ExtentX        =   13018
      _ExtentY        =   476
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7382
            Text            =   "Status"
            TextSave        =   "Status"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            AutoSize        =   2
            TextSave        =   "7/02/2002"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            AutoSize        =   2
            TextSave        =   "8:32 PM"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbToolBar 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   7380
      _ExtentX        =   13018
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   2
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "New"
            Object.ToolTipText     =   "New Game"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Options"
            Object.ToolTipText     =   "Game Options"
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame2 
      Height          =   4575
      Left            =   0
      MousePointer    =   12  'No Drop
      TabIndex        =   3
      Top             =   480
      Width           =   4695
      Begin VB.PictureBox PlayField 
         AutoRedraw      =   -1  'True
         DrawWidth       =   3
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   4095
         Left            =   120
         MouseIcon       =   "frmMain.frx":088E
         MousePointer    =   12  'No Drop
         ScaleHeight     =   269
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   293
         TabIndex        =   4
         Top             =   240
         Width           =   4455
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000A&
      Height          =   4575
      Left            =   4800
      TabIndex        =   2
      Top             =   480
      Width           =   2535
      Begin VB.Frame GameFrame 
         Caption         =   "Hall of Records"
         Height          =   3495
         Left            =   120
         TabIndex        =   5
         Top             =   960
         Width           =   2295
         Begin VB.ListBox GameList 
            BackColor       =   &H00FFC0FF&
            Enabled         =   0   'False
            Height          =   3180
            ItemData        =   "frmMain.frx":1158
            Left            =   120
            List            =   "frmMain.frx":115A
            TabIndex        =   6
            Top             =   240
            Width           =   2055
         End
      End
      Begin VB.Label PlayerScores 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Batang"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   1800
         TabIndex        =   10
         Top             =   600
         Width           =   495
      End
      Begin VB.Label PlayerScores 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Batang"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   1800
         TabIndex        =   9
         Top             =   240
         Width           =   495
      End
      Begin VB.Label PlayerNames 
         Caption         =   "Player 2"
         BeginProperty Font 
            Name            =   "Batang"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   480
         TabIndex        =   8
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label PlayerNames 
         Caption         =   "Player 1"
         BeginProperty Font 
            Name            =   "Batang"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   480
         TabIndex        =   7
         Top             =   240
         Width           =   1215
      End
      Begin VB.Shape Shape2 
         FillColor       =   &H00FF0000&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   120
         Top             =   600
         Width           =   255
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   120
         Top             =   240
         Width           =   255
      End
   End
   Begin VB.Menu mnuGame 
      Caption         =   "&Game"
      Begin VB.Menu mnuGameAbort 
         Caption         =   "&Abort"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuGameDisconnect 
         Caption         =   "&Disconnect"
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuGameNew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuGameOptions 
         Caption         =   "&Options..."
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuGameExit 
         Caption         =   "E&xit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpContents 
         Caption         =   "&Contents"
      End
      Begin VB.Menu mnuHelpSearchForHelpOn 
         Caption         =   "&Search For Help On..."
      End
      Begin VB.Menu mnuHelpBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About "
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function OSWinHelp% Lib "User32" _
    Alias "WinHelpA" (ByVal hWnd&, ByVal HelpFile$, ByVal wCommand%, dwData As Any)

Private Sub Form_Load()
' When the game starts, set all the defaults.
    ' Settings for Playing Field Frame
    Me.Frame2.Move 0, 30, 338, 344
    ' Settings for Playing Field
    Me.PlayField.Move 60, 150, 4940, 4930
    Me.PlayField.BackColor = vbWhite
    ' Settings for Scores Frame
    Me.Frame1.Move 5 + Me.Frame2.Width, 30, 170, Me.Frame2.Height
    ' Settings for GameFrame
    Me.GameFrame.Move Me.Frame2.Left + 100, Me.Shape2.Top + Me.Shape2.Height + 100, _
                      2340, 4100
    ' Settings for GameList
    Me.GameList.Move Me.GameFrame.Left + 5, 200, 2120, 3900
    ' Settings for ToolBar
    Me.tbToolBar.ImageList = Me.ImageList1
    Me.tbToolBar.Buttons(NEWGAME).Image = "new"
    Me.tbToolBar.Buttons(OPTIONS).Image = "options"
    ' Settings for Form
    Me.Width = 7850
    Me.Height = 6600
    ' Disable Game->Abort menu option
    Me.mnuGameAbort.Enabled = False
    ' Disable Game->Disconnect menu option
    Me.mnuGameDisconnect.Enabled = False
    ' Turn off maximise and resize capabilities on window
    VBRemoveMenu Me, rmMaximize
    VBRemoveMenu Me, rmSize
    ' Display Hall of Records
    RefreshHall
    ' Default start game is a 6 x 6 grid:
    MaxLines = 6
    SquareSize = 46
    MaxSquares = MaxLines - 1
    FirstPoint = SquareSize
    LastPoint = MaxLines * SquareSize
    ' Default game type is vs PC
    GameType = VS
    ' Default is sound on.
    Sound = True
End Sub

Private Sub PlayField_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
' Player has clicked so is making their own move.
    OtherMove = False
    ' Process their move.
    Player_MouseDown Button, Shift, X, Y
End Sub

Private Sub sckConnect_DataArrival(ByVal bytesTotal As Long)
' Data has arrived at the client (Player 2).
    Dim sData As String

    ' Get the data.
    Me.sckConnect.GetData sData, vbString
    ' Process it.
    ProcessData sData
End Sub

Private Sub sckConnect_Close()
' The server (Player 1) has disconnected.
    Disconnection
End Sub

Private Sub sckConnection_DataArrival(Index As Integer, ByVal bytesTotal As Long)
' The server (Player 1) is receiving data from the client (Player 2).
    Dim sData As String

    ' Get the data.
    Me.sckConnection(Index).GetData sData, vbString
    ' Process it.
    ProcessData sData
End Sub

Private Sub sckConnection_Close(Index As Integer)
' The client (Player 2) has disconnected.
    Disconnection
End Sub

Private Sub sckConnection_ConnectionRequest(Index As Integer, ByVal requestID As Long)
' The server (Player 1) receives a request to connect from the client (Player 2).
' Index = listening connection, should always be 0
    If Index = LISTENER Then
        ' Create a new winsock control for the client (Player 2).
        Load Me.sckConnection(PLAYER2)
        Me.sckConnection(PLAYER2).LocalPort = TCPPort
        Me.sckConnection(PLAYER2).Accept requestID
        Me.sbStatusBar.SimpleText = "Connected to " & Me.sckConnection(LISTENER).RemoteHostIP
        Player2Connected = True
        ' 1. server (Player 1) sends it's player name to client (Player 2)
        Me.sckConnection(PLAYER2).SendData "NAME:" & Me.PlayerNames(PLAYER1).Caption
    End If
End Sub

Private Sub tbToolBar_ButtonClick(ByVal Button As MSComctlLib.Button)
' Player clicked the New Game toolbar button.
    On Error Resume Next
    Select Case Button.Key
        Case "New"
            StartNewGame
        Case "Options"
            ' display the Options dialog
            SetOptions
            frmOptions.Show vbModal, Me
    End Select
End Sub

Private Sub mnuGameAbort_Click()
' Player chose Abort from the Game menu.
    AbortGame ""
End Sub

Private Sub mnuGameDisconnect_Click()
' Someone clicked the Disconnect menu option to break a connection.
    Disconnection
End Sub

Private Sub mnuGameNew_Click()
' Start a new game.
    StartNewGame
End Sub

Private Sub mnuGameOptions_Click()
' Display the Game Options dialog box.
    ' display the Options dialog
    SetOptions
    frmOptions.Show vbModal, Me
End Sub

Private Sub mnuGameExit_Click()
' Close the game down.
    Unload Me
End Sub

Private Sub mnuHelpAbout_Click()
' Display Help->About.
    frmAbout.Show vbModal, Me
End Sub

Private Sub mnuHelpSearchForHelpOn_Click()
' Display Help Index.
    Dim nRet As Integer

    On Error Resume Next
    nRet = OSWinHelp(Me.hWnd, App.HelpFile, 261, 0)
    If Err Then
        MsgBox Err.Description
    End If
End Sub

Private Sub mnuHelpContents_Click()
' Display Help Contents.
    Dim nRet As Integer

    On Error Resume Next
    nRet = OSWinHelp(Me.hWnd, App.HelpFile, 3, 0)
    If Err Then
        MsgBox Err.Description
    End If
End Sub
