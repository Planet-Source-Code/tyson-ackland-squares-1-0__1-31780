VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Options"
   ClientHeight    =   4740
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4875
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4740
   ScaleWidth      =   4875
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frSound 
      Height          =   615
      Left            =   120
      TabIndex        =   19
      Top             =   0
      Width           =   1215
      Begin VB.CheckBox chkSound 
         Caption         =   "Sound"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Value           =   1  'Checked
         Width           =   975
      End
   End
   Begin VB.OptionButton TwoPlayerMethod 
      Caption         =   "vs Computer"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   18
      Top             =   1560
      Value           =   -1  'True
      Width           =   1335
   End
   Begin VB.Frame GameSizeframe 
      Caption         =   "Game Grid Size"
      Height          =   615
      Left            =   120
      TabIndex        =   13
      Top             =   600
      Width           =   3735
      Begin VB.OptionButton GameSizeOption 
         Caption         =   "12 x 12"
         Height          =   255
         Index           =   3
         Left            =   2760
         TabIndex        =   17
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton GameSizeOption 
         Caption         =   "6 x 6"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   735
      End
      Begin VB.OptionButton GameSizeOption 
         Caption         =   "8 x 8"
         Height          =   255
         Index           =   1
         Left            =   960
         TabIndex        =   15
         Top             =   240
         Width           =   735
      End
      Begin VB.OptionButton GameSizeOption 
         Caption         =   "10 x 10"
         Height          =   255
         Index           =   2
         Left            =   1800
         TabIndex        =   14
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   960
      TabIndex        =   12
      Top             =   4320
      Width           =   735
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   4320
      Width           =   735
   End
   Begin VB.OptionButton TwoPlayerMethod 
      Caption         =   "Internet"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   4
      Top             =   2280
      Width           =   1335
   End
   Begin VB.OptionButton TwoPlayerMethod 
      Caption         =   "Share Mouse on this PC"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   1920
      Width           =   2055
   End
   Begin VB.Frame frInternetOptions 
      Caption         =   "Internet Options"
      Enabled         =   0   'False
      Height          =   1575
      Left            =   120
      TabIndex        =   2
      Top             =   2640
      Width           =   4695
      Begin VB.TextBox tbPort 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1800
         TabIndex        =   10
         Text            =   "9999"
         Top             =   1200
         Width           =   975
      End
      Begin VB.OptionButton optConnect 
         Caption         =   "Join a Game"
         Enabled         =   0   'False
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   8
         Top             =   840
         Width           =   1215
      End
      Begin VB.OptionButton optConnect 
         Caption         =   "Host a Game"
         Enabled         =   0   'False
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.TextBox tbRemoteIP 
         Enabled         =   0   'False
         Height          =   285
         Left            =   3240
         TabIndex        =   5
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label lbPort 
         Caption         =   "Enter a TCP/IP Port:"
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Label lbRemoteIP 
         Caption         =   "Remote IP Address:"
         Enabled         =   0   'False
         Height          =   255
         Left            =   1680
         TabIndex        =   6
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label lbHostIP 
         Caption         =   "Your IP Address: 000.000.000.000"
         Enabled         =   0   'False
         Height          =   255
         Left            =   2040
         TabIndex        =   3
         Top             =   360
         Width           =   2535
      End
   End
   Begin VB.Label lbTwoPlayerMethods 
      Caption         =   "Two Player Methods:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   1815
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkSound_Click()
    If Me.chkSound.Value Then
        Sound = True
    Else
        Sound = False
    End If
End Sub

Private Sub cmdCancel_Click()
' Close the Options dialog.
    Me.Hide
End Sub

Private Sub cmdOK_Click()
' Process the chosen options and close the Options dialog.
    If GameType = NET And Not Player2Connected And Pnames(PLAYER1) = "" Then
        ' Player wants Internet play
        Pnames(PLAYER1) = ""
        Pnames(PLAYER2) = ""
        TCPPort = Me.tbPort.Text
        fMainForm.mnuGameDisconnect.Enabled = True
        fMainForm.mnuGameNew.Enabled = False
        fMainForm.tbToolBar.Buttons(NEWGAME).Enabled = False
        If Me.optConnect(HOST).Value Then
            ' Player is hosting the game as Player 1
            MeServer = True
            ' Close the Winsock control that allows you to connect to the server.
            fMainForm.sckConnect.Close
            ' Reset the Winsock control that listens for connections.
            fMainForm.sckConnection(LISTENER).Close
            fMainForm.sckConnection(LISTENER).LocalPort = TCPPort
            fMainForm.sckConnection(LISTENER).Listen
            ' Update the status bar.
            fMainForm.sbStatusBar.SimpleText = "Hosting..."
            ' Prompt for server (Player 1) name
            GetPlayerName PLAYER1
            If Pnames(PLAYER1) = "" Then
                Disconnection
            Else
                fMainForm.PlayerNames(PLAYER1).Caption = Pnames(PLAYER1)
            End If
        Else
            ' Player is joining a hosted game as Player 2
            MeServer = False
            ' Prompt for client (Player 2) name
            GetPlayerName PLAYER2
            If Pnames(PLAYER2) = "" Then
                Disconnection
            Else
                fMainForm.PlayerNames(PLAYER2).Caption = Pnames(PLAYER2)
            End If
            ' Reset the Winsock control and try to connect.
            fMainForm.sckConnect.Close
            fMainForm.sckConnect.RemoteHost = Me.tbRemoteIP.Text
            fMainForm.sckConnect.RemotePort = TCPPort
            fMainForm.sckConnect.Connect
        End If
    End If
    Me.Hide
End Sub

Private Sub GameSizeOption_Click(Index As Integer)
' Set the various limits based on the player's grid size selection.
    Select Case Index
        Case 0
            ' 6 x 6 grid
            MaxLines = 6
            SquareSize = 46
        Case 1
            ' 8 x 8 grid
            MaxLines = 8
            SquareSize = 37
        Case 2
            ' 10 x 10 grid
            MaxLines = 10
            SquareSize = 30
        Case 3
            ' 12 x 12 grid
            MaxLines = 12
            SquareSize = 25
    End Select
    MaxSquares = MaxLines - 1
    FirstPoint = SquareSize
    LastPoint = MaxLines * SquareSize
End Sub

Private Sub optConnect_Click(Index As Integer)
' For Internet play only, if the player hosts, disable the join options and vice versa.
    Select Case Index
        Case HOST
            ' The player decides to host
            Me.lbHostIP.Enabled = True
            Me.lbRemoteIP.Enabled = False
            Me.tbRemoteIP.Enabled = False
            If Me.tbPort.Text = "" Then
                Me.cmdOK.Enabled = False
            Else
                Me.cmdOK.Enabled = True
            End If
        Case JOIN
            ' The player decides to join a hosted game
            Me.lbHostIP.Enabled = False
            Me.lbRemoteIP.Enabled = True
            Me.tbRemoteIP.Enabled = True
            ' Is there a remote IP address entered yet?
            If Me.tbRemoteIP.Text = "" Then
                ' No, so disable OK button.
                Me.cmdOK.Enabled = False
            Else
                ' Yes, so enable the OK button.
                If Me.tbPort.Text = "" Then
                    Me.cmdOK.Enabled = False
                Else
                    Me.cmdOK.Enabled = True
                End If
            End If
    End Select
End Sub

Private Sub tbPort_Change()
    If Me.tbPort.Text = "" Then
        Me.cmdOK.Enabled = False
    Else
        If Me.optConnect(HOST) Then
            Me.cmdOK.Enabled = True
        ElseIf Me.tbRemoteIP.Text <> "" Then
            Me.cmdOK.Enabled = True
        End If
    End If
End Sub

Private Sub tbRemoteIP_Change()
' If a value is entered in the remote IP address text box, enable the OK button
' otherwise disable it.
    If Me.tbRemoteIP.Text = "" Then
        Me.cmdOK.Enabled = False
    ElseIf Me.tbPort.Text = "" Then
        Me.cmdOK.Enabled = False
    Else
        Me.cmdOK.Enabled = True
    End If
End Sub

Private Sub TwoPlayerMethod_Click(Index As Integer)
' If non-Internet play is chosen, disable the Internet options and vice versa.
    Select Case Index
        Case SHARE
            ' Non-internet play, ie share the mouse
            GameType = SHARE
            DisableInternetOptions
            Me.cmdOK.Enabled = True
        Case NET
            ' Internet play
            Me.frInternetOptions.Enabled = True
            Me.optConnect(HOST).Enabled = True
            If Me.optConnect(HOST).Value Then
                Me.lbHostIP.Enabled = True
                If Me.tbPort.Text = "" Then
                    Me.cmdOK.Enabled = False
                Else
                    Me.cmdOK.Enabled = True
                End If
            End If
            Me.optConnect(JOIN).Enabled = True
            If Me.optConnect(JOIN).Value Then
                Me.lbRemoteIP.Enabled = True
                Me.tbRemoteIP.Enabled = True
                If Me.tbRemoteIP.Text = "" Or Me.tbPort.Text = "" Then
                    Me.cmdOK.Enabled = False
                Else
                    Me.cmdOK.Enabled = True
                End If
            End If
            Me.lbPort.Enabled = True
            Me.tbPort.Enabled = True
            GameType = NET
        Case VS
            ' vs PC play.
            GameType = VS
            DisableInternetOptions
            Me.cmdOK.Enabled = True
    End Select
End Sub
