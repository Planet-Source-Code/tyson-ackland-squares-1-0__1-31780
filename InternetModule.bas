Attribute VB_Name = "InternetModule"
Option Explicit

' Internet Game Constants
Public Const LISTENER As Integer = 0 ' Listening socket on server
Public Const HOST As Integer = 0 ' Player hosts the game.
Public Const JOIN As Integer = 1 ' Player joins the game.
Private Const COMMAND As Integer = 0 ' First parameter in all signals
Public Const LOOPBACK As String = "127.0.0.1" ' Offline IP address

' Game Variables
Public MeServer As Boolean ' Who is the server?
Public Player2Connected As Boolean ' True if Player 2 has connected.
Public TCPPort As Long ' TCP/IP port chosen
Public OtherMove As Boolean ' True if we're doing the other player's move.
Public Wins(NUMPLAYERS) As Integer ' Record of wins for each player.

Public Sub ProcessNetWinner(WinningPlayer As Byte, _
                            Player1Update As Byte, Player2Update As Byte, _
                            ServerSound As String, ClientSound As String)
' Display winner dialog, update score file, notify other player if a NET game and
' play a sound effect.
    fMainForm.Refresh
    If Sound Then
        If MeServer Then
            sndPlaySound32 ServerSound, 0
        Else
            sndPlaySound32 ClientSound, 0
        End If
    End If
    ' Show dialog.
    MsgBox fMainForm.PlayerNames(WinningPlayer).Caption & " is the winner!", vbOKOnly, "Game Over"
    If MeServer Then
        'fMainForm.mnuGameOptions.Enabled = True ' Game over so re-enable the File->Options menu.
        'fMainForm.tbToolBar.Buttons(OPTIONS).Enabled = True ' And re-enable the Options toolbar button.
        ' Update score file.
        UpdateScoreFile fMainForm.PlayerNames(PLAYER1).Caption, Player1Update, _
                        fMainForm.PlayerNames(PLAYER2).Caption, Player2Update
        ' Send the updated score to the client (Player 2).
        fMainForm.sckConnection(PLAYER2).SendData HallofRecordsStream
    End If
End Sub

Public Sub EndOfNetGame()
' Invoked if the Internet game completes.
    fMainForm.mnuGameAbort.Enabled = False ' Disable the Game->Abort option
    fMainForm.sbStatusBar.SimpleText = "Game Over" ' Update status bar.
    DeactivateField vbNoDrop ' No more mouse events wanted.
    If Scores(PLAYER1) > Scores(PLAYER2) Then
        ProcessNetWinner PLAYER1, 1, 0, WINNER, LOSER
    Else
        ProcessNetWinner PLAYER2, 0, 1, LOSER, WINNER
    End If
End Sub

Public Function HallofRecordsStream() As String
' Reads the local Hall of Records and loads it into the server's (Player 1)
' Hall of Records list.  It returns just the record concerning the
' server (Player 1) and client (Player 2) as a string.
    Dim ScoreData As ScoreRec   ' Hold each record as they are read
    Dim Position As Integer     ' Record number in the file
    Dim FileNum As Integer      ' File handle
    
    ' Open Hall of Records score file
    FileNum = FreeFile
    Open SCOREFILE For Random Access Read As FileNum Len = Len(ScoreData)
    fMainForm.GameList.Clear ' Empty the local Hall of Records list
    HallofRecordsStream = "RECORD:"
    ' Now read through the Hall of Records score file
    Position = 1
    Get FileNum, Position, ScoreData
    Do Until EOF(FileNum)
        With ScoreData
            ' Place entry in local Hall of Records list
            fMainForm.GameList.AddItem RTrim(.Players(PLAYER1)) & " (" & .Wins(PLAYER1) & ") vs " & _
                                       RTrim(.Players(PLAYER2)) & " (" & .Wins(PLAYER2) & ")", Position - 1
            If RTrim(.Players(PLAYER1)) = fMainForm.PlayerNames(PLAYER1).Caption And _
               RTrim(.Players(PLAYER2)) = fMainForm.PlayerNames(PLAYER2).Caption Or _
               RTrim(.Players(PLAYER1)) = fMainForm.PlayerNames(PLAYER2).Caption And _
               RTrim(.Players(PLAYER2)) = fMainForm.PlayerNames(PLAYER1).Caption Then
                ' Add entry to stream string
                HallofRecordsStream = HallofRecordsStream & RTrim(.Players(PLAYER1)) & _
                                      " (" & .Wins(PLAYER1) & ") vs " & _
                                      RTrim(.Players(PLAYER2)) & " (" & .Wins(PLAYER2) & _
                                      ")"
                ' Load scores into RAM
                Wins(PLAYER1) = .Wins(PLAYER1)
                Wins(PLAYER2) = .Wins(PLAYER2)
            End If
        End With
        Position = Position + 1
        Get FileNum, Position, ScoreData
    Loop
    Close (FileNum)
End Function

Public Sub InitialiseNetGame()
' Initialise everything required for playing a new game over the net.
    ResetArrays ' Initialise all counts
    ' Initialise player scores
    ResetScore PLAYER1
    ResetScore PLAYER2
    ResetTurns ' Player 1 always goes first
    ' Set player colours - blue and red
    Pcolour(PLAYER1) = vbRed
    Pcolour(PLAYER2) = vbBlue
    fMainForm.mnuGameAbort = True ' Enable Game->Abort menu option
    ResetGrid ' Draw the grid
End Sub

Public Sub Disconnection()
' Someone clicked the Disconnect button to break a connection.
    GameType = VS ' Set default game type again.
    ' Forget the player names.
    Pnames(PLAYER1) = ""
    Pnames(PLAYER2) = ""
    fMainForm.mnuGameAbort.Enabled = False ' Disable the Game->Abort menu option.
    fMainForm.mnuGameOptions.Enabled = True ' Re-enable the Game->Options menu option.
    fMainForm.tbToolBar.Buttons(OPTIONS).Enabled = True ' Re-enable the Options toolbar button.
    fMainForm.mnuGameDisconnect.Enabled = False ' Disable the Game->Disconnect menu option
    fMainForm.mnuGameNew.Enabled = True ' Re-enable the Game->New menu option.
    fMainForm.tbToolBar.Buttons(NEWGAME).Enabled = True ' Re-enable the New Game toolbar button.
    RefreshHall ' Reload local Hall of Records file:
    If MeServer Then
        fMainForm.sckConnection(LISTENER).Close ' close listening connection
        If Player2Connected Then
            fMainForm.sckConnection(PLAYER2).Close ' close connection to client
            Unload fMainForm.sckConnection(PLAYER2) ' free up RAM used by this Winsock control
            Player2Connected = False
        End If
    Else
        fMainForm.sckConnect.Close ' close connection to server
        fMainForm.mnuGameNew.Enabled = True ' re-enable Game->New menu option
        fMainForm.tbToolBar.Buttons(NEWGAME).Enabled = True ' disable New Game toolbar button
        Player2Connected = False
    End If
    fMainForm.sbStatusBar.SimpleText = "Disconnected." ' Reflect disconnection in status bar.
    fMainForm.Refresh ' Refresh the score listing before playing sound.
    If Sound Then
        sndPlaySound32 DISCONNECT, 0
    End If
End Sub

Public Sub ProcessData(sData As String)
' This is the main signal processor for the internet game.  It takes the signal from
' the other player and extracts the command and any parameters.
' Commands are generically COMMAND:PARAMETER1:PARAMETER2:...
'
' NAME:PlayerName - for accepting the other player's name.
' RECORD:PlayerPairScore - for accepting the other player's score record for this pair.
' READY: - for accepting that the other player is ready to start a game.
' NEW:MaxLines:SquareSize:MaxSquares:FirstPoint:LastPoint - for accepting the dimensions
'                                                           of a new game.
' TURN:PlayerName - for accepting whose turn it is now.
' MOVE:X:Y - for accepting the mouse click coordinates from the other player.
' GAMEOVER: - for accepting that the other player declares the game over.
' ABORT: - for accepting that the other player has aborted the game.
'
' Order of events in a game:
' 1. server (Player 1) sends it's player name to client (Player 2)
' 2. client (Player 2) accepts name and sends it's player name to server (Player 1)
' 3. server (Player 1) accepts name and sends high score list to client (Player 2)
' 4. client (Player 2) receives score list and sends ready to server (Player 1)
' 5. server (Player 1) receives ready signal and enables the Game->New
'    menu option and toolbar button.
' 6. server (Player 1) starts a new game.  Initialise the game locally and then
'    tell the client (Player 2) to do the same.
' 7. client (Player 2) initialises the game locally and then tells the server
'    (Player 1) to have a turn.
' 8. Whoever receives this event checks they are the player whose turn it is
'    by looking at the supplied parameter.  If they are not the player, they send
'    the same signal back again.  This allows for one player getting two turns
'    in a row.  If it is this player's turn, activate their play field so they
'    can move.
    Dim Fields() As String

    ' Split the received data into it's components
    Fields = Split(sData, ":")
    Select Case Fields(COMMAND)
        Case "NAME"
            If MeServer Then
                ' 3. server (Player 1) accepts name and sends high score record to client (Player 2)
                fMainForm.PlayerNames(PLAYER2).Caption = Fields(1)
                fMainForm.sckConnection(PLAYER2).SendData HallofRecordsStream
            Else
                ' 2. client (Player 2) accepts name and sends it's player name to server (Player 1)
                fMainForm.sbStatusBar.SimpleText = "Connected to " & fMainForm.sckConnect.RemoteHostIP
                fMainForm.PlayerNames(PLAYER1).Caption = Fields(1)
                Player2Connected = True
                fMainForm.sckConnect.SendData "NAME:" & fMainForm.PlayerNames(PLAYER2).Caption
            End If
        Case "RECORD"
            ' 4. client (Player 2) receives score list and sends ready signal to
            ' server (Player 1).
            fMainForm.GameList.Clear
            fMainForm.GameList.AddItem Fields(1)
            fMainForm.sbStatusBar.SimpleText = "Waiting on " & _
                                               fMainForm.PlayerNames(PLAYER1).Caption & _
                                               " to start a game."
            fMainForm.sckConnect.SendData "READY:"
        Case "READY"
            ' 5. server (Player 1) receives ready signal and enables the Game->New
            ' menu option and toolbar button.
            fMainForm.mnuGameNew.Enabled = True
            fMainForm.tbToolBar.Buttons(NEWGAME).Enabled = True
            fMainForm.sbStatusBar.SimpleText = fMainForm.PlayerNames(PLAYER2).Caption & _
                                               " is ready to start a game."
        Case "NEW"
            ' 7. client (Player 2) initialises the game locally and then tells the server
            '    (Player 1) to have a turn.
            MaxLines = Val(Fields(1))
            SquareSize = Val(Fields(2))
            MaxSquares = Val(Fields(3))
            FirstPoint = Val(Fields(4))
            LastPoint = Val(Fields(5))
            InitialiseNetGame
            fMainForm.sckConnect.SendData "TURN:" & fMainForm.PlayerNames(Turn).Caption & ":1"
        Case "TURN"
            ' 8. Whoever receives this event checks they are the player whose turn it is
            ' by looking at the supplied parameter.  If they are not the player, they send
            ' the same signal back again.  This allows for one player getting two turns
            ' in a row.  If it is this player's turn, activate their play field so they
            ' can move.
            If MeServer Then
                If Fields(1) = fMainForm.PlayerNames(PLAYER1).Caption Then
                    ActivateField
                    If Fields(2) = "0" Then
                        SwapTurns
                    End If
                Else
                    ' Pass back to client (Player 2)
                    SwapTurns
                    fMainForm.sckConnection(PLAYER2).SendData sData
                End If
            ElseIf Fields(1) = fMainForm.PlayerNames(PLAYER2).Caption Then
                ActivateField
                SwapTurns
                fMainForm.sbStatusBar.SimpleText = Fields(1) & "'s Turn" ' Update status bar.
            Else
                ' Pass back to server (Player 1)
                SwapTurns
                fMainForm.sckConnect.SendData sData
            End If
        Case "MOVE"
            ' Process the other player's move.
            OtherMove = True
            Player_MouseDown 0, 0, Val(Fields(1)), Val(Fields(2))
        Case "GAMEOVER"
            EndOfNetGame ' The other player declares the game over.
        Case "ABORT"
            AbortGame Fields(1) ' The other player has aborted the game.
    End Select
End Sub
