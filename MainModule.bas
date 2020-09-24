Attribute VB_Name = "MainModule"
Option Explicit

' Squares v1.0
' Author: Tyson Ackland
' Date: 26th January, 2002
' email: ittybittytoes@bigfoot.com
' www: www.users.bigpond.com/typatty
' ICQ: 3588438
' Language: Visual Basic 6.0 SP5
'
' Well, I guess this is the most appropriate place for the introduction to my
' code since this project starts with a Sub Main (below).
' I am a UNIX Sys Admin born and bred, well bred mainly from my Uni days when
' we learned C and shell scripting on DECvaxen running Ultrix.  The next 12 years
' until now was spent managing UNIX networks with the Microsoft world being a
' fairly fledgling attachment to any networks I managed.  Now, the tides are
' changing and the Microsoft world is the desktop OS of choice, a defacto
' standard if you will.
' So, here I am, facing a new future and learning a new language.  Why VB I hear
' you ask?  Well, I have asked many people what the dominant language is in
' MS-based Tech Support teams and VB (or it's variants eg VBA and VBscript) was
' the answer.  By learning VB, I should come close to killing these three birds
' with one stone.
' I have always been a game fan so it is no surprise that I have chosen a game
' as my learning theme.  It really doesn't matter too much what you choose to
' do as any Windows application teaches a new developer fundamental issues eg
' the VB language itself, the IDE, intro to the Windows API, creating Help,
' file I/O, graphics and in this particular case I added something I wasn't
' already skilled in - client/server programming.  VB certainly made this
' quite easy with the Winsock control doing all the low level stuff.  Apart from
' this, I also downloaded the simplest examples of client/server TCP/IP
' applications I could find off the net - the simpler the better so I didn't
' have to wade through guff (like in this, now large, set of code) just to learn
' how to get a basic connection between two running programs.
' Finally, whenever I got stuck, my colleagues on comp.lang.basic.visual.misc
' were very quick to answer my questions which surprised me!  If you are a newbie
' don't pass by this valuable resource - when I was lost looking for Help within
' the IDE, this newsgroup sorted me out.

' Declaration necessary to use Sleep API
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
' Declaration necessary to use sndPlaySound32 API
Declare Function sndPlaySound32 Lib "winmm.dll" Alias "sndPlaySoundA" _
    (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

' Game Constants
Public Const COLOURCURSOR As Integer = 101 ' Cursor for play area
Public Const PLAYER1 As Integer = 0 ' Player 1
Public Const PLAYER2 As Integer = 1 ' Player 2
Public Const NUMPLAYERS As Integer = 2 ' Number of players
Public Const SHARE As Integer = 0 ' Shared mouse two player game.
Public Const NET As Integer = 1 ' Internet two player game.
Public Const VS As Integer = 2 ' vs the PC game.
Public Const NEWGAME As Integer = 1 ' New button on toolbar.
Public Const OPTIONS As Integer = 2 ' Options button on toolbar.
Public Const HALFASEC As Integer = 500 ' Half a second in milliseconds.
Public Const ASEC As Integer = 1000 ' One second in milliseconds.
Public Const ABORT As String = "abort.wav" ' Abort game sound.
Public Const BORED As String = "snore.wav" ' Game snores from boredom.
Public Const DISCONNECT As String = "disconnect.wav" ' Disconnection sound.
Public Const INTRO As String = "splash.wav" ' Splash screen sound.
Public Const LOSER As String = "loser.wav" ' Losing sound.
Public Const PCLOSS As String = "scream.wav" ' Computer player screams.
Public Const PCWIN As String = "cackle.wav" ' Computer player giggles.
Public Const PENCIL As String = "scratch.wav" ' Pencil sound.
Public Const SQUARE As String = "square.wav" ' Drew a square sound.
Public Const WINNER As String = "winner.wav" ' Winning sound.

' Game Variables
Public fMainForm As frmMain
Public MaxLines As Integer 'We will use a MaxLines by MaxLines grid.
Public SquareSize As Integer ' Space between each grid point.
Public MaxSquares As Integer ' There will always be MaxLines - 1 by MaxLines - 1 squares.
Public FirstPoint As Integer ' Coordinates of the first grid point (SquareSize, SquareSize)
Public LastPoint As Integer ' Coordinates of the last grid point.
Public HLines(1 To 12, 1 To 12) As Boolean ' Stores the horizontal lines drawn
Public VLines(1 To 12, 1 To 12) As Boolean ' Stores the vertical lines drawn
Public Pnames(2) As String ' Remember input names.
Public SquaresArray(1 To 11, 1 To 11) As Boolean ' Stores the squares already completed
Public GameType As Integer ' Type of game - SHARE, NET or VS
Public Sound As Boolean ' Does the player want sound or not?
Public Pcolour(NUMPLAYERS) As Long ' Player colours
Public Scores(NUMPLAYERS) As Integer ' Keep track of players' scores
Public Turn As Byte ' Whose turn is it?
Public Moves As Integer ' Number of moves left in the game

Public Sub ActivateField()
' Bring the play field back to life and set the mouse cursor.
    fMainForm.PlayField.Enabled = True ' Set play field to accept events.
    ' Change cursor to pencil.
    fMainForm.PlayField.MousePointer = vbCustom
    fMainForm.PlayField.MouseIcon = LoadResPicture(COLOURCURSOR, vbResCursor)
    fMainForm.PlayField.Refresh ' Refresh the play field.
End Sub

Public Sub DeactivateField(MouseShape As Integer)
' Kill the play field and set the mouse cursor.
    fMainForm.PlayField.Enabled = False ' Disable any more activity on the play field
    fMainForm.PlayField.MousePointer = MouseShape ' Change mouse icon
    fMainForm.Refresh ' Refresh the play field.
End Sub

Public Sub CheckEndOfGame(ValidMove As Boolean, Scored As Boolean)
' Is the game over?
    If Moves > 0 Then
        ' If they scored, they get another go!
        If ValidMove And Not Scored Then
            SwapTurns
            If Moves = 0 Then
                EndGame
            End If
        End If
    Else
        EndGame
    End If
End Sub

Public Sub DisableInternetOptions()
' Disable the Internet options in Options dialog.
    frmOptions.frInternetOptions.Enabled = False
    frmOptions.optConnect(HOST).Enabled = False
    frmOptions.lbHostIP.Enabled = False
    frmOptions.optConnect(JOIN).Enabled = False
    frmOptions.lbRemoteIP.Enabled = False
    frmOptions.tbRemoteIP.Enabled = False
    frmOptions.lbPort.Enabled = False
    frmOptions.tbPort.Enabled = False
End Sub

Public Sub AbortGame(PlayerName As String)
' Stop the game, someone has aborted it.
    fMainForm.mnuGameAbort.Enabled = False ' Disable Game->Abort menu option.
    fMainForm.sbStatusBar.SimpleText = "Game Aborted" ' Update status bar.
    DeactivateField vbNoDrop
    If Sound Then
        sndPlaySound32 ABORT, 0
    End If
    If GameType = NET Then
        If MeServer Then
            fMainForm.mnuGameNew.Enabled = True ' Re-enable Game->New menu option.
            fMainForm.tbToolBar.Buttons(NEWGAME).Enabled = True ' Re-enable the New Game toolbar button.
            If PlayerName = "" Then
                ' We signal the client and record who aborted.
                fMainForm.sckConnection(PLAYER2).SendData "ABORT:" & fMainForm.PlayerNames(PLAYER2).Caption
            End If
        Else
            If PlayerName = "" Then
                ' We signal the server and record who aborted.
                fMainForm.sckConnect.SendData "ABORT:" & fMainForm.PlayerNames(PLAYER1).Caption
            End If
        End If
    Else
        fMainForm.mnuGameNew.Enabled = True ' Re-enable Game->New menu option.
        fMainForm.tbToolBar.Buttons(NEWGAME).Enabled = True ' Re-enable the New Game toolbar button.
    End If
End Sub

Public Sub SwapTurns()
' It is now the other player's turn.
    If Turn = 0 Then
        Turn = 1
    Else
        Turn = 0
    End If
    fMainForm.sbStatusBar.SimpleText = fMainForm.PlayerNames(Turn).Caption & "'s Turn"
    fMainForm.PlayField.Refresh
    Do While GameType = VS And Turn = 1 And Moves > 0
        ComputerMove
    Loop
End Sub

Public Sub ProcessWinner(WinningPlayer As Byte, _
                         Player1Update As Byte, Player2Update As Byte, _
                         WAVSound As String)
' Show winner in dialog, update score file and play sound.
    ' Show winner in dialog.
    MsgBox fMainForm.PlayerNames(WinningPlayer).Caption & " is the winner!", vbOKOnly, _
           "Game Over"
    ' Update the score file.
    UpdateScoreFile fMainForm.PlayerNames(PLAYER1).Caption, Player1Update, _
                    fMainForm.PlayerNames(PLAYER2).Caption, Player2Update
    fMainForm.Refresh
    If Sound And GameType = VS Then
        sndPlaySound32 WAVSound, 0
    End If
End Sub

Public Sub EndGame()
' Game Over.
    fMainForm.mnuGameOptions.Enabled = True ' Re-enable the File->Options menu.
    fMainForm.tbToolBar.Buttons(OPTIONS).Enabled = True ' Re-enable Options toolbar button.
    fMainForm.mnuGameAbort.Enabled = False ' Disable the Game->Abort option.
    fMainForm.sbStatusBar.SimpleText = "Game Over" ' Update status bar.
    DeactivateField vbNoDrop
    fMainForm.mnuGameNew.Enabled = True ' Re-enable the Game->New menu option.
    fMainForm.tbToolBar.Buttons(NEWGAME).Enabled = True ' Re-enable the New Game toolbar button.
    ' Who won?
    If Scores(PLAYER1) > Scores(PLAYER2) Then
        ProcessWinner PLAYER1, 1, 0, PCLOSS
    Else
        ProcessWinner PLAYER2, 0, 1, PCWIN
    End If
    RefreshHall ' Display the Hall of Records list.
End Sub

Public Sub RefreshHall()
' Loads the SCOREFILE into the Hall of Records list.
    Dim ScoreData As ScoreRec ' Hold each record as they are read
    Dim Position As Integer ' Record number in the file
    Dim FileNum As Integer ' File handle
    
    ' Open score file.
    FileNum = FreeFile
    Open SCOREFILE For Random Access Read As FileNum Len = Len(ScoreData)
    ' Refresh the hall of records.
    fMainForm.GameList.Clear
    Position = 1
    Get FileNum, Position, ScoreData
    Do Until EOF(FileNum)
        With ScoreData
            fMainForm.GameList.AddItem RTrim(.Players(PLAYER1)) & " (" & .Wins(PLAYER1) & ") vs " & _
                                       RTrim(.Players(PLAYER2)) & " (" & .Wins(PLAYER2) & ")", Position - 1
        End With
        Position = Position + 1
        Get FileNum, Position, ScoreData
    Loop
    Close (FileNum) ' Close the score file.
End Sub

Sub ResetGrid()
' Draws the grid on the play field.
    Dim xpoint As Single, ypoint As Single
    
    fMainForm.PlayField.Cls ' Clear the grid.
    fMainForm.PlayField.DrawWidth = 5 ' Draw fat dots so they remain visible even as the grid becomes filled.
    For ypoint = FirstPoint To LastPoint Step SquareSize
        For xpoint = FirstPoint To LastPoint Step SquareSize
            fMainForm.PlayField.PSet (xpoint, ypoint) ' Draw a dot at this grid point.
        Next xpoint
    Next ypoint
    ' Draw thin lines so the grid points remain visible even as the grid becomes filled.
    fMainForm.PlayField.DrawWidth = 3
End Sub

Sub ResetArrays()
' Set all array elements to empty for a new game
    Dim i As Byte, j As Byte
    
    ' Set no horizontal lines drawn
    For i = 1 To MaxSquares
        For j = 1 To MaxLines
            HLines(i, j) = False
        Next j
    Next i
    ' Set no vertical lines drawn
    For i = 1 To MaxLines
        For j = 1 To MaxSquares
            VLines(i, j) = False
        Next j
    Next i
    ' Set no squares drawn
    For i = 1 To MaxSquares
        For j = 1 To MaxSquares
            SquaresArray(i, j) = False
        Next j
    Next i
End Sub

Sub ResetScore(num As Integer)
' Initialise the player's score.
    Scores(num) = 0 ' Reset score for Player num to 0.
    fMainForm.PlayerScores(num).Caption = 0 ' Display the initialised score.
End Sub

Sub ResetTurns()
' Initialise whose starts the game.
    Turn = PLAYER1 ' Player 1 always goes first.
    Moves = MaxLines * MaxSquares * 2 ' Total possible moves in a game.
    ' Update status bar with whose turn it is.
    fMainForm.sbStatusBar.SimpleText = fMainForm.PlayerNames(Turn).Caption & "'s Turn"
End Sub

Public Sub SetOptions()
    ' Set the Sound status to whatever was selected previously.
    If Sound Then
        frmOptions.chkSound.Value = 1
    Else
        frmOptions.chkSound.Value = 0
    End If
    ' Set the grid size to whatever the previously selected.
    Select Case MaxLines
        Case 6
            frmOptions.GameSizeOption(0).Value = True
        Case 8
            frmOptions.GameSizeOption(1).Value = True
        Case 10
            frmOptions.GameSizeOption(2).Value = True
        Case 12
            frmOptions.GameSizeOption(3).Value = True
    End Select
    ' Set the Game Type to whatever the previous game used.
    frmOptions.TwoPlayerMethod(GameType).Value = True
    ' Show the local IP address.
    If fMainForm.sckConnection(LISTENER).LocalIP = LOOPBACK Then
        frmOptions.lbHostIP.Caption = "You are not currently online."
    Else
        frmOptions.lbHostIP.Caption = "Your IP address is " & _
                                      fMainForm.sckConnection(LISTENER).LocalIP
    End If
    If fMainForm.mnuGameNew.Enabled Then
        ' We are not in the middle of a game.
        ' Are we currently online?
        If GameType = VS Or GameType = SHARE Then
            ' No - VS or SHARE game
            frmOptions.GameSizeframe.Enabled = True
            frmOptions.GameSizeOption(0).Enabled = True
            frmOptions.GameSizeOption(1).Enabled = True
            frmOptions.GameSizeOption(2).Enabled = True
            frmOptions.GameSizeOption(3).Enabled = True
            frmOptions.lbTwoPlayerMethods.Enabled = True
            frmOptions.TwoPlayerMethod(VS).Enabled = True
            frmOptions.TwoPlayerMethod(SHARE).Enabled = True
            If fMainForm.sckConnection(LISTENER).LocalIP = LOOPBACK Then
                frmOptions.TwoPlayerMethod(NET).Enabled = False
            Else
                frmOptions.TwoPlayerMethod(NET).Enabled = True
            End If
            frmOptions.frInternetOptions.Enabled = False
            frmOptions.optConnect(HOST).Enabled = False
            frmOptions.lbHostIP.Enabled = False
            frmOptions.optConnect(JOIN).Enabled = False
            frmOptions.lbRemoteIP.Enabled = False
            frmOptions.lbPort.Enabled = False
            frmOptions.tbPort.Enabled = False
        ElseIf Player2Connected Then
            ' Yes and we are already connected.
            frmOptions.TwoPlayerMethod(VS).Enabled = False
            frmOptions.TwoPlayerMethod(SHARE).Enabled = False
            frmOptions.TwoPlayerMethod(NET).Enabled = False
            frmOptions.TwoPlayerMethod(NET).Value = True
            ' Enable and disable the relevant options.
            If Not MeServer Then
                frmOptions.GameSizeframe.Enabled = False
                frmOptions.GameSizeOption(0).Enabled = False
                frmOptions.GameSizeOption(1).Enabled = False
                frmOptions.GameSizeOption(2).Enabled = False
                frmOptions.GameSizeOption(3).Enabled = False
            Else
                frmOptions.GameSizeframe.Enabled = True
                frmOptions.GameSizeOption(0).Enabled = True
                frmOptions.GameSizeOption(1).Enabled = True
                frmOptions.GameSizeOption(2).Enabled = True
                frmOptions.GameSizeOption(3).Enabled = True
            End If
            frmOptions.frInternetOptions.Enabled = False
            frmOptions.optConnect(HOST).Enabled = False
            frmOptions.optConnect(JOIN).Enabled = False
            frmOptions.lbHostIP.Enabled = False
            frmOptions.lbRemoteIP.Enabled = False
            frmOptions.tbRemoteIP.Enabled = False
            frmOptions.tbRemoteIP.Enabled = False
            frmOptions.lbPort.Enabled = False
            frmOptions.tbPort.Enabled = False
        Else
            ' Yes but we are not connected yet.
            frmOptions.GameSizeframe.Enabled = True
            frmOptions.GameSizeOption(0).Enabled = True
            frmOptions.GameSizeOption(1).Enabled = True
            frmOptions.GameSizeOption(2).Enabled = True
            frmOptions.GameSizeOption(3).Enabled = True
            frmOptions.lbTwoPlayerMethods.Enabled = True
            frmOptions.TwoPlayerMethod(VS).Enabled = True
            frmOptions.TwoPlayerMethod(SHARE).Enabled = True
            frmOptions.TwoPlayerMethod(NET).Enabled = True
        End If
    Else
        ' We are in the middle of a game.
        If Not MeServer Then
            frmOptions.GameSizeframe.Enabled = False
            frmOptions.GameSizeOption(0).Enabled = False
            frmOptions.GameSizeOption(1).Enabled = False
            frmOptions.GameSizeOption(2).Enabled = False
            frmOptions.GameSizeOption(3).Enabled = False
        End If
        frmOptions.lbTwoPlayerMethods.Enabled = False
        frmOptions.TwoPlayerMethod(VS).Enabled = False
        frmOptions.TwoPlayerMethod(SHARE).Enabled = False
        frmOptions.TwoPlayerMethod(NET).Enabled = False
        frmOptions.frInternetOptions.Enabled = False
        frmOptions.optConnect(HOST).Enabled = False
        frmOptions.lbHostIP.Enabled = False
        frmOptions.optConnect(JOIN).Enabled = False
        frmOptions.lbRemoteIP.Enabled = False
        frmOptions.tbRemoteIP.Enabled = False
        frmOptions.lbPort.Enabled = False
        frmOptions.tbPort.Enabled = False
    End If
End Sub

Public Sub InitialiseGame()
    ResetArrays ' Initialise all counts.
    ' Initialise player scores
    ResetScore PLAYER1
    ResetScore PLAYER2
    ' Get player names
    Pnames(PLAYER1) = ""
    Pnames(PLAYER2) = ""
    ' Get Player 1's name.
    GetPlayerName PLAYER1
    If Pnames(PLAYER1) = "" Then
        ' Pnames is only empty if player cancelled the dialog.
        AbortGame ""
        Exit Sub
    Else
        fMainForm.PlayerNames(PLAYER1).Caption = Pnames(PLAYER1)
    End If
    If GameType = VS Then
        ' Player 2 is the computer.
        fMainForm.PlayerNames(PLAYER2).Caption = "Computer"
    Else
        ' Get Player 2's name.
        GetPlayerName PLAYER2
        If Pnames(PLAYER2) = "" Then
            ' Pnames is only empty if player cancelled the dialog.
            AbortGame ""
            Exit Sub
        Else
            fMainForm.PlayerNames(PLAYER2).Caption = Pnames(PLAYER2)
        End If
    End If
    ResetTurns ' Player 1 always goes first.
    RefreshHall ' Load the Hall of Records.
    ' Set player colours - blue and red
    Pcolour(PLAYER1) = vbRed
    Pcolour(PLAYER2) = vbBlue
    fMainForm.mnuGameAbort = True ' Enable Game->Abort menu option.
    fMainForm.mnuGameNew.Enabled = False ' Disable the Game->New menu option.
    fMainForm.tbToolBar.Buttons(NEWGAME).Enabled = False ' Disable the New Game toolbar button.
    ResetGrid ' Draw the grid
    ActivateField ' Prepare to receive events.
End Sub

Public Sub StartNewGame()
' Start a new game.
    ' During a game, only sound option may still be changed.
    frmOptions.GameSizeframe.Enabled = False
    frmOptions.GameSizeOption(0).Enabled = False
    frmOptions.GameSizeOption(1).Enabled = False
    frmOptions.GameSizeOption(2).Enabled = False
    frmOptions.GameSizeOption(3).Enabled = False
    frmOptions.TwoPlayerMethod(VS).Enabled = False
    frmOptions.TwoPlayerMethod(SHARE).Enabled = False
    frmOptions.TwoPlayerMethod(NET).Enabled = False
    Select Case GameType
        Case SHARE
            InitialiseGame
        Case NET
            ' 6. server (Player 1) starts a new game.  Initialise the game locally and then
            '    tell the client (Player 2) to do the same.
            InitialiseNetGame
            fMainForm.mnuGameNew.Enabled = False ' Disable Game->New menu option.
            fMainForm.tbToolBar.Buttons(NEWGAME).Enabled = False ' Disable New Game toolbar button.
            ' tell client (Player 2) to initialise.
            fMainForm.sckConnection(PLAYER2).SendData "NEW:" & _
                                                      MaxLines & ":" & _
                                                      SquareSize & ":" & _
                                                      MaxSquares & ":" & _
                                                      FirstPoint & ":" & _
                                                      LastPoint
        Case VS
            Randomize ' Seed the random number generator with the current time.
            InitialiseGame
    End Select
End Sub

Public Sub GetPlayerName(num As Integer)
' Ask for the player's name
    frmInput.Caption = "Enter a Name for Player " & num + 1 ' Dialog title.
    frmInput.lbPrompt.Caption = "Player " & num + 1 & "'s Name:" ' Dialog prompt.
    frmInput.Show vbModal, frmOptions ' Lock main window while dialog is displayed.
End Sub

Public Function CompleteSquares() As Boolean
' Now we have to check if the player has just completed a square or two:
    Dim i As Byte, j As Byte
    
    CompleteSquares = False
    For i = 1 To MaxSquares
        For j = 1 To MaxSquares
            ' Check all four sides of the square and that it is a new square.
            If HLines(i, j) And HLines(i, j + 1) And _
               VLines(i, j) And VLines(i + 1, j) And _
               Not SquaresArray(i, j) Then
                SquaresArray(i, j) = True
                CompleteSquares = True
                Scores(Turn) = Scores(Turn) + 1
                fMainForm.PlayerScores(Turn).Caption = Scores(Turn) ' Update the displayed score.
                ' Now draw a square with a 1 pixel margin within the border lines:
                fMainForm.PlayField.Line (i * SquareSize + fMainForm.PlayField.DrawWidth + 1, _
                                j * SquareSize + fMainForm.PlayField.DrawWidth + 1)- _
                                ((i + 1) * SquareSize - fMainForm.PlayField.DrawWidth - 1, _
                                (j + 1) * SquareSize - fMainForm.PlayField.DrawWidth - 1), _
                                Pcolour(Turn), BF
                fMainForm.PlayField.Refresh
                If Sound Then
                    sndPlaySound32 SQUARE, 0
                End If
            End If
        Next j
    Next i
End Function

Public Sub RedrawLines()
' Redraws all drawn lines on the grid so far.  This is to replace the green line which
' indicated the most recent line drawn.
    Dim i As Byte, j As Byte
    
    ' Horizontal lines first:
    For i = 1 To MaxSquares
        For j = 1 To MaxLines
            If HLines(i, j) Then
                fMainForm.PlayField.Line (i * SquareSize, j * SquareSize)-((i + 1) * SquareSize, j * SquareSize), vbBlack
            End If
        Next j
    Next i
    ' Vertical lines next:
    For i = 1 To MaxLines
        For j = 1 To MaxSquares
            If VLines(i, j) Then
                fMainForm.PlayField.Line (i * SquareSize, j * SquareSize)-(i * SquareSize, (j + 1) * SquareSize), vbBlack
            End If
        Next j
    Next i
End Sub

Public Sub Player_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
' This is the primary event handler in this game.  Basically, it's function is to
' take the coordinates of the mouse click and work out which two points in the grid
' are closest to it and then draw a line between them.  The user has in mind the line
' all the time but we have to choose the correct two end-points to be able to then draw
' that line.
' Now, when calculating the distance between any two points, we use the pythagorus
' formula: c^2 = a^2 + b^2.
    Dim xpoint As Single            ' x grid point being compared
    Dim ypoint As Single            ' y grid point being compared
    Dim c As Single                 ' distance from grid point to mouse click
    Dim mindist As Single           ' distance from closest grid point to mouse click
    Dim nextmindist As Single       ' distance from next closest grid point to mouse click
    Dim minx As Single, miny As Single ' closest grid point to mouse click
    Dim xnext As Single, ynext As Single ' next closest grid point to mouse click
    Dim i As Byte, j As Byte        ' array counters
    Dim Scored As Boolean           ' Did they score?
    Dim ValidMove As Boolean        ' Did they move legally?
    
    RedrawLines ' Redraw the lines.
    ' Initialise variables
    minx = FirstPoint
    miny = FirstPoint
    mindist = 10000                  ' some arbitrary large number
    nextmindist = 10000              ' some arbitrary large number
    ValidMove = False                ' Player has not made a valid move yet.
    ' Compare every point in the grid to the mouse click to find the closest
    ' two grid points (if only computers could see!):
    For ypoint = FirstPoint To LastPoint Step SquareSize
        For xpoint = FirstPoint To LastPoint Step SquareSize
            c = Sqr((X - xpoint) ^ 2 + (Y - ypoint) ^ 2) ' Pythagorus calculation.
            If c <= mindist Then
                ' We have found a closer grid point!  Save the old point incase
                ' it is the second-closest point:
                nextmindist = mindist
                xnext = minx
                ynext = miny
                mindist = c
                minx = xpoint
                miny = ypoint
            End If
            ' We may have skipped the previous if statement if this grid point
            ' isn't the closest but it could be the second-closest:
            If c < nextmindist And c <> mindist Then
                ' We have found a new, second-closest grid point:
                nextmindist = c
                xnext = xpoint
                ynext = ypoint
            End If
        Next xpoint
    Next ypoint
    If minx = xnext Then
        ' The player selected a vertical line
        If (miny < ynext) Then
            If VLines(minx / SquareSize, miny / SquareSize) Then
                MsgBox "That move is already taken.", vbCritical, "Illegal Move"
            Else
                ValidMove = True
                Moves = Moves - 1
                fMainForm.PlayField.Line (minx, miny)-(xnext, ynext), vbGreen ' Draw the line.
                VLines(minx / SquareSize, miny / SquareSize) = True ' Remember this line.
            End If
        Else
            If VLines(minx / SquareSize, ynext / SquareSize) Then
                MsgBox "That move is already taken.", vbCritical, "Illegal Move"
            Else
                ValidMove = True
                Moves = Moves - 1
                fMainForm.PlayField.Line (minx, miny)-(xnext, ynext), vbGreen ' Draw the line.
                VLines(minx / SquareSize, ynext / SquareSize) = True ' Remember the line.
            End If
        End If
    Else
        ' The player selected an horizontal line
        If (minx < xnext) Then
            If HLines(minx / SquareSize, miny / SquareSize) Then
                MsgBox "That move is already taken.", vbCritical, "Illegal Move"
            Else
                ValidMove = True
                Moves = Moves - 1
                fMainForm.PlayField.Line (minx, miny)-(xnext, ynext), vbGreen ' Draw the line.
                HLines(minx / SquareSize, miny / SquareSize) = True ' Remember the line.
            End If
        Else
            If HLines(xnext / SquareSize, miny / SquareSize) Then
                MsgBox "That move is already taken.", vbCritical, "Illegal Move"
            Else
                ValidMove = True
                Moves = Moves - 1
                fMainForm.PlayField.Line (minx, miny)-(xnext, ynext), vbGreen ' Draw the line.
                HLines(xnext / SquareSize, miny / SquareSize) = True ' Remember the line.
            End If
        End If
    End If
    fMainForm.PlayField.Refresh
    If ValidMove Then
        If Sound Then
            sndPlaySound32 PENCIL, 0
        End If
        Scored = CompleteSquares
        ' If no scoring moves have been made half-way through the game, snore!
        If Sound And Moves = MaxLines * MaxSquares And Scores(PLAYER1) = 0 And _
           Scores(PLAYER2) = 0 Then
            sndPlaySound32 BORED, 0
        End If
        If GameType = NET Then
            If OtherMove Then
                ' We just processed our opponent's move, is it our turn now?
                If Moves > 0 Then
                    If Scored Then
                        ' Pass straight back to other player for their second turn.
                        If MeServer Then
                            fMainForm.sckConnection(PLAYER2).SendData "TURN:" & _
                                                        fMainForm.PlayerNames(PLAYER2).Caption & ":0"
                        Else
                            fMainForm.sckConnect.SendData "TURN:" & _
                                                        fMainForm.PlayerNames(PLAYER1).Caption & ":0"
                        End If
                    Else
                        ActivateField
                        SwapTurns
                    End If
                Else
                    ' Game over so tell the other player
                    If MeServer Then
                        fMainForm.sckConnection(PLAYER2).SendData "GAMEOVER:"
                    Else
                        fMainForm.sckConnect.SendData "GAMEOVER:"
                    End If
                    EndOfNetGame
                End If
            Else
                DeactivateField vbNoDrop
                SwapTurns
                ' Send the move to the other player for processing.
                If MeServer Then
                    fMainForm.sckConnection(PLAYER2).SendData "MOVE:" & Str(X) & ":" & Str(Y)
                Else
                    fMainForm.sckConnect.SendData "MOVE:" & Str(X) & ":" & Str(Y)
                End If
            End If
        Else
            CheckEndOfGame ValidMove, Scored
        End If
    End If
End Sub

Sub Main()
' This is where it all begins!
    frmSplash.Show
    frmSplash.Refresh
    Set fMainForm = New frmMain
    Load fMainForm
    
    ' Play an intro sound.
    sndPlaySound32 INTRO, 0
    
    ' Give 'em a chance to see the splash screen
    Sleep ASEC
    Unload frmSplash
    
    fMainForm.Show
End Sub
