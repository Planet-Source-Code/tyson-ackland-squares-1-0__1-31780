Attribute VB_Name = "CumulativeScoreIOModule"
Option Explicit

Public Const SCOREFILE = "scores.rec" ' File name for scores
Private Const MAXNAMELEN As Integer = 20 ' Maximum length of a player name

' ScoreRec is the record block used in the Score file
Public Type ScoreRec
    Players(NUMPLAYERS) As String * MAXNAMELEN ' Player names
    Wins(NUMPLAYERS) As Integer ' Player game wins
    Games As Integer ' Total games played by the pair
End Type

Public Sub UpdateScoreFile(Player1Name As String, Player1Update As Byte, _
                           Player2Name As String, Player2Update As Byte)
' Opens the score file, seeks through it for the relevant record, updates it or creates
' it and then closes the file.
    Dim FileNum As Integer
    Dim Position As Integer
    Dim ScoreData As ScoreRec
    Dim Found As Boolean
    
    ' Open the score file
    FileNum = FreeFile
    Open SCOREFILE For Random Access Read Write As FileNum Len = Len(ScoreData)
    ' Locate the player-pair entry
    Found = False
    Position = 1
    Get FileNum, Position, ScoreData
    Do Until EOF(FileNum)
        With ScoreData
            If Player1Name = RTrim(.Players(PLAYER1)) And _
               Player2Name = RTrim(.Players(PLAYER2)) Then
                ' Found existing entry
                Found = True
                .Wins(PLAYER1) = .Wins(PLAYER1) + Player1Update
                .Wins(PLAYER2) = .Wins(PLAYER2) + Player2Update
                .Games = .Games + 1
                Put FileNum, Position, ScoreData ' Update the file
                Exit Do
            ElseIf Player1Name = RTrim(.Players(PLAYER2)) And _
                   Player2Name = RTrim(.Players(PLAYER1)) Then
                ' Found existing entry
                Found = True
                .Wins(PLAYER1) = .Wins(PLAYER1) + Player2Update
                .Wins(PLAYER2) = .Wins(PLAYER2) + Player1Update
                .Games = .Games + 1
                Put FileNum, Position, ScoreData ' Update the file
                Exit Do
            End If
        End With
        Position = Position + 1
        Get FileNum, Position, ScoreData
    Loop
    If Not Found Then
        ' We need to create a new entry
        With ScoreData
            .Players(PLAYER1) = Player1Name
            .Players(PLAYER2) = Player2Name
            .Wins(PLAYER1) = Player1Update
            .Wins(PLAYER2) = Player2Update
            .Games = 1
        End With
        Put FileNum, Position, ScoreData ' Update the file
    End If
    Close FileNum ' Close the file
End Sub
