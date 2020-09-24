Attribute VB_Name = "VsPCModule"
Option Explicit

' Movement record (x1, y1) - (x2, y2)
Public Type MoveRec
    x1 As Single
    y1 As Single
    x2 As Single
    y2 As Single
End Type

Public Function ScoringMove() As Boolean
' Look for a 3 sided square:
    Dim NumSides As Byte, i As Byte, j As Byte
    Dim x1 As Single, y1 As Single, x2 As Single, y2 As Single
    
    ScoringMove = False
    For i = 1 To MaxSquares
        For j = 1 To MaxSquares
            NumSides = 0
            If HLines(i, j) Then
                NumSides = NumSides + 1
            Else
                ' This might be the missing 4th side so remember it for the computer move!
                x1 = i * SquareSize
                y1 = j * SquareSize
                x2 = (i + 1) * SquareSize
                y2 = y1
            End If
            If HLines(i, j + 1) Then
                NumSides = NumSides + 1
            Else
                ' This might be the missing 4th side so remember it for the computer move!
                x1 = i * SquareSize
                y1 = (j + 1) * SquareSize
                x2 = (i + 1) * SquareSize
                y2 = y1
            End If
            If VLines(i, j) Then
                NumSides = NumSides + 1
            Else
                ' This might be the missing 4th side so remember it for the computer move!
                x1 = i * SquareSize
                y1 = j * SquareSize
                x2 = x1
                y2 = (j + 1) * SquareSize
            End If
            If VLines(i + 1, j) Then
                NumSides = NumSides + 1
            Else
                ' This might be the missing 4th side so remember it for the computer move!
                x1 = (i + 1) * SquareSize
                y1 = j * SquareSize
                x2 = x1
                y2 = (j + 1) * SquareSize
            End If
            ' Did we find one?
            If NumSides = 3 Then
                fMainForm.PlayField.Line (x1, y1)-(x2, y2), vbGreen ' Draw the 4th line of the square!
                fMainForm.PlayField.Refresh
                If Sound Then
                    sndPlaySound32 PENCIL, 0
                End If
                ' Remember all the lines in the square.  3 are already known but it will
                ' be faster to just set them all again blindly then to code looking for
                ' the one which is unset and then set it.
                ' Top Wall
                HLines(i, j) = True
                ' Bottom Wall
                HLines(i, j + 1) = True
                ' Left Wall
                VLines(i, j) = True
                ' Right Wall
                VLines(i + 1, j) = True
                ' Record the move.
                Moves = Moves - 1
                ScoringMove = True
                Exit Function ' We only want to score one square at a time.
            End If
        Next j
    Next i
End Function

Public Function SidesBelow(i As Byte, j As Byte) As Byte
' Returns the number of sides already drawn for the square below HLines(i,j).
    SidesBelow = 0
    ' Top horizontal.
    If HLines(i, j) Then
        SidesBelow = SidesBelow + 1
    End If
    ' Bottom horizontal.
    If HLines(i, j + 1) Then
        SidesBelow = SidesBelow + 1
    End If
    ' Left vertical.
    If VLines(i, j) Then
        SidesBelow = SidesBelow + 1
    End If
    ' Right vertical.
    If VLines(i + 1, j) Then
        SidesBelow = SidesBelow + 1
    End If
End Function

Public Function SidesAbove(i As Byte, j As Byte) As Byte
' Returns the number of sides already drawn for the square above HLines(i,j).
    SidesAbove = 0
    ' Bottom horizontal.
    If HLines(i, j) Then
        SidesAbove = SidesAbove + 1
    End If
    ' Top horizontal.
    If HLines(i, j - 1) Then
        SidesAbove = SidesAbove + 1
    End If
    ' Left vertical.
    If VLines(i, j - 1) Then
        SidesAbove = SidesAbove + 1
    End If
    ' Right vertical.
    If VLines(i + 1, j - 1) Then
        SidesAbove = SidesAbove + 1
    End If
End Function

Public Function SidesRight(i As Byte, j As Byte) As Byte
' Returns the number of sides already drawn for the square to the right of VLines(i,j).
    SidesRight = 0
    ' Left vertical.
    If VLines(i, j) Then
        SidesRight = SidesRight + 1
    End If
    ' Right vertical.
    If VLines(i + 1, j) Then
        SidesRight = SidesRight + 1
    End If
    ' Top horizontal.
    If HLines(i, j) Then
        SidesRight = SidesRight + 1
    End If
    ' Bottom horizontal.
    If HLines(i, j + 1) Then
        SidesRight = SidesRight + 1
    End If
End Function

Public Function SidesLeft(i As Byte, j As Byte) As Byte
' Returns the number of sides already drawn for the square to the right of VLines(i,j).
    SidesLeft = 0
    ' Right vertical.
    If VLines(i, j) Then
        SidesLeft = SidesLeft + 1
    End If
    ' Left vertical.
    If VLines(i - 1, j) Then
        SidesLeft = SidesLeft + 1
    End If
    ' Top horizontal.
    If HLines(i - 1, j) Then
        SidesLeft = SidesLeft + 1
    End If
    ' Bottom horizontal.
    If HLines(i - 1, j + 1) Then
        SidesLeft = SidesLeft + 1
    End If
End Function

Public Sub ComputerMove()
' Computer Player moves.  Returns True if they scored and so get another turn.
' 1. Look for scoring moves ie squares with 3 borders.
' 2. Choose a random move.  Only proviso here is do not create a 3rd border of a square
'    if we can help it.
    Dim i As Byte, j As Byte ' array looping counters.
    Dim Ignore As Boolean ' Needed to accept the results of CompleteSquares.
    Dim ComputerChoice As Integer ' Which move the computer will make.
    Dim NumSides As Byte ' How many sides are already drawn on this square?
    Dim NumSafeMoves As Integer ' How many safe moves have we found?
    Dim NumPossibleMoves As Integer ' How many possible moves have we found?
    Dim SafeMoves(264) As MoveRec ' Array of safe moves.
    Dim PossibleMoves(264) As MoveRec ' Array of possible moves.

    DeactivateField vbHourglass
    Sleep HALFASEC
    RedrawLines ' Redraw entire grid of lines to remove green line.
    ' Check for a scoring move first.
    If ScoringMove Then
        Ignore = CompleteSquares
    Else
        ' If there were no scoring moves, this time go through all the lines and
        ' choose a random one that is part of a square that has only 1 side already
        ' drawn.  Our move will then make the second side so that the next player won't
        ' be able to score from it.
        NumSafeMoves = 0
        NumPossibleMoves = 0
        ' Check horizontal lines first.
        For i = 1 To MaxSquares
            For j = 1 To MaxLines
                ' Don't bother analysing previously made moves.
                If Not HLines(i, j) Then
                    ' Add this move to the array of possible moves.
                    With PossibleMoves(NumPossibleMoves)
                        .x1 = i * SquareSize
                        .y1 = j * SquareSize
                        .x2 = (i + 1) * SquareSize
                        .y2 = .y1
                    End With
                    NumPossibleMoves = NumPossibleMoves + 1
                    If j = 1 Then
                        ' We are checking the first row of horizontal lines.
                        NumSides = SidesBelow(i, j)
                    ElseIf j = MaxLines Then
                        ' We are checking the last row of horizontal lines.
                        NumSides = SidesAbove(i, j)
                    Else
                        ' We are checking the remaining rows of horizontal lines.
                        NumSides = SidesAbove(i, j)
                        ' Only bother checking the square below if the square above is also
                        ' a safe move.
                        If NumSides <= 1 Then
                            NumSides = SidesBelow(i, j)
                        End If
                    End If
                    If NumSides <= 1 Then
                        ' Add this move to the list of safe moves.
                        With SafeMoves(NumSafeMoves)
                            .x1 = i * SquareSize
                            .y1 = j * SquareSize
                            .x2 = (i + 1) * SquareSize
                            .y2 = .y1
                        End With
                        NumSafeMoves = NumSafeMoves + 1
                    End If
                End If
            Next j
        Next i
        ' Check vertical lines next.
        For i = 1 To MaxLines
            For j = 1 To MaxSquares
                ' Don't bother analysing previously made moves.
                If Not VLines(i, j) Then
                    ' Add this move to the array of possible moves.
                    With PossibleMoves(NumPossibleMoves)
                        .x1 = i * SquareSize
                        .y1 = j * SquareSize
                        .x2 = .x1
                        .y2 = (j + 1) * SquareSize
                    End With
                    NumPossibleMoves = NumPossibleMoves + 1
                    If i = 1 Then
                        ' We are checking the first column of vertical lines.
                        NumSides = SidesRight(i, j)
                    ElseIf i = MaxLines Then
                        ' We are checking the last column of vertical lines.
                        NumSides = SidesLeft(i, j)
                    Else
                        ' We are checking the remaining columns of vertical lines.
                        NumSides = SidesLeft(i, j)
                        ' Only bother checking the square to the right if the square to
                        ' the left is also a safe move.
                        If NumSides <= 1 Then
                            NumSides = SidesRight(i, j)
                        End If
                    End If
                    If NumSides <= 1 Then
                        ' Add this move to the list of safe moves.
                        With SafeMoves(NumSafeMoves)
                            .x1 = i * SquareSize
                            .y1 = j * SquareSize
                            .x2 = .x1
                            .y2 = (j + 1) * SquareSize
                        End With
                        NumSafeMoves = NumSafeMoves + 1
                    End If
                End If
            Next j
        Next i
        If NumSafeMoves > 0 Then
            ' We found possible safe moves so choose a random one from our SafeMoves
            ' array.  Random number is between 0 and NumSafeMoves.
            ComputerChoice = Int(NumSafeMoves * Rnd)
            With SafeMoves(ComputerChoice)
                fMainForm.PlayField.Line (.x1, .y1)-(.x2, .y2), vbGreen
                If .x1 = .x2 Then
                    ' We did a vertical line so remember it.
                    If .y1 < .y2 Then
                        VLines(.x1 / SquareSize, .y1 / SquareSize) = True
                    Else
                        VLines(.x1 / SquareSize, .y2 / SquareSize) = True
                    End If
                Else
                    ' We did an horizontal line so remember it.
                    If .x1 < .x2 Then
                        HLines(.x1 / SquareSize, .y1 / SquareSize) = True
                    Else
                        HLines(.x2 / SquareSize, .y1 / SquareSize) = True
                    End If
                End If
            End With
            fMainForm.PlayField.Refresh
            If Sound Then
                sndPlaySound32 PENCIL, 0
            End If
            Moves = Moves - 1
            ' If no scoring moves have been made half-way through the game, snore!
            If Sound And Moves = MaxLines * MaxSquares And Scores(PLAYER1) = 0 And _
               Scores(PLAYER2) = 0 Then
                sndPlaySound32 BORED, 0
            End If
            SwapTurns
        ElseIf NumPossibleMoves > 0 Then
            ' No safe moves were found so choose a random move from our PossibleMoves
            ' array.  Random number is between 0 and NumPossibleMoves.
            ComputerChoice = Int(NumPossibleMoves * Rnd)
            With PossibleMoves(ComputerChoice)
                fMainForm.PlayField.Line (.x1, .y1)-(.x2, .y2), vbGreen
                If .x1 = .x2 Then
                    ' We did a vertical line so remember it.
                    If .y1 < .y2 Then
                        VLines(.x1 / SquareSize, .y1 / SquareSize) = True
                    Else
                        VLines(.x1 / SquareSize, .y2 / SquareSize) = True
                    End If
                Else
                    ' We did an horizontal line so remember it.
                    If .x1 < .x2 Then
                        HLines(.x1 / SquareSize, .y1 / SquareSize) = True
                    Else
                        HLines(.x2 / SquareSize, .y1 / SquareSize) = True
                    End If
                End If
            End With
            fMainForm.PlayField.Refresh
            If Sound Then
                sndPlaySound32 PENCIL, 0
            End If
            Moves = Moves - 1
            ' If no scoring moves have been made half-way through the game, snore!
            If Sound And Moves = MaxLines * MaxSquares And _
               Scores(PLAYER1) = 0 And Scores(PLAYER2) = 0 Then
                sndPlaySound32 BORED, 0
            End If
            SwapTurns
        End If
    End If
    ActivateField
End Sub
