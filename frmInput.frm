VERSION 5.00
Begin VB.Form frmInput 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Input Dialog"
   ClientHeight    =   915
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   3225
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   915
   ScaleWidth      =   3225
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox tbEntry 
      Height          =   285
      Left            =   1440
      MaxLength       =   20
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   1080
      TabIndex        =   1
      Top             =   480
      Width           =   855
   End
   Begin VB.Label lbPrompt 
      Caption         =   "Player Name:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmInput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub OKButton_Click()
' Get the player's name.
    If Me.Caption = "Enter a Name for Player 1" Then
        If Me.tbEntry.Text = "" Then
            Pnames(PLAYER1) = "Player 1"
        Else
            Pnames(PLAYER1) = Me.tbEntry.Text
        End If
    Else
        If Me.tbEntry.Text = "" Then
            Pnames(PLAYER2) = "Player 2"
        Else
            Pnames(PLAYER2) = Me.tbEntry.Text
        End If
    End If
    Unload Me
End Sub
