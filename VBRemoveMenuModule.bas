Attribute VB_Name = "VBRemoveMenuModule"
Option Explicit

' Declarations necessary to work with the Control Box menus
Private Declare Function RemoveMenu Lib "User32" (ByVal hMenu As Long, _
                                                  ByVal nPosition As Long, _
                                                  ByVal wFlags As Long) As Long
Private Declare Function GetSystemMenu Lib "User32" (ByVal hWnd As Long, _
                                                     ByVal revert As Long) As Long
Private Declare Function GetWindowLong Lib "User32" _
                             Alias "GetWindowLongA" (ByVal hWnd As Long, _
                                                     ByVal lIndex As Long) As Long
Private Declare Function SetWindowLong Lib "User32" _
                             Alias "SetWindowLongA" (ByVal hWnd As Long, _
                                                     ByVal lIndex As Long, _
                                                     ByVal dwNewLong As Long) As Long
Private Const MF_BYPOSITION = &H400
Private Const MAXIMIXE_BUTTON = &HFFFEFFFF
Private Const MINIMIZE_BUTTON = &HFFFDFFFF
Private Const GWL_STYLE = (-16)

' Enumeration used when calling VBRemoveMenu
Public Enum RemoveMenuEnum
    rmMove = 1
    rmSize = 2
    rmMinimize = 3
    rmMaximize = 4
    rmClose = 6
End Enum

Public Sub VBRemoveMenu(ByVal TargetForm As Form, ByVal MenuToRemove As RemoveMenuEnum)
' This routine removes the specified menu item from the control menu and the
' corresponding functionality from the form.
'
' Parameters:
' TargetForm - the form to perform the operation on
' MenuToRemove - Enum specifying which menu to remove
    Dim hSysMenu As Long
    Dim lStyle As Long
    
    hSysMenu = GetSystemMenu(TargetForm.hWnd, 0&)
    RemoveMenu hSysMenu, MenuToRemove, MF_BYPOSITION

    Select Case MenuToRemove
        Case rmClose
            ' when removing the Close menu, also
            ' remove the separator over it
            RemoveMenu hSysMenu, MenuToRemove - 1, MF_BYPOSITION
        Case rmMinimize, rmMaximize
            ' get the current window style
            lStyle = GetWindowLong(TargetForm.hWnd, GWL_STYLE)
            If MenuToRemove = rmMaximize Then
                ' turn off bits for Maximize arrow button
                lStyle = lStyle And MAXIMIXE_BUTTON
            Else
                ' turn off bits for Minimize arrow button
                lStyle = lStyle And MINIMIZE_BUTTON
            End If
            ' set the new window style
            SetWindowLong TargetForm.hWnd, GWL_STYLE, lStyle
    End Select
End Sub
