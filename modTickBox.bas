'==============================================================================
' Tickbox Module: modTickBox
'==============================================================================
' Description: This module demonstrates usage of the clsTickBox class module
'
'  Note: Keybinding "Tickbox_Insert" to a ctrl+key combination is recommended
'       it is bound to Ctrl+R by default
'
' Author: Matthew Snow
' Version: 1.1
' Last Modified: 23 Feb 2024
'==============================================================================

Option Explicit

Dim Tickbox As clsTickBox

Public Sub Tickbox_Click()
    ' Handles the click event for the tickbox
    
    ' Check if the tick object is already instantiated
    ' this will likely only call when the workbook is reopened
    If Tickbox Is Nothing Then
        Set Tickbox = New clsTickBox
    End If
    
    ' Call the Click method of the tick object passing the caller information
    Tickbox.Click Application.caller
End Sub

Public Sub Tickbox_Insert()
    ' Inserts a tickbox in the selected cell
    
    ' Create a new instance of clsTickBox
    Set Tickbox = New clsTickBox
    
    '====================================================================
    ' Set references to ticked and unticked shapes
    ' Default shape/s can be created using tick.CreateDefaultShape(true/false)
    ' or the user can provide a shape themselves
    '====================================================================
    Dim tickShp As Shape: Set tickShp = Tickbox.CreateDefaultShape(True)
    Dim untickShp As Shape: Set untickShp = Tickbox.CreateDefaultShape(False)

    ' Set the target cell as the currently selected cell
    Dim targetCell As Range: Set targetCell = Selection
    
    ' Create the tickbox in the target cell with the specified shapes and macro
    ' see clsTickBox for more information regarding usage
    Tickbox.Create tickShp, untickShp, targetCell, "Tickbox_Click"
End Sub
