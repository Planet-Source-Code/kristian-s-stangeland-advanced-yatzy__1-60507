VERSION 5.00
Begin VB.UserControl usrBoard 
   AutoRedraw      =   -1  'True
   ClientHeight    =   6255
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4245
   ScaleHeight     =   417
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   283
End
Attribute VB_Name = "usrBoard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Copyright (C) 2004 Kristian. S.Stangeland

'This program is free software; you can redistribute it and/or
'modify it under the terms of the GNU General Public License
'as published by the Free Software Foundation; either version 2
'of the License, or (at your option) any later version.

'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'GNU General Public License for more details.

'You should have received a copy of the GNU General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place - Suite 330, Boston, MA  02111-1307, USA.

' Class holding all the lines
Public WithEvents Lines As clsDatabase
Attribute Lines.VB_VarHelpID = -1

Private Type CellPos
    LineIndex As Long
    CellIndex As Long
End Type

' Public events
Public Event Paint(bIgnore As Boolean)
Public Event Redrawed()
Public Event Resize()
Public Event CellMouseOver(Button As Integer, Shift As Integer, LineIndex As Long, CellIndex As Long)
Public Event CellMouseUp(Button As Integer, Shift As Integer, LineIndex As Long, CellIndex As Long)
Public Event CellMouseDown(Button As Integer, Shift As Integer, LineIndex As Long, CellIndex As Long)

' Public variables
Public IgnoreDrawing As Boolean

' Constants needed for the drawing
Private pCellHeight As Long
Private pCellFrame As Long
Private pFrameWidth As Long
Private pFrameHeight As Long
Private pVariableSize As Boolean
Private pAutoResize As Boolean

Property Get AutoResize() As Boolean
    
    AutoResize = pAutoResize

End Property

Property Let AutoResize(ByVal lNewValue As Boolean)

    pAutoResize = lNewValue

End Property

Property Get VariableSize() As Boolean
    
    VariableSize = pVariableSize

End Property

Property Let VariableSize(ByVal lNewValue As Boolean)

    pVariableSize = lNewValue

End Property

Property Get CellHeight() As Long
    
    CellHeight = pCellHeight

End Property

Property Let CellHeight(ByVal lNewValue As Long)

    pCellHeight = lNewValue

End Property

Property Get CellFrame() As Long
    
    pCellFrame = pCellFrame

End Property

Property Let CellFrame(ByVal lNewValue As Long)

    pCellFrame = lNewValue

End Property

Property Get FrameWidth() As Long
    
    FrameWidth = pFrameWidth

End Property

Property Let FrameWidth(ByVal lNewValue As Long)

    pFrameWidth = lNewValue

End Property

Property Get FrameHeight() As Long
    
    FrameHeight = pFrameHeight

End Property

Property Let FrameHeight(ByVal lNewValue As Long)

    pFrameHeight = lNewValue

End Property

Private Sub Lines_CellChange(LineIndex As Long, CellIndex As Long, RowChanged As Boolean)

    ' Just redraw
    DrawAll

End Sub

Private Sub Lines_LineChange(LineIndex As Long, RowChanged As Boolean)

    ' Just redraw
    DrawAll

End Sub

Private Sub UserControl_Initialize()

    ' Create the database
    Set Lines = New clsDatabase
    
End Sub

Public Sub DrawAll()

    Dim X As Long, Y As Long, LineIndex As Long, CellIndex As Long, bIgnore As Boolean
    
    ' If the control is set to ignore drawings
    If IgnoreDrawing Then
        Exit Sub
    End If
        
    ' Inform about the control ready to redraw
    RaiseEvent Paint(bIgnore)
    
    ' If the user requested the operation to be ignored, do so
    If bIgnore Then
        Exit Sub
    End If
    
    ' Now, alot of the following code may result in the drawall-function to be called again - ignore those calls
    IgnoreDrawing = True
    
    ' Clear control
    UserControl.Cls
    
    ' Start with the frame height
    Y = FrameHeight
    
    ' Go through all elements
    For LineIndex = 0 To Lines.LineCount - 1
        
        ' Start with the frame width
        X = FrameWidth
        
        ' Clear the line rect
        Lines.LineWidth(LineIndex) = 0
        Lines.LineHeight(LineIndex) = 0
        
        Select Case Lines.LineType(LineIndex)
        Case 0 ' If it's a normal type, draw all cells
        
            For CellIndex = 0 To Lines.CellCount(LineIndex) - 1
        
                ' Draw the cell
                X = X + DrawCell(LineIndex, CellIndex, X, Y) + Lines.CellSpace(LineIndex, CellIndex)
        
            Next
            
            ' Take into account that the cells has a certain height
            Y = Y + CellHeight
            
            ' Increse line height
            Lines.LineHeight(LineIndex) = CellHeight

        End Select
        
        ' Increse the Y-position
        Y = Y + Lines.LineSpace(LineIndex)
        
    Next
        
    ' Set variables back to normal
    IgnoreDrawing = False
    
    ' We are finished redrawing
    RaiseEvent Redrawed

    ' Automatically give resize requests
    If pAutoResize Then
        UserControl.Width = (X + 1) * Screen.TwipsPerPixelX
        UserControl.Height = (Y + 1) * Screen.TwipsPerPixelY
    End If

End Sub

Public Function DrawCell(LineIndex As Long, CellIndex As Long, X As Long, Y As Long) As Long

    Dim CellWidth As Long, CellText As String

    Select Case Lines.CellType(LineIndex, CellIndex)
    Case 0 ' Only draw the cell if it is of a normal type
    
        ' Get cell text
        CellText = Lines.CellText(LineIndex, CellIndex)
    
        ' Get the cell size
        If pVariableSize Then
            ' In this case we calculate the size of each element dynamically
            CellWidth = UserControl.TextWidth(CellText) + (CellFrame * 2)
        Else
            ' Here we just harvest a static size and use it
            CellWidth = Lines.CellWidth(LineIndex, CellIndex)
        End If
        
        ' Draw the background
        UserControl.Line (X, Y)-(X + CellWidth, Y + CellHeight), Lines.CellBackground(LineIndex, CellIndex), BF
    
        ' Draw the box surounding the cell, if needed
        If Lines.CellBorder(LineIndex, CellIndex) Then
        
            ' Draw the box
            UserControl.Line (X, Y)-(X + CellWidth - 1, Y + CellHeight - 1), Lines.CellBorderColorDark(LineIndex, CellIndex), B
            UserControl.Line (X, Y + CellHeight - 1)-(X + CellWidth, Y + CellHeight - 1), Lines.CellBorderColorLight(LineIndex, CellIndex)
            UserControl.Line (X + CellWidth - 1, Y)-(X + CellWidth - 1, Y + CellHeight - 1), Lines.CellBorderColorLight(LineIndex, CellIndex)
        
        End If
        
        ' Draw text
        UserControl.CurrentX = X + pCellFrame
        UserControl.CurrentY = Y + (pCellHeight / 2) - (UserControl.TextHeight(CellText) / 2)
        UserControl.Print CellText
    
        ' Set information
        If pVariableSize = False Then
            Lines.CellWidth(LineIndex, CellIndex) = CellWidth
        End If
    
        ' Return the size
        DrawCell = CellWidth
    
    End Select

End Function

Private Function FindCellByPos(lngX As Single, lngY As Single) As CellPos

    Dim LineIndex As Long, CellIndex As Long, CellWidth As Long, X As Long, Y As Long
    
    ' Start with the frame height
    Y = FrameHeight
    
    ' Go through all elements
    For LineIndex = 0 To Lines.LineCount - 1
    
        ' Start with the frame width
        X = FrameWidth
        
        Select Case Lines.LineType(LineIndex)
        Case 0 ' If it's a normal type
        
            ' Find the cell
            For CellIndex = 0 To Lines.CellCount(LineIndex) - 1
            
                ' Get the width of the cell
                CellWidth = Lines.CellWidth(LineIndex, CellIndex)
            
                ' See if the position is inside this cell
                If lngX >= X And lngX <= X + CellWidth And lngY >= Y And lngY <= Y + CellHeight Then
                    FindCellByPos.LineIndex = LineIndex
                    FindCellByPos.CellIndex = CellIndex
                    
                    ' Nothing else to do
                    Exit Function
                End If
                
                ' Increse the X-position
                X = X + IIf(Lines.CellType(LineIndex, CellIndex) = 0, CellWidth, 0) + Lines.CellSpace(LineIndex, CellIndex)
            
            Next
            
            ' Take into account that this line has a hight
            Y = Y + CellHeight
            
        End Select
            
        ' Increse the Y-position
        Y = Y + Lines.LineSpace(LineIndex)
        
    Next
    
    ' No results
    FindCellByPos.LineIndex = -1

End Function

Private Sub MouseEvent(lngEvent As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)

    Dim CellPos As CellPos
    
    CellPos = FindCellByPos(X, Y)
    
    ' Proceed only if we have found a cell
    If CellPos.LineIndex >= 0 Then
    
        ' Invoke event
        Select Case lngEvent
        Case 0: RaiseEvent CellMouseDown(Button, Shift, CellPos.LineIndex, CellPos.CellIndex)
        Case 1: RaiseEvent CellMouseUp(Button, Shift, CellPos.LineIndex, CellPos.CellIndex)
        Case 2: RaiseEvent CellMouseOver(Button, Shift, CellPos.LineIndex, CellPos.CellIndex)
        End Select
        
    End If

End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    ' Test event
    MouseEvent 0, Button, Shift, X, Y

End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    ' Test event
    MouseEvent 1, Button, Shift, X, Y

End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    ' Test event
    ' MouseEvent 2, Button, Shift, X, Y

End Sub

Private Sub UserControl_Resize()

    ' The control has been resized
    RaiseEvent Resize

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    ' Load saved properties
    pCellHeight = PropBag.ReadProperty("CellHeight", 17)
    pCellFrame = PropBag.ReadProperty("CellFrame", 2)
    pFrameWidth = PropBag.ReadProperty("FrameWidth", 0)
    pFrameHeight = PropBag.ReadProperty("FrameHeight", 0)
    pVariableSize = PropBag.ReadProperty("VariableSize", True)
    pAutoResize = PropBag.ReadProperty("AutoResize", False)

End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    ' Save properties
    PropBag.WriteProperty "CellHeight", pCellHeight
    PropBag.WriteProperty "CellFrame", pCellFrame
    PropBag.WriteProperty "FrameWidth", pFrameWidth
    PropBag.WriteProperty "FrameHeight", pFrameHeight
    PropBag.WriteProperty "VariableSize", pVariableSize
    PropBag.WriteProperty "AutoResize", pAutoResize

End Sub


