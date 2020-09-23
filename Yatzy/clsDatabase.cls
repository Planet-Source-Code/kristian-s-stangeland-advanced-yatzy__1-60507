VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDatabase"
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

Private Type Cell
    Text As String
    Type As Long
    Background As Long
    BorderColorLight As Long
    BorderColorDark As Long
    Space As Long
    Width As Long
    Border As Boolean
    Enabled As Boolean
End Type

Private Type Line
    Cells() As Cell
    CellCount As Long
    LineType As Long
    LineSpace As Long
    LineWidth As Long
    LineHeight As Long
End Type

' Public events
Public Event CellChange(LineIndex As Long, CellIndex As Long, RowChanged As Boolean)
Public Event LineChange(LineIndex As Long, RowChanged As Boolean)

' The private array holding all the cells and rows
Private Lines() As Line

' The amout of lines
Public LineCount As Long
Public Parent As Object

Public Property Get CellText(ByVal Index As Long, ByVal CellIndex As Long) As String

    If IsCell(Index, CellIndex) Then
        CellText = Lines(Index).Cells(CellIndex).Text
    End If
    
End Property

Public Property Let CellText(ByVal Index As Long, ByVal CellIndex As Long, ByVal vNewText As String)

    If IsCell(Index, CellIndex) Then
        Lines(Index).Cells(CellIndex).Text = vNewText
        
        ' The cell has been changed
        RaiseEvent CellChange(Index, CellIndex, False)
    End If
    
End Property

Public Property Get CellEnabled(ByVal Index As Long, ByVal CellIndex As Long) As Boolean

    If IsCell(Index, CellIndex) Then
        CellEnabled = Lines(Index).Cells(CellIndex).Enabled
    End If
    
End Property

Public Property Let CellEnabled(ByVal Index As Long, ByVal CellIndex As Long, ByVal vNewEnabled As Boolean)

    If IsCell(Index, CellIndex) Then
        Lines(Index).Cells(CellIndex).Enabled = vNewEnabled
        
        ' The cell has been changed
        RaiseEvent CellChange(Index, CellIndex, False)
    End If
    
End Property

Public Property Get CellSpace(ByVal Index As Long, ByVal CellIndex As Long) As Long

    If IsCell(Index, CellIndex) Then
        CellSpace = Lines(Index).Cells(CellIndex).Space
    End If
    
End Property

Public Property Let CellSpace(ByVal Index As Long, ByVal CellIndex As Long, ByVal vNewCellSpace As Long)

    If IsCell(Index, CellIndex) Then
        Lines(Index).Cells(CellIndex).Space = vNewCellSpace
        
        ' The cell has been changed
        RaiseEvent CellChange(Index, CellIndex, False)
    End If
    
End Property

Public Property Get CellType(ByVal Index As Long, ByVal CellIndex As Long) As Long

    If IsCell(Index, CellIndex) Then
        CellType = Lines(Index).Cells(CellIndex).Type
    End If
    
End Property

Public Property Let CellType(ByVal Index As Long, ByVal CellIndex As Long, ByVal vNewCellType As Long)

    If IsCell(Index, CellIndex) Then
        Lines(Index).Cells(CellIndex).Type = vNewCellType
        
        ' The cell has been changed
        RaiseEvent CellChange(Index, CellIndex, False)
    End If
    
End Property

Public Property Get CellBackground(ByVal Index As Long, ByVal CellIndex As Long) As Long

    If IsCell(Index, CellIndex) Then
        CellBackground = Lines(Index).Cells(CellIndex).Background
    End If
    
End Property

Public Property Let CellBackground(ByVal Index As Long, ByVal CellIndex As Long, ByVal vNewCellBackground As Long)

    If IsCell(Index, CellIndex) Then
        Lines(Index).Cells(CellIndex).Background = vNewCellBackground
        
        ' The cell has been changed
        RaiseEvent CellChange(Index, CellIndex, False)
    End If
    
End Property

Public Property Get CellBorderColorLight(ByVal Index As Long, ByVal CellIndex As Long) As Long

    If IsCell(Index, CellIndex) Then
        CellBorderColorLight = Lines(Index).Cells(CellIndex).BorderColorLight
    End If
    
End Property

Public Property Let CellBorderColorLight(ByVal Index As Long, ByVal CellIndex As Long, ByVal vNewCellBorderColor As Long)

    If IsCell(Index, CellIndex) Then
        Lines(Index).Cells(CellIndex).BorderColorLight = vNewCellBorderColor
        
        ' The cell has been changed
        RaiseEvent CellChange(Index, CellIndex, False)
    End If
    
End Property

Public Property Get CellBorderColorDark(ByVal Index As Long, ByVal CellIndex As Long) As Long

    If IsCell(Index, CellIndex) Then
        CellBorderColorDark = Lines(Index).Cells(CellIndex).BorderColorDark
    End If
    
End Property

Public Property Let CellBorderColorDark(ByVal Index As Long, ByVal CellIndex As Long, ByVal vNewCellBorderColor As Long)

    If IsCell(Index, CellIndex) Then
        Lines(Index).Cells(CellIndex).BorderColorDark = vNewCellBorderColor
        
        ' The cell has been changed
        RaiseEvent CellChange(Index, CellIndex, False)
    End If
    
End Property

Public Property Get CellWidth(ByVal Index As Long, ByVal CellIndex As Long) As Long

    If IsCell(Index, CellIndex) Then
        CellWidth = Lines(Index).Cells(CellIndex).Width
    End If
    
End Property

Public Property Let CellWidth(ByVal Index As Long, ByVal CellIndex As Long, ByVal vNewCellWidth As Long)

    If IsCell(Index, CellIndex) Then
        Lines(Index).Cells(CellIndex).Width = vNewCellWidth
        
        ' The cell has been changed
        RaiseEvent CellChange(Index, CellIndex, False)
    End If
    
End Property

Public Property Get CellBorder(ByVal Index As Long, ByVal CellIndex As Long) As Boolean

    If IsCell(Index, CellIndex) Then
        CellBorder = Lines(Index).Cells(CellIndex).Border
    End If
    
End Property

Public Property Let CellBorder(ByVal Index As Long, ByVal CellIndex As Long, ByVal vNewCellBorder As Boolean)

    If IsCell(Index, CellIndex) Then
        Lines(Index).Cells(CellIndex).Border = vNewCellBorder
        
        ' The cell has been changed
        RaiseEvent CellChange(Index, CellIndex, False)
    End If
    
End Property

Public Property Get CellCount(ByVal Index As Long) As Long

    If IsLine(Index) Then
        CellCount = Lines(Index).CellCount
    End If
    
End Property

Public Property Get LineType(ByVal Index As Long) As Long

    If IsLine(Index) Then
        LineType = Lines(Index).LineType
    End If
    
End Property

Public Property Let LineType(ByVal Index As Long, ByVal lLineType As Long)

    If IsLine(Index) Then
        Lines(Index).LineType = lLineType
        
        ' The line has been changed
        RaiseEvent LineChange(Index, False)
    End If
    
End Property

Public Property Get LineSpace(ByVal Index As Long) As Long

    If IsLine(Index) Then
        LineSpace = Lines(Index).LineSpace
    End If
    
End Property

Public Property Let LineSpace(ByVal Index As Long, ByVal lLineSpace As Long)

    If IsLine(Index) Then
        Lines(Index).LineSpace = lLineSpace
        
        ' The line has been changed
        RaiseEvent LineChange(Index, False)
    End If
    
End Property

Public Property Get LineWidth(ByVal Index As Long) As Long

    If IsLine(Index) Then
        LineWidth = Lines(Index).LineWidth
    End If
    
End Property

Public Property Let LineWidth(ByVal Index As Long, ByVal lLineWidth As Long)

    If IsLine(Index) Then
        Lines(Index).LineWidth = lLineWidth
        
        ' The line has been changed
        RaiseEvent LineChange(Index, False)
    End If
    
End Property

Public Property Get LineHeight(ByVal Index As Long) As Long

    If IsLine(Index) Then
        LineHeight = Lines(Index).LineHeight
    End If
    
End Property

Public Property Let LineHeight(ByVal Index As Long, ByVal lLineHeight As Long)

    If IsLine(Index) Then
        Lines(Index).LineHeight = lLineHeight
        
        ' The line has been changed
        RaiseEvent LineChange(Index, False)
    End If
    
End Property

Public Function AddLine(Optional ByVal Index As Long = -1, Optional ByVal LineType As Long) As Long

    Dim Tell As Long
    
    ' If the index is below zero, use the amout of lines insted
    If Index < 0 Then
        Index = LineCount
    End If
    
    ' Reallocate array to make place for the new line
    ReDim Preserve Lines(LineCount)
    
    ' Move elements upwards if necessary
    For Tell = LineCount To Index + 1 Step -1
        LSet Lines(Tell) = Lines(Tell - 1)
    Next
    
    ' The amout of lines has incresed
    LineCount = LineCount + 1
    
    ' Set the properties of this line
    Lines(Index).LineType = LineType

    ' Return the line index
    AddLine = Index
    
    ' The line has been changed
    RaiseEvent LineChange(Index, True)

End Function

Public Function AnalyzeDatabase() As Long

    Dim SumIndex As Long, Sum As Long

    ' Check and see if everything in region 1 is entered in
    If RegionFilled(1) Then
    
        ' Set the sum of the entire board
        CellText(LineCount - 1, Parent.Players.CurrentPlayer + 1) = _
         SumOfRegion(FindCellInRow(0, "sum:", 1, vbTextCompare), _
          LineCount - 1, Parent.Players.CurrentPlayer + 1)
        
        ' Check and see if we are finished
        If Parent.CheckForWinners Then
            AnalyzeDatabase = 1
            Exit Function
        End If
        
    Else
    
        ' If the other region insted is filled, calulcate it
        If RegionFilled(0) Then
        
            ' Get the index of the sum cell
            SumIndex = FindCellInRow(0, "Sum:", , vbTextCompare)
            
            ' Get the sum of the cells
            Sum = SumOfRegion(0, SumIndex - 1, Parent.Players.CurrentPlayer + 1)
            
            ' Set it to the sum of the cells
            CellText(SumIndex, Parent.Players.CurrentPlayer + 1) = Sum
        
            ' Check the bonus
            If Sum >= Parent.Script.Execute(Parent.Types.GameCode("BonusCondition"), "") Then
            
                ' Add the bonus
                CellText(SumIndex + 1, Parent.Players.CurrentPlayer + 1) = Parent.Script.Execute(Parent.Types.GameCode("BonusValue"), "")
            
            Else
            
                ' No bonus
                CellText(SumIndex + 1, Parent.Players.CurrentPlayer + 1) = 0
            
            End If
        
        End If
    
    End If

End Function

Public Function RegionFilled(Region As Long) As Boolean

    Dim Tell As Long, Seperator As Long
    
    Select Case Region
    Case 0 ' Check region 1
    
        ' See if all cells are filled
        For Tell = 1 To FindCellInRow(0, "Sum:", , vbTextCompare) - 1
            If CellText(Tell, Parent.Players.CurrentPlayer + 1) = "" Then
                Exit Function
            End If
        Next
    
        ' Yup, the region is filled
        RegionFilled = True
    
    Case 1 ' Check region 2
    
        ' Find seperator
        Seperator = FindLineByType(1)
    
        ' See if all cells are filled
        For Tell = Seperator + 1 To LineCount - 2
            If CellText(Tell, Parent.Players.CurrentPlayer + 1) = "" Then
                Exit Function
            End If
        Next
        
        ' The region is filled
        RegionFilled = True
    
    End Select

End Function

Public Sub RemoveLine(Optional ByVal Index As Long = -1)

    Dim Tell As Long

    ' If the index is below zero or outside of the total lines, use the amout of lines insted
    If Index < 0 Or Index > LineCount - 1 Then
        Index = LineCount - 2
    End If
    
    ' Move elements downwards if necessary
    For Tell = Index To LineCount - 2
        LSet Lines(Tell) = Lines(Tell + 1)
    Next
    
    ' The amout of lines has decresed
    LineCount = LineCount - 1
    
    ' Reallocate array to make place for the new line
    If LineCount <= 0 Then
        Erase Lines
    Else
        ReDim Preserve Lines(LineCount - 1)
    End If

    ' All lines are changed
    RaiseEvent LineChange(Index, True)
    
End Sub

Public Sub AddEmptyCells(lngLineIndex As Long, lngAmout As Long, lngCellWidth As Long, Optional lngCellSpace As Long)

    Dim Tell As Long

    ' Add the amout of cells
    For Tell = 1 To lngAmout
        AddCell lngLineIndex, , , "", lngCellSpace, lngCellWidth
    Next

End Sub

Public Function AddCell(ByVal Index As Long, Optional ByVal CellIndex As Long = -1, Optional CellType As Long, Optional CellText As String, Optional CellSpace As Long, Optional CellWidth As Long)

    Dim Tell As Long

    ' If the index is below zero, use the amout of lines insted
    If Index < 0 Then
        Index = LineCount
    End If
    
    ' If the index is below zero, use the amout of lines insted
    If CellIndex < 0 Then
        CellIndex = Lines(Index).CellCount
    End If
    
    ' Reallocate array to make place for the new cell
    ReDim Preserve Lines(Index).Cells(CellIndex)
    
    ' Move elements upwards if necessary
    For Tell = Lines(Index).CellCount To CellIndex + 1 Step -1
        LSet Lines(Index).Cells(Tell) = Lines(Index).Cells(Tell - 1)
    Next
    
    ' The amout of cells has incresed
    Lines(Index).CellCount = Lines(Index).CellCount + 1
    
    ' Set the properties of this cell
    Lines(Index).Cells(CellIndex).Type = CellType
    Lines(Index).Cells(CellIndex).Text = CellText
    Lines(Index).Cells(CellIndex).Space = CellSpace
    Lines(Index).Cells(CellIndex).Width = CellWidth
    Lines(Index).Cells(CellIndex).Background = vbButtonFace
    Lines(Index).Cells(CellIndex).BorderColorLight = vb3DLight
    Lines(Index).Cells(CellIndex).BorderColorDark = vb3DShadow
    Lines(Index).Cells(CellIndex).Border = True

    ' Return the cell index
    CellIndex = Index
    
    ' All cells are changed
    RaiseEvent CellChange(Index, CellIndex, True)

End Function

Public Sub RemoveCell(ByVal Index As Long, Optional ByVal CellIndex As Long = -1)

    Dim Tell As Long

    ' If the index is below zero or outside of the total lines, use the amout of lines insted
    If Index < 0 Or Index > LineCount - 1 Then
        Index = LineCount - 2
    End If
    
    ' Move elements downwards if necessary
    For Tell = CellIndex To Lines(Index).CellCount - 2
        LSet Lines(Index).Cells(Tell) = Lines(Index).Cells(Tell + 1)
    Next
    
    ' Reallocate array to make place for the new line
    If Index < 0 Then
        Erase Lines(Index).Cells
    Else
        ReDim Preserve Lines(Index).Cells(Index)
    End If
    
    ' The amout of lines has decresed
    Lines(Index).CellCount = Lines(Index).CellCount - 1
    
    ' All cells are changed
    RaiseEvent CellChange(Index, CellIndex, True)
    
End Sub

Public Sub ClearLines()

    ' Clear all lines
    Erase Lines

    ' Clear line count
    LineCount = 0
    
    ' All cells are changed
    RaiseEvent LineChange(-1, True)

End Sub

Public Sub ClearCells(ByVal Index As Long)

    ' Clear all cell
    Erase Lines(Index).Cells
    
    ' Clear the count
    Lines(Index).CellCount = 0
    
    ' All cells are changed
    RaiseEvent CellChange(Index, -1, True)
    
End Sub

Public Function SumOfRegion(StartIndex As Long, EndIndex As Long, Row As Long) As Long

    Dim Tell As Long
    
    ' Search through the specified lines
    For Tell = StartIndex To EndIndex
    
        ' Add the value of that line
        SumOfRegion = SumOfRegion + Val(CellText(Tell, Row))

    Next

End Function

Public Function FindCellInRow(Row As Long, sText As String, Optional StartIndex As Long, Optional Compare As VbCompareMethod) As Long

    Dim Tell As Long
    
    ' Search through all lines
    For Tell = StartIndex To LineCount - 1
    
        ' Compare cell
        If StrComp(CellText(Tell, Row), sText, Compare) = 0 Then
        
            ' Return line index if match
            FindCellInRow = Tell
        
            ' Nothing else to do
            Exit Function
        
        End If
    Next

    ' No results
    FindCellInRow = -1

End Function

Public Function FindLineByType(lngType As Long, Optional StartIndex As Long) As Long

    Dim Tell As Long
    
    ' Search through all lines
    For Tell = StartIndex To LineCount - 1
    
        ' Compare type
        If LineType(Tell) = lngType Then
        
            ' Return line index if match
            FindLineByType = Tell
        
            ' Nothing else to do
            Exit Function
        
        End If
    Next

    ' No results
    FindLineByType = -1

End Function

Public Function IsLine(ByVal Index As Long) As Boolean

    ' Return wether or not this line is valid
    IsLine = CBool(Index >= 0 And Index < LineCount)

End Function

Public Function IsCell(ByVal Index As Long, ByVal CellIndex As Long) As Boolean

    ' Return wether or not this cell is valid
    If IsLine(Index) Then
        IsCell = CBool(CellIndex >= 0 And CellIndex < Lines(Index).CellCount)
    End If

End Function


