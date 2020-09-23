VERSION 5.00
Begin VB.UserControl usrDices 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   Begin VB.PictureBox picBase 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   2895
      Left            =   0
      ScaleHeight     =   193
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   257
      TabIndex        =   0
      Top             =   0
      Width           =   3855
   End
End
Attribute VB_Name = "usrDices"
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

Private Type Dice
    Value As Long
    BackColor As Long
    BorderColor As Long
    CircleColor As Long
    CircleBorder As Long
End Type

' Events
Public Event DiceClick(DiceIndex As Long)

' The private array holding the number of each dice
Private pDices() As Dice

' Different settings
Private pDiceCount As Long
Private pDiceWidth As Long
Private pDiceHeight As Long
Private pDiceSpace As Long
Private pDiceStartValue As Long
Private pDiceResetValue As Boolean

Public Property Get DiceStartValue() As Long

    DiceStartValue = pDiceStartValue

End Property

Public Property Let DiceStartValue(ByVal vNewValue As Long)

    pDiceStartValue = vNewValue

End Property

Public Property Get DiceResetValue() As Boolean

    DiceResetValue = pDiceResetValue

End Property

Public Property Let DiceResetValue(ByVal vNewValue As Boolean)

    pDiceResetValue = vNewValue

End Property

Public Property Get DiceSpace() As Long

    DiceSpace = pDiceSpace

End Property

Public Property Let DiceSpace(ByVal vNewValue As Long)

    pDiceSpace = vNewValue

End Property

Public Property Get DiceHeight() As Long

    DiceHeight = pDiceHeight

End Property

Public Property Let DiceHeight(ByVal vNewValue As Long)

    pDiceHeight = vNewValue

End Property

Public Property Get DiceWidth() As Long

    DiceWidth = pDiceWidth

End Property

Public Property Let DiceWidth(ByVal vNewValue As Long)

    pDiceWidth = vNewValue

End Property

Public Property Get DiceCount() As Long

    DiceCount = pDiceCount

End Property

Public Property Let DiceCount(ByVal vNewValue As Long)
    
    ' Reallocate dices
    AllocateDices vNewValue

End Property

Public Property Get DiceValue(ByVal Index As Long) As Long

    DiceValue = pDices(Index).Value

End Property

Public Property Let DiceValue(ByVal Index As Long, ByVal vNewValue As Long)

    If IsDice(Index) Then
        
        ' Set value
        pDices(Index).Value = vNewValue
    
        ' Redraw all dices
        DrawAll
    
    End If

End Property

Public Property Get DiceBackColor(ByVal Index As Long) As Long

    DiceBackColor = pDices(Index).BackColor

End Property

Public Property Let DiceBackColor(ByVal Index As Long, ByVal vNewValue As Long)

    If IsDice(Index) Then
        
        ' Set value
        pDices(Index).BackColor = vNewValue
    
        ' Redraw all dices
        DrawAll
    
    End If

End Property

Public Property Get DiceBorderColor(ByVal Index As Long) As Long

    DiceBorderColor = pDices(Index).BorderColor

End Property

Public Property Let DiceBorderColor(ByVal Index As Long, ByVal vNewValue As Long)

    If IsDice(Index) Then
        
        ' Set value
        pDices(Index).BorderColor = vNewValue
    
        ' Redraw all dices
        DrawAll
    
    End If

End Property

Public Property Get DiceCircleBorder(ByVal Index As Long) As Long

    DiceCircleBorder = pDices(Index).CircleBorder

End Property

Public Property Let DiceCircleBorder(ByVal Index As Long, ByVal vNewValue As Long)

    If IsDice(Index) Then
        
        ' Set value
        pDices(Index).CircleBorder = vNewValue
    
        ' Redraw all dices
        DrawAll
    
    End If

End Property

Public Property Get DiceCircleColor(ByVal Index As Long) As Long

    DiceCircleColor = pDices(Index).CircleColor

End Property

Public Property Let DiceCircleColor(ByVal Index As Long, ByVal vNewValue As Long)

    If IsDice(Index) Then
        
        ' Set value
        pDices(Index).CircleColor = vNewValue
    
        ' Redraw all dices
        DrawAll
    
    End If

End Property

Public Sub AllocateDices(Amout As Long)

    Dim Tell As Long

    ' Reallocate array
    If Amout < 0 Then
        Erase pDices
    Else
        
        ReDim pDices(Amout)
    
        ' Set default values
        For Tell = 0 To Amout
            pDices(Tell).BackColor = vbWhite
        Next
    
    End If

    ' Set the value
    pDiceCount = Amout

    ' Redraw dices
    DrawAll

End Sub

Public Function IsDice(ByVal Index As Long) As Boolean

    ' Return wether or not it is a valid index
    IsDice = CBool(Index >= 0 And Index <= DiceCount)

End Function

Public Function DiceArray() As Variant

    Dim Tell As Long, Number As Long, tempArray(1 To 6) As Variant
    
    For Number = 1 To 6
    
        For Tell = 0 To pDiceCount
            If pDices(Tell).Value = Number Then
            
                ' Increse the amout
                tempArray(Number) = Val(tempArray(Number)) + 1
            
            End If
        Next
    
    Next
    
    ' Return the array
    DiceArray = tempArray

End Function

Public Sub ResetDices(Optional SelectColor As Long)

    Dim Tell As Long
    
    ' Reset all unselected dices
    For Tell = 0 To DiceCount
        If DiceBackColor(Tell) <> SelectColor Then
            DiceValue(Tell) = 0
        End If
    Next

End Sub

Public Sub DrawAll()

    Dim Tell As Long

    ' Clear picture box
    picBase.Cls

    ' Draw all dices
    For Tell = 0 To pDiceCount
        DrawDice picBase, (Tell * (pDiceWidth + pDiceSpace)) + DiceSpace, DiceSpace, pDiceWidth, pDiceHeight, pDices(Tell).Value, 0, pDices(Tell).BackColor, pDices(Tell).CircleColor, pDices(Tell).CircleBorder, pDices(Tell).BorderColor
    Next

End Sub

Public Sub DrawDice(Control As Object, x As Long, y As Long, Width As Long, Height As Long, Number As Long, Optional PointSpace As Double, Optional Background As Long = vbWhite, Optional CircleColor As Long = vbBlack, Optional CircleBorder As Long = vbBlack, Optional BorderColor As Long = vbBlack)

    Dim Radius As Double, cX As Long, cY As Long

    ' Firstly draw the dice itself
    Control.Line (x, y)-(x + Width, y + Height), Background, BF

    ' Then draw the border
    Control.FillStyle = 1
    Control.Line (x, y)-(x + Width, y + Height), BorderColor, B

    ' Calculate the radius of the circles
    Radius = Width / 10

    ' Set the fillstyle and fillcolor
    Control.FillColor = CircleColor
    Control.FillStyle = 0
    
    ' The position of the center
    cX = x + (Width / 2)
    cY = y + (Height / 2)
    
    ' Default space
    If PointSpace <= 0 Then
        PointSpace = (Radius * 2.5)
    End If
    
    ' All these dices has a point in the middle
    If Number = 1 Or Number = 3 Or Number = 5 Then
        Control.Circle (cX, cY), Radius, CircleBorder
    End If

    ' And finaly draw the circles
    Select Case Number
    Case 2, 3
        Control.Circle (cX + PointSpace, cY - PointSpace), Radius, CircleBorder
        Control.Circle (cX - PointSpace, cY + PointSpace), Radius, CircleBorder
    
    Case 4, 5
        Control.Circle (cX - PointSpace, cY - PointSpace), Radius, CircleBorder
        Control.Circle (cX + PointSpace, cY - PointSpace), Radius, CircleBorder
        Control.Circle (cX - PointSpace, cY + PointSpace), Radius, CircleBorder
        Control.Circle (cX + PointSpace, cY + PointSpace), Radius, CircleBorder
    
    Case 6
        Control.Circle (cX - PointSpace, cY - PointSpace), Radius, CircleBorder
        Control.Circle (cX - PointSpace, cY), Radius, CircleBorder
        Control.Circle (cX - PointSpace, cY + PointSpace), Radius, CircleBorder
        Control.Circle (cX + PointSpace, cY - PointSpace), Radius, CircleBorder
        Control.Circle (cX + PointSpace, cY), Radius, CircleBorder
        Control.Circle (cX + PointSpace, cY + PointSpace), Radius, CircleBorder
    
    End Select

End Sub

Public Function ToString() As String

    Dim Tell As Long

    ' Create a string of the dice array
    For Tell = 0 To pDiceCount
        ' Add value and backcolor
        ToString = ToString & pDices(Tell).Value & "," & pDices(Tell).BackColor & IIf(Tell < pDiceCount, ",", "")
    Next

End Function

Private Sub picBase_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    Dim DiceIndex As Double
    
    DiceIndex = x / (DiceWidth + DiceSpace)
 
    ' Do not overload array
    If Fix(DiceIndex) <= DiceCount Then
 
        ' Check and see if we're in the right spot
        If (DiceIndex - Fix(DiceIndex)) * (DiceWidth + DiceSpace) <= DiceWidth Then
    
            ' Don't proceed if the user has click to high or too low
            If y >= DiceSpace And y <= picBase.Height - DiceSpace Then
                ' Raise the event
                RaiseEvent DiceClick(CLng(Fix(DiceIndex)))
            End If
        
        End If
    
    End If

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    ' Load saved properties
    DiceCount = PropBag.ReadProperty("DiceCount", 4)
    DiceWidth = PropBag.ReadProperty("DiceWidth", 40)
    DiceHeight = PropBag.ReadProperty("DiceHeight", 40)
    DiceSpace = PropBag.ReadProperty("DiceSpace", 8)
    DiceStartValue = PropBag.ReadProperty("DiceStartValue", 3)
    DiceResetValue = PropBag.ReadProperty("DiceResetValue", True)

End Sub

Private Sub UserControl_Resize()

    ' Resize the picture box
    picBase.Width = UserControl.ScaleWidth
    picBase.Height = UserControl.ScaleHeight

End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    ' Save properties
    PropBag.WriteProperty "DiceCount", DiceCount
    PropBag.WriteProperty "DiceWidth", DiceWidth
    PropBag.WriteProperty "DiceHeight", DiceHeight
    PropBag.WriteProperty "DiceSpace", DiceSpace
    PropBag.WriteProperty "DiceStartValue", DiceStartValue
    PropBag.WriteProperty "DiceResetValue", DiceResetValue

End Sub
