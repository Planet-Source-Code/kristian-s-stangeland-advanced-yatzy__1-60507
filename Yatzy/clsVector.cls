VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsVector"
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

' Private class holding the pScript
Private pScript As New clsScript
Private pSource As Variant

Public Property Get Source(Optional Index As Variant) As Variant

    ' If the index is missing, then we're talking about the variable pSource
    If IsMissing(Index) Then
        Source = pSource
    Else
        ' If not, then it is the element IN the array pSource it is refering to
        Source = pSource(Val(Index))
    End If

End Property

Public Property Let Source(Optional Index As Variant, ByVal vNewValue As Variant)

    ' If the index is missing, then we're talking about the variable pSource
    If IsMissing(Index) Then
        pSource = vNewValue
    Else
        ' If not, then it is the element IN the array pSource it is refering to
        pSource(Val(Index)) = vNewValue
    End If

End Property

Public Function Correspond(sCondition As String, Optional Operator As Long) As Variant

    Dim Tell As Long, Temp As Variant
    
    ' Check and see if all the elements correspond to a certain condition
    For Tell = LBound(pSource) To UBound(pSource)

        Temp = pScript.Execute(sCondition, pSource(Tell))

        If Tell = LBound(pSource) Then
            ' Don't combine with operator at the first element
            Correspond = Temp
        Else
        
            ' Combine with the specified operator
            Select Case Operator
            Case 0: Correspond = Correspond And Temp
            Case 1: Correspond = Correspond Or Temp
            Case 2: Correspond = Correspond Xor Temp
            Case 3: Correspond = Correspond Eqv Temp
            Case 4: Correspond = Correspond Imp Temp
            End Select
        
        End If
        
    Next

End Function

Public Sub Keep(StartIndex As Long, EndIndex As Long)

    Dim Tell As Long, Temp() As Variant
    
    ' We cannot use redim to erase an array
    If StartIndex > EndIndex Then
    
        ' This instruct us to simply erase the array
        Erase pSource
        
        ' Nothing else to do
        Exit Sub
    
    End If
    
    ' Reallocate array
    ReDim Preserve Temp(StartIndex To EndIndex)
    
    ' Set the array
    For Tell = StartIndex To EndIndex
        Temp(Tell) = pSource(Tell)
    Next
    
    ' Set the pSource to this new array
    pSource = Temp
    
End Sub

' AllowCount is the amout of elements that are allowed to not be excluded. This variable is ignored if AllowCount is below 0
Public Sub Exclude(sCondition As String, Optional ByVal AllowCount As Long = -1, Optional TruePart, Optional FalsePart = 0)

    Dim Tell As Long

    ' Loop through all elements including those that correspond to the condition, thereby removing everything else
    For Tell = LBound(pSource) To UBound(pSource)
    
        If AllowCount <> 0 Then
            If pScript.Execute(sCondition, pSource(Tell)) Then
            
                ' Remove this element
                pSource(Tell) = FalsePart

            Else
            
                ' This element has been allowed, decrese allow count if not below zero
                If AllowCount > 0 Then
                    AllowCount = AllowCount - 1
                End If
                
                ' Set the element to something, if requested
                If Not IsMissing(TruePart) Then
                
                    ' Set the element
                    pSource(Tell) = TruePart
                
                End If
            
            End If
        Else
        
            ' Remove this element
            pSource(Tell) = 0
        
        End If
    
    Next

End Sub

Public Function Condition(sCondition As String) As Variant

    Dim Tell As Long

    ' Count the amout of elements that correspond to this condition
    For Tell = LBound(pSource) To UBound(pSource)
    
        If pScript.Execute(sCondition, pSource(Tell)) Then

            Condition = Condition + 1
            
        End If
    
    Next

End Function

Public Function Clone() As clsVector

    ' Create a new class of the same type
    Set Clone = New clsVector
    
    ' Use the source of this class
    Clone.Source = pSource

End Function

Public Property Get Count() As Long

    On Error Resume Next
    
    ' When the array is uninitialized, return -1
    Count = -1
    
    ' Return the count
    Count = UBound(pSource) - LBound(pSource) + 1

End Property

Public Function Summarize() As Variant

    Dim Tell As Long
    
    ' Loop through all adding them together
    For Tell = LBound(pSource) To UBound(pSource)
        ' Add the element
        Summarize = Summarize + (pSource(Tell) * Tell)
    Next

End Function

Public Function Multiply() As Variant

    Dim Tell As Long
    
    ' Always start with one
    Multiply = 1
    
    ' Loop through all adding them together
    For Tell = LBound(pSource) To UBound(pSource)
        ' Don't go further if the number is zero
        If pSource(Tell) = 0 Then
            ' Add the element
            Multiply = Multiply * (pSource(Tell) * Tell)
        End If
    Next

End Function

Public Function Divide() As Variant

    On Error Resume Next
    Dim Tell As Long
    
    ' Always start with the first element
    Multiply = pSource(LBound(pSource))
    
    ' Loop through all adding them together
    For Tell = LBound(pSource) + 1 To UBound(pSource)
        If pSource(Tell) <> 0 Then 'We can't divide by zero
            ' Add the element
            Divide = Divide / (pSource(Tell) * Tell)
        End If
    Next

End Function
