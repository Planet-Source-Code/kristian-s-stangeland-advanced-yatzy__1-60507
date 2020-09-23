VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsType"
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

Private Type GameType
    Region As Long
    Name As String
    Code As String
End Type

' Array of game types
Private Games() As GameType

' The amout of game types
Public GameCount As Long

' Different public objects
Public Parent As Object
Public FileName As String

Public Property Get GameRegion(ByVal Index As Variant) As Long

    Index = GetIndex(Index)

    If IsGame(Index) Then
        GameRegion = Games(Index).Region
    End If

End Property

Public Property Let GameRegion(ByVal Index As Variant, ByVal vNewValue As Long)
    
    Index = GetIndex(Index)

    If IsGame(Index) Then
        Games(Index).Region = vNewValue
    End If

End Property

Public Property Get GameName(ByVal Index As Variant) As String

    Index = GetIndex(Index)

    If IsGame(Index) Then
        GameName = Games(Index).Name
    End If

End Property

Public Property Let GameName(ByVal Index As Variant, ByVal vNewValue As String)

    Index = GetIndex(Index)

    If IsGame(Index) Then
        Games(Index).Name = vNewValue
    End If

End Property

Public Property Get GameCode(ByVal Index As Variant) As String

    Index = GetIndex(Index)

    If IsGame(Index) Then
        GameCode = Games(Index).Code
    End If

End Property

Public Property Let GameCode(ByVal Index As Variant, ByVal vNewValue As String)

    Index = GetIndex(Index)

    If IsGame(Index) Then
        Games(Index).Code = vNewValue
    End If

End Property

Public Function CalulateValue(ByVal Index As Variant, DiceArray As Variant) As Long

    ' Use vector class
    Dim Vector As New clsVector

    ' Initialize class
    Vector.Source = DiceArray
    
    ' Get the index to calculate from
    Index = GetIndex(Index)

    If IsGame(Index) Then
        CalulateValue = Parent.Script.Execute(GameCode(Index), Vector)
    End If

End Function

Public Function AddType(lngRegion As Long, Name As String, Code As String, Optional Index As Variant) As Long
    
    Dim Tell As Long
    
    ' Default index
    If IsMissing(Index) Then
        Index = GameCount
    End If
    
    ' When the index is numeric and under zero, use default index
    If Val(Index) < 0 Then
        Index = GameCount
    End If
    
    ' Get the index
    Index = GetIndex(Index)

    ' Only proceed if this is a valid index
    If IsGame(Index) Then
    
        ' Reallocate array
        ReDim Preserve Games(GameCount)
        
        ' Move elements upwards if necessary
        For Tell = GameCount To Index + 1 Step -1
            LSet Games(Tell) = Games(Tell - 1)
        Next
        
        ' Set properties
        Games(Index).Code = Code
        Games(Index).Name = Name
        Games(Index).Region = lngRegion
        
        ' Increse the game count
        GameCount = GameCount + 1
    
        ' Return the line index
        AddType = Index
    
    End If

End Function

Public Sub RemoveType(ByVal Index As Variant)
    
    Dim Tell As Long
    
    ' Reallocate array
    Index = GetIndex(Index)
    
    If IsGame(Index) Then
    
        ' Move elements downwards if necessary
        For Tell = Index To GameCount - 2
            LSet Games(Tell) = Games(Tell + 1)
        Next
        
        ' Decresed the game count
        GameCount = GameCount - 1
        
        ' Reallocate array to make place for the new game type
        If GameCount <= 0 Then
            Erase Games
        Else
            ReDim Preserve Games(GameCount - 1)
        End If

    End If

End Sub

Public Sub ClearGames()

    ' Clear array
    Erase Games
    
    ' Clear count
    GameCount = 0

End Sub

Public Function IsGame(ByVal Index As Long) As Boolean

    ' Return wether or not this index is valid
    IsGame = CBool(Index >= 0 And Index < GameCount)

End Function

Public Sub LoadData(sFile As String)

    Dim Free As Long

    ' Only go further if this file exists
    If Dir(sFile) <> "" Then
    
        ' Get a free handle
        Free = FreeFile
    
        Open sFile For Binary As #Free
        
            Get #Free, , GameCount
            
            ' Resize array and get data
            If GameCount > 0 Then
                ReDim Games(GameCount - 1)
            
                ' Retrive data
                Get #Free, , Games
            End If
            
        Close #Free
    
    End If

    ' Save the file name
    FileName = sFile

End Sub

Public Sub SaveData(sFile As String)

    Dim Free As Long

    ' Get a free handle
    Free = FreeFile
    
    ' If this file exists, delete it
    If Dir(sFile) <> "" Then
        Kill sFile
    End If

    Open sFile For Binary As #Free
        
        ' Write data
        Put #Free, , GameCount
        Put #Free, , Games
        
    Close #Free

End Sub

Public Function FindType(sName As String) As Long

    Dim Tell As Long
    
    ' Loop through all elements and look for the name
    For Tell = 0 To GameCount - 1
    
        If LCase(sName) = LCase(Games(Tell).Name) Then
        
            ' Return the result
            FindType = Tell
            
            ' Nothing more to do
            Exit Function
        
        End If
    
    Next

    ' No results
    FindType = -1

End Function

Private Function GetIndex(Index As Variant) As Long

    If IsNumeric(Index) Then
        GetIndex = Index
    Else
        GetIndex = FindType(CStr(Index))
    End If

End Function



