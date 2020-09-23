VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPlayers"
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

Private Type Player
    Name As String
    Throws As Long
    Source As String
    Local As Boolean
    Disabled As Boolean
End Type

' Public events
Public Event PlayerChanged()

' Array holding all the player names
Private Player() As Player
Private pCurrentPlayer As Long

' The amout of players
Public PlayerCount As Long
Public JustChanged As Boolean
Public Parent As Object

Public Property Get CurrentPlayer() As Long

    CurrentPlayer = pCurrentPlayer

End Property

Public Property Let CurrentPlayer(ByVal vNewValue As Long)

    ' Set the value
    pCurrentPlayer = vNewValue
    
    ' If this value is over the amout of players, reset it
    If pCurrentPlayer > PlayerCount - 1 Then
        pCurrentPlayer = 0
    End If
    
    ' Tell about these changes
    RaiseEvent PlayerChanged

End Property

Public Property Get PlayerName(Optional ByVal Index As Variant) As String

    If IsMissing(Index) Then
        Index = CurrentPlayer
    End If

    ' Get the index
    Index = GetIndex(Index)

    If IsPlayer(Index) Then
        PlayerName = Player(Index).Name
    End If

End Property

Public Property Let PlayerName(Optional ByVal Index As Variant, ByVal vNewValue As String)

    If IsMissing(Index) Then
        Index = CurrentPlayer
    End If

    ' Get the index
    Index = GetIndex(Index)

    If IsPlayer(Index) Then
        
        Player(Index).Name = vNewValue
    
        ' Tell about these changes
        RaiseEvent PlayerChanged
    
    End If

End Property

Public Property Get PlayerThrows(Optional ByVal Index As Variant) As Long

    If IsMissing(Index) Then
        Index = CurrentPlayer
    End If
    
    ' Get the index
    Index = GetIndex(Index)

    If IsPlayer(Index) Then
        PlayerThrows = Player(Index).Throws
    End If

End Property

Public Property Let PlayerThrows(Optional ByVal Index As Variant, ByVal vNewValue As Long)

    If IsMissing(Index) Then
        Index = CurrentPlayer
    End If
    
    ' Get the index
    Index = GetIndex(Index)

    If IsPlayer(Index) Then
    
        Player(Index).Throws = vNewValue
        
        ' Tell about these changes
        RaiseEvent PlayerChanged
        
    End If

End Property

Public Property Get PlayerSource(Optional ByVal Index As Variant) As String

    If IsMissing(Index) Then
        Index = CurrentPlayer
    End If
    
    ' Get the index
    Index = GetIndex(Index)

    If IsPlayer(Index) Then
        PlayerSource = Player(Index).Source
    End If

End Property

Public Property Let PlayerSource(Optional ByVal Index As Variant, ByVal vNewValue As String)

    If IsMissing(Index) Then
        Index = CurrentPlayer
    End If
    
    ' Get the index
    Index = GetIndex(Index)

    If IsPlayer(Index) Then
        Player(Index).Source = vNewValue
    End If

End Property

Public Property Get PlayerDisabled(Optional ByVal Index As Variant) As Boolean

    If IsMissing(Index) Then
        Index = CurrentPlayer
    End If
    
    ' Get the index
    Index = GetIndex(Index)

    If IsPlayer(Index) Then
        PlayerLocal = Player(Index).Disabled
    End If

End Property

Public Property Let PlayerDisabled(Optional ByVal Index As Variant, ByVal vNewValue As Boolean)

    If IsMissing(Index) Then
        Index = CurrentPlayer
    End If
    
    ' Get the index
    Index = GetIndex(Index)

    If IsPlayer(Index) Then
        Player(Index).Disabled = vNewValue
    End If

End Property

Public Property Get PlayerLocal(Optional ByVal Index As Variant) As Boolean

    If IsMissing(Index) Then
        Index = CurrentPlayer
    End If
    
    ' Get the index
    Index = GetIndex(Index)

    If IsPlayer(Index) Then
        PlayerLocal = Player(Index).Local
    End If

End Property

Public Property Let PlayerLocal(Optional ByVal Index As Variant, ByVal vNewValue As Boolean)

    If IsMissing(Index) Then
        Index = CurrentPlayer
    End If
    
    ' Get the index
    Index = GetIndex(Index)

    If IsPlayer(Index) Then
        Player(Index).Local = vNewValue
    End If

End Property

Public Function ToString(Optional sName As String) As String

    Dim Tell As Long
    
    ' Add the required information to the string
    For Tell = 0 To PlayerCount - 1
        ToString = ToString & PlayerName(Tell) & "," & CLng(PlayerSource(Tell) = sName) & IIf(Tell < PlayerCount - 1, ",", "")
    Next

End Function

Public Function Clone() As clsPlayers

    Dim Tell As Long

    ' Create the new class
    Set Clone = New clsPlayers
    
    ' Recreate the content
    For Tell = 0 To PlayerCount - 1
        
        ' Add the player
        Clone.AddPlayer Player(Tell).Name
    
        ' Set all the settings
        Clone.PlayerSource(Tell) = Player(Tell).Source
        Clone.PlayerThrows(Tell) = Player(Tell).Throws
        Clone.PlayerLocal(Tell) = Player(Tell).Local
    
    Next

End Function

Public Function FindEnabled(ByVal StartIndex As Long, Optional EndIndex As Long = -1, Optional bMoveToStart As Boolean) As Long

    Dim Tell As Long
    
    ' Default value
    If EndIndex < 0 Then
        EndIndex = PlayerCount - 1
    End If
    
    ' Firstly, try to find an enabled player upwards
    For Tell = StartIndex To PlayerCount - 1
    
        ' See if it's enabled
        If Not Player(Tell).Disabled Then
    
            ' Return index
            FindEnabled = Tell
            
            ' Nothing more to do
            Exit Function
    
        End If
    
    Next
    
    ' Nothing was found
    FindEnabled = -1
    
    ' If we are instructed to try to find a enabled player in the rest, do so
    If bMoveToStart Then
        
        ' Find the rest
        FindEnabled = Me.FindEnabled(0, StartIndex - 1, False)
    
    End If

End Function

Public Function IsPlayer(ByVal Index As Long) As Boolean

    ' Return wether or not the passed index is valid
    IsPlayer = CBool(Index >= 0 And Index < PlayerCount)

End Function

Public Sub RemovePlayer(Optional ByVal Index As Long = -1)

    Dim Tell As Long

    ' If the index is below zero or outside of the total lines, use the amout of lines insted
    If Index < 0 Or Index > PlayerCount - 1 Then
        Index = PlayerCount - 2
    End If
    
    ' Move elements downwards if necessary
    For Tell = Index To PlayerCount - 2
        LSet Player(Tell) = Player(Tell + 1)
    Next
    
    ' The amout of lines has decresed
    PlayerCount = PlayerCount - 1
    
    ' Reallocate array to make place for the new line
    If PlayerCount <= 0 Then
        Erase Player
    Else
        ReDim Preserve Player(PlayerCount - 1)
    End If

End Sub

Public Function AddPlayer(sName As String) As Long

    ' Reallocate array
    ReDim Preserve Player(0 To PlayerCount)
    
    ' Add name
    Player(PlayerCount).Name = sName
    
    ' Return the index
    AddPlayer = PlayerCount
    
    ' New player has been added
    PlayerCount = PlayerCount + 1

End Function

Public Function FindPlayer(sName As String) As Long

    Dim Tell As Long
    
    ' Loop through all elements and look for the name
    For Tell = 0 To PlayerCount - 1
    
        If LCase(sName) = LCase(Player(Tell).Name) Then
        
            ' Return the result
            FindPlayer = Tell
            
            ' Nothing more to do
            Exit Function
        
        End If
    
    Next

    ' No results
    FindPlayer = -1

End Function

Private Function GetIndex(Index As Variant) As Long

    If IsNumeric(Index) Then
        GetIndex = Index
    Else
        GetIndex = FindPlayer(CStr(Index))
    End If

End Function

Public Sub ClearPlayers()

    ' Clear array
    Erase Player
    
    ' Clear count
    PlayerCount = 0

End Sub
