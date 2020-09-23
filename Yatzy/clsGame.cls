VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGame"
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

' Public event
Public Event PlayerListRecived()
Public Event ClientListRecived()
Public Event PlayerChanged()

Public Types As New clsType
Public FileSystem As New clsFileSystem
Public Script As New clsScript
Public Language As New clsLanguage
Public LocalInfo As New clsLocalInfo

' Classes with events
Public WithEvents Players As clsPlayers
Attribute Players.VB_VarHelpID = -1
Public WithEvents NetGame As clsNetGame
Attribute NetGame.VB_VarHelpID = -1

' Classes refering to controls or other classes
Public DataBase As clsDatabase
Public Board As usrBoard
Public Dices As usrDices

Public Sub NewGame()

    Dim Tell As Long, lngLine As Long, ret As Long, Region As Long

    ' Clear the database
    DataBase.ClearLines
    
    ' Don't draw the changes (the result would be flickering)
    Board.IgnoreDrawing = True
    
    ' Add the first line
    ret = DataBase.AddLine()
    
    ' Add cells
    DataBase.AddCell ret, , , "Type:", 8, 73
    
    ' Add all the different player names
    For Tell = 0 To Players.PlayerCount - 1
        DataBase.AddCell ret, , , Players.PlayerName(Tell), 0, 40
    Next
    
    ' Add region 0 and 1
    For Region = 0 To 1
    
        For Tell = 0 To Types.GameCount - 1
            If Types.GameRegion(Tell) = Region Then
        
                ' Add line
                AddLine Types.GameName(Tell) & ":"
        
            End If
        Next
    
        If Region = 0 Then
            ' Add lines
            AddLine "Sum:"
            AddLine "Bonus:"
            
            ' Add seperator
            DataBase.LineSpace(DataBase.AddLine(, 1)) = 8
        Else
        
            ' Add the "sum"-line
            AddLine "Sum:"
        
        End If
    
    Next
    
    ' Only the server or a game hosted localy should set the current player
    If NetGame.NetType = Net_Server Or NetGame.NetStatus = Net_LocalGame Then
        ' Go to first player
        Players.CurrentPlayer = -1
        NextPlayer
    Else
        ' It will be the first player
        Players.CurrentPlayer = 0
    End If
    
    ' Call the initialize code
    Script.Execute Types.GameCode("Initialize"), ""
    
    ' Draw the changes
    Board.IgnoreDrawing = False
    Board.AutoResize = True
    Board.DrawAll
    Board.AutoResize = False

End Sub

Public Function CheckForWinners() As Boolean

    Dim Tell As Long, CellText As String, Winner As Long, WinnerIndex As Long
    
    ' Automatically assume that the first player is the winner (the code below will prove if this is correct)
    Winner = Val(DataBase.CellText(DataBase.LineCount - 1, 1))
    WinnerIndex = 0
    
    ' Go through all cells at the last line
    For Tell = 1 To DataBase.CellCount(DataBase.LineCount - 1) - 1
        
        ' Get the cell text
        CellText = DataBase.CellText(DataBase.LineCount - 1, Tell)
        
        ' If one of those cells are empty, then we're not finished with the game
        If CellText = "" Then
            Exit Function
        Else
        
            ' If not, then check and see if this is the winner
            If Winner < Val(CellText) Then
                ' This is the winner
                Winner = Val(CellText)
                WinnerIndex = Tell - 1
            End If
        
        End If
        
    Next
    
    ' Tell the user who won
    MsgBox Language.Constant("Congratulations") & ", " & Players.PlayerName(WinnerIndex) & ", " & Language.Constant("PlayerWon"), vbInformation

    ' Return that someone has won
    CheckForWinners = True

End Function

Public Sub NextPlayer()
    
    If NetGame.NetStatus = Net_Playing And NetGame.NetType = Net_Client Then
        ' Clients must send a next player request to the server
        NetGame.Broadcast "PLAYERFINISHED"
        Exit Sub
    End If
    
    ' Analyze database to find wether or not someone has won
    If DataBase.AnalyzeDatabase = 1 Then
        ' Don't enable anything or go to next player
        Exit Sub
    End If
    
    ' Go to next player and give three throws
    Players.CurrentPlayer = Players.FindEnabled(Players.CurrentPlayer + 1, , True)
    
    ' Tell about this
    Players.JustChanged = True
    
    ' If we're supposed to not take account for earlier throws
    If Dices.DiceResetValue Then
        Players.PlayerThrows = Dices.DiceStartValue
    Else
        Players.PlayerThrows = Players.PlayerThrows + Dices.DiceStartValue
    End If

    ' If this is a server and a net game, inform clients
    If NetGame.NetStatus = Net_Playing And NetGame.NetType = Net_Server Then
        NetGame.Broadcast "NEXTPLAYER " & Players.CurrentPlayer & ", " & Players.PlayerThrows & ", True"
    End If

End Sub

Private Sub AddLine(sName As String)

    Dim ret As Long

    ' Add the line
    ret = DataBase.AddLine()
    DataBase.AddCell ret, , 0, sName, 8, 73
    DataBase.AddEmptyCells ret, Players.PlayerCount, 40

End Sub

Private Sub Class_Initialize()

    ' Create players class
    Set Players = New clsPlayers
    
    ' Create net class
    Set NetGame = New clsNetGame
    
    ' Set parents
    Set Language.Parent = Me
    Set Players.Parent = Me
    Set Types.Parent = Me
    Set NetGame.Parent = Me

End Sub

Private Sub NetGame_ClientClosed(Index As Long)

    Dim Tell As Long
    
    ' Remove all players belonging to this client
    For Tell = 0 To Players.PlayerCount - 1
        If Players.PlayerSource(Tell) = NetGame.Clients(Index).Name Then
    
           ' Disable this player
           Players.PlayerDisabled(Tell) = True
    
        End If
    Next

    ' If the current player is no disabled, go further on
    If Players.PlayerDisabled Then
        NextPlayer
    End If

End Sub

Private Sub NetGame_ClientList(aPlayers As Variant)

    Dim Tell As Long
        
    ' Clear the collection
    Set NetGame.ClientList = New Collection
    
    For Tell = LBound(aPlayers) To UBound(aPlayers)
        ' Add the client to the list
        NetGame.ClientList.Add aPlayers(Tell)
    Next
    
    ' Update the chat window
    frmChat.UpdateList
    
    ' Raise our event
    RaiseEvent ClientListRecived

End Sub

Private Sub NetGame_PlayerList(aPlayers As Variant, bResetObject As Boolean)

    Dim Tell As Long, Index As Long
    
    ' If we are going to update our player list, do so
    If bResetObject Then
    
        ' Merely remove all players
        Players.ClearPlayers
        
        ' Add each player in the list
        For Tell = LBound(aPlayers) To UBound(aPlayers) Step 2
            
            ' Add the player element
            Index = Players.AddPlayer(CStr(aPlayers(Tell)))
            
            ' Set properties
            Players.PlayerLocal(Index) = CBool(aPlayers(Tell + 1))
            
        Next
    
    End If

    ' Raise our event
    RaiseEvent PlayerListRecived

End Sub

Private Sub NetGame_Dices(aDices As Variant)

    Dim Tell As Long
    
    ' Simply change the local dices
    For Tell = LBound(aDices) To UBound(aDices) Step 2
        Dices.DiceValue(Tell / 2) = Val(aDices(Tell)) ' Change value
        Dices.DiceBackColor(Tell / 2) = Val(aDices(Tell + 1))  ' Change backcolor
    Next

End Sub

Private Sub NetGame_Message(Sender As String, Reciver As String, Message As String)

    Dim Form As Form

    ' Show the message
    If Reciver = "" Then
        ' This is a message in the global chat room
        frmChat.AddLine Sender & ": " & Message
        
    ElseIf Reciver <> Sender Then
    
        ' This is a private message, show it in a private form if don't exist already
        For Each Form In Forms
        
            ' This must be a chat form
            If TypeOf Form Is frmChat Then
            
                If Form.Reciver = Sender Then
                    
                    ' Add the line
                    Form.AddLine Sender & ": " & Message
                    
                    ' Show it
                    Form.Show
                    
                    ' We are finished
                    Exit Sub
                    
                End If
                
            End If
                        
        Next
        
        ' We didn't find our form, so make a new one
        Set Form = New frmChat
        
        ' Set the reciver and showlist
        Form.Reciver = Sender
        Form.ShowList = False
        
        ' Show the form
        Form.Show
        
        ' Add line
        Form.AddLine Sender & ": " & Message
        
    End If
    
End Sub

Private Sub Players_PlayerChanged()

    If NetGame.NetStatus = Net_Playing And NetGame.NetType = Net_Server Then
    
        ' We must inform our clients
        NetGame.Broadcast "NEXTPLAYER " & Players.CurrentPlayer & ", " & Players.PlayerThrows & ", " & IIf(Players.JustChanged, "True", "False")
    
    End If

    ' Raise our event
    RaiseEvent PlayerChanged

End Sub


