VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCommands"
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

Public Parent As clsNetGame
Public Connection As clsConnection
Public File As String

Public Sub Helo(sName As String)

    If Parent.NetType = Net_Server Then

        ' The client has introduced himself
        If Parent.FindClient(sName) < 0 Then
        
            ' Set the name of this connection to the paramenter
            Connection.name = sName
            
            ' Be nice and respond
            SendStat 200, "Name ok."
            
            ' A client has been added
            Parent.InvokeEvent 5
            
            ' Update list
            Parent.UpdateClientList
            Parent.UpdatePlayerList
        
        Else
        
            ' Tell the client that it must find another name to use
            SendStat 500, "Name not unique."
        
        End If
    
    End If
        
End Sub

Public Sub Ping()

    ' Pong the server/client
    Connection.SendData "PONG"

End Sub

Public Sub Pong()

    ' We've gotten a pong
    Parent.InvokeEvent 6, Connection.name

End Sub

Public Sub RemovePlayer(lngIndex As Long)

    Select Case Parent.NetType
    Case Net_Server
    
        ' Check and see if this player is allowed to do so
        If Parent.Parent.Players.PlayerSource(lngIndex) = Connection.name Then
        
            ' Make the request
            Parent.RemovePlayer lngIndex
        
        Else
            
            ' Player is'nt allowed to do this action
            SendStat 420, "Player not authorized"
        
        End If
    
    End Select

End Sub

Public Sub SendFile(ByVal sFile As String, ByVal Append As Boolean)

    ' Set the variable
    File = Parent.Parent.FileSystem.ValidPath(App.Path) & sFile

    ' If we're not appending and the file exists, the file must be deleted
    If Append = False And Dir(File) <> "" Then
        ' Delete file
        Kill File
    End If

End Sub

Public Sub WriteFile(ByVal Index As Long, Data As String)

    Dim Free As Long, sBuffer As String
    
    ' Get a free socket
    Free = FreeFile
    
    ' Convert the HEX-string to ASCII
    sBuffer = ConvertAscii(Data)

    ' Write the data to the file
    Open File For Binary As #Free
    
        ' Write section
        Put #Free, Index, sBuffer
    
    Close #Free
    
End Sub

' We're simply going to pass the message further on
Public Sub PrivMsg(sSender As String, sReciver As String, sMessage As String)

    Dim UserID As Long

    Select Case Parent.NetType
    Case Net_Server
    
        ' Be sure that this message isn't spoofed
        If sSender <> Connection.name Then
        
            ' We'll simply change the name to the correct one
            sSender = Connection.name
        
        End If
    
        ' Broadcast the message if there are no reciver
        If sReciver = "" Then
        
            Parent.Broadcast "PRIVMSG " & PrepareString(sSender) & ", " & PrepareString(sReciver) & ", " & PrepareString(sMessage)
        
            ' Tell the sender about the success
            SendStat 200, "Message sent."
        
        Else
        
            ' Find the user
            UserID = Parent.FindClient(sReciver)
            
            ' The userID must be a valid index
            If UserID >= 0 Then
            
                ' Send the message to the reciver
                Parent.Clients(UserID).SendData "PRIVMSG " & PrepareString(sSender) & ", " & _
                 PrepareString(sReciver) & ", " & PrepareString(sMessage)
                 
                ' Tell the sender about the success
                SendStat 200, "Message sent."
            
            Else
                
                ' Could not locate user
                SendStat 501, "User not found."
            
            End If
        
        End If
    
    End Select

    ' We'll inform about this message nevertheless
    Parent.InvokeEvent 0, sSender, sReciver, sMessage

End Sub

Public Sub Dices(sDices As String)
    
    Dim aDices As Variant, Tell As Long
    
    If Parent.NetType = Net_Server Then
    
        ' We must then broadcast this message to all clients except the one we got it from
        Parent.Broadcast "DICES " & PrepareString(sDices), Connection.name
        
    End If
    
    ' All the dice numbers are saved in an string delmimted by commas
    aDices = Split(sDices, ",")

    ' Inform about this command
    Parent.InvokeEvent 1, aDices

End Sub

' The game is getting started
Public Sub StatusChanged(lngStatus As Long)

    ' Change the status
    If Parent.NetType = Net_Client Then
        Parent.NetStatus = lngStatus
    Else
        ' This function cannot be sent to servers
        SendStat 400, "Premission denied"
    End If
    
End Sub

Public Sub GameType(ByVal sGameType As String)

    ' Simply change the game type
    If Parent.NetType = Net_Client Then
        Parent.Parent.Types.LoadData Parent.Parent.FileSystem.ValidPath(App.Path) & sGameType
    Else
        ' This function cannot be sent to servers
        SendStat 400, "Premission denied"
    End If

End Sub

Public Sub DiceResetValue(ByVal bDiceResetValue As Boolean)

    ' Simply change the property
    If Parent.NetType = Net_Client Then
        Parent.Parent.Dices.DiceResetValue = bDiceResetValue
    Else
        ' This function cannot be sent to servers
        SendStat 400, "Premission denied"
    End If

End Sub

Public Sub Querry(LineIndex As Long, CellIndex As Long, Value As String)

    ' Set the cell
    Parent.Parent.DataBase.CellText(LineIndex, CellIndex) = Value

    ' If this is a server, broadcast the querry to everyone except the original sender
    If Parent.NetType = Net_Server Then
        Parent.Broadcast "QUERRY " & LineIndex & ", " & CellIndex & ", " & PrepareString(Value), Connection.name
    End If

End Sub

' The player turn is changed
Public Sub NextPlayer(lngPlayer As Long, lngTurns As Long, JustChanged As Boolean)

    ' Simply change the player id
    If Parent.NetType = Net_Client Then
        Parent.Parent.Players.CurrentPlayer = lngPlayer
        Parent.Parent.Players.PlayerThrows = lngTurns
        Parent.Parent.Players.JustChanged = JustChanged
    Else
        ' This function cannot be sent to servers
        SendStat 400, "Premission denied"
    End If

End Sub

' A player request to go further on
Public Sub PlayerFinished()

    ' Go to next player if server
    If Parent.NetType = Net_Server Then
        Parent.Parent.NextPlayer
    End If

End Sub

Public Sub AddPlayer(sName As String)

    Dim Index As Long

    If Parent.NetType = Net_Server Then

        ' The player name must be unique
        If Parent.Parent.Players.FindPlayer(sName) < 0 Then
            
            ' Add the player name
            Index = Parent.Parent.Players.AddPlayer(sName)
        
            ' We must of course know the real source of this player
            Parent.Parent.Players.PlayerSource(Index) = Connection.name
            
            ' Update the player list to all clients
            Parent.UpdatePlayerList
            
            ' Tell the sender about the success
            SendStat 200, "Player added."
        
        Else
            ' Error adding player
            SendStat 503, "Player name not unique."
        End If

    End If

End Sub

Public Sub Stat(Number As Long, sMessage As String)

    ' Simply inform about the status message
    Parent.InvokeEvent 3, Number, sMessage

End Sub

Public Sub ClientList(sList As Variant)

    ' The client list comes as an array where the elements is delimitered with comma
    If Parent.NetType = Net_Client Then
        Parent.InvokeEvent 4, Split(CStr(sList), ",")
    Else
        ' This function cannot be sent to servers
        SendStat 400, "Premission denied"
    End If

End Sub

Public Sub PlayerList(sList As Variant)

    ' The player list comes as an array where the elements is delimitered with comma
    If Parent.NetType = Net_Client Then
        Parent.InvokeEvent 2, Split(CStr(sList), ",")
    Else
        ' This function cannot be sent to servers
        SendStat 400, "Premission denied"
    End If

End Sub

Public Function PrepareString(ByVal sData As String) As String

    ' replacing the existing with double quotation marks and replace new lines with the appropriate code
    sData = Replace(sData, Chr(34), Chr(34) & " & Chr(34) & " & Chr(34))
    sData = Replace(sData, vbCrLf, Chr(34) & " & vbCrLf & " & Chr(34))
    
    ' Enclose the string with quotation marks
    PrepareString = Chr(34) & sData & Chr(34)

End Function

Public Sub SendStat(lngNumber As Long, sMessage As String)

    ' Send the status respond
    Connection.SendData "Stat " & lngNumber & ", " & PrepareString(sMessage)

End Sub

