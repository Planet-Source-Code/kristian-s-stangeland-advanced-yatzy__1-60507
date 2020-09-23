VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 1  'vbDataSource
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsNetGame"
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

' Status codes for net games
Public Enum Net_Status
    Net_LocalGame
    Net_Joining
    Net_Playing
End Enum

' Server or client?
Public Enum Net_Type
    Net_Server
    Net_Client
End Enum

' Public accessible classes
Public WithEvents PrimarySocket As clsSocket
Attribute PrimarySocket.VB_VarHelpID = -1

' Events
Public Event DataRecived(Index As Long, Data As String)
Public Event DataBroadCast(Data As String)
Public Event DataSent(Index As Long, Data As String)
Public Event Message(Sender As String, Reciver As String, Message As String)
Public Event Dices(aDices As Variant)
Public Event PlayerList(aPlayers As Variant, bResetObject As Boolean)
Public Event ClientList(aPlayers As Variant)
Public Event Respond(Number As Long, Message As String)
Public Event DebugEvent(Text As String)
Public Event PongRecived(ClientName As String)
Public Event ClientClosed(Index As Long)
Public Event ClientAdded()
Public Event Connected()
Public Event StatusChanged()
Public Event ErrorConnect()
Public Event Disconnect()

' Client objects
Private pClients() As clsConnection

' The amout of clients in our array
Public ClientCount As Long

Public Script As New clsScript
Public Commands As New clsCommands
Public Parent As clsGame

' Variables for clients
Public Name As String
Public Connected As Boolean
Public ClientList As Collection

' Private variables
Private pStatus As Net_Status
Private pType As Net_Type
Private pBuffer As String
Private bWait As Boolean
Private bSendWait As Boolean
Private bPingWait As Boolean
Private lngPingTime As Long
Private lngRespond As Long

Public Sub InvokeEvent(EventID As Long, ParamArray Paramenters() As Variant)

    ' Invoke the event based on the event ID
    Select Case EventID
    Case 0: RaiseEvent Message(CStr(Paramenters(0)), CStr(Paramenters(1)), CStr(Paramenters(2)))
    Case 1: RaiseEvent Dices(Paramenters(0))
    Case 2: RaiseEvent PlayerList(Paramenters(0), (pType = Net_Client))
    Case 4: RaiseEvent ClientList(Paramenters(0))
    Case 5: RaiseEvent ClientAdded
    Case 6: RaiseEvent PongRecived(CStr(Paramenters(0)))
    Case 7: RaiseEvent DataSent(CLng(Val(Paramenters(0))), CStr(Paramenters(1)))
    Case 3
    
        ' We've gotten a respond, check and see if anybody is waiting for one
        If bWait Then
        
            ' Set the respond
            lngRespond = CLng(Paramenters(0))
        
            ' Don't wait anymore
            bWait = False
        
        End If
    
        ' Send response
        RaiseEvent Respond(CLng(Paramenters(0)), CStr(Paramenters(1)))
    
    End Select

End Sub

' All the events the client array recive
Public Sub SocketEvent(Index As Long, EventType As String, ParamArray Paramenters() As Variant)

    ' The buffer used to recive the data
    Dim sBuffer As String, Tell As Long, Temp As Long, aLines As Variant

    Select Case EventType
    Case "OnSendComplete"
    
        ' We're done sending the data
        If bSendWait Then
            ' Ok, tell the waiting function that we've send it
            bSendWait = False
        End If
    
    Case "OnClose"

        ' Find the deleted player and remove it
        For Tell = 1 To ClientList.Count - 1
        
            If ClientList(Tell) = pClients(Index).Name Then
                
                ' Remove it from the list
                ClientList.Remove Tell
        
                ' Remove the associated players
                For Temp = 0 To Parent.Players.PlayerCount - 1
                    If Parent.Players.PlayerSource = pClients(Index).Name Then
                        ' Disable the player so it isn't being selected
                        Parent.Players.PlayerDisabled(Temp) = True
                    End If
                Next
                
                ' Update the lists
                UpdatePlayerList
                UpdateClientList
                
                ' We've got nothing more to do
                Exit For
                
            End If
        
        Next
            
        ' Free all resources
        Set pClients(Index) = Nothing
        
        ' Raise the appropriate event
        RaiseEvent ClientClosed(Index)
    
    Case "OnDataArrival"
    
        ' Resive the data
        pClients(Index).Socket.GetData sBuffer
    
        ' Spiit the returned data into different lines
        aLines = Split(pClients(Index).Buffer & sBuffer, vbCrLf)
    
        ' If we actually have a line
        If UBound(aLines) > LBound(aLines) Then
        
            ' We have now used the buffer
            pClients(Index).Buffer = ""
        
            ' Process each line
            For Tell = LBound(aLines) To UBound(aLines) - 1
        
                ' The data has been recived
                RaiseEvent DataRecived(Index, CStr(aLines(Tell)))
            
                ' Analyze the command
                ProcessCommand pClients(Index), CStr(aLines(Tell))
            
            Next
            
        End If
        
        ' The last element is always to be saved for later
        pClients(Index).Buffer = pClients(Index).Buffer & aLines(UBound(aLines))
    
    End Select

End Sub

Public Sub RemovePlayer(lngIndex As Long)

    ' Only go further if this player actually exists
    If Parent.Players.IsPlayer(lngIndex) Then

        Select Case NetType
        Case Net_Client
            
            ' Check if player is local
            If Parent.Players.PlayerLocal(lngIndex) Then
            
                ' Send the request of removing the player
                Broadcast "REMOVEPLAYER " & lngIndex
            
            End If
        
        Case Net_Server
        
            ' If we're the server, simply add the player
            Parent.Players.RemovePlayer lngIndex
            
            ' Update the player list to all clients
            UpdatePlayerList
        
        End Select
    
    End If
    
End Sub

Public Sub AddPlayer(sName As String)

    Dim lngIndex As Long

    Select Case NetType
    Case Net_Client
    
        ' Simply ask the server to add a player
        Broadcast "ADDPLAYER " & Commands.PrepareString(sName)
    
    Case Net_Server
    
        ' If we're the server, add the player
        lngIndex = Parent.Players.AddPlayer(sName)
        
        ' Set the player source, that is, the server
        Parent.Players.PlayerSource(lngIndex) = Name
        
        ' Update the player list to all clients
        UpdatePlayerList
    
    End Select
    
End Sub

Public Function LocalPlayer() As Boolean

    If NetStatus = Net_LocalGame Then
    
        ' It is ALWAYS local
        LocalPlayer = True
    
    Else
    
        If NetType = Net_Client Then
        
            ' Return wether or not it is a local player's turn
            LocalPlayer = Parent.Players.PlayerLocal
            
        Else
        
            ' Check the source against our name
            LocalPlayer = CBool(Parent.Players.PlayerSource = Name)
            
        End If

    End If

End Function

Public Function WaitForRespond() As Long

    ' Wait until the variable is set to false
    Do While bWait
        ' Wait a bit
        Sleep 10
        DoEvents
    Loop
    
    ' Return the respond
    WaitForRespond = lngRespond

End Function

Public Sub StartWaitForRespond()

    ' Set the variable
    bWait = True

End Sub

Public Sub StartWaitForSend()

    ' Set the variable
    bSendWait = True

End Sub

Public Sub WaitForSend()
    
    ' Wait until the variable is set to false
    Do While bSendWait
        ' Wait a bit
        Sleep 10
        DoEvents
    Loop

End Sub

Public Function IsClient(ByVal Index As Long) As Boolean

    ' Return wether or not this index is valid
    IsClient = CBool(Index >= 0 And Index < ClientCount)

End Function

Public Function GenerateClientList() As String

    Dim Tell As Long
    
    ' The server will always be on the list
    GenerateClientList = Me.Name & IIf(ClientCount > 0, ",", "")
    
    For Tell = 0 To ClientCount - 1
        ' Add the name, together with the delimiter
        GenerateClientList = GenerateClientList & pClients(Tell).Name & IIf(Tell < ClientCount - 1, ",", "")
    Next

End Function

Public Sub KillClients()

    ' Erase all clients
    Erase pClients
    
    ' Erase the count
    ClientCount = 0

End Sub

Public Property Get Clients(ByVal Index As Long) As clsConnection

    ' Return the client based on the index
    If IsClient(Index) Then
        Set Clients = pClients(Index)
    End If

End Property

Public Property Let Clients(ByVal Index As Long, ByVal vNewObject As clsConnection)

    ' Return the client based on the index
    If IsClient(Index) Then
        Set pClients(Index) = vNewObject
    End If

End Property

Public Property Get NetStatus() As Net_Status

    NetStatus = pStatus

End Property

Public Property Let NetStatus(ByVal vNewValue As Net_Status)

    ' Change the status
    pStatus = vNewValue
    
    If NetType = Net_Server And vNewValue = Net_Playing Then
        ' Inform everybody that we're beginning
        Broadcast "STATUSCHANGED " & vNewValue
    End If
    
    ' Inform about the change
    RaiseEvent StatusChanged

End Property

Public Property Get NetType() As Net_Type

    NetType = pType

End Property

Public Property Let NetType(ByVal vNewValue As Net_Type)

    pType = vNewValue

End Property

Public Sub KillClient(ByVal Index As Long)

    ' Remove client
    pClients(Index).Socket.CloseSocket

End Sub

Public Sub StartServer(Optional Port As Long = 8100)

    ' Create the class object to use
    Set PrimarySocket = New clsSocket
    
    ' Tell the class that we're a client
    NetType = Net_Server
    
    ' The port to use
    PrimarySocket.LocalPort = Port
    
    ' Listen after connections
    PrimarySocket.Listen
    
    ' Update client list (the server, for starters)
    UpdateClientList
    
    ' Inform debug window
    RaiseEvent DebugEvent("Server started")

End Sub

Public Sub SendDices(sDices As String)

    ' We must be playing
    If NetStatus = Net_Playing Then
    
        ' Broadcast the message telling about our dices
        Broadcast "DICES " & Commands.PrepareString(sDices)
    
    End If

End Sub

Public Sub Connect(sServer As String, Optional Port As Long = 8100)

    ' Create the new socket
    Set PrimarySocket = New clsSocket

    ' Tell the class that we're a client
    NetType = Net_Client

    ' Connect to the server
    PrimarySocket.Connect sServer, Port
    
    ' Inform debug window
    RaiseEvent DebugEvent("Connecting")

End Sub

Public Sub CloseConnection()

    Dim Tell As Long

    ' Simply close the connection if there is one
    If Not PrimarySocket Is Nothing Then
        PrimarySocket.CloseSocket
    End If

    ' And kill the class
    Set PrimarySocket = Nothing
    
    ' Also, if server, kill the clients
    If pType = Net_Server Then
        ' Simply erase the array
        KillClients
    End If

End Sub

Public Sub Broadcast(sMessage As String, Optional IgnoreClient As String, Optional bWaitForSend As Boolean)

    Dim Tell As Long

    ' Ignore the broadcast if the message is empty
    If sMessage = "" Then
        Exit Sub
    End If

    ' Tell about the data to be broadcasted
    RaiseEvent DataBroadCast(sMessage)

    ' Two different approaches, depending of the connection role
    Select Case pType
    Case Net_Client
    
        ' Simply send the message to the server
        PrimarySocket.SendData sMessage & vbCrLf
    
    Case Net_Server

        ' Send to all clients
        For Tell = 0 To ClientCount - 1
            ' Ignore the sockets that are not initialized
            If (Not pClients(Tell).Socket Is Nothing) And (pClients(Tell).Name <> IgnoreClient Or IgnoreClient = "") Then
            
                ' Start the wait for send prosedure, if requested
                If bWaitForSend Then
                    StartWaitForSend
                End If
                
                ' Send the data
                pClients(Tell).Socket.SendData sMessage & vbCrLf
                
                ' Wait until the data is sent, if requested
                If bWaitForSend Then
                    WaitForSend
                End If
            
            End If
        Next

    End Select

End Sub

Public Function FindClient(sName As String) As Long

    Dim Tell As Long
    
    ' Go through all clients and look for the name
    For Tell = 0 To ClientCount - 1
        If Not pClients(Tell).Socket Is Nothing Then
    
            ' See if the name is what we're looking for
            If pClients(Tell).Name = sName Then
                
                ' Return the index
                FindClient = Tell
                
                ' Nothing else to do
                Exit Function
            
            End If
    
        End If
    Next

    ' No results
    FindClient = -1
    
End Function

Public Sub SetSettings(sGameType As String, bDiceSave As Boolean)

    ' Send the settings
    Broadcast "GAMETYPE " & Commands.PrepareString(sGameType), , True
    Broadcast "DICERESETVALUE " & IIf(bDiceSave, "True", "False"), , True

End Sub

Public Sub SendClientList(Connection As clsConnection)

    ' Send the list to this player
    Connection.SendData "CLIENTLIST " & Commands.PrepareString(GenerateClientList)

End Sub

Public Sub SendPlayerList(Connection As clsConnection)

    ' Send list to this player
    Connection.SendData "PLAYERLIST " & Commands.PrepareString(Parent.Players.ToString(Connection.Name))

End Sub

Public Sub SendQuerry(LineIndex As Long, CellIndex As Long, sText As String)

    ' Send querry to server/clients
    Broadcast "QUERRY " & LineIndex & ", " & CellIndex & ", " & Commands.PrepareString(sText)

End Sub

Public Sub SendMessage(sMessage As String, sReciver As String)

    Dim refConnection As New clsConnection

    ' We'll do different actions depending of the current user type
    Select Case NetType
    Case Net_Server
    
        ' Set the reference class
        Set refConnection.Socket = PrimarySocket
        refConnection.Name = Me.Name
        
        ' Inform that it is from the server
        Set Commands.Connection = refConnection
        
        ' Simply let the command class take care if this
        Commands.PrivMsg Name, sReciver, sMessage
        
        ' Clean up
        Set Commands.Connection = Nothing
        Set refConnection = Nothing
    
    Case Net_Client
    
        ' Send command to server
        Broadcast "PRIVMSG " & Commands.PrepareString(Name) & ", " & Commands.PrepareString(sReciver) & ", " & Commands.PrepareString(sMessage)
    
    End Select
    
End Sub

Public Sub UpdateClientList()

    Dim Tell As Long
    
    ' Check and see if we are the server
    If NetType = Net_Server Then
    
        ' Send the client list to all players
        For Tell = 0 To ClientCount - 1
            SendClientList pClients(Tell)
        Next

        ' Update our list as well
        InvokeEvent 4, Split(GenerateClientList, ",")
        
    End If

End Sub

Public Sub UpdatePlayerList()

    Dim Tell As Long
    
    ' Check and see if we are the server
    If NetType = Net_Server Then
    
        ' Send the player list to all players
        For Tell = 0 To ClientCount - 1
            SendPlayerList pClients(Tell)
        Next

        ' Update our list as well
        InvokeEvent 2, Split(Parent.Players.ToString(Name), ",")
        
    End If

End Sub

Public Sub SendPing(Optional Connection As clsConnection)

    ' React differently depending on what connection we are on
    Select Case NetType
    Case Net_Client
    
        ' Send the ping
        Broadcast "PING"
    
        ' Wait for an answer
        bPingWait = True
        lngPingTime = Timer
    
    Case Net_Server

        ' Send the ping to the specific client
        Connection.SendData "PING"
        Connection.PingWaiting = True
        Connection.PingTime = Timer

    End Select

End Sub

Public Sub SendFile(Connection As clsConnection, sFile As String, Optional bBroadcast As Boolean)

    Dim Index As Long, sCommand As String, FileSystem As New clsFileSystem
    
    ' Prepare command
    sCommand = "SENDFILE " & Commands.PrepareString(FileSystem.GetFileName(sFile)) & ", False"
    
    ' Firstly, tell the user that we're going to send a file
    If bBroadcast Then
        ' Broadcast the command
        Broadcast sCommand, , True
    Else
        ' Send the command to that specific user
        Connection.SendData sCommand
    End If
    
    ' Begin at the first index
    Index = 1
    
    ' Loop until we've send all
    Do
    
        ' Generate the command to send
        sCommand = FileSection(sFile, Index)
    
        ' Only proceed if there's anything to send
        If sCommand <> "" Then
        
            ' Start the waiting prosedure
            StartWaitForSend
        
            ' Send the generated command
            If bBroadcast Then
                Broadcast sCommand, , True
            Else
                Connection.SendData sCommand
            End If
            
        Else
        
            ' We're done
            Exit Do
            
        End If
    
        ' Go to next section when the data has been sent
        WaitForSend
    Loop

End Sub

Public Function FileSection(sFile As String, Index As Long) As String

    Dim sData As String
    
    ' Retrive the data
    sData = Parent.FileSystem.LoadFile(sFile, Index, 1024)

    ' Generate the write-commnand
    If sData <> "" Then
        FileSection = "WRITEFILE " & Index & ", " & Commands.PrepareString(ConvertHex(sData))
    End If

    ' Automatically increse index
    Index = Index + Len(sData)

End Function

Private Function NewClient() As Long

    Dim Tell As Long
    
    ' Firstly, check if we have any free clients
    For Tell = 0 To ClientCount - 1
        If pClients(Tell).Socket Is Nothing Then
            
            ' We have a free socket
            NewClient = Tell
            
            ' Nothing else to do
            Exit Function
            
        End If
    Next

    ' If not, allocate a new client
    ReDim Preserve pClients(ClientCount)
    
    ' Set information of this client
    Set pClients(ClientCount) = New clsConnection
    Set pClients(ClientCount).Trapper = Me
    pClients(ClientCount).Index = ClientCount
    
    ' Return the index
    NewClient = ClientCount
    
    ' Increse the client count
    ClientCount = ClientCount + 1

End Function

Private Sub ProcessCommand(Connection As clsConnection, sBuffer As String)

    ' Ignore if the buffer is empty
    If sBuffer = "" Then
        Exit Sub
    End If
    
    ' Set the current connection
    Set Commands.Connection = Connection
    
    ' Execute script to call the correct functions
    Script.Execute "Net.Commands." & sBuffer, ""
    
    ' Clean up connection
    Set Commands.Connection = Nothing
    
End Sub

Private Sub Class_Initialize()

    ' Set parents of different objects
    Set Commands.Parent = Me

End Sub

Private Sub PrimarySocket_OnClose()

    ' Nope, we're not connected
    Connected = False
    
    ' We are disconnected
    RaiseEvent Disconnect

End Sub

Private Sub PrimarySocket_OnConnect()

    ' Start the waiting prosedure
    StartWaitForRespond

    ' This must be a client, so send the helo-command
    PrimarySocket.SendData "HELO " & Commands.PrepareString(Name) & vbCrLf

    ' Wait for a respond
    If WaitForRespond <> 200 Then
        
        ' Tell about the error
        RaiseEvent ErrorConnect
        
        ' We are doomed
        Exit Sub
        
    End If
    
    ' Yup, we are connected
    Connected = True
    
    ' Tell the above class so
    RaiseEvent Connected

End Sub

Private Sub PrimarySocket_OnConnectionRequest(ByVal requestID As Long)

    Dim lngIndex As Long

    ' The server is reciving connection request. Always accept.
    If pType = Net_Server Then
    
        ' Get a index of a new socket
        lngIndex = NewClient
        
        ' Create a socket
        Set pClients(lngIndex).Socket = New clsSocket
        
        ' Accept the new connection
        pClients(lngIndex).Socket.Accept requestID
        
        ' Inform debug window
        RaiseEvent DebugEvent("Connection request " & requestID)

    End If

End Sub

Private Sub PrimarySocket_OnDataArrival(ByVal bytesTotal As Long)

    Dim Tell As Long, sBuffer As String, aLines, refConnection As New clsConnection
    
    ' We've gotten a command from the server, get the data
    PrimarySocket.GetData sBuffer

    ' Spiit the returned data into different lines
    aLines = Split(pBuffer & sBuffer, vbCrLf)

    ' If we actually have a line
    If UBound(aLines) > LBound(aLines) Then
    
        ' We have now used the buffer
        pBuffer = ""
    
        ' Process each line
        For Tell = LBound(aLines) To UBound(aLines) - 1
            
            ' The data has been recived
            RaiseEvent DataRecived(-1, CStr(aLines(Tell)))
            
            ' Make it possible for the command class to send messages back to the server
            Set refConnection.Socket = PrimarySocket
        
            ' Analyze the command
            ProcessCommand refConnection, CStr(aLines(Tell))
        
            ' Clean up connection
            Set refConnection = Nothing
        
        Next
        
    End If
    
    ' The last element is always to be saved for later
    pBuffer = pBuffer & aLines(UBound(aLines))

End Sub

Private Sub PrimarySocket_OnError(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)

    ' Tell about the error
    MsgBox "Error " & Number & ": " & Description, vbCritical, "Error"

End Sub

Private Sub PrimarySocket_OnSendComplete()

    ' We're done sending the data
    If bSendWait Then
        ' Ok, tell the waiting function that we've send it
        bSendWait = False
    End If

End Sub
