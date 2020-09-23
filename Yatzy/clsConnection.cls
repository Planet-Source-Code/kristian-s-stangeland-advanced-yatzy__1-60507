VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsConnection"
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

' Internal class
Public WithEvents Socket As clsSocket
Attribute Socket.VB_VarHelpID = -1
Public Index As Long
Public Name As String
Public Buffer As String
Public PingWaiting As Boolean
Public PingTime As Long
Public Trapper As Object

Public Sub SendData(sData As String)

    If Socket.State = sckConnected Then
    
        ' Raise event
        Trapper.InvokeEvent 7, Index, sData
    
        ' Send the data
        Socket.SendData sData & vbCrLf
    
    End If

End Sub

Private Sub Socket_OnClose()

    If Not Trapper Is Nothing Then
        Trapper.SocketEvent Index, "OnClose"
    End If

End Sub

Private Sub Socket_OnConnect()

    If Not Trapper Is Nothing Then
        Trapper.SocketEvent Index, "OnConnect"
    End If

End Sub

Private Sub Socket_OnConnectionRequest(ByVal requestID As Long)

    If Not Trapper Is Nothing Then
        Trapper.SocketEvent Index, "OnConnectionRequest", requestID
    End If

End Sub

Private Sub Socket_OnDataArrival(ByVal bytesTotal As Long)

    If Not Trapper Is Nothing Then
        Trapper.SocketEvent Index, "OnDataArrival", bytesTotal
    End If

End Sub

Private Sub Socket_OnSendComplete()
    
    If Not Trapper Is Nothing Then
        Trapper.SocketEvent Index, "OnSendComplete"
    End If

End Sub

Private Sub Socket_OnSendProgress(ByVal bytesSent As Long, ByVal bytesRemaining As Long)
    
    If Not Trapper Is Nothing Then
        Trapper.SocketEvent Index, "OnSendProgress"
    End If

End Sub

Private Sub Socket_OnError(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)

    If Not Trapper Is Nothing Then
        Trapper.SocketEvent Index, "OnError", Number, Description, Scode, Source, HelpFile, HelpContext
    End If

End Sub
