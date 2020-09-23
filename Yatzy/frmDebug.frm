VERSION 5.00
Begin VB.Form frmDebug 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Debug"
   ClientHeight    =   4740
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   6315
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   316
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   421
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtDebug 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4695
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   0
      Width           =   6255
   End
End
Attribute VB_Name = "frmDebug"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
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

Public WithEvents NetGame As clsNetGame
Attribute NetGame.VB_VarHelpID = -1

Public Sub AddLine(sLine As String)

    ' Add the line
    txtDebug.Text = txtDebug.Text & sLine & vbCrLf

    ' Scroll to the last line
    txtDebug.SelStart = Len(txtDebug.Text)

End Sub

Private Sub Form_Initialize()
        
    ' Set all controls in the form to be what is described in the language pack
    Language.SetLanguageInForm Me

End Sub

Private Sub Form_Resize()

    ' Only proceed if the window is visible
    If Me.WindowState <> 1 Then
        ' Change the size of the textbox so it filles the form
        txtDebug.Width = Me.ScaleWidth
        txtDebug.Height = Me.ScaleHeight
    End If

End Sub

Private Sub NetGame_ClientClosed(Index As Long)

    ' Inform about this event
    AddLine "Client " & Index & " closed"

End Sub

Private Sub NetGame_Connected()

    ' Add the event
    AddLine "Connected"

End Sub

Private Sub NetGame_DataBroadCast(Data As String)

    ' Add the broadcasted data
    AddLine "<" & Data

End Sub

Private Sub NetGame_DataRecived(Index As Long, Data As String)

    ' Add the recived data
    AddLine Index & "> " & Data

End Sub

Private Sub NetGame_DataSent(Index As Long, Data As String)

    ' Add the sent data
    AddLine Index & "<" & Data

End Sub

Private Sub NetGame_DebugEvent(Text As String)

    ' Add the event
    AddLine Text
    
End Sub

Private Sub NetGame_Disconnect()

    ' Add the event
    AddLine "Disconnected"

End Sub

Private Sub NetGame_ErrorConnect()

    ' Add the event
    AddLine "Error connecting"

End Sub
