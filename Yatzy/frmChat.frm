VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmChat 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Chat"
   ClientHeight    =   5340
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   7260
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   356
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   484
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin ComctlLib.ListView lstUsers 
      Height          =   4695
      Left            =   5520
      TabIndex        =   3
      Top             =   120
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   8281
      View            =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.TextBox txtMessages 
      Height          =   4695
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   120
      Width           =   5415
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "&Send"
      Height          =   285
      Left            =   6150
      TabIndex        =   1
      Top             =   4920
      Width           =   945
   End
   Begin VB.TextBox txtSend 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   4920
      Width           =   6015
   End
End
Attribute VB_Name = "frmChat"
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

' Reference to the net game class
Public NetGame As clsNetGame
Public Reciver As String
Public ShowList As Boolean

Public Sub AddLine(sLine As String)

    ' Add the new line to the text box
    txtMessages.Text = Right(txtMessages.Text, 5000) & sLine & vbCrLf

    ' Set the caret to the last position, thereby moving the scroll completely down
    txtMessages.SelStart = Len(txtMessages.Text)

End Sub

Public Sub UpdateList()

    Dim Client As Variant

    ' Clear the list over users
    lstUsers.ListItems.Clear
    
    ' Add all clients
    For Each Client In NetGame.ClientList
        lstUsers.ListItems.Add , , Client
    Next

End Sub

Private Sub cmdSend_Click()

    ' This command will send the text onto the server or all the clients
    NetGame.SendMessage txtSend.Text, Reciver
    
    If Reciver <> "" Then
        ' We won't get a message back, so add it to the text box
        AddLine NetGame.Name & ": " & txtSend.Text
    End If
    
    ' Clear the textbox and set it to the focus
    txtSend.Text = ""
    txtSend.SetFocus

End Sub

Private Sub Form_Load()

    ' Reference the class
    Set NetGame = frmMain.Game.NetGame

End Sub

Private Sub Form_Initialize()

    ' Set all controls in the form to be what is described in the language pack
    Language.SetLanguageInForm Me

End Sub

Private Sub Form_Resize()

    ' Only resize if the form isn't minimized
    If Me.WindowState <> 1 Then

        ' Resize the messages textbox
        txtMessages.Width = Me.ScaleWidth - txtMessages.Left - 8 - IIf(ShowList, lstUsers.Width, 0)
        txtMessages.Height = Me.ScaleHeight - txtMessages.Top - txtSend.Height - 16
    
        ' The send textbox must also be resized
        txtSend.Width = txtMessages.Width - cmdSend.Width - 2 + IIf(ShowList, lstUsers.Width, 0)
        txtSend.Top = txtMessages.Top + txtMessages.Height + 8
        cmdSend.Left = txtSend.Width + txtSend.Left + 2
        cmdSend.Top = txtSend.Top

        If ShowList Then
            ' Show it, thus also resizing it
            lstUsers.Visible = True
            lstUsers.Left = txtMessages.Left + txtMessages.Width
            lstUsers.Height = txtMessages.Height
        Else
            ' Hide the list of users
            lstUsers.Visible = False
        End If

    End If

End Sub

Private Sub lstUsers_ItemClick(ByVal Item As ComctlLib.ListItem)

    ' Create a conversation with this player only
    Dim frmNew As frmChat, Form As Form
    
    ' We cannot conversate with our self
    If Item.Text = NetGame.Name Then
        ' Simply stop doing anything more
        Exit Sub
    End If
    
    ' Check all open forms
    For Each Form In Forms
    
        ' This must be a chat form
        If TypeOf Form Is frmChat Then
    
            ' Check the reciver
            If Form.Reciver = Item.Text Then
            
                ' We have already opened a form with this user, show it and forget about the rest
                Form.Show
            
                ' Nothing more to do
                Exit Sub
            
            End If
        
        End If
    
    Next
    
    ' Create new form
    Set frmNew = New frmChat
    
    ' Don't show the list
    frmNew.ShowList = False
    
    ' Set the reciver
    frmNew.Reciver = Item.Text
    
    ' Show the form
    frmNew.Show
    
    
End Sub

Private Sub txtSend_KeyDown(KeyCode As Integer, Shift As Integer)

    ' If the user has pressed ENTER, send the text
    If KeyCode = 13 Then
    
        ' Send the text
        cmdSend_Click
    
        ' Remove that ugly pling
        KeyCode = 0
        
    End If

End Sub
