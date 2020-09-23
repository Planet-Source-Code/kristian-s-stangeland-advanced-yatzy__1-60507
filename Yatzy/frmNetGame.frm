VERSION 5.00
Begin VB.Form frmNetGame 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Net game"
   ClientHeight    =   6405
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   7365
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6405
   ScaleWidth      =   7365
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frameServer 
      Caption         =   "&Server:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3495
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   7095
      Begin VB.PictureBox picServer 
         BorderStyle     =   0  'None
         Height          =   2895
         Left            =   240
         ScaleHeight     =   2895
         ScaleWidth      =   6615
         TabIndex        =   2
         Top             =   480
         Width           =   6615
         Begin VB.CommandButton cmdRemovePlayer 
            Caption         =   "&Remove player"
            Height          =   375
            Left            =   4920
            TabIndex        =   19
            Top             =   2400
            Visible         =   0   'False
            Width           =   1695
         End
         Begin VB.CommandButton cmdKickPlayer 
            Caption         =   "&Kick player"
            Height          =   375
            Left            =   0
            TabIndex        =   18
            Top             =   2400
            Width           =   3135
         End
         Begin VB.CommandButton cmdAddPlayer 
            Caption         =   "&Add player"
            Height          =   375
            Left            =   3240
            TabIndex        =   17
            Top             =   2400
            Width           =   1695
         End
         Begin VB.ListBox lstPlayers 
            Height          =   2010
            Left            =   3240
            TabIndex        =   5
            Top             =   360
            Width           =   3375
         End
         Begin VB.ListBox lstClients 
            Height          =   2010
            Left            =   0
            TabIndex        =   3
            Top             =   360
            Width           =   3135
         End
         Begin VB.Label lblPlayers 
            Caption         =   "&Players:"
            Height          =   255
            Left            =   3240
            TabIndex        =   6
            Top             =   0
            Width           =   2775
         End
         Begin VB.Label lblClients 
            Caption         =   "&Clients:"
            Height          =   255
            Left            =   0
            TabIndex        =   4
            Top             =   0
            Width           =   2775
         End
      End
   End
   Begin VB.Frame frameStart 
      Caption         =   "S&tart:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   120
      TabIndex        =   0
      Top             =   3720
      Width           =   7095
      Begin VB.PictureBox picStart 
         BorderStyle     =   0  'None
         Height          =   1695
         Left            =   240
         ScaleHeight     =   1695
         ScaleWidth      =   6615
         TabIndex        =   7
         Top             =   480
         Width           =   6615
         Begin VB.CommandButton cmdStart 
            Caption         =   "&Start"
            Height          =   375
            Left            =   4680
            TabIndex        =   16
            Top             =   1320
            Width           =   1815
         End
         Begin VB.TextBox txtPort 
            Height          =   285
            Left            =   1800
            TabIndex        =   15
            Text            =   "8100"
            Top             =   720
            Width           =   4695
         End
         Begin VB.TextBox txtHostName 
            Height          =   285
            Left            =   1800
            TabIndex        =   13
            Top             =   360
            Width           =   4695
         End
         Begin VB.CommandButton cmdConnect 
            Caption         =   "&Connect"
            Height          =   375
            Left            =   2760
            TabIndex        =   11
            Top             =   1320
            Width           =   1815
         End
         Begin VB.CommandButton cmdListen 
            Caption         =   "&Listen"
            Height          =   375
            Left            =   840
            TabIndex        =   10
            Top             =   1320
            Width           =   1815
         End
         Begin VB.TextBox txtName 
            Height          =   285
            Left            =   1800
            TabIndex        =   9
            Top             =   0
            Width           =   4695
         End
         Begin VB.Label lblPort 
            Caption         =   "&Port:"
            Height          =   255
            Left            =   0
            TabIndex        =   14
            Top             =   720
            Width           =   1815
         End
         Begin VB.Label lblHostname 
            Caption         =   "&Host name:"
            Height          =   255
            Left            =   0
            TabIndex        =   12
            Top             =   360
            Width           =   1935
         End
         Begin VB.Label lblName 
            Caption         =   "&Name:"
            Height          =   255
            Left            =   0
            TabIndex        =   8
            Top             =   0
            Width           =   1935
         End
      End
   End
End
Attribute VB_Name = "frmNetGame"
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

' The mighty game class
Public WithEvents Game As clsGame
Attribute Game.VB_VarHelpID = -1

' The net class
Public WithEvents Net As clsNetGame
Attribute Net.VB_VarHelpID = -1

Private Sub cmdAddPlayer_Click()

    Dim sPlayer As String
    
    ' Get the player name
    sPlayer = InputBox(Language.Constant("WritePlayerName"))
    
    If sPlayer <> "" Then
    
        If Game.Players.FindPlayer(sPlayer) < 0 Then
    
            ' Add the player
            Net.AddPlayer sPlayer
        
        Else
        
            ' Show the error message
            MsgBox Language.Constant("PlayerUnique"), vbCritical, "Error"
            
        End If
    
    End If

End Sub

Private Sub cmdKickPlayer_Click()

    ' Only proceed if a player is selected on the list
    If lstClients.ListIndex >= 1 Then
    
        ' Kick the player
        Net.KillClient lstClients.ListIndex - 1
    
    Else
        
        ' This is most likely the server itself
        MsgBox Language.Constant("CannotKick"), vbCritical, "Error"
    
    End If

End Sub

Private Sub cmdRemovePlayer_Click()

    ' Only proceed if a player is selected on the list
    If lstPlayers.ListIndex >= 0 Then
    
        ' Remove selected item
        Net.RemovePlayer lstPlayers.ListIndex
        
    End If

End Sub

Private Sub Form_Initialize()
        
    ' Set all controls in the form to be what is described in the language pack
    Language.SetLanguageInForm Me

End Sub

Private Sub cmdConnect_Click()

    ' The name CANNOT be empty
    If txtName.Text = "" Then
        MsgBox Language.Constant("CannotBeEmpty"), vbCritical, Language.Constant("NameError")
        Exit Sub
    End If

    ' Show the chat form
    frmChat.ShowList = True
    frmChat.Show
    
    ' Set name of client/server
    Net.Name = txtName.Text

    ' Connect to the server
    Net.Connect txtHostName.Text, Val(txtPort.Text)

    ' Players are joing the channel
    Net.NetStatus = Net_Joining

    ' Enable and disable other commands
    cmdListen.Enabled = False
    cmdConnect.Enabled = False
    frameServer.Enabled = True
    cmdAddPlayer.Enabled = True
    cmdRemovePlayer.Enabled = True
    txtHostName.Enabled = False
    txtName.Enabled = False
    txtPort.Enabled = False

End Sub

Private Sub cmdListen_Click()

    ' The name CANNOT be empty
    If txtName.Text = "" Then
        MsgBox Language.Constant("CannotBeEmpty"), vbCritical, Language.Constant("NameError")
        Exit Sub
    End If

    ' Show the chat form
    frmChat.ShowList = True
    frmChat.Show
    
    ' Set name of client/server
    Net.Name = txtName.Text

    ' Start listening after connections
    Net.StartServer Val(txtPort.Text)

    ' Players are joing the channel
    Net.NetStatus = Net_Joining

    ' Enable and disable other commands
    cmdStart.Enabled = True
    cmdListen.Enabled = False
    cmdConnect.Enabled = False
    frameServer.Enabled = True
    cmdKickPlayer.Enabled = True
    cmdAddPlayer.Enabled = True
    cmdRemovePlayer.Enabled = True
    txtHostName.Enabled = False
    txtName.Enabled = False
    txtPort.Enabled = False

End Sub

Private Sub cmdStart_Click()

    Dim Tell As Long

     ' If we're supposed to save the throws
    frmMain.usrDices.DiceResetValue = (frmPlayers.cmbSaveThrows.ListIndex = 0)

    ' Use selected game type
    frmMain.Game.Types.LoadData frmMain.Game.FileSystem.ValidPath(App.Path) & frmPlayers.cmbGameTypes.Text
    
    ' Send the game file to all players
    Net.SendFile Nothing, frmMain.Game.Types.FileName, True
    
    ' Set all clients settings regarding the game type and dice reset value
    Net.SetSettings frmPlayers.cmbGameTypes.Text, frmMain.usrDices.DiceResetValue
    
    ' The game is commencing
    Net.NetStatus = Net_Playing
    
End Sub

Private Sub Form_Load()

    ' Reference the game class and net class
    Set Game = frmMain.Game
    Set Net = frmMain.Game.NetGame

    ' The state of all controls when the form is first loaded
    cmdStart.Enabled = False
    cmdConnect.Enabled = True
    cmdListen.Enabled = True
    cmdKickPlayer.Enabled = False
    cmdAddPlayer.Enabled = False
    frameServer.Enabled = False
    
    ' Clear all players
    Game.Players.ClearPlayers
    
    ' Set default properties
    cmdStart.Caption = Language.Constant("Start")
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    ' Exit connection
    Game.NetGame.CloseConnection

    ' Simply show the player form
    frmPlayers.Show

End Sub

Private Sub Game_ClientListRecived()

    Dim Client As Variant

    ' Clear list
    lstClients.Clear

    ' Update local list
    For Each Client In Net.ClientList
        lstClients.AddItem Client
    Next

End Sub

Private Sub Game_PlayerListRecived()

    Dim Tell As Long

    ' Clear list
    lstPlayers.Clear
    
    ' Update local list
    For Tell = 0 To Game.Players.PlayerCount - 1
        lstPlayers.AddItem Game.Players.PlayerName(Tell)
    Next

End Sub

Private Sub lstPlayers_KeyDown(KeyCode As Integer, Shift As Integer)

    ' If we are about to delete a player
    If KeyCode = vbKeyDelete And lstPlayers.ListIndex >= 0 Then
    
        ' Remove selected item
        Net.RemovePlayer lstPlayers.ListIndex
    
    End If

End Sub

Private Sub Net_Connected()

    ' We are connected
    frameServer.Enabled = True

End Sub

Private Sub Net_StatusChanged()

    If Net.NetStatus = Net_Playing Then

        ' Show the game form
        frmMain.Show
        
        ' Hide this form, if required
        If Net.NetType = Net_Client Then
            frmNetGame.Hide
        Else
            ' The start-command is now renamed as restart
            cmdStart.Caption = Language.Constant("Restart")
        End If
        
        ' Create a new game
        frmMain.InitializeControls
        frmMain.Game.NewGame
        
        ' Draw dices
        frmMain.usrDices.DrawAll
    
        ' Draw board
        frmMain.usrBoard.DrawAll

    End If
    
End Sub

