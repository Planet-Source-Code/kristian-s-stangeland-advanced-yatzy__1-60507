VERSION 5.00
Begin VB.Form frmPlayers 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Start game"
   ClientHeight    =   4305
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   6615
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4305
   ScaleWidth      =   6615
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClear 
      Caption         =   "&Clear players"
      Height          =   375
      Left            =   2400
      TabIndex        =   8
      Top             =   3480
      Width           =   1815
   End
   Begin VB.CommandButton cmdStartNet 
      Caption         =   "Start &net game"
      Height          =   375
      Left            =   4320
      TabIndex        =   7
      Top             =   3480
      Width           =   1815
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "&Start local game"
      Height          =   375
      Left            =   4320
      TabIndex        =   1
      Top             =   3000
      Width           =   1815
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add player"
      Height          =   375
      Left            =   2400
      TabIndex        =   0
      Top             =   3000
      Width           =   1815
   End
   Begin VB.CommandButton cmdAddType 
      Caption         =   "&Add type"
      Height          =   375
      Left            =   480
      TabIndex        =   3
      Top             =   3480
      Width           =   1815
   End
   Begin VB.ComboBox cmbSaveThrows 
      Height          =   315
      ItemData        =   "frmPlayers.frx":0000
      Left            =   480
      List            =   "frmPlayers.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   2400
      Width           =   5655
   End
   Begin VB.ListBox lstPlayers 
      Height          =   1425
      Left            =   480
      TabIndex        =   5
      Top             =   600
      Width           =   5655
   End
   Begin VB.ComboBox cmbGameTypes 
      Height          =   315
      Left            =   480
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   2040
      Width           =   5655
   End
   Begin VB.CommandButton cmdEditType 
      Caption         =   "&Edit type"
      Height          =   375
      Left            =   480
      TabIndex        =   2
      Top             =   3000
      Width           =   1815
   End
End
Attribute VB_Name = "frmPlayers"
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

Public Sub UpdateComboBox()

    Dim File As Variant

    ' Update combo box
    cmbGameTypes.Clear
    
    For Each File In frmMain.Game.FileSystem.GetFolderList(App.Path, "*.gtf")
        cmbGameTypes.AddItem File
    Next
    
    ' Select the first element, if there are any
    If cmbGameTypes.ListCount > 0 Then
        cmbGameTypes.ListIndex = 0
    End If

End Sub

Private Sub cmdAdd_Click()

    Dim sName As String

    ' Get the name of this element
    sName = InputBox(Language.Constant("WritePlayerName"))
    
    ' Simply add a element if the name is valid
    If sName <> "" Then
        lstPlayers.AddItem sName
    End If

End Sub

Private Sub cmdAddType_Click()

    ' Edit the current type
    Set frmGameType.GameType = New clsType
    
    ' Show the form
    frmGameType.Show
    frmGameType.RefreshList

End Sub

Private Sub cmdClear_Click()

    ' Clear all players in list
    lstPlayers.Clear

End Sub

Private Sub cmdEditType_Click()

    ' Edit the current type
    Set frmGameType.GameType = New clsType
    
    ' Reference this file to the form
    frmGameType.Reference = frmMain.Game.FileSystem.ValidPath(App.Path) & cmbGameTypes.Text
    
    ' Load the type into the class
    frmGameType.GameType.LoadData frmGameType.Reference
    
    ' Show the form
    frmGameType.Show
    frmGameType.RefreshList

End Sub

Private Sub cmdStart_Click()

    Dim Tell As Long

    ' There must be at least two players
    If lstPlayers.ListCount < 2 Then
        MsgBox Language.Constant("MorePlayerNeeded"), vbCritical, "Error"
        Exit Sub
    End If

    ' Make it clear that this is not a game played over any net
    frmMain.Game.NetGame.NetStatus = Net_LocalGame

    ' If we're supposed to save the throws
    frmMain.usrDices.DiceResetValue = (cmbSaveThrows.ListIndex = 0)

    ' Use selected game type
    frmMain.Game.Types.LoadData frmMain.Game.FileSystem.ValidPath(App.Path) & cmbGameTypes.Text

    ' Remove last players
    frmMain.Game.Players.ClearPlayers

    ' Add all the players
    For Tell = 0 To lstPlayers.ListCount - 1
        frmMain.Game.Players.AddPlayer lstPlayers.List(Tell)
    Next
    
    ' Start the game
    frmPlayers.Hide
    frmMain.Show
    
    ' Create a new game
    frmMain.InitializeControls
    frmMain.Game.NewGame
    
    ' Draw dices
    frmMain.usrDices.DrawAll

    ' Draw board
    frmMain.usrBoard.DrawAll
    
End Sub

Private Sub cmdStartNet_Click()

    ' Show dialog for starting a net game
    frmNetGame.Show
    
    ' Hide this dialog
    Me.Hide

End Sub

Private Sub Form_Load()

    ' Update controls
    UpdateComboBox
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    ' Exit the application
    ExitApp

End Sub

Private Sub Form_Initialize()

    ' Set all controls in the form to be what is described in the language pack
    Language.SetLanguageInForm Me

    ' Update list box
    cmbSaveThrows.Clear
    cmbSaveThrows.AddItem Language.Constant("DoNotSaveThrows")
    cmbSaveThrows.AddItem Language.Constant("SaveThrow")
    cmbSaveThrows.ListIndex = 0
    
End Sub

Private Sub lstPlayers_KeyDown(KeyCode As Integer, Shift As Integer)

    ' Simply ignore all errors
    On Error Resume Next

    ' If the user has pressed delete
    If KeyCode = vbKeyDelete Then
    
        ' Delete the selected item
        lstPlayers.RemoveItem lstPlayers.ListIndex
    
    End If

End Sub
