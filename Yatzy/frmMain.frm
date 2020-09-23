VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Yatzy"
   ClientHeight    =   6525
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8205
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   435
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   547
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdThrow 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Throw (3)"
      Height          =   375
      Left            =   4800
      TabIndex        =   3
      Top             =   4200
      Width           =   2895
   End
   Begin pYatzy.usrDices usrDices 
      Height          =   855
      Left            =   4320
      TabIndex        =   2
      Top             =   3240
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   1508
      DiceCount       =   4
      DiceWidth       =   40
      DiceHeight      =   40
      DiceSpace       =   8
      DiceStartValue  =   3
      DiceResetValue  =   -1  'True
   End
   Begin pYatzy.usrBoard usrBoard 
      Height          =   6015
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   10610
      CellHeight      =   17
      CellFrame       =   2
      FrameWidth      =   0
      FrameHeight     =   0
      VariableSize    =   -1  'True
      AutoResize      =   0   'False
   End
   Begin VB.PictureBox picYatzy 
      Height          =   2775
      Left            =   4320
      Picture         =   "frmMain.frx":08CA
      ScaleHeight     =   2715
      ScaleWidth      =   3675
      TabIndex        =   0
      Top             =   120
      Width           =   3735
   End
End
Attribute VB_Name = "frmMain"
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

' Game class
Public WithEvents Game As clsGame
Attribute Game.VB_VarHelpID = -1

Private Sub cmdThrow_Click()

    Dim Tell As Long, Wait As Long
    
    ' Only local players can throw dices
    If Game.NetGame.LocalPlayer = False Then
        Exit Sub
    End If
    
    ' Disable command button
    cmdThrow.Enabled = False
    
    ' Inform about this change
    Game.Players.JustChanged = False
    
    ' Firstly reset all unmarked dices
    usrDices.ResetDices vbYellow
    
    ' Create seed
    Randomize Timer
    
    ' Randomize all dices
    For Tell = 0 To usrDices.DiceCount
        
        ' Only throw if the dice isn't selected
        If usrDices.DiceBackColor(Tell) = vbWhite Then
        
            ' Simply give the process of throwing the dices a litte more action
            For Wait = 0 To Rnd * 10

                ' Set the dice to a random number
                usrDices.DiceValue(Tell) = 1 + Fix(Rnd * 5.99)
                
                ' Wait a bit
                Sleep 25
                DoEvents
                
            Next
        
        End If
    Next
    
    ' One less throw
    Game.Players.PlayerThrows = Game.Players.PlayerThrows - 1
    
    ' Enable command button if required
    If Game.Players.PlayerThrows > 0 Then
        cmdThrow.Enabled = True
    End If
    
    ' Send dices if required
    If Game.NetGame.NetStatus = Net_Playing Then
    
        ' Send the dices
        Game.NetGame.SendDices Game.Dices.ToString
    
    End If
        
End Sub

Public Sub InitializeControls()

    ' Reference classes
    Set Game.DataBase = usrBoard.Lines
    Set Game.DataBase.Parent = Game
    Set Game.Board = usrBoard
    Set Game.Dices = usrDices

    ' Set properties
    usrBoard.VariableSize = False
    usrBoard.CellFrame = 3

End Sub

Private Sub Form_Load()

    ' Set all controls in the form to be what is described in the language pack
    Language.SetLanguageInForm Me
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    ' Show the start game form
    frmPlayers.Show

End Sub

Private Sub Game_PlayerChanged()

    Dim Tell As Long, Player As Long

    ' Update command button
    cmdThrow.Caption = Language.Constant("Throw") & " (" & Game.Players.PlayerThrows & ")"

    ' Reset dices
    If Game.Players.JustChanged = True Then
        ' Reset command button and dices
        usrDices.AllocateDices usrDices.DiceCount
        cmdThrow.Enabled = True
        
        ' Ignore drawing
        usrBoard.IgnoreDrawing = True
        
        ' Show the current player
        For Tell = 0 To Game.DataBase.LineCount - 1
        
            ' Go through all players
            For Player = 0 To Game.Players.PlayerCount - 1
                Game.DataBase.CellBackground(Tell, Player + 1) = IIf(Player = Game.Players.CurrentPlayer, vbYellow, vbButtonFace)
            Next
        
        Next
        
        ' We're finished changing the background color
        usrBoard.IgnoreDrawing = False
        usrBoard.DrawAll
    
    End If

End Sub

Private Sub picYatzy_Click()

    ' Inform about who made this program
    MsgBox Language.Constant("MadeBy") & " Kristian S. Stangeland", vbInformation, Language.Constant("Author")

End Sub

Private Sub usrBoard_Resize()

    ' Resize the form after the control
    Me.Width = (usrBoard.Width + picYatzy.Width + 32) * Screen.TwipsPerPixelX
    Me.Height = (usrBoard.Height + 44) * Screen.TwipsPerPixelY
 
    ' Move controls
    picYatzy.Left = usrBoard.Left + usrBoard.Width + 8
    usrDices.Left = picYatzy.Left
    cmdThrow.Left = picYatzy.Left + (picYatzy.Width / 2) - (cmdThrow.Width / 2)
 
End Sub

Private Sub usrBoard_CellMouseUp(Button As Integer, Shift As Integer, LineIndex As Long, CellIndex As Long)

    Dim sCellText As String, sText As String

    ' Only respond to clicks at the first line (for now) and when the player has throwed
    If CellIndex = 0 And LineIndex > 0 And Game.Players.JustChanged = False And Game.NetGame.LocalPlayer Then
    
        ' Get the text of this cell
        sCellText = Game.DataBase.CellText(LineIndex, CellIndex)
    
        ' Ignore sum and bonus
        If Not (LCase(sCellText) = "sum:" Or LCase(sCellText) = "bonus:") Then
    
            ' Only go futher if the user hasn't something yet in the database
            If Game.DataBase.CellText(LineIndex, Game.Players.CurrentPlayer + 1) = "" Then
            
                ' Calucate the value
                sText = Game.Types.CalulateValue(Left(sCellText, Len(sCellText) - 1), usrDices.DiceArray)
                
                ' Set the value locally
                Game.DataBase.CellText(LineIndex, Game.Players.CurrentPlayer + 1) = sText
                
                If Game.NetGame.NetStatus = Net_Playing Then
                    ' Let the other databases that are connected to the client have this new data too
                    Game.NetGame.SendQuerry LineIndex, Game.Players.CurrentPlayer + 1, sText
                End If
                
                ' Next player
                Game.NextPlayer
            
            End If
        
        End If
    
    End If

End Sub

Private Sub usrDices_DiceClick(DiceIndex As Long)

    ' Only local players can interact with this control
    If Game.NetGame.LocalPlayer Then

        ' "Invert" the background color to indicate selecting
        usrDices.DiceBackColor(DiceIndex) = IIf(usrDices.DiceBackColor(DiceIndex) = vbWhite, vbYellow, vbWhite)

    End If

End Sub



