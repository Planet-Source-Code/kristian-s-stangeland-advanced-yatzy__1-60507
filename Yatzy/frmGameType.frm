VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmGameType 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Edit game type"
   ClientHeight    =   4905
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   7170
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   327
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   478
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picControls 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   1320
      ScaleHeight     =   375
      ScaleWidth      =   5655
      TabIndex        =   1
      Top             =   4320
      Width           =   5655
      Begin VB.CommandButton cmdOK 
         Caption         =   "&Ok"
         Height          =   375
         Left            =   4320
         TabIndex        =   5
         Top             =   0
         Width           =   1335
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   375
         Left            =   2880
         TabIndex        =   4
         Top             =   0
         Width           =   1335
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Height          =   375
         Left            =   1440
         TabIndex        =   3
         Top             =   0
         Width           =   1335
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit"
         Height          =   375
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   1335
      End
   End
   Begin ComctlLib.ListView lstGameTypes 
      Height          =   3975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   7011
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Region"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   2
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Code"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "frmGameType"
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

' The reference
Public Reference As String
Public GameType As clsType

' Handling dialoges
Public Dialog As New clsDialog
Public OpenSave As New clsOpenSave

Public Sub RefreshList()

    Dim Tell As Long, Item As ListItem
    
    ' Clear list
    lstGameTypes.ListItems.Clear
    
    ' Add elements
    For Tell = 0 To GameType.GameCount - 1
    
        ' Add first item
        Set Item = lstGameTypes.ListItems.Add(, , GameType.GameRegion(Tell))
        
        ' Add subitems
        Item.SubItems(1) = GameType.GameName(Tell)
        Item.SubItems(2) = GameType.GameCode(Tell)
        
    Next

End Sub

Private Sub cmdAdd_Click()

    Dim RetProp As New PropertyBag

    ' Initialize the dialog
    Set Dialog.ReferenceForm = frmAdd
    
    ' Show the form
    Set RetProp = Dialog.ShowDialog(New PropertyBag, "Add element")
    
    If RetProp.ReadProperty("Returned", "") = "OK" Then
    
        ' Add the information
        GameType.AddType RetProp.ReadProperty("txtRegion", ""), RetProp.ReadProperty("txtName", ""), RetProp.ReadProperty("txtCode", ""), lstGameTypes.SelectedItem.Index
    
    End If

    ' Update the list
    RefreshList

End Sub

Private Sub cmdCancel_Click()

    ' Hide the form
    Unload Me

End Sub

Private Sub cmdEdit_Click()

    Dim RetProp As New PropertyBag, SendProp As New PropertyBag, Item As ListItem

    ' Initialize the dialog
    Set Dialog.ReferenceForm = frmAdd
    
    ' Get the selected item
    Set Item = lstGameTypes.SelectedItem
    
    ' Make ready the information to send
    SendProp.WriteProperty "txtRegion", Item.Text
    SendProp.WriteProperty "txtName", Item.SubItems(1)
    SendProp.WriteProperty "txtCode", Item.SubItems(2)
    
    ' Show the form
    Set RetProp = Dialog.ShowDialog(SendProp, "Edit element")
    
    If RetProp.ReadProperty("Returned", "") = "OK" Then
    
        ' Change the information
        GameType.GameRegion(Item.Index - 1) = RetProp.ReadProperty("txtRegion", "")
        GameType.GameName(Item.Index - 1) = RetProp.ReadProperty("txtName", "")
        GameType.GameCode(Item.Index - 1) = RetProp.ReadProperty("txtCode", "")
    
    End If

    ' Update the list
    RefreshList

End Sub

Private Sub cmdOK_Click()

    ' If there is a reference, save the type
    If Reference <> "" Then
    
        ' Save the data
        GameType.SaveData Reference
    
    Else
    
        ' If not, let the user desice what it should be called
        OpenSave.SaveFile Me.hwnd, "Save game"
        
        ' If the user hasn't pressed cancel
        If OpenSave.File <> "" Then
            
            ' Save the file
            GameType.SaveData OpenSave.File
            
        End If
    
    End If
    
    ' Update the combo box in the players form
    frmPlayers.UpdateComboBox
    
    ' Close the form
    Unload Me

End Sub

Private Sub Form_Resize()

    ' Only resize if we aren't minimize
    If Me.WindowState <> 1 Then

        ' Resize the list box
        lstGameTypes.Width = Me.ScaleWidth - lstGameTypes.Left - 8
        lstGameTypes.Height = Me.ScaleHeight - lstGameTypes.Top - picControls.Height - 32
    
        ' Move control box
        picControls.Left = Me.ScaleWidth - picControls.Width - 8
        picControls.Top = Me.ScaleHeight - picControls.Height - 16
    
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)

    ' Enable the other form
    frmPlayers.Enabled = True

End Sub

Public Sub RemoveSelected(ListView As ListView)

    Dim Tell As Long
    
    ' Delete all selected
    For Tell = 1 To ListView.ListItems.Count
    
        ' Check and see if we've reach the boundaries of all the elements
        If Tell > ListView.ListItems.Count Then
            Exit For
        End If
    
        ' If the item is selected, ..
        If ListView.ListItems(Tell).Selected Then
        
            ' ... remove the item in the database
            GameType.RemoveType Tell - 1
        
        End If
    
    Next

    ' Update the list view
    RefreshList

End Sub

Private Sub lstGameTypes_KeyDown(KeyCode As Integer, Shift As Integer)

    ' If the user pressed the delete key ...
    If KeyCode = vbKeyDelete Then
    
        ' ... delete the selected items
        RemoveSelected lstGameTypes
    
    End If

End Sub

Private Sub Form_Initialize()
    
    ' Set all controls in the form to be what is described in the language pack
    Language.SetLanguageInForm Me

End Sub
