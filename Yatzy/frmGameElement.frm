VERSION 5.00
Begin VB.Form frmAdd 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Add element"
   ClientHeight    =   3540
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6645
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3540
   ScaleWidth      =   6645
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3840
      TabIndex        =   0
      Top             =   2760
      Width           =   2295
   End
   Begin VB.TextBox txtCode 
      Height          =   1245
      Left            =   1680
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Top             =   1200
      Width           =   4455
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   1680
      TabIndex        =   3
      Top             =   2760
      Width           =   2175
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   1680
      TabIndex        =   2
      Top             =   840
      Width           =   4455
   End
   Begin VB.TextBox txtRegion 
      Height          =   285
      Left            =   1680
      TabIndex        =   1
      Top             =   480
      Width           =   4455
   End
   Begin VB.Label lblCode 
      Caption         =   "&Code:"
      Height          =   255
      Left            =   360
      TabIndex        =   7
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label lblName 
      Caption         =   "&Name:"
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label lblRegion 
      Caption         =   "&Region:"
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   480
      Width           =   1335
   End
End
Attribute VB_Name = "frmAdd"
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

Private Sub cmdCancel_Click()

    Me.Tag = "CANCEL"
    Me.Hide

End Sub

Private Sub cmdOK_Click()

    Me.Tag = "OK"
    Me.Hide

End Sub

Private Sub Form_Initialize()
    
    ' Set all controls in the form to be what is described in the language pack
    Language.SetLanguageInForm Me

End Sub
