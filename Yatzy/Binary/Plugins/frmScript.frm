VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmScript 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Run Script"
   ClientHeight    =   4395
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   5340
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   293
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   356
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin RichTextLib.RichTextBox txtCode 
      Height          =   3615
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   6376
      _Version        =   393217
      BorderStyle     =   0
      Enabled         =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmScript.frx":0000
   End
   Begin VB.PictureBox picToolbar 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   120
      ScaleHeight     =   375
      ScaleWidth      =   5055
      TabIndex        =   0
      Top             =   3840
      Width           =   5055
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Cancel"
         Height          =   375
         Left            =   2520
         TabIndex        =   1
         Top             =   0
         Width           =   2535
      End
      Begin VB.CommandButton cmdRun 
         Caption         =   "&Run"
         Height          =   375
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   2535
      End
   End
End
Attribute VB_Name = "frmScript"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Game As Object
Public Script As Object

Private Sub cmdClose_Click()

    ' Exit application
    Unload Me

End Sub

Private Sub cmdRun_Click()

    ' Add and run code
    Script.Reset
    Script.AddObject "Game", Game
    Script.AddCode txtCode.Text

End Sub

Private Sub Form_Load()

    ' Create script engine and TLI-app
    Set Script = CreateObject("MSScriptControl.ScriptControl")

    ' Set language
    Script.Language = "VBScript"

End Sub

Private Sub Form_Resize()

    ' Only proceed if the form isn't minimized, since changing the controls at that point will cause an error
    If Me.WindowState <> 1 Then
    
        ' Resize textbox
        txtCode.Width = Me.ScaleWidth - txtCode.Left - 8
        txtCode.Height = Me.ScaleHeight - txtCode.Top - picToolbar.Height - 32
    
        ' Move tool bar
        picToolbar.Left = (Me.ScaleWidth / 2) - (picToolbar.Width / 2)
        picToolbar.Top = txtCode.Top + txtCode.Height + 16
    
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)

    ' Deallocate variables
    Set Script = Nothing
    Set Game = Nothing

End Sub
