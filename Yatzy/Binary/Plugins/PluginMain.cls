VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PluginMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' The plug-in's name. This is the string that will identify your class for the application and other plugins.
Public Property Get Name() As String

    Name = "Script"

End Property

' Gets the description of the plug-in (never used)
Public Property Get Description() As String

    Description = "Gives access to the games internal classes in run time"

End Property

' Is invoked at start. Game is the main object of this program
Public Sub Initialize(Game As Object)

    ' Show the form
    frmScript.Show
    
    ' Reference the game object
    Set frmScript.Game = Game

End Sub


