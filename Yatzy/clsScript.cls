VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsScript"
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

Public ScriptEngine As Object

Public Property Get ScriptForms() As Collection

    Set ScriptForms = Forms

End Property

Public Property Get ScriptPlugins() As Collection

    Set ScriptPlugins = Plugins

End Property

Public Function AddControl(Class As String, Name As String, Parent As Object) As Object

    ' Add a control to a form
    Set AddControl = Parent.Controls.Add(Class, Name)

End Function

Public Function Execute(sCode As String, Paramenter As Variant) As Variant

    ' Reinitialize the engine
    ScriptEngine.Reset
    
    ' Don't allow unsecure objects to be called
    ScriptEngine.UseSafeSubset = True
    ScriptEngine.AllowUI = False
    
    ' Add needed objects
    ScriptEngine.AddObject "Game", frmMain.Game
    ScriptEngine.AddObject "Form", frmMain
    ScriptEngine.AddObject "Net", frmMain.Game.NetGame
    
    ' Add the code
    ScriptEngine.AddCode "Function Value(Array)" & vbNewLine & sCode & vbNewLine & "End Function"

    ' Run the code
    Execute = ScriptEngine.Run("Value", Paramenter)

End Function

Private Sub Class_Initialize()

    ' Initialize the script
    Set ScriptEngine = CreateObject("MSScriptControl.ScriptControl")

    ' Initialize the engine
    ScriptEngine.Language = "VBScript"

End Sub

Public Sub Quit()

    ' Exit application
    ExitApp

End Sub
