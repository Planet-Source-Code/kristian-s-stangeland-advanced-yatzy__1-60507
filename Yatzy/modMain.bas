Attribute VB_Name = "modMain"
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

Declare Function InitCommonControlsEx Lib "comctl32.dll" (iccex As tagInitCommonControlsEx) As Boolean
Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Type tagInitCommonControlsEx
   lngSize As Long
   lngICC As Long
End Type
   
Private Const ICC_USEREX_CLASSES = &H200

' Plugins
Public Plugins As New Collection
Public Language As clsLanguage

Public Sub Main()
    
    ' Necessary for XP-style
    InitCommonControlsVB
    
    ' Create the game
    Set frmMain.Game = New clsGame

    ' Load language file
    LoadLanguage frmMain.Game
    
    ' Make this class public
    Set Language = frmMain.Game.Language

    ' Set icons
    Set frmPlayers.Icon = frmMain.Icon
    Set frmNetGame.Icon = frmMain.Icon

    ' Initialize different classes and properties
    frmMain.InitializeControls

    ' Load plugins
    LoadPlugins
    
    ' Process all commands
    ProcessCommands Command$
    
    ' Show the start game form
    frmPlayers.Show

End Sub

Public Sub LoadLanguage(Game As clsGame)

    ' Load the current language file
    Game.Language.LoadLanguagePack Game.FileSystem.ValidPath(App.Path) & GetSetting("Yatzy", "General", _
     "Language", Game.Language.EnumLanguagePacks(App.Path, "*.lpk")(1))

End Sub

Public Sub ProcessCommands(sCommands As String)

    Dim aElements As Variant, aCommand As Variant, Tell As Long
    
    If sCommands = "" Then
        ' There must be something to process
        Exit Sub
    End If
    
    ' Split elements by the space character - do not split inside brackets enclosed with qoutation marks
    aElements = SplitX(sCommands, "/")
    
    ' Go through all elements
    For Tell = LBound(aElements) To UBound(aElements)
    
        ' Only continue if there is anything in the element
        If Len(aElements(Tell)) > 0 Then
    
            ' Split the command
            aCommand = SplitX(CStr(aElements(Tell)), " ")
        
            ' Do the appropriate action
            Select Case aCommand(0)
            Case "debug"
                
                ' Show and start the debug window
                frmDebug.Show
                
                ' Set the reference to the main net-game class
                Set frmDebug.NetGame = frmMain.Game.NetGame
            
            End Select
        
        End If
    
    Next

End Sub

Public Sub ExitApp()

    Dim Form As Form

    ' Clear class
    Set frmMain.Game = Nothing

    ' Remove all forms
    For Each Form In Forms
        Unload Form
    Next

End Sub

Public Sub LoadPlugins()

    Dim sFile As Variant, Plugin As Object, strClassName As String, FileSystem As clsFileSystem
    
    ' Get the class handling communication to the file system
    Set FileSystem = frmMain.Game.FileSystem
    
    ' Find all files in the plugins folder
    For Each sFile In FileSystem.GetFolderList(App.Path & "\Plugins\", "*.dll")
        
        ' Only load dll-files
        If FileSystem.GetExtension(CStr(sFile)) = "dll" Then
    
            ' Clear all errors
            Err.Clear
            
            ' How the class is registered in the registry
            strClassName = FileSystem.GetFileBase(CStr(sFile)) & ".PluginMain"
            
            ' Try to create the object
            Set Plugin = CreateObject(strClassName)
    
            If Err = 429 Then ' ERROR: ActiveX component can't create object
                ' Try to register the object
                Shell "regsvr32 " & Chr(34) & FileSystem.ValidPath(App.Path) & "Plugins\" & CStr(sFile) & Chr(34)
                
                ' Load the plugin again
                Set Plugin = CreateObject(strClassName)
            End If
    
            ' Add plugin
            Plugins.Add Plugin, Plugin.Name
    
            ' Initialize plugin
            Plugin.Initialize frmMain.Game
        
        End If
        
    Next

End Sub

Public Function InitCommonControlsVB() As Boolean

   On Local Error Resume Next
   Dim iccex As tagInitCommonControlsEx
   
   ' Ensure CC available:
   With iccex
       .lngSize = LenB(iccex)
       .lngICC = ICC_USEREX_CLASSES
   End With
   
   InitCommonControlsEx iccex
   InitCommonControlsVB = (Err.Number = 0)

End Function
