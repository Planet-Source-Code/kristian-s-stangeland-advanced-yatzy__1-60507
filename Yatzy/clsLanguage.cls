VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLanguage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

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

Private Type FormProperties
    Name As String
    Code As String
End Type

' Our parent, that is, the game object
Public Parent As clsGame
Public Constant As New Collection

' The script object that is handling the process of parsing and analyzing the language file
Private Script As New clsScript

' Private data about the language
Private FormProperties() As FormProperties

Public Sub LoadLanguagePack(strFile As String)

    Dim Free As Long, sLine As String, lIndex As Long, aLine As Variant
    
    ClearLanguagePack
    Free = FreeFile
    
    ' No current block
    lIndex = -1
    
    Open strFile For Input As #Free
    
        Do Until EOF(Free)
    
            Line Input #Free, sLine
        
            sLine = LTrim(sLine)
        
            Select Case Mid(sLine, 1, 1)
            Case "["
        
                ' Count the block
                lIndex = lIndex + 1
    
                ' Increse array
                ReDim Preserve FormProperties(lIndex)
        
                ' Set the name of this block
                FormProperties(lIndex).Name = Mid(sLine, 2, Len(sLine) - 2)
                
            Case ";"
                ' Jump over, since this is a comment
                
            Case Else ' Add the current line
                
                ' Don't go further if we haven't yet created a block to put the line into
                If lIndex >= 0 Then
                
                    ' Add all constants to the constant collection if specified
                    If LCase(FormProperties(lIndex).Name) = "constants" Then
    
                        ' Process the line
                        aLine = Split(sLine, "=", 2)
                        
                        ' Add the constant
                        Constant.Add Script.ScriptEngine.Eval(aLine(1)), Trim(aLine(0))
    
                    Else
                    
                        ' Else, add it to the current block
                        FormProperties(lIndex).Code = FormProperties(lIndex).Code & sLine & vbCrLf
                    
                    End If
                
                End If
            
            End Select
        
        Loop
        
    Close #Free

End Sub

Public Function SetLanguageInForm(Form As Object) As Long

    Dim aLines As Variant, Tell As Long, Temp As Long
    
    ' Loop through all form objects to find this form
    For Tell = LBound(FormProperties) To UBound(FormProperties)
    
        If FormProperties(Tell).Name = Form.Name Then
    
            ' Get all lines of this form
            aLines = Split(FormProperties(Tell).Code, vbCrLf)
        
            ' Go through all lines and add the needed code
            For Temp = LBound(aLines) To UBound(aLines)
                
                ' There is no point of adding the code if this is an empty line
                If aLines(Temp) <> "" Then
                
                    ' Add the code to the line
                    aLines(Temp) = "Array." & aLines(Temp)
                
                End If
                
            Next
        
            ' Execute code
            Script.Execute Join(aLines, vbCrLf), Form
        
            ' We have nothing else to do
            Exit Function
            
        End If
    
    Next

End Function

Public Function Clone() As clsLanguage

    Set Clone = New clsLanguage

End Function

Public Function EnumLanguagePacks(ByVal sFolder As String, ByVal sExtension As String) As Collection

    sFolder = Parent.FileSystem.ValidPath(sFolder)
    
    If Dir(sFolder) = "" Then
        Exit Function
    End If
    
    ' Find all language files
    Set EnumLanguagePacks = Parent.FileSystem.GetFolderList(sFolder, sExtension)

End Function

Public Sub ClearLanguagePack()

    ' Remove array, thus deleting all content
    Erase FormProperties

End Sub
