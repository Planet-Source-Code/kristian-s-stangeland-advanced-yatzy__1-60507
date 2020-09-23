VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFileSystem"
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

Public Function GetFolderList(Path As String, ByVal Extension As String, Optional Attributers As VbFileAttribute = vbNormal) As Collection

    Dim sFile As String, AllowAll As Boolean
    
    ' Create new collection
    Set GetFolderList = New Collection
    
    ' Set variables
    sFile = Dir(ValidPath(Path), Attributers)
    
    ' Search through directory
    Do Until sFile = ""
        
        ' Only proceed if, we by the internal regular expressions, find the file to follow a specific pattern
        If sFile Like Extension Then
            GetFolderList.Add sFile
        End If
    
        ' Find the next file in the folder
        sFile = Dir
    Loop

End Function

Public Function GetFileName(File As String) As String

    GetFileName = Right(File, Len(File) - InStrRev(File, "\"))

End Function

Public Function GetFileBase(File As String) As String

    Dim Buff As String
    
    Buff = GetFileName(File)
    GetFileBase = Left(Buff, InStr(1, Buff, ".") - 1)

End Function

Public Function ValidPath(Path As String) As String

    ValidPath = Path & IIf(Right(Path, 1) = "\", "", "\")

End Function

Public Function GetExtension(File As String) As String

    GetExtension = Right(File, Len(File) - InStrRev(File, "."))

End Function

Public Sub SaveFile(Path As String, Data As String, Append As Boolean)

    Dim Free As Long
    
    ' If we're not appending and the file exist, delete it
    If Not Append And Dir(Path) <> "" Then
        Kill Path
    End If
    
    ' Get a free socket
    Free = FreeFile
    
    ' Write data to file
    Open Path For Append As Free
        Print #Free, Data
    Close Free

End Sub

Public Function LoadFile(Path As String, Optional Index As Long = 1, Optional ByVal Lenght As Long = -1) As String

    Dim Free As Long, Buff As String
    
    ' Get a free handle
    Free = FreeFile

    ' Open file
    Open Path For Binary Access Read Lock Write As #Free
        
        ' Check and see if the lenght isn't a bit to long
        If LOF(Free) - Index < Lenght Then
            ' Accommodate for the new lenght
            Lenght = LOF(Free) - Index + 1
            
            ' If the new lenght is below or equal to zero, then we really don't need to get any data
            If Lenght <= 0 Then
                Exit Function
            End If
            
        End If
        
        ' Don't proceed if the amout of data to be read is zero
        If Lenght <> 0 Then
        
            ' Allocate buffer according to the variable lenght if its not negative, if else, use lenght of file
            Buff = Space(IIf(Lenght >= 0, Lenght, LOF(Free)))
            
            ' Load data into the buffer from the specified index
            Get #Free, Index, Buff
        
        End If
        
    Close #Free
    
    ' Return file content
    LoadFile = Buff

End Function
