Attribute VB_Name = "modStrings"
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

' Used to split up a string with a certain delimiter, and avoiding areas enclosed by quotation marks
Public Function SplitX(Text As String, Delimiter As String) As Variant

    Dim Tell As Long, Last As Long, Arr() As String, Cnt As Long
    
    If Text = "" Then
        Exit Function
    End If
    
    Last = 1
    
    Do Until Tell >= Len(Text)
        
        Tell = InStrX(Last, Text, Delimiter)
        
        If Tell = 0 Then
            Tell = Len(Text) + 1
        End If
        
        ReDim Preserve Arr(Cnt)
        
        Arr(Cnt) = Mid(Text, Last, Tell - Last)
        Cnt = Cnt + 1
        Last = Tell + 1
    Loop
    
    SplitX = Arr

End Function

' Used by the above code
Public Function InStrX(ByVal Begin As Integer, Str As Variant, Optional SearchFor As String = " ") As Integer

    Dim Tell As Long, Buff As String, OneChar As String, DontLook As Boolean
    
    For Tell = Begin To Len(Str)
    
        Buff = Mid(Str, Tell, Len(SearchFor))
        OneChar = Mid(Buff, 1, 1)
        
        If OneChar = Chr(34) Then DontLook = Not DontLook
        
        If DontLook = False And Buff = SearchFor Then
            InStrX = Tell
            Exit Function
        End If
        
    Next

End Function

Public Function ConvertHex(sData As String) As String

    Dim Tell As Long, sHex As String
    
    ' Allocate space
    ConvertHex = Space(Len(sData) * 2)
    
    ' Go through all characters and convert it to hexadecimal
    For Tell = 1 To Len(sData)
        
        ' The hexadecimal to put in
        sHex = Hex(Asc(Mid(sData, Tell, 1)))
        
        ' Make sure that it's two characters
        If Len(sHex) < 2 Then
            sHex = String(2 - Len(sHex), "0") & sHex
        End If
        
        ' Set the to characters
        Mid(ConvertHex, ((Tell - 1) * 2) + 1, 2) = sHex
    
    Next

End Function

Public Function ConvertAscii(sData As String)

    Dim Tell As Long

    ' Allocate enough space
    ConvertAscii = Space(Len(sData) / 2)

    For Tell = 1 To Len(sData) Step 2
    
        ' Convert the hexadecimal numbers to real ascii
        Mid(ConvertAscii, (Tell + 1) / 2, 1) = Chr(Val("&H" & Mid(sData, Tell, 2)))
    
    Next

End Function

