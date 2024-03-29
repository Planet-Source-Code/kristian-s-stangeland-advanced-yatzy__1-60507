VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Copyright (C) 2004 Kristian. S. Stangeland

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

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public ReferenceForm As Form

Public Function ShowDialog(Properties As PropertyBag, Optional Caption As String) As PropertyBag

    Dim txtBox As Object

    ' Initialize form
    ReferenceForm.Tag = ""
    
    ' If we have receved a caption, use it
    If Caption <> "" Then
        ReferenceForm.Caption = Caption
    End If
    
    ' Create a new PropertyBag
    Set ShowDialog = New PropertyBag
    
    ' Loop through all textboxes and give them the passed property
    For Each txtBox In ReferenceForm.Controls
        If TypeOf txtBox Is TextBox Then
            txtBox.Text = Properties.ReadProperty(txtBox.Name, "")
        End If
    Next
        
    ' Show form
    ReferenceForm.Show
    
    ' Wait for it to finish
    Do Until ReferenceForm.Tag <> ""
        Sleep 10
        DoEvents
    Loop
    
    ' Write what the form "returned"
    ShowDialog.WriteProperty "Returned", UCase(ReferenceForm.Tag), ""
    
    ' The user has pressed OK
    If UCase(ReferenceForm.Tag) = "OK" Then
    
        ' Add all textboxes to the propertybag
        For Each txtBox In ReferenceForm.Controls
            If TypeOf txtBox Is TextBox Then
                ShowDialog.WriteProperty txtBox.Name, txtBox.Text, ""
            End If
        Next
    
    End If

End Function



