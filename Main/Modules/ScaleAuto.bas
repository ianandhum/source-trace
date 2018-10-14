Attribute VB_Name = "ScalingModule"
Private curr_obj As Object
Private x_size As Double
Private y_size As Double
Const DevEnvWidth = 1366
Const DevEnvHeight = 768
Option Explicit

'*****************************************************************************************
'                           LICENSE INFORMATION
'*****************************************************************************************
'   FormResize 3.0 by RadialApps based on FormControl Version 2.0
'   Code module for resizing a form based on screen size, then resizing the
'   controls based on the forms size
'
'   Copyright (C) 2007
'   Richard L. McCutchen
'   Email: richard@psychocoder.net
'   Created: AUG99
'
'   This program is free software: you can redistribute it and/or modify
'   it under the terms of the GNU General Public License as published by
'   the Free Software Foundation, either version 3 of the License, or
'   (at your option) any later version.
'
'   This program is distributed in the hope that it will be useful,
'   but WITHOUT ANY WARRANTY; without even the implied warranty of
'   MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'   GNU General Public License for more details.
'
'   You should have received a copy of the GNU General Public License
'   along with this program.  If not, see <http://www.gnu.org/licenses/>.
'*****************************************************************************************

Public Type ControlInitial
    Name As String
    index As Integer
    Left As Integer
    Top As Integer
    width As Integer
    height As Integer
    fontsize As Integer
End Type

Public Sub ResizeControls(frm As Form, InitialControlList() As ControlInitial, Optional AspectRatio As Boolean)
On Error Resume Next
Dim i As Integer
'   Get ratio of initial form size to current form size
x_size = frm.height / InitialControlList(0).height
y_size = frm.width / InitialControlList(0).width

If AspectRatio = True Then
    Dim minnum As Double
    If x_size > y_size Then
        minnum = y_size
    Else
        minnum = x_size
    End If
End If
                

'Loop though all the objects on the form
'Based on the upper bound of the # of controls

If AspectRatio = False Then

For i = 1 To UBound(InitialControlList)
    'Grad each control individually
    For Each curr_obj In frm
    
        'Check to make sure its the right control
        If curr_obj.Name = InitialControlList(i).Name Then
        If curr_obj.index = InitialControlList(i).index Or InitialControlList(i).index = -1 Then
            'Then resize the control
             With curr_obj
                .fontsize = InitialControlList(i).fontsize * x_size
                .Left = InitialControlList(i).Left * y_size
                .width = InitialControlList(i).width * y_size
                .height = InitialControlList(i).height * x_size
                .Top = InitialControlList(i).Top * x_size
             End With
        End If
        End If
    'Get the next control
    Next curr_obj
Next i

Else

For i = 1 To UBound(InitialControlList)
    'Grad each control individually
    For Each curr_obj In frm
    
        'Check to make sure its the right control
        If curr_obj.Name = InitialControlList(i).Name Then
        If curr_obj.index = InitialControlList(i).index Or InitialControlList(i).index = -1 Then
            'Then resize the control
             With curr_obj
                .fontsize = InitialControlList(i).fontsize * minnum
                If TypeName(.Container) = TypeName(frm) Then
                .Left = (InitialControlList(i).Left * minnum) + (((frm.width - InitialControlList(0).width) / 2) * (y_size - minnum) / y_size)
                .Top = InitialControlList(i).Top * minnum + (((frm.height - InitialControlList(0).height) / 2) * (x_size - minnum) / x_size)
                Else
                .Left = (InitialControlList(i).Left * minnum)
                .Top = InitialControlList(i).Top * minnum
                End If
                .width = InitialControlList(i).width * minnum
                .height = InitialControlList(i).height * minnum
             End With
        End If
        End If
    'Get the next control
    Next curr_obj
Next i

End If
End Sub

Public Function GetLocation(frm As Form) As ControlInitial()
On Error Resume Next
Dim InitialControlList() As ControlInitial
Dim i As Integer
i = 1
ReDim Preserve InitialControlList(0)
InitialControlList(0).Name = "TheForm"
InitialControlList(0).height = frm.height
InitialControlList(0).width = frm.width

'   Load the current positions of each object into a user defined type array.
'   This information will be used to rescale them in the Resize function
'Loop through each control
For Each curr_obj In frm
'Resize the Array by 1, and preserve
'the original objects in the array
On Error Resume Next
    ReDim Preserve InitialControlList(i)
    With InitialControlList(i)
        .Name = curr_obj.Name
        If curr_obj.index <> "" Then
            .index = curr_obj.index
        Else
            .index = -1
        End If
        .fontsize = curr_obj.fontsize
        .Left = curr_obj.Left
        .Top = curr_obj.Top
        .width = curr_obj.width
        .height = curr_obj.height
    End With
    i = i + 1
Next curr_obj
    
'   This is what the object sizes will be compared to on rescaling.
    GetLocation = InitialControlList
End Function

Public Sub ReSizePosForm(frm As Form, iHeight As Integer, iWidth As Integer, iLeft As Integer, iTop As Integer, Optional AspectRatio As Boolean = True)
Dim h_scale, w_scale As Double
h_scale = (Screen.height / Screen.TwipsPerPixelY) / DevEnvHeight
w_scale = (Screen.width / Screen.TwipsPerPixelX) / DevEnvWidth

If AspectRatio = True Then
    If h_scale > w_scale Then
        h_scale = w_scale
    Else
        w_scale = h_scale
    End If
End If

    frm.height = iHeight * h_scale
    frm.width = iWidth * w_scale
    frm.Top = iTop * h_scale
    frm.Left = iLeft * w_scale
End Sub



