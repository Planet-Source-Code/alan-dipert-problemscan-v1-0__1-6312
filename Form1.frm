VERSION 5.00
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ProblemScan 1 by Alan Dipert"
   ClientHeight    =   4170
   ClientLeft      =   4485
   ClientTop       =   4050
   ClientWidth     =   3120
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4170
   ScaleWidth      =   3120
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      BackColor       =   &H000080FF&
      Caption         =   "Info"
      Height          =   375
      Left            =   2280
      MaskColor       =   &H000080FF&
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3720
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H000080FF&
      Caption         =   "Scan for Trojan Server"
      Height          =   375
      Left            =   120
      MaskColor       =   &H000080FF&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3720
      Width           =   2055
   End
   Begin VB.DriveListBox Drive1 
      Appearance      =   0  'Flat
      BackColor       =   &H00004000&
      ForeColor       =   &H000080FF&
      Height          =   315
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   975
   End
   Begin VB.DirListBox Dir1 
      Appearance      =   0  'Flat
      BackColor       =   &H00004000&
      ForeColor       =   &H000080FF&
      Height          =   1440
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   2895
   End
   Begin VB.FileListBox File1 
      Appearance      =   0  'Flat
      BackColor       =   &H00004000&
      ForeColor       =   &H000080FF&
      Height          =   1395
      Left            =   120
      TabIndex        =   1
      Top             =   1920
      Width           =   2895
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00004000&
      ForeColor       =   &H000080FF&
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   3360
      Width           =   2895
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Scan"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   2400
      TabIndex        =   7
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   "Problem"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   375
      Left            =   1200
      TabIndex        =   5
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ooooooo      ooo           ooo
' oo    oo    oo oo         oo oo
' oo    oo   oo   oo       oo   oo
' ooooooo   ooooooooo     ooooooooo
' oo       oo       oo   oo       oo
' oo      oo         oo oo         oo
'   Programmer's Association of AOL
'        http://paa.11net.com

'###########################################
'# Info                                    #
'#                                         #
'# Creator: Alan Dipert                    #
'# Date: Feb 12, 2000                      #
'###########################################
'# Description                             #
'#                                         #
'# This example shows you how to make a    #
'# primitive trojan/virus scanner, and     #
'# how to break files down into binary and #
'# search them for certain strings.        #
'###########################################
'# How To Use                              #
'#                                         #
'# Just run it, and select a file.         #
'# You can also find trojans yourself, and #
'# get strings from them to add to the list#
'# of trojans that this program scans for. #
'# Enjoy!                                  #
'###########################################

Private Sub Command1_Click()
Screen.MousePointer = vbHourglass
    'ProblemScan v1.0 by Alan Dipert
    'Currently Scans for the following trojan servers:
        'NetBus v1.6
        'NetBus v1.7
        'NetBus v2.1 Pro
        'Subseven v1.8
        'Subseven v1.9
        'Subseven v2.0
        'Subseven v2.1
        'Subseven v2.1 GOLD
        'Back Orifice v1.2
        'Back Orifice v1.2 Encrypted variant
        'Back Orifice v2000
    
    'buffer and binary breakdown variables
    Dim filename As String, buffer As String
    
    'message box variables
    Dim resultsMSG As String, msgResult As VbMsgBoxResult, deletion As String
    
    
    'boolean declarations - "Are these strings in the file?"
    Dim netbus16total As Integer, netbus17total As Integer, netbus21total As Integer
    Dim subseven18total As Integer, subseven19total As Integer, subseven20total As Integer
    Dim subseven21total As Integer, subseven21goldtotal As Integer, bo12total As Integer, bo2ktotal As Integer
    Dim bo12encryptedtotal As Integer
    
    'NetBus versions
    Dim netbus16 As Boolean, netbus17 As Boolean, netbus21 As Boolean
    
    'SubSeven versions
    Dim subseven18 As Boolean, subseven19 As Boolean, subseven20 As Boolean, subseven21 As Boolean, subseven21gold As Boolean
    
    'Back Orifice versions
    Dim bo12 As Boolean, bo2k As Boolean, bo12encrypted As Boolean
    
    'other
    
    'begin NetBus series scan sequence
    Dim server As String
    netbus16total = 0
    server = "No trojan server was found."
    deletion = "The file has been deleted."
    filename = Text1.Text
    If filename = "" Then
        MsgBox "Please select a file to scan.", vbOKOnly, Scanner
        GoTo end1
    End If
    Open filename For Binary As 1
        Do While Not EOF(1)
            buffer = Space(4096)
            Get 1, , buffer
            DoEvents
            If InStr(buffer, "����Y��E�������U��") Then 'the weird looking code is a netbus v1.6 specific line of code
                netbus16total = netbus16total + 1
            End If
            If InStr(buffer, "S�ء�tE �8 t���@���") Then
                netbus16total = netbus16total + 1
            End If
            If InStr(buffer, "����M�3�3���}���E�") Then
                netbus16total = netbus16total + 1
            End If
            If InStr(buffer, "�Q�=�U��R�E���Q") Then
                netbus16total = netbus16total + 1
            End If
        Loop
    Close 1
    'if the above strings were found a total of 5 times, then the file is netbus v1.6
    If netbus16total = 4 Then
        netbus16 = True
        Screen.MousePointer = vbNormal
    End If
    
    netbus17total = 0
    Open filename For Binary As 1
        Do While Not EOF(1)
            buffer = Space(4096)
            Get 1, , buffer
            DoEvents
            If InStr(buffer, "�E����������������_^") Then
                netbus17total = netbus17total + 1
            End If
            If InStr(buffer, "���-v��P���������w") Then
                netbus17total = netbus17total + 1
            End If
            If InStr(buffer, "���U����SV�M�U��E�") Then
                netbus17total = netbus17total + 1
            End If
            If InStr(buffer, "����F�C��t�V�P") Then
                netbus17total = netbus17total + 1
            End If
        Loop
    Close 1
    If netbus17total = 4 Then
        netbus17 = True
        Screen.MousePointer = vbNormal
    End If
    
    netbus21total = 0
    Open filename For Binary As 1
        Do While Not EOF(1)
            buffer = Space(4096)
            Get 1, , buffer
            DoEvents
            If InStr(buffer, "�������q%����������") Then
                netbus21total = netbus21total + 1
            End If
            If InStr(buffer, "y���ڋ�3ҋ���u���Ƅ�") Then
                netbus21total = netbus21total + 1
            End If
            If InStr(buffer, "�E�P��R��f����u���") Then
                netbus21total = netbus21total + 1
            End If
            If InStr(buffer, "���U�E��6���CN���") Then
                netbus21total = netbus21total + 1
            End If
        Loop
    Close 1
    If netbus21total = 4 Then
        netbus21 = True
        Screen.MousePointer = vbNormal
    End If
    
    
    'begin Subseven series scan sequence
    subseven18total = 0
    Open filename For Binary As 1
        Do While Not EOF(1)
            buffer = Space(4096)
            Get 1, , buffer
            DoEvents
            If InStr(buffer, "�����Kwp��-ĺ�y�$��u:f��At") Then
                subseven18total = subseven18total + 1
            End If
            If InStr(buffer, "���EnumDisplayN<��h�") Then
                subseven18total = subseven18total + 1
            End If
            If InStr(buffer, "���z%}_�Qh`�B`K+Ť;�p��ta��ɖzB��cD�oWxC�B") Then
                subseven18total = subseven18total + 1
            End If
            If InStr(buffer, "T�{�B��bQ��hV@�LDWD�����8") Then
                subseven18total = subseven18total + 1
            End If
        Loop
    Close 1
    If subseven18total = 4 Then
            subseven18 = True
            Screen.MousePointer = vbNormal
    End If
    
        
    subseven19total = 0
    Open filename For Binary As 1
        Do While Not EOF(1)
            buffer = Space(4096)
            Get 1, , buffer
            DoEvents
            If InStr(buffer, "#�w���d�r�9��G�") Then
                subseven19total = subseven19total + 1
            End If
            If InStr(buffer, "W���H5S���c��") Then
                subseven19total = subseven19total + 1
            End If
            If InStr(buffer, "�ٲ���v��$�Pj��V^�K*5tj���S") Then
                subseven19total = subseven19total + 1
            End If
            If InStr(buffer, "����1���|>�%�@") Then
                subseven19total = subseven19total + 1
            End If
        Loop
    Close 1
    If subseven19total = 4 Then
            subseven19 = True
            Screen.MousePointer = vbNormal
    End If
        

    subseven20total = 0
    Open filename For Binary As 1
        Do While Not EOF(1)
            buffer = Space(4096)
            Get 1, , buffer
            DoEvents
            If InStr(buffer, "�Sec���WJLvhU") Then
                subseven20total = subseven20total + 1
            End If
            If InStr(buffer, "⎆�u�QoP:խ�!") Then
                subseven20total = subseven20total + 1
            End If
            If InStr(buffer, "�UD.3I���H�pe��") Then
                subseven20total = subseven20total + 1
            End If
            If InStr(buffer, "a*��ˊc�@�c4���;") Then
                subseven20total = subseven20total + 1
            End If
        Loop
    Close 1
    If subseven20total = 4 Then
            subseven20 = True
            Screen.MousePointer = vbNormal
    End If
    If subseven20total = 8 Then
            subseven20 = True
            Screen.MousePointer = vbNormal
    End If
    
    
        
    subseven21total = 0
    Open filename For Binary As 1
        Do While Not EOF(1)
            buffer = Space(4096)
            Get 1, , buffer
            DoEvents
            If InStr(buffer, "`��#(4d�'��_��+�~Z����Z\�") Then
                subseven21total = subseven21total + 1
            End If
            If InStr(buffer, "V��j������S�3��,E������") Then
                subseven21total = subseven21total + 1
            End If
            If InStr(buffer, "Fш��r���jL���") Then
                subseven21total = subseven21total + 1
            End If
            If InStr(buffer, "_���K%���$W�!��L") Then
                subseven21total = subseven21total + 1
            End If
        Loop
    Close 1
    If subseven21total = 4 Then
            subseven21 = True
            Screen.MousePointer = vbNormal
    End If
    
    subseven21goldtotal = 0
    Open filename For Binary As 1
        Do While Not EOF(1)
            buffer = Space(4096)
            Get 1, , buffer
            DoEvents
            If InStr(buffer, "�Y#rCX6��C��Left") Then
                subseven21goldtotal = subseven21goldtotal + 1
            End If
            If InStr(buffer, "�`P+��d�<ۜ�yg") Then
                subseven21goldtotal = subseven21goldtotal + 1
            End If
            If InStr(buffer, "��<��Q)O���gm�s8�") Then
                subseven21goldtotal = subseven21goldtotal + 1
            End If
            If InStr(buffer, "�r�k�1r���J���###") Then
                subseven21goldtotal = subseven21goldtotal + 1
            End If
        Loop
    Close 1
    If subseven21goldtotal = 4 Then
            subseven21gold = True
            Screen.MousePointer = vbNormal
    End If
    
    
    'begin back orifice series scan sequence
    bo12total = 0
    Open filename For Binary As 1
        Do While Not EOF(1)
            buffer = Space(4096)
            Get 1, , buffer
            DoEvents
            If InStr(buffer, "�E؍�0���S�U�P�u�Q") Then
                bo12total = bo12total + 1
            End If
            If InStr(buffer, "���$B P��   ��W��$B �����") Then
                bo12total = bo12total + 1
            End If
            If InStr(buffer, "s�ƃ������0") Then
                bo12total = bo12total + 1
            End If
            If InStr(buffer, "[��Ët$�l$V�z����U�z����W�") Then
                bo12total = bo12total + 1
            End If
        Loop
    Close 1
    If bo12total = 3 Then
            bo12 = True
            Screen.MousePointer = vbNormal
    End If
    
    bo12encryptedtotal = 0
    Open filename For Binary As 1
        Do While Not EOF(1)
            buffer = Space(4096)
            Get 1, , buffer
            DoEvents
            If InStr(buffer, "��P�U�RW��������V�") Then
                bo12encryptedtotal = bo12encryptedtotal + 1
            End If
            If InStr(buffer, "�E�u3��$��@t��u") Then
                bo12encryptedtotal = bo12encryptedtotal + 1
            End If
            If InStr(buffer, "�������7���C����:�") Then
                bo12encryptedtotal = bo12encryptedtotal + 1
            End If
            If InStr(buffer, "��t,�t7t���t���t�") Then
                bo12encryptedtotal = bo12encryptedtotal + 1
            End If
        Loop
    Close 1
    If bo12encryptedtotal = 4 Then
            bo12encrypted = True
            Screen.MousePointer = vbNormal
    End If
    
    bo2ktotal = 0
    Open filename For Binary As 1
        Do While Not EOF(1)
            buffer = Space(4096)
            Get 1, , buffer
            DoEvents
            If InStr(buffer, "�u�E�P�E�P�E�P�E�P�E�P�E�P�E�P�E�P�u��u���") Then
                bo2ktotal = bo2ktotal + 1
            End If
            If InStr(buffer, "��;�~b�|���@}W����B�������|��h") Then
                bo2ktotal = bo2ktotal + 1
            End If
            If InStr(buffer, "����u��u������h�") Then
                bo2ktotal = bo2ktotal + 1
            End If
            If InStr(buffer, "�uPPPPjS�Ջ�S��13�;�u;��") Then
                bo2ktotal = bo2ktotal + 1
            End If
        Loop
    Close 1
    If bo2ktotal = 4 Then
            bo2k = True
            Screen.MousePointer = vbNormal
    End If
    

    'netbus message box
    If netbus16 = True Then server = "The scanned file is the NetBus v1.6 server." & Chr(10) & " It is highly recommended that you delete this file immediately." & Chr(10)
    If netbus17 = True Then server = "The scanned file is the NetBus v1.7 server." & Chr(10) & " It is highly recommended that you delete this file immediately." & Chr(10)
    If netbus21 = True Then server = "The scanned file is the NetBus v2.10 Pro server." & Chr(10) & " It is highly recommended that you delete this file immediately." & Chr(10)
    
    
    'subseven message box
    If subseven18 = True Then server = "The scanned file is the Subseven v1.8 server." & Chr(10) & " It is highly recommended that you delete this file immediately." & Chr(10)
    If subseven19 = True Then server = "The scanned file is the Subseven v1.9 server." & Chr(10) & " It is highly recommended that you delete this file immediately." & Chr(10)
    If subseven20 = True Then server = "The scanned file is the Subseven v2.0 server." & Chr(10) & " It is highly recommended that you delete this file immediately." & Chr(10)
    If subseven21 = True Then server = "The scanned file is the Subseven v2.1 server." & Chr(10) & " It is highly recommended that you delete this file immediately." & Chr(10)
    If subseven21gold = True Then server = "The scanned file is the Subseven v2.1 Gold server." & Chr(10) & " It is highly recommended that you delete this file immediately." & Chr(10)
    
    'back orifice message box
    If bo12 = True Then server = "The scanned file is the Back Orifice v1.2 server." & Chr(10) & " It is highly recommended that you delete this file immediately." & Chr(10)
    If bo2k = True Then server = "The scanned file is the Back Orifice 2000 server." & Chr(10) & " It is highly recommended that you delete this file immediately." & Chr(10)
    If bo12encrypted = True Then server = "The scanned file is an encrypted form of the Back Orifice v1.2 Server." & Chr(10) & " It is highly recommended that you delete this file immediately." & Chr(10)
    
    If server = "No trojan server was found." Then
        Screen.MousePointer = vbNormal
        MsgBox server, vbInformation, "Scanner"
    Else
        msgResult = MsgBox(server & " Would you like to delete this file?", vbCritical Or vbYesNo, "Scanner")
        If msgResult = vbYes Then
            Kill filename
            MsgBox deletion, vbInformation, "File Deleted"
            'deletes scanned file if it is a trojan and user selects "Yes, I want to delete"
        End If
    End If
end1:
        Screen.MousePointer = vbNormal 'this returns the mouse pointer to its normal state
End Sub

Private Sub Command2_Click()
Form2.Show
End Sub

Private Sub Dir1_Change()
    File1.Path = Dir1
    ChDir File1.Path
End Sub

Private Sub Drive1_Change()
    Dir1 = Drive1
End Sub

Private Sub File1_Click()
    If Right(File1.Path, 1) = "\" Then
        Text1.Text = File1.Path & File1
    Else
        Text1.Text = File1.Path & "\" & File1
    End If
    filepath = Text1.Text
End Sub

Private Sub File1_DblClick()
    Command1_Click
End Sub

