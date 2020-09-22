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
            If InStr(buffer, "¯şÿÿYë‹EüèàıÿÿëUèÄ") Then 'the weird looking code is a netbus v1.6 specific line of code
                netbus16total = netbus16total + 1
            End If
            If InStr(buffer, "S‹Ø¡¼tE €8 t‹Ãè@õÿÿ") Then
                netbus16total = netbus16total + 1
            End If
            If InStr(buffer, "ıÿëMà3Ò3ÀèÏ}ıÿ‹Eè") Then
                netbus16total = netbus16total + 1
            End If
            If InStr(buffer, "ÿQë=‹Uü‹R‹Eø‹ÿQ") Then
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
            If InStr(buffer, "EøèÏËşÿÃééÅşÿëğ‹Æ_^") Then
                netbus17total = netbus17total + 1
            End If
            If InStr(buffer, "‹Ãè-vÿÿPèÿ¼ıÿë‹Ãè¶w") Then
                netbus17total = netbus17total + 1
            End If
            If InStr(buffer, "U‹ìƒÄèSV‰Mô‰Uø‰Eü") Then
                netbus17total = netbus17total + 1
            End If
            If InStr(buffer, "Àƒà‰F‹C…Àt‹V‰P") Then
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
            If InStr(buffer, "•øşÿÿ‹èq%ÿÿ•øşÿÿ‹Ç") Then
                netbus21total = netbus21total + 1
            End If
            If InStr(buffer, "yşÿ‹Ú‹ğ3Ò‹ÆèÑuşÿ‹Æ„Û") Then
                netbus21total = netbus21total + 1
            End If
            If InStr(buffer, "‹E‹Pü‹R‹Æf»ñÿèuºûÿ") Then
                netbus21total = netbus21total + 1
            End If
            If InStr(buffer, "ÿÿUô‹Eüè6ÅÿÿCN…¬ş") Then
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
            If InStr(buffer, "ÒúÙáèKwpüÛ-Äº™yØ$›İu:f÷ğAt") Then
                subseven18total = subseven18total + 1
            End If
            If InStr(buffer, "äªò€EnumDisplayN<ùÀhø") Then
                subseven18total = subseven18total + 1
            End If
            If InStr(buffer, "ûÏøz%}_äQh`úB`K+Å¤;şp¥‹ta»ÇÉ–zBÆècDä‚oWxC›B") Then
                subseven18total = subseven18total + 1
            End If
            If InStr(buffer, "Tû{ÓBª¦bQæèhV@ÈLDWDåû¼ı8") Then
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
            If InStr(buffer, "#éwêÊëdìrí9îïGĞ") Then
                subseven19total = subseven19total + 1
            End If
            If InStr(buffer, "W¥ëH5SœÒËcœÛ") Then
                subseven19total = subseven19total + 1
            End If
            If InStr(buffer, "¦Ù²’äÏv…$ÌPj‹»V^‰K*5tjÆĞŸS") Then
                subseven19total = subseven19total + 1
            End If
            If InStr(buffer, "‰°Ùÿ1÷ªÊ|>%ù@") Then
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
            If InStr(buffer, "ùSecèòˆà·WJLvhU") Then
                subseven20total = subseven20total + 1
            End If
            If InStr(buffer, "â†ÿuùQoP:Õ­ê£!") Then
                subseven20total = subseven20total + 1
            End If
            If InStr(buffer, "ÖUD.3I¾øHßpeŒ") Then
                subseven20total = subseven20total + 1
            End If
            If InStr(buffer, "a*•³ËŠcÖ@Ñc4ƒÀß;") Then
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
            If InStr(buffer, "`ÄŞ#(4dÛ'ĞĞ_ŠŠ+Ë~ZÑßìÒZ\Ë") Then
                subseven21total = subseven21total + 1
            End If
            If InStr(buffer, "VÁêj‰†€îøˆSÃ3î‹ú,EøÈÊ„á") Then
                subseven21total = subseven21total + 1
            End If
            If InStr(buffer, "FÑˆ’Ÿr¨‹ŠjL‚É€") Then
                subseven21total = subseven21total + 1
            End If
            If InStr(buffer, "_üœŞK%¼Íò$Wß!øL") Then
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
            If InStr(buffer, "şY#rCX6ÆòCªËLeft") Then
                subseven21goldtotal = subseven21goldtotal + 1
            End If
            If InStr(buffer, "Ê`P+„€dÕ<Ûœ¬yg") Then
                subseven21goldtotal = subseven21goldtotal + 1
            End If
            If InStr(buffer, "Ø<¬öQ)O†œògmÂs8¿") Then
                subseven21goldtotal = subseven21goldtotal + 1
            End If
            If InStr(buffer, "ÆrèkÍ1r€œğJãÄä###") Then
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
            If InStr(buffer, "EØ0ÜÿÿSUìP‰uìQ") Then
                bo12total = bo12total + 1
            End If
            If InStr(buffer, "Ãÿ¼$B Pèæ   ƒÄWÿ¸$B ¸ÿÿÿÿ") Then
                bo12total = bo12total + 1
            End If
            If InStr(buffer, "s‹ÆƒàçÁø˜0") Then
                bo12total = bo12total + 1
            End If
            If InStr(buffer, "[ƒÄÃ‹t$‹l$Vè´zÿÿƒÄUè«zÿÿƒÄWè¢") Then
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
            If InStr(buffer, "ÿÿPUüRWè¥şÿÿƒÄ‹ğV") Then
                bo12encryptedtotal = bo12encryptedtotal + 1
            End If
            If InStr(buffer, "öEüu3Àë$öÃ@töÇu") Then
                bo12encryptedtotal = bo12encryptedtotal + 1
            End If
            If InStr(buffer, "¬ƒıÿÿ—ÿ7¾ÿ—C¾ÿ—Ÿ:¾") Then
                bo12encryptedtotal = bo12encryptedtotal + 1
            End If
            If InStr(buffer, "Ûót,¾t7t£Ûët«Ûït»") Then
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
            If InStr(buffer, "ÿuEÈPEÜPEÀPEÌPEìPEüPEøPEôPÿuğÿuÄèó") Then
                bo2ktotal = bo2ktotal + 1
            End If
            If InStr(buffer, "ƒÄ;Ã~b‹|Ëƒú@}W‰•ŒËBôşÿÿ‰|Ë‹h") Then
                bo2ktotal = bo2ktotal + 1
            End If
            If InStr(buffer, "£Œ¬ÿuøÿuüèıüÿÿh¤") Then
                bo2ktotal = bo2ktotal + 1
            End If
            If InStr(buffer, "ğuPPPPjSÿÕ‹ğSÿ„13í;õu;éı") Then
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

