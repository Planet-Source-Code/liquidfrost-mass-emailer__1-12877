VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Main 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mass Mailer By Cause"
   ClientHeight    =   5070
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9495
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5070
   ScaleWidth      =   9495
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox DATA 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   2295
      Left            =   2760
      MultiLine       =   -1  'True
      TabIndex        =   14
      Top             =   2640
      Width           =   4935
   End
   Begin VB.OptionButton Option6 
      Caption         =   "SERVER-4"
      Height          =   375
      Left            =   8040
      TabIndex        =   25
      Top             =   4080
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Clear Receivers List"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   5880
      TabIndex        =   23
      Top             =   2040
      Width           =   3495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add Receivers E-mail To List"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   5880
      TabIndex        =   21
      Top             =   1320
      Width           =   3495
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      Height          =   3735
      Left            =   120
      TabIndex        =   20
      Top             =   240
      Width           =   2295
   End
   Begin VB.OptionButton Option5 
      Caption         =   "SERVER-3"
      Height          =   375
      Left            =   8040
      TabIndex        =   19
      Top             =   3600
      Width           =   1215
   End
   Begin VB.OptionButton Option3 
      Caption         =   "CUSTOM"
      Height          =   375
      Left            =   8040
      TabIndex        =   18
      Top             =   4560
      Width           =   1335
   End
   Begin VB.TextBox MAIL_FROM 
      Appearance      =   0  'Flat
      BackColor       =   &H80000014&
      Height          =   285
      Left            =   2760
      TabIndex        =   8
      Text            =   "Blah@hotmail.com"
      Top             =   840
      Width           =   2775
   End
   Begin VB.TextBox RCPT_TO 
      Appearance      =   0  'Flat
      BackColor       =   &H80000014&
      Height          =   285
      Left            =   2760
      TabIndex        =   7
      Top             =   1440
      Width           =   2775
   End
   Begin VB.CommandButton SEND_MAIL 
      Caption         =   "&Send Mail"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   600
      Width           =   3495
   End
   Begin VB.CommandButton MAIL_RESET 
      Appearance      =   0  'Flat
      Caption         =   "Remove Selected Email"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1680
      Width           =   3495
   End
   Begin VB.TextBox STATUS 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   5880
      Locked          =   -1  'True
      TabIndex        =   4
      TabStop         =   0   'False
      Text            =   "Status-Idle"
      Top             =   240
      Width           =   3495
   End
   Begin VB.TextBox FROM 
      Appearance      =   0  'Flat
      BackColor       =   &H80000014&
      Height          =   285
      Left            =   2760
      TabIndex        =   3
      Top             =   240
      Width           =   2775
   End
   Begin VB.TextBox SUBJECT 
      Appearance      =   0  'Flat
      BackColor       =   &H80000014&
      Height          =   285
      Left            =   2760
      TabIndex        =   2
      Top             =   2040
      Width           =   2775
   End
   Begin VB.CommandButton CANCEL_SEND 
      Caption         =   "C&ancel Send"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   960
      Width           =   3495
   End
   Begin VB.TextBox SMTP_HOST 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3960
      TabIndex        =   0
      Top             =   3840
      Visible         =   0   'False
      Width           =   3015
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   4320
      Top             =   3840
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Frame Frame5 
      Caption         =   "Subject"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2640
      TabIndex        =   13
      Top             =   1800
      Width           =   3015
   End
   Begin VB.Frame Frame4 
      Caption         =   "Receivers E-mail to add"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2640
      TabIndex        =   12
      Top             =   1200
      Width           =   3015
   End
   Begin VB.Frame Frame3 
      Caption         =   "Your-Mail Address"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2640
      TabIndex        =   11
      Top             =   600
      Width           =   3015
   End
   Begin VB.Frame Frame2 
      Caption         =   "Your Name"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2640
      TabIndex        =   10
      Top             =   0
      Width           =   3015
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   5760
      TabIndex        =   9
      Top             =   0
      Width           =   3735
   End
   Begin VB.OptionButton Option4 
      Caption         =   "SERVER-2"
      Height          =   375
      Left            =   8040
      TabIndex        =   16
      Top             =   3120
      Width           =   1215
   End
   Begin VB.OptionButton Option1 
      Caption         =   "SERVER-1"
      Height          =   375
      Left            =   8040
      TabIndex        =   17
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Frame Frame7 
      Caption         =   "Receivers E-mail Addresses"
      Height          =   4455
      Left            =   0
      TabIndex        =   22
      Top             =   0
      Width           =   2535
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "0 addresses in the list"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   4080
         Width           =   2295
      End
   End
   Begin VB.Frame Frame8 
      Height          =   2655
      Left            =   7920
      TabIndex        =   24
      Top             =   2400
      Width           =   1575
   End
   Begin VB.Frame Frame9 
      Height          =   615
      Left            =   0
      TabIndex        =   26
      Top             =   4440
      Width           =   2535
      Begin VB.CommandButton Command4 
         Caption         =   "Save List"
         Height          =   255
         Left            =   1320
         TabIndex        =   28
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Load List"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   240
         Width           =   1095
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4440
      Top             =   3000
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   "txt"
      Filter          =   "*.txt"
   End
   Begin VB.Frame Frame6 
      Caption         =   "Message"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   2640
      TabIndex        =   15
      Top             =   2400
      Width           =   5175
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim strName, strFAV As String
Dim curMONEY As Currency
Dim Progress
Dim Green_Light As Boolean
Dim DATAFile As String
Dim Helo_Ok As Boolean
Dim do_cancel As Boolean

Private Sub Command1_Click()
List1.AddItem (RCPT_TO.Text)
If List1.ListCount < 2 Then
Label1.Caption = "1" & " address in the list"
Else
Label1.Caption = List1.ListCount & (" addresses in the list")
End If
End Sub

Private Sub Command2_Click()
List1.Clear
Label1.Caption = List1.ListCount & (" addresses in the list")
End Sub

Private Sub Command3_Click()
CommonDialog1.ShowOpen
If CommonDialog1.FileTitle = "" Then ' so theres no text overflow if the user canceles open file
Exit Sub
Else
Dim MyString As String
    On Error Resume Next
    List1.Clear
    directory$ = CommonDialog1.FileTitle
    Open directory$ For Input As #1
        While Not EOF(1)
                Input #1, MyString$
        DoEvents
        List1.AddItem MyString$
    Wend
    Close #1
    End If
    If List1.ListCount < 2 Then
Label1.Caption = "1" & " address in the list"
Else
Label1.Caption = List1.ListCount & (" addresses in the list")
End If
End Sub

Private Sub Command4_Click()
CommonDialog1.ShowSave
Dim SaveList As Long
    On Error Resume Next
      directory$ = CommonDialog1.FileTitle
        Open directory$ For Output As #1
        For SaveList& = 0 To List1.ListCount - 1
        Print #1, List1.List(SaveList&)
       Next SaveList&
      Close #1
End Sub

Private Sub Form_Load()
 Progress = 0
    do_cancel = False
    Load LOG_FORM


On Error Resume Next
Open (WinSysPath) & "\Setting.txt" For Input As #1

Line Input #1, a
Line Input #1, b
Line Input #1, c
Line Input #1, d
Line Input #1, e
Line Input #1, f

FROM.Text = a
MAIL_FROM.Text = b
RCPT_TO.Text = c
SUBJECT.Text = d
DATA.Text = e
SMTP_HOST.Text = f


Close #1

If SMTP_HOST.Text = "mx-a-rwc.mail.home.com" Then
Option1.Value = 1
End If
If SMTP_HOST.Text = "mailgate.shopping.com" Then
Option4.Value = 1
End If
If SMTP_HOST.Text = "smtp.astreet.com" Then
Option5.Value = 1
End If
If SMTP_HOST.Text = "mx1.pogo.com" Then
Option6.Value = 1
End If


If Option1 + Option4 + Option5 + Option6.Value = 0 Then
Frame1.Caption = "Outgoing Mail (SMTP):  " + SMTP_HOST.Text
If SMTP_HOST.Text = "" Then
Frame1.Caption = "No Outgoing SMTP Mail Server Selected."
End If
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Open (WinSysPath) & "\Setting.txt" For Output As #1


Print #1, FROM
Print #1, MAIL_FROM
Print #1, RCPT_TO
Print #1, SUBJECT
Print #1, DATA
Print #1, SMTP_HOST
Close #1
End

End Sub



Private Sub CANCEL_SEND_Click()
    Winsock1.Close
    Winsock1.RemotePort = 0
    Winsock1.LocalPort = 0
    do_cancel = True
    Progress = 0
    STATUS.ForeColor = &H0&
    STATUS.Text = "Canceled."
    Log "Canceled Sending."
End Sub

Private Sub Check1_Click()


End Sub

Private Sub Check2_Click()

End Sub



Private Sub Form_Terminate()
    Unload Me
    Me.Hide
    Unload LOG_FORM
    LOG_FORM.Hide
    End
End Sub



Private Sub MAIL_RESET_Click()
If List1.ListIndex = -1 Then
MsgBox "You must select a name first", vbInformation
Exit Sub
Else
List1.RemoveItem (List1.ListIndex)
End If
    
End Sub
Private Function CheckText() As Boolean
    Dim bReturn As Boolean
    Dim oCtl As Control
    bReturn = True
    For Each oCtl In Me.Controls
        If TypeOf oCtl Is TextBox Then
            If oCtl.Text = "" And oCtl.Name <> "STATUS" Then
                bReturn = False
                Log "No text in " & oCtl.Name
            End If
        End If
    Next oCtl
    
    CheckText = bReturn
    
End Function

Private Sub REPLY_TO_Change()

End Sub

Private Sub Option1_Click()
SMTP_HOST.Text = "mx-a-rwc.mail.home.com"
Frame1.Caption = "Outgoing Mail (SMTP):  SERVER-1"
End Sub


Private Sub Option3_Click()
 strName = InputBox("                                                                                                               Enter Outgoing Mail (SMTP) Server", "CUSTOM")
    SMTP_HOST.Text = strName
Frame1.Caption = "Outgoing Mail (SMTP):  " & strName
If SMTP_HOST.Text = "" Then
Option1.SetFocus
End If

End Sub

Private Sub Option4_Click()
SMTP_HOST.Text = "mailgate.shopping.com"
Frame1.Caption = "Outgoing Mail (SMTP):  SERVER-2"
End Sub

Private Sub Option5_Click()
SMTP_HOST.Text = "smtp.astreet.com"
Frame1.Caption = "Outgoing Mail (SMTP):  SERVER-3"
End Sub

Private Sub Option6_Click()
SMTP_HOST.Text = "mx1.pogo.com"
Frame1.Caption = "Outgoing Mail (SMTP):  SERVER-4"
End Sub

Private Sub SEND_MAIL_Click()
Winsock1.Close
If SMTP_HOST.Text = "" Then
Option1.Value = 1
MsgBox "You did not select a server so I selected one for you.", vbInformation
End If
If List1.ListCount < 1 Then
MsgBox "Please add one or more receivers email addresses to the list.", vbInformation
End If
For X = 0 To List1.ListCount - 1
List1.ListIndex = X


    Green_Light = False
    Progress = 0
    Helo_Ok = False
    do_cancel = False
    On Error Resume Next
    
   
        
    If InStr(1, MAIL_FROM, "@") = 0 Then
        MsgBox "Your email address must contain an @ character"
        MAIL_FROM.SetFocus
        MAIL_FROM.SelStart = 0
        MAIL_FROM.SelLength = Len(MAIL_FROM)
        Log "Error, no @ in Your email address, stoping send."
        Exit Sub
    End If
    
    Winsock1.Close
    Winsock1.Connect SMTP_HOST, "25"
    
    Do While Winsock1.State <> sckConnected
        DoEvents
        STATUS.Text = "Connecting.. Please wait."
        If do_cancel = True Then STATUS = "Canceled...": do_cancel = False: Winsock1.SendData "QUIT" & Chr$(13) & Chr$(10): Exit Sub
    Loop
    STATUS.ForeColor = &H0&
    
    STATUS.Text = "Connected.."
    Log "Connected to " & SMTP_HOST & "."
    
    Do While Green_Light = False
        DoEvents
        STATUS.Text = "Waiting for reply..."
        If do_cancel = True Then STATUS = "Canceled...": do_cancel = False: Winsock1.SendData "QUIT" & Chr$(13) & Chr$(10): Exit Sub
    Loop
    
    Winsock1.SendData "HELO " & Mid(MAIL_FROM, InStr(1, MAIL_FROM, "@") + 1, Len(MAIL_FROM)) & Chr$(13) & Chr$(10)
    Log "HELO " & Mid(MAIL_FROM, InStr(1, MAIL_FROM, "@") + 1, Len(MAIL_FROM))
    
    Do While Helo_Ok = False
        DoEvents
        STATUS.Text = "Waiting for reply..."
        If do_cancel = True Then STATUS = "Canceled...": do_cancel = False: Winsock1.SendData "QUIT" & Chr$(13) & Chr$(10): Exit Sub
    Loop
    
    Winsock1.SendData "MAIL FROM: <" & MAIL_FROM & ">" & Chr$(13) & Chr$(10)
    Log "MAIL FROM: " & MAIL_FROM
    
    Do While Progress <> 1
        DoEvents
        STATUS.Text = "Sending data."
        If do_cancel = True Then STATUS = "Canceled...": do_cancel = False: Winsock1.SendData "QUIT" & Chr$(13) & Chr$(10): Exit Sub
    Loop
    
    
    Winsock1.SendData "RCPT TO: <" & List1.Text & ">" & Chr$(13) & Chr$(10)
    Log "RCPT TO: " & List1.Text
    
    'GREEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEN
    Do While Progress <> 2
        DoEvents
        STATUS.Text = "Sending data.."
        If do_cancel = True Then STATUS = "Canceled...": do_cancel = False: Winsock1.SendData "QUIT" & Chr$(13) & Chr$(10): Exit Sub
    Loop
    
    Winsock1.SendData "DATA" & Chr$(13) & Chr$(10)
    Log "DATA"
    
    Do While Progress <> 3
        DoEvents
        STATUS.Text = "Setting up body transfer..."
        If do_cancel = True Then STATUS = "Canceled...": do_cancel = False: Winsock1.SendData "QUIT" & Chr$(13) & Chr$(10): Exit Sub
    Loop
    
    Winsock1.SendData GenerateMessageID(Mid(MAIL_FROM, InStr(1, MAIL_FROM, "@") + 1, Len(MAIL_FROM))) & Chr$(13) & Chr$(10)
    Winsock1.SendData "DATE: " & Format(Now, "h:mm:ss") & Chr$(13) & Chr$(10)
    Winsock1.SendData "FROM: " & FROM & " <" & MAIL_FROM & ">" & Chr$(13) & Chr$(10)

    Winsock1.SendData "TO: <" & List1.Text & ">" & Chr$(13) & Chr$(10)
    ' GREEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEEN
    Winsock1.SendData "SUBJECT: " & SUBJECT & Chr$(13) & Chr$(10)
    Winsock1.SendData "MIME-Version: 1.0" & Chr$(13) & Chr$(10)
    Winsock1.SendData "Content-Type: text/plain; charset=us-ascii" & Chr$(13) & Chr$(10)
    Winsock1.SendData Chr$(13) & Chr$(10)
    
    Winsock1.SendData DATA & Chr$(13) & Chr$(10)
    Log DATA
    Winsock1.SendData Chr$(13) & Chr$(10) & "." & Chr$(13) & Chr$(10)
    
    Log Chr$(13) & Chr$(10) & "."
    
    Do While Progress <> 4
        DoEvents
        STATUS.Text = "Sending data..."
        If do_cancel = True Then STATUS = "Canceled...": do_cancel = False: Winsock1.SendData "QUIT" & Chr$(13) & Chr$(10): Exit Sub
    Loop
    
    Winsock1.SendData "QUIT" & Chr$(13) & Chr$(10)
    STATUS.ForeColor = &HC000&
    STATUS.Text = "Mail Sent to " & List1.Text
    Winsock1.Close
    Winsock1.RemotePort = 0
    Winsock1.LocalPort = 0
    do_cancel = False

    If List1.ListIndex = List1.ListCount - 1 Then
    PlaySound 0, "C:\windows\MEDIA\Tada.wav"
    STATUS.Text = "E-Mail Sent to Every address in the list."
    End If
    
Next X
End Sub

Private Sub SMTP_HOST_Change()
If SMTP_HOST.Text = "mx-a-rwc.mail.home.com" Then
Option1.Value = 1
End If
If SMTP_HOST.Text = "mailgate.shopping.com" Then
Option4.Value = 1
End If
If SMTP_HOST.Text = "smtp.astreet.com" Then
Option5.Value = 1
End If
If SMTP_HOST.Text = "mx1.pogo.com" Then
Option6.Value = 1
End If
End Sub

Private Sub STATUS_Change()
If STATUS.Text = "Canceled." Then
PlaySound 0, "C:\windows\MEDIA\Chord.wav"
End If


End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
    Dim Reply
On Error GoTo retry:
retry:
    Winsock1.GetData DATAFile
On Error Resume Next
    Reply = Mid(DATAFile, 1, 3)
    
    Log DATAFile
  
    Select Case Reply
        Case 250, 354
            Progress = Progress + 1
            Helo_Ok = True
        Case 220
            Green_Light = True
        Case 503
            Log "Error, helo command failed, or was never sent."
        Case 451
            MsgBox "The site you are attempting to send to requires that the hostname (blah@HOSTNAME.com) actually exists." & vbCrLf & vbCrLf & "This means that you cannot use " & Me.RCPT_TO & " as the fake from address."
            Log "The site you are attempting to send to requires that the hostname (blah@HOSTNAME.com) actually exists." & vbCrLf & vbCrLf & "This means that you cannot use " & Me.RCPT_TO & " as the fake from address."
            CANCEL_SEND_Click
    End Select
        
End Sub

Private Sub Log(ByVal sText As String)
    
    With LOG_FORM.LOG_TEXT
        .SelStart = Len(.Text)
        .SelText = sText & Chr$(13) & Chr$(10)
        .SelLength = 0
    End With

End Sub

Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    Log Number & " / " & Description
End Sub

Private Function GenerateMessageID(ByVal sHost As String) As String
    Dim idnum As Double
    Dim sMessageID As String
    
    sMessageID = "Message-ID: "
    
  
    Randomize Int(CDbl((Now))) + Timer
    
    idnum = GetRandom(9999999999999#, 99999999999999#)
    
    sMessageID = sMessageID & CStr(idnum)
    
    idnum = GetRandom(9999, 99999)
    
    sMessageID = sMessageID & "." & CStr(idnum) & ".qmail@" & sHost
    
    GenerateMessageID = sMessageID
    
End Function
Private Function GetRandom(ByVal dFrom As Double, ByVal dTo As Double) As Double

    Dim X As Double
    Randomize
    X = dTo - dFrom
    GetRandom = Int((X * Rnd) + 1) + dFrom
    
End Function
