VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form frm_messenger 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Web Messenger - <TYPE>"
   ClientHeight    =   8805
   ClientLeft      =   150
   ClientTop       =   735
   ClientWidth     =   7260
   Icon            =   "frm_messenger.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8805
   ScaleWidth      =   7260
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frm_messages 
      Caption         =   "Messages"
      Height          =   7275
      Left            =   60
      TabIndex        =   4
      Top             =   60
      Width           =   7095
      Begin SHDocVwCtl.WebBrowser web_messages 
         Height          =   6795
         Left            =   180
         TabIndex        =   5
         Top             =   300
         Width           =   6735
         ExtentX         =   11880
         ExtentY         =   11986
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   1
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   ""
      End
   End
   Begin VB.Frame frm_out 
      Caption         =   "Send a message"
      Height          =   795
      Left            =   60
      TabIndex        =   2
      Top             =   7440
      Width           =   7095
      Begin VB.TextBox txt_send 
         Height          =   375
         Left            =   180
         TabIndex        =   3
         ToolTipText     =   "Type the message you want to send in here"
         Top             =   240
         Width           =   6855
      End
   End
   Begin MSComctlLib.ProgressBar pro_info 
      Height          =   315
      Left            =   3720
      TabIndex        =   1
      Top             =   8400
      Width           =   3435
      _ExtentX        =   6059
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin MSComctlLib.StatusBar sta_info 
      Align           =   2  'Align Bottom
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   8310
      Width           =   7260
      _ExtentX        =   12806
      _ExtentY        =   873
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   6350
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Bevel           =   2
            Object.Width           =   6350
         EndProperty
      EndProperty
   End
   Begin MSWinsockLib.Winsock win_tcpip 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer tim_timout 
      Left            =   480
      Top             =   0
   End
   Begin VB.Timer tim_refinfo 
      Left            =   960
      Top             =   0
   End
   Begin MSComDlg.CommonDialog cdlg 
      Left            =   1440
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu men_file 
      Caption         =   "File"
      Begin VB.Menu men_setup 
         Caption         =   "Setup"
      End
      Begin VB.Menu Seperator1 
         Caption         =   "-"
      End
      Begin VB.Menu men_exit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu men_messages 
      Caption         =   "Messages"
      Begin VB.Menu men_includeuser 
         Caption         =   "Include User Name"
      End
      Begin VB.Menu men_changeuser 
         Caption         =   "Change User Name"
      End
      Begin VB.Menu Seperator2 
         Caption         =   "-"
      End
      Begin VB.Menu men_savemess 
         Caption         =   "Save Messages"
      End
      Begin VB.Menu Seperator3 
         Caption         =   "-"
      End
      Begin VB.Menu men_clear 
         Caption         =   "Clear Messages"
      End
   End
End
Attribute VB_Name = "frm_messenger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'********************************
Const timer_interval = 50
Const timer_max = 800
Const typing = "<T>"
Dim current_sec As Integer
'********************************
'ICON CONSTANTS
Const ico_smile = ":)"
Const ico_sad = ":("
Const ico_disgust = ":|"
Const ico_beer = "(B)"
'********************************

Private Sub men_changeuser_Click()
    my_name = InputBox("Please enter your name", "Web Messenger - User", , input_x, input_y)
    If my_name = "" Then my_name = "User"
    Me.Caption = "Web Messenger - " & cliserv & " : " & my_name
End Sub

Private Sub men_clear_Click()
    setup_message start_message
    web_messages.Refresh
End Sub

Private Sub Form_Load()
    current_sec = 0
    tim_timout.Interval = timer_interval
    pro_info.max = timer_max
    tim_refinfo.Interval = 3000
    begin_session
End Sub

Private Sub setup_message(pos As Integer)
        Select Case pos
            Case Is = start_message
                Open App.Path + messages_path + messages_file For Output As #1
                    Print #1, "<html>"
                    Print #1, "<head>"
                    Print #1, "<title>Web Messenger Message</title>"
                    Print #1, "<meta http-equiv=""Content-Type"" content=""text/html; charset=iso-8859-1"">"
                    Print #1, "</head>"
                    Print #1, "<body bgcolor=""#FFFFFF"">"
                Close #1
            Case Is = end_message
                Open App.Path + messages_path + messages_file For Append As #1
                    Print #1, "</p>"
                    Print #1, "</body>"
                    Print #1, "</html>"
                Close #1
        End Select
End Sub

Private Sub begin_session()
    '*****************
    setup_message start_message
    web_messages.Navigate (App.Path + messages_path + messages_file)
    '*****************
    toggle_controls
    load_settings
    men_includeuser.Checked = True
    frm_messenger.Caption = "Web Messenger - " & cliserv & " : " & my_name
    Select Case cliserv
        Case Is = iam_server
            With win_tcpip
                .LocalPort = Val(local_port)
                .Listen
            End With
            sta_info.Panels(1).Text = "Waiting for Client..."
        Case Is = iam_client
            With win_tcpip
                .RemoteHost = remote_ip
                .RemotePort = Val(remote_port)
            End With
            sta_info.Panels(1).Text = "Looking for Server..."
    End Select
    tim_timout.Enabled = True
End Sub
    
Private Sub end_session()
    setup_message end_message
    save_settings
    frm_main.Show
    frm_main.sta_info.Panels(1).Text = "Welcome to Web Messenger " & my_name
End Sub

Private Sub Form_Unload(Cancel As Integer)
    end_session
End Sub

Private Sub load_settings()
    With frm_messenger
        load_window (.Caption)
        If win_top <> 0 Then .Top = win_top
        If win_left <> 0 Then .Left = win_left
    End With
End Sub

Private Sub save_settings()
    With frm_messenger
        save_window .Caption, .Top, .Left
    End With
End Sub

Private Sub men_exit_Click()
    Unload Me
End Sub

Private Sub men_includeuser_Click()
    men_includeuser.Checked = Not men_includeuser.Checked
End Sub

Private Sub men_setup_Click()
    load frm_setup
    frm_setup.Show
End Sub

Private Sub tim_refinfo_Timer()
    sta_info.Panels(1).Text = ""
End Sub
Private Sub tim_timout_Timer()
    current_sec = increment_counter(current_sec, timer_max)
    If current_sec <> 0 Then
        pro_info = current_sec
        If cliserv = iam_client Then
            If win_tcpip.State <> sckClosed Then win_tcpip.Close
            win_tcpip.Connect
        End If
    Else
        snd_events(event_onerror).start
        MsgBox "Web-Messenger was unable to connect", vbOKOnly, "Connection Timout"
        frm_main.Show
        Unload Me
    End If
End Sub

Private Sub txt_send_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case Is = vbKeyReturn
            If txt_send <> "" Then
                snd_events(event_onreturn).start
                If men_includeuser.Checked = True Then txt_send = my_name & " says: " & txt_send
                win_tcpip.SendData txt_send
                write_message txt_send, False
                txt_send = ""
            End If
        Case Else
            snd_events(event_ontype).start
            win_tcpip.SendData typing
    End Select
End Sub

Private Sub win_tcpip_Close()
    snd_events(event_ondisconnect).start
    MsgBox "The connection was lost", vbOKOnly, "Connection Error"
    end_session
    Unload Me
End Sub

Private Sub win_tcpip_Connect()
    snd_events(event_onconnect).start
    tim_timout.Enabled = False
    pro_info = 0
    sta_info.Panels(1).Text = "Connected"
    toggle_controls
End Sub

Private Sub win_tcpip_DataArrival(ByVal bytesTotal As Long)
    Dim statustype As Integer
    Dim strdata As String
    Dim strData_passed As String
    win_tcpip.GetData strdata
    statustype = InStr(1, strdata, typing, vbTextCompare)
    If statustype <> 0 Then sta_info.Panels(1).Text = cliserv & " is typing..."
    strData_passed = Replace(strdata, typing, "")
    'PRINT MESSAGE
    If strData_passed <> "" Then write_message strData_passed, True
End Sub

Private Sub write_message(message As String, recieved As Boolean)
    'REPLACE ICONS
    For replace_recog = 0 To noof_icons - 1
        With msg_icons(replace_recog)
            message = Replace(message, .icon_recogstr, "<img src=""" & .icon_filename & """ width=12 height=12>")
        End With
    Next replace_recog
    'CHANGE COLOUR AND PLAY SOUND
    Dim text_colour As String
    If recieved = True Then
        text_colour = Str(vbRed)
        snd_events(event_onrx).start
    Else
        text_colour = Str(vbBlue)
        snd_events(event_onsend).start
    End If
    Static msgno As Long
    msgno = msgno + 1
    Open App.Path + messages_path + messages_file For Append As #1
        Print #1, "<a name=""BM" & msgno & """></a><p><font color=""" & text_colour & """><b>" & message & "</b></font></p>"
    Close #1
    web_messages.Refresh
    web_messages.Navigate App.Path + messages_path + messages_file & "#BM" & msgno
End Sub

Private Sub win_tcpip_ConnectionRequest(ByVal requestID As Long)
    If cliserv = iam_server Then
        If win_tcpip.State <> sckClosed Then win_tcpip.Close
        win_tcpip.Accept requestID
        tim_timout.Enabled = False
        pro_info = 0
        sta_info.Panels(1).Text = "Connected"
        toggle_controls
    End If
End Sub

Private Sub toggle_controls()
    men_messages.Enabled = Not men_messages.Enabled
    frm_messages.Enabled = Not frm_messages.Enabled
    frm_out.Enabled = Not frm_out.Enabled
    tim_refinfo.Enabled = Not tim_refinfo.Enabled
End Sub
