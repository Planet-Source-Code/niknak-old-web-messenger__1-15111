VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frm_main 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Web Messenger - Main menu"
   ClientHeight    =   4425
   ClientLeft      =   150
   ClientTop       =   735
   ClientWidth     =   7815
   Icon            =   "frm_main.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   4425
   ScaleWidth      =   7815
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frm_help 
      Caption         =   "Help me!"
      Height          =   3735
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   2055
      Begin VB.Image Image1 
         Height          =   480
         Left            =   720
         Picture         =   "frm_main.frx":0442
         Top             =   360
         Width           =   480
      End
      Begin VB.Label lbl_helpme 
         Alignment       =   2  'Center
         Caption         =   $"frm_main.frx":0884
         Height          =   2415
         Left            =   120
         TabIndex        =   3
         Top             =   1080
         Width           =   1815
      End
   End
   Begin VB.Frame frm_client 
      Caption         =   "I want to be a Client"
      Height          =   1815
      Left            =   2280
      TabIndex        =   1
      Top             =   2040
      Width           =   5415
      Begin VB.CommandButton cmd_begin 
         Caption         =   "Begin"
         Height          =   855
         Index           =   1
         Left            =   4560
         Picture         =   "frm_main.frx":0993
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox txt_serport 
         Height          =   375
         Left            =   1920
         TabIndex        =   9
         Text            =   "Server's TCP/IP Port"
         Top             =   1320
         Width           =   2535
      End
      Begin VB.TextBox txt_serip 
         Height          =   375
         Left            =   1920
         TabIndex        =   8
         Text            =   "Server's IP Address"
         Top             =   840
         Width           =   2535
      End
      Begin VB.Label lbl_clienthelp 
         Caption         =   "You must obtain the details below from the Server before you begin.  These are essential to connect successfully."
         Height          =   495
         Left            =   840
         TabIndex        =   15
         Top             =   240
         Width           =   4455
      End
      Begin VB.Label lbl_serport 
         Caption         =   "Server's TCP/IP Port"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label lbl_serip 
         Caption         =   "Server's IP Address"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   960
         Width           =   1575
      End
      Begin VB.Image Image2 
         Height          =   480
         Left            =   120
         Picture         =   "frm_main.frx":0DD5
         Top             =   240
         Width           =   480
      End
   End
   Begin VB.Frame frm_server 
      Caption         =   "I want to be the Server"
      Height          =   1815
      Left            =   2280
      TabIndex        =   0
      Top             =   120
      Width           =   5415
      Begin VB.CommandButton cmd_begin 
         Caption         =   "Begin"
         Height          =   855
         Index           =   0
         Left            =   4560
         Picture         =   "frm_main.frx":1217
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox txt_myport 
         Height          =   375
         Left            =   1920
         TabIndex        =   5
         Text            =   "My TCP/IP Port"
         Top             =   1320
         Width           =   2535
      End
      Begin VB.TextBox txt_myip 
         Height          =   375
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   4
         Text            =   "My IP Address"
         Top             =   840
         Width           =   2535
      End
      Begin VB.Label lbl_serverhelp 
         Caption         =   "Please make sure that you inform the Client of your IP Address and TCP/IP Port before you choose to begin."
         Height          =   495
         Left            =   840
         TabIndex        =   14
         Top             =   240
         Width           =   4455
      End
      Begin VB.Label lbl_myport 
         Caption         =   "My TCP/IP Port"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label lbl_myip 
         Caption         =   "My IP Address"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   960
         Width           =   1575
      End
      Begin VB.Image Image3 
         Height          =   480
         Left            =   120
         Picture         =   "frm_main.frx":1659
         Top             =   240
         Width           =   480
      End
   End
   Begin MSWinsockLib.Winsock win_details 
      Left            =   60
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar sta_info 
      Align           =   2  'Align Bottom
      Height          =   495
      Left            =   0
      TabIndex        =   16
      Top             =   3930
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   873
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   6853
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Bevel           =   2
            Object.Width           =   6853
         EndProperty
      EndProperty
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
      Begin VB.Menu men_changeuser 
         Caption         =   "Change User Name"
      End
   End
End
Attribute VB_Name = "frm_main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_begin_Click(index As Integer)
    local_port = txt_myport
    remote_ip = txt_serip
    remote_port = txt_serport
    If index = 0 Then
        cliserv = iam_server
    Else
        cliserv = iam_client
    End If
    frm_messenger.Show
    Me.Hide
End Sub

Private Sub Form_Load()
    my_name = InputBox("Please enter your name", "Web Messenger - User", , input_x, input_y)
    If my_name = "" Then my_name = "User"
    sta_info.Panels(1).Text = "Welcome to Web Messenger " & my_name
    load_settings
    load_options
End Sub

Private Sub Form_Unload(Cancel As Integer)
    save_settings
    save_options
    snd_events(event_onunload).start
End Sub

Private Sub load_settings()
    With frm_main
        load_window (.Caption)
        If win_top <> 0 Then .Top = win_top
        If win_left <> 0 Then .Left = win_left
    End With
End Sub

Private Sub save_settings()
    With frm_main
        save_window .Caption, .Top, .Left
    End With
End Sub

Private Sub save_options()
    SaveSetting App.ProductName, frm_main.Caption, "Options", "SAVED"
    SaveSetting App.ProductName, frm_main.Caption, "My IP Address", txt_myip.Text
    SaveSetting App.ProductName, frm_main.Caption, "My TCP/IP Port", txt_myport.Text
    SaveSetting App.ProductName, frm_main.Caption, "Server's IP Address", txt_serip.Text
    SaveSetting App.ProductName, frm_main.Caption, "Server's TCP/IP Port", txt_serport.Text
End Sub

Private Sub load_options()
    If GetSetting(App.ProductName, frm_main.Caption, "Options") = "SAVED" Then
        txt_myport = GetSetting(App.ProductName, frm_main.Caption, "My TCP/IP Port")
        txt_serip = GetSetting(App.ProductName, frm_main.Caption, "Server's IP Address")
        txt_serport = GetSetting(App.ProductName, frm_main.Caption, "Server's TCP/IP Port")
    Else
        txt_myport = "My TCP/IP Port"
        txt_serip = "Server's IP Address"
        txt_serport = "Server's TCP/IP Port"
    End If
    txt_myip = win_details.LocalIP
End Sub

Private Sub men_changeuser_Click()
    my_name = InputBox("Please enter your name", "Web Messenger - User", , input_x, input_y)
    If my_name = "" Then my_name = "User"
    sta_info.Panels(1).Text = "Welcome to Web Messenger " & my_name
End Sub

Private Sub men_exit_Click()
    Unload Me
End Sub

Private Sub men_setup_Click()
    load frm_setup
    frm_setup.Show
End Sub
