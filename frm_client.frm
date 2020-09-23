VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frm_client 
   Caption         =   "Web Messenger - Client"
   ClientHeight    =   5355
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   4425
   Icon            =   "frm_client.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5355
   ScaleWidth      =   4425
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frm_out 
      Caption         =   "Send a message"
      Height          =   915
      Left            =   60
      TabIndex        =   3
      Top             =   3840
      Width           =   4275
      Begin VB.TextBox txt_send 
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   4035
      End
   End
   Begin VB.Frame frm_messages 
      Caption         =   "Messages"
      Height          =   3615
      Left            =   60
      TabIndex        =   2
      Top             =   120
      Width           =   4275
      Begin RichTextLib.RichTextBox rtb_messages 
         Height          =   3135
         Left            =   180
         TabIndex        =   5
         Top             =   300
         Width           =   3915
         _ExtentX        =   6906
         _ExtentY        =   5530
         _Version        =   393217
         Enabled         =   -1  'True
         ScrollBars      =   2
         TextRTF         =   $"frm_client.frx":0442
      End
   End
   Begin MSComctlLib.ProgressBar pro_info 
      Height          =   315
      Left            =   2220
      TabIndex        =   1
      Top             =   4980
      Width           =   1875
      _ExtentX        =   3307
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
      Top             =   4860
      Width           =   4425
      _ExtentX        =   7805
      _ExtentY        =   873
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Enabled         =   0   'False
            Object.Width           =   3678
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   3678
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frm_client"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
