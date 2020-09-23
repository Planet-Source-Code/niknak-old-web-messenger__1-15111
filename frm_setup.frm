VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frm_setup 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Sound Events"
   ClientHeight    =   5490
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3570
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5490
   ScaleWidth      =   3570
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmd_change 
      Caption         =   "Change"
      Height          =   435
      Left            =   60
      TabIndex        =   3
      Top             =   4980
      Width           =   1035
   End
   Begin VB.CommandButton cmd_cancel 
      Caption         =   "Cancel"
      Height          =   435
      Left            =   1260
      TabIndex        =   2
      Top             =   4980
      Width           =   1035
   End
   Begin VB.CommandButton cmd_ok 
      Caption         =   "Ok"
      Height          =   435
      Left            =   2460
      TabIndex        =   1
      Top             =   4980
      Width           =   1035
   End
   Begin MSComctlLib.TreeView trv_sndevents 
      Height          =   4815
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   3435
      _ExtentX        =   6059
      _ExtentY        =   8493
      _Version        =   393217
      Style           =   7
      Checkboxes      =   -1  'True
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComDlg.CommonDialog cdlg 
      Left            =   60
      Top             =   60
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frm_setup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_cancel_Click()
    Unload Me
End Sub

Private Sub cmd_change_Click()
    With trv_sndevents
        If .Nodes.Count > 0 Then
            If .SelectedItem.index <> 0 Then
                change_sound .SelectedItem.index
            End If
        End If
    End With
End Sub

Private Sub cmd_ok_Click()
    For saveevents = 0 To noof_events - 1
        snd_events(saveevents).snd_enabled = trv_sndevents.Nodes.Item(saveevents + 1).Checked
        snd_events(saveevents).save App.ProductName
    Next saveevents
    Unload Me
End Sub

Private Sub Form_Load()
    load_settings
    refresh_events
End Sub

Private Sub load_settings()
    With frm_setup
        load_window (.Caption)
        If win_top <> 0 Then .Top = win_top
        If win_left <> 0 Then .Left = win_left
    End With
End Sub

Private Sub save_settings()
    With frm_setup
        save_window .Caption, .Top, .Left
    End With
End Sub

Private Sub refresh_events()
    trv_sndevents.Nodes.Clear
    For addnodes = 1 To noof_events
        With trv_sndevents
            .Nodes.Add , , snd_events(addnodes - 1).snd_name, snd_events(addnodes - 1).snd_name
            .Nodes.Item(addnodes).Checked = snd_events(addnodes - 1).snd_enabled
        End With
    Next addnodes
End Sub

Private Sub change_sound(index As Integer)
    Dim filename As String
        With cdlg
            .CancelError = True
            On Error GoTo ErrHandler
            .Flags = cdlOFNHideReadOnly
            .Filter = "Wave File (*.wav)|*.wav"
            .ShowSave
            snd_events(index - 1).filename .filename
        End With
    Exit Sub
ErrHandler:
    Exit Sub
End Sub

Private Sub Form_Unload(Cancel As Integer)
    save_settings
End Sub

Private Sub trv_sndevents_NodeCheck(ByVal Node As MSComctlLib.Node)
    With trv_sndevents
        If .SelectedItem.index > 0 Then snd_events(.SelectedItem.index).snd_enabled = .SelectedItem.Checked
    End With
End Sub
