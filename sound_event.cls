VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "sound_event"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'GENERAL DECLARATIONS
Private Declare Function sndPlaySound Lib "winmm" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Option Explicit
Public snd_name As String
Public snd_enabled As Boolean
Private snd_filename As String
Private snd_lastchanged As Date

'SETS OR RETURNS THE EVENT SOUND FILE NAME
Public Function filename(Optional i_filename As String) As String
    If i_filename <> "" Then
        Dim fs
        Set fs = CreateObject("Scripting.FileSystemObject")
        If fs.fileexists(i_filename) Then
            snd_filename = i_filename
            snd_lastchanged = Now
            filename = "True"
        Else
            filename = "False"
        End If
    Else
        filename = snd_filename
    End If
End Function

'RETURNS THE DATE THE FILENAME WAS LAST CHANGED
Public Property Get lastchanged() As Date
    lastchanged = snd_lastchanged
End Property

'ACTIVATES THE EVENT
Public Sub start()
    If snd_enabled = True Then sndPlaySound snd_filename, 1
End Sub

'SAVES THE EVENT MEMBERS TO THE SYSTEM REGISTRY
Public Sub save(appname As String)
    SaveSetting appname, snd_name, "snd_filename", snd_filename
    SaveSetting appname, snd_name, "snd_lastchanged", CStr(snd_lastchanged)
    SaveSetting appname, snd_name, "snd_enabled", snd_enabled
End Sub

'LOADS THE EVENT MEMBERS FROM THE SYSTEM REGISTRY
Public Sub load(appname As String)
    If GetSetting(appname, snd_name, "snd_filename") <> "" Then
        snd_filename = GetSetting(appname, snd_name, "snd_filename")
        snd_lastchanged = CDate(GetSetting(appname, snd_name, "snd_lastchanged"))
        snd_enabled = CBool(GetSetting(appname, snd_name, "snd_enabled"))
    End If
End Sub

'CONSTRUCTOR
Private Sub Class_Initialize()
    snd_lastchanged = Now
End Sub
