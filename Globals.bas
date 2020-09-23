Attribute VB_Name = "Globals"
Global cliserv As String
Global win_top As Long
Global win_left As Long
Global remote_port As String
Global local_port As String
Global remote_ip As String
Global my_name As String
Global input_x As Long
Global input_y As Long
Global Const iam_server = "Server"
Global Const iam_client = "Client"
'********************************
'ADDINS DIRECTORIES
Global Const icons_path = "\Icons\"
Global Const sounds_path = "\Sounds\"
'SOUND EVENTS
Global Const noof_events = 9
Global snd_events(noof_events) As sound_event
'EVENT LIST
Global Const event_onload = 0
Global Const event_onunload = 1
Global Const event_onrx = 2
Global Const event_onsend = 3
Global Const event_onerror = 4
Global Const event_onconnect = 5
Global Const event_ondisconnect = 6
Global Const event_ontype = 7
Global Const event_onreturn = 8
'MESSAGE HTML FILE PATH
Global Const messages_path = "\Messages\"
Global Const messages_file = "1.htm"
Global Const start_message = 1
Global Const end_message = 2
'ICON DEFS
Global Const noof_icons = 11
Global msg_icons(noof_icons) As icon_defs
'ICONS LIST
Global Const icon_smile = 0
Global Const icon_sad = 1
Global Const icon_beer = 2
Global Const icon_disgust = 3
Global Const icon_love = 4
Global Const icon_smileo = 5
Global Const icon_smilep = 6
Global Const icon_wink = 7
Global Const icon_unload = 8
Global Const icon_crooked = 9
Global Const icon_coctail = 10
'********************************

Public Sub setup_icons()
    For init_icons = 0 To noof_icons
        Set msg_icons(init_icons) = New icon_defs
        With msg_icons(init_icons)
            Select Case init_icons
                Case Is = icon_smile
                    .icon_filename = App.Path & icons_path & "smile.gif"
                    .icon_recogstr = ":)"
                Case Is = icon_sad
                    .icon_filename = App.Path & icons_path & "sad.gif"
                    .icon_recogstr = ":("
                Case Is = icon_disgust
                    .icon_filename = App.Path & icons_path & "disgust.gif"
                    .icon_recogstr = ":|"
                Case Is = icon_smilep
                    .icon_filename = App.Path & icons_path & "smilep.gif"
                    .icon_recogstr = ":p"
                Case Is = icon_wink
                    .icon_filename = App.Path & icons_path & "wink.gif"
                    .icon_recogstr = ";)"
                Case Is = icon_crooked
                    .icon_filename = App.Path & icons_path & "crooked.gif"
                    .icon_recogstr = ":/"
                Case Is = icon_smileo
                    .icon_filename = App.Path & icons_path & "smileo.gif"
                    .icon_recogstr = ":o"
                Case Is = icon_beer
                    .icon_filename = App.Path & icons_path & "beer.gif"
                    .icon_recogstr = "(B)"
                Case Is = icon_coctail
                    .icon_filename = App.Path & icons_path & "coctail.gif"
                    .icon_recogstr = "(C)"
                Case Is = icon_love
                    .icon_filename = App.Path & icons_path & "love.gif"
                    .icon_recogstr = "(L)"
                Case Is = icon_unlove
                    .icon_filename = App.Path & icons_path & "unlove.gif"
                    .icon_recogstr = "(U)"
            End Select
        End With
    Next init_icons
End Sub

Public Sub setup_events()
    For init_events = 0 To noof_events
        Set snd_events(init_events) = New sound_event
        With snd_events(init_events)
            Select Case init_events
                Case Is = event_onload
                    .snd_name = "sound event onload"
                Case Is = event_onunload
                    .snd_name = "sound event unload"
                Case Is = event_onrx
                    .snd_name = "sound event rx"
                Case Is = event_onsend
                    .snd_name = "sound event send"
                Case Is = event_onerror
                    .snd_name = "sound event error"
                Case Is = event_onconnect
                    .snd_name = "sound event connect"
                Case Is = event_ondisconnect
                    .snd_name = "sound event disconnect"
                Case Is = event_ontype
                    .snd_name = "sound event type"
                Case Is = event_onreturn
                    .snd_name = "sound event return"
            End Select
        End With
    Next init_events
End Sub

Public Sub save_window(window As String, save_top As Long, save_left As Long)
    SaveSetting App.ProductName, "windows", window, "SAVED"
    SaveSetting App.ProductName, "windows", window & " top", save_top
    SaveSetting App.ProductName, "windows", window & " left", save_left
End Sub

Public Sub load_window(window As String)
    'CHECK IF SETTINGS EXIST FIRST
    If GetSetting(App.ProductName, "windows", window) = "SAVED" Then
        'RETRIEVE SETTINGS
        win_top = Val(GetSetting(App.ProductName, "windows", window & " top"))
        win_left = Val(GetSetting(App.ProductName, "windows", window & " left"))
    Else
        win_top = 0
        win_left = 0
    End If
End Sub

Public Function increment_counter(counter As Integer, max As Integer) As Integer
    If counter < max Then
        counter = counter + 1
    Else
        counter = 0
    End If
    increment_counter = counter
End Function
