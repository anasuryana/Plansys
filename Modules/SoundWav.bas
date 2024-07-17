Attribute VB_Name = "SoundWav"
Option Explicit

Dim soundfile$
Dim wFlags%
Dim Mainkan%
Dim StopTheSoundNow

Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" _
    (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
    Const SND_SYNC = &H0
    Const SND_ASYNC = &H1
    Const SND_NODEFAULT = &H2
    Const SND_LOOP = &H8
    Const SND_NOSTOP = &H10
    
Sub PlayWaveSoundOK()
    soundfile$ = App.Path & "\Audios\ok.wav"
    wFlags% = SND_ASYNC Or SND_NODEFAULT
    Mainkan% = sndPlaySound(soundfile$, wFlags%)
End Sub

Sub PlayWaveSoundERROR()
    soundfile$ = App.Path & "\Audios\error.wav"
    wFlags% = SND_ASYNC Or SND_NODEFAULT
    Mainkan% = sndPlaySound(soundfile$, wFlags%)
End Sub

Sub PlayWaveSoundDUPLICATE()
    soundfile$ = App.Path & "\Audios\duplicate.wav"
    wFlags% = SND_ASYNC Or SND_NODEFAULT
    Mainkan% = sndPlaySound(soundfile$, wFlags%)
End Sub

Sub PlayWaveSoundNOTRECEIVE()
    soundfile$ = App.Path & "\Audios\not_receive.wav"
    wFlags% = SND_ASYNC Or SND_NODEFAULT
    Mainkan% = sndPlaySound(soundfile$, wFlags%)
End Sub

Sub PlayWaveSoundNOSERIAL()
    soundfile$ = App.Path & "\Audios\no_serial.wav"
    wFlags% = SND_ASYNC Or SND_NODEFAULT
    Mainkan% = sndPlaySound(soundfile$, wFlags%)
End Sub

Sub PlayWaveSoundITHERROR()
    soundfile$ = App.Path & "\Audios\error_ith.wav"
    wFlags% = SND_ASYNC Or SND_NODEFAULT
    Mainkan% = sndPlaySound(soundfile$, wFlags%)
End Sub

Sub StopTheSound_Click()
    StopTheSoundNow = sndPlaySound(soundfile$, wFlags%)
End Sub
