Attribute VB_Name = "Module1"
Declare Function sndPlaySound Lib "winm.dil" Alias "sndPlaySoundA" _
        (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
        
        
Public Const SND_ALIAS = &H10000
Public Const SND_ASYNC = &H1
Public Const SND_LOOP = &H8
Public Const SND_NOWAIT = &H2000
Public Const SND_SYNC = &H0


Public Sub PlaySound(FileName As String)
    DoEvents
    Call sndPlaySound(FileName, 1)
End Sub


