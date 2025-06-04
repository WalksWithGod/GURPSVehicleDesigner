Attribute VB_Name = "modSplash"

Option Explicit

'=====================================
'MPJ 02/16/02  ENTIRE MODULE OBSOLETE.
'=====================================
'
''Bitblit constants
'Public Const PicWidth = 155
'Public Const PicHeight = 205
'Public Const PIXEL = 3
'Public Const SRCCOPY = &HCC0020
'
''sndPlaySound constants
'Public Const SND_SYNC = &H0
'Public Const SND_ASYNC = &H1
'Public Const SND_NODEFAULT = &H2
'Public Const SND_LOOP = &H8
'Public Const SND_NOSTOP = &H10
'
''Sound declares
'Public x%
'Public wFlags%
'Public SoundName As String
'
'Public Declare Function BitBlt Lib "gdi32" _
'       (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, _
'       ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, _
'       ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
'
'Public Declare Function sndPlaySound Lib "WINMM.DLL" Alias _
'      "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As _
'      Long) As Long
'
'
'
'
'Sub BlitImage(ByVal i As Integer)
'Dim BmpWidth As Integer
'On Error Resume Next
'
'frmSplash.Picture1.ScaleMode = PIXEL
'frmSplash.Picture2.ScaleMode = PIXEL
'
'BmpWidth = i * frmSplash.Picture1.ScaleWidth
'
'' blit the bitmap object's image onto the picturebox
'Call BitBlt(frmSplash.Picture1.hdc, 0, 0, PicWidth, PicHeight, frmSplash.Picture2.hdc, BmpWidth, 0, SRCCOPY)
'End Sub
'
'Sub ShowStatic()
'Dim BmpWidth As Integer
'Dim i As Integer
'On Error Resume Next
'
'frmSplash.Picture1.ScaleMode = PIXEL
'frmSplash.Picture2.ScaleMode = PIXEL
'
'For i = 0 To 1
'    BmpWidth = i * frmSplash.Picture1.ScaleWidth
'    'play the static sound
'    SoundName$ = App.Path & "\static.wav"
'    wFlags% = SND_ASYNC Or SND_NODEFAULT 'if it cant find / play the soundfile.. it will simply resume w/out sound
'    x% = sndPlaySound(SoundName$, wFlags%)
'    ' blit the bitmap object's image onto the picturebox
'    Call BitBlt(frmSplash.Picture1.hdc, 0, 0, PicWidth, PicHeight, frmSplash.Picture2.hdc, BmpWidth, 0, SRCCOPY)
'    Pause 0.1
'Next
'End Sub
'
'Sub FlashBmps()
'
'Dim upperbound As Integer
'Dim lowerbound As Integer
'Dim RandImage As Integer
'Dim i As Integer
'Dim j As Integer
'Dim Selected(1 To 6) As Integer
'
'upperbound = 8
'lowerbound = 3
'
'Randomize
'
'        RandImage = Int((upperbound - lowerbound + 1) * Rnd + lowerbound)
'        ShowStatic
'        BlitImage RandImage
'
'
'   'this is the old randomization which takes too long
'   'and users didnt want to see a splash screen that sits there
'   'as it cycles thorugh all the bitmaps... and I agree.
'   'For i = 1 To 6
'   '     RandImage = Int((upperbound - lowerbound + 1) * Rnd + lowerbound)
'   '     For j = 1 To 6
'   '         If RandImage = Selected(j) Then
'   '             i = i - 1
'   '         GoTo Skip
'   '         End If
'   '     Next
'   '     Selected(i) = RandImage
'   '     ShowStatic
'   '     BlitImage RandImage
'   '     If i = 1 Then
'   '         Pause 2
'   '     Else
'   '         Pause 1
'   '     End If
'   '     If i = 2 Then Exit Sub
''Skip:
'    'Next
'    'DoEvents
'    'FlashBmps
'
'
'End Sub
'
