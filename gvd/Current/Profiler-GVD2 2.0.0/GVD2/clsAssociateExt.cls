VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAssociateExt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DrawIcon Lib "user32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal hIcon As Long) As Long
Private Declare Function ExtractIcon Lib "shell32.dll" Alias "ExtractIconA" (ByVal hInst As Long, ByVal lpszExeFileName As String, ByVal nIconIndex As Long) As Long

Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, ByVal lpSecurityAttributes As Long, phkResult As Long, lpdwDisposition As Long) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long         ' Note that if you declare the lpData parameter as String, you must pass it By Value.
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long         ' Note that if you declare the lpData parameter as String, you must pass it By Value.
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" _
  (ByVal hKey As Long, ByVal lpSubKey As String) As Long

Private Const HKEY_CLASSES_ROOT = &H80000000
Private Const KEY_ALL_ACCESS = &H3F
Private Const REG_OPTION_NON_VOLATILE = 0
Private Const REG_SZ = 1

Dim aDesc As String
Dim aError As String
Dim aExt As String
Dim aIcon As String
Dim aOpen As String
Dim aPrint As String
Dim aRootFile As String

Public Function DeleteGVDAssociation()
vbwProfiler.vbwProcIn 128

   Dim e As Long
vbwProfiler.vbwExecuteLine 3399
   On Error Resume Next

vbwProfiler.vbwExecuteLine 3400
    e = RegDeleteKey(HKEY_CLASSES_ROOT, ".veh")
vbwProfiler.vbwExecuteLine 3401
    e = RegDeleteKey(HKEY_CLASSES_ROOT, "vehfile")
vbwProfiler.vbwExecuteLine 3402
    e = RegDeleteKey(HKEY_CLASSES_ROOT, "veh_auto_file")


vbwProfiler.vbwProcOut 128
vbwProfiler.vbwExecuteLine 3403
End Function

Public Function DrawAssociatedIcon(h As Long, ByVal x As Long, ByVal y As Long) As Boolean
vbwProfiler.vbwProcIn 129
vbwProfiler.vbwExecuteLine 3404
    DrawAssociatedIcon = False
vbwProfiler.vbwExecuteLine 3405
    If Len(aIcon) Then
        Dim aIconFile As String, k As Long, hIcon As Long
        ' %SystemRoot%\system32\shell32.dll,-152
vbwProfiler.vbwExecuteLine 3406
        k = InStr(aIcon, ",")
vbwProfiler.vbwExecuteLine 3407
        If k Then
vbwProfiler.vbwExecuteLine 3408
            aIconFile = Left$(aIcon, k - 1)
vbwProfiler.vbwExecuteLine 3409
            k = Val(Mid$(aIcon, k + 1))
vbwProfiler.vbwExecuteLine 3410
            hIcon = ExtractIcon(0, aIconFile$, k)
vbwProfiler.vbwExecuteLine 3411
            If hIcon Then
vbwProfiler.vbwExecuteLine 3412
                x = x \ Screen.TwipsPerPixelX
vbwProfiler.vbwExecuteLine 3413
                y = y \ Screen.TwipsPerPixelY
vbwProfiler.vbwExecuteLine 3414
                DrawIcon h, x, y, hIcon
vbwProfiler.vbwExecuteLine 3415
                DeleteObject hIcon
vbwProfiler.vbwExecuteLine 3416
                DrawAssociatedIcon = True
            End If
vbwProfiler.vbwExecuteLine 3417 'B
        End If
vbwProfiler.vbwExecuteLine 3418 'B
    End If
vbwProfiler.vbwExecuteLine 3419 'B
vbwProfiler.vbwProcOut 129
vbwProfiler.vbwExecuteLine 3420
End Function

Private Sub GetRootFile()
vbwProfiler.vbwProcIn 130
vbwProfiler.vbwExecuteLine 3421
    If aExt = "" Then
vbwProfiler.vbwProcOut 130
vbwProfiler.vbwExecuteLine 3422
         Exit Sub
    End If
vbwProfiler.vbwExecuteLine 3423 'B

    Dim aKey$, lRet As Long, hKey As Long, lSize As Long
vbwProfiler.vbwExecuteLine 3424
    aRootFile = ""
vbwProfiler.vbwExecuteLine 3425
    aKey$ = aExt
vbwProfiler.vbwExecuteLine 3426
    lRet = RegOpenKeyEx(HKEY_CLASSES_ROOT, aKey$, 0, KEY_ALL_ACCESS, hKey)
vbwProfiler.vbwExecuteLine 3427
    If hKey Then
vbwProfiler.vbwExecuteLine 3428
        aRootFile = Space$(255)
vbwProfiler.vbwExecuteLine 3429
        lSize = 255
vbwProfiler.vbwExecuteLine 3430
        lRet = RegQueryValueEx(hKey, "", 0&, REG_SZ, ByVal aRootFile, lSize)
vbwProfiler.vbwExecuteLine 3431
        lRet = RegCloseKey(hKey)
vbwProfiler.vbwExecuteLine 3432
        lSize = InStr(aRootFile, Chr$(0))
vbwProfiler.vbwExecuteLine 3433
        If lSize Then
vbwProfiler.vbwExecuteLine 3434
            aRootFile = Left$(aRootFile, lSize - 1)
        Else
vbwProfiler.vbwExecuteLine 3435 'B
vbwProfiler.vbwExecuteLine 3436
            aRootFile = ""
        End If
vbwProfiler.vbwExecuteLine 3437 'B
    End If
vbwProfiler.vbwExecuteLine 3438 'B
vbwProfiler.vbwExecuteLine 3439
    If aRootFile = "" Then
vbwProfiler.vbwExecuteLine 3440
        aRootFile = Mid$(aExt, 2) & "file"
vbwProfiler.vbwExecuteLine 3441
        SetKey aExt$, aRootFile
    End If
vbwProfiler.vbwExecuteLine 3442 'B
vbwProfiler.vbwProcOut 130
vbwProfiler.vbwExecuteLine 3443
End Sub

Private Sub SetKey(aKey$, aDat$)
vbwProfiler.vbwProcIn 131
    Dim lRet As Long, hKey As Long, lDisp As Long
vbwProfiler.vbwExecuteLine 3444
    lRet = RegCreateKeyEx(HKEY_CLASSES_ROOT, aKey$, 0, vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, 0, hKey, lDisp)
vbwProfiler.vbwExecuteLine 3445
    If hKey Then
vbwProfiler.vbwExecuteLine 3446
        If Len(aDat$) Then
vbwProfiler.vbwExecuteLine 3447
            lRet = RegSetValueEx(hKey, "", 0&, REG_SZ, ByVal aDat$, Len(aDat$))
        Else
vbwProfiler.vbwExecuteLine 3448 'B
vbwProfiler.vbwExecuteLine 3449
            lRet = RegDeleteValue(hKey, "")
        End If
vbwProfiler.vbwExecuteLine 3450 'B
vbwProfiler.vbwExecuteLine 3451
        lRet = RegCloseKey(hKey)
    End If
vbwProfiler.vbwExecuteLine 3452 'B
vbwProfiler.vbwProcOut 131
vbwProfiler.vbwExecuteLine 3453
End Sub

Public Sub SetAssociation()
vbwProfiler.vbwProcIn 132
vbwProfiler.vbwExecuteLine 3454
    aError = ""
vbwProfiler.vbwExecuteLine 3455
    If aExt = "" Then
vbwProfiler.vbwExecuteLine 3456
         aError = "No extension specified for {GetAssocation} method."
vbwProfiler.vbwProcOut 132
vbwProfiler.vbwExecuteLine 3457
         Exit Sub
    End If
vbwProfiler.vbwExecuteLine 3458 'B
    Dim hKey As Long, lRet As Long, aKey$, aDat$, lSize As Long, lDisp As Long

vbwProfiler.vbwExecuteLine 3459
    Call GetRootFile

vbwProfiler.vbwExecuteLine 3460
    aKey$ = aRootFile
vbwProfiler.vbwExecuteLine 3461
    Call SetKey(aKey$, aDesc)

vbwProfiler.vbwExecuteLine 3462
    aKey$ = aRootFile & "\DefaultIcon"
vbwProfiler.vbwExecuteLine 3463
    Call SetKey(aKey$, aIcon)

vbwProfiler.vbwExecuteLine 3464
    aKey$ = aRootFile & "\shell\open\command"
vbwProfiler.vbwExecuteLine 3465
    Call SetKey(aKey$, aOpen)

    'aKey$ = aRootFile & "\shell\print\command"
    'Call SetKey(aKey$, aPrint)
vbwProfiler.vbwProcOut 132
vbwProfiler.vbwExecuteLine 3466
End Sub

Private Function GetKey(aKey$) As String
vbwProfiler.vbwProcIn 133
    Dim lRet As Long, lSize As Long, hKey As Long, aDat$
vbwProfiler.vbwExecuteLine 3467
    lRet = RegOpenKeyEx(HKEY_CLASSES_ROOT, aKey$, 0, KEY_ALL_ACCESS, hKey)
vbwProfiler.vbwExecuteLine 3468
    If hKey Then
vbwProfiler.vbwExecuteLine 3469
        aDat$ = Space$(255)
vbwProfiler.vbwExecuteLine 3470
        lSize = 255
vbwProfiler.vbwExecuteLine 3471
        lRet = RegQueryValueEx(hKey, "", 0&, REG_SZ, ByVal aDat$, lSize)
vbwProfiler.vbwExecuteLine 3472
        lRet = RegCloseKey(hKey)
vbwProfiler.vbwExecuteLine 3473
        lSize = InStr(aDat$, Chr$(0))
vbwProfiler.vbwExecuteLine 3474
        If lSize Then
vbwProfiler.vbwExecuteLine 3475
            aDat$ = Left$(aDat$, lSize - 1)
        Else
vbwProfiler.vbwExecuteLine 3476 'B
vbwProfiler.vbwExecuteLine 3477
            aDat$ = ""
        End If
vbwProfiler.vbwExecuteLine 3478 'B
    End If
vbwProfiler.vbwExecuteLine 3479 'B
vbwProfiler.vbwExecuteLine 3480
    GetKey = aDat$
vbwProfiler.vbwProcOut 133
vbwProfiler.vbwExecuteLine 3481
End Function

Public Sub GetAssociation()
vbwProfiler.vbwProcIn 134
vbwProfiler.vbwExecuteLine 3482
    aError = ""
vbwProfiler.vbwExecuteLine 3483
    If aExt = "" Then
vbwProfiler.vbwExecuteLine 3484
         aError = "No extension specified for {GetAssocation} method."
vbwProfiler.vbwProcOut 134
vbwProfiler.vbwExecuteLine 3485
         Exit Sub
    End If
vbwProfiler.vbwExecuteLine 3486 'B
    Dim hKey As Long, lRet As Long, aKey$, lSize As Long
vbwProfiler.vbwExecuteLine 3487
    aDesc = ""
vbwProfiler.vbwExecuteLine 3488
    aIcon = ""
vbwProfiler.vbwExecuteLine 3489
    aOpen = ""
vbwProfiler.vbwExecuteLine 3490
    aPrint = ""

vbwProfiler.vbwExecuteLine 3491
    Call GetRootFile

vbwProfiler.vbwExecuteLine 3492
    aKey$ = aRootFile
vbwProfiler.vbwExecuteLine 3493
    aDesc = GetKey(aKey$)

vbwProfiler.vbwExecuteLine 3494
    aKey$ = aRootFile & "\DefaultIcon"
vbwProfiler.vbwExecuteLine 3495
    aIcon = GetKey(aKey$)

vbwProfiler.vbwExecuteLine 3496
    aKey$ = aRootFile & "\shell\open\command"
vbwProfiler.vbwExecuteLine 3497
    aOpen = GetKey(aKey$)

vbwProfiler.vbwExecuteLine 3498
    aKey$ = aRootFile & "\shell\print\command"
vbwProfiler.vbwExecuteLine 3499
    aPrint = GetKey(aKey$)
vbwProfiler.vbwProcOut 134
vbwProfiler.vbwExecuteLine 3500
End Sub

Public Property Get LastError() As String
vbwProfiler.vbwProcIn 135
vbwProfiler.vbwExecuteLine 3501
    LastError = aError
vbwProfiler.vbwProcOut 135
vbwProfiler.vbwExecuteLine 3502
End Property

Public Property Get Description() As String
vbwProfiler.vbwProcIn 136
vbwProfiler.vbwExecuteLine 3503
    Description = aDesc
vbwProfiler.vbwProcOut 136
vbwProfiler.vbwExecuteLine 3504
End Property

Public Property Let Description(ByVal Param As String)
vbwProfiler.vbwProcIn 137
vbwProfiler.vbwExecuteLine 3505
    aDesc = Param
vbwProfiler.vbwProcOut 137
vbwProfiler.vbwExecuteLine 3506
End Property

Public Property Get Extension() As String
vbwProfiler.vbwProcIn 138
vbwProfiler.vbwExecuteLine 3507
    Extension = aExt
vbwProfiler.vbwProcOut 138
vbwProfiler.vbwExecuteLine 3508
End Property

Public Property Let Extension(ByVal Param As String)
vbwProfiler.vbwProcIn 139
vbwProfiler.vbwExecuteLine 3509
    aExt = LCase$(Param)
vbwProfiler.vbwExecuteLine 3510
    If Len(aExt) Then
vbwProfiler.vbwExecuteLine 3511
        If Left$(aExt, 1) <> "." Then
vbwProfiler.vbwExecuteLine 3512
             aExt = "." & aExt
        End If
vbwProfiler.vbwExecuteLine 3513 'B
    End If
vbwProfiler.vbwExecuteLine 3514 'B
vbwProfiler.vbwProcOut 139
vbwProfiler.vbwExecuteLine 3515
End Property

Public Property Get DefaultIcon() As String
vbwProfiler.vbwProcIn 140
vbwProfiler.vbwExecuteLine 3516
    DefaultIcon = aIcon
vbwProfiler.vbwProcOut 140
vbwProfiler.vbwExecuteLine 3517
End Property

Public Property Let DefaultIcon(ByVal Param As String)
vbwProfiler.vbwProcIn 141
vbwProfiler.vbwExecuteLine 3518
    aIcon = Param
vbwProfiler.vbwProcOut 141
vbwProfiler.vbwExecuteLine 3519
End Property

Public Property Get OpenCommand() As String
vbwProfiler.vbwProcIn 142
vbwProfiler.vbwExecuteLine 3520
    OpenCommand = aOpen
vbwProfiler.vbwProcOut 142
vbwProfiler.vbwExecuteLine 3521
End Property

Public Property Let OpenCommand(ByVal Param As String)
vbwProfiler.vbwProcIn 143
vbwProfiler.vbwExecuteLine 3522
    aOpen = Param
vbwProfiler.vbwProcOut 143
vbwProfiler.vbwExecuteLine 3523
End Property

Public Property Get PrintCommand() As String
vbwProfiler.vbwProcIn 144
vbwProfiler.vbwExecuteLine 3524
    PrintCommand = aPrint
vbwProfiler.vbwProcOut 144
vbwProfiler.vbwExecuteLine 3525
End Property

Public Property Let PrintCommand(ByVal Param As String)
vbwProfiler.vbwProcIn 145
vbwProfiler.vbwExecuteLine 3526
    aPrint = Param
vbwProfiler.vbwProcOut 145
vbwProfiler.vbwExecuteLine 3527
End Property


