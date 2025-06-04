Attribute VB_Name = "modTextOutput"
Option Explicit
Private sBreak As String
Private sLineBreak As String
Private bSlimline As Boolean

Private Const CREATED_WITH = "Created with GURPS Vehicle Designer 2.0"
Private Const GVD_URL = "http://www.makosoft.com/gvd"


Public Function createGURPSText(ByVal sType As String) As String
vbwProfiler.vbwProcIn 39
    Dim sOutput As String
    Dim sTemp As String
    Dim sTagline As String

    #If DEBUG_MODE Then
vbwProfiler.vbwExecuteLine 1051
        MsgBox "modTextOutput:createGURPSText() - Function not available in Debug Mode."
vbwProfiler.vbwProcOut 39
vbwProfiler.vbwExecuteLine 1052
        Exit Function
    #End If

    'jaw 2000.06.25
    'reformed to select case to allow for additional exports to be easily added
vbwProfiler.vbwExecuteLine 1053
    Select Case sType
'vbwLine 1054:        Case "Text"
        Case IIf(vbwProfiler.vbwExecuteLine(1054), VBWPROFILER_EMPTY, _
        "Text")
vbwProfiler.vbwExecuteLine 1055
            sBreak = Chr(13) + Chr(10) + Chr(13) + Chr(10)
vbwProfiler.vbwExecuteLine 1056
            sTagline = CREATED_WITH + Chr(13) + Chr(10) + GVD_URL
vbwProfiler.vbwExecuteLine 1057
            sLineBreak = Chr(13) + Chr(10)
'vbwLine 1058:        Case "Text Slim"
        Case IIf(vbwProfiler.vbwExecuteLine(1058), VBWPROFILER_EMPTY, _
        "Text Slim")
vbwProfiler.vbwExecuteLine 1059
            sBreak = Chr(13) + Chr(10) '+ Chr(13) + Chr(10)
vbwProfiler.vbwExecuteLine 1060
            sTagline = CREATED_WITH + Chr(13) + Chr(10) + GVD_URL + Chr(13) + Chr(10)
vbwProfiler.vbwExecuteLine 1061
            sLineBreak = Chr(13) + Chr(10)
vbwProfiler.vbwExecuteLine 1062
            bSlimline = True
'vbwLine 1063:        Case "Class HTML"
        Case IIf(vbwProfiler.vbwExecuteLine(1063), VBWPROFILER_EMPTY, _
        "Class HTML")
vbwProfiler.vbwExecuteLine 1064
            sBreak = "<BR> <BR>" & vbNewLine & vbNewLine
vbwProfiler.vbwExecuteLine 1065
            sTagline = CREATED_WITH & "<BR>" & GVD_URL & vbCrLf & "</body></html>"
vbwProfiler.vbwExecuteLine 1066
            sLineBreak = "<BR>"
'vbwLine 1067:        Case "New HTML"
        Case IIf(vbwProfiler.vbwExecuteLine(1067), VBWPROFILER_EMPTY, _
        "New HTML")
vbwProfiler.vbwExecuteLine 1068
            sBreak = "<BR> <BR>" & vbNewLine & vbNewLine
vbwProfiler.vbwExecuteLine 1069
            sTagline = CREATED_WITH & "<BR>" & GVD_URL & vbCrLf & "</body></html>"
vbwProfiler.vbwExecuteLine 1070
            sLineBreak = "<BR>"
        Case Else
vbwProfiler.vbwExecuteLine 1071 'B
vbwProfiler.vbwProcOut 39
vbwProfiler.vbwExecuteLine 1072
            Exit Function
    End Select
vbwProfiler.vbwExecuteLine 1073 'B
vbwProfiler.vbwExecuteLine 1074
    On Error Resume Next
    'get header, vehicle name, copyright info and description
vbwProfiler.vbwExecuteLine 1075
    sOutput = GetHeaderOutput(sType)
vbwProfiler.vbwExecuteLine 1076
    sOutput = sOutput + "Subassemblies and Body Features: " + GetSubassemblyOutput + GetBodyFeatures + sBreak
vbwProfiler.vbwExecuteLine 1077
    sTemp = GetCustomComponentsOutput
vbwProfiler.vbwExecuteLine 1078
    If sTemp <> "" Then
vbwProfiler.vbwExecuteLine 1079
         sOutput = sOutput + "Custom Components: " + sTemp + sBreak
    End If
vbwProfiler.vbwExecuteLine 1080 'B
vbwProfiler.vbwExecuteLine 1081
    sTemp = GetPropulsionOutput
vbwProfiler.vbwExecuteLine 1082
    If sTemp <> "" Then
vbwProfiler.vbwExecuteLine 1083
         sOutput = sOutput + "Propulsion: " + sTemp + sBreak
    End If
vbwProfiler.vbwExecuteLine 1084 'B
vbwProfiler.vbwExecuteLine 1085
    sTemp = GetAerostaticLiftOutput
vbwProfiler.vbwExecuteLine 1086
    If sTemp <> "" Then
vbwProfiler.vbwExecuteLine 1087
         sOutput = sOutput + "Aerostatic Lift: " + sTemp + sBreak
    End If
vbwProfiler.vbwExecuteLine 1088 'B
vbwProfiler.vbwExecuteLine 1089
    sTemp = GetWeaponryOutput
vbwProfiler.vbwExecuteLine 1090
    If sTemp <> "" Then
vbwProfiler.vbwExecuteLine 1091
         sOutput = sOutput + "Weaponry: " + sTemp + sBreak
    End If
vbwProfiler.vbwExecuteLine 1092 'B
vbwProfiler.vbwExecuteLine 1093
    sTemp = GetWeaponLinksOutput
vbwProfiler.vbwExecuteLine 1094
    If sTemp <> "" Then
vbwProfiler.vbwExecuteLine 1095
         sOutput = sOutput + "Weapon Links: " + sTemp + sBreak
    End If
vbwProfiler.vbwExecuteLine 1096 'B
vbwProfiler.vbwExecuteLine 1097
    sTemp = GetWeaponAccessoriesOutput
vbwProfiler.vbwExecuteLine 1098
    If sTemp <> "" Then
vbwProfiler.vbwExecuteLine 1099
         sOutput = sOutput + "Weapon Accessories: " + sTemp + sBreak
    End If
vbwProfiler.vbwExecuteLine 1100 'B
vbwProfiler.vbwExecuteLine 1101
    sTemp = GetCommunicationsOutput
vbwProfiler.vbwExecuteLine 1102
    If sTemp <> "" Then
vbwProfiler.vbwExecuteLine 1103
         sOutput = sOutput + "Communications: " + sTemp + sBreak
    End If
vbwProfiler.vbwExecuteLine 1104 'B
vbwProfiler.vbwExecuteLine 1105
    sTemp = GetSensorsOutput
vbwProfiler.vbwExecuteLine 1106
    If sTemp <> "" Then
vbwProfiler.vbwExecuteLine 1107
         sOutput = sOutput + "Sensors: " + sTemp + sBreak
    End If
vbwProfiler.vbwExecuteLine 1108 'B
vbwProfiler.vbwExecuteLine 1109
    sTemp = GetAudioVisualOutput
vbwProfiler.vbwExecuteLine 1110
    If sTemp <> "" Then
vbwProfiler.vbwExecuteLine 1111
         sOutput = sOutput + "Audio/Visual: " + sTemp + sBreak
    End If
vbwProfiler.vbwExecuteLine 1112 'B
vbwProfiler.vbwExecuteLine 1113
    sTemp = GetNavigationOutput
vbwProfiler.vbwExecuteLine 1114
    If sTemp <> "" Then
vbwProfiler.vbwExecuteLine 1115
         sOutput = sOutput + "Navigation: " + sTemp + sBreak
    End If
vbwProfiler.vbwExecuteLine 1116 'B
vbwProfiler.vbwExecuteLine 1117
    sTemp = GetTargetingOutput
vbwProfiler.vbwExecuteLine 1118
    If sTemp <> "" Then
vbwProfiler.vbwExecuteLine 1119
         sOutput = sOutput + "Targeting: " + sTemp + sBreak
    End If
vbwProfiler.vbwExecuteLine 1120 'B
vbwProfiler.vbwExecuteLine 1121
    sTemp = GetECMOutput
vbwProfiler.vbwExecuteLine 1122
    If sTemp <> "" Then
vbwProfiler.vbwExecuteLine 1123
         sOutput = sOutput + "ECM: " + sTemp + sBreak
    End If
vbwProfiler.vbwExecuteLine 1124 'B
vbwProfiler.vbwExecuteLine 1125
    sTemp = GetComputersOutput
vbwProfiler.vbwExecuteLine 1126
    If sTemp <> "" Then
vbwProfiler.vbwExecuteLine 1127
         sOutput = sOutput + "Computers: " + sTemp + sBreak
    End If
vbwProfiler.vbwExecuteLine 1128 'B
vbwProfiler.vbwExecuteLine 1129
    sTemp = GetSoftwareOutput
vbwProfiler.vbwExecuteLine 1130
    If sTemp <> "" Then
vbwProfiler.vbwExecuteLine 1131
         sOutput = sOutput + "Software: " + sTemp + sBreak
    End If
vbwProfiler.vbwExecuteLine 1132 'B
vbwProfiler.vbwExecuteLine 1133
    sTemp = GetMiscellaneousOutput
vbwProfiler.vbwExecuteLine 1134
    If sTemp <> "" Then
vbwProfiler.vbwExecuteLine 1135
         sOutput = sOutput + "Miscellaneous: " + sTemp + sBreak
    End If
vbwProfiler.vbwExecuteLine 1136 'B
vbwProfiler.vbwExecuteLine 1137
    sTemp = GetVehicleControlsOutput
vbwProfiler.vbwExecuteLine 1138
    If sTemp <> "" Then
vbwProfiler.vbwExecuteLine 1139
         sOutput = sOutput + "Vehicle Controls: " + sTemp + sBreak
    End If
vbwProfiler.vbwExecuteLine 1140 'B
vbwProfiler.vbwExecuteLine 1141
    sTemp = GetNeuralInterfaceSystemOutput
vbwProfiler.vbwExecuteLine 1142
    If sTemp <> "" Then
vbwProfiler.vbwExecuteLine 1143
         sOutput = sOutput + "Neural Interfaces: " + sTemp + sBreak
    End If
vbwProfiler.vbwExecuteLine 1144 'B
vbwProfiler.vbwExecuteLine 1145
    sTemp = GetCrewStationsOutput
vbwProfiler.vbwExecuteLine 1146
    If sTemp <> "" Then
vbwProfiler.vbwExecuteLine 1147
         sOutput = sOutput + "Crew Stations: " + sTemp + sBreak
    End If
vbwProfiler.vbwExecuteLine 1148 'B
vbwProfiler.vbwExecuteLine 1149
    sTemp = GetOccupancyOutput
vbwProfiler.vbwExecuteLine 1150
    If sTemp <> "" Then
vbwProfiler.vbwExecuteLine 1151
         sOutput = sOutput + "Occupancy: " + sTemp + sBreak
    End If
vbwProfiler.vbwExecuteLine 1152 'B
vbwProfiler.vbwExecuteLine 1153
    sTemp = GetAccomodationsOutput
vbwProfiler.vbwExecuteLine 1154
    If sTemp <> "" Then
vbwProfiler.vbwExecuteLine 1155
         sOutput = sOutput + "Accommodations: " + sTemp + sBreak
    End If
vbwProfiler.vbwExecuteLine 1156 'B
vbwProfiler.vbwExecuteLine 1157
    sTemp = GetEnvironmentalSystemsOutput
vbwProfiler.vbwExecuteLine 1158
    If sTemp <> "" Then
vbwProfiler.vbwExecuteLine 1159
         sOutput = sOutput + "Environmental Systems: " + sTemp + sBreak
    End If
vbwProfiler.vbwExecuteLine 1160 'B
vbwProfiler.vbwExecuteLine 1161
    sTemp = GetSafetySystemsOutput
vbwProfiler.vbwExecuteLine 1162
    If sTemp <> "" Then
vbwProfiler.vbwExecuteLine 1163
         sOutput = sOutput + "Safety Systems: " + sTemp + sBreak
    End If
vbwProfiler.vbwExecuteLine 1164 'B
vbwProfiler.vbwExecuteLine 1165
    sTemp = GetPowerSystemsOutPut
vbwProfiler.vbwExecuteLine 1166
    If sTemp <> "" Then
vbwProfiler.vbwExecuteLine 1167
         sOutput = sOutput + "Power Systems: " + sTemp + sBreak
    End If
vbwProfiler.vbwExecuteLine 1168 'B
vbwProfiler.vbwExecuteLine 1169
    sTemp = GetFuelOutput
vbwProfiler.vbwExecuteLine 1170
    If sTemp <> "" Then
vbwProfiler.vbwExecuteLine 1171
         sOutput = sOutput + "Fuel: " + sTemp + sBreak
    End If
vbwProfiler.vbwExecuteLine 1172 'B
vbwProfiler.vbwExecuteLine 1173
    sTemp = GetSpaceOutput
vbwProfiler.vbwExecuteLine 1174
    If sTemp <> "" Then
vbwProfiler.vbwExecuteLine 1175
         sOutput = sOutput + sTemp + sBreak
    End If
vbwProfiler.vbwExecuteLine 1176 'B
vbwProfiler.vbwExecuteLine 1177
    sTemp = GetSurfaceAreaOutput
vbwProfiler.vbwExecuteLine 1178
    If sTemp <> "" Then
vbwProfiler.vbwExecuteLine 1179
         sOutput = sOutput + "Surface Area: " + sTemp + sBreak
    End If
vbwProfiler.vbwExecuteLine 1180 'B
vbwProfiler.vbwExecuteLine 1181
    sTemp = GetStructureOutput
vbwProfiler.vbwExecuteLine 1182
    If sTemp <> "" Then
vbwProfiler.vbwExecuteLine 1183
         sOutput = sOutput + "Structure: " + sTemp + sBreak
    End If
vbwProfiler.vbwExecuteLine 1184 'B
vbwProfiler.vbwExecuteLine 1185
    sTemp = GetHitPointsOutput
vbwProfiler.vbwExecuteLine 1186
    If sTemp <> "" Then
vbwProfiler.vbwExecuteLine 1187
         sOutput = sOutput + "Hit Points: " + sTemp + sBreak
    End If
vbwProfiler.vbwExecuteLine 1188 'B
vbwProfiler.vbwExecuteLine 1189
    sTemp = GetStructuralOptionsOutput
vbwProfiler.vbwExecuteLine 1190
    If sTemp <> "" Then
vbwProfiler.vbwExecuteLine 1191
         sOutput = sOutput + sTemp + sBreak
    End If
vbwProfiler.vbwExecuteLine 1192 'B
vbwProfiler.vbwExecuteLine 1193
    sTemp = GetArmorOutput
vbwProfiler.vbwExecuteLine 1194
    If sTemp <> "" Then
vbwProfiler.vbwExecuteLine 1195
         sOutput = sOutput + sTemp + sBreak
    End If
vbwProfiler.vbwExecuteLine 1196 'B
vbwProfiler.vbwExecuteLine 1197
    sTemp = GetSurfaceFeaturesOutput
vbwProfiler.vbwExecuteLine 1198
    If sTemp <> "" Then
vbwProfiler.vbwExecuteLine 1199
         sOutput = sOutput + sTemp + sBreak
    End If
vbwProfiler.vbwExecuteLine 1200 'B
vbwProfiler.vbwExecuteLine 1201
    sTemp = GetDefensiveSurfaceFeatures
vbwProfiler.vbwExecuteLine 1202
    If sTemp <> "" Then
vbwProfiler.vbwExecuteLine 1203
         sOutput = sOutput + sTemp + sBreak
    End If
vbwProfiler.vbwExecuteLine 1204 'B
vbwProfiler.vbwExecuteLine 1205
    sTemp = GetOtherSurfaceFeatures
vbwProfiler.vbwExecuteLine 1206
    If sTemp <> "" Then
vbwProfiler.vbwExecuteLine 1207
         sOutput = sOutput + sTemp + sBreak
    End If
vbwProfiler.vbwExecuteLine 1208 'B
vbwProfiler.vbwExecuteLine 1209
    sTemp = GetTopDeckSurfaceFeatures
vbwProfiler.vbwExecuteLine 1210
    If sTemp <> "" Then
vbwProfiler.vbwExecuteLine 1211
         sOutput = sOutput + sTemp + sBreak
    End If
vbwProfiler.vbwExecuteLine 1212 'B
vbwProfiler.vbwExecuteLine 1213
    sTemp = GetWeaponBaysAndHardpoints
vbwProfiler.vbwExecuteLine 1214
    If sTemp <> "" Then
vbwProfiler.vbwExecuteLine 1215
         sOutput = sOutput + sTemp + sBreak
    End If
vbwProfiler.vbwExecuteLine 1216 'B
vbwProfiler.vbwExecuteLine 1217
    sTemp = GetVisionAndDetailsOutput
vbwProfiler.vbwExecuteLine 1218
    If sTemp <> "" Then
vbwProfiler.vbwExecuteLine 1219
         sOutput = sOutput + sTemp + sBreak
    End If
vbwProfiler.vbwExecuteLine 1220 'B
vbwProfiler.vbwExecuteLine 1221
    sTemp = GetStatisticsOutput
vbwProfiler.vbwExecuteLine 1222
    If sTemp <> "" Then
vbwProfiler.vbwExecuteLine 1223
         sOutput = sOutput + "Statistics: " + sTemp + sBreak
    End If
vbwProfiler.vbwExecuteLine 1224 'B
vbwProfiler.vbwExecuteLine 1225
    sTemp = GetPerformanceOutput
vbwProfiler.vbwExecuteLine 1226
    If sTemp <> "" Then
vbwProfiler.vbwExecuteLine 1227
         sOutput = sOutput + sTemp + sBreak
    End If
vbwProfiler.vbwExecuteLine 1228 'B
vbwProfiler.vbwExecuteLine 1229
    sTemp = GetDetailedWeaponStats
vbwProfiler.vbwExecuteLine 1230
    If sTemp <> "" Then
vbwProfiler.vbwExecuteLine 1231
         sOutput = sOutput + sTemp + vbNewLine
    End If
vbwProfiler.vbwExecuteLine 1232 'B

    '//add our tag line
vbwProfiler.vbwExecuteLine 1233
    sOutput = sOutput + sTagline
    'jaw 2000.06.25
vbwProfiler.vbwExecuteLine 1234
    If bSlimline Then
vbwProfiler.vbwExecuteLine 1235
        sOutput = RemoveParenthetical(sOutput)
    End If
vbwProfiler.vbwExecuteLine 1236 'B
vbwProfiler.vbwExecuteLine 1237
    createGURPSText = sOutput
vbwProfiler.vbwProcOut 39
vbwProfiler.vbwExecuteLine 1238
End Function

Private Function GetHeaderOutput(ByVal sType As String) As String
vbwProfiler.vbwProcIn 40
    Dim sOutput As String
    Dim sTemp As String
'    With m_oCurrentVeh.Description
'        sTemp = .Header
'        If sTemp <> "" Then sOutput = sTemp + sBreak
'        sTemp = .NickName
'        If sTemp <> "" Then sOutput = sOutput + "Name: " + sTemp + sLineBreak
'        sTemp = .ClassName
'        If sTemp <> "" Then sOutput = sOutput + "Class: " & sTemp + sLineBreak
'        sTemp = .category
'        If sTemp <> "" Then
'            If .subcategory <> "" Then
'                sOutput = sOutput + "Category: " & sTemp & "  Subcategory: " & .subcategory & sLineBreak
'            Else
'                sOutput = sOutput + "Category: " & sTemp & sLineBreak
'            End If
'        End If
'
'        sTemp = .CopyrightDate
'        If sTemp <> "" Then sOutput = sOutput + "Copyright (c) " + sTemp + sLineBreak
'        sTemp = .author
'        If sTemp <> "" Then
'            sOutput = sOutput + "by " + sTemp
'
'            sTemp = .email
'            If sTemp <> "" Then
'                sOutput = sOutput + " " + "<" + sTemp + ">" + sLineBreak
'            Else
'                sOutput = sOutput + sLineBreak
'            End If
'       End If
'        sTemp = .url
'        If sTemp <> "" Then sOutput = sOutput + "http://" + sTemp + sLineBreak
'
'        sTemp = .VehicleDescription
'        If sTemp <> "" Then sOutput = sOutput + sLineBreak + sTemp + sBreak
'
'    End With
'
'    'JAW 2000.06.25
'    'change header to include doc head tags for HTML
'    Select Case sType
'        Case "Class HTML", "New HTML"
'            GetHeaderOutput = "<html><head><title>" & m_oCurrentVeh.Description.Header _
'                & ", " & m_oCurrentVeh.Description.ClassName & "-class " & _
'                m_oCurrentVeh.Description.category & " " & m_oCurrentVeh.Description.subcategory _
'                & "</title></head><body>" & vbCrLf & sOutput
'        Case Else
'            GetHeaderOutput = sOutput
'    End Select
vbwProfiler.vbwProcOut 40
vbwProfiler.vbwExecuteLine 1239
End Function

Private Function CHOPCHOP(ByVal s As String) As String
vbwProfiler.vbwProcIn 41
    Dim GooGooGaga As Long
    Dim ScoobyDoo As Boolean
    Dim tempbyte() As Byte
    Dim bFlag As Boolean
    Dim i As Long
vbwProfiler.vbwExecuteLine 1240
    On Error GoTo errorhandler
    '//this routine mangles the Print Output if the program is not registered
vbwProfiler.vbwExecuteLine 1241
    Randomize
vbwProfiler.vbwExecuteLine 1242
    tempbyte = ChopCheck
vbwProfiler.vbwExecuteLine 1243
    If (IsEmpty(tempbyte) = False) And (UBound(tempbyte) - LBound(tempbyte) = UBound(gsRegNum) - LBound(gsRegNum)) Then
vbwProfiler.vbwExecuteLine 1244
        For i = 1 To UBound(gsRegNum)
vbwProfiler.vbwExecuteLine 1245
            If tempbyte(i) = gsRegNum(i) Then
vbwProfiler.vbwExecuteLine 1246
                bFlag = True
            Else
vbwProfiler.vbwExecuteLine 1247 'B
vbwProfiler.vbwExecuteLine 1248
                bFlag = False
vbwProfiler.vbwExecuteLine 1249
                Exit For
            End If
vbwProfiler.vbwExecuteLine 1250 'B
vbwProfiler.vbwExecuteLine 1251
        Next

vbwProfiler.vbwExecuteLine 1252
        If bFlag Then
vbwProfiler.vbwExecuteLine 1253
            CHOPCHOP = s
vbwProfiler.vbwProcOut 41
vbwProfiler.vbwExecuteLine 1254
            Exit Function
'vbwLine 1255:        ElseIf ScoobyDoo Then
        ElseIf vbwProfiler.vbwExecuteLine(1255) Or ScoobyDoo Then
vbwProfiler.vbwExecuteLine 1256
            GoTo ScoobySnack
        Else
vbwProfiler.vbwExecuteLine 1257 'B

        End If
vbwProfiler.vbwExecuteLine 1258 'B
    End If
vbwProfiler.vbwExecuteLine 1259 'B
    '//mangle time
vbwProfiler.vbwExecuteLine 1260
    If Len(s) <= 1 Then
vbwProfiler.vbwExecuteLine 1261
        CHOPCHOP = ""
    Else
vbwProfiler.vbwExecuteLine 1262 'B
vbwProfiler.vbwExecuteLine 1263
        For i = 1 To Len(s)
vbwProfiler.vbwExecuteLine 1264
            Mid(s, i, 1) = Chr(Int((255 - 0 + 1) * Rnd))
vbwProfiler.vbwExecuteLine 1265
        Next
    End If
vbwProfiler.vbwExecuteLine 1266 'B

    '//add some fake code that never gets executed
vbwProfiler.vbwExecuteLine 1267
    If GooGooGaga = -74439050 Then
vbwProfiler.vbwExecuteLine 1268
        GooGooGaga = 85858859
    End If
vbwProfiler.vbwExecuteLine 1269 'B
vbwProfiler.vbwExecuteLine 1270
    CHOPCHOP = s
vbwProfiler.vbwExecuteLine 1271
    If ScoobyDoo Then
vbwProfiler.vbwExecuteLine 1272
         GoTo ScoobySnack
    End If
vbwProfiler.vbwExecuteLine 1273 'B
vbwProfiler.vbwProcOut 41
vbwProfiler.vbwExecuteLine 1274
    Exit Function
ScoobySnack:
        'note that this never gets called because ScoobyDoo always evaluates to False
vbwProfiler.vbwExecuteLine 1275
        Resume Next
errorhandler:
vbwProfiler.vbwProcOut 41
vbwProfiler.vbwExecuteLine 1276
End Function

Public Function ChopCheck() As Byte()
vbwProfiler.vbwProcIn 42
    Dim tempbyte() As Byte
    Dim i As Long
    Dim j As Single
    Dim sTName As String
    Dim lngtotal As Single
    Dim sRegNumber As String

vbwProfiler.vbwExecuteLine 1277
    ReDim tempbyte(1)

#If DEBUG_MODE Then
vbwProfiler.vbwExecuteLine 1278
    MsgBox "modTextOutput:ChopCheck() -- Function not available in debug mode."
vbwProfiler.vbwProcOut 42
vbwProfiler.vbwExecuteLine 1279
    Exit Function
#End If
vbwProfiler.vbwExecuteLine 1280
    On Error GoTo errorhandler
    '//one of the local reg number checkers.  There will be several of these so
    ' so that a hacker will have to do some serious code hacking to disable all
    ' of them

    'here's the reg key formula
    '1- the user's reg name and key are accepted into a byte array with each
    '   letter being actually the ascii code for that letter.  Total them up
vbwProfiler.vbwExecuteLine 1281
    For i = 1 To UBound(gsRegName)
vbwProfiler.vbwExecuteLine 1282
        lngtotal = lngtotal + gsRegName(i)
        'at the same time total the ascii value for every even valued ascii code
vbwProfiler.vbwExecuteLine 1283
        If gsRegName(i) Mod 2 = 0 Then
vbwProfiler.vbwExecuteLine 1284
            lngtotal = lngtotal + gsRegName(i)
        End If
vbwProfiler.vbwExecuteLine 1285 'B
vbwProfiler.vbwExecuteLine 1286
    Next
    '2 - the RegID is actually just a modifier to prevent two people having the same
    '    name winding up with the same ID.  This ID is unique and alone can be used
    '   to identify a user.  Multiply this to the total
vbwProfiler.vbwExecuteLine 1287
    lngtotal = lngtotal * gsRegID
    '3- take the ascii value of the typename of the Body and multiply that to it
vbwProfiler.vbwExecuteLine 1288
    sTName = TypeName(m_oCurrentVeh.Body) '(BODY_KEY))
vbwProfiler.vbwExecuteLine 1289
    For i = 1 To Len(sTName)
vbwProfiler.vbwExecuteLine 1290
        lngtotal = lngtotal * Asc(Mid(sTName, i, 1))
vbwProfiler.vbwExecuteLine 1291
    Next
    '6- take a random seed to generate the seeded random number and multiply that
vbwProfiler.vbwExecuteLine 1292
    Rnd -1
vbwProfiler.vbwExecuteLine 1293
    Randomize 9921988
vbwProfiler.vbwExecuteLine 1294
    lngtotal = lngtotal * Rnd()
    '8- return this as a byte array that we can compare with our current one
    'how do we split this up into seperate bytes? well we know our ascii values
    'must be between 48-57, 65-90 and 97-122
    'well, we can generate a random reg code based on each number in the string
    'representation using the random seed of each number
vbwProfiler.vbwExecuteLine 1295
    For i = 1 To Len(Str(lngtotal))
vbwProfiler.vbwExecuteLine 1296
        j = Rnd()
vbwProfiler.vbwExecuteLine 1297
        If j <= 0.33 Then
vbwProfiler.vbwExecuteLine 1298
            ReDim Preserve tempbyte(i)
vbwProfiler.vbwExecuteLine 1299
            Rnd -1
vbwProfiler.vbwExecuteLine 1300
            Randomize Asc(Mid(Str(lngtotal), i, 1))
vbwProfiler.vbwExecuteLine 1301
            tempbyte(i) = Int((57 - 48 + 1) * Rnd + 48)
vbwProfiler.vbwExecuteLine 1302
            sRegNumber = sRegNumber & Chr(tempbyte(i))
'vbwLine 1303:        ElseIf j <= 0.66 Then
        ElseIf vbwProfiler.vbwExecuteLine(1303) Or j <= 0.66 Then
vbwProfiler.vbwExecuteLine 1304
            ReDim Preserve tempbyte(i)
vbwProfiler.vbwExecuteLine 1305
            Rnd -1
vbwProfiler.vbwExecuteLine 1306
            Randomize Asc(Mid(Str(lngtotal), i, 1))
vbwProfiler.vbwExecuteLine 1307
            tempbyte(i) = Int((90 - 65 + 1) * Rnd + 65)
vbwProfiler.vbwExecuteLine 1308
            sRegNumber = sRegNumber & Chr(tempbyte(i))
        Else
vbwProfiler.vbwExecuteLine 1309 'B
vbwProfiler.vbwExecuteLine 1310
            ReDim Preserve tempbyte(i)
vbwProfiler.vbwExecuteLine 1311
            Rnd -1
vbwProfiler.vbwExecuteLine 1312
            Randomize Asc(Mid(Str(lngtotal), i, 1))
vbwProfiler.vbwExecuteLine 1313
            tempbyte(i) = Int((122 - 97 + 1) * Rnd + 97)
vbwProfiler.vbwExecuteLine 1314
            sRegNumber = sRegNumber & Chr(tempbyte(i))
        End If
vbwProfiler.vbwExecuteLine 1315 'B
vbwProfiler.vbwExecuteLine 1316
    Next

vbwProfiler.vbwExecuteLine 1317
    ChopCheck = tempbyte
vbwProfiler.vbwProcOut 42
vbwProfiler.vbwExecuteLine 1318
    Exit Function
errorhandler:
vbwProfiler.vbwExecuteLine 1319
        ReDim tempbyte(1)
vbwProfiler.vbwExecuteLine 1320
        ChopCheck = tempbyte
vbwProfiler.vbwProcOut 42
vbwProfiler.vbwExecuteLine 1321
End Function

Private Function GetSubassemblyOutput() As String
vbwProfiler.vbwProcIn 43
    Dim sOutput As String
    Dim element As Object
vbwProfiler.vbwExecuteLine 1322
    On Error GoTo err

'todo: fix
'    For Each element In m_oCurrentVeh.Components
'        Select Case element.Datatype
'
'            Case Wheel, Skid, Track, Hydrofoil, Hovercraft, _
'                Leg, Arm, AutogyroRotor, TTRotor, CARotor, MMRotor, _
'                Wing, Mast, Superstructure, Turret, Popturret, _
'                OpenMount, Gasbag, Pod, SolarPanel, equipmentPod
'
'                sOutput = sOutput + element.PrintOutput + " "
'
'        End Select
'    Next

vbwProfiler.vbwExecuteLine 1323
    GetSubassemblyOutput = sOutput
vbwProfiler.vbwProcOut 43
vbwProfiler.vbwExecuteLine 1324
    Exit Function
err:
vbwProfiler.vbwExecuteLine 1325
    Debug.Print "modTextOutput:GetSubassemblyOutput -- Error #" & err.Number & " " & err.Description

vbwProfiler.vbwProcOut 43
vbwProfiler.vbwExecuteLine 1326
End Function

Private Function GetCustomComponentsOutput() As String
vbwProfiler.vbwProcIn 44
     Dim sOutput As String
    Dim element As Object

vbwProfiler.vbwExecuteLine 1327
    On Error GoTo err
'    For Each element In m_oCurrentVeh.Components
'        If TypeOf element Is clsSimpleCustom Then
'           sOutput = sOutput + element.PrintOutput + " "
'        End If
'    Next
'
'    GetCustomComponentsOutput = sOutput
'Exit Function
err:
vbwProfiler.vbwExecuteLine 1328
    Debug.Print "modTextOutput:GetCustomComponentsOutput -- Error #" & err.Number & " " & err.Description

vbwProfiler.vbwProcOut 44
vbwProfiler.vbwExecuteLine 1329
End Function

Private Function GetPropulsionOutput() As String
vbwProfiler.vbwProcIn 45
    Dim sOutput As String
    Dim element As Object

'    On Error GoTo err
'    For Each element In m_oCurrentVeh.Components
'        Select Case element.Datatype
'
'            Case WheeledDrivetrain, AllWheelDriveWheeledDrivetrain, _
'                FlexibodyDrivetrain, TrackedDrivetrain, LegDrivetrain, _
'                CARRotorDrivetrain, MMRRotorDrivetrain, TTRRotorDrivetrain, _
'                OrnithopterDrivetrain, AerialPropeller, DuctedFan, PaddleWheel, _
'                ScrewPropeller, lightScrewPropeller, DuctedPropeller, Hydrojet, _
'                MHDTunnel, RopeHarness, YokeandPoleHarness, ShaftandCollarHarness, _
'                WhiffletreeHarness, MagLevLifter, Turbojet, Turbofan, Ramjet, _
'                TurboRamjet, Hyperfan, FusionAirRam, StandardThruster, _
'                SuperThruster, MegaThruster, LiquidFuelRocket, MOXRocket, _
'                IonDrive, FissionRocket, FusionRocket, OptimizedFusion, _
'                AntimatterThermal, AntimatterPion, RowingPositions, ForeandAftRig, _
'                SquareRig, FullRig, AerialSail, AerialSailForeAftRig, lightSail, SolidRocketEngine, _
'                OrionEngine, TeleportationDrive, Hyperdrive, JumpDrive, _
'                WarpDrive, QuantumConveyor, SubQuantumConveyor, _
'                TwoQuantumConveyor
'
'                sOutput = sOutput + element.PrintOutput + " "
'
'        End Select
'    Next
'
'    GetPropulsionOutput = sOutput
'Exit Function
err:
vbwProfiler.vbwExecuteLine 1330
    Debug.Print "modTextOutput:GetPropulsionOutput -- Error #" & err.Number & " " & err.Description

vbwProfiler.vbwProcOut 45
vbwProfiler.vbwExecuteLine 1331
End Function

Private Function GetAerostaticLiftOutput() As String
vbwProfiler.vbwProcIn 46
    Dim sOutput As String
    Dim element As Object

vbwProfiler.vbwExecuteLine 1332
    On Error GoTo err
'    For Each element In m_oCurrentVeh.Components
'        Select Case element.Datatype
'
'            Case ContraGravGenerator, HotAir, Hydrogen, Helium
'
'                sOutput = sOutput + element.PrintOutput + " "
'
'        End Select
'    Next

vbwProfiler.vbwExecuteLine 1333
    GetAerostaticLiftOutput = sOutput
vbwProfiler.vbwProcOut 46
vbwProfiler.vbwExecuteLine 1334
Exit Function
err:
vbwProfiler.vbwExecuteLine 1335
    Debug.Print "modTextOutput:GetAerostaticLiftOutput -- Error #" & err.Number & " " & err.Description

vbwProfiler.vbwProcOut 46
vbwProfiler.vbwExecuteLine 1336
End Function

Private Function GetWeaponryOutput() As String
vbwProfiler.vbwProcIn 47
    Dim sOutput As String
    Dim element As Object

vbwProfiler.vbwExecuteLine 1337
    On Error GoTo err
vbwProfiler.vbwExecuteLine 1338
    For Each element In m_oCurrentVeh.Components
vbwProfiler.vbwExecuteLine 1339
        Select Case element.Datatype

'vbwLine 1340:            Case StoneThrower, BoltThrower, RepeatingBoltThrower, MuzzleLoader, BreechLoader, ManualRepeater, Revolver, MechanicalGatling, SlowAutoloader, FastAutoloader, lightAutomatic, HeavyAutomatic, ElectricGatling, BlueGreenLaser, RainbowLaser, Laser, UVLaser, IRLaser, Disruptor, ChargedParticleBeam, NeutralParticleBeam, Flamer, Screamer, Stunner, ParalysisBeam, XRayLaser, FusionBeam, GravityBeam, AntiparticleBeam, Graser, Disintegrator, Displacer, BeamedPowerTransmitter, MilitaryParalysisBeam, EnergyDrill, IronBomb, RetardedBomb, SmartBomb, SelfDestructSystem, ContactMine, ProximityMine, PressureTriggerMine, CommandTriggerMine, SmartTriggerMine, UnGuidedMissile, UnGuidedTorpedo, GuidedMissile, GuidedTorpedo, FlameThrower, WaterCannon, DisposableLauncher, MuzzleloadingLauncher, BreechloadingLauncher, ManualRepeaterLauncher, SlowAutoLoaderLauncher, FastAutoLoaderLauncher, RevolverLauncher, lightAutomaticLauncher, HeavyAutomaticLauncher
            Case IIf(vbwProfiler.vbwExecuteLine(1340), VBWPROFILER_EMPTY, _
        StoneThrower), BoltThrower, RepeatingBoltThrower, MuzzleLoader, BreechLoader, ManualRepeater, Revolver, MechanicalGatling, SlowAutoloader, FastAutoloader, lightAutomatic, HeavyAutomatic, ElectricGatling, BlueGreenLaser, RainbowLaser, Laser, UVLaser, IRLaser, Disruptor, ChargedParticleBeam, NeutralParticleBeam, Flamer, Screamer, Stunner, ParalysisBeam, XRayLaser, FusionBeam, GravityBeam, AntiparticleBeam, Graser, Disintegrator, Displacer, BeamedPowerTransmitter, MilitaryParalysisBeam, EnergyDrill, IronBomb, RetardedBomb, SmartBomb, SelfDestructSystem, ContactMine, ProximityMine, PressureTriggerMine, CommandTriggerMine, SmartTriggerMine, UnGuidedMissile, UnGuidedTorpedo, GuidedMissile, GuidedTorpedo, FlameThrower, WaterCannon, DisposableLauncher, MuzzleloadingLauncher, BreechloadingLauncher, ManualRepeaterLauncher, SlowAutoLoaderLauncher, FastAutoLoaderLauncher, RevolverLauncher, lightAutomaticLauncher, HeavyAutomaticLauncher

vbwProfiler.vbwExecuteLine 1341
                sOutput = sOutput + element.PrintOutput + " "

        End Select
vbwProfiler.vbwExecuteLine 1342 'B
vbwProfiler.vbwExecuteLine 1343
    Next

vbwProfiler.vbwExecuteLine 1344
    For Each element In m_oCurrentVeh.Components
vbwProfiler.vbwExecuteLine 1345
        If TypeOf element Is clsweaponAmmunition Then
vbwProfiler.vbwExecuteLine 1346
            sOutput = sOutput + element.PrintOutput + " "

        End If
vbwProfiler.vbwExecuteLine 1347 'B
vbwProfiler.vbwExecuteLine 1348
    Next

vbwProfiler.vbwExecuteLine 1349
    GetWeaponryOutput = sOutput
vbwProfiler.vbwProcOut 47
vbwProfiler.vbwExecuteLine 1350
    Exit Function
err:
vbwProfiler.vbwExecuteLine 1351
    Debug.Print "modTextOutput:GetWeaponryOutput -- Error #" & err.Number & " " & err.Description

vbwProfiler.vbwProcOut 47
vbwProfiler.vbwExecuteLine 1352
End Function

Private Function GetWeaponAccessoriesOutput() As String
vbwProfiler.vbwProcIn 48
    Dim sOutput As String
    Dim element As Object

vbwProfiler.vbwExecuteLine 1353
    On Error GoTo err
vbwProfiler.vbwExecuteLine 1354
    For Each element In m_oCurrentVeh.Components
vbwProfiler.vbwExecuteLine 1355
        Select Case element.Datatype

'vbwLine 1356:            Case PartialStabilizationGear, FullStabilizationGear, UniversalMount, CasemateMount, DoorMount, Cyberslave, AntiBlastMagazine
            Case IIf(vbwProfiler.vbwExecuteLine(1356), VBWPROFILER_EMPTY, _
        PartialStabilizationGear), FullStabilizationGear, UniversalMount, CasemateMount, DoorMount, Cyberslave, AntiBlastMagazine

vbwProfiler.vbwExecuteLine 1357
                sOutput = sOutput + element.PrintOutput + " "

        End Select
vbwProfiler.vbwExecuteLine 1358 'B
vbwProfiler.vbwExecuteLine 1359
    Next

vbwProfiler.vbwExecuteLine 1360
    GetWeaponAccessoriesOutput = sOutput
vbwProfiler.vbwProcOut 48
vbwProfiler.vbwExecuteLine 1361
    Exit Function
err:
vbwProfiler.vbwExecuteLine 1362
    Debug.Print "modTextOutput:GetWeaponAccessoriesOutput -- Error #" & err.Number & " " & err.Description

vbwProfiler.vbwProcOut 48
vbwProfiler.vbwExecuteLine 1363
End Function

Private Function GetWeaponLinksOutput() As String
vbwProfiler.vbwProcIn 49
    Dim sOutput As String
    Dim element As Object
    Dim sKeyArray() As String
    Dim i As Long

vbwProfiler.vbwExecuteLine 1364
    On Error GoTo err
vbwProfiler.vbwExecuteLine 1365
    For Each element In m_oCurrentVeh.Components
vbwProfiler.vbwExecuteLine 1366
        If TypeOf element Is clsWeaponLink Then
            '//append the weapons that are in the link
vbwProfiler.vbwExecuteLine 1367
                sKeyArray = element.getcurrentkeys
vbwProfiler.vbwExecuteLine 1368
                If sKeyArray(1) = "" Then
                Else
vbwProfiler.vbwExecuteLine 1369 'B
vbwProfiler.vbwExecuteLine 1370
                    sOutput = sOutput + element.Key & " controls "
vbwProfiler.vbwExecuteLine 1371
                    For i = 1 To UBound(sKeyArray)
vbwProfiler.vbwExecuteLine 1372
                        sOutput = sOutput + m_oCurrentVeh.Components(sKeyArray(i)).Description & ", "
vbwProfiler.vbwExecuteLine 1373
                    Next
                    '//delete the last "," and replace it with "."
vbwProfiler.vbwExecuteLine 1374
                    sOutput = Left(sOutput, Len(sOutput) - 2)
vbwProfiler.vbwExecuteLine 1375
                    sOutput = sOutput + ".  "

                End If
vbwProfiler.vbwExecuteLine 1376 'B

        End If
vbwProfiler.vbwExecuteLine 1377 'B
vbwProfiler.vbwExecuteLine 1378
    Next

vbwProfiler.vbwExecuteLine 1379
    GetWeaponLinksOutput = sOutput
vbwProfiler.vbwProcOut 49
vbwProfiler.vbwExecuteLine 1380
Exit Function
err:
vbwProfiler.vbwExecuteLine 1381
    Debug.Print "modTextOutput:GetWeaponLinksOutput -- Error #" & err.Number & " " & err.Description

vbwProfiler.vbwProcOut 49
vbwProfiler.vbwExecuteLine 1382
End Function
Private Function GetCommunicationsOutput() As String
vbwProfiler.vbwProcIn 50
    Dim sOutput As String
    Dim element As Object
vbwProfiler.vbwExecuteLine 1383
    On Error GoTo err
vbwProfiler.vbwExecuteLine 1384
    For Each element In m_oCurrentVeh.Components
vbwProfiler.vbwExecuteLine 1385
        Select Case element.Datatype

'vbwLine 1386:            Case RadioDirectionFinder, RadioCommunicator, TightBeamRadio, VLFRadio, CellularPhone, CellularPhonewithRadio, RadioJammer, ElfReceiver, LaserCommunicator, NeutrinoCommunicator, GravityRippleCommunicator
            Case IIf(vbwProfiler.vbwExecuteLine(1386), VBWPROFILER_EMPTY, _
        RadioDirectionFinder), RadioCommunicator, TightBeamRadio, VLFRadio, CellularPhone, CellularPhonewithRadio, RadioJammer, ElfReceiver, LaserCommunicator, NeutrinoCommunicator, GravityRippleCommunicator

vbwProfiler.vbwExecuteLine 1387
                sOutput = sOutput + element.PrintOutput + " "

        End Select
vbwProfiler.vbwExecuteLine 1388 'B
vbwProfiler.vbwExecuteLine 1389
    Next

vbwProfiler.vbwExecuteLine 1390
    GetCommunicationsOutput = sOutput
vbwProfiler.vbwProcOut 50
vbwProfiler.vbwExecuteLine 1391
    Exit Function
err:
vbwProfiler.vbwExecuteLine 1392
    Debug.Print "modTextOutput:GetCommunicationsOutput -- Error #" & err.Number & " " & err.Description

vbwProfiler.vbwProcOut 50
vbwProfiler.vbwExecuteLine 1393
End Function
Private Function GetSensorsOutput() As String
vbwProfiler.vbwProcIn 51
    Dim sOutput As String
    Dim element As Object

vbwProfiler.vbwExecuteLine 1394
    On Error GoTo err
vbwProfiler.vbwExecuteLine 1395
    For Each element In m_oCurrentVeh.Components
vbwProfiler.vbwExecuteLine 1396
        Select Case element.Datatype

'vbwLine 1397:            Case Headlight, Searchlight, InfraredSearchlight, AstronomicalInstruments, Telescope, lightAmplification, LowlightTV, ExtendableSensorPeriscope, Radar, Ladar, NavigationalRadar, AntiCollisionRadar, AESA, LowResImagingRadar, HiResImagingRadar, ActiveSonar, PassiveSonar, PassiveInfrared, Thermograph, PassiveRadar, PESA, Geophone, MAD, MultiScanner, ChemScanner, RadScanner, BioScanner, GravScanner, RangingSoundDetector, SurveillanceSoundDetector, MeteorologicalInstruments, LowResPlanetarySurveyArray, MedResPlanetarySurveyArray, HighResPlanetarySurveyArray
            Case IIf(vbwProfiler.vbwExecuteLine(1397), VBWPROFILER_EMPTY, _
        Headlight), Searchlight, InfraredSearchlight, AstronomicalInstruments, Telescope, lightAmplification, LowlightTV, ExtendableSensorPeriscope, Radar, Ladar, NavigationalRadar, AntiCollisionRadar, AESA, LowResImagingRadar, HiResImagingRadar, ActiveSonar, PassiveSonar, PassiveInfrared, Thermograph, PassiveRadar, PESA, Geophone, MAD, MultiScanner, ChemScanner, RadScanner, BioScanner, GravScanner, RangingSoundDetector, SurveillanceSoundDetector, MeteorologicalInstruments, LowResPlanetarySurveyArray, MedResPlanetarySurveyArray, HighResPlanetarySurveyArray

vbwProfiler.vbwExecuteLine 1398
                sOutput = sOutput + element.PrintOutput + " "

        End Select
vbwProfiler.vbwExecuteLine 1399 'B
vbwProfiler.vbwExecuteLine 1400
    Next

vbwProfiler.vbwExecuteLine 1401
    GetSensorsOutput = sOutput
vbwProfiler.vbwProcOut 51
vbwProfiler.vbwExecuteLine 1402
    Exit Function
err:
vbwProfiler.vbwExecuteLine 1403
    Debug.Print "modTextOutput:GetSensorsOutput -- Error #" & err.Number & " " & err.Description

vbwProfiler.vbwProcOut 51
vbwProfiler.vbwExecuteLine 1404
End Function
Private Function GetAudioVisualOutput() As String
vbwProfiler.vbwProcIn 52
    Dim sOutput As String
    Dim element As Object
vbwProfiler.vbwExecuteLine 1405
    On Error GoTo err
vbwProfiler.vbwExecuteLine 1406
    For Each element In m_oCurrentVeh.Components
vbwProfiler.vbwExecuteLine 1407
        Select Case element.Datatype

'vbwLine 1408:            Case SoundSystem, FlightRecorder, VehicleCamera, DigitalVehicleCamera, ReconCamera, DigitalReconCamera
            Case IIf(vbwProfiler.vbwExecuteLine(1408), VBWPROFILER_EMPTY, _
        SoundSystem), FlightRecorder, VehicleCamera, DigitalVehicleCamera, ReconCamera, DigitalReconCamera

vbwProfiler.vbwExecuteLine 1409
                sOutput = sOutput + element.PrintOutput + " "

        End Select
vbwProfiler.vbwExecuteLine 1410 'B
vbwProfiler.vbwExecuteLine 1411
    Next

vbwProfiler.vbwExecuteLine 1412
    GetAudioVisualOutput = sOutput
vbwProfiler.vbwProcOut 52
vbwProfiler.vbwExecuteLine 1413
    Exit Function
err:
vbwProfiler.vbwExecuteLine 1414
    Debug.Print "modTextOutput:GetAudioVisualOutput -- Error #" & err.Number & " " & err.Description

vbwProfiler.vbwProcOut 52
vbwProfiler.vbwExecuteLine 1415
End Function

Private Function GetNavigationOutput() As String
vbwProfiler.vbwProcIn 53
    Dim sOutput As String
    Dim element As Object
vbwProfiler.vbwExecuteLine 1416
    On Error GoTo err
vbwProfiler.vbwExecuteLine 1417
    For Each element In m_oCurrentVeh.Components
vbwProfiler.vbwExecuteLine 1418
        Select Case element.Datatype

'vbwLine 1419:            Case NavigationInstruments, AutoPilot, IFF, Transponder, INS, GPS, MilitaryGPS, TFR
            Case IIf(vbwProfiler.vbwExecuteLine(1419), VBWPROFILER_EMPTY, _
        NavigationInstruments), AutoPilot, IFF, Transponder, INS, GPS, MilitaryGPS, TFR

vbwProfiler.vbwExecuteLine 1420
                sOutput = sOutput + element.PrintOutput + " "

        End Select
vbwProfiler.vbwExecuteLine 1421 'B
vbwProfiler.vbwExecuteLine 1422
    Next

vbwProfiler.vbwExecuteLine 1423
    GetNavigationOutput = sOutput
vbwProfiler.vbwProcOut 53
vbwProfiler.vbwExecuteLine 1424
    Exit Function
err:
vbwProfiler.vbwExecuteLine 1425
    Debug.Print "modTextOutput:GetNavigationOutput -- Error #" & err.Number & " " & err.Description

vbwProfiler.vbwProcOut 53
vbwProfiler.vbwExecuteLine 1426
End Function

Private Function GetTargetingOutput() As String
vbwProfiler.vbwProcIn 54
    Dim sOutput As String
    Dim element As Object

vbwProfiler.vbwExecuteLine 1427
    On Error GoTo err
vbwProfiler.vbwExecuteLine 1428
    For Each element In m_oCurrentVeh.Components
vbwProfiler.vbwExecuteLine 1429
        Select Case element.Datatype

'vbwLine 1430:            Case ImprovedOpticalBombSight, AdvancedOpticalBombSight, OpticalBombSight, FireDirectionCenter, HUDWAC, PupilHUDWAC, LaserRangeFinder, LaserDesignator, LaserSpotTracker
            Case IIf(vbwProfiler.vbwExecuteLine(1430), VBWPROFILER_EMPTY, _
        ImprovedOpticalBombSight), AdvancedOpticalBombSight, OpticalBombSight, FireDirectionCenter, HUDWAC, PupilHUDWAC, LaserRangeFinder, LaserDesignator, LaserSpotTracker

vbwProfiler.vbwExecuteLine 1431
                sOutput = sOutput + element.PrintOutput + " "

        End Select
vbwProfiler.vbwExecuteLine 1432 'B
vbwProfiler.vbwExecuteLine 1433
    Next

vbwProfiler.vbwExecuteLine 1434
    GetTargetingOutput = sOutput
vbwProfiler.vbwProcOut 54
vbwProfiler.vbwExecuteLine 1435
    Exit Function
err:
vbwProfiler.vbwExecuteLine 1436
    Debug.Print "modTextOutput:GetTargetingOutput -- Error #" & err.Number & " " & err.Description

vbwProfiler.vbwProcOut 54
vbwProfiler.vbwExecuteLine 1437
End Function

Private Function GetECMOutput() As String
vbwProfiler.vbwProcIn 55
    Dim sOutput As String
    Dim element As Object

vbwProfiler.vbwExecuteLine 1438
    On Error GoTo err
vbwProfiler.vbwExecuteLine 1439
    For Each element In m_oCurrentVeh.Components
vbwProfiler.vbwExecuteLine 1440
        Select Case element.Datatype

'vbwLine 1441:            Case RadarDetector, LaserSensor, LaserRadarDetector, AreaRadarJammer, DeceptiveRadarJammer, InfraredJammer, ChaffDecoyDischarger, SmokeDecoyDischarger, FlareDecoyDischarger, SonarDecoyDischarger, HotSmokeDecoyDischarger, PrismDecoyDischarger, BlackOutGasDecoyDischarger, RadarReflector, BlipEnhancer, TEMPEST
            Case IIf(vbwProfiler.vbwExecuteLine(1441), VBWPROFILER_EMPTY, _
        RadarDetector), LaserSensor, LaserRadarDetector, AreaRadarJammer, DeceptiveRadarJammer, InfraredJammer, ChaffDecoyDischarger, SmokeDecoyDischarger, FlareDecoyDischarger, SonarDecoyDischarger, HotSmokeDecoyDischarger, PrismDecoyDischarger, BlackOutGasDecoyDischarger, RadarReflector, BlipEnhancer, TEMPEST

vbwProfiler.vbwExecuteLine 1442
                sOutput = sOutput + element.PrintOutput + " "

        End Select
vbwProfiler.vbwExecuteLine 1443 'B
vbwProfiler.vbwExecuteLine 1444
    Next

vbwProfiler.vbwExecuteLine 1445
    GetECMOutput = sOutput
vbwProfiler.vbwProcOut 55
vbwProfiler.vbwExecuteLine 1446
    Exit Function
err:
vbwProfiler.vbwExecuteLine 1447
    Debug.Print "modTextOutput:GetECMOutput -- Error #" & err.Number & " " & err.Description

vbwProfiler.vbwProcOut 55
vbwProfiler.vbwExecuteLine 1448
End Function

Private Function GetComputersOutput() As String
vbwProfiler.vbwProcIn 56
    Dim sOutput As String
    Dim element As Object

vbwProfiler.vbwExecuteLine 1449
    On Error GoTo err
vbwProfiler.vbwExecuteLine 1450
    For Each element In m_oCurrentVeh.Components
vbwProfiler.vbwExecuteLine 1451
        Select Case element.Datatype

'vbwLine 1452:            Case MacroFrame, MainFrame, MicroFrame, MiniComputer, SmallComputer, Terminal
            Case IIf(vbwProfiler.vbwExecuteLine(1452), VBWPROFILER_EMPTY, _
        MacroFrame), MainFrame, MicroFrame, MiniComputer, SmallComputer, Terminal

vbwProfiler.vbwExecuteLine 1453
               sOutput = sOutput + element.PrintOutput + " "

        End Select
vbwProfiler.vbwExecuteLine 1454 'B
vbwProfiler.vbwExecuteLine 1455
    Next

vbwProfiler.vbwExecuteLine 1456
    GetComputersOutput = sOutput
vbwProfiler.vbwProcOut 56
vbwProfiler.vbwExecuteLine 1457
    Exit Function
err:
vbwProfiler.vbwExecuteLine 1458
    Debug.Print "modTextOutput:GetComputersOutput -- Error #" & err.Number & " " & err.Description

vbwProfiler.vbwProcOut 56
vbwProfiler.vbwExecuteLine 1459
End Function

Private Function GetSoftwareOutput() As String
vbwProfiler.vbwProcIn 57
     Dim sOutput As String
    Dim element As Object

vbwProfiler.vbwExecuteLine 1460
    On Error GoTo err
vbwProfiler.vbwExecuteLine 1461
    For Each element In m_oCurrentVeh.Components
vbwProfiler.vbwExecuteLine 1462
        If TypeOf element Is clssoftware Then
vbwProfiler.vbwExecuteLine 1463
            sOutput = sOutput + element.PrintOutput + " "
        End If
vbwProfiler.vbwExecuteLine 1464 'B
vbwProfiler.vbwExecuteLine 1465
    Next


vbwProfiler.vbwExecuteLine 1466
GetSoftwareOutput = sOutput
vbwProfiler.vbwProcOut 57
vbwProfiler.vbwExecuteLine 1467
Exit Function
err:
vbwProfiler.vbwExecuteLine 1468
    Debug.Print "modTextOutput:GetSoftwareOutput -- Error #" & err.Number & " " & err.Description

vbwProfiler.vbwProcOut 57
vbwProfiler.vbwExecuteLine 1469
End Function


Private Function GetNeuralInterfaceSystemOutput() As String
vbwProfiler.vbwProcIn 58
    Dim sOutput As String
    Dim element As Object

vbwProfiler.vbwExecuteLine 1470
    On Error GoTo err

vbwProfiler.vbwExecuteLine 1471
     For Each element In m_oCurrentVeh.Components
vbwProfiler.vbwExecuteLine 1472
        If TypeOf element Is clsneuralinterfacesystem Then
vbwProfiler.vbwExecuteLine 1473
            sOutput = sOutput + element.PrintOutput + " "
        End If
vbwProfiler.vbwExecuteLine 1474 'B
vbwProfiler.vbwExecuteLine 1475
    Next

vbwProfiler.vbwExecuteLine 1476
    GetNeuralInterfaceSystemOutput = sOutput
vbwProfiler.vbwProcOut 58
vbwProfiler.vbwExecuteLine 1477
    Exit Function
err:
vbwProfiler.vbwExecuteLine 1478
    Debug.Print "modTextOutput:GetNeuralInterfaceSystemOutput -- Error #" & err.Number & " " & err.Description

vbwProfiler.vbwProcOut 58
vbwProfiler.vbwExecuteLine 1479
End Function

Private Function GetMiscellaneousOutput() As String
vbwProfiler.vbwProcIn 59
    Dim sOutput As String
    Dim element As Object

vbwProfiler.vbwExecuteLine 1480
    On Error GoTo err
vbwProfiler.vbwExecuteLine 1481
    For Each element In m_oCurrentVeh.Components
vbwProfiler.vbwExecuteLine 1482
        Select Case element.Datatype

'vbwLine 1483:            Case ArmMotor, FireExtinguisherSystem, FullFireSuppressionSystem, CompactFireSuppressionSystem, BilgePump, CompleteWorkshop, MechanicWorkshop, EngineeringWorkshop, ElectronicsWorkshop, ArmouryWorkshop, CompleteMiniWorkshop, ScienceLab, MiniMechanicWorkshop, MiniElectronicsWorkshop, MiniEngineeringWorkshop, MiniArmouryWorkshop
            Case IIf(vbwProfiler.vbwExecuteLine(1483), VBWPROFILER_EMPTY, _
        ArmMotor), FireExtinguisherSystem, FullFireSuppressionSystem, CompactFireSuppressionSystem, BilgePump, CompleteWorkshop, MechanicWorkshop, EngineeringWorkshop, ElectronicsWorkshop, ArmouryWorkshop, CompleteMiniWorkshop, ScienceLab, MiniMechanicWorkshop, MiniElectronicsWorkshop, MiniEngineeringWorkshop, MiniArmouryWorkshop

vbwProfiler.vbwExecuteLine 1484
                sOutput = sOutput + element.PrintOutput + " "

'vbwLine 1485:            Case ExtendableLadder, Crane, Winch, PowerShovel, WreckingCrane, ForkLift, VehicularBridge, LaunchCatapult, SkyHook, Bore, SuperBore, EnergyDrill, TractorBeam, PressorBeam, CombinationBeam, CraneWithElectroMagnet, OperatingRoom, StretcherPallet, EmergencySupportUnit, EmergencylightsandSiren, CryonicCapsule, Automed, DiagnosisTable, Stage, Hall, BarRoom, ConferenceRoom, MovieScreenandProjector, MovieScreenandProjectorSmall, HoloventureZone, CargoRamp, Airlock, MembraneAirlock, Forcelock, PassageTube, ArmoredPassageTube, TeleportProjector, BrigsandRestraints, BurglarAlarm, HighSecurityAlarm, MutableLicensePlate, OilSprayer, PaintSprayer, SmokeScreen, SpikeDropper, VehicleBay, HangerBay, DryDock, SpaceDock, ExternalCradle, ArrestorHook, VehicularParachute, RefuellingProbe, RefuellingDrogue, FuelElectrolysisSystem, HydrogenFuelScoop, AtmosphereProcessor, NuclearDamper, SmallRealityStabilizer, MediumRealityStabilizer, HeavyRealityStabilizer, ModularSocket, Module
            Case IIf(vbwProfiler.vbwExecuteLine(1485), VBWPROFILER_EMPTY, _
        ExtendableLadder), Crane, Winch, PowerShovel, WreckingCrane, ForkLift, VehicularBridge, LaunchCatapult, SkyHook, Bore, SuperBore, EnergyDrill, TractorBeam, PressorBeam, CombinationBeam, CraneWithElectroMagnet, OperatingRoom, StretcherPallet, EmergencySupportUnit, EmergencylightsandSiren, CryonicCapsule, Automed, DiagnosisTable, Stage, Hall, BarRoom, ConferenceRoom, MovieScreenandProjector, MovieScreenandProjectorSmall, HoloventureZone, CargoRamp, Airlock, MembraneAirlock, Forcelock, PassageTube, ArmoredPassageTube, TeleportProjector, BrigsandRestraints, BurglarAlarm, HighSecurityAlarm, MutableLicensePlate, OilSprayer, PaintSprayer, SmokeScreen, SpikeDropper, VehicleBay, HangerBay, DryDock, SpaceDock, ExternalCradle, ArrestorHook, VehicularParachute, RefuellingProbe, RefuellingDrogue, FuelElectrolysisSystem, HydrogenFuelScoop, AtmosphereProcessor, NuclearDamper, SmallRealityStabilizer, MediumRealityStabilizer, HeavyRealityStabilizer, ModularSocket, Module

vbwProfiler.vbwExecuteLine 1486
                sOutput = sOutput + element.PrintOutput + " "

        End Select
vbwProfiler.vbwExecuteLine 1487 'B
vbwProfiler.vbwExecuteLine 1488
    Next

vbwProfiler.vbwExecuteLine 1489
    GetMiscellaneousOutput = sOutput
vbwProfiler.vbwProcOut 59
vbwProfiler.vbwExecuteLine 1490
    Exit Function
err:
vbwProfiler.vbwExecuteLine 1491
    Debug.Print "modTextOutput:GetMiscellaneousOutput -- Error #" & err.Number & " " & err.Description

vbwProfiler.vbwProcOut 59
vbwProfiler.vbwExecuteLine 1492
End Function

Private Function GetVehicleControlsOutput() As String
vbwProfiler.vbwProcIn 60
    Dim sOutput As String
    Dim element As Object

vbwProfiler.vbwExecuteLine 1493
    On Error GoTo err
vbwProfiler.vbwExecuteLine 1494
    For Each element In m_oCurrentVeh.Components
vbwProfiler.vbwExecuteLine 1495
        Select Case element.Datatype

'vbwLine 1496:            Case PrimitiveManeuverControl, ElectronicDivingControl, ComputerizedDivingControl, MechanicalManeuverControl, ElectronicManeuverControl, ComputerizedManeuverControl, MechanicalDivingControl
            Case IIf(vbwProfiler.vbwExecuteLine(1496), VBWPROFILER_EMPTY, _
        PrimitiveManeuverControl), ElectronicDivingControl, ComputerizedDivingControl, MechanicalManeuverControl, ElectronicManeuverControl, ComputerizedManeuverControl, MechanicalDivingControl

vbwProfiler.vbwExecuteLine 1497
                sOutput = sOutput + element.PrintOutput + " "

        End Select
vbwProfiler.vbwExecuteLine 1498 'B
vbwProfiler.vbwExecuteLine 1499
    Next

vbwProfiler.vbwExecuteLine 1500
    GetVehicleControlsOutput = sOutput
vbwProfiler.vbwProcOut 60
vbwProfiler.vbwExecuteLine 1501
    Exit Function
err:
vbwProfiler.vbwExecuteLine 1502
    Debug.Print "modTextOutput:GetVehicleControlsOutput -- Error #" & err.Number & " " & err.Description

vbwProfiler.vbwProcOut 60
vbwProfiler.vbwExecuteLine 1503
End Function

Private Function GetCrewStationsOutput() As String
vbwProfiler.vbwProcIn 61
    Dim sOutput As String
    Dim element As Object

vbwProfiler.vbwExecuteLine 1504
    On Error GoTo err
vbwProfiler.vbwExecuteLine 1505
    For Each element In m_oCurrentVeh.Components
vbwProfiler.vbwExecuteLine 1506
        Select Case element.Datatype

'vbwLine 1507:            Case CrampedCrewStation, NormalCrewStation, RoomyCrewStation, CycleCrewStation, HarnessCrewStation
            Case IIf(vbwProfiler.vbwExecuteLine(1507), VBWPROFILER_EMPTY, _
        CrampedCrewStation), NormalCrewStation, RoomyCrewStation, CycleCrewStation, HarnessCrewStation


vbwProfiler.vbwExecuteLine 1508
                sOutput = sOutput + element.PrintOutput + " "

        End Select
vbwProfiler.vbwExecuteLine 1509 'B
vbwProfiler.vbwExecuteLine 1510
    Next

vbwProfiler.vbwExecuteLine 1511
    GetCrewStationsOutput = sOutput
vbwProfiler.vbwProcOut 61
vbwProfiler.vbwExecuteLine 1512
    Exit Function
err:
vbwProfiler.vbwExecuteLine 1513
    Debug.Print "modTextOutput:GetCrewStationsOutput -- Error #" & err.Number & " " & err.Description

vbwProfiler.vbwProcOut 61
vbwProfiler.vbwExecuteLine 1514
End Function

Private Function GetOccupancyOutput() As String
vbwProfiler.vbwProcIn 62

    Dim sOutput As String
    Dim element As Object

vbwProfiler.vbwExecuteLine 1515
    On Error GoTo err
vbwProfiler.vbwExecuteLine 1516
    With m_oCurrentVeh.crew
vbwProfiler.vbwExecuteLine 1517
         If .numshifts > 1 Then
vbwProfiler.vbwExecuteLine 1518
              sOutput = NumericToString(.numshifts) & " shifts. "
         End If
vbwProfiler.vbwExecuteLine 1519 'B
vbwProfiler.vbwExecuteLine 1520
        If .numcaptains > 0 Then
vbwProfiler.vbwExecuteLine 1521
             sOutput = sOutput & NumericToString(.numcaptains) & " captains. "
        End If
vbwProfiler.vbwExecuteLine 1522 'B
vbwProfiler.vbwExecuteLine 1523
        If .NumOfficers > 0 Then
vbwProfiler.vbwExecuteLine 1524
             sOutput = sOutput & NumericToString(.NumOfficers) & " officers. "
        End If
vbwProfiler.vbwExecuteLine 1525 'B
vbwProfiler.vbwExecuteLine 1526
        If .NumCrewStationOperators > 0 Then
vbwProfiler.vbwExecuteLine 1527
             sOutput = sOutput & NumericToString(.NumCrewStationOperators) & " crew station operators. "
        End If
vbwProfiler.vbwExecuteLine 1528 'B
vbwProfiler.vbwExecuteLine 1529
        If .NumWeaponLoaders > 0 Then
vbwProfiler.vbwExecuteLine 1530
             sOutput = sOutput & NumericToString(.NumWeaponLoaders) & " weapon loaders. "
        End If
vbwProfiler.vbwExecuteLine 1531 'B
vbwProfiler.vbwExecuteLine 1532
        If .NumRowers > 0 Then
vbwProfiler.vbwExecuteLine 1533
             sOutput = sOutput & NumericToString(.NumRowers) & " rowers. "
        End If
vbwProfiler.vbwExecuteLine 1534 'B
vbwProfiler.vbwExecuteLine 1535
        If .NumSailors > 0 Then
vbwProfiler.vbwExecuteLine 1536
             sOutput = sOutput & NumericToString(.NumSailors) & " sailors. "
        End If
vbwProfiler.vbwExecuteLine 1537 'B
vbwProfiler.vbwExecuteLine 1538
        If .NumRiggers > 0 Then
vbwProfiler.vbwExecuteLine 1539
             sOutput = sOutput & NumericToString(.NumRiggers) & " sail riggers. "
        End If
vbwProfiler.vbwExecuteLine 1540 'B
vbwProfiler.vbwExecuteLine 1541
        If .NumFuelStokers > 0 Then
vbwProfiler.vbwExecuteLine 1542
             sOutput = sOutput & NumericToString(.NumFuelStokers) & " fuel stokers. "
        End If
vbwProfiler.vbwExecuteLine 1543 'B
vbwProfiler.vbwExecuteLine 1544
        If .NumMechanics > 0 Then
vbwProfiler.vbwExecuteLine 1545
             sOutput = sOutput & NumericToString(.NumMechanics) & " mechanics. "
        End If
vbwProfiler.vbwExecuteLine 1546 'B
vbwProfiler.vbwExecuteLine 1547
        If .NumServiceCrewmen > 0 Then
vbwProfiler.vbwExecuteLine 1548
             sOutput = sOutput & NumericToString(.NumServiceCrewmen) & " service crewmen. "
        End If
vbwProfiler.vbwExecuteLine 1549 'B
vbwProfiler.vbwExecuteLine 1550
        If .NumMedics > 0 Then
vbwProfiler.vbwExecuteLine 1551
             sOutput = sOutput & NumericToString(.NumMedics) & " medics. "
        End If
vbwProfiler.vbwExecuteLine 1552 'B
vbwProfiler.vbwExecuteLine 1553
        If .NumScientists > 0 Then
vbwProfiler.vbwExecuteLine 1554
             sOutput = sOutput & NumericToString(.NumScientists) & " scientists. "
        End If
vbwProfiler.vbwExecuteLine 1555 'B
vbwProfiler.vbwExecuteLine 1556
        If .NumAuxiliaryVehicleCrew > 0 Then
vbwProfiler.vbwExecuteLine 1557
             sOutput = sOutput & NumericToString(.NumAuxiliaryVehicleCrew) & " auxiliary vehicle crewmen. "
        End If
vbwProfiler.vbwExecuteLine 1558 'B
vbwProfiler.vbwExecuteLine 1559
        If .NumStewards > 0 Then
vbwProfiler.vbwExecuteLine 1560
             sOutput = sOutput & NumericToString(.NumStewards) & " stewards. "
        End If
vbwProfiler.vbwExecuteLine 1561 'B
vbwProfiler.vbwExecuteLine 1562
        If .NumLuxury > 0 Then
vbwProfiler.vbwExecuteLine 1563
             sOutput = sOutput & NumericToString(.NumLuxury) & " luxury class passengers. "
        End If
vbwProfiler.vbwExecuteLine 1564 'B
vbwProfiler.vbwExecuteLine 1565
        If .NumFirstClass > 0 Then
vbwProfiler.vbwExecuteLine 1566
             sOutput = sOutput & NumericToString(.NumFirstClass) & " first class passengers. "
        End If
vbwProfiler.vbwExecuteLine 1567 'B
vbwProfiler.vbwExecuteLine 1568
        If .NumSecondClass > 0 Then
vbwProfiler.vbwExecuteLine 1569
             sOutput = sOutput & NumericToString(.NumSecondClass) & " second class passengers. "
        End If
vbwProfiler.vbwExecuteLine 1570 'B
vbwProfiler.vbwExecuteLine 1571
        If .NumSteerage > 0 Then
vbwProfiler.vbwExecuteLine 1572
             sOutput = sOutput & NumericToString(.NumSteerage) & " steerage passengers. "
        End If
vbwProfiler.vbwExecuteLine 1573 'B

        ' append whether its short or long Occupancy
vbwProfiler.vbwExecuteLine 1574
        sOutput = .Occupancy & ". " & sOutput
vbwProfiler.vbwExecuteLine 1575
   End With

vbwProfiler.vbwExecuteLine 1576
    GetOccupancyOutput = sOutput
vbwProfiler.vbwProcOut 62
vbwProfiler.vbwExecuteLine 1577
    Exit Function
err:
vbwProfiler.vbwExecuteLine 1578
    Debug.Print "modTextOutput:GetOccupancyOutput -- Error #" & err.Number & " " & err.Description

vbwProfiler.vbwProcOut 62
vbwProfiler.vbwExecuteLine 1579
End Function

Private Function GetAccomodationsOutput() As String
vbwProfiler.vbwProcIn 63
    Dim sOutput As String
    Dim element As Object

vbwProfiler.vbwExecuteLine 1580
    On Error GoTo err
vbwProfiler.vbwExecuteLine 1581
    For Each element In m_oCurrentVeh.Components
vbwProfiler.vbwExecuteLine 1582
        Select Case element.Datatype

'vbwLine 1583:            Case CrampedSeat, NormalSeat, RoomySeat, CrampedStandingRoom, NormalStandingRoom, RoomyStandingRoom, CycleSeat, Hammock, Bunk, Cabin, LuxuryCabin, Suite, LuxurySuite, SmallGalley
            Case IIf(vbwProfiler.vbwExecuteLine(1583), VBWPROFILER_EMPTY, _
        CrampedSeat), NormalSeat, RoomySeat, CrampedStandingRoom, NormalStandingRoom, RoomyStandingRoom, CycleSeat, Hammock, Bunk, Cabin, LuxuryCabin, Suite, LuxurySuite, SmallGalley


vbwProfiler.vbwExecuteLine 1584
                sOutput = sOutput + element.PrintOutput + " "

        End Select
vbwProfiler.vbwExecuteLine 1585 'B
vbwProfiler.vbwExecuteLine 1586
    Next

vbwProfiler.vbwExecuteLine 1587
    GetAccomodationsOutput = sOutput
vbwProfiler.vbwProcOut 63
vbwProfiler.vbwExecuteLine 1588
    Exit Function
err:
vbwProfiler.vbwExecuteLine 1589
    Debug.Print "modTextOutput:GetAccomodationsOutput -- Error #" & err.Number & " " & err.Description

vbwProfiler.vbwProcOut 63
vbwProfiler.vbwExecuteLine 1590
End Function

Private Function GetEnvironmentalSystemsOutput() As String
vbwProfiler.vbwProcIn 64
    Dim sOutput As String
    Dim element As Object

vbwProfiler.vbwExecuteLine 1591
    On Error GoTo err
vbwProfiler.vbwExecuteLine 1592
    For Each element In m_oCurrentVeh.Components
vbwProfiler.vbwExecuteLine 1593
        Select Case element.Datatype

'vbwLine 1594:            Case TotalLifeSystem, ArtificialGravityUnit, EnvironmentalControl, NBCKit, LimitedLifeSystem, FullLifeSystem
            Case IIf(vbwProfiler.vbwExecuteLine(1594), VBWPROFILER_EMPTY, _
        TotalLifeSystem), ArtificialGravityUnit, EnvironmentalControl, NBCKit, LimitedLifeSystem, FullLifeSystem

vbwProfiler.vbwExecuteLine 1595
                sOutput = sOutput + element.PrintOutput + " "
        End Select
vbwProfiler.vbwExecuteLine 1596 'B
vbwProfiler.vbwExecuteLine 1597
    Next

    'append the provisions to it
vbwProfiler.vbwExecuteLine 1598
    For Each element In m_oCurrentVeh.Components
vbwProfiler.vbwExecuteLine 1599
        If element.Datatype = Provisions Then
vbwProfiler.vbwExecuteLine 1600
            sOutput = sOutput + element.PrintOutput + " "
        End If
vbwProfiler.vbwExecuteLine 1601 'B
vbwProfiler.vbwExecuteLine 1602
    Next
vbwProfiler.vbwExecuteLine 1603
    GetEnvironmentalSystemsOutput = sOutput
vbwProfiler.vbwProcOut 64
vbwProfiler.vbwExecuteLine 1604
    Exit Function
err:
vbwProfiler.vbwExecuteLine 1605
    Debug.Print "modTextOutput:EnvironmentalSystemsOutput -- Error #" & err.Number & " " & err.Description
vbwProfiler.vbwProcOut 64
vbwProfiler.vbwExecuteLine 1606
End Function

Private Function GetSafetySystemsOutput() As String
vbwProfiler.vbwProcIn 65
    Dim sOutput As String
    Dim element As Object

vbwProfiler.vbwExecuteLine 1607
    On Error GoTo err
vbwProfiler.vbwExecuteLine 1608
    For Each element In m_oCurrentVeh.Components
vbwProfiler.vbwExecuteLine 1609
        Select Case element.Datatype
'vbwLine 1610:            Case EjectionSeat, CrewEscapeCapsule, Airbag, CrashWeb, WombTank, GravityWeb, GravCompensator
            Case IIf(vbwProfiler.vbwExecuteLine(1610), VBWPROFILER_EMPTY, _
        EjectionSeat), CrewEscapeCapsule, Airbag, CrashWeb, WombTank, GravityWeb, GravCompensator

vbwProfiler.vbwExecuteLine 1611
                sOutput = sOutput + element.PrintOutput + " "
        End Select
vbwProfiler.vbwExecuteLine 1612 'B
vbwProfiler.vbwExecuteLine 1613
    Next
vbwProfiler.vbwExecuteLine 1614
    GetSafetySystemsOutput = sOutput
vbwProfiler.vbwProcOut 65
vbwProfiler.vbwExecuteLine 1615
    Exit Function
err:
vbwProfiler.vbwExecuteLine 1616
    Debug.Print "modTextOutput:GetSafetySystemsOutput -- Error #" & err.Number & " " & err.Description

vbwProfiler.vbwProcOut 65
vbwProfiler.vbwExecuteLine 1617
End Function

Public Function GetPowerSystemsOutPut() As String
vbwProfiler.vbwProcIn 66
    Dim oProfile As clsProfilePower
    Dim oGroup As clsSupplyConsumeGroup
    Dim iGroupCount As Long
    Dim i As Long
    Dim iSupplierCount As Long
    Dim iConsumerCount As Long
    Dim sTemp As String
    Dim j As Long

vbwProfiler.vbwExecuteLine 1618
    On Error GoTo err
vbwProfiler.vbwExecuteLine 1619
    For Each oProfile In m_oCurrentVeh.Profiles
vbwProfiler.vbwExecuteLine 1620
        sTemp = sTemp & "Profile " & oProfile.Key & vbNewLine
vbwProfiler.vbwExecuteLine 1621
        iGroupCount = oProfile.groupcount
vbwProfiler.vbwExecuteLine 1622
        If iGroupCount > 0 Then
vbwProfiler.vbwExecuteLine 1623
            For i = 1 To iGroupCount
vbwProfiler.vbwExecuteLine 1624
                Set oGroup = oProfile.Group(i)
                ' get the suppliers
vbwProfiler.vbwExecuteLine 1625
                sTemp = sTemp & " Suppliers " & vbNewLine
vbwProfiler.vbwExecuteLine 1626
                iSupplierCount = oGroup.SupplierCount
vbwProfiler.vbwExecuteLine 1627
                For j = 1 To iSupplierCount
vbwProfiler.vbwExecuteLine 1628
                    sTemp = sTemp & m_oCurrentVeh.Components(oGroup.Supplier(j)).Description
vbwProfiler.vbwExecuteLine 1629
                Next
                ' get the consumers
vbwProfiler.vbwExecuteLine 1630
                sTemp = sTemp & " Consumers " & vbNewLine
vbwProfiler.vbwExecuteLine 1631
                iConsumerCount = oGroup.ConsumerCount
vbwProfiler.vbwExecuteLine 1632
                For j = 1 To iConsumerCount
vbwProfiler.vbwExecuteLine 1633
                    sTemp = sTemp & m_oCurrentVeh.Components(oGroup.consumer(j)).Description
vbwProfiler.vbwExecuteLine 1634
                Next
vbwProfiler.vbwExecuteLine 1635
                sTemp = sTemp
vbwProfiler.vbwExecuteLine 1636
            Next
       End If
vbwProfiler.vbwExecuteLine 1637 'B

vbwProfiler.vbwExecuteLine 1638
        GetPowerSystemsOutPut = sTemp
        ' get the keys for each supplier in each group


        ' get the keys for all consumers attached to each group


vbwProfiler.vbwExecuteLine 1639
    Next ' next profile
vbwProfiler.vbwExecuteLine 1640
    Set oGroup = Nothing
vbwProfiler.vbwExecuteLine 1641
    Set oProfile = Nothing
vbwProfiler.vbwProcOut 66
vbwProfiler.vbwExecuteLine 1642
    Exit Function
err:
vbwProfiler.vbwExecuteLine 1643
    Debug.Print "modTextOutput:GetPowerSystemsOutput -- Error #" & err.Number & " " & err.Description

vbwProfiler.vbwProcOut 66
vbwProfiler.vbwExecuteLine 1644
End Function
'Private Function GetPowerSystemsOutPut() As String
'  MPJ 05/27/02 ENTIRE FUNCTION OBSOLETE and NON FUNCTIONAL WITH NEW POWER SYSTEM PROFILES
'    Dim sOutput As String
'    Dim element As Object
'    Dim sKeyArray() As String
'    Dim i As Long
'    Dim sngPowerRemaining As Single
'
'    'todo: this is simplified because we can get this info directly from the
'    ' m_SC_Groups in each profile
'
'    ' todo: however, we do need seperate writes ups for each profile
'
'    For Each element In m_oCurrentVeh.Components
'        Select Case element.Datatype
'
'            Case MuscleEngine, GasolineEngine, HPGasolineEngine, _
'                TurboGasolineEngine, SuperGasolineEngine, TurboHPGasolineEngine, _
'                SuperHPGasolineEngine, StandardDieselEngine, _
'                TurboStandardDieselEngine, MarineDieselEngine, _
'                HPDieselEngine, TurboHPDieselEngine, CeramicEngine, _
'                TurboCeramicEngine, SuperCeramicEngine, HPCeramicEngine, _
'                TurboHPCeramicEngine, SuperHPCeramicEngine, _
'                HydrogenCombustionEngine, EarlySteamEngine, _
'                ForcedDraftSteamEngine, TripleExpansionSteamEngine, SteamTurbine
'
'                sOutput = sOutput + element.PrintOutput + " "
'
'                '//append the systems that it powers
'                sKeyArray = element.GetCurrentConsumptionSystemKeys
'                If sKeyArray(1) = "" Then
'                Else
'                    sOutput = sOutput + "Powers the "
'                    For i = 1 To UBound(sKeyArray)
'                        sOutput = sOutput + m_oCurrentVeh.Components(sKeyArray(i)).Description & ", "
'                    Next
'                    '//delete the last "," and replace it with "."
'                    sOutput = Left(sOutput, Len(sOutput) - 2)
'                    sngPowerRemaining = element.Output - element.PowerConsumed
'                    If sngPowerRemaining > 0 Then
'                        sOutput = sOutput + " with " & sngPowerRemaining & " kW in reserve."
'                    Else
'                        sOutput = sOutput + ".  "
'                    End If
'                End If
'
'            Case StandardGasTurbine, HPGasTurbine, OptimizedGasTurbine, _
'                StandardMHDTurbine, HPMHDTurbine, FuelCell, FissionReactor, _
'                RTGReactor, NPU, FusionReactor, AntimatterReactor, _
'                TotalConversionPowerPlant, CosmicPowerPlant, Soulburner, _
'                ElementalFurnace, ManaEngine, Carnivore, Herbivore, Omnivore, _
'                Vampire, ClockWork, LeadAcidBattery, AdvancedBattery, _
'                Flywheel, RechargeablePowerCell, PowerCell, Snorkel, _
'                ElectricContactPower, LaserBeamedPowerReceiver, _
'                MaserBeamedPowerReceiver, SolarCellArray
'
'                sOutput = sOutput + element.PrintOutput + " "
'
'                '//append the systems that it powers
'                sKeyArray = element.GetCurrentConsumptionSystemKeys
'                If sKeyArray(1) = "" Then
'                Else
'                    sOutput = sOutput + "Powers the "
'                    For i = 1 To UBound(sKeyArray)
'                        sOutput = sOutput + m_oCurrentVeh.Components(sKeyArray(i)).Description & ", "
'                    Next
'                    '//delete the last "," and replace it with "."
'                    sOutput = Left(sOutput, Len(sOutput) - 2)
'                    sngPowerRemaining = element.Output - element.PowerConsumed
'                    If sngPowerRemaining > 0 Then
'                        sOutput = sOutput + " with " & sngPowerRemaining & " kW in reserve."
'                    Else
'                        sOutput = sOutput + ".  "
'                    End If
'                End If
'        End Select
'
'    Next
'
'
'
'    GetPowerSystemsOutPut = sOutput
'End Function

Private Function GetFuelOutput() As String
vbwProfiler.vbwProcIn 67
    Dim sOutput As String
    Dim element As Object

    'TODO: Is this where endurance gets spit out?
    '
' 'find the endurance of the engine
'    mvarEndurance = 0 'reset the variable
'    If mvarFuelStorageKeyChain(1) <> "" Then
'        For i = 1 To UBound(mvarFuelStorageKeyChain)
'            mvarEndurance = mvarEndurance + Veh.Components(mvarFuelStorageKeyChain(i)).capacity
'        Next
'        mvarEndurance = mvarEndurance / mvarFuelConsumption
'    Else
'        mvarEndurance = 0
'    End If

vbwProfiler.vbwExecuteLine 1645
    On Error GoTo err
vbwProfiler.vbwExecuteLine 1646
    For Each element In m_oCurrentVeh.Components
vbwProfiler.vbwExecuteLine 1647
        Select Case element.Datatype
'vbwLine 1648:            Case AntiMatterBay, CoalBunker, WoodBunker, StandardTank, lightTank, UltralightTank, StandardSelfSealingTank, lightSelfSealingTank, UltralightSelfSealingTank
            Case IIf(vbwProfiler.vbwExecuteLine(1648), VBWPROFILER_EMPTY, _
        AntiMatterBay), CoalBunker, WoodBunker, StandardTank, lightTank, UltralightTank, StandardSelfSealingTank, lightSelfSealingTank, UltralightSelfSealingTank

vbwProfiler.vbwExecuteLine 1649
                sOutput = sOutput + element.PrintOutput + " "
        End Select
vbwProfiler.vbwExecuteLine 1650 'B
vbwProfiler.vbwExecuteLine 1651
    Next
vbwProfiler.vbwExecuteLine 1652
    GetFuelOutput = sOutput
vbwProfiler.vbwProcOut 67
vbwProfiler.vbwExecuteLine 1653
    Exit Function
err:
vbwProfiler.vbwExecuteLine 1654
    Debug.Print "modTextOutput:GetFuelOutput -- Error #" & err.Number & " " & err.Description

vbwProfiler.vbwProcOut 67
vbwProfiler.vbwExecuteLine 1655
End Function

Private Function GetSurfaceAreaOutput() As String
vbwProfiler.vbwProcIn 68
    Dim sOutput As String
    Dim element As Object
    Dim totalsurfacearea As Single

vbwProfiler.vbwExecuteLine 1656
    On Error GoTo err
vbwProfiler.vbwExecuteLine 1657
    For Each element In m_oCurrentVeh.Components
vbwProfiler.vbwExecuteLine 1658
        Select Case element.Datatype
'vbwLine 1659:            Case Body, Wheel, Skid, Track, Hydrofoil, Hovercraft, Leg, Arm, AutogyroRotor, TTRotor, CARotor, MMRotor, Wing, Mast, Superstructure, Turret, Popturret, OpenMount, Gasbag, Pod
            Case IIf(vbwProfiler.vbwExecuteLine(1659), VBWPROFILER_EMPTY, _
        Body), Wheel, Skid, Track, Hydrofoil, Hovercraft, Leg, Arm, AutogyroRotor, TTRotor, CARotor, MMRotor, Wing, Mast, Superstructure, Turret, Popturret, OpenMount, Gasbag, Pod

vbwProfiler.vbwExecuteLine 1660
                totalsurfacearea = totalsurfacearea + element.SurfaceArea
vbwProfiler.vbwExecuteLine 1661
                sOutput = sOutput + element.abbrev + " " + Format(element.SurfaceArea, Settings.FormatString) + ". "
        End Select
vbwProfiler.vbwExecuteLine 1662 'B
vbwProfiler.vbwExecuteLine 1663
    Next
vbwProfiler.vbwExecuteLine 1664
    GetSurfaceAreaOutput = sOutput + "total " + Format(totalsurfacearea, Settings.FormatString) + "."
vbwProfiler.vbwProcOut 68
vbwProfiler.vbwExecuteLine 1665
    Exit Function
err:
vbwProfiler.vbwExecuteLine 1666
    Debug.Print "modTextOutput:GetSurfaceAreaOutput -- Error #" & err.Number & " " & err.Description
vbwProfiler.vbwProcOut 68
vbwProfiler.vbwExecuteLine 1667
End Function

Private Function GetStructureOutput() As String
vbwProfiler.vbwProcIn 69
    Dim sOutput As String
    Dim element As Object
    Dim sBodyStruct As String
    Dim tOutput As String

vbwProfiler.vbwExecuteLine 1668
On Error GoTo err
    'get the structure of the body first

vbwProfiler.vbwExecuteLine 1669
    With m_oCurrentVeh.Components(BODY_KEY)
vbwProfiler.vbwExecuteLine 1670
        sBodyStruct = element.Description + " - " + .FrameStrength + " frame" + " with " + .Materials + " materials. "
vbwProfiler.vbwExecuteLine 1671
    End With

vbwProfiler.vbwExecuteLine 1672
    sOutput = sBodyStruct

vbwProfiler.vbwExecuteLine 1673
    For Each element In m_oCurrentVeh.Components
vbwProfiler.vbwExecuteLine 1674
        Select Case element.Datatype
            'note Open Mount, Mast and Gasbag are not included here
'vbwLine 1675:            Case Body, Wheel, Skid, Track, Hydrofoil, Hovercraft, Leg, Arm, AutogyroRotor, TTRotor, CARotor, MMRotor, Wing, Superstructure, Turret, Popturret, Pod
            Case IIf(vbwProfiler.vbwExecuteLine(1675), VBWPROFILER_EMPTY, _
        Body), Wheel, Skid, Track, Hydrofoil, Hovercraft, Leg, Arm, AutogyroRotor, TTRotor, CARotor, MMRotor, Wing, Superstructure, Turret, Popturret, Pod

vbwProfiler.vbwExecuteLine 1676
                With element
vbwProfiler.vbwExecuteLine 1677
                    tOutput = element.abbrev + " - " + .FrameStrength + " frame" + " with " + .Materials + " materials. "
vbwProfiler.vbwExecuteLine 1678
                End With

                'only print this if its different than the Body's structure
vbwProfiler.vbwExecuteLine 1679
                If tOutput <> sBodyStruct Then
vbwProfiler.vbwExecuteLine 1680
                    sOutput = sOutput + " " + tOutput
                End If
vbwProfiler.vbwExecuteLine 1681 'B
        End Select
vbwProfiler.vbwExecuteLine 1682 'B
vbwProfiler.vbwExecuteLine 1683
    Next

vbwProfiler.vbwExecuteLine 1684
    GetStructureOutput = sOutput
vbwProfiler.vbwProcOut 69
vbwProfiler.vbwExecuteLine 1685
    Exit Function
err:
vbwProfiler.vbwExecuteLine 1686
    Debug.Print "modTextOutput:GetStructureOutput -- Error #" & err.Number & " " & err.Description
vbwProfiler.vbwExecuteLine 1687
    Resume Next
vbwProfiler.vbwProcOut 69
vbwProfiler.vbwExecuteLine 1688
End Function
Private Function GetHitPointsOutput() As String
vbwProfiler.vbwProcIn 70
    Dim sOutput As String
    Dim element As Object

vbwProfiler.vbwExecuteLine 1689
    On Error GoTo err
vbwProfiler.vbwExecuteLine 1690
    For Each element In m_oCurrentVeh.Components
vbwProfiler.vbwExecuteLine 1691
        Select Case element.Datatype
'vbwLine 1692:            Case Body, Wheel, Skid, Track, Hydrofoil, Hovercraft, Leg, Arm, AutogyroRotor, TTRotor, CARotor, MMRotor, Wing, Mast, Superstructure, Turret, Popturret, OpenMount, Gasbag, Pod, equipmentPod, SolarPanel
            Case IIf(vbwProfiler.vbwExecuteLine(1692), VBWPROFILER_EMPTY, _
        Body), Wheel, Skid, Track, Hydrofoil, Hovercraft, Leg, Arm, AutogyroRotor, TTRotor, CARotor, MMRotor, Wing, Mast, Superstructure, Turret, Popturret, OpenMount, Gasbag, Pod, equipmentPod, SolarPanel

vbwProfiler.vbwExecuteLine 1693
                sOutput = sOutput + element.abbrev + " " + Format(element.HitPoints) + ", "
        End Select
vbwProfiler.vbwExecuteLine 1694 'B
vbwProfiler.vbwExecuteLine 1695
    Next
vbwProfiler.vbwExecuteLine 1696
    GetHitPointsOutput = Left(sOutput, Len(sOutput) - 2) + "."
vbwProfiler.vbwProcOut 70
vbwProfiler.vbwExecuteLine 1697
    Exit Function
err:
vbwProfiler.vbwExecuteLine 1698
    Debug.Print "modTextOutput:GetHitPointsOutput -- Error #" & err.Number & " " & err.Description

vbwProfiler.vbwProcOut 70
vbwProfiler.vbwExecuteLine 1699
End Function

Private Function GetArmorOutput() As String
vbwProfiler.vbwProcIn 71
    Dim sOutput As String
    Dim element As Object

vbwProfiler.vbwExecuteLine 1700
    On Error GoTo err
vbwProfiler.vbwExecuteLine 1701
    For Each element In m_oCurrentVeh.Components
vbwProfiler.vbwExecuteLine 1702
        Select Case element.Datatype
'vbwLine 1703:            Case ArmorComplexFacing, ArmorBasicFacing, ArmorOpenFrame, ArmorGunShield, ArmorLocation, ArmorComponent, ArmorOverall, ArmorWheelGuard
            Case IIf(vbwProfiler.vbwExecuteLine(1703), VBWPROFILER_EMPTY, _
        ArmorComplexFacing), ArmorBasicFacing, ArmorOpenFrame, ArmorGunShield, ArmorLocation, ArmorComponent, ArmorOverall, ArmorWheelGuard

vbwProfiler.vbwExecuteLine 1704
                 sOutput = sOutput + m_oCurrentVeh.Components(element.LogicalParent).CustomDescription + " armor: " + element.PrintOutput + vbNewLine
        End Select
vbwProfiler.vbwExecuteLine 1705 'B
vbwProfiler.vbwExecuteLine 1706
    Next
vbwProfiler.vbwExecuteLine 1707
    GetArmorOutput = sOutput
vbwProfiler.vbwProcOut 71
vbwProfiler.vbwExecuteLine 1708
    Exit Function
err:
vbwProfiler.vbwExecuteLine 1709
    Debug.Print "modTextOutput:GetArmorOutput -- Error #" & err.Number & " " & err.Description

vbwProfiler.vbwProcOut 71
vbwProfiler.vbwExecuteLine 1710
End Function

Private Function GetStatisticsOutput() As String
vbwProfiler.vbwProcIn 72
    Dim sTemp As String

vbwProfiler.vbwExecuteLine 1711
    On Error GoTo err
vbwProfiler.vbwExecuteLine 1712
    With m_oCurrentVeh.Stats
vbwProfiler.vbwExecuteLine 1713
        sTemp = "Empty weight " + Format(.EmptyWeight, Settings.FormatString) + " lbs., "
vbwProfiler.vbwExecuteLine 1714
        If .UsualInternalPayload <> 0 Then
vbwProfiler.vbwExecuteLine 1715
             sTemp = sTemp + "Internal payload " + Format(.UsualInternalPayload, Settings.FormatString) + " lbs., "
        End If
vbwProfiler.vbwExecuteLine 1716 'B
vbwProfiler.vbwExecuteLine 1717
        sTemp = sTemp + "Loaded weight " + Format(.LoadedWeight, Settings.FormatString) + " lbs., "
vbwProfiler.vbwExecuteLine 1718
        If .SubmergedWeight <> 0 Then
vbwProfiler.vbwExecuteLine 1719
             sTemp = sTemp + "Submerged weight " + Format(.SubmergedWeight, Settings.FormatString) + " lbs., "
        End If
vbwProfiler.vbwExecuteLine 1720 'B
vbwProfiler.vbwExecuteLine 1721
        sTemp = sTemp + "Volume " + Format(.TotalVolume, Settings.FormatString) + " cf. "
vbwProfiler.vbwExecuteLine 1722
        sTemp = sTemp + "Size modifier " + Format(.SizeModifier) + ". "
vbwProfiler.vbwExecuteLine 1723
        sTemp = sTemp + "Cost $" + Format(.TotalPrice, Settings.FormatString) + ". "
vbwProfiler.vbwExecuteLine 1724
        sTemp = sTemp + "HT " + Format(.StructuralHealth)
vbwProfiler.vbwExecuteLine 1725
    End With
vbwProfiler.vbwExecuteLine 1726
    GetStatisticsOutput = sTemp
vbwProfiler.vbwProcOut 72
vbwProfiler.vbwExecuteLine 1727
    Exit Function
err:
vbwProfiler.vbwExecuteLine 1728
    Debug.Print "modTextOutput:GetStatisticsOutput -- Error #" & err.Number & " " & err.Description

vbwProfiler.vbwProcOut 72
vbwProfiler.vbwExecuteLine 1729
End Function

Private Function GetSpaceOutput() As String
   'access, empty and cargo space
vbwProfiler.vbwProcIn 73
    Dim sAccessOutput As String
    Dim sEmptyOutput As String
    Dim sCargoOutput As String
    Dim element As Object
    Dim sOutput As String

vbwProfiler.vbwExecuteLine 1730
    On Error GoTo err

vbwProfiler.vbwExecuteLine 1731
    For Each element In m_oCurrentVeh.Components
vbwProfiler.vbwExecuteLine 1732
        Select Case element.Datatype

'vbwLine 1733:            Case Body, Wheel, Skid, Track, Hydrofoil, Hovercraft, Leg, Arm, AutogyroRotor, TTRotor, CARotor, MMRotor, Wing, Mast, Superstructure, Turret, Popturret, OpenMount, Gasbag, Pod, SolarPanel, equipmentPod
            Case IIf(vbwProfiler.vbwExecuteLine(1733), VBWPROFILER_EMPTY, _
        Body), Wheel, Skid, Track, Hydrofoil, Hovercraft, Leg, Arm, AutogyroRotor, TTRotor, CARotor, MMRotor, Wing, Mast, Superstructure, Turret, Popturret, OpenMount, Gasbag, Pod, SolarPanel, equipmentPod

vbwProfiler.vbwExecuteLine 1734
                If element.EmptySpace <> 0 Then
vbwProfiler.vbwExecuteLine 1735
                    sEmptyOutput = sEmptyOutput + element.abbrev + " " + Format(element.EmptySpace, Settings.FormatString) + " cf, "
                End If
vbwProfiler.vbwExecuteLine 1736 'B
vbwProfiler.vbwExecuteLine 1737
                If element.AccessSpace <> 0 Then
vbwProfiler.vbwExecuteLine 1738
                    sAccessOutput = sAccessOutput + element.abbrev + " " + Format(element.AccessSpace, Settings.FormatString) + " cf, "
                End If
vbwProfiler.vbwExecuteLine 1739 'B
'vbwLine 1740:            Case Cargo
            Case IIf(vbwProfiler.vbwExecuteLine(1740), VBWPROFILER_EMPTY, _
        Cargo)
vbwProfiler.vbwExecuteLine 1741
                sCargoOutput = sCargoOutput + element.PrintOutput + " "

        End Select
vbwProfiler.vbwExecuteLine 1742 'B
vbwProfiler.vbwExecuteLine 1743
    Next
vbwProfiler.vbwExecuteLine 1744
    sCargoOutput = Left(sCargoOutput, Len(sCargoOutput) - 1)
vbwProfiler.vbwExecuteLine 1745
    If sCargoOutput <> "" Then
vbwProfiler.vbwExecuteLine 1746
       sOutput = "Space: " + sCargoOutput
    End If
vbwProfiler.vbwExecuteLine 1747 'B
vbwProfiler.vbwExecuteLine 1748
    If sAccessOutput <> "" Then
vbwProfiler.vbwExecuteLine 1749
        sAccessOutput = Left(sAccessOutput, Len(sAccessOutput) - 2)
vbwProfiler.vbwExecuteLine 1750
        sAccessOutput = "(" + sAccessOutput + ")"
vbwProfiler.vbwExecuteLine 1751
        If sOutput = "" Then
vbwProfiler.vbwExecuteLine 1752
             sOutput = "Space: "
        End If
vbwProfiler.vbwExecuteLine 1753 'B
vbwProfiler.vbwExecuteLine 1754
        sOutput = sOutput + " Access space " + sAccessOutput + "."
    End If
vbwProfiler.vbwExecuteLine 1755 'B
vbwProfiler.vbwExecuteLine 1756
    If sEmptyOutput <> "" Then
vbwProfiler.vbwExecuteLine 1757
        sEmptyOutput = Left(sEmptyOutput, Len(sEmptyOutput) - 2)
vbwProfiler.vbwExecuteLine 1758
        sEmptyOutput = "(" + sEmptyOutput + ")"
vbwProfiler.vbwExecuteLine 1759
        If sOutput = "" Then
vbwProfiler.vbwExecuteLine 1760
             sOutput = "Space: "
        End If
vbwProfiler.vbwExecuteLine 1761 'B
vbwProfiler.vbwExecuteLine 1762
        sOutput = sOutput + " Empty space " + sEmptyOutput + "."
    End If
vbwProfiler.vbwExecuteLine 1763 'B
vbwProfiler.vbwExecuteLine 1764
    GetSpaceOutput = sOutput
       'this will error if the component doesnt have Access or Emtpyspace properties.
        'so will just resume past the error
vbwProfiler.vbwProcOut 73
vbwProfiler.vbwExecuteLine 1765
        Exit Function
err:
vbwProfiler.vbwExecuteLine 1766
    Debug.Print "modTextOutput:GetSpaceOutput -- Error #" & err.Number & " " & err.Description
vbwProfiler.vbwExecuteLine 1767
    Resume Next
vbwProfiler.vbwProcOut 73
vbwProfiler.vbwExecuteLine 1768
End Function

Private Function GetStructuralOptionsOutput() As String
vbwProfiler.vbwProcIn 74

Dim element As Object
Dim sOutput As String
Dim bControlledInstability As Boolean
Dim bImprovedSuspension As Boolean

vbwProfiler.vbwExecuteLine 1769
On Error GoTo err
    '//get the structural options that are stored in the body
vbwProfiler.vbwExecuteLine 1770
    With m_oCurrentVeh.Options
vbwProfiler.vbwExecuteLine 1771
        If .RollStabilizers Then

        End If
vbwProfiler.vbwExecuteLine 1772 'B
vbwProfiler.vbwExecuteLine 1773
        If m_oCurrentVeh.surface.Submersible = True Then

        End If
vbwProfiler.vbwExecuteLine 1774 'B
vbwProfiler.vbwExecuteLine 1775
    End With
    '//now get the rest of the structural options from the various other subsassemblies
vbwProfiler.vbwExecuteLine 1776
    For Each element In m_oCurrentVeh.Components
vbwProfiler.vbwExecuteLine 1777
        Select Case element.Datatype
'vbwLine 1778:            Case Body, Superstructure, Popturret, Turret
            Case IIf(vbwProfiler.vbwExecuteLine(1778), VBWPROFILER_EMPTY, _
        Body), Superstructure, Popturret, Turret
vbwProfiler.vbwExecuteLine 1779
                If element.Compartmentalization <> "none" Then
                    'compartmentalization
vbwProfiler.vbwExecuteLine 1780
                    sOutput = sOutput & " " & element.Compartmentalization & " compartmentalization in " & element.abbrev & "."
                End If
vbwProfiler.vbwExecuteLine 1781 'B
'vbwLine 1782:            Case Wing
            Case IIf(vbwProfiler.vbwExecuteLine(1782), VBWPROFILER_EMPTY, _
        Wing)
                'folding wings or rotors
vbwProfiler.vbwExecuteLine 1783
                If element.Folding Then
vbwProfiler.vbwExecuteLine 1784
                    sOutput = sOutput & " Folding wings."
                End If
vbwProfiler.vbwExecuteLine 1785 'B
vbwProfiler.vbwExecuteLine 1786
                If element.VariableSweep <> "none" Then
vbwProfiler.vbwExecuteLine 1787
                    sOutput = sOutput & element.VariableSweep & " variable sweep wings."
                End If
vbwProfiler.vbwExecuteLine 1788 'B
                'controlled instability
vbwProfiler.vbwExecuteLine 1789
                If (element.ControlledInstability) And (bControlledInstability = False) Then
vbwProfiler.vbwExecuteLine 1790
                    sOutput = sOutput & " Controlled instability."
vbwProfiler.vbwExecuteLine 1791
                    bControlledInstability = True
                End If
vbwProfiler.vbwExecuteLine 1792 'B
'vbwLine 1793:            Case TTRotor, AutogyroRotor, MMRotor, CARotor
            Case IIf(vbwProfiler.vbwExecuteLine(1793), VBWPROFILER_EMPTY, _
        TTRotor), AutogyroRotor, MMRotor, CARotor
                'folding wings or rotors
vbwProfiler.vbwExecuteLine 1794
                If element.Folding Then
vbwProfiler.vbwExecuteLine 1795
                    sOutput = sOutput & " Folding rotors."
                End If
vbwProfiler.vbwExecuteLine 1796 'B
                'controlled instability
vbwProfiler.vbwExecuteLine 1797
                If (element.ControlledInstability) And (bControlledInstability = False) Then
vbwProfiler.vbwExecuteLine 1798
                    sOutput = sOutput & " Controlled instability."
vbwProfiler.vbwExecuteLine 1799
                    bControlledInstability = True
                End If
vbwProfiler.vbwExecuteLine 1800 'B
'vbwLine 1801:            Case Track, Skid, Leg
            Case IIf(vbwProfiler.vbwExecuteLine(1801), VBWPROFILER_EMPTY, _
        Track), Skid, Leg
vbwProfiler.vbwExecuteLine 1802
                If (element.ImprovedSuspension) And (bImprovedSuspension = False) Then
vbwProfiler.vbwExecuteLine 1803
                    sOutput = sOutput & " Improved Suspension."
vbwProfiler.vbwExecuteLine 1804
                    bImprovedSuspension = True
                End If
vbwProfiler.vbwExecuteLine 1805 'B

'vbwLine 1806:            Case Wheel
            Case IIf(vbwProfiler.vbwExecuteLine(1806), VBWPROFILER_EMPTY, _
        Wheel)
vbwProfiler.vbwExecuteLine 1807
                If (element.ImprovedSuspension) And (bImprovedSuspension = False) Then
vbwProfiler.vbwExecuteLine 1808
                    sOutput = sOutput & " Improved Suspension."
vbwProfiler.vbwExecuteLine 1809
                    bImprovedSuspension = True
                End If
vbwProfiler.vbwExecuteLine 1810 'B
vbwProfiler.vbwExecuteLine 1811
                If element.ImprovedBrakes Then
vbwProfiler.vbwExecuteLine 1812
                    sOutput = sOutput & " Improved Brakes."
                End If
vbwProfiler.vbwExecuteLine 1813 'B
vbwProfiler.vbwExecuteLine 1814
                If element.AllwheelSteering Then
vbwProfiler.vbwExecuteLine 1815
                    sOutput = sOutput & " All wheel steering."
                End If
vbwProfiler.vbwExecuteLine 1816 'B
vbwProfiler.vbwExecuteLine 1817
                If element.Smartwheels Then
vbwProfiler.vbwExecuteLine 1818
                    sOutput = sOutput & " Smart wheels."
                End If
vbwProfiler.vbwExecuteLine 1819 'B
        End Select
vbwProfiler.vbwExecuteLine 1820 'B
vbwProfiler.vbwExecuteLine 1821
    Next

vbwProfiler.vbwExecuteLine 1822
    If sOutput <> "" Then
vbwProfiler.vbwExecuteLine 1823
         sOutput = "Structural Options: " + sOutput
    End If
vbwProfiler.vbwExecuteLine 1824 'B

vbwProfiler.vbwExecuteLine 1825
    GetStructuralOptionsOutput = sOutput
vbwProfiler.vbwProcOut 74
vbwProfiler.vbwExecuteLine 1826
    Exit Function
err:
vbwProfiler.vbwExecuteLine 1827
        Debug.Print "modTextOutput:GetStructuralOptionsOutput --  Error #" & err.Number & " " & TypeName(element) & " " & err.Description
vbwProfiler.vbwExecuteLine 1828
        Resume Next
vbwProfiler.vbwProcOut 74
vbwProfiler.vbwExecuteLine 1829
End Function

Private Function GetBodyFeatures() As String
vbwProfiler.vbwProcIn 75
Dim sOutput As String
Dim element As Object
Dim sSlope As String
Dim sOutput2 As String

'On Error GoTo err
'    With m_oCurrentVeh.surface
'        If .FloatationHull Then
'            sOutput = sOutput & " Floatation hull " ' todo: fix (rating " & m_oCurrentVeh.stats.FloatationRating & " lbs)."
'        End If
'        If .SubmarineLines Then
'            sOutput = sOutput & " Submarine lines."
'        End If
'        If .HydrodynamicLines <> "none" Then
'            sOutput = sOutput & " " & .HydrodynamicLines & " hydrodynamic lines."
'        End If
'        If .Catamaran Then
'            sOutput = sOutput & " Catamaran."
'        End If
'        If .Trimaran Then
'            sOutput = sOutput & " Trimaran."
'        End If
'        If .StreamLining <> "none" Then
'            sOutput = sOutput & " " & .StreamLining & " streamlining."
'        End If
'    End With
'    '//to get the slope i must check the Body, Superstructure and Turrets and Popturrets
'    For Each element In Vehicle
'        Select Case element.Datatype
'            Case Body, Superstructure, Turret, Popturret
'                sSlope = ""
'                If element.slopef <> "none" Then
'                    If sSlope <> "" Then
'                        sSlope = sSlope & ", front " & element.slopef
'                    Else
'                        sSlope = sSlope & "front " & element.slopef
'                    End If
'                End If
'                If element.slopeb <> "none" Then
'                    If sSlope <> "" Then
'                        sSlope = sSlope & ", back " & element.slopeb
'                    Else
'                        sSlope = sSlope & "back " & element.slopeb
'                    End If
'                End If
'                If element.slopel <> "none" Then
'                    If sSlope <> "" Then
'                        sSlope = sSlope & ", left " & element.slopel
'                    Else
'                        sSlope = sSlope & "left " & element.slopel
'                    End If
'                End If
'                If element.SlopeR <> "none" Then
'                    If sSlope <> "" Then
'                        sSlope = sSlope & ", right " & element.SlopeR
'                    Else
'                        sSlope = sSlope & "right " & element.SlopeR
'                    End If
'                End If
'                If sSlope <> "" Then
'                    sOutput2 = sOutput2 + " Slope on " & element.abbrev & ": " & sSlope & "."
'                End If
'        End Select
'    Next
'
'    sOutput = sOutput + sOutput2
'    If sOutput <> "" Then sOutput = sOutput
'    GetBodyFeatures = sOutput
'    Exit Function
'err:
'    Debug.Print "modTextOutput:GetBodyFeatures --  Error #" & err.Number & " " & TypeName(element) & " " & err.Description
vbwProfiler.vbwProcOut 75
vbwProfiler.vbwExecuteLine 1830
End Function

Private Function GetSurfaceFeaturesOutput() As String
vbwProfiler.vbwProcIn 76
Dim sOutput As String
Dim element As Object

vbwProfiler.vbwExecuteLine 1831
On Error GoTo err
vbwProfiler.vbwExecuteLine 1832
With m_oCurrentVeh.surface
vbwProfiler.vbwExecuteLine 1833
    If .Sealed Then
vbwProfiler.vbwExecuteLine 1834
         sOutput = sOutput + "sealed. "
    End If
vbwProfiler.vbwExecuteLine 1835 'B

vbwProfiler.vbwExecuteLine 1836
    If .Sealed = False Then
vbwProfiler.vbwExecuteLine 1837
        If .WaterProof Then
vbwProfiler.vbwExecuteLine 1838
            sOutput = sOutput + "waterproofed. "
        End If
vbwProfiler.vbwExecuteLine 1839 'B
    End If
vbwProfiler.vbwExecuteLine 1840 'B

    'concealment and stealth
vbwProfiler.vbwExecuteLine 1841
    If .Camouflage Then
vbwProfiler.vbwExecuteLine 1842
         sOutput = sOutput + "camouflage. "
    End If
vbwProfiler.vbwExecuteLine 1843 'B
vbwProfiler.vbwExecuteLine 1844
    If .infraredcloaking <> "none" Then
vbwProfiler.vbwExecuteLine 1845
         sOutput = sOutput + .infraredcloaking + " infrared cloaking. "
    End If
vbwProfiler.vbwExecuteLine 1846 'B
vbwProfiler.vbwExecuteLine 1847
    If .EmissionCloaking <> "none" Then
vbwProfiler.vbwExecuteLine 1848
         sOutput = sOutput + .EmissionCloaking + " emission cloaking. "
    End If
vbwProfiler.vbwExecuteLine 1849 'B
vbwProfiler.vbwExecuteLine 1850
    If .SoundBaffling <> "none" Then
vbwProfiler.vbwExecuteLine 1851
         sOutput = sOutput + .SoundBaffling + " sound baffling. "
    End If
vbwProfiler.vbwExecuteLine 1852 'B
vbwProfiler.vbwExecuteLine 1853
    If .stealth <> "none" Then
vbwProfiler.vbwExecuteLine 1854
         sOutput = sOutput + .stealth + " stealth. "
    End If
vbwProfiler.vbwExecuteLine 1855 'B
vbwProfiler.vbwExecuteLine 1856
    If .LiquidCrystal Then
vbwProfiler.vbwExecuteLine 1857
         sOutput = sOutput + "liquid crystal skin. "
    End If
vbwProfiler.vbwExecuteLine 1858 'B
vbwProfiler.vbwExecuteLine 1859
    If .Chameleon <> "none" Then
vbwProfiler.vbwExecuteLine 1860
         sOutput = sOutput + .Chameleon + " chameleon system. "
    End If
vbwProfiler.vbwExecuteLine 1861 'B
vbwProfiler.vbwExecuteLine 1862
    If .PsiShielding Then
vbwProfiler.vbwExecuteLine 1863
         sOutput = sOutput + "Psi Shielding. "
    End If
vbwProfiler.vbwExecuteLine 1864 'B
vbwProfiler.vbwExecuteLine 1865
    If m_oCurrentVeh.Components(BODY_KEY).liftingbody Then
vbwProfiler.vbwExecuteLine 1866
         sOutput = sOutput + "Lifting Body. "
    End If
vbwProfiler.vbwExecuteLine 1867 'B
vbwProfiler.vbwExecuteLine 1868
    If m_oCurrentVeh.Components(BODY_KEY).FlexibodyOption Then
vbwProfiler.vbwExecuteLine 1869
         sOutput = sOutput + "Flexibody. "
    End If
vbwProfiler.vbwExecuteLine 1870 'B

vbwProfiler.vbwExecuteLine 1871
    For Each element In m_oCurrentVeh.Components
vbwProfiler.vbwExecuteLine 1872
        Select Case element.Datatype
'vbwLine 1873:            Case ForceScreen, DeflectorField, VariableForceScreen
            Case IIf(vbwProfiler.vbwExecuteLine(1873), VBWPROFILER_EMPTY, _
        ForceScreen), DeflectorField, VariableForceScreen
vbwProfiler.vbwExecuteLine 1874
                sOutput = sOutput + element.PrintOutput

        End Select
vbwProfiler.vbwExecuteLine 1875 'B
vbwProfiler.vbwExecuteLine 1876
    Next
vbwProfiler.vbwExecuteLine 1877
    If .bMagicLevitation Then
vbwProfiler.vbwExecuteLine 1878
         sOutput = sOutput + "Magic Levitation. "
    End If
vbwProfiler.vbwExecuteLine 1879 'B
vbwProfiler.vbwExecuteLine 1880
    If .bAntigravityCoating Then
vbwProfiler.vbwExecuteLine 1881
         sOutput = sOutput + "Antigravity Coating. "
    End If
vbwProfiler.vbwExecuteLine 1882 'B
vbwProfiler.vbwExecuteLine 1883
    If .bSuperScienceCoating Then
vbwProfiler.vbwExecuteLine 1884
         sOutput = sOutput + "Super Science Coating. "
    End If
vbwProfiler.vbwExecuteLine 1885 'B
vbwProfiler.vbwExecuteLine 1886
End With

vbwProfiler.vbwExecuteLine 1887
If sOutput <> "" Then
vbwProfiler.vbwExecuteLine 1888
     sOutput = "Surface Features: " + sOutput
End If
vbwProfiler.vbwExecuteLine 1889 'B
vbwProfiler.vbwExecuteLine 1890
GetSurfaceFeaturesOutput = sOutput
vbwProfiler.vbwProcOut 76
vbwProfiler.vbwExecuteLine 1891
Exit Function
err:
vbwProfiler.vbwExecuteLine 1892
    Debug.Print "modTextOutput:GetSufaceFeaturesOutput --  Error #" & err.Number & " " & TypeName(element) & " " & err.Description
vbwProfiler.vbwProcOut 76
vbwProfiler.vbwExecuteLine 1893
End Function

Private Function GetOtherSurfaceFeatures() As String
vbwProfiler.vbwProcIn 77
    Dim sOutput As String
    Dim element As Object
    Dim sDefensive As String
    Dim sOutput2 As String

vbwProfiler.vbwExecuteLine 1894
    On Error GoTo err
vbwProfiler.vbwExecuteLine 1895
    With m_oCurrentVeh.Options
    'other
vbwProfiler.vbwExecuteLine 1896
        If .Convertible <> "none" Then
vbwProfiler.vbwExecuteLine 1897
             sOutput = sOutput + .Convertible + ". "
        End If
vbwProfiler.vbwExecuteLine 1898 'B
vbwProfiler.vbwExecuteLine 1899
        If .Ram Then
vbwProfiler.vbwExecuteLine 1900
             sOutput = sOutput + "Ram. "
        End If
vbwProfiler.vbwExecuteLine 1901 'B
vbwProfiler.vbwExecuteLine 1902
        If .Bulldozer Then
vbwProfiler.vbwExecuteLine 1903
             sOutput = sOutput + "Bulldozer. "
        End If
vbwProfiler.vbwExecuteLine 1904 'B
vbwProfiler.vbwExecuteLine 1905
        If .Plow Then
vbwProfiler.vbwExecuteLine 1906
             sOutput = sOutput + "Plow. "
        End If
vbwProfiler.vbwExecuteLine 1907 'B
vbwProfiler.vbwExecuteLine 1908
        If .Hitch Then
vbwProfiler.vbwExecuteLine 1909
             sOutput = sOutput + "Hitch. "
        End If
vbwProfiler.vbwExecuteLine 1910 'B
vbwProfiler.vbwExecuteLine 1911
        If .Pin <> "none" Then
vbwProfiler.vbwExecuteLine 1912
             sOutput = sOutput + .Pin + " pin. "
        End If
vbwProfiler.vbwExecuteLine 1913 'B

vbwProfiler.vbwExecuteLine 1914
        For Each element In m_oCurrentVeh.Components
vbwProfiler.vbwExecuteLine 1915
            If TypeOf element Is clsWheel Then
vbwProfiler.vbwExecuteLine 1916
                If element.Wheelblades <> "none" Then
vbwProfiler.vbwExecuteLine 1917
                     sOutput = sOutput + element.Wheelblades & " wheelblades. "
                End If
vbwProfiler.vbwExecuteLine 1918 'B
vbwProfiler.vbwExecuteLine 1919
                If element.snowtires Then
vbwProfiler.vbwExecuteLine 1920
                     sOutput = sOutput + "Snow tires. "
                End If
vbwProfiler.vbwExecuteLine 1921 'B
vbwProfiler.vbwExecuteLine 1922
                If element.racingtires Then
vbwProfiler.vbwExecuteLine 1923
                     sOutput = sOutput + "Racing tires. "
                End If
vbwProfiler.vbwExecuteLine 1924 'B
vbwProfiler.vbwExecuteLine 1925
                If element.PunctureResistant Then
vbwProfiler.vbwExecuteLine 1926
                     sOutput = sOutput + "Puncture resistant tires. "
                End If
vbwProfiler.vbwExecuteLine 1927 'B
            End If
vbwProfiler.vbwExecuteLine 1928 'B
vbwProfiler.vbwExecuteLine 1929
        Next
vbwProfiler.vbwExecuteLine 1930
    End With
vbwProfiler.vbwExecuteLine 1931
    If sOutput <> "" Then
vbwProfiler.vbwExecuteLine 1932
         sOutput = "Other Surface Features: " + sOutput
    End If
vbwProfiler.vbwExecuteLine 1933 'B
vbwProfiler.vbwExecuteLine 1934
    GetOtherSurfaceFeatures = sOutput
vbwProfiler.vbwProcOut 77
vbwProfiler.vbwExecuteLine 1935
    Exit Function
err:
vbwProfiler.vbwExecuteLine 1936
    Debug.Print "modTextOutput:GetOtherSurfaceFeatures --  Error #" & err.Number & " " & TypeName(element) & " " & err.Description
vbwProfiler.vbwProcOut 77
vbwProfiler.vbwExecuteLine 1937
End Function
Private Function GetDefensiveSurfaceFeatures() As String
vbwProfiler.vbwProcIn 78
Dim element As Object
Dim sDefensive As String
Dim sOutput2 As String

vbwProfiler.vbwExecuteLine 1938
    On Error GoTo err
    'defensive surface features found in Armor classes
vbwProfiler.vbwExecuteLine 1939
    For Each element In m_oCurrentVeh.Components
vbwProfiler.vbwExecuteLine 1940
        Select Case element.Datatype

'vbwLine 1941:            Case ArmorComplexFacing, ArmorBasicFacing, ArmorGunShield, ArmorLocation, ArmorOverall, ArmorWheelGuard
            Case IIf(vbwProfiler.vbwExecuteLine(1941), VBWPROFILER_EMPTY, _
        ArmorComplexFacing), ArmorBasicFacing, ArmorGunShield, ArmorLocation, ArmorOverall, ArmorWheelGuard

vbwProfiler.vbwExecuteLine 1942
                 sDefensive = ""
vbwProfiler.vbwExecuteLine 1943
                If element.rap Then
vbwProfiler.vbwExecuteLine 1944
                    If sDefensive <> "" Then
vbwProfiler.vbwExecuteLine 1945
                        sDefensive = sDefensive & ", reactive armor"
                    Else
vbwProfiler.vbwExecuteLine 1946 'B
vbwProfiler.vbwExecuteLine 1947
                        sDefensive = sDefensive & "reactive armor"
                    End If
vbwProfiler.vbwExecuteLine 1948 'B
                End If
vbwProfiler.vbwExecuteLine 1949 'B
vbwProfiler.vbwExecuteLine 1950
                If element.electrified Then
vbwProfiler.vbwExecuteLine 1951
                    If sDefensive <> "" Then
vbwProfiler.vbwExecuteLine 1952
                        sDefensive = sDefensive & ", electrified"
                    Else
vbwProfiler.vbwExecuteLine 1953 'B
vbwProfiler.vbwExecuteLine 1954
                        sDefensive = sDefensive & "electrified"
                    End If
vbwProfiler.vbwExecuteLine 1955 'B
                End If
vbwProfiler.vbwExecuteLine 1956 'B
vbwProfiler.vbwExecuteLine 1957
                If element.thermal Then
vbwProfiler.vbwExecuteLine 1958
                    If sDefensive <> "" Then
vbwProfiler.vbwExecuteLine 1959
                        sDefensive = sDefensive & ", thermal superconductor armor"
                    Else
vbwProfiler.vbwExecuteLine 1960 'B
vbwProfiler.vbwExecuteLine 1961
                        sDefensive = sDefensive & "thermal superconductor armor"
                    End If
vbwProfiler.vbwExecuteLine 1962 'B
                End If
vbwProfiler.vbwExecuteLine 1963 'B
vbwProfiler.vbwExecuteLine 1964
                If element.radiation Then
vbwProfiler.vbwExecuteLine 1965
                    If sDefensive <> "" Then
vbwProfiler.vbwExecuteLine 1966
                        sDefensive = sDefensive & ", radiation shielding"
                    Else
vbwProfiler.vbwExecuteLine 1967 'B
vbwProfiler.vbwExecuteLine 1968
                        sDefensive = sDefensive & "radiation shielding"
                    End If
vbwProfiler.vbwExecuteLine 1969 'B
                End If
vbwProfiler.vbwExecuteLine 1970 'B
vbwProfiler.vbwExecuteLine 1971
                If element.coating <> "none" Then
vbwProfiler.vbwExecuteLine 1972
                    If sDefensive <> "" Then
vbwProfiler.vbwExecuteLine 1973
                        sDefensive = sDefensive & ", " & element.coating & " coating"
                    Else
vbwProfiler.vbwExecuteLine 1974 'B
vbwProfiler.vbwExecuteLine 1975
                        sDefensive = sDefensive & element.coating & " coating"
                    End If
vbwProfiler.vbwExecuteLine 1976 'B
                End If
vbwProfiler.vbwExecuteLine 1977 'B
vbwProfiler.vbwExecuteLine 1978
                If sDefensive <> "" Then
vbwProfiler.vbwExecuteLine 1979
                    sOutput2 = sOutput2 + " On " & m_oCurrentVeh.Components(element.LogicalParent).Description & ": " & sDefensive & "."
                End If
vbwProfiler.vbwExecuteLine 1980 'B
        End Select
vbwProfiler.vbwExecuteLine 1981 'B
vbwProfiler.vbwExecuteLine 1982
    Next

vbwProfiler.vbwExecuteLine 1983
    If sOutput2 <> "" Then
vbwProfiler.vbwExecuteLine 1984
         sOutput2 = "Defensive Surface Features: " + sOutput2
    End If
vbwProfiler.vbwExecuteLine 1985 'B
vbwProfiler.vbwExecuteLine 1986
    GetDefensiveSurfaceFeatures = sOutput2
vbwProfiler.vbwProcOut 78
vbwProfiler.vbwExecuteLine 1987
    Exit Function
err:
vbwProfiler.vbwExecuteLine 1988
    Debug.Print "modTextOutput:GetDefensiveSurfaceFeatures --  Error #" & err.Number & " " & TypeName(element) & " " & err.Description
vbwProfiler.vbwProcOut 78
vbwProfiler.vbwExecuteLine 1989
End Function

Private Function GetTopDeckSurfaceFeatures() As String
vbwProfiler.vbwProcIn 79
Dim element As Object
Dim sTopDeck As String
Dim sOutput2 As String
Dim sDeckType As String

vbwProfiler.vbwExecuteLine 1990
    On Error GoTo err
vbwProfiler.vbwExecuteLine 1991
    For Each element In m_oCurrentVeh.Components
vbwProfiler.vbwExecuteLine 1992
        Select Case element.Datatype

'vbwLine 1993:            Case Body, Superstructure
            Case IIf(vbwProfiler.vbwExecuteLine(1993), VBWPROFILER_EMPTY, _
        Body), Superstructure
vbwProfiler.vbwExecuteLine 1994
            If element.TopDeck Then
vbwProfiler.vbwExecuteLine 1995
                 sTopDeck = ""
vbwProfiler.vbwExecuteLine 1996
                If element.covereddeckarea <> 0 Then
vbwProfiler.vbwExecuteLine 1997
                    If sTopDeck <> "" Then
vbwProfiler.vbwExecuteLine 1998
                        sTopDeck = sTopDeck & ", " & Format(element.covereddeckarea, Settings.FormatString) & "sq ft covered"
                    Else
vbwProfiler.vbwExecuteLine 1999 'B
vbwProfiler.vbwExecuteLine 2000
                        sTopDeck = sTopDeck & Format(element.covereddeckarea, Settings.FormatString) & "sq ft covered"
                    End If
vbwProfiler.vbwExecuteLine 2001 'B
                End If
vbwProfiler.vbwExecuteLine 2002 'B
vbwProfiler.vbwExecuteLine 2003
                If element.FlightDeckArea <> 0 Then
                    'get the decktype
vbwProfiler.vbwExecuteLine 2004
                    If element.flightdeckoption = "none" Then
vbwProfiler.vbwExecuteLine 2005
                        sDeckType = "flight deck"
                    Else
vbwProfiler.vbwExecuteLine 2006 'B
vbwProfiler.vbwExecuteLine 2007
                        sDeckType = element.flightdeckoption
                    End If
vbwProfiler.vbwExecuteLine 2008 'B
vbwProfiler.vbwExecuteLine 2009
                    If sTopDeck <> "" Then
vbwProfiler.vbwExecuteLine 2010
                        sTopDeck = sTopDeck & ", " & Format(element.FlightDeckArea, Settings.FormatString) & " sq ft " & sDeckType & " with a length of " & Format(element.flightdecklength, Settings.FormatString) & " ft"
                    Else
vbwProfiler.vbwExecuteLine 2011 'B
vbwProfiler.vbwExecuteLine 2012
                        sTopDeck = sTopDeck & Format(element.FlightDeckArea, Settings.FormatString) & "sq ft " & sDeckType & " with a length of " & Format(element.flightdecklength, Settings.FormatString) & " ft"
                    End If
vbwProfiler.vbwExecuteLine 2013 'B
                End If
vbwProfiler.vbwExecuteLine 2014 'B

vbwProfiler.vbwExecuteLine 2015
                If sTopDeck <> "" Then
vbwProfiler.vbwExecuteLine 2016
                    sOutput2 = sOutput2 + " On " & element.Description & ": " & sTopDeck & "."
                End If
vbwProfiler.vbwExecuteLine 2017 'B
        End If
vbwProfiler.vbwExecuteLine 2018 'B
        End Select
vbwProfiler.vbwExecuteLine 2019 'B
vbwProfiler.vbwExecuteLine 2020
    Next

vbwProfiler.vbwExecuteLine 2021
    If sOutput2 <> "" Then
vbwProfiler.vbwExecuteLine 2022
         sOutput2 = "Top Deck: " + sOutput2
    End If
vbwProfiler.vbwExecuteLine 2023 'B
vbwProfiler.vbwExecuteLine 2024
    GetTopDeckSurfaceFeatures = sOutput2
vbwProfiler.vbwProcOut 79
vbwProfiler.vbwExecuteLine 2025
    Exit Function
err:
vbwProfiler.vbwExecuteLine 2026
    Debug.Print "modTextOutput:GetTopDeckSurfaceFeatures --  Error #" & err.Number & " " & TypeName(element) & " " & err.Description
vbwProfiler.vbwProcOut 79
vbwProfiler.vbwExecuteLine 2027
End Function

Private Function GetWeaponBaysAndHardpoints() As String
vbwProfiler.vbwProcIn 80
    Dim sOutput As String
    Dim sOutput2 As String
    Dim element As Object
    Dim sngLoad As Single
    Dim sngLoad2 As Single

vbwProfiler.vbwExecuteLine 2028
    On Error GoTo err
vbwProfiler.vbwExecuteLine 2029
    For Each element In m_oCurrentVeh.Components
vbwProfiler.vbwExecuteLine 2030
        Select Case element.Datatype
'vbwLine 2031:            Case WeaponBay
            Case IIf(vbwProfiler.vbwExecuteLine(2031), VBWPROFILER_EMPTY, _
        WeaponBay)
vbwProfiler.vbwExecuteLine 2032
                sOutput = sOutput + element.abbrev + " "
vbwProfiler.vbwExecuteLine 2033
                sngLoad = sngLoad + element.loadcapacity
'vbwLine 2034:            Case HardPoint
            Case IIf(vbwProfiler.vbwExecuteLine(2034), VBWPROFILER_EMPTY, _
        HardPoint)
vbwProfiler.vbwExecuteLine 2035
                sOutput2 = sOutput2 + element.abbrev + " "
vbwProfiler.vbwExecuteLine 2036
                sngLoad2 = sngLoad2 + element.loadcapacity
        End Select
vbwProfiler.vbwExecuteLine 2037 'B
vbwProfiler.vbwExecuteLine 2038
    Next

vbwProfiler.vbwExecuteLine 2039
    If sOutput <> "" Then
vbwProfiler.vbwExecuteLine 2040
        sOutput = "Weapon bays: " & sOutput & "Total weapon bay load " & Format(sngLoad, Settings.FormatString) & " lbs."

    End If
vbwProfiler.vbwExecuteLine 2041 'B
vbwProfiler.vbwExecuteLine 2042
    If sOutput2 <> "" Then
vbwProfiler.vbwExecuteLine 2043
        sOutput2 = "Hardpoints: " & sOutput2 & "Total hardpoint load " & Format(sngLoad2, Settings.FormatString) & " lbs."
    End If
vbwProfiler.vbwExecuteLine 2044 'B
vbwProfiler.vbwExecuteLine 2045
    If (sOutput <> "") And (sOutput2 <> "") Then
vbwProfiler.vbwExecuteLine 2046
        sOutput = sOutput & vbNewLine & vbNewLine & sOutput2
    Else
vbwProfiler.vbwExecuteLine 2047 'B
vbwProfiler.vbwExecuteLine 2048
        sOutput = sOutput & sOutput2
    End If
vbwProfiler.vbwExecuteLine 2049 'B
vbwProfiler.vbwExecuteLine 2050
    GetWeaponBaysAndHardpoints = sOutput
vbwProfiler.vbwProcOut 80
vbwProfiler.vbwExecuteLine 2051
    Exit Function
err:
vbwProfiler.vbwExecuteLine 2052
    Debug.Print "modTextOutput:GetWeaponBaysAndHardpoints --  Error #" & err.Number & " " & TypeName(element) & " " & err.Description
vbwProfiler.vbwProcOut 80
vbwProfiler.vbwExecuteLine 2053
End Function

Private Function GetVisionAndDetailsOutput() As String
    'user entered vision and details
vbwProfiler.vbwProcIn 81
    Dim sOutput As String

vbwProfiler.vbwExecuteLine 2054
    On Error GoTo err
vbwProfiler.vbwExecuteLine 2055
    With m_oCurrentVeh.Description
vbwProfiler.vbwExecuteLine 2056
        If .Details <> "" Then
vbwProfiler.vbwExecuteLine 2057
            sOutput = "Details: " + .Details
        End If
vbwProfiler.vbwExecuteLine 2058 'B
vbwProfiler.vbwExecuteLine 2059
        If .Vision <> "" Then
vbwProfiler.vbwExecuteLine 2060
            sOutput = sOutput + vbNewLine + "Vision: " + .Vision
        End If
vbwProfiler.vbwExecuteLine 2061 'B
vbwProfiler.vbwExecuteLine 2062
    End With

vbwProfiler.vbwExecuteLine 2063
    GetVisionAndDetailsOutput = sOutput
vbwProfiler.vbwProcOut 81
vbwProfiler.vbwExecuteLine 2064
    Exit Function
err:
vbwProfiler.vbwExecuteLine 2065
    Debug.Print "modTextOutput:GetVisionDetailsOutput --  Error #" & err.Number & " " & err.Description
vbwProfiler.vbwProcOut 81
vbwProfiler.vbwExecuteLine 2066
End Function
Private Function GetPerformanceOutput() As String
vbwProfiler.vbwProcIn 82
Dim element As Object
Dim sOutput As String

vbwProfiler.vbwExecuteLine 2067
On Error GoTo err
vbwProfiler.vbwExecuteLine 2068
For Each element In m_oCurrentVeh.PerformanceProfiles
vbwProfiler.vbwExecuteLine 2069
    If element.Datatype = PERFORMANCEPROFILE Then

vbwProfiler.vbwExecuteLine 2070
        With element
vbwProfiler.vbwExecuteLine 2071
            Select Case .PerformanceType
'vbwLine 2072:                Case "Air"
                Case IIf(vbwProfiler.vbwExecuteLine(2072), VBWPROFILER_EMPTY, _
        "Air")
vbwProfiler.vbwExecuteLine 2073
                    sOutput = sOutput + .Key + ": "
vbwProfiler.vbwExecuteLine 2074
                    sOutput = sOutput + "Stall Speed " & Format(.aStallSpeed, "standard") & " mph."
vbwProfiler.vbwExecuteLine 2075
                    sOutput = sOutput + " Aerial motive thrust " & Format(.aMotiveThrust, "standard") & " lbs."
vbwProfiler.vbwExecuteLine 2076
                    sOutput = sOutput + " Aerodynamic drag " & .aDrag & "."
vbwProfiler.vbwExecuteLine 2077
                    sOutput = sOutput + " Top speed " & Format(.aTopSpeed, "standard") & " mph."
vbwProfiler.vbwExecuteLine 2078
                    sOutput = sOutput + " aAccel " & Format(.aAcceleration, "standard") & " mph/s."
vbwProfiler.vbwExecuteLine 2079
                    sOutput = sOutput + " aMR " & .aManeuverability & "."
vbwProfiler.vbwExecuteLine 2080
                    sOutput = sOutput + " aSR " & .aStability & "."
vbwProfiler.vbwExecuteLine 2081
                    sOutput = sOutput + " aDecel " & Format(.aDeceleration, "standard") & " mph/s."

'vbwLine 2082:                Case "Ground"
                Case IIf(vbwProfiler.vbwExecuteLine(2082), VBWPROFILER_EMPTY, _
        "Ground")
vbwProfiler.vbwExecuteLine 2083
                    sOutput = sOutput + .Key + ": "
vbwProfiler.vbwExecuteLine 2084
                    sOutput = sOutput + " Speed " & Format(.gTopSpeed, "standard") & " mph."
vbwProfiler.vbwExecuteLine 2085
                    sOutput = sOutput + " gAccel " & Format(.gAcceleration, "standard") & " mph/s."
vbwProfiler.vbwExecuteLine 2086
                    sOutput = sOutput + " gDecel " & Format(.gDeceleration, "standard") & " mph/s."
vbwProfiler.vbwExecuteLine 2087
                    sOutput = sOutput + " gSR " & .gStability & "."
vbwProfiler.vbwExecuteLine 2088
                    sOutput = sOutput + " gMR " & .gManeuverability & "."
vbwProfiler.vbwExecuteLine 2089
                    sOutput = sOutput + " " & Format(.gPressureDescription, "standard") & " ground pressure."
vbwProfiler.vbwExecuteLine 2090
                    sOutput = sOutput + " Off road speed " & Format(.gOffRoad, "standard") & " mph/s."


'vbwLine 2091:                Case "Hovercraft"
                Case IIf(vbwProfiler.vbwExecuteLine(2091), VBWPROFILER_EMPTY, _
        "Hovercraft")
vbwProfiler.vbwExecuteLine 2092
                    sOutput = sOutput + .Key + ": "
vbwProfiler.vbwExecuteLine 2093
                    sOutput = sOutput + " Hover Altitude " & .hHoverAltitude & " feet."
vbwProfiler.vbwExecuteLine 2094
                    sOutput = sOutput + " Thrust " & Format(.hMotiveThrust, "standard") & " lbs."
vbwProfiler.vbwExecuteLine 2095
                    sOutput = sOutput + " Speed " & Format(.hTopSpeed, "standard") & " mph."
vbwProfiler.vbwExecuteLine 2096
                    sOutput = sOutput + " Drag " & .hDrag & "."
vbwProfiler.vbwExecuteLine 2097
                    sOutput = sOutput + " hAccel " & Format(.hAcceleration, "standard") & " mph/s."
vbwProfiler.vbwExecuteLine 2098
                    sOutput = sOutput + " hSR " & .hstability & "."
vbwProfiler.vbwExecuteLine 2099
                    sOutput = sOutput + " hMR " & .hmaneuverability & " g."
vbwProfiler.vbwExecuteLine 2100
                    sOutput = sOutput + " hDecel " & Format(.hDeceleration, "standard") & " mph/s."


'vbwLine 2101:                Case "Mag-Lev"
                Case IIf(vbwProfiler.vbwExecuteLine(2101), VBWPROFILER_EMPTY, _
        "Mag-Lev")
vbwProfiler.vbwExecuteLine 2102
                    sOutput = sOutput + .Key + ": "
vbwProfiler.vbwExecuteLine 2103
                    sOutput = sOutput + " Thrust " & Format(.mlMotiveThrust, "standard") & " lbs."
vbwProfiler.vbwExecuteLine 2104
                    sOutput = sOutput + " Speed " & Format(.mlTopSpeed, "standard") & " mph."
vbwProfiler.vbwExecuteLine 2105
                    sOutput = sOutput + " Stall Speed " & Format(.mlStallSpeed, "standard") & " mph."
vbwProfiler.vbwExecuteLine 2106
                    sOutput = sOutput + " mDrag " & .mlDrag & "."
vbwProfiler.vbwExecuteLine 2107
                    sOutput = sOutput + " mAccel " & Format(.mlAcceleration, "standard") & " mph/s."
vbwProfiler.vbwExecuteLine 2108
                    sOutput = sOutput + " mSR " & .mlStability & "."
vbwProfiler.vbwExecuteLine 2109
                    sOutput = sOutput + " mMR " & .mlManeuverability & "."
vbwProfiler.vbwExecuteLine 2110
                    sOutput = sOutput + " mDecel " & Format(.mlDeceleration, "standard") & " mph/s."

'vbwLine 2111:                Case "Water"
                Case IIf(vbwProfiler.vbwExecuteLine(2111), VBWPROFILER_EMPTY, _
        "Water")
vbwProfiler.vbwExecuteLine 2112
                    sOutput = sOutput + .Key + ": "
vbwProfiler.vbwExecuteLine 2113
                    sOutput = sOutput + " Hydrodynamic drag " & Format(.wHydroDrag, "standard") & "."
vbwProfiler.vbwExecuteLine 2114
                    sOutput = sOutput + " Aquatic motive thrust " & Format(.wTotalAquaticThrust, "standard") & " lbs."
vbwProfiler.vbwExecuteLine 2115
                    sOutput = sOutput + " Speed " & Format(.wTopSpeed, "standard") & " mph."
vbwProfiler.vbwExecuteLine 2116
                    sOutput = sOutput + " Hydrofoil Speed " & Format(.wHydrofoilSpeed, "standard") & " mph."
vbwProfiler.vbwExecuteLine 2117
                    sOutput = sOutput + " Planing Speed " & Format(.wPlaningSpeed, "standard") & " mph."
vbwProfiler.vbwExecuteLine 2118
                    sOutput = sOutput + " wAccel " & Format(.wAcceleration, "standard") & " mph/s."
vbwProfiler.vbwExecuteLine 2119
                    sOutput = sOutput + " wMR " & .wManeuverability & "."
vbwProfiler.vbwExecuteLine 2120
                    sOutput = sOutput + " wSR  " & .wStability & "."
vbwProfiler.vbwExecuteLine 2121
                    sOutput = sOutput + " wDecel " & Format(.wDeceleration, "standard") & " mph/s."
vbwProfiler.vbwExecuteLine 2122
                    If .wIDeceleration > 0 Then
vbwProfiler.vbwExecuteLine 2123
                        sOutput = sOutput + " Incr wDecel " & Format(.wIDeceleration, "standard") & " mph/s."
                    End If
vbwProfiler.vbwExecuteLine 2124 'B
vbwProfiler.vbwExecuteLine 2125
                    sOutput = sOutput + " wDraft " & Format(.wDraft, "standard") & " feet."

'vbwLine 2126:                Case "Submerged"
                Case IIf(vbwProfiler.vbwExecuteLine(2126), VBWPROFILER_EMPTY, _
        "Submerged")
vbwProfiler.vbwExecuteLine 2127
                    sOutput = sOutput + .Key + ": "
vbwProfiler.vbwExecuteLine 2128
                    sOutput = sOutput + "suThrust " & Format(.suTotalAquaticThrust, "standard") & " lbs."
vbwProfiler.vbwExecuteLine 2129
                    sOutput = sOutput + " suDrag " & .suHydroDrag & "."
vbwProfiler.vbwExecuteLine 2130
                    sOutput = sOutput + " suSpeed " & Format(.suTopSpeed, "standard") & " mph."
vbwProfiler.vbwExecuteLine 2131
                    sOutput = sOutput + " suAccel " & Format(.suAcceleration, "standard") & " mph/s."
vbwProfiler.vbwExecuteLine 2132
                    sOutput = sOutput + " suDecel " & Format(.suDeceleration, "standard") & " mph/s."
vbwProfiler.vbwExecuteLine 2133
                    If .suIDeceleration > 0 Then
vbwProfiler.vbwExecuteLine 2134
                        sOutput = sOutput + " Incr suDecel " & Format(.suIDeceleration, "standard") & " mph/s."
                    End If
vbwProfiler.vbwExecuteLine 2135 'B
vbwProfiler.vbwExecuteLine 2136
                    sOutput = sOutput + " suSR " & .suStability & "."
vbwProfiler.vbwExecuteLine 2137
                    sOutput = sOutput + " suMR " & .suManeuverability & "."
vbwProfiler.vbwExecuteLine 2138
                    sOutput = sOutput + " Draft " & Format(.suDraft, "standard") & " feet."
vbwProfiler.vbwExecuteLine 2139
                    If .suCrushDepth = -1 Then
vbwProfiler.vbwExecuteLine 2140
                        sOutput = sOutput & "No Crush Depth"
                    Else
vbwProfiler.vbwExecuteLine 2141 'B
vbwProfiler.vbwExecuteLine 2142
                        sOutput = sOutput + " Crush Depth " & Format(.suCrushDepth, "standard") & " yards."
                     End If
vbwProfiler.vbwExecuteLine 2143 'B

'vbwLine 2144:               Case "Space"
               Case IIf(vbwProfiler.vbwExecuteLine(2144), VBWPROFILER_EMPTY, _
        "Space")
vbwProfiler.vbwExecuteLine 2145
                    sOutput = sOutput + .Key + ": "
vbwProfiler.vbwExecuteLine 2146
                    sOutput = sOutput + " Thrust " & Format(.sMotiveThrust, "standard") & " lbs."
vbwProfiler.vbwExecuteLine 2147
                    sOutput = sOutput + " sAccel " & Format(.sAccelerationG, "standard") & " g."
vbwProfiler.vbwExecuteLine 2148
                    sOutput = sOutput + " sAccel " & Format(.sAccelerationMPH, "standard") & " mph/s."
vbwProfiler.vbwExecuteLine 2149
                    sOutput = sOutput + " Turn Around " & Format(.sTurnAroundTime, "standard") & " secs."
vbwProfiler.vbwExecuteLine 2150
                    sOutput = sOutput + " sMR " & Format(.sManeuverability, "standard") & "."
vbwProfiler.vbwExecuteLine 2151
                    sOutput = sOutput + " Hyper " & Format(.sHyperSpeed, "standard") & " parsecs per day."
vbwProfiler.vbwExecuteLine 2152
                    sOutput = sOutput + " Warp " & Format(.sWarpSpeed, "standard") & " parsecs per day."
vbwProfiler.vbwExecuteLine 2153
                    If .sJumpDriveable Then
vbwProfiler.vbwExecuteLine 2154
                        sOutput = sOutput + " Has jump capabilities."
                    End If
vbwProfiler.vbwExecuteLine 2155 'B
vbwProfiler.vbwExecuteLine 2156
                    If .sTeleportationDriveable Then
vbwProfiler.vbwExecuteLine 2157
                        sOutput = sOutput + " Has teleportation drive capabilities."
                    End If
vbwProfiler.vbwExecuteLine 2158 'B
            End Select
vbwProfiler.vbwExecuteLine 2159 'B
vbwProfiler.vbwExecuteLine 2160
            sOutput = sOutput + vbNewLine + vbNewLine
vbwProfiler.vbwExecuteLine 2161
        End With
    End If
vbwProfiler.vbwExecuteLine 2162 'B
vbwProfiler.vbwExecuteLine 2163
Next

vbwProfiler.vbwExecuteLine 2164
GetPerformanceOutput = sOutput
vbwProfiler.vbwProcOut 82
vbwProfiler.vbwExecuteLine 2165
Exit Function
err:
vbwProfiler.vbwExecuteLine 2166
    Debug.Print "modTextOutput:GetPerformanceOutput --  Error #" & err.Number & " " & TypeName(element) & " " & err.Description
vbwProfiler.vbwProcOut 82
vbwProfiler.vbwExecuteLine 2167
End Function

Private Function GetDetailedWeaponStats() As String
vbwProfiler.vbwProcIn 83
Dim element As Object
Dim sOutput As String
Dim aHeader, bHeader, cHeader As Boolean
Dim gun() As Variant
Dim i, j, k, l As Long
Dim iLength As Long
Dim iOldLength As Long
Dim iPropID As Long


vbwProfiler.vbwExecuteLine 2168
On Error GoTo err

vbwProfiler.vbwExecuteLine 2169
i = 1
vbwProfiler.vbwExecuteLine 2170
j = 1
vbwProfiler.vbwExecuteLine 2171
k = 1
vbwProfiler.vbwExecuteLine 2172
l = 1
    '//guns and artillery
vbwProfiler.vbwExecuteLine 2173
    bHeader = False
vbwProfiler.vbwExecuteLine 2174
    ReDim gun(1 To 15, 1)
vbwProfiler.vbwExecuteLine 2175
    For Each element In m_oCurrentVeh.Components
vbwProfiler.vbwExecuteLine 2176
        Select Case element.Datatype
'vbwLine 2177:            Case StoneThrower, BoltThrower, RepeatingBoltThrower, MuzzleLoader, BreechLoader, ManualRepeater, Revolver, MechanicalGatling, SlowAutoloader, FastAutoloader, lightAutomatic, HeavyAutomatic, ElectricGatling
            Case IIf(vbwProfiler.vbwExecuteLine(2177), VBWPROFILER_EMPTY, _
        StoneThrower), BoltThrower, RepeatingBoltThrower, MuzzleLoader, BreechLoader, ManualRepeater, Revolver, MechanicalGatling, SlowAutoloader, FastAutoloader, lightAutomatic, HeavyAutomatic, ElectricGatling


vbwProfiler.vbwExecuteLine 2178
                If bHeader = False Then
                    'we havent printed our header yet so do it now
vbwProfiler.vbwExecuteLine 2179
                    gun(1, 1) = "Name"
vbwProfiler.vbwExecuteLine 2180
                    gun(2, 1) = "Malf"
vbwProfiler.vbwExecuteLine 2181
                    gun(3, 1) = "Type"
vbwProfiler.vbwExecuteLine 2182
                    gun(4, 1) = "Damage"
vbwProfiler.vbwExecuteLine 2183
                    gun(5, 1) = "SS"
vbwProfiler.vbwExecuteLine 2184
                    gun(6, 1) = "Acc"
vbwProfiler.vbwExecuteLine 2185
                    gun(7, 1) = "1/2D"
vbwProfiler.vbwExecuteLine 2186
                    gun(8, 1) = "Max"
vbwProfiler.vbwExecuteLine 2187
                    gun(9, 1) = "RoF"
vbwProfiler.vbwExecuteLine 2188
                    gun(10, 1) = "Weight"
vbwProfiler.vbwExecuteLine 2189
                    gun(11, 1) = "Cost"
vbwProfiler.vbwExecuteLine 2190
                    gun(12, 1) = "WPS"
vbwProfiler.vbwExecuteLine 2191
                    gun(13, 1) = "VPS"
vbwProfiler.vbwExecuteLine 2192
                    gun(14, 1) = "CPS"
vbwProfiler.vbwExecuteLine 2193
                    gun(15, 1) = "Ldrs."
vbwProfiler.vbwExecuteLine 2194
                    bHeader = True
                End If
vbwProfiler.vbwExecuteLine 2195 'B
vbwProfiler.vbwExecuteLine 2196
                i = i + 1
vbwProfiler.vbwExecuteLine 2197
                ReDim Preserve gun(1 To 15, i)
vbwProfiler.vbwExecuteLine 2198
                gun(1, i) = element.CustomDescription
vbwProfiler.vbwExecuteLine 2199
                gun(2, i) = element.Malfunction
vbwProfiler.vbwExecuteLine 2200
                gun(3, i) = element.TypeDamage1
vbwProfiler.vbwExecuteLine 2201
                gun(4, i) = element.Damage1
vbwProfiler.vbwExecuteLine 2202
                gun(5, i) = element.SnapShot
vbwProfiler.vbwExecuteLine 2203
                gun(6, i) = element.Accuracy
vbwProfiler.vbwExecuteLine 2204
                gun(7, i) = element.halfDamage
vbwProfiler.vbwExecuteLine 2205
                gun(8, i) = element.MaxRange
vbwProfiler.vbwExecuteLine 2206
                gun(9, i) = element.sRoF
vbwProfiler.vbwExecuteLine 2207
                gun(10, i) = element.Weight
vbwProfiler.vbwExecuteLine 2208
                gun(11, i) = element.Cost
vbwProfiler.vbwExecuteLine 2209
                gun(12, i) = element.WPS
vbwProfiler.vbwExecuteLine 2210
                gun(13, i) = element.VPS
vbwProfiler.vbwExecuteLine 2211
                gun(14, i) = element.CPS
vbwProfiler.vbwExecuteLine 2212
                gun(15, i) = element.Loaders
        End Select
vbwProfiler.vbwExecuteLine 2213 'B
vbwProfiler.vbwExecuteLine 2214
    Next
    '//now we must pad each row item with spaces so that they are all the same length
vbwProfiler.vbwExecuteLine 2215
    If gun(1, 1) <> "" Then
vbwProfiler.vbwExecuteLine 2216
        For iPropID = 1 To 15
vbwProfiler.vbwExecuteLine 2217
            iLength = 0
vbwProfiler.vbwExecuteLine 2218
            iOldLength = 0
vbwProfiler.vbwExecuteLine 2219
            For j = 1 To i
vbwProfiler.vbwExecuteLine 2220
                iLength = Maximum(iLength, Len(gun(iPropID, j)))
vbwProfiler.vbwExecuteLine 2221
            Next
vbwProfiler.vbwExecuteLine 2222
            iLength = iLength + 1 '//we need 1 space seperation
vbwProfiler.vbwExecuteLine 2223
            For j = 1 To i
vbwProfiler.vbwExecuteLine 2224
                iOldLength = Len(gun(iPropID, j))
vbwProfiler.vbwExecuteLine 2225
                For k = 1 To iLength - iOldLength
vbwProfiler.vbwExecuteLine 2226
                    gun(iPropID, j) = gun(iPropID, j) & " "
vbwProfiler.vbwExecuteLine 2227
                Next
vbwProfiler.vbwExecuteLine 2228
            Next
vbwProfiler.vbwExecuteLine 2229
        Next
        '//finally we can output it all
vbwProfiler.vbwExecuteLine 2230
        For j = 1 To i
vbwProfiler.vbwExecuteLine 2231
            sOutput = sOutput + sLineBreak
vbwProfiler.vbwExecuteLine 2232
            sOutput = sOutput & gun(1, j) & gun(2, j) & gun(3, j) & gun(4, j) & gun(5, j) & gun(6, j) & gun(7, j) & gun(8, j) & gun(9, j) & gun(10, j) & gun(11, j) & gun(12, j) & gun(13, j) & gun(14, j) & gun(15, j)
vbwProfiler.vbwExecuteLine 2233
        Next
vbwProfiler.vbwExecuteLine 2234
        sOutput = sOutput + sLineBreak
vbwProfiler.vbwExecuteLine 2235
        gun(1, 1) = ""
    End If
vbwProfiler.vbwExecuteLine 2236 'B
    '////////////////////////////////////////////////////////////////////
    '//Beam Weapons
vbwProfiler.vbwExecuteLine 2237
    bHeader = False
vbwProfiler.vbwExecuteLine 2238
    i = 1
vbwProfiler.vbwExecuteLine 2239
    ReDim gun(1 To 12, 1)
vbwProfiler.vbwExecuteLine 2240
    For Each element In m_oCurrentVeh.Components
vbwProfiler.vbwExecuteLine 2241
        Select Case element.Datatype
'vbwLine 2242:            Case BlueGreenLaser, RainbowLaser, Laser, UVLaser, IRLaser, Disruptor, ChargedParticleBeam, NeutralParticleBeam, Flamer, Screamer, Stunner, ParalysisBeam, XRayLaser, FusionBeam, GravityBeam, AntiparticleBeam, Graser, Disintegrator, Displacer, BeamedPowerTransmitter, MilitaryParalysisBeam, EnergyDrill
            Case IIf(vbwProfiler.vbwExecuteLine(2242), VBWPROFILER_EMPTY, _
        BlueGreenLaser), RainbowLaser, Laser, UVLaser, IRLaser, Disruptor, ChargedParticleBeam, NeutralParticleBeam, Flamer, Screamer, Stunner, ParalysisBeam, XRayLaser, FusionBeam, GravityBeam, AntiparticleBeam, Graser, Disintegrator, Displacer, BeamedPowerTransmitter, MilitaryParalysisBeam, EnergyDrill

vbwProfiler.vbwExecuteLine 2243
                If bHeader = False Then
                    'we havent printed our header yet so do it now
vbwProfiler.vbwExecuteLine 2244
                    gun(1, 1) = "Name"
vbwProfiler.vbwExecuteLine 2245
                    gun(2, 1) = "Malf"
vbwProfiler.vbwExecuteLine 2246
                    gun(3, 1) = "Type"
vbwProfiler.vbwExecuteLine 2247
                    gun(4, 1) = "Damage"
vbwProfiler.vbwExecuteLine 2248
                    gun(5, 1) = "SS"
vbwProfiler.vbwExecuteLine 2249
                    gun(6, 1) = "Acc"
vbwProfiler.vbwExecuteLine 2250
                    gun(7, 1) = "1/2D"
vbwProfiler.vbwExecuteLine 2251
                    gun(8, 1) = "Max"
vbwProfiler.vbwExecuteLine 2252
                    gun(9, 1) = "RoF"
vbwProfiler.vbwExecuteLine 2253
                    gun(10, 1) = "Weight"
vbwProfiler.vbwExecuteLine 2254
                    gun(11, 1) = "Cost"
vbwProfiler.vbwExecuteLine 2255
                    gun(12, 1) = "Power"
vbwProfiler.vbwExecuteLine 2256
                    bHeader = True
                End If
vbwProfiler.vbwExecuteLine 2257 'B
vbwProfiler.vbwExecuteLine 2258
                i = i + 1
vbwProfiler.vbwExecuteLine 2259
                ReDim Preserve gun(1 To 12, i)
vbwProfiler.vbwExecuteLine 2260
                gun(1, i) = element.CustomDescription
vbwProfiler.vbwExecuteLine 2261
                gun(2, i) = element.Malfunction
vbwProfiler.vbwExecuteLine 2262
                gun(3, i) = element.TypeDamage
vbwProfiler.vbwExecuteLine 2263
                gun(4, i) = element.Damage
vbwProfiler.vbwExecuteLine 2264
                gun(5, i) = element.SnapShot
vbwProfiler.vbwExecuteLine 2265
                gun(6, i) = element.Accuracy
vbwProfiler.vbwExecuteLine 2266
                gun(7, i) = element.halfDamage
vbwProfiler.vbwExecuteLine 2267
                gun(8, i) = element.MaxRange
vbwProfiler.vbwExecuteLine 2268
                gun(9, i) = element.rof
vbwProfiler.vbwExecuteLine 2269
                gun(10, i) = element.Weight
vbwProfiler.vbwExecuteLine 2270
                gun(11, i) = element.Cost
vbwProfiler.vbwExecuteLine 2271
                gun(12, i) = element.PowerReqt
        End Select
vbwProfiler.vbwExecuteLine 2272 'B
vbwProfiler.vbwExecuteLine 2273
    Next
    '//now we must pad each row item with spaces so that they are all the same length
vbwProfiler.vbwExecuteLine 2274
    If gun(1, 1) <> "" Then
vbwProfiler.vbwExecuteLine 2275
        For iPropID = 1 To 12
vbwProfiler.vbwExecuteLine 2276
            iLength = 0
vbwProfiler.vbwExecuteLine 2277
            iOldLength = 0
vbwProfiler.vbwExecuteLine 2278
            For j = 1 To i
vbwProfiler.vbwExecuteLine 2279
                iLength = Maximum(iLength, Len(gun(iPropID, j)))
vbwProfiler.vbwExecuteLine 2280
            Next
vbwProfiler.vbwExecuteLine 2281
            iLength = iLength + 1 '//we need 1 space seperation
vbwProfiler.vbwExecuteLine 2282
            For j = 1 To i
vbwProfiler.vbwExecuteLine 2283
                iOldLength = Len(gun(iPropID, j))
vbwProfiler.vbwExecuteLine 2284
                For k = 1 To iLength - iOldLength
vbwProfiler.vbwExecuteLine 2285
                    gun(iPropID, j) = gun(iPropID, j) & " "
vbwProfiler.vbwExecuteLine 2286
                Next
vbwProfiler.vbwExecuteLine 2287
            Next
vbwProfiler.vbwExecuteLine 2288
        Next
        '//finally we can output it all
vbwProfiler.vbwExecuteLine 2289
        For j = 1 To i
vbwProfiler.vbwExecuteLine 2290
            sOutput = sOutput + sLineBreak
vbwProfiler.vbwExecuteLine 2291
            sOutput = sOutput & gun(1, j) & gun(2, j) & gun(3, j) & gun(4, j) & gun(5, j) & gun(6, j) & gun(7, j) & gun(8, j) & gun(9, j) & gun(10, j) & gun(11, j) & gun(12, j)
vbwProfiler.vbwExecuteLine 2292
        Next
vbwProfiler.vbwExecuteLine 2293
        sOutput = sOutput + sLineBreak
vbwProfiler.vbwExecuteLine 2294
        gun(1, 1) = ""
    End If
vbwProfiler.vbwExecuteLine 2295 'B
    '////////////////////////////////////////////////////////////////
    '//Bombs, missiles and torps
vbwProfiler.vbwExecuteLine 2296
    bHeader = False
vbwProfiler.vbwExecuteLine 2297
    i = 1
vbwProfiler.vbwExecuteLine 2298
    ReDim gun(1 To 13, 1)
vbwProfiler.vbwExecuteLine 2299
    For Each element In m_oCurrentVeh.Components
vbwProfiler.vbwExecuteLine 2300
        Select Case element.Datatype
'vbwLine 2301:            Case IronBomb, RetardedBomb, SmartBomb, SelfDestructSystem, ContactMine, ProximityMine, PressureTriggerMine, CommandTriggerMine, SmartTriggerMine, UnGuidedMissile, UnGuidedTorpedo, GuidedMissile, GuidedTorpedo
            Case IIf(vbwProfiler.vbwExecuteLine(2301), VBWPROFILER_EMPTY, _
        IronBomb), RetardedBomb, SmartBomb, SelfDestructSystem, ContactMine, ProximityMine, PressureTriggerMine, CommandTriggerMine, SmartTriggerMine, UnGuidedMissile, UnGuidedTorpedo, GuidedMissile, GuidedTorpedo

vbwProfiler.vbwExecuteLine 2302
                If bHeader = False Then
                    'we havent printed our header yet so do it now
vbwProfiler.vbwExecuteLine 2303
                    gun(1, 1) = "Name"
vbwProfiler.vbwExecuteLine 2304
                    gun(2, 1) = "Malf"
vbwProfiler.vbwExecuteLine 2305
                    gun(3, 1) = "Guid"
vbwProfiler.vbwExecuteLine 2306
                    gun(4, 1) = "Type"
vbwProfiler.vbwExecuteLine 2307
                    gun(5, 1) = "Damage"
vbwProfiler.vbwExecuteLine 2308
                    gun(6, 1) = "Spd"
vbwProfiler.vbwExecuteLine 2309
                    gun(7, 1) = "End"
vbwProfiler.vbwExecuteLine 2310
                    gun(8, 1) = "Max"
vbwProfiler.vbwExecuteLine 2311
                    gun(9, 1) = "Min"
vbwProfiler.vbwExecuteLine 2312
                    gun(10, 1) = "Skill"
vbwProfiler.vbwExecuteLine 2313
                    gun(11, 1) = "WPS"
vbwProfiler.vbwExecuteLine 2314
                    gun(12, 1) = "VPS"
vbwProfiler.vbwExecuteLine 2315
                    gun(13, 1) = "CPS"

vbwProfiler.vbwExecuteLine 2316
                    bHeader = True
                End If
vbwProfiler.vbwExecuteLine 2317 'B
vbwProfiler.vbwExecuteLine 2318
                i = i + 1
vbwProfiler.vbwExecuteLine 2319
                ReDim Preserve gun(1 To 13, i)
vbwProfiler.vbwExecuteLine 2320
                gun(1, i) = element.CustomDescription
vbwProfiler.vbwExecuteLine 2321
                gun(2, i) = element.Malfunction
vbwProfiler.vbwExecuteLine 2322
                gun(3, i) = element.GuidanceSystem
vbwProfiler.vbwExecuteLine 2323
                gun(4, i) = element.TypeDamage1
vbwProfiler.vbwExecuteLine 2324
                gun(5, i) = element.Damage1
vbwProfiler.vbwExecuteLine 2325
                gun(6, i) = element.Speed
vbwProfiler.vbwExecuteLine 2326
                gun(7, i) = element.Endurance
vbwProfiler.vbwExecuteLine 2327
                gun(8, i) = element.MaxRange
vbwProfiler.vbwExecuteLine 2328
                gun(9, i) = element.MinRange
vbwProfiler.vbwExecuteLine 2329
                gun(10, i) = element.Skill
vbwProfiler.vbwExecuteLine 2330
                gun(11, i) = element.Weight / element.Quantity
vbwProfiler.vbwExecuteLine 2331
                gun(12, i) = element.Volume
vbwProfiler.vbwExecuteLine 2332
                gun(13, i) = element.Cost / element.Quantity
        End Select
vbwProfiler.vbwExecuteLine 2333 'B
vbwProfiler.vbwExecuteLine 2334
    Next
    '//now we must pad each row item with spaces so that they are all the same length
vbwProfiler.vbwExecuteLine 2335
    If gun(1, 1) <> "" Then
vbwProfiler.vbwExecuteLine 2336
        For iPropID = 1 To 13
vbwProfiler.vbwExecuteLine 2337
            iLength = 0
vbwProfiler.vbwExecuteLine 2338
            iOldLength = 0
vbwProfiler.vbwExecuteLine 2339
            For j = 1 To i
vbwProfiler.vbwExecuteLine 2340
                iLength = Maximum(iLength, Len(gun(iPropID, j)))
vbwProfiler.vbwExecuteLine 2341
            Next
vbwProfiler.vbwExecuteLine 2342
            iLength = iLength + 1 '//we need 1 space seperation
vbwProfiler.vbwExecuteLine 2343
            For j = 1 To i
vbwProfiler.vbwExecuteLine 2344
                iOldLength = Len(gun(iPropID, j))
vbwProfiler.vbwExecuteLine 2345
                For k = 1 To iLength - iOldLength
vbwProfiler.vbwExecuteLine 2346
                    gun(iPropID, j) = gun(iPropID, j) & " "
vbwProfiler.vbwExecuteLine 2347
                Next
vbwProfiler.vbwExecuteLine 2348
            Next
vbwProfiler.vbwExecuteLine 2349
        Next
        '//finally we can output it all
vbwProfiler.vbwExecuteLine 2350
        For j = 1 To i
vbwProfiler.vbwExecuteLine 2351
            sOutput = sOutput + sLineBreak
vbwProfiler.vbwExecuteLine 2352
            sOutput = sOutput & gun(1, j) & gun(2, j) & gun(3, j) & gun(4, j) & gun(5, j) & gun(6, j) & gun(7, j) & gun(8, j) & gun(9, j) & gun(10, j) & gun(11, j) & gun(12, j) & gun(13, j)
vbwProfiler.vbwExecuteLine 2353
        Next
vbwProfiler.vbwExecuteLine 2354
        sOutput = sOutput + sLineBreak
vbwProfiler.vbwExecuteLine 2355
        gun(1, 1) = ""
     End If
vbwProfiler.vbwExecuteLine 2356 'B
     '/////////////////////////////////////////////////////////////
     '//Liquid projectors
vbwProfiler.vbwExecuteLine 2357
     bHeader = False
vbwProfiler.vbwExecuteLine 2358
     i = 1
vbwProfiler.vbwExecuteLine 2359
     ReDim gun(1 To 13, 1)
vbwProfiler.vbwExecuteLine 2360
     For Each element In m_oCurrentVeh.Components
vbwProfiler.vbwExecuteLine 2361
        Select Case element.Datatype
'vbwLine 2362:            Case FlameThrower, WaterCannon
            Case IIf(vbwProfiler.vbwExecuteLine(2362), VBWPROFILER_EMPTY, _
        FlameThrower), WaterCannon

vbwProfiler.vbwExecuteLine 2363
                If bHeader = False Then
                    'we havent printed our header yet so do it now
vbwProfiler.vbwExecuteLine 2364
                    gun(1, 1) = "Name"
vbwProfiler.vbwExecuteLine 2365
                    gun(2, 1) = "Malf"
vbwProfiler.vbwExecuteLine 2366
                    gun(3, 1) = "Type"
vbwProfiler.vbwExecuteLine 2367
                    gun(4, 1) = "Damage"
vbwProfiler.vbwExecuteLine 2368
                    gun(5, 1) = "SS"
vbwProfiler.vbwExecuteLine 2369
                    gun(6, 1) = "Acc"
vbwProfiler.vbwExecuteLine 2370
                    gun(7, 1) = "1/2D"
vbwProfiler.vbwExecuteLine 2371
                    gun(8, 1) = "Max"
vbwProfiler.vbwExecuteLine 2372
                    gun(9, 1) = "RoF"
vbwProfiler.vbwExecuteLine 2373
                    gun(10, 1) = "Weight"
vbwProfiler.vbwExecuteLine 2374
                    gun(11, 1) = "Cost"
vbwProfiler.vbwExecuteLine 2375
                    gun(12, 1) = "WPS"
vbwProfiler.vbwExecuteLine 2376
                    gun(13, 1) = "CPS"
vbwProfiler.vbwExecuteLine 2377
                    bHeader = True
                End If
vbwProfiler.vbwExecuteLine 2378 'B
vbwProfiler.vbwExecuteLine 2379
                i = i + 1
vbwProfiler.vbwExecuteLine 2380
                ReDim Preserve gun(1 To 13, i)
vbwProfiler.vbwExecuteLine 2381
                gun(1, i) = element.CustomDescription
vbwProfiler.vbwExecuteLine 2382
                gun(2, i) = element.Malfunction
vbwProfiler.vbwExecuteLine 2383
                gun(3, i) = element.TypeDamage
vbwProfiler.vbwExecuteLine 2384
                gun(4, i) = element.Damage
vbwProfiler.vbwExecuteLine 2385
                gun(5, i) = element.SnapShot
vbwProfiler.vbwExecuteLine 2386
                gun(6, i) = element.Accuracy
vbwProfiler.vbwExecuteLine 2387
                gun(7, i) = element.halfDamage
vbwProfiler.vbwExecuteLine 2388
                gun(8, i) = element.MaxRange
vbwProfiler.vbwExecuteLine 2389
                gun(9, i) = element.rof
vbwProfiler.vbwExecuteLine 2390
                gun(10, i) = element.Weight
vbwProfiler.vbwExecuteLine 2391
                gun(11, i) = element.Cost
vbwProfiler.vbwExecuteLine 2392
                gun(12, i) = element.WPS
vbwProfiler.vbwExecuteLine 2393
                gun(13, i) = element.CPS
        End Select
vbwProfiler.vbwExecuteLine 2394 'B
vbwProfiler.vbwExecuteLine 2395
    Next
    '//now we must pad each row item with spaces so that they are all the same length
vbwProfiler.vbwExecuteLine 2396
    If gun(1, 1) <> "" Then
vbwProfiler.vbwExecuteLine 2397
        For iPropID = 1 To 13
vbwProfiler.vbwExecuteLine 2398
            iLength = 0
vbwProfiler.vbwExecuteLine 2399
            iOldLength = 0
vbwProfiler.vbwExecuteLine 2400
            For j = 1 To i
vbwProfiler.vbwExecuteLine 2401
                iLength = Maximum(iLength, Len(gun(iPropID, j)))
vbwProfiler.vbwExecuteLine 2402
            Next
vbwProfiler.vbwExecuteLine 2403
            iLength = iLength + 1 '//we need 1 space seperation
vbwProfiler.vbwExecuteLine 2404
            For j = 1 To i
vbwProfiler.vbwExecuteLine 2405
                iOldLength = Len(gun(iPropID, j))
vbwProfiler.vbwExecuteLine 2406
                For k = 1 To iLength - iOldLength
vbwProfiler.vbwExecuteLine 2407
                    gun(iPropID, j) = gun(iPropID, j) & " "
vbwProfiler.vbwExecuteLine 2408
                Next
vbwProfiler.vbwExecuteLine 2409
            Next
vbwProfiler.vbwExecuteLine 2410
        Next
        '//finally we can output it all
vbwProfiler.vbwExecuteLine 2411
        For j = 1 To i
vbwProfiler.vbwExecuteLine 2412
            sOutput = sOutput + sLineBreak
vbwProfiler.vbwExecuteLine 2413
            sOutput = sOutput & gun(1, j) & gun(2, j) & gun(3, j) & gun(4, j) & gun(5, j) & gun(6, j) & gun(7, j) & gun(8, j) & gun(9, j) & gun(10, j) & gun(11, j) & gun(12, j) & gun(13, j)
vbwProfiler.vbwExecuteLine 2414
        Next
vbwProfiler.vbwExecuteLine 2415
        sOutput = sOutput + sLineBreak
vbwProfiler.vbwExecuteLine 2416
        gun(1, 1) = ""
    End If
vbwProfiler.vbwExecuteLine 2417 'B
    '///////////////////////////////////////////////////////////////
    '//Launchers
vbwProfiler.vbwExecuteLine 2418
    bHeader = False
vbwProfiler.vbwExecuteLine 2419
    i = 1
vbwProfiler.vbwExecuteLine 2420
    ReDim gun(1 To 7, 1)
vbwProfiler.vbwExecuteLine 2421
    For Each element In m_oCurrentVeh.Components
vbwProfiler.vbwExecuteLine 2422
        Select Case element.Datatype
'vbwLine 2423:            Case DisposableLauncher, MuzzleloadingLauncher, BreechloadingLauncher, ManualRepeaterLauncher, SlowAutoLoaderLauncher, FastAutoLoaderLauncher, RevolverLauncher, lightAutomaticLauncher, HeavyAutomaticLauncher
            Case IIf(vbwProfiler.vbwExecuteLine(2423), VBWPROFILER_EMPTY, _
        DisposableLauncher), MuzzleloadingLauncher, BreechloadingLauncher, ManualRepeaterLauncher, SlowAutoLoaderLauncher, FastAutoLoaderLauncher, RevolverLauncher, lightAutomaticLauncher, HeavyAutomaticLauncher

vbwProfiler.vbwExecuteLine 2424
                If bHeader = False Then
                    'we havent printed our header yet so do it now
vbwProfiler.vbwExecuteLine 2425
                    gun(1, 1) = "Name"
vbwProfiler.vbwExecuteLine 2426
                    gun(2, 1) = "SS"
vbwProfiler.vbwExecuteLine 2427
                    gun(3, 1) = "RoF"
vbwProfiler.vbwExecuteLine 2428
                    gun(4, 1) = "Weight"
vbwProfiler.vbwExecuteLine 2429
                    gun(5, 1) = "Cost"
vbwProfiler.vbwExecuteLine 2430
                    gun(6, 1) = "Ldrs."
vbwProfiler.vbwExecuteLine 2431
                    gun(7, 1) = "Rating"
vbwProfiler.vbwExecuteLine 2432
                    bHeader = True
                End If
vbwProfiler.vbwExecuteLine 2433 'B
vbwProfiler.vbwExecuteLine 2434
                i = i + 1
vbwProfiler.vbwExecuteLine 2435
                ReDim Preserve gun(1 To 7, i)
vbwProfiler.vbwExecuteLine 2436
                gun(1, i) = element.CustomDescription
vbwProfiler.vbwExecuteLine 2437
                gun(2, i) = element.SnapShot
vbwProfiler.vbwExecuteLine 2438
                gun(3, i) = element.rof
vbwProfiler.vbwExecuteLine 2439
                gun(4, i) = element.Weight
vbwProfiler.vbwExecuteLine 2440
                gun(5, i) = element.Cost
vbwProfiler.vbwExecuteLine 2441
                gun(6, i) = element.Loaders
vbwProfiler.vbwExecuteLine 2442
                gun(7, i) = element.MaxLoad
        End Select
vbwProfiler.vbwExecuteLine 2443 'B
vbwProfiler.vbwExecuteLine 2444
    Next
    '//now we must pad each row item with spaces so that they are all the same length
vbwProfiler.vbwExecuteLine 2445
    If gun(1, 1) <> "" Then
vbwProfiler.vbwExecuteLine 2446
        For iPropID = 1 To 7
vbwProfiler.vbwExecuteLine 2447
            iLength = 0
vbwProfiler.vbwExecuteLine 2448
            iOldLength = 0
vbwProfiler.vbwExecuteLine 2449
            For j = 1 To i
vbwProfiler.vbwExecuteLine 2450
                iLength = Maximum(iLength, Len(gun(iPropID, j)))
vbwProfiler.vbwExecuteLine 2451
            Next
vbwProfiler.vbwExecuteLine 2452
            iLength = iLength + 1 '//we need 1 space seperation
vbwProfiler.vbwExecuteLine 2453
            For j = 1 To i
vbwProfiler.vbwExecuteLine 2454
                iOldLength = Len(gun(iPropID, j))
vbwProfiler.vbwExecuteLine 2455
                For k = 1 To iLength - iOldLength
vbwProfiler.vbwExecuteLine 2456
                    gun(iPropID, j) = gun(iPropID, j) & " "
vbwProfiler.vbwExecuteLine 2457
                Next
vbwProfiler.vbwExecuteLine 2458
            Next
vbwProfiler.vbwExecuteLine 2459
        Next
        '//finally we can output it all
vbwProfiler.vbwExecuteLine 2460
        For j = 1 To i
vbwProfiler.vbwExecuteLine 2461
            sOutput = sOutput + sLineBreak
vbwProfiler.vbwExecuteLine 2462
            sOutput = sOutput & gun(1, j) & gun(2, j) & gun(3, j) & gun(4, j) & gun(5, j) & gun(6, j) & gun(7, j)
vbwProfiler.vbwExecuteLine 2463
        Next
vbwProfiler.vbwExecuteLine 2464
        sOutput = sOutput + sLineBreak
vbwProfiler.vbwExecuteLine 2465
        gun(1, 1) = ""
    End If
vbwProfiler.vbwExecuteLine 2466 'B
    '//add a new line and send then return the entire output value
vbwProfiler.vbwExecuteLine 2467
    sOutput = sOutput + vbNewLine
vbwProfiler.vbwExecuteLine 2468
    GetDetailedWeaponStats = sOutput
vbwProfiler.vbwProcOut 83
vbwProfiler.vbwExecuteLine 2469
   Exit Function
err:
vbwProfiler.vbwExecuteLine 2470
    Debug.Print "modTextOutput:GetDetailedWeaponStats --  Error #" & err.Number & " " & TypeName(element) & " " & err.Description
vbwProfiler.vbwProcOut 83
vbwProfiler.vbwExecuteLine 2471
End Function

Private Function NumericToString(ByVal nNumber As Variant) As String
'//this function accepts a number and if that number is between
'1 and 10 it will convert them to "One" and "Ten" for instance. If
'the number is greater than 10 it will just return the number formatted
'as a string
vbwProfiler.vbwProcIn 84
Dim retval As String

'NOTE: This is currently only set up to handle longs and not decimals
vbwProfiler.vbwExecuteLine 2472
If nNumber >= 1 And nNumber <= 10 Then
vbwProfiler.vbwExecuteLine 2473
    If nNumber = 1 Then
vbwProfiler.vbwExecuteLine 2474
        retval = "One"
        'retval = "" 'if its a 1 we'll jsut leave it blank since its assumed to be 1 unless noted
'vbwLine 2475:    ElseIf nNumber = 2 Then
    ElseIf vbwProfiler.vbwExecuteLine(2475) Or nNumber = 2 Then
vbwProfiler.vbwExecuteLine 2476
        retval = "Two"
'vbwLine 2477:    ElseIf nNumber = 3 Then
    ElseIf vbwProfiler.vbwExecuteLine(2477) Or nNumber = 3 Then
vbwProfiler.vbwExecuteLine 2478
        retval = "Three"
'vbwLine 2479:    ElseIf nNumber = 4 Then
    ElseIf vbwProfiler.vbwExecuteLine(2479) Or nNumber = 4 Then
vbwProfiler.vbwExecuteLine 2480
        retval = "Four"
'vbwLine 2481:    ElseIf nNumber = 5 Then
    ElseIf vbwProfiler.vbwExecuteLine(2481) Or nNumber = 5 Then
vbwProfiler.vbwExecuteLine 2482
        retval = "Five"
'vbwLine 2483:    ElseIf nNumber = 6 Then
    ElseIf vbwProfiler.vbwExecuteLine(2483) Or nNumber = 6 Then
vbwProfiler.vbwExecuteLine 2484
        retval = "Six"
'vbwLine 2485:    ElseIf nNumber = 7 Then
    ElseIf vbwProfiler.vbwExecuteLine(2485) Or nNumber = 7 Then
vbwProfiler.vbwExecuteLine 2486
        retval = "Seven"
'vbwLine 2487:    ElseIf nNumber = 8 Then
    ElseIf vbwProfiler.vbwExecuteLine(2487) Or nNumber = 8 Then
vbwProfiler.vbwExecuteLine 2488
        retval = "Eight"
'vbwLine 2489:    ElseIf nNumber = 9 Then
    ElseIf vbwProfiler.vbwExecuteLine(2489) Or nNumber = 9 Then
vbwProfiler.vbwExecuteLine 2490
        retval = "Nine"
'vbwLine 2491:    ElseIf nNumber = 10 Then
    ElseIf vbwProfiler.vbwExecuteLine(2491) Or nNumber = 10 Then
vbwProfiler.vbwExecuteLine 2492
        retval = "Ten"
    End If
vbwProfiler.vbwExecuteLine 2493 'B
Else
vbwProfiler.vbwExecuteLine 2494 'B
vbwProfiler.vbwExecuteLine 2495
    retval = "(" + Format(nNumber) + ")"
End If
vbwProfiler.vbwExecuteLine 2496 'B

vbwProfiler.vbwExecuteLine 2497
NumericToString = retval

vbwProfiler.vbwProcOut 84
vbwProfiler.vbwExecuteLine 2498
End Function

Private Function RemoveParenthetical(ByVal strIn As String)
    'JAW 2000.060.26
vbwProfiler.vbwProcIn 85
    Dim varSplit As Variant
    Dim i As Integer
    Dim strTemp As String
    Dim strLocation As String

vbwProfiler.vbwExecuteLine 2499
    varSplit = Split(strIn, "(")
vbwProfiler.vbwExecuteLine 2500
    For i = 1 To UBound(varSplit)
vbwProfiler.vbwExecuteLine 2501
        strTemp = varSplit(i - 1)
vbwProfiler.vbwExecuteLine 2502
        If InStr(1, strTemp, ")") Then
'                If Left(strTemp, 1) = "$" Then
'                strLocation = ""
'            Else
'                strLocation = Left(strTemp, 2)
'            End If
'        varSplit(i - 1) = "[" & strLocation & "]" & Split(strTemp, ")")(1)
vbwProfiler.vbwExecuteLine 2503
            varSplit(i - 1) = Split(strTemp, ")")(1)
        End If
vbwProfiler.vbwExecuteLine 2504 'B
vbwProfiler.vbwExecuteLine 2505
    Next i
vbwProfiler.vbwExecuteLine 2506
    For i = 1 To UBound(varSplit)
vbwProfiler.vbwExecuteLine 2507
        RemoveParenthetical = RemoveParenthetical & varSplit(i - 1)
vbwProfiler.vbwExecuteLine 2508
    Next i

vbwProfiler.vbwProcOut 85
vbwProfiler.vbwExecuteLine 2509
End Function


