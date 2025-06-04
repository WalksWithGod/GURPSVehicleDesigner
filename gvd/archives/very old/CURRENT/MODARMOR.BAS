Attribute VB_Name = "modArmor"
Option Explicit
Public TempTL As Integer ' the temporary tech level of the vehicle
Public ArmorSetting As String
Public TurretOrSuper As Boolean
Dim FinalWeight As Double
Dim FinalCost As Double
Public UserDR As Integer
Public DRPoints As Integer
Public CSide As Integer ' the connected side

Sub LoadArmor(Key As String)

Dim i As Integer ' 1 = complex 0 = basic

CSide = colVehicle.Item(Key).ConnectedSide
frmArmor.Tag = Key
' Load all the data
   With colVehicle.Item(Key)
            frmArmor.Material(0).Text = .FMaterial
            frmArmor.Quality(0).Text = .FQuality
            frmArmor.PD(0).Text = .FPD
            frmArmor.DR(0).Text = .FDR
            frmArmor.Material(1).Text = .BMaterial
            frmArmor.Quality(1).Text = .BQuality
            frmArmor.PD(1).Text = .BPD
            frmArmor.DR(1).Text = .BDR
            frmArmor.Material(2).Text = .LMaterial
            frmArmor.Quality(2).Text = .LQuality
            frmArmor.PD(2).Text = .LPD
            frmArmor.DR(2).Text = .LDR
            frmArmor.Material(3).Text = .RMaterial
            frmArmor.Quality(3).Text = .RQuality
            frmArmor.PD(3).Text = .RPD
            frmArmor.DR(3).Text = .RDR
            frmArmor.Material(4).Text = .TMaterial
            frmArmor.Quality(4).Text = .TQuality
            frmArmor.PD(4).Text = .TPD
            frmArmor.DR(4).Text = .TDR
            frmArmor.Material(5).Text = .UMaterial
            frmArmor.Quality(5).Text = .UQuality
            frmArmor.PD(5).Text = .UPD
            frmArmor.DR(5).Text = .UDR
    End With
'populate all the "quality" combo's with the topmost combo's text
    'For i = 0 To 5
        FillQualityCombo (0)
    'Next
' Gray out the appropriate boxes
    If colVehicle.Item(Key).ArmorOption = "Basic" Then
         If colVehicle.Item(Key).Parent <> "1_" Then
           ' Its a Super or Turret, disable the connected side
            frmArmor.Side(CSide).Value = True
            frmArmor.Side(CSide).Enabled = True
            frmArmor.Material(CSide).Text = "None"
            frmArmor.Material(CSide).Enabled = False
            frmArmor.Quality(CSide).Text = "None"
            frmArmor.Quality(CSide).Enabled = False
            frmArmor.PD(CSide).Text = "0"
            frmArmor.PD(CSide).Enabled = False
            frmArmor.DR(CSide).Text = "0"
            frmArmor.DR(CSide).Enabled = False
            For i = 1 To 5
                If i = CSide Then
                Else
                    frmArmor.Side(i).Value = False
                    frmArmor.Side(i).Enabled = False
                    frmArmor.Material(i).Text = frmArmor.Material(0).Text
                    frmArmor.Material(i).Enabled = False
                    frmArmor.Quality(i).Text = frmArmor.Quality(0).Text
                    frmArmor.Quality(i).Enabled = False
                    frmArmor.PD(i).Text = frmArmor.PD(0).Text
                    frmArmor.PD(i).Enabled = False
                End If
            Next
            ' populate the material combo
            If TempTL <= 6 Then
                    frmArmor.Material(0).AddItem "Wood"
                    frmArmor.Material(0).AddItem "Metal"
                    frmArmor.Material(0).AddItem "Nonrigid"
            Else ' if its greater than or equal to 7
                    frmArmor.Material(0).AddItem "Wood"
                    frmArmor.Material(0).AddItem "Metal"
                    frmArmor.Material(0).AddItem "Ablative"
                    frmArmor.Material(0).AddItem "Nonrigid"
                    frmArmor.Material(0).AddItem "Composite"
                    frmArmor.Material(0).AddItem "Laminate"
            End If
            ' fill the quality combo based on value of the Material(0)
            FillMaterialCombo (0)
        End If
     Else ' We are using Complex Armor Option
            'populate all the "material" combo's
                        ' Fill the Material Combo Boxes according to tech level
            If TempTL <= 6 Then
                    For i = 0 To 5
                        frmArmor.Material(i).AddItem "Wood"
                        frmArmor.Material(i).AddItem "Metal"
                        frmArmor.Material(i).AddItem "Nonrigid"
                    Next
             Else ' if its greater than or equal to 7
                    For i = 0 To 5
                        frmArmor.Material(i).AddItem "Wood"
                        frmArmor.Material(i).AddItem "Metal"
                        frmArmor.Material(i).AddItem "Ablative"
                        frmArmor.Material(i).AddItem "Nonrigid"
                        frmArmor.Material(i).AddItem "Composite"
                        frmArmor.Material(i).AddItem "Laminate"
                    Next
              End If
            ' fill the quality combo's based on whats in the Material's
            For i = 0 To 5
                FillMaterialCombo (i)
            Next
            ' disable just the connected side if Super
            ' or Turret since this is complex option
            If colVehicle.Item(Key).Parent <> "1_" Then
                frmArmor.Side(CSide).Value = True
                frmArmor.Side(CSide).Enabled = False
                frmArmor.Material(CSide).Text = "None"
                frmArmor.Material(CSide).Enabled = False
                frmArmor.Quality(CSide).Text = "None"
                frmArmor.Quality(CSide).Enabled = False
                frmArmor.PD(CSide).Text = "0"
                frmArmor.PD(CSide).Enabled = False
            End If
           ' Set checkOption button to True
           frmArmor.chkOption.Value = 1
    End If
    If colVehicle.Item(Key).Parent = "1_" Then
        For i = 1 To 5
            frmArmor.Side(i).Enabled = False
        Next
        frmArmor.Show vbModal, frmDesigner
    End If
frmArmor.Show vbModal, frmDesigner

End Sub

Sub FillMaterialCombo(Index As Integer)
Dim MaterialCombo As Combobox
Dim Selected As String ' holds the users selected Armor material
Dim arrQuality() As Integer 'holds the list of suitable quality
Dim iSelected As Integer 'holds converted Selected string
Dim element As Variant 'one element of the arrQuality array
Dim i As Integer 'counter
Dim Count As Integer 'another counter
Dim TextFlag As Boolean
Dim TempText As String

Const Cheap = 1
Const Standard = 2
Const Expensive = 3
Const Advanced = 4
Const Fireproof = 8
Const Reflex = 5
Const TL6 = 6
Const TL7 = 7

'get the type of armor that the user selected
Selected = frmArmor.Material(Index).Text
'convert the Selected into an integer
Select Case Selected
    Case "Wood"
        iSelected = 1
    Case "Metal"
        iSelected = 2
    Case "Ablative"
        iSelected = 3
    Case "Nonrigid"
        iSelected = 4
    Case "Composite"
        iSelected = 5
    Case "Laminate"
        iSelected = 6
End Select

Count = 1 ' init the counter
'given the tech level, produce list of quality types
For i = 1 To UBound(ArmorMatrix)
    If ArmorMatrix(i).TL <= TempTL Then
        If ArmorMatrix(i).MaterialType = iSelected Then
            If ArmorMatrix(i).WeightMod <> 0 Then
                ReDim Preserve arrQuality(Count)
                arrQuality(Count) = ArmorMatrix(i).Quality
                Count = Count + 1
                If Count > 5 Then Exit For Else
            End If
        End If
    End If
Next

' Clear the combo box's list
TempText = frmArmor.Quality(Index)
frmArmor.Quality(Index).Clear
TextFlag = False

'fill the Quality combo with the list of available items
frmArmor.Quality(Index).AddItem "None"
For Each element In arrQuality
    Select Case element
        Case Cheap
            frmArmor.Quality(Index).AddItem "Cheap"
            If TempText = "Cheap" Then TextFlag = True
        Case Standard
            frmArmor.Quality(Index).AddItem "Standard"
            If TempText = "Standard" Then TextFlag = True
        Case Expensive
            frmArmor.Quality(Index).AddItem "Expensive"
            If TempText = "Expensive" Then TextFlag = True
        Case Advanced
            frmArmor.Quality(Index).AddItem "Advanced"
            If TempText = "Advanced" Then TextFlag = True
        Case Fireproof
            frmArmor.Quality(Index).AddItem "Fireproof"
            If TempText = "Fireproof" Then TextFlag = True
        Case Reflex
            frmArmor.Quality(Index).AddItem "Reflex"
            If TempText = "Reflex" Then TextFlag = True
        Case TL6
            frmArmor.Quality(Index).AddItem "TL6-"
            If TempText = "TL6" Then TextFlag = True
        Case TL7
            frmArmor.Quality(Index).AddItem "TL7+"
            If TempText = "TL7" Then TextFlag = True
    End Select
Next
' Set the text property of the combo since we clearedit
If TextFlag Then
    frmArmor.Quality(Index) = TempText
Else
    frmArmor.Quality(Index) = "None"
End If

' Fill each Material combo box with the value of the topmost combo
If Index = 0 Then
    If frmArmor.chkOption.Value = False Then
        For Each MaterialCombo In frmArmor.Material()
            If GetCurrentSide = MaterialCombo.Index Then
            Else
                'update the text for each combo
                MaterialCombo.Text = Selected
                ' since this changed, we must update the Quality as well
                FillQualityCombo (0)
            End If
        Next
    End If
End If
End Sub


Sub FillQualityCombo(Index As Integer)
Dim QualityCombo As Combobox
' Fill each combo box with the value of the topmost combo
If Index = 0 Then
    If frmArmor.chkOption.Value = False Then
        For Each QualityCombo In frmArmor.Quality()
            If GetCurrentSide = QualityCombo.Index Then
            Else
                QualityCombo.Text = frmArmor.Quality(Index).Text
            End If
        Next
    End If
End If
End Sub
Function GetCurrentSide()
Dim Face As OptionButton

For Each Face In frmArmor.Side()
    If Face.Value = True Then
        GetCurrentSide = Face.Index
        Exit For
    End If
    Next
End Function


Sub SaveArmor(Key As String)
' This routine calculates the Cost and Weight of the armor
Dim Area As Double 'holds the surface area
Dim CostModifier(0 To 6) As Single
Dim WeightModifier(0 To 6) As Single
Dim iWeight(0 To 6) As Double
Dim iCost(0 To 6) As Double
Dim SelectedMaterial(0 To 6) As String
Dim SelectedQuality(0 To 6) As String
Dim iSelectedQuality(0 To 6) As Integer
Dim iSelectedMaterial(0 To 6) As Integer
Dim Count As Integer
Dim i As Integer

' re-init variables
FinalCost = 0
FinalWeight = 0
' Get the surface area of the vehicle
Area = colVehicle(colVehicle.Item(Key).Parent).SurfaceArea
' There are just two paths, one for Complex and one for Basic
Select Case ArmorSetting
    Case "Complex"
            For Count = 0 To 5
            ' Get the quality and material of the armor
            SelectedMaterial(Count) = frmArmor.Material(Count).Text
            SelectedQuality(Count) = frmArmor.Quality(Count).Text
            'convert the Selected into an integer
            Select Case SelectedMaterial(Count)
                Case "Wood"
                    iSelectedMaterial(Count) = 1
                Case "Metal"
                    iSelectedMaterial(Count) = 2
                Case "Ablative"
                    iSelectedMaterial(Count) = 3
                Case "Nonrigid"
                    iSelectedMaterial(Count) = 4
                Case "Composite"
                    iSelectedMaterial(Count) = 5
                Case "Laminate"
                    iSelectedMaterial(Count) = 6
            End Select
            Select Case SelectedQuality(Count)
                Case "Cheap"
                    iSelectedQuality(Count) = 1
                Case "Standard"
                    iSelectedQuality(Count) = 2
                Case "Expensive"
                    iSelectedQuality(Count) = 3
                Case "Advanced"
                    iSelectedQuality(Count) = 4
                Case "Fireproof"
                    iSelectedQuality(Count) = 8
                Case "Reflex"
                    iSelectedQuality(Count) = 5
                Case "TL6-"
                    iSelectedQuality(Count) = 6
                Case "TL7-"
                    iSelectedQuality(Count) = 7
            End Select
            ' Get the Cost and Weight Modifiers
                For i = 1 To UBound(ArmorMatrix)
                    If ArmorMatrix(i).TL = TempTL Then
                        If ArmorMatrix(i).MaterialType = iSelectedMaterial(Count) Then
                            If ArmorMatrix(i).Quality = iSelectedQuality(Count) Then
                                CostModifier(Count) = ArmorMatrix(i).Cost
                                WeightModifier(Count) = ArmorMatrix(i).WeightMod
                             End If
                         End If
                    End If
                Next
                ' Get the Cost and weight of each face
                iWeight(Count) = Val(frmArmor.DR(Count)) * (Area / 5) * WeightModifier(Count)
                iCost(Count) = iWeight(Count) * CostModifier(Count)
                FinalWeight = FinalWeight + iWeight(Count)
                FinalCost = FinalCost + iCost(Count)
            Next
                ' Get the Final Cost and Final Weight
                FinalWeight = FinalWeight
                FinalCost = FinalCost
                frmArmor.StatusBar1.Panels(1).Text = "Cost: " & FinalCost
                frmArmor.StatusBar1.Panels(2).Text = "Weight: " & FinalWeight
    
    Case "Basic"
        ' Get the quality and material of the armor
        SelectedMaterial(0) = frmArmor.Material(0).Text
        SelectedQuality(0) = frmArmor.Quality(0).Text
        'convert the Selected into an integer
        Select Case SelectedMaterial(0)
            Case "Wood"
                iSelectedMaterial(0) = 1
            Case "Metal"
                iSelectedMaterial(0) = 2
            Case "Ablative"
                iSelectedMaterial(0) = 3
            Case "Nonrigid"
                iSelectedMaterial(0) = 4
            Case "Composite"
                iSelectedMaterial(0) = 5
            Case "Laminate"
                iSelectedMaterial(0) = 6
        End Select
        Select Case SelectedQuality(0)
            Case "Cheap"
                iSelectedQuality(0) = 1
            Case "Standard"
                iSelectedQuality(0) = 2
            Case "Expensive"
                iSelectedQuality(0) = 3
            Case "Advanced"
                iSelectedQuality(0) = 4
            Case "Fireproof"
                iSelectedQuality(0) = 8
            Case "Reflex"
                iSelectedQuality(0) = 5
            Case "TL6-"
                iSelectedQuality(0) = 6
            Case "TL7-"
                iSelectedQuality(0) = 7
        End Select
        ' Get the Cost and Weight Modifiers
            For i = 1 To UBound(ArmorMatrix)
                If ArmorMatrix(i).TL = TempTL Then
                    If ArmorMatrix(i).MaterialType = iSelectedMaterial(0) Then
                        If ArmorMatrix(i).Quality = iSelectedQuality(0) Then
                            CostModifier(0) = ArmorMatrix(i).Cost
                            WeightModifier(0) = ArmorMatrix(i).WeightMod
                         End If
                     End If
                End If
            Next
            ' Get the Final Cost and Final Weight
            FinalWeight = UserDR * Area * WeightModifier(0)
            FinalCost = FinalWeight * CostModifier(0)
            frmArmor.StatusBar1.Panels(1).Text = "Cost: " & FinalCost
            frmArmor.StatusBar1.Panels(2).Text = "Weight: " & FinalWeight
End Select

' Save the armor option in the collection item
colVehicle.Item(frmArmor.Tag).ArmorOption = ArmorSetting
End Sub
Sub CalcLocationArmor(Key)

' This routine calculates the Cost and Weight of the armor
Dim Area As Double 'holds the surface area
Dim CostModifier As Single
Dim WeightModifier As Single
Dim SelectedMaterial As String
Dim SelectedQuality As String
Dim i As Integer ' counter

' Get the surface area of the vehicle
Area = colVehicle.Item(frmArmor.Tag).SurfaceArea

   ' Get the quality and material of the armor
   SelectedMaterial = frmArmor.Material(0).Text
   SelectedQuality = frmArmor.Quality(0).Text
   ' Get the Cost and Weight Modifiers
       For i = 1 To UBound(ArmorMatrix)
           If ArmorMatrix(i).TL = TempTL Then
               If ArmorMatrix(i).MaterialType = SelectedMaterial Then
                   If ArmorMatrix(i).Quality = SelectedQuality Then
                       CostModifier = ArmorMatrix(i).Cost
                       WeightModifier = ArmorMatrix(i).WeightMod
                    End If
                End If
           End If
       Next
           FinalWeight = Val(frmArmor.DR(0).Text) * Area * WeightModifier
           FinalCost = FinalWeight * CostModifier

' Save the armor option in the collection item
colVehicle.Item(Key).ArmorOption = ArmorSetting
' Save the FinalWeight and FinalCost
colVehicle.Item(Key).Weight = FinalWeight
colVehicle.Item(Key).Cost = FinalCost
End Sub
