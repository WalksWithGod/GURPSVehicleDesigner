Attribute VB_Name = "modPerformance"
Option Explicit


Public Function CalcTotalContragravLift()
Dim element As Object
Dim Templift As Double

For Each element In colVehicle
    If TypeOf element Is clsContraGravGenerator Then
        Templift = element.Lift
    End If
Next
CalcTotalContragravLift = Templift
End Function
