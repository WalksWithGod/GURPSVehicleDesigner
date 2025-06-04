VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsArmor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'local variable(s) to hold property value(s)
Private mvarMaterial As String 'local copy
Private mvarQuality As String 'local copy
Private mvarDR As Integer 'local copy
Private mvarRAP As Boolean 'local copy
Private mvarElectrified As Boolean 'local copy
Private mvarThermal As Boolean 'local copy
Private mvarPR As Integer 'local copy
Private mvarWeight As Double 'local copy
Private mvarCost As Double 'local copy
Private mvarArmorOption As String 'local copy
Private mvarFCost As Double 'local copy
Private mvarBCost As Double 'local copy
Private mvarLCost As Double 'local copy
Private mvarRCost As Double 'local copy
Private mvarTCost As Double 'local copy
Private mvarUCost As Double 'local copy
Private mvarFMaterial As String 'local copy
Private mvarBMaterial As String 'local copy
Private mvarLMaterial As String 'local copy
Private mvarRMaterial As String 'local copy
Private mvarTMaterial As String 'local copy
Private mvarUMaterial As String 'local copy
Private mvarFQuality As String 'local copy
Private mvarBQuality As String 'local copy
Private mvarLQuality As String 'local copy
Private mvarRQuality As String 'local copy
Private mvarTQuality As String 'local copy
Private mvarUQuality As String 'local copy
Private mvarFDR As Integer 'local copy
Private mvarBDR As Integer 'local copy
Private mvarLDR As Integer 'local copy
Private mvarRDR As Integer 'local copy
Private mvarTDR As Integer 'local copy
Private mvarUDR As Integer 'local copy
Private mvarFRAP As Boolean 'local copy
Private mvarBRAP As Boolean 'local copy
Private mvarLRAP As Boolean 'local copy
Private mvarRRAP As Boolean 'local copy
Private mvarTRAP As Boolean 'local copy
Private mvarURAP As Boolean 'local copy
Private mvarFElectrified As Boolean 'local copy
Private mvarBElectrified As Boolean 'local copy
Private mvarLElectrified As Boolean 'local copy
Private mvarRElectrified As Boolean 'local copy
Private mvarTElectrified As Boolean 'local copy
Private mvarUElectrified As Boolean 'local copy
Private mvarFThermal As Boolean 'local copy
Private mvarBThermal As Boolean 'local copy
Private mvarLThermal As Boolean 'local copy
Private mvarRThermal As Boolean 'local copy
Private mvarTThermal As Boolean 'local copy
Private mvarUThermal As Boolean 'local copy
Private mvarFPR As Integer 'local copy
Private mvarBPR As Integer 'local copy
Private mvarTPR As Integer 'local copy
Private mvarRPR As Integer 'local copy
Private mvarLPR As Integer 'local copy
Private mvarUPR As Integer 'local copy
Private mvarFWeight As Double 'local copy
Private mvarBWeight As Double 'local copy
Private mvarLWeight As Double 'local copy
Private mvarRWeight As Double 'local copy
Private mvarTWeight As Double 'local copy
Private mvarUWeight As Double 'local copy
'local variable(s) to hold property value(s)
Private mvarTL As Integer 'local copy
Private mvarText As String 'local copy
Private mvarParent As String 'local copy
Private mvarLocation As String 'local copy
Private mvarKey As String 'local copy
'local variable(s) to hold property value(s)
Private mvarImage As Integer 'local copy
'local variable(s) to hold property value(s)
Private mvarSelectedImage As Integer 'local copy
'local variable(s) to hold property value(s)
Private mvarRadiation As Boolean 'local copy
Private mvarTRadiation As Boolean 'local copy
Private mvarURadiation As Boolean 'local copy
Private mvarFRadiation As Boolean 'local copy
Private mvarBRadiation As Boolean 'local copy
Private mvarLRadiation As Boolean 'local copy
Private mvarRRadiation As Boolean 'local copy
Private mvarCoating As String 'local copy
Private mvarTCoating As String 'local copy
Private mvarUCoating As String 'local copy
Private mvarFCoating As String 'local copy
Private mvarBCoating As String 'local copy
Private mvarLCoating As String 'local copy
Private mvarRCoating As String 'local copy
'local variable(s) to hold property value(s)
Private mvarConnectedSide As Integer 'local copy
Public Property Let ConnectedSide(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ConnectedSide = 5
    mvarConnectedSide = vData
End Property



Public Property Get ConnectedSide() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ConnectedSide
        ConnectedSide = mvarConnectedSide
End Property



Public Property Let RCoating(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.RCoating = 5
    mvarRCoating = vData
End Property


Public Property Get RCoating() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.RCoating
    RCoating = mvarRCoating
End Property



Public Property Let LCoating(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.LCoating = 5
    mvarLCoating = vData
End Property


Public Property Get LCoating() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.LCoating
    LCoating = mvarLCoating
End Property



Public Property Let BCoating(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.BCoating = 5
    mvarBCoating = vData
End Property


Public Property Get BCoating() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.BCoating
    BCoating = mvarBCoating
End Property



Public Property Let FCoating(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FCoating = 5
    mvarFCoating = vData
End Property


Public Property Get FCoating() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FCoating
    FCoating = mvarFCoating
End Property



Public Property Let UCoating(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.UCoating = 5
    mvarUCoating = vData
End Property


Public Property Get UCoating() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.UCoating
    UCoating = mvarUCoating
End Property



Public Property Let TCoating(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.TCoating = 5
    mvarTCoating = vData
End Property


Public Property Get TCoating() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.TCoating
    TCoating = mvarTCoating
End Property



Public Property Let Coating(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Coating = 5
    mvarCoating = vData
End Property


Public Property Get Coating() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Coating
    Coating = mvarCoating
End Property



Public Property Let RRadiation(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.RRadiation = 5
    mvarRRadiation = vData
End Property


Public Property Get RRadiation() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.RRadiation
    RRadiation = mvarRRadiation
End Property



Public Property Let LRadiation(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.LRadiation = 5
    mvarLRadiation = vData
End Property


Public Property Get LRadiation() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.LRadiation
    LRadiation = mvarLRadiation
End Property



Public Property Let BRadiation(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.BRadiation = 5
    mvarBRadiation = vData
End Property


Public Property Get BRadiation() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.BRadiation
    BRadiation = mvarBRadiation
End Property



Public Property Let FRadiation(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FRadiation = 5
    mvarFRadiation = vData
End Property


Public Property Get FRadiation() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FRadiation
    FRadiation = mvarFRadiation
End Property



Public Property Let URadiation(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.URadiation = 5
    mvarURadiation = vData
End Property


Public Property Get URadiation() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.URadiation
    URadiation = mvarURadiation
End Property



Public Property Let TRadiation(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.TRadiation = 5
    mvarTRadiation = vData
End Property


Public Property Get TRadiation() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.TRadiation
    TRadiation = mvarTRadiation
End Property



Public Property Let Radiation(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Radiation = 5
    mvarRadiation = vData
End Property


Public Property Get Radiation() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Radiation
    Radiation = mvarRadiation
End Property



Public Property Let SelectedImage(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SelectedImage = 5
    mvarSelectedImage = vData
End Property


Public Property Get SelectedImage() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SelectedImage
    SelectedImage = mvarSelectedImage
End Property



Public Property Let Image(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Image = 5
    mvarImage = vData
End Property


Public Property Get Image() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Image
    Image = mvarImage
End Property



Public Property Let Key(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.key = 5
    mvarKey = vData
End Property


Public Property Get Key() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.key
    Key = mvarKey
End Property



Public Property Let Location(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Location = 5
    mvarLocation = vData
End Property


Public Property Get Location() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Location
    Location = mvarLocation
End Property



Public Property Let Parent(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Parent = 5
    mvarParent = vData
End Property


Public Property Get Parent() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Parent
    Parent = mvarParent
End Property



Public Property Let Text(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Text = 5
    mvarText = vData
End Property


Public Property Get Text() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Text
    Text = mvarText
End Property



Public Property Let TL(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.TL = 5
    mvarTL = vData
End Property


Public Property Get TL() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.TL
    TL = mvarTL
End Property




Public Property Let ArmorOption(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Option = 5
    mvarArmorOption = vData
End Property


Public Property Get ArmorOption() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Option
    ArmorOption = mvarArmorOption
End Property
Public Property Let UWeight(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.UWeight = 5
    mvarUWeight = vData
End Property


Public Property Get UWeight() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.UWeight
    UWeight = mvarUWeight
End Property



Public Property Let TWeight(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.TWeight = 5
    mvarTWeight = vData
End Property


Public Property Get TWeight() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.TWeight
    TWeight = mvarTWeight
End Property



Public Property Let RWeight(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.RWeight = 5
    mvarRWeight = vData
End Property


Public Property Get RWeight() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.RWeight
    RWeight = mvarRWeight
End Property



Public Property Let LWeight(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.LWeight = 5
    mvarLWeight = vData
End Property


Public Property Get LWeight() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.LWeight
    LWeight = mvarLWeight
End Property



Public Property Let BWeight(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.BWeight = 5
    mvarBWeight = vData
End Property


Public Property Get BWeight() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.BWeight
    BWeight = mvarBWeight
End Property



Public Property Let FWeight(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FWeight = 5
    mvarFWeight = vData
End Property


Public Property Get FWeight() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FWeight
    FWeight = mvarFWeight
End Property



Public Property Let UPD(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.UPR = 5
    mvarUPR = vData
End Property


Public Property Get UPD() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.UPR
    UPD = mvarUPR
End Property



Public Property Let LPD(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.LPR = 5
    mvarLPR = vData
End Property


Public Property Get LPD() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.LPR
    LPD = mvarLPR
End Property



Public Property Let RPD(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.RPR = 5
    mvarRPR = vData
End Property


Public Property Get RPD() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.RPR
    RPD = mvarRPR
End Property



Public Property Let TPD(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.TPR = 5
    mvarTPR = vData
End Property


Public Property Get TPD() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.TPR
    TPD = mvarTPR
End Property



Public Property Let BPD(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.BPR = 5
    mvarBPR = vData
End Property


Public Property Get BPD() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.BPR
    BPD = mvarBPR
End Property



Public Property Let FPD(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FPR = 5
    mvarFPR = vData
End Property


Public Property Get FPD() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FPR
    FPD = mvarFPR
End Property



Public Property Let UThermal(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.UThermal = 5
    mvarUThermal = vData
End Property


Public Property Get UThermal() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.UThermal
    UThermal = mvarUThermal
End Property



Public Property Let TThermal(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.TThermal = 5
    mvarTThermal = vData
End Property


Public Property Get TThermal() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.TThermal
    TThermal = mvarTThermal
End Property



Public Property Let RThermal(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.RThermal = 5
    mvarRThermal = vData
End Property


Public Property Get RThermal() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.RThermal
    RThermal = mvarRThermal
End Property



Public Property Let LThermal(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.LThermal = 5
    mvarLThermal = vData
End Property


Public Property Get LThermal() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.LThermal
    LThermal = mvarLThermal
End Property



Public Property Let BThermal(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.BThermal = 5
    mvarBThermal = vData
End Property


Public Property Get BThermal() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.BThermal
    BThermal = mvarBThermal
End Property



Public Property Let FThermal(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FThermal = 5
    mvarFThermal = vData
End Property


Public Property Get FThermal() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FThermal
    FThermal = mvarFThermal
End Property



Public Property Let UElectrified(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.UElectrified = 5
    mvarUElectrified = vData
End Property


Public Property Get UElectrified() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.UElectrified
    UElectrified = mvarUElectrified
End Property



Public Property Let TElectrified(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.TElectrified = 5
    mvarTElectrified = vData
End Property


Public Property Get TElectrified() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.TElectrified
    TElectrified = mvarTElectrified
End Property



Public Property Let RElectrified(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.RElectrified = 5
    mvarRElectrified = vData
End Property


Public Property Get RElectrified() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.RElectrified
    RElectrified = mvarRElectrified
End Property



Public Property Let LElectrified(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.LElectrified = 5
    mvarLElectrified = vData
End Property


Public Property Get LElectrified() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.LElectrified
    LElectrified = mvarLElectrified
End Property



Public Property Let BElectrified(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.BElectrified = 5
    mvarBElectrified = vData
End Property


Public Property Get BElectrified() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.BElectrified
    BElectrified = mvarBElectrified
End Property



Public Property Let FElectrified(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FElectrified = 5
    mvarFElectrified = vData
End Property


Public Property Get FElectrified() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FElectrified
    FElectrified = mvarFElectrified
End Property



Public Property Let URAP(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.URAP = 5
    mvarURAP = vData
End Property


Public Property Get URAP() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.URAP
    URAP = mvarURAP
End Property



Public Property Let TRAP(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.TRAP = 5
    mvarTRAP = vData
End Property


Public Property Get TRAP() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.TRAP
    TRAP = mvarTRAP
End Property



Public Property Let RRAP(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.RRAP = 5
    mvarRRAP = vData
End Property


Public Property Get RRAP() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.RRAP
    RRAP = mvarRRAP
End Property



Public Property Let LRAP(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.LRAP = 5
    mvarLRAP = vData
End Property


Public Property Get LRAP() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.LRAP
    LRAP = mvarLRAP
End Property



Public Property Let BRAP(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.BRAP = 5
    mvarBRAP = vData
End Property


Public Property Get BRAP() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.BRAP
    BRAP = mvarBRAP
End Property



Public Property Let FRAP(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FRAP = 5
    mvarFRAP = vData
End Property


Public Property Get FRAP() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FRAP
    FRAP = mvarFRAP
End Property



Public Property Let UDR(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.UDR = 5
    mvarUDR = vData
End Property


Public Property Get UDR() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.UDR
    UDR = mvarUDR
End Property



Public Property Let TDR(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.TDR = 5
    mvarTDR = vData
End Property


Public Property Get TDR() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.TDR
    TDR = mvarTDR
End Property



Public Property Let RDR(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.RDR = 5
    mvarRDR = vData
End Property


Public Property Get RDR() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.RDR
    RDR = mvarRDR
End Property



Public Property Let LDR(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.LDR = 5
    mvarLDR = vData
End Property


Public Property Get LDR() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.LDR
    LDR = mvarLDR
End Property



Public Property Let BDR(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.BDR = 5
    mvarBDR = vData
End Property


Public Property Get BDR() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.BDR
    BDR = mvarBDR
End Property



Public Property Let FDR(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FDR = 5
    mvarFDR = vData
End Property


Public Property Get FDR() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FDR
    FDR = mvarFDR
End Property



Public Property Let UQuality(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.UQuality = 5
    mvarUQuality = vData
End Property


Public Property Get UQuality() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.UQuality
    UQuality = mvarUQuality
End Property



Public Property Let TQuality(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.TQuality = 5
    mvarTQuality = vData
End Property


Public Property Get TQuality() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.TQuality
    TQuality = mvarTQuality
End Property



Public Property Let RQuality(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.RQuality = 5
    mvarRQuality = vData
End Property


Public Property Get RQuality() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.RQuality
    RQuality = mvarRQuality
End Property



Public Property Let LQuality(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.LQuality = 5
    mvarLQuality = vData
End Property


Public Property Get LQuality() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.LQuality
    LQuality = mvarLQuality
End Property



Public Property Let BQuality(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.BQuality = 5
    mvarBQuality = vData
End Property


Public Property Get BQuality() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.BQuality
    BQuality = mvarBQuality
End Property



Public Property Let FQuality(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FQuality = 5
    mvarFQuality = vData
End Property


Public Property Get FQuality() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FQuality
    FQuality = mvarFQuality
End Property



Public Property Let UMaterial(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.UMaterial = 5
    mvarUMaterial = vData
End Property


Public Property Get UMaterial() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.UMaterial
    UMaterial = mvarUMaterial
End Property



Public Property Let TMaterial(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.TMaterial = 5
    mvarTMaterial = vData
End Property


Public Property Get TMaterial() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.TMaterial
    TMaterial = mvarTMaterial
End Property



Public Property Let RMaterial(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.RMaterial = 5
    mvarRMaterial = vData
End Property


Public Property Get RMaterial() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.RMaterial
    RMaterial = mvarRMaterial
End Property



Public Property Let LMaterial(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.LMaterial = 5
    mvarLMaterial = vData
End Property


Public Property Get LMaterial() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.LMaterial
    LMaterial = mvarLMaterial
End Property



Public Property Let BMaterial(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.BMaterial = 5
    mvarBMaterial = vData
End Property


Public Property Get BMaterial() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.BMaterial
    BMaterial = mvarBMaterial
End Property



Public Property Let FMaterial(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FMaterial = 5
    mvarFMaterial = vData
End Property


Public Property Get FMaterial() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FMaterial
    FMaterial = mvarFMaterial
End Property



Public Property Let UCost(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.UCost = 5
    mvarUCost = vData
End Property


Public Property Get UCost() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.UCost
    UCost = mvarUCost
End Property



Public Property Let TCost(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.TCost = 5
    mvarTCost = vData
End Property


Public Property Get TCost() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.TCost
    TCost = mvarTCost
End Property



Public Property Let RCost(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.RCost = 5
    mvarRCost = vData
End Property


Public Property Get RCost() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.RCost
    RCost = mvarRCost
End Property



Public Property Let LCost(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.LCost = 5
    mvarLCost = vData
End Property


Public Property Get LCost() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.LCost
    LCost = mvarLCost
End Property



Public Property Let BCost(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.BCost = 5
    mvarBCost = vData
End Property


Public Property Get BCost() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.BCost
    BCost = mvarBCost
End Property



Public Property Let FCost(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FCost = 5
    mvarFCost = vData
End Property


Public Property Get FCost() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FCost
    FCost = mvarFCost
End Property







Public Property Let Cost(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Cost = 5
    mvarCost = vData
End Property


Public Property Get Cost() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Cost
    Cost = mvarCost
End Property



Public Property Let Weight(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Weight = 5
    mvarWeight = vData
End Property


Public Property Get Weight() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Weight
    Weight = mvarWeight
End Property



Public Property Let PR(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.PR = 5
    mvarPR = vData
End Property


Public Property Get PR() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.PR
    PR = mvarPR
End Property



Public Property Let Thermal(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Thermal = 5
    mvarThermal = vData
End Property


Public Property Get Thermal() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Thermal
    Thermal = mvarThermal
End Property



Public Property Let Electrified(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Electrified = 5
    mvarElectrified = vData
End Property


Public Property Get Electrified() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Electrified
    Electrified = mvarElectrified
End Property



Public Property Let RAP(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.RAP = 5
    mvarRAP = vData
End Property


Public Property Get RAP() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.RAP
    RAP = mvarRAP
End Property



Public Property Let DR(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DR = 5
    mvarDR = vData
End Property


Public Property Get DR() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DR
    DR = mvarDR
End Property



Public Property Let Quality(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Quality = 5
    mvarQuality = vData
End Property


Public Property Get Quality() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Quality
    Quality = mvarQuality
End Property



Public Property Let Material(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Material = 5
    mvarMaterial = vData
End Property


Public Property Get Material() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Material
    Material = mvarMaterial
End Property



Private Sub Class_Initialize()
ConnectedSide = 5 'connectedside always inits as the Bottom side
FMaterial = "Wood"
FQuality = "Standard"
ArmorOption = "Location"
Radiaton = False
Coating = False
Thermal = False

FRAP = False
FElectrified = False
FThermal = False
FDR = 1
FPD = 1


End Sub

Public Function PropQuery(ByVal Row As Integer) As Integer
Const Combobox = 1
Const Textbox = 2
Const CommandButton = 3
' clear the Properties form's combobox
frmProperties.Combo1.Clear
    
Select Case ArmorOption
    
    Case "Location"
        Select Case Row
            Case 1
                frmProperties.Combo1.AddItem "Location"
                frmProperties.Combo1.AddItem "Basic"
                frmProperties.Combo1.AddItem "Complex"
                PropQuery = Combobox
            'Case 2
                'frmProperties.Combo1.AddItem "Location"
                'frmProperties.Combo1.AddItem "By Facing"
                'PropQuery = Textbox
            Case 3
                frmProperties.Combo1.AddItem "Wood"
                frmProperties.Combo1.AddItem "Metal"
                frmProperties.Combo1.AddItem "Ablative"
                frmProperties.Combo1.AddItem "Nonrigid"
                frmProperties.Combo1.AddItem "Composite"
                frmProperties.Combo1.AddItem "Laminate"
                PropQuery = Combobox
            Case 4
                frmProperties.Combo1.AddItem "Cheap"
                frmProperties.Combo1.AddItem "Standard"
                frmProperties.Combo1.AddItem "Expensive"
                frmProperties.Combo1.AddItem "Advanced"
                frmProperties.Combo1.AddItem "Fireproof"
                PropQuery = Combobox
            Case 5 To 8
                frmProperties.Combo1.AddItem "True"
                frmProperties.Combo1.AddItem "False"
                PropQuery = Combobox
            Case 2, 10 To 13
                PropQuery = Textbox
        End Select
    Case Else
        Select Case Row
            Case 1
                frmProperties.Combo1.AddItem "Location"
                frmProperties.Combo1.AddItem "Basic"
                frmProperties.Combo1.AddItem "Complex"
                PropQuery = Combobox
            'Case 2
                'frmProperties.Combo1.AddItem "Location"
                'frmProperties.Combo1.AddItem "By Facing"
                'PropQuery = Textbox
            Case 2 To 4
                PropQuery = CommandButton
            Case 5 To 8
                frmProperties.Combo1.AddItem "True"
                frmProperties.Combo1.AddItem "False"
                PropQuery = Combobox
            Case 2, 10 To 13
                PropQuery = Textbox
        End Select
End Select
    
End Function

Public Sub PropUpdate(iRow As Integer, vNewValue As Variant)
Select Case iRow
    Case 1
        ArmorOption = vNewValue
    Case 2
        Material = vNewValue
    Case 3
        Quality = vNewValue
    Case 4
        'Responsive = vNewValue
    Case 5
        'Robotic = vNewValue
    Case 6
        'Biomechanical = vNewValue
    Case 7
        LivingMetal = vNewValue
    Case 10
        ' textbox for DR
    Case 11
        ' textbox for PD
    Case 12
        ' just a textbox for cost
    Case 13
        ' just a textbox for weight
End Select
'StatsUpdate
End Sub


Public Sub StatsUpdate()

    
    ' Calculate the cost
    'mvarCost = BasicDesignCost(mvarSurfaceArea, mvarTL, mvarFrameStrength, mvarMaterials, mvarResponsive, mvarRobotic, mvarBiomechanical, mvarLivingMetal)
    ' Calculate the weight
    'mvarWeight = BasicDesignWeight(mvarSurfaceArea, mvarTL, mvarFrameStrength, mvarMaterials)
    ' Calculate the DR
    'mvarHitPoints = CalcHitPoints(TypeName(Me), mvarFrameStrength, mvarSurfaceArea)
    ' Calculate the PR
End Sub

Public Sub QueryParent()
' if the object has a parent, query it and check to see if
' more stats/property updates are needed for other objects in the collection
If mvarParent <> "" Then
    colVehicle.Item(Parent).StatsUpdate
End If

End Sub

Public Sub QueryChild()
' if the object has children, query it and check to see if
' more stats/property updates are needed for other objects in the collection
' (see the StatusUpdate property for help on checking for childeren.  Can i use that one in place of this?)

End Sub




