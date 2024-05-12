Type ComponentCount
    swComponent As SldWorks.Component2
    Count As Double
End Type


Type PathItem
    PathName As String
    ConfName As String
    
    Text As String
    Count As Double
    IsPrinted As Boolean
End Type

Type MainModel
  Paths() As PathItem
  MaxTreeLevel As Integer
  
End Type

Type PartItem
  Designation As String
  Name As String
  Material As String
  Product As String
  Count As Double
  FullCount As Double
End Type

Type CutItem
  Width As Double
  Material As String
  OutFileName As String
End Type


Enum SheetMetalOptions_e
    ExportFlatPatternGeometry = 1
    IncludeHiddenEdges = 2
    ExportBendLines = 4
    IncludeSketches = 8
    MergeCoplanarFaces = 16
    ExportLibraryFeatures = 32
    ExportFormingTools = 64
    ExportBoundingBox = 2048
End Enum


Const SKIP_EXISTING_FILES As Boolean = False

Const OUT_FILENAME_TEMPLATE As String = "<_FileName_>_<_FeatureName_>_<_ConfName_>.dxf"

Const OUT_PATHNAME_TEMPLATE As String = "DXFs\<$CLPRP:Ìàòåðèàë>\<$CLPRP:Òîëùèíà ëèñòîâîãî ìàòåðèàëà>"

Const OUT_NAME_TEMPLATE As String = "DXFs\<$CLPRP:Ìàòåðèàë>\<$CLPRP:Òîëùèíà ëèñòîâîãî ìàòåðèàëà>\<_FileName_>_<_FeatureName_>_<_ConfName_>.dxf"

Const FLAT_PATTERN_OPTIONS As Integer = SheetMetalOptions_e.ExportBendLines + SheetMetalOptions_e.ExportFlatPatternGeometry

Dim swApp As Object
Dim mModel As MainModel

Sub main()

  Set swApp = Application.SldWorks
'try_:
'    On Error GoTo catch_
    Dim swModel As SldWorks.ModelDoc2
    Set swModel = swApp.ActiveDoc
    
    If swModel Is Nothing Then
        Err.Raise vbError, "", "Please open assembly"
    End If
    
    If (Not swModel.GetType() = swDocumentTypes_e.swDocASSEMBLY) Then
        Err.Raise vbError, "", "Please open assembly"
    End If

    Dim swAssy As SldWorks.AssemblyDoc
    Set swAssy = swModel
    
    swAssy.ResolveAllLightWeightComponents True
    
    Dim swComponents As Variant
    swComponents = swAssy.GetComponents(True)
    
    TraverseAssedmbly swModel, swComponents
    
End Sub


Sub TraverseAssedmbly(swRootModel As SldWorks.ModelDoc2, swComponents As Variant)
    Dim swModel As SldWorks.ModelDoc2
    Dim firstPathItem As PathItem
    
    firstPathItem = PathItems_Get(swRootModel, swRootModel.ConfigurationManager.ActiveConfiguration.Name)
    PathItems_Add firstPathItem
    
    Dim swSubComponents() As SldWorks.Component2
    swSubComponents = VariantToComponents(swComponents)
    
    TraverseComponents swSubComponents
End Sub

Sub TraverseComponents(swComponents() As SldWorks.Component2)

    If (Not swComponents) = -1 Then
        Exit Sub
    End If
    
    Dim i As Integer
    Dim swComponent As SldWorks.Component2
    For i = 0 To UBound(swComponents)
        Set swComponent = swComponents(i)
        
    Next i
End Sub

Function VariantToComponents(swComponents As Variant) As SldWorks.Component2
    
    Dim swSubComponents() As SldWorks.Component2
    Dim swSubComponent As SldWorks.Component2
    If Not IsEmpty(swComponents) Then
        Dim i As Integer
        For i = 0 To UBound(swComponents)
            Set swSubComponent = swComponents(i)
            If (Not swSubComponents) = -1 Then
                ReDim swSubComponents(0)
            Else
                ReDim Preserve swSubComponents(UBound(swSubComponents) + 1)
            End If
            
            Set swSubComponents(UBound(swSubComponents)) = swSubComponent
        Next i
    End If
    
    VariantToComponents = swSubComponents
End Function


Function Components_Distinct(swComponents() As SldWorks.Component2) As ComponentCount()
  
End Function

Sub PathItems_Add(pItem As PathItem)
    If (Not mModel.Paths) = -1 Then
        ReDim mModel.Paths(0)
    Else
        ReDim Preserve mModel.Paths(UBound(mModel.Paths) + 1)
    End If
    
    Set mModel.Paths(UBound(mModel.Paths)) = PathItem
End Sub

Sub PathItems_Release()
    If (Not mModel.Paths) = -1 Then
        Exit Sub
    End If
    
    ReDim Preserve mModel.Paths(UBound(mModel.Paths) - 1)
End Sub

Function PathItems_Get(swModel As SldWorks.Model2, ConfName As String) As PathItem
    Dim res_ As PathItem
    
    res_.Count = 1
    res_.Text = GetFileNameWithoutExtension(swModel.GetPathName)
    res_.IsPrinted = False
    
    PathItems_Get = res_
End Function

Function GetPathItem_Text(swModel As SldWorks.Model2, ConfName As String) As String
  GetOutPathItemName = GetFileNameWithoutExtension(swModel.GetPathName)
End Function

Function GetModelPropertyValue(model As SldWorks.ModelDoc2, ConfName As String, prpName As String) As String
    
    Dim prpVal As String
    Dim swCustPrpMgr As SldWorks.CustomPropertyManager
    
    Set swCustPrpMgr = model.Extension.CustomPropertyManager(ConfName)
    prpVal = GetPropertyValue(swCustPrpMgr, prpName)
    
    If prpVal = "" Then
        Set swCustPrpMgr = model.Extension.CustomPropertyManager("")
        prpVal = GetPropertyValue(swCustPrpMgr, prpName)
    End If
    
    GetModelPropertyValue = prpVal
    
End Function

Function GetPropertyValue(custPrpMgr As SldWorks.CustomPropertyManager, prpName As String) As String
    Dim resVal As String
    custPrpMgr.Get2 prpName, "", resVal
    GetPropertyValue = resVal
End Function

Function GetFileNameWithoutExtension(path As String) As String
    GetFileNameWithoutExtension = Mid(path, InStrRev(path, "\") + 1, InStrRev(path, ".") - InStrRev(path, "\") - 1)
End Function
