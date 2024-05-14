Type ComponentCount
    swComponent As SldWorks.Component2
    count As Integer
End Type


Type PartCacheItem
    PathName As String
    confName As String
    
    IsSheetMetal As Boolean
    IsBad As Boolean
    
    Designation As String
    Name As String
    Material As String
    Product As String
    Width As String
    Razdel As String
    
    OutFileName As String
End Type

Type PathItem
    PathName As String
    confName As String
    
    Text As String
    count As Double
    IsPrinted As Boolean
End Type

Type PartItem
  PathItems() As PathItem
  
  Designation As String
  Name As String
  Material As String
  Product As String
  Width As String
  count As Double
  Razdel As String
End Type

Type MainModel
  PathItems() As PathItem
  Parts() As PartItem
  PartsCache() As PartCacheItem
End Type

Type CutItem
  Width As String
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


Const IS_TEST As Boolean = False

Const INCLUDE_WITHOUT_CUT As Boolean = True

Const SKIP_EXISTING_FILES As Boolean = False

Const SHEETMETAL_WIDTH_PROPERTYNAME As String = "Толщина листового металла"

'Const OUT_FILENAME_TEMPLATE As String = "<_FileName_>_<_FeatureName_>_<_ConfName_>.dxf"

'Const OUT_PATHNAME_TEMPLATE As String = "DXFs\<$CLPRP:Материал>\<$CLPRP:Толщина листового металла>"

Const OUT_NAME_TEMPLATE As String = "DXFs\<$CLPRP:Материал>\<$CLPRP:Толщина листового металла>\<_FileName_>_<_FeatureName_>_<_ConfName_>.dxf"

Const FLAT_PATTERN_OPTIONS As Integer = SheetMetalOptions_e.ExportBendLines + SheetMetalOptions_e.ExportFlatPatternGeometry

Dim swApp As Object
Dim mModel As MainModel

Dim prevS As String


Sub main()

  Set swApp = Application.SldWorks
try_:
    On Error GoTo catch_
    Dim swModel As SldWorks.ModelDoc2
    Set swModel = swApp.ActiveDoc
    
    If swModel Is Nothing Then
        Err.Raise vbError, "", "Please open assembly"
    End If
    
    If (Not swModel.GetType() = swDocumentTypes_e.swDocASSEMBLY) Then
        Err.Raise vbError, "", "Please open assembly"
    End If

    Erase mModel.Parts
    Erase mModel.PathItems
    Erase mModel.PartsCache
    
    Dim swAssy As SldWorks.AssemblyDoc
    Set swAssy = swModel
    
    swAssy.ResolveAllLightWeightComponents True
    
    Dim firstPathItem As PathItem
    
    firstPathItem = PathItems_Get(swModel, swModel.ConfigurationManager.ActiveConfiguration.Name)
    PathItems_Add mModel.PathItems, firstPathItem
    
    prevS = ""
   
    Traverse swModel, 1, swAssy.GetComponents(True)
    
    
    PrintOut mModel
    
    GoTo finally_
catch_:
    swApp.SendMsgToUser2 Err.Description, swMessageBoxIcon_e.swMbStop, swMessageBoxBtn_e.swMbOk
finally_:
End Sub

Sub Traverse(swRootModel As SldWorks.ModelDoc2, count As Integer, subs As Variant)

    Dim componentCounts() As ComponentCount
    componentCounts = Components_Distinct(VariantToComponents(subs))
    
    If (Not componentCounts) = -1 Then
        Exit Sub
    End If
    
    prevS = prevS & " "
    Dim swModel As SldWorks.ModelDoc2
    Dim tmpComponentCount As ComponentCount
    Dim swComponent As SldWorks.Component2
    Dim tmpPathItem As PathItem
    
    Dim i As Integer
    For i = 0 To UBound(componentCounts)
        tmpComponentCount = componentCounts(i)
        Set swComponent = tmpComponentCount.swComponent
        
        Debug.Print prevS & swComponent.Name2 & " |Count=" & tmpComponentCount.count & " Conf=" & swComponent.ReferencedConfiguration
        
        ProcessComponent swRootModel, swComponent, tmpComponentCount.count
        
        tmpPathItem = PathItems_Get(swComponent.GetModelDoc2(), swComponent.ReferencedConfiguration)
        tmpPathItem.count = tmpComponentCount.count
        
        PathItems_Add mModel.PathItems, tmpPathItem

        Traverse swRootModel, tmpComponentCount.count, swComponent.GetChildren()
        
        PathItems_Release mModel.PathItems
    Next i
    
    prevS = Left(prevS, Len(prevS) - 1)
End Sub

Sub ProcessComponent(rootModel As SldWorks.ModelDoc2, swComponent As SldWorks.Component2, count As Integer)

    Dim swModel As SldWorks.ModelDoc2
    Set swModel = swComponent.GetModelDoc2()
    If (swModel.GetType() = swDocumentTypes_e.swDocASSEMBLY) Then
        If (Component_IsSheetMetal(swComponent)) Then
            Err.Raise vbError, "", swModel.GetTitle() & vbCrLf & "Configuration=" & conf & vbCrLf & " assembly with sheet metal not supported"
        End If
    End If
    
    Dim cutOut As CutItem
    Dim cacheItem As PartCacheItem
    If (swModel.GetType() = swDocumentTypes_e.swDocPART) Then
        cacheItem = ProcessPart(rootModel, swComponent)
        If (cacheItem.IsSheetMetal = True) Then
            If (cacheItem.IsBad = False) Or (INCLUDE_WITHOUT_CUT = True) Then
                Part_AddItem cacheItem, count
                'Debug.Print "!!!!"
            End If
        End If
    End If
End Sub

Function ProcessPart(rootModel As SldWorks.ModelDoc2, swComponent As SldWorks.Component2) As PartCacheItem

    Dim cacheIndex As Integer
    cacheIndex = PartCache_IndexOf(swComponent)
    If (cacheIndex <> -1) Then
        ProcessPart = mModel.PartsCache(cacheIndex)
        Exit Function
    End If

    Dim tmpRes As PartCacheItem
    tmpRes = PartCache_Get(rootModel, swComponent)
    
    PartCache_Add mModel.PartsCache, tmpRes
    
    ProcessPart = tmpRes
End Function

Function PartCache_IndexOf(swComponent As SldWorks.Component2) As Integer
    PartCache_IndexOf = -1
    If (Not mModel.PartsCache) = -1 Then
        Exit Function
    End If
    
    Dim i As Integer
    
    Dim tmpItem As PartCacheItem
    Dim tmpPathName As String
    Dim tmpRefConf As String
    
    tmpPathName = swComponent.GetPathName()
    tmpRefConf = swComponent.ReferencedConfiguration
    
    For i = 0 To UBound(mModel.PartsCache)
        tmpItem = mModel.PartsCache(i)
       
        If tmpItem.PathName = tmpPathName And tmpItem.confName = tmpRefConf Then
            PartCache_IndexOf = i
            Exit Function
        End If
    Next
End Function

Sub PartCache_Add(items() As PartCacheItem, item As PartCacheItem)
    If (Not items) = -1 Then
        ReDim items(0)
    Else
        ReDim Preserve items(UBound(items) + 1)
    End If
    
    items(UBound(items)) = item
End Sub

Function PartCache_Get(rootModel As SldWorks.ModelDoc2, swComponent As SldWorks.Component2) As PartCacheItem

    Dim swModel As SldWorks.ModelDoc2
    Dim confName As String
    Dim cutOut As CutItem
 
    Set swModel = swComponent.GetModelDoc2
    confName = swComponent.ReferencedConfiguration
    
    PartCache_Get.PathName = swModel.GetPathName()
    PartCache_Get.confName = confName
    
    PartCache_Get.IsBad = False
    PartCache_Get.IsSheetMetal = False
    
    PartCache_Get.OutFileName = ""
    PartCache_Get.Width = ""
    
    PartCache_Get.Material = GetModelPropertyValue(swModel, confName, "мтех_Наименование_материала")
    PartCache_Get.Designation = GetModelPropertyValue(swModel, confName, "Обозначение")
    PartCache_Get.Name = GetModelPropertyValue(swModel, confName, "Наименование")
    PartCache_Get.Product = GetModelPropertyValue(swModel, confName, "Прибор")
    PartCache_Get.Razdel = GetModelPropertyValue(swModel, confName, "Раздел")
    
try_:
    On Error GoTo catch_
    If (Component_IsSheetMetal(swComponent)) Then
        PartCache_Get.IsSheetMetal = True
    End If
    
    
    If (PartCache_Get.IsSheetMetal = True) Then
        cutOut = SheetMetalModel_Process(rootModel, swModel, confName)
        PartCache_Get.Width = cutOut.Width
        PartCache_Get.OutFileName = cutOut.OutFileName
    End If
    
    GoTo finally_
catch_:
    swApp.SendMsgToUser2 Err.Description, swMessageBoxIcon_e.swMbStop, swMessageBoxBtn_e.swMbOk
    PartCache_Get.IsBad = True
finally_:
    
End Function


Sub PrintOut(outModel As MainModel)
    Dim maxLen As Integer
    Dim tmpMaxLen As Integer
  
    maxLen = -1
    Dim tmpPart As PartItem
    
    Dim i As Integer
    If (Not outModel.Parts) = -1 Then
        Exit Sub
    End If
    For i = 0 To UBound(outModel.Parts)
        tmpPart = outModel.Parts(i)
        tmpMaxLen = UBound(tmpPart.PathItems)
        If (tmpMaxLen > maxLen) Then
            maxLen = tmpMaxLen
        End If
    Next i

  Debug.Print "MaxLen=" & maxLen
  
  Dim tmpExcel As Object
  Dim tmpWorkbook As Object
  Dim tmpWorkSheet As Object
try_:
    On Error GoTo catch_
    
    Set tmpExcel = CreateObject("Excel.Application")
    tmpExcel.Visible = True
    Set tmpWorkbook = tmpExcel.Workbooks.Add
    
    Set tmpWorkSheet = tmpWorkbook.Sheets(1)
    
    Dim recordNumber As Integer
    recordNumber = 0
    
    Dim i1 As Integer
    
    Dim rowIndex As Integer
    Dim colIndex As Integer
    Dim tmpPartItem As PartItem
    Dim fullCount As Double
    Dim tFirstRowIndex As Integer
    Dim tFirstColIndex As Integer
    Dim tLastRowIndex As Integer
    Dim tLastColIndex As Integer
    
    
    fullCount = 1
    
    rowIndex = 3
    colIndex = 1
    tmpWorkSheet.Cells(rowIndex, colIndex).Value = "№ п/п"
    colIndex = colIndex + 1
    For i1 = 0 To maxLen
        tmpWorkSheet.Cells(rowIndex, colIndex).Value = "Сборка"
        colIndex = colIndex + 1
        tmpWorkSheet.Cells(rowIndex, colIndex).Value = "Кол."
        colIndex = colIndex + 1
    Next i1
    tmpWorkSheet.Cells(rowIndex, colIndex).Value = "Номер детали"
    colIndex = colIndex + 1
    tmpWorkSheet.Cells(rowIndex, colIndex).Value = "Наименование"
    colIndex = colIndex + 1
    tmpWorkSheet.Cells(rowIndex, colIndex).Value = "Материал"
    colIndex = colIndex + 1
    tmpWorkSheet.Cells(rowIndex, colIndex).Value = "Применяемость"
    colIndex = colIndex + 1
    tmpWorkSheet.Cells(rowIndex, colIndex).Value = "Примечание"
    colIndex = colIndex + 1
    tmpWorkSheet.Cells(rowIndex, colIndex).Value = "Толщина"
    colIndex = colIndex + 1
    tmpWorkSheet.Cells(rowIndex, colIndex).Value = "Кол-во на комплект"
    colIndex = colIndex + 1
    
    rowIndex = rowIndex + 1
    colIndex = 1
    
    If (UBound(outModel.Parts) > -1) Then
        tFirstRowIndex = rowIndex
        tFirstColIndex = 1
        For i = 0 To UBound(outModel.Parts)
            colIndex = 1
            recordNumber = recordNumber + 1
            tmpWorkSheet.Cells(rowIndex, colIndex).Value = recordNumber
            colIndex = colIndex + 1
            
            tmpPartItem = outModel.Parts(i)
            fullCount = tmpPartItem.count
            For i1 = 0 To maxLen
                If (i1 <= UBound(tmpPartItem.PathItems)) Then
                    tmpWorkSheet.Cells(rowIndex, colIndex).Value = tmpPartItem.PathItems(i1).Text
                    colIndex = colIndex + 1
                    tmpWorkSheet.Cells(rowIndex, colIndex).Value = tmpPartItem.PathItems(i1).count
                    colIndex = colIndex + 1
                    fullCount = fullCount * tmpPartItem.PathItems(i1).count
                Else
                    colIndex = colIndex + 1
                    colIndex = colIndex + 1
                End If
            Next i1
            
            tmpWorkSheet.Cells(rowIndex, colIndex).Value = tmpPartItem.Designation
            colIndex = colIndex + 1
            tmpWorkSheet.Cells(rowIndex, colIndex).Value = tmpPartItem.Name
            colIndex = colIndex + 1
            tmpWorkSheet.Cells(rowIndex, colIndex).Value = tmpPartItem.Material
            colIndex = colIndex + 1
            tmpWorkSheet.Cells(rowIndex, colIndex).Value = tmpPartItem.Product
            colIndex = colIndex + 1
            tmpWorkSheet.Cells(rowIndex, colIndex).Value = ""
            colIndex = colIndex + 1
            tmpWorkSheet.Cells(rowIndex, colIndex).Value = tmpPartItem.Width
            colIndex = colIndex + 1
            
            tmpWorkSheet.Cells(rowIndex, colIndex).Value = fullCount
            colIndex = colIndex + 1
            rowIndex = rowIndex + 1
            
            tLastRowIndex = rowIndex - 1
            tLastColIndex = colIndex - 1
            
        Next i
        tmpWorkSheet.Range(tmpWorkSheet.Cells(tFirstRowIndex, tFirstColIndex), tmpWorkSheet.Cells(tLastRowIndex, tLastColIndex)).Font.Italic = True
        tmpWorkSheet.Range(tmpWorkSheet.Cells(tFirstRowIndex, tFirstColIndex), tmpWorkSheet.Cells(tLastRowIndex, tLastColIndex)).RowHeight = 24
        tmpWorkSheet.Range(tmpWorkSheet.Cells(tFirstRowIndex, tFirstColIndex), tmpWorkSheet.Cells(tLastRowIndex, tLastColIndex)).EntireColumn.AutoFit
        
'        Dim rFirstRowIndex As Integer
'        Dim rLastRowIndex As Integer
'        rFirstRowIndex = tFirstRowIndex
'        rLastRowIndex = tFirstRowIndex
'        For i1 = 0 To maxLen
'            colIndex = i1 + 2
'            rowIndex = tFirstRowIndex
'            rFirstRowIndex = rowIndex
'            rLastRowIndex = rowIndex
'            For i = 0 To UBound(outModel.Parts)
'                rowIndex = rowIndex + i
'                tmpPartItem = outModel.Parts(i)
'                If (i1 <= UBound(tmpPartItem.PathItems)) Then
'                    If (rFirstRowIndex <> rLastRowIndex) Then
'                        tmpWorkSheet.Range(tmpWorkSheet.Cells(rFirstRowIndex, colIndex), tmpWorkSheet.Cells(rLastRowIndex, colIndex)).Merge
'                    End If
'                    rFirstRowIndex = rowIndex
'                    rLastRowIndex = rowIndex
'                Else
'                    If (tmpPartItem.PathItems(i).IsPrinted = False) Then
'                        If (rFirstRowIndex <> rLastRowIndex) Then
'                            tmpWorkSheet.Range(tmpWorkSheet.Cells(rFirstRowIndex, colIndex), tmpWorkSheet.Cells(rLastRowIndex, colIndex)).Merge
'                        End If
'                        rFirstRowIndex = rowIndex
'                        rLastRowIndex = rowIndex
'                    Else
'                        rLastRowIndex = rowIndex
'                    End If
'                End If
'            Next i
'        Next i1
    End If
    
    GoTo finally_
catch_:
    swApp.SendMsgToUser2 Err.Description, swMessageBoxIcon_e.swMbStop, swMessageBoxBtn_e.swMbOk
finally_:

End Sub

Sub Parts_Add(swComponent As SldWorks.Component2, count As Integer, cut As CutItem)
    Dim res As PartItem
    
    Dim swModel As SldWorks.ModelDoc2
    Set swModel = swComponent.GetModelDoc2()
    
    res.Designation = GetModelPropertyValue(swModel, swComponent.ReferencedConfiguration, "Обозначение")
    res.Name = GetModelPropertyValue(swModel, swComponent.ReferencedConfiguration, "Наименование")
    res.Product = GetModelPropertyValue(swModel, swComponent.ReferencedConfiguration, "Прибор")
    res.Width = cut.Width
    res.Material = cut.Material
    res.Razdel = GetModelPropertyValue(swModel, swComponent.ReferencedConfiguration, "Раздел")
    res.count = count
    
    
    
    Dim tmpPathItem As PathItem
    Dim tmpNewPathItem As PathItem
    Dim i As Integer
    For i = 0 To UBound(mModel.PathItems)
        tmpPathItem = mModel.PathItems(i)
        
        tmpNewPathItem.confName = tmpPathItem.confName
        tmpNewPathItem.count = tmpPathItem.count
        tmpNewPathItem.IsPrinted = tmpPathItem.IsPrinted
        tmpNewPathItem.PathName = tmpPathItem.PathName
        tmpNewPathItem.Text = tmpPathItem.Text
        
        PathItems_Add res.PathItems, tmpNewPathItem
       
        mModel.PathItems(i).IsPrinted = True
    Next i
    
    PartItems_Add mModel.Parts, res
    
    
    'GetPropertyValue(CutFeat.CustomPropertyManager, "ìòåõ_Íàèìåíîâàíèå_ìàòåðèàëà")
    
    'Debug.Print "DES=" & res.Designation
    'Debug.Print "Name=" & res.Name
    'Debug.Print "Product=" & res.Product
    'Debug.Print "Width=" & res.Width
    'Debug.Print "Material=" & res.Material
    'Debug.Print "Razdel=" & res.Razdel
    
End Sub

Sub Part_AddItem(cacheItem As PartCacheItem, count As Integer)
    Dim res As PartItem
    
   
    res.Designation = cacheItem.Designation
    res.Name = cacheItem.Name
    res.Product = cacheItem.Product
    res.Width = cacheItem.Width
    res.Material = cacheItem.Material
    res.Razdel = cacheItem.Razdel
    res.count = count
    
    Dim tmpPathItem As PathItem
    Dim tmpNewPathItem As PathItem
    Dim i As Integer
    For i = 0 To UBound(mModel.PathItems)
        tmpPathItem = mModel.PathItems(i)
        
        tmpNewPathItem.confName = tmpPathItem.confName
        tmpNewPathItem.count = tmpPathItem.count
        tmpNewPathItem.IsPrinted = tmpPathItem.IsPrinted
        tmpNewPathItem.PathName = tmpPathItem.PathName
        tmpNewPathItem.Text = tmpPathItem.Text
        
        PathItems_Add res.PathItems, tmpNewPathItem
       
        mModel.PathItems(i).IsPrinted = True
    Next i
    
    PartItems_Add mModel.Parts, res
End Sub

Function Component_IsSheetMetal(comp As SldWorks.Component2) As Boolean
    
    Dim vBodies As Variant
    vBodies = comp.GetBodies3(swBodyType_e.swSolidBody, Empty)
    
    If Not IsEmpty(vBodies) Then
        
        Dim i As Integer
        
        For i = 0 To UBound(vBodies)
            Dim swBody As SldWorks.Body2
            Set swBody = vBodies(i)
            
            If False <> swBody.IsSheetMetal() Then
                Component_IsSheetMetal = True
                Exit Function
            End If
            
        Next
    End If
    
    Component_IsSheetMetal = False
End Function

Function VariantToComponents(source As Variant) As SldWorks.Component2()
    
    Dim swComponents() As SldWorks.Component2
    Dim swComponent As SldWorks.Component2
    If Not IsEmpty(source) Then
        Dim i As Integer
        For i = 0 To UBound(source)
            Set swComponent = source(i)
            If (Not swComponents) = -1 Then
                ReDim swComponents(0)
            Else
                ReDim Preserve swComponents(UBound(swComponents) + 1)
            End If
            
            Set swComponents(UBound(swComponents)) = swComponent
        Next i
    End If
    
    VariantToComponents = swComponents
End Function


Function Components_Distinct(swComponents() As SldWorks.Component2) As ComponentCount()

    If (Not swComponents) = -1 Then
        Exit Function
    End If
    
  Dim res() As ComponentCount
  Dim swComponent As SldWorks.Component2
  Dim tmpComponentCount As ComponentCount
  
    
  Dim i As Integer
  Dim findIndex As Integer
  Dim cName As String
  
  For i = 0 To UBound(swComponents)
    Set swComponent = swComponents(i)
    cName = swComponent.Name2
        
        If (swComponent.IsSuppressed() = False) Then
            findIndex = ComponentCounts_IndexOf(res, swComponent)
            If (findIndex = -1) Then
                findIndex = ComponentCounts_Add(res, swComponent)
            End If
            
            Set res(findIndex).swComponent = swComponent
            res(findIndex).count = res(findIndex).count + 1
        End If
  Next i
 
  Components_Distinct = res
End Function


Function ComponentCounts_IndexOf(componentCounts() As ComponentCount, swComponent As SldWorks.Component2) As Integer
    ComponentCounts_IndexOf = -1
    If (Not componentCounts) = -1 Then
        Exit Function
    End If
    
    Dim i As Integer
    
    Dim tmpComponentCount As ComponentCount
    Dim tmpswComponent As SldWorks.Component2
    
    For i = 0 To UBound(componentCounts)
        tmpComponentCount = componentCounts(i)
        Set tmpswComponent = tmpComponentCount.swComponent
        
        If tmpswComponent.GetPathName() = swComponent.GetPathName() And tmpswComponent.ReferencedConfiguration = swComponent.ReferencedConfiguration Then
            ComponentCounts_IndexOf = i
            Exit Function
        End If
    Next
End Function

Function ComponentCounts_Add(componentCounts() As ComponentCount, swComponent As SldWorks.Component2) As Integer
    If (Not componentCounts) = -1 Then
        ReDim componentCounts(0)
    Else
        ReDim Preserve componentCounts(UBound(componentCounts) + 1)
    End If
    
    Set componentCounts(UBound(componentCounts)).swComponent = swComponent
    componentCounts(UBound(componentCounts)).count = 0
    
    ComponentCounts_Add = UBound(componentCounts)
End Function

Sub PathItems_Add(items() As PathItem, pItem As PathItem)
    If (Not items) = -1 Then
        ReDim items(0)
    Else
        ReDim Preserve items(UBound(items) + 1)
    End If
    
    items(UBound(items)) = pItem
End Sub

Sub PathItems_Release(items() As PathItem)
    If (Not items) = -1 Then
        Exit Sub
    End If
    
    ReDim Preserve items(UBound(items) - 1)
End Sub


Sub PartItems_Add(items() As PartItem, item As PartItem)
    If (Not items) = -1 Then
        ReDim items(0)
    Else
        ReDim Preserve items(UBound(items) + 1)
    End If
    
    items(UBound(items)) = item
End Sub

Sub PartItems_Release(items() As PartItem)
    If (Not items) = -1 Then
        Exit Sub
    End If
    
    ReDim Preserve items(UBound(items) - 1)
End Sub

Function PathItems_Get(swModel As SldWorks.ModelDoc2, confName As String) As PathItem
    Dim res_ As PathItem
    
    res_.count = 1
    res_.Text = GetModelPropertyValue(swModel, confName, "Обозначение")

    If (Trim(res_.Text) = "") Then
        res_.Text = swModel.GetTitle() & "_" & confName
    Else
        res_.Text = res_.Text & " " & GetModelPropertyValue(swModel, confName, "Наименование")
    End If
    
    'GetFileNameWithoutExtension (swModel.GetPathName)
    res_.IsPrinted = False
    
    PathItems_Get = res_
End Function

Function GetModelPropertyValue(model As SldWorks.ModelDoc2, confName As String, prpName As String) As String
    
    Dim prpVal As String
    Dim swCustPrpMgr As SldWorks.CustomPropertyManager
    
    Set swCustPrpMgr = model.Extension.CustomPropertyManager(confName)
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


Function SheetMetalModel_Process(rootModel As SldWorks.ModelDoc2, sheetMetalModel As SldWorks.ModelDoc2, conf As String) As CutItem
  
    Dim vCutListFeats As Variant
    
    Dim ModelTitle As String
    Dim outed As Boolean
    outed = False
    
    vCutListFeats = GetCutListFeatures(sheetMetalModel)
    
    ModelTitle = sheetMetalModel.GetTitle
    If Not IsEmpty(vCutListFeats) Then
        
        Dim vFlatPatternFeats As Variant
        vFlatPatternFeats = GetFlatPatternFeatures(sheetMetalModel)
        
        If Not IsEmpty(vFlatPatternFeats) Then
            
            Dim swProcessedCutListsFeats() As SldWorks.Feature
            
            Dim i As Integer
    
            For i = 0 To UBound(vFlatPatternFeats)
                
                Dim swFlatPatternFeat As SldWorks.Feature
                Dim swFlatPattern As SldWorks.FlatPatternFeatureData
                
                Set swFlatPatternFeat = vFlatPatternFeats(i)
                
                Set swFlatPattern = swFlatPatternFeat.GetDefinition
                
                Dim swFixedEnt As SldWorks.Entity
                
                Set swFixedEnt = swFlatPattern.FixedFace2
                
                Dim swBody As SldWorks.Body2
                
                If Not (swFixedEnt Is Nothing) Then
                    If TypeOf swFixedEnt Is SldWorks.Face2 Then
                        Dim swFixedFace As SldWorks.Face2
                        Set swFixedFace = swFixedEnt
                        Set swBody = swFixedFace.GetBody
                    ElseIf TypeOf swFixedEnt Is SldWorks.Edge Then
                        Dim swFixedEdge As SldWorks.Edge
                        Set swFixedEdge = swFixedEnt
                        Set swBody = swFixedEdge.GetBody
                    ElseIf TypeOf swFixedEnt Is SldWorks.Vertex Then
                        Dim swFixedVert As SldWorks.Vertex
                        Set swFixedVert = swFixedEnt
                        Set swBody = swFixedVert.GetBody
                    End If
                End If
                
                Dim swCutListFeat As SldWorks.Feature
                Set swCutListFeat = FindCutListFeature(vCutListFeats, swBody)
                
                If Not swCutListFeat Is Nothing Then
                    
                    Dim isUnique As Boolean
                                        
                    If (Not swProcessedCutListsFeats) = -1 Then
                        isUnique = True
                    ElseIf Not ContainsSwObject(swProcessedCutListsFeats, swCutListFeat) Then
                        isUnique = True
                    Else
                        isUnique = False
                    End If
                    
                    If isUnique Then
                        
                        If (Not swProcessedCutListsFeats) = -1 Then
                            ReDim swProcessedCutListsFeats(0)
                        Else
                            ReDim Preserve swProcessedCutListsFeats(UBound(swProcessedCutListsFeats) + 1)
                        End If
                        
                        Set swProcessedCutListsFeats(UBound(swProcessedCutListsFeats)) = swCutListFeat
                        
                        Dim OutFileName As String
                        OutFileName = ComposeOutFileName(OUT_NAME_TEMPLATE, rootModel, sheetMetalModel, conf, swFlatPatternFeat, swCutListFeat)
                        
                        If (outed) Then
                            Err.Raise vbError, "", ModelTitle & vbCrLf & "Configuration=" & conf & vbCrLf & "Multiple flat pattern not supported."
                        End If
                        
                        outed = True
                        
                        If Not SKIP_EXISTING_FILES Or Not FileExists(OutFileName) Then
                            If Not IS_TEST Then
                                ExportFlatPattern sheetMetalModel, swFlatPatternFeat, OutFileName, FLAT_PATTERN_OPTIONS, conf
                            End If
                        End If
                        
                        SheetMetalModel_Process.Width = GetPropertyValue(swCutListFeat.CustomPropertyManager, SHEETMETAL_WIDTH_PROPERTYNAME)
                        SheetMetalModel_Process.OutFileName = OutFileName
                    End If
                    
                Else
                    Err.Raise vbError, "", ModelTitle & vbCrLf & "Configuration=" & conf & vbCrLf & "Failed to find cut-list for flat pattern " & swFlatPatternFeat.Name
                End If
                
            Next
            
        Else
            Err.Raise vbError, "", ModelTitle & vbCrLf & "Configuration=" & conf & vbCrLf & "No flat pattern features found."
        End If
        
    Else
        Err.Raise vbError, "", ModelTitle & vbCrLf & "Configuration=" & conf & vbCrLf & "No cut-list items found."
    End If
    
End Function

Function FileExists(filePath As String) As Boolean
    FileExists = Dir(filePath) <> ""
End Function

Function GetCutListFeatures(model As SldWorks.ModelDoc2) As Variant
    GetCutListFeatures = GetFeaturesByType(model, "CutListFolder")
End Function

Function GetFlatPatternFeatures(model As SldWorks.ModelDoc2) As Variant
    GetFlatPatternFeatures = GetFeaturesByType(model, "FlatPattern")
End Function

Function ResolveToken(token As String, rootModel As SldWorks.ModelDoc2, sheetMetalModel As SldWorks.ModelDoc2, conf As String, flatPatternFeat As SldWorks.Feature, cutListFeat As SldWorks.Feature) As String
    
    Const FILE_NAME_TOKEN As String = "_FileName_"
    Const ASSM_FILE_NAME_TOKEN As String = "_AssmFileName_"
    Const FEAT_NAME_TOKEN As String = "_FeatureName_"
    Const CONF_NAME_TOKEN As String = "_ConfName_"
    
    Const PRP_TOKEN As String = "$PRP:"
    Const CUT_LIST_PRP_TOKEN As String = "$CLPRP:"
    Const ASM_PRP_TOKEN As String = "$ASSMPRP:"
    
    Select Case LCase(token)
        Case LCase(FILE_NAME_TOKEN)
            ResolveToken = GetFileNameWithoutExtension(sheetMetalModel.GetPathName)
        Case LCase(FEAT_NAME_TOKEN)
            ResolveToken = flatPatternFeat.Name
        Case LCase(CONF_NAME_TOKEN)
            ResolveToken = conf
        Case LCase(ASSM_FILE_NAME_TOKEN)
            If rootModel.GetPathName() = "" Then
                Err.Raise vbError, "", "Assembly must be saved to use " & ASSM_FILE_NAME_TOKEN
            End If
            ResolveToken = GetFileNameWithoutExtension(rootModel.GetPathName())
        Case Else
            
            Dim prpName As String
                        
            If Left(token, Len(PRP_TOKEN)) = PRP_TOKEN Then
                prpName = Right(token, Len(token) - Len(PRP_TOKEN))
                ResolveToken = GetModelPropertyValue(sheetMetalModel, conf, prpName)
            ElseIf Left(token, Len(ASM_PRP_TOKEN)) = ASM_PRP_TOKEN Then
                prpName = Right(token, Len(token) - Len(ASM_PRP_TOKEN))
                ResolveToken = GetModelPropertyValue(rootModel, rootModel.ConfigurationManager.ActiveConfiguration.Name, prpName)
            ElseIf Left(token, Len(CUT_LIST_PRP_TOKEN)) = CUT_LIST_PRP_TOKEN Then
                prpName = Right(token, Len(token) - Len(CUT_LIST_PRP_TOKEN))
                ResolveToken = GetPropertyValue(cutListFeat.CustomPropertyManager, prpName)
            Else
                Err.Raise vbError, "", "Unrecognized token: " & token
            End If
            
    End Select
    
End Function

Function FindCutListFeature(vCutListFeats As Variant, body As SldWorks.Body2) As SldWorks.Feature
    
    Dim i As Integer
    
    For i = 0 To UBound(vCutListFeats)
        
        Dim swCutListFeat As SldWorks.Feature
        Set swCutListFeat = vCutListFeats(i)
        
        Dim swBodyFolder As SldWorks.BodyFolder
        Set swBodyFolder = swCutListFeat.GetSpecificFeature2
        
        Dim vBodies As Variant
        
        vBodies = swBodyFolder.GetBodies
        
        If ContainsSwObject(vBodies, body) Then
            Set FindCutListFeature = swCutListFeat
        End If
            
    Next
End Function

Function ContainsSwObject(vArr As Variant, obj As Object) As Boolean
    
    If Not IsEmpty(vArr) Then
    
        Dim i As Integer
        
        For i = 0 To UBound(vArr)
            
            Dim swObj As Object
            Set swObj = vArr(i)
            
            If swApp.IsSame(swObj, obj) = swObjectEquality.swObjectSame Then
                ContainsSwObject = True
                Exit Function
            End If
        Next
        
    End If
    
    ContainsSwObject = False
    
End Function

Function GetFeaturesByType(model As SldWorks.ModelDoc2, typeName As String) As Variant
    
    Dim swFeats() As SldWorks.Feature
    
    Dim swFeat As SldWorks.Feature
    
    Set swFeat = model.FirstFeature
    
    Do While Not swFeat Is Nothing
        
        If typeName = "CutListFolder" And swFeat.GetTypeName2() = "SolidBodyFolder" Then
            Dim swBodyFolder As SldWorks.BodyFolder
            Set swBodyFolder = swFeat.GetSpecificFeature2
            swBodyFolder.UpdateCutList
        End If
        
        ProcessFeature swFeat, swFeats, typeName

        Set swFeat = swFeat.GetNextFeature
        
    Loop
    
    If (Not swFeats) = -1 Then
        GetFeaturesByType = Empty
    Else
        GetFeaturesByType = swFeats
    End If
    
End Function

Sub ProcessFeature(thisFeat As SldWorks.Feature, featsArr() As SldWorks.Feature, typeName As String)
    
    If thisFeat.GetTypeName2() = typeName Then
    
        If (Not featsArr) = -1 Then
            ReDim featsArr(0)
            Set featsArr(0) = thisFeat
        Else
            Dim i As Integer
            
            For i = 0 To UBound(featsArr)
                If swApp.IsSame(featsArr(i), thisFeat) = swObjectEquality.swObjectSame Then
                    Exit Sub
                End If
            Next
            
            ReDim Preserve featsArr(UBound(featsArr) + 1)
            Set featsArr(UBound(featsArr)) = thisFeat
        End If
    
    End If
    
    Dim swSubFeat As SldWorks.Feature
    Set swSubFeat = thisFeat.GetFirstSubFeature
        
    While Not swSubFeat Is Nothing
        ProcessFeature swSubFeat, featsArr, typeName
        Set swSubFeat = swSubFeat.GetNextSubFeature
    Wend
        
End Sub

Sub ExportFlatPattern(part As SldWorks.PartDoc, flatPattern As SldWorks.Feature, outFilePath As String, opts As SheetMetalOptions_e, conf As String)
    
    Dim swModel As SldWorks.ModelDoc2
    Set swModel = part
    
    Dim error As ErrObject
    Dim hide As Boolean

    Dim modelPath As String
    Dim partPathName As String
    
    Dim activateDocOption As Long
    Dim activateErrors As Long
try_:
    
    On Error GoTo catch_

    If False = swModel.Visible Then
        hide = True
        swModel.Visible = True
    End If
    
    modelPath = part.GetPathName
    partPathName = swModel.GetPathName()
    
    activateErrors = 0
    swApp.ActivateDoc3 modelPath, False, swRebuildOnActivation_e.swDontRebuildActiveDoc, activateErrors
    
    swModel.FeatureManager.EnableFeatureTree = False
    swModel.FeatureManager.EnableFeatureTreeWindow = False
    swModel.ActiveView.EnableGraphicsUpdate = False
    
    Dim curConf As String
    
    curConf = swModel.ConfigurationManager.ActiveConfiguration.Name
    
    If curConf <> conf Then
        If False = swModel.ShowConfiguration2(conf) Then
            Err.Raise vbError, "", "Failed to activate configuration"
        End If
    End If
    
    Dim outDir As String
    outDir = Left(outFilePath, InStrRev(outFilePath, "\"))
    
    CreateDirectories outDir
    
    
    If modelPath = "" Then
        Err.Raise vbError, "", "Part document must be saved"
    End If
    
    If False <> flatPattern.Select2(False, -1) Then
        If False = part.ExportToDWG2(outFilePath, modelPath, swExportToDWG_e.swExportToDWG_ExportSheetMetal, True, Empty, False, False, opts, Empty) Then
            Err.Raise vbError, "", "Failed to export flat pattern"
        End If
    Else
        Err.Raise vbError, "", "Failed to select flat-pattern"
    End If
    
    swModel.ShowConfiguration2 curConf
    
    GoTo finally_
    
catch_:
    Set error = Err
finally_:

    swModel.FeatureManager.EnableFeatureTree = True
    swModel.FeatureManager.EnableFeatureTreeWindow = True
    swModel.ActiveView.EnableGraphicsUpdate = True
    
    If hide Then
        swApp.CloseDoc swModel.GetTitle
    End If
    
    If Not error Is Nothing Then
        Err.Raise error.Number, error.source, error.Description, error.HelpFile, error.HelpContext
    End If
    
End Sub

Sub CreateDirectories(Path As String)

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If fso.FolderExists(Path) Then
        Exit Sub
    End If

    CreateDirectories fso.GetParentFolderName(Path)
    
    fso.CreateFolder Path
    
End Sub

Function GetFullPath(model As SldWorks.ModelDoc2, Path As String)
    
    GetFullPath = Path
        
    If IsPathRelative(Path) Then
        
        If Left(Path, 1) <> "\" Then
            Path = "\" & Path
        End If
        
        Dim modelPath As String
        Dim modelDir As String
        
        modelPath = model.GetPathName
        
        modelDir = Left(modelPath, InStrRev(modelPath, "\") - 1)
        
        GetFullPath = modelDir & Path
        
    End If
    
End Function

Function IsPathRelative(Path As String)
    IsPathRelative = Mid(Path, 2, 1) <> ":" And Not IsPathUnc(Path)
End Function

Function IsPathUnc(Path As String)
    IsPathUnc = Left(Path, 2) = "\\"
End Function

Function ComposeOutFileName(template As String, rootModel As SldWorks.ModelDoc2, sheetMetalModel As SldWorks.ModelDoc2, conf As String, flatPatternFeat As SldWorks.Feature, cutListFeat As SldWorks.Feature) As String

    Dim regEx As Object
    Set regEx = CreateObject("VBScript.RegExp")
    
    regEx.Global = True
    regEx.IgnoreCase = True
    regEx.Pattern = "<[^>]*>"
    
    Dim regExMatches As Object
    Set regExMatches = regEx.Execute(template)
    
    Dim i As Integer
    
    Dim OutFileName As String
    OutFileName = template
    
    For i = regExMatches.count - 1 To 0 Step -1
        
        Dim regExMatch As Object
        Set regExMatch = regExMatches.item(i)
                    
        Dim tokenName As String
        tokenName = Mid(regExMatch.Value, 2, Len(regExMatch.Value) - 2)
        
        OutFileName = Left(OutFileName, regExMatch.FirstIndex) & ResolveToken(tokenName, rootModel, sheetMetalModel, conf, flatPatternFeat, cutListFeat) & Right(OutFileName, Len(OutFileName) - (regExMatch.FirstIndex + regExMatch.Length))
    Next
    
    ComposeOutFileName = ReplaceInvalidPathSymbols(GetFullPath(rootModel, OutFileName))
    
End Function

Function ReplaceInvalidPathSymbols(Path As String) As String
    
    Const REPLACE_SYMB As String = "_"
    
    Dim res As String
    res = Right(Path, Len(Path) - Len("X:\"))
    
    Dim drive As String
    drive = Left(Path, Len("X:\"))
    
    Dim invalidSymbols As Variant
    invalidSymbols = Array("/", ":", "*", "?", """", "<", ">", "|")
    
    Dim i As Integer
    For i = 0 To UBound(invalidSymbols)
        Dim invalidSymb As String
        invalidSymb = CStr(invalidSymbols(i))
        res = Replace(res, invalidSymb, REPLACE_SYMB)
    Next
    
    ReplaceInvalidPathSymbols = drive + res
    
End Function

Function GetFileNameWithoutExtension(Path As String) As String
    GetFileNameWithoutExtension = Mid(Path, InStrRev(Path, "\") + 1, InStrRev(Path, ".") - InStrRev(Path, "\") - 1)
End Function


