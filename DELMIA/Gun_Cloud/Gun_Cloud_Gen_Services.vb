
Sub CATMain()
  Dim currDoc As Document
  Dim rootProd As Product
  Dim toolTag As Tag
  
  ' Get hold of root product
  Set currDoc = CATIA.ActiveDocument
  Set rootProd = currDoc.Product
  
  Set toolTag = getToolTag(rootProd)
  If toolTag Is Nothing Then
    MsgBox "No Tool Tag defined in Gun", vbOKOnly, "ERROR"
  End If
End Sub

Function getToolTag(iProd As Product) As Product
  Dim gunDoc As Document
  Dim docSelection As Selection
  Dim ii As Integer
  Dim selCnt As Integer
  Dim typeStr As String
  Dim toolTag As Tag
  
  Set toolTag = Nothing
  Set getToolTag = Nothing
  Set gunDoc = iProd.ReferenceProduct.Parent
  Set docSelection = gunDoc.Selection
  ' Search for all the Frames of interest in the Gun product
  docSelection.Search ("Name='Frames Of Interest*',all")
  If docSelection.Count2 > 0 Then
    selCnt = docSelection.Count2
    For ii = 1 To selCnt
        On Error Resume Next
            ' Type cast the object to Tag to isolate tags
            ' This will eliminate Tag groups which are also collected by query
            Set toolTag = docSelection.Item2(ii).Value
            If Not toolTag Is Nothing Then
                toolTag.GetType typeStr
                ' Checks if tag type is tool
                ' Search stops if yes as Gun is supposed to have only one Tool Tag
                If typeStr = "Tool" Then
                    Set getToolTag = toolTag
                End If
            End If
    Next
  End If
  
End Function
