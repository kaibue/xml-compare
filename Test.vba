Option Explicit

' Structure to hold difference information
Public Type XMLDifference
    DifferenceType As String
    XPath1 As String
    XPath2 As String
    Value1 As String
    Value2 As String
    Description As String
    ElementType As String
End Type

' Main function to compare two S1000D XML objects
Public Function CompareS1000DXML(xml1 As MSXML2.DOMDocument60, xml2 As MSXML2.DOMDocument60, _
                                Optional outputTable As String = "XMLDifferences") As Long
    
    Dim differences() As XMLDifference
    Dim diffCount As Long
    diffCount = 0
    
    ' Validate input XML documents
    If xml1 Is Nothing Or xml2 Is Nothing Then
        MsgBox "One or both XML objects are Nothing", vbCritical
        Exit Function
    End If
    
    If xml1.documentElement Is Nothing Or xml2.documentElement Is Nothing Then
        MsgBox "One or both XML documents have no root element", vbCritical
        Exit Function
    End If
    
    ' Start comparison from root elements
    CompareNodes xml1.documentElement, xml2.documentElement, "", differences, diffCount
    
    ' Create output table with results
    If diffCount > 0 Then
        CreateDifferenceTable differences, diffCount, outputTable
    Else
        MsgBox "No differences found between the XML documents", vbInformation
    End If
    
    CompareS1000DXML = diffCount
End Function

' Recursive function to compare XML nodes deeply
Private Sub CompareNodes(node1 As MSXML2.IXMLDOMNode, node2 As MSXML2.IXMLDOMNode, _
                        currentPath As String, ByRef differences() As XMLDifference, _
                        ByRef diffCount As Long)
    
    Dim xpath As String
    Dim i As Long, j As Long
    Dim found As Boolean
    Dim attr1 As MSXML2.IXMLDOMAttribute
    Dim attr2 As MSXML2.IXMLDOMAttribute
    
    ' Build XPath for current node
    If node1 Is Nothing And node2 Is Nothing Then Exit Sub
    
    If node1 Is Nothing Then
        xpath = currentPath & "/" & node2.nodeName
        AddDifference differences, diffCount, "MISSING_NODE", "", xpath, "", GetNodeValue(node2), _
                     "Node exists in XML2 but missing in XML1", GetS1000DElementType(node2.nodeName)
        Exit Sub
    End If
    
    If node2 Is Nothing Then
        xpath = currentPath & "/" & node1.nodeName
        AddDifference differences, diffCount, "EXTRA_NODE", xpath, "", GetNodeValue(node1), "", _
                     "Node exists in XML1 but missing in XML2", GetS1000DElementType(node1.nodeName)
        Exit Sub
    End If
    
    xpath = currentPath & "/" & node1.nodeName
    
    ' Compare node names
    If node1.nodeName <> node2.nodeName Then
        AddDifference differences, diffCount, "NODE_NAME", xpath, currentPath & "/" & node2.nodeName, _
                     node1.nodeName, node2.nodeName, "Different node names at same position", _
                     GetS1000DElementType(node1.nodeName)
    End If
    
    ' Compare node values (text content)
    Dim value1 As String, value2 As String
    value1 = GetNodeValue(node1)
    value2 = GetNodeValue(node2)
    
    If Trim(value1) <> Trim(value2) Then
        If Len(Trim(value1)) > 0 Or Len(Trim(value2)) > 0 Then
            AddDifference differences, diffCount, "NODE_VALUE", xpath, xpath, value1, value2, _
                         "Different text content in " & GetS1000DDescription(node1.nodeName), _
                         GetS1000DElementType(node1.nodeName)
        End If
    End If
    
    ' Compare attributes
    CompareAttributes node1, node2, xpath, differences, diffCount
    
    ' Compare child nodes
    CompareChildNodes node1, node2, xpath, differences, diffCount
End Sub

' Compare attributes of two nodes
Private Sub CompareAttributes(node1 As MSXML2.IXMLDOMNode, node2 As MSXML2.IXMLDOMNode, _
                             xpath As String, ByRef differences() As XMLDifference, _
                             ByRef diffCount As Long)
    
    Dim attr1 As MSXML2.IXMLDOMAttribute
    Dim attr2 As MSXML2.IXMLDOMAttribute
    Dim i As Long, found As Boolean
    
    ' Check attributes in node1
    If Not node1.Attributes Is Nothing Then
        For i = 0 To node1.Attributes.length - 1
            Set attr1 = node1.Attributes.Item(i)
            Set attr2 = node2.Attributes.getNamedItem(attr1.Name)
            
            If attr2 Is Nothing Then
                AddDifference differences, diffCount, "MISSING_ATTRIBUTE", _
                             xpath & "/@" & attr1.Name, "", attr1.Value, "", _
                             "Attribute '" & attr1.Name & "' missing in XML2", _
                             GetS1000DElementType(node1.nodeName)
            ElseIf attr1.Value <> attr2.Value Then
                AddDifference differences, diffCount, "ATTRIBUTE_VALUE", _
                             xpath & "/@" & attr1.Name, xpath & "/@" & attr2.Name, _
                             attr1.Value, attr2.Value, _
                             "Different value for attribute '" & attr1.Name & "' in " & GetS1000DDescription(node1.nodeName), _
                             GetS1000DElementType(node1.nodeName)
            End If
        Next i
    End If
    
    ' Check for extra attributes in node2
    If Not node2.Attributes Is Nothing Then
        For i = 0 To node2.Attributes.length - 1
            Set attr2 = node2.Attributes.Item(i)
            If node1.Attributes Is Nothing Then
                Set attr1 = Nothing
            Else
                Set attr1 = node1.Attributes.getNamedItem(attr2.Name)
            End If
            
            If attr1 Is Nothing Then
                AddDifference differences, diffCount, "EXTRA_ATTRIBUTE", _
                             "", xpath & "/@" & attr2.Name, "", attr2.Value, _
                             "Attribute '" & attr2.Name & "' exists in XML2 but not in XML1", _
                             GetS1000DElementType(node2.nodeName)
            End If
        Next i
    End If
End Sub

' Compare child nodes of two parent nodes
Private Sub CompareChildNodes(parent1 As MSXML2.IXMLDOMNode, parent2 As MSXML2.IXMLDOMNode, _
                             currentPath As String, ByRef differences() As XMLDifference, _
                             ByRef diffCount As Long)
    
    Dim child1 As MSXML2.IXMLDOMNode
    Dim child2 As MSXML2.IXMLDOMNode
    Dim children1 As MSXML2.IXMLDOMNodeList
    Dim children2 As MSXML2.IXMLDOMNodeList
    Dim i As Long, j As Long
    Dim found As Boolean
    Dim matchedNodes2() As Boolean
    
    Set children1 = parent1.childNodes
    Set children2 = parent2.childNodes
    
    If children2.length > 0 Then
        ReDim matchedNodes2(children2.length - 1)
    End If
    
    ' Compare each child in parent1 with children in parent2
    For i = 0 To children1.length - 1
        Set child1 = children1.Item(i)
        
        ' Skip text nodes that are just whitespace
        If child1.nodeType = NODE_TEXT And Trim(child1.nodeValue) = "" Then
            GoTo NextChild1
        End If
        
        found = False
        
        ' Try to find matching node in parent2
        For j = 0 To children2.length - 1
            Set child2 = children2.Item(j)
            
            ' Skip already matched nodes and whitespace text nodes
            If UBound(matchedNodes2) >= j Then
                If matchedNodes2(j) Then GoTo NextChild2
            End If
            
            If child2.nodeType = NODE_TEXT And Trim(child2.nodeValue) = "" Then
                GoTo NextChild2
            End If
            
            ' Check if nodes match (same name and key attributes for S1000D)
            If NodesMatch(child1, child2) Then
                found = True
                If UBound(matchedNodes2) >= j Then matchedNodes2(j) = True
                CompareNodes child1, child2, currentPath, differences, diffCount
                Exit For
            End If
            
NextChild2:
        Next j
        
        ' If no match found, it's a missing node in XML2
        If Not found Then
            CompareNodes child1, Nothing, currentPath, differences, diffCount
        End If
        
NextChild1:
    Next i
    
    ' Check for extra nodes in parent2
    For j = 0 To children2.length - 1
        If UBound(matchedNodes2) >= j Then
            If Not matchedNodes2(j) Then
                Set child2 = children2.Item(j)
                If Not (child2.nodeType = NODE_TEXT And Trim(child2.nodeValue) = "") Then
                    CompareNodes Nothing, child2, currentPath, differences, diffCount
                End If
            End If
        End If
    Next j
End Sub

' Check if two nodes match based on S1000D logic
Private Function NodesMatch(node1 As MSXML2.IXMLDOMNode, node2 As MSXML2.IXMLDOMNode) As Boolean
    NodesMatch = False
    
    If node1.nodeName <> node2.nodeName Then Exit Function
    
    ' For S1000D, check key identifying attributes
    Dim keyAttrs As Variant
    keyAttrs = Array("id", "infoCode", "dmCode", "pmCode", "applicRefId", "reasonForUpdateRefId")
    
    Dim attr As Variant
    For Each attr In keyAttrs
        Dim val1 As String, val2 As String
        val1 = GetAttributeValue(node1, CStr(attr))
        val2 = GetAttributeValue(node2, CStr(attr))
        
        If val1 <> "" Or val2 <> "" Then
            If val1 <> val2 Then Exit Function
        End If
    Next attr
    
    NodesMatch = True
End Function

' Get attribute value safely
Private Function GetAttributeValue(node As MSXML2.IXMLDOMNode, attrName As String) As String
    GetAttributeValue = ""
    If Not node.Attributes Is Nothing Then
        Dim attr As MSXML2.IXMLDOMAttribute
        Set attr = node.Attributes.getNamedItem(attrName)
        If Not attr Is Nothing Then GetAttributeValue = attr.Value
    End If
End Function

' Get meaningful text content from a node
Private Function GetNodeValue(node As MSXML2.IXMLDOMNode) As String
    If node Is Nothing Then
        GetNodeValue = ""
        Exit Function
    End If
    
    ' For elements with only text content
    If node.childNodes.length = 1 And node.firstChild.nodeType = NODE_TEXT Then
        GetNodeValue = node.firstChild.nodeValue
    ElseIf node.childNodes.length = 0 And node.nodeType = NODE_TEXT Then
        GetNodeValue = node.nodeValue
    Else
        GetNodeValue = ""
    End If
End Function

' Add a difference to the collection
Private Sub AddDifference(ByRef differences() As XMLDifference, ByRef diffCount As Long, _
                         diffType As String, xpath1 As String, xpath2 As String, _
                         value1 As String, value2 As String, description As String, _
                         elementType As String)
    
    diffCount = diffCount + 1
    ReDim Preserve differences(diffCount - 1)
    
    With differences(diffCount - 1)
        .DifferenceType = diffType
        .XPath1 = xpath1
        .XPath2 = xpath2
        .Value1 = Left(value1, 255) ' Truncate for table display
        .Value2 = Left(value2, 255)
        .Description = description
        .ElementType = elementType
    End With
End Sub

' Get S1000D-specific element description
Private Function GetS1000DDescription(nodeName As String) As String
    Select Case UCase(nodeName)
        Case "DMODULE": GetS1000DDescription = "Data Module"
        Case "IDENTANDSTATUSECTION": GetS1000DDescription = "Identification and Status Section"
        Case "DMADDRES": GetS1000DDescription = "Data Module Address"
        Case "DMSTATUS": GetS1000DDescription = "Data Module Status"
        Case "CONTENT": GetS1000DDescription = "Content Section"
        Case "DMDESCR": GetS1000DDescription = "Data Module Description"
        Case "DMTITLE": GetS1000DDescription = "Data Module Title"
        Case "TECHNAME": GetS1000DDescription = "Technical Name"
        Case "INFONAME": GetS1000DDescription = "Information Name"
        Case "PROCEDURALREQUIREMENTSSECTION": GetS1000DDescription = "Procedural Requirements Section"
        Case "MAINFUNCTSECTION": GetS1000DDescription = "Main Function Section"
        Case "LEVELLEDPARA": GetS1000DDescription = "Levelled Paragraph"
        Case "TITLE": GetS1000DDescription = "Title"
        Case "PARA": GetS1000DDescription = "Paragraph"
        Case "STEP1": GetS1000DDescription = "Step 1"
        Case "STEP2": GetS1000DDescription = "Step 2"
        Case "STEP3": GetS1000DDescription = "Step 3"
        Case "WARNING": GetS1000DDescription = "Warning"
        Case "CAUTION": GetS1000DDescription = "Caution"
        Case "NOTE": GetS1000DDescription = "Note"
        Case "TABLE": GetS1000DDescription = "Table"
        Case "GRAPHIC": GetS1000DDescription = "Graphic"
        Case "HOTSPOT": GetS1000DDescription = "Hotspot"
        Case "APPLICABILITY": GetS1000DDescription = "Applicability"
        Case "APPLIC": GetS1000DDescription = "Applicability Reference"
        Case Else: GetS1000DDescription = "Element (" & nodeName & ")"
    End Select
End Function

' Get S1000D element type for categorization
Private Function GetS1000DElementType(nodeName As String) As String
    Select Case UCase(nodeName)
        Case "DMODULE", "PMODULE": GetS1000DElementType = "Module"
        Case "IDENTANDSTATUSECTION", "DMADDRES", "DMSTATUS": GetS1000DElementType = "Identification"
        Case "CONTENT", "PROCEDURALREQUIREMENTSSECTION", "MAINFUNCTSECTION": GetS1000DElementType = "Content Structure"
        Case "LEVELLEDPARA", "PARA", "STEP1", "STEP2", "STEP3": GetS1000DElementType = "Text Content"
        Case "WARNING", "CAUTION", "NOTE": GetS1000DElementType = "Advisory"
        Case "TABLE", "GRAPHIC", "HOTSPOT": GetS1000DElementType = "Media"
        Case "APPLICABILITY", "APPLIC": GetS1000DElementType = "Applicability"
        Case "TITLE", "DMTITLE", "TECHNAME", "INFONAME": GetS1000DElementType = "Title"
        Case Else: GetS1000DElementType = "Other"
    End Select
End Function

' Create Excel table with differences
Private Sub CreateDifferenceTable(differences() As XMLDifference, diffCount As Long, tableName As String)
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim i As Long
    
    ' Create new worksheet or use existing
    On Error Resume Next
    Set ws = ActiveWorkbook.Worksheets(tableName)
    On Error GoTo 0
    
    If ws Is Nothing Then
        Set ws = ActiveWorkbook.Worksheets.Add
        ws.Name = tableName
    Else
        ws.Cells.Clear
    End If
    
    ' Create headers
    With ws
        .Range("A1:H1").Value = Array("Difference Type", "XPath XML1", "XPath XML2", _
                                     "Value XML1", "Value XML2", "Description", _
                                     "Element Type", "Row")
        
        ' Format headers
        With .Range("A1:H1")
            .Font.Bold = True
            .Interior.ColorIndex = 15
            .Borders.Weight = xlThin
        End With
        
        ' Add data
        For i = 0 To diffCount - 1
            .Cells(i + 2, 1).Value = differences(i).DifferenceType
            .Cells(i + 2, 2).Value = differences(i).XPath1
            .Cells(i + 2, 3).Value = differences(i).XPath2
            .Cells(i + 2, 4).Value = differences(i).Value1
            .Cells(i + 2, 5).Value = differences(i).Value2
            .Cells(i + 2, 6).Value = differences(i).Description
            .Cells(i + 2, 7).Value = differences(i).ElementType
            .Cells(i + 2, 8).Value = i + 1
        Next i
        
        ' Create table
        Set tbl = .ListObjects.Add(xlSrcRange, .Range("A1:H" & (diffCount + 1)), , xlYes)
        tbl.Name = "XMLDifferences"
        tbl.TableStyle = "TableStyleMedium2"
        
        ' Auto-fit columns
        .Columns("A:H").AutoFit
        
        ' Add filters
        .Range("A1:H1").AutoFilter
        
        ' Freeze panes
        .Range("A2").Select
        ActiveWindow.FreezePanes = True
    End With
    
    MsgBox diffCount & " differences found and exported to worksheet '" & tableName & "'", vbInformation
End Sub

' Example usage function
Public Sub ExampleUsage()
    Dim xml1 As MSXML2.DOMDocument60
    Dim xml2 As MSXML2.DOMDocument60
    Dim filePath1 As String, filePath2 As String
    
    Set xml1 = New MSXML2.DOMDocument60
    Set xml2 = New MSXML2.DOMDocument60
    
    ' Load your S1000D XML files
    filePath1 = "C:\Path\To\Your\First\S1000D\File.xml"
    filePath2 = "C:\Path\To\Your\Second\S1000D\File.xml"
    
    xml1.Load filePath1
    xml2.Load filePath2
    
    If xml1.parseError.ErrorCode <> 0 Then
        MsgBox "Error loading XML1: " & xml1.parseError.reason
        Exit Sub
    End If
    
    If xml2.parseError.ErrorCode <> 0 Then
        MsgBox "Error loading XML2: " & xml2.parseError.reason
        Exit Sub
    End If
    
    ' Compare the XML files
    Dim diffCount As Long
    diffCount = CompareS1000DXML(xml1, xml2, "S1000D_Differences")
    
    Set xml1 = Nothing
    Set xml2 = Nothing
End Sub
