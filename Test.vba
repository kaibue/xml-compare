' S1000D XML File Comparison System for MS Access
' This module provides functionality to compare S1000D XML files and record differences

Option Compare Database
Option Explicit

' First, create the comparison results table (run this once)
Sub CreateComparisonTable()
    Dim db As DAO.Database
    Dim tdf As DAO.TableDef
    
    Set db = CurrentDb()
    
    ' Delete table if it exists
    On Error Resume Next
    db.TableDefs.Delete "allDifferences"
    On Error GoTo 0
    
    ' Create new table
    Set tdf = db.CreateTableDef("allDifferences")
    
    With tdf
        .Fields.Append .CreateField("ID", dbLong)
        .Fields("ID").Attributes = dbAutoIncrField
        .Fields.Append .CreateField("ComparisonDate", dbDate)
        .Fields.Append .CreateField("FileName", dbText, 255)
        .Fields.Append .CreateField("Folder1Path", dbText, 255)
        .Fields.Append .CreateField("Folder2Path", dbText, 255)
        .Fields.Append .CreateField("ChangeType", dbText, 50) ' Added, Deleted, Modified, Moved
        .Fields.Append .CreateField("ElementPath", dbText, 500) ' XPath to the element
        .Fields.Append .CreateField("ElementName", dbText, 100)
        .Fields.Append .CreateField("AttributeName", dbText, 100)
        .Fields.Append .CreateField("OldValue", dbMemo)
        .Fields.Append .CreateField("NewValue", dbMemo)
        .Fields.Append .CreateField("S1000DSection", dbText, 100) ' e.g., dmodule, pmEntry, etc.
        .Fields.Append .CreateField("DataModuleCode", dbText, 50) ' DMC if applicable
        .Fields.Append .CreateField("IssueInfo", dbText, 50)
        .Fields.Append .CreateField("SecurityClass", dbText, 20)
        
        ' Create primary key
        .Fields.Append .CreateField("PrimaryKey", dbLong)
        .Fields("PrimaryKey").Attributes = dbAutoIncrField
        
        Dim idx As DAO.Index
        Set idx = .CreateIndex("PrimaryKey")
        idx.Fields = "PrimaryKey"
        idx.Primary = True
        .Indexes.Append idx
    End With
    
    db.TableDefs.Append tdf
    
    MsgBox "allDifferences table created successfully!"
End Sub

' Main comparison function
Public Sub CompareS1000DFiles(folder1Path As String, folder2Path As String, Optional filePattern As String = "*.xml")
    Dim fso As Object
    Dim folder1 As Object
    Dim file As Object
    Dim file1Path As String, file2Path As String
    Dim fileName As String
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If Not fso.FolderExists(folder1Path) Or Not fso.FolderExists(folder2Path) Then
        MsgBox "One or both folder paths do not exist!"
        Exit Sub
    End If
    
    Set folder1 = fso.GetFolder(folder1Path)
    
    ' Loop through all XML files in folder1
    For Each file In folder1.Files
        If file.Name Like filePattern Then
            fileName = file.Name
            file1Path = folder1Path & "\" & fileName
            file2Path = folder2Path & "\" & fileName
            
            If fso.FileExists(file2Path) Then
                ' Compare the two files
                Call CompareS1000DFile(file1Path, file2Path, fileName)
            Else
                ' File only exists in folder1
                Call RecordFileDifference(fileName, folder1Path, folder2Path, "FileOnlyInFolder1", "", "", "", "", "", "")
            End If
        End If
    Next file
    
    ' Check for files that only exist in folder2
    Dim folder2 As Object
    Set folder2 = fso.GetFolder(folder2Path)
    
    For Each file In folder2.Files
        If file.Name Like filePattern Then
            fileName = file.Name
            file1Path = folder1Path & "\" & fileName
            
            If Not fso.FileExists(file1Path) Then
                Call RecordFileDifference(fileName, folder1Path, folder2Path, "FileOnlyInFolder2", "", "", "", "", "", "")
            End If
        End If
    Next file
    
    MsgBox "S1000D file comparison completed! Check the allDifferences table for results."
End Sub

' Compare individual S1000D XML files
Private Sub CompareS1000DFile(file1Path As String, file2Path As String, fileName As String)
    Dim xml1 As Object, xml2 As Object
    Dim doc1 As Object, doc2 As Object
    
    Set xml1 = CreateObject("MSXML2.DOMDocument.6.0")
    Set xml2 = CreateObject("MSXML2.DOMDocument.6.0")
    
    xml1.async = False
    xml2.async = False
    xml1.validateOnParse = False
    xml2.validateOnParse = False
    
    If Not xml1.Load(file1Path) Then
        MsgBox "Error loading " & file1Path & ": " & xml1.parseError.reason
        Exit Sub
    End If
    
    If Not xml2.Load(file2Path) Then
        MsgBox "Error loading " & file2Path & ": " & xml2.parseError.reason
        Exit Sub
    End If
    
    ' Extract S1000D metadata
    Dim dmc1 As String, dmc2 As String
    Dim issue1 As String, issue2 As String
    Dim security1 As String, security2 As String
    
    dmc1 = ExtractDataModuleCode(xml1)
    dmc2 = ExtractDataModuleCode(xml2)
    issue1 = ExtractIssueInfo(xml1)
    issue2 = ExtractIssueInfo(xml2)
    security1 = ExtractSecurityClass(xml1)
    security2 = ExtractSecurityClass(xml2)
    
    ' Compare root elements
    Call CompareNodes(xml1.documentElement, xml2.documentElement, "", fileName, _
                      Left(file1Path, InStrRev(file1Path, "\")), _
                      Left(file2Path, InStrRev(file2Path, "\")), _
                      dmc1, issue1, security1)
End Sub

' Recursive function to compare XML nodes
Private Sub CompareNodes(node1 As Object, node2 As Object, currentPath As String, _
                        fileName As String, folder1 As String, folder2 As String, _
                        dmc As String, issueInfo As String, securityClass As String)
    
    Dim newPath As String
    Dim i As Integer, j As Integer
    Dim found As Boolean
    Dim s1000dSection As String
    
    If node1 Is Nothing And node2 Is Nothing Then Exit Sub
    
    ' Build XPath
    If currentPath = "" Then
        newPath = "/" & node1.nodeName
    Else
        newPath = currentPath & "/" & node1.nodeName
    End If
    
    ' Determine S1000D section
    s1000dSection = GetS1000DSection(node1.nodeName)
    
    ' Compare node existence
    If node1 Is Nothing Then
        Call RecordFileDifference(fileName, folder1, folder2, "Added", newPath, _
                                node2.nodeName, "", "", GetNodeValue(node2), s1000dSection, dmc, issueInfo, securityClass)
        Exit Sub
    End If
    
    If node2 Is Nothing Then
        Call RecordFileDifference(fileName, folder1, folder2, "Deleted", newPath, _
                                node1.nodeName, "", GetNodeValue(node1), "", s1000dSection, dmc, issueInfo, securityClass)
        Exit Sub
    End If
    
    ' Compare node names
    If node1.nodeName <> node2.nodeName Then
        Call RecordFileDifference(fileName, folder1, folder2, "Modified", newPath, _
                                "NodeName", "", node1.nodeName, node2.nodeName, s1000dSection, dmc, issueInfo, securityClass)
    End If
    
    ' Compare attributes
    Call CompareAttributes(node1, node2, newPath, fileName, folder1, folder2, s1000dSection, dmc, issueInfo, securityClass)
    
    ' Compare text content (for leaf nodes)
    If node1.childNodes.Length = 0 And node2.childNodes.Length = 0 Then
        If Trim(node1.Text) <> Trim(node2.Text) Then
            Call RecordFileDifference(fileName, folder1, folder2, "Modified", newPath, _
                                    node1.nodeName, "TextContent", node1.Text, node2.Text, s1000dSection, dmc, issueInfo, securityClass)
        End If
    Else
        ' Compare child nodes
        Dim childMap1 As Object, childMap2 As Object
        Set childMap1 = CreateObject("Scripting.Dictionary")
        Set childMap2 = CreateObject("Scripting.Dictionary")
        
        ' Build maps of child nodes
        For i = 0 To node1.childNodes.Length - 1
            If node1.childNodes(i).nodeType = 1 Then ' Element node
                Dim key1 As String
                key1 = GetNodeKey(node1.childNodes(i))
                If Not childMap1.exists(key1) Then
                    Set childMap1(key1) = CreateObject("Scripting.Dictionary")
                End If
                childMap1(key1)(childMap1(key1).Count) = node1.childNodes(i)
            End If
        Next
        
        For i = 0 To node2.childNodes.Length - 1
            If node2.childNodes(i).nodeType = 1 Then ' Element node
                Dim key2 As String
                key2 = GetNodeKey(node2.childNodes(i))
                If Not childMap2.exists(key2) Then
                    Set childMap2(key2) = CreateObject("Scripting.Dictionary")
                End If
                childMap2(key2)(childMap2(key2).Count) = node2.childNodes(i)
            End If
        Next
        
        ' Compare child nodes
        Dim key As Variant
        For Each key In childMap1.Keys
            If childMap2.exists(key) Then
                ' Compare matching children
                Dim maxCount As Integer
                maxCount = Application.Max(childMap1(key).Count, childMap2(key).Count)
                For i = 0 To maxCount - 1
                    Dim child1 As Object, child2 As Object
                    Set child1 = Nothing
                    Set child2 = Nothing
                    
                    If i < childMap1(key).Count Then Set child1 = childMap1(key)(i)
                    If i < childMap2(key).Count Then Set child2 = childMap2(key)(i)
                    
                    Call CompareNodes(child1, child2, newPath, fileName, folder1, folder2, dmc, issueInfo, securityClass)
                Next
            Else
                ' Child only exists in node1
                For i = 0 To childMap1(key).Count - 1
                    Call CompareNodes(childMap1(key)(i), Nothing, newPath, fileName, folder1, folder2, dmc, issueInfo, securityClass)
                Next
            End If
        Next
        
        ' Check for children that only exist in node2
        For Each key In childMap2.Keys
            If Not childMap1.exists(key) Then
                For i = 0 To childMap2(key).Count - 1
                    Call CompareNodes(Nothing, childMap2(key)(i), newPath, fileName, folder1, folder2, dmc, issueInfo, securityClass)
                Next
            End If
        Next
    End If
End Sub

' Compare attributes between two nodes
Private Sub CompareAttributes(node1 As Object, node2 As Object, elementPath As String, _
                             fileName As String, folder1 As String, folder2 As String, _
                             s1000dSection As String, dmc As String, issueInfo As String, securityClass As String)
    
    Dim attr1 As Object, attr2 As Object
    Dim attrMap1 As Object, attrMap2 As Object
    Dim i As Integer
    
    Set attrMap1 = CreateObject("Scripting.Dictionary")
    Set attrMap2 = CreateObject("Scripting.Dictionary")
    
    ' Build attribute maps
    If Not node1.Attributes Is Nothing Then
        For i = 0 To node1.Attributes.Length - 1
            Set attr1 = node1.Attributes(i)
            attrMap1(attr1.Name) = attr1.Value
        Next
    End If
    
    If Not node2.Attributes Is Nothing Then
        For i = 0 To node2.Attributes.Length - 1
            Set attr2 = node2.Attributes(i)
            attrMap2(attr2.Name) = attr2.Value
        Next
    End If
    
    ' Compare attributes
    Dim attrName As Variant
    For Each attrName In attrMap1.Keys
        If attrMap2.exists(attrName) Then
            If attrMap1(attrName) <> attrMap2(attrName) Then
                Call RecordFileDifference(fileName, folder1, folder2, "Modified", elementPath, _
                                        node1.nodeName, CStr(attrName), attrMap1(attrName), attrMap2(attrName), _
                                        s1000dSection, dmc, issueInfo, securityClass)
            End If
        Else
            Call RecordFileDifference(fileName, folder1, folder2, "Deleted", elementPath, _
                                    node1.nodeName, CStr(attrName), attrMap1(attrName), "", _
                                    s1000dSection, dmc, issueInfo, securityClass)
        End If
    Next
    
    ' Check for new attributes in node2
    For Each attrName In attrMap2.Keys
        If Not attrMap1.exists(attrName) Then
            Call RecordFileDifference(fileName, folder1, folder2, "Added", elementPath, _
                                    node2.nodeName, CStr(attrName), "", attrMap2(attrName), _
                                    s1000dSection, dmc, issueInfo, securityClass)
        End If
    Next
End Sub

' Record a difference in the database
Private Sub RecordFileDifference(fileName As String, folder1 As String, folder2 As String, _
                                changeType As String, elementPath As String, elementName As String, _
                                attributeName As String, oldValue As String, newValue As String, _
                                s1000dSection As String, Optional dmc As String = "", _
                                Optional issueInfo As String = "", Optional securityClass As String = "")
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    
    Set db = CurrentDb()
    Set rs = db.OpenRecordset("allDifferences")
    
    rs.AddNew
    rs!ComparisonDate = Now()
    rs!fileName = fileName
    rs!Folder1Path = folder1
    rs!Folder2Path = folder2
    rs!ChangeType = changeType
    rs!ElementPath = elementPath
    rs!ElementName = elementName
    rs!AttributeName = attributeName
    rs!OldValue = Left(oldValue, 65000) ' Memo field limit
    rs!NewValue = Left(newValue, 65000)
    rs!S1000DSection = s1000dSection
    rs!DataModuleCode = dmc
    rs!IssueInfo = issueInfo
    rs!SecurityClass = securityClass
    rs.Update
    
    rs.Close
End Sub

' Helper functions for S1000D specific parsing

Private Function ExtractDataModuleCode(xmlDoc As Object) As String
    Dim node As Object
    Set node = xmlDoc.SelectSingleNode("//dmIdent/dmCode")
    If Not node Is Nothing Then
        ExtractDataModuleCode = GetDMCString(node)
    Else
        ExtractDataModuleCode = ""
    End If
End Function

Private Function ExtractIssueInfo(xmlDoc As Object) As String
    Dim node As Object
    Set node = xmlDoc.SelectSingleNode("//dmIdent/issueInfo")
    If Not node Is Nothing Then
        ExtractIssueInfo = node.getAttribute("issueNumber") & "-" & node.getAttribute("inWork")
    Else
        ExtractIssueInfo = ""
    End If
End Function

Private Function ExtractSecurityClass(xmlDoc As Object) As String
    Dim node As Object
    Set node = xmlDoc.SelectSingleNode("//dmStatus/security")
    If Not node Is Nothing Then
        ExtractSecurityClass = node.getAttribute("securityClassification")
    Else
        ExtractSecurityClass = ""
    End If
End Function

Private Function GetDMCString(dmCodeNode As Object) As String
    If dmCodeNode Is Nothing Then
        GetDMCString = ""
        Exit Function
    End If
    
    GetDMCString = dmCodeNode.getAttribute("modelIdentCode") & "-" & _
                   dmCodeNode.getAttribute("systemDiffCode") & "-" & _
                   dmCodeNode.getAttribute("systemCode") & "-" & _
                   dmCodeNode.getAttribute("subSystemCode") & _
                   dmCodeNode.getAttribute("subSubSystemCode") & "-" & _
                   dmCodeNode.getAttribute("assyCode") & "-" & _
                   dmCodeNode.getAttribute("disassyCode") & _
                   dmCodeNode.getAttribute("disassyCodeVariant") & "-" & _
                   dmCodeNode.getAttribute("infoCode") & _
                   dmCodeNode.getAttribute("infoCodeVariant") & "-" & _
                   dmCodeNode.getAttribute("itemLocationCode") & _
                   dmCodeNode.getAttribute("learnCode") & _
                   dmCodeNode.getAttribute("learnEventCode")
End Function

Private Function GetS1000DSection(nodeName As String) As String
    Select Case LCase(nodeName)
        Case "dmodule"
            GetS1000DSection = "Data Module"
        Case "pmentry"
            GetS1000DSection = "Publication Module"
        Case "dmlentry"
            GetS1000DSection = "Data Management List"
        Case "scormcontentpackage"
            GetS1000DSection = "SCORM Content Package"
        Case "identandstatusection"
            GetS1000DSection = "Identification and Status"
        Case "content"
            GetS1000DSection = "Content"
        Case "procedure"
            GetS1000DSection = "Procedure"
        Case "description"
            GetS1000DSection = "Description"
        Case "fault"
            GetS1000DSection = "Fault"
        Case "crew"
            GetS1000DSection = "Crew"
        Case "frontmatter"
            GetS1000DSection = "Front Matter"
        Case Else
            GetS1000DSection = "Other"
    End Select
End Function

Private Function GetNodeKey(node As Object) As String
    ' Create a unique key for the node based on S1000D conventions
    Dim key As String
    key = node.nodeName
    
    ' Add important identifying attributes for S1000D elements
    Select Case LCase(node.nodeName)
        Case "dmref", "pmref"
            If Not node.SelectSingleNode("dmRefIdent/dmCode") Is Nothing Then
                key = key & "_" & GetDMCString(node.SelectSingleNode("dmRefIdent/dmCode"))
            End If
        Case "step"
            If node.hasAttribute("id") Then
                key = key & "_" & node.getAttribute("id")
            End If
        Case "para"
            If node.hasAttribute("id") Then
                key = key & "_" & node.getAttribute("id")
            End If
        Case "figure", "table"
            If node.hasAttribute("id") Then
                key = key & "_" & node.getAttribute("id")
            End If
    End Select
    
    GetNodeKey = key
End Function

Private Function GetNodeValue(node As Object) As String
    If node.hasChildNodes Then
        GetNodeValue = Left(node.xml, 500) ' Truncate for storage
    Else
        GetNodeValue = node.Text
    End If
End Function

' Utility function to start comparison with folder selection
Public Sub SelectFoldersAndCompare()
    Dim folder1 As String, folder2 As String
    
    ' You can replace these with folder picker dialogs
    folder1 = InputBox("Enter path to first folder:", "Folder 1", "C:\S1000D\Folder1")
    If folder1 = "" Then Exit Sub
    
    folder2 = InputBox("Enter path to second folder:", "Folder 2", "C:\S1000D\Folder2")
    If folder2 = "" Then Exit Sub
    
    Call CompareS1000DFiles(folder1, folder2)
End Sub

# Query to view results
Public Sub ViewComparisonResults()
    DoCmd.OpenQuery "SELECT * FROM allDifferences ORDER BY FileName, ElementPath"
End Sub
