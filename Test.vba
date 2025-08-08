Public Function CompareXMLFiles(file1Path As String, file2Path As String) As Boolean
    On Error GoTo ErrorHandler
    
    ' Clear previous results
    DoCmd.RunSQL "DELETE FROM allDifferences"
    
    ' Read and format both XML files
    Dim file1Content As String
    Dim file2Content As String
    Dim file1Lines As Collection
    Dim file2Lines As Collection
    
    file1Content = ReadTextFile(file1Path)
    file2Content = ReadTextFile(file2Path)
    
    If file1Content = "" Or file2Content = "" Then
        MsgBox "Error reading one or both files"
        CompareXMLFiles = False
        Exit Function
    End If
    
    ' Format XML content into lines
    Set file1Lines = FormatXMLToLines(file1Content)
    Set file2Lines = FormatXMLToLines(file2Content)
    
    ' Calculate differences
    Dim differences As Collection
    Set differences = CalculateDifferences(file1Lines, file2Lines)
    
    ' Output differences to table
    Call OutputDifferencesToTable(differences, file1Path, file2Path)
    
    MsgBox "Comparison complete. Found " & differences.Count & " differences."
    CompareXMLFiles = True
    Exit Function
    
ErrorHandler:
    MsgBox "Error in CompareXMLFiles: " & Err.Description
    CompareXMLFiles = False
End Function

Private Function ReadTextFile(filePath As String) As String
    On Error GoTo ErrorHandler
    
    Dim fileNum As Integer
    Dim fileContent As String
    
    fileNum = FreeFile
    Open filePath For Input As #fileNum
    fileContent = Input$(LOF(fileNum), fileNum)
    Close #fileNum
    
    ReadTextFile = fileContent
    Exit Function
    
ErrorHandler:
    If fileNum <> 0 Then Close #fileNum
    ReadTextFile = ""
End Function

Private Function FormatXMLToLines(xmlContent As String) As Collection
    On Error GoTo ErrorHandler
    
    Dim lines As Collection
    Set lines = New Collection
    
    ' Split content by line breaks and clean up
    Dim rawLines As Variant
    Dim i As Long
    Dim line As String
    
    rawLines = Split(xmlContent, vbCrLf)
    If UBound(rawLines) = 0 Then
        rawLines = Split(xmlContent, vbLf)
    End If
    
    For i = 0 To UBound(rawLines)
        line = Trim(rawLines(i))
        If Len(line) > 0 Then
            lines.Add line
        End If
    Next i
    
    Set FormatXMLToLines = lines
    Exit Function
    
ErrorHandler:
    Set FormatXMLToLines = New Collection
End Function

Private Function CalculateDifferences(lines1 As Collection, lines2 As Collection) As Collection
    On Error GoTo ErrorHandler
    
    Dim differences As Collection
    Set differences = New Collection
    
    ' Get Longest Common Subsequence
    Dim lcs As Collection
    Set lcs = GetLCS(lines1, lines2)
    
    Dim i As Long, j As Long, lcsIndex As Long
    i = 1: j = 1: lcsIndex = 1
    
    Do While i <= lines1.Count Or j <= lines2.Count
        If lcsIndex <= lcs.Count And i <= lines1.Count And j <= lines2.Count Then
            If lines1(i) = lcs(lcsIndex) And lines2(j) = lcs(lcsIndex) Then
                ' Unchanged line
                Dim unchangedDiff As DiffItem
                Set unchangedDiff = New DiffItem
                unchangedDiff.DiffType = "unchanged"
                unchangedDiff.LeftLineNum = i
                unchangedDiff.LeftContent = lines1(i)
                unchangedDiff.RightLineNum = j
                unchangedDiff.RightContent = lines2(j)
                differences.Add unchangedDiff
                
                i = i + 1
                j = j + 1
                lcsIndex = lcsIndex + 1
                GoTo ContinueLoop
            End If
        End If
        
        If i <= lines1.Count And (lcsIndex > lcs.Count Or lines1(i) <> lcs(lcsIndex)) Then
            ' Removed line
            Dim removedDiff As DiffItem
            Set removedDiff = New DiffItem
            removedDiff.DiffType = "removed"
            removedDiff.LeftLineNum = i
            removedDiff.LeftContent = lines1(i)
            removedDiff.RightLineNum = 0
            removedDiff.RightContent = ""
            differences.Add removedDiff
            i = i + 1
        ElseIf j <= lines2.Count And (lcsIndex > lcs.Count Or lines2(j) <> lcs(lcsIndex)) Then
            ' Added line
            Dim addedDiff As DiffItem
            Set addedDiff = New DiffItem
            addedDiff.DiffType = "added"
            addedDiff.LeftLineNum = 0
            addedDiff.LeftContent = ""
            addedDiff.RightLineNum = j
            addedDiff.RightContent = lines2(j)
            differences.Add addedDiff
            j = j + 1
        End If
        
ContinueLoop:
    Loop
    
    ' Identify modified lines (consecutive removed/added pairs with high similarity)
    Set differences = IdentifyModifiedLines(differences)
    
    Set CalculateDifferences = differences
    Exit Function
    
ErrorHandler:
    Set CalculateDifferences = New Collection
End Function

Private Function IdentifyModifiedLines(differences As Collection) As Collection
    On Error GoTo ErrorHandler
    
    Dim result As Collection
    Set result = New Collection
    
    Dim i As Long
    i = 1
    
    Do While i <= differences.Count
        If i < differences.Count Then
            Dim currentDiff As DiffItem
            Dim nextDiff As DiffItem
            Set currentDiff = differences(i)
            Set nextDiff = differences(i + 1)
            
            If currentDiff.DiffType = "removed" And nextDiff.DiffType = "added" Then
                ' Check similarity
                Dim similarity As Double
                similarity = CalculateSimilarity(currentDiff.LeftContent, nextDiff.RightContent)
                
                If similarity > 0.5 Then
                    ' Treat as modified line
                    Dim modifiedDiff As DiffItem
                    Set modifiedDiff = New DiffItem
                    modifiedDiff.DiffType = "modified"
                    modifiedDiff.LeftLineNum = currentDiff.LeftLineNum
                    modifiedDiff.LeftContent = currentDiff.LeftContent
                    modifiedDiff.RightLineNum = nextDiff.RightLineNum
                    modifiedDiff.RightContent = nextDiff.RightContent
                    result.Add modifiedDiff
                    i = i + 2
                    GoTo ContinueLoop2
                End If
            End If
        End If
        
        result.Add differences(i)
        i = i + 1
        
ContinueLoop2:
    Loop
    
    Set IdentifyModifiedLines = result
    Exit Function
    
ErrorHandler:
    Set IdentifyModifiedLines = differences
End Function

Private Function CalculateSimilarity(str1 As String, str2 As String) As Double
    On Error GoTo ErrorHandler
    
    Dim longer As String
    Dim shorter As String
    
    If Len(str1) > Len(str2) Then
        longer = str1
        shorter = str2
    Else
        longer = str2
        shorter = str1
    End If
    
    If Len(longer) = 0 Then
        CalculateSimilarity = 1#
        Exit Function
    End If
    
    Dim editDistance As Long
    editDistance = LevenshteinDistance(longer, shorter)
    
    CalculateSimilarity = (Len(longer) - editDistance) / Len(longer)
    Exit Function
    
ErrorHandler:
    CalculateSimilarity = 0
End Function

Private Function LevenshteinDistance(str1 As String, str2 As String) As Long
    On Error GoTo ErrorHandler
    
    Dim len1 As Long, len2 As Long
    len1 = Len(str1)
    len2 = Len(str2)
    
    If len1 = 0 Then
        LevenshteinDistance = len2
        Exit Function
    End If
    If len2 = 0 Then
        LevenshteinDistance = len1
        Exit Function
    End If
    
    Dim matrix() As Long
    ReDim matrix(0 To len2, 0 To len1)
    
    Dim i As Long, j As Long
    
    For i = 0 To len2
        matrix(i, 0) = i
    Next i
    
    For j = 0 To len1
        matrix(0, j) = j
    Next j
    
    For i = 1 To len2
        For j = 1 To len1
            If Mid(str2, i, 1) = Mid(str1, j, 1) Then
                matrix(i, j) = matrix(i - 1, j - 1)
            Else
                matrix(i, j) = Application.Min(Application.Min(matrix(i - 1, j - 1) + 1, matrix(i, j - 1) + 1), matrix(i - 1, j) + 1)
            End If
        Next j
    Next i
    
    LevenshteinDistance = matrix(len2, len1)
    Exit Function
    
ErrorHandler:
    LevenshteinDistance = 999999
End Function

Private Function GetLCS(arr1 As Collection, arr2 As Collection) As Collection
    On Error GoTo ErrorHandler
    
    Dim m As Long, n As Long
    m = arr1.Count
    n = arr2.Count
    
    ' Create DP table
    Dim dp() As Long
    ReDim dp(0 To m, 0 To n)
    
    Dim i As Long, j As Long
    
    For i = 1 To m
        For j = 1 To n
            If arr1(i) = arr2(j) Then
                dp(i, j) = dp(i - 1, j - 1) + 1
            Else
                dp(i, j) = Application.Max(dp(i - 1, j), dp(i, j - 1))
            End If
        Next j
    Next i
    
    ' Reconstruct LCS
    Dim lcs As Collection
    Set lcs = New Collection
    
    i = m
    j = n
    
    Do While i > 0 And j > 0
        If arr1(i) = arr2(j) Then
            lcs.Add arr1(i), , 1  ' Add at beginning
            i = i - 1
            j = j - 1
        ElseIf dp(i - 1, j) > dp(i, j - 1) Then
            i = i - 1
        Else
            j = j - 1
        End If
    Loop
    
    Set GetLCS = lcs
    Exit Function
    
ErrorHandler:
    Set GetLCS = New Collection
End Function

Private Sub OutputDifferencesToTable(differences As Collection, file1Path As String, file2Path As String)
    On Error GoTo ErrorHandler
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    
    Set db = CurrentDb
    Set rs = db.OpenRecordset("allDifferences", dbOpenDynaset)
    
    Dim diff As DiffItem
    Dim i As Long
    
    For i = 1 To differences.Count
        Set diff = differences(i)
        
        rs.AddNew
        rs!ComparisonID = Now() & "_" & i  ' Unique identifier
        rs!File1Path = file1Path
        rs!File2Path = file2Path
        rs!DiffType = diff.DiffType
        rs!LeftLineNumber = IIf(diff.LeftLineNum = 0, Null, diff.LeftLineNum)
        rs!LeftContent = diff.LeftContent
        rs!RightLineNumber = IIf(diff.RightLineNum = 0, Null, diff.RightLineNum)
        rs!RightContent = diff.RightContent
        rs!ComparisonDateTime = Now()
        rs.Update
    Next i
    
    rs.Close
    Set rs = Nothing
    Set db = Nothing
    Exit Sub
    
ErrorHandler:
    If Not rs Is Nothing Then rs.Close
    MsgBox "Error writing to table: " & Err.Description
End Sub

' Class Module: DiffItem
' Create this as a separate class module named "DiffItem"
'
' Public DiffType As String
' Public LeftLineNum As Long
' Public LeftContent As String
' Public RightLineNum As Long
' Public RightContent As String
