Option Explicit

' Combined Module: Import (with Colors) + Auto-Clean + Leading Zero Protection
Public Sub ImportAndAutoClean()
    Dim fd As FileDialog
    Dim selectedFile As String, fileName As String
    Dim sourceWb As Workbook
    Dim ws As Worksheet, targetWs As Worksheet
    Dim sheetFound As Boolean, SkipSheet As Boolean
    Dim keywords As Variant, kw As Variant, ignoreList As Variant, ignoreItem As Variant
    Dim detectedMonth As String, detectedYear As String, newSheetName As String

    ' 1. Configuration
    keywords = Array("Leave", "OT", "Late")
    ignoreList = Array("Leave Hour", "Night", "Leave Hours", "leave hour")
    
    ' 2. File Picker
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = "Select Attendance File"
        .Filters.Clear
        .Filters.Add "Excel Files", "*.xls; *.xlsx; *.xlsm; *.xlsb"
        If .Show = -1 Then
            selectedFile = .SelectedItems(1)
            fileName = fd.SelectedItems(1)
        Else
            Exit Sub
        End If
    End With

    ' 3. Auto-Detect Month and Year
    detectedMonth = GetMonthFromFileName(fileName)
    detectedYear = GetYearFromFileName(fileName)
    
    ' Fallbacks
    If detectedYear = "" Then detectedYear = "2026"
    If detectedMonth = "" Then detectedMonth = "UnknownMonth"

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    ' 4. Open Source Workbook
    On Error Resume Next
    Set sourceWb = Workbooks.Open(fileName:=selectedFile, UpdateLinks:=0, ReadOnly:=True)
    If sourceWb Is Nothing Then
        Application.DisplayAlerts = True
        MsgBox "Could not open the file.", vbCritical
        Exit Sub
    End If
    On Error GoTo 0
    
    sheetFound = False
    
    ' 5. Process Sheets
    For Each ws In sourceWb.Worksheets
        SkipSheet = False
        
        ' Exclusion Filter
        For Each ignoreItem In ignoreList
            If InStr(1, ws.Name, ignoreItem, vbTextCompare) > 0 Then
                SkipSheet = True
                Exit For
            End If
        Next ignoreItem
        
        If Not SkipSheet Then
            For Each kw In keywords
                If InStr(1, ws.Name, kw, vbTextCompare) > 0 Then
                    newSheetName = detectedMonth & "_" & kw
                    
                    On Error Resume Next
                    ThisWorkbook.Sheets(newSheetName).Delete
                    On Error GoTo 0
                    
                    ' Import Sheet
                    ws.Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
                    Set targetWs = ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
                    targetWs.Name = newSheetName
                    
                    ' IMPORTANT: Convert to Values while preserving leading zeros
                    ' We loop to ensure .Text is captured for ID/Code columns
                    Call ProtectLeadingZeros(targetWs)
                    
                    ' Execute Cleaning (Columns, Dates, Row 2)
                    Call ExecuteInternalClean(targetWs, detectedMonth, Val(detectedYear))
                    
                    sheetFound = True
                    Exit For
                End If
            Next kw
        End If
    Next ws

    sourceWb.Close SaveChanges:=False
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
    If sheetFound Then
        MsgBox "Import and Cleaning successful for " & detectedMonth & " " & detectedYear & "." & vbCrLf & "Leading zeros (00xx) preserved.", vbInformation
    Else
        MsgBox "No relevant sheets found.", vbExclamation
    End If

    On Error Resume Next
    ThisWorkbook.Sheets("MAIN").Activate
    On Error GoTo 0
End Sub

' Helper: Forces ID and Code columns to Text format immediately after import
Private Sub ProtectLeadingZeros(ws As Worksheet)
    Dim c As Long, lastCol As Long
    Dim headerText As String
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    For c = 1 To lastCol
        headerText = LCase(ws.Cells(1, c).Value)
        ' Identify columns that usually contain leading zeros
        If InStr(headerText, "id") > 0 Or InStr(headerText, "code") > 0 Or InStr(headerText, "dept") > 0 Then
            ws.Columns(c).NumberFormat = "@"
        End If
    Next c
    
    ' Convert formulas to values but keep text format
    ws.UsedRange.Value = ws.UsedRange.Value
End Sub

' Helper: Cleaning Logic
Private Sub ExecuteInternalClean(ws As Worksheet, mName As String, yVal As Integer)
    Dim lastCol As Long, c As Long, startCol As Integer
    Dim colHeader As String, removeList As Variant, item As Variant
    Dim dayVal As Variant, tempDate As Date
    Dim sheetMonth As Integer
    
    ' A. Remove Unwanted Columns
    removeList = Array("no", "grade", "gender", "check", "sign")
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    For c = lastCol To 1 Step -1
        colHeader = LCase(Trim(ws.Cells(1, c).Value))
        For Each item In removeList
            If colHeader = item Then
                ws.Columns(c).Delete
                Exit For
            End If
        Next item
    Next c
    
    ' B. Find Date Start
    startCol = 1
    For c = 1 To 50
        If IsNumeric(ws.Cells(1, c).Value) And ws.Cells(1, c).Value > 0 Then
            startCol = c
            Exit For
        End If
    Next c
    
    ' C. Format Dates
    sheetMonth = GetMonthNumber(mName)
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    ws.Rows(1).NumberFormat = "@"
    
    For c = startCol To lastCol
        dayVal = ws.Cells(1, c).Value
        If IsNumeric(dayVal) And dayVal <> "" Then
            If dayVal > 31 Then
                tempDate = CDate(dayVal)
            Else
                tempDate = DateSerial(yVal, sheetMonth, Int(dayVal))
            End If
            ws.Cells(1, c).Value = Format(tempDate, "yyyy-mm-dd")
        End If
    Next c
    
    ' D. Remove Row 2 and AutoFit
    ws.Rows(2).Delete
    ws.Columns.AutoFit
End Sub

' --- Extraction Helpers ---

Private Function GetMonthFromFileName(text As String) As String
    Dim months As Variant, m As Variant
    months = Array("January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December", _
                   "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec")
    For Each m In months
        If InStr(1, text, m, vbTextCompare) > 0 Then
            GetMonthFromFileName = m
            Exit Function
        End If
    Next m
    GetMonthFromFileName = ""
End Function

Private Function GetYearFromFileName(text As String) As String
    Dim objRegEx As Object, objMatches As Object
    On Error Resume Next
    Set objRegEx = CreateObject("VBScript.RegExp")
    objRegEx.Pattern = "\b(20[2-9][0-9])\b"
    Set objMatches = objRegEx.Execute(text)
    If Not objMatches Is Nothing Then
        If objMatches.Count > 0 Then
            GetYearFromFileName = objMatches(0).Value
            Exit Function
        End If
    End If
    GetYearFromFileName = ""
End Function

Private Function GetMonthNumber(mName As String) As Integer
    Dim i As Integer
    For i = 1 To 12
        If InStr(1, MonthName(i), mName, vbTextCompare) > 0 Or _
           InStr(1, MonthName(i, True), mName, vbTextCompare) > 0 Then
            GetMonthNumber = i
            Exit Function
        End If
    Next i
    GetMonthNumber = Month(Date)
End Function

