Sub TransformAndCleanDataForPostgres()
    Dim ws As Worksheet, wsTarget As Worksheet
    Dim lastRow As Long, lastCol As Long, totalCol As Long
    Dim dateStartCol As Integer, lastDateCol As Long
    Dim r As Long, c As Long, outRow As Long
    Dim empID As Variant, dailyVal As Variant, cellColor As Long, empTotal As Variant
    Dim dictColors As Object
    Dim targetSheetName As String, dateHeader As String, valHeader As String
    Dim tCol As Integer, headerVal As String
    Dim obd As Variant, rsd As Variant
    Dim isLeave As Boolean, isOT As Boolean, isLate As Boolean
    Dim sheetsToDelete As Collection
    Dim sheetName As Variant
    
    ' Dynamic Column Trackers for OT/Late
    Dim colID As Integer, colName As Integer, colOnboard As Integer, colResign As Integer
    Dim colFactory As Integer, colGroup As Integer, colDept As Integer
    
    Set sheetsToDelete = New Collection
    
    ' Application optimizations
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual

    ' Loop through all sheets in the workbook
    For Each ws In ThisWorkbook.Worksheets
        If UCase(ws.Name) <> "MAIN" And Right(ws.Name, 7) <> "_result" Then
            
            isLeave = (InStr(1, ws.Name, "Leave", vbTextCompare) > 0)
            isOT = (InStr(1, ws.Name, "OT", vbTextCompare) > 0)
            isLate = (InStr(1, ws.Name, "Late", vbTextCompare) > 0)
            
            If Not (isLeave Or isOT Or isLate) Then GoTo SkipSheet
            
            targetSheetName = ws.Name & "_result"
            On Error Resume Next
            ThisWorkbook.Worksheets(targetSheetName).Delete
            On Error GoTo 0
            
            Set wsTarget = ThisWorkbook.Worksheets.Add(After:=ws)
            wsTarget.Name = targetSheetName
            
            ' 1. Write STATIC Postgres Headers
            If isLeave Then
                dateHeader = "Leave Date"
                valHeader = "Leave Hour"
            ElseIf isOT Then
                dateHeader = "OT Date"
                valHeader = "OT Hour"
            ElseIf isLate Then
                dateHeader = "Late Date"
                valHeader = "Late Value"
            End If
            
            Dim headers As Variant
            headers = Array("ID No", "English Name", "Onboard Date", "Resign Date", "Factory", "Group Code", "Department", dateHeader, valHeader, "Total")
            wsTarget.Range("A1:J1").Value = headers
            
            ' Pre-format columns (Data Integrity, keeping leading zeros and SQL dates)
            wsTarget.Columns("F:F").NumberFormat = "@"
            wsTarget.Columns("C:D").NumberFormat = "yyyy-mm-dd"
            wsTarget.Columns("H:H").NumberFormat = "yyyy-mm-dd"
            
            lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
            lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
            outRow = 2
            
            ' =========================================================
            ' BRANCH 1: STRICT LOGIC FOR LEAVE SHEETS (Restored to Working Version)
            ' =========================================================
            If isLeave Then
                dateStartCol = 22 ' Fixed at Col V
                
                ' Map the 14 Leave colors dynamically to target columns K to X
                Set dictColors = CreateObject("Scripting.Dictionary")
                Dim targetLeaveCol As Integer
                targetLeaveCol = 11 ' Column K
                
                For c = 8 To 21
                    cellColor = ws.Cells(1, c).Interior.Color
                    If Not dictColors.Exists(cellColor) Then dictColors.Add cellColor, targetLeaveCol
                    wsTarget.Cells(1, targetLeaveCol).Value = ws.Cells(1, c).Value
                    targetLeaveCol = targetLeaveCol + 1
                Next c
                
                ' Find "Total" Column
                totalCol = 0
                For c = dateStartCol To lastCol
                    If InStr(1, ws.Cells(1, c).Value, "Total", vbTextCompare) > 0 Then totalCol = c: Exit For
                Next c
                If totalCol > 0 Then lastDateCol = totalCol - 1 Else lastDateCol = lastCol
                
                ' Loop rows and write strict data
                For r = 2 To lastRow
                    empID = ws.Cells(r, 1).Value
                    If Trim(empID) <> "" Then
                        empTotal = IIf(totalCol > 0, ws.Cells(r, totalCol).Value, 0)
                        
                        For c = dateStartCol To lastDateCol
                            dailyVal = ws.Cells(r, c).Value
                            
                            If IsNumeric(dailyVal) And dailyVal > 0 Then
                                ' Strict block copy for A-G
                                wsTarget.Range("A" & outRow & ":G" & outRow).Value = ws.Range("A" & r & ":G" & r).Value
                                wsTarget.Cells(outRow, 6).Value = ws.Cells(r, 6).text ' Fix leading 00
                                
                                ' Date Formatting
                                obd = ws.Cells(r, 3).Value
                                If Len(CStr(obd)) = 8 And IsNumeric(obd) Then wsTarget.Cells(outRow, 3).Value = DateSerial(Left(obd, 4), Mid(obd, 5, 2), Right(obd, 2))
                                rsd = ws.Cells(r, 4).Value
                                If Len(CStr(rsd)) = 8 And IsNumeric(rsd) Then wsTarget.Cells(outRow, 4).Value = DateSerial(Left(rsd, 4), Mid(rsd, 5, 2), Right(rsd, 2))
                                
                                ' Write Date, Value, Total
                                wsTarget.Cells(outRow, 8).Value = ws.Cells(1, c).Value
                                wsTarget.Cells(outRow, 9).Value = dailyVal
                                wsTarget.Cells(outRow, 10).Value = empTotal
                                
                                ' Output mapped leave values based on source color
                                cellColor = ws.Cells(r, c).Interior.Color
                                If dictColors.Exists(cellColor) Then
                                    tCol = dictColors(cellColor)
                                    wsTarget.Cells(outRow, tCol).Value = dailyVal
                                End If
                                
                                outRow = outRow + 1
                            End If
                        Next c
                    End If
                Next r
                Set dictColors = Nothing

            ' =========================================================
            ' BRANCH 2: DYNAMIC LOGIC FOR OT & LATE SHEETS
            ' =========================================================
            Else ' isOT or isLate
                ' Reset trackers
                colID = 0: colName = 0: colOnboard = 0: colResign = 0
                colFactory = 0: colGroup = 0: colDept = 0
                
                ' Scan headers dynamically
                For c = 1 To lastCol
                    headerVal = UCase(Trim(ws.Cells(1, c).Value))
                    If headerVal Like "*ID*" Then colID = c
                    If headerVal Like "*ENGLISH*" Or headerVal Like "*NAME*" Then colName = c
                    If headerVal Like "*ONB*" And headerVal Like "*DATE*" Then colOnboard = c
                    If headerVal Like "*RESIGN*" Then colResign = c
                    If headerVal Like "*FACTORY*" Then colFactory = c
                    If headerVal Like "*GROUP*" Then colGroup = c
                    If headerVal Like "*DEPT*" Or headerVal Like "*DEPARTMENT*" Then colDept = c
                Next c
                
                ' Scan for dates dynamically
                dateStartCol = 0
                For c = 1 To lastCol
                    headerVal = Trim(ws.Cells(1, c).text)
                    If IsDate(headerVal) Or headerVal Like "*-*-*" Or headerVal Like "*/*/*" Then dateStartCol = c: Exit For
                Next c
                If dateStartCol = 0 Then dateStartCol = 8 ' Fallback
                
                ' Find "Total" Column
                totalCol = 0
                For c = dateStartCol To lastCol
                    If InStr(1, ws.Cells(1, c).Value, "Total", vbTextCompare) > 0 Then totalCol = c: Exit For
                Next c
                If totalCol > 0 Then lastDateCol = totalCol - 1 Else lastDateCol = lastCol
                
                ' Loop rows and write dynamic data
                For r = 2 To lastRow
                    empID = ws.Cells(r, IIf(colID > 0, colID, 1)).Value
                    If Trim(empID) <> "" Then
                        empTotal = IIf(totalCol > 0, ws.Cells(r, totalCol).Value, 0)
                        
                        For c = dateStartCol To lastDateCol
                            dailyVal = ws.Cells(r, c).Value
                            
                            If IsNumeric(dailyVal) And dailyVal > 0 Then
                                ' Map values using dynamic column trackers
                                If colID > 0 Then wsTarget.Cells(outRow, 1).Value = ws.Cells(r, colID).Value
                                If colName > 0 Then wsTarget.Cells(outRow, 2).Value = ws.Cells(r, colName).Value
                                If colFactory > 0 Then wsTarget.Cells(outRow, 5).Value = ws.Cells(r, colFactory).Value
                                If colGroup > 0 Then wsTarget.Cells(outRow, 6).Value = ws.Cells(r, colGroup).text
                                If colDept > 0 Then wsTarget.Cells(outRow, 7).Value = ws.Cells(r, colDept).Value
                                
                                ' Date Formatting
                                If colOnboard > 0 Then
                                    obd = ws.Cells(r, colOnboard).Value
                                    If Len(CStr(obd)) = 8 And IsNumeric(obd) Then wsTarget.Cells(outRow, 3).Value = DateSerial(Left(obd, 4), Mid(obd, 5, 2), Right(obd, 2))
                                End If
                                If colResign > 0 Then
                                    rsd = ws.Cells(r, colResign).Value
                                    If Len(CStr(rsd)) = 8 And IsNumeric(rsd) Then wsTarget.Cells(outRow, 4).Value = DateSerial(Left(rsd, 4), Mid(rsd, 5, 2), Right(rsd, 2))
                                End If
                                
                                ' Write Date, Value, Total
                                wsTarget.Cells(outRow, 8).Value = ws.Cells(1, c).Value
                                wsTarget.Cells(outRow, 9).Value = dailyVal
                                wsTarget.Cells(outRow, 10).Value = empTotal
                                
                                outRow = outRow + 1
                            End If
                        Next c
                    End If
                Next r
            End If
            
            ' AutoFit columns and flag source sheet for deletion
            wsTarget.UsedRange.Columns.AutoFit
            sheetsToDelete.Add ws.Name
            
        End If
SkipSheet:
    Next ws
    
    ' Delete original source sheets
    For Each sheetName In sheetsToDelete
        ThisWorkbook.Worksheets(sheetName).Delete
    Next sheetName
    
    ' Restore Application Settings
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
    MsgBox "Processing complete. Leave sheets used strict mapping, OT/Late used dynamic mapping.", vbInformation, "Process Finished"
    
    On Error Resume Next
    ThisWorkbook.Sheets("MAIN").Activate
    On Error GoTo 0
End Sub
