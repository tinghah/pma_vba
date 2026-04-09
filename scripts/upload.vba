Option Explicit

' PostgreSQL Connection Configuration
Private Const DB_HOST As String = "172.23.86.119"
Private Const DB_PORT As String = "5432"
Private Const DB_NAME As String = "pma_hr"
Private Const DB_USER As String = "postgres"
Private Const DB_PASS As String = "Abc123"

Public Sub UploadToPostgres()
    Dim conn As Object
    Dim ws As Worksheet
    Dim lastRow As Long, lastCol As Long
    Dim r As Long, c As Long
    Dim sql As String, dbColumns As String, valString As String, updateString As String
    Dim successCount As Long
    Dim connStr As String
    Dim dateColName As String ' To hold leave_date, ot_date, or late_date for the PK
    
    Set conn = CreateObject("ADODB.Connection")
    
    ' Connection String
    connStr = "Driver={PostgreSQL Unicode};Server=" & DB_HOST & ";Port=" & DB_PORT & ";Database=" & DB_NAME & ";Uid=" & DB_USER & ";Pwd=" & DB_PASS & ";"
    
    On Error Resume Next
    conn.Open connStr
    If conn.State = 0 Then
        MsgBox "Connection Failed! Check if PostgreSQL Unicode Driver is installed.", vbCritical
        Exit Sub
    End If
    On Error GoTo 0
    
    Application.ScreenUpdating = False
    
    For Each ws In ThisWorkbook.Worksheets
        If Right(ws.Name, 7) = "_result" Then
            Dim tableType As String
            If InStr(1, ws.Name, "Leave", vbTextCompare) > 0 Then
                tableType = "leave": dateColName = "leave_date"
            ElseIf InStr(1, ws.Name, "OT", vbTextCompare) > 0 Then
                tableType = "ot": dateColName = "ot_date"
            ElseIf InStr(1, ws.Name, "Late", vbTextCompare) > 0 Then
                tableType = "late": dateColName = "late_date"
            Else
                GoTo NextSheet
            End If
            
            lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
            lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
            
            ' Process Rows
            For r = 2 To lastRow
                dbColumns = "": valString = "": updateString = ""
                
                For c = 1 To lastCol
                    Dim rawHeader As String, colName As String, cellVal As Variant, formattedVal As String
                    rawHeader = LCase(Replace(ws.Cells(1, c).Value, " ", "_"))
                    colName = GetDBColumnName(rawHeader, tableType)
                    
                    If colName <> "" Then
                        cellVal = ws.Cells(r, c).Value
                        
                        ' Format Value
                        If IsEmpty(cellVal) Or Trim(CStr(cellVal)) = "" Then
                            formattedVal = "NULL"
                        ElseIf IsDate(cellVal) Then
                            formattedVal = "'" & Format(cellVal, "yyyy-mm-dd") & "'"
                        ElseIf IsNumeric(cellVal) Then
                            formattedVal = Replace(CStr(cellVal), ",", ".") ' Ensure dot decimal
                        Else
                            formattedVal = "'" & Replace(CStr(cellVal), "'", "''") & "'"
                        End If
                        
                        dbColumns = dbColumns & colName & ","
                        valString = valString & formattedVal & ","
                        ' Build update string for "ON CONFLICT"
                        If colName <> "id_no" And colName <> dateColName Then
                            updateString = updateString & colName & "=EXCLUDED." & colName & ","
                        End If
                    End If
                Next c
                
                If Len(dbColumns) > 0 Then
                    dbColumns = Left(dbColumns, Len(dbColumns) - 1)
                    valString = Left(valString, Len(valString) - 1)
                    updateString = Left(updateString, Len(updateString) - 1)
                    
                    ' UPSERT SQL: Allows daily overwriting of existing data
                    sql = "INSERT INTO " & tableType & " (" & dbColumns & ") VALUES (" & valString & ") " & _
                          "ON CONFLICT (id_no, " & dateColName & ") DO UPDATE SET " & updateString & ";"
                    
                    On Error Resume Next
                    conn.Execute sql
                    If Err.Number = 0 Then
                        successCount = successCount + 1
                    Else
                        Debug.Print "Error in Sheet " & ws.Name & " Row " & r & ": " & Err.Description
                        Debug.Print "SQL: " & sql
                    End If
                    On Error GoTo 0
                End If
            Next r
        End If
NextSheet:
    Next ws
    
    conn.Close
    Application.ScreenUpdating = True
    MsgBox "Done! Successfully processed " & successCount & " records.", vbInformation
End Sub

Private Function GetDBColumnName(header As String, tableType As String) As String
    ' Standard mappings
    Select Case header
        Case "id_no", "factory", "department", "total", "english_name", "onboard_date", "resign_date": GetDBColumnName = header
        Case "group_code": GetDBColumnName = "group_code"
        
        ' Specific Type mappings
        Case "leave_date", "ot_date", "late_date": GetDBColumnName = header
        Case "leave_hour", "ot_hour", "late_value": GetDBColumnName = header
        
        ' Leave Column sub-types
        Case "casual", "unpaid", "ssb_sick", "sick", "earned", "og", "maternity", "paternity", "official", "paid_injury", "unpaid_injury", "other", "business", "absent"
            GetDBColumnName = header
        Case Else: GetDBColumnName = ""
    End Select
End Function

