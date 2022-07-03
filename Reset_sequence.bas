Attribute VB_Name = "Reset_sequence"
Sub Refresh_button()
    
    Call Update_Dashboard
    '"Updated Dashboard" msg
    Msgbox Sheet_Formulas.Range("Formulas_Dashboard_updated")
            
End Sub
Sub Update_Dashboard()

    Application.ScreenUpdating = False
        
    Sheet_Dashboard.Activate
    Sheet_Dashboard.Unprotect
            
    Call Unhide_rows
    Call Unhide_columns
    Call Refresh_tables
    Call Hide_rows
    Call Hide_columns
        
    Sheet_Dashboard.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True _
            , AllowSorting:=True, AllowFiltering:=True, AllowUsingPivotTables:=True
        
    Application.ScreenUpdating = True

End Sub

Sub Refresh_tables()

    Sheet_Inputs_Data.ListObjects("FinalTable").QueryTable.Refresh BackgroundQuery:=False
    Sheet_PivotTable_Employee.PivotTables("PivotPracownik").PivotCache.Refresh
    Sheet_PivotTable_Date.PivotTables("PivotDate").PivotCache.Refresh
    
End Sub

Sub Hide_rows()

    Dim X As Long

    'Hiding a row if it is empty, in employee table
    For X = 7 To 46
        If Cells(X, "C").Value = "" Then
        Rows(X).Hidden = True
        End If
    Next X
    
    'Hiding a row if it is empty, in date table
    For X = 53 To 115
        If Cells(X, "C").Value = "" Then
        Rows(X).Hidden = True
        End If
    Next X
 
End Sub

Sub Hide_columns()

    Dim Y As Long

    'Hiding a column if there is no rate
    For Y = 5 To 21
        If Cells(5, Y).Value = 0 Or _
        Cells(5, Y).Value = "" Then
        Columns(Y).Hidden = True
        End If
    Next Y
 
End Sub

Sub Unhide_rows()

    Dim X As Long
    
    'Unhiding the rows in employee and date table
    For X = 7 To 115
        Rows(X).Hidden = False
    Next X

End Sub

Sub Unhide_columns()

    Dim Y As Long

    'Unhiding the columns
    For Y = 5 To 21
        Columns(Y).Hidden = False
    Next Y
    
End Sub

Sub Reset_table()
    
    'Deletion of old data
    With Sheet_Import.ListObjects("TableData")
        If Not .DataBodyRange Is Nothing Then
            .DataBodyRange.Delete
        End If
    End With

End Sub
