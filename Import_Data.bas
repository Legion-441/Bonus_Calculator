Attribute VB_Name = "Import_Data"
Sub Import_button()
'When user click the import button

Dim fileToOpen As Variant
Dim fileFilterPattern As String
Dim wsMaster As Worksheet
Dim wbTextImport As Workbook

    
    'I DON'T KNOW WHY IT DOESN'T WORK:
    '   Set wsMaster = ThisWorkbook.Sheet_Import
    Set wsMaster = ThisWorkbook.Worksheets("IMPORT")
            
    'Sheet_Config "store details" validation
    If Range("Config_Store_Name_Number") = "" Or _
        Range("Config_Cafe_format") = "" Or _
        Range("Config_Device_1") = "" Or _
        Range("Config_Device_2") = "" Or _
        Range("Config_Surname") = "" Or _
        Range("Config_Deputy") = "" _
    Then

        '"Enter store details" msg
         Msgbox Sheet_Formulas.Range("Formulas_Enter_store_details")
    
    Else 'Starts importing data
        
        
        Application.ScreenUpdating = False
        
        'Sets filter pattern
        fileFilterPattern = "Text Files (*.txt; *.csv), *.txt; *.csv"
        fileToOpen = Application.GetOpenFilename(fileFilterPattern)
        
        If fileToOpen = False Then
            '"No file" msg
            Msgbox Sheet_Formulas.Range("Formula_No_file")
            Application.ScreenUpdating = True
        Else
            'Reset_table subroutine
            Call Reset_table
            
            'Creates an additional workbook
            Workbooks.OpenText _
                    FileName:=fileToOpen, _
                    DataType:=xlDelimited, _
                    Local:=True
            Set wbTextImport = ActiveWorkbook
            
            'Processes imported data
            Rows(1).EntireRow.Delete
            Columns(1).TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, _
                TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
                Semicolon:=False, Comma:=True, Space:=False, Other:=False, OtherChar:="F", _
                FieldInfo:=Array(Array(1, 9), Array(2, 4), Array(3, 2), Array(4, 2), Array(5, 1), _
                Array(6, 9), Array(7, 9), Array(8, 9), Array(9, 9), Array(10, 9), Array(11, 9), _
                Array(12, 2), Array(13, 9), Array(14, 9), Array(15, 9)), TrailingMinusNumbers:=True
            
            'Transfers imorted data to Sheet_import
            wbTextImport.Worksheets(1).Range("A1").CurrentRegion.Copy wsMaster.Range("A2")
            
            'Closes additional workbook
            wbTextImport.Close False
            
            'Replaces_UTF_with_W1250 & Update_Dashboard subroutine
            Call Replace_UTF_with_W1250
            Call Update_Dashboard
            
            '"Data loaded" msg
            Msgbox Sheet_Formulas.Range("Formulas_Data_loaded")
            
        End If
    
    End If

End Sub


