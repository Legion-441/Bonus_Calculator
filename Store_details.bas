Attribute VB_Name = "Store_details"
Sub ChangeStoreDetails_button()
'When user click the change store details button
    
    'Sheet_Config "store details" validation
    If Range("Config_Store_Name_Number") = "" Or _
            Range("Config_Cafe_format") = "" Or _
            Range("Config_Device_1") = "" Or _
            Range("Config_Device_2") = "" Or _
            Range("Config_Surname") = "" Or _
            Range("Config_Deputy") = "" _
        Then
            'Disabling the "Cancel" button
            ufStoreDetails.Button_Cancel.Enabled = False
    End If
    
    'Transfers data from config sheet to userform
    ufStoreDetails.TextBox_StoreName.Value = Range("Config_Store_Name_Number") 'Store name
        
    If Range("Config_Cafe_format") <> "" Then
    ufStoreDetails.ComboBox_Format.Value = Range("Config_Cafe_format")
    End If
    'Cafe format
        
    If Range("Config_Device_1") <> "" Then
    ufStoreDetails.CheckBox_Device1.Value = Range("Config_Device_1")
    End If
    'Presence of a device 1
        
    If Range("Config_Device_2") <> "" Then
    ufStoreDetails.CheckBox_Device2.Value = Range("Config_Device_2")
    End If
    'Presence of a device 2
        
    If Range("Config_Deputy") <> "" Then
    ufStoreDetails.CheckBox_Deputy.Value = Range("Config_Deputy")
    End If
    'Presence of a deputy
        
    If Range("Config_Surname") = False Then
    ufStoreDetails.OptionButton_Payroll.Value = True
    End If
    'Display mode
        
    ufStoreDetails.TextBox_StartDate = Format(Range("Config_Start"), "dd/mm/yy") 'Start date
    ufStoreDetails.TextBox_EndDate = Format(Range("Config_End"), "dd/mm/yy") 'End date
        
    'Userform pops up
    ufStoreDetails.Show
    
End Sub
