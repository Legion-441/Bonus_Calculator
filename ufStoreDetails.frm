VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufStoreDetails 
   Caption         =   "Store Details"
   ClientHeight    =   6405
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6105
   OleObjectBlob   =   "ufStoreDetails.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ufStoreDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()
'When the userform opens

    'Filling the list with Cafe formats
    ComboBox_Format.List = Array("S", "L", "XL")
    'Language
    Label_StoreName.Caption = Sheet_Formulas.Range("Formulas_Store_name_number")
    Label_CaffeeFormat.Caption = Sheet_Formulas.Range("Formulas_Caffee_Format")
    Label_Device.Caption = Sheet_Formulas.Range("Formulas_Device")
    CheckBox_Device1.Caption = Sheet_Formulas.Range("Formulas_Device1")
    CheckBox_Device2.Caption = Sheet_Formulas.Range("Formulas_Device2")
    Label_DisplayMode.Caption = Sheet_Formulas.Range("Formulas_Display_mode")
    OptionButton_Surname.Caption = Sheet_Formulas.Range("Formulas_Surnames")
    OptionButton_Payroll.Caption = Sheet_Formulas.Range("Formulas_Payroll")
    Label_RunningStore.Caption = Sheet_Formulas.Range("Formulas_Running_store")
    CheckBox_Manager.Caption = Sheet_Formulas.Range("Formulas_Manager")
    CheckBox_Deputy.Caption = Sheet_Formulas.Range("Formulas_Deputy")
    Label_DateRange.Caption = Sheet_Formulas.Range("Formulas_Date_range")
    Button_Confirm.Caption = Sheet_Formulas.Range("Formulas_Confirm")
    Button_Cancel.Caption = Sheet_Formulas.Range("Formulas_Cancel")
    
End Sub

Private Sub Button_Cancel_Click()
'When user click the cancel button

    'Exits the userform
    Unload ufStoreDetails
    
End Sub

Private Sub Button_Confirm_Click()
'When user click the confirm button

    'Userform inputs validation
    If ufStoreDetails.TextBox_StoreName = "" Or _
        ufStoreDetails.ComboBox_Format = "" Or _
        ufStoreDetails.OptionButton_Surname.Value = ufStoreDetails.OptionButton_Payroll.Value Or _
        ufStoreDetails.TextBox_StartDate.Value = "" Or _
        ufStoreDetails.TextBox_EndDate.Value = "" _
    Then
        '"Fill Form" msg
        Msgbox Sheet_Formulas.Range("Formulas_Fill_Form")
    Else
        'Transfer data to config sheet
        Application.ScreenUpdating = False
        Range("Config_Store_Name_Number") = ufStoreDetails.TextBox_StoreName.Value
        Range("Config_Cafe_format") = ufStoreDetails.ComboBox_Format.Value
        Range("Config_Device_1") = ufStoreDetails.CheckBox_Device1.Value
        Range("Config_Device_2") = ufStoreDetails.CheckBox_Device2.Value
        Range("Config_Surname") = ufStoreDetails.OptionButton_Surname.Value
        Range("Config_Deputy") = ufStoreDetails.CheckBox_Deputy.Value
        Range("Config_Start") = Format(ufStoreDetails.TextBox_StartDate.Value, "dd/mm/yy")
        Range("Config_End") = Format(ufStoreDetails.TextBox_EndDate.Value, "dd/mm/yy")
        
        'Update dashboard
        Call Update_Dashboard
            
        'Exits the userform
        Unload ufStoreDetails
        
        '"Updated store details" msg
        Msgbox Sheet_Formulas.Range("Formulas_Updated_store_details")
        
    End If
    
End Sub
