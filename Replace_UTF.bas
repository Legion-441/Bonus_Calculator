Attribute VB_Name = "Replace_UTF"
Sub Replace_UTF_with_W1250()
Attribute Replace_UTF_with_W1250.VB_ProcData.VB_Invoke_Func = " \n14"

    Sheet_Import.Visible = True
    Sheet_Import.Select

        Cells.replace What:="ƒÖ", Replacement:="π", LookAt:=xlPart, SearchOrder _
            :=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
            
        Cells.replace What:="ƒá", Replacement:="Ê", LookAt:=xlPart, SearchOrder _
            :=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
            
        Cells.replace What:="ƒô", Replacement:="Í", LookAt:=xlPart, SearchOrder _
            :=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
            
        Cells.replace What:="≈Ç", Replacement:="≥", LookAt:=xlPart, SearchOrder _
            :=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
            
        Cells.replace What:="≈Ñ", Replacement:="Ò", LookAt:=xlPart, SearchOrder _
            :=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
            
        Cells.replace What:="√≥", Replacement:="Û", LookAt:=xlPart, SearchOrder _
            :=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
            
        Cells.replace What:="≈õ", Replacement:="ú", LookAt:=xlPart, SearchOrder _
            :=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
            
        Cells.replace What:="≈∫", Replacement:="ü", LookAt:=xlPart, SearchOrder _
            :=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
            
        Cells.replace What:="≈º", Replacement:="ø", LookAt:=xlPart, SearchOrder _
            :=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
    
    

        Cells.replace What:="ƒÑ", Replacement:="•", LookAt:=xlPart, SearchOrder _
            :=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
            
        Cells.replace What:="ƒÜ", Replacement:="∆", LookAt:=xlPart, SearchOrder _
            :=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
            
        Cells.replace What:="√ì", Replacement:="”", LookAt:=xlPart, SearchOrder _
            :=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
            
        Cells.replace What:="≈ö", Replacement:="å", LookAt:=xlPart, SearchOrder _
            :=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
            
        Cells.replace What:="≈π", Replacement:="Ø", LookAt:=xlPart, SearchOrder _
            :=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
        
        Cells.replace What:="≈ª", Replacement:="è", LookAt:=xlPart, SearchOrder _
            :=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
     
        Cells.replace What:="ƒ", Replacement:=" ", LookAt:=xlPart, SearchOrder _
            :=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
            
        Cells.replace What:="≈", Replacement:="£", LookAt:=xlPart, SearchOrder _
            :=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
            
    Sheet_Import.Visible = False
    Sheet_Dashboard.Select
            
End Sub

