Attribute VB_Name = "Replace_UTF"
Sub Replace_UTF_with_W1250()
Attribute Replace_UTF_with_W1250.VB_ProcData.VB_Invoke_Func = " \n14"

    Sheet_Import.Visible = True
    Sheet_Import.Select

        Cells.replace What:="ą", Replacement:="�", LookAt:=xlPart, SearchOrder _
            :=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
            
        Cells.replace What:="ć", Replacement:="�", LookAt:=xlPart, SearchOrder _
            :=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
            
        Cells.replace What:="ę", Replacement:="�", LookAt:=xlPart, SearchOrder _
            :=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
            
        Cells.replace What:="ł", Replacement:="�", LookAt:=xlPart, SearchOrder _
            :=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
            
        Cells.replace What:="ń", Replacement:="�", LookAt:=xlPart, SearchOrder _
            :=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
            
        Cells.replace What:="ó", Replacement:="�", LookAt:=xlPart, SearchOrder _
            :=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
            
        Cells.replace What:="ś", Replacement:="�", LookAt:=xlPart, SearchOrder _
            :=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
            
        Cells.replace What:="ź", Replacement:="�", LookAt:=xlPart, SearchOrder _
            :=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
            
        Cells.replace What:="ż", Replacement:="�", LookAt:=xlPart, SearchOrder _
            :=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
    
    

        Cells.replace What:="Ą", Replacement:="�", LookAt:=xlPart, SearchOrder _
            :=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
            
        Cells.replace What:="Ć", Replacement:="�", LookAt:=xlPart, SearchOrder _
            :=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
            
        Cells.replace What:="Ó", Replacement:="�", LookAt:=xlPart, SearchOrder _
            :=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
            
        Cells.replace What:="Ś", Replacement:="�", LookAt:=xlPart, SearchOrder _
            :=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
            
        Cells.replace What:="Ź", Replacement:="�", LookAt:=xlPart, SearchOrder _
            :=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
        
        Cells.replace What:="Ż", Replacement:="�", LookAt:=xlPart, SearchOrder _
            :=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
     
        Cells.replace What:="�", Replacement:="�", LookAt:=xlPart, SearchOrder _
            :=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
            
        Cells.replace What:="�", Replacement:="�", LookAt:=xlPart, SearchOrder _
            :=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
            
    Sheet_Import.Visible = False
    Sheet_Dashboard.Select
            
End Sub

