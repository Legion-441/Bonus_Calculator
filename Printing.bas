Attribute VB_Name = "Printing"
Const LRMargin As Single = 0.25
Const TBMargin As Single = 0.25
Const WorkspaceRatio As Single = (11.6929134 - 2 * LRMargin) / (8.26771654 - 2 * TBMargin)
    
    
Sub PageSetup()

    'Setup printing page
    With Sheet_Dashboard.PageSetup
        .Zoom = False
        .Orientation = xlLandscape
        .PaperSize = xlPaperA4
        .HeaderMargin = Application.InchesToPoints(0)
        .FooterMargin = Application.InchesToPoints(0)
        .LeftMargin = Application.InchesToPoints(LRMargin)
        .RightMargin = Application.InchesToPoints(LRMargin)
        .TopMargin = Application.InchesToPoints(TBMargin)
        .BottomMargin = Application.InchesToPoints(TBMargin)
        .FitToPagesTall = 1
        .FitToPagesWide = 1
    End With
        
End Sub


Sub PrintAll()
'When user click the "Print all" button

    Dim PA_1p As Range
    Dim PA_2p As Range
    Dim TableRatio As Single

    Set PA_1p = Range("Print_All_1page")
    Set PA_2p = Range("Print_All_2pages")
        
    TableRatio = PA_1p.Width / PA_1p.Height

    Application.ScreenUpdating = False
    
    Call PageSetup

    'Debug print
    Debug.Print "Workspace Ratio:"; WorkspaceRatio
    Debug.Print "Table Ratio:"; TableRatio
    
    'Set single or two-sided printing based on the dimension ratio
    If TableRatio < WorkspaceRatio Then
            PA_2p.PrintOut copies:=1
            Debug.Print "2 pages"
        Else
            PA_1p.PrintOut copies:=1
            Debug.Print "1 page"
    End If

    Application.ScreenUpdating = True
    
End Sub

Sub PrintEmployee()
'when user click the employee "Print" button


    Application.ScreenUpdating = False
    
    Call PageSetup
    
    Range("Print_Employees").PrintOut copies:=1
    
    Application.ScreenUpdating = True
    
End Sub

Sub PrintDate()
'when user click the date "Print" button

    Application.ScreenUpdating = False

    Call PageSetup

    Range("Print_Date").PrintOut copies:=1
    
    Application.ScreenUpdating = True
    
End Sub
