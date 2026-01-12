Attribute VB_Name = "RunnerModule"
Public Function RunAutomation(moduleName As String, macroName As String, automationFile As String, automationPath As String, _
    automationSheet, sapSession As Long, sapSessionCell As String, statementDate As String, statementDateCell As String)
    
    Dim wbDashboard As Workbook
    Dim shtDashboard As Worksheet
    
    Set wbDashboard = ThisWorkbook
    Set shtDashboard = wbDashboard.Sheets("Dashboard")
    
    Dim proceedRun As VbMsgBoxResult
	
    proceedRun = MsgBox("Running the following Automation:" & vbNewLine & vbNewLine & automationFile & "'!" & moduleName & "." & macroName _
        & vbNewLine & vbNewLine & "Are you sure you want to proceed??", vbYesNo)
    
    If proceedRun = vbYes Then
        
        Dim automationApp As New Excel.Application
        
        automationApp.Workbooks.Open fileName:=automationPath & "\" & automationFile
        automationApp.Workbooks(automationFile).Windows(1).Visible = True
        automationApp.Visible = True
        
        Dim wbAutomation As Workbook
        Dim shtAutomation As Worksheet
        
        Set wbAutomation = automationApp.Workbooks(automationFile)
        Set shtAutomation = wbAutomation.Sheets(automationSheet)
        
        Dim runSession As Range
        Dim runDate As Range
        
        Set runSession = shtAutomation.Range(sapSessionCell)
        Set runDate = shtAutomation.Range(statementDateCell)
        
        runSession.Value = sapSession - 1
        runSession.Font.Color = vbWhite
        runDate.Value = statementDate
        
        If Not postingDate = "" Then
            If postingDateCell = "" Then
                MsgBox "Posting Date set but no Cell to update. Please include Cell."
                Exit Function
            End If
            Dim runPostDate As Range
            Set runPostDate = shtAutomation.Range(postingDateCell)
            runPostDate.Value = postingDate
        End If
        
        'automationApp.Workbooks(automationFile).Activate
        'automationApp.Workbooks(automationFile).Close SaveChanges:=True
        
    ElseIf proceedRun = vbNo Then
        Exit Function
    End If

End Function

Public Function OpenAutomationFile(automationFile As String, automationPath As String)
    
    Dim proceedRun As VbMsgBoxResult
    
    proceedRun = MsgBox("Opening the following Automation in a new Excel Instance:" & vbNewLine & vbNewLine & automationFile & _
        vbNewLine & vbNewLine & "Are you sure you want to proceed??", vbYesNo)
    
    If proceedRun = vbYes Then
        Dim automationApp As New Excel.Application
        automationApp.Workbooks.Open fileName:=automationPath & "\" & automationFile
        automationApp.Workbooks(automationFile).Windows(1).Visible = True
        automationApp.Visible = True
    ElseIf proceedRun = vbNo Then
        Exit Function
    End If

End Function