Private Sub CommandButton1_Click()
    Dim strProgramName As String
    Dim strArgument As String
    Dim strHost As String
    Dim strUser As String
    Dim strPass As String
    
    On Error Resume Next
    
Dim myRange As Range
Dim cell As Range
Set myRange = Selection
'    MsgBox myRange.Rows.count

'strProgramName = "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"

strProgramName = Cells(1, 1).Value
    
    For Each cell In myRange
        'MsgBox "AT: " & cell.Value
        
        strHost = Cells(cell.Row, 1).Value
        strUser = Cells(cell.Row, 2).Value
        strPass = Cells(cell.Row, 3).Value
        
        strArgument = strUser & "@" & strHost & " -pw " & strPass

        Call Shell("""" & strProgramName & """ """ & strArgument & """", vbNormalFocus)
    
    Next cell
    
    
    'MsgBox count & " item(s) selected"

    'strProgramName = "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
    'strArgument = "https://www.coursera.org/programs/gcp-pathways-licences-bxhqk"



    Call Shell("""" & strProgramName & """ """ & strArgument & """", vbNormalFocus)
End Sub
