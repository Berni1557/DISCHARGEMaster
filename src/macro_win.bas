'Attribute VB_Name = "Modul1"

Sub openScan()
    'Attribute openScan.VB_ProcData.VB_Invoke_Func = "o\n14"
    Column = ActiveCell.Column
    Row = ActiveCell.Row
    Dim SeriesInstanceUID As String
    StudyInstanceUID = ActiveSheet.Cells(Row, 4)
    SeriesInstanceUID = ActiveSheet.Cells(Row, 5)
    
    'MsgBox ("Open StudyInstanceUID:" + StudyInstanceUID + " SeriesInstanceUID: " + SeriesInstanceUID)
    answer = MsgBox("Open StudyInstanceUID:" + StudyInstanceUID + " SeriesInstanceUID: " + SeriesInstanceUID, vbQuestion + vbYesNo + vbDefaultButton2, "Message Box Title")
    If answer = vbYes Then
        Dim ReturnValue
        Dim command As String
        command = "C:\python37\python " + "H:\cloud\cloud_data\Projects\CACSFilter\src\openScan.py " + StudyInstanceUID + " " + SeriesInstanceUID
        ReturnValue = Shell(command)
    End If
End Sub

Sub sendEmail()
'Attribute sendEmail.VB_ProcData.VB_Invoke_Func = "m\n14"
    filepath = "H:\cloud\cloud_data\Projects\CACSFilter\src\text_tmp.txt"
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set file = fs.CreateTextFile(filepath, True)
    For Each cell In Selection
        Row = cell.Row
        Site = ActiveSheet.Cells(Row, 3)
        PatientID = ActiveSheet.Cells(Row, 4)
        StudyInstanceUID = ActiveSheet.Cells(Row, 5)
        SeriesInstanceUID = ActiveSheet.Cells(Row, 6)
        ProblemSummary = ActiveSheet.Cells(Row, 7)
        Problem = ActiveSheet.Cells(Row, 8)
        DateQuery = ActiveSheet.Cells(Row, 9)
        DateAnswer = ActiveSheet.Cells(Row, 10)
        Results = ActiveSheet.Cells(Row, 11)
        Status = ActiveSheet.Cells(Row, 12)
        ResponsiblePersonProblem = ActiveSheet.Cells(Row, 13)
        file.WriteLine ("---")
        file.WriteLine (CStr(Row))
        file.WriteLine (Site)
        file.WriteLine (PatientID)
        file.WriteLine (StudyInstanceUID)
        file.WriteLine (SeriesInstanceUID)
        file.WriteLine (ProblemSummary)
        file.WriteLine (Problem)
        file.WriteLine (DateQuery)
        file.WriteLine (DateAnswer)
        file.WriteLine (Results)
        file.WriteLine (Status)
        file.WriteLine (ResponsiblePersonProblem)
    Next cell
    file.Close
    
    Dim command As String
    command = "C:\python37\python " + "H:\cloud\cloud_data\Projects\CACSFilter\src\sendEmail.py " + filepath
    MsgBox (command)
    ReturnValue = Shell(command)

End Sub


