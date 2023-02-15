Attribute VB_Name = "ReadingCalc"
Option Explicit

Dim NextTick As Date, t As Date, PreviousTimerValue As Date, SelectedCell As String


Sub StartClock()
    If Range("B2").Value > 0 Then
        PreviousTimerValue = Range("B2").Value
    Else:
        PreviousTimerValue = 0
    End If
    t = Time
    Call ExcelStopWatch
End Sub

Private Sub ExcelStopWatch()
    Range("B2").Value = Format(Time - t + PreviousTimerValue, "hh:mm:ss")
    NextTick = Now + TimeValue("00:00:01")
    Application.OnTime NextTick, "ExcelStopWatch"
End Sub

Sub StopClock()
    If Range("B2").Value > 0 Then
        On Error Resume Next
        Application.OnTime earliesttime:=NextTick, procedure:="ExcelStopWatch", schedule:=False
        Call AddResultToTable(Application.InputBox("Enter student's name:"), Range("B2").Value, Application.InputBox("Enter words read correctly:"))
        Range("B2").Value = 0
    End If
End Sub

Sub ClearAll()
    On Error Resume Next
    Application.OnTime earliesttime:=NextTick, procedure:="ExcelStopWatch", schedule:=False
    Range("B2").Value = 0
    Dim table As Range
    Set table = ActiveSheet.ListObjects("ResultsTable").Range
    table.ListObject.DataBodyRange.Rows.Delete
End Sub

Private Sub AddResultToTable(StudentName As String, TimeTaken As Date, CorrectWords As Integer)

    Dim table As Range
    Set table = ActiveSheet.ListObjects("ResultsTable").Range

    Dim LastRow As Long

    LastRow = table.Find(What:="*", _
    After:=table.Cells(1), _
    Lookat:=xlPart, _
    LookIn:=xlFormulas, _
    SearchOrder:=xlByRows, _
    SearchDirection:=xlPrevious, _
    MatchCase:=False).Row - (table.Rows(1).Row - 2)

    table.Cells(LastRow, 1).Value = StudentName
    table.Cells(LastRow, 2).Value = TimeTaken
    table.Cells(LastRow, 3).Value = CorrectWords
    table.Cells(LastRow, 4).Value = Round(CorrectWords / ((Minute(TimeTaken) * 60 + Second(TimeTaken)) / 60), 0)

End Sub

