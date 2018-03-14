VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DataManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Sub Test()
   Dim wb As Workbook
   Set wb = Workbooks.Add
End Sub



Public Sub Test2()
    Dim path As String
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    Dim day_nums
    day_nums = Array(31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31)
    For i = 1 To 12
        path = "C:\Users\whens\Desktop\"
        path = path & CStr(i) & ".xlsx"
        Set Wk = Workbooks.Add
        With Wk
            .SaveAs Filename:=path
            num = day_nums(i - 1)
            For d = 1 To num
                Set ws = .Sheets.Add(After:=.Sheets(.Sheets.Count))
                ws.Name = CStr(d)
            Next d
            .Sheets(1).Delete
            .Save
            .Close
        End With
    Next i
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.DisplayAlerts = True
End Sub
