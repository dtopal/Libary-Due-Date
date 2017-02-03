Attribute VB_Name = "Module1"
Function DueDate(InputDay As Date)
    
    'Takes InputDay and adds 30 days
    DueDate = DateAdd("d", 30, InputDay)
    
    'If DueDate is on a Sunday, changes due date to next day
    Dim DueDay As Integer
    DueDay = Weekday(DueDate)
    If DueDay = 1 Then
        DueDate = DateAdd("d", 1, DueDate)
    End If
    
    'Build queue of Holiday dates
    'Set Sheet = ActiveWorkbook.Sheets("Holidays").Activate
    Set HolidayList = CreateObject("System.Collections.Queue")
    
    Dim dateRange As Range, cell As Range
    Set dateRange = Worksheets("Holidays").Range("A1:A365")
    
    For Each cell In dateRange
        If cell.Value <> "" Then
            'Dim currDate As Date
            'currDate = DateValue(cell.Value)
            'HolidayList.Enqueue currDate
            HolidayList.Enqueue cell.Value
        End If
    Next
    
    'Check if date is in the HolidayList
    Do While HolidayList.Contains(DueDate)
        DueDate = DateAdd("d", 1, DueDate)
    Loop
    
    
    
    
    
    
    'Formats DueDate as mm/dd/yyyy
    DueDate = Format(DueDate, "General Date")
End Function
