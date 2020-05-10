
Sub try()

 Dim last_row As Long
 ' define a  variable whose name is week with data type string
 Dim week As String
 ' define a  variable whose name iselec_lastrow with data type  long
 Dim elec_lastrow As Long
 ' define a  variable whose name iscash_lastrow with data typelong
 Dim cash_lastrow As Long
 'define a  variable whose name is prog_lastrow with data type long
 Dim prog_lastrow As Long
  'define a  variable whose name is opt_latrow with data typelong
 Dim opt_lastrow As Long
  'define a  variable whose name is dt with data type date
 Dim dt As Date
  

Worksheets("elec").Cells.Clear
Worksheets("cash").Cells.Clear
Worksheets("prog").Cells.Clear
Worksheets("opt").Cells.Clear

last_row = lastrownum2(Worksheets("weekly task list"))


Sheets("weekly task list").Range("A1").AutoFilter field:=7, Criteria1:="<>"
Worksheets("weekly task list").Range("a1:u" & last_row).Copy Destination:=Worksheets("ELEC").Range("A1")


Sheets("weekly task list").Range("A1").AutoFilter field:=8, Criteria1:="<>"
Worksheets("weekly task list").Range("a1:u" & last_row).Copy Destination:=Worksheets("cash").Range("A1")

Sheets("weekly task list").Range("A1").AutoFilter field:=9, Criteria1:="<>"
Worksheets("weekly task list").Range("a1:u" & last_row).Copy Destination:=Worksheets("prog").Range("A1")


Sheets("weekly task list").Range("A1").AutoFilter field:=10, Criteria1:="<>"
Worksheets("weekly task list").Range("a1:u" & last_row).Copy Destination:=Worksheets("opt").Range("A1")


Sheets("elec").Range("a1").EntireRow.Insert
   With Sheets("elec").Range("a1")
       .Value = current_month
       .NumberFormat = "@"
       .Value = "ELECTRONIC:"
       .Font.FontStyle = "bold"
       .Font.Size = 10
       .HorizontalAlignment = xlLeft
    End With

Sheets("cash").Range("a1").EntireRow.Insert
   With Sheets("cash").Range("a1")
       .Value = current_month
       .NumberFormat = "@"
       .Value = "CASH:"
       .Font.FontStyle = "bold"
       .Font.Size = 10
       .HorizontalAlignment = xlLeft
    End With
    
Sheets("prog").Range("a1").EntireRow.Insert
   With Sheets("prog").Range("a1")
       .Value = current_month
       .NumberFormat = "@"
       .Value = "PROGRAMS:"
       .Font.FontStyle = "bold"
       .Font.Size = 10
       .HorizontalAlignment = xlLeft
    End With
    
    Sheets("opt").Range("a1").EntireRow.Insert
   With Sheets("opt").Range("a1")
       .Value = current_month
       .NumberFormat = "@"
       .Value = "OPTIONS:"
       .Font.FontStyle = "bold"
       .Font.Size = 10
       .HorizontalAlignment = xlLeft
    End With
    


Sheets("elec").Columns("b").EntireColumn.Delete
Sheets("cash").Columns("b").EntireColumn.Delete
Sheets("prog").Columns("b").EntireColumn.Delete
Sheets("opt").Columns("b").EntireColumn.Delete

Sheets("elec").Columns("k").EntireColumn.Delete
Sheets("cash").Columns("k").EntireColumn.Delete
Sheets("prog").Columns("k").EntireColumn.Delete
Sheets("opt").Columns("k").EntireColumn.Delete

Sheets("elec").Columns("k").EntireColumn.Delete
Sheets("cash").Columns("k").EntireColumn.Delete
Sheets("prog").Columns("k").EntireColumn.Delete
Sheets("opt").Columns("k").EntireColumn.Delete

elec_lastrow = lastrownum2(Sheets("elec"))
cash_lastrow = lastrownum2(Sheets("cash"))
prog_lastrow = lastrownum2(Sheets("prog"))
opt_lastrow = lastrownum2(Sheets("opt"))

color_alt_rows2 (Sheets("elec").Range("a2:o" & elec_lastrow))
color_alt_rows2 (Sheets("cash").Range("a2:o" & cash_lastrow))
color_alt_rows2 (Sheets("prog").Range("a2:o" & prog_lastrow))
color_alt_rows2 (Sheets("opt").Range("a2:o" & opt_lastrow))

Sheets("elec").Range("a2:o2").Font.Color = RGB(83, 141, 213)
Sheets("cash").Range("a2:o2").Font.Color = RGB(83, 141, 213)
Sheets("prog").Range("a2:o2").Font.Color = RGB(83, 141, 213)
Sheets("opt").Range("a2:o2").Font.Color = RGB(83, 141, 213)

Sheets("elec").Range("a2:o2").Font.Color = RGB(0, 0, 0)
Sheets("cash").Range("a2:o2").Font.Color = RGB(0, 0, 0)
Sheets("prog").Range("a2:o2").Font.Color = RGB(0, 0, 0)
Sheets("opt").Range("a2:o2").Font.Color = RGB(0, 0, 0)

''from class the cell alignment **
Sheets("elec").Cells.HorizontalAlignment = xlLeft
Sheets("cash").Cells.HorizontalAlignment = xlLeft
Sheets("prog").Cells.HorizontalAlignment = xlLeft
Sheets("opt").Cells.HorizontalAlignment = xlLeft


Sheets("elec").Cells.VerticalAlignment = xlTop
Sheets("cash").Cells.VerticalAlignment = xlTop
Sheets("prog").Cells.VerticalAlignment = xlTop
Sheets("opt").Cells.VerticalAlignment = xlTop


'''''column width from class

'first column
Sheets("elec").Columns("a").ColumnWidth = 30
Sheets("cash").Columns("a").ColumnWidth = 30
Sheets("prog").Columns("a").ColumnWidth = 30
Sheets("opt").Columns("a").ColumnWidth = 30

Sheets("elec").Columns("b").ColumnWidth = 10
Sheets("cash").Columns("b").ColumnWidth = 10
Sheets("prog").Columns("b").ColumnWidth = 10
Sheets("opt").Columns("b").ColumnWidth = 10

Sheets("elec").Columns("c").ColumnWidth = 15
Sheets("cash").Columns("c").ColumnWidth = 15
Sheets("prog").Columns("c").ColumnWidth = 15
Sheets("opt").Columns("c").ColumnWidth = 15

Sheets("elec").Columns("d:j").ColumnWidth = 10
Sheets("cash").Columns("d:j").ColumnWidth = 10
Sheets("prog").Columns("d:j").ColumnWidth = 10
Sheets("opt").Columns("d:j").ColumnWidth = 10

Sheets("elec").Columns("k").ColumnWidth = 50
Sheets("cash").Columns("k").ColumnWidth = 50
Sheets("prog").Columns("k").ColumnWidth = 50
Sheets("opt").Columns("k").ColumnWidth = 50

Sheets("elec").Columns("k").WrapText = True
Sheets("cash").Columns("k").WrapText = True
Sheets("prog").Columns("k").WrapText = True
Sheets("opt").Columns("k").WrapText = True


Sheets("elec").Columns("o").ColumnWidth = 10
Sheets("cash").Columns("o").ColumnWidth = 10
Sheets("prog").Columns("o").ColumnWidth = 10
Sheets("opt").Columns("o").ColumnWidth = 10


''''from the class change the column height
Sheets("elec").Rows("1:" & elec_lastrow).RowHeight = 20
Sheets("cash").Rows("1:" & elec_lastrow).RowHeight = 20
Sheets("prog").Rows("1:" & elec_lastrow).RowHeight = 20
Sheets("opt").Rows("1:" & elec_lastrow).RowHeight = 20


dt = InputBox("please enter the date: ")
html_body = RangetoHTML(Sheets("elec").Range("a1:o" & elec_lastrow)) & "</table>" & "<br></br></br></br>" & "<table width ='70'>" & RangetoHTML(Sheets("cash").Range("a1:o" & cash_lastrow)) & "</table>" & "<br></br></br></br>" _
& RangetoHTML(Sheets("prog").Range("a1:o" & prog_lastrow)) & "</table>" & "<br></br></br></br>" _
& RangetoHTML(Sheets("opt").Range("a1:o" & opt_lastrow)) & "</table>" & "<br></br></br></br>"
            

Set outlookmail = New outlook.Application
Set Email = outlook.CreateItem(olMailItem)

tempfilepath = Environ$("temp") & "\"

With New FileSystemObject
If .FileExists(tempfilepath & "weekly_tracker.xlsx") Then
   .DeleteFile tempfilepath & "weekly_tracker.xlsx"
   
   End If
   
Workbooks("weeklytask.xlsm").Sheets("elec").Copy
ActiveWorkbook.SaveAs tempfilepath & "weekly_tracker.xlsx"
Workbooks("weeklytask.xlsm").Sheets("Cash").Copy After:=Workbooks("weekly_tracker.xlsx").Sheets(1)
Workbooks("weeklytask.xlsm").Sheets("prog").Copy After:=Workbooks("weekly_tracker.xlsx").Sheets(2)
Workbooks("weeklytask.xlsm").Sheets("opt").Copy After:=Workbooks("weekly_tracker.xlsx").Sheets(3)


Workbooks("weekly_tracker.xlsx").Save


 With Email
                .to = "hgao62@uwo.ca"   ' change this to any other people's email address you want to send(you can send to me if you want)
                .cc = "g471692309@gmail.com"  ' change this to any other people's email you want to copy
                .Subject = "title"
                .htmlbody = html_body
                .Display
            End With
            Set outlookmail = Nothing
            Set outlookapp = Nothing
            
            Workbooks("weekly_tracker.xlsx").Close savechanges:=True
        
        End With
    
End Sub





Function RangetoHTML(rng As Range)
' Changed by Ron de Bruin 28-Oct-2006
' Working in Office 2000-2010
    Dim fso As Object
    Dim ts As Object
    Dim TempFile As String
    Dim TempWB As Workbook
    Dim rngArea As Range
    Dim lngRow As Long
 
    TempFile = Environ$("temp") & "/" & Format(Now, "dd-mm-yy h-mm-ss") & ".htm"
 
    'Copy the range and create a new workbook to past the data in
'''    rng2.Copy
      rng.Copy
      
    Set TempWB = Workbooks.Add(1)
    With TempWB.Sheets(1)
            .Cells(1).PasteSpecial Paste:=8
            .Cells(1).PasteSpecial xlPasteValues, , False, False
            .Cells(1).PasteSpecial xlPasteFormats, , False, False
            .Cells(1).Select
            Application.CutCopyMode = False
            
    End With
 
    'Publish the sheet to a htm file
    With TempWB.PublishObjects.Add( _
        SourceType:=xlSourceRange, _
        Filename:=TempFile, _
        sheet:=TempWB.Sheets(1).Name, _
        Source:=TempWB.Sheets(1).UsedRange.Address, _
        HtmlType:=xlHtmlStatic)
        .Publish (True)
    End With
 
    'Read all data from the htm file into RangetoHTML
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.GetFile(TempFile).OpenAsTextStream(1, -2)
    RangetoHTML = ts.ReadAll
    ts.Close
    RangetoHTML = Replace(RangetoHTML, "align=center x:publishsource=", _
                          "align=left x:publishsource=")
 
    'Close TempWB
    TempWB.Close 0
    Kill TempFile
 
    'Delete the htm file we used in this function
'''    Kill TempFile
 
    Set ts = Nothing
    Set fso = Nothing
    Set TempWB = Nothing

    
End Function


'------------------------------------------------------------------------------
'------------------------------------------------------------------------------

Public Function lastrownum2(sheet As Worksheet) As Long
      If Application.WorksheetFunction.CountA(sheet.Cells) <> 0 Then
         lastrownum2 = sheet.Cells.Find(what:="*", _
                                       LookIn:=xlFormulas, _
                                       searchorder:=xlByRows, _
                                       searchdirection:=xlPrevious).Row
       Else
         lastrownum2 = 1
       End If
       
End Function

'------------------------------------------------------------------------------
'------------------------------------------------------------------------------


Sub color_alt_rows2(rng As Range)

     Application.ScreenUpdating = False
     rng.Interior.ColorIndex = xlNone
     rng.FormatConditions.Add Type:=xlExpression, Formula1:="=mod(row(),2)"
     rng.FormatConditions(1).Interior.Color = RGB(220, 230, 241)
     
End Sub




