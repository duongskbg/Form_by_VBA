VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "DD_SmartTools"
   ClientHeight    =   4365
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10695
   OleObjectBlob   =   "DD_form_V3.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Function SheetExists(sname) As Boolean
'   Returns TRUE if sheet exists in the active workbook
    Dim X As Object
    On Error Resume Next
    Set X = ActiveWorkbook.Sheets(sname)
    If Err = 0 Then SheetExists = True _
        Else SheetExists = False
End Function

Private Sub btn_cancel_Click()
    frame_col.Visible = False
End Sub

Private Sub btn_chcol_Click()
    frame_col.Visible = True
    'Dim ws As Worksheet
    'Dim lc, lr, i As Long
    'Set ws = ActiveSheet
    'Sheets("SD-Bueller Zero Report").Select
    'Columns("AL:AL").Select
    'Selection.Copy
    'If cb_col1.Value <> "" Then

        'Columns(cb_col1.Value).Select
        'Selection.PasteSpecial Paste:=xlPasteFormats
        'Columns(cb_col2.Value).Select
        'Selection.PasteSpecial Paste:=xlPasteFormats
        'Columns(cb_col3.Value).Select
        'Selection.PasteSpecial Paste:=xlPasteFormats
    'End If
    'If cb_col2.Value <> "" Then
        'Columns(cb_col2.Value).Select
        'Selection.PasteSpecial Paste:=xlPasteFormats
    'End If
    'If cb_col3.Value <> "" Then
        'Columns(cb_col3.Value).Select
        'Selection.PasteSpecial Paste:=xlPasteFormats
    'End If
    'If cb_col1.Value = "" And cb_col2.Value = "" And cb_col3.Value = "" Then
        'Columns("AL:AL").Select
        'Selection.Copy
        'Columns("I:AK").Select
        'Selection.PasteSpecial Paste:=xlPasteFormats
    'End If


End Sub

Private Sub btn_ok_Click()
    If cb_col1.Value <> "" Then
        UserForm1.opt_col1.Caption = "Column 1: " + cb_col1.Value
    End If
    If cb_col2.Value <> "" Then
        UserForm1.opt_col2.Caption = "Column 2: " + cb_col2.Value
    End If
    If cb_col3.Value <> "" Then
        UserForm1.opt_col3.Caption = "Column 3: " + cb_col3.Value
    End If
    'MsgBox cb_col1.Value
    Dim ws As Worksheet
    Dim lc, lr, i As Long
    Set ws = ActiveSheet
    Sheets("SD-Bueller Zero Report").Select
    lc = ws.Cells(7, ws.Columns.Count).End(xlToLeft).Column
    lr = ws.Cells(ws.Rows.Count, "D").End(xlUp).Row
    If cb_col1.Value <> "" Then
        'Columns(cb_col1.Value).Select
        'Selection.Interior.Color = 65535
        Cells(1, 1) = cb_col1.ListIndex + 9
    End If
    If cb_col2.Value <> "" Then
        'Columns(cb_col2.Value).Select
        'Selection.Interior.Color = 65535
        Cells(1, 2) = cb_col2.ListIndex + 9
        UserForm1.opt_col2.Enabled = True
    End If
    If cb_col3.Value <> "" Then
        'Columns(cb_col3.Value).Select
        'Selection.Interior.Color = 65535
        Cells(1, 3) = cb_col3.ListIndex + 9
        UserForm1.opt_col3.Enabled = True
        UserForm1.opt_col3.Value = True
    End If
    
    If cb_col2.Value = "" Then
        UserForm1.opt_col1.Value = True
        UserForm1.opt_col2.Enabled = False
        UserForm1.opt_col3.Enabled = False
    ElseIf cb_col3.Value = "" Then
        UserForm1.opt_col2.Value = True
        UserForm1.opt_col3.Enabled = False
    End If
    
    frame_col.Visible = False

    
    
    
    
End Sub

Private Sub btn_quit_Click()
    Unload Me
End Sub



Private Sub btn_start_Click()
Dim ws As Worksheet
Dim rng, chon1, chon2, color_rg As Range
Dim k, lr, temp, dem, DK, i, val, val2, l, dred, lc As Long

If SheetExists("New_Data") Then
    Application.DisplayAlerts = False
    Sheets("New_Data").Select
    ActiveWindow.SelectedSheets.Delete
    Application.DisplayAlerts = True
End If
Sheets.Add
ActiveSheet.Name = "New_Data"
Sheets("SD-Bueller Zero Report").Select
'Sheet1.Select
Range(Cells(1, 1), Cells(3, 53)).Select
Selection.Copy Destination:=Sheets("New_Data").Cells(1, 1)
lr = Cells(Rows.Count, "D").End(xlUp).Row
lc = 53
If opt_col2.Enabled = False Then
    temp = 1
ElseIf opt_col3.Enabled = False Then
    temp = 2
Else
    temp = 3
End If

For i = 7 To lr Step 5
    DK = 0
    dem = 0
    Set color_rg = Range(Cells(i - 3, 2), Cells(i + 1, 2))
    For k = 1 To temp
        val = Cells(1, k).Value2
        DK = Cells(i, val).Value2
        
        If DK < 0 Then
            dem = dem + 1
        End If
        If dem = 0 Then
            color_rg.Interior.Color = vbGreen
        Else
            color_rg.Interior.Color = vbRed
        End If
    Next k
Next i

If opt_col1.Value = True Then
    val2 = Cells(1, 1).Value2
ElseIf opt_col2.Value = True Then
    val2 = Cells(1, 2).Value2
Else
    val2 = Cells(1, 3).Value2
End If

Dim dem_pt As Long
dem_pt = 0
For l = 7 To lr Step 5
    Set color_rg = Range(Cells(l - 3, 2), Cells(l + 1, 2))
    If color_rg.Interior.Color = vbRed Then
        dred = l
        dem_pt = dem_pt + 1
    End If
Next l

'gan gia tri
Dim arr(5000) As Long
Dim ar, ck, stt, bb, new_re, last_stt As Long
Dim chonn, table2 As Range
ar = 0
For ck = 7 To dred Step 5
    
    Set color_rg = Range(Cells(ck - 3, 2), Cells(ck + 1, 2))
    If color_rg.Interior.Color = vbRed And ar < dem_pt Then
        arr(ar) = Cells(ck, val2).Value2
        ar = ar + 1
    End If
    
Next ck


    
    
'sap xep
Dim arr2() As Long
If opt_1.Value = True Then
    MergeSort arr, 0, dem_pt - 1
ElseIf opt_2.Value = True Then
    MergeSort2 arr, 0, dem_pt - 1
End If

bb = 0
ar = 0
new_re = 4
last_stt = 3 + (dem_pt * 5)

For i = 0 To (dred - 3)
    If arr(i) <> arr(i + 1) Then
        For k = 7 To lr Step 5
            Set color_rg = Range(Cells(k - 3, 2), Cells(k + 1, 2))
           'Set ch = Range(Cells(4, 1), Cells(4, lc))
            Set chonn = Range(Cells(k - 3, 1), Cells(k + 1, lc))
            DK = Cells(k, val2).Value2
            If color_rg.Interior.Color = vbRed And arr(i) = DK Then
            
            'Range("A1").Copy Destination:=Range("A2")'
                chonn.Select
                Selection.Copy Destination:=Sheets("New_Data").Cells(new_re, 1)
                'Sheets("New_Data").Select
                'Cells(new_re,1).Select
                'ActiveSheet.Paste
                
                new_re = new_re + 5
                Sheets("SD-Bueller Zero Report").Select
            End If
            
        Next k
    End If
Next i

Set table2 = Range(Cells(4, 1), Cells(lr, lc))
'Range("A4").Activate
table2.Select
Selection.AutoFilter
table2.AutoFilter Field:=2, Criteria1:=RGB(0, _
    255, 0), Operator:=xlFilterCellColor


'update stt
Sheets("New_Data").Select
stt = 1
For i = 4 To last_stt Step 5
    Cells(i, 2).Value2 = stt
    stt = stt + 1
Next i


End Sub





Private Sub opt_1_Click()
    If opt_1.Value = True Then
        opt_1.BackColor = &H80FFFF
        opt_2.BackColor = &H8000000F
    End If
End Sub


Private Sub opt_2_Click()
    If opt_2.Value = True Then
        opt_2.BackColor = &H80FFFF
        opt_1.BackColor = &H8000000F
    End If
End Sub

Private Sub opt_col1_Click()
    If opt_col1.Value = True Then
        opt_col1.BackColor = &H80FFFF
        opt_col2.BackColor = &H8000000F
        opt_col3.BackColor = &H8000000F
        
    End If
End Sub


Private Sub opt_col2_Click()
    If opt_col2.Value = True Then
        opt_col2.BackColor = &H80FFFF
        opt_col1.BackColor = &H8000000F
        opt_col3.BackColor = &H8000000F
        
    End If
End Sub


Private Sub opt_col3_Click()
    If opt_col3.Value = True Then
        opt_col3.BackColor = &H80FFFF
        opt_col1.BackColor = &H8000000F
        opt_col2.BackColor = &H8000000F
        
    End If
End Sub


Private Sub UserForm_Initialize()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    Sheets("SD-Bueller Zero Report").Select
    Cells(1, 1) = ""
    Cells(1, 2) = ""
    Cells(1, 3) = ""
    opt_col1.Caption = "Column 1"
    opt_col2.Caption = "Column 2"
    opt_col3.Caption = "Column 3"
    opt_col2.Enabled = True
    opt_col3.Enabled = True
    LoadCombos cb_col1
    LoadCombos cb_col2
    LoadCombos cb_col3
    opt_1.Value = True
    opt_col3.Value = True
    
End Sub

Public Sub LoadCombos(pcb As ComboBox)
With pcb
    .Clear
    .AddItem "I"
    .AddItem "J"
    .AddItem "K"
    .AddItem "L"
    .AddItem "M"
    .AddItem "N"
    .AddItem "O"
    .AddItem "P"
    .AddItem "Q"
    .AddItem "R"
    .AddItem "S"
    .AddItem "T"
    .AddItem "U"
    .AddItem "V"
    .AddItem "W"
    .AddItem "X"
    .AddItem "Y"
    .AddItem "Z"
    .AddItem "AA"
    .AddItem "AB"
    .AddItem "AC"
    .AddItem "AD"
    .AddItem "AE"
    .AddItem "AF"
    .AddItem "AG"
    .AddItem "AH"
    .AddItem "AI"
    .AddItem "AJ"
    .AddItem "AK"
    .AddItem "AL"
    '.Text = "Choose"
End With
End Sub

'=================================================================================================


Public Sub MergeSort(ByRef list() As Long, ByVal first_index As Long, ByVal last_index As Long)
    Dim middle As Long


    If (last_index > first_index) Then
        ' Recursively sort the two halves of the list.
        middle = (first_index + last_index) \ 2
        MergeSort list, first_index, middle
        MergeSort list, middle + 1, last_index


        ' Merge the results.
        Merge list, first_index, middle, last_index
    End If
End Sub

Public Sub MergeSort2(ByRef list() As Long, ByVal first_index As Long, ByVal last_index As Long)
    Dim middle As Long


    If (last_index > first_index) Then
        ' Recursively sort the two halves of the list.
        middle = (first_index + last_index) \ 2
        MergeSort2 list, first_index, middle
        MergeSort2 list, middle + 1, last_index


        ' Merge the results.
        Merge2 list, first_index, middle, last_index
    End If
End Sub

Public Sub Merge2(ByRef list() As Long, ByVal beginning As Long, ByVal middle As Long, ByVal ending As Long)
Dim temp_Array() As Long
Dim temp As Long
Dim counterA As Long
Dim counterB As Long
Dim counterMain As Long
Dim i As Long
Dim tempCounter As Long
Dim n As Long


    ' Copy the array into a temporary array.
    ReDim temp_Array(beginning To ending)
'    CopyMemory temp_Array(beginning), list(beginning), _
'        (ending - beginning + 1) * Len(list(beginning))
    For i = beginning To ending
        temp_Array(i) = list(i)
    Next i
        
       


    ' counterA and counterB mark the next item to save
    ' in the first and second halves of the list.
    counterA = beginning
    counterB = middle + 1


    ' counterMain is the index where we will put the
    ' next item in the merged list.
    counterMain = beginning
    Do While (counterA <= middle) And (counterB <= ending)
        ' Find the smaller of the two items at the front
        ' of the two sublists.
        If (temp_Array(counterA) >= temp_Array(counterB)) _
            Then
            ' The smaller value is in the first half.
            list(counterMain) = temp_Array(counterA)
            counterA = counterA + 1
        Else
            ' The smaller value is in the second half.
            list(counterMain) = temp_Array(counterB)
            counterB = counterB + 1
        End If
        counterMain = counterMain + 1
    Loop


    ' Copy any remaining items from the first half.
    If counterA <= middle Then
'        CopyMemory list(counterMain), temp_Array(counterA), _
'            (middle - counterA + 1) * Len(list(beginning))


        n = 0
        For i = counterA To middle
            list(counterMain + n) = temp_Array(i)
            n = n + 1
        Next
            


        
    End If


    ' Copy any remaining items from the second half.
    If counterB <= ending Then
    
        n = 0
        For i = counterB To ending
            list(counterMain + n) = temp_Array(i)
            n = n + 1
        Next


    End If
End Sub

' Merge two sorted sublists.
Public Sub Merge(ByRef list() As Long, ByVal beginning As Long, ByVal middle As Long, ByVal ending As Long)
Dim temp_Array() As Long
Dim temp As Long
Dim counterA As Long
Dim counterB As Long
Dim counterMain As Long
Dim i As Long
Dim tempCounter As Long
Dim n As Long


    ' Copy the array into a temporary array.
    ReDim temp_Array(beginning To ending)
'    CopyMemory temp_Array(beginning), list(beginning), _
'        (ending - beginning + 1) * Len(list(beginning))
    For i = beginning To ending
        temp_Array(i) = list(i)
    Next i
        
       


    ' counterA and counterB mark the next item to save
    ' in the first and second halves of the list.
    counterA = beginning
    counterB = middle + 1


    ' counterMain is the index where we will put the
    ' next item in the merged list.
    counterMain = beginning
    Do While (counterA <= middle) And (counterB <= ending)
        ' Find the smaller of the two items at the front
        ' of the two sublists.
        If (temp_Array(counterA) <= temp_Array(counterB)) _
            Then
            ' The smaller value is in the first half.
            list(counterMain) = temp_Array(counterA)
            counterA = counterA + 1
        Else
            ' The smaller value is in the second half.
            list(counterMain) = temp_Array(counterB)
            counterB = counterB + 1
        End If
        counterMain = counterMain + 1
    Loop


    ' Copy any remaining items from the first half.
    If counterA <= middle Then
'        CopyMemory list(counterMain), temp_Array(counterA), _
'            (middle - counterA + 1) * Len(list(beginning))


        n = 0
        For i = counterA To middle
            list(counterMain + n) = temp_Array(i)
            n = n + 1
        Next
            


        
    End If


    ' Copy any remaining items from the second half.
    If counterB <= ending Then
    
        n = 0
        For i = counterB To ending
            list(counterMain + n) = temp_Array(i)
            n = n + 1
        Next


    End If
End Sub

'=================================================================================================
