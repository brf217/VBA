
' paste results for all other results except the first one
Sub paste_result()
    Range("paste_from").Select
    Selection.Copy
    Selection.End(xlDown).Select
    Selection.End(xlDown).Select
    ActiveCell.Offset(1, 0).Range("A1").Select
    Selection.PasteSpecial Paste:=xlPasteValues
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Application.CutCopyMode = False
    Calculate
End Sub


' sub to paste the initial pair with ID = 1
Sub first_row_paste()
    Range("paste_from").Select
    Selection.Copy
    Range("initial_paste").Select
     Selection.PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    Calculate
End Sub


' call the functions above as needed to fill the worksheet
Sub main_function()

    Range("paste_range").ClearContents
    Range("ID") = 1

' turn off screen updating
Application.ScreenUpdating = False

' boilerplate block to allow bloomberg to finish its calcs
    While Range("calcs_ready_test") <> 0
        Application.Calculate
        Application.RTD.RefreshData
        DoEvents
    Wend
    
' if the id pair = 1, call the first_row_paste macro to seed the first row
  If Range("ID") = 1 Then Call first_row_paste
  
  While Range("ID") < Range("max_pair_id")
    Range("ID") = Range("ID") + 1
      While Range("calcs_ready_test") <> 0
          Application.Calculate
          Application.RTD.RefreshData
          DoEvents
      Wend
    Call paste_result
    Wend
    
' Turn screen updating back on
Application.ScreenUpdating = True

End Sub

Sub call_loop()

' Define scenarios that will be run
Dim input_arr(3) As String
input_arr(1) = "SOFR-FWD"
input_arr(2) = "CUSTOM"
input_arr(3) = "FLAT"

' Start of loop through scenarios and maturities

For i = 1 To 3
    For j = 1 To 2
        ' Go to the control sheet and set the inputs
        Sheets("Control").Select
        Range("curve_selection") = input_arr(i)
        Range("years_back_fr_mat") = j
        
        ' Call the main_function to run the pairs for the selected scenario
        Call main_function
        
        ' Paste the completed run to the proper sheet
        Range("paste_area").Copy
        
        ' Construct and select the sheet name based on the variable names
        Sheets(input_arr(i) + "-" + CStr(j)).Activate
        Range("B5").Select
            
        ' Paste into the destination sheet in the cell selected above
        Selection.PasteSpecial Paste:=xlPasteValues
        Selection.PasteSpecial Paste:=xlPasteFormats
        Selection.PasteSpecial Paste:=xlPasteColumnWidths
        
    Next j
Next i
   
End Sub

