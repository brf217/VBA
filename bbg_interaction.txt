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
    
' if the id pair = 1, call the initial_paste macro to seed the first row
  If Range("ID") = 1 Then Call initial_paste
  
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


