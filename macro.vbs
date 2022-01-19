Sub switch_display_sheet()
' you can switch display sheet "DS1" and "DS2"

If (ThisWorkbook.Sheets("DS1").Visible = False) Then
    ThisWorkbook.Sheets("DS1").Visible = True   'display
    ThisWorkbook.Sheets("DS2").Visible = False  'undisplay
Else
    ThisWorkbook.Sheets("DS1").Visible = False
    ThisWorkbook.Sheets("DS2").Visible = True
End If

End Sub
