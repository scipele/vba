Sub acct()
' Changes the current selection to accounting format
    Selection.NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    Selection.NumberFormat = "_(* #,##0_);_(* (#,##0);_(* ""-""_);_(@_)"
End Sub