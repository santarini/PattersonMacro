Sub ServiceSortMain()
    Call CatSort
    Call StatusSort
    Call DateValidator
    Call CurrencyValidator
    Call DataRecode
    Call createPivotTable
    Call generateReport
    MsgBox "Success!", vbInformation
End Sub
