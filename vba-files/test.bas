Attribute VB_Name = "Module1"

Public Sub Alert()
    MsgBox "message",,"title" 

End Sub

    Sub ApplyTemplateWithDropdown()
        Dim ruleSheet As Worksheet
        Dim dataSheet As Worksheet
        Dim dropDownSheet As Worksheet
        Dim transactionCol As Integer
        Dim lastRow As Long, lastRuleRow As Long, i As Long, j As Long
        Dim templateName As String
    
        Set dataSheet = ThisWorkbook.Sheets("Data")
        Set ruleSheet = ThisWorkbook.Sheets("Rules")
        Set dropDownSheet = ThisWorkbook.Sheets("Templates")
    
        ' Identify the "Transaction No." column
        For i = 1 To dataSheet.Cells(1, dataSheet.Columns.Count).End(xlToLeft).Column
            If dataSheet.Cells(1, i).Value = "Transaction No." Then
                transactionCol = i
                Exit For
            End If
        Next i
    
        lastRow = dataSheet.Cells(dataSheet.Rows.Count, transactionCol).End(xlUp).Row
        lastRuleRow = ruleSheet.Cells(ruleSheet.Rows.Count, "A").End(xlUp).Row
    
        ' Apply rules based on template
        For i = 2 To lastRow
            For j = 2 To lastRuleRow
                If ruleSheet.Cells(j, "A").Value = templateName And _
                   dataSheet.Cells(i, transactionCol).Value = ruleSheet.Cells(j, "B").Value Then
                    dataSheet.Cells(i, transactionCol).Value = ruleSheet.Cells(j, "C").Value
                End If
            Next j
        Next i
    
        MsgBox "Template applied successfully!"
    End Sub