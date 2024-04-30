Attribute VB_Name = "Module1"

Public Sub ShowHelp(control As IRibbonControl)
    MsgBox "
    MUK:
    Az AL oszlopban lévő sorokat módosítja:
    1.'AK' oszlopban ('TR type'), a'Bank account'-ra szűrve, a TR számot 1-ről 0,5-re kell átírni
    2.'E' oszlopban, ahol a 'Document No'  S/0*-val kezdődik és a az 'AL' oszlopban a TR szám 1, ott a 0,5-re kell átírni a TR számot
    3.'AO' oszlopban, ahol a 'Ledger Entry Document No' értéke nem üres (tehát van benne vmilyen érték), ott a TR számot ki kell nullázni.
    RIVERSIDE:
    Az AL oszlopban lévő sorokat módosítja:
    1.'H' oszlopban rászűrök a 'BA-PS-ESCROWACC'-ra és a 'W' oszlopban <0 ( kivéve Bankkötség ezt kiveszem a szűrésből az 'L' oszlopban(Description)) és az AJ oszlopban (PS makes bank transfer) átírom a TR számot 1-re, illetve az AL oszlopban 1,5-re
    2.Rászűrök az AL szlopban ('TR Type') 'Bank account'-ra és az AJ oszlop=0 (PS makes bank transfer) 'AL' oszlopban átírom a TR számot 0,5-re
    3.'AL' oszlop ('TR type') rászűrök 'DEPR'-re TR számoz átírom 0,2-re AL oszlopban.", , "Leírás"
End Sub

Public Sub MUK_TransactionCounter(control As IRibbonControl)
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    
    ' A munkafüzet aktív munkalapjának kiválasztása
    Set ws = ActiveSheet
    
    ' Az utolsó sor meghatározása az AL oszlopban
    lastRow = ws.Cells(ws.Rows.Count, "AL").End(xlUp).Row
    
    ' Végigmegyünk minden soron
    For i = 1 To lastRow
        ' Az AK oszlopban 'Bank account' szűrése
        If ws.Cells(i, "AK").Value = "Bank account" And ws.Cells(i, "AL").Value = 1 Then
            ws.Cells(i, "AL").Value = 0.5
        End If
        
        ' Az E oszlopban 'Document No' szűrése
        If Left(ws.Cells(i, "E").Value, 3) = "S/0" And ws.Cells(i, "AL").Value = 1 Then
            ws.Cells(i, "AL").Value = 0.5
        End If
        
        ' Az AO oszlopban 'Ledger Entry Document No' ellenőrzése
        If ws.Cells(i, "AO").Value <> "" And ws.Cells(i, "AL").Value <> 0 Then
            ws.Cells(i, "AL").Value = 0
        End If
    Next i
End Sub
    
Public Sub Riverside_TransactionCounter(control As IRibbonControl)
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    
    Set ws = ActiveSheet
    lastRow = ws.Cells(ws.Rows.Count, "AL").End(xlUp).Row
    
    For i = 1 To lastRow
        ' Első feltétel
        If ws.Cells(i, "H").Value = "BA-PS-ESCROWACC" And ws.Cells(i, "W").Value < 0 And ws.Cells(i, "L").Value <> "Bankköltség" And ws.Cells(i, "AJ").Value = "PS makes bank transfer" Then
            ws.Cells(i, "AL").Value = 1.5
            ws.Cells(i, "AJ").Value = 1
        End If
        
        ' Második feltétel
        If ws.Cells(i, "AL").Value = "Bank account" And ws.Cells(i, "AJ").Value = 0 Then
            ws.Cells(i, "AL").Value = 0.5
        End If
        
        ' Harmadik feltétel
        If ws.Cells(i, "AL").Value = "DEPR" Then
            ws.Cells(i, "AL").Value = 0.2
        End If
    Next i
End Sub