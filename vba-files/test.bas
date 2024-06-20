Public Sub ShowHelp(control As IRibbonControl)
    Dim message As String
    message = "MUK:" & vbCrLf & _
              "    Az AL oszlopban lévo sorokat módosítja:" & vbCrLf & _
              "    1. 'AK' oszlopban ('TR type'), a 'Bank account'-ra szurve, a TR számot 1-rol 0,5-re kell átírni" & vbCrLf & _
              "    2. 'E' oszlopban, ahol a 'Document No'  S/0*-val kezdodik és az 'AL' oszlopban a TR szám 1, ott a 0,5-re kell átírni a TR számot" & vbCrLf & _
              "    3. 'AO' oszlopban, ahol a 'Ledger Entry Document No' értéke nem üres (tehát van benne vmilyen érték), ott a TR számot ki kell nullázni." & vbCrLf & _
              "RIVERSIDE:" & vbCrLf & _
              "    Az AL oszlopban lévo sorokat módosítja:" & vbCrLf & _
              "    1. 'H' oszlopban rászurök a 'BA-PS-ESCROWACC'-ra és a 'W' oszlopban <0 ( kivéve Bankkötség – ezt kiveszem a szurésbol az 'L' oszlopban(Description)) és az AJ oszlopban (PS makes bank transfer) átírom a TR számot 1-re, illetve az AL oszlopban 1,5-re" & vbCrLf & _
              "    2. Rászurök az AK szlopban ('TR Type') 'Bank account'-ra és az AJ oszlop=0 (PS makes bank transfer) 'AL' oszlopban átírom a TR számot 0,5-re" & vbCrLf & _
              "    3. 'AK' oszlop ('TR type') rászurök 'DEPR'-re TR számoz átírom 0,2-re AL oszlopban."
              
    MsgBox message, vbInformation, "Részletes Funkció Leírás"
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
        If ws.Cells(i, "A").Value <> "uniqueID" And ws.Cells(i, "A").Value <> "" Then
            ' Az AK oszlopban 'Bank account' szurése
            If ws.Cells(i, "AK").Value = "Bank Account" And ws.Cells(i, "AL").Value = 1 Then
                ws.Cells(i, "AL").Value = 0.5
            End If
            
            ' Az E oszlopban 'Document No' szurése
            If Left(ws.Cells(i, "E").Value, 3) = "S/0" And ws.Cells(i, "AL").Value = 1 Then
                ws.Cells(i, "AL").Value = 0.5
            End If
            
            ' Az AO oszlopban 'Ledger Entry Document No' ellenorzése
            If ws.Cells(i, "AO").Value <> "" And ws.Cells(i, "AL").Value <> 0 Then
                ws.Cells(i, "AL").Value = 0
            End If
        End If
    Next i
    MsgBox "Kész", , ""
End Sub
    
Public Sub Riverside_TransactionCounter(control As IRibbonControl)
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    
    Set ws = ActiveSheet
    lastRow = ws.Cells(ws.Rows.Count, "AL").End(xlUp).Row
    
    For i = 1 To lastRow
        If ws.Cells(i, "A").Value <> "uniqueID" And ws.Cells(i, "A").Value <> "" Then
            ' Elso feltétel
            If ws.Cells(i, "H").Value = "BA-PS-ESCROWACC" And ws.Cells(i, "W").Value < 0 And (InStr(ws.Cells(i, "L").Value, "Bankköltség") > 0 Or InStr(ws.Cells(i, "L").Value, "Bankktg")) Then
                ws.Cells(i, "AL").Value = 1.5
                ws.Cells(i, "AJ").Value = 1
            End If
            
            ' Második feltétel
            If ws.Cells(i, "AK").Value = "Bank account" And ws.Cells(i, "AJ").Value = 0 Then
                ws.Cells(i, "AL").Value = 0.5
            End If
            
            ' Harmadik feltétel
            If ws.Cells(i, "AK").Value = "DEPR" Then
                ws.Cells(i, "AL").Value = 0.2
            End If
        End If

    Next i

    MsgBox "Kész", , ""
End Sub
