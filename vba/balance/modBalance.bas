Attribute VB_Name = "modBalance"
Option Explicit

' Seule macro restante du projet : lire un fichier de balance tiers et en détecter les
' colonnes n'est possible par aucune formule, sur aucune plateforme - un minimum de code
' est incontournable ici. Cette macro se contente d'écrire les colonnes brutes
' (Compte/Libellé/Année N-1/Année N) dans Balance_Data ; toute la classification
' (cycle, lecture écologique, variation) est ensuite calculée par formule sur cette
' feuille, que les données proviennent de cet import ou d'un collage manuel.
'
' Optionnel testFilePath : permet de piloter l'import sans la boîte de dialogue
' (utilisé par les tests headless - cf. tests\Verify-Balance.ps1)
Public Sub ImporterBalance(Optional ByVal testFilePath As String = "")
    Dim clientName As String
    clientName = Trim$(CStr(modUtil.ReadRangeValue("Bal_ClientName")))
    If Len(clientName) = 0 Then
        modUtil.NotifyUser "Le nom du client est requis.", vbExclamation, "Import balance"
        Exit Sub
    End If
    If Not IsNumeric(modUtil.ReadRangeValue("Bal_Year")) Then
        modUtil.NotifyUser "Veuillez saisir une année valide.", vbExclamation, "Import balance"
        Exit Sub
    End If

    Dim filePath As String
    If Len(testFilePath) > 0 Then
        filePath = testFilePath
    Else
        Dim fPick As Variant
        fPick = Application.GetOpenFilename("Fichiers balance (*.xlsx;*.xls;*.csv), *.xlsx;*.xls;*.csv", , "Choisir le fichier de balance")
        If VarType(fPick) = vbBoolean Then Exit Sub
        filePath = CStr(fPick)
    End If

    If Len(Dir$(filePath)) = 0 Then
        modUtil.NotifyUser "Fichier introuvable : " & filePath, vbCritical, "Import balance"
        Exit Sub
    End If

    Dim ext As String
    ext = LCase$(Mid$(filePath, InStrRev(filePath, ".") + 1))

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    Dim srcWb As Workbook
    If ext = "csv" Then
        Workbooks.OpenText Filename:=filePath, DataType:=xlDelimited, Comma:=True, Semicolon:=False, Tab:=False, Local:=True
        Set srcWb = ActiveWorkbook
    Else
        Set srcWb = Workbooks.Open(filePath, ReadOnly:=True)
    End If

    Dim srcSheet As Worksheet
    Set srcSheet = srcWb.Sheets(1)
    Dim dataArr As Variant
    dataArr = srcSheet.UsedRange.Value2

    Dim nRows As Long, nCols As Long
    nRows = UBound(dataArr, 1)
    nCols = UBound(dataArr, 2)

    Dim headers() As String
    ReDim headers(1 To nCols)
    Dim c As Long
    For c = 1 To nCols
        headers(c) = LCase$(Trim$(CStr(dataArr(1, c))))
    Next c

    Dim accountCol As Long, nameCol As Long
    accountCol = FindHeaderCol(headers, Array("compte", "account", "numero"))
    nameCol = FindHeaderCol(headers, Array("libellé", "libelle", "account name", "description"))

    If accountCol = 0 Or nameCol = 0 Then
        modUtil.NotifyUser "Format de fichier non reconnu - colonnes Compte/Libellé manquantes.", vbCritical, "Import balance"
        srcWb.Close False
        Application.ScreenUpdating = True: Application.DisplayAlerts = True
        Exit Sub
    End If

    Dim yearVals() As Long, yearColIdx() As Long, yearCount As Long
    ReDim yearVals(1 To nCols): ReDim yearColIdx(1 To nCols)
    yearCount = 0
    For c = 1 To nCols
        Dim hdr As String
        hdr = Trim$(CStr(dataArr(1, c)))
        If Len(hdr) = 4 And IsNumeric(hdr) Then
            Dim yv As Long
            yv = CLng(hdr)
            If yv >= 2000 And yv <= Year(Date) + 1 Then
                yearCount = yearCount + 1
                yearVals(yearCount) = yv
                yearColIdx(yearCount) = c
            End If
        End If
    Next c

    If yearCount = 0 Then
        modUtil.NotifyUser "Aucune colonne d'année détectée dans le fichier.", vbCritical, "Import balance"
        srcWb.Close False
        Application.ScreenUpdating = True: Application.DisplayAlerts = True
        Exit Sub
    End If

    Dim targetYear As Long
    targetYear = CLng(modUtil.ReadRangeValue("Bal_Year"))
    Dim curColIdx As Long, prevColIdx As Long
    curColIdx = 0: prevColIdx = 0
    For c = 1 To yearCount
        If yearVals(c) = targetYear Then curColIdx = yearColIdx(c)
        If yearVals(c) = targetYear - 1 Then prevColIdx = yearColIdx(c)
    Next c

    If curColIdx = 0 Then
        modUtil.NotifyUser "L'année " & targetYear & " n'existe pas dans le fichier.", vbCritical, "Import balance"
        srcWb.Close False
        Application.ScreenUpdating = True: Application.DisplayAlerts = True
        Exit Sub
    End If

    Dim startCell As Range
    Set startCell = modUtil.SafeRange("BalanceData_FirstRow")
    Dim dataWs As Worksheet
    Set dataWs = startCell.Worksheet
    Dim firstRow As Long, maxRows As Long
    firstRow = startCell.Row
    maxRows = 1000
    dataWs.Range(dataWs.Cells(firstRow, startCell.Column), dataWs.Cells(firstRow + maxRows - 1, startCell.Column + 3)).ClearContents

    Dim r As Long, outR As Long, importCount As Long
    outR = firstRow: importCount = 0
    For r = 2 To nRows
        If outR > firstRow + maxRows - 1 Then Exit For
        Dim accNum As String, accName As String
        accNum = Trim$(CStr(dataArr(r, accountCol)))
        accName = Trim$(CStr(dataArr(r, nameCol)))
        If Len(accNum) > 0 Then
            Dim curVal As Double, prevVal As Double
            curVal = modUtil.ParseFrenchNumber(dataArr(r, curColIdx))
            If prevColIdx > 0 Then
                prevVal = modUtil.ParseFrenchNumber(dataArr(r, prevColIdx))
            Else
                prevVal = 0
            End If

            ' La colonne Compte est déjà au format Texte (posé par New-BalanceDataSheet) -
            ' assigner la chaîne brute suffit à conserver les zéros de tête.
            dataWs.Cells(outR, startCell.Column).Value = accNum
            dataWs.Cells(outR, startCell.Column + 1).Value = accName
            dataWs.Cells(outR, startCell.Column + 2).Value = prevVal
            dataWs.Cells(outR, startCell.Column + 3).Value = curVal
            outR = outR + 1
            importCount = importCount + 1
        End If
    Next r

    srcWb.Close False
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

    modUtil.SafeRange("Bal_ImportStatus").Value = importCount & " ligne(s) importée(s) pour l'année " & targetYear & _
        " (comparée à " & (targetYear - 1) & "). Colonnes détectées : Compte=" & headers(accountCol) & ", Libellé=" & headers(nameCol) & "." & vbCrLf & _
        "La classification (cycle, lecture écologique) est automatique - consultez l'onglet Analyse."

    If Len(testFilePath) = 0 Then
        modUtil.NotifyUser importCount & " ligne(s) importée(s) avec succès.", vbInformation, "Import terminé"
    End If
End Sub

Private Function FindHeaderCol(ByRef headers() As String, ByVal candidates As Variant) As Long
    Dim n As Long, h As Long
    For n = LBound(candidates) To UBound(candidates)
        For h = LBound(headers) To UBound(headers)
            If InStr(1, headers(h), CStr(candidates(n))) > 0 Then
                FindHeaderCol = h
                Exit Function
            End If
        Next h
    Next n
    FindHeaderCol = 0
End Function
