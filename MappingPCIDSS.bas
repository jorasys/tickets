'==============================================================================
' MACRO : RemplirPlanDeControle
' Logique lettre->rang :
'   3.5.1.a -> cle "3.5.1|1"  (1ere occurrence de 3.5.1 dans le Plan)
'   3.5.1.b -> cle "3.5.1|2"  (2eme occurrence)
'   3.5.1   -> cle "3.5.1|0"  (identifiant sans lettre, rang 0 = traitement special)
'==============================================================================

Private Const SHEET_PLAN      As String = "Plan de controle"
Private Const SHEET_DOCS      As String = "Liste des Docs"
Private Const PLAN_ID_COL     As Long = 3
Private Const PLAN_RESULT_COL As Long = 22
Private Const PLAN_START_ROW  As Long = 6
Private Const DOCS_ID_COL     As Long = 8
Private Const DOCS_FILE_COL   As Long = 3
Private Const DOCS_DATE_COL   As Long = 5
Private Const DOCS_START_ROW  As Long = 2

'==============================================================================
' Structure logique retournee par ParseIdentifiant :
'   sBase  = partie numerique  ex. "3.5.1"
'   iRang  = 0 si pas de lettre, 1 pour 'a', 2 pour 'b', etc.
'==============================================================================
Private Type TPciId
    sBase  As String
    iRang  As Integer   ' 0 = sans lettre suffixe
End Type

'------------------------------------------------------------------------------
Public Sub RemplirPlanDeControle()

    Dim wsPlan As Worksheet
    Dim wsDocs As Worksheet

    On Error Resume Next
    Set wsPlan = ThisWorkbook.Worksheets(SHEET_PLAN)
    Set wsDocs = ThisWorkbook.Worksheets(SHEET_DOCS)
    On Error GoTo 0

    If wsPlan Is Nothing Then
        MsgBox "Feuille introuvable : """ & SHEET_PLAN & """", vbCritical: Exit Sub
    End If
    If wsDocs Is Nothing Then
        MsgBox "Feuille introuvable : """ & SHEET_DOCS & """", vbCritical: Exit Sub
    End If

    Dim lastRowDocs As Long
    Dim lastRowPlan As Long
    lastRowDocs = DernierelignePleine(wsDocs, DOCS_ID_COL)
    lastRowPlan = DernierelignePleine(wsPlan, PLAN_ID_COL)

    If lastRowDocs < DOCS_START_ROW Then
        MsgBox "Aucune donnee dans la colonne H de """ & SHEET_DOCS & """.", vbInformation
        Exit Sub
    End If
    If lastRowPlan < PLAN_START_ROW Then
        MsgBox "Aucune donnee dans la colonne C de """ & SHEET_PLAN & """.", vbInformation
        Exit Sub
    End If

    ' -------------------------------------------------------------------------
    ' Etape 1 : Construire le dictionnaire  cle -> "Filename(Date)"
    '   Cle = "3.5.1|1"  pour 3.5.1.a
    '         "3.5.1|2"  pour 3.5.1.b
    '         "3.5.1|0"  pour 3.5.1  (sans lettre)
    ' -------------------------------------------------------------------------
    Dim dictMap As Object
    Set dictMap = CreateObject("Scripting.Dictionary")
    dictMap.CompareMode = vbTextCompare

    Dim iDoc   As Long
    Dim cellH  As String
    Dim sFile  As String
    Dim sDate  As String
    Dim sEntry As String
    Dim aIds() As String
    Dim j      As Integer
    Dim pid    As TPciId
    Dim sCle   As String

    For iDoc = DOCS_START_ROW To lastRowDocs

        cellH = Trim(CStr(wsDocs.Cells(iDoc, DOCS_ID_COL).Value))
        If cellH = "" Then GoTo NextDoc

        sFile = Trim(CStr(wsDocs.Cells(iDoc, DOCS_FILE_COL).Value))
        sDate = Trim(CStr(wsDocs.Cells(iDoc, DOCS_DATE_COL).Value))
        sEntry = sFile & " (" & sDate & ")"

        ' Normalisation des separateurs
        cellH = Replace(cellH, Chr(13) & Chr(10), Chr(10))
        cellH = Replace(cellH, Chr(13), Chr(10))
        cellH = Replace(cellH, ",", Chr(10))
        cellH = Replace(cellH, ";", Chr(10))
        aIds = Split(cellH, Chr(10))

        For j = 0 To UBound(aIds)

            pid = ParseIdentifiant(Trim(aIds(j)))
            If pid.sBase = "" Then GoTo NextId

            sCle = pid.sBase & "|" & pid.iRang

            Debug.Print "  [Doc " & iDoc & "] '" & Trim(aIds(j)) & _
                        "' -> base='" & pid.sBase & "'  rang=" & pid.iRang & _
                        "  cle='" & sCle & "'"

            If dictMap.Exists(sCle) Then
                dictMap(sCle) = dictMap(sCle) & Chr(10) & sEntry
            Else
                dictMap(sCle) = sEntry
            End If

NextId:
        Next j
NextDoc:
    Next iDoc

    Debug.Print "Dictionnaire : " & dictMap.Count & " cle(s)"

    ' -------------------------------------------------------------------------
    ' Etape 2 : Parcourir le Plan de controle
    '   - Pour chaque ligne, extraire la base de l'identifiant
    '   - Compter les occurrences de cette base (compteur incremente)
    '   - Chercher la cle "base|N" dans le dictionnaire
    '   - Si rang 0 existe aussi (identifiant sans lettre), l'appliquer a toutes
    '     les occurrences de la base (comportement de repli)
    ' -------------------------------------------------------------------------
    Dim dictOccurrences As Object
    Set dictOccurrences = CreateObject("Scripting.Dictionary")
    dictOccurrences.CompareMode = vbTextCompare

    Dim iPlan         As Long
    Dim sPlanRaw      As String
    Dim planPid       As TPciId
    Dim iOccurrence   As Integer
    Dim sCleRang      As String
    Dim sCleBase      As String   ' cle rang 0 (identifiant sans lettre)
    Dim sCellActuelle As String
    Dim sAjout        As String
    Dim nbMaj         As Long
    nbMaj = 0

    For iPlan = PLAN_START_ROW To lastRowPlan

        sPlanRaw = Trim(CStr(wsPlan.Cells(iPlan, PLAN_ID_COL).Value))
        If sPlanRaw = "" Then GoTo NextPlan

        ' Parser l'identifiant de la ligne du Plan
        ' (le Plan ne contient normalement pas de lettre, mais on parse quand meme)
        planPid = ParseIdentifiant(sPlanRaw)
        If planPid.sBase = "" Then GoTo NextPlan

        ' Incrementer le compteur d'occurrences pour cette base
        If dictOccurrences.Exists(planPid.sBase) Then
            dictOccurrences(planPid.sBase) = dictOccurrences(planPid.sBase) + 1
        Else
            dictOccurrences(planPid.sBase) = 1
        End If
        iOccurrence = dictOccurrences(planPid.sBase)

        ' Construire les deux cles candidates
        sCleRang = planPid.sBase & "|" & iOccurrence   ' ex. "3.5.1|2"
        sCleBase = planPid.sBase & "|0"                ' ex. "3.5.1|0" (sans lettre)

        ' Determiner la valeur a ajouter (priorite au rang exact, repli sur rang 0)
        sAjout = ""
        If dictMap.Exists(sCleRang) Then
            sAjout = dictMap(sCleRang)
            Debug.Print "  [Plan " & iPlan & "] '" & sPlanRaw & _
                        "' occ=" & iOccurrence & " -> cle '" & sCleRang & "' TROUVEE"
        ElseIf dictMap.Exists(sCleBase) Then
            sAjout = dictMap(sCleBase)
            Debug.Print "  [Plan " & iPlan & "] '" & sPlanRaw & _
                        "' occ=" & iOccurrence & " -> repli cle '" & sCleBase & "'"
        Else
            Debug.Print "  [Plan " & iPlan & "] '" & sPlanRaw & _
                        "' occ=" & iOccurrence & " -> aucune correspondance"
            GoTo NextPlan
        End If

        ' Ecriture dans la cellule cible
        sCellActuelle = Trim(CStr(wsPlan.Cells(iPlan, PLAN_RESULT_COL).Value))
        If sCellActuelle = "" Then
            wsPlan.Cells(iPlan, PLAN_RESULT_COL).Value = sAjout
        Else
            wsPlan.Cells(iPlan, PLAN_RESULT_COL).Value = sCellActuelle & Chr(10) & sAjout
        End If
        wsPlan.Cells(iPlan, PLAN_RESULT_COL).WrapText = True
        nbMaj = nbMaj + 1

NextPlan:
    Next iPlan

    Debug.Print "=== Fin : " & nbMaj & " cellule(s) mise(s) a jour ==="
    MsgBox nbMaj & " cellule(s) du Plan de controle mise(s) a jour.", _
           vbInformation, "Mapping PCI DSS"

End Sub

'==============================================================================
' ReinitialiserColonneResultat  (inchange)
'==============================================================================
Public Sub ReinitialiserColonneResultat()

    Dim wsPlan As Worksheet
    On Error Resume Next
    Set wsPlan = ThisWorkbook.Worksheets(SHEET_PLAN)
    On Error GoTo 0
    If wsPlan Is Nothing Then
        MsgBox "Feuille introuvable : """ & SHEET_PLAN & """", vbCritical: Exit Sub
    End If

    Dim lastRow As Long
    lastRow = DernierelignePleine(wsPlan, PLAN_ID_COL)
    If lastRow < PLAN_START_ROW Then Exit Sub

    wsPlan.Range(wsPlan.Cells(PLAN_START_ROW, PLAN_RESULT_COL), _
                 wsPlan.Cells(lastRow, PLAN_RESULT_COL)).ClearContents
    MsgBox "Colonne de resultat reinitialise.", vbInformation

End Sub

'==============================================================================
' ParseIdentifiant
' Analyse un identifiant PCI DSS brut et retourne (sBase, iRang).
'
' Exemples :
'   "3.5.1.a"  -> sBase="3.5.1"  iRang=1
'   "3.5.1.b"  -> sBase="3.5.1"  iRang=2
'   "3.5.1"    -> sBase="3.5.1"  iRang=0
'   "10.2.1.a" -> sBase="10.2.1" iRang=1
'   ""         -> sBase=""       iRang=0
'
' Logique :
'   Le suffixe lettre est reconnu quand le dernier segment apres le dernier "."
'   est une lettre unique (a-z ou A-Z).
'==============================================================================
Private Function ParseIdentifiant(raw As String) As TPciId

    Dim result As TPciId
    result.sBase = ""
    result.iRang = 0

    ' Nettoyage de base
    Dim s As String
    s = Trim(raw)
    s = Replace(s, Chr(160), "")
    s = Replace(s, Chr(9), "")
    s = Trim(s)

    If Len(s) = 0 Then
        ParseIdentifiant = result: Exit Function
    End If

    ' Doit commencer par un chiffre
    If Not (Left(s, 1) >= "0" And Left(s, 1) <= "9") Then
        ParseIdentifiant = result: Exit Function
    End If

    ' Doit contenir au moins un point
    If InStr(s, ".") = 0 Then
        ParseIdentifiant = result: Exit Function
    End If

    ' Dernier segment apres le dernier "."
    Dim iLastDot    As Integer
    Dim sLastSeg    As String
    iLastDot = 0
    Dim k As Integer
    For k = Len(s) To 1 Step -1
        If Mid(s, k, 1) = "." Then
            iLastDot = k
            Exit For
        End If
    Next k

    sLastSeg = Mid(s, iLastDot + 1)   ' segment apres le dernier point

    ' Si le dernier segment est une lettre unique -> suffixe lettre
    If Len(sLastSeg) = 1 Then
        Dim c As String
        c = LCase(sLastSeg)
        If c >= "a" And c <= "z" Then
            result.sBase = Left(s, iLastDot - 1)   ' tout sauf ".X"
            result.iRang = Asc(c) - Asc("a") + 1   ' a=1, b=2, c=3...
            ParseIdentifiant = result
            Exit Function
        End If
    End If

    ' Sinon : identifiant sans lettre suffixe
    result.sBase = s
    result.iRang = 0
    ParseIdentifiant = result

End Function

'==============================================================================
' DernierelignePleine  (inchange)
'==============================================================================
Private Function DernierelignePleine(ws As Worksheet, col As Long) As Long
    DernierelignePleine = ws.Cells(ws.Rows.Count, col).End(xlUp).Row
End Function