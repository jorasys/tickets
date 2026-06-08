' Variable publique pour stocker le cookie de session (si nécessaire pour d'autres macros)
Public RTSessionCookie As String

Sub CheckAndUpdateTickets()
    Dim wsTickets As Worksheet
    Dim wsCurrent As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim ticketID As String
    Dim status As String
    Dim found As Boolean
    Dim currentLastRow As Long
    Dim j As Long

    ' Référencer la feuille TICKETS
    On Error Resume Next
    Set wsTickets = ThisWorkbook.Worksheets("TICKETS")
    If wsTickets Is Nothing Then
        MsgBox "Feuille 'TICKETS' non trouvée."
        Exit Sub
    End If

    ' Référencer la feuille courante (active)
    Set wsCurrent = ThisWorkbook.ActiveSheet
    If wsCurrent Is Nothing Then
        MsgBox "Aucune feuille active."
        Exit Sub
    End If
    On Error GoTo 0

    ' Trouver la dernière ligne dans TICKETS
    lastRow = wsTickets.Cells(wsTickets.Rows.Count, 13).End(xlUp).Row  ' Colonne M
    Debug.Print "Dernière ligne TICKETS: " & lastRow

    ' Trouver la dernière ligne dans la feuille courante (colonne B)
    currentLastRow = wsCurrent.Cells(wsCurrent.Rows.Count, 2).End(xlUp).Row  ' Colonne B
    Debug.Print "Dernière ligne feuille courante: " & currentLastRow

    ' Boucler sur chaque ligne de TICKETS
    For i = 2 To lastRow  ' Supposant ligne 1 en-tête
        If Trim(wsTickets.Cells(i, 2).Value) = "" Then  ' Colonne B vide
            ticketID = Trim(wsTickets.Cells(i, 13).Value)  ' Colonne M
            status = Trim(wsTickets.Cells(i, 3).Value)    ' Colonne C
            If ticketID <> "" Then
                Debug.Print "Vérification ticket: " & ticketID & " (ligne " & i & ")"

                ' Vérifier si ticketID existe dans colonne B de la feuille courante
                found = False
                For j = 2 To currentLastRow
                    If Trim(wsCurrent.Cells(j, 2).Value) = ticketID Then
                        found = True
                        Exit For
                    End If
                Next j

                If found And LCase(status) = "resolved" Then
                    wsTickets.Cells(i, 2).Value = "à réouvrir (macro)"
                    Debug.Print "Marqué à réouvrir"
                ElseIf Not found And (LCase(status) = "new" Or LCase(status) = "open" Or LCase(status) = "stalled") Then
                    wsTickets.Cells(i, 2).Value = "à fermer (macro)"
                    Debug.Print "Marqué à fermer"
                Else
                    Debug.Print "Aucune action pour ce ticket"
                End If
            End If
        End If
    Next i

    MsgBox "Vérification terminée."
    Debug.Print "Vérification terminée."
End Sub