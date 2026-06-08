Attribute VB_Name = "CompletionTickets"
Sub CompleterTickets()
    Debug.Print "Starting CompleterTickets"
    Dim wsMain As Worksheet
    Dim wsTickets As Worksheet
    Dim lastRowMain As Long
    Dim lastRowTickets As Long
    Dim i As Long, j As Long
    Dim foundMatch As Boolean
    Dim ipArray() As String
    Dim qidArray() As String
    Dim k As Long, m As Long
    Dim currentIP As String
    Dim currentQID As String
    
    ' D�finir les feuilles de travail
    Set wsMain = ActiveSheet ' Feuille principale active
    On Error Resume Next
    Set wsTickets = Worksheets("TICKETS") ' Feuille TICKETS
    On Error GoTo 0
    If wsTickets Is Nothing Then
        MsgBox "Erreur: Feuille TICKETS non trouvée.", vbExclamation
        Exit Sub
    End If
    Debug.Print "Worksheets set: Main=" & wsMain.Name & ", Tickets=" & wsTickets.Name
    
    ' D�terminer la derni�re ligne de chaque feuille
    lastRowMain = wsMain.Cells(wsMain.Rows.Count, "D").End(xlUp).Row
    lastRowTickets = wsTickets.Cells(wsTickets.Rows.Count, "E").End(xlUp).Row
    
    Application.ScreenUpdating = False ' D�sactiver la mise � jour de l'�cran pour am�liorer les performances
    
    ' Parcourir chaque ligne de la feuille principale
    For i = 2 To lastRowMain ' Commencer � la ligne 2 en supposant que la ligne 1 est un en-t�te
        foundMatch = False
        
        ' V�rifier si la ligne est masqu�e
        If wsMain.Rows(i).Hidden Then
            GoTo NextMainRow
        End If
        
        ' V�rifier si la cellule B contient d�j� une valeur
        If Not IsEmpty(wsMain.Cells(i, "B").value) Then
            ' Si la cellule B contient d�j� une valeur, passer � la ligne suivante
            GoTo NextMainRow
        End If
        
        ' R�cup�rer l'adresse IP et le QID de la ligne actuelle
        currentIP = Replace(Trim(wsMain.Cells(i, "D").value), " ", "")
    Debug.Print "Processing row " & i & ": IP=" & currentIP & ", QID=" & currentQID
        currentQID = Replace(Trim(wsMain.Cells(i, "J").value), " ", "")
        
        ' V�rifier que les valeurs ne sont pas vides
        If currentIP = "" Or currentQID = "" Then
            GoTo NextMainRow
        End If
        
        ' Parcourir chaque ligne de la feuille TICKETS pour trouver une correspondance
        For j = 2 To lastRowTickets ' Commencer � la ligne 2 en supposant que la ligne 1 est un en-t�te
            ' R�cup�rer les donn�es de la feuille TICKETS
            Dim ticketIPs As String
            Dim ticketQIDs As String
            
    Debug.Print "Checking ticket row " & j & ": IPs=" & ticketIPs & ", QIDs=" & ticketQIDs
            ticketIPs = Replace(Trim(wsTickets.Cells(j, "E").value), " ", "")
            ticketQIDs = Replace(Trim(wsTickets.Cells(j, "D").value), " ", "")
            
            ' V�rifier que les valeurs ne sont pas vides
            If ticketIPs = "" Or ticketQIDs = "" Then
                GoTo NextTicketRow
            End If
            
            ' R�cup�rer et diviser les adresses IP de la feuille TICKETS
'            Dim tempTicketIPs As String
'            tempTicketIPs = Replace(ticketIPs, ":", ",")
            ipArray = Split(ticketIPs, ",")

            ' R�cup�rer et diviser les QID de la feuille TICKETS
            qidArray = Split(ticketQIDs, ",")

    Debug.Print "Arrays: IP count=" & UBound(ipArray)+1 & ", QID count=" & UBound(qidArray)+1
            If UBound(ipArray) < 0 Or UBound(qidArray) < 0 Then GoTo NextTicketRow
            
            ' V�rifier chaque combinaison d'IP et de QID
            For k = 0 To UBound(ipArray)
                Dim trimmedIP As String
                trimmedIP = Replace(Trim(ipArray(k)), " ", "")
'                Debug.Print trimmedIP
                For m = 0 To UBound(qidArray)
                    Dim trimmedQID As String
                    trimmedQID = Replace(Trim(qidArray(m)), " ", "")
    Debug.Print "Comparing IP: " & trimmedIP & " vs " & currentIP & ", QID: " & trimmedQID & " vs " & currentQID
                    
'                    Debug.Print "Comparing: IP=" & trimmedIP & " vs " & currentIP & ", QID=" & trimmedQID & " vs " & currentQID & ", " & (StrComp(trimmedQID, currentQID, 1) = 0)
'                    Debug.Print InStr("192.168.202.25", "192.168.202.25:5031/tcp")
                    If (InStr(trimmedIP, currentIP) > 0 Or InStr(currentIP, trimmedIP) > 0) And StrComp(trimmedQID, currentQID, 1) = 0 Then
                        Debug.Print "Comparing: IP=" & trimmedIP & " vs " & currentIP & ", QID=" & trimmedQID & " vs " & currentQID
                        ' Correspondance trouv�e, copier le num�ro de ticket dans la colonne B
    Debug.Print "Match found for row " & i & " with ticket " & wsTickets.Cells(j, "A").value
                        wsMain.Cells(i, "B").value = wsTickets.Cells(j, "A").value
                        foundMatch = True
                        Exit For ' Sortir de la boucle une fois la correspondance trouv�e
                    End If
                Next m
                
                If foundMatch Then Exit For
            Next k
            
            If foundMatch Then Exit For ' Sortir de la boucle principale si une correspondance est trouv�e
NextTicketRow:
        Next j
        
        ' Pas d'action si aucune correspondance n'est trouv�e
NextMainRow:
    Next i
    
    Application.ScreenUpdating = True ' R�activer la mise � jour de l'�cran
    
    Debug.Print "CompleterTickets finished"
    MsgBox "Remplissage des num�ros de ticket termin�!", vbInformation
End Sub
