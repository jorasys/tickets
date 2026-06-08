' Variable publique pour stocker le cookie de session
Public RTSessionCookie As String

Sub SetRTSessionCookie()
    Dim cookieValue As String

    ' Demander la valeur du cookie de session
    cookieValue = InputBox("Entrez la valeur du cookie RT_SID_w-ha.com.7114:", "Cookie de session RT")
    If cookieValue = "" Then Exit Sub

    ' Stocker le cookie complet
    RTSessionCookie = "RT_SID_w-ha.com.7114=" & cookieValue

    MsgBox "Cookie de session défini: " & RTSessionCookie

    ' Maintenant, récupérer le ticket
    GetTicket
End Sub

Sub GetTicket()
    Dim httpRequest As Object
    Dim url As String
    Dim proxy As String
    Dim response As String

    ' Vérifier si le cookie est défini
    If RTSessionCookie = "" Then
        MsgBox "Cookie de session non défini. Exécutez d'abord SetRTSessionCookie."
        Exit Sub
    End If

    ' URL du ticket
    url = "http://rt.w-ha.com/REST/1.0/ticket/139399"

    ' Proxy
    proxy = "http://127.0.0.1:8080"

    ' Créer l'objet MSXML2.ServerXMLHTTP
    Set httpRequest = CreateObject("MSXML2.ServerXMLHTTP.6.0")

    ' Configurer le proxy
    httpRequest.setProxy 2, proxy  ' 2 = HTTP proxy

    ' Ouvrir la connexion GET
    httpRequest.Open "GET", url, False

    ' Définir les en-têtes
    httpRequest.setRequestHeader "Host", "rt.w-ha.com"
    httpRequest.setRequestHeader "Accept-Encoding", "gzip, deflate, br"
    httpRequest.setRequestHeader "Connection", "keep-alive"
    httpRequest.setRequestHeader "Cookie", RTSessionCookie

    ' Envoyer la requête
    httpRequest.Send

    ' Récupérer la réponse
    If httpRequest.Status = 200 Then
        response = httpRequest.ResponseText
        MsgBox "Ticket récupéré: " & Left(response, 500)
    Else
        MsgBox "Erreur: " & httpRequest.Status & " - " & httpRequest.StatusText
    End If

    ' Nettoyer
    Set httpRequest = Nothing
End Sub

Sub UpdateAllTickets()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim ticketID As String
    Dim queueValue As String

    ' Vérifier si le cookie est défini
    If RTSessionCookie = "" Then
        MsgBox "Cookie de session non défini. Exécutez d'abord SetRTSessionCookie."
        Exit Sub
    End If

    ' Référencer la feuille TICKETS
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("TICKETS")
    If ws Is Nothing Then
        MsgBox "Feuille 'TICKETS' non trouvée."
        Exit Sub
    End If
    On Error GoTo 0

    ' Trouver la dernière ligne
    lastRow = ws.Cells(ws.Rows.Count, 13).End(xlUp).Row  ' Colonne M
    Debug.Print "Dernière ligne trouvée: " & lastRow

    ' Boucler sur chaque ligne
    For i = 2 To lastRow  ' Supposant que la ligne 1 est l'en-tête
        ticketID = Trim(ws.Cells(i, 13).Value)  ' Colonne M
        If ticketID <> "" Then
            Debug.Print "Traitement du ticket: " & ticketID & " (ligne " & i & ")"
            Dim ticketData As Variant
            ticketData = GetTicketData(ticketID)
            If Not IsEmpty(ticketData) Then
                ws.Cells(i, 10).Value = ticketData(0)  ' Colonne J: Queue
                ws.Cells(i, 3).Value = ticketData(1)   ' Colonne C: Status
                ws.Cells(i, 11).NumberFormat = "@"    ' Forcer format texte pour LastUpdated
                ws.Cells(i, 11).Value = ticketData(2)  ' Colonne K: LastUpdated
                Debug.Print "Queue: " & ticketData(0) & ", Status: " & ticketData(1) & ", LastUpdated: " & ticketData(2)
            Else
                ws.Cells(i, 10).Value = "Erreur"
                ws.Cells(i, 3).Value = "Erreur"
                ws.Cells(i, 11).Value = "Erreur"
                Debug.Print "Erreur pour le ticket: " & ticketID
            End If
        Else
            Debug.Print "Ticket ID vide à la ligne " & i
        End If
    Next i

    MsgBox "Mise à jour terminée."
    Debug.Print "Mise à jour terminée."
End Sub

Function GetTicketData(ticketID As String) As Variant
    Dim httpRequest As Object
    Dim url As String
    Dim proxy As String
    Dim response As String
    Dim lines() As String
    Dim line As Variant
    Dim queueValue As String
    Dim statusValue As String
    Dim lastUpdatedValue As String

    ' URL du ticket
    url = "http://rt.w-ha.com/REST/1.0/ticket/" & ticketID

    ' Proxy
    proxy = "http://127.0.0.1:8080"

    ' Créer l'objet MSXML2.ServerXMLHTTP
    Set httpRequest = CreateObject("MSXML2.ServerXMLHTTP.6.0")

    ' Configurer le proxy
    httpRequest.setProxy 2, proxy  ' 2 = HTTP proxy

    ' Ouvrir la connexion GET
    httpRequest.Open "GET", url, False

    ' Définir les en-têtes
    httpRequest.setRequestHeader "Host", "rt.w-ha.com"
    httpRequest.setRequestHeader "Accept-Encoding", "gzip, deflate, br"
    httpRequest.setRequestHeader "Connection", "keep-alive"
    httpRequest.setRequestHeader "Cookie", RTSessionCookie

    ' Envoyer la requête
    httpRequest.Send

    ' Récupérer la réponse
    If httpRequest.Status = 200 Then
        response = httpRequest.ResponseText
        ' Parser la réponse pour extraire Queue, Status et LastUpdated
        lines = Split(response, vbLf)
        For Each line In lines
            If Left(Trim(line), 6) = "Queue:" Then
                queueValue = Trim(Mid(line, 7))
            ElseIf Left(Trim(line), 7) = "Status:" Then
                statusValue = Trim(Mid(line, 8))
            ElseIf Left(Trim(line), 12) = "LastUpdated:" Then
                lastUpdatedValue = Trim(Mid(line, 13))
            End If
        Next line
        If queueValue <> "" And statusValue <> "" And lastUpdatedValue <> "" Then
            GetTicketData = Array(queueValue, statusValue, lastUpdatedValue)
        Else
            GetTicketData = Empty
        End If
    Else
        GetTicketData = Empty
    End If

    ' Nettoyer
    Set httpRequest = Nothing
End Function