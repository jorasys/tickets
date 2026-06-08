Sub TestHttpRequestViaProxy()
    Dim httpRequest As Object
    Dim url As String
    Dim proxy As String
    Dim response As String

    ' URL cible
    url = "https://orasys.fr"

    ' Proxy (remplacez par votre proxy si nécessaire)
    proxy = "http://127.0.0.1:8080"

    ' Créer l'objet MSXML2.ServerXMLHTTP
    Set httpRequest = CreateObject("MSXML2.ServerXMLHTTP.6.0")

    ' Configurer pour ignorer les erreurs SSL
    httpRequest.setOption 2, 13056  ' SXH_SERVER_CERT_IGNORE_ALL_SERVER_ERRORS

    ' Configurer le proxy
    httpRequest.setProxy 2, proxy  ' 2 = HTTP proxy

    ' Ouvrir la connexion
    httpRequest.Open "GET", url, False

    ' Envoyer la requête
    httpRequest.Send

    ' Récupérer la réponse
    If httpRequest.Status = 200 Then
        response = httpRequest.ResponseText
        MsgBox "Réponse réussie: " & Left(response, 500)  ' Afficher les premiers 500 caractères
    Else
        MsgBox "Erreur: " & httpRequest.Status & " - " & httpRequest.StatusText
    End If

    ' Nettoyer
    Set httpRequest = Nothing
End Sub