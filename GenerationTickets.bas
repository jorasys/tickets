Public Sub GenerationTickets()
    Dim ws As Worksheet
    Dim lastRow As Long
    
    ' Initialisation
    Set ws = ActiveSheet
    lastRow = ws.Cells(ws.Rows.Count, 4).End(xlUp).Row
    ' Système
    Dim stagingSystEnvData As New EnvData
    stagingSystEnvData.Name = "PREPROD"
    stagingSystEnvData.Team = "Système"
    Dim prodSystEnvData As New EnvData
    prodSystEnvData.Name = "PROD"
    prodSystEnvData.Team = "Système"
    ' Réseau
    Dim stagingErsoEnvData As New EnvData
    stagingErsoEnvData.Name = "PREPROD"
    stagingErsoEnvData.Team = "Réseau"
    Dim prodErsoEnvData As New EnvData
    prodErsoEnvData.Name = "PROD"
    prodErsoEnvData.Team = "Réseau"
    ' Windows-vm
    Dim stagingWinvEnvData As New EnvData
    stagingWinvEnvData.Name = "PREPROD"
    stagingWinvEnvData.Team = "Windows-vm"
    Dim prodWinvEnvData As New EnvData
    prodWinvEnvData.Name = "PROD"
    prodWinvEnvData.Team = "Windows-vm"
    ' Other
    Dim stagingOzerEnvData As New EnvData
    stagingOzerEnvData.Name = "PREPROD"
    stagingOzerEnvData.Team = "Autre"
    Dim prodOzerEnvData As New EnvData
    prodOzerEnvData.Name = "PROD"
    prodOzerEnvData.Team = "Autre"

    ' Parcourir les lignes visibles
    For Each ligne In ws.Range("D2:D" & lastRow).Rows
        If Not ligne.EntireRow.Hidden Then
            Dim dataInstance As New Data
            dataInstance.Initialize ligne
            dataInstance.Format
            Select Case dataInstance.Team
                Case "Système"
                    If dataInstance.IsStaging Then
                        stagingSystEnvData.AddData dataInstance
                    Else
                        prodSystEnvData.AddData dataInstance
                    End If
                Case "Réseau"
                    If dataInstance.IsStaging Then
                        stagingErsoEnvData.AddData dataInstance
                    Else
                        prodErsoEnvData.AddData dataInstance
                    End If
                Case "Windows-vm"
                    If dataInstance.IsStaging Then
                        stagingWinvEnvData.AddData dataInstance
                    Else
                        prodWinvEnvData.AddData dataInstance
                    End If
                Case "Autre"
                    If dataInstance.IsStaging Then
                        stagingOzerEnvData.AddData dataInstance
                    Else
                        prodOzerEnvData.AddData dataInstance
                    End If
                Case Else
                    Debug.Print "L'équipe n'est pas définie - " & dataInstance.Team
            End Select
        End If
    Next
    If stagingSystEnvData.Count > 0 Then
        Dim stagingSystTicket As New Ticket
        stagingSystTicket.Initialize stagingSystEnvData
        stagingSystTicket.AjouterLigneTicket
        stagingSystTicket.Export
    End If
    If prodSystEnvData.Count > 0 Then
        Dim prodSystTicket As New Ticket
        prodSystTicket.Initialize prodSystEnvData
        prodSystTicket.AjouterLigneTicket
        prodSystTicket.Export
    End If
    If stagingErsoEnvData.Count > 0 Then
        Dim stagingErsoTicket As New Ticket
        stagingErsoTicket.Initialize stagingErsoEnvData
        stagingErsoTicket.AjouterLigneTicket
        stagingErsoTicket.Export
    End If
    If prodErsoEnvData.Count > 0 Then
        Dim prodErsoTicket As New Ticket
        prodErsoTicket.Initialize prodErsoEnvData
        prodErsoTicket.AjouterLigneTicket
        prodErsoTicket.Export
    End If
    If stagingOzerEnvData.Count > 0 Then
        Dim stagingOzerTicket As New Ticket
        stagingOzerTicket.Initialize stagingOzerEnvData
        stagingOzerTicket.AjouterLigneTicket
        stagingOzerTicket.Export
    End If
    If prodOzerEnvData.Count > 0 Then
        Dim prodOzerTicket As New Ticket
        prodOzerTicket.Initialize prodOzerEnvData
        prodOzerTicket.AjouterLigneTicket
        prodOzerTicket.Export
    End If
   
End Sub

