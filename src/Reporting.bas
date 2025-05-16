Attribute VB_Name = "Reporting"
Sub reportingFonds()
UserForm1.Show
End Sub
Sub GenererRapportFond(f As String)
    
'%%%% On Ètablit la connexion
Dim cnn As ADODB.Connection
Dim MaTable As ADODB.Recordset
Dim MaTable2 As ADODB.Recordset
Const ACCDB_Fournisseur = "Microsoft.ACE.OLEDB.12.0" ' Fournisseur de données.
Dim path As String
Dim Fichier_Source As String
path = ThisWorkbook.path & "\" & "basededonnees" & ".accdb"
Fichier_Source = path 'Fichier accdb ou mdb.

' Initialisation d'un objet connexion.
Set cnn = New ADODB.Connection
' Declaration du curseur control qui enregistre la navigation, qui contrôle _
la visibilité des changements dans la db et qui controle son update
cnn.CursorLocation = adUseServer
' Initialisation de la connexion à la base avec declaration du fournisseur :
cnn.Open "Provider= " & ACCDB_Fournisseur & ";" _
& "Data Source=" & Fichier_Source _
& ";", "", "", adAsyncConnect
While (cnn.State = adStateConnecting): DoEvents: Wend ' Attente de la connexion.


'%%%% Faire la jointure entre tables rendements mensuels et poids des actifs _
pour obtenir les rendements mensuels de chaque fond
Dim requete As String
requete = "SELECT rendements_mensuels.*, Parts_actifs.code_actif, Parts_actifs.Parts_" & LCase(f) & _
         " INTO rdts_actifs " & _
         "FROM rendements_mensuels " & _
         "INNER JOIN Parts_actifs ON rendements_mensuels.actifs = Parts_actifs.actif;"
' Exécution de la requête SQL
cnn.Execute requete

Dim i As Long, j As Long

' Dans une nouvelle feuille excel qu'on crée
Dim ws As Worksheet
Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
ws.Name = "Rendements " & f


'%%%% CALCUL DES RENDEMENTS MENSUELS DES FONDS


' On récupère la date de debut du fond
requete = "SELECT [Date] FROM pilotage_fonds WHERE Fonds = '" & f & "';"
Set MaTable = cnn.Execute(requete)

Dim datedebut As Date
datedebut = MaTable.Fields("Date")
' Si le fond ne commence pas au debut du mois, on calculera les rdts mensuels à partir du mois suivant
If Day(datedebut) <> 1 Then datedebut = DateAdd("m", 1, datedebut)

MaTable.Close
Set MaTable = Nothing ' On ferme la table et la connexion et on libère de la mémoire

Dim moisactuel As Date
moisactuel = "01/03/2024"
Dim nbmois As Integer ' Nb de mois depuis la date de début du fond
nbmois = DateDiff("m", datedebut, moisactuel)


Dim col As Long  'nb de colonne de dates de ma table access

requete = "SELECT * FROM rdts_actifs;"
Set MaTable = cnn.Execute(requete)
For col = (MaTable.Fields.count - 2) - nbmois To MaTable.Fields.count - 2 - 1 ' Boucle pour parcourir toutes les entitÈs mois de la table access _
on enlève la dernière colonne (poids de l'actif dans le fond) et l'avant dernière (code actif) et -1 pour obtenir l'indice correspondant
    
    requete = "SELECT SUM([" & MaTable.Fields(col).Name & "] * [Parts_" & LCase(f) & "]) AS rendement_fond FROM rdts_actifs;"
    Set MaTable2 = cnn.Execute(requete)

    ws.Cells(col + 10, 2).Value = MaTable2.Fields("rendement_fond").Value

    MaTable2.Close
    Set MaTable2 = Nothing

Next col

MaTable.Close
Set MaTable = Nothing

' On supprime la table jointure sur access
cnn.Execute "DROP TABLE rdts_actifs;"


' On inscrit les dates allant de février 2019 à fevrier 2024 dans la colonne 1
datedebut = DateSerial(2019, 2, 28) ' Date de départ en février 2019
moisactuel = datedebut ' Commencer par la date de début

i = 11 ' Commencer à la ligne 11
Do While moisactuel <= DateSerial(2024, 2, 29)
    ws.Cells(i, 1).Value = Format(moisactuel, "mm/yyyy") ' Écriture de la date au format mois/année
    moisactuel = DateAdd("m", 1, moisactuel) ' Passage au mois suivant
    i = i + 1 ' Passage à la ligne suivante dans la feuille de calcul
Loop

' On aura besoin du fichier Indice pour les rdts des marchés fr et américains
Dim wb As Workbook
Dim wsI As Worksheet
Dim chemin As String
chemin = ThisWorkbook.path & "\Indices.xlsm"
Set wb = Workbooks.Open(chemin)
Set wsI = wb.Worksheets("rdts_mensuels")
If f = "Alpha" Or f = "Gamma" Or f = "Omega" Then ' Ajouter graphique NYSE
        ' Copie colle la colonne 3 (rdts de NYSE) à la colonne 3
        wsI.Range(wsI.Cells(2, 3), wsI.Cells(wsI.Cells(wsI.Rows.count, 1).End(xlUp).row, 3)).Copy Destination:=ws.Cells(11, 3)
End If
If f = "Omicron" Or f = "Theta" Then ' Ajouter graphique FR CAC40
        ' Copie colle la colonne 5 (rdts du FR CAC40) à la colonne 3
        wsI.Range(wsI.Cells(2, 5), wsI.Cells(wsI.Cells(wsI.Rows.count, 1).End(xlUp).row, 5)).Copy Destination:=ws.Cells(11, 3)
End If
If f = "Omega" Or f = "Theta" Then ' Ajouter graphique Nasdaq
        ' Copie colle colonne 4 (rdts du Nasdaq) à la colonne 4
        wsI.Range(wsI.Cells(2, 4), wsI.Cells(wsI.Cells(wsI.Rows.count, 1).End(xlUp).row, 4)).Copy Destination:=ws.Cells(11, 4)
End If
ws.Range("C10:D" & ws.Cells(ws.Rows.count, 1).End(xlUp).row).Font.Color = RGB(255, 255, 255)   ' Blanc pour cacher les rdts mensuels des marchés
wb.Close SaveChanges:=False ' On ferme le fichier des indices

Do While IsEmpty(ws.Cells(11, 2))
    ws.Rows(11).Delete
Loop

'%%%% Création du graphique pour le fond
Dim chartObj As ChartObject
Dim chartRange As Range
Dim chart As chart

' On aura entre 2 et 3 courbes (series) sur le graphique : courbe du fond, et les courbes de NYSE, CAC40 et/ou Nasdaq
Dim series As series
Dim seriesIndex As Integer
seriesIndex = 1

' On positionne le graphique
Set chartObj = ws.ChartObjects.Add(Left:=ws.Cells(12, 3).Left + 15, _
                                   Top:=ws.Cells(12, 3).Top, _
                                   Width:=380, _
                                   Height:=400)
' Graphique à partir des dates (col 1), des rdts mensuels (col 2) et des rdts des indices (col 3 et 4)
Set chartRange = ws.Range(ws.Cells(11, 1), ws.Cells(ws.Cells(ws.Rows.count, 1).End(xlUp).row, 4))

Set chart = chartObj.chart
' Ajout des données au graphique
With chart
    .SetSourceData Source:=chartRange
    .ChartType = xlLine
    .HasTitle = True
    .ChartTitle.Text = "Evolution du rendement mensuel de " & f

    For Each series In .SeriesCollection
        series.Format.Line.Weight = 1.5
        If seriesIndex = 1 Then series.Name = f ' titre de la premiere courbe : fond
        If seriesIndex = 2 Then
            If f = "Alpha" Or f = "Gamma" Or f = "Omega" Then
                series.Name = "NYSE" ' titre de la 2ème courbe : NYSE
            ElseIf f = "Omicron" Or f = "Theta" Then
                series.Name = "CAC40" ' titre de la 2ème courbe : CAC40
            End If
        End If
        If seriesIndex = 3 Then
            If f = "Omega" Or f = "Theta" Then
                series.Name = "NASDAQ" ' titre de la 3ème courbe : Nasdaq
            Else
                series.Delete ' pas de troisième courbe
            End If
        End If
        seriesIndex = seriesIndex + 1
    Next series
End With
    
'%%%% Maintenant on va chercher les infos de la table pilotage_fonds
requete = "SELECT * FROM pilotage_fonds WHERE Fonds='" & f & "';"
Set MaTable = cnn.Execute(requete)

For j = 0 To MaTable.Fields.count - 3
    ws.Cells(j + 1, 1).Value = MaTable.Fields(j).Name
    ws.Cells(j + 1, 2).Value = MaTable.Fields(j).Value
Next j

' Je ferme ma table et je libère de la mémoire
MaTable.Close
Set MaTable = Nothing


'%%%% REMPLISSAGE CELLULES

With ws
    ' On ajoute les titres
    .Cells(10, 1).Value = "Date"
    .Cells(10, 2).Value = "Rendements mensuels " & f
    
    .Cells(6, 1).Value = "Rendement moyen"
    .Cells(7, 1).Value = "VolatilitÈ"
    .Cells(8, 1).Value = "Rendement moyen"
    .Cells(9, 1).Value = "5 derniers mois"
    
    .Cells(8, 4).Value = "Investissement moyen"
    .Cells(8, 5).Value = "Ecart-type"
    .Cells(8, 6).Value = "Min"
    .Cells(8, 7).Value = "Max"
    
    .Cells(10, 6).Value = "Nombre d'investisseurs"

    ' On calcule les rdt moyens et la volat
    .Cells(6, 2).Value = Application.WorksheetFunction.Average(ws.Range("B11:B" & ws.Cells(ws.Rows.count, 2).End(xlUp).row))
    .Cells(7, 2).Value = Application.WorksheetFunction.StDev(ws.Range("B11:B" & ws.Cells(ws.Rows.count, 2).End(xlUp).row))
    .Cells(8, 2).Value = Application.WorksheetFunction.Average(ws.Range("B" & ws.Cells(ws.Rows.count, 2).End(xlUp).row - 4 & ":B" & ws.Cells(ws.Rows.count, 2).End(xlUp).row))
End With




' Requête pour obtenir toutes les valeurs calculées
requete = "SELECT AVG(Somme_" & LCase(f) & ") AS Moyenne, " & _
          "STDEV(Somme_" & LCase(f) & ") AS EcartType, " & _
          "MIN(Somme_" & LCase(f) & ") AS Minimum, " & _
          "MAX(Somme_" & LCase(f) & ") AS Maximum, " & _
          "COUNT(*) AS NombreInvestisseurs " & _
          "FROM pilotage_investisseurs " & _
          "WHERE Somme_" & LCase(f) & " <> 0;"

' Exécution de la requête SQL
Set MaTable = cnn.Execute(requete)

' Récupération des valeurs et écriture dans les cellules correspondantes
ws.Cells(9, 4).Value = MaTable.Fields("Moyenne").Value
ws.Cells(9, 5).Value = MaTable.Fields("EcartType").Value
ws.Cells(9, 6).Value = MaTable.Fields("Minimum").Value
ws.Cells(9, 7).Value = MaTable.Fields("Maximum").Value
ws.Cells(10, 7).Value = MaTable.Fields("NombreInvestisseurs").Value

MaTable.Close ' On a fini avec notre table
Set MaTable = Nothing


'%%%% Ajout du logo dans la feuille de calcul
Dim img As Shape
Dim imgPath As String
imgPath = ThisWorkbook.path & "\" & "logo" & ".jpg" ' Chemin de l'image
Set img = ws.Shapes.AddPicture(imgPath, msoFalse, msoCTrue, _
ws.Cells(1, 7).Left + 10, ws.Cells(1, 7).Top, -0.5, -0.5)
' Redimensionner l'image
img.LockAspectRatio = msoTrue ' Verrouiller les proportions de l'image
img.Width = 70
img.Height = 70



'%%%% FORMATAGE CELLULES
With ws
    .Range("B3").NumberFormat = "dd/mm/yyyy" ' Formater dates
    .Range("B4, D9:G9").NumberFormat = "#,##0.00Ä" ' Formater somme et stats en euros
    .Range("B6:B8").NumberFormat = "0.00%" ' Formater rdts et volat en %
    .Range("B11:B" & ws.Cells(ws.Rows.count, "B").End(xlUp).row).NumberFormat = "General" ' Rdts mensuels en nb standard
    
    .Range("A1:A7, D8:G8, A10:B10").Font.Bold = True ' Formater les titres en gras
    .Range("A8:B8, A9").Font.Underline = xlUnderlineStyleSingle ' Souligner le texte
    
    .Cells.Font.Size = 9 ' Rapetisser ecritures
    .Cells(6, 1).Font.Size = 7
    .Cells(8, 1).Font.Size = 7
    .Cells(9, 1).Font.Size = 7
    .Cells(10, 2).Font.Size = 8

    .Columns("B:G").ColumnWidth = 15 ' Ajuster la largeur des colonnes
    .Columns("A").ColumnWidth = 8
    .Cells.HorizontalAlignment = xlCenter ' Texte centré
    .Cells(6, 1).HorizontalAlignment = xlLeft
End With

' Plage info du fond en tableau (quadrillages fins)
With ws.Range("A1:B7").Borders
    .LineStyle = xlContinuous
    .Weight = xlHairline
End With

' Plage des stats en tableau (quadrillages gros en rouge)
With ws.Range("D8:G9").Borders
    .LineStyle = xlContinuous
    .Weight = xlThick
    .Color = RGB(255, 0, 0)
End With

' Plage rdts mensuels en tableau (quadrillages fins)
With ws.Range("A10:B" & ws.Cells(ws.Rows.count, 1).End(xlUp).row).Borders
    .LineStyle = xlContinuous
    .Weight = xlHairline
End With



' On ferme la connexion et on libère la mémoire
cnn.Close
Set cnn = Nothing

' Suppression des gridlines de toutes les cellules non utilis_es sur la feuille
ws.Application.ActiveWindow.DisplayGridlines = False
            

'%%%% Exportation la feuille en PDF
cheminPDF = ThisWorkbook.path & "\Reporting " & f & ".pdf"
' Formatage du PDF
With ws.PageSetup
    .PrintArea = ws.Range("A1:H" & ws.Cells(ws.Rows.count, 1).End(xlUp).row).Address ' Impression de la plage de cellules utilisée
    .Orientation = xlLandscape ' Orientation en paysage
    .FitToPagesWide = 1 ' Tout doit tenir sur une page
    .FitToPagesTall = False ' Ne pas ajuster la hauteur
End With
ws.ExportAsFixedFormat Type:=xlTypePDF, Filename:=cheminPDF, Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False

'%%%% Ecrire un mail
Dim adresse As String           'Chaine de caractres pour le maill
Dim OutApp As Object            'Ouverture de l'application Outlook
Dim OutMail As Object

' Récupération de l'adresse email
adresse = "lunabaladi@gmail.com" '"projetLGLM@outlook.fr"

' Affectation de l'application Outlook
Set OutApp = CreateObject("Outlook.Application")
Set OutMail = OutApp.CreateItem(0)
    
' Rendre visible le mail
With OutMail
    .To = adresse
    .CC = ""
    .BCC = ""
    .Subject = "Reporting Fond " & f
    .HTMLbody = "Veuillez trouver ci-joint ‡ ce mail le reporting dÈsirÈ du fond " & f & "."
End With
' Condition v_rifiant la pr_sence de la pice jointee
If cheminPDF <> "" Then
    'Ajoute de la pice jointee
    OutMail.Attachments.Add cheminPDF
    ' Sauvegarde l 'email avant l'envoi
    OutMail.Save
    ' Envoie l'email
    OutMail.Send
End If

Application.DisplayAlerts = False ' Désactive les alertes pour éviter la confirmation de suppression
ws.Delete ' Supprime la feuille ws
Application.DisplayAlerts = True ' Réactive les alertes

End Sub

Sub reportingTopClient()
Dim ws As Worksheet
Dim cheminPDF As String
Dim outlookApp As Object
Dim email As Object
Dim dbPath As String
Dim cnn As Object
Dim MaTable As Object
Dim requete As String
Dim iRow, totalAge As Integer
Dim i, j, k As Integer
Dim AVG As Double

' Chemin de la base de donn_es Access
dbPath = ThisWorkbook.path & "\basededonnees.accdb"

' Connexion _ la base de donn_es Access
Set cnn = CreateObject("ADODB.Connection")
cnn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & dbPath


' SQL pour rÈcupÈrer les 10 meilleurs clients, la somme qu'ils ont investie dans chaque fond, et leur age
requete = "SELECT TOP 10 " & _
      "Nom & ' ' & Prenom AS Client," & _
      "Somme_investie_totale, " & _
      "Somme_Alpha / Somme_investie_totale AS Part_Alpha, " & _
      "Somme_Omega / Somme_investie_totale AS Part_Omega, " & _
      "Somme_Omicron / Somme_investie_totale AS Part_Omicron, " & _
      "Somme_Gamma / Somme_investie_totale AS Part_Gamma, " & _
      "Somme_Theta / Somme_investie_totale AS Part_Theta, " & _
      "Year(NOW) - Year(Date_naissance) AS Age " & _
      "FROM pilotage_investisseurs " & _
      "ORDER BY Somme_investie_totale DESC;"

' Ex_cution de la requtee
Set MaTable = cnn.Execute(requete)

' Cr_ation d'une nouvelle feuille pour le rapport
Set ws = ThisWorkbook.Sheets.Add

iRow = 1
Do While Not MaTable.EOF ' Boucle tant qu'il reste des enregistrements _ traiter dans le jeu de r_sultats
        For i = 0 To MaTable.Fields.count - 2 ' Boucle pour parcourir tous les champs
            ws.Cells(iRow + 1, i + 1).Value = MaTable.Fields(i).Value ' On inscrit les valeurs ‡ partir de la ligne 2, colonne 1.
        Next i
        iRow = iRow + 1 ' On passe _ la ligne suivante
        totalAge = totalAge + MaTable.Fields("Age").Value
        MaTable.MoveNext ' DÈplace le pointeur du jeu de r_sultats _ l'enregistrement suivant
Loop
ws.Cells(20, 2).Value = totalAge / 10 ' On calcule l'age moyen

' Requete top 3 dÈpartements qui reviennent parmi les clients ayant les meilleures sommes investies
requete = "SELECT TOP 3 " & _
       "Nom_departement AS Departement " & _
       "FROM pilotage_investisseurs " & _
       "ORDER BY Somme_investie_totale DESC;"
Set MaTable = cnn.Execute(requete)
iRow = 21
Do While Not MaTable.EOF ' Boucle tant qu'il reste des enregistrements _ traiter dans le jeu de r_sultats
        For i = 0 To MaTable.Fields.count - 1 ' Boucle pour parcourir tous les champs
            ws.Cells(iRow, 2).Value = MaTable.Fields(i).Value ' On les inscrit ‡ partir de la ligne 21, colonne 2.
        Next i
        iRow = iRow + 1
        MaTable.MoveNext ' DÈplace le pointeur du jeu de r_sultats _ l'enregistrement suivant
Loop

Dim fonds As Variant
Dim fond As Variant
fonds = Array("Alpha", "Gamma", "Omega", "Omicron", "Theta")
i = 25 ' ligne o˘ noter le meilleur investisseur
j = 3 ' colonne o˘ noter part fond
k = 14 ' ligne o˘ noter le fond

For Each fond In fonds
    ' On selectionne le client qui a la somme la plus importante dans le fond
    requete = "SELECT " & _
        "Nom & ' ' & Prenom AS Meilleur_Client " & _
        "FROM pilotage_investisseurs " & _
        "WHERE Somme_" & fond & " = (SELECT MAX(Somme_" & fond & ") FROM pilotage_investisseurs);"
    Set MaTable = cnn.Execute(requete)
    With ws
        .Cells(i, 2).Value = MaTable.Fields("Meilleur_Client").Value ' On l'inscrit en colonne 2
        .Cells(i, 1).Value = "Meilleur client chez " & fond
        ' Remplissage autres cellules
        .Cells(1, j).Value = "Part " & fond
        .Cells(k, 1).Value = fond
    End With
i = i + 1
j = j + 1
k = k + 1
Next fond

' Edition d'un tableau reportant la rÈpartition moyenne des investissements en fonction du fond
For i = 1 To 5 ' Boucle sur le nombre de fonds
    AVG = 0
    For j = 1 To 10 ' Boucle sur le nombre de clients
        AVG = AVG + ws.Cells(j + 1, 2 + i).Value
    Next j
    AVG = AVG / 10 ' Moyenne arithm_tique
    ws.Cells(13 + i, 2).Value = AVG
Next i


With ws
    .Name = "Reporting"
    
    ' %%%% Remplissage autres cellules
    .Cells(1, 1).Value = "Classement des 10 meilleurs clients de Lama-Gestion"
    .Cells(20, 1).Value = "Age moyen des meilleurs clients"
    .Cells(21, 1).Value = "Top 3 dÈpartements des meilleurs clients"
    .Cells(13, 1).Value = "RÈpartition moyenne des investissements des meilleurs clients"
    
    ' %%%% Formatage cellules
    .Rows("1:19").HorizontalAlignment = xlCenter ' On centre les Ècritures dans les cellules
    .Cells.Font.Size = 11 ' Police ‡ 11.
    .Range("A1:G1,A1:A11").Font.Bold = True ' En gras les titres

    With .Range(.Cells(1, 1), .Cells(1, 2))
        .Merge
        .Font.Color = RGB(255, 0, 0) ' Rouge
        .Font.Size = 10
    End With
    With .Range(Cells(13, 1), Cells(13, 2))
        .Merge
        .Font.Bold = True
        .Font.Size = 9
    End With
    .Rows("13").AutoFit
    With .Range(Cells(1, 1), Cells(11, 7)).Borders
        .LineStyle = xlHairline
        .Color = RGB(0, 0, 0)
        .Weight = xlHairline
    End With
    With ws.Range(Cells(13, 1), Cells(18, 2)).Borders
        .LineStyle = xlHairline
        .Color = RGB(0, 0, 0)
        .Weight = xlHairline
    End With
    .Range("A20:A29").Font.Underline = xlUnderlineStyleSingle
    .Range("A20:B29").Font.Size = 10
    .Range("A21").Font.Size = 8
    
    .Range(Cells(2, 3), Cells(11, 7)).NumberFormat = "0.00%"
    .Range(Cells(14, 2), Cells(18, 2)).NumberFormat = "0.00%"
    .Range(Cells(2, 2), Cells(11, 2)).NumberFormat = "#,##0Ä"
    
    .Columns("A").ColumnWidth = 23
    .Columns("B").ColumnWidth = 15
    .Columns("C:G").ColumnWidth = 10

End With


' Formatage des cellules avec age moyen, top 3 dÈpartements, et meillleur client par fonds
ws.Range(Cells(20, 1), Cells(21, 2)).Borders.LineStyle = xlHairline
ws.Range(Cells(20, 1), Cells(21, 2)).Borders.Color = vbBlack
ws.Range(Cells(20, 1), Cells(21, 2)).Borders.Weight = xlHairline

ws.Range(Cells(22, 2), Cells(23, 2)).Borders.LineStyle = xlHairline
ws.Range(Cells(22, 2), Cells(23, 2)).Borders.Color = vbBlack
ws.Range(Cells(22, 2), Cells(23, 2)).Borders.Weight = xlHairline

ws.Range(Cells(25, 1), Cells(29, 2)).Borders.LineStyle = xlHairline
ws.Range(Cells(25, 1), Cells(29, 2)).Borders.Color = vbBlack
ws.Range(Cells(25, 1), Cells(29, 2)).Borders.Weight = xlHairline


' %%%% InsÈrer le graphique
Dim chartObj As ChartObject
Dim chartDataRange As Range
Dim chart As chart

' DÈfinition de la plage de donnÈes pour le graphique
Set chartDataRange = ws.Range(Cells(14, 1), Cells(18, 2))

' Insertion du graphique en camembert
Set chartObj = ws.ChartObjects.Add(Left:=ws.Cells(13, 4).Left, _
                                    Width:=300, _
                                    Top:=ws.Cells(13, 4).Top, _
                                    Height:=190)
Set chart = chartObj.chart

' Configuration du graphique
With chart
    .SetSourceData Source:=chartDataRange
    .ChartType = xlPie

    ' Titre du graphique
    .HasTitle = True
    .ChartTitle.Text = "RÈpartition moyenne des investissements"

    ' Style et couleur optionnels
    .ApplyLayout (6) ' Style prÈdÈfini de couleurs
End With


'%%%% Ajout du logo dans la feuille de calcul
Dim img As Shape
Dim imgPath As String
imgPath = ThisWorkbook.path & "\" & "logo" & ".jpg" ' Chemin de l'image
Set img = ws.Shapes.AddPicture(imgPath, msoFalse, msoCTrue, _
ws.Cells(1, 8).Left + 10, ws.Cells(1, 8).Top, -0.5, -0.5)
' Redimensionner l'image
img.LockAspectRatio = msoTrue ' Verrouiller les proportions de l'image
img.Width = 70
img.Height = 70


' Exportation de la feuille en PDF
cheminPDF = ThisWorkbook.path & "\RapportClients.pdf"
    
' Formater le PDF
    
With ws.PageSetup
        .PrintArea = "A1:I29" 'Impression de la colonne A ‡ I
        .Orientation = xlLandscape ' Orientation pdf en paysage
        .FitToPagesWide = 1 ' Tout doit tenir sur une page
        .FitToPagesTall = False ' Ne pas ajuster la hauteur
End With
    
ws.ExportAsFixedFormat Type:=xlTypePDF, Filename:=cheminPDF
    

' Pr_paration de l'envoi par mail
Set outlookApp = CreateObject("Outlook.Application")
Set email = outlookApp.CreateItem(0)

'''''Ecrire un mail

Dim adresse As String           'Chaine de caractres pour le maill
Dim OutApp As Object            'Ouverture de l'application Outlook
Dim OutMail As Object

' R_cup_ration de l'adresse email
adresse = "lunabaladi@gmail.com" '"projetLGLM@outlook.fr"

' Affectation de l'application Outlook
Set OutApp = CreateObject("Outlook.Application")
Set OutMail = OutApp.CreateItem(0)
    
' Rendre visible le mail
With OutMail
        .To = adresse
        .CC = ""
        .BCC = ""
        .Subject = "Reporting meilleurs clients "
        .HTMLbody = "Veuillez trouver ci-joint ‡ ce mail le reporting sur nos meilleurs clients."
End With
' Condition v_rifiant la pr_sence de la pice jointee
If cheminPDF <> "" Then
    'Ajoute de la pice jointee
    OutMail.Attachments.Add cheminPDF

    'Sauvegarde l 'email avant l'envoi
    OutMail.Save
    'Envoie l'email
    OutMail.Send
End If

' Nettoyage
Application.DisplayAlerts = False
ws.Delete
Application.DisplayAlerts = True

' Fermeture de la connexion
MaTable.Close
Set MaTable = Nothing
cnn.Close
Set cnn = Nothing

End Sub
Sub reportingBoutique()

Dim cnn As ADODB.Connection
Dim MaTable As ADODB.Recordset
Const ACCDB_Fournisseur = "Microsoft.ACE.OLEDB.12.0" ' Fournisseur de donn_es.
Dim path As String
Dim Fichier_Source As String
path = ThisWorkbook.path & "\" & "basededonnees" & ".accdb"
Fichier_Source = path 'Fichier accdb ou mdb.

' Initialisation d'un objet connexion.
Set cnn = New ADODB.Connection
' Déclaration du curseur control qui enregistre la navigation, qui controle _
la visibilit_ des changements dans la db et qui controle son update
cnn.CursorLocation = adUseServer
' Initialisation de la connexion ˆ la base avec declaration du fournisseur :
cnn.Open "Provider= " & ACCDB_Fournisseur & ";" _
& "Data Source=" & Fichier_Source _
& ";", "", "", adAsyncConnect
While (cnn.State = adStateConnecting): DoEvents: Wend ' Attente de la connexion.


'%%%% Partie fonds

Dim ws As Worksheet
Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
ws.Name = "Reporting Boutique"

' On va chercher les infos dans la table pilotage_fonds de Access
Dim requete As String
requete = "SELECT * FROM pilotage_fonds;"
Set MaTable = cnn.Execute(requete)

Dim i As Long
For i = 1 To MaTable.Fields.count ' On copie colle les en-tetes de colonne
    ws.Cells(1, i).Value = MaTable.Fields(i - 1).Name
Next i
ws.Cells(2, 1).CopyFromRecordset MaTable ' On copie colle les donn_es de la table

With ws ' Ajuster les formats des colonnes
    .Range("C2:C6").NumberFormat = "dd-mm-yyyy"
    .Range("D2:D8").NumberFormat = "#,##0.00 Ä"
End With

' On ferme la table et la connexion et on libère de la mémoire
MaTable.Close
Set MaTable = Nothing

' On calcule la somme totale gérée dans une nouvelle colonne
ws.Cells(8, 3).Value = "Somme gÈrÈe"
ws.Cells(8, 4).Value = ws.Cells(2, 4).Value + ws.Cells(3, 4).Value + _
ws.Cells(4, 4).Value + ws.Cells(5, 4).Value + ws.Cells(6, 4).Value

' Créer le diagramme poids du fond
Set cht = ws.ChartObjects.Add(Left:=ws.Cells(10, 2).Left, Width:=200, _
                                Top:=ws.Cells(10, 1).Top, Height:=150)
' Définir la plage de données pour le diagramme
Dim chartRange As Range
Set chartRange = ws.Range("A2:A6")
' Ajouter une nouvelle série au graphique et définir ses valeurs
With cht.chart.SeriesCollection.NewSeries
    .XValues = chartRange
    .Values = ws.Range("G2:G6")
End With
' Définir le type de graphique
cht.chart.ChartType = xlPie
' Afficher le graphique
cht.Visible = True


'%%%% Partie employÈs

' On va chercher les infos dans pilotage_fonds
Set MaTable = New ADODB.Recordset
MaTable.Open "SELECT * FROM pilotage_fonds;", cnn

Dim gerants As Object
Dim fonds As Object
Dim j As Long
Dim gerant As Variant
Dim fond As Variant

' Initialisation des dictionnaires pour stocker les donnÈes uniques
Set gerants = CreateObject("Scripting.Dictionary")
Set fonds = CreateObject("Scripting.Dictionary")
' Parcours des donn_es pour extraire les informations uniques
Do While Not MaTable.EOF
    gerants(MaTable.Fields("Gerant").Value) = 1
    fonds(MaTable.Fields("Gerant").Value) = fonds(MaTable.Fields("Gerant").Value) & "/" & MaTable.Fields("Fonds").Value
    MaTable.MoveNext
Loop

' Ecriture des informations dans la feuille ws
i = 2
For Each gerant In gerants.Keys
    ws.Cells(24, i).Value = gerant ' Ecriture des noms de gérants en ligne 24
    i = i + 1
Next gerant
j = 2
For Each fond In fonds.Keys ' Ecriture des fonds associÈs à chaque gérant en ligne 25
    ws.Cells(25, j).Value = Mid(fonds(fond), 2) ' Supprimer le premier "/"
    j = j + 1
Next fond
' Ajout de titres
ws.Cells(24, 1).Value = "GÈrants"
ws.Cells(25, 1).Value = "Fonds"

' On ferme la table et la connexion et on libre de la m_moiree
MaTable.Close
Set MaTable = Nothing



'%%%% Partie investisseurs
' On cherche les infos dans la table pilotage_investisseurs
Set MaTable = New ADODB.Recordset
requete = "SELECT Num_client, Nom, Prenom, Mail, Date_naissance, Adresse, nom_commune, " & _
         " nom_region, somme_investie_totale FROM pilotage_investisseurs;"
Set MaTable = cnn.Execute(requete)

With ws ' Copie des données dans Excel
    For i = 0 To MaTable.Fields.count - 1
        .Cells(30, i + 1).Value = MaTable.Fields(i).Name ' Enttes de colonnee
    Next i
    .Range("A31").CopyFromRecordset MaTable ' Donn_es

    ' Ajuster le format des colonnes
    
    .Range("E31:E83").NumberFormat = "dd-mm-yyyy"
    .Range("I31:I83").NumberFormat = "#,##0.00 Ä"
    .Range("F31:F83").NumberFormat = "@"

    ' Ajuster le nom des colonnes
    .Cells(30, 1).Value = "NumÈro cient"
    .Cells(30, 5).Value = "Date de naissance"
    .Cells(30, 6).Value = "Code postal"
    .Cells(30, 7).Value = "Ville"
    .Cells(30, 8).Value = "RÈgion"
    .Cells(30, 9).Value = "Somme investie totale"

    ' Ajuster colonnes
    .UsedRange.ColumnWidth = 15
    .Columns(1).ColumnWidth = 7
    .Columns(2).ColumnWidth = 12
    .Columns(3).ColumnWidth = 10
    .Columns(5).ColumnWidth = 9
    .Columns(6).ColumnWidth = 6
    .Columns(9).ColumnWidth = 13

    .UsedRange.Font.Size = 9
End With


Dim plages() As Variant
' DÈfinition des plages de cellules ‡ formater (quadrillages fins)
plages = Array("A1:G6", "A24:E25", "A30:I83")

For i = LBound(plages) To UBound(plages) ' Parcours de chaque plage de cellules
    With ws.Range(plages(i)).Borders
        .LineStyle = xlContinuous
        .Weight = xlHairline
        If i = 1 Then
            .Color = RGB(255, 0, 0) ' 1ère plage en rouge
        End If
    End With
Next i
ws.Range("C8:D8").Interior.Color = RGB(255, 255, 0)


' Ajouter logo du fond
Dim img As Shape
Dim imgPath As String
imgPath = ThisWorkbook.path & "\" & "logo" & ".jpg" ' Chemin de l'image
Set img = ws.Shapes.AddPicture(imgPath, msoFalse, msoCTrue, _
ws.Cells(1, 9).Left, ws.Cells(1, 9).Top, -0.5, -0.5)
' Redimensionner l'image
img.LockAspectRatio = msoTrue ' Verrouiller les proportions de l'image
img.Width = 70
img.Height = 70

' Formatage ajoutÈ sur la colonne des poids des fonds
ws.Range(Cells(2, 7), Cells(6, 7)).NumberFormat = "0.##"

' On ferme la table et la connexion et on libère de la mémoire
MaTable.Close
Set MaTable = Nothing

cnn.Close
Set cnn = Nothing


Dim cheminPDF As String
cheminPDF = ThisWorkbook.path & "\ReportingDashboard.pdf"
    
' Formatage du PDF
With ws.PageSetup
    .PrintArea = ws.Range("A1:I85").Address ' Impression de la plage de cellules utilisée
    .Orientation = xlLandscape ' Orientation en paysage
    .FitToPagesWide = 1 ' Tout doit tenir sur une page
    .FitToPagesTall = False ' Ne pas ajuster la hauteur
End With
    
ws.ExportAsFixedFormat Type:=xlTypePDF, Filename:=cheminPDF
    
' Pr_paration de l'envoi par mail
Set outlookApp = CreateObject("Outlook.Application")
Set email = outlookApp.CreateItem(0)

'''''Ecrire un mail
Dim adresse As String           'Chaine de caractres pour le maill
Dim OutApp As Object            'Ouverture de l'application Outlook
Dim OutMail As Object

' R_cup_ration de l'adresse email
adresse = "lunabaladi@gmail.com" '"projetLGLM@outlook.fr"

'Affectation de l'application Outlook
Set OutApp = CreateObject("Outlook.Application")
Set OutMail = OutApp.CreateItem(0)
    
'Rendre visible le mail
With OutMail
    .To = adresse
    .CC = ""
    .BCC = ""
    .Subject = "Reporting Dashboard "
    .HTMLbody = "Veuillez trouver ci-joint ‡ ce mail le reporting de la boutique : vous lirez d'abord les donnÈes rÈsumant les diffÈrents fonds, une liste de tous les clients et de tous les gÈrants."
End With
'Condition v_rifiant la pr_sence de la pice jointee
If cheminPDF <> "" Then
        'Ajoute de la pice jointee
        OutMail.Attachments.Add cheminPDF
        'Sauvegarde l 'email avant l'envoi
        OutMail.Save
        'Envoie l'email
        OutMail.Send
End If

Application.DisplayAlerts = False ' Désactive les alertes pour éviter la confirmation de suppression
ws.Delete ' Supprime la feuille ws
Application.DisplayAlerts = True ' Réactive les alertes

End Sub
Function CompterSelec(lst As MSForms.ListBox, count As Integer)

Dim i As Integer
Dim nombreSelections As Integer
nombreSelections = 0

For i = 0 To count - 1
    If lst.Selected(i) Then
        nombreSelections = nombreSelections + 1
    End If
Next i

CompterSelec = nombreSelections

End Function
Sub reportingClient()
UserFormRpClient.Show
End Sub
Sub GenererReportingClient()

Dim ws As Worksheet
Dim dbPath As String

Dim cnn As ADODB.Connection

Dim MaTable0 As ADODB.Recordset
Dim MaTable1 As ADODB.Recordset
Dim MaTable2 As ADODB.Recordset

Dim requete As String

Dim nbclient As Integer

Dim liste As MSForms.ListBox
Dim clientSelect As Integer

Dim i As Integer

Dim moisactuel As Date
Dim datedebut As Date

Dim nbmois As Integer ' Nb de mois depuis la date de dÈbut du fond

Dim indice_mois As Long  ' Nb de colonnes de dates de ma table access

Dim fonds As Variant
fonds = Array("alpha", "omega", "omicron", "theta", "gamma")
Dim f As Integer ' Renverra ‡ l'indice du fond (de 0 ‡ 4 pour les 5 fonds)
Dim rdtMoy() As Double ' On crÈe un tableau o˘ on stockera les rendements moyens des fonds
ReDim rdtMoy(0 To UBound(fonds)) ' On ajoutera le rendement du fond ‡ ce tableau, ‡ mesure qu'on parcourt les fonds

Dim cpt As Integer ' Compteur du nb de fois qu'on parcourt un fond

Dim img As Shape
Dim imgPath As String

Dim outlookApp As Object
Dim email As Object
Dim cheminPDF As String

Dim adresse As String       ' Chaine de caractres pour le maill
Dim OutApp As Object        ' Ouverture de l'application Outlook
Dim OutMail As Object


' Chemin de la base de donn_es Access
dbPath = ThisWorkbook.path & "\basededonnees.accdb"

' Connexion _ la base de donn_es Access
Set cnn = New ADODB.Connection
cnn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & dbPath
cnn.Open

'On compte le nombre de client total
If UserFormRpClient.ListBox1.ListCount = 0 Then
    Exit Sub
Else
    nbclient = UserFormRpClient.ListBox1.ListCount
End If


'On compte le nombre de client s_lectionn_
Set liste = UserFormRpClient.ListBox1
clientSelect = CompterSelec(liste, nbclient)

Dim selectedValue As String

i = 0

Do While clientSelect > 0 Or i > (nbclient - 1) ' Jusqu'‡ qu'il n'y ai plus de rp ‡ faire ou que tous les clients ont ÈtÈ parcourus

    If UserFormRpClient.ListBox1.Selected(i) Then
        ' Cr_ation d'une nouvelle feuille pour le rapport
        Set ws = ThisWorkbook.Sheets.Add
        ws.Name = "Reporting Client"
    
        selectedValue = UserFormRpClient.ListBox1.List(i)
        
        With ws ' Mise en forme de la feuille
            .Cells(1, 1).Value = "NumÈro client"
            .Cells(1, 2).Value = "PrÈnom"
            .Cells(1, 3).Value = "Nom"
            .Cells(1, 4).Value = "Mail"
            .Cells(1, 5).Value = "Ville"
            .Cells(1, 6).Value = "RÈgion"
            .Cells(1, 7).Value = "Somme totale"
            .Cells(4, 2).Value = "Fond"
            .Cells(4, 3).Value = "Somme"
            .Cells(4, 4).Value = "Rendement moyen"
            .Cells(11, 1).Value = "Rdts mensuels chez"
        End With
    
        ' On sÈlectionne les infos du client choisi
        requete = "SELECT Num_client, Prenom, Nom, Mail, nom_commune, nom_region, " & _
        "somme_investie_totale FROM pilotage_investisseurs WHERE Num_client = '" _
        & selectedValue & "';"
        ' ExÈcution de la requete SQL
        Set MaTable0 = New ADODB.Recordset
        MaTable0.Open requete, cnn, adLockReadOnly
        
        ' V_rification que des donn_es ont _t_ trouv_es
        If Not MaTable0.EOF Then
            ' Placement des donn_es dans la feuille Excel
            With ws
                .Cells(2, 1).Value = MaTable0!Num_client
                .Cells(2, 2).Value = MaTable0!prenom
                .Cells(2, 3).Value = MaTable0!Nom
                .Cells(2, 4).Value = MaTable0!Mail
                .Cells(2, 5).Value = MaTable0!nom_commune
                .Cells(2, 6).Value = MaTable0!nom_region
                .Cells(2, 7).Value = MaTable0!somme_investie_totale
                .Cells(2, 7).NumberFormat = "#,##0 Ä" ' Cellule somme totale en euro
            End With
        End If
    
        MaTable0.Close

        ' On inscrit les dates allant de février 2019 à fevrier 2024 dans la colonne 1
        datedebut = DateSerial(2019, 2, 28) ' Date de départ en février 2019
        moisactuel = datedebut ' Commencer par la date de début

        indice_mois = 12 ' Commencer à la ligne 12
        Do While moisactuel <= DateSerial(2024, 2, 29)
            ws.Cells(indice_mois, 1).Value = Format(moisactuel, "mm/yyyy") ' Écriture de la date au format mois/année
            moisactuel = DateAdd("m", 1, moisactuel) ' Passage au mois suivant
            ws.Cells(indice_mois, 1).Font.Color = RGB(128, 0, 0)
            indice_mois = indice_mois + 1 ' Passage à la ligne suivante dans la feuille de calcul
        Loop

        moisactuel = "01/03/2024"
        
        cpt = 0 ' On initialise le compteur ‡ 0.
        
        For f = LBound(fonds) To UBound(fonds) ' Parcourir chaque fond pour r_cup_rer et afficher les montants investis
            requete = "SELECT somme_" & fonds(f) & ", somme_investie_totale FROM pilotage_investisseurs WHERE Num_client = '" & selectedValue & "';"
    
            ' Ex_cution de la requte pour le fond actuell
            MaTable0.Open requete, cnn, adLockReadOnly
    
            ' V_rification que des donn_es ont _t_ trouv_es
            If Not MaTable0.EOF Then
                If MaTable0.Fields(0).Value > 0 Then
                    
                    With ws.Cells(f + 5, 2)
                        .Value = UCase(fonds(f)) ' Ecrire le nom du fond
                        .Borders.LineStyle = xlContinuous ' Ajouter bordures
                        .Borders.Weight = xlHairline
                    End With
                    With ws.Cells(f + 5, 3)
                        .Value = MaTable0.Fields(0).Value ' Ecrire le montant investi si non nul
                        .Borders.LineStyle = xlContinuous ' Ajouter bordure
                        .Borders.Weight = xlHairline
                        .NumberFormat = "#,##0 Ä" ' En euro le montant
                    End With
                
                
                    ' %%%% Faire la jointure entre tables rendements mensuels et poids des actifs _
                    pour obtenir les rendements mensuels de chaque fond

                    cnn.Execute "SELECT rendements_mensuels.*, Parts_actifs.code_actif, Parts_actifs.Parts_" & fonds(f) & _
                        " INTO rdts_actifs " & _
                        "FROM rendements_mensuels " & _
                        "INNER JOIN Parts_actifs ON rendements_mensuels.actifs = Parts_actifs.actif;"

                    ' On rÈcupËère la date de debut du fond
                    Set MaTable1 = cnn.Execute("SELECT [Date] FROM pilotage_fonds WHERE Fonds = '" & fonds(f) & "';")
                    datedebut = MaTable1.Fields("Date")
                    ' Si le fond ne commence pas au debut du mois, on calculera les rdts mensuels à partir du mois suivant
                    If Day(datedebut) <> 1 Then datedebut = DateAdd("m", 1, datedebut)
                    MaTable1.Close
                    Set MaTable1 = Nothing ' On ferme la table et la connexion et on libère de la mémoire
                    nbmois = DateDiff("m", datedebut, moisactuel) ' =Nb de mois depuis le dÈbut du fond

                    ' On va sÈlectionner les rendements mensuels et calculer les rdts de chaque fond
                    Set MaTable1 = cnn.Execute("SELECT * FROM rdts_actifs;")

                    For indice_mois = (MaTable1.Fields.count - 2) - nbmois To MaTable1.Fields.count - 2 - 1
                    ' Boucle pour parcourir toutes les entitÈs mois de la table access. On enlËve la derniËre et l'avant derniËre colonnes _
                    (poids de l'actif dans le fond et code actif)
    
                        Set MaTable2 = cnn.Execute("SELECT SUM([" & MaTable1.Fields(indice_mois).Name & "] * " & _
                        "[Parts_" & fonds(f) & "]) AS rendement_fond FROM rdts_actifs;") ' On fait la moyenne pondÈrÈe du rendement mensuel pour le fond

                
                        ws.Cells(indice_mois + 11, f + 2).Value = MaTable2.Fields("rendement_fond").Value ' On l'inscrit dans notre tableau _
                        au mois correspondant(‡ partir de la ligne 12), colonne du fond correspondant (‡ partir de la colonne 2)
                        ' On a donc inscrit les rdts mensuels des fonds. Apres on les remplacera par les rdts mensuels du client dans les fonds.
                        
                        MaTable2.Close
                        Set MaTable2 = Nothing

                    Next indice_mois
                
                    ' On calcule le rendement moyen du fond sur toute la pÈriode fev 19 -> fev 24 (lignes 12 ‡ 72)
                    rdtMoy(f) = Application.WorksheetFunction.Average(ws.Range(ws.Cells(12, f + 2), ws.Cells(72, f + 2)))
                
                
                    ' Formatage cellules de titres :
                    With ws.Cells(11, f + 2)
                        .Value = UCase(fonds(f)) ' On inscrit le nom du fond
                        .Borders.LineStyle = xlContinuous ' Bordures
                        .Borders.Weight = xlHairline
                        .Font.Bold = True ' Gras
                    End With
                    With ws.Cells(f + 5, 4)
                        ' Rendement moyen du client dans un fond = rdt moyen du fond * (montant du client investi dans le fond / montant total investi)
                        .Value = rdtMoy(f) * (MaTable0.Fields(0).Value / MaTable0.Fields(1).Value)
                        .Borders.LineStyle = xlContinuous ' Bordures
                        .Borders.Weight = xlHairline
                        .NumberFormat = "0.00%" ' En % le rendement
                    End With
                    With ws.Cells(indice_mois, f + 2)
                        For indice_mois = 12 To 72 ' Pour tous les mois
                        ' Rendement mensuel du client dans un fond = rdt mensuel du fond *(montant du client investi dans le fond / montant total investi)
                            .Value = .Value * (MaTable0.Fields(0).Value / MaTable0.Fields(1).Value)
                            If .Value = 0 Then .Value = "" ' On ne veut pas afficher des 0 lorsqu'il n'y a pas de rdt mensuel
                        Next indice_mois
                    End With
            
                    MaTable1.Close
                    Set MaTable1 = Nothing

                    ' On supprime la table jointure sur access
                    cnn.Execute "DROP TABLE rdts_actifs;"
                    
                    cpt = cpt + 1
        
                End If
            End If
        
            MaTable0.Close
        
        Next f


        ' %%%% Formatage cellules

        With ws.Range("A1:G1") ' Plage infos du client
            .Font.Bold = True ' Gras
            .Font.Color = RGB(128, 0, 0) ' Couleur bordeaux
        End With
        With ws.Range("A1:G2")
            .Borders.LineStyle = xlHairline ' Appliquer bordure
            .Borders.Weight = xlContinuous
        End With
        With ws.Range("B4:D4") ' Plage fond, somme, rendement moyen
            .Borders.LineStyle = xlContinuous ' Bordure
            .Borders.Weight = xlHairline
            .Font.Bold = True ' Gras
            .Font.Color = RGB(128, 0, 0) ' Couleur bordeaux
        End With
        
        ' Appliquer des bordures ‡ la plage de rendements mensuels
        With ws.Range(ws.Cells(11, 1), ws.Cells(72, cpt + 1)).Borders
            .LineStyle = xlContinuous
            .Weight = xlHairline
        End With
        
        
        ws.Cells.HorizontalAlignment = xlCenter ' Aligne texte au centre
        ws.Range("E2:F2").HorizontalAlignment = xlLeft ' Ville et rÈgion alignÈes ‡ gauche pour pouvoir les lire
        ws.Cells.Font.Size = 10 ' Police ‡ 10
        ws.Cells.ColumnWidth = 15 ' Largeur colonnes ‡ 15

        ' Ajouter logo du fond
        imgPath = ThisWorkbook.path & "\" & "logo" & ".jpg" ' Chemin de l'image
        Set img = ws.Shapes.AddPicture(imgPath, msoFalse, msoCTrue, _
        ws.Cells(4, 7).Left, ws.Cells(4, 7).Top, -0.5, -0.5)
        ' Redimensionner l'image
        img.LockAspectRatio = msoTrue ' Verrouiller les proportions de l'image
        img.Width = 70
        img.Height = 70


        ' %%%% Exportation de la feuille en PDF
        cheminPDF = ThisWorkbook.path & "\ReportingClient.pdf"
        ' Formatage du PDF
        With ws.PageSetup
           .PrintArea = "$A:$G" 'Impression de la colonne A ‡ G
            .Orientation = xlLandscape ' Orientation pdf en paysage
            .FitToPagesWide = 1 ' Tout doit tenir sur une page
            .FitToPagesTall = 1 ' Limiter ‡ trois pages en hauteur
        End With
        ws.ExportAsFixedFormat Type:=xlTypePDF, Filename:=cheminPDF
    
    
        ' %%%% Ecrire un mail
        ' Pr_paration de l'envoi par mail
        Set outlookApp = CreateObject("Outlook.Application")
        Set email = outlookApp.CreateItem(0)

        ' RÈcupÈration de l'adresse email
        adresse = "lunabaladi@gmail.com" '"projetLGLM@outlook.fr"

        ' Affectation de l'application Outlook
        Set OutApp = CreateObject("Outlook.Application")
        Set OutMail = OutApp.CreateItem(0)
    
        ' Rendre visible le mail
        With OutMail
            .To = adresse
            .CC = ""
            .BCC = ""
            .Subject = "Reporting Client " & Cells(2, 2).Value & " " & Cells(2, 3).Value
            .HTMLbody = "Veuillez trouver ci-joint ‡ ce mail le reporting sur le client numÈro " & selectedValue & "."
        End With
        ' Condition v_rifiant la pr_sence de la pice jointee
        If cheminPDF <> "" Then
            ' Ajout de la pice jointee
            OutMail.Attachments.Add cheminPDF
            ' Sauvegarde l 'email avant l'envoi
            OutMail.Save
            '   Envoie l'email
            OutMail.Send
        End If
        
        
        ' Nettoyage
        Application.DisplayAlerts = False
        ws.Delete
        Application.DisplayAlerts = True

        clientSelect = clientSelect - 1 'un reporting marqu_ comme fait
    End If

    i = i + 1 'passer au num_ro client suivant

Loop

 ' Fermeture de la connexion
cnn.Close
Set cnn = Nothing

End Sub


