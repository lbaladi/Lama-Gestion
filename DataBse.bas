Attribute VB_Name = "DataBse"
Sub modifDB()

Dim cnn As ADODB.Connection
Dim jointure As ADODB.Recordset
Dim path As String
Dim Fichier_Source As String
Const ACCDB_Fournisseur = "Microsoft.ACE.OLEDB.12.0"
path = ThisWorkbook.path & "\" & "basededonnees" & ".accdb"
Fichier_Source = path 'Fichier accdb ou mdb.
' Initialisation d'un objet connexion.
Set cnn = New ADODB.Connection
'Declaration du curseur control qui enregistre la navigation, qui controlela visibilte des changements dans la db et qui controle son update
cnn.CursorLocation = adUseServer
' Initialisation de la connexion a la base avec declaration du fournisseur :
cnn.Open "Provider= " & ACCDB_Fournisseur & ";" _
& "Data Source=" & Fichier_Source _
& ";", "", "", adAsyncConnect
While (cnn.State = adStateConnecting): DoEvents: Wend ' Attente de la connexion.

' Définir la requête SQL: jointure de regions, departments et city_adm dans une table: info_regions
Dim requete As String

requete = "SELECT city_adm.code_insee, city_adm.nom_commune, " & _
             "departments.code_departement, departments.nom_departement, " & _
             "regions.code_region, regions.nom_region " & _
             "INTO info_regions " & _
             "FROM ((city_adm " & _
             "INNER JOIN departments ON city_adm.code_departement = departments.code_departement) " & _
             "INNER JOIN regions ON departments.code_region = regions.code_region);"
' Exécution de la requête SQL
cnn.Execute requete
' Suppression des anciennes tables.
cnn.Execute "DROP TABLE departments;"
cnn.Execute "DROP TABLE city_adm;"
cnn.Execute "DROP TABLE regions;"

' Définir la requête SQL: on combine les infos de tous les fonds en une table pilotage_fonds
requete = "SELECT Fonds, Gerant, Date, Taille, Marche, Devise " & _
         "INTO pilotage_fonds " & _
         "FROM (" & _
         "SELECT Fonds, Gerant, Date, Taille, Marche, Devise FROM alpha_fonds " & _
         "UNION ALL " & _
         "SELECT Fonds, Gerant, Date, Taille, Marche, Devise FROM gamma_fonds " & _
         "UNION ALL " & _
         "SELECT Fonds, Gerant, Date, Taille, Marche, Devise FROM omega_fonds " & _
         "UNION ALL " & _
         "SELECT Fonds, Gerant, Date, Taille, Marche, Devise FROM omicron_fonds " & _
         "UNION ALL " & _
         "SELECT Fonds, Gerant, Date, Taille, Marche, Devise FROM theta_fonds);"
' Exécution de la requête SQL
cnn.Execute requete
' Suppression des anciennes tables.
cnn.Execute "DROP TABLE alpha_fonds;"
cnn.Execute "DROP TABLE gamma_fonds;"
cnn.Execute "DROP TABLE omega_fonds;"
cnn.Execute "DROP TABLE omicron_fonds;"
cnn.Execute "DROP TABLE theta_fonds;"

' Ajouter une nouvelle colonne TotalTaille à la table pilotage_fonds
cnn.Execute "ALTER TABLE pilotage_fonds ADD COLUMN TotalTaille DOUBLE;"
' Calcul de la somme totale des tailles des fonds
requete = "SELECT SUM(Taille) AS TotalTaille FROM pilotage_fonds;"
TotalTaille = cnn.Execute(requete).Fields(0).Value
requete = "UPDATE pilotage_fonds SET TotalTaille = " & TotalTaille & ";"
cnn.Execute requete
' Ajout de la colonne poids_boutique=taille du fond/TotalTaille à la table pilotage_fonds
requete = "ALTER TABLE pilotage_fonds ADD COLUMN poids_boutique DOUBLE;"
cnn.Execute requete
requete = "UPDATE pilotage_fonds SET poids_boutique = Taille / TotalTaille;"
cnn.Execute requete
' On supprime la colonne TotalTaille
requete = "ALTER TABLE pilotage_fonds DROP COLUMN TotalTaille;"
cnn.Execute requete

' Définir la requête SQL: jointure de tous les actifs avec leur numéro d'actif et leur part dans une table: parts_actifs
requete = "SELECT actifs.actif, actifs.code_actif, " & _
         "alpha_actifs.Parts_alpha, gamma_actifs.Parts_gamma, " & _
         "omega_actifs.Parts_omega, omicron_actifs.Parts_omicron, " & _
         "theta_actifs.Parts_theta " & _
         "INTO Parts_actifs " & _
         "FROM (((((actifs " & _
         "LEFT JOIN alpha_actifs ON actifs.actif = alpha_actifs.Actifs) " & _
         "LEFT JOIN gamma_actifs ON actifs.actif = gamma_actifs.Actifs) " & _
         "LEFT JOIN omega_actifs ON actifs.actif = omega_actifs.Actifs) " & _
         "LEFT JOIN omicron_actifs ON actifs.actif = omicron_actifs.Actifs) " & _
         "LEFT JOIN theta_actifs ON actifs.actif = theta_actifs.Actifs);"
' Exécution de la requête SQL
cnn.Execute requete
' Suppression des anciennes tables.
cnn.Execute "DROP TABLE alpha_actifs;"
cnn.Execute "DROP TABLE gamma_actifs;"
cnn.Execute "DROP TABLE omega_actifs;"
cnn.Execute "DROP TABLE omicron_actifs;"
cnn.Execute "DROP TABLE theta_actifs;"
cnn.Execute "DROP TABLE actifs;"

' Définir la requête SQL: jointure des investisseurs de tous les fonds avec leurs sommes investies _
dans chaque fond dans une table: info_investisseurs
requete = "SELECT alpha_investisseurs.Num_client, alpha_investisseurs.Nom, alpha_investisseurs.Prenom, " & _
         "alpha_investisseurs.Mail, alpha_investisseurs.Date_naissance, alpha_investisseurs.Adresse, alpha_investisseurs.Somme AS Somme_alpha, " & _
         "gamma_investisseurs.Somme AS Somme_gamma, omega_investisseurs.Somme AS Somme_omega, " & _
         "omicron_investisseurs.Somme AS Somme_omicron, theta_investisseurs.Somme AS Somme_theta " & _
         "INTO info_investisseurs " & _
         "FROM ((((alpha_investisseurs " & _
         "LEFT JOIN gamma_investisseurs ON alpha_investisseurs.Nom = gamma_investisseurs.Nom) " & _
         "LEFT JOIN omega_investisseurs ON alpha_investisseurs.Nom = omega_investisseurs.Nom) " & _
         "LEFT JOIN omicron_investisseurs ON alpha_investisseurs.Nom = omicron_investisseurs.Nom) " & _
         "LEFT JOIN theta_investisseurs ON alpha_investisseurs.Nom = theta_investisseurs.Nom);"
' Exécution de la requête SQL
cnn.Execute requete
' Suppression des anciennes tables.
cnn.Execute "DROP TABLE alpha_investisseurs;"
cnn.Execute "DROP TABLE gamma_investisseurs;"
cnn.Execute "DROP TABLE omega_investisseurs;"
cnn.Execute "DROP TABLE omicron_investisseurs;"
cnn.Execute "DROP TABLE theta_investisseurs;"
' Ajout de la colonne somme_investie_totale à la table info_investisseurs
requete = "ALTER TABLE info_investisseurs ADD COLUMN somme_investie_totale DOUBLE;"
cnn.Execute requete
' Calcul de la somme investie totale pour chaque client
requete = "UPDATE info_investisseurs SET somme_investie_totale = Somme_alpha + Somme_gamma + Somme_omega + Somme_omicron + Somme_theta;"
cnn.Execute requete
' Définition de la requête SQL pour la jointure des investisseurs avec leurs coordonnées à partir de leur _
code postal dans une table: pilotage_investisseur
requete = "SELECT info_investisseurs.Num_client, info_investisseurs.Nom, info_investisseurs.Prenom, " & _
             "info_investisseurs.Mail, info_investisseurs.Date_naissance, info_investisseurs.Adresse, " & _
             "info_regions.nom_commune, info_regions.code_departement, info_regions.nom_departement, " & _
             "info_regions.code_region, info_regions.nom_region, info_investisseurs.Somme_alpha, " & _
             "info_investisseurs.Somme_gamma, info_investisseurs.Somme_omega, info_investisseurs.Somme_omicron, " & _
             "info_investisseurs.Somme_theta, info_investisseurs.somme_investie_totale " & _
             "INTO pilotage_investisseurs " & _
             "FROM info_investisseurs " & _
             "LEFT JOIN info_regions ON info_investisseurs.Adresse = info_regions.code_insee;"
' Exécution de la requête SQL
cnn.Execute requete
' Suppression des anciennes tables.
cnn.Execute "DROP TABLE info_investisseurs;"
cnn.Execute "DROP TABLE info_regions;"



' On ferme la connexion
cnn.Close
Set cnn = Nothing
End Sub

Sub macro_rappro()
ImportListings
Rapprochement
End Sub
Sub ImportListings()

'%%%% Dans un premier temps on doit importer le fichier Listing New et le déposer sur Excel
Dim fso As Object
Dim ts As Object
Dim txtPath As String
Dim textLine As String
Dim row As Long
Dim ForReading As Integer
Dim data As Variant
Dim col As Integer
Dim ws1, ws2, ws3, wsR  As Worksheet
    
' Chemin d'accs au fichier TXT
txtPath = ThisWorkbook.path & "\Nasdaq Listing NEW.rtf"
    
' Créer un objet FileSystemObject
Set fso = CreateObject("Scripting.FileSystemObject")
    
' Ouvrir le fichier TXT en mode lecture
ForReading = 1
Set ts = fso.OpenTextFile(txtPath, ForReading)
    
' Initialiser le numéro de la premire lignee
row = 1
    
' Créer une feuille suppl_mentaire ö notre workbook qui r_pertorie les titres du listing NEW
Set ws1 = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
ws1.Name = "Listing New"
    
' Lire le fichier ligne par ligne
Do While Not ts.AtEndOfStream
    textLine = ts.ReadLine
    If InStr(textLine, ",") > 0 Then
        data = Split(textLine, ",")
        For col = 0 To UBound(data)
            ws1.Cells(row, col + 1) = data(col)
        Next col
    Else
        ws1.Cells(row, 1) = textLine
    End If
    ' Passer à la ligne suivante
    row = row + 1
Loop

' On s'assure qu'il n' y a pas de lignes en plus qui se sont importées du fichier rtf
Do Until ws1.Cells(1, 1).Value = "Symbol"
    ws1.Rows(1).Delete
Loop

' Fermer le fichier
ts.Close
    
' Nettoyer
Set ts = Nothing
Set fso = Nothing

'%%%% Ensuite on récupère la liste old dans listing market pour la déposer dans une seconde feuillee
Dim wbM As Workbook
Dim wsN As Worksheet
Dim chemin As String
Dim observ As Integer

chemin = ThisWorkbook.path & "\Listing market.xlsx"
Set wbM = Workbooks.Open(chemin)
Set wsN = wbM.Worksheets("Nasdaq")

Set ws2 = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
ws2.Name = "Listing old"
' On recopie la colonne des titres dans wsN sur ws2 qui est une sheet crŽe temporairement sur le wb
observ = wsN.Cells(Rows.count, 1).End(xlUp).row - 1
ws2.Cells(2, 1).Resize(observ, 1).Value = wsN.Cells(2, 1).Resize(observ, 1).Value

wbM.Close SaveChanges:=False

End Sub
Sub Rapprochement()
Dim ws1 As Worksheet, ws2 As Worksheet, ws3 As Worksheet, wsR As Worksheet
Dim cell As Range
Dim valeursUniques As Collection
Dim valeur As Variant
Dim i As Long
Dim observ As Integer
' Initialiser les feuilles
Set ws1 = ThisWorkbook.Sheets("Listing New")
Set ws2 = ThisWorkbook.Sheets("Listing old")
Set ws3 = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
ws3.Name = "titres presents"
Set valeursUniques = New Collection
    
' Parcourir la colonne 1 de la feuille 2
For Each cell In ws2.Range("A1:A" & ws2.Cells(ws2.Rows.count, 1).End(xlUp).row)
    If cell.Value <> "" And IsError(cell.Value) = False Then
        On Error Resume Next
        valeursUniques.Add cell.Value, CStr(cell.Value)
        On Error GoTo 0
    End If
Next cell

' Parcourir la colonne 1 de la feuille 1
For Each cell In ws1.Range("A1:A" & ws1.Cells(ws1.Rows.count, 1).End(xlUp).row)
    If cell.Value <> "" And IsError(cell.Value) = False Then
        On Error Resume Next
        valeursUniques.Add cell.Value, CStr(cell.Value)
        On Error GoTo 0
    End If
Next cell
    
' Ecrire les valeurs uniques dans la feuille 3, colonne 1
i = 1
For Each valeur In valeursUniques
    ws3.Cells(i, 1).Value = valeur
    ' Vérifier la présence de la valeur dans ws1 et ws2
    Set celluleTrouvee1 = ws1.Columns(1).Find(valeur, LookIn:=xlValues, LookAt:=xlWhole)
    Set celluleTrouvee2 = ws2.Columns(1).Find(valeur, LookIn:=xlValues, LookAt:=xlWhole)
    ' Ecrire la localisation de la valeur
    If (Not celluleTrouvee1 Is Nothing) And (Not celluleTrouvee2 Is Nothing) Then
        ws3.Cells(i, 2).Value = "ws1 + ws2"
    ElseIf Not celluleTrouvee1 Is Nothing Then
        ws3.Cells(i, 2).Value = "ws1"
    ElseIf Not celluleTrouvee2 Is Nothing Then
        ws3.Cells(i, 2).Value = "ws2"
    End If
    i = i + 1
Next valeur

' Trier nos données
With ws3.Sort
    .SortFields.Clear
    .SortFields.Add Key:=ws3.Range("A1:A" & ws3.Cells(ws3.Rows.count, 1).End(xlUp).row), Order:=xlAscending
    .SetRange ws3.Range("A1:B" & ws3.Cells(ws3.Rows.count, 1).End(xlUp).row)
    .Header = xlNo
    .MatchCase = False
    .Apply
End With


'%%%%% Création d'une 4ème feuille wsR qui restera définitivement sur notre code

'Mise en page de la feuille
Set wsR = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
wsR.Name = "Rapprochement"
wsR.Cells(1, 1).Value = "Rapprochement NASDAQ listing et NEW listing"
wsR.Cells(2, 1).Value = "Données manquantes dans NEW"
wsR.Cells(2, 2).Value = "Données manquantes dans OLD"

observ = ws3.Cells(1, ws3.Columns.count).End(xlUp).Column - 1
' Dernières lignes remplies aux colonnes 1 et 2
Dim lastrow1, lastrow2 As Integer
lastrow1 = 2
lastrow2 = 2
For i = 1 To ws3.Cells(ws3.Rows.count, 1).End(xlUp).row
    If ws3.Cells(i, 2).Value = "ws1" Then
    ' On écrit en dessous de lastrow à la colonne 2 puis on actualise la valeur de cette ligne
        wsR.Cells(lastrow1 + 1, 2).Value = ws3.Cells(i, 1).Value
        lastrow1 = lastrow1 + 1
    ElseIf ws3.Cells(i, 2).Value = "ws2" Then
    ' On écrit en dessous de lastrow à la colonne 1 puis on actualise la valeur de cette ligne
        wsR.Cells(lastrow2 + 1, 1).Value = ws3.Cells(i, 1).Value
        lastrow2 = lastrow2 + 1
    End If
Next i

'Suppression des feuilles 3 feuilles créées
Application.DisplayAlerts = False
ws1.Delete
ws2.Delete
ws3.Delete
Application.DisplayAlerts = True

'On s'assure que la feuille de Rapprochement est placée en 2
wsR.Move After:=ThisWorkbook.Sheets(2)

End Sub
Sub create_database() ' sub appelée par le bouton CREATION DATABASE

ThisWorkbook.Worksheets(1).Cells.ClearContents

createDb_ADO
createTables_ADO
RemplirTables_regions
RemplirTables_fonds
RemplirTablesActifs
RemplirActif

MsgBox "Votre database est stockée ici : " & ThisWorkbook.path & "\" & "basededonnees" & ".accdb"

End Sub
Sub createDb_ADO() ' Creation Database

Dim sConnectStr As String
Dim adoCat As ADOX.Catalog
Dim sNameDb As String
Dim sPathDb As String
Dim sFullPath As String

sNameDb = "basededonnees" ' Nom de la db

' Affectation de la connexion a access et du chemin de la db
Set adoCat = New ADOX.Catalog
sFullPath = ThisWorkbook.path & "\" & sNameDb & ".accdb"
sConnectStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & sFullPath

' Création de la database
Set adoCat = New ADOX.Catalog
adoCat.Create sConnectStr
Set adoCat = Nothing

End Sub
Sub createTables_ADO() ' Création tables
Dim dbs As DataBase
' Ouvrir la base de données
Set dbs = OpenDatabase(ThisWorkbook.path & "\" & "basededonnees" & ".accdb")

'%%%% Création de la table REGIONS
dbs.Execute "CREATE TABLE regions (nom_region CHAR, code_region INTEGER);"
'%%%% Création de la table DEPARTEMENTS
dbs.Execute "CREATE TABLE departments (nom_departement CHAR, code_region INTEGER, code_departement CHAR);"
'%%%% Création de la table city
dbs.Execute "CREATE TABLE city_adm (nom_commune CHAR, code_insee CHAR, type_commune CHAR, code_departement CHAR);"

'%%%% Création des tables des différents fonds:
Dim i As Long, j As Long
Dim tablePrefixes As Variant
Dim tableSuffixes As Variant
Dim queries As Variant
' Tableaux pour stocker les préfixes et suffixes des tables
tablePrefixes = Array("alpha", "gamma", "omega", "omicron", "theta")
tableSuffixes = Array("_fonds", "_actifs", "_investisseurs")
' Requête SQL de création des tables avec les entités à chaque fois
queries = Array( _
    "CREATE TABLE %prefix%_fonds (Fonds TEXT, Gerant TEXT, [Date] DATETIME, Taille DOUBLE, Marche TEXT, Devise TEXT);", _
    "CREATE TABLE %prefix%_actifs (Actifs TEXT, Parts_%prefix% DOUBLE);", _
    "CREATE TABLE %prefix%_investisseurs (Num_client TEXT, Nom TEXT, Prenom TEXT, Date_naissance DATETIME, Somme DOUBLE);")
' Boucler à travers les préfixes et suffixes des tables pour créer toutes les tables
For i = LBound(tablePrefixes) To UBound(tablePrefixes)
    For j = LBound(tableSuffixes) To UBound(tableSuffixes)
        ' Remplacer les marqueurs de position dans la requête SQL par les préfixes et suffixes actuels
        Dim query As String
        query = Replace(queries(j), "%prefix%", tablePrefixes(i))
        ' Exécuter la requête SQL correspondante
        ' Cas particulier Theta_investisseur  n'a que 3 entités
        If tablePrefixes(i) = "theta" And tableSuffixes(j) = "_investisseurs" Then
            dbs.Execute "CREATE TABLE theta_investisseurs (Nom TEXT, Prenom TEXT, Somme DOUBLE);"
        ' Cas particulier Alpha_investisseurs a 7 entités
        ElseIf tablePrefixes(i) = "alpha" And tableSuffixes(j) = "_investisseurs" Then
            dbs.Execute "CREATE TABLE alpha_investisseurs (Num_client TEXT, Nom TEXT, Prenom TEXT, Mail TEXT, Date_naissance DATETIME, Adresse CHAR, Somme DOUBLE);"
        Else
            dbs.Execute query
        End If
    Next j
Next i

'%%%% Création des tables des rendements des actifs présents dans le fichier DATASET
Dim wb As Workbook
Dim ws As Worksheet
Dim chemin As String
chemin = ThisWorkbook.path & "\DATASET.xlsm"
Set wb = Workbooks.Open(chemin)
Set ws = wb.Worksheets("Univers actifs")
Dim col As Long
Dim cpt As Integer

' Création de 4 tables rendements_actifs. On divise le nb de colonnes en 4 car dépassement de capacité
col = WorksheetFunction.RoundUp(ws.Cells(1, ws.Columns.count).End(xlToLeft).Column / 4, 0)
For i = 1 To 4
    dbs.Execute "CREATE TABLE rendements_actifs" & i & " ([Date] DATETIME);"
    If i = 4 Then
        ' Dernière table : For j = 3*col + 2 to 4*col
        For j = cpt * col + 2 To (cpt + 1) * col
            dbs.Execute "ALTER TABLE rendements_actifs" & i & " ADD COLUMN [" & "Actif " & j - 1 & "] TEXT;"
        Next j
    Else
        ' Les tables se suivent : _
        Première table : For j = 2 to col + 1 _
        Deuxième table : For j = col + 2 To 2 * col + 1 _
        Troisième table : For j = 2 * col + 2 To 3 * col + 1
        For j = cpt * col + 2 To (cpt + 1) * col + 1
            dbs.Execute "ALTER TABLE rendements_actifs" & i & " ADD COLUMN [" & "Actif " & j - 1 & "] TEXT;"
        Next j
    End If
    cpt = cpt + 1
Next i


'%%%% Création table rdts mensuels
Set ws = wb.Worksheets("rdts_mensuels_t")
dbs.Execute "CREATE TABLE rendements_mensuels" & " (actifs CHAR);"

For i = 2 To 12
    dbs.Execute "ALTER TABLE rendements_mensuels" & " ADD COLUMN [" & i & "-19 " & "] TEXT;"
Next i

For j = 20 To 23
    For i = 1 To 12
        dbs.Execute "ALTER TABLE rendements_mensuels" & " ADD COLUMN [" & i & "-" & j & "] TEXT;"
    Next i
Next j

For i = 1 To 2
    dbs.Execute "ALTER TABLE rendements_mensuels" & " ADD COLUMN [" & i & "-24 " & "] TEXT;"
Next i

' On ferme le fichier DATASET
wb.Close SaveChanges:=False

'%%%% Création de la table qui énumère les actifs
dbs.Execute "CREATE TABLE actifs" & "(actif CHAR, code_actif CHAR);"


dbs.Close
End Sub
' Procédure pour remplir les tables regions, départements et city_adm
Sub RemplirTableRegions(ByVal tableName As String, ByVal colonnes As Variant)

Dim cnn As ADODB.Connection
Dim MaTable As ADODB.Recordset
Dim path As String
Dim Fichier_Source As String
Dim wb As Workbook
Dim ws As Worksheet
Const ACCDB_Fournisseur = "Microsoft.ACE.OLEDB.12.0"
path = ThisWorkbook.path & "\" & "basededonnees" & ".accdb"
Fichier_Source = path 'Fichier accdb ou mdb.
' Initialisation d'un objet connexion.
Set cnn = New ADODB.Connection
'Declaration du curseur control qui enregistre la navigation, qui controlela visibilte des changements dans la db et qui controle son update
cnn.CursorLocation = adUseServer
' Initialisation de la connexion a la base avec declaration du fournisseur :
cnn.Open "Provider= " & ACCDB_Fournisseur & ";" _
& "Data Source=" & Fichier_Source _
& ";", "", "", adAsyncConnect
While (cnn.State = adStateConnecting): DoEvents: Wend ' Attente de la connexion.

' Ouverture de la table en lecture/écriture.
Set MaTable = New ADODB.Recordset
MaTable.Open "SELECT * FROM " & tableName, cnn, adOpenKeyset, adLockPessimistic, adCmdText
' Ouverture du fichier Excel pour lecture.
Set wb = Workbooks.Open(ThisWorkbook.path & "\" & tableName & ".csv")
Set ws = wb.Worksheets(1)

Dim i As Long, j As Long
' On remplit les données de la table à partir des valeurs dans excel
For i = 2 To ws.UsedRange.Rows.count
    MaTable.AddNew
    For j = LBound(colonnes) To UBound(colonnes)
        MaTable.Fields(colonnes(j)).Value = ws.Cells(i, j + 1).Value
    Next j
Next i
' On ferme le fichier Excel
wb.Close SaveChanges:=False
' Enregistrement des modifications.
MaTable.Update
' Fermeture de la table et on libère de la mémoire
MaTable.Close
Set MaTable = Nothing
' Fermeture de la connexion et on libère de la mémoire
cnn.Close
Set cnn = Nothing

End Sub
Sub RemplirTables_regions() ' Macro qui exécute la procédure pour les trois tables
RemplirTableRegions "regions", Array("nom_region", "code_region")
RemplirTableRegions "departments", Array("nom_departement", "code_region", "code_departement")
RemplirTableRegions "city_adm", Array("nom_commune", "code_insee", "type_commune", "code_departement")
End Sub
Sub RemplirTables_fonds() ' Macro qui remplit les tables de tous les fonds

Dim cnn As ADODB.Connection ' Objet representant une connexion à la base.
Dim MaTable As ADODB.Recordset ' Objet representant la table de la base.
Dim path As String
Dim Fichier_Source As String

Dim wb As Workbook
Dim ws1 As Worksheet, ws2 As Worksheet

Dim i As Long
Dim tableName As Variant
Dim fondsSheetName As String, actifsSheetName As String, investisseursSheetName As String

Const ACCDB_Fournisseur = "Microsoft.ACE.OLEDB.12.0" ' Fournisseur de donnees.
path = ThisWorkbook.path & "\" & "basededonnees" & ".accdb"
Fichier_Source = path 'Fichier accdb ou mdb.
' Initialisation d'un objet connexion.
Set cnn = New ADODB.Connection
cnn.CursorLocation = adUseServer
' Initialisation de la connexion à la base avec déclaration du fournisseur :
cnn.Open "Provider= " & ACCDB_Fournisseur & ";" _
& "Data Source=" & Fichier_Source _
& ";", "", "", adAsyncConnect
While (cnn.State = adStateConnecting): DoEvents: Wend ' Attente de la connexion.

For Each tableName In Array("Alpha", "Gamma", "Omega", "Omicron", "Theta")
    ' Récupérer le nom de la feuille de calcul correspondante
    fondsSheetName = tableName & "_fonds"
    actifsSheetName = tableName & "_actifs"
    investisseursSheetName = tableName & "_investisseurs"
    ' Ouverture du classeur contenant les données
    Set wb = Workbooks.Open(ThisWorkbook.path & "\" & tableName & ".xlsm")
    ' Récupérer les feuilles de calcul
    Set ws1 = wb.Worksheets(1)
    Set ws2 = wb.Worksheets(2)
    
    '%%%% Remplir la table des fonds
    Set MaTable = New ADODB.Recordset
    MaTable.Open "SELECT * FROM " & fondsSheetName, cnn, adOpenKeyset, adLockPessimistic, adCmdText
    MaTable.AddNew
    ' Remplir les données de la première ligne de la première feuille de calcul
    For i = 1 To ws1.UsedRange.Columns.count
        MaTable.Fields(i - 1).Value = ws1.Cells(2, i)
    Next i
    MaTable.Update
    MaTable.Close
    Set MaTable = Nothing
    
    '%%%% Remplir la table des actifs
    Set MaTable = New ADODB.Recordset
    MaTable.Open "SELECT * FROM " & actifsSheetName, cnn, adOpenKeyset, adLockPessimistic, adCmdText
    ' Remplir les données à partir de la 5ème ligne de la feuille de calcul (les actifs et leurs parts)
    For i = 5 To ws1.UsedRange.Rows.count
        MaTable.AddNew
        MaTable!actifs = ws1.Cells(i, 1)
        MaTable.Fields("Parts_" & tableName).Value = ws1.Cells(i, 2)
    Next i
    MaTable.Update
    MaTable.Close
    Set MaTable = Nothing

    '%%%% Remplir la table des investisseurs
    Set MaTable = New ADODB.Recordset
    MaTable.Open "SELECT * FROM " & investisseursSheetName, cnn, adOpenKeyset, adLockPessimistic, adCmdText
    ' Remplir les données à partir de la deuxième ligne de la deuxième feuille de calcul (les infos des clients)
    For i = 2 To ws2.UsedRange.Rows.count
        MaTable.AddNew
        ' Cas particulier table theta
        If tableName = "Theta" Then
            MaTable!Nom = ws2.Cells(i, 1)
            MaTable!prenom = ws2.Cells(i, 2)
            MaTable!Somme = ws2.Cells(i, 3)
        ' Cas particulier table Alpha
        ElseIf tableName = "Alpha" Then
            MaTable!Num_client = ws2.Cells(i, 1)
            MaTable!Nom = ws2.Cells(i, 2)
            MaTable!prenom = ws2.Cells(i, 3)
            MaTable!Mail = ws2.Cells(i, 4)
            MaTable!Date_naissance = ws2.Cells(i, 5)
            MaTable!adresse = ws2.Cells(i, 6)
            MaTable!Somme = ws2.Cells(i, 7)
        Else
            MaTable!Num_client = ws2.Cells(i, 1)
            MaTable!Nom = ws2.Cells(i, 2)
            MaTable!prenom = ws2.Cells(i, 3)
            MaTable!Date_naissance = ws2.Cells(i, 4)
            MaTable!Somme = ws2.Cells(i, 5)
        End If
    Next i
    MaTable.Update
    MaTable.Close
    Set MaTable = Nothing
    
    ' Fermer le classeur
    wb.Close SaveChanges:=False
Next tableName

'Ferme la connexion.
cnn.Close
'Libère l'objet de la mémoire
Set cnn = Nothing
End Sub
' Procédure pour remplir les tables des rendements
Sub RemplirTableActifs(ByVal tableName As String, ByVal cpt As Integer)

Dim cnn As ADODB.Connection
Dim MaTable As ADODB.Recordset
Const ACCDB_Fournisseur = "Microsoft.ACE.OLEDB.12.0" ' Fournisseur de données.
Dim path As String
Dim Fichier_Source As String
path = ThisWorkbook.path & "\" & "basededonnees" & ".accdb"
Fichier_Source = path 'Fichier accdb ou mdb.

Dim wb As Workbook
Set wb = Workbooks.Open(ThisWorkbook.path & "\DATASET.xlsm")
Dim ws As Worksheet
Set ws = wb.Worksheets("Univers actifs")
' On divise le nb de colonnes en 4
Dim col As Long
col = WorksheetFunction.RoundUp(ws.Cells(1, ws.Columns.count).End(xlToLeft).Column / 4, 0)

' Initialisation d'un objet connexion.
Set cnn = New ADODB.Connection
'Declaration du curseur control qui enregistre la navigation, qui contrôle la visibilté des changements dans la db et qui controle son update
cnn.CursorLocation = adUseServer
' Initialisation de la connexion à la base avec declaration du fournisseur :
cnn.Open "Provider= " & ACCDB_Fournisseur & ";" _
& "Data Source=" & Fichier_Source _
& ";", "", "", adAsyncConnect
While (cnn.State = adStateConnecting): DoEvents: Wend ' Attente de la connexion.
    
Set MaTable = New ADODB.Recordset
MaTable.Open "SELECT * FROM " & tableName, cnn, adOpenKeyset, adLockPessimistic, adCmdText
MaTable.AddNew ' Passe la table en mode Ajout.

' On remplit les données à partir de la ligne 2
Dim i As Long, j As Long
For i = 2 To ws.UsedRange.Rows.count
    MaTable.AddNew
    MaTable![Date] = ws.Cells(i, 1)
    ' Dernière table va de 3*col + 2 et se finit à la colonne 4*col
    If cpt = 3 Then
        For j = cpt * col + 2 To (cpt + 1) * col
            MaTable.Fields(j - cpt * col - 1).Value = ws.Cells(i, j)
        Next j
    Else
        ' Les tables se suivent: _
        Table 1 va de la colonne 2 à col + 1 _
        Table 2 va de la colonne col + 2 à 2*col + 1 _
        Table 3 va de la colonne 2*col + 2 à 3*col + 1
        For j = cpt * col + 2 To (cpt + 1) * col + 1
            MaTable.Fields(j - cpt * col - 1).Value = ws.Cells(i, j)
        Next j
    End If
Next i

' On ferme le fichier, la table et la connexion et on libère de la mémoire
MaTable.Update
MaTable.Close
Set MaTable = Nothing
wb.Close SaveChanges:=False
cnn.Close
Set cnn = Nothing

End Sub
Sub RemplirTablesActifs() ' On exécute la procédure pour les quatre tables de rendements
RemplirTableActifs "rendements_actifs1", 0
RemplirTableActifs "rendements_actifs2", 1
RemplirTableActifs "rendements_actifs3", 2
RemplirTableActifs "rendements_actifs4", 3
End Sub
Sub RemplirActif()
'-------------------------------------------------------------------------------
Dim cnn As ADODB.Connection ' Objet representant une connexion a la base.
Dim MaTable As ADODB.Recordset ' Objet representant la table de la base.
Dim path As String
Dim Fichier_Source As String
Dim wb As Workbook
Dim ws As Worksheet
Const ACCDB_Fournisseur = "Microsoft.ACE.OLEDB.12.0" ' Fournisseur de donnees.

path = ThisWorkbook.path & "\" & "basededonnees" & ".accdb"
Fichier_Source = path  'Fichier accdb ou mdb.

' Initialisation d'un objet connexion.
Set cnn = New ADODB.Connection
'Declaration du curseur control qui enregistre la navigation, qui controlela visibilte des changements dans la db et qui controle son update
cnn.CursorLocation = adUseServer

' Initialisation de la connexion a la base avec declaration du fournisseur :
cnn.Open "Provider= " & ACCDB_Fournisseur & ";" _
& "Data Source=" & Fichier_Source _
& ";", "", "", adAsyncConnect
While (cnn.State = adStateConnecting): DoEvents: Wend ' Attente de la connexion.


'%%%% Remplissage de la table qui associe actif à son numéro
Set MaTable = New ADODB.Recordset ' Initialisation d'un objet table.
' Ouverture de la table en lecture/ecriture :
MaTable.Open "SELECT * FROM actifs", cnn, adOpenKeyset, adLockPessimistic, adCmdText
' On recopie les données du fichier dataset feuille code actif à partir de la ligne 1
Set wb = Workbooks.Open(ThisWorkbook.path & "\DATASET.xlsm")
Set ws = wb.Worksheets("code actifs")
Dim i As Long, j As Long
For i = 1 To ws.UsedRange.Rows.count
    MaTable.AddNew
    MaTable!Actif = ws.Cells(i, 1)
    MaTable!code_actif = ws.Cells(i, 2)
Next i

' On ferme la table et on libère de la mémoire.
MaTable.Update
MaTable.Close
Set MaTable = Nothing

'%%%% Remplissage de la table rendements mensuels
Set ws = wb.Worksheets("rdts_mensuels_t")
Set MaTable = New ADODB.Recordset
MaTable.Open "SELECT * FROM rendements_mensuels", cnn, adOpenKeyset, adLockPessimistic, adCmdText
' Remplir les données à partir de la ligne 2
For i = 2 To ws.UsedRange.Rows.count
    MaTable.AddNew
    MaTable!actifs = ws.Cells(i, 1)
    For j = 2 To ws.UsedRange.Columns.count ' Commence à la deuxième colonne
        MaTable.Fields(j - 1).Value = ws.Cells(i, j).Value ' Remplit les champs dans la table Access à partir de la deuxième colonne
    Next j
Next i
MaTable.Update
MaTable.Close
Set MaTable = Nothing

' On ferme le fichier et la connexion et on libère de la mémoire.
wb.Close SaveChanges:=False
cnn.Close
Set cnn = Nothing

End Sub



