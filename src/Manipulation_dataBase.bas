Attribute VB_Name = "Manipulation_dataBase"
Sub NewInvest()
UserFormDeposit.Show 'lance le UserForm qui permet de saisir l'identit� du client et la somme qu'il d�sire d�poser puis lance DepotProc
End Sub
Sub DepotProc()
 Dim dbPath As String
    Dim conn As ADODB.Connection
    Dim rst As ADODB.Recordset
    Dim sql As String
    Dim i As Integer
    Dim fonds As Variant
    Dim amount As Double
    Dim tailleTotale As Double

    ' Chemin de la base de donn�es Access
    dbPath = ThisWorkbook.path & "\basededonnees.accdb"

    ' Connexion � la base de donn�es Access
    Set conn = New ADODB.Connection
    conn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & dbPath
    conn.Open

    ' Liste des fonds disponibles
    fonds = Array("Alpha", "Omega", "Omicron", "Gamma", "Theta")

    ' It�ration sur chaque fond pour effectuer la mise � jour
    For i = LBound(fonds) To UBound(fonds)
        ' R�cup�ration du montant � ajouter pour le fonds courant, bas� sur les informations saisies par l'utilisateur
        ' Si les textbox sont vides ou alors contiennent des cha�nes de caract�res, la valeur retenue est 0
        If UserFormDeposit.Controls("TextBox" & fonds(i)).Value = "" Or Not IsNumeric(UserFormDeposit.Controls("TextBox" & fonds(i)).Value) Then
            amount = 0
        Else
            amount = CDec(UserFormDeposit.TextBoxSomme.Value) * CDbl(UserFormDeposit.Controls("TextBox" & fonds(i)).Value) 'Val est utilis� pour convertir une cha�ne String en valeur num�rique, Controles indique la zone de texte
        End If
        
        If amount > 0 Then ' Permet d'�viter d'effectuer des requ�tes inutiles si amount = 0
            ' Mise � jour de la somme investie par le client dans le fonds courant
            sql = "UPDATE pilotage_investisseurs SET Somme_" & LCase(fonds(i)) & " = Somme_" & LCase(fonds(i)) & " + " & amount & _
                  ", Somme_investie_totale = Somme_investie_totale + " & amount & _
                  " WHERE Nom = '" & UCase(UserFormDeposit.TextBoxNom.Value) & "' AND Prenom = '" & UCase(UserFormDeposit.TextBoxPrenom.Value) & "'"
            conn.Execute sql

            ' Mise � jour de la taille du fond dans pilotage_fonds
            sql = "UPDATE pilotage_fonds SET Taille = Taille + " & amount & " WHERE Fonds = '" & fonds(i) & "'"
            conn.Execute sql
        End If
    Next i
    
' Obtention de la taille totale de gestion
sql = "SELECT SUM(Taille) AS Total FROM pilotage_fonds;"
Set rst = conn.Execute(sql)
If Not rst.EOF Then
    tailleTotale = rst.Fields("Total").Value
End If
rst.Close

Dim tailleTotaleConv As String
tailleTotaleConv = Replace(CStr(tailleTotale), ",", ".") ' Access utilise le point pour les d�cimales donc besoin de convertir la valeur
    
    ' Mise � jour des poids relatifs des fonds dans la boutique
    For i = LBound(fonds) To UBound(fonds)
        sql = "UPDATE pilotage_fonds SET Poids_boutique = (Taille / " & tailleTotaleConv & ") WHERE Fonds = '" & fonds(i) & "'"
        conn.Execute sql
    Next i

    conn.Close
End Sub
Sub SupInvest()
UserFormWithDraw.Show 'lance le UserForm qui permet � l'utilisateur de saisir l'idendit� du client et la somme qu'il souhaite retirer puis lance RetraitProc
End Sub
Sub RetraitProc()
Dim dbPath As String
    Dim conn As ADODB.Connection
    Dim rst As ADODB.Recordset
    Dim sql As String
    Dim i As Integer
    Dim fonds As Variant
    Dim amount As Double
    Dim tailleTotale As Double

    ' Chemin de la base de donn�es Access
    dbPath = ThisWorkbook.path & "\basededonnees.accdb"

    ' Connexion � la base de donn�es Access
    Set conn = New ADODB.Connection
    conn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & dbPath
    conn.Open

    ' Liste des fonds disponibles
    fonds = Array("Alpha", "Omega", "Omicron", "Gamma", "Theta")

    ' It�ration sur chaque fond pour effectuer la mise � jour
    For i = LBound(fonds) To UBound(fonds)
        ' R�cup�ration du montant � ajouter pour le fonds courant, bas� sur l'input de l'utilisateur
        ' Si les textbox sont vides ou alors contiennent des cha�nes de caract�res, la valeur retenue est 0
        If UserFormWithDraw.Controls("TextBox" & fonds(i)).Value = "" Or Not IsNumeric(UserFormWithDraw.Controls("TextBox" & fonds(i)).Value) Then
            amount = 0
        Else
            amount = CDec(UserFormWithDraw.TextBoxSomme.Value) * CDbl(UserFormWithDraw.Controls("TextBox" & fonds(i)).Value) 'Val est utilis� pour convertir une cha�ne String en valeur num�rique, Controles indique la zone de texte
        End If
        
        If amount > 0 Then ' Permet d'�viter d'effectuer des requ�tes inutiles si amount = 0
            ' Mise � jour de la somme investie par le client dans le fonds courant
            sql = "UPDATE pilotage_investisseurs SET Somme_" & LCase(fonds(i)) & " = Somme_" & LCase(fonds(i)) & " - " & amount & _
                  ", Somme_investie_totale = Somme_investie_totale - " & amount & _
                  " WHERE Nom = '" & UCase(UserFormWithDraw.TextBoxNom.Value) & "' AND Prenom = '" & UCase(UserFormWithDraw.TextBoxPrenom.Value) & "'"
            conn.Execute sql

            ' Mise � jour de la taille du fond dans pilotage_fonds
            sql = "UPDATE pilotage_fonds SET Taille = Taille - " & amount & " WHERE Fonds = '" & fonds(i) & "'"
            conn.Execute sql
        End If
    Next i


    ' Obtention de la taille totale de gestion
    sql = "SELECT SUM(Taille) AS Total FROM pilotage_fonds"
    Set rst = conn.Execute(sql)
    If Not rst.EOF Then
        tailleTotale = rst.Fields("Total").Value
    End If
    rst.Close
    
    Dim tailleTotaleConv As String
    tailleTotaleConv = Replace(CStr(tailleTotale), ",", ".") ' Access utilise le point pour les d�cimales donc besoin de convertir la valeur
    
    ' Mise � jour des poids relatifs des fonds dans la boutique
    For i = LBound(fonds) To UBound(fonds)
        sql = "UPDATE pilotage_fonds SET Poids_boutique = (Taille / " & tailleTotaleConv & ") WHERE Fonds = '" & fonds(i) & "'"
        conn.Execute sql
    Next i

    conn.Close
End Sub
Sub SupClient()
UserFormSupC.Show 'lance le Userform o� l'utilisateur rentre l'identit� du client qui quitte le fond puis lance DeleteProc
End Sub
Sub Deleteproc()

Dim dbPath As String
Dim conn As ADODB.Connection
Dim rst As ADODB.Recordset
Dim sql As String
Dim i As Integer

' Chemin de la base de donn_es Access
dbPath = ThisWorkbook.path & "\basededonnees.accdb"

' Connexion _ la base de donn_es Access
Set conn = New ADODB.Connection
conn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & dbPath
conn.Open

'Requte SQL qui modifie le pilotage fond car l'argent du client a �t� retir�e
Dim fonds As Variant
Dim amount As Double
fonds = Array("Alpha", "Omega", "Omicron", "Gamma", "Theta")

For i = LBound(fonds) To UBound(fonds)
    sql = "SELECT Somme_" & LCase(fonds(i)) & " AS montant FROM pilotage_investisseurs WHERE Nom = '" _
    & UCase(UserFormSupC.TextBoxNom.Value) & "' AND Prenom = '" & UCase(UserFormSupC.TextBoxPrenom.Value) & "';"
    Set rst = conn.Execute(sql)
    amount = rst.Fields("montant")
    Dim amountSTR As String
    amountSTR = Replace(CStr(amount), ",", ".") ' Besoin de convertir car Access aux normes anglaises
    'On soustrait le montant � la taille du fonds
    sql = "UPDATE pilotage_fonds SET Taille = Taille - " & amountSTR & " WHERE Fonds = '" & fonds(i) & "';"
    conn.Execute sql
    rst.Close
Next i

    'On update �galement la taille du fonds par rapport � la boutique
    ' Obtention de la taille totale de gestion
    sql = "SELECT SUM(Taille) AS Total FROM pilotage_fonds;"
    Set rst = conn.Execute(sql)
    If Not rst.EOF Then
        tailleTotale = rst.Fields("Total").Value
    End If
    Dim tailleTotaleConv As String
    tailleTotaleConv = Replace(CStr(tailleTotale), ",", ".") ' Access utilise le point pour les d�cimales donc besoin de convertir la valeur
    
    sql = "UPDATE pilotage_fonds SET poids_boutique = (Taille / " & tailleTotaleConv & ")"
    
' Requte SQL qui supprime la ligne dans pilotage investisseurs du client qui a quitte le fondd
sql = "DELETE FROM pilotage_investisseurs WHERE Nom = '" & UserFormSupC.TextBoxNom.Value & "' AND Prenom = '" & UserFormSupC.TextBoxPrenom.Value & "';"
conn.Execute sql

conn.Close
End Sub
Sub NewClient()
UserFormNewC.Show 'lance le UserForm o� l'utilisateur rentre les informations du nouveau client puis lance Addproc
End Sub
Sub Addproc()
Dim dbPath As String
Dim conn As ADODB.Connection
Dim rst As ADODB.Recordset
Dim sql As String
Dim tailleTotale As Double ' n�cessaire pour modifier le poids des fonds sur la boutique

' Chemin de la base de donn�es Access
dbPath = ThisWorkbook.path & "\basededonnees.accdb"

' Connexion � la base de donn�es Access
Set conn = New ADODB.Connection
conn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & dbPath
conn.Open

Dim Alpha, Omega, Omicron, Theta, Gamma, Somme As Double


'Condition pour v�rifier si les valeurs ins�r�es sont bien num�riques, puis les convertir en type Decimal afin de les manipuler
' Si les textbox sont vides ou alors contiennent des cha�nes de caract�res, la valeur retenue est 0
If UserFormNewC.TextBoxAlpha.Value = "" Or Not IsNumeric(UserFormNewC.TextBoxAlpha.Value) Then
    Alpha = 0
Else
    Alpha = CDec(UserFormNewC.TextBoxAlpha.Value)
End If

If UserFormNewC.TextBoxOmega.Value = "" Or Not IsNumeric(UserFormNewC.TextBoxOmega.Value) Then
    Omega = 0
Else
    Omega = CDec(UserFormNewC.TextBoxOmega.Value)
End If

If UserFormNewC.TextBoxOmicron.Value = "" Or Not IsNumeric(UserFormNewC.TextBoxOmicron.Value) Then
    Omicron = 0
Else
    Omicron = CDec(UserFormNewC.TextBoxOmicron.Value)
End If

If UserFormNewC.TextBoxTheta.Value = "" Or Not IsNumeric(UserFormNewC.TextBoxTheta.Value) Then
    Theta = 0
Else
    Theta = CDec(UserFormNewC.TextBoxTheta.Value)
End If

If UserFormNewC.TextBoxGamma.Value = "" Or Not IsNumeric(UserFormNewC.TextBoxGamma.Value) Then
    Gamma = 0
Else
    Gamma = CDec(UserFormNewC.TextBoxGamma.Value)
End If

'Somme est correct car on fait d�j� une v�rification dans la private sub
Somme = CDbl(UserFormNewC.TextBoxSomme.Value) '' Aucun probl�me pour faire des op�rations entre des Double et des Decimal


' G�n�rer un num�ro client
Dim NC As String
NC = Mid(UserFormNewC.TextBoxNom.Value, 1, 4) & Mid(UserFormNewC.TextBoxPrenom.Value, 1, 2)
NC = UCase(NC)
NC = NC & "1"

' Formater le nom et le pr�nom
Dim nomf, prenom As String
nomf = UCase(UserFormNewC.TextBoxNom.Value)
prenom = UCase(UserFormNewC.TextBoxPrenom.Value)

' Requ�te SQL qui ajoute toutes les informations sur une nouvelle ligne
' Pour les sommes de chaque fonds, il suffit de multiplier la somme par les valeurs de chaque textbox fonds
sql = "INSERT INTO pilotage_investisseurs (Num_client, Nom, Prenom, Mail, nom_Commune, nom_departement, nom_region, Somme_alpha, Somme_omega, Somme_omicron, Somme_gamma, Somme_theta, somme_investie_totale) VALUES ('" _
        & NC & "', '" & nomf & " ' ,'" & prenom & "', '" & UserFormNewC.TextBoxMail.Value & "', '" & UserFormNewC.TextBoxVille.Value & "', '" _
        & UserFormNewC.TextBoxDepartement.Value & "', '" & UserFormNewC.TextBoxRegion.Value & "', " _
        & Somme * Alpha & ", " & Somme * Omega _
        & ", " & Somme * Omicron & ", " & Somme * Gamma _
        & ", " & Somme * Theta & ", " & Somme & ");"
conn.Execute sql

'Requte SQL qui modifie les infos des fonds d� au nouvel investissementt
' La requte modifie la taillle de chaque fond d� � la somme  investi par le nouveau client et elle calcule ensuite le nouveau poids de chaque fondd
' par rapport au montant total sous gestion
Dim fondsNames As Variant
Dim fondsValues As Variant
Dim i As Integer

' Nom des fonds
fondsNames = Array("Alpha", "Omega", "Omicron", "Gamma", "Theta")

' Valeurs associ�es de chaque fond, suppose qu'il existe des TextBox pour chaque
fondsValues = Array(Somme * Alpha, Somme * Omega, Somme * Omicron, Somme * Gamma, Somme * Theta)

For i = LBound(fondsNames) To UBound(fondsNames)
    ' Construction et ex�cution de la requ�te SQL pour chaque fond
    sql = "UPDATE pilotage_fonds SET Taille = Taille + " & fondsValues(i) & " WHERE Fonds = '" & fondsNames(i) & "';"
    conn.Execute sql
Next i

' Obtention de la taille totale de gestion
sql = "SELECT SUM(Taille) AS Total FROM pilotage_fonds;"
Set rst = conn.Execute(sql)
If Not rst.EOF Then
    tailleTotale = rst.Fields("Total").Value
End If

Dim tailleTotaleConv As String
tailleTotaleConv = Replace(CStr(tailleTotale), ",", ".") ' Access utilise le point pour les d�cimales donc besoin de convertir la valeur
' Mise � jour du poids_boutique
sql = "UPDATE pilotage_fonds SET poids_boutique = (Taille / " & tailleTotaleConv & ")"
conn.Execute sql

conn.Close
Set conn = Nothing
End Sub



