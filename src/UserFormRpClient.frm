VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormRpClient 
   Caption         =   "UserForm3"
   ClientHeight    =   6972
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   17976
   OleObjectBlob   =   "UserFormRpClient.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserFormRpClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()

Dim cpt, i As Integer
cpt = 0
For i = 0 To Me.ListBox1.ListCount - 1
    If Me.ListBox1.Selected(i) = True Then
        cpt = cpt + 1
        Exit For
    End If
Next i

If cpt = 0 Then
    MsgBox "Aucun client sélectionné"
    Unload Me
    Exit Sub

Else
    Call GenererReportingClient
End If


Me.Hide
End Sub

Private Sub CommandButton2_Click()
Unload Me
End Sub

Private Sub ListBox1_Click()

End Sub


Private Sub UserForm_Initialize()
    Dim conn As Object, rs As Object
    Dim cheminDB As String
    Dim sql As String

    ' Chemin vers la base de donn_es, adapt_ pour tre dynamiquee
    cheminDB = ThisWorkbook.path & "\basededonnees.accdb"

    ' D_finition de la requte SQL pour s_lectionner les num_ros de clientt
    sql = "SELECT num_client FROM pilotage_investisseurs"

    ' Cr_ation d'un nouvel objet connexion
    Set conn = CreateObject("ADODB.Connection")

    ' Ouverture de la connexion ö la base de donn_es Access
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & cheminDB

    ' Cr_ation d'un objet Recordset pour contenir les r_sultats de la requtee
    Set rs = CreateObject("ADODB.Recordset")

    ' Ex_cution de la requtee
    rs.Open sql, conn

    ' V_rification que le recordset contient des donn_es
    If Not rs.EOF Then
        rs.MoveFirst ' S'assurer de commencer au d_but du jeu de r_sultats
        With Me.ListBox1
            .Clear ' Nettoie les entr_es existantes
            While Not rs.EOF
                ' Ajout de chaque num_ro de client dans la ComboBox
                .AddItem rs!Num_client
                rs.MoveNext ' Passer ö l'enregistrement suivant
            Wend
            .MultiSelect = fmMultiSelectMulti
        End With
    End If

    ' Fermeture du recordset et de la connexion
    rs.Close
    conn.Close

    ' Lib_ration des ressources
    Set rs = Nothing
    Set conn = Nothing
End Sub

