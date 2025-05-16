VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Reporting Fond"
   ClientHeight    =   6612
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   13548
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub UserForm_Initialize()
With Me.ComboBoxFd
.AddItem "Theta"
.AddItem "Alpha"
.AddItem "Omega"
.AddItem "Gamma"
.AddItem "Omicron"
End With
End Sub
Private Sub ButtonAnnuler_Click()
Unload Me
End Sub
Private Sub ButtonValider_Click()
Dim fond As String
Dim choix As Integer
'Attribuer la valeur selectionn_e _ la variable fond
fond = Me.ComboBoxFd.Value

'Vérification erreur
Do While fond = ""
    Me.Hide
   choix = MsgBox("Veuillez choisir un fond ou d'annuler sinon", vbOKCancel + vbExclamation, "Sélection requise") 'vbExclamation fait apparaître msgbox en message d'alerte
   If choix = vbOK Then
        Me.Show
    ElseIf choix = vbCancel Then
        Exit Sub
    End If
Loop

'Appel _ la macro pour g_n_rer le rapport
Call GenererRapportFond(fond)

Unload Me

End Sub

