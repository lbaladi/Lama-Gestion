VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormNewC 
   Caption         =   "UserForm2"
   ClientHeight    =   7692
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   15780
   OleObjectBlob   =   "UserFormNewC.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserFormNewC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()

Dim choix As Integer
Dim sumShare As Double

Do While TextBoxNom.Value = "" Or TextBoxPrenom.Value = ""
    Me.Hide
    choix = MsgBox("Veuillez entrer l'idendité du nouveau client", vbOKCancel + vbExclamation, "Identité requise")
    If choix = vbCancel Then
        Exit Sub
    ElseIf choix = vbOK Then
        Me.Show
    End If
Loop

Do While TextBoxSomme.Value = "" Or Not IsNumeric(TextBoxSomme.Value)
    choix = MsgBox("Veuillez entrer la somme investie par le nouveau client", vbOKCancel + vbExclamation, "Somme requise")
    If choix = vbOK Then
        TextBoxSomme.Value = InputBox("Somme investie par le nouveau client :", "Somme à investir")
    ElseIf choix = vbCancel Then
        Unload Me
        Exit Sub
    End If
Loop
    
Dim Alpha, Omega, Omicron, Theta, Gamma As Double

'Condition pour vérifier si les valeurs insérés sont bien numériques, puis les convertis en double afin de les manipuler
If TextBoxAlpha.Value = "" Or Not IsNumeric(TextBoxAlpha.Value) Then
    Alpha = 0
Else
    Alpha = CDbl(TextBoxAlpha.Value)
End If

If TextBoxOmega.Value = "" Or Not IsNumeric(TextBoxOmega.Value) Then
    Omega = 0
Else
    Omega = CDbl(TextBoxOmega.Value)
End If

If TextBoxOmicron.Value = "" Or Not IsNumeric(TextBoxOmicron.Value) Then
    Omicron = 0
Else
    Omicron = CDbl(TextBoxOmicron.Value)
End If

If TextBoxTheta.Value = "" Or Not IsNumeric(TextBoxTheta.Value) Then
    Theta = 0
Else
    Theta = CDbl(TextBoxTheta.Value)
End If

If TextBoxGamma.Value = "" Or Not IsNumeric(TextBoxGamma.Value) Then
    Gamma = 0
Else
    Gamma = CDbl(TextBoxGamma.Value)
End If

sumShare = Alpha + Omicron + Omega + Theta + Gamma
Do Until sumShare = 1
    Me.Hide
    choix = MsgBox("Veuillez revoir la répartition des fonds", vbCancel + vbExclamation, "Mauvaise répartition des fonds")
    If choix = vbIgnore Then
        Exit Sub
    ElseIf choix = vbAbort Then
        Exit Sub
    ElseIf choix = vbRetry Then
        Me.Show
    End If
Loop

Call Addproc

Me.Hide
End Sub
Private Sub CommandButton2_Click()
Unload Me
End Sub
Private Sub UserForm_Click()

End Sub
