VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormWithDraw 
   Caption         =   "UserForm4"
   ClientHeight    =   8304.001
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   12096
   OleObjectBlob   =   "UserFormWithDraw.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserFormWithDraw"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()

Dim choix As Integer
Dim sumShare As Double

'Si l'utilisateur n'a pas saisis l'identité du client, il faut remplir à nouveau le userform
Do While TextBoxNom.Value = "" Or TextBoxPrenom.Value = ""
    Me.Hide
    choix = MsgBox("Veuillez entrer l'idendité du  client", vbOKCancel + vbExclamation, "Identité requise")
    If choix = vbCancel Then
        Exit Sub
    ElseIf choix = vbOK Then
        Me.Show
    End If
Loop

'Si la somme n'a pas été saisi, ou si une chaine de caractère est présente dans la textbox, on soumet alors une inputbox
' à l'utilisateur pour qu'il puisse saisir la somme à investir
Do While TextBoxSomme.Value = "" Or Not IsNumeric(TextBoxSomme.Value)
    choix = MsgBox("Veuillez entrer la somme ajoutée par le client", vbOKCancel + vbExclamation, "Somme requise")
    If choix = vbOK Then
        TextBoxSomme.Value = InputBox("Somme retirée par le client :", "Somme à retirer")
    ElseIf choix = vbCancel Then
        Unload Me
        Exit Sub
    End If
Loop
    
Dim Alpha, Omega, Omicron, Theta, Gamma As Double

'Condition pour vérifier si les valeurs insérées sont bien numériques, puis les convertis en double afin de les manipuler
' Si la textbox est vide, alors on considère que la proportion est égale à 0
If TextBoxAlpha.Value = "" Or Not IsNumeric(TextBoxAlpha.Value) Then
    Alpha = 0
Else
    Alpha = Val(Me.Controls("TextBoxAlpha"))
End If

If TextBoxOmega.Value = "" Or Not IsNumeric(TextBoxOmega.Value) Then
    Omega = 0
Else
    Omega = Val(Me.Controls("TextBoxOmega"))
End If

If TextBoxOmicron.Value = "" Or Not IsNumeric(TextBoxOmicron.Value) Then
    Omicron = 0
Else
    Omicron = Val(Me.Controls("TextBoxOmicron"))
End If

If TextBoxTheta.Value = "" Or Not IsNumeric(TextBoxTheta.Value) Then
    Theta = 0
Else
    Theta = Val(Me.Controls("TextBoxTheta"))
End If

If TextBoxGamma.Value = "" Or Not IsNumeric(TextBoxGamma.Value) Then
    Gamma = 0
Else
    Gamma = Val(Me.Controls("TextBoxGamma"))
End If

sumShare = Alpha + Omicron + Omega + Theta + Gamma
Do While sumShare <> 1
    Me.Hide
    choix = MsgBox("Veuillez revoir la répartition des fonds", vbOKCancel + vbExclamation, "Mauvaise répartition des fonds")
    If choix = vbCancel Then
        Exit Sub
    ElseIf choix = vbOK Then
        Me.Show
    End If
Loop

Call RetraitProc

Me.Hide
End Sub

Private Sub CommandButton2_Click()
Unload Me
End Sub


Private Sub UserForm_Click()

End Sub

