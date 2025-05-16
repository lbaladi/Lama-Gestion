VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormSupC 
   Caption         =   "UserForm3"
   ClientHeight    =   4908
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   18156
   OleObjectBlob   =   "UserFormSupC.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserFormSupC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
Dim choix As Integer

Do While TextBoxNom.Value = "" Or TextBoxPrenom.Value = ""
    Me.Hide
    choix = MsgBox("Veuillez donner l'identité complète du client désirant quitter la boutique", vbOKCancel + vbExclamation, "Identité requise")
    If choix = vbOK Then
        Me.Show
    ElseIf choix = vbCancel Then
        Unload Me
        Exit Sub
    End If
Loop

Call Deleteproc

Me.Hide

End Sub

Private Sub CommandButton2_Click()
Unload Me
End Sub

Private Sub UserForm_Click()

End Sub


