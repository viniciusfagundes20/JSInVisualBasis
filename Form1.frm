VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "teste"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

    Call rodarScript

End Sub


Private Sub rodarScript()

    Dim Script As New ScriptControl
    
    With Script
        
        .Language = "JScript"
        .AddCode " function Sum(a, b) { return a + b; } "
        .AddCode " function Sub(a, b) { return a - b; } "
        
        Debug.Print .Run("Sum", 1, 2)
        
        Debug.Print .Run("Sub", 5, 2)
        
    End With

End Sub
