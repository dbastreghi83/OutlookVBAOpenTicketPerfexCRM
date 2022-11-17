VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmPerfexCRM 
   Caption         =   "Open Ticket in PerfexCRM"
   ClientHeight    =   7656
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   8856.001
   OleObjectBlob   =   "frmPerfexCRM.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmPerfexCRM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btnOpenTicket_Click()
    PerfexCRM_OpenTicketPost txtSubject.Value, txtName.Value, txtEmail.Value, txtPriority.Value, txtMessage.Value, txtCC.Value
    Me.Hide
End Sub

Private Sub UserForm_Initialize()
    txtPriority.AddItem "1"
    txtPriority.AddItem "2"
    txtPriority.AddItem "3"
End Sub
