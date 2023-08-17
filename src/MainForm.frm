VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MainForm 
   Caption         =   "RemoveProperty 23.1"
   ClientHeight    =   7560
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3960
   OleObjectBlob   =   "MainForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Private Sub btnCancel_Click()

  ExitApp
    
End Sub

Private Sub btnRun_Click()

  Dim Props As Collection
  Dim I As Integer
  
  Set Props = New Collection
  For I = 0 To Me.ListBoxProperty.ListCount - 1
    If Me.ListBoxProperty.Selected(I) Then
      Props.Add Me.ListBoxProperty.List(I)
    End If
  Next
  
  RunExecution Props, Me.chkCommon.Value
    
End Sub

Private Sub UserForm_Initialize()

  InitForm

End Sub
