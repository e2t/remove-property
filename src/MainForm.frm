VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MainForm 
   Caption         =   "������� ��������"
   ClientHeight    =   2415
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3930
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
    RunExecution cmbProperty.Text, currentDoc
End Sub

Private Sub UserForm_Initialize()
    cmbProperty.AddItem "������������"
    cmbProperty.AddItem "�����������"
    cmbProperty.AddItem "���������"
    cmbProperty.AddItem "����������"
    cmbProperty.AddItem "��������"
    cmbProperty.AddItem "������"
    cmbProperty.AddItem "����������"
    cmbProperty.AddItem "�����"
    cmbProperty.AddItem "����������"
    cmbProperty.AddItem "�����"
    cmbProperty.AddItem "������"

    'cmbProperty.AddItem "�������"
    'cmbProperty.AddItem "��� ���������"
    'cmbProperty.AddItem "��������"
    'cmbProperty.AddItem "���������"
    'cmbProperty.AddItem "��������"
    'cmbProperty.AddItem "�����������"
    'cmbProperty.AddItem "��������"
    'cmbProperty.AddItem "�������������"
    'cmbProperty.AddItem "�����������"
End Sub
