VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MainForm 
   Caption         =   "Удалить свойство"
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
    cmbProperty.AddItem "Наименование"
    cmbProperty.AddItem "Обозначение"
    cmbProperty.AddItem "Заготовка"
    cmbProperty.AddItem "Типоразмер"
    cmbProperty.AddItem "Материал"
    cmbProperty.AddItem "Формат"
    cmbProperty.AddItem "Примечание"
    cmbProperty.AddItem "Масса"
    cmbProperty.AddItem "Разработал"
    cmbProperty.AddItem "Длина"
    cmbProperty.AddItem "Ширина"

    'cmbProperty.AddItem "Пометка"
    'cmbProperty.AddItem "Тип документа"
    'cmbProperty.AddItem "Начертил"
    'cmbProperty.AddItem "Изменение"
    'cmbProperty.AddItem "Утвердил"
    'cmbProperty.AddItem "Техконтроль"
    'cmbProperty.AddItem "Проверил"
    'cmbProperty.AddItem "Нормоконтроль"
    'cmbProperty.AddItem "Организация"
End Sub
