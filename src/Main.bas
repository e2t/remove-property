Attribute VB_Name = "Main"
Option Explicit

Dim swApp As Object
Public currentDoc As ModelDoc2

Sub Main()
    Set swApp = Application.SldWorks
    Set currentDoc = swApp.ActiveDoc
    If currentDoc Is Nothing Then
        Exit Sub
    End If
    If currentDoc.GetType = swDocDRAWING Then
        MsgBox "Макрос не работает в чертежах.", vbExclamation
        Exit Sub
    End If
    MainForm.Show
End Sub

Sub RunExecution(propertyName As String, doc As ModelDoc2)
    Dim count As Integer
    count = RemovePropertyFromAllConfigurations(propertyName, doc)
    If count = 0 Then
        MsgBox "Свойство не найдено ни в одной конфигурации.", vbExclamation
    Else
        doc.SetSaveFlag
        MsgBox "Свойство было удалено " & Str(count) & " раз.", vbInformation
    End If
End Sub

Function RemovePropertyFromAllConfigurations(propertyName As String, doc As ModelDoc2) As Integer
    Dim i As Variant
    Dim count As Integer
    count = 0
    For Each i In doc.GetConfigurationNames
        Dim prpMgr As CustomPropertyManager
        Set prpMgr = doc.Extension.CustomPropertyManager(i)
        If prpMgr.Delete2(propertyName) = swCustomInfoDeleteResult_OK Then
            count = count + 1
        End If
    Next
    RemovePropertyFromAllConfigurations = count
End Function

Function ExitApp() 'mask for button
    Unload MainForm
    End
End Function
