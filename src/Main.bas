Attribute VB_Name = "Main"
Option Explicit

Dim swApp As Object
Dim CurrentDoc As ModelDoc2

Sub Main()
    Set swApp = Application.SldWorks
    Set CurrentDoc = swApp.ActiveDoc
    If Not CurrentDoc Is Nothing Then
        If CurrentDoc.GetType = swDocDRAWING Then
            MsgBox "Макрос не работает в чертежах.", vbExclamation
        Else
            MainForm.Show
        End If
    End If
End Sub

Function InitForm() 'hide
    Dim Props As Variant
    Dim I As Variant

    Props = GetAllProperties(CurrentDoc)
    QuickSort Props, LBound(Props), UBound(Props)
    MainForm.ListBoxProperty.Clear
    For Each I In Props
        MainForm.ListBoxProperty.AddItem I
    Next
End Function

Sub RunExecution(Props As Collection, IsCommonInclude As Boolean)
    Dim I As Integer
    
    For I = 1 To Props.Count
        RemovePropertyFromAllConfigurations Props(I), CurrentDoc, IsCommonInclude
    Next
    CurrentDoc.SetSaveFlag
    InitForm
End Sub

Function ExitApp() 'mask for button
    Unload MainForm
    End
End Function
