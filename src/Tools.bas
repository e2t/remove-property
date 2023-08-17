Attribute VB_Name = "Tools"
Option Explicit

Sub RemovePropertyFromAllConfigurations(PropertyName As String, Doc As ModelDoc2, _
                                        IsCommonInclude As Boolean)
    Dim I As Variant
    Dim Count As Integer 'not used
    
    Count = 0
    For Each I In Doc.GetConfigurationNames
        RemoveProperty PropertyName, Doc, I, Count
    Next
    If IsCommonInclude Then
        RemoveProperty PropertyName, Doc, "", Count
    End If
End Sub

Sub RemoveProperty(PropertyName As String, Doc As ModelDoc2, ByVal Conf As String, ByRef Count As Integer)
    Dim PrpMgr As CustomPropertyManager
    
    Set PrpMgr = Doc.Extension.CustomPropertyManager(Conf)
    If PrpMgr.Delete2(PropertyName) = swCustomInfoDeleteResult_OK Then
        Count = Count + 1
    End If
End Sub

Sub AppendProperties(PropMgr As CustomPropertyManager, ByRef Props As Dictionary)
    Dim I As Variant
    Dim Prop As String
    Dim AllProps As Variant
    
    AllProps = PropMgr.GetNames
    If IsEmpty(AllProps) Then
        Exit Sub
    End If
    For Each I In AllProps
        Prop = StrConv(I, vbProperCase)
        If Not Props.Exists(Prop) Then
            Props.Add Prop, 0
        End If
    Next
End Sub

Function GetAllProperties(Doc As ModelDoc2) As Variant
    Dim Props As Dictionary
    Dim DocExt As ModelDocExtension
    Dim I As Variant
    
    Set Props = New Dictionary
    Set DocExt = Doc.Extension
    AppendProperties DocExt.CustomPropertyManager(""), Props

    For Each I In Doc.GetConfigurationNames
        AppendProperties DocExt.CustomPropertyManager(I), Props
    Next
    
    GetAllProperties = Props.Keys
End Function
