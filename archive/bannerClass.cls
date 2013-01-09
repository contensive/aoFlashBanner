VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "bannerClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private banWidth As String
Private banHeight As String

Private Main As Object
Private CSV As Object

Public Function Execute(CsvObject As Object, MainObject As Object, OptionString As String, FilterInput As String) As String
    Set CSV = CsvObject
    
    Call Init(MainObject)
    
    Execute = GetContent(OptionString)
    
End Function

Public Sub Init(MainObject As Object)
    Set Main = MainObject
    
    Exit Sub

ErrorTrap:
    Call HandleError("bannerClass", "Init", Err.Number, Err.Source, Err.Description, True, False)
    End Sub

Public Function GetContent(OptionString As String) As String
    On Error GoTo ErrorTrap
    
    Dim Stream As String
    
    banWidth = Main.GetAggrOption("Banner Width", OptionString)
    banHeight = Main.GetAggrOption("Banner Height", OptionString)
    
    If banWidth = "" Then
        banWidth = "100%"
    ElseIf (InStr(1, banWidth, "px") = 0) And (InStr(1, banWidth, "%") = 0) Then
        banWidth = banWidth & "px"
    End If
    If banHeight = "" Then
        banHeight = "300px"
    ElseIf (InStr(1, banHeight, "px") = 0) Then
        If InStr(1, banHeight, "%") <> 0 Then
            banHeight = Replace(banHeight, "%", "")
        End If
        banHeight = banHeight & "px"
    End If

    Stream = Stream & "<div style=""width:" & banWidth & ";height:" & banHeight & """ id=""bannerContainer"">"
    Stream = Stream & "<div id=""flashBanner""></div>"
    Stream = Stream & "</div>"
    
    GetContent = Stream
    
    Exit Function
ErrorTrap:
    Call HandleError("bannerClass", "GetContent", Err.Number, Err.Source, Err.Description, True, False)
    End Function
