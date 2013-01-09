VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "responseClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private CSV As Object
Private Main As Object

Public Function Execute(CsvObject As Object, MainObject As Object, OptionString As String, FilterInput As String) As String
    Set CSV = CsvObject
    
    Call Init(MainObject)
    
    Execute = GetContent(OptionString)
    
End Function

Public Sub Init(MainObject As Object)
    Set Main = MainObject
    
    Exit Sub
ErrorTrap:
    Call HandleError("responseClass", "Init", Err.Number, Err.Source, Err.Description, True, False)
End Sub

Public Function GetContent(OptionString As String) As String
    On Error GoTo ErrorTrap

    Dim CS As Long
    Dim Stream As String
    Dim Color As String
    Dim Position As Double
    Dim Caption As String
    Dim Link As String
    Dim Image As String

    Stream = Main.ReadBake("flashBanner")

    If Stream = "" Then
        CS = Main.OpenCSContent("Flash Banners", "", "SortOrder")
        If Main.IsCSOK(CS) Then
            Stream = "<images>" & vbCrLf
            Do While Main.CSOK(CS)
            
                Color = Main.GetCS(CS, "CaptionColor")
                Position = Main.GetCSNumber(CS, "Position")
                Caption = Main.GetCS(CS, "Caption")
                Link = Main.GetCS(CS, "Link")
                Image = Main.GetCS(CS, "ImageFileName")
                
                If Image <> "" Then
                    Image = Main.ServerFilePath & Image
                End If
                If Color = "" Then
                    Color = "#404040"
                End If
                If Position = 0 Then
                    Position = 37
                End If
                
                If Image <> "" Then
                    Stream = Stream & "<item>" & vbCrLf
                    Stream = Stream & "<headline color=""" & Color & """ y=""" & Position & """><![CDATA[" & vbCrLf
                    Stream = Stream & Caption & vbCrLf
                    Stream = Stream & "]]></headline>" & vbCrLf
                    If Link <> "" Then
                        Stream = Stream & "<link>" & Link & "</link>" & vbCrLf
                    End If
                    Stream = Stream & "<location>" & Image & "</location>" & vbCrLf
                    Stream = Stream & "</item>" & vbCrLf
                End If
                
                Call Main.NextCSRecord(CS)
            Loop
            Stream = Stream & "</images>" & vbCrLf
        End If
        Call Main.CloseCS(CS)
        
        Call Main.SaveBake("flashBanner", Stream, "Flash Banners")
        
    End If
    
    GetContent = Stream
    
    Exit Function
ErrorTrap:
    Call HandleError("responseClass", "GetContent", Err.Number, Err.Source, Err.Description, True, False)
End Function

