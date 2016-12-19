VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsErrorLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim colErrorLog As Collection

Private Sub Class_Initialize()
    Set colErrorLog = New Collection
End Sub

Private Sub Class_Terminate()
    Set colErrorLog = Nothing
End Sub

Public Property Get Count() As Variant
    Count = colErrorLog.Count
End Property

Public Property Get get_Log() As Variant
    Dim l As Long
    Dim strAusgabe As String
    
    For l = 1 To colErrorLog.Count
        strAusgabe = strAusgabe & vbTab & colErrorLog.Item(l) & vbNewLine
    Next l
    
    get_Log = strAusgabe
End Property

Public Sub add(ByVal ErrorSource As String, ByVal ErrorDescription As String)
    colErrorLog.add (Now & vbTab & Environ("UserName") & vbTab & ErrorSource & vbTab & ErrorDescription)
End Sub
