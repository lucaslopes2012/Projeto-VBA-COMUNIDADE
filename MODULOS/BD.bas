Attribute VB_Name = "BD"
Option Explicit
Global BD As New ADODB.Connection, OPBD As New ADODB.Recordset
Global PathBD As String
Public Sub CONECTION()
  'ARQ = ThisWorkbook.Path
  'ARQ = Left(ARQ, InStr(ARQ, "\ADM")) & "ADM\BASE DE DADOS\SIS_CONTROLE_DELIVERY1.accdb;"
  BD.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;" & "Data Source=" & PathBD
  BD.Open
End Sub
Public Sub CLOSECONECTION()
  If OPBD.State = 1 Then OPBD.Close: Set OPBD = Nothing
  If BD.State = 1 Then BD.Close: Set BD = Nothing
End Sub
