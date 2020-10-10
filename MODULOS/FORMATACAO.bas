Attribute VB_Name = "FORMATACAO"
Option Explicit
Global CAMPOS(1 To 10) As Object
Public Sub CARREGACOMBOBOXS()
  Dim CBBOXESTADO As Range, CONT As Integer
  
  'BASE COMBOBOX ESTADOS
  CONT = COMOBOBOX.Range("A1").CurrentRegion.Rows.Count
  Set CBBOXESTADO = COMOBOBOX.Range(COMOBOBOX.Cells(1, 1), COMOBOBOX.Cells(CONT, 1))
  FormAssociados.CBBESTADO.RowSource = CBBOXESTADO.Address(, , , True)
  
  'BASE COMBOBOX ESTADOS
  CONT = COMOBOBOX.Range("A1").CurrentRegion.Rows.Count
  Set CBBOXESTADO = COMOBOBOX.Range(COMOBOBOX.Cells(1, 1), COMOBOBOX.Cells(CONT, 1))
  FormAssociados.CBBRECEBEDOR.RowSource = CBBOXESTADO.Address(, , , True)
End Sub

Public Sub FORMATACAMPOS()
  Set CAMPOS(1) = FormCadastro.TxtNomecompleto
  Set CAMPOS(2) = FormCadastro.TxtApelido
  Set CAMPOS(3) = FormCadastro.TxtCPF
  Set CAMPOS(4) = FormCadastro.TxtCNPJ
  Set CAMPOS(5) = FormCadastro.TxtTelCel
  Set CAMPOS(6) = FormCadastro.TxtNomeTitular
  Set CAMPOS(7) = FormCadastro.CBBBANCO
  Set CAMPOS(8) = FormCadastro.TxtAgencia
  Set CAMPOS(9) = FormCadastro.CBBTIPOCONTA
  Set CAMPOS(10) = FormCadastro.TxtNconta
End Sub

Public Sub EXIBICAMPOS(ByVal CODIGO As Integer)
  
  Dim SQL As String, I As Integer
  
  SQL = "Select * From Colaborador Where Código = " & CODIGO & ";"
  
  BDCONSULTA.CursorType = adOpenKeyset
  Call MACROS.BANCO(SQL)
  
  
  Call FORMATACAMPOS

  For I = 1 To 10
    With BDCONSULTA
     Select Case .Fields(I).Name
       Case Is = "Nome_Completo"
         CAMPOS(I).Value = .Fields.Item(I).Value
       Case Is = "Apelido"
         CAMPOS(I).Value = .Fields.Item(I).Value
       Case Is = "CPF"
         CAMPOS(I).Value = .Fields.Item(I).Value
       Case Is = "CNPJ"
         CAMPOS(I).Value = .Fields.Item(I).Value
       Case Is = "Contato"
         CAMPOS(I).Value = .Fields.Item(I).Value
       Case Is = "Titular_Conta"
         CAMPOS(I).Value = .Fields.Item(I).Value
       Case Is = "Banco"
         If .Fields.Item(I).Value = "S/ CONTA" Then
           FormCadastro.CkBoxSConta = True
         Else: CAMPOS(I).Value = .Fields.Item(I).Value
         End If
       Case Is = "Agencia_Conta"
         CAMPOS(I).Value = .Fields.Item(I).Value
       Case Is = "Tipo_Conta"
         CAMPOS(I).Value = .Fields.Item(I).Value
       Case Is = "Numero_Conta"
         CAMPOS(I).Value = .Fields.Item(I).Value
     End Select
    End With
  Next
  
  MACROS.FECHA_BANCO
End Sub
Public Sub CARREGALISTBOX()
   Dim SQL As String, Lin As Integer
   
   SQL = "SELECT * FROM Colaborador ORDER BY Apelido;"
   
   BDCONSULTA.CursorType = adOpenKeyset
   Call MACROS.BANCO(SQL)
   
   
   On Error GoTo ErrorHandler
     With FormBoard.GRID_LISTA
      .AddItem
      .List = LISTAGEMBASE.Range("A1:K1").Value
      .Clear
      Lin = 0
      Do Until BDCONSULTA.EOF
        .AddItem
        .List(Lin, 0) = "" & BDCONSULTA!Código
        .List(Lin, 1) = "" & BDCONSULTA!Apelido
        .List(Lin, 2) = "" & BDCONSULTA!Nome_Completo
        .List(Lin, 3) = "" & BDCONSULTA!CPF
        .List(Lin, 4) = "" & BDCONSULTA!CNPJ
        .List(Lin, 5) = "" & BDCONSULTA!Contato
        .List(Lin, 6) = "" & BDCONSULTA!Titular_Conta
        .List(Lin, 7) = "" & BDCONSULTA!BANCO
        .List(Lin, 8) = "" & BDCONSULTA!Agencia_Conta
        .List(Lin, 9) = "" & BDCONSULTA!Tipo_Conta
        .List(Lin, 10) = "" & BDCONSULTA!Numero_Conta
        Lin = Lin + 1
        BDCONSULTA.MoveNext
      Loop
      .ColumnWidths = "0 pt;100 pt;160 pt;70 pt;90 pt;80 pt;160 pt;60 pt;50 pt;70 pt;40 pt"
    End With
  Call MACROS.FECHA_BANCO
ErrorHandler: If Err.Number <> 0 Then Call MACROS.FECHA_BANCO
End Sub
Public Sub CARREGALABELS()
   Dim I As Integer
   With FormBoard
     Set LABELS(1) = .LB1
     Set LABELS(2) = .LB2
     Set LABELS(3) = .LB3
     Set LABELS(4) = .LB4
     Set LABELS(5) = .LB5
     Set LABELS(6) = .LB6
     Set LABELS(7) = .LB7
     Set LABELS(8) = .LB8
     Set LABELS(9) = .LB9
     Set LABELS(10) = .LB10
   End With
   
   For I = 1 To 10
     LABELS(I).Caption = LISTAGEMBASE.Cells(1, I + 1).Value
   Next
  
End Sub
