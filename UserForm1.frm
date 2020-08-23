VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   9330
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15705
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cbbcidade_Change()
Dim base As Range
On Error Resume Next

Set base = Listas.Range("A1").CurrentRegion
cbbestado.Value = WorksheetFunction.VLookup(cbbcidade.Value, base, 2, 0)


End Sub

Private Sub CBBPESQUISARCIDADE_AfterUpdate()
Call FiltroAvancado
End Sub



Private Sub CBBPESQUISARESTADO_AfterUpdate()
Call FiltroAvancado
End Sub



Private Sub CBBPESQUISARESTADOCIVIL_AfterUpdate()
Call FiltroAvancado
End Sub



Private Sub CBBPESQUISARSEXO_AfterUpdate()
Call FiltroAvancado
End Sub



Private Sub CMBAPAGAR_Click()
For I = 1 To 21
    campos(I).Value = ""
Next

base.Range("imagem").ClearContents
imgFoto.Picture = LoadPicture("")
opbnao.Value = True
txtNome.SetFocus
End Sub

Private Sub CMBDELETAR_Click()
Dim resposta As VbMsgBoxResult
Dim nivel As Integer

nivel = BDUSUARIO.Range("NIVELATUAL").Value

    If nivel < 3 Then
        MsgBox "Esta pessoa não tem permissão para deletar registros!", vbCritical, "Atenção"
        Exit Sub
    End If


If UserForm1.LbModo.Caption = "Novo Item" Then Exit Sub

resposta = MsgBox("Tem certeza de que deseja deletar esse usuário?", vbYesNo + vbQuestion, "Atenção")

If resposta = vbYes Then Call Deletar: CMBNOVO_Click


End Sub

Private Sub CMBDELETARIMG_Click()
imgFoto.Picture = LoadPicture("")
base.Range("IMAGEM").Value = ""
End Sub

Private Sub cmbDeletarusuario_Click()
Dim nivel As Integer
If listaUsers.ListIndex = -1 Then Exit Sub
Dim resposta As VbMsgBoxResult

nivel = BDUSUARIO.Range("NIVELATUAL").Value

If nivel < 4 Then
MsgBox "Esta pessoa não tem permissão para configurar usuários!", vbCritical, "Atenção"
Exit Sub
End If

resposta = MsgBox("Tem certeza de que deseja deletar esse usuário?", vbYesNo + vbQuestion, "Atenção")

If resposta = vbYes Then Call DeletarUsuario(listaUsers.Value)
    
End Sub

Private Sub CMBEDITARUSUARIO_Click()
If listaUsers.ListIndex = -1 Then Exit Sub
Dim nivel As Integer
nivel = BDUSUARIO.Range("NIVELATUAL").Value

If nivel < 4 Then
MsgBox "Esta pessoa não tem permissão para configurar usuários!", vbCritical, "Atenção"
Exit Sub
End If

editarUsuario = True
UserForm2.Show
End Sub

Private Sub CMBFILTRAR_Click()
Call FiltroAvancado
End Sub

Private Sub CMBImagem_Click()
Dim caminho As Variant

caminho = Application.GetOpenFilename("Selecione a Imagem,*.bmp", , "Sistema de Cadastro")

If caminho = False Then Exit Sub
base.Range("imagem").Value = caminho
imgFoto.Picture = LoadPicture(caminho)

End Sub

Private Sub CMBLIMPAR_Click()

txtPESQUISARNOME.Value = ""
TXTPESQUISARCPF.Value = ""
CBBPESQUISARCIDADE.Value = ""
CBBPESQUISARESTADO.Value = ""
CBBPESQUISARESTADOCIVIL.Value = ""
CBBPESQUISARSEXO.Value = ""
TXTPESQUISARPAIS.Value = ""
IMGPESQUISARIMAGEM.Picture = LoadPicture("")


End Sub

Private Sub CMBNOVO_Click()
MultiPage1.Value = 0
LbModo.Caption = "Novo Item"
LBID.Caption = base.Range("ID").Value
CMBAPAGAR_Click

End Sub

Private Sub cmbNovoUsuario_Click()
Dim nivel As Integer

nivel = BDUSUARIO.Range("NIVELATUAL").Value

If nivel < 4 Then
MsgBox "Esta pessoa não tem permissão para configurar usuários!", vbCritical, "Atenção"
Exit Sub
End If


editarUsuario = False
UserForm2.Show
CMBUSUARIO_Click
End Sub

Private Sub CMBPESQUISAR_Click()
MultiPage1.Value = 1
CMBLIMPAR_Click
Call FiltroAvancado
txtPESQUISARNOME.SetFocus
End Sub

Private Sub cmbSair_Click()
Application.Visible = True
ThisWorkbook.Save
Unload Me
End Sub

Private Sub CMBSALVAR_Click()
Dim nivel As Integer
Dim I As Integer, DataN As Date
Dim emBranco As Boolean

nivel = BDUSUARIO.Range("NIVELATUAL").Value

    If nivel < 2 Then
        MsgBox "Esta pessoa não tem permissão para adicionar ou editar registros!", vbCritical, "Atenção"
        Exit Sub
    End If

Call AtribuiCampos

    For I = 1 To 21
        campos(I).BackColor = &HFFFFFF
        
        Select Case I
        Case Is = 1, 2, 3, 4, 5, 6, 7, 8, 9, 12, 13, 14, 15, 16, _
        17, 18
            
            If campos(I).Value = "" Then
               campos(I).BackColor = &HC0C0FF
               emBranco = True
            End If
        End Select
    Next

'CAMPO DEFICIENCIA
    If opbsim.Value = True Then
            If txtdeficiencia = "" Then
               txtdeficiencia.BackColor = &HC0C0FF
               emBranco = True
            End If
    End If
    
If emBranco = True Then
    MsgBox "Preencha os campos obrigatórios!", vbCritical, "Atenção"
    Exit Sub
End If


'VERIFICA A DATA
If IsDate(txtdatanasc.Value) = False Then
    MsgBox "Preencha uma data válida!", vbCritical, "Atenção"
    txtdatanasc.BackColor = &HC0C0FF
    txtdatanasc.Value = ""
    txtdatanasc.SetFocus
    Exit Sub
End If

Call PreencherBase
CMBAPAGAR_Click
CMBNOVO_Click

End Sub



Private Sub CMBUSUARIO_Click()

Dim nivel As Integer
Dim I As Integer, baseUsuarios As Range

nivel = BDUSUARIO.Range("NIVELATUAL").Value

If nivel < 4 Then
MsgBox "Esta pessoa não tem permissão para configurar usuários!", vbCritical, "Atenção"
Exit Sub
End If

MultiPage1.Value = 2

I = BDUSUARIO.Range("A1").CurrentRegion.Rows.Count
Set baseUsuarios = BDUSUARIO.Range(BDUSUARIO.Cells(2, 1), BDUSUARIO.Cells(I, 3))

listaUsers.RowSource = baseUsuarios.Address(, , , True)



End Sub

Private Sub Label54_Click()
CMBNOVO_Click
End Sub

Private Sub Label55_Click()
CMBPESQUISAR_Click
End Sub

Private Sub Label56_Click()
CMBUSUARIO_Click
End Sub

Private Sub Label57_Click()
cmbSenha_Click
End Sub

Private Sub Label58_Click()
cmbSair_Click
End Sub




End Sub

Private Sub opbnao_Click()
If opbnao.Value = True Then txtdeficiencia.Enabled = False
End Sub

Private Sub opbsim_Click()
If opbsim.Value = True Then
    txtdeficiencia.Enabled = True
    txtdeficiencia.SetFocus
End If
End Sub


Private Sub txtaltura_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

If KeyAscii = 44 Then Exit Sub
If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0

End Sub

Private Sub txtcep_Change()
Dim cep As String, cep2 As String, cep3 As String
Dim I  As Integer, j As Integer, n As Integer

cep = txtcep.Value
txtcep.MaxLength = 9

I = Len(cep)

    For j = 1 To I
        If IsNumeric(Mid(cep, j, 1)) Then
            cep2 = cep2 & Mid(cep, j, 1)
        End If
    Next
    
I = Len(cep2)
    For j = 1 To I
        cep3 = cep3 & Mid(cep2, j, 1)
        If j = 6 Then
        n = Len(cep3) - 1
        cep3 = Left(cep3, n) & "-" & Right(cep3, 1)
        End If
    Next

txtcep.Value = cep3

End Sub

Private Sub txtcep_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
End Sub

Private Sub txtcpf_Change()

Dim CPF As String, CPF2 As String, CPF3 As String
Dim I  As Integer, j As Integer, n As Integer

CPF = txtcpf.Value
txtcpf.MaxLength = 14

I = Len(CPF)

    For j = 1 To I
        If IsNumeric(Mid(CPF, j, 1)) Then
            CPF2 = CPF2 & Mid(CPF, j, 1)
        End If
    Next
    
I = Len(CPF2)
    For j = 1 To I
        CPF3 = CPF3 & Mid(CPF2, j, 1)
        If j = 4 Or j = 7 Then
        n = Len(CPF3) - 1
            CPF3 = Left(CPF3, n) & "." & Right(CPF3, 1)
        ElseIf j = 10 Then
         n = Len(CPF3) - 1
            CPF3 = Left(CPF3, n) & "-" & Right(CPF3, 1)
        End If
    Next

txtcpf.Value = CPF3

End Sub

Private Sub txtcpf_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0

End Sub

Private Sub txtdatanasc_Change()
Dim dataNasc As String, datanasc2 As String, datanasc3 As String
Dim I  As Integer, j As Integer, n As Integer

dataNasc = txtdatanasc.Value
txtdatanasc.MaxLength = 10

I = Len(dataNasc)

    For j = 1 To I
        If IsNumeric(Mid(dataNasc, j, 1)) Then
            datanasc2 = datanasc2 & Mid(dataNasc, j, 1)
        End If
    Next
    
I = Len(datanasc2)
    For j = 1 To I
        datanasc3 = datanasc3 & Mid(datanasc2, j, 1)
        If j = 3 Or j = 5 Then
        n = Len(datanasc3) - 1
            datanasc3 = Left(datanasc3, n) & "/" & Right(datanasc3, 1)
        End If
    Next

txtdatanasc.Value = datanasc3
End Sub

Private Sub txtdatanasc_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
End Sub


Private Sub TXTPESQUISARCPF_AfterUpdate()
Call FiltroAvancado
End Sub



Private Sub txtPESQUISARNOME_AfterUpdate()
Call FiltroAvancado
End Sub


Private Sub TXTPESQUISARPAIS_AfterUpdate()
Call FiltroAvancado
End Sub



Private Sub txtRenda_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
If KeyAscii = 44 Then Exit Sub
If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
End Sub

Private Sub txtRg_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
End Sub

Private Sub UserForm_Initialize()
Call CarregarCombobox
Call AtribuiCampos
LBID.Caption = base.Range("id").Value
MultiPage1.Value = 0
txtNome.SetFocus
txtpais.Value = "Brasil"
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

'    If Workbooks.Count = 1 Then
'        Application.Visible = True
'       Application.Quit
'    Else
'        Application.Visible = True
'        ThisWorkbook.Close True
'    End If

End Sub
