VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Form 
   Caption         =   "Formulário"
   ClientHeight    =   9090
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12465
   OleObjectBlob   =   "FormEstudo.frx":0000
   StartUpPosition =   1  'CenterOwner
   WhatsThisHelp   =   -1  'True
End
Attribute VB_Name = "Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Sub ATUALIZA()
  Dim n As Integer
  Dim B As Range
  n = Range("A1").CurrentRegion.Rows.Count
  Set B = Range(Cells(2, 1), Cells(n, 4))
  ListBox1.RowSource = B.Address
End Sub

Private Sub BtFechar_Click()
  Unload Me
End Sub

Private Sub BtnImagem_Click()
    Dim C As Variant
    C = Application.GetOpenFilename(",*.bmp")
    If C = False Then Exit Sub
    Range("I1").Value = C
    ImgLoad.Picture = LoadPicture(C)
End Sub

Private Sub BtnNovo_Click()
    txtNome = Empty
    TxtModelo = Empty
    CBBMarca = Empty
    CBBCatgeoria = Empty
    Range("i1").Value = ""
    ImgLoad.Picture = LoadPicture("")
    ImMarca.Picture = LoadPicture("")
    txtNome.SetFocus
End Sub

Private Sub BtnSalvar_Click()
  Dim I As Integer
  
  I = Range("A1").CurrentRegion.Rows.Count + 1
  
  Cells(I, 1).Value = txtNome.Value
  Cells(I, 2).Value = TxtModelo.Value
  Cells(I, 3).Value = CBBCatgeoria.Value
  Cells(I, 4).Value = CBBMarca.Value
  Cells(I, 5).Value = Range("i1").Value
  
  Call ATUALIZA
  
  MsgBox "Produto Cadastrado Com Sucesso", vbInformation, "Informação"
  
End Sub

Private Sub CBBMarca_Change()
  Dim Cam As String
  
  Cam = ThisWorkbook.Path & "\Fotos\" & CBBMarca.Value & ".bmp"
  
  If CBBMarca.Value = "" Then
     ImMarca.Picture = LoadPicture("")
  Else: ImMarca.Picture = LoadPicture(Cam)
  End If
  
End Sub

Private Sub ListBox1_Click()
 Dim n As Integer
 n = ListBox1.ListIndex + 2
 txtNome.Value = Cells(n, 1).Value
 TxtModelo.Value = Cells(n, 2).Value
 CBBCatgeoria.Value = Cells(n, 3).Value
 CBBMarca.Value = Cells(n, 4).Value
 ImgLoad.Picture = LoadPicture(Cells(n, 5).Value)
End Sub


Private Sub UserForm_Initialize()
  Application.WindowState = xlMaximized
  With CBBCatgeoria
     .AddItem "Telefonia"
     .AddItem "Informática"
     .AddItem "Eletrodoméstica"
     .AddItem "TVS"
  End With
  With CBBMarca
     .AddItem "Apple"
     .AddItem "Brastemp"
     .AddItem "Consul"
     .AddItem "Dell"
     .AddItem "Electrolux"
     .AddItem "HP"
     .AddItem "LG"
     .AddItem "Motorola"
     .AddItem "Samsung"
     .AddItem "Sony"
  End With
  Call ATUALIZA
End Sub

