VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmBordero 
   Caption         =   "Remessa de Borderôs para Cobrança"
   ClientHeight    =   3960
   ClientLeft      =   1230
   ClientTop       =   1470
   ClientWidth     =   6705
   LinkTopic       =   "Form9"
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3960
   ScaleWidth      =   6705
   Begin VB.ComboBox cmbBordero 
      Height          =   315
      Left            =   4455
      TabIndex        =   5
      Text            =   "cmbBordero"
      Top             =   225
      Width           =   1500
   End
   Begin VB.CommandButton cmdGravacao 
      Caption         =   "&Gravar"
      Height          =   420
      Left            =   4635
      TabIndex        =   17
      Top             =   1575
      Width           =   1230
   End
   Begin VB.CommandButton cmdSaida 
      Caption         =   "&Retornar"
      Height          =   420
      Left            =   4635
      TabIndex        =   18
      Top             =   3375
      Width           =   1230
   End
   Begin VB.CommandButton cmdAdicao 
      Caption         =   "&Adicionar"
      Height          =   330
      Left            =   1620
      TabIndex        =   13
      Top             =   1935
      Width           =   1320
   End
   Begin VB.CommandButton cmdRemocao 
      Caption         =   "Remo&ver"
      Height          =   330
      Left            =   1620
      TabIndex        =   14
      Top             =   2385
      Width           =   1320
   End
   Begin VB.CommandButton cmdRemocaoTotal 
      Caption         =   "Remover &Todos"
      Height          =   330
      Left            =   1620
      TabIndex        =   15
      Top             =   2835
      Width           =   1320
   End
   Begin VB.Frame Frame1 
      Caption         =   "Tipo de Operação"
      Height          =   600
      Left            =   1755
      TabIndex        =   8
      Top             =   630
      Width           =   4065
      Begin VB.OptionButton optTipoOperacao 
         Caption         =   "Desconto"
         Height          =   195
         Index           =   2
         Left            =   2610
         TabIndex        =   11
         Top             =   270
         Width           =   1140
      End
      Begin VB.OptionButton optTipoOperacao 
         Caption         =   "Cobrança"
         Height          =   195
         Index           =   1
         Left            =   1395
         TabIndex        =   10
         Top             =   270
         Value           =   -1  'True
         Width           =   1050
      End
      Begin VB.OptionButton optTipoOperacao 
         Caption         =   "Caução"
         Height          =   195
         Index           =   0
         Left            =   315
         TabIndex        =   9
         Top             =   270
         Width           =   960
      End
   End
   Begin VB.ComboBox cmbAgencia 
      Height          =   315
      Left            =   2430
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   225
      Width           =   1860
   End
   Begin VB.ComboBox cmbBanco 
      Height          =   315
      Left            =   225
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   225
      Width           =   2040
   End
   Begin VB.ListBox lstBordero 
      Height          =   2205
      Left            =   3105
      TabIndex        =   16
      Top             =   1575
      Width           =   1230
   End
   Begin VB.ListBox lstNumeroTitulo 
      Height          =   2205
      Left            =   225
      TabIndex        =   12
      Top             =   1575
      Width           =   1230
   End
   Begin MSMask.MaskEdBox mskData 
      Height          =   330
      Left            =   240
      TabIndex        =   7
      Top             =   900
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   582
      _Version        =   327680
      MaxLength       =   8
      Format          =   "ddddd"
      Mask            =   "##/##/##"
      PromptChar      =   " "
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Data de Remessa"
      Height          =   195
      Left            =   225
      TabIndex        =   6
      Top             =   675
      Width           =   1275
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Agência"
      Height          =   195
      Left            =   2460
      TabIndex        =   2
      Top             =   0
      Width           =   585
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Banco"
      Height          =   195
      Left            =   225
      TabIndex        =   0
      Top             =   45
      Width           =   465
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Borderô"
      Height          =   195
      Left            =   4455
      TabIndex        =   4
      Top             =   0
      Width           =   555
   End
End
Attribute VB_Name = "frmBordero"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'************************************************
' Este código foi retirado da Seriallink.com    *
' www.seriallink.com                            *
' UESLEI R. VALENTINI (ueslei@seriallink.com)   *
' Última Revisão: 25/08/99                      *
'************************************************
'+---------------------------------------------------------+
'| Projeto:    CPR - Contas a Pagar e a Receber            |
'| Autor:      Adilson da Silva Lima                       |
'| Data:       17/07/1997                                  |
'+---------------------------------------------------------+
'| Descrição:  Remessa de títulos para cobrança  bancária  |
'|             (borderô)                                   |
'+---------------------------------------------------------+

Private Sub cmbAgencia_Click()
   If cmbBanco.ListIndex = -1 Then
      MsgBox "Selecione um banco", 16, "Erro!"
      cmbBanco.SetFocus
      Exit Sub
   End If
   cmbBordero.Clear
   Dim Agencia As String
   Selecao = "select * from Agencia where IdBanco="
   Selecao = Selecao & cmbBanco.ItemData(cmbBanco.ListIndex)
   Selecao = Selecao & " and Nome=" & """"
   Selecao = Selecao & cmbAgencia.List(cmbAgencia.ListIndex) & """"
   Set Tabela = Banco.OpenRecordset(Selecao, dbOpenSnapshot)
   Agencia = Tabela("IdAgencia")
   Tabela.Close
   Set Consulta = Banco.QueryDefs("SelBorderos")
   Consulta("parIdBanco") = cmbBanco.ItemData(cmbBanco.ListIndex)
   Consulta("parIdAgencia") = Agencia
   
   Set Tabela = Consulta.OpenRecordset()
   Do Until Tabela.EOF
      cmbBordero.AddItem Tabela("NumeroBordero")
      Tabela.MoveNext
   Loop
   Tabela.Close
End Sub


Private Sub cmbBanco_Click()
   cmbAgencia.Clear
   cmbBordero.Clear
   mskData = Format("  /  /  ", "dd-mmm-yy")
   lstBordero.Clear
   Set Consulta = Banco.QueryDefs("SelAgencias")
   Consulta("parIdBanco") = cmbBanco.ItemData(cmbBanco.ListIndex)
   Set Tabela = Consulta.OpenRecordset()
   Do Until Tabela.EOF
      cmbAgencia.AddItem Tabela("Nome")
      cmbAgencia.ItemData(cmbAgencia.NewIndex) = Tabela("IdAgencia")
      Tabela.MoveNext
   Loop
   Tabela.Close

End Sub


Private Sub cmbBordero_Click()
   If cmbBanco.ListIndex = -1 Then
      MsgBox "Selecione um banco", 16, "Erro!"
      cmbBanco.SetFocus
      Exit Sub
   End If
   If cmbAgencia.ListIndex = -1 Then
      MsgBox "Selecione uma agência", 16, "Erro!"
      cmbAgencia.SetFocus
      Exit Sub
   End If
   lstBordero.Clear
   Dim Agencia As String
   Selecao = "select * from Agencia where IdBanco="
   Selecao = Selecao & cmbBanco.ItemData(cmbBanco.ListIndex)
   Selecao = Selecao & " and Nome=" & """"
   Selecao = Selecao & cmbAgencia.List(cmbAgencia.ListIndex) & """"
   Set Tabela = Banco.OpenRecordset(Selecao, dbOpenSnapshot)
   Agencia = Tabela("IdAgencia")
   Tabela.Close
   lstBordero.Clear
   Set Consulta = Banco.QueryDefs("SelRemessaBordero")
   Consulta("parIdBanco") = cmbBanco.ItemData(cmbBanco.ListIndex)
   Consulta("parIdAgencia") = Agencia
   Consulta("parNumeroBordero") = cmbBordero.List(cmbBordero.ListIndex)
  
   Set Tabela = Consulta.OpenRecordset()
   mskData = Tabela("DataRemessa")
   optTipoOperacao(Tabela("TipoOperacao")).Value = True
   Do Until Tabela.EOF
      lstBordero.AddItem Tabela("NumeroTitulo")
      Tabela.MoveNext
   Loop
   Tabela.Close
   '
End Sub


Private Sub cmdAdicao_Click()
   If lstNumeroTitulo.ListIndex = -1 Then
      MsgBox "Selecione um título...", 16, "Adicionar"
      lstNumeroTitulo.SetFocus
      Exit Sub
   End If
   lstBordero.AddItem lstNumeroTitulo.List(lstNumeroTitulo.ListIndex)
   lstNumeroTitulo.RemoveItem lstNumeroTitulo.ListIndex
End Sub

Private Sub cmdGravacao_Click()
   If Not IsDate(Trim$(mskData)) Then
      MsgBox "Data inválida", 16, "Erro"
      mskData.SetFocus
      Exit Sub
   End If
   If cmbBanco.ListIndex = -1 Then
      MsgBox "Selecione um banco", 16, "Erro!"
      cmbBanco.SetFocus
      Exit Sub
   End If
   If cmbAgencia.ListIndex = -1 Then
      MsgBox "Selecione uma agência", 16, "Erro!"
      cmbAgencia.SetFocus
      Exit Sub
   End If
   If Trim$(cmbBordero.Text) = "" Then
      MsgBox "Informe o número do borderô", 16, "Erro!"
      cmbBordero.SetFocus
      Exit Sub
   End If
   
   On Error GoTo ErrGravaBordero
   Dim Agencia As String
   Selecao = "select * from Agencia where IdBanco="
   Selecao = Selecao & cmbBanco.ItemData(cmbBanco.ListIndex)
   Selecao = Selecao & " and Nome=" & """"
   Selecao = Selecao & cmbAgencia.List(cmbAgencia.ListIndex) & """"
   Set Tabela = Banco.OpenRecordset(Selecao, dbOpenSnapshot)
   Agencia = Tabela("IdAgencia")
   Tabela.Close
   Area.BeginTrans
   Set Consulta = Banco.QueryDefs("DelRemessa")
   Consulta("parIdBanco") = cmbBanco.ItemData(cmbBanco.ListIndex)
   Consulta("parIdAgencia") = Agencia
   Consulta("parNumeroBordero") = cmbBordero.Text
   Consulta.Execute
   If Trim$(cmbBordero.Text) = "" Then
      Area.CommitTrans
      Exit Sub
   End If
   
   For contador = 0 To lstBordero.ListCount - 1
      Set Consulta = Banco.QueryDefs("InsRemessa")
      Consulta("parIdBanco") = cmbBanco.ItemData(cmbBanco.ListIndex)
      Consulta("parIdAgencia") = Agencia
      Consulta("parNumeroTitulo") = lstBordero.List(contador)
      Consulta("parNumeroBordero") = cmbBordero.Text
      Consulta("parDataRemessa") = mskData
      For i = 0 To 2
         If optTipoOperacao(i).Value = True Then
            Consulta("parTipoOperacao") = i
         End If
      Next i

      Consulta.Execute
   Next contador
   Area.CommitTrans
   cmbBanco.SetFocus
   Exit Sub
   
ErrGravaBordero:
   MsgBox "Erro na gravação do borderô", 16, "Gravação"
   Area.Rollback
   Exit Sub
End Sub

Private Sub cmdRemocao_Click()
   If lstBordero.ListIndex = -1 Then
      MsgBox "Selecione um título...", 16, "Adicionar"
      lstBordero.SetFocus
      Exit Sub
   End If
   lstNumeroTitulo.AddItem lstBordero.List(lstBordero.ListIndex)
   lstBordero.RemoveItem lstBordero.ListIndex
End Sub


Private Sub cmdRemocaoTotal_Click()
   For contador = 0 To lstBordero.ListCount - 1
      lstNumeroTitulo.AddItem lstBordero.List(contador)
      lstBordero.RemoveItem contador
   Next
End Sub

Private Sub cmdSaida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   ComboBanco cmbBanco
   TituloSemBordero
End Sub


Private Sub TituloSemBordero()
   lstNumeroTitulo.Clear
   Set Tabela = Banco.OpenRecordset("SelTituloSemBordero", dbOpenSnapshot)
   Do Until Tabela.EOF
      lstNumeroTitulo.AddItem Tabela("NumeroTitulo")
      Tabela.MoveNext
   Loop
   Tabela.Close
End Sub
