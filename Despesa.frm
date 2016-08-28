VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmDespesa 
   Caption         =   "Previsão de Despesas"
   ClientHeight    =   1965
   ClientLeft      =   465
   ClientTop       =   1665
   ClientWidth     =   7005
   LinkTopic       =   "Form12"
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1965
   ScaleWidth      =   7005
   Begin VB.CommandButton cmdSaida 
      Caption         =   "&Retornar"
      Height          =   420
      Left            =   5580
      TabIndex        =   14
      Top             =   1440
      Width           =   1230
   End
   Begin VB.CommandButton cmdExclusao 
      Caption         =   "&Exclusão"
      Height          =   420
      Left            =   5625
      TabIndex        =   13
      Top             =   810
      Width           =   1230
   End
   Begin VB.CommandButton cmdInclusao 
      Caption         =   "&Incluir"
      Height          =   420
      Left            =   5625
      TabIndex        =   12
      Top             =   225
      Width           =   1230
   End
   Begin VB.TextBox txtDescricao 
      Height          =   285
      Left            =   135
      TabIndex        =   11
      Top             =   1485
      Width           =   5325
   End
   Begin VB.TextBox txtIdLancamento 
      Height          =   285
      Left            =   1350
      TabIndex        =   3
      Top             =   225
      Width           =   1050
   End
   Begin VB.TextBox txtIDDespesa 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   225
      Width           =   1050
   End
   Begin MSMask.MaskEdBox mskPrevisaoPagamento 
      Height          =   330
      Left            =   1350
      TabIndex        =   7
      Top             =   855
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   582
      _Version        =   327680
      MaxLength       =   8
      Format          =   "dd-mm-yy"
      Mask            =   "99/99/99"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox mskData 
      Height          =   330
      Left            =   120
      TabIndex        =   5
      Top             =   855
      Width           =   1050
      _ExtentX        =   1852
      _ExtentY        =   582
      _Version        =   327680
      MaxLength       =   8
      Format          =   "dd-mm-yy"
      Mask            =   "99/99/99"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox mskValor 
      Height          =   330
      Left            =   2640
      TabIndex        =   9
      Top             =   840
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   582
      _Version        =   327680
      Format          =   "#,##0.00;(#,##0.00)"
      PromptChar      =   "_"
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Descrição"
      Height          =   195
      Left            =   135
      TabIndex        =   10
      Top             =   1260
      Width           =   720
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Valor Previsto"
      Height          =   195
      Left            =   2610
      TabIndex        =   8
      Top             =   630
      Width           =   975
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Previsão"
      Height          =   195
      Left            =   1350
      TabIndex        =   6
      Top             =   630
      Width           =   615
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Data"
      Height          =   195
      Left            =   135
      TabIndex        =   4
      Top             =   630
      Width           =   345
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Lançamento"
      Height          =   195
      Left            =   1350
      TabIndex        =   2
      Top             =   45
      Width           =   885
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Número"
      Height          =   195
      Left            =   135
      TabIndex        =   0
      Top             =   45
      Width           =   555
   End
End
Attribute VB_Name = "frmDespesa"
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
'+-----------------------------------------------------+
'| Projeto:    CPR - Contas a Pagar e a Receber        |
'| Autor:      Adilson da Silva Lima                   |
'| Data:       17/07/1997                              |
'+-----------------------------------------------------+
'| Descrição:  Formulário para cadastro de previsão  de|
'|             despesas                                |
'+-----------------------------------------------------+

Private Sub cmdExclusao_Click()
   If MsgBox("Confirme a exclusão da despesa", 36, "Exclusão") <> 6 Then
      Exit Sub
   End If
   On Error GoTo ErrExcluiDespesa
   ErroPar = False
   MensErro = "Falha na abertura da consulta"
   Set Consulta = Banco.QueryDefs("DelDespesa")
   ErroPar = True
   
   Area.BeginTrans
   MensErro = "Falha na alimentação dos parâmetros"

   Consulta("parIdDespesa") = txtIDDespesa
   
   MensErro = "Falha na gravação dos dados"

   Consulta.Execute
   Area.CommitTrans
   
   txtIdLancamento = ""
   mskPrevisaoPagamento = "  /  /  "
   mskData = "  /  /  "
   txtDescricao = ""
   mskValor = 0

   Exit Sub

ErrExcluiDespesa:
   MsgBox MensErro, 16, "Exclusão"
   If ErroPar Then Area.Rollback
   Exit Sub
End Sub

Private Sub cmdInclusao_Click()
   If Not ConsisteData(mskData) Then mskData.SetFocus
   If Not ConsisteData(mskPrevisaoPagamento) Then mskPrevisaoPagamento.SetFocus
   If Trim(txtDescricao) = "" Then
      MsgBox "Informe a descrição", 16, "Inclusão"
      txtDescricao.SetFocus
      Exit Sub
   End If
   If Val(mskValor) = 0 Then
      MsgBox "Informe o valor", 16, "Inclusão"
      mskValor.SetFocus
      Exit Sub
   End If

   On Error GoTo ErrIncluiDespesa
   ErroPar = False
   MensErro = "Falha na abertura da consulta"
   Set Consulta = Banco.QueryDefs("InsDespesa")
   ErroPar = True
   
   Area.BeginTrans
   MensErro = "Falha na alimentação dos parâmetros"

   Consulta("parIdLancamento") = txtIdLancamento
   Consulta("parIdDespesa") = txtIDDespesa
   Consulta("parPrevisaoPagamento") = mskPrevisaoPagamento
   Consulta("parDataLancamento") = mskData
   Consulta("parValorPrevisto") = mskValor
   Consulta("parDescricao") = txtDescricao
   
   MensErro = "Falha na gravação dos dados"

   Consulta.Execute
   Area.CommitTrans
   
   txtIdLancamento = ""
   mskPrevisaoPagamento = "  /  /  "
   mskData = "  /  /  "
   txtDescricao = ""
   mskValor = 0

   Exit Sub

ErrIncluiDespesa:
   MsgBox MensErro, 16, "Inclusão"
   If ErroPar Then Area.Rollback
   Exit Sub
End Sub

Private Sub cmdSaida_Click()
   Unload Me
End Sub

Private Sub Form_Load()

End Sub

Private Sub txtIDDespesa_LostFocus()
   If Trim$(txtIDDespesa) = "" Then txtIDDespesa = 0
   Selecao = "SELECT * from Despesa where IdDespesa="
   Selecao = Selecao & txtIDDespesa
   Set Tabela = Banco.OpenRecordset(Selecao, dbOpenSnapshot)
   
   If Tabela.RecordCount = 0 Then
      MsgBox "Despesa inexistente...", 16, "Erro"
      Exit Sub
   End If
   mskData = Tabela("DataLancamento")
   txtIdLancamento = Tabela("IdLancamento")
   mskPrevisaoPagamento = Tabela("PrevisaoPagamento")
   txtDescricao = Tabela("Descricao")
   mskValor = Tabela("ValorPrevisto")
   Tabela.Close

End Sub


Private Sub txtIdLancamento_LostFocus()
   If Trim$(txtIdLancamento) = "" Then txtIdLancamento = 0
   Selecao = "SELECT * from Lancamento where IdLancamento="
   Selecao = Selecao & txtIdLancamento
   Set Tabela = Banco.OpenRecordset(Selecao, dbOpenSnapshot)
   If Tabela.RecordCount = 0 Then Exit Sub
   txtDescricao = Tabela("Descricao")
   mskValor = Tabela("Valor")
   mskData = Tabela("Data")
 
   Tabela.Close

End Sub
