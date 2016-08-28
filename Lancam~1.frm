VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmLancamento 
   Caption         =   "Lançamentos Contábeis"
   ClientHeight    =   3795
   ClientLeft      =   270
   ClientTop       =   915
   ClientWidth     =   7815
   LinkTopic       =   "Form11"
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3795
   ScaleWidth      =   7815
   Begin VB.TextBox txtContaCredito 
      Height          =   285
      Left            =   120
      TabIndex        =   10
      Top             =   2040
      Width           =   1470
   End
   Begin VB.TextBox txtContaDebito 
      Height          =   285
      Left            =   120
      TabIndex        =   7
      Top             =   1500
      Width           =   1470
   End
   Begin VB.CommandButton cmdInclusao 
      Caption         =   "&Incluir"
      Height          =   420
      Left            =   6435
      TabIndex        =   16
      Top             =   855
      Width           =   1230
   End
   Begin VB.CommandButton cmdExclusao 
      Caption         =   "&Excluir"
      Height          =   420
      Left            =   6435
      TabIndex        =   17
      Top             =   1440
      Width           =   1230
   End
   Begin VB.CommandButton cmdSaida 
      Caption         =   "&Retornar"
      Height          =   420
      Left            =   6435
      TabIndex        =   18
      Top             =   2610
      Width           =   1230
   End
   Begin VB.ComboBox cmbConta 
      Height          =   315
      Left            =   1755
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   900
      Width           =   4425
   End
   Begin VB.TextBox txtDescricao 
      Height          =   285
      Left            =   135
      TabIndex        =   13
      Top             =   2655
      Width           =   5190
   End
   Begin VB.TextBox txtIdLancamento 
      Height          =   285
      Left            =   135
      TabIndex        =   3
      Top             =   900
      Width           =   1185
   End
   Begin MSMask.MaskEdBox mskValor 
      Height          =   330
      Left            =   120
      TabIndex        =   15
      Top             =   3240
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   582
      _Version        =   327680
      Format          =   "#,##0.00;(#,##0.00)"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox mskData 
      Height          =   330
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   582
      _Version        =   327680
      MaxLength       =   8
      Format          =   "ddddd"
      Mask            =   "##/##/##"
      PromptChar      =   " "
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Valor"
      Height          =   195
      Left            =   135
      TabIndex        =   14
      Top             =   3015
      Width           =   360
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Descrição"
      Height          =   195
      Left            =   135
      TabIndex        =   12
      Top             =   2430
      Width           =   720
   End
   Begin VB.Label lblContaCredito 
      Height          =   195
      Left            =   1740
      TabIndex        =   11
      Top             =   2100
      Width           =   3585
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Conta Crédito"
      Height          =   195
      Left            =   135
      TabIndex        =   9
      Top             =   1845
      Width           =   960
   End
   Begin VB.Label lblContaDebito 
      Height          =   195
      Left            =   1740
      TabIndex        =   8
      Top             =   1530
      Width           =   3585
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Conta Débito"
      Height          =   195
      Left            =   135
      TabIndex        =   6
      Top             =   1260
      Width           =   930
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Contas de Lançamento"
      Height          =   195
      Left            =   1755
      TabIndex        =   4
      Top             =   675
      Width           =   1650
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Número Lançamento"
      Height          =   195
      Left            =   135
      TabIndex        =   2
      Top             =   675
      Width           =   1485
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Data de Lançamento"
      Height          =   195
      Left            =   135
      TabIndex        =   0
      Top             =   45
      Width           =   1500
   End
End
Attribute VB_Name = "frmLancamento"
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
'| Descrição:  Formulário para lançamentos de contas       |
'+---------------------------------------------------------+

Private Sub cmbConta_Click()
   If cmbConta.ListIndex = -1 Then Exit Sub
   If Trim(txtContaDebito) = "" Then
      txtContaDebito = Trim(Left(cmbConta.List(cmbConta.ListIndex), 9))
   Else
      If Trim(txtContaCredito) = "" Then
         txtContaCredito = Trim(Left(cmbConta.List(cmbConta.ListIndex), 9))
      End If
   End If
End Sub

Private Sub cmdExclusao_Click()
   If Trim(txtIdLancamento) = "" Then
      MsgBox "Informe o número do lançamento", 16, "Exclusão"
      txtIdLancamento.SetFocus
      Exit Sub
   End If
   If MsgBox("Confirme a exclusão do lançamento", 36, "Exclusão") <> 6 Then
      Exit Sub
   End If
   On Error GoTo ErrExcluiLancamento

   Set Consulta = Banco.QueryDefs("DelLancamento")
   Consulta("parIdLancamento") = txtIdLancamento
   
   Consulta.Execute
   
   txtContaCredito = ""
   txtContaDebito = ""
   lblContaCredito = ""
   lblContaDebito = ""
   txtDescricao = ""
   mskValor = 0
   cmbConta.ListIndex = -1
   txtIdLancamento.SetFocus
   Exit Sub
   
ErrExcluiLancamento:
   MsgBox MensErro, 16, "Falha na exclusão do lancamento"
   Exit Sub

End Sub

Private Sub cmdInclusao_Click()
   If Not ConsisteData(mskData) Then mskData.SetFocus
   If Trim(txtDescricao) = "" Then
      MsgBox "Informe a descrição do lançamento", 16, "Inclusão"
      txtDescricao.SetFocus
      Exit Sub
   End If
   If Trim(txtContaDebito) = Trim(txtContaCredito) Then
      MsgBox "A conta de débito é a mesma da conta de crédito...", 16, Titulo
      txtContaDebito.SetFocus
      Exit Sub
   End If
   If Val(mskValor) = 0 Then
      MsgBox "Informe o valor do lançamento", 16, "Inclusão"
      mskValor.SetFocus
      Exit Sub
   End If
   
   Set Consulta = Banco.QueryDefs("SelContaPlano")
   Consulta("parConta") = txtContaDebito
   Set Tabela = Consulta.OpenRecordset()
   If Tabela.RecordCount = 0 Then
      MsgBox "Conta de débito inexistente...", 16, "Inclusão"
      Tabela.Close
      txtContaDebito.SetFocus
      Exit Sub
   End If
   Tabela.Close
   
   Set Consulta = Banco.QueryDefs("SelContaPlano")
   Consulta("parConta") = txtContaCredito
   Set Tabela = Consulta.OpenRecordset()
   If Tabela.RecordCount = 0 Then
      MsgBox "Conta de crédito inexistente...", 16, "Inclusão"
      Tabela.Close
      txtContaCredito.SetFocus
      Exit Sub
   End If
   Tabela.Close

   On Error GoTo ErrGravaLancamento
   ErroPar = False
   MensErro = "Falha na abertura da consulta"
   Set Consulta = Banco.QueryDefs("InsLancamento")
   ErroPar = True
   
   Area.BeginTrans
   MensErro = "Falha na alimentação dos parâmetros"

   Consulta("parData") = mskData
   Consulta("parContaCredito") = txtContaCredito
   Consulta("parContaDebito") = txtContaDebito
   Consulta("parDescricao") = txtDescricao
   Consulta("parValor") = mskValor

   MensErro = "Falha na gravação dos dados"

   Consulta.Execute
   Area.CommitTrans
   
   txtContaCredito = ""
   txtContaDebito = ""
   lblContaCredito = ""
   lblContaDebito = ""
   txtDescricao = ""
   mskValor = 0
   cmbConta.ListIndex = -1

   Exit Sub

ErrGravaLancamento:
   MsgBox MensErro, 16, "Atualização de Agência"
   If ErroPar Then Area.Rollback
   Exit Sub

End Sub

Private Sub cmdSaida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Selecao = "Select * from PlanoContas where not Consolidacao"
   Set Tabela = Banco.OpenRecordset(Selecao, dbOpenSnapshot)
   Do Until Tabela.EOF
      cmbConta.AddItem Tabela("Conta") & String(9 - Len(Tabela("Conta")), " ") & "." & Tabela("Descricao")
      Tabela.MoveNext
   Loop
   Tabela.Close
End Sub

Private Sub txtIdLancamento_LostFocus()
   If Trim$(txtIdLancamento) = "" Then txtIdLancamento = 0
   ' Apesar de existir uma consulta na base de dados CPR.MDB para selecionar um _
     lançamento, preferimos executar diretamente o recordset a fim de exemplificar a _
     montagem de uma declaração SQL um pouco mais complexa
   Selecao = "SELECT L.Data, L.ContaCredito, L.ContaDebito, L.Descricao, "
   Selecao = Selecao & " L.Valor, L.IdLancamento, D.Descricao AS Debito, "
   Selecao = Selecao & " C.Descricao AS Credito FROM Lancamento AS L, "
   Selecao = Selecao & " PlanoContas As D, PlanoContas As C"
   Selecao = Selecao & " WHERE L.IdLancamento="
   Selecao = Selecao & txtIdLancamento
   Selecao = Selecao & " and L.ContaDebito=D.Conta "
   Selecao = Selecao & " and L.ContaCredito = C.Conta"

   Set Tabela = Banco.OpenRecordset(Selecao, dbOpenSnapshot)
   
   If Tabela.RecordCount = 0 Then Exit Sub
   txtContaDebito = Tabela("ContaDebito")
   lblContaDebito = Tabela("Debito")
   txtContaCredito = Tabela("ContaCredito")
   lblContaCredito = Tabela("Credito")
   txtDescricao = Tabela("Descricao")
   mskValor = Tabela("Valor")
   mskData = Tabela("Data")
   
   Tabela.Close
End Sub
