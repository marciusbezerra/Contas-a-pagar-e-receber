VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmChequePreDatado 
   Caption         =   "Cheques Pré-Datados"
   ClientHeight    =   4530
   ClientLeft      =   180
   ClientTop       =   1500
   ClientWidth     =   7485
   LinkTopic       =   "Form7"
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4530
   ScaleWidth      =   7485
   Begin VB.CommandButton cmdInclusao 
      Caption         =   "&Incluir"
      Height          =   420
      Left            =   5895
      TabIndex        =   23
      Top             =   315
      Width           =   1230
   End
   Begin VB.CommandButton cmdAlteracao 
      Caption         =   "&Alterar"
      Height          =   420
      Left            =   5895
      TabIndex        =   24
      Top             =   945
      Width           =   1230
   End
   Begin VB.CommandButton cmdExclusao 
      Caption         =   "&Excluir"
      Height          =   420
      Left            =   5895
      TabIndex        =   25
      Top             =   1665
      Width           =   1230
   End
   Begin VB.CommandButton cmdSaida 
      Caption         =   "&Retornar"
      Height          =   420
      Left            =   5940
      TabIndex        =   26
      Top             =   3285
      Width           =   1230
   End
   Begin VB.ComboBox cmbCodinome 
      Height          =   315
      Left            =   225
      TabIndex        =   1
      Top             =   270
      Width           =   2490
   End
   Begin VB.Frame Frame1 
      Caption         =   "Conta Corrente"
      Height          =   2805
      Left            =   2880
      TabIndex        =   14
      Top             =   900
      Width           =   2760
      Begin VB.ComboBox cmbContaCorrente 
         Height          =   315
         Left            =   135
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   2070
         Width           =   2490
      End
      Begin VB.ComboBox cmbAgencia 
         Height          =   315
         Left            =   135
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   1260
         Width           =   2490
      End
      Begin VB.ComboBox cmbBanco 
         Height          =   315
         Left            =   135
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   495
         Width           =   2490
      End
      Begin VB.Label Label6 
         Caption         =   "Conta"
         Height          =   195
         Left            =   135
         TabIndex        =   19
         Top             =   1845
         Width           =   1230
      End
      Begin VB.Label Label4 
         Caption         =   "Agencia"
         Height          =   195
         Left            =   135
         TabIndex        =   17
         Top             =   1035
         Width           =   1230
      End
      Begin VB.Label Label1 
         Caption         =   "Banco"
         Height          =   195
         Left            =   135
         TabIndex        =   15
         Top             =   270
         Width           =   1230
      End
   End
   Begin VB.TextBox txtDescricao 
      Height          =   285
      Left            =   225
      TabIndex        =   22
      Top             =   4095
      Width           =   5460
   End
   Begin VB.TextBox txtNumeroCheque 
      Height          =   285
      Left            =   225
      TabIndex        =   5
      Top             =   945
      Width           =   1410
   End
   Begin MSMask.MaskEdBox mskCPF_CGC 
      Height          =   330
      Left            =   2970
      TabIndex        =   3
      Top             =   270
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   582
      _Version        =   327680
      PromptInclude   =   0   'False
      MaxLength       =   18
      Mask            =   "##.###.###/####-##"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox mskEmissao 
      Height          =   330
      Left            =   240
      TabIndex        =   7
      Top             =   1500
      Width           =   1050
      _ExtentX        =   1852
      _ExtentY        =   582
      _Version        =   327680
      MaxLength       =   8
      Format          =   "dd-mm-yy"
      Mask            =   "99/99/99"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox mskLancamento 
      Height          =   330
      Left            =   240
      TabIndex        =   9
      Top             =   2160
      Width           =   1050
      _ExtentX        =   1852
      _ExtentY        =   582
      _Version        =   327680
      MaxLength       =   8
      Format          =   "dd-mm-yy"
      Mask            =   "99/99/99"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox mskDeposito 
      Height          =   330
      Left            =   240
      TabIndex        =   11
      Top             =   2760
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
      Left            =   240
      TabIndex        =   13
      Top             =   3420
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   582
      _Version        =   327680
      Format          =   "#,##0.00;(#,##0.00)"
      PromptChar      =   "_"
   End
   Begin VB.Label Label5 
      Caption         =   "Valor"
      Height          =   195
      Left            =   225
      TabIndex        =   12
      Top             =   3195
      Width           =   1230
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "Depósito"
      Height          =   195
      Left            =   225
      TabIndex        =   10
      Top             =   2565
      Width           =   630
   End
   Begin VB.Label Label9 
      Caption         =   "Emissão"
      Height          =   195
      Left            =   225
      TabIndex        =   6
      Top             =   1305
      Width           =   1230
   End
   Begin VB.Label Label8 
      Caption         =   "Lançamento"
      Height          =   195
      Left            =   225
      TabIndex        =   8
      Top             =   1935
      Width           =   1230
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Clientes Cadastrados"
      Height          =   195
      Left            =   225
      TabIndex        =   0
      Top             =   45
      Width           =   1485
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "CPF/CGC"
      Height          =   195
      Left            =   2970
      TabIndex        =   2
      Top             =   45
      Width           =   705
   End
   Begin VB.Label Label11 
      Caption         =   "Descrição"
      Height          =   195
      Left            =   225
      TabIndex        =   21
      Top             =   3870
      Width           =   1230
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Número do Cheque"
      Height          =   195
      Left            =   225
      TabIndex        =   4
      Top             =   720
      Width           =   1380
   End
End
Attribute VB_Name = "frmChequePreDatado"
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
'| Descrição:  Controle de cheques pré-datados             |
'+---------------------------------------------------------+


Private Sub cmdAlteracao_Click()
   If Not ConsisteData(mskEmissao) Then mskEmissao.SetFocus
   If Not ConsisteData(mskLancamento) Then mskLancamento.SetFocus
   If Not ConsisteData(mskDeposito) Then mskDeposito.SetFocus
   If Trim(txtNumeroCheque) = "" Then
      MsgBox "Informe o número do cheque", 16, "Alteração"
      txtNumeroCheque.SetFocus
      Exit Sub
   End If
   If Trim(txtDescricao) = "" Then
      MsgBox "Informe a descrição", 16, "Alteração"
      txtDescricao.SetFocus
      Exit Sub
   End If
   If Val(mskValor) = 0 Then
      MsgBox "Informe o valor", 16, "Alteração"
      mskValor.SetFocus
      Exit Sub
   End If
   If cmbBanco.ListIndex = -1 Then
      MsgBox "Selecione o banco", 16, "Alteração"
      cmbBanco.SetFocus
      Exit Sub
   End If
   If cmbAgencia.ListIndex = -1 Then
      MsgBox "Selecione a agência", 16, "Alteração"
      cmbAgencia.SetFocus
      Exit Sub
   End If
   If cmbContaCorrente.ListIndex = -1 Then
      MsgBox "Selecione a conta corrente", 16, "Alteração"
      cmbContaCorrente.SetFocus
      Exit Sub
   End If
   
   Dim Agencia As String
   Selecao = "select * from Agencia where IdBanco="
   Selecao = Selecao & cmbBanco.ItemData(cmbBanco.ListIndex)
   Selecao = Selecao & " and Nome=" & """"
   Selecao = Selecao & cmbAgencia.List(cmbAgencia.ListIndex) & """"
   Set Tabela = Banco.OpenRecordset(Selecao, dbOpenSnapshot)
   Agencia = Tabela("IdAgencia")
   Tabela.Close
   
   On Error GoTo ErrAlteraCheque
   ErroPar = False
   MensErro = "Falha na abertura da consulta"
   Set Consulta = Banco.QueryDefs("UpdChequePreDatado")
   ErroPar = True
   
   Area.BeginTrans
   MensErro = "Falha na alimentação dos parâmetros"
   
   Consulta("parNumeroCheque") = txtNumeroCheque
   Consulta("parCPF_CGC") = mskCPF_CGC
   Consulta("parBanco") = cmbBanco.ItemData(cmbBanco.ListIndex)
   Consulta("parAgencia") = Agencia
   Consulta("parContaCorrente") = cmbContaCorrente
   Consulta("parEmissao") = mskEmissao
   Consulta("parLancamento") = mskLancamento
   Consulta("parDeposito") = mskDeposito
   Consulta("parDescricao") = txtDescricao
   Consulta("parValor") = mskValor

   MensErro = "Falha na gravação dos dados"

   Consulta.Execute
   Area.CommitTrans
   
   txtNumeroCheque = ""
   mskCPF_CGC = ""
   mskEmissao = "  /  /  "
   mskLancamento = "  /  /  "
   mskDeposito = "  /  /  "
   txtDescricao = ""
   mskValor = 0

   Exit Sub

ErrAlteraCheque:
   MsgBox MensErro, 16, "Alteração"
   If ErroPar Then Area.Rollback
   Exit Sub
End Sub

Private Sub cmdExclusao_Click()
   If Trim(mskCPF_CGC) = "" Then
      MsgBox "Informe o número de CPF/CGC", 16, "Exclusão"
      mskCPF_CGC.SetFocus
      Exit Sub
   End If
   If Trim(txtNumeroCheque) = "" Then
      MsgBox "Informe o número do cheque", 16, "Exclusão"
      txtNumeroCheque.SetFocus
      Exit Sub
   End If
    
   On Error GoTo ErrExcluiCheque
   ErroPar = False
   MensErro = "Falha na abertura da consulta"
   Set Consulta = Banco.QueryDefs("DelChequePreDatado")
   ErroPar = True
   
   Area.BeginTrans
   MensErro = "Falha na alimentação dos parâmetros"
   
   Consulta("parNumeroCheque") = txtNumeroCheque
   Consulta("parCPF_CGC") = mskCPF_CGC

   MensErro = "Falha na gravação dos dados"

   Consulta.Execute
   Area.CommitTrans
   
   txtNumeroCheque = ""
   mskCPF_CGC = ""
   mskEmissao = "  /  /  "
   mskLancamento = "  /  /  "
   mskDeposito = "  /  /  "
   txtDescricao = ""
   mskValor = 0

   Exit Sub

ErrExcluiCheque:
   MsgBox MensErro, 16, "Exclusão"
   If ErroPar Then Area.Rollback
   Exit Sub
End Sub

Private Sub cmdInclusao_Click()
   If Not ConsisteData(mskEmissao) Then mskEmissao.SetFocus
   If Not ConsisteData(mskLancamento) Then mskLancamento.SetFocus
   If Not ConsisteData(mskDeposito) Then mskDeposito.SetFocus
   If Trim(txtNumeroCheque) = "" Then
      MsgBox "Informe o númnero do cheque", 16, "Inclusão"
      txtNumeroCheque.SetFocus
      Exit Sub
   End If
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
   If cmbBanco.ListIndex = -1 Then
      MsgBox "Selecione o banco", 16, "Inclusão"
      cmbBanco.SetFocus
      Exit Sub
   End If
   If cmbAgencia.ListIndex = -1 Then
      MsgBox "Selecione a agência", 16, "Inclusão"
      cmbAgencia.SetFocus
      Exit Sub
   End If
   If cmbContaCorrente.ListIndex = -1 Then
      MsgBox "Selecione a conta corrente", 16, "Inclusão"
      cmbContaCorrente.SetFocus
      Exit Sub
   End If
   
   Dim Agencia As String
   Selecao = "select * from Agencia where IdBanco="
   Selecao = Selecao & cmbBanco.ItemData(cmbBanco.ListIndex)
   Selecao = Selecao & " and Nome=" & """"
   Selecao = Selecao & cmbAgencia.List(cmbAgencia.ListIndex) & """"
   Set Tabela = Banco.OpenRecordset(Selecao, dbOpenSnapshot)
   Agencia = Tabela("IdAgencia")
   Tabela.Close
   
'   On Error GoTo ErrIncluiCheque
   ErroPar = False
   MensErro = "Falha na abertura da consulta"
   Set Consulta = Banco.QueryDefs("InsChequePreDatado")
   ErroPar = True
   
   Area.BeginTrans
   MensErro = "Falha na alimentação dos parâmetros"

   Consulta("parNumeroCheque") = txtNumeroCheque
   Consulta("parCPF_CGC") = mskCPF_CGC
   Consulta("parBanco") = cmbBanco.ItemData(cmbBanco.ListIndex)
   Consulta("parAgencia") = Agencia
   Consulta("parContaCorrente") = cmbContaCorrente
   Consulta("parEmissao") = mskEmissao
   Consulta("parLancamento") = mskLancamento
   Consulta("parDeposito") = mskDeposito
   Consulta("parDescricao") = txtDescricao
   Consulta("parValor") = mskValor

   MensErro = "Falha na gravação dos dados"

   Consulta.Execute
   Area.CommitTrans
   
   txtNumeroCheque = ""
   mskCPF_CGC = ""
   cmbBanco.ListIndex = -1
   cmbAgencia.ListIndex = -1
   cmbContaCorrente.ListIndex = -1
   mskEmissao = "  /  /  "
   mskLancamento = "  /  /  "
   mskDeposito = "  /  /  "
   txtDescricao = ""
   mskValor = 0

   Exit Sub

ErrIncluiCheque:
   MsgBox MensErro, 16, "Inclusão"
   If ErroPar Then Area.Rollback
   Exit Sub

End Sub

Private Sub cmdSaida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   ComboCliente cmbCodinome
   ComboBanco cmbBanco
   If cmbBanco.ListCount > 0 Then
      cmbBanco.ListIndex = 0
   End If
End Sub
Private Sub cmbCodinome_Click()
   Dim Selecao As String
   Selecao = "select * from Cliente where Codinome=" & """"
   Selecao = Selecao & cmbCodinome.Text & """"

   Set Tabela = Banco.OpenRecordset(Selecao, dbOpenSnapshot)
   
   If Tabela.RecordCount = 0 Then Exit Sub
   mskCPF_CGC = Tabela("CPF_CGC")
   Tabela.Close
End Sub

Private Sub mskCPF_CGC_GotFocus()
   mskCPF_CGC.Mask = "##############"
End Sub

Private Sub mskCPF_CGC_LostFocus()
   Dim temp As String
   temp = mskCPF_CGC
   Select Case Len(mskCPF_CGC)
      Case Is = 11
         mskCPF_CGC.Mask = "###.###.###-##"
      Case Is = 14
         mskCPF_CGC.Mask = "##.###.###/####-##"
   End Select
   mskCPF_CGC = temp
   
   If Trim$(temp) = "" Then Exit Sub
   Set Consulta = Banco.QueryDefs("SelCliente")
   Consulta("parCPF_CGC") = temp
   Set Tabela = Consulta.OpenRecordset()
   If Tabela.RecordCount = 0 Then
      Tabela.Close
      Exit Sub
   End If
   For contador = 0 To cmbCodinome.ListCount - 1
      If cmbCodinome.List(contador) = Tabela("Codinome") Then
         cmbCodinome.ListIndex = contador
         Exit For
      End If
   Next
   Tabela.Close
  
End Sub



Private Sub cmbBanco_Click()
      
   cmbAgencia.Clear
   cmbContaCorrente.Clear
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

Private Sub cmbAgencia_Click()
   If cmbBanco.ListIndex = -1 Then
      MsgBox "Selecione um banco", 16, "Erro!"
      cmbBanco.SetFocus
      Exit Sub
   End If
   Dim Agencia As String
   Selecao = "select * from Agencia where IdBanco="
   Selecao = Selecao & cmbBanco.ItemData(cmbBanco.ListIndex)
   Selecao = Selecao & " and Nome=" & """"
   Selecao = Selecao & cmbAgencia.List(cmbAgencia.ListIndex) & """"
   Set Tabela = Banco.OpenRecordset(Selecao, dbOpenSnapshot)
   Agencia = Tabela("IdAgencia")
   Tabela.Close
   Set Consulta = Banco.QueryDefs("SelContasCorrentes")
   Consulta("parIdBanco") = cmbBanco.ItemData(cmbBanco.ListIndex)
   Consulta("parIdAgencia") = Agencia
   
   Set Tabela = Consulta.OpenRecordset()
   Do Until Tabela.EOF
      cmbContaCorrente.AddItem Tabela("Conta")
      Tabela.MoveNext
   Loop
   Tabela.Close
End Sub


Private Sub txtNumeroCheque_LostFocus()
   ' Nesta rotina notamos a declaração de um recordset local ("Temp")
   Dim temp As Recordset, Temp1 As Recordset
   If Trim$(txtNumeroCheque) = "" Then Exit Sub
   Set Consulta = Banco.QueryDefs("SelChequePreDatado")
   Consulta("parCPF_CGC") = mskCPF_CGC
   Consulta("parNumeroCheque") = txtNumeroCheque

   Set Temp1 = Consulta.OpenRecordset()
   
   If Temp1.RecordCount = 0 Then Exit Sub
   
   For contador = 0 To cmbBanco.ListCount - 1
      If cmbBanco.ItemData(contador) = Temp1("Banco") Then
         cmbBanco.ListIndex = contador
         Exit For
      End If
   Next
   
   mskEmissao = Temp1("Emissao")
   mskLancamento = Temp1("Lancamento")
   mskDeposito = Temp1("Deposito")
   mskValor = Temp1("Valor")
   txtDescricao = Temp1("Descricao")
      
   Selecao = "select * from Agencia where IdBanco="
   Selecao = Selecao & Temp1("Banco")
   Selecao = Selecao & " and IdAgencia=" & """"
   Selecao = Selecao & Temp1("Agencia") & """"
   Set temp = Banco.OpenRecordset(Selecao, dbOpenSnapshot)
   For contador = 0 To cmbAgencia.ListCount - 1
      If temp("Nome") = cmbAgencia.List(contador) Then
         cmbAgencia.ListIndex = contador
         Exit For
      End If
   Next
   temp.Close
   For contador = 0 To cmbContaCorrente.ListCount - 1
      If cmbContaCorrente.List(contador) = Temp1("ContaCorrente") Then
         cmbContaCorrente.ListIndex = contador
         Exit For
      End If
   Next
   Temp1.Close
End Sub
