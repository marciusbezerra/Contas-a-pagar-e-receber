VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmRecebimento 
   Caption         =   "Recebimentos"
   ClientHeight    =   3870
   ClientLeft      =   210
   ClientTop       =   1425
   ClientWidth     =   6735
   LinkTopic       =   "Form6"
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3870
   ScaleWidth      =   6735
   Begin VB.CommandButton cmdSaida 
      Caption         =   "&Retornar"
      Height          =   420
      Left            =   5220
      TabIndex        =   29
      Top             =   3105
      Width           =   1230
   End
   Begin VB.CommandButton cmdExclusao 
      Caption         =   "&Excluir"
      Height          =   420
      Left            =   5175
      TabIndex        =   28
      Top             =   1710
      Width           =   1230
   End
   Begin VB.CommandButton cmdAlteracao 
      Caption         =   "&Alterar"
      Height          =   420
      Left            =   5175
      TabIndex        =   27
      Top             =   990
      Width           =   1230
   End
   Begin VB.CommandButton cmdInclusao 
      Caption         =   "&Incluir"
      Height          =   420
      Left            =   5175
      TabIndex        =   26
      Top             =   360
      Width           =   1230
   End
   Begin VB.ComboBox cmbCodinome 
      Height          =   315
      Left            =   180
      TabIndex        =   1
      Top             =   360
      Width           =   2490
   End
   Begin VB.TextBox txtNumeroTitulo 
      Height          =   285
      Left            =   180
      TabIndex        =   5
      Top             =   1035
      Width           =   1230
   End
   Begin VB.Frame Frame1 
      Caption         =   "Tipo"
      Height          =   915
      Left            =   1845
      TabIndex        =   6
      Top             =   810
      Width           =   1725
      Begin VB.OptionButton optTipo 
         Caption         =   "Duplicata"
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   7
         Top             =   225
         Value           =   -1  'True
         Width           =   1230
      End
      Begin VB.OptionButton optTipo 
         Caption         =   "Nota Promissória"
         Height          =   195
         Index           =   1
         Left            =   135
         TabIndex        =   8
         Top             =   450
         Width           =   1500
      End
      Begin VB.OptionButton optTipo 
         Caption         =   "Outros"
         Height          =   195
         Index           =   2
         Left            =   135
         TabIndex        =   9
         Top             =   675
         Width           =   1230
      End
   End
   Begin MSMask.MaskEdBox mskCPF_CGC 
      Height          =   330
      Left            =   2925
      TabIndex        =   3
      Top             =   360
      Width           =   1860
      _ExtentX        =   3281
      _ExtentY        =   582
      _Version        =   327680
      PromptInclude   =   0   'False
      MaxLength       =   18
      Mask            =   "##.###.###/####-##"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox mskValor 
      Height          =   330
      Left            =   180
      TabIndex        =   17
      Top             =   2700
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   582
      _Version        =   327680
      Format          =   "#,##0.00;(#,##0.00)"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox mskDesconto 
      Height          =   330
      Left            =   180
      TabIndex        =   21
      Top             =   3300
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   582
      _Version        =   327680
      Format          =   "#,##0.00;(#,##0.00)"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox mskAcrescimo 
      Height          =   330
      Left            =   3420
      TabIndex        =   25
      Top             =   3300
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   582
      _Version        =   327680
      Format          =   "#,##0.00;(#,##0.00)"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox mskLancamento 
      Height          =   330
      Left            =   1920
      TabIndex        =   13
      Top             =   2040
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   582
      _Version        =   327680
      MaxLength       =   8
      Format          =   "ddddd"
      Mask            =   "##/##/##"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox mskVencimento 
      Height          =   330
      Left            =   3540
      TabIndex        =   15
      Top             =   2040
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   582
      _Version        =   327680
      MaxLength       =   8
      Format          =   "ddddd"
      Mask            =   "##/##/##"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox mskPagamento 
      Height          =   330
      Left            =   1920
      TabIndex        =   19
      Top             =   2700
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   582
      _Version        =   327680
      MaxLength       =   8
      Format          =   "ddddd"
      Mask            =   "##/##/##"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox mskValidadeDesconto 
      Height          =   330
      Left            =   1920
      TabIndex        =   23
      Top             =   3300
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   582
      _Version        =   327680
      MaxLength       =   8
      Format          =   "ddddd"
      Mask            =   "##/##/##"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox mskEmissao 
      Height          =   330
      Left            =   180
      TabIndex        =   11
      Top             =   2040
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   582
      _Version        =   327680
      MaxLength       =   8
      Format          =   "ddddd"
      Mask            =   "##/##/##"
      PromptChar      =   " "
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Clientes Cadastrados"
      Height          =   195
      Left            =   180
      TabIndex        =   0
      Top             =   135
      Width           =   1485
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "CPF/CGC"
      Height          =   195
      Left            =   2925
      TabIndex        =   2
      Top             =   135
      Width           =   705
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Número do Título"
      Height          =   195
      Left            =   180
      TabIndex        =   4
      Top             =   810
      Width           =   1245
   End
   Begin VB.Label Label4 
      Caption         =   "Emissão"
      Height          =   195
      Left            =   180
      TabIndex        =   10
      Top             =   1845
      Width           =   1230
   End
   Begin VB.Label Label5 
      Caption         =   "Lançamento"
      Height          =   195
      Left            =   1935
      TabIndex        =   12
      Top             =   1845
      Width           =   1230
   End
   Begin VB.Label Label6 
      Caption         =   "Vencimento"
      Height          =   195
      Left            =   3510
      TabIndex        =   14
      Top             =   1845
      Width           =   1230
   End
   Begin VB.Label Label7 
      Caption         =   "Valor"
      Height          =   195
      Left            =   180
      TabIndex        =   16
      Top             =   2475
      Width           =   1230
   End
   Begin VB.Label Label8 
      Caption         =   "Pagamento"
      Height          =   195
      Left            =   1935
      TabIndex        =   18
      Top             =   2475
      Width           =   1230
   End
   Begin VB.Label Label9 
      Caption         =   "Desconto"
      Height          =   195
      Left            =   180
      TabIndex        =   20
      Top             =   3105
      Width           =   1230
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "Validade Desconto"
      Height          =   195
      Left            =   1935
      TabIndex        =   22
      Top             =   3105
      Width           =   1350
   End
   Begin VB.Label Label11 
      Caption         =   "Acréscimo"
      Height          =   195
      Left            =   3420
      TabIndex        =   24
      Top             =   3105
      Width           =   1230
   End
End
Attribute VB_Name = "frmRecebimento"
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
'+----------------------------------------------------------+
'| Projeto:    CPR - Contas a Pagar e a Receber             |
'| Autor:      Adilson da Silva Lima                        |
'| Data:       17/07/1997                                   |
'+----------------------------------------------------------+
'| Descrição:  Formulário para cadastro de títulos a receber|
'+----------------------------------------------------------+

Private Sub cmbCodinome_Click()
   Dim Selecao As String
   Selecao = "select * from Cliente where Codinome=" & """"
   Selecao = Selecao & cmbCodinome.Text & """"

   Set Tabela = Banco.OpenRecordset(Selecao, dbOpenSnapshot)

   
   If Tabela.RecordCount = 0 Then Exit Sub
   mskCPF_CGC = Tabela("CPF_CGC")
   Tabela.Close
End Sub

Private Sub cmdInclusao_Click()
   Set Consulta = Banco.QueryDefs("SelCliente")
   Consulta("parCPF_CGC") = mskCPF_CGC
   Set Tabela = Consulta.OpenRecordset()
   If Tabela.RecordCount = 0 Then
      MsgBox "CPF/CGC não cadastrado", 16, "Inclusão"
      Tabela.Close
      mdkCPF_CGC.SetFocus
      Exit Sub
   End If
   Tabela.Close
   Consulta.Close
   If txtNumeroTitulo = "" Then
      MsgBox "Informe o número do título", 16, "Inclusão"
      txtNumeroTitulo.SetFocus
      Exit Sub
   End If
   
   Selecao = "SELECT * from Recebimento where NumeroTitulo=" & """"
   Selecao = Selecao & txtNumeroTitulo & """"
   Set Tabela = Banco.OpenRecordset(Selecao, dbOpenSnapshot)
   If Tabela.RecordCount > 0 Then
      MsgBox "Título já cadastrado", 16, "Inclusão"
      Tabela.Close
      txtNumeroTitulo.SetFocus
      Exit Sub
   End If
   Tabela.Close
   If mskEmissao = "  /  /  " Then
      MsgBox "Informe a data de emissão", 16, "Inclusão"
      mskEmissao.SetFocus
      Exit Sub
   End If
   If mskVencimento = "  /  /  " Then
      MsgBox "Informe a data de vencimento", 16, "Inclusão"
      mskVencimento.SetFocus
      Exit Sub
   End If
   If Not ConsisteData(mskEmissao) Then mskmskEmissao.SetFocus
   If Not ConsisteData(mskVencimento) Then mskVencimento.SetFocus
   If Not mskLancamento = "  /  /  " Then
      If Not ConsisteData(mskLancamento) Then mskLancamento.SetFocus
      Exit Sub
   End If
   If Not mskPagamento = "  /  /  " Then
      If Not ConsisteData(mskPagamento) Then mskPagamento.SetFocus
      Exit Sub
   End If
   If Not mskValidadeDesconto = "  /  /  " Then
      If Not ConsisteData(mskValidadeDesconto) Then mskValidadeDesconto.SetFocus
      Exit Sub
   End If
   If Val(mskValor) = 0 Then
      MsgBox "Informe o valor do título", 16, "Inclusão"
      mskValor.SetFocus
      Exit Sub
   End If
   
   On Error GoTo ErrIncluiRecebimento
   ErroPar = False
   MensErro = "Falha na abertura da consulta"
   
   Set Consulta = Banco.QueryDefs("InsRecebimento")
   ErroPar = True
   
   Area.BeginTrans
   MensErro = "Falha na alimentação dos parâmetros"
   
   Consulta("parNumeroTitulo") = txtNumeroTitulo
   Consulta("parCPF_CGC") = mskCPF_CGC
   Consulta("parValor") = mskValor
   Consulta("parEmissao") = mskEmissao
   Consulta("parVencimento") = mskVencimento
   Consulta("parLancamento") = mskLancamento
   Consulta("parPagamento") = mskPagamento
   Consulta("parDesconto") = mskDesconto
   Consulta("parAcrescimo") = mskAcrescimo
   Consulta("parValidadeDesconto") = mskValidadeDesconto
   For contador = 0 To 2
      If optTipo(contador).Value = True Then
         Consulta("parTipo") = contador
      End If
   Next

   MensErro = "Falha na gravação dos dados"

   Consulta.Execute
   Area.CommitTrans
   
   txtNumeroTitulo = ""
   mskCPF_CGC = ""
   mskValor = 0
   mskEmissao = "  /  /  "
   mskVencimento = "  /  /  "
   mskLancamento = "  /  /  "
   mskPagamento = "  /  /  "
   mskDesconto = ""
   mskAcrescimo = ""
   mskValidadeDesconto = "  /  /  "

   Exit Sub

ErrIncluiRecebimento:
   MsgBox MensErro, 16, "Inclusão"
   If ErroPar Then Area.Rollback
   Exit Sub
End Sub

Private Sub cmdSaida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   ComboCliente cmbCodinome
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
   ' Caso o CPF/CGC não exista, o recordset é fechado e a rotina interrompida
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


Private Sub txtNumeroTitulo_LostFocus()
   Selecao = "SELECT * from Recebimento where NumeroTitulo=" & """"
   Selecao = Selecao & txtNumeroTitulo & """"
   Selecao = Selecao & " and CPF_CGC="
   Selecao = Selecao & mskCPF_CGC

   Set Tabela = Banco.OpenRecordset(Selecao, dbOpenSnapshot)
   
   If Tabela.RecordCount = 0 Then Exit Sub
   optTipo(Tabela("Tipo")).Value = True
   mskEmissao = Tabela("Emissao")
   If Not IsNull(Tabela("Lancamento")) Then mskLancamento = Tabela("Lancamento")
   mskVencimento = Tabela("Vencimento")
   If Not IsNull(Tabela("Pagamento")) Then mskPagamento = Tabela("Pagamento")
   If Not IsNull(Tabela("ValidadeDesconto")) Then mskValidadeDesconto = Tabela("ValidadeDesconto")
   mskValor = Tabela("Valor")
   If Not IsNull(Tabela("Desconto")) Then mskDesconto = Tabela("Desconto")
   If Not IsNull(Tabela("Acrescimo")) Then mskAcrescimo = Tabela("Acrescimo")
   Tabela.Close

End Sub
