VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form frmPlanoContas 
   Caption         =   "Manutenção do Plano de Contas"
   ClientHeight    =   3630
   ClientLeft      =   870
   ClientTop       =   1560
   ClientWidth     =   6825
   LinkTopic       =   "Form7"
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3630
   ScaleWidth      =   6825
   Begin VB.CommandButton cmdImpressao 
      Caption         =   "&Impressão"
      Height          =   420
      Left            =   5400
      TabIndex        =   12
      Top             =   2400
      Width           =   1230
   End
   Begin VB.TextBox txtSuperior 
      Height          =   285
      Left            =   1680
      TabIndex        =   5
      Top             =   2520
      Width           =   1470
   End
   Begin VB.TextBox txtConta 
      Height          =   285
      Left            =   60
      TabIndex        =   3
      Top             =   2520
      Width           =   1470
   End
   Begin VB.ListBox lstPlanoContas 
      Height          =   1815
      Left            =   45
      TabIndex        =   1
      Top             =   300
      Width           =   4935
   End
   Begin VB.CommandButton cmdExclusao 
      Caption         =   "&Exclusao"
      Height          =   420
      Left            =   5400
      TabIndex        =   11
      Top             =   1710
      Width           =   1230
   End
   Begin VB.CommandButton cmdSaida 
      Caption         =   "&Retornar"
      Height          =   420
      Left            =   5400
      TabIndex        =   13
      Top             =   3060
      Width           =   1230
   End
   Begin VB.CommandButton cmdInclusao 
      Caption         =   "&Inclusao"
      Height          =   420
      Left            =   5400
      TabIndex        =   9
      Top             =   360
      Width           =   1230
   End
   Begin VB.CommandButton cmdAlteracao 
      Caption         =   "&Alteracao"
      Height          =   420
      Left            =   5400
      TabIndex        =   10
      Top             =   1035
      Width           =   1230
   End
   Begin VB.CheckBox chkConsolidacao 
      Caption         =   "Consolidação"
      Height          =   195
      Left            =   3540
      TabIndex        =   6
      Top             =   2520
      Value           =   1  'Checked
      Width           =   1530
   End
   Begin VB.TextBox txtDescricao 
      Height          =   285
      Left            =   45
      TabIndex        =   8
      Top             =   3165
      Width           =   5010
   End
   Begin Crystal.CrystalReport Relatorio 
      Left            =   5100
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "C:\Erica_VB5\CPR\plano.rpt"
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Conta Superior"
      Height          =   195
      Left            =   1680
      TabIndex        =   4
      Top             =   2280
      Width           =   1050
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Descrição"
      Height          =   195
      Left            =   45
      TabIndex        =   7
      Top             =   2940
      Width           =   720
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Número da Conta"
      Height          =   195
      Left            =   60
      TabIndex        =   2
      Top             =   2280
      Width           =   1245
   End
   Begin VB.Label Label1 
      Caption         =   "Plano de Contas"
      Height          =   195
      Left            =   45
      TabIndex        =   0
      Top             =   90
      Width           =   1230
   End
End
Attribute VB_Name = "frmPlanoContas"
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
'+--------------------------------------------------------+
'| Projeto:    CPR - Contas a Pagar e a Receber           |
'| Autor:      Adilson da Silva Lima                      |
'| Data:       17/07/1997                                 |
'+--------------------------------------------------------+
'| Descrição:  Formulário para cadastro do plano de contas|
'| OBSERVAÇÃO: No próximo volume de dessa série estaremos |
'|             apresentando uma solução mais refinada uti-|
'|             o controle TREEVIEW, onde esse mesmo plano |
'|             de contas será montado em uma    estrutura |
'|             hierárquica.                               |
'+--------------------------------------------------------+

Private Sub cmdAlteracao_Click()
   GravaContaPlano 2
End Sub

Private Sub cmdExclusao_Click()
   If lstPlanoContas.ListIndex = -1 Then
      MsgBox "Selecione a conta...", 16, "Exclusão"
      lstPlanoContas.SetFocus
      Exit Sub
   End If
   If MsgBox("Confirme a exclusão da conta", 36, "Exclusão") <> 6 Then
      Exit Sub
   End If
   On Error GoTo ErrExcluiConta

   Set Consulta = Banco.QueryDefs("DelPlanoContas")
   Consulta("parConta") = Left(lstPlanoContas.List(lstPlanoContas.ListIndex), 9)
   Consulta.Execute
   ListaPlano
   Exit Sub
   
ErrExcluiConta:
   MsgBox MensErro, 16, "Falha na exclusão da conta", "Exclusão"
   Exit Sub
End Sub

Private Sub cmdInclusao_Click()
   GravaContaPlano 1
End Sub

Private Sub cmdSaida_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    ListaPlano
End Sub

Private Sub ListaPlano()
   lstPlanoContas.Clear
   Set Tabela = Banco.OpenRecordset("SelPlanoContas")
   ' A aplicação assume que o maior comprimento para número de conta é _
     9: a função STRING retorna um caractere informado no segundo parâ- _
     metro repetido um determinado número de vezes (esse número é infor- _
     do pelo primeiro parâmetro da função)
   Do Until Tabela.EOF
      lstPlanoContas.AddItem Tabela("Conta") & String(9 - Len(Tabela("Conta")), " ") & "." & Tabela("Descricao")
      Tabela.MoveNext
   Loop
   Tabela.Close

End Sub

Private Sub lstPlanoContas_Click()

   Set Consulta = Banco.QueryDefs("SelContaPlano")
   Consulta("parConta") = Left(lstPlanoContas.List(lstPlanoContas.ListIndex), 9)
   
   Set Tabela = Consulta.OpenRecordset()
   
   txtConta = Tabela("Conta")
   txtSuperior = Tabela("Superior")
   txtDescricao = Tabela("Descricao")
   chkConsolidacao = Tabela("Consolidacao") * -1
   
   Tabela.Close


End Sub

Private Sub mskConta_Change()

End Sub
Private Sub GravaContaPlano(Operacao As Integer)
   Dim Titulo As String
   If Operacao = 1 Then
      Titulo = "Inclusão"
   Else
      Titulo = "Alteração"
   End If
   If Trim$(txtConta) = "" Then
      MsgBox "Informe o número da conta", 16, Titulo
      txtConta.SetFocus
      Exit Sub
   End If
   If Trim$(txtDescricao) = "" Then
      MsgBox "Informe a descricao da conta", 16, Titulo
      txtDescricao.SetFocus
      Exit Sub
   End If
   
   Set Consulta = Banco.QueryDefs("SelContaPlano")
   Consulta("parConta") = txtSuperior
   
   Set Tabela = Consulta.OpenRecordset()
   If Tabela.RecordCount = 0 Then
      MsgBox "Conta superior inexistente...", 16, Titulo
      Tabela.Close
      txtSuperior.SetFocus
      Exit Sub
   End If
   Tabela.Close
   
   Dim MensErro As String, ErroPar As Integer
   On Error GoTo ErrGravaContaPlano
   ErroPar = False
   MensErro = "Falha na abertura da consulta"
   If Operacao = 1 Then
      Set Consulta = Banco.QueryDefs("InsPlanoContas")
   Else
      Set Consulta = Banco.QueryDefs("UpdPlanoContas")
   End If
   ErroPar = True
   
   Area.BeginTrans
   MensErro = "Falha na alimentação dos parâmetros"
   Consulta("parConta") = txtConta
   Consulta("parDescricao") = txtDescricao
   If Trim(txtSuperior) = "" Then txtSuperior = 0
   Consulta("parSuperior") = txtSuperior
   Consulta("parConsolidacao") = chkConsolidacao * -1
   
   MensErro = "Falha na gravação dos dados"

   Consulta.Execute

   Area.CommitTrans
   
   txtConta = ""
   txtDescricao = ""
   txtSuperior = ""
   chkConsolidacao = 0
   ListaPlano
   lstPlanoContas.SetFocus
   Exit Sub

ErrGravaContaPlano:
   MsgBox MensErro, 16, "Atualização de Conta"
   If ErroPar Then Area.Rollback
   Exit Sub
End Sub

Private Sub cmdImpressao_Click()
   Relatorio.Action = 1
End Sub

