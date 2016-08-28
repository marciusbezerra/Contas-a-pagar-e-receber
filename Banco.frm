VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmBanco 
   Caption         =   "Cadastro de Bancos e Contas Correntes"
   ClientHeight    =   4455
   ClientLeft      =   240
   ClientTop       =   1470
   ClientWidth     =   7695
   LinkTopic       =   "Form10"
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4455
   ScaleWidth      =   7695
   Begin VB.CommandButton cmdSaida 
      Caption         =   "&Retornar"
      Height          =   420
      Left            =   6300
      TabIndex        =   41
      Top             =   450
      Width           =   1230
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4185
      Left            =   180
      TabIndex        =   0
      Top             =   180
      Width           =   5910
      _ExtentX        =   10425
      _ExtentY        =   7382
      _Version        =   327680
      Tab             =   1
      TabHeight       =   520
      TabCaption(0)   =   "&Bancos"
      Tab(0).ControlCount=   9
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label4"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label12"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdExclusaoBanco"
      Tab(0).Control(3).Enabled=   -1  'True
      Tab(0).Control(4)=   "cmdAlteracaoBanco"
      Tab(0).Control(4).Enabled=   -1  'True
      Tab(0).Control(5)=   "cmbInclusaobanco"
      Tab(0).Control(5).Enabled=   -1  'True
      Tab(0).Control(6)=   "txtIdBanco"
      Tab(0).Control(6).Enabled=   -1  'True
      Tab(0).Control(7)=   "txtNomeBanco"
      Tab(0).Control(7).Enabled=   -1  'True
      Tab(0).Control(8)=   "lstBanco"
      Tab(0).Control(8).Enabled=   -1  'True
      TabCaption(1)   =   "&Agências"
      Tab(1).ControlCount=   21
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label11"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label8"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label13"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label14"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label7"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Label6"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Label5"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Label3"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Label15"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "mskFAX"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "mskTelefone"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "cmbAgencia"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "cmdExclusaoAgencia"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "cmdAlteracaoAgencia"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "cmdInclusaoAgencia"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "txtInternet"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "txtRamal"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "txtNomeAgencia"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "txtGerente"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).Control(19)=   "txtNumeroAgencia"
      Tab(1).Control(19).Enabled=   0   'False
      Tab(1).Control(20)=   "txtEndereco"
      Tab(1).Control(20).Enabled=   0   'False
      TabCaption(2)   =   "&Contas"
      Tab(2).ControlCount=   10
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label1"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Label10"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Label9"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "mskLimite"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "chkTipo"
      Tab(2).Control(4).Enabled=   -1  'True
      Tab(2).Control(5)=   "txtConta"
      Tab(2).Control(5).Enabled=   -1  'True
      Tab(2).Control(6)=   "lstContaCorrente"
      Tab(2).Control(6).Enabled=   -1  'True
      Tab(2).Control(7)=   "cmdExclusaoConta"
      Tab(2).Control(7).Enabled=   -1  'True
      Tab(2).Control(8)=   "cmdAlteracaoConta"
      Tab(2).Control(8).Enabled=   -1  'True
      Tab(2).Control(9)=   "cmdInclusaoConta"
      Tab(2).Control(9).Enabled=   -1  'True
      Begin VB.TextBox txtEndereco 
         Height          =   285
         Left            =   360
         TabIndex        =   17
         Top             =   1905
         Width           =   3705
      End
      Begin VB.ListBox lstBanco 
         Height          =   840
         Left            =   -74775
         TabIndex        =   2
         Top             =   900
         Width           =   4065
      End
      Begin VB.TextBox txtNomeBanco 
         Height          =   285
         Left            =   -74775
         MaxLength       =   30
         TabIndex        =   6
         Top             =   2700
         Width           =   3840
      End
      Begin VB.TextBox txtIdBanco 
         Height          =   285
         Left            =   -74775
         MaxLength       =   4
         TabIndex        =   4
         Top             =   2070
         Width           =   825
      End
      Begin VB.CommandButton cmbInclusaobanco 
         Caption         =   "Incluir"
         Height          =   330
         Left            =   -70545
         TabIndex        =   7
         Top             =   900
         Width           =   1140
      End
      Begin VB.CommandButton cmdAlteracaoBanco 
         Caption         =   "Alterar"
         Height          =   330
         Left            =   -70545
         TabIndex        =   8
         Top             =   1395
         Width           =   1140
      End
      Begin VB.CommandButton cmdExclusaoBanco 
         Caption         =   "Excluir"
         Height          =   330
         Left            =   -70545
         TabIndex        =   9
         Top             =   1890
         Width           =   1140
      End
      Begin VB.TextBox txtNumeroAgencia 
         Height          =   285
         Left            =   360
         MaxLength       =   5
         TabIndex        =   13
         Top             =   1305
         Width           =   825
      End
      Begin VB.TextBox txtGerente 
         Height          =   285
         Left            =   360
         TabIndex        =   19
         Top             =   2490
         Width           =   3705
      End
      Begin VB.TextBox txtNomeAgencia 
         Height          =   285
         Left            =   1350
         TabIndex        =   15
         Top             =   1305
         Width           =   2715
      End
      Begin VB.TextBox txtRamal 
         Height          =   285
         Left            =   1575
         TabIndex        =   23
         Top             =   3075
         Width           =   645
      End
      Begin VB.TextBox txtInternet 
         Height          =   285
         Left            =   360
         TabIndex        =   27
         Top             =   3660
         Width           =   2220
      End
      Begin VB.CommandButton cmdInclusaoAgencia 
         Caption         =   "Incluir"
         Height          =   330
         Left            =   4365
         TabIndex        =   28
         Top             =   675
         Width           =   1140
      End
      Begin VB.CommandButton cmdAlteracaoAgencia 
         Caption         =   "Alterar"
         Height          =   330
         Left            =   4365
         TabIndex        =   29
         Top             =   1305
         Width           =   1140
      End
      Begin VB.CommandButton cmdExclusaoAgencia 
         Caption         =   "Excluir"
         Height          =   330
         Left            =   4365
         TabIndex        =   30
         Top             =   1890
         Width           =   1140
      End
      Begin VB.ComboBox cmbAgencia 
         Height          =   315
         Left            =   360
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   675
         Width           =   3885
      End
      Begin VB.CommandButton cmdInclusaoConta 
         Caption         =   "Incluir"
         Height          =   330
         Left            =   -71280
         TabIndex        =   38
         Top             =   945
         Width           =   1140
      End
      Begin VB.CommandButton cmdAlteracaoConta 
         Caption         =   "Alterar"
         Height          =   330
         Left            =   -71265
         TabIndex        =   39
         Top             =   1485
         Width           =   1140
      End
      Begin VB.CommandButton cmdExclusaoConta 
         Caption         =   "Excluir"
         Height          =   330
         Left            =   -71265
         TabIndex        =   40
         Top             =   2070
         Width           =   1140
      End
      Begin VB.ListBox lstContaCorrente 
         Height          =   840
         Left            =   -74415
         TabIndex        =   32
         Top             =   945
         Width           =   2985
      End
      Begin VB.TextBox txtConta 
         Height          =   285
         Left            =   -74415
         TabIndex        =   34
         Top             =   2115
         Width           =   2715
      End
      Begin VB.CheckBox chkTipo 
         Caption         =   "Conta de Movimento"
         Height          =   195
         Left            =   -72705
         TabIndex        =   37
         Top             =   2790
         Width           =   1770
      End
      Begin MSMask.MaskEdBox mskLimite 
         Height          =   330
         Left            =   -74415
         TabIndex        =   36
         Top             =   2745
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   582
         _Version        =   327680
         Format          =   "#,##0.00;(#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskTelefone 
         Height          =   315
         Left            =   360
         TabIndex        =   21
         ToolTipText     =   "Telefone da residência"
         Top             =   3060
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         _Version        =   327680
         PromptInclude   =   0   'False
         MaxLength       =   10
         Format          =   "c"
         Mask            =   "####-##-##"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox mskFAX 
         Height          =   315
         Left            =   2460
         TabIndex        =   25
         ToolTipText     =   "Telefone da residência"
         Top             =   3060
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         _Version        =   327680
         PromptInclude   =   0   'False
         MaxLength       =   10
         Format          =   "c"
         Mask            =   "####-##-##"
         PromptChar      =   " "
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Endereço"
         Height          =   195
         Left            =   360
         TabIndex        =   16
         Top             =   1680
         Width           =   690
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Bancos Cadastrados"
         Height          =   195
         Left            =   -74775
         TabIndex        =   1
         Top             =   675
         Width           =   1470
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Nome"
         Height          =   195
         Left            =   -74775
         TabIndex        =   5
         Top             =   2475
         Width           =   420
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Número"
         Height          =   195
         Left            =   -74775
         TabIndex        =   3
         Top             =   1845
         Width           =   555
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Agências Cadastradas"
         Height          =   195
         Left            =   360
         TabIndex        =   10
         Top             =   450
         Width           =   1590
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Número"
         Height          =   195
         Left            =   360
         TabIndex        =   12
         Top             =   1080
         Width           =   555
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Gerente"
         Height          =   195
         Left            =   360
         TabIndex        =   18
         Top             =   2265
         Width           =   570
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Nome"
         Height          =   195
         Left            =   1350
         TabIndex        =   14
         Top             =   1080
         Width           =   420
      End
      Begin VB.Label Label14 
         Caption         =   "Ramal"
         Height          =   195
         Left            =   1575
         TabIndex        =   22
         Top             =   2850
         Width           =   690
      End
      Begin VB.Label Label13 
         Caption         =   "Telefone"
         Height          =   195
         Left            =   360
         TabIndex        =   20
         Top             =   2850
         Width           =   825
      End
      Begin VB.Label Label8 
         Caption         =   "FAX"
         Height          =   195
         Left            =   2475
         TabIndex        =   24
         Top             =   2850
         Width           =   825
      End
      Begin VB.Label Label11 
         Caption         =   "Internet"
         Height          =   195
         Left            =   360
         TabIndex        =   26
         Top             =   3435
         Width           =   1230
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Contas Cadastradas"
         Height          =   195
         Left            =   -74415
         TabIndex        =   31
         Top             =   720
         Width           =   1425
      End
      Begin VB.Label Label10 
         Caption         =   "Conta"
         Height          =   195
         Left            =   -74415
         TabIndex        =   33
         Top             =   1890
         Width           =   825
      End
      Begin VB.Label Label1 
         Caption         =   "Limite"
         Height          =   195
         Left            =   -74415
         TabIndex        =   35
         Top             =   2520
         Width           =   825
      End
   End
End
Attribute VB_Name = "frmBanco"
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
'| Descrição:  Formulário para cadastro de bancos, agências|
'|             e contas correntes                          |
'+---------------------------------------------------------+

Private Sub cmbAgencia_Click()
   If lstBanco.ListIndex = -1 Then Exit Sub
   Set Consulta = Banco.QueryDefs("SelAgencia")
   Consulta("parIdBanco") = lstBanco.ItemData(lstBanco.ListIndex)
   Consulta("parIdAgencia") = cmbAgencia.ItemData(cmbAgencia.ListIndex)
   Set Tabela = Consulta.OpenRecordset()
   
   txtNumeroAgencia = Tabela("IdAgencia")
   txtNomeAgencia = Tabela("Nome")
   txtGerente = Tabela("IdAgencia")
   If Not IsNull(Tabela("Endereco")) Then txtEndereco = Tabela("Endereco")
   If Not IsNull(Tabela("Telefone")) Then mskTelefone.Text = Tabela("Telefone")
   If Not IsNull(Tabela("Ramal")) Then txtRamal = Tabela("Ramal")
   If Not IsNull(Tabela("FAX")) Then mskFAX = Tabela("FAX")
   If Not IsNull(Tabela("Internet")) Then txtInternet = Tabela("Internet")
   
   Tabela.Close

End Sub

Private Sub cmbInclusaobanco_Click()
   GravaBanco 1
   lstBanco.SetFocus
End Sub

Private Sub cmdAlteracaoAgencia_Click()
   GravaAgencia 2
   If cmbAgencia.ListIndex = -1 Then
      MsgBox "Selecione a agência", 16, Titulo
      cmbAgencia.SetFocus
      Exit Sub
   End If
   cmbAgencia.SetFocus
End Sub

Private Sub cmdAlteracaoBanco_Click()
   GravaBanco 2
   lstBanco.SetFocus
End Sub

Private Sub cmdAlteracaoConta_Click()
   GravaConta 2
   txtConta = ""
   mskLimite = 0
   lstContaCorrente.SetFocus
End Sub

Private Sub cmdExclusaoAgencia_Click()
   If lstBanco.ListIndex = -1 Then
      MsgBox "Selecione o banco da agência", 16, Titulo
      lstBanco.SetFocus
      Exit Sub
   End If
   If cmbAgencia.ListIndex = -1 Then
      MsgBox "Selecione a agência", 16, Titulo
      cmbAgencia.SetFocus
      Exit Sub
   End If
   If MsgBox("Confirme a exclusão da agencia", 36, "Exclusão") <> 6 Then
      Exit Sub
   End If
   On Error GoTo ErrExcluiAgencia

   Set Consulta = Banco.QueryDefs("DelAgencia")
   Consulta("parIdBanco") = lstBanco.ItemData(lstBanco.ListIndex)
   Consulta("parIdAgencia") = cmbAgencia.ItemData(cmbAgencia.ListIndex)

   Consulta.Execute
   ComboAgencia cmbAgencia, lstBanco.ItemData(lstBanco.ListIndex)
   cmbAgencia.SetFocus
   Exit Sub
   
ErrExcluiAgencia:
   MsgBox MensErro, 16, "Falha na exclusão da agência"
   Exit Sub
End Sub

Private Sub cmdExclusaoBanco_Click()
   If lstBanco.ListIndex = -1 Then
      MsgBox "Selecione o banco a ser excluído", 16, "Exclusão"
      lstBanco.SetFocus
      Exit Sub
   End If
   If MsgBox("Confirme a exclusão do banco", 36, "Exclusão") <> 6 Then
      Exit Sub
   End If
   On Error GoTo ErrExcluiBanco
   
   Set Consulta = Banco.QueryDefs("DelBanco")
   Consulta("parIdBanco") = lstBanco.ItemData(lstBanco.ListIndex)
   
   Consulta.Execute
   
   ComboBanco lstBanco
   lstBanco.SetFocus
   Exit Sub
   
ErrExcluiBanco:
   MsgBox MensErro, 16, "Falha na exclusão do banco"
   Exit Sub
End Sub

Private Sub cmdInclusaoAgencia_Click()
   GravaAgencia 1
   cmbAgencia.SetFocus
End Sub

Private Sub cmdInclusaoConta_Click()
   GravaConta 1
   txtConta = ""
   mskLimite = 0
   lstContaCorrente.SetFocus
End Sub

Private Sub cmdSaida_Click()
   Unload Me
End Sub
Private Sub Form_Load()
   ComboBanco lstBanco
   ' Se existir pelo um banco cadastrado, seleciona automaticamente _
     o primeiro banco da lista
   If lstBanco.ListCount > 0 Then
      lstBanco.ListIndex = 0
   End If
End Sub
Private Sub GravaBanco(Operacao As Integer)
   Dim Titulo As String
   If Operacao = 1 Then
      Titulo = "Inclusão"
   Else
      Titulo = "Alteração"
   End If
   If Trim$(txtIdBanco) = "" Then
      MsgBox "Informe o número do banco", 16, Titulo
      txtIdBanco.SetFocus
      Exit Sub
   End If
   If Trim$(txtNomeBanco) = "" Then
      MsgBox "Informe o nome do banco", 16, Titulo
      txtNomeBanco.SetFocus
      Exit Sub
   End If
   Dim MensErro As String, ErroPar As Integer
   On Error GoTo ErrGravaBanco
   ErroPar = False
   MensErro = "Falha na abertura da consulta"
   If Operacao = 1 Then
      Set Consulta = Banco.QueryDefs("InsBanco")
   Else
      Set Consulta = Banco.QueryDefs("UpdBanco")
   End If
   ErroPar = True
   
   Area.BeginTrans
   MensErro = "Falha na alimentação dos parâmetros"
   Consulta("parIdBanco") = txtIdBanco
   Consulta("parNome") = txtNomeBanco
   MensErro = "Falha na gravação dos dados"

   Consulta.Execute

   Area.CommitTrans
   
   ComboBanco lstBanco
   Exit Sub

ErrGravaBanco:
   MsgBox MensErro, 16, "Atualização de Banco"
   If ErroPar Then Area.Rollback
   Exit Sub
End Sub
Private Sub lstBanco_Click()
   Dim Selecao As String
   Selecao = "select * from Banco where IdBanco="
   Selecao = Selecao & lstBanco.ItemData(lstBanco.ListIndex)

   Set Tabela = Banco.OpenRecordset(Selecao, dbOpenSnapshot)

   txtIdBanco = Tabela("IdBanco")
   txtNomeBanco = Tabela("Nome")
   Tabela.Close
End Sub

Private Sub lstContaCorrente_Click()
   If lstBanco.ListIndex = -1 Then Exit Sub
   If cmbAgencia.ListIndex = -1 Then Exit Sub
   txtConta = ""
   mskLimite = 0
   Set Consulta = Banco.QueryDefs("SelContaCorrente")
   Consulta("parIdBanco") = lstBanco.ItemData(lstBanco.ListIndex)
   Consulta("parIdAgencia") = cmbAgencia.ItemData(cmbAgencia.ListIndex)
   Consulta("parConta") = lstContaCorrente.List(lstContaCorrente.ListIndex)
   Set Tabela = Consulta.OpenRecordset()
   
   txtConta = Tabela("Conta")
   If Not IsNull(Tabela("Limite")) Then mskLimite = Format$(Tabela("Limite"), "###,###,##0.00")
   
   chkTipo = Tabela("Tipo")
   
   Tabela.Close
End Sub


Private Sub SSTab1_Click(PreviousTab As Integer)
   If PreviousTab = 0 Then
      If lstBanco.ListIndex = -1 Then
         MsgBox "Selecione um banco...", 16, "Atenção!"
         lstBanco.SetFocus
         Exit Sub
      End If
      ComboAgencia cmbAgencia, lstBanco.ItemData(lstBanco.ListIndex)
   End If
   If PreviousTab = 1 Then
      If lstBanco.ListIndex = -1 Then
         MsgBox "Selecione um banco...", 16, "Atenção!"
         lstBanco.SetFocus
         Exit Sub
      End If
      If cmbAgencia.ListIndex = -1 Then
         MsgBox "Selecione uma agência...", 16, "Atenção!"
         cmbAgencia.SetFocus
         Exit Sub
      End If
      ComboContaCorrente lstContaCorrente, lstBanco.ItemData(lstBanco.ListIndex), cmbAgencia.ItemData(cmbAgencia.ListIndex)
   End If
End Sub

Private Sub txtIdBanco_KeyPress(KeyAscii As Integer)
   ' Aproveitamos o evento "KEYPRESS" de "txtIdBanco" para ilustrar como você pode _
     anular o acionamento de uma tecla: a função ISNUMERIC verifica se uma determinada _
     expressão é numérica, enquanto as função CHR retorna o caractere correspondente a _
     um número na table ASCII.
   If Not IsNumeric(Chr(KeyAscii)) Then KeyAscii = 0
End Sub
Private Sub txtNomeBanco_KeyPress(KeyAscii As Integer)
   ' Neste evento exemplificamos como modificar o valor de uma tecla acionada pelo _
     usuário, trocando o caractere digitado por outro. Nesse caso, estamos convertendo _
     qualquer digitada para sua respectiva maiúscula.
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
Private Sub GravaAgencia(Operacao As Integer)
   Dim Titulo As String
   If Operacao = 1 Then
      Titulo = "Inclusão"
   Else
      Titulo = "Alteração"
   End If
   If lstBanco.ListIndex = -1 Then
      MsgBox "Selecione o banco da agência", 16, Titulo
      lstBanco.SetFocus
      Exit Sub
   End If
   If Trim$(txtNumeroAgencia) = "" Then
      MsgBox "Informe o número da agência", 16, Titulo
      txtNumeroAgencia.SetFocus
      Exit Sub
   End If
   If Trim$(txtNomeAgencia) = "" Then
      MsgBox "Informe o nome da agência", 16, Titulo
      txtNomeAgencia.SetFocus
      Exit Sub
   End If
   Dim MensErro As String, ErroPar As Integer
   
   On Error GoTo ErrGravaAgencia
   ErroPar = False
   MensErro = "Falha na abertura da consulta"
   If Operacao = 1 Then
      Set Consulta = Banco.QueryDefs("InsAgencia")
   Else
      Set Consulta = Banco.QueryDefs("UpdAgencia")
   End If
   ErroPar = True
   
   Area.BeginTrans
   MensErro = "Falha na alimentação dos parâmetros"
   Consulta("parIdBanco") = lstBanco.ItemData(lstBanco.ListIndex)
   Consulta("parIdAgencia") = txtNumeroAgencia
   Consulta("parNome") = txtNomeBanco
   Consulta("parEndereco") = txtEndereco
   Consulta("parGerente") = txtGerente
   Consulta("parTelefone") = mskTelefone
   Consulta("parFAX") = mskFAX
   Consulta("parRamal") = txtRamal
   Consulta("parInternet") = txtInternet

   MensErro = "Falha na gravação dos dados"

   Consulta.Execute

   Area.CommitTrans
   
   ComboBanco lstBanco
   Exit Sub

ErrGravaAgencia:
   MsgBox MensErro, 16, "Atualização de Agência"
   If ErroPar Then Area.Rollback
   Exit Sub
End Sub

Private Sub GravaConta(Operacao As Integer)
   Dim Titulo As String
   If Operacao = 1 Then
      Titulo = "Inclusão"
   Else
      Titulo = "Alteração"
   End If
   If lstBanco.ListIndex = -1 Then
      MsgBox "Selecione o banco", 16, Titulo
      lstBanco.SetFocus
      Exit Sub
   End If
   If cmbAgencia.ListIndex = -1 Then
      MsgBox "Selecione a agencia", 16, Titulo
      cmbAgencia.SetFocus
      Exit Sub
   End If
   If Trim$(txtConta) = "" Then
      MsgBox "Informe o número da conta corrente", 16, Titulo
      txtConta.SetFocus
      Exit Sub
   End If
   Dim MensErro As String, ErroPar As Integer
   
   On Error GoTo ErrGravaConta
   ErroPar = False
   MensErro = "Falha na abertura da consulta"
   If Operacao = 1 Then
      Set Consulta = Banco.QueryDefs("InsContaCorrente")
   Else
      Set Consulta = Banco.QueryDefs("UpdContaCorrente")
   End If
   ErroPar = True
   
   Area.BeginTrans
   MensErro = "Falha na alimentação dos parâmetros"
   Consulta("parIdBanco") = lstBanco.ItemData(lstBanco.ListIndex)
   Consulta("parIdAgencia") = cmbAgencia.ItemData(cmbAgencia.ListIndex)
   Consulta("parConta") = txtConta
   Consulta("parLimite") = Format(mskLimite, "0.00")
   Consulta("parTipo") = chkTipo
   MensErro = "Falha na gravação dos dados"

   Consulta.Execute

   Area.CommitTrans
   
   ComboContaCorrente lstContaCorrente, lstBanco.ItemData(lstBanco.ListIndex), cmbAgencia.ItemData(cmbAgencia.ListIndex)
   Exit Sub

ErrGravaConta:
   MsgBox MensErro, 16, "Atualização de Conta Corrente"
   If ErroPar Then Area.Rollback
   Exit Sub
End Sub
