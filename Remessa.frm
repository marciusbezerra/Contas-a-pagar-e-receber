VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmRemessa 
   Caption         =   "Remessa de Títulos para Cartório"
   ClientHeight    =   3075
   ClientLeft      =   255
   ClientTop       =   1260
   ClientWidth     =   6090
   LinkTopic       =   "Form8"
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3075
   ScaleWidth      =   6090
   Begin VB.CommandButton cmdInclusao 
      Caption         =   "&Incluir"
      Height          =   420
      Left            =   4470
      TabIndex        =   11
      Top             =   135
      Width           =   1230
   End
   Begin VB.CommandButton cmdAlteracao 
      Caption         =   "&Alterar"
      Height          =   420
      Left            =   4470
      TabIndex        =   12
      Top             =   810
      Width           =   1230
   End
   Begin VB.CommandButton cmdExclusao 
      Caption         =   "&Excluir"
      Height          =   420
      Left            =   4470
      TabIndex        =   13
      Top             =   1485
      Width           =   1230
   End
   Begin VB.CommandButton cmdSaida 
      Caption         =   "&Retornar"
      Height          =   420
      Left            =   4470
      TabIndex        =   14
      Top             =   2340
      Width           =   1230
   End
   Begin VB.CommandButton cmdCadastro 
      Caption         =   "&Cadastro"
      Height          =   330
      Left            =   2790
      TabIndex        =   2
      Top             =   315
      Width           =   1140
   End
   Begin VB.TextBox txtNumeroTitulo 
      Height          =   285
      Left            =   135
      MaxLength       =   18
      TabIndex        =   4
      Top             =   990
      Width           =   1500
   End
   Begin VB.ComboBox cmbCartorio 
      Height          =   315
      Left            =   135
      TabIndex        =   1
      Top             =   315
      Width           =   2490
   End
   Begin MSMask.MaskEdBox mskRemessa 
      Height          =   330
      Left            =   120
      TabIndex        =   6
      Top             =   1620
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   582
      _Version        =   327680
      MaxLength       =   8
      Format          =   "ddddd"
      Mask            =   "##/##/##"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox mskProtesto 
      Height          =   330
      Left            =   1560
      TabIndex        =   8
      Top             =   1620
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   582
      _Version        =   327680
      MaxLength       =   8
      Format          =   "ddddd"
      Mask            =   "##/##/##"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox mskBaixa 
      Height          =   330
      Left            =   3000
      TabIndex        =   10
      Top             =   1620
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   582
      _Version        =   327680
      MaxLength       =   8
      Format          =   "ddddd"
      Mask            =   "##/##/##"
      PromptChar      =   " "
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Número do Título"
      Height          =   195
      Left            =   135
      TabIndex        =   3
      Top             =   765
      Width           =   1245
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Remessa"
      Height          =   195
      Left            =   135
      TabIndex        =   5
      Top             =   1395
      Width           =   660
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Protesto"
      Height          =   195
      Left            =   1575
      TabIndex        =   7
      Top             =   1395
      Width           =   585
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Baixa"
      Height          =   195
      Left            =   3015
      TabIndex        =   9
      Top             =   1395
      Width           =   390
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Cartórios Cadastrados"
      Height          =   195
      Left            =   135
      TabIndex        =   0
      Top             =   90
      Width           =   1545
   End
End
Attribute VB_Name = "frmRemessa"
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
'| Descrição:  Formulário para remessa de títulos  para    |
'|             cartório                                    |
'+---------------------------------------------------------+


Private Sub cmdAlteracao_Click()
   If cmbCartorio.ListIndex = -1 Then
      MsgBox "Informe o cartório...", 16, Titulo
      cmbCartorio.SetFocus
      Exit Sub
   End If
   If Not mskRemessa = "  /  /  " Then
      If Not ConsisteData(mskRemessa) Then mskRemessa.SetFocus
   End If
   If Not mskBaixa = "  /  /  " Then
      If Not ConsisteData(mskBaixa) Then mskBaixa.SetFocus
   End If
   If Not mskProtesto = "  /  /  " Then
      If Not ConsisteData(mskProtesto) Then mskProtesto.SetFocus
   End If
   If Trim$(txtNumeroTitulo) = "" Then
      MsgBox "Informe o número do título", 16, Titulo
      txtNumeroTitulo.SetFocus
      Exit Sub
   End If
   Selecao = "SELECT * from Recebimento where NumeroTitulo=" & """"
   Selecao = Selecao & txtNumeroTitulo & """"

   Set Tabela = Banco.OpenRecordset(Selecao, dbOpenSnapshot)
   
   If Tabela.RecordCount = 0 Then
      MsgBox "Título não cadastrado...", 16, Titulo
      Tabela.Close
      txtNumeroTitulo.SetFocus
      Exit Sub
   End If
   Dim MensErro As String, ErroPar As Integer
   On Error GoTo ErrGravaRemessaCar
   ErroPar = False
   MensErro = "Falha na abertura da consulta"
   Set Consulta = Banco.QueryDefs("UpdRemessaCartorio")
   ErroPar = True
   
   Area.BeginTrans
   MensErro = "Falha na alimentação dos parâmetros"

   Consulta("parIdCartorio") = cmbCartorio.ItemData(cmbCartorio.ListIndex)
   Consulta("parNumeroTitulo") = txtNumeroTitulo
   Consulta("parRemessa") = mskRemessa
   Consulta("parBaixa") = mskBaixa
   Consulta("parProtesto") = mskProtesto
 
   MensErro = "Falha na gravação dos dados"

   Consulta.Execute
   Area.CommitTrans
   
   mskBaixa = "  /  /  "
   mskRemessa = "  /  /  "
   mskProtesto = "  /  /  "
   txtNumeroTitulo = ""
   txtNumeroTitulo.SetFocus

   Exit Sub

ErrGravaRemessaCar:
   MsgBox MensErro, 16, "Alteração"
   If ErroPar Then Area.Rollback
   Exit Sub
End Sub

Private Sub cmdCadastro_Click()
   frmCartorio.Show
End Sub

Private Sub cmdInclusao_Click()
   If cmbCartorio.ListIndex = -1 Then
      MsgBox "Informe o cartório...", 16, "Atenção"
      cmbCartorio.SetFocus
      Exit Sub
   End If
   If Not mskRemessa = "  /  /  " Then
      If Not ConsisteData(mskRemessa) Then mskRemessa.SetFocus
   End If
   If Not mskBaixa = "  /  /  " Then
      If Not ConsisteData(mskBaixa) Then mskBaixa.SetFocus
   End If
   If Not mskProtesto = "  /  /  " Then
      If Not ConsisteData(mskProtesto) Then mskProtesto.SetFocus
   End If
   If Trim$(txtNumeroTitulo) = "" Then
      MsgBox "Informe o número do título", 16, Titulo
      txtNumeroTitulo.SetFocus
      Exit Sub
   End If
   Selecao = "SELECT * from Recebimento where NumeroTitulo=" & """"
   Selecao = Selecao & txtNumeroTitulo & """"

   Set Tabela = Banco.OpenRecordset(Selecao, dbOpenSnapshot)
   
   If Tabela.RecordCount = 0 Then
      MsgBox "Título não cadastrado...", 16, "Atenção"
      Tabela.Close
      txtNumeroTitulo.SetFocus
      Exit Sub
   End If
   Dim MensErro As String, ErroPar As Integer
   On Error GoTo ErrGravaRemessaCartorio
   ErroPar = False
   MensErro = "Falha na abertura da consulta"
   Set Consulta = Banco.QueryDefs("InsRemessaCartorio")
   ErroPar = True
   
   Area.BeginTrans
   MensErro = "Falha na alimentação dos parâmetros"

   Consulta("parIdCartorio") = cmbCartorio.ItemData(cmbCartorio.ListIndex)
   Consulta("parNumeroTitulo") = txtNumeroTitulo
   Consulta("parRemessa") = mskRemessa
   Consulta("parBaixa") = mskBaixa
   Consulta("parProtesto") = mskProtesto

   MensErro = "Falha na gravação dos dados"

   Consulta.Execute
   Area.CommitTrans
   
   mskBaixa = ""
   mskRemessa = ""
   mskProtesto = ""
   txtNumeroTitulo = ""
   txtNumeroTitulo.SetFocus

   Exit Sub

ErrGravaRemessaCartorio:
   MsgBox MensErro, 16, "Atualização de Remessa para Cartório"
   If ErroPar Then Area.Rollback
   Exit Sub

End Sub

Private Sub cmdSaida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Set Tabela = Banco.OpenRecordset("SelCartorios", dbOpenSnapshot)
   Do Until Tabela.EOF
      cmbCartorio.AddItem Tabela("Nome")
      cmbCartorio.ItemData(cmbCartorio.NewIndex) = Tabela("IdCartorio")
      Tabela.MoveNext
   Loop
   Tabela.Close
End Sub

Private Sub txtNumeroTitulo_LostFocus()
   If Trim$(txtNumeroTitulo) = "" Then Exit Sub
   Selecao = "SELECT * from RemessaCartorio where NumeroTitulo=" & """"
   Selecao = Selecao & txtNumeroTitulo & """"

   Set Tabela = Banco.OpenRecordset(Selecao, dbOpenSnapshot)
   
   If Tabela.RecordCount = 0 Then Exit Sub
   If Not IsNull(Tabela("Baixa")) Then mskBaixa = Tabela("Baixa")
   If Not IsNull(Tabela("Remessa")) Then mskRemessa = Tabela("Remessa")
   If Not IsNull(Tabela("Protesto")) Then mskProtesto = Tabela("Protesto")
   Tabela.Close

End Sub
