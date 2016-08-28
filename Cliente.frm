VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form frmCliente 
   Caption         =   "Cadastro de Clientes"
   ClientHeight    =   5295
   ClientLeft      =   465
   ClientTop       =   825
   ClientWidth     =   7680
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5295
   ScaleWidth      =   7680
   Begin Crystal.CrystalReport Relatorio 
      Left            =   5640
      Top             =   240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "C:\Erica_VB5\CPR\cliente.rpt"
   End
   Begin VB.CommandButton cmdImpressao 
      Caption         =   "&Imprimir"
      Height          =   420
      Left            =   6180
      TabIndex        =   35
      Top             =   2460
      Width           =   1230
   End
   Begin VB.CommandButton cmdHistorico 
      Caption         =   "&Histórico de Cliente"
      Height          =   420
      Left            =   4410
      TabIndex        =   37
      Top             =   4635
      Width           =   1815
   End
   Begin VB.CommandButton cmdSaida 
      Caption         =   "&Retornar"
      Height          =   420
      Left            =   6210
      TabIndex        =   36
      Top             =   3360
      Width           =   1230
   End
   Begin VB.CommandButton cmdExclusao 
      Caption         =   "&Excluir"
      Height          =   420
      Left            =   6210
      TabIndex        =   34
      Top             =   1800
      Width           =   1230
   End
   Begin VB.CommandButton cmdAlteracao 
      Caption         =   "&Alterar"
      Height          =   420
      Left            =   6210
      TabIndex        =   33
      Top             =   1080
      Width           =   1230
   End
   Begin VB.CommandButton cmdInclusao 
      Caption         =   "&Incluir"
      Height          =   420
      Left            =   6210
      TabIndex        =   32
      Top             =   450
      Width           =   1230
   End
   Begin VB.TextBox txtInternet 
      Height          =   285
      Left            =   1845
      MaxLength       =   20
      TabIndex        =   17
      Top             =   4680
      Width           =   2220
   End
   Begin VB.TextBox txtRamal 
      Height          =   285
      Left            =   5310
      TabIndex        =   15
      Top             =   4050
      Width           =   645
   End
   Begin VB.TextBox txtDDD 
      Height          =   285
      Left            =   3060
      TabIndex        =   13
      Top             =   4050
      Width           =   555
   End
   Begin VB.TextBox txtContato 
      Height          =   285
      Left            =   225
      MaxLength       =   20
      TabIndex        =   12
      Top             =   4050
      Width           =   2220
   End
   Begin VB.TextBox txtEstado 
      Height          =   285
      Left            =   4860
      MaxLength       =   2
      TabIndex        =   11
      Top             =   3420
      Width           =   510
   End
   Begin VB.TextBox txtCidade 
      Height          =   285
      Left            =   1485
      MaxLength       =   30
      TabIndex        =   10
      Top             =   3420
      Width           =   3210
   End
   Begin VB.TextBox txtBairro 
      Height          =   285
      Left            =   3600
      MaxLength       =   20
      TabIndex        =   8
      Top             =   2790
      Width           =   2220
   End
   Begin VB.TextBox txtLogradouro 
      Height          =   285
      Left            =   225
      MaxLength       =   30
      TabIndex        =   7
      Top             =   2790
      Width           =   3210
   End
   Begin VB.TextBox txtInscricaoEstadual 
      Height          =   285
      Left            =   225
      MaxLength       =   20
      TabIndex        =   6
      Top             =   2160
      Width           =   2220
   End
   Begin VB.TextBox txtRazaoSocial 
      Height          =   285
      Left            =   225
      MaxLength       =   50
      TabIndex        =   5
      Top             =   1575
      Width           =   5325
   End
   Begin VB.TextBox txtCodinome 
      Height          =   285
      Left            =   225
      MaxLength       =   20
      TabIndex        =   4
      Top             =   990
      Width           =   2175
   End
   Begin VB.ComboBox cmbCodinome 
      Height          =   315
      Left            =   225
      TabIndex        =   0
      Top             =   360
      Width           =   2490
   End
   Begin MSMask.MaskEdBox mskFAX 
      Height          =   330
      Left            =   225
      TabIndex        =   16
      Top             =   4680
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   582
      _Version        =   327680
      PromptInclude   =   0   'False
      MaxLength       =   10
      Mask            =   "####-##-##"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox mskTelefone 
      Height          =   330
      Left            =   3960
      TabIndex        =   14
      Top             =   4050
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   582
      _Version        =   327680
      PromptInclude   =   0   'False
      MaxLength       =   10
      Mask            =   "####-##-##"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox mskCEP 
      Height          =   330
      Left            =   225
      TabIndex        =   9
      Top             =   3420
      Width           =   1050
      _ExtentX        =   1852
      _ExtentY        =   582
      _Version        =   327680
      PromptInclude   =   0   'False
      MaxLength       =   9
      Mask            =   "#####-###"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox mskCPF_CGC 
      Height          =   330
      Left            =   2970
      TabIndex        =   1
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
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      Caption         =   "Internet"
      Height          =   195
      Left            =   1845
      TabIndex        =   31
      Top             =   4455
      Width           =   540
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      Caption         =   "FAX"
      Height          =   195
      Left            =   225
      TabIndex        =   30
      Top             =   4455
      Width           =   300
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      Caption         =   "Ramal"
      Height          =   195
      Left            =   5310
      TabIndex        =   29
      Top             =   3825
      Width           =   450
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      Caption         =   "Telefone"
      Height          =   195
      Left            =   3960
      TabIndex        =   28
      Top             =   3825
      Width           =   630
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "DDD"
      Height          =   195
      Left            =   3060
      TabIndex        =   27
      Top             =   3825
      Width           =   360
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "Contato"
      Height          =   195
      Left            =   225
      TabIndex        =   26
      Top             =   3825
      Width           =   555
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "Estado"
      Height          =   195
      Left            =   4860
      TabIndex        =   25
      Top             =   3195
      Width           =   495
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Cidade"
      Height          =   195
      Left            =   1485
      TabIndex        =   24
      Top             =   3195
      Width           =   495
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "CEP"
      Height          =   195
      Left            =   225
      TabIndex        =   23
      Top             =   3195
      Width           =   315
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Bairro"
      Height          =   195
      Left            =   3600
      TabIndex        =   22
      Top             =   2565
      Width           =   405
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Logradouro"
      Height          =   195
      Left            =   225
      TabIndex        =   21
      Top             =   2565
      Width           =   810
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Inscrição Estadual"
      Height          =   195
      Left            =   225
      TabIndex        =   20
      Top             =   1980
      Width           =   1305
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Razão Social"
      Height          =   195
      Left            =   225
      TabIndex        =   19
      Top             =   1395
      Width           =   945
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Codinome"
      Height          =   195
      Left            =   225
      TabIndex        =   18
      Top             =   765
      Width           =   705
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "CPF/CGC"
      Height          =   195
      Left            =   2970
      TabIndex        =   3
      Top             =   135
      Width           =   705
   End
   Begin VB.Label lblCodinome 
      AutoSize        =   -1  'True
      Caption         =   "Clientes Cadastrados"
      Height          =   195
      Left            =   225
      TabIndex        =   2
      Top             =   135
      Width           =   1485
   End
End
Attribute VB_Name = "frmCliente"
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
'| Descrição:  Formulário para cadastro de clientes utili-|
'|             zando como metedologia de acesso e   grava-|
'|             ção "queries" (consultas) pré-desenvolvidas|
'|             na base de dados CPR.MDB                   |
'+--------------------------------------------------------+

Private Sub GravaCliente(Operacao As Integer)
' A ROTINA GravaCliente, COMO SEU PRÓPRIO NOME INDICA, EFETUA A GRAVAÇÃO _
  DOS DADOS NA TABELA "Cliente" DE "CPR.MDB"
' O parâmetro "operação" serve para indicar se a operação é de inclusão de _
  novo cliente (valor 1) ou de alteração dos dados de um cliente já cadas- _
  trado (qualquer valor diferente de 1)
  
   Dim Titulo As String ' Variável para configuração do título de janelas de mensagem
   If Operacao = 1 Then
      Titulo = "Inclusão"
   Else
      Titulo = "Alteração"
   End If
   ' Os quatro próximos IF´s efetuam a verificação de controles com preen- _
     chimento obrigatório antes da gravação do cliente na base de dados
   If Trim$(mskCPF_CGC) = "" Then
      MsgBox "Informe o número de CPF/CGC", 16, Titulo
      mskCPF_CGC.SetFocus
      Exit Sub
   End If
   If Len(Trim$(mskCPF_CGC)) <> 11 And Len(Trim$(mskCPF_CGC)) <> 14 Then
      MsgBox "Corrija o número de CPF/CGC", 16, Titulo
      mskCPF_CGC.SetFocus
      Exit Sub
   End If
   If Trim$(txtRazaoSocial) = "" Then
      MsgBox "Informe a razão social do cliente", 16, Titulo
      txtRazaoSocial.SetFocus
      Exit Sub
   End If
   If Trim$(txtCodinome) = "" Then
      MsgBox "Informe o codinome (apelido) do cliente", 16, Titulo
      txtCodinome.SetFocus
      Exit Sub
   End If
   'O Visual Basic permite declarar variáveis de memória em qualquer ponto _
    do código:
   Dim MensErro As String, ErroPar As Integer
   ' Com a declaração "ON ERROR.." e as variáveis "ErroPar" e "MensErro" _
     tratamos quaisquer erros que porventura ocorram na execução a partir _
     deste ponto do código
   On Error GoTo ErrGravaCliente  'Quando ocorrer um erro, o controle será desviado _
'                                   para o rótulo ErrGravaCliente
   ErroPar = False
   MensErro = "Falha na abertura da consulta"
   If Operacao = 1 Then
      Set Consulta = Banco.QueryDefs("InsCliente")
   Else
      Set Consulta = Banco.QueryDefs("UpdCliente")
   End If
   ErroPar = True
   ' Início da transação de gravação dos dados do cliente
   Area.BeginTrans
   MensErro = "Falha na alimentação dos parâmetros"
   Consulta("parCPF_CGC") = mskCPF_CGC.Text
   Consulta("parRazaoSocial") = txtRazaoSocial
   Consulta("parCodinome") = txtCodinome
   Consulta("parInscricaoEstadual") = txtInscricaoEstadual
   Consulta("parLogradouro") = txtLogradouro
   Consulta("parBairro") = txtBairro
   Consulta("parCEP") = mskCEP
   Consulta("parCidade") = txtCidade
   Consulta("parEstado") = txtEstado
   Consulta("parContato") = txtContato
   Consulta("parDDD") = txtDDD
   Consulta("parTelefone") = mskTelefone
   Consulta("parRamal") = txtRamal
   Consulta("parFAX") = mskFAX
   Consulta("parInternet") = txtInternet
   MensErro = "Falha na gravação dos dados"
   ' Execução da consulta após a alimentação de sxeus parâmetros
   Consulta.Execute
   ' Efetivação da transação
   Area.CommitTrans
   ' Após a gravação, o combobox "cmbCodinome" é devidamente atualizado _
     invocando-se a rotina "ComboCliente"
   ComboCliente cmbCodinome
   Exit Sub ' Este comando evita que o bloco de tratamento de erros seja _
'             executado, pois, uma vez que o controle chegou até aqui, _
'             nenhum erro ocorreu na gravação dos dados do cliente
   
' O rótulo "ErrGravaCliente:" trata o erro ocorrido, emite uma mensagem _
  para o usuário e desfaz a transação (ROLLBACK) de gravação dos dados
ErrGravaCliente:
   MsgBox MensErro, 16, "Atualização de Cliente"
   If ErroPar Then Area.Rollback
   Exit Sub
End Sub
Private Sub cmbCodinome_Click()
   ' Montagem da declaração SQL para leitura do cliente selecionado no combo
   Dim Selecao As String
   Selecao = "select * from Cliente where Codinome=" & """"
   Selecao = Selecao & cmbCodinome.Text & """"
   ' Execução da declaração
   Set Tabela = Banco.OpenRecordset(Selecao, dbOpenSnapshot)
   ' Alimentação dos controles a partir das colunas do recordset "Tabela"
   mskCPF_CGC = Tabela("CPF_CGC")
   txtCodinome = Tabela("Codinome")
   txtRazaoSocial = Tabela("RazaoSocial")
   ' Aqui a função ISNULL permite verificar o conteúdo das colunas do _
     Recordset: essa verificação é necessária, pois, se tentarmos copiar _
     um conteúdo nulo para um controle, ocorrerá um erro
   If Not IsNull(Tabela("InscricaoEstadual")) Then txtInscricaoEstadual = Tabela("InscricaoEstadual")
   If Not IsNull(Tabela("Logradouro")) Then txtLogradouro = Tabela("Logradouro")
   If Not IsNull(Tabela("Bairro")) Then txtBairro = Tabela("Bairro")
   If Not IsNull(Tabela("CEP")) Then mskCEP = Tabela("CEP")
   If Not IsNull(Tabela("Cidade")) Then txtCidade = Tabela("Cidade")
   If Not IsNull(Tabela("Estado")) Then txtEstado = Tabela("Estado")
   If Not IsNull(Tabela("Contato")) Then txtContato = Tabela("Contato")
   If Not IsNull(Tabela("DDD")) Then txtDDD = Tabela("DDD")
   If Not IsNull(Tabela("Telefone")) Then mskTelefone = Tabela("Telefone")
   If Not IsNull(Tabela("Ramal")) Then txtRamal = Tabela("Ramal")
   If Not IsNull(Tabela("FAX")) Then mskFAX = Tabela("FAX")
   If Not IsNull(Tabela("Internet")) Then txtInternet = Tabela("Internet")
   ' O método CLOSE fecha o recordset, evitando que ele fique aberto após _
     sua utilização
   Tabela.Close

End Sub
Private Sub cmbCodinome_KeyPress(KeyAscii As Integer)
   ' Por definição, o deslocamento do foco de um controle para outro _
     é feito mediante o pressionamento da tecla <TAB>. Como alternativa, _
     podemos fazer com que esse deslocamento seja feito também com o _
     pressionamento da tecla <ENTER>, mas essa facilidade só pode ser _
     implementada via código. Esse bloco exemplifica esse recurso.
   If KeyAscii = 13 Then
      KeyAscii = 0
      mskCPF_CGC.SetFocus
      Exit Sub
   End If
End Sub
Private Sub cmdAlteracao_Click()
   ' Invoca a rotina GRAVACLIENTE passando como argumento o valor 2: _
     qualquer valor diferente de 1 corresponde à operação de alteração
   GravaCliente 2
End Sub
Private Sub cmdAlteracao_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   mdiCPR.barMensagem.Panels(1).Text = "Altera os dados do cliente com os dados informados"
End Sub
Private Sub cmdExclusao_Click()
   'Para excluir um cliente, é necessário que o usuário informe o CGC/CPF _
    do cliente: isso pode ser feito clicando-se em um dos codinomes do combo _
    CMBCODINOME, que edita os dados para os controles, entre eles o CGC/CPF
   If Trim$(mskCPF_CGC) = "" Then
      MsgBox "Informe o número de CPF/CGC", 16, "Exclusão"
      mskCPF_CGC.SetFocus
      Exit Sub
   End If
   If MsgBox("Confirme a exclusão do cliente", 36, "Exclusão") <> 6 Then
      Exit Sub
   End If
   On Error GoTo ErrExcluiCliente
   ' Invoca a consulta DELCLIENTE, previamente escrita e gravada em CPR.MDB.
   Set Consulta = Banco.QueryDefs("DelCliente")
   Consulta("parCPF_CGC") = mskCPF_CGC.Text
   Consulta.Execute
   ComboCliente cmbCodinome
   Exit Sub ' Evita que o bloco para tratamento de erros seja executado quando _
'             a exclusão foi bem sucedida
   
ErrExcluiCliente:
   MsgBox MensErro, 16, "Falha na exclusão do Cliente", Titulo
   Exit Sub
End Sub
Private Sub cmdExclusao_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   ' Explica a finalidade do botão escrevendo uma mensagem na barra de status _
     do formulário principal (mdiCPR)
   mdiCPR.barMensagem.Panels(1).Text = "Exclui o cliente selecionado da base de dados"
End Sub
Private Sub cmdHistorico_Click()
   ' Invoa o formulário de históricos de clientes
   frmHistoricoCliente.Show
End Sub

Private Sub cmdImpressao_Click()
   Relatorio.Action = 1
End Sub

Private Sub cmdInclusao_Click()
   ' Invoca a rotina GRAVACLIENTE passando como argumento o valor 1, _
     correspondente à operação de inclusão
   GravaCliente 1
End Sub
Private Sub cmdInclusao_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   mdiCPR.barMensagem.Panels(1).Text = "Inclui novo cliente com os dados informados"
End Sub
Private Sub cmdSaida_Click()
   Unload Me
End Sub
Private Sub cmdSaida_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   mdiCPR.barMensagem.Panels(1).Text = "Fecha o cadastro de clientes e retorna à janela principal"
End Sub


Private Sub Form_Load()
   ' INVOCA A ROTINA ComboCliente PARA PREENCHER O COMBOBOX cmbCodinome _
     COM OS CODINOMES DE CLIENTES JÁ EXISTENTES NA TABELA "CLIENTE" DA _
     BASE DE DADOS "CPR.MDB"
   ComboCliente cmbCodinome
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   mdiCPR.barMensagem.Panels(1).Text = ""
End Sub


Private Sub lblCodinome_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   mdiCPR.barMensagem.Panels(1).Text = "Lista de clientes já cadastrados no sistema"
End Sub
Private Sub mskCPF_CGC_GotFocus()
   mskCPF_CGC.Mask = "##############"
End Sub

Private Sub mskCPF_CGC_LostFocus()
   ' As linhas iniciais da rotina "mskCPF_CGC_LostFocus" testam o conteúdo _
     digitado no controle "mskCPF_CGC": se o seu comprimento (LEN) for igual _
     a 11, trata-se de CPF; se for igual a 14, trata-se de CGC. Nesses casos, _
     a máscara do controle é configurado para o formato compatível (CPF ou _
     CGC) - note a utilização da variável "temp" para salvar o conteúdo do _
     controle antes da reconfiguração da máscara e para devolver esse con- _
     teúdo após essa reconfiguração.
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
   ' Note que o usuário pode editar os dados de um cliente através de um _
     clique no combobox CMBCODINOME ou informando o CPF/CGC no controle _
     mskCPF_CGC.
   Set Consulta = Banco.QueryDefs("SelCliente")
   Consulta("parCPF_CGC") = temp
   Set Tabela = Consulta.OpenRecordset()
   ' Caso o CPF/CGC não exista, o recordset é fechado e a rotina interrompida
   If Tabela.RecordCount = 0 Then
      Tabela.Close
      Exit Sub
   End If
   ' Copia os dados do cliente para os controles exatamente como em _
     cmbCodinome_Click
   txtCodinome = Tabela("Codinome")
   txtRazaoSocial = Tabela("RazaoSocial")
   txtInscricaoEstadual = Tabela("InscricaoEstadual")
   txtLogradouro = Tabela("Logradouro")
   txtBairro = Tabela("Bairro")
   mskCEP = Tabela("CEP")
   txtCidade = Tabela("Cidade")
   txtEstado = Tabela("Estado")
   txtContato = Tabela("Contato")
   txtDDD = Tabela("DDD")
   mskTelefone = Tabela("Telefone")
   txtRamal = Tabela("Ramal")
   mskFAX = Tabela("FAX")
   txtInternet = Tabela("Internet")
   Tabela.Close
   
End Sub
Private Sub txtCodinome_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   mdiCPR.barMensagem.Panels(1).Text = "Codinome (""apelido"") do cliente"
End Sub
Private Sub txtContato_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   mdiCPR.barMensagem.Panels(1).Text = "Nome da pessoa de contato no cliente"
End Sub
Private Sub txtDDD_KeyPress(KeyAscii As Integer)
   If KeyAscii < 48 Or KeyAscii > 57 Then
      KeyAscii = 0
   End If
End Sub
Private Sub txtDDD_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   mdiCPR.barMensagem.Panels(1).Text = "Código de dsicagem direta à distância"
End Sub
Private Sub txtInscricaoEstadual_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   mdiCPR.barMensagem.Panels(1).Text = "Inscrição estadual do cliente"
End Sub
Private Sub txtInternet_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   mdiCPR.barMensagem.Panels(1).Text = "Endereço do clinte na Internet"
End Sub
Private Sub txtLogradouro_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   mdiCPR.barMensagem.Panels(1).Text = "Rua, avenida, praça, etc."
End Sub
Private Sub txtRazaoSocial_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   mdiCPR.barMensagem.Panels(1).Text = "Razão social do cliente"
End Sub
