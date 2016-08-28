VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmFornecedor 
   Caption         =   "Cadastro de Fornecedores"
   ClientHeight    =   5730
   ClientLeft      =   465
   ClientTop       =   570
   ClientWidth     =   7815
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5730
   ScaleWidth      =   7815
   Begin VB.CommandButton cmdHistorico 
      Caption         =   "&Histórico Fornecedor"
      Height          =   420
      Left            =   2790
      TabIndex        =   42
      Top             =   2010
      Width           =   1860
   End
   Begin VB.TextBox txtPrazoEntrega 
      Height          =   285
      Left            =   5085
      MaxLength       =   3
      TabIndex        =   20
      Top             =   5265
      Width           =   645
   End
   Begin VB.TextBox txtCondicoesDesconto 
      Height          =   285
      Left            =   2655
      MaxLength       =   20
      TabIndex        =   19
      Top             =   5265
      Width           =   2220
   End
   Begin VB.TextBox txtCondicoesPagamento 
      Height          =   285
      Left            =   180
      MaxLength       =   20
      TabIndex        =   18
      Top             =   5265
      Width           =   2220
   End
   Begin VB.CommandButton cmdSaida 
      Caption         =   "&Retornar"
      Height          =   420
      Left            =   6345
      TabIndex        =   38
      Top             =   3915
      Width           =   1230
   End
   Begin VB.CommandButton cmdExclusao 
      Caption         =   "&Excluir"
      Height          =   420
      Left            =   6345
      TabIndex        =   37
      Top             =   1755
      Width           =   1230
   End
   Begin VB.CommandButton cmdAlteracao 
      Caption         =   "&Alterar"
      Height          =   420
      Left            =   6345
      TabIndex        =   36
      Top             =   1035
      Width           =   1230
   End
   Begin VB.CommandButton cmdInclusao 
      Caption         =   "&Incluir"
      Height          =   420
      Left            =   6345
      TabIndex        =   35
      Top             =   315
      Width           =   1230
   End
   Begin VB.TextBox txtInternet 
      Height          =   285
      Left            =   1620
      MaxLength       =   20
      TabIndex        =   17
      Top             =   4635
      Width           =   2220
   End
   Begin VB.TextBox txtRamal 
      Height          =   285
      Left            =   5175
      MaxLength       =   5
      TabIndex        =   15
      Top             =   4005
      Width           =   645
   End
   Begin VB.TextBox txtDDD 
      Height          =   285
      Left            =   3060
      MaxLength       =   4
      TabIndex        =   13
      Top             =   4005
      Width           =   555
   End
   Begin VB.TextBox txtVendedor 
      Height          =   285
      Left            =   225
      MaxLength       =   20
      TabIndex        =   12
      Top             =   4005
      Width           =   2220
   End
   Begin VB.TextBox txtEstado 
      Height          =   285
      Left            =   4950
      MaxLength       =   2
      TabIndex        =   11
      Top             =   3375
      Width           =   510
   End
   Begin VB.TextBox txtCidade 
      Height          =   285
      Left            =   1575
      TabIndex        =   10
      Top             =   3375
      Width           =   3210
   End
   Begin VB.TextBox txtBairro 
      Height          =   285
      Left            =   3600
      TabIndex        =   8
      Top             =   2745
      Width           =   2220
   End
   Begin VB.TextBox txtLogradouro 
      Height          =   285
      Left            =   225
      MaxLength       =   30
      TabIndex        =   7
      Top             =   2745
      Width           =   3210
   End
   Begin VB.TextBox txtInscricaoEstadual 
      Height          =   285
      Left            =   225
      MaxLength       =   20
      TabIndex        =   6
      Top             =   2115
      Width           =   2220
   End
   Begin VB.TextBox txtRazaoSocial 
      Height          =   285
      Left            =   225
      MaxLength       =   50
      TabIndex        =   5
      Top             =   1530
      Width           =   5325
   End
   Begin VB.TextBox txtCodinome 
      Height          =   285
      Left            =   225
      MaxLength       =   20
      TabIndex        =   4
      Top             =   945
      Width           =   2175
   End
   Begin VB.ComboBox cmbCodinome 
      Height          =   315
      Left            =   225
      TabIndex        =   0
      Top             =   315
      Width           =   2490
   End
   Begin MSMask.MaskEdBox mskFAX 
      Height          =   330
      Left            =   225
      TabIndex        =   16
      Top             =   4635
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   582
      _Version        =   327680
      PromptInclude   =   0   'False
      MaxLength       =   10
      Mask            =   "####-##-##"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox mskTelefone 
      Height          =   330
      Left            =   3825
      TabIndex        =   14
      Top             =   4005
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
      Top             =   3375
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
      Top             =   315
      Width           =   1860
      _ExtentX        =   3281
      _ExtentY        =   582
      _Version        =   327680
      PromptInclude   =   0   'False
      MaxLength       =   18
      Mask            =   "##.###.###/####-##"
      PromptChar      =   " "
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      Caption         =   "Prazo de Entrega"
      Height          =   195
      Left            =   5085
      TabIndex        =   41
      Top             =   5040
      Width           =   1230
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      Caption         =   "Condições de Desconto"
      Height          =   195
      Left            =   2655
      TabIndex        =   40
      Top             =   5040
      Width           =   1710
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      Caption         =   "Condições de Pagamento"
      Height          =   195
      Left            =   225
      TabIndex        =   39
      Top             =   5040
      Width           =   1830
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      Caption         =   "Internet"
      Height          =   195
      Left            =   1620
      TabIndex        =   34
      Top             =   4410
      Width           =   540
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      Caption         =   "FAX"
      Height          =   195
      Left            =   225
      TabIndex        =   33
      Top             =   4410
      Width           =   300
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      Caption         =   "Ramal"
      Height          =   195
      Left            =   5175
      TabIndex        =   32
      Top             =   3780
      Width           =   450
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      Caption         =   "Telefone"
      Height          =   195
      Left            =   3960
      TabIndex        =   31
      Top             =   3780
      Width           =   630
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "DDD"
      Height          =   195
      Left            =   3060
      TabIndex        =   30
      Top             =   3780
      Width           =   360
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "Vendedor"
      Height          =   195
      Left            =   225
      TabIndex        =   29
      Top             =   3780
      Width           =   690
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "Estado"
      Height          =   195
      Left            =   4950
      TabIndex        =   28
      Top             =   3150
      Width           =   495
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Cidade"
      Height          =   195
      Left            =   1575
      TabIndex        =   27
      Top             =   3150
      Width           =   495
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "CEP"
      Height          =   195
      Left            =   225
      TabIndex        =   26
      Top             =   3150
      Width           =   315
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Bairro"
      Height          =   195
      Left            =   3600
      TabIndex        =   25
      Top             =   2520
      Width           =   405
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Logradouro"
      Height          =   195
      Left            =   225
      TabIndex        =   24
      Top             =   2520
      Width           =   810
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Inscrição Estadual"
      Height          =   195
      Left            =   225
      TabIndex        =   23
      Top             =   1935
      Width           =   1305
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Razão Social"
      Height          =   195
      Left            =   225
      TabIndex        =   22
      Top             =   1350
      Width           =   945
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Codinome"
      Height          =   195
      Left            =   225
      TabIndex        =   21
      Top             =   720
      Width           =   705
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "CPF/CGC"
      Height          =   195
      Left            =   2970
      TabIndex        =   3
      Top             =   90
      Width           =   705
   End
   Begin VB.Label lblCodinome 
      AutoSize        =   -1  'True
      Caption         =   "Fornecedores Cadastrados"
      Height          =   195
      Left            =   225
      TabIndex        =   2
      Top             =   90
      Width           =   1905
   End
End
Attribute VB_Name = "frmFornecedor"
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
'| Descrição:  Formulário para cadastro de fornecedores u-|
'|             tilizando como metodologia de acesso e     |
'|             gravação "queries" (consultas)             |
'|             pré-desenvolvidas na base de dados CPR.MDB |
'+--------------------------------------------------------+

Private Sub GravaFornecedor(Operacao As Integer)
' A ROTINA GravaFornecedor, COMO SEU PRÓPRIO NOME INDICA, EFETUA A GRAVAÇÃO _
  DOS DADOS NA TABELA "Fornecedor" DE "CPR.MDB"
  
   Dim Titulo As String ' Variável para configuração do título de janelas de mensagem
   If Operacao = 1 Then
      Titulo = "Inclusão"
   Else
      Titulo = "Alteração"
   End If
   ' Os quatro próximos IF´s efetuam a verificação de controles com preen- _
     chimento obrigatório antes da gravação na base de dados
   If Trim$(mskCPF_CGC) = "" Then
      MsgBox "Informe o número de CPF/CGC", 16, Titulo
      mskCPF_CGC.SetFocus
      Exit Sub
   End If
   If Len(Trim$(mskCPF_CGC)) <> 11 And Len(Trim$(mskCPF_CGC)) <> 15 Then
      MsgBox "Corrija o número de CPF/CGC", 16, Titulo
      mskCPF_CGC.SetFocus
      Exit Sub
   End If
   If Trim$(txtRazaoSocial) = "" Then
      MsgBox "Informe a razão social do Fornecedor", 16, Titulo
      txtRazaoSocial.SetFocus
      Exit Sub
   End If
   If Trim$(txtCodinome) = "" Then
      MsgBox "Informe o codinome (apelido) do Fornecedor", 16, Titulo
      txtCodinome.SetFocus
      Exit Sub
   End If
   'O Visual Basic permite declarar variáveis de memória em qualquer ponto _
    do código:
   Dim MensErro As String, ErroPar As Integer
   ' Com a declaração "ON ERROR.." e as variáveis "ErroPar" e "MensErro" _
     tratamos quaisquer erros que porventura ocorram na execução a partir _
     deste ponto do código
   On Error GoTo ErrGravaFornecedor  'Quando ocorrer um erro, o controle será desviado _
'                                   para o rótulo ErrGravaFornecedor
   ErroPar = False
   MensErro = "Falha na abertura da consulta"
   If Operacao = 1 Then
      Set Consulta = Banco.QueryDefs("InsFornecedor")
   Else
      Set Consulta = Banco.QueryDefs("UpdFornecedor")
   End If
   ErroPar = True
   ' Início da transação de gravação dos dados
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
   Consulta("parVendedor") = txtVendedor
   Consulta("parDDD") = txtDDD
   Consulta("parTelefone") = mskTelefone
   Consulta("parRamal") = txtRamal
   Consulta("parFAX") = mskFAX
   Consulta("parInternet") = txtInternet
   Consulta("parCondicoesPagamento") = txtCondicoesPagamento
   Consulta("parCondicoesDesconto") = txtCondicoesDesconto
   Consulta("parPrazoEntrega") = txtPrazoEntrega
   
   MensErro = "Falha na gravação dos dados"
   ' Execução da consulta após a alimentação de sxeus parâmetros
   Consulta.Execute
   ' Efetivação da transação
   Area.CommitTrans
   ' Após a gravação, o combobox "cmbCodinome" é devidamente atualizado _
     invocando-se a rotina "ComboFornecedor"
   ComboFornecedor cmbCodinome
   Exit Sub ' Este comando evita que o bloco de tratamento de erros seja _
'             executado, pois, uma vez que o controle chegou até aqui, _
'             nenhum erro ocorreu na gravação dos dados
   
' O rótulo trata o erro ocorrido, emite uma mensagem para o usuário _
  e desfaz a transação (ROLLBACK) de gravação dos dados
ErrGravaFornecedor:
   MsgBox MensErro, 16, "Atualização de Fornecedor"
   If ErroPar Then Area.Rollback
   Exit Sub
End Sub
Private Sub cmbCodinome_Click()
   ' Montagem da declaração SQL para leitura do fornecedor selecionado no combo
   Dim Selecao As String
   Selecao = "select * from Fornecedor where Codinome=" & """"
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
   If Not IsNull(Tabela("Vendedor")) Then txtVendedor = Tabela("Vendedor")
   If Not IsNull(Tabela("DDD")) Then txtDDD = Tabela("DDD")
   If Not IsNull(Tabela("Telefone")) Then mskTelefone = Tabela("Telefone")
   If Not IsNull(Tabela("Ramal")) Then txtRamal = Tabela("Ramal")
   If Not IsNull(Tabela("FAX")) Then mskFAX = Tabela("FAX")
   If Not IsNull(Tabela("Internet")) Then txtInternet = Tabela("Internet")
   If Not IsNull(Tabela("CondicoesPagamento")) Then txtCondicoesPagamento = Tabela("CondicoesPagamento")
   If Not IsNull(Tabela("CondicoesDesconto")) Then txtCondicoesDesconto = Tabela("CondicoesDesconto")
   If Not IsNull(Tabela("PrazoEntrega")) Then txtPrazoEntrega = Tabela("PrazoEntrega")
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
   ' Invoca a rotina GRAVAFornecedor passando como argumento o valor 2: _
     qualquer valor diferente de 1 corresponde à operação de alteração
   GravaFornecedor 2
   ' Efetuada a gravação o foco direcionado para o combobox CMBCODINOME
   cmbCodinome.SetFocus
End Sub
Private Sub cmdAlteracao_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   mdiCPR.barMensagem.Panels(1).Text = "Altera os dados do Fornecedor com os dados informados"
End Sub
Private Sub cmdExclusao_Click()
   'Para excluir um Fornecedor, é necessário que o usuário informe o CGC/CPF: _
    isso pode ser feito clicando-se em um dos codinomes do combo CMBCODINOME, _
    que edita os dados para os controles, entre eles o CGC/CPF
   If Trim$(mskCPF_CGC) = "" Then
      MsgBox "Informe o número de CPF/CGC", 16, "Exclusão"
      mskCPF_CGC.SetFocus
      Exit Sub
   End If
   If MsgBox("Confirme a exclusão do Fornecedor", 36, "Exclusão") <> 6 Then
      Exit Sub
   End If
   On Error GoTo ErrExcluiFornecedor
   ' Invoca a consulta DELFORNECEDOR, previamente escrita e gravada em CPR.MDB.
   Set Consulta = Banco.QueryDefs("DelFornecedor")
   Consulta("parCPF_CGC") = mskCPF_CGC.Text
   Consulta.Execute
   ComboFornecedor cmbCodinome
   Exit Sub ' Evita que o bloco para tratamento de erros seja executado quando _
'             a exclusão foi bem sucedida
   
ErrExcluiFornecedor:
   MsgBox MensErro, 16, "Falha na exclusão do Fornecedor"
   Exit Sub
End Sub
Private Sub cmdExclusao_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   ' Explica a finalidade do botão escrevendo uma mensagem na barra de status _
     do formulário principal (mdiCPR)
   mdiCPR.barMensagem.Panels(1).Text = "Exclui o Fornecedor selecionado da base de dados"
End Sub
Private Sub cmdHistorico_Click()
   ' Invoca o formulário de históricos de fornecedores
   frmHistoricoFornecedor.Show
End Sub
Private Sub cmdInclusao_Click()
   ' Invoca a rotina GRAVAFORNECEDOR passando como argumento o valor 1, _
     correspondente à operação de inclusão
   GravaFornecedor 1
   ' Efetuada a gravação o foco direcionado para o combobox CMBCODINOME
   cmbCodinome.SetFocus
End Sub
Private Sub cmdInclusao_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   mdiCPR.barMensagem.Panels(1).Text = "Inclui novo Fornecedor com os dados informados"
End Sub
Private Sub cmdSaida_Click()
   Unload Me
End Sub
Private Sub cmdSaida_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   mdiCPR.barMensagem.Panels(1).Text = "Fecha o cadastro de Fornecedors e retorna à janela principal"
End Sub
Private Sub Form_Load()
   ' INVOCA A ROTINA ComboFornecedor PARA PREENCHER O COMBOBOX cmbCodinome _
     COM OS CODINOMES DE FornecedorS JÁ EXISTENTES NA TABELA "FORNECEDOR" DA _
     BASE DE DADOS "CPR.MDB"
   ComboFornecedor cmbCodinome
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   mdiCPR.barMensagem.Panels(1).Text = ""
End Sub


Private Sub lblCodinome_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   mdiCPR.barMensagem.Panels(1).Text = "Lista de Fornecedors já cadastrados no sistema"
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
   ' Note que o usuário pode editar os dados de um fornecedor através de um _
     clique no combobox CMBCODINOME ou informando o CPF/CGC no controle _
     mskCPF_CGC.
   Set Consulta = Banco.QueryDefs("SelFornecedor")
   Consulta("parCPF_CGC") = temp
   Set Tabela = Consulta.OpenRecordset()
   ' Caso o CPF/CGC não exista, o recordset é fechado e a rotina interrompida
   If Tabela.RecordCount = 0 Then
      Tabela.Close
      Exit Sub
   End If
   ' Copia os dados para os controles exatamente como em cmbCodinome_Click
   txtCodinome = Tabela("Codinome")
   txtRazaoSocial = Tabela("RazaoSocial")
   txtInscricaoEstadual = Tabela("InscricaoEstadual")
   txtLogradouro = Tabela("Logradouro")
   txtBairro = Tabela("Bairro")
   mskCEP = Tabela("CEP")
   txtCidade = Tabela("Cidade")
   txtEstado = Tabela("Estado")
   txtVendedor = Tabela("Contato")
   txtDDD = Tabela("DDD")
   mskTelefone = Tabela("Telefone")
   txtRamal = Tabela("Ramal")
   mskFAX = Tabela("FAX")
   txtInternet = Tabela("Internet")
   Tabela.Close
   
End Sub
Private Sub txtCodinome_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   mdiCPR.barMensagem.Panels(1).Text = "Codinome (""apelido"") do Fornecedor"
End Sub

Private Sub txtVendedor_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   mdiCPR.barMensagem.Panels(1).Text = "Nome da pessoa de contato no Fornecedor"
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
   mdiCPR.barMensagem.Panels(1).Text = "Inscrição estadual do Fornecedor"
End Sub
Private Sub txtInternet_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   mdiCPR.barMensagem.Panels(1).Text = "Endereço do clinte na Internet"
End Sub
Private Sub txtLogradouro_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   mdiCPR.barMensagem.Panels(1).Text = "Rua, avenida, praça, etc."
End Sub
Private Sub txtRazaoSocial_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   mdiCPR.barMensagem.Panels(1).Text = "Razão social do Fornecedor"
End Sub


