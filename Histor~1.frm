VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmHistoricoCliente 
   Caption         =   "Históricos de Clientes"
   ClientHeight    =   4455
   ClientLeft      =   255
   ClientTop       =   1530
   ClientWidth     =   6255
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4455
   ScaleWidth      =   6255
   Begin VB.CommandButton cmdInclusao 
      Caption         =   "&Incluir"
      Height          =   420
      Left            =   4860
      TabIndex        =   12
      Top             =   315
      Width           =   1230
   End
   Begin VB.CommandButton cmdAlteracao 
      Caption         =   "&Alterar"
      Height          =   420
      Left            =   4860
      TabIndex        =   13
      Top             =   945
      Width           =   1230
   End
   Begin VB.CommandButton cmdExclusao 
      Caption         =   "&Excluir"
      Height          =   420
      Left            =   4860
      TabIndex        =   14
      Top             =   1665
      Width           =   1230
   End
   Begin VB.CommandButton cmdSaida 
      Caption         =   "&Retornar"
      Height          =   420
      Left            =   4860
      TabIndex        =   15
      Top             =   3510
      Width           =   1230
   End
   Begin VB.TextBox txtDescricao 
      Height          =   2130
      Left            =   180
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   11
      Top             =   2205
      Width           =   4380
   End
   Begin VB.TextBox txtAssunto 
      Height          =   285
      Left            =   180
      MaxLength       =   20
      TabIndex        =   9
      Top             =   1575
      Width           =   3255
   End
   Begin VB.ComboBox cmbOrdem 
      Height          =   315
      Left            =   180
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   900
      Width           =   915
   End
   Begin VB.ComboBox cmbCodinome 
      Height          =   315
      Left            =   180
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   315
      Width           =   2490
   End
   Begin MSMask.MaskEdBox mskData 
      Height          =   330
      Left            =   1215
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
   Begin MSMask.MaskEdBox mskCPF_CGC 
      Height          =   330
      Left            =   2760
      TabIndex        =   3
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
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Data"
      Height          =   195
      Left            =   1260
      TabIndex        =   6
      Top             =   720
      Width           =   345
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Descrição"
      Height          =   195
      Left            =   180
      TabIndex        =   10
      Top             =   1980
      Width           =   720
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Assunto"
      Height          =   195
      Left            =   180
      TabIndex        =   8
      Top             =   1350
      Width           =   570
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Ordem"
      Height          =   195
      Left            =   180
      TabIndex        =   4
      Top             =   720
      Width           =   465
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Clientes Cadastrados"
      Height          =   195
      Left            =   180
      TabIndex        =   0
      Top             =   90
      Width           =   1485
   End
   Begin VB.Label Label2 
      Caption         =   "CPF/CGC"
      Height          =   195
      Left            =   2805
      TabIndex        =   2
      Top             =   90
      Width           =   1230
   End
End
Attribute VB_Name = "frmHistoricoCliente"
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
'| Descrição:  Formulário para cadastro de históricos  de |
'|             clientes                                   |
'+--------------------------------------------------------+

Private Sub cmbCodinome_Click()
   Dim Selecao As String
   Selecao = "select * from Cliente where Codinome=" & """"
   Selecao = Selecao & cmbCodinome.Text & """"
   Set Tabela = Banco.OpenRecordset(Selecao, dbOpenSnapshot)
   ' Configuração da máscara a partir da coluna "Tabela("CPF_CGC")"
   Select Case Len(Tabela("CPF_CGC"))
      Case Is = 11
         mskCPF_CGC.Mask = "###.###.###-##"
      Case Is = 14
         mskCPF_CGC.Mask = "##.###.###/####-##"
   End Select
   mskCPF_CGC.Text = Tabela("CPF_CGC")
   Tabela.Close
  
   'Executa seleção na tabela de históricos a partir do CPF/CGC do _
    cliente selecionado (utilizando o mesmo objeto recordset)
   Selecao = "select * from HistoricoCliente where CPF_CGC="
   Selecao = Selecao & mskCPF_CGC
   
   Set Tabela = Banco.OpenRecordset(Selecao, dbOpenSnapshot)
   ' Limpa o combo CMBORDEM  e alimenta-o com todos os números de ordem _
     dos históricos do cliente selecionados
   cmbOrdem.Clear
   Do Until Tabela.EOF
      cmbOrdem.AddItem Tabela("Ordem")
      Tabela.MoveNext
   Loop
   ' O método CLOSE fecha o recordset, evitando que ele fique aberto após _
     sua utilização
   Tabela.Close


End Sub

Private Sub cmbOrdem_Click()
   Selecao = "SELECT * From HistoricoCliente WHERE CPF_CGC = "
   Selecao = Selecao & mskCPF_CGC
   Selecao = Selecao & " and Ordem = " & cmbOrdem
   
   Set Tabela = Banco.OpenRecordset(Selecao, dbOpenSnapshot)
   If Tabela.RecordCount = 0 Then
      Tabela.Close
      Exit Sub
   End If
   mskData = Tabela("data")
   txtAssunto = Tabela("Assunto")
   txtDescricao = Tabela("Descricao")
   Tabela.Close
End Sub

Private Sub cmdAlteracao_Click()
   GravaHistoricoCliente 2
   cmbOrdem.SetFocus
End Sub

Private Sub cmdExclusao_Click()
   If Trim$(mskCPF_CGC) = "" Then
      MsgBox "Informe o número de CPF/CGC", 16, "Exclusão"
      mskCPF_CGC.SetFocus
      Exit Sub
   End If
   If cmbOrdem.ListIndex = -1 Then
      MsgBox "Informe o número de ordem", 16, "Exclusão"
      cmbOrdem.SetFocus
      Exit Sub
   End If
   If MsgBox("Confirme a exclusão do histórico", 36, "Exclusão") <> 6 Then
      Exit Sub
   End If
   On Error GoTo ErrExcluiHistoricoCliente
   Set Consulta = Banco.QueryDefs("DelHistoricoCliente")
   Consulta("parCPF_CGC") = mskCPF_CGC.Text
   Consulta("parOrdem") = cmbOrdem.Text
   Consulta.Execute
   mskCPF_CGC.Text = ""
   mskData = ""
   txtAssunto = ""
   txtDescricao = ""
   cmbCodinome_Click
   Exit Sub
   
ErrExcluiHistoricoCliente:
   MsgBox MensErro, 16, "Falha na exclusão do histórico"
   Exit Sub
End Sub

Private Sub cmdInclusao_Click()
   GravaHistoricoCliente 1
   cmbOrdem.SetFocus
End Sub

Private Sub cmdSaida_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   ' A rotina ComboCliente foi invocada na leitura do formulário-chamador _
     (frmCliente) e, aqui, uma segunda vez...
   ' Note que em FRMCLIENTE também possuímos um combobox denominado CMBCODINOME: _
     esses dois controles possuem comportamento independente!
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
End Sub
Private Sub mskData_LostFocus()
   ' Em vez de tratar a data através do evento "ValidationError", _
     preferimos ilustrar aqui a utilização de uma função genérica construída _
     para a consistência de datas.
   If Not ConsisteData(mskData) Then mskData.SetFocus
End Sub
Private Sub GravaHistoricoCliente(Operacao As Integer)
   Dim Titulo As String
   If Operacao = 1 Then
      Titulo = "Inclusão"
   Else
      Titulo = "Alteração"
   End If
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
   If Trim$(mskData) = "" Then
      MsgBox "Informe a data do histórico", 16, Titulo
      mskData.SetFocus
      Exit Sub
   End If
   If Operacao <> 1 And cmbOrdem.ListIndex = -1 Then
      MsgBox "Informe o número de ordem", 16, Titulo
      cmbOrdem.SetFocus
      Exit Sub
   End If
   Dim MensErro As String, ErroPar As Integer
   On Error GoTo ErrGravaHistoricoCliente
   ErroPar = False
   MensErro = "Falha na abertura da consulta"
   If Operacao = 1 Then
      Set Consulta = Banco.QueryDefs("InsHistoricoCliente")
   Else
      Set Consulta = Banco.QueryDefs("UpdHistoricoCliente")
   End If
   ErroPar = True

   Area.BeginTrans
   MensErro = "Falha na alimentação dos parâmetros"
   If Operacao <> 1 Then
      Consulta("parOrdem") = cmbOrdem.Text
   End If
   Consulta("parCPF_CGC") = mskCPF_CGC.Text
   Consulta("parData") = mskData
   Consulta("parAssunto") = txtAssunto
   Consulta("parDescricao") = txtDescricao
   
   MensErro = "Falha na gravação dos dados"

   Consulta.Execute

   Area.CommitTrans
   cmbCodinome_Click
   cmbCodinome.SetFocus
   Exit Sub

ErrGravaHistoricoCliente:
   MsgBox MensErro, 16, "Atualização de Histórico de Cliente"
   If ErroPar Then Area.Rollback
   Exit Sub
End Sub


