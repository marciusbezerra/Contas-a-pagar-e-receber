VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "COMCTL32.OCX"
Begin VB.MDIForm mdiCPR 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Sistema de Contas a Pagar e a Receber"
   ClientHeight    =   5220
   ClientLeft      =   300
   ClientTop       =   870
   ClientWidth     =   9480
   LinkTopic       =   "MDIForm1"
   WindowState     =   2  'Maximized
   Begin ComctlLib.StatusBar barMensagem 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   4905
      Width           =   9480
      _ExtentX        =   16722
      _ExtentY        =   556
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   6
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   5292
            MinWidth        =   5292
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   1
            Enabled         =   0   'False
            Object.Width           =   1235
            MinWidth        =   1235
            TextSave        =   "CAPS"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   2
            Object.Width           =   1235
            MinWidth        =   1235
            TextSave        =   "NUM"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   3
            Object.Width           =   1235
            MinWidth        =   1235
            TextSave        =   "INS"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel5 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   6
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "30/08/99"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel6 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   5
            Object.Width           =   1411
            MinWidth        =   1411
            TextSave        =   "11:45"
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu menCadastros 
      Caption         =   "&Cadastros"
      Begin VB.Menu menClientes 
         Caption         =   "&Clientes"
         Shortcut        =   ^C
      End
      Begin VB.Menu menFornecedores 
         Caption         =   "&Fornecedores"
         Shortcut        =   ^F
      End
      Begin VB.Menu menBancos 
         Caption         =   "&Bancos"
         Shortcut        =   ^B
      End
   End
   Begin VB.Menu menContabilidade 
      Caption         =   "C&ontabilidade"
      Begin VB.Menu menPlano 
         Caption         =   "&Plano de Contas"
         Shortcut        =   ^P
      End
      Begin VB.Menu menLancamentos 
         Caption         =   "&Lançamentos"
         Shortcut        =   ^L
      End
   End
   Begin VB.Menu menContasPag 
      Caption         =   "Contas a &Pagar"
      Begin VB.Menu menDespesas 
         Caption         =   "&Despesas"
         Shortcut        =   ^D
      End
      Begin VB.Menu menPagamentos 
         Caption         =   "&Pagamentos"
         Shortcut        =   ^G
      End
   End
   Begin VB.Menu menContasRec 
      Caption         =   "Contas a &Receber"
      Begin VB.Menu menRecebimentos 
         Caption         =   "&Recebimentos"
         Shortcut        =   ^R
      End
      Begin VB.Menu menCartorios 
         Caption         =   "&Cartórios"
         Shortcut        =   ^A
      End
      Begin VB.Menu menBorderos 
         Caption         =   "&Borderôs"
         Shortcut        =   ^O
      End
      Begin VB.Menu menCheques 
         Caption         =   "Cheques &Pré-Datados"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu menSaida 
      Caption         =   "&Encerrar "
   End
End
Attribute VB_Name = "mdiCPR"
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

Private Sub MDIForm_Load()

End Sub

Private Sub menBancos_Click()
   frmBanco.Show
End Sub


Private Sub menBorderos_Click()
frmBordero.Show
End Sub

Private Sub menCartorios_Click()
   frmRemessa.Show
End Sub

Private Sub menCheques_Click()
frmChequePreDatado.Show
End Sub

Private Sub menClientes_Click()
   frmCliente.Show
End Sub


Private Sub menDespesas_Click()
   frmDespesa.Show
End Sub

Private Sub menFornecedores_Click()
   frmFornecedor.Show
End Sub

Private Sub menLancamentos_Click()
   frmLancamento.Show
End Sub

Private Sub menPagamentos_Click()
frmPagamento.Show
End Sub

Private Sub menPlano_Click()
   frmPlanoContas.Show
End Sub

Private Sub menRecebimentos_Click()
   frmRecebimento.Show
End Sub

Private Sub menSaida_Click()
   End
End Sub


