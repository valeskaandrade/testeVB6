VERSION 5.00
Begin VB.Form frmTela 
   Caption         =   "Consulta Países"
   ClientHeight    =   8295
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14565
   LinkTopic       =   "Form1"
   ScaleHeight     =   8295
   ScaleWidth      =   14565
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtResultado 
      Height          =   6855
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   960
      Width           =   14175
   End
   Begin VB.CommandButton cmdSalvarDados 
      Caption         =   "Salvar Dados"
      Height          =   495
      Left            =   2280
      TabIndex        =   1
      Top             =   120
      Width           =   2055
   End
   Begin VB.CommandButton cmdBaixarDados 
      Caption         =   "Baixar Dados"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label lblResultado 
      Caption         =   "Resultado:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   855
   End
End
Attribute VB_Name = "frmTela"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBaixarDados_Click()
    On Error GoTo tratarErro
    
    Screen.MousePointer = vbHourglass
    
    cmdSalvarDados.Enabled = BaixarDados
    txtResultado.Text = sTextoTela
    
    Screen.MousePointer = vbDefault
    If Trim(Msg) <> "" Then MsgBox sMsg
    Exit Sub
    
tratarErro:
    Screen.MousePointer = vbDefault
    sMsg = "Erro: " & Err.Number & "-" & Err.Description & "Origem: cmdBaixarDados_Click"
    MsgBox sMsg
    GravarLog sMsg
    
End Sub

Private Sub cmdSalvarDados_Click()
    On Error GoTo tratarErro
     Screen.MousePointer = vbHourglass
    SalvarDados
    Screen.MousePointer = vbDefault
    MsgBox "Fim com sucesso!"
    
    Exit Sub
    
tratarErro:
    Screen.MousePointer = vbDefault
    sMsg = "Erro: " & Err.Number & "-" & Err.Description & "Origem: cmdSalvarDados_Click"
    MsgBox sMsg
    GravarLog sMsg
 
End Sub

Private Sub Form_Load()
    cmdSalvarDados.Enabled = False
End Sub
