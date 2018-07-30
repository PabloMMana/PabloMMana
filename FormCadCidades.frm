VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Begin VB.Form FormCadCidades 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Cidades"
   ClientHeight    =   5055
   ClientLeft      =   45
   ClientTop       =   675
   ClientWidth     =   7245
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5055
   ScaleWidth      =   7245
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4935
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   6975
      Begin VB.Frame Frame2 
         Height          =   3255
         Left            =   120
         TabIndex        =   5
         Top             =   1560
         Width           =   6735
         Begin MSFlexGridLib.MSFlexGrid MSFlexGridCidade 
            Height          =   2895
            Left            =   120
            TabIndex        =   6
            Top             =   240
            Width           =   6495
            _ExtentX        =   11456
            _ExtentY        =   5106
            _Version        =   393216
            Cols            =   4
         End
      End
      Begin VB.ComboBox CmbUF 
         Height          =   330
         Left            =   120
         TabIndex        =   3
         Top             =   1080
         Width           =   1215
      End
      Begin VB.TextBox TextCidade 
         Height          =   315
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   2895
      End
      Begin VB.Label Label2 
         Caption         =   "UF"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Cidade"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   1575
      End
   End
   Begin VB.Menu MnuCad 
      Caption         =   "&Cadastrar |"
   End
   Begin VB.Menu MnuSair 
      Caption         =   "&Sair"
   End
End
Attribute VB_Name = "FormCadCidades"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cn As New ADODB.Connection

Private Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Declare Function GetSystemMenu Lib "user32" (ByVal hWnd As Long, ByVal bRevert As Long) As Long

Private Const MF_BYPOSITION = &H400&


Private Sub RemoveMenus()
  Dim hMenu As Long
  hMenu = GetSystemMenu(hWnd, False)
  DeleteMenu hMenu, 6, MF_BYPOSITION
End Sub


Private Sub Form_Load()

  Dim hMenu As Long
  hMenu = GetSystemMenu(hWnd, False)
  DeleteMenu hMenu, 6, MF_BYPOSITION

Call StringConexao
Call AJUSTAGRID
Call PRECOMBO
Call Pregrid

End Sub

Private Sub MnuCad_Click()
Call Cadastrar
End Sub

Private Sub MnuSair_Click()
Unload Me
End Sub

Private Sub Cadastrar()
Dim conexao As New ADODB.Command
Dim rsdados As New ADODB.Recordset
Dim ssql As String

cn.Open

ssql = "INSERT INTO TABCIDADE (NOME,IDUF)VALUES ('" & TextCidade.Text & "' , '" & CmbUF.ListIndex & "' +  1 )"


conexao.CommandType = adCmdText
conexao.CommandText = ssql
Set conexao.ActiveConnection = cn

rsdados.Open conexao, , adOpenDynamic, adLockOptimistic

Call Pregrid



End Sub

Public Sub Pregrid()

Dim conexao As New ADODB.Command
Dim rsdados As New ADODB.Recordset
Dim ssql As String


ssql = "select C.IDCIDADE,C.NOME,U.NOME NOMEUF from tabcidade C INNER JOIN tabuf U ON C.IDUF=U.IDUF"


conexao.CommandType = adCmdText
conexao.CommandText = ssql
Set conexao.ActiveConnection = cn

rsdados.Open conexao, , adOpenDynamic, adLockOptimistic


MSFlexGridCidade.Rows = Contalinhas + 1

I = 1

Do While Not rsdados.EOF


MSFlexGridCidade.Row = I
MSFlexGridCidade.Col = 1
MSFlexGridCidade.Text = rsdados!idcidade


MSFlexGridCidade.Row = I
MSFlexGridCidade.Col = 2
MSFlexGridCidade.Text = rsdados!nome

MSFlexGridCidade.Row = I
MSFlexGridCidade.Col = 3
MSFlexGridCidade.Text = rsdados!NOMEUF




I = I + 1
rsdados.MoveNext

Loop

cn.Close

End Sub

Public Function Contalinhas() As String

Dim conexao As New ADODB.Command
Dim rsdados As New ADODB.Recordset
Dim ssql As String

ssql = "SELECT count(idCIDADE) cidade  FROM TABCIDADE"


conexao.CommandType = adCmdText
conexao.CommandText = ssql
Set conexao.ActiveConnection = cn

rsdados.Open conexao, , adOpenDynamic, adLockOptimistic

Contalinhas = rsdados!cidade



End Function

Public Function StringConexao()
'--String de Conexao

cn.ConnectionString = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=CORCLIENTE"

cn.Open

End Function




Private Sub PRECOMBO()

Dim conexao As New ADODB.Command
Dim rsdados As New ADODB.Recordset
Dim ssql As String


ssql = "select * from tabuf "


conexao.CommandType = adCmdText
conexao.CommandText = ssql
Set conexao.ActiveConnection = cn

rsdados.Open conexao, , adOpenDynamic, adLockOptimistic

Do While Not rsdados.EOF

CmbUF.AddItem (rsdados!nome)
rsdados.MoveNext

Loop
End Sub



Public Sub AJUSTAGRID()



'---------Nomes das Colunas
MSFlexGridCidade.TextMatrix(0, 1) = "Código"
MSFlexGridCidade.TextMatrix(0, 2) = "Nome Cidade"
MSFlexGridCidade.TextMatrix(0, 3) = "Estado"


'---------Ajustar tamanho das colunas

MSFlexGridCidade.ColWidth(0) = 0
MSFlexGridCidade.ColWidth(1) = 600
MSFlexGridCidade.ColWidth(2) = 2600
MSFlexGridCidade.ColWidth(3) = 600
End Sub
