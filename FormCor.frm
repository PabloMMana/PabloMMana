VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Begin VB.Form FormCor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Corretores"
   ClientHeight    =   7620
   ClientLeft      =   45
   ClientTop       =   675
   ClientWidth     =   8685
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
   ScaleHeight     =   7620
   ScaleWidth      =   8685
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   4215
      Left            =   120
      TabIndex        =   7
      Top             =   3120
      Width           =   8415
      Begin MSFlexGridLib.MSFlexGrid MSFlexGridcOR 
         Height          =   3735
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   8175
         _ExtentX        =   14420
         _ExtentY        =   6588
         _Version        =   393216
         Cols            =   4
      End
   End
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
      Height          =   2775
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   8415
      Begin MSMask.MaskEdBox textCPF 
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   2040
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   450
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin VB.TextBox TextNome 
         Height          =   315
         Left            =   240
         TabIndex        =   0
         Top             =   1320
         Width           =   5655
      End
      Begin VB.TextBox TextCod 
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         Height          =   315
         Left            =   240
         TabIndex        =   8
         Top             =   600
         Width           =   1935
      End
      Begin VB.Label Label3 
         Caption         =   "C.P.F."
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   1800
         Width           =   1815
      End
      Begin VB.Label Label2 
         Caption         =   "Nome"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   1080
         Width           =   2535
      End
      Begin VB.Label Label1 
         Caption         =   "Código"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Menu MnuCad 
      Caption         =   "&Cadastrar |"
   End
   Begin VB.Menu MnuSair 
      Caption         =   "&Sair"
   End
End
Attribute VB_Name = "FormCor"
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

Call UltimoReg

Call Pregrid

Call CadCPF

End Sub

Public Sub AJUSTAGRID()



'---------Nomes das Colunas
MSFlexGridcOR.TextMatrix(0, 1) = "Cod. Corretor"
MSFlexGridcOR.TextMatrix(0, 2) = "Nome do Corretor"
MSFlexGridcOR.TextMatrix(0, 3) = "C.P.F"


'---------Ajustar tamanho das colunas

MSFlexGridcOR.ColWidth(0) = 0
MSFlexGridcOR.ColWidth(1) = 800
MSFlexGridcOR.ColWidth(2) = 2600
MSFlexGridcOR.ColWidth(3) = 2600



End Sub

Private Sub MnuCad_Click()

If TextNome.Text = "" Then
MsgBox "O nome não pode ser em banco!", vbCritical, "Aviso"

ElseIf textCPF.Text = "___.___.___-__" Then

MsgBox "O CPF não pode ser em banco!", vbCritical, "Aviso"


Else

Dim conexao As New ADODB.Command
Dim rsdados As New ADODB.Recordset
Dim ssql As String

    ssql = "INSERT INTO TABCORRETORES (NOME,CPF) VALUES ('" & TextNome.Text & "','" & textCPF.Text & "')"

conexao.CommandType = adCmdText
conexao.CommandText = ssql
Set conexao.ActiveConnection = cn
rsdados.Open conexao, , adOpenDynamic, adLockOptimistic

Call Limpa

Call Pregrid


End If
End Sub

Public Sub Limpa()


Call UltimoReg
textCPF.Text = "___.___.___-__"
TextNome.Text = ""


End Sub

Private Sub MnuSair_Click()
cn.Close
Unload Me
End Sub

Public Sub CadCPF()

textCPF.Mask = "###.###.###-##"

End Sub



Public Function StringConexao()
'--String de Conexao
cn.ConnectionString = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=CORCLIENTE"

cn.Open

End Function


Public Function Pregrid()

' PREENCHE GRID COM CORRETORES

Dim conexao As New ADODB.Command
Dim rsdados As New ADODB.Recordset
Dim ssql As String

ssql = "SELECT IDCORRETOR,NOME,CPF FROM TABCORRETORES"


conexao.CommandType = adCmdText
conexao.CommandText = ssql
Set conexao.ActiveConnection = cn

rsdados.Open conexao, , adOpenDynamic, adLockOptimistic

MSFlexGridcOR.Rows = Contalinhas + 1

I = 1
Do While Not rsdados.EOF

MSFlexGridcOR.Row = I
MSFlexGridcOR.Col = 1
MSFlexGridcOR.Text = rsdados!idcorretor

MSFlexGridcOR.Row = I
MSFlexGridcOR.Col = 2
MSFlexGridcOR.Text = rsdados!nome

MSFlexGridcOR.Row = I
MSFlexGridcOR.Col = 3
MSFlexGridcOR.Text = rsdados!cpf

I = I + 1
rsdados.MoveNext
Loop


End Function

Public Function Contalinhas() As String

Dim conexao As New ADODB.Command
Dim rsdados As New ADODB.Recordset
Dim ssql As String

ssql = "SELECT count(idcorretor) total FROM TABCORRETORES"


conexao.CommandType = adCmdText
conexao.CommandText = ssql
Set conexao.ActiveConnection = cn

rsdados.Open conexao, , adOpenDynamic, adLockOptimistic

Contalinhas = rsdados!total

End Function



Private Sub UltimoReg()

Dim conexao As New ADODB.Command
Dim rsdados As New ADODB.Recordset
Dim ssql As String


ssql = "SELECT max(idcorretor) idcorretor  FROM TABCORRETORES"


conexao.CommandType = adCmdText
conexao.CommandText = ssql
Set conexao.ActiveConnection = cn

rsdados.Open conexao, , adOpenDynamic, adLockOptimistic

If IsNull(rsdados!idcorretor) Then
TextCod.Text = 1
Else
TextCod.Text = rsdados!idcorretor + 1
End If

End Sub

Private Sub MSFlexGridcOR_DblClick()
num = MSFlexGridcOR.TextMatrix(MSFlexGridcOR.RowSel, 1)
Dim conexao As New ADODB.Command
Dim rsdados As New ADODB.Recordset
Dim ssql As String

If MsgBox("Deseja realmente excluir esse corretor?", vbYesNo, "Aviso") = vbYes Then

ssql = "DELETE FROM TABCORRETORES WHERE IDCORRETOR='" & num & "'"


conexao.CommandType = adCmdText
conexao.CommandText = ssql
Set conexao.ActiveConnection = cn

rsdados.Open conexao, , adOpenDynamic, adLockOptimistic
Call Pregrid
Call UltimoReg
cn.Close
Else
Call Pregrid
cn.Close
End If

cn.Open
End Sub
