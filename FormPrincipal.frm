VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Begin VB.Form FormPrincipal 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Cliente/Corretores"
   ClientHeight    =   8550
   ClientLeft      =   150
   ClientTop       =   480
   ClientWidth     =   12135
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
   ScaleHeight     =   8550
   ScaleWidth      =   12135
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TextNomeCli 
      Height          =   315
      Left            =   360
      TabIndex        =   3
      Top             =   2280
      Width           =   5895
   End
   Begin VB.TextBox TextNomeCor 
      Height          =   315
      Left            =   360
      TabIndex        =   2
      Top             =   1560
      Width           =   5895
   End
   Begin VB.TextBox TextCodCor 
      Height          =   315
      Left            =   360
      TabIndex        =   1
      Top             =   840
      Width           =   3015
   End
   Begin VB.Frame Frame1 
      Height          =   8295
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11895
      Begin VB.CommandButton ComdPesq 
         Caption         =   "Pesquisar"
         Height          =   255
         Left            =   7200
         TabIndex        =   15
         Top             =   2760
         Width           =   1455
      End
      Begin VB.ComboBox CmbCidade 
         Height          =   330
         Left            =   7200
         TabIndex        =   12
         Top             =   2160
         Width           =   3375
      End
      Begin VB.ComboBox CmbEstado 
         Height          =   330
         Left            =   7200
         TabIndex        =   11
         Top             =   1320
         Width           =   3375
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   255
         Left            =   7200
         TabIndex        =   10
         Top             =   720
         Width           =   255
      End
      Begin VB.Frame Frame2 
         Height          =   4935
         Left            =   240
         TabIndex        =   9
         Top             =   3360
         Width           =   11535
         Begin VB.CommandButton ComdExclu 
            Caption         =   "Desativar"
            Height          =   255
            Left            =   10440
            TabIndex        =   14
            Top             =   4440
            Width           =   975
         End
         Begin MSFlexGridLib.MSFlexGrid MSFlexGridDet 
            Height          =   4095
            Left            =   120
            TabIndex        =   13
            Top             =   240
            Width           =   11295
            _ExtentX        =   19923
            _ExtentY        =   7223
            _Version        =   393216
         End
      End
      Begin MSMask.MaskEdBox TextMasCPF 
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   2880
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   450
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin VB.Label Label7 
         Caption         =   "UF"
         Height          =   255
         Left            =   7200
         TabIndex        =   18
         Top             =   1920
         Width           =   1695
      End
      Begin VB.Label Label6 
         Caption         =   "Cidade"
         Height          =   255
         Left            =   7200
         TabIndex        =   17
         Top             =   1080
         Width           =   2415
      End
      Begin VB.Label Label5 
         Caption         =   "Ativo?"
         Height          =   255
         Left            =   7200
         TabIndex        =   16
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "C.P.F. Cliente"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   2640
         Width           =   2295
      End
      Begin VB.Label Label3 
         Caption         =   "Nome Cliente"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   1920
         Width           =   1935
      End
      Begin VB.Label Label2 
         Caption         =   "Nome Corretor"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   1200
         Width           =   2175
      End
      Begin VB.Label Label1 
         Caption         =   "Código Corretor"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   480
         Width           =   1575
      End
   End
   Begin VB.Menu MnuCorretor 
      Caption         =   "&Cadastrar Corretor |"
   End
   Begin VB.Menu MnuCliente 
      Caption         =   "&Cadastrar Clientes |"
   End
   Begin VB.Menu mnuCadCid 
      Caption         =   "&Cadastro de Cidades |"
   End
   Begin VB.Menu MnuSair 
      Caption         =   "&Sair"
   End
End
Attribute VB_Name = "FormPrincipal"
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

Private Sub ComdExclu_Click()
INICIAL = 1
FINAL = MSFlexGridDet.Rows

For I = INICIAL To FINAL - 1

If MSFlexGridDet.TextMatrix(I, 8) = "X" Then

   cpf = MSFlexGridDet.TextMatrix(I, 2)
   Desativar (cpf)
    
    End If
Next I


Call ComdPesq_Click

End Sub

Public Function Desativar(cpf As String)

Dim conexao As New ADODB.Command
Dim rsdados As New ADODB.Recordset
Dim ssql As String

ssql = "UPDATE  TABCLIENTES  SET ATIVO=1 WHERE CPF='" & cpf & "'"


conexao.CommandType = adCmdText
conexao.CommandText = ssql
Set conexao.ActiveConnection = cn

rsdados.Open conexao, , adOpenDynamic, adLockOptimistic

End Function


Private Sub ComdPesq_Click()




Dim conexao As New ADODB.Command
Dim rsdados As New ADODB.Recordset
Dim ssql As String


If Check1.Value = 1 Then
ATIVO = 0
Else
ATIVO = 1
End If


ssql = "SELECT C.NOME NOMECLI ,C.CPF ,CASE C.ATIVO WHEN 0 THEN 'Sim' else 'Não' end Ativo, CO.NOME NOMECOR,CO.IDCORRETOR," _
& " CI.NOME NOMECI,U.NOME NOMEUF" _
& " FROM TABCLICOR T" _
& " INNER JOIN TABCLIENTES C ON T.IDCLIENTE=C.IDCLIENTE" _
& " INNER JOIN TABCORRETORES CO ON T.IDCORRETOR=CO.IDCORRETOR" _
& " INNER JOIN TABCIDADE CI ON C.IDCIDADE=CI.IDCIDADE" _
& " INNER JOIN TABUF U ON CI.IDUF=U.IDUF" _
& " WHERE C.CPF='" & TextMasCPF.Text & "'" _
& " OR C.NOME = '" & TextNomeCli.Text & "'" _
& " OR C.ATIVO=" & ATIVO & "" _
& " OR CO.NOME='" & TextNomeCor.Text & "'" _
& " OR CI.NOME= '" & Cmbcidade.Text & "'" _
& " OR U.NOME='" & CmbEstado.Text & "' "





conexao.CommandType = adCmdText
conexao.CommandText = ssql
Set conexao.ActiveConnection = cn

rsdados.Open conexao, , adOpenDynamic, adLockOptimistic



MSFlexGridDet.Rows = ContalinhaspESQ + 1

I = 1

Do While Not rsdados.EOF

MSFlexGridDet.Row = I
MSFlexGridDet.Col = 1
MSFlexGridDet.Text = rsdados!nomeCLI

MSFlexGridDet.Row = I
MSFlexGridDet.Col = 2
MSFlexGridDet.Text = rsdados!cpf

MSFlexGridDet.Row = I
MSFlexGridDet.Col = 3
MSFlexGridDet.Text = rsdados!ATIVO
MSFlexGridDet.Row = I
MSFlexGridDet.Col = 4
MSFlexGridDet.Text = rsdados!nomecor
MSFlexGridDet.Row = I
MSFlexGridDet.Col = 5
MSFlexGridDet.Text = rsdados!idcorretor
MSFlexGridDet.Row = I
MSFlexGridDet.Col = 6
MSFlexGridDet.Text = rsdados!NOMEUF
MSFlexGridDet.Row = I
MSFlexGridDet.Col = 7
MSFlexGridDet.Text = rsdados!NOMECI

I = I + 1
rsdados.MoveNext

Loop


End Sub



Private Sub Form_Load()

   Dim hMenu As Long
  hMenu = GetSystemMenu(hWnd, False)
  DeleteMenu hMenu, 6, MF_BYPOSITION
 Call RemoveMenus
 Call Maskcpf
Call AJUSTAGRID
Call StringConexao
Call Pregrid

End Sub

Private Sub Maskcpf()
TextMasCPF.Mask = "###.###.###-##"
End Sub

Public Sub AJUSTAGRID()

MSFlexGridDet.Cols = 9

'---------Nomes das Colunas
MSFlexGridDet.TextMatrix(0, 1) = "Nome do Cliente"
MSFlexGridDet.TextMatrix(0, 2) = "CPF"
MSFlexGridDet.TextMatrix(0, 3) = "Ativo?"
MSFlexGridDet.TextMatrix(0, 4) = "Nome do Corretor"
MSFlexGridDet.TextMatrix(0, 5) = "Cod. do Corretor"
MSFlexGridDet.TextMatrix(0, 6) = "U.F."
MSFlexGridDet.TextMatrix(0, 7) = "Cidade"
MSFlexGridDet.TextMatrix(0, 8) = ""

'---------Ajustar tamanho das colunas

MSFlexGridDet.ColWidth(0) = 0
MSFlexGridDet.ColWidth(1) = 2200
MSFlexGridDet.ColWidth(2) = 1600
MSFlexGridDet.ColWidth(3) = 600
MSFlexGridDet.ColWidth(4) = 2000
MSFlexGridDet.ColWidth(5) = 1600
MSFlexGridDet.ColWidth(6) = 600
MSFlexGridDet.ColWidth(7) = 2200
MSFlexGridDet.ColWidth(8) = 200

End Sub

Private Sub mnuCadCid_Click()
FormCadCidades.Show
End Sub

Private Sub MnuCadEst_Click()

End Sub

Private Sub MnuCliente_Click()
FormClientes.Show
End Sub

Private Sub MnuCorretor_Click()
FormCor.Show
End Sub

Private Sub MnuSair_Click()
' Sair do aplicativo
If MsgBox("Deseja realmente sair do sistema?", vbYesNo, "Aviso") = vbYes Then
Unload Me
End If
End Sub

Private Sub MSFlexGridDet_Click()
'--------Seleciona linha para exclusão

If MSFlexGridDet.TextMatrix(MSFlexGridDet.Row, 8) = "X" Then

MSFlexGridDet.TextMatrix(MSFlexGridDet.Row, 8) = ""
Else

MSFlexGridDet.TextMatrix(MSFlexGridDet.Row, 8) = "X"

End If
End Sub

Public Function StringConexao()
'--String de Conexao
cn.ConnectionString = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=CORCLIENTE"
cn.Open

End Function


Public Sub Pregrid()


Dim conexao As New ADODB.Command
Dim rsdados As New ADODB.Recordset
Dim ssql As String



ssql = "SELECT C.NOME NOMECLI ,C.CPF ,CASE C.ATIVO WHEN 0 THEN 'Sim' else 'Não' end Ativo, CO.NOME NOMECOR,CO.IDCORRETOR," _
& " CI.NOME NOMECI,U.NOME NOMEUF" _
& " FROM TABCLICOR T" _
& " INNER JOIN TABCLIENTES C ON T.IDCLIENTE=C.IDCLIENTE" _
& " INNER JOIN TABCORRETORES CO ON T.IDCORRETOR=CO.IDCORRETOR" _
& " INNER JOIN TABCIDADE CI ON C.IDCIDADE=CI.IDCIDADE" _
& " INNER JOIN TABUF U ON CI.IDUF=U.IDUF"




conexao.CommandType = adCmdText
conexao.CommandText = ssql
Set conexao.ActiveConnection = cn

rsdados.Open conexao, , adOpenDynamic, adLockOptimistic

MSFlexGridDet.Rows = Contalinhas + 1

I = 1

Do While Not rsdados.EOF

MSFlexGridDet.Row = I
MSFlexGridDet.Col = 1
MSFlexGridDet.Text = rsdados!nomeCLI
MSFlexGridDet.Row = I
MSFlexGridDet.Col = 2
MSFlexGridDet.Text = rsdados!cpf
MSFlexGridDet.Row = I
MSFlexGridDet.Col = 3
MSFlexGridDet.Text = rsdados!ATIVO
MSFlexGridDet.Row = I
MSFlexGridDet.Col = 4
MSFlexGridDet.Text = rsdados!nomecor
MSFlexGridDet.Row = I
MSFlexGridDet.Col = 5
MSFlexGridDet.Text = rsdados!idcorretor
MSFlexGridDet.Row = I
MSFlexGridDet.Col = 6
MSFlexGridDet.Text = rsdados!NOMEUF
MSFlexGridDet.Row = I
MSFlexGridDet.Col = 7
MSFlexGridDet.Text = rsdados!NOMECI

I = I + 1
rsdados.MoveNext

Loop


End Sub

Public Function Contalinhas() As String

Dim conexao As New ADODB.Command
Dim rsdados As New ADODB.Recordset
Dim ssql As String

ssql = "SELECT count(C.IDCLIENTE) AS NOME" _
& " FROM TABCLICOR T" _
& " INNER JOIN TABCLIENTES C ON T.IDCLIENTE=C.IDCLIENTE" _
& " INNER JOIN TABCORRETORES CO ON T.IDCORRETOR=CO.IDCORRETOR" _
& " INNER JOIN TABCIDADE CI ON C.IDCIDADE=CI.IDCIDADE" _
& " INNER JOIN TABUF U ON CI.IDUF=U.IDUF"



conexao.CommandType = adCmdText
conexao.CommandText = ssql
Set conexao.ActiveConnection = cn

rsdados.Open conexao, , adOpenDynamic, adLockOptimistic

If IsNull(rsdados!nome) Then
Contalinhas = 0
Else

Contalinhas = rsdados!nome

End If
End Function

Public Function ContalinhaspESQ() As String

If Check1.Value = 1 Then
ATIVO = 0
Else
ATIVO = 1
End If

Dim conexao As New ADODB.Command
Dim rsdados As New ADODB.Recordset
Dim ssql As String

ssql = "SELECT count(C.IDCLIENTE) AS NOME" _
& " FROM TABCLICOR T" _
& " INNER JOIN TABCLIENTES C ON T.IDCLIENTE=C.IDCLIENTE" _
& " INNER JOIN TABCORRETORES CO ON T.IDCORRETOR=CO.IDCORRETOR" _
& " INNER JOIN TABCIDADE CI ON C.IDCIDADE=CI.IDCIDADE" _
& " INNER JOIN TABUF U ON CI.IDUF=U.IDUF" _
& " WHERE C.CPF='" & TextMasCPF.Text & "'" _
& " OR C.NOME = '" & TextNomeCli.Text & "'" _
& " OR C.ATIVO=" & ATIVO & "" _
& " OR CO.NOME='" & TextNomeCor.Text & "'" _
& " OR CI.NOME= '" & Cmbcidade.Text & "'" _
& " OR U.NOME='" & CmbEstado.Text & "' "


conexao.CommandType = adCmdText
conexao.CommandText = ssql
Set conexao.ActiveConnection = cn

rsdados.Open conexao, , adOpenDynamic, adLockOptimistic

If IsNull(rsdados!nome) Or (rsdados.EOF And rsdados.BOF) Then
ContalinhaspESQ = 0
Else

ContalinhaspESQ = rsdados!nome

End If

End Function

