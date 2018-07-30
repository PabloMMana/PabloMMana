VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Begin VB.Form FormClientes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Clientes"
   ClientHeight    =   9165
   ClientLeft      =   45
   ClientTop       =   675
   ClientWidth     =   8805
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
   ScaleHeight     =   9165
   ScaleWidth      =   8805
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
      Height          =   8655
      Left            =   120
      TabIndex        =   7
      Top             =   240
      Width           =   8535
      Begin VB.TextBox TextCodigo 
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         Height          =   315
         Left            =   120
         TabIndex        =   15
         Top             =   480
         Width           =   855
      End
      Begin MSMask.MaskEdBox MaskcpfCLI 
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   1680
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   450
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin VB.ComboBox CmbUF 
         Height          =   330
         Left            =   3600
         TabIndex        =   5
         Top             =   2880
         Width           =   2895
      End
      Begin VB.ComboBox Cmbcidade 
         Height          =   330
         Left            =   120
         TabIndex        =   4
         Top             =   2880
         Width           =   2775
      End
      Begin VB.TextBox TextEnd 
         Height          =   285
         Left            =   120
         TabIndex        =   3
         Top             =   2280
         Width           =   8055
      End
      Begin VB.TextBox TextNome 
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Top             =   1080
         Width           =   8055
      End
      Begin VB.ComboBox CmbCor 
         Height          =   330
         Left            =   1200
         TabIndex        =   0
         Top             =   480
         Width           =   2655
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGridCli 
         Height          =   4935
         Left            =   120
         TabIndex        =   6
         Top             =   3360
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   8705
         _Version        =   393216
         Cols            =   7
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label8 
         Caption         =   "Código"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label7 
         Caption         =   "UF do Cliente"
         Height          =   255
         Left            =   3600
         TabIndex        =   14
         Top             =   2640
         Width           =   1215
      End
      Begin VB.Label Label6 
         Caption         =   "Cidade do Cliente"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   2640
         Width           =   2415
      End
      Begin VB.Label Label5 
         Caption         =   "Endereço do Cliente"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   2040
         Width           =   1695
      End
      Begin VB.Label Label4 
         Caption         =   "Nome do Cliente"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   840
         Width           =   2415
      End
      Begin VB.Label Label3 
         Caption         =   "C.P.F"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Corretor"
         Height          =   255
         Left            =   1200
         TabIndex        =   9
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Label1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   240
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
Attribute VB_Name = "FormClientes"
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


Public Function StringConexao()
'--String de Conexao

cn.ConnectionString = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=CORCLIENTE"

cn.Open

End Function



Private Sub CmbCor_DblClick()
'pega o id do item selecionado
num = CmbCor.ListIndex + 1


Dim conexao As New ADODB.Command
Dim rsdados As New ADODB.Recordset
Dim ssql As String

ssql = "SELECT * FROM TABCORRETORES where idcorretor='" & num & "'"


conexao.CommandType = adCmdText
conexao.CommandText = ssql
Set conexao.ActiveConnection = cn

rsdados.Open conexao, , adOpenDynamic, adLockOptimistic



End Sub

Private Sub Form_Load()

  Dim hMenu As Long
  hMenu = GetSystemMenu(hWnd, False)
  DeleteMenu hMenu, 6, MF_BYPOSITION

Call StringConexao
Call PreComboCor
Call PreComboCidade
Call PreComboUF
Call Maskcpf
Call AJUSTAGRID
Call PREEGRID
Call UltimoReg
End Sub

Private Sub UltimoReg()

Dim conexao As New ADODB.Command
Dim rsdados As New ADODB.Recordset
Dim ssql As String


ssql = "SELECT max(idcliente) idcliente FROM TABclientes"


conexao.CommandType = adCmdText
conexao.CommandText = ssql
Set conexao.ActiveConnection = cn

rsdados.Open conexao, , adOpenDynamic, adLockOptimistic


If IsNull(rsdados!idcliente) Then
TextCodigo.Text = 1
Else
TextCodigo.Text = rsdados!idcliente + 1
End If

End Sub



Public Sub AJUSTAGRID()



'---------Nomes das Colunas
MSFlexGridCli.TextMatrix(0, 1) = "Codigo"
MSFlexGridCli.TextMatrix(0, 2) = "Nome do Cliente"
MSFlexGridCli.TextMatrix(0, 3) = "CPF"
MSFlexGridCli.TextMatrix(0, 4) = "Endereço"
MSFlexGridCli.TextMatrix(0, 5) = "Cidade"
MSFlexGridCli.TextMatrix(0, 6) = "UF"


'---------Ajustar tamanho das colunas

MSFlexGridCli.ColWidth(0) = 0
MSFlexGridCli.ColWidth(1) = 600
MSFlexGridCli.ColWidth(2) = 2200
MSFlexGridCli.ColWidth(3) = 1200
MSFlexGridCli.ColWidth(4) = 3200
MSFlexGridCli.ColWidth(5) = 1000
MSFlexGridCli.ColWidth(6) = 600

End Sub


Public Sub PREEGRID()

Dim conexao As New ADODB.Command
Dim rsdados As New ADODB.Recordset
Dim ssql As String

ssql = "SELECT C.IDCLIENTE,C.NOME Nome,C.CPF,C.Endereco,CI.NOME Cidade,U.NOME UF FROM TABCLIENTES C INNER JOIN TABUF U ON C.IDUF = U.IDUF INNER JOIN TABCIDADE CI ON C.IDCIDADE=CI.IDCIDADE"


conexao.CommandType = adCmdText
conexao.CommandText = ssql
Set conexao.ActiveConnection = cn

rsdados.Open conexao, , adOpenDynamic, adLockOptimistic

MSFlexGridCli.Rows = Contalinhas + 1

I = 1
Do While Not rsdados.EOF

MSFlexGridCli.Row = I
MSFlexGridCli.Col = 1
MSFlexGridCli.Text = rsdados!idcliente

MSFlexGridCli.Row = I
MSFlexGridCli.Col = 2
MSFlexGridCli.Text = rsdados!nome

MSFlexGridCli.Row = I
MSFlexGridCli.Col = 3
MSFlexGridCli.Text = rsdados!cpf

MSFlexGridCli.Row = I
MSFlexGridCli.Col = 4
MSFlexGridCli.Text = rsdados!endereco

MSFlexGridCli.Row = I
MSFlexGridCli.Col = 5
MSFlexGridCli.Text = rsdados!cidade

MSFlexGridCli.Row = I
MSFlexGridCli.Col = 6
MSFlexGridCli.Text = rsdados!uf


I = I + 1
rsdados.MoveNext
Loop


End Sub

Private Sub MnuCad_Click()

If CmbCor.Text = "" Then
MsgBox "O corretor deve ser selecionado!", vbCritical, "Aviso"
ElseIf TextNome.Text = "" Then
MsgBox "O nome do cliente deve ser preenchido!", vbCritical, "Aviso"
ElseIf MaskcpfCLI.Text = "___.___.___-__" Then
MsgBox "O CPF do cliente deve ser preenchido!", vbCritical, "Aviso"
ElseIf TextEnd.Text = "" Then
MsgBox "O endereço do cliente deve ser preenchido!", vbCritical, "Aviso"
ElseIf Cmbcidade.Text = "" Then
MsgBox "A cidade do cliente deve ser selecionado!", vbCritical, "Aviso"
ElseIf CmbUF.Text = "" Then
MsgBox "A estado do cliente deve ser selecinado!", vbCritical, "Aviso"
ElseIf VerCli(MaskcpfCLI.Text) = True Then
MsgBox "O cliente já está cadastrado!", vbCritical, "Aviso"
Else
Call InserTabCli(CmbCor.Text, TextNome.Text, MaskcpfCLI.Text, TextEnd.Text, Cmbcidade.ListIndex, CmbUF.ListIndex)
End If
PREEGRID
End Sub

Private Function VerCli(cpf As String) As Boolean
Dim conexao As New ADODB.Command
Dim rsdados As New ADODB.Recordset
Dim ssql As String

ssql = "SELECT * FROM TABclientes WHERE CPF='" & cpf & "'"


conexao.CommandType = adCmdText
conexao.CommandText = ssql
Set conexao.ActiveConnection = cn
rsdados.Open conexao, , adOpenDynamic, adLockOptimistic


If rsdados.BOF And rsdados.EOF Then

VerCli = False

Else

VerCli = True

End If


End Function


Private Sub MnuSair_Click()
cn.Close
Unload Me
End Sub



Private Sub PreComboCor()

Dim conexao As New ADODB.Command
Dim rsdados As New ADODB.Recordset
Dim ssql As String

ssql = "SELECT * FROM TABCORRETORES"


conexao.CommandType = adCmdText
conexao.CommandText = ssql
Set conexao.ActiveConnection = cn
rsdados.Open conexao, , adOpenDynamic, adLockOptimistic


I = 1

Do While Not rsdados.EOF

CmbCor.AddItem (rsdados!nome)

rsdados.MoveNext

Loop


End Sub

Private Sub PreComboCidade()

Dim conexao As New ADODB.Command
Dim rsdados As New ADODB.Recordset
Dim ssql As String

ssql = "SELECT * FROM TABcidade"


conexao.CommandType = adCmdText
conexao.CommandText = ssql
Set conexao.ActiveConnection = cn

rsdados.Open conexao, , adOpenDynamic, adLockOptimistic
I = 1

Do While Not rsdados.EOF

Cmbcidade.AddItem (rsdados!nome)

rsdados.MoveNext

Loop
End Sub


Private Sub PreComboUF()

Dim conexao As New ADODB.Command
Dim rsdados As New ADODB.Recordset
Dim ssql As String

ssql = "SELECT * FROM TABUF"


conexao.CommandType = adCmdText
conexao.CommandText = ssql
Set conexao.ActiveConnection = cn

rsdados.Open conexao, , adOpenDynamic, adLockOptimistic
I = 1

Do While Not rsdados.EOF

CmbUF.AddItem (rsdados!nome)

rsdados.MoveNext

Loop
End Sub

Private Sub Maskcpf()
MaskcpfCLI.Mask = "###.###.###-##"
End Sub

Private Function InserClicor(idcor As String, idcli As String)

Dim conexao As New ADODB.Command
Dim rsdados As New ADODB.Recordset
Dim ssql As String




ssql = "INSERT INTO TABCLICOR (IDCORRETOR,IDCLIENTE) VALUES ('" & idcor & "','" & idcli & "' )"


conexao.CommandType = adCmdText
conexao.CommandText = ssql
Set conexao.ActiveConnection = cn

rsdados.Open conexao, , adOpenDynamic, adLockOptimistic


FormPrincipal.Pregrid



End Function


Private Function InserTabCli(corr As String, nome As String, cpf As String, ende As String, cid As String, est As String)

Dim conexao As New ADODB.Command
Dim rsdados As New ADODB.Recordset
Dim ssql As String




ssql = "INSERT INTO TABCLIENTES (NOME,CPF,ENDERECO,IDCIDADE,IDUF,ATIVO) VALUES ('" & nome & "','" & cpf & "','" & ende & "','" & cid & "' +1 ,'" & est & "' +1,0)"


conexao.CommandType = adCmdText
conexao.CommandText = ssql
Set conexao.ActiveConnection = cn

rsdados.Open conexao, , adOpenDynamic, adLockOptimistic

corr = CmbCor.ListIndex
corr = corr + 1



Call InserClicor(corr, TextCodigo.Text)


End Function


Public Function Contalinhas() As String

Dim conexao As New ADODB.Command
Dim rsdados As New ADODB.Recordset
Dim ssql As String

ssql = "SELECT count(idCLIENTE) total FROM TABCLIENTES"


conexao.CommandType = adCmdText
conexao.CommandText = ssql
Set conexao.ActiveConnection = cn

rsdados.Open conexao, , adOpenDynamic, adLockOptimistic

Contalinhas = rsdados!total

End Function

Private Sub MSFlexGridCli_DblClick()
num = MSFlexGridCli.TextMatrix(MSFlexGridCli.RowSel, 1)
Dim conexao As New ADODB.Command
Dim rsdados As New ADODB.Recordset
Dim ssql As String

If MsgBox("Deseja realmente excluir esse cliente?", vbYesNo, "Aviso") = vbYes Then

ssql = "DELETE FROM TABclientes WHERE IDCLIENTE='" & num & "'"


conexao.CommandType = adCmdText
conexao.CommandText = ssql
Set conexao.ActiveConnection = cn

rsdados.Open conexao, , adOpenDynamic, adLockOptimistic
Call PREEGRID
Call UltimoReg

Else
Call PREEGRID
cn.Close
End If

'cn.Open
End Sub
