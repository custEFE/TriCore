VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TriCore Games"
   ClientHeight    =   9270
   ClientLeft      =   3735
   ClientTop       =   3210
   ClientWidth     =   14835
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9270
   ScaleWidth      =   14835
   StartUpPosition =   2  '��Ļ����
   Begin VB.TextBox NumInp 
      Height          =   375
      Index           =   3
      Left            =   6840
      MaxLength       =   4
      TabIndex        =   49
      Text            =   "6"
      Top             =   7680
      Width           =   495
   End
   Begin VB.Frame Frame10 
      Caption         =   "9-��Ϸ�����ź�"
      Height          =   1335
      Left            =   240
      TabIndex        =   47
      Top             =   7320
      Width           =   4695
      Begin VB.CommandButton InsM 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000E&
         Caption         =   "��������"
         Height          =   495
         Index           =   23
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   48
         Top             =   480
         Width           =   4215
      End
   End
   Begin VB.Frame Frame9 
      Caption         =   "8-���������ź�"
      Height          =   2295
      Left            =   9840
      TabIndex        =   12
      Top             =   5760
      Width           =   4695
      Begin VB.CheckBox Check1 
         Caption         =   "����"
         Height          =   375
         Left            =   360
         TabIndex        =   46
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton InsM 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         Caption         =   "��������"
         Height          =   495
         Index           =   22
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   45
         Top             =   1440
         Width           =   4215
      End
      Begin VB.CommandButton InsM 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         Caption         =   "��������"
         Height          =   495
         Index           =   21
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   44
         Top             =   840
         Width           =   4215
      End
   End
   Begin VB.Frame Frame8 
      Caption         =   "7-��������"
      Height          =   1455
      Left            =   5040
      TabIndex        =   11
      Top             =   5760
      Width           =   4695
      Begin VB.CommandButton InsM 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000E&
         Caption         =   "��������"
         Height          =   495
         Index           =   20
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   480
         Width           =   4215
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "6-������������֪ͨ"
      Height          =   1455
      Left            =   240
      TabIndex        =   10
      Top             =   5760
      Width           =   4695
      Begin VB.CommandButton InsM 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         Caption         =   "��������"
         Height          =   495
         Index           =   19
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   480
         Width           =   4215
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "5-������������ʱ��"
      Height          =   3255
      Left            =   9840
      TabIndex        =   9
      Top             =   2400
      Width           =   4695
      Begin VB.OptionButton Option2 
         Caption         =   "��"
         Height          =   375
         Index           =   2
         Left            =   2520
         TabIndex        =   41
         Top             =   2400
         Width           =   1095
      End
      Begin VB.OptionButton Option2 
         Caption         =   "��"
         Height          =   375
         Index           =   1
         Left            =   1680
         TabIndex        =   40
         Top             =   2400
         Width           =   975
      End
      Begin VB.OptionButton Option2 
         Caption         =   "������"
         Height          =   375
         Index           =   0
         Left            =   480
         TabIndex        =   39
         Top             =   2400
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.CommandButton InsM 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000E&
         Caption         =   "��������"
         Height          =   495
         Index           =   18
         Left            =   2400
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   1680
         Width           =   1935
      End
      Begin VB.CommandButton InsM 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000E&
         Caption         =   "��������"
         Height          =   495
         Index           =   17
         Left            =   2400
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   1200
         Width           =   1935
      End
      Begin VB.CommandButton InsM 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000E&
         Caption         =   "��������"
         Height          =   495
         Index           =   16
         Left            =   2400
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   720
         Width           =   1935
      End
      Begin VB.CommandButton InsM 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000E&
         Caption         =   "��������"
         Height          =   495
         Index           =   15
         Left            =   2400
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   240
         Width           =   1935
      End
      Begin VB.CommandButton InsM 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000E&
         Caption         =   "��������"
         Height          =   495
         Index           =   14
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   1680
         Width           =   1935
      End
      Begin VB.CommandButton InsM 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000E&
         Caption         =   "��������"
         Height          =   495
         Index           =   13
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   1200
         Width           =   1935
      End
      Begin VB.CommandButton InsM 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         Caption         =   "��������"
         Height          =   495
         Index           =   12
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   720
         Width           =   1935
      End
      Begin VB.CommandButton InsM 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         Caption         =   "��������"
         Height          =   495
         Index           =   11
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "��ʼ����"
      Default         =   -1  'True
      Height          =   495
      Left            =   9840
      TabIndex        =   5
      Top             =   8280
      Width           =   4695
   End
   Begin VB.Frame Frame5 
      Caption         =   "4-�����źź�ʼ��������ʱ��"
      Height          =   3255
      Left            =   5040
      TabIndex        =   4
      Top             =   2400
      Width           =   4695
      Begin VB.ComboBox Combo1 
         Height          =   300
         ItemData        =   "Form1.frx":3332
         Left            =   1680
         List            =   "Form1.frx":3348
         TabIndex        =   30
         Text            =   "1.25"
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "����ʱ��"
         Height          =   375
         Left            =   600
         TabIndex        =   29
         Top             =   1200
         Width           =   855
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "3-����ָʾ�ź�"
      Height          =   3255
      Left            =   240
      TabIndex        =   3
      Top             =   2400
      Width           =   4695
      Begin VB.OptionButton Option1 
         Caption         =   "�����ʽ��"
         Height          =   180
         Index           =   1
         Left            =   2280
         TabIndex        =   28
         Top             =   2760
         Width           =   1695
      End
      Begin VB.OptionButton Option1 
         Caption         =   "�����ʽһ"
         Height          =   180
         Index           =   0
         Left            =   2280
         TabIndex        =   27
         Top             =   2400
         Value           =   -1  'True
         Width           =   1695
      End
      Begin VB.TextBox NumInp 
         Height          =   375
         Index           =   0
         Left            =   1320
         MaxLength       =   4
         TabIndex        =   26
         Text            =   "0"
         Top             =   2400
         Width           =   495
      End
      Begin VB.CommandButton InsM 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000E&
         Caption         =   "��������"
         Height          =   495
         Index           =   10
         Left            =   2280
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   1680
         Width           =   1935
      End
      Begin VB.CommandButton InsM 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000E&
         Caption         =   "��������"
         Height          =   495
         Index           =   9
         Left            =   2280
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   1200
         Width           =   1935
      End
      Begin VB.CommandButton InsM 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000E&
         Caption         =   "��������"
         Height          =   495
         Index           =   8
         Left            =   2280
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   720
         Width           =   1935
      End
      Begin VB.CommandButton InsM 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000E&
         Caption         =   "��������"
         Height          =   495
         Index           =   7
         Left            =   2280
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   240
         Width           =   1935
      End
      Begin VB.CommandButton InsM 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000E&
         Caption         =   "��������"
         Height          =   495
         Index           =   6
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   1680
         Width           =   1935
      End
      Begin VB.CommandButton InsM 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000E&
         Caption         =   "��������"
         Height          =   495
         Index           =   5
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   1200
         Width           =   1935
      End
      Begin VB.CommandButton InsM 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         Caption         =   "��������"
         Height          =   495
         Index           =   4
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   720
         Width           =   1935
      End
      Begin VB.CommandButton InsM 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         Caption         =   "��������"
         Height          =   495
         Index           =   3
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Label2 
         Caption         =   "����ʱ��  ������0�򲥷�һ��"
         Height          =   615
         Left            =   360
         TabIndex        =   25
         Top             =   2400
         Width           =   855
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "2-��������ʱ��"
      Height          =   2055
      Left            =   9840
      TabIndex        =   2
      Top             =   240
      Width           =   4695
      Begin VB.TextBox NumInp 
         Height          =   375
         Index           =   2
         Left            =   3000
         MaxLength       =   4
         TabIndex        =   15
         Text            =   "3"
         Top             =   840
         Width           =   495
      End
      Begin VB.TextBox NumInp 
         Height          =   375
         Index           =   1
         Left            =   1800
         MaxLength       =   4
         TabIndex        =   14
         Text            =   "1.5"
         Top             =   840
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "�������         ����"
         Height          =   375
         Left            =   960
         TabIndex        =   16
         Top             =   960
         Width           =   2655
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "1-������ʼ"
      Height          =   2055
      Left            =   5040
      TabIndex        =   1
      Top             =   240
      Width           =   4695
      Begin VB.CommandButton InsM 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000E&
         Caption         =   "��������"
         Height          =   495
         Index           =   2
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   360
         Width           =   4215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "0-������"
      Height          =   2055
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   4695
      Begin VB.CommandButton InsM 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000E&
         Caption         =   "��������"
         Height          =   495
         Index           =   1
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   960
         Width           =   4215
      End
      Begin VB.CommandButton InsM 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000E&
         Caption         =   "��������"
         Height          =   495
         Index           =   0
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   360
         Width           =   4215
      End
   End
   Begin VB.Label Status 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "���ļ�"
      Height          =   375
      Left            =   0
      TabIndex        =   50
      Top             =   9000
      Width           =   14895
   End
   Begin VB.Label Label1 
      Caption         =   "ѭ��������"
      Height          =   375
      Left            =   5640
      TabIndex        =   6
      Top             =   7800
      Width           =   1095
   End
   Begin VB.Menu File 
      Caption         =   "�ļ�"
      Begin VB.Menu NewFile 
         Caption         =   "�½����"
         Shortcut        =   ^N
      End
      Begin VB.Menu OpenFile 
         Caption         =   "�����..."
         Shortcut        =   ^O
      End
      Begin VB.Menu SaveFile 
         Caption         =   "�������..."
         Shortcut        =   ^S
      End
      Begin VB.Menu line1 
         Caption         =   "-"
      End
      Begin VB.Menu ExitFile 
         Caption         =   "�˳�"
      End
   End
   Begin VB.Menu Helping 
      Caption         =   "����"
      Begin VB.Menu About 
         Caption         =   "����"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub About_Click()
MsgBox "��Ȩ���� 2017 ����Ⱥ", vbInformation
End Sub

Private Sub Command1_Click()
'=====��Ч�Լ��
If InsM(3).Caption = "��������" Then MsgBox "��3ʱ�ε�1������������룡", vbCritical: Exit Sub
For i = 3 To 10
  If InsM(i).Caption <> "��������" And InsM(i + 8).Caption = "��������" Then MsgBox "��5ʱ�α������3ʱ�ζ�Ӧ��", vbCritical: Exit Sub
Next i
If InsM(19).Caption = "��������" Then MsgBox "��6ʱ�ε�1������������룡", vbCritical: Exit Sub
If Check1.Value = 1 And (InsM(21).Caption = "��������" Or InsM(22).Caption = "��������") Then MsgBox "�����8ʱ�ι�ѡ������2����������ȫ�����룡", vbCritical: Exit Sub
'=====
Form1.Enabled = False
Form1.WindowState = 1
Form2.Show
End Sub


Private Sub ExitFile_Click()
Unload Me
End Sub

Private Sub Form_Load()
'��ʼ��
CurrentFile = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
q = MsgBox("�Ƿ񱣴浱ǰ��ϣ�", vbYesNoCancel + vbExclamation)
If q = vbYes Then
  SaveFile_Click
End If
If q = vbCancel Then Cancel = -1
If Not Cancel = -1 Then End
End Sub

Private Sub InsM_Click(Index As Integer)
Form3.Show
SelInd = Index

End Sub

Private Sub NewFile_Click()
If MsgBox("�Ƿ������ǰ��ϲ��½���", vbYesNo + vbExclamation) = vbNo Then Exit Sub
For i = 0 To 23
  InsM(i).Caption = "��������"
Next i

NumInp(0).Text = "0"
NumInp(1).Text = "1.5"
NumInp(2).Text = "3"
NumInp(3).Text = "6"
Option1(0).Value = True
Option2(0).Value = True
Combo1.Text = "1.25"
Check1.Value = False

CurrentFile = ""
Status.Caption = "���ļ�"
End Sub


Private Sub NumInp_LostFocus(Index As Integer)
If IsNumeric(NumInp(Index).Text) = False Then
  MsgBox "����Ĳ������֣����������롣", vbExclamation
  NumInp(Index).Text = "0"
End If
End Sub

Private Sub OpenFile_Click()
sOpenFileName = GetDialog("open", "���ļ�", "*.tri", "TriCore����ļ�(*.tri)" & Chr$(0) & "*.tri" & Chr$(0))
If sOpenFileName = "" Then Exit Sub
Open sOpenFileName For Input As #1 '
Line Input #1, firstline
If firstline <> "TriCore File" Then
  MsgBox "�Ƿ�tri�ļ���", vbCritical
  Exit Sub
End If

Dim arr(35) As String
i = 0
Do While Not EOF(1)
Line Input #1, arr(i)
i = i + 1
Loop
Close #1

If i < 35 Then
  MsgBox "�Ƿ�tri�ļ���", vbCritical
  Exit Sub
End If

CurrentFile = sOpenFileName

For i = 0 To 23
  InsM(i).Caption = arr(i)
Next i

NumInp(0).Text = arr(24)
NumInp(1).Text = arr(25)
NumInp(2).Text = arr(26)
NumInp(3).Text = arr(27)
Option1(0).Value = StrToBoo(arr(28))
Option1(1).Value = StrToBoo(arr(29))
Option2(0).Value = StrToBoo(arr(30))
Option2(1).Value = StrToBoo(arr(31))
Option2(2).Value = StrToBoo(arr(32))
Combo1.Text = arr(33)
Check1.Value = Val(arr(34))

Status.Caption = "��ǰ�ļ�" & CurrentFile
End Sub

Private Sub SaveFile_Click()
If CurrentFile = "" Then
  sSaveFileName = GetDialog("save", "�����ļ�", "*.tri", "TriCore����ļ�(*.tri)" & Chr$(0) & "*.tri" & Chr$(0))
Else
  sSaveFileName = CurrentFile
End If

If sSaveFileName = "" Then Exit Sub

Open sSaveFileName For Output As #1
Print #1, "TriCore File"
For i = 0 To 23
  Print #1, InsM(i).Caption
Next i
Print #1, NumInp(0).Text
Print #1, NumInp(1).Text
Print #1, NumInp(2).Text
Print #1, NumInp(3).Text
Print #1, Option1(0).Value
Print #1, Option1(1).Value
Print #1, Option2(0).Value
Print #1, Option2(1).Value
Print #1, Option2(2).Value
Print #1, Combo1.Text
Print #1, Check1.Value

Close #1
CurrentFile = sSaveFileName
Status.Caption = "����ɹ���" & sSaveFileName
End Sub
Function StrToBoo(str As String) As Boolean
If str = "True" Or Trim(str) = "1" Then StrToBoo = True Else StrToBoo = False
End Function
