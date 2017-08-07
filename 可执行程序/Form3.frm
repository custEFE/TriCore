VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form Form3 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "插入声音"
   ClientHeight    =   2040
   ClientLeft      =   4905
   ClientTop       =   3720
   ClientWidth     =   5250
   ControlBox      =   0   'False
   Icon            =   "Form3.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2040
   ScaleWidth      =   5250
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command4 
      Caption         =   "确定"
      Default         =   -1  'True
      Height          =   615
      Left            =   1440
      TabIndex        =   4
      Top             =   1200
      Width           =   2535
   End
   Begin VB.CommandButton Command3 
      Caption         =   "试听"
      Height          =   495
      Left            =   3480
      TabIndex        =   2
      Top             =   360
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "删除"
      Height          =   495
      Left            =   2040
      TabIndex        =   1
      Top             =   360
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "打开"
      Height          =   495
      Left            =   600
      TabIndex        =   0
      Top             =   360
      Width           =   1215
   End
   Begin WMPLibCtl.WindowsMediaPlayer MP 
      Height          =   375
      Left            =   1680
      TabIndex        =   3
      Top             =   2160
      Width           =   1455
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   2566
      _cy             =   661
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
sOpenFileName = GetDialog("open", "打开文件", "", "音频文件(*.mp3|*.wav|*.aac)" & Chr$(0) & "*.mp3;*.wav;*.aac" & Chr$(0))
If Not sOpenFileName = "" Then Form1.InsM(SelInd).Caption = sOpenFileName
Command3.Enabled = True
End Sub

Private Sub Command2_Click()
Form1.InsM(SelInd).Caption = "插入声音"
Command3.Enabled = False
End Sub

Private Sub Command3_Click()
MP.URL = Form1.InsM(SelInd).Caption
MP.Controls.play
End Sub

Private Sub Command4_Click()
MP.Controls.stop
Me.Hide
Form1.SetFocus
End Sub

Private Sub Form_Activate()
If Form1.InsM(SelInd).Caption = "插入声音" Then Command3.Enabled = False Else Command3.Enabled = True
End Sub

Private Sub Form_Load()
SetParent Form3.hWnd, Form1.hWnd
End Sub
