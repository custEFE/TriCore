VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form Form2 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "正在播放"
   ClientHeight    =   4290
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   10950
   ControlBox      =   0   'False
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4290
   ScaleWidth      =   10950
   StartUpPosition =   1  '所有者中心
   Begin VB.Timer Timer2 
      Left            =   10080
      Top             =   1800
   End
   Begin VB.TextBox Info 
      Height          =   3735
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   240
      Width           =   9255
   End
   Begin VB.Timer Timer1 
      Left            =   9960
      Top             =   1080
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "停止播放"
      Height          =   495
      Left            =   9600
      TabIndex        =   1
      Top             =   360
      Width           =   1215
   End
   Begin WMPLibCtl.WindowsMediaPlayer MP2 
      Height          =   4575
      Left            =   6240
      TabIndex        =   2
      Top             =   4440
      Width           =   4095
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
      _cx             =   7223
      _cy             =   8070
   End
   Begin WMPLibCtl.WindowsMediaPlayer MP1 
      Height          =   4455
      Left            =   1440
      TabIndex        =   0
      Top             =   4440
      Width           =   4095
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
      _cx             =   7223
      _cy             =   7858
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim LP, ST, S3No, S5No, S8LP As Integer
Dim BlankTime, StartTime, DurTime As Single
Dim S5Yes, S8No, isPlaying As Boolean
Private Sub Command1_Click()
Form1.WindowState = 0
Unload Form2
End Sub

Private Sub Form_Load()
'======初始化
Info.Text = "开始播放..."
Randomize Timer
S8LP = 3 + Int(Rnd(1) * 4)
S8No = True
IntPlay
DoPlay
End Sub

Private Sub Form_Unload(Cancel As Integer)
Form1.Enabled = True
End Sub

Private Sub Info_Change()
Info.SelStart = Len(Info.Text)
End Sub

Private Sub Info_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub


Private Sub MP1_PlayStateChange(ByVal NewState As Long)
If NewState = 1 Then
  If isPlaying = True Then
    ST = ST + 1
    isPlaying = False
    DoPlay
  Else
    MP1.Controls.play
  End If
End If
If NewState = 3 Then isPlaying = True
Debug.Print NewState & "----" & MP1.URL
End Sub
Private Sub IntPlay()
LP = 1
ST = 1
Randomize Timer
End Sub

Private Sub DoPlay()
re:
Select Case ST
'===============================
Case 1
Info.Text = Info.Text & vbCrLf & "执行第" & LP & "循环 第1时段-宣布开始：播放" & GetFileName(Form1.InsM(2).Caption)
If Form1.InsM(2).Caption <> "插入声音" Then
  MP1.URL = Form1.InsM(2).Caption
Else
  ST = ST + 1
  GoTo re
End If

'===============================
Case 2
Info.Text = Info.Text & vbCrLf & "执行第" & LP & "循环 第2时段-动作间隔时间"

If Form1.InsM(0).Caption <> "插入声音" And LP = 1 Then
  MP2.settings.playCount = 99999
  MP2.URL = Form1.InsM(0).Caption
End If

Randomize Timer
ta = Val(Form1.NumInp(1))
tb = Val(Form1.NumInp(2))
dur = ta + Int(Rnd(1) * Abs(ta - tb) * 100) / 100
Silent (dur)
Info.Text = Info.Text & dur & "秒"

'===============================
Case 3
Info.Text = Info.Text & vbCrLf & "执行第" & LP & "循环 第3时段-动作指示信号-播放"
i = 3
Do Until Form1.InsM(i).Caption = "插入声音"
  i = i + 1
Loop

Randomize Timer

plen = Val(Form1.NumInp(0).Text)
If plen = 0 Then
  MP1.settings.playCount = 1
  Info.Text = Info.Text & "1遍"
ElseIf plen > 0 Then
  MP1.settings.playCount = 9999
  PlayLen (Val(Form1.NumInp(0).Text))
  Info.Text = Info.Text & plen & "秒"
End If
Select Case i
  Case 4
    MP1.URL = Form1.InsM(3).Caption
    S3No = 0
  Case 5
    If Rnd(1) >= 0.5 Then
      MP1.URL = Form1.InsM(4).Caption
      S3No = 1
    Else
      MP1.URL = Form1.InsM(3).Caption
      S3No = 0
    End If
  Case 6 To 11
    If Form1.Option1(0).Value = True Then
       r = Rnd(1) * 10
       If r < 4 Then MP1.URL = Form1.InsM(3).Caption: S3No = 0
       If r >= 4 And r < 8 Then MP1.URL = Form1.InsM(4).Caption: S3No = 1
       If r > 8 Then
         n = Int(Rnd(1) * (i - 5)) + 5
         MP1.URL = Form1.InsM(n).Caption
         S3No = n - 3
       End If
    Else
      n = Int(Rnd(1) * (i - 3)) + 3
      MP1.URL = Form1.InsM(n).Caption
      S3No = n - 3
    End If
End Select

Info.Text = Info.Text & "第" & (S3No + 1) & "号声音：" & GetFileName(Form1.InsM(S3No + 3).Caption)

'===============================
Case 4
Info.Text = Info.Text & vbCrLf & "执行第" & LP & "循环 第4时段-听到信号后开始做动作的时间-"
Silent (Val(Form1.Combo1.Text))
Info.Text = Info.Text & Form1.Combo1.Text & "秒"
'===============================
Case 5
Info.Text = Info.Text & vbCrLf & "执行第" & LP & "循环 第5时段-持续做动作的时间-"

MP1.settings.playCount = 1

  If Form1.Option2(0).Value = True Or (Form1.Option2(1).Value = True And S3No = S5No) Or (Form1.Option2(2).Value = True And S3No <> S5No) Then
    MP1.URL = Form1.InsM(S3No + 11).Caption
    Info.Text = Info.Text & "播放" & GetFileName(Form1.InsM(S3No + 11).Caption)
    S5Yes = True
    S5No = S3No
  Else
    Info.Text = Info.Text & "跳过播放"
    ST = ST + 1
    S5Yes = False
    S5No = S3No
    GoTo re
  End If

'===============================
Case 6
Info.Text = Info.Text & vbCrLf & "执行第" & LP & "循环 第6时段-不触发动作的通知-"
If S5Yes = True Then
  Info.Text = Info.Text & "跳过播放"
  ST = ST + 1
  GoTo re
Else
  Info.Text = Info.Text & "播放" & GetFileName(Form1.InsM(19).Caption)
  MP1.URL = Form1.InsM(19).Caption
End If
'===============================
Case 7
Info.Text = Info.Text & vbCrLf & "执行第" & LP & "循环 第7时段-动作结束-"
If Form1.InsM(20).Caption = "插入声音" Then
  Info.Text = Info.Text & "跳过播放"
  ST = ST + 1
  GoTo re
Else
  Info.Text = Info.Text & "播放" & GetFileName(Form1.InsM(20).Caption)
  MP1.URL = Form1.InsM(20).Caption
End If
'===============================
Case 8
Info.Text = Info.Text & vbCrLf & "执行第" & LP & "循环 第8时段-交换动作信号-"
If Form1.Check1.Value = 0 Then
  Info.Text = Info.Text & "跳过"
  ST = ST + 1
  GoTo re
Else
  If LP = S8LP Then
    If S8No = True Then
      MP1.URL = Form1.InsM(21).Caption
      Info.Text = Info.Text & "播放" & GetFileName(Form1.InsM(21).Caption)
      If Form1.InsM(1).Caption <> "插入声音" Then
        MP2.URL = Form1.InsM(1).Caption
        Info.Text = Info.Text & "背景播放" & GetFileName(Form1.InsM(1).Caption)
      End If
      S8No = False
    Else
      MP1.URL = Form1.InsM(22).Caption
      Info.Text = Info.Text & "播放" & GetFileName(Form1.InsM(22).Caption)
      If Form1.InsM(0).Caption <> "插入声音" Then
        MP2.URL = Form1.InsM(0).Caption
        Info.Text = Info.Text & "背景播放" & GetFileName(Form1.InsM(0).Caption)
      End If
      S8No = True
    End If

    S8LP = S8LP + 3 + Int(Rnd(1) * 4)
  Else
    Info.Text = Info.Text & "跳过，将在第" & S8LP & "次循环时播放"
    ST = ST + 1
    GoTo re
  End If
End If

'===============================
Case 9
  mlp = Val(Form1.NumInp(3).Text)
If LP < mlp Then
    ST = 2
    LP = LP + 1
    GoTo re
Else
    MP2.Controls.stop
    If Form1.InsM(23).Caption <> "插入声音" Then
      MP1.URL = Form1.InsM(23).Caption
      Info.Text = Info.Text & vbCrLf & "执行第9时段-游戏结束信号：播放" & GetFileName(Form1.InsM(23).Caption)
    End If

End If
'===============================
Case 10
    Info.Text = Info.Text & vbCrLf & "执行完毕！"
    Command1.Caption = "关闭"
End Select
End Sub

Private Sub Timer1_Timer()
DurTime = Timer - StartTime
If DurTime >= BlankTime Then
  Timer1.Interval = 0
  ST = ST + 1
  DoPlay
End If
End Sub
Private Sub Silent(duration As Single)
StartTime = Timer
BlankTime = duration
Timer1.Interval = 10
End Sub
Private Sub PlayLen(duration As Single)
StartTime = Timer
BlankTime = duration
Timer2.Interval = 10
End Sub
Function GetFileName(FullPath As String)
GetFileName = "【" & Mid$(FullPath, InStrRev(FullPath, "\") + 1) & "】"
End Function

Private Sub Timer2_Timer()
DurTime = Timer - StartTime
If DurTime >= BlankTime Then
  Timer2.Interval = 0
  MP1.Controls.stop
End If
End Sub
