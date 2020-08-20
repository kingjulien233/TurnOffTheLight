VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00E0E0E0&
   Caption         =   "关灯"
   ClientHeight    =   4290
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4905
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4290
   ScaleWidth      =   4905
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton CommandO 
      Caption         =   "确定!"
      Height          =   495
      Left            =   1800
      TabIndex        =   38
      Top             =   2760
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   240
      TabIndex        =   36
      Top             =   1920
      Visible         =   0   'False
      Width           =   4455
   End
   Begin VB.CommandButton CommandR 
      Caption         =   "重新开始!"
      Height          =   495
      Left            =   3480
      TabIndex        =   32
      Top             =   2640
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton CommandM 
      Caption         =   "返回菜单!"
      Height          =   495
      Left            =   3480
      TabIndex        =   31
      Top             =   3360
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton CommandE 
      Caption         =   "退出游戏!"
      Height          =   495
      Left            =   1800
      TabIndex        =   30
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton CommandP 
      Caption         =   "输入密码!"
      Height          =   495
      Left            =   1800
      TabIndex        =   29
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton CommandS 
      Caption         =   "开始游戏!"
      Height          =   495
      Left            =   1800
      TabIndex        =   28
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton CommandN 
      Caption         =   "下一关!"
      Height          =   495
      Left            =   1200
      TabIndex        =   25
      Top             =   3240
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label26 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "TurnOffTheLight  NaiveGames  Copyright2014"
      Height          =   255
      Left            =   120
      TabIndex        =   40
      Top             =   3960
      Width           =   4575
   End
   Begin VB.Label LabelEND 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "恭喜通关！"
      BeginProperty Font 
         Name            =   "隶书"
         Size            =   36
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   360
      TabIndex        =   39
      Top             =   1080
      Visible         =   0   'False
      Width           =   4215
   End
   Begin VB.Label LabelW 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "请输入密码"
      BeginProperty Font 
         Name            =   "华文楷体"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1440
      TabIndex        =   37
      Top             =   1200
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label LabelT 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "关灯"
      BeginProperty Font 
         Name            =   "隶书"
         Size            =   42
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1440
      TabIndex        =   35
      Top             =   240
      Width           =   1935
   End
   Begin VB.Label LabelP2 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "华文楷体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3480
      TabIndex        =   34
      Top             =   2040
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label LabelP1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "密码"
      BeginProperty Font 
         Name            =   "华文楷体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      TabIndex        =   33
      Top             =   1440
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label LabelN 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "华文楷体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3720
      TabIndex        =   27
      Top             =   840
      Width           =   735
   End
   Begin VB.Label LabelL 
      BackColor       =   &H00E0E0E0&
      Caption         =   "关卡"
      BeginProperty Font 
         Name            =   "华文楷体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      TabIndex        =   26
      Top             =   240
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label25 
      BackColor       =   &H00000000&
      Height          =   495
      Left            =   2760
      TabIndex        =   24
      Top             =   2520
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label24 
      BackColor       =   &H00000000&
      Height          =   495
      Left            =   2160
      TabIndex        =   23
      Top             =   2520
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label23 
      BackColor       =   &H00000000&
      Height          =   495
      Left            =   1560
      TabIndex        =   22
      Top             =   2520
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label22 
      BackColor       =   &H00000000&
      Height          =   495
      Left            =   960
      TabIndex        =   21
      Top             =   2520
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label21 
      BackColor       =   &H00000000&
      Height          =   495
      Left            =   360
      TabIndex        =   20
      Top             =   2520
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label20 
      BackColor       =   &H00000000&
      Height          =   495
      Left            =   2760
      TabIndex        =   19
      Top             =   1920
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label19 
      BackColor       =   &H00000000&
      Height          =   495
      Left            =   2160
      TabIndex        =   18
      Top             =   1920
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label18 
      BackColor       =   &H00000000&
      Height          =   495
      Left            =   1560
      TabIndex        =   17
      Top             =   1920
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label17 
      BackColor       =   &H00000000&
      Height          =   495
      Left            =   960
      TabIndex        =   16
      Top             =   1920
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label16 
      BackColor       =   &H00000000&
      Height          =   495
      Left            =   360
      TabIndex        =   15
      Top             =   1920
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label15 
      BackColor       =   &H00000000&
      Height          =   495
      Left            =   2760
      TabIndex        =   14
      Top             =   1320
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label14 
      BackColor       =   &H00000000&
      Height          =   495
      Left            =   2160
      TabIndex        =   13
      Top             =   1320
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label13 
      BackColor       =   &H00000000&
      Height          =   495
      Left            =   1560
      TabIndex        =   12
      Top             =   1320
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label12 
      BackColor       =   &H00000000&
      Height          =   495
      Left            =   960
      TabIndex        =   11
      Top             =   1320
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label11 
      BackColor       =   &H00000000&
      Height          =   495
      Left            =   360
      TabIndex        =   10
      Top             =   1320
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label10 
      BackColor       =   &H00000000&
      Height          =   495
      Left            =   2760
      TabIndex        =   9
      Top             =   720
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label9 
      BackColor       =   &H00000000&
      Height          =   495
      Left            =   2160
      TabIndex        =   8
      Top             =   720
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label8 
      BackColor       =   &H00000000&
      Height          =   495
      Left            =   1560
      TabIndex        =   7
      Top             =   720
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label7 
      BackColor       =   &H00000000&
      Height          =   495
      Left            =   960
      TabIndex        =   6
      Top             =   720
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label6 
      BackColor       =   &H00000000&
      Height          =   495
      Left            =   360
      TabIndex        =   5
      Top             =   720
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label5 
      BackColor       =   &H00000000&
      Height          =   495
      Left            =   2760
      TabIndex        =   4
      Top             =   120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label4 
      BackColor       =   &H00000000&
      Height          =   495
      Left            =   2160
      TabIndex        =   3
      Top             =   120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      Height          =   495
      Left            =   1560
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Height          =   495
      Left            =   960
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a As Integer

Private Sub CommandE_Click()
End
End Sub

Private Sub CommandM_Click()
Label1.Visible = False
Label2.Visible = False
Label3.Visible = False
Label4.Visible = False
Label5.Visible = False
Label6.Visible = False
Label7.Visible = False
Label8.Visible = False
Label9.Visible = False
Label10.Visible = False
Label11.Visible = False
Label12.Visible = False
Label13.Visible = False
Label14.Visible = False
Label15.Visible = False
Label16.Visible = False
Label17.Visible = False
Label18.Visible = False
Label19.Visible = False
Label20.Visible = False
Label21.Visible = False
Label22.Visible = False
Label23.Visible = False
Label24.Visible = False
Label25.Visible = False
LabelN.Visible = False
LabelL.Visible = False
LabelT.Visible = True
LabelP1.Visible = False
LabelP2.Visible = False
CommandM.Visible = False
CommandN.Visible = False
CommandR.Visible = False
CommandS.Visible = True
CommandP.Visible = True
CommandE.Visible = True
End Sub

Private Sub CommandN_Click()
a = a + 1
LabelN = a
If a = 2 Then
  Label1.BackColor = &HFFFFFF
  Label2.BackColor = &HFFFFFF
  Label3.BackColor = &H0&
  Label4.BackColor = &H0&
  Label5.BackColor = &H0&
  Label6.BackColor = &HFFFFFF
  Label7.BackColor = &H0&
  Label8.BackColor = &H0&
  Label9.BackColor = &H0&
  Label10.BackColor = &H0&
  Label11.BackColor = &H0&
  Label12.BackColor = &H0&
  Label13.BackColor = &H0&
  Label14.BackColor = &H0&
  Label15.BackColor = &H0&
  Label16.BackColor = &H0&
  Label17.BackColor = &H0&
  Label18.BackColor = &HFFFFFF
  Label19.BackColor = &H0&
  Label20.BackColor = &H0&
  Label21.BackColor = &H0&
  Label22.BackColor = &HFFFFFF
  Label23.BackColor = &HFFFFFF
  Label24.BackColor = &HFFFFFF
  Label25.BackColor = &H0&
End If
If a = 3 Then
  Label1.BackColor = &H0&
  Label2.BackColor = &H0&
  Label3.BackColor = &H0&
  Label4.BackColor = &H0&
  Label5.BackColor = &H0&
  Label6.BackColor = &H0&
  Label7.BackColor = &HFFFFFF
  Label8.BackColor = &HFFFFFF
  Label9.BackColor = &HFFFFFF
  Label10.BackColor = &H0&
  Label11.BackColor = &HFFFFFF
  Label12.BackColor = &H0&
  Label13.BackColor = &HFFFFFF
  Label14.BackColor = &H0&
  Label15.BackColor = &HFFFFFF
  Label16.BackColor = &H0&
  Label17.BackColor = &HFFFFFF
  Label18.BackColor = &HFFFFFF
  Label19.BackColor = &HFFFFFF
  Label20.BackColor = &H0&
  Label21.BackColor = &H0&
  Label22.BackColor = &H0&
  Label23.BackColor = &H0&
  Label24.BackColor = &H0&
  Label25.BackColor = &H0&
End If
If a = 4 Then
  Label1.BackColor = &H0&
  Label2.BackColor = &H0&
  Label3.BackColor = &H0&
  Label4.BackColor = &H0&
  Label5.BackColor = &H0&
  Label6.BackColor = &H0&
  Label7.BackColor = &H0&
  Label8.BackColor = &H0&
  Label9.BackColor = &HFFFFFF
  Label10.BackColor = &HFFFFFF
  Label11.BackColor = &H0&
  Label12.BackColor = &H0&
  Label13.BackColor = &H0&
  Label14.BackColor = &H0&
  Label15.BackColor = &H0&
  Label16.BackColor = &H0&
  Label17.BackColor = &HFFFFFF
  Label18.BackColor = &HFFFFFF
  Label19.BackColor = &H0&
  Label20.BackColor = &H0&
  Label21.BackColor = &H0&
  Label22.BackColor = &H0&
  Label23.BackColor = &HFFFFFF
  Label24.BackColor = &HFFFFFF
  Label25.BackColor = &HFFFFFF
End If
If a = 5 Then
  Label1.BackColor = &HFFFFFF
  Label2.BackColor = &H0&
  Label3.BackColor = &H0&
  Label4.BackColor = &H0&
  Label5.BackColor = &HFFFFFF
  Label6.BackColor = &H0&
  Label7.BackColor = &HFFFFFF
  Label8.BackColor = &H0&
  Label9.BackColor = &HFFFFFF
  Label10.BackColor = &H0&
  Label11.BackColor = &H0&
  Label12.BackColor = &HFFFFFF
  Label13.BackColor = &H0&
  Label14.BackColor = &HFFFFFF
  Label15.BackColor = &H0&
  Label16.BackColor = &H0&
  Label17.BackColor = &H0&
  Label18.BackColor = &H0&
  Label19.BackColor = &H0&
  Label20.BackColor = &H0&
  Label21.BackColor = &H0&
  Label22.BackColor = &H0&
  Label23.BackColor = &H0&
  Label24.BackColor = &H0&
  Label25.BackColor = &H0&
  LabelP1.Visible = True
  LabelP2.Visible = True
  LabelP2 = "soeasy"
End If
If a = 6 Then
  Label1.BackColor = &H0&
  Label2.BackColor = &H0&
  Label3.BackColor = &H0&
  Label4.BackColor = &H0&
  Label5.BackColor = &H0&
  Label6.BackColor = &HFFFFFF
  Label7.BackColor = &HFFFFFF
  Label8.BackColor = &H0&
  Label9.BackColor = &HFFFFFF
  Label10.BackColor = &HFFFFFF
  Label11.BackColor = &H0&
  Label12.BackColor = &H0&
  Label13.BackColor = &H0&
  Label14.BackColor = &H0&
  Label15.BackColor = &H0&
  Label16.BackColor = &H0&
  Label17.BackColor = &H0&
  Label18.BackColor = &H0&
  Label19.BackColor = &H0&
  Label20.BackColor = &H0&
  Label21.BackColor = &H0&
  Label22.BackColor = &H0&
  Label23.BackColor = &H0&
  Label24.BackColor = &H0&
  Label25.BackColor = &H0&
  LabelP1.Visible = False
  LabelP2.Visible = False
End If
If a = 7 Then
  Label1.BackColor = &H0&
  Label2.BackColor = &H0&
  Label3.BackColor = &H0&
  Label4.BackColor = &HFFFFFF
  Label5.BackColor = &H0&
  Label6.BackColor = &H0&
  Label7.BackColor = &H0&
  Label8.BackColor = &HFFFFFF
  Label9.BackColor = &H0&
  Label10.BackColor = &H0&
  Label11.BackColor = &H0&
  Label12.BackColor = &H0&
  Label13.BackColor = &H0&
  Label14.BackColor = &HFFFFFF
  Label15.BackColor = &H0&
  Label16.BackColor = &H0&
  Label17.BackColor = &HFFFFFF
  Label18.BackColor = &HFFFFFF
  Label19.BackColor = &HFFFFFF
  Label20.BackColor = &H0&
  Label21.BackColor = &H0&
  Label22.BackColor = &H0&
  Label23.BackColor = &H0&
  Label24.BackColor = &H0&
  Label25.BackColor = &H0&
End If
If a = 8 Then
  Label1.BackColor = &H0&
  Label2.BackColor = &H0&
  Label3.BackColor = &H0&
  Label4.BackColor = &H0&
  Label5.BackColor = &H0&
  Label6.BackColor = &H0&
  Label7.BackColor = &HFFFFFF
  Label8.BackColor = &HFFFFFF
  Label9.BackColor = &HFFFFFF
  Label10.BackColor = &H0&
  Label11.BackColor = &HFFFFFF
  Label12.BackColor = &HFFFFFF
  Label13.BackColor = &HFFFFFF
  Label14.BackColor = &HFFFFFF
  Label15.BackColor = &HFFFFFF
  Label16.BackColor = &H0&
  Label17.BackColor = &HFFFFFF
  Label18.BackColor = &HFFFFFF
  Label19.BackColor = &HFFFFFF
  Label20.BackColor = &H0&
  Label21.BackColor = &H0&
  Label22.BackColor = &H0&
  Label23.BackColor = &H0&
  Label24.BackColor = &H0&
  Label25.BackColor = &H0&
End If
If a = 9 Then
  Label1.BackColor = &H0&
  Label2.BackColor = &H0&
  Label3.BackColor = &H0&
  Label4.BackColor = &H0&
  Label5.BackColor = &HFFFFFF
  Label6.BackColor = &H0&
  Label7.BackColor = &HFFFFFF
  Label8.BackColor = &H0&
  Label9.BackColor = &H0&
  Label10.BackColor = &H0&
  Label11.BackColor = &H0&
  Label12.BackColor = &H0&
  Label13.BackColor = &HFFFFFF
  Label14.BackColor = &H0&
  Label15.BackColor = &H0&
  Label16.BackColor = &H0&
  Label17.BackColor = &HFFFFFF
  Label18.BackColor = &H0&
  Label19.BackColor = &H0&
  Label20.BackColor = &HFFFFFF
  Label21.BackColor = &H0&
  Label22.BackColor = &H0&
  Label23.BackColor = &H0&
  Label24.BackColor = &HFFFFFF
  Label25.BackColor = &H0&
End If
If a = 10 Then
  Label1.BackColor = &HFFFFFF
  Label2.BackColor = &HFFFFFF
  Label3.BackColor = &HFFFFFF
  Label4.BackColor = &HFFFFFF
  Label5.BackColor = &HFFFFFF
  Label6.BackColor = &HFFFFFF
  Label7.BackColor = &HFFFFFF
  Label8.BackColor = &H0&
  Label9.BackColor = &H0&
  Label10.BackColor = &H0&
  Label11.BackColor = &H0&
  Label12.BackColor = &H0&
  Label13.BackColor = &H0&
  Label14.BackColor = &HFFFFFF
  Label15.BackColor = &HFFFFFF
  Label16.BackColor = &H0&
  Label17.BackColor = &HFFFFFF
  Label18.BackColor = &HFFFFFF
  Label19.BackColor = &HFFFFFF
  Label20.BackColor = &H0&
  Label21.BackColor = &H0&
  Label22.BackColor = &H0&
  Label23.BackColor = &HFFFFFF
  Label24.BackColor = &H0&
  Label25.BackColor = &H0&
  LabelP1.Visible = True
  LabelP2.Visible = True
  LabelP2 = "halflife"
End If
If a = 11 Then
  Label1.BackColor = &H0&
  Label2.BackColor = &H0&
  Label3.BackColor = &H0&
  Label4.BackColor = &H0&
  Label5.BackColor = &H0&
  Label6.BackColor = &H0&
  Label7.BackColor = &HFFFFFF
  Label8.BackColor = &HFFFFFF
  Label9.BackColor = &HFFFFFF
  Label10.BackColor = &H0&
  Label11.BackColor = &H0&
  Label12.BackColor = &HFFFFFF
  Label13.BackColor = &HFFFFFF
  Label14.BackColor = &HFFFFFF
  Label15.BackColor = &H0&
  Label16.BackColor = &H0&
  Label17.BackColor = &HFFFFFF
  Label18.BackColor = &HFFFFFF
  Label19.BackColor = &HFFFFFF
  Label20.BackColor = &H0&
  Label21.BackColor = &H0&
  Label22.BackColor = &H0&
  Label23.BackColor = &H0&
  Label24.BackColor = &H0&
  Label25.BackColor = &H0&
  LabelP1.Visible = False
  LabelP2.Visible = False
End If
If a = 12 Then
  Label1.BackColor = &H0&
  Label2.BackColor = &H0&
  Label3.BackColor = &HFFFFFF
  Label4.BackColor = &HFFFFFF
  Label5.BackColor = &HFFFFFF
  Label6.BackColor = &H0&
  Label7.BackColor = &H0&
  Label8.BackColor = &H0&
  Label9.BackColor = &HFFFFFF
  Label10.BackColor = &HFFFFFF
  Label11.BackColor = &HFFFFFF
  Label12.BackColor = &H0&
  Label13.BackColor = &H0&
  Label14.BackColor = &H0&
  Label15.BackColor = &HFFFFFF
  Label16.BackColor = &HFFFFFF
  Label17.BackColor = &HFFFFFF
  Label18.BackColor = &H0&
  Label19.BackColor = &H0&
  Label20.BackColor = &H0&
  Label21.BackColor = &HFFFFFF
  Label22.BackColor = &HFFFFFF
  Label23.BackColor = &HFFFFFF
  Label24.BackColor = &H0&
  Label25.BackColor = &H0&
End If
If a = 13 Then
  Label1.BackColor = &HFFFFFF
  Label2.BackColor = &H0&
  Label3.BackColor = &H0&
  Label4.BackColor = &H0&
  Label5.BackColor = &HFFFFFF
  Label6.BackColor = &H0&
  Label7.BackColor = &HFFFFFF
  Label8.BackColor = &H0&
  Label9.BackColor = &HFFFFFF
  Label10.BackColor = &H0&
  Label11.BackColor = &H0&
  Label12.BackColor = &H0&
  Label13.BackColor = &HFFFFFF
  Label14.BackColor = &H0&
  Label15.BackColor = &H0&
  Label16.BackColor = &H0&
  Label17.BackColor = &HFFFFFF
  Label18.BackColor = &H0&
  Label19.BackColor = &HFFFFFF
  Label20.BackColor = &H0&
  Label21.BackColor = &HFFFFFF
  Label22.BackColor = &H0&
  Label23.BackColor = &H0&
  Label24.BackColor = &H0&
  Label25.BackColor = &HFFFFFF
End If
If a = 14 Then
  Label1.BackColor = &HFFFFFF
  Label2.BackColor = &HFFFFFF
  Label3.BackColor = &H0&
  Label4.BackColor = &HFFFFFF
  Label5.BackColor = &HFFFFFF
  Label6.BackColor = &HFFFFFF
  Label7.BackColor = &HFFFFFF
  Label8.BackColor = &H0&
  Label9.BackColor = &HFFFFFF
  Label10.BackColor = &HFFFFFF
  Label11.BackColor = &HFFFFFF
  Label12.BackColor = &HFFFFFF
  Label13.BackColor = &H0&
  Label14.BackColor = &HFFFFFF
  Label15.BackColor = &HFFFFFF
  Label16.BackColor = &HFFFFFF
  Label17.BackColor = &HFFFFFF
  Label18.BackColor = &H0&
  Label19.BackColor = &HFFFFFF
  Label20.BackColor = &HFFFFFF
  Label21.BackColor = &HFFFFFF
  Label22.BackColor = &HFFFFFF
  Label23.BackColor = &H0&
  Label24.BackColor = &HFFFFFF
  Label25.BackColor = &HFFFFFF
End If
If a = 15 Then
  Label1.BackColor = &H0&
  Label2.BackColor = &H0&
  Label3.BackColor = &H0&
  Label4.BackColor = &H0&
  Label5.BackColor = &H0&
  Label6.BackColor = &H0&
  Label7.BackColor = &HFFFFFF
  Label8.BackColor = &HFFFFFF
  Label9.BackColor = &H0&
  Label10.BackColor = &H0&
  Label11.BackColor = &H0&
  Label12.BackColor = &HFFFFFF
  Label13.BackColor = &HFFFFFF
  Label14.BackColor = &HFFFFFF
  Label15.BackColor = &H0&
  Label16.BackColor = &H0&
  Label17.BackColor = &H0&
  Label18.BackColor = &HFFFFFF
  Label19.BackColor = &HFFFFFF
  Label20.BackColor = &H0&
  Label21.BackColor = &H0&
  Label22.BackColor = &H0&
  Label23.BackColor = &H0&
  Label24.BackColor = &H0&
  Label25.BackColor = &H0&
  LabelP1.Visible = True
  LabelP2.Visible = True
  LabelP2 = "limbo"
End If
If a = 16 Then
  Label1.BackColor = &HFFFFFF
  Label2.BackColor = &H0&
  Label3.BackColor = &HFFFFFF
  Label4.BackColor = &H0&
  Label5.BackColor = &HFFFFFF
  Label6.BackColor = &HFFFFFF
  Label7.BackColor = &H0&
  Label8.BackColor = &HFFFFFF
  Label9.BackColor = &H0&
  Label10.BackColor = &H0&
  Label11.BackColor = &H0&
  Label12.BackColor = &H0&
  Label13.BackColor = &H0&
  Label14.BackColor = &HFFFFFF
  Label15.BackColor = &HFFFFFF
  Label16.BackColor = &H0&
  Label17.BackColor = &HFFFFFF
  Label18.BackColor = &H0&
  Label19.BackColor = &HFFFFFF
  Label20.BackColor = &H0&
  Label21.BackColor = &HFFFFFF
  Label22.BackColor = &H0&
  Label23.BackColor = &H0&
  Label24.BackColor = &HFFFFFF
  Label25.BackColor = &H0&
  LabelP1.Visible = False
  LabelP2.Visible = False
End If
If a = 17 Then
  Label1.BackColor = &H0&
  Label2.BackColor = &H0&
  Label3.BackColor = &HFFFFFF
  Label4.BackColor = &H0&
  Label5.BackColor = &H0&
  Label6.BackColor = &HFFFFFF
  Label7.BackColor = &H0&
  Label8.BackColor = &H0&
  Label9.BackColor = &HFFFFFF
  Label10.BackColor = &H0&
  Label11.BackColor = &H0&
  Label12.BackColor = &HFFFFFF
  Label13.BackColor = &HFFFFFF
  Label14.BackColor = &HFFFFFF
  Label15.BackColor = &HFFFFFF
  Label16.BackColor = &HFFFFFF
  Label17.BackColor = &H0&
  Label18.BackColor = &H0&
  Label19.BackColor = &H0&
  Label20.BackColor = &H0&
  Label21.BackColor = &HFFFFFF
  Label22.BackColor = &HFFFFFF
  Label23.BackColor = &H0&
  Label24.BackColor = &HFFFFFF
  Label25.BackColor = &H0&
End If
If a = 18 Then
  Label1.BackColor = &H0&
  Label2.BackColor = &H0&
  Label3.BackColor = &H0&
  Label4.BackColor = &H0&
  Label5.BackColor = &HFFFFFF
  Label6.BackColor = &H0&
  Label7.BackColor = &H0&
  Label8.BackColor = &H0&
  Label9.BackColor = &HFFFFFF
  Label10.BackColor = &HFFFFFF
  Label11.BackColor = &H0&
  Label12.BackColor = &H0&
  Label13.BackColor = &HFFFFFF
  Label14.BackColor = &H0&
  Label15.BackColor = &HFFFFFF
  Label16.BackColor = &H0&
  Label17.BackColor = &HFFFFFF
  Label18.BackColor = &H0&
  Label19.BackColor = &H0&
  Label20.BackColor = &HFFFFFF
  Label21.BackColor = &HFFFFFF
  Label22.BackColor = &HFFFFFF
  Label23.BackColor = &HFFFFFF
  Label24.BackColor = &HFFFFFF
  Label25.BackColor = &HFFFFFF
End If
If a = 19 Then
  Label1.BackColor = &H0&
  Label2.BackColor = &HFFFFFF
  Label3.BackColor = &H0&
  Label4.BackColor = &HFFFFFF
  Label5.BackColor = &H0&
  Label6.BackColor = &HFFFFFF
  Label7.BackColor = &H0&
  Label8.BackColor = &H0&
  Label9.BackColor = &HFFFFFF
  Label10.BackColor = &H0&
  Label11.BackColor = &H0&
  Label12.BackColor = &HFFFFFF
  Label13.BackColor = &HFFFFFF
  Label14.BackColor = &HFFFFFF
  Label15.BackColor = &HFFFFFF
  Label16.BackColor = &HFFFFFF
  Label17.BackColor = &H0&
  Label18.BackColor = &HFFFFFF
  Label19.BackColor = &H0&
  Label20.BackColor = &H0&
  Label21.BackColor = &H0&
  Label22.BackColor = &HFFFFFF
  Label23.BackColor = &HFFFFFF
  Label24.BackColor = &H0&
  Label25.BackColor = &H0&
End If
If a = 20 Then
  Label1.BackColor = &HFFFFFF
  Label2.BackColor = &HFFFFFF
  Label3.BackColor = &HFFFFFF
  Label4.BackColor = &HFFFFFF
  Label5.BackColor = &HFFFFFF
  Label6.BackColor = &HFFFFFF
  Label7.BackColor = &HFFFFFF
  Label8.BackColor = &HFFFFFF
  Label9.BackColor = &HFFFFFF
  Label10.BackColor = &HFFFFFF
  Label11.BackColor = &HFFFFFF
  Label12.BackColor = &HFFFFFF
  Label13.BackColor = &HFFFFFF
  Label14.BackColor = &HFFFFFF
  Label15.BackColor = &HFFFFFF
  Label16.BackColor = &HFFFFFF
  Label17.BackColor = &HFFFFFF
  Label18.BackColor = &HFFFFFF
  Label19.BackColor = &HFFFFFF
  Label20.BackColor = &HFFFFFF
  Label21.BackColor = &HFFFFFF
  Label22.BackColor = &HFFFFFF
  Label23.BackColor = &HFFFFFF
  Label24.BackColor = &HFFFFFF
  Label25.BackColor = &HFFFFFF
End If
If a = 21 Then
  Label1.Visible = False
  Label2.Visible = False
  Label3.Visible = False
  Label4.Visible = False
  Label5.Visible = False
  Label6.Visible = False
  Label7.Visible = False
  Label8.Visible = False
  Label9.Visible = False
  Label10.Visible = False
  Label11.Visible = False
  Label12.Visible = False
  Label13.Visible = False
  Label14.Visible = False
  Label15.Visible = False
  Label16.Visible = False
  Label17.Visible = False
  Label18.Visible = False
  Label19.Visible = False
  Label20.Visible = False
  Label21.Visible = False
  Label22.Visible = False
  Label23.Visible = False
  Label24.Visible = False
  Label25.Visible = False
  LabelN.Visible = False
  LabelL.Visible = False
  LabelEND.Visible = True
  CommandE.Visible = True
  CommandM.Visible = False
  CommandN.Visible = False
  CommandR.Visible = False
  CommandN.Visible = False
End If
CommandN.Visible = False
End Sub

Private Sub CommandO_Click()
If Text1.Text = "soeasy" Then
  a = 5
  Label1.BackColor = &HFFFFFF
  Label2.BackColor = &H0&
  Label3.BackColor = &H0&
  Label4.BackColor = &H0&
  Label5.BackColor = &HFFFFFF
  Label6.BackColor = &H0&
  Label7.BackColor = &HFFFFFF
  Label8.BackColor = &H0&
  Label9.BackColor = &HFFFFFF
  Label10.BackColor = &H0&
  Label11.BackColor = &H0&
  Label12.BackColor = &HFFFFFF
  Label13.BackColor = &H0&
  Label14.BackColor = &HFFFFFF
  Label15.BackColor = &H0&
  Label16.BackColor = &H0&
  Label17.BackColor = &H0&
  Label18.BackColor = &H0&
  Label19.BackColor = &H0&
  Label20.BackColor = &H0&
  Label21.BackColor = &H0&
  Label22.BackColor = &H0&
  Label23.BackColor = &H0&
  Label24.BackColor = &H0&
  Label25.BackColor = &H0&
  Label1.Visible = True
  Label2.Visible = True
  Label3.Visible = True
  Label4.Visible = True
  Label5.Visible = True
  Label6.Visible = True
  Label7.Visible = True
  Label8.Visible = True
  Label9.Visible = True
  Label10.Visible = True
  Label11.Visible = True
  Label12.Visible = True
  Label13.Visible = True
  Label14.Visible = True
  Label15.Visible = True
  Label16.Visible = True
  Label17.Visible = True
  Label18.Visible = True
  Label19.Visible = True
  Label20.Visible = True
  Label21.Visible = True
  Label22.Visible = True
  Label23.Visible = True
  Label24.Visible = True
  Label25.Visible = True
  LabelN.Visible = True
  LabelL.Visible = True
  LabelW.Visible = False
  LabelP1.Visible = True
  LabelP2.Visible = True
  Text1.Visible = False
  CommandO.Visible = False
  CommandM.Visible = True
  CommandR.Visible = True
  LabelP2 = "soeasy"
  LabelN = a
  Text1.Text = ""
Else
  If Text1.Text = "halflife" Then
    a = 10
    Label1.BackColor = &HFFFFFF
    Label2.BackColor = &HFFFFFF
    Label3.BackColor = &HFFFFFF
    Label4.BackColor = &HFFFFFF
    Label5.BackColor = &HFFFFFF
    Label6.BackColor = &HFFFFFF
    Label7.BackColor = &HFFFFFF
    Label8.BackColor = &H0&
    Label9.BackColor = &H0&
    Label10.BackColor = &H0&
    Label11.BackColor = &H0&
    Label12.BackColor = &H0&
    Label13.BackColor = &H0&
    Label14.BackColor = &HFFFFFF
    Label15.BackColor = &HFFFFFF
    Label16.BackColor = &H0&
    Label17.BackColor = &HFFFFFF
    Label18.BackColor = &HFFFFFF
    Label19.BackColor = &HFFFFFF
    Label20.BackColor = &H0&
    Label21.BackColor = &H0&
    Label22.BackColor = &H0&
    Label23.BackColor = &HFFFFFF
    Label24.BackColor = &H0&
    Label25.BackColor = &H0&
    Label1.Visible = True
    Label2.Visible = True
    Label3.Visible = True
    Label4.Visible = True
    Label5.Visible = True
    Label6.Visible = True
    Label7.Visible = True
    Label8.Visible = True
    Label9.Visible = True
    Label10.Visible = True
    Label11.Visible = True
    Label12.Visible = True
    Label13.Visible = True
    Label14.Visible = True
    Label15.Visible = True
    Label16.Visible = True
    Label17.Visible = True
    Label18.Visible = True
    Label19.Visible = True
    Label20.Visible = True
    Label21.Visible = True
    Label22.Visible = True
    Label23.Visible = True
    Label24.Visible = True
    Label25.Visible = True
    LabelN.Visible = True
    LabelL.Visible = True
    LabelW.Visible = False
    LabelP1.Visible = True
    LabelP2.Visible = True
    Text1.Visible = False
    CommandO.Visible = False
    CommandM.Visible = True
    CommandR.Visible = True
    LabelP2 = "halflife"
    LabelN = a
    Text1.Text = ""
  Else
    If Text1.Text = "limbo" Then
      a = 15
      Label1.BackColor = &H0&
      Label2.BackColor = &H0&
      Label3.BackColor = &H0&
      Label4.BackColor = &H0&
      Label5.BackColor = &H0&
      Label6.BackColor = &H0&
      Label7.BackColor = &HFFFFFF
      Label8.BackColor = &HFFFFFF
      Label9.BackColor = &H0&
      Label10.BackColor = &H0&
      Label11.BackColor = &H0&
      Label12.BackColor = &HFFFFFF
      Label13.BackColor = &HFFFFFF
      Label14.BackColor = &HFFFFFF
      Label15.BackColor = &H0&
      Label16.BackColor = &H0&
      Label17.BackColor = &H0&
      Label18.BackColor = &HFFFFFF
      Label19.BackColor = &HFFFFFF
      Label20.BackColor = &H0&
      Label21.BackColor = &H0&
      Label22.BackColor = &H0&
      Label23.BackColor = &H0&
      Label24.BackColor = &H0&
      Label25.BackColor = &H0&
      Label1.Visible = True
      Label2.Visible = True
      Label3.Visible = True
      Label4.Visible = True
      Label5.Visible = True
      Label6.Visible = True
      Label7.Visible = True
      Label8.Visible = True
      Label9.Visible = True
      Label10.Visible = True
      Label11.Visible = True
      Label12.Visible = True
      Label13.Visible = True
      Label14.Visible = True
      Label15.Visible = True
      Label16.Visible = True
      Label17.Visible = True
      Label18.Visible = True
      Label19.Visible = True
      Label20.Visible = True
      Label21.Visible = True
      Label22.Visible = True
      Label23.Visible = True
      Label24.Visible = True
      Label25.Visible = True
      LabelN.Visible = True
      LabelL.Visible = True
      LabelW.Visible = False
      LabelP1.Visible = True
      LabelP2.Visible = True
      Text1.Visible = False
      CommandO.Visible = False
      CommandM.Visible = True
      CommandR.Visible = True
      LabelP2 = "limbo"
      LabelN = a
      Text1.Text = ""
    Else
      Text1.Text = ""
      LabelT.Visible = True
      LabelW.Visible = False
      Text1.Visible = False
      CommandO.Visible = False
      CommandS.Visible = True
      CommandP.Visible = True
      CommandE.Visible = True
    End If
  End If
End If
End Sub

Private Sub CommandP_Click()
Text1.Visible = True
LabelT.Visible = False
LabelW.Visible = True
CommandO.Visible = True
CommandS.Visible = False
CommandP.Visible = False
CommandE.Visible = False
End Sub

Private Sub CommandR_Click()
If a = 1 Then
  Label1.BackColor = &H0&
  Label2.BackColor = &H0&
  Label3.BackColor = &H0&
  Label4.BackColor = &H0&
  Label5.BackColor = &H0&
  Label6.BackColor = &H0&
  Label7.BackColor = &H0&
  Label8.BackColor = &HFFFFFF
  Label9.BackColor = &H0&
  Label10.BackColor = &H0&
  Label11.BackColor = &H0&
  Label12.BackColor = &HFFFFFF
  Label13.BackColor = &HFFFFFF
  Label14.BackColor = &HFFFFFF
  Label15.BackColor = &H0&
  Label16.BackColor = &H0&
  Label17.BackColor = &H0&
  Label18.BackColor = &HFFFFFF
  Label19.BackColor = &H0&
  Label20.BackColor = &H0&
  Label21.BackColor = &H0&
  Label22.BackColor = &H0&
  Label23.BackColor = &H0&
  Label24.BackColor = &H0&
  Label25.BackColor = &H0&
End If
If a = 2 Then
  Label1.BackColor = &HFFFFFF
  Label2.BackColor = &HFFFFFF
  Label3.BackColor = &H0&
  Label4.BackColor = &H0&
  Label5.BackColor = &H0&
  Label6.BackColor = &HFFFFFF
  Label7.BackColor = &H0&
  Label8.BackColor = &H0&
  Label9.BackColor = &H0&
  Label10.BackColor = &H0&
  Label11.BackColor = &H0&
  Label12.BackColor = &H0&
  Label13.BackColor = &H0&
  Label14.BackColor = &H0&
  Label15.BackColor = &H0&
  Label16.BackColor = &H0&
  Label17.BackColor = &H0&
  Label18.BackColor = &HFFFFFF
  Label19.BackColor = &H0&
  Label20.BackColor = &H0&
  Label21.BackColor = &H0&
  Label22.BackColor = &HFFFFFF
  Label23.BackColor = &HFFFFFF
  Label24.BackColor = &HFFFFFF
  Label25.BackColor = &H0&
End If
If a = 3 Then
  Label1.BackColor = &H0&
  Label2.BackColor = &H0&
  Label3.BackColor = &H0&
  Label4.BackColor = &H0&
  Label5.BackColor = &H0&
  Label6.BackColor = &H0&
  Label7.BackColor = &HFFFFFF
  Label8.BackColor = &HFFFFFF
  Label9.BackColor = &HFFFFFF
  Label10.BackColor = &H0&
  Label11.BackColor = &HFFFFFF
  Label12.BackColor = &H0&
  Label13.BackColor = &HFFFFFF
  Label14.BackColor = &H0&
  Label15.BackColor = &HFFFFFF
  Label16.BackColor = &H0&
  Label17.BackColor = &HFFFFFF
  Label18.BackColor = &HFFFFFF
  Label19.BackColor = &HFFFFFF
  Label20.BackColor = &H0&
  Label21.BackColor = &H0&
  Label22.BackColor = &H0&
  Label23.BackColor = &H0&
  Label24.BackColor = &H0&
  Label25.BackColor = &H0&
End If
If a = 4 Then
  Label1.BackColor = &H0&
  Label2.BackColor = &H0&
  Label3.BackColor = &H0&
  Label4.BackColor = &H0&
  Label5.BackColor = &H0&
  Label6.BackColor = &H0&
  Label7.BackColor = &H0&
  Label8.BackColor = &H0&
  Label9.BackColor = &HFFFFFF
  Label10.BackColor = &HFFFFFF
  Label11.BackColor = &H0&
  Label12.BackColor = &H0&
  Label13.BackColor = &H0&
  Label14.BackColor = &H0&
  Label15.BackColor = &H0&
  Label16.BackColor = &H0&
  Label17.BackColor = &HFFFFFF
  Label18.BackColor = &HFFFFFF
  Label19.BackColor = &H0&
  Label20.BackColor = &H0&
  Label21.BackColor = &H0&
  Label22.BackColor = &H0&
  Label23.BackColor = &HFFFFFF
  Label24.BackColor = &HFFFFFF
  Label25.BackColor = &HFFFFFF
End If
If a = 5 Then
  Label1.BackColor = &HFFFFFF
  Label2.BackColor = &H0&
  Label3.BackColor = &H0&
  Label4.BackColor = &H0&
  Label5.BackColor = &HFFFFFF
  Label6.BackColor = &H0&
  Label7.BackColor = &HFFFFFF
  Label8.BackColor = &H0&
  Label9.BackColor = &HFFFFFF
  Label10.BackColor = &H0&
  Label11.BackColor = &H0&
  Label12.BackColor = &HFFFFFF
  Label13.BackColor = &H0&
  Label14.BackColor = &HFFFFFF
  Label15.BackColor = &H0&
  Label16.BackColor = &H0&
  Label17.BackColor = &H0&
  Label18.BackColor = &H0&
  Label19.BackColor = &H0&
  Label20.BackColor = &H0&
  Label21.BackColor = &H0&
  Label22.BackColor = &H0&
  Label23.BackColor = &H0&
  Label24.BackColor = &H0&
  Label25.BackColor = &H0&
End If
If a = 6 Then
  Label1.BackColor = &H0&
  Label2.BackColor = &H0&
  Label3.BackColor = &H0&
  Label4.BackColor = &H0&
  Label5.BackColor = &H0&
  Label6.BackColor = &HFFFFFF
  Label7.BackColor = &HFFFFFF
  Label8.BackColor = &H0&
  Label9.BackColor = &HFFFFFF
  Label10.BackColor = &HFFFFFF
  Label11.BackColor = &H0&
  Label12.BackColor = &H0&
  Label13.BackColor = &H0&
  Label14.BackColor = &H0&
  Label15.BackColor = &H0&
  Label16.BackColor = &H0&
  Label17.BackColor = &H0&
  Label18.BackColor = &H0&
  Label19.BackColor = &H0&
  Label20.BackColor = &H0&
  Label21.BackColor = &H0&
  Label22.BackColor = &H0&
  Label23.BackColor = &H0&
  Label24.BackColor = &H0&
  Label25.BackColor = &H0&
End If
If a = 7 Then
  Label1.BackColor = &H0&
  Label2.BackColor = &H0&
  Label3.BackColor = &H0&
  Label4.BackColor = &HFFFFFF
  Label5.BackColor = &H0&
  Label6.BackColor = &H0&
  Label7.BackColor = &H0&
  Label8.BackColor = &HFFFFFF
  Label9.BackColor = &H0&
  Label10.BackColor = &H0&
  Label11.BackColor = &H0&
  Label12.BackColor = &H0&
  Label13.BackColor = &H0&
  Label14.BackColor = &HFFFFFF
  Label15.BackColor = &H0&
  Label16.BackColor = &H0&
  Label17.BackColor = &HFFFFFF
  Label18.BackColor = &HFFFFFF
  Label19.BackColor = &HFFFFFF
  Label20.BackColor = &H0&
  Label21.BackColor = &H0&
  Label22.BackColor = &H0&
  Label23.BackColor = &H0&
  Label24.BackColor = &H0&
  Label25.BackColor = &H0&
End If
If a = 8 Then
  Label1.BackColor = &H0&
  Label2.BackColor = &H0&
  Label3.BackColor = &H0&
  Label4.BackColor = &H0&
  Label5.BackColor = &H0&
  Label6.BackColor = &H0&
  Label7.BackColor = &HFFFFFF
  Label8.BackColor = &HFFFFFF
  Label9.BackColor = &HFFFFFF
  Label10.BackColor = &H0&
  Label11.BackColor = &HFFFFFF
  Label12.BackColor = &HFFFFFF
  Label13.BackColor = &HFFFFFF
  Label14.BackColor = &HFFFFFF
  Label15.BackColor = &HFFFFFF
  Label16.BackColor = &H0&
  Label17.BackColor = &HFFFFFF
  Label18.BackColor = &HFFFFFF
  Label19.BackColor = &HFFFFFF
  Label20.BackColor = &H0&
  Label21.BackColor = &H0&
  Label22.BackColor = &H0&
  Label23.BackColor = &H0&
  Label24.BackColor = &H0&
  Label25.BackColor = &H0&
End If
If a = 9 Then
  Label1.BackColor = &H0&
  Label2.BackColor = &H0&
  Label3.BackColor = &H0&
  Label4.BackColor = &H0&
  Label5.BackColor = &HFFFFFF
  Label6.BackColor = &H0&
  Label7.BackColor = &HFFFFFF
  Label8.BackColor = &H0&
  Label9.BackColor = &H0&
  Label10.BackColor = &H0&
  Label11.BackColor = &H0&
  Label12.BackColor = &H0&
  Label13.BackColor = &HFFFFFF
  Label14.BackColor = &H0&
  Label15.BackColor = &H0&
  Label16.BackColor = &H0&
  Label17.BackColor = &HFFFFFF
  Label18.BackColor = &H0&
  Label19.BackColor = &H0&
  Label20.BackColor = &HFFFFFF
  Label21.BackColor = &H0&
  Label22.BackColor = &H0&
  Label23.BackColor = &H0&
  Label24.BackColor = &HFFFFFF
  Label25.BackColor = &H0&
End If
If a = 10 Then
  Label1.BackColor = &HFFFFFF
  Label2.BackColor = &HFFFFFF
  Label3.BackColor = &HFFFFFF
  Label4.BackColor = &HFFFFFF
  Label5.BackColor = &HFFFFFF
  Label6.BackColor = &HFFFFFF
  Label7.BackColor = &HFFFFFF
  Label8.BackColor = &H0&
  Label9.BackColor = &H0&
  Label10.BackColor = &H0&
  Label11.BackColor = &H0&
  Label12.BackColor = &H0&
  Label13.BackColor = &H0&
  Label14.BackColor = &HFFFFFF
  Label15.BackColor = &HFFFFFF
  Label16.BackColor = &H0&
  Label17.BackColor = &HFFFFFF
  Label18.BackColor = &HFFFFFF
  Label19.BackColor = &HFFFFFF
  Label20.BackColor = &H0&
  Label21.BackColor = &H0&
  Label22.BackColor = &H0&
  Label23.BackColor = &HFFFFFF
  Label24.BackColor = &H0&
  Label25.BackColor = &H0&
End If
If a = 11 Then
  Label1.BackColor = &H0&
  Label2.BackColor = &H0&
  Label3.BackColor = &H0&
  Label4.BackColor = &H0&
  Label5.BackColor = &H0&
  Label6.BackColor = &H0&
  Label7.BackColor = &HFFFFFF
  Label8.BackColor = &HFFFFFF
  Label9.BackColor = &HFFFFFF
  Label10.BackColor = &H0&
  Label11.BackColor = &H0&
  Label12.BackColor = &HFFFFFF
  Label13.BackColor = &HFFFFFF
  Label14.BackColor = &HFFFFFF
  Label15.BackColor = &H0&
  Label16.BackColor = &H0&
  Label17.BackColor = &HFFFFFF
  Label18.BackColor = &HFFFFFF
  Label19.BackColor = &HFFFFFF
  Label20.BackColor = &H0&
  Label21.BackColor = &H0&
  Label22.BackColor = &H0&
  Label23.BackColor = &H0&
  Label24.BackColor = &H0&
  Label25.BackColor = &H0&
  LabelP1.Visible = False
  LabelP2.Visible = False
End If
If a = 12 Then
  Label1.BackColor = &H0&
  Label2.BackColor = &H0&
  Label3.BackColor = &HFFFFFF
  Label4.BackColor = &HFFFFFF
  Label5.BackColor = &HFFFFFF
  Label6.BackColor = &H0&
  Label7.BackColor = &H0&
  Label8.BackColor = &H0&
  Label9.BackColor = &HFFFFFF
  Label10.BackColor = &HFFFFFF
  Label11.BackColor = &HFFFFFF
  Label12.BackColor = &H0&
  Label13.BackColor = &H0&
  Label14.BackColor = &H0&
  Label15.BackColor = &HFFFFFF
  Label16.BackColor = &HFFFFFF
  Label17.BackColor = &HFFFFFF
  Label18.BackColor = &H0&
  Label19.BackColor = &H0&
  Label20.BackColor = &H0&
  Label21.BackColor = &HFFFFFF
  Label22.BackColor = &HFFFFFF
  Label23.BackColor = &HFFFFFF
  Label24.BackColor = &H0&
  Label25.BackColor = &H0&
End If
If a = 13 Then
  Label1.BackColor = &HFFFFFF
  Label2.BackColor = &H0&
  Label3.BackColor = &H0&
  Label4.BackColor = &H0&
  Label5.BackColor = &HFFFFFF
  Label6.BackColor = &H0&
  Label7.BackColor = &HFFFFFF
  Label8.BackColor = &H0&
  Label9.BackColor = &HFFFFFF
  Label10.BackColor = &H0&
  Label11.BackColor = &H0&
  Label12.BackColor = &H0&
  Label13.BackColor = &HFFFFFF
  Label14.BackColor = &H0&
  Label15.BackColor = &H0&
  Label16.BackColor = &H0&
  Label17.BackColor = &HFFFFFF
  Label18.BackColor = &H0&
  Label19.BackColor = &HFFFFFF
  Label20.BackColor = &H0&
  Label21.BackColor = &HFFFFFF
  Label22.BackColor = &H0&
  Label23.BackColor = &H0&
  Label24.BackColor = &H0&
  Label25.BackColor = &HFFFFFF
End If
If a = 14 Then
  Label1.BackColor = &HFFFFFF
  Label2.BackColor = &HFFFFFF
  Label3.BackColor = &H0&
  Label4.BackColor = &HFFFFFF
  Label5.BackColor = &HFFFFFF
  Label6.BackColor = &HFFFFFF
  Label7.BackColor = &HFFFFFF
  Label8.BackColor = &H0&
  Label9.BackColor = &HFFFFFF
  Label10.BackColor = &HFFFFFF
  Label11.BackColor = &HFFFFFF
  Label12.BackColor = &HFFFFFF
  Label13.BackColor = &H0&
  Label14.BackColor = &HFFFFFF
  Label15.BackColor = &HFFFFFF
  Label16.BackColor = &HFFFFFF
  Label17.BackColor = &HFFFFFF
  Label18.BackColor = &H0&
  Label19.BackColor = &HFFFFFF
  Label20.BackColor = &HFFFFFF
  Label21.BackColor = &HFFFFFF
  Label22.BackColor = &HFFFFFF
  Label23.BackColor = &H0&
  Label24.BackColor = &HFFFFFF
  Label25.BackColor = &HFFFFFF
End If
If a = 15 Then
  Label1.BackColor = &H0&
  Label2.BackColor = &H0&
  Label3.BackColor = &H0&
  Label4.BackColor = &H0&
  Label5.BackColor = &H0&
  Label6.BackColor = &H0&
  Label7.BackColor = &HFFFFFF
  Label8.BackColor = &HFFFFFF
  Label9.BackColor = &H0&
  Label10.BackColor = &H0&
  Label11.BackColor = &H0&
  Label12.BackColor = &HFFFFFF
  Label13.BackColor = &HFFFFFF
  Label14.BackColor = &HFFFFFF
  Label15.BackColor = &H0&
  Label16.BackColor = &H0&
  Label17.BackColor = &H0&
  Label18.BackColor = &HFFFFFF
  Label19.BackColor = &HFFFFFF
  Label20.BackColor = &H0&
  Label21.BackColor = &H0&
  Label22.BackColor = &H0&
  Label23.BackColor = &H0&
  Label24.BackColor = &H0&
  Label25.BackColor = &H0&
  LabelP1.Visible = True
  LabelP2.Visible = True
  LabelP2 = "limbo"
End If
If a = 16 Then
  Label1.BackColor = &HFFFFFF
  Label2.BackColor = &H0&
  Label3.BackColor = &HFFFFFF
  Label4.BackColor = &H0&
  Label5.BackColor = &HFFFFFF
  Label6.BackColor = &HFFFFFF
  Label7.BackColor = &H0&
  Label8.BackColor = &HFFFFFF
  Label9.BackColor = &H0&
  Label10.BackColor = &H0&
  Label11.BackColor = &H0&
  Label12.BackColor = &H0&
  Label13.BackColor = &H0&
  Label14.BackColor = &HFFFFFF
  Label15.BackColor = &HFFFFFF
  Label16.BackColor = &H0&
  Label17.BackColor = &HFFFFFF
  Label18.BackColor = &H0&
  Label19.BackColor = &HFFFFFF
  Label20.BackColor = &H0&
  Label21.BackColor = &HFFFFFF
  Label22.BackColor = &H0&
  Label23.BackColor = &H0&
  Label24.BackColor = &HFFFFFF
  Label25.BackColor = &H0&
  LabelP1.Visible = False
  LabelP2.Visible = False
End If
If a = 17 Then
  Label1.BackColor = &H0&
  Label2.BackColor = &H0&
  Label3.BackColor = &HFFFFFF
  Label4.BackColor = &H0&
  Label5.BackColor = &H0&
  Label6.BackColor = &HFFFFFF
  Label7.BackColor = &H0&
  Label8.BackColor = &H0&
  Label9.BackColor = &HFFFFFF
  Label10.BackColor = &H0&
  Label11.BackColor = &H0&
  Label12.BackColor = &HFFFFFF
  Label13.BackColor = &HFFFFFF
  Label14.BackColor = &HFFFFFF
  Label15.BackColor = &HFFFFFF
  Label16.BackColor = &HFFFFFF
  Label17.BackColor = &H0&
  Label18.BackColor = &H0&
  Label19.BackColor = &H0&
  Label20.BackColor = &H0&
  Label21.BackColor = &HFFFFFF
  Label22.BackColor = &HFFFFFF
  Label23.BackColor = &H0&
  Label24.BackColor = &HFFFFFF
  Label25.BackColor = &H0&
End If
If a = 18 Then
  Label1.BackColor = &H0&
  Label2.BackColor = &H0&
  Label3.BackColor = &H0&
  Label4.BackColor = &H0&
  Label5.BackColor = &HFFFFFF
  Label6.BackColor = &H0&
  Label7.BackColor = &H0&
  Label8.BackColor = &H0&
  Label9.BackColor = &HFFFFFF
  Label10.BackColor = &HFFFFFF
  Label11.BackColor = &H0&
  Label12.BackColor = &H0&
  Label13.BackColor = &HFFFFFF
  Label14.BackColor = &H0&
  Label15.BackColor = &HFFFFFF
  Label16.BackColor = &H0&
  Label17.BackColor = &HFFFFFF
  Label18.BackColor = &H0&
  Label19.BackColor = &H0&
  Label20.BackColor = &HFFFFFF
  Label21.BackColor = &HFFFFFF
  Label22.BackColor = &HFFFFFF
  Label23.BackColor = &HFFFFFF
  Label24.BackColor = &HFFFFFF
  Label25.BackColor = &HFFFFFF
End If
If a = 19 Then
  Label1.BackColor = &H0&
  Label2.BackColor = &HFFFFFF
  Label3.BackColor = &H0&
  Label4.BackColor = &HFFFFFF
  Label5.BackColor = &H0&
  Label6.BackColor = &HFFFFFF
  Label7.BackColor = &H0&
  Label8.BackColor = &H0&
  Label9.BackColor = &HFFFFFF
  Label10.BackColor = &H0&
  Label11.BackColor = &H0&
  Label12.BackColor = &HFFFFFF
  Label13.BackColor = &HFFFFFF
  Label14.BackColor = &HFFFFFF
  Label15.BackColor = &HFFFFFF
  Label16.BackColor = &HFFFFFF
  Label17.BackColor = &H0&
  Label18.BackColor = &HFFFFFF
  Label19.BackColor = &H0&
  Label20.BackColor = &H0&
  Label21.BackColor = &H0&
  Label22.BackColor = &HFFFFFF
  Label23.BackColor = &HFFFFFF
  Label24.BackColor = &H0&
  Label25.BackColor = &H0&
End If
If a = 20 Then
  Label1.BackColor = &HFFFFFF
  Label2.BackColor = &HFFFFFF
  Label3.BackColor = &HFFFFFF
  Label4.BackColor = &HFFFFFF
  Label5.BackColor = &HFFFFFF
  Label6.BackColor = &HFFFFFF
  Label7.BackColor = &HFFFFFF
  Label8.BackColor = &HFFFFFF
  Label9.BackColor = &HFFFFFF
  Label10.BackColor = &HFFFFFF
  Label11.BackColor = &HFFFFFF
  Label12.BackColor = &HFFFFFF
  Label13.BackColor = &HFFFFFF
  Label14.BackColor = &HFFFFFF
  Label15.BackColor = &HFFFFFF
  Label16.BackColor = &HFFFFFF
  Label17.BackColor = &HFFFFFF
  Label18.BackColor = &HFFFFFF
  Label19.BackColor = &HFFFFFF
  Label20.BackColor = &HFFFFFF
  Label21.BackColor = &HFFFFFF
  Label22.BackColor = &HFFFFFF
  Label23.BackColor = &HFFFFFF
  Label24.BackColor = &HFFFFFF
  Label25.BackColor = &HFFFFFF
End If
CommandN.Visible = False
End Sub

Private Sub CommandS_Click()
a = 1
Label1.Visible = True
Label2.Visible = True
Label3.Visible = True
Label4.Visible = True
Label5.Visible = True
Label6.Visible = True
Label7.Visible = True
Label8.Visible = True
Label9.Visible = True
Label10.Visible = True
Label11.Visible = True
Label12.Visible = True
Label13.Visible = True
Label14.Visible = True
Label15.Visible = True
Label16.Visible = True
Label17.Visible = True
Label18.Visible = True
Label19.Visible = True
Label20.Visible = True
Label21.Visible = True
Label22.Visible = True
Label23.Visible = True
Label24.Visible = True
Label25.Visible = True
LabelN.Visible = True
LabelL.Visible = True
LabelT.Visible = False
CommandM.Visible = True
CommandR.Visible = True
CommandS.Visible = False
CommandP.Visible = False
CommandE.Visible = False
Label1.BackColor = &H0&
Label2.BackColor = &H0&
Label3.BackColor = &H0&
Label4.BackColor = &H0&
Label5.BackColor = &H0&
Label6.BackColor = &H0&
Label7.BackColor = &H0&
Label8.BackColor = &HFFFFFF
Label9.BackColor = &H0&
Label10.BackColor = &H0&
Label11.BackColor = &H0&
Label12.BackColor = &HFFFFFF
Label13.BackColor = &HFFFFFF
Label14.BackColor = &HFFFFFF
Label15.BackColor = &H0&
Label16.BackColor = &H0&
Label17.BackColor = &H0&
Label18.BackColor = &HFFFFFF
Label19.BackColor = &H0&
Label20.BackColor = &H0&
Label21.BackColor = &H0&
Label22.BackColor = &H0&
Label23.BackColor = &H0&
Label24.BackColor = &H0&
Label25.BackColor = &H0&
LabelN = a
End Sub

Private Sub Label1_Click()
If Label1.BackColor = &H0& Then
  Label1.BackColor = &HFFFFFF
Else
  If Label1.BackColor = &HFFFFFF Then
    Label1.BackColor = &H0&
  End If
End If
If Label2.BackColor = &H0& Then
  Label2.BackColor = &HFFFFFF
Else
  If Label2.BackColor = &HFFFFFF Then
    Label2.BackColor = &H0&
  End If
End If
If Label6.BackColor = &H0& Then
  Label6.BackColor = &HFFFFFF
Else
  If Label6.BackColor = &HFFFFFF Then
    Label6.BackColor = &H0&
  End If
End If
If Label1.BackColor = &H0& And Label2.BackColor = &H0& And Label3.BackColor = &H0& And Label4.BackColor = &H0& And Label5.BackColor = &H0& And Label6.BackColor = &H0& And Label7.BackColor = &H0& And Label8.BackColor = &H0& And Label9.BackColor = &H0& And Label10.BackColor = &H0& And Label11.BackColor = &H0& And Label12.BackColor = &H0& And Label13.BackColor = &H0& And Label14.BackColor = &H0& And Label15.BackColor = &H0& And Label16.BackColor = &H0& And Label17.BackColor = &H0& And Label18.BackColor = &H0& And Label19.BackColor = &H0& And Label20.BackColor = &H0& And Label21.BackColor = &H0& And Label22.BackColor = &H0& And Label23.BackColor = &H0& And Label24.BackColor = &H0& And Label25.BackColor = &H0& Then
  CommandN.Visible = True
End If
End Sub

Private Sub Label10_Click()
If Label5.BackColor = &H0& Then
  Label5.BackColor = &HFFFFFF
Else
  If Label5.BackColor = &HFFFFFF Then
    Label5.BackColor = &H0&
  End If
End If
If Label9.BackColor = &H0& Then
  Label9.BackColor = &HFFFFFF
Else
  If Label9.BackColor = &HFFFFFF Then
    Label9.BackColor = &H0&
  End If
End If
If Label10.BackColor = &H0& Then
  Label10.BackColor = &HFFFFFF
Else
  If Label10.BackColor = &HFFFFFF Then
    Label10.BackColor = &H0&
  End If
End If
If Label15.BackColor = &H0& Then
  Label15.BackColor = &HFFFFFF
Else
  If Label15.BackColor = &HFFFFFF Then
    Label15.BackColor = &H0&
  End If
End If
If Label1.BackColor = &H0& And Label2.BackColor = &H0& And Label3.BackColor = &H0& And Label4.BackColor = &H0& And Label5.BackColor = &H0& And Label6.BackColor = &H0& And Label7.BackColor = &H0& And Label8.BackColor = &H0& And Label9.BackColor = &H0& And Label10.BackColor = &H0& And Label11.BackColor = &H0& And Label12.BackColor = &H0& And Label13.BackColor = &H0& And Label14.BackColor = &H0& And Label15.BackColor = &H0& And Label16.BackColor = &H0& And Label17.BackColor = &H0& And Label18.BackColor = &H0& And Label19.BackColor = &H0& And Label20.BackColor = &H0& And Label21.BackColor = &H0& And Label22.BackColor = &H0& And Label23.BackColor = &H0& And Label24.BackColor = &H0& And Label25.BackColor = &H0& Then
  CommandN.Visible = True
End If
End Sub

Private Sub Label11_Click()
If Label6.BackColor = &H0& Then
  Label6.BackColor = &HFFFFFF
Else
  If Label6.BackColor = &HFFFFFF Then
    Label6.BackColor = &H0&
  End If
End If
If Label11.BackColor = &H0& Then
  Label11.BackColor = &HFFFFFF
Else
  If Label11.BackColor = &HFFFFFF Then
    Label11.BackColor = &H0&
  End If
End If
If Label12.BackColor = &H0& Then
  Label12.BackColor = &HFFFFFF
Else
  If Label12.BackColor = &HFFFFFF Then
    Label12.BackColor = &H0&
  End If
End If
If Label16.BackColor = &H0& Then
  Label16.BackColor = &HFFFFFF
Else
  If Label16.BackColor = &HFFFFFF Then
    Label16.BackColor = &H0&
  End If
End If
If Label1.BackColor = &H0& And Label2.BackColor = &H0& And Label3.BackColor = &H0& And Label4.BackColor = &H0& And Label5.BackColor = &H0& And Label6.BackColor = &H0& And Label7.BackColor = &H0& And Label8.BackColor = &H0& And Label9.BackColor = &H0& And Label10.BackColor = &H0& And Label11.BackColor = &H0& And Label12.BackColor = &H0& And Label13.BackColor = &H0& And Label14.BackColor = &H0& And Label15.BackColor = &H0& And Label16.BackColor = &H0& And Label17.BackColor = &H0& And Label18.BackColor = &H0& And Label19.BackColor = &H0& And Label20.BackColor = &H0& And Label21.BackColor = &H0& And Label22.BackColor = &H0& And Label23.BackColor = &H0& And Label24.BackColor = &H0& And Label25.BackColor = &H0& Then
  CommandN.Visible = True
End If
End Sub

Private Sub Label12_Click()
If Label7.BackColor = &H0& Then
  Label7.BackColor = &HFFFFFF
Else
  If Label7.BackColor = &HFFFFFF Then
    Label7.BackColor = &H0&
  End If
End If
If Label11.BackColor = &H0& Then
  Label11.BackColor = &HFFFFFF
Else
  If Label11.BackColor = &HFFFFFF Then
    Label11.BackColor = &H0&
  End If
End If
If Label12.BackColor = &H0& Then
  Label12.BackColor = &HFFFFFF
Else
  If Label12.BackColor = &HFFFFFF Then
    Label12.BackColor = &H0&
  End If
End If
If Label13.BackColor = &H0& Then
  Label13.BackColor = &HFFFFFF
Else
  If Label13.BackColor = &HFFFFFF Then
    Label13.BackColor = &H0&
  End If
End If
If Label17.BackColor = &H0& Then
  Label17.BackColor = &HFFFFFF
Else
  If Label17.BackColor = &HFFFFFF Then
    Label17.BackColor = &H0&
  End If
End If
If Label1.BackColor = &H0& And Label2.BackColor = &H0& And Label3.BackColor = &H0& And Label4.BackColor = &H0& And Label5.BackColor = &H0& And Label6.BackColor = &H0& And Label7.BackColor = &H0& And Label8.BackColor = &H0& And Label9.BackColor = &H0& And Label10.BackColor = &H0& And Label11.BackColor = &H0& And Label12.BackColor = &H0& And Label13.BackColor = &H0& And Label14.BackColor = &H0& And Label15.BackColor = &H0& And Label16.BackColor = &H0& And Label17.BackColor = &H0& And Label18.BackColor = &H0& And Label19.BackColor = &H0& And Label20.BackColor = &H0& And Label21.BackColor = &H0& And Label22.BackColor = &H0& And Label23.BackColor = &H0& And Label24.BackColor = &H0& And Label25.BackColor = &H0& Then
  CommandN.Visible = True
End If
End Sub

Private Sub Label13_Click()
If Label8.BackColor = &H0& Then
  Label8.BackColor = &HFFFFFF
Else
  If Label8.BackColor = &HFFFFFF Then
    Label8.BackColor = &H0&
  End If
End If
If Label12.BackColor = &H0& Then
  Label12.BackColor = &HFFFFFF
Else
  If Label12.BackColor = &HFFFFFF Then
    Label12.BackColor = &H0&
  End If
End If
If Label13.BackColor = &H0& Then
  Label13.BackColor = &HFFFFFF
Else
  If Label13.BackColor = &HFFFFFF Then
    Label13.BackColor = &H0&
  End If
End If
If Label14.BackColor = &H0& Then
  Label14.BackColor = &HFFFFFF
Else
  If Label14.BackColor = &HFFFFFF Then
    Label14.BackColor = &H0&
  End If
End If
If Label18.BackColor = &H0& Then
  Label18.BackColor = &HFFFFFF
Else
  If Label18.BackColor = &HFFFFFF Then
    Label18.BackColor = &H0&
  End If
End If
If Label1.BackColor = &H0& And Label2.BackColor = &H0& And Label3.BackColor = &H0& And Label4.BackColor = &H0& And Label5.BackColor = &H0& And Label6.BackColor = &H0& And Label7.BackColor = &H0& And Label8.BackColor = &H0& And Label9.BackColor = &H0& And Label10.BackColor = &H0& And Label11.BackColor = &H0& And Label12.BackColor = &H0& And Label13.BackColor = &H0& And Label14.BackColor = &H0& And Label15.BackColor = &H0& And Label16.BackColor = &H0& And Label17.BackColor = &H0& And Label18.BackColor = &H0& And Label19.BackColor = &H0& And Label20.BackColor = &H0& And Label21.BackColor = &H0& And Label22.BackColor = &H0& And Label23.BackColor = &H0& And Label24.BackColor = &H0& And Label25.BackColor = &H0& Then
  CommandN.Visible = True
End If
End Sub

Private Sub Label14_Click()
If Label9.BackColor = &H0& Then
  Label9.BackColor = &HFFFFFF
Else
  If Label9.BackColor = &HFFFFFF Then
    Label9.BackColor = &H0&
  End If
End If
If Label13.BackColor = &H0& Then
  Label13.BackColor = &HFFFFFF
Else
  If Label13.BackColor = &HFFFFFF Then
    Label13.BackColor = &H0&
  End If
End If
If Label14.BackColor = &H0& Then
  Label14.BackColor = &HFFFFFF
Else
  If Label14.BackColor = &HFFFFFF Then
    Label14.BackColor = &H0&
  End If
End If
If Label15.BackColor = &H0& Then
  Label15.BackColor = &HFFFFFF
Else
  If Label15.BackColor = &HFFFFFF Then
    Label15.BackColor = &H0&
  End If
End If
If Label19.BackColor = &H0& Then
  Label19.BackColor = &HFFFFFF
Else
  If Label19.BackColor = &HFFFFFF Then
    Label19.BackColor = &H0&
  End If
End If
If Label1.BackColor = &H0& And Label2.BackColor = &H0& And Label3.BackColor = &H0& And Label4.BackColor = &H0& And Label5.BackColor = &H0& And Label6.BackColor = &H0& And Label7.BackColor = &H0& And Label8.BackColor = &H0& And Label9.BackColor = &H0& And Label10.BackColor = &H0& And Label11.BackColor = &H0& And Label12.BackColor = &H0& And Label13.BackColor = &H0& And Label14.BackColor = &H0& And Label15.BackColor = &H0& And Label16.BackColor = &H0& And Label17.BackColor = &H0& And Label18.BackColor = &H0& And Label19.BackColor = &H0& And Label20.BackColor = &H0& And Label21.BackColor = &H0& And Label22.BackColor = &H0& And Label23.BackColor = &H0& And Label24.BackColor = &H0& And Label25.BackColor = &H0& Then
  CommandN.Visible = True
End If
End Sub

Private Sub Label15_Click()
If Label10.BackColor = &H0& Then
  Label10.BackColor = &HFFFFFF
Else
  If Label10.BackColor = &HFFFFFF Then
    Label10.BackColor = &H0&
  End If
End If
If Label14.BackColor = &H0& Then
  Label14.BackColor = &HFFFFFF
Else
  If Label14.BackColor = &HFFFFFF Then
    Label14.BackColor = &H0&
  End If
End If
If Label15.BackColor = &H0& Then
  Label15.BackColor = &HFFFFFF
Else
  If Label15.BackColor = &HFFFFFF Then
    Label15.BackColor = &H0&
  End If
End If
If Label20.BackColor = &H0& Then
  Label20.BackColor = &HFFFFFF
Else
  If Label20.BackColor = &HFFFFFF Then
    Label20.BackColor = &H0&
  End If
End If
If Label1.BackColor = &H0& And Label2.BackColor = &H0& And Label3.BackColor = &H0& And Label4.BackColor = &H0& And Label5.BackColor = &H0& And Label6.BackColor = &H0& And Label7.BackColor = &H0& And Label8.BackColor = &H0& And Label9.BackColor = &H0& And Label10.BackColor = &H0& And Label11.BackColor = &H0& And Label12.BackColor = &H0& And Label13.BackColor = &H0& And Label14.BackColor = &H0& And Label15.BackColor = &H0& And Label16.BackColor = &H0& And Label17.BackColor = &H0& And Label18.BackColor = &H0& And Label19.BackColor = &H0& And Label20.BackColor = &H0& And Label21.BackColor = &H0& And Label22.BackColor = &H0& And Label23.BackColor = &H0& And Label24.BackColor = &H0& And Label25.BackColor = &H0& Then
  CommandN.Visible = True
End If
End Sub

Private Sub Label16_Click()
If Label11.BackColor = &H0& Then
  Label11.BackColor = &HFFFFFF
Else
  If Label11.BackColor = &HFFFFFF Then
    Label11.BackColor = &H0&
  End If
End If
If Label16.BackColor = &H0& Then
  Label16.BackColor = &HFFFFFF
Else
  If Label16.BackColor = &HFFFFFF Then
    Label16.BackColor = &H0&
  End If
End If
If Label17.BackColor = &H0& Then
  Label17.BackColor = &HFFFFFF
Else
  If Label17.BackColor = &HFFFFFF Then
    Label17.BackColor = &H0&
  End If
End If
If Label21.BackColor = &H0& Then
  Label21.BackColor = &HFFFFFF
Else
  If Label21.BackColor = &HFFFFFF Then
    Label21.BackColor = &H0&
  End If
End If
End Sub

Private Sub Label17_Click()
If Label12.BackColor = &H0& Then
  Label12.BackColor = &HFFFFFF
Else
  If Label12.BackColor = &HFFFFFF Then
    Label12.BackColor = &H0&
  End If
End If
If Label16.BackColor = &H0& Then
  Label16.BackColor = &HFFFFFF
Else
  If Label16.BackColor = &HFFFFFF Then
    Label16.BackColor = &H0&
  End If
End If
If Label17.BackColor = &H0& Then
  Label17.BackColor = &HFFFFFF
Else
If Label17.BackColor = &HFFFFFF Then
    Label17.BackColor = &H0&
  End If
End If
If Label18.BackColor = &H0& Then
  Label18.BackColor = &HFFFFFF
Else
  If Label18.BackColor = &HFFFFFF Then
    Label18.BackColor = &H0&
  End If
End If
If Label22.BackColor = &H0& Then
  Label22.BackColor = &HFFFFFF
Else
  If Label22.BackColor = &HFFFFFF Then
    Label22.BackColor = &H0&
  End If
End If
If Label1.BackColor = &H0& And Label2.BackColor = &H0& And Label3.BackColor = &H0& And Label4.BackColor = &H0& And Label5.BackColor = &H0& And Label6.BackColor = &H0& And Label7.BackColor = &H0& And Label8.BackColor = &H0& And Label9.BackColor = &H0& And Label10.BackColor = &H0& And Label11.BackColor = &H0& And Label12.BackColor = &H0& And Label13.BackColor = &H0& And Label14.BackColor = &H0& And Label15.BackColor = &H0& And Label16.BackColor = &H0& And Label17.BackColor = &H0& And Label18.BackColor = &H0& And Label19.BackColor = &H0& And Label20.BackColor = &H0& And Label21.BackColor = &H0& And Label22.BackColor = &H0& And Label23.BackColor = &H0& And Label24.BackColor = &H0& And Label25.BackColor = &H0& Then
  CommandN.Visible = True
End If
End Sub

Private Sub Label18_Click()
If Label13.BackColor = &H0& Then
  Label13.BackColor = &HFFFFFF
Else
  If Label13.BackColor = &HFFFFFF Then
    Label13.BackColor = &H0&
  End If
End If
If Label17.BackColor = &H0& Then
  Label17.BackColor = &HFFFFFF
Else
  If Label17.BackColor = &HFFFFFF Then
    Label17.BackColor = &H0&
  End If
End If
If Label18.BackColor = &H0& Then
  Label18.BackColor = &HFFFFFF
Else
  If Label18.BackColor = &HFFFFFF Then
    Label18.BackColor = &H0&
  End If
End If
If Label19.BackColor = &H0& Then
  Label19.BackColor = &HFFFFFF
Else
  If Label19.BackColor = &HFFFFFF Then
    Label19.BackColor = &H0&
  End If
End If
If Label23.BackColor = &H0& Then
  Label23.BackColor = &HFFFFFF
Else
  If Label23.BackColor = &HFFFFFF Then
    Label23.BackColor = &H0&
  End If
End If
If Label1.BackColor = &H0& And Label2.BackColor = &H0& And Label3.BackColor = &H0& And Label4.BackColor = &H0& And Label5.BackColor = &H0& And Label6.BackColor = &H0& And Label7.BackColor = &H0& And Label8.BackColor = &H0& And Label9.BackColor = &H0& And Label10.BackColor = &H0& And Label11.BackColor = &H0& And Label12.BackColor = &H0& And Label13.BackColor = &H0& And Label14.BackColor = &H0& And Label15.BackColor = &H0& And Label16.BackColor = &H0& And Label17.BackColor = &H0& And Label18.BackColor = &H0& And Label19.BackColor = &H0& And Label20.BackColor = &H0& And Label21.BackColor = &H0& And Label22.BackColor = &H0& And Label23.BackColor = &H0& And Label24.BackColor = &H0& And Label25.BackColor = &H0& Then
  CommandN.Visible = True
End If
End Sub

Private Sub Label19_Click()
If Label14.BackColor = &H0& Then
  Label14.BackColor = &HFFFFFF
Else
  If Label14.BackColor = &HFFFFFF Then
    Label14.BackColor = &H0&
  End If
End If
If Label18.BackColor = &H0& Then
  Label18.BackColor = &HFFFFFF
Else
  If Label18.BackColor = &HFFFFFF Then
    Label18.BackColor = &H0&
  End If
End If
If Label19.BackColor = &H0& Then
  Label19.BackColor = &HFFFFFF
Else
  If Label19.BackColor = &HFFFFFF Then
    Label19.BackColor = &H0&
  End If
End If
If Label20.BackColor = &H0& Then
  Label20.BackColor = &HFFFFFF
Else
  If Label20.BackColor = &HFFFFFF Then
    Label20.BackColor = &H0&
  End If
End If
If Label24.BackColor = &H0& Then
  Label24.BackColor = &HFFFFFF
Else
  If Label24.BackColor = &HFFFFFF Then
    Label24.BackColor = &H0&
  End If
End If
If Label1.BackColor = &H0& And Label2.BackColor = &H0& And Label3.BackColor = &H0& And Label4.BackColor = &H0& And Label5.BackColor = &H0& And Label6.BackColor = &H0& And Label7.BackColor = &H0& And Label8.BackColor = &H0& And Label9.BackColor = &H0& And Label10.BackColor = &H0& And Label11.BackColor = &H0& And Label12.BackColor = &H0& And Label13.BackColor = &H0& And Label14.BackColor = &H0& And Label15.BackColor = &H0& And Label16.BackColor = &H0& And Label17.BackColor = &H0& And Label18.BackColor = &H0& And Label19.BackColor = &H0& And Label20.BackColor = &H0& And Label21.BackColor = &H0& And Label22.BackColor = &H0& And Label23.BackColor = &H0& And Label24.BackColor = &H0& And Label25.BackColor = &H0& Then
  CommandN.Visible = True
End If
End Sub

Private Sub Label2_Click()
If Label1.BackColor = &H0& Then
  Label1.BackColor = &HFFFFFF
Else
  If Label1.BackColor = &HFFFFFF Then
    Label1.BackColor = &H0&
  End If
End If
If Label2.BackColor = &H0& Then
  Label2.BackColor = &HFFFFFF
Else
  If Label2.BackColor = &HFFFFFF Then
    Label2.BackColor = &H0&
  End If
End If
If Label3.BackColor = &H0& Then
  Label3.BackColor = &HFFFFFF
Else
  If Label3.BackColor = &HFFFFFF Then
    Label3.BackColor = &H0&
  End If
End If
If Label7.BackColor = &H0& Then
  Label7.BackColor = &HFFFFFF
Else
  If Label7.BackColor = &HFFFFFF Then
    Label7.BackColor = &H0&
  End If
End If
If Label1.BackColor = &H0& And Label2.BackColor = &H0& And Label3.BackColor = &H0& And Label4.BackColor = &H0& And Label5.BackColor = &H0& And Label6.BackColor = &H0& And Label7.BackColor = &H0& And Label8.BackColor = &H0& And Label9.BackColor = &H0& And Label10.BackColor = &H0& And Label11.BackColor = &H0& And Label12.BackColor = &H0& And Label13.BackColor = &H0& And Label14.BackColor = &H0& And Label15.BackColor = &H0& And Label16.BackColor = &H0& And Label17.BackColor = &H0& And Label18.BackColor = &H0& And Label19.BackColor = &H0& And Label20.BackColor = &H0& And Label21.BackColor = &H0& And Label22.BackColor = &H0& And Label23.BackColor = &H0& And Label24.BackColor = &H0& And Label25.BackColor = &H0& Then
  CommandN.Visible = True
End If
End Sub

Private Sub Label20_Click()
If Label15.BackColor = &H0& Then
  Label15.BackColor = &HFFFFFF
Else
  If Label15.BackColor = &HFFFFFF Then
    Label15.BackColor = &H0&
  End If
End If
If Label19.BackColor = &H0& Then
  Label19.BackColor = &HFFFFFF
Else
  If Label19.BackColor = &HFFFFFF Then
    Label19.BackColor = &H0&
  End If
End If
If Label20.BackColor = &H0& Then
  Label20.BackColor = &HFFFFFF
Else
  If Label20.BackColor = &HFFFFFF Then
    Label20.BackColor = &H0&
  End If
End If
If Label25.BackColor = &H0& Then
  Label25.BackColor = &HFFFFFF
Else
  If Label25.BackColor = &HFFFFFF Then
    Label25.BackColor = &H0&
  End If
End If
If Label1.BackColor = &H0& And Label2.BackColor = &H0& And Label3.BackColor = &H0& And Label4.BackColor = &H0& And Label5.BackColor = &H0& And Label6.BackColor = &H0& And Label7.BackColor = &H0& And Label8.BackColor = &H0& And Label9.BackColor = &H0& And Label10.BackColor = &H0& And Label11.BackColor = &H0& And Label12.BackColor = &H0& And Label13.BackColor = &H0& And Label14.BackColor = &H0& And Label15.BackColor = &H0& And Label16.BackColor = &H0& And Label17.BackColor = &H0& And Label18.BackColor = &H0& And Label19.BackColor = &H0& And Label20.BackColor = &H0& And Label21.BackColor = &H0& And Label22.BackColor = &H0& And Label23.BackColor = &H0& And Label24.BackColor = &H0& And Label25.BackColor = &H0& Then
  CommandN.Visible = True
End If
End Sub

Private Sub Label21_Click()
If Label16.BackColor = &H0& Then
  Label16.BackColor = &HFFFFFF
Else
  If Label16.BackColor = &HFFFFFF Then
    Label16.BackColor = &H0&
  End If
End If
If Label21.BackColor = &H0& Then
  Label21.BackColor = &HFFFFFF
Else
  If Label21.BackColor = &HFFFFFF Then
    Label21.BackColor = &H0&
  End If
End If
If Label22.BackColor = &H0& Then
  Label22.BackColor = &HFFFFFF
Else
  If Label22.BackColor = &HFFFFFF Then
    Label22.BackColor = &H0&
  End If
End If
If Label1.BackColor = &H0& And Label2.BackColor = &H0& And Label3.BackColor = &H0& And Label4.BackColor = &H0& And Label5.BackColor = &H0& And Label6.BackColor = &H0& And Label7.BackColor = &H0& And Label8.BackColor = &H0& And Label9.BackColor = &H0& And Label10.BackColor = &H0& And Label11.BackColor = &H0& And Label12.BackColor = &H0& And Label13.BackColor = &H0& And Label14.BackColor = &H0& And Label15.BackColor = &H0& And Label16.BackColor = &H0& And Label17.BackColor = &H0& And Label18.BackColor = &H0& And Label19.BackColor = &H0& And Label20.BackColor = &H0& And Label21.BackColor = &H0& And Label22.BackColor = &H0& And Label23.BackColor = &H0& And Label24.BackColor = &H0& And Label25.BackColor = &H0& Then
  CommandN.Visible = True
End If
End Sub

Private Sub Label22_Click()
If Label17.BackColor = &H0& Then
  Label17.BackColor = &HFFFFFF
Else
  If Label17.BackColor = &HFFFFFF Then
    Label17.BackColor = &H0&
  End If
End If
If Label21.BackColor = &H0& Then
  Label21.BackColor = &HFFFFFF
Else
  If Label21.BackColor = &HFFFFFF Then
    Label21.BackColor = &H0&
  End If
End If
If Label22.BackColor = &H0& Then
  Label22.BackColor = &HFFFFFF
Else
  If Label22.BackColor = &HFFFFFF Then
    Label22.BackColor = &H0&
  End If
End If
If Label23.BackColor = &H0& Then
  Label23.BackColor = &HFFFFFF
Else
  If Label23.BackColor = &HFFFFFF Then
    Label23.BackColor = &H0&
  End If
End If
If Label1.BackColor = &H0& And Label2.BackColor = &H0& And Label3.BackColor = &H0& And Label4.BackColor = &H0& And Label5.BackColor = &H0& And Label6.BackColor = &H0& And Label7.BackColor = &H0& And Label8.BackColor = &H0& And Label9.BackColor = &H0& And Label10.BackColor = &H0& And Label11.BackColor = &H0& And Label12.BackColor = &H0& And Label13.BackColor = &H0& And Label14.BackColor = &H0& And Label15.BackColor = &H0& And Label16.BackColor = &H0& And Label17.BackColor = &H0& And Label18.BackColor = &H0& And Label19.BackColor = &H0& And Label20.BackColor = &H0& And Label21.BackColor = &H0& And Label22.BackColor = &H0& And Label23.BackColor = &H0& And Label24.BackColor = &H0& And Label25.BackColor = &H0& Then
  CommandN.Visible = True
End If
End Sub

Private Sub Label23_Click()
If Label18.BackColor = &H0& Then
  Label18.BackColor = &HFFFFFF
Else
  If Label18.BackColor = &HFFFFFF Then
    Label18.BackColor = &H0&
  End If
End If
If Label22.BackColor = &H0& Then
  Label22.BackColor = &HFFFFFF
Else
  If Label22.BackColor = &HFFFFFF Then
    Label22.BackColor = &H0&
  End If
End If
If Label23.BackColor = &H0& Then
  Label23.BackColor = &HFFFFFF
Else
  If Label23.BackColor = &HFFFFFF Then
    Label23.BackColor = &H0&
  End If
End If
If Label24.BackColor = &H0& Then
  Label24.BackColor = &HFFFFFF
Else
  If Label24.BackColor = &HFFFFFF Then
    Label24.BackColor = &H0&
  End If
End If
If Label1.BackColor = &H0& And Label2.BackColor = &H0& And Label3.BackColor = &H0& And Label4.BackColor = &H0& And Label5.BackColor = &H0& And Label6.BackColor = &H0& And Label7.BackColor = &H0& And Label8.BackColor = &H0& And Label9.BackColor = &H0& And Label10.BackColor = &H0& And Label11.BackColor = &H0& And Label12.BackColor = &H0& And Label13.BackColor = &H0& And Label14.BackColor = &H0& And Label15.BackColor = &H0& And Label16.BackColor = &H0& And Label17.BackColor = &H0& And Label18.BackColor = &H0& And Label19.BackColor = &H0& And Label20.BackColor = &H0& And Label21.BackColor = &H0& And Label22.BackColor = &H0& And Label23.BackColor = &H0& And Label24.BackColor = &H0& And Label25.BackColor = &H0& Then
  CommandN.Visible = True
End If
End Sub

Private Sub Label24_Click()
If Label19.BackColor = &H0& Then
  Label19.BackColor = &HFFFFFF
Else
  If Label19.BackColor = &HFFFFFF Then
    Label19.BackColor = &H0&
  End If
End If
If Label23.BackColor = &H0& Then
  Label23.BackColor = &HFFFFFF
Else
  If Label23.BackColor = &HFFFFFF Then
    Label23.BackColor = &H0&
  End If
End If
If Label24.BackColor = &H0& Then
  Label24.BackColor = &HFFFFFF
Else
  If Label24.BackColor = &HFFFFFF Then
    Label24.BackColor = &H0&
  End If
End If
If Label25.BackColor = &H0& Then
  Label25.BackColor = &HFFFFFF
Else
  If Label25.BackColor = &HFFFFFF Then
    Label25.BackColor = &H0&
  End If
End If
If Label1.BackColor = &H0& And Label2.BackColor = &H0& And Label3.BackColor = &H0& And Label4.BackColor = &H0& And Label5.BackColor = &H0& And Label6.BackColor = &H0& And Label7.BackColor = &H0& And Label8.BackColor = &H0& And Label9.BackColor = &H0& And Label10.BackColor = &H0& And Label11.BackColor = &H0& And Label12.BackColor = &H0& And Label13.BackColor = &H0& And Label14.BackColor = &H0& And Label15.BackColor = &H0& And Label16.BackColor = &H0& And Label17.BackColor = &H0& And Label18.BackColor = &H0& And Label19.BackColor = &H0& And Label20.BackColor = &H0& And Label21.BackColor = &H0& And Label22.BackColor = &H0& And Label23.BackColor = &H0& And Label24.BackColor = &H0& And Label25.BackColor = &H0& Then
  CommandN.Visible = True
End If
End Sub

Private Sub Label25_Click()
If Label20.BackColor = &H0& Then
  Label20.BackColor = &HFFFFFF
Else
  If Label20.BackColor = &HFFFFFF Then
    Label20.BackColor = &H0&
  End If
End If
If Label24.BackColor = &H0& Then
  Label24.BackColor = &HFFFFFF
Else
  If Label24.BackColor = &HFFFFFF Then
    Label24.BackColor = &H0&
  End If
End If
If Label25.BackColor = &H0& Then
  Label25.BackColor = &HFFFFFF
Else
  If Label25.BackColor = &HFFFFFF Then
    Label25.BackColor = &H0&
  End If
End If
If Label1.BackColor = &H0& And Label2.BackColor = &H0& And Label3.BackColor = &H0& And Label4.BackColor = &H0& And Label5.BackColor = &H0& And Label6.BackColor = &H0& And Label7.BackColor = &H0& And Label8.BackColor = &H0& And Label9.BackColor = &H0& And Label10.BackColor = &H0& And Label11.BackColor = &H0& And Label12.BackColor = &H0& And Label13.BackColor = &H0& And Label14.BackColor = &H0& And Label15.BackColor = &H0& And Label16.BackColor = &H0& And Label17.BackColor = &H0& And Label18.BackColor = &H0& And Label19.BackColor = &H0& And Label20.BackColor = &H0& And Label21.BackColor = &H0& And Label22.BackColor = &H0& And Label23.BackColor = &H0& And Label24.BackColor = &H0& And Label25.BackColor = &H0& Then
  CommandN.Visible = True
End If
End Sub

Private Sub Label3_Click()
If Label2.BackColor = &H0& Then
  Label2.BackColor = &HFFFFFF
Else
  If Label2.BackColor = &HFFFFFF Then
    Label2.BackColor = &H0&
  End If
End If
If Label3.BackColor = &H0& Then
  Label3.BackColor = &HFFFFFF
Else
  If Label3.BackColor = &HFFFFFF Then
    Label3.BackColor = &H0&
  End If
End If
If Label4.BackColor = &H0& Then
  Label4.BackColor = &HFFFFFF
Else
  If Label4.BackColor = &HFFFFFF Then
    Label4.BackColor = &H0&
  End If
End If
If Label8.BackColor = &H0& Then
  Label8.BackColor = &HFFFFFF
Else
  If Label8.BackColor = &HFFFFFF Then
    Label8.BackColor = &H0&
  End If
End If
If Label1.BackColor = &H0& And Label2.BackColor = &H0& And Label3.BackColor = &H0& And Label4.BackColor = &H0& And Label5.BackColor = &H0& And Label6.BackColor = &H0& And Label7.BackColor = &H0& And Label8.BackColor = &H0& And Label9.BackColor = &H0& And Label10.BackColor = &H0& And Label11.BackColor = &H0& And Label12.BackColor = &H0& And Label13.BackColor = &H0& And Label14.BackColor = &H0& And Label15.BackColor = &H0& And Label16.BackColor = &H0& And Label17.BackColor = &H0& And Label18.BackColor = &H0& And Label19.BackColor = &H0& And Label20.BackColor = &H0& And Label21.BackColor = &H0& And Label22.BackColor = &H0& And Label23.BackColor = &H0& And Label24.BackColor = &H0& And Label25.BackColor = &H0& Then
  CommandN.Visible = True
End If
End Sub

Private Sub Label4_Click()
If Label3.BackColor = &H0& Then
  Label3.BackColor = &HFFFFFF
Else
  If Label3.BackColor = &HFFFFFF Then
    Label3.BackColor = &H0&
  End If
End If
If Label4.BackColor = &H0& Then
  Label4.BackColor = &HFFFFFF
Else
  If Label4.BackColor = &HFFFFFF Then
    Label4.BackColor = &H0&
  End If
End If
If Label5.BackColor = &H0& Then
  Label5.BackColor = &HFFFFFF
Else
  If Label5.BackColor = &HFFFFFF Then
    Label5.BackColor = &H0&
  End If
End If
If Label9.BackColor = &H0& Then
  Label9.BackColor = &HFFFFFF
Else
  If Label9.BackColor = &HFFFFFF Then
    Label9.BackColor = &H0&
  End If
End If
If Label1.BackColor = &H0& And Label2.BackColor = &H0& And Label3.BackColor = &H0& And Label4.BackColor = &H0& And Label5.BackColor = &H0& And Label6.BackColor = &H0& And Label7.BackColor = &H0& And Label8.BackColor = &H0& And Label9.BackColor = &H0& And Label10.BackColor = &H0& And Label11.BackColor = &H0& And Label12.BackColor = &H0& And Label13.BackColor = &H0& And Label14.BackColor = &H0& And Label15.BackColor = &H0& And Label16.BackColor = &H0& And Label17.BackColor = &H0& And Label18.BackColor = &H0& And Label19.BackColor = &H0& And Label20.BackColor = &H0& And Label21.BackColor = &H0& And Label22.BackColor = &H0& And Label23.BackColor = &H0& And Label24.BackColor = &H0& And Label25.BackColor = &H0& Then
  CommandN.Visible = True
End If
End Sub

Private Sub Label5_Click()
If Label4.BackColor = &H0& Then
  Label4.BackColor = &HFFFFFF
Else
  If Label4.BackColor = &HFFFFFF Then
    Label4.BackColor = &H0&
  End If
End If
If Label5.BackColor = &H0& Then
  Label5.BackColor = &HFFFFFF
Else
  If Label5.BackColor = &HFFFFFF Then
    Label5.BackColor = &H0&
  End If
End If
If Label10.BackColor = &H0& Then
  Label10.BackColor = &HFFFFFF
Else
  If Label10.BackColor = &HFFFFFF Then
    Label10.BackColor = &H0&
  End If
End If
If Label1.BackColor = &H0& And Label2.BackColor = &H0& And Label3.BackColor = &H0& And Label4.BackColor = &H0& And Label5.BackColor = &H0& And Label6.BackColor = &H0& And Label7.BackColor = &H0& And Label8.BackColor = &H0& And Label9.BackColor = &H0& And Label10.BackColor = &H0& And Label11.BackColor = &H0& And Label12.BackColor = &H0& And Label13.BackColor = &H0& And Label14.BackColor = &H0& And Label15.BackColor = &H0& And Label16.BackColor = &H0& And Label17.BackColor = &H0& And Label18.BackColor = &H0& And Label19.BackColor = &H0& And Label20.BackColor = &H0& And Label21.BackColor = &H0& And Label22.BackColor = &H0& And Label23.BackColor = &H0& And Label24.BackColor = &H0& And Label25.BackColor = &H0& Then
  CommandN.Visible = True
End If
End Sub

Private Sub Label6_Click()
If Label1.BackColor = &H0& Then
  Label1.BackColor = &HFFFFFF
Else
  If Label1.BackColor = &HFFFFFF Then
    Label1.BackColor = &H0&
  End If
End If
If Label6.BackColor = &H0& Then
  Label6.BackColor = &HFFFFFF
Else
  If Label6.BackColor = &HFFFFFF Then
    Label6.BackColor = &H0&
  End If
End If
If Label7.BackColor = &H0& Then
  Label7.BackColor = &HFFFFFF
Else
  If Label7.BackColor = &HFFFFFF Then
    Label7.BackColor = &H0&
  End If
End If
If Label11.BackColor = &H0& Then
  Label11.BackColor = &HFFFFFF
Else
  If Label11.BackColor = &HFFFFFF Then
    Label11.BackColor = &H0&
  End If
End If
If Label1.BackColor = &H0& And Label2.BackColor = &H0& And Label3.BackColor = &H0& And Label4.BackColor = &H0& And Label5.BackColor = &H0& And Label6.BackColor = &H0& And Label7.BackColor = &H0& And Label8.BackColor = &H0& And Label9.BackColor = &H0& And Label10.BackColor = &H0& And Label11.BackColor = &H0& And Label12.BackColor = &H0& And Label13.BackColor = &H0& And Label14.BackColor = &H0& And Label15.BackColor = &H0& And Label16.BackColor = &H0& And Label17.BackColor = &H0& And Label18.BackColor = &H0& And Label19.BackColor = &H0& And Label20.BackColor = &H0& And Label21.BackColor = &H0& And Label22.BackColor = &H0& And Label23.BackColor = &H0& And Label24.BackColor = &H0& And Label25.BackColor = &H0& Then
  CommandN.Visible = True
End If
End Sub

Private Sub Label7_Click()
If Label2.BackColor = &H0& Then
  Label2.BackColor = &HFFFFFF
Else
  If Label2.BackColor = &HFFFFFF Then
    Label2.BackColor = &H0&
  End If
End If
If Label6.BackColor = &H0& Then
  Label6.BackColor = &HFFFFFF
Else
  If Label6.BackColor = &HFFFFFF Then
    Label6.BackColor = &H0&
  End If
End If
If Label7.BackColor = &H0& Then
  Label7.BackColor = &HFFFFFF
Else
  If Label7.BackColor = &HFFFFFF Then
    Label7.BackColor = &H0&
  End If
End If
If Label8.BackColor = &H0& Then
  Label8.BackColor = &HFFFFFF
Else
  If Label8.BackColor = &HFFFFFF Then
    Label8.BackColor = &H0&
  End If
End If
If Label12.BackColor = &H0& Then
  Label12.BackColor = &HFFFFFF
Else
  If Label12.BackColor = &HFFFFFF Then
    Label12.BackColor = &H0&
  End If
End If
If Label1.BackColor = &H0& And Label2.BackColor = &H0& And Label3.BackColor = &H0& And Label4.BackColor = &H0& And Label5.BackColor = &H0& And Label6.BackColor = &H0& And Label7.BackColor = &H0& And Label8.BackColor = &H0& And Label9.BackColor = &H0& And Label10.BackColor = &H0& And Label11.BackColor = &H0& And Label12.BackColor = &H0& And Label13.BackColor = &H0& And Label14.BackColor = &H0& And Label15.BackColor = &H0& And Label16.BackColor = &H0& And Label17.BackColor = &H0& And Label18.BackColor = &H0& And Label19.BackColor = &H0& And Label20.BackColor = &H0& And Label21.BackColor = &H0& And Label22.BackColor = &H0& And Label23.BackColor = &H0& And Label24.BackColor = &H0& And Label25.BackColor = &H0& Then
  CommandN.Visible = True
End If
End Sub

Private Sub Label8_Click()
If Label3.BackColor = &H0& Then
  Label3.BackColor = &HFFFFFF
Else
  If Label3.BackColor = &HFFFFFF Then
    Label3.BackColor = &H0&
  End If
End If
If Label7.BackColor = &H0& Then
  Label7.BackColor = &HFFFFFF
Else
  If Label7.BackColor = &HFFFFFF Then
    Label7.BackColor = &H0&
  End If
End If
If Label8.BackColor = &H0& Then
  Label8.BackColor = &HFFFFFF
Else
  If Label8.BackColor = &HFFFFFF Then
    Label8.BackColor = &H0&
  End If
End If
If Label9.BackColor = &H0& Then
  Label9.BackColor = &HFFFFFF
Else
  If Label9.BackColor = &HFFFFFF Then
    Label9.BackColor = &H0&
  End If
End If
If Label13.BackColor = &H0& Then
  Label13.BackColor = &HFFFFFF
Else
  If Label13.BackColor = &HFFFFFF Then
    Label13.BackColor = &H0&
  End If
End If
If Label1.BackColor = &H0& And Label2.BackColor = &H0& And Label3.BackColor = &H0& And Label4.BackColor = &H0& And Label5.BackColor = &H0& And Label6.BackColor = &H0& And Label7.BackColor = &H0& And Label8.BackColor = &H0& And Label9.BackColor = &H0& And Label10.BackColor = &H0& And Label11.BackColor = &H0& And Label12.BackColor = &H0& And Label13.BackColor = &H0& And Label14.BackColor = &H0& And Label15.BackColor = &H0& And Label16.BackColor = &H0& And Label17.BackColor = &H0& And Label18.BackColor = &H0& And Label19.BackColor = &H0& And Label20.BackColor = &H0& And Label21.BackColor = &H0& And Label22.BackColor = &H0& And Label23.BackColor = &H0& And Label24.BackColor = &H0& And Label25.BackColor = &H0& Then
  CommandN.Visible = True
End If
End Sub

Private Sub Label9_Click()
If Label4.BackColor = &H0& Then
  Label4.BackColor = &HFFFFFF
Else
  If Label4.BackColor = &HFFFFFF Then
    Label4.BackColor = &H0&
  End If
End If
If Label8.BackColor = &H0& Then
  Label8.BackColor = &HFFFFFF
Else
  If Label8.BackColor = &HFFFFFF Then
    Label8.BackColor = &H0&
  End If
End If
If Label9.BackColor = &H0& Then
  Label9.BackColor = &HFFFFFF
Else
  If Label9.BackColor = &HFFFFFF Then
    Label9.BackColor = &H0&
  End If
End If
If Label10.BackColor = &H0& Then
  Label10.BackColor = &HFFFFFF
Else
  If Label10.BackColor = &HFFFFFF Then
    Label10.BackColor = &H0&
  End If
End If
If Label14.BackColor = &H0& Then
  Label14.BackColor = &HFFFFFF
Else
  If Label14.BackColor = &HFFFFFF Then
    Label14.BackColor = &H0&
  End If
End If
If Label1.BackColor = &H0& And Label2.BackColor = &H0& And Label3.BackColor = &H0& And Label4.BackColor = &H0& And Label5.BackColor = &H0& And Label6.BackColor = &H0& And Label7.BackColor = &H0& And Label8.BackColor = &H0& And Label9.BackColor = &H0& And Label10.BackColor = &H0& And Label11.BackColor = &H0& And Label12.BackColor = &H0& And Label13.BackColor = &H0& And Label14.BackColor = &H0& And Label15.BackColor = &H0& And Label16.BackColor = &H0& And Label17.BackColor = &H0& And Label18.BackColor = &H0& And Label19.BackColor = &H0& And Label20.BackColor = &H0& And Label21.BackColor = &H0& And Label22.BackColor = &H0& And Label23.BackColor = &H0& And Label24.BackColor = &H0& And Label25.BackColor = &H0& Then
  CommandN.Visible = True
End If
End Sub
