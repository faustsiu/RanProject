VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Attack"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5910
   LinkMode        =   1  '來源
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   5910
   StartUpPosition =   3  '系統預設值
   Begin VB.ComboBox Combo3 
      Height          =   300
      ItemData        =   "Form1.frx":0000
      Left            =   1560
      List            =   "Form1.frx":000A
      TabIndex        =   19
      Top             =   1200
      Width           =   1455
   End
   Begin VB.CommandButton Command6 
      Caption         =   "StopMission"
      Height          =   255
      Left            =   2040
      TabIndex        =   18
      Top             =   840
      Width           =   975
   End
   Begin VB.ComboBox Combo2 
      Height          =   300
      ItemData        =   "Form1.frx":0026
      Left            =   120
      List            =   "Form1.frx":003C
      TabIndex        =   17
      Top             =   120
      Width           =   1575
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      Left            =   120
      TabIndex        =   16
      Text            =   "Combo1"
      Top             =   480
      Width           =   1575
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Set Regen Level"
      Height          =   255
      Left            =   3240
      TabIndex        =   15
      Top             =   1200
      Width           =   1335
   End
   Begin VB.TextBox Text8 
      Height          =   270
      Left            =   3600
      TabIndex        =   11
      Text            =   "50"
      Top             =   840
      Width           =   975
   End
   Begin VB.TextBox Text7 
      Height          =   270
      Left            =   3600
      TabIndex        =   10
      Text            =   "10"
      Top             =   480
      Width           =   975
   End
   Begin VB.TextBox Text6 
      Height          =   270
      Left            =   3600
      TabIndex        =   9
      Text            =   "1800"
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Run Game"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1920
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Mission"
      Height          =   255
      Left            =   2040
      TabIndex        =   7
      Top             =   480
      Width           =   975
   End
   Begin VB.TextBox Text5 
      Height          =   270
      Left            =   2040
      TabIndex        =   6
      Text            =   "0"
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Attack"
      Height          =   255
      Left            =   4800
      TabIndex        =   5
      Top             =   840
      Width           =   975
   End
   Begin VB.TextBox Text4 
      Height          =   270
      Left            =   4800
      TabIndex        =   4
      Text            =   "Text4"
      Top             =   480
      Width           =   975
   End
   Begin VB.TextBox Text3 
      Height          =   270
      Left            =   4800
      TabIndex        =   3
      Text            =   "Text3"
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Gate Entry"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Height          =   270
      Left            =   120
      TabIndex        =   1
      Text            =   "0"
      Top             =   840
      Width           =   975
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   5280
      Top             =   2280
   End
   Begin VB.TextBox Text1 
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   2280
      Width           =   5655
   End
   Begin VB.Label Label3 
      Caption         =   "SP"
      Height          =   255
      Left            =   3240
      TabIndex        =   14
      Top             =   840
      Width           =   375
   End
   Begin VB.Label Label2 
      Caption         =   "MP"
      Height          =   255
      Left            =   3240
      TabIndex        =   13
      Top             =   480
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "HP"
      Height          =   255
      Left            =   3240
      TabIndex        =   12
      Top             =   120
      Width           =   375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub socket_hook Lib "ws2_32.dll" (ByVal switch As Integer)

Private Declare Function GetHookerStatus Lib "ws2_hook.dll" (ByVal enable As Long) As Long
Private Declare Sub AutoRegenPP Lib "ws2_hook.dll" (ByVal enable As Integer, ByVal HP As Integer, ByVal MP As Integer, ByVal SP As Integer)
Private Declare Sub AutoPicking Lib "ws2_hook.dll" (ByVal enable As Integer)
Private Declare Sub EntryArea Lib "ws2_hook.dll" (ByVal GateNumber As Integer)
Private Declare Sub AutoMission Lib "ws2_hook.dll" (ByVal enable As Integer)
Private Declare Sub AttackTest Lib "ws2_hook.dll" (ByVal Skill As Long, ByVal Delay As Long)
Private Declare Sub GetSysInfo Lib "ws2_hook.dll" (ByVal Info As Integer)

Dim gMapID As Integer



Private Sub Combo1_Click()
    If Left(Combo1.Text, 2) <> "  " Then
        EntryArea Val(Left(Combo1.Text, 2))
    End If

End Sub

Private Sub Combo2_Click()
    AutoPicking (Val(Left(Combo2.Text, 1)))
End Sub

Private Sub Combo3_Click()
    GetSysInfo (Val(Left(Combo3.Text, 1)))
End Sub

Private Sub Command1_Click()
    EntryArea (Val(Text2.Text))
End Sub

Private Sub Command2_Click()
    AttackTest Val(Text3.Text), Val(Text4.Text)
    
End Sub

Private Sub Command3_Click()
        AutoMission Val(Text5.Text)
End Sub

Private Sub Command4_Click()
    Shell ("Game.exe /app_run")
End Sub

Private Sub Command5_Click()
    AutoRegenPP 1, Val(Text6.Text), Val(Text7.Text), Val(Text8.Text)
End Sub

Private Sub Command6_Click()
 AutoMission 0
End Sub

Private Sub Form_Load()
    gMapID = -1
    Combo2.ListIndex = 1
    Combo3.ListIndex = 0
    socket_hook (1)
    AutoPicking (1)
End Sub

Function ToAreaName(ByVal Number As Integer) As String
   If Number = 3 Then
        ToAreaName = "03"
    ElseIf Number = 16 Then
        ToAreaName = "16"
    ElseIf Number = 28 Then
        ToAreaName = "28"
    ElseIf Number = 29 Then
        ToAreaName = "29"
    End If

End Function


Private Sub Form_Terminate()
    socket_hook (0)
End Sub

Private Sub Timer1_Timer()
    Area = GetHookerStatus(0)
    If Area >= 0 And Area <= 300 Then
        MapName (GetHookerStatus(0))
    End If
'    MapName (120)
    Text1.Text = GetHookerStatus(0) & ":" & GetHookerStatus(1) & ":" & GetHookerStatus(2) & ":" & GetHookerStatus(3) & ":" & GetHookerStatus(4)
'    Text1.Text = hook_test(0) & ":" & Hex$(hook_test(1)) & ":" & Hex$(hook_test(2)) & ":" & Hex$(hook_test(3))
    
End Sub



Private Sub MapName(ByVal NewMapID As Integer)
    If NewMapID <> gMapID Then
    
    gMapID = NewMapID
    
    Combo1.Clear
    Select Case gMapID
    
Case 0
 Combo1.AddItem "  聖門本館1F"
 Combo1.AddItem "6 聖門學院"

Case 2
 Combo1.AddItem "  聖門學院"
 Combo1.AddItem "1 聖門本館1F"
 Combo1.AddItem "2 聖門洞"
 Combo1.AddItem "3 綜合修煉院"
 Combo1.AddItem "20學院廣場"

Case 100
 Combo1.AddItem "  學習1館"

Case 101
 Combo1.AddItem "  學習2館"

Case 103
 Combo1.AddItem "  學習3館"

Case 104
 Combo1.AddItem "  歷史館"

Case 105
 Combo1.AddItem "  圖書館"

Case 106
 Combo1.AddItem "  社團室"

Case 107
 Combo1.AddItem "  科學館"

Case 102
 Combo1.AddItem "  宿舍"

Case 3
 Combo1.AddItem "  聖門洞"
 Combo1.AddItem "1 聖門學院"
 Combo1.AddItem "2 玄巖洞"
 Combo1.AddItem "3 鳳凰洞"
 Combo1.AddItem "4 虎令學院"
 Combo1.AddItem "12商洞"

Case 4
 Combo1.AddItem "  玄巖本館1F"

Case 5
 Combo1.AddItem "  玄巖學院"
 Combo1.AddItem "20學院廣場"

Case 120
 Combo1.AddItem "  學習2館"

Case 121
 Combo1.AddItem "  學習1館"

Case 123
 Combo1.AddItem "  藝術館"

Case 122
 Combo1.AddItem "  圖書館"

Case 124
 Combo1.AddItem "  學生1館"

Case 125
 Combo1.AddItem "  學生2館"

Case 126
 Combo1.AddItem "  宿舍"

Case 128
 Combo1.AddItem "  宿舍別館"

Case 6
 Combo1.AddItem "  玄巖洞"
 Combo1.AddItem "2 聖門洞"

Case 7
 Combo1.AddItem "  鳳凰本館1F"

Case 8
 Combo1.AddItem "  鳳凰學院"
 Combo1.AddItem "20學院廣場"

Case 130
 Combo1.AddItem "  學習2館"

Case 131
 Combo1.AddItem "  學習1館"

Case 136
 Combo1.AddItem "  整修館"

Case 135
 Combo1.AddItem "  科學館"

Case 137
 Combo1.AddItem "  實驗館"

Case 138
 Combo1.AddItem "  實習館"

Case 141
 Combo1.AddItem "  輔助器材管理館"

Case 142
 Combo1.AddItem "  圖書館"

Case 134
 Combo1.AddItem "  宿舍"

Case 9
 Combo1.AddItem "  鳳凰洞"
 Combo1.AddItem "2 聖門洞"

Case 10
 Combo1.AddItem "  虎令學院"
 Combo1.AddItem "1 虎令本館1F"
 Combo1.AddItem "2 聖門洞"

Case 11
 Combo1.AddItem "  虎令本館1F"
 Combo1.AddItem "1 虎令本館2F"
 Combo1.AddItem "6 虎令學院"

Case 12
 Combo1.AddItem "  虎令本館2F"
 Combo1.AddItem "1 虎令本館1F"
 Combo1.AddItem "3 虎令本館3F"
 Combo1.AddItem "5 虎令本館B1"

Case 13
 Combo1.AddItem "  虎令本館3F"
 Combo1.AddItem "3 虎令本館2F"
 Combo1.AddItem "5 虎令本館B2"

Case 14
 Combo1.AddItem "  虎令本館B1"
 Combo1.AddItem "1 虎令本館2F"

Case 15
 Combo1.AddItem "  虎令本館B2"
 Combo1.AddItem "1 虎令本館3F"
 Combo1.AddItem "2 虎令本館B3"

Case 30
 Combo1.AddItem "  虎令本館B3"
 Combo1.AddItem "0 本館"
 Combo1.AddItem "2 虎令本館B2"
 Combo1.AddItem "4 異界虎令本館B3"

Case 17
 Combo1.AddItem "  聖東隧道"

Case 16
 Combo1.AddItem "  商洞"
 Combo1.AddItem "0 本館"
 Combo1.AddItem "1 聖門洞"
 Combo1.AddItem "8 綜合碼頭"
 Combo1.AddItem "10中洞"
 Combo1.AddItem "11商洞3號隧道"
 Combo1.AddItem "15綜合運動場"
 Combo1.AddItem "18火炎地獄"

Case 161
 Combo1.AddItem "  1號 地下停車場"

Case 162
 Combo1.AddItem "  2號 地下停車場"

Case 18
 Combo1.AddItem "  綜合碼頭"
 Combo1.AddItem "1 商洞"
 Combo1.AddItem "3 青基地1F_A區"
 Combo1.AddItem "5 青基地1F_B區"
 Combo1.AddItem "6 青基地1F_C區"

Case 19
 Combo1.AddItem "  青基地1F_A區"
 Combo1.AddItem "0 綜合碼頭"
 Combo1.AddItem "1 青基地2F_A區"

Case 20
 Combo1.AddItem "  青基地2F_A區"
 Combo1.AddItem "0 青基地3F_A區"
 Combo1.AddItem "1 青基地1F_A區"
 Combo1.AddItem "2 青基地1F_A區"

Case 21
 Combo1.AddItem "  青基地3F_A區"
 Combo1.AddItem "0 青基地2F_A區"

Case 150
 Combo1.AddItem "  青基地1F_B區"
 Combo1.AddItem "0 綜合碼頭"
 Combo1.AddItem "1 青基地2F_B區"

Case 152
 Combo1.AddItem "  青基地2F_B區"
 Combo1.AddItem "0 青基地3F_B區"
 Combo1.AddItem "1 青基地1F_B區"

Case 154
 Combo1.AddItem "  青基地3F_B區"
 Combo1.AddItem "0 青基地2F_B區"

Case 151
 Combo1.AddItem "  青基地1F_C區"
 Combo1.AddItem "0 綜合碼頭"
 Combo1.AddItem "1 青基地2F_C區"

Case 153
 Combo1.AddItem "  青基地2F_C區"
 Combo1.AddItem "0 青基地3F_C區"
 Combo1.AddItem "1 青基地1F_C區"

Case 155
 Combo1.AddItem "  青基地3F_C區"
 Combo1.AddItem "0 青基地2F_C區"

Case 26
 Combo1.AddItem "  商洞3號隧道"
 Combo1.AddItem "0 商洞"
 Combo1.AddItem "1 聖財團私立監獄"

Case 27
 Combo1.AddItem "  聖財團私立監獄"
 Combo1.AddItem "0 商洞3號隧道"
 Combo1.AddItem "1 本館"
 Combo1.AddItem "4 監獄實驗場"
 Combo1.AddItem "5 死牢入口"

Case 28
 Combo1.AddItem "  中洞"
 Combo1.AddItem "0 商洞"
 Combo1.AddItem "2 本洞"

Case 29
 Combo1.AddItem "  本洞"
 Combo1.AddItem "0 中洞"
 Combo1.AddItem "1 本館"
 Combo1.AddItem "2 聖財團1F"

Case 22
 Combo1.AddItem "  學院廣場"
 Combo1.AddItem "1 聖門學院"
 Combo1.AddItem "2 玄巖學院"
 Combo1.AddItem "3 鳳凰學院"
 Combo1.AddItem "11聖門電算室入口"
 Combo1.AddItem "12玄巖電算室入口"
 Combo1.AddItem "13鳳凰電算室入口"
 Combo1.AddItem "14商洞電算室入口"
 Combo1.AddItem "  財團附屬停車場B"
 Combo1.AddItem "  財團附屬停車場C"
 Combo1.AddItem "  財團附屬停車場D"
 Combo1.AddItem "95財團附屬停車場E"
 Combo1.AddItem "5 休息會館"
 Combo1.AddItem "  結婚宴會廳"
 Combo1.AddItem "  活動宴會廳"

Case 23
 Combo1.AddItem "  綜合修煉院"
 Combo1.AddItem "2 聖門學院"
 Combo1.AddItem "  玄巖學院"
 Combo1.AddItem "  鳳凰學院"

Case 31
 Combo1.AddItem "  結婚宴會廳"
 Combo1.AddItem "0 學院廣場"

Case 201
 Combo1.AddItem "  聖門電算室入口"
 Combo1.AddItem "0 學院廣場"
 Combo1.AddItem "1 聖門電算室"

Case 211
 Combo1.AddItem "  聖門洞電算室"

Case 202
 Combo1.AddItem "  玄巖電算室入口"
 Combo1.AddItem "0 學院廣場"
 Combo1.AddItem "1 玄巖洞電算室"

Case 212
 Combo1.AddItem "  玄巖洞電算室"

Case 203
 Combo1.AddItem "  鳳凰電算室入口"
 Combo1.AddItem "0 學院廣場"
 Combo1.AddItem "1 鳳凰電算室"

Case 213
 Combo1.AddItem "  鳳凰洞電算室"

Case 204
 Combo1.AddItem "  商洞電算室入口"
 Combo1.AddItem "0 學院廣場"

Case 214
 Combo1.AddItem "  商洞電算室"

Case 32
 Combo1.AddItem "  監獄實驗場"
 Combo1.AddItem "0 聖財團私立監獄"
 Combo1.AddItem "1 本館"
 Combo1.AddItem "2 第七實驗室"

Case 33
 Combo1.AddItem "  第七實驗室"
 Combo1.AddItem "0 本館"
 Combo1.AddItem "  監獄實驗場"

Case 34
 Combo1.AddItem "  涉谷"

Case 46
 Combo1.AddItem "  聖財團1F"
 Combo1.AddItem "0 本洞"
 Combo1.AddItem "1 本館"
 Combo1.AddItem "5 聖財團地下室"
 Combo1.AddItem "2 聖財團2F"
 Combo1.AddItem "3 聖財團30F"

Case 47
 Combo1.AddItem "  聖財團2F"
 Combo1.AddItem "2 聖財團1F"
 Combo1.AddItem "4 聖財團3F"

Case 48
 Combo1.AddItem "  聖財團3F"
 Combo1.AddItem "1 聖財團2F"

Case 35
 Combo1.AddItem "  聖財團30F"
 Combo1.AddItem "1 本館"
 Combo1.AddItem "2 聖財團50F"
 Combo1.AddItem "3 聖財團1F"

Case 36
 Combo1.AddItem "  聖財團50F"
 Combo1.AddItem "1 本館"
 Combo1.AddItem "  聖財團30F"
 Combo1.AddItem "4 聖財團80F"
 Combo1.AddItem "  聖財團90F"

Case 37
 Combo1.AddItem "  聖財團90F"
 Combo1.AddItem "0 聖財團50F"
 Combo1.AddItem "1 本館"
 Combo1.AddItem "  聖財團80樓"

Case 97
 Combo1.AddItem "  財團附屬停車場B"

Case 98
 Combo1.AddItem "  財團附屬停車場C"

Case 99
 Combo1.AddItem "  財團附屬停車場D"

Case 95
 Combo1.AddItem "  財團附屬停車場E1"

Case 38
 Combo1.AddItem "  聖財團大樓_左翼"
 Combo1.AddItem "0 聖財團80樓"
 Combo1.AddItem "1 本館"
 Combo1.AddItem "2 聖財團密室左區"

Case 39
 Combo1.AddItem "  聖財團大樓_右翼"
 Combo1.AddItem "0 聖財團80樓"
 Combo1.AddItem "2 聖財團密室右區"

Case 40
 Combo1.AddItem "  聖財團密室左區"
 Combo1.AddItem "  聖財團80樓"
 Combo1.AddItem "  聖財團大樓_左翼"

Case 41
 Combo1.AddItem "  聖財團密室右區"
 Combo1.AddItem "  聖財團大樓_右翼"
 Combo1.AddItem "  異界殿堂1F"

Case 42
 Combo1.AddItem "  聖財團地下室"
 Combo1.AddItem "0 聖財團1F"

Case 43
 Combo1.AddItem "  異界殿堂1F"

Case 44
 Combo1.AddItem "  異界殿堂2F"

Case 45
 Combo1.AddItem "  異界神殿"

Case 50
 Combo1.AddItem "  神殿"

Case 215
 Combo1.AddItem "  休閒會館"
 Combo1.AddItem "0 學院廣場"

Case 220
 Combo1.AddItem "  綜合運動場"
 Combo1.AddItem "0 商洞"

Case 51
 Combo1.AddItem "  聖財團80樓"
 Combo1.AddItem "0 聖財團90F"
 Combo1.AddItem "1 本館"
 Combo1.AddItem "2 聖財團1F"
 Combo1.AddItem "3 聖財團50F"
 Combo1.AddItem "5 聖財團大樓_左翼"
 Combo1.AddItem "7 聖財團大樓_右翼"

Case 216
 Combo1.AddItem "  聖門電算密室"

Case 217
 Combo1.AddItem "  玄嚴電算密室"

Case 218
 Combo1.AddItem "  鳳凰電算密室"

Case 219
 Combo1.AddItem "  商洞電算密室"

Case 225
 Combo1.AddItem "  火炎地獄"
 Combo1.AddItem "0 商洞"
 Combo1.AddItem "1 本館"
 Combo1.AddItem "2 極凍領域"
 Combo1.AddItem "4 雷電幻境"
 Combo1.AddItem "6 巫毒神殿"

Case 226
 Combo1.AddItem "  極凍領域"
 Combo1.AddItem "  火炎地獄"

Case 227
 Combo1.AddItem "  雷電幻境"
 Combo1.AddItem "  火炎地獄"

Case 228
 Combo1.AddItem "  巫毒神殿"
 Combo1.AddItem "  火炎地獄"
 Combo1.AddItem "3 虛空要塞"

Case 229
 Combo1.AddItem "  虛空要塞"
 Combo1.AddItem "  巫毒神殿"

Case 230
 Combo1.AddItem "  活動宴會廳"
 Combo1.AddItem "  學院廣場"
 
Case 231
 Combo1.AddItem "  學院競技場"

Case 232
 Combo1.AddItem "  隱藏聖門本館"

Case 233
 Combo1.AddItem "  隱藏玄嚴本館"

Case 234
 Combo1.AddItem "  隱藏鳳凰本館"

Case 235
 Combo1.AddItem "  財團附屬停車場E2"

Case 236
 Combo1.AddItem "  死牢入口"
 Combo1.AddItem "0 聖財團私立監獄"
 Combo1.AddItem "1 大廳"
 Combo1.AddItem "2 學院廣場"

Case 237
 Combo1.AddItem "  大廳"
 Combo1.AddItem "0 死牢入口"
 Combo1.AddItem "1 重刑區"
 Combo1.AddItem "2 死牢-特案區"

Case 238
 Combo1.AddItem "  重刑區"
 Combo1.AddItem "0 大廳"
 Combo1.AddItem "1 骯髒區"

Case 239
 Combo1.AddItem "  骯髒區"
 Combo1.AddItem "0 重刑區"
 Combo1.AddItem "1 活體區"

Case 240
 Combo1.AddItem "  活體區"
 Combo1.AddItem "0 骯髒區"
 Combo1.AddItem "1 生化區"

Case 241
 Combo1.AddItem "  生化區"
 Combo1.AddItem "0 活體區"
 Combo1.AddItem "1 憎惡區"

Case 242
 Combo1.AddItem "  憎惡區"
 Combo1.AddItem "0 生化區"

Case 243
 Combo1.AddItem "  死牢-特案區"
 Combo1.AddItem "  死牢-幽禁區"

Case 244
 Combo1.AddItem "  死牢-幽禁區"
 Combo1.AddItem "  死牢-特案區"

Case 245
 Combo1.AddItem "  異界虎令本館B3"
 Combo1.AddItem "0 異界虎令本館B2"
 Combo1.AddItem "1 虎令本館B3"
 Combo1.AddItem "2 本館"

Case 246
 Combo1.AddItem "  異界虎令本館B2"
 Combo1.AddItem "1 異界虎令本館B3"
 Combo1.AddItem "3 異界虎令本館B1"

Case 247
 Combo1.AddItem "  異界虎令本館B1"
 Combo1.AddItem "1 異界虎令本館B2"
 Combo1.AddItem "2 本館"
 Combo1.AddItem "3 異界虎令本館"

Case 248
 Combo1.AddItem "  異界虎令本館"
 Combo1.AddItem "1 異界虎令本館B1"
 Combo1.AddItem "3 異界虎令操場"

Case 249
 Combo1.AddItem "  異界虎令操場"
 Combo1.AddItem "1 異界虎令本館"

End Select

    Combo1.ListIndex = 0

End If
End Sub

