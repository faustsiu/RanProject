VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Attack"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5910
   LinkMode        =   1  '�ӷ�
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   5910
   StartUpPosition =   3  '�t�ιw�]��
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
 Combo1.AddItem "  �t�����]1F"
 Combo1.AddItem "6 �t���ǰ|"

Case 2
 Combo1.AddItem "  �t���ǰ|"
 Combo1.AddItem "1 �t�����]1F"
 Combo1.AddItem "2 �t���}"
 Combo1.AddItem "3 ��X�׷Ұ|"
 Combo1.AddItem "20�ǰ|�s��"

Case 100
 Combo1.AddItem "  �ǲ�1�]"

Case 101
 Combo1.AddItem "  �ǲ�2�]"

Case 103
 Combo1.AddItem "  �ǲ�3�]"

Case 104
 Combo1.AddItem "  ���v�]"

Case 105
 Combo1.AddItem "  �Ϯ��]"

Case 106
 Combo1.AddItem "  ���Ϋ�"

Case 107
 Combo1.AddItem "  ����]"

Case 102
 Combo1.AddItem "  �J��"

Case 3
 Combo1.AddItem "  �t���}"
 Combo1.AddItem "1 �t���ǰ|"
 Combo1.AddItem "2 ���ɬ}"
 Combo1.AddItem "3 ��Ĭ}"
 Combo1.AddItem "4 ��O�ǰ|"
 Combo1.AddItem "12�Ӭ}"

Case 4
 Combo1.AddItem "  ���ɥ��]1F"

Case 5
 Combo1.AddItem "  ���ɾǰ|"
 Combo1.AddItem "20�ǰ|�s��"

Case 120
 Combo1.AddItem "  �ǲ�2�]"

Case 121
 Combo1.AddItem "  �ǲ�1�]"

Case 123
 Combo1.AddItem "  ���N�]"

Case 122
 Combo1.AddItem "  �Ϯ��]"

Case 124
 Combo1.AddItem "  �ǥ�1�]"

Case 125
 Combo1.AddItem "  �ǥ�2�]"

Case 126
 Combo1.AddItem "  �J��"

Case 128
 Combo1.AddItem "  �J�٧O�]"

Case 6
 Combo1.AddItem "  ���ɬ}"
 Combo1.AddItem "2 �t���}"

Case 7
 Combo1.AddItem "  ��ĥ��]1F"

Case 8
 Combo1.AddItem "  ��ľǰ|"
 Combo1.AddItem "20�ǰ|�s��"

Case 130
 Combo1.AddItem "  �ǲ�2�]"

Case 131
 Combo1.AddItem "  �ǲ�1�]"

Case 136
 Combo1.AddItem "  ����]"

Case 135
 Combo1.AddItem "  ����]"

Case 137
 Combo1.AddItem "  �����]"

Case 138
 Combo1.AddItem "  ����]"

Case 141
 Combo1.AddItem "  ���U�����޲z�]"

Case 142
 Combo1.AddItem "  �Ϯ��]"

Case 134
 Combo1.AddItem "  �J��"

Case 9
 Combo1.AddItem "  ��Ĭ}"
 Combo1.AddItem "2 �t���}"

Case 10
 Combo1.AddItem "  ��O�ǰ|"
 Combo1.AddItem "1 ��O���]1F"
 Combo1.AddItem "2 �t���}"

Case 11
 Combo1.AddItem "  ��O���]1F"
 Combo1.AddItem "1 ��O���]2F"
 Combo1.AddItem "6 ��O�ǰ|"

Case 12
 Combo1.AddItem "  ��O���]2F"
 Combo1.AddItem "1 ��O���]1F"
 Combo1.AddItem "3 ��O���]3F"
 Combo1.AddItem "5 ��O���]B1"

Case 13
 Combo1.AddItem "  ��O���]3F"
 Combo1.AddItem "3 ��O���]2F"
 Combo1.AddItem "5 ��O���]B2"

Case 14
 Combo1.AddItem "  ��O���]B1"
 Combo1.AddItem "1 ��O���]2F"

Case 15
 Combo1.AddItem "  ��O���]B2"
 Combo1.AddItem "1 ��O���]3F"
 Combo1.AddItem "2 ��O���]B3"

Case 30
 Combo1.AddItem "  ��O���]B3"
 Combo1.AddItem "0 ���]"
 Combo1.AddItem "2 ��O���]B2"
 Combo1.AddItem "4 ���ɪ�O���]B3"

Case 17
 Combo1.AddItem "  �t�F�G�D"

Case 16
 Combo1.AddItem "  �Ӭ}"
 Combo1.AddItem "0 ���]"
 Combo1.AddItem "1 �t���}"
 Combo1.AddItem "8 ��X�X�Y"
 Combo1.AddItem "10���}"
 Combo1.AddItem "11�Ӭ}3���G�D"
 Combo1.AddItem "15��X�B�ʳ�"
 Combo1.AddItem "18�����a��"

Case 161
 Combo1.AddItem "  1�� �a�U������"

Case 162
 Combo1.AddItem "  2�� �a�U������"

Case 18
 Combo1.AddItem "  ��X�X�Y"
 Combo1.AddItem "1 �Ӭ}"
 Combo1.AddItem "3 �C��a1F_A��"
 Combo1.AddItem "5 �C��a1F_B��"
 Combo1.AddItem "6 �C��a1F_C��"

Case 19
 Combo1.AddItem "  �C��a1F_A��"
 Combo1.AddItem "0 ��X�X�Y"
 Combo1.AddItem "1 �C��a2F_A��"

Case 20
 Combo1.AddItem "  �C��a2F_A��"
 Combo1.AddItem "0 �C��a3F_A��"
 Combo1.AddItem "1 �C��a1F_A��"
 Combo1.AddItem "2 �C��a1F_A��"

Case 21
 Combo1.AddItem "  �C��a3F_A��"
 Combo1.AddItem "0 �C��a2F_A��"

Case 150
 Combo1.AddItem "  �C��a1F_B��"
 Combo1.AddItem "0 ��X�X�Y"
 Combo1.AddItem "1 �C��a2F_B��"

Case 152
 Combo1.AddItem "  �C��a2F_B��"
 Combo1.AddItem "0 �C��a3F_B��"
 Combo1.AddItem "1 �C��a1F_B��"

Case 154
 Combo1.AddItem "  �C��a3F_B��"
 Combo1.AddItem "0 �C��a2F_B��"

Case 151
 Combo1.AddItem "  �C��a1F_C��"
 Combo1.AddItem "0 ��X�X�Y"
 Combo1.AddItem "1 �C��a2F_C��"

Case 153
 Combo1.AddItem "  �C��a2F_C��"
 Combo1.AddItem "0 �C��a3F_C��"
 Combo1.AddItem "1 �C��a1F_C��"

Case 155
 Combo1.AddItem "  �C��a3F_C��"
 Combo1.AddItem "0 �C��a2F_C��"

Case 26
 Combo1.AddItem "  �Ӭ}3���G�D"
 Combo1.AddItem "0 �Ӭ}"
 Combo1.AddItem "1 �t�]�Ψp�ߺʺ�"

Case 27
 Combo1.AddItem "  �t�]�Ψp�ߺʺ�"
 Combo1.AddItem "0 �Ӭ}3���G�D"
 Combo1.AddItem "1 ���]"
 Combo1.AddItem "4 �ʺ������"
 Combo1.AddItem "5 ���c�J�f"

Case 28
 Combo1.AddItem "  ���}"
 Combo1.AddItem "0 �Ӭ}"
 Combo1.AddItem "2 ���}"

Case 29
 Combo1.AddItem "  ���}"
 Combo1.AddItem "0 ���}"
 Combo1.AddItem "1 ���]"
 Combo1.AddItem "2 �t�]��1F"

Case 22
 Combo1.AddItem "  �ǰ|�s��"
 Combo1.AddItem "1 �t���ǰ|"
 Combo1.AddItem "2 ���ɾǰ|"
 Combo1.AddItem "3 ��ľǰ|"
 Combo1.AddItem "11�t���q��ǤJ�f"
 Combo1.AddItem "12���ɹq��ǤJ�f"
 Combo1.AddItem "13��Ĺq��ǤJ�f"
 Combo1.AddItem "14�Ӭ}�q��ǤJ�f"
 Combo1.AddItem "  �]�Ϊ��ݰ�����B"
 Combo1.AddItem "  �]�Ϊ��ݰ�����C"
 Combo1.AddItem "  �]�Ϊ��ݰ�����D"
 Combo1.AddItem "95�]�Ϊ��ݰ�����E"
 Combo1.AddItem "5 �𮧷|�]"
 Combo1.AddItem "  ���B�b�|�U"
 Combo1.AddItem "  ���ʮb�|�U"

Case 23
 Combo1.AddItem "  ��X�׷Ұ|"
 Combo1.AddItem "2 �t���ǰ|"
 Combo1.AddItem "  ���ɾǰ|"
 Combo1.AddItem "  ��ľǰ|"

Case 31
 Combo1.AddItem "  ���B�b�|�U"
 Combo1.AddItem "0 �ǰ|�s��"

Case 201
 Combo1.AddItem "  �t���q��ǤJ�f"
 Combo1.AddItem "0 �ǰ|�s��"
 Combo1.AddItem "1 �t���q���"

Case 211
 Combo1.AddItem "  �t���}�q���"

Case 202
 Combo1.AddItem "  ���ɹq��ǤJ�f"
 Combo1.AddItem "0 �ǰ|�s��"
 Combo1.AddItem "1 ���ɬ}�q���"

Case 212
 Combo1.AddItem "  ���ɬ}�q���"

Case 203
 Combo1.AddItem "  ��Ĺq��ǤJ�f"
 Combo1.AddItem "0 �ǰ|�s��"
 Combo1.AddItem "1 ��Ĺq���"

Case 213
 Combo1.AddItem "  ��Ĭ}�q���"

Case 204
 Combo1.AddItem "  �Ӭ}�q��ǤJ�f"
 Combo1.AddItem "0 �ǰ|�s��"

Case 214
 Combo1.AddItem "  �Ӭ}�q���"

Case 32
 Combo1.AddItem "  �ʺ������"
 Combo1.AddItem "0 �t�]�Ψp�ߺʺ�"
 Combo1.AddItem "1 ���]"
 Combo1.AddItem "2 �ĤC�����"

Case 33
 Combo1.AddItem "  �ĤC�����"
 Combo1.AddItem "0 ���]"
 Combo1.AddItem "  �ʺ������"

Case 34
 Combo1.AddItem "  �A��"

Case 46
 Combo1.AddItem "  �t�]��1F"
 Combo1.AddItem "0 ���}"
 Combo1.AddItem "1 ���]"
 Combo1.AddItem "5 �t�]�Φa�U��"
 Combo1.AddItem "2 �t�]��2F"
 Combo1.AddItem "3 �t�]��30F"

Case 47
 Combo1.AddItem "  �t�]��2F"
 Combo1.AddItem "2 �t�]��1F"
 Combo1.AddItem "4 �t�]��3F"

Case 48
 Combo1.AddItem "  �t�]��3F"
 Combo1.AddItem "1 �t�]��2F"

Case 35
 Combo1.AddItem "  �t�]��30F"
 Combo1.AddItem "1 ���]"
 Combo1.AddItem "2 �t�]��50F"
 Combo1.AddItem "3 �t�]��1F"

Case 36
 Combo1.AddItem "  �t�]��50F"
 Combo1.AddItem "1 ���]"
 Combo1.AddItem "  �t�]��30F"
 Combo1.AddItem "4 �t�]��80F"
 Combo1.AddItem "  �t�]��90F"

Case 37
 Combo1.AddItem "  �t�]��90F"
 Combo1.AddItem "0 �t�]��50F"
 Combo1.AddItem "1 ���]"
 Combo1.AddItem "  �t�]��80��"

Case 97
 Combo1.AddItem "  �]�Ϊ��ݰ�����B"

Case 98
 Combo1.AddItem "  �]�Ϊ��ݰ�����C"

Case 99
 Combo1.AddItem "  �]�Ϊ��ݰ�����D"

Case 95
 Combo1.AddItem "  �]�Ϊ��ݰ�����E1"

Case 38
 Combo1.AddItem "  �t�]�Τj��_���l"
 Combo1.AddItem "0 �t�]��80��"
 Combo1.AddItem "1 ���]"
 Combo1.AddItem "2 �t�]�αK�ǥ���"

Case 39
 Combo1.AddItem "  �t�]�Τj��_�k�l"
 Combo1.AddItem "0 �t�]��80��"
 Combo1.AddItem "2 �t�]�αK�ǥk��"

Case 40
 Combo1.AddItem "  �t�]�αK�ǥ���"
 Combo1.AddItem "  �t�]��80��"
 Combo1.AddItem "  �t�]�Τj��_���l"

Case 41
 Combo1.AddItem "  �t�]�αK�ǥk��"
 Combo1.AddItem "  �t�]�Τj��_�k�l"
 Combo1.AddItem "  ���ɷ���1F"

Case 42
 Combo1.AddItem "  �t�]�Φa�U��"
 Combo1.AddItem "0 �t�]��1F"

Case 43
 Combo1.AddItem "  ���ɷ���1F"

Case 44
 Combo1.AddItem "  ���ɷ���2F"

Case 45
 Combo1.AddItem "  ���ɯ���"

Case 50
 Combo1.AddItem "  ����"

Case 215
 Combo1.AddItem "  �𶢷|�]"
 Combo1.AddItem "0 �ǰ|�s��"

Case 220
 Combo1.AddItem "  ��X�B�ʳ�"
 Combo1.AddItem "0 �Ӭ}"

Case 51
 Combo1.AddItem "  �t�]��80��"
 Combo1.AddItem "0 �t�]��90F"
 Combo1.AddItem "1 ���]"
 Combo1.AddItem "2 �t�]��1F"
 Combo1.AddItem "3 �t�]��50F"
 Combo1.AddItem "5 �t�]�Τj��_���l"
 Combo1.AddItem "7 �t�]�Τj��_�k�l"

Case 216
 Combo1.AddItem "  �t���q��K��"

Case 217
 Combo1.AddItem "  ���Y�q��K��"

Case 218
 Combo1.AddItem "  ��Ĺq��K��"

Case 219
 Combo1.AddItem "  �Ӭ}�q��K��"

Case 225
 Combo1.AddItem "  �����a��"
 Combo1.AddItem "0 �Ӭ}"
 Combo1.AddItem "1 ���]"
 Combo1.AddItem "2 ������"
 Combo1.AddItem "4 �p�q�۹�"
 Combo1.AddItem "6 �Ŭr����"

Case 226
 Combo1.AddItem "  ������"
 Combo1.AddItem "  �����a��"

Case 227
 Combo1.AddItem "  �p�q�۹�"
 Combo1.AddItem "  �����a��"

Case 228
 Combo1.AddItem "  �Ŭr����"
 Combo1.AddItem "  �����a��"
 Combo1.AddItem "3 ��ŭn��"

Case 229
 Combo1.AddItem "  ��ŭn��"
 Combo1.AddItem "  �Ŭr����"

Case 230
 Combo1.AddItem "  ���ʮb�|�U"
 Combo1.AddItem "  �ǰ|�s��"
 
Case 231
 Combo1.AddItem "  �ǰ|�v�޳�"

Case 232
 Combo1.AddItem "  ���øt�����]"

Case 233
 Combo1.AddItem "  ���å��Y���]"

Case 234
 Combo1.AddItem "  ���û�ĥ��]"

Case 235
 Combo1.AddItem "  �]�Ϊ��ݰ�����E2"

Case 236
 Combo1.AddItem "  ���c�J�f"
 Combo1.AddItem "0 �t�]�Ψp�ߺʺ�"
 Combo1.AddItem "1 �j�U"
 Combo1.AddItem "2 �ǰ|�s��"

Case 237
 Combo1.AddItem "  �j�U"
 Combo1.AddItem "0 ���c�J�f"
 Combo1.AddItem "1 ���D��"
 Combo1.AddItem "2 ���c-�S�װ�"

Case 238
 Combo1.AddItem "  ���D��"
 Combo1.AddItem "0 �j�U"
 Combo1.AddItem "1 ��ż��"

Case 239
 Combo1.AddItem "  ��ż��"
 Combo1.AddItem "0 ���D��"
 Combo1.AddItem "1 �����"

Case 240
 Combo1.AddItem "  �����"
 Combo1.AddItem "0 ��ż��"
 Combo1.AddItem "1 �ͤư�"

Case 241
 Combo1.AddItem "  �ͤư�"
 Combo1.AddItem "0 �����"
 Combo1.AddItem "1 ���c��"

Case 242
 Combo1.AddItem "  ���c��"
 Combo1.AddItem "0 �ͤư�"

Case 243
 Combo1.AddItem "  ���c-�S�װ�"
 Combo1.AddItem "  ���c-�ոT��"

Case 244
 Combo1.AddItem "  ���c-�ոT��"
 Combo1.AddItem "  ���c-�S�װ�"

Case 245
 Combo1.AddItem "  ���ɪ�O���]B3"
 Combo1.AddItem "0 ���ɪ�O���]B2"
 Combo1.AddItem "1 ��O���]B3"
 Combo1.AddItem "2 ���]"

Case 246
 Combo1.AddItem "  ���ɪ�O���]B2"
 Combo1.AddItem "1 ���ɪ�O���]B3"
 Combo1.AddItem "3 ���ɪ�O���]B1"

Case 247
 Combo1.AddItem "  ���ɪ�O���]B1"
 Combo1.AddItem "1 ���ɪ�O���]B2"
 Combo1.AddItem "2 ���]"
 Combo1.AddItem "3 ���ɪ�O���]"

Case 248
 Combo1.AddItem "  ���ɪ�O���]"
 Combo1.AddItem "1 ���ɪ�O���]B1"
 Combo1.AddItem "3 ���ɪ�O�޳�"

Case 249
 Combo1.AddItem "  ���ɪ�O�޳�"
 Combo1.AddItem "1 ���ɪ�O���]"

End Select

    Combo1.ListIndex = 0

End If
End Sub

