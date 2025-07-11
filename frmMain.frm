VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   "MQTT Client"
   ClientHeight    =   11475
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   16095
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11475
   ScaleWidth      =   16095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.Frame Frame4 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Caption         =   "Frame4"
      BeginProperty Font 
         Name            =   "等线"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   8040
      TabIndex        =   40
      Top             =   1080
      Width           =   7575
      Begin VB.TextBox txtState 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "等线"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   7080
         TabIndex        =   41
         Text            =   "0"
         Top             =   120
         Width           =   375
      End
      Begin VB.Label Label23 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Log"
         BeginProperty Font 
            Name            =   "等线"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   44
         Top             =   120
         Width           =   1215
      End
   End
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      BeginProperty Font 
         Name            =   "等线"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   2775
      Left            =   480
      TabIndex        =   31
      Top             =   8280
      Width           =   7335
      Begin VB.TextBox txtTopicGET 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "等线"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1800
         TabIndex        =   33
         Text            =   "AAA"
         Top             =   360
         Width           =   3255
      End
      Begin VB.TextBox txtRx 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   1455
         Left            =   1800
         MultiLine       =   -1  'True
         TabIndex        =   32
         Top             =   960
         Width           =   5055
      End
      Begin VB.Label Label21 
         Height          =   495
         Left            =   1680
         TabIndex        =   38
         Top             =   240
         Width           =   3495
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Topic"
         BeginProperty Font 
            Name            =   "等线"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   37
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Messages"
         BeginProperty Font 
            Name            =   "等线"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   36
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label cmdGET 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0059911C&
         BackStyle       =   0  'Transparent
         Caption         =   "Subscribe"
         BeginProperty Font 
            Name            =   "等线"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   5400
         TabIndex        =   35
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label17 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   1695
         Left            =   1680
         TabIndex        =   34
         Top             =   840
         Width           =   5295
      End
      Begin VB.Label border_cmdGET 
         BackColor       =   &H0059911C&
         Height          =   495
         Left            =   5280
         TabIndex        =   39
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      BeginProperty Font 
         Name            =   "等线"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   3375
      Left            =   480
      TabIndex        =   21
      Top             =   4680
      Width           =   7335
      Begin VB.TextBox txtMsg 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   1455
         Left            =   1800
         MultiLine       =   -1  'True
         TabIndex        =   26
         Top             =   960
         Width           =   5055
      End
      Begin VB.TextBox txtTopic 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "等线"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1800
         TabIndex        =   23
         Text            =   "AAA"
         Top             =   360
         Width           =   5055
      End
      Begin VB.Label cmdPing 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Ping"
         BeginProperty Font 
            Name            =   "等线"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4560
         TabIndex        =   43
         Top             =   2760
         Width           =   975
      End
      Begin VB.Label border_cmdPing 
         BackColor       =   &H00C0C0C0&
         Height          =   495
         Left            =   4440
         TabIndex        =   42
         Top             =   2640
         Width           =   1215
      End
      Begin VB.Label Label15 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   1695
         Left            =   1680
         TabIndex        =   30
         Top             =   840
         Width           =   5295
      End
      Begin VB.Label cmdPUT 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0059911C&
         BackStyle       =   0  'Transparent
         Caption         =   "Send"
         BeginProperty Font 
            Name            =   "等线"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   5880
         TabIndex        =   28
         Top             =   2760
         Width           =   975
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Messages"
         BeginProperty Font 
            Name            =   "等线"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   27
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Topic"
         BeginProperty Font 
            Name            =   "等线"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   25
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label11 
         Height          =   495
         Left            =   1680
         TabIndex        =   24
         Top             =   240
         Width           =   5295
      End
      Begin VB.Label border_cmdPUT 
         BackColor       =   &H0059911C&
         Height          =   495
         Left            =   5760
         TabIndex        =   29
         Top             =   2640
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Caption         =   "Port:"
      BeginProperty Font 
         Name            =   "等线"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3375
      Left            =   480
      TabIndex        =   2
      Top             =   1080
      Width           =   7335
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "等线"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1800
         TabIndex        =   10
         Text            =   "60"
         Top             =   2160
         Width           =   5055
      End
      Begin VB.TextBox txtIdentifier 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "等线"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1800
         TabIndex        =   8
         Text            =   "mqttid0000"
         Top             =   1560
         Width           =   5055
      End
      Begin VB.TextBox Port 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "等线"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1800
         TabIndex        =   6
         Text            =   "1883"
         Top             =   960
         Width           =   5055
      End
      Begin VB.TextBox cmbBroker 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "等线"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1800
         TabIndex        =   3
         Text            =   "127.0.0.1"
         Top             =   360
         Width           =   5055
      End
      Begin VB.Label txtMQTTState 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "State"
         BeginProperty Font 
            Name            =   "等线"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000F&
         Height          =   255
         Left            =   1680
         TabIndex        =   22
         Top             =   2760
         Width           =   1695
      End
      Begin VB.Label cmdDisconnect 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Disconnect"
         BeginProperty Font 
            Name            =   "等线"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3840
         TabIndex        =   19
         Top             =   2760
         Width           =   1335
      End
      Begin VB.Label cmdConnect 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0059911C&
         BackStyle       =   0  'Transparent
         Caption         =   "Connect"
         BeginProperty Font 
            Name            =   "等线"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   5640
         TabIndex        =   17
         Top             =   2760
         Width           =   1215
      End
      Begin VB.Label Label8 
         Height          =   495
         Left            =   1680
         TabIndex        =   15
         Top             =   2040
         Width           =   5295
      End
      Begin VB.Label Label7 
         Height          =   495
         Left            =   1680
         TabIndex        =   14
         Top             =   1440
         Width           =   5295
      End
      Begin VB.Label Label6 
         Height          =   495
         Left            =   1680
         TabIndex        =   13
         Top             =   840
         Width           =   5295
      End
      Begin VB.Label Label5 
         Height          =   495
         Left            =   1680
         TabIndex        =   12
         Top             =   240
         Width           =   5295
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Keep Alive"
         BeginProperty Font 
            Name            =   "等线"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   2160
         Width           =   1215
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Client ID"
         BeginProperty Font 
            Name            =   "等线"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Port"
         BeginProperty Font 
            Name            =   "等线"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Host"
         BeginProperty Font 
            Name            =   "等线"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label border_cmdConnect 
         BackColor       =   &H0059911C&
         Height          =   495
         Left            =   5520
         TabIndex        =   18
         Top             =   2640
         Width           =   1455
      End
      Begin VB.Label border_cmdDisconnect 
         BackColor       =   &H00C0C0C0&
         Height          =   495
         Left            =   3600
         TabIndex        =   20
         Top             =   2640
         Width           =   1815
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   5880
      Top             =   360
   End
   Begin VB.Timer Timer2 
      Interval        =   50000
      Left            =   6360
      Top             =   360
   End
   Begin VB.TextBox log 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   9255
      Left            =   8160
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   1680
      Width           =   7455
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   5400
      Top             =   360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemoteHost      =   "127.0.0.1"
      RemotePort      =   1883
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000000C0&
      BorderWidth     =   2
      X1              =   15480
      X2              =   15840
      Y1              =   600
      Y2              =   240
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000C0&
      BorderWidth     =   2
      X1              =   15480
      X2              =   15840
      Y1              =   240
      Y2              =   600
   End
   Begin VB.Label exit 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "等线"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   15360
      TabIndex        =   5
      Top             =   120
      Width           =   615
   End
   Begin VB.Label title 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Visual Basic 6.0 - MQTT 3.1.1"
      BeginProperty Font 
         Name            =   "等线"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0059911C&
      Height          =   375
      Left            =   480
      TabIndex        =   1
      Top             =   480
      Width           =   4515
   End
   Begin VB.Label Label9 
      Height          =   9495
      Left            =   8040
      TabIndex        =   16
      Top             =   1560
      Width           =   7575
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' udk
' Linux_index@outlook.com

' 顶部声明
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
    (Destination As Any, Source As Any, ByVal length As Long)

Private Declare Function MultiByteToWideChar Lib "kernel32" ( _
    ByVal CodePage As Long, _
    ByVal dwFlags As Long, _
    ByVal lpMultiByteStr As Long, _
    ByVal cbMultiByte As Long, _
    ByVal lpWideCharStr As Long, _
    ByVal cchWideChar As Long) As Long

Private Const CP_UTF8 As Long = 65001

' 声明常量
Private Const MQTT_CONNECT As Integer = &H10
Private Const MQTT_CONNACK As Integer = &H20
Private Const MQTT_PUBLISH As Integer = &H30
Private Const MQTT_SUBSCRIBE As Integer = &H82
Private Const MQTT_SUBACK As Integer = &H90
Private Const MQTT_UNSUBSCRIBE As Integer = &HA2
Private Const MQTT_PINGREQ As Integer = &HC0
Private Const MQTT_PINGRESP As Integer = &HD0
Private Const MQTT_DISCONNECT As Integer = &HE0
Private Const SOCKET_CONNECTED As Integer = 7

' 全局变量
Dim m_Connected As Boolean


' 使用Windows API的UTF-8转换函数 (更可靠)
Private Function UTF8ToStr(b() As Byte) As String
    Dim length As Long
    Dim result As String
    
    ' 获取所需缓冲区长度
    length = MultiByteToWideChar(CP_UTF8, 0, VarPtr(b(LBound(b))), UBound(b) - LBound(b) + 1, 0, 0)
    
    If length > 0 Then
        result = String$(length, 0)
        MultiByteToWideChar CP_UTF8, 0, VarPtr(b(LBound(b))), UBound(b) - LBound(b) + 1, StrPtr(result), length
        UTF8ToStr = result
    Else
        UTF8ToStr = ""
    End If
End Function


Private Function StringToUTF8(ByVal str As String) As Byte()
    With CreateObject("ADODB.Stream")
        .Type = 2  ' adTypeText
        .Charset = "utf-8"
        .Open
        .WriteText str
        .Position = 0
        .Type = 1  ' adTypeBinary
        .Position = 3  ' 跳过BOM
        StringToUTF8 = .Read
        .Close
    End With
End Function

' 字节数组十六进制显示函数
Private Function HexDumpBytes(data() As Byte) As String
    Dim i As Long
    Dim hexStr As String
    Dim asciiStr As String
    Dim result As String
    Dim lineCount As Long
    
    For i = LBound(data) To UBound(data)
        hexStr = hexStr & Right("0" & Hex(data(i)), 2) & " "
        
        If data(i) >= 32 And data(i) <= 126 Then
            asciiStr = asciiStr & Chr(data(i))
        Else
            asciiStr = asciiStr & "."
        End If
        
        If (i + 1) Mod 16 = 0 Then
            result = result & hexStr & "  " & asciiStr & vbCrLf
            hexStr = ""
            asciiStr = ""
            lineCount = lineCount + 1
        End If
    Next i
    
    If hexStr <> "" Then
        ' 确保对齐
        While Len(hexStr) < 48
            hexStr = hexStr & "   "
        Wend
        result = result & hexStr & "  " & asciiStr
    End If
    
    HexDumpBytes = result
End Function

Private Sub border_cmdConnect_Click()
cmdConnect_Click
End Sub

Private Sub border_cmdDisconnect_Click()
cmdDisconnect_Click
End Sub

Private Sub border_cmdGET_Click()
cmdGET_Click
End Sub

Private Sub border_cmdPing_Click()
cmdPing_Click
End Sub

Private Sub border_cmdPUT_Click()
cmdPUT_Click
End Sub

Private Sub exit_Click()
End
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Static FO(2): If Button <> 0 Then Me.Move Me.Left - FO(0) + x, Me.Top - FO(1) + y Else: FO(0) = x: FO(1) = y
End Sub

Private Sub Form_Load()
    txtIdentifier = "mqttid" & Format(Time, "mm") & Trim(Format(Time, "ss"))
    m_Connected = False
    log.Text = "Application started" & vbCrLf
    DoEvents
End Sub

' 关闭连接
Private Sub cmdClose()
    On Error Resume Next
    log.Text = log.Text & "Closing connection..." & vbCrLf
    Winsock1.Close
    m_Connected = False
    While Winsock1.State <> 0
        DoEvents
    Wend
    log.Text = log.Text & "Connection closed" & vbCrLf
End Sub

' 打开连接
Private Sub cmdOpen()
    On Error Resume Next
    log.Text = log.Text & "Opening TCP connection to " & cmbBroker.Text & vbCrLf
    cmdClose
    Winsock1.LocalPort = 0
    Winsock1.RemotePort = Port.Text ' 设置远程端口
    Winsock1.RemoteHost = cmbBroker.Text
    Winsock1.Connect
End Sub

' 断开MQTT连接
Private Sub cmdDisconnect_Click()
    If Winsock1.State = SOCKET_CONNECTED And m_Connected Then
        ' 构建DISCONNECT报文: [0xE0, 0x00]
        Dim discPacket(1) As Byte
        discPacket(0) = MQTT_DISCONNECT
        discPacket(1) = &H0
        
        log.Text = log.Text & "Sending MQTT DISCONNECT" & vbCrLf
        Winsock1.SendData discPacket
    End If
    
    Winsock1.Close
    m_Connected = False
End Sub

' 建立MQTT连接
Private Sub cmdConnect_Click()
    Dim MsgLen As Integer
    Dim IDLen As Integer
    Dim Identifier As String
    Dim i As Long  ' 声明循环变量
    
    Identifier = txtIdentifier
    IDLen = Len(Identifier)
    ' 计算剩余长度（不包括固定头）
    MsgLen = 10 + 2 + IDLen  ' 协议名(6) + 协议级别(1) + 标志(1) + KeepAlive(2) + 客户端ID长度(2) + ID
    
    ' 正确计算数组大小：固定头2字节 + 剩余长度部分
    Dim connectPacket() As Byte
    ReDim connectPacket(0 To MsgLen + 1)  ' 索引0到MsgLen+1，共MsgLen+2个元素
    
    ' 固定头: 类型(0x10) + 剩余长度(MsgLen)
    connectPacket(0) = MQTT_CONNECT  ' &H10
    connectPacket(1) = MsgLen
    
    ' 可变头部和负载
    Dim pos As Long: pos = 2  ' 从索引2开始
    
    ' 协议名: 4字节 "MQTT" (前面有2字节长度)
    connectPacket(pos) = 0: pos = pos + 1
    connectPacket(pos) = 4: pos = pos + 1
    connectPacket(pos) = Asc("M"): pos = pos + 1
    connectPacket(pos) = Asc("Q"): pos = pos + 1
    connectPacket(pos) = Asc("T"): pos = pos + 1
    connectPacket(pos) = Asc("T"): pos = pos + 1
    
    ' 协议级别: 4 (MQTT 3.1.1)
    connectPacket(pos) = 4: pos = pos + 1
    
    ' 连接标志: 这里使用清除会话 &H02
    connectPacket(pos) = &H2: pos = pos + 1
    
    ' 保持时间: 60秒 (0x003C)
    connectPacket(pos) = &H0: pos = pos + 1   ' 高字节
    connectPacket(pos) = &H3C: pos = pos + 1  ' 低字节
    
    ' 客户端ID长度 (高字节0，低字节为IDLen)
    connectPacket(pos) = 0: pos = pos + 1
    connectPacket(pos) = IDLen: pos = pos + 1
    
    ' 客户端ID
    Dim idBytes() As Byte
    idBytes = StrConv(Identifier, vbFromUnicode)
    For i = 0 To IDLen - 1
        connectPacket(pos) = idBytes(i): pos = pos + 1
    Next i
    
    ' 记录日志
    Dim packetStr As String
    packetStr = StrConv(connectPacket, vbUnicode)
    log.Text = log.Text & "Sending CONNECT: " & HexDump(Left(packetStr, MsgLen + 2)) & vbCrLf
    
    ' 连接并发送
    cmdOpen
    While Winsock1.State <> SOCKET_CONNECTED
        DoEvents
    Wend
    
    Winsock1.SendData connectPacket
End Sub

' 发送PING请求
Private Sub cmdPing_Click()
    If Winsock1.State = SOCKET_CONNECTED And m_Connected Then
        ' 构建字节数组: [0xC0, 0x00]
        Dim pingReq(1) As Byte
        pingReq(0) = MQTT_PINGREQ  ' &HC0
        pingReq(1) = &H0
        
        log.Text = log.Text & "Sending PINGREQ" & vbCrLf
        Winsock1.SendData pingReq
    End If
End Sub

' 发布消息
Private Sub cmdPUT_Click()
    If Winsock1.State <> SOCKET_CONNECTED Or Not m_Connected Then Exit Sub
    If Len(txtMsg) < 1 Then Exit Sub
    
    ' 转换为UTF-8字节数组
    Dim topicBytes() As Byte
    Dim msgBytes() As Byte
    topicBytes = StringToUTF8(txtTopic)
    msgBytes = StringToUTF8(txtMsg)
    
    Dim TopicLength As Long
    Dim MsgLength As Long
    TopicLength = UBound(topicBytes) + 1
    MsgLength = UBound(msgBytes) + 1
    
    Dim remainingLength As Long
    remainingLength = 2 + TopicLength + MsgLength
    
    ' 构建报文
    Dim publishPacket() As Byte
    ReDim publishPacket(0 To remainingLength + 1)
    
    publishPacket(0) = MQTT_PUBLISH
    publishPacket(1) = remainingLength
    
    ' 主题长度(高位在前)
    publishPacket(2) = TopicLength \ 256
    publishPacket(3) = TopicLength Mod 256
    
    ' 复制主题
    Dim i As Long  ' 声明循环变量
    For i = 0 To TopicLength - 1
        publishPacket(4 + i) = topicBytes(i)
    Next i
    
    ' 复制消息
    For i = 0 To MsgLength - 1
        publishPacket(4 + TopicLength + i) = msgBytes(i)
    Next i
    
    Winsock1.SendData publishPacket
    log.Text = log.Text & "PUBLISH to [" & txtTopic & "]: " & txtMsg & vbCrLf
End Sub

' 订阅主题
Private Sub cmdGET_Click()
    Dim remainingLength As Long
    Dim TopicLength As Long
    Dim i As Long  ' 声明循环变量
    Static msgID As Integer
    
    If Not m_Connected Then
        log.Text = log.Text & "ERROR: Not connected to MQTT broker" & vbCrLf
        MsgBox "Please establish MQTT connection first"
        Exit Sub
    End If
    
    ' 转换为UTF-8
    Dim topicBytes() As Byte
    topicBytes = StringToUTF8(txtTopicGET)
    TopicLength = UBound(topicBytes) + 1
    
    remainingLength = 2 + 2 + TopicLength + 1  ' 消息ID(2) + 主题长度(2) + 主题 + QoS(1)
    
    msgID = (msgID + 1) Mod 65536
    If msgID = 0 Then msgID = 1
    
    ' 使用字节数组构建报文
    Dim byteArr() As Byte
    ReDim byteArr(0 To remainingLength + 1)  ' 固定头2字节 + 剩余长度
    
    ' 固定头: 类型(0x82) + 剩余长度
    byteArr(0) = MQTT_SUBSCRIBE  ' &H82
    byteArr(1) = remainingLength
    
    ' 可变头: 消息ID (高字节在前)
    byteArr(2) = msgID \ 256
    byteArr(3) = msgID Mod 256
    
    ' 主题长度 (高字节在前)
    byteArr(4) = TopicLength \ 256
    byteArr(5) = TopicLength Mod 256
    
    ' 主题内容 (UTF-8字节)
    For i = 0 To TopicLength - 1
        byteArr(6 + i) = topicBytes(i)
    Next i
    
    ' QoS 级别
    byteArr(6 + TopicLength) = 0  ' QoS 0
    
    ' 记录日志
    Dim logMsg As String
    logMsg = "Subscribing to: " & txtTopic & " (ID:" & msgID & ")" & vbCrLf
    log.Text = log.Text & logMsg
    
    ' 发送字节数组
    Winsock1.SendData byteArr
    log.Text = log.Text & "SUBSCRIBE sent (" & UBound(byteArr) + 1 & " bytes)" & vbCrLf
End Sub

' 接收数据处理
Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
    Dim byteData() As Byte
    ReDim byteData(0 To bytesTotal - 1)
    Winsock1.GetData byteData, vbArray + vbByte
    
    ' 记录原始字节
    log.Text = log.Text & "Received " & bytesTotal & "B: " & HexDumpBytes(byteData) & vbCrLf
    
    
    Dim MsgType As Byte
    MsgType = byteData(0) And &HF0  ' 取高4位
    
    ' 处理PUBLISH
    If MsgType = MQTT_PUBLISH Then
        Dim pos As Long: pos = 1
        Dim remainingLength As Long
        Dim multiplier As Long: multiplier = 1
        Dim digit As Byte
        
        ' 解析剩余长度 (正确计算长度字节数)
        Do
            digit = byteData(pos)
            remainingLength = remainingLength + (digit And 127) * multiplier
            multiplier = multiplier * 128
            pos = pos + 1
        Loop While (digit And 128) <> 0
        
        ' 读取主题长度 (高位在前)
        Dim TopicLen As Integer
        TopicLen = byteData(pos) * 256 + byteData(pos + 1)
        pos = pos + 2
        
        ' 提取主题(UTF-8) - 确保只复制主题部分
        Dim topicBytes() As Byte
        ReDim topicBytes(0 To TopicLen - 1)
        CopyMemory topicBytes(0), byteData(pos), TopicLen
        pos = pos + TopicLen
        
        ' 提取消息内容(UTF-8) - 精确计算长度
        Dim payloadBytes() As Byte
        Dim payloadLen As Long
        payloadLen = remainingLength - 2 - TopicLen
        If payloadLen > 0 Then
            ReDim payloadBytes(0 To payloadLen - 1)
            CopyMemory payloadBytes(0), byteData(pos), payloadLen
        Else
            ReDim payloadBytes(0)
        End If
        
        ' 使用API转换UTF-8
        Dim topic As String
        Dim payload As String
        topic = UTF8ToStr(topicBytes)
        payload = UTF8ToStr(payloadBytes)
        
        log.Text = log.Text & "Topic: " & topic & ", Message: " & payload & vbCrLf
        txtRx = topic & vbCrLf & vbCrLf & payload
    End If
    
    ' 处理CONNACK
    If MsgType = MQTT_CONNACK Then
        If UBound(byteData) >= 3 Then
            Dim connCode As Integer
            connCode = byteData(3)
            If connCode = 0 Then
                m_Connected = True
                log.Text = log.Text & "MQTT connection established" & vbCrLf
            Else
                m_Connected = False
                log.Text = log.Text & "Connection refused: code " & connCode & vbCrLf
            End If
        End If
    End If
    
    ' 处理SUBACK
    If MsgType = MQTT_SUBACK Then
        log.Text = log.Text & "SUBACK received" & vbCrLf
        If UBound(byteData) >= 4 Then
            Dim returnCode As Integer
            Dim packetID As Integer
            packetID = byteData(2) * 256 + byteData(3)
            returnCode = byteData(4)
            
            log.Text = log.Text & "Packet ID: " & packetID & ", Return code: " & Hex(returnCode) & vbCrLf
            
            If returnCode < &H80 Then
                log.Text = log.Text & "Subscribe successful (QoS " & returnCode & ")" & vbCrLf
            Else
                log.Text = log.Text & "Subscribe failed: " & Hex(returnCode) & vbCrLf
            End If
        End If
    End If
    
    ' 处理PING响应
    If MsgType = MQTT_PINGRESP Then
        log.Text = log.Text & "PINGRESP received" & vbCrLf
    End If
End Sub

' 十六进制转储函数
Private Function HexDump(data As String) As String
    Dim i As Integer
    Dim result As String
    Dim hexPart As String
    Dim asciiPart As String
    
    For i = 1 To Len(data)
        Dim char As String
        char = Mid(data, i, 1)
        hexPart = hexPart & Right("0" & Hex(Asc(char)), 2) & " "
        
        If Asc(char) >= 32 And Asc(char) <= 126 Then
            asciiPart = asciiPart & char
        Else
            asciiPart = asciiPart & "."
        End If
        
        If i Mod 16 = 0 Then
            result = result & hexPart & "  " & asciiPart & vbCrLf
            hexPart = ""
            asciiPart = ""
        End If
    Next i
    
    If Len(hexPart) > 0 Then
        While Len(hexPart) < 48
            hexPart = hexPart & "   "
        Wend
        result = result & hexPart & "  " & asciiPart
    End If
    
    HexDump = result
End Function

' 定时更新状态
Private Sub Timer1_Timer()
    txtState = Winsock1.State
    If m_Connected And txtState = "7" Then
        txtMQTTState.Caption = "Connected"
    Else
        txtMQTTState.Caption = "Disconnected"
        m_Connected = False
    End If
End Sub

' 定时发送PING
Private Sub Timer2_Timer()
    If m_Connected Then
        cmdPing_Click
    End If
End Sub

Private Sub log_Change()
log.SelStart = Len(log.Text)
End Sub

