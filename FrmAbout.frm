VERSION 5.00
Begin VB.Form FrmAbout 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "����"
   ClientHeight    =   3030
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4560
   BeginProperty Font 
      Name            =   "΢���ź�"
      Size            =   9
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  '����ȱʡ
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "ע�⣺��ǰ�汾Ϊ�����汾��ʹ�øð汾����������һ�к����ʹ���߸������߲��е��κ����Σ�"
      Height          =   615
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   4215
      WordWrap        =   -1  'True
   End
   Begin VB.Label LblAuthor 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "���ߣ������"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   1080
   End
   Begin VB.Label LblVer 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Alpha Version"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Ceasium Meeting Recorder"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2385
   End
End
Attribute VB_Name = "FrmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
LblVer.Caption = Versionstring
Me.Show
End Sub
