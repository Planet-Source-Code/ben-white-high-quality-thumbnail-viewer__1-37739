VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Thumbnail Viewer"
   ClientHeight    =   1455
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2775
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1455
   ScaleWidth      =   2775
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cboQuality 
      Height          =   315
      ItemData        =   "main.frx":0000
      Left            =   1440
      List            =   "main.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   360
      Width           =   1215
   End
   Begin VB.PictureBox picHigh 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      ForeColor       =   &H80000008&
      Height          =   1215
      Left            =   120
      ScaleHeight     =   1185
      ScaleWidth      =   1185
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Load Image"
      Height          =   315
      Left            =   1440
      TabIndex        =   1
      Top             =   1020
      Width           =   1215
   End
   Begin VB.PictureBox picTemp 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   2400
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label labSpeed 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   1440
      TabIndex        =   5
      Top             =   750
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Quality"
      Height          =   255
      Left            =   1440
      TabIndex        =   4
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim bolLoaded As Boolean

Private Declare Function GetTickCount Lib "kernel32" () As Long

Private Sub cboQuality_Click()
    Dim strFilename, tick

    If Not bolLoaded Then Exit Sub
    
    cboQuality.Refresh
    
    'Start timer
    tick = GetTickCount()
    labSpeed.Caption = "Loading..."
    'Get Thumbnail
    ViewImage "", picTemp, picHigh, cboQuality.ListIndex
    'display operation time
    labSpeed.Caption = GetTickCount() - tick & "ms"
End Sub

Private Sub cmdLoad_Click()

    Dim strFilename, tick
    
    'Get filename
    strFilename = ShowOpen(Me.hwnd)
    
    If strFilename <> "" Then
        'Start timer
        tick = GetTickCount()
        labSpeed.Caption = "Loading..."
        'Get Thumbnail
        ViewImage strFilename, picTemp, picHigh, cboQuality.ListIndex
        'display operation time
        labSpeed.Caption = GetTickCount() - tick & "ms"
        bolLoaded = True
    End If
    
End Sub

Private Sub Form_Load()

    cboQuality.AddItem "Poor", 0
    cboQuality.AddItem "Low", 1
    cboQuality.AddItem "Average", 2
    cboQuality.AddItem "Fair", 3
    cboQuality.AddItem "Best", 4
    
    cboQuality.ListIndex = 0
End Sub
