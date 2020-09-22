VERSION 5.00
Object = "*\AVectorPreviewCtl.vbp"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5625
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   8145
   LinkTopic       =   "Form1"
   ScaleHeight     =   5625
   ScaleWidth      =   8145
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Draw Mode"
      Height          =   1245
      Left            =   6300
      TabIndex        =   6
      Top             =   2100
      Width           =   1665
      Begin VB.OptionButton OpDraw 
         Alignment       =   1  'Right Justify
         Caption         =   "Filled"
         Height          =   255
         Index           =   2
         Left            =   90
         TabIndex        =   9
         Top             =   810
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.OptionButton OpDraw 
         Alignment       =   1  'Right Justify
         Caption         =   "Colored Border"
         Height          =   255
         Index           =   1
         Left            =   90
         TabIndex        =   8
         Top             =   540
         Width           =   1455
      End
      Begin VB.OptionButton OpDraw 
         Alignment       =   1  'Right Justify
         Caption         =   "Black Lines"
         Height          =   255
         Index           =   0
         Left            =   90
         TabIndex        =   7
         Top             =   270
         Width           =   1455
      End
   End
   Begin VB.CheckBox chkTransparent 
      Alignment       =   1  'Right Justify
      Caption         =   "Transparent"
      Height          =   285
      Left            =   6330
      TabIndex        =   5
      Top             =   1650
      Width           =   1635
   End
   Begin VB.CheckBox chkProgress 
      Alignment       =   1  'Right Justify
      Caption         =   "Show Progress"
      Height          =   285
      Left            =   6300
      TabIndex        =   4
      Top             =   3600
      Value           =   1  'Checked
      Width           =   1635
   End
   Begin VB.CheckBox chkRender 
      Alignment       =   1  'Right Justify
      Caption         =   "Rendered"
      Height          =   285
      Left            =   6330
      TabIndex        =   3
      Top             =   1320
      Value           =   1  'Checked
      Width           =   1635
   End
   Begin MSComDlg.CommonDialog CDialog1 
      Left            =   7440
      Top             =   4860
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Save Pic"
      Height          =   465
      Left            =   6330
      TabIndex        =   2
      Top             =   660
      Width           =   1305
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Open"
      Height          =   465
      Left            =   6330
      TabIndex        =   1
      Top             =   180
      Width           =   1305
   End
   Begin VectorPreviewCtl.VectorPreview VectorPreview1 
      Height          =   5295
      Left            =   90
      TabIndex        =   0
      Top             =   180
      Width           =   6075
      _ExtentX        =   10716
      _ExtentY        =   9340
      ShowProgress    =   -1  'True
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub chkProgress_Click()
VectorPreview1.ShowProgress = CBool(chkProgress.Value)
End Sub

Private Sub chkRender_Click()
VectorPreview1.Rendered = CBool(chkRender.Value)
End Sub


Private Sub chkTransparent_Click()
If chkTransparent.Value = 0 Then
VectorPreview1.Transparent = 255
Else
VectorPreview1.Transparent = 100
End If
End Sub


Private Sub Command1_Click()
Dim fName As String

CDialog1.Filter = "Compatible Vectors Files|*.plt;*.eps;*.ai"
CDialog1.ShowOpen

fName = CDialog1.FileName

If Len(fName) > 0 Then
VectorPreview1.OpenVectorFile fName
End If
End Sub


Private Sub Command2_Click()
VectorPreview1.SaveImage App.Path & "\Test.bmp"
MsgBox "saved as " & App.Path & "\Test.bmp"
End Sub


Private Sub OpDraw_Click(Index As Integer)
VectorPreview1.DrawMode = Index
End Sub


