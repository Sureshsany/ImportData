VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmConfig 
   Caption         =   "Form1"
   ClientHeight    =   5820
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10905
   LinkTopic       =   "Form1"
   ScaleHeight     =   5820
   ScaleWidth      =   10905
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdConnect 
      Caption         =   "Connect"
      Height          =   375
      Left            =   4200
      TabIndex        =   6
      Top             =   2280
      Width           =   1455
   End
   Begin VB.ComboBox cboTable 
      Height          =   315
      Left            =   1080
      TabIndex        =   5
      Text            =   "Combo1"
      Top             =   3240
      Width           =   2415
   End
   Begin VB.TextBox txtPword 
      Height          =   375
      Left            =   2040
      TabIndex        =   2
      Top             =   2280
      Width           =   1815
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "Browse (MS Access)"
      Height          =   615
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   2175
   End
   Begin MSComDlg.CommonDialog ComDia 
      Left            =   8280
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label2 
      Caption         =   "Table List"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   3240
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Password (Only if required)"
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   2280
      Width           =   1575
   End
   Begin VB.Label lblDBPath 
      Caption         =   "Label1"
      Height          =   735
      Left            =   360
      TabIndex        =   1
      Top             =   1200
      Width           =   10215
   End
End
Attribute VB_Name = "frmConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
