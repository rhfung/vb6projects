VERSION 5.00
Begin VB.Form frmDim 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Start Point"
   ClientHeight    =   1605
   ClientLeft      =   2760
   ClientTop       =   3705
   ClientWidth     =   1770
   Icon            =   "dlgDimensionChange.frx":0000
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1605
   ScaleWidth      =   1770
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtZ 
      Height          =   285
      Left            =   360
      TabIndex        =   7
      Top             =   840
      Width           =   1335
   End
   Begin VB.CommandButton cmdDimenReset 
      Caption         =   "&Reset"
      Height          =   375
      Left            =   960
      TabIndex        =   5
      Top             =   1200
      Width           =   735
   End
   Begin VB.TextBox txtY 
      Height          =   285
      Left            =   360
      TabIndex        =   4
      Top             =   480
      Width           =   1335
   End
   Begin VB.TextBox txtX 
      Height          =   285
      Left            =   360
      TabIndex        =   3
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton cmdDimenOk 
      Caption         =   "&OK"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "&Z:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Width           =   255
   End
   Begin VB.Label lblDimenY 
      Caption         =   "&Y:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   255
   End
   Begin VB.Label lblDimenX 
      Caption         =   "&X:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   255
   End
End
Attribute VB_Name = "frmDim"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

