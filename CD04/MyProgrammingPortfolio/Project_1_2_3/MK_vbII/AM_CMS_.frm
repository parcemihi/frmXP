VERSION 5.00
Begin VB.Form frmAtlanticContact 
   Caption         =   "Atlantic Marketing - Contact Management System"
   ClientHeight    =   4410
   ClientLeft      =   2325
   ClientTop       =   1950
   ClientWidth     =   5865
   LinkTopic       =   "Form2"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4410
   ScaleWidth      =   5865
   Begin VB.TextBox txtClientID 
      Height          =   285
      Left            =   2160
      TabIndex        =   1
      Top             =   240
      Width           =   1815
   End
   Begin VB.TextBox txtEstimatedSales 
      Height          =   285
      Left            =   2160
      TabIndex        =   10
      Top             =   2040
      Width           =   1815
   End
   Begin VB.TextBox txtNotes 
      Height          =   1335
      Left            =   960
      TabIndex        =   11
      Top             =   2400
      Width           =   4695
   End
   Begin VB.TextBox txtDateAdded 
      Height          =   285
      Left            =   2160
      TabIndex        =   9
      Top             =   1680
      Width           =   1815
   End
   Begin VB.TextBox txtTelephone 
      Height          =   285
      Left            =   2160
      TabIndex        =   8
      Top             =   1320
      Width           =   1815
   End
   Begin VB.TextBox txtFirstName 
      Height          =   285
      Left            =   2160
      TabIndex        =   7
      Top             =   960
      Width           =   1815
   End
   Begin VB.TextBox txtLastName 
      Height          =   285
      Left            =   2160
      TabIndex        =   6
      Top             =   600
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Client ID: "
      Height          =   255
      Index           =   6
      Left            =   240
      TabIndex        =   13
      Top             =   240
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Estimated Annual Sales:  "
      Height          =   255
      Index           =   5
      Left            =   240
      TabIndex        =   12
      Top             =   2040
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Telephone Number:  "
      Height          =   255
      Index           =   4
      Left            =   240
      TabIndex        =   5
      Top             =   1320
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Notes:"
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   4
      Top             =   2400
      Width           =   615
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Date Added:  "
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   3
      Top             =   1680
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "First Name:  "
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   2
      Top             =   960
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "Last Name:  "
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   1815
   End
End
Attribute VB_Name = "frmAtlanticContact"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
