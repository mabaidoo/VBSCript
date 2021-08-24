VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form PrintPreview 
   BackColor       =   &H00FFFFFF&
   Caption         =   "PrintPreview"
   ClientHeight    =   7035
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   16260
   Icon            =   "PrintPreview.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7035
   ScaleWidth      =   16260
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog PrintDialog 
      Left            =   14400
      Top             =   3480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Close 
      Caption         =   "Close"
      Height          =   375
      Left            =   13320
      TabIndex        =   3
      Top             =   1200
      Width           =   1935
   End
   Begin VB.CommandButton Print_btn 
      Caption         =   "Print"
      Height          =   375
      Left            =   13320
      TabIndex        =   2
      Top             =   600
      Width           =   1935
   End
   Begin RichTextLib.RichTextBox preview_printerrors 
      Height          =   4695
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   12975
      _ExtentX        =   22886
      _ExtentY        =   8281
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"PrintPreview.frx":9647
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Validation Errors"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   3135
   End
End
Attribute VB_Name = "PrintPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Close_Click()
    PrintPreview.Hide
End Sub

Private Sub Print_btn_Click()
    PrintDialog.ShowPrinter
    If PrintDialog.CancelError = False Then
        preview_printerrors.SelPrint (Printer.hDC)
    End If
            
            
End Sub
Private Sub form_load()
    preview_printerrors.Text = ValidationForm.Errors.Text
    
End Sub
