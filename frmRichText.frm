VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmRichText 
   Caption         =   "Form1"
   ClientHeight    =   4455
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6900
   LinkTopic       =   "Form1"
   ScaleHeight     =   4455
   ScaleWidth      =   6900
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Open Text to Picture"
      Height          =   375
      Left            =   2100
      TabIndex        =   1
      Top             =   3750
      Width           =   1890
   End
   Begin RichTextLib.RichTextBox myrtf1 
      Height          =   3345
      Left            =   150
      TabIndex        =   0
      Top             =   210
      Width           =   6555
      _ExtentX        =   11562
      _ExtentY        =   5900
      _Version        =   393217
      Enabled         =   -1  'True
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"frmRichText.frx":0000
   End
End
Attribute VB_Name = "frmRichText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim OpenCommonDialog As New CommonDialog
Private Const WM_PASTE = &H302
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Sub Command1_Click()
    frmFormatText.Show
    End Sub



Sub ConvertImage()
    Dim theimage As StdPicture
    Set theimage = ConvertTextToImage(gTextToImage)
   
    ' insert picture
    Clipboard.Clear
    Clipboard.SetData theimage, vbCFBitmap
    SendMessage myrtf1.hwnd, WM_PASTE, 0, 0
   
End Sub

Function ConvertTextToImage(s As String) As StdPicture
   Dim x As PictureBox
   
    ' add temporary picture control
    Set x = Me.Controls.Add("VB.Picturebox", "picTemp")
   
    With x
        .Font.Size = frmFormatText.txtFormatText.FontSize '  tText1.FontSize
        .Font = frmFormatText.txtFormatText.Font
       .ForeColor = frmFormatText.txtFormatText.ForeColor   ' Text1.ForeColor
       .FontBold = BoldFont
       .FontItalic = ItalicFont
       .FontUnderline = ItalicFont
        .Visible = True
        .BackColor = myrtf1.BackColor
        .Width = .TextWidth(s)
        .Height = Trim(.TextHeight(s))
        .AutoRedraw = True
        .BorderStyle = 0
        
        '.FontUnderline = True
        x.Print s
        .AutoRedraw = False
    End With

    Set ConvertTextToImage = x.Image
   
    Me.Controls.Remove "pictemp" ' remove temporary picture control
End Function

