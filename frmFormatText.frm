VERSION 5.00
Begin VB.Form frmFormatText 
   BackColor       =   &H00FFFFFF&
   Caption         =   "  Text to Image"
   ClientHeight    =   3960
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5580
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3960
   ScaleWidth      =   5580
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdToolBar 
      BackColor       =   &H00FFFFFF&
      Height          =   405
      Index           =   4
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   60
      Width           =   435
   End
   Begin VB.CommandButton cmdToolBar 
      BackColor       =   &H00FFFFFF&
      Height          =   405
      Index           =   3
      Left            =   1740
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   60
      Width           =   435
   End
   Begin VB.CommandButton cmdToolBar 
      BackColor       =   &H00FFFFFF&
      Height          =   405
      Index           =   2
      Left            =   1170
      MaskColor       =   &H00FF00FF&
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   60
      Width           =   435
   End
   Begin VB.CommandButton cmdToolBar 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   1
      Left            =   660
      MaskColor       =   &H00FF00FF&
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   90
      Width           =   435
   End
   Begin VB.CommandButton cmdToolBar 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   0
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   90
      Width           =   435
   End
   Begin VB.CommandButton cmdInsert 
      Caption         =   "Insert"
      Height          =   435
      Left            =   4110
      TabIndex        =   3
      Top             =   870
      Width           =   1305
   End
   Begin VB.TextBox txtFormatText 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2265
      Left            =   195
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   1485
      Width           =   5190
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   -30
      TabIndex        =   1
      Top             =   0
      Width           =   6030
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Text "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   855
      TabIndex        =   2
      Top             =   945
      Width           =   1800
   End
End
Attribute VB_Name = "frmFormatText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Dim LastKeyClicked As Integer 'This used for triming spaces from text to image transformation
Dim OpenCommonDialog As New CommonDialog



Private Sub cmdInsert_Click()
    'gfrm = "frmFormatText"
    If LastKeyClicked = 13 Then '13 enter key
      gTextToImage = Left(txtFormatText, Len(txtFormatText) - 2)  'Removes trailing spaces from textbox
      Else
      gTextToImage = txtFormatText
      End If
      
      frmRichText.ConvertImage
      Me.Visible = False
      
    End Sub

Private Sub cmdToolBar_Click(Index As Integer)
On Error GoTo bail
Select Case Index

  Case 0 '"Bold"
  txtFormatText.FontBold = Not txtFormatText.FontBold
   BoldFont = txtFormatText.FontBold
   
    Case 1 ' "Italic"
     txtFormatText.FontItalic = Not txtFormatText.FontItalic
     ItalicFont = txtFormatText.FontItalic
       
    Case 2 '"Underline"
    txtFormatText.FontUnderline = Not txtFormatText.FontUnderline
    UnderlineFont = txtFormatText.FontUnderline
       
 Case 3
      
        Dim dlg As New clsCommonDialog
        dlg.ShowColor
        txtFormatText.ForeColor = dlg.Color
    Set dlg = Nothing

           
 Case 4   '"Font"
        
       
        OpenCommonDialog.Flags = cdlCFBoth Or cdlCFEffects
        With txtFormatText
            OpenCommonDialog.FontName = .FontName
            OpenCommonDialog.FontSize = .FontSize
            OpenCommonDialog.Bold = .FontBold
           OpenCommonDialog.Italic = .FontItalic
            OpenCommonDialog.Underline = .FontUnderline
            OpenCommonDialog.rgbResult = .ForeColor
        End With
        OpenCommonDialog.ShowFont
        With txtFormatText
            .FontName = OpenCommonDialog.FontName
            .FontSize = OpenCommonDialog.FontSize
            .FontBold = OpenCommonDialog.Bold
            .FontItalic = OpenCommonDialog.Italic
            .FontUnderline = OpenCommonDialog.Underline
            .ForeColor = OpenCommonDialog.rgbResult
     End With
      
    End Select
    
bail:
    End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 2 Then

Select Case KeyCode
    Case vbKeyT
     cmdInsert_Click
   
    End Select
   End If
End Sub

Private Sub Form_Load()
    On Error Resume Next
        
    Label2.ForeColor = vbBlue
    Label1.Width = Me.Width
    Me.cmdInsert.Enabled = False
cmdToolBar(0).Picture = LoadResPicture(101, vbResBitmap)
cmdToolBar(1).Picture = LoadResPicture(102, vbResBitmap)
cmdToolBar(2).Picture = LoadResPicture(103, vbResBitmap)
cmdToolBar(3).Picture = LoadResPicture(104, vbResBitmap)
cmdToolBar(4).Picture = LoadResPicture(105, vbResBitmap)


End Sub

'To control the space under coversation on of text to image
Private Sub txtFormatText_KeyPress(KeyAscii As Integer)
   If KeyAscii = 9 Then KeyAscii = 0  'Hack to keep ctrl-i from moving a tab
    LastKeyClicked = KeyAscii
    'cmdInsert.Visible = True
    Me.cmdInsert.Enabled = True
    End Sub

Private Sub txtFormatText_KeyUp(KeyCode As Integer, Shift As Integer)
   If Shift <> 2 Then Exit Sub
  Select Case KeyCode
  Case vbKeyB  '66  ' bold ctrl-b
  cmdToolBar_Click (9)
  Shift = 0
  Case vbKeyI    'italic ctrl-i
   cmdToolBar_Click (10)
   Shift = 0
   KeyCode = 0
   Case vbKeyU '85  'underline UUUUUUUUUUUUUUUUUUU
    cmdToolBar_Click (11)
    Shift = 0
   Case vbKeyL  '76  'color        LLLLLLLLLLLLLLLLLL
    cmdToolBar_Click (15)
  
  Case vbKeyD '68 'ctrl-d font dialog
  cmdToolBar_Click (16)
  Shift = 0
  End Select
End Sub

