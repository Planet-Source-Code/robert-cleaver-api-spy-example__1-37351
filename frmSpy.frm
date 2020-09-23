VERSION 5.00
Begin VB.Form frmSpy 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Api Spy Thing"
   ClientHeight    =   3405
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4365
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "frmSpy.frx":0000
   ScaleHeight     =   3405
   ScaleWidth      =   4365
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Spy"
      Height          =   3360
      Left            =   60
      MouseIcon       =   "frmSpy.frx":030A
      TabIndex        =   0
      Top             =   0
      Width           =   4260
      Begin VB.CommandButton cmdCont 
         Caption         =   "Cont."
         Height          =   240
         Left            =   90
         MouseIcon       =   "frmSpy.frx":0614
         TabIndex        =   10
         Top             =   900
         Width           =   900
      End
      Begin VB.Frame Frame3 
         Caption         =   "Window Information"
         Height          =   1710
         Left            =   1050
         TabIndex        =   3
         Top             =   180
         Width           =   3120
         Begin VB.TextBox txtWindowText 
            Alignment       =   2  'Center
            BackColor       =   &H80000000&
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   45
            TabIndex        =   9
            Text            =   "--"
            Top             =   1380
            Width           =   3030
         End
         Begin VB.TextBox txtWindowClassName 
            Alignment       =   2  'Center
            BackColor       =   &H80000000&
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   45
            TabIndex        =   7
            Text            =   "--"
            Top             =   960
            Width           =   3030
         End
         Begin VB.TextBox txtWindowHandle 
            Alignment       =   2  'Center
            BackColor       =   &H80000000&
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   45
            TabIndex        =   5
            Text            =   "--"
            Top             =   480
            Width           =   3030
         End
         Begin VB.Label Label3 
            Caption         =   "Window Text:"
            Height          =   180
            Left            =   90
            TabIndex        =   8
            Top             =   1200
            Width           =   1425
         End
         Begin VB.Label Label2 
            Caption         =   "Window Class Name:"
            Height          =   180
            Left            =   90
            TabIndex        =   6
            Top             =   720
            Width           =   2925
         End
         Begin VB.Label Label1 
            Caption         =   "Window Handle:"
            Height          =   180
            Left            =   90
            TabIndex        =   4
            Top             =   240
            Width           =   2910
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Drag Me"
         Height          =   675
         Left            =   90
         TabIndex        =   1
         Top             =   180
         Width           =   900
         Begin VB.PictureBox picDrag 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   360
            Left            =   255
            Picture         =   "frmSpy.frx":091E
            ScaleHeight     =   360
            ScaleWidth      =   405
            TabIndex        =   2
            Top             =   240
            Width           =   405
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Parent List"
         Height          =   1395
         Left            =   120
         TabIndex        =   11
         Top             =   1860
         Width           =   4035
         Begin VB.ListBox lstParents 
            Height          =   1035
            Left            =   120
            TabIndex        =   12
            Top             =   240
            Width           =   3795
         End
      End
   End
End
Attribute VB_Name = "frmSpy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCont_Click()
    frmContCode.Show vbModal
End Sub

Private Sub picDrag_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    picDrag.MousePointer = 99
    Me.MousePointer = 99
    picDrag.Picture = Me.Picture
    InformationNow = True
End Sub

Private Sub picDrag_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If InformationNow = True Then
        Call GetWindowInformation(WindowHandle&, WindowClassName$, WindowText$, lstParents)
        txtWindowHandle.Text = WindowHandle&
        txtWindowClassName.Text = WindowClassName$
        txtWindowText.Text = WindowText$
    End If
End Sub

Private Sub picDrag_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim WindowHandle&
Dim WindowClassName$, WindowText$
    picDrag.MousePointer = 0
    Me.MousePointer = 0
    picDrag.Picture = cmdCont.MouseIcon
    Call GetWindowInformation(WindowHandle&, WindowClassName$, WindowText$, lstParents)
    txtWindowHandle.Text = WindowHandle&
    txtWindowClassName.Text = WindowClassName$
    txtWindowText.Text = WindowText$
    InformationNow = False
End Sub
