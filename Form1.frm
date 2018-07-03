VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FF8080&
   Caption         =   "Insert X to field"
   ClientHeight    =   3225
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11940
   LinkTopic       =   "Form1"
   ScaleHeight     =   3225
   ScaleWidth      =   11940
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF8080&
      Height          =   1215
      Left            =   5760
      TabIndex        =   15
      Top             =   1320
      Width           =   4095
      Begin VB.OptionButton after 
         BackColor       =   &H00FF8080&
         Caption         =   "After field value"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   6
         Top             =   120
         Width           =   3255
      End
      Begin VB.OptionButton befor 
         BackColor       =   &H00FF8080&
         Caption         =   "Befor field value"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   7
         Top             =   600
         Value           =   -1  'True
         Width           =   3375
      End
   End
   Begin VB.TextBox x 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   480
      Left            =   10560
      TabIndex        =   8
      Text            =   "0"
      Top             =   1680
      Width           =   975
   End
   Begin VB.TextBox tot 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   480
      Left            =   4680
      TabIndex        =   5
      Top             =   1560
      Width           =   975
   End
   Begin VB.OptionButton manualpath 
      BackColor       =   &H00FF8080&
      Caption         =   "Manula path"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Top             =   720
      Width           =   2295
   End
   Begin VB.OptionButton defaultpath 
      BackColor       =   &H00FF8080&
      Caption         =   "Default path"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Value           =   -1  'True
      Width           =   2295
   End
   Begin VB.TextBox from 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   480
      Left            =   3000
      TabIndex        =   4
      Top             =   1560
      Width           =   975
   End
   Begin VB.TextBox field 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   480
      Left            =   1320
      TabIndex        =   3
      Top             =   1560
      Width           =   615
   End
   Begin VB.TextBox path 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   3000
      TabIndex        =   2
      Top             =   720
      Width           =   8655
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Run"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   10
      Top             =   2640
      Width           =   11655
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "(0/0)-%0"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   330
      Left            =   240
      TabIndex        =   16
      Top             =   2160
      Width           =   1200
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "g:\rezamohit\list.xlsx"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   330
      Left            =   3000
      TabIndex        =   14
      Top             =   240
      Width           =   3300
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "X = "
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   1
      Left            =   9960
      TabIndex        =   13
      Top             =   1800
      Width           =   600
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "To"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   0
      Left            =   4200
      TabIndex        =   12
      Top             =   1560
      Width           =   300
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "From"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   0
      Left            =   2160
      TabIndex        =   11
      Top             =   1560
      Width           =   600
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Filde"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   360
      TabIndex        =   9
      Top             =   1560
      Width           =   750
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim oExcel As Object
Dim AD As String
If defaultpath.Value = True Then
    xpath = App.path & "\list.xlsx"
Else
    xpath = path.Text
End If

Set oExcel = GetObject(xpath)




f = field.Text

For j = from.Text To tot.Text

If befor.Value = True Then
    oExcel.ActiveSheet.Range(f & j).Value = x.Text & oExcel.ActiveSheet.Range(f & j).Value
ElseIf after.Value = True Then
    oExcel.ActiveSheet.Range(f & j).Value = oExcel.ActiveSheet.Range(f & j).Value & x.Text
End If
y = (j * 100) / Val(tot.Text)

Label5.Caption = "( " & j & " / " & tot.Text & " ) - %" & y

Next j
MsgBox "The Excel on the '" & xpath & "' path changed", vbInformation, "Operation complete"
oExcel.Parent.Windows(1).Visible = True
oExcel.Close


End Sub

Private Sub defaultpath_Click()
path.Enabled = False

End Sub

Private Sub Form_Load()
Label4.Caption = App.path & "\list.xlsx"

End Sub

Private Sub manualpath_Click()
path.Enabled = True

End Sub
