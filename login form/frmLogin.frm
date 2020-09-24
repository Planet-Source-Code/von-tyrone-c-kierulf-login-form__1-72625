VERSION 5.00
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Begin VB.Form frmLogin 
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   2535
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5550
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   2535
   ScaleWidth      =   5550
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Refresh"
      Height          =   615
      Left            =   120
      TabIndex        =   7
      Top             =   4200
      Width           =   1575
   End
   Begin MCI.MMControl MMControl1 
      Height          =   735
      Left            =   600
      TabIndex        =   6
      Top             =   2880
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   1296
      _Version        =   393216
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin VB.CommandButton cmdcancel 
      Caption         =   "Exit"
      Height          =   375
      Left            =   3720
      TabIndex        =   3
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "OK"
      Height          =   375
      Left            =   2400
      TabIndex        =   2
      Top             =   1785
      Width           =   1215
   End
   Begin VB.TextBox txtpass 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   2400
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1320
      Width           =   2535
   End
   Begin VB.TextBox txtuser 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   2400
      TabIndex        =   0
      Top             =   840
      Width           =   2535
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackColor       =   &H00004000&
      BackStyle       =   0  'Transparent
      Caption         =   "&User Name:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   0
      Left            =   510
      TabIndex        =   5
      Top             =   945
      Width           =   1350
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackColor       =   &H00004000&
      BackStyle       =   0  'Transparent
      Caption         =   "&Password:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   1
      Left            =   510
      TabIndex        =   4
      Top             =   1305
      Width           =   1245
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   50
      Height          =   2655
      Left            =   600
      Shape           =   1  'Square
      Top             =   8040
      Width           =   9735
   End
   Begin VB.Image Image2 
      Height          =   765
      Left            =   840
      Top             =   7920
      Width           =   5250
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   50
      Height          =   3255
      Left            =   3120
      Shape           =   3  'Circle
      Top             =   7680
      Width           =   4695
   End
   Begin VB.Image Image3 
      Height          =   2550
      Left            =   0
      Picture         =   "frmLogin.frx":0000
      Top             =   0
      Width           =   5550
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public LoginSuccess As Boolean
Dim RS As ADODB.Recordset
Private Sub cmdcancel_Click()
  On Error Resume Next
 
   
      End
   
End Sub

Private Sub cmdok_Click()
    Set RS = New ADODB.Recordset
    RS.CursorLocation = adUseClient
    RS.Open "SELECT * FROM tblAccess WHERE User = '" + txtuser + "' and Pass = '" + txtpass + "' ", CN, adOpenDynamic, adLockOptimistic
    
   
        If (txtpass.Text = "") And (txtuser.Text = "") Then '1
                MsgBox "Please Type Username and Password          ", vbOKOnly + vbInformation, " Access Denied"

                txtpass.Text = ""
                txtuser.Text = ""
                txtuser.SetFocus
                

        ElseIf (txtuser.Text = "") Then
                MsgBox "Please Type Username          ", vbOKOnly + vbInformation, " Access Denied"
                
                txtuser.Text = ""
                txtuser.SetFocus
        ElseIf (txtpass.Text = "") Then
        
                MsgBox "Please Type Password          ", vbOKOnly + vbInformation, " Access Denied"
                
                txtpass.Text = ""
                txtpass.SetFocus
                
                'Exit Sub
       'SendKeys "{Home}+{End}"
' ----------------------------------------------------------------------------------
            
        Else   'This else part signifies that txtuser & txtpas has been filled up...
        With RS
            
             
           
             If RS.RecordCount = 1 Then '2
                 
              
                    Unload Me
                    'cmdRefresh_Click
                    MsgBox "Access Code Accepted!          ", vbInformation, "Access Granted"
                   frmMain.Show
                   Unload Me
            Else
            'login denied
            MsgBox "Password and/or Username  Mismatch!            ", vbInformation, " Access Denied"
            'Call DeniedLog 'audio prompt "access denied"
            txtpass.Text = ""
            txtuser.Text = ""
            txtuser.SetFocus
            'SendKeys "{Home}+{End}"
                
            End If '2
        
        
        End With
        
        End If '1
  '----------------- 'Below recycle line' ( might still use them later )
 
  
  'LoginSuccess = True
  'If LoginSuccess = True Then '3
 'End If '3
End Sub
Private Sub Form_Load()
 Me.Show
    Set CN = New ADODB.Connection
    CN.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\MasterFile.mdb;Persist Security Info=False;Jet OLEDB:Database Password=MLEVDQ48L2"
    'CN.Close
End Sub
