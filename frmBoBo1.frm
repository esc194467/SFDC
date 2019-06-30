VERSION 5.00
Object = "{AC957AC8-A491-11CE-A69B-0000E8A490E7}#1.0#0"; "wdtvocx.ocx"
Begin VB.Form frmBoBo1 
   Caption         =   "Currently Booked On"
   ClientHeight    =   8670
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9915
   LinkTopic       =   "Form1"
   ScaleHeight     =   8670
   ScaleWidth      =   9915
   StartUpPosition =   3  'Windows Default
   Begin SAPTableView.SAPTableView stvBoBo 
      Height          =   4095
      Left            =   1080
      TabIndex        =   5
      Top             =   2160
      Width           =   7215
      _Version        =   3
      _ExtentX        =   12726
      _ExtentY        =   7223
      _StockProps     =   109
      BackColor       =   16777215
      BeginProperty CtrlVColsVCollFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty CtrlVRowsVCollFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox CheckNumber 
      BackColor       =   &H8000000B&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   360
      Width           =   1575
   End
   Begin VB.TextBox PersonName 
      BackColor       =   &H8000000B&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   1
      Top             =   360
      Width           =   4335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   3600
      TabIndex        =   0
      Top             =   7560
      Width           =   2415
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Check Number"
      Height          =   195
      Left            =   240
      TabIndex        =   4
      Top             =   120
      Width           =   1065
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Name"
      Height          =   195
      Left            =   2160
      TabIndex        =   3
      Top             =   120
      Width           =   420
   End
End
Attribute VB_Name = "frmBoBo1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim BoBoTableColumn As SAPTableFactoryCtrl.Column



Private Sub Command1_Click()


   'Set export parameters fo SAPReadPrevBookings
   SAPReadBoBo.Exports("PERSON") = "00029008"
   SAPReadBoBo.Exports("TIMEZONE") = "UTC+8"

   'Call the RFC
   If SAPReadBoBo.Call = False Then
        'set pointer to standard
        MousePointer = vbDefault
        
        MsgBox SAPReadBoBo.Exception, vbExclamation, DialogTitleLogon
        'Exit Sub
    Else
        'BoBoTable.Refresh
        
   End If


 Exit Sub
 

End Sub

Private Sub Form_Load()

   Dim TabColIndex, DispColCount, ColCount As Integer
   Dim ColumnName As String
   
   'Create the BoBo Table and View Objects
   Set BoBoTable = SAPReadBoBo.Tables.Item("BOBO")
   If BoBoTable.Views.Count = 0 Then
    BoBoTable.Views.Add frmBoBo.stvBoBo.object
    Set BoBoview = BoBoTable.Views.Item(1)
   End If
   
   BoBoview.Columns.Height = 600
   
   ColCount = BoBoTable.Columns.Count
   DispColCount = ColCount
   
   For TabColIndex = 1 To ColCount
    Set BoBoViewColumn = BoBoview.Columns.Item(TabColIndex)
    Set BoBoTableColumn = BoBoTable.Columns.Item(TabColIndex)
    ColumnName = BoBoTableColumn.Name
    Debug.Print ColumnName
    Select Case ColumnName
        Case "ORDER_NO"
            BoBoViewColumn.Header = "Order"
            BoBoViewColumn.Type = tavColumnGeneral
            BoBoViewColumn.Protection = True
        Case "ACTIVITY"
            BoBoViewColumn.Header = "Activity"
            BoBoViewColumn.Type = tavColumnGeneral
            BoBoViewColumn.Protection = True
       Case "ON_DATE"
            BoBoViewColumn.Header = "Clock-On" & Chr(10) & "Date"
            BoBoViewColumn.Type = tavColumnGeneral
            BoBoViewColumn.Width = 12
            BoBoViewColumn.Protection = True
       Case "ON_TIME"
            BoBoViewColumn.Header = "Clock-On" & Chr(10) & "Time"
            BoBoViewColumn.Type = tavColumnGeneral
            BoBoViewColumn.Width = 12
            BoBoViewColumn.Protection = True
        Case "FIN_CONF"
            BoBoViewColumn.Header = "Completed"
            BoBoViewColumn.Type = tavColumnBoolean
            BoBoViewColumn.Width = 12
        Case "CLOCK_OFF"
            BoBoViewColumn.Header = "Clock-Off"
            BoBoViewColumn.Type = tavColumnBoolean
            BoBoViewColumn.Width = 12
        Case Else
            BoBoViewColumn.Protection = True
            BoBoViewColumn.Visible = False
            DispColCount = DispColCount - 1
    End Select
    
 Next
     
   BoBoTable.FreeTable 'remove existing data in table
 
End Sub
