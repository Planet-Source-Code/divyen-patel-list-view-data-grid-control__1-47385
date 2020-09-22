VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "List View Grid Control Test Form ...."
   ClientHeight    =   8490
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   11850
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8490
   ScaleWidth      =   11850
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin LIST_VIEW_DATA_GRID.LISTVIEW_DATAGRID LISTVIEW_DATAGRID1 
      Height          =   7455
      Left            =   240
      TabIndex        =   2
      Top             =   480
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   13150
      FORECOLOR       =   -2147483640
      AllowColumnReorder=   -1  'True
      BACKCOLOR       =   16443364
   End
   Begin VB.Label Label4 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2640
      TabIndex        =   1
      Top             =   120
      Width           =   2295
   End
   Begin VB.Label Label3 
      Caption         =   "Selected Customer Id :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    'SET CONNECTION STRING
            LISTVIEW_DATAGRID1.CONNECTION_STRING = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\DATABASE.MDB;Persist Security Info=False"
    'FILL RECORD IN TO LISTVIEW BY QUERY
            LISTVIEW_DATAGRID1.FILL_RECORDS "SELECT * FROM CUSTOMER"
    'SET THE SORT KEY AND SORT INDEX
            LISTVIEW_DATAGRID1.SORTKEY = True
            LISTVIEW_DATAGRID1.SORT_ON_INDEX = 0
    'WITH THIS PROPERRY YOU CAN REORDER COLUMN AT RUNTIME
            LISTVIEW_DATAGRID1.AllowColumnReorder = True
End Sub



Private Sub LISTVIEW_DATAGRID1_GRIDCLICK()
            Label4.Caption = LISTVIEW_DATAGRID1.SELECTED_KEY
End Sub


Private Sub LISTVIEW_DATAGRID1_GRIDKEYDOWN(KEYCODE As Integer)
If KEYCODE = 13 Then
        MsgBox "SELECTED RECORD'S KEY IS : " + LISTVIEW_DATAGRID1.SELECTED_KEY
End If
End Sub
