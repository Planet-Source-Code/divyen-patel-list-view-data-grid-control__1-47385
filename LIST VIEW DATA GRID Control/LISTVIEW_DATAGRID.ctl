VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.UserControl LISTVIEW_DATAGRID 
   ClientHeight    =   3705
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4785
   ScaleHeight     =   3705
   ScaleWidth      =   4785
   ToolboxBitmap   =   "LISTVIEW_DATAGRID.ctx":0000
   Begin MSComctlLib.ListView LIST_VIEW_GRID 
      Height          =   3375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   5953
      View            =   3
      Arrange         =   1
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Unicode MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
End
Attribute VB_Name = "LISTVIEW_DATAGRID"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim DB As New ADODB.Connection
Dim RS As New ADODB.Recordset
Dim I As Integer
Dim STR As String
Public Event GRIDCLICK()
Public Event GRIDKEYDOWN(KEYCODE As Integer)

Private Sub LIST_VIEW_GRID_ItemClick(ByVal Item As MSComctlLib.ListItem)
        RaiseEvent GRIDCLICK
        SELECTED_KEY
End Sub
Private Sub LIST_VIEW_GRID_KeyDown(KEYCODE As Integer, Shift As Integer)
        RaiseEvent GRIDKEYDOWN(KEYCODE)
End Sub


Private Sub UserControl_GetDataMember(DataMember As String, Data As Object)
    Data = RS
End Sub

Private Sub UserControl_Initialize()
    ' ALLOW TO SELECT FULL ROW WHEN YOU CLICK OR SELECT PERTICULAR ITEM
     LIST_VIEW_GRID.FullRowSelect = True
End Sub
Private Sub UserControl_Resize()
    'RESIZE GRID ACCORDING TO USERCONTROL'S SIZE
    LIST_VIEW_GRID.Width = UserControl.Width
    LIST_VIEW_GRID.Height = UserControl.Height
End Sub
Public Property Let CONNECTION_STRING(ByVal vNewValue As String)
   'OPENS THE CONNECTION
    DB.Open vNewValue
    PropertyChanged "CONNECTION_STRING"
    
End Property

Public Function FILL_RECORDS(STR_QRY As String)
   'OPEN THE RECORDSET OF THE SPECIFIED QUERY
        RS.Open STR_QRY, DB, adOpenKeyset, adLockOptimistic
   
   
   
   
   'FILL THE RECORD IN THE LISTVIEW CONTROL
With LIST_VIEW_GRID
            'ADDS THE HEADERS AS FIELD'S CAPTION
            For I = 1 To RS.Fields.Count
                    .ColumnHeaders.Add I, , RS.Fields(I - 1).Name
                    .ColumnHeaders.Item(I).Width = RS.Fields(I - 1).ActualSize * 120
            Next
            
            While RS.EOF <> True
                        'ADD KEY RECORD , FIRST RECORD MUST BE THE KEY FIELD
                        .ListItems.Add 1, , RS.Fields(0).Value & " "
                        'ADD SUB ITEMS UNDER KEY RECORD
                        For I = 1 To RS.Fields.Count - 1
                            If Len(RS.Fields(I).Value) > 0 Then
                                    .ListItems(1).ListSubItems.Add I, , RS.Fields(I).Value
                            Else
                                    .ListItems(1).ListSubItems.Add I, , Chr(0)
                            End If
                        Next
                RS.MoveNext
            Wend
End With
End Function
Public Function SELECTED_KEY() As Variant
    'FUNCTION RETURNS THE SELECTED RECORD'S KEY
    SELECTED_KEY = LIST_VIEW_GRID.SelectedItem
End Function


Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
        PropBag.WriteProperty "FORECOLOR", LIST_VIEW_GRID.FORECOLOR
        PropBag.WriteProperty "AllowColumnReorder", LIST_VIEW_GRID.AllowColumnReorder
        PropBag.WriteProperty "BACKCOLOR", LIST_VIEW_GRID.BACKCOLOR
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
        LIST_VIEW_GRID.FORECOLOR = PropBag.ReadProperty("FORECOLOR")
        LIST_VIEW_GRID.AllowColumnReorder = PropBag.ReadProperty("AllowColumnReorder")
        LIST_VIEW_GRID.BACKCOLOR = PropBag.ReadProperty("BACKCOLOR")
End Sub



Public Property Let SORTKEY(ByVal vNewValue As Boolean)
        LIST_VIEW_GRID.Sorted = True
End Property

Public Property Let SORT_ON_INDEX(ByVal vNewValue As Integer)
        LIST_VIEW_GRID.SORTKEY = vNewValue
End Property

Public Property Let AllowColumnReorder(ByVal vNewValue As Boolean)
        LIST_VIEW_GRID.AllowColumnReorder = vNewValue
        PropertyChanged "AllowColumnReorder"
End Property

Public Property Get BACKCOLOR() As OLE_COLOR
    BACKCOLOR = LIST_VIEW_GRID.BACKCOLOR
End Property

Public Property Let BACKCOLOR(ByVal vNewValue As OLE_COLOR)
        LIST_VIEW_GRID.BACKCOLOR = vNewValue
        PropertyChanged "BACKCOLOR"
End Property

Public Property Get FORECOLOR() As OLE_COLOR
        FORECOLOR = LIST_VIEW_GRID.FORECOLOR
End Property

Public Property Let FORECOLOR(ByVal vNewValue As OLE_COLOR)
        LIST_VIEW_GRID.FORECOLOR = vNewValue
End Property
