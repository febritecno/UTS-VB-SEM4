VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BackColor       =   &H80000001&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   9435
   ClientLeft      =   3240
   ClientTop       =   1830
   ClientWidth     =   14085
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9435
   ScaleWidth      =   14085
   ShowInTaskbar   =   0   'False
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   9120
      Top             =   5760
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command6 
      Caption         =   "OPEN"
      Height          =   855
      Left            =   10560
      TabIndex        =   18
      Top             =   7200
      Width           =   2295
   End
   Begin VB.PictureBox Picture1 
      DataSource      =   "Adodc1"
      Height          =   4935
      Left            =   9600
      ScaleHeight     =   4875
      ScaleWidth      =   3795
      TabIndex        =   17
      Top             =   1800
      Width           =   3855
   End
   Begin VB.TextBox Text5 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3840
      TabIndex        =   4
      Top             =   4680
      Width           =   4335
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3840
      TabIndex        =   3
      Top             =   4080
      Width           =   4335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Tambah"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3480
      Picture         =   "Form1.frx":0000
      TabIndex        =   9
      Top             =   5520
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Edit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5520
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5520
      TabIndex        =   7
      Top             =   5520
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Hapus"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6600
      TabIndex        =   6
      Top             =   5520
      Width           =   975
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Keluar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7560
      TabIndex        =   5
      Top             =   5520
      Width           =   975
   End
   Begin VB.TextBox Text3 
      Height          =   885
      Left            =   3840
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   3000
      Width           =   3615
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3840
      TabIndex        =   1
      Top             =   2400
      Width           =   4335
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3840
      TabIndex        =   0
      Top             =   1800
      Width           =   4335
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   120
      Top             =   8280
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=F:\USER\Videos\uts\db.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=F:\USER\Videos\uts\db.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "dvd_film"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Form1.frx":0407
      Height          =   2895
      Left            =   1800
      TabIndex        =   10
      Top             =   6360
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   5106
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   6
      BeginProperty Column00 
         DataField       =   "kode_dvd"
         Caption         =   "kode_dvd"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "judul"
         Caption         =   "judul"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "deskripsi"
         Caption         =   "deskripsi"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "genre"
         Caption         =   "genre"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "author"
         Caption         =   "author"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "img"
         Caption         =   "img"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   915.024
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1739.906
         EndProperty
      EndProperty
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000005&
      BorderStyle     =   3  'Dot
      BorderWidth     =   2
      DrawMode        =   6  'Mask Pen Not
      FillStyle       =   4  'Upward Diagonal
      Height          =   5175
      Left            =   9480
      Top             =   1680
      Width           =   4095
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "AUTHOR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1800
      TabIndex        =   16
      Top             =   4800
      Width           =   1935
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "DESKRIPSI"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1800
      TabIndex        =   15
      Top             =   3120
      Width           =   1935
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "JUDUL DVD"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   1800
      TabIndex        =   14
      Top             =   2400
      Width           =   1935
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "KODE DVD"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   1800
      TabIndex        =   13
      Top             =   1800
      Width           =   1935
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "ENTRY DVD FILM"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   495
      Left            =   3720
      TabIndex        =   12
      Top             =   360
      Width           =   6495
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "KATAGORI"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1800
      TabIndex        =   11
      Top             =   4080
      Width           =   1935
   End
   Begin VB.Menu abt 
      Caption         =   "About"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'===========================
'Febrian Dwi Putra

Private Sub abt_Click()
Dialog.Show
End Sub

'1602040627

'UTS

'FTI - B

'============================================

Private Sub Command1_Click()
mati (True)
Call control(True, True, True, False, False, True, True, True, True, True, True)
Text1.SetFocus
kosong
End Sub

Private Sub Command2_Click()
If Command2.Caption = "Edit" Then
Call control(False, True, True, False, True, False, True, True, True, True, True)
Command2.Caption = "Save Edit"
Text2.SetFocus
Else
If Text1 = "" And Text2 = "" And Text3 = "" And Text4 = "" And Text5 = "" Then
    MsgBox "Masih ada data yang kosong..!!!", vbCritical, "Error!"
        Else
        
db
With Adodc1.Recordset
    !kode_dvd = Text1
    !judul = Text2
    !deskripsi = Text3
    !genre = Text4
    !author = Text5
    .Update
End With
Call control(False, False, False, True, True, False, True, True, False, False, False)
Command2.Caption = "Edit"

    End If
End If

End Sub

Private Sub Command3_Click()
If Text1.Text = "" And Text2.Text = "" And Text3.Text = "" And Text4 = "" And Text5 = "" Then
MsgBox "error"
Else
With Adodc1.Recordset
    .AddNew
    !kode_dvd = Text1
    !judul = Text2
    !deskripsi = Text3
    !genre = Text4
    !author = Text5
    .Update
    End With
    MsgBox "Disimpan!", vbOKOnly, "Berhasil!"
        kosong
        Call control(False, False, False, True, True, False, True, True, False, False, False)
        
End If
End Sub

Private Sub Command4_Click()
Dim hapus As String
db
    If Adodc1.Recordset.RecordCount <> 0 Then
        hapus = MsgBox("Yakin akan dihapus?", vbYesNo, "Peringatan...!")
        If hapus = vbYes Then
            If Adodc1.Recordset.EOF Then
                MsgBox "kosong"
            Else
                Adodc1.Recordset.Delete
                Adodc1.Recordset.MoveNext
                Call Form_Load
                    On Error GoTo eek
                Kill App.Path & "\img\upload_" & Text1.Text & ".jpeg"
eek:
                    If err.Number <> 0 Then
                    MsgBox "Gambar Kosong / Hilang"
                    End
                    Else
                    Me.Show
                    End If
                End If
        End If
    Else
        MsgBox "Data kosong...", vbInformation, "Informasi!"
End If

End Sub

Private Sub Command5_Click()
Unload Me
End Sub

Sub kosong()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
End Sub

Private Sub Command6_Click()
On Error GoTo bujat
With CommonDialog1

.Filter = "Semua Gambar |*.jpg;*.jpeg;*.gif;*.bmp;*.png"

.ShowOpen

Picture1.Picture = LoadPicture(CommonDialog1.FileName)

SavePicture Picture1.Picture, App.Path & "\img\upload_" & Text1.Text & ".jpeg"
    Picture1.ScaleMode = 3
    Picture1.AutoRedraw = True
    Picture1.PaintPicture Picture1.Picture, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, 0, 0, Picture1.Picture.Width / 26.46, Picture1.Picture.Height / 26.46

bujat:
If err.Number <> 0 Then
   If err.Number = 2755 Then Exit Sub
   MsgBox "Pilih Gambar Yang Benar", vbCritical, "Error"
End If

End With
End Sub
Private Sub DataGrid1_Click()
isi
mati (False)
Call Form_Load
End Sub


Private Sub DataGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
isi
mati (False)
Call Form_Load
End Sub

Private Sub DataGrid1_KeyUp(KeyCode As Integer, Shift As Integer)
isi
mati (False)
Call Form_Load
End Sub
Private Sub Form_Load()
Adodc1.Visible = False
isi
mati (False)
Call control(False, False, False, True, True, False, True, True, False, False, False)
End Sub
Sub isi()
Dim pathgambar As String
If Adodc1.Recordset.EOF Then
MsgBox "Data Masih Kosong, Silahkan Mengisi", vbCritical, "Info"
Me.Show
Else
Text1.Text = Adodc1.Recordset.Fields("kode_dvd")
Text2.Text = Adodc1.Recordset.Fields("judul")
Text3.Text = Adodc1.Recordset.Fields("deskripsi")
Text4.Text = Adodc1.Recordset.Fields("genre")
Text5.Text = Adodc1.Recordset.Fields("author")

pathgambar = App.Path & "\img\upload_" & Text1.Text & ".jpeg"
                
        On Error GoTo err
        
        Picture1.Picture = LoadPicture(pathgambar)
        Picture1.ScaleMode = 3
        Picture1.AutoRedraw = True
        Picture1.PaintPicture Picture1.Picture, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, 0, 0, Picture1.Picture.Width / 26.46, Picture1.Picture.Height / 26.46
err:
        If err.Number <> 0 Then
         MsgBox "Gambar Kosong, Edit dan Open Gambar", vbCritical, "Info"
         Else
         Me.Show
         End If

End If
Command2.Caption = "Edit"
End Sub

Sub mati(x)
Text1.Enabled = x
Text2.Enabled = x
Text3.Enabled = x
Text4.Enabled = x
Text5.Enabled = x
End Sub

Function control(t1, t2, t3, a1, a2, a3, a4, a5, d3, d4, img)
Text1.Enabled = t1
Text2.Enabled = t2
Text3.Enabled = t3
Text4.Enabled = d3
Text5.Enabled = d4
Command6.Enabled = img
Command1.Enabled = a1
Command2.Enabled = a2
Command3.Enabled = a3
Command4.Enabled = a4
Command5.Enabled = a5
End Function

