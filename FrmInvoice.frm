VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FrmInvoice 
   BackColor       =   &H00FFFF00&
   Caption         =   "Invoice"
   ClientHeight    =   4965
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13620
   LinkTopic       =   "Form1"
   ScaleHeight     =   4965
   ScaleWidth      =   13620
   StartUpPosition =   1  'CenterOwner
   Begin Crystal.CrystalReport crInvoice 
      Left            =   5880
      Top             =   4200
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton btnKeluar 
      Caption         =   "&Keluar"
      Height          =   495
      Left            =   4440
      TabIndex        =   13
      Top             =   4200
      Width           =   1215
   End
   Begin MSComCtl2.DTPicker dtpTanggalInvoice 
      Height          =   495
      Left            =   1680
      TabIndex        =   12
      Top             =   2640
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
      _Version        =   393216
      Format          =   103481345
      CurrentDate     =   41623
   End
   Begin MSComCtl2.DTPicker dtpTermin 
      Height          =   495
      Left            =   1680
      TabIndex        =   11
      Top             =   3360
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
      _Version        =   393216
      Format          =   103481345
      CurrentDate     =   41623
   End
   Begin VB.TextBox txtNoInvoice 
      Height          =   495
      Left            =   1680
      MaxLength       =   15
      TabIndex        =   5
      Top             =   1200
      Width           =   1800
   End
   Begin VB.CommandButton btnCetak 
      Caption         =   "&Cetak"
      Height          =   495
      Left            =   2640
      TabIndex        =   4
      Top             =   4200
      Width           =   1215
   End
   Begin VB.CommandButton btnBatal 
      Caption         =   "&Batal"
      Height          =   495
      Left            =   1440
      TabIndex        =   3
      Top             =   4200
      Width           =   1215
   End
   Begin VB.CommandButton btnTambah 
      Caption         =   "&Tambah"
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   4200
      Width           =   1215
   End
   Begin VB.ComboBox cmbNoPO 
      Height          =   315
      Left            =   1680
      TabIndex        =   0
      Top             =   1920
      Width           =   2295
   End
   Begin MSDataGridLib.DataGrid dtgTmpInvoice 
      Bindings        =   "FrmInvoice.frx":0000
      Height          =   2655
      Left            =   4320
      TabIndex        =   1
      Top             =   1200
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   4683
      _Version        =   393216
      AllowArrows     =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      TabAction       =   2
      WrapCellPointer =   -1  'True
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
      ColumnCount     =   10
      BeginProperty Column00 
         DataField       =   "No_invoice"
         Caption         =   "No_invoice"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "No_po"
         Caption         =   "No_po"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "Tanggal_invoice"
         Caption         =   "Tanggal_invoice"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "Termin"
         Caption         =   "Termin"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "Kode_pelanggan"
         Caption         =   "Kode_pelanggan"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "Kode_barang"
         Caption         =   "Kode_barang"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "Harga_satuan"
         Caption         =   "Harga_satuan"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column07 
         DataField       =   "Kuantitas"
         Caption         =   "Kuantitas"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column08 
         DataField       =   "Total"
         Caption         =   "Total"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column09 
         DataField       =   "Keterangan"
         Caption         =   "Keterangan"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   1154,835
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1244,976
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1275,024
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1080
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1365,165
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1065,26
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1725,165
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   915,024
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   1739,906
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc adoTmpInvoice 
      Height          =   495
      Left            =   10680
      Top             =   3360
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   873
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Penjualan Tunai\penjualan_tunai.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Penjualan Tunai\penjualan_tunai.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from tmp_invoice"
      Caption         =   "Ado Tmp Invoice"
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
   Begin VB.Label Label1 
      Caption         =   "No Invoice"
      Height          =   495
      Left            =   240
      TabIndex        =   10
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Termin"
      Height          =   495
      Left            =   240
      TabIndex        =   9
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Tanggal Invoice"
      Height          =   495
      Left            =   240
      TabIndex        =   8
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "No PO"
      Height          =   495
      Left            =   240
      TabIndex        =   7
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      Caption         =   "Invoice"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   6
      Top             =   240
      Width           =   13095
   End
End
Attribute VB_Name = "FrmInvoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub bersih()
    ' Membersihkan TextBox
    For Each X In Me
        If TypeOf X Is TextBox Then
            X.Text = ""
        End If
    Next
End Sub

Sub txtHidup()
    ' Mengaktifkan Item
    For Each X In Me
        If TypeOf X Is TextBox Then
            X.Enabled = True
        End If
    Next
    txtNoInvoice.Enabled = False
    cmbNoPO.Enabled = True
    dtpTanggalInvoice.Enabled = True
    dtpTermin.Enabled = True
End Sub

Sub txtMati()
    ' Menonaktifkan item
    For Each X In Me
        If TypeOf X Is TextBox Then
            X.Enabled = False
        End If
    Next
    cmbNoPO.Enabled = False
    dtpTanggalInvoice.Enabled = False
    dtpTermin.Enabled = False
End Sub

Sub noInvoiceOtomatis()
    On Error Resume Next
    Dim nomor As String
    koneksi
    qry = "select No_invoice from invoice order by No_invoice DESC"
    Set rsInvoice = kon.Execute(qry, , adCmdText)
    With rsInvoice
        If .RecordCount = 0 Then
            nomor = "NV" + Format(Date, "YYMMDD") + "001"
        Else
            If Mid(!No_invoice, 3, 6) <> Format(Date, "YYMMDD") Then
                nomor = "NV" + Format(Date, "YYMMDD") + "001"
            Else
                hitung = Right(!No_invoice, 3) + 1
                nomor = "NV" + Format(Date, "YYMMDD") + Right("000" & hitung, 3)
            End If
        End If
    End With
    txtNoInvoice.Text = nomor
    kon.Close
    On Error GoTo 0
End Sub

Sub ambilDataNoPO()
    On Error Resume Next
    cmbNoPO.Clear
    koneksi
    rsTransaksi.Open "select No_po from transaksi group by No_po having count(*) >= 1", kon
    Do While Not rsTransaksi.EOF
        cmbNoPO.AddItem rsTransaksi!No_po
        rsTransaksi.MoveNext
    Loop
    kon.Close
    On Error GoTo 0
End Sub

Sub hapusTmpInvoice()
    On Error Resume Next
    koneksi
    qry = "delete from tmp_invoice"
    Set rsTmpInvoice = kon.Execute(qry, , adCmdText)
    kon.Close
    adoTmpInvoice.Refresh
    On Error GoTo 0
End Sub

Sub salinTmpInvoiceKeInvoice()
    koneksi
    qry = "insert into invoice select * from tmp_invoice"
    Set rsInvoice = kon.Execute(qry, , adCmdText)
    kon.Close
End Sub

Sub tampilGridTransaksiNoInvoice()
    koneksi
    qry = "insert into tmp_invoice select '" & txtNoInvoice.Text & "' as No_invoice, No_po, Kode_pelanggan, Kode_barang, Harga_satuan, Kuantitas, Total, Keterangan from transaksi where No_po='" & cmbNoPO.Text & "'"
    Set rsTmpInvoice = kon.Execute(qry, , adCmdText)
    kon.Close
    adoTmpInvoice.Refresh
End Sub

Sub tampilGridTransaksiTanggalInvoice()
    On Error Resume Next
    koneksi
    qry = "insert into tmp_invoice select '" & txtNoInvoice.Text & "' as No_invoice, No_po, '" & dtpTanggalInvoice.Value & "' as Tanggal_invoice, Kode_pelanggan, Kode_barang, Harga_satuan, Kuantitas, Total, Keterangan from transaksi where No_po='" & cmbNoPO.Text & "'"
    Set rsTmpInvoice = kon.Execute(qry, , adCmdText)
    kon.Close
    adoTmpInvoice.Refresh
    On Error GoTo 0
End Sub

Sub tampilGridTransaksiTermin()
    On Error Resume Next
    koneksi
    qry = "insert into tmp_invoice select '" & txtNoInvoice.Text & "' as No_invoice, No_po, '" & dtpTanggalInvoice.Value & "' as Tanggal_invoice, '" & dtpTermin.Value & "' as Termin, Kode_pelanggan, Kode_barang, Harga_satuan, Kuantitas, Total, Keterangan from transaksi where No_po='" & cmbNoPO.Text & "'"
    Set rsTmpInvoice = kon.Execute(qry, , adCmdText)
    kon.Close
    adoTmpInvoice.Refresh
    On Error GoTo 0
End Sub

Sub cetakInvoice()
    On Error Resume Next
    With crInvoice
        .ReportFileName = App.Path & "\Laporan\lapInvoice.rpt"
        .SelectionFormula = ""
        .ParameterFields(0) = "formulaNoInvoice;" & txtNoInvoice.Text & ";True"
        .ParameterFields(1) = "formulaKodePelanggan;" & dtgTmpInvoice.Columns(4).Value & ";True"
        .ParameterFields(2) = "formulaTanggalInvoice;" & dtgTmpInvoice.Columns(2).Value & ";True"
        .ParameterFields(3) = "formulaTermin;" & dtgTmpInvoice.Columns(3).Value & ";True"
        .RetrieveDataFiles
        .WindowState = crptMaximized
        .Action = 1
    End With
    On Error GoTo 0
End Sub

Private Sub btnBatal_Click()
    bersih
    txtMati
    btnBatal.Enabled = False
    btnTambah.Enabled = True
    btnCetak.Enabled = False
End Sub

Private Sub btnCetak_Click()
    salinTmpInvoiceKeInvoice
    cetakInvoice
    hapusTmpInvoice
    bersih
    txtMati
    btnBatal.Enabled = False
    btnTambah.Enabled = True
    btnCetak.Enabled = False
End Sub

Private Sub btnKeluar_Click()
    Unload Me
    hapusTmpInvoice
End Sub

Private Sub btnTambah_Click()
    txtHidup
    bersih
    btnTambah.Enabled = True
    btnBatal.Enabled = True
    btnCetak.Enabled = True
    noInvoiceOtomatis
    ambilDataNoPO
End Sub

Private Sub cmbNoPO_Click()
    hapusTmpInvoice
    tampilGridTransaksiNoInvoice
End Sub

Private Sub dtpTanggalInvoice_Click()
    hapusTmpInvoice
    tampilGridTransaksiTanggalInvoice
End Sub

Private Sub dtpTermin_Click()
    hapusTmpInvoice
    tampilGridTransaksiTermin
End Sub

Private Sub Form_Activate()
    txtMati
    btnTambah.SetFocus
    txtNoInvoice.Enabled = False
    btnBatal.Enabled = False
    btnCetak.Enabled = False
End Sub
