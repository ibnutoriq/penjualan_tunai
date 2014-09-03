Attribute VB_Name = "MdlPenjualanTunai"
Public kon As New ADODB.Connection
Public rsBarang As New ADODB.Recordset
Public rsInvoice As New ADODB.Recordset
Public rsJurnal As New ADODB.Recordset
Public rsPengguna As New ADODB.Recordset
Public rsPelanggan As New ADODB.Recordset
Public rsPenerimaanPembayaran As New ADODB.Recordset
Public rsPerkiraan As New ADODB.Recordset
Public rsSuratJalan As New ADODB.Recordset
Public rsTransaksi As New ADODB.Recordset
Public rsTmpTransaksi As New ADODB.Recordset
Public rsTmpInvoice As New ADODB.Recordset
Public qry As String

Sub koneksi()
    Set rsBarang = New ADODB.Recordset
    Set rsInvoice = New ADODB.Recordset
    Set rsJurnal = New ADODB.Recordset
    Set rsPengguna = New ADODB.Recordset
    Set rsPelanggan = New ADODB.Recordset
    Set rsPenerimaanPembayaran = New ADODB.Recordset
    Set rsPerkiraan = New ADODB.Recordset
    Set rsSuratJalan = New ADODB.Recordset
    Set rsTransaksi = New ADODB.Recordset
    Set rsTmpTransaksi = New ADODB.Recordset
    Set rsTmpInvoice = New ADODB.Recordset
    kon.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Penjualan Tunai\penjualan_tunai.mdb;Persist Security Info=False"
    kon.Open
End Sub
