Imports System.Data.SqlClient
'Imports Microsoft.Office.Interop
'Imports Microsoft.Office.Interop.Excel
Imports Excel = Microsoft.Office.Interop.Excel
Imports Microsoft.VisualBasic
'Imports System.Data.OleDb
Imports System.Globalization

'santo
Imports System.IO
Imports System.Reflection
Imports NiceLabel.SDK

''crys Report
'Imports CrystalDecisions.CrystalReports.Engine
'Imports CrystalDecisions.Shared

Imports System.Data
Imports System.IO.Directory
Imports Microsoft.Office.Interop

'Imports System.Data
Imports System.Data.OleDb

'export excel
Imports System.Linq
Imports Microsoft.Office.Core
Imports ExcelAutoFormat = Microsoft.Office.Interop.Excel.XlRangeAutoFormat
Imports System.Xml.XPath
Imports System.Xml
'Imports System.Math

Public Class Main
    Public Shared koneksi As SqlConnection
    Public Shared database As String
    Public role As String
    Public q As Object
    Public PackQuantity As Long
    Dim visualCheck As Integer
    Dim actualQty As Integer
    Dim targetQty As Integer
    Dim finishVision As Boolean
    Dim scanComponents As Boolean
    Dim tempPPnumber As String
    Dim tempComponentNo As String
    Private wk As String
    Dim excel As String
    Dim oleCon As OleDbConnection
    Dim adapteroledb As OleDbDataAdapter
    Dim dtoledb As DataSet
    Dim hasnotbeenfound As Boolean
    Dim hasnotbeenindatabase As Boolean
    Dim doyouwanttoreprintit As Boolean
    Dim PPPackingdoyouwanttoreprintit As Boolean
    Dim CustomerCountryNotFound As Boolean
    Dim ReqDelvdt As Boolean
    Dim cscinput As Boolean
    Dim atsinput As Boolean
    Dim qtyperboxinput1 As Boolean
    Dim qtyperboxinput2 As Boolean
    Dim duplicateSQOO As Boolean
    Dim duplicateCustomer As Boolean
    Dim variable_Q As Integer
    Dim hitung As Integer = 0
    Dim qtyboxinput As Integer = 0
    Dim missingMadein As Integer = 0
    Dim sekaliOF As Integer = 0
    Dim seconds As Integer = 0
    Dim minutes As Integer = 0
    Dim hours As Integer = 0
    Dim countRuby As Integer = 0

    'santo
    Dim label_printer As ILabel
    Dim label1_printer As ILabel
    Dim label2_printer As ILabel
    Dim label3_printer As ILabel
    Dim label4_printer As ILabel
    Dim label5_printer As ILabel
    Dim label6_printer As ILabel

    'Fuji
    Dim label_rotary_printer As ILabel
    Dim label_front_long_printer As ILabel
    Dim label_front_short_printer As ILabel
    Dim label_side_printer As ILabel
    Dim label_carton_printer As ILabel
    Dim label_out_side_printer As ILabel

    Dim printers As IList(Of IPrinter)
    Dim selected_Printer As IPrinter

    'Ruby
    Dim label_performance_small_ruby As ILabel
    Dim label_performance_big_ruby As ILabel
    Dim label_packaging_ruby As ILabel
    Dim label_outside_ruby As ILabel


    Dim seq_SerialNumber As Integer

    Dim var_PrintLabel_klik As Boolean = False

    Public Shared Sub koneksi_db()
        Try
            'database = "Data Source=DESKTOP-4PHNBDD;initial catalog=MES_DB;integrated security=true"
            'database = "Data Source=10.155.128.185;initial catalog=MES_DB;Persist Security Info=True;User ID =HA;Password=HA@123"
            database = "Data Source=10.155.128.71;
            initial catalog=SGRAC_MES;
            Persist Security Info=True;
            User ID=SGRAC;
            Password=SGRAC@123;
            Max Pool Size=5000;
            Pooling=True"
            koneksi = New SqlConnection(database)
            If koneksi.State = ConnectionState.Closed Then koneksi.Open() Else koneksi.Close()
        Catch ex As Exception
            'MsgBox("Please Contact IT Team. This is Database Problem -> " + ex.Message)
            Log_form.Log_text.Text = Log_form.Log_text.Text & Chr(13) & Chr(13) & "Please Contact IT Team. This is Database Problem -> " + ex.Message
        End Try
    End Sub

    Private Sub btnExit1_Click(sender As Object, e As EventArgs) Handles Command39.Click
        On Error GoTo Err_btnExit1_Click
        Me.Close()
Exit_btnExit1_Click:
        Exit Sub
Err_btnExit1_Click:
        'MsgBox Err.Description
        Resume Exit_btnExit1_Click
    End Sub

    Private Sub btnExit2_Click(sender As Object, e As EventArgs) Handles Command192.Click
        On Error GoTo Err_btnExit2_Click
        Me.Close()
Exit_btnExit2_Click:
        Exit Sub
Err_btnExit2_Click:
        'MsgBox Err.Description
        Resume Exit_btnExit2_Click
    End Sub
    Private Sub cek_newVersion()
        Dim NewVer As String
        Dim path1 As String = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile)
        Dim sourceLocation = path1 & "\Box\SANTO\_build\version"
        Dim xNewLocataion = Application.StartupPath()
        NewVer = My.Computer.FileSystem.ReadAllText(sourceLocation & "\version.txt")

        Try

            If Application.ProductVersion <> NewVer Then

                Dim result_quest As DialogResult = MessageBox.Show("You are using old version of application !" & Environment.NewLine & "Do you want to Update?",
            "Application Version",
             MessageBoxButtons.YesNo,
             MessageBoxIcon.Question)

                If result_quest = DialogResult.Yes Then
                    'close app


                    CopyDirectory(path1 & "\Box\SANTO\_build\updater\", xNewLocataion)
                    'System.Threading.Thread.Sleep(5000)
                    'MessageBox.Show("Updater terdownload")
                    Dim proc As New System.Diagnostics.Process()
                    proc = Process.Start("updater.exe", "")
                    Me.Close()

                End If
            End If
        Catch ex As Exception

        End Try

        'LABEL COPY
        Try
            Dim a = path1 & "\Box\SANTO\_build\label"
            'If Not System.IO.File.Exists("Label\tes.nlbl") Then File.Copy(a & "\tes.nlbl", "Label\tes.nlbl", True)
            'Fuji
            If Not System.IO.File.Exists("Label\Fuji Carton Label.nlbl") Then File.Copy(a & "\Fuji Carton Label.nlbl", "Label\Fuji Carton Label.nlbl", True)
            If Not System.IO.File.Exists("Label\Fuji Front Label long.nlbl") Then File.Copy(a & "\Fuji Front Label long.nlbl", "Label\Fuji Front Label long.nlbl", True)
            If Not System.IO.File.Exists("Label\Fuji Front Label short.nlbl") Then File.Copy(a & "\Fuji Front Label short.nlbl", "Label\Fuji Front Label short.nlbl", True)
            If Not System.IO.File.Exists("Label\Fuji Outside Grouping Label.nlbl") Then File.Copy(a & "\Fuji Outside Grouping Label.nlbl", "Label\Fuji Outside Grouping Label.nlbl", True)
            If Not System.IO.File.Exists("Label\Fuji Product Label side label.nlbl") Then File.Copy(a & "\Fuji Product Label side label.nlbl", "Label\Fuji Product Label side label.nlbl", True)
            If Not System.IO.File.Exists("Label\Fuji Rotary Handle Label.nlbl") Then File.Copy(a & "\Fuji Rotary Handle Label.nlbl", "Label\Fuji Rotary Handle Label.nlbl", True)

            'Ruby
            If Not System.IO.File.Exists("Label\Ruby Performace Label Small.nlbl") Then File.Copy(a & "\Ruby Performace Label Small.nlbl", "Label\Ruby Performace Label Small.nlbl", True)
            If Not System.IO.File.Exists("Label\Ruby Performace Label Big.nlbl") Then File.Copy(a & "\Ruby Performace Label Big.nlbl", "Label\Ruby Performace Label Big.nlbl", True)
            If Not System.IO.File.Exists("Label\Ruby Packaging Label.nlbl") Then File.Copy(a & "\Ruby Packaging Label.nlbl", "Label\Ruby Packaging Label.nlbl", True)
            If Not System.IO.File.Exists("Label\Ruby Out Side Label.nlbl") Then File.Copy(a & "\Ruby Out Side Label.nlbl", "Label\Ruby Out Side Label.nlbl", True)


        Catch ex As Exception

        End Try
    End Sub

    Public Sub CopyDirectory(ByVal sourcePath As String, ByVal destinationPath As String)

        Dim sourceDirectoryInfo As New System.IO.DirectoryInfo(sourcePath)

        ' If the destination folder don't exist then create it
        If Not System.IO.Directory.Exists(destinationPath) Then
            System.IO.Directory.CreateDirectory(destinationPath)
        End If

        Dim fileSystemInfo As System.IO.FileSystemInfo
        For Each fileSystemInfo In sourceDirectoryInfo.GetFileSystemInfos

            Dim destinationFileName As String =
                System.IO.Path.Combine(destinationPath, fileSystemInfo.Name)

            ' Now check whether its a file or a folder and take action accordingly
            If TypeOf fileSystemInfo Is System.IO.FileInfo Then
                System.IO.File.Copy(fileSystemInfo.FullName, destinationFileName, True)
            Else
                ' Recursively call the mothod to copy all the neste folders
                CopyDirectory(fileSystemInfo.FullName, destinationFileName)
            End If
        Next

    End Sub

    Private Sub loading(ByVal a)
        Loading_Form.Loading_Progress.Value = a
        Application.DoEvents()
    End Sub

    Private Sub Main_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Dim dtset As New DataSet
        Dim sqlMaxRecord As String = "SELECT max([RecordRuby]) as maxRuby FROM [printingRecord] WHERE [date] = '" & DateTime.Now.ToString("yyyy-MM-dd") & "'"
        Dim adapt = New SqlDataAdapter(sqlMaxRecord, Main.koneksi)
        adapt.Fill(dtset)
        If IsDBNull(dtset.Tables(0).Rows(0).Item("maxRuby")) Then
            countRuby = 1
        Else
            countRuby = dtset.Tables(0).Rows(0).Item("maxRuby") + 1
        End If

        LoginForm.Hide()
        Loading_Form.Show()
        Me.Hide()

        loading(10)

        baca_chkbox_last()

        testPrintQty.SelectedItem = "2"

        cek_newVersion()

        Me.Text = "MES Local Adaptation Center v " & Application.ProductVersion

        'TODO: This line of code loads data into the 'SGRAC_MESDataSet.customerDatabase' table. You can move, or remove it, as needed.
        Me.CustomerDatabaseTableAdapter.Fill(Me.SGRAC_MESDataSet.customerDatabase)
        'TODO: This line of code loads data into the 'SGRAC_MESDataSet.madeINtext' table. You can move, or remove it, as needed.
        Me.MadeINtextTableAdapter.Fill(Me.SGRAC_MESDataSet.madeINtext)
        'TODO: This line of code loads data into the 'SGRAC_MESDataSet.Users' table. You can move, or remove it, as needed.
        Me.UsersTableAdapter.Fill(Me.SGRAC_MESDataSet.Users)
        'TODO: This line of code loads data into the 'SGRAC_MESDataSet.Workstations' table. You can move, or remove it, as needed.
        Me.WorkstationsTableAdapter.Fill(Me.SGRAC_MESDataSet.Workstations)

        'Call koneksi_db()
        'Try
        '    Dim sc_workstations As New SqlCommand("select * from Workstations order by wkName asc", koneksi)
        '    Dim sc_technician As New SqlCommand("select * from Users order by Technician asc", koneksi)
        '    Dim adapter_ws As New SqlDataAdapter(sc_workstations)
        '    Dim adapter_tech As New SqlDataAdapter(sc_technician)
        '    Dim table_tech As New DataTable()
        '    Dim table_ws As New DataTable()

        '    adapter_ws.Fill(table_ws)
        '    adapter_tech.Fill(table_tech)

        'ComboBox1.DataSource = table_ws
        'ComboBox11.DataSource = table_ws
        'ComboBox4.DataSource = table_ws

        'ComboBox1.DisplayMember = "wkname"
        'ComboBox11.DisplayMember = "wkname"
        'ComboBox4.DisplayMember = "wkname"

        'ComboBox2.DataSource = table_tech
        'ComboBox5.DataSource = table_tech
        'ComboBox10.DataSource = table_tech

        'ComboBox2.DisplayMember = "Technician"
        'ComboBox5.DisplayMember = "Technician"
        'ComboBox10.DisplayMember = "Technician"

        'Catch ex As Exception
        '    MessageBox.Show(ex.Message)
        'End Try

        'workstation.Text = My.Settings.WorkStationBind

        loading(20)

        'santo
        InitializePrintEngine()

        Dim appPath As String = Application.StartupPath()
        'read NiceLabel file
        'label_printer = PrintEngineFactory.PrintEngine.OpenLabel("SimpleSample1.nlbl")

        label_printer = PrintEngineFactory.PrintEngine.OpenLabel(appPath & "\Label\" & "generic.nlbl")
        label1_printer = PrintEngineFactory.PrintEngine.OpenLabel(appPath & "\Label\" & "adaptation.nlbl")
        label2_printer = PrintEngineFactory.PrintEngine.OpenLabel(appPath & "\Label\" & "testReport.nlbl")
        label3_printer = PrintEngineFactory.PrintEngine.OpenLabel(appPath & "\Label\" & "COC.nlbl")
        loading(30)
        label4_printer = PrintEngineFactory.PrintEngine.OpenLabel(appPath & "\Label\" & "Loose2.nlbl")
        label5_printer = PrintEngineFactory.PrintEngine.OpenLabel(appPath & "\Label\" & "Components.nlbl")
        label6_printer = PrintEngineFactory.PrintEngine.OpenLabel(appPath & "\Label\" & "Loose2.nlbl")

        loading(40)

        'Fuji
        label_side_printer = PrintEngineFactory.PrintEngine.OpenLabel(appPath & "\Label\" & "Fuji Product Label side label.nlbl")
        label_rotary_printer = PrintEngineFactory.PrintEngine.OpenLabel(appPath & "\Label\" & "Fuji Rotary Handle Label.nlbl")
        label_front_long_printer = PrintEngineFactory.PrintEngine.OpenLabel(appPath & "\Label\" & "Fuji Front Label long.nlbl")
        label_front_short_printer = PrintEngineFactory.PrintEngine.OpenLabel(appPath & "\Label\" & "Fuji Front Label short.nlbl")
        label_carton_printer = PrintEngineFactory.PrintEngine.OpenLabel(appPath & "\Label\" & "Fuji Carton Label.nlbl")
        label_out_side_printer = PrintEngineFactory.PrintEngine.OpenLabel(appPath & "\Label\" & "Fuji Outside Grouping Label.nlbl")

        'Ruby
        label_performance_small_ruby = PrintEngineFactory.PrintEngine.OpenLabel(appPath & "\Label\" & "Ruby Performace Label Small.nlbl")
        label_performance_big_ruby = PrintEngineFactory.PrintEngine.OpenLabel(appPath & "\Label\" & "Ruby Performace Label Big.nlbl")
        label_packaging_ruby = PrintEngineFactory.PrintEngine.OpenLabel(appPath & "\Label\" & "Ruby Packaging Label.nlbl")
        label_outside_ruby = PrintEngineFactory.PrintEngine.OpenLabel(appPath & "\Label\" & "Ruby Out Side Label.nlbl")

        loading(50)

        printers = PrintEngineFactory.PrintEngine.Printers
        'filling up list of comboBox
        For i = 0 To printers.Count - 1
            listprinter.Items.Add(printers.Item(i).Name)
            listprinter1.Items.Add(printers.Item(i).Name)
            listprinter2.Items.Add(printers.Item(i).Name)
            listprinter3.Items.Add(printers.Item(i).Name)

            listAvalablePrinter.Items.Add(printers.Item(i).Name)

            'fuji
            cbx_Carton.Items.Add(printers.Item(i).Name)
            cbx_outside.Items.Add(printers.Item(i).Name)
            cbx_Rotary.Items.Add(printers.Item(i).Name)
            cbx_front.Items.Add(printers.Item(i).Name)
            cbx_fuji_side_label.Items.Add(printers.Item(i).Name)

            'Ruby
            cbxPerfomaceRuby.Items.Add(printers.Item(i).Name)
            cbxPackagingRuby.Items.Add(printers.Item(i).Name)
            cbxOutsideRuby.Items.Add(printers.Item(i).Name)

        Next

        loading(60)
        'show list printer
        If role = "admin" Then
            listprinter.Enabled = True
        Else
            listprinter.Enabled = False
        End If

        If role = "admin" Then
            listprinter1.Enabled = True
        Else
            listprinter1.Enabled = False
        End If

        If role = "admin" Then
            listprinter2.Enabled = True
        Else
            listprinter2.Enabled = False
        End If

        If role = "admin" Then
            listprinter3.Enabled = True
        Else
            listprinter3.Enabled = False
        End If


        'initial selected Printer label 1
        selected_Printer = printers.Item(0)
        'label_printer.PrintSettings.PrinterName = selected_Printer.Name
        'label1_printer.PrintSettings.PrinterName = selected_Printer.Name
        'label2_printer.PrintSettings.PrinterName = selected_Printer.Name
        'label3_printer.PrintSettings.PrinterName = selected_Printer.Name

        listprinter.SelectedItem = selected_Printer.Name
        listprinter1.SelectedItem = selected_Printer.Name
        listprinter2.SelectedItem = selected_Printer.Name
        listprinter3.SelectedItem = selected_Printer.Name

        technicianName.Text = ""
        user.Text = ""
        testUser.Text = ""

        workstation.Text = ""
        workstation2.Text = ""
        workstation3.Text = ""


        madeInEnglish.Text = ""

        If role = "admin" Then
            AccessAdmin()
        Else
            AccessOperator()
        End If

        finishVision = False
        scanComponents = False
        CounterItems.Text = 1

        loading(70)

        'show check box
        en_print.Checked = My.Settings.checkbox

        If en_print.Checked = True Then
            Manual_Print_Product.Visible = True
            Print_Product_x.Visible = True
        Else
            Manual_Print_Product.Visible = False
            Print_Product_x.Visible = False
        End If

        DGV_Quality()

        DGV_MasterPlantCode()

        loading(90)

        Report_Tab.SelectedIndex = 3
        Report_Tab.SelectedIndex = 2
        Report_Tab.SelectedIndex = 1
        Report_Tab.SelectedIndex = 0

        workstationFuji.Text = ""



        loading(100)

        Loading_Form.Close()
        Me.Show()

    End Sub

    Sub resetValmsgbox()
        hasnotbeenfound = 0
        hasnotbeenindatabase = 0
        doyouwanttoreprintit = 0
        PPPackingdoyouwanttoreprintit = 0
        atsinput = 0
        cscinput = 0
        qtyperboxinput1 = 0
        qtyperboxinput2 = 0
        duplicateSQOO = 0
        duplicateCustomer = 0
    End Sub

    Sub AccessAdmin() 'Can Edit, Import to DB
        upload_open_orders.Visible = True
        Command109.Visible = True
        Command121.Visible = True
        Command147.Visible = True
        Command111.Visible = True
        Command113.Visible = True
        Command279.Visible = True
        Command577.Visible = True
        Command575.Visible = True
        listprinter.Enabled = True
        listprinter1.Enabled = True
        listprinter2.Enabled = True
        listprinter3.Enabled = True
        DeleteDuplicateBOM.Visible = True
        'Manual_Print_Product.Visible = True
        'Print_Product_x.Visible = True
        en_print.Visible = True
        'My.Settings.Reset()

        Button5.Visible = True
        Active_Quality_issue.Enabled = True
        btn_add.Enabled = True
        btn_delete.Enabled = True
        Button5.Enabled = True
        btn_add.Enabled = True
        btn_delete.Enabled = True
        Button7.Enabled = True
        Button6.Enabled = True
        btn_export_QI_Trace.Enabled = True

        deleteRow.Enabled = True

        'fuji 
        cbx_fuji_side_label.Enabled = True
        cbx_Rotary.Enabled = True
        cbx_outside.Enabled = True
        cbx_front.Enabled = True
        cbx_Carton.Enabled = True

        'tools
        btn_log.Visible = True

        btn_NewComponents.Visible = True
        Button19.Visible = True

    End Sub

    Sub AccessOperator() 'Just View and cannot Edit
        upload_open_orders.Visible = False
        Command109.Visible = False
        Command121.Visible = False
        Command147.Visible = False
        Command111.Visible = False
        Command113.Visible = False
        Command279.Visible = False
        Command577.Visible = False
        Command575.Visible = False
        listprinter.Enabled = False
        listprinter1.Enabled = False
        listprinter2.Enabled = False
        listprinter3.Enabled = False
        DeleteDuplicateBOM.Visible = False
        'Manual_Print_Product.Visible = False
        'Print_Product_x.Visible = False
        en_print.Visible = False
        'Button5.Visible = False
        Active_Quality_issue.Enabled = False
        btn_add.Enabled = False
        btn_delete.Enabled = False
        Button5.Enabled = False
        btn_add.Enabled = False
        btn_delete.Enabled = False
        Button7.Enabled = False
        Button6.Enabled = False
        btn_export_QI_Trace.Enabled = False
        deleteRow.Enabled = False

        'fuji
        cbx_fuji_side_label.Enabled = False
        cbx_Rotary.Enabled = False
        cbx_outside.Enabled = False
        cbx_front.Enabled = False
        cbx_Carton.Enabled = False

        'tools
        btn_log.Visible = False

        btn_NewComponents.Visible = False
        Button19.Visible = False

    End Sub

    'santo
    'initialitation of printer engine
    Private Sub InitializePrintEngine()
        Try
            Dim sdkFilesPath As String = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "..\\..\\..\\SDKFiles")
            'MsgBox(sdkFilesPath)

            If My.Computer.FileSystem.DirectoryExists(sdkFilesPath) Then
                PrintEngineFactory.SDKFilesPath = sdkFilesPath
            End If
            PrintEngineFactory.PrintEngine.Initialize()
            'MsgBox("test")
        Catch ex As Exception
            MsgBox("Initialization of the SDK failed." + Environment.NewLine + Environment.NewLine + ex.ToString())
            Me.Close()
        End Try

    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles open_orders_table.Click
        Form1.open_form("select [ID]
        ,[Order]
        ,[Material]
        ,[Descr]
        ,[Reqmts qty]
        ,[Stor  loc]
        ,[Status]
        ,[PeggedReqt] from openOrders")
        Form1.Show()
    End Sub

    Private Sub OpenComponentslist_Click(sender As Object, e As EventArgs) Handles OpenComponentslist.Click
        Form1.open_form("select [ID]
        ,[Material]
        ,[Code] from Componentslist")
        Form1.Show()
    End Sub

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Command108.Click
        Form1.open_form("select [ID]
      ,[Material]
      ,[Description]
      ,[DescriptionFR]
      ,[DescriptionES]
      ,[DescriptionZH]
      ,[DescriptionRU]
      ,[Range]
      ,[EAN13]
      ,[logo1]
      ,[logo2]
      ,[logo3]
      ,[logo4]
      ,[logo5]
      ,[logo6]
      ,[productImage] from NSXMasterdata")
        Form1.Show()
    End Sub

    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Command37.Click
        '  Form1.open_form("select [autonumber]
        ',[PP]
        ',[date]
        ',[time]
        ',[user]
        ',[From]
        ',[To]
        ',[QRCodeFuji]
        ',[Data] from printingRecord")

        Form1.open_form("select * from printingRecord")
        Form1.Show()
    End Sub

    Private Sub Button27_Click(sender As Object, e As EventArgs) Handles Command3.Click
        Form1.open_form("select [ID]
      ,[Created on]
      ,[Entered by]
      ,[Type]
      ,[Order]
      ,[Start time]
      ,[Basic fin]
      ,[Committed]
      ,[SchedStart]
      ,[Finish tme]
      ,[PrCtr]
      ,[Material]
      ,[Scheduled finish]
      ,[Scheduled start]
      ,[Basic start date]
      ,[Basic finish date]
      ,[Item quantity]
      ,[OUM]
      ,[Material description]
      ,[Description]
      ,[CDI]
      ,[Req dlv dt]
      ,[Purchase order number]
      ,[POitem]
      ,[GI time]
      ,[Stag  time]
      ,[GI date]
      ,[Mat av dt]
      ,[Confirmed qty]
      ,[SU]
      ,[Customer]
      ,[City]
      ,[Name 2]
      ,[Name 1]
      ,[Item]
      ,[SO no]
      ,[Customer Material Number]
      ,[Status]
      ,[Stat]
      ,[range]
      ,[OTD status]
      ,[OTD status2]
      ,[Entity] from PPList")
        Form1.Show()
    End Sub

    Private Sub Button26_Click(sender As Object, e As EventArgs) Handles Command146.Click
        Form1.open_form("select [Id]
      ,[customer code]
      ,[customer name]
      ,[country]
      ,[country short name] from customerDatabase")
        Form1.Show()
    End Sub

    Private Sub Button25_Click(sender As Object, e As EventArgs) Handles Command141.Click
        Form1.open_form("select [ID]
      ,[PP]
      ,[print date]
      ,[print time]
      ,[User] from printingRecordPacking")
        Form1.Show()
    End Sub

    Private Sub Button36_Click(sender As Object, e As EventArgs) Handles Command34.Click
        Form1.open_form("select [ID]
      ,[Technician]
      ,[Short name] from Users")
        Form1.Show()
    End Sub

    Private Sub Button35_Click(sender As Object, e As EventArgs) Handles Command43.Click
        Form1.open_form("select [range start]
      ,[range end]
      ,[label]
      ,[special]
      ,[maxBOX]
      ,[maxPallet]
      ,[Default Made-In]
      ,[ProductImage] from labelSelectionTable")
        Form1.Show()
    End Sub

    Private Sub Button34_Click(sender As Object, e As EventArgs) Handles Command264.Click
        '  Form1.open_form("select [ID]
        ',[workstation]
        ',[report type]
        ',[printer]
        ',[printer type]
        ',[dpi] from printerTable")
        '  Form1.Show()
        Printer_Setting.Show()
    End Sub

    Private Sub Button48_Click(sender As Object, e As EventArgs) Handles Command204.Click
        Form1.open_form("select [ID]
      ,[PP]
      ,[date]
      ,[time]
      ,[User] from printingRecordCOC")
        Form1.Show()
    End Sub

    Private Sub Button47_Click(sender As Object, e As EventArgs) Handles Command206.Click
        Form1.open_form("select [ID]
      ,[PP]
      ,[date]
      ,[time]
      ,[User] from printingRecordTest")
        Form1.Show()
    End Sub

    Private Sub Button46_Click(sender As Object, e As EventArgs) Handles Command278.Click
        Form1.open_form("select [ID]
      ,[BOM number]
      ,[description]
      ,[range]
      ,[category]
      ,[material]
      ,[type]
      ,[poles] from BOM")
        Form1.Show()
    End Sub
    Dim selectWorkstation As Integer
    Private Sub workstation_event()

        Dim sql As String
        Dim ds As New DataSet
        Dim sql2 As String
        Dim ds2 As New DataSet
        Dim sql3 As String
        Dim ds3 As New DataSet
        Dim sql4 As String
        Dim ds4 As New DataSet

        Dim sql_fuji(10), sql_ruby(5) As String
        Dim ds1_fuji, ds2_fuji, ds3_fuji, ds4_fuji, ds5_fuji As New DataSet
        Dim ds1_ruby, ds2_ruby, ds3_ruby As New DataSet

        Dim selectedPrinterPackaging As String = "Microsoft Print to PDF"
        Dim selectedPrinterProduct As String = "Microsoft Print to PDF"
        Dim selectedPrinterCOC As String = "Microsoft Print to PDF"
        Dim selectedPrinterTest As String = "Microsoft Print to PDF"
        'Fuji
        Dim selectedPrinterFujiSIdeLabel As String = "Microsoft Print to PDF"
        Dim selectedPrinterFujifrontLongLabel As String = "Microsoft Print to PDF"
        Dim selectedPrinterFujiRotaryLabel As String = "Microsoft Print to PDF"
        Dim selectedPrinterFujiCartonLabel As String = "Microsoft Print to PDF"
        Dim selectedPrinterFujiOutSIdeLabel As String = "Microsoft Print to PDF"
        Dim selectedPrinterFujifrontLabel As String = "Microsoft Print to PDF"

        'Ruby
        Dim selectedPrinterRubyPerformance As String = "Microsoft Print to PDF"
        Dim selectedPrinterRubyPackaging As String = "Microsoft Print to PDF"
        Dim selectedPrinterRubyOutSide As String = "Microsoft Print to PDF"



        'label_performance_small_ruby.PrintSettings.PrinterName = cbxPerfomaceRuby.Text
        'label_performance_big_ruby.PrintSettings.PrinterName = cbxPerfomaceRuby.Text
        'label_packaging_ruby.PrintSettings.PrinterName = cbxPackagingRuby.Text
        'label_outside_ruby.PrintSettings.PrinterName = cbxOutsideRuby.Text

        Try
            sql = "SELECT [printer]  FROM [dbo].[printerTable] where [workstation]= '" &
            workstation.SelectedValue.ToString & "' AND [report type]= 'product label'"
            Dim adapter = New SqlDataAdapter(sql, Main.koneksi)
            adapter.Fill(ds)
            If ds.Tables(0).Rows.Count = 0 Then
                MsgBox("Product Label Printer not set in Database !")
                'Exit Sub
                selectWorkstation = 0
            Else
                selectedPrinterProduct = ds.Tables(0).Rows(0).Item("printer").ToString
                selectWorkstation = selectWorkstation Or 1
            End If

        Catch ex As Exception
        End Try
        Try
            'Select Packaging Label printer
            sql2 = "SELECT [printer]  FROM [dbo].[printerTable] where [workstation]= '" &
            workstation.SelectedValue.ToString & "' AND [report type]= 'packing label'"
            Dim adapter2 = New SqlDataAdapter(sql2, Main.koneksi)
            adapter2.Fill(ds2)

            If ds2.Tables(0).Rows.Count = 0 Then
                MsgBox("Packaging Label Printer not set in Database !")
                selectWorkstation = 0
            Else
                selectedPrinterPackaging = ds2.Tables(0).Rows(0).Item("printer").ToString
                selectWorkstation = selectWorkstation Or 2
            End If
        Catch ex As Exception
        End Try

        Try
            'Select COC Label printer
            sql3 = "SELECT [printer]  FROM [dbo].[printerTable] where [workstation]= '" &
            workstation.SelectedValue.ToString & "' AND [report type]= 'COC'"
            Dim adapter3 = New SqlDataAdapter(sql3, Main.koneksi)
            adapter3.Fill(ds3)

            If ds3.Tables(0).Rows.Count = 0 Then
                MsgBox("COC Label Printer not set in Database !")
                selectWorkstation = 0
            Else
                selectedPrinterCOC = ds3.Tables(0).Rows(0).Item("printer").ToString
                selectWorkstation = selectWorkstation Or 4
            End If
        Catch ex As Exception
        End Try

        Try
            'Select COC Label printer
            sql4 = "SELECT [printer]  FROM [dbo].[printerTable] where [workstation]= '" &
            workstation.SelectedValue.ToString & "' AND [report type]= 'test report'"
            Dim adapter4 = New SqlDataAdapter(sql4, Main.koneksi)
            adapter4.Fill(ds4)

            If ds4.Tables(0).Rows.Count = 0 Then
                MsgBox("Test report Label Printer not set in Database !")
                selectWorkstation = 0
            Else
                selectedPrinterTest = ds4.Tables(0).Rows(0).Item("printer").ToString
                selectWorkstation = selectWorkstation Or 8
            End If
        Catch ex As Exception
        End Try

        'Fuji side label
        Try
            sql_fuji(1) = "SELECT [printer]  FROM [dbo].[printerTable] where [workstation]= '" &
            workstation.SelectedValue.ToString & "' AND [report type]= 'side label'"
            Dim adapter_fuji_side = New SqlDataAdapter(sql_fuji(1), Main.koneksi)
            adapter_fuji_side.Fill(ds1_fuji)

            If ds1_fuji.Tables(0).Rows.Count = 0 Then
                MsgBox("Product Label Printer not set in Database !")
            Else
                selectedPrinterFujiSIdeLabel = ds1_fuji.Tables(0).Rows(0).Item("printer").ToString
            End If

        Catch ex As Exception
        End Try

        Try
            sql_fuji(2) = "SELECT [printer]  FROM [dbo].[printerTable] where [workstation]= '" &
            workstation.SelectedValue.ToString & "' AND [report type]= 'rotary label'"
            Dim adapter_fuji_rotary = New SqlDataAdapter(sql_fuji(2), Main.koneksi)
            adapter_fuji_rotary.Fill(ds2_fuji)

            If ds2_fuji.Tables(0).Rows.Count = 0 Then
                MsgBox("Fuji Rotary Label Printer not set in Database !")
            Else
                selectedPrinterFujiRotaryLabel = ds2_fuji.Tables(0).Rows(0).Item("printer").ToString
            End If

        Catch ex As Exception
        End Try
        'front
        Try
            sql_fuji(3) = "SELECT [printer]  FROM [dbo].[printerTable] where [workstation]= '" &
            workstation.SelectedValue.ToString & "' AND [report type]= 'front label'"
            Dim adapter_fuji_front = New SqlDataAdapter(sql_fuji(3), Main.koneksi)
            adapter_fuji_front.Fill(ds3_fuji)

            If ds3_fuji.Tables(0).Rows.Count = 0 Then
                MsgBox("Fuji Front Label Printer not set in Database !")
            Else
                selectedPrinterFujifrontLabel = ds3_fuji.Tables(0).Rows(0).Item("printer").ToString
            End If

        Catch ex As Exception
        End Try

        'carton
        Try
            sql_fuji(4) = "SELECT [printer]  FROM [dbo].[printerTable] where [workstation]= '" &
            workstation.SelectedValue.ToString & "' AND [report type]= 'carton label'"
            Dim adapter_fuji_carton = New SqlDataAdapter(sql_fuji(4), Main.koneksi)
            adapter_fuji_carton.Fill(ds4_fuji)

            If ds4_fuji.Tables(0).Rows.Count = 0 Then
                MsgBox("Fuji Carton Label Printer not set in Database !")
            Else
                selectedPrinterFujiCartonLabel = ds4_fuji.Tables(0).Rows(0).Item("printer").ToString
            End If

        Catch ex As Exception
        End Try


        'Ruby
        'performance
        Try
            sql_ruby(1) = "SELECT [printer]  FROM [dbo].[printerTable] where [workstation]= '" &
            workstation.SelectedValue.ToString & "' AND [report type]= 'Ruby performance label'"
            Dim adapter_ruby_performance = New SqlDataAdapter(sql_ruby(1), Main.koneksi)
            adapter_ruby_performance.Fill(ds1_ruby)

            If ds1_ruby.Tables(0).Rows.Count = 0 Then
                'MsgBox("Ruby Performance Label Printer not set in Database !")
            Else
                selectedPrinterRubyPerformance = ds1_ruby.Tables(0).Rows(0).Item("printer").ToString
            End If
        Catch ex As Exception
        End Try

        'packaging
        Try
            sql_ruby(2) = "SELECT [printer]  FROM [dbo].[printerTable] where [workstation]= '" &
            workstation.SelectedValue.ToString & "' AND [report type]= 'Ruby packaging label'"
            Dim adapter_ruby_packaging = New SqlDataAdapter(sql_ruby(2), Main.koneksi)
            adapter_ruby_packaging.Fill(ds2_ruby)

            If ds2_ruby.Tables(0).Rows.Count = 0 Then
                'MsgBox("Ruby Performance Label Printer not set in Database !")
            Else
                selectedPrinterRubyPackaging = ds2_ruby.Tables(0).Rows(0).Item("printer").ToString
            End If
        Catch ex As Exception
        End Try

        'Out Side
        Try
            sql_ruby(3) = "SELECT [printer]  FROM [dbo].[printerTable] where [workstation]= '" &
            workstation.SelectedValue.ToString & "' AND [report type]= 'Ruby Outside label'"
            Dim adapter_ruby_outside = New SqlDataAdapter(sql_ruby(3), Main.koneksi)
            adapter_ruby_outside.Fill(ds3_ruby)

            If ds3_ruby.Tables(0).Rows.Count = 0 Then
                'MsgBox("Ruby Performance Label Printer not set in Database !")
            Else
                selectedPrinterRubyOutSide = ds3_ruby.Tables(0).Rows(0).Item("printer").ToString
            End If
        Catch ex As Exception
        End Try


        listprinter.ResetText()
        listprinter1.ResetText()
        listprinter2.ResetText()
        listprinter3.ResetText()

        'Fuji
        cbx_Carton.ResetText()
        cbx_front.ResetText()
        cbx_fuji_side_label.ResetText()
        cbx_outside.ResetText()
        cbx_Rotary.ResetText()

        'Ruby
        cbxPerfomaceRuby.ResetText()
        cbxPackagingRuby.ResetText()
        cbxOutsideRuby.ResetText()



        'listprinter.Refresh()
        'listprinter1.Refresh()
        'listprinter2.Refresh()
        'listprinter3.Refresh()

        Dim items() As String = (From item As String In listprinter.Items() Select item).ToArray

        Dim index1 = Array.IndexOf(items, selectedPrinterProduct)
        Dim index2 = Array.IndexOf(items, selectedPrinterPackaging)
        Dim index3 = Array.IndexOf(items, selectedPrinterCOC)
        Dim index4 = Array.IndexOf(items, selectedPrinterTest)
        'fuji
        Dim index_fuji1 = Array.IndexOf(items, selectedPrinterFujiSIdeLabel)
        Dim index_fuji2 = Array.IndexOf(items, selectedPrinterFujiRotaryLabel)
        Dim index_fuji3 = Array.IndexOf(items, selectedPrinterFujifrontLabel)
        Dim index_fuji4 = Array.IndexOf(items, selectedPrinterFujiCartonLabel)
        Dim index_fuji5 = Array.IndexOf(items, selectedPrinterFujiOutSIdeLabel)
        'ruby
        Dim index_ruby1 = Array.IndexOf(items, selectedPrinterRubyPerformance)
        Dim index_ruby2 = Array.IndexOf(items, selectedPrinterRubyPackaging)
        Dim index_ruby3 = Array.IndexOf(items, selectedPrinterRubyOutSide)

        'Ruby set
        'performance label
        If index_ruby1 >= 0 Then
            cbxPerfomaceRuby.SelectedText = selectedPrinterRubyPerformance
            label_performance_small_ruby.PrintSettings.PrinterName = selectedPrinterRubyPerformance
            label_performance_big_ruby.PrintSettings.PrinterName = selectedPrinterRubyPerformance
        Else
            cbxPerfomaceRuby.SelectedText = "Microsoft Print to PDF"
            label_performance_small_ruby.PrintSettings.PrinterName = "Microsoft Print to PDF"
            label_performance_big_ruby.PrintSettings.PrinterName = "Microsoft Print to PDF"
            'MsgBox("Fuji Side label Printer are choosen (" & selectedPrinterFujiSIdeLabel & ") Not installed in this PC!" & vbNewLine & "Printer will use: Microsoft Print to PDF")
        End If
        'packaging label
        If index_ruby2 >= 0 Then
            cbxPackagingRuby.SelectedText = selectedPrinterRubyPackaging
            label_packaging_ruby.PrintSettings.PrinterName = selectedPrinterRubyPackaging
        Else
            cbxPackagingRuby.SelectedText = "Microsoft Print to PDF"
            label_packaging_ruby.PrintSettings.PrinterName = "Microsoft Print to PDF"
            'MsgBox("Fuji Side label Printer are choosen (" & selectedPrinterFujiSIdeLabel & ") Not installed in this PC!" & vbNewLine & "Printer will use: Microsoft Print to PDF")
        End If
        'packaging label
        If index_ruby3 >= 0 Then
            cbxOutsideRuby.SelectedText = selectedPrinterRubyOutSide
            label_outside_ruby.PrintSettings.PrinterName = selectedPrinterRubyOutSide
        Else
            cbxOutsideRuby.SelectedText = "Microsoft Print to PDF"
            label_outside_ruby.PrintSettings.PrinterName = "Microsoft Print to PDF"
            'MsgBox("Fuji Side label Printer are choosen (" & selectedPrinterFujiSIdeLabel & ") Not installed in this PC!" & vbNewLine & "Printer will use: Microsoft Print to PDF")
        End If




        'fuji set
        If index_fuji1 >= 0 Then
            cbx_fuji_side_label.SelectedText = selectedPrinterFujiSIdeLabel
            label_side_printer.PrintSettings.PrinterName = selectedPrinterFujiSIdeLabel

        Else
            cbx_fuji_side_label.SelectedText = "Microsoft Print to PDF"
            label_side_printer.PrintSettings.PrinterName = "Microsoft Print to PDF"
            'MsgBox("Fuji Side label Printer are choosen (" & selectedPrinterFujiSIdeLabel & ") Not installed in this PC!" & vbNewLine & "Printer will use: Microsoft Print to PDF")
        End If

        If index_fuji2 >= 0 Then
            cbx_Rotary.SelectedText = selectedPrinterFujiRotaryLabel
            label_rotary_printer.PrintSettings.PrinterName = selectedPrinterFujiRotaryLabel

        Else
            cbx_Rotary.SelectedText = "Microsoft Print to PDF"
            label_rotary_printer.PrintSettings.PrinterName = "Microsoft Print to PDF"
            'MsgBox("Fuji Rotary label Printer are choosen (" & selectedPrinterFujiRotaryLabel & ") Not installed in this PC!" & vbNewLine & "Printer will use: Microsoft Print to PDF")
        End If

        If index_fuji3 >= 0 Then
            cbx_front.SelectedText = selectedPrinterFujifrontLabel
            label_front_long_printer.PrintSettings.PrinterName = selectedPrinterFujifrontLabel
            label_front_short_printer.PrintSettings.PrinterName = selectedPrinterFujifrontLabel
        Else
            cbx_front.SelectedText = "Microsoft Print to PDF"
            label_front_long_printer.PrintSettings.PrinterName = "Microsoft Print to PDF"
            label_front_short_printer.PrintSettings.PrinterName = "Microsoft Print to PDF"
            'MsgBox("Fuji Front label Printer are choosen (" & selectedPrinterFujifrontLabel & ") Not installed in this PC!" & vbNewLine & "Printer will use: Microsoft Print to PDF")
        End If

        If index_fuji4 >= 0 Then
            cbx_Carton.SelectedText = selectedPrinterFujiCartonLabel
            cbx_outside.SelectedText = selectedPrinterFujiCartonLabel
            label_carton_printer.PrintSettings.PrinterName = selectedPrinterFujiCartonLabel
            label_out_side_printer.PrintSettings.PrinterName = selectedPrinterFujiCartonLabel
        Else
            cbx_Carton.SelectedText = "Microsoft Print to PDF"
            cbx_outside.SelectedText = "Microsoft Print to PDF"
            label_carton_printer.PrintSettings.PrinterName = "Microsoft Print to PDF"
            label_out_side_printer.PrintSettings.PrinterName = "Microsoft Print to PDF"
            'MsgBox("Fuji Carton and Outside Box label Printer are choosen (" & selectedPrinterFujiCartonLabel & ") Not installed in this PC!" & vbNewLine & "Printer will use: Microsoft Print to PDF")
        End If
        '----------------------------------------------------------------------
        If index1 >= 0 Then
            listprinter.SelectedText = selectedPrinterProduct
            label_printer.PrintSettings.PrinterName = selectedPrinterProduct

        Else
            listprinter.SelectedText = "Microsoft Print to PDF"
            label_printer.PrintSettings.PrinterName = "Microsoft Print to PDF"
            MsgBox("Product label Printer are choosen (" & selectedPrinterProduct & ") Not installed in this PC!" & vbNewLine & "Printer will use: Microsoft Print to PDF")
        End If

        If index2 >= 0 Then
            listprinter1.SelectedText = selectedPrinterPackaging
            label1_printer.PrintSettings.PrinterName = selectedPrinterPackaging
        Else
            listprinter1.SelectedText = "Microsoft Print to PDF"
            label1_printer.PrintSettings.PrinterName = "Microsoft Print to PDF"
            MsgBox("Packaging label Printer are choosen (" & selectedPrinterPackaging & ") not installed in this PC!" & vbNewLine & "Printer will use: Microsoft Print to PDF")
        End If

        If index3 >= 0 Then
            listprinter2.SelectedText = selectedPrinterCOC
            label2_printer.PrintSettings.PrinterName = selectedPrinterCOC
        Else
            listprinter2.SelectedText = "Microsoft Print to PDF"
            label2_printer.PrintSettings.PrinterName = "Microsoft Print to PDF"
            MsgBox("COC label Printer are choosen (" & selectedPrinterCOC & ") not installed in this PC!" & vbNewLine & "Printer will use: Microsoft Print to PDF")
        End If

        If index4 >= 0 Then
            listprinter3.SelectedText = selectedPrinterTest
            label3_printer.PrintSettings.PrinterName = selectedPrinterTest
        Else
            listprinter3.SelectedText = "Microsoft Print to PDF"
            label3_printer.PrintSettings.PrinterName = "Microsoft Print to PDF"
            MsgBox("Test Report  Printer are choosen (" & selectedPrinterTest & ") not installed in this PC!" & vbNewLine & "Printer will use: Microsoft Print to PDF")
        End If
        'Catch ex As Exception

        'End Try
        If selectWorkstation < 15 Then
            MsgBox("Wrong Workstation was choosen !")
            PPnumberEntry.Enabled = False
        Else
            PPnumberEntry.Enabled = True
        End If


        Application.DoEvents()
        'update
        'Listprinter2_SelectedIndexChanged()
        'Listprinter3_SelectedIndexChanged()
    End Sub
    Private Sub workstation_SelectedIndexChanged() Handles workstation.SelectedIndexChanged
        workstation2.SelectedItem = workstation.SelectedItem
        workstation3.SelectedItem = workstation.SelectedItem
        workstationFuji.SelectedItem = workstation.SelectedItem

        Report_Tab.SelectedIndex = 3
        Report_Tab.SelectedIndex = 2
        Report_Tab.SelectedIndex = 1
        Report_Tab.SelectedIndex = 0

        'Select Product Label printer
        workstation_event()

    End Sub

    Private Sub ComboBox4_SelectedIndexChanged(sender As Object, e As EventArgs) Handles workstation2.SelectedIndexChanged
        workstation.SelectedItem = workstation2.SelectedItem
        workstation3.SelectedItem = workstation2.SelectedItem
    End Sub

    Private Sub ComboBox11_SelectedIndexChanged(sender As Object, e As EventArgs) Handles workstation3.SelectedIndexChanged
        workstation.SelectedItem = workstation3.SelectedItem
        workstation2.SelectedItem = workstation3.SelectedItem
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles technicianName.SelectedIndexChanged
        user.SelectedItem = technicianName.SelectedItem
        testUser.SelectedItem = technicianName.SelectedItem
        technicianNameFuji.SelectedItem = technicianName.SelectedItem
        Dim row_v As DataRowView = technicianName.SelectedItem
        If (Not row_v Is Nothing) Then
            Dim row As DataRow = row_v.Row
            Dim itemName As String = row(2).ToString()
            technicianShortName.Text = itemName
            technicianShortNameFuji.Text = itemName
            technicianShortNameRuby.Text = itemName

        End If

    End Sub

    Private Sub ComboBox5_SelectedIndexChanged(sender As Object, e As EventArgs) Handles user.SelectedIndexChanged
        If String.IsNullOrEmpty(Me.technicianName.SelectedValue) = False Then
            'technicianName.SelectedItem = user.SelectedItem
            'testUser.SelectedItem = user.SelectedItem
        End If
    End Sub

    Private Sub ComboBox10_SelectedValueChanged(sender As Object, e As EventArgs) Handles testUser.SelectedIndexChanged
        If String.IsNullOrEmpty(testUser.SelectedValue) = False Then
            Me.user.SelectedItem = Me.testUser.SelectedItem
            Me.technicianName.SelectedItem = Me.testUser.SelectedItem
        End If
    End Sub

    'Private Sub Button13_Click(sender As Object, e As EventArgs) Handles export_open_orders.Click
    '    ExportToExcel("SELECT * FROM openOrders")
    '    'Dim Dialog As New SaveFileDialog
    '    'Dim sql As String
    '    'Dialog.Filter = "Microsoft Excel 97-2003|*.xls"
    '    'If (Dialog.ShowDialog = DialogResult.OK) Then
    '    '    'MsgBox(Dialog.FileName)
    '    '    Dim xlApp As Excel.Application
    '    '    Dim xlWorkBook As Excel.Workbook
    '    '    Dim xlWorkSheet As Excel.Worksheet
    '    '    Dim misValue As Object = System.Reflection.Missing.Value
    '    '    'Dim i As Integer
    '    '    'Dim j As Integer

    '    '    xlApp = New Excel.Application
    '    '    xlWorkBook = xlApp.Workbooks.Add(misValue)
    '    '    xlWorkSheet = xlWorkBook.Sheets("Sheet1")

    '    '    Call Main.koneksi_db()
    '    '    sql = "SELECT * FROM Workstations"
    '    '    Dim adapter As New SqlDataAdapter(sql, Main.koneksi)
    '    '    Dim ds As New DataSet
    '    '    adapter.Fill(ds)

    '    '    'MsgBox(ds.Tables(0).Rows(i).Item(j).ToString)

    '    '    'For i = 0 To ds.Tables(0).Rows.Count - 1
    '    '    '    For j = 0 To ds.Tables(0).Columns.Count - 1
    '    '    '        xlWorkSheet.Cells(i + 1, j + 1) = ds.Tables(0).Rows(i).Item(j).ToString
    '    '    '    Next
    '    '    'Next

    '    '    xlApp = New Microsoft.Office.Interop.Excel.Application
    '    '    xlWorkBook = xlApp.Workbooks.Add(misValue)
    '    '    xlWorkSheet = xlWorkBook.Sheets("Sheet1")
    '    '    xlWorkSheet.SaveAs(Dialog.FileName)
    '    '    xlWorkBook.Close()
    '    '    xlApp.Quit()
    '    '    Main.koneksi.Close()

    '    'End If
    'End Sub

    Private Sub ComboBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles madeInEnglish.SelectedIndexChanged
        Dim row_v As DataRowView = madeInEnglish.SelectedItem
        If (Not row_v Is Nothing) Then
            Dim row As DataRow = row_v.Row
            Dim CN As String = row(3).ToString()
            Dim RU As String = row(4).ToString()
            madeInChinese.Text = CN
            madeInRussian.Text = RU
        End If
    End Sub

    Private Sub ComboBox6_SelectedIndexChanged(sender As Object, e As EventArgs) Handles country.SelectedIndexChanged
        Dim ds As New DataSet
        Dim dt As New DataTable
        Dim sql As String = "select * from [customerDatabase] where [country] ='" & country.Text & "'"
        Dim adapter As New SqlDataAdapter(sql, Main.koneksi)
        adapter.Fill(ds)
        adapter.Fill(dt)

        If ds.Tables(0).Rows.Count = 0 Then
            Exit Sub
        End If

        If ds.Tables(0).Rows.Count = 1 Then
            Me.customer.Text = ds.Tables(0).Rows(0).Item("customer name").ToString
            Me.Customer_code.Text = ds.Tables(0).Rows(0).Item("customer code").ToString
        End If

        If ds.Tables(0).Rows.Count > 0 Then
            countryShortName.Text = ds.Tables(0).Rows(0).Item("country short name").ToString
        End If

        Me.testCustomer.DisplayMember = "customer name"
        Me.testCustomer.ValueMember = "customer name"
        Me.testCustomer.DataSource = dt

        'Dim dsa As New DataSet
        'Dim dt As New DataTable
        'If String.IsNullOrEmpty(country.Text) = False Then

        '    Call koneksi_db()
        '    Dim adapter As New SqlDataAdapter("Select distinct [customer name], [country short name],[customer code],[customer name] FROM customerDatabase WHERE [country] = '" & Me.country.SelectedValue & "'", Main.koneksi)
        '    adapter.Fill(dsa)
        '    adapter.Fill(dt)

        '    If dsa.Tables(0).Rows.Count = 0 Then Exit Sub

        '    If dsa.Tables(0).Rows.Count = 1 Then
        '        Me.customer.Text = dsa.Tables(0).Rows(0).Item("customer name").ToString()
        '        Me.testCustomer.Text = dsa.Tables(0).Rows(0).Item("customer name").ToString()
        '        Me.Customer_code.Text = dsa.Tables(0).Rows(0).Item("customer code").ToString()
        '        Me.countryShortName.Text = dsa.Tables(0).Rows(0).Item("country short name").ToString()
        '    End If

        '    If (dsa.Tables(0).Rows.Count > 1) Then
        '        Me.countryShortName.Text = dsa.Tables(0).Rows(0).Item("country short name").ToString()
        '        Me.customer.DisplayMember = "customer name"
        '        Me.customer.ValueMember = "customer name"
        '        Me.customer.DataSource = dt
        '        Me.testCustomer.DisplayMember = "customer name"
        '        Me.testCustomer.ValueMember = "customer name"
        '        Me.testCustomer.DataSource = dt
        '    End If

        'End If

        Timer2.Enabled = True 'santo added
    End Sub

    Private Sub PPnumberEntry_KeyPress(sender As Object, e As PreviewKeyDownEventArgs) Handles PPnumberEntry.PreviewKeyDown 'sudah sesuai dengan VBA
        'Try

        If Len(Me.PPnumberEntry.Text) = 12 Then Me.PPnumberEntry.Text = Microsoft.VisualBasic.Right(Me.PPnumberEntry.Text, 11)

        'reset  santo added
        CustomerCountryNotFound = 0
        ReqDelvdt = 0
        variable_Q = 0

        'UpdateScanRecords()

        'MsgBox(Convert.ToChar(Keys.Tab))

        If (e.KeyData = Keys.Tab Or e.KeyData = Keys.Enter) And Len(Me.PPnumberEntry.Text) >= 11 Then
            'MsgBox(PPnumberEntry.Text)
            'If e.KeyChar = Chr(9) Then

            If PPnumberEntry.Text <> "" Then
                If tempPPnumber <> PPnumberEntry.Text Then
                    tempPPnumber = PPnumberEntry.Text
                Else
                    resetValmsgbox()
                    PP.Text = ""
                    PP.Text = tempPPnumber
                    'ComponentNo.Select()
                    CompToQuality.Enabled = True
                    CompToQuality.Select()
                    Exit Sub
                End If
            End If

            'MsgBox(tempPPnumber)

            resetValmsgbox()

            If String.IsNullOrEmpty(workstation.Text) = False And String.IsNullOrEmpty(technicianName.Text) = False Then
                If IsNumeric(PPnumberEntry.Text) Then
                    clean()
                    afterPPinput()
                    Me.testPP.Text = Me.PPnumberEntry.Text
                    loadTestData()
                    visualCheck = 2

                    If checkConfig() = True Then
                        updateDataVisual()
                    End If

                    If Me.CounterCpts.Text = Me.LabelQuantitycpt.Text And CInt(Me.CounterCpts.Text) > 0 Then
                        If visualCheck = 2 Then
                            Me.CounterItems.Text = CInt(Me.CounterItems.Text) + 1

                            'If Me.CounterItems.Text = Me.LabelQuantityitem.Text Then
                            '    Dim adapter As New SqlDataAdapter
                            '    Dim query = "UPDATE [PPList] SET [FinishDate] = '" & DateTime.Now.ToString("dd.MM.yyyy") & "' WHERE [Order] = '" & Me.PPnumberEntry.Text & "' and FinishDate is null"
                            '    adapter = New SqlDataAdapter(query, Main.koneksi)
                            '    adapter.SelectCommand.ExecuteNonQuery()

                            '    Dim queryGiOn = "Update PPList SET [On time]= 'On Time' WHERE convert(datetime, [GI date], 103) > convert(datetime, [FinishDate], 103) AND [On Time] is null and [FinishDate] is not null and [Order] = '" & Me.PPnumberEntry.Text & "'"
                            '    Dim querySchOn = "Update PPList SET [On timeM]= 'On Time' WHERE convert(datetime, [Scheduled finish], 103) > convert(datetime, [FinishDate], 103) AND [On TimeM] is null and [FinishDate] is not null and [Order] = '" & Me.PPnumberEntry.Text & "'"

                            '    Dim queryGiLate = "Update PPList SET [Late]= 'Late' WHERE convert(datetime, [GI date], 103) <= convert(datetime, [FinishDate], 103) AND [Late] is null and [FinishDate] is not null and [Order] = '" & Me.PPnumberEntry.Text & "'"
                            '    Dim querySchLate = "Update PPList SET [LateM]= 'Late' WHERE convert(datetime, [Scheduled finish], 103) <= convert(datetime, [FinishDate], 103) AND [Late] is null and [FinishDate] is not null and [Order] = '" & Me.PPnumberEntry.Text & "'"

                            '    adapter = New SqlDataAdapter(queryGiOn, Main.koneksi)
                            '    adapter.SelectCommand.ExecuteNonQuery()

                            '    adapter = New SqlDataAdapter(querySchOn, Main.koneksi)
                            '    adapter.SelectCommand.ExecuteNonQuery()

                            '    adapter = New SqlDataAdapter(queryGiLate, Main.koneksi)
                            '    adapter.SelectCommand.ExecuteNonQuery()

                            '    adapter = New SqlDataAdapter(querySchLate, Main.koneksi)
                            '    adapter.SelectCommand.ExecuteNonQuery()
                            'End If

                            UpdateScanRecords()
                            Me.CounterCpts.Text = 0
                        End If
                    End If
                    'UpdateComponents()
                    'Try
                    '    Dim MinCmd As SqlCommand = New SqlCommand("select MIN([Check Components]) from Components where [order]=" & Me.PPnumberEntry.Text & "", Main.koneksi)
                    '    Dim min As Integer = Convert.ToInt64(MinCmd.ExecuteScalar())
                    '    Me.CounterItems.Text = min.ToString
                    '    Me.LabelCheckCom.Text = min.ToString
                    'Catch ex As Exception

                    'End Try
                    'Dim MinComponentScan As Integer

                    'Check Jumlah Component Scan

                    'For j = 0 To Me.DataGridView1.Rows.Count - 2
                    '    If j = 0 Then
                    '        MinComponentScan = DataGridView1.Rows(j).Cells(4).Value
                    '    End If

                    '    If MinComponentScan > DataGridView1.Rows(j).Cells(4).Value Then
                    '        MinComponentScan = DataGridView1.Rows(j).Cells(4).Value
                    '    End If
                    'Next
                    'Dim hasil As Integer
                    'Dim sum As Integer = 0
                    'For k = 0 To Me.DataGridView1.Rows.Count - 2
                    '    hasil = DataGridView1.Rows(k).Cells(4).Value - MinComponentScan
                    '    If hasil > 0 Then
                    '        sum = sum + hasil
                    '    End If
                    'Next
                    'CounterCpts.Text = sum

                    'end Check Jumlah Component Scan

                    ' Load data auto print
                    AutoPrintProductLabelFrom.Text = CounterItems.Text
                Else
                    MsgBox("Sorry The PP must be number")
                    Me.PPnumberEntry.Text = ""
                End If
            Else
                MsgBox("Sorry you Must select workstations and technician first")
            End If
            'make sure labele selected
            'If selectedLabel.Text = "" Then selectLabel()
        End If
        'Try
        '    Dim ds2 As New DataSet
        '    Dim adapter2 = New SqlDataAdapter("Select * FROM [PPList] WHERE [Order] = '" & Me.PPnumberEntry.Text & "'", Main.koneksi)
        '    adapter2.Fill(ds2)
        '    If Not IsDBNull(ds2.Tables(0).Rows(0).Item("Name 1")) Then
        '        Me.customer.Text = ds2.Tables(0).Rows(0).Item("Name 1").ToString
        '    End If
        'Catch ex As Exception

        'End Try
        'Catch ex As Exception
        '    MsgBox("Check PP Number " & ex.Message)
        'End Try

        'select label



    End Sub

    Sub clean() 'Clean data 'sudah sesuai dengan VBA
        Fuji_QR_Product_Label.Text = ""
        SaveData.Text = ""
        select_quantity.Checked = False
        sekaliOF = 0
        hitung = 0
        missingMadein = 0
        TextBox_Neutral.Text = ""
        LabelCheckCom.Text = 0
        Me.DataGridView1.Rows.Clear()
        Me.SD.Text = ""
        Me.StartTestQuantity.Text = 0
        Me.boxQty.Text = 0
        Me.Label468.Text = 0
        Me.Picture2.Text = ""
        Me.dateToPrint.Text = ""
        Me.finishDate.Text = ""
        Me.testQuantity.Text = 0
        Me.cek_testQuantity.Text = 0
        Me.testCountry.Text = ""
        Me.testCRD.Text = ""
        Me.testCreationDate.Text = ""
        Me.testCustomer.Text = ""
        Me.testCustomerCode.Text = ""
        Me.testMaterial.Text = ""
        Me.testSO.Text = ""
        Me.testSOitem.Text = ""
        Me.checkTestPrint.CheckState = 1
        Me.checkCOCPrint.CheckState = 1
        Me.testCustPO.Text = ""
        Me.testCustPOitem.Text = ""
        Me.testDescription.Text = ""
        Me.productRange.Text = ""
        Me.custMaterial.Text = ""
        Me.breakerDevice.Text = ""
        Me.breakerPole.Text = ""
        Me.breakerRef.Text = ""
        Me.breakerType.Text = ""

        Me.micrologicRef.Text = ""
        Me.micrologicType.Text = ""

        Me.plugRef.Text = ""
        Me.plugType.Text = ""

        Me.chassisRef.Text = ""

        Me.drawoutFixed.Text = ""
        Me.topConnect.Text = ""
        Me.bottomConnect.Text = ""

        Me.MXtype.Text = ""
        Me.XFtype.Text = ""

        Me.MCH.Text = ""
        Me.XF.Text = ""

        Me.MX.Text = ""
        Me.MX2.Text = ""
        Me.MN.Text = ""
        Me.SDE2.Text = ""
        Me.remoteReset.Text = ""

        Me.header.Text = ""
        Me.range.Text = ""
        Me.description.Text = ""
        Me.EAN13.Text = ""
        Me.mat1.Text = ""
        Me.mat2.Text = ""
        Me.mat3.Text = ""
        Me.mat4.Text = ""
        Me.mat5.Text = ""
        Me.mat6.Text = ""
        Me.mat7.Text = ""
        Me.mat8.Text = ""
        Me.mat9.Text = ""
        Me.mat10.Text = ""

        Me.descr1.Text = ""
        Me.descr2.Text = ""
        Me.descr3.Text = ""
        Me.descr4.Text = ""
        Me.descr5.Text = ""
        Me.descr6.Text = ""
        Me.descr7.Text = ""
        Me.descr8.Text = ""
        Me.descr9.Text = ""
        Me.descr10.Text = ""
        Me.CounterItems.Text = 0
        Me.LabelQuantityitem.Text = 0
        Me.CounterCpts.Text = 0
        Me.LabelQuantitycpt.Text = 0
        Me.Quantity.Text = 0

        Me.qty1.Text = ""
        Me.qty2.Text = ""
        Me.qty3.Text = ""
        Me.qty4.Text = ""
        Me.qty5.Text = ""
        Me.qty6.Text = ""
        Me.qty7.Text = ""
        Me.qty8.Text = ""
        Me.qty9.Text = ""
        Me.qty10.Text = ""
        Me.palletQty.Text = 0

        Me.quantityLabel.Text = 0
        Me.StartLabel.Text = 0
        Me.StartPackingLabel.Text = 0
        Me.StartPackingLabel2.Text = 0
        Me.labelQty2.Text = 0
        Me.labelQty.Text = 0

        Me.warning.Text = ""
        Me.warning.Visible = False

        Me.checkMadeInChina.CheckState = 1
        Me.checkAssembledSingapore.CheckState = 1
        Me.madeInSingapore.Text = "Assembled by Schneider Electric in Singapore"

        Me.checkVBP.CheckState = 0
        Me.checkVCPO.CheckState = 0
        Me.checkVSPO.CheckState = 0
        Me.checkVSPDchassis.CheckState = 0
        Me.checkVSPDall.CheckState = 0
        Me.checkVPEC.CheckState = 0
        Me.checkVPOC.CheckState = 0
        Me.checkDAE.CheckState = 0
        Me.checkVDC.CheckState = 0
        Me.checkIPA.CheckState = 0
        Me.checkIPBO.CheckState = 0
        Me.checkVICV.CheckState = 0


        Me.checkCDM.CheckState = 0
        Me.checkCB.CheckState = 0
        Me.checkCDP.CheckState = 0
        Me.checkCP.CheckState = 0
        Me.checkOP.CheckState = 0
        Me.checkPhaseBarrier.CheckState = 0

        Me.checkRotary.CheckState = 0
        Me.checkEF.CheckState = 0
        Me.CD.Text = ""
        Me.CE.Text = ""
        Me.CT.Text = ""
        Me.checkPTE.CheckState = 0
        Me.checkR.CheckState = 0
        Me.checkRr.CheckState = 0

        Me.checkM2C.CheckState = 0
        Me.checkCOM.CheckState = 0
        Me.checkATS.CheckState = 0
        Me.checkAD.CheckState = 0
        Me.checkBattery.CheckState = 0
        Me.checkexternalCT.CheckState = 0

        Me.cluster.Text = ""
        Me.checkStandard.CheckState = 0
        Me.checkLow.CheckState = 0
        Me.checkHigh.CheckState = 0
        Me.checkLT.CheckState = 0

        Me.SDE1.CheckState = 0
        Me.SDE2count.Text = ""
        Me.IC_OF.Text = ""
        Me.PF.Text = ""

        Me.checkScrew.CheckState = 1
        Me.checkLabel.CheckState = 1
        Me.checkLeaflet.CheckState = 1

        Me.fixingScrew.Text = ""


        Me.madeInChinese.Text = ""
        Me.madeInEnglish.SelectedText = ""
        Me.madeInRussian.Text = ""

        Me.description.Text = ""
        Me.descriptionChinese.Text = ""
        Me.descriptionFrench.Text = ""
        Me.descriptionRussian.Text = ""
        Me.descriptionSpanish.Text = ""
        Me.technicalDescription.Text = ""

        Me.logo1value.Text = ""
        Me.logo2value.Text = ""
        Me.logo3value.Text = ""

        Me.logo3value.Text = ""
        Me.logo4value.Text = ""
        Me.logo5value.Text = ""

        Me.AutoPrintProductLabelFrom.Text = 0
        Me.AutoPrintPackagingLabelFrom.Text = 0

        CleanComp()

    End Sub

    'Private Sub PPnumberEntry_AfterUpdate() 'Update of the PP
    '    If Me.PPnumberEntry.Text <> "" Then
    '        tempPPnumber = Me.PPnumberEntry.Text
    '    End If

    '    If Len(Me.PPnumberEntry.Text) = 12 Then Me.PPnumberEntry.Text = Microsoft.VisualBasic.Right(Me.PPnumberEntry.Text, 11)

    '    clean
    '    afterPPinput
    '    Me.testPP = Me.PPnumberEntry
    '    loadTestData
    '    UpdateComponents()
    '    visualCheck = 2


    '    If checkConfig = True Then

    '        updateDataVisual

    '    End If

    '    If CInt(Me.CounterCpts.Text) = CInt(Me.LabelQuantitycpt.Text) And Integer.Parse(Me.CounterCpts.Text) > 0 Then
    '        If visualCheck = 2 Then



    '            Me.CounterItems.Text = Integer.Parse(Me.CounterItems.Text) + 1
    '            UpdateScanRecords()

    '            Me.CounterCpts.Text = 0
    '        End If
    '    End If
    'End Sub

    Function checkConfig() As Boolean
        'Updateby Extendoffice 20161109
        '    Application.ScreenUpdating = False
        Dim FilePath As String
        FilePath = ""
        On Error Resume Next
        FilePath = Dir("C:\config.ini")
        On Error GoTo 0
        If FilePath = "" Then
            checkConfig = False
            '        MsgBox "File doesn't exist", vbInformation, "Kutools for Excel"
        Else
            checkConfig = True
            '        MsgBox "File exist", vbInformation, "Kutools for Excel"
        End If
        '    Application.ScreenUpdating = False
    End Function

    Sub componentNo_Scan()

        If Me.ComponentNo.Text <> "" Then
            tempComponentNo = Me.ComponentNo.Text
        End If

        Dim ds As New DataSet
        Dim ds1 As New DataSet
        Dim ds2 As New DataSet
        Dim dsWorkstation As New DataSet
        Dim member As Integer
        Dim dateCode As String
        Dim barcode As String
        Dim chosenLabel As String
        Dim start As Long
        Dim qty As Long
        Dim pindah As Boolean = False
        Dim startStrBC As Integer
        Dim startStrdC As Integer

        If Me.LabelQuantitycpt.Text = 0 And Me.ComponentNo.Text = Me.PP.Text Then pindah = True

        barcode = Me.ComponentNo.Text

        If Len(Me.ComponentNo.Text) = 24 Then Me.ComponentNo.Text = Microsoft.VisualBasic.Right(Microsoft.VisualBasic.Left(Me.ComponentNo.Text, 11), 10)

        If Len(Me.ComponentNo.Text) = 10 Then
            dateCode = Microsoft.VisualBasic.Left(Me.ComponentNo.Text, 5)
            Me.ComponentNo.Text = Microsoft.VisualBasic.Right(Me.ComponentNo.Text, 5)
        ElseIf Len(Me.ComponentNo.Text) >= 11 Then
            dateCode = Microsoft.VisualBasic.Right(Microsoft.VisualBasic.Left(Me.ComponentNo.Text, 6), 5)
            Me.ComponentNo.Text = Microsoft.VisualBasic.Left(Microsoft.VisualBasic.Right(Me.ComponentNo.Text, 6), 5)
        Else
            dateCode = "NOData"
        End If

        If ComponentNo.Text = "" Then
            Exit Sub
        End If

        If CInt(Me.CounterItems.Text) < Me.Quantity.Text Then

            Dim adapter = New SqlDataAdapter("Select [ID],[Material], [Code], [Check Components], [order], [Reqmts qty],[Limit],[barcode],[datecode] FROM [Components] WHERE [Order] = " & Me.PPnumberEntry.Text & " and Workstation = '" & Me.workstation.SelectedValue & "'", Main.koneksi)
            adapter.Fill(ds)

            If (ds.Tables(0).Rows.Count > 0) Then
                If Me.CounterCpts.Text < CInt(Me.LabelQuantitycpt.Text) Then
                    For i = 0 To ds.Tables(0).Rows.Count - 1
                        If Me.ComponentNo.Text = ds.Tables(0).Rows(i).Item("Material").ToString Or Me.ComponentNo.Text = ds.Tables(0).Rows(i).Item("code").ToString Then
                            member = 1
                            If ds.Tables(0).Rows(i).Item("Check Components") < ds.Tables(0).Rows(i).Item("Limit") Then
                                If Not IsDBNull(ds.Tables(0).Rows(i).Item("barcode")) Then startStrBC = InStr(ds.Tables(0).Rows(i).Item("barcode"), barcode)
                                If Not IsDBNull(ds.Tables(0).Rows(i).Item("dateCode")) Then startStrdC = InStr(ds.Tables(0).Rows(i).Item("dateCode"), dateCode)
                                If startStrBC > 0 And Not IsDBNull(startStrBC) And startStrdC > 0 And Not IsDBNull(startStrdC) Then
                                    ds.Tables(0).Rows(i).Item("barcode") = Microsoft.VisualBasic.Left(ds.Tables(0).Rows(i).Item("barcode"), startStrBC - 1) & Replace(ds.Tables(0).Rows(i).Item("barcode"), "[", "[" & CInt(Me.CounterItems.Text) + 1 & ",", startStrBC, 1)
                                    ds.Tables(0).Rows(i).Item("dateCode") = Microsoft.VisualBasic.Left(ds.Tables(0).Rows(i).Item("dateCode"), startStrdC - 1) & Replace(ds.Tables(0).Rows(i).Item("dateCode"), "[", "[" & CInt(Me.CounterItems.Text) + 1 & ",", startStrdC, 1)
                                    adapter.UpdateCommand = New SqlCommandBuilder(adapter).GetUpdateCommand()
                                    adapter.Update(ds)
                                Else
                                    ds.Tables(0).Rows(i).Item("barcode") = barcode & "[" & CInt(Me.CounterItems.Text) + 1 & "]" & ds.Tables(0).Rows(i).Item("barcode")
                                    ds.Tables(0).Rows(i).Item("dateCode") = dateCode & "[" & CInt(Me.CounterItems.Text) + 1 & "]" & ds.Tables(0).Rows(i).Item("dateCode")
                                    adapter.UpdateCommand = New SqlCommandBuilder(adapter).GetUpdateCommand()
                                    adapter.Update(ds)
                                End If

                                Dim adapter2 = New SqlDataAdapter("SELECT * FROM [PPList] WHERE [Order] = '" & Me.PPnumberEntry.Text & "'", Main.koneksi)
                                adapter2.Fill(ds2)

                                If ds2.Tables(0).Rows.Count > 0 Then
                                    If ds2.Tables(0).Rows(0).Item("PrCtr") = 10 Then
                                        If ds.Tables(0).Rows(i).Item("Limit") - ds.Tables(0).Rows(i).Item("Check Components") >= 5 Then
                                            Me.CounterCpts.Text = Me.CounterCpts.Text + 5
                                            ds.Tables(0).Rows(i).Item("Check Components") = ds.Tables(0).Rows(i).Item("Check Components") + 5
                                            adapter.UpdateCommand = New SqlCommandBuilder(adapter).GetUpdateCommand()
                                            adapter.Update(ds)
                                        Else
                                            Me.CounterCpts.Text = Me.CounterCpts.Text + ds.Tables(0).Rows(i).Item("Limit") - ds.Tables(0).Rows(i).Item("Check Components")
                                            ds.Tables(0).Rows(i).Item("Check Components") = ds.Tables(0).Rows(i).Item("Limit")
                                            adapter.UpdateCommand = New SqlCommandBuilder(adapter).GetUpdateCommand()
                                            adapter.Update(ds)
                                        End If
                                    Else
                                        Me.CounterCpts.Text = Me.CounterCpts.Text + 1
                                        ds.Tables(0).Rows(i).Item("Check Components") = ds.Tables(0).Rows(i).Item("Check Components") + 1
                                        adapter.UpdateCommand = New SqlCommandBuilder(adapter).GetUpdateCommand()
                                        adapter.Update(ds)
                                    End If
                                End If

                                'Refresh_DGV()
                                If hasnotbeenfound <> 1 And hasnotbeenindatabase <> 1 Then Refresh_DGV() 'santo added

                                'Exit Sub

                                If Me.range.Text = "Kitting" And Me.CounterCpts.Text = Me.LabelQuantitycpt.Text And Me.CounterItems.Text = Me.Check.Text Then
                                    Exit Sub
                                End If

                                '################################# if all components have been scanned trigger camera ##############################
                                If pindah = True Then
                                    finishVision = False

                                    If Me.CounterCpts.Text = Me.LabelQuantitycpt.Text Then
                                        'apakah ini
                                        'Me.CounterItems = Me.CounterItems + 1
                                        adapter2 = New SqlDataAdapter("SELECT * FROM Workstations WHERE wkname= '" & Me.workstation.SelectedValue & "'", Main.koneksi)
                                        adapter2.Fill(dsWorkstation)

                                        'qty label
                                        Me.quantityLabel.Text = (CInt(Me.CounterItems.Text) + 1) * CInt(Me.Quantity.Text) / CInt(Me.Check.Text)

                                        Me.testQuantity.Text = (CInt(Me.CounterItems.Text) + 1) * CInt(Me.Quantity.Text) / CInt(Me.Check.Text)
                                        'santo cek
                                        Me.cek_testQuantity.Text = (CInt(Me.CounterItems.Text) + 1) * CInt(Me.Quantity.Text) / CInt(Me.Check.Text)

                                        'UpdateScanRecords
                                        If dsWorkstation.Tables(0).Rows.Count > 0 Then
                                            If dsWorkstation.Tables(0).Rows(i).Item("onelabel") >= 1 And Me.range.Text <> "Kitting" Or dsWorkstation.Tables(0).Rows(i).Item("onelabel") >= 1 And CInt(Me.CounterItems.Text) + 1 = Me.Check.Text Then
                                                selectedPrinter = ""
                                                If checkAllDataOk() = False Then Exit Sub
                                                chosenLabel = selectLabel()

                                                If chosenLabel = "" Then Exit Sub
                                                start = (CInt(Me.CounterItems.Text) + 1) * CInt(Me.Quantity.Text) / CInt(Me.Check.Text)
                                                qty = (CInt(Me.CounterItems.Text) + 1) * CInt(Me.Quantity.Text) / CInt(Me.Check.Text)
                                                printProductLabel(chosenLabel)
                                                'If ds2.Tables(0).Rows(i).Item("onelabel") >= 2 Then
                                                '##################3 skip by tuguss ####################
                                                'Dim AppAccess As Object
                                                ''MsgBox Me.databaseToUpdatePath2.Caption
                                                'Set AppAccess = GetObject(Me.databaseToUpdatePath2.Caption)
                                                'AppAccess.Forms![Main Menu].PP = ""
                                                'AppAccess.Forms![Main Menu].PP = Me.PPnumberEntry
                                                'AppAccess.Forms![Main Menu].PP.SetFocus

                                                'Set AppAccess = Nothing
                                                '##################3 skip by tuguss ####################
                                                'End If
                                            End If
                                        End If
                                    End If
                                End If

                                checkTrigger()

                                'Update [Limit]
                                On Error GoTo 3

                                If Me.range.Text = "Kitting" And CInt(Me.LabelQuantitycpt.Text) > CInt(Me.CounterCpts.Text) Then
                                Else
                                    If (ds.Tables(0).Rows.Count > 0) Then
                                        If Me.range.Text = "Kitting" Then
                                            For i2 = 0 To ds.Tables(0).Rows.Count - 1
                                                If ds.Tables(0).Rows(i2).Item("reqmts qty") - ds.Tables(0).Rows(i2).Item("Check Components") < CInt(Me.boxQty.Text) Then
                                                    Me.LabelQuantitycpt.Text = (ds.Tables(0).Rows(i2).Item("reqmts qty") - ds.Tables(0).Rows(i2).Item("Check Components")) * ds.Tables(0).Rows.Count
                                                    'MsgBox("suspect 1")
                                                    ds.Tables(0).Rows(i2).Item("Limit") = ds.Tables(0).Rows(i2).Item("reqmts qty")
                                                Else
                                                    ds.Tables(0).Rows(i2).Item("Limit") = ds.Tables(0).Rows(i2).Item("Limit") + CInt(Me.boxQty.Text)
                                                End If

                                                adapter.UpdateCommand = New SqlCommandBuilder(adapter).GetUpdateCommand()
                                                adapter.Update(ds)
                                            Next
                                        Else
                                            For i3 = 0 To ds.Tables(0).Rows.Count - 1
                                                ds.Tables(0).Rows(i3).Item("Limit") = (CInt(Me.CounterItems.Text) + 1) * ds.Tables(0).Rows(i3).Item("reqmts qty") / CInt(Me.Check.Text)
                                                adapter.UpdateCommand = New SqlCommandBuilder(adapter).GetUpdateCommand()
                                                adapter.Update(ds)
                                            Next
                                        End If

                                    Else

                                    End If
                                End If

                                '################################# if all components have been scanned trigger camera ##############################

3:
                                '---------------> If all items have been scanned
                                If Me.CounterItems.Text = Me.Check.Text Then
                                    Me.ComponentNo.Text = ""
                                    'Me.PPnumberEntry.Text = "" tidak perlu dikosongin andy minta gitu
                                    Me.PPnumberEntry.Select()
                                    'If Not IsDBNull(dsWorkstation.Tables(0).Rows(i).Item("onelabel")) Then
                                    '    If dsWorkstation.Tables(0).Rows(i).Item("onelabel") >= 1 Then
                                    '        Me.PPnumberEntry.Text = ""
                                    '        Me.ComponentNo.Select()
                                    '        Me.PPnumberEntry.Select()
                                    '    End If
                                    'End If

                                    '################ skip insert by tuguss #####################
                                    'DoCmd.SetWarnings False


                                    'sql2 = "INSERT INTO [ScanRecords" & Me.workstation & "] IN 'D:\1. KERJAAN\DT 2019\Adaptation Hub\ScanRecords.accdb'" &
                                    '            "SELECT * FROM Components WHERE [Components].[Workstation]='" & Me.workstation & "' ;"

                                    'DoCmd.RunSQL(sql2)

                                    'DoCmd.SetWarnings True

                                    '##########################################################

                                    Exit Sub
                                End If

                                Me.ComponentNo.Text = ""
                                'Me.ComponentNo.Select()
                                CompToQuality.Select()
                                Exit Sub

                                'End If

                            End If
                        End If
                    Next

                Else
                    'If Convert.ToDecimal(LabelQuantitycpt.Text) = 0 Then
                    '    afterPPinput()
                    'Else
                    MsgBox("All components have been scanned")

                    Me.ComponentNo.Text = ""
                    'Me.ComponentNo.Select()
                    CompToQuality.Select()
                    Exit Sub
                    'End If
                End If

            Else
                'MsgBox("No data found")
                Me.ComponentNo.Text = ""
                'Me.ComponentNo.Select()
                CompToQuality.Select()
                Exit Sub
            End If

        Else

            'If Me.CounterItems.Text = Me.LabelQuantityitem.Text Then
            '    Dim adapter As New SqlDataAdapter
            '    Dim query = "UPDATE [PPList] SET [FinishDate] = '" & DateTime.Now.ToString("dd.MM.yyyy") & "' WHERE [Order] = '" & Me.PPnumberEntry.Text & "' and FinishDate is null"
            '    adapter = New SqlDataAdapter(query, Main.koneksi)
            '    adapter.SelectCommand.ExecuteNonQuery()

            '    Dim queryGiOn = "Update PPList SET [On time]= 'On Time' WHERE convert(datetime, [GI date], 103) > convert(datetime, [FinishDate], 103) AND [On Time] is null and [FinishDate] is not null and [Order] = '" & Me.PPnumberEntry.Text & "'"
            '    Dim querySchOn = "Update PPList SET [On timeM]= 'On Time' WHERE convert(datetime, [Scheduled finish], 103) > convert(datetime, [FinishDate], 103) AND [On TimeM] is null and [FinishDate] is not null and [Order] = '" & Me.PPnumberEntry.Text & "'"

            '    Dim queryGiLate = "Update PPList SET [Late]= 'Late' WHERE convert(datetime, [GI date], 103) <= convert(datetime, [FinishDate], 103) AND [Late] is null and [FinishDate] is not null and [Order] = '" & Me.PPnumberEntry.Text & "'"
            '    Dim querySchLate = "Update PPList SET [LateM]= 'Late' WHERE convert(datetime, [Scheduled finish], 103) <= convert(datetime, [FinishDate], 103) AND [Late] is null and [FinishDate] is not null and [Order] = '" & Me.PPnumberEntry.Text & "'"

            '    adapter = New SqlDataAdapter(queryGiOn, Main.koneksi)
            '    adapter.SelectCommand.ExecuteNonQuery()

            '    adapter = New SqlDataAdapter(querySchOn, Main.koneksi)
            '    adapter.SelectCommand.ExecuteNonQuery()

            '    adapter = New SqlDataAdapter(queryGiLate, Main.koneksi)
            '    adapter.SelectCommand.ExecuteNonQuery()

            '    adapter = New SqlDataAdapter(querySchLate, Main.koneksi)
            '    adapter.SelectCommand.ExecuteNonQuery()
            'End If

            MsgBox("All Breakers have been adapted")
            Me.ComponentNo.Text = ""
            'Me.ComponentNo.Select()
            CompToQuality.Select()
            Exit Sub

        End If

        If member = 1 Then

            If Me.CounterItems.Text = Me.LabelQuantityitem.Text Then
                MsgBox("The component has already been scanned")
            Else
                If CInt(Me.CounterCpts.Text) = CInt(Me.LabelQuantitycpt.Text) Or CInt(Me.CounterCpts.Text) = 0 Then

                    Me.PPnumberEntry.Text = ""
                    Me.PPnumberEntry.Select()
                    Me.PPnumberEntry.Text = tempPPnumber
                    Me.PPnumberEntry.Select()

                    CleanComp()
                    UpdateComponents()

                    If checkConfig() = True Then
                        updateDataVisual()
                    End If

                Else
                    MsgBox("The component " & Me.ComponentNo.Text & " has been scanned!")
                    ComponentNo.Text = ""
                    Exit Sub
                End If


                Me.ComponentNo.Text = ""
                Me.ComponentNo.Select()
                'CompToQuality.Select()
                Me.ComponentNo.Text = tempComponentNo
                'CompToQuality.Text = tempComponentNo
                Me.ComponentNo.Select()
                'CompToQuality.Select()
                componentNo_Scan()

                Exit Sub
            End If

        Else
            MsgBox("Wrong Components", vbExclamation, "The component is not part of the product")
            If barcode.Length > 128 Then
                barcode = Microsoft.VisualBasic.Left(barcode, 127)
            End If
            Dim sql As String = "INSERT INTO ScanError (workstation, technician, day, PP, BarcodeScanned) VALUES ('" & Me.workstation.SelectedValue & "','" & Me.technicianName.SelectedValue & "','" & DateTime.Now & "','" & Me.PPnumberEntry.Text & "','" & barcode & "')"
            Dim cmd = New SqlCommand(sql, Main.koneksi)
            cmd.ExecuteNonQuery()

        End If

        Me.ComponentNo.Text = ""
        'Me.ComponentNo.Select()
        CompToQuality.Select()

        AutoPrintProductLabelFrom.Text = "0"
        AutoPrintPackagingLabelFrom.Text = "0"

        'If header.Text.Contains("BW") Then CheckProductLabelPrinting.Checked = False

    End Sub

    Private Sub cek_fuji_barcode()
        SaveData.Text = CompToQuality.Text
        'Dim phrase As String = SaveData.Text.Substring(SaveData.Text.IndexOf(" ") + 1, SaveData.Text.Length - SaveData.Text.IndexOf(" "))
        Dim a As Integer = SaveData.Text.IndexOf(" ")
        Dim idx_ref As Integer = SaveData.Text.IndexOf("f=")
        Dim idx_sn As Integer = SaveData.Text.IndexOf("n=TC")
        Dim b As Integer = SaveData.Text.Length
        Dim p As String = SaveData.Text.Substring(a + 1, b - a - 1)
        p = p.Replace(" ", "")
        'MessageBox.Show("tes")
        If CompToQuality.Text.Contains("E ") Then
            Dim QrCodeFujiSideLabel = header.Text & "/sn=SG" & p 'Microsoft.VisualBasic.Right(SaveData.Text, 9)
            Fuji_QR_Product_Label.Text = QrCodeFujiSideLabel
            CompToQuality.Text = SaveData.Text.Substring(0, SaveData.Text.IndexOf("E"))

        ElseIf CompToQuality.Text.Contains(" ") Then
            Dim QrCodeFujiSideLabel = header.Text & "/sn=SG" & p 'Microsoft.VisualBasic.Right(SaveData.Text, 13)
            Fuji_QR_Product_Label.Text = QrCodeFujiSideLabel
            CompToQuality.Text = SaveData.Text.Substring(0, SaveData.Text.IndexOf(" "))

        ElseIf CompToQuality.Text.Contains("http://go2se.com/ref") Then
            p = SaveData.Text.Substring(idx_sn + 1, b - idx_sn - 1)
            p = p.Replace("=TC", "")

            Dim QrCodeFujiSideLabel = header.Text & "/sn=SG" & p 'Microsoft.VisualBasic.Right(SaveData.Text, 13)
            Fuji_QR_Product_Label.Text = QrCodeFujiSideLabel

            Dim s As String = SaveData.Text
            Dim i As Integer = idx_ref
            Dim f As String = s.Substring(i + 2, s.IndexOf("/s", i + 1) - i - 2)
            'MsgBox(f)
            CompToQuality.Text = f

        End If
    End Sub

    Private Sub CompToQuality_TextChanged(sender As Object, e As EventArgs) Handles CompToQuality.TextChanged
        'santo add 16 maret 21
        'If Me.CompToQuality.Text.Length = 22 Or Me.CompToQuality.Text.Length = 19 Then cek_fuji_barcode()
    End Sub
    Private Sub ComponentNo_TextChanged(sender As Object, e As EventArgs) Handles ComponentNo.TextChanged

        If Me.ComponentNo.Text.Length <= 3 Then Exit Sub  'untuk antisipasi error enter

        'coba santo
        If CounterCpts.Text <> LabelQuantitycpt.Text Then AutoPrintProductLabelFrom.Text = CounterItems.Text


        If Me.ComponentNo.Text <> "" Then
            tempComponentNo = Me.ComponentNo.Text
        End If


        If ComponentNo.Text <> "" Then
            Dim ds As New DataSet
            Dim ds1 As New DataSet
            Dim ds2 As New DataSet
            Dim dss3 As New DataSet
            Dim dsWorkstation As New DataSet
            Dim member As Integer
            Dim dateCode As String
            Dim barcode As String
            Dim chosenLabel As String
            Dim start As Long
            Dim qty As Long
            Dim pindah As Boolean = False
            Dim startStrBC As Integer
            Dim startStrdC As Integer

            If Me.LabelQuantitycpt.Text = 0 And Me.ComponentNo.Text = Me.PP.Text Then pindah = True

            barcode = Me.ComponentNo.Text

            If Len(Me.ComponentNo.Text) = 24 Then Me.ComponentNo.Text = Microsoft.VisualBasic.Right(Microsoft.VisualBasic.Left(Me.ComponentNo.Text, 11), 10)

            If Len(Me.ComponentNo.Text) = 10 Then
                dateCode = Microsoft.VisualBasic.Left(Me.ComponentNo.Text, 5)
                Me.ComponentNo.Text = Microsoft.VisualBasic.Right(Me.ComponentNo.Text, 5)
            ElseIf Len(Me.ComponentNo.Text) >= 11 Then
                dateCode = Microsoft.VisualBasic.Right(Microsoft.VisualBasic.Left(Me.ComponentNo.Text, 6), 5)
                Me.ComponentNo.Text = Microsoft.VisualBasic.Left(Microsoft.VisualBasic.Right(Me.ComponentNo.Text, 6), 5)
            Else
                dateCode = "NOData"
            End If

            If ComponentNo.Text = "" Then
                Exit Sub
            End If

            If CInt(Me.CounterItems.Text) < Me.Quantity.Text Then

                Dim adapter = New SqlDataAdapter("Select [ID],[Material], [Code], [Check Components], [order], [Reqmts qty],[Limit],[barcode],[datecode] FROM [Components] WHERE [Order] = " & Me.PPnumberEntry.Text & "and Workstation = '" & Me.workstation.SelectedValue & "'", Main.koneksi)
                adapter.Fill(ds)

                If (ds.Tables(0).Rows.Count > 0) Then
                    If Me.CounterCpts.Text < CInt(Me.LabelQuantitycpt.Text) Then
                        For i = 0 To ds.Tables(0).Rows.Count - 1
                            If Me.ComponentNo.Text = ds.Tables(0).Rows(i).Item("Material").ToString Or Me.ComponentNo.Text = ds.Tables(0).Rows(i).Item("code").ToString Then
                                member = 1
                                If ds.Tables(0).Rows(i).Item("Check Components") < ds.Tables(0).Rows(i).Item("Limit") Then
                                    If Not IsDBNull(ds.Tables(0).Rows(i).Item("barcode")) Then startStrBC = InStr(ds.Tables(0).Rows(i).Item("barcode"), barcode)
                                    If Not IsDBNull(ds.Tables(0).Rows(i).Item("dateCode")) Then startStrdC = InStr(ds.Tables(0).Rows(i).Item("dateCode"), dateCode)
                                    If startStrBC > 0 And Not IsDBNull(startStrBC) And startStrdC > 0 And Not IsDBNull(startStrdC) Then
                                        ds.Tables(0).Rows(i).Item("barcode") = Microsoft.VisualBasic.Left(ds.Tables(0).Rows(i).Item("barcode"), startStrBC - 1) & Replace(ds.Tables(0).Rows(i).Item("barcode"), "[", "[" & CInt(Me.CounterItems.Text) + 1 & ",", startStrBC, 1)
                                        ds.Tables(0).Rows(i).Item("dateCode") = Microsoft.VisualBasic.Left(ds.Tables(0).Rows(i).Item("dateCode"), startStrdC - 1) & Replace(ds.Tables(0).Rows(i).Item("dateCode"), "[", "[" & CInt(Me.CounterItems.Text) + 1 & ",", startStrdC, 1)
                                        adapter.UpdateCommand = New SqlCommandBuilder(adapter).GetUpdateCommand()
                                        adapter.Update(ds)
                                    Else
                                        ds.Tables(0).Rows(i).Item("barcode") = barcode & "[" & CInt(Me.CounterItems.Text) + 1 & "]" & ds.Tables(0).Rows(i).Item("barcode")
                                        ds.Tables(0).Rows(i).Item("dateCode") = dateCode & "[" & CInt(Me.CounterItems.Text) + 1 & "]" & ds.Tables(0).Rows(i).Item("dateCode")
                                        adapter.UpdateCommand = New SqlCommandBuilder(adapter).GetUpdateCommand()
                                        adapter.Update(ds)
                                    End If

                                    Dim adapter2 = New SqlDataAdapter("SELECT * FROM [PPList] WHERE [Order] = '" & Me.PPnumberEntry.Text & "'", Main.koneksi)
                                    adapter2.Fill(ds2)

                                    If ds2.Tables(0).Rows.Count > 0 Then
                                        'santo add last additional
                                        Dim abc As Integer = 0
                                        Dim bcd As String = ds.Tables(0).Rows(i).Item("Material")
                                        If bcd = "47840" Or bcd = "47837" Or bcd = "47883" Or bcd = "47884" Then abc = 1
                                        '////////////////////////////////////////////////////////////////////////

                                        If ds2.Tables(0).Rows(0).Item("PrCtr") = 10 Or abc = 1 Then
                                            If ds.Tables(0).Rows(i).Item("Limit") - ds.Tables(0).Rows(i).Item("Check Components") >= 5 And abc <> 1 Then
                                                Me.CounterCpts.Text = Me.CounterCpts.Text + 5
                                                ds.Tables(0).Rows(i).Item("Check Components") = ds.Tables(0).Rows(i).Item("Check Components") + 5
                                                adapter.UpdateCommand = New SqlCommandBuilder(adapter).GetUpdateCommand()
                                                adapter.Update(ds)
                                            Else
                                                Me.CounterCpts.Text = Me.CounterCpts.Text + ds.Tables(0).Rows(i).Item("Limit") - ds.Tables(0).Rows(i).Item("Check Components")
                                                ds.Tables(0).Rows(i).Item("Check Components") = ds.Tables(0).Rows(i).Item("Limit")
                                                adapter.UpdateCommand = New SqlCommandBuilder(adapter).GetUpdateCommand()
                                                adapter.Update(ds)
                                            End If
                                        Else
                                            Me.CounterCpts.Text = Me.CounterCpts.Text + 1
                                            ds.Tables(0).Rows(i).Item("Check Components") = ds.Tables(0).Rows(i).Item("Check Components") + 1
                                            adapter.UpdateCommand = New SqlCommandBuilder(adapter).GetUpdateCommand()
                                            adapter.Update(ds)
                                        End If
                                    End If

                                    'Refresh_DGV()
                                    If hasnotbeenfound <> 1 And hasnotbeenindatabase <> 1 Then Refresh_DGV() 'santo added

                                    'Exit Sub

                                    If Me.range.Text = "Kitting" And Me.CounterCpts.Text = Me.LabelQuantitycpt.Text And Me.CounterItems.Text = Me.Check.Text Then
                                        Exit Sub
                                    End If

                                    '################################# if all components have been scanned trigger camera ##############################
                                    If pindah = True Then
                                        finishVision = False

                                        If Me.CounterCpts.Text = Me.LabelQuantitycpt.Text Then

                                            'Me.CounterItems = Me.CounterItems + 1
                                            adapter2 = New SqlDataAdapter("SELECT * FROM Workstations WHERE wkname= '" & Me.workstation.SelectedValue & "'", Main.koneksi)
                                            adapter2.Fill(dsWorkstation)
                                            'apakah ini
                                            'qty label
                                            Me.quantityLabel.Text = (CInt(Me.CounterItems.Text) + 1) * CInt(Me.Quantity.Text) / CInt(Me.Check.Text)

                                            Me.testQuantity.Text = (CInt(Me.CounterItems.Text) + 1) * CInt(Me.Quantity.Text) / CInt(Me.Check.Text)
                                            'santo cek
                                            Me.cek_testQuantity.Text = (CInt(Me.CounterItems.Text) + 1) * CInt(Me.Quantity.Text) / CInt(Me.Check.Text)

                                            'UpdateScanRecords
                                            If dsWorkstation.Tables(0).Rows.Count > 0 Then
                                                If dsWorkstation.Tables(0).Rows(i).Item("onelabel") >= 1 And Me.range.Text <> "Kitting" Or dsWorkstation.Tables(0).Rows(i).Item("onelabel") >= 1 And CInt(Me.CounterItems.Text) + 1 = Me.Check.Text Then

                                                    selectedPrinter = ""
                                                    If checkAllDataOk() = False Then Exit Sub
                                                    chosenLabel = selectLabel()

                                                    If chosenLabel = "" Then Exit Sub
                                                    start = (CInt(Me.CounterItems.Text) + 1) * CInt(Me.Quantity.Text) / CInt(Me.Check.Text)
                                                    qty = (CInt(Me.CounterItems.Text) + 1) * CInt(Me.Quantity.Text) / CInt(Me.Check.Text)
                                                    printProductLabel(chosenLabel)

                                                End If
                                            End If
                                        End If
                                    End If

                                    checkTrigger()

                                    'Update [Limit]
                                    On Error GoTo 3

                                    If Me.range.Text = "Kitting" And CInt(Me.LabelQuantitycpt.Text) > CInt(Me.CounterCpts.Text) Then
                                    Else
                                        If (ds.Tables(0).Rows.Count > 0) Then
                                            If Me.range.Text = "Kitting" Then
                                                For i2 = 0 To ds.Tables(0).Rows.Count - 1
                                                    If ds.Tables(0).Rows(i2).Item("reqmts qty") - ds.Tables(0).Rows(i2).Item("Check Components") < CInt(Me.boxQty.Text) Then
                                                        Me.LabelQuantitycpt.Text = (ds.Tables(0).Rows(i2).Item("reqmts qty") - ds.Tables(0).Rows(i2).Item("Check Components")) * ds.Tables(0).Rows.Count
                                                        MsgBox("suspect 1")
                                                        ds.Tables(0).Rows(i2).Item("Limit") = ds.Tables(0).Rows(i2).Item("reqmts qty")
                                                    Else
                                                        ds.Tables(0).Rows(i2).Item("Limit") = ds.Tables(0).Rows(i2).Item("Limit") + CInt(Me.boxQty.Text)
                                                    End If

                                                    adapter.UpdateCommand = New SqlCommandBuilder(adapter).GetUpdateCommand()
                                                    adapter.Update(ds)
                                                Next
                                            Else
                                                For i3 = 0 To ds.Tables(0).Rows.Count - 1
                                                    ds.Tables(0).Rows(i3).Item("Limit") = (CInt(Me.CounterItems.Text) + 1) * ds.Tables(0).Rows(i3).Item("reqmts qty") / CInt(Me.Check.Text)
                                                    adapter.UpdateCommand = New SqlCommandBuilder(adapter).GetUpdateCommand()
                                                    adapter.Update(ds)
                                                Next
                                            End If

                                        Else

                                        End If
                                    End If

                                    '################################# if all components have been scanned trigger camera ##############################

3:
                                    '---------------> If all items have been scanned

                                    If Me.CounterItems.Text = Me.Check.Text Then
                                        Me.ComponentNo.Text = ""
                                        'Me.PPnumberEntry.Text = ""
                                        Me.PPnumberEntry.Select()

                                        Exit Sub
                                    End If

                                    Me.ComponentNo.Text = ""
                                    'Me.ComponentNo.Select()
                                    CompToQuality.Select()
                                    Exit Sub

                                    'End If

                                End If
                            End If
                        Next

                    Else

                        'If Convert.ToDecimal(LabelQuantitycpt.Text) = 0 Then
                        '    afterPPinput()
                        'Else
                        MsgBox(" All components have been scanned")

                        Me.ComponentNo.Text = ""
                        'Me.ComponentNo.Select()
                        CompToQuality.Select()
                        Exit Sub
                        'End If

                    End If

                Else
                    MsgBox("No data found")
                    Me.ComponentNo.Text = ""
                    'Me.ComponentNo.Select()
                    CompToQuality.Select()
                    Exit Sub
                End If

            Else


                MsgBox("All Breakers have been adapted")
                Me.ComponentNo.Text = ""
                'Me.ComponentNo.Select()
                CompToQuality.Select()
                Exit Sub

            End If

            If member = 1 Then
                If Me.CounterItems.Text = Me.LabelQuantityitem.Text Then
                    MsgBox("The component has already been scanned")
                Else
                    If CInt(Me.CounterCpts.Text) = CInt(Me.LabelQuantitycpt.Text) Or CInt(Me.CounterCpts.Text) = 0 Then

                        Me.PPnumberEntry.Text = ""
                        Me.PPnumberEntry.Select()
                        Me.PPnumberEntry.Text = tempPPnumber
                        Me.PPnumberEntry.Select()

                        CleanComp()
                        UpdateComponents()

                        If checkConfig() = True Then
                            updateDataVisual()
                        End If

                    Else
                        MsgBox("The component " & Me.ComponentNo.Text & " has been scanned!")
                        ComponentNo.Text = ""
                        Exit Sub
                    End If


                    Me.ComponentNo.Text = ""
                    Me.ComponentNo.Select()
                    'CompToQuality.Select()
                    Me.ComponentNo.Text = tempComponentNo
                    'CompToQuality.Text = tempComponentNo
                    Me.ComponentNo.Select()
                    'CompToQuality.Select()
                    componentNo_Scan()

                    Exit Sub
                End If

            Else
                MsgBox("Wrong Components", vbExclamation, "The component is not part of the product")
                If barcode.Length > 128 Then
                    barcode = Microsoft.VisualBasic.Left(barcode, 127)
                End If


                Dim sql As String = "INSERT INTO ScanError (workstation, technician, day, PP, BarcodeScanned) VALUES ('" & Me.workstation.SelectedValue & "','" & Me.technicianName.SelectedValue & "','" & DateTime.Now & "','" & Me.PPnumberEntry.Text & "','" & barcode & "')"
                    Dim cmd = New SqlCommand(sql, Main.koneksi)
                    cmd.ExecuteNonQuery()



            End If
            Me.ComponentNo.Text = ""
            'Me.ComponentNo.Select()
            CompToQuality.Select()

            AutoPrintProductLabelFrom.Text = "0"
            AutoPrintPackagingLabelFrom.Text = "0"
        End If
    End Sub

    Private Sub DGV_Quality()
        DataGridView5.Rows.Clear()
        Dim sqlDB As String = "SELECT * FROM [QualityIssue]"
        Dim ds As New DataSet
        adapter = New SqlDataAdapter(sqlDB, Main.koneksi)
        adapter.Fill(ds)
        If ds.Tables(0).Rows.Count > 0 Then
            DataGridView5.ColumnCount = 4
            DataGridView5.Columns(0).Name = "NO"
            DataGridView5.Columns(1).Name = "References"
            DataGridView5.Columns(2).Name = "Start Of Impact"
            DataGridView5.Columns(3).Name = "End  Of Impact"
            For r = 0 To ds.Tables(0).Rows.Count - 1
                Dim row As String() = New String() {(r + 1).ToString(), ds.Tables(0).Rows(r).Item("References").ToString(), ds.Tables(0).Rows(r).Item("Start Of Impact").ToString(), ds.Tables(0).Rows(r).Item("End Of Impact").ToString()}
                DataGridView5.Rows.Add(row)
            Next
        End If
    End Sub

    Private Sub DGV_MasterPlantCode()
        DataGridView6.Rows.Clear()
        Dim sqlDB As String = "SELECT * FROM [MasterPlantCode]"
        Dim ds As New DataSet
        adapter = New SqlDataAdapter(sqlDB, Main.koneksi)
        adapter.Fill(ds)
        If ds.Tables(0).Rows.Count > 0 Then
            DataGridView6.ColumnCount = 3
            DataGridView6.Columns(0).Name = "NO"
            DataGridView6.Columns(1).Name = "Plant Code"
            DataGridView6.Columns(2).Name = "Name Of Plant Code"
            For r = 0 To ds.Tables(0).Rows.Count - 1
                Dim row As String() = New String() {(r + 1).ToString(), ds.Tables(0).Rows(r).Item("Plant Code").ToString(), ds.Tables(0).Rows(r).Item("Name Plant Code").ToString()}
                DataGridView6.Rows.Add(row)
            Next
        End If

        'santo
        'update_combo_plantCode()

    End Sub

    Private Sub Refresh_DGV()
        DataGridView1.Rows.Clear()
        'Dim sql As String = "SELECT [ID],[Order],[Material], [Descr], [Reqmts qty], [Check Components], [Workstation] FROM Components where [Order]=" & Me.PPnumberEntry.Text & " and [Workstation]='" & Me.workstation.SelectedValue & "' order by material ASC"
        'Dim sql As String = "SELECT [ID],[Order],[Material], [Descr], [Reqmts qty], [Check Components], [Workstation] FROM Components where [Order]=" & Me.PPnumberEntry.Text & "  order by material ASC"
        'Dim sql As String = "SELECT [ID],[Order],[Material], [Descr], [Reqmts qty], [Check Components], [Workstation],[Limit] FROM Components where [Order]=" & Me.PPnumberEntry.Text & ""
        Dim sql As String = "SELECT [ID],[Order],[Material], [Descr], [Reqmts qty], [Check Components], [Workstation],[Limit] FROM Components where [Order]=" & Me.PPnumberEntry.Text & " and [Workstation]='" & Me.workstation.SelectedValue & "'"
        Dim ds As New DataSet
        adapter = New SqlDataAdapter(sql, Main.koneksi)
        adapter.Fill(ds)
        If ds.Tables(0).Rows.Count > 0 Then
            DataGridView1.ColumnCount = 5
            DataGridView1.Columns(0).Name = "Order"
            DataGridView1.Columns(1).Name = "Description"
            DataGridView1.Columns(2).Name = "Reqmts Qty"
            DataGridView1.Columns(3).Name = "Material"
            DataGridView1.Columns(4).Name = "Check Com"
            For r = 0 To ds.Tables(0).Rows.Count - 1
                If CounterItems.Text = LabelQuantityitem.Text Then
                    Dim row As String() = New String() {ds.Tables(0).Rows(r).Item("Order").ToString(), ds.Tables(0).Rows(r).Item("Descr").ToString(), ds.Tables(0).Rows(r).Item("Reqmts qty").ToString(), ds.Tables(0).Rows(r).Item("Material").ToString(), LabelQuantityitem.Text}
                    DataGridView1.Rows.Add(row)
                    If ds.Tables(0).Rows(r).Item("Check Components").ToString() = ds.Tables(0).Rows(r).Item("Limit").ToString() Then
                        DataGridView1.Rows(r).Cells(4).Style.BackColor = Color.SkyBlue
                    Else
                        DataGridView1.Rows(r).Cells(4).Style.BackColor = Color.White
                    End If
                Else
                    'If Not String.IsNullOrEmpty(ds.Tables(0).Rows(r).Item(1).ToString()) Then
                    Dim row As String() = New String() {ds.Tables(0).Rows(r).Item("Order").ToString(), ds.Tables(0).Rows(r).Item("Descr").ToString(), ds.Tables(0).Rows(r).Item("Reqmts qty").ToString(), ds.Tables(0).Rows(r).Item("Material").ToString(), ds.Tables(0).Rows(r).Item("Check Components").ToString()}
                    DataGridView1.Rows.Add(row)
                    If ds.Tables(0).Rows(r).Item("Check Components").ToString() = ds.Tables(0).Rows(r).Item("Limit").ToString() Then
                        DataGridView1.Rows(r).Cells(4).Style.BackColor = Color.SkyBlue
                    Else
                        DataGridView1.Rows(r).Cells(4).Style.BackColor = Color.White
                    End If
                End If
            Next

        End If
        Label155.Text = "Total Rows : " & DataGridView1.Rows.Count - 1
    End Sub

    Private Sub afterPPinput() ' sudah sesuai dengan VBA
        Try
            Me.dateCode.Text = dateCode2()
            Dim adapter As SqlDataAdapter
            Dim ds As New DataSet
            Dim Matt As Integer
            Dim Check As Double
            Dim cluster As Double = 0
            'Refresh_DGV()
            If hasnotbeenfound <> 1 And hasnotbeenindatabase <> 1 Then Refresh_DGV() 'santo added

            If Me.PPnumberEntry.Text = "" Or String.IsNullOrEmpty(Me.PPnumberEntry.Text) Then
                Exit Sub
            End If

            Dim sql As String = "SELECT * FROM openOrders where [Order]='" & Me.PPnumberEntry.Text & "'"
            'Dim sql As String = "SELECT o.[Order],o.[Material],o.[Descr],o.[Reqmts qty],o.[PeggedReqt] FROM openOrders o where o.[Order]='" & Me.PPnumberEntry.Text & "'"
            'Dim sql As String = "SELECT o.[Order],o.[Material],o.[Descr],o.[Reqmts qty],o.[PeggedReqt] FROM openOrders o,Componentslist c, Components Comp where o.[Order]='" & Me.PPnumberEntry.Text & "' and o.[Material]=c.[Material] and o.[Material]=Comp.[Material]"
            'Dim CkComp As String = "SELECT [Check Components] FROM Components where [Order]=" & Me.PPnumberEntry.Text & ""
            'dim sql = "SELECT [Order],[Material],[Descr],[Reqmts qty],[PeggedReqt] FROM openOrders where [Order]=@PP"
            adapter = New SqlDataAdapter(sql, Main.koneksi)
            adapter.Fill(ds)

            If ds.Tables(0).Rows.Count > 0 Then

                header.Text = ds.Tables(0).Rows(0).Item("peggedreqt").ToString
                'tambahan biar fuji tidak print
                If header.Text.Contains("BW") Then
                    CheckProductLabelPrinting.Checked = False
                    autoPrint.Checked = False
                    autoPrint2.Checked = False
                End If

                'If variable_Q = 0 Then
                If InStr(header.Text, "CSC") <> 0 Then
                    If cscinput = False Then
                        q = InputBox("CSC : How many breakers to adapt?", "CSC")
                        cscinput = True
                    End If
                ElseIf InStr(header.Text, "ATS") <> 0 Then
                    If atsinput = False Then
                        q = InputBox("ATS : How many breakers to adapt?", "ATS")
                        atsinput = True
                    End If
                End If

                variable_Q = q  ' santo added
                'End If

                'MsgBox("test:" & ds.Tables(0).Rows(0).Item("peggedreqt").ToString)

                Check = 10000000
                Matt = 0

                LabelQuantitycpt.Text = ds.Tables(0).Rows.Count.ToString

                For i = 0 To ds.Tables(0).Rows.Count - 1

                    Dim looping As Integer = i

                    If Microsoft.VisualBasic.Left(ds.Tables(0).Rows(i).Item("peggedreqt").ToString, 3) = "GCR" Then Me.header.Text = ds.Tables(0).Rows(i).Item("peggedreqt").ToString

                    If Not IsDBNull(ds.Tables(0).Rows(i).Item("material")) Or Not String.IsNullOrEmpty(ds.Tables(0).Rows(i).Item("Material").ToString) Then

                        If CInt(ds.Tables(0).Rows(i).Item("reqmts qty")) < Check Then
                            Check = CInt(ds.Tables(0).Rows(i).Item("reqmts qty"))
                        End If

                        If Matt = 0 Then Me.mat1.Text = ds.Tables(0).Rows(i).Item("material").ToString
                        If Matt = 1 Then Me.mat2.Text = ds.Tables(0).Rows(i).Item("material").ToString
                        If Matt = 2 Then Me.mat3.Text = ds.Tables(0).Rows(i).Item("material").ToString
                        If Matt = 3 Then Me.mat4.Text = ds.Tables(0).Rows(i).Item("material").ToString
                        If Matt = 4 Then Me.mat5.Text = ds.Tables(0).Rows(i).Item("material").ToString
                        If Matt = 5 Then Me.mat6.Text = ds.Tables(0).Rows(i).Item("material").ToString
                        If Matt = 6 Then Me.mat7.Text = ds.Tables(0).Rows(i).Item("material").ToString
                        If Matt = 7 Then Me.mat8.Text = ds.Tables(0).Rows(i).Item("material").ToString
                        If Matt = 8 Then Me.mat9.Text = ds.Tables(0).Rows(i).Item("material").ToString
                        If Matt = 9 Then Me.mat10.Text = ds.Tables(0).Rows(i).Item("material").ToString

                        If Matt = 0 Then Me.descr1.Text = ds.Tables(0).Rows(i).Item("descr").ToString
                        If Matt = 1 Then Me.descr2.Text = ds.Tables(0).Rows(i).Item("descr").ToString
                        If Matt = 2 Then Me.descr3.Text = ds.Tables(0).Rows(i).Item("descr").ToString
                        If Matt = 3 Then Me.descr4.Text = ds.Tables(0).Rows(i).Item("descr").ToString
                        If Matt = 4 Then Me.descr5.Text = ds.Tables(0).Rows(i).Item("descr").ToString
                        If Matt = 5 Then Me.descr6.Text = ds.Tables(0).Rows(i).Item("descr").ToString
                        If Matt = 6 Then Me.descr7.Text = ds.Tables(0).Rows(i).Item("descr").ToString
                        If Matt = 7 Then Me.descr8.Text = ds.Tables(0).Rows(i).Item("descr").ToString
                        If Matt = 8 Then Me.descr9.Text = ds.Tables(0).Rows(i).Item("descr").ToString
                        If Matt = 9 Then Me.descr10.Text = ds.Tables(0).Rows(i).Item("descr").ToString

                        If Matt = 0 Then Me.qty1.Text = ds.Tables(0).Rows(i).Item("reqmts qty").ToString
                        If Matt = 1 Then Me.qty2.Text = ds.Tables(0).Rows(i).Item("reqmts qty").ToString
                        If Matt = 2 Then Me.qty3.Text = ds.Tables(0).Rows(i).Item("reqmts qty").ToString
                        If Matt = 3 Then Me.qty4.Text = ds.Tables(0).Rows(i).Item("reqmts qty").ToString
                        If Matt = 4 Then Me.qty5.Text = ds.Tables(0).Rows(i).Item("reqmts qty").ToString
                        If Matt = 5 Then Me.qty6.Text = ds.Tables(0).Rows(i).Item("reqmts qty").ToString
                        If Matt = 6 Then Me.qty7.Text = ds.Tables(0).Rows(i).Item("reqmts qty").ToString
                        If Matt = 7 Then Me.qty8.Text = ds.Tables(0).Rows(i).Item("reqmts qty").ToString
                        If Matt = 8 Then Me.qty9.Text = ds.Tables(0).Rows(i).Item("reqmts qty").ToString
                        If Matt = 9 Then Me.qty10.Text = ds.Tables(0).Rows(i).Item("reqmts qty").ToString
                        Matt = Matt + 1
                    End If

                    If InStr(1, ds.Tables(0).Rows(i).Item("Descr").ToString, "SENSOR PLUG", 1) > 0 Or InStr(1, ds.Tables(0).Rows(i).Item("Descr").ToString, "Sensor plug", 1) > 0 Then
                        plugRef.Text = ds.Tables(0).Rows(i).Item("material").ToString
                        If Microsoft.VisualBasic.Left(ds.Tables(0).Rows(i).Item("Descr").ToString, 5) = "SENSO" Then
                            plugType.Text = Microsoft.VisualBasic.Right(ds.Tables(0).Rows(i).Item("Descr").ToString, 5)
                        Else
                            plugType.Text = Microsoft.VisualBasic.Left(ds.Tables(0).Rows(i).Item("Descr").ToString, 5)
                        End If
                    End If

                    Dim qtyComponent As Double
                    qtyComponent = ds.Tables(0).Rows(i).Item("reqmts qty")

                    If Not String.IsNullOrEmpty(ds.Tables(0).Rows(i).Item("Material").ToString) Then
                        fillReport(ds.Tables(0).Rows(i).Item("Material").ToString, qtyComponent)
                    End If

                    If ds.Tables(0).Rows(i).Item("Material").ToString = "33166" Then cluster = ds.Tables(0).Rows(i).Item("reqmts qty")

                    checkNeutral(ds.Tables(0).Rows(i).Item("material").ToString)
                Next

                If chassisRef.Text <> "" And String.IsNullOrEmpty(chassisRef.Text) = False Then
                    drawoutFixed.Text = "Drawout"
                Else
                    drawoutFixed.Text = "Fixed"
                End If

                Me.Check.Text = Check

                quantityLabel.Text = Check

                Me.cluster.Text = cluster / CInt(quantityLabel.Text)

                If Me.SDE2count.Text <> "" And Not String.IsNullOrEmpty(SDE2count.Text) Then SDE2count.Text = CInt(SDE2count.Text) / CInt(quantityLabel.Text)

                'If Me.SDE1.CheckState <> 1 Then Me.SDE1.CheckState = CInt(Me.SDE1.Text) / CInt(Me.quantityLabel.Text)

                If Me.SD.Text <> "" And Not String.IsNullOrEmpty(SD.Text) Then SD.Text = CInt(SD.Text) / CInt(quantityLabel.Text)

                'If Microsoft.VisualBasic.Left(range.Text, 2) = "NT" Then
                '    IC_OF.Text = 4
                'End If

                If sekaliOF = 0 Then
                    If Microsoft.VisualBasic.Left(Me.breakerDevice.Text, 2) <> "NS" And Microsoft.VisualBasic.Left(Me.breakerDevice.Text, 2) <> "NT" Then
                        If Me.IC_OF.Text <> "" And String.IsNullOrEmpty(Me.IC_OF.Text) = False Then
                            If CInt(Me.IC_OF.Text) > 0 Then
                                Me.IC_OF.Text = (CInt(Me.IC_OF.Text) / CInt(Me.quantityLabel.Text)) * 4
                                Me.IC_OF.Text = CInt(Me.IC_OF.Text) + 4
                                sekaliOF = sekaliOF + 1
                            End If
                        Else
                            '75
                            If Microsoft.VisualBasic.Left(Me.breakerDevice.Text, 2) <> "NS" And Not header.Text.Contains("NS") Then Me.IC_OF.Text = 4
                            sekaliOF = sekaliOF + 1
                        End If
                    Else
                        If Me.IC_OF.Text <> "" And String.IsNullOrEmpty(Me.IC_OF.Text) = False Then
                            If CInt(Me.IC_OF.Text) > 0 Then Me.IC_OF.Text = (CInt(Me.IC_OF.Text) / CInt(Me.quantityLabel.Text))
                            sekaliOF = sekaliOF + 1
                        End If
                    End If
                End If

                'If Microsoft.VisualBasic.Left(Me.breakerDevice.Text, 2) <> "NS" And Microsoft.VisualBasic.Left(Me.breakerDevice.Text, 2) <> "NT" Then
                '    'If Not String.IsNullOrEmpty(IC_OF.Text) Then
                '    Dim xxx As String = IC_OF.Text
                '    If xxx <> "" And Not String.IsNullOrEmpty(xxx) Then
                '        IC_OF.Text = (CInt(IC_OF.Text) / CInt(quantityLabel.Text)) * 4
                '        IC_OF.Text = CInt(IC_OF.Text) + 4
                '    Else
                '        If Microsoft.VisualBasic.Left(Me.breakerDevice.Text, 2) <> "NS" Then IC_OF.Text = 4
                '    End If
                '    'If xxx <> "" Or CInt(xxx) > 0 Then 'santo adited
                '    '    IC_OF.Text = (CInt(IC_OF.Text) / CInt(quantityLabel.Text)) * 4
                '    '    IC_OF.Text = CInt(IC_OF.Text) + 4
                '    'Else
                '    '    MsgBox(IC_OF.Text)
                '    '    If Microsoft.VisualBasic.Left(Me.breakerDevice.Text, 2) <> "NS" Then IC_OF.Text = 4
                '    'End If
                '    ''End If
                'Else
                '    If Microsoft.VisualBasic.Left(range.Text, 2) = "NT" Then
                '        IC_OF.Text = 4
                '    Else
                '        If Not String.IsNullOrEmpty(IC_OF.Text) Then IC_OF.Text = CInt(IC_OF.Text) / CInt(quantityLabel.Text)
                '    End If
                'End If

                'Fill the PF number
                If PF.Text <> "" And Not String.IsNullOrEmpty(PF.Text) Then PF.Text = CInt(PF.Text) / CInt(quantityLabel.Text)

                'Fill the CD number 
                If CD.Text <> "" And Not String.IsNullOrEmpty(CD.Text) Then CD.Text = CInt(CD.Text) / CInt(quantityLabel.Text)

                'Fill the CT number
                If CT.Text <> "" And Not String.IsNullOrEmpty(CT.Text) Then CT.Text = CInt(CT.Text) / CInt(quantityLabel.Text)

                'Fill the CE number
                If CE.Text <> "" And Not String.IsNullOrEmpty(CE.Text) Then CE.Text = CInt(CE.Text) / CInt(quantityLabel.Text)

                'Fill the fixing crews number
                If fixingScrew.Text <> "" And Not String.IsNullOrEmpty(fixingScrew.Text) Then fixingScrew.Text = CInt(fixingScrew.Text) / CInt(quantityLabel.Text)

                If checkLow.CheckState = 1 Or checkHigh.CheckState = 1 Or checkLT.CheckState = 1 Then checkStandard.CheckState = 0

                If selectLabel() = "" Then Exit Sub

                PP.Text = PPnumberEntry.Text
                Me.user.Text = Me.technicianName.Text
                preparePackingLabel()

                If autoPrint.CheckState = 1 Then

                    If checkAllDataOk() = False Then Exit Sub
                    '27092019
                    'If checkDuplicatePP() = True Then Exit Sub

                    'DoCmd.SelectObject acReport, Trim(selectLabel), True
                    'DoCmd.OpenReport selectLabel, acViewPreview, , , acHidden

                    'If Me.EAN13.Text = "" Or String.IsNullOrEmpty(EAN13.Text) = True Then Reports(Trim(selectLabel) Then)![text70].Visible = False

                    'DoCmd.SelectObject acReport, Trim(selectLabel), True
                    'DoCmd.PrintOut , , , , Me.quantityLabel
                    'DoCmd.Close acReport, Trim(selectLabel)
                    'DoCmd.RunCommand acCmdWindowHide

                End If
                ''remove
                'If header.Text.Contains("BW") Then CheckProductLabelPrinting.Checked = False

            Else
                If hasnotbeenfound = False Then
                    DataGridView1.Rows.Clear() 'santo added
                    MsgBox("This PP number has not been found")
                    hasnotbeenfound = 1

                End If
                Exit Sub
            End If
        Catch ex As Exception
            'MsgBox("Fail PP Input " & ex.Message)
        End Try
    End Sub

    Function checkAllDataOk() 'sudah sesuai dengan VBA

        checkAllDataOk = True

        If Me.technicianName.Text = "" Or String.IsNullOrEmpty(Me.technicianName.Text) = True Then
            checkAllDataOk = False
            Timer1.Enabled = True
            MsgBox("No Technician Name")
        End If
        If Me.technicianShortName.Text = "" Or String.IsNullOrEmpty(Me.technicianShortName.Text) = True Then
            checkAllDataOk = False
            Timer1.Enabled = True
            MsgBox("No Technician Short Name")
        End If
        If Me.PPnumberEntry.Text = "" Or String.IsNullOrEmpty(Me.PPnumberEntry.Text) = True Then
            checkAllDataOk = False
            Timer1.Enabled = True
            MsgBox("No PP Number Entry")
        End If
        If Me.header.Text = "" Or String.IsNullOrEmpty(Me.header.Text) = True Then
            checkAllDataOk = False
            Timer1.Enabled = True
            MsgBox("No Header")
        End If
        If Me.CheckProductLabelPrinting.CheckState = 1 And Me.quantityLabel.Text = 0 Or Me.CheckProductLabelPrinting.CheckState = 1 And String.IsNullOrEmpty(Me.quantityLabel.Text) Then checkAllDataOk = False

        Try
            'check the data
            If CInt(Me.StartPackingLabel.Text) > CInt(Me.labelQty.Text) Then
                'MsgBox(" Error on packaging labels quantity")
                checkAllDataOk = False
                Exit Function
            End If
        Catch ex As Exception
        End Try


        If CInt(Me.labelQty.Text) > CInt(Me.Label468.Text) Then
            'MsgBox(" Error on packaging labels quantity, number of pallets=" & Me.Label468.Text)
            Me.labelQty.Text = Me.Label468.Text
            checkAllDataOk = False
            Exit Function
        End If

        If Me.CheckProductLabelPrinting.CheckState = 1 Or Me.CheckProductLabelPrinting.CheckState = 0 And Me.checkPackagingPrinting.CheckState = 0 Then

            If CInt(Me.StartLabel.Text) > CInt(Me.quantityLabel.Text) Then
                MsgBox("Error on Product label quantity")
                checkAllDataOk = False
                Exit Function
            End If

            If CInt(Me.quantityLabel.Text) > CInt(Me.Quantity.Text) Then
                MsgBox("Error on Product label quantity, Product number =" & Me.Quantity.Text)
                'qty label
                Me.quantityLabel.Text = Me.Quantity.Text
                checkAllDataOk = False
                Exit Function
            End If

            If CInt(Me.quantityLabel.Text) > (CInt(Me.CounterItems.Text) + 1) * CInt(Me.Quantity.Text) / CInt(Me.Check.Text) Then
                'Timer1.Enabled = True
                'MsgBox("You need to scan " & Me.quantityLabel.Text - Me.CounterItems.Text & " more item(s)")
                'qty label
                Me.quantityLabel.Text = Me.CounterItems.Text
                checkAllDataOk = False
                Exit Function
            End If

        End If

        If checkAllDataOk = False Then
            Timer1.Enabled = True
            MsgBox("Some data are missing, please check your entry")
        End If

    End Function

    Function checkDuplicatePP() 'sudah sesuai dengan VBA
        Dim cmd As New SqlCommand
        checkDuplicatePP = False
        Call Main.koneksi_db()
        Dim sql As String
        sql = "Select * FROM printingRecord WHERE [PP] = '" & Me.PPnumberEntry.Text & "';"
        Dim adapter As New SqlDataAdapter(sql, Main.koneksi)
        Dim ds As New DataSet
        adapter.Fill(ds)

        If ds.Tables(0).Rows.Count = 0 Then
            'cmd = New SqlCommand("insert into printingRecord([pp],[date],[time],[user],[from],[to]) values(@pp,@date,@time,@user,@from,@to)", Main.koneksi)
            'cmd.Parameters.AddWithValue("@pp", Me.PPnumberEntry.Text)
            'cmd.Parameters.AddWithValue("@date", DateTime.Now.ToString("yyyy-MM-dd"))
            'cmd.Parameters.AddWithValue("@time", DateTime.Now.ToString("HH:mm:ss"))
            'cmd.Parameters.AddWithValue("@user", Me.technicianShortName.Text)
            'cmd.Parameters.AddWithValue("@from", Me.StartLabel.Text)
            'cmd.Parameters.AddWithValue("@to", Me.quantityLabel.Text)
            'cmd.ExecuteNonQuery()

        Else
            Dim Answer As MsgBoxResult
            'If doyouwanttoreprintit = False Then
            Answer = MsgBox("This Product Label has already been printed, do you want to reprint it?", vbQuestion + vbYesNo, "This PP has been printed before")

            If Answer = vbNo Then
                checkDuplicatePP = True
                Exit Function
            Else
                PPPackingdoyouwanttoreprintit = 1
                'cmd = New SqlCommand("insert into printingRecord([pp],[date],[time],[user],[from],[to],[Data]) values(@pp,@date,@time,@user,@from,@to,@data)", Main.koneksi)
                'cmd.Parameters.AddWithValue("@pp", Me.PPnumberEntry.Text)
                'cmd.Parameters.AddWithValue("@date", DateTime.Now.ToString("yyyy-MM-dd"))
                'cmd.Parameters.AddWithValue("@time", DateTime.Now.ToString("HH:mm:ss"))
                'cmd.Parameters.AddWithValue("@user", Me.technicianShortName.Text)
                'cmd.Parameters.AddWithValue("@from", Me.StartLabel.Text)
                'cmd.Parameters.AddWithValue("@to", Me.quantityLabel.Text)
                'cmd.Parameters.AddWithValue("@data", LoginForm.strHostName & " - " & Application.ProductVersion)
                'cmd.ExecuteNonQuery()

                checkDuplicatdePP = False
            End If
            'doyouwanttoreprintit = True
            'End If
        End If
    End Function

    Sub fillReport(description As String, qty As Double)
        description = Replace(description, "'", " ", 1)
        Call Main.koneksi_db()
        Dim str As String
        str = "SELECT * FROM BOM where [BOM number]='" & description & "'"
        Dim adapter As New SqlDataAdapter(str, Main.koneksi)
        Dim ds As New DataSet
        adapter.Fill(ds)

        If ds.Tables(0).Rows.Count = 0 Then Exit Sub

        If ds.Tables(0).Rows(0).Item("Category").ToString = "breaker" Then
            breakerRef.Text = ds.Tables(0).Rows(0).Item("Material").ToString
            breakerPole.Text = ds.Tables(0).Rows(0).Item("poles").ToString
            If ds.Tables(0).Rows(0).Item("range").ToString <> "MVS" Then breakerType.Text = ds.Tables(0).Rows(0).Item("Type").ToString()
        End If

        If ds.Tables(0).Rows(0).Item("Category").ToString = "label" Then
            breakerType.Text = ds.Tables(0).Rows(0).Item("Type").ToString()
            If Mid(ds.Tables(0).Rows(0).Item("description").ToString, 6, 1) <> "0" Then
                breakerDevice.Text = Microsoft.VisualBasic.Left(ds.Tables(0).Rows(0).Item("description").ToString, 5)
            Else
                breakerDevice.Text = Microsoft.VisualBasic.Left(ds.Tables(0).Rows(0).Item("description").ToString, 6)
            End If
        End If

        If ds.Tables(0).Rows(0).Item("Category").ToString = "micrologic" Then
            micrologicRef.Text = ds.Tables(0).Rows(0).Item("Material").ToString
            micrologicType.Text = ds.Tables(0).Rows(0).Item("Type").ToString
        End If

        If ds.Tables(0).Rows(0).Item("Category").ToString = "chassis" Then
            chassisRef.Text = ds.Tables(0).Rows(0).Item("Material").ToString
        End If

        If ds.Tables(0).Rows(0).Item("Category").ToString = "top connection" Then
            topConnect.Text = ds.Tables(0).Rows(0).Item("Type").ToString
        End If

        If ds.Tables(0).Rows(0).Item("Category").ToString = "bottom connection" Then
            bottomConnect.Text = ds.Tables(0).Rows(0).Item("Type").ToString
        End If

        If ds.Tables(0).Rows(0).Item("Category").ToString = "MCH" Then
            MCH.Text = ds.Tables(0).Rows(0).Item("Type").ToString
            If Microsoft.VisualBasic.Left(ds.Tables(0).Rows(0).Item("range").ToString, 2) = "NS" Then
                XF.Text = ds.Tables(0).Rows(0).Item("Type").ToString
                MX.Text = ds.Tables(0).Rows(0).Item("Type").ToString
            End If
        End If

        If ds.Tables(0).Rows(0).Item("Category").ToString = "XF" Then
            XF.Text = ds.Tables(0).Rows(0).Item("Type").ToString
            XFtype.Text = ds.Tables(0).Rows(0).Item("poles").ToString
        End If

        If ds.Tables(0).Rows(0).Item("Category").ToString = "MX" Then
            MX.Text = ds.Tables(0).Rows(0).Item("Type").ToString
            MXtype.Text = ds.Tables(0).Rows(0).Item("poles").ToString
        End If


        If ds.Tables(0).Rows(0).Item("Category").ToString = "MN" Then
            MN.Text = ds.Tables(0).Rows(0).Item("Type").ToString
        End If

        If ds.Tables(0).Rows(0).Item("Category").ToString = "MX2" Then
            MX2.Text = ds.Tables(0).Rows(0).Item("Type").ToString
        End If

        If ds.Tables(0).Rows(0).Item("Category").ToString = "SDE2" Then
            SDE2.Text = ds.Tables(0).Rows(0).Item("Type").ToString
            SDE2count.Text = qty
        End If

        If Microsoft.VisualBasic.Left(ds.Tables(0).Rows(0).Item("range").ToString, 2) = "NW" Or ds.Tables(0).Rows(0).Item("range").ToString = "NT" Or ds.Tables(0).Rows(0).Item("range").ToString = "MVS" Then
            If ds.Tables(0).Rows(0).Item("Category").ToString = "breaker" Then
                If ds.Tables(0).Rows(0).Item("Type").ToString = "HA" Or ds.Tables(0).Rows(0).Item("Type").ToString = "NA" Then
                    SDE1.CheckState = 0
                Else
                    SDE1.CheckState = 1
                End If
            End If
        Else
            If ds.Tables(0).Rows(0).Item("Category").ToString = "SDE1" Then SDE1.CheckState = 1
        End If
        'Dim tes As String = ds.Tables(0).Rows(0).Item("Category").ToString

        If ds.Tables(0).Rows(0).Item("Category").ToString = "OF" Then
            If IC_OF.Text = "4" Or IC_OF.Text = "8" Or IC_OF.Text = "12" Then

            Else
                'IC_OF.Text = CInt(IC_OF.Text) + qty
                IC_OF.Text = qty
            End If
            'IC_OF.Text = qty
        End If

        If ds.Tables(0).Rows(0).Item("Category").ToString = "OF individual" Then
            IC_OF.Text = qty
        End If

        If ds.Tables(0).Rows(0).Item("Category").ToString = "fixingScrews" Then
            fixingScrew.Text = qty
        End If

        If ds.Tables(0).Rows(0).Item("Category").ToString = "PF" Then
            PF.Text = qty
        End If

        If ds.Tables(0).Rows(0).Item("Category").ToString = "SD" Then
            SD.Text = qty
        End If


        If ds.Tables(0).Rows(0).Item("Category").ToString = "remote reset" Then
            Me.remoteReset.Text = ds.Tables(0).Rows(0).Item("Type").ToString
        End If

        If ds.Tables(0).Rows(0).Item("Category").ToString = "VBP" Then
            checkVBP.CheckState = 1
        End If

        If ds.Tables(0).Rows(0).Item("Category").ToString = "VCPO" Then
            checkVCPO.CheckState = 1
        End If

        If ds.Tables(0).Rows(0).Item("Category").ToString = "VSPO" Then
            checkVSPO.CheckState = 1
        End If

        If ds.Tables(0).Rows(0).Item("Category").ToString = "VSPDchassis" Then
            checkVSPDchassis.CheckState = 1
        End If

        If ds.Tables(0).Rows(0).Item("Category").ToString = "VSPDall" Then
            checkVSPDall.CheckState = 1
        End If

        If ds.Tables(0).Rows(0).Item("Category").ToString = "VPEC" Then
            checkVPEC.CheckState = 1
        End If

        'tidak ada
        If ds.Tables(0).Rows(0).Item("Category").ToString = "VPOC" Then
            checkVPOC.CheckState = 1
        End If

        'tidak ada
        If ds.Tables(0).Rows(0).Item("Category").ToString = "DAE" Then
            checkDAE.CheckState = 1
        End If

        'tidak ada
        If ds.Tables(0).Rows(0).Item("Category").ToString = "VDC" Then
            checkVDC.CheckState = 1
        End If

        'tidak ada
        If ds.Tables(0).Rows(0).Item("Category").ToString = "IPA" Then
            checkIPA.CheckState = 1
        End If

        'tidak ada
        If ds.Tables(0).Rows(0).Item("Category").ToString = "IPBO" Then
            checkIPBO.CheckState = 1
        End If

        'tidak ada
        If ds.Tables(0).Rows(0).Item("Category").ToString = "VICV" Then
            checkVICV.CheckState = 1
        End If

        If ds.Tables(0).Rows(0).Item("Category").ToString = "CDM" Then
            checkCDM.CheckState = 1
        End If

        If ds.Tables(0).Rows(0).Item("Category").ToString = "CB" Then
            checkCB.CheckState = 1
        End If

        If ds.Tables(0).Rows(0).Item("Category").ToString = "CDP" Then
            checkCDP.CheckState = 1
        End If

        If ds.Tables(0).Rows(0).Item("Category").ToString = "CP" Then
            checkCP.CheckState = 1
        End If

        'tidak ada
        If ds.Tables(0).Rows(0).Item("Category").ToString = "OP" Then
            checkOP.CheckState = 1
        End If

        If ds.Tables(0).Rows(0).Item("Category").ToString = "PhaseBarrier" Then
            checkPhaseBarrier.CheckState = 1
        End If

        If ds.Tables(0).Rows(0).Item("Category").ToString = "Rotary" Then
            checkRotary.CheckState = 1
        End If

        'tidak ada
        If ds.Tables(0).Rows(0).Item("Category").ToString = "EF" Then
            checkEF.CheckState = 1
        End If

        If ds.Tables(0).Rows(0).Item("Category").ToString = "CE" Then
            CE.Text = qty
        End If

        If ds.Tables(0).Rows(0).Item("Category").ToString = "CD" Then
            CD.Text = qty
        End If

        If ds.Tables(0).Rows(0).Item("Category").ToString = "CT" Then
            CT.Text = qty
        End If

        If ds.Tables(0).Rows(0).Item("Category").ToString = "PTE" Then
            checkPTE.CheckState = 1
        End If

        If ds.Tables(0).Rows(0).Item("Category").ToString = "R" Then
            checkR.CheckState = 1
        End If

        If ds.Tables(0).Rows(0).Item("Category").ToString = "Rr" Then
            checkRr.CheckState = 1
        End If

        If ds.Tables(0).Rows(0).Item("Category").ToString = "M2C" Then
            checkM2C.CheckState = 1
        End If

        If ds.Tables(0).Rows(0).Item("Category").ToString = "COM" Then
            checkCOM.CheckState = 1
        End If

        If ds.Tables(0).Rows(0).Item("Category").ToString = "ATS" Then
            checkATS.CheckState = 1
        End If

        If ds.Tables(0).Rows(0).Item("Category").ToString = "AD" Then
            checkAD.CheckState = 1
        End If

        If ds.Tables(0).Rows(0).Item("Category").ToString = "battery" Then
            checkBattery.CheckState = 1
        End If

        If ds.Tables(0).Rows(0).Item("Category").ToString = "externalCT" Then
            checkexternalCT.CheckState = 1
        End If

        If ds.Tables(0).Rows(0).Item("Category").ToString = "standard" Then
            checkStandard.CheckState = 1
        End If

        If ds.Tables(0).Rows(0).Item("Category").ToString = "low" Then
            checkLow.CheckState = 1
        End If

        If ds.Tables(0).Rows(0).Item("Category").ToString = "high" Then
            checkHigh.CheckState = 1
        End If

        'tidak ada
        If ds.Tables(0).Rows(0).Item("Category").ToString = "LT" Then
            checkLT.CheckState = 1
        End If

    End Sub

    Private Sub boxQty_TextChanged() Handles boxQty.TextChanged
        convert_box_pallet()
    End Sub

    Private Sub Check157_Click(sender As Object, e As EventArgs) Handles checkPackagingPrinting2.Click
        If checkPackagingPrinting2.CheckState = 1 Then checkPackagingPrinting.CheckState = 1
        If checkPackagingPrinting2.CheckState = 0 Then checkPackagingPrinting.CheckState = 0
    End Sub

    Private Sub Check158_Click(sender As Object, e As EventArgs) Handles CheckProductLabelPrinting.Click
        If CheckProductLabelPrinting.CheckState = 1 Then CheckProductLabelPrinting2.CheckState = 1
        If CheckProductLabelPrinting.CheckState = 0 Then CheckProductLabelPrinting2.CheckState = 0
    End Sub

    Private Sub checkAssembledSingapore_CheckedChanged(sender As Object, e As EventArgs) Handles checkAssembledSingapore.CheckedChanged
        If checkAssembledSingapore.CheckState = 0 Then
            madeInSingapore.Text = ""
        Else
            madeInSingapore.Text = "Assembled by Schneider Electric in Singapore"
        End If
    End Sub
    Public Sub checkPallet_CheckedChanged() Handles checkPallet.CheckedChanged
        'Dim qty As Double
        'Dim pallet As Double
        'Dim box As Double
        'Dim dec As Double

        'qty = Me.Quantity.Text
        'pallet = Me.palletQty.Text
        box = Me.boxQty.Text

        convert_box_pallet()

        ''If CheckBox "Max quantity per pallet" not checked
        'If Me.checkPallet.Checked = 0 Then
        '    Me.checkBox.Checked = 1
        '    Me.checkPallet.Checked = 0
        '    Me.boxQty.Font = New Font(palletQty.Font, FontStyle.Bold)
        '    Me.palletQty.Font = New Font(palletQty.Font, FontStyle.Regular)

        'Else

        '    Me.checkBox.Checked = 0
        '    Me.checkPallet.Checked = 1
        '    Me.boxQty.Font = New Font(palletQty.Font, FontStyle.Regular)
        '    Me.palletQty.Font = New Font(palletQty.Font, FontStyle.Bold)

        'End If
        'Me.labelQty2 = Me.labelQty
        'Me.Label468.Text = Me.labelQty.Text
        'Me.StartPackingLabel.Text = 1
        'Me.StartPackingLabel2.Text = 1
    End Sub
    Private Sub convert_box_pallet()

        Dim qty As Double
        Dim pallet As Double
        Dim box As Double
        Dim dec As Double

        Dim vqty As Double
        Dim vpallet As Double
        Dim vbox As Double
        'MsgBox("checkBoX")
        qty = Double.TryParse(Quantity.Text, vqty)
        pallet = Double.TryParse(palletQty.Text, vpallet)
        box = Double.TryParse(boxQty.Text, vbox)

        qty = vqty
        pallet = vpallet
        box = vbox


        'MsgBox("checkBoX")
        'If CheckBox "Max quantity per box" not checked
        Try
            If checkBox.CheckState = 0 Then
                checkBox.CheckState = 0
                checkPallet.CheckState = 1
                boxQty.Font = New Font(boxQty.Font, FontStyle.Regular)
                palletQty.Font = New Font(palletQty.Font, FontStyle.Bold)

                'Number of Label to print
                dec = (qty / pallet) - Math.Floor(qty / pallet)
                If dec = 0 Then
                    labelQty.Text = Math.Floor(qty / pallet)
                End If
                If dec > 0 Then
                    labelQty.Text = Math.Floor(qty / pallet) + 1
                End If


                'If CheckBox "Max quantity per box" checked
            Else

                checkBox.CheckState = 1
                checkPallet.CheckState = 0
                boxQty.Font = New Font(boxQty.Font, FontStyle.Bold)
                palletQty.Font = New Font(palletQty.Font, FontStyle.Regular)

                'Number of Label to print
                dec = (qty / box) - Math.Floor(qty / box)
                If dec = 0 Then labelQty.Text = Math.Floor(qty / box)
                If dec > 0 Then labelQty.Text = Math.Floor(qty / box) + 1

            End If
        Catch ex As Exception

        End Try

        'labelQty2.Text = labelQty.Text
        Label468.Text = labelQty.Text
        StartPackingLabel.Text = 1
        StartPackingLabel2.Text = 1
    End Sub


    Public Sub checkBox_CheckedChanged() Handles checkBox.CheckedChanged
        convert_box_pallet()
    End Sub

    Private Sub UpdateComponents()
        ' Item = Circuit Breaker to be adapted
        ' Component = Component that is part of the breaker according to the BOM
        Call Main.koneksi_db()
        Dim sql As String
        Dim sql2 As String
        Dim sqlPP As String
        Dim count As Integer
        Dim count2 As Integer
        Dim SensorCounter As Integer
        Dim countDelete As Integer
        Dim ds As New DataSet
        Dim ds2 As New DataSet
        Dim dsPP As New DataSet
        Dim adapter As New SqlDataAdapter
        Dim sekali As Boolean = False

        SensorCounter = 0
        count = 0
        count2 = 0

        If Me.PPnumberEntry.Text = "" Or String.IsNullOrEmpty(Me.PPnumberEntry.Text) = True Then Exit Sub

        ' Avoid blank lines in the components table
        If Me.workstation.Text = "" Or String.IsNullOrEmpty(Me.workstation.Text) = True Then
            MsgBox("Please choose the workstation")
            Exit Sub
        End If

        sql = "Select [Order],[Material],[Descr],[Reqmts qty],[PeggedReqt] FROM openOrders WHERE [Order] = '" & Me.PPnumberEntry.Text & "' ;"
        adapter = New SqlDataAdapter(sql, Main.koneksi)
        adapter.Fill(ds)

        If ds.Tables(0).Rows.Count = 0 Then Exit Sub

        'Call Main.koneksi_db()
        sql = "SELECT * FROM [PPList] WHERE [Order] = '" & Me.PPnumberEntry.Text & "'"
        adapter = New SqlDataAdapter(sql, Main.koneksi)
        adapter.Fill(dsPP)

        'Call Main.koneksi_db()
        'sql = "SELECT * FROM [PPList] WHERE [Order] = '" & Me.PPnumberEntry.Text & "'"
        'adapter = New SqlDataAdapter(sql, Main.koneksi)
        'adapter.Fill(dsPP)

        'If ds.Tables(0).Rows.Count = 0 Then Exit Sub

        On Error GoTo 1
        Dim value1 As Long
        Long.TryParse(dsPP.Tables(0).Rows(0).Item("Material").ToString, value1)
        If dsPP.Tables(0).Rows.Count > 0 Then
            If value1 > 29999 And value1 < 56000 Or value1 > 56999 Then
                If hitung = 0 Then
                    'If qtyperboxinput1 = False Then
                    qtyboxinput = InputBox("What is the max qty per box")
                    Me.boxQty.Text = qtyboxinput
                    qtyperboxinput1 = True
                    qtyperboxinput2 = True
                    'dipindah kesini pada tanggal 19 desember 2019
                    Me.checkBox.CheckState = 1
                    checkBox_CheckedChanged()
                    Me.Check.Text = Me.labelQty2.Text
                    'End If
                    hitung = hitung + 1
                End If
                'awalnya disini tanggal 19 desember 2019
                'Me.checkBox.CheckState = 1
                'checkBox_CheckedChanged()
                'Me.Check.Text = Me.labelQty2.Text
            End If
        End If
1:
        'On Error GoTo 0
        'If ds.Tables(0).Rows.Count > 0 Then
        '    'MsgBox(dsPP.Tables(0).Rows(0).Item("PrCtr").ToString)
        '    If String.IsNullOrEmpty(dsPP.Tables(0).Rows(0).Item("Material").ToString) = False And String.IsNullOrEmpty(dsPP.Tables(0).Rows(0).Item("PrCtr").ToString) = False Then
        '        If dsPP.Tables(0).Rows(0).Item("PrCtr") = 10 Or Microsoft.VisualBasic.Left(dsPP.Tables(0).Rows(0).Item("Material").ToString, 3) = "LV4" And
        '            dsPP.Tables(0).Rows(0).Item("PrCtr") = 4 Or dsPP.Tables(0).Rows(0).Item("Material") = "GCR_KEYLOCKS_MM" Then
        '            If qtyperboxinput2 = False Then
        '                Me.boxQty.Text = InputBox("what is the max qty per box")
        '                qtyperboxinput2 = True
        '            End If
        '            Me.checkBox.CheckState = 1
        '            checkBox_CheckedChanged()
        '            Me.Check.Text = Me.labelQty2.Text
        '        End If
        '    End If
        'End If

        On Error GoTo 0
        If dsPP.Tables(0).Rows.Count > 0 Then
            'MsgBox(dsPP.Tables(0).Rows(0).Item("PrCtr").ToString)
            If String.IsNullOrEmpty(dsPP.Tables(0).Rows(0).Item("Material").ToString) = False And String.IsNullOrEmpty(dsPP.Tables(0).Rows(0).Item("PrCtr").ToString) = False Then
                If dsPP.Tables(0).Rows(0).Item("PrCtr") = 10 Or Microsoft.VisualBasic.Left(dsPP.Tables(0).Rows(0).Item("Material").ToString, 3) = "LV4" And
                    dsPP.Tables(0).Rows(0).Item("PrCtr") = 4 Or dsPP.Tables(0).Rows(0).Item("Material") = "GCR_KEYLOCKS_MM" Then
                    'MsgBox(hitung.ToString )
                    'If qtyperboxinput2 = False Then
                    If hitung = 0 Then
                        qtyboxinput = InputBox("what is the max qty per box")
                        Me.boxQty.Text = qtyboxinput
                        qtyperboxinput2 = True
                        qtyperboxinput1 = True
                        'dipindah kesini pada tanggal 19 desember 2019
                        Me.checkBox.CheckState = 1
                        checkBox_CheckedChanged()
                        Me.Check.Text = Me.labelQty2.Text
                        hitung = hitung + 1
                    End If
                    'End If
                    'awalnya disini tanggal 19 desember 2019
                    'Me.checkBox.CheckState = 1
                    'checkBox_CheckedChanged()
                    'Me.Check.Text = Me.labelQty2.Text
                End If
            End If
        End If

        'Check if components has not been scanned yet
        Dim dsItemScan As New DataSet
        sql = "SELECT [ItemScanned] FROM [ScanRecords] WHERE [Order]= '" & Me.PPnumberEntry.Text & "';"
        adapter = New SqlDataAdapter(sql, Main.koneksi)
        adapter.Fill(dsItemScan)

        'Initialization
        If dsItemScan.Tables(0).Rows.Count > 0 Then
            If Not IsDBNull(dsItemScan.Tables(0).Rows(0).Item("Itemscanned")) Then
                Me.CounterItems.Text = dsItemScan.Tables(0).Rows(0).Item("Itemscanned").ToString 'qty label
                Me.quantityLabel.Text = CInt(dsItemScan.Tables(0).Rows(0).Item("Itemscanned").ToString) * CInt(Me.Quantity.Text) / CInt(Me.Check.Text)
                Me.testQuantity.Text = CInt(dsItemScan.Tables(0).Rows(0).Item("Itemscanned").ToString) * CInt(Me.Quantity.Text) / CInt(Me.Check.Text)
                'santo cek
                Me.cek_testQuantity.Text = CInt(dsItemScan.Tables(0).Rows(0).Item("Itemscanned").ToString) * CInt(Me.Quantity.Text) / CInt(Me.Check.Text)
            End If
        Else
            Me.CounterItems.Text = 0
            Me.quantityLabel.Text = 0
            Me.testQuantity.Text = 0
            'santo_cek
            Me.cek_testQuantity.Text = 0
        End If

        wk = Me.workstation.SelectedValue

        'balik:
        'Order by [description] Is used to delete the highest Sensor Plug rate
        'sql2 = "Select * FROM [Components] WHERE [Order] = " & Me.PPnumberEntry.Text & " ORDER BY [Descr];"
        sql2 = "Select * FROM [Components] WHERE [Order] = " & Me.PPnumberEntry.Text & " ORDER BY [Descr];"
        Dim adapter2 = New SqlDataAdapter(sql2, Main.koneksi)

        adapter2.Fill(ds2)

        'Check If the PP has been recorded Or Not
        If ds2.Tables(0).Rows.Count = 0 Then
            'insert the New components
            'DoCmd.SetWarnings False
            'Dim ds3 As New DataSet
            'Dim sqlCount As String
            'sqlCount = "Select [openOrders].[Order], [openOrders].[Material], [openOrders].[Descr], [openOrders].[Reqmts qty], [Componentslist].[Code] FROM [Componentslist] INNER JOIN [openOrders] ON [Componentslist].[Material]=[openOrders].[Material] WHERE [openOrders].[Order] = '" & Me.PPnumberEntry.Text & "'"
            'Dim adapter3 = New SqlDataAdapter(sqlCount, Main.koneksi)
            'adapter3.Fill(ds3)
            'MsgBox(ds3.Tables(0).Rows.Count)

            'adapter3.Fill(ds3)

            'If ds3.Tables(0).Rows.Count <> ds2.Tables(0).Rows.Count Or ds2.Tables(0).Rows.Count = 0 Then
            Dim ter = New SqlDataAdapter("INSERT INTO [Components] ([Order],[Material],[Descr],[Reqmts qty],[Code]) Select [openOrders].[Order], [openOrders].[Material], [openOrders].[Descr], [openOrders].[Reqmts qty], [Componentslist].[Code] FROM [Componentslist] INNER JOIN [openOrders] ON [Componentslist].[Material]=[openOrders].[Material] WHERE [openOrders].[Order] = '" & Me.PPnumberEntry.Text & "'", Main.koneksi)
            ter.SelectCommand.ExecuteNonQuery()
            'End If

            'recovering the items scanned
            If dsItemScan.Tables(0).Rows.Count > 0 Then
                For i = 0 To ds2.Tables(0).Rows.Count - 1
                    ds2.Tables(0).Rows(i).Item("Check Components") = ds2.Tables(0).Rows(i).Item("reqmts qty") / CInt(Me.Check.Text) * dsItemScan.Tables(0).Rows(i).Item("Itemscanned")
                    adapter2.UpdateCommand = New SqlCommandBuilder(adapter2).GetUpdateCommand()
                    adapter2.Update(ds2)
                Next
            End If
            'End If
            'GoTo balik
            'End If
            'End If
            'DoCmd.SetWarnings True

            'sql2 = "Select * FROM [Components] WHERE [Order] = " & Me.PPnumberEntry.Text & " and [Workstation] is null ORDER BY [Descr];"
            'adapter2 = New SqlDataAdapter(sql2, Main.koneksi)

            'adapter2.Fill(ds2)

            'Update components table 
        ElseIf ds2.Tables(0).Rows.Count > 0 Then
            count = 0
            count2 = 0
            SensorCounter = 0
            Dim tes_a As Integer
            tes_a = ds2.Tables(0).Rows.Count

            On Error Resume Next

            For i = 0 To tes_a - 1
                'For i = 0 To ds2.Tables(0).Rows.Count - 1
                'If ds2.Tables(0).Rows.Count - countDelete = i Then Exit For
                'If i >= 17 Then Exit For 'ini aku yang tambahin
                'Delete the highest SENSOR PLUG line (except for ATS and CSC)
                'On Error GoTo next_step

                If Microsoft.VisualBasic.Right(ds2.Tables(0).Rows(i).Item("Descr").ToString, 11) = "SENSOR PLUG" Then
                    If InStr(Me.header.Text, "CSC") = 0 And InStr(Me.header.Text, "ATS") = 0 Then
                        If SensorCounter >= 1 Then
                            ds2.Tables(0).Rows(i).Delete()
                            countDelete = i
                            Continue For
                        Else
                            SensorCounter = SensorCounter + 1
                        End If
                    End If
                End If

                'next_step:
                'If ds2.Tables(0).Rows(0).Item("date") < ds2.Tables(0).Rows(i).Item("date") Then
                '    ds2.Tables(0).Rows(i).Delete()
                '    Continue For
                'End If

                If countDelete <> 0 Then
                    If sekali = False Then
                        ds2.Tables(0).Rows(i).Item("Workstation") = wk
                        count = count + CInt(ds2.Tables(0).Rows(i).Item("reqmts qty").ToString)
                        count2 = count2 + CInt(ds2.Tables(0).Rows(i).Item("Check Components").ToString)
                        adapter2.UpdateCommand = New SqlCommandBuilder(adapter2).GetUpdateCommand()
                        adapter2.Update(ds2)
                        sekali = True
                    End If
                End If

                'MsgBox(i) 

                ds2.Tables(0).Rows(i).Item("Workstation") = wk

                'counting the number of components
                count = count + CInt(ds2.Tables(0).Rows(i).Item("reqmts qty").ToString)
                count2 = count2 + CInt(ds2.Tables(0).Rows(i).Item("Check Components").ToString)

                adapter2.UpdateCommand = New SqlCommandBuilder(adapter2).GetUpdateCommand()
                adapter2.Update(ds2)

            Next

            If Me.range.Text = "Kitting" Then
                Me.LabelQuantitycpt.Text = CInt(Me.boxQty.Text) * ds2.Tables(0).Rows.Count 'terakhir di command
                For i = 0 To ds2.Tables(0).Rows.Count - 1
                    If CInt(Me.CounterItems.Text) + 1 < Me.Check.Text Then
                        ds2.Tables(0).Rows(i).Item("Limit") = (CInt(Me.CounterItems.Text) + 1) * CInt(Me.boxQty.Text)
                        adapter2.UpdateCommand = New SqlCommandBuilder(adapter2).GetUpdateCommand()
                        adapter2.Update(ds2)
                        'MsgBox("A")
                    Else
                        ds2.Tables(0).Rows(i).Item("Limit") = ds2.Tables(0).Rows(i).Item("reqmts qty")
                        adapter2.UpdateCommand = New SqlCommandBuilder(adapter2).GetUpdateCommand()
                        adapter2.Update(ds2)
                        'MsgBox("B")
                    End If
                Next
            Else
                'suspect 3
                Me.LabelQuantitycpt.Text = count / CInt(Me.Check.Text)  'terakhir di command
                'Me.LabelQuantitycpt.Text = count / CInt(Me.Quantity.Text) 'terakhir di command
                ' MsgBox(LabelQuantitycpt.Text)
                For i = 0 To ds2.Tables(0).Rows.Count - 1
                    ds2.Tables(0).Rows(i).Item("Limit") = (CInt(Me.CounterItems.Text) + 1) * CInt(ds2.Tables(0).Rows(i).Item("Reqmts qty")) / CInt(Me.Check.Text)
                    adapter2.UpdateCommand = New SqlCommandBuilder(adapter2).GetUpdateCommand()
                    adapter2.Update(ds2)
                    'MsgBox("C")
                Next
            End If

            'Else
            '    Dim ds3 As New DataSet
            '    Dim sqlCek As String
            '    'sqlCek = "Select * FROM [Components] WHERE [Order] = " & Me.PPnumberEntry.Text & " and [Workstation] = '" & Me.workstation.SelectedValue & "' ORDER BY [Descr];"
            '    sqlCek = "Select * FROM [Components] WHERE [Order] = " & Me.PPnumberEntry.Text & " ORDER BY [Descr];"
            '    Dim adapter3 = New SqlDataAdapter(sqlCek, Main.koneksi)

            '    adapter3.Fill(ds3)

            '    If ds3.Tables(0).Rows.Count <> 0 Then
            '        count = 0
            '        count2 = 0
            '        For i = 0 To ds3.Tables(0).Rows.Count - 1
            '            'Delete the highest SENSOR PLUG line (except for ATS and CSC)
            '            If Microsoft.VisualBasic.Right(ds3.Tables(0).Rows(i).Item("Descr").ToString, 11) = "SENSOR PLUG" Then
            '                If InStr(Me.header.Text, "CSC") = 0 And InStr(Me.header.Text, "ATS") = 0 Then
            '                    If SensorCounter >= 1 Then
            '                        ds.Tables(0).Rows(i).Delete()
            '                    Else
            '                        SensorCounter = SensorCounter + 1
            '                    End If
            '                End If
            '            End If

            '            ds3.Tables(0).Rows(i).Item("Workstation") = wk

            '            'counting the number of components
            '            count = count + CInt(ds3.Tables(0).Rows(i).Item("reqmts qty").ToString)
            '            count2 = count2 + CInt(ds3.Tables(0).Rows(i).Item("Check Components").ToString)

            '            adapter3.UpdateCommand = New SqlCommandBuilder(adapter3).GetUpdateCommand()
            '            adapter3.Update(ds3)
            '        Next

            '        If Me.range.Text = "Kitting" Then
            '            Me.LabelQuantitycpt.Text = CInt(Me.boxQty.Text) * ds3.Tables(0).Rows.Count 'terakhir di command
            '            For i = 0 To ds3.Tables(0).Rows.Count - 1
            '                If CInt(Me.CounterItems.Text) + 1 < Me.Check.Text Then
            '                    ds3.Tables(0).Rows(i).Item("Limit") = (CInt(Me.CounterItems.Text) + 1) * CInt(Me.boxQty.Text)
            '                    adapter3.UpdateCommand = New SqlCommandBuilder(adapter3).GetUpdateCommand()
            '                    adapter3.Update(ds3)
            '                    'MsgBox("A")
            '                Else
            '                    ds3.Tables(0).Rows(i).Item("Limit") = ds3.Tables(0).Rows(i).Item("reqmts qty")
            '                    adapter3.UpdateCommand = New SqlCommandBuilder(adapter3).GetUpdateCommand()
            '                    adapter3.Update(ds3)
            '                    'MsgBox("B")
            '                End If
            '            Next
            '        Else
            '            'suspect 3
            '            Me.LabelQuantitycpt.Text = count / CInt(Me.Check.Text)  'terakhir di command
            '            'Me.LabelQuantitycpt.Text = count / CInt(Me.Quantity.Text) 'terakhir di command
            '            ' MsgBox(LabelQuantitycpt.Text)
            '            For i = 0 To ds3.Tables(0).Rows.Count - 1
            '                ds3.Tables(0).Rows(i).Item("Limit") = (CInt(Me.CounterItems.Text) + 1) * CInt(ds3.Tables(0).Rows(i).Item("Reqmts qty")) / CInt(Me.Check.Text)
            '                adapter3.UpdateCommand = New SqlCommandBuilder(adapter3).GetUpdateCommand()
            '                adapter3.Update(ds3)
            '                'MsgBox("C")
            '            Next
            '        End If
            '    Else
            '        Me.LabelQuantitycpt.Text = 0       'santo command
            '        MsgBox("No Components found")      'santo command  
            '    End If
        End If

        'Refresh_DGV()
        If hasnotbeenfound <> 1 And hasnotbeenindatabase <> 1 Then Refresh_DGV() 'santo added

        Me.LabelQuantityitem.Text = Me.Check.Text
        Me.CounterCpts.Text = count2 - CInt(Me.CounterItems.Text) * CInt(Me.LabelQuantitycpt.Text)  'terakhir di command

        Me.StartLabel.Text = 1

        'Me.ComponentNo.Select()
        CompToQuality.Select()
    End Sub

    Sub updateDataVisual()
        visualCheck = 0
        actualQty = 0
        targetQty = 0

        Dim queryCheck As String
        Dim queryUpdate As String
        Dim ds As New DataSet

        queryCheck = "SELECT * FROM [Components] WHERE [Material] = '29450' AND [Order] = " & Me.PPnumberEntry.Text & " And [Workstation] = '" & Me.workstation.Text & "';"
        Dim adapter = New SqlDataAdapter(queryCheck, Main.koneksi)
        adapter.Fill(ds)

        If ds.Tables(0).Rows.Count > 0 And checkConfig() = True Then
            visualCheck = 1
        Else
            visualCheck = 2
        End If


        'DoCmd.SetWarnings False
        Call koneksi_db()
        queryUpdate = "UPDATE [tblVisualCheck] SET [PP_No] = " & Me.PPnumberEntry.Text & ", [Material_Name] = '" & Me.header.Text & "',[Visual_Check] = " & visualCheck & ", [Actual_Qty] = " & CInt(Me.CounterItems.Text) & ", [Target_Qty] = " & CInt(Me.LabelQuantityitem.Text) & ", [flag] = 1, [Tech_Name] = '" & Me.technicianName.Text & "';"
        Dim cmd = New SqlDataAdapter(queryUpdate, Main.koneksi)
        cmd.SelectCommand.ExecuteNonQuery()


        'DoCmd.RunSQL queryUpdate
        'DoCmd.SetWarnings True

        If CInt(Me.CounterCpts.Text) = CInt(Me.LabelQuantitycpt.Text) And CInt(Me.CounterCpts.Text) > 0 Then
            '    MsgBox "SDFDS"
            updateTriggerVisual()
        End If


    End Sub

    Sub updateTriggerVisual()


        Me.Text593.Text = "PLEASE PUT BREAKER TO VISION JIG AND PRESS TRIGGER BUTTON TO SCAN THE LABEL!"

        Dim queryUpdate As String

        'DoCmd.SetWarnings False

        queryUpdate = "UPDATE [tblVisualCheck] SET [trigger] = 1;"

        'DoCmd.RunSQL queryUpdate
        'DoCmd.SetWarnings True
        scanComponents = True

    End Sub

    Sub checkStatusVision()
        Dim queryCheck As String
        Dim queryUpdate As String
        Dim ds As New DataSet

        queryCheck = "SELECT * FROM [tblVisualCheck] WHERE [Status_Vision] = 'PASS';"
        Dim adapter = New SqlDataAdapter(queryCheck, Main.koneksi)
        adapter.Fill(ds)

        If ds.Tables(0).Rows.Count > 0 Then

            'Me.Label590.Caption = "PLEASE SCAN ANOTHER BREAKER!"
            Me.Text593.Text = "PLEASE SCAN ANOTHER BREAKER!"
            finishVision = True
            checkTrigger()

            'DoCmd.SetWarnings False
            queryUpdate = "UPDATE [tblVisualCheck] SET [Status_Vision] = 'IDLE';"

            'DoCmd.RunSQL queryUpdate
            'DoCmd.SetWarnings True
            scanComponents = False
        End If
    End Sub

    Function dateCode2()
        Dim VBAWeekNum As Integer = DatePart(DateInterval.WeekOfYear, Date.Today, FirstDayOfWeek.Monday, FirstWeekOfYear.FirstFourDays)

        If Len(VBAWeekNum) = 1 Then VBAWeekNum = "0" & VBAWeekNum

        dateCode2 = Date.Now.Year & "-W" & VBAWeekNum & "-" & DateAndTime.Weekday(DateTime.Now, vbMonday)

    End Function
    'untuk LA4202
    Function dateCode3()
        Dim VBAWeekNum As Integer = DatePart(DateInterval.WeekOfYear, Date.Today, FirstDayOfWeek.Monday, FirstWeekOfYear.FirstFourDays)

        If Len(VBAWeekNum) = 1 Then VBAWeekNum = "0" & VBAWeekNum

        dateCode3 = Date.Now.Year & VBAWeekNum & DateAndTime.Weekday(DateTime.Now, vbMonday)

    End Function
    'untuk Ruby
    Function dateCode4()
        Dim VBAWeekNum As Integer = DatePart(DateInterval.WeekOfYear, Date.Today, FirstDayOfWeek.Monday, FirstWeekOfYear.FirstFourDays)

        If Len(VBAWeekNum) = 1 Then VBAWeekNum = "0" & VBAWeekNum

        dateCode4 = Date.Now.Year.ToString().Substring(2, 2) & VBAWeekNum & DateAndTime.Weekday(DateTime.Now, vbMonday)

    End Function

    Sub updateTraceability()
        Try
            Dim ds2 As New DataSet
            Dim dss3 As New DataSet
            Dim sql2a = "SELECT * FROM [PPList] WHERE [Order] = '" & Me.PPnumberEntry.Text & "'"
            Dim adapter2 = New SqlDataAdapter(sql2a, Main.koneksi)
            adapter2.Fill(ds2)

            'Problem save data
            'Dim strArrGiDate() As String
            'Dim stringGiDate As String = ds2.Tables(0).Rows(0).Item("GI date").ToString()
            'strArrGiDate = stringGiDate.Split(".")
            'Dim CombineGiDate As String = strArrGiDate(1) + "." + strArrGiDate(0) + "." + strArrGiDate(2)

            'Dim CombineGiDate As String = ds2.Tables(0).Rows(0).Item("GI date").ToString()

            'Dim strArrSchFini() As String
            'Dim stringSchFini As String = ds2.Tables(0).Rows(0).Item("Scheduled finish").ToString()
            'strArrSchFini = stringSchFini.Split(".")
            'Dim CombineSchFini As String = strArrSchFini(1) + "." + strArrSchFini(0) + "." + strArrSchFini(2)

            'Dim CombineSchFini As String = ds2.Tables(0).Rows(0).Item("Scheduled finish").ToString()

            Dim sql = "INSERT INTO [dbo].[ProductionOrders](
            [GI Date]
            ,[PP]
            ,[Quantityadapted]
            ,[Quantity]
            ,[workstation]
            ,[range]
            ,[Customer]
            ,[City]
            ,[SchedFinishDate]
            ,[Material]
            ,[Description]
            ,[Name1]
            ,[SO no]
            ,[Item]
            ,[Entity]) 
            values(
            '" & ds2.Tables(0).Rows(0).Item("GI date").ToString() & "'
            ,'" & ds2.Tables(0).Rows(0).Item("order") & "'
            ,'" & CounterItems.Text & "'
            ,'" & ds2.Tables(0).Rows(0).Item("Item quantity") & "'
            ,'" & workstation.SelectedValue & "'
            ,'" & ds2.Tables(0).Rows(0).Item("range") & "'
            ,'" & ds2.Tables(0).Rows(0).Item("Customer") & "'
            ,'" & ds2.Tables(0).Rows(0).Item("City") & "'
            ,'" & ds2.Tables(0).Rows(0).Item("Scheduled finish").ToString() & "'
            ,'" & ds2.Tables(0).Rows(0).Item("Material") & "'
            ,'" & ds2.Tables(0).Rows(0).Item("Description") & "'
            ,'" & ds2.Tables(0).Rows(0).Item("Name 1") & "'
            ,'" & ds2.Tables(0).Rows(0).Item("SO no") & "'
            ,'" & ds2.Tables(0).Rows(0).Item("Item") & "'
            ,'" & ds2.Tables(0).Rows(0).Item("Entity") & "')"

            Dim ter = New SqlDataAdapter(sql, Main.koneksi)
            ter.SelectCommand.ExecuteNonQuery()

            'cast('" & Convert.ToDateTime(CombineGiDate).ToString("MM/dd/yyyy") & "' as datetime)
            ',cast('" & Convert.ToDateTime(CombineSchFini).ToString("MM/dd/yyyy") & "' as datetime)

            'If ter.SelectCommand.ExecuteNonQuery() Then
            '    MessageBox.Show("Berhasil")
            'Else
            '    MessageBox.Show("Gagal")
            'End If

            Dim sql2 = "INSERT INTO [dbo].[DailyProduction]([Employee],[WorkingSection]) values(@emp,@ws)"
            Dim ter2 = New SqlDataAdapter(sql2, Main.koneksi)
            ter2.SelectCommand.Parameters.AddWithValue("@emp", technicianName.SelectedValue)
            ter2.SelectCommand.Parameters.AddWithValue("@ws", workstation.SelectedValue)
            ter2.SelectCommand.ExecuteNonQuery()

            Dim adapters3 = New SqlDataAdapter("SELECT * FROM [Components] WHERE [Order] = '" & Me.PPnumberEntry.Text & "'", Main.koneksi)
            adapters3.Fill(dss3)
            If dss3.Tables(0).Rows.Count > 0 Then
                For i3 = 0 To dss3.Tables(0).Rows.Count - 1
                    Dim sql3 = "INSERT INTO [dbo].[ComponentOrders]
                    ([Order]
                    ,[Code]
                    ,[Barcode]
                    ,[Datecode]
                    ,[Material]
                    ,[Workstation]
                    ,[Descr]
                    ,[Reqmts qty]
                    ,[Check Components])
                    values(@order,@code,@barcode,@datecode,@mat,@ws,@ds,@req,@check)"
                    Dim ter3 = New SqlDataAdapter(sql3, Main.koneksi)
                    ter3.SelectCommand.Parameters.AddWithValue("@order", dss3.Tables(0).Rows(i3).Item("order"))
                    ter3.SelectCommand.Parameters.AddWithValue("@code", dss3.Tables(0).Rows(i3).Item("code"))
                    ter3.SelectCommand.Parameters.AddWithValue("@barcode", dss3.Tables(0).Rows(i3).Item("barcode"))
                    ter3.SelectCommand.Parameters.AddWithValue("@datecode", dss3.Tables(0).Rows(i3).Item("datecode"))
                    ter3.SelectCommand.Parameters.AddWithValue("@mat", dss3.Tables(0).Rows(i3).Item("material"))
                    ter3.SelectCommand.Parameters.AddWithValue("@ws", dss3.Tables(0).Rows(i3).Item("workstation"))
                    ter3.SelectCommand.Parameters.AddWithValue("@ds", dss3.Tables(0).Rows(i3).Item("descr"))
                    ter3.SelectCommand.Parameters.AddWithValue("@req", dss3.Tables(0).Rows(i3).Item("reqmts qty"))
                    ter3.SelectCommand.Parameters.AddWithValue("@check", dss3.Tables(0).Rows(i3).Item("check components"))
                    ter3.SelectCommand.ExecuteNonQuery()
                Next
            End If

            Dim queryGiOn = "Update ProductionOrders SET [On Time]= 'On Time' where convert(datetime, [GI Date], 103) > convert(datetime, [FinishDate], 103) and [On Time] is null and [FinishDate] is not null and [pp] = '" & Me.PPnumberEntry.Text & "'"
            Dim queryschon = "update ProductionOrders set [On TimeM]= 'On Time' where convert(datetime, [SchedFinishDate], 103) > convert(datetime, [FinishDate], 103) and [On TimeM] is null and [FinishDate] is not null and [pp] = '" & Me.PPnumberEntry.Text & "'"

            Dim querygilate = "update ProductionOrders set [Late]= 'Late' where convert(datetime, [GI Date], 103) <= convert(datetime, [FinishDate], 103) and [Late] is null and [FinishDate] is not null and [pp] = '" & Me.PPnumberEntry.Text & "'"
            Dim queryschlate = "update ProductionOrders set [LateM]= 'Late' where convert(datetime, [SchedFinishDate], 103) <= convert(datetime, [FinishDate], 103) and [LateM] is null and [FinishDate] is not null and [pp] = '" & Me.PPnumberEntry.Text & "'"

            adapters2 = New SqlDataAdapter(queryGiOn, Main.koneksi)
            adapters2.SelectCommand.ExecuteNonQuery()

            adapters2 = New SqlDataAdapter(queryschon, Main.koneksi)
            adapters2.SelectCommand.ExecuteNonQuery()

            adapters2 = New SqlDataAdapter(querygilate, Main.koneksi)
            adapters2.SelectCommand.ExecuteNonQuery()

            adapters2 = New SqlDataAdapter(queryschlate, Main.koneksi)
            adapters2.SelectCommand.ExecuteNonQuery()
        Catch ex As Exception
            'MessageBox.Show(ex.ToString)
            'Log_data("Update Trecebility:" & ex.ToString)
        End Try
    End Sub

    Sub checkTrigger()
        scanComponents = False
        'Dim dateCode As String
        'Dim workRS As DAO.Recordset
        'Dim workRS2 As DAO.Recordset
        'Dim rs As DAO.Recordset

        'Dim sql As String
        'Dim sql2 As String
        'Dim sql3 As String
        'Dim ctrl As Control
        'Dim report As report
        'Dim member As Integer
        'Dim dateCode As String
        'Dim barcode As String
        'Dim chosenLabel As String
        'Dim start As Long
        'Dim qty As Long

        '------> If all components have been scanned
        If Me.CounterCpts.Text = Me.LabelQuantitycpt.Text Then
            If visualCheck = 1 And finishVision = False Then
                updateTriggerVisual()
                Exit Sub
            End If

            If CInt(Me.CounterItems.Text) < CInt(Me.LabelQuantityitem.Text) - 1 Then

                Me.PPnumberEntry.Text = ""
                Me.PPnumberEntry.Select()
                Me.PPnumberEntry.Text = tempPPnumber
                Me.PPnumberEntry.Select()
                '        PPnumberEntry_Scan
                CleanComp()
                UpdateComponents()
            End If

            'Update counters
            Me.CounterItems.Text = CInt(Me.CounterItems.Text) + 1

            'If Me.CounterItems.Text = Me.LabelQuantityitem.Text Then
            '    Dim adapter As New SqlDataAdapter
            '    Dim query = "UPDATE [PPList] SET [FinishDate] = '" & DateTime.Now.ToString("dd.MM.yyyy") & "' WHERE [Order] = '" & Me.PPnumberEntry.Text & "' and FinishDate is null"
            '    adapter = New SqlDataAdapter(query, Main.koneksi)
            '    adapter.SelectCommand.ExecuteNonQuery()

            '    Dim queryGiOn = "Update PPList SET [On time]= 'On Time' WHERE convert(datetime, [GI date], 103) > convert(datetime, [FinishDate], 103) AND [On Time] is null and [FinishDate] is not null and [Order] = '" & Me.PPnumberEntry.Text & "'"
            '    Dim querySchOn = "Update PPList SET [On timeM]= 'On Time' WHERE convert(datetime, [Scheduled finish], 103) > convert(datetime, [FinishDate], 103) AND [On TimeM] is null and [FinishDate] is not null and [Order] = '" & Me.PPnumberEntry.Text & "'"

            '    Dim queryGiLate = "Update PPList SET [Late]= 'Late' WHERE convert(datetime, [GI date], 103) <= convert(datetime, [FinishDate], 103) AND [Late] is null and [FinishDate] is not null and [Order] = '" & Me.PPnumberEntry.Text & "'"
            '    Dim querySchLate = "Update PPList SET [LateM]= 'Late' WHERE convert(datetime, [Scheduled finish], 103) <= convert(datetime, [FinishDate], 103) AND [Late] is null and [FinishDate] is not null and [Order] = '" & Me.PPnumberEntry.Text & "'"

            '    adapter = New SqlDataAdapter(queryGiOn, Main.koneksi)
            '    adapter.SelectCommand.ExecuteNonQuery()

            '    adapter = New SqlDataAdapter(querySchOn, Main.koneksi)
            '    adapter.SelectCommand.ExecuteNonQuery()

            '    adapter = New SqlDataAdapter(queryGiLate, Main.koneksi)
            '    adapter.SelectCommand.ExecuteNonQuery()

            '    adapter = New SqlDataAdapter(querySchLate, Main.koneksi)
            '    adapter.SelectCommand.ExecuteNonQuery()
            'End If

            UpdateScanRecords()

            Me.CounterCpts.Text = 0

        End If '<--- end if start all component has been scanned
    End Sub

    Sub CleanComp() 'sudah sesuai dengan VBA
        Try
            Dim adapter As New SqlDataAdapter
            Dim ds As New DataSet
            adapter = New SqlDataAdapter("SELECT * FROM [Components] WHERE [Workstation] = '" & workstation.SelectedValue & "'", Main.koneksi)
            adapter.Fill(ds)

            If ds.Tables(0).Rows.Count > 0 Then
                For i = 0 To ds.Tables(0).Rows.Count - 1
                    If ds.Tables(0).Rows(i).Item("reqmts qty") > ds.Tables(0).Rows(i).Item("Check Components") And ds.Tables(0).Rows(i).Item("Check Components") <> 0 Then
                        adapter = New SqlDataAdapter("UPDATE [Components] SET [Workstation] = NULL WHERE Workstation = '" & workstation.SelectedValue & "'", Main.koneksi)
                        adapter.SelectCommand.ExecuteNonQuery()
                        Exit Sub
                    ElseIf ds.Tables(0).Rows(i).Item("Check Components") = 0 Then
                        adapter = New SqlDataAdapter("Delete FROM Components WHERE Workstation = '" & workstation.SelectedValue & "'", Main.koneksi)
                        adapter.SelectCommand.ExecuteNonQuery()
                        Exit Sub
                    End If
                Next
            Else
                Exit Sub
            End If

            Dim cmd = New SqlCommand("delete from [Components] where [workstation] = '" & workstation.SelectedValue & "'", Main.koneksi)
            cmd.ExecuteNonQuery()

        Catch ex As Exception
            MsgBox("Clean Component " + ex.Message)
        End Try
    End Sub

    Sub UpdateScanRecords()
        Try

            If header.Text.Contains("BW") Then Btn_Fuji_Side_label_Click(sender, e)
            'Dim sql As String
            Dim ds As New DataSet
            Dim ds2 As New DataSet
            Dim dsFuji As New DataSet
            Dim dsCheckComp As New DataSet

            Dim adapter = New SqlDataAdapter("SELECT [ID],[ItemScanned] FROM [ScanRecords] WHERE [Order]='" & Me.PPnumberEntry.Text & "'", Main.koneksi)
            adapter.Fill(ds)

            If ds.Tables(0).Rows.Count > 0 Then
                ds.Tables(0).Rows(0).Item("ItemScanned") = Me.CounterItems.Text
                adapter.UpdateCommand = New SqlCommandBuilder(adapter).GetUpdateCommand
                adapter.Update(ds)

                Dim adapter2 = New SqlDataAdapter("SELECT [ID],[ItemScanned] FROM [ScanRecords] WHERE [Order]='" & Me.PPnumberEntry.Text & "'", Main.koneksi)
                adapter2.Fill(ds2)
                If ds2.Tables(0).Rows.Count > 0 Then
                    If ds2.Tables(0).Rows(0).Item("ItemScanned").ToString = LabelQuantityitem.Text Then
                        updateTraceability()
                    End If
                End If
            Else
                Dim sql As String = "insert into [ScanRecords]([Order],[ItemScanned]) values('" & Me.PPnumberEntry.Text & "','" & Me.CounterItems.Text & "')"
                Dim cmd As New SqlCommand(sql, Main.koneksi)
                cmd.ExecuteNonQuery()

                Dim adapter2 = New SqlDataAdapter("SELECT [ID],[ItemScanned] FROM [ScanRecords] WHERE [Order]='" & Me.PPnumberEntry.Text & "'", Main.koneksi)
                adapter2.Fill(ds2)
                If ds2.Tables(0).Rows.Count > 0 Then
                    If ds2.Tables(0).Rows(0).Item("ItemScanned").ToString = LabelQuantityitem.Text Then
                        updateTraceability()
                    End If
                End If
            End If

            Dim adapterFuji = New SqlDataAdapter("SELECT * FROM [MasterFuji] WHERE [FCSRef]='" & Me.header.Text & "'", Main.koneksi)
            adapterFuji.Fill(dsFuji)
            If dsFuji.Tables(0).Rows.Count > 0 Then
                'santo fuji edit
                Dim QrCodeFujiSideLabel = Fuji_QR_Product_Label.Text
                'Fuji_QR_Product_Label.Text = QrCodeFujiSideLabel
                Dim adapterCheckComp = New SqlDataAdapter("select * from ComponentsFuji where [order]='" & PPnumberEntry.Text & "' and RefFuji='" & header.Text & "' and QrCodeFuji='" & QrCodeFujiSideLabel & "'", Main.koneksi)
                adapterCheckComp.Fill(dsCheckComp)
                If dsCheckComp.Tables(0).Rows.Count = 0 Then
                    Dim queryInsert = "insert into [ComponentsFuji]([order],[RefFuji],[WorkStation],[Check Components], 
                    [2ndScan],[QRCodeFuji]) values ('" & Me.PPnumberEntry.Text & "',(select DISTINCT PeggedReqt from openOrders 
                    where [order]='" & Me.PPnumberEntry.Text & "'),'" & Me.workstation.Text & "',0,0,'" & QrCodeFujiSideLabel & "')"
                    Dim adapterInsert = New SqlDataAdapter(queryInsert, Main.koneksi)
                    adapterInsert.SelectCommand.ExecuteNonQuery()
                    'MessageBox.Show("Print Fuji Side Label")
                    'Btn_Fuji_Side_label_Click(sender, e)
                End If
            End If

        Catch ex As Exception
            MsgBox("Scan Record Fail " & ex.Message)
        End Try
    End Sub

    Sub loadTestData() 'sudah sesuai dengan VBA

        Dim sql As String
        Dim ds As New DataSet

        If Me.testPP.Text = "" Or String.IsNullOrEmpty(Me.testPP.Text) = True Then Exit Sub

        sql = "Select distinct * " &
        "FROM PPList " &
        "WHERE [Order] = '" & Me.testPP.Text & "';"
        Dim adapter = New SqlDataAdapter(sql, Main.koneksi)
        adapter.Fill(ds)

        If ds.Tables(0).Rows.Count = 0 Then
            If hasnotbeenindatabase = False Then
                DataGridView1.Rows.Clear() 'santo added
                MsgBox("This PP number is not in the database")
                hasnotbeenindatabase = 1
            End If
            Exit Sub
        ElseIf ds.Tables(0).Rows.Count > 1 Then
            If duplicateSQOO = False Then
                MsgBox("There is duplicate records in the PP database (SQOO), please check")
                duplicateSQOO = True
            End If
            GoTo loncat
        Else
loncat:
            Me.testSO.Text = ds.Tables(0).Rows(0).Item("SO no").ToString
            Me.testSOitem.Text = ds.Tables(0).Rows(0).Item("Item").ToString
            Me.txt_SO_item.Text = ds.Tables(0).Rows(0).Item("Item").ToString
            Me.testMaterial.Text = ds.Tables(0).Rows(0).Item("Material").ToString
            If InStr(Me.header.Text, "CSC") <> 0 Or InStr(Me.header.Text, "ATS") <> 0 Then
                Me.testQuantity.Text = q
                'santo cek
                Me.cek_testQuantity.Text = q
            Else
                Me.testQuantity.Text = ds.Tables(0).Rows(0).Item("item Quantity").ToString
                'santo cek
                Me.cek_testQuantity.Text = ds.Tables(0).Rows(0).Item("item Quantity").ToString
            End If
            Me.StartTestQuantity.Text = 1
            Me.testCustomer.Text = ds.Tables(0).Rows(0).Item("Name 1").ToString
            Me.testCustomerCode.Text = ds.Tables(0).Rows(0).Item("customer").ToString
            If Not IsDBNull(ds.Tables(0).Rows(0).Item("Req dlv dt")) Then
                Me.testCRD.Text = Convert_To_Date_Format(ds.Tables(0).Rows(0).Item("Req dlv dt"))
                'Me.testCRD.Text = Date.FromOADate(ds.Tables(0).Rows(0).Item("Req dlv dt")).ToString("dd/MM/yyyy")
            Else
                If ReqDelvdt = 0 Then
                    MsgBox("Req dlv dt is Empty")
                    ReqDelvdt = 1
                End If
            End If
            If Not IsDBNull(ds.Tables(0).Rows(0).Item("Created On")) Then
                'Me.testCreationDate.Text = Date.FromOADate(ds.Tables(0).Rows(0).Item("Created On")).ToString("dd/MM/yyyy")
                Me.testCreationDate.Text = Convert_To_Date_Format(ds.Tables(0).Rows(0).Item("Created On"))
            Else
                MsgBox("Created On is Empty")
            End If
            Me.testCustPO.Text = ds.Tables(0).Rows(0).Item("Purchase order number").ToString
            Me.testCustPOitem.Text = ds.Tables(0).Rows(0).Item("POitem").ToString
            Me.testDescription.Text = ds.Tables(0).Rows(0).Item("description").ToString

            If Microsoft.VisualBasic.Left(Me.testDescription.Text, 2) = "NW" Then
                Me.productRange.Text = "NW Masterpact"
                Me.breakerDevice.Text = Microsoft.VisualBasic.Left(Me.testDescription.Text, 4)
                If Mid(Me.testDescription.Text, 10, 1) = "D" Then Me.drawoutFixed.Text = "Drawout"
                If Mid(Me.testDescription.Text, 10, 1) = "F" Then Me.drawoutFixed.Text = "Fixed"

            End If

            If Microsoft.VisualBasic.Left(Me.testDescription.Text, 2) = "NS" And Mid(Me.testDescription.Text, 3, 1) <> "X" Then
                Me.productRange.Text = "Compact NS"
                If String.IsNullOrEmpty(Mid(Me.testDescription.Text, 6, 1)) <> "0" Then
                    Me.breakerDevice.Text = Microsoft.VisualBasic.Left(Me.testDescription.Text, 5)
                Else
                    '75
                    'Me.breakerDevice.Text = Microsoft.VisualBasic.Left(Me.testDescription.Text, 6)
                    Me.breakerDevice.Text = Microsoft.VisualBasic.Left(Me.testDescription.Text, 6)
                    Dim cek_last As String = Me.breakerDevice.Text
                    Dim sss As Boolean = Char.IsDigit(cek_last(5))
                    If Not sss Then
                        cek_last = cek_last.Substring(0, cek_last.Length - 1)
                        'MessageBox.Show(cek_last)
                        Me.breakerDevice.Text = cek_last

                    End If
                End If

                If Mid(Me.testDescription.Text, 10, 1) = "D" Then Me.drawoutFixed.Text = "Drawout"
                If Mid(Me.testDescription.Text, 10, 1) = "F" Then Me.drawoutFixed.Text = "Fixed"
                If Mid(Me.testDescription.Text, 11, 1) = "D" Then Me.drawoutFixed.Text = "Drawout"
                If Mid(Me.testDescription.Text, 11, 1) = "F" Then Me.drawoutFixed.Text = "Fixed"
            End If

            If Microsoft.VisualBasic.Left(Me.testDescription.Text, 2) = "NT" Then
                Me.productRange.Text = "NT Masterpact"
                Me.breakerDevice.Text = Microsoft.VisualBasic.Left(Me.testDescription.Text, 4)
                If Mid(Me.testDescription.Text, 10, 1) = "D" Then Me.drawoutFixed.Text = "Drawout"
                If Mid(Me.testDescription.Text, 10, 1) = "F" Then Me.drawoutFixed.Text = "Fixed"
            End If


            If Microsoft.VisualBasic.Left(Me.testDescription.Text, 3) = "MVS" Then
                Me.productRange.Text = "MVS Easypact"
                Me.breakerDevice.Text = Microsoft.VisualBasic.Left(Me.testDescription.Text, 5)
                If Mid(Me.testDescription.Text, 9, 1) = "D" Then Me.drawoutFixed.Text = "Drawout"
                If Mid(Me.testDescription.Text, 9, 1) = "F" Then Me.drawoutFixed.Text = "Fixed"
            End If

            'Jinlong Req
            If Microsoft.VisualBasic.Left(Me.testDescription.Text, 3) = "MTZ" Then
                If testMaterial.Text.Contains("MTZ1") Then
                    Me.productRange.Text = "MTZ1 Masterpact"
                ElseIf testMaterial.Text.Contains("MTZ2") Then
                    Me.productRange.Text = "MTZ2 Masterpact"
                ElseIf testMaterial.Text.Contains("MTZ3") Then
                    Me.productRange.Text = "MTZ3 Masterpact"
                Else
                    Me.productRange.Text = "MTZ Masterpact"
                End If
            End If

            If Microsoft.VisualBasic.Left(Me.testDescription.Text, 3) = "EV" Then
                Me.productRange.Text = "MVS Easypact"
            End If


            If Me.testCountry.Text = "" Then
                Dim dsCountry As New DataSet
                sql = "Select * " &
                "FROM customerDatabase " &
                "WHERE [customer code] = '" & Me.testCustomerCode.Text & "';"

                adapter = New SqlDataAdapter(sql, Main.koneksi)
                adapter.Fill(dsCountry)
                Try
                    If dsCountry.Tables(0).Rows.Count = 1 Then
                        Me.testCountry.Text = dsCountry.Tables(0).Rows(0).Item("country")
                    End If
                Catch ex As Exception
                    If CustomerCountryNotFound = 0 Then MsgBox("Customer Country Not Found ")
                    CustomerCountryNotFound = 1
                End Try
            End If
        End If
    End Sub

    Private Sub TestMaterial_TextChanged(sender As Object, e As EventArgs) Handles testMaterial.TextChanged
        Me.productRange.Text = ""
        Me.breakerDevice.Text = ""

        If Microsoft.VisualBasic.Left(Me.Material.Text, 2) = "NW" Then Me.productRange.Text = "NW Masterpact"
        If Microsoft.VisualBasic.Left(Me.Material.Text, 2) = "NS" Then Me.productRange.Text = "Compact NS"
        If Microsoft.VisualBasic.Left(Me.Material.Text, 2) = "NT" Then Me.productRange.Text = "NT Masterpact"

        Me.breakerDevice.Text = Microsoft.VisualBasic.Left(Me.Material.Text, 4)
    End Sub

    Private Sub TestPP_TextChanged(sender As Object, e As EventArgs) Handles testPP.TextChanged
        clean()
        Me.PPnumberEntry.Text = Me.testPP.Text
        Me.PP.Text = Me.testPP.Text

        afterPPinput()
        loadTestData()
        preparePackingLabel()
        UpdateComponents()
    End Sub

    Function checkDuplicateTest() 'Check whether the Test report has already been printed or not

        Dim ds As New DataSet

        checkDuplicateTest = False

        If CInt(Me.StartTestQuantity.Text) > CInt(Me.testQuantity.Text) Then
            MsgBox("Error on Test report quantity")
            checkDuplicateTest = True
            Exit Function
        End If

        If CInt(Me.testQuantity.Text) > CInt(Me.CounterItems.Text) * CInt(Me.Quantity.Text) / CInt(Me.Check.Text) Then
            MsgBox("You need to scan " & CInt(Me.testQuantity.Text) - CInt(Me.CounterItems.Text) * CInt(Me.Quantity.Text) / CInt(Me.Check.Text) & " more item(s)")
            Me.testQuantity.Text = CInt(Me.CounterItems.Text) * CInt(Me.Quantity.Text) / CInt(Me.Check.Text)
            'santo cek
            Me.cek_testQuantity.Text = CInt(Me.CounterItems.Text) * CInt(Me.Quantity.Text) / CInt(Me.Check.Text)

            checkDuplicateTest = True
            Exit Function
        End If

        Dim adapter = New SqlDataAdapter("Select * FROM [printingRecordTest] WHERE [PP] = '" & Me.testPP.Text & "'", Main.koneksi)
        adapter.Fill(ds)

        If ds.Tables(0).Rows.Count = 0 Then
            adapter.InsertCommand.Parameters.AddWithValue("PP", Me.PPnumberEntry.Text)
            adapter.InsertCommand.Parameters.AddWithValue("DATE", DateTime.Now.ToString("yyyy-MM-dd"))
            adapter.InsertCommand.Parameters.AddWithValue("TIME", DateTime.Now.ToString("HH:mm:ss"))
            adapter.InsertCommand.Parameters.AddWithValue("user", Me.technicianShortName.Text)
            adapter.InsertCommand.Parameters.AddWithValue("From", Me.StartLabel.Text)
            adapter.InsertCommand.Parameters.AddWithValue("To", Me.quantityLabel.Text)
            adapter.InsertCommand.ExecuteNonQuery()

        Else
            Dim Answer As MsgBoxResult
            Answer = MsgBox("This report has already been printed (TEST), do you want to reprint it?", vbQuestion + vbYesNo, "This report has been printed before")

            If Answer = vbNo Then
                checkDuplicateTest = True
                Exit Function

            Else

                adapter.UpdateCommand.Parameters.AddWithValue("PP", Me.PPnumberEntry.Text)
                adapter.UpdateCommand.Parameters.AddWithValue("DATE", DateTime.Now.ToString("yyyy-MM-dd"))
                adapter.UpdateCommand.Parameters.AddWithValue("TIME", DateTime.Now.ToString("HH:mm:ss"))
                adapter.UpdateCommand.Parameters.AddWithValue("user", Me.technicianShortName.Text)
                adapter.UpdateCommand.Parameters.AddWithValue("From", Me.StartLabel.Text)
                adapter.UpdateCommand.Parameters.AddWithValue("To", Me.quantityLabel.Text)
                adapter.UpdateCommand.ExecuteNonQuery()

                checkDuplicateTest = False
            End If
        End If
    End Function

    Function checkDuplicatePPpacking() ' Check if the Packing label has already been printed out 'sudah sesuai dengan VBA

        Dim ds As New DataSet

        checkDuplicatePPpacking = False

        Dim adapter = New SqlDataAdapter("Select * FROM [printingRecordPacking] WHERE [PP] = '" & Me.PP.Text & "'", Main.koneksi)
        adapter.Fill(ds)

        'If not
        If ds.Tables(0).Rows.Count = 0 Then
            'Dim sqlNew As String = "insert into [printingRecordPacking] ([PP],[print date],[print time],[User]) values ('" & Me.PP.Text & "',getDate(),getDate(),'" & Me.user.Text & "')"
            'Dim cmdNew = New SqlCommand(sqlNew, Main.koneksi)
            'cmdNew.ExecuteNonQuery 
            Exit Function
        Else
            'If PPPackingdoyouwanttoreprintit = False Then
            Dim Answer As MsgBoxResult
            Answer = MsgBox("The packaging label has already been printed, do you want to reprint it?", vbQuestion + vbYesNo, "Packaging label previously printed")

            If Answer = vbNo Then
                checkDuplicatePPpacking = True
                Exit Function
            Else
                doyouwanttoreprintit = 1
                adapter = New SqlDataAdapter("update [printingRecordPacking] set PP = '" & Me.PP.Text & "',[print date]=getDate(),[print time]=getDate(),[User]='" & Me.user.Text & "' where PP='" & Me.PP.Text & "'", Main.koneksi)
                adapter.SelectCommand.ExecuteNonQuery()
                checkDuplicatePPpacking = False
            End If
            'PPPackingdoyouwanttoreprintit = 1
            'End If
        End If
    End Function

    Private Sub PP_TextChanged(sender As Object, e As EventArgs) Handles PP.TextChanged
        clean()
        Me.PPnumberEntry.Text = Me.PP.Text
        Me.testPP.Text = Me.PP.Text


        afterPPinput()
        loadTestData()
        preparePackingLabel() 'Useless /included in afterPPinput
        UpdateComponents()
        boxQty_TextChanged()
    End Sub

    Function Convert_To_Date_Format(dateNotOK As String)
        Dim childAgeAsdouble As Double
        Dim value As String
        If Integer.TryParse(dateNotOK, childAgeAsdouble) Then
            value = Date.FromOADate(dateNotOK).ToString("dd/MM/yyyy")
        Else
            If InStr(dateNotOK, ".") Then
                value = dateNotOK.ToString().Replace(".", "/")
            Else
                value = dateNotOK.ToString
            End If
        End If
        Return value
    End Function

    Private Sub preparePackingLabel() 'sudah sesuai dengan VBA

        Dim ds As New DataSet
        Dim adapter As New SqlDataAdapter

        'Clean previous data
        Me.Quantity.Text = 0
        Me.country.Text = ""
        Me.CRD.Text = ""
        Me.creationDate.Text = ""
        Me.localexport.Text = ""
        Me.customer.Text = ""
        Me.Customer_code.Text = ""
        Me.Material.Text = ""
        Me.custMaterial.Text = ""
        Me.SO.Text = ""
        Me.SO_item.Text = ""
        Me.palletQty.Text = 0
        Me.boxQty.Text = 0
        Me.labelQty.Text = 0
        Me.checkBox.CheckState = 0
        Me.checkPallet.CheckState = 0
        Me.custPO.Text = ""
        Me.sysconfID.Text = ""
        Me.sysconfLine.Text = ""
        Me.packingDescription.Text = ""
        Me.custPOitem.Text = ""
        Me.palletQty.Font = New Font(palletQty.Font, FontStyle.Regular)
        Me.boxQty.Font = New Font(boxQty.Font, FontStyle.Regular)
        Me.finishDate.Text = ""
        Me.dateToPrint.Text = ""
        Me.StartPackingLabel.Text = ""
        Me.labelQty2.Text = 0

        If String.IsNullOrEmpty(Me.PP.Text) = True Then Exit Sub

        adapter = New SqlDataAdapter("Select * FROM [PPList] WHERE [Order] = '" & Me.PP.Text & "'", Main.koneksi)
        adapter.Fill(ds)

        If ds.Tables(0).Rows.Count = 1 Then
            'Fill the form for packaging label
            Me.SO.Text = ds.Tables(0).Rows(0).Item("SO no").ToString
            Me.SO_item.Text = ds.Tables(0).Rows(0).Item("Item").ToString
            Me.txt_SO_item.Text = ds.Tables(0).Rows(0).Item("Item").ToString
            Me.Material.Text = ds.Tables(0).Rows(0).Item("Material").ToString
            If InStr(Me.header.Text, "CSC") <> 0 Or InStr(Me.header.Text, "ATS") <> 0 Then
                Me.Quantity.Text = q
            Else
                'MsgBox(ds.Tables(0).Rows(0).Item("Item quantity").ToString)
                Me.Quantity.Text = ds.Tables(0).Rows(0).Item("Item quantity").ToString
            End If
            Me.StartPackingLabel.Text = 1

            'Me.customer.Text = ds.Tables(0).Rows(0).Item("Name 1").ToString
            Me.customer.Text = ""
            cust_TextBox1.Text = ds.Tables(0).Rows(0).Item("Name 1").ToString
            Me.Customer_code.Text = ds.Tables(0).Rows(0).Item("customer").ToString
            'If Not String.IsNullOrEmpty(ds.Tables(0).Rows(0).Item("Req dlv dt")) Then
            If Not IsDBNull(ds.Tables(0).Rows(0).Item("Req dlv dt")) Then
                Me.CRD.Text = Convert_To_Date_Format(ds.Tables(0).Rows(0).Item("Req dlv dt").ToString)
                'Dim childAgeAsdouble As Double
                'If Integer.TryParse(ds.Tables(0).Rows(0).Item("Req dlv dt"), childAgeAsdouble) Then
                'Me.CRD.Text = Date.FromOADate(ds.Tables(0).Rows(0).Item("Req dlv dt")).ToString("dd/MM/yyyy")
                'Else
                '    If InStr(ds.Tables(0).Rows(0).Item("Req dlv dt").ToString(), ".") Then
                '        Me.CRD.Text = ds.Tables(0).Rows(0).Item("Req dlv dt").ToString().Replace(".", "/")
                '    Else
                '        Me.CRD.Text = ds.Tables(0).Rows(0).Item("Req dlv dt").ToString()
                '    End If
                'End If
            Else
                If ReqDelvdt = 0 Then
                    MsgBox("Req dlv dt is Empty")
                    ReqDelvdt = 1
                End If
            End If
            If Not IsDBNull(ds.Tables(0).Rows(0).Item("Created on")) Then
                'Me.creationDate.Text = Date.FromOADate(ds.Tables(0).Rows(0).Item("Created on")).ToString("dd/MM/yyyy")
                Me.creationDate.Text = Convert_To_Date_Format(ds.Tables(0).Rows(0).Item("Created on"))
            Else
                MsgBox("Created on is Empty")
            End If
            Me.custPO.Text = ds.Tables(0).Rows(0).Item("Purchase order number").ToString
            Me.custPOitem.Text = ds.Tables(0).Rows(0).Item("POitem").ToString
            Me.sysconfID.Text = ds.Tables(0).Rows(0).Item("Purchase order number").ToString
            Me.sysconfLine.Text = ds.Tables(0).Rows(0).Item("POitem").ToString
            Me.packingDescription.Text = ds.Tables(0).Rows(0).Item("description").ToString
            Me.custMaterial.Text = ds.Tables(0).Rows(0).Item("Customer Material Number").ToString
            If String.IsNullOrEmpty(Me.custMaterial.Text) = True Or Me.custMaterial.Text = "" Then Me.custMaterial.Text = Me.Material.Text
            If Not IsDBNull(ds.Tables(0).Rows(0).Item("Scheduled finish")) Then
                'Me.finishDate.Text = Date.FromOADate(ds.Tables(0).Rows(0).Item("Scheduled finish")).ToString("dd/MM/yyyy")
                Me.finishDate.Text = Convert_To_Date_Format(ds.Tables(0).Rows(0).Item("Scheduled finish"))
            Else
                MsgBox("Scheduled finish is Empty")
            End If
        End If

        If Me.country.Text = "" Then
            Dim dsa As New DataSet
            Call koneksi_db()
            Dim adap2 = New SqlDataAdapter("Select * FROM customerDatabase WHERE [customer code] = '" & Customer_code.Text & "'", Main.koneksi)
            adap2.Fill(dsa)

            If dsa.Tables(0).Rows.Count = 1 Then

                'Me.country.SelectedValue = dsa.Tables(0).Rows(0).Item("country").ToString
                Me.country.Text = dsa.Tables(0).Rows(0).Item("country").ToString
                Me.countryShortName.Text = dsa.Tables(0).Rows(0).Item("country short name").ToString
                Me.Customer_code.Text = dsa.Tables(0).Rows(0).Item("Customer code").ToString
            ElseIf dsa.Tables(0).Rows.Count > 1 Then
                Me.country.SelectedValue = dsa.Tables(0).Rows(0).Item("country").ToString
                Me.countryShortName.Text = dsa.Tables(0).Rows(0).Item("country short name").ToString
                Me.Customer_code.Text = dsa.Tables(0).Rows(0).Item("Customer code").ToString
                If duplicateCustomer = False Then
                    MsgBox("You have duplicate customer database. Please Check customer code:" & Me.Customer_code.Text)
                    duplicateCustomer = True
                End If
            End If
        End If

        If Microsoft.VisualBasic.Left(Me.countryShortName.Text, 2) = "SG" Then
            Me.localexport.Text = "LOCAL"
            Me.dateToPrint.Text = Me.CRD.Text
            If String.IsNullOrEmpty(Me.dateToPrint.Text) And Me.dateToPrint.Text = "" Then Me.dateToPrint.Text = Me.finishDate.Text
        End If

        If Microsoft.VisualBasic.Left(Me.countryShortName.Text, 2) <> "SG" Then
            Me.localexport.Text = "EXPORT"
            Me.dateToPrint.Text = Me.finishDate.Text
        End If

        'MsgBox(ds.Tables(0).Rows(0).Item("Material").ToString)

        '************ define label split
        Dim part = Me.Material.Text
        Dim dss As New DataSet

        Dim labelsql As String = "Select * FROM [labelSelectionTable] WHERE [range start] <= '" & Me.Material.Text & "' and [range end] >= '" & Me.Material.Text & "'"
        Dim adap As New SqlDataAdapter(labelsql, Main.koneksi)
        adap.Fill(dss)

        If dss.Tables(0).Rows.Count = 0 Then Exit Sub

        If qtyperboxinput1 = False And qtyperboxinput2 = False Then
            Me.boxQty.Text = dss.Tables(0).Rows(0).Item("maxBox").ToString
        Else
            Me.boxQty.Text = qtyboxinput
        End If
        Me.palletQty.Text = dss.Tables(0).Rows(0).Item("maxPallet").ToString
        Dim madeIn = ""
        madeIn = dss.Tables(0).Rows(0).Item("Default Made-In").ToString

        If String.IsNullOrEmpty(madeIn) = True Or madeIn = "" Then

            If missingMadein = 0 Then
                MsgBox("Missing made-in information, please check the label selection table")
                missingMadein = 1
            End If

        Else
            Dim ds2 As New DataSet

            Dim adapter2 = New SqlDataAdapter("Select * FROM [madeINtext] WHERE [country] = '" & madeIn.ToUpper & "'", Main.koneksi)
            adapter2.Fill(ds2)

            If ds2.Tables(0).Rows.Count = 0 Then

                MsgBox("Missing/wrong made-in information, please check the label selection table")

            Else
                Me.madeInEnglish.SelectedValue = ds2.Tables(0).Rows(0).Item("EN").ToString
                Me.madeInChinese.Text = ds2.Tables(0).Rows(0).Item("ZH").ToString
                Me.madeInRussian.Text = ds2.Tables(0).Rows(0).Item("RU").ToString
            End If
        End If


        Dim qty As Integer
        Dim pallet As Integer
        Dim box As Integer
        'Dim Vqty As Long
        'Dim Vpallet As Long
        'Dim Vbox As Long

        'Long.TryParse(Me.Quantity.Text, Vqty)
        'Long.TryParse(Me.palletQty.Text, Vpallet)
        'Long.TryParse(Me.boxQty.Text, Vbox)

        qty = CInt(Me.Quantity.Text)
        pallet = CInt(Me.palletQty.Text)
        box = CInt(Me.boxQty.Text)

        If qty < pallet Then

            Dim dec = (qty / box) - (qty / box)
            If dec = 0 Then Me.labelQty.Text = Convert.ToInt32(qty / box)
            If dec > 0 Then Me.labelQty.Text = Convert.ToInt32(qty / box) + 1
            Me.boxQty.Font = New Font(boxQty.Font, FontStyle.Bold)
            Me.checkBox.CheckState = 1
        End If

        If qty = pallet Then
            Dim dec = (qty / pallet) - (qty / pallet)
            If dec = 0 Then Me.labelQty.Text = Convert.ToInt32(qty / pallet)
            If dec > 0 Then Me.labelQty.Text = Convert.ToInt32(qty / pallet) + 1
            Me.palletQty.Font = New Font(palletQty.Font, FontStyle.Bold)
            Me.checkPallet.CheckState = 1
        End If

        If qty > pallet Then
            Dim dec = (qty / pallet) - (qty / pallet)
            If dec = 0 Then Me.labelQty.Text = Convert.ToInt32(qty / pallet)
            If dec > 0 Then Me.labelQty.Text = Convert.ToInt32(qty / pallet) + 1
            Me.palletQty.Font = New Font(palletQty.Font, FontStyle.Bold)
            Me.checkPallet.CheckState = 1
        End If

        'Copy the data
        Me.labelQty2.Text = Me.labelQty.Text
        Me.Label468.Text = Me.labelQty.Text
        Me.StartPackingLabel2.Text = Me.StartPackingLabel.Text

        ' Product label
        'qty label
        Me.quantityLabel.Text = Me.Quantity.Text
        Me.StartLabel.Text = 1

        'Automatic printing
        If Me.autoPrint2.CheckState = 1 Then

            Dim chosenLabel As String

            chosenLabel = selectLabel()

            If chosenLabel = "" Then Exit Sub

            printPackingLabel(chosenLabel)

        End If
    End Sub

    Private Sub PreviewLabelPacking_Click(sender As Object, e As EventArgs) Handles previewLabelPacking.Click
        reload_printer()
        'set to NiceLabel Variable
        'label1_printer.Variables("Name").SetValue("111111111")
        label1_setValue()
        'declaration of Preview
        Dim LabelPreviewSettings As ILabelPreviewSettings = New LabelPreviewSettings()

        'setting preview format
        LabelPreviewSettings.ImageFormat = "PNG"
        LabelPreviewSettings.Width = Form_preview.pictureBoxPreview.Width                   ' Width Of image To generate
        LabelPreviewSettings.Height = Form_preview.pictureBoxPreview.Height                 ' Height Of image To generate

        ' Generate Preview File
        Dim imageObj As Object = label1_printer.GetLabelPreview(LabelPreviewSettings)

        'Display image in UI
        If TypeOf imageObj Is Byte() Then
            Form_preview.pictureBoxPreview.Image = ByteToImage(CType(imageObj, Byte()))
        ElseIf TypeOf imageObj Is String Then
            Form_preview.pictureBoxPreview.ImageLocation = CStr(imageObj)
        End If

        Form_preview.Show()

    End Sub

    'Byte to Image convertion
    Private Function ByteToImage(ByVal bytes As Byte()) As Bitmap
        Dim memoryStream As MemoryStream = New MemoryStream()
        memoryStream.Write(bytes, 0, Convert.ToInt32(bytes.Length))
        Dim bm As Bitmap = New Bitmap(memoryStream, False)
        memoryStream.Dispose()
        Return bm
    End Function

    Private Sub SO_TextChanged(sender As Object, e As EventArgs) Handles SO.TextChanged
        If String.IsNullOrEmpty(SO.Text) = False Then
            Me.localexport.Text = ""
            If Microsoft.VisualBasic.Left(Me.SO.Text, 2) = 12 Then Me.localexport.Text = "local"
            If Microsoft.VisualBasic.Left(Me.SO.Text, 2) = 14 Then Me.localexport.Text = "export"
        End If
    End Sub

    Function selectLabel() 'Select the label type ' sudah sesuai dengan VBA

        Dim ds As New DataSet
        Dim ds2 As New DataSet
        Dim ds3 As New DataSet
        Dim Category As String

        selectLabel = ""

        Category = Me.header.Text

        Dim adapter2 = New SqlDataAdapter("Select * from [NSXMasterdata] Where [material] = '" & Category & "'", Main.koneksi)
        adapter2.Fill(ds2)

        If ds2.Tables(0).Rows.Count > 0 Then

            selectLabel = "generic"
            selectedLabel.Text = selectLabel

            Me.range.Text = ds2.Tables(0).Rows(0).Item("range").ToString
            Me.description.Text = ds2.Tables(0).Rows(0).Item("description").ToString
            Me.descriptionChinese.Text = ds2.Tables(0).Rows(0).Item("descriptionZH").ToString
            Me.descriptionFrench.Text = ds2.Tables(0).Rows(0).Item("descriptionFR").ToString
            Me.descriptionRussian.Text = ds2.Tables(0).Rows(0).Item("descriptionRU").ToString
            Me.descriptionSpanish.Text = ds2.Tables(0).Rows(0).Item("descriptionES").ToString
            Me.technicalDescription.Text = ds2.Tables(0).Rows(0).Item("Technical description").ToString

            Me.EAN13.Text = ds2.Tables(0).Rows(0).Item("EAN13").ToString
            'MsgBox(EAN13.Text)

            Me.Picture2.Text = ds2.Tables(0).Rows(0).Item("productImage").ToString

            Me.logo1value.Text = ds2.Tables(0).Rows(0).Item("logo1").ToString
            Me.logo2value.Text = ds2.Tables(0).Rows(0).Item("logo2").ToString
            Me.logo3value.Text = ds2.Tables(0).Rows(0).Item("logo3").ToString

            'Try
            Me.logo4value.Text = ds2.Tables(0).Rows(0).Item("logo4").ToString
                Me.logo5value.Text = ds2.Tables(0).Rows(0).Item("logo5").ToString
                Me.logo6value.Text = ds2.Tables(0).Rows(0).Item("logo6").ToString
            'Catch ex As Exception
            'Log_data("Logo :" & DateTime.Today.ToString & ex.ToString)
            ' End Try



        Else

            adapter2 = New SqlDataAdapter("Select * FROM [labelSelectionTable] WHERE [range start] <= '" & Category & "' AND [range end] >= '" & Category & "'", Main.koneksi)
            adapter2.Fill(ds2)

            If ds2.Tables(0).Rows.Count = 0 Then
                'santo add
                If Not header.Text.Contains("BW") Then MsgBox("no label has been found for this range of product")
            Else

                selectLabel = ds2.Tables(0).Rows(0).Item("Label").ToString
                selectedLabel.Text = selectLabel

                If ds2.Tables(0).Rows(0).Item("special").ToString <> "" Or String.IsNullOrEmpty(ds2.Tables(0).Rows(0).Item("special").ToString) = False Then
                    Me.warning.Text = ds2.Tables(0).Rows(0).Item("special").ToString
                    Me.range.Text = ds2.Tables(0).Rows(0).Item("special").ToString
                    Me.warning.Visible = True
                End If

                Dim Order As String
                Order = Me.PPnumberEntry.Text

                Dim adapter = New SqlDataAdapter("Select * FROM [PPList] WHERE [Order] = '" & Order & "'", Main.koneksi)
                adapter.Fill(ds)

                If ds.Tables(0).Rows.Count > 0 Then
                    Me.description.Text = ds.Tables(0).Rows(0).Item("description").ToString
                End If

                If selectLabel = "generic" Then
                    adapter = New SqlDataAdapter("Select [openOrders].[Material], [NSXMasterdata].* FROM [openOrders] INNER JOIN [NSXMasterdata] ON [openOrders].[Material]=[NSXMasterdata].[Material] WHERE [openOrders].[Order] = '" & Me.PPnumberEntry.Text & "'", Main.koneksi)
                    adapter.Fill(ds3)

                    If ds3.Tables(0).Rows.Count >= 1 Then
                        'MsgBox("masuk")
                        Me.description.Text = ds3.Tables(0).Rows(0).Item("description").ToString
                        Me.descriptionChinese.Text = ds3.Tables(0).Rows(0).Item("descriptionZH").ToString
                        Me.descriptionFrench.Text = ds3.Tables(0).Rows(0).Item("descriptionFR").ToString
                        Me.descriptionRussian.Text = ds3.Tables(0).Rows(0).Item("descriptionRU").ToString
                        Me.descriptionSpanish.Text = ds3.Tables(0).Rows(0).Item("descriptionES").ToString
                        Me.technicalDescription.Text = ds3.Tables(0).Rows(0).Item("Technical description").ToString

                        Me.EAN13.Text = ds3.Tables(0).Rows(0).Item("EAN13").ToString
                        'MsgBox(EAN13.Text)
                        Me.Picture2.Text = ds3.Tables(0).Rows(0).Item("productImage").ToString

                        Me.logo1value.Text = ds3.Tables(0).Rows(0).Item("logo1").ToString
                        Me.logo2value.Text = ds3.Tables(0).Rows(0).Item("logo2").ToString
                        Me.logo3value.Text = ds3.Tables(0).Rows(0).Item("logo3").ToString

                        'Try
                        Me.logo4value.Text = ds2.Tables(0).Rows(0).Item("logo4").ToString
                            Me.logo5value.Text = ds2.Tables(0).Rows(0).Item("logo5").ToString
                            Me.logo6value.Text = ds2.Tables(0).Rows(0).Item("logo6").ToString
                        'Catch ex As Exception
                        'Log_data("Logo :" & DateTime.Today.ToString & ex.ToString)
                        'End Try


                    Else
                        Me.Picture2.Text = ds2.Tables(0).Rows(0).Item("productImage").ToString
                    End If
                End If
            End If
        End If
    End Function

    Private Sub printPackingLabel(chosenLabel As String) 'Print packaging  & product label 'sudah sesuai dengan VBA
        Try

            ' Check whether the Packaging label has already printed or not
            ' 27092019
            'If checkDuplicatePPpacking() = True Then Exit Sub

            Dim start As Long
            Dim startprint As Integer
            Dim endprint As Integer
            Dim pallet As Long
            Dim box As Long
            Dim totalqty As Long
            Dim labelQty2 As Long
            Dim labelNumber As Long
            Dim sql As String
            Dim ds As New DataSet

            Dim checkQty = 0
            totalqty = CInt(Me.Quantity.Text)
            pallet = CInt(Me.palletQty.Text)
            box = CInt(Me.boxQty.Text)
            labelQty2 = 0
            labelNumber = 1

            Dim selectedPrinterPacking = ""
            Dim selectedPrinterProduct = ""
            Dim selectedPackingDpi = ""
            Dim packingLabel = ""


            'Select the right printerProduct
            sql = "Select * " &
                    "FROM [printerTable] " &
                    "WHERE [workstation] = '" & Me.workstation.SelectedValue & "'" &
                    "AND [report type] ='" & "product label" & "';"
            Dim adapter = New SqlDataAdapter(sql, Main.koneksi)
            adapter.Fill(ds)

            'If ds.Tables(0).Rows.Count = 0 Then
            '    MsgBox("printer not found")
            '    Exit Sub
            'End If

            selectedPrinterProduct = ds.Tables(0).Rows(0).Item("Printer").ToString

            'Select the right printerPacking
            sql = "Select * " &
                    "FROM [printerTable] " &
                    "WHERE [workstation] = '" & Me.workstation.SelectedValue & "'" &
                    "AND [report type] ='" & "packing label" & "';"

            'If ds.Tables(0).Rows.Count = 0 Then
            '    MsgBox("printer not found")
            '    Exit Sub
            'End If

            '#################3 SIMULATE PRINTER BY TUGUSS ######################
            selectedPrinterPacking = ds.Tables(0).Rows(0).Item("Printer").ToString
            '####################################################################

            selectedPackingDpi = ds.Tables(0).Rows(0).Item("dpi").ToString

            If CInt(selectedPackingDpi) = 600 Then
                packingLabel = "adaptation packaging label"
            Else
                packingLabel = "adaptation packaging label"
            End If

            '------- Print Out Packaging labels -----------

            '--------If "max qty per pallet" checkbox Checked
            If Me.checkPallet.CheckState = 1 Then

                'number of Packaging label to be printed
                'dec = (totalqty / pallet) - Int(totalqty / pallet)
                'If dec = 0 Then Me.labelQty = (totalqty / pallet)
                'If dec > 0 Then Me.labelQty = Int(totalqty / pallet) + 1

                startprint = Me.StartPackingLabel.Text
                endprint = Me.labelQty.Text
                start = 1


                Do While totalqty > 0

                    'If totalquantity >= Quantity max per pallet
                    If totalqty >= pallet Then labelQty2 = pallet
                    'If totalquantity < Quantity max per pallet
                    If totalqty < pallet Then labelQty2 = totalqty

                    'DoCmd.SelectObject acReport, packingLabel, True
                    'DoCmd.OpenReport packingLabel, acViewPreview, , , acHidden
                    'Reports(packingLabel)![packageqty].Caption = labelQty2
                    'Reports(packingLabel)![Label#].Caption = labelNumber
                    'Reports(packingLabel)![Text27].Caption = Me.Label468.Caption

                    'For SG customers
                    If Me.countryShortName.Text = "SG" Then
                        'Reports(packingLabel)![countryShortName].BackColor = 0
                        'Reports(packingLabel)![countryShortName].ForeColor = 16777215
                    Else
                        'Reports(packingLabel)![countryShortName].BackColor = 16777215
                        'Reports(packingLabel)![countryShortName].ForeColor = 0
                    End If

                    '
                    If start >= startprint And start <= endprint Then 'Useful in case of reprinting selected labels
                        'Set Application.Printer = Application.Printers(selectedPrinterPacking)
                        'Printer.PrintQuality = 300
                        'DoCmd.SelectObject acReport, packingLabel, True
                        'DoCmd.PrintOut , , , acHigh, 1
                    End If

                    'DoCmd.RunCommand acCmdWindowHide
                    'DoCmd.Close acReport, packingLabel


                    start = start + 1
                    labelNumber = labelNumber + 1
                    totalqty = totalqty - labelQty2

                Loop

            End If


            '-------If "max qty per box" checkbox is Checked
            If Me.checkBox.CheckState = 1 Then

                'number of Packaging label to be printed
                'dec = (totalqty / box) - Int(totalqty / box)
                'If dec = 0 Then Me.labelQty = (totalqty / box)
                'If dec > 0 Then Me.labelQty = Int(totalqty / box) + 1

                startprint = Me.StartPackingLabel.Text
                endprint = Me.labelQty.Text
                start = 1


                Do While totalqty > 0

                    If totalqty >= box Then labelQty2 = box
                    If totalqty < box Then labelQty2 = totalqty

                    'DoCmd.SelectObject acReport, packingLabel, True
                    'DoCmd.OpenReport packingLabel, acViewPreview, , , acHidden
                    'Reports(packingLabel)![packageqty].Caption = labelQty2
                    'Reports(packingLabel)![Label#].Caption = labelNumber
                    'Reports(packingLabel)![Text27].Caption = Me.Label468.Caption

                    'For SG customers
                    If Me.countryShortName.Text = "SG" Then
                        'Reports(packingLabel)![countryShortName].BackColor = 0
                        'Reports(packingLabel)![countryShortName].ForeColor = 16777215
                    Else
                        'Reports(packingLabel)![countryShortName].BackColor = 16777215
                        'Reports(packingLabel)![countryShortName].ForeColor = 0
                    End If


                    If start >= startprint And start <= endprint Then 'Useful in case of reprinting selected labels
                        'Set Application.Printer = Application.Printers(selectedPrinterPacking)
                        'Printer.PrintQuality = 300
                        'DoCmd.SelectObject acReport, packingLabel, True
                        'DoCmd.PrintOut , , , acHigh, 1
                    End If

                    'DoCmd.RunCommand acCmdWindowHide
                    'DoCmd.Close acReport, packingLabel

                    labelNumber = labelNumber + 1
                    totalqty = totalqty - labelQty2
                    start = start + 1

                Loop

            End If

            'Me.PPnumberEntry.SetFocus
            'Me.ComponentNo.Select()
            CompToQuality.Select()

            '------- Print Out Product labels ----------_
            ' If CheckBox Product label printing is checked
            If Me.CheckProductLabelPrinting2.CheckState = 1 Then
                start = Me.StartLabel.Text
                labelQty2 = Me.quantityLabel.Text
                printProductLabel(chosenLabel)
            End If
        Catch ex As Exception

        End Try
    End Sub

    Function read_serialNumber()
        Dim sql2 As String
        Dim adapter2 As SqlDataAdapter
        Dim ds2 As New DataSet

        sql2 = " Select Case [sequence] FROM SGRAC_MES.dbo.serialNumber WHERE [day] like '%" & DateTime.Now.ToString("yyyy-MM-dd") & "%'"
        adapter2 = New SqlDataAdapter(sql2, Main.koneksi)
        adapter2.Fill(ds2)

        Dim seq As Integer = Convert.ToDecimal(ds2.Tables(0).Rows(0).Item("sequence").ToString)
        read_serialNumber = seq
    End Function


    Function checkSerialNumber() 'Check the serial number or add one 'sudah sesuai dengan VBA

        Dim seq_old As Integer

        Dim sql As String
        Dim adapter As SqlDataAdapter
        Dim ds As New DataSet
        Dim cmd As SqlCommand

        'sql = "Select * FROM [serialNumber] WHERE [day] like %" & DateTime.Now.ToString("yyyy-MM-dd") & "%"
        sql = "SELECT *  FROM [dbo].[serialNumber] WHERE [day] like '%" & DateTime.Now.ToString("yyyy-MM-dd") & "%'"
        adapter = New SqlDataAdapter(sql, Main.koneksi)
        adapter.Fill(ds)



        'MsgBox(ds.Tables(0).Rows.Count)

        If ds.Tables(0).Rows.Count = 0 Then
            cmd = New SqlCommand("insert into serialNumber([day],[sequence]) values(@day,@seq)", Main.koneksi)
            cmd.Parameters.AddWithValue("@day", DateTime.Now.ToString("yyyy/MM/dd"))
            cmd.Parameters.AddWithValue("@seq", 1)
            cmd.ExecuteNonQuery()
            checkSerialNumber = 1
        Else
            seq_old = Convert.ToDecimal(ds.Tables(0).Rows(0).Item("sequence").ToString)
            checkSerialNumber = ds.Tables(0).Rows(0).Item("sequence")
            cmd = New SqlCommand("update [serialNumber] set [sequence]=@seq where [day]=@day", Main.koneksi)
            cmd.Parameters.AddWithValue("@day", DateTime.Now.ToString("yyyy/MM/dd"))
            cmd.Parameters.AddWithValue("@seq", 1 + seq_old)
            cmd.ExecuteNonQuery()


            Dim sql2 As String
            Dim adapter2 As SqlDataAdapter
            Dim ds2 As New DataSet

            sql2 = " Select Case [sequence] FROM SGRAC_MES.dbo.serialNumber WHERE [day] like '%" & DateTime.Now.ToString("yyyy-MM-dd") & "%'"
            adapter2 = New SqlDataAdapter(sql, Main.koneksi)
            adapter2.Fill(ds2)

            Dim seq As Integer = Convert.ToDecimal(ds2.Tables(0).Rows(0).Item("sequence").ToString)
            checkSerialNumber = seq
        End If

    End Function
    'santo
    Private Sub Listprinter_SelectedIndexChanged(sender As Object, e As EventArgs) Handles listprinter.SelectedIndexChanged
        'Selection of printer
        If printers.Count > 0 Then
            selected_Printer = printers.Item(listprinter.SelectedIndex)
            label_printer.PrintSettings.PrinterName = selected_Printer.Name
            listprinter.SelectedItem = selected_Printer.Name
        End If
    End Sub

    'santo
    Private Sub printProductLabel(chosenLabel As String) ' sudah sesuai dengan VBA

        Dim sql As String
        Dim ds As New DataSet
        Dim Category As String
        Dim ds3 As New DataSet
        Dim sql3 As String
        Dim iii As String

        Category = Me.header.Text


        'sql = "SELECT [printer]  FROM [dbo].[printerTable] where [workstation]= '" &
        '    workstation.SelectedValue.ToString & "' AND [report type]= 'packing label'"
        'Dim adapter = New SqlDataAdapter(sql, Main.koneksi)
        'adapter.Fill(ds)

        'If ds.Tables(0).Rows.Count = 0 Then
        '    MsgBox("Printer not set !")
        '    Exit Sub
        'End If

        'Dim selectedPrinter = ds.Tables(0).Rows(0).Item("printer").ToString

        '************** Display the right picture *******************************
        If chosenLabel = "generic" Then

            'If Me.EAN13.Text = "" Or String.IsNullOrEmpty(Me.EAN13.Text) = True Or Me.EAN13.Text = "111" Then
            '    Me.EAN13.Text = "111"
            'End If

            If Me.countryShortName.Text.IndexOf("KR") <> -1 Then
                'NEW FOR KR CERTIFICATION NUMBER Stephen
                sql3 = "Select [certificationnumber] from [dbo].[KC_certification_table] " &
                "Where [Reference] ='" & Category & "'"
                Dim adapter3 = New SqlDataAdapter(sql3, Main.koneksi)
                adapter3.Fill(ds3)

                If ds3.Tables(0).Rows.Count > 0 Then
                    iii = ds3.Tables(0).Rows(0).Item("certificationnumber").ToString
                    label_printer.Variables("KRtextcertif1").SetValue(iii)
                    label_printer.Variables("Korea_cerft").SetValue(1)
                Else
                    label_printer.Variables("KRtextcertif1").SetValue("   ")
                    label_printer.Variables("Korea_cerft").SetValue(0)
                End If

            End If
        End If


        'If chosenLabel = "loose" Then
        'Reports(Trim(chosenLabel))![certifKoreaIdentification].Visible = False
        'Reports(Trim(chosenLabel))![KRtextcertif1].Visible = False
        'Reports(Trim(chosenLabel))![certifUAidentification].Visible = False
        'Reports(Trim(chosenLabel))![certifTRidentification].Visible = False
        'Reports(Trim(chosenLabel))![certifEACidentification].Visible = False
        'End If

        'If Me.checkCombine.CheckState = 1 Then 'One product label which combine all the data

        '    Dim firstSerial = checkSerialNumber(1)

        '    Dim yearNum As Integer = 19
        '    Dim thisweek As Integer = 12
        '    If Len(thisweek) = 1 Then thisweek = "0" & thisweek

        '    Dim dayNum = 12

        '    Dim j = firstSerial
        '    Dim serial = 0

        '    If Len(j) = 1 Then serial = "000" & j
        '    If Len(j) = 2 Then serial = "00" & j
        '    If Len(j) = 3 Then serial = "0" & j


        If chosenLabel = "loose2" And Me.labelQty.Text > 1 Then

            '        Dim PrintCount As Integer = 0

            '        For i = 1 To Me.labelQty.Text - 1

            '            'Reports(Trim(chosenLabel))![Label13].Caption = Me.boxQty
            '            'print
            '            'DoCmd.SelectObject acReport, Trim(chosenLabel), True
            '            'DoCmd.PrintOut , , , , 1

            '            PrintCount = PrintCount + CInt(Me.boxQty.Text)

            '        Next

            '        'Reports(Trim(chosenLabel))![Label13].Caption = Me.Quantity - PrintCount
            '        'print
            '        'DoCmd.SelectObject acReport, Trim(chosenLabel), True
            '        'DoCmd.PrintOut , , , , 1


        Else
            '        Dim traceability = "SG" & yearNum & thisweek & dayNum & serial
            '        'Reports(Trim(chosenLabel))![Label13].Caption = Me.Quantity
            'Try
            '    If workstation.Text = "Kitting , Pendent & LC Contactor" Then
            '        label_printer.Variables("Label13").SetValue(Me.Quantity.Text)
            '    End If
            'Catch ex As Exception
            'End Try

            '        If chosenLabel = "loose" Then
            '            'Reports(Trim(chosenLabel))![serialNum].Caption = traceability
            '            'Reports(Trim(chosenLabel))![ItemNumber].Caption = Me.Quantity
            '        End If

            '        '        DoCmd.SelectObject acReport, Trim(chosenLabel), True
            '        'DoCmd.PrintOut , , , , 1

        End If

        'End If


        'If Me.checkCombine.CheckState = 0 Then

        '    Dim firstSerial = checkSerialNumber(Quantity)

        '    Dim yearNum As Integer = 19
        '    Dim thisweek As Integer = 12
        '    If Len(thisweek) = 1 Then thisweek = "0" & thisweek

        '    Dim dayNum = 12

        '    Dim i = startQty
        '    Dim j = firstSerial
        '    Dim maxValue = 1 + Quantity
        '    Dim serial = 0

        '    Do While i < maxValue

        '        If Len(j) = 1 Then serial = "000" & j
        '        If Len(j) = 2 Then serial = "00" & j
        '        If Len(j) = 3 Then serial = "0" & j

        '        Dim traceability = "SG" & yearNum & thisweek & dayNum & serial

        '        If chosenLabel = "generic" Or chosenLabel = "loose" Then
        '            'Reports(Trim(chosenLabel))![ItemNumber].Caption = i
        '            'Reports(Trim(chosenLabel))![serialNum].Caption = traceability
        '            'Reports(Trim(chosenLabel))![Text27].Caption = Me.Quantity
        '        End If

        '        '    DoCmd.SelectObject acReport, Trim(chosenLabel), True
        '        'DoCmd.PrintOut , , , , 1

        '        i = i + 1
        '        j = j + 1

        '    Loop

        'End If

        'DoCmd.RunCommand acCmdWindowHide
        'DoCmd.Close acReport, Trim(chosenLabel)

    End Sub



    'santo
    'Set to Variable of NiceLabel
    Private Sub Label_SetValue()
        Dim ds3 As New DataSet
        Dim sql3 As String
        Dim iii As String
        Dim Category As String
        Category = Me.header.Text

        Try
            If Me.countryShortName.Text.IndexOf("KR") <> -1 Then
                'NEW FOR KR CERTIFICATION NUMBER Stephen
                sql3 = "Select [certificationnumber] from [dbo].[KC_certification_table] " &
                    "Where [Reference] ='" & Category & "'"
                Dim adapter3 = New SqlDataAdapter(sql3, Main.koneksi)
                adapter3.Fill(ds3)

                If ds3.Tables(0).Rows.Count > 0 Then
                    iii = ds3.Tables(0).Rows(0).Item("certificationnumber").ToString
                    label_printer.Variables("KRtextcertif1").SetValue(iii)
                    label_printer.Variables("Korea_cerft").SetValue(1)
                Else
                    label_printer.Variables("KRtextcertif1").SetValue("   ")
                    label_printer.Variables("Korea_cerft").SetValue(0)
                End If

            End If
        Catch ex As Exception
        End Try

        Try
            'label_printer.Variables("Product ID").SetValue("111111111")
            label_printer.Variables("Product ID").SetValue(Microsoft.VisualBasic.Left(header.Text, 19))
            If description.Text <> "" Then
                label_printer.Variables("Product Desc Eng").SetValue(Microsoft.VisualBasic.Left(description.Text, 33))
            Else
                label_printer.Variables("Product Desc Eng").SetValue("   ")
            End If
        Catch ex As Exception

        End Try

        Try
            label_printer.Variables("Comertial Name").SetValue(Microsoft.VisualBasic.Left(range.Text, 18))
            label_printer.Variables("EAN13").SetValue(EAN13.Text)
            label_printer.Variables("PP Number").SetValue(PPnumberEntry.Text)
            label_printer.Variables("technicianShortName").SetValue(technicianShortName.Text)
        Catch ex As Exception

        End Try


        Try
            label_printer.Variables("country").SetValue(madeInEnglish.Text)
            label_printer.Variables("country").SetValue(madeInEnglish.SelectedValue.ToString)
            'label_printer.Variables("country").SetValue(madeInEnglish.SelectedValue.ToString)
        Catch ex As Exception
            MsgBox("Made In not selected !")
            'If range.Text.IndexOf("CVS") <> -1 Or range.Text.IndexOf("NSX") Or range.Text.IndexOf("NT") Then
            '    madeInEnglish.Text = "Made in China"
            'Else
            '    madeInEnglish.Text = "Made in India"
            'End If
            'label_printer.Variables("country").SetValue(Microsoft.VisualBasic.Left(madeInEnglish.Text, 22))
        End Try

        'checkbox Made in fuction
        If checkMadeInChina.Checked = False Then label_printer.Variables("country").SetValue("    ")

        'AssembledBySing
        If checkAssembledSingapore.Checked = True Then label_printer.Variables("AssembledBySing").SetValue(1)



        label_printer.Variables("Date Print").SetValue("SG-" & dateCode.Text)
        label_printer.Variables("Logistic Ref").SetValue(header.Text)
        label_printer.Variables("Made in Ch").SetValue(madeInChinese.Text)
        label_printer.Variables("Made in rus").SetValue(madeInRussian.Text)

        label_printer.Variables("Made in singapore").SetValue(madeInSingapore.Text)
        label_printer.Variables("LabelQuantityitem").SetValue(LabelQuantityitem.Text)

        label_printer.Variables("Technical Dis").SetValue(Microsoft.VisualBasic.Left(technicalDescription.Text, 18) & " ")
        label_printer.Variables("Pic_contactor").SetValue(Picture2.Text)
        label_printer.Variables("date code").SetValue("SG-" & dateCode.Text)

        label_printer.Variables("Product Desc Fr").SetValue(descriptionFrench.Text)
        label_printer.Variables("Product Desc Span").SetValue(descriptionSpanish.Text)
        label_printer.Variables("Product Desc Ch").SetValue(descriptionChinese.Text)
        label_printer.Variables("Product Desc Rus").SetValue(descriptionRussian.Text)
        label_printer.Variables("cust Material").SetValue(Microsoft.VisualBasic.Left(custMaterial.Text, 19))

        Dim SG_sign As String
        Dim test As String = Microsoft.VisualBasic.Right(dateCode.Text, 2)
        If dateCode.Text(7) <> "-" Then

            'If test.IndexOf("-") <> -1 Then
            'SG_sign = "SG" & dateCode.Text(2) & dateCode.Text(3) & dateCode.Text(6) & dateCode.Text(7) & test
            'Else
            SG_sign = "SG" & dateCode.Text(2) & dateCode.Text(3) & dateCode.Text(6) & dateCode.Text(7) & Microsoft.VisualBasic.Right(dateCode.Text, 1) & "0"
            ' End If
        Else

            ' If test.IndexOf("-") <> -1 Then
            SG_sign = "SG" & dateCode.Text(2) & dateCode.Text(3) & "0" & dateCode.Text(6) & Microsoft.VisualBasic.Right(dateCode.Text, 1) & "0"
            ' Else
            'SG_sign = "SG" & dateCode.Text(2) & dateCode.Text(3) & "0" & dateCode.Text(6) & test
        End If

        'End If

        label_printer.Variables("SG Sign").SetValue(SG_sign)

        label_printer.Variables("Logo1").SetValue(logo1value.Text)
        label_printer.Variables("Logo2").SetValue(logo2value.Text)
        label_printer.Variables("Logo3").SetValue(logo3value.Text)

        Try
            label_printer.Variables("Logo4").SetValue(logo4value.Text)
            label_printer.Variables("Logo5").SetValue(logo5value.Text)
            label_printer.Variables("Logo6").SetValue(logo6value.Text)
        Catch ex As Exception
            'Log_data("Logo :" & DateTime.Today.ToString & ex.ToString)
        End Try


        'compare Neutral L
        'Dim strList1 As List(Of String) = New List(Of String)(New String() {"47595", "47596", "47597", "LV847595", "LV847596", "LV847597"})
        'Dim Neutral As String
        'Dim pole As String
        'Neutral = Material.Text
        Try
            'If Neutral = "47595" Or Neutral = "47596" Or Neutral = "47597" Or Neutral = "LV847595" Or Neutral = "LV847596" Or Neutral = "LV847597" Then
            '    label_printer.Variables("Neutral").SetValue("R")
            '    TextBox_Neutral.Text = "R"
            'ElseIf instr(description.Text, "4P") Then
            '    label_printer.Variables("Neutral").SetValue("L")
            '    TextBox_Neutral.Text = "L"
            'Else
            '    label_printer.Variables("Neutral").SetValue("")
            '    TextBox_Neutral.Text = ""
            'End If
            If InStr(TextBox_Neutral.Text, "R") = 0 Then
                If InStr(description.Text, "4P") Then
                    'label_printer.Variables("Neutral").SetValue("L")
                    TextBox_Neutral.Text = "L"
                Else
                    'label_printer.Variables("Neutral").SetValue("")
                    'TextBox_Neutral.Text = ""
                End If
            End If

            label_printer.Variables("Neutral").SetValue(TextBox_Neutral.Text)

            'tambahan kusus csc
            If header.Text.IndexOf("CSC") > 0 Then label_printer.Variables("LabelQuantityitem").SetValue(Quantity.Text)

        Catch ex As Exception

        End Try

        'tambahan santo
        Try
            label_printer.Variables("Label13").SetValue("1")

            'If workstation.Text = "Kitting , Pendent & LC Contactor" And checkCombine.Checked = True Then
            If checkCombine.Checked = True Then
                label_printer.Variables("Label13").SetValue(Me.Quantity.Text)
            End If
        Catch ex As Exception
            'MsgBox("Pls Unchecked check Combine!")
        End Try

        Try
            If select_quantity.Checked = True Then
                label_printer.Variables("Label13").SetValue(Me.Quantity.Text)
            End If
        Catch ex As Exception

        End Try

        Try
            label_printer.Variables("SQ00_item").SetValue(Convert.ToDecimal(txt_SO_item.Text).ToString("0000"))
        Catch ex As Exception

        End Try


        Try
            label_printer.Variables("SQ00_Description").SetValue(Microsoft.VisualBasic.Left(Me.testDescription.Text, 24))
        Catch ex As Exception

        End Try
        Try
            Dim date_2 As String

            Dim VBAWeekNum As Integer = DatePart(DateInterval.WeekOfYear, Date.Today, FirstDayOfWeek.Monday, FirstWeekOfYear.FirstFourDays)

            If Len(VBAWeekNum) = 1 Then VBAWeekNum = "0" & VBAWeekNum

            date_2 = Date.Now.ToString("yy") & VBAWeekNum & DateAndTime.Weekday(DateTime.Now, vbMonday)

            'label_printer.Variables("DataMatrix").SetValue("SG" & date_2 & Convert.ToDecimal(CounterItems.Text).ToString("0000"))

            Dim _ID As String = Convert.ToDecimal(seq_SerialNumber.ToString).ToString("0000")

            'label_printer.Variables("DataMatrix").SetValue("SG" & date_2 & _ID)
            label_printer.Variables("Matrix").SetValue("SG" & date_2 & _ID)
        Catch ex As Exception

        End Try

    End Sub

    Private Sub checkNeutral(a As String)
        Dim Neutral As String
        'Dim pole As String
        Neutral = a
        'Try
        '    If InStr(description.Text, "4P") Then
        '        label_printer.Variables("Neutral").SetValue("L")
        '        TextBox_Neutral.Text = "L"
        '    Else
        If Neutral = "47595" Or Neutral = "47596" Or Neutral = "47597" Or Neutral = "LV847595" Or Neutral = "LV847596" Or Neutral = "LV847597" Then
            'label_printer.Variables("Neutral").SetValue("R")
            TextBox_Neutral.Text = "R"
        End If
        '        Else
        '            label_printer.Variables("Neutral").SetValue("")
        '            TextBox_Neutral.Text = ""
        '        End If
        '    End If

        'Catch ex As Exception

        'End Try
    End Sub



    'Product label Printing ....................
    Private Sub Product_Label_print()

        'select Label
        'selectLabel()

        'MsgBox(label_printer.PrintSettings.PrinterName.ToString)

        'select generic or Loose2
        printProductLabel(selectedLabel.Text)

        '====================================================================================================
        'If selectedLabel.Text.ToLower = "generic" Or selectedLabel.Text = "" Then
        If selectedLabel.Text.ToLower = "generic" Then

            seq_SerialNumber = checkSerialNumber()
            'set value to label
            Label_SetValue()


            'select generic or Loose2
            'printProductLabel(selectedLabel.Text)

            'printProductLabel(chosenLabel As String, startQty As Integer, Quantity As Integer)

            Dim a As Integer
            If Convert.ToDecimal(quantityLabel.Text) >= Convert.ToDecimal(StartLabel.Text) Then
                For a = Convert.ToDecimal(StartLabel.Text) To Convert.ToDecimal(quantityLabel.Text)
                    label_printer.Variables("CounterItems").SetValue(a.ToString)
                    Try
                        label_printer.Variables("DataMatrix").SetValue("SG" & dateCode3() & Convert.ToDecimal(a.ToString).ToString("0000"))
                    Catch ex As Exception

                    End Try
                    'printing with quantity 
                    Dim qty As Integer = 1
                    Try
                        label_printer.Print(qty)
                    Catch ex As Exception
                        MsgBox("Printing Label Cancel " & ex.Message)
                    End Try
                Next
                cmd = New SqlCommand("insert into printingRecord([pp],[date],[time],[user],[from],[to],[Data],[Seq]) values(@pp,@date,@time,@user,@from,@to,@data,@Seq)", Main.koneksi)
                cmd.Parameters.AddWithValue("@pp", Me.PPnumberEntry.Text)
                cmd.Parameters.AddWithValue("@date", DateTime.Now.ToString("yyyy-MM-dd"))
                cmd.Parameters.AddWithValue("@time", DateTime.Now.ToString("HH:mm:ss"))
                cmd.Parameters.AddWithValue("@user", Me.technicianShortName.Text)

                If var_PrintLabel_klik = True Then
                    cmd.Parameters.AddWithValue("@from", CInt(Me.StartLabel.Text))
                    var_PrintLabel_klik = False
                Else
                    cmd.Parameters.AddWithValue("@from", CInt(Me.CounterItems.Text) + 1)
                End If


                cmd.Parameters.AddWithValue("@to", Me.LabelQuantityitem.Text)
                cmd.Parameters.AddWithValue("@data", LoginForm.strHostName & " - " & Application.ProductVersion)

                Dim date_2 As String

                Dim VBAWeekNum As Integer = DatePart(DateInterval.WeekOfYear, Date.Today, FirstDayOfWeek.Monday, FirstWeekOfYear.FirstFourDays)

                If Len(VBAWeekNum) = 1 Then VBAWeekNum = "0" & VBAWeekNum

                date_2 = Date.Now.ToString("yy") & VBAWeekNum & DateAndTime.Weekday(DateTime.Now, vbMonday)

                Dim _ID As String = Convert.ToDecimal(seq_SerialNumber.ToString).ToString("0000")
                Dim _seq As String = "SG" & date_2 & _ID

                cmd.Parameters.AddWithValue("@Seq", _seq)

                'cmd.Parameters.AddWithValue("@Seq", seq_SerialNumber.ToString)
                cmd.ExecuteNonQuery()
            Else
                MsgBox("The Component has not been Scanned !" & vbNewLine & "Pls Scan The Component First")
            End If



        ElseIf selectedLabel.Text.ToLower = "loose2" And Convert.ToDecimal(Me.labelQty.Text) >= 1 Then
            '===============================================================================================
            If Me.checkCombine.Checked = True Then 'One product label which combine all the data

                Dim PrintCount As Integer = 0
                Label6_SetValue()
                PrintCount = 0
                For i = 1 To Me.labelQty.Text - 1

                    label6_printer.Variables("label13").SetValue(Me.boxQty.Text)
                    set_for_combine(Convert.ToDecimal(Me.boxQty.Text))
                    'printing with quantity 
                    Dim qty As Integer = 1
                    Try
                        label6_printer.Print(qty)
                    Catch ex As Exception
                        MsgBox("Printing Label Cancel " & ex.Message)
                    End Try

                    PrintCount = PrintCount + CInt(Me.boxQty.Text)

                Next

                'label6_printer.Variables("label13").SetValue(Convert.ToDecimal(Me.boxQty.Text) - PrintCount)
                label6_printer.Variables("label13").SetValue(Convert.ToDecimal(Me.Quantity.Text) - PrintCount)
                set_for_combine(Convert.ToDecimal(Me.Quantity.Text) - PrintCount)

                Try
                    label6_printer.Print(1)
                Catch ex As Exception
                    MsgBox("Printing Label Cancel " & ex.Message)
                End Try

                '===================================================================================================
            ElseIf Me.checkCombine.Checked = False Then
                Dim maxvalue As Integer = 0
                Dim i As Integer = 0
                '        firstSerial = checkSerialNumber(Quantity)

                '        yearNum = Format(Of Date, "yy")
                '        thisweek = Format(Of Date, "ww")
                '        If Len(thisweek) = 1 Then thisweek = "0" & thisweek
                '        dayNum = Weekday(Of Date, vbMonday)

                '        i = startQty
                '        j = firstSerial
                maxvalue = 1 + Convert.ToDecimal(Quantity.Text)
                '        serial = 0

                Dim indx As Integer = 0

                If description.Text.ToLower.IndexOf("pendant") >= 0 Or description.Text.ToLower.IndexOf("lc") >= 0 Then
                    For indx = Convert.ToDecimal(StartLabel.Text) To Convert.ToDecimal(quantityLabel.Text)
                        Label6_SetValue()
                        label6_printer.Variables("label13").SetValue(1)


                        Try
                            label6_printer.Print(1)
                        Catch ex As Exception
                            MsgBox("Printing Label Cancel " & ex.Message)
                        End Try
                    Next
                Else
                    Do While i < maxvalue

                        '            If Len(j) = 1 Then serial = "000" & j
                        '            If Len(j) = 2 Then serial = "00" & j
                        '            If Len(j) = 3 Then serial = "0" & j

                        '            traceability = "SG" & yearNum & thisweek & dayNum & serial

                        '            If chosenLabel = "generic" Or chosenLabel = "loose" Then
                        '                Reports(Trim(chosenLabel))![ItemNumber].Caption = i
                        '                Reports(Trim(chosenLabel))![serialNum].Caption = traceability
                        '                Reports(Trim(chosenLabel))![Text27].Caption = Me.Quantity
                        '            End If

                        Label6_SetValue()
                        label6_printer.Variables("label13").SetValue(1)

                        If (PPnumberEntry.Text.Length >= 10 And Me.DataGridView1.Rows.Count <= 0) Or select_quantity.Checked = True Then
                            label6_printer.Variables("label13").SetValue(Me.Quantity.Text)
                            i = maxvalue
                        End If

                        Try
                            label6_printer.Print(1)
                        Catch ex As Exception
                            MsgBox("Printing Label Cancel " & ex.Message)
                        End Try
                        i = i + 1
                        '            j = j + 1

                    Loop
                End If
            End If


        End If


    End Sub

    'set for combine label only
    Private Sub set_for_combine(a As Integer)
        Dim vqty1 As Integer
        Dim vqty2 As Integer
        Dim vqty3 As Integer
        Dim vqty4 As Integer
        Dim vqty5 As Integer
        Dim vqty6 As Integer
        Dim vqty7 As Integer
        Dim vqty8 As Integer
        Dim vqty9 As Integer
        Dim vqty10 As Integer

        label6_printer.Variables("qty1").SetValue(" ")
        label6_printer.Variables("qty2").SetValue(" ")
        label6_printer.Variables("qty3").SetValue(" ")
        label6_printer.Variables("qty4").SetValue(" ")
        label6_printer.Variables("qty5").SetValue(" ")
        label6_printer.Variables("qty6").SetValue(" ")
        label6_printer.Variables("qty7").SetValue(" ")
        label6_printer.Variables("qty8").SetValue(" ")
        label6_printer.Variables("qty9").SetValue(" ")
        label6_printer.Variables("qty10").SetValue(" ")

        'If qty1.Text <> "" Then vqty1 = a * (Convert.ToDecimal(qty1.Text) / Convert.ToDecimal(LabelQuantityitem.Text))
        'If qty2.Text <> "" Then vqty2 = a * (Convert.ToDecimal(qty2.Text) / Convert.ToDecimal(LabelQuantityitem.Text))
        'If qty3.Text <> "" Then vqty3 = a * (Convert.ToDecimal(qty3.Text) / Convert.ToDecimal(LabelQuantityitem.Text))
        'If qty4.Text <> "" Then vqty4 = a * (Convert.ToDecimal(qty4.Text) / Convert.ToDecimal(LabelQuantityitem.Text))
        'If qty5.Text <> "" Then vqty5 = a * (Convert.ToDecimal(qty5.Text) / Convert.ToDecimal(LabelQuantityitem.Text))
        'If qty6.Text <> "" Then vqty6 = a * (Convert.ToDecimal(qty6.Text) / Convert.ToDecimal(LabelQuantityitem.Text))
        'If qty7.Text <> "" Then vqty7 = a * (Convert.ToDecimal(qty7.Text) / Convert.ToDecimal(LabelQuantityitem.Text))
        'If qty8.Text <> "" Then vqty8 = a * (Convert.ToDecimal(qty8.Text) / Convert.ToDecimal(LabelQuantityitem.Text))
        'If qty9.Text <> "" Then vqty9 = a * (Convert.ToDecimal(qty9.Text) / Convert.ToDecimal(LabelQuantityitem.Text))
        'If qty10.Text <> "" Then vqty10 = a * (Convert.ToDecimal(qty10.Text) / Convert.ToDecimal(LabelQuantityitem.Text))

        '=IIf([Forms]![Main]![qty1]="","",([Forms]![Main]![qty1]/[Forms]![Main]![quantity])*[Label13].[Caption])

        If qty1.Text <> "" Then vqty1 = a * (Convert.ToDecimal(qty1.Text) / Convert.ToDecimal(Quantity.Text))
        If qty2.Text <> "" Then vqty2 = a * (Convert.ToDecimal(qty2.Text) / Convert.ToDecimal(Quantity.Text))
        If qty3.Text <> "" Then vqty3 = a * (Convert.ToDecimal(qty3.Text) / Convert.ToDecimal(Quantity.Text))
        If qty4.Text <> "" Then vqty4 = a * (Convert.ToDecimal(qty4.Text) / Convert.ToDecimal(Quantity.Text))
        If qty5.Text <> "" Then vqty5 = a * (Convert.ToDecimal(qty5.Text) / Convert.ToDecimal(Quantity.Text))
        If qty6.Text <> "" Then vqty6 = a * (Convert.ToDecimal(qty6.Text) / Convert.ToDecimal(Quantity.Text))
        If qty7.Text <> "" Then vqty7 = a * (Convert.ToDecimal(qty7.Text) / Convert.ToDecimal(Quantity.Text))
        If qty8.Text <> "" Then vqty8 = a * (Convert.ToDecimal(qty8.Text) / Convert.ToDecimal(Quantity.Text))
        If qty9.Text <> "" Then vqty9 = a * (Convert.ToDecimal(qty9.Text) / Convert.ToDecimal(Quantity.Text))
        If qty10.Text <> "" Then vqty10 = a * (Convert.ToDecimal(qty10.Text) / Convert.ToDecimal(Quantity.Text))

        If vqty1 <> 0 Then label6_printer.Variables("qty1").SetValue(vqty1)
        If vqty2 <> 0 Then label6_printer.Variables("qty2").SetValue(vqty2)
        If vqty3 <> 0 Then label6_printer.Variables("qty3").SetValue(vqty3)
        If vqty4 <> 0 Then label6_printer.Variables("qty4").SetValue(vqty4)
        If vqty5 <> 0 Then label6_printer.Variables("qty5").SetValue(vqty5)
        If vqty6 <> 0 Then label6_printer.Variables("qty6").SetValue(vqty6)
        If vqty7 <> 0 Then label6_printer.Variables("qty7").SetValue(vqty7)
        If vqty8 <> 0 Then label6_printer.Variables("qty8").SetValue(vqty8)
        If vqty9 <> 0 Then label6_printer.Variables("qty9").SetValue(vqty9)
        If vqty10 <> 0 Then label6_printer.Variables("qty10").SetValue(vqty10)
    End Sub


    'Packaging label Printing ..................
    Private Sub Packaging_label_print()
        'Set to Variable of NiceLabel
        '
        'load custemer

        'If cust_TextBox1.Text <> "" Then customer.Text = cust_TextBox1.Text  'dikomen santo

        label1_setValue()

        Dim a As Integer
        'Dim sisa As Integer
        Dim quantity_show As Integer

        For a = Convert.ToDecimal(StartPackingLabel.Text) To Convert.ToDecimal(labelQty.Text)

            If Convert.ToDecimal(Quantity.Text) <= Convert.ToDecimal(boxQty.Text) Then
                quantity_show = Convert.ToDecimal(Quantity.Text)
            Else
                'sisa = Convert.ToDecimal(Quantity.Text) - (a * Convert.ToDecimal(boxQty.Text))
                If (a * Convert.ToDecimal(boxQty.Text)) > Convert.ToDecimal(Quantity.Text) Then
                    quantity_show = Convert.ToDecimal(Quantity.Text) Mod Convert.ToDecimal(boxQty.Text)
                Else
                    quantity_show = Convert.ToDecimal(boxQty.Text)
                End If
            End If

            label1_printer.Variables("quantity").SetValue(quantity_show.ToString)
            label1_printer.Variables("Package").SetValue(a.ToString)

            'printing with quantity 
            'Dim qty As Integer = 1
            Try
                label1_printer.Print(1)
                Call koneksi_db()
                Dim sqlNew As String = "insert into [printingRecordPacking] ([PP],[print date],[print time], [user]) values ('" & Me.PP.Text & "',getDate(),getDate(),'" & Me.user.Text & "')"
                Dim cmdNew = New SqlCommand(sqlNew, Main.koneksi)
                cmdNew.ExecuteNonQuery()
            Catch ex As Exception
                MsgBox("Printing Packaging Cancel " & ex.Message)
            End Try

        Next


    End Sub
    Private Sub PrintLabel_Click() Handles printLabel.Click

        var_PrintLabel_klik = True

        Dim a As Integer = 0

        If PPnumberEntry.Text.Length >= 10 And Me.DataGridView1.Rows.Count <= 0 Then
            a = 1
        End If



        If Convert.ToDecimal(CounterItems.Text) > 0 Or a = 1 Then

            If CheckProductLabelPrinting.Checked = True Then
                If checkDuplicatePP() = False Then
                    Product_Label_print()
                End If
            End If

            If checkPackagingPrinting.Checked = True Then
                If checkDuplicatePPpacking() = False Then
                    If String.IsNullOrEmpty(country.Text) Then
                        MsgBox("Country is empty, Please Select First!")
                        Report_Tab.SelectedIndex = 1
                        Exit Sub
                    Else
                        Packaging_label_print()
                    End If
                End If
            End If
        Else
            MsgBox("This Product not yet printed!" & vbNewLine & "You have to scan it first")
        End If


    End Sub

    Private Sub listprinter1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles listprinter1.SelectedIndexChanged
        'Selection of printer
        If printers.Count > 0 Then
            selected_Printer = printers.Item(listprinter1.SelectedIndex)
            label1_printer.PrintSettings.PrinterName = selected_Printer.Name
            listprinter1.SelectedItem = selected_Printer.Name
        End If
    End Sub

    Private Sub listprinter2_SelectionChangeCommitted(sender As Object, e As EventArgs) Handles listprinter2.SelectionChangeCommitted
        'Selection of printer
        If printers.Count > 0 Then
            selected_Printer = printers.Item(listprinter2.SelectedIndex)
            label2_printer.PrintSettings.PrinterName = selected_Printer.Name
            listprinter2.SelectedItem = selected_Printer.Name
        End If
    End Sub

    Private Sub listprinter3_SelectionChangeCommitted(sender As Object, e As EventArgs) Handles listprinter3.SelectionChangeCommitted
        'Selection of printer
        If printers.Count > 0 Then
            selected_Printer = printers.Item(listprinter3.SelectedIndex)
            label3_printer.PrintSettings.PrinterName = selected_Printer.Name
            listprinter3.SelectedItem = selected_Printer.Name
        End If
    End Sub
    'santo
    'label 1 setvalue
    Private Sub label1_setValue()
        Application.DoEvents()
        Me.Refresh()

        label1_printer.Variables("Product ID").SetValue(Microsoft.VisualBasic.Left(Material.Text, 19))
        'label1_printer.Variables("custMaterial").SetValue(Microsoft.VisualBasic.Left(custMaterial.Text, 19))
        label1_printer.Variables("custMaterial").SetValue(custMaterial.Text)
        label1_printer.Variables("Product Desc Eng").SetValue(Microsoft.VisualBasic.Left(packingDescription.Text, 30))
        label1_printer.Variables("PP Number").SetValue(PP.Text)
        label1_printer.Variables("SO Number").SetValue(SO.Text)
        label1_printer.Variables("Cust PO").SetValue(custPO.Text)

        If CRD.Text <> "00/00/0000" Then
            label1_printer.Variables("Req Date").SetValue(Microsoft.VisualBasic.Left(CRD.Text, 10))
        Else
            MsgBox("Request Date empty !")
            label1_printer.Variables("Req Date").SetValue(" ")
        End If
        'aaa
        label1_printer.Variables("TotalQuantity").SetValue(Microsoft.VisualBasic.Left(Quantity.Text, 4))
        label1_printer.Variables("country").SetValue(Microsoft.VisualBasic.Left(country.Text, 11))
        label1_printer.Variables("Cust code").SetValue(Microsoft.VisualBasic.Left(Customer_code.Text, 12))
        'label1_printer.Variables("labelQty").SetValue(labelQty.Text)
        label1_printer.Variables("labelQty").SetValue(Microsoft.VisualBasic.Left(Label468.Text, 3))
        'label1_printer.Variables("Customer").SetValue(customer.SelectedText)

        Try
            'label1_printer.Variables("Customer").SetValue(customer.SelectedValue.ToString)
            label1_printer.Variables("Customer").SetValue(Microsoft.VisualBasic.Left(customer.Text, 35))
        Catch ex As Exception
            MsgBox("Customer No Found")
        End Try

        Try
            label1_printer.Variables("user").SetValue(Microsoft.VisualBasic.Left(user.Text, 18))
        Catch ex As Exception
            MsgBox("user No Found")
        End Try

        label1_printer.Variables("item").SetValue(SO_item.Text)

        'Try
        '    label1_printer.Variables("Package").SetValue(pack.SelectedValue.ToString)
        'Catch ex As Exception
        '    MsgBox("Package Not Selected !")
        'End Try

        label1_printer.Variables("countryShortName").SetValue(Microsoft.VisualBasic.Left(countryShortName.Text, 2))
        'label1_printer.Variables("PP Number").SetValue(Customer_code.Text)

        Try
            'label1_printer.Variables("SO_Barcode").SetValue("0" & SO.Text & "000" & Convert.ToDecimal(SO_item.Text).ToString("000"))
            label1_printer.Variables("SO_Barcode").SetValue("0" & SO.Text & Convert.ToDecimal(SO_item.Text).ToString("000000"))
            'MessageBox.Show("0" & SO.Text & "000" & Convert.ToDecimal(SO_item.Text).ToString("000"))
        Catch ex As Exception
            'MessageBox.Show("Packaging Label Need to Update for SO Baarcode !")
        End Try

    End Sub
    Private Sub Command29_Click(sender As Object, e As EventArgs) Handles Command29.Click

        Me.labelQty2 = Me.labelQty
        Me.StartPackingLabel2 = Me.StartPackingLabel

        ''If checkAllDataOk() = False Then donothing
        'If checkAllDataOk() Then
        'End If
        ''Printing........
        'If CheckProductLabelPrinting2.Checked = True Then Product_Label_print()

        'If checkPackagingPrinting2.Checked = True Then Packaging_label_print()

        PrintLabel_Click()


    End Sub
    'santo
    'label2 setvalue
    Public Const do_check As String = "{\rtf1\deff0{\fonttbl{\f0 Calibri;}{\f1 Times New Roman;}{\f2\fcharset2 Wingdings 2;}}{\colortbl ;\red0\green0\blue255 ;}{\*\defchp \f1}{\stylesheet {\ql\f1 Normal;}{\*\cs1\f1 Default Paragraph Font;}{\*\cs2\sbasedon1\f1 Line Number;}{\*\cs3\ul\f1\cf1 Hyperlink;}{\*\ts4\tsrowd\f1\ql\tscellpaddfl3\tscellpaddl108\tscellpaddfb3\tscellpaddfr3\tscellpaddr108\tscellpaddft3\tsvertalt\cltxlrtb Normal Table;}{\*\ts5\tsrowd\sbasedon4\f1\ql\trbrdrt\brdrs\brdrw10\trbrdrl\brdrs\brdrw10\trbrdrb\brdrs\brdrw10\trbrdrr\brdrs\brdrw10\trbrdrh\brdrs\brdrw10\trbrdrv\brdrs\brdrw10\tscellpaddfl3\tscellpaddl108\tscellpaddfr3\tscellpaddr108\tsvertalt\cltxlrtb Table Simple 1;}}{\*\listoverridetable}\nouicompat\splytwnine\htmautsp\sectd\marglsxn1440\margrsxn1440\margtsxn1440\margbsxn1440\pard\plain\ql{\field{\*\fldinst{\f1\fs32\cf0 SYMBOL 82 \\f ""Wingdings 2"" \\s 11}}{\fldrslt{\f2\fs22\cf0 R}}}\f1\par}"
    Public Const no_check As String = "{\rtf1\deff0{\fonttbl{\f0 Calibri;}{\f1 Times New Roman;}{\f2\fcharset2 Wingdings 2;}}{\colortbl ;\red0\green0\blue255 ;}{\*\defchp \f1}{\stylesheet {\ql\f1 Normal;}{\*\cs1\f1 Default Paragraph Font;}{\*\cs2\sbasedon1\f1 Line Number;}{\*\cs3\ul\f1\cf1 Hyperlink;}{\*\ts4\tsrowd\f1\ql\tscellpaddfl3\tscellpaddl108\tscellpaddfb3\tscellpaddfr3\tscellpaddr108\tscellpaddft3\tsvertalt\cltxlrtb Normal Table;}{\*\ts5\tsrowd\sbasedon4\f1\ql\trbrdrt\brdrs\brdrw10\trbrdrl\brdrs\brdrw10\trbrdrb\brdrs\brdrw10\trbrdrr\brdrs\brdrw10\trbrdrh\brdrs\brdrw10\trbrdrv\brdrs\brdrw10\tscellpaddfl3\tscellpaddl108\tscellpaddfr3\tscellpaddr108\tsvertalt\cltxlrtb Table Simple 1;}}{\*\listoverridetable}\nouicompat\splytwnine\htmautsp\sectd\marglsxn1440\margrsxn1440\margtsxn1440\margbsxn1440\pard\plain\ql{\field{\*\fldinst{\f1\cf0 SYMBOL 163 \\f ""Wingdings 2"" \\s 11}}{\fldrslt{\f2\fs22\cf0 \u163\'a3}}}\f1\par}"
    Private Sub Label2_setValue()
        Application.DoEvents()
        'label2_printer.Variables("testQuantity").SetValue(testQuantity.Text)
        label2_printer.Variables("testQuantity").SetValue(Quantity.Text)
        label2_printer.Variables("StartTestQuantity").SetValue(StartTestQuantity.Text)

        label2_printer.Variables("testCustomer").SetValue(testCustomer.Text)
        'label2_printer.Variables("testCustomer").SetValue(testCustomer.SelectedValue.ToString)
        label2_printer.Variables("testMaterial").SetValue(testMaterial.Text)
        label2_printer.Variables("testDescription").SetValue(testDescription.Text)
        label2_printer.Variables("testPP").SetValue(testPP.Text)
        label2_printer.Variables("testSO").SetValue(testSO.Text)
        label2_printer.Variables("testSOitem").SetValue(testSOitem.Text)
        label2_printer.Variables("testCustPO").SetValue(testCustPO.Text)
        label2_printer.Variables("testCustPOitem").SetValue(testCustPOitem.Text)
        label2_printer.Variables("Quantity").SetValue(Quantity.Text)
        label2_printer.Variables("breakerref").SetValue(breakerRef.Text)
        label2_printer.Variables("breakerdevice").SetValue(breakerDevice.Text)
        label2_printer.Variables("breakertype").SetValue(breakerType.Text)
        label2_printer.Variables("breakerpole").SetValue(breakerPole.Text)
        label2_printer.Variables("drawoutFixed").SetValue(drawoutFixed.Text)
        label2_printer.Variables("micrologicref").SetValue(micrologicRef.Text)
        label2_printer.Variables("micrologictype").SetValue(micrologicType.Text)
        label2_printer.Variables("plugref").SetValue(plugRef.Text)
        label2_printer.Variables("plugtype").SetValue(plugType.Text)
        label2_printer.Variables("chassisref").SetValue(chassisRef.Text)
        label2_printer.Variables("topConnect").SetValue(topConnect.Text)
        label2_printer.Variables("bottomConnect").SetValue(bottomConnect.Text)
        label2_printer.Variables("MCH").SetValue(MCH.Text)

        Dim z As String = ""
        Dim x As String = ""
        If XFtype.Text <> "" Then z = " "
        If MXtype.Text <> "" Then x = " "

        label2_printer.Variables("XF").SetValue(XFtype.Text & z & XF.Text)
        label2_printer.Variables("MX").SetValue(MXtype.Text & x & MX.Text)
        label2_printer.Variables("MN").SetValue(MN.Text)
        label2_printer.Variables("MX2").SetValue(MX2.Text)
        label2_printer.Variables("remoteReset").SetValue(remoteReset.Text)
        label2_printer.Variables("SDE2count").SetValue(SDE2count.Text)
        label2_printer.Variables("OF").SetValue(IC_OF.Text)
        label2_printer.Variables("SD").SetValue(SD.Text)
        label2_printer.Variables("PF").SetValue(PF.Text)
        label2_printer.Variables("CE").SetValue(CE.Text)
        label2_printer.Variables("CD").SetValue(CD.Text)
        label2_printer.Variables("CT").SetValue(CT.Text)
        label2_printer.Variables("cluster").SetValue(cluster.Text)
        label2_printer.Variables("fixingScrew").SetValue(fixingScrew.Text)
        label2_printer.Variables("text584").SetValue(Text584.Text)
        label2_printer.Variables("text586").SetValue(Text586.Text)
        label2_printer.Variables("text588").SetValue(Text588.Text)

        label2_printer.Variables("combo578").SetValue(Combo578.Text)
        label2_printer.Variables("combo580").SetValue(Combo580.Text)
        label2_printer.Variables("combo582").SetValue(Combo582.Text)


        'label2_printer.Variables("XFtype").SetValue(XFtype.Text)
        'label2_printer.Variables("MXtype").SetValue(MXtype.Text)


        If SDE1.Checked = True Then
            label2_printer.Variables("SDE1").SetValue(do_check)
        Else
            label2_printer.Variables("SDE1").SetValue(no_check)
        End If

        If checkStandard.Checked = True Then
            label2_printer.Variables("checkStandard").SetValue(do_check)
        Else
            label2_printer.Variables("checkStandard").SetValue(no_check)
        End If

        If checkLow.Checked = True Then
            label2_printer.Variables("checkLow").SetValue(do_check)
        Else
            label2_printer.Variables("checkLow").SetValue(no_check)
        End If

        If checkHigh.Checked = True Then
            label2_printer.Variables("checkHigh").SetValue(do_check)
        Else
            label2_printer.Variables("checkHigh").SetValue(no_check)
        End If

        If checkLT.Checked = True Then
            label2_printer.Variables("checkLT").SetValue(do_check)
        Else
            label2_printer.Variables("checkLT").SetValue(no_check)
        End If

        If checkVBP.Checked = True Then
            label2_printer.Variables("checkVBP").SetValue(do_check)
        Else
            label2_printer.Variables("checkVBP").SetValue(no_check)
        End If

        If checkVCPO.Checked = True Then
            label2_printer.Variables("checkVCPO").SetValue(do_check)
        Else
            label2_printer.Variables("checkVCPO").SetValue(no_check)
        End If

        If checkVSPO.Checked = True Then
            label2_printer.Variables("checkVSPO").SetValue(do_check)
        Else
            label2_printer.Variables("checkVSPO").SetValue(no_check)
        End If

        If checkVSPDchassis.Checked = True Then
            label2_printer.Variables("checkVSPDchassis").SetValue(do_check)
        Else
            label2_printer.Variables("checkVSPDchassis").SetValue(no_check)
        End If

        If checkVSPDall.Checked = True Then
            label2_printer.Variables("checkVSPDall").SetValue(do_check)
        Else
            label2_printer.Variables("checkVSPDall").SetValue(no_check)
        End If

        If checkVPEC.Checked = True Then
            label2_printer.Variables("checkVPEC").SetValue(do_check)
        Else
            label2_printer.Variables("checkVPEC").SetValue(no_check)
        End If

        If checkVPOC.Checked = True Then
            label2_printer.Variables("checkVPOC").SetValue(do_check)
        Else
            label2_printer.Variables("checkVPOC").SetValue(no_check)
        End If

        If checkDAE.Checked = True Then
            label2_printer.Variables("checkDAE").SetValue(do_check)
        Else
            label2_printer.Variables("checkDAE").SetValue(no_check)
        End If

        If checkVDC.Checked = True Then
            label2_printer.Variables("checkVDC").SetValue(do_check)
        Else
            label2_printer.Variables("checkVDC").SetValue(no_check)
        End If

        If checkIPA.Checked = True Then
            label2_printer.Variables("checkIPA").SetValue(do_check)
        Else
            label2_printer.Variables("checkIPA").SetValue(no_check)
        End If

        If checkIPBO.Checked = True Then
            label2_printer.Variables("checkIPBO").SetValue(do_check)
        Else
            label2_printer.Variables("checkIPBO").SetValue(no_check)
        End If

        If checkVICV.Checked = True Then
            label2_printer.Variables("checkVICV").SetValue(do_check)
        Else
            label2_printer.Variables("checkVICV").SetValue(no_check)
        End If

        If checkCDM.Checked = True Then
            label2_printer.Variables("checkCDM").SetValue(do_check)
        Else
            label2_printer.Variables("checkCDM").SetValue(no_check)
        End If

        If checkCB.Checked = True Then
            label2_printer.Variables("checkCB").SetValue(do_check)
        Else
            label2_printer.Variables("checkCB").SetValue(no_check)
        End If

        If checkCDP.Checked = True Then
            label2_printer.Variables("checkCDP").SetValue(do_check)
        Else
            label2_printer.Variables("checkCDP").SetValue(no_check)
        End If


        If checkCP.Checked = True Then
            label2_printer.Variables("checkCP").SetValue(do_check)
        Else
            label2_printer.Variables("checkCP").SetValue(no_check)
        End If

        If checkOP.Checked = True Then
            label2_printer.Variables("checkOP").SetValue(do_check)
        Else
            label2_printer.Variables("checkOP").SetValue(no_check)
        End If

        If checkPhaseBarrier.Checked = True Then
            label2_printer.Variables("checkPhaseBarrier").SetValue(do_check)
        Else
            label2_printer.Variables("checkPhaseBarrier").SetValue(no_check)
        End If

        If checkRotary.Checked = True Then
            label2_printer.Variables("checkRotary").SetValue(do_check)
        Else
            label2_printer.Variables("checkRotary").SetValue(no_check)
        End If

        If checkEF.Checked = True Then
            label2_printer.Variables("checkEF").SetValue(do_check)
        Else
            label2_printer.Variables("checkEF").SetValue(no_check)
        End If

        If checkPTE.Checked = True Then
            label2_printer.Variables("checkPTE").SetValue(do_check)
        Else
            label2_printer.Variables("checkPTE").SetValue(no_check)
        End If

        If checkR.Checked = True Then
            label2_printer.Variables("checkR").SetValue(do_check)
        Else
            label2_printer.Variables("checkR").SetValue(no_check)
        End If

        If checkRr.Checked = True Then
            label2_printer.Variables("checkRr").SetValue(do_check)
        Else
            label2_printer.Variables("checkRr").SetValue(no_check)
        End If

        If checkM2C.Checked = True Then
            label2_printer.Variables("checkM2C").SetValue(do_check)
        Else
            label2_printer.Variables("checkM2C").SetValue(no_check)
        End If

        If checkCOM.Checked = True Then
            label2_printer.Variables("checkCOM").SetValue(do_check)
        Else
            label2_printer.Variables("checkCOM").SetValue(no_check)
        End If

        If checkATS.Checked = True Then
            label2_printer.Variables("checkATS").SetValue(do_check)
        Else
            label2_printer.Variables("checkATS").SetValue(no_check)
        End If

        If checkAD.Checked = True Then
            label2_printer.Variables("checkAD").SetValue(do_check)
        Else
            label2_printer.Variables("checkAD").SetValue(no_check)
        End If

        If checkBattery.Checked = True Then
            label2_printer.Variables("checkBattery").SetValue(do_check)
        Else
            label2_printer.Variables("checkBattery").SetValue(no_check)
        End If

        If checkexternalCT.Checked = True Then
            label2_printer.Variables("checkexternalCT").SetValue(do_check)
        Else
            label2_printer.Variables("checkexternalCT").SetValue(no_check)
        End If

        If checkScrew.Checked = True Then
            label2_printer.Variables("checkScrew").SetValue(do_check)
        Else
            label2_printer.Variables("checkScrew").SetValue(no_check)
        End If

        If checkLeaflet.Checked = True Then
            label2_printer.Variables("checkLeaflet").SetValue(do_check)
        Else
            label2_printer.Variables("checkLeaflet").SetValue(no_check)
        End If

        If checkLabel.Checked = True Then
            label2_printer.Variables("checkLabel").SetValue(do_check)
        Else
            label2_printer.Variables("checkLabel").SetValue(no_check)
        End If
    End Sub

    Private Sub Command209_Click(sender As Object, e As EventArgs) Handles Command209.Click

        'If Convert.ToDecimal(Me.testQuantity.Text) > Convert.ToDecimal(Me.cek_testQuantity.Text) Then
        '    testPP.Select()
        '    SendKeys.Send("{ENTER}")
        '    'Dim result1 As DialogResult = MessageBox.Show("Max Test Print Quantity is  " & Me.cek_testQuantity.Text &
        '    '                                              Chr(13) & "Continue to print?",
        '    '"Max Test Print Quantity!",
        '    'MessageBoxButtons.YesNo,
        '    'MessageBoxIcon.Question)

        '    'If result1 = DialogResult.Yes Then
        '    '    Me.testQuantity.Text = Me.cek_testQuantity.Text
        '    'Else
        '    Exit Sub
        '    'End If
        'End If

        If Convert.ToDecimal(StartTestQuantity.Text) = 0 Or Convert.ToDecimal(Me.testPrintQty.Text) = 0 Then
            MsgBox("you have to input test report print qty")
            Exit Sub
        End If

        If Convert.ToDecimal(Me.testQuantity.Text) > Convert.ToDecimal(Me.Quantity.Text) Then
            MsgBox("The Number of printing higher then the number of Quantity !")
            Me.testQuantity.Text = Me.Quantity.Text
            testQuantity.Select()
            Exit Sub
        End If

        progress_printing(30)
        'Set to Variable of NiceLabel
        Label2_setValue()
        progress_printing(50)
        ' Dim a As Integer
        For a = Convert.ToDecimal(StartTestQuantity.Text) To Convert.ToDecimal(testQuantity.Text)
            'MsgBox(testPrintQty.SelectedItem)
            'printing with quantity 
            Dim qty As Integer = Convert.ToDecimal(testPrintQty.Text)
            label2_printer.Variables("CounterItemTest").SetValue(a)
            Try
                label2_printer.Print(qty)
                cmd = New SqlCommand("insert into printingRecordTest([pp],[date],[time],[user],[from],[to]) values(@pp,@date,@time,@user,@from,@to)", Main.koneksi)
                cmd.Parameters.AddWithValue("@pp", Me.PPnumberEntry.Text)
                cmd.Parameters.AddWithValue("@date", DateTime.Now.ToString("yyyy-MM-dd"))
                cmd.Parameters.AddWithValue("@time", DateTime.Now.ToString("HH:mm:ss"))
                cmd.Parameters.AddWithValue("@user", Me.technicianShortName.Text)
                cmd.Parameters.AddWithValue("@from", Me.StartTestQuantity.Text)
                cmd.Parameters.AddWithValue("@to", Me.testQuantity.Text)
                cmd.ExecuteNonQuery()
                progress_printing(90)
            Catch ex As Exception
                MsgBox("Printing TEST Cancel " & ex.Message)
            End Try
        Next

        'For a = 1 To Convert.ToDecimal(testPrintQty.SelectedItem)
        '    'MsgBox(testPrintQty.SelectedItem)
        '    'printing with quantity 
        '    Dim qty As Integer = Convert.ToDecimal(testPrintQty.Text)
        '    label2_printer.Variables("Pint_seq").SetValue(a)
        '    Try
        '        label2_printer.Print(qty)
        '        cmd = New SqlCommand("insert into printingRecordTest([pp],[date],[time],[user],[from],[to]) values(@pp,@date,@time,@user,@from,@to)", Main.koneksi)
        '        cmd.Parameters.AddWithValue("@pp", Me.PPnumberEntry.Text)
        '        cmd.Parameters.AddWithValue("@date", DateTime.Now.ToString("yyyy-MM-dd"))
        '        cmd.Parameters.AddWithValue("@time", DateTime.Now.ToString("HH:mm:ss"))
        '        cmd.Parameters.AddWithValue("@user", Me.technicianShortName.Text)
        '        cmd.Parameters.AddWithValue("@from", Me.StartTestQuantity.Text)
        '        cmd.Parameters.AddWithValue("@to", Me.testQuantity.Text)
        '        cmd.ExecuteNonQuery()
        '    Catch ex As Exception
        '        MsgBox("Printing TEST Cancel " & ex.Message)
        '    End Try
        'Next
        progress_printing(100)
    End Sub
    Public Const COC_SD_Main As String = "We hereby certify that the switch disconnector delivered (as listed below) are assembled and tested at Schneider Electric Logistics Asia Pte Ltd Adaptation Centre. We guarantee their conformity with Schneider Electric technical specification and with the standards and regulations for switch disconnector as per IEC 60947-3."
    Public Const COC_CB_Main As String = "We hereby certify that the air circuit breaker delivered (as listed below) are assembled and tested at Schneider Electric Logistics Asia Pte Ltd Adaptation Centre. We guarantee their conformity with Schneider Electric technical specification and with the standards and regulations for circuit breaker as per IEC 60947-2."
    Public Const COC_CB_foot As String = "The Air Circuit Breaker delivered:"
    Public Const COC_SD_foot As String = "The switch disconnector delivered:"

    'Public Sub Log_data(ByVal data As String)
    '    Dim file As System.IO.StreamWriter
    '    file = My.Computer.FileSystem.OpenTextFileWriter("Log.txt", True)
    '    file.WriteLine(data)
    '    file.Close()
    'End Sub

    Dim buka_packaging_tab As Integer


    Private Sub label3_setvalue()

        'If buka_packaging_tab = 0 Then
        'Report_Tab.SelectedIndex = 1
        'Report_Tab.SelectedIndex = 2
        'buka_packaging_tab = 1
        'End If

        Try
            label3_printer.Variables("StartTestQuantity").SetValue(StartTestQuantity.Text)
        Catch ex As Exception
            'Log_data("COC Start Qty: " & ex.ToString)
        End Try

        Try
            label3_printer.Variables("testQuantity").SetValue(Quantity.Text)
        Catch ex As Exception
            'Log_data("COC Error testQty: " & ex.ToString)
        End Try


        Dim COCreport As String = ""
        'If Me.micrologicRef.Text = "" Or String.IsNullOrEmpty(Me.micrologicRef.Text) = True Then

        'Try
        '    'reset value
        '    label3_printer.Variables("COC_SD").SetValue("0")
        '    label3_printer.Variables("COC_CB").SetValue("0")
        'Catch ex As Exception
        '    Log_data("COC COC SD = 0: " & ex.ToString)
        'End Try

        'Try
        '    If Me.breakerType.Text = "NA" Or Me.breakerType.Text = "HA" Or Me.breakerType.Text = "HF" Then
        '        COCreport = "COC SD"
        '        'santo coco jinlong
        '        label3_printer.Variables("COC_SD").SetValue("1")

        '    Else
        '        COCreport = "COC CB"
        '        'santo COC jinlong
        '        label3_printer.Variables("COC_CB").SetValue("1")

        '    End If
        'Catch ex As Exception
        '    Log_data("COC COC Data in 1: " & ex.ToString)
        'End Try


        Try
            If Me.breakerType.Text = "NA" Or Me.breakerType.Text = "HA" Or Me.breakerType.Text = "HF" Or Me.breakerType.Text = "H1-S" Then
                COCreport = "COC SD"

                label3_printer.Variables("COC_Main").SetValue(COC_SD_Main)
                label3_printer.Variables("COC_foot").SetValue(COC_SD_foot)

            Else
                COCreport = "COC CB"

                label3_printer.Variables("COC_Main").SetValue(COC_CB_Main)
                label3_printer.Variables("COC_foot").SetValue(COC_CB_foot)

            End If
        Catch ex As Exception
            'Log_data("COC COC Data in 2: " & ex.ToString)
        End Try
        Try
            label3_printer.Variables("Adress").SetValue(testCustomer.Text)
            label3_printer.Variables("Product").SetValue(productRange.Text)
            label3_printer.Variables("Item").SetValue(StartTestQuantity.Text)
            label3_printer.Variables("tes Material").SetValue(Microsoft.VisualBasic.Left(testMaterial.Text, 14))
            label3_printer.Variables("tes Description").SetValue(testDescription.Text)
            label3_printer.Variables("tes SO").SetValue(testSO.Text)
            label3_printer.Variables("tes Production").SetValue(testPP.Text)
            label3_printer.Variables("tes Customer PO").SetValue(testCustPO.Text)
            label3_printer.Variables("tes SOitem").SetValue(testSOitem.Text)
            label3_printer.Variables("testCustPOitem").SetValue(testCustPOitem.Text)
        Catch ex As Exception
            'Log_data("COC last: " & ex.ToString)
        End Try

    End Sub

    Private Sub Label4_SetValue()

        label4_printer.Variables("Product ID").SetValue(header.Text)
        label4_printer.Variables("Label13").SetValue(LabelQuantityitem.Text)

        label4_printer.Variables("mat1").SetValue(mat1.Text)
        label4_printer.Variables("mat2").SetValue(mat2.Text)
        label4_printer.Variables("mat3").SetValue(mat3.Text)
        label4_printer.Variables("mat4").SetValue(mat4.Text)
        label4_printer.Variables("mat5").SetValue(mat5.Text)
        label4_printer.Variables("mat6").SetValue(mat6.Text)
        label4_printer.Variables("mat7").SetValue(mat7.Text)
        label4_printer.Variables("mat8").SetValue(mat8.Text)
        label4_printer.Variables("mat9").SetValue(mat9.Text)
        label4_printer.Variables("mat10").SetValue(mat10.Text)

        label4_printer.Variables("qty1").SetValue(qty1.Text)
        label4_printer.Variables("qty2").SetValue(qty2.Text)
        label4_printer.Variables("qty3").SetValue(qty3.Text)
        label4_printer.Variables("qty4").SetValue(qty4.Text)
        label4_printer.Variables("qty5").SetValue(qty5.Text)
        label4_printer.Variables("qty6").SetValue(qty6.Text)
        label4_printer.Variables("qty7").SetValue(qty7.Text)
        label4_printer.Variables("qty8").SetValue(qty8.Text)
        label4_printer.Variables("qty9").SetValue(qty9.Text)
        label4_printer.Variables("qty10").SetValue(qty10.Text)

        label4_printer.Variables("descr1").SetValue(descr1.Text)
        label4_printer.Variables("descr2").SetValue(descr2.Text)
        label4_printer.Variables("descr3").SetValue(descr3.Text)
        label4_printer.Variables("descr4").SetValue(descr4.Text)
        label4_printer.Variables("descr5").SetValue(descr5.Text)
        label4_printer.Variables("descr6").SetValue(descr6.Text)
        label4_printer.Variables("descr7").SetValue(descr7.Text)
        label4_printer.Variables("descr8").SetValue(descr8.Text)
        label4_printer.Variables("descr9").SetValue(descr9.Text)
        label4_printer.Variables("descr10").SetValue(descr10.Text)

        label4_printer.Variables("warning").SetValue(warning.Text)
        label4_printer.Variables("Made in singapore").SetValue(madeInSingapore.Text)
        Try
            label4_printer.Variables("country").SetValue(madeInEnglish.SelectedValue.ToString)
        Catch ex As Exception
            MsgBox("Made In not selected !")
            'If range.Text.IndexOf("CVS") <> -1 Or range.Text.IndexOf("NSX") Or range.Text.IndexOf("NT") Then
            '    madeInEnglish.Text = "Made in China"
            'Else
            '    madeInEnglish.Text = "Made in India"
            'End If
            'label4_printer.Variables("country").SetValue(madeInEnglish.SelectedValue.ToString)
        End Try
        label4_printer.Variables("date code").SetValue("SG-" & dateCode.Text)
        label4_printer.Variables("PPnumberEntry").SetValue(PPnumberEntry.Text)
        label4_printer.Variables("technicianShortName").SetValue(technicianShortName.Text)

        label4_printer.Variables("PP Number").SetValue(PPnumberEntry.Text)
        label4_printer.Variables("LabelQuantityitem").SetValue(LabelQuantityitem.Text)
        label4_printer.Variables("CounterItems").SetValue(CounterItems.Text)

        label4_printer.Variables("header").SetValue(header.Text)
        label4_printer.Variables("Made in Ch").SetValue(madeInChinese.Text)
        label4_printer.Variables("Made in Rus").SetValue(madeInRussian.Text)
        Dim SG_sign As String
        Dim test As String = Microsoft.VisualBasic.Right(dateCode.Text, 2)
        If dateCode.Text(7) <> "-" Then

            'If test.IndexOf("-") <> -1 Then
            'SG_sign = "SG" & dateCode.Text(2) & dateCode.Text(3) & dateCode.Text(6) & dateCode.Text(7) & test
            'Else
            SG_sign = "SG" & dateCode.Text(2) & dateCode.Text(3) & dateCode.Text(6) & dateCode.Text(7) & Microsoft.VisualBasic.Right(dateCode.Text, 1) & "0"
            ' End If
        Else

            ' If test.IndexOf("-") <> -1 Then
            SG_sign = "SG" & dateCode.Text(2) & dateCode.Text(3) & "0" & dateCode.Text(6) & Microsoft.VisualBasic.Right(dateCode.Text, 1) & "0"
            ' Else
            'SG_sign = "SG" & dateCode.Text(2) & dateCode.Text(3) & "0" & dateCode.Text(6) & test
        End If
        label4_printer.Variables("SG Sign").SetValue(SG_sign)

        label4_printer.Variables("Logo1").SetValue(logo1value.Text)
        label4_printer.Variables("Logo2").SetValue(logo2value.Text)
        label4_printer.Variables("Logo3").SetValue(logo3value.Text)

        Try
            label_printer.Variables("Logo4").SetValue(logo4value.Text)
            label_printer.Variables("Logo5").SetValue(logo5value.Text)
            label_printer.Variables("Logo6").SetValue(logo6value.Text)
        Catch ex As Exception
            'Log_data("Logo :" & DateTime.Today.ToString & ex.ToString)
        End Try



        'compare Neutral L
        'Dim strList1 As List(Of String) = New List(Of String)(New String() {"47595", "47596", "47597", "LV847595", "LV847596", "LV847597"})
        Dim Neutral As String
        'Dim pole As String
        Neutral = Material.Text
        Try
            If Neutral = "47595" Or Neutral = "47596" Or Neutral = "47597" Or Neutral = "LV847595" Or Neutral = "LV847596" Or Neutral = "LV847597" Then
                label4_printer.Variables("Neutral").SetValue("R")
            ElseIf instr(description.Text, "4P") Then
                label4_printer.Variables("Neutral").SetValue("L")
            Else
                label4_printer.Variables("Neutral").SetValue("")
            End If
        Catch ex As Exception

        End Try



    End Sub

    Private Sub Label5_SetValue()
        Dim vqty1 As Integer
        Dim vqty2 As Integer
        Dim vqty3 As Integer
        Dim vqty4 As Integer
        Dim vqty5 As Integer
        Dim vqty6 As Integer
        Dim vqty7 As Integer
        Dim vqty8 As Integer
        Dim vqty9 As Integer
        Dim vqty10 As Integer

        label5_printer.Variables("qty1").SetValue(" ")
        label5_printer.Variables("qty2").SetValue(" ")
        label5_printer.Variables("qty3").SetValue(" ")
        label5_printer.Variables("qty4").SetValue(" ")
        label5_printer.Variables("qty5").SetValue(" ")
        label5_printer.Variables("qty6").SetValue(" ")
        label5_printer.Variables("qty7").SetValue(" ")
        label5_printer.Variables("qty8").SetValue(" ")
        label5_printer.Variables("qty9").SetValue(" ")
        label5_printer.Variables("qty10").SetValue(" ")

        If Not String.IsNullOrEmpty(mat1.Text) And qty1.Text <> "" And Not String.IsNullOrEmpty(qty1.Text) And Not String.IsNullOrWhiteSpace(qty1.Text) Then vqty1 = Convert.ToDecimal(qty1.Text) / Convert.ToDecimal(LabelQuantityitem.Text)
        If Not String.IsNullOrEmpty(mat2.Text) And qty2.Text <> "" And Not String.IsNullOrEmpty(qty2.Text) And Not String.IsNullOrWhiteSpace(qty2.Text) Then vqty2 = Convert.ToDecimal(qty2.Text) / Convert.ToDecimal(LabelQuantityitem.Text)
        If Not String.IsNullOrEmpty(mat3.Text) And qty3.Text <> "" And Not String.IsNullOrEmpty(qty3.Text) And Not String.IsNullOrWhiteSpace(qty3.Text) Then vqty3 = Convert.ToDecimal(qty3.Text) / Convert.ToDecimal(LabelQuantityitem.Text)
        If Not String.IsNullOrEmpty(mat4.Text) And qty4.Text <> "" And Not String.IsNullOrEmpty(qty4.Text) And Not String.IsNullOrWhiteSpace(qty4.Text) Then vqty4 = Convert.ToDecimal(qty4.Text) / Convert.ToDecimal(LabelQuantityitem.Text)
        If Not String.IsNullOrEmpty(mat5.Text) And qty5.Text <> "" And Not String.IsNullOrEmpty(qty5.Text) And Not String.IsNullOrWhiteSpace(qty5.Text) Then vqty5 = Convert.ToDecimal(qty5.Text) / Convert.ToDecimal(LabelQuantityitem.Text)
        If Not String.IsNullOrEmpty(mat6.Text) And qty6.Text <> "" And Not String.IsNullOrEmpty(qty6.Text) And Not String.IsNullOrWhiteSpace(qty6.Text) Then vqty6 = Convert.ToDecimal(qty6.Text) / Convert.ToDecimal(LabelQuantityitem.Text)
        If Not String.IsNullOrEmpty(mat7.Text) And qty7.Text <> "" And Not String.IsNullOrEmpty(qty7.Text) And Not String.IsNullOrWhiteSpace(qty7.Text) Then vqty7 = Convert.ToDecimal(qty7.Text) / Convert.ToDecimal(LabelQuantityitem.Text)
        If Not String.IsNullOrEmpty(mat8.Text) And qty8.Text <> "" And Not String.IsNullOrEmpty(qty8.Text) And Not String.IsNullOrWhiteSpace(qty8.Text) Then vqty8 = Convert.ToDecimal(qty8.Text) / Convert.ToDecimal(LabelQuantityitem.Text)
        If Not String.IsNullOrEmpty(mat9.Text) And qty9.Text <> "" And Not String.IsNullOrEmpty(qty9.Text) And Not String.IsNullOrWhiteSpace(qty9.Text) Then vqty9 = Convert.ToDecimal(qty9.Text) / Convert.ToDecimal(LabelQuantityitem.Text)
        If Not String.IsNullOrEmpty(mat10.Text) And qty10.Text <> "" And Not String.IsNullOrEmpty(qty10.Text) And Not String.IsNullOrWhiteSpace(qty10.Text) Then vqty10 = Convert.ToDecimal(qty10.Text) / Convert.ToDecimal(LabelQuantityitem.Text)

        label5_printer.Variables("Header").SetValue(Microsoft.VisualBasic.Left(header.Text, 16))
        label5_printer.Variables("description").SetValue(description.Text)
        label5_printer.Variables("PPNumberEntry").SetValue(PPnumberEntry.Text)

        label5_printer.Variables("mat1").SetValue(Microsoft.VisualBasic.Left(mat1.Text, 14))
        label5_printer.Variables("mat2").SetValue(Microsoft.VisualBasic.Left(mat2.Text, 14))
        label5_printer.Variables("mat3").SetValue(Microsoft.VisualBasic.Left(mat3.Text, 14))
        label5_printer.Variables("mat4").SetValue(Microsoft.VisualBasic.Left(mat4.Text, 14))
        label5_printer.Variables("mat5").SetValue(Microsoft.VisualBasic.Left(mat5.Text, 14))
        label5_printer.Variables("mat6").SetValue(Microsoft.VisualBasic.Left(mat6.Text, 14))
        label5_printer.Variables("mat7").SetValue(Microsoft.VisualBasic.Left(mat7.Text, 14))
        label5_printer.Variables("mat8").SetValue(Microsoft.VisualBasic.Left(mat8.Text, 14))
        label5_printer.Variables("mat9").SetValue(Microsoft.VisualBasic.Left(mat9.Text, 14))
        label5_printer.Variables("mat10").SetValue(Microsoft.VisualBasic.Left(mat10.Text, 14))

        If vqty1 <> 0 Then label5_printer.Variables("qty1").SetValue(vqty1)
        If vqty2 <> 0 Then label5_printer.Variables("qty2").SetValue(vqty2)
        If vqty3 <> 0 Then label5_printer.Variables("qty3").SetValue(vqty3)
        If vqty4 <> 0 Then label5_printer.Variables("qty4").SetValue(vqty4)
        If vqty5 <> 0 Then label5_printer.Variables("qty5").SetValue(vqty5)
        If vqty6 <> 0 Then label5_printer.Variables("qty6").SetValue(vqty6)
        If vqty7 <> 0 Then label5_printer.Variables("qty7").SetValue(vqty7)
        If vqty8 <> 0 Then label5_printer.Variables("qty8").SetValue(vqty8)
        If vqty9 <> 0 Then label5_printer.Variables("qty9").SetValue(vqty9)
        If vqty10 <> 0 Then label5_printer.Variables("qty10").SetValue(vqty10)

        'label5_printer.Variables("qty1").SetValue(qty1.Text)
        'label5_printer.Variables("qty2").SetValue(qty2.Text)
        'label5_printer.Variables("qty3").SetValue(qty3.Text)
        'label5_printer.Variables("qty4").SetValue(qty4.Text)
        'label5_printer.Variables("qty5").SetValue(qty5.Text)
        'label5_printer.Variables("qty6").SetValue(qty6.Text)
        'label5_printer.Variables("qty7").SetValue(qty7.Text)
        'label5_printer.Variables("qty8").SetValue(qty8.Text)
        'label5_printer.Variables("qty9").SetValue(qty9.Text)
        'label5_printer.Variables("qty10").SetValue(qty10.Text)

        label5_printer.Variables("descr1").SetValue(Microsoft.VisualBasic.Left(descr1.Text, 50))
        label5_printer.Variables("descr2").SetValue(Microsoft.VisualBasic.Left(descr2.Text, 50))
        label5_printer.Variables("descr3").SetValue(Microsoft.VisualBasic.Left(descr3.Text, 50))
        label5_printer.Variables("descr4").SetValue(Microsoft.VisualBasic.Left(descr4.Text, 50))
        label5_printer.Variables("descr5").SetValue(Microsoft.VisualBasic.Left(descr5.Text, 50))
        label5_printer.Variables("descr6").SetValue(Microsoft.VisualBasic.Left(descr6.Text, 50))
        label5_printer.Variables("descr7").SetValue(Microsoft.VisualBasic.Left(descr7.Text, 50))
        label5_printer.Variables("descr8").SetValue(Microsoft.VisualBasic.Left(descr8.Text, 50))
        label5_printer.Variables("descr9").SetValue(Microsoft.VisualBasic.Left(descr9.Text, 50))
        label5_printer.Variables("descr10").SetValue(Microsoft.VisualBasic.Left(descr10.Text, 50))

    End Sub

    Private Sub Label6_SetValue()
        Try

            Dim vqty1 As Integer = 0
            Dim vqty2 As Integer = 0
            Dim vqty3 As Integer = 0
            Dim vqty4 As Integer = 0
            Dim vqty5 As Integer = 0
            Dim vqty6 As Integer = 0
            Dim vqty7 As Integer = 0
            Dim vqty8 As Integer = 0
            Dim vqty9 As Integer = 0
            Dim vqty10 As Integer = 0


            label6_printer.Variables("qty1").SetValue(" ")
            label6_printer.Variables("qty2").SetValue(" ")
            label6_printer.Variables("qty3").SetValue(" ")
            label6_printer.Variables("qty4").SetValue(" ")
            label6_printer.Variables("qty5").SetValue(" ")
            label6_printer.Variables("qty6").SetValue(" ")
            label6_printer.Variables("qty7").SetValue(" ")
            label6_printer.Variables("qty8").SetValue(" ")
            label6_printer.Variables("qty9").SetValue(" ")
            label6_printer.Variables("qty10").SetValue(" ")

            If qty1.Text <> "" And Not String.IsNullOrEmpty(qty1.Text) And Not String.IsNullOrWhiteSpace(qty1.Text) Then vqty1 = Convert.ToDecimal(qty1.Text) / Convert.ToDecimal(LabelQuantityitem.Text)
            If qty2.Text <> "" And Not String.IsNullOrEmpty(qty2.Text) And Not String.IsNullOrWhiteSpace(qty2.Text) Then vqty2 = Convert.ToDecimal(qty2.Text) / Convert.ToDecimal(LabelQuantityitem.Text)
            If qty3.Text <> "" And Not String.IsNullOrEmpty(qty3.Text) And Not String.IsNullOrWhiteSpace(qty3.Text) Then vqty3 = Convert.ToDecimal(qty3.Text) / Convert.ToDecimal(LabelQuantityitem.Text)
            If qty4.Text <> "" And Not String.IsNullOrEmpty(qty4.Text) And Not String.IsNullOrWhiteSpace(qty4.Text) Then vqty4 = Convert.ToDecimal(qty4.Text) / Convert.ToDecimal(LabelQuantityitem.Text)
            If qty5.Text <> "" And Not String.IsNullOrEmpty(qty5.Text) And Not String.IsNullOrWhiteSpace(qty5.Text) Then vqty5 = Convert.ToDecimal(qty5.Text) / Convert.ToDecimal(LabelQuantityitem.Text)
            If qty6.Text <> "" And Not String.IsNullOrEmpty(qty6.Text) And Not String.IsNullOrWhiteSpace(qty6.Text) Then vqty6 = Convert.ToDecimal(qty6.Text) / Convert.ToDecimal(LabelQuantityitem.Text)
            If qty7.Text <> "" And Not String.IsNullOrEmpty(qty7.Text) And Not String.IsNullOrWhiteSpace(qty7.Text) Then vqty7 = Convert.ToDecimal(qty7.Text) / Convert.ToDecimal(LabelQuantityitem.Text)
            If qty8.Text <> "" And Not String.IsNullOrEmpty(qty8.Text) And Not String.IsNullOrWhiteSpace(qty8.Text) Then vqty8 = Convert.ToDecimal(qty8.Text) / Convert.ToDecimal(LabelQuantityitem.Text)
            If qty9.Text <> "" And Not String.IsNullOrEmpty(qty9.Text) And Not String.IsNullOrWhiteSpace(qty9.Text) Then vqty9 = Convert.ToDecimal(qty9.Text) / Convert.ToDecimal(LabelQuantityitem.Text)
            If qty10.Text <> "" And Not String.IsNullOrEmpty(qty10.Text) And Not String.IsNullOrWhiteSpace(qty10.Text) Then vqty10 = Convert.ToDecimal(qty10.Text) / Convert.ToDecimal(LabelQuantityitem.Text)

            If vqty1 <> 0 Then label6_printer.Variables("qty1").SetValue(vqty1)
            If vqty2 <> 0 Then label6_printer.Variables("qty2").SetValue(vqty2)
            If vqty3 <> 0 Then label6_printer.Variables("qty3").SetValue(vqty3)
            If vqty4 <> 0 Then label6_printer.Variables("qty4").SetValue(vqty4)
            If vqty5 <> 0 Then label6_printer.Variables("qty5").SetValue(vqty5)
            If vqty6 <> 0 Then label6_printer.Variables("qty6").SetValue(vqty6)
            If vqty7 <> 0 Then label6_printer.Variables("qty7").SetValue(vqty7)
            If vqty8 <> 0 Then label6_printer.Variables("qty8").SetValue(vqty8)
            If vqty9 <> 0 Then label6_printer.Variables("qty9").SetValue(vqty9)
            If vqty10 <> 0 Then label6_printer.Variables("qty10").SetValue(vqty10)

            Try
                label6_printer.Variables("Product ID").SetValue(Microsoft.VisualBasic.Left(header.Text, 18))
            Catch ex As Exception
                MsgBox("PID - " & ex.Message)
            End Try
            Try
                label6_printer.Variables("range").SetValue(range.Text)
            Catch ex As Exception
                MsgBox("range - " & ex.Message)
            End Try
            Try
                label6_printer.Variables("PPNumberEntry").SetValue(PPnumberEntry.Text)
            Catch ex As Exception
                MsgBox("ppnumber - " & ex.Message)
            End Try
            Try
                label6_printer.Variables("datecode").SetValue("SG-" & dateCode.Text)
            Catch ex As Exception
                MsgBox("sg - " & ex.Message)
            End Try
            Try
                label6_printer.Variables("technicianShortName").SetValue("(" & technicianShortName.Text & ")")
            Catch ex As Exception
                MsgBox("shortname - " & ex.Message)
            End Try
            Try
                'label6_printer.Variables("madeInEnglish").SetValue(madeInEnglish.SelectedText)
                label6_printer.Variables("madeInEnglish").SetValue(madeInEnglish.Text)
            Catch ex As Exception
                MsgBox("madein - " & ex.Message)
            End Try

            label6_printer.Variables("mat1").SetValue(Microsoft.VisualBasic.Left(mat1.Text, 14))
            label6_printer.Variables("mat2").SetValue(Microsoft.VisualBasic.Left(mat2.Text, 14))
            label6_printer.Variables("mat3").SetValue(Microsoft.VisualBasic.Left(mat3.Text, 14))
            label6_printer.Variables("mat4").SetValue(Microsoft.VisualBasic.Left(mat4.Text, 14))
            label6_printer.Variables("mat5").SetValue(Microsoft.VisualBasic.Left(mat5.Text, 14))
            label6_printer.Variables("mat6").SetValue(Microsoft.VisualBasic.Left(mat6.Text, 14))
            label6_printer.Variables("mat7").SetValue(Microsoft.VisualBasic.Left(mat7.Text, 14))
            label6_printer.Variables("mat8").SetValue(Microsoft.VisualBasic.Left(mat8.Text, 14))
            label6_printer.Variables("mat9").SetValue(Microsoft.VisualBasic.Left(mat9.Text, 14))
            label6_printer.Variables("mat10").SetValue(Microsoft.VisualBasic.Left(mat10.Text, 14))

            label6_printer.Variables("descr1").SetValue(Microsoft.VisualBasic.Left(descr1.Text, 50))
            label6_printer.Variables("descr2").SetValue(Microsoft.VisualBasic.Left(descr2.Text, 50))
            label6_printer.Variables("descr3").SetValue(Microsoft.VisualBasic.Left(descr3.Text, 50))
            label6_printer.Variables("descr4").SetValue(Microsoft.VisualBasic.Left(descr4.Text, 50))
            label6_printer.Variables("descr5").SetValue(Microsoft.VisualBasic.Left(descr5.Text, 50))
            label6_printer.Variables("descr6").SetValue(Microsoft.VisualBasic.Left(descr6.Text, 50))
            label6_printer.Variables("descr7").SetValue(Microsoft.VisualBasic.Left(descr7.Text, 50))
            label6_printer.Variables("descr8").SetValue(Microsoft.VisualBasic.Left(descr8.Text, 50))
            label6_printer.Variables("descr9").SetValue(Microsoft.VisualBasic.Left(descr9.Text, 50))
            label6_printer.Variables("descr10").SetValue(Microsoft.VisualBasic.Left(descr10.Text, 50))

            If (PPnumberEntry.Text.Length >= 10 And Me.DataGridView1.Rows.Count <= 0) Or select_quantity.Checked = True Then
                label6_printer.Variables("label13").SetValue(Me.Quantity.Text)
            End If

        Catch ex As Exception
            MsgBox("Label 6 - " & ex.Message)
        End Try
    End Sub

    Private Sub progress_printing(ByVal a As Integer)
        If a > 0 And a < 100 Then
            Printing.Show()
        Else
            Printing.Close()
        End If

        Printing.ProgressBar1.Value = a
        Application.DoEvents()
    End Sub

    Private Sub Command191_Click(sender As Object, e As EventArgs) Handles Command191.Click
        'Report_Tab.SelectedIndex = 1
        'Report_Tab.SelectedIndex = 2

        progress_printing(30)

        'Set to Variable of NiceLabel
        'label3_printer.Variables("Name").SetValue("333333")
        label3_setvalue()
        'printing with quantity 
        'Dim qty As Integer = 1
        progress_printing(70)
        ' Dim a As Integer
        For a = Convert.ToDecimal(StartTestQuantity.Text) To Convert.ToDecimal(testQuantity.Text)
            'MsgBox(testPrintQty.SelectedItem)
            'printing with quantity 
            Dim qty As Integer = Convert.ToDecimal(COCprintQty.Text)
            Try
                label3_printer.Variables("StartTestQuantity").SetValue(a)
            Catch ex As Exception
                'Log_data("COC StartTestQTY : " & ex.ToString)
            End Try

            Try
                label3_printer.Print(COCprintQty.Text)
                progress_printing(90)
                cmd = New SqlCommand("insert into printingRecordCOC([pp],[date],[time],[user]) values(@pp,@date,@time,@user)", Main.koneksi)
                cmd.Parameters.AddWithValue("@pp", Me.PPnumberEntry.Text)
                cmd.Parameters.AddWithValue("@date", DateTime.Now.ToString("yyyy-MM-dd"))
                cmd.Parameters.AddWithValue("@time", DateTime.Now.ToString("HH:mm:ss"))
                cmd.Parameters.AddWithValue("@user", Me.technicianShortName.Text)
                cmd.ExecuteNonQuery()
            Catch ex As Exception
                'abc
                'MsgBox("Printing COC Cancel " & ex.Message)
            End Try
        Next
        progress_printing(100)
    End Sub

    Private Sub Command31_Click(sender As Object, e As EventArgs) Handles Command31.Click
        Me.Close()
    End Sub

    Private Sub PreviewTest_Click(sender As Object, e As EventArgs) Handles previewTest.Click
        reload_printer()
        'set to NiceLabel Variable
        'label2_printer.Variables("Name").SetValue("222222222222")

        'declaration of Preview
        Dim LabelPreviewSettings As ILabelPreviewSettings = New LabelPreviewSettings()

        'setting preview format
        LabelPreviewSettings.ImageFormat = "PNG"
        LabelPreviewSettings.Width = Form_preview.pictureBoxPreview.Width                   ' Width Of image To generate
        LabelPreviewSettings.Height = Form_preview.pictureBoxPreview.Height                 ' Height Of image To generate

        ' Generate Preview File
        Dim imageObj As Object = label2_printer.GetLabelPreview(LabelPreviewSettings)

        'Display image in UI
        If TypeOf imageObj Is Byte() Then
            Form_preview.pictureBoxPreview.Image = ByteToImage(CType(imageObj, Byte()))
        ElseIf TypeOf imageObj Is String Then
            Form_preview.pictureBoxPreview.ImageLocation = CStr(imageObj)
        End If

        Form_preview.Show()
    End Sub

    Private Sub PreviewCOC_Click(sender As Object, e As EventArgs) Handles previewCOC.Click
        reload_printer()

        'set to NiceLabel Variable
        label3_setvalue()

        'declaration of Preview
        Dim LabelPreviewSettings As ILabelPreviewSettings = New LabelPreviewSettings()

        'setting preview format
        LabelPreviewSettings.ImageFormat = "PNG"
        LabelPreviewSettings.Width = Form_preview.pictureBoxPreview.Width                   ' Width Of image To generate
        LabelPreviewSettings.Height = Form_preview.pictureBoxPreview.Height                 ' Height Of image To generate

        ' Generate Preview File
        Dim imageObj As Object = label3_printer.GetLabelPreview(LabelPreviewSettings)

        'Display image in UI
        If TypeOf imageObj Is Byte() Then
            Form_preview.pictureBoxPreview.Image = ByteToImage(CType(imageObj, Byte()))
        ElseIf TypeOf imageObj Is String Then
            Form_preview.pictureBoxPreview.ImageLocation = CStr(imageObj)
        End If

        Form_preview.Show()
    End Sub


    Private Sub PreviewLabel_Click(sender As Object, e As EventArgs) Handles previewLabel.Click
        Try
            reload_printer()
            'set to NiceLabel Variable
            'label2_printer.Variables("Name").SetValue("222222222222")
            If selectedLabel.Text.ToLower = "generic" Then Label_SetValue()
            If selectedLabel.Text.ToLower = "loose2" Then Label6_SetValue() 'fix tersangka masalah ada disini
            'declaration of Preview
            Dim LabelPreviewSettings As ILabelPreviewSettings = New LabelPreviewSettings()

            'setting preview format
            LabelPreviewSettings.ImageFormat = "PNG"
            LabelPreviewSettings.Width = Form_preview.pictureBoxPreview.Width                   ' Width Of image To generate
            LabelPreviewSettings.Height = Form_preview.pictureBoxPreview.Height                 ' Height Of image To generate
            Dim imageObj As Object
            If selectedLabel.Text.ToLower = "generic" Or selectedLabel.Text = "" Then
                ' generic
                imageObj = label_printer.GetLabelPreview(LabelPreviewSettings)
            Else
                ' loose2
                imageObj = label6_printer.GetLabelPreview(LabelPreviewSettings)
            End If


            'Display image in UI
            If TypeOf imageObj Is Byte() Then
                Form_preview.pictureBoxPreview.Image = ByteToImage(CType(imageObj, Byte()))
            ElseIf TypeOf imageObj Is String Then
                Form_preview.pictureBoxPreview.ImageLocation = CStr(imageObj)
            End If

            Form_preview.Show()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    'santo
    Private Sub Main_Closed(sender As Object, e As EventArgs) Handles Me.Closed
        LoginForm.Close()
        Form1.Close()
        Application.Exit()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim confirm = MessageBox.Show("Are You Sure For Logout?", "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If confirm = Windows.Forms.DialogResult.Yes Then
            Me.Hide()
            LoginForm.Show()
            LoginForm.textboxusername.Text = ""
            LoginForm.textboxpassword.Text = ""
        ElseIf confirm = Windows.Forms.DialogResult.Yes Then
            Exit Sub
        End If
        role = ""
    End Sub

    'Private Sub Upload_open_orders_Click(sender As Object, e As EventArgs) Handles upload_open_orders.Click
    '    Call koneksi_db()
    '    Dim cmd As New SqlCommand
    '    Dim existCmd As New SqlCommand
    '    Dim exist As New Integer
    '    Dim count As New Integer
    '    Try
    '        OpenFileDialog1.InitialDirectory = My.Computer.FileSystem.SpecialDirectories.MyDocuments
    '        OpenFileDialog1.Filter = "Microsoft Excel 97-2007(*.xls)|*.xls|Excel File(*.xlsx)|*.xlsx"
    '        'OpenFileDialog1.Filter = "Microsoft Excel 97-2007|*.xls"
    '        If OpenFileDialog1.ShowDialog(Me) = System.Windows.Forms.DialogResult.OK Then
    '            Dim xlApp As New Microsoft.Office.Interop.Excel.Application
    '            Dim xlWorkBook As Microsoft.Office.Interop.Excel.Workbook = xlApp.Workbooks.Open(OpenFileDialog1.FileName)
    '            Dim SheetName As String = xlWorkBook.Worksheets(1).Name.ToString
    '            Dim fi As New IO.FileInfo(OpenFileDialog1.FileName)
    '            Dim fileName As String = OpenFileDialog1.FileName
    '            excel = fi.FullName
    '            Dim asd As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & excel & ";Extended Properties='Excel 8.0;HDR=No;IMEX=1;'"
    '            oleCon = New OleDbConnection(asd)
    '            adapteroledb = New OleDbDataAdapter("select * from [" & SheetName & "$]", oleCon)
    '            dtoledb = New DataSet
    '            adapteroledb.Fill(dtoledb)
    '            Form1.open_form_import_pp(dtoledb)
    '            Form1.Show()

    '            existCmd = New SqlCommand("TRUNCATE TABLE openOrders", Main.koneksi)
    '            existCmd.ExecuteNonQuery()

    '            For i = 5 To dtoledb.Tables(0).Rows.Count - 1 ' pengulangan untuk upload
    '                If String.IsNullOrEmpty(dtoledb.Tables(0).Rows(i).Item(1).ToString()) Then
    '                    dtoledb.Tables(0).Rows(i).Item(1) = ""
    '                End If
    '                If String.IsNullOrEmpty(dtoledb.Tables(0).Rows(i).Item(2).ToString()) Then
    '                    dtoledb.Tables(0).Rows(i).Item(2) = ""
    '                End If
    '                If String.IsNullOrEmpty(dtoledb.Tables(0).Rows(i).Item(3).ToString()) Then
    '                    dtoledb.Tables(0).Rows(i).Item(3) = ""
    '                End If
    '                If String.IsNullOrEmpty(dtoledb.Tables(0).Rows(i).Item(4).ToString()) Then
    '                    dtoledb.Tables(0).Rows(i).Item(4) = ""
    '                End If
    '                If String.IsNullOrEmpty(dtoledb.Tables(0).Rows(i).Item(5).ToString()) Then
    '                    dtoledb.Tables(0).Rows(i).Item(5) = ""
    '                End If
    '                If String.IsNullOrEmpty(dtoledb.Tables(0).Rows(i).Item(6).ToString()) Then
    '                    dtoledb.Tables(0).Rows(i).Item(6) = ""
    '                End If
    '                If String.IsNullOrEmpty(dtoledb.Tables(0).Rows(i).Item(7).ToString()) Then
    '                    dtoledb.Tables(0).Rows(i).Item(7) = ""
    '                End If

    '                'memasukkan ke db
    '                cmd = New SqlCommand("insert into openOrders([order],[material],[descr],[reqmts qty],[stor  loc],[status],[peggedreqt]) values(@order,@material,@descr,@qty,@str,@status,@pegged)", Main.koneksi)
    '                cmd.Parameters.AddWithValue("@order", dtoledb.Tables(0).Rows(i).Item(1).ToString())
    '                cmd.Parameters.AddWithValue("@material", dtoledb.Tables(0).Rows(i).Item(2).ToString())
    '                cmd.Parameters.AddWithValue("@descr", dtoledb.Tables(0).Rows(i).Item(3).ToString())
    '                cmd.Parameters.AddWithValue("@qty", dtoledb.Tables(0).Rows(i).Item(4).ToString())
    '                cmd.Parameters.AddWithValue("@str", dtoledb.Tables(0).Rows(i).Item(5).ToString())
    '                cmd.Parameters.AddWithValue("@status", dtoledb.Tables(0).Rows(i).Item(6).ToString())
    '                cmd.Parameters.AddWithValue("@pegged", dtoledb.Tables(0).Rows(i).Item(7).ToString())
    '                cmd.ExecuteNonQuery()
    '                'If count = 100 Or count = 1000 Or count = 2000 Or count = 3000 Or count = 4000 Or count = 5000 Or count = 6000 Or count = 7000 Then
    '                '    MsgBox("still on progress, " + count.ToString + " records")
    '                'End If
    '                count = count + 1
    '            Next
    '            MsgBox("Success Upload COOIS " + count.ToString + " records")
    '        End If
    '    Catch ex As Exception
    '        MessageBox.Show(ex.Message)
    '    End Try
    'End Sub

    'Private Sub Command121_Click(sender As Object, e As EventArgs) Handles Command121.Click
    '    Call koneksi_db()
    '    Dim cmd As New SqlCommand
    '    Dim existCmd As New SqlCommand
    '    Dim exist As New Integer
    '    Dim count As New Integer
    '    Try
    '        OpenFileDialog1.InitialDirectory = My.Computer.FileSystem.SpecialDirectories.MyDocuments
    '        OpenFileDialog1.Filter = "Microsoft Excel 97-2007(*.xls)|*.xls|Excel File(*.xlsx)|*.xlsx"
    '        'OpenFileDialog1.Filter = "Microsoft Excel 97-2007|*.xls"
    '        If OpenFileDialog1.ShowDialog(Me) = System.Windows.Forms.DialogResult.OK Then
    '            Dim xlApp As New Microsoft.Office.Interop.Excel.Application
    '            Dim xlWorkBook As Microsoft.Office.Interop.Excel.Workbook = xlApp.Workbooks.Open(OpenFileDialog1.FileName)
    '            Dim SheetName As String = xlWorkBook.Worksheets(1).Name.ToString
    '            Dim fi As New IO.FileInfo(OpenFileDialog1.FileName)
    '            Dim fileName As String = OpenFileDialog1.FileName
    '            excel = fi.FullName
    '            Dim asd As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & excel & ";Extended Properties='Excel 8.0;HDR=No;IMEX=1;'"
    '            oleCon = New OleDbConnection(asd)
    '            adapteroledb = New OleDbDataAdapter("select * from [" & SheetName & "$]", oleCon)
    '            dtoledb = New DataSet
    '            adapteroledb.Fill(dtoledb)
    '            Form1.open_form_import_pplist(dtoledb)
    '            Form1.Show()

    '            existCmd = New SqlCommand("TRUNCATE TABLE PPList", Main.koneksi)
    '            existCmd.ExecuteNonQuery()

    '            For i = 1 To dtoledb.Tables(0).Rows.Count - 1 ' pengulangan untuk upload

    '                'memasukkan ke db
    '                cmd = New SqlCommand("insert into [pplist] ([created on],[entered by],[type],[order],[start time],[basic fin],[committed],[schedstart],[finish tme],[prctr],[material],[scheduled finish] ,[scheduled start],[basic start date],[basic finish date],[item quantity],[oum],[material description],[description],[cdi],[req dlv dt],[purchase order number],[poitem],[gi time],[stag  time],[gi date],[mat av dt],[confirmed qty],[su],[customer],[city],[name 2],[name 1],[item],[so no],[customer material number],[status],[stat]) values(@create,@enter,@type,@order,@start,@basic,@com,@sched,@finish,@prctr,@material,@sched_finish,@sched_start,@basic_finish,@basic_start,@item_qty,@oum,@material_desc,@desc,@cdi,@req,@purchase,@po,@gi_ti,@stag,@gi_da,@mat,@confirm,@su,@customer,@city,@name2,@name1,@item,@so,@cusmat,@status,@stat)", Main.koneksi)
    '                cmd.Parameters.AddWithValue("@create", Convert.ToDateTime(New DateTime().FromOADate(dtoledb.Tables(0).Rows(i).Item(0).ToString())))
    '                cmd.Parameters.AddWithValue("@enter", dtoledb.Tables(0).Rows(i).Item(1).ToString())
    '                cmd.Parameters.AddWithValue("@type", dtoledb.Tables(0).Rows(i).Item(2).ToString())
    '                cmd.Parameters.AddWithValue("@order", dtoledb.Tables(0).Rows(i).Item(3).ToString())
    '                cmd.Parameters.AddWithValue("@start", Convert.ToDateTime(New DateTime().FromOADate(dtoledb.Tables(0).Rows(i).Item(4).ToString())))
    '                cmd.Parameters.AddWithValue("@basic", Convert.ToDateTime(New DateTime().FromOADate(dtoledb.Tables(0).Rows(i).Item(5).ToString())))
    '                cmd.Parameters.AddWithValue("@com", dtoledb.Tables(0).Rows(i).Item(6).ToString())
    '                cmd.Parameters.AddWithValue("@sched", Convert.ToDateTime(New DateTime().FromOADate(dtoledb.Tables(0).Rows(i).Item(7).ToString())))
    '                cmd.Parameters.AddWithValue("@finish", Convert.ToDateTime(New DateTime().FromOADate(dtoledb.Tables(0).Rows(i).Item(8).ToString())))
    '                cmd.Parameters.AddWithValue("@prctr", dtoledb.Tables(0).Rows(i).Item(9))
    '                cmd.Parameters.AddWithValue("@material", dtoledb.Tables(0).Rows(i).Item(10).ToString())
    '                cmd.Parameters.AddWithValue("@sched_finish", Convert.ToDateTime(New DateTime().FromOADate(dtoledb.Tables(0).Rows(i).Item(11).ToString())))
    '                cmd.Parameters.AddWithValue("@sched_start", Convert.ToDateTime(New DateTime().FromOADate(dtoledb.Tables(0).Rows(i).Item(12).ToString())))
    '                cmd.Parameters.AddWithValue("@basic_start", Convert.ToDateTime(New DateTime().FromOADate(dtoledb.Tables(0).Rows(i).Item(13).ToString())))
    '                cmd.Parameters.AddWithValue("@basic_finish", Convert.ToDateTime(New DateTime().FromOADate(dtoledb.Tables(0).Rows(i).Item(14).ToString())))
    '                cmd.Parameters.AddWithValue("@item_qty", dtoledb.Tables(0).Rows(i).Item(15))
    '                cmd.Parameters.AddWithValue("@oum", dtoledb.Tables(0).Rows(i).Item(16).ToString())
    '                cmd.Parameters.AddWithValue("@material_desc", dtoledb.Tables(0).Rows(i).Item(17).ToString())
    '                cmd.Parameters.AddWithValue("@desc", dtoledb.Tables(0).Rows(i).Item(18).ToString())
    '                cmd.Parameters.AddWithValue("@cdi", dtoledb.Tables(0).Rows(i).Item(19).ToString())
    '                If String.IsNullOrEmpty(dtoledb.Tables(0).Rows(i).Item(20).ToString()) Then
    '                    cmd.Parameters.AddWithValue("@req", "")
    '                Else
    '                    cmd.Parameters.AddWithValue("@req", Convert.ToDateTime(New DateTime().FromOADate(dtoledb.Tables(0).Rows(i).Item(20).ToString())))
    '                End If
    '                cmd.Parameters.AddWithValue("@purchase", dtoledb.Tables(0).Rows(i).Item(21).ToString())
    '                cmd.Parameters.AddWithValue("@po", dtoledb.Tables(0).Rows(i).Item(22).ToString())
    '                cmd.Parameters.AddWithValue("@gi_ti", Convert.ToDateTime(New DateTime().FromOADate(dtoledb.Tables(0).Rows(i).Item(23).ToString())))
    '                cmd.Parameters.AddWithValue("@stag", Convert.ToDateTime(New DateTime().FromOADate(dtoledb.Tables(0).Rows(i).Item(24).ToString())))
    '                cmd.Parameters.AddWithValue("@gi_da", Convert.ToDateTime(New DateTime().FromOADate(dtoledb.Tables(0).Rows(i).Item(25).ToString())))
    '                cmd.Parameters.AddWithValue("@mat", Convert.ToDateTime(New DateTime().FromOADate(dtoledb.Tables(0).Rows(i).Item(26).ToString())))
    '                cmd.Parameters.AddWithValue("@confirm", dtoledb.Tables(0).Rows(i).Item(27))
    '                cmd.Parameters.AddWithValue("@su", dtoledb.Tables(0).Rows(i).Item(28).ToString())
    '                cmd.Parameters.AddWithValue("@customer", dtoledb.Tables(0).Rows(i).Item(29).ToString())
    '                cmd.Parameters.AddWithValue("@city", dtoledb.Tables(0).Rows(i).Item(30).ToString())
    '                cmd.Parameters.AddWithValue("@name2", dtoledb.Tables(0).Rows(i).Item(31).ToString())
    '                cmd.Parameters.AddWithValue("@name1", dtoledb.Tables(0).Rows(i).Item(32).ToString())
    '                cmd.Parameters.AddWithValue("@item", dtoledb.Tables(0).Rows(i).Item(33))
    '                cmd.Parameters.AddWithValue("@so", dtoledb.Tables(0).Rows(i).Item(34))
    '                cmd.Parameters.AddWithValue("@cusmat", dtoledb.Tables(0).Rows(i).Item(35).ToString())
    '                cmd.Parameters.AddWithValue("@status", dtoledb.Tables(0).Rows(i).Item(36).ToString())
    '                cmd.Parameters.AddWithValue("@stat", dtoledb.Tables(0).Rows(i).Item(37).ToString())
    '                cmd.ExecuteNonQuery()
    '                'if count = 100 or count = 1000 or count = 2000 or count = 3000 or count = 4000 or count = 5000 or count = 6000 or count = 7000 then
    '                '    msgbox("still on progress, " + count.tostring + " records")
    '                'end if
    '                count = count + 1

    '            Next
    '            MsgBox("Success Upload SQOO " + count.ToString + " records")
    '        End If
    '    Catch ex As Exception
    '        'MessageBox.Show(ex.Message)
    '    End Try
    'End Sub

    Private Sub Command155_Click(sender As Object, e As EventArgs) Handles Command155.Click
        'Form2.Show()
        OpenFileDialog1.InitialDirectory = My.Computer.FileSystem.SpecialDirectories.MyDocuments
        If OpenFileDialog1.ShowDialog(Me) = System.Windows.Forms.DialogResult.OK Then
            Dim xlApp As New Microsoft.Office.Interop.Excel.Application
            Dim xlWorkBook As Microsoft.Office.Interop.Excel.Workbook = xlApp.Workbooks.Open(OpenFileDialog1.FileName)
            Dim SheetName As String = xlWorkBook.Worksheets(1).Name.ToString
            Dim excelpath As String = OpenFileDialog1.FileName
            Dim koneksiExcel As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & excelpath & ";Extended Properties='Excel 8.0;HDR=No;IMEX=1;'"
            oleCon = New OleDbConnection(koneksiExcel)
            oleCon.Open()

            Dim queryExcel As String = "select * from [" & SheetName & "$]"
            Dim cmd As OleDbCommand = New OleDbCommand(queryExcel, oleCon)
            Dim rd As OleDbDataReader

            Call koneksi_db()

            Using bulkCopy As SqlBulkCopy = New SqlBulkCopy(Main.koneksi)
                bulkCopy.DestinationTableName = "dbo.PPList"
                Try
                    rd = cmd.ExecuteReader
                    bulkCopy.ColumnMappings.Add(0, 1)
                    bulkCopy.ColumnMappings.Add(1, 2)
                    bulkCopy.ColumnMappings.Add(2, 3)
                    bulkCopy.ColumnMappings.Add(3, 4)
                    bulkCopy.ColumnMappings.Add(4, 5)
                    bulkCopy.ColumnMappings.Add(5, 6)
                    bulkCopy.ColumnMappings.Add(6, 7)
                    bulkCopy.ColumnMappings.Add(7, 8)
                    bulkCopy.ColumnMappings.Add(8, 9)
                    bulkCopy.ColumnMappings.Add(9, 10)
                    bulkCopy.ColumnMappings.Add(10, 11)
                    bulkCopy.ColumnMappings.Add(11, 12)
                    bulkCopy.ColumnMappings.Add(12, 13)
                    bulkCopy.ColumnMappings.Add(13, 14)
                    bulkCopy.ColumnMappings.Add(14, 15)
                    bulkCopy.ColumnMappings.Add(15, 16)
                    bulkCopy.ColumnMappings.Add(16, 17)
                    bulkCopy.ColumnMappings.Add(17, 18)
                    bulkCopy.ColumnMappings.Add(18, 19)
                    bulkCopy.ColumnMappings.Add(19, 20)
                    bulkCopy.ColumnMappings.Add(20, 21)
                    bulkCopy.ColumnMappings.Add(21, 22)
                    bulkCopy.ColumnMappings.Add(22, 23)
                    bulkCopy.ColumnMappings.Add(23, 24)
                    bulkCopy.ColumnMappings.Add(24, 25)
                    bulkCopy.ColumnMappings.Add(25, 26)
                    bulkCopy.ColumnMappings.Add(26, 27)
                    bulkCopy.ColumnMappings.Add(27, 28)
                    bulkCopy.ColumnMappings.Add(28, 29)
                    bulkCopy.ColumnMappings.Add(29, 30)
                    bulkCopy.ColumnMappings.Add(30, 31)
                    bulkCopy.ColumnMappings.Add(31, 32)
                    bulkCopy.ColumnMappings.Add(32, 33)
                    bulkCopy.ColumnMappings.Add(33, 34)
                    bulkCopy.ColumnMappings.Add(34, 35)
                    bulkCopy.ColumnMappings.Add(35, 36)
                    bulkCopy.ColumnMappings.Add(36, 37)
                    bulkCopy.ColumnMappings.Add(37, 38)

                    bulkCopy.WriteToServer(rd)

                    Dim Sql = "INSERT INTO [Upload_History] (upload) Values ('SQOO');"
                    Dim insert = New SqlCommand(Sql, Main.koneksi)
                    insert.ExecuteNonQuery()

                    rd.Close()
                    MsgBox("Add Daily SQOO Success !")
                Catch ex As Exception
                    MsgBox("Add SQOO Fail" & ex.Message)

                End Try
            End Using
        End If

        'Delete Duplicate
        DeleteDup_Click()
    End Sub

    Private Sub CounterCpts_TextChanged(sender As Object, e As EventArgs) Handles CounterCpts.TextChanged

        If CInt(CounterCpts.Text) < 0 Then CounterCpts.Text = 0
        If CounterCpts.Text = LabelQuantitycpt.Text And autoPrint.Checked = True Then
            'If CounterCpts.Text = LabelQuantitycpt.Text Then
            'Dim sql As String = "select DISTINCT PeggedReqt from openOrders, MasterFuji where openOrders.[order]='" & Me.PPnumberEntry.Text & "' and MasterFuji.Ref=openOrders.PeggedReqt"
            'Dim ds As New DataSet
            'adapter = New SqlDataAdapter(sql, Main.koneksi)
            'adapter.Fill(ds)
            'If ds.Tables(0).Rows.Count > 0 Then
            '    MessageBox.Show("Printing Label Fuji")
            '    Exit Sub
            'End If

            'End If

            'MsgBox("Printing .... ")
            ' Product_Label_print()
            'Packaging_label_print()

            '    If chosenLabel = "loose2" And Me.labelQty.Text > 1 Then

            '        Dim PrintCount As Integer = 0

            '        For i = 1 To Me.labelQty.Text - 1

            '            'Reports(Trim(chosenLabel))![Label13].Caption = Me.boxQty
            '            'print
            '            'DoCmd.SelectObject acReport, Trim(chosenLabel), True
            '            'DoCmd.PrintOut , , , , 1

            '            PrintCount = PrintCount + CInt(Me.boxQty.Text)

            '        Next

            '        'Reports(Trim(chosenLabel))![Label13].Caption = Me.Quantity - PrintCount
            '        'print
            '        'DoCmd.SelectObject acReport, Trim(chosenLabel), True
            '        'DoCmd.PrintOut , , , , 1

            '    Else
            '        Dim traceability = "SG" & yearNum & thisweek & dayNum & serial
            '        'Reports(Trim(chosenLabel))![Label13].Caption = Me.Quantity

            '        If chosenLabel = "loose" Then
            '            'Reports(Trim(chosenLabel))![serialNum].Caption = traceability
            '            'Reports(Trim(chosenLabel))![ItemNumber].Caption = Me.Quantity
            '        End If

            '        '        DoCmd.SelectObject acReport, Trim(chosenLabel), True
            '        'DoCmd.PrintOut , , , , 1

            '    End If

            'End If

            'asdf

            'Else 
            'Printing Label
            'If CheckProductLabelPrinting.Checked = True Then
            AutoPrintProductLabelFrom.Text = Convert.ToDecimal(AutoPrintProductLabelFrom.Text) + 1
            If selectedLabel.Text.ToLower = "generic" Then
                seq_SerialNumber = checkSerialNumber()
                Label_SetValue() 'generic
            End If
            'Label4_SetValue() 'Loose2
            If selectedLabel.Text.ToLower = "loose2" Then Label6_SetValue()
            Try
                label_printer.Variables("CounterItems").SetValue(AutoPrintProductLabelFrom.Text)
            Catch ex As Exception

            End Try

            Try
                label_printer.Variables("DataMatrix").SetValue("SG" & dateCode3() & Convert.ToDecimal(AutoPrintProductLabelFrom.Text).ToString("0000"))
            Catch ex As Exception

            End Try

            'label_printer.Variables("CounterItems").SetValue(CounterItems.Text)
            'Try


            'If selectedLabel.Text.ToLower <> "loose2" Then
            If selectedLabel.Text.ToLower = "generic" Then
                'seq_SerialNumber = checkSerialNumber()
                label_printer.Print(1)
            ElseIf selectedLabel.Text.ToLower = "loose2" Then
                label6_printer.Variables("Label13").SetValue(1)
                label6_printer.Print(1)
            End If

            'Catch ex As Exception
            ' MsgBox("Printing Product Label Failed !" & Chr(13) & " Pls do reprint or check printer connection")
            'End Try
            'End If

            'Save to DB every printing
            'Try
            '    Dim cmd = New SqlCommand("insert into printingRecord([pp],[date],[time],[user],[from],[to],[QRCodeFuji],[Data],[Seq]) values(@pp,@date,@time,@user,@from,@to,@QRCodeFuji,@data,@Seq)", Main.koneksi)
            '    cmd.Parameters.AddWithValue("@pp", Me.PPnumberEntry.Text)
            '    cmd.Parameters.AddWithValue("@date", DateTime.Now.ToString("yyyy-MM-dd"))
            '    cmd.Parameters.AddWithValue("@time", DateTime.Now.ToString("HH:mm:ss"))
            '    cmd.Parameters.AddWithValue("@user", Me.technicianShortName.Text)
            '    cmd.Parameters.AddWithValue("@from", CInt(Me.CounterItems.Text) + 1)
            '    cmd.Parameters.AddWithValue("@to", Me.LabelQuantityitem.Text)
            '    cmd.Parameters.AddWithValue("@QRCodeFuji", Me.Fuji_QR_Product_Label.Text)
            '    cmd.Parameters.AddWithValue("@data", LoginForm.strHostName & " - " & Application.ProductVersion)

            '    Dim date_2 As String

            '    Dim VBAWeekNum As Integer = DatePart(DateInterval.WeekOfYear, Date.Today, FirstDayOfWeek.Monday, FirstWeekOfYear.FirstFourDays)

            '    If Len(VBAWeekNum) = 1 Then VBAWeekNum = "0" & VBAWeekNum

            '    date_2 = Date.Now.ToString("yy") & VBAWeekNum & DateAndTime.Weekday(DateTime.Now, vbMonday)

            '    Dim _ID As String = Convert.ToDecimal(seq_SerialNumber.ToString).ToString("0000")
            '    Dim _seq As String = "SG" & date_2 & _ID

            '    cmd.Parameters.AddWithValue("@Seq", _seq)

            '    'cmd.Parameters.AddWithValue("@Seq", seq_SerialNumber.ToString)
            '    cmd.ExecuteNonQuery()


            '    'cmd = New SqlCommand("insert into printingRecord([pp],[date],[time],[user],[from],[to],[Data],[Seq]) values(@pp,@date,@time,@user,@from,@to,@data,@Seq)", Main.koneksi)
            '    'cmd.Parameters.AddWithValue("@pp", Me.PPnumberEntry.Text)
            '    'cmd.Parameters.AddWithValue("@date", DateTime.Now.ToString("yyyy-MM-dd"))
            '    'cmd.Parameters.AddWithValue("@time", DateTime.Now.ToString("HH:mm:ss"))
            '    'cmd.Parameters.AddWithValue("@user", Me.technicianShortName.Text)

            '    'If var_PrintLabel_klik = True Then
            '    '    cmd.Parameters.AddWithValue("@from", CInt(Me.StartLabel.Text))
            '    '    var_PrintLabel_klik = False
            '    'Else
            '    '    cmd.Parameters.AddWithValue("@from", CInt(Me.CounterItems.Text) + 1)
            '    'End If


            '    'cmd.Parameters.AddWithValue("@to", Me.LabelQuantityitem.Text)
            '    'cmd.Parameters.AddWithValue("@data", LoginForm.strHostName & " - " & Application.ProductVersion)
            '    'cmd.Parameters.AddWithValue("@Seq", seq_SerialNumber.ToString)
            '    'cmd.ExecuteNonQuery()

            'Catch ex As Exception

            'End Try
            'If CounterItems.Text > 0 Then



            'End If
            'end if



            'If checkPackagingPrinting.Checked = True Then
            '    AutoPrintPackagingLabelFrom.Text = Convert.ToDecimal(AutoPrintPackagingLabelFrom.Text) + 1
            '    label1_setValue()
            '    label1_printer.Variables("Package").SetValue(AutoPrintProductLabelFrom.Text)
            '    Try
            '        label1_printer.Print(1)
            '    Catch ex As Exception
            '        MsgBox("Printing Packaging Label Failed !")
            '    End Try
            'End If

            'CounterItems.Text = CInt(CounterItems.Text) + 1
            'CounterCpts.Text = 0
            'LabelCheckCom.Text = CInt(LabelCheckCom.Text) + 1
            'DataGridView1.BackgroundColor = Color.White
            'For i = 0 To DataGridView1.Rows.Count - 2
            '    DataGridView1.Rows(i).Cells(4).Style.BackColor = Color.White
            'Next
        End If
        Try
            Dim cmd = New SqlCommand("insert into printingRecord([pp],[date],[time],[user],[from],[to],[QRCodeFuji],[Data],[Seq]) values(@pp,@date,@time,@user,@from,@to,@QRCodeFuji,@data,@Seq)", Main.koneksi)
            cmd.Parameters.AddWithValue("@pp", Me.PPnumberEntry.Text)
            cmd.Parameters.AddWithValue("@date", DateTime.Now.ToString("yyyy-MM-dd"))
            cmd.Parameters.AddWithValue("@time", DateTime.Now.ToString("HH:mm:ss"))
            cmd.Parameters.AddWithValue("@user", Me.technicianShortName.Text)
            cmd.Parameters.AddWithValue("@from", CInt(Me.CounterItems.Text) + 1)
            cmd.Parameters.AddWithValue("@to", Me.LabelQuantityitem.Text)
            cmd.Parameters.AddWithValue("@QRCodeFuji", Me.Fuji_QR_Product_Label.Text)
            cmd.Parameters.AddWithValue("@data", LoginForm.strHostName & " - " & Application.ProductVersion)

            Dim date_2 As String

            Dim VBAWeekNum As Integer = DatePart(DateInterval.WeekOfYear, Date.Today, FirstDayOfWeek.Monday, FirstWeekOfYear.FirstFourDays)

            If Len(VBAWeekNum) = 1 Then VBAWeekNum = "0" & VBAWeekNum

            date_2 = Date.Now.ToString("yy") & VBAWeekNum & DateAndTime.Weekday(DateTime.Now, vbMonday)

            Dim _ID As String = Convert.ToDecimal(seq_SerialNumber.ToString).ToString("0000")
            Dim _seq As String = "SG" & date_2 & _ID

            cmd.Parameters.AddWithValue("@Seq", _seq)

            'cmd.Parameters.AddWithValue("@Seq", seq_SerialNumber.ToString)
            cmd.ExecuteNonQuery()


            'cmd = New SqlCommand("insert into printingRecord([pp],[date],[time],[user],[from],[to],[Data],[Seq]) values(@pp,@date,@time,@user,@from,@to,@data,@Seq)", Main.koneksi)
            'cmd.Parameters.AddWithValue("@pp", Me.PPnumberEntry.Text)
            'cmd.Parameters.AddWithValue("@date", DateTime.Now.ToString("yyyy-MM-dd"))
            'cmd.Parameters.AddWithValue("@time", DateTime.Now.ToString("HH:mm:ss"))
            'cmd.Parameters.AddWithValue("@user", Me.technicianShortName.Text)

            'If var_PrintLabel_klik = True Then
            '    cmd.Parameters.AddWithValue("@from", CInt(Me.StartLabel.Text))
            '    var_PrintLabel_klik = False
            'Else
            '    cmd.Parameters.AddWithValue("@from", CInt(Me.CounterItems.Text) + 1)
            'End If


            'cmd.Parameters.AddWithValue("@to", Me.LabelQuantityitem.Text)
            'cmd.Parameters.AddWithValue("@data", LoginForm.strHostName & " - " & Application.ProductVersion)
            'cmd.Parameters.AddWithValue("@Seq", seq_SerialNumber.ToString)
            'cmd.ExecuteNonQuery()

        Catch ex As Exception

        End Try
    End Sub

    Private Sub Command111_Click(sender As Object, e As EventArgs) Handles Command111.Click
        OpenFileDialog1.InitialDirectory = My.Computer.FileSystem.SpecialDirectories.MyDocuments
        If OpenFileDialog1.ShowDialog(Me) = System.Windows.Forms.DialogResult.OK Then
            Dim xlApp As New Microsoft.Office.Interop.Excel.Application
            Dim xlWorkBook As Microsoft.Office.Interop.Excel.Workbook = xlApp.Workbooks.Open(OpenFileDialog1.FileName)
            Dim SheetName As String = xlWorkBook.Worksheets(1).Name.ToString
            Dim excelpath As String = OpenFileDialog1.FileName
            Dim koneksiExcel As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & excelpath & ";Extended Properties='Excel 8.0;HDR=No;IMEX=1;'"
            oleCon = New OleDbConnection(koneksiExcel)
            oleCon.Open()

            Dim queryExcel As String = "select * from [" & SheetName & "$]"
            Dim cmd As OleDbCommand = New OleDbCommand(queryExcel, oleCon)
            Dim rd As OleDbDataReader

            Call koneksi_db()

            Dim deleteReset As New SqlCommand("Delete from [dbo].[Users] DBCC CHECKIDENT ('[Users]',RESEED,0)", Main.koneksi)

            Try
                deleteReset.ExecuteNonQuery()
                'Show Msg Box
                'Timer1.Enabled = True
                'MsgBox("database Delete & Reset ID !")
            Catch ex As Exception
                MsgBox("Delete DB Fail " & ex.Message)
            End Try

            Using bulkCopy As SqlBulkCopy = New SqlBulkCopy(Main.koneksi)
                bulkCopy.DestinationTableName = "dbo.Users"
                Try
                    rd = cmd.ExecuteReader
                    bulkCopy.ColumnMappings.Add(0, 1)
                    bulkCopy.ColumnMappings.Add(1, 2)
                    bulkCopy.WriteToServer(rd)
                    rd.Close()
                    MsgBox("Upload Technician Success")
                Catch ex As Exception
                    MsgBox(ex.Message)
                End Try
            End Using
        End If
    End Sub

    Private Sub Command121_Click(sender As Object, e As EventArgs) Handles Command121.Click
        OpenFileDialog1.InitialDirectory = My.Computer.FileSystem.SpecialDirectories.MyDocuments
        If OpenFileDialog1.ShowDialog(Me) = System.Windows.Forms.DialogResult.OK Then
            Dim xlApp As New Microsoft.Office.Interop.Excel.Application
            Dim xlWorkBook As Microsoft.Office.Interop.Excel.Workbook = xlApp.Workbooks.Open(OpenFileDialog1.FileName)
            Dim SheetName As String = xlWorkBook.Worksheets(1).Name.ToString
            Dim excelpath As String = OpenFileDialog1.FileName
            Dim koneksiExcel As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & excelpath & ";Extended Properties='Excel 8.0;HDR=No;IMEX=1;'"
            oleCon = New OleDbConnection(koneksiExcel)
            oleCon.Open()

            Dim queryExcel As String = "select * from [" & SheetName & "$]"
            Dim cmd As OleDbCommand = New OleDbCommand(queryExcel, oleCon)
            Dim rd As OleDbDataReader

            Call koneksi_db()

            Dim deleteReset As New SqlCommand("Delete from [dbo].[PPList] DBCC CHECKIDENT ('[PPList]',RESEED,0)", Main.koneksi)
            Try
                deleteReset.ExecuteNonQuery()
                'Show Msg Box
                'Timer1.Enabled = True
                'MsgBox("database Deleted & Reset ID !")
            Catch ex As Exception
                MsgBox("Delete DB Fail " & ex.Message)
            End Try

            Using bulkCopy As SqlBulkCopy = New SqlBulkCopy(Main.koneksi)
                bulkCopy.DestinationTableName = "dbo.PPList"
                Try
                    rd = cmd.ExecuteReader
                    bulkCopy.ColumnMappings.Add(0, 1)
                    bulkCopy.ColumnMappings.Add(1, 2)
                    bulkCopy.ColumnMappings.Add(2, 3)
                    bulkCopy.ColumnMappings.Add(3, 4)
                    bulkCopy.ColumnMappings.Add(4, 5)
                    bulkCopy.ColumnMappings.Add(5, 6)
                    bulkCopy.ColumnMappings.Add(6, 7)
                    bulkCopy.ColumnMappings.Add(7, 8)
                    bulkCopy.ColumnMappings.Add(8, 9)
                    bulkCopy.ColumnMappings.Add(9, 10)
                    bulkCopy.ColumnMappings.Add(10, 11)
                    bulkCopy.ColumnMappings.Add(11, 12)
                    bulkCopy.ColumnMappings.Add(12, 13)
                    bulkCopy.ColumnMappings.Add(13, 14)
                    bulkCopy.ColumnMappings.Add(14, 15)
                    bulkCopy.ColumnMappings.Add(15, 16)
                    bulkCopy.ColumnMappings.Add(16, 17)
                    bulkCopy.ColumnMappings.Add(17, 18)
                    bulkCopy.ColumnMappings.Add(18, 19)
                    bulkCopy.ColumnMappings.Add(19, 20)
                    bulkCopy.ColumnMappings.Add(20, 21)
                    bulkCopy.ColumnMappings.Add(21, 22)
                    bulkCopy.ColumnMappings.Add(22, 23)
                    bulkCopy.ColumnMappings.Add(23, 24)
                    bulkCopy.ColumnMappings.Add(24, 25)
                    bulkCopy.ColumnMappings.Add(25, 26)
                    bulkCopy.ColumnMappings.Add(26, 27)
                    bulkCopy.ColumnMappings.Add(27, 28)
                    bulkCopy.ColumnMappings.Add(28, 29)
                    bulkCopy.ColumnMappings.Add(29, 30)
                    bulkCopy.ColumnMappings.Add(30, 31)
                    bulkCopy.ColumnMappings.Add(31, 32)
                    bulkCopy.ColumnMappings.Add(32, 33)
                    bulkCopy.ColumnMappings.Add(33, 34)
                    bulkCopy.ColumnMappings.Add(34, 35)
                    bulkCopy.ColumnMappings.Add(35, 36)
                    bulkCopy.ColumnMappings.Add(36, 37)
                    bulkCopy.ColumnMappings.Add(37, 38)

                    bulkCopy.WriteToServer(rd)

                    Dim Sql = "INSERT INTO [Upload_History] (upload) Values ('SQOO');"
                    Dim insert = New SqlCommand(Sql, Main.koneksi)
                    insert.ExecuteNonQuery()

                    rd.Close()
                    MsgBox("Upload SQOO Success")
                Catch ex As Exception
                    MsgBox("Upload SQOO Fail " & ex.Message)
                End Try
            End Using

            btn_delete_dot_sq00_Click(sender, e)
        End If
    End Sub

    Private Sub Upload_open_orders_Click(sender As Object, e As EventArgs) Handles upload_open_orders.Click
        OpenFileDialog1.InitialDirectory = My.Computer.FileSystem.SpecialDirectories.MyDocuments
        If OpenFileDialog1.ShowDialog(Me) = System.Windows.Forms.DialogResult.OK Then
            Dim xlApp As New Microsoft.Office.Interop.Excel.Application
            Dim xlWorkBook As Microsoft.Office.Interop.Excel.Workbook = xlApp.Workbooks.Open(OpenFileDialog1.FileName)
            Dim SheetName As String = xlWorkBook.Worksheets(1).Name.ToString
            Dim excelpath As String = OpenFileDialog1.FileName
            Dim koneksiExcel As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & excelpath & ";Extended Properties='Excel 8.0;HDR=No;IMEX=1;'"
            oleCon = New OleDbConnection(koneksiExcel)
            oleCon.Open()

            Dim queryExcel As String = "select * from [" & SheetName & "$]"
            Dim cmd As OleDbCommand = New OleDbCommand(queryExcel, oleCon)
            Dim rd As OleDbDataReader

            Call koneksi_db()

            Dim deleteReset As New SqlCommand("Delete from [dbo].[openOrders] DBCC CHECKIDENT ('[openOrders]',RESEED,0)", Main.koneksi)

            Try
                deleteReset.ExecuteNonQuery()
                'Show Msg Box
                'Timer1.Enabled = True
                'MsgBox("database Delete & Reset ID !")
            Catch ex As Exception
                MsgBox("Delete DB Fail " & ex.Message)
            End Try

            Using bulkCopy As SqlBulkCopy = New SqlBulkCopy(Main.koneksi)
                bulkCopy.DestinationTableName = "dbo.openOrders"
                bulkCopy.BulkCopyTimeout = 120
                Try
                    rd = cmd.ExecuteReader
                    bulkCopy.ColumnMappings.Add(0, 1)
                    bulkCopy.ColumnMappings.Add(1, 2)
                    bulkCopy.ColumnMappings.Add(2, 3)
                    bulkCopy.ColumnMappings.Add(3, 4)
                    bulkCopy.ColumnMappings.Add(4, 5)
                    bulkCopy.ColumnMappings.Add(5, 6)
                    bulkCopy.ColumnMappings.Add(6, 7)

                    bulkCopy.WriteToServer(rd)

                    Dim Sql = "INSERT INTO [Upload_History] (upload) Values ('COOIS');"
                    Dim insert = New SqlCommand(Sql, Main.koneksi)
                    insert.ExecuteNonQuery()

                    rd.Close()
                    MsgBox("Add COOIS Successed !")
                Catch ex As Exception
                    MsgBox("Add COOIS Fail " & ex.Message)
                End Try

                Try
                    Dim Sql2 = "UPDATE openOrders set [Reqmts qty] = REPLACE([Reqmts qty],'.','') WHERE [Reqmts qty] LIKE '%.%';"
                    Dim insert2 = New SqlCommand(Sql2, Main.koneksi)
                    'insert2.ExecuteNonQuery()

                    Dim afected As Integer = insert2.ExecuteNonQuery()
                    If afected > 0 Then MsgBox("Deleting '.' Succeed, Rows Affected:" & afected.ToString)
                Catch ex As Exception
                    MsgBox("Deleting '.' Failed !")
                End Try
            End Using
        End If
    End Sub
    'Remove last space
    Private Sub Picture2_TextChanged(sender As Object, e As EventArgs) Handles Picture2.TextChanged
        Try
            If Picture2.Text(Picture2.Text.Length - 1) = " " Then Picture2.Text = Picture2.Text.Remove(Picture2.Text.Length - 1)
            'If Picture2.Text = "CVS C2 4P" Then Picture2.Text = "CVS C2 3P"   ' ini harus diperbaiki
        Catch ex As Exception
            'MsgBox(ex.Message)
        End Try

    End Sub

    Private Sub AutoPrint_CheckedChanged(sender As Object, e As EventArgs) Handles autoPrint.CheckedChanged
        autoPrint2.Checked = autoPrint.Checked
    End Sub

    Private Sub AutoPrint2_CheckedChanged(sender As Object, e As EventArgs) Handles autoPrint2.CheckedChanged
        autoPrint.Checked = autoPrint2.Checked
    End Sub

    Private Sub CheckPackagingPrinting2_CheckedChanged(sender As Object, e As EventArgs) Handles checkPackagingPrinting2.CheckedChanged
        checkPackagingPrinting.Checked = checkPackagingPrinting2.Checked
    End Sub

    Private Sub CheckPackagingPrinting_CheckedChanged(sender As Object, e As EventArgs) Handles checkPackagingPrinting.CheckedChanged
        checkPackagingPrinting2.Checked = checkPackagingPrinting.Checked
        'boxQty_TextChanged()
    End Sub

    Private Sub CheckProductLabelPrinting_CheckedChanged(sender As Object, e As EventArgs) Handles CheckProductLabelPrinting.CheckedChanged
        CheckProductLabelPrinting2.Checked = CheckProductLabelPrinting.Checked
    End Sub

    Private Sub CheckProductLabelPrinting2_CheckedChanged(sender As Object, e As EventArgs) Handles CheckProductLabelPrinting2.CheckedChanged
        CheckProductLabelPrinting.Checked = CheckProductLabelPrinting2.Checked
    End Sub

    Sub AddComponent()

        Dim sql As String
        Dim Material As String

        On Error GoTo Addcomponent_Error

        Dim Mat = InputBox("Please enter the Material number :", "Add Component into the Components List")
        Dim bar = InputBox("Please scan the barcode, if necessary:", "Add Component into the Components List")

        If Len(bar) > 6 Then
            If Len(bar) = 10 Then
                bar = Microsoft.VisualBasic.Right(Me.ComponentNo.Text, 5)
            ElseIf Len(bar) >= 11 Then
                bar = Microsoft.VisualBasic.Left(Microsoft.VisualBasic.Right(bar, 6), 5)
            End If
        ElseIf Len(bar) < 5 Then
            MsgBox("The barcode is not correct")
            Exit Sub
        End If

        If Mat = "" Or String.IsNullOrEmpty(Mat) Then
            MsgBox("Please enter a correct Material number")
            Exit Sub
        Else
            Dim dsexist = New DataSet
            Dim exist = "select * from [Componentslist] where [material]='" & Mat & "'"
            Dim adapterexist = New SqlDataAdapter(exist, koneksi)
            adapterexist.Fill(dsexist)
            If dsexist.Tables(0).Rows.Count > 0 Then
                Dim deleteExist = "delete from [Componentslist] where [material] = '" & Mat & "'"
                Dim cmddelete = New SqlCommand(deleteExist, Main.koneksi)
                cmddelete.ExecuteNonQuery()
                sql = "INSERT INTO [Componentslist] (Material,code) Values ('" & Mat & "', '" & bar & "');"
            Else
                sql = "INSERT INTO [Componentslist] (Material,code) Values ('" & Mat & "', '" & bar & "');"
            End If
            Dim cmd = New SqlCommand(sql, Main.koneksi)
            cmd.ExecuteNonQuery()

            MsgBox("Component: " & Mat & " with barcode number : " & bar & " added")

            Exit Sub

        End If
Addcomponent_Error:
        Exit Sub
    End Sub

    Sub DeleteComponent()

        Dim sql As String
        Dim Mat As String

        On Error GoTo Deletecomponent_Error

        Mat = InputBox("Please enter the Material number to delete:", "Delete Component from the Components List")

        If Mat = "" Or String.IsNullOrEmpty(Mat) Then
            MsgBox("Please enter a correct Material number")
            Exit Sub
        Else

            sql = "Delete FROM [Componentslist] WHERE Material='" & Mat & "';"
            Dim cmd = New SqlCommand(sql, Main.koneksi)
            cmd.ExecuteNonQuery()

            MsgBox("Component: " & Mat & " deleted")

            Exit Sub

        End If

Deletecomponent_Error:
        Exit Sub

    End Sub

    Sub UpdateComponent()

        Dim sql As String
        Dim Material As String
        Dim barcode As String
        Dim code As String
        Dim ds As New DataSet


        Material = InputBox("Please key the Material number you want to update", "Material Number")

        If Material = "" Or String.IsNullOrEmpty(Material) Then
            MsgBox("Please enter a correct Material number")
            Exit Sub
        End If

        sql = "SELECT * FROM [ComponentsList] WHERE Material= '" & Material & "';"
        Dim adapter = New SqlDataAdapter(sql, Main.koneksi)
        adapter.Fill(ds)


        If ds.Tables(0).Rows.Count > 0 Then

            barcode = InputBox("Please key or scan the barcode", "New Barcode")

            If Len(barcode) >= 10 Then
                If Len(barcode) = 10 Then
                    code = Microsoft.VisualBasic.Right(Me.ComponentNo.Text, 5)
                ElseIf Len(barcode) >= 11 Then
                    code = Microsoft.VisualBasic.Left(Microsoft.VisualBasic.Right(barcode, 6), 5)
                End If
            ElseIf Len(barcode) < 10 Then
                MsgBox("The barcode is too short")
                Exit Sub
            End If

            sql = "INSERT INTO ComponentUpdateRecord (workstation, day, Material, formercode,BarcodeScanned, newcode) " &
    "VALUES ('" & Me.workstation.Text & "', '" & DateTime.Now & "', '" & Material & "', '" & ds.Tables(0).Rows(0).Item("code").ToString & "', '" & barcode & "', '" & code & "');"
            Dim cmd = New SqlCommand(sql, Main.koneksi)
            cmd.ExecuteNonQuery()

            sql = "UPDATE ComponentsList SET code='" & code & "' WHERE Material ='" & Material & "' ;"
            Dim cmd1 = New SqlCommand(sql, Main.koneksi)
            cmd.ExecuteNonQuery()

            sql = "UPDATE Components SET code='" & code & "' WHERE Material ='" & Material & "' ;"
            Dim cmd2 = New SqlCommand(sql, Main.koneksi)
            cmd.ExecuteNonQuery()


        Else
            Dim Answer = MsgBox(Material & " not found, Do you want to add it?", vbQuestion + vbYesNo + vbDefaultButton2, "Material not found")

            If Answer = vbYes Then
                AddComponent()
            End If

        End If

    End Sub


    Private Sub Command576_Click(sender As Object, e As EventArgs) Handles Command576.Click
        AddComponent()
    End Sub

    Private Sub Command575_Click(sender As Object, e As EventArgs) Handles Command575.Click
        DeleteComponent()
    End Sub

    Private Sub Command577_Click(sender As Object, e As EventArgs) Handles Command577.Click
        UpdateComponent()
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        SendKeys.Send("{ENTER}")
        Timer1.Enabled = False
    End Sub
    Private Sub TabPage3_Enter(sender As Object, e As EventArgs) Handles TabPage3.Enter
        Cek_OF_Click()
    End Sub
    Private Sub TabPage6_Enter(sender As Object, e As EventArgs) Handles TabPage6.Enter
        DGV_Quality()
        DGV_MasterPlantCode()
    End Sub
    Private Sub TabPage7_Enter(sender As Object, e As EventArgs) Handles TabPage7.Enter
        DGV_QualityIssue_refresh()
    End Sub
    Private Sub TabPage4_Enter(sender As Object, e As EventArgs) Handles TabPage4.Enter
        Dim Sql = "SELECT TOP 1 * FROM [Upload_History] where upload='COOIS' order by datetime DESC"
        Dim adapter As New SqlDataAdapter(Sql, Main.koneksi)
        Dim ds As New DataSet
        adapter.Fill(ds)
        latestUpdate.Text = ds.Tables(0).Rows(0).Item("datetime").ToString

        Dim Sql1 = "SELECT TOP 1 * FROM [Upload_History] where upload='SQOO' order by datetime DESC"
        Dim adapter1 As New SqlDataAdapter(Sql1, Main.koneksi)
        Dim ds1 As New DataSet
        adapter1.Fill(ds1)
        latestUpdatePacking.Text = ds1.Tables(0).Rows(0).Item("datetime").ToString
    End Sub
    Private Sub TabPage1_Enter(sender As Object, e As EventArgs) Handles TabPage1.Enter
        'If cust_TextBox1.Text <> "" Then customer.Text = cust_TextBox1.Text
    End Sub
    Private Sub TabPage2_Enter(sender As Object, e As EventArgs) Handles TabPage2.Enter
        If cust_TextBox1.Text <> "" Then customer.Text = cust_TextBox1.Text
        Timer2.Enabled = True
        'boxQty_TextChanged()
    End Sub



    Private Sub Command293_Click(sender As Object, e As EventArgs) Handles Command293.Click
        'selected_Printer = printers.Item(listprinter2.SelectedIndex)
        label2_printer.PrintSettings.PrinterName = "Microsoft Print to PDF"
        Label2_setValue()

        ' Dim a As Integer
        For a = 1 To Convert.ToDecimal(testPrintQty.SelectedItem)
            'MsgBox(testPrintQty.SelectedItem)
            'printing with quantity 
            Dim qty As Integer = 1
            label2_printer.Variables("Pint_seq").SetValue(a)
            Try
                label2_printer.Print(qty)
            Catch ex As Exception
                MsgBox("Printing Cancel " & ex.Message)
            End Try
        Next

        'listprinter.SelectedItem = selected_Printer.Name
        'MsgBox(selected_Printer.Name)
    End Sub

    Private Sub Command292_Click(sender As Object, e As EventArgs) Handles Command292.Click
        label3_printer.PrintSettings.PrinterName = "Microsoft Print to PDF"
        label3_setvalue()

        ' Dim a As Integer
        'For a = 1 To Convert.ToDecimal(COCprintQty.SelectedItem)
        'MsgBox(testPrintQty.SelectedItem)
        'printing with quantity 
        ' Dim qty As Integer = 1
        'label2_printer.Variables("Pint_seq").SetValue(a)
        Try
            label3_printer.Print(COCprintQty.SelectedItem)
        Catch ex As Exception
            MsgBox("Printing Cancel " & ex.Message)
        End Try
        'Next
    End Sub

    Private Sub Command110_Click(sender As Object, e As EventArgs) Handles Command110.Click
        ExportToExcel("SELECT * FROM NSXMasterdata")
    End Sub

    Private Sub ExportToExcel(sql_command As String)
        Dim Dialog As New SaveFileDialog

        Dim datatableMain As New System.Data.DataTable()
        'Dim sql As String 

        Dim f As FolderBrowserDialog = New FolderBrowserDialog
        Dialog.Filter = "Microsoft Excel 97-2003|*.xls"
        If (Dialog.ShowDialog = DialogResult.OK) Then
            'Try
            Dim oExcel As Excel.Application
            Dim oBook As Excel.Workbook
            Dim oSheet As Excel.Worksheet

            oExcel = CreateObject("Excel.Application")
            oBook = oExcel.Workbooks.Add(Type.Missing)
            oSheet = oBook.Worksheets(1)

            Dim dc As System.Data.DataColumn
            Dim dr As System.Data.DataRow
            Dim colIndex As Integer = 0
            Dim rowIndex As Integer = 0

            Dim adapter As New SqlCommand(sql_command, Main.koneksi)
            Dim dataAdapter As New SqlClient.SqlDataAdapter()

            dataAdapter.SelectCommand = adapter
            Try
                'adapter.ExecuteNonQuery()
                dataAdapter.Fill(datatableMain)
            Catch ex As Exception
                MsgBox("Export Fail " & ex.Message)
                Exit Sub
            End Try

            Dim a As Integer = 0
            Dim b As Integer = 0
            'Dim c As Integer
            'c = datatableMain.Columns.Count + datatableMain.Rows.Count

            'Set final path
            Dim fileName As String = Dialog.FileName '+ ".xls"
            Dim finalPath = f.SelectedPath + fileName


            'Export the Columns to excel file
            For Each dc In datatableMain.Columns
                colIndex = colIndex + 1
                oSheet.Cells(1, colIndex) = dc.ColumnName
            Next
            'show progress
            Progress.Show()


            'Export the rows to excel file
            For Each dr In datatableMain.Rows
                rowIndex = rowIndex + 1
                colIndex = 0
                'data pregress bar
                Progress.ProgressBar1.Value = 100 * rowIndex / datatableMain.Rows.Count
                Progress.Label1.Text = "Pregress: " & Progress.ProgressBar1.Value & " %"
                For Each dc In datatableMain.Columns
                    colIndex = colIndex + 1
                    oSheet.Cells(rowIndex + 1, colIndex) = dr(dc.ColumnName)

                Next
            Next

            oSheet.Columns.AutoFit()
            'Save file in final path
            oBook.SaveAs(finalPath)

            'Release the objects
            releaseObject(oSheet)
            oBook.Close(False, Type.Missing, Type.Missing)
            releaseObject(oBook)
            oExcel.Quit()
            releaseObject(oExcel)

            'Some time Office application does not quit after automation: 
            'so i am calling GC.Collect method.
            GC.Collect()

            MessageBox.Show("Export done successfully!")
            Progress.Close()

            'Catch ex As Exception
            'MsgBox("Export to Excel Failed: " & ex.Message)
            Progress.Close()
            'End Try

        End If
    End Sub
    Private Sub releaseObject(ByVal obj As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
        Catch ex As Exception
        Finally
            obj = Nothing
        End Try
    End Sub

    Private Sub Command38_Click(sender As Object, e As EventArgs) Handles Command38.Click
        ExportToExcel("SELECT * FROM printingRecord")
    End Sub

    Private Sub Command153_Click(sender As Object, e As EventArgs) Handles Command153.Click
        ExportToExcel("SELECT * FROM PPList")
    End Sub

    Private Sub Command148_Click(sender As Object, e As EventArgs) Handles Command148.Click
        ExportToExcel("SELECT * FROM customerDatabase")
    End Sub

    Private Sub Command142_Click(sender As Object, e As EventArgs) Handles Command142.Click
        ExportToExcel("SELECT * FROM printingRecordPacking")
    End Sub

    Private Sub Command112_Click(sender As Object, e As EventArgs) Handles Command112.Click
        ExportToExcel("SELECT * FROM Users")
    End Sub

    Private Sub Command114_Click(sender As Object, e As EventArgs) Handles Command114.Click
        ExportToExcel("SELECT * FROM labelSelectionTable")
    End Sub

    Private Sub Command205_Click(sender As Object, e As EventArgs) Handles Command205.Click
        ExportToExcel("SELECT * FROM printingRecordCOC")
    End Sub

    Private Sub Command207_Click(sender As Object, e As EventArgs) Handles Command207.Click
        ExportToExcel("SELECT * FROM printingRecordTest")
    End Sub

    Private Sub Command280_Click(sender As Object, e As EventArgs) Handles Command280.Click
        ExportToExcel("SELECT * FROM BOM")
    End Sub

    Private Sub Export_open_orders_Click(sender As Object, e As EventArgs) Handles export_open_orders.Click
        'ExportToExcel("SELECT * FROM openOrders")
        'MsgBox("Data is too huge Use SSMs")
        'Export_to_Excel.Show()
        'Form1.open_form("select [ID]
        Main.open_form("select [ID]
        ,[Order]
        ,[Material]
        ,[Descr]
        ,[Reqmts qty]
        ,[Stor  loc]
        ,[Status]
        ,[PeggedReqt] from openOrders")
        'Form1.Show()
    End Sub

    'santo  Export to Excel Open Order
    Public Shared Sub open_form(query)
        Form1.publicQuery = query
        Call Main.koneksi_db()
        Try
            Dim sc As New SqlCommand(query, Main.koneksi)
            Dim adapter As New SqlDataAdapter(sc)
            Dim ds As New DataSet

            adapter.Fill(ds)
            Form1.DataGridView1.DataSource = ds.Tables(0)
            Form1.DataGridView1.Rows(0).Selected = True

            adapter.UpdateCommand = New SqlCommandBuilder(adapter).GetUpdateCommand
            adapter.Update(ds)



            'Creating DataTable.
            Dim dt_x As New System.Data.DataTable()
            dt_x = Form1.DataGridView1.DataSource

            ''Adding the Columns.
            'For Each column As DataGridViewColumn In Form1.DataGridView1.Columns
            '    dt_x.Columns.Add(column.HeaderText, column.ValueType)
            'Next

            ''Adding the Rows.
            'For Each row As DataGridViewRow In Form1.DataGridView1.Rows
            '    dt_x.Rows.Add()
            '    For Each cell As DataGridViewCell In row.Cells
            '        dt_x.Rows(dt_x.Rows.Count - 1)(cell.ColumnIndex) = cell.Value.ToString()
            '    Next
            'Next

            Dim objDlg As New SaveFileDialog
            objDlg.Filter = "Excel File|*.xls"
            objDlg.OverwritePrompt = False
            If objDlg.ShowDialog = DialogResult.OK Then
                Dim filepath As String = objDlg.FileName
                Main.ExportToExcel_ww(dt_x, filepath)
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub


    Private Sub ExportToExcel_ww(ByVal dtTemp As DataTable, ByVal filepath As String)
        Dim strFileName As String = filepath
        If System.IO.File.Exists(strFileName) Then
            If (MessageBox.Show("Do you want to replace from the existing file?", "Export to Excel", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) = System.Windows.Forms.DialogResult.Yes) Then
                System.IO.File.Delete(strFileName)
            Else
                Return
            End If

        End If
        Dim _excel As New Excel.Application
        Dim wBook As Excel.Workbook
        Dim wSheet As Excel.Worksheet

        wBook = _excel.Workbooks.Add()
        wSheet = wBook.ActiveSheet()

        Dim dt As System.Data.DataTable = dtTemp
        Dim dc As System.Data.DataColumn
        Dim dr As System.Data.DataRow
        Dim colIndex As Integer = 0
        Dim rowIndex As Integer = 0
        'If CheckBox1.Checked Then
        '    For Each dc In dt.Columns
        '        colIndex = colIndex + 1
        '        wSheet.Cells(1, colIndex) = dc.ColumnName
        '    Next
        'End If
        Dim progress_data As Integer
        Progress.Show()

        For Each dr In dt.Rows
            rowIndex = rowIndex + 1
            colIndex = 0

            progress_data = 100 * rowIndex / dt.Rows.Count
            Progress.Label1.Text = "Progress: " & progress_data & " %"
            Progress.ProgressBar1.Value = progress_data

            For Each dc In dt.Columns
                colIndex = colIndex + 1
                wSheet.Cells(rowIndex + 1, colIndex) = dr(dc.ColumnName)
            Next
        Next
        wSheet.Columns.AutoFit()
        wBook.SaveAs(strFileName)

        ReleaseObject_ww(wSheet)
        wBook.Close(False)
        ReleaseObject_ww(wBook)
        _excel.Quit()
        ReleaseObject_ww(_excel)
        GC.Collect()

        MessageBox.Show("File Export Successfully!")
        Progress.Close()
    End Sub
    Private Sub ReleaseObject_ww(ByVal o As Object)
        Try
            While (System.Runtime.InteropServices.Marshal.ReleaseComObject(o) > 0)
            End While
        Catch
        Finally
            o = Nothing
        End Try
    End Sub

    Private Sub Label154_Click(sender As Object, e As EventArgs)
        MsgBox("MES@2019")
    End Sub


    Private Sub Command323_Click(sender As Object, e As EventArgs) Handles Command323.Click
        OpenFileDialog1.InitialDirectory = My.Computer.FileSystem.SpecialDirectories.MyDocuments
        If OpenFileDialog1.ShowDialog(Me) = System.Windows.Forms.DialogResult.OK Then
            Dim xlApp As New Microsoft.Office.Interop.Excel.Application
            Dim xlWorkBook As Microsoft.Office.Interop.Excel.Workbook = xlApp.Workbooks.Open(OpenFileDialog1.FileName)
            Dim SheetName As String = xlWorkBook.Worksheets(1).Name.ToString
            Dim excelpath As String = OpenFileDialog1.FileName
            Dim koneksiExcel As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & excelpath & ";Extended Properties='Excel 8.0;HDR=No;IMEX=1;'"
            oleCon = New OleDbConnection(koneksiExcel)
            oleCon.Open()

            Dim queryExcel As String = "select * from [" & SheetName & "$]"
            Dim cmd As OleDbCommand = New OleDbCommand(queryExcel, oleCon)
            Dim rd As OleDbDataReader

            Call koneksi_db()

            Using bulkCopy As SqlBulkCopy = New SqlBulkCopy(Main.koneksi)
                bulkCopy.DestinationTableName = "dbo.openOrders"
                bulkCopy.BulkCopyTimeout = 120
                Try
                    rd = cmd.ExecuteReader
                    bulkCopy.ColumnMappings.Add(0, 1)
                    bulkCopy.ColumnMappings.Add(1, 2)
                    bulkCopy.ColumnMappings.Add(2, 3)
                    bulkCopy.ColumnMappings.Add(3, 4)
                    bulkCopy.ColumnMappings.Add(4, 5)
                    bulkCopy.ColumnMappings.Add(5, 6)
                    bulkCopy.ColumnMappings.Add(6, 7)

                    bulkCopy.WriteToServer(rd)

                    Dim Sql = "INSERT INTO [Upload_History] (upload) Values ('COOIS');"
                    Dim insert = New SqlCommand(Sql, Main.koneksi)
                    insert.ExecuteNonQuery()

                    rd.Close()
                    MsgBox("Upload New COOIS Successed !")
                Catch ex As Exception
                    MsgBox("Upload COOIS Fail " & ex.Message)
                End Try
            End Using
        End If
    End Sub

    Private Sub AutoPrintProductLabelFrom_TextChanged(sender As Object, e As EventArgs) Handles AutoPrintProductLabelFrom.TextChanged
        If Convert.ToDecimal(AutoPrintPackagingLabelFrom.Text) > 1 Then
            quantityLabel.Text = Convert.ToDecimal(AutoPrintPackagingLabelFrom.Text - 1)
        ElseIf Convert.ToDecimal(AutoPrintPackagingLabelFrom.Text) = 1 Then
            quantityLabel.Text = 1
        End If
    End Sub

    Private Sub CounterItems_TextChanged(sender As Object, e As EventArgs)
        quantityLabel.Text = Convert.ToDecimal(CounterItems.Text)
        'MsgBox("berubah")
    End Sub

    Private Sub DeleteDup_Click() Handles deleteDup.Click
        Dim sql As String = "with deleteDup as(select *, ROW_NUMBER() over (partition by [order] order by id) as RowNumber from [SGRAC_MES].[dbo].[PPList]) delete from deleteDup where RowNumber > 1"
        Dim cmd As New SqlCommand(sql, Main.koneksi)
        Dim count As Integer = cmd.ExecuteNonQuery()
        MsgBox("The Number Of Duplicate data has been executed : " & count & " Records")
    End Sub

    'Private Sub Customer_SelectedIndexChanged(sender As Object, e As EventArgs) Handles customer.SelectedIndexChanged
    '    Dim ds As New DataSet
    '    Dim dt As New DataTable
    '    Dim adapter As New SqlDataAdapter("Select * FROM customerDatabase WHERE [customer name] = '" & Me.customer.Text & "'", Main.koneksi)
    '    adapter.Fill(ds)

    '    If ds.Tables(0).Rows.Count = 1 Then
    '        Customer_code.Text = ds.Tables(0).Rows(0).Item("customer code").ToString
    '    End If

    'End Sub

    Private Sub StartPackingLabel2_TextChanged(sender As Object, e As EventArgs) Handles StartPackingLabel2.TextChanged
        StartPackingLabel.Text = StartPackingLabel2.Text
    End Sub

    Private Sub LabelQty2_TextChanged(sender As Object, e As EventArgs) Handles labelQty2.TextChanged
        labelQty.Text = labelQty2.Text
    End Sub

    Private Sub StartPackingLabel_TextChanged(sender As Object, e As EventArgs) Handles StartPackingLabel.TextChanged
        StartPackingLabel2.Text = StartPackingLabel.Text
    End Sub

    Private Sub LabelQty_TextChanged(sender As Object, e As EventArgs) Handles labelQty.TextChanged
        labelQty2.Text = labelQty.Text
    End Sub

    Private Sub PalletQty_TextChanged() Handles palletQty.TextChanged
        convert_box_pallet()
    End Sub

    Private Sub Quantity_TextChanged(sender As Object, e As EventArgs) Handles Quantity.TextChanged
        PalletQty_TextChanged()
        quantityLabel.Text = Quantity.Text
    End Sub

    Private Sub Listprinter2_SelectedIndexChanged() Handles listprinter2.SelectedIndexChanged
        'Selection of printer
        If printers.Count > 0 Then
            selected_Printer = printers.Item(listprinter2.SelectedIndex)
            label2_printer.PrintSettings.PrinterName = selected_Printer.Name
            listprinter2.SelectedItem = selected_Printer.Name
        End If
        'MsgBox("Test Report")
    End Sub

    Private Sub Listprinter3_SelectedIndexChanged() Handles listprinter3.SelectedIndexChanged
        'Selection of printer
        If printers.Count > 0 Then
            selected_Printer = printers.Item(listprinter3.SelectedIndex)
            label3_printer.PrintSettings.PrinterName = selected_Printer.Name
            listprinter3.SelectedItem = selected_Printer.Name
        End If
        'MsgBox("COC")
    End Sub

    Private Sub TestCountry_SelectedIndexChanged(sender As Object, e As EventArgs) Handles testCountry.SelectedIndexChanged
        Dim dt As New DataTable
        Dim sql As String = "select * from [customerDatabase] where [country] ='" & testCountry.Text & "'"
        Dim adapter As New SqlDataAdapter(sql, Main.koneksi)
        adapter.Fill(dt)

        Me.customer.DisplayMember = "customer name"
        Me.customer.ValueMember = "customer name"
        Me.customer.DataSource = dt
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        'If cust_TextBox1.Text <> "" Then customer.Text = cust_TextBox1.Text
    End Sub

    Private Sub Timer2_Tick(sender As Object, e As EventArgs) Handles Timer2.Tick
        If cust_TextBox1.Text <> "" Then customer.Text = cust_TextBox1.Text
        If cust_TextBox1.Text <> "" Then testCustomer.Text = cust_TextBox1.Text
        Timer2.Enabled = False
    End Sub

    Private Sub Command109_Click(sender As Object, e As EventArgs) Handles Command109.Click
        OpenFileDialog1.InitialDirectory = My.Computer.FileSystem.SpecialDirectories.MyDocuments
        If OpenFileDialog1.ShowDialog(Me) = System.Windows.Forms.DialogResult.OK Then
            Dim xlApp As New Microsoft.Office.Interop.Excel.Application
            Dim xlWorkBook As Microsoft.Office.Interop.Excel.Workbook = xlApp.Workbooks.Open(OpenFileDialog1.FileName)
            Dim SheetName As String = xlWorkBook.Worksheets(1).Name.ToString
            Dim excelpath As String = OpenFileDialog1.FileName
            Dim koneksiExcel As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & excelpath & ";Extended Properties='Excel 8.0;HDR=No;IMEX=1;'"
            oleCon = New OleDbConnection(koneksiExcel)
            oleCon.Open()

            Dim queryExcel As String = "select * from [" & SheetName & "$]"
            Dim cmd As OleDbCommand = New OleDbCommand(queryExcel, oleCon)
            Dim rd As OleDbDataReader

            Call koneksi_db()

            Dim deleteReset As New SqlCommand("Delete from [dbo].[NSXMasterdata] DBCC CHECKIDENT ('[NSXMasterdata]',RESEED,0)", Main.koneksi)

            Try
                deleteReset.ExecuteNonQuery()
                'Show Msg Box
                'Timer1.Enabled = True
                'MsgBox("database Delete & Reset ID !")
            Catch ex As Exception
                MsgBox("Delete DB Fail " & ex.Message)
            End Try

            Using bulkCopy As SqlBulkCopy = New SqlBulkCopy(Main.koneksi)
                bulkCopy.DestinationTableName = "dbo.NSXMasterdata"
                bulkCopy.BulkCopyTimeout = 300 '5 menit
                Try
                    rd = cmd.ExecuteReader
                    bulkCopy.ColumnMappings.Add(0, 1)
                    bulkCopy.ColumnMappings.Add(1, 2)
                    bulkCopy.ColumnMappings.Add(2, 3)
                    bulkCopy.ColumnMappings.Add(3, 4)
                    bulkCopy.ColumnMappings.Add(4, 5)
                    bulkCopy.ColumnMappings.Add(5, 6)
                    bulkCopy.ColumnMappings.Add(6, 7)
                    bulkCopy.ColumnMappings.Add(7, 8)
                    bulkCopy.ColumnMappings.Add(8, 9)
                    bulkCopy.ColumnMappings.Add(9, 10)
                    bulkCopy.ColumnMappings.Add(10, 11)
                    bulkCopy.ColumnMappings.Add(11, 12)
                    bulkCopy.ColumnMappings.Add(12, 13)
                    bulkCopy.ColumnMappings.Add(13, 14)
                    bulkCopy.ColumnMappings.Add(14, 15)
                    bulkCopy.ColumnMappings.Add(15, 16)
                    bulkCopy.WriteToServer(rd)
                    rd.Close()
                    MsgBox("Upload NSX Master Data Success")
                Catch ex As Exception
                    MsgBox(ex.Message)
                End Try
            End Using
        End If
    End Sub

    Private Sub Command147_Click(sender As Object, e As EventArgs) Handles Command147.Click
        OpenFileDialog1.InitialDirectory = My.Computer.FileSystem.SpecialDirectories.MyDocuments
        If OpenFileDialog1.ShowDialog(Me) = System.Windows.Forms.DialogResult.OK Then
            Dim xlApp As New Microsoft.Office.Interop.Excel.Application
            Dim xlWorkBook As Microsoft.Office.Interop.Excel.Workbook = xlApp.Workbooks.Open(OpenFileDialog1.FileName)
            Dim SheetName As String = xlWorkBook.Worksheets(1).Name.ToString
            Dim excelpath As String = OpenFileDialog1.FileName
            Dim koneksiExcel As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & excelpath & ";Extended Properties='Excel 8.0;HDR=No;IMEX=1;'"
            oleCon = New OleDbConnection(koneksiExcel)
            oleCon.Open()

            Dim queryExcel As String = "select * from [" & SheetName & "$]"
            Dim cmd As OleDbCommand = New OleDbCommand(queryExcel, oleCon)
            Dim rd As OleDbDataReader

            Call koneksi_db()

            Dim deleteReset As New SqlCommand("Delete from [dbo].[customerDatabase] DBCC CHECKIDENT ('[customerDatabase]',RESEED,0)", Main.koneksi)

            Try
                deleteReset.ExecuteNonQuery()
                'Show Msg Box
                'Timer1.Enabled = True
                'MsgBox("database Delete & Reset ID !")
            Catch ex As Exception
                MsgBox("Delete DB Fail " & ex.Message)
            End Try

            Using bulkCopy As SqlBulkCopy = New SqlBulkCopy(Main.koneksi)
                bulkCopy.DestinationTableName = "dbo.customerDatabase"
                bulkCopy.BulkCopyTimeout = 300 '5 menit
                Try
                    rd = cmd.ExecuteReader
                    bulkCopy.ColumnMappings.Add(0, 1)
                    bulkCopy.ColumnMappings.Add(1, 2)
                    bulkCopy.ColumnMappings.Add(2, 3)
                    bulkCopy.ColumnMappings.Add(3, 4)
                    bulkCopy.WriteToServer(rd)
                    rd.Close()
                    MsgBox("Upload Customer Success")
                Catch ex As Exception
                    MsgBox(ex.Message)
                End Try
            End Using
        End If
    End Sub

    Private Sub Command113_Click(sender As Object, e As EventArgs) Handles Command113.Click
        OpenFileDialog1.InitialDirectory = My.Computer.FileSystem.SpecialDirectories.MyDocuments
        If OpenFileDialog1.ShowDialog(Me) = System.Windows.Forms.DialogResult.OK Then
            Dim xlApp As New Microsoft.Office.Interop.Excel.Application
            Dim xlWorkBook As Microsoft.Office.Interop.Excel.Workbook = xlApp.Workbooks.Open(OpenFileDialog1.FileName)
            Dim SheetName As String = xlWorkBook.Worksheets(1).Name.ToString
            Dim excelpath As String = OpenFileDialog1.FileName
            Dim koneksiExcel As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & excelpath & ";Extended Properties='Excel 8.0;HDR=No;IMEX=1;'"
            oleCon = New OleDbConnection(koneksiExcel)
            oleCon.Open()

            Dim queryExcel As String = "select * from [" & SheetName & "$]"
            Dim cmd As OleDbCommand = New OleDbCommand(queryExcel, oleCon)
            Dim rd As OleDbDataReader

            Call koneksi_db()

            Dim deleteReset As New SqlCommand("Delete from [dbo].[labelSelectionTable]", Main.koneksi)

            Try
                deleteReset.ExecuteNonQuery()
                'Show Msg Box
                'Timer1.Enabled = True
                'MsgBox("database Delete & Reset ID !")
            Catch ex As Exception
                MsgBox("Delete DB Fail " & ex.Message)
            End Try

            Using bulkCopy As SqlBulkCopy = New SqlBulkCopy(Main.koneksi)
                bulkCopy.DestinationTableName = "dbo.labelSelectionTable"
                bulkCopy.BulkCopyTimeout = 300 '5 menit
                Try
                    rd = cmd.ExecuteReader
                    bulkCopy.WriteToServer(rd)
                    rd.Close()
                    MsgBox("Upload Label Selection Success")
                Catch ex As Exception
                    MsgBox(ex.Message)
                End Try
            End Using
        End If
    End Sub

    Private Sub Command279_Click(sender As Object, e As EventArgs) Handles Command279.Click
        OpenFileDialog1.InitialDirectory = My.Computer.FileSystem.SpecialDirectories.MyDocuments
        If OpenFileDialog1.ShowDialog(Me) = System.Windows.Forms.DialogResult.OK Then
            Dim xlApp As New Microsoft.Office.Interop.Excel.Application
            Dim xlWorkBook As Microsoft.Office.Interop.Excel.Workbook = xlApp.Workbooks.Open(OpenFileDialog1.FileName)
            Dim SheetName As String = xlWorkBook.Worksheets(1).Name.ToString
            Dim excelpath As String = OpenFileDialog1.FileName
            Dim koneksiExcel As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & excelpath & ";Extended Properties='Excel 8.0;HDR=No;IMEX=1;'"
            oleCon = New OleDbConnection(koneksiExcel)
            oleCon.Open()

            Dim queryExcel As String = "select * from [" & SheetName & "$]"
            Dim cmd As OleDbCommand = New OleDbCommand(queryExcel, oleCon)
            Dim rd As OleDbDataReader

            Call koneksi_db()

            Dim deleteReset As New SqlCommand("Delete from [dbo].[BOM] DBCC CHECKIDENT ('[BOM]',RESEED,0)", Main.koneksi)

            Try
                deleteReset.ExecuteNonQuery()
                'Show Msg Box
                'Timer1.Enabled = True
                'MsgBox("database Delete & Reset ID !")
            Catch ex As Exception
                MsgBox("Delete DB Fail " & ex.Message)
            End Try

            Using bulkCopy As SqlBulkCopy = New SqlBulkCopy(Main.koneksi)
                bulkCopy.DestinationTableName = "dbo.BOM"
                bulkCopy.BulkCopyTimeout = 300 '5 menit
                Try
                    rd = cmd.ExecuteReader
                    bulkCopy.ColumnMappings.Add(0, 1)
                    bulkCopy.ColumnMappings.Add(1, 2)
                    bulkCopy.ColumnMappings.Add(2, 3)
                    bulkCopy.ColumnMappings.Add(3, 4)
                    bulkCopy.ColumnMappings.Add(4, 5)
                    bulkCopy.ColumnMappings.Add(5, 6)
                    bulkCopy.ColumnMappings.Add(6, 7)
                    bulkCopy.WriteToServer(rd)
                    rd.Close()
                    MsgBox("Upload BOM Success")
                Catch ex As Exception
                    MsgBox(ex.Message)
                End Try
            End Using
        End If
    End Sub

    Private Sub Command418_Click(sender As Object, e As EventArgs) Handles Command418.Click
        Dim qty As Integer = Convert.ToDecimal(Quantity.Text)
        'Printing Components
        Try
            label5_printer.PrintSettings.PrinterName = label_printer.PrintSettings.PrinterName
            Label5_SetValue()
            If qty = 0 Then qty = 1
            label5_printer.Print(qty)
        Catch ex As Exception
            MsgBox("Print Failed")
        End Try

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        reload_printer()
        Label5_SetValue()

        'declaration of Preview
        Dim LabelPreviewSettings As ILabelPreviewSettings = New LabelPreviewSettings()

        'setting preview format
        LabelPreviewSettings.ImageFormat = "PNG"
        LabelPreviewSettings.Width = Form_preview.pictureBoxPreview.Width                   ' Width Of image To generate
        LabelPreviewSettings.Height = Form_preview.pictureBoxPreview.Height                 ' Height Of image To generate

        ' Generate Preview File
        Dim imageObj As Object = label5_printer.GetLabelPreview(LabelPreviewSettings)

        'Display image in UI
        If TypeOf imageObj Is Byte() Then
            Form_preview.pictureBoxPreview.Image = ByteToImage(CType(imageObj, Byte()))
        ElseIf TypeOf imageObj Is String Then
            Form_preview.pictureBoxPreview.ImageLocation = CStr(imageObj)
        End If

        Form_preview.Show()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        reload_printer()
    End Sub

    Sub reload_printer()
        Dim appPath As String = Application.StartupPath()
        label_printer = PrintEngineFactory.PrintEngine.OpenLabel(appPath & "\Label\" & "generic.nlbl")
        label1_printer = PrintEngineFactory.PrintEngine.OpenLabel(appPath & "\Label\" & "adaptation.nlbl")
        label2_printer = PrintEngineFactory.PrintEngine.OpenLabel(appPath & "\Label\" & "testReport.nlbl")
        label3_printer = PrintEngineFactory.PrintEngine.OpenLabel(appPath & "\Label\" & "COC.nlbl")

        'label4_printer = PrintEngineFactory.PrintEngine.OpenLabel(appPath & "\Label\" & "Loose2.nlbl")
        label5_printer = PrintEngineFactory.PrintEngine.OpenLabel(appPath & "\Label\" & "Components.nlbl")
        label6_printer = PrintEngineFactory.PrintEngine.OpenLabel(appPath & "\Label\" & "Loose2.nlbl")

        'Fuji
        label_side_printer = PrintEngineFactory.PrintEngine.OpenLabel(appPath & "\Label\" & "Fuji Product Label side label.nlbl")
        label_rotary_printer = PrintEngineFactory.PrintEngine.OpenLabel(appPath & "\Label\" & "Fuji Rotary Handle Label.nlbl")
        label_front_long_printer = PrintEngineFactory.PrintEngine.OpenLabel(appPath & "\Label\" & "Fuji Front Label long.nlbl")
        label_front_short_printer = PrintEngineFactory.PrintEngine.OpenLabel(appPath & "\Label\" & "Fuji Front Label short.nlbl")
        label_carton_printer = PrintEngineFactory.PrintEngine.OpenLabel(appPath & "\Label\" & "Fuji Carton Label.nlbl")
        label_out_side_printer = PrintEngineFactory.PrintEngine.OpenLabel(appPath & "\Label\" & "Fuji Outside Grouping Label.nlbl")

        'Ruby
        label_performance_small_ruby = PrintEngineFactory.PrintEngine.OpenLabel(appPath & "\Label\" & "Ruby Performace Label Small.nlbl")
        label_performance_big_ruby = PrintEngineFactory.PrintEngine.OpenLabel(appPath & "\Label\" & "Ruby Performace Label Big.nlbl")
        label_packaging_ruby = PrintEngineFactory.PrintEngine.OpenLabel(appPath & "\Label\" & "Ruby Packaging Label.nlbl")
        label_outside_ruby = PrintEngineFactory.PrintEngine.OpenLabel(appPath & "\Label\" & "Ruby Out Side Label.nlbl")


        printers = PrintEngineFactory.PrintEngine.Printers

        label_printer.PrintSettings.PrinterName = listprinter.Text
        label1_printer.PrintSettings.PrinterName = listprinter1.Text
        label2_printer.PrintSettings.PrinterName = listprinter2.Text
        label3_printer.PrintSettings.PrinterName = listprinter3.Text

        'fuji
        label_side_printer.PrintSettings.PrinterName = cbx_fuji_side_label.Text
        label_rotary_printer.PrintSettings.PrinterName = cbx_Rotary.Text
        label_front_long_printer.PrintSettings.PrinterName = cbx_front.Text
        label_front_short_printer.PrintSettings.PrinterName = cbx_front.Text
        label_carton_printer.PrintSettings.PrinterName = cbx_Carton.Text
        label_out_side_printer.PrintSettings.PrinterName = cbx_outside.Text

        'ruby
        label_performance_small_ruby.PrintSettings.PrinterName = cbxPerfomaceRuby.Text
        label_performance_big_ruby.PrintSettings.PrinterName = cbxPerfomaceRuby.Text
        label_packaging_ruby.PrintSettings.PrinterName = cbxPackagingRuby.Text
        label_outside_ruby.PrintSettings.PrinterName = cbxOutsideRuby.Text

    End Sub

    'Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
    '    Dim qty As Integer = 1
    '    'Printing Components
    '    Try
    '        label6_printer.PrintSettings.PrinterName = label_printer.PrintSettings.PrinterName
    '        Label6_SetValue()
    '        If qty = 0 Then qty = 1
    '        label6_printer.Print(qty)
    '    Catch ex As Exception
    '        MsgBox("Print Failed")
    '    End Try
    'End Sub

    'Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
    '    reload_printer()
    '    Label6_SetValue()

    '    'declaration of Preview
    '    Dim LabelPreviewSettings As ILabelPreviewSettings = New LabelPreviewSettings()

    '    'setting preview format
    '    LabelPreviewSettings.ImageFormat = "PNG"
    '    LabelPreviewSettings.Width = Form_preview.pictureBoxPreview.Width                   ' Width Of image To generate
    '    LabelPreviewSettings.Height = Form_preview.pictureBoxPreview.Height                 ' Height Of image To generate

    '    ' Generate Preview File
    '    Dim imageObj As Object = label6_printer.GetLabelPreview(LabelPreviewSettings)

    '    'Display image in UI
    '    If TypeOf imageObj Is Byte() Then
    '        Form_preview.pictureBoxPreview.Image = ByteToImage(CType(imageObj, Byte()))
    '    ElseIf TypeOf imageObj Is String Then
    '        Form_preview.pictureBoxPreview.ImageLocation = CStr(imageObj)
    '    End If

    '    Form_preview.Show()
    'End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub Cek_OF_Click() Handles Cek_OF.Click
        Dim str2 As String
        str2 = "SELECT *  FROM openOrders where [order]='" & testPP.Text & "'"
        Dim adapter2 As New SqlDataAdapter(str2, Main.koneksi)
        Dim ds2 As New DataSet
        adapter2.Fill(ds2)
        'Call Main.koneksi_db()



        Dim str As String

        TipeMaterial.Text = Microsoft.VisualBasic.Left(Me.testDescription.Text, 2)
        jmlOF.Text = 0

        For i As Integer = 0 To ds2.Tables(0).Rows.Count - 1
            str = "SELECt [category] FROM BOM where [material]  = '" & ds2.Tables(0).Rows(i).Item("Material").ToString & "'"
            Dim adapter As New SqlDataAdapter(str, Main.koneksi)
            Dim ds As New DataSet
            adapter.Fill(ds)

            If ds.Tables(0).Rows.Count > 0 Then
                ' MsgBox(ds.Tables(0).Rows(0).Item("category").ToString)
                If ds.Tables(0).Rows(0).Item("category").ToString = "OF" Or ds.Tables(0).Rows(0).Item("category").ToString = "OF Individual" Then
                    jmlOF.Text = (CInt(jmlOF.Text) + CInt(ds2.Tables(0).Rows(i).Item("Reqmts qty")))
                End If
            End If
        Next

        'If CInt(jmlOF.Text) > 0 And CInt(LabelQuantityitem.Text) > 0 Then
        '    'If TipeMaterial.Text = "NS" Then
        '    '    IC_OF.Text = (CInt(jmlOF.Text) / CInt(LabelQuantityitem.Text)) * 4
        '    'End If

        '    'If TipeMaterial.Text = "NT" Or TipeMaterial.Text = "NV" Or TipeMaterial.Text = "NW" Then
        '    '    IC_OF.Text = (CInt(jmlOF.Text) / CInt(LabelQuantityitem.Text)) * 4
        '    '    IC_OF.Text = CInt(IC_OF.Text) + 4
        '    'End If
        '    'If TipeMaterial.Text = "NS" Or TipeMaterial.Text = "NT" Or TipeMaterial.Text = "NV" Or TipeMaterial.Text = "NW" Then
        '    If TipeMaterial.Text = "NS" Or TipeMaterial.Text = "NV" Or TipeMaterial.Text = "NW" Then
        '        IC_OF.Text = (CInt(jmlOF.Text) / CInt(LabelQuantityitem.Text)) * 4
        '        'IC_OF.Text = (CInt(jmlOF.Text) / CInt(quantityLabel.Text)) * 4
        '        IC_OF.Text = CInt(IC_OF.Text) + 4
        '    ElseIf TipeMaterial.Text = "NT" Then
        '        IC_OF.Text = 4
        '    Else
        '        IC_OF.Text = 4
        '    End If
        'Else
        '    'If TipeMaterial.Text = "NS" Then
        '    '    IC_OF.Text = 0
        '    'End If

        '    'If TipeMaterial.Text = "NT" Or TipeMaterial.Text = "NV" Or TipeMaterial.Text = "NW" Then
        '    '    IC_OF.Text = 4
        '    'End If
        '    If TipeMaterial.Text = "NT" Or TipeMaterial.Text = "NV" Or TipeMaterial.Text = "NW" Then
        '        IC_OF.Text = 4
        '    ElseIf TipeMaterial.Text = "NS" Then
        '        IC_OF.Text = ""
        '    End If
        'End If
    End Sub


    Private Sub LabelQuantityitem_GotFocus(sender As Object, e As EventArgs) Handles LabelQuantityitem.GotFocus
        'ComponentNo.Select()
        CompToQuality.Select()
    End Sub

    Private Sub DeleteDuplicateBOM_Click(sender As Object, e As EventArgs) Handles DeleteDuplicateBOM.Click
        Dim str2 As String = "declare @abc datetime
set @abc = (SELECT TOP (1) [Date]
  FROM [SGRAC_MES].[dbo].[Components] where [order]='" & testPP.Text & "' ORDER BY [ID] DESC);

  DELETE FROM [SGRAC_MES].[dbo].[Components] where [order]='" & testPP.Text & "' and [date] = @abc;"

        'Dim str2 As String
        'str2 = "DELETE FROM [SGRAC_MES].[dbo].[Components] where [order]='" & testPP.Text & "' and [date] ='" & b & "'"
        Dim adapter2 As New SqlDataAdapter(str2, Main.koneksi)
        Dim ds2 As New DataSet
        adapter2.Fill(ds2)

        PPnumberEntry.Select()
        SendKeys.Send("{ENTER}")

    End Sub

    Private Sub Print_Label_Outside_Click(sender As Object, e As EventArgs) Handles Print_Label_Outside.Click
        '******************************************************************


        Dim PrintCount As Integer = 0
        Label_SetValue()
        PrintCount = 0
        For i = 1 To Me.labelQty.Text - 1

            label_printer.Variables("label13").SetValue(Me.boxQty.Text)
            Dim qty As Integer = 1
            Try
                label_printer.Print(qty)
            Catch ex As Exception
                MsgBox("Printing Label Cancel " & ex.Message)
            End Try

            PrintCount = PrintCount + CInt(Me.boxQty.Text)

        Next
        label_printer.Variables("label13").SetValue(Convert.ToDecimal(Me.Quantity.Text) - PrintCount)

        Try
            label_printer.Print(1)
        Catch ex As Exception
            MsgBox("Printing Label Cancel " & ex.Message)
        End Try
        '********************************************************************************
    End Sub

    ' Private Sub TextBox1_TextChanged(sender As Object, e As PreviewKeyDownEventArgs) Handles TextBox1.PreviewKeyDown, TextBox1.TextChanged
    'Call Main.koneksi_db()
    'If (e.KeyData = Keys.Tab Or e.KeyData = Keys.Enter) And Len(TextBox1.Text) >= 11 Then
    '    DataGridView2.Rows.Clear()
    '    DataGridView3.Rows.Clear()
    '    DataGridView4.Rows.Clear()

    '    Try
    '        Dim query = "
    '            SELECT 
    '            convert(varchar, [Day], 101) as Day,
    '            convert(varchar, [FinishDate], 101) as Finish_Date,
    '            convert(varchar, [GI date], 101) as GI_Date,
    '            convert(varchar, [SchedFinishDate], 101) as Scheduled_Finish,
    '            [PP],
    '            [Quantityadapted],
    '            [Quantity],
    '            [workstation],
    '            [Material],
    '            [Description],
    '            [range],
    '            [Customer],
    '            [Name1],
    '            [Entity],
    '            [City],
    '            [SO no],
    '            [Item],
    '            [On time],
    '            [On timeM]
    '            FROM [SGRAC_MES].[dbo].[ProductionOrders] 
    '            where [PP]='" & TextBox1.Text & "'"
    '        Dim sc As New SqlCommand(query, Main.koneksi)
    '        Dim adapter As New SqlDataAdapter(sc)
    '        Dim ds As New DataSet
    '        adapter.Fill(ds)

    '        DataGridView2.ColumnCount = 19
    '        DataGridView2.Columns(0).Name = "Day"
    '        DataGridView2.Columns(1).Name = "Finish_Date"
    '        DataGridView2.Columns(2).Name = "GI_Date"
    '        DataGridView2.Columns(3).Name = "Scheduled_Finish"
    '        DataGridView2.Columns(4).Name = "PP"
    '        DataGridView2.Columns(5).Name = "Adapted"
    '        DataGridView2.Columns(6).Name = "Required"
    '        DataGridView2.Columns(7).Name = "WorkStation"
    '        DataGridView2.Columns(8).Name = "Material"
    '        DataGridView2.Columns(9).Name = "Description"
    '        DataGridView2.Columns(10).Name = "Range"
    '        DataGridView2.Columns(11).Name = "Customer"
    '        DataGridView2.Columns(12).Name = "Name 1"
    '        DataGridView2.Columns(13).Name = "Entity"
    '        DataGridView2.Columns(14).Name = "City"
    '        DataGridView2.Columns(15).Name = "SO no"
    '        DataGridView2.Columns(16).Name = "Item"
    '        DataGridView2.Columns(17).Name = "On time"
    '        DataGridView2.Columns(18).Name = "On timeM"
    '        For r = 0 To ds.Tables(0).Rows.Count - 1
    '            Dim row As String() = New String() {
    '                ds.Tables(0).Rows(r).Item("Day").ToString(),
    '                ds.Tables(0).Rows(r).Item("Finish_Date").ToString(),
    '                ds.Tables(0).Rows(r).Item("GI_Date").ToString(),
    '                ds.Tables(0).Rows(r).Item("Scheduled_Finish").ToString(),
    '                ds.Tables(0).Rows(r).Item("PP").ToString(),
    '                ds.Tables(0).Rows(r).Item("Quantityadapted").ToString(),
    '                ds.Tables(0).Rows(r).Item("Quantity").ToString(),
    '                ds.Tables(0).Rows(r).Item("workstation").ToString(),
    '                ds.Tables(0).Rows(r).Item("Material").ToString(),
    '                ds.Tables(0).Rows(r).Item("Description").ToString(),
    '                ds.Tables(0).Rows(r).Item("range").ToString(),
    '                ds.Tables(0).Rows(r).Item("Customer").ToString(),
    '                ds.Tables(0).Rows(r).Item("Name1").ToString(),
    '                ds.Tables(0).Rows(r).Item("Entity").ToString(),
    '                ds.Tables(0).Rows(r).Item("City").ToString(),
    '                ds.Tables(0).Rows(r).Item("SO no").ToString(),
    '                ds.Tables(0).Rows(r).Item("Item").ToString(),
    '                ds.Tables(0).Rows(r).Item("On time").ToString(),
    '                ds.Tables(0).Rows(r).Item("On timeM").ToString()
    '            }
    '            DataGridView2.Rows.Add(row)
    '        Next

    '        Try
    '            Dim query2 = "
    '            Select * From DailyProduction 
    '            Where workingsection ='" & ds.Tables(0).Rows(0).Item("workstation").ToString() & "' 
    '            AND Convert(varchar,Jour,103) = '" & ds.Tables(0).Rows(0).Item("Finish_Date").ToString() & "'"
    '            Dim sc2 As New SqlCommand(query2, Main.koneksi)
    '            Dim adapter2 As New SqlDataAdapter(sc2)
    '            Dim ds2 As New DataSet
    '            adapter2.Fill(ds2)

    '            DataGridView3.ColumnCount = 3
    '            DataGridView3.Columns(0).Name = "Jour"
    '            DataGridView3.Columns(1).Name = "Employee"
    '            DataGridView3.Columns(2).Name = "WorkingSection"
    '            For r = 0 To ds2.Tables(0).Rows.Count - 1
    '                Dim row As String() = New String() {
    '                    ds2.Tables(0).Rows(r).Item("Jour").ToString(),
    '                    ds2.Tables(0).Rows(r).Item("Employee").ToString(),
    '                    ds2.Tables(0).Rows(r).Item("WorkingSection").ToString()
    '                }
    '                DataGridView3.Rows.Add(row)
    '            Next

    '        Catch ex As Exception
    '            MessageBox.Show(ex.Message)
    '        End Try

    '    Catch ex As Exception
    '        MessageBox.Show(ex.Message)
    '    End Try

    '    Try
    '        Dim query = "
    '            SELECT [Order]
    '            ,[Date]
    '            ,[Workstation]
    '            ,[Material]
    '            ,[Code]
    '            ,[Barcode]
    '            ,[Datecode]
    '            ,[Descr]
    '            ,[Reqmts qty]
    '            ,[Check Components]
    '            FROM [SGRAC_MES].[dbo].[ComponentOrders] 
    '            where [order]=" & TextBox1.Text
    '        Dim sc As New SqlCommand(query, Main.koneksi)
    '        Dim adapter As New SqlDataAdapter(sc)
    '        Dim ds As New DataSet
    '        adapter.Fill(ds)

    '        DataGridView4.ColumnCount = 10
    '        DataGridView4.Columns(0).Name = "PP"
    '        DataGridView4.Columns(1).Name = "Date"
    '        DataGridView4.Columns(2).Name = "Workstation"
    '        DataGridView4.Columns(3).Name = "Material"
    '        DataGridView4.Columns(4).Name = "Code"
    '        DataGridView4.Columns(5).Name = "Barcode"
    '        DataGridView4.Columns(6).Name = "Datecode"
    '        DataGridView4.Columns(7).Name = "Description"
    '        DataGridView4.Columns(8).Name = "Reqmts aty"
    '        DataGridView4.Columns(9).Name = "Check Components"
    '        For r = 0 To ds.Tables(0).Rows.Count - 1
    '            Dim row As String() = New String() {
    '                ds.Tables(0).Rows(r).Item("Order").ToString(),
    '                ds.Tables(0).Rows(r).Item("Date").ToString(),
    '                ds.Tables(0).Rows(r).Item("Workstation").ToString(),
    '                ds.Tables(0).Rows(r).Item("Material").ToString(),
    '                ds.Tables(0).Rows(r).Item("Code").ToString(),
    '                ds.Tables(0).Rows(r).Item("Barcode").ToString(),
    '                ds.Tables(0).Rows(r).Item("Datecode").ToString(),
    '                ds.Tables(0).Rows(r).Item("Descr").ToString(),
    '                ds.Tables(0).Rows(r).Item("Reqmts qty").ToString(),
    '                ds.Tables(0).Rows(r).Item("Check Components").ToString()
    '            }
    '            DataGridView4.Rows.Add(row)
    '        Next

    '    Catch ex As Exception
    '        MessageBox.Show(ex.Message)
    '    End Try

    '    'MsgBox(TextBox1.Text)
    'End If
    '  End Sub


    Private Sub Print_Product_x_Click(sender As Object, e As EventArgs) Handles Print_Product_x.Click
        Dim PrintCount As Integer = 0
        Dim i As Integer = 0
        Label_SetValue()
        PrintCount = 0
        For i = Convert.ToDecimal(StartLabel.Text) To Convert.ToDecimal(quantityLabel.Text) - 1

            'label_printer.Variables("label13").SetValue(Me.boxQty.Text)
            label_printer.Variables("label13").SetValue("1")
            label_printer.Variables("CounterItems").SetValue(i)

            Try
                label_printer.Variables("DataMatrix").SetValue("SG" & dateCode3() & Convert.ToDecimal(i.ToString).ToString("0000"))
            Catch ex As Exception

            End Try

            Dim qty As Integer = 1
            Try
                label_printer.Print(qty)
            Catch ex As Exception
                MsgBox("Printing Label Cancel " & ex.Message)
            End Try

            PrintCount = PrintCount + CInt(Me.boxQty.Text)

        Next
        ' label_printer.Variables("label13").SetValue(Convert.ToDecimal(Me.Quantity.Text) - PrintCount)
        label_printer.Variables("label13").SetValue("1")
        label_printer.Variables("CounterItems").SetValue(i)

        Try
            label_printer.Variables("DataMatrix").SetValue("SG" & dateCode3() & Convert.ToDecimal(i.ToString).ToString("0000"))
        Catch ex As Exception

        End Try

        Try
            label_printer.Print(1)
        Catch ex As Exception
            MsgBox("Printing Label Cancel " & ex.Message)
        End Try

        StartLabel.Text = Convert.ToDecimal(StartLabel.Text) + 1
        quantityLabel.Text = StartLabel.Text
    End Sub

    Private Sub Manual_Print_Product_Click(sender As Object, e As EventArgs) Handles Manual_Print_Product.Click
        label1_setValue()

        Dim a As Integer
        'Dim sisa As Integer
        Dim quantity_show As Integer

        For a = Convert.ToDecimal(StartPackingLabel.Text) To Convert.ToDecimal(labelQty.Text)

            If Convert.ToDecimal(Quantity.Text) <= Convert.ToDecimal(boxQty.Text) Then
                quantity_show = Convert.ToDecimal(Quantity.Text)
            Else
                'sisa = Convert.ToDecimal(Quantity.Text) - (a * Convert.ToDecimal(boxQty.Text))
                If (a * Convert.ToDecimal(boxQty.Text)) > Convert.ToDecimal(Quantity.Text) Then
                    quantity_show = Convert.ToDecimal(Quantity.Text) Mod Convert.ToDecimal(boxQty.Text)
                Else
                    quantity_show = Convert.ToDecimal(boxQty.Text)
                End If
            End If

            label1_printer.Variables("quantity").SetValue(quantity_show.ToString)
            label1_printer.Variables("Package").SetValue(a.ToString)

            'printing with quantity 
            'Dim qty As Integer = 1
            Try
                label1_printer.Print(1)
                Call koneksi_db()
                Dim sqlNew As String = "insert into [printingRecordPacking] ([PP],[print date],[print time], [user]) values ('" & Me.PP.Text & "',getDate(),getDate(),'" & Me.user.Text & "')"
                Dim cmdNew = New SqlCommand(sqlNew, Main.koneksi)
                cmdNew.ExecuteNonQuery()
            Catch ex As Exception
                MsgBox("Printing Packaging Cancel " & ex.Message)
            End Try

        Next

        StartPackingLabel.Text = Convert.ToDecimal(StartPackingLabel.Text) + 1
        labelQty.Text = StartPackingLabel.Text

    End Sub

    Private Sub En_print_CheckedChanged(sender As Object, e As EventArgs) Handles en_print.CheckedChanged

        'My.Settings.Save()

        If en_print.Checked = True Then
            Manual_Print_Product.Visible = True
            Print_Product_x.Visible = True
        Else
            Manual_Print_Product.Visible = False
            Print_Product_x.Visible = False
        End If

    End Sub

    Private Sub Trace_Click(sender As Object, e As EventArgs) Handles Trace.Click


        Call Main.koneksi_db()
        'e.KeyData = Keys.Tab Or e.KeyData = Keys.Enter) And
        If Len(TextBox1.Text) >= 33 Then
            DataGridView2.Rows.Clear()
            DataGridView3.Rows.Clear()
            DataGridView4.Rows.Clear()
            Dim dsTrace As New DataSet
            Dim sql = "select DISTINCT ProductionOrders.PP from ProductionOrders, printingRecord where ProductionOrders.PP=printingRecord.PP and printingRecord.QRCodeFuji='" & TextBox2.Text & "'"
            Dim adapter = New SqlDataAdapter(sql, Main.koneksi)
            adapter.Fill(dsTrace)
            If dsTrace.Tables(0).Rows.Count > 0 Then
                TextBox1.Text = dsTrace.Tables(0).Rows(0).Item("PP")
                'Trace.PerformClick()
            Else
                MessageBox.Show("Sorry Data Not Found")
            End If

        ElseIf TextBox1.Text.Contains("SG") Then

            Dim ds As New DataSet
            Dim queryCheck = "select PP FROM SGRAC_MES.dbo.printingRecord pr WHERE Seq ='" & TextBox1.Text & "'"
            Dim CheckAdap = New SqlDataAdapter(queryCheck, Main.koneksi)
            CheckAdap.Fill(ds)

            If ds.Tables(0).Rows.Count > 0 Then
                TextBox1.Text = ds.Tables(0).Rows(0).Item("PP")
            End If

        End If


        If Len(TextBox1.Text) <= 12 Then
            DataGridView2.Rows.Clear()
            DataGridView3.Rows.Clear()
            DataGridView4.Rows.Clear()

            Try
                Dim query = "
                    SELECT 
                    convert(varchar, [Day], 101) as Day,
                    convert(varchar, [FinishDate], 101) as Finish_Date,
                    convert(varchar, [GI date], 101) as GI_Date,
                    convert(varchar, [SchedFinishDate], 101) as Scheduled_Finish,
                    [PP],
                    [Quantityadapted],
                    [Quantity],
                    [workstation],
                    [Material],
                    [Description],
                    [range],
                    [Customer],
                    [Name1],
                    [Entity],
                    [City],
                    [SO no],
                    [Item],
                    [On time],
                    [On timeM]
                    FROM [SGRAC_MES].[dbo].[ProductionOrders] 
                    where [PP]='" & TextBox1.Text & "'"
                Dim sc As New SqlCommand(query, Main.koneksi)
                Dim adapter As New SqlDataAdapter(sc)
                Dim ds As New DataSet
                adapter.Fill(ds)

                DataGridView2.ColumnCount = 19
                DataGridView2.Columns(0).Name = "Day"
                DataGridView2.Columns(1).Name = "Finish_Date"
                DataGridView2.Columns(2).Name = "GI_Date"
                DataGridView2.Columns(3).Name = "Scheduled_Finish"
                DataGridView2.Columns(4).Name = "PP"
                DataGridView2.Columns(5).Name = "Adapted"
                DataGridView2.Columns(6).Name = "Required"
                DataGridView2.Columns(7).Name = "WorkStation"
                DataGridView2.Columns(8).Name = "Material"
                DataGridView2.Columns(9).Name = "Description"
                DataGridView2.Columns(10).Name = "Range"
                DataGridView2.Columns(11).Name = "Customer"
                DataGridView2.Columns(12).Name = "Name 1"
                DataGridView2.Columns(13).Name = "Entity"
                DataGridView2.Columns(14).Name = "City"
                DataGridView2.Columns(15).Name = "SO no"
                DataGridView2.Columns(16).Name = "Item"
                DataGridView2.Columns(17).Name = "On time"
                DataGridView2.Columns(18).Name = "On timeM"
                For r = 0 To ds.Tables(0).Rows.Count - 1
                    Dim row As String() = New String() {
                        ds.Tables(0).Rows(r).Item("Day").ToString(),
                        ds.Tables(0).Rows(r).Item("Finish_Date").ToString(),
                        ds.Tables(0).Rows(r).Item("GI_Date").ToString(),
                        ds.Tables(0).Rows(r).Item("Scheduled_Finish").ToString(),
                        ds.Tables(0).Rows(r).Item("PP").ToString(),
                        ds.Tables(0).Rows(r).Item("Quantityadapted").ToString(),
                        ds.Tables(0).Rows(r).Item("Quantity").ToString(),
                        ds.Tables(0).Rows(r).Item("workstation").ToString(),
                        ds.Tables(0).Rows(r).Item("Material").ToString(),
                        ds.Tables(0).Rows(r).Item("Description").ToString(),
                        ds.Tables(0).Rows(r).Item("range").ToString(),
                        ds.Tables(0).Rows(r).Item("Customer").ToString(),
                        ds.Tables(0).Rows(r).Item("Name1").ToString(),
                        ds.Tables(0).Rows(r).Item("Entity").ToString(),
                        ds.Tables(0).Rows(r).Item("City").ToString(),
                        ds.Tables(0).Rows(r).Item("SO no").ToString(),
                        ds.Tables(0).Rows(r).Item("Item").ToString(),
                        ds.Tables(0).Rows(r).Item("On time").ToString(),
                        ds.Tables(0).Rows(r).Item("On timeM").ToString()
                    }
                    DataGridView2.Rows.Add(row)
                Next

                Try
                    Dim query2 = "
                    Select * From DailyProduction 
                    Where workingsection ='" & ds.Tables(0).Rows(0).Item("workstation").ToString() & "' 
                    AND Convert(varchar,Jour,101) = '" & ds.Tables(0).Rows(0).Item("Finish_Date").ToString() & "'"
                    Dim sc2 As New SqlCommand(query2, Main.koneksi)
                    Dim adapter2 As New SqlDataAdapter(sc2)
                    Dim ds2 As New DataSet
                    adapter2.Fill(ds2)

                    DataGridView3.ColumnCount = 3
                    DataGridView3.Columns(0).Name = "Jour"
                    DataGridView3.Columns(1).Name = "Employee"
                    DataGridView3.Columns(2).Name = "WorkingSection"
                    For r = 0 To ds2.Tables(0).Rows.Count - 1
                        Dim row As String() = New String() {
                            ds2.Tables(0).Rows(r).Item("Jour").ToString(),
                            ds2.Tables(0).Rows(r).Item("Employee").ToString(),
                            ds2.Tables(0).Rows(r).Item("WorkingSection").ToString()
                        }
                        DataGridView3.Rows.Add(row)
                    Next

                Catch ex As Exception
                    MessageBox.Show(ex.Message)
                End Try

            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try

            Try
                Dim query = "
                    SELECT [Order]
                    ,[Date]
                    ,[Workstation]
                    ,[Material]
                    ,[Code]
                    ,[Barcode]
                    ,[Datecode]
                    ,[Descr]
                    ,[Reqmts qty]
                    ,[Check Components]
                    FROM [SGRAC_MES].[dbo].[ComponentOrders] 
                    where [order]=" & TextBox1.Text
                Dim sc As New SqlCommand(query, Main.koneksi)
                Dim adapter As New SqlDataAdapter(sc)
                Dim ds As New DataSet
                adapter.Fill(ds)

                DataGridView4.ColumnCount = 10
                DataGridView4.Columns(0).Name = "PP"
                DataGridView4.Columns(1).Name = "Date"
                DataGridView4.Columns(2).Name = "Workstation"
                DataGridView4.Columns(3).Name = "Material"
                DataGridView4.Columns(4).Name = "Code"
                DataGridView4.Columns(5).Name = "Barcode"
                DataGridView4.Columns(6).Name = "Datecode"
                DataGridView4.Columns(7).Name = "Description"
                DataGridView4.Columns(8).Name = "Reqmts aty"
                DataGridView4.Columns(9).Name = "Check Components"
                For r = 0 To ds.Tables(0).Rows.Count - 1
                    Dim row As String() = New String() {
                        ds.Tables(0).Rows(r).Item("Order").ToString(),
                        ds.Tables(0).Rows(r).Item("Date").ToString(),
                        ds.Tables(0).Rows(r).Item("Workstation").ToString(),
                        ds.Tables(0).Rows(r).Item("Material").ToString(),
                        ds.Tables(0).Rows(r).Item("Code").ToString(),
                        ds.Tables(0).Rows(r).Item("Barcode").ToString(),
                        ds.Tables(0).Rows(r).Item("Datecode").ToString(),
                        ds.Tables(0).Rows(r).Item("Descr").ToString(),
                        ds.Tables(0).Rows(r).Item("Reqmts qty").ToString(),
                        ds.Tables(0).Rows(r).Item("Check Components").ToString()
                    }
                    DataGridView4.Rows.Add(row)
                Next

            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try

            'MsgBox(TextBox1.Text)
        End If
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        OpenFileDialog1.InitialDirectory = My.Computer.FileSystem.SpecialDirectories.MyDocuments
        If OpenFileDialog1.ShowDialog(Me) = System.Windows.Forms.DialogResult.OK Then
            Dim xlApp As New Microsoft.Office.Interop.Excel.Application
            Dim xlWorkBook As Microsoft.Office.Interop.Excel.Workbook = xlApp.Workbooks.Open(OpenFileDialog1.FileName)
            Dim SheetName As String = xlWorkBook.Worksheets(1).Name.ToString
            Dim excelpath As String = OpenFileDialog1.FileName
            Dim koneksiExcel As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & excelpath & ";Extended Properties='Excel 8.0;HDR=No;IMEX=1;'"
            oleCon = New OleDbConnection(koneksiExcel)
            oleCon.Open()

            Dim queryExcel As String = "select * from [" & SheetName & "$A2:C]"
            Dim cmd As OleDbCommand = New OleDbCommand(queryExcel, oleCon)
            Dim rd As OleDbDataReader

            Call koneksi_db()

            Dim deleteReset As New SqlCommand("Delete from [dbo].[QualityIssue]", Main.koneksi)

            Try
                deleteReset.ExecuteNonQuery()
            Catch ex As Exception
                MsgBox("Delete DB Fail " & ex.Message)
            End Try

            Using bulkCopy As SqlBulkCopy = New SqlBulkCopy(Main.koneksi)
                bulkCopy.DestinationTableName = "dbo.QualityIssue"
                bulkCopy.BulkCopyTimeout = 300 '5 menit
                Try
                    rd = cmd.ExecuteReader
                    bulkCopy.ColumnMappings.Add(0, 1)
                    bulkCopy.ColumnMappings.Add(1, 2)
                    bulkCopy.ColumnMappings.Add(2, 3)
                    bulkCopy.WriteToServer(rd)
                    rd.Close()
                    MsgBox("Upload Quality Issue Success")
                    DGV_Quality()
                Catch ex As Exception
                    MsgBox(ex.Message)
                End Try
            End Using
        End If
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Check_Quality.Click
        If txt_Ref.Text <> "" And txt_PlanCode_DateCode.Text <> "" Then
            Try

                checkQualityIssue(txt_Ref.Text, txt_PlanCode_DateCode.Text)

                'Santo  
                'Simpan_QualityIssueRecord()
            Catch ex As Exception
                'MsgBox(ex.ToString)
            End Try
        End If
    End Sub

    Function checkQualityIssue(comp, pcdc) As String
        Try
            Dim ds As New DataSet
            Dim dsMCode As New DataSet
            Dim SelectQuality As String = "select * from [QualityIssue] where [references] ='" & comp.ToString.ToUpper & "'"
            Dim adapterQuality As New SqlDataAdapter(SelectQuality, Main.koneksi)
            adapterQuality.Fill(ds)
            Dim hasil As String = ""
            Dim tampungpcdc As String = ""

            If ds.Tables(0).Rows.Count = 0 Then
                Lbl_Result.Text = "OK"
            End If

            If ds.Tables(0).Rows.Count >= 1 Then
                Dim length = Len(pcdc.ToString.ToUpper) - 5
                Dim PlantCode = Microsoft.VisualBasic.Left(pcdc.ToString.ToUpper, length)
                Dim DateCode = Microsoft.VisualBasic.Right(pcdc.ToString.ToUpper, 5)

                'dim var_OK_NOK As String

                For i = 0 To ds.Tables(0).Rows.Count - 1
                    Dim DBlengthStart = Len(ds.Tables(0).Rows(i).Item("Start Of Impact").ToString) - 5
                    Dim DBStartPC = Microsoft.VisualBasic.Left(ds.Tables(0).Rows(i).Item("Start Of Impact").ToString, DBlengthStart)
                    Dim DBStartDC = Microsoft.VisualBasic.Right(ds.Tables(0).Rows(i).Item("Start Of Impact").ToString, 5)
                    Dim DBlengthEnd = Len(ds.Tables(0).Rows(i).Item("End Of Impact").ToString) - 5
                    Dim DBEndPC = Microsoft.VisualBasic.Left(ds.Tables(0).Rows(i).Item("End Of Impact").ToString, DBlengthEnd)
                    Dim DBEndDC = Microsoft.VisualBasic.Right(ds.Tables(0).Rows(i).Item("End Of Impact").ToString, 5)
                    If PlantCode = DBStartPC Then
                        If DBStartDC <> tampungpcdc Then
                            If CInt(DBStartDC) <= CInt(DateCode) And CInt(DBEndDC) >= CInt(DateCode) Then
                                hasil = "NOK"
                                Dim Query = "INSERT INTO [dbo].[QualityIssueRecord] ([PP],[References],[PlantCodeDateCode],[Remark]) VALUES ('" & PPnumberEntry.Text & "','" & comp.ToUpper & "','" & pcdc.ToUpper & "','" & hasil & "')"
                                Dim InsertBro = New SqlDataAdapter(Query, Main.koneksi)
                                InsertBro.SelectCommand.ExecuteNonQuery()
                                Lbl_Result.Text = hasil
                                Exit Function
                            Else
                                hasil = "OK"
                                'Dim Query = "INSERT INTO [dbo].[QualityIssueRecord] ([PP],[References],[PlantCodeDateCode],[Remark]) VALUES ('" & PPnumberEntry.Text & "','" & comp.ToUpper & "','" & pcdc.ToUpper & "','" & hasil & "')"
                                'Dim InsertBro = New SqlDataAdapter(Query, Main.koneksi)
                                'InsertBro.SelectCommand.ExecuteNonQuery()
                            End If
                            tampungpcdc = DBStartDC
                        End If
                    Else
                        hasil = "OK"
                    End If
                Next

                If hasil = "OK" Then
                    Dim Query = "INSERT INTO [dbo].[QualityIssueRecord] ([PP],[References],[PlantCodeDateCode],[Remark]) VALUES ('" & PPnumberEntry.Text & "','" & comp.ToUpper & "','" & pcdc.ToUpper & "','" & hasil & "')"
                    Dim InsertBro = New SqlDataAdapter(Query, Main.koneksi)
                    InsertBro.SelectCommand.ExecuteNonQuery()
                End If

                Lbl_Result.Text = hasil

                Dim SelectMPlant As String = "select * from [MasterPlantCode] where [Plant Code] ='" & PlantCode & "'"
                Dim adapterMPlant As New SqlDataAdapter(SelectMPlant, Main.koneksi)
                adapterMPlant.Fill(dsMCode)

                If dsMCode.Tables(0).Rows.Count = 1 Then
                    lbl_PlantCode.Text = "Plant = " + dsMCode.Tables(0).Rows(0).Item("Name Plant Code").ToString()
                Else
                    lbl_PlantCode.Text = "Plant Code Not In DB"
                End If

            End If

        Catch ex As Exception
            'MsgBox(ex.ToString)
        End Try
        checkQualityIssue = Lbl_Result.Text
    End Function

    Private Sub Result_TextChanged(sender As Object, e As EventArgs) Handles Lbl_Result.TextChanged, lbl_PlantCode.TextChanged
        If Lbl_Result.Text = "OK" Then
            Lbl_Result.ForeColor = Color.Green
            Lbl_Result.BackColor = Color.LightGreen
        ElseIf Lbl_Result.Text = "NOK" Then
            Lbl_Result.ForeColor = Color.Red
            Lbl_Result.BackColor = Color.LightCoral
        End If
    End Sub

    Private Sub Txt_QRCode_TextChanged(sender As Object, e As PreviewKeyDownEventArgs) Handles txt_QRCode.PreviewKeyDown
        'If txt_QRCode.TextLength >= 300 And (e.KeyData = Keys.Enter Or e.KeyData = Keys.Tab) Then
        'If txt_QRCode.TextLength >= 32 And (e.KeyData = Keys.Enter Or e.KeyData = Keys.Tab) Then
        If (e.KeyData = Keys.Enter Or e.KeyData = Keys.Tab) Then
            If txt_QRCode.Text.Length >= 50 Then
                MsgBox("Wrong QRCode !")
                Exit Sub
            End If
            Combo_PlanCode.Text = ""
            Txt_Barcode.Text = ""
            txt_Ref.Text = txt_QRCode.Text.Substring(23, 8)
            txt_PlanCode_DateCode.Text = txt_QRCode.Text.Substring(7, 7)
            txt_QRCode.SelectAll()
        End If

        'arif 
        'If txt_QRCode.TextLength >= 300 And (e.KeyData = Keys.Enter Or e.KeyData = Keys.Tab) Then
        '    Combo_PlanCode.Text = ""
        '    Txt_Barcode.Text = ""
        '    txt_Ref.Text = Microsoft.VisualBasic.Right(Microsoft.VisualBasic.Left(txt_QRCode.Text, 31), 8)
        '    txt_PlanCode_DateCode.Text = txt_QRCode.Text.Substring(7, 7)
        '    txt_QRCode.SelectAll()
        'End If
    End Sub

    Private Sub Txt_Barcode_TextChanged(sender As Object, e As PreviewKeyDownEventArgs) Handles Txt_Barcode.PreviewKeyDown
        If e.KeyData = Keys.Enter Or e.KeyData = Keys.Tab Then
            If Txt_Barcode.TextLength >= 10 Then
                txt_QRCode.Text = ""
                txt_Ref.Text = Txt_Barcode.Text.Substring(6, 5)
                If Combo_PlanCode.Text = "" Then
                    MsgBox("Please Select PlanCode !")
                End If
                txt_PlanCode_DateCode.Text = Combo_PlanCode.Text & Txt_Barcode.Text.Substring(1, 5)
                Txt_Barcode.SelectAll()
            End If
        End If
    End Sub

    Private Sub Txt_PlanCode_DateCode_TextChanged(sender As Object, e As PreviewKeyDownEventArgs) Handles txt_PlanCode_DateCode.PreviewKeyDown
        If e.KeyData = Keys.Enter Or e.KeyData = Keys.Tab Then
            If txt_PlanCode_DateCode.TextLength >= 6 And txt_Ref.Text <> "" Then
                Button6_Click(sender, e)
            End If
        End If
    End Sub

    Private Sub update_combo_plantCode()
        Combo_PlanCode.Items.Clear()
        Dim i As String
        Dim a As Integer

        For a = 0 To DataGridView6.Rows.Count - 1
            'Combo_PlanCode.Items.Add(DataGridView6.Item(1, a).Value.ToString)
        Next
    End Sub
    Private Sub Combo_PlanCode_SelectedIndexChanged(sender As Object, e As EventArgs) Handles Combo_PlanCode.Click
        Dim Adap = New SqlDataAdapter("Select [Plant Code] from MasterPlantCode", Main.koneksi)
        Dim dt = New DataTable
        Try
            Adap.Fill(dt)
            Combo_PlanCode.DataSource = dt
            Combo_PlanCode.ValueMember = "Plant Code"
            Combo_PlanCode.DisplayMember = "Plant Code"
        Catch ex As Exception

        End Try

        If Txt_Barcode.TextLength >= 10 And Combo_PlanCode.Text <> "" Then
            Txt_Barcode_TextChanged(sender, e)
        End If
    End Sub

    Private Sub Button6_Click_1(sender As Object, e As EventArgs) Handles btn_add.Click
        If refI.Text <> "" And startimpactI.Text <> "" And endimpactI.Text <> "" Then
            Try
                Dim ds As New DataSet
                Dim queryCheck = "select * from QualityIssue where [References] = '" & refI.Text.ToUpper & "' and [Start Of Impact]='" & startimpactI.Text.ToUpper & "' and [End Of Impact]='" & endimpactI.Text.ToUpper & "'"
                Dim CheckAdap = New SqlDataAdapter(queryCheck, Main.koneksi)
                CheckAdap.Fill(ds)

                If ds.Tables(0).Rows.Count = 0 Then
                    Dim Query = "INSERT INTO QualityIssue ([References], [Start Of Impact], [End Of Impact]) 
                    VALUES ('" & refI.Text.ToUpper & "','" & startimpactI.Text.ToUpper & "','" & endimpactI.Text.ToUpper & "')"
                    Dim InsertBro = New SqlDataAdapter(Query, Main.koneksi)
                    If InsertBro.SelectCommand.ExecuteNonQuery() Then
                        DGV_Quality()
                        refI.Text = ""
                        startimpactI.Text = ""
                        endimpactI.Text = ""
                    End If
                Else
                    MsgBox("References in DB")
                    refI.Text = ""
                    startimpactI.Text = ""
                    endimpactI.Text = ""
                End If
            Catch ex As Exception

            End Try
        End If
    End Sub
    Function After_str(value As String, a As String) As String
        ' Get index of argument and return substring after its position.
        Dim posA As Integer = value.LastIndexOf(a)
        If posA = -1 Then
            Return ""
        End If
        Dim adjustedPosA As Integer = posA + a.Length
        If adjustedPosA >= value.Length Then
            Return ""
        End If
        Return value.Substring(adjustedPosA)
    End Function

    Private Sub TextBox2_TextChanged(sender As Object, e As PreviewKeyDownEventArgs) Handles CompToQuality.PreviewKeyDown

        Dim tampungSplit As String = ""
        If e.KeyData = Keys.Enter Or e.KeyData = 9 Then

            'jinlong request
            If Me.CompToQuality.Text.Contains("!") Then
                Dim str As String = ""
                str = Microsoft.VisualBasic.Right(Me.CompToQuality.Text, 14)
                Dim new_str As String = str.Replace("!", "")

                Dim result As String = After_str(new_str, "241")
                If result.Length > 8 Then
                    result = result.Substring(0, 8)
                End If

                result = "241" & result

                Dim q2 = "SELECT * FROM SGRAC_MES.dbo.NewScanningComponent nsc WHERE [QR code] ='" & result & "'"
                Dim dsfuji2 As New DataSet
                Dim adapt2 = New SqlDataAdapter(q2, Main.koneksi)
                adapt2.Fill(dsfuji2)

                If dsfuji2.Tables(0).Rows.Count > 0 Then
                    Dim material As String = dsfuji2.Tables(0).Rows(0).Item("Material").ToString()
                    Me.CompToQuality.Text = material
                Else
                    MsgBox("QR Code Not Found and pls Scan the Barcode !")
                    Me.CompToQuality.Text = ""
                End If
                'ElseIf Me.CompToQuality.Text.Length = 13 Then
            ElseIf Me.CompToQuality.Text.Length >= 5 Then
                Dim str As String = ""
                str = Me.CompToQuality.Text

                Dim q2 = "SELECT * FROM SGRAC_MES.dbo.NewScanningComponent nsc WHERE [Reference] ='" & str & "'"
                Dim dsfuji2 As New DataSet
                Dim adapt2 = New SqlDataAdapter(q2, Main.koneksi)
                adapt2.Fill(dsfuji2)

                If dsfuji2.Tables(0).Rows.Count > 0 Then
                    Dim material As String = dsfuji2.Tables(0).Rows(0).Item("Material").ToString()
                    Me.CompToQuality.Text = material
                Else
                    MsgBox("New scan EAN13 Not Found and pls Contact Leader!")
                    Me.CompToQuality.Text = ""
                End If

            End If

            'Cek Fuji breaker
            If header.Text.Contains("BW") Then
                Dim q = "select * from FujiBreakerCheck where [EAN13]='" & Me.CompToQuality.Text & "'"
                Dim dsfuji As New DataSet
                Dim adapt = New SqlDataAdapter(q, Main.koneksi)

                adapt.Fill(dsfuji)
                If dsfuji.Tables(0).Rows.Count > 0 Then
                    Dim a As String = "Wrong Scan!" + Chr(13) + "Please Scan QR Code"
                    MessageBox.Show(a, "Important", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Me.CompToQuality.Text = ""
                    Exit Sub
                End If
            End If

            Dim componentScanned As String = Me.CompToQuality.Text

            'santo tambah
            'If Me.CompToQuality.Text.Length = 22 Or Me.CompToQuality.Text.Length = 19 Then cek_fuji_barcode()
            If Me.CompToQuality.Text.Length = 22 Or Me.CompToQuality.Text.Length = 19 Or Me.CompToQuality.Text.Length = 48 Then cek_fuji_barcode()

            'Fuji handle repeat breaker
            If header.Text.Contains("BW") And componentScanned.Length > 20 Then
                Dim datcode_bearker As String = After_str(componentScanned, " ")
                'Dim q = "select * from printingRecord where [PP]='" & Me.PPnumberEntry.Text & "' AND [qrcodefuji] like '%" & datcode_bearker & "'"
                Dim q = "select * from printingRecord where [qrcodefuji] like '%" & datcode_bearker & "'"
                Dim dsfuji As New DataSet
                Dim adapt = New SqlDataAdapter(q, Main.koneksi)

                adapt.Fill(dsfuji)
                If dsfuji.Tables(0).Rows.Count > 0 Then
                    Dim a As String = "Breaker Already Scanned!" + Chr(13) + "Please Scan Another Breaker"
                    MessageBox.Show(a, "Important", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Me.CompToQuality.Text = ""
                    Exit Sub
                End If
            End If

            'abcde
            If Active_Quality_issue.Checked = True Then

                'mungkin salah
                'If Len(Me.CompToQuality.Text) = 22 Then
                '    SaveData.Text = Microsoft.VisualBasic.Right(Me.CompToQuality.Text, 14)
                '    tampungSplit = Microsoft.VisualBasic.Left(Me.CompToQuality.Text, 8)
                '    ComponentNo.Text = Microsoft.VisualBasic.Left(Me.CompToQuality.Text, 8)
                'End If

                'arif
                If Len(Me.CompToQuality.Text) = 24 Then tampungSplit = Microsoft.VisualBasic.Right(Microsoft.VisualBasic.Left(Me.CompToQuality.Text, 11), 10)
                If Len(Me.CompToQuality.Text) = 10 Then
                    tampungSplit = Microsoft.VisualBasic.Right(Me.CompToQuality.Text, 5)
                ElseIf Len(Me.CompToQuality.Text) >= 11 Then
                    tampungSplit = Microsoft.VisualBasic.Left(Microsoft.VisualBasic.Right(Me.CompToQuality.Text, 6), 5)
                Else
                    tampungSplit = CompToQuality.Text
                End If

                Dim sql As String = "select * from [Componentslist] where [code]='" & tampungSplit & "' or [material]='" & tampungSplit & "'"
                Dim dsSql As New DataSet
                Dim sqlAdap = New SqlDataAdapter(sql, Main.koneksi)
                sqlAdap.Fill(dsSql)
                If dsSql.Tables(0).Rows.Count > 0 Then
                    tampungSplit = dsSql.Tables(0).Rows(0).Item("Material").ToString
                Else
                    MsgBox("Wrong Components", vbExclamation, "The component is not part of the product")
                End If

                'santo
                If Convert.ToDecimal(CounterItems.Text) = Convert.ToDecimal(LabelQuantityitem.Text) Then
                    MsgBox("All Breakers have been adapted !")
                    CompToQuality.Text = ""
                    Exit Sub
                End If

                Try
                    'Dim queryCheck = "select * from QualityIssue where [references]='" & CompToQuality.Text & "'"
                    Dim queryCheck = "select * from QualityIssue where [references]='" & tampungSplit & "'"
                    Dim ds As New DataSet
                    Dim CheckAdap = New SqlDataAdapter(queryCheck, Main.koneksi)

                    CheckAdap.Fill(ds)

                    If ds.Tables(0).Rows.Count >= 1 Then
                        'Dim PlantCode = InputBox("Please enter the Plant Code + Date Code :" & vbLf & vbLf & "Example: AF19514", "Enter Plant Code & Date Code")
                        Dim PlantCode
                        'santo

                        Input_BoX.ShowDialog()
                        PlantCode = Input_BoX.txt_PlanCode_DateCode.Text

                        'Dim Query = "INSERT INTO [dbo].[QualityIssueRecord] ([PP],[References],[PlantCodeDateCode],[Remark]) VALUES ('" & PPnumberEntry.Text & "','" & tampungSplit.ToUpper & "','" & PlantCode.ToString.ToUpper & "','" & checkQualityIssue(tampungSplit, PlantCode) & "')"
                        'Dim InsertBro = New SqlDataAdapter(Query, Main.koneksi)
                        'InsertBro.SelectCommand.ExecuteNonQuery()

                        If checkQualityIssue(tampungSplit, PlantCode) <> "OK" Then

                            MessageBox.Show("Product Impected based on LA Info!" & vbCrLf & "Please Liaise with CS & Q", "Attention !",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Stop)

                            CompToQuality.Text = ""
                            'CompToQuality.Select()
                            CompToQuality.Enabled = False
                            PPnumberEntry.Select()
                            PPnumberEntry.SelectAll()
                            Exit Sub

                        Else
                            ComponentNo.Text = CompToQuality.Text
                            CompToQuality.Text = ""
                            CompToQuality.Select()
                            'Lbl_Result.Text = "Result"
                            'lbl_PlantCode.Text = "Plant :"
                        End If
                    Else
                        ComponentNo.Text = CompToQuality.Text
                        CompToQuality.Text = ""
                        CompToQuality.Select()
                    End If
                Catch ex As Exception

                End Try
            Else
                ComponentNo.Text = CompToQuality.Text
                CompToQuality.Text = ""
                CompToQuality.Select()
            End If
        End If
    End Sub

    Private Sub Simpan_QualityIssueRecord()
        Dim pp_No As String

        'If Active_Quality_issue.Checked = True Then
        '    pp_No = PPnumberEntry.Text
        'Else
        pp_No = ""
        'End If

        Dim Query = "INSERT INTO [dbo].[QualityIssueRecord]
           ([PP]
           ,[References]
           ,[PlantCodeDateCode]
           ,[Remark]
           ,[Datetime])
            VALUES ('" & pp_No & "','" & txt_Ref.Text.ToUpper & "','" & txt_PlanCode_DateCode.Text.ToUpper & "','" & Lbl_Result.Text & "'," & "getdate() )"

        Dim InsertBro = New SqlDataAdapter(Query, Main.koneksi)
        InsertBro.SelectCommand.ExecuteNonQuery()
    End Sub

    Private Sub DGV_QualityIssue_refresh()
        ' Dim con As New connectionstring='your connection string

        Dim cmd As New SqlCommand("SELECT * FROM [SGRAC_MES].[dbo].[QualityIssueRecord] order by [datetime] DESC", Main.koneksi)

        Dim adapter As New SqlDataAdapter(cmd)

        Dim table As New DataTable

        adapter.Fill(table)

        DGV_QI_Trace.DataSource = table
    End Sub

    Private Sub Btn_trace_QI_Click(sender As Object, e As EventArgs) Handles btn_trace_QI.Click

        If txt_PP_QI_Trace.TextLength <= 5 Then txt_PP_QI_Trace.Text = "***"

        Dim cmd As New SqlCommand("SELECT * FROM [SGRAC_MES].[dbo].[QualityIssueRecord]
        where [pp] = '" & txt_PP_QI_Trace.Text & " ' OR [References]='" & txt_Ref_QI_Trace.Text & "'", Main.koneksi)

        Dim adapter As New SqlDataAdapter(cmd)

        Dim table As New DataTable

        adapter.Fill(table)

        DGV_QI_Trace.DataSource = table
    End Sub

    Private Sub Txt_Ref_QI_Trace_TextChanged(sender As Object, e As PreviewKeyDownEventArgs) Handles txt_Ref_QI_Trace.PreviewKeyDown
        If e.KeyData = Keys.Enter Or e.KeyData = Keys.Tab Then
            Btn_trace_QI_Click(sender, e)
        End If
    End Sub

    Private Sub Txt_PP_QI_Trace_TextChanged(sender As Object, e As PreviewKeyDownEventArgs) Handles txt_PP_QI_Trace.PreviewKeyDown
        If e.KeyData = Keys.Enter Or e.KeyData = Keys.Tab Then
            Btn_trace_QI_Click(sender, e)
        End If
    End Sub

    Private Sub Btn_delete_Click(sender As Object, e As EventArgs) Handles btn_delete.Click
        Dim i = DataGridView5.CurrentRow.Index
        Dim SelectReferences As String = DataGridView5.Item(1, i).Value.ToString
        Dim SelectStartDC As String = DataGridView5.Item(2, i).Value.ToString
        Dim SelectEndDC As String = DataGridView5.Item(3, i).Value.ToString

        Dim result As DialogResult = MessageBox.Show("Are You sure going to delete ?", "Info", MessageBoxButtons.YesNo)
        If result = DialogResult.No Then
            Exit Sub
        End If

        Dim Query = "DELETE FROM [dbo].[QualityIssue]
        WHERE [References]='" & SelectReferences & "' AND [Start Of Impact] = '" & SelectStartDC & "' AND [End Of Impact] = '" & SelectEndDC & "'"

        Dim InsertBro = New SqlDataAdapter(Query, Main.koneksi)
        InsertBro.SelectCommand.ExecuteNonQuery()

        DGV_Quality()

    End Sub

    Private Sub Button7_Click_1(sender As Object, e As EventArgs) Handles Button7.Click
        If plantcodeI.Text <> "" And plantcodenameI.Text <> "" Then
            Try
                Dim ds As New DataSet
                Dim queryCheck = "select * from MasterPlantCode where [Plant Code] = '" & plantcodeI.Text.ToUpper & "'"
                Dim CheckAdap = New SqlDataAdapter(queryCheck, Main.koneksi)
                CheckAdap.Fill(ds)

                If ds.Tables(0).Rows.Count = 0 Then
                    Dim Query = "INSERT INTO MasterPlantCode ([Plant Code], [Name Plant Code]) 
                    VALUES ('" & plantcodeI.Text.ToUpper & "','" & plantcodenameI.Text.ToUpper & "')"
                    Dim InsertBro = New SqlDataAdapter(Query, Main.koneksi)
                    If InsertBro.SelectCommand.ExecuteNonQuery() Then
                        DGV_MasterPlantCode()
                        plantcodeI.Text = ""
                        plantcodenameI.Text = ""
                    End If
                Else
                    MsgBox("References in DB")
                    plantcodeI.Text = ""
                    plantcodenameI.Text = ""
                End If
            Catch ex As Exception

            End Try
        End If
    End Sub

    Private Sub Button6_Click_2(sender As Object, e As EventArgs) Handles Button6.Click
        Dim i = DataGridView6.CurrentRow.Index
        Dim SelectPC As String = DataGridView6.Item(1, i).Value.ToString

        Dim result As DialogResult = MessageBox.Show("Are You sure going to delete ?", "Info", MessageBoxButtons.YesNo)
        If result = DialogResult.No Then
            Exit Sub
        End If

        Dim Query = "DELETE FROM [dbo].[MasterPlantCode] WHERE [Plant Code]='" & SelectPC & "'"
        Dim DeleteBro = New SqlDataAdapter(Query, Main.koneksi)
        DeleteBro.SelectCommand.ExecuteNonQuery()

        DGV_MasterPlantCode()
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click

        Button8.Text = "Please Wait..."
        Button8.Enabled = False

        Dim Dialog As New SaveFileDialog
        Dialog.Filter = "Microsoft excel 97-2003|*.xls"
        If (Dialog.ShowDialog = DialogResult.OK) Then
            Dim xlApp As New Excel.Application
            Dim xlWorkbook As Excel.Workbook = xlApp.Workbooks.Add()
            Dim xlWorksheet As Excel.Worksheet = CType(xlWorkbook.Sheets("sheet1"), Excel.Worksheet)

            xlWorksheet.Cells(1, 1) = "References"
            xlWorksheet.Cells(1, 2) = "Start Of Impact (Plant Code + Date Code ) Ex : AF19511"
            xlWorksheet.Cells(1, 3) = "End Of Impact (Plant Code + Date Code ) Ex : AF20111"
            xlWorksheet.Cells(2, 1) = "paste the data from this line"
            xlApp = New Microsoft.Office.Interop.Excel.Application
            xlWorksheet = xlWorkbook.Sheets("Sheet1")
            xlWorksheet.SaveAs(Dialog.FileName)
            If System.IO.File.Exists(Dialog.FileName) Then
                MsgBox("Export Success")
            End If
            xlWorkbook.Close()
            xlApp.Quit()
        End If

        Button8.Text = "Export Tamplate"
        Button8.Enabled = True
    End Sub

    Private Sub Btn_export_QI_Trace_Click(sender As Object, e As EventArgs) Handles btn_export_QI_Trace.Click
        Try
            btn_export_QI_Trace.Text = "Please Wait..."
            btn_export_QI_Trace.Enabled = False

            SaveFileDialog1.Filter = "Excel Document (*.xlsx)|*.xlsx"
            If SaveFileDialog1.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
                Dim xlApp As Microsoft.Office.Interop.Excel.Application
                Dim xlWorkBook As Microsoft.Office.Interop.Excel.Workbook
                Dim xlWorkSheet As Microsoft.Office.Interop.Excel.Worksheet
                Dim misValue As Object = System.Reflection.Missing.Value
                Dim i As Integer
                Dim j As Integer

                xlApp = New Microsoft.Office.Interop.Excel.Application
                xlWorkBook = xlApp.Workbooks.Add(misValue)
                xlWorkSheet = xlWorkBook.Sheets("sheet1")

                For i = 0 To DGV_QI_Trace.RowCount - 2
                    For j = 0 To DGV_QI_Trace.ColumnCount - 1
                        For k As Integer = 1 To DGV_QI_Trace.Columns.Count
                            xlWorkSheet.Cells(1, k) = DGV_QI_Trace.Columns(k - 1).HeaderText
                            xlWorkSheet.Cells(i + 2, j + 1) = DGV_QI_Trace(j, i).Value.ToString()
                        Next
                    Next
                Next

                xlWorkSheet.SaveAs(SaveFileDialog1.FileName)
                xlWorkBook.Close()
                xlApp.Quit()

                releaseObject_export_excel(xlApp)
                releaseObject_export_excel(xlWorkBook)
                releaseObject_export_excel(xlWorkSheet)

                MsgBox("Successfully saved" & vbCrLf & "File are saved at : " & SaveFileDialog1.FileName, MsgBoxStyle.Information, "Information")

                btn_export_QI_Trace.Text = "Export"
                btn_export_QI_Trace.Enabled = True
            End If
        Catch ex As Exception
            MessageBox.Show("Failed to save !!!", "Error Message", MessageBoxButtons.OK, MessageBoxIcon.Error)
            btn_export_QI_Trace.Text = "Export"
            btn_export_QI_Trace.Enabled = True
            Return
        End Try
    End Sub

    Private Sub releaseObject_export_excel(ByVal obj As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
        Finally
            GC.Collect()
        End Try
    End Sub

    Private Sub Btn_export_Tracecibility_Click(sender As Object, e As EventArgs) Handles btn_export_Tracecibility.Click
        Try
            btn_export_Tracecibility.Text = "Please Wait..."
            btn_export_Tracecibility.Enabled = False

            SaveFileDialog1.Filter = "Excel Document (*.xlsx)|*.xlsx"
            If SaveFileDialog1.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
                Dim xlApp As Microsoft.Office.Interop.Excel.Application
                Dim xlWorkBook As Microsoft.Office.Interop.Excel.Workbook
                Dim xlWorkSheet As Microsoft.Office.Interop.Excel.Worksheet
                Dim misValue As Object = System.Reflection.Missing.Value
                Dim i As Integer
                Dim j As Integer

                xlApp = New Microsoft.Office.Interop.Excel.Application
                xlWorkBook = xlApp.Workbooks.Add(misValue)

                xlWorkSheet = xlWorkBook.Worksheets.Add()
                xlWorkSheet.Name = "Components Scanned"
                xlWorkSheet = xlWorkBook.Worksheets.Add()
                xlWorkSheet.Name = "Technicians on the Workstation"
                xlWorkSheet = xlWorkBook.Worksheets.Add()
                xlWorkSheet.Name = "PP Info"

                xlWorkSheet = xlWorkBook.Sheets("PP Info")

                For i = 0 To DataGridView2.RowCount - 2
                    For j = 0 To DataGridView2.ColumnCount - 1
                        For k As Integer = 1 To DataGridView2.Columns.Count
                            xlWorkSheet.Cells(1, k) = DataGridView2.Columns(k - 1).HeaderText
                            xlWorkSheet.Cells(i + 2, j + 1) = DataGridView2(j, i).Value.ToString()
                        Next
                    Next
                Next

                xlWorkSheet = xlWorkBook.Sheets("Technicians on the Workstation")

                For i = 0 To DataGridView3.RowCount - 2
                    For j = 0 To DataGridView3.ColumnCount - 1
                        For k As Integer = 1 To DataGridView3.Columns.Count
                            xlWorkSheet.Cells(1, k) = DataGridView3.Columns(k - 1).HeaderText
                            xlWorkSheet.Cells(i + 2, j + 1) = DataGridView3(j, i).Value.ToString()
                        Next
                    Next
                Next

                xlWorkSheet = xlWorkBook.Sheets("Components Scanned")

                For i = 0 To DataGridView4.RowCount - 2
                    For j = 0 To DataGridView4.ColumnCount - 1
                        For k As Integer = 1 To DataGridView4.Columns.Count
                            xlWorkSheet.Cells(1, k) = DataGridView4.Columns(k - 1).HeaderText
                            xlWorkSheet.Cells(i + 2, j + 1) = DataGridView4(j, i).Value.ToString()
                        Next
                    Next
                Next


                xlWorkSheet.SaveAs(SaveFileDialog1.FileName)
                xlWorkBook.Close()
                xlApp.Quit()

                releaseObject_export_excel(xlApp)
                releaseObject_export_excel(xlWorkBook)
                releaseObject_export_excel(xlWorkSheet)

                MsgBox("Successfully saved" & vbCrLf & "File are saved at : " & SaveFileDialog1.FileName, MsgBoxStyle.Information, "Information")

                btn_export_Tracecibility.Text = "Export"
                btn_export_Tracecibility.Enabled = True
            End If
        Catch ex As Exception
            MessageBox.Show("Failed to save !!!", "Error Message", MessageBoxButtons.OK, MessageBoxIcon.Error)
            btn_export_Tracecibility.Text = "Export"
            btn_export_Tracecibility.Enabled = True
            Return
        End Try
    End Sub

    Private Sub Btn_QI_Search_Click(sender As Object, e As EventArgs) Handles btn_QI_Search.Click
        If TXT_Ref_Search_QI.Text <> "" Then
            Dim sqlSearch = "select * from QualityIssue where [References]='" & TXT_Ref_Search_QI.Text & "'"
            Dim ds As New DataSet
            Dim CheckAdap = New SqlDataAdapter(sqlSearch, Main.koneksi)
            CheckAdap.Fill(ds)

            If ds.Tables(0).Rows.Count = 0 Then
                MsgBox("No Data Found")
            Else
                DataGridView5.Rows.Clear()
                DataGridView5.ColumnCount = 4
                DataGridView5.Columns(0).Name = "NO"
                DataGridView5.Columns(1).Name = "References"
                DataGridView5.Columns(2).Name = "Start Of Impact"
                DataGridView5.Columns(3).Name = "End  Of Impact"
                For r = 0 To ds.Tables(0).Rows.Count - 1
                    Dim row As String() = New String() {(r + 1).ToString(), ds.Tables(0).Rows(r).Item("References").ToString(), ds.Tables(0).Rows(r).Item("Start Of Impact").ToString(), ds.Tables(0).Rows(r).Item("End Of Impact").ToString()}
                    DataGridView5.Rows.Add(row)
                Next
            End If
        Else
            DGV_Quality()
        End If
    End Sub

    Private Sub TXT_Ref_Search_QI_TextChanged(sender As Object, e As PreviewKeyDownEventArgs) Handles TXT_Ref_Search_QI.PreviewKeyDown
        If e.KeyData = Keys.Enter Or e.KeyData = Keys.Tab Then
            If TXT_Ref_Search_QI.Text <> "" Then
                Dim sqlSearch = "select * from QualityIssue where [References]='" & TXT_Ref_Search_QI.Text & "'"
                Dim ds As New DataSet
                Dim CheckAdap = New SqlDataAdapter(sqlSearch, Main.koneksi)
                CheckAdap.Fill(ds)

                If ds.Tables(0).Rows.Count = 0 Then
                    MsgBox("No Data Found")
                Else
                    DataGridView5.Rows.Clear()
                    DataGridView5.ColumnCount = 4
                    DataGridView5.Columns(0).Name = "NO"
                    DataGridView5.Columns(1).Name = "References"
                    DataGridView5.Columns(2).Name = "Start Of Impact"
                    DataGridView5.Columns(3).Name = "End  Of Impact"
                    For r = 0 To ds.Tables(0).Rows.Count - 1
                        Dim row As String() = New String() {(r + 1).ToString(), ds.Tables(0).Rows(r).Item("References").ToString(), ds.Tables(0).Rows(r).Item("Start Of Impact").ToString(), ds.Tables(0).Rows(r).Item("End Of Impact").ToString()}
                        DataGridView5.Rows.Add(row)
                    Next
                End If
            Else
                DGV_Quality()
            End If
        End If
    End Sub

    Private Sub DeleteRow_Click(sender As Object, e As EventArgs) Handles deleteRow.Click

        Dim i = DGV_QI_Trace.CurrentRow.Index
        Dim SelectPC As String = DGV_QI_Trace.Item(0, i).Value.ToString

        Dim result As DialogResult = MessageBox.Show("Are You sure going to delete ?", "Info", MessageBoxButtons.YesNo)
        If result = DialogResult.No Then
            Exit Sub
        End If

        Dim Query = "DELETE FROM [dbo].[QualityIssueRecord] WHERE [ID]='" & SelectPC & "'"
        Dim DeleteBro = New SqlDataAdapter(Query, Main.koneksi)
        DeleteBro.SelectCommand.ExecuteNonQuery()

        DGV_QualityIssue_refresh()
    End Sub

    Private Sub Active_Quality_issue_CheckedChanged(sender As Object, e As EventArgs) Handles Active_Quality_issue.CheckedChanged
        Try
            Using writer As New StreamWriter("chkbox.txt", False)
                writer.Write(Active_Quality_issue.Checked.ToString)
            End Using
        Catch ex As Exception

        End Try

    End Sub

    Private Sub baca_chkbox_last()
        Dim line As String
        Try
            Using reader As New StreamReader("chkbox.txt")
                line = reader.ReadLine()
                'MsgBox(line.ToString)
                If line = "True" Then
                    Active_Quality_issue.Checked = True
                Else
                    Active_Quality_issue.Checked = False
                End If
            End Using
        Catch ex As Exception

        End Try
    End Sub

    Private Sub masterfuji_Click(sender As Object, e As EventArgs) Handles masterfuji.Click
        OpenFileDialog1.InitialDirectory = My.Computer.FileSystem.SpecialDirectories.MyDocuments
        If OpenFileDialog1.ShowDialog(Me) = System.Windows.Forms.DialogResult.OK Then
            Dim xlApp As New Microsoft.Office.Interop.Excel.Application
            Dim xlWorkBook As Microsoft.Office.Interop.Excel.Workbook = xlApp.Workbooks.Open(OpenFileDialog1.FileName)
            Dim SheetName As String = xlWorkBook.Worksheets(1).Name.ToString
            Dim SheetName2 As String = xlWorkBook.Worksheets(2).Name.ToString
            Dim excelpath As String = OpenFileDialog1.FileName
            Dim koneksiExcel As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & excelpath & ";Extended Properties='Excel 8.0;HDR=No;IMEX=1;'"
            oleCon = New OleDbConnection(koneksiExcel)
            oleCon.Open()

            Dim queryExcel As String = "select * from [" & SheetName & "$]"
            Dim cmd As OleDbCommand = New OleDbCommand(queryExcel, oleCon)
            Dim rd As OleDbDataReader

            Dim queryExcel2 As String = "select * from [" & SheetName2 & "$]"
            Dim cmd2 As OleDbCommand = New OleDbCommand(queryExcel2, oleCon)
            Dim rd2 As OleDbDataReader

            Call koneksi_db()

            Dim deleteReset As New SqlCommand("Delete from [dbo].[MasterFuji]", Main.koneksi)

            Try
                deleteReset.ExecuteNonQuery()
            Catch ex As Exception
                MsgBox("Delete DB Fail " & ex.Message)
            End Try

            Using bulkCopy As SqlBulkCopy = New SqlBulkCopy(Main.koneksi)
                bulkCopy.DestinationTableName = "dbo.MasterFuji"
                Try
                    rd = cmd.ExecuteReader
                    bulkCopy.WriteToServer(rd)

                    rd.Close()
                    MsgBox("Upload Master Data FUJI Success !")
                Catch ex As Exception
                    MsgBox("Upload Master Fuji Fail" & ex.Message)
                End Try
            End Using

            Dim deleteReset1 As New SqlCommand("Delete from [dbo].[MasterFujiLabelling]", Main.koneksi)

            Try
                deleteReset1.ExecuteNonQuery()
            Catch ex As Exception
                MsgBox("Delete DB Fail " & ex.Message)
            End Try

            Using bulkCopy2 As SqlBulkCopy = New SqlBulkCopy(Main.koneksi)
                bulkCopy2.DestinationTableName = "dbo.MasterFujiLabelling"
                Try
                    rd2 = cmd2.ExecuteReader
                    bulkCopy2.WriteToServer(rd2)

                    rd2.Close()
                    MsgBox("Upload Master Data FUJI Labelling Success !")
                Catch ex As Exception
                    MsgBox("Upload Master Fuji Fail" & ex.Message)
                End Try
            End Using
        End If
    End Sub

    Private Sub Refresh_DGV_Fuji()
        DataGridView7.Rows.Clear()
        'DataGridView8.Rows.Clear()
        Dim sql As String = "select [Order], [RefFuji],[Check Components],[QRCodeFuji]  from ComponentsFuji where [Order]='" & Me.PPFujiEntry.Text & "' and Workstation='" & Me.workstationFuji.Text & "' and [2ndScan]= 0 "
        Dim ds As New DataSet
        adapter = New SqlDataAdapter(sql, Main.koneksi)
        adapter.Fill(ds)
        If ds.Tables(0).Rows.Count > 0 Then
            DataGridView7.ColumnCount = 4
            DataGridView7.Columns(0).Name = "Order"
            DataGridView7.Columns(1).Name = "FCS Refrence"
            DataGridView7.Columns(2).Name = "QR Code"
            DataGridView7.Columns(3).Name = "Check Scan"
            For r = 0 To ds.Tables(0).Rows.Count - 1
                Dim row As String() = New String() {ds.Tables(0).Rows(r).Item("Order").ToString(), ds.Tables(0).Rows(r).Item("RefFuji").ToString(), ds.Tables(0).Rows(r).Item("QRCodeFuji").ToString(), ds.Tables(0).Rows(r).Item("Check Components").ToString()}
                DataGridView7.Rows.Add(row)
                If ds.Tables(0).Rows(r).Item("Check Components").ToString() = 1 Then
                    DataGridView7.Rows(r).Cells(3).Style.BackColor = Color.SkyBlue
                Else
                    DataGridView7.Rows(r).Cells(3).Style.BackColor = Color.White
                End If
            Next
        End If
    End Sub

    Private Sub PPFujiEntry_KeyPress(sender As Object, e As PreviewKeyDownEventArgs) Handles PPFujiEntry.PreviewKeyDown 'sudah sesuai dengan VBA
        If Len(Me.PPFujiEntry.Text) = 12 Then Me.PPFujiEntry.Text = Microsoft.VisualBasic.Right(Me.PPFujiEntry.Text, 11)

        If (e.KeyData = Keys.Tab Or e.KeyData = Keys.Enter) And Len(Me.PPFujiEntry.Text) >= 11 Then
            If String.IsNullOrEmpty(workstationFuji.Text) = False And String.IsNullOrEmpty(technicianNameFuji.Text) = False Then
                If IsNumeric(PPFujiEntry.Text) Then
                    Dim sqlCekFuji As String = "select * from MasterFuji, MasterFujiLabelling, openOrders where MasterFuji.FCSRef=MasterFujiLabelling.FCSRef and openOrders.PeggedReqt=MasterFuji.FCSRef and openOrders.[Order]='" & Me.PPFujiEntry.Text & "'"
                    Dim dsCekFuji As New DataSet
                    adapter = New SqlDataAdapter(sqlCekFuji, Main.koneksi)
                    adapter.Fill(dsCekFuji)
                    If (dsCekFuji.Tables(0).Rows.Count > 0) Then
                        Dim sqlCekForInsert As String = "select * from openOrders, ComponentsFuji where openOrders.[Order]='" & Me.PPFujiEntry.Text & "' and ComponentsFuji.RefFuji=openOrders.PeggedReqt and ComponentsFuji.Workstation='" & Me.workstationFuji.Text & "' and ComponentsFuji.[2ndScan]=0"
                        Dim dsCekForInsert As New DataSet
                        adapter = New SqlDataAdapter(sqlCekForInsert, Main.koneksi)
                        adapter.Fill(dsCekForInsert)
                        If (dsCekForInsert.Tables(0).Rows.Count > 0) Then
                            Refresh_DGV_Fuji()
                            ScanLabel.Enabled = True
                            ScanLabel2.Enabled = False
                            ScanLabel.Select()
                            'If dsCekForInsert.Tables(0).Rows(0).Item("Check Components").ToString = 1 Then
                            '    Refresh_DGV_Fuji()
                            '    Refresh_DGV_Fuji_2()
                            '    ScanLabel.Enabled = False
                            '    ScanLabel2.Select()
                            'Else
                            '    Refresh_DGV_Fuji()
                            '    ScanLabel.Enabled = True
                            '    ScanLabel.Select()
                            '    ScanLabel2.Enabled = False
                            'End If
                        Else
                            MessageBox.Show("No Data Found")
                            'Dim queryInsert = "insert into [ComponentsFuji]([order],[RefFuji],[WorkStation],[Check Components], 
                            '[2ndScan]) values ('" & Me.PPFujiEntry.Text & "',(select DISTINCT PeggedReqt from openOrders 
                            'where [order]='" & Me.PPFujiEntry.Text & "'),'" & Me.workstationFuji.Text & "',0,0)"
                            'adapter = New SqlDataAdapter(queryInsert, Main.koneksi)
                            'adapter.SelectCommand.ExecuteNonQuery()
                            'Refresh_DGV_Fuji()
                            'ScanLabel.Enabled = True
                            'ScanLabel.Select()
                            'ScanLabel2.Enabled = False
                        End If
                        'Keluarkan semua data dibawah ini
                        FCSRef.Text = dsCekFuji.Tables(0).Rows(0).Item("FCSRef").ToString
                        Me.dateCodeFuji.Text = dateCode2()

                        'this textbox for rotaryhandle
                        QRForRotaryHandle.Text = dsCekFuji.Tables(0).Rows(0).Item("QRRotaryHandleNumber").ToString
                        Rotary1.Text = dsCekFuji.Tables(0).Rows(0).Item("From1").ToString
                        Rotary2.Text = dsCekFuji.Tables(0).Rows(0).Item("From2").ToString
                        Rotary3.Text = dsCekFuji.Tables(0).Rows(0).Item("FrameNumber").ToString
                        Rotary5.Text = dsCekFuji.Tables(0).Rows(0).Item("BC1").ToString
                        Rotary6.Text = dsCekFuji.Tables(0).Rows(0).Item("BC2").ToString
                        Rotary7.Text = dsCekFuji.Tables(0).Rows(0).Item("BC3").ToString
                        Rotary4.Text = dsCekFuji.Tables(0).Rows(0).Item("RatedCurrent").ToString
                        LotNumberRotaryHandle.Text = dsCekFuji.Tables(0).Rows(0).Item("RotaryHandleLabelNumber").ToString
                        IdxLabelRotaryHandle.Text = "-"

                        'this textbox for Front Cover Label
                        QRForCoverLabel.Text = dsCekFuji.Tables(0).Rows(0).Item("QRFrontLabelNumber").ToString
                        Front1.Text = dsCekFuji.Tables(0).Rows(0).Item("From1").ToString
                        Front2.Text = dsCekFuji.Tables(0).Rows(0).Item("From2").ToString
                        Front3.Text = dsCekFuji.Tables(0).Rows(0).Item("FrameNumber").ToString
                        Front5.Text = dsCekFuji.Tables(0).Rows(0).Item("BC1").ToString
                        Front6.Text = dsCekFuji.Tables(0).Rows(0).Item("BC2").ToString
                        Front7.Text = dsCekFuji.Tables(0).Rows(0).Item("BC3").ToString
                        Front4.Text = dsCekFuji.Tables(0).Rows(0).Item("RatedCurrent").ToString
                        LotNumberCoverLabel.Text = dsCekFuji.Tables(0).Rows(0).Item("FrontLabelNumber").ToString
                        IdxLabelCoverLabel.Text = "-"
                        Dim range_label As String = dsCekFuji.Tables(0).Rows(0).Item("FigureNumber").ToString
                        If Convert.ToDecimal(range_label) <= 56 Or (Convert.ToDecimal(range_label) >= 73 And Convert.ToDecimal(range_label) <= 88) Then
                            FujiLabelType.Text = "Short"
                        Else
                            FujiLabelType.Text = "Long"
                        End If

                        'this textbox for Carton Label
                        Carton1.Text = dsCekFuji.Tables(0).Rows(0).Item("From1").ToString
                        Carton2.Text = dsCekFuji.Tables(0).Rows(0).Item("From2").ToString
                        Carton32.Text = dsCekFuji.Tables(0).Rows(0).Item("FrameNumber")
                        Carton32.Text = Carton32.Text.ToString.Substring(Carton32.Text.Length - 2)
                        Carton4.Text = dsCekFuji.Tables(0).Rows(0).Item("RatedCurrent").ToString
                        'baca Sq00
                        Dim sql_box As String = "SELECT* FROM PPList WHERE [Order] ='" & Me.PPFujiEntry.Text & "'"
                        Dim ds_box As New DataSet
                        Dim adapter_box = New SqlDataAdapter(sql_box, Main.koneksi)
                        adapter_box.Fill(ds_box)



                        FujiSO.Text = ds_box.Tables(0).Rows(0).Item("SO no").ToString
                        FujiLineNo.Text = ds_box.Tables(0).Rows(0).Item("item").ToString
                        fujiQty.Text = ds_box.Tables(0).Rows(0).Item("Item quantity").ToString

                        GroupingBox.Text = dsCekFuji.Tables(0).Rows(0).Item("GroupingBox").ToString
                        PO_Group_BOX.Text = ds_box.Tables(0).Rows(0).Item("Purchase order number").ToString
                        'PO_Group_BOX.Text = ds_box.Tables(0).Rows(0).Item("Order").ToString


                        Dim sqlCounterItemsFuji As String = "select * from ComponentsFuji where 
                        ComponentsFuji.[Order]='" & Me.PPFujiEntry.Text & "' and ComponentsFuji.Workstation='" & Me.workstationFuji.Text & "' 
                        and ComponentsFuji.[Check Components]=1 and ComponentsFuji.[2ndScan]= 0 "
                        Dim ds3 As New DataSet
                        adapter3 = New SqlDataAdapter(sqlCounterItemsFuji, Main.koneksi)
                        adapter3.Fill(ds3)
                        CounterItemsFuji.Text = ds3.Tables(0).Rows.Count

                        ScanLabel.Text = ""
                    Else
                        MessageBox.Show("This PP not Fuji")
                    End If
                Else
                    MsgBox("Sorry The PP must be number")
                    Me.PPFujiEntry.Text = ""
                End If
            Else
                MsgBox("Sorry you Must select workstations and technician first")
            End If
        End If


    End Sub

    Private Sub ScanLabel_KeyPress(sender As Object, e As PreviewKeyDownEventArgs) Handles ScanLabel.PreviewKeyDown
        'edit santo
        If (e.KeyData = Keys.Tab Or e.KeyData = Keys.Enter) And (Len(Me.ScanLabel.Text) >= 33 Or Len(Me.ScanLabel.Text) >= 29) Then
            'If (e.KeyData = Keys.Tab Or e.KeyData = Keys.Enter) And Len(Me.ScanLabel.Text) >= 33 Then
            'ScanLabel.Text = ScanLabel.Text.Replace(Keys.Tab, "")

            Dim SplitData = Microsoft.VisualBasic.Left(Me.ScanLabel.Text, 14)
            Dim sql As String = "select ComponentsFuji.[Order], ComponentsFuji.[RefFuji],ComponentsFuji.[Check Components],
            MasterFujiLabelling.QRFrontLabelNumber,MasterFujiLabelling.QRRotaryHandleNumber,MasterFujiLabelling.TripUnitLabelNumber from ComponentsFuji, MasterFujiLabelling where 
            ComponentsFuji.[Order]='" & Me.PPFujiEntry.Text & "' and ComponentsFuji.Workstation='" & Me.workstationFuji.Text & "' 
            and ComponentsFuji.QRCodeFuji='" & Me.ScanLabel.Text & "' and MasterFujiLabelling.FCSRef = ComponentsFuji.RefFuji 
            and ComponentsFuji.[2ndScan]= 0 "
            Dim ds As New DataSet
            adapter = New SqlDataAdapter(sql, Main.koneksi)
            adapter.Fill(ds)
            If ds.Tables(0).Rows.Count > 0 Then
                If ds.Tables(0).Rows(0).Item("Check Components").ToString() = 0 Then
                    Dim queryUpdate = "UPDATE [ComponentsFuji] SET [Check Components] = 1 where 
                    [Order]='" & Me.PPFujiEntry.Text & "' 
                    and Workstation='" & Me.workstationFuji.Text & "' 
                    and QRCodeFuji='" & Me.ScanLabel.Text & "' and [2ndScan]= 0"
                    adapter = New SqlDataAdapter(queryUpdate, Main.koneksi)
                    If adapter.SelectCommand.ExecuteNonQuery().ToString() = 1 Then
                        'Refresh_DGV_Fuji()

                        ''Baca scan data
                        Dim sql2 As String = "SELECT * FROM SGRAC_MES.dbo.MasterFuji WHERE FCSRef ='" & FCSRef.Text & "'"
                        Dim ds2 As New DataSet
                        adapter2 = New SqlDataAdapter(sql2, Main.koneksi)
                        adapter2.Fill(ds2)

                        Dim scan_rotary As String = ds2.Tables(0).Rows(0).Item("ScanRotaryHandleLabel").ToString()
                        Dim scan_front As String = ds2.Tables(0).Rows(0).Item("ScanFrontLabel").ToString()
                        Dim scan_trip As String = ds2.Tables(0).Rows(0).Item("ScanTripUnitLabel").ToString()

                        Dim maxScan As Integer

                        'MsgBox(scan_rotary & " " & scan_front & " " & scan_trip & " ")

                        If chk_scan_tripUnit.Checked = True Then
                            maxScan = 3
                        Else
                            maxScan = 2
                        End If


                        For r = 1 To maxScan
                            If r = 1 And scan_rotary = "1" Then
                                Dim NamaLabel As String = "Rotary Handle Label"
                                Dim queryInsert = "insert into [ComponentsFuji]([order],[RefFuji],[WorkStation],[Check Components], 
                                [2ndScan],[2ndScanName],[2ndQrCode],[QRCodeFuji]) values ('" & Me.PPFujiEntry.Text & "','" & SplitData & "',
                                '" & Me.workstationFuji.Text & "',0,1,'" & NamaLabel & "',
                                '" & ds.Tables(0).Rows(0).Item("QRRotaryHandleNumber").ToString() & "','" & Me.ScanLabel.Text & "')"
                                adapter = New SqlDataAdapter(queryInsert, Main.koneksi)
                                adapter.SelectCommand.ExecuteNonQuery()
                            ElseIf r = 2 And scan_front = "1" Then
                                Dim NamaLabel As String = "Front Cover Label"
                                Dim queryInsert = "insert into [ComponentsFuji]([order],[RefFuji],[WorkStation],[Check Components], 
                                [2ndScan],[2ndScanName],[2ndQrCode],[QRCodeFuji]) values ('" & Me.PPFujiEntry.Text & "','" & SplitData & "',
                                '" & Me.workstationFuji.Text & "',0,1,'" & NamaLabel & "',
                                '" & ds.Tables(0).Rows(0).Item("QRFrontLabelNumber").ToString() & "','" & Me.ScanLabel.Text & "')"
                                adapter = New SqlDataAdapter(queryInsert, Main.koneksi)
                                adapter.SelectCommand.ExecuteNonQuery()
                                'ElseIf chk_scan_tripUnit.Checked = True And scan_trip = "1" Then
                            ElseIf r = 3 And chk_scan_tripUnit.Checked = True Then
                                Dim NamaLabel As String = "Trip Unit Label"
                                Dim queryInsert = "insert into [ComponentsFuji]([order],[RefFuji],[WorkStation],[Check Components], 
                                [2ndScan],[2ndScanName],[2ndQrCode],[QRCodeFuji]) values ('" & Me.PPFujiEntry.Text & "','" & SplitData & "',
                                '" & Me.workstationFuji.Text & "',0,1,'" & NamaLabel & "',
                                '" & ds.Tables(0).Rows(0).Item("TripUnitLabelNumber").ToString() & "','" & Me.ScanLabel.Text & "')"
                                adapter = New SqlDataAdapter(queryInsert, Main.koneksi)
                                adapter.SelectCommand.ExecuteNonQuery()
                            End If
                        Next

                        'For r = 1 To maxScan
                        '    If r = 1 And scan_rotary = "1" Then
                        '        Dim NamaLabel As String = "Rotary Handle Label"
                        '        Dim queryInsert = "insert into [ComponentsFuji]([order],[RefFuji],[WorkStation],[Check Components], 
                        '        [2ndScan],[2ndScanName],[2ndQrCode],[QRCodeFuji]) values ('" & Me.PPFujiEntry.Text & "','" & SplitData & "',
                        '        '" & Me.workstationFuji.Text & "',0,1,'" & NamaLabel & "',
                        '        '" & ds.Tables(0).Rows(0).Item("QRFrontLabelNumber").ToString() & "','" & Me.ScanLabel.Text & "')"
                        '        adapter = New SqlDataAdapter(queryInsert, Main.koneksi)
                        '        adapter.SelectCommand.ExecuteNonQuery()
                        '    ElseIf r = 2 And scan_front = "1" Then
                        '        Dim NamaLabel As String = "Front Cover Label"
                        '        Dim queryInsert = "insert into [ComponentsFuji]([order],[RefFuji],[WorkStation],[Check Components], 
                        '        [2ndScan],[2ndScanName],[2ndQrCode],[QRCodeFuji]) values ('" & Me.PPFujiEntry.Text & "','" & SplitData & "',
                        '        '" & Me.workstationFuji.Text & "',0,1,'" & NamaLabel & "',
                        '        '" & ds.Tables(0).Rows(0).Item("QRRotaryHandleNumber").ToString() & "','" & Me.ScanLabel.Text & "')"
                        '        adapter = New SqlDataAdapter(queryInsert, Main.koneksi)
                        '        adapter.SelectCommand.ExecuteNonQuery()
                        '        'ElseIf chk_scan_tripUnit.Checked = True And scan_trip = "1" Then
                        '    ElseIf r = 3 And chk_scan_tripUnit.Checked = True Then
                        '        Dim NamaLabel As String = "Trip Unit Label"
                        '        Dim queryInsert = "insert into [ComponentsFuji]([order],[RefFuji],[WorkStation],[Check Components], 
                        '        [2ndScan],[2ndScanName],[2ndQrCode],[QRCodeFuji]) values ('" & Me.PPFujiEntry.Text & "','" & SplitData & "',
                        '        '" & Me.workstationFuji.Text & "',0,1,'" & NamaLabel & "',
                        '        '" & ds.Tables(0).Rows(0).Item("TripUnitLabelNumber").ToString() & "','" & Me.ScanLabel.Text & "')"
                        '        adapter = New SqlDataAdapter(queryInsert, Main.koneksi)
                        '        adapter.SelectCommand.ExecuteNonQuery()
                        '    End If
                        'Next
                        Refresh_DGV_Fuji()
                        Refresh_DGV_Fuji_2()
                        'fuji printing label
                        btn_rotary_print_Click(sender, e)
                        btn_front_print_Click(sender, e)
                        btn_carton_print_Click(sender, e)

                        'MessageBox.Show("Print 3 Label")
                        ScanLabel2.Enabled = True
                        ScanLabel2.Text = ""
                        ScanLabel2.Select()
                        ScanLabel.Enabled = False


                        'this textbox for sidelabel
                        QRSideLabel.Text = ScanLabel.Text
                        FCSRefSideLabel.Text = FCSRef.Text
                        DateCodeSideLabel.Text = DateTime.Now.ToString("yyMMdd")
                    End If
                Else
                    Dim sql2 As String = "select * from ComponentsFuji where 
                    ComponentsFuji.[Order]='" & Me.PPFujiEntry.Text & "' and ComponentsFuji.Workstation='" & Me.workstationFuji.Text & "' 
                    and ComponentsFuji.QRCodeFuji='" & Me.ScanLabel.Text & "' and ComponentsFuji.[2ndScan]= 1 "
                    Dim ds2 As New DataSet
                    adapter2 = New SqlDataAdapter(sql2, Main.koneksi)
                    adapter2.Fill(ds2)
                    If ds2.Tables(0).Rows.Count > 0 Then

                        'santo edit
                        Dim cek_0 As String = "x"
                        Dim cek_1 As String = "x"
                        Dim cek_2 As String = "x"
                        Try
                            cek_0 = ds2.Tables(0).Rows(0).Item("Check Components").ToString()
                            cek_1 = ds2.Tables(0).Rows(1).Item("Check Components").ToString()
                            cek_2 = ds2.Tables(0).Rows(2).Item("Check Components").ToString()
                        Catch ex As Exception

                        End Try

                        If cek_0 = "0" Or cek_1 = "0" Or cek_2 = "0" Then
                            Refresh_DGV_Fuji_2()
                            ScanLabel2.Enabled = True
                            ScanLabel2.Text = ""
                            ScanLabel2.Select()
                            ScanLabel.Enabled = False

                        Else
                            MessageBox.Show("All Labels Have been Scan!")
                        End If
                    End If
                End If
                Dim sqlCounterItemsFuji As String = "select * from ComponentsFuji where 
                    ComponentsFuji.[Order]='" & Me.PPFujiEntry.Text & "' and ComponentsFuji.Workstation='" & Me.workstationFuji.Text & "' 
                    and ComponentsFuji.[Check Components]=1 and ComponentsFuji.[2ndScan]= 0 "
                Dim ds3 As New DataSet
                adapter3 = New SqlDataAdapter(sqlCounterItemsFuji, Main.koneksi)
                adapter3.Fill(ds3)
                CounterItemsFuji.Text = ds3.Tables(0).Rows.Count
            Else
                MessageBox.Show("Sorry Wrong QR Code")
            End If
        End If
    End Sub

    Private Sub Refresh_DGV_Fuji_2()
        ScanLabel2.Text = ""
        DataGridView8.Rows.Clear()
        Dim sql As String = "select [Order], [RefFuji],[Check Components],[2ndScanName],[2ndQrCode] from ComponentsFuji 
        where [Order]='" & Me.PPFujiEntry.Text & "' and Workstation='" & Me.workstationFuji.Text & "' and [2ndScan]=1 
        and [QRCodeFuji]='" & ScanLabel.Text & "' order by ID"
        Dim ds As New DataSet
        adapter = New SqlDataAdapter(sql, Main.koneksi)
        adapter.Fill(ds)
        If ds.Tables(0).Rows.Count > 0 Then
            DataGridView8.ColumnCount = 5
            DataGridView8.Columns(0).Name = "Order"
            DataGridView8.Columns(1).Name = "FCS Refrence"
            DataGridView8.Columns(2).Name = "Name"
            DataGridView8.Columns(3).Name = "Code"
            DataGridView8.Columns(4).Name = "Check Scan"
            For r = 0 To ds.Tables(0).Rows.Count - 1
                Dim row As String() = New String() {ds.Tables(0).Rows(r).Item("Order").ToString(), ds.Tables(0).Rows(r).Item("RefFuji").ToString(), ds.Tables(0).Rows(r).Item("2ndScanName").ToString(), ds.Tables(0).Rows(r).Item("2ndQrCode").ToString(), ds.Tables(0).Rows(r).Item("Check Components").ToString()}
                DataGridView8.Rows.Add(row)
                If ds.Tables(0).Rows(r).Item("Check Components").ToString() = 1 Then
                    DataGridView8.Rows(r).Cells(4).Style.BackColor = Color.SkyBlue
                Else
                    DataGridView8.Rows(r).Cells(4).Style.BackColor = Color.White
                End If
            Next
        End If
    End Sub

    Private Sub ScanLabel2_KeyPress(sender As Object, e As PreviewKeyDownEventArgs) Handles ScanLabel2.PreviewKeyDown
        If (e.KeyData = Keys.Tab Or e.KeyData = Keys.Enter) And Len(Me.ScanLabel2.Text) >= 9 Then
            Dim sql As String = "select * from ComponentsFuji where [Order]= " & PPFujiEntry.Text & " and [2ndScan]=1 and [QRCodeFuji]='" & ScanLabel.Text & "' order by ID"
            Dim ds As New DataSet
            adapter = New SqlDataAdapter(sql, Main.koneksi)
            adapter.Fill(ds)
            If ds.Tables(0).Rows.Count > 0 Then
                For r = 0 To ds.Tables(0).Rows.Count - 1
                    If r = 0 Then
                        If ds.Tables(0).Rows(r).Item("Check Components").ToString() = 0 And ds.Tables(0).Rows(r).Item("2ndQrCode").ToString() = ScanLabel2.Text And ds.Tables(0).Rows(r).Item("2ndScanName").ToString() = "Rotary Handle Label" Then
                            'MessageBox.Show("update rotary handle")
                            Dim queryUpdate = "UPDATE [ComponentsFuji] SET [Check Components] = 1 where 
                            [Order]='" & Me.PPFujiEntry.Text & "' and Workstation='" & Me.workstationFuji.Text & "' 
                            and [2ndQrCode]= '" & ScanLabel2.Text & "' and [QRCodeFuji]='" & ScanLabel.Text & "'"
                            adapter = New SqlDataAdapter(queryUpdate, Main.koneksi)
                            If adapter.SelectCommand.ExecuteNonQuery().ToString() = 1 Then
                                Refresh_DGV_Fuji_2()
                            End If
                        End If
                    ElseIf r = 1 Then
                        If ds.Tables(0).Rows(r - 1).Item("Check Components").ToString() = 0 And ds.Tables(0).Rows(r).Item("2ndQrCode").ToString() = ScanLabel2.Text And ds.Tables(0).Rows(r).Item("2ndScanName").ToString() = "Front Cover Label" Then
                            MessageBox.Show("Please Scan Rotary Handle First")
                        ElseIf ds.Tables(0).Rows(r - 1).Item("Check Components").ToString() = 1 And ds.Tables(0).Rows(r).Item("2ndQrCode").ToString() = ScanLabel2.Text And ds.Tables(0).Rows(r).Item("2ndScanName").ToString() = "Front Cover Label" Then
                            Dim queryUpdate = "UPDATE [ComponentsFuji] SET [Check Components] = 1 where 
                            [Order]='" & Me.PPFujiEntry.Text & "' and Workstation='" & Me.workstationFuji.Text & "' 
                            and [2ndQrCode]= '" & ScanLabel2.Text & "' and [QRCodeFuji]='" & ScanLabel.Text & "'"
                            adapter = New SqlDataAdapter(queryUpdate, Main.koneksi)
                            If adapter.SelectCommand.ExecuteNonQuery().ToString() = 1 Then
                                Refresh_DGV_Fuji_2()
                                If chk_scan_tripUnit.Checked = False Then
                                    Dim sql2 As String = "select * from ComponentsFuji where [Order]= " & PPFujiEntry.Text & " and [Check Components]=0"
                                    Dim ds2 As New DataSet
                                    adapter2 = New SqlDataAdapter(sql2, Main.koneksi)
                                    adapter2.Fill(ds2)
                                    If ds2.Tables(0).Rows.Count > 0 Then
                                        ScanLabel.Text = ""
                                        ScanLabel2.Text = ""
                                        ScanLabel.Enabled = True
                                        ScanLabel2.Enabled = False
                                        DataGridView8.Rows.Clear()
                                        ScanLabel.Select()
                                    Else
                                        ScanLabel.Text = ""
                                        ScanLabel2.Text = ""
                                        PPFujiEntry.Text = ""
                                        PPFujiEntry.Select()
                                        ScanLabel.Enabled = False
                                        ScanLabel2.Enabled = False
                                        DataGridView7.Rows.Clear()
                                        DataGridView8.Rows.Clear()
                                        FCSRef.Text = ""
                                    End If
                                End If
                            End If
                        End If
                    Else
                        If ds.Tables(0).Rows(r - 1).Item("Check Components").ToString() = 0 And ds.Tables(0).Rows(r).Item("2ndQrCode").ToString() = ScanLabel2.Text And ds.Tables(0).Rows(r).Item("2ndScanName").ToString() = "Trip Unit Label" Then
                            MessageBox.Show("Please Scan Front Cover Label First")
                        ElseIf ds.Tables(0).Rows(r - 1).Item("Check Components").ToString() = 1 And ds.Tables(0).Rows(r).Item("2ndQrCode").ToString() = ScanLabel2.Text And ds.Tables(0).Rows(r).Item("2ndScanName").ToString() = "Trip Unit Label" Then
                            Dim queryUpdate = "UPDATE [ComponentsFuji] SET [Check Components] = 1 where 
                            [Order]='" & Me.PPFujiEntry.Text & "' and Workstation='" & Me.workstationFuji.Text & "' 
                            and [2ndQrCode]= '" & ScanLabel2.Text & "' and [QRCodeFuji]='" & ScanLabel.Text & "'"
                            adapter = New SqlDataAdapter(queryUpdate, Main.koneksi)
                            If adapter.SelectCommand.ExecuteNonQuery().ToString() = 1 Then
                                Refresh_DGV_Fuji_2()
                            End If

                            Dim sql2 As String = "select * from ComponentsFuji where [Order]= " & PPFujiEntry.Text & " and [Check Components]=0"
                            Dim ds2 As New DataSet
                            adapter2 = New SqlDataAdapter(sql2, Main.koneksi)
                            adapter2.Fill(ds2)
                            If ds2.Tables(0).Rows.Count > 0 Then
                                ScanLabel.Text = ""
                                ScanLabel2.Text = ""
                                ScanLabel.Enabled = True
                                ScanLabel2.Enabled = False
                                DataGridView8.Rows.Clear()
                                ScanLabel.Select()
                            Else
                                ScanLabel.Text = ""
                                ScanLabel2.Text = ""
                                PPFujiEntry.Text = ""
                                PPFujiEntry.Select()
                                ScanLabel.Enabled = False
                                ScanLabel2.Enabled = False
                                DataGridView7.Rows.Clear()
                                DataGridView8.Rows.Clear()
                                FCSRef.Text = ""
                            End If
                        End If
                    End If
                Next
                ScanLabel2.Text = ""
                Refresh_DGV_Fuji()
            Else
                MessageBox.Show("Sorry Wrong QR Code")
            End If
        End If
    End Sub

    Private Sub Tab_Fuji_Enter(sender As Object, e As EventArgs) Handles Tab_Fuji.Enter
        'MessageBox.Show("testing")
    End Sub

    Private Sub workstationFuji_SelectedIndexChanged(sender As Object, e As EventArgs) Handles workstationFuji.SelectedIndexChanged
        workstation.SelectedIndex = workstationFuji.SelectedIndex

        'workstation_event()
    End Sub

    Private Sub cbx_fuji_side_label_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbx_fuji_side_label.SelectedIndexChanged
        'label_side_printer.PrintSettings.PrinterName = cbx_fuji_side_label.SelectedText

        If printers.Count > 0 Then
            selected_Printer = printers.Item(cbx_fuji_side_label.SelectedIndex)
            label_side_printer.PrintSettings.PrinterName = selected_Printer.Name
            cbx_fuji_side_label.SelectedItem = selected_Printer.Name
        End If
    End Sub

    Private Sub cbx_Carton_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbx_Carton.SelectedIndexChanged
        'label_carton_printer.PrintSettings.PrinterName = cbx_Carton.SelectedText

        If printers.Count > 0 Then
            selected_Printer = printers.Item(cbx_Carton.SelectedIndex)
            label_carton_printer.PrintSettings.PrinterName = selected_Printer.Name
            cbx_Carton.SelectedItem = selected_Printer.Name
        End If
    End Sub

    Private Sub cbx_front_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbx_front.SelectedIndexChanged
        'label_front_short_printer.PrintSettings.PrinterName = cbx_front.SelectedText
        'label_front_long_printer.PrintSettings.PrinterName = cbx_front.SelectedText

        If printers.Count > 0 Then
            selected_Printer = printers.Item(cbx_front.SelectedIndex)
            label_front_short_printer.PrintSettings.PrinterName = selected_Printer.Name
            label_front_long_printer.PrintSettings.PrinterName = selected_Printer.Name
            cbx_front.SelectedItem = selected_Printer.Name
        End If

    End Sub

    Private Sub cbx_outside_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbx_outside.SelectedIndexChanged
        'label_out_side_printer.PrintSettings.PrinterName = cbx_outside.SelectedText
        If printers.Count > 0 Then
            selected_Printer = printers.Item(cbx_outside.SelectedIndex)
            label_out_side_printer.PrintSettings.PrinterName = selected_Printer.Name
            cbx_outside.SelectedItem = selected_Printer.Name
        End If
    End Sub

    Private Sub cbx_Rotary_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbx_Rotary.SelectedIndexChanged
        'label_rotary_printer.PrintSettings.PrinterName = cbx_Rotary.SelectedText

        If printers.Count > 0 Then
            selected_Printer = printers.Item(cbx_Rotary.SelectedIndex)
            label_rotary_printer.PrintSettings.PrinterName = selected_Printer.Name
            cbx_Rotary.SelectedItem = selected_Printer.Name
        End If

    End Sub

    'Set Value Side Label

    Private Sub Fuji_side_SetValue()

        Dim date_code As String = dateCode.Text

        'Fuji_QR_Product_Label.Text = header.Text & ":/sn=" & "SG" & date_code.Substring(2, 2) & date_code.Substring(6, 2) & date_code.Substring(date_code.Length - 1, 1)
        Try

            label_side_printer.Variables("QR Product Label").SetValue(Fuji_QR_Product_Label.Text)
            label_side_printer.Variables("1 2").SetValue(header.Text)
        Catch ex As Exception
            MsgBox("Fuji Side Label " & ex.Message)
        End Try
    End Sub

    Private Sub Fuji_Rotary_setvalue()
        Try

            label_rotary_printer.Variables("1").SetValue(Rotary1.Text)
            label_rotary_printer.Variables("2").SetValue(Rotary2.Text)
            label_rotary_printer.Variables("3").SetValue(Rotary3.Text)
            label_rotary_printer.Variables("4").SetValue(Rotary4.Text)
            label_rotary_printer.Variables("5").SetValue(Rotary5.Text)
            label_rotary_printer.Variables("6").SetValue(Rotary6.Text)
            label_rotary_printer.Variables("7").SetValue(Rotary7.Text)

            label_rotary_printer.Variables("QR Code Rotary Handle Number").SetValue(QRForRotaryHandle.Text)
            label_rotary_printer.Variables("Rotary Handle Label Number").SetValue(LotNumberRotaryHandle.Text)
        Catch ex As Exception
            MsgBox("Fuji Rotary Label " & ex.Message)
        End Try
    End Sub

    Private Sub Fuji_Front_long_setvalue()
        Try

            label_front_long_printer.Variables("1").SetValue(Front1.Text)
            label_front_long_printer.Variables("2").SetValue(Front2.Text)
            label_front_long_printer.Variables("3").SetValue(Front3.Text)
            label_front_long_printer.Variables("4").SetValue(Front4.Text)
            label_front_long_printer.Variables("5").SetValue(Front5.Text)
            label_front_long_printer.Variables("6").SetValue(Front6.Text)
            label_front_long_printer.Variables("7").SetValue(Front7.Text)

            label_front_long_printer.Variables("QR Code Front Label Number").SetValue(QRForCoverLabel.Text)
            label_front_long_printer.Variables("Front Label Number").SetValue(LotNumberCoverLabel.Text)
        Catch ex As Exception
            MsgBox("Fuji front Label " & ex.Message)
        End Try
    End Sub
    Private Sub Fuji_Front_short_setvalue()
        Try

            label_front_short_printer.Variables("1").SetValue(Front1.Text)
            label_front_short_printer.Variables("2").SetValue(Front2.Text)
            label_front_short_printer.Variables("3").SetValue(Front3.Text)
            label_front_short_printer.Variables("4").SetValue(Front4.Text)
            label_front_short_printer.Variables("5").SetValue(Front5.Text)
            label_front_short_printer.Variables("6").SetValue(Front6.Text)
            label_front_short_printer.Variables("7").SetValue(Front7.Text)

            label_front_short_printer.Variables("QR Code Front Label Number").SetValue(QRForCoverLabel.Text)
            label_front_short_printer.Variables("Front Label Number").SetValue(LotNumberCoverLabel.Text)
        Catch ex As Exception
            MsgBox("Fuji front Label " & ex.Message)
        End Try
    End Sub
    Private Sub Fuji_carton_setvalue()
        Try

            label_carton_printer.Variables("1").SetValue(Carton1.Text)
            label_carton_printer.Variables("2").SetValue(Carton2.Text)
            label_carton_printer.Variables("3 2").SetValue(Carton32.Text)
            label_carton_printer.Variables("4").SetValue(Carton4.Text)

        Catch ex As Exception
            MsgBox("Fuji carton Label " & ex.Message)
        End Try
    End Sub
    Private Sub Fuji_outside_setvalue()
        Try

            label_out_side_printer.Variables("1").SetValue(Carton1.Text)
            label_out_side_printer.Variables("2").SetValue(Carton2.Text)
            label_out_side_printer.Variables("3 2").SetValue(Carton32.Text)
            label_out_side_printer.Variables("4").SetValue(Carton4.Text)

            label_out_side_printer.Variables("pp").SetValue(PPFujiEntry.Text)
            label_out_side_printer.Variables("co").SetValue(FujiSO.Text)
            label_out_side_printer.Variables("SO_Line_Number").SetValue(FujiLineNo.Text)

            label_out_side_printer.Variables("PO").SetValue(PO_Group_BOX.Text)
            label_out_side_printer.Variables("QR_SO").SetValue("0" & FujiSO.Text & Convert.ToDecimal(FujiLineNo.Text).ToString("000000"))


        Catch ex As Exception
            MsgBox("Fuji Outside Label " & ex.Message)
        End Try
    End Sub

    Private Sub btn_preview_fuji_side_label_Click(sender As Object, e As EventArgs) Handles btn_preview_fuji_side_label.Click
        reload_printer()
        Fuji_side_SetValue()

        'declaration of Preview
        Dim LabelPreviewSettings As ILabelPreviewSettings = New LabelPreviewSettings()

        'setting preview format
        LabelPreviewSettings.ImageFormat = "PNG"
        LabelPreviewSettings.Width = Form_preview.pictureBoxPreview.Width                   ' Width Of image To generate
        LabelPreviewSettings.Height = Form_preview.pictureBoxPreview.Height                 ' Height Of image To generate

        ' Generate Preview File
        Dim imageObj As Object = label_side_printer.GetLabelPreview(LabelPreviewSettings)

        'Display image in UI
        If TypeOf imageObj Is Byte() Then
            Form_preview.pictureBoxPreview.Image = ByteToImage(CType(imageObj, Byte()))
        ElseIf TypeOf imageObj Is String Then
            Form_preview.pictureBoxPreview.ImageLocation = CStr(imageObj)
        End If

        Form_preview.Show()
    End Sub

    Private Sub btn_rotary_handle_preview_Click(sender As Object, e As EventArgs) Handles btn_rotary_handle_preview.Click
        reload_printer()
        Fuji_Rotary_setvalue()

        'declaration of Preview
        Dim LabelPreviewSettings As ILabelPreviewSettings = New LabelPreviewSettings()

        'setting preview format
        LabelPreviewSettings.ImageFormat = "PNG"
        LabelPreviewSettings.Width = Form_preview.pictureBoxPreview.Width                   ' Width Of image To generate
        LabelPreviewSettings.Height = Form_preview.pictureBoxPreview.Height                 ' Height Of image To generate

        ' Generate Preview File
        Dim imageObj As Object = label_rotary_printer.GetLabelPreview(LabelPreviewSettings)

        'Display image in UI
        If TypeOf imageObj Is Byte() Then
            Form_preview.pictureBoxPreview.Image = ByteToImage(CType(imageObj, Byte()))
        ElseIf TypeOf imageObj Is String Then
            Form_preview.pictureBoxPreview.ImageLocation = CStr(imageObj)
        End If

        Form_preview.Show()
    End Sub

    Private Sub btn_front_label_preview_Click(sender As Object, e As EventArgs) Handles btn_front_label_preview.Click
        reload_printer()
        Fuji_Front_long_setvalue()
        Fuji_Front_short_setvalue()

        'declaration of Preview
        Dim LabelPreviewSettings As ILabelPreviewSettings = New LabelPreviewSettings()

        'setting preview format
        LabelPreviewSettings.ImageFormat = "PNG"
        LabelPreviewSettings.Width = Form_preview.pictureBoxPreview.Width                   ' Width Of image To generate
        LabelPreviewSettings.Height = Form_preview.pictureBoxPreview.Height                 ' Height Of image To generate

        ' Generate Preview File
        Dim imageObj As Object = label_front_short_printer.GetLabelPreview(LabelPreviewSettings)

        'Display image in UI
        If TypeOf imageObj Is Byte() Then
            Form_preview.pictureBoxPreview.Image = ByteToImage(CType(imageObj, Byte()))
        ElseIf TypeOf imageObj Is String Then
            Form_preview.pictureBoxPreview.ImageLocation = CStr(imageObj)
        End If

        Form_preview.Show()
    End Sub

    Private Sub btn_carton_label_preview_Click(sender As Object, e As EventArgs) Handles btn_carton_label_preview.Click
        reload_printer()
        Fuji_carton_setvalue()

        'declaration of Preview
        Dim LabelPreviewSettings As ILabelPreviewSettings = New LabelPreviewSettings()

        'setting preview format
        LabelPreviewSettings.ImageFormat = "PNG"
        LabelPreviewSettings.Width = Form_preview.pictureBoxPreview.Width                   ' Width Of image To generate
        LabelPreviewSettings.Height = Form_preview.pictureBoxPreview.Height                 ' Height Of image To generate

        ' Generate Preview File
        Dim imageObj As Object = label_carton_printer.GetLabelPreview(LabelPreviewSettings)

        'Display image in UI
        If TypeOf imageObj Is Byte() Then
            Form_preview.pictureBoxPreview.Image = ByteToImage(CType(imageObj, Byte()))
        ElseIf TypeOf imageObj Is String Then
            Form_preview.pictureBoxPreview.ImageLocation = CStr(imageObj)
        End If

        Form_preview.Show()
    End Sub

    Private Sub btn_outside_label_preview_Click(sender As Object, e As EventArgs) Handles btn_outside_label_preview.Click
        reload_printer()
        Fuji_outside_setvalue()
        cek_manual_group_print()
        'declaration of Preview
        Dim LabelPreviewSettings As ILabelPreviewSettings = New LabelPreviewSettings()

        'setting preview format
        LabelPreviewSettings.ImageFormat = "PNG"
        LabelPreviewSettings.Width = Form_preview.pictureBoxPreview.Width                   ' Width Of image To generate
        LabelPreviewSettings.Height = Form_preview.pictureBoxPreview.Height                 ' Height Of image To generate

        ' Generate Preview File
        Dim imageObj As Object = label_out_side_printer.GetLabelPreview(LabelPreviewSettings)

        'Display image in UI
        If TypeOf imageObj Is Byte() Then
            Form_preview.pictureBoxPreview.Image = ByteToImage(CType(imageObj, Byte()))
        ElseIf TypeOf imageObj Is String Then
            Form_preview.pictureBoxPreview.ImageLocation = CStr(imageObj)
        End If

        Form_preview.Show()
    End Sub

    Private Sub btn_rotary_print_Click(sender As Object, e As EventArgs) Handles btn_rotary_print.Click

        Try
            'label5_printer.PrintSettings.PrinterName = label_printer.PrintSettings.PrinterName
            Fuji_Rotary_setvalue()
            If qty = 0 Then qty = 1
            label_rotary_printer.Print(qty)
        Catch ex As Exception
            MsgBox("Print Failed" & ex.ToString)
        End Try
    End Sub

    Private Sub Btn_Fuji_Side_label_Click(sender As Object, e As EventArgs) Handles Btn_Fuji_Side_label.Click

        Try
            Fuji_side_SetValue()
            If qty = 0 Then qty = 1
            label_side_printer.Print(qty)
            'updateTraceability()
        Catch ex As Exception
            MsgBox("Print Failed" & ex.ToString)
        End Try
    End Sub

    Private Sub btn_front_print_Click(sender As Object, e As EventArgs) Handles btn_front_print.Click

        Try
            Fuji_Front_long_setvalue()
            Fuji_Front_short_setvalue()
            If qty = 0 Then qty = 1
            If FujiLabelType.Text = "Short" Then
                label_front_short_printer.Print(qty)
            Else
                label_front_long_printer.Print(qty)
            End If

        Catch ex As Exception
            MsgBox("Print Failed" & ex.ToString)
        End Try
    End Sub

    Private Sub btn_carton_print_Click(sender As Object, e As EventArgs) Handles btn_carton_print.Click
        Try
            Fuji_carton_setvalue()
            If qty = 0 Then qty = 1
            label_carton_printer.Print(qty)
        Catch ex As Exception
            MsgBox("Print Failed" & ex.ToString)
        End Try
    End Sub

    Private Sub Print_OutSide_Label()
        Try
            Fuji_outside_setvalue()
            If qty = 0 Then qty = 1
            label_out_side_printer.Print(qty)
        Catch ex As Exception
            MsgBox("Print Failed" & ex.ToString)
        End Try
    End Sub
    Private Sub cek_manual_group_print()
        If Convert.ToDecimal(No_Box.Text) = 1 Then
            If Convert.ToDecimal(fujiQty.Text) < Convert.ToDecimal(GroupingBox.Text) Then
                label_out_side_printer.Variables("pcs").SetValue(fujiQty.Text)
            Else
                label_out_side_printer.Variables("pcs").SetValue(GroupingBox.Text)
            End If
        ElseIf Convert.ToDecimal(No_Box.Text) <> 0 Then
            Dim cek_pcs As Integer = Convert.ToDecimal(No_Box.Text) * Convert.ToDecimal(GroupingBox.Text)
            If cek_pcs <= Convert.ToDecimal(fujiQty.Text) Then
                Dim cek_2 As Integer = Convert.ToDecimal(fujiQty.Text)
                label_out_side_printer.Variables("pcs").SetValue(GroupingBox.Text)
            ElseIf (cek_pcs - Convert.ToDecimal(fujiQty.Text)) < Convert.ToDecimal(GroupingBox.Text) Then
                label_out_side_printer.Variables("pcs").SetValue(Convert.ToDecimal(fujiQty.Text) - (cek_pcs - Convert.ToDecimal(GroupingBox.Text)))
            Else
                MessageBox.Show("No. Box is too High !")
            End If
        End If
    End Sub
    Private Sub btn_outside_print_Click(sender As Object, e As EventArgs) Handles btn_outside_print.Click
        cek_manual_group_print()
        Print_OutSide_Label()
    End Sub

    Private Sub Button10_Click_1(sender As Object, e As EventArgs) Handles Button10.Click
        Form1.open_form("select * from MasterFuji")
        Form1.Show()
    End Sub

    Private Sub Button11_Click_1(sender As Object, e As EventArgs) Handles Button11.Click
        Form1.open_form("select * from MasterFujiLabelling")
        Form1.Show()
    End Sub

    Private Sub CounterItemsFuji_TextChanged(sender As Object, e As EventArgs) Handles CounterItemsFuji.TextChanged
        Try
            Dim a As Integer = Convert.ToDecimal(CounterItemsFuji.Text)
            Dim b As Integer = Convert.ToDecimal(CounterItemsFuji.Text)
            Dim c As Integer = Convert.ToDecimal(GroupingBox.Text)
            Dim d As Integer
            Dim es As Integer

            es = Int(a / c) + 1
            No_Box.Text = es

            d = b Mod c
            If a = 1 Or (1 = d And b > c) Then
                'MessageBox.Show("Auto Print Group")
                cek_manual_group_print()
                Print_OutSide_Label()
            End If
        Catch ex As Exception

        End Try

    End Sub

    Private Sub TraceFujiProduct_Click(sender As Object, e As EventArgs) Handles TraceFujiProduct.Click
        Call Main.koneksi_db()
        If Len(TextBox2.Text) >= 33 Then
            DataGridView2.Rows.Clear()
            DataGridView3.Rows.Clear()
            DataGridView4.Rows.Clear()
            Dim dsTrace As New DataSet
            Dim sql = "select DISTINCT ProductionOrders.PP from ProductionOrders, printingRecord where ProductionOrders.PP=printingRecord.PP and printingRecord.QRCodeFuji='" & TextBox2.Text & "'"
            Dim adapter = New SqlDataAdapter(sql, Main.koneksi)
            adapter.Fill(dsTrace)
            If dsTrace.Tables(0).Rows.Count > 0 Then
                TextBox1.Text = dsTrace.Tables(0).Rows(0).Item("PP")
                Trace.PerformClick()
            Else
                MessageBox.Show("Sorry Data Not Found")
            End If
        End If
    End Sub

    Private Sub btn_export_BOM_Fuji_Click(sender As Object, e As EventArgs) Handles btn_export_BOM_Fuji.Click
        ExportToExcel("SELECT * FROM MasterFuji")
    End Sub

    Private Sub btn_Export_Labelling_fuji_Click(sender As Object, e As EventArgs) Handles btn_Export_Labelling_fuji.Click
        ExportToExcel("SELECT * FROM MasterFujiLabelling")
    End Sub

    Private Sub Button12_Click_1(sender As Object, e As EventArgs) Handles btn_log.Click
        Log_form.Show()
    End Sub

    Private Sub Button12_Click_2(sender As Object, e As EventArgs) Handles Button12.Click
        Try
            Dim Sql2 = "UPDATE openOrders set [Reqmts qty] = REPLACE([Reqmts qty],'.','') WHERE [Reqmts qty] LIKE '%.%';"
            Dim insert2 = New SqlCommand(Sql2, Main.koneksi)
            'insert2.ExecuteNonQuery()

            Dim afected As Integer = insert2.ExecuteNonQuery()
            If afected > 0 Then MsgBox("Deleting '.' Succeed, Rows Affected:" & afected.ToString)
        Catch ex As Exception
            MsgBox("Deleting '.' Failed !")
        End Try
    End Sub

    Private Sub btn_delete_dot_sq00_Click(sender As Object, e As EventArgs) Handles btn_delete_dot_sq00.Click
        Try
            Dim Sql2 = "UPDATE PPList set [Item quantity] = REPLACE([Item quantity],'.','') WHERE [Item quantity] LIKE '%.%';"
            Dim insert2 = New SqlCommand(Sql2, Main.koneksi)
            'insert2.ExecuteNonQuery()

            Dim afected As Integer = insert2.ExecuteNonQuery()
            If afected > 0 Then MsgBox("Deleting '.' Succeed, Rows Affected:" & afected.ToString)
        Catch ex As Exception
            MsgBox("Deleting '.' Failed !")
        End Try
    End Sub


    Private Sub print_progress(ByVal a As Integer)
        Progress_print_all.Value = a
        Application.DoEvents()
    End Sub

    Private Sub Button13_Click(sender As Object, e As EventArgs)
        Dim sql2 As String = "SELECT * FROM SGRAC_MES.dbo.MasterFuji WHERE FCSRef ='" & FCSRef.Text & "'"
        Dim ds2 As New DataSet
        adapter2 = New SqlDataAdapter(sql2, Main.koneksi)
        adapter2.Fill(ds2)

        Dim scan_rotary As String = ds2.Tables(0).Rows(0).Item("ScanRotaryHandleLabel").ToString()
        Dim scan_front As String = ds2.Tables(0).Rows(0).Item("ScanFrontLabel").ToString()
        Dim scan_trip As String = ds2.Tables(0).Rows(0).Item("ScanTripUnitLabel").ToString()

        MsgBox(scan_rotary & " " & scan_front & " " & scan_trip & " ")

    End Sub

    Private Sub Label157_Click(sender As Object, e As EventArgs) Handles Label157.Click

    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        Dim confirm = MessageBox.Show("Are You Sure For Logout?", "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If confirm = Windows.Forms.DialogResult.Yes Then
            Me.Hide()
            LoginForm.Show()
            LoginForm.textboxusername.Text = ""
            LoginForm.textboxpassword.Text = ""
        ElseIf confirm = Windows.Forms.DialogResult.Yes Then
            Exit Sub
        End If
        role = ""
    End Sub

    '--Code Ruby Start From Here--'

    Private Sub PPRubyEntry_KeyPress(sender As Object, e As PreviewKeyDownEventArgs) Handles PPRubyEntry.PreviewKeyDown 'sudah sesuai dengan VBA
        If Len(Me.PPRubyEntry.Text) = 12 Then Me.PPRubyEntry.Text = Microsoft.VisualBasic.Right(Me.PPRubyEntry.Text, 11)

        If (e.KeyData = Keys.Tab Or e.KeyData = Keys.Enter) And Len(Me.PPRubyEntry.Text) >= 11 Then
            If String.IsNullOrEmpty(workstationRuby.Text) = False And String.IsNullOrEmpty(technicianNameRuby.Text) = False Then
                If IsNumeric(PPRubyEntry.Text) Then
                    TextBox4.Text = 0
                    TextBox5.Text = 0
                    Dim sqlCekRuby As String = "select * from MasterRuby, PPList, MasterRangeLabelRuby, MasterPackagingLabelRuby where PPList.Material=MasterRuby.CRRef and PPList.[Order]='" & Me.PPRubyEntry.Text & "' and PPList.[Material]=MasterRangeLabelRuby.[CRRef] and PPList.[Material]=MasterPackagingLabelRuby.CRRef"
                    Dim dsCekRuby As New DataSet
                    adapter = New SqlDataAdapter(sqlCekRuby, Main.koneksi)
                    adapter.Fill(dsCekRuby)
                    If (dsCekRuby.Tables(0).Rows.Count > 0) Then
                        Dim sqlCekForInsert As String = "select * from PPList, ComponentsRuby where PPlist.[Order]='" & Me.PPRubyEntry.Text & "' and ComponentsRuby.RefRuby=PPList.Material and ComponentsRuby.Workstation='" & Me.workstationRuby.Text & "'"
                        Dim dsCekForInsert As New DataSet
                        adapter = New SqlDataAdapter(sqlCekForInsert, Main.koneksi)
                        adapter.Fill(dsCekForInsert)

                        ComRefRuby.Text = dsCekRuby.Tables(0).Rows(0).Item("CRRef").ToString
                        Me.dateCodeRuby.Text = dateCode2()

                        Dim row_v As DataRowView = technicianNameRuby.SelectedItem
                        If (Not row_v Is Nothing) Then
                            Dim row As DataRow = row_v.Row
                            Dim itemName As String = row(2).ToString()
                            technicianShortNameRuby.Text = itemName
                        End If

                        'Keluarkan semua data dibawah ini
                        RatedCurrRuby.Text = dsCekRuby.Tables(0).Rows(0).Item("RatedCurr").ToString
                        RatedVoltRuby.Text = dsCekRuby.Tables(0).Rows(0).Item("RatedVolt").ToString
                        UtilizationRuby.Text = dsCekRuby.Tables(0).Rows(0).Item("Utilization").ToString
                        RatedFreq.Text = dsCekRuby.Tables(0).Rows(0).Item("RatedFreq").ToString
                        ShortCircuitRuby.Text = dsCekRuby.Tables(0).Rows(0).Item("ShortCircuit").ToString
                        wotpc.Text = dsCekRuby.Tables(0).Rows(0).Item("WOTPC").ToString
                        warningRuby.Text = dsCekRuby.Tables(0).Rows(0).Item("Warning").ToString

                        'santo add
                        Phase_N_Logo.Text = dsCekRuby.Tables(0).Rows(0).Item("Phase N logo").ToString
                        active_under_QRCode.Text = dsCekRuby.Tables(0).Rows(0).Item("Active under QR Code").ToString

                        SORuby.Text = dsCekRuby.Tables(0).Rows(0).Item("SO no").ToString
                        SoLineRuby.Text = dsCekRuby.Tables(0).Rows(0).Item("item").ToString
                        POGroupBoxRuby.Text = dsCekRuby.Tables(0).Rows(0).Item("Purchase order number").ToString

                        PerfLabel.Text = dsCekRuby.Tables(0).Rows(0).Item("PerfLabel").ToString
                        QtyPerfLabel.Text = dsCekRuby.Tables(0).Rows(0).Item("QTY3").ToString
                        SERefRuby.Text = dsCekRuby.Tables(0).Rows(0).Item("SeRef").ToString
                        QtySERefRuby.Text = dsCekRuby.Tables(0).Rows(0).Item("QTY1").ToString
                        ScrewAcc.Text = dsCekRuby.Tables(0).Rows(0).Item("ScrewAcc").ToString
                        QtyScrewAcc.Text = dsCekRuby.Tables(0).Rows(0).Item("QTY2").ToString
                        EXT24.Text = dsCekRuby.Tables(0).Rows(0).Item("EXT24").ToString
                        QtyEXT24.Text = dsCekRuby.Tables(0).Rows(0).Item("QTY6").ToString
                        RS485.Text = dsCekRuby.Tables(0).Rows(0).Item("ModRS485").ToString
                        QtyRS485.Text = dsCekRuby.Tables(0).Rows(0).Item("QTY7").ToString
                        CartonBoxRuby.Text = dsCekRuby.Tables(0).Rows(0).Item("CartonBox").ToString
                        QtyCartonBox.Text = dsCekRuby.Tables(0).Rows(0).Item("QTY8").ToString
                        ProductTypeRuby.Text = dsCekRuby.Tables(0).Rows(0).Item("ProductType").ToString

                        LangCn.Text = dsCekRuby.Tables(0).Rows(0).Item("cn").ToString
                        LangEn.Text = dsCekRuby.Tables(0).Rows(0).Item("en").ToString
                        LangEs.Text = dsCekRuby.Tables(0).Rows(0).Item("es").ToString
                        LangFr.Text = dsCekRuby.Tables(0).Rows(0).Item("fr").ToString
                        LangKz.Text = dsCekRuby.Tables(0).Rows(0).Item("kz").ToString
                        LangRu.Text = dsCekRuby.Tables(0).Rows(0).Item("ru").ToString
                        RangeNameRuby.Text = dsCekRuby.Tables(0).Rows(0).Item("RangeName").ToString
                        EAN13Ruby.Text = dsCekRuby.Tables(0).Rows(0).Item("EAN13").ToString
                        QTYSingleBoxRuby.Text = dsCekRuby.Tables(0).Rows(0).Item("QTY").ToString
                        MadeInCountryRuby.Text = dsCekRuby.Tables(0).Rows(0).Item("MadeInCountry").ToString
                        SiteAddressRuby.Text = dsCekRuby.Tables(0).Rows(0).Item("SiteAddress").ToString
                        ZipCodeRuby.Text = dsCekRuby.Tables(0).Rows(0).Item("ZipCode").ToString
                        NameChRuby.Text = dsCekRuby.Tables(0).Rows(0).Item("NameCh").ToString
                        AddressChRuby.Text = dsCekRuby.Tables(0).Rows(0).Item("AddressCh").ToString
                        LogisticRefRuby.Text = dsCekRuby.Tables(0).Rows(0).Item("LogisticRef").ToString
                        CIPRuby.Text = dsCekRuby.Tables(0).Rows(0).Item("CIP").ToString

                        Dim pir As String = ""

                        pir = dsCekRuby.Tables(0).Rows(0).Item("ProductPicture").ToString
                        'pir = pir.Replace("1P", "4P")
                        'pir = pir.Replace("2P", "4P")
                        'pir = pir.Replace("3P", "4P")
                        ProductImageRuby.Text = pir

                        If (dsCekForInsert.Tables(0).Rows.Count > 0) Then
                            Refresh_DGV_Ruby()
                            Refresh_DGV_Ruby2()
                        Else
                            Dim queryInsert = "insert into [ComponentsRuby]([order],[RefRuby],[WorkStation], 
                            [Qty]) values ('" & Me.PPRubyEntry.Text & "',(select Top 1 material from PPList 
                            where [order]='" & Me.PPRubyEntry.Text & "'),'" & Me.workstationRuby.Text & "','" & dsCekRuby.Tables(0).Rows(0).Item("QTY3").ToString & "')"
                            Dim adapterInsert = New SqlDataAdapter(queryInsert, Main.koneksi)
                            adapterInsert.SelectCommand.ExecuteNonQuery()
                            Refresh_DGV_Ruby()
                            Refresh_DGV_Ruby2()
                        End If

                        DCRuby.Text = "-"
                        ScanRuby.Enabled = True
                        ScanRuby.Text = ""

                    Else
                        MessageBox.Show("This PP not Ruby")
                    End If
                Else
                    MsgBox("Sorry The PP must be number")
                    Me.PPFujiEntry.Text = ""
                End If
            Else
                MsgBox("Sorry you Must select workstations and technician first")
            End If
        End If
    End Sub

    Private Sub Refresh_DGV_Ruby()
        DataGridView9.Rows.Clear()
        Dim sql As String = "select CR.[Order],CR.[RefRuby],CR.[Check Components],CR.[ScanAgain], p.[Item quantity], CR.[process] 
        from ComponentsRuby CR, PPList p 
        where CR.[Order]='" & Me.PPRubyEntry.Text & "' and CR.Workstation='" & Me.workstationRuby.Text & "' and CR.Component is null and p.[Order]='" & Me.PPRubyEntry.Text & "'
        and [process] = 0"
        Dim ds As New DataSet
        adapter = New SqlDataAdapter(sql, Main.koneksi)
        adapter.Fill(ds)

        Dim maxRuby As String = "select Max([process]) as process 
        from ComponentsRuby where [Order]='" & Me.PPRubyEntry.Text & "' and Workstation='" & Me.workstationRuby.Text & "'"
        Dim dsMax As New DataSet
        adapterMax = New SqlDataAdapter(maxRuby, Main.koneksi)
        adapterMax.Fill(dsMax)

        QTY_Process.Text = dsMax.Tables(0).Rows(0).Item("Process").ToString()

        If ds.Tables(0).Rows.Count = 0 Then
            Dim sqlIQ As String = "select [Item quantity] from PPList where [Order]='" & Me.PPRubyEntry.Text & "'"
            Dim dsIQ As New DataSet
            adapterIQ = New SqlDataAdapter(sqlIQ, Main.koneksi)
            adapterIQ.Fill(dsIQ)
            If dsMax.Tables(0).Rows(0).Item("Process") < dsIQ.Tables(0).Rows(0).Item("Item quantity") Then
                Dim queryInsert = "insert into [ComponentsRuby]([order],[RefRuby],[WorkStation], 
                [Qty]) values ('" & Me.PPRubyEntry.Text & "',(select Top 1 material from PPList 
                where [order]='" & Me.PPRubyEntry.Text & "'),'" & Me.workstationRuby.Text & "','" & QtyPerfLabel.Text & "')"
                Dim adapterInsert = New SqlDataAdapter(queryInsert, Main.koneksi)
                adapterInsert.SelectCommand.ExecuteNonQuery()
                Refresh_DGV_Ruby()
                Exit Sub
            End If

            Dim sql2 As String = "select CR.[Order],CR.[RefRuby],CR.[Check Components],CR.[ScanAgain], p.[Item quantity], CR.[process] 
            from ComponentsRuby CR, PPList p 
            where CR.[Order]='" & Me.PPRubyEntry.Text & "' and CR.Workstation='" & Me.workstationRuby.Text & "' and CR.Component is null and p.[Order]='" & Me.PPRubyEntry.Text & "'
            and [process] = " & dsMax.Tables(0).Rows(0).Item("Process").ToString()
            Dim ds2 As New DataSet
            adapter2 = New SqlDataAdapter(sql2, Main.koneksi)
            adapter2.Fill(ds2)
            QTY_Ruby.Text = ds2.Tables(0).Rows(0).Item("Item quantity").ToString()

            If ds2.Tables(0).Rows.Count > 0 Then
                DataGridView9.ColumnCount = 6
                DataGridView9.Columns(0).Name = "Order"
                DataGridView9.Columns(1).Name = "Com Ref"
                DataGridView9.Columns(2).Name = "QTY"
                DataGridView9.Columns(3).Name = "Process QTY"
                DataGridView9.Columns(4).Name = "Scan Product"
                DataGridView9.Columns(5).Name = "Scan Again"
                For r = 0 To ds2.Tables(0).Rows.Count - 1
                    Dim row As String() = New String() {
                        ds2.Tables(0).Rows(r).Item("Order").ToString(),
                        ds2.Tables(0).Rows(r).Item("RefRuby").ToString(),
                        ds2.Tables(0).Rows(r).Item("Item quantity").ToString(),
                        dsMax.Tables(0).Rows(0).Item("Process").ToString(),
                        ds2.Tables(0).Rows(r).Item("Check Components").ToString(),
                        ds2.Tables(0).Rows(r).Item("ScanAgain").ToString()
                    }
                    DataGridView9.Rows.Add(row)

                    If Convert.ToInt32(ds2.Tables(0).Rows(r).Item("Check Components").ToString()) = 1 Then
                        DataGridView9.Rows(r).Cells(4).Style.BackColor = Color.SkyBlue
                    Else
                        DataGridView9.Rows(r).Cells(4).Style.BackColor = Color.White
                    End If

                    If Convert.ToInt32(ds2.Tables(0).Rows(r).Item("ScanAgain").ToString()) Mod 2 = 0 Then
                        DataGridView9.Rows(r).Cells(5).Style.BackColor = Color.SkyBlue
                    Else
                        DataGridView9.Rows(r).Cells(5).Style.BackColor = Color.White
                    End If

                Next
            End If
        Else
            QTY_Ruby.Text = ds.Tables(0).Rows(0).Item("Item quantity").ToString()

            If ds.Tables(0).Rows.Count > 0 Then
                DataGridView9.ColumnCount = 6
                DataGridView9.Columns(0).Name = "Order"
                DataGridView9.Columns(1).Name = "Com Ref"
                DataGridView9.Columns(2).Name = "QTY"
                DataGridView9.Columns(3).Name = "Process QTY"
                DataGridView9.Columns(4).Name = "Scan Product"
                DataGridView9.Columns(5).Name = "Scan Again"
                For r = 0 To ds.Tables(0).Rows.Count - 1
                    Dim row As String() = New String() {
                        ds.Tables(0).Rows(r).Item("Order").ToString(),
                        ds.Tables(0).Rows(r).Item("RefRuby").ToString(),
                        ds.Tables(0).Rows(r).Item("Item quantity").ToString(),
                        dsMax.Tables(0).Rows(0).Item("Process").ToString(),
                        ds.Tables(0).Rows(r).Item("Check Components").ToString(),
                        ds.Tables(0).Rows(r).Item("ScanAgain").ToString()
                    }
                    DataGridView9.Rows.Add(row)

                    If Convert.ToInt32(ds.Tables(0).Rows(r).Item("Check Components").ToString()) = 1 Then
                        DataGridView9.Rows(r).Cells(4).Style.BackColor = Color.SkyBlue
                    Else
                        DataGridView9.Rows(r).Cells(4).Style.BackColor = Color.White
                    End If

                    If Convert.ToInt32(ds.Tables(0).Rows(r).Item("ScanAgain").ToString()) Mod 2 = 0 Then
                        DataGridView9.Rows(r).Cells(5).Style.BackColor = Color.SkyBlue
                    Else
                        DataGridView9.Rows(r).Cells(5).Style.BackColor = Color.White
                    End If

                Next
            End If
        End If

        If Convert.ToInt32(QTY_Process.Text) >= Convert.ToInt32(QTY_Ruby.Text) Then
            ScanRuby.Enabled = False
            ScanLabelRuby.Enabled = False
        End If
    End Sub

    '------------------- Refresh DGV Ruby V2 -------------------------------------------------
    Private Sub Refresh_DGV_Ruby2()
        DataGridView10.Rows.Clear()
        Dim sql As String = "select [Order],[RefRuby],[Component],[Check Components],[Desc] from ComponentsRuby where [Order]='" & Me.PPRubyEntry.Text & "' and Workstation='" & Me.workstationRuby.Text & "' and Component is not null and [process] = 0"
        Dim ds As New DataSet
        adapter = New SqlDataAdapter(sql, Main.koneksi)
        adapter.Fill(ds)
        If ds.Tables(0).Rows.Count > 0 Then
            DataGridView10.ColumnCount = 5
            DataGridView10.Columns(0).Name = "Order"
            DataGridView10.Columns(1).Name = "Com Ref"
            DataGridView10.Columns(2).Name = "Component"
            DataGridView10.Columns(3).Name = "Description"
            DataGridView10.Columns(4).Name = "Scan"
            For r = 0 To ds.Tables(0).Rows.Count - 1
                Dim row As String() = New String() {ds.Tables(0).Rows(r).Item("Order").ToString(), ds.Tables(0).Rows(r).Item("RefRuby").ToString(), ds.Tables(0).Rows(r).Item("Component").ToString(), ds.Tables(0).Rows(r).Item("Desc").ToString(), ds.Tables(0).Rows(r).Item("Check Components").ToString()}
                DataGridView10.Rows.Add(row)
                If ds.Tables(0).Rows(r).Item("Check Components").ToString() = 1 Then
                    DataGridView10.Rows(r).Cells(4).Style.BackColor = Color.SkyBlue
                Else
                    DataGridView10.Rows(r).Cells(4).Style.BackColor = Color.White
                End If
            Next
        End If
    End Sub

    Private Sub masterruby_Click(sender As Object, e As EventArgs) Handles masterruby.Click
        OpenFileDialog1.InitialDirectory = My.Computer.FileSystem.SpecialDirectories.MyDocuments
        If OpenFileDialog1.ShowDialog(Me) = System.Windows.Forms.DialogResult.OK Then
            Dim xlApp As New Microsoft.Office.Interop.Excel.Application
            Dim xlWorkBook As Microsoft.Office.Interop.Excel.Workbook = xlApp.Workbooks.Open(OpenFileDialog1.FileName)
            Dim SheetName As String = xlWorkBook.Worksheets(1).Name.ToString
            Dim excelpath As String = OpenFileDialog1.FileName
            Dim koneksiExcel As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & excelpath & ";Extended Properties='Excel 8.0;HDR=No;IMEX=1;'"
            oleCon = New OleDbConnection(koneksiExcel)
            oleCon.Open()

            Dim queryExcel As String = "select * from [" & SheetName & "$]"
            Dim cmd As OleDbCommand = New OleDbCommand(queryExcel, oleCon)
            Dim rd As OleDbDataReader

            Call koneksi_db()

            Dim deleteReset As New SqlCommand("Delete from [dbo].[MasterRuby]", Main.koneksi)

            Try
                deleteReset.ExecuteNonQuery()
            Catch ex As Exception
                MsgBox("Delete DB Fail " & ex.Message)
            End Try

            Using bulkCopy As SqlBulkCopy = New SqlBulkCopy(Main.koneksi)
                bulkCopy.DestinationTableName = "dbo.MasterRuby"
                Try
                    rd = cmd.ExecuteReader
                    bulkCopy.WriteToServer(rd)

                    rd.Close()
                    MsgBox("Upload Master Data RUBY Successed !")
                Catch ex As Exception
                    MsgBox("Upload Master RUBY Fail" & ex.Message)
                End Try
            End Using
        End If
    End Sub

    Private Sub Button14_Click(sender As Object, e As EventArgs) Handles Button14.Click
        OpenFileDialog1.InitialDirectory = My.Computer.FileSystem.SpecialDirectories.MyDocuments
        If OpenFileDialog1.ShowDialog(Me) = System.Windows.Forms.DialogResult.OK Then
            Dim xlApp As New Microsoft.Office.Interop.Excel.Application
            Dim xlWorkBook As Microsoft.Office.Interop.Excel.Workbook = xlApp.Workbooks.Open(OpenFileDialog1.FileName)
            Dim SheetName As String = xlWorkBook.Worksheets(1).Name.ToString
            Dim excelpath As String = OpenFileDialog1.FileName
            Dim koneksiExcel As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & excelpath & ";Extended Properties='Excel 8.0;HDR=No;IMEX=1;'"
            oleCon = New OleDbConnection(koneksiExcel)
            oleCon.Open()

            Dim queryExcel As String = "select * from [" & SheetName & "$]"
            Dim cmd As OleDbCommand = New OleDbCommand(queryExcel, oleCon)
            Dim rd As OleDbDataReader

            Call koneksi_db()

            Dim deleteReset As New SqlCommand("Delete from [dbo].[MasterMaterialRuby]", Main.koneksi)

            Try
                deleteReset.ExecuteNonQuery()
            Catch ex As Exception
                MsgBox("Delete DB Fail " & ex.Message)
            End Try

            Using bulkCopy As SqlBulkCopy = New SqlBulkCopy(Main.koneksi)
                bulkCopy.DestinationTableName = "dbo.MasterMaterialRuby"
                Try
                    rd = cmd.ExecuteReader
                    bulkCopy.WriteToServer(rd)

                    rd.Close()
                    MsgBox("Upload Master Material RUBY Successed !")
                Catch ex As Exception
                    MsgBox("Upload Master Material Fail" & ex.Message)
                End Try
            End Using
        End If
    End Sub

    Private Sub Button15_Click(sender As Object, e As EventArgs) Handles Button15.Click
        OpenFileDialog1.InitialDirectory = My.Computer.FileSystem.SpecialDirectories.MyDocuments
        If OpenFileDialog1.ShowDialog(Me) = System.Windows.Forms.DialogResult.OK Then
            Dim xlApp As New Microsoft.Office.Interop.Excel.Application
            Dim xlWorkBook As Microsoft.Office.Interop.Excel.Workbook = xlApp.Workbooks.Open(OpenFileDialog1.FileName)
            Dim SheetName As String = xlWorkBook.Worksheets(1).Name.ToString
            Dim excelpath As String = OpenFileDialog1.FileName
            Dim koneksiExcel As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & excelpath & ";Extended Properties='Excel 8.0;HDR=No;IMEX=1;'"
            oleCon = New OleDbConnection(koneksiExcel)
            oleCon.Open()

            Dim queryExcel As String = "select * from [" & SheetName & "$]"
            Dim cmd As OleDbCommand = New OleDbCommand(queryExcel, oleCon)
            Dim rd As OleDbDataReader

            Call koneksi_db()

            Dim deleteReset As New SqlCommand("Delete from [dbo].[MasterRangeLabelRuby]", Main.koneksi)

            Try
                deleteReset.ExecuteNonQuery()
            Catch ex As Exception
                MsgBox("Delete DB Fail " & ex.Message)
            End Try

            Using bulkCopy As SqlBulkCopy = New SqlBulkCopy(Main.koneksi)
                bulkCopy.DestinationTableName = "dbo.MasterRangeLabelRuby"
                Try
                    rd = cmd.ExecuteReader
                    bulkCopy.WriteToServer(rd)

                    rd.Close()
                    MsgBox("Upload Master Range Label RUBY Successed !")
                Catch ex As Exception
                    MsgBox("Upload Master Range Label Fail" & ex.Message)
                End Try
            End Using
        End If
    End Sub

    Private Sub Button16_Click(sender As Object, e As EventArgs) Handles Button16.Click
        OpenFileDialog1.InitialDirectory = My.Computer.FileSystem.SpecialDirectories.MyDocuments
        If OpenFileDialog1.ShowDialog(Me) = System.Windows.Forms.DialogResult.OK Then
            Dim xlApp As New Microsoft.Office.Interop.Excel.Application
            Dim xlWorkBook As Microsoft.Office.Interop.Excel.Workbook = xlApp.Workbooks.Open(OpenFileDialog1.FileName)
            Dim SheetName As String = xlWorkBook.Worksheets(1).Name.ToString
            Dim excelpath As String = OpenFileDialog1.FileName
            Dim koneksiExcel As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & excelpath & ";Extended Properties='Excel 8.0;HDR=No;IMEX=1;'"
            oleCon = New OleDbConnection(koneksiExcel)
            oleCon.Open()

            Dim queryExcel As String = "select * from [" & SheetName & "$]"
            Dim cmd As OleDbCommand = New OleDbCommand(queryExcel, oleCon)
            Dim rd As OleDbDataReader

            Call koneksi_db()

            Dim deleteReset As New SqlCommand("Delete from [dbo].[MasterPackagingLabelRuby]", Main.koneksi)

            Try
                deleteReset.ExecuteNonQuery()
            Catch ex As Exception
                MsgBox("Delete DB Fail " & ex.Message)
            End Try

            Using bulkCopy As SqlBulkCopy = New SqlBulkCopy(Main.koneksi)
                bulkCopy.DestinationTableName = "dbo.MasterPackagingLabelRuby"
                Try
                    rd = cmd.ExecuteReader
                    bulkCopy.WriteToServer(rd)

                    rd.Close()
                    MsgBox("Upload Master Packaging Label RUBY Successed !")
                Catch ex As Exception
                    MsgBox("Upload Master Packaging Label Fail" & ex.Message)
                End Try
            End Using
        End If
    End Sub



    Sub exeption()
        Dim DataSn As String

        If Len(ScanRuby.Text) >= 52 Then
            DataSn = ScanRuby.Text().Substring(39, 13)
            DCRuby.Text = DataSn.Substring(2)
            YearDCRuby.Text = DataSn.Substring(2, 2)
            WeekDCRuby.Text = DataSn.Substring(4, 2)
            DayDCRuby.Text = DataSn.Substring(6, 1)
            IDNumberRuby.Text = DataSn.Substring(9, 4)
            QRCodeRuby.Text = ScanRuby.Text
        ElseIf Len(ScanRuby.Text) < 70 Then
            If ScanRuby.Text.Length = 34 Then
                DataSn = ScanRuby.Text().Substring(0, 13)
                QRCodeRuby.Text = "http://go2se.com/ref=" + ComRefRuby.Text + "/sn=" + ScanRuby.Text.Substring(0, 13)
                txtCXMatrix.Text = ScanRuby.Text.Substring(13, 21)
            End If

            If Not ScanRuby.Text.Contains("http://go2se.com/ref=") Then
                DCRuby.Text = DataSn.Substring(2)
                YearDCRuby.Text = DataSn.Substring(2, 2)
                WeekDCRuby.Text = DataSn.Substring(4, 2)
                DayDCRuby.Text = DataSn.Substring(6, 1)
                IDNumberRuby.Text = DataSn.Substring(9, 4)
            End If
        End If
    End Sub


    Private Sub ScanRuby_TextChanged(sender As Object, e As PreviewKeyDownEventArgs) Handles ScanRuby.PreviewKeyDown
        If (e.KeyData = Keys.Tab Or e.KeyData = Keys.Enter) And (Len(Me.ScanRuby.Text) >= 34) Then

            If Convert.ToInt32(QTY_Process.Text) < Convert.ToInt32(QTY_Ruby.Text) Then
                Dim QRParse As String
                Dim ComRefParse As String

                Dim dtset As New DataSet
                Dim sqlMaxRecord As String = "SELECT max([RecordRuby]) as maxRuby FROM [printingRecord] WHERE [date] = '" & DateTime.Now.ToString("yyyy-MM-dd") & "'"
                Dim adapt = New SqlDataAdapter(sqlMaxRecord, Main.koneksi)
                adapt.Fill(dtset)
                If IsDBNull(dtset.Tables(0).Rows(0).Item("maxRuby")) Then
                    countRuby = 1
                Else
                    countRuby = dtset.Tables(0).Rows(0).Item("maxRuby") + 1
                End If

                'If Len(ScanRuby.Text) >= 70 And TextBox5.Text = 1 Then
                If Len(ScanRuby.Text) >= 52 And TextBox5.Text = 1 Then
                    ComRefParse = ScanRuby.Text().Substring(21, 14)
                    If ComRefParse = ComRefRuby.Text Then
                        'QRParse = ScanRuby.Text().Substring(52, 21)
                        QRParse = txtCXMatrix.Text
                    End If
                ElseIf Len(ScanRuby.Text) < 70 And TextBox5.Text = 0 Then
                    QRParse = ScanRuby.Text().Substring(13)
                    Dim sqlCekprint As String = "select * from printingRecord where [QRCodeRuby] like '%" & ScanRuby.Text().Substring(0, 13) & "%'"
                    Dim dsCekprint As New DataSet
                    adapterCheckprint = New SqlDataAdapter(sqlCekprint, Main.koneksi)
                    adapterCheckprint.Fill(dsCekprint)
                    If dsCekprint.Tables(0).Rows.Count > 0 Then
                        ScanRuby.Text = ""
                        MessageBox.Show("This QR have been Scanned.")
                        Exit Sub
                    End If
                End If

                Dim sqlGetSERef As String = "select SERef from CRMetrix where [QRCodeRuby]='" & QRParse & "'"
                Dim dsGetSERef As New DataSet
                adapterGetSERef = New SqlDataAdapter(sqlGetSERef, Main.koneksi)
                adapterGetSERef.Fill(dsGetSERef)
                If dsGetSERef.Tables(0).Rows.Count = 0 Then
                    ScanRuby.Text = ""
                    MessageBox.Show("Wrong Breaker!")
                    Exit Sub
                End If

                Dim SERefDB As String = dsGetSERef.Tables(0).Rows(0).Item("SERef")

                If SERefDB = SERefRuby.Text Then
                    Dim sqlGetDataRef As String = "select DISTINCT peggedreqt from openOrders where [order]='" & PPRubyEntry.Text & "'"
                    Dim dsGetDataRef As New DataSet
                    adapterGetDataRef = New SqlDataAdapter(sqlGetDataRef, Main.koneksi)
                    adapterGetDataRef.Fill(dsGetDataRef)

                    Dim DataRef As String = dsGetDataRef.Tables(0).Rows(0).Item("peggedreqt")

                    Dim DataSn As String = ScanRuby.Text().Substring(0, 13)


                    Dim sqlCekRuby As String = "select * from ComponentsRuby where [order]=" & PPRubyEntry.Text & "and [RefRuby]='" & DataRef & "' and workstation='" & workstationRuby.Text & "' and Component is null and [process] = 0"
                    Dim dsCekRuby As New DataSet
                    adapter = New SqlDataAdapter(sqlCekRuby, Main.koneksi)
                    adapter.Fill(dsCekRuby)
                    If (dsCekRuby.Tables(0).Rows.Count > 0) Then

                        'ambil QRCODE
                        If ScanRuby.Text.Length = 34 Then
                            'QRCodeRuby.Text = "http://go2se.com/ref=" + ComRefRuby.Text + "/sn=" + ScanRuby.Text ' + "07" + countRuby.ToString("D4") '+ "07"
                            QRCodeRuby.Text = "http://go2se.com/ref=" + ComRefRuby.Text + "/sn=" + ScanRuby.Text.Substring(0, 13)
                            txtCXMatrix.Text = ScanRuby.Text.Substring(13, 21)
                        End If

                        'If ScanRuby.Text.Length = 79 Then
                        If ScanRuby.Text.Length = 73 Then
                            'QRCodeRuby.Text = "http://go2se.com/ref=" + ComRefRuby.Text + "/sn=" + ScanRuby.Text.Substring(39, 34) '+ "07" + countRuby.ToString("D4") '+ "07"
                            QRCodeRuby.Text = "http://go2se.com/ref=" + ComRefRuby.Text + "/sn=" + ScanRuby.Text.Substring(39, 13)
                        End If

                        If Not ScanRuby.Text.Contains("http://go2se.com/ref=") Then
                            DCRuby.Text = DataSn.Substring(2)
                            YearDCRuby.Text = DataSn.Substring(2, 2)
                            WeekDCRuby.Text = DataSn.Substring(4, 2)
                            DayDCRuby.Text = DataSn.Substring(6, 1)
                            IDNumberRuby.Text = DataSn.Substring(9, 4)
                        End If

                        If (dsCekRuby.Tables(0).Rows(0).Item("Check Components") = 1 And dsCekRuby.Tables(0).Rows(0).Item("ScanAgain") = dsCekRuby.Tables(0).Rows(0).Item("Qty") And Convert.ToInt32(QTY_Process.Text) >= Convert.ToInt32(QTY_Ruby.Text)) Then
                            MsgBox("All Product have been scanned")
                            ScanRuby.Text = ""
                            ScanLabelRuby.Text = ""
                            ScanLabelRuby.Enabled = True
                            ScanRuby.Enabled = False
                        Else
                            If dsCekRuby.Tables(0).Rows(0).Item("Check Components") = 1 And dsCekRuby.Tables(0).Rows(0).Item("ScanAgain") = 0 Then
                                Dim queryUpdate = "UPDATE [ComponentsRuby] SET [ScanAgain] = " & Convert.ToInt32(dsCekRuby.Tables(0).Rows(0).Item("ScanAgain")) + 1 & " 
                                where [Order]='" & Me.PPRubyEntry.Text & "' 
                                and Workstation='" & Me.workstationRuby.Text & "' 
                                and RefRuby='" & DataRef & "' 
                                and Component is null and [process]=0"
                                adapter = New SqlDataAdapter(queryUpdate, Main.koneksi)
                                If adapter.SelectCommand.ExecuteNonQuery().ToString() = 1 Then
                                    Refresh_DGV_Ruby()
                                    ScanRuby.Text = ""
                                    ScanRuby.Select()
                                    If dsCekRuby.Tables(0).Rows(0).Item("Qty") = 1 Then
                                        Dim queryInsert = "INSERT INTO [ComponentsRuby] ([Order],[RefRuby],[Component],[Qty],[workstation],[Desc]) 
                                    values ( " & Me.PPRubyEntry.Text & ",'" & DataRef & "', (select screwacc from masterruby where CRRef = '" & DataRef & "'), 1,'" & workstationRuby.Text & "','Screw Acc')"
                                        Dim adapterInsert = New SqlDataAdapter(queryInsert, Main.koneksi)
                                        If adapterInsert.SelectCommand.ExecuteNonQuery() > 0 Then
                                            Refresh_DGV_Ruby2()
                                            ScanRuby.Text = ""
                                            ScanLabelRuby.Text = ""
                                            ScanLabelRuby.Enabled = True
                                            ScanRuby.Enabled = False
                                            ScanLabelRuby.Select()
                                        End If
                                    Else
                                        CheckBoxNR.Checked = True
                                        CheckBoxNL.Checked = False
                                        PrintPerformanceRuby_Click(sender, e)
                                    End If
                                End If
                            ElseIf dsCekRuby.Tables(0).Rows(0).Item("Check Components") = 1 And dsCekRuby.Tables(0).Rows(0).Item("ScanAgain") = 1 Then
                                Dim queryUpdate = "UPDATE [ComponentsRuby] SET [ScanAgain] = " & Convert.ToInt32(dsCekRuby.Tables(0).Rows(0).Item("ScanAgain")) + 1 & " 
                                where [Order]='" & Me.PPRubyEntry.Text & "' 
                                and Workstation='" & Me.workstationRuby.Text & "' 
                                and RefRuby='" & DataRef & "' 
                                and Component is null and [process]=0"
                                adapter = New SqlDataAdapter(queryUpdate, Main.koneksi)
                                If adapter.SelectCommand.ExecuteNonQuery().ToString() = 1 Then
                                    Refresh_DGV_Ruby()
                                    ScanRuby.Text = ""
                                    ScanRuby.Select()
                                    Dim queryInsert = "INSERT INTO [ComponentsRuby] ([Order],[RefRuby],[Component],[Qty],[workstation],[Desc]) 
                                values ( " & Me.PPRubyEntry.Text & ",'" & DataRef & "', (select screwacc from masterruby where CRRef = '" & DataRef & "'), 1,'" & workstationRuby.Text & "','Screw Acc')"
                                    Dim adapterInsert = New SqlDataAdapter(queryInsert, Main.koneksi)

                                    If adapterInsert.SelectCommand.ExecuteNonQuery() > 0 Then
                                        Refresh_DGV_Ruby2()
                                        ScanRuby.Text = ""
                                        ScanLabelRuby.Text = ""
                                        ScanLabelRuby.Enabled = True
                                        ScanRuby.Enabled = False
                                        ScanLabelRuby.Select()
                                    End If
                                End If
                                'ElseIf dsCekRuby.Tables(0).Rows(0).Item("Check Components") = 1 And dsCekRuby.Tables(0).Rows(0).Item("ScanAgain") = 1 And dsCekRuby.Tables(0).Rows(0).Item("ScanAgain") <> dsCekRuby.Tables(0).Rows(0).Item("Qty") Then
                                '    Dim queryUpdate = "UPDATE [ComponentsRuby] SET [Check Components] = " & Convert.ToInt32(dsCekRuby.Tables(0).Rows(0).Item("Check Components")) + 1 & "
                                '    where [Order]='" & Me.PPRubyEntry.Text & "' 
                                '    and Workstation='" & Me.workstationRuby.Text & "' 
                                '    and RefRuby='" & DataRef & "' 
                                '    and Component is null"
                                '    adapter = New SqlDataAdapter(queryUpdate, Main.koneksi)
                                '    If adapter.SelectCommand.ExecuteNonQuery().ToString() = 1 Then
                                '        ' MessageBox.Show("Print " & QtyPerfLabel.Text & " Perf 1 Label")
                                '        'print ruby performance
                                '        CheckBoxNL.Checked = False
                                '        CheckBoxNR.Checked = True
                                '        PrintPerformanceRuby_Click(sender, e)

                                '        Refresh_DGV_Ruby()
                                '        ScanRuby.Text = ""
                                '        ScanRuby.Select()
                                '    End If
                            Else
                                Dim queryUpdate = "UPDATE [ComponentsRuby] SET [Check Components] = 1
                                where [Order]='" & Me.PPRubyEntry.Text & "' 
                                and Workstation='" & Me.workstationRuby.Text & "' 
                                and RefRuby='" & DataRef & "' 
                                and Component is null and [process] = 0"
                                adapter = New SqlDataAdapter(queryUpdate, Main.koneksi)
                                If adapter.SelectCommand.ExecuteNonQuery().ToString() = 1 Then

                                    'MessageBox.Show("Print " & QtyPerfLabel.Text & " Perf 2 Label")
                                    'print ruby performance
                                    CheckBoxNR.Checked = False
                                    CheckBoxNL.Checked = True
                                    PrintPerformanceRuby_Click(sender, e)

                                    Refresh_DGV_Ruby()
                                    ScanRuby.Text = ""
                                    ScanRuby.Select()

                                    TextBox5.Text = 1
                                End If
                            End If
                        End If
                    End If
                Else
                    MessageBox.Show("Wrong QR Code")
                End If
            Else
                MessageBox.Show("This Breaker Already Scan")
                exeption()
                'ScanRuby.Text = ""
            End If
        End If
    End Sub

    Private Sub ScanLabelRuby_TextChanged(sender As Object, e As PreviewKeyDownEventArgs) Handles ScanLabelRuby.PreviewKeyDown
        If (e.KeyData = Keys.Tab Or e.KeyData = Keys.Enter) Then
            'add scan QR Code for acessories
            Dim accesories_QR As String
            If ScanLabelRuby.Text.Length >= 33 Then
                accesories_QR = Microsoft.VisualBasic.Left(ScanLabelRuby.Text, 5)
                If accesories_QR = "G4332" Then
                    ScanLabelRuby.Text = "TPCDIO15"
                End If

                If accesories_QR = "G4333" Then
                    ScanLabelRuby.Text = "TPCCOM16"
                End If
            End If


            Dim sqlCekRuby As String = "select * from MasterRuby where [CRRef]='" & ComRefRuby.Text & "'"
            Dim dsCekRuby As New DataSet
            adapter = New SqlDataAdapter(sqlCekRuby, Main.koneksi)
            adapter.Fill(dsCekRuby)

            Dim sqlCekComponentsRuby As String = "select * from ComponentsRuby where [order]=" & PPRubyEntry.Text & "and [RefRuby]='" & ComRefRuby.Text & "' and workstation='" & workstationRuby.Text & "' and Component ='" & ScanLabelRuby.Text & "' and [check components] = 0"
            Dim dsCekComponentsRuby As New DataSet
            adapter = New SqlDataAdapter(sqlCekComponentsRuby, Main.koneksi)
            adapter.Fill(dsCekComponentsRuby)
            If (dsCekComponentsRuby.Tables(0).Rows.Count > 0) Then
                If dsCekRuby.Tables(0).Rows(0).Item("ScrewAcc").ToString() = ScanLabelRuby.Text Then
                    Dim queryUpdateScrew = "UPDATE [ComponentsRuby] SET [check components] = 1 
                    where [Order]='" & Me.PPRubyEntry.Text & "' and refruby='" & ComRefRuby.Text & "' 
                    and workstation='" & workstationRuby.Text & "' and component = '" & ScanLabelRuby.Text & "'"
                    Dim adapterUpdateScrew = New SqlDataAdapter(queryUpdateScrew, Main.koneksi)
                    If adapterUpdateScrew.SelectCommand.ExecuteNonQuery().ToString() > 0 Then
                        Dim queryInsert = "INSERT INTO [ComponentsRuby] 
	                    ([Order],[RefRuby],[Component],[Qty],[workstation],[desc]) 
                        values 
	                    ( '" & Me.PPRubyEntry.Text & "','" & ComRefRuby.Text & "', (select ext24 from masterruby where CRRef = '" & ComRefRuby.Text & "'), 1,'" & workstationRuby.Text & "',(select descr from openorders where material = (select ext24 from masterruby where CRRef = '" & ComRefRuby.Text & "') and [order]='" & PPRubyEntry.Text & "')),
	                    ( '" & Me.PPRubyEntry.Text & "','" & ComRefRuby.Text & "', (select modrs485 from masterruby where CRRef = '" & ComRefRuby.Text & "'), 1,'" & workstationRuby.Text & "',(select descr from openorders where material = (select modrs485 from masterruby where CRRef = '" & ComRefRuby.Text & "') and [order]='" & PPRubyEntry.Text & "')),
	                    ( '" & Me.PPRubyEntry.Text & "','" & ComRefRuby.Text & "', (select ean13 from MasterPackagingLabelRuby where CRRef = '" & ComRefRuby.Text & "'), 1,'" & workstationRuby.Text & "','Packaging Label')"
                        Dim adapterInsert = New SqlDataAdapter(queryInsert, Main.koneksi)
                        If adapterInsert.SelectCommand.ExecuteNonQuery() > 0 Then
                            Refresh_DGV_Ruby2()
                            'MessageBox.Show("Print Single Box")
                            PrintPackagingRuby_Click(sender, e)

                            ScanLabelRuby.Text = ""
                            ScanLabelRuby.Select()
                        End If
                    End If
                Else
                    Dim queryUpdateScrew = "UPDATE [ComponentsRuby] SET [check components] = 1 
                    where [Order]='" & Me.PPRubyEntry.Text & "' and refruby='" & ComRefRuby.Text & "' 
                    and workstation='" & workstationRuby.Text & "' and component = '" & ScanLabelRuby.Text & "'"
                    Dim adapterUpdateScrew = New SqlDataAdapter(queryUpdateScrew, Main.koneksi)
                    If adapterUpdateScrew.SelectCommand.ExecuteNonQuery().ToString() > 0 Then
                        Refresh_DGV_Ruby2()
                        If QtyPerfLabel.Text = 2 Then
                            TextBox4.Text = Convert.ToInt32(TextBox4.Text) + 1
                        Else
                            TextBox4.Text = Convert.ToInt32(TextBox4.Text) + 3
                        End If
                        ScanLabelRuby.Text = ""
                        ScanLabelRuby.Select()
                    End If
                End If

                Dim sqlCekComponentsRubyUpdate As String = "select * from ComponentsRuby where [order]=" & PPRubyEntry.Text & "and [RefRuby]='" & ComRefRuby.Text & "' and workstation='" & workstationRuby.Text & "' and [check components] = 0"
                Dim dsCekComponentsRubyUpdate As New DataSet
                adapter = New SqlDataAdapter(sqlCekComponentsRubyUpdate, Main.koneksi)
                adapter.Fill(dsCekComponentsRubyUpdate)
                If (dsCekComponentsRubyUpdate.Tables(0).Rows.Count = 0) Then
                    Dim queryUpdateScrew = "UPDATE [ComponentsRuby] SET [process] = " & Convert.ToInt32(QTY_Process.Text) + 1 & " 
                    where [Order]='" & Me.PPRubyEntry.Text & "' and refruby='" & ComRefRuby.Text & "' 
                    and workstation='" & workstationRuby.Text & "' and [process]=0"
                    Dim adapterUpdateScrew = New SqlDataAdapter(queryUpdateScrew, Main.koneksi)
                    If adapterUpdateScrew.SelectCommand.ExecuteNonQuery().ToString() > 0 Then
                        Dim sqlMax As String = "select MAX([process]) as process 
                        from ComponentsRuby
                        where [Order]='" & Me.PPRubyEntry.Text & "' and Workstation='" & Me.workstationRuby.Text & "' and [RefRuby]='" & ComRefRuby.Text & "'"
                        Dim dsMax As New DataSet
                        adapterMax = New SqlDataAdapter(sqlMax, Main.koneksi)
                        adapterMax.Fill(dsMax)
                        If Convert.ToInt32(QTY_Ruby.Text) <= Convert.ToInt32(dsMax.Tables(0).Rows(0).Item("process")) Then
                            TextBox4.Text = 5
                        Else
                            loopRuby()
                        End If
                    End If
                End If
            End If
        End If
    End Sub

    Private Sub TextBox4_TextChanged(sender As Object, e As EventArgs) Handles TextBox4.TextChanged
        If TextBox4.Text = 3 Then
            'MessageBox.Show("Print Group Box")
            PrintOutsideRuby_Click(sender, e)

            'clearAllRuby()
            'Dim queryInsert = "INSERT INTO [ComponentsRuby] 
            ' ([Order],[RefRuby],[Component],[Qty],[workstation]) 
            'values 
            ' ('" & Me.PPRubyEntry.Text & "','" & ComRefRuby.Text & "', (select cartonbox from masterruby where CRRef = '" & ComRefRuby.Text & "'), 1,'" & workstationRuby.Text & "')"
            'Dim adapterInsert = New SqlDataAdapter(queryInsert, Main.koneksi)
            'If adapterInsert.SelectCommand.ExecuteNonQuery() > 0 Then
            '    Refresh_DGV_Ruby2()
            'End If
        ElseIf TextBox4.Text = 5 Then
            PrintOutsideRuby_Click(sender, e)
            clearAllRuby()
        End If
    End Sub

    Sub loopRuby()
        ScanLabelRuby.Text = ""
        ScanRuby.Text = ""
        ScanLabelRuby.Enabled = False
        ScanRuby.Enabled = True
        ScanRuby.Select()
        TextBox5.Text = 0
        TextBox4.Text = 0

        Dim queryInsert = "insert into [ComponentsRuby]([order],[RefRuby],[WorkStation], 
        [Qty]) values ('" & Me.PPRubyEntry.Text & "',(select Top 1 material from PPList 
        where [order]='" & Me.PPRubyEntry.Text & "'),'" & Me.workstationRuby.Text & "','" & QtyPerfLabel.Text & "')"
        Dim adapterInsert = New SqlDataAdapter(queryInsert, Main.koneksi)
        adapterInsert.SelectCommand.ExecuteNonQuery()
        Refresh_DGV_Ruby()
        Refresh_DGV_Ruby2()
    End Sub

    Sub clearAllRuby()
        TextBox5.Text = 0
        TextBox4.Text = 0
        PPRubyEntry.Text = ""
        ScanLabelRuby.Text = ""
        ScanRuby.Text = ""
        ScanLabelRuby.Enabled = False
        ScanRuby.Enabled = False
        PPRubyEntry.Enabled = True
        ComRefRuby.Text = ""
        DataGridView10.Rows.Clear()
        DataGridView9.Rows.Clear()
        PPRubyEntry.Select()

        RatedCurrRuby.Text = ""
        RatedVoltRuby.Text = ""
        UtilizationRuby.Text = ""
        RatedFreq.Text = ""
        ShortCircuitRuby.Text = ""

        PerfLabel.Text = ""
        QtyPerfLabel.Text = ""
        SERefRuby.Text = ""
        QtySERefRuby.Text = ""
        ScrewAcc.Text = ""
        QtyScrewAcc.Text = ""
        EXT24.Text = ""
        QtyEXT24.Text = ""
        RS485.Text = ""
        QtyRS485.Text = ""
        CartonBoxRuby.Text = ""
        QtyCartonBox.Text = ""

        LangCn.Text = ""
        LangEn.Text = ""
        LangEs.Text = ""
        LangFr.Text = ""
        LangKz.Text = ""
        LangRu.Text = ""
        RangeNameRuby.Text = ""
        EAN13Ruby.Text = ""
        QTYSingleBoxRuby.Text = ""
        MadeInCountryRuby.Text = ""
        SiteAddressRuby.Text = ""
        ZipCodeRuby.Text = ""
        NameChRuby.Text = ""
        AddressChRuby.Text = ""
        LogisticRefRuby.Text = ""
        CIPRuby.Text = ""
        ProductImageRuby.Text = ""
        DCRuby.Text = ""
        ProductTypeRuby.Text = ""
        wotpc.Text = ""
        warningRuby.Text = ""
        POGroupBoxRuby.Text = ""
        SoLineRuby.Text = ""
        SORuby.Text = ""
        QRCodeRuby.Text = ""
        YearDCRuby.Text = ""
        WeekDCRuby.Text = ""
        DayDCRuby.Text = ""
        IDNumberRuby.Text = ""
        QTY_Ruby.Text = ""
        QTY_Process.Text = ""
        txtCXMatrix.Text = ""
    End Sub

    Private Sub Ruby_Performance_Small_setvalue(visible As String)
        Try

            'label_performance_small_ruby.Variables("Category").SetValue(Category_Ruby.Text)
            label_performance_small_ruby.Variables("CR").SetValue(ComRefRuby.Text)
            label_performance_small_ruby.Variables("DateCode").SetValue(DCRuby.Text)
            label_performance_small_ruby.Variables("Frequency").SetValue(RatedFreq.Text)
            label_performance_small_ruby.Variables("Icw").SetValue(ShortCircuitRuby.Text)
            label_performance_small_ruby.Variables("Ie").SetValue(RatedCurrRuby.Text)
            label_performance_small_ruby.Variables("QR_Code_Ruby").SetValue(QRCodeRuby.Text)
            label_performance_small_ruby.Variables("Ue").SetValue(RatedVoltRuby.Text)
            label_performance_small_ruby.Variables("Utilization").SetValue(UtilizationRuby.Text)
            label_performance_small_ruby.Variables("Visible").SetValue(visible)

            If Visible_N.Text = "N" Then
                label_performance_small_ruby.Variables("Visible N").SetValue(visible)
            End If

            label_performance_small_ruby.Variables("VisibleBlackL").SetValue("Y")
            label_performance_small_ruby.Variables("VisibleBlackR").SetValue("Y")

            If Phase_N_Logo.Text = "No Printing" Then
                If visible = "R" Then
                    label_performance_small_ruby.Variables("VisibleBlackR").SetValue("")
                ElseIf visible = "L" Then
                    label_performance_small_ruby.Variables("VisibleBlackL").SetValue("")
                End If
            End If

            label_performance_small_ruby.Variables("Automatic").SetValue(wotpc.Text)
            label_performance_small_ruby.Variables("insulation").SetValue(warningRuby.Text)

            label_performance_small_ruby.Variables("Active").SetValue(active_under_QRCode.Text)

        Catch ex As Exception
            MsgBox("Ruby Performance label Small " & ex.Message)
        End Try
    End Sub

    Private Sub Ruby_Performance_Big_setvalue(visible As String)
        Try

            label_performance_big_ruby.Variables("CR").SetValue(ComRefRuby.Text)
            label_performance_big_ruby.Variables("DateCode").SetValue(DCRuby.Text)
            label_performance_big_ruby.Variables("Frequency").SetValue(RatedFreq.Text)
            label_performance_big_ruby.Variables("Icw").SetValue(ShortCircuitRuby.Text)
            label_performance_big_ruby.Variables("Ie").SetValue(RatedCurrRuby.Text)
            label_performance_big_ruby.Variables("QR_Code_Ruby").SetValue(QRCodeRuby.Text)
            label_performance_big_ruby.Variables("Ue").SetValue(RatedVoltRuby.Text)
            label_performance_big_ruby.Variables("Utilization").SetValue(UtilizationRuby.Text)
            label_performance_big_ruby.Variables("Visible").SetValue(visible)

            If Visible_N.Text = "N" Then
                label_performance_big_ruby.Variables("Visible N").SetValue(visible)
            End If

            label_performance_big_ruby.Variables("VisibleBlackL").SetValue("Y")
            label_performance_big_ruby.Variables("VisibleBlackR").SetValue("Y")

            If Phase_N_Logo.Text = "No Printing" Then
                If visible = "R" Then
                    label_performance_big_ruby.Variables("VisibleBlackR").SetValue("")
                ElseIf visible = "L" Then
                    label_performance_big_ruby.Variables("VisibleBlackL").SetValue("")
                End If
            End If

            label_performance_big_ruby.Variables("Automatic").SetValue(wotpc.Text)
            label_performance_big_ruby.Variables("insulation").SetValue(warningRuby.Text)

            label_performance_big_ruby.Variables("Active").SetValue(active_under_QRCode.Text)

        Catch ex As Exception
            MsgBox("Ruby Performance label Big " & ex.Message)
        End Try
    End Sub

    Private Sub Ruby_Packaging_setvalue()
        Try

            label_packaging_ruby.Variables("CR").SetValue(ComRefRuby.Text)
            label_packaging_ruby.Variables("Product_Desc_English").SetValue(LangEn.Text)
            label_packaging_ruby.Variables("Product_Desc_Chinese").SetValue(LangCn.Text)
            label_packaging_ruby.Variables("Product_Desc_France").SetValue(LangFr.Text)
            label_packaging_ruby.Variables("Product_Desc_Espanol").SetValue(LangEs.Text)
            label_packaging_ruby.Variables("Product_Desc_Rusia").SetValue(LangRu.Text)
            label_packaging_ruby.Variables("Product_Desc_Kazakh").SetValue(LangKz.Text)

            label_packaging_ruby.Variables("RangeName").SetValue(RangeNameRuby.Text)
            label_packaging_ruby.Variables("Technical_specification").SetValue(technicianNameRuby.Text)

            label_packaging_ruby.Variables("Barcode").SetValue(EAN13Ruby.Text)
            label_packaging_ruby.Variables("MadeInCountry").SetValue(MadeInCountryRuby.Text)
            label_packaging_ruby.Variables("Factory_Adress").SetValue(SiteAddressRuby.Text)
            label_packaging_ruby.Variables("ZipCode").SetValue(ZipCodeRuby.Text)
            label_packaging_ruby.Variables("NameChRuby").SetValue(NameChRuby.Text)
            label_packaging_ruby.Variables("AddrsChRuby").SetValue(AddressChRuby.Text)
            label_packaging_ruby.Variables("Date_Code").SetValue(DCRuby.Text)
            label_packaging_ruby.Variables("Logistic_Ref").SetValue(LogisticRefRuby.Text)
            label_packaging_ruby.Variables("CIP").SetValue(CIPRuby.Text)

            label_packaging_ruby.Variables("Product_Pic").SetValue(ProductImageRuby.Text)

        Catch ex As Exception
            MsgBox("Ruby Packaging Label  " & ex.Message)
        End Try
    End Sub

    Private Sub Ruby_Outside_setvalue()
        Try

            label_outside_ruby.Variables("CR").SetValue(ComRefRuby.Text)
            label_outside_ruby.Variables("Product_Desc_English").SetValue(LangEn.Text)
            label_outside_ruby.Variables("Product_Desc_Chinese").SetValue(LangCn.Text)
            label_outside_ruby.Variables("Product_Desc_France").SetValue(LangFr.Text)
            label_outside_ruby.Variables("Product_Desc_Espanol").SetValue(LangEs.Text)
            label_outside_ruby.Variables("Product_Desc_Rusia").SetValue(LangRu.Text)
            label_outside_ruby.Variables("Product_Desc_Kazakh").SetValue(LangKz.Text)

            label_outside_ruby.Variables("RangeName").SetValue(RangeNameRuby.Text)
            label_outside_ruby.Variables("Technical_specification").SetValue(technicianNameRuby.Text)

            label_outside_ruby.Variables("Barcode").SetValue(EAN13Ruby.Text)
            label_outside_ruby.Variables("MadeInCountry").SetValue(MadeInCountryRuby.Text)
            label_outside_ruby.Variables("Factory_Adress").SetValue(SiteAddressRuby.Text)
            label_outside_ruby.Variables("ZipCode").SetValue(ZipCodeRuby.Text)
            label_outside_ruby.Variables("NameChRuby").SetValue(NameChRuby.Text)
            label_outside_ruby.Variables("AddrsChRuby").SetValue(AddressChRuby.Text)
            label_outside_ruby.Variables("Date_Code").SetValue(DCRuby.Text)
            label_outside_ruby.Variables("Logistic_Ref").SetValue(LogisticRefRuby.Text)
            label_outside_ruby.Variables("CIP").SetValue(CIPRuby.Text)

            label_outside_ruby.Variables("Product_Pic").SetValue(ProductImageRuby.Text)

            label_outside_ruby.Variables("Po").SetValue(POGroupBoxRuby.Text)
            label_outside_ruby.Variables("So_Line_Number").SetValue(SoLineRuby.Text)
            label_outside_ruby.Variables("SO").SetValue(SORuby.Text)
            label_outside_ruby.Variables("PP").SetValue(PPRubyEntry.Text)

            Dim SOBarcode As String = "0" & SORuby.Text & Convert.ToDecimal(SoLineRuby.Text).ToString("000000")
            label_outside_ruby.Variables("SO_Barcode").SetValue(SOBarcode)
            'label_out_side_printer.Variables("QR_SO").SetValue("0" & FujiSO.Text & Convert.ToDecimal(FujiLineNo.Text).ToString("000000"))

        Catch ex As Exception
            MsgBox("Ruby Outside Label  " & ex.Message)
        End Try
    End Sub

    Private Sub PreviewPerformanceRuby_Click(sender As Object, e As EventArgs) Handles PreviewPerformanceRuby.Click
        reload_printer()
        If CheckBoxNR.Checked Then
            Ruby_Performance_Small_setvalue("R")
            Ruby_Performance_Big_setvalue("R")
        Else
            Ruby_Performance_Small_setvalue("L")
            Ruby_Performance_Big_setvalue("L")
        End If


        'declaration of Preview
        Dim LabelPreviewSettings As ILabelPreviewSettings = New LabelPreviewSettings()

        'setting preview format
        LabelPreviewSettings.ImageFormat = "PNG"
        LabelPreviewSettings.Width = Form_preview.pictureBoxPreview.Width                   ' Width Of image To generate
        LabelPreviewSettings.Height = Form_preview.pictureBoxPreview.Height                 ' Height Of image To generate

        Dim imageObj As Object
        ' Generate Preview File
        If ProductTypeRuby.Text = "Ruby 1" Then
            imageObj = label_performance_small_ruby.GetLabelPreview(LabelPreviewSettings)
        Else
            imageObj = label_performance_big_ruby.GetLabelPreview(LabelPreviewSettings)
        End If


        'Display image in UI
        If TypeOf imageObj Is Byte() Then
            Form_preview.pictureBoxPreview.Image = ByteToImage(CType(imageObj, Byte()))
        ElseIf TypeOf imageObj Is String Then
            Form_preview.pictureBoxPreview.ImageLocation = CStr(imageObj)
        End If

        Form_preview.Show()
    End Sub

    Private Sub PreviewPackagingRuby_Click(sender As Object, e As EventArgs) Handles PreviewPackagingRuby.Click
        reload_printer()
        Ruby_Packaging_setvalue()

        'declaration of Preview
        Dim LabelPreviewSettings As ILabelPreviewSettings = New LabelPreviewSettings()

        'setting preview format
        LabelPreviewSettings.ImageFormat = "PNG"
        LabelPreviewSettings.Width = Form_preview.pictureBoxPreview.Width                   ' Width Of image To generate
        LabelPreviewSettings.Height = Form_preview.pictureBoxPreview.Height                 ' Height Of image To generate

        ' Generate Preview File
        Dim imageObj As Object = label_packaging_ruby.GetLabelPreview(LabelPreviewSettings)

        'Display image in UI
        If TypeOf imageObj Is Byte() Then
            Form_preview.pictureBoxPreview.Image = ByteToImage(CType(imageObj, Byte()))
        ElseIf TypeOf imageObj Is String Then
            Form_preview.pictureBoxPreview.ImageLocation = CStr(imageObj)
        End If

        Form_preview.Show()
    End Sub

    Private Sub PreviewOutsideRuby_Click(sender As Object, e As EventArgs) Handles PreviewOutsideRuby.Click
        reload_printer()
        Ruby_Outside_setvalue()

        'declaration of Preview
        Dim LabelPreviewSettings As ILabelPreviewSettings = New LabelPreviewSettings()

        'setting preview format
        LabelPreviewSettings.ImageFormat = "PNG"
        LabelPreviewSettings.Width = Form_preview.pictureBoxPreview.Width                   ' Width Of image To generate
        LabelPreviewSettings.Height = Form_preview.pictureBoxPreview.Height                 ' Height Of image To generate

        ' Generate Preview File
        Dim imageObj As Object = label_outside_ruby.GetLabelPreview(LabelPreviewSettings)

        'Display image in UI
        If TypeOf imageObj Is Byte() Then
            Form_preview.pictureBoxPreview.Image = ByteToImage(CType(imageObj, Byte()))
        ElseIf TypeOf imageObj Is String Then
            Form_preview.pictureBoxPreview.ImageLocation = CStr(imageObj)
        End If

        Form_preview.Show()
    End Sub

    Private Sub PrintPerformanceRuby_Click(sender As Object, e As EventArgs) Handles PrintPerformanceRuby.Click

        Dim cmd = New SqlCommand("insert into printingRecord([pp],[date],[time],[user],[from],[to],[QRCodeRuby],[Data],[RecordRuby]) values(@pp,@date,@time,@user,@from,@to,@qrcoderuby,@data,@recordruby)", Main.koneksi)
        cmd.Parameters.AddWithValue("@pp", Me.PPRubyEntry.Text)
        cmd.Parameters.AddWithValue("@date", DateTime.Now.ToString("yyyy-MM-dd"))
        cmd.Parameters.AddWithValue("@time", DateTime.Now.ToString("HH:mm:ss"))
        cmd.Parameters.AddWithValue("@user", Me.technicianShortNameRuby.Text)
        cmd.Parameters.AddWithValue("@from", Convert.ToInt32(QTY_Process.Text) + 1)
        cmd.Parameters.AddWithValue("@to", QTY_Ruby.Text)
        cmd.Parameters.AddWithValue("@qrcoderuby", Me.QRCodeRuby.Text)
        cmd.Parameters.AddWithValue("@data", LoginForm.strHostName & " - " & Application.ProductVersion)
        cmd.Parameters.AddWithValue("@recordruby", countRuby)
        cmd.ExecuteNonQuery()

        If DCRuby.Text.Length > 5 Then

            If ProductTypeRuby.Text = "Ruby 1" Then
                If CheckBoxNR.Checked Then
                    Try

                        Ruby_Performance_Small_setvalue("R")
                        label_performance_small_ruby.Print(1)
                    Catch ex As Exception
                        MsgBox("Print Failed" & ex.ToString)
                    End Try
                End If

                If CheckBoxNL.Checked Then
                    Try

                        Ruby_Performance_Small_setvalue("L")
                        label_performance_small_ruby.Print(1)
                    Catch ex As Exception
                        MsgBox("Print Failed" & ex.ToString)
                    End Try
                End If
            ElseIf ProductTypeRuby.Text = "Ruby 2" Then
                If CheckBoxNR.Checked Then
                    Try

                        Ruby_Performance_Big_setvalue("R")
                        label_performance_big_ruby.Print(1)
                    Catch ex As Exception
                        MsgBox("Print Failed" & ex.ToString)
                    End Try
                End If

                If CheckBoxNL.Checked Then
                    Try

                        Ruby_Performance_Big_setvalue("L")
                        label_performance_big_ruby.Print(1)
                    Catch ex As Exception
                        MsgBox("Print Failed" & ex.ToString)
                    End Try
                End If

            Else
                MsgBox("You Need to Scan PP Order First !")
            End If
        Else
            MsgBox("You Need to Scan The Breaker QR-Code First !")
        End If
    End Sub

    Private Sub PrintPackagingRuby_Click(sender As Object, e As EventArgs) Handles PrintPackagingRuby.Click
        If DCRuby.Text.Length > 5 Then
            Try
                Ruby_Packaging_setvalue()
                label_packaging_ruby.Print(1)
            Catch ex As Exception
                MsgBox("Print Failed" & ex.ToString)
            End Try
        Else
            MsgBox("You Need to Scan The Breaker QR-Code First !")
        End If
    End Sub

    Private Sub PrintOutsideRuby_Click(sender As Object, e As EventArgs) Handles PrintOutsideRuby.Click
        If DCRuby.Text.Length > 5 Then
            Try
                Ruby_Outside_setvalue()
                label_outside_ruby.Print(1)
            Catch ex As Exception
                MsgBox("Print Failed" & ex.ToString)
            End Try
        Else
            MsgBox("You Need to Scan The Breaker QR-Code First !")
        End If
    End Sub

    Private Sub workstationRuby_SelectedValueChanged(sender As Object, e As EventArgs) Handles workstationRuby.SelectedValueChanged
        'workstation_event()
    End Sub

    Private Sub Button17_Click(sender As Object, e As EventArgs) Handles Button17.Click
        OpenFileDialog1.InitialDirectory = My.Computer.FileSystem.SpecialDirectories.MyDocuments
        If OpenFileDialog1.ShowDialog(Me) = System.Windows.Forms.DialogResult.OK Then
            Dim xlApp As New Microsoft.Office.Interop.Excel.Application
            Dim xlWorkBook As Microsoft.Office.Interop.Excel.Workbook = xlApp.Workbooks.Open(OpenFileDialog1.FileName)
            Dim SheetName As String = xlWorkBook.Worksheets(1).Name.ToString
            Dim excelpath As String = OpenFileDialog1.FileName
            Dim koneksiExcel As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & excelpath & ";Extended Properties='Excel 8.0;HDR=No;IMEX=1;'"
            oleCon = New OleDbConnection(koneksiExcel)
            oleCon.Open()

            Dim queryExcel As String = "select * from [" & SheetName & "$]"
            Dim cmd As OleDbCommand = New OleDbCommand(queryExcel, oleCon)
            Dim rd As OleDbDataReader

            Call koneksi_db()

            Dim deleteReset As New SqlCommand("Delete from [dbo].[CRMetrix]", Main.koneksi)

            Try
                deleteReset.ExecuteNonQuery()
            Catch ex As Exception
                MsgBox("Delete DB Fail " & ex.Message)
            End Try

            Using bulkCopy As SqlBulkCopy = New SqlBulkCopy(Main.koneksi)
                bulkCopy.DestinationTableName = "dbo.CRMetrix"
                Try
                    rd = cmd.ExecuteReader
                    bulkCopy.WriteToServer(rd)

                    rd.Close()
                    MsgBox("Upload Master CR Metrix Successed !")
                Catch ex As Exception
                    MsgBox("Upload Master CR Metrix Fail" & ex.Message)
                End Try
            End Using
        End If
    End Sub

    Private Sub ProductImageRuby_TextChanged(sender As Object, e As EventArgs) Handles ProductImageRuby.TextChanged
        If ProductImageRuby.Text.Contains("3P") Then
            Visible_N.Text = ""
        Else
            Visible_N.Text = "N"
        End If
    End Sub

    Private Sub cbxPerfomaceRuby_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbxPerfomaceRuby.SelectedIndexChanged
        If printers.Count > 0 Then
            selected_Printer = printers.Item(cbxPerfomaceRuby.SelectedIndex)
            label_performance_small_ruby.PrintSettings.PrinterName = selected_Printer.Name
            label_performance_big_ruby.PrintSettings.PrinterName = selected_Printer.Name
            cbxPerfomaceRuby.SelectedItem = selected_Printer.Name
        End If
    End Sub

    Private Sub cbxPackagingRuby_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbxPackagingRuby.SelectedIndexChanged
        If printers.Count > 0 Then
            selected_Printer = printers.Item(cbxPackagingRuby.SelectedIndex)
            label_packaging_ruby.PrintSettings.PrinterName = selected_Printer.Name
            cbxPackagingRuby.SelectedItem = selected_Printer.Name
        End If
    End Sub

    Private Sub cbxOutsideRuby_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbxOutsideRuby.SelectedIndexChanged
        If printers.Count > 0 Then
            selected_Printer = printers.Item(cbxOutsideRuby.SelectedIndex)
            label_outside_ruby.PrintSettings.PrinterName = selected_Printer.Name
            cbxOutsideRuby.SelectedItem = selected_Printer.Name
        End If
    End Sub

    Private Sub Button18_Click(sender As Object, e As EventArgs) Handles Button18.Click
        ExportToExcel("SELECT * FROM Componentslist")
    End Sub

    Private Sub Button19_Click(sender As Object, e As EventArgs) Handles Button19.Click
        OpenFileDialog1.InitialDirectory = My.Computer.FileSystem.SpecialDirectories.MyDocuments
        If OpenFileDialog1.ShowDialog(Me) = System.Windows.Forms.DialogResult.OK Then
            Dim xlApp As New Microsoft.Office.Interop.Excel.Application
            Dim xlWorkBook As Microsoft.Office.Interop.Excel.Workbook = xlApp.Workbooks.Open(OpenFileDialog1.FileName)
            Dim SheetName As String = xlWorkBook.Worksheets(1).Name.ToString
            Dim excelpath As String = OpenFileDialog1.FileName
            Dim koneksiExcel As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & excelpath & ";Extended Properties='Excel 8.0;HDR=No;IMEX=1;'"
            oleCon = New OleDbConnection(koneksiExcel)
            oleCon.Open()

            Dim queryExcel As String = "select * from [" & SheetName & "$]"
            Dim cmd As OleDbCommand = New OleDbCommand(queryExcel, oleCon)
            Dim rd As OleDbDataReader

            Call koneksi_db()

            Using bulkCopy As SqlBulkCopy = New SqlBulkCopy(Main.koneksi)
                bulkCopy.DestinationTableName = "dbo.Componentslist"
                bulkCopy.BulkCopyTimeout = 120
                Try
                    rd = cmd.ExecuteReader
                    bulkCopy.ColumnMappings.Add(0, 1)
                    bulkCopy.ColumnMappings.Add(1, 2)
                    'bulkCopy.ColumnMappings.Add(2, 3)
                    'bulkCopy.ColumnMappings.Add(3, 4)
                    ' bulkCopy.ColumnMappings.Add(4, 5)
                    'bulkCopy.ColumnMappings.Add(5, 6)
                    'bulkCopy.ColumnMappings.Add(6, 7)

                    bulkCopy.WriteToServer(rd)

                    ' Dim Sql = "INSERT INTO [Upload_History] (upload) Values ('COOIS');"
                    ' Dim insert = New SqlCommand(Sql, Main.koneksi)
                    'insert.ExecuteNonQuery()

                    rd.Close()
                    MsgBox("Upload New Component List Successed !")
                Catch ex As Exception
                    MsgBox("Upload Component List Failed " & ex.Message)
                End Try
            End Using
        End If
    End Sub

    Private Sub Button20_Click(sender As Object, e As EventArgs) Handles Button20.Click
        OpenFileDialog1.InitialDirectory = My.Computer.FileSystem.SpecialDirectories.MyDocuments
        If OpenFileDialog1.ShowDialog(Me) = System.Windows.Forms.DialogResult.OK Then
            Dim xlApp As New Microsoft.Office.Interop.Excel.Application
            Dim xlWorkBook As Microsoft.Office.Interop.Excel.Workbook = xlApp.Workbooks.Open(OpenFileDialog1.FileName)
            Dim SheetName As String = xlWorkBook.Worksheets(1).Name.ToString
            Dim excelpath As String = OpenFileDialog1.FileName
            Dim koneksiExcel As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & excelpath & ";Extended Properties='Excel 8.0;HDR=No;IMEX=1;'"
            oleCon = New OleDbConnection(koneksiExcel)
            oleCon.Open()

            Dim queryExcel As String = "select * from [" & SheetName & "$]"
            Dim cmd As OleDbCommand = New OleDbCommand(queryExcel, oleCon)
            Dim rd As OleDbDataReader

            Call koneksi_db()

            Using bulkCopy As SqlBulkCopy = New SqlBulkCopy(Main.koneksi)
                bulkCopy.DestinationTableName = "dbo.NSXMasterdata"
                bulkCopy.BulkCopyTimeout = 120
                Try
                    rd = cmd.ExecuteReader
                    bulkCopy.ColumnMappings.Add(0, 1)
                    bulkCopy.ColumnMappings.Add(1, 2)
                    bulkCopy.ColumnMappings.Add(2, 3)
                    bulkCopy.ColumnMappings.Add(3, 4)
                    bulkCopy.ColumnMappings.Add(4, 5)
                    bulkCopy.ColumnMappings.Add(5, 6)
                    bulkCopy.ColumnMappings.Add(6, 7)
                    bulkCopy.ColumnMappings.Add(7, 8)
                    bulkCopy.ColumnMappings.Add(8, 9)
                    bulkCopy.ColumnMappings.Add(9, 10)
                    bulkCopy.ColumnMappings.Add(10, 11)
                    bulkCopy.ColumnMappings.Add(11, 12)
                    bulkCopy.ColumnMappings.Add(12, 13)
                    bulkCopy.ColumnMappings.Add(13, 14)
                    bulkCopy.ColumnMappings.Add(14, 15)
                    bulkCopy.ColumnMappings.Add(15, 16)
                    'bulkCopy.ColumnMappings.Add(16, 17)

                    bulkCopy.WriteToServer(rd)

                    ' Dim Sql = "INSERT INTO [Upload_History] (upload) Values ('COOIS');"
                    ' Dim insert = New SqlCommand(Sql, Main.koneksi)
                    'insert.ExecuteNonQuery()

                    rd.Close()
                    MsgBox("Upload New NSX Master Data Successed !")
                Catch ex As Exception
                    MsgBox("Upload NSX Master Data Failed " & ex.Message)
                End Try
            End Using
        End If
    End Sub

    Private Sub btn_NewComponents_Click(sender As Object, e As EventArgs) Handles btn_NewComponents.Click
        DataGridUpdateDb.Show()
    End Sub

    Private Sub Button13_Click_1(sender As Object, e As EventArgs) Handles Button13.Click
        Dim confirm = MessageBox.Show("Are You Sure For Logout?", "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If confirm = Windows.Forms.DialogResult.Yes Then
            Me.Hide()
            LoginForm.Show()
            LoginForm.textboxusername.Text = ""
            LoginForm.textboxpassword.Text = ""
        ElseIf confirm = Windows.Forms.DialogResult.Yes Then
            Exit Sub
        End If
        role = ""
    End Sub

End Class
