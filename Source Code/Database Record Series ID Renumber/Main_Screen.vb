Imports Microsoft.Win32
Imports System.IO

Public Class Main_Screen
    Inherits System.Windows.Forms.Form

    Dim WithEvents Worker1 As Worker
    Public Delegate Sub WorkerhHandler(ByVal Result As String)
    Public Delegate Sub WorkerProgresshHandler(ByVal switch As Integer)

    Private splash_loader As Splash_Screen
    Public dataloaded As Boolean = False

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        Worker1 = New Worker
        AddHandler Worker1.WorkerComplete, AddressOf WorkerHandler
        AddHandler Worker1.WorkerProgress, AddressOf WorkerProgressHandler
    End Sub

    Public Sub New(ByVal splash As Splash_Screen)
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        splash_loader = splash
        Worker1 = New Worker
        AddHandler Worker1.WorkerComplete, AddressOf WorkerHandler
        AddHandler Worker1.WorkerProgress, AddressOf WorkerProgressHandler
    End Sub
    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents Timer1 As System.Windows.Forms.Timer
    Friend WithEvents PictureBox5 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox4 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox3 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox2 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Friend WithEvents ImageList1 As System.Windows.Forms.ImageList
    Friend WithEvents Timer2 As System.Windows.Forms.Timer
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents NotifyIcon1 As System.Windows.Forms.NotifyIcon
    Friend WithEvents ContextMenu1 As System.Windows.Forms.ContextMenu
    Friend WithEvents MenuItem2 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem3 As System.Windows.Forms.MenuItem
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents SaveFileDialog1 As System.Windows.Forms.SaveFileDialog
    Friend WithEvents FolderBrowserDialog1 As System.Windows.Forms.FolderBrowserDialog
    Friend WithEvents Button3 As System.Windows.Forms.Button
    Friend WithEvents Button4 As System.Windows.Forms.Button
    Friend WithEvents MenuItem1 As System.Windows.Forms.MenuItem
    Friend WithEvents OpenFileDialog1 As System.Windows.Forms.OpenFileDialog
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents Label24 As System.Windows.Forms.Label
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents DatabaseToUse As System.Windows.Forms.TextBox
    Friend WithEvents SeriesStart As System.Windows.Forms.NumericUpDown
    Friend WithEvents SeriesPrefix As System.Windows.Forms.TextBox
    Friend WithEvents SeriesSuffix As System.Windows.Forms.TextBox
    Friend WithEvents TableToUse As System.Windows.Forms.TextBox
    Friend WithEvents ColumnSort As System.Windows.Forms.TextBox
    Friend WithEvents ColumnRenumber As System.Windows.Forms.TextBox
    Friend WithEvents analysisend As System.Windows.Forms.Label
    Friend WithEvents recordsrenumbered As System.Windows.Forms.Label
    Friend WithEvents analysisstart As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents ComboBox1 As System.Windows.Forms.ComboBox
    Friend WithEvents ColumnSortOrder As System.Windows.Forms.TextBox
    Friend WithEvents totalrecords As System.Windows.Forms.Label
    Friend WithEvents Label12 As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(Main_Screen))
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.totalrecords = New System.Windows.Forms.Label
        Me.Label12 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.PictureBox5 = New System.Windows.Forms.PictureBox
        Me.PictureBox4 = New System.Windows.Forms.PictureBox
        Me.PictureBox3 = New System.Windows.Forms.PictureBox
        Me.PictureBox2 = New System.Windows.Forms.PictureBox
        Me.PictureBox1 = New System.Windows.Forms.PictureBox
        Me.analysisend = New System.Windows.Forms.Label
        Me.Label15 = New System.Windows.Forms.Label
        Me.recordsrenumbered = New System.Windows.Forms.Label
        Me.analysisstart = New System.Windows.Forms.Label
        Me.Label6 = New System.Windows.Forms.Label
        Me.Label5 = New System.Windows.Forms.Label
        Me.Button3 = New System.Windows.Forms.Button
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.ImageList1 = New System.Windows.Forms.ImageList(Me.components)
        Me.Timer2 = New System.Windows.Forms.Timer(Me.components)
        Me.Label8 = New System.Windows.Forms.Label
        Me.NotifyIcon1 = New System.Windows.Forms.NotifyIcon(Me.components)
        Me.ContextMenu1 = New System.Windows.Forms.ContextMenu
        Me.MenuItem1 = New System.Windows.Forms.MenuItem
        Me.MenuItem2 = New System.Windows.Forms.MenuItem
        Me.MenuItem3 = New System.Windows.Forms.MenuItem
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.Label9 = New System.Windows.Forms.Label
        Me.SaveFileDialog1 = New System.Windows.Forms.SaveFileDialog
        Me.FolderBrowserDialog1 = New System.Windows.Forms.FolderBrowserDialog
        Me.Button4 = New System.Windows.Forms.Button
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog
        Me.Button1 = New System.Windows.Forms.Button
        Me.DatabaseToUse = New System.Windows.Forms.TextBox
        Me.Label24 = New System.Windows.Forms.Label
        Me.Label13 = New System.Windows.Forms.Label
        Me.SeriesStart = New System.Windows.Forms.NumericUpDown
        Me.SeriesPrefix = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.SeriesSuffix = New System.Windows.Forms.TextBox
        Me.TableToUse = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.ColumnSort = New System.Windows.Forms.TextBox
        Me.Label7 = New System.Windows.Forms.Label
        Me.ColumnRenumber = New System.Windows.Forms.TextBox
        Me.Label10 = New System.Windows.Forms.Label
        Me.ComboBox1 = New System.Windows.Forms.ComboBox
        Me.ColumnSortOrder = New System.Windows.Forms.TextBox
        Me.Panel1.SuspendLayout()
        CType(Me.SeriesStart, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.Color.White
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel1.Controls.Add(Me.totalrecords)
        Me.Panel1.Controls.Add(Me.Label12)
        Me.Panel1.Controls.Add(Me.Label1)
        Me.Panel1.Controls.Add(Me.PictureBox5)
        Me.Panel1.Controls.Add(Me.PictureBox4)
        Me.Panel1.Controls.Add(Me.PictureBox3)
        Me.Panel1.Controls.Add(Me.PictureBox2)
        Me.Panel1.Controls.Add(Me.PictureBox1)
        Me.Panel1.Controls.Add(Me.analysisend)
        Me.Panel1.Controls.Add(Me.Label15)
        Me.Panel1.Controls.Add(Me.recordsrenumbered)
        Me.Panel1.Controls.Add(Me.analysisstart)
        Me.Panel1.Controls.Add(Me.Label6)
        Me.Panel1.Controls.Add(Me.Label5)
        Me.Panel1.Location = New System.Drawing.Point(8, 264)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(288, 104)
        Me.Panel1.TabIndex = 11
        '
        'totalrecords
        '
        Me.totalrecords.BackColor = System.Drawing.Color.Transparent
        Me.totalrecords.ForeColor = System.Drawing.Color.Black
        Me.totalrecords.Location = New System.Drawing.Point(136, 24)
        Me.totalrecords.Name = "totalrecords"
        Me.totalrecords.Size = New System.Drawing.Size(144, 16)
        Me.totalrecords.TabIndex = 43
        Me.totalrecords.Text = "0"
        '
        'Label12
        '
        Me.Label12.BackColor = System.Drawing.Color.Transparent
        Me.Label12.ForeColor = System.Drawing.Color.Black
        Me.Label12.Location = New System.Drawing.Point(8, 24)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(136, 16)
        Me.Label12.TabIndex = 42
        Me.Label12.Text = "Total Records:"
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.Color.White
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.LimeGreen
        Me.Label1.Location = New System.Drawing.Point(216, 80)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(66, 16)
        Me.Label1.TabIndex = 27
        Me.Label1.Text = "Waiting"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'PictureBox5
        '
        Me.PictureBox5.BackgroundImage = CType(resources.GetObject("PictureBox5.BackgroundImage"), System.Drawing.Image)
        Me.PictureBox5.Location = New System.Drawing.Point(192, 80)
        Me.PictureBox5.Name = "PictureBox5"
        Me.PictureBox5.Size = New System.Drawing.Size(16, 16)
        Me.PictureBox5.TabIndex = 26
        Me.PictureBox5.TabStop = False
        '
        'PictureBox4
        '
        Me.PictureBox4.BackgroundImage = CType(resources.GetObject("PictureBox4.BackgroundImage"), System.Drawing.Image)
        Me.PictureBox4.Location = New System.Drawing.Point(176, 80)
        Me.PictureBox4.Name = "PictureBox4"
        Me.PictureBox4.Size = New System.Drawing.Size(16, 16)
        Me.PictureBox4.TabIndex = 25
        Me.PictureBox4.TabStop = False
        '
        'PictureBox3
        '
        Me.PictureBox3.BackgroundImage = CType(resources.GetObject("PictureBox3.BackgroundImage"), System.Drawing.Image)
        Me.PictureBox3.Location = New System.Drawing.Point(160, 80)
        Me.PictureBox3.Name = "PictureBox3"
        Me.PictureBox3.Size = New System.Drawing.Size(16, 16)
        Me.PictureBox3.TabIndex = 24
        Me.PictureBox3.TabStop = False
        '
        'PictureBox2
        '
        Me.PictureBox2.BackgroundImage = CType(resources.GetObject("PictureBox2.BackgroundImage"), System.Drawing.Image)
        Me.PictureBox2.Location = New System.Drawing.Point(144, 80)
        Me.PictureBox2.Name = "PictureBox2"
        Me.PictureBox2.Size = New System.Drawing.Size(16, 16)
        Me.PictureBox2.TabIndex = 23
        Me.PictureBox2.TabStop = False
        '
        'PictureBox1
        '
        Me.PictureBox1.BackgroundImage = CType(resources.GetObject("PictureBox1.BackgroundImage"), System.Drawing.Image)
        Me.PictureBox1.Location = New System.Drawing.Point(128, 80)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(16, 16)
        Me.PictureBox1.TabIndex = 22
        Me.PictureBox1.TabStop = False
        '
        'analysisend
        '
        Me.analysisend.BackColor = System.Drawing.Color.Transparent
        Me.analysisend.ForeColor = System.Drawing.Color.Black
        Me.analysisend.Location = New System.Drawing.Point(136, 56)
        Me.analysisend.Name = "analysisend"
        Me.analysisend.Size = New System.Drawing.Size(144, 16)
        Me.analysisend.TabIndex = 41
        '
        'Label15
        '
        Me.Label15.BackColor = System.Drawing.Color.Transparent
        Me.Label15.ForeColor = System.Drawing.Color.Black
        Me.Label15.Location = New System.Drawing.Point(8, 56)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(152, 16)
        Me.Label15.TabIndex = 37
        Me.Label15.Text = "Analysis End:"
        '
        'recordsrenumbered
        '
        Me.recordsrenumbered.BackColor = System.Drawing.Color.Transparent
        Me.recordsrenumbered.ForeColor = System.Drawing.Color.Black
        Me.recordsrenumbered.Location = New System.Drawing.Point(136, 40)
        Me.recordsrenumbered.Name = "recordsrenumbered"
        Me.recordsrenumbered.Size = New System.Drawing.Size(144, 16)
        Me.recordsrenumbered.TabIndex = 35
        Me.recordsrenumbered.Text = "0"
        '
        'analysisstart
        '
        Me.analysisstart.BackColor = System.Drawing.Color.Transparent
        Me.analysisstart.ForeColor = System.Drawing.Color.Black
        Me.analysisstart.Location = New System.Drawing.Point(136, 8)
        Me.analysisstart.Name = "analysisstart"
        Me.analysisstart.Size = New System.Drawing.Size(144, 16)
        Me.analysisstart.TabIndex = 34
        '
        'Label6
        '
        Me.Label6.BackColor = System.Drawing.Color.Transparent
        Me.Label6.ForeColor = System.Drawing.Color.Black
        Me.Label6.Location = New System.Drawing.Point(8, 40)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(136, 16)
        Me.Label6.TabIndex = 31
        Me.Label6.Text = "Records Renumbered:"
        '
        'Label5
        '
        Me.Label5.BackColor = System.Drawing.Color.Transparent
        Me.Label5.ForeColor = System.Drawing.Color.Black
        Me.Label5.Location = New System.Drawing.Point(8, 8)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(136, 16)
        Me.Label5.TabIndex = 30
        Me.Label5.Text = "Analysis Start:"
        '
        'Button3
        '
        Me.Button3.BackColor = System.Drawing.Color.FromArgb(CType(224, Byte), CType(224, Byte), CType(224, Byte))
        Me.Button3.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button3.Location = New System.Drawing.Point(96, 224)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(120, 23)
        Me.Button3.TabIndex = 54
        Me.Button3.Text = "Renumber Records"
        '
        'ImageList1
        '
        Me.ImageList1.ImageSize = New System.Drawing.Size(16, 16)
        Me.ImageList1.ImageStream = CType(resources.GetObject("ImageList1.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.ImageList1.TransparentColor = System.Drawing.Color.Transparent
        '
        'Timer2
        '
        Me.Timer2.Interval = 1000
        '
        'Label8
        '
        Me.Label8.BackColor = System.Drawing.Color.Transparent
        Me.Label8.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label8.ForeColor = System.Drawing.Color.FromArgb(CType(64, Byte), CType(64, Byte), CType(64, Byte))
        Me.Label8.Location = New System.Drawing.Point(96, 376)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(96, 16)
        Me.Label8.TabIndex = 33
        Me.Label8.TextAlign = System.Drawing.ContentAlignment.TopRight
        Me.ToolTip1.SetToolTip(Me.Label8, "Current System Time")
        '
        'NotifyIcon1
        '
        Me.NotifyIcon1.ContextMenu = Me.ContextMenu1
        Me.NotifyIcon1.Icon = CType(resources.GetObject("NotifyIcon1.Icon"), System.Drawing.Icon)
        Me.NotifyIcon1.Text = "Resting..."
        Me.NotifyIcon1.Visible = True
        '
        'ContextMenu1
        '
        Me.ContextMenu1.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem1, Me.MenuItem2, Me.MenuItem3})
        '
        'MenuItem1
        '
        Me.MenuItem1.Index = 0
        Me.MenuItem1.Text = "Display Main Screen"
        '
        'MenuItem2
        '
        Me.MenuItem2.Index = 1
        Me.MenuItem2.Text = "-"
        '
        'MenuItem3
        '
        Me.MenuItem3.Index = 2
        Me.MenuItem3.Text = "Exit Application"
        '
        'Label9
        '
        Me.Label9.BackColor = System.Drawing.Color.Transparent
        Me.Label9.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label9.ForeColor = System.Drawing.Color.FromArgb(CType(64, Byte), CType(64, Byte), CType(64, Byte))
        Me.Label9.Location = New System.Drawing.Point(208, 376)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(88, 16)
        Me.Label9.TabIndex = 76
        Me.Label9.Text = "BUILD 20051102.1"
        Me.Label9.TextAlign = System.Drawing.ContentAlignment.TopRight
        Me.ToolTip1.SetToolTip(Me.Label9, "Current System Time")
        '
        'SaveFileDialog1
        '
        Me.SaveFileDialog1.DefaultExt = "mdb"
        Me.SaveFileDialog1.Filter = "MSAccess files|*.mdb"
        Me.SaveFileDialog1.Title = "Database Save Path"
        '
        'FolderBrowserDialog1
        '
        Me.FolderBrowserDialog1.Description = "Select the folder to rename the files in"
        '
        'Button4
        '
        Me.Button4.BackColor = System.Drawing.Color.FromArgb(CType(224, Byte), CType(224, Byte), CType(224, Byte))
        Me.Button4.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button4.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button4.ForeColor = System.Drawing.Color.Black
        Me.Button4.Location = New System.Drawing.Point(8, 376)
        Me.Button4.Name = "Button4"
        Me.Button4.Size = New System.Drawing.Size(88, 16)
        Me.Button4.TabIndex = 59
        Me.Button4.Text = "KILL PROCESSES"
        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.DefaultExt = "mdb"
        Me.OpenFileDialog1.Filter = "MSAccess | *.mdb"
        Me.OpenFileDialog1.Title = "Select the Access database to manipulate"
        '
        'Button1
        '
        Me.Button1.BackColor = System.Drawing.Color.FromArgb(CType(0, Byte), CType(192, Byte), CType(192, Byte))
        Me.Button1.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button1.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button1.ForeColor = System.Drawing.Color.Black
        Me.Button1.Location = New System.Drawing.Point(240, 24)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(56, 20)
        Me.Button1.TabIndex = 61
        Me.Button1.Text = "BROWSE"
        '
        'DatabaseToUse
        '
        Me.DatabaseToUse.BackColor = System.Drawing.Color.White
        Me.DatabaseToUse.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.DatabaseToUse.ForeColor = System.Drawing.Color.Black
        Me.DatabaseToUse.Location = New System.Drawing.Point(8, 24)
        Me.DatabaseToUse.Name = "DatabaseToUse"
        Me.DatabaseToUse.ReadOnly = True
        Me.DatabaseToUse.Size = New System.Drawing.Size(224, 20)
        Me.DatabaseToUse.TabIndex = 60
        Me.DatabaseToUse.Text = ""
        '
        'Label24
        '
        Me.Label24.BackColor = System.Drawing.Color.Transparent
        Me.Label24.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label24.ForeColor = System.Drawing.Color.Black
        Me.Label24.Location = New System.Drawing.Point(8, 8)
        Me.Label24.Name = "Label24"
        Me.Label24.Size = New System.Drawing.Size(144, 16)
        Me.Label24.TabIndex = 62
        Me.Label24.Text = "DATABASE TO USE"
        Me.Label24.TextAlign = System.Drawing.ContentAlignment.BottomLeft
        '
        'Label13
        '
        Me.Label13.BackColor = System.Drawing.Color.Transparent
        Me.Label13.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label13.ForeColor = System.Drawing.Color.Black
        Me.Label13.Location = New System.Drawing.Point(120, 168)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(72, 16)
        Me.Label13.TabIndex = 64
        Me.Label13.Text = "SERIES START"
        Me.Label13.TextAlign = System.Drawing.ContentAlignment.BottomLeft
        '
        'SeriesStart
        '
        Me.SeriesStart.BackColor = System.Drawing.Color.White
        Me.SeriesStart.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.SeriesStart.ForeColor = System.Drawing.Color.Black
        Me.SeriesStart.Location = New System.Drawing.Point(120, 184)
        Me.SeriesStart.Maximum = New Decimal(New Integer() {10000, 0, 0, 0})
        Me.SeriesStart.Minimum = New Decimal(New Integer() {10000, 0, 0, -2147483648})
        Me.SeriesStart.Name = "SeriesStart"
        Me.SeriesStart.Size = New System.Drawing.Size(64, 20)
        Me.SeriesStart.TabIndex = 63
        Me.SeriesStart.Value = New Decimal(New Integer() {1, 0, 0, 0})
        '
        'SeriesPrefix
        '
        Me.SeriesPrefix.BackColor = System.Drawing.Color.White
        Me.SeriesPrefix.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.SeriesPrefix.ForeColor = System.Drawing.Color.Black
        Me.SeriesPrefix.Location = New System.Drawing.Point(8, 184)
        Me.SeriesPrefix.Name = "SeriesPrefix"
        Me.SeriesPrefix.Size = New System.Drawing.Size(104, 20)
        Me.SeriesPrefix.TabIndex = 65
        Me.SeriesPrefix.Text = ""
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.Color.Transparent
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.Color.Black
        Me.Label2.Location = New System.Drawing.Point(8, 168)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(72, 16)
        Me.Label2.TabIndex = 66
        Me.Label2.Text = "SERIES PREFIX"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.BottomLeft
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.Color.Transparent
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.ForeColor = System.Drawing.Color.White
        Me.Label3.Location = New System.Drawing.Point(224, 168)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(72, 16)
        Me.Label3.TabIndex = 68
        Me.Label3.Text = "SERIES SUFFIX"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.BottomLeft
        '
        'SeriesSuffix
        '
        Me.SeriesSuffix.BackColor = System.Drawing.Color.White
        Me.SeriesSuffix.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.SeriesSuffix.ForeColor = System.Drawing.Color.Black
        Me.SeriesSuffix.Location = New System.Drawing.Point(192, 184)
        Me.SeriesSuffix.Name = "SeriesSuffix"
        Me.SeriesSuffix.Size = New System.Drawing.Size(104, 20)
        Me.SeriesSuffix.TabIndex = 67
        Me.SeriesSuffix.Text = ""
        '
        'TableToUse
        '
        Me.TableToUse.BackColor = System.Drawing.Color.White
        Me.TableToUse.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.TableToUse.ForeColor = System.Drawing.Color.Black
        Me.TableToUse.Location = New System.Drawing.Point(8, 64)
        Me.TableToUse.Name = "TableToUse"
        Me.TableToUse.ReadOnly = True
        Me.TableToUse.Size = New System.Drawing.Size(288, 20)
        Me.TableToUse.TabIndex = 69
        Me.TableToUse.Text = ""
        '
        'Label4
        '
        Me.Label4.BackColor = System.Drawing.Color.Transparent
        Me.Label4.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.ForeColor = System.Drawing.Color.Black
        Me.Label4.Location = New System.Drawing.Point(8, 48)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(144, 16)
        Me.Label4.TabIndex = 70
        Me.Label4.Text = "TABLE TO EXAMINE"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.BottomLeft
        '
        'ColumnSort
        '
        Me.ColumnSort.BackColor = System.Drawing.Color.White
        Me.ColumnSort.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.ColumnSort.ForeColor = System.Drawing.Color.Black
        Me.ColumnSort.Location = New System.Drawing.Point(8, 104)
        Me.ColumnSort.Name = "ColumnSort"
        Me.ColumnSort.ReadOnly = True
        Me.ColumnSort.Size = New System.Drawing.Size(240, 20)
        Me.ColumnSort.TabIndex = 71
        Me.ColumnSort.Text = ""
        '
        'Label7
        '
        Me.Label7.BackColor = System.Drawing.Color.Transparent
        Me.Label7.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.ForeColor = System.Drawing.Color.Black
        Me.Label7.Location = New System.Drawing.Point(8, 88)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(144, 16)
        Me.Label7.TabIndex = 72
        Me.Label7.Text = "COLUMN TO SORT RECORDS ON"
        Me.Label7.TextAlign = System.Drawing.ContentAlignment.BottomLeft
        '
        'ColumnRenumber
        '
        Me.ColumnRenumber.BackColor = System.Drawing.Color.White
        Me.ColumnRenumber.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.ColumnRenumber.ForeColor = System.Drawing.Color.Black
        Me.ColumnRenumber.Location = New System.Drawing.Point(8, 144)
        Me.ColumnRenumber.Name = "ColumnRenumber"
        Me.ColumnRenumber.ReadOnly = True
        Me.ColumnRenumber.Size = New System.Drawing.Size(288, 20)
        Me.ColumnRenumber.TabIndex = 73
        Me.ColumnRenumber.Text = ""
        '
        'Label10
        '
        Me.Label10.BackColor = System.Drawing.Color.Transparent
        Me.Label10.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label10.ForeColor = System.Drawing.Color.Black
        Me.Label10.Location = New System.Drawing.Point(8, 128)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(144, 16)
        Me.Label10.TabIndex = 74
        Me.Label10.Text = "COLUMN TO RENUMBER"
        Me.Label10.TextAlign = System.Drawing.ContentAlignment.BottomLeft
        '
        'ComboBox1
        '
        Me.ComboBox1.Location = New System.Drawing.Point(248, 88)
        Me.ComboBox1.Name = "ComboBox1"
        Me.ComboBox1.Size = New System.Drawing.Size(48, 21)
        Me.ComboBox1.TabIndex = 0
        '
        'ColumnSortOrder
        '
        Me.ColumnSortOrder.BackColor = System.Drawing.Color.White
        Me.ColumnSortOrder.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.ColumnSortOrder.ForeColor = System.Drawing.Color.Black
        Me.ColumnSortOrder.Location = New System.Drawing.Point(256, 104)
        Me.ColumnSortOrder.Name = "ColumnSortOrder"
        Me.ColumnSortOrder.ReadOnly = True
        Me.ColumnSortOrder.Size = New System.Drawing.Size(39, 20)
        Me.ColumnSortOrder.TabIndex = 77
        Me.ColumnSortOrder.Text = ""
        '
        'Main_Screen
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.Color.FromArgb(CType(192, Byte), CType(255, Byte), CType(255, Byte))
        Me.BackgroundImage = CType(resources.GetObject("$this.BackgroundImage"), System.Drawing.Image)
        Me.ClientSize = New System.Drawing.Size(306, 400)
        Me.Controls.Add(Me.ColumnSortOrder)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.Label9)
        Me.Controls.Add(Me.ColumnRenumber)
        Me.Controls.Add(Me.Label10)
        Me.Controls.Add(Me.ColumnSort)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.TableToUse)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.SeriesSuffix)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.SeriesPrefix)
        Me.Controls.Add(Me.Label13)
        Me.Controls.Add(Me.SeriesStart)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.DatabaseToUse)
        Me.Controls.Add(Me.Label24)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.Button4)
        Me.Controls.Add(Me.Button3)
        Me.ForeColor = System.Drawing.Color.Black
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MaximumSize = New System.Drawing.Size(312, 432)
        Me.MinimumSize = New System.Drawing.Size(312, 432)
        Me.Name = "Main_Screen"
        Me.ShowInTaskbar = False
        Me.Text = "DB Record Series ID Renumber"
        Me.Panel1.ResumeLayout(False)
        CType(Me.SeriesStart, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region


    Private current_light As Integer = 0
    Private current_colour As Integer = 0
    Private currently_working As Boolean = False




    Private Sub Error_Handler(ByVal ex As Exception, Optional ByVal identifier_msg As String = "")
        Try
            If ex.Message.IndexOf("Thread was being aborted") < 0 Then
                Dim Display_Message1 As New Display_Message("The Application encountered the following problem: " & vbCrLf & ex.ToString)
                Display_Message1.ShowDialog()
                Dim dir As DirectoryInfo = New DirectoryInfo((Application.StartupPath & "\").Replace("\\", "\") & "Error Logs")
                If dir.Exists = False Then
                    dir.Create()
                End If
                Dim filewriter As StreamWriter = New StreamWriter((Application.StartupPath & "\").Replace("\\", "\") & "Error Logs\" & Format(Now(), "yyyyMMdd") & "_Error_Log.txt", True)
                filewriter.WriteLine("#" & Format(Now(), "dd/MM/yyyy hh:mm:ss tt") & " - " & ex.ToString)
                filewriter.Flush()
                filewriter.Close()
            End If
        Catch exc As Exception

            MsgBox("An error occurred in Database Record Series ID Renumber's error handling routine. The application will try to recover from this serious error.", MsgBoxStyle.Critical, "Critical Error Encountered")
        End Try
    End Sub

    Public Sub Load_Registry_Values()
        Try
            Dim configflag As Boolean
            configflag = False
            Dim str As String
            Dim keyflag1 As Boolean = False
            Dim oReg As RegistryKey = Registry.LocalMachine
            Dim keys() As String = oReg.GetSubKeyNames()
            System.Array.Sort(keys)

            For Each str In keys
                If str.Equals("Software\Database Record Series ID Renumber") = True Then
                    keyflag1 = True
                    Exit For
                End If
            Next str

            If keyflag1 = False Then
                oReg.CreateSubKey("Software\Database Record Series ID Renumber")
            End If

            keyflag1 = False

            Dim oKey As RegistryKey = oReg.OpenSubKey("Software\Database Record Series ID Renumber", True)

            str = oKey.GetValue("DatabaseToUse")
            If Not IsNothing(str) And Not (str = "") Then
                DatabaseToUse.Text = str
            Else
                configflag = True
                oKey.SetValue("DatabaseToUse", "")
                DatabaseToUse.Text = ""
            End If

            str = oKey.GetValue("TableToUse")
            If Not IsNothing(str) And Not (str = "") Then
                TableToUse.Text = str
            Else
                configflag = True
                oKey.SetValue("TableToUse", "")
                TableToUse.Text = ""
            End If

            str = oKey.GetValue("ColumnSort")
            If Not IsNothing(str) And Not (str = "") Then
                ColumnSort.Text = str
            Else
                configflag = True
                oKey.SetValue("ColumnSort", "")
                ColumnSort.Text = ""
            End If

            str = oKey.GetValue("ColumnSortOrder")
            If Not IsNothing(str) And Not (str = "") Then
                ColumnSortOrder.Text = str
            Else
                configflag = True
                oKey.SetValue("ColumnSortOrder", "asc")
                ColumnSortOrder.Text = "asc"
            End If

            str = oKey.GetValue("ColumnRenumber")
            If Not IsNothing(str) And Not (str = "") Then
                ColumnRenumber.Text = str
            Else
                configflag = True
                oKey.SetValue("ColumnRenumber", "")
                ColumnRenumber.Text = ""
            End If

            str = oKey.GetValue("SeriesPrefix")
            If Not IsNothing(str) And Not (str = "") Then
                SeriesPrefix.Text = str
            Else
                configflag = True
                oKey.SetValue("SeriesPrefix", "")
                SeriesPrefix.Text = ""
            End If

            str = oKey.GetValue("SeriesSuffix")
            If Not IsNothing(str) And Not (str = "") Then
                SeriesSuffix.Text = str
            Else
                configflag = True
                oKey.SetValue("SeriesSuffix", "")
                SeriesSuffix.Text = ""
            End If

            str = oKey.GetValue("SeriesStart")
            If Not IsNothing(str) And Not (str = "") Then
                SeriesStart.Value = CInt(str)
            Else
                configflag = True
                oKey.SetValue("SeriesStart", "1")
                SeriesStart.Value = 1
            End If

            oKey.Close()
            oReg.Close()

        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub

    Private Sub Save_Registry_Values()
        Try
            Dim oReg As RegistryKey = Registry.LocalMachine
            Dim oKey As RegistryKey = oReg.OpenSubKey("Software\Database Record Series ID Renumber", True)

            oKey.SetValue("DatabaseToUse", DatabaseToUse.Text)
            oKey.SetValue("TableToUse", TableToUse.Text)
            oKey.SetValue("DatabaseToUse", DatabaseToUse.Text)
            oKey.SetValue("ColumnSort", ColumnSort.Text)
            oKey.SetValue("ColumnSortOrder", ColumnSortOrder.Text)
            oKey.SetValue("ColumnRenumber", ColumnRenumber.Text)
            oKey.SetValue("SeriesPrefix", SeriesPrefix.Text)
            oKey.SetValue("SeriesSuffix", SeriesSuffix.Text)
            oKey.SetValue("SeriesStart", SeriesStart.Value.ToString)

            oKey.Close()
            oReg.Close()
        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub

    Private Sub run_green_lights()
        Try
            Label1.ForeColor = Color.LimeGreen
            Label1.Text = "Waiting"


            current_light = current_light - 1
            If current_light < 1 Then
                current_light = 5
            End If
            current_colour = 0
            PictureBox1.Image = ImageList1.Images(1)
            PictureBox2.Image = ImageList1.Images(1)
            PictureBox3.Image = ImageList1.Images(1)
            PictureBox4.Image = ImageList1.Images(1)
            PictureBox5.Image = ImageList1.Images(1)

            Select Case current_light
                Case 0

                    PictureBox1.Image = ImageList1.Images(0)
                Case 1

                    PictureBox2.Image = ImageList1.Images(0)
                Case 2

                    PictureBox3.Image = ImageList1.Images(0)
                Case 3

                    PictureBox4.Image = ImageList1.Images(0)
                Case 4

                    PictureBox5.Image = ImageList1.Images(0)
                Case 5

                    PictureBox1.Image = ImageList1.Images(0)
            End Select

            current_light = current_light + 1
            If current_light > 5 Then
                current_light = 1
            End If
        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub

    Private Sub run_orange_lights()
        Try
            Label1.ForeColor = Color.DarkOrange
            Label1.Text = "Working"

            current_light = current_light - 1
            If current_light < 1 Then
                current_light = 5
            End If
            current_colour = 1
            PictureBox1.Image = ImageList1.Images(3)
            PictureBox2.Image = ImageList1.Images(3)
            PictureBox3.Image = ImageList1.Images(3)
            PictureBox4.Image = ImageList1.Images(3)
            PictureBox5.Image = ImageList1.Images(3)
            Select Case current_light
                Case 0
                    PictureBox1.Image = ImageList1.Images(2)
                Case 1
                    PictureBox2.Image = ImageList1.Images(2)
                Case 2
                    PictureBox3.Image = ImageList1.Images(2)
                Case 3
                    PictureBox4.Image = ImageList1.Images(2)
                Case 4
                    PictureBox5.Image = ImageList1.Images(2)
                Case 5
                    PictureBox1.Image = ImageList1.Images(2)
            End Select

            current_light = current_light + 1
            If current_light > 5 Then
                current_light = 1
            End If
        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub

    Private Sub run_lights()
        Try
            If current_colour = 1 Then
                Select Case current_light
                    Case 0
                        PictureBox5.Image = ImageList1.Images(3)
                        PictureBox1.Image = ImageList1.Images(2)
                    Case 1
                        PictureBox1.Image = ImageList1.Images(3)
                        PictureBox2.Image = ImageList1.Images(2)
                    Case 2
                        PictureBox2.Image = ImageList1.Images(3)
                        PictureBox3.Image = ImageList1.Images(2)
                    Case 3
                        PictureBox3.Image = ImageList1.Images(3)
                        PictureBox4.Image = ImageList1.Images(2)
                    Case 4
                        PictureBox4.Image = ImageList1.Images(3)
                        PictureBox5.Image = ImageList1.Images(2)
                    Case 5
                        PictureBox5.Image = ImageList1.Images(3)
                        PictureBox1.Image = ImageList1.Images(2)
                End Select
            Else
                Select Case current_light
                    Case 0
                        PictureBox5.Image = ImageList1.Images(1)
                        PictureBox1.Image = ImageList1.Images(0)
                    Case 1
                        PictureBox1.Image = ImageList1.Images(1)
                        PictureBox2.Image = ImageList1.Images(0)
                    Case 2
                        PictureBox2.Image = ImageList1.Images(1)
                        PictureBox3.Image = ImageList1.Images(0)
                    Case 3
                        PictureBox3.Image = ImageList1.Images(1)
                        PictureBox4.Image = ImageList1.Images(0)
                    Case 4
                        PictureBox4.Image = ImageList1.Images(1)
                        PictureBox5.Image = ImageList1.Images(0)
                    Case 5
                        PictureBox5.Image = ImageList1.Images(1)
                        PictureBox1.Image = ImageList1.Images(0)
                End Select

            End If

            current_light = current_light + 1
            If current_light > 5 Then
                current_light = 1
            End If
        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub

    Private Sub Timer2_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer2.Tick
        Try
            run_lights()
            'Label8.Text = Format(Now(), "dd/MM/yyyy hh:mm:ss tt")
            Label8.Text = Format(Now(), "dd/MM/yyyy HH:mm:ss")
        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub

    Private Sub Main_Screen_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            '  Label8.Text = Format(Now(), "dd/MM/yyyy hh:mm:ss tt")
            Label8.Text = Format(Now(), "dd/MM/yyyy HH:mm:ss")

            Load_Registry_Values()
            Timer2.Start()
            dataloaded = True
            splash_loader.Visible = False
        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub

    Private Sub exit_application()
        Try
            Save_Registry_Values()
            If Worker1.WorkerThread Is Nothing = False Then
                Worker1.WorkerThread.Abort()
                Worker1.Dispose()
            End If
            NotifyIcon1.Dispose()
            Application.Exit()
        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub

    Private Sub Main_Screen_closed(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Closed
        Try
            exit_application()
        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub


    Public Sub WorkerHandler(ByVal Result As String)
        Try
            currently_working = False
            analysisend.Text = Format(Now(), "dd/MM/yyyy hh:mm:ss tt")
            NotifyIcon1.Text = "Resting... "
            run_green_lights()
        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub

    Public Sub WorkerProgressHandler(ByVal switch As Integer)
        Try
            Select Case switch
                Case 1
                    recordsrenumbered.Text = CLng(recordsrenumbered.Text) + 1
                Case 2
                    totalrecords.Text = CLng(totalrecords.Text) + 1
            End Select

        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub

    Private Sub run_worker()
        run_orange_lights()
        analysisstart.Text = ""
        analysisend.Text = ""
        recordsrenumbered.Text = "0"
        totalrecords.Text = "0"
        analysisstart.Text = Format(Now(), "dd/MM/yyyy hh:mm:ss tt")

        Worker1.DatabaseToUse = DatabaseToUse.Text
        Worker1.TableToUse = TableToUse.Text
        Worker1.ColumnSort = ColumnSort.Text
        Worker1.ColumnSortOrder = ColumnSortOrder.Text
        Worker1.ColumnRenumber = ColumnRenumber.Text
        Worker1.SeriesPrefix = SeriesPrefix.Text
        Worker1.SeriesSuffix = SeriesSuffix.Text
        Worker1.SeriesStart = SeriesStart.Value
        Worker1.ChooseThreads(1)
        currently_working = True
    End Sub



    Private Sub MenuItem3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem3.Click
        Try
            exit_application()
        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub

    Private Sub show_application()
        Try
            Me.Opacity = 1

            Me.BringToFront()
            Me.Refresh()
            Me.WindowState = FormWindowState.Normal

        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub

    Private Sub NotifyIcon1_dblclick(ByVal sender As Object, ByVal e As System.EventArgs) Handles NotifyIcon1.DoubleClick
        show_application()
    End Sub
    Private Sub NotifyIcon1_snglclick(ByVal sender As Object, ByVal e As System.EventArgs) Handles NotifyIcon1.Click
        show_application()
    End Sub

    Private Sub MenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem1.Click
        show_application()
    End Sub

    Private Sub Main_Screen_resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Resize
        Try

            If Me.WindowState = FormWindowState.Minimized Then
                NotifyIcon1.Visible = True
                Me.Opacity = 0
            End If
        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub

    Private Sub force_check()
        Try
            NotifyIcon1.Text = "Renumbering Records..."
            If currently_working = False Then
                run_worker()
            End If
        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub


    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        force_check()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Try
            Dim result As DialogResult = OpenFileDialog1.ShowDialog()
            If Not result = DialogResult.Cancel Then
                DatabaseToUse.Text = OpenFileDialog1.FileName
                Try
                    Dim Conn As Data.OleDb.OleDbConnection = New Data.OleDb.OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DatabaseToUse.Text & ";")
                    Conn.Open()
                    Dim schemaTable As DataTable = Conn.GetOleDbSchemaTable(Data.OleDb.OleDbSchemaGuid.Tables, New Object() {Nothing, Nothing, Nothing, "TABLE"})

                    Dim frm_Select_Table_Dialog As New Select_Table_Dialog
                    frm_Select_Table_Dialog.Activate()
                    frm_Select_Table_Dialog.Label1.Text = "Please select the table that will have its records renumbered."
                    frm_Select_Table_Dialog.TableChoice = schemaTable

                    If frm_Select_Table_Dialog.ShowDialog() = DialogResult.OK Then
                        If Not frm_Select_Table_Dialog.Selected_Table.SelectedItem = Nothing Then
                            TableToUse.Text = frm_Select_Table_Dialog.Selected_Table.SelectedItem.ToString
                        End If
                    End If

                    frm_Select_Table_Dialog.Dispose()



                    Dim frm_Select_Table_Dialog2 As New Select_Column_Dialog
                    Dim schemaTable2 As DataTable = Conn.GetOleDbSchemaTable(Data.OleDb.OleDbSchemaGuid.Columns, New Object() {Nothing, Nothing, TableToUse.Text, Nothing})
                    frm_Select_Table_Dialog2.Activate()
                    frm_Select_Table_Dialog2.Label1.Text = "Please select the column that will be used to sort the existing records on."
                    frm_Select_Table_Dialog2.TableChoice = schemaTable2

                    If frm_Select_Table_Dialog2.ShowDialog() = DialogResult.OK Then
                        If Not frm_Select_Table_Dialog2.Selected_Table.SelectedItem = Nothing Then
                            ColumnSort.Text = frm_Select_Table_Dialog2.Selected_Table.SelectedItem.ToString
                        End If
                    End If
                    frm_Select_Table_Dialog2.Dispose()

                    Dim sorted As New Select_Order_Dialog
                    sorted.Activate()
                    sorted.Selected_Table.SelectedIndex = 0

                    If sorted.ShowDialog() = DialogResult.OK Then
                        If Not sorted.Selected_Table.SelectedItem = Nothing Then
                            ColumnSortOrder.Text = sorted.Selected_Table.SelectedItem.ToString
                            If ColumnSortOrder.Text = "Ascending" Then ColumnSortOrder.Text = "asc"
                            If ColumnSortOrder.Text = "Descending" Then ColumnSortOrder.Text = "desc"
                        End If
                    End If
                    sorted.Dispose()

                    Dim frm_Select_Table_Dialog3 As New Select_Column_Dialog
                    'Dim schemaTable2 As DataTable = Conn.GetOleDbSchemaTable(Data.OleDb.OleDbSchemaGuid.Columns, New Object() {Nothing, Nothing, TableToUse.Text, Nothing})
                    frm_Select_Table_Dialog3.Activate()
                    frm_Select_Table_Dialog3.Label1.Text = "Please select the column that will be renumbered using the new series"
                    frm_Select_Table_Dialog3.TableChoice = schemaTable2

                    If frm_Select_Table_Dialog3.ShowDialog() = DialogResult.OK Then
                        If Not frm_Select_Table_Dialog3.Selected_Table.SelectedItem = Nothing Then
                            ColumnRenumber.Text = frm_Select_Table_Dialog3.Selected_Table.SelectedItem.ToString
                        End If
                    End If
                    frm_Select_Table_Dialog3.Dispose()

                    Conn.Close()


                Catch dberror As OleDb.OleDbException
                    Error_Handler(dberror, "Selecting Database Table")

                Catch othererror As Exception
                    Error_Handler(othererror, "Selecting Database Table")
                End Try
            End If

        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Try
            If Worker1.WorkerThread Is Nothing = False Then
                Worker1.WorkerThread.Abort()
                Worker1.Dispose()
            End If
        Catch ex As Exception
            Error_Handler(ex)
        Finally
            currently_working = False
            analysisend.Text = Format(Now(), "dd/MM/yyyy hh:mm:ss tt")
            NotifyIcon1.Text = "Resting... "
            run_green_lights()
        End Try

    End Sub



End Class
