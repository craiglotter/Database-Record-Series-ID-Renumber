Public Class Select_Order_Dialog
    Inherits System.Windows.Forms.Form


#Region " Windows Form Designer generated code "

    'Public Sub New()
    '    MyBase.New()

    '    'This call is required by the Windows Form Designer.
    '    InitializeComponent()

    '    'Add any initialization after the InitializeComponent() call
    '    Type = 1
    'End Sub

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
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
    Friend WithEvents Selected_Table As System.Windows.Forms.ComboBox
    Friend WithEvents accept_button As System.Windows.Forms.Button
    Protected Friend WithEvents Label1 As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(Select_Order_Dialog))
        Me.accept_button = New System.Windows.Forms.Button
        Me.Selected_Table = New System.Windows.Forms.ComboBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.SuspendLayout()
        '
        'accept_button
        '
        Me.accept_button.BackColor = System.Drawing.Color.FromArgb(CType(224, Byte), CType(224, Byte), CType(224, Byte))
        Me.accept_button.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.accept_button.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.accept_button.Location = New System.Drawing.Point(160, 24)
        Me.accept_button.Name = "accept_button"
        Me.accept_button.TabIndex = 0
        Me.accept_button.Text = "Accept"
        '
        'Selected_Table
        '
        Me.Selected_Table.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.Selected_Table.Items.AddRange(New Object() {"Ascending", "Descending"})
        Me.Selected_Table.Location = New System.Drawing.Point(8, 24)
        Me.Selected_Table.MaxDropDownItems = 20
        Me.Selected_Table.Name = "Selected_Table"
        Me.Selected_Table.Size = New System.Drawing.Size(144, 21)
        Me.Selected_Table.Sorted = True
        Me.Selected_Table.TabIndex = 1
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.Color.Transparent
        Me.Label1.Location = New System.Drawing.Point(8, 8)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(232, 16)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "Please select the sort order required"
        '
        'Select_Order_Dialog
        '
        Me.AcceptButton = Me.accept_button
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.Color.FromArgb(CType(192, Byte), CType(255, Byte), CType(255, Byte))
        Me.BackgroundImage = CType(resources.GetObject("$this.BackgroundImage"), System.Drawing.Image)
        Me.ClientSize = New System.Drawing.Size(248, 62)
        Me.ControlBox = False
        Me.Controls.Add(Me.Selected_Table)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.accept_button)
        Me.ForeColor = System.Drawing.Color.Black
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MaximumSize = New System.Drawing.Size(256, 96)
        Me.MinimizeBox = False
        Me.MinimumSize = New System.Drawing.Size(256, 96)
        Me.Name = "Select_Order_Dialog"
        Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Select the Sort Order"
        Me.ResumeLayout(False)

    End Sub

#End Region



End Class
