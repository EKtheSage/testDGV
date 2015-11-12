Imports System.Windows.Forms
Imports System.Runtime.InteropServices
Imports ExcelDna.Integration.CustomUI
Imports ExcelDna.Integration
Imports Microsoft.Office.Interop


Module Globals
    Public ReadOnly Property Application As Excel.Application
        Get
            Application = CType(ExcelDnaUtil.Application, Excel.Application)
        End Get
    End Property
End Module

Public Class dgvForm
    Inherits Form

    Friend WithEvents TabControl1 As TabControl

    Friend WithEvents tpPlaceHolder As TabPage

    Friend WithEvents tpTriangle As TabPage

    Friend WithEvents dgvTriangle As DataGridView

    Public Sub New()
        InitializeComponent()

        tpPlaceHolder.Select()
    End Sub

    Private Sub InitializeComponent()
        Me.TabControl1 = New System.Windows.Forms.TabControl()
        Me.tpPlaceHolder = New System.Windows.Forms.TabPage()
        Me.tpTriangle = New System.Windows.Forms.TabPage()
        Me.dgvTriangle = New System.Windows.Forms.DataGridView()
        Me.TabControl1.SuspendLayout()
        Me.tpTriangle.SuspendLayout()
        CType(Me.dgvTriangle, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'TabControl1
        '
        Me.TabControl1.Controls.Add(Me.tpPlaceHolder)
        Me.TabControl1.Controls.Add(Me.tpTriangle)
        Me.TabControl1.Location = New System.Drawing.Point(12, 51)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(692, 388)
        Me.TabControl1.TabIndex = 0
        '
        'tpPlaceHolder
        '
        Me.tpPlaceHolder.Location = New System.Drawing.Point(4, 22)
        Me.tpPlaceHolder.Name = "tpPlaceHolder"
        Me.tpPlaceHolder.Padding = New System.Windows.Forms.Padding(3)
        Me.tpPlaceHolder.Size = New System.Drawing.Size(684, 362)
        Me.tpPlaceHolder.TabIndex = 0
        Me.tpPlaceHolder.Text = "Placeholder"
        Me.tpPlaceHolder.UseVisualStyleBackColor = True
        '
        'tpTriangle
        '
        Me.tpTriangle.Controls.Add(Me.dgvTriangle)
        Me.tpTriangle.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.tpTriangle.Location = New System.Drawing.Point(4, 22)
        Me.tpTriangle.Name = "tpTriangle"
        Me.tpTriangle.Padding = New System.Windows.Forms.Padding(3)
        Me.tpTriangle.Size = New System.Drawing.Size(684, 362)
        Me.tpTriangle.TabIndex = 1
        Me.tpTriangle.Text = "Triangle"
        Me.tpTriangle.UseVisualStyleBackColor = True
        '
        'dgvTriangle
        '
        Me.dgvTriangle.AllowUserToAddRows = False
        Me.dgvTriangle.AllowUserToDeleteRows = False
        Me.dgvTriangle.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvTriangle.Location = New System.Drawing.Point(7, 41)
        Me.dgvTriangle.Name = "dgvTriangle"
        Me.dgvTriangle.Size = New System.Drawing.Size(580, 281)
        Me.dgvTriangle.TabIndex = 0
        '
        'dgvForm
        '
        Me.ClientSize = New System.Drawing.Size(716, 451)
        Me.Controls.Add(Me.TabControl1)
        Me.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Name = "dgvForm"
        Me.Text = "DataGridView Test"
        Me.TabControl1.ResumeLayout(False)
        Me.tpTriangle.ResumeLayout(False)
        CType(Me.dgvTriangle, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    Private Sub dgvForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        AddHandler dgvTriangle.DataBindingComplete, AddressOf dgvTriangle_DataBindingComplete
        'First get the body of triangle
        Dim rng As String = "Sheet1!B2:K11"
        Dim wkst As ExcelReference = CType(XlCall.Excel(XlCall.xlfEvaluate, rng), ExcelReference)
        Dim selectVal As Object(,) = CType(wkst.GetValue, Object(,))

        Dim dt As Data.DataTable = New Data.DataTable()


        'need to multiply ATA to triangle
        For i As Integer = 0 To selectVal.GetUpperBound(0)
            For j As Integer = 0 To selectVal.GetUpperBound(1)

            Next
        Next


        For i As Integer = 0 To selectVal.GetUpperBound(1)
            dt.Columns.Add("Age " & i + 1, GetType(Double))
        Next
        For i As Integer = 0 To selectVal.GetUpperBound(0)
            Dim row As DataRow = dt.NewRow
            For j As Integer = 0 To selectVal.GetUpperBound(1)
                row.Item(j) = CType(selectVal(i, j), Double)
            Next
            dt.Rows.Add(row)
        Next

        'get column headers
        rng = "Sheet1!B1:K1"
        wkst = CType(XlCall.Excel(XlCall.xlfEvaluate, rng), ExcelReference)
        selectVal = CType(wkst.GetValue, Object(,))

        dgvTriangle.DataSource = dt
        dgvTriangle.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders
        For i As Integer = 0 To selectVal.GetUpperBound(1)
            dgvTriangle.Columns(i).HeaderText = dt.Columns(i).ColumnName
        Next


    End Sub

    Private Sub dgvTriangle_DataBindingComplete(sender As Object, e As DataGridViewBindingCompleteEventArgs) Handles dgvTriangle.DataBindingComplete
        Dim style As DataGridViewCellStyle = New DataGridViewCellStyle
        style.Alignment = DataGridViewContentAlignment.MiddleRight
        Dim rng As Excel.Range = CType(Application.ActiveWorkbook.Worksheets("Sheet1"),
            Excel.Worksheet).Range("A2:A11")

        For Each row As DataGridViewRow In dgvTriangle.Rows
            row.HeaderCell.Value = CType(rng.Cells(row.Index + 1, 1), Excel.Range).Value.ToString
            row.HeaderCell.Style = style
            row.Resizable = DataGridViewTriState.False
        Next

        dgvTriangle.ClearSelection()
        dgvTriangle.CurrentCell = Nothing
        dgvTriangle.ResumeLayout()
    End Sub
End Class

<ComVisible(True)>
Public Class myRibbon
    Inherits ExcelRibbon

End Class

Public Module macros
    Public frm As dgvForm = New dgvForm

    Public Sub loadForm()
        frm.Show()
    End Sub
End Module