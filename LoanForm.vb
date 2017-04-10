Public Class LoanForm
    Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

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
    Friend WithEvents txtAmount As System.Windows.Forms.TextBox
    Friend WithEvents txtRate As System.Windows.Forms.TextBox
    Friend WithEvents txtDuration As System.Windows.Forms.TextBox
    Friend WithEvents bttnShowPayment As System.Windows.Forms.Button
    Friend WithEvents txtPayment As System.Windows.Forms.TextBox
    Friend WithEvents chkPayEarly As System.Windows.Forms.CheckBox
    Friend WithEvents bttnExit As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.Container

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.txtPayment = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.txtRate = New System.Windows.Forms.TextBox()
        Me.bttnExit = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.chkPayEarly = New System.Windows.Forms.CheckBox()
        Me.txtDuration = New System.Windows.Forms.TextBox()
        Me.bttnShowPayment = New System.Windows.Forms.Button()
        Me.txtAmount = New System.Windows.Forms.TextBox()
        Me.SuspendLayout()
        '
        'txtPayment
        '
        Me.txtPayment.Location = New System.Drawing.Point(160, 152)
        Me.txtPayment.Name = "txtPayment"
        Me.txtPayment.ReadOnly = True
        Me.txtPayment.Size = New System.Drawing.Size(112, 23)
        Me.txtPayment.TabIndex = 8
        Me.txtPayment.Text = ""
        '
        'Label4
        '
        Me.Label4.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(8, 152)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(128, 23)
        Me.Label4.TabIndex = 7
        Me.Label4.Text = "Monthly Payment"
        '
        'txtRate
        '
        Me.txtRate.BackColor = System.Drawing.Color.AntiqueWhite
        Me.txtRate.Location = New System.Drawing.Point(160, 72)
        Me.txtRate.Name = "txtRate"
        Me.txtRate.TabIndex = 5
        Me.txtRate.Text = ""
        '
        'bttnExit
        '
        Me.bttnExit.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.bttnExit.Font = New System.Drawing.Font("Verdana", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.bttnExit.Location = New System.Drawing.Point(176, 208)
        Me.bttnExit.Name = "bttnExit"
        Me.bttnExit.Size = New System.Drawing.Size(136, 32)
        Me.bttnExit.TabIndex = 10
        Me.bttnExit.Text = "Exit"
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(8, 8)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(136, 23)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Loan Amount"
        '
        'Label2
        '
        Me.Label2.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(8, 40)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(152, 23)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "Duration (in months)"
        '
        'Label3
        '
        Me.Label3.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(8, 72)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(128, 23)
        Me.Label3.TabIndex = 4
        Me.Label3.Text = "Interest Rate"
        '
        'chkPayEarly
        '
        Me.chkPayEarly.BackColor = System.Drawing.SystemColors.ActiveBorder
        Me.chkPayEarly.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkPayEarly.Location = New System.Drawing.Point(8, 112)
        Me.chkPayEarly.Name = "chkPayEarly"
        Me.chkPayEarly.Size = New System.Drawing.Size(168, 16)
        Me.chkPayEarly.TabIndex = 6
        Me.chkPayEarly.Text = "Early Payment"
        '
        'txtDuration
        '
        Me.txtDuration.BackColor = System.Drawing.Color.AntiqueWhite
        Me.txtDuration.Location = New System.Drawing.Point(160, 40)
        Me.txtDuration.Name = "txtDuration"
        Me.txtDuration.TabIndex = 3
        Me.txtDuration.Text = ""
        '
        'bttnShowPayment
        '
        Me.bttnShowPayment.Font = New System.Drawing.Font("Verdana", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.bttnShowPayment.Location = New System.Drawing.Point(8, 208)
        Me.bttnShowPayment.Name = "bttnShowPayment"
        Me.bttnShowPayment.Size = New System.Drawing.Size(136, 32)
        Me.bttnShowPayment.TabIndex = 9
        Me.bttnShowPayment.Text = "Show Payment"
        '
        'txtAmount
        '
        Me.txtAmount.BackColor = System.Drawing.Color.AntiqueWhite
        Me.txtAmount.Location = New System.Drawing.Point(160, 8)
        Me.txtAmount.Name = "txtAmount"
        Me.txtAmount.TabIndex = 1
        Me.txtAmount.Text = ""
        '
        'LoanForm
        '
        Me.AcceptButton = Me.bttnShowPayment
        Me.AutoScaleBaseSize = New System.Drawing.Size(7, 16)
        Me.CancelButton = Me.bttnExit
        Me.ClientSize = New System.Drawing.Size(320, 245)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.Label4, Me.Label3, Me.Label2, Me.Label1, Me.bttnExit, Me.txtPayment, Me.bttnShowPayment, Me.chkPayEarly, Me.txtDuration, Me.txtRate, Me.txtAmount})
        Me.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Name = "LoanForm"
        Me.Text = "Loan Calculator"
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub bttnShowPayment_Click(ByVal sender As System.Object, _
                   ByVal e As System.EventArgs) Handles bttnShowPayment.Click
        Dim Payment As Single
        Dim LoanIRate As Single
        Dim LoanDuration As Integer
        Dim LoanAmount As Integer
        Dim payEarly As DueDate

        If IsNumeric(txtAmount.Text) Then
            LoanAmount = txtAmount.Text
        Else
            MsgBox("Please enter a valid amount")
            Exit Sub
        End If
        If IsNumeric(txtRate.Text) Then
            LoanIRate = 0.01 * txtRate.Text / 12
        Else
            MsgBox("Invalid interest rate, please re-enter")
            Exit Sub
        End If
        If IsNumeric(txtDuration.Text) Then
            LoanDuration = txtDuration.Text
        Else
            MsgBox("Please specify the loan’s duration " & _
                    "as a number of months")
            Exit Sub
        End If

        If chkPayEarly.Checked Then
            payEarly = DueDate.BegOfPeriod
        Else
            payEarly = DueDate.EndOfPeriod
        End If
        Payment = Pmt(LoanIRate, LoanDuration, _
                  -LoanAmount, 0, payEarly)
        txtPayment.Text = Payment.ToString("#.00")
    End Sub

    Private Sub bttnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bttnExit.Click
        End
    End Sub
End Class
