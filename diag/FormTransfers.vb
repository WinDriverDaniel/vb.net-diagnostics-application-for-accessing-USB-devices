' Note: This code sample is provided AS-IS and as a guiding sample only.

Imports System.Windows.Forms

Public Class FormTransfers
    Inherits System.Windows.Forms.Form

    Private m_bIsControl As Boolean
    Private m_bIsRead As Boolean
    Private m_dwBuffSize As UInt32
    Private m_buffer() As Byte = Nothing
    Private m_pSetupPacket(7) As Byte

#Region " Windows Form Designer generated code "

    Public Sub New(ByVal bIsRead As Boolean, ByVal bIsControl As Boolean)
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        pnlBufSize.Enabled = Not (bIsControl)
        txtBufSize.Enabled = Not (bIsControl)

        pnlData.Enabled = Not (bIsRead)
        txtData.Enabled = Not (bIsRead)

        pnlSetupPacket.Enabled = bIsControl
        txtRequest.Enabled = bIsControl
        txtType.Enabled = bIsControl
        txtwIndex.Enabled = bIsControl
        txtwLength.Enabled = bIsControl
        txtwValue.Enabled = bIsControl
        m_bIsControl = bIsControl
        m_bIsRead = bIsRead

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
    Friend WithEvents pnlData As System.Windows.Forms.Panel
    Friend WithEvents txtData As System.Windows.Forms.TextBox
    Friend WithEvents lblData As System.Windows.Forms.Label
    Friend WithEvents pnlBufSize As System.Windows.Forms.Panel
    Friend WithEvents txtBufSize As System.Windows.Forms.TextBox
    Friend WithEvents lblBufSize As System.Windows.Forms.Label
    Friend WithEvents btCancel As System.Windows.Forms.Button
    Friend WithEvents pnlSetupPacket As System.Windows.Forms.Panel
    Friend WithEvents txtwLength As System.Windows.Forms.TextBox
    Friend WithEvents lblwLength As System.Windows.Forms.Label
    Friend WithEvents txtwIndex As System.Windows.Forms.TextBox
    Friend WithEvents lblwIndex As System.Windows.Forms.Label
    Friend WithEvents txtwValue As System.Windows.Forms.TextBox
    Friend WithEvents lblwValue As System.Windows.Forms.Label
    Friend WithEvents txtRequest As System.Windows.Forms.TextBox
    Friend WithEvents lblRequest As System.Windows.Forms.Label
    Friend WithEvents txtType As System.Windows.Forms.TextBox
    Friend WithEvents lblType As System.Windows.Forms.Label
    Friend WithEvents lblSetupPacket As System.Windows.Forms.Label
    Friend WithEvents btSubmit As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.pnlData = New System.Windows.Forms.Panel
        Me.txtData = New System.Windows.Forms.TextBox
        Me.lblData = New System.Windows.Forms.Label
        Me.pnlBufSize = New System.Windows.Forms.Panel
        Me.txtBufSize = New System.Windows.Forms.TextBox
        Me.lblBufSize = New System.Windows.Forms.Label
        Me.btCancel = New System.Windows.Forms.Button
        Me.pnlSetupPacket = New System.Windows.Forms.Panel
        Me.txtwLength = New System.Windows.Forms.TextBox
        Me.lblwLength = New System.Windows.Forms.Label
        Me.txtwIndex = New System.Windows.Forms.TextBox
        Me.lblwIndex = New System.Windows.Forms.Label
        Me.txtwValue = New System.Windows.Forms.TextBox
        Me.lblwValue = New System.Windows.Forms.Label
        Me.txtRequest = New System.Windows.Forms.TextBox
        Me.lblRequest = New System.Windows.Forms.Label
        Me.txtType = New System.Windows.Forms.TextBox
        Me.lblType = New System.Windows.Forms.Label
        Me.lblSetupPacket = New System.Windows.Forms.Label
        Me.btSubmit = New System.Windows.Forms.Button
        Me.pnlData.SuspendLayout()
        Me.pnlBufSize.SuspendLayout()
        Me.pnlSetupPacket.SuspendLayout()
        Me.SuspendLayout()
        '
        'pnlData
        '
        Me.pnlData.Controls.Add(Me.txtData)
        Me.pnlData.Controls.Add(Me.lblData)
        Me.pnlData.Location = New System.Drawing.Point(12, 46)
        Me.pnlData.Name = "pnlData"
        Me.pnlData.Size = New System.Drawing.Size(360, 32)
        Me.pnlData.TabIndex = 26
        '
        'txtData
        '
        Me.txtData.Enabled = False
        Me.txtData.Location = New System.Drawing.Point(112, 8)
        Me.txtData.Name = "txtData"
        Me.txtData.Size = New System.Drawing.Size(240, 20)
        Me.txtData.TabIndex = 3
        Me.txtData.Text = ""
        '
        'lblData
        '
        Me.lblData.Font = New System.Drawing.Font("Microsoft Sans Serif", _
            8.25!, CType((System.Drawing.FontStyle.Bold Or _
            System.Drawing.FontStyle.Underline), System.Drawing.FontStyle), _
            System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblData.Location = New System.Drawing.Point(8, 8)
        Me.lblData.Name = "lblData"
        Me.lblData.Size = New System.Drawing.Size(64, 16)
        Me.lblData.TabIndex = 2
        Me.lblData.Text = "Data (hex):"
        '
        'pnlBufSize
        '
        Me.pnlBufSize.Controls.Add(Me.txtBufSize)
        Me.pnlBufSize.Controls.Add(Me.lblBufSize)
        Me.pnlBufSize.Location = New System.Drawing.Point(12, 14)
        Me.pnlBufSize.Name = "pnlBufSize"
        Me.pnlBufSize.Size = New System.Drawing.Size(168, 32)
        Me.pnlBufSize.TabIndex = 25
        '
        'txtBufSize
        '
        Me.txtBufSize.Location = New System.Drawing.Point(112, 8)
        Me.txtBufSize.Name = "txtBufSize"
        Me.txtBufSize.Size = New System.Drawing.Size(40, 20)
        Me.txtBufSize.TabIndex = 1
        Me.txtBufSize.Text = ""
        Me.txtBufSize.TextAlign = _
            System.Windows.Forms.HorizontalAlignment.Center
        '
        'lblBufSize
        '
        Me.lblBufSize.Font = New System.Drawing.Font("Microsoft Sans Serif", _
            8.25!, CType((System.Drawing.FontStyle.Bold Or _
            System.Drawing.FontStyle.Underline), System.Drawing.FontStyle), _
            System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblBufSize.Location = New System.Drawing.Point(8, 8)
        Me.lblBufSize.Name = "lblBufSize"
        Me.lblBufSize.Size = New System.Drawing.Size(96, 16)
        Me.lblBufSize.TabIndex = 0
        Me.lblBufSize.Text = "Buffer size (hex):"
        '
        'btCancel
        '
        Me.btCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btCancel.Font = New System.Drawing.Font("Microsoft Sans Serif", _
            8.25!, System.Drawing.FontStyle.Bold, _
            System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btCancel.Location = New System.Drawing.Point(108, 166)
        Me.btCancel.Name = "btCancel"
        Me.btCancel.Size = New System.Drawing.Size(75, 32)
        Me.btCancel.TabIndex = 24
        Me.btCancel.Text = "Cancel"
        '
        'pnlSetupPacket
        '
        Me.pnlSetupPacket.Controls.Add(Me.txtwLength)
        Me.pnlSetupPacket.Controls.Add(Me.lblwLength)
        Me.pnlSetupPacket.Controls.Add(Me.txtwIndex)
        Me.pnlSetupPacket.Controls.Add(Me.lblwIndex)
        Me.pnlSetupPacket.Controls.Add(Me.txtwValue)
        Me.pnlSetupPacket.Controls.Add(Me.lblwValue)
        Me.pnlSetupPacket.Controls.Add(Me.txtRequest)
        Me.pnlSetupPacket.Controls.Add(Me.lblRequest)
        Me.pnlSetupPacket.Controls.Add(Me.txtType)
        Me.pnlSetupPacket.Controls.Add(Me.lblType)
        Me.pnlSetupPacket.Controls.Add(Me.lblSetupPacket)
        Me.pnlSetupPacket.Location = New System.Drawing.Point(12, 78)
        Me.pnlSetupPacket.Name = "pnlSetupPacket"
        Me.pnlSetupPacket.Size = New System.Drawing.Size(312, 72)
        Me.pnlSetupPacket.TabIndex = 22
        '
        'txtwLength
        '
        Me.txtwLength.Enabled = False
        Me.txtwLength.Location = New System.Drawing.Point(248, 48)
        Me.txtwLength.MaxLength = 4
        Me.txtwLength.Name = "txtwLength"
        Me.txtwLength.Size = New System.Drawing.Size(48, 20)
        Me.txtwLength.TabIndex = 8
        Me.txtwLength.Text = ""
        Me.txtwLength.TextAlign = _
            System.Windows.Forms.HorizontalAlignment.Center
        '
        'lblwLength
        '
        Me.lblwLength.Font = New System.Drawing.Font("Microsoft Sans Serif", _
            8.25!, System.Drawing.FontStyle.Bold, _
            System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblwLength.Location = New System.Drawing.Point(248, 32)
        Me.lblwLength.Name = "lblwLength"
        Me.lblwLength.Size = New System.Drawing.Size(48, 16)
        Me.lblwLength.TabIndex = 12
        Me.lblwLength.Text = "wLength"
        '
        'txtwIndex
        '
        Me.txtwIndex.Enabled = False
        Me.txtwIndex.Location = New System.Drawing.Point(184, 48)
        Me.txtwIndex.MaxLength = 4
        Me.txtwIndex.Name = "txtwIndex"
        Me.txtwIndex.Size = New System.Drawing.Size(48, 20)
        Me.txtwIndex.TabIndex = 7
        Me.txtwIndex.Text = ""
        Me.txtwIndex.TextAlign = _
            System.Windows.Forms.HorizontalAlignment.Center
        '
        'lblwIndex
        '
        Me.lblwIndex.Font = New System.Drawing.Font("Microsoft Sans Serif", _
            8.25!, System.Drawing.FontStyle.Bold, _
            System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblwIndex.Location = New System.Drawing.Point(184, 32)
        Me.lblwIndex.Name = "lblwIndex"
        Me.lblwIndex.Size = New System.Drawing.Size(48, 16)
        Me.lblwIndex.TabIndex = 11
        Me.lblwIndex.Text = "wIndex"
        '
        'txtwValue
        '
        Me.txtwValue.Enabled = False
        Me.txtwValue.Location = New System.Drawing.Point(120, 48)
        Me.txtwValue.MaxLength = 4
        Me.txtwValue.Name = "txtwValue"
        Me.txtwValue.Size = New System.Drawing.Size(48, 20)
        Me.txtwValue.TabIndex = 6
        Me.txtwValue.Text = ""
        Me.txtwValue.TextAlign = _
            System.Windows.Forms.HorizontalAlignment.Center
        '
        'lblwValue
        '
        Me.lblwValue.Font = New System.Drawing.Font("Microsoft Sans Serif", _
            8.25!, System.Drawing.FontStyle.Bold, _
            System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblwValue.Location = New System.Drawing.Point(120, 32)
        Me.lblwValue.Name = "lblwValue"
        Me.lblwValue.Size = New System.Drawing.Size(48, 16)
        Me.lblwValue.TabIndex = 10
        Me.lblwValue.Text = "wValue"
        '
        'txtRequest
        '
        Me.txtRequest.Enabled = False
        Me.txtRequest.Location = New System.Drawing.Point(64, 48)
        Me.txtRequest.MaxLength = 2
        Me.txtRequest.Name = "txtRequest"
        Me.txtRequest.Size = New System.Drawing.Size(32, 20)
        Me.txtRequest.TabIndex = 5
        Me.txtRequest.Text = ""
        Me.txtRequest.TextAlign = _
            System.Windows.Forms.HorizontalAlignment.Center
        '
        'lblRequest
        '
        Me.lblRequest.Font = New System.Drawing.Font("Microsoft Sans Serif", _
            8.25!, System.Drawing.FontStyle.Bold, _
            System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblRequest.Location = New System.Drawing.Point(56, 32)
        Me.lblRequest.Name = "lblRequest"
        Me.lblRequest.Size = New System.Drawing.Size(48, 16)
        Me.lblRequest.TabIndex = 9
        Me.lblRequest.Text = "Request"
        '
        'txtType
        '
        Me.txtType.Enabled = False
        Me.txtType.Location = New System.Drawing.Point(8, 48)
        Me.txtType.MaxLength = 2
        Me.txtType.Name = "txtType"
        Me.txtType.Size = New System.Drawing.Size(32, 20)
        Me.txtType.TabIndex = 4
        Me.txtType.Text = ""
        Me.txtType.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'lblType
        '
        Me.lblType.Font = New System.Drawing.Font("Microsoft Sans Serif", _
            8.25!, System.Drawing.FontStyle.Bold, _
            System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblType.Location = New System.Drawing.Point(8, 32)
        Me.lblType.Name = "lblType"
        Me.lblType.Size = New System.Drawing.Size(32, 16)
        Me.lblType.TabIndex = 3
        Me.lblType.Text = "Type"
        '
        'lblSetupPacket
        '
        Me.lblSetupPacket.Font = New _
            System.Drawing.Font("Microsoft Sans Serif", 8.25!, _
            CType((System.Drawing.FontStyle.Bold Or _
            System.Drawing.FontStyle.Underline), System.Drawing.FontStyle), _
            System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblSetupPacket.Location = New System.Drawing.Point(8, 8)
        Me.lblSetupPacket.Name = "lblSetupPacket"
        Me.lblSetupPacket.Size = New System.Drawing.Size(296, 16)
        Me.lblSetupPacket.TabIndex = 17
        Me.lblSetupPacket.Text = "Setup Packet:"
        '
        'btSubmit
        '
        Me.btSubmit.Font = New System.Drawing.Font("Microsoft Sans Serif", _
            8.25!, System.Drawing.FontStyle.Bold, _
            System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btSubmit.Location = New System.Drawing.Point(20, 166)
        Me.btSubmit.Name = "btSubmit"
        Me.btSubmit.Size = New System.Drawing.Size(75, 32)
        Me.btSubmit.TabIndex = 23
        Me.btSubmit.Text = "Submit"
        '
        'FormTransfers
        '
        Me.AcceptButton = Me.btSubmit
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.CancelButton = Me.btCancel
        Me.ClientSize = New System.Drawing.Size(384, 213)
        Me.Controls.Add(Me.pnlData)
        Me.Controls.Add(Me.pnlBufSize)
        Me.Controls.Add(Me.btCancel)
        Me.Controls.Add(Me.pnlSetupPacket)
        Me.Controls.Add(Me.btSubmit)
        Me.Name = "FormTransfers"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Read/Write device's Pipes"
        Me.pnlData.ResumeLayout(False)
        Me.pnlBufSize.ResumeLayout(False)
        Me.pnlSetupPacket.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Public Function GetInput(ByRef dwBuffSize As UInt32, ByRef buffer() As _
        Byte, ByRef pSetupPacket() As Byte) As Boolean

        Dim result As DialogResult = ShowDialog()
        While (result = System.Windows.Forms.DialogResult.Retry)
            result = ShowDialog()
        End While
        If (result <> System.Windows.Forms.DialogResult.OK) Then
            Return False
        End If

        dwBuffSize = m_dwBuffSize
        buffer = m_buffer
        pSetupPacket = m_pSetupPacket

        Return True

    End Function

    Private Sub TranslateInput()

    Dim iBuffSize As Int32
        If (m_bIsControl) Then
            GetSetupPacketData()
            iBuffSize = Convert.ToInt32(txtwLength.Text, 16)
        Else
            iBuffSize = Convert.ToInt32(txtBufSize.Text, 16)
        End If

        If (iBuffSize > 0) Then
            ReDim m_buffer(iBuffSize - 1)
        End If

        If (Not (m_bIsRead)) Then
            'padding the first bytes if necessary
            Dim str As String = PadBuffer(txtData.Text, txtData.Text.Length, _
                2 * iBuffSize)
            Dim i As Int32
            For i = 0 To (iBuffSize - 1)
                m_buffer(i) = Convert.ToByte(str.Substring(2 * i, 2), 16)
            Next i
        End If
    m_dwBuffSize = Convert.ToUInt32(iBuffSize)
    End Sub

    Private Function PadBuffer(ByVal str As String, ByVal fromIndex As Int32, _
        ByVal toIndex As Int32) As String
        Dim i As Int32

        For i = fromIndex To (toIndex - 1)
            str = "0" & str
        Next i

        Return str
    End Function

    Private Sub GetSetupPacketData()
        Dim _type As String = PadBuffer(txtType.Text, txtType.Text.Length, 2)
        Dim _request As String = PadBuffer(txtRequest.Text, _
            txtRequest.Text.Length, 2)
        Dim _wValue As String = PadBuffer(txtwValue.Text, _
            txtwValue.Text.Length, 4)
        Dim _wIndex As String = PadBuffer(txtwIndex.Text, _
            txtwIndex.Text.Length, 4)
        Dim _wLength As String = PadBuffer(txtwLength.Text, _
            txtwLength.Text.Length, 4)


        m_pSetupPacket(0) = Convert.ToByte(_type, 16)
        m_pSetupPacket(1) = Convert.ToByte(_request, 16)
        m_pSetupPacket(2) = Convert.ToByte(_wValue.Substring(2, 2), 16)
        m_pSetupPacket(3) = Convert.ToByte(_wValue.Substring(0, 2), 16)
        m_pSetupPacket(4) = Convert.ToByte(_wIndex.Substring(2, 2), 16)
        m_pSetupPacket(5) = Convert.ToByte(_wIndex.Substring(0, 2), 16)
        m_pSetupPacket(6) = Convert.ToByte(_wLength.Substring(2, 2), 16)
        m_pSetupPacket(7) = Convert.ToByte(_wLength.Substring(0, 2), 16)
    End Sub

    Private Sub btSubmit_Click(ByVal sender As System.Object, ByVal e As _
        System.EventArgs) Handles btSubmit.Click
        DialogResult = System.Windows.Forms.DialogResult.OK
        Try
            TranslateInput()
        Catch ex As Exception
            MessageBox.Show(String.Concat("The text is not a valid hex number.", _
                "Please re-enter, or press Cancel to exit"), _
                "Input Entry Error", MessageBoxButtons.OK, _
                MessageBoxIcon.Exclamation)
            DialogResult = System.Windows.Forms.DialogResult.Retry
        End Try
    End Sub

    Private Sub SetupPacketInputChanged(ByVal txtInput As TextBox)
        If (txtInput.Text.Length = txtInput.MaxLength) Then
            SelectNextControl(txtInput, True, True, True, False)
        End If
    End Sub

    Private Sub txtType_TextChanged(ByVal sender As System.Object, ByVal e As _
        System.EventArgs) Handles txtType.TextChanged
        SetupPacketInputChanged(txtType)
    End Sub

    Private Sub txtRequest_TextChanged(ByVal sender As System.Object, ByVal e _
        As System.EventArgs) Handles txtRequest.TextChanged
        SetupPacketInputChanged(txtRequest)
    End Sub

    Private Sub txtwValue_TextChanged(ByVal sender As System.Object, ByVal e _
        As System.EventArgs) Handles txtwValue.TextChanged
        SetupPacketInputChanged(txtwValue)
    End Sub

    Private Sub txtwIndex_TextChanged(ByVal sender As System.Object, ByVal e _
        As System.EventArgs) Handles txtwIndex.TextChanged
        SetupPacketInputChanged(txtwIndex)
    End Sub

    Private Sub txtwLength_TextChanged(ByVal sender As System.Object, ByVal e _
        As System.EventArgs) Handles txtwLength.TextChanged
        SetupPacketInputChanged(txtwLength)
    End Sub
End Class

