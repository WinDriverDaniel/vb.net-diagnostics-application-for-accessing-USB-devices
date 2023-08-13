' Note: This code sample is provided AS-IS and as a guiding sample only.

Option Explicit On
Imports System.Windows.Forms
Imports System.Drawing
Imports Jungo.wdapi_dotnet
Imports wdu_err = Jungo.wdapi_dotnet.WD_ERROR_CODES
Imports Jungo.usb_lib

Public Class UsbSample
    Inherits System.Windows.Forms.Form

    Delegate Sub safeLogTextCallBack(ByVal sMsg As String)
    Delegate Sub safeButtonsAccessCallBack()
    Private Const APP_NAME As String = ".NET USB Sample"

    ' WinDriver license registration string
    ' TODO: When using a registered WinDriver version, replace the license string
    '     below with the development license in order to use on the development
    '     machine.
    '     Once you require to distribute the driver's package to other machines,
    '     please replace the string with a distribution license */
    Private Const DEFAULT_LICENSE_STRING As String = "12345abcde1234.license"

    'TODO: If you have renamed the WinDriver kernel module
    '(windrvr1610.sys), change the driver name below accordingly
    Private Const DEFAULT_DRIVER_NAME As String = "windrvr1610"
    Private Const DEFAULT_VENDOR_ID As Short = 0
    Private Const DEFAULT_PRODUCT_ID As Short = 0
    Public Const TIME_OUT As Int32 = 30000
    Private uDevManager As UsbDeviceManager
    Private Delegate Sub D_ATTACH_GUI_CALLBACK(ByVal pDev As UsbDevice)
    Private Delegate Sub D_DETACH_GUI_CALLBACK(ByVal pDev As UsbDevice)


#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()
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
    Friend WithEvents btExit As System.Windows.Forms.Button
    Friend WithEvents lblNumDevices As System.Windows.Forms.Label
    Friend WithEvents lblNumDevicesText As System.Windows.Forms.Label
    Friend WithEvents btLogClear As System.Windows.Forms.Button
    Friend WithEvents btActiveAltSetChange As System.Windows.Forms.Button
    Friend WithEvents lbLog As System.Windows.Forms.Label
    Friend WithEvents txtLog As System.Windows.Forms.TextBox
    Friend WithEvents tabDevices As System.Windows.Forms.TabControl
    Friend WithEvents btPipeWrite As System.Windows.Forms.Button
    Friend WithEvents btPipeRead As System.Windows.Forms.Button
    Friend WithEvents btPipeListen As System.Windows.Forms.Button
    Friend WithEvents btPipeReset As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.btExit = New System.Windows.Forms.Button
        Me.lblNumDevices = New System.Windows.Forms.Label
        Me.lblNumDevicesText = New System.Windows.Forms.Label
        Me.btLogClear = New System.Windows.Forms.Button
        Me.btActiveAltSetChange = New System.Windows.Forms.Button
        Me.lbLog = New System.Windows.Forms.Label
        Me.txtLog = New System.Windows.Forms.TextBox
        Me.tabDevices = New System.Windows.Forms.TabControl
        Me.btPipeWrite = New System.Windows.Forms.Button
        Me.btPipeRead = New System.Windows.Forms.Button
        Me.btPipeListen = New System.Windows.Forms.Button
        Me.btPipeReset = New System.Windows.Forms.Button
        Me.SuspendLayout()
        '
        'btExit
        '
        Me.btExit.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btExit.Location = New System.Drawing.Point(688, 536)
        Me.btExit.Name = "btExit"
        Me.btExit.Size = New System.Drawing.Size(88, 40)
        Me.btExit.TabIndex = 35
        Me.btExit.Text = "&Exit"
        '
        'lblNumDevices
        '
        Me.lblNumDevices.Font = New _
            System.Drawing.Font("Microsoft Sans Serif", 9.0!, _
            System.Drawing.FontStyle.Bold, _
            System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblNumDevices.Location = New System.Drawing.Point(172, 26)
        Me.lblNumDevices.Name = "lblNumDevices"
        Me.lblNumDevices.Size = New System.Drawing.Size(16, 16)
        Me.lblNumDevices.TabIndex = 25
        Me.lblNumDevices.Text = "0"
        Me.lblNumDevices.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblNumDevicesText
        '
        Me.lblNumDevicesText.ImageAlign = _
            System.Drawing.ContentAlignment.MiddleLeft
        Me.lblNumDevicesText.Location = New System.Drawing.Point(20, 18)
        Me.lblNumDevicesText.Name = "lblNumDevicesText"
        Me.lblNumDevicesText.Size = New System.Drawing.Size(152, 32)
        Me.lblNumDevicesText.TabIndex = 24
        Me.lblNumDevicesText.Text = "Number of attached devices:"
        Me.lblNumDevicesText.TextAlign = _
            System.Drawing.ContentAlignment.MiddleCenter
        '
        'btLogClear
        '
        Me.btLogClear.Location = New System.Drawing.Point(684, 394)
        Me.btLogClear.Name = "btLogClear"
        Me.btLogClear.Size = New System.Drawing.Size(88, 40)
        Me.btLogClear.TabIndex = 32
        Me.btLogClear.Text = "&Clear Log"
        '
        'btActiveAltSetChange
        '
        Me.btActiveAltSetChange.Location = New System.Drawing.Point(780, 298)
        Me.btActiveAltSetChange.Name = "btActiveAltSetChange"
        Me.btActiveAltSetChange.Size = New System.Drawing.Size(88, 48)
        Me.btActiveAltSetChange.TabIndex = 31
        Me.btActiveAltSetChange.Text = "Change active &alternate setting"
        '
        'lbLog
        '
        Me.lbLog.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, _
            System.Drawing.FontStyle.Bold, _
            System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbLog.Location = New System.Drawing.Point(20, 378)
        Me.lbLog.Name = "lbLog"
        Me.lbLog.Size = New System.Drawing.Size(72, 16)
        Me.lbLog.TabIndex = 33
        Me.lbLog.Text = "Log"
        '
        'txtLog
        '
        Me.txtLog.AutoSize = False
        Me.txtLog.Location = New System.Drawing.Point(20, 394)
        Me.txtLog.Multiline = True
        Me.txtLog.Name = "txtLog"
        Me.txtLog.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtLog.Size = New System.Drawing.Size(648, 192)
        Me.txtLog.TabIndex = 34
        Me.txtLog.Text = ""
        '
        'tabDevices
        '
        Me.tabDevices.ItemSize = New System.Drawing.Size(350, 18)
        Me.tabDevices.Location = New System.Drawing.Point(20, 66)
        Me.tabDevices.Name = "tabDevices"
        Me.tabDevices.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.tabDevices.SelectedIndex = 0
        Me.tabDevices.Size = New System.Drawing.Size(738, 280)
        Me.tabDevices.SizeMode = System.Windows.Forms.TabSizeMode.Fixed
        Me.tabDevices.TabIndex = 26
        '
        'btPipeWrite
        '
        Me.btPipeWrite.Location = New System.Drawing.Point(780, 138)
        Me.btPipeWrite.Name = "btPipeWrite"
        Me.btPipeWrite.Size = New System.Drawing.Size(88, 40)
        Me.btPipeWrite.TabIndex = 28
        Me.btPipeWrite.Text = "&Write To Pipe"
        '
        'btPipeRead
        '
        Me.btPipeRead.Location = New System.Drawing.Point(780, 90)
        Me.btPipeRead.Name = "btPipeRead"
        Me.btPipeRead.Size = New System.Drawing.Size(88, 40)
        Me.btPipeRead.TabIndex = 27
        Me.btPipeRead.Text = "&Read From Pipe"
        '
        'btPipeListen
        '
        Me.btPipeListen.Location = New System.Drawing.Point(780, 186)
        Me.btPipeListen.Name = "btPipeListen"
        Me.btPipeListen.Size = New System.Drawing.Size(88, 48)
        Me.btPipeListen.TabIndex = 29
        Me.btPipeListen.Text = "&Listen To Pipe"
        '
        'btPipeReset
        '
        Me.btPipeReset.Location = New System.Drawing.Point(780, 242)
        Me.btPipeReset.Name = "btPipeReset"
        Me.btPipeReset.Size = New System.Drawing.Size(88, 48)
        Me.btPipeReset.TabIndex = 30
        Me.btPipeReset.Text = "Re&set Pipe"
        '
        'UsbSample
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(888, 605)
        Me.Controls.Add(Me.btExit)
        Me.Controls.Add(Me.lblNumDevices)
        Me.Controls.Add(Me.lblNumDevicesText)
        Me.Controls.Add(Me.btLogClear)
        Me.Controls.Add(Me.btActiveAltSetChange)
        Me.Controls.Add(Me.lbLog)
        Me.Controls.Add(Me.txtLog)
        Me.Controls.Add(Me.tabDevices)
        Me.Controls.Add(Me.btPipeWrite)
        Me.Controls.Add(Me.btPipeRead)
        Me.Controls.Add(Me.btPipeListen)
        Me.Controls.Add(Me.btPipeReset)
        Me.Name = "UsbSample"
        Me.Text = "USB .NET Sample"
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub UsbSample_Load(ByVal sender As System.Object, ByVal e As _
        System.EventArgs) Handles MyBase.Load

        Dim dDeviceAttachCb As D_USER_ATTACH_CALLBACK = AddressOf _
            UserDeviceAttach
        Dim dDeviceDetachCb As D_USER_DETACH_CALLBACK = AddressOf _
            UserDeviceDetach
        uDevManager = New UsbDeviceManager(dDeviceAttachCb, dDeviceDetachCb, _
            Convert.ToUInt16(DEFAULT_VENDOR_ID), _
            Convert.ToUInt16(DEFAULT_PRODUCT_ID), DEFAULT_DRIVER_NAME, _
            DEFAULT_LICENSE_STRING)

        UpdateButtons()

    End Sub

    Private Sub UserDeviceAttach(ByVal pDev As UsbDevice)
        Dim AttachCb As D_ATTACH_GUI_CALLBACK = AddressOf DeviceAttachGuiCb
        Invoke(AttachCb, New Object() {pDev})
    End Sub

    Private Sub DeviceAttachGuiCb(ByVal pDev As UsbDevice)
        lblNumDevices.Text = uDevManager.GetNumOfDevicesAttached().ToString()
        Dim tabPage As New DeviceTabPage(Me, pDev)
        tabDevices.Controls.Add(tabPage)
    End Sub

    Private Sub UserDeviceDetach(ByVal pDev As UsbDevice)
        Dim DetachCb As D_DETACH_GUI_CALLBACK = AddressOf DeviceDetachGuiCb
        Invoke(DetachCb, New Object() {pDev})
    End Sub

    Private Sub DeviceDetachGuiCb(ByVal pDev As UsbDevice)
        Dim numTabs As Int32 = tabDevices.TabCount
        Dim i As Int32 = 0
        Dim temp As DeviceTabPage

        lblNumDevices.Text = uDevManager.GetNumOfDevicesAttached().ToString()
        Do
            temp = CType(tabDevices.TabPages(i), DeviceTabPage)
            i = i + 1
        Loop Until (temp.GetUsbDev() Is pDev)
        tabDevices.Controls.Remove(temp)
    End Sub

    Private Sub UsbSample_Closing(ByVal eventSender As System.Object, _
        ByVal eventArgs As System.EventArgs) Handles MyBase.Closed

        Dim lblClose As New Label
        lblClose.Text = "Please wait untill " & APP_NAME & "." & _
            Environment.NewLine & "shuts down..."
        lblClose.Font = New System.Drawing.Font(lblClose.Font.FontFamily, 18, _
           FontStyle.Bold)
        lblClose.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        lblClose.Size = New System.Drawing.Size(400, 200)
        lblClose.BorderStyle = System.Windows.Forms.BorderStyle.None
        lblClose.Location = New System.Drawing.Point(200, 100)

        DisableButtons()
        tabDevices.Visible = False
        Me.Controls.Add(lblClose)
        Refresh()
        uDevManager.Dispose()
    End Sub

    Public Function GetActiveTab() As DeviceTabPage
        If (tabDevices.TabCount = 0) Then
            Return Nothing
        End If

        Return CType(tabDevices.GetControl(tabDevices.SelectedIndex), _
            DeviceTabPage)
    End Function

    Private Sub tabDevices_SelectedIndexChanged(ByVal sender As System.Object, _
        ByVal e As System.EventArgs) Handles tabDevices.SelectedIndexChanged

        UpdateButtons()
    End Sub

    Private Sub DisableButtons()
        btPipeWrite.Enabled = False
        btPipeRead.Enabled = False
        btPipeListen.Enabled = False
        btPipeReset.Enabled = False
        btActiveAltSetChange.Enabled = False
    End Sub

    Public Sub UpdateButtons()
        If (tabDevices.InvokeRequired()) Then
            Dim cb As New safeButtonsAccessCallBack(AddressOf UpdateButtons)
            Me.Invoke(cb)

        Else
            If (tabDevices.TabCount = 0) Then
                DisableButtons()
                Return
            End If
            Dim activePage As DeviceTabPage = GetActiveTab()
            If (activePage Is Nothing) Then
                Return
            End If
            Dim pipe As UsbPipe = activePage.GetActivePipe()

            btPipeWrite_Update(pipe)
            btPipeRead_Update(pipe)
            btPipeListen_Update(pipe)
            btPipeReset_Update(pipe)
            btActiveAltSetChange_Update(activePage)
        End If
    End Sub

    Private Sub btPipeWrite_Update(ByVal Pipe As UsbPipe)
        btPipeWrite.Enabled = Not (Pipe.IsPipeDirectionIn())

        If (Pipe.IsInUse() AndAlso btPipeWrite.Enabled) Then
            btPipeWrite.Text = "Stop &Writing"
        Else
            btPipeWrite.Text = "&Write To Pipe"
        End If

    End Sub

    Private Sub btPipeRead_Update(ByVal Pipe As UsbPipe)
        btPipeRead.Enabled = Not (Pipe.IsPipeDirectionOut()) AndAlso _
            Not (Pipe.IsContiguous())

        If (Pipe.IsInUse() AndAlso btPipeRead.Enabled) Then
            btPipeRead.Text = "Stop &Reading"
        Else
            btPipeRead.Text = "&Read From Pipe"
        End If

    End Sub

    Private Sub btPipeListen_Update(ByVal Pipe As UsbPipe)
        btPipeListen.Enabled = Not (Pipe.IsPipeDirectionOut() OrElse _
           Pipe.IsControlPipe()) AndAlso Not (Pipe.IsInUse() AndAlso _
           Not (Pipe.IsContiguous()))

        If (Pipe.IsInUse() AndAlso btPipeListen.Enabled) Then
            btPipeListen.Text = "Stop &Listening"
        Else
            btPipeListen.Text = "&Listen To Pipe"
        End If

    End Sub

    Private Sub btPipeReset_Update(ByVal Pipe As UsbPipe)
        btPipeReset.Enabled = Not (Pipe.IsControlPipe()) AndAlso Not _
        (Pipe.IsInUse())
    End Sub

    Public Sub btActiveAltSetChange_Update(ByVal activePage As DeviceTabPage)
        Dim dwNumOfAltSettings As UInt32 = _
            activePage.GetUsbDev().GetNumOfAlternateSettingsTotal()
        btActiveAltSetChange.Enabled = ((Convert.ToInt32(dwNumOfAltSettings) _
            > 1) AndAlso Not (activePage.GetUsbDev().IsDeviceTransferring()))
    End Sub

    Private Sub TransferCompletion(ByVal pipe As UsbPipe)
        UpdateButtons()

        Dim transferType As String
        If (pipe.GetfRead()) Then
            transferType = "read"
        Else
            transferType = "written"
        End If

        If (Convert.ToInt64(pipe.GetTransferStatus()) = _
            wdu_err.WD_STATUS_SUCCESS) Then
            TraceMsg(String.Format("Transfer completed successfully! " & _
                "Data {0}: {1} ", transferType, _
                DisplayHexBuffer(pipe.GetBuffer(), pipe.GetBytesTransferred())))
        Else
            Dim dwStatus As UInt32 = pipe.GetTransferStatus()
            ErrMsg(String.Format("Transfer Failed! Error {0}: {1} ", _
                dwStatus.ToString("X"), _
                utils.Stat2Str(dwStatus)))
        End If

    End Sub

    Private Sub ListenCompletion(ByVal pipe As UsbPipe)
        Dim dwStatus As UInt32 = pipe.GetTransferStatus()
        Dim IsListenStopped As Boolean = ((Convert.ToInt64(dwStatus) _
            = wdu_err.WD_IRP_CANCELED) AndAlso (Not pipe.IsContiguous()))
        If (Convert.ToInt64(dwStatus) <> wdu_err.WD_STATUS_SUCCESS AndAlso _
            Not IsListenStopped) Then

            ErrMsg(String.Format("Transfer Failed! Error {0}: {1} ", _
                dwStatus.ToString("X"), utils.Stat2Str(dwStatus)))
            UpdateButtons()
        Else
            TraceMsg(String.Format("{0}", DisplayHexBuffer(pipe.GetBuffer(), _
                pipe.GetBytesTransferred())))
        End If
    End Sub

    Private Sub SingleTransfer(ByVal fRead As Boolean)
        Dim activeTab As DeviceTabPage = GetActiveTab()
        If (activeTab Is Nothing) Then
            Return
        End If

        Dim currUsbDev As UsbDevice = activeTab.GetUsbDev()
        Dim activePipe As UsbPipe = activeTab.GetActivePipe()
        Dim dwPipeNum As UInt32 = activeTab.GetActivePipeNum()
        Dim fControl As Boolean = activePipe.IsControlPipe()
        Dim dwBuffSize As UInt32 = Convert.ToUInt32(0)
        Dim buffer() As Byte = Nothing
        Dim pSetupPacket(7) As Byte
        Dim dwOptions As UInt32 = Convert.ToUInt32(0)
        Dim userTransCompletion As D_USER_TRANSFER_COMPLETION = AddressOf _
            TransferCompletion
    Dim dwTimeOut As UInt32 = Convert.ToUInt32(TIME_OUT)

        Dim frmTransfers As New FormTransfers(fRead, fControl)

        If (frmTransfers.GetInput(dwBuffSize, buffer, pSetupPacket) = False) _
            Then
            Return
        End If

        Dim transferType As String
        If (fRead) Then
            transferType = "reading from"
        Else
            transferType = "writing to"
        End If

        TraceMsg(String.Format("began {0} {1} pipe number 0x{2:X}", _
            transferType, activeTab.GetUsbDev().DeviceDescription(), dwPipeNum))

        If (activePipe.IsControlPipe()) Then
            Dim dwBytesTransferred As UInt32 = Convert.ToUInt32(0)
            activePipe.UsbPipeTransfer(fRead, dwOptions, buffer, dwBuffSize, _
                dwBytesTransferred, pSetupPacket, dwTimeOut)
            TransferCompletion(activePipe)
        Else
            activePipe.UsbPipeTransferAsync(fRead, dwOptions, buffer, _
                dwBuffSize, dwTimeOut, userTransCompletion)
        End If

    End Sub
    Private Function DisplayHexBuffer(ByVal buff() As Byte, ByVal dwBuffSize _
        As UInt32) As String

        Dim i As Int32
        Dim display As String = ""
        For i = 0 To (Convert.ToInt32(dwBuffSize) - 1)
            display = String.Concat(display, buff(i).ToString("X"), " ")
        Next i

        display = String.Concat(display, Environment.NewLine)
        Return display
    End Function

    Private Sub btPipeRead_Click(ByVal sender As System.Object, ByVal e As _
        System.EventArgs) Handles btPipeRead.Click
        Dim activePage As DeviceTabPage = GetActiveTab()
        If (activePage Is Nothing) Then
            Return
        End If
        Dim activePipe As UsbPipe = activePage.GetActivePipe()

        If (activePipe.IsInUse()) Then
            activePipe.HaltTransferOnPipe()
        Else
            SingleTransfer(True)
        End If

        UpdateButtons()
    End Sub

    Private Sub btPipeWrite_Click(ByVal sender As System.Object, ByVal e As _
        System.EventArgs) Handles btPipeWrite.Click

        Dim activePage As DeviceTabPage = GetActiveTab()
        If (activePage Is Nothing) Then
            Return
        End If

        Dim activePipe As UsbPipe = activePage.GetActivePipe()

        If (activePipe.IsInUse()) Then
            activePipe.HaltTransferOnPipe()
        Else
            SingleTransfer(False)
        End If

        UpdateButtons()
    End Sub

    Private Sub btPipeListen_Click(ByVal sender As System.Object, ByVal e As _
        System.EventArgs) Handles btPipeListen.Click

        Dim activePage As DeviceTabPage = GetActiveTab()
        If (activePage Is Nothing) Then
            Return
        End If

        Dim activePipe As UsbPipe = activePage.GetActivePipe()
        Dim userTransCompletion As D_USER_TRANSFER_COMPLETION = AddressOf _
            ListenCompletion

        If (activePipe.IsInUse()) Then
            activePipe.SetContiguous(False)
            activePipe.HaltTransferOnPipe()
        Else
            Dim dwOptions As UInt32 = Convert.ToUInt32(0)
            activePipe.SetContiguous(True)
            TraceMsg(String.Format("began listening to {0} pipe number 0x{1:X}", _
                activePage.GetUsbDev().DeviceDescription(), _
                activePipe.GetPipeNum()))
            activePipe.UsbPipeTransferAsync(True, dwOptions, _
                Convert.ToUInt32(TIME_OUT), userTransCompletion)

        End If
        UpdateButtons()

    End Sub

    Private Sub btPipeReset_Click(ByVal sender As System.Object, ByVal e As _
    System.EventArgs) Handles btPipeReset.Click
        Dim activePage As DeviceTabPage = GetActiveTab()
        If (activePage Is Nothing) Then
            Return
        End If

        Dim activePipe As UsbPipe = activePage.GetActivePipe()
        TraceMsg(String.Format("reseting {0} pipe number 0x{1:X}", _
            activePage.GetUsbDev().DeviceDescription(), _
            activePipe.GetPipeNum()))
        activePipe.ResetPipe()
    End Sub

    Private Sub btActiveAltSetChange_Click(ByVal sender As System.Object, _
        ByVal e As System.EventArgs) Handles btActiveAltSetChange.Click

        Dim activePage As DeviceTabPage = GetActiveTab()
        If (activePage Is Nothing) Then
            Return
        End If

        Dim currUsbDev As UsbDevice = activePage.GetUsbDev()
        Dim result As DialogResult
        Dim frmChangeSettings As New FormChangeSettings(currUsbDev)

        result = frmChangeSettings.ShowDialog()
        If (result = System.Windows.Forms.DialogResult.OK) Then
            Dim newInterface As UInt32 = frmChangeSettings.GetChosenInterface()
            Dim newSetting As UInt32 = frmChangeSettings.GetChosenSetting()
            Dim dwStatus As UInt32 = _
                currUsbDev.ChangeAlternateSetting(newInterface, newSetting)
            If (Convert.ToInt64(dwStatus) <> wdu_err.WD_STATUS_SUCCESS) Then
                ErrMsg(String.Format("Failed to change the device's " & _
                    "setting to the one chosen: {0}", _
                    utils.Stat2Str(dwStatus)))
            Else
                activePage.UpdateAlternateSettings()
            End If
        End If

    End Sub

    Private Sub btLogClear_Click(ByVal sender As System.Object, ByVal e As _
        System.EventArgs) Handles btLogClear.Click
        txtLog.Clear()
    End Sub

    Private Sub btExit_Click(ByVal sender As System.Object, ByVal e As _
        System.EventArgs) Handles btExit.Click
        Close()
        Dispose()
    End Sub

    Private Sub SafeLogText(ByVal sMsg as String)
        ' InvokeRequired required compares the thread ID of the
        ' calling thread to the thread ID of the creating thread.
        ' If these threads are different, it returns true.
        if txtLog.InvokeRequired Then
            Dim cb As New safeLogTextCallBack(AddressOf SafeLogText)
            Me.Invoke(cb, new object() { sMsg })
        Else
            txtLog.AppendText(sMsg)
        End If
    End Sub

    Public Sub TraceMsg(ByVal sMsg As String)
        SafeLogText(sMsg)
        SafeLogText(Environment.NewLine)
    End Sub

    Public Sub ErrMsg(ByVal sMsg As String)
        SafeLogText(sMsg)
        SafeLogText(Environment.NewLine)
    End Sub
End Class

