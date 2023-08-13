' Note: This code sample is provided AS-IS and as a guiding sample only.

Imports System
Imports System.Drawing
Imports System.Collections
Imports System.ComponentModel
Imports System.Windows.Forms
Imports Jungo.wdapi_dotnet
Imports Jungo.usb_lib


Public Class FormChangeSettings
    Inherits System.Windows.Forms.Form

    Private settingsArr(,) As Int32
    Private i32ChosenInterface As Int32
    Private i32ChosenSetting As Int32
#Region " Windows Form Designer generated code "

    Public Sub New(ByRef usbDev As UsbDevice)
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        Dim numOfAltSettings As Int32 = _
            Convert.ToInt32(usbDev.GetNumOfAlternateSettingsTotal())
        Dim currInterfaceIndex As Int32 = _
            Convert.ToInt32(usbDev.GetCurrInterfaceIndex())
        Dim currAltSetting As Int32 = _
            Convert.ToInt32(usbDev.GetCurrAlternateSettingNum())
        ReDim settingsArr(numOfAltSettings, 1)

        Dim currIndex As Int32 = 0
        Dim i As Int32 = 0
        Dim interfac As Int32
        Dim altSetting As Int32

        For interfac = 0 To (Convert.ToInt32(usbDev.GetNumOfInteraces) - 1)
            Dim interfaceNumber As Int32 = _
                usbDev.GetInterfaceNumberByIndex(Convert.ToUInt32(interfac))

            For altSetting = 0 To (Convert.ToInt32(usbDev.GetNumOfAlternateSettingsPerInterface(Convert.ToUInt32(interfaceNumber))) - 1)
                cmboAltSettings.Items.Add("Interface " & _
                    interfaceNumber.ToString() & ", Alternate Setting " & _
                    altSetting.ToString())

                settingsArr(i, 0) = interfaceNumber
                settingsArr(i, 1) = altSetting
                i = i + 1

                If interfac <= currInterfaceIndex And altSetting < currAltSetting Then
                    currIndex = currIndex + 1
                End If

            Next altSetting

        Next interfac

        cmboAltSettings.SelectedIndex = currIndex
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
    Friend WithEvents btCancel As System.Windows.Forms.Button
    Friend WithEvents btSubmit As System.Windows.Forms.Button
    Friend WithEvents cmboAltSettings As System.Windows.Forms.ComboBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.btCancel = New System.Windows.Forms.Button
        Me.btSubmit = New System.Windows.Forms.Button
        Me.cmboAltSettings = New System.Windows.Forms.ComboBox
        Me.SuspendLayout()
        '
        'btCancel
        '
        Me.btCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btCancel.Font = New System.Drawing.Font("Microsoft Sans Serif", _
            8.25!, System.Drawing.FontStyle.Bold, _
            System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btCancel.Location = New System.Drawing.Point(120, 88)
        Me.btCancel.Name = "btCancel"
        Me.btCancel.Size = New System.Drawing.Size(64, 24)
        Me.btCancel.TabIndex = 22
        Me.btCancel.Text = "Cancel"
        '
        'btSubmit
        '
        Me.btSubmit.Font = New System.Drawing.Font("Microsoft Sans Serif", _
            8.25!, System.Drawing.FontStyle.Bold, _
            System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btSubmit.Location = New System.Drawing.Point(32, 88)
        Me.btSubmit.Name = "btSubmit"
        Me.btSubmit.Size = New System.Drawing.Size(64, 24)
        Me.btSubmit.TabIndex = 21
        Me.btSubmit.Text = "Submit"
        '
        'cmboAltSettings
        '
        Me.cmboAltSettings.ItemHeight = 13
        Me.cmboAltSettings.Location = New System.Drawing.Point(32, 32)
        Me.cmboAltSettings.Name = "cmboAltSettings"
        Me.cmboAltSettings.Size = New System.Drawing.Size(152, 21)
        Me.cmboAltSettings.TabIndex = 24
        Me.cmboAltSettings.Text = "Choose Setting"
        '
        'FormChangeSettings
        '
        Me.AcceptButton = Me.btSubmit
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.CancelButton = Me.btCancel
        Me.ClientSize = New System.Drawing.Size(232, 149)
        Me.Controls.Add(Me.cmboAltSettings)
        Me.Controls.Add(Me.btCancel)
        Me.Controls.Add(Me.btSubmit)
        Me.Name = "FormChangeSettings"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Change device's Settings"
        Me.ResumeLayout(False)
    End Sub

#End Region

    Public Function GetChosenInterface() As UInt32
        Return Convert.ToUInt32(i32ChosenInterface)
    End Function

    Public Function GetChosenSetting() As UInt32
        Return Convert.ToUInt32(i32ChosenSetting)
    End Function

    Private Sub btSubmit_Click(ByVal sender As System.Object, _
        ByVal e As System.EventArgs) Handles btSubmit.Click

        Dim index As Int32 = cmboAltSettings.SelectedIndex
        i32ChosenInterface = settingsArr(index, 0)
        i32ChosenSetting = settingsArr(index, 1)
        DialogResult = System.Windows.Forms.DialogResult.OK

    End Sub
End Class

