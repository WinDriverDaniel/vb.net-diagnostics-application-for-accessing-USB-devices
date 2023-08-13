' Note: This code sample is provided AS-IS and as a guiding sample only.

Option Explicit On
Imports System
Imports System.Drawing
Imports System.Collections
Imports System.ComponentModel
Imports System.Windows.Forms
Imports Jungo.wdapi_dotnet
Imports Jungo.usb_lib

Public Class DeviceTabPage
    Inherits System.Windows.Forms.TabPage

    Private usbDev As UsbDevice
    Private pipesListView As System.Windows.Forms.ListView
    Private iSelectedPipeIndex As Int32 = -1
    Private ParentForm As UsbSample

    Public Sub New(ByRef Parent As UsbSample, ByRef pUsbDev As UsbDevice)

        ParentForm = Parent
        pipesListView = New System.Windows.Forms.ListView
        usbDev = pUsbDev

        RightToLeft = System.Windows.Forms.RightToLeft.No
        Text = usbDev.DeviceDescription()
        CreateListViewSettings()
        UpdatePipesListView()
        Controls.Add(pipesListView)
        AddHandler pipesListView.SelectedIndexChanged, AddressOf _
            PipeListSelectedIndexChanged

    End Sub

    Private Sub CreateListViewSettings()

        pipesListView.Location = New Point(8, 8)
        pipesListView.Name = "pipesListView"
        pipesListView.HideSelection = False
        pipesListView.Size = New _
            System.Drawing.Size(ParentForm.tabDevices.Width - 16, _
            ParentForm.tabDevices.Height - 16)
        pipesListView.View = View.Details
        pipesListView.MultiSelect = False
        pipesListView.GridLines = False
        pipesListView.FullRowSelect = True
        pipesListView.LabelEdit = False
        pipesListView.HeaderStyle = ColumnHeaderStyle.Nonclickable
        pipesListView.Columns.Add("Pipe", _
            CType(pipesListView.Width / 4 - 1, Integer), _
            HorizontalAlignment.Left)
        pipesListView.Columns.Add("Type", _
            CType(pipesListView.Width / 4 - 1, Integer), _
            HorizontalAlignment.Left)
        pipesListView.Columns.Add("Direction", _
            CType(pipesListView.Width / 4 - 1, Integer), _
            HorizontalAlignment.Left)
        pipesListView.Columns.Add("Max Packet Size", _
            CType(pipesListView.Width / 4 - 1, Integer), _
            HorizontalAlignment.Left)

    End Sub


    Public Sub UpdatePipesListView()

        Dim index As Int32 = 0
        Dim currUsbPipe As UsbPipe

        pipesListView.BeginUpdate()
        pipesListView.Items.Clear()

        For Each currUsbPipe In usbDev.GetpPipesList()
            Dim pipeListItem As New ListViewItem("0x" & _
                currUsbPipe.GetPipeNum().ToString("X"))
            pipeListItem.SubItems.Add(PipeTypeToString(currUsbPipe))
            pipeListItem.SubItems.Add(PipeDirectionToString(currUsbPipe))
            pipeListItem.SubItems.Add(String.Concat("0x", _
                (currUsbPipe.GetPipeMaxPacketSz()).ToString("X")))
            pipesListView.Items.Add(pipeListItem)
            index = index + 1
        Next

        pipesListView.Items(0).Selected = True
        pipesListView.EndUpdate()

    End Sub

    Sub PipeListSelectedIndexChanged(ByVal sender As System.Object, _
        ByVal e As System.EventArgs) 'Handles pipesListView.SelectedIndexChanged

        If (pipesListView.SelectedItems.Count > 0) Then
            Dim Index As Int32 = pipesListView.SelectedItems(0).Index
            If (iSelectedPipeIndex = -1 OrElse iSelectedPipeIndex <> Index) Then
                iSelectedPipeIndex = Index
                If (ParentForm.GetActiveTab() Is Me) Then
                    ParentForm.UpdateButtons()
                End If
            End If
        End If

    End Sub

    Public Sub UpdateAlternateSettings()

        UpdatePipesListView()
        Text = usbDev.DeviceDescription()

    End Sub

    Public Function GetActivePipe() As UsbPipe
        Return CType((usbDev.GetpPipesList())(iSelectedPipeIndex), UsbPipe)
    End Function

    Public Function GetActivePipeNum() As UInt32
        Return (CType((usbDev.GetpPipesList())(iSelectedPipeIndex), _
            UsbPipe)).GetPipeNum()
    End Function

    Public Function GetUsbDev() As UsbDevice
        Return usbDev
    End Function

    Public Sub SetUsbDev(ByRef newUsbDev As UsbDevice)
        usbDev = newUsbDev
    End Sub

    Private Function PipeTypeToString(ByVal pipe As UsbPipe) As String
        Dim strPipeType As String

        If (pipe.IsControlPipe()) Then
            strPipeType = "Control"
        ElseIf (pipe.IsBulkPipe()) Then
            strPipeType = "Bulk"

        ElseIf (pipe.IsInterruptPipe()) Then
            strPipeType = "Interrupt"

        ElseIf (pipe.IsIsochronousPipe()) Then
            strPipeType = "Isochronous"

        Else
            strPipeType = "N/A"
        End If
        Return strPipeType

    End Function

    Private Function PipeDirectionToString(ByVal pipe As UsbPipe) As String
        Dim strPipeDirection As String

        If (pipe.IsPipeDirectionIn()) Then
            strPipeDirection = "In"

        ElseIf (pipe.IsPipeDirectionOut()) Then
            strPipeDirection = "Out"

        ElseIf (pipe.IsPipeDirectionInOut()) Then
            strPipeDirection = "In/Out"

        Else
            strPipeDirection = "N/A"
        End If
        Return strPipeDirection

    End Function

End Class

