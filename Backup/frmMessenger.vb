
Public Class frmMsnGers
    'this code is by no means bullet proof and may be used a giudeline
    'on how to perform such a task.   Gogs[05-21-2004].
    Inherits System.Windows.Forms.Form
    Private Const cnstLocalInifile = "c:\users.ini"
    Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

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

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents lblStatus As System.Windows.Forms.StatusBar
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents txtMessage As System.Windows.Forms.TextBox
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents cmdRemove As System.Windows.Forms.Button
    Friend WithEvents cmdAdd As System.Windows.Forms.Button
    Friend WithEvents txtManual As System.Windows.Forms.TextBox
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents cmdSend As System.Windows.Forms.Button
    Friend WithEvents cmdClear As System.Windows.Forms.Button
    Friend WithEvents cmdSave As System.Windows.Forms.Button
    Friend WithEvents GroupBox4 As System.Windows.Forms.GroupBox
    Friend WithEvents lstnames As System.Windows.Forms.ListBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmMsnGers))
        Me.lblStatus = New System.Windows.Forms.StatusBar
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.txtMessage = New System.Windows.Forms.TextBox
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.cmdRemove = New System.Windows.Forms.Button
        Me.cmdAdd = New System.Windows.Forms.Button
        Me.txtManual = New System.Windows.Forms.TextBox
        Me.GroupBox3 = New System.Windows.Forms.GroupBox
        Me.cmdSend = New System.Windows.Forms.Button
        Me.cmdClear = New System.Windows.Forms.Button
        Me.cmdSave = New System.Windows.Forms.Button
        Me.GroupBox4 = New System.Windows.Forms.GroupBox
        Me.lstnames = New System.Windows.Forms.ListBox
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.GroupBox4.SuspendLayout()
        Me.SuspendLayout()
        '
        'lblStatus
        '
        Me.lblStatus.Location = New System.Drawing.Point(0, 339)
        Me.lblStatus.Name = "lblStatus"
        Me.lblStatus.Size = New System.Drawing.Size(516, 22)
        Me.lblStatus.TabIndex = 8
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.txtMessage)
        Me.GroupBox1.Location = New System.Drawing.Point(176, 0)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(326, 296)
        Me.GroupBox1.TabIndex = 9
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "2.Enter Message Text to Send"
        '
        'txtMessage
        '
        Me.txtMessage.Location = New System.Drawing.Point(8, 20)
        Me.txtMessage.Multiline = True
        Me.txtMessage.Name = "txtMessage"
        Me.txtMessage.Size = New System.Drawing.Size(310, 267)
        Me.txtMessage.TabIndex = 8
        Me.txtMessage.Text = ""
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.cmdRemove)
        Me.GroupBox2.Controls.Add(Me.cmdAdd)
        Me.GroupBox2.Controls.Add(Me.txtManual)
        Me.GroupBox2.Location = New System.Drawing.Point(8, 296)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(216, 45)
        Me.GroupBox2.TabIndex = 10
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Add Names[Computer/User] to List"
        '
        'cmdRemove
        '
        Me.cmdRemove.Location = New System.Drawing.Point(192, 16)
        Me.cmdRemove.Name = "cmdRemove"
        Me.cmdRemove.Size = New System.Drawing.Size(17, 16)
        Me.cmdRemove.TabIndex = 18
        Me.cmdRemove.Text = "-"
        '
        'cmdAdd
        '
        Me.cmdAdd.Location = New System.Drawing.Point(168, 16)
        Me.cmdAdd.Name = "cmdAdd"
        Me.cmdAdd.Size = New System.Drawing.Size(17, 16)
        Me.cmdAdd.TabIndex = 17
        Me.cmdAdd.Text = "+"
        '
        'txtManual
        '
        Me.txtManual.Location = New System.Drawing.Point(8, 16)
        Me.txtManual.Name = "txtManual"
        Me.txtManual.Size = New System.Drawing.Size(150, 20)
        Me.txtManual.TabIndex = 16
        Me.txtManual.Text = ""
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.cmdSend)
        Me.GroupBox3.Controls.Add(Me.cmdClear)
        Me.GroupBox3.Controls.Add(Me.cmdSave)
        Me.GroupBox3.Location = New System.Drawing.Point(224, 296)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(278, 46)
        Me.GroupBox3.TabIndex = 11
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "3. Actions"
        '
        'cmdSend
        '
        Me.cmdSend.Location = New System.Drawing.Point(208, 15)
        Me.cmdSend.Name = "cmdSend"
        Me.cmdSend.Size = New System.Drawing.Size(61, 21)
        Me.cmdSend.TabIndex = 15
        Me.cmdSend.Text = "Send"
        '
        'cmdClear
        '
        Me.cmdClear.Location = New System.Drawing.Point(96, 16)
        Me.cmdClear.Name = "cmdClear"
        Me.cmdClear.Size = New System.Drawing.Size(105, 21)
        Me.cmdClear.TabIndex = 14
        Me.cmdClear.Text = "Clear Message"
        '
        'cmdSave
        '
        Me.cmdSave.Location = New System.Drawing.Point(8, 16)
        Me.cmdSave.Name = "cmdSave"
        Me.cmdSave.Size = New System.Drawing.Size(81, 21)
        Me.cmdSave.TabIndex = 13
        Me.cmdSave.Text = "Save List"
        '
        'GroupBox4
        '
        Me.GroupBox4.Controls.Add(Me.lstnames)
        Me.GroupBox4.Location = New System.Drawing.Point(8, 0)
        Me.GroupBox4.Name = "GroupBox4"
        Me.GroupBox4.Size = New System.Drawing.Size(163, 294)
        Me.GroupBox4.TabIndex = 12
        Me.GroupBox4.TabStop = False
        Me.GroupBox4.Text = "1. Select Recipients"
        '
        'lstnames
        '
        Me.lstnames.Cursor = System.Windows.Forms.Cursors.Hand
        Me.lstnames.HorizontalScrollbar = True
        Me.lstnames.Location = New System.Drawing.Point(16, 16)
        Me.lstnames.Name = "lstnames"
        Me.lstnames.SelectionMode = System.Windows.Forms.SelectionMode.MultiSimple
        Me.lstnames.Size = New System.Drawing.Size(137, 264)
        Me.lstnames.Sorted = True
        Me.lstnames.TabIndex = 10
        '
        'frmMsnGers
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(516, 361)
        Me.Controls.Add(Me.GroupBox4)
        Me.Controls.Add(Me.GroupBox3)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.lblStatus)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.Name = "frmMsnGers"
        Me.Text = "Simple Send"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox4.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub cmdAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAdd.Click
        Try
            '---------------------------------------------------------------
            ' add new name to list 
            '---------------------------------------------------------------
            If Len(txtManual.Text) > 0 Then
                lstnames.Items.Add(txtManual.Text)
            End If
        Catch
            'make sure user knows something unexpected happened
            MsgBox(Err.Description, vbOKOnly, Windows.Forms.Application.ProductName)
        End Try
    End Sub

    Private Sub cmdRemove_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdRemove.Click
        Try
            '---------------------------------------------------------------
            ' Loop through list removing selected items
            '---------------------------------------------------------------
            Dim aryListSelected As String
            Dim arySelected() As String
            Dim intNames As Integer
            Const mySeperator = ":--:"

            If lstnames.SelectedItems.Count > 0 Then
                For Each strSeleted As String In lstnames.SelectedItems
                    If Len(aryListSelected) > 0 Then ' if removing more than one insert seperator
                        aryListSelected &= mySeperator
                    End If
                    aryListSelected &= lstnames.SelectedIndex.ToString 'append selected position
                Next strSeleted
                arySelected = Split(aryListSelected, mySeperator) 'make array of positions
            End If

            'loop treough array and take out based on original position
            For intLoop As Integer = 1 To intNames
                lstnames.Items.RemoveAt(Val(arySelected(intLoop - 1)))
            Next intLoop

        Catch ex As Exception
            'make sure user knows something unexpected happened
            Call MessageBox.Show(ex.Message, Windows.Forms.Application.ProductName, MessageBoxButtons.OK)
        End Try
    End Sub

    Private Sub cmdSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSave.Click
        '---------------------------------------------------------------
        ' save list to file on c drive 
        ' if this file is deleted file in app directory is used
        '---------------------------------------------------------------
        Dim liFilenum As Byte
        Dim tmpString As String
        Try
            liFilenum = FreeFile() 'declaraion out of try to remain in scope in finally
            FileOpen(liFilenum, cnstLocalInifile, OpenMode.Output, OpenAccess.Write, OpenShare.LockWrite)
            For intLoop As Integer = 1 To lstnames.Items.Count
                tmpString = lstnames.Items.Item(intLoop - 1) & vbCrLf
                Print(liFilenum, tmpString)
            Next intLoop
        Catch ex As Exception
            'make sure user knows something unexpected happened
            Call MessageBox.Show(ex.Message, Windows.Forms.Application.ProductName, MessageBoxButtons.OK)
        Finally
            If liFilenum > 0 Then FileClose(liFilenum) 'good idea to make sure file is closed
        End Try
    End Sub

    Private Sub cmdSend_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSend.Click

        '-------------------------------------------------------------------------------
        ' send message using net send 
        ' the messenger service needs to be running on both sender/recipient machines
        '-------------------------------------------------------------------------------
        Dim strStatus As String
        Const iChunk = 512

        Try

            Dim strMessage As String = txtMessage.Text

            'list of likely error messages for user to see
            Select Case True
                Case Len(strMessage) = 0
                    strStatus = "No message to send!"
                    lblStatus.Text = strStatus
                    Exit Sub
                Case lstnames.Items.Count = 0
                    strStatus = "No Recipients Available"
                    lblStatus.Text = strStatus
                    Exit Sub
                Case lstnames.SelectedItems.Count = 0
                    strStatus = "No Recipients Selected"
                    lblStatus.Text = strStatus
                    Exit Sub
                Case lstnames.SelectedItems.Count > 0
                    'tests passed initialise front end message
                    strStatus = "Message sent :"
                Case Else
                    strStatus = "Something Screwy!!!"
                    lblStatus.Text = strStatus
                    Exit Sub
            End Select


            'chop up big messages
            If lstnames.SelectedItems.Count > 0 Then
                For Each strSeleted As String In lstnames.SelectedItems 'somehow a list item is in fact a string
                    Dim strname As String = " " & strSeleted
                    Dim lngPos As Long = 1
                    Do
                        Dim lretval As Long = Shell("net send " & strname & " """ & Mid(strMessage, lngPos, iChunk) & """", vbHide)
                        lngPos += iChunk
                        'pause to stop out of sequence message chunks
                        Sleep(100)
                    Loop Until lngPos > Len(strMessage)
                    strStatus &= strname.ToString
                Next strSeleted
            End If

            strStatus &= " @ " & TimeOfDay.ToString ' append timestamp

            lblStatus.Text = strStatus

        Catch ex As Exception
            lblStatus.Text = Err.Description 'pass handled error to user nicely
        Finally
        End Try
    End Sub

    Private Sub frmMsnGers_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        '-------------------------------------------------------------------------------
        ' load app and read initial user list
        '-------------------------------------------------------------------------------

        Dim liFilenum As Byte
        Dim inifilename As String

        Try
            Dim fileLocalPath As New System.IO.FileInfo("c:\users.ini")
            If fileLocalPath.Exists Then
                inifilename = "c:\users.ini"
            Else
                inifilename = Application.StartupPath & IIf(Application.StartupPath.EndsWith("\"), "", "\") & "users.ini"
            End If

            Dim fileUsingPath As New System.IO.FileInfo(inifilename)
            If fileUsingPath.Exists Then
                liFilenum = FreeFile() 'get next available freefile number

                'open the file open for reading
                FileOpen(liFilenum, inifilename, OpenMode.Input, OpenAccess.Read, OpenShare.Default)

                Do While Not EOF(liFilenum)   ' Loop until end of file.
                    lstnames.Items.Add(LineInput(liFilenum).ToString) 'Read line into list.
                Loop
            End If
        Catch fex As System.IO.FileNotFoundException
            'not a problem in no user list file found
        Catch ex As Exception
            Call MessageBox.Show(ex.Message, Windows.Forms.Application.ProductName, MessageBoxButtons.OK)
        Finally
            If liFilenum > 0 Then FileClose(liFilenum) 'good idea to make sure file is closed
        End Try
    End Sub

    Private Sub txtMessage_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtMessage.TextChanged


    End Sub
End Class