Imports System.ComponentModel
Imports Microsoft.Office.Interop.Outlook
Imports Microsoft.Office.Interop
Imports System.Text.RegularExpressions

Public Class Email
    Dim ProcessoFinalizado As Boolean
    Dim Retorno As New List(Of String)()
    Dim Itens As Integer
    Dim CalculaTempo As Integer
    Dim Contador As Integer
    Dim calc As Integer = 0

    'Private Sub AbrirBTN_Click(sender As Object, e As EventArgs)

    '    Procurar.ShowDialog()
    '    ArquivoTXT.Text = Procurar.FileName

    'End Sub

    'Private Sub RunBTN_Click(sender As Object, e As EventArgs)

    '    If ArquivoTXT.Text = "" Then
    '        MsgBox("Selecione o arquivo a ser executado.", MsgBoxStyle.Critical, "Erro")
    '        Exit Sub
    '    Else
    '        RUN_pst()
    '    End If

    'End Sub

    'Private Sub RUN_pst()
    '    Dim file As String = IO.Path.GetFileNameWithoutExtension(Procurar.FileName)
    '    Try
    '        Dim mailItems As IEnumerable(Of MailItem) = readPst(ArquivoTXT.Text, file)
    '        For Each mailItem As MailItem In mailItems
    '            Console.WriteLine(mailItem.SenderName + " - " + mailItem.Subject)
    '            DataGridView1.Rows.Add(mailItem.SenderName + " - " + mailItem.Subject)
    '        Next
    '    Catch ex As System.Exception
    '        Console.WriteLine(ex.Message)
    '    End Try
    '    Console.ReadLine()
    'End Sub

    'Private Shared Function readPst(pstFilePath As String, pstName As String) As IEnumerable(Of MailItem)
    '    Dim mailItems As New List(Of MailItem)()
    '    Dim app As New Application()
    '    Dim outlookNs As [NameSpace] = app.GetNamespace("MAPI")
    '    ' Add PST file (Outlook Data File) to Default Profile
    '    outlookNs.AddStore(pstFilePath)
    '    Dim rootFolder As MAPIFolder = outlookNs.Stores(pstName).GetRootFolder()
    '    ' Traverse through all folders in the PST file
    '    ' TODO: This is not recursive, refactor
    '    Dim subFolders As Folders = rootFolder.Folders
    '    For Each folder As Folder In subFolders
    '        Dim items As Items = folder.Items
    '        For Each item As Object In items
    '            If TypeOf item Is MailItem Then
    '                Dim mailItem As MailItem = TryCast(item, MailItem)
    '                mailItems.Add(mailItem)
    '            End If
    '        Next
    '    Next
    '    ' Remove PST file from Default Profile
    '    outlookNs.RemoveStore(rootFolder)
    '    Return mailItems
    '    'End Function
    '    ****write this code for getting the mails ***************

    '    Dim objOL As Outlook.Application
    '    Dim objNS As Outlook.NameSpace
    '    Dim objFolder As Outlook.Folders
    '    Dim Item As Object
    '    Dim myItems As Outlook.Items
    '    Dim x As Int16

    'objOL = New Outlook.Application()
    'objNS = objOL.GetNamespace("MAPI")

    '    Dim olfolder As Outlook.MAPIFolder
    'olfolder = objOL.GetNamespace("MAPI").PickFolder
    'myItems = olfolder.Items

    '    Dim i As Integer
    'For x = 1 To myItems.Count
    '    messagebox.show(myItems.Item(x).SenderName) 
    '    messagebox.show myItems.Item(x).SenderEmailAddress)
    '    messagebox.showmyItems.Item(x).Subject)
    '    messagebox.showmyItems.Item(x).Body)
    '    messagebox.show myItems.Item(x).to)
    '    messagebox.showmyItems.Item(x).ReceivedByName)
    '    messagebox.show myItems.Item(x).ReceivedOnBehalfOfName)
    '    messagebox.showmyItems.Item(x).ReplyRecipientNames)
    '    messagebox.showmyItems.Item(x).SentOnBehalfOfName)
    '    messagebox.showmyItems.Item(x).CC)
    '    messagebox.show myItems.Item(x).ReceivedTime)
    'Next x

    '***********for getting attachments use this code******************

    '    Dim Atmt As Outlook.Attachment

    'For Each Atmt In Item.Attachments
    '    filename = "C:\Email Attachments\" + Atmt.FileName
    '    Atmt.SaveAsFile(filename)
    'Next Atmt

    Private Function ExtractEmailAddressesFromString(ByVal source As String) As String()
        Dim mc As MatchCollection
        Dim i As Integer

        ' expression garnered from www.regexlib.com - thanks guys!
        mc = Regex.Matches(source, _
            "([a-zA-Z0-9_\-\.]+)@([a-zA-Z0-9_\-\.]+)\.([a-zA-Z]{2,5})")
        Dim results(mc.Count - 1) As String
        For i = 0 To results.Length - 1
            results(i) = mc(i).Value
        Next

        Return results
    End Function

    Private Sub RUN_READ(ByVal processoAtivo As BackgroundWorker)
    
        Dim objOL As Application
        Dim objNS As Outlook.NameSpace
        Dim myItems As Items
        Dim x As Integer

        objOL = New Application()
        objNS = objOL.GetNamespace("MAPI")

        Dim olfolder As MAPIFolder
        olfolder = objOL.GetNamespace("MAPI").PickFolder
        myItems = olfolder.Items

        Itens = myItems.Count

        For x = 1 To myItems.Count

            On Error Resume Next

            If (processoAtivo.CancellationPending = True) Then Exit For

            Threading.Thread.Sleep(50)
            'processoAtivo.ReportProgress((x / myItems.Count) * 100)
            processoAtivo.ReportProgress(x)

            For Each I In ExtractEmailAddressesFromString(myItems.Item(x).SenderEmailAddress.ToString().Trim())
                'Emails.Rows.Add(I.ToString())
                Retorno.Add(I.ToString())
            Next

            For Each I In ExtractEmailAddressesFromString(myItems.Item(x).to.ToString().Trim())
                'Emails.Rows.Add(I.ToString())
                Retorno.Add(I.ToString())
            Next

            For Each I In ExtractEmailAddressesFromString(myItems.Item(x).CC.ToString().Trim())
                'Emails.Rows.Add(I.ToString())
                Retorno.Add(I.ToString())
            Next

            For Each I In ExtractEmailAddressesFromString(myItems.Item(x).Body.ToString().Trim())
                'Emails.Rows.Add(I.ToString())
                Retorno.Add(I.ToString())
            Next

        Next x

        ProcessoFinalizado = True
        
    End Sub

    Private Sub ExeBTN_Click(sender As Object, e As EventArgs) Handles ExeBTN.Click

        Contador = 0
        CalculaTempo = 0
        Timer1.Start()
        ToolStripProgressBar1.Value = 0
        ToolStripStatusLabel1.Text = "Extração . . . "
        GeraExtracao

    End Sub

    Private Sub ExportBTN_Click(sender As Object, e As EventArgs) Handles ExportBTN.Click

        Try
            If Emails.Rows.Count > 0 Then

                SaveFileDialog1.Filter = "Excel (*.csv) |*.csv"
                SaveFileDialog1.ShowDialog()

                Dim StrExport As String = ""

                For Each C As DataGridViewColumn In Emails.Columns
                    StrExport &= """" & C.HeaderText & ""","
                Next
                StrExport = StrExport.Substring(0, StrExport.Length - 1)
                StrExport &= Environment.NewLine

                For Each R As DataGridViewRow In Emails.Rows
                    For Each C As DataGridViewCell In R.Cells
                        If Not C.Value Is Nothing Then
                            StrExport &= """" & C.Value.ToString.Trim() & ""","
                        Else
                            StrExport &= """" & "" & ""","
                        End If
                    Next
                    StrExport = StrExport.Substring(0, StrExport.Length - 1)
                    StrExport &= Environment.NewLine
                Next

                Dim tw As IO.TextWriter = New IO.StreamWriter(SaveFileDialog1.FileName)
                tw.Write(StrExport)
                tw.Close()
                MsgBox("Operação Finalizada", , "Finalizado")
            End If

        Catch ex As System.Exception
            MsgBox(ex.Message)
        End Try
        
    End Sub

    Private Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker.DoWork

        Try
            Dim segundoPlano As BackgroundWorker
            segundoPlano = CType(sender, BackgroundWorker)
            RUN_READ(segundoPlano)
            If (BackgroundWorker.CancellationPending = True) Then e.Cancel = True

        Catch ex As System.Exception
            ProcessoFinalizado = False
            MessageBox.Show(ex.Message, "Alerta de Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End Try

    End Sub

    Private Sub GeraExtracao()

        Try
            ExeBTN.Enabled = False
            ExportBTN.Enabled = False
            Emails.Enabled = False
            CancBTN.Visible = True
            Me.BackgroundWorker.RunWorkerAsync()

        Catch ex As System.Exception
            ExeBTN.Enabled = True
            ExportBTN.Enabled = True
            Emails.Enabled = True
            CancBTN.Visible = False
            MessageBox.Show(ex.Message, "Alerta de Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End Try

    End Sub

    Private Sub BackgroundWorker_ProgressChanged(sender As Object, e As ProgressChangedEventArgs) Handles BackgroundWorker.ProgressChanged
        
        If CalculaTempo = 2 Then
            Contador += 1
            calc = (((CalculaTempo * Contador) * Itens) / e.ProgressPercentage) / 60
            CalculaTempo = 0
        End If

        ToolStripProgressBar1.Maximum = Itens
        ToolStripProgressBar1.Value = e.ProgressPercentage
        ToolStripStatusLabel2.Text = e.ProgressPercentage & " de " & Itens & " . . . Tempo Estmiado: " & calc & " min(s)"

    End Sub
    Private Sub BackgroundWorker_RunWorkerCompleted(sender As Object, e As RunWorkerCompletedEventArgs) Handles BackgroundWorker.RunWorkerCompleted


        For Each i In Retorno
            Dim valida As Boolean = True
            If Emails.RowCount > 0 Then
                For Each row As DataGridViewRow In Emails.Rows
                    If row.Cells(0).Value IsNot Nothing AndAlso row.Cells(0).Value.ToString.Trim() = i.ToString.Trim() Then
                        valida = False
                    End If
                Next
            End If
            If valida = True Then Emails.Rows.Add((i.ToString()))
        Next

        If (e.Cancelled = True) Then
            ToolStripStatusLabel1.Text = "Extração Cancelada!"
        Else
            If ProcessoFinalizado = True Then
                ToolStripStatusLabel1.Text = "Extração Finalizada!"

            Else
                ToolStripStatusLabel1.Text = "Extração com Erro!"
            End If
        End If

        ToolStripProgressBar1.Value = Itens
        ToolStripStatusLabel2.Text = String.Empty
        ExeBTN.Enabled = True
        Emails.Enabled = True
        CancBTN.Visible = False
        ExportBTN.Enabled = True

    End Sub

    Private Sub CancBTN_Click(sender As Object, e As EventArgs) Handles CancBTN.Click
        BackgroundWorker.CancelAsync()
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        
        CalculaTempo += 1

    End Sub

End Class
