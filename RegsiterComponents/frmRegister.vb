Imports System
Imports System.IO
Imports System.Management
Imports System.Environment



Public Class FrmRegisterComponents


#Region "  Class Members "
    Public bit64 As Boolean
    Public sys32 As Boolean
    Public winnt As Boolean
    Public finished As Boolean


#End Region

    Private Sub FrmRegisterComponents_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        lblos.Text = "Environment OS Version: " & Environment.OSVersion.ToString()
        lblmachine.Text = "Computer Name: " & System.Environment.MachineName
        lblplat.Text = "Platform: " & System.Environment.OSVersion.Platform.ToString
        Dim i As Int32

        i = IntPtr.Size
        Select Case i
            Case 4
                bit64 = False
                Label2.Text = "64 Bit System? " & bit64

            Case 8
                bit64 = True
                Label2.Text = "64 Bit System? " & bit64
            Case Else
                Label2.Text = "64 Bit System? Unknown"

        End Select
        If i = 4 Then

        Else
            bit64 = True

        End If



    End Sub

    Private Sub btnRegister_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRegister.Click
        'mark - these files are within vista/win7 but not registered, method created to unregister then register components
        finished = False
        RegisterComponent("msdatgrd.ocx", True, True)
        Application.DoEvents()
        RegisterComponent("msdatgrd.ocx", False, True)
        Application.DoEvents()
        lblcomponents.Visible = True
        lblcomponents.Text = "Registered - msdatgrd.ocx"

        RegisterComponent("msdbrptr.dll", True, True)
        Application.DoEvents()
        RegisterComponent("msdbrptr.dll", False, True)
        Application.DoEvents()
        lblcomponents.Text = "Registered - msdbrptr.dll"

        RegisterComponent("msflxgrd.ocx", True, True)
        Application.DoEvents()

        RegisterComponent("msflxgrd.ocx", False, True)
        Application.DoEvents()
        lblcomponents.Text = "Registered - msflxgrd.ocx"

        RegisterComponent("msmask32.ocx", True, True)
        Application.DoEvents()
        RegisterComponent("msmask32.ocx", False, True)
        Application.DoEvents()
        lblcomponents.Text = "Registered - msmask32.ocx"

        RegisterComponent("msstdfmt.dll", True, True)
        Application.DoEvents()
        RegisterComponent("msstdfmt.dll", False, True)
        Application.DoEvents()
        lblcomponents.Text = "Registered - msstdfmt.dll"



        RegisterComponent("msderun.dll", True, True)
        Application.DoEvents()
        RegisterComponent("msderun.dll", False, True)
        Application.DoEvents()
        lblcomponents.Text = "Registered - msderun.dll"

        RegisterComponent("xceedzip.dll", True, True)
        Application.DoEvents()
        RegisterComponent("xceedzip.dll", False, True)
        Application.DoEvents()
        lblcomponents.Text = "Registered - xceedzip.dll"

        RegisterComponent("xceedftp.dll", True, True)
        Application.DoEvents()
        RegisterComponent("xceedftp.dll", False, True)
        Application.DoEvents()
        lblcomponents.Text = "Registered - xceedftp.dll"

        RegisterComponent("mscomct2.ocx", True, True)
        Application.DoEvents()
        RegisterComponent("mscomct2.ocx", False, True)
        Application.DoEvents()
        lblcomponents.Text = "Registered - mscomct2.ocx"

        RegisterComponent("mscomctl.ocx", True, True)
        Application.DoEvents()
        RegisterComponent("mscomctl.ocx", False, True)
        Application.DoEvents()
        lblcomponents.Text = "Registered - mscomctl.ocx"


        RegisterComponent("comdlg32.ocx", True, True)
        Application.DoEvents()
        RegisterComponent("comdlg32.ocx", False, True)
        Application.DoEvents()
        lblcomponents.Text = "Registered - comdlg32.ocx"

        RegisterComponent("Msdatrep.ocx", True, True)
        Application.DoEvents()
        RegisterComponent("Msdatrep.ocx", False, True)
        Application.DoEvents()
        lblcomponents.Text = "Registered - Msdatrep.ocx"

        RegisterComponent("Msbind.dll", True, True)
        Application.DoEvents()
        RegisterComponent("Msbind.dll", False, True)
        Application.DoEvents()
        lblcomponents.Text = "Registered - Msbind.dll"

        RegisterComponent("sqldmo.dll", True, True)
        Application.DoEvents()
        RegisterComponent("sqldmo.dll", False, True)
        Application.DoEvents()
        lblcomponents.Text = "Registered - sqldmo.dll"


        RegisterComponent("msflxgrd.ocx", True, True)
        Application.DoEvents()
        RegisterComponent("msflxgrd.ocx", False, True)
        Application.DoEvents()
        lblcomponents.Text = "Registered - sqldmo.dll"

        finished = True
        MsgBox("Components Successfully Registered", vbInformation)
        Application.DoEvents()
        Me.lblcomponents.Visible = False

    End Sub
    Function ReportFolderStatus(ByVal fldr) As Boolean 'chk if folder exist

        Dim fso, msg
        fso = CreateObject("Scripting.FileSystemObject")
        If (fso.FolderExists(fldr)) Then
            msg = fldr & " exists."
            ReportFolderStatus = True

        Else
            msg = fldr & " doesn't exist."
            ReportFolderStatus = False

        End If
        'ReportFolderStatus = msg
    End Function

    Private Sub RegisterComponent(ByVal sFileName As String, Optional ByVal bUnRegister As Boolean = False, Optional ByVal bHideResults As Boolean = True)
        '    If Len(Dir$(sFileName)) = 0 Then
        '        'File is missing
        '        MsgBox "Unable to locate file " & sFileName, vbCritical

        'chk if file exist
        'Dim filename As String
        'filename = "C:\windows\windows32\" & sFileName


        Try
            If finished = True Then Exit Sub

        
        If bit64 = True Then
                'folder exist?
                sFileName = "C:\Windows\SysWow64\" & sFileName


                If ReportFolderStatus("C:\Windows\SysWow64") = True Then
                'Folder found lets find files & register 
tryagain:

                    If File.Exists(sFileName) Then

                        If bUnRegister Then
                            'Unregister a component
                            If bHideResults Then
                                'Hide results
                                Shell("regsvr32 /s /u " & """" & sFileName & """")
                            Else
                                'Show results
                                Shell("regsvr32 /u " & """" & sFileName & """")
                            End If

                        Else
                            'Register a component
                            If bHideResults Then
                                'Hide results
                                Shell("regsvr32 /s " & """" & sFileName & """")
                            Else
                                'Show results
                                Shell("regsvr32 " & """" & sFileName & """")
                            End If
                        End If
                    Else 'folder not found
                        'if xceed stuff copy from c:\eds\treasury\support
                        If sFileName = "C:\Windows\SysWow64\xceedzip.dll" Then
                            Dim FSO
                            FSO = CreateObject("Scripting.FileSystemObject")
                            FSO.CopyFile("c:\eds\treasury\support\xceedzip.dll", "C:\Windows\SysWow64\")
                            GoTo tryagain
                        End If
                        If sFileName = "C:\Windows\SysWow64\xceedftp.dll" Then
                            Dim FSO
                            FSO = CreateObject("Scripting.FileSystemObject")
                            FSO.CopyFile("c:\eds\treasury\support\xceedftp.dll", "C:\Windows\SysWow64\")
                            GoTo tryagain
                        End If

                        'show in listbox file not found - manual needed
                        ListBox1.Items.Add(sFileName)
                    End If
            Else
                    '  MsgBox("Default folder: ""SysWow64"" was not found on this 64Bit System, Please register components Manually", MsgBoxStyle.Information)
                    Exit Sub
            End If

                Exit Sub

            End If

            '========================================================================================================end wow

            If ReportFolderStatus("C:\Windows\system32") = True Then
                sFileName = "C:\Windows\system32\" & sFileName
tryagain1:
                If File.Exists(sFileName) Then

                    If bUnRegister Then
                        'Unregister a component
                        If bHideResults Then
                            'Hide results
                            Shell("regsvr32 /s /u " & """" & sFileName & """")
                        Else
                            'Show results
                            Shell("regsvr32 /u " & """" & sFileName & """")
                        End If
                    Else
                        'Register a component
                        If bHideResults Then
                            'Hide results
                            Shell("regsvr32 /s " & """" & sFileName & """")
                        Else
                            'Show results
                            Shell("regsvr32 " & """" & sFileName & """")
                        End If
                    End If

                Else 'file not found
                    'if xceed stuff copy from c:\eds\treasury\support
                    If sFileName = "C:\Windows\system32\xceedzip.dll" Then
                        Dim FSO0
                        FSO0 = CreateObject("Scripting.FileSystemObject")
                        FSO0.CopyFile("c:\eds\treasury\support\xceedzip.dll", "C:\Windows\system32\")
                        GoTo tryagain1
                    End If
                    If sFileName = "C:\Windows\system32\xceedftp.dll" Then
                        Dim FSO0
                        FSO0 = CreateObject("Scripting.FileSystemObject")
                        FSO0.CopyFile("c:\eds\treasury\support\xceedftp.dll", "C:\Windows\system32\")
                        GoTo tryagain1
                    End If

                    'show in listbox file not found - manual needed
                    ListBox1.Items.Add(sFileName)
                    ' MsgBox("Default folder: ""SysWow64"" was not found on this 64Bit System, Please register components Manually", MsgBoxStyle.Information)

                    Exit Sub
                End If
            End If






            If ReportFolderStatus("C:\Windows\Winnt\system32") = True Then
                sFileName = "C:\Windows\Winnt\system32\" & sFileName
tryagain2:
                If File.Exists(sFileName) Then

                    If bUnRegister Then
                        'Unregister a component
                        If bHideResults Then
                            'Hide results
                            Shell("regsvr32 /s /u " & """" & sFileName & """")
                        Else
                            'Show results
                            Shell("regsvr32 /u " & """" & sFileName & """")
                        End If
                    Else
                        'Register a component
                        If bHideResults Then
                            'Hide results
                            Shell("regsvr32 /s " & """" & sFileName & """")
                        Else
                            'Show results
                            Shell("regsvr32 " & """" & sFileName & """")
                        End If
                    End If


                Else 'folder not found

                    'if xceed stuff copy from c:\eds\treasury\support
                    If sFileName = "C:\Windows\Winnt\system32\xceedzip.dll" Then
                        Dim FSO1
                        FSO1 = CreateObject("Scripting.FileSystemObject")
                        FSO1.CopyFile("c:\eds\treasury\support\xceedzip.dll", "C:\Windows\Winnt\system32\")
                        GoTo tryagain2
                    End If
                    If sFileName = "C:\Windows\Winnt\system32\xceedftp.dll" Then
                        Dim FSO1
                        FSO1 = CreateObject("Scripting.FileSystemObject")
                        FSO1.CopyFile("c:\eds\treasury\support\xceedftp.dll", "C:\Windows\Winnt\system32\")
                        GoTo tryagain2
                    End If

                    'show in listbox file not found - manual needed
                    ListBox1.Items.Add(sFileName)
                    ' MsgBox("Default folder: ""SysWow64"" was not found on this 64Bit System, Please register components Manually", MsgBoxStyle.Information)

                    Exit Sub
                End If
            End If


        Catch ex As Exception
            MsgBox(Err.Description & Err.Number & "Error Registration Process")
            Exit Sub
        End Try

    End Sub


    Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click
        Me.Dispose()
    End Sub
End Class
