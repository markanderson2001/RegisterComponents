Imports System
Imports System.IO
Imports System.Management
Imports System.Environment


Public Class FrmRegisterComponents

    Private Sub FrmRegisterComponents_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        lblos.Text = "Environment OS Version: " & Environment.OSVersion.ToString()
        lblmachine.Text = "Machinename: " & System.Environment.MachineName
        lblplat.Text = "Platform: " & System.Environment.OSVersion.Platform.ToString
        Dim i As Int32
        Dim bit64 As Boolean
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


        MsgBox("Components Successfully Registered", vbInformation)
        Application.DoEvents()
        Me.lblcomponents.Visible = False

    End Sub

    Private Sub RegisterComponent(ByVal sFileName As String, Optional ByVal bUnRegister As Boolean = False, Optional ByVal bHideResults As Boolean = True)
        '    If Len(Dir$(sFileName)) = 0 Then
        '        'File is missing
        '        MsgBox "Unable to locate file " & sFileName, vbCritical

        'chk if file exist
        'Dim filename As String
        'filename = "C:\windows\windows32\" & sFileName
        If File.Exists("C:\windows\system32\" & sFileName) Then


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
        Else
            'first try if 64 bit system in wow folder
            If File.Exists("C:\windows\SysWoW64\" & sFileName) Then
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
            Else
                'first try if 64 bit system in wow folder
                If File.Exists("C:\windows\Winnt\system32" & sFileName) Then
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
                End If

                'show in listbox file not found - manual needed
                ListBox1.Items.Add(sFileName)
            End If
        End If


    End Sub

  
    Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click
        Me.Dispose()
    End Sub
End Class
