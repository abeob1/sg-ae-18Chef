Module modProcess

    Public Sub Start()
        Dim sFuncName As String = "Start()"
        Dim sErrDesc As String = String.Empty

        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("calling ReadExcel()", sFuncName)

            Console.WriteLine("Reading text values")

            UploadFiles(sErrDesc)

            'Send Error Email if Datable has rows.
            If p_oDtError.Rows.Count > 0 Then
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling EmailTemplate_Error()", sFuncName)
                EmailTemplate_Error()
            End If
            p_oDtError.Rows.Clear()

            'Send Success Email if Datable has rows..
            If p_oDtSuccess.Rows.Count > 0 Then
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling EmailTemplate_Success()", sFuncName)
                EmailTemplate_Success()
            End If
            p_oDtSuccess.Rows.Clear()


            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
        Catch ex As Exception
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            End
        End Try
    End Sub

    Private Function UploadFiles(ByRef sErrDesc As String) As Long

        Dim sFuncName As String = "UploadFiles"
        Dim oDVHeaderData As DataView = New DataView
        Dim oDVDetailsData As DataView = New DataView
        Dim oDVItemData As DataView = New DataView
        Dim oDVCollections As DataView = New DataView

        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Upload funciton", sFuncName)

            p_oDtSuccess = CreateDataTable("FileName", "Status")
            p_oDtError = CreateDataTable("FileName", "Status", "ErrDesc")

            Dim DirInfo As New System.IO.DirectoryInfo(p_oCompDef.sInboxDir)
            Dim files() As System.IO.FileInfo

            files = DirInfo.GetFiles("NewItem_*.txt")

            For Each file As System.IO.FileInfo In files
                sErrDesc = String.Empty

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("File Name is: " & file.Name.ToUpper, sFuncName)
                Console.WriteLine("Reading File: " & file.Name.ToUpper)

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Read header text file into Dataview", sFuncName)
                oDVItemData = ReadTextFile(file.FullName, "ITEM")

                If Not oDVDetailsData Is Nothing Then
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ProcessItemDetails()", sFuncName)
                    Console.WriteLine("Processing file " & file.Name)
                    If ProcessItemDetails(oDVItemData, file, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                Else
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("No Data's found in text file :" & file.Name & ". Please check the datas in header and detail file", sFuncName)
                    Continue For
                End If
            Next

            files = DirInfo.GetFiles("SALESHDR_*.txt")

            For Each file As System.IO.FileInfo In files
                sErrDesc = String.Empty

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("File Name is: " & file.Name.ToUpper, sFuncName)
                Console.WriteLine("Reading File: " & file.Name.ToUpper)

                Dim sFileDate As String = String.Empty
                Dim sFileName As String = String.Empty
                sFileName = file.FullName

                Dim k As Integer = file.Name.IndexOf("_")
                sFileDate = file.Name.Substring(k, Len(file.Name) - k)

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Read header text file into Dataview", sFuncName)
                oDVHeaderData = ReadTextFile(file.FullName, "")

                If Not oDVHeaderData Is Nothing Then
                    Dim files_detail() As System.IO.FileInfo
                    Dim sDetailFile As String = "SALESHDET" & sFileDate
                    files_detail = DirInfo.GetFiles(sDetailFile)

                    For Each DetailFile As System.IO.FileInfo In files_detail

                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("File Name is: " & DetailFile.Name.ToUpper, sFuncName)
                        Console.WriteLine("Reading File: " & DetailFile.Name.ToUpper)

                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Read detail text file into Dataview", sFuncName)
                        oDVDetailsData = ReadTextFile(DetailFile.FullName, "")

                        If Not oDVDetailsData Is Nothing Then
                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ProcessInvoiceFiles()", sFuncName)
                            Console.WriteLine("Processing files " & file.Name & " and " & DetailFile.Name)
                            If ProcessInvoiceFiles(file, DetailFile, oDVHeaderData, oDVDetailsData, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                        Else
                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("No Data's found in text file :" & file.Name & ". Please check the datas in header and detail file", sFuncName)
                            Continue For
                        End If
                    Next
                End If

            Next

            files = DirInfo.GetFiles("PAYMENT_*.txt")

            For Each file As System.IO.FileInfo In files
                sErrDesc = String.Empty

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("File Name is: " & file.Name.ToUpper, sFuncName)
                Console.WriteLine("Reading File: " & file.Name.ToUpper)

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Read header text file into Dataview", sFuncName)
                oDVItemData = ReadTextFile(file.FullName, "")

                If Not oDVDetailsData Is Nothing Then
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ProcessCollectionDetails()", sFuncName)
                    Console.WriteLine("Processing file " & file.Name)
                    If ProcessCollectionDetails(oDVItemData, file, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                Else
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("No Data's found in text file :" & file.Name & ". Please check the datas in header and detail file", sFuncName)
                    Continue For
                End If
            Next

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Function Completed Successfully.", sFuncName)
            UploadFiles = RTN_SUCCESS

        Catch ex As Exception
            UploadFiles = RTN_ERROR
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Error in Uplodiang AR file.", sFuncName)
            Call WriteToLogFile(ex.Message, sFuncName)
        End Try
    End Function

End Module
