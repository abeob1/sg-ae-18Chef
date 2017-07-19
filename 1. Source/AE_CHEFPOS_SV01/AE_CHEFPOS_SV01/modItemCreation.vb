Module modItemCreation

    Private dtItemGroup As DataTable

    Public Function ProcessItemDetails(ByVal oDv As DataView, ByVal file As System.IO.FileInfo, ByRef sErrDesc As String) As Long
        Dim sFuncName As String = "ProcessItemDetails"
        Dim sSQL As String = String.Empty

        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ConnectToTargetCompany()", sFuncName)
            Console.WriteLine("Connecting Company")
            If ConnectToCompany(p_oCompany, p_oCompDef.sSAPDBName, p_oCompDef.sSAPUser, p_oCompDef.sSAPPwd, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

            If p_oCompany.Connected Then
                Console.WriteLine("Company connected to " & p_oCompany.CompanyDB)

                sSQL = "SELECT ""ItmsGrpCod"",UPPER(""ItmsGrpNam"") AS ""ItmsGrpNam"" FROM " & p_oCompDef.sSAPDBName & ".""OITB"" "
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("EXECUTING SQL :" & sSQL, sFuncName)
                dtItemGroup = ExecuteQueryReturnDataTable(sSQL, p_oCompDef.sSAPDBName)

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling StartTransaction()", sFuncName)
                If StartTransaction(sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

                If oDv.Count > 0 Then
                    Console.WriteLine("Creating Items")

                    Dim oItems As SAPbobsCOM.Items
                    oItems = p_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems)

                    For i As Integer = 0 To oDv.Count - 1
                        Dim sItemCode As String = String.Empty
                        Dim sItmGrpNam As String = String.Empty
                        Dim sItmGrpCod As String = String.Empty
                        sItemCode = oDv(i)(0).ToString.Trim()
                        sItmGrpNam = oDv(i)(2).ToString.Trim()

                        If oItems.GetByKey(sItemCode) Then
                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Item Code " & sItemCode & " already exists in SAP", sFuncName)
                        Else

                            dtItemGroup.DefaultView.RowFilter = "ItmsGrpNam = '" & sItmGrpNam.ToUpper() & "'"
                            If dtItemGroup.DefaultView.Count = 0 Then
                                sErrDesc = "Item Group name :: " & sItmGrpNam & " Not exists in SAP."
                                Console.WriteLine(sErrDesc)
                                Call WriteToLogFile(sErrDesc, sFuncName)
                                Throw New ArgumentException(sErrDesc)
                            Else
                                sItmGrpCod = dtItemGroup.DefaultView.Item(0)(0).ToString().Trim()
                            End If

                            oItems.ItemCode = sItemCode
                            oItems.ItemName = oDv(i)(1).ToString.Trim()
                            oItems.ItemsGroupCode = sItmGrpCod
                            oItems.Frozen = SAPbobsCOM.BoYesNoEnum.tYES
                            oItems.Valid = SAPbobsCOM.BoYesNoEnum.tNO

                            If oItems.Add() <> 0 Then
                                sErrDesc = "Error " & p_oCompany.GetLastErrorDescription
                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug(sErrDesc, sFuncName)
                                Console.WriteLine("Error while creating Item " & sItemCode)
                                Throw New ArgumentException(sErrDesc)
                            Else
                                Console.WriteLine("ItemCode " & sItemCode & " created successfully")
                                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("ItemCode " & sItemCode & " created successfully", sFuncName)
                            End If
                        End If

                    Next
                End If

            End If

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling CommitTransaction", sFuncName)
            If CommitTransaction(sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling FileMoveToArchive()", sFuncName)
            FileMoveToArchive(file, file.FullName, RTN_SUCCESS)

            'Insert Success Notificaiton into Table..
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling AddDataToTable()", sFuncName)
            AddDataToTable(p_oDtSuccess, file.Name, "Success")
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("File successfully uploaded" & file.FullName, sFuncName)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            ProcessItemDetails = RTN_SUCCESS
        Catch ex As Exception
            sErrDesc = ex.Message
            Call WriteToLogFile(sErrDesc, sFuncName)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling RollbackTransaction", sFuncName)
            If RollbackTransaction(sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

            'Insert Error Description into Table
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling AddDataToTable()", sFuncName)
            AddDataToTable(p_oDtError, file.Name, "Error", sErrDesc)
            'error condition

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with Error", sFuncName)
            ProcessItemDetails = RTN_ERROR
        End Try
    End Function

End Module
