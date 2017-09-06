Module modInvoice

    Private dtBPMaster As DataTable
    Private dtCheckType As DataTable
    Private dtVatGroup As DataTable
    Private dtCostCenter As DataTable

    Public Function ProcessInvoiceFiles(ByVal file_Header As System.IO.FileInfo, ByVal file_Detail As System.IO.FileInfo, ByVal file_Payment As System.IO.FileInfo, ByVal oDvHeader As DataView, ByVal oDvDeatil As DataView, ByVal oDvCollections As DataView, ByRef sErrDesc As String) As Long
        Dim sFuncName As String = "ProcessInvoiceFiles"
        Dim sSQL As String = String.Empty

        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ConnectToTargetCompany()", sFuncName)
            Console.WriteLine("Connecting Company")
            If ConnectToCompany(p_oCompany, p_oCompDef.sSAPDBName, p_oCompDef.sSAPUser, p_oCompDef.sSAPPwd, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

            If p_oCompany.Connected Then
                Console.WriteLine("Company connected to " & p_oCompany.CompanyDB)

                sSQL = "SELECT DISTINCT ""CardCode"",UPPER(""CardCode"") AS ""UPPERCARDCODE"" FROM " & p_oCompDef.sSAPDBName & ".""OCRD"" "
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("EXECUTING SQL :" & sSQL, sFuncName)
                dtBPMaster = ExecuteQueryReturnDataTable(sSQL, p_oCompDef.sSAPDBName)

                sSQL = "SELECT UPPER(""Code"") AS ""CHECKTYPE"",""Name"" AS ""U_SAPACCOUNT"" FROM " & p_oCompDef.sSAPDBName & ".""@CHECKTYPE"" "
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("EXECUTING SQL :" & sSQL, sFuncName)
                dtCheckType = ExecuteQueryReturnDataTable(sSQL, p_oCompDef.sSAPDBName)

                sSQL = "SELECT ""ItemCode"",""VatGourpSa"" FROM " & p_oCompany.CompanyDB & ".""OITM"" WHERE ""frozenFor""='N'"
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("EXECUTING  SQL :" & sSQL, sFuncName)
                dtVatGroup = ExecuteQueryReturnDataTable(sSQL, p_oCompany.CompanyDB)

                sSQL = "SELECT ""PrcCode"",UPPER(""PrcName"") AS ""PrcName"",""DimCode"" FROM ""OPRC"" "
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("EXECUTING  SQL :" & sSQL, sFuncName)
                dtCostCenter = ExecuteQueryReturnDataTable(sSQL, p_oCompany.CompanyDB)

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling StartTransaction", sFuncName)
                If StartTransaction(sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Grouping header data based on fileid", sFuncName)
                Dim oDtHeader_Group As DataTable = oDvHeader.Table.DefaultView.ToTable(True, "FILEID")
                For i As Integer = 0 To oDtHeader_Group.Rows.Count - 1
                    If Not (oDtHeader_Group.Rows(i).Item(0).ToString.Trim() = String.Empty Or oDtHeader_Group.Rows(i).Item(0).ToString.ToUpper().Trim() = "FILEID") Then
                        oDvHeader.RowFilter = "FILEID ='" & oDtHeader_Group.Rows(i).Item(0).ToString.Trim() & "' "

                        If oDvHeader.Count > 0 Then
                            Dim sInvDocEntry As String = String.Empty
                            Dim oDtHeader_Grouped As DataTable
                            oDtHeader_Grouped = oDvHeader.ToTable
                            Dim oDvHeader_Grouped As DataView = New DataView(oDtHeader_Grouped)

                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling CreateInvoice()", sFuncName)
                            Console.WriteLine("Processing Invoice for file id " & oDtHeader_Group.Rows(i).Item(0).ToString.Trim())
                            If CreateInvoice(oDvHeader_Grouped, oDvDeatil, sInvDocEntry, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

                            If sInvDocEntry <> "" Then
                                oDvCollections.RowFilter = "FILEID ='" & oDtHeader_Group.Rows(i).Item(0).ToString.Trim() & "' "

                                If oDvCollections.Count > 0 Then
                                    Dim oDtCollections_Grouped As DataTable
                                    oDtCollections_Grouped = oDvCollections.ToTable
                                    Dim oDvCollections_Grouped As DataView = New DataView(oDtCollections_Grouped)

                                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling CreatePayment()", sFuncName)
                                    If CreatePayment(oDvCollections_Grouped, sInvDocEntry, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                                End If
                            End If


                        End If

                    End If
                Next

                'processing payment
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Grouping Payment data based on fileid", sFuncName)
                Dim oDtPay_Group As DataTable = oDvCollections.Table.DefaultView.ToTable(True, "FILEID")
                For i As Integer = 0 To oDtPay_Group.Rows.Count - 1
                    If Not (oDtPay_Group.Rows(i).Item(0).ToString.Trim() = String.Empty Or oDtPay_Group.Rows(i).Item(0).ToString.ToUpper().Trim() = "FILEID") Then

                    End If
                Next
            End If

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling CommitTransaction", sFuncName)
            If CommitTransaction(sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling FileMoveToArchive()", sFuncName)
            FileMoveToArchive(file_Header, file_Header.FullName, RTN_SUCCESS)
            FileMoveToArchive(file_Detail, file_Detail.FullName, RTN_SUCCESS)
            FileMoveToArchive(file_Payment, file_Payment.FullName, RTN_SUCCESS)

            'Insert Success Notificaiton into Table..
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling AddDataToTable()", sFuncName)
            AddDataToTable(p_oDtSuccess, file_Header.Name, "Success")
            AddDataToTable(p_oDtSuccess, file_Detail.Name, "Success")
            AddDataToTable(p_oDtSuccess, file_Payment.Name, "Success")

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Header File " & file_Header.FullName & " & Detail file " & file_Detail.FullName & " & Payment file " & file_Payment.FullName & " uploaded successfully", sFuncName)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            ProcessInvoiceFiles = RTN_SUCCESS
        Catch ex As Exception
            sErrDesc = ex.Message
            Call WriteToLogFile(sErrDesc, sFuncName)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling RollbackTransaction", sFuncName)
            If RollbackTransaction(sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

            'Insert Error Description into Table
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling AddDataToTable()", sFuncName)
            AddDataToTable(p_oDtError, file_Header.Name, "Error", sErrDesc)
            AddDataToTable(p_oDtError, file_Detail.Name, "Error", sErrDesc)
            AddDataToTable(p_oDtError, file_Payment.Name, "Error", sErrDesc)
            'error condition

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling FileMoveToArchive()", sFuncName)
            FileMoveToArchive(file_Header, file_Header.FullName, RTN_ERROR)
            FileMoveToArchive(file_Detail, file_Detail.FullName, RTN_ERROR)
            FileMoveToArchive(file_Payment, file_Payment.FullName, RTN_ERROR)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            ProcessInvoiceFiles = RTN_ERROR
        End Try
    End Function

    Public Function ProcessInvoiceFiles_WithoutPayment(ByVal file_Header As System.IO.FileInfo, ByVal file_Detail As System.IO.FileInfo, ByVal oDvHeader As DataView, ByVal oDvDeatil As DataView, ByRef sErrDesc As String) As Long
        Dim sFuncName As String = "ProcessInvoiceFiles_WithoutPayment"
        Dim sSQL As String = String.Empty

        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ConnectToTargetCompany()", sFuncName)
            Console.WriteLine("Connecting Company")
            If ConnectToCompany(p_oCompany, p_oCompDef.sSAPDBName, p_oCompDef.sSAPUser, p_oCompDef.sSAPPwd, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

            If p_oCompany.Connected Then
                Console.WriteLine("Company connected to " & p_oCompany.CompanyDB)

                sSQL = "SELECT DISTINCT ""CardCode"",UPPER(""CardCode"") AS ""UPPERCARDCODE"" FROM " & p_oCompDef.sSAPDBName & ".""OCRD"" "
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("EXECUTING SQL :" & sSQL, sFuncName)
                dtBPMaster = ExecuteQueryReturnDataTable(sSQL, p_oCompDef.sSAPDBName)

                sSQL = "SELECT UPPER(""U_CHECKTYPE"") AS ""CHECKTYPE"",U_SAPACCOUNT FROM " & p_oCompDef.sSAPDBName & ".""@AE_CHECKTYPE"" "
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("EXECUTING SQL :" & sSQL, sFuncName)
                dtCheckType = ExecuteQueryReturnDataTable(sSQL, p_oCompDef.sSAPDBName)

                sSQL = "SELECT ""ItemCode"",""VatGourpSa"" FROM " & p_oCompany.CompanyDB & ".""OITM"" WHERE ""frozenFor""='N'"
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("EXECUTING  SQL :" & sSQL, sFuncName)
                dtVatGroup = ExecuteQueryReturnDataTable(sSQL, p_oCompany.CompanyDB)

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling StartTransaction", sFuncName)
                If StartTransaction(sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Grouping header data based on fileid", sFuncName)
                Dim oDtHeader_Group As DataTable = oDvHeader.Table.DefaultView.ToTable(True, "FILEID")
                For i As Integer = 0 To oDtHeader_Group.Rows.Count - 1
                    If Not (oDtHeader_Group.Rows(i).Item(0).ToString.Trim() = String.Empty Or oDtHeader_Group.Rows(i).Item(0).ToString.ToUpper().Trim() = "FILEID") Then
                        oDvHeader.RowFilter = "FILEID ='" & oDtHeader_Group.Rows(i).Item(0).ToString.Trim() & "' "

                        If oDvHeader.Count > 0 Then
                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling CreateInvoice()", sFuncName)
                            Dim oDtHeader_Grouped As DataTable
                            oDtHeader_Grouped = oDvHeader.ToTable
                            Dim oDvHeader_Grouped As DataView = New DataView(oDtHeader_Grouped)

                            If CreateInvoice(oDvHeader_Grouped, oDvDeatil, "", sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                        End If

                    End If
                Next
            End If

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling CommitTransaction", sFuncName)
            If CommitTransaction(sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling FileMoveToArchive()", sFuncName)
            FileMoveToArchive(file_Header, file_Header.FullName, RTN_SUCCESS)
            FileMoveToArchive(file_Detail, file_Detail.FullName, RTN_SUCCESS)

            'Insert Success Notificaiton into Table..
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling AddDataToTable()", sFuncName)
            AddDataToTable(p_oDtSuccess, file_Header.Name, "Success")
            AddDataToTable(p_oDtSuccess, file_Detail.Name, "Success")

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Header File " & file_Header.FullName & " & Detail file " & file_Detail.FullName & " uploaded successfully", sFuncName)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            ProcessInvoiceFiles_WithoutPayment = RTN_SUCCESS
        Catch ex As Exception
            sErrDesc = ex.Message
            Call WriteToLogFile(sErrDesc, sFuncName)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling RollbackTransaction", sFuncName)
            If RollbackTransaction(sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

            'Insert Error Description into Table
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling AddDataToTable()", sFuncName)
            AddDataToTable(p_oDtError, file_Header.Name, "Error", sErrDesc)
            AddDataToTable(p_oDtError, file_Detail.Name, "Error", sErrDesc)
            'error condition

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling FileMoveToArchive()", sFuncName)
            FileMoveToArchive(file_Header, file_Header.FullName, RTN_ERROR)
            FileMoveToArchive(file_Detail, file_Detail.FullName, RTN_ERROR)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            ProcessInvoiceFiles_WithoutPayment = RTN_ERROR
        End Try
    End Function

    Private Function CreateInvoice(ByVal oDvHeader As DataView, ByVal oDvDetail As DataView, ByRef sInvDocEntry As String, ByRef sErrDesc As String) As Long
        Dim sFuncName As String = "CreateInvoice"
        Dim sFileId As String = String.Empty
        Dim sCardCode As String = String.Empty
        Dim sConcept As String = String.Empty
        Dim sBrand As String = String.Empty
        Dim sOutLet As String = String.Empty
        Dim sOcrCode As String = String.Empty
        Dim sOcrCode2 As String = String.Empty
        Dim sOcrCode3 As String = String.Empty
        Dim bIsLineAdded As Boolean = False

        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            sFileId = oDvHeader(0)(0).ToString().Trim()

            oDvDetail.RowFilter = Nothing
            oDvDetail.RowFilter = "FILEID = '" & sFileId & "'"
            Dim odt As New DataTable
            odt = oDvDetail.ToTable
            Dim oDvDetail_Grouped As DataView = New DataView(odt)

            Dim oInvoice As SAPbobsCOM.Documents
            oInvoice = p_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices)

            sCardCode = "C" & oDvHeader(0)(4).ToString.Trim()
            dtBPMaster.DefaultView.RowFilter = "UPPERCARDCODE = '" & sCardCode.ToUpper() & "'"
            If dtBPMaster.DefaultView.Count = 0 Then
                sErrDesc = "CardCode :: " & sCardCode & " Not exists in SAP."
                Console.WriteLine(sErrDesc)
                Call WriteToLogFile(sErrDesc, sFuncName)
                Throw New ArgumentException(sErrDesc)
            Else
                sCardCode = dtBPMaster.DefaultView.Item(0)(0).ToString().Trim()
            End If

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("CardCode is " & sCardCode, sFuncName)

            Dim iIndex As Integer = oDvHeader(0)(5).ToString.IndexOf(" ")
            Dim sDate As String
            If iIndex > -1 Then
                sDate = oDvHeader(0)(5).ToString.Substring(0, iIndex)
            Else
                sDate = oDvHeader(0)(5).ToString
            End If
            Dim dtDocDate As Date
            Dim format() = {"dd/MM/yyyy", "d/M/yyyy", "dd-MM-yyyy", "dd.MM.yyyy", "yyyyMMdd", "MMddYYYY", "M/dd/yyyy", "MM/dd/YYYY"}
            Date.TryParseExact(sDate, format, System.Globalization.DateTimeFormatInfo.InvariantInfo, Globalization.DateTimeStyles.None, dtDocDate)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Assigning header values", sFuncName)

            sConcept = oDvHeader(0)(2).ToString.Trim()
            sBrand = oDvHeader(0)(3).ToString.Trim()
            sOutLet = oDvHeader(0)(4).ToString.Trim()

            dtCostCenter.DefaultView.RowFilter = "PrcName = '" & sBrand.ToUpper() & "' AND DimCode = 1 "
            If dtCostCenter.DefaultView.Count = 0 Then
                sOcrCode = String.Empty
            Else
                sOcrCode = dtCostCenter.DefaultView.Item(0)(0).ToString().Trim()
            End If

            dtCostCenter.DefaultView.RowFilter = "PrcName = '" & sConcept.ToUpper() & "' AND DimCode = 2 "
            If dtCostCenter.DefaultView.Count = 0 Then
                sOcrCode2 = String.Empty
            Else
                sOcrCode2 = dtCostCenter.DefaultView.Item(0)(0).ToString().Trim()
            End If

            dtCostCenter.DefaultView.RowFilter = "PrcName = '" & sOutLet.ToUpper() & "' AND DimCode = 3 "
            If dtCostCenter.DefaultView.Count = 0 Then
                sOcrCode3 = String.Empty
            Else
                sOcrCode3 = dtCostCenter.DefaultView.Item(0)(0).ToString().Trim()
            End If

            oInvoice.CardCode = sCardCode
            oInvoice.DocDate = dtDocDate
            oInvoice.NumAtCard = oDvHeader(0)(4).ToString.Trim() & "-" & sDate & "-" & oDvHeader(0)(7).ToString.Trim()
            oInvoice.UserFields.Fields.Item("U_Concept").Value = sConcept
            oInvoice.UserFields.Fields.Item("U_BRAND").Value = sBrand
            oInvoice.UserFields.Fields.Item("U_Outlet").Value = sOutLet
            oInvoice.UserFields.Fields.Item("U_POSNo").Value = oDvHeader(0)(4).ToString.Trim()
            oInvoice.UserFields.Fields.Item("U_MEALPERIOD").Value = oDvHeader(0)(6).ToString.Trim()
            oInvoice.UserFields.Fields.Item("U_HOUR").Value = oDvHeader(0)(7).ToString.Trim()
            oInvoice.UserFields.Fields.Item("U_Covers").Value = oDvHeader(0)(14).ToString.Trim()
            oInvoice.UserFields.Fields.Item("U_NetTables").Value = oDvHeader(0)(15).ToString.Trim()
            oInvoice.DocTotal = CDbl(oDvHeader(0)(8))

            Dim iCount As Integer = 0
            If oDvDetail_Grouped.Count > 0 Then
                For i As Integer = 0 To oDvDetail_Grouped.Count - 1
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Processing detail line " & i + 1, sFuncName)

                    Dim sItemCode As String = String.Empty
                    Dim sDeliveryMode As String = String.Empty
                    Dim sAccountCode As String = String.Empty
                    Dim sVatGroup As String = String.Empty
                    Dim iAdjustment As Integer = 0
                    Dim sAdjustment As String = String.Empty

                    sItemCode = oDvDetail_Grouped(i)(2).ToString.Trim()
                    sDeliveryMode = oDvDetail_Grouped(i)(6).ToString().Trim()
                    sAdjustment = oDvDetail_Grouped(i)(7).ToString().Trim()

                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Delivery type for item " & sItemCode & " is " & sDeliveryMode, sFuncName)

                    sAccountCode = String.Empty
                    dtCheckType.DefaultView.RowFilter = "CHECKTYPE = '" & sDeliveryMode.ToUpper() & "'"
                    If dtCheckType.DefaultView.Count = 0 Then
                        sErrDesc = "Account code not found for checktype :: " & sDeliveryMode
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug(sErrDesc, sFuncName)
                        sErrDesc = ""
                    Else
                        sAccountCode = dtCheckType.DefaultView.Item(0)(1).ToString.Trim()
                    End If

                    dtVatGroup.DefaultView.RowFilter = "ItemCode = '" & sItemCode & "'"
                    If dtVatGroup.DefaultView.Count = 0 Then
                        sErrDesc = "ItemCode :: " & sItemCode & " provided does not exist in SAP."
                        Call WriteToLogFile(sErrDesc, sFuncName)
                        Throw New ArgumentException(sErrDesc)
                    Else
                        sVatGroup = dtVatGroup.DefaultView.Item(0)(1).ToString().Trim()
                    End If

                    If iCount > 0 Then
                        oInvoice.Lines.Add()
                    End If
                    oInvoice.Lines.ItemCode = sItemCode
                    oInvoice.Lines.Quantity = CDbl(oDvDetail_Grouped(i)(9))
                    'oInvoice.Lines.UnitPrice = CDbl(oDvDetail_Grouped(i)(8))
                    oInvoice.Lines.WarehouseCode = oDvDetail_Grouped(i)(4).ToString.Trim()
                    If Not (sOcrCode = String.Empty) Then
                        oInvoice.Lines.CostingCode = sOcrCode
                    End If
                    If Not (sOcrCode2 = String.Empty) Then
                        oInvoice.Lines.CostingCode2 = sOcrCode2
                    End If
                    If Not (sOcrCode3 = String.Empty) Then
                        oInvoice.Lines.CostingCode3 = sOcrCode3
                    End If
                    'oInvoice.Lines.COGSCostingCode3 = oDvDetail_Grouped(i)(4).ToString.Trim()
                    If Not (sAccountCode = String.Empty) Then
                        oInvoice.Lines.AccountCode = sAccountCode
                    End If
                    If sAdjustment <> "" And sAdjustment.ToUpper() <> "REFUND CHECK" Then
                        Try
                            iAdjustment = sAdjustment
                        Catch ex As Exception
                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Adjustment column is not integer", sFuncName)
                        End Try
                        If iAdjustment >= 1 And iAdjustment <= 50 Then
                            oInvoice.Lines.COGSAccountCode = p_oCompDef.sAdjAct1to50
                        End If
                        If iAdjustment >= 51 And iAdjustment <= 99 Then
                            oInvoice.Lines.COGSAccountCode = p_oCompDef.sAdjAct51t99
                        End If
                        If iAdjustment >= 100 And iAdjustment <= 150 Then
                            oInvoice.Lines.COGSAccountCode = p_oCompDef.sAdjAct100to150
                        End If
                        If iAdjustment >= 151 And iAdjustment <= 255 Then
                            oInvoice.Lines.COGSAccountCode = p_oCompDef.sAdjAct151to254
                        End If
                    End If
                    If sVatGroup <> "" Then
                        oInvoice.Lines.VatGroup = sVatGroup
                    End If
                    oInvoice.Lines.LineTotal = CDbl(oDvDetail_Grouped(i)(10))
                    oInvoice.Lines.UserFields.Fields.Item("U_POSNo").Value = oDvDetail_Grouped(i)(4).ToString.Trim()
                    oInvoice.Lines.UserFields.Fields.Item("U_MarketSeg").Value = oDvDetail_Grouped(i)(5).ToString.Trim()
                    oInvoice.Lines.UserFields.Fields.Item("U_ReasonCode").Value = oDvDetail_Grouped(i)(7).ToString.Trim()
                    oInvoice.Lines.UserFields.Fields.Item("U_SalesHour").Value = oDvHeader(0)(7).ToString.Trim()
                    oInvoice.Lines.UserFields.Fields.Item("U_MealPeriod").Value = oDvHeader(0)(6).ToString.Trim()

                    If sAdjustment.ToUpper() = "REFUND CHECK" Then
                        If iCount > 0 Then
                            oInvoice.Lines.Add()
                        End If
                        oInvoice.Lines.ItemCode = sItemCode
                        oInvoice.Lines.Quantity = (-1) * CDbl(oDvDetail_Grouped(i)(9))
                        oInvoice.Lines.WarehouseCode = oDvDetail_Grouped(i)(4).ToString.Trim()
                        If Not (sOcrCode = String.Empty) Then
                            oInvoice.Lines.CostingCode = sOcrCode
                        End If
                        If Not (sOcrCode2 = String.Empty) Then
                            oInvoice.Lines.CostingCode2 = sOcrCode2
                        End If
                        If Not (sOcrCode3 = String.Empty) Then
                            oInvoice.Lines.CostingCode3 = sOcrCode3
                        End If
                        If Not (sAccountCode = String.Empty) Then
                            oInvoice.Lines.AccountCode = sAccountCode
                        End If
                        If Not (p_oCompDef.sRefundAct = String.Empty) Then
                            oInvoice.Lines.COGSAccountCode = p_oCompDef.sRefundAct
                        End If

                        If sVatGroup <> "" Then
                            oInvoice.Lines.VatGroup = sVatGroup
                        End If
                        oInvoice.Lines.LineTotal = CDbl(oDvDetail_Grouped(i)(10))
                        oInvoice.Lines.UserFields.Fields.Item("U_POSNo").Value = oDvDetail_Grouped(i)(4).ToString.Trim()
                        oInvoice.Lines.UserFields.Fields.Item("U_MarketSeg").Value = oDvDetail_Grouped(i)(5).ToString.Trim()
                        oInvoice.Lines.UserFields.Fields.Item("U_ReasonCode").Value = oDvDetail_Grouped(i)(7).ToString.Trim()
                        oInvoice.Lines.UserFields.Fields.Item("U_SalesHour").Value = oDvHeader(0)(7).ToString.Trim()
                        oInvoice.Lines.UserFields.Fields.Item("U_MealPeriod").Value = oDvHeader(0)(6).ToString.Trim()

                        'adding 2nd line if refund
                        If iCount > 0 Then
                            oInvoice.Lines.Add()
                        End If
                        oInvoice.Lines.ItemCode = sItemCode
                        oInvoice.Lines.Quantity = CDbl(oDvDetail_Grouped(i)(9))
                        oInvoice.Lines.WarehouseCode = oDvDetail_Grouped(i)(4).ToString.Trim()
                        If Not (sOcrCode = String.Empty) Then
                            oInvoice.Lines.CostingCode = sOcrCode
                        End If
                        If Not (sOcrCode2 = String.Empty) Then
                            oInvoice.Lines.CostingCode2 = sOcrCode2
                        End If
                        If Not (sOcrCode3 = String.Empty) Then
                            oInvoice.Lines.CostingCode3 = sOcrCode3
                        End If
                        If Not (sAccountCode = String.Empty) Then
                            oInvoice.Lines.AccountCode = sAccountCode
                        End If
                        If Not (p_oCompDef.sRefundAct = String.Empty) Then
                            oInvoice.Lines.COGSAccountCode = p_oCompDef.sRefundAct
                        End If

                        If sVatGroup <> "" Then
                            oInvoice.Lines.VatGroup = sVatGroup
                        End If
                        oInvoice.Lines.LineTotal = 0.0
                        oInvoice.Lines.UserFields.Fields.Item("U_POSNo").Value = oDvDetail_Grouped(i)(4).ToString.Trim()
                        oInvoice.Lines.UserFields.Fields.Item("U_MarketSeg").Value = oDvDetail_Grouped(i)(5).ToString.Trim()
                        oInvoice.Lines.UserFields.Fields.Item("U_ReasonCode").Value = oDvDetail_Grouped(i)(7).ToString.Trim()
                        oInvoice.Lines.UserFields.Fields.Item("U_SalesHour").Value = oDvHeader(0)(7).ToString.Trim()
                        oInvoice.Lines.UserFields.Fields.Item("U_MealPeriod").Value = oDvHeader(0)(6).ToString.Trim()
                    End If

                    bIsLineAdded = True
                    iCount = iCount + 1
                Next
            Else
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("No Detail file is found for file id " & sFileId, sFuncName)
            End If
            If Not (oDvHeader(0)(9).ToString = String.Empty) Then
                If (CDbl(oDvHeader(0)(9).ToString.Trim() <> 0)) Then
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Assigning Service charges", sFuncName)
                    If iCount > 0 Then
                        oInvoice.Lines.Add()
                    End If

                    dtVatGroup.DefaultView.RowFilter = "ItemCode = '" & p_oCompDef.sServChargeItem & "'"
                    If dtVatGroup.DefaultView.Count = 0 Then
                        sErrDesc = "ItemCode :: " & p_oCompDef.sServChargeItem & " provided does not exist in SAP."
                        Call WriteToLogFile(sErrDesc, sFuncName)
                        Throw New ArgumentException(sErrDesc)
                    End If
                    oInvoice.Lines.ItemCode = p_oCompDef.sServChargeItem
                    oInvoice.Lines.Quantity = 1
                    oInvoice.Lines.UnitPrice = CDbl(oDvHeader(0)(9))
                    bIsLineAdded = True
                    iCount = iCount + 1
                End If
            End If
            If Not (oDvHeader(0)(12).ToString = String.Empty) Then
                If (CDbl(oDvHeader(0)(12).ToString.Trim() <> 0)) Then
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Assigning Excess charges", sFuncName)
                    If iCount > 0 Then
                        oInvoice.Lines.Add()
                    End If

                    dtVatGroup.DefaultView.RowFilter = "ItemCode = '" & p_oCompDef.sExcessItem & "'"
                    If dtVatGroup.DefaultView.Count = 0 Then
                        sErrDesc = "ItemCode :: " & p_oCompDef.sExcessItem & " provided does not exist in SAP."
                        Call WriteToLogFile(sErrDesc, sFuncName)
                        Throw New ArgumentException(sErrDesc)
                    End If
                    oInvoice.Lines.ItemCode = p_oCompDef.sExcessItem
                    oInvoice.Lines.Quantity = 1
                    oInvoice.Lines.UnitPrice = CDbl(oDvHeader(0)(12))
                    bIsLineAdded = True
                    iCount = iCount + 1
                End If
            End If
            If Not (oDvHeader(0)(13).ToString = String.Empty) Then
                If (CDbl(oDvHeader(0)(13).ToString.Trim() <> 0)) Then
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Assigning Tipping charges", sFuncName)
                    If iCount > 0 Then
                        oInvoice.Lines.Add()
                    End If

                    dtVatGroup.DefaultView.RowFilter = "ItemCode = '" & p_oCompDef.sTippingItem & "'"
                    If dtVatGroup.DefaultView.Count = 0 Then
                        sErrDesc = "ItemCode :: " & p_oCompDef.sTippingItem & " provided does not exist in SAP."
                        Call WriteToLogFile(sErrDesc, sFuncName)
                        Throw New ArgumentException(sErrDesc)
                    End If
                    oInvoice.Lines.ItemCode = p_oCompDef.sTippingItem
                    oInvoice.Lines.Quantity = 1
                    oInvoice.Lines.UnitPrice = CDbl(oDvHeader(0)(13))
                    bIsLineAdded = True
                    iCount = iCount + 1
                End If
            End If
            If Not (oDvHeader(0)(11).ToString = String.Empty) Then
                If (CDbl(oDvHeader(0)(11).ToString.Trim() <> 0)) Then
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Assigning Rounding charges", sFuncName)
                    If iCount > 0 Then
                        oInvoice.Lines.Add()
                    End If

                    dtVatGroup.DefaultView.RowFilter = "ItemCode = '" & p_oCompDef.sRoundingItem & "'"
                    If dtVatGroup.DefaultView.Count = 0 Then
                        sErrDesc = "ItemCode :: " & p_oCompDef.sRoundingItem & " provided does not exist in SAP."
                        Call WriteToLogFile(sErrDesc, sFuncName)
                        Throw New ArgumentException(sErrDesc)
                    End If
                    oInvoice.Lines.ItemCode = p_oCompDef.sRoundingItem
                    oInvoice.Lines.Quantity = 1
                    oInvoice.Lines.UnitPrice = CDbl(oDvHeader(0)(11))
                    bIsLineAdded = True
                    iCount = iCount + 1
                End If
            End If
            If bIsLineAdded = True Then
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Attempting to Add invoice", sFuncName)
                If oInvoice.Add() <> 0 Then
                    sErrDesc = "Error " & p_oCompany.GetLastErrorDescription
                    Console.WriteLine("Error while adding Invoice")
                    Throw New ArgumentException(sErrDesc)
                Else
                    Dim iDocEntry As Integer
                    iDocEntry = p_oCompany.GetNewObjectKey()
                    sInvDocEntry = iDocEntry

                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oInvoice)

                    Console.WriteLine("Invoice document Created successfully :: " & iDocEntry)
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Invoice document Created successfully :: " & iDocEntry, sFuncName)
                End If
            End If

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            CreateInvoice = RTN_SUCCESS
        Catch ex As Exception
            sErrDesc = ex.Message
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            CreateInvoice = RTN_ERROR
        End Try
    End Function

    Private Function CreatePayment(ByVal oDv As DataView, ByVal sInvDocEntry As String, ByRef sErrDesc As String) As Long
        Dim sFuncName As String = "CreatePayment"
        Dim sSQL As String = String.Empty
        Dim sPosTenderCode As String = String.Empty
        Dim oRs As SAPbobsCOM.Recordset = Nothing
        Dim oPayments As SAPbobsCOM.IPayments = Nothing
        oPayments = p_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oIncomingPayments)
        Dim bIsLineAdded As Boolean = False
        Dim sCardCode As String = String.Empty
        Dim sFileId As String = String.Empty

        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            sFileId = oDv(0)(0).ToString.Trim()

            oRs = p_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            sSQL = "SELECT ""CardCode"" FROM ""OINV"" WHERE ""DocEntry"" = '" & sInvDocEntry & "' "
            oRs.DoQuery(sSQL)
            If oRs.RecordCount > 0 Then
                sCardCode = oRs.Fields.Item("CardCode").Value
            Else
                sErrDesc = "CardCode not found for creating payment for file id " & sFileId
                Throw New ArgumentException(sErrDesc)
            End If

            'oPayments.DocObjectCode = SAPbobsCOM.BoPaymentsObjectType.bopot_IncomingPayments
            oPayments.DocType = SAPbobsCOM.BoRcptTypes.rCustomer

            Dim iIndex As Integer = oDv(0)(2).ToString.IndexOf(" ")
            Dim sDate As String
            If iIndex > -1 Then
                sDate = oDv(0)(2).ToString.Substring(0, iIndex)
            Else
                sDate = oDv(0)(2).ToString
            End If
            Dim dtDocDate As Date
            Dim format() = {"dd/MM/yyyy", "d/M/yyyy", "dd-MM-yyyy", "dd.MM.yyyy", "yyyyMMdd", "MMddYYYY", "M/dd/yyyy", "MM/dd/YYYY"}
            Date.TryParseExact(sDate, format, System.Globalization.DateTimeFormatInfo.InvariantInfo, Globalization.DateTimeStyles.None, dtDocDate)

            oPayments.CardCode = sCardCode
            oPayments.DocDate = dtDocDate
            oPayments.UserFields.Fields.Item("U_WHSCode").Value = oDv(0)(1).ToString.Trim()
            oPayments.UserFields.Fields.Item("U_POSNo").Value = oDv(0)(1).ToString.Trim()

            Console.WriteLine("Selecting Payment methods")
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Selecting Payment methods", sFuncName)

            Dim dTotalPayAmt As Double = 0.0
            Dim dPayAmt As Double = 0.0
            For i As Integer = 0 To oDv.Count - 1
                If Not (oDv(i)(0).ToString = String.Empty) Then
                    If Not (oDv(i)(4).ToString.Trim = String.Empty) Then
                        dPayAmt = oDv(i)(4).ToString.Trim
                    Else
                        dPayAmt = 0.0
                    End If

                    dTotalPayAmt = Math.Round(dTotalPayAmt, 2) + Math.Round(dPayAmt, 2)
                End If
            Next

            If dTotalPayAmt > 0.0 Then
                oPayments.Invoices.DocEntry = Convert.ToInt32(sInvDocEntry)
                oPayments.Invoices.InvoiceType = SAPbobsCOM.BoRcptInvTypes.it_Invoice
                oPayments.Invoices.DocLine = 0
                oPayments.Invoices.DiscountPercent = 0.0
                oPayments.Invoices.SumApplied = dTotalPayAmt
                oPayments.Invoices.Add()

                For j As Integer = 0 To oDv.Count - 1
                    sPosTenderCode = oDv(j)(3).ToString.Trim()

                    sSQL = "SELECT T0.""CreditCard"" FROM ""OCRC"" T0 WHERE UPPER(T0.""CardName"") ='" & sPosTenderCode.ToUpper() & "'"
                    oRs.DoQuery(sSQL)
                    If oRs.RecordCount > 0 Then
                        oPayments.CreditCards.CreditCard = oRs.Fields.Item("CreditCard").Value
                        oPayments.CreditCards.CreditCardNumber = oDv(j)(0).ToString.Trim()
                        oPayments.CreditCards.CreditSum = CDbl(oDv(j)(4))
                        Dim sCrdtValidDt As Date = "9999-12-01"
                        oPayments.CreditCards.CardValidUntil = sCrdtValidDt
                        oPayments.CreditCards.VoucherNum = oDv(j)(1).ToString.Trim() & "-" & DateTime.Now.ToString("yyyyMMdd")
                        oPayments.CreditCards.Add()
                    Else
                        sErrDesc = "Credit card details for : " & sPosTenderCode & " Not found"
                        Throw New ArgumentException(sErrDesc)
                    End If
                Next

                bIsLineAdded = True
            Else
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Total payment amount for file id " & sFileId & " is 0. No Payment will be created", sFuncName)
            End If

            If bIsLineAdded = True Then
                If oPayments.Add() <> 0 Then
                    sErrDesc = p_oCompany.GetLastErrorDescription()
                    Throw New ArgumentException(sErrDesc)
                Else
                    Dim iDocEntry As Integer
                    p_oCompany.GetNewObjectCode(iDocEntry)

                    Console.WriteLine("Payment document successfully created. DocEntry is :: " & iDocEntry)
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Payment document successfully created. DocEntry is :: " & iDocEntry, sFuncName)

                End If
            End If

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            CreatePayment = RTN_SUCCESS
        Catch ex As Exception
            sErrDesc = ex.Message
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            CreatePayment = RTN_ERROR
        End Try
    End Function

    Private Function CreatePayment_Working_Backup(ByVal oDv As DataView, ByVal sInvDocEntry As String, ByRef sErrDesc As String) As Long
        Dim sFuncName As String = "CreatePayment_Working_Backup"
        Dim sSQL As String = String.Empty
        Dim sPosTenderCode As String = String.Empty
        Dim oRs As SAPbobsCOM.Recordset = Nothing
        Dim oPayments As SAPbobsCOM.IPayments = Nothing
        oPayments = p_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oIncomingPayments)
        Dim bIsLineAdded As Boolean = False
        Dim sCardCode As String = String.Empty
        Dim sFileId As String = String.Empty

        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            sFileId = oDv(0)(0).ToString.Trim()

            oRs = p_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            sSQL = "SELECT ""CardCode"" FROM ""OINV"" WHERE ""DocEntry"" = '" & sInvDocEntry & "' "
            oRs.DoQuery(sSQL)
            If oRs.RecordCount > 0 Then
                sCardCode = oRs.Fields.Item("CardCode").Value
            Else
                sErrDesc = "CardCode not found for creating payment for file id " & sFileId
                Throw New ArgumentException(sErrDesc)
            End If

            'oPayments.DocObjectCode = SAPbobsCOM.BoPaymentsObjectType.bopot_IncomingPayments
            oPayments.DocType = SAPbobsCOM.BoRcptTypes.rCustomer

            Dim iIndex As Integer = oDv(0)(2).ToString.IndexOf(" ")
            Dim sDate As String
            If iIndex > -1 Then
                sDate = oDv(0)(2).ToString.Substring(0, iIndex)
            Else
                sDate = oDv(0)(2).ToString
            End If
            Dim dtDocDate As Date
            Dim format() = {"dd/MM/yyyy", "d/M/yyyy", "dd-MM-yyyy", "dd.MM.yyyy", "yyyyMMdd", "MMddYYYY", "M/dd/yyyy", "MM/dd/YYYY"}
            Date.TryParseExact(sDate, format, System.Globalization.DateTimeFormatInfo.InvariantInfo, Globalization.DateTimeStyles.None, dtDocDate)

            oPayments.CardCode = sCardCode
            oPayments.DocDate = dtDocDate
            oPayments.UserFields.Fields.Item("U_WHSCode").Value = oDv(0)(1).ToString.Trim()
            oPayments.UserFields.Fields.Item("U_POSNo").Value = oDv(0)(1).ToString.Trim()

            Console.WriteLine("Selecting Payment methods")
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Selecting Payment methods", sFuncName)

            Dim dTotalPayAmt As Double = 0.0
            Dim dPayAmt As Double = 0.0
            For i As Integer = 0 To oDv.Count - 1
                If Not (oDv(i)(0).ToString = String.Empty) Then
                    If Not (oDv(i)(4).ToString.Trim = String.Empty) Then
                        dPayAmt = oDv(i)(4).ToString.Trim
                    Else
                        dPayAmt = 0.0
                    End If

                    dTotalPayAmt = Math.Round(dTotalPayAmt, 2) + Math.Round(dPayAmt, 2)
                End If
            Next

            If dTotalPayAmt > 0.0 Then
                oPayments.Invoices.DocEntry = Convert.ToInt32(sInvDocEntry)
                oPayments.Invoices.InvoiceType = SAPbobsCOM.BoRcptInvTypes.it_Invoice
                oPayments.Invoices.DocLine = 0
                oPayments.Invoices.SumApplied = dTotalPayAmt
                oPayments.Invoices.Add()

                For j As Integer = 0 To oDv.Count - 1
                    sPosTenderCode = oDv(j)(3).ToString.Trim()

                    sSQL = "SELECT T0.""CreditCard"" FROM ""OCRC"" T0 WHERE UPPER(T0.""CardName"") ='" & sPosTenderCode.ToUpper() & "'"
                    oRs.DoQuery(sSQL)
                    If oRs.RecordCount > 0 Then
                        oPayments.CreditCards.CreditCard = oRs.Fields.Item("CreditCard").Value
                        oPayments.CreditCards.CreditCardNumber = oDv(j)(0).ToString.Trim()
                        oPayments.CreditCards.CreditSum = CDbl(oDv(j)(4))
                        Dim sCrdtValidDt As Date = "9999-12-01"
                        oPayments.CreditCards.CardValidUntil = sCrdtValidDt
                        oPayments.CreditCards.VoucherNum = oDv(j)(1).ToString.Trim() & "-" & DateTime.Now.ToString("yyyyMMdd")
                        oPayments.CreditCards.Add()
                    Else
                        sErrDesc = "Credit card details for : " & sPosTenderCode & " Not found"
                        Throw New ArgumentException(sErrDesc)
                    End If
                Next

                bIsLineAdded = True
            Else
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Total payment amount for file id " & sFileId & " is 0. No Payment will be created", sFuncName)
            End If

            If bIsLineAdded = True Then
                If oPayments.Add() <> 0 Then
                    sErrDesc = p_oCompany.GetLastErrorDescription()
                    Throw New ArgumentException(sErrDesc)
                Else
                    Dim iDocEntry As Integer
                    p_oCompany.GetNewObjectCode(iDocEntry)

                    Console.WriteLine("Payment document successfully created. DocEntry is :: " & iDocEntry)
                    If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Payment document successfully created. DocEntry is :: " & iDocEntry, sFuncName)

                End If
            End If

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            CreatePayment_Working_Backup = RTN_SUCCESS
        Catch ex As Exception
            sErrDesc = ex.Message
            Call WriteToLogFile(sErrDesc, sFuncName)
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            CreatePayment_Working_Backup = RTN_ERROR
        End Try
    End Function

End Module
