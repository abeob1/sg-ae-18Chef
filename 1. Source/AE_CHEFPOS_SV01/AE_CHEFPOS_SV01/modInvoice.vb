Module modInvoice

    Private dtBPMaster As DataTable
    Private dtCheckType As DataTable
    Private dtVatGroup As DataTable

    Public Function ProcessInvoiceFiles(ByVal file_Header As System.IO.FileInfo, ByVal file_Detail As System.IO.FileInfo, ByVal oDvHeader As DataView, ByVal oDvDeatil As DataView, ByRef sErrDesc As String) As Long
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
                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling CreateAPInvoice_NonCap()", sFuncName)
                            Dim oDtHeader_Grouped As DataTable
                            oDtHeader_Grouped = oDvHeader.ToTable
                            Dim oDvHeader_Grouped As DataView = New DataView(oDtHeader_Grouped)

                            If CreateInvoice(oDvHeader_Grouped, oDvDeatil, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
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
            'error condition

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling FileMoveToArchive()", sFuncName)
            FileMoveToArchive(file_Header, file_Header.FullName, RTN_ERROR)
            FileMoveToArchive(file_Detail, file_Detail.FullName, RTN_ERROR)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            ProcessInvoiceFiles = RTN_ERROR
        End Try
    End Function

    Private Function CreateInvoice(ByVal oDvHeader As DataView, ByVal oDvDetail As DataView, ByRef sErrDesc As String) As Long
        Dim sFuncName As String = "CreateInvoice"
        Dim sFileId As String = String.Empty
        Dim sCardCode As String = String.Empty
        Dim bIsLineAdded As Boolean = False

        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            sFileId = oDvHeader(0)(0).ToString().Trim()

            oDvDetail.RowFilter = Nothing
            oDvDetail.RowFilter = "FILEID = '" & sFileId & "'"
            Dim odt As New DataTable
            odt = oDvDetail.ToTable
            Dim oDvDetail_Grouped As DataView = New DataView(odt)

            If oDvDetail_Grouped.Count > 0 Then
                Dim oInvoice As SAPbobsCOM.Documents
                oInvoice = p_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices)

                sCardCode = "S" & oDvHeader(0)(4).ToString.Trim()
                dtBPMaster.DefaultView.RowFilter = "UPPERCARDCODE = '" & sCardCode.ToUpper() & "'"
                If dtBPMaster.DefaultView.Count = 0 Then
                    sErrDesc = "CardCode :: " & sCardCode & " Not exists in SAP."
                    Console.WriteLine(sErrDesc)
                    Call WriteToLogFile(sErrDesc, sFuncName)
                    Throw New ArgumentException(sErrDesc)
                Else
                    sCardCode = dtBPMaster.DefaultView.Item(0)(0).ToString().Trim()
                End If

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

                oInvoice.CardCode = sCardCode
                oInvoice.DocDate = dtDocDate
                oInvoice.NumAtCard = oDvHeader(0)(4).ToString.Trim() & "-" & sDate & "-" & oDvHeader(0)(7).ToString.Trim()
                oInvoice.UserFields.Fields.Item("U_Concept").Value = oDvHeader(0)(2).ToString.Trim()
                oInvoice.UserFields.Fields.Item("U_BRAND").Value = oDvHeader(0)(3).ToString.Trim()
                oInvoice.UserFields.Fields.Item("U_Outlet").Value = oDvHeader(0)(4).ToString.Trim()
                oInvoice.UserFields.Fields.Item("U_POSNo").Value = oDvHeader(0)(4).ToString.Trim()
                oInvoice.UserFields.Fields.Item("U_MEALPERIOD").Value = oDvHeader(0)(6).ToString.Trim()
                oInvoice.UserFields.Fields.Item("U_HOUR").Value = oDvHeader(0)(7).ToString.Trim()
                oInvoice.UserFields.Fields.Item("U_Covers").Value = oDvHeader(0)(14).ToString.Trim()
                oInvoice.UserFields.Fields.Item("U_NetTables").Value = oDvHeader(0)(15).ToString.Trim()
                oInvoice.DocTotal = CDbl(oDvHeader(0)(8))

                Dim iCount As Integer = 0
                For i As Integer = 0 To oDvDetail_Grouped.Count - 1
                    Dim sItemCode As String = String.Empty
                    Dim sDeliveryMode As String = String.Empty
                    Dim sAccountCode As String = String.Empty
                    Dim sVatGroup As String = String.Empty

                    sItemCode = oDvDetail_Grouped(i)(2).ToString.Trim()
                    sDeliveryMode = oDvDetail_Grouped(i)(6).ToString().Trim()

                    dtCheckType.DefaultView.RowFilter = "CHECKTYPE = '" & sDeliveryMode.ToUpper() & "'"
                    If dtCheckType.DefaultView.Count = 0 Then
                        sErrDesc = "Account code not found for checktype :: " & sDeliveryMode
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug(sErrDesc, sFuncName)
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
                    oInvoice.Lines.UnitPrice = CDbl(oDvDetail_Grouped(i)(8))
                    oInvoice.Lines.WarehouseCode = oDvDetail_Grouped(i)(4).ToString.Trim()
                    oInvoice.Lines.CostingCode = oDvHeader(0)(3).ToString.Trim()
                    oInvoice.Lines.CostingCode2 = oDvHeader(0)(2).ToString.Trim()
                    oInvoice.Lines.CostingCode3 = oDvDetail_Grouped(i)(4).ToString.Trim()
                    oInvoice.Lines.COGSCostingCode3 = oDvDetail_Grouped(i)(4).ToString.Trim()
                    If Not (sAccountCode = String.Empty) Then
                        oInvoice.Lines.AccountCode = sAccountCode
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
                    bIsLineAdded = True
                    iCount = iCount + 1
                Next
                If Not (oDvHeader(0)(9).ToString = String.Empty) Then
                    If (CDbl(oDvHeader(0)(9).ToString.Trim() <> 0)) Then
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
                If Not (oDvHeader(0)(11).ToString = String.Empty) Then
                    If (CDbl(oDvHeader(0)(11).ToString.Trim() <> 0)) Then
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
                If Not (oDvHeader(0)(12).ToString = String.Empty) Then
                    If (CDbl(oDvHeader(0)(12).ToString.Trim() <> 0)) Then
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

                If bIsLineAdded = True Then
                    If oInvoice.Add() <> 0 Then
                        sErrDesc = "Error " & p_oCompany.GetLastErrorDescription
                        Console.WriteLine("Error while adding Invoice")
                        Throw New ArgumentException(sErrDesc)
                    Else
                        Dim iDocEntry As Integer
                        iDocEntry = p_oCompany.GetNewObjectKey()
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(oInvoice)

                        Console.WriteLine("Invoice document Created successfully :: " & iDocEntry)
                        If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Invoice document Created successfully :: " & iDocEntry, sFuncName)
                    End If
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

End Module
