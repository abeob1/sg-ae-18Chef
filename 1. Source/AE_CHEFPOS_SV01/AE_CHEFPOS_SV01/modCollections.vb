Module modCollections

    Private dtCheckType As DataTable
    Private dtTenderCode As DataTable

    Public Function ProcessCollectionDetails(ByVal oDvCollections As DataView, ByVal file As System.IO.FileInfo, ByRef sErrDesc As String) As Long
        Dim sFuncName As String = "ProcessCollectionDetails"
        Dim sSQL As String = String.Empty

        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting function", sFuncName)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling ConnectToTargetCompany()", sFuncName)
            Console.WriteLine("Connecting Company")
            If ConnectToCompany(p_oCompany, p_oCompDef.sSAPDBName, p_oCompDef.sSAPUser, p_oCompDef.sSAPPwd, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

            If p_oCompany.Connected Then
                Console.WriteLine("Company connected to " & p_oCompany.CompanyDB)

                sSQL = "SELECT UPPER(""U_CHECKTYPE"") AS ""CHECKTYPE"",U_SAPACCOUNT FROM " & p_oCompDef.sSAPDBName & ".""@AE_CHECKTYPE"" "
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("EXECUTING SQL :" & sSQL, sFuncName)
                dtCheckType = ExecuteQueryReturnDataTable(sSQL, p_oCompDef.sSAPDBName)

                sSQL = "SELECT UPPER(T0.""U_POS_TENDER_CODE"") AS ""U_POS_TENDER_CODE"", UPPER(T0.""U_SAP_TENDER_CODE"") AS ""U_SAP_TENDER_CODE"" FROM " & p_oCompDef.sSAPDBName & ".""@AE_TENDERCODE""  T0 "
                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("EXECUTING SQL :" & sSQL, sFuncName)
                dtTenderCode = ExecuteQueryReturnDataTable(sSQL, p_oCompDef.sSAPDBName)

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling StartTransaction", sFuncName)
                If StartTransaction(sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

                If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Grouping header data based on fileid", sFuncName)
                Dim oDtGroup As DataTable = oDvCollections.Table.DefaultView.ToTable(True, "FILEID")
                For i As Integer = 0 To oDtGroup.Rows.Count - 1
                    If Not (oDtGroup.Rows(i).Item(0).ToString.Trim() = String.Empty Or oDtGroup.Rows(i).Item(0).ToString.ToUpper().Trim() = "FILEID") Then
                        oDvCollections.RowFilter = "FILEID ='" & oDtGroup.Rows(i).Item(0).ToString.Trim() & "' "

                        If oDvCollections.Count > 0 Then
                            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling CreatePayment()", sFuncName)
                            Dim oDtCollections_Grouped As DataTable
                            oDtCollections_Grouped = oDvCollections.ToTable
                            Dim oDvCollections_Grouped As DataView = New DataView(oDtCollections_Grouped)

                            If CreatePayment(oDvCollections_Grouped, sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)
                        End If

                    End If
                Next
            End If

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling CommitTransaction", sFuncName)
            If CommitTransaction(sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling FileMoveToArchive()", sFuncName)
            FileMoveToArchive(file, file.FullName, RTN_SUCCESS)

            'Insert Success Notificaiton into Table..
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling AddDataToTable()", sFuncName)
            AddDataToTable(p_oDtSuccess, file.Name, "Success")

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with SUCCESS", sFuncName)
            ProcessCollectionDetails = RTN_SUCCESS
        Catch ex As Exception
            sErrDesc = ex.Message
            Call WriteToLogFile(sErrDesc, sFuncName)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling RollbackTransaction", sFuncName)
            If RollbackTransaction(sErrDesc) <> RTN_SUCCESS Then Throw New ArgumentException(sErrDesc)

            'Insert Error Description into Table
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling AddDataToTable()", sFuncName)
            AddDataToTable(p_oDtError, file.Name, "Error", sErrDesc)
            'error condition

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Calling FileMoveToArchive()", sFuncName)
            FileMoveToArchive(file, file.FullName, RTN_ERROR)

            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Completed with ERROR", sFuncName)
            ProcessCollectionDetails = RTN_ERROR
        End Try
    End Function

    Private Function CreatePayment(ByVal oDv As DataView, ByRef sErrDesc As String) As Long
        Dim sFuncName As String = "CreatePayment"
        Dim sSQL As String = String.Empty
        Dim sPosTenderCode As String = String.Empty
        Dim oRs As SAPbobsCOM.Recordset = Nothing
        Dim oPayments As SAPbobsCOM.IPayments = Nothing
        oPayments = p_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oIncomingPayments)
        Dim bIsLineAdded As Boolean = False

        Try
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Starting Function", sFuncName)

            oRs = p_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            oPayments.DocType = SAPbobsCOM.BoRcptTypes.rAccount

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

            oPayments.DocDate = dtDocDate
            oPayments.UserFields.Fields.Item("U_WHSCode").Value = oDv(0)(1).ToString.Trim()
            oPayments.UserFields.Fields.Item("U_POSNo").Value = oDv(0)(1).ToString.Trim()

            Console.WriteLine("Selecting Payment methods")
            If p_iDebugMode = DEBUG_ON Then Call WriteToLogFile_Debug("Selecting Payment methods", sFuncName)

            Dim iCount As Integer = 0
            For j As Integer = 0 To oDv.Count - 1
                Dim sSAPTenderCode As String = String.Empty
                Dim sAccountCode As String = String.Empty
                sPosTenderCode = oDv(j)(3).ToString.Trim()
                dtCheckType.DefaultView.RowFilter = "CHECKTYPE = '" & sPosTenderCode.ToUpper() & "'"
                If dtCheckType.DefaultView.Count = 0 Then
                    sErrDesc = "Account code not exists for check type  :: " & sPosTenderCode & ". please check checktype table."
                    Console.WriteLine(sErrDesc)
                    Call WriteToLogFile(sErrDesc, sFuncName)
                    Throw New ArgumentException(sErrDesc)
                Else
                    sAccountCode = dtCheckType.DefaultView.Item(0)(1).ToString().Trim()
                End If

                If iCount > 1 Then
                    oPayments.AccountPayments.Add()
                End If

                oPayments.AccountPayments.AccountCode = sAccountCode
                oPayments.AccountPayments.GrossAmount = CDbl(oDv(j)(4))
                bIsLineAdded = True
                iCount = iCount + 1
            Next

            For j As Integer = 0 To oDv.Count - 1
                Dim sSAPTenderCode As String = String.Empty
                sPosTenderCode = oDv(j)(3).ToString.Trim()
                dtTenderCode.DefaultView.RowFilter = "U_POS_TENDER_CODE = '" & sPosTenderCode.ToUpper() & "'"
                If dtTenderCode.DefaultView.Count = 0 Then
                    sErrDesc = "Tendercode  :: " & sPosTenderCode & " Not exists in table."
                    Console.WriteLine(sErrDesc)
                    Call WriteToLogFile(sErrDesc, sFuncName)
                    Throw New ArgumentException(sErrDesc)
                Else
                    sSAPTenderCode = dtTenderCode.DefaultView.Item(0)(1).ToString().Trim()
                End If

                sSQL = "SELECT T0.""CreditCard"" FROM ""OCRC"" T0 WHERE UPPER(T0.""CardName"") ='" & sSAPTenderCode.ToUpper() & "'"
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
                    sErrDesc = "Credit card details for : " & sSAPTenderCode & " Not found"
                    Throw New ArgumentException(sErrDesc)
                End If
            Next

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

End Module
