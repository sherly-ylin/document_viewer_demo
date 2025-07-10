
Private Function GetMergeDataOR(pField As String) As String
    On Error GoTo errHandler:
    
    Const mstrMethodName As String = "GetMergeDataOR"
    Dim mstrTemp As String
    Dim mintMachineNum As Integer
    Dim mintLineNum As Integer
    Dim mstrPrefix As String
    Dim mobjDom As MSXML2.DOMDocument
    Dim mobjNode As MSXML2.IXMLDOMNode
    Dim TagIndex As String
    Dim mobjDBServer As Object
    Dim mstrSQL As String
    Dim mstrXML As String
    Dim mRS As ADODB.Recordset
    Dim mRSAT As ADODB.Recordset
    Dim mobjCommon As Object
    Dim mstrTagPrefix As String
    Dim mIndex As Long
    Dim mint As Long
    Dim mstrAddendumname As String
    Dim mstrCatalogServer As String
    Dim mstrAdjTags As String
    Dim mstrFields As String
                       
    mstrCatalogServer = ""
    
    ''need to hold values in array and check before querying
    If cflgInMTS = True Then
        Set mobjCommon = cobjContext.CreateInstance("SNCommon.CCommon")
        Set mobjDBServer = cobjContext.CreateInstance("SNDBService.CDBServer")
    Else
        Set mobjCommon = CreateObject("SNCommon.CCommon")
        Set mobjDBServer = CreateObject("SNDBService.CDBServer")
    End If
    
    mstrTagPrefix = Mid(pField, 1, 2)
    If InStr(pField, ":") > 0 Then
        TagIndex = Right(pField, Len(pField) - InStrRev(pField, ":"))
        If InStr(pField, "ADDENDUMNEEDED") > 0 And InStr(pField, "-") > 0 Then
            mIndex = CInt(Left(TagIndex, InStrRev(TagIndex, "-") - 1))
            mstrAddendumname = Right(pField, Len(pField) - InStrRev(pField, "-"))
        End If
        pField = UCase(Mid(pField, 1, InStrRev(pField, ":") - 1))
    End If
        
    mstrXML = GetMergeDataXML(mstrTagPrefix)
    
    Select Case mstrTagPrefix
           
        Case "OR"
            If UCase(pField) = "OR.SUMADJ" Then
                If TagIndex = "" Then
                    GetMergeDataOR = ""
                Else
                    TagIndex = Replace(TagIndex, "_", ",")
                    mstrSQL = ""
                    mstrSQL = mstrSQL & " SELECT dbo.fn_FormatCurrency(sum(AdjAmount)) as SumAdj"
                    mstrSQL = mstrSQL & " FROM (SELECT OA.OrderAdjTypeID, case when OAT.CreditDebitInd = 'D' then OA.AdjAmount else OA.AdjAmount * -1 end as AdjAmount"
                    mstrSQL = mstrSQL & "       FROM SNOrderAdjustment OA with (nolock)"
                    mstrSQL = mstrSQL & "       JOIN SNOrderAdjType OAT with (nolock) on OAT.OrderAdjTypeID = OA.OrderAdjTypeID"
                    mstrSQL = mstrSQL & "       WHERE OA.OrderID = " & clngObjectKey
                    mstrSQL = mstrSQL & "       UNION"
                    mstrSQL = mstrSQL & "       SELECT 0, SubTotalAmount"
                    mstrSQL = mstrSQL & "       FROM SNOrder with (nolock)"
                    mstrSQL = mstrSQL & "       WHERE OrderID = " & clngObjectKey
                    mstrSQL = mstrSQL & "       )d"
                    mstrSQL = mstrSQL & " WHERE OrderAdjTypeID in (" & TagIndex & ")"
                    
                    Set mRS = mobjDBServer.DoQuery(CStr(evDBConnString), CStr(mstrSQL))
                    
                    GetMergeDataOR = mRS("SumAdj")
                End If
            ElseIf UCase(pField) = "OR.DELJOBOPTIONMAX" Then
                If TagIndex = "" Then
                    GetMergeDataOR = ""
                Else
                    mstrSQL = ""
                    mstrSQL = mstrSQL & " SELECT top 1 dbo.fn_FormatXMLChars(isnull(OV.Name,'')) as Name"
                    mstrSQL = mstrSQL & " FROM SNDelJobOption JO with (nolock)"
                    mstrSQL = mstrSQL & " JOIN SNDelJobOptionValue OV with (nolock) on OV.DelJobOptionTypeID = JO.DelJobOptionTypeID and OV.DelJobOptionValueID = JO.OptionValue"
                    mstrSQL = mstrSQL & " WHERE JO.DeliveryJobID in (SELECT DeliveryJobID FROM SNDeliveryJob with (nolock) WHERE OrderID = " & clngObjectKey & ")"
                    mstrSQL = mstrSQL & "   AND JO.DelJobOptionTypeID = " & TagIndex
                    mstrSQL = mstrSQL & " ORDER BY OV.SeqNbr desc"
                    
                    Set mRS = mobjDBServer.DoQuery(CStr(evDBConnString), CStr(mstrSQL))
                    
                    If mRS.RecordCount > 0 Then
                        GetMergeDataOR = mRS("Name")
                    Else
                        GetMergeDataOR = ""
                    End If
                End If
            Else
                 If Len(mstrXML) = 0 Then
                    
                    mstrAdjTags = ""
                     
                    mstrSQL = ""
                    mstrSQL = mstrSQL & " SELECT *, -1 AdjustOn FROM SNSPMAdjustType with (nolock) "
                    mstrSQL = mstrSQL & " WHERE charIndex(cast(SPMAdjustTypeID as nvarchar), (SELECT ActiveSPMAdjustTypeIDs FROM SNOrder with (nolock) "
                    If UCase(cstrObjectClassName) = "CORDER" Then
                         mstrSQL = mstrSQL & " Where OrderID = " & clngObjectKey
                    Else
                         mstrSQL = mstrSQL & " Where OrderID = (SELECT OrderID FROM SNDeliveryJob with (nolock) WHERE DeliveryJobID = " & clngObjectKey & ")"
                    End If
                    mstrSQL = mstrSQL & " )) > 0 "
                    mstrSQL = mstrSQL & " UNION "
                    mstrSQL = mstrSQL & " SELECT *, 0 AdjustOn FROM SNSPMAdjustType with (nolock) "
                    mstrSQL = mstrSQL & " WHERE charIndex(cast(SPMAdjustTypeID as nvarchar), (SELECT ActiveSPMAdjustTypeIDs FROM SNOrder with (nolock) "
                    If UCase(cstrObjectClassName) = "CORDER" Then
                         mstrSQL = mstrSQL & " Where OrderID = " & clngObjectKey
                    Else
                         mstrSQL = mstrSQL & " Where OrderID = (SELECT OrderID FROM SNDeliveryJob with (nolock) WHERE DeliveryJobID = " & clngObjectKey & ")"
                    End If
                    mstrSQL = mstrSQL & " )) <= 0 "
                    Set mRSAT = mobjDBServer.DoQuery(CStr(evDBConnString), CStr(mstrSQL))
                    
                    For mint = 1 To mRSAT.RecordCount
                         mstrAdjTags = mstrAdjTags & " + '<OR.ADJ" & mRSAT("SPMAdjustTypeID") & "NAME>' + dbo.fn_FormatXMLChars('" & mRSAT("Name") & "') + '</OR.ADJ" & mRSAT("SPMAdjustTypeID") & "NAME>'"
                         mstrAdjTags = mstrAdjTags & " + '<OR.ADJ" & mRSAT("SPMAdjustTypeID") & "CODENAME>' + dbo.fn_FormatXMLChars('" & mRSAT("CodeName") & "') + '</OR.ADJ" & mRSAT("SPMAdjustTypeID") & "CODENAME>'"
                         mstrAdjTags = mstrAdjTags & " + '<OR.ADJ" & mRSAT("SPMAdjustTypeID") & "CONTRACTLABEL>' + case when " & mRSAT("AdjustOn") & " = -1 then dbo.fn_FormatXMLChars('" & mRSAT("ContractLabelOn") & "') else dbo.fn_FormatXMLChars('" & mRSAT("ContractLabelOff") & "') end + '</OR.ADJ" & mRSAT("SPMAdjustTypeID") & "CONTRACTLABEL>'"
                         
                         mRSAT.MoveNext
                    Next
                    
                    Set mRSAT = Nothing
                    
                    mstrSQL = ""
                    mstrSQL = mstrSQL & " SELECT LeaseAdjustTypeID FROM SNLeaseAdjustType with (nolock) "
                    Set mRSAT = mobjDBServer.DoQuery(CStr(evDBConnString), CStr(mstrSQL))
                    
                    For mint = 1 To mRSAT.RecordCount
                         mstrAdjTags = mstrAdjTags & " + '<OR.LEASEADJ" & mRSAT("LeaseAdjustTypeID") & "FACTOR>' + case when isnull(ST.IsLeased,0) = 0 and OD.SaleTypeID > 0 then 'n/a' else case when isnull(OD.Lease_Adj" & mRSAT("LeaseAdjustTypeID") & "ActiveInd,0) = 0 then '' else case when isnull(OD.Lease_Adj" & mRSAT("LeaseAdjustTypeID") & "Factor,0) <= 0 then isnull(nullif(dbo.fn_formatnumber(OD.Lease_Adj" & mRSAT("LeaseAdjustTypeID") & "Factor * -1,5),'0.00000'),'') else '(' + isnull(nullif(dbo.fn_formatnumber(OD.Lease_Adj" & mRSAT("LeaseAdjustTypeID") & "Factor,5),'0.00000'),'') + ')' end end end + '</OR.LEASEADJ" & mRSAT("LeaseAdjustTypeID") & "FACTOR>'"
                         
                         mRSAT.MoveNext
                    Next
                    
                
                     mstrSQL = ""
                     mstrSQL = mstrSQL & " SELECT '<OR>' "
                     mstrSQL = mstrSQL & " + '<OR.ACCESSORIES>' + dbo.fn_FormatXMLChars(isnull((select stuff((select ', ' + Model from SNOrderLine with (nolock) where OrderID = OD.OrderID and SeqNbr >= case when OD.AccessoryOnlyInd = 0 then 2 else 1 end order by SeqNbr for xml path('')),1,2,'')),'')) + '</OR.ACCESSORIES>'"
                     mstrSQL = mstrSQL & " + '<OR.ADJCOSTAMT>' + dbo.fn_FormatCurrency(((SELECT SUM(ISNULL(BuyPrice,0) * Quantity) FROM SNOrderLine with (nolock) WHERE ISNULL(ServiceInd,0) = 0 AND OrderID = OD.OrderID) - OD.SRCommCSMPCreditValue)) + '</OR.ADJCOSTAMT>'"
                     mstrSQL = mstrSQL & " + '<OR.ADJGPAMT>' + dbo.fn_FormatCurrency(OD.AdjGPAmt) + '</OR.ADJGPAMT>'"
                     mstrSQL = mstrSQL & " + '<OR.ADJGPPCT>' + dbo.fn_formatnumber(OD.AdjGPPCT * 100,2) + '%' + '</OR.ADJGPPCT>'"
                     mstrSQL = mstrSQL & " + '<OR.AMTFINANCED>' + case when ST.IsLeased = 0 and OD.SaleTypeID > 0 then 'n/a' else dbo.fn_FormatCurrency(OD.AmtFinanced) end + '</OR.AMTFINANCED>'"
                     mstrSQL = mstrSQL & " + '<OR.AMTFIN-LBO>' + case when isnull(OD.AmtFinanced,0) = 0 then '' else dbo.fn_FormatCurrency(OD.AmtFinanced - OD.BuyoutAmt) end + '</OR.AMTFIN-LBO>'"
                     mstrSQL = mstrSQL & " + '<OR.BASEBILLFREQANX>' +  case when isnull(BT.PaymentsPerYear,0) = 1 then 'X' else '' end + '</OR.BASEBILLFREQANX>'"
                     mstrSQL = mstrSQL & " + '<OR.BASEBILLFREQMOX>' +  case when isnull(BT.PaymentsPerYear,0) = 12 then 'X' else '' end + '</OR.BASEBILLFREQMOX>'"
                     mstrSQL = mstrSQL & " + '<OR.BASEBILLFREQQTRX>' +  case when isnull(BT.PaymentsPerYear,0) = 4 then 'X' else '' end + '</OR.BASEBILLFREQQTRX>'"
                     mstrSQL = mstrSQL & " + '<OR.BASEBILLFREQSAX>' +  case when isnull(BT.PaymentsPerYear,0) = 2 then 'X' else '' end + '</OR.BASEBILLFREQSAX>'"
                     mstrSQL = mstrSQL & " + '<OR.BASEBILLINLSNOX>' +  case when isnull(BT.IncludeInLease,0) = 0 then 'X' else '' end + '</OR.BASEBILLINLSNOX>'"
                     mstrSQL = mstrSQL & " + '<OR.BASEBILLINLSYESX>' +  case when isnull(BT.IncludeInLease,0) = -1 then 'X' else '' end + '</OR.BASEBILLINLSYESX>'"
                     mstrSQL = mstrSQL & " + '<OR.BASEGPAMT>' + dbo.fn_FormatCurrency(OD.GPTotal) + '</OR.BASEGPAMT>'"
                     mstrSQL = mstrSQL & " + '<OR.BASEGPPCT>' + dbo.fn_formatnumber(OD.TotalGPPCT * 100,2) + '%' + '</OR.BASEGPPCT>'"
                     mstrSQL = mstrSQL & " + '<OR.BASELEASEPAYMENT>' + case when isnull(ST.IsLeased,0) = 0 and OD.SaleTypeID > 0 then 'n/a' else dbo.fn_FormatCurrency(isnull(OD.AmtFinanced ,0) * isnull(OD.LeaseFactor ,0))  end + '</OR.BASELEASEPAYMENT>'"
                     mstrSQL = mstrSQL & " + '<OR.BASELEASEPAYMENTNOSIGN>' + case when isnull(ST.IsLeased,0) = 0 and OD.SaleTypeID > 0 then 'n/a' else replace(dbo.fn_FormatCurrency(isnull(OD.AmtFinanced ,0) * isnull(OD.LeaseFactor ,0)),'$','') end + '</OR.BASELEASEPAYMENTNOSIGN>'"
                     mstrSQL = mstrSQL & " + '<OR.BIDDESKLABEL>' + case when isnull(OD.BidDeskDTReqSRepReview,'1/1/1900') <> '1/1/1900' then case when isnull(OD.BidDeskDTSRepReviewed,'1/1/1900') = '1/1/1900' then 'Requested Rep Review (' + convert(varchar,OD.BidDeskDTReqSRepReview,1) + ' ' + right(convert(varchar,OD.BidDeskDTReqSRepReview,100),7) + ')' else case upper(OD.BidDeskSRepStatus) when 'A' then 'Accepted By Rep (' + convert(varchar,OD.BidDeskDTSRepReviewed,1) + ' ' + right(convert(varchar,OD.BidDeskDTSRepReviewed,100),7) + ')' when 'R' then 'Rejected By Rep (' + convert(varchar,OD.BidDeskDTSRepReviewed,1) + ' ' + right(convert(varchar,OD.BidDeskDTSRepReviewed,100),7) + ')' else 'Unrecognized Status (' + OD.BidDeskSRepStatus + ')' end end else '' end + '</OR.BIDDESKLABEL>'"
                     mstrSQL = mstrSQL & " + '<OR.BILLINGCUSTOMERNAME>' + dbo.fn_FormatXMLChars(isnull(OD.BillingCustomerName ,'')) + '</OR.BILLINGCUSTOMERNAME>'"
                     mstrSQL = mstrSQL & " + '<OR.BILLINGOMDCUSTOMERNBR>' + dbo.fn_FormatXMLChars(isnull(BCU.OMDCustomerNbr ,'')) + '</OR.BILLINGOMDCUSTOMERNBR>'"
                     mstrSQL = mstrSQL & " + '<OR.BILLINGCUSTOMERASSOCIATION>' + case when (select isnull(ct.TypeInd,'') from SNCompany co with (nolock) left join SNCompanyType ct with (nolock) on ct.CompanyTypeID = co.CompanyTypeID where co.CompanyID = BCU.CompanyID) = 'A' then 'X' else '' end + '</OR.BILLINGCUSTOMERASSOCIATION>'"
                     mstrSQL = mstrSQL & " + '<OR.BILLINGCUSTOMERCORPORATION>' + case when (select isnull(ct.TypeInd,'') from SNCompany co with (nolock) left join SNCompanyType ct with (nolock) on ct.CompanyTypeID = co.CompanyTypeID where co.CompanyID = BCU.CompanyID) = 'C' then 'X' else '' end + '</OR.BILLINGCUSTOMERCORPORATION>'"
                     mstrSQL = mstrSQL & " + '<OR.BILLINGCUSTOMERLLC>' + case when (select isnull(ct.TypeInd,'') from SNCompany co with (nolock) left join SNCompanyType ct with (nolock) on ct.CompanyTypeID = co.CompanyTypeID where co.CompanyID = BCU.CompanyID) = 'L' then 'X' else '' end + '</OR.BILLINGCUSTOMERLLC>'"
                     mstrSQL = mstrSQL & " + '<OR.BILLINGCUSTOMERMUNICIPALITY>' + case when (select isnull(ct.TypeInd,'') from SNCompany co with (nolock) left join SNCompanyType ct with (nolock) on ct.CompanyTypeID = co.CompanyTypeID where co.CompanyID = BCU.CompanyID) = 'M' then 'X' else '' end + '</OR.BILLINGCUSTOMERMUNICIPALITY>'"
                     mstrSQL = mstrSQL & " + '<OR.BILLINGCUSTOMERPARTNERSHIP>' + case when (select isnull(ct.TypeInd,'') from SNCompany co with (nolock) left join SNCompanyType ct with (nolock) on ct.CompanyTypeID = co.CompanyTypeID where co.CompanyID = BCU.CompanyID) = 'T' then 'X' else '' end + '</OR.BILLINGCUSTOMERPARTNERSHIP>'"
                     mstrSQL = mstrSQL & " + '<OR.BILLINGCUSTOMERPROPRIETORSHIP>' + case when (select isnull(ct.TypeInd,'') from SNCompany co with (nolock) left join SNCompanyType ct with (nolock) on ct.CompanyTypeID = co.CompanyTypeID where co.CompanyID = BCU.CompanyID) = 'P' then 'X' else '' end + '</OR.BILLINGCUSTOMERPROPRIETORSHIP>'"
                     mstrSQL = mstrSQL & " + '<OR.BILLINGCUSTOMERTYPE>' + dbo.fn_FormatXMLChars(isnull((select Name from SNCompanyType with (nolock) where CompanyTypeID = (select CompanyTypeID from SNCompany with (nolock) where CompanyID = BCU.CompanyID)),'')) + '</OR.BILLINGCUSTOMERTYPE>'"
                     mstrSQL = mstrSQL & " + '<OR.BILLINGCUSTOMERFEDID>' + dbo.fn_FormatXMLChars(isnull((select FedIDNbr from SNCompany with (nolock) where CompanyID = BCU.CompanyID),'')) + '</OR.BILLINGCUSTOMERFEDID>'"
                     mstrSQL = mstrSQL & " + '<OR.BILLINGADDRESS1>' + dbo.fn_FormatXMLChars(isnull(OD.BillingAddress1 ,'')) + '</OR.BILLINGADDRESS1>'"
                     mstrSQL = mstrSQL & " + '<OR.BILLINGADDRESS2>' + dbo.fn_FormatXMLChars(isnull(OD.BillingAddress2 ,'')) + '</OR.BILLINGADDRESS2>'"
                     mstrSQL = mstrSQL & " + '<OR.BILLINGCITY>' + dbo.fn_FormatXMLChars(isnull(OD.BillingCity ,'')) + '</OR.BILLINGCITY>'"
                     mstrSQL = mstrSQL & " + '<OR.BILLINGSTATE>' + dbo.fn_FormatXMLChars(isnull(OD.BillingState ,'')) + '</OR.BILLINGSTATE>'"
                     mstrSQL = mstrSQL & " + '<OR.BILLINGZIP>' + dbo.fn_FormatXMLChars(isnull(OD.BillingPostalCode ,'')) + '</OR.BILLINGZIP>'"
                     mstrSQL = mstrSQL & " + '<OR.BILLINGADDRESSLABEL>' + dbo.fn_FormatXMLChars(case when len(OD.BillingCity) > 0 then OD.BillingAddress1 + case when LEN(OD.BillingAddress2) > 0 then CHAR(13) + CHAR(10) + OD.BillingAddress2 else '' end  + CHAR(13) + CHAR(10) + OD.BillingCity + ', ' + OD.BillingState + ' ' + OD.BillingPostalCode else '' end) + '</OR.BILLINGADDRESSLABEL>'"
                     mstrSQL = mstrSQL & " + '<OR.BILLINGADDRESSONELINE>' + dbo.fn_FormatXMLChars(case when len(OD.BillingCity) > 0 then OD.BillingAddress1 + case when LEN(OD.BillingAddress2) > 0 then OD.BillingAddress2 else '' end  + ', ' + OD.BillingCity + ', ' + OD.BillingState + ' ' + OD.BillingPostalCode else '' end) + '</OR.BILLINGADDRESSONELINE>'"
                     mstrSQL = mstrSQL & " + '<OR.BILLINGCONTACTFIRSTNAME>' + case when isnull(OD.BillToContactSelect,-1) > 0 then dbo.fn_FormatXMLChars(isnull((select p.FirstName from SNContact c with (nolock) join SNPerson p with (nolock) on p.PersonID = c.PersonID where c.ContactID = OD.BillToContactSelect),'')) else dbo.fn_FormatXMLChars(isnull(OD.BillingContactName ,'')) end + '</OR.BILLINGCONTACTFIRSTNAME>'"
                     mstrSQL = mstrSQL & " + '<OR.BILLINGCONTACTLASTNAME>' + case when isnull(OD.BillToContactSelect,-1) > 0 then dbo.fn_FormatXMLChars(isnull((select p.LastName from SNContact c with (nolock) join SNPerson p with (nolock) on p.PersonID = c.PersonID where c.ContactID = OD.BillToContactSelect),'')) else '' end + '</OR.BILLINGCONTACTLASTNAME>'"
                     mstrSQL = mstrSQL & " + '<OR.BILLINGCONTACTNAME>' + dbo.fn_FormatXMLChars(isnull(OD.BillingContactName ,'')) + '</OR.BILLINGCONTACTNAME>'"
                     mstrSQL = mstrSQL & " + '<OR.BILLINGCONTACTTITLE>' + dbo.fn_FormatXMLChars(isnull(case when OD.BillToContactSelect > 0 then (SELECT Title FROM SNContact with (nolock) WHERE ContactID = OD.BillToContactSelect) else '' end,'')) + '</OR.BILLINGCONTACTTITLE>'"
                     mstrSQL = mstrSQL & " + '<OR.BILLINGCONTACTEMAILADDR>' + dbo.fn_FormatXMLChars(isnull(OD.BillingContactEmailAddr ,'')) + '</OR.BILLINGCONTACTEMAILADDR>'"
                     mstrSQL = mstrSQL & " + '<OR.BILLINGCONTACTPHONENBR>' + dbo.fn_FormatPhoneNumber(OD.BillingContactPhoneNbr) + case when isnull(OD.BillingContactPhoneExt,'') <> '' then 'x' + isnull(OD.BillingContactPhoneExt,'') else '' end + '</OR.BILLINGCONTACTPHONENBR>'"
                     mstrSQL = mstrSQL & " + '<OR.BILLINGCONTACTFAXNBR>' + dbo.fn_FormatPhoneNumber(OD.BillingContactFaxNbr) + '</OR.BILLINGCONTACTFAXNBR>'"
                     mstrSQL = mstrSQL & " + '<OR.BILLINGCONTACTCELLNBR>' + dbo.fn_FormatPhoneNumber(OD.BillingContactCellNbr) + '</OR.BILLINGCONTACTCELLNBR>'"
                     mstrSQL = mstrSQL & " + '<OR.BLENDESCTEXT>' + case when isnull(OD.SPM_BlendEscInd,0) = -1 then 'Service rates are fixed for the term of the contract' else '' end + '</OR.BLENDESCTEXT>'"
                     mstrSQL = mstrSQL & " + '<OR.BOTDESCRIPTION>' + dbo.fn_FormatXMLChars(isnull(OD.ITP_BOT_Description,'')) + '</OR.BOTDESCRIPTION>'"
                     mstrSQL = mstrSQL & " + '<OR.BOTDOWNPAYMENT>' + dbo.fn_FormatCurrency(isnull(OD.ITPBOTDownPayment,0)) + '</OR.BOTDOWNPAYMENT>'"
                     mstrSQL = mstrSQL & " + '<OR.BOTHOURLYRATE>' + dbo.fn_FormatCurrency(isnull(OD.ITP_BOT_HourlyRate,0)) + '</OR.BOTHOURLYRATE>'"
                     mstrSQL = mstrSQL & " + '<OR.BOTHOURS>' + case when isnull(OD.ITP_BOTHrs,-1) = -1 then '' else dbo.fn_FormatNumber(OD.ITP_BOTHrs,0) end + '</OR.BOTHOURS>'"
                     mstrSQL = mstrSQL & " + '<OR.BOTORIGINALRATE>' + dbo.fn_FormatCurrency(isnull((select MoRepCost from SNServicesItem with (nolock) where ServicesItemID = (select top 1 ITP_BlockOfTimeServicesItemID from SNSystemParm with (nolock))),0)) + '</OR.BOTORIGINALRATE>'"
                     mstrSQL = mstrSQL & " + '<OR.BOTPPAYMENT>' + dbo.fn_FormatCurrency(case when isnull(BTB.IsInLeaseInd,0) = -1 then isnull(OD.LeaseFactor ,0) * isnull(OD.ITP_BOT_TotalCharge,0) else isnull(OD.ITP_BOT_TotalCharge,0) end) + '</OR.BOTPPAYMENT>'"
                     mstrSQL = mstrSQL & " + '<OR.BOTPAYMENTTERM>' + dbo.fn_FormatXMLChars(isnull(BTB.Name,'')) + '</OR.BOTPAYMENTTERM>'"
                     mstrSQL = mstrSQL & " + '<OR.BOTTOTALCHARGE>' + dbo.fn_FormatCurrency(isnull(OD.ITP_BOT_TotalCharge,0)) + '</OR.BOTTOTALCHARGE>'"
                     mstrSQL = mstrSQL & " + '<OR.BRANCH>' + dbo.fn_FormatXMLChars(isnull(BR.BranchName ,'')) + '</OR.BRANCH>'"
                     mstrSQL = mstrSQL & " + '<OR.BRANCHABBREV>' + dbo.fn_FormatXMLChars(isnull(BR.Abbreviation ,'')) + '</OR.BRANCHABBREV>'"
                     mstrSQL = mstrSQL & " + '<OR.BRANCHLEGALNAME>' + dbo.fn_FormatXMLChars(isnull(BR.LegalName ,'')) + '</OR.BRANCHLEGALNAME>'"
                     mstrSQL = mstrSQL & " + '<OR.BRANCHADDRESS>' + dbo.fn_FormatXMLChars(case when len(BR.City) > 0 then BR.Addr1 + case when LEN(BR.Addr2) > 0 then CHAR(10) + CHAR(13) + BR.Addr2 else '' end  + CHAR(10) + CHAR(13) + BR.City + ', ' + BR.State + ' ' + BR.PostalCode else '' end) + '</OR.BRANCHADDRESS>'"
                     mstrSQL = mstrSQL & " + '<OR.BRANCHADDR1>' + dbo.fn_FormatXMLChars(isnull(BR.Addr1 ,'')) + '</OR.BRANCHADDR1>'"
                     mstrSQL = mstrSQL & " + '<OR.BRANCHADDR2>' + dbo.fn_FormatXMLChars(isnull(BR.Addr2 ,'')) + '</OR.BRANCHADDR2>'"
                     mstrSQL = mstrSQL & " + '<OR.BRANCHCITY>' + dbo.fn_FormatXMLChars(isnull(BR.City ,'')) + '</OR.BRANCHCITY>'"
                     mstrSQL = mstrSQL & " + '<OR.BRANCHSTATE>' + dbo.fn_FormatXMLChars(isnull(BR.State ,'')) + '</OR.BRANCHSTATE>'"
                     mstrSQL = mstrSQL & " + '<OR.BRANCHPOSTALCODE>' + dbo.fn_FormatXMLChars(isnull(BR.PostalCode ,'')) + '</OR.BRANCHPOSTALCODE>'"
                     mstrSQL = mstrSQL & " + '<OR.BRANCHPHONE>' + dbo.fn_FormatPhoneNumber(BR.Phone) + '</OR.BRANCHPHONE>'"
                     mstrSQL = mstrSQL & " + '<OR.BRANCHFAX>' + dbo.fn_FormatPhoneNumber(BR.Fax) + '</OR.BRANCHFAX>'"
                     mstrSQL = mstrSQL & " + '<OR.BRANCHIMAGELARGE>' + dbo.fn_FormatXMLChars(isnull(BR.ImageLarge,'')) + '</OR.BRANCHIMAGELARGE>'"
                     mstrSQL = mstrSQL & " + '<OR.BRANCHIMAGESMALL>' + dbo.fn_FormatXMLChars(isnull(BR.ImageSmall,'')) + '</OR.BRANCHIMAGESMALL>'"
                     mstrSQL = mstrSQL & " + '<OR.BREAKDOWNAMTFUNDED>' + dbo.fn_FormatCurrency(OD.BreakdownAmtFunded) + '</OR.BREAKDOWNAMTFUNDED>'"
                     mstrSQL = mstrSQL & " + '<OR.BREAKDOWNAMTINVOICED>' + dbo.fn_FormatCurrency(OD.BreakdownAmtInvoiced) + '</OR.BREAKDOWNAMTINVOICED>'"
                     mstrSQL = mstrSQL & " + '<OR.BREAKDOWNBOOKINGTOTAL>' + dbo.fn_FormatCurrency(OD.BreakdownBookingTotal) + '</OR.BREAKDOWNBOOKINGTOTAL>'"
                     mstrSQL = mstrSQL & " + '<OR.BREAKDOWNEQTOTAL>' + dbo.fn_FormatCurrency(OD.EquipmentTotal) + '</OR.BREAKDOWNEQTOTAL>'"
                     mstrSQL = mstrSQL & " + '<OR.BREAKDOWNFUNDINGFINANCED>' + dbo.fn_FormatCurrency(OD.PaperCost) + '</OR.BREAKDOWNFUNDINGFINANCED>'"
                     mstrSQL = mstrSQL & " + '<OR.BREAKDOWNFUNDEDRATE>' + case when isnull(OD.PaperRate ,0) = 0 then '' else dbo.fn_formatnumber(OD.PaperRate,5) end + '</OR.BREAKDOWNFUNDEDRATE>'"
                     mstrSQL = mstrSQL & " + '<OR.BREAKDOWNHOUSE>' + dbo.fn_FormatCurrency(OD.BreakdownHouse) + '</OR.BREAKDOWNHOUSE>'"
                     mstrSQL = mstrSQL & " + '<OR.BUSINESS>' + dbo.fn_FormatXMLChars(isnull(BU.Name ,'')) + '</OR.BUSINESS>'"
                     mstrSQL = mstrSQL & " + '<OR.BUSINESSADDR1>' + dbo.fn_FormatXMLChars(isnull(BU.Address1 ,'')) + '</OR.BUSINESSADDR1>'"
                     mstrSQL = mstrSQL & " + '<OR.BUSINESSADDR2>' + dbo.fn_FormatXMLChars(isnull(BU.Address2 ,'')) + '</OR.BUSINESSADDR2>'"
                     mstrSQL = mstrSQL & " + '<OR.BUSINESSCITY>' + dbo.fn_FormatXMLChars(isnull(BU.City ,'')) + '</OR.BUSINESSCITY>'"
                     mstrSQL = mstrSQL & " + '<OR.BUSINESSSTATE>' + dbo.fn_FormatXMLChars(isnull(BU.StateCode ,'')) + '</OR.BUSINESSSTATE>'"
                     mstrSQL = mstrSQL & " + '<OR.BUSINESSPOSTALCODE>' + dbo.fn_FormatXMLChars(isnull(BU.PostalCode ,'')) + '</OR.BUSINESSPOSTALCODE>'"
                     mstrSQL = mstrSQL & " + '<OR.BUSINESSPHONE>' + dbo.fn_FormatPhoneNumber(BU.Phone) + '</OR.BUSINESSPHONE>'"
                     mstrSQL = mstrSQL & " + '<OR.BUSINESSFAX>' + dbo.fn_FormatPhoneNumber(BU.PhoneFax) + '</OR.BUSINESSFAX>'"
                     mstrSQL = mstrSQL & " + '<OR.BUSINESSIMAGELARGE>' + dbo.fn_FormatXMLChars(isnull(BU.ImageLarge,'')) + '</OR.BUSINESSIMAGELARGE>'"
                     mstrSQL = mstrSQL & " + '<OR.BUSINESSIMAGESMALL>' + dbo.fn_FormatXMLChars(isnull(BU.ImageSmall,'')) + '</OR.BUSINESSIMAGESMALL>'"
                     mstrSQL = mstrSQL & " + '<OR.BUYTOTAL>' + dbo.fn_FormatCurrency((SELECT SUM(ISNULL(BuyPrice,0) * Quantity) FROM SNOrderLine with (nolock) WHERE ISNULL(ServiceInd,0) = 0 AND OrderID = OD.OrderID)) + '</OR.BUYTOTAL>'"
                     mstrSQL = mstrSQL & " + '<OR.BUYTOTALPCT>' + dbo.fn_FormatCurrency(((SELECT SUM(ISNULL(BuyPrice,0) * Quantity) FROM SNOrderLine with (nolock) WHERE ISNULL(ServiceInd,0) = 0 AND OrderID = OD.OrderID) + OD.CSMPPCTAmt)) + '</OR.BUYTOTALPCT>'"
                     mstrSQL = mstrSQL & " + '<OR.COMMTOTAL>' + dbo.fn_FormatCurrency((SELECT SUM(ISNULL(EarnedCommAmt,0)) FROM SNSRCommEarned with (nolock) WHERE OrderID = OD.OrderID)) + '</OR.COMMTOTAL>'"
                     mstrSQL = mstrSQL & " + '<OR.COMMPRIMARYREPTOTAL>' + dbo.fn_FormatCurrency((SELECT SUM(ISNULL(EarnedCommAmt,0)) FROM SNSRCommEarned with (nolock) WHERE OrderID = OD.OrderID AND UserID = OD.SalesRepUserID)) + '</OR.COMMPRIMARYREPTOTAL>'"
                     'mstrSQL = mstrSQL & " + '<OR.CREDITAPPFCO>' + dbo.fn_FormatXMLChars(isnull((select co.Name from SNCreditApplication ca with (nolock) join SNFinanceCompany fc with (nolock) on fc.FCOID = ca.FCOID join SNCompany co with (nolock) on co.CompanyID = fc.CompanyID where ca.OrderID = OD.OrderID and ca.SelectedSource = -1),'')) + '</OR.CREDITAPPFCO>'"
                     mstrSQL = mstrSQL & " + '<OR.CREDITAPPFCO>' + dbo.fn_FormatXMLChars(replace(isnull(CFC.RPTName,''),'_',' ')) + '</OR.CREDITAPPFCO>'"
                     mstrSQL = mstrSQL & " + '<OR.CREDITAPPFCOADDRESSONELINE>' + dbo.fn_FormatXMLChars(case when len(isnull(CFC.RPTCity,'')) > 0 then isnull(CFC.RPTAddress1,'') + case when LEN(isnull(CFC.RPTAddress2,'')) > 0 then ' ' + isnull(CFC.RPTAddress2,'') else '' end  + ', ' +  isnull(CFC.RPTCity,'') + ', ' + isnull(CFC.RPTState,'') + ' ' + isnull(CFC.RPTZip,'') else '' end) + '</OR.CREDITAPPFCOADDRESSONELINE>'"
                     mstrSQL = mstrSQL & " + '<OR.CREDITAPPFCOADDRESS1>' + dbo.fn_FormatXMLChars(isnull(CFC.RPTAddress1,'')) + '</OR.CREDITAPPFCOADDRESS1>'"
                     mstrSQL = mstrSQL & " + '<OR.CREDITAPPFCOADDRESS2>' + dbo.fn_FormatXMLChars(isnull(CFC.RPTAddress2,'')) + '</OR.CREDITAPPFCOADDRESS2>'"
                     mstrSQL = mstrSQL & " + '<OR.CREDITAPPFCOCITY>' + dbo.fn_FormatXMLChars(isnull(CFC.RPTCity,'')) + '</OR.CREDITAPPFCOCITY>'"
                     mstrSQL = mstrSQL & " + '<OR.CREDITAPPFCOSTATE>' + dbo.fn_FormatXMLChars(isnull(CFC.RPTState,'')) + '</OR.CREDITAPPFCOSTATE>'"
                     mstrSQL = mstrSQL & " + '<OR.CREDITAPPFCOZIP>' + dbo.fn_FormatXMLChars(isnull(CFC.RPTZip,'')) + '</OR.CREDITAPPFCOZIP>'"
                     mstrSQL = mstrSQL & " + '<OR.CSMPAMTALL>' + dbo.fn_FormatCurrency((OD.SRCommCSMPCreditValue + OD.CSMPPCTAmt)) + '</OR.CSMPAMTALL>'"
                     mstrSQL = mstrSQL & " + '<OR.CSMPCONTRACT>' + dbo.fn_FormatXMLChars(isnull(OD.SRCommCSMPContractNbr,'')) + '</OR.CSMPCONTRACT>'"
                     mstrSQL = mstrSQL & " + '<OR.CSMPCREDITAMT>' + dbo.fn_FormatCurrency(OD.SRCommCSMPCreditValue) + '</OR.CSMPCREDITAMT>'"
                     mstrSQL = mstrSQL & " + '<OR.CSMPCREDITNAME>' + case when OD.SRCommCSMPLevelID = -100 then 'Custom' else case when  OD.SRCommCSMPLevelID = -102 then 'Strategic' else dbo.fn_FormatXMLChars(isnull(CSMP.LevelName,'')) end end  + '</OR.CSMPCREDITNAME>'"
                     mstrSQL = mstrSQL & " + '<OR.CURRENTMONTHLYCOST>' + case when isnull(OD.CurrentMonthlyCost,0) = 0 then '' else dbo.fn_FormatCurrency(isnull(OD.CurrentMonthlyCost,0)) end + '</OR.CURRENTMONTHLYCOST>'"
                     mstrSQL = mstrSQL & " + '<OR.CUSTOMERPO>' + dbo.fn_FormatXMLChars(isnull(OD.CustomerPO,'')) + '</OR.CUSTOMERPO>'"
                     mstrSQL = mstrSQL & " + '<OR.CUSTOMERUPGRADEX>' + case when isnull(OD.CustomerUpgradeind,0) = -1 then 'X' else '' end + '</OR.CUSTOMERUPGRADEX>'"
                     mstrSQL = mstrSQL & " + '<OR.DEAR>' + dbo.fn_FormatXMLChars(isnull(OD.Dear,'')) + '</OR.DEAR>'"
                     mstrSQL = mstrSQL & " + '<OR.DELIVERYFEEELEVATORX>' + case when isnull(OD.ElevatorInd,0) = -1 then 'X' else '' end + '</OR.DELIVERYFEEELEVATORX>'"
                     mstrSQL = mstrSQL & " + '<OR.DELIVERYFEENBRFLOORS>' + dbo.fn_formatnumber(OD.DeliveryNbrFloors,0) + '</OR.DELIVERYFEENBRFLOORS>'"
                     mstrSQL = mstrSQL & " + '<OR.DELIVERYMETHOD>' + dbo.fn_FormatXMLChars(isnull((select stuff((select '; ' + isnull(Method,'') from SNDeliveryJob with (nolock) where OrderID = OD.OrderID order by DeliveryJobCount for xml path('')),1,2,'')),'')) + '</OR.DELIVERYMETHOD>'"
                     mstrSQL = mstrSQL & " + '<OR.DMACCESSORYPRICE>' + dbo.fn_FormatCurrency(isnull(OD.DMTotalAddFeesAmt,0)) + '</OR.DMACCESSORYPRICE>'"
                     mstrSQL = mstrSQL & " + '<OR.DMDOWNPAYMENT>' + dbo.fn_FormatCurrency(isnull(OD.DMDownPaymentAmt,0)) + '</OR.DMDOWNPAYMENT>'"
                     mstrSQL = mstrSQL & " + '<OR.DMDOWNPAYMENTDT>' + dbo.fn_FormatDate(OD.DTClosed,3,0,0) + '</OR.DMDOWNPAYMENTDT>'"
                     mstrSQL = mstrSQL & " + '<OR.DMPAYMENTTERM>' + dbo.fn_FormatXMLChars(isnull(DMB.Name,'')) + '</OR.DMPAYMENTTERM>'"
                     mstrSQL = mstrSQL & " + '<OR.DMPRIMARYPRICE>' + dbo.fn_FormatCurrency(isnull(OD.DMTotalServicesAmt,0)) + '</OR.DMPRIMARYPRICE>'"
                     mstrSQL = mstrSQL & " + '<OR.DMSALESTAX>' + dbo.fn_FormatCurrency(isnull(OD.DMTotalSalesTaxAmt,0)) + '</OR.DMSALESTAX>'"
                     mstrSQL = mstrSQL & " + '<OR.DMSUBTOTAL>' + dbo.fn_FormatCurrency(isnull(OD.DMTotalOrderAmt,0)) + '</OR.DMSUBTOTAL>'"
                     mstrSQL = mstrSQL & " + '<OR.DMSUMSERVICES>' + dbo.fn_FormatCurrency(isnull((select sum(case when s.Type = 'P' then l.TotalSellAmt else 0 end) from SNOrderDMLine l with (nolock) join SNDMServices s with (nolock) on s.DMServiceID = l.DMServiceID where l.OrderID = OD.OrderID),0)) + '</OR.DMSUMSERVICES>'"
                     mstrSQL = mstrSQL & " + '<OR.DMSUMFEES>' + dbo.fn_FormatCurrency(isnull((select sum(case when s.Type <> 'P' then l.TotalSellAmt else 0 end) from SNOrderDMLine l with (nolock) join SNDMServices s with (nolock) on s.DMServiceID = l.DMServiceID where l.OrderID = OD.OrderID),0)) + '</OR.DMSUMFEES>'"
                     mstrSQL = mstrSQL & " + '<OR.DMTOTAL>' + dbo.fn_FormatCurrency(isnull(OD.DMTotalSalesTaxAmt,0) + isnull(OD.DMTotalOrderAmt,0)) + '</OR.DMTOTAL>'"
                     mstrSQL = mstrSQL & " + '<OR.DOCFEE>' + case when isnull(ST.IsLeased,0) = 0 and OD.SaleTypeID > 0 then 'n/a' else dbo.fn_FormatCurrency(isnull(LP.DocFee,0))  end + '</OR.DOCFEE>'"
                     mstrSQL = mstrSQL & " + '<OR.DTCDAY>' + cast(datepart(dd,OD.DTClosed) as nvarchar) + '</OR.DTCDAY>'"
                     mstrSQL = mstrSQL & " + '<OR.DTCMONTH>' + cast(datepart(mm,OD.DTClosed) as nvarchar) + '</OR.DTCMONTH>'"
                     mstrSQL = mstrSQL & " + '<OR.DTCMONTHNAME>' + datename(mm,OD.DTClosed) + '</OR.DTCMONTHNAME>'"
                     mstrSQL = mstrSQL & " + '<OR.DTCLOSED>' + dbo.fn_FormatDate(OD.DTClosed,3,0,0)  + '</OR.DTCLOSED>'"
                     mstrSQL = mstrSQL & " + '<OR.DTCYEAR>' + cast(datepart(yy,OD.DTClosed) as nvarchar) + '</OR.DTCYEAR>'"
                     mstrSQL = mstrSQL & " + '<OR.DTPROPOSAL>' + dbo.fn_FormatDate(PR.DTProposal,3,0,0)  + '</OR.DTPROPOSAL>'"
                     mstrSQL = mstrSQL & " + '<OR.EOTFIXPCTX>' + case when isnull(LP.isEOT_FixedPCT,0) = -1 then 'X' else '' end + '</OR.EOTFIXPCTX>'"
                     mstrSQL = mstrSQL & " + '<OR.EOTFMVX>' + case when isnull(LP.isEOT_FMV,0) = -1 then 'X' else '' end + '</OR.EOTFMVX>'"
                     mstrSQL = mstrSQL & " + '<OR.EOTOPTION>' + case when isnull(LP.isEOT_FMV,0) = -1 then 'FMV' else case when isnull(LP.isEOT_BuckOut,0) = -1 then '$Out' else case when isnull(LP.isEOT_FixedPCT,0) = -1 then 'Fixed PCT' else case when isnull(LP.isEOT_NoOption,0) = -1 then 'No Option' else '' end end end end + '</OR.EOTOPTION>'"
                     mstrSQL = mstrSQL & " + '<OR.EOTOUTX>' + case when isnull(LP.isEOT_BuckOut,0) = -1 then 'X' else '' end + '</OR.EOTOUTX>'"
                     mstrSQL = mstrSQL & " + '<OR.EXPCLOSEDT>' + dbo.fn_FormatDate(PR.DTExpectedClose,3,0,0)  + '</OR.EXPCLOSEDT>'"
                     mstrSQL = mstrSQL & " + '<OR.FCOAPPAPPROVEDATE>' + dbo.fn_FormatDate(OD.FCOApprovedDT,3,0,0)  + '</OR.FCOAPPAPPROVEDATE>'"
                     mstrSQL = mstrSQL & " + '<OR.FCOAPPNBR>' + dbo.fn_FormatXMLChars(isnull(OD.FCOAppNbr ,'')) + '</OR.FCOAPPNBR>'"
                     mstrSQL = mstrSQL & " + '<OR.FCOAPPRVALIDTHRUDATE>' + dbo.fn_FormatDate(OD.FCOApprovalValidThruDT,3,0,0)  + '</OR.FCOAPPRVALIDTHRUDATE>'"
                     mstrSQL = mstrSQL & " + '<OR.FCOAPPLDATE>' + dbo.fn_FormatDate(OD.FCOApplicationDT,3,0,0)  + '</OR.FCOAPPLDATE>'"
                     mstrSQL = mstrSQL & " + '<OR.FCOAPPRCONDITIONS>' + dbo.fn_FormatXMLChars(cast(isnull(OD.FCOConditions ,'') as nvarchar(max))) + '</OR.FCOAPPRCONDITIONS>'"
                     mstrSQL = mstrSQL & " + '<OR.FCOADDRESS1>' + dbo.fn_FormatXMLChars(isnull(FC.RPTAddress1,'')) + '</OR.FCOADDRESS1>'"
                     mstrSQL = mstrSQL & " + '<OR.FCOADDRESS2>' + dbo.fn_FormatXMLChars(isnull(FC.RPTAddress2,'')) + '</OR.FCOADDRESS2>'"
                     mstrSQL = mstrSQL & " + '<OR.FCOCITY>' + dbo.fn_FormatXMLChars(isnull(FC.RPTCity,'')) + '</OR.FCOCITY>'"
                     mstrSQL = mstrSQL & " + '<OR.FCOSTATE>' + dbo.fn_FormatXMLChars(isnull(FC.RPTState,'')) + '</OR.FCOSTATE>'"
                     mstrSQL = mstrSQL & " + '<OR.FCOZIP>' + dbo.fn_FormatXMLChars(isnull(FC.RPTZip,'')) + '</OR.FCOZIP>'"
                     mstrSQL = mstrSQL & " + '<OR.FCOPHONENBR>' + dbo.fn_FormatPhoneNumber(FC.AppPhoneNbr) + '</OR.FCOPHONENBR>'"
                     mstrSQL = mstrSQL & " + '<OR.FCOLEGALNAME>' + dbo.fn_FormatXMLChars(case when isnull(OD.FCOLegalName,'') = '' then case when isnull(OD.BillingCustomerName,'') = '' then  isnull(BCU.RPTName ,'') else OD.BillingCustomerName end else OD.FCOLegalName end) + '</OR.FCOLEGALNAME>'"
                     mstrSQL = mstrSQL & " + '<OR.FCOLEGALNAMEASIS>' + dbo.fn_FormatXMLChars(isnull(OD.FCOLegalName ,'')) + '</OR.FCOLEGALNAMEASIS>'"
                     mstrSQL = mstrSQL & " + '<OR.FCONAME>' + dbo.fn_FormatXMLChars(replace(isnull(CO.Name,''),'_',' ')) + '</OR.FCONAME>'"
                     mstrSQL = mstrSQL & " + '<OR.GENDAY>' + cast(datepart(dd,getdate()) as nvarchar) + '</OR.GENDAY>'"
                     mstrSQL = mstrSQL & " + '<OR.GENMONTH>' + cast(datepart(mm,getdate()) as nvarchar) + '</OR.GENMONTH>'"
                     mstrSQL = mstrSQL & " + '<OR.GENMONTHNAME>' + datename(mm,getdate()) + '</OR.GENMONTHNAME>'"
                     mstrSQL = mstrSQL & " + '<OR.GENYEAR>' + cast(datepart(yy,getdate()) as nvarchar) + '</OR.GENYEAR>'"
                     mstrSQL = mstrSQL & " + '<OR.HASCUSTOMEREQX>' + case when exists(select * from SNOrderLine with (nolock) where OrderID = OD.OrderID and OwnershipInd = 'C') then 'X' else '' end + '</OR.HASCUSTOMEREQX>'"
                     mstrSQL = mstrSQL & " + '<OR.HASDEALEREQX>' + case when exists(select * from SNOrderLine with (nolock) where OrderID = OD.OrderID and isnull(OwnershipInd,'D') = 'D' and BPProductCategoryID not in (select ProductCategoryID from SNProductCategory with (nolock) where isnull(IsNotMachine,0) = -1)) then 'X' else '' end + '</OR.HASDEALEREQX>'"
                     mstrSQL = mstrSQL & " + '<OR.HASEXISTINGEQLABEL>' + case when exists(select * from SNOrderLine with (nolock) where OrderID = OD.OrderID and isnull(OwnershipInd,'') = 'D' and BPProductCategoryID not in (select ProductCategoryID from SNProductCategory with (nolock) where isnull(IsNotMachine,0) = -1)) then 'Existing EQ' else '' end + '</OR.HASEXISTINGEQLABEL>'"
                     mstrSQL = mstrSQL & " + '<OR.HASEXISTINGEQX>' + case when exists(select * from SNOrderLine with (nolock) where OrderID = OD.OrderID and isnull(OwnershipInd,'') = 'D' and BPProductCategoryID not in (select ProductCategoryID from SNProductCategory with (nolock) where isnull(IsNotMachine,0) = -1)) then 'X' else '' end + '</OR.HASEXISTINGEQX>'"
                     mstrSQL = mstrSQL & " + '<OR.HASNEWEQX>' + case when exists(select * from SNOrderLine with (nolock) where OrderID = OD.OrderID and isnull(OwnershipInd,'') = '' and isnull(ServiceInd,0) = 0) then 'X' else '' end + '</OR.HASNEWEQX>'"
                     mstrSQL = mstrSQL & " + '<OR.HASSERVICEX>' + case when exists(select * from SNOBDServiceLine with (nolock) where OBDID = OD.OrderID) then 'X' else '' end + '</OR.HASSERVICEX>'"
                     mstrSQL = mstrSQL & " + '<OR.HASSERVICEXPTONLY>' + case when isnull(BT.IncludeInLease,0) = -1 then case when exists(select * from SNOBDServiceLine with (nolock) where OBDID = OD.OrderID) then 'X' else '' end else '' end + '</OR.HASSERVICEXPTONLY>'"
                     mstrSQL = mstrSQL & " + '<OR.HASSOFTWAREX>' + case when exists(select * from SNOrderLine with (nolock) where OrderID = OD.OrderID and BPProductCategoryID in (select ProductCategoryID from SNProductCategory with (nolock) where isnull(IsNotMachine,0) = -1)) then 'X' else '' end + '</OR.HASSOFTWAREX>'"
                     mstrSQL = mstrSQL & " + '<OR.HASUSEDEQX>' + case when exists(select * from SNOrderLine with (nolock) where OrderID = OD.OrderID and isnull(Condition,'') = 'U') then 'X' else '' end + '</OR.HASUSEDEQX>'"
                     mstrSQL = mstrSQL & " + '<OR.IMMEDIATEMANAGERNAME>' +  dbo.fn_FormatXMLChars(isnull((select top 1 isnull(pr.FirstName,'') + ' ' + case when isnull(pr.MiddleName,'') <> '' then pr.MiddleName + ' ' else '' end + isnull(pr.LastName,'') from SNUser u with (nolock) join SNHierarchyNode p with (nolock) on p.ObjectKey = u.UserID join SNHierarchyNode c with (nolock) on c.ParentHierarchyNodeID = p.HierarchyNodeID join SNPerson pr with (nolock) on pr.PersonID = u.PersonID where c.ObjectKey = OD.SalesRepUserID),isnull(SMP.FirstName,'') + ' ' + case when isnull(SMP.MiddleName,'') <> '' then  SMP.MiddleName + ' ' else '' end + isnull(SMP.LastName,''))) + '</OR.IMMEDIATEMANAGERNAME>'"
                     mstrSQL = mstrSQL & " + '<OR.ITPDEALERCOST>' + dbo.fn_FormatCurrency(isnull(OD.ITPTotalDealerCostAmt,0)) + '</OR.ITPDEALERCOST>'"
                     mstrSQL = mstrSQL & " + '<OR.ITPDOWNPAYMENT>' + dbo.fn_FormatCurrency(isnull(OD.ITPDownPaymentAmt,0)) + '</OR.ITPDOWNPAYMENT>'"
                     mstrSQL = mstrSQL & " + '<OR.ITPDOWNPAYMENTDT>' + dbo.fn_FormatDate(OD.DTClosed,3,0,0) + '</OR.ITPDOWNPAYMENTDT>'"
                     mstrSQL = mstrSQL & " + '<OR.ITPGP>' + dbo.fn_FormatCurrency(isnull(OD.ITPTotalGP,0)) + '</OR.ITPGP>'"
                     mstrSQL = mstrSQL & " + '<OR.ITPPAYMENT>' + dbo.fn_FormatCurrency(case when isnull(IPB.IsInLeaseInd,0) = -1 then isnull(OD.LeaseFactor ,0) * isnull(OD.ITPTotalSellAmt,0) else isnull(OD.ITPTotalSellAmt,0) end) + '</OR.ITPPAYMENT>'"
                     mstrSQL = mstrSQL & " + '<OR.ITPPAYMENTTERM>' + dbo.fn_FormatXMLChars(isnull(IPB.Name,'')) + '</OR.ITPPAYMENTTERM>'"
                     mstrSQL = mstrSQL & " + '<OR.ITPTOTALGPPCT>' + dbo.fn_FormatNumber(OD.ITPTotalGPPCT * 100,2) + '%' + '</OR.ITPTOTALGPPCT>'"
                     mstrSQL = mstrSQL & " + '<OR.ITPREPCOST>' + dbo.fn_FormatCurrency(isnull(OD.ITPTotalRepCostAmt,0)) + '</OR.ITPREPCOST>'"
                     mstrSQL = mstrSQL & " + '<OR.ITPSELLCOST>' + dbo.fn_FormatCurrency(isnull(OD.ITPTotalSellAmt,0)) + '</OR.ITPSELLCOST>'"
                     mstrSQL = mstrSQL & " + '<OR.LEASEACHX>' + case when isnull(OD.LeaseACHInd,0) = -1 then 'X' else '' end + '</OR.LEASEACHX>'"
                     mstrSQL = mstrSQL & " + '<OR.LEASEACHXNO>' + case when isnull(OD.LeaseACHInd,0) = -1 then '' else 'X' end + '</OR.LEASEACHXNO>'"
                     mstrSQL = mstrSQL & " + '<OR.LEASEACHBANKACCTNBR>' + dbo.fn_FormatXMLChars(isnull(OD.LeaseACHBankAcctNbr ,'')) + '</OR.LEASEACHBANKACCTNBR>'"
                     mstrSQL = mstrSQL & " + '<OR.LEASEACHBANKROUTENBR>' + dbo.fn_FormatXMLChars(isnull(OD.LeaseACHBankRouteNbr ,'')) + '</OR.LEASEACHBANKROUTENBR>'"
                     mstrSQL = mstrSQL & " + '<OR.LEASEADJUSTTYPES>' + dbo.fn_FormatXMLChars(isnull(stuff((select ', ' + Name from SNLeaseAdjustType with (nolock) where ',' + isnull(OD.ActiveLeaseAdjustTypeIDs,'') + ',' like '%,' + cast(LeaseAdjustTypeID as nvarchar) + ',%' for xml path('')),1,2,''),'')) + '</OR.LEASEADJUSTTYPES>'"
                     mstrSQL = mstrSQL & " + '<OR.LEASEBASEFACTOR>' + case when isnull(ST.IsLeased,0) = 0 and OD.SaleTypeID > 0 then 'n/a' else isnull(nullif(dbo.fn_formatnumber((select RateFactor from SNLeasePrice with (nolock) where LeasePriceID = OD.LeaseFactorSelect),5),'0.00000'),'') end + '</OR.LEASEBASEFACTOR>'"
                     mstrSQL = mstrSQL & " + '<OR.LEASECOMMENCEDT>' + dbo.fn_FormatDate(OD.LeaseCommenceDT,3,0,0) + '</OR.LEASECOMMENCEDT>'"
                     mstrSQL = mstrSQL & " + '<OR.LEASEDCAX>' + case when isnull(OD.isDCAInd,0) = -1 then 'X' else '' end + '</OR.LEASEDCAX>'"
                     mstrSQL = mstrSQL & " + '<OR.LEASEDCAXNO>' + case when isnull(OD.isDCAInd,0) = -1 then '' else 'X' end + '</OR.LEASEDCAXNO>'"
                     mstrSQL = mstrSQL & " + '<OR.LEASEFACTOR>' + case when isnull(ST.IsLeased,0) = 0 and OD.SaleTypeID > 0 then 'n/a' else case when isnull(OD.LeaseFactor ,0) = 0 then '' else cast(OD.LeaseFactor as nvarchar) end end + '</OR.LEASEFACTOR>'"
                     mstrSQL = mstrSQL & " + '<OR.LEASEFIRSTPYMTDUEDT>' + dbo.fn_FormatDate(OD.LeaseFirstPymtDueDT,3,0,0) + '</OR.LEASEFIRSTPYMTDUEDT>'"
                     mstrSQL = mstrSQL & " + '<OR.LEASEFIRSTUSAGEDT>' + dbo.fn_FormatDate(OD.LeaseFirstUsageDT,3,0,0) + '</OR.LEASEFIRSTUSAGEDT>'"
                     mstrSQL = mstrSQL & " + '<OR.LEASEFIRSTUSAGEPERIOD>' + substring(dbo.fn_FormatDate(OD.LeaseFirstUsageDT,3,0,0),1,5) + ' - ' + substring(dbo.fn_FormatDate(OD.LeaseFirstPymtDueDT,3,0,0),1,5) + '</OR.LEASEFIRSTUSAGEPERIOD>'"
                     mstrSQL = mstrSQL & " + '<OR.LEASEINTERIMDAYS>' + dbo.fn_formatnumber(isnull(OD.NbrInterimRentDays,0),0) + '</OR.LEASEINTERIMDAYS>'"
                     mstrSQL = mstrSQL & " + '<OR.LEASEINTERIMEQPERDIEM>' + case when isnull(ST.IsLeased,0) = 0 and OD.SaleTypeID > 0 then 'n/a' else dbo.fn_FormatCurrency(isnull(round((isnull(OD.AmtFinanced,0) * isnull(OD.LeaseFactor,0)) / 30, 2),0))  end + '</OR.LEASEINTERIMEQPERDIEM>'"
                     mstrSQL = mstrSQL & " + '<OR.LEASEINTERIMEQ25RENT>' + case when isnull(ST.IsLeased,0) = 0 and OD.SaleTypeID > 0 then 'n/a' else dbo.fn_FormatCurrency((isnull(round((isnull(OD.AmtFinanced,0) * isnull(OD.LeaseFactor,0)) / 30, 2),0) * isnull(OD.NbrInterimRentDays,0)) * .25)  end + '</OR.LEASEINTERIMEQ25RENT>'"
                     mstrSQL = mstrSQL & " + '<OR.LEASEINTERIMEQRENT>' + case when isnull(ST.IsLeased,0) = 0 and OD.SaleTypeID > 0 then 'n/a' else dbo.fn_FormatCurrency(isnull(round((isnull(OD.AmtFinanced,0) * isnull(OD.LeaseFactor,0)) / 30, 2),0) * isnull(OD.NbrInterimRentDays,0))  end + '</OR.LEASEINTERIMEQRENT>'"
                     mstrSQL = mstrSQL & " + '<OR.LEASEINTERIMPERDIEM>' + case when isnull(ST.IsLeased,0) = 0 and OD.SaleTypeID > 0 then 'n/a' else dbo.fn_FormatCurrency(isnull(round(isnull(OD.PaymentAmtMonthlyFin,0) / 30, 2),0))  end + '</OR.LEASEINTERIMPERDIEM>'"
                     mstrSQL = mstrSQL & " + '<OR.LEASEINTERIMRENT>' + case when isnull(ST.IsLeased,0) = 0 and OD.SaleTypeID > 0 then 'n/a' else dbo.fn_FormatCurrency((isnull(round((isnull(OD.AmtFinanced,0) * isnull(OD.LeaseFactor,0)) / 30, 2),0) + isnull(round((isnull(OD.PaymentAmtMonthlyFin,0) / 30) - ((isnull(OD.AmtFinanced,0) * isnull(OD.LeaseFactor,0)) / 30), 2),0)) * isnull(OD.NbrInterimRentDays,0))  end + '</OR.LEASEINTERIMRENT>'"
                     mstrSQL = mstrSQL & " + '<OR.LEASEINTERIMRENTLESS25>' + case when isnull(ST.IsLeased,0) = 0 and OD.SaleTypeID > 0 then 'n/a' else dbo.fn_FormatCurrency(((isnull(round((isnull(OD.AmtFinanced,0) * isnull(OD.LeaseFactor,0)) / 30, 2),0) + isnull(round((isnull(OD.PaymentAmtMonthlyFin,0) / 30) - ((isnull(OD.AmtFinanced,0) * isnull(OD.LeaseFactor,0)) / 30), 2),0)) * isnull(OD.NbrInterimRentDays,0)) - round(((isnull(round((isnull(OD.AmtFinanced,0) * isnull(OD.LeaseFactor,0)) / 30, 2),0) * isnull(OD.NbrInterimRentDays,0)) * .25),2))  end + '</OR.LEASEINTERIMRENTLESS25>'"
                     mstrSQL = mstrSQL & " + '<OR.LEASEINTERIMSVCPERDIEM>' + case when isnull(ST.IsLeased,0) = 0 and OD.SaleTypeID > 0 then 'n/a' else dbo.fn_FormatCurrency(isnull(round((isnull(OD.PaymentAmtMonthlyFin,0) / 30) - ((isnull(OD.AmtFinanced,0) * isnull(OD.LeaseFactor,0)) / 30), 2),0))  end + '</OR.LEASEINTERIMSVCPERDIEM>'"
                     mstrSQL = mstrSQL & " + '<OR.LEASEINTERIMSVCRENT>' + case when isnull(ST.IsLeased,0) = 0 and OD.SaleTypeID > 0 then 'n/a' else dbo.fn_FormatCurrency(isnull(round((isnull(OD.PaymentAmtMonthlyFin,0) / 30) - ((isnull(OD.AmtFinanced,0) * isnull(OD.LeaseFactor,0)) / 30), 2),0) * isnull(OD.NbrInterimRentDays,0))  end + '</OR.LEASEINTERIMSVCRENT>'"
                     mstrSQL = mstrSQL & " + '<OR.LEASEINSTALLDT>' + dbo.fn_FormatDate(OD.InstallDT,3,0,0) + '</OR.LEASEINSTALLDT>'"
                     mstrSQL = mstrSQL & " + '<OR.LEASEMASTERAGREEMENTNBR>' + dbo.fn_FormatXMLChars(isnull(OD.LeaseMasterAgreementNbr ,'')) + '</OR.LEASEMASTERAGREEMENTNBR>'"
                     mstrSQL = mstrSQL & " + '<OR.LEASEPRICELEVEL>' + case when isnull(ST.IsLeased,0) = 0 and OD.SaleTypeID > 0 then 'n/a' else dbo.fn_FormatXMLChars(isnull(LPL.Name,'')) end + '</OR.LEASEPRICELEVEL>'"
                     mstrSQL = mstrSQL & " + '<OR.LEASEPRODUCT>' + dbo.fn_FormatXMLChars(isnull(LP.Name ,'')) + '</OR.LEASEPRODUCT>'"
                     mstrSQL = mstrSQL & " + '<OR.LEASEPYMTINCLUDES>' + dbo.fn_FormatXMLChars(case when isnull(SBT.FundedInLeaseInd,0) = -1 or isnull(IPB.IsInLeaseInd,0) = -1 or isnull(BTB.IsInLeaseInd,0) = -1 or isnull(DMB.IsInLeaseInd,0) = -1 then '* Payment includes ' + case when isnull(SBT.FundedInLeaseInd,0) = -1 then 'IT Services' else '' end + case when isnull(IPB.IsInLeaseInd,0) = -1 then case when isnull(SBT.FundedInLeaseInd,0) = -1 then ' and ' else '' end + 'IT Products' else '' end + case when isnull(BTB.IsInLeaseInd,0) = -1 then case when isnull(SBT.FundedInLeaseInd,0) = -1 or isnull(IPB.IsInLeaseInd,0) = -1 then ' and ' else '' end + 'Block of Time' else '' end + case when isnull(DMB.IsInLeaseInd,0) = -1 then case when isnull(SBT.FundedInLeaseInd,0) = -1 or isnull(IPB.IsInLeaseInd,0) = -1 or isnull(BTB.IsInLeaseInd,0) = -1 then ' and ' else '' end + 'Document Management' else '' end else '' end) + '</OR.LEASEPYMTINCLUDES>'"
                     mstrSQL = mstrSQL & " + '<OR.LEASESCHEDULENBR>' + dbo.fn_FormatXMLChars(isnull(OD.LeaseScheduleNbr ,'')) + '</OR.LEASESCHEDULENBR>'"
                     mstrSQL = mstrSQL & " + '<OR.LEASESERVICEPAYMENT>' + case when isnull(ST.IsLeased,0) = 0 and OD.SaleTypeID > 0 then 'n/a' else dbo.fn_FormatCurrency(case when BT.IncludeInLease = -1 then isnull(OD.MaintMonthValueActual,0) else 0 end)  end + '</OR.LEASESERVICEPAYMENT>'"
                     mstrSQL = mstrSQL & " + '<OR.LEASESECURITYDEPAMT>' + dbo.fn_FormatCurrency(isnull(OD.LeaseSecurityDepAmt,0)) + '</OR.LEASESECURITYDEPAMT>'"
                     mstrSQL = mstrSQL & " + '<OR.LEASETAXEXEMPTX>' + case when isnull(OD.isTaxExemptInd,0) = -1 then 'X' else '' end + '</OR.LEASETAXEXEMPTX>'"
                     mstrSQL = mstrSQL & " + '<OR.LEASETAXEXEMPTXNO>' + case when isnull(OD.isTaxExemptInd,0) = -1 then '' else 'X' end + '</OR.LEASETAXEXEMPTXNO>'"
                     mstrSQL = mstrSQL & " + '<OR.LEASETAXEXEMPTYESNO>' + case when isnull(OD.isTaxExemptInd,0) = -1 then 'Yes' else 'No' end + '</OR.LEASETAXEXEMPTYESNO>'"
                     mstrSQL = mstrSQL & " + '<OR.MAINTBASEPAYMENT>' + dbo.fn_FormatCurrency(OD.MaintBasePayment) + '</OR.MAINTBASEPAYMENT>'"
                     mstrSQL = mstrSQL & " + '<OR.MAINTBASEPAYMENTY2>' + case when isnull(OD.SPM_BlendEscInd,0) = -1 then '' else dbo.fn_FormatCurrency(OD.MaintBasePayment * (OD.SPM_Y2EscalatePCT / 100)) end + '</OR.MAINTBASEPAYMENTY2>'"
                     mstrSQL = mstrSQL & " + '<OR.MAINTBASEPAYMENTY3>' + case when isnull(OD.SPM_BlendEscInd,0) = -1 then '' else dbo.fn_FormatCurrency((OD.MaintBasePayment * (OD.SPM_Y2EscalatePCT / 100)) * (OD.SPM_Y3EscalatePCT / 100)) end + '</OR.MAINTBASEPAYMENTY3>'"
                     mstrSQL = mstrSQL & " + '<OR.MAINTBASEPAYMENTY4>' + case when isnull(OD.SPM_BlendEscInd,0) = -1 then '' else dbo.fn_FormatCurrency(((OD.MaintBasePayment * (OD.SPM_Y2EscalatePCT / 100)) * (OD.SPM_Y3EscalatePCT / 100)) * (OD.SPM_Y4EscalatePCT / 100)) end + '</OR.MAINTBASEPAYMENTY4>'"
                     mstrSQL = mstrSQL & " + '<OR.MAINTBASEPAYMENTY5>' + case when isnull(OD.SPM_BlendEscInd,0) = -1 then '' else dbo.fn_FormatCurrency((((OD.MaintBasePayment * (OD.SPM_Y2EscalatePCT / 100)) * (OD.SPM_Y3EscalatePCT / 100)) * (OD.SPM_Y4EscalatePCT / 100)) * (OD.SPM_Y5EscalatePCT / 100)) end + '</OR.MAINTBASEPAYMENTY5>'"
                     mstrSQL = mstrSQL & " + '<OR.MAINTBASEPAYMENTY6>' + case when isnull(OD.SPM_BlendEscInd,0) = -1 then '' else dbo.fn_FormatCurrency(((((OD.MaintBasePayment * (OD.SPM_Y2EscalatePCT / 100)) * (OD.SPM_Y3EscalatePCT / 100)) * (OD.SPM_Y4EscalatePCT / 100)) * (OD.SPM_Y5EscalatePCT / 100)) * (OD.SPM_Y6EscalatePCT / 100)) end + '</OR.MAINTBASEPAYMENTY6>'"
                     mstrSQL = mstrSQL & " + '<OR.MAINTBASEPAYMENTY2PROJECTED>' + dbo.fn_FormatCurrency(OD.MaintBasePayment * (CT.Y2EscalatePCT / 100)) + '</OR.MAINTBASEPAYMENTY2PROJECTED>'"
                     mstrSQL = mstrSQL & " + '<OR.MAINTBASEPAYMENTY3PROJECTED>' + dbo.fn_FormatCurrency((OD.MaintBasePayment * (CT.Y2EscalatePCT / 100)) * (CT.Y3EscalatePCT / 100)) + '</OR.MAINTBASEPAYMENTY3PROJECTED>'"
                     mstrSQL = mstrSQL & " + '<OR.MAINTBASEPAYMENTY4PROJECTED>' + dbo.fn_FormatCurrency(((OD.MaintBasePayment * (CT.Y2EscalatePCT / 100)) * (CT.Y3EscalatePCT / 100)) * (CT.Y4EscalatePCT / 100)) + '</OR.MAINTBASEPAYMENTY4PROJECTED>'"
                     mstrSQL = mstrSQL & " + '<OR.MAINTBASEPAYMENTY5PROJECTED>' + dbo.fn_FormatCurrency((((OD.MaintBasePayment * (CT.Y2EscalatePCT / 100)) * (CT.Y3EscalatePCT / 100)) * (CT.Y4EscalatePCT / 100)) * (CT.Y5EscalatePCT / 100)) + '</OR.MAINTBASEPAYMENTY5PROJECTED>'"
                     mstrSQL = mstrSQL & " + '<OR.MAINTBASEPAYMENTY6PROJECTED>' + dbo.fn_FormatCurrency(((((OD.MaintBasePayment * (CT.Y2EscalatePCT / 100)) * (CT.Y3EscalatePCT / 100)) * (CT.Y4EscalatePCT / 100)) * (CT.Y5EscalatePCT / 100)) * (CT.Y6EscalatePCT / 100)) + '</OR.MAINTBASEPAYMENTY6PROJECTED>'"
                     mstrSQL = mstrSQL & " + '<OR.MAINTBILLINGCYCLE>' + dbo.fn_FormatXMLChars(isnull(BT.Name,'')) + '</OR.MAINTBILLINGCYCLE>'"
                     mstrSQL = mstrSQL & " + '<OR.MAINTBILLINGCYCLELEASE>' + dbo.fn_FormatXMLChars(case when isnull(BT.IncludeInLease,0) = -1 then case LP.PaymentFrequency when 1 then 'Annually' when 2 then 'Semi-Annually' when 4 then 'Quarterly' else 'Monthly' end else isnull(BT.Name,'') end) + '</OR.MAINTBILLINGCYCLELEASE>'"
                     mstrSQL = mstrSQL & " + '<OR.MAINTBILLINGCYCLEOTHER>' + dbo.fn_FormatXMLChars(isnull(BT.OtherName,'')) + '</OR.MAINTBILLINGCYCLEOTHER>'"
                     mstrSQL = mstrSQL & " + '<OR.MAINTBILLINGCYCLEOTHERCAPS>' + dbo.fn_FormatXMLChars(upper(isnull(BT.OtherName,''))) + '</OR.MAINTBILLINGCYCLEOTHERCAPS>'"
                     mstrSQL = mstrSQL & " + '<OR.MAINTCONTRACTDESC>' + dbo.fn_FormatXMLChars(isnull(CT.Description,'')) + '</OR.MAINTCONTRACTDESC>'"
                     mstrSQL = mstrSQL & " + '<OR.MAINTCONTRACTLEGALTEXT>' + dbo.fn_FormatXMLChars(isnull(CT.LegalText,'')) + '</OR.MAINTCONTRACTLEGALTEXT>'"
                     mstrSQL = mstrSQL & " + '<OR.MAINTCONTRACTTYPE>' + dbo.fn_FormatXMLChars(isnull(CT.Name,'')) + '</OR.MAINTCONTRACTTYPE>'"
                     mstrSQL = mstrSQL & " + '<OR.MAINTENDDATEV2>' + (case when OD.MaintStartDateV2 > 2 then dbo.fn_FormatDate(DATEADD(M,OD.MaintNumMonths,DATEADD(D,OD.MaintStartDateV2,'12/30/1899')),3,0,0) else '' end)  + '</OR.MAINTENDDATEV2>'"
                     mstrSQL = mstrSQL & " + '<OR.MAINTFIXEDANNUAL>' + dbo.fn_FormatCurrency((select sum(BundleQuantity * FixedAmount) from SNOrderLine with (nolock) where OrderID = OD.OrderID)) + '</OR.MAINTFIXEDANNUAL>'"
                     mstrSQL = mstrSQL & " + '<OR.MAINTNUMMONTHS>' + cast(isnull(OD.MaintNumMonths ,'') as nvarchar) + '</OR.MAINTNUMMONTHS>'"
                     mstrSQL = mstrSQL & " + '<OR.MAINTOVERBILLINGCYCLE>' + dbo.fn_FormatXMLChars(case when OD.MaintOverBillingCycle = -1 then isnull(BT.Name,'') else isnull(OBT.Name,'') end) + '</OR.MAINTOVERBILLINGCYCLE>'"
                     mstrSQL = mstrSQL & " + '<OR.MAINTOVERBILLINGCYCLEOTHER>' + dbo.fn_FormatXMLChars(case when OD.MaintOverBillingCycle = -1 then isnull(BT.OtherName,'') else isnull(OBT.OtherName,'') end) + '</OR.MAINTOVERBILLINGCYCLEOTHER>'"
                     mstrSQL = mstrSQL & " + '<OR.MAINTOVERBILLINGCYCLEOTHERCAPS>' + dbo.fn_FormatXMLChars(upper(case when OD.MaintOverBillingCycle = -1 then isnull(BT.OtherName,'') else isnull(OBT.OtherName,'') end)) + '</OR.MAINTOVERBILLINGCYCLEOTHERCAPS>'"
                     mstrSQL = mstrSQL & " + '<OR.MAINTSTARTDATE>' + dbo.fn_FormatDate(OD.MaintStartDate,3,0,0) + '</OR.MAINTSTARTDATE>'"
                     mstrSQL = mstrSQL & " + '<OR.MAINTSTARTDATEV2>' + (case when OD.MaintStartDateV2 > 1 then dbo.fn_FormatDate(DATEADD(D,OD.MaintStartDateV2,'12/30/1899'),3,0,0) else '' end)  + '</OR.MAINTSTARTDATEV2>'"
                     mstrSQL = mstrSQL & " + '<OR.MAINTSTARTDATEV2METRO>' + (case when OD.MaintStartDateV2 > 1 then dbo.fn_FormatDate(DATEADD(D,OD.MaintStartDateV2,'12/30/1899'),3,0,0) else 'Install Date' end)  + '</OR.MAINTSTARTDATEV2METRO>'"
                     mstrSQL = mstrSQL & " + '<OR.MAINTSTARTDATEY2>' + case when isnull(OD.SPM_BlendEscInd,0) = -1 then '' else dbo.fn_FormatDate(dateadd(yy,1,OD.MaintStartDate),3,0,0) end + '</OR.MAINTSTARTDATEY2>'"
                     mstrSQL = mstrSQL & " + '<OR.MAINTSTARTDATEY3>' + case when isnull(OD.SPM_BlendEscInd,0) = -1 then '' else dbo.fn_FormatDate(dateadd(yy,2,OD.MaintStartDate),3,0,0) end + '</OR.MAINTSTARTDATEY3>'"
                     mstrSQL = mstrSQL & " + '<OR.MAINTSTARTDATEY4>' + case when isnull(OD.SPM_BlendEscInd,0) = -1 then '' else dbo.fn_FormatDate(dateadd(yy,3,OD.MaintStartDate),3,0,0) end + '</OR.MAINTSTARTDATEY4>'"
                     mstrSQL = mstrSQL & " + '<OR.MAINTSTARTDATEY5>' + case when isnull(OD.SPM_BlendEscInd,0) = -1 then '' else dbo.fn_FormatDate(dateadd(yy,4,OD.MaintStartDate),3,0,0) end + '</OR.MAINTSTARTDATEY5>'"
                     mstrSQL = mstrSQL & " + '<OR.MAINTSTARTDATEY6>' + case when isnull(OD.SPM_BlendEscInd,0) = -1 then '' else dbo.fn_FormatDate(dateadd(yy,5,OD.MaintStartDate),3,0,0) end + '</OR.MAINTSTARTDATEY6>'"
                     mstrSQL = mstrSQL & " + '<OR.MAINTSTARTDATEY2PROJECTED>' + dbo.fn_FormatDate(dateadd(yy,1,OD.MaintStartDate),3,0,0) + '</OR.MAINTSTARTDATEY2PROJECTED>'"
                     mstrSQL = mstrSQL & " + '<OR.MAINTSTARTDATEY3PROJECTED>' + dbo.fn_FormatDate(dateadd(yy,2,OD.MaintStartDate),3,0,0) + '</OR.MAINTSTARTDATEY3PROJECTED>'"
                     mstrSQL = mstrSQL & " + '<OR.MAINTSTARTDATEY4PROJECTED>' + dbo.fn_FormatDate(dateadd(yy,3,OD.MaintStartDate),3,0,0) + '</OR.MAINTSTARTDATEY4PROJECTED>'"
                     mstrSQL = mstrSQL & " + '<OR.MAINTSTARTDATEY5PROJECTED>' + dbo.fn_FormatDate(dateadd(yy,4,OD.MaintStartDate),3,0,0) + '</OR.MAINTSTARTDATEY5PROJECTED>'"
                     mstrSQL = mstrSQL & " + '<OR.MAINTSTARTDATEY6PROJECTED>' + dbo.fn_FormatDate(dateadd(yy,5,OD.MaintStartDate),3,0,0) + '</OR.MAINTSTARTDATEY6PROJECTED>'"
                     mstrSQL = mstrSQL & " + '<OR.MAINTTOTALBASEPAYMENT>' + dbo.fn_FormatCurrency(OD.MaintBasePayment * (12/isnull(BT.PaymentsPerYear,12))) + '</OR.MAINTTOTALBASEPAYMENT>'"
                     mstrSQL = mstrSQL & " + '<OR.MGRAPPROVEDLABEL>' + case when isnull(OD.ApprovedByUserID,-1) < 0 then '' else case isnull(OD.ManagerApproved,0) when 1 then 'Approved By Manager' when 2 then 'Declined By Manager' else 'No Action By Manager' end + '(' + dbo.fn_FormatXMLChars((select isnull(cPR.FirstName,'') + ' ' + case when isnull(cPR.MiddleName,'') <> '' then cPR.MiddleName + ' ' else '' end + isnull(cPR.LastName,'') from SNUser cUS with (nolock) join SNPerson cPR with (nolock) on cPR.PersonID = cUS.PersonID where cUS.UserID = OD.ApprovedByUserID)) + ')' end + '</OR.MGRAPPROVEDLABEL>'"
                     mstrSQL = mstrSQL & " + '<OR.MNSINDYESNO>' + case when isnull(OD.IsMNSInd,0) = -1 then 'Yes' else 'No' end + '</OR.MNSINDYESNO>'"
                     mstrSQL = mstrSQL & " + '<OR.MNSINDYESX>' + case when isnull(OD.IsMNSInd,0) = -1 then 'X' else '' end + '</OR.MNSINDYESX>'"
                     mstrSQL = mstrSQL & " + '<OR.MPSINDYESNO>' + case when isnull(OD.IsMPSInd,0) = -1 then 'Yes' else 'No' end + '</OR.MPSINDYESNO>'"
                     mstrSQL = mstrSQL & " + '<OR.MPSINDYESX>' + case when isnull(OD.IsMPSInd,0) = -1 then 'X' else '' end + '</OR.MPSINDYESX>'"
                     mstrSQL = mstrSQL & " + '<OR.NBRLEASEPYMTS>' + case when ST.IsLeased = 0 and OD.SaleTypeID > 0 then 'n/a' else case when isnull(OD.NbrLeasePymts,0) = 0 then '' else cast(OD.NbrLeasePymts as nvarchar) end end + '</OR.NBRLEASEPYMTS>'"
                     mstrSQL = mstrSQL & " + '<OR.NBRLEASEPYMTSCAN>' + case when ST.IsLeased = 0 and OD.SaleTypeID > 0 then 'n/a' else case when isnull(OD.NbrLeasePymts,0) = 0 then '' else cast((isnull(LP.PaymentFrequency,12) * OD.NbrLeasePymts)/12 as nvarchar) end end + '</OR.NBRLEASEPYMTSCAN>'"
                     mstrSQL = mstrSQL & " + '<OR.NBRPYMTSNBRDEF>' + case when isnull(OD.NbrLeasePymts + OD.NbrDeferredPymts,0) = 0 then '' else cast((OD.NbrLeasePymts + OD.NbrDeferredPymts) as nvarchar) end + '</OR.NBRPYMTSNBRDEF>'"
                     mstrSQL = mstrSQL & " + '<OR.NBRPYMTSTBLW>' + case when ST.IsLeased = 0 and OD.SaleTypeID > 0 then 'n/a' else case when isnull(OD.NbrDeferredPymts,0) = 0 then case when isnull(OD.LeaseOverride,0) = -1 then dbo.fn_FormatNumber(isnull(OD.OverrideLeaseTerm,0),0) else case when isnull(OD.PaymentNbrFin,0) = 0 then '' else dbo.fn_FormatNumber(isnull(OD.NbrLeasePymts,0),0) end end else dbo.fn_FormatNumber(isnull(OD.NbrDeferredPymts,0),0) + ' @' + char(13) + case when isnull(OD.LeaseOverride,0) = -1 then dbo.fn_FormatNumber(isnull(OD.OverrideLeaseTerm,0) - isnull(OD.NbrDeferredPymts,0),0) else dbo.fn_FormatNumber(isnull(OD.PaymentNbrFin,0) - isnull(OD.NbrDeferredPymts,0),0) end + ' @' end end + '</OR.NBRPYMTSTBLW>'"
                     mstrSQL = mstrSQL & " + '<OR.DSLSPYMTSUM>' "
                     mstrSQL = mstrSQL & " + case when isnull(ST.IsLeased,0) = 0 and OD.SaleTypeID > 0 "
                     mstrSQL = mstrSQL & "          then '' "
                     mstrSQL = mstrSQL & "        else case when isnull(OD.NbrDeferredPymts,0) = 0 and isnull(LP.isDefStep,0) = 0 "
                     mstrSQL = mstrSQL & "                    then case when OD.PaymentAmtMonthlyFin = 0 then '' "
                     mstrSQL = mstrSQL & "                              else case when isnull(BT.IncludeInLease,0) = -1 "
                     mstrSQL = mstrSQL & "                                          then dbo.fn_FormatCurrency(isnull(OD.AmtFinanced,0) * isnull(OD.LeaseFactor,0)) "
                     mstrSQL = mstrSQL & "                                        else dbo.fn_FormatCurrency(OD.PaymentAmtMonthlyFin) "
                     mstrSQL = mstrSQL & "                                   end "
                     mstrSQL = mstrSQL & "                         end "
                     mstrSQL = mstrSQL & "                  when isnull(OD.NbrDeferredPymts,0) <> 0 and isnull(LP.isDefStep,0) = 0  "
                     mstrSQL = mstrSQL & "                    then '$0.00' + char(13) + case when isnull(BT.IncludeInLease,0) = -1  "
                     mstrSQL = mstrSQL & "                                                     then dbo.fn_FormatCurrency(isnull(OD.AmtFinanced,0) * isnull(OD.LeaseFactor,0)) "
                     mstrSQL = mstrSQL & "                                                   else dbo.fn_FormatCurrency(OD.PaymentAmtMonthlyFin) "
                     mstrSQL = mstrSQL & "                                              end "
                     mstrSQL = mstrSQL & "                  when isnull(OD.NbrDeferredPymts,0) = 0 and isnull(LP.isDefStep,0) <> 0 "
                     mstrSQL = mstrSQL & "                    then cast(LP.DS1NbrPymts as nvarchar) + ' @ ' +  dbo.fn_FormatCurrency(isnull(OD.AmtFinanced,0) * isnull(OD.LeaseFactor,0) * LP.DS1PCTPymt) "
                     mstrSQL = mstrSQL & "                         + case when LP.DS2NbrPymts > 0 "
                     mstrSQL = mstrSQL & "                                  then ' then ' + case when LP.DS3NbrPymts > 0 "
                     mstrSQL = mstrSQL & "                                                         then cast(LP.DS2NbrPymts as nvarchar) "
                     mstrSQL = mstrSQL & "                                                       else cast(OD.NbrLeasePymts - LP.DS1NbrPymts as nvarchar)"
                     mstrSQL = mstrSQL & "                                                  end + ' @ ' +  dbo.fn_FormatCurrency(isnull(OD.AmtFinanced,0) * isnull(OD.LeaseFactor,0) * LP.DS2PCTPymt) "
                     mstrSQL = mstrSQL & "                                else '' "
                     mstrSQL = mstrSQL & "                           end "
                     mstrSQL = mstrSQL & "                         + case when LP.DS3NbrPymts > 0 "
                     mstrSQL = mstrSQL & "                                  then ' then ' + cast(OD.NbrLeasePymts - (LP.DS1NbrPymts + LP.DS2NbrPymts) as nvarchar) + ' @ ' "
                     mstrSQL = mstrSQL & "                                       +  dbo.fn_FormatCurrency(isnull(OD.AmtFinanced,0) * isnull(OD.LeaseFactor,0) * LP.DS3PCTPymt) "
                     mstrSQL = mstrSQL & "                                else '' "
                     mstrSQL = mstrSQL & "                           end "
                     mstrSQL = mstrSQL & "                  else '' "
                     mstrSQL = mstrSQL & "             end "
                     mstrSQL = mstrSQL & "   end "
                     mstrSQL = mstrSQL & " + '</OR.DSLSPYMTSUM>'"
                     mstrSQL = mstrSQL & " + '<OR.NEWCUSTOMER>' + case when isnull(OD.NewCustomerInd,0) = -1 then 'Yes' else 'No' end + '</OR.NEWCUSTOMER>'"
                     mstrSQL = mstrSQL & " + '<OR.NEWCUSTOMERX>' + case when isnull(OD.NewCustomerInd,0) = -1 then 'X' else '' end + '</OR.NEWCUSTOMERX>'"
                     mstrSQL = mstrSQL & " + '<OR.NEWPLACEMENTX>' + case when isnull(OD.NewPlacementInd,0) = -1 then 'X' else '' end + '</OR.NEWPLACEMENTX>'"
                     mstrSQL = mstrSQL & " + '<OR.NOTE>' + dbo.fn_FormatXMLChars(ISNULL(cast(OD.Note as nvarchar(max)),'')) + '</OR.NOTE>'"
                     mstrSQL = mstrSQL & " + '<OR.NUMMACHINECFG>' + dbo.fn_formatnumber(isnull((select sum(Quantity) from SNOrderLine with (nolock) where OrderID = OD.OrderID and isPrimaryInd = -1),0),0) + '</OR.NUMMACHINECFG>'"
                     mstrSQL = mstrSQL & " + '<OR.ORDERID>' + cast(OD.OrderID as nvarchar) + '</OR.ORDERID>'"
                     mstrSQL = mstrSQL & " + '<OR.OVEBILLFREQANX>' + case when OD.MaintOverBillingCycle = -1 and BT.PaymentsPerYear = 1 then 'X' else case when OBT.PaymentsPerYear = 1 then 'X'  else '' end end + '</OR.OVEBILLFREQANX>'"
                     mstrSQL = mstrSQL & " + '<OR.OVEBILLFREQMOX>' + case when OD.MaintOverBillingCycle = -1 and BT.PaymentsPerYear = 12 then 'X' else case when OBT.PaymentsPerYear = 12 then 'X'  else '' end end + '</OR.OVEBILLFREQMOX>'"
                     mstrSQL = mstrSQL & " + '<OR.OVEBILLFREQQTRX>' + case when OD.MaintOverBillingCycle = -1 and BT.PaymentsPerYear = 4 then 'X' else case when OBT.PaymentsPerYear = 4 then 'X'  else '' end end + '</OR.OVEBILLFREQQTRX>'"
                     mstrSQL = mstrSQL & " + '<OR.OVEBILLFREQSAX>' + case when OD.MaintOverBillingCycle = -1 and BT.PaymentsPerYear = 2 then 'X' else case when OBT.PaymentsPerYear = 2 then 'X'  else '' end end + '</OR.OVEBILLFREQSAX>'"
                     mstrSQL = mstrSQL & " + '<OR.OVERRIDECHARGEBACK>' + dbo.fn_FormatCurrency(OD.MaintTotalChargeback) + '</OR.OVERRIDECHARGEBACK>'"
                     mstrSQL = mstrSQL & " + '<OR.PAYMENTAMTMONTHLYFIN>' + case when isnull(ST.IsLeased,0) = 0 and OD.SaleTypeID > 0 then 'n/a' else dbo.fn_FormatCurrency(OD.PaymentAmtMonthlyFin) end + '</OR.PAYMENTAMTMONTHLYFIN>'"
                     mstrSQL = mstrSQL & " + '<OR.PAYMENTAMTMONTHLYFINNOSIGN>' + case when isnull(ST.IsLeased,0) = 0 and OD.SaleTypeID > 0 then 'n/a' else replace(dbo.fn_FormatCurrency(OD.PaymentAmtMonthlyFin),'$','') end + '</OR.PAYMENTAMTMONTHLYFINNOSIGN>'"
                     mstrSQL = mstrSQL & " + '<OR.PAYMENTAMTMONTHLYFINY2>' + case when isnull(OD.SPM_BlendEscInd,0) = -1 then '' else case when isnull(ST.IsLeased,0) = 0 and OD.SaleTypeID > 0 then 'n/a' else dbo.fn_FormatCurrency((OD.AmtFinanced * OD.LeaseFactor) + (OD.MaintBasePayment * (OD.SPM_Y2EscalatePCT / 100))) end end + '</OR.PAYMENTAMTMONTHLYFINY2>'"
                     mstrSQL = mstrSQL & " + '<OR.PAYMENTAMTMONTHLYFINY3>' + case when isnull(OD.SPM_BlendEscInd,0) = -1 then '' else case when isnull(ST.IsLeased,0) = 0 and OD.SaleTypeID > 0 then 'n/a' else dbo.fn_FormatCurrency((OD.AmtFinanced * OD.LeaseFactor) + ((OD.MaintBasePayment * (OD.SPM_Y2EscalatePCT / 100)) * (OD.SPM_Y3EscalatePCT / 100))) end end + '</OR.PAYMENTAMTMONTHLYFINY3>'"
                     mstrSQL = mstrSQL & " + '<OR.PAYMENTAMTMONTHLYFINY4>' + case when isnull(OD.SPM_BlendEscInd,0) = -1 then '' else case when isnull(ST.IsLeased,0) = 0 and OD.SaleTypeID > 0 then 'n/a' else dbo.fn_FormatCurrency((OD.AmtFinanced * OD.LeaseFactor) + (((OD.MaintBasePayment * (OD.SPM_Y2EscalatePCT / 100)) * (OD.SPM_Y3EscalatePCT / 100)) * (OD.SPM_Y4EscalatePCT / 100))) end end + '</OR.PAYMENTAMTMONTHLYFINY4>'"
                     mstrSQL = mstrSQL & " + '<OR.PAYMENTAMTMONTHLYFINY5>' + case when isnull(OD.SPM_BlendEscInd,0) = -1 then '' else case when isnull(ST.IsLeased,0) = 0 and OD.SaleTypeID > 0 then 'n/a' else dbo.fn_FormatCurrency((OD.AmtFinanced * OD.LeaseFactor) + ((((OD.MaintBasePayment * (OD.SPM_Y2EscalatePCT / 100)) * (OD.SPM_Y3EscalatePCT / 100)) * (OD.SPM_Y4EscalatePCT / 100)) * (OD.SPM_Y5EscalatePCT / 100))) end end + '</OR.PAYMENTAMTMONTHLYFINY5>'"
                     mstrSQL = mstrSQL & " + '<OR.PAYMENTAMTMONTHLYFINY6>' + case when isnull(OD.SPM_BlendEscInd,0) = -1 then '' else case when isnull(ST.IsLeased,0) = 0 and OD.SaleTypeID > 0 then 'n/a' else dbo.fn_FormatCurrency((OD.AmtFinanced * OD.LeaseFactor) + (((((OD.MaintBasePayment * (OD.SPM_Y2EscalatePCT / 100)) * (OD.SPM_Y3EscalatePCT / 100)) * (OD.SPM_Y4EscalatePCT / 100)) * (OD.SPM_Y5EscalatePCT / 100)) * (OD.SPM_Y6EscalatePCT / 100))) end end + '</OR.PAYMENTAMTMONTHLYFINY6>'"
                     mstrSQL = mstrSQL & " + '<OR.PAYMENTAMTMONTHLYFINY2PROJECTED>' + case when isnull(ST.IsLeased,0) = 0 and OD.SaleTypeID > 0 then 'n/a' else dbo.fn_FormatCurrency((OD.AmtFinanced * OD.LeaseFactor) + (OD.MaintBasePayment * (CT.Y2EscalatePCT / 100))) end + '</OR.PAYMENTAMTMONTHLYFINY2PROJECTED>'"
                     mstrSQL = mstrSQL & " + '<OR.PAYMENTAMTMONTHLYFINY3PROJECTED>' + case when isnull(ST.IsLeased,0) = 0 and OD.SaleTypeID > 0 then 'n/a' else dbo.fn_FormatCurrency((OD.AmtFinanced * OD.LeaseFactor) + ((OD.MaintBasePayment * (CT.Y2EscalatePCT / 100)) * (CT.Y3EscalatePCT / 100))) end + '</OR.PAYMENTAMTMONTHLYFINY3PROJECTED>'"
                     mstrSQL = mstrSQL & " + '<OR.PAYMENTAMTMONTHLYFINY4PROJECTED>' + case when isnull(ST.IsLeased,0) = 0 and OD.SaleTypeID > 0 then 'n/a' else dbo.fn_FormatCurrency((OD.AmtFinanced * OD.LeaseFactor) + (((OD.MaintBasePayment * (CT.Y2EscalatePCT / 100)) * (CT.Y3EscalatePCT / 100)) * (CT.Y4EscalatePCT / 100))) end + '</OR.PAYMENTAMTMONTHLYFINY4PROJECTED>'"
                     mstrSQL = mstrSQL & " + '<OR.PAYMENTAMTMONTHLYFINY5PROJECTED>' + case when isnull(ST.IsLeased,0) = 0 and OD.SaleTypeID > 0 then 'n/a' else dbo.fn_FormatCurrency((OD.AmtFinanced * OD.LeaseFactor) + ((((OD.MaintBasePayment * (CT.Y2EscalatePCT / 100)) * (CT.Y3EscalatePCT / 100)) * (CT.Y4EscalatePCT / 100)) * (CT.Y5EscalatePCT / 100))) end + '</OR.PAYMENTAMTMONTHLYFINY5PROJECTED>'"
                     mstrSQL = mstrSQL & " + '<OR.PAYMENTAMTMONTHLYFINY6PROJECTED>' + case when isnull(ST.IsLeased,0) = 0 and OD.SaleTypeID > 0 then 'n/a' else dbo.fn_FormatCurrency((OD.AmtFinanced * OD.LeaseFactor) + (((((OD.MaintBasePayment * (CT.Y2EscalatePCT / 100)) * (CT.Y3EscalatePCT / 100)) * (CT.Y4EscalatePCT / 100)) * (CT.Y5EscalatePCT / 100)) * (CT.Y6EscalatePCT / 100))) end + '</OR.PAYMENTAMTMONTHLYFINY6PROJECTED>'"
                     mstrSQL = mstrSQL & " + '<OR.PAYMENTAMTMONTHLYFINWITHSERVICE>' + case when ST.IsLeased = 0 and OD.SaleTypeID > 0 then 'n/a' else case when (isnull(OD.AmtFinanced ,0) * isnull(OD.LeaseFactor ,0)) + OD.MaintBasePayment = 0 then '' else dbo.fn_FormatCurrency((isnull(OD.AmtFinanced ,0) * isnull(OD.LeaseFactor ,0)) + OD.MaintBasePayment) end end + '</OR.PAYMENTAMTMONTHLYFINWITHSERVICE>'"
                     mstrSQL = mstrSQL & " + '<OR.PAYMENTNBRFIN>' + case when isnull(OD.LeaseOverride,0) = -1 then cast(isnull(OD.OverrideLeaseTerm,0) as nvarchar) else case when isnull(OD.PaymentNbrFin ,0) = 0 then '' else cast(OD.PaymentNbrFin as nvarchar) end end + '</OR.PAYMENTNBRFIN>'"
                     mstrSQL = mstrSQL & " + '<OR.PAYMENTNBRFINMINUSONE>' + case when isnull(OD.LeaseOverride,0) = -1 then cast(isnull(OD.OverrideLeaseTerm - 1,0) as nvarchar) else case when isnull(OD.PaymentNbrFin ,0) = 0 then '' else cast(OD.PaymentNbrFin - 1 as nvarchar) end end + '</OR.PAYMENTNBRFINMINUSONE>'"
                     mstrSQL = mstrSQL & " + '<OR.PAYMENTWITHSERVICE>' + case when ST.IsLeased = 0 and OD.SaleTypeID > 0 then 'n/a' else case when BT.IncludeInLease = -1 then case when OD.PaymentAmtMonthlyFin + OD.MaintBasePayment = 0 then '' else dbo.fn_FormatCurrency(OD.PaymentAmtMonthlyFin + OD.MaintBasePayment) end else case when OD.PaymentAmtMonthlyFin = 0 then '' else dbo.fn_FormatCurrency(OD.PaymentAmtMonthlyFin) end end end + '</OR.PAYMENTWITHSERVICE>'"
                     mstrSQL = mstrSQL & " + '<OR.PAYMENTFUNDEDWITHSERVICE>' + case when ST.IsLeased = 0 and OD.SaleTypeID > 0 then 'n/a' else case when BT.FundedInLeaseInd = -1 then case when OD.PaymentAmtMonthlyFin + OD.MaintBasePayment = 0 then '' else dbo.fn_FormatCurrency(OD.PaymentAmtMonthlyFin + OD.MaintBasePayment) end else case when OD.PaymentAmtMonthlyFin = 0 then '' else dbo.fn_FormatCurrency(OD.PaymentAmtMonthlyFin) end end end + '</OR.PAYMENTFUNDEDWITHSERVICE>'"
                     mstrSQL = mstrSQL & " + '<OR.PBCHARGEBACK>' + dbo.fn_FormatCurrency(OD.PBChargeBackAmt) + '</OR.PBCHARGEBACK>'"
                     mstrSQL = mstrSQL & " + '<OR.PRICECONFORM>' + case when isnull(OD.PricingMethod,0) <> 0 then 'Pricing is Non Conforming - ' else '' end + '</OR.PRICECONFORM>'"
                     mstrSQL = mstrSQL & " + '<OR.PRICELEVEL>' + dbo.fn_FormatXMLChars(isnull(PL.Name ,'')) + '</OR.PRICELEVEL>'"
                     mstrSQL = mstrSQL & " + '<OR.PRICELEVELDESCRIPTION>' + dbo.fn_FormatXMLChars(isnull(PL.Description ,'')) + '</OR.PRICELEVELDESCRIPTION>'"
                     mstrSQL = mstrSQL & " + '<OR.PRICELEVELLEGALTEXT>' + dbo.fn_FormatXMLChars(isnull(PL.LegalText ,'')) + '</OR.PRICELEVELLEGALTEXT>'"
                     mstrSQL = mstrSQL & " + '<OR.PRICEMETHOD>' + case isnull(OD.PricingMethod,0) when 0 then 'Point Book' when 1 then 'Promo/Special' when 2 then 'Bid Desk' when 3 then 'Branch Manager' when 4 then 'Double Contest' else '' end + '</OR.PRICEMETHOD>'"
                     mstrSQL = mstrSQL & " + '<OR.PROPOSALNAME>' + dbo.fn_FormatXMLChars(isnull(OD.ProposalName,'')) + '</OR.PROPOSALNAME>'"
                     mstrSQL = mstrSQL & " + '<OR.PS24MOPYMT>' + case when isnull(ST.IsLeased,0) = 0 and OD.SaleTypeID > 0 then 'n/a' else dbo.fn_FormatCurrency(OD.PS24MoPymt) end + '</OR.PS24MOPYMT>'"
                     mstrSQL = mstrSQL & " + '<OR.PS36MOPYMT>' + case when isnull(ST.IsLeased,0) = 0 and OD.SaleTypeID > 0 then 'n/a' else dbo.fn_FormatCurrency(OD.PS36MoPymt) end + '</OR.PS36MOPYMT>'"
                     mstrSQL = mstrSQL & " + '<OR.PS48MOPYMT>' + case when isnull(ST.IsLeased,0) = 0 and OD.SaleTypeID > 0 then 'n/a' else dbo.fn_FormatCurrency(OD.PS48MoPymt) end + '</OR.PS48MOPYMT>'"
                     mstrSQL = mstrSQL & " + '<OR.PS60MOPYMT>' + case when isnull(ST.IsLeased,0) = 0 and OD.SaleTypeID > 0 then 'n/a' else dbo.fn_FormatCurrency(OD.PS60MoPymt) end + '</OR.PS60MOPYMT>'"
                     mstrSQL = mstrSQL & " + '<OR.PYMTFREQANX>' + case when LP.PaymentFrequency = 1 then 'X' else '' end + '</OR.PYMTFREQANX>'"
                     mstrSQL = mstrSQL & " + '<OR.PYMTFREQLONG>' + case when LP.PaymentFrequency = 12 then 'Monthly' else"
                     mstrSQL = mstrSQL & "                         case when LP.PaymentFrequency = 4 then 'Quarterly' else"
                     mstrSQL = mstrSQL & "                         case when LP.PaymentFrequency = 2 then 'Semi-Annual' else"
                     mstrSQL = mstrSQL & "                         case when LP.PaymentFrequency = 1 then 'Annual' else '??Unknown' end end end end + '</OR.PYMTFREQLONG>'"
                     mstrSQL = mstrSQL & " + '<OR.PYMTFREQMOX>' + case when LP.PaymentFrequency = 12 then 'X' else '' end + '</OR.PYMTFREQMOX>'"
                     mstrSQL = mstrSQL & " + '<OR.PYMTFREQQTRX>' + case when LP.PaymentFrequency = 4 then 'X' else '' end + '</OR.PYMTFREQQTRX>'"
                     mstrSQL = mstrSQL & " + '<OR.PYMTFREQSAX>' + case when LP.PaymentFrequency = 2 then 'X' else '' end + '</OR.PYMTFREQSAX>'"
                     mstrSQL = mstrSQL & " + '<OR.PYMTFREQSHORT>' + case when LP.PaymentFrequency = 12 then 'MO' else"
                     mstrSQL = mstrSQL & "                         case when LP.PaymentFrequency = 4 then 'QT' else"
                     mstrSQL = mstrSQL & "                         case when LP.PaymentFrequency = 2 then 'SA' else"
                     mstrSQL = mstrSQL & "                         case when LP.PaymentFrequency = 1 then 'AN' else '??' end end end end + '</OR.PYMTFREQSHORT>'"
                     mstrSQL = mstrSQL & " + '<OR.PYMTPLUSSERVPLUSPM>' + dbo.fn_FormatCurrency(OD.PaymentAmtMonthlyFin + (select sum(isnull(BundleQuantity,1) * isnull(MonthlyMeter,0)) from SNOrderLine with (nolock) where OrderID = OD.OrderID)) + '</OR.PYMTPLUSSERVPLUSPM>'"
                     mstrSQL = mstrSQL & " + '<OR.PYMTSTBL>' + case when isnull(ST.IsLeased,0) = 0 and OD.SaleTypeID > 0 then 'n/a' else case when isnull(OD.NbrDeferredPymts,0) = 0 then case when OD.PaymentAmtMonthlyFin = 0 then '' else case when isnull(BT.IncludeInLease,0) = -1 then dbo.fn_FormatCurrency(isnull(OD.AmtFinanced,0) * isnull(OD.LeaseFactor,0)) else dbo.fn_FormatCurrency(OD.PaymentAmtMonthlyFin) end end else '$0.00' + char(13) + case when isnull(BT.IncludeInLease,0) = -1 then dbo.fn_FormatCurrency(isnull(OD.AmtFinanced,0) * isnull(OD.LeaseFactor,0)) else dbo.fn_FormatCurrency(OD.PaymentAmtMonthlyFin) end end end + '</OR.PYMTSTBL>'"
                     mstrSQL = mstrSQL & " + '<OR.PYMTSTBLDEF>' + case when isnull(ST.IsLeased,0) = 0 and OD.SaleTypeID > 0 then 'n/a' else case when isnull(OD.NbrDeferredPymts,0) > 0 then cast(isnull(OD.NbrDeferredPymts,0) as nvarchar) + ' @ $0.00' + char(13) + cast(OD.NbrLeasePymts - isnull(OD.NbrDeferredPymts,0) as nvarchar) + ' @ ' + dbo.fn_FormatCurrency(OD.PaymentAmtMonthlyFin) else cast(OD.NbrLeasePymts as nvarchar) + ' @ ' + dbo.fn_FormatCurrency(OD.PaymentAmtMonthlyFin) end end + '</OR.PYMTSTBLDEF>'"
                     mstrSQL = mstrSQL & " + '<OR.PYMTSTBLTOTAL>' + case when isnull(ST.IsLeased,0) = 0 and OD.SaleTypeID > 0 then 'n/a' else case when isnull(OD.NbrDeferredPymts,0) = 0 then case when OD.PaymentAmtMonthlyFin = 0 then '' else dbo.fn_FormatCurrency(OD.PaymentAmtMonthlyFin) end else '$0.00' + char(13) + dbo.fn_FormatCurrency(OD.PaymentAmtMonthlyFin) end end + '</OR.PYMTSTBLTOTAL>'"
                     mstrSQL = mstrSQL & " + '<OR.REPCOSTANDEXPENSES>' + dbo.fn_FormatCurrency((select sum(BuyPrice * Quantity) from SNOrderLine with (nolock) where OrderID = OD.OrderID) + (select sum(case when isnull(oat.IsSoftCost,0) = -1 then case when oat.CreditDebitInd = 'C' then -1 * oa.AdjAmount else oa.AdjAmount end else 0 end) from SNOrderAdjustment oa with (nolock) join SNOrderAdjType oat with (nolock) on oat.OrderAdjTypeID = oa.OrderAdjTypeID where oa.OrderID = OD.OrderID)) + '</OR.REPCOSTANDEXPENSES>'"
                     mstrSQL = mstrSQL & " + '<OR.SALESMANAGERNAME>' +  dbo.fn_FormatXMLChars(isnull(SMP.FirstName,'') + ' ' + case when isnull(SMP.MiddleName,'') <> '' then  SMP.MiddleName + ' ' else '' end + isnull(SMP.LastName,'')) + '</OR.SALESMANAGERNAME>'"
                     mstrSQL = mstrSQL & " + '<OR.SALESREP>' + dbo.fn_FormatXMLChars(isnull(SRP.FirstName,'') + ' ' + case when isnull(SRP.MiddleName,'') <> '' then  SRP.MiddleName + ' ' else '' end + isnull(SRP.LastName,'')) + '</OR.SALESREP>'"
                     mstrSQL = mstrSQL & " + '<OR.SALESREPCELLPHONE>' + dbo.fn_FormatPhoneNumber((select p.PhoneNbr from SNPhone p with (nolock) join SNEntityPhone e with (nolock) on e.PhoneID = p.PhoneID where e.EntityID = SRP.EntityID and e.PhoneTypeID = (select PhoneTypeID from SNPhoneType with (nolock) where SystemLabel = 'Person.CellPhone'))) + '</OR.SALESREPCELLPHONE>'"
                     mstrSQL = mstrSQL & " + '<OR.SALESREPEMAIL>' + dbo.fn_FormatXMLChars(isnull(SRP.EmailAddress ,'')) + '</OR.SALESREPEMAIL>'"
                     mstrSQL = mstrSQL & " + '<OR.SALESREPPHONE>' + dbo.fn_FormatPhoneNumber(PH.PhoneNbr) + '</OR.SALESREPPHONE>'"
                     mstrSQL = mstrSQL & " + '<OR.SALESREPPHONEEXT>' + dbo.fn_FormatXMLChars(isnull(PH.PhoneExt,'')) + '</OR.SALESREPPHONEEXT>'"
                     mstrSQL = mstrSQL & " + '<OR.SALESREPTITLE>' + dbo.fn_FormatXMLChars(isnull(SR.Title,'')) + '</OR.SALESREPTITLE>'"
                     mstrSQL = mstrSQL & " + '<OR.SALESREPXREF>' + dbo.fn_FormatXMLChars(isnull(EN.ImportXref ,'')) + '</OR.SALESREPXREF>'"
                     mstrSQL = mstrSQL & " + '<OR.SERVICECOMMENTS>' + dbo.fn_FormatXMLChars(isnull(OD.SvcComments,'')) + '</OR.SERVICECOMMENTS>'"
                     mstrSQL = mstrSQL & " + '<OR.SRCCOMMAMT>' + dbo.fn_FormatCurrency(OD.SalesRepCommissionAmt) + '</OR.SRCCOMMAMT>'"
                     mstrSQL = mstrSQL & " + '<OR.SRCCOMMBONUS>' + dbo.fn_FormatCurrency(OD.SRCommMaintBonus) + '</OR.SRCCOMMBONUS>'"
                     mstrSQL = mstrSQL & " + '<OR.SRCCOMMPCT>' + dbo.fn_formatnumber(OD.SRCommPCTPaid * 100,2) + '%' + '</OR.SRCCOMMPCT>'"
                     mstrSQL = mstrSQL & " + '<OR.SRCCOMMSVCBONUS>' + dbo.fn_FormatCurrency(OD.SRCommMaintService) + '</OR.SRCCOMMSVCBONUS>'"
                     mstrSQL = mstrSQL & " + '<OR.SRCINVOICENBR>' + dbo.fn_FormatXMLChars(isnull(OD.InvoiceNbr,'')) + '</OR.SRCINVOICENBR>'"
                     mstrSQL = mstrSQL & " + '<OR.SRCLEVELNAME>' + dbo.fn_FormatXMLChars(isnull(CLVL.Name,'')) + '</OR.SRCLEVELNAME>'"
                     mstrSQL = mstrSQL & " + '<OR.SRCNOTE>' + dbo.fn_FormatXMLChars(cast(isnull(OD.SRCommNote,'') as nvarchar(max))) + '</OR.SRCNOTE>'"
                     mstrSQL = mstrSQL & " + '<OR.SRCOTHERSALESREPCOMMAMT>' + dbo.fn_FormatCurrency(OD.SplitSRCommissionAmt) + '</OR.SRCOTHERSALESREPCOMMAMT>'"
                     mstrSQL = mstrSQL & " + '<OR.SRCOTHERSALESREPPCT>' + dbo.fn_formatnumber(OD.OtherSalesRepPCT,2) + '%' + '</OR.SRCOTHERSALESREPPCT>'"
                     mstrSQL = mstrSQL & " + '<OR.SRCPRIAMT>' + dbo.fn_FormatCurrency((select sum(EarnedCommAmt) from SNSRCommEarned with (nolock) where OrderID = OD.OrderID and UserID = OD.SalesRepUserID)) + '</OR.SRCPRIAMT>'"
                     mstrSQL = mstrSQL & " + '<OR.SRCPRIAMTEQCOMM>' + dbo.fn_FormatCurrency(isnull(OD.AdjGPAmt,0) * isnull(OD.SRCommPCTPaid,0)) + '</OR.SRCPRIAMTEQCOMM>'"
                     mstrSQL = mstrSQL & " + '<OR.SRCPRIMSALESREPCOMMAMT>' + dbo.fn_FormatCurrency(OD.PrimarySRCommissionAmt) + '</OR.SRCPRIMSALESREPCOMMAMT>'"
                     mstrSQL = mstrSQL & " + '<OR.SRCSALESORDER>' + dbo.fn_FormatXMLChars(isnull(OD.SalesOrderNbr,'')) + '</OR.SRCSALESORDER>'"
                     mstrSQL = mstrSQL & " + '<OR.SRCSEGMENTBONUSAMT>' + dbo.fn_FormatCurrency(OD.SegmentBonusValue) + '</OR.SRCSEGMENTBONUSAMT>'"
                     mstrSQL = mstrSQL & " + '<OR.SRCSERVICEAMT>' + dbo.fn_FormatCurrency(OD.SRCommServiceValue) + '</OR.SRCSERVICEAMT>'"
                     mstrSQL = mstrSQL & " + '<OR.SRCSERVICERATE>' + dbo.fn_formatnumber(OD.SRCommServiceRate * 100,2) + '%' + '</OR.SRCSERVICERATE>'"
                     mstrSQL = mstrSQL & " + '<OR.SRCSPLITREPNAME>' + dbo.fn_FormatXMLChars(isnull((select isnull(oPR.FirstName,'') + ' ' + isnull(oPR.LastName,'') from SNUser oUS with (nolock) join SNPerson oPR with (nolock) on oPR.PersonID = oUS.PersonID where oUS.UserID = OD.OtherSalesRepUserID),'')) + '</OR.SRCSPLITREPNAME>'"
                     mstrSQL = mstrSQL & " + '<OR.SRCTOTALCOMMAMT>' + dbo.fn_FormatCurrency(OD.SRCommTotalAmt) + '</OR.SRCTOTALCOMMAMT>'"
                     mstrSQL = mstrSQL & " + '<OR.STATUSNOTE>' + dbo.fn_FormatXMLChars(isnull((select top 1 Note from SNOrderHistory h with (nolock) where h.OrderID = OD.OrderID and h.ToStatusID = OD.OrderStatusID order by DTChanged desc),'')) + '</OR.STATUSNOTE>'"
                     mstrSQL = mstrSQL & " + '<OR.SUBTOTALSELLPRICE>' + dbo.fn_FormatCurrency(OD.SubTotalAmount) + '</OR.SUBTOTALSELLPRICE>'"
                     mstrSQL = mstrSQL & " + '<OR.SUBTOTAL24MOPYMT>' + dbo.fn_FormatCurrency((select sum(val) from (select cast(convert(nvarchar(50),convert(money,coalesce(l.LineTotal * o.Lease24moFactor,0)),1) as money) val from SNOrderLine l with (nolock) join SNOrder o with (nolock) on o.OrderID = l.OrderID where isnull(l.ServiceInd,0) = 0 and l.OrderID = OD.OrderID) t)) + '</OR.SUBTOTAL24MOPYMT>'"
                     mstrSQL = mstrSQL & " + '<OR.SUBTOTAL36MOPYMT>' + dbo.fn_FormatCurrency((select sum(val) from (select cast(convert(nvarchar(50),convert(money,coalesce(l.LineTotal * o.Lease36moFactor,0)),1) as money) val from SNOrderLine l with (nolock) join SNOrder o with (nolock) on o.OrderID = l.OrderID where isnull(l.ServiceInd,0) = 0 and l.OrderID = OD.OrderID) t)) + '</OR.SUBTOTAL36MOPYMT>'"
                     mstrSQL = mstrSQL & " + '<OR.SUBTOTAL48MOPYMT>' + dbo.fn_FormatCurrency((select sum(val) from (select cast(convert(nvarchar(50),convert(money,coalesce(l.LineTotal * o.Lease48moFactor,0)),1) as money) val from SNOrderLine l with (nolock) join SNOrder o with (nolock) on o.OrderID = l.OrderID where isnull(l.ServiceInd,0) = 0 and l.OrderID = OD.OrderID) t)) + '</OR.SUBTOTAL48MOPYMT>'"
                     mstrSQL = mstrSQL & " + '<OR.SUBTOTAL60MOPYMT>' + dbo.fn_FormatCurrency((select sum(val) from (select cast(convert(nvarchar(50),convert(money,coalesce(l.LineTotal * o.Lease60moFactor,0)),1) as money) val from SNOrderLine l with (nolock) join SNOrder o with (nolock) on o.OrderID = l.OrderID where isnull(l.ServiceInd,0) = 0 and l.OrderID = OD.OrderID) t)) + '</OR.SUBTOTAL60MOPYMT>'"
                     mstrSQL = mstrSQL & " + '<OR.SUBTOTALLEASEPYMT>' + dbo.fn_FormatCurrency((select sum(val) from (select cast(convert(nvarchar(50),convert(money,coalesce(l.LineTotal * o.LeaseFactor,0)),1) as money) val from SNOrderLine l with (nolock) join SNOrder o with (nolock) on o.OrderID = l.OrderID where isnull(l.ServiceInd,0) = 0 and l.OrderID = OD.OrderID) t)) + '</OR.SUBTOTALLEASEPYMT>'"
                     mstrSQL = mstrSQL & " + '<OR.SUBTOTALUNITLEASEPYMT>' + dbo.fn_FormatCurrency((select sum(val) from (select cast(convert(nvarchar(50),convert(money,coalesce(l.SellPrice * o.LeaseFactor,0)),1) as money) val from SNOrderLine l with (nolock) join SNOrder o with (nolock) on o.OrderID = l.OrderID where isnull(l.ServiceInd,0) = 0 and l.OrderID = OD.OrderID) t)) + '</OR.SUBTOTALUNITLEASEPYMT>'"
                     mstrSQL = mstrSQL & " + '<OR.SVANNUALGP>' + case when isnull(OD.SVANNUALGP,0) = 0 then '' else dbo.fn_FormatCurrency(OD.SVANNUALGP) end + '</OR.SVANNUALGP>'"
                     mstrSQL = mstrSQL & " + '<OR.SVANNUALPRICE>' + case when isnull(OD.SVANNUALPRICE,0) = 0 then '' else dbo.fn_FormatCurrency(OD.SVANNUALPRICE) end + '</OR.SVANNUALPRICE>'"
                     mstrSQL = mstrSQL & " + '<OR.SVBASEBILLFREQANX>' + case when isnull(SBT.PaymentsPerYear,0) = 1 then 'X' else '' end + '</OR.SVBASEBILLFREQANX>'"
                     mstrSQL = mstrSQL & " + '<OR.SVBASEBILLFREQMOX>' + case when isnull(SBT.PaymentsPerYear,0) = 12 then 'X' else '' end + '</OR.SVBASEBILLFREQMOX>'"
                     mstrSQL = mstrSQL & " + '<OR.SVBASEBILLFREQQTRX>' + case when isnull(SBT.PaymentsPerYear,0) = 4 then 'X' else '' end + '</OR.SVBASEBILLFREQQTRX>'"
                     mstrSQL = mstrSQL & " + '<OR.SVBASEBILLFREQSAX>' + case when isnull(SBT.PaymentsPerYear,0) = 2 then 'X' else '' end + '</OR.SVBASEBILLFREQSAX>'"
                     mstrSQL = mstrSQL & " + '<OR.SVBILLCYCLENAME>' + dbo.fn_FormatXMLChars(isnull(SBT.Name,'')) + '</OR.SVBILLCYCLENAME>'"
                     mstrSQL = mstrSQL & " + '<OR.SVBILLCYCLENAMEOTHER>' + dbo.fn_FormatXMLChars(isnull(SBT.OtherName,'')) + '</OR.SVBILLCYCLENAMEOTHER>'"
                     mstrSQL = mstrSQL & " + '<OR.SVBILLCYCLEPPY>' + case when isnull(SBT.PaymentsPerYear,0) = 0 then '' else dbo.fn_FORMATNumber(SBT.PaymentsPerYear,0) end + '</OR.SVBILLCYCLEPPY>'"
                     mstrSQL = mstrSQL & " + '<OR.SVCOSTPERUSER>' + case when isnull(OD.SVCOSTPERUSER,0) = 0 then '' else dbo.fn_FormatCurrency(OD.SVCOSTPERUSER) end + '</OR.SVCOSTPERUSER>'"
                     mstrSQL = mstrSQL & " + '<OR.SVENDDT>' + case when isnull(OD.SVSTARTDT,'1/1/1900') = '1/1/1900' or isnull(OD.SVTERM,0) = 0 then '' else convert(nvarchar, dateadd(mm,OD.SVTERM,OD.SVSTARTDT), 101) end + '</OR.SVENDDT>'"
                     mstrSQL = mstrSQL & " + '<OR.SVMOCOST>' + case when isnull(OD.SVMOCOST,0) = 0 then '' else dbo.fn_FormatCurrency(OD.SVMOCOST) end + '</OR.SVMOCOST>'"
                     mstrSQL = mstrSQL & " + '<OR.SVMOGP>' + case when isnull(OD.SVMOGP,0) = 0 then '' else dbo.fn_FormatCurrency(OD.SVMOGP) end + '</OR.SVMOGP>'"
                     mstrSQL = mstrSQL & " + '<OR.SVMOREPCOST>' + case when isnull(OD.SVMOREPCOST,0) = 0 then '' else dbo.fn_FormatCurrency(OD.SVMOREPCOST) end + '</OR.SVMOREPCOST>'"
                     mstrSQL = mstrSQL & " + '<OR.SVMOSELLPRICE>' + case when isnull(OD.SVMOSELLPRICE,0) = 0 then '' else dbo.fn_FormatCurrency(OD.SVMOSELLPRICE) end + '</OR.SVMOSELLPRICE>'"
                     mstrSQL = mstrSQL & " + '<OR.SVMOSELLPRICENOSIGN>' + case when isnull(OD.SVMOSELLPRICE,0) = 0 then '' else replace(dbo.fn_FormatCurrency(OD.SVMOSELLPRICE),'$','') end + '</OR.SVMOSELLPRICENOSIGN>'"
                     mstrSQL = mstrSQL & " + '<OR.SVNBRUSERS>' + case when isnull(OD.SVNBRUSERS,0) = 0 then '' else dbo.fn_FormatNumber(OD.SVNBRUSERS,0) end + '</OR.SVNBRUSERS>'"
                     mstrSQL = mstrSQL & " + '<OR.SVPERIODSELLPRICE>' + case when isnull(OD.SVMOSELLPRICE,0) = 0 or isnull(SBT.PaymentsPerYear,0) = 0 then '' else dbo.fn_FormatCurrency(OD.SVMOSELLPRICE * (12/isnull(SBT.PaymentsPerYear,12))) end + '</OR.SVPERIODSELLPRICE>'"
                     mstrSQL = mstrSQL & " + '<OR.SVSTARTDT>' + case when isnull(OD.SVSTARTDT,'1/1/1900') = '1/1/1900' then '' else convert(nvarchar, OD.SVSTARTDT, 101) end + '</OR.SVSTARTDT>'"
                     mstrSQL = mstrSQL & " + '<OR.SVTERM>' + case when isnull(OD.SVTERM,0) = 0 then '' else dbo.fn_FormatNumber(OD.SVTERM,0) end + '</OR.SVTERM>'"
                     mstrSQL = mstrSQL & " + '<OR.SVTOTALGP>' + case when isnull(OD.SVTOTALGP,0) = 0 then '' else dbo.fn_FormatCurrency(OD.SVTOTALGP) end + '</OR.SVTOTALGP>'"
                     mstrSQL = mstrSQL & " + '<OR.SVTOTALVALUE>' + case when isnull(OD.SVTOTALVALUE,0) = 0 then '' else dbo.fn_FormatCurrency(OD.SVTOTALVALUE) end + '</OR.SVTOTALVALUE>'"
                     mstrSQL = mstrSQL & " + '<OR.TOTALADVLEASEPYMT>' + case when isnull(ST.IsLeased,0) = 0 and OD.SaleTypeID > 0 then 'n/a' else dbo.fn_FormatCurrency((isnull(OD.AmtFinanced ,0) * isnull(OD.LeaseFactor ,0)) * isnull(LP.NbrDown,0))  end + '</OR.TOTALADVLEASEPYMT>'"
                     mstrSQL = mstrSQL & " + '<OR.TOTALDEALERCOST>' + dbo.fn_FormatCurrency((select sum(OL.DealerCostAmt) from SNOrderLine OL with (nolock) where OL.OrderID = OD.OrderID)) + '</OR.TOTALDEALERCOST>'"
                     mstrSQL = mstrSQL & " + '<OR.TOTALMSRP>' + dbo.fn_FormatCurrency(ISNULL(OD.TotalMSRP,0)) + '</OR.TOTALMSRP>'"
                     mstrSQL = mstrSQL & " + '<OR.TOTALMSRPPCTFIN>' + dbo.fn_formatnumber(ISNULL(case when OD.TotalMSRP > 0 and OD.AmtFinanced > 0 then OD.AmtFinanced/OD.TotalMSRP else OD.TotalMSRP end,0) * 100,0) + '%' + '</OR.TOTALMSRPPCTFIN>'"
                     mstrSQL = mstrSQL & " + '<OR.TOTALNUMADVLEASEPYMT>' + cast(isnull(LP.NbrDown,0) as nvarchar) + '</OR.TOTALNUMADVLEASEPYMT>'"
                     mstrSQL = mstrSQL & " + '<OR.TOTALNUMDEFLEASEPYMT>' + cast(isnull(LP.NbrDeferred,0) as nvarchar) + '</OR.TOTALNUMDEFLEASEPYMT>'"
                     mstrSQL = mstrSQL & " + '<OR.TOTALSELLNOTAX>' + dbo.fn_FormatCurrency(TotalOrderAmount - isnull((select AdjAmount from SNOrderAdjustment oa with (nolock) join SNOrderAdjType oat with (nolock) on oat.OrderAdjTypeID = oa.OrderAdjTypeID where oa.OrderID = OD.OrderID and oat.Taxind = -1),0)) + '</OR.TOTALSELLNOTAX>'"
                     mstrSQL = mstrSQL & " + '<OR.TOTALSELLPRICE>' + dbo.fn_FormatCurrency(TotalOrderAmount) + '</OR.TOTALSELLPRICE>'"
                     mstrSQL = mstrSQL & " + '<OR.TOTALSERVICECONTRACT>' + dbo.fn_FormatCurrency(OD.MaintTotalValue) + '</OR.TOTALSERVICECONTRACT>'"
                     mstrSQL = mstrSQL & " + '<OR.TRAVELCATEGORY>' + dbo.fn_FormatXMLChars(isnull(TC.Name,'')) + '</OR.TRAVELCATEGORY>'"
                     mstrSQL = mstrSQL & " + '<OR.TRAVELCATEGORYDOCDESC>' + dbo.fn_FormatXMLChars(isnull(TC.DocDescription,'')) + '</OR.TRAVELCATEGORYDOCDESC>'"
                     mstrSQL = mstrSQL & " + '<OR.TRAVELCATEGORYMINCHARGE>' + dbo.fn_FormatCurrency(TC.MinimumCharge) + '</OR.TRAVELCATEGORYMINCHARGE>'"
                     mstrSQL = mstrSQL & " + '<OR.TRAVELCATEGORYMILLAGECHARGE>' + dbo.fn_FormatCurrency(TC.MillageCharge) + '</OR.TRAVELCATEGORYMILLAGECHARGE>'"
                     mstrSQL = mstrSQL & " + '<OR.TRAVELCATEGORYPERPERSONDAILYCOST>' + dbo.fn_FormatCurrency(TC.DailyCostPerPersonAmt) + '</OR.TRAVELCATEGORYPERPERSONDAILYCOST>'"
                     mstrSQL = mstrSQL & " + '<OR.TRAVELFEENBRDAYS>' + dbo.fn_formatnumber(OD.TravelNbrNights,0) + '</OR.TRAVELFEENBRDAYS>'"
                     mstrSQL = mstrSQL & " + '<OR.TRAVELFEENBRMILES>' + dbo.fn_formatnumber(OD.TravelNbrMiles,0) + '</OR.TRAVELFEENBRMILES>'"
                     mstrSQL = mstrSQL & " + '<OR.TRAVELFEENBRPEOPLE>' + dbo.fn_formatnumber(OD.TravelNbrPeople,0) + '</OR.TRAVELFEENBRPEOPLE>'"
                     mstrSQL = mstrSQL & " + '<OR.TYPEOFOBJECT>' + case when OS.SystemLabel = 'PROP' then 'Proposal' else 'Order' end + '</OR.TYPEOFOBJECT>'"
                     mstrSQL = mstrSQL & " + '<OR.TYPEOFSALE>' + dbo.fn_FormatXMLChars(isnull(ST.Name,'')) + '</OR.TYPEOFSALE>'"
                     mstrSQL = mstrSQL & " + '<OR.TYPEORDERID>' + case when OS.SystemLabel = 'PROP' then 'P-' else '' end  + cast(OD.OrderID as nvarchar) + '</OR.TYPEORDERID>'"
                     mstrSQL = mstrSQL & " + '<OR.UPGRADEFCONAME>' + dbo.fn_FormatXMLChars(replace(isnull(UCO.Name,''),'_',' ')) + '</OR.UPGRADEFCONAME>'"
                     mstrSQL = mstrSQL & " + '<OR.UPGRADEFCOADDRESSONELINE>' + dbo.fn_FormatXMLChars(case when len(isnull(UFC.RPTCity,'')) > 0 then isnull(UFC.RPTAddress1,'') + case when LEN(isnull(UFC.RPTAddress2,'')) > 0 then ' ' + isnull(UFC.RPTAddress2,'') else '' end  + ', ' +  isnull(UFC.RPTCity,'') + ', ' + isnull(UFC.RPTState,'') + ' ' + isnull(UFC.RPTZip,'') else '' end) + '</OR.UPGRADEFCOADDRESSONELINE>'"
                     mstrSQL = mstrSQL & " + '<OR.UPGRADEFCOADDRESS1>' + dbo.fn_FormatXMLChars(isnull(UFC.RPTAddress1,'')) + '</OR.UPGRADEFCOADDRESS1>'"
                     mstrSQL = mstrSQL & " + '<OR.UPGRADEFCOADDRESS2>' + dbo.fn_FormatXMLChars(isnull(UFC.RPTAddress2,'')) + '</OR.UPGRADEFCOADDRESS2>'"
                     mstrSQL = mstrSQL & " + '<OR.UPGRADEFCOCITY>' + dbo.fn_FormatXMLChars(isnull(UFC.RPTCity,'')) + '</OR.UPGRADEFCOCITY>'"
                     mstrSQL = mstrSQL & " + '<OR.UPGRADEFCOSTATE>' + dbo.fn_FormatXMLChars(isnull(UFC.RPTState,'')) + '</OR.UPGRADEFCOSTATE>'"
                     mstrSQL = mstrSQL & " + '<OR.UPGRADEFCOZIP>' + dbo.fn_FormatXMLChars(isnull(UFC.RPTZip,'')) + '</OR.UPGRADEFCOZIP>'"
                     mstrSQL = mstrSQL & " + '<OR.UPGRADELSNBR>' + dbo.fn_FormatXMLChars(case when isnull(OD.LeaseUpgradeInd,0) = -1 then  isnull(OD.UpgradeFCOLeaseNbr,'') else '' end) + '</OR.UPGRADELSNBR>'"
                     mstrSQL = mstrSQL & " + '<OR.UPGRADENOX>' + case when isnull(OD.LeaseUpgradeInd,0) = 0 then 'X' else '' end + '</OR.UPGRADENOX>'"
                     mstrSQL = mstrSQL & " + '<OR.UPGRADETYPE>' + isnull((select Name from SNUpgradeType with (nolock) where UpgradeTypeID = OD.UpgradeTypeID),'') + '</OR.UPGRADETYPE>'"
                     mstrSQL = mstrSQL & " + '<OR.UPGRADEYESNO>' + case when isnull(OD.LeaseUpgradeInd,0) = -1 then 'Yes' else 'No' end + '</OR.UPGRADEYESNO>'"
                     mstrSQL = mstrSQL & " + '<OR.UPGRADEYESX>' + case when isnull(OD.LeaseUpgradeInd,0) = -1 then 'X' else '' end + '</OR.UPGRADEYESX>'"
                     If InStr(pField, "ADDENDUMNEEDED") > 0 And mIndex > 0 Then  'Dynamic mergefield based on index
                         mstrSQL = mstrSQL & " + '<OR.ISOLINEADDENDUMNEEDED>' + dbo.fn_FormatXMLChars(case when (SELECT COUNT(*) FROM SNOrderLine with (nolock) WHERE OrderID = OD.OrderID) > " & mIndex & " then 'ADDENDUM IS REQUIRED' + case when '" & mstrAddendumname & "' <> '' then ' (" & mstrAddendumname & ")' else '' end else '' end) + '</OR.ISOLINEADDENDUMNEEDED>'"
                         mstrSQL = mstrSQL & " + '<OR.ISOLINEADDENDUMNEEDEDX>' + case when (SELECT COUNT(*) FROM SNOrderLine with (nolock) WHERE OrderID = OD.OrderID) > " & mIndex & " then 'X' else '' end + '</OR.ISOLINEADDENDUMNEEDEDX>'"
                         mstrSQL = mstrSQL & " + '<OR.PRIMACHADDENDUMNEEDED>' + dbo.fn_FormatXMLChars(case when (SELECT SUM(SNOrderLine.Quantity) From SNOrderLine with (nolock) "
                         mstrSQL = mstrSQL & "           Left JOIN SNAssocCatalogType with (nolock) ON SNAssocCatalogType.AssocCatalogItemTypeID = SNOrderLine.AssocCatalogItemTypeID"
                         mstrSQL = mstrSQL & "           LEFT JOIN SNCatalogItem with (nolock) ON SNCatalogItem.CatalogItemID = SNOrderLine.CatalogItemID"
                         mstrSQL = mstrSQL & "           Where OrderID = OD.OrderID"
                         mstrSQL = mstrSQL & "           AND (SNAssocCatalogType.PrimaryInd = -1  OR  SNCatalogItem.PrimaryInd = -1"
                         mstrSQL = mstrSQL & "           OR (SNOrderLine.CatalogItemID = -100 and IsPrimaryInd = -1))) > " & mIndex & " then 'ADDENDUM IS REQUIRED' + case when '" & mstrAddendumname & "' <> '' then ' (" & mstrAddendumname & ")' else '' end else '' end) + '</OR.PRIMACHADDENDUMNEEDED>'"
                         mstrSQL = mstrSQL & " + '<OR.PRIMACHADDENDUMNEEDEDX>' + case when (SELECT SUM(SNOrderLine.Quantity) From SNOrderLine with (nolock) "
                         mstrSQL = mstrSQL & "           Left JOIN SNAssocCatalogType with (nolock) ON SNAssocCatalogType.AssocCatalogItemTypeID = SNOrderLine.AssocCatalogItemTypeID"
                         mstrSQL = mstrSQL & "           LEFT JOIN SNCatalogItem with (nolock) ON SNCatalogItem.CatalogItemID = SNOrderLine.CatalogItemID"
                         mstrSQL = mstrSQL & "           Where OrderID = OD.OrderID"
                         mstrSQL = mstrSQL & "           AND (SNAssocCatalogType.PrimaryInd = -1  OR  SNCatalogItem.PrimaryInd = -1"
                         mstrSQL = mstrSQL & "           OR (SNOrderLine.CatalogItemID = -100 and IsPrimaryInd = -1))) > " & mIndex & " then 'X' else '' end + '</OR.PRIMACHADDENDUMNEEDEDX>'"
                     End If
                     mstrSQL = mstrSQL & mstrAdjTags
                     mstrSQL = mstrSQL & " + '<OR.TOTALMOAMT>' + dbo.fn_FormatCurrency(OD.SPMTotalMoAdjAmt) + '</OR.TOTALMOAMT>'"
                     mstrSQL = mstrSQL & " + '<OR.TOTALMOAMTACTUAL>' + dbo.fn_FormatCurrency(OD.SPMTotalMoAdjAmtActual) + '</OR.TOTALMOAMTACTUAL>'"
                     mstrSQL = mstrSQL & " + '<OR.WTMOREVAMT>' + case when isnull(OD.WTMORevAmt,0) = 0 then '' else dbo.fn_FormatCurrency(OD.WTMORevAmt) end + '</OR.WTMOREVAMT>'"
                     mstrSQL = mstrSQL & " + '<OR.WTMOGPAMT>' + case when isnull(OD.WTMOGPAmt,0) = 0 then '' else dbo.fn_FormatCurrency(OD.WTMOGPAmt) end + '</OR.WTMOGPAMT>'"
                     mstrSQL = mstrSQL & " + '<OR.WTTERM>' + case when isnull(OD.WTTerm,0) = 0 then '' else dbo.fn_FormatNumber(OD.WTTerm,0) end + '</OR.WTTERM>'"
                     mstrSQL = mstrSQL & " + '<OR.CREDITAPPNOTE>' + dbo.fn_FormatXMLChars(isnull(OD.CreditAppNote ,'')) + '</OR.CREDITAPPNOTE>'"
                     mstrSQL = mstrSQL & " + '</OR>'"
                     mstrSQL = mstrSQL & " FROM SNOrder AS OD with (nolock) "
                     mstrSQL = mstrSQL & " LEFT JOIN SNContractType as CT with (nolock) on CT.ContractTypeID = OD.MaintContractTypeID "
                     mstrSQL = mstrSQL & " LEFT JOIN SNBillingCycleType as BT with (nolock) on BT.BillingCycleTypeID = OD.MaintBillingCycle "
                     mstrSQL = mstrSQL & " LEFT JOIN SNBillingCycleType as OBT with (nolock) on OBT.BillingCycleTypeID = OD.MaintOverBillingCycle"
                     mstrSQL = mstrSQL & " LEFT JOIN SNBillingCycleType as SBT with (nolock) on SBT.BillingCycleTypeID = OD.SVBillingCycleID"
                     mstrSQL = mstrSQL & " LEFT JOIN SNLeaseProduct as LP with (nolock) on LP.LeaseProductID = OD.LeaseProductID "
                     mstrSQL = mstrSQL & " LEFT JOIN SNSaleType as ST with (nolock) on ST.SaleTypeID = OD.SaleTypeID "
                     mstrSQL = mstrSQL & " LEFT JOIN SNOrderStatus as OS with (nolock) on OS.OrderStatusID = OD.OrderStatusID "
                     mstrSQL = mstrSQL & " LEFT JOIN SNPriceLevel as PL with (nolock) on PL.PriceLevelID = OD.PriceLevelID "
                     mstrSQL = mstrSQL & " LEFT JOIN SNLeasePriceLevel as LPL with (nolock) on LPL.LeasePriceLevelID = OD.LeasePriceLevelID "
                     mstrSQL = mstrSQL & " LEFT JOIN SNCustomer as BCU with (nolock) on BCU.CustomerID = OD.BillToAccountSelect "
                     mstrSQL = mstrSQL & " LEFT JOIN SNUser as SR with (nolock) on SR.UserID = OD.SalesRepUserID "
                     mstrSQL = mstrSQL & " LEFT JOIN SNPerson as SRP with (nolock) on SRP.PersonID = SR.PersonID "
                     mstrSQL = mstrSQL & " LEFT JOIN SNEntityPhone as EP with (nolock) on EP.EntityID = SRP.EntityID and EP.PhoneTypeID = (Select PhoneTypeID FROM SNPhoneType with (nolock) where Name = 'Office') "
                     mstrSQL = mstrSQL & " LEFT JOIN SNPhone as PH with (nolock) on PH.PhoneID = EP.PhoneID "
                     mstrSQL = mstrSQL & " LEFT JOIN SNEntity as EN with (nolock) on EN.EntityID = SRP.EntityID "
                     mstrSQL = mstrSQL & " LEFT JOIN SNBranch as BR with (nolock) on BR.BranchID = SR.BranchID "
                     mstrSQL = mstrSQL & " LEFT JOIN SNBusiness as BU with (nolock) on BU.BusinessID = BR.BusinessID "
                     mstrSQL = mstrSQL & " LEFT JOIN SNProspect as PR with (nolock) on PR.ProspectID = OD.ProspectID "
                     mstrSQL = mstrSQL & " LEFT JOIN SNFinanceCompany as FC with (nolock) on FC.FCOID = OD.FCOID"
                     mstrSQL = mstrSQL & " LEFT JOIN SNCompany as CO with (nolock) on CO.CompanyID = FC.CompanyID"
                     mstrSQL = mstrSQL & " LEFT JOIN SNFinanceCompany as UFC with (nolock) on UFC.FCOID = OD.UpgradeFCOID"
                     mstrSQL = mstrSQL & " LEFT JOIN SNCompany as UCO with (nolock) on UCO.CompanyID = UFC.CompanyID"
                     mstrSQL = mstrSQL & " LEFT JOIN SNSRCommCSMPLevel as CSMP with (nolock) on CSMP.SRCommCSMPLevelID = OD.SRCommCSMPLevelID"
                     mstrSQL = mstrSQL & " LEFT JOIN SNSRCommLevel as CLVL with (nolock) on CLVL.SRCommLevelID = OD.SRCommLevelID"
                     mstrSQL = mstrSQL & " LEFT JOIN SNUser as SMU with (nolock) on SMU.UserID = BR.SalesManagerUserID"
                     mstrSQL = mstrSQL & " LEFT JOIN SNPerson as SMP with (nolock) on SMP.PersonID = SMU.PersonID"
                     mstrSQL = mstrSQL & " LEFT JOIN SNDMPaymentTerms as IPB with (nolock) on IPB.DMPaymentTermsID = OD.ITPPaymentTermsID"
                     mstrSQL = mstrSQL & " LEFT JOIN SNDMPaymentTerms as BTB with (nolock) on BTB.DMPaymentTermsID = OD.ITP_BOT_PaymentTermsID"
                     mstrSQL = mstrSQL & " LEFT JOIN SNDMPaymentTerms as DMB with (nolock) on DMB.DMPaymentTermsID = OD.DMPaymentTermsID"
                     mstrSQL = mstrSQL & " LEFT JOIN SNCreditApplication as CA with (nolock) on CA.OrderID = OD.OrderID AND CA.SelectedSource = -1"
                     mstrSQL = mstrSQL & " LEFT JOIN SNFinanceCompany as CFC with (nolock) on CFC.FCOID = CA.FCOID"
                     mstrSQL = mstrSQL & " LEFT JOIN SNTravelCategory as TC with (nolock) on TC.TravelCategoryID = OD.TravelCategoryID"
                     If UCase(cstrObjectClassName) = "CORDER" Then
                         mstrSQL = mstrSQL & " Where OD.OrderID = " & clngObjectKey
                     Else
                         mstrSQL = mstrSQL & " Where OD.OrderID = (SELECT OrderID FROM SNDeliveryJob with (nolock) WHERE DeliveryJobID = " & clngObjectKey & ")"
                     End If
                     
                     Set mRS = mobjDBServer.DoQuery(CStr(evDBConnString), CStr(mstrSQL))
                   If mRS.RecordCount > 0 Then
                        mstrXML = "<ORDER>" & mRS.GetString(2) & "</ORDER>"
                   Else
                        mstrXML = "<ORDER><OR></OR></ORDER>"
                   End If
                
                   Set mobjDom = mobjCommon.RetrieveDOM(mstrXML)
                   If mobjDom.parseError.errorCode <> 0 Then
                        GetMergeDataOR = pField
                   End If
                   carrXML(clngTagCount, 0) = mstrTagPrefix
                   carrXML(clngTagCount, 1) = mstrXML
                   clngTagCount = clngTagCount + 1
                Else
                
                   If InStr(pField, "ADDENDUMNEEDED") > 0 And mIndex > 0 Then  'Dynamic mergefield based on index
                         mstrSQL = "SELECT '<OR>'"
                         mstrSQL = mstrSQL & " + '<OR.ISOLINEADDENDUMNEEDED>' + dbo.fn_FormatXMLChars(case when (SELECT COUNT(*) FROM SNOrderLine with (nolock) WHERE OrderID = OD.OrderID) > " & mIndex & " then 'ADDENDUM ' + case when '" & mstrAddendumname & "' <> '' then '(" & mstrAddendumname & ") ' else '' end + 'IS REQUIRED' else '' end) + '</OR.ISOLINEADDENDUMNEEDED>'"
                         mstrSQL = mstrSQL & " + '<OR.ISOLINEADDENDUMNEEDEDX>' + case when (SELECT COUNT(*) FROM SNOrderLine with (nolock) WHERE OrderID = OD.OrderID) > " & mIndex & " then 'X' else '' end + '</OR.ISOLINEADDENDUMNEEDEDX>'"
                         mstrSQL = mstrSQL & " + '<OR.PRIMACHADDENDUMNEEDED>' + dbo.fn_FormatXMLChars(case when (SELECT SUM(SNOrderLine.Quantity) From SNOrderLine with (nolock) "
                         mstrSQL = mstrSQL & "           Left JOIN SNAssocCatalogType with (nolock) ON SNAssocCatalogType.AssocCatalogItemTypeID = SNOrderLine.AssocCatalogItemTypeID"
                         mstrSQL = mstrSQL & "           LEFT JOIN SNCatalogItem with (nolock) ON SNCatalogItem.CatalogItemID = SNOrderLine.CatalogItemID"
                         mstrSQL = mstrSQL & "           Where OrderID = OD.OrderID"
                         mstrSQL = mstrSQL & "           AND (SNAssocCatalogType.PrimaryInd = -1  OR  SNCatalogItem.PrimaryInd = -1"
                         mstrSQL = mstrSQL & "           OR (SNOrderLine.CatalogItemID = -100 and IsPrimaryInd = -1))) > " & mIndex & " then 'ADDENDUM ' + case when '" & mstrAddendumname & "' <> '' then '(" & mstrAddendumname & ") ' else '' end + 'IS REQUIRED' else '' end) + '</OR.PRIMACHADDENDUMNEEDED>'"
                         mstrSQL = mstrSQL & " + '<OR.PRIMACHADDENDUMNEEDEDX>' + case when (SELECT SUM(SNOrderLine.Quantity) From SNOrderLine with (nolock) "
                         mstrSQL = mstrSQL & "           Left JOIN SNAssocCatalogType with (nolock) ON SNAssocCatalogType.AssocCatalogItemTypeID = SNOrderLine.AssocCatalogItemTypeID"
                         mstrSQL = mstrSQL & "           LEFT JOIN SNCatalogItem with (nolock) ON SNCatalogItem.CatalogItemID = SNOrderLine.CatalogItemID"
                         mstrSQL = mstrSQL & "           Where OrderID = OD.OrderID"
                         mstrSQL = mstrSQL & "           AND (SNAssocCatalogType.PrimaryInd = -1  OR  SNCatalogItem.PrimaryInd = -1"
                         mstrSQL = mstrSQL & "           OR (SNOrderLine.CatalogItemID = -100 and IsPrimaryInd = -1))) > " & mIndex & " then 'X' else '' end + '</OR.PRIMACHADDENDUMNEEDEDX>'"
                   
                         mstrSQL = mstrSQL & " + '</OR>'"
                         mstrSQL = mstrSQL & " FROM SNOrder AS OD with (nolock) "
                         If UCase(cstrObjectClassName) = "CORDER" Then
                             mstrSQL = mstrSQL & " Where OD.OrderID = " & clngObjectKey
                         Else
                             mstrSQL = mstrSQL & " Where OD.OrderID = (SELECT OrderID FROM SNDeliveryJob with (nolock) WHERE DeliveryJobID = " & clngObjectKey & ")"
                         End If
                           
                         Set mRS = mobjDBServer.DoQuery(CStr(evDBConnString), CStr(mstrSQL))
                         If mRS.RecordCount > 0 Then
                              mstrXML = "<ORDER>" & mRS.GetString(2) & "</ORDER>"
                         Else
                              mstrXML = "<ORDER><OR></OR></ORDER>"
                         End If
                   End If
                   
                   Set mobjDom = mobjCommon.RetrieveDOM(mstrXML)
                   If mobjDom.parseError.errorCode <> 0 Then
                        GetMergeDataOR = pField
                   End If
                End If
                
                GetMergeDataOR = mobjDom.selectSingleNode("//" & mstrTagPrefix & "/" & UCase(pField)).Text
            End If
          
        Case Else
        mobjCommon.LogActivity cstrObjectName & "." & mstrMethodName, _
                            Now, _
                            1, _
                            pField & ", index:" & TagIndex, _
                            ""
    End Select
    
Cleanup:
    Set mobjCommon = Nothing

    Exit Function
    
errHandler:
    GetMergeDataOR = ""
        
    If TagIndex = "" Or InStr(TagIndex, "-") > 0 Then
        If InStr(TagIndex, "-") > 0 Then
            If IsNumeric(Left(TagIndex, InStr(TagIndex, "-") - 1)) = True And IsNumeric(Mid(TagIndex, InStr(TagIndex, "-") + 1)) = True Then
                If Left(TagIndex, InStr(TagIndex, "-") - 1) <= 1 And Mid(TagIndex, InStr(TagIndex, "-") + 1) <= 1 Then
                    mobjCommon.LogActivity cstrObjectName & "." & mstrMethodName, _
                                    Now, _
                                    1, _
                                    pField & ", index:" & TagIndex, _
                                    ""
                End If
            End If
        Else
            mobjCommon.LogActivity cstrObjectName & "." & mstrMethodName, _
                                Now, _
                                1, _
                                pField & ", index:" & TagIndex, _
                                ""
        End If
        GoTo Cleanup:
    End If
End Function