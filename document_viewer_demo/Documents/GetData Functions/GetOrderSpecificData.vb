Private Function GetOrderSpecificData(pField As String, pIndex As Integer, _
                                pAddendumName As String, pCountDown As Integer, _
                                pDJSeqNbr As Long, pcol As Collection, _
                                pcurOB_BuyAmount As Currency, pcurOB_SellAmount As Currency, _
                                pcurOB_LeasePymt As Currency, _
                                pLineIx As Integer, pMachineIx As Integer, pOffsetIx As Integer, pInt As Integer) As String
    Dim mstrReturn As String
    Dim mIndex As Integer
    Dim mstrAddendumname As String
    Dim mcurTotal As Currency
    Dim mint As Integer
    Dim mdatTemp As Date
    Dim mlngTotal As Long
    Dim mrsSums As ADODB.Recordset
    Dim mintCountDown As Integer
    Dim mintLineIx As Integer
    Dim mintMachineIx As Integer
    Dim mstrTemp As String
    Dim mintRowIx As Integer
    Dim mstrMethod(10) As String
    Dim mbolFound As Boolean
    Dim mintDJSeqNbr As Long
    Dim mstrErrLocation As String
    Dim mlngLocTypeID As Long
    Dim mobjDelJobOption As Object
    Dim mintOffsetIx As Integer
    Dim mTmpObj As Object
    Dim mcurOB_BuyAmount As Currency
    Dim mcurOB_SellAmount As Currency
    Dim mcurOB_LeasePymt As Currency
    
    mcurOB_BuyAmount = pcurOB_BuyAmount
    mcurOB_SellAmount = pcurOB_SellAmount
    mcurOB_LeasePymt = pcurOB_LeasePymt
    
    mIndex = pIndex
    mintLineIx = pLineIx
    mintMachineIx = pMachineIx
    mstrAddendumname = pAddendumName
    mintCountDown = pCountDown
    mintDJSeqNbr = pDJSeqNbr
    mintOffsetIx = pOffsetIx
    mint = pInt
    
    Select Case UCase(pField)
        Case "OR.ADDRESSLABEL"
            mstrReturn = cobjOrder.AddressLabel
        Case "OR.ADJCOMMBASEAMT"
            mstrReturn = Format(cobjOrder.ADJCOMMBASEAMT, "$###,###,##0.00")
        Case "OR.ADJGPAMT"
            mstrReturn = Format(cobjOrder.AdjGPAmt, "$###,###,##0.00")
        Case "OR.TOTALCSMPAMT"
            mstrReturn = Format(cobjOrder.CSMPPCTAmt, "$###,###,##0.00")
        
        Case "OR.CSMPAMTALL"
            mstrReturn = Format(cobjOrder.SRCommCSMPCreditValue + cobjOrder.CSMPPCTAmt, "$###,###,##0.00")
        
        Case "OR.ADJGPPCT"
            If cobjOrder.ADJGPPCT <> 0 Then
                mstrReturn = Format(cobjOrder.ADJGPPCT / 100, "##0.00%")
            Else
                mstrReturn = "0.00%"
            End If
        Case "OR.BIDDESKLABEL"
            If Object.SysParms.biddeskusergroupid < 1 Or cobjOrder.orderid < 1 Then
                mstrReturn = ""
            Else
                If cobjOrder.BidDeskDTReqSRepReview <> "1/1/1900" And cobjOrder.BidDeskDTSRepReviewed = "1/1/1900" Then
                    mstrReturn = "Requested Rep Review (" & Format(cobjOrder.BidDeskDTReqSRepReview, "mm/dd/yy hh:mm ampm") & ")"
                End If
                If cobjOrder.BidDeskDTReqSRepReview <> "1/1/1900" And cobjOrder.BidDeskDTSRepReviewed <> "1/1/1900" Then
                    Select Case UCase(cobjOrder.BidDeskSRepStatus)
                        Case "A"
                            mstrReturn = "Accepted By Rep (" & Format(cobjOrder.BidDeskDTSRepReviewed, "mm/dd/yy hh:mm ampm") & ")"
                        Case "R"
                            mstrReturn = "Rejected By Rep (" & Format(cobjOrder.BidDeskDTSRepReviewed, "mm/dd/yy hh:mm ampm") & ")"
                        Case Else
                            mstrReturn = "Unrecognized Status (" & cobjOrder.BidDeskSRepStatus & ")"
                    End Select
                End If
            End If
        Case "OR.CHALLENGES"
            mstrReturn = cobjOrder.Challenges
        Case "OR.CUSTOMERNAME"
            mstrReturn = cobjOrder.CustomerName
        Case "OR.DTCLOSED"
            mstrReturn = Format(cobjOrder.DtClosed, "mm/dd/yyyy")
        Case "OR.EXPCLOSEDT"
            mstrReturn = Format(cobjOrder.Prospect.DTExpectedClose, "mm/dd/yyyy")
        Case "OR.CONTRACTEDTHRUDT"
            If cobjOrder.Prospect.CONTRACTEDTHRUDT <> "1/1/1900" Then
                mstrReturn = Format(cobjOrder.Prospect.CONTRACTEDTHRUDT, "mm/dd/yyyy")
            Else
                mstrReturn = ""
            End If
        Case "OR.DTDEMO"
            If cobjOrder.Prospect.DTDEMO <> "1/1/1900" Then
                mstrReturn = Format(cobjOrder.Prospect.DTDEMO, "mm/dd/yyyy")
            Else
                mstrReturn = ""
            End If
        Case "OR.DTPROPOSAL"
            If cobjOrder.Prospect.DTPROPOSAL <> "1/1/1900" Then
                mstrReturn = Format(cobjOrder.Prospect.DTPROPOSAL, "mm/dd/yyyy")
            Else
                mstrReturn = ""
            End If
        Case "OR.NEXTSTEP"
             mstrReturn = cobjOrder.Prospect.NEXTSTEP
        Case "OR.GOALS"
            mstrReturn = cobjOrder.Goals
        Case "OR.ORDERID"
            mstrReturn = CStr(cobjOrder.orderid)
        Case "OR.TYPEORDERID"
            If UCase(cobjOrder.OrderStatus.SystemLabel) = "PROP" Then
                mstrReturn = "P-" & CStr(cobjOrder.orderid)
            Else
                mstrReturn = CStr(cobjOrder.orderid)
            End If
        Case "OR.RECOMMENDATIONS"
            mstrReturn = cobjOrder.Recommendations
        Case "OR.TYPEOFOBJECT"
            If UCase(cobjOrder.OrderStatus.SystemLabel) = "PROP" Then
                mstrReturn = "Proposal"
            Else
                mstrReturn = "Order"
            End If
        Case "OR.CONTACTNAME"
            mstrReturn = FormatTitle(cobjOrder.ContactName)
            If InStr(mstrReturn, ",") > 0 Then
                mstrReturn = Trim(Mid(mstrReturn, InStr(mstrReturn, ",") + 1)) & " " & Trim(Mid(mstrReturn, 1, InStr(mstrReturn, ",") - 1))
            End If
        Case "OR.DEAR"
            mstrReturn = FormatTitle(cobjOrder.Dear)
        Case "OR.CONTACTEMAILADDR"
            mstrReturn = LCase(cobjOrder.ContactEmailAddr)
        Case "OR.CONTACTPHONENBR"
            If Len(cobjOrder.ContactPhoneNbr) > 0 Then
                mstrReturn = Format(cobjOrder.ContactPhoneNbr, "(###) ###-####")
            Else
                mstrReturn = ""
            End If
        Case "OR.CONTACTFAXNBR"
            If Len(cobjOrder.ContactFaxNbr) > 0 Then
                mstrReturn = Format(cobjOrder.ContactFaxNbr, "(###) ###-####")
            Else
                mstrReturn = ""
            End If
        Case "OR.BILLINGCUSTOMERNAME"
            mstrReturn = cobjOrder.BillingCustomerName
        Case "OR.BILLINGOMDCUSTOMERNBR"
            mstrReturn = cobjOrder.BillToCustomer.OMDCustomerNbr

        Case "OR.BILLINGCONTACTEMAILADDR"
            mstrReturn = LCase(cobjOrder.BillingContactEmailAddr)
        Case "OR.BILLINGADDRESSLABEL"
            mstrReturn = cobjOrder.BILLINGAddressLabel
        Case "OR.BILLINGADDRESS1"
            mstrReturn = cobjOrder.BILLINGAddress1
        Case "OR.BILLINGADDRESS2"
            mstrReturn = cobjOrder.BILLINGAddress2
        Case "OR.BILLINGCITY"
            mstrReturn = cobjOrder.BILLINGcity
        Case "OR.BILLINGSTATE"
            mstrReturn = cobjOrder.BILLINGState
        Case "OR.BILLINGZIP"
            mstrReturn = cobjOrder.BILLINGPostalCode
        Case "OR.BILLINGADDRESSONELINE"
            mstrReturn = cobjOrder.BillingAddressOneLine
        Case "OR.BILLINGCONTACTNAME"
            mstrReturn = FormatTitle(cobjOrder.BILLINGContactName)
            If InStr(mstrReturn, ",") > 0 Then
                mstrReturn = Trim(Mid(mstrReturn, InStr(mstrReturn, ",") + 1)) & " " & Trim(Mid(mstrReturn, 1, InStr(mstrReturn, ",") - 1))
            End If
        Case "OR.BILLINGCONTACTPHONENBR"
            If Len(cobjOrder.BillingContactPhoneNbr) > 0 Then
                mstrReturn = Format(cobjOrder.BillingContactPhoneNbr, "(###) ###-####")
            Else
                mstrReturn = ""
            End If
        Case "OR.BILLINGCONTACTFAXNBR"
            If Len(cobjOrder.BillingContactFaxNbr) > 0 Then
                mstrReturn = Format(cobjOrder.BillingContactFaxNbr, "(###) ###-####")
            Else
                mstrReturn = ""
            End If
        Case "OR.BILLINGCONTACTCELLNBR"
            If Len(cobjOrder.BillingContactCellNbr) > 0 Then
                mstrReturn = Format(cobjOrder.BillingContactCellNbr, "(###) ###-####")
            Else
                mstrReturn = ""
            End If
        Case "OR.FCOLEGALNAME"
            mstrReturn = Trim(cobjOrder.FCOLegalName)
            '9/15/2013 TES - Make sure that there is a name for this field.
            If mstrReturn = "" Then
                mstrReturn = Trim(cobjOrder.BillingCustomerName)
            End If
            If mstrReturn = "" Then
                mstrReturn = Trim(cobjOrder.BillToCustomer.Company.Name)
            End If
        Case "OR.UPGRADETYPE" 'JB Add
            mstrReturn = cobjOrder.UpgradeType.Name
        Case "OR.FCOAPPNBR"
            mstrReturn = cobjOrder.FCOAppNbr
        Case "OR.FCOAPPLDATE"
            If cobjOrder.FCOApplicationDT = "1/1/1900" Then
                mstrReturn = ""
            Else
                mstrReturn = Format(cobjOrder.FCOApplicationDT, "mm/dd/yyyy")
            End If
        Case "OR.FCOAPPAPPROVEDATE"
            If cobjOrder.FCOApprovedDT = "1/1/1900" Then
                mstrReturn = ""
            Else
                mstrReturn = Format(cobjOrder.FCOApprovedDT, "mm/dd/yyyy")
            End If
        Case "OR.FCOAPPRCONDITIONS"
            mstrReturn = cobjOrder.fcoConditions
        Case "OR.FCOAPPRVALIDTHRUDATE"
            If cobjOrder.FCOApprovalValidThruDT = "1/1/1900" Then
                mstrReturn = ""
            Else
                mstrReturn = Format(cobjOrder.FCOApprovalValidThruDT, "mm/dd/yyyy")
            End If
        
        Case "OR.ISOLINEADDENDUMNEEDED"
            'Provide a message if the order has more lines then a document is capable of displaying.
            If cobjOrder.OrderLines.Count > mIndex Then
                If Len(mstrAddendumname) > 0 Then
                    mstrReturn = "ADDENDUM (" & mstrAddendumname & ") IS REQUIRED"
                Else
                    mstrReturn = "ADDENDUM IS REQUIRED"
                End If
            Else
                mstrReturn = ""
            End If
        Case "OR.ISOLINEADDENDUMNEEDEDX"
            'Provide a message if the order has more lines then a document is capable of displaying.
            If cobjOrder.OrderLines.Count > mIndex Then
                mstrReturn = "X"
            Else
                mstrReturn = ""
            End If
        Case "OR.PRIMACHADDENDUMNEEDED"
            
            If cobjOrder.GetNumberOfPrimaryMachines > mIndex Then
                If Len(mstrAddendumname) > 0 Then
                    mstrReturn = "ADDENDUM (" & mstrAddendumname & ") IS REQUIRED"
                Else
                    mstrReturn = "ADDENDUM IS REQUIRED"
                End If
            Else
                mstrReturn = ""
            End If
        Case "OR.PRIMACHADDENDUMNEEDEDX"
            'Provide a message if the order has more primary lines then a document is capable of displaying.
            If cobjOrder.GetNumberOfPrimaryMachines > mIndex Then
                mstrReturn = "X"
            Else
                mstrReturn = ""
            End If
        Case "OR.SUMADJ"
            mstrReturn = mstrAddendumname
        Case "OR.MGRAPPROVEDLABEL"
            If cobjOrder.ApprovedByUser.UserID > 1 Then
                Select Case cobjOrder.ManagerApproved
                    Case 1
                        mstrReturn = "Approved By Manager(" & cobjOrder.ApprovedByUser.Person.Fullname & ")"
                    Case 2
                        mstrReturn = "Declined By Manager(" & cobjOrder.ApprovedByUser.Person.Fullname & ")"
                    Case 0
                        mstrReturn = "No Action By Manager(" & cobjOrder.ApprovedByUser.Person.Fullname & ")"
                End Select
            Else
                mstrReturn = ""
            End If
        Case "OR.TYPEOFOBJECT"
            If cobjOrder.OrderStatus.SystemLabel = "PROP" Then
                mstrReturn = "Proposal"
            Else
                mstrReturn = "Order"
            End If
        Case "OR.PRICELEVEL"
            mstrReturn = cobjOrder.PriceLevel.Name
        Case "OR.PRICEMETHOD"
            Select Case cobjOrder.PricingMethod
                Case 0
                    mstrReturn = "Point Book"
                Case 1
                    mstrReturn = "Promo/Special"
                Case 2
                    mstrReturn = "Bid Desk"
                Case 3
                    mstrReturn = "Branch Manager"
                Case 4
                    mstrReturn = "Double Contest"
                Case Else
                    mstrReturn = ""
            End Select
        Case "OR.PRICECONFORM"
            If cobjOrder.PricingMethod <> 0 Then
                mstrReturn = "Pricing is Non Conforming - "
            Else
                mstrReturn = ""
            End If
            
        Case "OR.NOTE"
            mstrReturn = cobjOrder.Note
        Case "OR.TERRITORY"
            mstrReturn = FormatTitle(cobjOrder.Territory.TerritoryName)
        'Case "OR.BRANCH"
        '    mstrReturn = FormatTitle(cobjOrder.Territory.Branch.Name)
        Case "OR.SALESMANAGERNAME"
            mstrReturn = cobjOrder.SalesRepUser.Branch.SalesManagerUser.Person.Fullname
        Case "OR.SALESREP"
            mstrReturn = cobjOrder.SalesRepUser.Person.Fullname
        Case "OR.SALESREPEMAIL"
            mstrReturn = LCase(cobjOrder.SalesRepUser.Person.emailaddress)
        Case "OR.SALESREPPHONE"
            mstrReturn = Format(cobjOrder.SalesRepUser.Person.OfficePhone.PhoneNbr, "(###) ###-####")
        Case "OR.SALESREPTITLE"
            mstrReturn = cobjOrder.SalesRepUser.Title
        Case "OR.SALESREPCELLPHONE"
            If Len(cobjUser.Person.CellPhone.PhoneNbr) > 0 Then
                mstrReturn = Format(cobjUser.Person.CellPhone.PhoneNbr, "(###) ###-####")
            Else
                mstrReturn = ""
            End If
        Case "OR.SALESREPXREF"
            mstrReturn = cobjOrder.SalesRepUser.Person.Entity.ImportXRef
        Case "OR.BUSINESS"
            mstrReturn = cobjOrder.SalesRepUser.Branch.Business.Name
        Case "OR.BUSINESSADDR1"
            mstrReturn = cobjOrder.SalesRepUser.Branch.Business.Address1
        Case "OR.BUSINESSADDR2"
            mstrReturn = cobjOrder.SalesRepUser.Branch.Business.Address2
        Case "OR.BUSINESSCITY"
            mstrReturn = cobjOrder.SalesRepUser.Branch.Business.City
        Case "OR.BUSINESSSTATE"
            mstrReturn = cobjOrder.SalesRepUser.Branch.Business.StateCode
        Case "OR.BUSINESSPOSTALCODE"
            mstrReturn = cobjOrder.SalesRepUser.Branch.Business.PostalCode
        Case "OR.BUSINESSPHONE"
            mstrReturn = cobjOrder.SalesRepUser.Branch.Business.Phone
        Case "OR.BUSINESSFAX"
            mstrReturn = cobjOrder.SalesRepUser.Branch.Business.PhoneFax
        Case "OR.BUSINESSIMAGESMALL"
            mstrReturn = cobjOrder.SalesRepUser.Branch.Business.ImageSmall
        Case "OR.BUSINESSIMAGELARGE"
            mstrReturn = cobjOrder.SalesRepUser.Branch.Business.ImageLarge
        Case "OR.BRANCH"
            mstrReturn = cobjOrder.SalesRepUser.Branch.BranchName
        Case "OR.BRANCHADDRESS"
            mstrReturn = cobjOrder.SalesRepUser.Branch.GetAddress
        Case "OR.BRANCHADDR1"
            mstrReturn = cobjOrder.SalesRepUser.Branch.Addr1
        Case "OR.BRANCHADDR2"
            mstrReturn = cobjOrder.SalesRepUser.Branch.Addr2
        Case "OR.BRANCHCITY"
            mstrReturn = cobjOrder.SalesRepUser.Branch.City
        Case "OR.BRANCHSTATE"
            mstrReturn = cobjOrder.SalesRepUser.Branch.State
        Case "OR.BRANCHPOSTALCODE"
            mstrReturn = cobjOrder.SalesRepUser.Branch.PostalCode
        Case "OR.BRANCHPHONE"
            mstrReturn = cobjOrder.SalesRepUser.Branch.Phone
        Case "OR.BRANCHFAX"
            mstrReturn = cobjOrder.SalesRepUser.Branch.Fax
        Case "OR.BRANCHABBREV"
            mstrReturn = cobjOrder.SalesRepUser.Branch.Abbreviation
        Case "OR.TYPEOFSALE"
            mstrReturn = FormatTitle(cobjOrder.SaleType.Name)
            
        Case "OR.PRESALESREPFULL"
            mstrReturn = cobjOrder.PreSaleAgentUser.Person.Fullname
        Case "OR.PRESALESREPABREV"
            mstrReturn = FormatTitle(cobjOrder.PreSaleAgentUser.Abbreviation)
        Case "OR.PRESALESREPEMAIL"
            mstrReturn = LCase(cobjOrder.PreSaleAgentUser.Person.emailaddress)
        Case "OR.PRESALESREPPHONE"
            mstrReturn = Format(cobjOrder.PreSaleAgentUser.Person.OfficePhone.PhoneNbr, "(###) ###-####")
        Case "OR.PRESALESREPXREF"
            mstrReturn = cobjOrder.PreSaleAgentUser.Person.Entity.ImportXRef
        
        Case "OR.SRCLEVELNAME"
            mstrReturn = FormatTitle(cobjOrder.SRCommLevel.Name)
        Case "OR.GPBOARDCREDITAMT"
            mstrReturn = Format(cobjOrder.GPBoardCreditAmt, "$###,###,##0.00")
        Case "OR.GPBASEAMT"
            mstrReturn = Format(cobjOrder.GPBaseAmt, "$###,###,##0.00")
        Case "OR.BASEGPAMT"
            mstrReturn = Format(cobjOrder.GPTotal, "$###,###,##0.00")
        Case "OR.SRCCOMMPCT"
            mstrReturn = Format(cobjOrder.SRCommPCTPaid, "##0.00%")
        
        Case "OR.SRCSEGMENTBONUSAMT"
            mstrReturn = Format(cobjOrder.SegmentBonusValue, "$###,###,##0.00")
        Case "OR.CSMPCREDITAMT"
            mstrReturn = Format(cobjOrder.SRCommCSMPCreditValue, "$###,###,##0.00")
        Case "OR.ADJCOSTAMT"
            'mstrReturn = Format(cobjOrder.Subtotalamount - cobjOrder.SRCommCSMPCreditValue, "$###,###,##0.00")
            mcurTotal = 0
            For mint = 1 To cobjOrder.OrderLines.Count
                mcurTotal = mcurTotal + (cobjOrder.OrderLines(mint).BuyPrice * cobjOrder.OrderLines(mint).Quantity)
            Next
            mstrReturn = Format(mcurTotal - cobjOrder.SRCommCSMPCreditValue, "$###,###,##0.00")
        Case "OR.CSMPCREDITNAME"
            If cobjOrder.SRComMCSMPLevelid = -100 Then
                mstrReturn = "Custom"
            ElseIf cobjOrder.SRComMCSMPLevelid = -102 Then
                mstrReturn = "Strategic"
            Else
                mstrReturn = cobjOrder.SRCommCSMPCredit.levelname
            End If
        Case "OR.CSMPCONTRACT"
            mstrReturn = cobjOrder.SRCommCSMPContractNbr
        Case "OR.CSMPDTAPPROVED"
            If cobjOrder.SRCommCSMPDTApproved = "1/1/1900" Then
                mstrReturn = ""
            Else
                mstrReturn = Format(cobjOrder.SRCommCSMPDTApproved, "mm/dd/yyyy")
            End If
        Case "OR.SRCSERVICEAMT"
            mstrReturn = Format(cobjOrder.SRCommServiceValue, "$###,###,##0.00")
        Case "OR.SRCSERVICERATE"
            mstrReturn = Format(cobjOrder.SRCommServiceRate, "##0.00%")
        
        Case "OR.SRCCOMMAMT"
            mstrReturn = Format(cobjOrder.SalesRepCommissionAmt, "$###,###,##0.00")
        Case "OR.SRCCOMMBONUS"
            mstrReturn = Format(cobjOrder.SRCommMaintBonus, "$###,###,##0.00")
         Case "OR.SRCCOMMSVCBONUS"
            mstrReturn = Format(cobjOrder.SRCommMaintService, "$###,###,##0.00")
        Case "OR.SRCTOTALCOMMAMT"
            'mstrReturn = Format(cobjOrder.SalesRepCommissionAmt + cobjOrder.SRCommMaintBonus, "$###,###,##0.00")
            mstrReturn = Format(cobjOrder.SRCommTotalAmt, "$###,###,##0.00")
        Case "OR.SRCSPLITREPNAME"
            mstrReturn = FormatTitle(cobjOrder.OtherSalesRepUser.Person.Fullname)
        Case "OR.SRCOTHERSALESREPPCT"
            If cobjOrder.OtherSalesRepPCT = 0 Then
                mstrReturn = ""
            Else
                mstrReturn = Format(cobjOrder.OtherSalesRepPCT / 100, "##0.00%")
            End If
        Case "OR.SRCPRIMSALESREPCOMMAMT"
            mstrReturn = Format(cobjOrder.PrimarySRCommissionAmt, "$###,###,##0.00")
        Case "OR.SRCOTHERSALESREPCOMMAMT"
            If cobjOrder.SplitSRCommissionAmt = 0 Then
                mstrReturn = ""
            Else
                mstrReturn = Format(cobjOrder.SplitSRCommissionAmt, "$###,###,##0.00")
            End If
        Case "OR.SRCISPRIMPAID"
            If cobjOrder.SRCommPaidPrimary = False Then
                mstrReturn = "No"
            Else
                mstrReturn = "Yes"
            End If
        Case "OR.SRCISOTHERPAID"
            If cobjOrder.SRCommPaidOther = False Then
                mstrReturn = "No"
            Else
                mstrReturn = "Yes"
            End If
        Case "OR.SRCDTPRIMPAIDSHORT"
            If cobjOrder.SRCommDTPaidPrimary = "1/1/1900" Then
                mstrReturn = ""
            Else
                mstrReturn = Format(cobjOrder.SRCommDTPaidPrimary, "mm/dd/yyyy")
            End If
        Case "OR.SRCDTPRIMPAIDLONG"
            If cobjOrder.SRCommDTPaidPrimary = "1/1/1900" Then
                mstrReturn = ""
            Else
                mstrReturn = Format(cobjOrder.SRCommDTPaidPrimary, "mmm d yyyy")
            End If
        Case "OR.SRCDTOTHERPAIDSHORT"
            If cobjOrder.SRCommDTPaidOther = "1/1/1900" Then
                mstrReturn = ""
            Else
                mstrReturn = Format(cobjOrder.SRCommDTPaidOther, "mm/dd/yyyy")
            End If
        Case "OR.SRCDTOTHERPAIDLONG"
            If cobjOrder.SRCommDTPaidOther = "1/1/1900" Then
                mstrReturn = ""
            Else
                mstrReturn = Format(cobjOrder.SRCommDTPaidOther, "mmm d yyyy")
            End If
        Case "OR.SRCSALESORDER"
            mstrReturn = cobjOrder.SalesOrderNbr
        Case "OR.SRCINVOICENBR"
            mstrReturn = cobjOrder.InvoiceNbr
        Case "OR.SRCNOTE"
            mstrReturn = cobjOrder.SRCommNote
        Case "OR.SRCISMINIMUMAPPLYN"
            If cobjOrder.SRCommMinimumAppliedInd = False Then
                mstrReturn = "No"
            Else
                mstrReturn = "Yes"
            End If
        Case "OR.SRCISMINIMUMAPPLMSG"
            If cobjOrder.SRCommMinimumAppliedInd = False Then
                mstrReturn = ""
            Else
                mstrReturn = "Minimum Applied"
            End If
        Case "OR.SUBTOTALSELLPRICE"
            mstrReturn = Format(cobjOrder.Subtotalamount, "$###,###,##0.00")
            
        Case "OR.SUBTOTALPLUSTAXABLEADJ"
            mcurTotal = cobjOrder.Subtotalamount
            For mint = 1 To cobjOrder.OrderAdjustments.Count
                If cobjOrder.OrderAdjustments.Item(mint).OrderAdjType.Taxable = True Then
                    If cobjOrder.OrderAdjustments.Item(mint).OrderAdjType.CreditDebitInd = "C" Then
                        mcurTotal = mcurTotal - cobjOrder.OrderAdjustments.Item(mint).AdjAmount
                    Else
                        mcurTotal = mcurTotal + cobjOrder.OrderAdjustments.Item(mint).AdjAmount
                    End If
                End If
            Next
            mstrReturn = Format(mcurTotal, "$###,##0.00")
        
        Case "OR.SUBACCUMPTS"
            mstrReturn = Format(cobjOrder.GetSubTotalAccumPts, "###,##0.00")
        Case "OR.SUBCONPTS"
            mstrReturn = Format(cobjOrder.GetSubTotalContestPts, "###,##0.00")
        Case "OR.SUBPAYPTS"
            mstrReturn = Format(cobjOrder.GetSubTotalPayPts, "###,##0.00")
        Case "OR.TOTALADJCOST"
            mcurTotal = 0
            For mint = 1 To cobjOrder.OrderLines.Count
                mcurTotal = mcurTotal + (cobjOrder.OrderLines(mint).Quantity * cobjOrder.OrderLines(mint).AdjDealerCostAmt)
            Next
            mstrReturn = Format(mcurTotal, "$###,###,##0.00")
        Case "OR.TOTALMSRP"
            mcurTotal = 0
            For mint = 1 To cobjOrder.OrderLines.Count
                mcurTotal = mcurTotal + (cobjOrder.OrderLines(mint).Quantity * cobjOrder.OrderLines(mint).CatalogItem.MFGPrice)
            Next
            mstrReturn = Format(mcurTotal, "$###,###,##0.00")
        Case "OR.TOTALMSRPPCTFIN"
            mcurTotal = 0
            For mint = 1 To cobjOrder.OrderLines.Count
                mcurTotal = mcurTotal + (cobjOrder.OrderLines(mint).Quantity * cobjOrder.OrderLines(mint).CatalogItem.MFGPrice)
            Next
            If cobjOrder.AmtFinanced <> 0 And mcurTotal <> 0 Then
                mstrReturn = Format(cobjOrder.AmtFinanced / mcurTotal, "###,##0%")
            Else
                mstrReturn = Format(mcurTotal, "$###,###,##0.00")
            End If
        
        Case "OR.MANUALENTRYLABEL"
            mstrReturn = ""
            For mint = 1 To cobjOrder.OrderLines.Count
                If cobjOrder.OrderLines(mint).CatalogItemID <= -100 Then
                    mstrReturn = "ORDER CONTAINS MANUALLY ENTERED PRODUCTS"
                    Exit For
                End If
            Next
        Case "OR.SoftCostAmt"
            mstrReturn = Format(cobjOrder.SoftCostAmt, "$###,###,##0.00")
        Case "OR.SoftCostPCT"
            mstrReturn = Format(cobjOrder.SoftCostPCT, "##0.00%")
        Case "OR.DEALERCOST"
            mstrReturn = Format(cobjOrder.DealerCostAmt, "$###,###,##0.00")
        Case "OR.TOTALSELLPRICE"
            mstrReturn = Format(cobjOrder.totalOrderAmount, "$###,###,##0.00")
        Case "OR.TOTALSELLNOTAX"
            mcurTotal = cobjOrder.totalOrderAmount
            For mint = 1 To cobjOrder.OrderAdjustments.Count
                If cobjOrder.OrderAdjustments.Item(mint).OrderAdjType.TaxInd = True Then
                    mcurTotal = mcurTotal - cobjOrder.OrderAdjustments.Item(mint).AdjAmount
                End If
            Next
            mstrReturn = Format(mcurTotal, "$###,###,##0.00")
        Case "OR.TOTALACCUMPTS"
            mstrReturn = Format(cobjOrder.TotalAccumPoints, "###,##0.00")
        Case "OR.TOTALCONPTS"
            mstrReturn = Format(cobjOrder.TotalContestPoints, "###,##0.00")
        Case "OR.TOTALPAYPTS"
            mstrReturn = Format(cobjOrder.TotalPayPoints, "###,##0.00")
        Case "OR.PICKUPINFO"
            mstrReturn = cobjOrder.GetOrderItemsPickupInfo
        Case "OR.REASONFORTECH"
            If Len(cobjOrder.ReasonForServiceTech) = 0 Then
                mstrReturn = "No Reason Given"
            Else
                mstrReturn = cobjOrder.ReasonForServiceTech
            End If
        Case "OR.DELIVERYINSTRUCTIONS"
            If Len(cobjOrder.DeliveryInstructions) = 0 Then
                mstrReturn = "Incomplete"
            Else
                mstrReturn = cobjOrder.DeliveryInstructions
            End If
        Case "OR.REPDELIVERY"
            If cobjOrder.RepDeliveryInd = True Then
                mstrReturn = "Yes"
            Else
                mstrReturn = "No"
            End If
        Case "OR.REPDELIVERYTEXT"
            If cobjOrder.RepDeliveryInd = True Then
                mstrReturn = "Sales Rep"
            Else
                mstrReturn = ""
            End If
        Case "OR.NETWORKINGIND"
            If cobjOrder.NetworkingInd = True Then
                mstrReturn = "Yes"
            Else
                mstrReturn = "No"
            End If
        Case "OR.TRUCKDELIVERY"
            If cobjOrder.TruckDeliveryInd = True Then
                mstrReturn = "Yes"
            Else
                mstrReturn = "No"
            End If
        Case "OR.TRUCKDELIVERYTEXT"
            If cobjOrder.TruckDeliveryInd = True Then
                mstrReturn = "Truck"
            Else
                mstrReturn = ""
            End If
        Case "OR.PICKUPDELIVERY"
            If cobjOrder.PickupDeliveryInd = True Then
                mstrReturn = "Yes"
            Else
                mstrReturn = "No"
            End If
        Case "OR.RELOCATE"
            If cobjOrder.RelocateInd = True Then
                mstrReturn = "Yes"
            Else
                mstrReturn = "No"
            End If
        Case "OR.DTPICKUP"
            If cobjOrder.DTPickup <> "1/1/1900" Then
                mstrReturn = Format(cobjOrder.DTPickup, "mm/dd/yyyy")
            Else
                mstrReturn = ""
            End If
        Case "OR.STAIRS"
            If cobjOrder.DeliveryStairsInd = True Then
                mstrReturn = "Yes"
            Else
                mstrReturn = "No"
            End If
        Case "OR.NBRSTAIRS"
            mstrReturn = Format(cobjOrder.DeliveryNbrStairs, "##0")
        Case "OR.DTCONVERSION"
            If cobjOrder.DTPickup <> "1/1/1900" Then
                mstrReturn = Format(cobjOrder.DTConversion, "mm/dd/yyyy")
            Else
                mstrReturn = ""
            End If
        Case "OR.NEWCUSTOMER"
            If cobjOrder.NewCustomerInd = True Then
                mstrReturn = "Yes"
            Else
                mstrReturn = "No"
            End If
        Case "OR.MPSIND"
            If cobjOrder.IsMPSInd = True Then
                mstrReturn = "Yes"
            Else
                mstrReturn = "No"
            End If
        Case "OR.MNSIND"
            If cobjOrder.IsMNSInd = True Then
                mstrReturn = "Yes"
            Else
                mstrReturn = "No"
            End If
        Case "OR.TECHREQUESTED"
            If cobjOrder.TechRequestedInd = True Then
                mstrReturn = "Yes"
            Else
                mstrReturn = "No"
            End If
        Case "OR.CSRREQUESTED"
            If cobjOrder.CSRRequestedInd = True Then
                mstrReturn = "Yes"
            Else
                mstrReturn = "No"
            End If
        Case "OR.DIGINSTREQUESTED"
            If cobjOrder.DigInstRequestedInd = True Then
                mstrReturn = "Yes"
            Else
                mstrReturn = "No"
            End If
        Case "OR.UPDATECONNECTEDEQ"
            If cobjOrder.UpdateConnectedEq = True Then
                mstrReturn = "Yes"
            Else
                mstrReturn = "No"
            End If
        Case "OR.FAXSETUP"
            If cobjOrder.FAXSetup = True Then
                mstrReturn = "Yes"
            Else
                mstrReturn = "No"
            End If
        Case "OR.STATEPRICING"
            If cobjOrder.StatePricingInd = True Then
                mstrReturn = "Yes"
            Else
                mstrReturn = "No"
            End If
        Case "OR.STATEPRICINGSTATE"
            mstrReturn = cobjOrder.StateCode
        Case "OR.CURRENTCUSTBUSINESS"
            mstrReturn = cobjOrder.CurrentCustBusiness.Name
        Case "OR.SPLITORDER"
            If cobjOrder.SplitOrderInd = True Then
                mstrReturn = "Yes"
            Else
                mstrReturn = "No"
            End If
        Case "OR.OTHERSALESREP"
            mstrReturn = cobjOrder.OtherSalesRepUser.Person.Fullname
        Case "OR.OTHERSALESREPPCT"
            mstrReturn = Format(cobjOrder.OtherSalesRepPCT, "percent")
        Case "OR.INVOICELINEDETAIL"
            If cobjOrder.InvoiceLineDetailInd = True Then
                mstrReturn = "Yes"
            Else
                mstrReturn = "No"
            End If
        Case "OR.TONERINCLUSIVE"
            If cobjOrder.TonerInclusiveInd = True Then
                mstrReturn = "Yes"
            Else
                mstrReturn = "No"
            End If
        Case "OR.CONVERSIONTYPE"
            mstrReturn = cobjOrder.ConversionType.Name
        Case "OR.DTORIGORDER"
            If cobjOrder.DTOrigOrder <> "1/1/1900" Then
                mstrReturn = Format(cobjOrder.DTOrigOrder, "mm/dd/yyyy")
            Else
                mstrReturn = ""
            End If
        Case "OR.RENTALEQUITY"
            mstrReturn = Format(cobjOrder.RentalEquity, "$###,###,##0.00")
        Case "OR.FINALMETER"
            mstrReturn = Format(cobjOrder.FinalMeterCount, "###,###,###,##0")
        Case "OR.ACCESSORYONLY"
            If cobjOrder.AccessoryOnlyInd = True Then
                mstrReturn = "Yes"
            Else
                mstrReturn = "No"
            End If
        Case "OR.ACCESSORIES"
            mstrReturn = cobjOrder.Accessories
            
        Case "OR.LEASEFACTOR"
            If cobjOrder.LeaseFactor = 0 Then
                mstrReturn = ""
            Else
                mstrReturn = Format(cobjOrder.LeaseFactor, "##0.00000")
            End If
            If cobjOrder.SaleType.isleased = False And cobjOrder.SaleTypeID > 0 Then
                mstrReturn = "n/a"
            End If
        
        Case "OR.LEASEPRODUCT"
            mstrReturn = cobjOrder.LeaseProduct.Name
            If cobjOrder.SaleType.isleased = False And cobjOrder.SaleTypeID > 0 Then
                mstrReturn = "n/a"
            End If
            
        Case "OR.EOTFMVX"
            If cobjOrder.LeaseProduct.isEOT_FMV = True Then
                mstrReturn = "X"
            Else
                mstrReturn = ""
            End If
        Case "OR.EOT$OUTX"
            If cobjOrder.LeaseProduct.isEOT_BuckOut = True Then
                mstrReturn = "X"
            Else
                mstrReturn = ""
            End If
        Case "OR.EOTFIXPCTX"
            If cobjOrder.LeaseProduct.isEOT_FixedPCT = True Then
                mstrReturn = "X"
            Else
                mstrReturn = ""
            End If
        Case "OR.EOTFIXEDPCT"
            If cobjOrder.LeaseProduct.isEOT_FixedPCT = True Then
                mstrReturn = Format(cobjOrder.LeaseProduct.EOT_FixedPCT, "##0.00%")
            Else
                mstrReturn = ""
            End If
        
        Case "OR.PYMTFREQLONG"
            mstrReturn = GetFreqLong(cobjOrder.LeaseProduct.PaymentFrequency)
        
        Case "OR.PYMTFREQSHORT"
            mstrReturn = GetFreqShort(cobjOrder.LeaseProduct.PaymentFrequency)
        
        Case "OR.PYMTFREQNBRPERYR"
            mstrReturn = CStr(cobjOrder.LeaseProduct.PaymentFrequency)
        
        Case "OR.PYMTFREQMOX"
            If cobjOrder.LeaseProduct.PaymentFrequency = 12 Then
                mstrReturn = "X"
            Else
                mstrReturn = ""
            End If
        
        Case "OR.PYMTFREQQTRX"
            If cobjOrder.LeaseProduct.PaymentFrequency = 4 Then
                mstrReturn = "X"
            Else
                mstrReturn = ""
            End If
        Case "OR.PYMTFREQSAX"
            If cobjOrder.LeaseProduct.PaymentFrequency = 2 Then
                mstrReturn = "X"
            Else
                mstrReturn = ""
            End If
        Case "OR.PYMTFREQANX"
            If cobjOrder.LeaseProduct.PaymentFrequency = 1 Then
                mstrReturn = "X"
            Else
                mstrReturn = ""
            End If
        
        Case "OR.BASEBILLINLSYESX"
            If cobjOrder.BillingCycleType.IncludeInLease = True Then
                mstrReturn = "X"
            Else
                mstrReturn = ""
            End If
        
        Case "OR.BASEBILLINLSNOX"
            If cobjOrder.BillingCycleType.IncludeInLease = False Then
                mstrReturn = "X"
            Else
                mstrReturn = ""
            End If
        
     'Bill Cycle check box for contract
        Case "OR.BASEBILLFREQMOX" 'JB Update to return monthly check on any bill cycle that has 12 payments per year
                                  'Removed below condition because it was not checking pymt freq box on pass throughs
            'If cobjOrder.BillingCycleType.PaymentsPerYear = 12 And UCase(cobjOrder.BillingCycleType.Name) = "MONTH" Then
            
            If cobjOrder.BillingCycleType.PaymentsPerYear = 12 Then
                mstrReturn = "X"
            Else
                mstrReturn = ""
            End If
        Case "OR.BASEBILLFREQQTRX"
            If cobjOrder.BillingCycleType.PaymentsPerYear = 4 Then
                mstrReturn = "X"
            Else
                mstrReturn = ""
            End If
        Case "OR.BASEBILLFREQSAX"
            If cobjOrder.BillingCycleType.PaymentsPerYear = 2 Then
                mstrReturn = "X"
            Else
                mstrReturn = ""
            End If
        Case "OR.BASEBILLFREQANX"
            If cobjOrder.BillingCycleType.PaymentsPerYear = 1 Then
                mstrReturn = "X"
            Else
                mstrReturn = ""
            End If
            
            
        'Overage base bill check boxes
            
            Case "OR.OVEBILLFREQMOX"
            If (cobjOrder.MaintOverBillingCycle = -1) And cobjOrder.BillingCycleType.PaymentsPerYear = 12 Then  'overage bill cycle same as base
              mstrReturn = "X"
            ElseIf cobjOrder.OverBillingCycleType.PaymentsPerYear = 12 Then
                mstrReturn = "X"
            Else
                mstrReturn = ""
            End If
        Case "OR.OVEBILLFREQQTRX"
            If (cobjOrder.MaintOverBillingCycle = -1) And cobjOrder.BillingCycleType.PaymentsPerYear = 4 Then  'overage bill cycle same as base
              mstrReturn = "X"
            ElseIf cobjOrder.OverBillingCycleType.PaymentsPerYear = 4 Then
                mstrReturn = "X"
            Else
                mstrReturn = ""
            End If
        Case "OR.OVEBILLFREQSAX"
            If (cobjOrder.MaintOverBillingCycle = -1) And cobjOrder.BillingCycleType.PaymentsPerYear = 2 Then  'overage bill cycle same as base
              mstrReturn = "X"
            ElseIf cobjOrder.OverBillingCycleType.PaymentsPerYear = 2 Then
                mstrReturn = "X"
            Else
                mstrReturn = ""
            End If
        Case "OR.OVEBILLFREQANX"
            If (cobjOrder.MaintOverBillingCycle = -1) And cobjOrder.BillingCycleType.PaymentsPerYear = 1 Then  'overage bill cycle same as base
              mstrReturn = "X"
            ElseIf cobjOrder.OverBillingCycleType.PaymentsPerYear = 1 Then
                mstrReturn = "X"
            Else
                mstrReturn = ""
            End If
        
        
        Case "OR.LEASEPRICELEVEL"
            mstrReturn = cobjOrder.LeasePriceLevel.Name
            If cobjOrder.SaleType.isleased = False And cobjOrder.SaleTypeID > 0 Then
                mstrReturn = "n/a"
            End If
        Case "OR.AMTFINANCED"
            If cobjOrder.AmtFinanced = 0 Then
                mstrReturn = ""
            Else
                mstrReturn = Format(cobjOrder.AmtFinanced, "$###,###,##0.00")
            End If
            If cobjOrder.SaleType.isleased = False And cobjOrder.SaleTypeID > 0 Then
                mstrReturn = "n/a"
            End If
        Case "OR.AMTFIN-LBO"
            If cobjOrder.AmtFinanced = 0 Then
                mstrReturn = ""
            Else
                mstrReturn = Format(cobjOrder.AmtFinanced - cobjOrder.buyoutamt, "$###,###,##0.00")
            End If
        Case "OR.BUYOUTAMT"
            If cobjOrder.buyoutamt = 0 Then
                mstrReturn = ""
            Else
                mstrReturn = Format(cobjOrder.buyoutamt, "$###,###,##0.00")
            End If
        Case "OR.FCOCODENAME"
            If cobjOrder.FCOID < 1 Then
                mstrReturn = ""
            Else
                mstrReturn = cobjOrder.FCO.CodeName
            End If
            If cobjOrder.SaleType.isleased = False And cobjOrder.SaleTypeID > 0 Then
                mstrReturn = "n/a"
            End If
        Case "OR.FCONAME"
            If cobjOrder.FCOID < 1 Then
                mstrReturn = ""
            Else
                mstrReturn = cobjOrder.FCO.Company.Name
            End If
            If cobjOrder.SaleType.isleased = False And cobjOrder.SaleTypeID > 0 Then
                mstrReturn = "n/a"
            End If
        Case "OR.NBRDEFPYMTS"
            If cobjOrder.NbrDeferredPymts = 0 Then
                mstrReturn = ""
            Else
                mstrReturn = Format(cobjOrder.NbrDeferredPymts, "##0")
            End If
            If cobjOrder.SaleType.isleased = False And cobjOrder.SaleTypeID > 0 Then
                mstrReturn = "n/a"
            End If
        Case "OR.NBRDEFPYMTSYESX"
            If cobjOrder.NbrDeferredPymts = 0 Then
                mstrReturn = ""
            Else
                mstrReturn = "X"
            End If
            If cobjOrder.SaleType.isleased = False And cobjOrder.SaleTypeID > 0 Then
                mstrReturn = ""
            End If
        Case "OR.NBRDEFPYMTSNOX"
            If cobjOrder.NbrDeferredPymts = 0 Then
                mstrReturn = "X"
            Else
                mstrReturn = ""
            End If
            If cobjOrder.SaleType.isleased = False And cobjOrder.SaleTypeID > 0 Then
                mstrReturn = ""
            End If
        Case "OR.NBRLEASEPYMTS"
            If cobjOrder.NbrLeasePymts = 0 Then
                mstrReturn = ""
            Else
                mstrReturn = Format(cobjOrder.NbrLeasePymts, "##0")
            End If
            If cobjOrder.SaleType.isleased = False And cobjOrder.SaleTypeID > 0 Then
                mstrReturn = "n/a"
            End If
        Case "OR.NBRPYMTS+NBRDEF"
            If cobjOrder.NbrLeasePymts + cobjOrder.NbrDeferredPymts = 0 Then
                mstrReturn = ""
            Else
                mstrReturn = Format(cobjOrder.NbrLeasePymts + cobjOrder.NbrDeferredPymts, "##0")
            End If
            If cobjOrder.SaleType.isleased = False And cobjOrder.SaleTypeID > 0 Then
                mstrReturn = "n/a"
            End If
        Case "OR.NBRPYMTSTBLW@"
            If cobjOrder.NbrDeferredPymts = 0 Then
                If cobjOrder.PaymentNbrFin = 0 Then
                    mstrReturn = ""
                Else
                    mstrReturn = Format(cobjOrder.NbrLeasePymts, "##0")
                End If
            Else
                    mstrReturn = Format(cobjOrder.NbrDeferredPymts, "##0") & " @" & vbCrLf & Format(cobjOrder.PaymentNbrFin, "##0") & " @"
            End If
            If cobjOrder.SaleType.isleased = False And cobjOrder.SaleTypeID > 0 Then
                mstrReturn = "n/a"
            End If
        Case "OR.PYMTSTBL"
            If cobjOrder.NbrDeferredPymts = 0 Then
                If cobjOrder.PaymentAmtMonthlyFin = 0 Then
                    mstrReturn = ""
                Else
                    mstrReturn = Format(cobjOrder.PaymentAmtMonthlyFin, "$###,###,##0.00")
                End If
            Else
                    mstrReturn = "$0.00" & vbCrLf & Format(cobjOrder.PaymentAmtMonthlyFin, "$###,###,##0.00")
            End If
            If cobjOrder.SaleType.isleased = False And cobjOrder.SaleTypeID > 0 Then
                mstrReturn = "n/a"
            End If
        Case "OR.PAYMENTNBRFIN"
            If cobjOrder.PaymentNbrFin = 0 Then
                mstrReturn = ""
            Else
                mstrReturn = Format(cobjOrder.PaymentNbrFin, "##0")
            End If
            If cobjOrder.SaleType.isleased = False And cobjOrder.SaleTypeID > 0 Then
                mstrReturn = "n/a"
            End If
        Case "OR.PAYMENTNBRFINMINUSONE" 'Term minus one to break out first payment
            If cobjOrder.PaymentNbrFin = 0 Then
                mstrReturn = ""
            Else
                mstrReturn = Format((cobjOrder.PaymentNbrFin - 1), "##0")
            End If
            If cobjOrder.SaleType.isleased = False And cobjOrder.SaleTypeID > 0 Then
                mstrReturn = "n/a"
            End If
        Case "OR.OVERRIDELEASETERM"
            If cobjOrder.OverrideLeaseTerm = 0 Then
                mstrReturn = ""
            Else
                mstrReturn = Format(cobjOrder.OverrideLeaseTerm, "##0")
            End If
            If cobjOrder.SaleType.isleased = False And cobjOrder.SaleTypeID > 0 Then
                mstrReturn = "n/a"
            End If
        Case "OR.PAYMENTAMTMONTHLYFIN"
            If cobjOrder.PaymentAmtMonthlyFin = 0 Then
                mstrReturn = ""
            Else
                mstrReturn = Format(cobjOrder.PaymentAmtMonthlyFin, "$###,###,##0.00")
            End If
            If cobjOrder.SaleType.isleased = False And cobjOrder.SaleTypeID > 0 Then
                mstrReturn = "n/a"
            End If
        Case "OR.PS24MOPYMT"
            If cobjOrder.PS24MoPymt = 0 Then
                mstrReturn = ""
            Else
                mstrReturn = Format(cobjOrder.PS24MoPymt, "$###,###,##0.00")
            End If
            If cobjOrder.SaleType.isleased = False And cobjOrder.SaleTypeID > 0 Then
                mstrReturn = "n/a"
            End If
        Case "OR.PS36MOPYMT"
            If cobjOrder.PS36MoPymt = 0 Then
                mstrReturn = ""
            Else
                mstrReturn = Format(cobjOrder.PS36MoPymt, "$###,###,##0.00")
            End If
            If cobjOrder.SaleType.isleased = False And cobjOrder.SaleTypeID > 0 Then
                mstrReturn = "n/a"
            End If
        Case "OR.PS48MOPYMT"
            If cobjOrder.PS48MoPymt = 0 Then
                mstrReturn = ""
            Else
                mstrReturn = Format(cobjOrder.PS48MoPymt, "$###,###,##0.00")
            End If
            If cobjOrder.SaleType.isleased = False And cobjOrder.SaleTypeID > 0 Then
                mstrReturn = "n/a"
            End If
        Case "OR.PS60MOPYMT"
            If cobjOrder.PS60MoPymt = 0 Then
                mstrReturn = ""
            Else
                mstrReturn = Format(cobjOrder.PS60MoPymt, "$###,###,##0.00")
            End If
            If cobjOrder.SaleType.isleased = False And cobjOrder.SaleTypeID > 0 Then
                mstrReturn = "n/a"
            End If
        Case "OR.BASELEASEPAYMENT"
            mstrReturn = Format(cobjOrder.AmtFinanced * cobjOrder.LeaseFactor, "$###,###,##0.00")
            If cobjOrder.SaleType.isleased = False And cobjOrder.SaleTypeID > 0 Then
                mstrReturn = "n/a"
            End If
        Case "OR.PAYMENTWITHTAX"
            mcurTotal = cobjOrder.PaymentAmtMonthlyFin
            If mcurTotal = 0 Then
                mstrReturn = ""
            Else
                mcurTotal = mcurTotal * (1 + cobjOrder.GetTaxRate)
                mstrReturn = Format(mcurTotal, "$###,###,##0.00")
            End If
            If cobjOrder.SaleType.isleased = False And cobjOrder.SaleTypeID > 0 Then
                mstrReturn = "n/a"
            End If
        Case "OR.TOTALSERVICECONTRACT" ' total service contract value
            
            If IsNull(cobjOrder.MaintTotalValue) = True Then
              mcurTotal = 0
            Else
              mcurTotal = cobjOrder.MaintTotalValue
            End If
            mstrReturn = Format(mcurTotal, "$###,###,##0.00")
        Case "OR.PAYMENTWITHSERVICE"
            mcurTotal = cobjOrder.PaymentAmtMonthlyFin
            If cobjOrder.BillingCycleType.IncludeInLease = True Then
                mcurTotal = mcurTotal + cobjOrder.MaintBasePayment
            End If
            If mcurTotal = 0 Then
                mstrReturn = ""
            Else
                mstrReturn = Format(mcurTotal, "$###,###,##0.00")
            End If
            If cobjOrder.SaleType.isleased = False And cobjOrder.SaleTypeID > 0 Then
                mstrReturn = "n/a"
            End If
        Case "OR.PAYMENTPLUSSERVICE"
            mcurTotal = cobjOrder.PaymentAmtMonthlyFin
            mcurTotal = mcurTotal + cobjOrder.MaintBasePayment
            mstrReturn = Format(mcurTotal, "$###,###,##0.00")
            If cobjOrder.SaleType.isleased = False And cobjOrder.SaleTypeID > 0 Then
                mstrReturn = "n/a"
            End If
        Case "OR.PYMTWSERVWPM"
            mcurTotal = cobjOrder.PaymentAmtMonthlyFin
            If cobjOrder.BillingCycleType.IncludeInLease = True Then
                mcurTotal = mcurTotal - cobjOrder.MaintBasePayment ' subtract maintenance - we do not want it
            End If
            For mint = 1 To cobjOrder.OrderLines.Count
                mcurTotal = mcurTotal + (cobjOrder.OrderLines(mint).BundleQuantity * cobjOrder.OrderLines(mint).MonthlyMeter)
            Next
            If mcurTotal = 0 Then
                mstrReturn = ""
            Else
                mstrReturn = Format(mcurTotal, "$###,###,##0.00")
            End If
        Case "OR.PYMTPLUSSERVPLUSPM" 'Payment will include service if bill cycle is pass through
            mcurTotal = cobjOrder.PaymentAmtMonthlyFin 'already includes maintenance
            For mint = 1 To cobjOrder.OrderLines.Count
                mcurTotal = mcurTotal + (cobjOrder.OrderLines(mint).BundleQuantity * cobjOrder.OrderLines(mint).MonthlyMeter)
            Next
            mstrReturn = Format(mcurTotal, "$###,###,##0.00")
        Case "OR.CUSTOMERPO"
            mstrReturn = cobjOrder.CustomerPO
        Case "OR.UPGRADEYESX"
            If cobjOrder.LeaseUpgradeInd = True Then
                mstrReturn = "X"
            Else
                mstrReturn = ""
            End If
        Case "OR.UPGRADENOX"
            If cobjOrder.LeaseUpgradeInd = False Then
                mstrReturn = "X"
            Else
                mstrReturn = ""
            End If
        Case "OR.UPGRADEYESNO"
            If cobjOrder.LeaseUpgradeInd = True Then
                mstrReturn = "Yes"
            Else
                mstrReturn = "No"
            End If
        Case "OR.UPGRADETYPE"
            If cobjOrder.LeaseUpgradeInd = True Then
                mstrReturn = cobjOrder.UpgradeType.Name
            Else
                mstrReturn = ""
            End If
        Case "OR.UPGRADELSNBR"
            If cobjOrder.LeaseUpgradeInd = True Then
                mstrReturn = cobjOrder.UpgradeFCOLeaseNbr
            Else
                mstrReturn = ""
            End If
        Case "OR.UPGRADEFCONAME"
            If cobjOrder.LeaseUpgradeInd = True Then
                mstrReturn = cobjOrder.UpgradeFCO.Company.Name
            Else
                mstrReturn = ""
            End If
        Case "OR.DTDEMOSTART"
            If cobjOrder.DTDemoStart <> "1/1/1900" Then
                mstrReturn = Format(cobjOrder.DTDemoStart, "mm/dd/yyyy")
            Else
                mstrReturn = ""
            End If
        Case "OR.DTDEMOEND"
            If cobjOrder.DTDemoEnd <> "1/1/1900" Then
                mstrReturn = Format(cobjOrder.DTDemoEnd, "mm/dd/yyyy")
            Else
                mstrReturn = ""
            End If
        Case "OR.PAYMENTNBRRENTAL"
            mstrReturn = Format(cobjOrder.PaymentNbrRental, "##0")
        Case "OR.PAYMENTAMTMONTHLYRENTAL"
            mstrReturn = Format(cobjOrder.PaymentAmtMonthlyRental, "$###,###,##0.00")
        Case "OR.COPYNBRRENTAL"
            mstrReturn = Format(cobjOrder.CopyNbrRental, "##0")
        Case "OR.OVERRAGERATERENTAL"
            mstrReturn = Format(cobjOrder.OverrageRateRental, "##0.00000")
        Case "OR.PURCHASETERMS"
            mstrReturn = cobjOrder.PurchaseTerms
        Case "OR.GROSSREVENUE"
            If cobjOrder.LeaseFactor = 0 Then
                mstrReturn = ""
            Else
                mstrReturn = Format(cobjOrder.GrossRevenue, "$###,###,##0.00")
            End If
        Case "OR.BUYTOTAL"
            mstrReturn = Format(cobjOrder.BuyTotal, "$###,###,##0.00")
        Case "OR.BUYTOTALPCT" ' JB Give full cost for CSMP PCT deals
            mstrReturn = Format((cobjOrder.BuyTotal + cobjOrder.CSMPPCTAmt), "$###,###,##0.00")
        Case "OR.NETPROFIT"
            mstrReturn = Format(cobjOrder.GrossRevenue - cobjOrder.BuyTotal - cobjOrder.GetTotalsByType("ALL")("TOTAL"), "$###,###,##0.00")
        Case "OR.GROSSPROFIT"
            If cobjOrder.GrossRevenue > 0 Then
                mstrReturn = Format((cobjOrder.GrossRevenue - cobjOrder.BuyTotal - cobjOrder.GetTotalsByType("ALL")("TOTAL")) / cobjOrder.GrossRevenue, "percent")
            Else
                mstrReturn = ""
            End If
        Case "OR.GROSSPROFITAMOUNT"
            mstrReturn = Format(cobjOrder.SellTotalHardware - cobjOrder.BuyTotal, "$###,###,##0.00")
        Case "OR.GROSSPROFITAMOUNTNF"
            mstrReturn = Format(cobjOrder.SellTotalHardware - cobjOrder.BuyTotal, "########0")
        Case "OR.GROSSREVENUECASH"
            If cobjOrder.LeaseFactor = 0 Then
                mstrReturn = Format(cobjOrder.GrossRevenue, "$###,###,##0.00")
            Else
                mstrReturn = ""
            End If
        Case "OR.MAINTENANCESUBTOTAL"
            mstrReturn = Format(cobjOrder.GetMaintenanceSubTotal, "$###,###,##0.00")
        Case "OR.TAXRATE"
            mstrReturn = Format(cobjOrder.GetTaxRate, "percent")
        Case "OR.MAINTENANCETAX"
            mstrReturn = Format(cobjOrder.GetMaintenanceSubTotal * cobjOrder.GetMaintenanceTaxRate, "$###,###,##0.00")
        Case "OR.MAINTENANCETOTAL"
            mstrReturn = Format(cobjOrder.GetMaintenanceSubTotal * (1 + cobjOrder.GetMaintenanceTaxRate), "$###,###,##0.00")
        Case "OR.MAINTCONTRACTTYPE"
            mstrReturn = cobjOrder.ContractType.Name
        Case "OR.MAINTCONTRACTDESC"
            mstrReturn = cobjOrder.ContractType.Description
        Case "OR.MAINTBILLINGCYCLE"
            mstrReturn = cobjOrder.BillingCycleType.Name
        Case "OR.MAINTBILLINGCYCLEOTHER"
            mstrReturn = cobjOrder.BillingCycleType.OtherName
        Case "OR.MAINTOVERBILLINGCYCLE"
            If cobjOrder.MaintOverBillingCycle = -1 Then
                'Same as base billing cycle
                mstrReturn = cobjOrder.BillingCycleType.Name
            Else
                mstrReturn = cobjOrder.OverBillingCycleType.Name
            End If
         Case "OR.MAINTOVERBILLINGCYCLEOTHER" ' JB added 03/24/2015
            If cobjOrder.MaintOverBillingCycle = -1 Then
                'Same as base billing cycle
                mstrReturn = cobjOrder.BillingCycleType.OtherName
            Else
                mstrReturn = cobjOrder.OverBillingCycleType.OtherName
            End If
        Case "OR.MAINTENANCEPYMT"
            mstrReturn = Format(cobjOrder.PaymentPerBillingCycle, "$###,###,##0.00")
        Case "OR.MAINTSTARTDATE"
            If cobjOrder.MaintStartDate <> "1/1/1900" Then
                mstrReturn = Format(cobjOrder.MaintStartDate, "mm/dd/yyyy")
            Else
                mstrReturn = ""
            End If
        Case "OR.MAINTENDDATE"
            If cobjOrder.MaintStartDate <> "1/1/1900" Then
                mstrReturn = Format(DateAdd("m", cobjOrder.MaintNumMonths, cobjOrder.MaintStartDate), "mm/dd/yyyy")
            Else
                mstrReturn = ""
            End If
        Case "OR.MAINTSTARTDATEV2"  'V2
            If cobjOrder.MaintStartDateV2 > 1 Then
                mstrReturn = Format(DateAdd("d", cobjOrder.MaintStartDateV2, "12/30/1899"), "mm/dd/yyyy")
            Else
                mstrReturn = ""
            End If
        Case "OR.MAINTENDDATEV2"    'V2
            If cobjOrder.MaintStartDateV2 > 1 Then
                mdatTemp = DateAdd("d", cobjOrder.MaintStartDateV2, "12/30/1899")
                mdatTemp = DateAdd("m", cobjOrder.MaintNumMonths, mdatTemp)
                mstrReturn = Format(mdatTemp, "mm/dd/yyyy")
            Else
                mstrReturn = ""
            End If
        Case "OR.MAINTNUMMONTHS"
            mstrReturn = cobjOrder.MaintNumMonths
        Case "OR.MAINTBASEPAYMENT"  'V2
            mstrReturn = Format(cobjOrder.MaintBasePayment, "$###,###,##0.00")
        Case "OR.MAINTFIXEDANNUAL"  'V2
            mcurTotal = 0
            For mint = 1 To cobjOrder.OrderLines.Count
                mcurTotal = mcurTotal + (cobjOrder.OrderLines(mint).BundleQuantity * cobjOrder.OrderLines(mint).FixedAmount)
            Next
            mstrReturn = Format(mcurTotal, "$###,###,##0.00")
        Case "OR.MAINTTOTALBWVOL"
            mlngTotal = 0
            For mint = 1 To cobjOrder.PrimaryOrderLines.Count
                mlngTotal = mlngTotal + cobjOrder.PrimaryOrderLines(mint).BaseCopiesBW
            Next
            mstrReturn = Format(CStr(mlngTotal), "###,###,##0")
        Case "OR.MAINTPYMTBW"
            mstrReturn = Format(cobjOrder.CopyChargePerBillingCycle("BW"), "$###,###,##0.00")
        Case "OR.MAINTPYMTCOLOR"
            mstrReturn = Format(cobjOrder.CopyChargePerBillingCycle("Color"), "$###,###,##0.00")
        Case "OR.MAINTPYMT2COLOR"
            mstrReturn = Format(cobjOrder.CopyChargePerBillingCycle("2Color"), "$###,###,##0.00")
        
        Case "OR.SALESREPCOMMISSIONAMT"
            mstrReturn = Format(cobjOrder.SalesRepCommissionAmt, "$###,###,##0.00")
        Case "OR.PRIMARYSRCOMMISSIONAMT"
            mstrReturn = Format(cobjOrder.PrimarySRCommissionAmt, "$###,###,##0.00")
        Case "OR.SPLITSRCOMMISSIONAMT"
            mstrReturn = Format(cobjOrder.SplitSRCommissionAmt, "$###,###,##0.00")
            
        Case "OR.SPECTSTE"
            Set mrsSums = cobjOrder.GetTotalsByType()
            If Not mrsSums.EOF And Not mrsSums.BOF Then mrsSums.MoveFirst
            For mint = 1 To mrsSums.RecordCount
                If mint = 1 Then
                    mstrReturn = Format(mrsSums("TOTAL") + cobjOrder.Subtotalamount, "$###,###,##0.00")
                End If
                mrsSums.MoveNext
            Next
            Set mrsSums = Nothing
            
        Case "OR.SUBTOTALSELL_EXP1"
            mcurTotal = cobjOrder.Subtotalamount
            Set mrsSums = cobjOrder.GetTotalsByType()
            If Not mrsSums.EOF And Not mrsSums.BOF Then mrsSums.MoveFirst
            For mint = 1 To mrsSums.RecordCount
                If mint = 1 Then
                    If mrsSums("CreditInd") <> 0 Then
                        mcurTotal = mcurTotal + mrsSums("TOTAL")
                    Else
                        mcurTotal = mcurTotal - mrsSums("TOTAL")
                    End If
                End If
                mrsSums.MoveNext
            Next
            Set mrsSums = Nothing
            mstrReturn = Format(mcurTotal, "$###,###,##0.00")
            
        'Gross Profit +/- total Expenses#1
        Case "OR.GP_EXP1"
            mcurTotal = cobjOrder.SellTotalHardware - cobjOrder.BuyTotal
            Set mrsSums = cobjOrder.GetTotalsByType()
            If Not mrsSums.EOF And Not mrsSums.BOF Then mrsSums.MoveFirst
            For mint = 1 To mrsSums.RecordCount
                If mint = 1 Then
                    If mrsSums("CreditInd") <> 0 Then
                        mcurTotal = mcurTotal + mrsSums("TOTAL")
                    Else
                        mcurTotal = mcurTotal - mrsSums("TOTAL")
                    End If
                End If
                mrsSums.MoveNext
            Next
            Set mrsSums = Nothing
            mstrReturn = Format(mcurTotal, "$###,###,##0.00")
            
        'Gross Profit + All expenses (+/-)
        Case "OR.GP_ALLEXP"
            mcurTotal = cobjOrder.SellTotalHardware - cobjOrder.BuyTotal
            Set mrsSums = cobjOrder.GetTotalsByType()
            If Not mrsSums.EOF And Not mrsSums.BOF Then mrsSums.MoveFirst
            For mint = 1 To mrsSums.RecordCount
                If mrsSums("CreditInd") <> 0 Then
                    mcurTotal = mcurTotal + mrsSums("TOTAL")
                Else
                    mcurTotal = mcurTotal - mrsSums("TOTAL")
                End If
                mrsSums.MoveNext
            Next
            Set mrsSums = Nothing
            mstrReturn = Format(mcurTotal, "$###,###,##0.00")
            
        Case "OR.REPCOSTANDEXPENSES" 'REP BASE COST +- Soft Costs
            mcurTotal = 0
            mint = 0
           
            For mint = 1 To cobjOrder.OrderLines.Count
              If cobjOrder.OrderLines.Item(mint).CatalogItemID < 0 Then 'manual items
                  mcurTotal = mcurTotal + (cobjOrder.OrderLines.Item(mint).BuyPrice * cobjOrder.OrderLines.Item(mint).Quantity)
              Else
                mcurTotal = mcurTotal + (cobjOrder.OrderLines.Item(mint).BuyPrice * cobjOrder.OrderLines.Item(mint).Quantity)
              End If
            Next
            
            For mintLineIx = 1 To cobjOrder.OrderAdjustments.Count
              If cobjOrder.OrderAdjustments.Item(mintLineIx).OrderAdjType.CreditDebitInd = "D" And cobjOrder.OrderAdjustments.Item(mintLineIx).OrderAdjType.IsSoftCost = True Then
                mcurTotal = mcurTotal + cobjOrder.OrderAdjustments.Item(mintLineIx).AdjAmount
              ElseIf cobjOrder.OrderAdjustments.Item(mintLineIx).OrderAdjType.IsSoftCost = True And cobjOrder.OrderAdjustments.Item(mintLineIx).OrderAdjType.CreditDebitInd = "C" Then
                mcurTotal = mcurTotal - cobjOrder.OrderAdjustments.Item(mintLineIx).AdjAmount
              Else
                mcurTotal = mcurTotal
              End If
            Next
            
        
            mstrReturn = Format(mcurTotal, "$###,###,##0.00")
            
         Case "OR.REPCOSTANDEXPENSESWITHCSMP" 'REP BASE COST +- Soft Costs
            mcurTotal = 0
            mint = 0
           
            For mint = 1 To cobjOrder.OrderLines.Count
              If cobjOrder.OrderLines.Item(mint).CatalogItemID < 0 Then 'manual items
                  mcurTotal = mcurTotal + (cobjOrder.OrderLines.Item(mint).BuyPrice * cobjOrder.OrderLines.Item(mint).Quantity)
              Else
                mcurTotal = mcurTotal + (cobjOrder.OrderLines.Item(mint).BuyPrice * cobjOrder.OrderLines.Item(mint).Quantity)
              End If
            Next
            
            For mintLineIx = 1 To cobjOrder.OrderAdjustments.Count
              If cobjOrder.OrderAdjustments.Item(mintLineIx).OrderAdjType.CreditDebitInd = "D" And cobjOrder.OrderAdjustments.Item(mintLineIx).OrderAdjType.IsSoftCost = True Then
                mcurTotal = mcurTotal + cobjOrder.OrderAdjustments.Item(mintLineIx).AdjAmount
              ElseIf cobjOrder.OrderAdjustments.Item(mintLineIx).OrderAdjType.IsSoftCost = True And cobjOrder.OrderAdjustments.Item(mintLineIx).OrderAdjType.CreditDebitInd = "C" Then
                mcurTotal = mcurTotal - cobjOrder.OrderAdjustments.Item(mintLineIx).AdjAmount
              Else
                mcurTotal = mcurTotal
              End If
            Next
            
            If (cobjOrder.SRCommCSMPCreditValue > 0) Then
              mcurTotal = mcurTotal - cobjOrder.SRCommCSMPCreditValue
            End If
            
           mstrReturn = Format(mcurTotal, "$###,###,##0.00")
         
        Case "OR.GSAAlert"
            mstrReturn = ""
            If Object.SysParms.ShowGSAInd = True Then
                For mint = 1 To cobjOrder.OrderLines.Count
                    If CCur(cobjOrder.OrderLines(mint).sellprice) < CCur(cobjOrder.OrderLines(mint).CatalogItem.GSAPrice) Then
                        mstrReturn = "GSA ALERT: Product sold below GSA."
                        Exit For
                    End If
                Next
            Else
                mstrReturn = ""
            End If
            
        Case "OR.SPECGP12"
            mcurTotal = cobjOrder.SellTotalHardware - cobjOrder.BuyTotal
            Set mrsSums = cobjOrder.GetTotalsByType()
            If Not mrsSums.EOF And Not mrsSums.BOF Then mrsSums.MoveFirst
            For mint = 1 To mrsSums.RecordCount
                If mint = 1 Or mint = 2 Then
                    mcurTotal = mcurTotal + mrsSums("TOTAL")
                End If
                mrsSums.MoveNext
            Next
            Set mrsSums = Nothing
            mstrReturn = Format(mcurTotal, "$###,###,##0.00")
            
        Case "OR.SPECGP23"
            mcurTotal = cobjOrder.SellTotalHardware - cobjOrder.BuyTotal
            Set mrsSums = cobjOrder.GetTotalsByType()
            If Not mrsSums.EOF And Not mrsSums.BOF Then mrsSums.MoveFirst
            For mint = 1 To mrsSums.RecordCount
                If mint = 2 Or mint = 3 Then
                    mcurTotal = mcurTotal + mrsSums("TOTAL")
                End If
                mrsSums.MoveNext
            Next
            Set mrsSums = Nothing
            mstrReturn = Format(mcurTotal, "$###,###,##0.00")
            
        Case "OR.NUMMACHINECFG"
            mintCountDown = cobjOrder.GetNumberOfPrimaryMachines
'            For mintLineIx = 1 To cobjOrder.PrimaryOrderLines.Count
'                For mintMachineIx = 1 To cobjOrder.PrimaryOrderLines.Item(mintLineIx).OrderMachines.Count
'                    mintCountDown = mintCountDown + 1
'                Next
'            Next
            mstrReturn = mintCountDown
            
        Case "OR.DELIVERYMETHOD"
            For mintLineIx = 1 To cobjOrder.DeliveryJobs.Count
                mstrTemp = Trim(cobjOrder.DeliveryJobs.Item(mintLineIx).Method)
                If mstrTemp <> "" Then
                    'See if we already have recorded this method
                    For mintRowIx = 0 To UBound(mstrMethod)
                        If mstrMethod(mintRowIx) = "" Then
                            Exit For
                        End If
                        If LCase(mstrMethod(mintRowIx)) = LCase(mstrTemp) Then
                            mbolFound = True
                            Exit For
                        End If
                    Next
                    'If no, then save it
                    If mbolFound = False Then
                        If mintRowIx <= UBound(mstrMethod) Then
                            mstrMethod(mintRowIx) = mstrTemp
                        End If
                    End If
                End If
            Next
            
            mstrReturn = ""
            For mintRowIx = 0 To UBound(mstrMethod)
                If mstrMethod(mintRowIx) <> "" Then
                    If mintRowIx > 0 Then
                        mstrReturn = mstrReturn & "; "
                    End If
                    mstrReturn = mstrReturn & mstrMethod(mintRowIx)
                End If
            Next
        
        'V2 - postage meter stuff
        Case "OR.POSTAGEMETERRENTTOTAL"     'Total for all line items
            mcurTotal = 0
            For mint = 1 To cobjOrder.OrderLines.Count
                mcurTotal = mcurTotal + (cobjOrder.OrderLines(mint).BundleQuantity * cobjOrder.OrderLines(mint).MonthlyMeter)
            Next
            mstrReturn = Format(mcurTotal, "$###,###,##0.00")
        Case "OR.POSTAGEANNUALTOTAL"        'Total for all line items
            mcurTotal = 0
            For mint = 1 To cobjOrder.OrderLines.Count
                mcurTotal = mcurTotal + (cobjOrder.OrderLines(mint).BundleQuantity * cobjOrder.OrderLines(mint).RateAndStructure)
            Next
            mstrReturn = Format(mcurTotal, "$###,###,##0.00")
            
        Case "OJ.COMPANYNAME"
            mstrReturn = cobjOrder.DeliveryJobs.Item(mintDJSeqNbr).CustomerName
        Case "OJ.CUSTOMERNBR"
            mstrReturn = cobjOrder.DeliveryJobs.Item(mintDJSeqNbr).Customer.OMDCustomerNbr
        
        Case "OJ.ADDRESSLABEL"
            mstrErrLocation = "1"
            If cobjOrder.DeliveryJobs.Count < mintDJSeqNbr Then
                mstrReturn = cobjOrder.BILLINGAddressLabel
            Else
                mstrErrLocation = "2"
                If cobjOrder.DeliveryJobs.Item(mintDJSeqNbr).CustomerSelect > 0 Then
                    mstrReturn = cobjOrder.DeliveryJobs.Item(mintDJSeqNbr).Customer.Company.Address.AddressLabel
                Else
                    mstrErrLocation = "3"
                    'TES 5/24/2011 - Repair reference to the manually entered address.
                    If cobjOrder.DeliveryJobs.Item(mintDJSeqNbr).CustomerSelect = -100 Then
                        mstrErrLocation = "4"
                        mlngLocTypeID = cobjOrder.DeliveryJobs.Item(mintDJSeqNbr).Entity.GetLocationTypeGivenLabel("Company.Address")
                        mstrReturn = cobjOrder.DeliveryJobs.Item(mintDJSeqNbr).Entity.GetLocationWithTypeID(mlngLocTypeID).AddressLabel
                    Else
                        mstrReturn = cobjOrder.BILLINGAddressLabel
                    End If
                End If
            End If
        Case "OJ.ADDRESS1"
            mstrErrLocation = "1"
            If cobjOrder.DeliveryJobs.Count < mintDJSeqNbr Then
                mstrReturn = cobjOrder.BILLINGAddress1
            Else
                mstrErrLocation = "2"
                If cobjOrder.DeliveryJobs.Item(mintDJSeqNbr).CustomerSelect > 0 Then
                    mstrReturn = cobjOrder.DeliveryJobs.Item(mintDJSeqNbr).Customer.Company.Address.Address1
                Else
                    mstrErrLocation = "3"
                    'TES 5/24/2011 - Repair reference to the manually entered address.
                    If cobjOrder.DeliveryJobs.Item(mintDJSeqNbr).CustomerSelect = -100 Then
                        mstrErrLocation = "4"
                        mlngLocTypeID = cobjOrder.DeliveryJobs.Item(mintDJSeqNbr).Entity.GetLocationTypeGivenLabel("Company.Address")
                        mstrReturn = cobjOrder.DeliveryJobs.Item(mintDJSeqNbr).Entity.GetLocationWithTypeID(mlngLocTypeID).Address1
                    Else
                        mstrReturn = cobjOrder.BILLINGAddress1
                    End If
                End If
            End If
            
        Case "OJ.ADDRESS2"
            mstrErrLocation = "1"
            If cobjOrder.DeliveryJobs.Count < mintDJSeqNbr Then
                mstrReturn = cobjOrder.BILLINGAddress2
            Else
                mstrErrLocation = "2"
                If cobjOrder.DeliveryJobs.Item(mintDJSeqNbr).CustomerSelect > 0 Then
                    mstrReturn = cobjOrder.DeliveryJobs.Item(mintDJSeqNbr).Customer.Company.Address.Address2
                Else
                    mstrErrLocation = "3"
                    'TES 5/24/2011 - Repair reference to the manually entered address.
                    If cobjOrder.DeliveryJobs.Item(mintDJSeqNbr).CustomerSelect = -100 Then
                        mstrErrLocation = "4"
                        mlngLocTypeID = cobjOrder.DeliveryJobs.Item(mintDJSeqNbr).Entity.GetLocationTypeGivenLabel("Company.Address")
                        mstrReturn = cobjOrder.DeliveryJobs.Item(mintDJSeqNbr).Entity.GetLocationWithTypeID(mlngLocTypeID).Address2
                    Else
                        mstrReturn = cobjOrder.BILLINGAddress2
                    End If
                End If
            End If
            
        Case "OJ.ADDRESSCITY"
            mstrErrLocation = "1"
            If cobjOrder.DeliveryJobs.Count < mintDJSeqNbr Then
                mstrReturn = cobjOrder.BILLINGcity
            Else
                mstrErrLocation = "2"
                If cobjOrder.DeliveryJobs.Item(mintDJSeqNbr).CustomerSelect > 0 Then
                    mstrReturn = cobjOrder.DeliveryJobs.Item(mintDJSeqNbr).Customer.Company.Address.City
                Else
                    mstrErrLocation = "3"
                    'TES 5/24/2011 - Repair reference to the manually entered address.
                    If cobjOrder.DeliveryJobs.Item(mintDJSeqNbr).CustomerSelect = -100 Then
                        mstrErrLocation = "4"
                        mlngLocTypeID = cobjOrder.DeliveryJobs.Item(mintDJSeqNbr).Entity.GetLocationTypeGivenLabel("Company.Address")
                        mstrReturn = cobjOrder.DeliveryJobs.Item(mintDJSeqNbr).Entity.GetLocationWithTypeID(mlngLocTypeID).City
                    Else
                        mstrReturn = cobjOrder.BILLINGcity
                    End If
                End If
            End If
            
        Case "OJ.ADDRESSSTATE"
            mstrErrLocation = "1"
            If cobjOrder.DeliveryJobs.Count < mintDJSeqNbr Then
                mstrReturn = cobjOrder.BILLINGState
            Else
                mstrErrLocation = "2"
                If cobjOrder.DeliveryJobs.Item(mintDJSeqNbr).CustomerSelect > 0 Then
                    mstrReturn = cobjOrder.DeliveryJobs.Item(mintDJSeqNbr).Customer.Company.Address.StateCode
                Else
                    mstrErrLocation = "3"
                    'TES 5/24/2011 - Repair reference to the manually entered address.
                    If cobjOrder.DeliveryJobs.Item(mintDJSeqNbr).CustomerSelect = -100 Then
                        mstrErrLocation = "4"
                        mlngLocTypeID = cobjOrder.DeliveryJobs.Item(mintDJSeqNbr).Entity.GetLocationTypeGivenLabel("Company.Address")
                        mstrReturn = cobjOrder.DeliveryJobs.Item(mintDJSeqNbr).Entity.GetLocationWithTypeID(mlngLocTypeID).StateCode
                    Else
                        mstrReturn = cobjOrder.BILLINGState
                    End If
                End If
            End If
            
        Case "OJ.ADDRESSPOSTALCODE"
            mstrErrLocation = "1"
            If cobjOrder.DeliveryJobs.Count < mintDJSeqNbr Then
                mstrReturn = cobjOrder.BILLINGPostalCode
            Else
                mstrErrLocation = "2"
                If cobjOrder.DeliveryJobs.Item(mintDJSeqNbr).CustomerSelect > 0 Then
                    mstrReturn = cobjOrder.DeliveryJobs.Item(mintDJSeqNbr).Customer.Company.Address.PostalCode
                Else
                    mstrErrLocation = "3"
                    'TES 5/24/2011 - Repair reference to the manually entered address.
                    If cobjOrder.DeliveryJobs.Item(mintDJSeqNbr).CustomerSelect = -100 Then
                        mstrErrLocation = "4"
                        mlngLocTypeID = cobjOrder.DeliveryJobs.Item(mintDJSeqNbr).Entity.GetLocationTypeGivenLabel("Company.Address")
                        mstrReturn = cobjOrder.DeliveryJobs.Item(mintDJSeqNbr).Entity.GetLocationWithTypeID(mlngLocTypeID).PostalCode
                    Else
                        mstrReturn = cobjOrder.BILLINGPostalCode
                    End If
                End If
            End If
            
            
        Case "OJ.PRIMARYPHONE"
            mstrReturn = Format(cobjOrder.DeliveryJobs.Item(mintDJSeqNbr).PriContactPhone, "(###) ###-####")
        Case "OJ.PRIMARYCONTACTNAME"
            mstrReturn = cobjOrder.DeliveryJobs.Item(mintDJSeqNbr).PriContactName
            If InStr(mstrReturn, ",") > 0 Then
                mstrReturn = Trim(Mid(mstrReturn, InStr(mstrReturn, ",") + 1)) & " " & Trim(Mid(mstrReturn, 1, InStr(mstrReturn, ",") - 1))
            End If
        Case "OJ.PRICONTACTFAX"
            mstrReturn = Format(cobjOrder.DeliveryJobs.Item(mintDJSeqNbr).PriContactFax, "(###) ###-####")
        Case "OJ.PRICONTACTCELL"
            mstrReturn = Format(cobjOrder.DeliveryJobs.Item(mintDJSeqNbr).PriContactCell, "(###) ###-####")
        Case "OJ.PRICONTACTEMAIL"
            mstrReturn = LCase(cobjOrder.DeliveryJobs.Item(mintDJSeqNbr).PriContactEmail)
            
        Case "OJ.ITCONTACTNAME"
            mstrReturn = cobjOrder.DeliveryJobs.Item(mintDJSeqNbr).ITContactName
            If InStr(mstrReturn, ",") > 0 Then
                mstrReturn = Trim(Mid(mstrReturn, InStr(mstrReturn, ",") + 1)) & " " & Trim(Mid(mstrReturn, 1, InStr(mstrReturn, ",") - 1))
            End If
        Case "OJ.ITPHONE"
            mstrReturn = Format(cobjOrder.DeliveryJobs.Item(mintDJSeqNbr).ITContactPhone, "(###) ###-####")
            
        Case "OJ.ITCONTACTFAX"
            mstrReturn = Format(cobjOrder.DeliveryJobs.Item(mintDJSeqNbr).ITContactFax, "(###) ###-####")
        Case "OJ.ITCONTACTCELL"
            mstrReturn = Format(cobjOrder.DeliveryJobs.Item(mintDJSeqNbr).ITContactCell, "(###) ###-####")
        Case "OJ.ITCONTACTEMAIL"
            mstrReturn = LCase(cobjOrder.DeliveryJobs.Item(mintDJSeqNbr).ITContactEmail)
        Case Else
            mstrReturn = GetOrderSpecificData2(pField, pIndex, pAddendumName, pCountDown, pDJSeqNbr, pcol, pcurOB_BuyAmount, pcurOB_SellAmount, pcurOB_LeasePymt, pLineIx, pMachineIx, pOffsetIx, pInt)
    End Select
        
    GoTo EndRoutine
    
EndRoutine:
    GetOrderSpecificData = mstrReturn

End Function
