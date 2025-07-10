Private Function GetOrderSpecificData3(pField As String, pIndex As Integer, _
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
        Case "ML.ADDRESSLABEL"
            'TES 8/3/2011 Repaired indexing error and added the safty if index vs count check
            If cobjOrder.DeliveryJobEQs.Count < mintLineIx Then
                mstrReturn = cobjOrder.BILLINGAddressLabel
            Else
                If cobjOrder.DeliveryJobEQs.Item(mintLineIx).DeliveryJob.CustomerSelect > 0 Then
                    mstrReturn = cobjOrder.DeliveryJobEQs.Item(mintLineIx).DeliveryJob.Customer.Company.Address.AddressLabel
                Else
                    'TES 5/24/2011 - Repair reference to the manually entered address.
                    If cobjOrder.DeliveryJobEQs.Item(mintLineIx).DeliveryJob.CustomerSelect = -100 Then
                        mlngLocTypeID = cobjOrder.DeliveryJobEQs.Item(mintLineIx).DeliveryJob.Entity.GetLocationTypeGivenLabel("Company.Address")
                        mstrReturn = cobjOrder.DeliveryJobEQs.Item(mintLineIx).DeliveryJob.Entity.GetLocationWithTypeID(mlngLocTypeID).AddressLabel
                    Else
                        mstrReturn = cobjOrder.BILLINGAddressLabel
                    End If
    
                End If
            End If
        Case "ML.ADDRESSONELINE"
            'TES 8/3/2011 Repaired indexing error and added the safty if index vs count check
            If cobjOrder.DeliveryJobEQs.Count < mintLineIx Then
                mstrReturn = cobjOrder.BILLINGAddressLabel
            Else
                If cobjOrder.DeliveryJobEQs.Item(mintLineIx).DeliveryJob.CustomerSelect > 0 Then
                    mstrReturn = cobjOrder.DeliveryJobEQs.Item(mintLineIx).DeliveryJob.Customer.Company.Address.AddressOneLine
                Else
                    'TES 5/24/2011 - Repair reference to the manually entered address.
                    If cobjOrder.DeliveryJobEQs.Item(mintLineIx).DeliveryJob.CustomerSelect = -100 Then
                        mlngLocTypeID = cobjOrder.DeliveryJobEQs.Item(mintLineIx).DeliveryJob.Entity.GetLocationTypeGivenLabel("Company.Address")
                        mstrReturn = cobjOrder.DeliveryJobEQs.Item(mintLineIx).DeliveryJob.Entity.GetLocationWithTypeID(mlngLocTypeID).AddressOneLine
                    Else
                        mstrReturn = cobjOrder.BillingAddressOneLine
                    End If
    
                End If
            End If
        Case "ML.MODEL"
            mstrReturn = cobjOrder.DeliveryJobEQs.Item(mintLineIx).Model
        Case "ML.PRODUCTCODE"
            mstrReturn = cobjOrder.DeliveryJobEQs.Item(mintLineIx).CatalogItem.OMDCode
        Case "ML.POWERSUPPLY"
            mstrReturn = cobjOrder.DeliveryJobEQs.Item(mintLineIx).CatalogItem.PowerSupply.Name
        Case "ML.MFG"
            If cobjOrder.DeliveryJobEQs.Item(mintLineIx).CatalogItemID = -100 Then
                If cobjOrder.DeliveryJobEQs.Item(mintLineIx).orderlineid > 0 Then
                    mstrReturn = cobjOrder.DeliveryJobEQs.Item(mintLineIx).OrderLine.ProductMfg.Name
                Else
                    mstrReturn = ""
                End If
            Else
                mstrReturn = cobjOrder.DeliveryJobEQs.Item(mintLineIx).CatalogItem.ProductMfg.Name
            End If
        Case "ML.DESCRIPTION"
            If cobjOrder.DeliveryJobEQs.Item(mintLineIx).CatalogItemID > 0 Then
                mstrReturn = cobjOrder.DeliveryJobEQs.Item(mintLineIx).CatalogItem.Description
            Else
                For mIndex = 1 To cobjOrder.OrderLines.Count
                    If cobjOrder.DeliveryJobEQs.Item(mintLineIx).BundleID = cobjOrder.OrderLines.Item(mIndex).BundleID And _
                       cobjOrder.DeliveryJobEQs.Item(mintLineIx).CatalogItemID = cobjOrder.OrderLines.Item(mIndex).CatalogItemID Then
                        mstrReturn = cobjOrder.OrderLines.Item(mIndex).Description
                        Exit For
                    End If
                Next
            End If
        Case "ML.QTY"
            For mIndex = 1 To cobjOrder.OrderLines.Count
                If cobjOrder.DeliveryJobEQs.Item(mintLineIx).BundleID = cobjOrder.OrderLines.Item(mIndex).BundleID And _
                   cobjOrder.DeliveryJobEQs.Item(mintLineIx).CatalogItemID = cobjOrder.OrderLines.Item(mIndex).CatalogItemID Then
                    mstrReturn = cobjOrder.OrderLines.Item(mIndex).PerBundleQuantity
                    Exit For
                End If
            Next
        
        Case "ML.UNITPRICE"
            For mIndex = 1 To cobjOrder.OrderLines.Count
                If cobjOrder.DeliveryJobEQs.Item(mintLineIx).BundleID = cobjOrder.OrderLines.Item(mIndex).BundleID And _
                   cobjOrder.DeliveryJobEQs.Item(mintLineIx).CatalogItemID = cobjOrder.OrderLines.Item(mIndex).CatalogItemID Then
                    mstrReturn = Format(cobjOrder.OrderLines.Item(mIndex).sellprice, "$###,###,##0.00")
                    Exit For
                End If
            Next
            
        Case "ML.SERIAL"
            mstrReturn = cobjOrder.DeliveryJobEQs.Item(mintLineIx).Serial
        Case "ML.ASSETTAG"
            mstrReturn = cobjOrder.DeliveryJobEQs.Item(mintLineIx).AssetTag
        Case "ML.ITEMNBR"
            If cobjOrder.DeliveryJobEQs.Item(mintLineIx).CatalogItemID = -100 Then
                Set mTmpObj = cobjOrder.DeliveryJobEQs.Item(mintLineIx).GetAssocOrderLine
                mstrReturn = RemovePreceedingUnderscore(mTmpObj.MFGItemNbr)
                Set mTmpObj = Nothing
            Else
                mstrReturn = RemovePreceedingUnderscore(cobjOrder.DeliveryJobEQs.Item(mintLineIx).CatalogItem.MFGItemNbr)
            End If
        Case "ML.TOTALMETER"
            If cobjOrder.DeliveryJobEQs.Item(mintLineIx).TotalMeter >= 0 Then
                mstrReturn = cobjOrder.DeliveryJobEQs.Item(mintLineIx).TotalMeter
            Else
                mstrReturn = ""
            End If
        Case "ML.BWMETER"
            If cobjOrder.DeliveryJobEQs.Item(mintLineIx).BWMeter >= 0 Then
                mstrReturn = cobjOrder.DeliveryJobEQs.Item(mintLineIx).BWMeter
            Else
                mstrReturn = ""
            End If
        Case "ML.COLORMETER"
            If cobjOrder.DeliveryJobEQs.Item(mintLineIx).ColorMeter >= 0 Then
                mstrReturn = cobjOrder.DeliveryJobEQs.Item(mintLineIx).ColorMeter
            Else
                mstrReturn = ""
            End If
            
        'Per machine delivery job info
        Case "MJ.UNIQUEID"
            mstrReturn = cobjOrder.DeliveryJobEQs.Item(mint).BundleID & "-" & cobjOrder.DeliveryJobEQs.Item(mint).MachineID
        Case "MJ.COMPANY"
            mstrReturn = cobjOrder.DeliveryJobs.Item(mintLineIx).CustomerName
        Case "MJ.CONTACT"
            mstrReturn = cobjOrder.DeliveryJobs.Item(mintLineIx).PriContactName
        Case "MJ.ADDR"
            mstrReturn = cobjOrder.DeliveryJobs.Item(mintLineIx).Address.AddressLabel
        Case "MJ.CONTACT"
            mstrReturn = cobjOrder.DeliveryJobs.Item(mintLineIx).PriContactName
        Case "MJ.PHONE"
            mstrReturn = cobjOrder.DeliveryJobs.Item(mintLineIx).PriContactPhone
        Case "MJ.CELL"
            mstrReturn = cobjOrder.DeliveryJobs.Item(mintLineIx).PriContactCell
        Case "MJ.FAX"
            mstrReturn = cobjOrder.DeliveryJobs.Item(mintLineIx).PriContactFax
        Case "MJ.EMAIL"
            mstrReturn = LCase(cobjOrder.DeliveryJobs.Item(mintLineIx).PriContactEmail)
        Case "MJ.ITCONTACT"
            mstrReturn = cobjOrder.DeliveryJobs.Item(mintLineIx).ITContactName
        Case "MJ.ITPHONE"
            mstrReturn = cobjOrder.DeliveryJobs.Item(mintLineIx).ITContactPhone
        Case "MJ.ITCELL"
            mstrReturn = cobjOrder.DeliveryJobs.Item(mintLineIx).ITContactCell
        Case "MJ.ITFAX"
            mstrReturn = cobjOrder.DeliveryJobs.Item(mintLineIx).ITContactFax
        Case "MJ.ITEMAIL"
            mstrReturn = LCase(cobjOrder.DeliveryJobs.Item(mintLineIx).ITContactEmail)
        Case "MJ.DELDATE"
            If cobjOrder.DeliveryJobs.Item(mintLineIx).ScheduledDate > 1 Then
                mstrReturn = Format(DateAdd("d", cobjOrder.DeliveryJobs.Item(mintLineIx).ScheduledDate, "12/30/1899"), "mm/dd/yyyy")
            Else
                mstrReturn = ""
            End If
        Case "MJ.REQDELDATE"
            If cobjOrder.DeliveryJobs.Item(mintLineIx).RequestedDate > 1 Then
                mstrReturn = Format(DateAdd("d", cobjOrder.DeliveryJobs.Item(mintLineIx).RequestedDate, "12/30/1899"), "mm/dd/yyyy")
            Else
                mstrReturn = ""
            End If
        Case "MJ.STAIRS"
            If cobjOrder.DeliveryJobs.Item(mintLineIx).StairsInd = True Then
                mstrReturn = "Yes"
            Else
                mstrReturn = "No"
            End If
        Case "MJ.NUMSTAIRS"
            If cobjOrder.DeliveryJobs.Item(mintLineIx).StairsInd = True Then
                mstrReturn = cobjOrder.DeliveryJobs.Item(mintLineIx).NumStairs
            Else
                mstrReturn = ""
            End If
        Case "MJ.ELEVATOR"
            If cobjOrder.DeliveryJobs.Item(mintLineIx).ElevatorInd = True Then
                mstrReturn = "Yes"
            Else
                mstrReturn = "No"
            End If
        
        Case "MO.SELVAL"
            mstrReturn = cobjOrder.DeliveryJobs.Item(mintLineIx).DelJobOptions.Item(mintOffsetIx).DelJobOptionValue.Name
            
        Case "MW.NOTES"
            If cobjOrder.DeliveryJobs.Item(mintLineIx).WorkItems.Item(mintOffsetIx).IsOn = True Then
                mstrReturn = cobjOrder.DeliveryJobs.Item(mintLineIx).WorkItems.Item(mintOffsetIx).Notes
            Else
                mstrReturn = ""
            End If
        Case "MW.INSTRUCTIONS"
            If cobjOrder.DeliveryJobs.Item(mintLineIx).WorkItems.Item(mintOffsetIx).IsOn = True Then
                mstrReturn = cobjOrder.DeliveryJobs.Item(mintLineIx).WorkItems.Item(mintOffsetIx).Instructions
            Else
                mstrReturn = ""
            End If
            
        Case "MP.VALUE"
            If cobjOrder.DeliveryJobs.Item(pIndex).WorkItemBySystemLabel("Pickup").IsOn = False Then
                mstrReturn = ""
            Else
                Select Case UCase(cobjOrder.DeliveryJobs.Item(pIndex).WorkItemBySystemLabel("Pickup").PickupMoveEqs.Item(pLineIx).PickupMoveType)
                    Case "P"
                        mstrReturn = "Pickup"
                    Case "M"
                        mstrReturn = "Move"
                    Case "W"
                        mstrReturn = "Wrap"
                    Case "L"
                        mstrReturn = "Leave"
                    Case ""
                        mstrReturn = ""
                    Case Else
                        mstrReturn = "UnkVal(" & UCase(cobjOrder.DeliveryJobs.Item(pIndex).WorkItemBySystemLabel("Pickup").PickupMoveEqs.Item(pLineIx).PickupMoveType) & ")"
                End Select
            End If
        Case "MP.MODEL"
            If cobjOrder.DeliveryJobs.Item(pIndex).WorkItemBySystemLabel("Pickup").IsOn = False Then
                mstrReturn = ""
            Else
                mstrReturn = cobjOrder.DeliveryJobs.Item(pIndex).WorkItemBySystemLabel("Pickup").PickupMoveEqs.Item(pLineIx).Model
            End If
        Case "MP.SERIAL"
            If cobjOrder.DeliveryJobs.Item(pIndex).WorkItemBySystemLabel("Pickup").IsOn = False Then
                mstrReturn = ""
            Else
                mstrReturn = cobjOrder.DeliveryJobs.Item(pIndex).WorkItemBySystemLabel("Pickup").PickupMoveEqs.Item(pLineIx).SerialNbr
            End If
        Case "MP.EQID"
            If cobjOrder.DeliveryJobs.Item(pIndex).WorkItemBySystemLabel("Pickup").IsOn = False Then
                mstrReturn = ""
            Else
                mstrReturn = cobjOrder.DeliveryJobs.Item(pIndex).WorkItemBySystemLabel("Pickup").PickupMoveEqs.Item(pLineIx).AssetTag
            End If
        Case "MP.HDCLEANINDX"
            If cobjOrder.DeliveryJobs.Item(pIndex).WorkItemBySystemLabel("Pickup").IsOn = False Then
                mstrReturn = ""
            Else
                If cobjOrder.DeliveryJobs.Item(pIndex).WorkItemBySystemLabel("Pickup").PickupMoveEqs.Item(pLineIx).HDCleaningReqInd = True Then
                    mstrReturn = "X"
                Else
                    mstrReturn = ""
                End If
            End If
        Case "MP.HDCLEANBILLCUSTINDX"
            If cobjOrder.DeliveryJobs.Item(pIndex).WorkItemBySystemLabel("Pickup").IsOn = False Then
                mstrReturn = ""
            Else
                If cobjOrder.DeliveryJobs.Item(pIndex).WorkItemBySystemLabel("Pickup").PickupMoveEqs.Item(pLineIx).HDCleaningBillCustInd = True Then
                    mstrReturn = "X"
                Else
                    mstrReturn = ""
                End If
            End If
        Case "MP.LEASERETURNX"
            If cobjOrder.DeliveryJobs.Item(pIndex).WorkItemBySystemLabel("Pickup").IsOn = False Then
                mstrReturn = ""
            Else
                If cobjOrder.DeliveryJobs.Item(pIndex).WorkItemBySystemLabel("Pickup").PickupMoveEqs.Item(pLineIx).LeaseReturnInd = True Then
                    mstrReturn = "X"
                Else
                    mstrReturn = ""
                End If
            End If
        Case "MP.DEALEROWNINDX"
            If cobjOrder.DeliveryJobs.Item(pIndex).WorkItemBySystemLabel("Pickup").IsOn = False Then
                mstrReturn = ""
            Else
                If cobjOrder.DeliveryJobs.Item(pIndex).WorkItemBySystemLabel("Pickup").PickupMoveEqs.Item(pLineIx).DealerOwnInd = True Then
                    mstrReturn = "X"
                Else
                    mstrReturn = ""
                End If
            End If
        Case "MP.FCONAME"
            If cobjOrder.DeliveryJobs.Item(pIndex).WorkItemBySystemLabel("Pickup").IsOn = False Then
                mstrReturn = ""
            Else
                mstrReturn = cobjOrder.DeliveryJobs.Item(pIndex).WorkItemBySystemLabel("Pickup").PickupMoveEqs.Item(pLineIx).LeaseCoName
            End If
        Case "MP.FCOLEASENBR"
            If cobjOrder.DeliveryJobs.Item(pIndex).WorkItemBySystemLabel("Pickup").IsOn = False Then
                mstrReturn = ""
            Else
                mstrReturn = cobjOrder.DeliveryJobs.Item(pIndex).WorkItemBySystemLabel("Pickup").PickupMoveEqs.Item(pLineIx).LeaseNbr
            End If
        Case "MP.FCOLEASEEXPDT"
            If cobjOrder.DeliveryJobs.Item(pIndex).WorkItemBySystemLabel("Pickup").IsOn = False Then
                mstrReturn = ""
            Else
                If cobjOrder.DeliveryJobs.Item(pIndex).WorkItemBySystemLabel("Pickup").PickupMoveEqs.Item(pLineIx).LeaseExpDate <> "1/1/1900" Then
                    mstrReturn = Format(cobjOrder.DeliveryJobs.Item(pIndex).WorkItemBySystemLabel("Pickup").PickupMoveEqs.Item(pLineIx).LeaseExpDate, "mm/dd/yyyy")
                Else
                    mstrReturn = ""
                End If
            End If
        
        Case "M1.VALUE"
            If cobjOrder.DeliveryJobs.Item(pIndex).WorkItemBySystemLabel("Pickup").IsOn = False Then
                mstrReturn = ""
            Else
                Select Case UCase(cobjOrder.DeliveryJobs.Item(pIndex).WorkItemBySystemLabel("Pickup").PickupMoveEqsTypes("PM").Item(pLineIx).PickupMoveType)
                    Case "P"
                        mstrReturn = "Pickup"
                    Case "M"
                        mstrReturn = "Move"
                    Case "W"
                        mstrReturn = "Wrap"
                    Case "L"
                        mstrReturn = "Leave"
                    Case ""
                        mstrReturn = ""
                    Case Else
                        mstrReturn = "UnkVal(" & UCase(cobjOrder.DeliveryJobs.Item(pIndex).WorkItemBySystemLabel("Pickup").PickupMoveEqsTypes("PM").Item(pLineIx).PickupMoveType) & ")"
                End Select
            End If
        Case "M1.MODEL"
            If cobjOrder.DeliveryJobs.Item(pIndex).WorkItemBySystemLabel("Pickup").IsOn = False Then
                mstrReturn = ""
            Else
                mstrReturn = cobjOrder.DeliveryJobs.Item(pIndex).WorkItemBySystemLabel("Pickup").PickupMoveEqsTypes("PM").Item(pLineIx).Model
            End If
        Case "M1.SERIAL"
            If cobjOrder.DeliveryJobs.Item(pIndex).WorkItemBySystemLabel("Pickup").IsOn = False Then
                mstrReturn = ""
            Else
                mstrReturn = cobjOrder.DeliveryJobs.Item(pIndex).WorkItemBySystemLabel("Pickup").PickupMoveEqsTypes("PM").Item(pLineIx).SerialNbr
            End If
        Case "M1.EQID"
            If cobjOrder.DeliveryJobs.Item(pIndex).WorkItemBySystemLabel("Pickup").IsOn = False Then
                mstrReturn = ""
            Else
                mstrReturn = cobjOrder.DeliveryJobs.Item(pIndex).WorkItemBySystemLabel("Pickup").PickupMoveEqsTypes("PM").Item(pLineIx).AssetTag
            End If
        Case "M1.HDCLEANINDX"
            If cobjOrder.DeliveryJobs.Item(pIndex).WorkItemBySystemLabel("Pickup").IsOn = False Then
                mstrReturn = ""
            Else
                If cobjOrder.DeliveryJobs.Item(pIndex).WorkItemBySystemLabel("Pickup").PickupMoveEqsTypes("PM").Item(pLineIx).HDCleaningReqInd = True Then
                    mstrReturn = "X"
                Else
                    mstrReturn = ""
                End If
            End If
        Case "M1.HDCLEANBILLCUSTINDX"
            If cobjOrder.DeliveryJobs.Item(pIndex).WorkItemBySystemLabel("Pickup").IsOn = False Then
                mstrReturn = ""
            Else
                If cobjOrder.DeliveryJobs.Item(pIndex).WorkItemBySystemLabel("Pickup").PickupMoveEqsTypes("PM").Item(pLineIx).HDCleaningBillCustInd = True Then
                    mstrReturn = "X"
                Else
                    mstrReturn = ""
                End If
            End If
        Case "M1.LEASERETURNX"
            If cobjOrder.DeliveryJobs.Item(pIndex).WorkItemBySystemLabel("Pickup").IsOn = False Then
                mstrReturn = ""
            Else
                If cobjOrder.DeliveryJobs.Item(pIndex).WorkItemBySystemLabel("Pickup").PickupMoveEqsTypes("PM").Item(pLineIx).LeaseReturnInd = True Then
                    mstrReturn = "X"
                Else
                    mstrReturn = ""
                End If
            End If
        Case "M1.DEALEROWNINDX"
            If cobjOrder.DeliveryJobs.Item(pIndex).WorkItemBySystemLabel("Pickup").IsOn = False Then
                mstrReturn = ""
            Else
                If cobjOrder.DeliveryJobs.Item(pIndex).WorkItemBySystemLabel("Pickup").PickupMoveEqsTypes("PM").Item(pLineIx).DealerOwnInd = True Then
                    mstrReturn = "X"
                Else
                    mstrReturn = ""
                End If
            End If
        Case "M1.FCONAME"
            If cobjOrder.DeliveryJobs.Item(pIndex).WorkItemBySystemLabel("Pickup").IsOn = False Then
                mstrReturn = ""
            Else
                mstrReturn = cobjOrder.DeliveryJobs.Item(pIndex).WorkItemBySystemLabel("Pickup").PickupMoveEqsTypes("PM").Item(pLineIx).LeaseCoName
            End If
        Case "M1.FCOLEASENBR"
            If cobjOrder.DeliveryJobs.Item(pIndex).WorkItemBySystemLabel("Pickup").IsOn = False Then
                mstrReturn = ""
            Else
                mstrReturn = cobjOrder.DeliveryJobs.Item(pIndex).WorkItemBySystemLabel("Pickup").PickupMoveEqsTypes("PM").Item(pLineIx).LeaseNbr
            End If
        Case "M1.FCOLEASEEXPDT"
            If cobjOrder.DeliveryJobs.Item(pIndex).WorkItemBySystemLabel("Pickup").IsOn = False Then
                mstrReturn = ""
            Else
                If cobjOrder.DeliveryJobs.Item(pIndex).WorkItemBySystemLabel("Pickup").PickupMoveEqsTypes("PM").Item(pLineIx).LeaseExpDate <> "1/1/1900" Then
                    mstrReturn = Format(cobjOrder.DeliveryJobs.Item(pIndex).WorkItemBySystemLabel("Pickup").PickupMoveEqsTypes("PM").Item(pLineIx).LeaseExpDate, "mm/dd/yyyy")
                Else
                    mstrReturn = ""
                End If
            End If
        
        Case "M2.VALUE"
            If cobjOrder.DeliveryJobs.Item(pIndex).WorkItemBySystemLabel("Pickup").IsOn = False Then
                mstrReturn = ""
            Else
                Select Case UCase(cobjOrder.DeliveryJobs.Item(pIndex).WorkItemBySystemLabel("Pickup").PickupMoveEqsTypes("WL").Item(pLineIx).PickupMoveType)
                    Case "P"
                        mstrReturn = "Pickup"
                    Case "M"
                        mstrReturn = "Move"
                    Case "W"
                        mstrReturn = "Wrap"
                    Case "L"
                        mstrReturn = "Leave"
                    Case ""
                        mstrReturn = ""
                    Case Else
                        mstrReturn = "UnkVal(" & UCase(cobjOrder.DeliveryJobs.Item(pIndex).WorkItemBySystemLabel("Pickup").PickupMoveEqsTypes("WL").Item(pLineIx).PickupMoveType) & ")"
                End Select
            End If
        Case "M2.MODEL"
            If cobjOrder.DeliveryJobs.Item(pIndex).WorkItemBySystemLabel("Pickup").IsOn = False Then
                mstrReturn = ""
            Else
                mstrReturn = cobjOrder.DeliveryJobs.Item(pIndex).WorkItemBySystemLabel("Pickup").PickupMoveEqsTypes("WL").Item(pLineIx).Model
            End If
        Case "M2.SERIAL"
            If cobjOrder.DeliveryJobs.Item(pIndex).WorkItemBySystemLabel("Pickup").IsOn = False Then
                mstrReturn = ""
            Else
                mstrReturn = cobjOrder.DeliveryJobs.Item(pIndex).WorkItemBySystemLabel("Pickup").PickupMoveEqsTypes("WL").Item(pLineIx).SerialNbr
            End If
        Case "M2.EQID"
            If cobjOrder.DeliveryJobs.Item(pIndex).WorkItemBySystemLabel("Pickup").IsOn = False Then
                mstrReturn = ""
            Else
                mstrReturn = cobjOrder.DeliveryJobs.Item(pIndex).WorkItemBySystemLabel("Pickup").PickupMoveEqsTypes("WL").Item(pLineIx).AssetTag
            End If
        Case "M2.HDCLEANINDX"
            If cobjOrder.DeliveryJobs.Item(pIndex).WorkItemBySystemLabel("Pickup").IsOn = False Then
                mstrReturn = ""
            Else
                If cobjOrder.DeliveryJobs.Item(pIndex).WorkItemBySystemLabel("Pickup").PickupMoveEqsTypes("WL").Item(pLineIx).HDCleaningReqInd = True Then
                    mstrReturn = "X"
                Else
                    mstrReturn = ""
                End If
            End If
        Case "M2.HDCLEANBILLCUSTINDX"
            If cobjOrder.DeliveryJobs.Item(pIndex).WorkItemBySystemLabel("Pickup").IsOn = False Then
                mstrReturn = ""
            Else
                If cobjOrder.DeliveryJobs.Item(pIndex).WorkItemBySystemLabel("Pickup").PickupMoveEqsTypes("WL").Item(pLineIx).HDCleaningBillCustInd = True Then
                    mstrReturn = "X"
                Else
                    mstrReturn = ""
                End If
            End If
        Case "M2.LEASERETURNX"
            If cobjOrder.DeliveryJobs.Item(pIndex).WorkItemBySystemLabel("Pickup").IsOn = False Then
                mstrReturn = ""
            Else
                If cobjOrder.DeliveryJobs.Item(pIndex).WorkItemBySystemLabel("Pickup").PickupMoveEqsTypes("WL").Item(pLineIx).LeaseReturnInd = True Then
                    mstrReturn = "X"
                Else
                    mstrReturn = ""
                End If
            End If
        Case "M2.DEALEROWNINDX"
            If cobjOrder.DeliveryJobs.Item(pIndex).WorkItemBySystemLabel("Pickup").IsOn = False Then
                mstrReturn = ""
            Else
                If cobjOrder.DeliveryJobs.Item(pIndex).WorkItemBySystemLabel("Pickup").PickupMoveEqsTypes("WL").Item(pLineIx).DealerOwnInd = True Then
                    mstrReturn = "X"
                Else
                    mstrReturn = ""
                End If
            End If
        Case "M2.FCONAME"
            If cobjOrder.DeliveryJobs.Item(pIndex).WorkItemBySystemLabel("Pickup").IsOn = False Then
                mstrReturn = ""
            Else
                mstrReturn = cobjOrder.DeliveryJobs.Item(pIndex).WorkItemBySystemLabel("Pickup").PickupMoveEqsTypes("WL").Item(pLineIx).LeaseCoName
            End If
        Case "M2.FCOLEASENBR"
            If cobjOrder.DeliveryJobs.Item(pIndex).WorkItemBySystemLabel("Pickup").IsOn = False Then
                mstrReturn = ""
            Else
                mstrReturn = cobjOrder.DeliveryJobs.Item(pIndex).WorkItemBySystemLabel("Pickup").PickupMoveEqsTypes("WL").Item(pLineIx).LeaseNbr
            End If
        Case "M2.FCOLEASEEXPDT"
            If cobjOrder.DeliveryJobs.Item(pIndex).WorkItemBySystemLabel("Pickup").IsOn = False Then
                mstrReturn = ""
            Else
                If cobjOrder.DeliveryJobs.Item(pIndex).WorkItemBySystemLabel("Pickup").PickupMoveEqsTypes("WL").Item(pLineIx).LeaseExpDate <> "1/1/1900" Then
                    mstrReturn = Format(cobjOrder.DeliveryJobs.Item(pIndex).WorkItemBySystemLabel("Pickup").PickupMoveEqsTypes("WL").Item(pLineIx).LeaseExpDate, "mm/dd/yyyy")
                Else
                    mstrReturn = ""
                End If
            End If
        
        Case "OZ.ASSETTAG"
            mstrReturn = cobjOrder.DeliveryJobEQs.Item(pIndex).AssetTag
        Case "OZ.BUNDLEID"
            mstrReturn = cobjOrder.DeliveryJobEQs.Item(pIndex).BundleID
        Case "OZ.DJNBR"
            mstrReturn = cobjOrder.DeliveryJobEQs.Item(pIndex).DeliveryJobCount
        Case "OZ.ISADDENDUMNEEDED"
            If cobjOrder.DeliveryJobEQs.Count > mIndex Then
                If Len(mstrAddendumname) > 0 Then
                    mstrReturn = "ADDENDUM (" & mstrAddendumname & ") IS REQUIRED"
                Else
                    mstrReturn = "ADDENDUM IS REQUIRED"
                End If
            Else
                mstrReturn = ""
            End If
        Case "OZ.MACHINEID"
            mstrReturn = cobjOrder.DeliveryJobEQs.Item(pIndex).MachineID
        Case "OZ.MFGNAME"
            mstrReturn = cobjOrder.DeliveryJobEQs.Item(pIndex).CatalogItem.ProductMfg.Name
        Case "OZ.MFGITEMNBR"
            mstrReturn = RemovePreceedingUnderscore(cobjOrder.DeliveryJobEQs.Item(pIndex).CatalogItem.MFGItemNbr)
        Case "OZ.MODEL"
            mstrReturn = cobjOrder.DeliveryJobEQs.Item(pIndex).Model
        Case "OZ.UNITPRICE"
            mstrReturn = Format(cobjOrder.DeliveryJobEQs.Item(pIndex).OrderLine.sellprice, "$###,###,##0.00")
        Case "OZ.SERIAL"
            mstrReturn = cobjOrder.DeliveryJobEQs.Item(pIndex).Serial
        Case Else
            mstrReturn = "Unk Field[" & pField & "]"
    End Select

EndRoutine:
    GetOrderSpecificData3 = mstrReturn

End Function