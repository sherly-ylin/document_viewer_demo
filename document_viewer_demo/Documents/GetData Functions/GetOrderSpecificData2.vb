Private Function GetOrderSpecificData2(pField As String, pIndex As Integer, _
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
    Dim mInt2 As Integer
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
        
        Case "OD.DELJOBOPTIONMAX"
            mlngTotal = -1
            For mintLineIx = 1 To cobjOrder.DeliveryJobs.Count
                Set mobjDelJobOption = cobjOrder.DeliveryJobs.Item(mintLineIx).DelJobOptionByTypeID(CLng(mIndex))
                If mobjDelJobOption.DelJobOptionValue.SeqNbr > mlngTotal Then
                    mlngTotal = mobjDelJobOption.DelJobOptionValue.SeqNbr
                    mstrReturn = mobjDelJobOption.DelJobOptionValue.Name
                End If
                Set mobjDelJobOption = Nothing
            Next
            
        Case "OS.COPIERBASEAMT"
            mstrReturn = Format(cobjOrder.CopierMoBaseAmt, "$###,###,##0.00")
        Case "OS.COPIERBASEVOLUME"
            mstrReturn = Format(cobjOrder.CopierMoBaseVolume, "###,###,##0")
        Case "OS.COPIERTOTALAMT"
            mstrReturn = Format(cobjOrder.CopierTotalBaseAmt, "$###,###,##0.00")
        Case "OS.COPIERTOTALVOLUME"
            mstrReturn = Format(cobjOrder.CopierTotalVolume, "###,###,##0")

        Case "OS.PRINTBASEAMT"
            mstrReturn = Format(cobjOrder.PrintMoBaseAmt, "$###,###,##0.00")
        Case "OS.PRINTBASEVOLUME"
            mstrReturn = Format(cobjOrder.PrintMoBaseVolume, "###,###,##0")
        Case "OS.PRINTTOTALAMT"
            mstrReturn = Format(cobjOrder.PrintTotalBaseAmt, "$###,###,##0.00")
        Case "OS.PRINTTOTALVOLUME"
            mstrReturn = Format(cobjOrder.PrintTotalVolume, "###,###,##0")

        Case "OS.PRODBASEAMT"
            mstrReturn = Format(cobjOrder.ProductionMoBaseAmt, "$###,###,##0.00")
        Case "OS.PRODBASEVOLUME"
            mstrReturn = Format(cobjOrder.ProductionMoBaseVolume, "###,###,##0")
        Case "OS.PRODTOTALAMT"
            mstrReturn = Format(cobjOrder.ProductionTotalBaseAmt, "$###,###,##0.00")
        Case "OS.PRODTOTALVOLUME"
            mstrReturn = Format(cobjOrder.ProductionTotalVolume, "###,###,##0")

        Case "OS.OTHERBASEAMT"
            mstrReturn = Format(cobjOrder.OtherMoBaseAmt, "$###,###,##0.00")
        Case "OS.OTHERBASEVOLUME"
            mstrReturn = Format(cobjOrder.OtherMoBaseVolume, "###,###,##0")
        Case "OS.OTHERTOTALAMT"
            mstrReturn = Format(cobjOrder.OtherTotalBaseAmt, "$###,###,##0.00")
        Case "OS.OTHERTOTALVOLUME"
            mstrReturn = Format(cobjOrder.OtherTotalVolume, "###,###,##0")

        Case "OS.PERIODVOLUME"
            mstrReturn = ""
            For mintLineIx = 1 To cobjOrder.OrderServiceMeters.Count
                If cobjOrder.OrderServiceMeters.Item(mintLineIx).MeterTypeID = mIndex Then
                    mstrReturn = Format(cobjOrder.OrderServiceMeters.Item(mintLineIx).BaseVolume * GetFreqMultiplier(cobjOrder.BillingCycleType.PaymentsPerYear), "###,###,##0")
                    Exit For
                End If
            Next
        
        Case "OS.PERIODTOTAL"
            mstrReturn = ""
            For mintLineIx = 1 To cobjOrder.OrderServiceMeters.Count
                If cobjOrder.OrderServiceMeters.Item(mintLineIx).MeterTypeID = mIndex Then
                    mstrReturn = Format(cobjOrder.OrderServiceMeters.Item(mintLineIx).BaseTotal * GetFreqMultiplier(cobjOrder.BillingCycleType.PaymentsPerYear), "$###,###,##0.00")
                    Exit For
                End If
            Next
            
        Case "OS.MONTHLYVOLUME"
            mstrReturn = ""
            For mintLineIx = 1 To cobjOrder.OrderServiceMeters.Count
                If cobjOrder.OrderServiceMeters.Item(mintLineIx).MeterTypeID = mIndex Then
                    mstrReturn = Format(cobjOrder.OrderServiceMeters.Item(mintLineIx).BaseVolume, "###,###,##0")
                    Exit For
                End If
            Next
        
        Case "OS.BASERATE"
            mstrReturn = ""
            For mintLineIx = 1 To cobjOrder.OrderServiceMeters.Count
                If cobjOrder.OrderServiceMeters.Item(mintLineIx).MeterTypeID = mIndex Then
                    mstrReturn = Format(cobjOrder.OrderServiceMeters.Item(mintLineIx).BaseRate, "$###,###,##0.00000")
                    Exit For
                End If
            Next
            
        Case "OS.BASETOTAL"
            mstrReturn = ""
            For mintLineIx = 1 To cobjOrder.OrderServiceMeters.Count
                If cobjOrder.OrderServiceMeters.Item(mintLineIx).MeterTypeID = mIndex Then
                    mstrReturn = Format(cobjOrder.OrderServiceMeters.Item(mintLineIx).BaseTotal, "$###,###,##0.00")
                    Exit For
                End If
            Next
            
        Case "OS.OVERAGERATE"
            mstrReturn = ""
            For mintLineIx = 1 To cobjOrder.OrderServiceMeters.Count
                If cobjOrder.OrderServiceMeters.Item(mintLineIx).MeterTypeID = mIndex Then
                    mstrReturn = Format(cobjOrder.OrderServiceMeters.Item(mintLineIx).OverageRate, "$###,###,##0.00000")
                    Exit For
                End If
            Next
            
        Case "OL.MFG"
            mstrReturn = pcol.Item(mIndex).CatalogItem.ProductMfg.Name
        Case "OL.MODEL"
            mstrReturn = pcol.Item(mIndex).Model
        Case "OL.QTY"
            mstrReturn = Format(pcol.Item(mIndex).Quantity, "###,##0")
        Case "OL.BUNDLEQTY"
            mstrReturn = Format(pcol.Item(mIndex).BundleQuantity, "###,##0")
        Case "OL.PERBUNDLEQTY"
            mstrReturn = Format(pcol.Item(mIndex).PerBundleQuantity, "###,##0")
        Case "OL.SELLAMOUNT"
            mstrReturn = Format(mcurOB_SellAmount, "$###,###,##0.00")
        Case "OL.BUYAMOUNT"
            mstrReturn = Format(mcurOB_BuyAmount, "$###,###,##0.00")
        Case "OL.LEASEPYMT"
            mstrReturn = Format(mcurOB_LeasePymt, "$###,###,##0.00")
        Case "OL.UNITPRICE"
            mstrReturn = Format(pcol.Item(mIndex).sellprice, "$###,###,##0.00")
        Case "OL.LINETOTAL"
            mstrReturn = Format(pcol.Item(mIndex).LineTotal, "$###,###,##0.00")
        
        Case "OL.DEALERCOST"
            mstrReturn = Format(pcol.Item(mIndex).DealerCostAmt, "$###,###,##0.00")
        Case "OL.ADJDEALERCOST"
            
            mstrReturn = Format(pcol.Item(mIndex).AdjDealerCostAmt, "$###,###,##0.00")
            
        Case "OL.NETDEALERCOST"
            If cobjOrder.SRComMCSMPLevelid = -101 Then
                mstrReturn = Format(cobjOrder.OrderLines.Item(mIndex).AdjDealerCostAmt, "$###,###,##0.00")
            Else
                mstrReturn = Format(cobjOrder.OrderLines.Item(mIndex).DealerCostAmt - cobjOrder.OrderLines.Item(mIndex).SRCommCSMPCreditValue, "$###,###,##0.00")
            End If
        Case "OL.NETREPCOST" 'JB Cost to Rep - any credits
            If cobjOrder.SRComMCSMPLevelid = -101 Then
                mstrReturn = Format(cobjOrder.OrderLines.Item(mIndex).AdjDealerCostAmt, "$###,###,##0.00")
            Else
                mstrReturn = Format((pcol.Item(mIndex).BuyPrice - cobjOrder.OrderLines.Item(mIndex).SRCommCSMPCreditValue), "$###,###,##0.00")
            End If
            
        Case "OL.TOTALSTANDARDBUY" 'JB returns total rep buy price at standard level
            mcurTotal = 0
            For mint = 1 To cobjOrder.OrderLines.Count
              If cobjOrder.OrderLines.Item(mint).CatalogItemID < 0 Then 'manually entered item
                mcurTotal = mcurTotal + (cobjOrder.OrderLines.Item(mint).BuyPrice * cobjOrder.OrderLines.Item(mint).Quantity)
              Else
                For mInt2 = 1 To pcol.Item(mint).CatalogItem.CatalogPrices.Count
                  If pcol.Item(mint).CatalogItem.CatalogPrices(mInt2).PriceLevel.IsStandardInd <> 0 Then
                    mcurTotal = mcurTotal + (pcol.Item(mint).CatalogItem.CatalogPrices(mInt2).BuyPrice * cobjOrder.OrderLines.Item(mint).Quantity)
                    Exit For
                  End If
                Next
              End If
            Next
            mstrReturn = Format(mcurTotal, "$###,###,##0.00")
            
        Case "OL.REVENUEABOVEBASE" 'JB UBT.VA use only
            mcurTotal = 0
            'First Get TotalRepBuy at Standard Price
            For mint = 1 To cobjOrder.OrderLines.Count
              If cobjOrder.OrderLines.Item(mint).CatalogItemID < 0 Then 'manually entered item
               mcurTotal = mcurTotal + (cobjOrder.OrderLines.Item(mint).BuyPrice * cobjOrder.OrderLines.Item(mint).Quantity)
              Else
                For mInt2 = 1 To pcol.Item(mint).CatalogItem.CatalogPrices.Count
                  If pcol.Item(mint).CatalogItem.CatalogPrices(mInt2).PriceLevel.IsStandardInd <> 0 Then
                    mcurTotal = mcurTotal + (pcol.Item(mint).CatalogItem.CatalogPrices(mInt2).BuyPrice * cobjOrder.OrderLines.Item(mint).Quantity)
                    Exit For
                  End If
                Next
              End If
            Next
            
            'Math is TotalStandardBuy - BoardCredit
           mstrReturn = Format((CCur(cobjOrder.BoardCreditAmt) - mcurTotal), "$###,###,##0.00")
          
        Case "OL.PERCENTOVERBASE"
            mcurTotal = 0
            'First Get TotalRepBuy at Standard Price
            For mint = 1 To cobjOrder.OrderLines.Count
              If cobjOrder.OrderLines.Item(mint).CatalogItemID < 0 Then 'manually entered item
               mcurTotal = mcurTotal + (cobjOrder.OrderLines.Item(mint).BuyPrice * cobjOrder.OrderLines.Item(mint).Quantity)
              Else
                For mInt2 = 1 To pcol.Item(mint).CatalogItem.CatalogPrices.Count
                  If pcol.Item(mint).CatalogItem.CatalogPrices(mInt2).PriceLevel.IsStandardInd <> 0 Then
                    mcurTotal = mcurTotal + (pcol.Item(mint).CatalogItem.CatalogPrices(mInt2).BuyPrice * cobjOrder.OrderLines.Item(mint).Quantity)
                    Exit For
                  End If
                Next
              End If
            Next
            If mcurTotal > 0 Then
                mstrReturn = Format(((cobjOrder.BoardCreditAmt - mcurTotal) / mcurTotal), "percent")
            Else
                mstrReturn = Format(0, "percent")
            End If
            
        Case "OL.LEASEPAYMENT"
            If cobjOrder.Subtotalamount = 0 Then
                mcurTotal = 0
            Else
                mcurTotal = (pcol.Item(mIndex).LineTotal / cobjOrder.Subtotalamount) * cobjOrder.PaymentAmtMonthlyFin
            End If
            mstrReturn = Format(mcurTotal, "$###,###,##0.00")
            
        Case "OL.UNITPRICEA"
            mcurTotal = pcol.Item(mIndex).sellprice
            For mintOffsetIx = 1 To pcol.Item(mIndex).OrderLineAdjusts.Count
                mcurTotal = mcurTotal + pcol.Item(mIndex).OrderLineAdjusts.Item(mintOffsetIx).AdjAmount
            Next
            mstrReturn = Format(mcurTotal, "$###,###,##0.00")
        Case "OL.LINETOTALA"
            mcurTotal = pcol.Item(mIndex).LineTotal
            For mintOffsetIx = 1 To pcol.Item(mIndex).OrderLineAdjusts.Count
                mcurTotal = mcurTotal + (pcol.Item(mIndex).OrderLineAdjusts.Item(mintOffsetIx).AdjAmount * pcol.Item(mIndex).Quantity)
            Next
            mstrReturn = Format(mcurTotal, "$###,###,##0.00")
        Case "OL.BUYUNITPRICEPCT"
            If pcol.Item(mIndex).BuyPrice = 0 Then
                mstrReturn = ""
            Else
              If pcol.Item(mIndex).CSMPPCTEntryPCT = 0 Then
                mstrReturn = Format(pcol.Item(mIndex).BuyPrice, "$###,###,##0.00")
             Else 'JB - return whole cost without credit reduction for percent credit CSMP
                mstrReturn = Format(pcol.Item(mIndex).DealerCostAmt, "$###,###,##0.00")
             End If
            End If
        Case "OL.BUYUNITPRICE"
            If pcol.Item(mIndex).BuyPrice = 0 Then
                mstrReturn = ""
            Else
                mstrReturn = Format(pcol.Item(mIndex).BuyPrice, "$###,###,##0.00")
            End If
        Case "OL.BUYLINETOTAL"
            If pcol.Item(mIndex).BuyPrice = 0 Or pcol.Item(mIndex).Quantity = 0 Then
                mstrReturn = ""
            Else
                mstrReturn = Format(pcol.Item(mIndex).BuyPrice * pcol.Item(mIndex).Quantity, "$###,###,##0.00")
            End If
        Case "OL.PRICELEVEL" 'JB - Shows Price Level per order line item
            If pcol.Item(mIndex).PriceLevelID > 0 Or pcol.Item(mIndex).Quantity = 0 Then
                mstrReturn = pcol.Item(mIndex).PriceLevel.Name
            Else
                mstrReturn = "N/A"
            End If
        Case "OL.MSRP"
            mstrReturn = Format(pcol.Item(mIndex).MSRP, "$###,###,##0.00")
        Case "OL.CSMP"
            mstrReturn = Format(pcol.Item(mIndex).SRCommCSMPCreditValue, "$###,###,##0.00")
        Case "OL.CSMPAMT"
            mstrReturn = Format(pcol.Item(mIndex).CSMPPCTAmt, "$###,###,##0.00")
        
        Case "OL.CSMPAMTALL"
            mstrReturn = Format(pcol.Item(mIndex).CSMPPCTAmt + pcol.Item(mIndex).SRCommCSMPCreditValue, "$###,###,##0.00")
        
        
        Case "OL.EXTMSRP"
            mstrReturn = Format(pcol.Item(mIndex).CatalogItem.MFGPrice * pcol.Item(mIndex).Quantity, "$###,###,##0.00")
        Case "OL.MANUALENTRYFLAG"
            If pcol.Item(mIndex).CatalogItemID <= -100 Then     'V2 order form manual entry
                mstrReturn = "*"
            Else
                mstrReturn = ""
            End If
        Case "OL.DESCRIPTION"
            If pcol.Item(mIndex).CatalogItemID <= -100 Then     'V2 order form manual entry
                mstrReturn = pcol.Item(mIndex).Description
            ElseIf pcol.Item(mIndex).CatalogItem.Description <> "" Then
                mstrReturn = pcol.Item(mIndex).CatalogItem.Description
            Else
                mstrReturn = pcol.Item(mIndex).Model
            End If
        Case "OL.MODEL-DESCRIPTION"                             'JB 06/19/2014 will give Model - Description unless same or one is blank
            If pcol.Item(mIndex).CatalogItemID <= -100 Then     'V2 order form manual entry
              If (pcol.Item(mIndex).Description = pcol.Item(mIndex).Model) And pcol.Item(mIndex).Description <> "" Then
                mstrReturn = pcol.Item(mIndex).Description
              ElseIf (pcol.Item(mIndex).Description = "" And pcol.Item(mIndex).Model <> "") Then
                mstrReturn = pcol.Item(mIndex).Model
              ElseIf (pcol.Item(mIndex).Description <> "" And pcol.Item(mIndex).Model = "") Then
                mstrReturn = pcol.Item(mIndex).Description
              ElseIf (InStr(pcol.Item(mIndex).Model, pcol.Item(mIndex).Description) <> 0) Then 'desc is already in model
                mstrReturn = pcol.Item(mIndex).Model
              Else
                mstrReturn = pcol.Item(mIndex).Model + " - " + pcol.Item(mIndex).Description
              End If
            Else
              If (pcol.Item(mIndex).CatalogItem.Description = pcol.Item(mIndex).CatalogItem.Model) And pcol.Item(mIndex).CatalogItem.Description <> "" Then
                mstrReturn = pcol.Item(mIndex).CatalogItem.Description
              ElseIf (pcol.Item(mIndex).CatalogItem.Description = "" And pcol.Item(mIndex).CatalogItem.Model <> "") Then
                mstrReturn = pcol.Item(mIndex).CatalogItem.Model
              ElseIf (pcol.Item(mIndex).CatalogItem.Description <> "" And pcol.Item(mIndex).CatalogItem.Model = "") Then
                mstrReturn = pcol.Item(mIndex).CatalogItem.Description
              ElseIf (InStr(pcol.Item(mIndex).CatalogItem.Model, pcol.Item(mIndex).CatalogItem.Description) <> 0) Then  'desc is already in model
                mstrReturn = pcol.Item(mIndex).CatalogItem.Model
              Else
                mstrReturn = pcol.Item(mIndex).CatalogItem.Model + " - " + pcol.Item(mIndex).CatalogItem.Description
              End If
            End If
        Case "OL.DESCRIPTIONONLY"
            If pcol.Item(mIndex).CatalogItemID <= -100 Then     'V2 order form manual entry
                mstrReturn = pcol.Item(mIndex).Description
            Else
                mstrReturn = pcol.Item(mIndex).CatalogItem.Description
            End If
        Case "OL.ACCUMPTS"
            If pcol.Item(mIndex).AccumPoints <> 0 Then
                mstrReturn = Format(pcol.Item(mIndex).AccumPoints, "###,##0.00")
            Else
                mstrReturn = ""
            End If
        Case "OL.CONPTS"
            If pcol.Item(mIndex).ContestPoints <> 0 Then
                mstrReturn = Format(pcol.Item(mIndex).ContestPoints, "###,##0.00")
            Else
                mstrReturn = ""
            End If
        Case "OL.PAYPTS"
            If pcol.Item(mIndex).PayPoints <> 0 Then
                mstrReturn = Format(pcol.Item(mIndex).PayPoints, "###,##0.00")
            Else
                mstrReturn = ""
            End If
        Case "OL.TACCUMPTS"
            If pcol.Item(mIndex).AccumPoints <> 0 Then
                mstrReturn = Format(pcol.Item(mIndex).Quantity * pcol.Item(mIndex).AccumPoints, "###,##0.00")
            Else
                mstrReturn = ""
            End If
        Case "OL.TCONPTS"
            If pcol.Item(mIndex).ContestPoints <> 0 Then
                mstrReturn = Format(pcol.Item(mIndex).Quantity * pcol.Item(mIndex).ContestPoints, "###,##0.00")
            Else
                mstrReturn = ""
            End If
        Case "OL.TPAYPTS"
            If pcol.Item(mIndex).PayPoints <> 0 Then
                mstrReturn = Format(pcol.Item(mIndex).Quantity * pcol.Item(mIndex).PayPoints, "###,##0.00")
            Else
                mstrReturn = ""
            End If
        Case "OL.COND"
            If pcol.Item(mIndex).Condition = "R" Then
                mstrReturn = "Ref"
            ElseIf pcol.Item(mIndex).Condition = "U" Then
                mstrReturn = "Used"
            Else
                mstrReturn = "New"
            End If
    'Catalog item info
        Case "OL.CI_DESCRIPTION"
            If pcol.Item(mIndex).CatalogItem.Description <> "" Then
                mstrReturn = pcol.Item(mIndex).CatalogItem.Description
            Else
                mstrReturn = pcol.Item(mIndex).Model
            End If
        Case "OL.CI_MSRP"
            mstrReturn = Format(pcol.Item(mIndex).CatalogItem.MFGPrice, "$###,###,##0.00")
        Case "OL.CI_NETWORKIND"
            If pcol.Item(mIndex).CatalogItemID < 1 Then
                mstrReturn = ""
            ElseIf pcol.Item(mIndex).CatalogItem.NetworkedInd = True Then
                mstrReturn = "Yes"
            Else
                mstrReturn = "No"
            End If
        Case "OL.CI_MFGITEMNBR"
            mstrReturn = RemovePreceedingUnderscore(pcol.Item(mIndex).CatalogItem.MFGItemNbr)
        
        Case "OL.CI_IMG_PRINT"
            mstrReturn = pcol.Item(mIndex).CatalogItem.MasterCatalogItem.ImageName_Print
            'mstrReturn = GetCatalogImageName(pcol.Item(mIndex).CatalogItem.MasterCatalogItemID, "P")
        Case "OL.CI_IMG_WEB"
            mstrReturn = pcol.Item(mIndex).CatalogItem.MasterCatalogItem.ImageName_Web
            'mstrReturn = GetCatalogImageName(pcol.Item(mIndex).CatalogItem.MasterCatalogItemID, "W")
        Case "OL.CI_IMG_THUMB"
            mstrReturn = pcol.Item(mIndex).CatalogItem.MasterCatalogItem.ImageName_Thumb
            'mstrReturn = GetCatalogImageName(pcol.Item(mIndex).CatalogItem.MasterCatalogItemID, "T")
        Case "OL.CI_MKTDESCRIPTION"
            mstrReturn = pcol.Item(mIndex).CatalogItem.MasterCatalogItem.MKTDescription
        Case "OL.CI_MKTFEATURES"
            mstrReturn = Trim(pcol.Item(mIndex).CatalogItem.MasterCatalogItem.MKTFeatures)
            If Mid(mstrReturn, 1, 2) = "-/" Then
                mstrReturn = Mid(mstrReturn, 3)
            End If
            mstrReturn = Replace(mstrReturn, "-/", vbCrLf)
        Case "OL.CI_PPMCOLORLTR"
            mstrReturn = Format(pcol.Item(mIndex).CatalogItem.MasterCatalogItem.PPMColorLtr, "###,###,##0")
        Case "OL.CI_PPMBWLTR"
            mstrReturn = Format(pcol.Item(mIndex).CatalogItem.MasterCatalogItem.PPMBWLtr, "###,###,##0")
        Case "OL.CI_PPMCOLORLGL"
            mstrReturn = Format(pcol.Item(mIndex).CatalogItem.MasterCatalogItem.PPMColorLgl, "###,###,##0")
        Case "OL.CI_PPMBWLGL"
            mstrReturn = Format(pcol.Item(mIndex).CatalogItem.MasterCatalogItem.PPMBWLgl, "###,###,##0")
                
        Case "OL.CI_OMDCODE"
            mstrReturn = pcol.Item(mIndex).CatalogItem.OMDCode
        Case "OL.CI_SALESPRICEBOTTOM"
            mstrReturn = Format(pcol.Item(mIndex).CatalogItem.SalesPriceBottom, "$###,###,##0.00")
        Case "OL.CI_SALESPRICETOP"
            mstrReturn = Format(pcol.Item(mIndex).CatalogItem.SalesPriceTop, "$###,###,##0.00")
        Case "OL.CI_POINTSBOTTOM"
            mstrReturn = Format(pcol.Item(mIndex).CatalogItem.PointsBottom, "###,##0.00")
        Case "OL.CI_POINTSTOP"
            mstrReturn = Format(pcol.Item(mIndex).CatalogItem.PointsTop, "###,##0.00")
        Case "OL.CI_WSWIDTH"
            mstrReturn = pcol.Item(mIndex).CatalogItem.WSWidth
        Case "OL.CI_WSDEPTH"
            mstrReturn = pcol.Item(mIndex).CatalogItem.WSDepth
        Case "OL.CI_WSHEIGHT"
            mstrReturn = pcol.Item(mIndex).CatalogItem.WSHeight
        Case "OL.CI_WSBEHIND"
            mstrReturn = pcol.Item(mIndex).CatalogItem.WSBehind
        Case "OL.CI_UNITWIDTH"
            mstrReturn = pcol.Item(mIndex).CatalogItem.UnitWidth
        Case "OL.CI_UNITDEPTH"
            mstrReturn = pcol.Item(mIndex).CatalogItem.UnitDepth
        Case "OL.CI_UNITHEIGHT"
            mstrReturn = pcol.Item(mIndex).CatalogItem.UnitHeight
        Case "OL.CI_UNITWEIGHT"
            mstrReturn = CStr(pcol.Item(mIndex).CatalogItem.UnitWeight)
        Case "OL.CI_PICKUPINFO"
            mstrReturn = pcol.Item(mIndex).CatalogItem.PickupInfo
        Case "OL.CI_PREINSTALLNOTICE"
            mstrReturn = pcol.Item(mIndex).CatalogItem.PreInstallNotice
        Case "OL.CI_GENERALNOTES"
            mstrReturn = pcol.Item(mIndex).CatalogItem.GeneralNotes
        Case "OL.CI_CONFIGINSTRUCTIONS"
            mstrReturn = pcol.Item(mIndex).CatalogItem.ConfigInstructions
        Case "OL.CI_TECHREQUIREDIND"
            If pcol.Item(mIndex).CatalogItem.TechRequiredInd Then
                mstrReturn = "Yes"
            Else
                mstrReturn = "No"
            End If
        Case "OL.CI_REPUNITS"
            mstrReturn = CStr(pcol.Item(mIndex).CatalogItem.RepUnits)
        Case "OL.CI_POWERSUPPLY"
            mstrReturn = pcol.Item(mIndex).CatalogItem.PowerSupply.Name
        Case "OL.GSAPrice"
            mstrReturn = Format(pcol.Item(mIndex).CatalogItem.GSAPrice, "$###,###,##0.0000")
        Case "OL.STANDARDPRICE" 'JB returns standard price regardless of order price level
             If pcol.Item(mIndex).CatalogItemID < 0 Then
                 mstrReturn = Format(pcol.Item(mIndex).BuyPrice, "$###,###,##0.00")
             Else
               For mint = 1 To pcol.Item(mIndex).CatalogItem.CatalogPrices.Count

                 If pcol.Item(mIndex).CatalogItem.CatalogPrices(mint).PriceLevel.IsStandardInd <> 0 Then
                    mstrReturn = Format(pcol.Item(mIndex).CatalogItem.CatalogPrices(mint).BuyPrice, "$###,###,##0.00")
                    Exit For
                 End If
              Next
            End If
    'end of Catalog Item info
        Case "OL.BASECOPIESBW"
            If pcol.Item(mIndex).BaseCopiesBW > 0 Then
                mstrReturn = Format(pcol.Item(mIndex).BaseCopiesBW, "###,###,##0")
            Else
                mstrReturn = ""
            End If
        Case "OL.BASECOPIESCOLOR"
            If pcol.Item(mIndex).BaseCopiesColor > 0 Then
                mstrReturn = Format(pcol.Item(mIndex).BaseCopiesColor, "###,###,##0")
            Else
                mstrReturn = ""
            End If
        Case "OL.BASECOPIESDUALCOLOR"
            If pcol.Item(mIndex).BaseCopiesDualColor > 0 Then
                mstrReturn = Format(pcol.Item(mIndex).BaseCopiesDualColor, "###,###,##0")
            Else
                mstrReturn = ""
            End If
        Case "OL.BASECHARGEBW"
            If pcol.Item(mIndex).BaseChargeBW > 0 Then
                mstrReturn = Format(pcol.Item(mIndex).BaseChargeBW, "$###,###,##0.0000")
            Else
                mstrReturn = ""
            End If
        Case "OL.BASECHARGECOLOR"
            If pcol.Item(mIndex).BaseChargeColor > 0 Then
                mstrReturn = Format(pcol.Item(mIndex).BaseChargeColor, "$###,###,##0.0000")
            Else
                mstrReturn = ""
            End If
        Case "OL.BASECHARGEDUALCOLOR"
            If pcol.Item(mIndex).BaseChargeDualColor > 0 Then
                mstrReturn = Format(pcol.Item(mIndex).BaseChargeDualColor, "$###,###,##0.0000")
            Else
                mstrReturn = ""
            End If
        Case "OL.OVERAGEBW"
            If pcol.Item(mIndex).OverageBW > 0 Then
                mstrReturn = Format(pcol.Item(mIndex).OverageBW, "$###,###,##0.0000")
            Else
                mstrReturn = ""
            End If
        Case "OL.OVERAGECOLOR"
            If pcol.Item(mIndex).OverageColor > 0 Then
                mstrReturn = Format(pcol.Item(mIndex).OverageColor, "$###,###,##0.0000")
            Else
                mstrReturn = ""
            End If
        Case "OL.OVERAGEDUALCOLOR"
            If pcol.Item(mIndex).OverageDualColor > 0 Then
                mstrReturn = Format(pcol.Item(mIndex).OverageDualColor, "$###,###,##0.0000")
            Else
                mstrReturn = ""
            End If
        Case "OL.FIXEDAMOUNT"
            If pcol.Item(mIndex).FixedAmount > 0 Then
                mstrReturn = Format(pcol.Item(mIndex).FixedAmount, "$###,###,##0.00")
            Else
                mstrReturn = ""
            End If
        Case "OL.MAINTTOTAL"
            mcurTotal = 0
            If pcol.Item(mIndex).FixedAmount > 0 Then
                mcurTotal = pcol.Item(mIndex).FixedAmount
            End If
            If pcol.Item(mIndex).BaseChargeBW > 0 Then
                mcurTotal = mcurTotal + _
                    (pcol.Item(mIndex).BaseChargeBW * pcol.Item(mIndex).BaseCopiesBW)
            End If
            If pcol.Item(mIndex).BaseChargeColor > 0 Then
                mcurTotal = mcurTotal + _
                    (pcol.Item(mIndex).BaseChargeColor * pcol.Item(mIndex).BaseCopiesColor)
            End If
            mstrReturn = Format(mcurTotal, "$###,###,##0.00")
        Case "OL.TOTALCHARGEBW"
            mcurTotal = 0
            If pcol.Item(mIndex).BaseChargeBW > 0 Then
                mcurTotal = mcurTotal + _
                    (pcol.Item(mIndex).BaseChargeBW * pcol.Item(mIndex).BaseCopiesBW)
            End If
            mstrReturn = Format(mcurTotal, "$###,###,##0.00")
        Case "OL.TOTALCHARGECOLOR"
            mcurTotal = 0
            If pcol.Item(mIndex).BaseChargeColor > 0 Then
                mcurTotal = mcurTotal + _
                    (pcol.Item(mIndex).BaseChargeColor * pcol.Item(mIndex).BaseCopiesColor)
            End If
            mstrReturn = Format(mcurTotal, "$###,###,##0.00")
        Case "OL.TOTALCHARGEDUALCOLOR"
            mcurTotal = 0
            If pcol.Item(mIndex).BaseChargeDualColor > 0 Then
                mcurTotal = mcurTotal + _
                    (pcol.Item(mIndex).BaseChargeDualColor * pcol.Item(mIndex).BaseCopiesDualColor)
            End If
            mstrReturn = Format(mcurTotal, "$###,###,##0.00")
            
        Case "OL.ANNUALSERVICE"         'V2
            mstrReturn = Format(cobjOrder.OrderLines(mIndex).FixedAmount, "$###,###,##0.00")
        Case "OL.TOTALANNUALSERVICE"    'V2
            mstrReturn = Format(cobjOrder.OrderLines(mIndex).BundleQuantity * cobjOrder.OrderLines(mIndex).FixedAmount, "$###,###,##0.00")
        
        'V2 - Postage Meter (custom for Upstate.NY)
        Case "OL.POSTAGEMETERRENT"          'Per unit for this line item
            mstrReturn = Format(cobjOrder.OrderLines(mIndex).MonthlyMeter, "$###,###,##0.00")
        Case "OL.POSTAGEMETERRENTLINE"      'Total for this line item
            mstrReturn = Format(cobjOrder.OrderLines(mIndex).BundleQuantity * cobjOrder.OrderLines(mIndex).MonthlyMeter, "$###,###,##0.00")
        Case "OL.POSTAGEANNUAL"             'Per unit for this line item
            mstrReturn = Format(cobjOrder.OrderLines(mIndex).RateAndStructure, "$###,###,##0.00")
        Case "OL.POSTAGEANNUALLINE"         'Total for this line item
            mstrReturn = Format(cobjOrder.OrderLines(mIndex).BundleQuantity * cobjOrder.OrderLines(mIndex).RateAndStructure, "$###,###,##0.00")
            
        Case "QA.NAME"
            mstrReturn = cobjOrder.OrderAdjustments.Item(mintLineIx).OrderAdjType.OrderAdjLabel
        Case "QA.AMT"
            If cobjOrder.OrderAdjustments.Item(mintLineIx).AdjAmount = 0 Then
                mstrReturn = ""
            ElseIf cobjOrder.OrderAdjustments.Item(mintLineIx).OrderAdjType.CreditDebitInd = "D" Then
                mstrReturn = Format(cobjOrder.OrderAdjustments.Item(mintLineIx).AdjAmount, "$###,###,##0.00")
            Else
                mstrReturn = Format(cobjOrder.OrderAdjustments.Item(mintLineIx).AdjAmount, "($###,###,##0.00)")
            End If
        Case "QA.OPAMT"
            If cobjOrder.OrderAdjustments.Item(mintLineIx).AdjAmount = 0 Then
                mstrReturn = ""
            ElseIf cobjOrder.OrderAdjustments.Item(mintLineIx).OrderAdjType.CreditDebitInd = "D" Then
                mstrReturn = Format(cobjOrder.OrderAdjustments.Item(mintLineIx).AdjAmount, "-$###,###,##0.00")
            Else
                mstrReturn = Format(cobjOrder.OrderAdjustments.Item(mintLineIx).AdjAmount, "$###,###,##0.00")
            End If
        Case "QA.LABELENTRY"
            If cobjOrder.OrderAdjustments.Item(mintLineIx).OrderAdjType.labelname <> "" Then
                mstrReturn = cobjOrder.OrderAdjustments.Item(mintLineIx).LabelEntry
            Else
                mstrReturn = ""
            End If
        Case "QA.LABELNAME"
            mstrReturn = cobjOrder.OrderAdjustments.Item(mintLineIx).OrderAdjType.labelname
        Case "QA.ACCUMPTS"
            If cobjOrder.OrderAdjustments.Item(mintLineIx).AccumPoints <> 0 Then
                mstrReturn = Format(cobjOrder.OrderAdjustments.Item(mintLineIx).AccumPoints, "###,###,##0.00")
            Else
                mstrReturn = ""
            End If
        Case "QA.CONPTS"
            If cobjOrder.OrderAdjustments.Item(mintLineIx).ContestPoints <> 0 Then
                mstrReturn = Format(cobjOrder.OrderAdjustments.Item(mintLineIx).ContestPoints, "###,###,##0.00")
            Else
                mstrReturn = ""
            End If
        Case "QA.PAYPTS"
            If cobjOrder.OrderAdjustments.Item(mintLineIx).PayPoints <> 0 Then
                mstrReturn = Format(cobjOrder.OrderAdjustments.Item(mintLineIx).PayPoints, "###,###,##0.00")
            Else
                mstrReturn = ""
            End If

        Case "CM.SRCUSER"
            mstrReturn = cobjOrder.SRCommEarned.Item(mintLineIx).User.Person.Fullname
        
        Case "CM.SRCROLE"
            mstrReturn = cobjOrder.SRCommEarned.Item(mintLineIx).SRCommRole.Name
            
        Case "CM.SRCSPLITPCT"
            If cobjOrder.SRCommEarned.Item(mintLineIx).SplitPCT <> 0 Then
                mstrReturn = Format(cobjOrder.SRCommEarned.Item(mintLineIx).SplitPCT, "##0.00%")
            Else
                mstrReturn = "0.00%"
            End If
        Case "CM.SRCAMT"
            If cobjOrder.SRCommEarned.Item(mintLineIx).EarnedCommAmt = 0 Then
                mstrReturn = ""
            Else
                mstrReturn = Format(cobjOrder.SRCommEarned.Item(mintLineIx).EarnedCommAmt, "$###,###,##0.00")
            End If
        Case "OR.SRCPRIAMT"
            mcurTotal = 0
            For mint = 1 To cobjOrder.SRCommEarned.Count
                If cobjOrder.SRCommEarned.Item(mint).UserID = cobjOrder.SalesRepUserID Then
                    mcurTotal = mcurTotal + cobjOrder.SRCommEarned.Item(mint).EarnedCommAmt
                End If
            Next
            mstrReturn = Format(mcurTotal, "$###,###,##0.00")
        
        Case "OR.SRCPRIAMTEQCOMM" 'JB Comm Pct * Adj GP For Primary Rep
            mcurTotal = 0
            mcurTotal = cobjOrder.AdjGPAmt * cobjOrder.SRCommPCTPaid
            mstrReturn = Format(mcurTotal, "$###,###,##0.00")
        
        Case "OR.SVMOCOST"
            If cobjOrder.SVMOCOST = 0 Then
                mstrReturn = ""
            Else
                mstrReturn = Format(cobjOrder.SVMOCOST, "$###,###,##0.00")
            End If
        Case "OR.SVMOREPCOST"
            If cobjOrder.SVMOREPCOST = 0 Then
                mstrReturn = ""
            Else
                mstrReturn = Format(cobjOrder.SVMOREPCOST, "$###,###,##0.00")
            End If
        Case "OR.SVMOSELLPRICE"
            If cobjOrder.SVMOSELLPRICE = 0 Then
                mstrReturn = ""
            Else
                mstrReturn = Format(cobjOrder.SVMOSELLPRICE, "$###,###,##0.00")
            End If
        Case "OR.SVMOGP"
            If cobjOrder.SVMOGP = 0 Then
                mstrReturn = ""
            Else
                mstrReturn = Format(cobjOrder.SVMOGP, "$###,###,##0.00")
            End If
        Case "OR.SVANNUALPRICE"
            If cobjOrder.SVANNUALPRICE = 0 Then
                mstrReturn = ""
            Else
                mstrReturn = Format(cobjOrder.SVANNUALPRICE, "$###,###,##0.00")
            End If
        Case "OR.SVANNUALGP"
            If cobjOrder.SVANNUALGP = 0 Then
                mstrReturn = ""
            Else
                mstrReturn = Format(cobjOrder.SVANNUALGP, "$###,###,##0.00")
            End If
        Case "OR.SVTERM"
            If cobjOrder.SVTERM = 0 Then
                mstrReturn = ""
            Else
                mstrReturn = Format(cobjOrder.SVTERM, "###,##0")
            End If
        Case "OR.SVNBRUSERS"
            If cobjOrder.SVNBRUSERS = 0 Then
                mstrReturn = ""
            Else
                mstrReturn = Format(cobjOrder.SVNBRUSERS, "###,##0")
            End If
        Case "OR.SVCOSTPERUSER"
            If cobjOrder.SVCOSTPERUSER = 0 Then
                mstrReturn = ""
            Else
                mstrReturn = Format(cobjOrder.SVCOSTPERUSER, "$###,###,##0.00")
            End If
        Case "OR.SVSTARTDT"
            If cobjOrder.SVSTARTDT = "1/1/1900" Then
                mstrReturn = ""
            Else
                mstrReturn = Format(cobjOrder.SVSTARTDT, "MM/DD/YYYY")
            End If
        Case "OR.SVTOTALVALUE"
            If cobjOrder.SVTOTALVALUE = 0 Then
                mstrReturn = ""
            Else
                mstrReturn = Format(cobjOrder.SVTOTALVALUE, "$###,###,##0.00")
            End If
        Case "OR.SVTOTALGP"
            If cobjOrder.SVTOTALGP = 0 Then
                mstrReturn = ""
            Else
                mstrReturn = Format(cobjOrder.SVTOTALGP, "$###,###,##0.00")
            End If
        Case "OR.WTMOREVAMT"
            If cobjOrder.WTMOREVAMT = 0 Then
                mstrReturn = ""
            Else
                mstrReturn = Format(cobjOrder.WTMOREVAMT, "$###,###,##0.00")
            End If
        Case "OR.WTMOGPAMT"
            If cobjOrder.WTMOGPAMT = 0 Then
                mstrReturn = ""
            Else
                mstrReturn = Format(cobjOrder.WTMOGPAMT, "$###,###,##0.00")
            End If
        Case "OR.WTTERM"
            If cobjOrder.WTTERM = 0 Then
                mstrReturn = ""
            Else
                mstrReturn = Format(cobjOrder.WTTERM, "###,##0")
            End If
        Case "OR.SVBILLCYCLENAME"
            mstrReturn = cobjOrder.SVBillingCycleType.Name
        Case "OR.SVBILLCYCLENAMEOTHER"
            mstrReturn = cobjOrder.SVBillingCycleType.OtherName
        Case "OR.SVBILLCYCLEPPY"
            If cobjOrder.SVBillingCycleType.PaymentsPerYear = 0 Then
                mstrReturn = ""
            Else
                mstrReturn = Format(cobjOrder.SVBillingCycleType.PaymentsPerYear, "###,##0")
            End If
        Case "OR.SVBASEBILLFREQMOX"
            If cobjOrder.SVBillingCycleType.PaymentsPerYear = 12 Then
                mstrReturn = "X"
            Else
                mstrReturn = ""
            End If
        Case "OR.SVBASEBILLFREQQTRX"
            If cobjOrder.SVBillingCycleType.PaymentsPerYear = 4 Then
                mstrReturn = "X"
            Else
                mstrReturn = ""
            End If
        Case "OR.SVBASEBILLFREQSAX"
            If cobjOrder.SVBillingCycleType.PaymentsPerYear = 2 Then
                mstrReturn = "X"
            Else
                mstrReturn = ""
            End If
        Case "OR.SVBASEBILLFREQANX"
            If cobjOrder.SVBillingCycleType.PaymentsPerYear = 1 Then
                mstrReturn = "X"
            Else
                mstrReturn = ""
            End If
        
        Case "IL.QUANTITY"
            mstrReturn = Format(pcol.Item(mIndex).Quantity, "###,###,##0")
        Case "IL.MOCOST"
            mstrReturn = Format(pcol.Item(mIndex).MOCOST, "$###,###,##0.00")
        Case "IL.MOTOTALCOST"
            mstrReturn = Format(pcol.Item(mIndex).MOCOST * pcol.Item(mIndex).Quantity, "$###,###,##0.00")
        Case "IL.MARKUPPCT"
            mstrReturn = Format(pcol.Item(mIndex).Markup, "###,##0.00%")
        Case "IL.MOREPCOST"
            mstrReturn = Format(pcol.Item(mIndex).MOREPCOST, "$###,###,##0.00")
        Case "IL.MOTOTALREPCOST"
            mstrReturn = Format(pcol.Item(mIndex).MOREPCOST * pcol.Item(mIndex).Quantity, "$###,###,##0.00")
        Case "IL.MOSELLPRICE"
            mstrReturn = Format(pcol.Item(mIndex).MOSELLPRICE, "$###,###,##0.00")
        Case "IL.MOTOTALSELLPRICE"
            mstrReturn = Format(pcol.Item(mIndex).MOSELLPRICE * pcol.Item(mIndex).Quantity, "$###,###,##0.00")
        Case "IL.MOGP"
            mstrReturn = Format(pcol.Item(mIndex).MOGP, "$###,###,##0.00")
        Case "IL.TOTALCONTRACTPRICE"
            mstrReturn = Format(pcol.Item(mIndex).MOTOTALPRICE, "$###,###,##0.00")
        Case "IL.TOTALCONTRACTGP"
            mstrReturn = Format(pcol.Item(mIndex).MOTOTALGP, "$###,###,##0.00")
        Case "IL.ANNUALPRICE"
            mstrReturn = Format(pcol.Item(mIndex).ANNUALPRICE, "$###,###,##0.00")
        Case "IL.ANNUALGP"
            mstrReturn = Format(pcol.Item(mIndex).ANNUALGP, "$###,###,##0.00")
        Case "IL.SI_SKU"
            mstrReturn = pcol.Item(mIndex).ServicesItem.SKU
        Case "IL.SI_VENDOR"
            mstrReturn = pcol.Item(mIndex).ServicesItem.VENDOR
        Case "IL.SI_MODEL"
            mstrReturn = pcol.Item(mIndex).ServicesItem.Model
        Case "IL.SI_DESCRIPTION"
            mstrReturn = pcol.Item(mIndex).ServicesItem.Description
        Case "IL.SI_MOCOST"
            mstrReturn = Format(pcol.Item(mIndex).ServicesItem.MOCOST, "$###,###,##0.00")
        Case "IL.SI_MARKUPPCT"
            mstrReturn = Format(pcol.Item(mIndex).ServicesItem.Markup, "###,##0.00%")
        Case "IL.SI_MOREPCOST"
            mstrReturn = Format(pcol.Item(mIndex).ServicesItem.MOREPCOST, "$###,###,##0.00")
        Case "IL.SI_MSRP"
            mstrReturn = Format(pcol.Item(mIndex).ServicesItem.MSRP, "$###,###,##0.00")

        Case "CM.SRCBC$"
            If cobjOrder.SRCommEarned.Item(mintLineIx).BoardCreditAmt = 0 Then
                mstrReturn = ""
            Else
                mstrReturn = Format(cobjOrder.SRCommEarned.Item(mintLineIx).BoardCreditAmt, "$###,###,##0.00")
            End If
        Case "CM.SRCREV$"
            If cobjOrder.SRCommEarned.Item(mintLineIx).GrossRevenueAmt = 0 Then
                mstrReturn = ""
            Else
                mstrReturn = Format(cobjOrder.SRCommEarned.Item(mintLineIx).GrossRevenueAmt, "$###,###,##0.00")
            End If
        Case "CM.SRCGP$"
            If cobjOrder.SRCommEarned.Item(mintLineIx).GPBoardCreditAmt = 0 Then
                mstrReturn = ""
            Else
                mstrReturn = Format(cobjOrder.SRCommEarned.Item(mintLineIx).GPBoardCreditAmt, "$###,###,##0.00")
            End If
        Case "CM.SRCAGP$"
            If cobjOrder.SRCommEarned.Item(mintLineIx).AdjGPAmt = 0 Then
                mstrReturn = ""
            Else
                mstrReturn = Format(cobjOrder.SRCommEarned.Item(mintLineIx).AdjGPAmt, "$###,###,##0.00")
            End If
            
        Case "OA.NAME"
            mstrReturn = cobjOrder.GetRelativeOrderAdjTypes.Item(mIndex).OrderAdjLabel
            If Len(Trim(cobjOrder.GetRelativeOrderAdjTypes.Item(mIndex).labelname)) > 0 Then
                mstrReturn = mstrReturn & ":[" & cobjOrder.OrderAdjustmentByTypeID(cobjOrder.GetRelativeOrderAdjTypes.Item(mIndex).OrderAdjTypeID).LabelEntry & "]"
            End If
        Case "OA.AMT"
        
            If cobjOrder.GetRelativeOrderAdjTypes.Item(mIndex).CreditDebitInd = "D" Then
                mstrReturn = IIf(cobjOrder.OrderAdjustmentByTypeID(cobjOrder.GetRelativeOrderAdjTypes.Item(mIndex).OrderAdjTypeID).AdjAmount = 0, "", Format(cobjOrder.OrderAdjustmentByTypeID(cobjOrder.GetRelativeOrderAdjTypes.Item(mIndex).OrderAdjTypeID).AdjAmount, "$###,###,##0.00"))
            Else
                mstrReturn = IIf(cobjOrder.OrderAdjustmentByTypeID(cobjOrder.GetRelativeOrderAdjTypes.Item(mIndex).OrderAdjTypeID).AdjAmount = 0, "", Format(cobjOrder.OrderAdjustmentByTypeID(cobjOrder.GetRelativeOrderAdjTypes.Item(mIndex).OrderAdjTypeID).AdjAmount, "($###,###,##0.00)"))
            End If
        Case "OA.ACCUMPTS"
            If cobjOrder.OrderAdjustmentByTypeID(cobjOrder.GetRelativeOrderAdjTypes.Item(mIndex).OrderAdjTypeID).AccumPoints <> 0 Then
                mstrReturn = Format(cobjOrder.OrderAdjustmentByTypeID(cobjOrder.GetRelativeOrderAdjTypes.Item(mIndex).OrderAdjTypeID).AccumPoints, "###,###,##0.00")
            Else
                mstrReturn = ""
            End If
        Case "OA.CONPTS"
            If cobjOrder.OrderAdjustmentByTypeID(cobjOrder.GetRelativeOrderAdjTypes.Item(mIndex).OrderAdjTypeID).ContestPoints <> 0 Then
                mstrReturn = Format(cobjOrder.OrderAdjustmentByTypeID(cobjOrder.GetRelativeOrderAdjTypes.Item(mIndex).OrderAdjTypeID).ContestPoints, "###,###,##0.00")
            Else
                mstrReturn = ""
            End If
        Case "OA.PAYPTS"
            If cobjOrder.OrderAdjustmentByTypeID(cobjOrder.GetRelativeOrderAdjTypes.Item(mIndex).OrderAdjTypeID).PayPoints <> 0 Then
                mstrReturn = Format(cobjOrder.OrderAdjustmentByTypeID(cobjOrder.GetRelativeOrderAdjTypes.Item(mIndex).OrderAdjTypeID).PayPoints, "###,###,##0.00")
            Else
                mstrReturn = ""
            End If
        Case "OA.SELECTOPT"
            mstrReturn = cobjOrder.OrderAdjustmentByTypeID(cobjOrder.GetRelativeOrderAdjTypes.Item(mIndex).OrderAdjTypeID).SelectOpt
        Case "OX.TYPE"
            mstrReturn = cobjOrder.OrderExpenseByID(cobjOrder.GetRelativeOrderExpenses.Item(mIndex).ExpensesID).Expenses.ExpenseType.ExpenseType
        Case "OX.NAME"
            mstrReturn = cobjOrder.OrderExpenseByID(cobjOrder.GetRelativeOrderExpenses.Item(mIndex).ExpensesID).Expenses.Name
        Case "OX.AMT"
            If cobjOrder.OrderExpenseByID(cobjOrder.GetRelativeOrderExpenses.Item(mIndex).ExpensesID).Amount <> 0 Then
                mstrReturn = Format(cobjOrder.OrderExpenseByID(cobjOrder.GetRelativeOrderExpenses.Item(mIndex).ExpensesID).Amount, "$###,###,##0.00")
            End If
        Case "OX.LABEL"
            mstrReturn = cobjOrder.OrderExpenseByID(cobjOrder.GetRelativeOrderExpenses.Item(mIndex).ExpensesID).Label
        Case "OX.EXPTYPE1TOTAL"
           
            
            Set mrsSums = cobjOrder.GetTotalsByType()
            If Not mrsSums.EOF And Not mrsSums.BOF Then mrsSums.MoveFirst
            For mint = 1 To mrsSums.RecordCount
                If mint = 1 Then
                    mstrReturn = Format(mrsSums("TOTAL"), "$###,###,##0.00")
                End If
                mrsSums.MoveNext
            Next
            Set mrsSums = Nothing
        Case "OX.EXPTYPE2TOTAL"
           
            Set mrsSums = cobjOrder.GetTotalsByType()
            If Not mrsSums.EOF And Not mrsSums.BOF Then mrsSums.MoveFirst
            For mint = 1 To mrsSums.RecordCount
                If mint = 2 Then
                    mstrReturn = Format(mrsSums("TOTAL"), "$###,###,##0.00")
                End If
                mrsSums.MoveNext
            Next
            Set mrsSums = Nothing
        Case "OX.EXPTYPE3TOTAL"
           
            Set mrsSums = cobjOrder.GetTotalsByType()
            If Not mrsSums.EOF And Not mrsSums.BOF Then mrsSums.MoveFirst
            For mint = 1 To mrsSums.RecordCount
                If mint = 3 Then
                    mstrReturn = Format(mrsSums("TOTAL"), "$###,###,##0.00")
                End If
                mrsSums.MoveNext
            Next
            Set mrsSums = Nothing
        Case "OE.MODEL"
            mstrReturn = cobjOrder.OrderEqMoves.Item(mIndex).Model
        Case "OE.PUORMOVE"
            If cobjOrder.OrderEqMoves.Item(mIndex).Pickupmoveind = "P" Then
                mstrReturn = "Pickup"
            ElseIf cobjOrder.OrderEqMoves.Item(mIndex).Pickupmoveind = "M" Then
                mstrReturn = "Move"
            Else
                mstrReturn = ""
            End If
        Case "OE.SERIAL"
            mstrReturn = cobjOrder.OrderEqMoves.Item(mIndex).SerialNbr
        Case "OE.WEIGHT"
            mstrReturn = ""
        
        Case "ON.MODEL"
            mstrReturn = cobjOrder.OrderLines.Item(mintLineIx).Model
        Case "ON.DESCRIPTION"
            If cobjOrder.OrderLines.Item(mintLineIx).CatalogItem.Description <> "" Then
                mstrReturn = cobjOrder.OrderLines.Item(mintLineIx).CatalogItem.Description
            Else
                mstrReturn = cobjOrder.OrderLines.Item(mintLineIx).Model
            End If
        Case "ON.ITEMNBR"
            mstrReturn = RemovePreceedingUnderscore(cobjOrder.OrderLines.Item(mintLineIx).CatalogItem.MFGItemNbr)
        Case "ON.UNITPRICE"
            mstrReturn = Format(cobjOrder.OrderLines.Item(mintLineIx).sellprice, "$###,###,##0.00")
                        
        Case "OM.SERIAL"
            mstrReturn = cobjOrder.PrimaryOrderLines.Item(mintLineIx).OrderMachines.Item(mintMachineIx).SerialNumber
        Case "OM.EQID"
            mstrReturn = cobjOrder.PrimaryOrderLines.Item(mintLineIx).OrderMachines.Item(mintMachineIx).AssetTag
        Case "OM.METER"
            mstrReturn = cobjOrder.PrimaryOrderLines.Item(mintLineIx).OrderMachines.Item(mintMachineIx).MeterRead
        Case "OM.COMPANYNAME", "OU.COMPANYNAME", "OV.COMPANYNAME"
            mstrReturn = cobjOrder.PrimaryOrderLines.Item(mintLineIx).OrderMachines.Item(mintMachineIx).CompanyName
        Case "OM.ADDRESSLABEL", "OU.ADDRESSLABEL", "OV.ADDRESSLABEL"
            mstrReturn = cobjOrder.PrimaryOrderLines.Item(mintLineIx).OrderMachines.Item(mintMachineIx).AddressLabel
        Case "OM.CONTACTNAME", "OU.CONTACTNAME", "OV.CONTACTNAME"
            mstrReturn = cobjOrder.PrimaryOrderLines.Item(mintLineIx).OrderMachines.Item(mintMachineIx).ContactName
            If InStr(mstrReturn, ",") > 0 Then
                mstrReturn = Trim(Mid(mstrReturn, InStr(mstrReturn, ",") + 1)) & " " & Trim(Mid(mstrReturn, 1, InStr(mstrReturn, ",") - 1))
            End If
        Case "OM.PHONE", "OU.PHONE", "OV.PHONE"
            mstrReturn = cobjOrder.PrimaryOrderLines.Item(mintLineIx).OrderMachines.Item(mintMachineIx).ContactPhone
        Case "OM.CELL", "OU.CELL", "OV.CELL"
            mstrReturn = cobjOrder.PrimaryOrderLines.Item(mintLineIx).OrderMachines.Item(mintMachineIx).ContactCell
        Case "OM.FAX", "OU.FAX", "OV.FAX"
            mstrReturn = cobjOrder.PrimaryOrderLines.Item(mintLineIx).OrderMachines.Item(mintMachineIx).ContactFAX
        Case "OM.EMAIL", "OU.EMAIL", "OV.EMAIL"
            mstrReturn = LCase(cobjOrder.PrimaryOrderLines.Item(mintLineIx).OrderMachines.Item(mintMachineIx).ContactEmail)
        Case "OM.DELIVERYDATE", "OU.DELIVERYDATE", "OV.DELIVERYDATE"
            mstrReturn = Format(cobjOrder.PrimaryOrderLines.Item(mintLineIx).OrderMachines.Item(mintMachineIx).DeliveryDate, "mm/dd/yyyy")
        
        Case "OM.REASONFORTECH", "OU.REASONFORTECH", "OV.REASONFORTECH"
            mstrReturn = cobjOrder.PrimaryOrderLines.Item(mintLineIx).OrderMachines.Item(mintMachineIx).ReasonForTech
        Case "OM.DELIVERYINSTRUCTIONS", "OU.DELIVERYINSTRUCTIONS", "OV.DELIVERYINSTRUCTIONS"
            mstrReturn = cobjOrder.PrimaryOrderLines.Item(mintLineIx).OrderMachines.Item(mintMachineIx).DeliveryInstructions
        
        Case "OM.STAIRS", "OU.STAIRS", "OV.STAIRS"
            If cobjOrder.PrimaryOrderLines.Item(mintLineIx).OrderMachines.Item(mintMachineIx).Stairs Then
                mstrReturn = "Yes"
            Else
                mstrReturn = "No"
            End If
        Case "OM.NUMSTAIRS", "OU.NUMSTAIRS", "OV.NUMSTAIRS"
            mstrReturn = cobjOrder.PrimaryOrderLines.Item(mintLineIx).OrderMachines.Item(mintMachineIx).NumStairs
        Case "OM.ELEVATOR", "OU.ELEVATOR", "OV.ELEVATOR"
            If cobjOrder.PrimaryOrderLines.Item(mintLineIx).OrderMachines.Item(mintMachineIx).Elevator Then
                mstrReturn = "Yes"
            Else
                mstrReturn = "No"
            End If
        Case "OM.TRUCKPICKUP"
            If cobjOrder.PrimaryOrderLines.Item(mintLineIx).OrderMachines.Item(mintMachineIx).TruckPickup Then
                mstrReturn = "Yes"
            Else
                mstrReturn = "No"
            End If
        Case "OM.NETREQ"
            If cobjOrder.PrimaryOrderLines.Item(mintLineIx).OrderMachines.Item(mintMachineIx).NetRequired Then
                mstrReturn = "Yes"
            Else
                mstrReturn = "No"
            End If
        Case "OM.SALESREPREQ"
            If cobjOrder.PrimaryOrderLines.Item(mintLineIx).OrderMachines.Item(mintMachineIx).SalesRepRequested Then
                mstrReturn = "Yes"
            Else
                mstrReturn = "No"
            End If
        Case "OM.CSRREQ"
            If cobjOrder.PrimaryOrderLines.Item(mintLineIx).OrderMachines.Item(mintMachineIx).CSRRequested Then
                mstrReturn = "Yes"
            Else
                mstrReturn = "No"
            End If
        Case "OM.UPDCONNECTEDEQ"
            If cobjOrder.PrimaryOrderLines.Item(mintLineIx).OrderMachines.Item(mintMachineIx).UpdateConnectedEq Then
                mstrReturn = "Yes"
            Else
                mstrReturn = "No"
            End If
        Case "OM.DIGINSTALL"
            If cobjOrder.PrimaryOrderLines.Item(mintLineIx).OrderMachines.Item(mintMachineIx).DigitalInstall Then
                mstrReturn = "Yes"
            Else
                mstrReturn = "No"
            End If
        Case "OM.FAXSETUP"
            If cobjOrder.PrimaryOrderLines.Item(mintLineIx).OrderMachines.Item(mintMachineIx).FAXSetup Then
                mstrReturn = "Yes"
            Else
                mstrReturn = "No"
            End If
        Case "OM.TECHREQUESTED"
            If cobjOrder.PrimaryOrderLines.Item(mintLineIx).OrderMachines.Item(mintMachineIx).TechRequested Then
                mstrReturn = "Yes"
            Else
                mstrReturn = "No"
            End If
        Case "OM.PICKUPTYPE1"
            If cobjOrder.PrimaryOrderLines.Item(mintLineIx).OrderMachines.Item(mintMachineIx).Pickup1Type = "P" Then
                mstrReturn = "Pickup"
            ElseIf cobjOrder.PrimaryOrderLines.Item(mintLineIx).OrderMachines.Item(mintMachineIx).Pickup1Type = "M" Then
                mstrReturn = "Move"
            Else
                mstrReturn = ""
            End If
        Case "OM.PICKUPMODEL1"
            mstrReturn = cobjOrder.PrimaryOrderLines.Item(mintLineIx).OrderMachines.Item(mintMachineIx).Pickup1Model
        Case "OM.PICKUPSERIAL1"
            mstrReturn = cobjOrder.PrimaryOrderLines.Item(mintLineIx).OrderMachines.Item(mintMachineIx).Pickup1Serial
        Case "OM.PICKUPASSETTAG1"
            mstrReturn = cobjOrder.PrimaryOrderLines.Item(mintLineIx).OrderMachines.Item(mintMachineIx).Pickup1AssetTag
        Case "OM.PICKUPTYPE2"
            If cobjOrder.PrimaryOrderLines.Item(mintLineIx).OrderMachines.Item(mintMachineIx).Pickup2Type = "P" Then
                mstrReturn = "Pickup"
            ElseIf cobjOrder.PrimaryOrderLines.Item(mintLineIx).OrderMachines.Item(mintMachineIx).Pickup2Type = "M" Then
                mstrReturn = "Move"
            Else
                mstrReturn = ""
            End If
        Case "OM.PICKUPMODEL2"
            mstrReturn = cobjOrder.PrimaryOrderLines.Item(mintLineIx).OrderMachines.Item(mintMachineIx).Pickup2Model
        Case "OM.PICKUPSERIAL2"
            mstrReturn = cobjOrder.PrimaryOrderLines.Item(mintLineIx).OrderMachines.Item(mintMachineIx).Pickup2Serial
        Case "OM.PICKUPASSETTAG2"
            mstrReturn = cobjOrder.PrimaryOrderLines.Item(mintLineIx).OrderMachines.Item(mintMachineIx).Pickup2AssetTag
        Case "OM.PICKUPTYPE3"
            If cobjOrder.PrimaryOrderLines.Item(mintLineIx).OrderMachines.Item(mintMachineIx).Pickup3Type = "P" Then
                mstrReturn = "Pickup"
            ElseIf cobjOrder.PrimaryOrderLines.Item(mintLineIx).OrderMachines.Item(mintMachineIx).Pickup3Type = "M" Then
                mstrReturn = "Move"
            Else
                mstrReturn = ""
            End If
        Case "OM.PICKUPMODEL3"
            mstrReturn = cobjOrder.PrimaryOrderLines.Item(mintLineIx).OrderMachines.Item(mintMachineIx).Pickup3Model
        Case "OM.PICKUPSERIAL3"
            mstrReturn = cobjOrder.PrimaryOrderLines.Item(mintLineIx).OrderMachines.Item(mintMachineIx).Pickup3Serial
        Case "OM.PICKUPASSETTAG3"
            mstrReturn = cobjOrder.PrimaryOrderLines.Item(mintLineIx).OrderMachines.Item(mintMachineIx).Pickup3AssetTag
            
        'Machine Pickup/Move
        Case "OO.TYPE"
            If mintOffsetIx = 1 Then
                mstrTemp = cobjOrder.PrimaryOrderLines.Item(mintLineIx).OrderMachines.Item(mintMachineIx).Pickup1Type
            ElseIf mintOffsetIx = 2 Then
                mstrTemp = cobjOrder.PrimaryOrderLines.Item(mintLineIx).OrderMachines.Item(mintMachineIx).Pickup2Type
            ElseIf mintOffsetIx = 3 Then
                mstrTemp = cobjOrder.PrimaryOrderLines.Item(mintLineIx).OrderMachines.Item(mintMachineIx).Pickup3Type
            End If
            If mstrTemp = "P" Then
                mstrReturn = "Pickup"
            ElseIf mstrTemp = "M" Then
                mstrReturn = "Move"
            End If
        Case "OO.MODEL"
            If mintOffsetIx = 1 Then
                mstrReturn = cobjOrder.PrimaryOrderLines.Item(mintLineIx).OrderMachines.Item(mintMachineIx).Pickup1Model
            ElseIf mintOffsetIx = 2 Then
                mstrReturn = cobjOrder.PrimaryOrderLines.Item(mintLineIx).OrderMachines.Item(mintMachineIx).Pickup2Model
            ElseIf mintOffsetIx = 3 Then
                mstrReturn = cobjOrder.PrimaryOrderLines.Item(mintLineIx).OrderMachines.Item(mintMachineIx).Pickup3Model
            End If
        Case "OO.SERIAL"
            If mintOffsetIx = 1 Then
                mstrReturn = cobjOrder.PrimaryOrderLines.Item(mintLineIx).OrderMachines.Item(mintMachineIx).Pickup1Serial
            ElseIf mintOffsetIx = 2 Then
                mstrReturn = cobjOrder.PrimaryOrderLines.Item(mintLineIx).OrderMachines.Item(mintMachineIx).Pickup2Serial
            ElseIf mintOffsetIx = 3 Then
                mstrReturn = cobjOrder.PrimaryOrderLines.Item(mintLineIx).OrderMachines.Item(mintMachineIx).Pickup3Serial
            End If
        Case "OO.ASSETTAG"
            If mintOffsetIx = 1 Then
                mstrReturn = cobjOrder.PrimaryOrderLines.Item(mintLineIx).OrderMachines.Item(mintMachineIx).Pickup1AssetTag
            ElseIf mintOffsetIx = 2 Then
                mstrReturn = cobjOrder.PrimaryOrderLines.Item(mintLineIx).OrderMachines.Item(mintMachineIx).Pickup2AssetTag
            ElseIf mintOffsetIx = 3 Then
                mstrReturn = cobjOrder.PrimaryOrderLines.Item(mintLineIx).OrderMachines.Item(mintMachineIx).Pickup3AssetTag
            End If
            
        'Location-specific, Machine Pickup
        Case "OU.MODEL1"
            If cobjOrder.PrimaryOrderLines.Item(mintLineIx).OrderMachines.Item(mintMachineIx).Pickup1Type = "P" Then
                mstrReturn = cobjOrder.PrimaryOrderLines.Item(mintLineIx).OrderMachines.Item(mintMachineIx).Pickup1Model
            Else
                mstrReturn = ""
            End If
        Case "OU.MODEL2"
            If cobjOrder.PrimaryOrderLines.Item(mintLineIx).OrderMachines.Item(mintMachineIx).Pickup2Type = "P" Then
                mstrReturn = cobjOrder.PrimaryOrderLines.Item(mintLineIx).OrderMachines.Item(mintMachineIx).Pickup2Model
            Else
                mstrReturn = ""
            End If
        Case "OU.MODEL3"
            If cobjOrder.PrimaryOrderLines.Item(mintLineIx).OrderMachines.Item(mintMachineIx).Pickup3Type = "P" Then
                mstrReturn = cobjOrder.PrimaryOrderLines.Item(mintLineIx).OrderMachines.Item(mintMachineIx).Pickup3Model
            Else
                mstrReturn = ""
            End If
        Case "OU.SERIAL1"
            If cobjOrder.PrimaryOrderLines.Item(mintLineIx).OrderMachines.Item(mintMachineIx).Pickup1Type = "P" Then
                mstrReturn = cobjOrder.PrimaryOrderLines.Item(mintLineIx).OrderMachines.Item(mintMachineIx).Pickup1Serial
            Else
                mstrReturn = ""
            End If
        Case "OU.SERIAL2"
            If cobjOrder.PrimaryOrderLines.Item(mintLineIx).OrderMachines.Item(mintMachineIx).Pickup2Type = "P" Then
                mstrReturn = cobjOrder.PrimaryOrderLines.Item(mintLineIx).OrderMachines.Item(mintMachineIx).Pickup2Serial
            Else
                mstrReturn = ""
            End If
        Case "OU.SERIAL3"
            If cobjOrder.PrimaryOrderLines.Item(mintLineIx).OrderMachines.Item(mintMachineIx).Pickup3Type = "P" Then
                mstrReturn = cobjOrder.PrimaryOrderLines.Item(mintLineIx).OrderMachines.Item(mintMachineIx).Pickup3Serial
            Else
                mstrReturn = ""
            End If
        Case "OU.ASSETTAG1"
            If cobjOrder.PrimaryOrderLines.Item(mintLineIx).OrderMachines.Item(mintMachineIx).Pickup1Type = "P" Then
                mstrReturn = cobjOrder.PrimaryOrderLines.Item(mintLineIx).OrderMachines.Item(mintMachineIx).Pickup1AssetTag
            Else
                mstrReturn = ""
            End If
        Case "OU.ASSETTAG2"
            If cobjOrder.PrimaryOrderLines.Item(mintLineIx).OrderMachines.Item(mintMachineIx).Pickup2Type = "P" Then
                mstrReturn = cobjOrder.PrimaryOrderLines.Item(mintLineIx).OrderMachines.Item(mintMachineIx).Pickup2AssetTag
            Else
                mstrReturn = ""
            End If
        Case "OU.ASSETTAG3"
            If cobjOrder.PrimaryOrderLines.Item(mintLineIx).OrderMachines.Item(mintMachineIx).Pickup3Type = "P" Then
                mstrReturn = cobjOrder.PrimaryOrderLines.Item(mintLineIx).OrderMachines.Item(mintMachineIx).Pickup3AssetTag
            Else
                mstrReturn = ""
            End If
            
        'Location-specific, Machine Move
        Case "OV.MODEL1"
            If cobjOrder.PrimaryOrderLines.Item(mintLineIx).OrderMachines.Item(mintMachineIx).Pickup1Type = "M" Then
                mstrReturn = cobjOrder.PrimaryOrderLines.Item(mintLineIx).OrderMachines.Item(mintMachineIx).Pickup1Model
            Else
                mstrReturn = ""
            End If
        Case "OV.MODEL2"
            If cobjOrder.PrimaryOrderLines.Item(mintLineIx).OrderMachines.Item(mintMachineIx).Pickup2Type = "M" Then
                mstrReturn = cobjOrder.PrimaryOrderLines.Item(mintLineIx).OrderMachines.Item(mintMachineIx).Pickup2Model
            Else
                mstrReturn = ""
            End If
        Case "OV.MODEL3"
            If cobjOrder.PrimaryOrderLines.Item(mintLineIx).OrderMachines.Item(mintMachineIx).Pickup3Type = "M" Then
                mstrReturn = cobjOrder.PrimaryOrderLines.Item(mintLineIx).OrderMachines.Item(mintMachineIx).Pickup3Model
            Else
                mstrReturn = ""
            End If
        Case "OV.SERIAL1"
            If cobjOrder.PrimaryOrderLines.Item(mintLineIx).OrderMachines.Item(mintMachineIx).Pickup1Type = "M" Then
                mstrReturn = cobjOrder.PrimaryOrderLines.Item(mintLineIx).OrderMachines.Item(mintMachineIx).Pickup1Serial
            Else
                mstrReturn = ""
            End If
        Case "OV.SERIAL2"
            If cobjOrder.PrimaryOrderLines.Item(mintLineIx).OrderMachines.Item(mintMachineIx).Pickup2Type = "M" Then
                mstrReturn = cobjOrder.PrimaryOrderLines.Item(mintLineIx).OrderMachines.Item(mintMachineIx).Pickup2Serial
            Else
                mstrReturn = ""
            End If
        Case "OV.SERIAL3"
            If cobjOrder.PrimaryOrderLines.Item(mintLineIx).OrderMachines.Item(mintMachineIx).Pickup3Type = "M" Then
                mstrReturn = cobjOrder.PrimaryOrderLines.Item(mintLineIx).OrderMachines.Item(mintMachineIx).Pickup3Serial
            Else
                mstrReturn = ""
            End If
        Case "OV.ASSETTAG1"
            If cobjOrder.PrimaryOrderLines.Item(mintLineIx).OrderMachines.Item(mintMachineIx).Pickup1Type = "M" Then
                mstrReturn = cobjOrder.PrimaryOrderLines.Item(mintLineIx).OrderMachines.Item(mintMachineIx).Pickup1AssetTag
            Else
                mstrReturn = ""
            End If
        Case "OV.ASSETTAG2"
            If cobjOrder.PrimaryOrderLines.Item(mintLineIx).OrderMachines.Item(mintMachineIx).Pickup2Type = "M" Then
                mstrReturn = cobjOrder.PrimaryOrderLines.Item(mintLineIx).OrderMachines.Item(mintMachineIx).Pickup2AssetTag
            Else
                mstrReturn = ""
            End If
        Case "OV.ASSETTAG3"
            If cobjOrder.PrimaryOrderLines.Item(mintLineIx).OrderMachines.Item(mintMachineIx).Pickup3Type = "M" Then
                mstrReturn = cobjOrder.PrimaryOrderLines.Item(mintLineIx).OrderMachines.Item(mintMachineIx).Pickup3AssetTag
            Else
                mstrReturn = ""
            End If
        Case Else
            mstrReturn = GetOrderSpecificData3(pField, pIndex, pAddendumName, pCountDown, pDJSeqNbr, pcol, pcurOB_BuyAmount, pcurOB_SellAmount, pcurOB_LeasePymt, pLineIx, pMachineIx, pOffsetIx, pInt)
    End Select
        
    GoTo EndRoutine
    
EndRoutine:
    GetOrderSpecificData2 = mstrReturn

End Function

