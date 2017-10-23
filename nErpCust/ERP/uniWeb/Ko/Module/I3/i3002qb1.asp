<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<!--'**********************************************************************************************
'*  1. Module Name            : Inventory
'*  2. Function Name          : 
'*  3. Program ID             : i3002qb1.asp
'*  4. Program Name           : 
'*  5. Program Desc           : 재고현황 상세조회 
'*  6. Comproxy List          :      
'*  7. Modified date(First)   : 2003/07/02
'*  8. Modified date(Last)    : 2003/07/02
'*  9. Modifier (First)       : Lee Seung Wook
'* 10. Modifier (Last)        : 
'* 11. Comment                :
'**********************************************************************************************-->
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" --> 
<%
Call LoadBasisGlobalInf()
Call LoadInfTB19029B("Q", "I", "NOCOOKIE", "QB")

On Error Resume Next
Err.Clear					

Call HideStatusWnd

Dim IntRetCD
Dim PvArr
Dim strPlantCd,strItemAcct,strSlCd,strItemCd
Dim ComboRow, ComboName


lgLngMaxRow       = Request("txtMaxRows")
lgMaxCount        = 100                  
lgStrPrevKeyIndex = UNICInt(Trim(Request("lgStrPrevKeyIndex")),0)   
lgErrorStatus     = "NO"
lgErrorPos        = ""                   
lgOpModeCRUD      = Request("txtMode") 

strPlantCd		= FilterVar(Request("txtPlantCd"), "''", "S")
strItemAcct		= FilterVar(Request("txtItemAcct"), "''", "S")
strSlCd			= FilterVar("%" & Trim(Request("txtSlCd")) & "%", "''", "S")
strItemCd		= FilterVar("%" & Trim(Request("txtItemCd")) & "%", "''", "S")


On Error Resume Next

Call SubOpenDB(lgObjConn)

Call SubCreateCommandObject(lgObjComm)

Call SubBizQuery()

Call SubCloseCommandObject(lgObjComm)    
Call SubCloseDB(lgObjConn)     

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()
	
	Dim iDx
	
	On Error Resume Next           
    Err.Clear
    
	Call SubMakeSQLStatements      
    
	If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then 
		IntRetCD = -1
		lgStrPrevKeyIndex = ""		
		Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)
		Call SetErrorStatus()

		Call SubCloseRs(lgObjRs)

		Response.End 
	Else
		IntRetCD = 1
		Call SubSkipRs(lgObjRs,lgMaxCount * lgStrPrevKeyIndex)
		
		lgstrData	= ""
		iDx			= 1

        ReDim PvArr(0)
        
        Do While Not lgObjRs.EOF
			
			ReDim Preserve PvArr(iDx-1)
        
            lgstrData = Chr(11) & ConvSPChars(lgObjRs(0)) & _
						Chr(11) & ConvSPChars(lgObjRs(1)) & _
						Chr(11) & ConvSPChars(lgObjRs(2)) & _
						Chr(11) & ConvSPChars(lgObjRs(3)) & _
						Chr(11) & ConvSPChars(lgObjRs(4)) & _
						Chr(11) & ConvSPChars(lgObjRs(5)) & _
						Chr(11) & ConvSPChars(lgObjRs(6)) & _
						Chr(11) & UniConvNumberDBToCompany(lgObjRs(7),ggQty.DecPoint,ggQty.RndPolicy, ggQty.RndUnit, 0) & _
						Chr(11) & ConvSPChars(lgObjRs(8)) & _
						Chr(11) & UniConvNumberDBToCompany(lgObjRs(9),ggQty.DecPoint,ggQty.RndPolicy, ggQty.RndUnit, 0) & _
						Chr(11) & UniConvNumberDBToCompany(lgObjRs(10),ggQty.DecPoint,ggQty.RndPolicy, ggQty.RndUnit, 0) & _
						Chr(11) & UniConvNumberDBToCompany(lgObjRs(11),ggQty.DecPoint,ggQty.RndPolicy, ggQty.RndUnit, 0) & _
						Chr(11) & UniConvNumberDBToCompany(lgObjRs(12),ggQty.DecPoint,ggQty.RndPolicy, ggQty.RndUnit, 0) & _
						Chr(11) & UniConvNumberDBToCompany(lgObjRs(13),ggQty.DecPoint,ggQty.RndPolicy, ggQty.RndUnit, 0) & _
						Chr(11) & UniConvNumberDBToCompany(lgObjRs(14),ggQty.DecPoint,ggQty.RndPolicy, ggQty.RndUnit, 0) & _
						Chr(11) & UniConvNumberDBToCompany(lgObjRs(15),ggQty.DecPoint,ggQty.RndPolicy, ggQty.RndUnit, 0) & _
						Chr(11) & UniConvNumberDBToCompany(lgObjRs(16),ggQty.DecPoint,ggQty.RndPolicy, ggQty.RndUnit, 0) & _
						Chr(11) & UniConvNumberDBToCompany(lgObjRs(17),ggQty.DecPoint,ggQty.RndPolicy, ggQty.RndUnit, 0) & _
						Chr(11) & UniConvNumberDBToCompany(lgObjRs(18),ggQty.DecPoint,ggQty.RndPolicy, ggQty.RndUnit, 0) & _
						Chr(11) & UniConvNumberDBToCompany(lgObjRs(19),ggQty.DecPoint,ggQty.RndPolicy, ggQty.RndUnit, 0) & _
						Chr(11) & UNIDateClientFormat(lgObjRs(20)) & _
						Chr(11) & UNIDateClientFormat(lgObjRs(21)) & _
						Chr(11) & UNIDateClientFormat(lgObjRs(22)) & _
						Chr(11) & UNIDateClientFormat(lgObjRs(23)) 
						
			Select Case ConvSPChars(lgObjRs(24))
				Case "S"
						lgstrData = lgstrData & Chr(11) & "표준단가"
						lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs(25), ggUnitCost.DecPoint, ggUnitCost.RndPolicy, ggUnitCost.RndUnit, 0)
				Case "M"
						lgstrData = lgstrData & Chr(11) & "이동평균단가"
						lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs(26), ggUnitCost.DecPoint, ggUnitCost.RndPolicy, ggUnitCost.RndUnit, 0)
            End Select						
						
			lgstrData = lgstrData &	Chr(11) & lgLngMaxRow + iDx & Chr(11) & Chr(12) 
			
			PvArr(iDx - 1) = lgstrData
		    lgObjRs.MoveNext
		    
		    iDx =  iDx + 1
	        If iDx > lgMaxCount Then
	           lgStrPrevKeyIndex = lgStrPrevKeyIndex + 1
	           Exit Do
	        End If   
		    
        Loop
        
        lgstrData = Join(PvArr, "")

    End If
    
    If iDx <= lgMaxCount Then
	   lgStrPrevKeyIndex = ""
	End If 

	Call SubHandleError("MR",lgObjConn,lgObjRs,Err)
    Call SubCloseRs(lgObjRs)  
    
    lgStrSQL = ""                                           
		
End Sub

'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements()

    On Error Resume Next
    Err.Clear
    
    lgStrSQL	=	" SELECT 	A.SL_CD,D.SL_NM,A.ITEM_CD,B.ITEM_NM,B.SPEC,A.TRACKING_NO,A.LOT_NO,A.LOT_SUB_NO,B.BASIC_UNIT," & _
					"			(A.GOOD_ON_HAND_QTY+A.BAD_ON_HAND_QTY+A.STK_ON_INSP_QTY+A.STK_ON_TRNS_QTY) AS ONHAND_QTY, " & _
					"			A.GOOD_ON_HAND_QTY,A.BAD_ON_HAND_QTY,A.STK_ON_INSP_QTY,A.STK_ON_TRNS_QTY, " & _
					"			(A.PREV_GOOD_QTY+A.PREV_BAD_QTY+A.PREV_STK_ON_INSP_QTY+A.PREV_STK_IN_TRNS_QTY) AS PREV_ONHAND_QTY, " & _
					"			A.PREV_GOOD_QTY,A.PREV_BAD_QTY,A.PREV_STK_ON_INSP_QTY,A.PREV_STK_IN_TRNS_QTY,A.PICKING_QTY," & _
					"			F.LAST_RCPT_DT,A.LAST_ISSUE_DT,A.EXPIARY_DT,A.LAST_PHY_INV_INSP_DT,E.PRC_CTRL_INDCTR,E.STD_PRC,E.MOVING_AVG_PRC " & _
					" FROM		I_ONHAND_STOCK_DETAIL A inner join B_ITEM B on A.item_cd = B.item_cd" & _
					"		 	inner join B_ITEM_BY_PLANT C on A.plant_cd = C.plant_cd and A.item_cd = C.item_cd" & _
					"			inner join B_STORAGE_LOCATION D on A.sl_cd = D.sl_cd" & _
					"			inner join I_MATERIAL_VALUATION E on A.plant_cd = E.plant_cd and A.item_cd = E.item_cd" & _
					"			left outer join (SELECT MAX(B.DOCUMENT_DT) AS LAST_RCPT_DT,A.ITEM_CD,A.PLANT_CD,A.SL_CD,A.TRACKING_NO,A.LOT_NO,A.LOT_SUB_NO" & _
					"			FROM I_GOODS_MOVEMENT_DETAIL A inner join I_GOODS_MOVEMENT_HEADER B " & _
					"			on A.ITEM_DOCUMENT_NO = B.ITEM_DOCUMENT_NO and A.DOCUMENT_YEAR = B.DOCUMENT_YEAR " & _
					"			and  A.DEBIT_CREDIT_FLAG = " & FilterVar("D", "''", "S") & "  and A.DELETE_FLAG = " & FilterVar("N", "''", "S") & " " & _
					"			GROUP BY B.DOCUMENT_DT,A.ITEM_CD,A.PLANT_CD,A.SL_CD,A.TRACKING_NO,A.LOT_NO,A.LOT_SUB_NO) F" & _
					"			on A.item_cd = F.item_cd and A.plant_cd = F.plant_cd and A.sl_cd = F.sl_cd and A.tracking_no = F.tracking_no " & _
					"			and A.lot_no = F.lot_no and A.lot_sub_no = F.lot_sub_no " & _ 
					" WHERE		A.PLANT_CD = "				& strPlantCd	& _
					" AND		C.ITEM_ACCT = "				& strItemAcct	& _
					" AND		A.SL_CD LIKE "				& strSlCd		& _
					" AND		A.ITEM_CD LIKE "			& strItemCd		& _
					" ORDER BY	A.SL_CD ASC,A.ITEM_CD ASC,A.TRACKING_NO ASC,A.LOT_NO ASC,A.LOT_SUB_NO ASC "	  
					 
End Sub  

'============================================================================================================
' Name : CommonOnTransactionAbort
' Desc : This Sub is called by OnTransactionAbort Error handler
'============================================================================================================
Sub CommonOnTransactionAbort()
    lgErrorStatus    = "YES"
End Sub

'============================================================================================================
' Name : SetErrorStatus
' Desc : This Sub set error status
'============================================================================================================
Sub SetErrorStatus()
    lgErrorStatus     = "YES"       
End Sub

'============================================================================================================
' Name : SubHandleError
' Desc : This Sub handle error
'============================================================================================================
Sub SubHandleError(pOpCode,pConn,pRs,pErr)
    On Error Resume Next                 
    Err.Clear                            

    Select Case pOpCode
        Case "MC"
            If CheckSYSTEMError(pErr,True) = True Then
               ObjectContext.SetAbort
               Call SetErrorStatus
            Else
               If CheckSQLError(pConn,True) = True Then
                  ObjectContext.SetAbort
                  Call SetErrorStatus
               End If
            End If
        Case "MD"
        Case "MR"
        Case "MU"
            If CheckSYSTEMError(pErr,True) = True Then
               ObjectContext.SetAbort
               Call SetErrorStatus
            Else
               If CheckSQLError(pConn,True) = True Then
                  ObjectContext.SetAbort
                  Call SetErrorStatus
               End If
            End If
        Case "MB"
			ObjectContext.SetAbort
            Call SetErrorStatus        
    End Select
End Sub


Function SetComboSplit(ByVal InitCombo)
	Dim ComboList
	Dim InitCode, InitName
	Dim iArrR
	
	ComboList = Split(Initcombo,Chr(12))
	InitCode  = Split(ComboList(0),Chr(11))
	InitName  = Split(ComboList(1),Chr(11))
	
	ReDim ComboList(1, Ubound(InitCode) - 1)
	
	For iArrR = 0 To Ubound(InitCode) - 1
		ComboList(0, iArrR) = InitCode(iArrR)
		ComboList(1, iArrR) = InitName(iArrR)
	Next
	SetComboSplit = ComboList
End Function
  

Response.Write "<Script language=vbs> "																								& vbCr         
Response.Write " With Parent "      																								& vbCr
Response.Write "	If """ & lgErrorStatus & """ = ""NO"" And """ & IntRetCd & """ <> -1 Then "										& vbCr
Response.Write "	.ggoSpread.Source	= .frm1.vspdData "																			& vbCr
Response.Write "	.lgStrPrevKeyIndex	=  """ & lgStrPrevKeyIndex & """"															& vbCr	 
Response.Write "	.ggoSpread.SSShowData  """ & lgstrData  & """"																	& vbCr
Response.Write "		If .frm1.vspdData.MaxRows < .parent.VisibleRowCnt(.frm1.vspdData, 0) And .lgStrPrevKeyIndex <> """" Then "	& vbCr
Response.Write "			.DbQuery						"																		& vbCr
Response.Write "		Else								"																		& vbCr
Response.Write "			.DbQueryOK						"																		& vbCr
Response.Write "		End If								"																		& vbCr
Response.Write "		.frm1.vspdData.focus				"																		& vbCr
Response.Write "    End If								"																			& vbCr
Response.Write " End With "             & vbCr		
Response.Write "</Script> "             & vbCr 

%>
 

