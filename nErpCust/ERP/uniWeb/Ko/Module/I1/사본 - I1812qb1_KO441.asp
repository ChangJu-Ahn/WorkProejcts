<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<!-- #Include file="../../ComAsp/LoadinfTB19029.asp" -->
<%
'**********************************************************************************************
'*  1. Module Name          : Sales Management
'*  2. Function Name        : 
'*  3. Program ID           : i1812qb1_KO441
'*  4. Program Name         : 창고별 수불현황 조회 
'*  5. Program Desc         : 
'*  6. Component List       : 
'*  7. Modified date(First) : 2007/07/31
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Lee Ho Jun
'* 10. Modifier (Last)      : 
'* 11. Comment
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change" 
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************

'On Error Resume Next

Call HideStatusWnd															'☜: 모든 작업 완료후 작업진행중 표시창을 Hide'
'Err.Clear

Call LoadBasisGlobalInf
Call LoadinfTB19029B("Q", "Q", "NOCOOKIE", "QB")

Dim IntRetCD
Dim PvArr
Dim NextKey1
Dim strNextKey1

Const C_SHEETMAXROWS_D = 100000

lgLngMaxRow     = Request("txtMaxRows") 
lgErrorStatus   = "NO"

Call HideStatusWnd 

'On Error Resume Next
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

        iDx = 0
        ReDim PvArr(C_SHEETMAXROWS_D)
        
        Do While Not lgObjRs.EOF
 
            If iDx = C_SHEETMAXROWS_D Then
'               NextKey1 = ConvSPChars(lgObjRs(0))
               Exit Do
            End If   

            lgstrData = Chr(11) & ConvSPChars(lgObjRs("sl_cd")) & _
						Chr(11) & ConvSPChars(lgObjRs("Item_Acct_CD")) & _
						Chr(11) & ConvSPChars(lgObjRs("Item_Acct_Nm")) & _
						Chr(11) & ConvSPChars(lgObjRs("item_cd")) & _
						Chr(11) & ConvSPChars(lgObjRs("Item_Nm")) & _
						Chr(11) & ConvSPChars(lgObjRs("BASIC_UNIT")) & _
						Chr(11) & UniConvNumberDBToCompany(lgObjRs("B_PRICE"), ggQty.DecPoint , ggQty.RndPolicy, ggQty.RndUnit, 0) & _
						Chr(11) & UniConvNumberDBToCompany(lgObjRs("TransStockQty"), ggQty.DecPoint , ggQty.RndPolicy, ggQty.RndUnit, 0) & _
						Chr(11) & UniConvNumberDBToCompany(lgObjRs("TransStockAmt"), ggQty.DecPoint , ggQty.RndPolicy, ggQty.RndUnit, 0) & _
						Chr(11) & UniConvNumberDBToCompany(lgObjRs("InProdQty"), ggQty.DecPoint , ggQty.RndPolicy, ggQty.RndUnit, 0) & _
						Chr(11) & UniConvNumberDBToCompany(lgObjRs("InProdAmt"), ggQty.DecPoint , ggQty.RndPolicy, ggQty.RndUnit, 0) & _
						Chr(11) & UniConvNumberDBToCompany(lgObjRs("InPurQty"), ggQty.DecPoint , ggQty.RndPolicy, ggQty.RndUnit, 0) & _
						Chr(11) & UniConvNumberDBToCompany(lgObjRs("InPurAmt"), ggQty.DecPoint , ggQty.RndPolicy, ggQty.RndUnit, 0) & _
						Chr(11) & UniConvNumberDBToCompany(lgObjRs("InExecQty"), ggQty.DecPoint , ggQty.RndPolicy, ggQty.RndUnit, 0) & _
						Chr(11) & UniConvNumberDBToCompany(lgObjRs("InExecAmt"), ggQty.DecPoint , ggQty.RndPolicy, ggQty.RndUnit, 0) & _
						Chr(11) & UniConvNumberDBToCompany(lgObjRs("InStockQty"), ggQty.DecPoint , ggQty.RndPolicy, ggQty.RndUnit, 0) & _
						Chr(11) & UniConvNumberDBToCompany(lgObjRs("InStockAmt"), ggQty.DecPoint , ggQty.RndPolicy, ggQty.RndUnit, 0) & _
						Chr(11) & UniConvNumberDBToCompany(lgObjRs("InSumQty"), ggQty.DecPoint , ggQty.RndPolicy, ggQty.RndUnit, 0) & _
						Chr(11) & UniConvNumberDBToCompany(lgObjRs("InSumAmt"), ggQty.DecPoint , ggQty.RndPolicy, ggQty.RndUnit, 0) & _
						Chr(11) & UniConvNumberDBToCompany(lgObjRs("OutProdQty"), ggQty.DecPoint , ggQty.RndPolicy, ggQty.RndUnit, 0) & _
						Chr(11) & UniConvNumberDBToCompany(lgObjRs("OutProdAmt"), ggQty.DecPoint , ggQty.RndPolicy, ggQty.RndUnit, 0) & _
						Chr(11) & UniConvNumberDBToCompany(lgObjRs("OutPurQty"), ggQty.DecPoint , ggQty.RndPolicy, ggQty.RndUnit, 0) & _
						Chr(11) & UniConvNumberDBToCompany(lgObjRs("OutPurAmt"), ggQty.DecPoint , ggQty.RndPolicy, ggQty.RndUnit, 0) & _
						Chr(11) & UniConvNumberDBToCompany(lgObjRs("OutExecQty"), ggQty.DecPoint , ggQty.RndPolicy, ggQty.RndUnit, 0) & _
						Chr(11) & UniConvNumberDBToCompany(lgObjRs("OutExecAmt"), ggQty.DecPoint , ggQty.RndPolicy, ggQty.RndUnit, 0) & _
						Chr(11) & UniConvNumberDBToCompany(lgObjRs("OutStockQty"), ggQty.DecPoint , ggQty.RndPolicy, ggQty.RndUnit, 0) & _
						Chr(11) & UniConvNumberDBToCompany(lgObjRs("OutStockAmt"), ggQty.DecPoint , ggQty.RndPolicy, ggQty.RndUnit, 0) & _
						Chr(11) & UniConvNumberDBToCompany(lgObjRs("OutSumQty"), ggQty.DecPoint , ggQty.RndPolicy, ggQty.RndUnit, 0) & _
						Chr(11) & UniConvNumberDBToCompany(lgObjRs("OutSumAmt"), ggQty.DecPoint , ggQty.RndPolicy, ggQty.RndUnit, 0) & _
						Chr(11) & UniConvNumberDBToCompany(lgObjRs("StockQty"), ggQty.DecPoint , ggQty.RndPolicy, ggQty.RndUnit, 0) & _
						Chr(11) & UniConvNumberDBToCompany(lgObjRs("StockAmt"), ggQty.DecPoint , ggQty.RndPolicy, ggQty.RndUnit, 0) & _
						Chr(11) & lgLngMaxRow + iDx & Chr(11) & Chr(12)
			
			PvArr(iDx) = lgstrData
			iDx = iDx + 1
		    lgObjRs.MoveNext
        Loop 
    End If

	lgstrData = Join(PvArr, "")

	Call SubHandleError("MR",lgObjConn,lgObjRs,Err)
    Call SubCloseRs(lgObjRs)                                             
		
End Sub

'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements()

    'On Error Resume Next
    'Err.Clear

'	If Trim(Request("txtItemCd")) <> "" Then
'		lgStrSQL = lgStrSQL & " and a.item_cd = " & FilterVar(Request("txtItemCd"),"","S")
'	End If
'			
'	If Trim(Request("txtprojectCode")) <> "" Then
'		lgStrSQL = lgStrSQL & " and a.tracking_no = " & FilterVar(Request("txtprojectCode"),"","S")
'	End If	

'Dim strfromYY, strfromMm
'Dim strToYY, strToMm
    
lgStrSQL = ""
lgStrSQL = lgStrSQL & vbCrLf & " "
lgStrSQL = lgStrSQL & vbCrLf & " SELECT DISTINCT "
lgStrSQL = lgStrSQL & vbCrLf & "    temp.Item_Nm, "
lgStrSQL = lgStrSQL & vbCrLf & "    temp.SL_CD, "
lgStrSQL = lgStrSQL & vbCrLf & "    B.ITEM_ACCT as Item_Acct_CD, "
lgStrSQL = lgStrSQL & vbCrLf & "    dbo.ufn_GetCodeName('P1001',B.ITEM_ACCT) as Item_Acct_Nm, "
lgStrSQL = lgStrSQL & vbCrLf & "    temp.item_cd, "
lgStrSQL = lgStrSQL & vbCrLf & "    temp.BASIC_UNIT, "
lgStrSQL = lgStrSQL & vbCrLf & "    ISNULL((select top 1 case a.prc_ctrl_indctr when 'M' then a.moving_avg_prc "
lgStrSQL = lgStrSQL & vbCrLf & "    				 when 'S' then a.std_prc "
lgStrSQL = lgStrSQL & vbCrLf & "    	end as aa "
lgStrSQL = lgStrSQL & vbCrLf & "    	from i_monthly_inventory a "
lgStrSQL = lgStrSQL & vbCrLf & "    	where a.plant_cd like " & FilterVar(Request("txtPlantCd") & "%","","S")
lgStrSQL = lgStrSQL & vbCrLf & "    	and item_cd = temp.item_cd "
lgStrSQL = lgStrSQL & vbCrLf & "    	and mnth_inv_year =  " & FilterVar(Request("txtToYY"),"","S")
lgStrSQL = lgStrSQL & vbCrLf & "    	and mnth_inv_month = " & FilterVar(Request("txtToMm"),"","S") & " ),0) as b_price, "
lgStrSQL = lgStrSQL & vbCrLf & "    SUM(temp.Inv_Qty) TransStockQty, "
lgStrSQL = lgStrSQL & vbCrLf & "    SUM(temp.Inv_Amt) TransStockAmt, "
lgStrSQL = lgStrSQL & vbCrLf & "    SUM(temp.MR_QTY) InProdQty, "
lgStrSQL = lgStrSQL & vbCrLf & "    SUM(temp.MR_AMT) InProdAmt, "
lgStrSQL = lgStrSQL & vbCrLf & "    SUM(temp.PR_QTY) InPurQty, "
lgStrSQL = lgStrSQL & vbCrLf & "    SUM(temp.PR_AMT) InPurAmt, "
lgStrSQL = lgStrSQL & vbCrLf & "    SUM(temp.OR_QTY) InExecQty, "
lgStrSQL = lgStrSQL & vbCrLf & "    SUM(temp.OR_AMT) InExecAmt, "
lgStrSQL = lgStrSQL & vbCrLf & "    SUM(temp.ST_DEB_QTY) InStockQty, "
lgStrSQL = lgStrSQL & vbCrLf & "    SUM(temp.ST_DEB_AMT) InStockAmt, "
lgStrSQL = lgStrSQL & vbCrLf & "    SUM(temp.MR_QTY+temp.PR_QTY+temp.OR_QTY+temp.ST_DEB_QTY) InSumQty, "
lgStrSQL = lgStrSQL & vbCrLf & "    SUM(temp.MR_AMT+temp.PR_AMT+temp.OR_AMT+temp.ST_DEB_AMT) InSumAmt, "
lgStrSQL = lgStrSQL & vbCrLf & "    SUM(temp.PI_QTY) OutProdQty, "
lgStrSQL = lgStrSQL & vbCrLf & "    SUM(temp.PI_AMT) OutProdAmt, "
lgStrSQL = lgStrSQL & vbCrLf & "    SUM(temp.DI_QTY) OutPurQty, "
lgStrSQL = lgStrSQL & vbCrLf & "    SUM(temp.DI_Amt) OutPurAmt, "
lgStrSQL = lgStrSQL & vbCrLf & "    SUM(temp.OI_QTY) OutExecQty, "
lgStrSQL = lgStrSQL & vbCrLf & "    SUM(temp.OI_AMT) OutExecAmt, "
lgStrSQL = lgStrSQL & vbCrLf & "    SUM(temp.ST_CRE_QTY) OutStockQty, "
lgStrSQL = lgStrSQL & vbCrLf & "    SUM(temp.ST_CRE_AMT) OutStockAmt, "
lgStrSQL = lgStrSQL & vbCrLf & "    SUM(temp.PI_QTY+temp.DI_QTY+temp.OI_QTY+temp.ST_CRE_QTY) OutSumQty, "
lgStrSQL = lgStrSQL & vbCrLf & "    SUM(temp.PI_AMT+temp.DI_AMT+temp.OI_AMT+temp.ST_CRE_AMT) OutSumAmt, "
lgStrSQL = lgStrSQL & vbCrLf & "    SUM(temp.TOTAL_QTY) StockQty,"
lgStrSQL = lgStrSQL & vbCrLf & "    SUM(temp.TOTAL_AMT) StockAmt"
lgStrSQL = lgStrSQL & vbCrLf & " FROM ("
lgStrSQL = lgStrSQL & vbCrLf & "    SELECT  S.SL_CD,"
lgStrSQL = lgStrSQL & vbCrLf & "        S.ITEM_CD, "
lgStrSQL = lgStrSQL & vbCrLf & "        T.ITEM_NM, "
lgStrSQL = lgStrSQL & vbCrLf & "        T.BASIC_UNIT, "
lgStrSQL = lgStrSQL & vbCrLf & "        ISNULL(INV_QTY,0) AS INV_QTY,   "
'lgStrSQL = lgStrSQL & vbCrLf & "        (CASE WHEN S.SL_CD IN ('SL60','SLA4') THEN 0 ELSE ISNULL(INV_AMT,0) END) INV_AMT, "
'lgStrSQL = lgStrSQL & vbCrLf & "        ISNULL(MR_QTY,0) AS MR_QTY, "
'lgStrSQL = lgStrSQL & vbCrLf & "        (CASE WHEN S.SL_CD IN ('SL60','SLA4') THEN 0 ELSE ISNULL(MR_AMT,0) END) AS MR_AMT, "
'lgStrSQL = lgStrSQL & vbCrLf & "        ISNULL(PR_QTY,0) AS PR_QTY, "
'lgStrSQL = lgStrSQL & vbCrLf & "        (CASE WHEN S.SL_CD IN ('SL60','SLA4') THEN 0 ELSE ISNULL(PR_AMT,0) END) AS PR_AMT, "
'lgStrSQL = lgStrSQL & vbCrLf & "        ISNULL(OR_QTY,0) AS OR_QTY, "
'lgStrSQL = lgStrSQL & vbCrLf & "        (CASE WHEN S.SL_CD IN ('SL60','SLA4') THEN 0 ELSE ISNULL(OR_AMT,0) END) AS OR_AMT, "
'lgStrSQL = lgStrSQL & vbCrLf & "        ISNULL(ST_DEB_QTY,0) AS ST_DEB_QTY, "
'lgStrSQL = lgStrSQL & vbCrLf & "        (CASE WHEN S.SL_CD IN ('SL60','SLA4') THEN 0 ELSE ISNULL(ST_DEB_AMT,0) END) AS ST_DEB_AMT,"
'lgStrSQL = lgStrSQL & vbCrLf & "        ISNULL(PI_QTY,0) AS PI_QTY, "
'lgStrSQL = lgStrSQL & vbCrLf & "        (CASE WHEN S.SL_CD IN ('SL60','SLA4') THEN 0 ELSE ISNULL(PI_AMT,0) END) AS PI_AMT, "
'lgStrSQL = lgStrSQL & vbCrLf & "        ISNULL(DI_QTY,0) AS DI_QTY, "
'lgStrSQL = lgStrSQL & vbCrLf & "        (CASE WHEN S.SL_CD IN ('SL60','SLA4') THEN 0 ELSE ISNULL(DI_AMT,0) END) AS DI_AMT, "
'lgStrSQL = lgStrSQL & vbCrLf & "        ISNULL(OI_QTY,0) AS OI_QTY, "
'lgStrSQL = lgStrSQL & vbCrLf & "        (CASE WHEN S.SL_CD IN ('SL60','SLA4') THEN 0 ELSE ISNULL(OI_AMT,0) END) AS OI_AMT, "
'lgStrSQL = lgStrSQL & vbCrLf & "        ISNULL(ST_CRE_QTY,0) AS ST_CRE_QTY, "
'lgStrSQL = lgStrSQL & vbCrLf & "        (CASE WHEN S.SL_CD IN ('SL60','SLA4') THEN 0 ELSE ISNULL(ST_CRE_AMT,0) END) AS ST_CRE_AMT,"
'lgStrSQL = lgStrSQL & vbCrLf & "        ( (ISNULL(INV_QTY,0)+MR_QTY+PR_QTY+OR_QTY+ST_DEB_QTY)-(PI_QTY+DI_QTY+OI_QTY+ST_CRE_QTY )) AS TOTAL_QTY,"
'lgStrSQL = lgStrSQL & vbCrLf & "        (CASE WHEN S.SL_CD IN ('SL60','SLA4') THEN 0 ELSE ( (ISNULL(INV_AMT,0)+MR_AMT+PR_AMT+OR_AMT+ST_DEB_AMT)-(PI_AMT+DI_AMT+OI_AMT+ST_CRE_AMT) ) END) AS TOTAL_AMT"
lgStrSQL = lgStrSQL & vbCrLf & "        ISNULL(INV_AMT,0) INV_AMT, "
lgStrSQL = lgStrSQL & vbCrLf & "        ISNULL(MR_QTY,0) AS MR_QTY, "
lgStrSQL = lgStrSQL & vbCrLf & "        ISNULL(MR_AMT,0) AS MR_AMT, "
lgStrSQL = lgStrSQL & vbCrLf & "        ISNULL(PR_QTY,0) AS PR_QTY, "
lgStrSQL = lgStrSQL & vbCrLf & "        ISNULL(PR_AMT,0) AS PR_AMT, "
lgStrSQL = lgStrSQL & vbCrLf & "        ISNULL(OR_QTY,0) AS OR_QTY, "
lgStrSQL = lgStrSQL & vbCrLf & "        ISNULL(OR_AMT,0) AS OR_AMT, "
lgStrSQL = lgStrSQL & vbCrLf & "        ISNULL(ST_DEB_QTY,0) AS ST_DEB_QTY, "
lgStrSQL = lgStrSQL & vbCrLf & "        ISNULL(ST_DEB_AMT,0) AS ST_DEB_AMT,"
lgStrSQL = lgStrSQL & vbCrLf & "        ISNULL(PI_QTY,0) AS PI_QTY, "
lgStrSQL = lgStrSQL & vbCrLf & "        ISNULL(PI_AMT,0) AS PI_AMT, "
lgStrSQL = lgStrSQL & vbCrLf & "        ISNULL(DI_QTY,0) AS DI_QTY, "
lgStrSQL = lgStrSQL & vbCrLf & "        ISNULL(DI_AMT,0) AS DI_AMT, "
lgStrSQL = lgStrSQL & vbCrLf & "        ISNULL(OI_QTY,0) AS OI_QTY, "
lgStrSQL = lgStrSQL & vbCrLf & "        ISNULL(OI_AMT,0) AS OI_AMT, "
lgStrSQL = lgStrSQL & vbCrLf & "        ISNULL(ST_CRE_QTY,0) AS ST_CRE_QTY, "
lgStrSQL = lgStrSQL & vbCrLf & "        ISNULL(ST_CRE_AMT,0) AS ST_CRE_AMT,"
lgStrSQL = lgStrSQL & vbCrLf & "        ((ISNULL(INV_QTY,0)+MR_QTY+PR_QTY+OR_QTY+ST_DEB_QTY)-(PI_QTY+DI_QTY+OI_QTY+ST_CRE_QTY )) AS TOTAL_QTY,"
lgStrSQL = lgStrSQL & vbCrLf & "        (ISNULL(INV_AMT,0)+MR_AMT+PR_AMT+OR_AMT+ST_DEB_AMT)-(PI_AMT+DI_AMT+OI_AMT+ST_CRE_AMT) AS TOTAL_AMT"
lgStrSQL = lgStrSQL & vbCrLf & "    FROM    ("
lgStrSQL = lgStrSQL & vbCrLf & "            ("
lgStrSQL = lgStrSQL & vbCrLf & " "
lgStrSQL = lgStrSQL & vbCrLf & "             SELECT     A.SL_CD,A.ITEM_CD,"
lgStrSQL = lgStrSQL & vbCrLf & "                       SUM(CASE WHEN A.TRNS_TYPE = 'MR' AND A.DEBIT_CREDIT_FLAG= 'D' THEN A.QTY"
lgStrSQL = lgStrSQL & vbCrLf & "                                         WHEN A.TRNS_TYPE = 'MR' AND A.DEBIT_CREDIT_FLAG= 'C' THEN A.QTY*(-1) ELSE 0 END) AS MR_QTY,"
lgStrSQL = lgStrSQL & vbCrLf & "                       SUM(CASE WHEN A.TRNS_TYPE = 'MR' AND A.DEBIT_CREDIT_FLAG= 'D' THEN A.AMOUNT"
lgStrSQL = lgStrSQL & vbCrLf & "                                         WHEN A.TRNS_TYPE = 'MR' AND A.DEBIT_CREDIT_FLAG= 'C' THEN A.AMOUNT*(-1) ELSE 0 END) AS MR_AMT,"
lgStrSQL = lgStrSQL & vbCrLf & "                       SUM(CASE WHEN A.TRNS_TYPE = 'PR' AND A.DEBIT_CREDIT_FLAG= 'D' THEN A.QTY"
lgStrSQL = lgStrSQL & vbCrLf & "                                         WHEN A.TRNS_TYPE = 'PR' AND A.DEBIT_CREDIT_FLAG= 'C' THEN A.QTY*(-1) ELSE 0 END) AS PR_QTY,"
lgStrSQL = lgStrSQL & vbCrLf & "                       SUM(CASE WHEN A.TRNS_TYPE = 'PR' AND A.DEBIT_CREDIT_FLAG= 'D' THEN A.AMOUNT"
lgStrSQL = lgStrSQL & vbCrLf & "                                         WHEN A.TRNS_TYPE = 'PR' AND A.DEBIT_CREDIT_FLAG= 'C' THEN A.AMOUNT*(-1) ELSE 0 END) AS PR_AMT,"
lgStrSQL = lgStrSQL & vbCrLf & "                       SUM(CASE WHEN A.TRNS_TYPE = 'OR' AND A.DEBIT_CREDIT_FLAG= 'D' THEN A.QTY"
lgStrSQL = lgStrSQL & vbCrLf & "                                         WHEN A.TRNS_TYPE = 'OR' AND A.DEBIT_CREDIT_FLAG= 'C' THEN A.QTY*(-1) ELSE 0 END) AS OR_QTY,"
lgStrSQL = lgStrSQL & vbCrLf & "                       SUM(CASE WHEN A.TRNS_TYPE = 'OR' AND A.DEBIT_CREDIT_FLAG= 'D' THEN A.AMOUNT"
lgStrSQL = lgStrSQL & vbCrLf & "                                         WHEN A.TRNS_TYPE = 'OR' AND A.DEBIT_CREDIT_FLAG= 'C' THEN A.AMOUNT*(-1) ELSE 0 END) AS OR_AMT,"
lgStrSQL = lgStrSQL & vbCrLf & "                       SUM(CASE WHEN A.TRNS_TYPE = 'ST' AND A.DEBIT_CREDIT_FLAG= 'D' THEN A.QTY ELSE 0 END) AS ST_DEB_QTY,"
lgStrSQL = lgStrSQL & vbCrLf & "                       SUM(CASE WHEN A.TRNS_TYPE = 'ST' AND A.DEBIT_CREDIT_FLAG= 'D' THEN A.AMOUNT ELSE 0 END ) AS ST_DEB_AMT,"
lgStrSQL = lgStrSQL & vbCrLf & "                       SUM(CASE WHEN A.TRNS_TYPE = 'PI' AND A.DEBIT_CREDIT_FLAG= 'C' THEN A.QTY "
lgStrSQL = lgStrSQL & vbCrLf & "                                         WHEN A.TRNS_TYPE = 'PI' AND A.DEBIT_CREDIT_FLAG= 'D' THEN A.QTY*(-1) ELSE 0 END ) AS PI_QTY,"
lgStrSQL = lgStrSQL & vbCrLf & "                       SUM(CASE WHEN A.TRNS_TYPE = 'PI' AND A.DEBIT_CREDIT_FLAG= 'C' THEN A.AMOUNT"
lgStrSQL = lgStrSQL & vbCrLf & "                                         WHEN A.TRNS_TYPE = 'PI' AND A.DEBIT_CREDIT_FLAG= 'D' THEN A.AMOUNT*(-1) ELSE 0 END ) AS PI_AMT,"
lgStrSQL = lgStrSQL & vbCrLf & "                       SUM(CASE WHEN A.TRNS_TYPE = 'DI' AND A.DEBIT_CREDIT_FLAG= 'C' THEN A.QTY "
lgStrSQL = lgStrSQL & vbCrLf & "                                         WHEN A.TRNS_TYPE = 'DI' AND A.DEBIT_CREDIT_FLAG= 'D' THEN A.QTY*(-1) ELSE 0 END ) AS DI_QTY,"
lgStrSQL = lgStrSQL & vbCrLf & "                       SUM(CASE WHEN A.TRNS_TYPE = 'DI' AND A.DEBIT_CREDIT_FLAG= 'C' THEN A.AMOUNT"
lgStrSQL = lgStrSQL & vbCrLf & "                                         WHEN A.TRNS_TYPE = 'DI' AND A.DEBIT_CREDIT_FLAG= 'D' THEN A.AMOUNT*(-1) ELSE 0 END ) AS DI_AMT,"
lgStrSQL = lgStrSQL & vbCrLf & "                       SUM(CASE WHEN A.TRNS_TYPE = 'OI' AND A.DEBIT_CREDIT_FLAG= 'C' THEN A.QTY "
lgStrSQL = lgStrSQL & vbCrLf & "                                         WHEN A.TRNS_TYPE = 'OI' AND A.DEBIT_CREDIT_FLAG= 'D' THEN A.QTY*(-1) ELSE 0 END ) AS OI_QTY,"
lgStrSQL = lgStrSQL & vbCrLf & "                       SUM(CASE WHEN A.TRNS_TYPE = 'OI' AND A.DEBIT_CREDIT_FLAG= 'C' THEN A.AMOUNT"
lgStrSQL = lgStrSQL & vbCrLf & "                                         WHEN A.TRNS_TYPE = 'OI' AND A.DEBIT_CREDIT_FLAG= 'D' THEN A.AMOUNT*(-1) ELSE 0 END ) AS OI_AMT,"
lgStrSQL = lgStrSQL & vbCrLf & "                       SUM(CASE WHEN A.TRNS_TYPE = 'ST' AND A.DEBIT_CREDIT_FLAG= 'C' THEN A.QTY ELSE 0 END) AS ST_CRE_QTY,"
lgStrSQL = lgStrSQL & vbCrLf & "                       SUM(CASE WHEN A.TRNS_TYPE = 'ST' AND A.DEBIT_CREDIT_FLAG= 'C' THEN A.AMOUNT ELSE 0 END ) AS ST_CRE_AMT,0 AS INV_QTY,0 AS INV_AMT"
lgStrSQL = lgStrSQL & vbCrLf & "                FROM I_GOODS_MOVEMENT_DETAIL A, I_GOODS_MOVEMENT_HEADER B, B_ITEM_BY_PLANT C"
lgStrSQL = lgStrSQL & vbCrLf & "                WHERE C.PLANT_CD = A.PLANT_CD"
lgStrSQL = lgStrSQL & vbCrLf & "                    AND C.ITEM_CD = A.ITEM_CD"
lgStrSQL = lgStrSQL & vbCrLf & "                    AND A.DELETE_FLAG = 'N'"
lgStrSQL = lgStrSQL & vbCrLf & "                    AND B.DOCUMENT_YEAR = A.DOCUMENT_YEAR"
lgStrSQL = lgStrSQL & vbCrLf & "                    AND B.ITEM_DOCUMENT_NO = A.ITEM_DOCUMENT_NO"
lgStrSQL = lgStrSQL & vbCrLf & "                    AND A.PLANT_CD LIKE " & FilterVar(Request("txtPlantCd") & "%","","S")
lgStrSQL = lgStrSQL & vbCrLf & "                    AND A.SL_CD LIKE " & FilterVar(Request("txtSlCd") & "%","","S")
lgStrSQL = lgStrSQL & vbCrLf & "                    AND C.ITEM_ACCT LIKE " & FilterVar(Request("txtItemAcct") & "%","","S")
'lgStrSQL = lgStrSQL & vbCrLf & "                    AND C.ITEM_CD >= " & FilterVar(Request("txtItemCd") & "%","","S")"
lgStrSQL = lgStrSQL & vbCrLf & "                    AND C.ITEM_CD LIKE " & FilterVar(Request("txtItemCd") & "%","","S")
'lgStrSQL = lgStrSQL & vbCrLf & "					And  A.Item_Cd Not Like 'X%'"
'lgStrSQL = lgStrSQL & vbCrLf & "					And c.Item_Acct <> '40'"
lgStrSQL = lgStrSQL & vbCrLf & "                    AND A.DOCUMENT_YEAR BETWEEN " & FilterVar(Request("txtFromYY"),"","S") & " AND " & FilterVar(Request("txtToYY"),"","S")
lgStrSQL = lgStrSQL & vbCrLf & "                    AND CONVERT(CHAR(6), B.DOCUMENT_DT, 112) BETWEEN " & FilterVar(Request("txtFromYY"),"","S") & "+" & FilterVar(Request("txtFromMM"),"","S") & " AND " & FilterVar(Request("txtToYY"),"","S") & "+" & FilterVar(Request("txtToMm"),"","S")
'2007/09/10 추가(같은 창고에 같은 품목 이동은 뺌) -----
lgStrSQL = lgStrSQL & vbCrLf & "                    AND NOT (A.TRNS_TYPE = 'ST' AND A.SL_CD = A.TRNS_SL_CD AND A.ITEM_CD = A.TRNS_ITEM_CD) "
' ----- ----- ----- ----- ----- ----- ----- ----- -----
lgStrSQL = lgStrSQL & vbCrLf & "                    AND (a.qty <> 0  or a.amount <> 0) "
lgStrSQL = lgStrSQL & vbCrLf & "                GROUP BY A.SL_CD, A.ITEM_CD)"
lgStrSQL = lgStrSQL & vbCrLf & "        UNION all "
lgStrSQL = lgStrSQL & vbCrLf & "            ("
lgStrSQL = lgStrSQL & vbCrLf & "                SELECT A.SL_CD, A.ITEM_CD,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0, A.GOOD_ON_HAND_QTY, "
'lgStrSQL = lgStrSQL & vbCrLf & "                (CASE WHEN C.PRC_CTRL_INDCTR = 'S' THEN  A.GOOD_ON_HAND_QTY * C.STD_PRC ELSE A.GOOD_ON_HAND_QTY * C.MOVING_AVG_PRC END) AS INV_AMT "
lgStrSQL = lgStrSQL & vbCrLf & "                (CASE WHEN ISNULL(C.INV_QTY,0) = 0 THEN  0 ELSE A.GOOD_ON_HAND_QTY * (C.INV_AMT/C.INV_QTY) END) AS INV_AMT "
lgStrSQL = lgStrSQL & vbCrLf & "                FROM I_ONHAND_STOCK_HISTORY A, B_ITEM_BY_PLANT B, I_MONTHLY_INVENTORY C"
lgStrSQL = lgStrSQL & vbCrLf & "                WHERE A.PLANT_CD = B.PLANT_CD"
lgStrSQL = lgStrSQL & vbCrLf & "                AND A.ITEM_CD = B.ITEM_CD"
lgStrSQL = lgStrSQL & vbCrLf & "                AND A.PLANT_CD = C.PLANT_CD"
lgStrSQL = lgStrSQL & vbCrLf & "                AND A.ITEM_CD = C.ITEM_CD"
lgStrSQL = lgStrSQL & vbCrLf & "                AND A.MNTH_INV_YEAR =   c.MNTH_INV_YEAR"
lgStrSQL = lgStrSQL & vbCrLf & "                AND A.MNTH_INV_MONTH = c.MNTH_INV_MONTH"
lgStrSQL = lgStrSQL & vbCrLf & "                AND A.PLANT_CD LIKE " & FilterVar(Request("txtPlantCd") & "%","","S")
lgStrSQL = lgStrSQL & vbCrLf & "                AND A.SL_CD LIKE " & FilterVar(Request("txtSlCd") & "%","","S")
lgStrSQL = lgStrSQL & vbCrLf & "                AND B.ITEM_ACCT LIKE " & FilterVar(Request("txtItemAcct") & "%","","S")
'lgStrSQL = lgStrSQL & vbCrLf & "                AND A.ITEM_CD >= " & FilterVar(Request("txtItemCd") & "%","","S")"
lgStrSQL = lgStrSQL & vbCrLf & "                AND A.ITEM_CD like " & FilterVar(Request("txtItemCd") & "%","","S")
lgStrSQL = lgStrSQL & vbCrLf & "                AND (A.GOOD_ON_HAND_QTY <> 0 OR C.INV_QTY <> 0  OR C.INV_AMT <> 0) "
'lgStrSQL = lgStrSQL & vbCrLf & "                And A.Item_Cd Not Like 'X%'"
'lgStrSQL = lgStrSQL & vbCrLf & "                And b.Item_Acct <> '40'"
lgStrSQL = lgStrSQL & vbCrLf & "                AND A.MNTH_INV_YEAR = CONVERT(CHAR(4), DATEADD(DAY,-1," & FilterVar(Request("txtFromYY"),"","S") & "+" & FilterVar(Request("txtFromMM"),"","S") & "+'01'), 112)"
lgStrSQL = lgStrSQL & vbCrLf & "                AND A.MNTH_INV_MONTH = CONVERT(CHAR(2), DATEADD(DAY,-1," & FilterVar(Request("txtFromYY"),"","S") & "+" & FilterVar(Request("txtFromMM"),"","S") & "+'01'), 110)"
lgStrSQL = lgStrSQL & vbCrLf & "                      )"
lgStrSQL = lgStrSQL & vbCrLf & "    ) S, B_ITEM T"
lgStrSQL = lgStrSQL & vbCrLf & " WHERE S.ITEM_CD *= T.ITEM_CD"
'lgStrSQL = lgStrSQL & vbCrLf & " AND S.PLANT_CD *= B.PLANT_CD"
'lgStrSQL = lgStrSQL & vbCrLf & " AND S.ITEM_CD *= B.ITEM_CD"
lgStrSQL = lgStrSQL & vbCrLf & " ) temp, B_STORAGE_LOCATION, B_ITEM_BY_PLANT B"
lgStrSQL = lgStrSQL & vbCrLf & " WHERE ( B_STORAGE_LOCATION.SL_CD LIKE " & FilterVar(Request("txtSlCd") & "%","","S") & " ) AND ( B_STORAGE_LOCATION.SL_CD = temp.sl_cd ) AND (temp.ITEM_CD *= B.ITEM_CD) AND B.PLANT_CD like " & FilterVar(Request("txtPlantCd") & "%","","S")
lgStrSQL = lgStrSQL & vbCrLf & "  GROUP BY temp.Item_Nm, temp.SL_CD, temp.item_cd, temp.BASIC_UNIT, B.ITEM_ACCT"
lgStrSQL = lgStrSQL & vbCrLf & "  HAVING SUM(temp.Inv_Amt) <> 0 Or SUM(temp.MR_AMT) <> 0 Or SUM(temp.PR_AMT) <> 0 Or SUM(temp.OR_AMT) <> 0 Or SUM(temp.ST_DEB_AMT) <> 0 Or SUM(temp.PI_AMT) <> 0 Or SUM(temp.DI_AMT) <> 0 Or SUM(temp.OI_AMT) <> 0 Or SUM(temp.ST_CRE_AMT) <> 0 Or SUM(temp.TOTAL_AMT) <> 0 Or SUM(temp.Inv_Qty) <> 0 Or SUM(temp.MR_QTY) <> 0 Or SUM(temp.PR_QTY) <> 0 Or SUM(temp.OR_QTY) <> 0 Or SUM(temp.ST_DEB_QTY) <> 0 Or SUM(temp.PI_QTY) <> 0 Or SUM(temp.DI_QTY) <> 0 Or SUM(temp.OI_QTY) <> 0 Or SUM(temp.ST_CRE_QTY) <> 0 Or SUM(temp.TOTAL_QTY) <> 0"
lgStrSQL = lgStrSQL & vbCrLf & "  ORDER BY temp.sl_cd ASC,temp.item_cd, b.item_acct ASC"

'Response.Write lgStrSQL
'Response.eND

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


Response.Write "<Script language=vbs> " & vbCr         
Response.Write " With Parent "      	& vbCr
Response.Write "	If """ & lgErrorStatus & """ = ""NO"" And """ & IntRetCd & """ <> -1 Then "	& vbCr
Response.Write "    .lgStrPrevKey  = """ & NextKey1 & """" & vbCr  
Response.Write "	.ggoSpread.Source	= .frm1.vspdData "				& vbCr
Response.Write "	.ggoSpread.SSShowDataByClip  """ & lgstrData  & """"        & vbCr
Response.Write "		If .frm1.vspdData.MaxRows < .parent.VisibleRowCnt(.frm1.vspdData, 0) And .lgStrPrevKey <> """" Then "	& vbCr
Response.Write "			.DbQuery						"				& vbCr
Response.Write "		Else								"				& vbCr
Response.Write "			.DbQueryOK						"				& vbCr
Response.Write "		End If								"				& vbCr
Response.Write "		.frm1.vspdData.focus				"				& vbCr
Response.Write "    End If								"				& vbCr
Response.Write " End With "             & vbCr		
Response.Write "</Script> "             & vbCr 
Response.End     

%>    
