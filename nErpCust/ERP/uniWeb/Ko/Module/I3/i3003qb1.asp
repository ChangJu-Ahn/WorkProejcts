<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<!--'**********************************************************************************************
'*  1. Module Name            : Inventory
'*  2. Function Name          : 
'*  3. Program ID             : i3003qb1.asp
'*  4. Program Name           : 
'*  5. Program Desc           : 담당자별 재고현황조회 
'*  6. Comproxy List          :      
'*  7. Modified date(First)   : 2003/07/03
'*  8. Modified date(Last)    : 2003/07/03
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
Dim strInvMgr,strPlantCd,strSlCd,strItemCd
Dim ComboRow, ComboName


lgLngMaxRow       = Request("txtMaxRows")
lgMaxCount        = 100                  
lgStrPrevKeyIndex = UNICInt(Trim(Request("lgStrPrevKeyIndex")),0)   
lgErrorStatus     = "NO"
lgErrorPos        = ""                   
lgOpModeCRUD      = Request("txtMode") 

strInvMgr		= FilterVar(Request("cboInvMgr"), "''", "S")
strPlantCd		= FilterVar("%" & Trim(Request("txtPlantCd")) & "%", "''", "S")
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
						Chr(11) & ConvSPChars(lgObjRs(7)) & _
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
						Chr(11) & lgLngMaxRow + iDx & Chr(11) & Chr(12) 
			
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
    
    lgStrSQL	=	" SELECT 	A.PLANT_CD,E.PLANT_NM,A.SL_CD,D.SL_NM,A.ITEM_CD,B.ITEM_NM,B.SPEC,B.BASIC_UNIT,A.TRACKING_NO," & _
					"			A.GOOD_ON_HAND_QTY,A.BAD_ON_HAND_QTY,A.STK_ON_INSP_QTY,A.STK_IN_TRNS_QTY, " & _
					"			A.PREV_GOOD_QTY,A.PREV_BAD_QTY,A.PREV_STK_ON_INSP_QTY,A.PREV_STK_IN_TRNS_QTY, " & _
					"			A.SCHD_RCPT_QTY,A.SCHD_ISSUE_QTY,F.PICKING_QTY " & _
					" FROM		I_ONHAND_STOCK A inner join B_ITEM B on A.item_cd = B.item_cd" & _
					"		 	inner join B_ITEM_BY_PLANT C on A.plant_cd = C.plant_cd and A.item_cd = C.item_cd" & _
					"			inner join B_STORAGE_LOCATION D on A.sl_cd = D.sl_cd" & _
					"			inner join B_PLANT E on A.plant_cd = E.plant_cd" & _
					"			inner join (SELECT SUM(PICKING_QTY) AS PICKING_QTY,PLANT_CD,SL_CD,ITEM_CD,TRACKING_NO FROM I_ONHAND_STOCK_DETAIL GROUP BY PLANT_CD,SL_CD,ITEM_CD,TRACKING_NO) F" & _ 
					"			on A.item_cd = F.item_cd and A.plant_cd = F.plant_cd and A.sl_cd = F.sl_cd and A.tracking_no = F.tracking_no" & _
					" WHERE		C.INV_MGR =  "				& strInvMgr		& _
					" AND		A.PLANT_CD LIKE  "			& strPlantCd	& _
					" AND		A.SL_CD LIKE "				& strSlCd		& _
					" AND		A.ITEM_CD LIKE "			& strItemCd		& _
					" ORDER BY	A.PLANT_CD ASC,A.SL_CD ASC,A.ITEM_CD ASC,A.TRACKING_NO ASC "	  
					 
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
 

