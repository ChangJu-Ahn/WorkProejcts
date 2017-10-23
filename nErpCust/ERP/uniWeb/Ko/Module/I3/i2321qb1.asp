<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<!--
'======================================================================================================
'*  1. Module Name          : Inventory
'*  2. Function Name        : 
'*  3. Program ID           : i2321qb1.asp
'*  4. Program Name         : Tracking별 재고현황조회 
'*  5. Program Desc         : 
'*  6. Modified date(First) : 2005/03/08
'*  7. Modified date(Last)  : 2005/10/28
'*  8. Modifier (First)     : Lee Seung Wook
'*  9. Modifier (Last)      : 
'* 10. Comment              : 
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'=======================================================================================================
-->
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" --> 

<%
Call LoadBasisGlobalInf()
Call LoadInfTB19029B("Q", "I", "NOCOOKIE", "MB")

On Error Resume Next
Err.Clear																	

Call HideStatusWnd

Dim IntRetCD

Dim strPlantCd
Dim strItemCd
Dim strTrackingNo
Dim strFlag
Dim strItemAcct

'---------------------------------------Common-----------------------------------------------------------
lgLngMaxRow       = Request("txtMaxRows")                                        
lgMaxCount        = 100
lgStrPrevKeyIndex = UNICInt(Trim(Request("lgStrPrevKeyIndex")),0)   
lgErrorStatus     = "NO"
lgErrorPos        = ""                                                          
lgOpModeCRUD      = Request("txtMode")                                          
'------ Developer Coding part (Start ) ------------------------------------------------------------------

On Error Resume Next

Call SubOpenDB(lgObjConn)

strPlantCd		= FilterVar(Request("txtPlantCd"),"''","S")
strItemCd		= FilterVar("%" & Trim(Request("txtItemCd")) & "%", "''", "S")
strTrackingNo	= FilterVar(Trim(Request("txtTrackingNo")),"''","S")
strItemAcct		= FilterVar(Trim(Request("txtItemAcct")),"''","S")
strFlag			= Request("txtFlag")

If strItemCd = "'%%'" then 
	Call SubBizQuery("FR")
Else
	Call SubBizQuery("AL") 	
End if

Call SubCloseDB(lgObjConn)     

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery(pType)
	
	Dim iDx
	Dim PvArr
	
	On Error Resume Next                                                            
    Err.Clear
    
	If pType = "FR" Then
		Call SubMakeSQLStatements("FR","",strPlantCd,strTrackingNo,strFlag,strItemAcct)
	Else
		Call SubMakeSQLStatements("AL",strItemCd,strPlantCd,strTrackingNo,strFlag,strItemAcct)
	End If                                 
    
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

        lgstrData = ""
        iDx       = 1
        ReDim PvArr(0)
        
        Do While Not lgObjRs.EOF
			ReDim Preserve PvArr(iDx-1)
        
            lgstrData = Chr(11) & ConvSPChars(lgObjRs(0)) & _
						Chr(11) & ConvSPChars(lgObjRs(1)) & _
						Chr(11) & ConvSPChars(lgObjRs(2)) & _
						Chr(11) & ConvSPChars(lgObjRs(3)) & _
						Chr(11) & ConvSPChars(lgObjRs(4)) & _
						Chr(11) & ConvSPChars(lgObjRs(5)) & _
						Chr(11) & UniConvNumberDBToCompany(lgObjRs(6), ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0) & _
						Chr(11) & UniConvNumberDBToCompany(lgObjRs(7), ggAmtOfMoney.DecPoint, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit, 0)
            
            Select Case ConvSPChars(lgObjRs(10))
				Case "S"
						lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs(8), ggUnitCost.DecPoint, ggUnitCost.RndPolicy, ggUnitCost.RndUnit, 0)
						lgstrData = lgstrData & Chr(11) & "표준단가" 
				Case "M"
						lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs(9), ggUnitCost.DecPoint, ggUnitCost.RndPolicy, ggUnitCost.RndUnit, 0)
						lgstrData = lgstrData & Chr(11) & "이동평균단가"
            End Select
            
			lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs(11), ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0) & _
									Chr(11) & UniConvNumberDBToCompany(lgObjRs(12), ggAmtOfMoney.DecPoint, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit, 0)
            
            Select Case ConvSPChars(lgObjRs(15))
				Case "S"
						lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs(13), ggUnitCost.DecPoint, ggUnitCost.RndPolicy, ggUnitCost.RndUnit, 0)
						lgstrData = lgstrData & Chr(11) & "표준단가"
				Case "M"
						lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs(14), ggUnitCost.DecPoint, ggUnitCost.RndPolicy, ggUnitCost.RndUnit, 0)
						lgstrData = lgstrData & Chr(11) & "이동평균단가"
            End Select
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
            lgstrData = lgstrData & Chr(11) & lgLngMaxRow + iDx & Chr(11) & Chr(12)

			PvArr(iDx-1) = lgstrData
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
Sub SubMakeSQLStatements(pDataType,pCode,pCode1,pCode2,pCode3,pCode4)

    On Error Resume Next                                                           
    Err.Clear                                                                      
    
	Dim iSelCount
	
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	lgStrSQL	= " SELECT	DISTINCT A.ITEM_CD,B.ITEM_NM, " _
				& " (CASE WHEN D.S_O_FLAG = " & FilterVar("MTO","''","S") _
				& " AND D.TRACKING_FLAG = " & FilterVar("Y","''","S") _
				& " THEN A.TRACKING_NO ELSE " & FilterVar("*","''","S") & "END) AS TRACKING_NO, " _
				& " B.SPEC,B.BASIC_UNIT," _
				& " C.LOCATION,A.TOT_STK_QTY,A.TOT_STK_VAL,A.STD_PRC,A.MOVING_AVG_PRC, " _
				& " A.PRC_CTRL_INDCTR,A.PREV_TOT_STK_QTY,A.PREV_TOT_STK_VAL,A.PREV_STD_PRC, " _
				& " A.PREV_MOVING_AVG_PRC,A.PREV_PRC_CTRL " _
				& " FROM I_MATERIAL_VALUATION A(nolock) inner join B_ITEM B(nolock) on A.ITEM_CD = B.ITEM_CD " _
				& " LEFT OUTER JOIN B_ITEM_BY_PLANT C(nolock) on A.PLANT_CD = C.PLANT_CD AND A.ITEM_CD = C.ITEM_CD " _
				& " LEFT OUTER JOIN B_ITEM_ACCT_TRACKING D(NOLOCK) ON C.ITEM_ACCT = D.ITEM_ACCT AND C.PLANT_CD = D.PLANT_CD " _
				& " WHERE A.PLANT_CD = " & pCode1 _
				& " AND (CASE WHEN D.S_O_FLAG = " & FilterVar("MTO","''","S") _
				& " AND D.TRACKING_FLAG = " & FilterVar("Y","''","S") _
				& " THEN A.TRACKING_NO ELSE " & FilterVar("*","''","S") & "END) = " _
				& " (CASE WHEN D.S_O_FLAG = " & FilterVar("MTO","''","S") _
				& " AND D.TRACKING_FLAG = " & FilterVar("Y","''","S") _
				& " THEN " & pCode2 & " ELSE " & FilterVar("*","''","S") & " END) " _
				& " AND C.ITEM_ACCT = " & pCode4
   
	If pCode3 = "Y" Then
		lgStrSQL = lgStrSQL & " AND (A.TOT_STK_QTY <> " & 0 _
							& " Or A.PREV_TOT_STK_QTY <> " & 0 & ")"
	End If

	Select Case pDataType

	  Case "FR"
		lgStrSQL	= lgStrSQL	& " ORDER BY A.ITEM_CD " 

	  Case "AL"
		lgStrSQL	= lgStrSQL	& " AND A.ITEM_CD LIKE " & pCode _
								& " ORDER BY A.ITEM_CD "
	End Select
End Sub  

'============================================================================================================
' Name : CommonOnTransactionCommit
' Desc : This Sub is called by OnTransactionCommit Error handler
'============================================================================================================
Sub CommonOnTransactionCommit()
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub

'============================================================================================================
' Name : CommonOnTransactionAbort
' Desc : This Sub is called by OnTransactionAbort Error handler
'============================================================================================================
Sub CommonOnTransactionAbort()
    lgErrorStatus    = "YES"
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub

'============================================================================================================
' Name : SetErrorStatus
' Desc : This Sub set error status
'============================================================================================================
Sub SetErrorStatus()
    lgErrorStatus     = "YES"                                                         
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
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

%>

<Script Language="VBScript">
	With parent
		
		If "<%=lgErrorStatus%>" = "NO" And "<%=IntRetCd%>" <> -1 Then
		    .ggoSpread.Source = .frm1.vspdData1
			.lgStrPrevKeyIndex = "<%=lgStrPrevKeyIndex%>"
			
			.ggoSpread.SSShowData "<%=ConvSPChars(lgstrData)%>"

        	If .frm1.vspdData1.MaxRows < .VisibleRowCnt(.frm1.vspdData1,0)  And .lgStrPrevKeyIndex <> "" Then	
				.DbQuery				
			Else
				.DbQueryOk				
			End If
			.frm1.vspdData1.focus
		End If   

	End With	
       
</Script>	  

