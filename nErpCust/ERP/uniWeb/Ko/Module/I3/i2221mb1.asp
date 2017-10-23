<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<!--
'======================================================================================================
'*  1. Module Name          : Inventory
'*  2. Function Name        : 
'*  3. Program ID           : i2221mb1.asp
'*  4. Program Name         : 품목별 재고현황조회 
'*  5. Program Desc         : 
'*  6. Modified date(First) : 2002/09/04
'*  7. Modified date(Last)  : 2005/02/17
'*  8. Modifier (First)     : Lee Seung Wook
'*  9. Modifier (Last)      : Lee Seung Wook
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

Dim strItemCd
Dim strBasicUnit

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

strItemCd   = FilterVar(Request("txtItemCd"), "''", "S")

Call SubBizQuery("AL")

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
    
	Call SubMakeSQLStatements("AL",strItemCd)                                  
    
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
						Chr(11) & UniConvNumberDBToCompany(lgObjRs(4), ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0) & _
						Chr(11) & UniConvNumberDBToCompany(lgObjRs(5), ggAmtOfMoney.DecPoint, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit, 0)
            
            Select Case ConvSPChars(lgObjRs(8))
				Case "S"
						lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs(6), ggUnitCost.DecPoint, ggUnitCost.RndPolicy, ggUnitCost.RndUnit, 0)
						lgstrData = lgstrData & Chr(11) & "표준단가" 
				Case "M"
						lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs(7), ggUnitCost.DecPoint, ggUnitCost.RndPolicy, ggUnitCost.RndUnit, 0)
						lgstrData = lgstrData & Chr(11) & "이동평균단가"
            End Select
            
			lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs(9), ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0) & _
									Chr(11) & UniConvNumberDBToCompany(lgObjRs(10), ggAmtOfMoney.DecPoint, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit, 0)
            
            Select Case ConvSPChars(lgObjRs(13))
				Case "S"
						lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs(11), ggUnitCost.DecPoint, ggUnitCost.RndPolicy, ggUnitCost.RndUnit, 0)
						lgstrData = lgstrData & Chr(11) & "표준단가"
				Case "M"
						lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs(12), ggUnitCost.DecPoint, ggUnitCost.RndPolicy, ggUnitCost.RndUnit, 0)
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
Sub SubMakeSQLStatements(pDataType,pCode)

    On Error Resume Next                                                           
    Err.Clear                                                                      
    
	Dim iSelCount

	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	Select Case pDataType
		Case "AL"
			lgStrSQL =	" SELECT A.plant_cd,C.plant_nm,A.tracking_No,B.location,A.tot_stk_qty,A.tot_stk_val,A.std_prc," & _
						" A.moving_avg_prc,A.prc_ctrl_indctr,A.prev_tot_stk_qty,A.prev_tot_stk_val," & _
						" A.prev_std_prc,A.prev_moving_avg_prc,A.prev_prc_ctrl" & _
						" FROM I_MATERIAL_VALUATION A, B_ITEM_BY_PLANT B, B_PLANT C " & _
						" WHERE A.plant_cd	= B.plant_cd "	& _
						" AND A.item_cd		= B.item_cd "	& _
						" AND A.plant_cd	= C.plant_cd "	& _
						" AND A.item_cd		= "	& pCode		& _
						" ORDER BY A.plant_cd "
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
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

