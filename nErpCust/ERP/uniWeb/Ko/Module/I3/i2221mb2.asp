<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<!--
'======================================================================================================
'*  1. Module Name          : Inventory
'*  2. Function Name        : 
'*  3. Program ID           : i2221mb2.asp
'*  4. Program Name         : 품목별 재고현황조회 
'*  5. Program Desc         : 
'*  6. Modified date(First) : 2002/09/04
'*  7. Modified date(Last)  : 2002/10/07
'*  8. Modifier (First)     : Ahn Jung Je
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

Dim strItemCd
Dim strPlantCd
Dim lngHeaderRow

'---------------------------------------Common-----------------------------------------------------------
lgLngMaxRow       = Request("txtMaxRows")                                      
lgMaxCount        = 100
lgStrPrevKeyIndex = UNICInt(Trim(Request("lgStrPrevKeyIndex")),0)   
lgErrorStatus     = "NO"
lgErrorPos        = ""                                                         
lgOpModeCRUD      = Request("txtMode")                                          
'------ Developer Coding part (Start ) ------------------------------------------------------------------


strMode = Request("txtMode")												

On Error Resume Next


Call SubOpenDB(lgObjConn)
	strPlantCd  = FilterVar(Request("txtPlantCd"), "''", "S")
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
    
		Call SubMakeSQLStatements("AL",strPlantCd,strItemCd)                                 
    
	If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then                 
		IntRetCD = -1
		lgStrPrevKeyIndex = ""		
		'Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)     
		Call SetErrorStatus()

		Call SubCloseRs(lgObjRs)
		
%>
<Script Language="VBScript">
		parent.frm1.vspdData1.focus
</Script>	
<%
	
		Response.End 
	Else
		IntRetCD = 1
		 Call SubSkipRs(lgObjRs,lgMaxCount * lgStrPrevKeyIndex)

        lgstrData = ""
        iDx       = 1
        Redim PvArr(0)
        
        Do While Not lgObjRs.EOF
			ReDim Preserve PvArr(iDx-1)

            lgstrData = Chr(11) & ConvSPChars(lgObjRs(0)) & _
						Chr(11) & ConvSPChars(lgObjRs(1)) & _
						Chr(11) & ConvSPChars(lgObjRs(2)) & _
						Chr(11) & UniConvNumberDBToCompany(lgObjRs(3), ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0) & _
						Chr(11) & UniConvNumberDBToCompany(lgObjRs(4), ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0) & _
						Chr(11) & UniConvNumberDBToCompany(lgObjRs(5), ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0) & _
						Chr(11) & UniConvNumberDBToCompany(lgObjRs(6), ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0) & _
						Chr(11) & UniConvNumberDBToCompany(lgObjRs(7), ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0) & _
						Chr(11) & UniConvNumberDBToCompany(lgObjRs(8), ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0) & _
						Chr(11) & UniConvNumberDBToCompany(lgObjRs(9), ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0) & _
						Chr(11) & UniConvNumberDBToCompany(lgObjRs(10), ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0) & _
						Chr(11) & UniConvNumberDBToCompany(lgObjRs(11), ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0) & _
						Chr(11) & UniConvNumberDBToCompany(lgObjRs(12), ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0) & _
						Chr(11) & UniConvNumberDBToCompany(lgObjRs(13), ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0) & _
						Chr(11) & UniConvNumberDBToCompany(lgObjRs(14), ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0) & _
						Chr(11) & lgLngMaxRow + iDx & Chr(11) & Chr(12)

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
Sub SubMakeSQLStatements(pDataType,pCode,pCode1)

    On Error Resume Next                                                            
    Err.Clear                                                                       
    
	Dim iSelCount

	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	Select Case pDataType
		Case "AL"
			lgStrSQL =	"select A.sl_cd, B.sl_nm, A.tracking_no, A.good_on_hand_qty, A.bad_on_hand_qty,A.stk_on_insp_qty, A.stk_in_trns_qty,A.schd_rcpt_qty,A.schd_issue_qty,A.prev_good_qty, A.prev_bad_qty, A.prev_stk_on_insp_qty,A.prev_stk_in_trns_qty, A.allocation_qty, D.picking_qty" & _
						" FROM I_ONHAND_STOCK A, B_STORAGE_LOCATION B, B_ITEM C,(SELECT PLANT_CD,SL_CD,ITEM_CD,TRACKING_NO,SUM(PICKING_QTY) PICKING_QTY FROM I_ONHAND_STOCK_DETAIL GROUP BY PLANT_CD,SL_CD,ITEM_CD,TRACKING_NO) D" & _
						" WHERE A.sl_cd = B.sl_cd " & _
						" AND A.item_cd = C.item_cd " & _
						" AND A.plant_cd *= D.plant_cd " & _
						" AND A.sl_cd *= D.sl_cd " & _
						" AND A.item_cd *= D.item_cd " & _
						" AND A.tracking_no *= D.tracking_no " & _
						" AND A.plant_cd = "     & pCode & _
						" AND A.item_cd = "      & pCode1 & _
						" ORDER BY A.sl_cd "
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
		    .ggoSpread.Source = .frm1.vspdData2
			.lgStrPrevKeyIndex2 = "<%=lgStrPrevKeyIndex%>"
			
			.ggoSpread.SSShowData "<%=ConvSPChars(lgstrData)%>"

        	If .frm1.vspdData2.MaxRows < .VisibleRowCnt(.frm1.vspdData2,0)  And .lgStrPrevKeyIndex2 <> "" Then	 ' GroupView 사이즈로 화면 Row수보다 쿼리가 작으면 다시 쿼리함 
				.DbDtlQuery(.frm1.vspdData1.ActiveRow)				
			Else
				.DbDtlQueryOk()				
			End If
			.frm1.vspdData1.focus
		End If   

	End With	
       
</Script>	  

