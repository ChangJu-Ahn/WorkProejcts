<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<!--
'======================================================================================================
'*  1. Module Name          : Inventory
'*  2. Function Name        : 
'*  3. Program ID           : i2321qb2.asp
'*  4. Program Name         : Tracking별 재고현황조회 
'*  5. Program Desc         : 
'*  6. Modified date(First) : 2005/03/10
'*  7. Modified date(Last)  : 
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

Dim strItemCd
Dim strPlantCd
Dim strTrackingNo
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
	strTrackingNo = FilterVar(Request("txtTrackingNo"), "''", "S")
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
    
		Call SubMakeSQLStatements("AL",strPlantCd,strItemCd,strTrackingNo)                                 
    
	If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then                 
		IntRetCD = -1
		lgStrPrevKeyIndex = ""		
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
						Chr(11) & ConvSPChars(lgObjRs(3)) & _
						Chr(11) & ConvSPChars(lgObjRs(4)) & _
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
						Chr(11) & UniConvNumberDBToCompany(lgObjRs(15), ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0) & _
						Chr(11) & UniConvNumberDBToCompany(lgObjRs(16), ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0) & _
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
Sub SubMakeSQLStatements(pDataType,pCode,pCode1,pCode2)

    On Error Resume Next                                                            
    Err.Clear                                                                       
    
	Dim iSelCount

	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	Select Case pDataType
		Case "AL"
			lgStrSQL =	"SELECT A.SL_CD, B.SL_NM, A.ITEM_CD, C.ITEM_NM, A.TRACKING_NO,A.GOOD_ON_HAND_QTY, A.BAD_ON_HAND_QTY," _
					 &	"A.STK_ON_INSP_QTY, A.STK_IN_TRNS_QTY,A.SCHD_RCPT_QTY,A.SCHD_ISSUE_QTY,A.PREV_GOOD_QTY," _
					 &	"A.PREV_BAD_QTY, A.PREV_STK_ON_INSP_QTY,A.PREV_STK_IN_TRNS_QTY, A.ALLOCATION_QTY, D.PICKING_QTY " _
					 &	"FROM I_ONHAND_STOCK A(nolock) join B_STORAGE_LOCATION B(nolock) on A.SL_CD = B.SL_CD " _
					 &	"join B_ITEM C(nolock) on A.ITEM_CD = C.ITEM_CD " _
					 &	"join (SELECT PLANT_CD,SL_CD,ITEM_CD,TRACKING_NO,SUM(PICKING_QTY) PICKING_QTY " _
					 &	"FROM I_ONHAND_STOCK_DETAIL(nolock) GROUP BY PLANT_CD,SL_CD,ITEM_CD,TRACKING_NO) D " _
					 &	"on A.PLANT_CD = D.PLANT_CD AND A.SL_CD = D.SL_CD " _
					 &	"AND A.ITEM_CD = D.ITEM_CD AND A.TRACKING_NO = D.TRACKING_NO " _
					 &	" WHERE		A.PLANT_CD = "		& pCode _
					 &	" AND		A.ITEM_CD = "		& pCode1
			If pCode2 <> "'*'" Then
				lgStrSQL = lgStrSQL & " AND A.TRACKING_NO = " & pCode2
			End If						 
				lgStrSQL = lgStrSQL & " ORDER BY	A.SL_CD "
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

