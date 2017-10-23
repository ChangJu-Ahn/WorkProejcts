<%@ LANGUAGE="VBSCRIPT" %>
<%Option Explicit    %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Prucurement
'*  2. Function Name        : 
'*  3. Program ID           : MC201pb1
'*  4. Program Name         : 납입지시대상 팝업 
'*  5. Program Desc         : 납입지시대상 팝업 
'*  6. Component List       : 
'*  7. Modified date(First) : 2003/02/24
'*  8. Modified date(Last)  : 2003/05/22
'*  9. Modifier (First)     : Lee Seung Wook
'* 10. Modifier (Last)      : Kang Su Hwan
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change" 
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************
-->

<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->

<%		
Call LoadBasisGlobalInf
Call LoadInfTB19029B("Q", "I","NOCOOKIE","PB")   
Call HideStatusWnd 
											
On Error Resume Next
Err.Clear

'---------------------------------------Common-----------------------------------------------------------

lgLngMaxRow       = Cint(Request("txtMaxRows"))                                        '☜: Read Operation Mode (CRUD)
lgStrPrevKeyIndex = UNICInt(Trim(Request("lgStrPrevKeyIndex")),0)   
lgErrorStatus     = "NO"
lgErrorPos        = ""                                                           '☜: Set to space
'------ Developer Coding part (Start ) ------------------------------------------------------------------
Dim IntRetCD
        
Dim strPlantCd
Dim strItemCd
Dim strTrackingNo
Dim strReqQty
Dim strBpCd

strPlantCd			= FilterVar(Request("txtPlantCd"), "''", "S")
strItemCd			= FilterVar(Request("txtItemCd"), "''", "S")
strTrackingNo		= FilterVar(Request("txtTrackingNo"), "''", "S")
strReqQty			= FilterVar(Request("hReqQty"), "''", "S")
strBpCd				= FilterVar(Request("hBpCd"), "''", "S")

'------ Developer Coding part (End   ) ------------------------------------------------------------------ 

Call SubOpenDB(lgObjConn)
Call SubCreateCommandObject(lgObjComm)

Call SubBizQuery("AL") 

Call SubCloseCommandObject(lgObjComm)    
Call SubCloseDB(lgObjConn)      

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery(pType)
	Const C_SHEETMAXROWS_D  = 100
	
	Dim iDx
	Dim PvArr
	
	On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear
    

    
    '---------- Developer Coding part (Start) ---------------------------------------------------------------
	Call SubMakeSQLStatements("AL",strPlantCd,strItemCd,strTrackingNo,strReqQty)           '☜ : Make sql statements
		
	
				
	If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then                    'If data not exists	
		IntRetCD = -1
		lgStrPrevKeyIndex = ""		
		Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
		Call SetErrorStatus()

		Call SubCloseRs(lgObjRs)

		Response.End 
	Else
		
	
		IntRetCD = 1
		 Call SubSkipRs(lgObjRs,C_SHEETMAXROWS_D * lgStrPrevKeyIndex)
        lgstrData = ""
        iDx       = 1
        
	    ReDim PvArr(C_SHEETMAXROWS_D - 1)
        Do While Not lgObjRs.EOF
        
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs(0))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs(1))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs(2))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs(3))
            lgstrData = lgstrData & Chr(11) & UniNumClientFormat(lgObjRs(4),ggQty.DecPoint,0)
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs(5))
            lgstrData = lgstrData & Chr(11) & UniNumClientFormat(lgObjRs(6),ggQty.DecPoint,0)
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs(7))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs(8))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs(9))
                        
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
            lgstrData = lgstrData & Chr(11) & lgLngMaxRow + iDx
            lgstrData = lgstrData & Chr(11) & Chr(12)

		    lgObjRs.MoveNext

            iDx =  iDx + 1
            If iDx-1 >= C_SHEETMAXROWS_D Then
               lgStrPrevKeyIndex = lgStrPrevKeyIndex + 1
               Exit Do
            End If   
               
            PvArr(iDx-1) = lgstrData	
			lgstrData = ""
        Loop 
        lgstrData = join(PvArr,"")
        
        
    End If
    If iDx <= C_SHEETMAXROWS_D Then
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
Sub SubMakeSQLStatements(pDataType,pCode,pCode1,pCode2,pCode3)

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
    
	Dim iSelCount
	
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
Select Case pDataType
		
		Case "AL"
			lgStrSQL = "SELECT A.PO_NO,A.PO_SEQ_NO,A.BP_CD,B.BP_NM,A.PO_QTY,A.PO_UNIT,A.BASE_QTY,A.BASE_UNIT,A.PUR_GRP,A.PR_NO" & _
			 " FROM		M_PRE_DLVY_ORD A,B_BIZ_PARTNER B" & _
			 " WHERE	A.BP_CD 			=	B.BP_CD" & _
			 " AND 		A.BASE_QTY			>=	A.PO_DLY_QTY" & _
			 " AND		A.PLANT_CD			=	" & strPlantCd & _
			 " AND		A.ITEM_CD			=	" & strItemCd & _
			 " AND		A.TRACKING_NO		=	" & strTrackingNo & _
			 " AND		A.BASE_QTY			>=	" & strReqQty & _
			 " AND		A.PO_DLY_QTY		<>	" & 0 & _
			 " AND		A.BP_CD				<>	" & strBpCd & _
			 " ORDER BY A.ITEM_CD "
End Select    
   
   
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
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
    lgErrorStatus     = "YES"                                                         '☜: Set error status
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub

'============================================================================================================
' Name : SubHandleError
' Desc : This Sub handle error
'============================================================================================================
Sub SubHandleError(pOpCode,pConn,pRs,pErr)
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

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
		    .ggoSpread.Source	= .frm1.vspdData
			.lgStrPrevKeyIndex	= "<%=lgStrPrevKeyIndex%>"
			.ggoSpread.SSShowData "<%=lgstrData%>"

        	If .frm1.vspdData.MaxRows < .VisibleRowCnt(.frm1.vspdData,0)  And .lgStrPrevKeyIndex <> "" Then	 ' GroupView 사이즈로 화면 Row수보다 쿼리가 작으면 다시 쿼리함 
				.DbQuery
				
			Else
				.DbQueryOk
				
			End If
			.frm1.vspdData.focus
		End If   

	End With	
       
</Script>	


	
