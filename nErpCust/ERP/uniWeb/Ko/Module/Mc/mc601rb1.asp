<%@ LANGUAGE="VBSCRIPT" %>
<%Option Explicit    %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Inventory
'*  2. Function Name        : 
'*  3. Program ID           : MC601RB1
'*  4. Program Name         : 납입지시대상 참조 
'*  5. Program Desc         : 납입지시대상 참조 
'*  6. Component List       : 
'*  7. Modified date(First) : 2003/03/03
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
Call LoadInfTB19029B("Q", "I","NOCOOKIE","RB")   
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
        
Dim strBpCd
Dim strDocumentDt1
Dim strDocumentDt2
Dim strDoTime

'------ Developer Coding part (End   ) ------------------------------------------------------------------ 

Call SubOpenDB(lgObjConn)
Call SubCreateCommandObject(lgObjComm)

Call SubBizQuery

Call SubCloseCommandObject(lgObjComm)    
Call SubCloseDB(lgObjConn)      

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()
	Const C_SHEETMAXROWS_D  = 100
	
	Dim iDx
	Dim PvArr
	
	On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear
    
    '---------- Developer Coding part (Start) ---------------------------------------------------------------
	Call SubMakeSQLStatements

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
        
            lgstrData = lgstrData & Chr(11) & UCase(ConvSPChars(lgObjRs(0)))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs(1))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs(2))
			lgstrData = lgstrData & Chr(11) & UCase(ConvSPChars(lgObjRs(3)))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs(4))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs(5))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs(6))
            lgstrData = lgstrData & Chr(11) & UniNumClientFormat(lgObjRs(7),ggQty.DecPoint,0)
            lgstrData = lgstrData & Chr(11) & UniNumClientFormat(lgObjRs(8),ggQty.DecPoint,0)
            lgstrData = lgstrData & Chr(11) & UCase(ConvSPChars(lgObjRs(9)))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs(10))
            lgstrData = lgstrData & Chr(11) & UNIDateClientFormat(lgObjRs(11))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs(12))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs(13))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs(14))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs(15))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs(16))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs(17))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs(18))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs(19))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs(20))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs(21))
            
            
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
            lgstrData = lgstrData & Chr(11) & lgLngMaxRow + iDx
            lgstrData = lgstrData & Chr(11) & Chr(12)

		    lgObjRs.MoveNext

            iDx =  iDx + 1
            If iDx-1 > C_SHEETMAXROWS_D Then
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
Sub SubMakeSQLStatements()

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
    
	Dim iSelCount
	
	'------ Developer Coding part (Start ) ------------------------------------------------------------------

	lgStrSQL = "SELECT A.PLANT_CD, B.PLANT_NM, A.PRODT_ORDER_NO, A.ITEM_CD, C.ITEM_NM, C.SPEC, A.PO_UNIT, A.DO_QTY_PO_UNIT, A.RCPT_QTY_PO_UNIT, D.SL_CD, F.SL_NM, A.DO_DATE, E.MINOR_NM, A.TRACKING_NO,G.RECV_INSPEC_FLG, A.PO_NO, A.PO_SEQ_NO, A.WC_CD, A.OPR_NO, A.SEQ, A.SUB_SEQ, D.PUR_GRP"
	lgStrSQL = lgStrSQL & " FROM	M_DLVY_ORD A, B_PLANT B, B_ITEM C, M_PRE_DLVY_ORD D, B_MINOR E, B_STORAGE_LOCATION F, B_ITEM_BY_PLANT G"
	lgStrSQL = lgStrSQL & " WHERE	A.PLANT_CD	=	B.PLANT_CD"
	lgStrSQL = lgStrSQL & " AND		A.ITEM_CD	= 	C.ITEM_CD"
	lgStrSQL = lgStrSQL & " AND 	A.PO_NO		=	D.PO_NO"
	lgStrSQL = lgStrSQL & " AND		A.PO_SEQ_NO	=	D.PO_SEQ_NO"
	lgStrSQL = lgStrSQL & " AND		A.DO_TIME	=	E.MINOR_CD"
	lgStrSQL = lgStrSQL & " AND		E.MAJOR_CD	=	" & FilterVar("M2110", "''", "S") & ""
	lgStrSQL = lgStrSQL & " AND		D.SL_CD		=	F.SL_CD"
	lgStrSQL = lgStrSQL & " AND		A.PLANT_CD	=	G.PLANT_CD"
	lgStrSQL = lgStrSQL & " AND		A.ITEM_CD	=	G.ITEM_CD"
	lgStrSQL = lgStrSQL & " AND 	A.DO_STATUS	=	" & "" & FilterVar("IS", "''", "S") & ""
	lgStrSQL = lgStrSQL & " AND		A.bp_cd		=	" & FilterVar(Request("txtBpCd"), "''", "S")

	IF Trim(Request("cboDoTime")) <> "" Then
	lgStrSQL = lgStrSQL & " AND		A.do_time	=	" & FilterVar(Request("cboDoTime"), "''", "S")
	End if
	If Trim(Request("txtDocumentDt1")) <> "" Then
	lgStrSQL = lgStrSQL & " AND		A.do_date	>=	" & FilterVar(Request("txtDocumentDt1"), "''", "S")
	End if
	If Trim(Request("txtDocumentDt2")) <> "" Then	
	lgStrSQL = lgStrSQL & " AND		A.do_date	<=	" & FilterVar(Request("txtDocumentDt2"), "''", "S")
	End if
	lgStrSQL = lgStrSQL & " ORDER BY A.PLANT_CD, A.PRODT_ORDER_NO, A.ITEM_CD, A.PO_NO, A.PO_SEQ_NO  "
   
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


	
