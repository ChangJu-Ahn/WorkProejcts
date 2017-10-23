<%@ LANGUAGE="VBSCRIPT" %>
<%Option Explicit    %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Prucurement
'*  2. Function Name        : 
'*  3. Program ID           : ma112qB1
'*  4. Program Name         : 미착경비현황조회 
'*  5. Program Desc         : 미착경비현황조회 
'*  6. Component List       : 
'*  7. Modified date(First) : 2003/06/20
'*  8. Modified date(Last)  : 2003/06/20
'*  9. Modifier (First)     : Kang Su Hwan
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
lgStrPrevKeyIndex = UNICInt(Trim(Request("lgStrPrevKey")),0)   
lgErrorStatus     = "NO"
lgErrorPos        = ""                                                           '☜: Set to space
'------ Developer Coding part (Start ) ------------------------------------------------------------------
Dim IntRetCD
        
Dim strChargeFrDt
Dim strChargeToDt
Dim strDocFrDt
Dim strDocToDt

strChargeFrDt	= UNIConvDate(Trim(Request("txtChargeFrDt")))
strChargeToDt	= UNIConvDate(Trim(Request("txtChargeToDt")))
strDocFrDt		= UNIConvDate(Trim(Request("txtDocFrDt")))
strDocToDt		= UNIConvDate(Trim(Request("txtDocToDt")))

'------ Developer Coding part (End   ) ------------------------------------------------------------------ 

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
	Const C_SHEETMAXROWS_D  = 100
	Dim iDx
	Dim PvArr
	
	On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear
    
    '---------- Developer Coding part (Start) ---------------------------------------------------------------
	Call SubMakeSQLStatements("AT",strChargeFrDt,strChargeToDt,strDocFrDt,strDocToDt)           '☜ : Make sql statements

	If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then                    'If data not exists	
		IntRetCD = -1
		lgStrPrevKeyIndex = ""		
		Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
		Call SetErrorStatus()
		Call SubCloseRs(lgObjRs)
		Response.Write "<script language=""vbscript"">" & vbcr
		Response.Write "	Call parent.SetFocusToDocument(""M"")" & vbcr
		Response.Write "	parent.frm1.txtChargeFrDt.focus" & vbcr
		Response.Write "</script>" & vbcr
		Response.End 
	Else
		IntRetCD = 1
        lgstrData = ""
        iDx       = -1
	    ReDim PvArr(C_SHEETMAXROWS_D - 1)
        
        Do While Not lgObjRs.EOF
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs(0))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs(1))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs(2))
            lgstrData = lgstrData & Chr(11) & lgLngMaxRow + iDx + 2
            lgstrData = lgstrData & Chr(11) & Chr(12)

            iDx =  iDx + 1
            If iDx >= C_SHEETMAXROWS_D Then
               lgStrPrevKeyIndex = lgStrPrevKeyIndex + 1
               Exit Do
            End If   
               
            PvArr(iDx) = lgstrData	
			lgstrData = ""

		    lgObjRs.MoveNext
        Loop 
    End If
    lgstrData = join(PvArr,"")
	
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
		Case "AT"
			lgstrsql =  "SELECT EXP.ITEM_ACCT,EXP.BASE_CHARGE_AMT,MVMT.COST_OF_DEVY " _
						& VBCR & " FROM (SELECT C.ITEM_ACCT,SUM(B.BASE_CHARGE_AMT) BASE_CHARGE_AMT FROM M_PURCHASE_CHARGE A " _
						& VBCR & VBTab & " INNER JOIN M_PURCHASE_EXPENSE_BY_ITEM B ON ( A.CHARGE_NO = B.CHARGE_NO ) " _
						& VBCR & VBTab & " LEFT OUTER JOIN B_ITEM_BY_PLANT C ON ( B.PLANT_CD=C.PLANT_CD AND B.ITEM_CD=C.ITEM_CD) " _
						& VBCR & VBTab & " WHERE A.CHARGE_DT >= '" & PCODE & " ' AND A.CHARGE_DT <= '" & PCODE1 & "' " _
						& VBCR & VBTab & " GROUP BY  C.ITEM_ACCT) EXP," _
						& VBCR & VBCR  & VBTab & " (SELECT C.ITEM_ACCT,SUM(B.COST_OF_DEVY) COST_OF_DEVY " _
						& VBCR & VBTab & " FROM I_GOODS_MOVEMENT_HEADER A " _
						& VBCR & VBTab & " INNER JOIN I_GOODS_MOVEMENT_DETAIL B ON ( A.ITEM_DOCUMENT_NO=B.ITEM_DOCUMENT_NO AND A.DOCUMENT_YEAR=B.DOCUMENT_YEAR) " _
						& VBCR & VBTab & " LEFT OUTER JOIN B_ITEM_BY_PLANT C ON ( B.PLANT_CD=C.PLANT_CD AND B.ITEM_CD=C.ITEM_CD) " _
						& VBCR & VBTab & " WHERE A.DOCUMENT_DT >=  " & FilterVar(PCODE2 , "''", "S") & " AND A.DOCUMENT_DT <=  " & FilterVar(PCODE3 , "''", "S") & "  " _
						& VBCR & VBTab & " AND B.TRNS_TYPE=" & FilterVar("PR", "''", "S") & " " _
						& VBCR & VBTab & " AND B.DELETE_FLAG=" & FilterVar("N", "''", "S") & "  " _
						& VBCR & VBTab & " GROUP BY  C.ITEM_ACCT) MVMT" _
						& VBCR & " WHERE EXP.ITEM_ACCT=MVMT.ITEM_ACCT"
End Select 

Response.Write "   <pre> " & vbcr
Response.Write lgstrsql & vbcr
Response.Write "   </pre>" & vbcr

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
		
		.frm1.txtHChargeFrDt.value	= "<%=Trim(Request("txtChargeFrDt"))%>"
		.frm1.txtHChargeToDt.value	= "<%=Trim(Request("txtChargeToDt"))%>"
		.frm1.txtHDocFrDt.value		= "<%=Trim(Request("txtDocFrDt"))%>"
		.frm1.txtHDocToDt.value		= "<%=Trim(Request("txtDocToDt"))%>"

		If "<%=lgErrorStatus%>" = "NO" And "<%=IntRetCd%>" <> -1 Then
		    .ggoSpread.Source	= .frm1.vspdData
			.lgStrPrevKeyIndex	= "<%=lgStrPrevKeyIndex%>"
			.ggoSpread.SSShowData "<%=lgstrData%>"

			.DbQueryOk

			.frm1.vspdData.focus
		End If   

	End With	
       
</Script>	

	
