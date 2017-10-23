<%@  LANGUAGE = VBSCript%>
<% Option Explicit%>
<!--'**********************************************************************************************
'*  1. Module Name          : Inventory List onhand stock detail
'*  2. Function Name        : 
'*  3. Program ID           : I1211pb1.asp
'*  4. Program Name         : 
'*  5. Program Desc         : 수불품목팝업 
'*  6. Comproxy List        : 
'                             +B25018ListPlant
'                             +B25019LookUpPlant
'*  7. Modified date(First) : 2002/04/03
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Lee Seung Wook
'* 10. Modifier (Last)      : Lee Seung Wook
'* 11. Comment              :
'**********************************************************************************************-->
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<%             
On Error Resume Next
Err.Clear

Call LoadBasisGlobalInf()
Call LoadInfTB19029B("Q", "I","NOCOOKIE","MB")
Call HideStatusWnd
         
Dim strBuildAsstNo
'Dim strBuildAsstNoNm
Dim IntRetCD

lgLngMaxRow       = Request("txtMaxRows")
lgMaxCount        = 100                  
lgStrPrevKeyIndex = UNICInt(Trim(Request("lgStrPrevKeyIndex")),0)   
lgErrorStatus     = "NO"
lgErrorPos        = ""                   
lgOpModeCRUD      = Request("txtMode")   

Call SubOpenDB(lgObjConn)

'Call SubCreateCommandObject(lgObjComm)

strBuildAsstNo  = FilterVar(Request("txtBuildAsstNo") & "%", "''", "S")
'strBuildAsstNoNm  = FilterVar("%" & Trim(Request("txtItemNm1")) & "%", "''", "S")

Call SubBizQuery()

 
'Call SubCloseCommandObject(lgObjComm)    
Call SubCloseDB(lgObjConn)      


'============================================================================================================
' Name : SubBizQuery
'============================================================================================================
Sub SubBizQuery()
 
 Dim iDx
 Dim PvArr

On Error Resume Next
Err.Clear

	'Call SubMakeSQLStatements("MR",strBuildAsstNo,strBuildAsstNoNm)
	Call SubMakeSQLStatements("MR",strBuildAsstNo)
	 
	If  FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
		IntRetCD = -1
		lgStrPrevKeyIndex = ""  
		Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT) 
		Call SetErrorStatus()
		Call SubCloseRs(lgObjRs)
'Call ServerMesgBox("a" , vbInformation, I_MKSCRIPT)		    
		Response.End 
	Else
	
		IntRetCD = 1
		'Call SubSkipRs(lgObjRs,lgMaxCount * lgStrPrevKeyIndex)

	    lgstrData = ""
	    iDx       = 1

	    Do While Not lgObjRs.EOF

	        lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs(0)) 
	        lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs(1)) 	        
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs(2))
			'lgstrData = lgstrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(lgObjRs(2),PopupParent.gCurrency,PopupParent.ggAmtOfMoneyNo,PopupParent.gLocRndPolicyNo,"X") 
			'lgstrData = lgstrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(lgObjRs(3),PopupParent.gCurrency,PopupParent.ggAmtOfMoneyNo,PopupParent.gLocRndPolicyNo,"X") 
			'lgstrData = lgstrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(lgObjRs(4),PopupParent.gCurrency,PopupParent.ggAmtOfMoneyNo,PopupParent.gLocRndPolicyNo,"X") 
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs(3))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs(4))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs(5))
			lgstrData = lgstrData & Chr(11) & lgLngMaxRow + iDx
			lgstrData = lgstrData & Chr(11) & Chr(12)
				
			lgObjRs.MoveNext

	        iDx =  iDx + 1
	        If iDx > lgMaxCount Then
	           lgStrPrevKeyIndex = lgStrPrevKeyIndex + 1
	           Exit Do
	        End If   
	    Loop 
'Call ServerMesgBox(lgstrData , vbInformation, I_MKSCRIPT)		

	
	End If
	   
	If iDx <= lgMaxCount Then
	   lgStrPrevKeyIndex = ""
	End If   
	
	Call SubHandleError("MR",lgObjConn,lgObjRs,Err)
	Call SubCloseRs(lgObjRs)                                             
'Call ServerMesgBox(lgstrData , vbInformation, I_MKSCRIPT)		    
	lgStrSQL = ""
    
End Sub 

'============================================================================================================
' Name : SubMakeSQLStatements
'============================================================================================================
'Sub SubMakeSQLStatements(pDataType,pCode,pCode1)
Sub SubMakeSQLStatements(pDataType,pCode)

    Dim iSelCount
    
    Const C_SHEETMAXROWS_D  = 100  
    
    
    On Error Resume Next                                                     
    Err.Clear                                                                
    
		iSelCount = C_SHEETMAXROWS_D + C_SHEETMAXROWS_D *  lgStrPrevKeyIndex + 1    
    
		lgStrSQL = "" 
		lgStrSQL = lgStrSQL & " SELECT	TOP " & iSelCount & " ACCT_CD, ACCT_NM, CTRL_VAL, "
		lgStrSQL = lgStrSQL & " 		SUM(DR_AMT) 'DR_AMT', "
		lgStrSQL = lgStrSQL & " 		SUM(CR_AMT) 'CR_AMT', "
		lgStrSQL = lgStrSQL & " 		SUM(DR_AMT)-SUM(CR_AMT) 'BAL_AMT' "
		lgStrSQL = lgStrSQL & " FROM	( "
		lgStrSQL = lgStrSQL & " 		SELECT	A.ACCT_CD, C.ACCT_NM, B.CTRL_VAL, B.ITEM_SEQ, "
		lgStrSQL = lgStrSQL & " 				ISNULL((CASE WHEN A.DR_CR_FG = 'DR' THEN ITEM_LOC_AMT END),0) 'DR_AMT', "
		lgStrSQL = lgStrSQL & " 				ISNULL((CASE WHEN A.DR_CR_FG = 'CR' THEN ITEM_LOC_AMT END),0) 'CR_AMT' "
		lgStrSQL = lgStrSQL & " 		FROM	A_GL_ITEM A (NOLOCK), "
		lgStrSQL = lgStrSQL & " 				A_GL_DTL  B (NOLOCK), "
		lgStrSQL = lgStrSQL & " 				A_ACCT    C (NOLOCK)  "
		lgStrSQL = lgStrSQL & " 		WHERE	A.GL_NO = B.GL_NO "
		lgStrSQL = lgStrSQL & " 				AND A.ITEM_SEQ = B.ITEM_SEQ "		
		lgStrSQL = lgStrSQL & " 				AND A.ACCT_CD = C.ACCT_CD "
		lgStrSQL = lgStrSQL & " 				AND B.CTRL_CD = 'CA' "
		lgStrSQL = lgStrSQL & " 				AND B.CTRL_VAL LIKE " & pCode		
		lgStrSQL = lgStrSQL & " 		) AA "
		lgStrSQL = lgStrSQL & " GROUP BY ACCT_CD, ACCT_NM, CTRL_VAL "
		lgStrSQL = lgStrSQL & " ORDER BY ACCT_CD, CTRL_VAL "

'Call ServerMesgBox(lgStrSQL , vbInformation, I_MKSCRIPT)
'response.end   
End Sub    

'============================================================================================================
' Name : CommonOnTransactionAbort
'============================================================================================================
Sub CommonOnTransactionAbort()
    lgErrorStatus    = "YES"
End Sub

'============================================================================================================
' Name : SetErrorStatus
'============================================================================================================
Sub SetErrorStatus()
    lgErrorStatus     = "YES"     
End Sub


'============================================================================================================
' Name : SubHandleError
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

	'.txtPlantNm.value = "<%=ConvSPChars(strPlantNm)%>"
	If "<%=lgErrorStatus%>" = "NO" And "<%=IntRetCd%>" <> -1 Then
	
		.ggoSpread.Source = .vspdData
		.lgStrPrevKeyIndex = "<%=lgStrPrevKeyIndex%>"
		.ggoSpread.SSShowData "<%=lgstrData%>"

		if .vspdData.MaxRows < .parent.VisibleRowCnt(.vspdData, 0) and .lgStrPrevKeyIndex <> "" Then
			.DbQuery
		Else
			.DbQueryOk
		End If
	End If   
End With 
</Script> 


 

