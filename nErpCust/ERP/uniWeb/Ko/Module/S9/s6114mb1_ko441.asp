<%@ LANGUAGE="VBSCRIPT"  TRANSACTION=Required%>
<%Option Explicit    %>
<!--
'**********************************************************************************************
'*  1. Module Name          : 영업
'*  2. Function Name        : 
'*  3. Program ID           : S6114MB1_KO441
'*  4. Program Name         : 수출입경비조회
'*  5. Program Desc         : 
'*  6. Component List       : 
'*  7. Modified date(First) : 2007/12/26
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : 
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change" 
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************
-->

<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/lgSvrVariables.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->

<Script Language=vbscript>
	Dim strVar1
	Dim strVar2
	Dim TempstrBizArea

	TempstrBizArea = "<%=Request("txtHConBizArea")%>"
	
	'공장명 불러오기 
	
	Call parent.CommonQueryRs("biz_area_cd,biz_area_nm","B_biz_area","biz_area_cd =  " & parent.FilterVar(TempstrBizArea , "''", "S") & "",strVar1,strVar2,strVar3,strVar4,strVar5,strVar6,strVar7)
	strVar1 = Replace(strVar1,Chr(11), "")
	strVar2 = Replace(strVar2,Chr(11), "")
	
	Parent.frm1.txtConBizAreaNM.Value = strVar2	
	
</Script>

<%
	
    Const C_SHEETMAXROWS_D = 100
    
    Call HideStatusWnd    
    Call LoadBasisGlobalInf()
    Call LoadInfTB19029B("I", "*","NOCOOKIE", "MB")
    
    lgErrorStatus            = "NO"
    lgErrorPos               = ""                                                           '☜: Set to space
    lgOpModeCRUD             = Request("txtMode") 
       
	Dim BizArea
    Dim documentDt
    Dim lgStrColorFlag
	
    lgLngMaxRow              = Request("txtMaxRows")     
    BizArea 				 = Trim(UCase(Request("txtHConBizArea")))
    documentDt               = Trim(Ucase(Request("txtYr")))  
	
	
    Call SubOpenDB(lgObjConn)                                                        '☜: Make a DB Connection

    Call SubBizQueryMulti()
    
    Call SubCloseDB(lgObjConn)  
	
'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()
    Dim iDx
    Dim iLoopMax
    Dim iKey1, baseDt

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
	
    Call SubMakeSQLStatements("")													 '☆ : Make sql statements

    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        lgStrPrevKeyIndex = ""
        Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
        Call SetErrorStatus()
    Else

        Call SubSkipRs(lgObjRs,C_SHEETMAXROWS_D * lgStrPrevKeyIndex)

        lgstrData = ""
        
        iDx       = 1
        
        Do While Not lgObjRs.EOF
                 
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("COL0"))      
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("COL1"))      
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("COL2"))                
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("COL3"))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("COL4"))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("COL5"))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("COL6"))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("COL7"))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("COL8"))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("COL9"))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("COL10"))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("COL11"))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("COL12"))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("COL13"))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("COL14"))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("COL15"))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("COL16"))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("COL17"))
				
            lgstrData = lgstrData & Chr(11) & lgLngMaxRow + iDx
            lgstrData = lgstrData & Chr(11) & Chr(12)
		    lgObjRs.MoveNext

			'If lgObjRs(0) = "12" OR rs0(0) = "22" Then
			'	If lgObjRs(0) > 0 Then
			'		lgStrColorFlag = lgStrColorFlag & CStr(iLoopCount+1) & gColSep & rs0(0) & gRowSep
			'	End If
			'End If

            iDx =  iDx + 1
            If iDx > C_SHEETMAXROWS_D Then
               lgStrPrevKeyIndex = lgStrPrevKeyIndex + 1
            
				Exit Do
            End If   
               
        Loop 
    End If
    
    If iDx <= C_SHEETMAXROWS_D Then
       lgStrPrevKeyIndex = ""
    End If   

	 Call SubHandleError("MR",lgObjConn,lgObjRs,Err)
     Call SubCloseRs(lgObjRs)                                                          '☜: Release RecordSSet

End Sub    
 

'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements(pComp)
   
	Dim iSelCount 
    Dim lgGroupIndex
	
	iSelCount = C_SHEETMAXROWS_D + C_SHEETMAXROWS_D * lgStrPrevKeyIndex + 1   
   
	lgStrSQL = " SELECT TBL.COL0,TBL.COL1,TBL.COL2,TBL.COL3,TBL.COL4,TBL.COL5,TBL.COL6,TBL.COL7,TBL.COL8,TBL.COL9,  "
	lgStrSQL = lgStrSQL & "	TBL.COL10, TBL.COL11, TBL.COL12,TBL.COL13,TBL.COL14,TBL.COL15,TBL.COL16,TBL.COL17 "
	lgStrSQL = lgStrSQL & "	FROM ( SELECT '11' AS COL0, '수입' AS COL1, JNL_NM AS COL2, C1.BP_NM  AS COL3, D1.BP_NM  AS COL4, "    
	lgStrSQL = lgStrSQL & "					SUM(CHARGE_DOC_AMT) AS COL5, "                        
	lgStrSQL = lgStrSQL & "					SUM(CASE DATEPART(m,CHARGE_DT) WHEN 01 THEN PAY_LOC_AMT ELSE 0 END)  AS COL6, "
	lgStrSQL = lgStrSQL & "					SUM(CASE DATEPART(m,CHARGE_DT) WHEN 02 THEN PAY_LOC_AMT ELSE 0 END)  AS COL7, "
	lgStrSQL = lgStrSQL & "					SUM(CASE DATEPART(m,CHARGE_DT) WHEN 03 THEN PAY_LOC_AMT ELSE 0 END)  AS COL8, "
	lgStrSQL = lgStrSQL & "					SUM(CASE DATEPART(m,CHARGE_DT) WHEN 04 THEN PAY_LOC_AMT ELSE 0 END)  AS COL9, "
	lgStrSQL = lgStrSQL & "					SUM(CASE DATEPART(m,CHARGE_DT) WHEN 05 THEN PAY_LOC_AMT ELSE 0 END)  AS COL10, "
	lgStrSQL = lgStrSQL & "					SUM(CASE DATEPART(m,CHARGE_DT) WHEN 06 THEN PAY_LOC_AMT ELSE 0 END)  AS COL11, "
	lgStrSQL = lgStrSQL & "					SUM(CASE DATEPART(m,CHARGE_DT) WHEN 07 THEN PAY_LOC_AMT ELSE 0 END)  AS COL12, "
	lgStrSQL = lgStrSQL & "					SUM(CASE DATEPART(m,CHARGE_DT) WHEN 08 THEN PAY_LOC_AMT ELSE 0 END)  AS COL13, "
	lgStrSQL = lgStrSQL & "					SUM(CASE DATEPART(m,CHARGE_DT) WHEN 09 THEN PAY_LOC_AMT ELSE 0 END)  AS COL14, "
	lgStrSQL = lgStrSQL & "					SUM(CASE DATEPART(m,CHARGE_DT) WHEN 10 THEN PAY_LOC_AMT ELSE 0 END)  AS COL15, "
	lgStrSQL = lgStrSQL & "					SUM(CASE DATEPART(m,CHARGE_DT) WHEN 11 THEN PAY_LOC_AMT ELSE 0 END)  AS COL16, "
	lgStrSQL = lgStrSQL & "					SUM(CASE DATEPART(m,CHARGE_DT) WHEN 12 THEN PAY_LOC_AMT ELSE 0 END)  AS COL17 "      	
	lgStrSQL = lgStrSQL & "			FROM  M_PURCHASE_CHARGE A1, A_JNL_ITEM B1, B_BIZ_PARTNER C1, B_BIZ_PARTNER D1 "
	lgStrSQL = lgStrSQL & "		WHERE A1.CHARGE_TYPE = B1.JNL_CD AND A1.PAYEE_CD = C1.BP_CD AND A1.BP_CD = D1.BP_CD "
	lgStrSQL = lgStrSQL & "		AND  BIZ_AREA =  " & FilterVar(Trim(BizArea), "''", "S") & "  "
	lgStrSQL = lgStrSQL & "     AND  YEAR(CHARGE_DT) = " & FilterVar(Trim(documentDt), "''", "S") & " "
	lgStrSQL = lgStrSQL & "		AND B1.JNL_TYPE = 'EC' "                
	lgStrSQL = lgStrSQL & "		GROUP BY JNL_NM, C1.BP_NM, D1.BP_NM, DATEPART(m,CHARGE_DT), BIZ_AREA, YEAR(CHARGE_DT) "	
	lgStrSQL = lgStrSQL & "		UNION ALL "
	lgStrSQL = lgStrSQL & "		SELECT '12' AS COL0, '수입소계' AS COL1, '' AS COL2, '' AS COL3, '' AS COL4,  "   
	lgStrSQL = lgStrSQL & "		       SUM(CHARGE_DOC_AMT) AS COL5, "
	lgStrSQL = lgStrSQL & "		       SUM(CASE DATEPART(m,CHARGE_DT) WHEN 01 THEN PAY_LOC_AMT ELSE 0 END)  AS COL6, "
	lgStrSQL = lgStrSQL & "		       SUM(CASE DATEPART(m,CHARGE_DT) WHEN 02 THEN PAY_LOC_AMT ELSE 0 END)  AS COL7, "
	lgStrSQL = lgStrSQL & "		       SUM(CASE DATEPART(m,CHARGE_DT) WHEN 03 THEN PAY_LOC_AMT ELSE 0 END)  AS COL8, "
	lgStrSQL = lgStrSQL & "		       SUM(CASE DATEPART(m,CHARGE_DT) WHEN 04 THEN PAY_LOC_AMT ELSE 0 END)  AS COL9, "
	lgStrSQL = lgStrSQL & "			   SUM(CASE DATEPART(m,CHARGE_DT) WHEN 05 THEN PAY_LOC_AMT ELSE 0 END)  AS COL10, " 
	lgStrSQL = lgStrSQL & "		       SUM(CASE DATEPART(m,CHARGE_DT) WHEN 06 THEN PAY_LOC_AMT ELSE 0 END)  AS COL11, "
	lgStrSQL = lgStrSQL & "		       SUM(CASE DATEPART(m,CHARGE_DT) WHEN 07 THEN PAY_LOC_AMT ELSE 0 END)  AS COL12, "
	lgStrSQL = lgStrSQL & "		       SUM(CASE DATEPART(m,CHARGE_DT) WHEN 08 THEN PAY_LOC_AMT ELSE 0 END)  AS COL13, "
	lgStrSQL = lgStrSQL & "		       SUM(CASE DATEPART(m,CHARGE_DT) WHEN 09 THEN PAY_LOC_AMT ELSE 0 END)  AS COL14, "
	lgStrSQL = lgStrSQL & "		       SUM(CASE DATEPART(m,CHARGE_DT) WHEN 10 THEN PAY_LOC_AMT ELSE 0 END)  AS COL15, "
	lgStrSQL = lgStrSQL & "		       SUM(CASE DATEPART(m,CHARGE_DT) WHEN 11 THEN PAY_LOC_AMT ELSE 0 END)  AS COL16, " 
	lgStrSQL = lgStrSQL & "		       SUM(CASE DATEPART(m,CHARGE_DT) WHEN 12 THEN PAY_LOC_AMT ELSE 0 END)  AS COL17 "	      				    	
	lgStrSQL = lgStrSQL & "		FROM  M_PURCHASE_CHARGE A1, A_JNL_ITEM B1, B_BIZ_PARTNER C1, B_BIZ_PARTNER D1 "
	lgStrSQL = lgStrSQL & "		WHERE A1.CHARGE_TYPE = B1.JNL_CD AND A1.PAYEE_CD = C1.BP_CD AND A1.BP_CD = D1.BP_CD "
	lgStrSQL = lgStrSQL & "		AND  BIZ_AREA =  " & FilterVar(Trim(BizArea), "''", "S") & "  "
	lgStrSQL = lgStrSQL & "     AND  YEAR(CHARGE_DT) = " & FilterVar(Trim(documentDt), "''", "S") & " "
	lgStrSQL = lgStrSQL & "		AND B1.JNL_TYPE = 'EC' HAVING SUM(CHARGE_DOC_AMT) > 0 "   
	lgStrSQL = lgStrSQL & "		UNION ALL "
	lgStrSQL = lgStrSQL & "		SELECT '21' AS COL0, '수출' AS COL1, JNL_NM AS COL2, C2.BP_NM AS COL3, "              
	lgStrSQL = lgStrSQL & "				CASE WHEN A2.PROCESS_STEP = 'ED' THEN (SELECT BP_NM FROM B_BIZ_PARTNER WHERE BP_CD IN "
	lgStrSQL = lgStrSQL & "														(SELECT APPLICANT FROM S_CC_HDR WHERE CC_NO = A2.BAS_NO)) "
	lgStrSQL = lgStrSQL & "					 WHEN A2.PROCESS_STEP = 'EB' THEN (SELECT BP_NM FROM B_BIZ_PARTNER WHERE BP_CD IN "
	lgStrSQL = lgStrSQL & "														(SELECT SOLD_TO_PARTY FROM S_BILL_HDR WHERE BILL_NO =  A2.BAS_NO)) "
	lgStrSQL = lgStrSQL & "			       END  AS COL4, "
	lgStrSQL = lgStrSQL & "			       SUM(CHARGE_DOC_AMT) AS COL5, "
	lgStrSQL = lgStrSQL & "			       SUM(CASE DATEPART(m,CHARGE_DT) WHEN 01 THEN CHARGE_DOC_AMT ELSE 0 END)  AS COL6, "
	lgStrSQL = lgStrSQL & "			       SUM(CASE DATEPART(m,CHARGE_DT) WHEN 02 THEN CHARGE_DOC_AMT ELSE 0 END)  AS COL7, "
	lgStrSQL = lgStrSQL & "			       SUM(CASE DATEPART(m,CHARGE_DT) WHEN 03 THEN CHARGE_DOC_AMT ELSE 0 END)  AS COL8, "
	lgStrSQL = lgStrSQL & "			       SUM(CASE DATEPART(m,CHARGE_DT) WHEN 04 THEN CHARGE_DOC_AMT ELSE 0 END)  AS COL9, "
	lgStrSQL = lgStrSQL & "			       SUM(CASE DATEPART(m,CHARGE_DT) WHEN 05 THEN CHARGE_DOC_AMT ELSE 0 END)  AS COL10,"
	lgStrSQL = lgStrSQL & "			       SUM(CASE DATEPART(m,CHARGE_DT) WHEN 06 THEN CHARGE_DOC_AMT ELSE 0 END)  AS COL11," 
	lgStrSQL = lgStrSQL & "			       SUM(CASE DATEPART(m,CHARGE_DT) WHEN 07 THEN CHARGE_DOC_AMT ELSE 0 END)  AS COL12,"
	lgStrSQL = lgStrSQL & "			       SUM(CASE DATEPART(m,CHARGE_DT) WHEN 08 THEN CHARGE_DOC_AMT ELSE 0 END)  AS COL13," 
	lgStrSQL = lgStrSQL & "			       SUM(CASE DATEPART(m,CHARGE_DT) WHEN 09 THEN CHARGE_DOC_AMT ELSE 0 END)  AS COL14," 
	lgStrSQL = lgStrSQL & "			       SUM(CASE DATEPART(m,CHARGE_DT) WHEN 10 THEN CHARGE_DOC_AMT ELSE 0 END)  AS COL15," 
	lgStrSQL = lgStrSQL & "			       SUM(CASE DATEPART(m,CHARGE_DT) WHEN 11 THEN CHARGE_DOC_AMT ELSE 0 END)  AS COL16,"  
	lgStrSQL = lgStrSQL & "			       SUM(CASE DATEPART(m,CHARGE_DT) WHEN 12 THEN CHARGE_DOC_AMT ELSE 0 END)  AS COL17	"      
	lgStrSQL = lgStrSQL & "		FROM S_SALES_CHARGE A2, A_JNL_ITEM B2, B_BIZ_PARTNER C2 "
	lgStrSQL = lgStrSQL & "		WHERE A2.CHARGE_CD = B2.JNL_CD AND A2.BP_CD = C2.BP_CD "
	lgStrSQL = lgStrSQL & "		AND  BIZ_AREA =  " & FilterVar(Trim(BizArea), "''", "S") & "  "
	lgStrSQL = lgStrSQL & "     AND  YEAR(CHARGE_DT) = " & FilterVar(Trim(documentDt), "''", "S") & " "
	lgStrSQL = lgStrSQL & "		AND B2.JNL_TYPE = 'EC' "
	lgStrSQL = lgStrSQL & "		GROUP BY JNL_NM, C2.BP_NM, DATEPART(m,CHARGE_DT), BIZ_AREA, YEAR(CHARGE_DT), A2.PROCESS_STEP, A2.BAS_NO "
	lgStrSQL = lgStrSQL & "		UNION ALL "
	lgStrSQL = lgStrSQL & "		SELECT '22' AS COL0, '수출소계' AS COL1, '' AS COL2, '' AS COL3, '' AS COL4, "               
	lgStrSQL = lgStrSQL & "			       SUM(CHARGE_DOC_AMT) AS COL5, "
	lgStrSQL = lgStrSQL & "			       SUM(CASE DATEPART(m,CHARGE_DT) WHEN 01 THEN CHARGE_DOC_AMT ELSE 0 END)  AS COL6,  "
	lgStrSQL = lgStrSQL & "			       SUM(CASE DATEPART(m,CHARGE_DT) WHEN 02 THEN CHARGE_DOC_AMT ELSE 0 END)  AS COL7, "
	lgStrSQL = lgStrSQL & "			       SUM(CASE DATEPART(m,CHARGE_DT) WHEN 03 THEN CHARGE_DOC_AMT ELSE 0 END)  AS COL8, "
	lgStrSQL = lgStrSQL & "			       SUM(CASE DATEPART(m,CHARGE_DT) WHEN 04 THEN CHARGE_DOC_AMT ELSE 0 END)  AS COL9, "
	lgStrSQL = lgStrSQL & "			       SUM(CASE DATEPART(m,CHARGE_DT) WHEN 05 THEN CHARGE_DOC_AMT ELSE 0 END)  AS COL10, "
	lgStrSQL = lgStrSQL & "			       SUM(CASE DATEPART(m,CHARGE_DT) WHEN 06 THEN CHARGE_DOC_AMT ELSE 0 END)  AS COL11, "
	lgStrSQL = lgStrSQL & "			       SUM(CASE DATEPART(m,CHARGE_DT) WHEN 07 THEN CHARGE_DOC_AMT ELSE 0 END)  AS COL12, "
	lgStrSQL = lgStrSQL & "			       SUM(CASE DATEPART(m,CHARGE_DT) WHEN 08 THEN CHARGE_DOC_AMT ELSE 0 END)  AS COL13, "
	lgStrSQL = lgStrSQL & "			       SUM(CASE DATEPART(m,CHARGE_DT) WHEN 09 THEN CHARGE_DOC_AMT ELSE 0 END)  AS COL14, "
	lgStrSQL = lgStrSQL & "			       SUM(CASE DATEPART(m,CHARGE_DT) WHEN 10 THEN CHARGE_DOC_AMT ELSE 0 END)  AS COL15, "
	lgStrSQL = lgStrSQL & "			       SUM(CASE DATEPART(m,CHARGE_DT) WHEN 11 THEN CHARGE_DOC_AMT ELSE 0 END)  AS COL16, " 
	lgStrSQL = lgStrSQL & "			       SUM(CASE DATEPART(m,CHARGE_DT) WHEN 12 THEN CHARGE_DOC_AMT ELSE 0 END)  AS COL17	 "     	
	lgStrSQL = lgStrSQL & "		FROM S_SALES_CHARGE A2, A_JNL_ITEM B2, B_BIZ_PARTNER C2 "
	lgStrSQL = lgStrSQL & "		WHERE A2.CHARGE_CD = B2.JNL_CD AND A2.BP_CD = C2.BP_CD "
	lgStrSQL = lgStrSQL & "		AND  BIZ_AREA =  " & FilterVar(Trim(BizArea), "''", "S") & "  "
	lgStrSQL = lgStrSQL & "     AND  YEAR(CHARGE_DT) = " & FilterVar(Trim(documentDt), "''", "S") & " "
	lgStrSQL = lgStrSQL & "		AND B2.JNL_TYPE = 'EC' HAVING SUM(CHARGE_DOC_AMT) > 0   "
	lgStrSQL = lgStrSQL & "	) TBL "
	lgStrSQL = lgStrSQL & "	 ORDER BY 1,2,3,4   "

Response.Write lgStrSQL
'Response.End                       

End Sub

'============================================================================================================
' Name : CommonOnTransactionCommit
' Desc : This Sub is called by OnTransactionCommit Error handler
'============================================================================================================
Sub CommonOnTransactionCommit()
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
    lgErrorStatus     = "YES"                                                         '☜: Set error status
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
    End Select
End Sub
%>

<Script Language="VBScript">
    
    Select Case "<%=lgOpModeCRUD %>"
       Case "<%=UID_M0001%>"                                                         '☜ : Query
          If Trim("<%=lgErrorStatus%>") = "NO" Then
              With Parent
                .ggoSpread.Source     = .frm1.vspdData
                .lgStrPrevKeyIndex = "<%=lgStrPrevKeyIndex%>"
                .ggoSpread.SSShowData "<%=lgstrData%>"
                .DBQueryOk        
	         End with
          End If   
       Case "<%=UID_M0002%>"                                                         '☜ : Save
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DBSaveOk
          Else
             Parent.SubSetErrPos(Trim("<%=lgErrorPos%>"))
          End If   
       Case "<%=UID_M0002%>"                                                         '☜ : Delete
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DbDeleteOk
          Else   
          End If   
    End Select    
    
       
</Script>	
 

    
    
