<%@ LANGUAGE=VBSCript TRANSACTION=Required%>

<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/IncServerAdoDb.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<%
    call LoadBasisGlobalInf()
    
    Call HideStatusWnd                                                               '¢Ð: Hide Processing message

    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '¢Ð: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '¢Ð: Read Operation Mode (CRUD)
    lgKeyStream       = Split(Request("txtKeyStream"),gColSep)

    lgLngMaxRow       = Request("txtMaxRows")                                        '¢Ð: Read Operation Mode (CRUD)

    Call SubOpenDB(lgObjConn)                                                        '¢Ð: Make a DB Connection
    
    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '¢Ð: Query
             Call SubBizQuery()
    End Select
    
    Call SubCloseDB(lgObjConn)                                                       '¢Ð: Close DB Connection

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()
    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status
    Call SubBizQueryMulti()
End Sub    
'============================================================================================================
' Name : SubBizSave
' Desc : Save Data 
'============================================================================================================
Sub SubBizSave()
    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status
End Sub
'============================================================================================================
' Name : SubBizDelete
' Desc : Delete DB data
'============================================================================================================
Sub SubBizDelete()
    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status
End Sub

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()
    Dim iDx
    Dim iLoopMax
    Dim iKey1
    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status

    Dim strWhere, strChang_dt
   
   
    If lgKeyStream(1) = "" then
       strWhere = strWhere & " AND H.DEPT_CD LIKE " & "" & FilterVar("%", "''", "S") & ""
    Else
       strWhere = strWhere & " AND H.DEPT_CD LIKE " & FilterVar(lgKeyStream(1), "''", "S")
    End if


    If  lgKeyStream(3) = "" then
        strWhere = strWhere & " AND H.EMP_NO LIKE  " & FilterVar(Trim(lgKeyStream(3)) & "%", "''", "S") & " " 
    else
        strWhere = strWhere & " AND H.EMP_NO  = " & FilterVar(lgKeyStream(3), "''", "S")
    end if
    	               
    lgStrSQL = "SELECT  H.EMP_NO, H.NAME EMP_NAME, D.ORG_CHANGE_ID, H.DEPT_CD, D.DEPT_NM, " 
    lgStrSQL = lgStrSQL & " ISNULL(V.DIR_INDIR,  H.DIR_INDIR) DIR_INDIR,  ISNULL(M.MINOR_NM, P.MINOR_NM)  DIR_INDIR_NM,  "    
    lgStrSQL = lgStrSQL & " ISNULL(V.COST_CD,  N.COST_CD) COST_CD,  ISNULL(Q.COST_NM, N.COST_NM) COST_NM, "       
    lgStrSQL = lgStrSQL & "  Z.BIZ_AREA_CD, Z.BIZ_AREA_NM  "
    lgStrSQL = lgStrSQL & " FROM   HAA010T  H "
         							     
    lgStrSQL = lgStrSQL & " LEFT OUTER  JOIN  C_INDV_COSTCENTER_KO441   V  ON    H.EMP_NO  = V.EMP_NO  "
    lgStrSQL = lgStrSQL & "				 AND  V.YYYYMM IN ( SELECT MAX(YYYYMM)  FROM  C_INDV_COSTCENTER_KO441 "        
    lgStrSQL = lgStrSQL & " 								WHERE   YYYYMM  < "    & FilterVar(lgKeyStream(0)& "01" , "''", "S") &  ")"  
    lgStrSQL = lgStrSQL & " LEFT OUTER  JOIN  B_COST_CENTER  Q  ON  V.COST_CD = Q.COST_CD  "    
    lgStrSQL = lgStrSQL & " LEFT OUTER  JOIN  B_MINOR  P ON   P.MAJOR_CD  ='H0071'  AND P.MINOR_CD = V.DIR_INDIR "        
    
   ' lgStrSQL = lgStrSQL & " LEFT OUTER  JOIN  C_INDV_COSTCENTER_KO441   V  ON    H.EMP_NO  = V.EMP_NO  "    
    lgStrSQL = lgStrSQL & " LEFT OUTER  JOIN  B_ACCT_DEPT   D  ON    H.DEPT_CD  = D.DEPT_CD  "
    lgStrSQL = lgStrSQL & " LEFT OUTER  JOIN  B_MINOR M ON   M.MAJOR_CD  ='H0071'  AND M.MINOR_CD = H.DIR_INDIR  "
    lgStrSQL = lgStrSQL & " LEFT OUTER  JOIN  B_COST_CENTER  N  ON  D.COST_CD = N.COST_CD "
    lgStrSQL = lgStrSQL & " LEFT OUTER  JOIN  B_BIZ_AREA Z  ON  N.BIZ_AREA_CD = Z.BIZ_AREA_CD   "
    lgStrSQL = lgStrSQL & " WHERE   H.EMP_NO NOT  IN ( SELECT EMP_NO FROM  C_INDV_COSTCENTER_KO441 "        
    lgStrSQL = lgStrSQL & "        						WHERE  YYYYMM = "  & FilterVar(lgKeyStream(0), "''", "S")   & ")"   
    lgStrSQL = lgStrSQL & "       AND  D.ORG_CHANGE_ID  IN (SELECT MAX(A.ORGID)"         
    lgStrSQL = lgStrSQL & " 								FROM  HORG_ABS A "       
    lgStrSQL = lgStrSQL & "                                WHERE  ORGDT <= "  & FilterVar(lgKeyStream(0)& "01" , "''", "S") &  ")"  
    lgStrSQL = lgStrSQL & " AND  (RETIRE_DT IS NULL OR RETIRE_DT >= "    & FilterVar(lgKeyStream(0)& "01" , "''", "S") &  ")"         
    lgStrSQL = lgStrSQL & " AND  D.ORG_CHANGE_ID IN (SELECT MAX(ORGID) FROM HORG_ABS WHERE  ORGDT <= " & FilterVar(lgKeyStream(0)& "01" , "''", "S") &  ")"                                               
    lgStrSQL = lgStrSQL &   strWhere
    lgStrSQL = lgStrSQL & " ORDER BY   H.EMP_NO "


	'call svrmsgbox  (lgStrSQL, 0,1)
   							
	                         
        
        
        
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '¢Ð : No data is found. 
        Call SetErrorStatus()
    Else

        lgstrData = ""

        Do While Not lgObjRs.EOF

            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("EMP_NO"))            
            lgstrData = lgstrData & Chr(11) & ""
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("EMP_NAME"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ORG_CHANGE_ID"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("DEPT_CD"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("DEPT_NM"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("DIR_INDIR"))
            lgstrData = lgstrData & Chr(11) & ""
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("DIR_INDIR_NM"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("COST_CD"))
            lgstrData = lgstrData & Chr(11) & ""
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("COST_NM"))            
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("BIZ_AREA_CD"))
            lgstrData = lgstrData & Chr(11) & ""
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("BIZ_AREA_NM"))                       
            lgstrData = lgstrData & Chr(11) & lgLngMaxRow + iDx
            lgstrData = lgstrData & Chr(11) & Chr(12)
            
		    lgObjRs.MoveNext
               
        Loop 
    End If
    
	Call SubHandleError("MR",lgObjConn,lgObjRs,Err)
    Call SubCloseRs(lgObjRs)                                                          '¢Ð: Release RecordSSet

End Sub    


'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data 
'============================================================================================================
Sub SubBizSaveMulti()

    Dim arrRowVal
    Dim arrColVal
    Dim iDx

    On Error Resume Next                                                             '¢Ð: Protect system from crashing

    Err.Clear                                                                        '¢Ð: Clear Error status
    
	arrRowVal = Split(Request("txtSpread"), gRowSep)                                 '¢Ð: Split Row    data
	
    For iDx = 1 To lgLngMaxRow
        arrColVal = Split(arrRowVal(iDx-1), gColSep)                                 '¢Ð: Split Column data
        
        Select Case arrColVal(0)
            Case "C"
                    Call SubBizSaveMultiCreate(arrColVal)                            '¢Ð: Create
            Case "U"
                    Call SubBizSaveMultiUpdate(arrColVal)                            '¢Ð: Update
            Case "D"
                    Call SubBizSaveMultiDelete(arrColVal)                            '¢Ð: Delete
        End Select
        
        If lgErrorStatus    = "YES" Then
           lgErrorPos = lgErrorPos & arrColVal(1) & gColSep
           Exit For
        End If
        
    Next

End Sub    

'============================================================================================================
' Name : SubBizSaveCreate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveMultiCreate(arrColVal)
End Sub
'============================================================================================================
' Name : SubBizSaveMultiUpdate
' Desc : Update Data from Db
'============================================================================================================
Sub SubBizSaveMultiUpdate(arrColVal)
End Sub

'============================================================================================================
' Name : SubBizSaveMultiDelete
' Desc : Delete Data from Db
'============================================================================================================
Sub SubBizSaveMultiDelete(arrColVal)
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
    lgErrorStatus     = "YES"                                                         '¢Ð: Set error status
End Sub
'============================================================================================================
' Name : SubHandleError
' Desc : This Sub handle error
'============================================================================================================
Sub SubHandleError(pOpCode,pConn,pRs,pErr)
    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status

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
       Case "<%=UID_M0001%>"                                                         '¢Ð : Save
          If Trim("<%=lgErrorStatus%>") = "NO" Then
              With Parent
                .ggoSpread.Source     = .frm1.vspdData
                .ggoSpread.SSShowData "<%=lgstrData%>"
                .DBAutoQueryOk
	         End with
          End If   
    End Select    
    
       
</Script>	
