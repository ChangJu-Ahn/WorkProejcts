<%@ LANGUAGE=VBSCript TRANSACTION=Required%>
<% Option Explicit%>

<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/lgSvrVariables.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->

<%
	Dim lgStrPrevKey
	Const C_SHEETMAXROWS_D = 100
    Call HideStatusWnd                                                               '¢Ð: Hide Processing message
    Call LoadBasisGlobalInf()  
    Call LoadInfTB19029B("I", "H","NOCOOKIE","MB")

    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '¢Ð: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '¢Ð: Read Operation Mode (CRUD)
    lgKeyStream       = Split(Request("txtKeyStream"),gColSep)

    lgLngMaxRow       = Request("txtMaxRows")                                        '¢Ð: Read Operation Mode (CRUD)
    lgStrPrevKey = UNICInt(Trim(Request("lgStrPrevKey")),0)							 '¢Ð: "0"(First),"1"(Second),"2"(Third),"3"(...)

    Call SubOpenDB(lgObjConn)                                                        '¢Ð: Make a DB Connection
    
    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '¢Ð: Query
             Call SubBizQuery()
        Case CStr(UID_M0002)                                                         '¢Ð: Save,Update
             Call SubBizSaveMulti()
        Case CStr(UID_M0003)                                                         '¢Ð: Delete
             Call SubBizDelete()
    End Select
    
    Call SubCloseDB(lgObjConn)                                                       '¢Ð: Close DB Connection

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()
    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status
    Call SubBizQueryCond()
    If lgErrorStatus <> "YES" Then
		Call SubBizQueryMulti()
	End If
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
' Name : SubBizQueryCond
' Desc : Verify Condition Value from Db
'============================================================================================================
Sub SubBizQueryCond()

    On Error Resume Next
    Err.Clear

	If lgKeyStream(0) <> ""   Then
   
		Call SubMakeSQLStatements("CB")
    
		If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
			Response.Write "<Script Language=VBScript>" & vbCrLf
				Response.Write "parent.frm1.txtBp_nmFrom.value = """"" & vbCrLf	
			Response.Write "</Script>" & vbCrLf
		    'Call DisplayMsgBox("125000", vbInformation, "", "", I_MKSCRIPT)
		    'Call SetErrorStatus()
		Else
			Response.Write "<Script Language=VBScript>" & vbCrLf
				Response.Write "parent.frm1.txtBp_NMFrom.value = """ & ConvSPChars(lgObjRs("BP_NM")) & """" & vbCrLf
			Response.Write "</Script>" & vbCrLf
		End If
	Else
		Response.Write "<Script Language=VBScript>" & vbCrLf
			Response.Write "parent.frm1.txtBp_nmFrom.value = """"" & vbCrLf	
		Response.Write "</Script>" & vbCrLf
	End If

End Sub

'============================================================================================================
' Name : SubBizQueryMulti
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()

    Dim iDx
    Dim iLoopMax
    Dim iKey1
    Dim strWhere
    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status

     strWhere = FilterVar(Trim(lgKeyStream(0)),"''","S")
    
    Call SubMakeSQLStatements("MR")                                 '¡Ù : Make sql statements
    
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        lgStrPrevKey = ""
        Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '¢Ð : No data is found. 
        Call SetErrorStatus()
    Else
        Call SubSkipRs(lgObjRs,C_SHEETMAXROWS_D * lgStrPrevKey)
        lgstrData = ""
        iDx       = 1
        
        Do While Not lgObjRs.EOF
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("BP_CD"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("BP_NM"))
            
            If ConvSPChars(lgObjRs("SCM_PUR_FLAG")) = "Y" Then
				lgstrData = lgstrData & Chr(11) & "1"
			ElseIf ConvSPChars(lgObjRs("SCM_PUR_FLAG")) = "N" Then
				lgstrData = lgstrData & Chr(11) & ""
			Else
				lgstrData = lgstrData & Chr(11) & ""	
            End If
            
            If ConvSPChars(lgObjRs("SCM_RET_FLAG")) = "Y" Then
				lgstrData = lgstrData & Chr(11) & "1"
			ElseIf ConvSPChars(lgObjRs("SCM_RET_FLAG")) = "N" Then
				lgstrData = lgstrData & Chr(11) & ""
			Else
				lgstrData = lgstrData & Chr(11) & ""	
            End If
            
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("BP_TYPE"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("BP_TYPE_NM"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("BP_RGST_NO"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("REPRE_NM"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("IND_TYPE"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("IND_TYPE_NM"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("IND_CLASS"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("IND_CLASS_NM"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("TEL_NO1"))
            
            If ConvSPChars(lgObjRs("USAGE_FLAG")) = "Y" Then
				lgstrData = lgstrData & Chr(11) & "1"
			ElseIf ConvSPChars(lgObjRs("USAGE_FLAG")) = "N" Then
				lgstrData = lgstrData & Chr(11) & ""
			Else
				lgstrData = lgstrData & Chr(11) & ""	
            End If
            
            
            
            lgstrData = lgstrData & Chr(11) & lgLngMaxRow + iDx
            lgstrData = lgstrData & Chr(11) & Chr(12)
		    lgObjRs.MoveNext
            iDx =  iDx + 1
            If iDx > C_SHEETMAXROWS_D Then
               lgStrPrevKey = lgStrPrevKey + 1
               Exit Do
            End If   
        Loop 
    End If
    
    If iDx <= C_SHEETMAXROWS_D Then
       lgStrPrevKey = ""
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
        CALL SVRMSGBOX(arrColVal ,0,1)
        
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
    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status
End Sub
'============================================================================================================
' Name : SubBizSaveMultiUpdate
' Desc : Update Data from Db
'============================================================================================================
Sub SubBizSaveMultiUpdate(arrColVal)
    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status

    lgStrSQL =            " UPDATE  B_BIZ_PARTNER "
    lgStrSQL = lgStrSQL & "    SET " 
    lgStrSQL = lgStrSQL & " SCM_PUR_FLAG	= " &  FilterVar(Trim(UCase(arrColVal(3))),"","S") & "," 
    lgStrSQL = lgStrSQL & " SCM_RET_FLAG	= " &  FilterVar(Trim(UCase(arrColVal(4))),"","S") & "," 
    lgStrSQL = lgStrSQL & " UPDT_USER_ID	= " &  FilterVar(gUsrId,"","S")                    & "," 
    lgStrSQL = lgStrSQL & " UPDT_DT			= " &  FilterVar(GetSvrDateTime,"''","S")
    lgStrSQL = lgStrSQL & " WHERE BP_CD		= " &  FilterVar(Trim(UCase(arrColVal(2))),"","S")
        
'    Response.Write lgStrSQL
        
    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	
	Call SubHandleError("MU",lgObjConn,lgObjRs,Err)    
    
End Sub

'============================================================================================================
' Name : SubBizSaveMultiDelete
' Desc : Delete Data from Db
'============================================================================================================
Sub SubBizSaveMultiDelete(arrColVal)
    On Error Resume Next
    Err.Clear
End Sub

'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements(pDataType)
    Dim iSelCount

    Select Case Mid(pDataType,1,1)
        Case "M"
           iSelCount = C_SHEETMAXROWS_D + C_SHEETMAXROWS_D *  lgStrPrevKey + 1
           
           Select Case Mid(pDataType,2,1)
               Case "R"
					lgStrSQL =                   "	SELECT	TOP " & iSelCount  & " A.* , (B.MINOR_NM)IND_TYPE_NM , (C.MINOR_NM) IND_CLASS_NM , (D.MINOR_NM) BP_TYPE_NM "
					lgStrSQL = lgStrSQL & "			  FROM	B_BIZ_PARTNER A, B_MINOR B , B_MINOR C , B_MINOR D "
					lgStrSQL = lgStrSQL & "			 WHERE 	IND_TYPE = B.MINOR_CD"
					lgStrSQL = lgStrSQL & "			   AND  IND_CLASS = C.MINOR_CD"
					lgStrSQL = lgStrSQL & "			   AND  BP_TYPE = D.MINOR_CD"
					lgStrSQL = lgStrSQL & "			   AND  B.MAJOR_CD = 'B9002' "
					lgStrSQL = lgStrSQL & "			   AND  C.MAJOR_CD = 'B9003' "
					lgStrSQL = lgStrSQL & "			   AND  D.MAJOR_CD = 'B9005' "	
					lgStrSQL = lgStrSQL & "			   AND  BP_TYPE IN ('S','CS') "					
					
					If lgkeystream(0) <> "" Then
						lgStrSQL = lgStrSQL & "             and 	a.BP_CD >= " & FilterVar(lgKeyStream(0) ,"''", "S")
					End If
					
					If lgkeystream(1) <> "" Then
						lgStrSQL = lgStrSQL & "            	and 	a.USAGE_FLAG = " & FilterVar(lgKeyStream(1),"''", "S")
					End If
					
					If lgkeystream(2) <> "" Then
						lgStrSQL = lgStrSQL & "            	and 	a.BP_NM LIKE " & FilterVar("%" & lgKeyStream(2) & "%","''", "S")
					End If
					
					
					
					lgStrSQL = lgStrSQL & "			  order BY A.BP_CD  "
           End Select     
           
        Case "C"
           
           Select Case Mid(pDataType,2,1)
               Case "B"
                    lgStrSQL =            " select bp_nm from b_biz_partner where bp_cd = " & FilterVar(lgKeyStream(0) & "" ,"''", "S")
               Case "C"
                    lgStrSQL =            " select bp_nm from b_biz_partner where bp_cd = " & FilterVar(lgKeyStream(1) & "" ,"''", "S")     
           End Select  
                   
    End Select
    
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
	Dim lsMsg
    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status

    Select Case pOpCode
        Case "MC"
        Case "MD"
        Case "MR"
        Case "MU"
    End Select
End Sub

%>

<Script Language="VBScript">
    
    Select Case "<%=lgOpModeCRUD %>"
       Case "<%=UID_M0001%>"                                                         '¢Ð : Query
          If Trim("<%=lgErrorStatus%>") = "NO" Then
              With Parent
                .ggoSpread.Source     = .frm1.vspdData
                .lgStrPrevKey    = "<%=lgStrPrevKey%>"
                .ggoSpread.SSShowData "<%=lgstrData%>"
                .DBQueryOk        
	         End with
          End If   
       Case "<%=UID_M0002%>"                                                         '¢Ð : Save
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DBSaveOk
          Else
             Parent.SubSetErrPos(Trim("<%=lgErrorPos%>"))
          End If   
       Case "<%=UID_M0002%>"                                                         '¢Ð : Delete
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DbDeleteOk
          Else   
          End If   
    End Select    
      
</Script>	