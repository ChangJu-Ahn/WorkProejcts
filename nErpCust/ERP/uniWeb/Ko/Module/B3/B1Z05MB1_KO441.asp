<%@ LANGUAGE="VBScript" CODEPAGE=949 TRANSACTION=Required %>
<% session.CodePage=949 %>

<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/IncServerAdoDb.asp" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->

<%

Const C_IG_CMD = 0
Const C_IG_SEQ = 1
Const C_IG_UD_MINOR_CD = 2
Const C_IG_SAMPLE_DATA = 3
Const C_IG_ROW = 4

	Dim lgStrPrevKey
	Dim lgCBM_DESCRIPTION
	Dim lgUSR_NM
	Dim lgDEV_PROD_GB
	Dim lgNOTE_DT
	

	Const C_SHEETMAXROWS_D = 100
    Dim lgSvrDateTime
    call LoadBasisGlobalInf()
    
    lgSvrDateTime = GetSvrDateTime    
    
	Call loadInfTB19029B( "I", "H","NOCOOKIE","MB")   
  
    Call HideStatusWnd 

    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '☜: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)
    lgKeyStream       = Split(Request("txtKeyStream"),gColSep)

    lgLngMaxRow       = Request("txtMaxRows")                                        '☜: Read Operation Mode (CRUD)
    lgStrPrevKey = UNICInt(Trim(Request("lgStrPrevKey")),0)                '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)
    
    Call SubOpenDB(lgObjConn)           
   
    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '☜: Query
             Call SubBizQuery()
        Case CStr(UID_M0002)                                                         '☜: Save,Update
             Call SubBizSaveMulti()
        Case CStr("CONFIRM")                                                         '☜: Delete
             Call SubBizConfirm()
        Case CStr("UNCONFIRM")                                                         '☜: Delete
             Call SubBizUnConfirm()
    End Select
'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()
     On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear
    Call SubBizQueryMulti()
End Sub    
'============================================================================================================
' Name : SubBizDelete
' Desc : Delete DB data
'============================================================================================================
Sub SubBizDelete()
     On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear            
End Sub

'============================================================================================================
' Name : SubBizQueryMulti
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()
    Dim iDx
    Dim iLoopMax
    Dim iKey1
    Dim strWhere, strFg
  
     On Error Resume Next    
    Err.Clear                                                               '☜: Clear Error status

	
    Call SubMakeSQLStatements("MR",strWhere,"X",C_LIKE)                              '☜ : Make sql statements

    If FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL&strWhere,"X","X") = False Then
       lgStrPrevKey = ""
       Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
       Call SetErrorStatus()
    Else
       Call SubSkipRs(lgObjRs,C_SHEETMAXROWS_D * lgStrPrevKey)
       lgstrData = ""
       iDx       = 1
       Do While Not lgObjRs.EOF

						lgCBM_DESCRIPTION = ConvSPChars(lgObjRs("CBM_DESCRIPTION"))
						lgUSR_NM 					= ConvSPChars(lgObjRs("USR_NM"))
						lgDEV_PROD_GB			= ConvSPChars(lgObjRs("DEV_PROD_GB"))
						lgNOTE_DT 				= ConvSPChars(lgObjRs("NOTE_DT"))


            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("SEQ"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("TODO_DOC"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("COMBO_YN"))            
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("UD_MAJOR_CD"))			
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("UD_MINOR_CD"))			
            lgstrData = lgstrData & Chr(11)
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("DATA_TEXT"))			
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("PROCESS_TYPE"))			
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("MES_USE_YN"))			
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("CDN_BIZ"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("CDN_BMP"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("CDN_PKG"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("CDN_PRD"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("CDN_TQC"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("REMARK"))
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

      If CheckSQLError(lgObjRs.ActiveConnection) = True Then
         ObjectContext.SetAbort
      End If
            
      Call SubHandleError("MR",lgObjConn,lgObjRs,Err)
      Call SubCloseRs(lgObjRs)    

End Sub    

'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data 
'============================================================================================================
Sub SubBizSaveMulti()
    Dim arrRowVal
    Dim arrColVal
    Dim iDx

    On Error Resume Next
    Err.Clear                                                                        '☜: Clear Error status
    
	arrRowVal = Split(Request("txtSpread"), gRowSep)                                 '☜: Split Row    data
	
    For iDx = 1 To lgLngMaxRow
        arrColVal = Split(arrRowVal(iDx-1), gColSep)                                 '☜: Split Column data
        
        Select Case arrColVal(C_IG_CMD)
            Case "U"
                    Call SubBizSaveMultiUpdate(arrColVal)                            '☜: Update
        End Select
        
        If lgErrorStatus    = "YES" Then
           lgErrorPos = lgErrorPos & arrColVal(C_IG_ROW) & gColSep
           Exit For
        End If
    Next
End Sub      

'============================================================================================================
' Name : SubBizSaveMultiUpdate
' Desc : Update Data from Db
'============================================================================================================
Sub SubBizSaveMultiUpdate(arrColVal)

		lgStrSQL = "UPDATE B_CDN_REQ_DTL_KO441 SET "
		lgStrSQL = lgStrSQL & " UD_MINOR_CD		=" & FilterVar(Trim(arrColVal(C_IG_UD_MINOR_CD)), "''", "S")     & ","
		lgStrSQL = lgStrSQL & " DATA_TEXT 		=" & FilterVar(Trim(arrColVal(C_IG_SAMPLE_DATA)), "''", "S")     & ","
		lgStrSQL = lgStrSQL & " UPDT_USER_ID	=" & FilterVar(gUsrId, "''", "S")                        & "," 
		lgStrSQL = lgStrSQL & " UPDT_DT				=" & FilterVar(lgSvrDateTime, "''", "S")   
		lgStrSQL = lgStrSQL & " WHERE ITEM_CD	=" & FilterVar(Request("txtItemCd"), "''", "S")
		lgStrSQL = lgStrSQL & " AND 	SEQ			=" & FilterVar(UCase(arrColVal(C_IG_SEQ)), "''", "S")

    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
	Call SubHandleError("MU",lgObjConn,lgObjRs,Err)
	
End Sub


'============================================================================================================
' Name : SubBizSaveCreate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizConfirm()
    On Error Resume Next
    Err.Clear                                                                        '☜: Clear Error status
  	Dim iStrSeq

		lgStrSQL = " If Exists(select CDN_BIZ " & vbcrlf
		lgStrSQL = lgStrSQL & " from B_CDN_REQ_DTL_KO441 " & vbcrlf
		lgStrSQL = lgStrSQL & " where exists(select CDN_BIZ from B_CDN_USER_KO441 where USER_ID=" & FilterVar(gUsrId,"''","S") & " and CDN_BIZ='Y' and CDN_BIZ=B_CDN_REQ_DTL_KO441.CDN_BIZ)) " & vbcrlf
		lgStrSQL = lgStrSQL & " update B_CDN_REQ_HDR_KO441 set " & vbcrlf
		lgStrSQL = lgStrSQL & " CDN_BIZ='Y' " & vbcrlf
		lgStrSQL = lgStrSQL & " where ITEM_CD=" & FilterVar(Request("txtItemCd"),"''","S") & vbcrlf
		lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
		Call SubHandleError("MC",lgObjConn,lgObjRs,Err)


		lgStrSQL = " If Exists(select CDN_BMP " & vbcrlf
		lgStrSQL = lgStrSQL & " from B_CDN_REQ_DTL_KO441 " & vbcrlf
		lgStrSQL = lgStrSQL & " where exists(select CDN_BIZ from B_CDN_USER_KO441 where USER_ID=" & FilterVar(gUsrId,"''","S") & " and CDN_BMP='Y' and CDN_BMP=B_CDN_REQ_DTL_KO441.CDN_BMP)) " & vbcrlf
		lgStrSQL = lgStrSQL & " update B_CDN_REQ_HDR_KO441 set " & vbcrlf
		lgStrSQL = lgStrSQL & " CDN_BMP='Y' " & vbcrlf
		lgStrSQL = lgStrSQL & " where ITEM_CD=" & FilterVar(Request("txtItemCd"),"''","S") & vbcrlf
		lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
		Call SubHandleError("MC",lgObjConn,lgObjRs,Err)

		lgStrSQL = " If Exists(select CDN_PKG " & vbcrlf
		lgStrSQL = lgStrSQL & " from B_CDN_REQ_DTL_KO441 " & vbcrlf
		lgStrSQL = lgStrSQL & " where exists(select CDN_PKG from B_CDN_USER_KO441 where USER_ID=" & FilterVar(gUsrId,"''","S") & " and CDN_PKG='Y' and CDN_PKG=B_CDN_REQ_DTL_KO441.CDN_PKG)) " & vbcrlf
		lgStrSQL = lgStrSQL & " update B_CDN_REQ_HDR_KO441 set " & vbcrlf
		lgStrSQL = lgStrSQL & " CDN_PKG='Y' " & vbcrlf
		lgStrSQL = lgStrSQL & " where ITEM_CD=" & FilterVar(Request("txtItemCd"),"''","S") & vbcrlf
		lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
		Call SubHandleError("MC",lgObjConn,lgObjRs,Err)

		lgStrSQL = " If Exists(select CDN_PRD " & vbcrlf
		lgStrSQL = lgStrSQL & " from B_CDN_REQ_DTL_KO441 " & vbcrlf
		lgStrSQL = lgStrSQL & " where exists(select CDN_PRD from B_CDN_USER_KO441 where USER_ID=" & FilterVar(gUsrId,"''","S") & " and CDN_PRD='Y' and CDN_PRD=B_CDN_REQ_DTL_KO441.CDN_PRD)) " & vbcrlf
		lgStrSQL = lgStrSQL & " update B_CDN_REQ_HDR_KO441 set " & vbcrlf
		lgStrSQL = lgStrSQL & " CDN_PRD='Y' " & vbcrlf
		lgStrSQL = lgStrSQL & " where ITEM_CD=" & FilterVar(Request("txtItemCd"),"''","S") & vbcrlf
		lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
		Call SubHandleError("MC",lgObjConn,lgObjRs,Err)

		lgStrSQL = " If Exists(select CDN_TQC " & vbcrlf
		lgStrSQL = lgStrSQL & " from B_CDN_REQ_DTL_KO441 " & vbcrlf
		lgStrSQL = lgStrSQL & " where exists(select CDN_TQC from B_CDN_USER_KO441 where USER_ID=" & FilterVar(gUsrId,"''","S") & " and CDN_TQC='Y' and CDN_TQC=B_CDN_REQ_DTL_KO441.CDN_TQC)) " & vbcrlf
		lgStrSQL = lgStrSQL & " update B_CDN_REQ_HDR_KO441 set " & vbcrlf
		lgStrSQL = lgStrSQL & " CDN_TQC='Y' " & vbcrlf
		lgStrSQL = lgStrSQL & " where ITEM_CD=" & FilterVar(Request("txtItemCd"),"''","S") & vbcrlf
		lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
		Call SubHandleError("MC",lgObjConn,lgObjRs,Err)

		lgStrSQL = " If exists(select CDN_BIZ " & vbcrlf
		lgStrSQL = lgStrSQL & " from B_CDN_REQ_HDR_KO441 " & vbcrlf
		lgStrSQL = lgStrSQL & " where  " & vbcrlf
		lgStrSQL = lgStrSQL & " CDN_BIZ = isnull((select distinct CDN_BIZ from B_CDN_REQ_DTL_KO441 where item_cd=" & FilterVar(Request("txtItemCd"),"''","S") & " and CDN_BIZ='Y'),'N') AND  " & vbcrlf
		lgStrSQL = lgStrSQL & " CDN_BMP = isnull((select distinct CDN_BMP from B_CDN_REQ_DTL_KO441 where item_cd=" & FilterVar(Request("txtItemCd"),"''","S") & " and CDN_BMP='Y'),'N') AND " & vbcrlf
		lgStrSQL = lgStrSQL & " CDN_PKG = isnull((select distinct CDN_PKG from B_CDN_REQ_DTL_KO441 where item_cd=" & FilterVar(Request("txtItemCd"),"''","S") & " and CDN_PKG='Y'),'N') AND " & vbcrlf
		lgStrSQL = lgStrSQL & " CDN_PRD = isnull((select distinct CDN_PRD from B_CDN_REQ_DTL_KO441 where item_cd=" & FilterVar(Request("txtItemCd"),"''","S") & " and CDN_PRD='Y'),'N') AND " & vbcrlf
		lgStrSQL = lgStrSQL & " CDN_TQC = isnull((select distinct CDN_TQC from B_CDN_REQ_DTL_KO441 where item_cd=" & FilterVar(Request("txtItemCd"),"''","S") & " and CDN_TQC='Y'),'N')) " & vbcrlf
		lgStrSQL = lgStrSQL & " update B_CDN_REQ_HDR_KO441 set " & vbcrlf
		lgStrSQL = lgStrSQL & " CONFIRM_FLG='Y' " & vbcrlf
		lgStrSQL = lgStrSQL & " where ITEM_CD=" & FilterVar(Request("txtItemCd"),"''","S") & vbcrlf
		lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
		Call SubHandleError("MC",lgObjConn,lgObjRs,Err)
  			
End Sub

Sub SubBizUnConfirm()
    On Error Resume Next
    Err.Clear                                                                        '☜: Clear Error status
  	Dim iStrSeq

		lgStrSQL = " If Exists(select CDN_BIZ " & vbcrlf
		lgStrSQL = lgStrSQL & " from B_CDN_REQ_DTL_KO441 " & vbcrlf
		lgStrSQL = lgStrSQL & " where exists(select CDN_BIZ from B_CDN_USER_KO441 where USER_ID=" & FilterVar(gUsrId,"''","S") & " and CDN_BIZ='Y' and CDN_BIZ=B_CDN_REQ_DTL_KO441.CDN_BIZ)) " & vbcrlf
		lgStrSQL = lgStrSQL & " update B_CDN_REQ_HDR_KO441 set " & vbcrlf
		lgStrSQL = lgStrSQL & " CDN_BIZ='N' " & vbcrlf
		lgStrSQL = lgStrSQL & " where ITEM_CD=" & FilterVar(Request("txtItemCd"),"''","S") & vbcrlf
		lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
		Call SubHandleError("MC",lgObjConn,lgObjRs,Err)


		lgStrSQL = " If Exists(select CDN_BMP " & vbcrlf
		lgStrSQL = lgStrSQL & " from B_CDN_REQ_DTL_KO441 " & vbcrlf
		lgStrSQL = lgStrSQL & " where exists(select CDN_BIZ from B_CDN_USER_KO441 where USER_ID=" & FilterVar(gUsrId,"''","S") & " and CDN_BMP='Y' and CDN_BMP=B_CDN_REQ_DTL_KO441.CDN_BMP)) " & vbcrlf
		lgStrSQL = lgStrSQL & " update B_CDN_REQ_HDR_KO441 set " & vbcrlf
		lgStrSQL = lgStrSQL & " CDN_BMP='N' " & vbcrlf
		lgStrSQL = lgStrSQL & " where ITEM_CD=" & FilterVar(Request("txtItemCd"),"''","S") & vbcrlf
		lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
		Call SubHandleError("MC",lgObjConn,lgObjRs,Err)

		lgStrSQL = " If Exists(select CDN_PKG " & vbcrlf
		lgStrSQL = lgStrSQL & " from B_CDN_REQ_DTL_KO441 " & vbcrlf
		lgStrSQL = lgStrSQL & " where exists(select CDN_PKG from B_CDN_USER_KO441 where USER_ID=" & FilterVar(gUsrId,"''","S") & " and CDN_PKG='Y' and CDN_PKG=B_CDN_REQ_DTL_KO441.CDN_PKG)) " & vbcrlf
		lgStrSQL = lgStrSQL & " update B_CDN_REQ_HDR_KO441 set " & vbcrlf
		lgStrSQL = lgStrSQL & " CDN_PKG='N' " & vbcrlf
		lgStrSQL = lgStrSQL & " where ITEM_CD=" & FilterVar(Request("txtItemCd"),"''","S") & vbcrlf
		lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
		Call SubHandleError("MC",lgObjConn,lgObjRs,Err)

		lgStrSQL = " If Exists(select CDN_PRD " & vbcrlf
		lgStrSQL = lgStrSQL & " from B_CDN_REQ_DTL_KO441 " & vbcrlf
		lgStrSQL = lgStrSQL & " where exists(select CDN_PRD from B_CDN_USER_KO441 where USER_ID=" & FilterVar(gUsrId,"''","S") & " and CDN_PRD='Y' and CDN_PRD=B_CDN_REQ_DTL_KO441.CDN_PRD)) " & vbcrlf
		lgStrSQL = lgStrSQL & " update B_CDN_REQ_HDR_KO441 set " & vbcrlf
		lgStrSQL = lgStrSQL & " CDN_PRD='N' " & vbcrlf
		lgStrSQL = lgStrSQL & " where ITEM_CD=" & FilterVar(Request("txtItemCd"),"''","S") & vbcrlf
		lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
		Call SubHandleError("MC",lgObjConn,lgObjRs,Err)

		lgStrSQL = " If Exists(select CDN_TQC " & vbcrlf
		lgStrSQL = lgStrSQL & " from B_CDN_REQ_DTL_KO441 " & vbcrlf
		lgStrSQL = lgStrSQL & " where exists(select CDN_TQC from B_CDN_USER_KO441 where USER_ID=" & FilterVar(gUsrId,"''","S") & " and CDN_TQC='Y' and CDN_TQC=B_CDN_REQ_DTL_KO441.CDN_TQC)) " & vbcrlf
		lgStrSQL = lgStrSQL & " update B_CDN_REQ_HDR_KO441 set " & vbcrlf
		lgStrSQL = lgStrSQL & " CDN_TQC='N' " & vbcrlf
		lgStrSQL = lgStrSQL & " where ITEM_CD=" & FilterVar(Request("txtItemCd"),"''","S") & vbcrlf
		lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
		Call SubHandleError("MC",lgObjConn,lgObjRs,Err)

		lgStrSQL = " update B_CDN_REQ_HDR_KO441 set " & vbcrlf
		lgStrSQL = lgStrSQL & " CONFIRM_FLG='N' " & vbcrlf
		lgStrSQL = lgStrSQL & " where ITEM_CD=" & FilterVar(Request("txtItemCd"),"''","S") & vbcrlf
		lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
		Call SubHandleError("MC",lgObjConn,lgObjRs,Err)
  			
End Sub

'============================================================================================================
' Name : SubBizSaveMultiDelete
' Desc : Delete Data from Db

'============================================================================================================
Sub SubBizSaveMultiDelete(arrColVal)
 
End Sub

'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements(pDataType,pCode,pCode1,pComp)
     Dim iSelCount

    Select Case Mid(pDataType,1,1)
        Case "M"
      
           iSelCount = C_SHEETMAXROWS_D + C_SHEETMAXROWS_D *  lgStrPrevKey + 1
           
           Select Case Mid(pDataType,2,1)
           
               Case "R"
                              
                       lgStrSQL = "Select TOP " & iSelCount  

                       lgStrSQL = lgStrSQL & " a.SEQ,   " & vbCrLf
                       lgStrSQL = lgStrSQL & " b.TODO_DOC,   " & vbCrLf
                       lgStrSQL = lgStrSQL & " case when b.COMBO_YN='Y' then 1 else 0 end COMBO_YN,  " & vbCrLf
                       lgStrSQL = lgStrSQL & " b.UD_MAJOR_CD,   " & vbCrLf
                       lgStrSQL = lgStrSQL & " a.UD_MINOR_CD,   " & vbCrLf
                       lgStrSQL = lgStrSQL & " a.DATA_TEXT,   " & vbCrLf
                       lgStrSQL = lgStrSQL & " b.PROCESS_TYPE,   " & vbCrLf
                       lgStrSQL = lgStrSQL & " case when b.MES_USE_YN='Y' then 1 else 0 end MES_USE_YN,   " & vbCrLf
                       lgStrSQL = lgStrSQL & " case when a.CDN_BIZ='Y' then 1 else 0 end CDN_BIZ,   " & vbCrLf
                       lgStrSQL = lgStrSQL & " case when a.CDN_BMP='Y' then 1 else 0 end CDN_BMP,   " & vbCrLf
                       lgStrSQL = lgStrSQL & " case when a.CDN_PKG='Y' then 1 else 0 end CDN_PKG,   " & vbCrLf
                       lgStrSQL = lgStrSQL & " case when a.CDN_PRD='Y' then 1 else 0 end CDN_PRD,   " & vbCrLf
                       lgStrSQL = lgStrSQL & " case when a.CDN_TQC='Y' then 1 else 0 end CDN_TQC,   " & vbCrLf
                       lgStrSQL = lgStrSQL & " a.REMARK,   " & vbCrLf
                       lgStrSQL = lgStrSQL & " c.CBM_DESCRIPTION,   " & vbCrLf
                       lgStrSQL = lgStrSQL & " d.USR_NM,   " & vbCrLf
                       lgStrSQL = lgStrSQL & " c.DEV_PROD_GB,   " & vbCrLf
                       lgStrSQL = lgStrSQL & " c.NOTE_DT   " & vbCrLf
                       
                       lgStrSQL = lgStrSQL & " from B_CDN_REQ_DTL_KO441 a (nolock)   " & vbCrLf
                       lgStrSQL = lgStrSQL & " left outer join B_CDN_DEPT_TODO_KO441 b (nolock) on (a.SEQ=b.SEQ)   " & vbCrLf
                       lgStrSQL = lgStrSQL & " left outer join B_CDN_REQ_HDR_KO441 c (nolock) on (a.ITEM_CD=c.ITEM_CD)    " & vbCrLf
                       lgStrSQL = lgStrSQL & " left outer join Z_USR_MAST_REC d (nolock) on (c.INSRT_USER_ID=d.USR_ID)   " & vbCrLf
                       lgStrSQL = lgStrSQL & " where a.ITEM_CD=" & FilterVar(Request("txtItemCd"),"''","S") & " and a.SEQ>=0  " & vbCrLf
											lgStrSQL = lgStrSQL & "  order by 1 "
          End Select
    End Select
End Sub
'============================================================================================================
' Name : CommonOnTransactionCommit
' Desc : This Sub is called by OnTransactionCommit Error handler
'============================================================================================================
Sub CommonOnTransactionCommit()
    Response.Write "<BR> Commit Event occur"
End Sub
'============================================================================================================
' Name : CommonOnTransactionAbort
' Desc : This Sub is called by OnTransactionAbort Error handler
'============================================================================================================
Sub CommonOnTransactionAbort()
    lgErrorStatus    = "YES"
    Response.Write "<BR> Abort Event occur"
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
        Case "MS"
                 Call DisplayMsgBox("800486", vbInformation, "", "", I_MKSCRIPT)     'Can not create(Demo code)        
                 ObjectContext.SetAbort
                 Call SetErrorStatus
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
        Case "M1"
                 Call DisplayMsgBox("173132", vbInformation, "코드그룹", "", I_MKSCRIPT)     '
                 ObjectContext.SetAbort
                 Call SetErrorStatus
        Case "M2"
                 Call DisplayMsgBox("173132", vbInformation, "공통코드", "", I_MKSCRIPT)     '
                 ObjectContext.SetAbort
                 Call SetErrorStatus
    End Select
End Sub

%>

<Script Language="VBScript">
    Select Case "<%=lgOpModeCRUD %>"
       Case "<%=UID_M0001%>"                                                         '☜ : Query
          If Trim("<%=lgErrorStatus%>") = "NO" Then
              With Parent
                .ggoSpread.Source     = .frm1.vspdData

                .frm1.txtCBMdescription.value    = "<%=lgCBM_DESCRIPTION%>"
                .frm1.txtInsUser.value    = "<%=lgUSR_NM%>"
                If "<%=lgDEV_PROD_GB%>" = "Y" Then
                	.frm1.rdoDP1.checked=true
                Else
                	.frm1.rdoDP2.checked=true
              	End If
                .frm1.txtNoteDt.value    = "<%=lgNOTE_DT%>"

                .lgStrPrevKey    = "<%=lgStrPrevKey%>"
                .ggoSpread.SSShowData "<%=lgstrData%>"
                .DBQueryOk        
	         End with
          End If   
       Case "UNCONFIRM","CONFIRM","<%=UID_M0002%>"                                                         '☜ : Save
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DBSaveOk
          Else
             Parent.SubSetErrPos(Trim("<%=lgErrorPos%>"))
          End If   
       Case "<%=UID_M0003%>"                                                         '☜ : Delete
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DbDeleteOk
          Else   
          End If   
    End Select       
       
</Script>	
