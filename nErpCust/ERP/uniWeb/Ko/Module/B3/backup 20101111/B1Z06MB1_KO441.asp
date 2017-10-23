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
	Dim lgStrPrevKey
	Dim orgChangID
	Const C_SHEETMAXROWS_D = 100
    Dim lgSvrDateTime
    call LoadBasisGlobalInf()

Const C_IG_CMD = 0
Const C_IG_ITEM_CD = 1
Const C_IG_FLAG = 2
Const C_IG_ROW = 3
    
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
        Case CStr(UID_M0003)                                                         '☜: Delete
             Call SubBizDelete()
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

            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("MES_SND_YN"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("USR_NM"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("NOTE_DT"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("DEV_PROD_GB"))			
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ITEM_CD"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ITEM_NM"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("CBM_DESCRIPTION"))			
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("CDN_BIZ"))			
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("CDN_BMP"))			
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("CDN_PKG"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("CDN_PRD"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("CDN_TQC"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("CONFIRM_FLG"))
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
            Case "C"
                    Call SubBizSaveMultiCreate(arrColVal)                            '☜: Create
            Case "U"
                    Call SubBizSaveMultiUpdate(arrColVal)                            '☜: Update
            Case "D"
                    Call SubBizSaveMultiDelete(arrColVal)                            '☜: Delete
        End Select
        
        If lgErrorStatus    = "YES" Then
           lgErrorPos = lgErrorPos & arrColVal(C_IG_ROW) & gColSep
           Exit For
        End If
    Next
End Sub      

'============================================================================================================
' Name : SubBizSaveCreate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveMultiCreate(arrColVal)
    On Error Resume Next
    Err.Clear                                                                        '☜: Clear Error status
  
		
End Sub
'============================================================================================================
' Name : SubBizSaveMultiUpdate
' Desc : Update Data from Db
'============================================================================================================
Sub SubBizSaveMultiUpdate(arrColVal)
 
		lgStrSQL = "UPDATE B_CDN_REQ_HDR_KO441 SET "
		
		If Trim(arrColVal(C_IG_FLAG)) = "1" Then
			lgStrSQL = lgStrSQL & " MES_SND_YN='Y' ," 
		Else
			lgStrSQL = lgStrSQL & " MES_SND_YN='N' ," 
		End If		
		lgStrSQL = lgStrSQL & " UPDT_USER_ID	=" & FilterVar(gUsrId, "''", "S")                        & "," 
		lgStrSQL = lgStrSQL & " UPDT_DT				=" & FilterVar(lgSvrDateTime, "''", "S")   
		lgStrSQL = lgStrSQL & " WHERE ITEM_CD =" & FilterVar(UCase(arrColVal(C_IG_ITEM_CD)), "''", "S")

    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
	Call SubHandleError("MU",lgObjConn,lgObjRs,Err)
	
End Sub

'============================================================================================================
' Name : SubBizSaveMultiDelete
' Desc : Delete Data from Db

'============================================================================================================
Sub SubBizSaveMultiDelete(arrColVal)
     On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status	
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

                       lgStrSQL = lgStrSQL & " case when a.MES_SND_YN='Y' then 1 else 0 end MES_SND_YN, " & vbcrlf
                       lgStrSQL = lgStrSQL & " b.USR_NM, " & vbcrlf
                       lgStrSQL = lgStrSQL & " a.NOTE_DT, " & vbcrlf
                       lgStrSQL = lgStrSQL & " a.ITEM_CD, " & vbcrlf
                       lgStrSQL = lgStrSQL & " a.ITEM_NM, " & vbcrlf
                       lgStrSQL = lgStrSQL & " a.CBM_DESCRIPTION, " & vbcrlf
                       lgStrSQL = lgStrSQL & " case when a.CDN_BIZ='Y' then 1 else 0 end CDN_BIZ, " & vbcrlf
                       lgStrSQL = lgStrSQL & " case when a.CDN_BMP='Y' then 1 else 0 end CDN_BMP, " & vbcrlf
                       lgStrSQL = lgStrSQL & " case when a.CDN_PKG='Y' then 1 else 0 end CDN_PKG, " & vbcrlf
                       lgStrSQL = lgStrSQL & " case when a.CDN_PRD='Y' then 1 else 0 end CDN_PRD, " & vbcrlf
                       lgStrSQL = lgStrSQL & " case when a.CDN_TQC='Y' then 1 else 0 end CDN_TQC, " & vbcrlf
                       lgStrSQL = lgStrSQL & " case when a.CONFIRM_FLG='Y' then 1 else 0 end CONFIRM_FLG," & vbcrlf
                       lgStrSQL = lgStrSQL & " case when a.DEV_PROD_GB='Y' then 'Development' else 'Production' end DEV_PROD_GB " & vbcrlf

                       lgStrSQL = lgStrSQL & " from B_CDN_REQ_HDR_KO441 a (nolock) " & vbcrlf
                       lgStrSQL = lgStrSQL & " left outer join Z_USR_MAST_REC b (nolock) on (a.INSRT_USER_ID=b.USR_ID) " & vbcrlf
                       lgStrSQL = lgStrSQL & " where 1=1 "
                       
                       If Trim(Request("txtItemCd")) <> "" Then
                       		lgStrSQL = lgStrSQL & " AND a.ITEM_CD=" & FilterVar(Request("txtItemCd"),"''","S") & VbCrLf
                       End If
                       If Trim(Request("txtFrDt")) <> "" Then
                       		lgStrSQL = lgStrSQL & " AND a.NOTE_DT>=" & FilterVar(UniConvDate(Request("txtFrDt")),"''","S") & VbCrLf
                       End If
                       If Trim(Request("txtToDt")) <> "" Then
                       		lgStrSQL = lgStrSQL & " AND a.NOTE_DT<=" & FilterVar(UniConvDate(Request("txtToDt")),"''","S") & VbCrLf
                       End If
                       If Trim(Request("txtDevice")) <> "" Then
                       		lgStrSQL = lgStrSQL & " AND a.CBM_DESCRIPTION =" & FilterVar(Request("txtDevice"),"''","S") & VbCrLf
                       End If
                       If Trim(Request("rdoDp")) <> "" Then
                       		lgStrSQL = lgStrSQL & " AND a.DEV_PROD_GB =" & FilterVar(Request("rdoDp"),"''","S") & VbCrLf
                       End If
                       If Trim(Request("rdoTrans")) <> "" Then
                       		lgStrSQL = lgStrSQL & " AND a.MES_SND_YN =" & FilterVar(Request("rdoTrans"),"''","S") & VbCrLf
                       End If
                       If Trim(Request("rdoCfmDept")) <> "" Then
                       		If Trim(Request("rdoCfmDept"))="SO" Then
                       			If Trim(Request("rdoCfm"))="Y" Then
	                       			lgStrSQL = lgStrSQL & " AND a.CDN_BIZ ='Y'" & VbCrLf
	                       		else
	                       			lgStrSQL = lgStrSQL & " AND a.CDN_BIZ ='N'" & VbCrLf
	                       		end if
                       		ElseIf Trim(Request("rdoCfmDept"))="TB" Then
                       			If Trim(Request("rdoCfm"))="Y" Then
		                       		lgStrSQL = lgStrSQL & " AND a.CDN_BMP ='Y'" & VbCrLf
	                       		else
		                       		lgStrSQL = lgStrSQL & " AND a.CDN_BMP ='N'" & VbCrLf
	                       		end if
                       		ElseIf Trim(Request("rdoCfmDept"))="TP" Then
                       			If Trim(Request("rdoCfm"))="Y" Then
		                       		lgStrSQL = lgStrSQL & " AND a.CDN_PKG ='Y'" & VbCrLf
	                       		else
		                       		lgStrSQL = lgStrSQL & " AND a.CDN_PKG ='N'" & VbCrLf
	                       		end if
                       		ElseIf Trim(Request("rdoCfmDept"))="QA" Then
                       			If Trim(Request("rdoCfm"))="Y" Then
		                       		lgStrSQL = lgStrSQL & " AND a.CDN_PRD ='Y'" & VbCrLf
	                       		else
		                       		lgStrSQL = lgStrSQL & " AND a.CDN_PRD ='N'" & VbCrLf
	                       		end if
                       		ElseIf Trim(Request("rdoCfmDept"))="PP" Then
                       			If Trim(Request("rdoCfm"))="Y" Then
		                       		lgStrSQL = lgStrSQL & " AND a.CDN_TQC ='Y'" & VbCrLf
	                       		else
		                       		lgStrSQL = lgStrSQL & " AND a.CDN_TQC ='N'" & VbCrLf
	                       		end if
                       		End If
                       End If
                       If Trim(Request("rdoCfm")) <> "" Then
                       		If Trim(Request("rdoCfm"))="Y" Then
	                       		lgStrSQL = lgStrSQL & " AND a.CONFIRM_FLG ='Y'" & VbCrLf
                       		Else
	                       		lgStrSQL = lgStrSQL & " AND a.CONFIRM_FLG ='N'" & VbCrLf
                       		End IF
                       End If

											 lgStrSQL = lgStrSQL & " order by a.ITEM_CD"
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
                .lgStrPrevKey    = "<%=lgStrPrevKey%>"
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
       Case "<%=UID_M0003%>"                                                         '☜ : Delete
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DbDeleteOk
          Else   
          End If   
    End Select       
       
</Script>	
