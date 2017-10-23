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
Const C_IG_USER_ID = 1
Const C_IG_CDN_RCV = 2
Const C_IG_CDN_BIZ = 3
Const C_IG_CDN_BMP = 4
Const C_IG_CDN_PKG = 5
Const C_IG_CDN_PRD = 6
Const C_IG_CDN_TQC = 7
Const C_IG_REMARK = 8
Const C_IG_ROW = 9

	Dim lgStrPrevKey
	Dim orgChangID
	Const C_SHEETMAXROWS_D = 100
    Dim lgSvrDateTime
    call LoadBasisGlobalInf()
    
    lgSvrDateTime = GetSvrDateTime    
    
	Call loadInfTB19029B( "I", "H","NOCOOKIE","MB")   
  
    Call HideStatusWnd 

    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '¢Ð: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '¢Ð: Read Operation Mode (CRUD)
    lgKeyStream       = Split(Request("txtKeyStream"),gColSep)

    lgLngMaxRow       = Request("txtMaxRows")                                        '¢Ð: Read Operation Mode (CRUD)
    lgStrPrevKey = UNICInt(Trim(Request("lgStrPrevKey")),0)                '¢Ð: "0"(First),"1"(Second),"2"(Third),"3"(...)
    
    Call SubOpenDB(lgObjConn)           
   
    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '¢Ð: Query
             Call SubBizQuery()
        Case CStr(UID_M0002)                                                         '¢Ð: Save,Update
             Call SubBizSaveMulti()
        Case CStr(UID_M0003)                                                         '¢Ð: Delete
             Call SubBizDelete()
    End Select
'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()
     On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear
    Call SubBizQueryMulti()
End Sub    
'============================================================================================================
' Name : SubBizDelete
' Desc : Delete DB data
'============================================================================================================
Sub SubBizDelete()
     On Error Resume Next                                                             '¢Ð: Protect system from crashing
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
    Err.Clear                                                               '¢Ð: Clear Error status

	
    Call SubMakeSQLStatements("MR",strWhere,"X",C_LIKE)                              '¢Ð : Make sql statements

    If FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL&strWhere,"X","X") = False Then
       lgStrPrevKey = ""
       Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '¢Ð : No data is found. 
       Call SetErrorStatus()
    Else
       Call SubSkipRs(lgObjRs,C_SHEETMAXROWS_D * lgStrPrevKey)
       lgstrData = ""
       iDx       = 1
       Do While Not lgObjRs.EOF

            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("USER_ID"))
            lgstrData = lgstrData & Chr(11)
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("USR_NM"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("CDN_RCV"))			
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
    Err.Clear                                                                        '¢Ð: Clear Error status
    
	arrRowVal = Split(Request("txtSpread"), gRowSep)                                 '¢Ð: Split Row    data
	
    For iDx = 1 To lgLngMaxRow
        arrColVal = Split(arrRowVal(iDx-1), gColSep)                                 '¢Ð: Split Column data
        
        Select Case arrColVal(C_IG_CMD)
            Case "C"
                    Call SubBizSaveMultiCreate(arrColVal)                            '¢Ð: Create
            Case "U"
                    Call SubBizSaveMultiUpdate(arrColVal)                            '¢Ð: Update
            Case "D"
                    Call SubBizSaveMultiDelete(arrColVal)                            '¢Ð: Delete
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
    Err.Clear                                                                        '¢Ð: Clear Error status
  
	
	If CommonQueryRs(" USR_ID "," Z_USR_MAST_REC ", " USR_ID=" & FilterVar(arrColVal(C_IG_USER_ID), "''", "S"),lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) Then
 		If lgF0="X" Then
			Call SubHandleError("M1",lgObjConn,lgObjRs,Err)
			Exit Sub
		End If
	End If
   	
		lgStrSQL = "INSERT INTO B_CDN_USER_KO441(USER_ID,CDN_RCV,CDN_BIZ,CDN_BMP,CDN_PKG,CDN_PRD,CDN_TQC,REMARK,INSRT_USER_ID,INSRT_DT,UPDT_USER_ID,UPDT_DT) "
		lgStrSQL = lgStrSQL & " VALUES("     
		lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(C_IG_USER_ID)), "''", "S")     & ","		
		If arrColVal(C_IG_CDN_RCV)="1" Then
			lgStrSQL = lgStrSQL & FilterVar("Y", "''", "S")     & ","
		Else
			lgStrSQL = lgStrSQL & FilterVar("N", "''", "S")     & ","
		End If
		If arrColVal(C_IG_CDN_BIZ)="1" Then
			lgStrSQL = lgStrSQL & FilterVar("Y", "''", "S")     & ","
		Else
			lgStrSQL = lgStrSQL & FilterVar("N", "''", "S")     & ","
		End If
		If arrColVal(C_IG_CDN_BMP)="1" Then
			lgStrSQL = lgStrSQL & FilterVar("Y", "''", "S")     & ","
		Else
			lgStrSQL = lgStrSQL & FilterVar("N", "''", "S")     & ","
		End If
		If arrColVal(C_IG_CDN_PKG)="1" Then
			lgStrSQL = lgStrSQL & FilterVar("Y", "''", "S")     & ","
		Else
			lgStrSQL = lgStrSQL & FilterVar("N", "''", "S")     & ","
		End If
		If arrColVal(C_IG_CDN_PRD)="1" Then
			lgStrSQL = lgStrSQL & FilterVar("Y", "''", "S")     & ","
		Else
			lgStrSQL = lgStrSQL & FilterVar("N", "''", "S")     & ","
		End If
		If arrColVal(C_IG_CDN_TQC)="1" Then
			lgStrSQL = lgStrSQL & FilterVar("Y", "''", "S")     & ","
		Else
			lgStrSQL = lgStrSQL & FilterVar("N", "''", "S")     & ","
		End If
		lgStrSQL = lgStrSQL & FilterVar(Trim(arrColVal(C_IG_REMARK)), "''", "S")     & ","
		lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''", "S")                        & "," 
		lgStrSQL = lgStrSQL & FilterVar(lgSvrDateTime, "''", "S")   & "," 
		lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''", "S")		   & "," 
		lgStrSQL = lgStrSQL & FilterVar(lgSvrDateTime, "''", "S")   
		lgStrSQL = lgStrSQL & ")"  

		lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
		Call SubHandleError("MC",lgObjConn,lgObjRs,Err)
		
End Sub
'============================================================================================================
' Name : SubBizSaveMultiUpdate
' Desc : Update Data from Db
'============================================================================================================
Sub SubBizSaveMultiUpdate(arrColVal)
 
	If CommonQueryRs(" USR_ID "," Z_USR_MAST_REC ", " USR_ID=" & FilterVar(arrColVal(C_IG_USER_ID), "''", "S"),lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) Then
 		If lgF0="X" Then
			Call SubHandleError("M1",lgObjConn,lgObjRs,Err)
			Exit Sub
		End If
	End If

		lgStrSQL = "UPDATE B_CDN_USER_KO441 SET "
		If arrColVal(C_IG_CDN_RCV)="1" Then
			lgStrSQL = lgStrSQL & " CDN_RCV	=" & FilterVar("Y", "''", "S")     & ","
		Else
			lgStrSQL = lgStrSQL & " CDN_RCV	=" & FilterVar("N", "''", "S")     & ","
		End If
		If arrColVal(C_IG_CDN_BIZ)="1" Then
			lgStrSQL = lgStrSQL & " CDN_BIZ			=" & FilterVar("Y", "''", "S")     & ","
		Else
			lgStrSQL = lgStrSQL & " CDN_BIZ			=" & FilterVar("N", "''", "S")     & ","
		End If
		If arrColVal(C_IG_CDN_BMP)="1" Then
			lgStrSQL = lgStrSQL & " CDN_BMP			=" & FilterVar("Y", "''", "S")     & ","
		Else
			lgStrSQL = lgStrSQL & " CDN_BMP			=" & FilterVar("N", "''", "S")     & ","
		End If
		If arrColVal(C_IG_CDN_PKG)="1" Then
			lgStrSQL = lgStrSQL & " CDN_PKG			=" & FilterVar("Y", "''", "S")     & ","
		Else
			lgStrSQL = lgStrSQL & " CDN_PKG			=" & FilterVar("N", "''", "S")     & ","
		End If
		If arrColVal(C_IG_CDN_PRD)="1" Then
			lgStrSQL = lgStrSQL & " CDN_PRD			=" & FilterVar("Y", "''", "S")     & ","
		Else
			lgStrSQL = lgStrSQL & " CDN_PRD			=" & FilterVar("N", "''", "S")     & ","
		End If
		If arrColVal(C_IG_CDN_TQC)="1" Then
			lgStrSQL = lgStrSQL & " CDN_TQC			=" & FilterVar("Y", "''", "S")     & ","
		Else
			lgStrSQL = lgStrSQL & " CDN_TQC			=" & FilterVar("N", "''", "S")     & ","
		End If
		lgStrSQL = lgStrSQL & " REMARK				=" & FilterVar(Trim(arrColVal(C_IG_REMARK)), "''", "S")     & ","
		lgStrSQL = lgStrSQL & " UPDT_USER_ID	=" & FilterVar(gUsrId, "''", "S")                        & "," 
		lgStrSQL = lgStrSQL & " UPDT_DT				=" & FilterVar(lgSvrDateTime, "''", "S")   
		lgStrSQL = lgStrSQL & " WHERE USER_ID=" & FilterVar(UCase(arrColVal(C_IG_USER_ID)), "''", "S")

    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
	Call SubHandleError("MU",lgObjConn,lgObjRs,Err)
	
End Sub

'============================================================================================================
' Name : SubBizSaveMultiDelete
' Desc : Delete Data from Db

'============================================================================================================
Sub SubBizSaveMultiDelete(arrColVal)
     On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status

    lgStrSQL = "DELETE  B_CDN_USER_KO441"
		lgStrSQL = lgStrSQL & " WHERE USER_ID=" & FilterVar(UCase(arrColVal(C_IG_USER_ID)), "''", "S")
    
    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
		Call SubHandleError("MD",lgObjConn,lgObjRs,Err)
	
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
                       lgStrSQL = lgStrSQL & " a.USER_ID,  " & vbCrLf
                       lgStrSQL = lgStrSQL & " b.USR_NM, " & vbCrLf
                       lgStrSQL = lgStrSQL & " case when a.CDN_RCV='Y' then 1 else 0 end CDN_RCV, " & vbCrLf
                       lgStrSQL = lgStrSQL & " case when a.CDN_BIZ='Y' then 1 else 0 end CDN_BIZ, " & vbCrLf
                       lgStrSQL = lgStrSQL & " case when a.CDN_BMP='Y' then 1 else 0 end CDN_BMP, " & vbCrLf
                       lgStrSQL = lgStrSQL & " case when a.CDN_PKG='Y' then 1 else 0 end CDN_PKG, " & vbCrLf
                       lgStrSQL = lgStrSQL & " case when a.CDN_PRD='Y' then 1 else 0 end CDN_PRD, " & vbCrLf
                       lgStrSQL = lgStrSQL & " case when a.CDN_TQC='Y' then 1 else 0 end CDN_TQC, " & vbCrLf
                       lgStrSQL = lgStrSQL & " a.REMARK " & vbCrLf
                       lgStrSQL = lgStrSQL & " from B_CDN_USER_KO441 a " & vbCrLf
                       lgStrSQL = lgStrSQL & " LEFT OUTER JOIN Z_USR_MAST_REC b on (a.USER_ID=b.USR_ID) " & vbCrLf

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
                 Call DisplayMsgBox("173132", vbInformation, "»ç¿ëÀÚID", "", I_MKSCRIPT)     '
                 ObjectContext.SetAbort
                 Call SetErrorStatus
        Case "M2"
                 Call DisplayMsgBox("202404", vbInformation, "", "", I_MKSCRIPT)     '
                 ObjectContext.SetAbort
                 Call SetErrorStatus
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
       Case "<%=UID_M0003%>"                                                         '¢Ð : Delete
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DbDeleteOk
          Else   
          End If   
    End Select       
       
</Script>	
