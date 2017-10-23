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
Const C_IG_TODO_DOC = 2
Const C_IG_COMBO_YN = 3
Const C_IG_UD_MAJOR_CD = 4
Const C_IG_UD_MINOR_CD = 5
Const C_IG_SAMPLE_DATA = 6
Const C_IG_PROCESS_TYPE = 7
Const C_IG_MES_USE_YN = 8
Const C_IG_CDN_BIZ = 9
Const C_IG_CDN_BMP = 10
Const C_IG_CDN_PKG = 11
Const C_IG_CDN_PRD = 12
Const C_IG_CDN_TQC = 13
Const C_IG_REMARK = 14
Const C_IG_ROW = 15

	Dim lgStrPrevKey
	Dim orgChangID
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

            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("SEQ"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("TODO_DOC"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("COMBO_YN"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("UD_MAJOR_CD"))
            lgstrData = lgstrData & Chr(11)
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("UD_MAJOR_NM"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("UD_MINOR_CD"))			
            lgstrData = lgstrData & Chr(11)
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("SAMPLE_DATA"))			
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
  	Dim iStrSeq
  	
	If arrColVal(C_IG_UD_MAJOR_CD) <> "" Then
	
		If CommonQueryRs(" UD_MAJOR_CD "," B_USER_DEFINED_MAJOR ", " UD_MAJOR_CD=" & FilterVar(arrColVal(C_IG_UD_MAJOR_CD), "''", "S"),lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) Then
	 		If lgF0="X" Then
				Call SubHandleError("M1",lgObjConn,lgObjRs,Err)
				Exit Sub
			End If
		End If
	
	End If

	If arrColVal(C_IG_UD_MINOR_CD) <> "" Then
	
		If CommonQueryRs(" UD_MINOR_CD "," B_USER_DEFINED_MINOR ", " UD_MAJOR_CD=" & FilterVar(arrColVal(C_IG_UD_MAJOR_CD),"''","S") & " AND UD_MINOR_CD=" & FilterVar(arrColVal(C_IG_UD_MINOR_CD),"''","S"),lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) Then
	 		If lgF0="X" Then
				Call SubHandleError("M2",lgObjConn,lgObjRs,Err)
				Exit Sub
			End If
		End If
	
	End If

		iStrSeq=""
		If CommonQueryRs(" ISNULL(MAX(SEQ),99)+1 "," B_CDN_DEPT_TODO_KO441 ", "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) Then
	 		iStrSeq = Replace(lgF0,chr(11),"")
		End If

		lgStrSQL = "INSERT INTO B_CDN_DEPT_TODO_KO441(SEQ,TODO_DOC,COMBO_YN,UD_MAJOR_CD,UD_MINOR_CD,SAMPLE_DATA,PROCESS_TYPE,MES_USE_YN,CDN_BIZ,CDN_BMP,CDN_PKG,CDN_PRD,CDN_TQC,REMARK,INSRT_USER_ID,INSRT_DT,UPDT_USER_ID,UPDT_DT) "
		lgStrSQL = lgStrSQL & " VALUES("     
		lgStrSQL = lgStrSQL & FilterVar(iStrSeq, "''", "S")     & ","		
		lgStrSQL = lgStrSQL & FilterVar(Trim(arrColVal(C_IG_TODO_DOC)), "''", "S")     & ","

		If arrColVal(C_IG_COMBO_YN)="1" Then
			lgStrSQL = lgStrSQL & FilterVar("Y", "''", "S")     & ","
		Else
			lgStrSQL = lgStrSQL & FilterVar("N", "''", "S")     & ","
		End If

		lgStrSQL = lgStrSQL & FilterVar(Trim(arrColVal(C_IG_UD_MAJOR_CD)), "''", "S")     & ","
		lgStrSQL = lgStrSQL & FilterVar(Trim(arrColVal(C_IG_UD_MINOR_CD)), "''", "S")     & ","
		lgStrSQL = lgStrSQL & FilterVar(Trim(arrColVal(C_IG_SAMPLE_DATA)), "''", "S")     & ","
		lgStrSQL = lgStrSQL & FilterVar(Trim(arrColVal(C_IG_PROCESS_TYPE)), "''", "S")     & ","

		If arrColVal(C_IG_MES_USE_YN)="1" Then
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
 
  	
	If arrColVal(C_IG_UD_MAJOR_CD) <> "" Then
	
		If CommonQueryRs(" UD_MAJOR_CD "," B_USER_DEFINED_MAJOR ", " UD_MAJOR_CD=" & FilterVar(arrColVal(C_IG_UD_MAJOR_CD), "''", "S"),lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) Then
	 		If lgF0="X" Then
				Call SubHandleError("M1",lgObjConn,lgObjRs,Err)
				Exit Sub
			End If
		End If
	
	End If

	If arrColVal(C_IG_UD_MINOR_CD) <> "" Then
	
		If CommonQueryRs(" UD_MINOR_CD "," B_USER_DEFINED_MINOR ", " UD_MAJOR_CD=" & FilterVar(arrColVal(C_IG_UD_MAJOR_CD),"''","S") & " AND UD_MINOR_CD=" & FilterVar(arrColVal(C_IG_UD_MINOR_CD),"''","S"),lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) Then
	 		If lgF0="X" Then
				Call SubHandleError("M2",lgObjConn,lgObjRs,Err)
				Exit Sub
			End If
		End If
	
	End If

		lgStrSQL = "UPDATE B_CDN_DEPT_TODO_KO441 SET "

		lgStrSQL = lgStrSQL & " TODO_DOC=" & FilterVar(Trim(arrColVal(C_IG_TODO_DOC)), "''", "S")     & ","

		If arrColVal(C_IG_COMBO_YN)="1" Then
			lgStrSQL = lgStrSQL & " COMBO_YN=" & FilterVar("Y", "''", "S")     & ","
		Else
			lgStrSQL = lgStrSQL & " COMBO_YN=" & FilterVar("N", "''", "S")     & ","
		End If

		lgStrSQL = lgStrSQL & " UD_MAJOR_CD=" & FilterVar(Trim(arrColVal(C_IG_UD_MAJOR_CD)), "''", "S")     & ","
		lgStrSQL = lgStrSQL & " UD_MINOR_CD=" & FilterVar(Trim(arrColVal(C_IG_UD_MINOR_CD)), "''", "S")     & ","
		lgStrSQL = lgStrSQL & " SAMPLE_DATA=" & FilterVar(Trim(arrColVal(C_IG_SAMPLE_DATA)), "''", "S")     & ","
		lgStrSQL = lgStrSQL & " PROCESS_TYPE=" & FilterVar(Trim(arrColVal(C_IG_PROCESS_TYPE)), "''", "S")     & ","

		If arrColVal(C_IG_MES_USE_YN)="1" Then
			lgStrSQL = lgStrSQL & " MES_USE_YN=" & FilterVar("Y", "''", "S")     & ","
		Else
			lgStrSQL = lgStrSQL & " MES_USE_YN=" & FilterVar("N", "''", "S")     & ","
		End If
		
		If arrColVal(C_IG_CDN_BIZ)="1" Then
			lgStrSQL = lgStrSQL & " CDN_BIZ=" & FilterVar("Y", "''", "S")     & ","
		Else
			lgStrSQL = lgStrSQL & " CDN_BIZ=" & FilterVar("N", "''", "S")     & ","
		End If
		
		If arrColVal(C_IG_CDN_BMP)="1" Then
			lgStrSQL = lgStrSQL & " CDN_BMP=" & FilterVar("Y", "''", "S")     & ","
		Else
			lgStrSQL = lgStrSQL & " CDN_BMP=" & FilterVar("N", "''", "S")     & ","
		End If
		
		If arrColVal(C_IG_CDN_PKG)="1" Then
			lgStrSQL = lgStrSQL & " CDN_PKG=" & FilterVar("Y", "''", "S")     & ","
		Else
			lgStrSQL = lgStrSQL & " CDN_PKG=" & FilterVar("N", "''", "S")     & ","
		End If
		
		If arrColVal(C_IG_CDN_PRD)="1" Then
			lgStrSQL = lgStrSQL & " CDN_PRD=" & FilterVar("Y", "''", "S")     & ","
		Else
			lgStrSQL = lgStrSQL & " CDN_PRD=" & FilterVar("N", "''", "S")     & ","
		End If
		
		If arrColVal(C_IG_CDN_TQC)="1" Then
			lgStrSQL = lgStrSQL & " CDN_TQC=" & FilterVar("Y", "''", "S")     & ","
		Else
			lgStrSQL = lgStrSQL & " CDN_TQC=" & FilterVar("N", "''", "S")     & ","
		End If

		lgStrSQL = lgStrSQL & " REMARK				=" & FilterVar(Trim(arrColVal(C_IG_REMARK)), "''", "S")     & ","
		lgStrSQL = lgStrSQL & " UPDT_USER_ID	=" & FilterVar(gUsrId, "''", "S")                        & "," 
		lgStrSQL = lgStrSQL & " UPDT_DT				=" & FilterVar(lgSvrDateTime, "''", "S")   
		lgStrSQL = lgStrSQL & " WHERE 	SEQ		=" & FilterVar(UCase(arrColVal(C_IG_SEQ)), "''", "S")

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

    lgStrSQL = "DELETE  B_CDN_DEPT_TODO_KO441"
		lgStrSQL = lgStrSQL & " WHERE SEQ=" & FilterVar(UCase(arrColVal(C_IG_SEQ)), "''", "S")
    
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

                       lgStrSQL = lgStrSQL & " a.SEQ  " & vbCrLf
                       lgStrSQL = lgStrSQL & " ,a.TODO_DOC  " & vbCrLf
                       lgStrSQL = lgStrSQL & " ,case when a.COMBO_YN='Y' then 1 else 0 end COMBO_YN  " & vbCrLf
                       lgStrSQL = lgStrSQL & " ,a.UD_MAJOR_CD  " & vbCrLf
                       lgStrSQL = lgStrSQL & " ,b.UD_MAJOR_NM  " & vbCrLf
                       lgStrSQL = lgStrSQL & " ,a.UD_MINOR_CD  " & vbCrLf
                       lgStrSQL = lgStrSQL & " ,a.SAMPLE_DATA  " & vbCrLf
                       lgStrSQL = lgStrSQL & " ,a.PROCESS_TYPE  " & vbCrLf
                       lgStrSQL = lgStrSQL & " ,case when a.MES_USE_YN='Y' then 1 else 0 end  MES_USE_YN  " & vbCrLf
                       lgStrSQL = lgStrSQL & " ,case when a.CDN_BIZ='Y' then 1 else 0 end CDN_BIZ  " & vbCrLf
                       lgStrSQL = lgStrSQL & " ,case when a.CDN_BMP='Y' then 1 else 0 end CDN_BMP  " & vbCrLf
                       lgStrSQL = lgStrSQL & " ,case when a.CDN_PKG='Y' then 1 else 0 end CDN_PKG  " & vbCrLf
                       lgStrSQL = lgStrSQL & " ,case when a.CDN_PRD='Y' then 1 else 0 end CDN_PRD  " & vbCrLf
                       lgStrSQL = lgStrSQL & " ,case when a.CDN_TQC='Y' then 1 else 0 end CDN_TQC  " & vbCrLf
                       lgStrSQL = lgStrSQL & " ,REMARK  " & vbCrLf

                       lgStrSQL = lgStrSQL & " from B_CDN_DEPT_TODO_KO441 a  " & vbCrLf
                       lgStrSQL = lgStrSQL & " left outer join B_USER_DEFINED_MAJOR  b on (a.UD_MAJOR_CD=b.UD_MAJOR_CD)  " & vbCrLf

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
