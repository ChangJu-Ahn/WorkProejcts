<%@ LANGUAGE=VBSCript TRANSACTION=Required%>
<!-- #Include file="../../inc/Adovbs.inc"  -->
<!-- #Include file="../../inc/incServerAdoDb.asp" -->
<!-- #Include file="../../inc/incServer.asp" -->
<!-- #Include file="../../inc/incSvrFuncSims.inc" -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<!-- #Include file="../../inc/incSvrVarSims.inc"  -->

<%
'------------------------------------------------------------------------------------------------------------------------------
' Common variables
'------------------------------------------------------------------------------------------------------------------------------

Call HideStatusWnd_uniSIMS
'==============================================================================
' Date Query Key 
'==============================================================================
Dim IsOpenPop



Dim strPass
Dim Enc1, Enc2
                                                               '☜: Hide Processing message
    '---------------------------------------Common-----------------------------------------------------------
    lgErrorStatus     = "NO"
    lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)
    lgKeyStream       = Split(Request("txtKeyStream"),gColSep)
    lgPrevNext        = Request("txtPrevNext")        


	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	err.Clear
        Set Enc2 = Server.CreateObject("EDCodeCom.EDCodeObj.1")

        if err.number <> 0 then
%>
            <script language=vbscript>
                msgbox "CreateObject error - EDCodeCom", vbExclamation , _
                    "<%=Request.Cookies("unierp")("gLogoName")%>" & " login error 6"
                msgbox "<%=err.description%>"
            </script>
<%
            response.end
        end if

	'------ Developer Coding part (End   ) ------------------------------------------------------------------ 

    Call SubOpenDB(lgObjConn)                                                        '☜: Make a DB Connection
    
    Select Case lgOpModeCRUD
        Case "UID_M0001"                                                         '☜: Query
             Call SubBizQuery()
        Case "UID_M0002"                                                     '☜: Save,Update
             Call SubBizSave()
        Case "UID_M0003"
             Call SubBizDelete()
    End Select

    Set Enc2 = Nothing
    Call SubCloseDB(lgObjConn)                                                       '☜: Close DB Connection

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()
    Dim iKey1

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status


    iKey1 = FilterVar(lgKeyStream(2), "''", "S")


    Call SubMakeSQLStatements("R",iKey1)                                       '☜ : Make sql statements

    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then                    'If data not exists
       If lgPrevNext = "" Then
          'Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)
          Call SetErrorStatus()
       ElseIf lgPrevNext = "P" Then
          Call DisplayMsgBox("900011", vbInformation, "", "", I_MKSCRIPT)
          lgPrevNext = ""
          Call SubBizQuery()
       ElseIf lgPrevNext = "N" Then
          Call DisplayMsgBox("900012", vbInformation, "", "", I_MKSCRIPT)
          lgPrevNext = ""
          Call SubBizQuery()
       End If
       
    Else
    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to update record

        strPass = ConvSPChars(lgObjRs("password"))
        If strPass <> "" Then
            strPass = Enc2.Decode(strPass)
        End If

%>
<Script Language=vbscript>
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	' Set condition area, contents area
	'--------------------------------------------------------------------------------------------------------
       With Parent	
            .frm1.txtuser_id.value = "<%=ConvSPChars(lgObjRs("uid"))%>"
            .frm1.txtemp_no1.value = "<%=ConvSPChars(lgObjRs("emp_no"))%>"
            .frm1.txtname1.value = "<%=ConvSPChars(lgObjRs("name"))%>"
            .frm1.txtpassword.value = "<%=ConvSPChars(strPass)%>"
            .frm1.txtpro_auth.value = "<%=ConvSPChars(lgObjRs("pro_auth"))%>"
            .frm1.txtdept_authv.value = "<%=ConvSPChars(lgObjRs("dept_auth"))%>"
            .frm1.txtuse_ynv.value = "<%=ConvSPChars(lgObjRs("use_yn"))%>"
       End With          
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
</Script>       
<%     
    End If
    
    Call SubCloseRs(lgObjRs)
    
End Sub    

'============================================================================================================
' Name : SubBizSave
' Desc : Date data 
'============================================================================================================
Sub SubBizSave()
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
    
    lgIntFlgMode = CInt(Request("txtFlgMode"))                                       '☜: Read Operayion Mode (CREATE, UPDATE)

    Select Case lgIntFlgMode
        Case  OPMD_CMODE                                                             '☜ : Create
              Call SubBizSaveSingleCreate()  
        Case  OPMD_UMODE                                                             '☜ : Update
              Call SubBizSaveSingleUpdate()
    End Select

End Sub	

'============================================================================================================
' Name : SubBizDelete
' Desc : Delete DB data
'============================================================================================================
Sub SubBizDelete()

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    '미등록된 사번 
    lgStrSQL = "emp_no = " & FilterVar(Request("txtemp_no1"), "''", "S")
    Call CommonQueryRs(" count(emp_no) "," E11002T ", lgStrSQL,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    if  Replace(lgF0, Chr(11), "") = "X" then
    else
        if Cint(Replace(lgF0, Chr(11), "")) <> 0 then
        else
			Call DisplayMsgBox("800481", vbInformation, "", "", I_MKSCRIPT)  
            lgErrorStatus = "YES"
            exit sub
        end if
    end if
    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to update record
    '--------------------------------------------------------------------------------------------------------
    lgStrSQL = "DELETE  E11002T"
    lgStrSQL = lgStrSQL & " WHERE emp_no = " & FilterVar(lgKeyStream(2), "''", "S")                              ' 사번char(10)

    'lgStrSQL = lgStrSQL & "   AND trip_cd = " & FilterVar(lgKeyStream(2),"''", "S")

    '---------- Developer Coding part (End  ) ---------------------------------------------------------------

    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
	Call SubHandleError("SD",lgObjConn,lgObjRs,Err)

End Sub

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()
End Sub    

'============================================================================================================
' Name : SubBizSaveCreate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveSingleCreate()

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
   
	'License체크 
    Call CommonQueryRs(" count(UID) "," E11002T ", " USE_YN = " & FilterVar("Y", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

    If  CInt(GetLicenseInfo) <= CInt(Replace(lgF0, Chr(11), "")) then
        Call DisplayMsgBox("210008", vbInformation, "", "", I_MKSCRIPT)
        lgErrorStatus = "YES"
        Exit sub
	End If        

    Call CommonQueryRs(" count(emp_no) "," E11002T ", " UID in (select res_no from  haa010t group by res_no having  count(emp_no)>1) and uid= " & FilterVar(Request("txtuser_id"), "''", "S"),lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)    
    if lgF0>0 then
		Call DisplayMsgBox("800501", vbInformation, "", "", I_MKSCRIPT)
		Response.End
	end if
    
    strPass = Request("txtpassword")
    if strPass = "" then
    else
        strPass = Enc2.Encode(strPass)
    end if
    
    '--------------------------------------------------------------------------------------------------------

    lgStrSQL = "INSERT INTO E11002T("
    lgStrSQL = lgStrSQL & " uid, "
    lgStrSQL = lgStrSQL & " emp_no, "
    lgStrSQL = lgStrSQL & " password, "
    lgStrSQL = lgStrSQL & " pro_auth, "
    lgStrSQL = lgStrSQL & " dept_auth, "
    lgStrSQL = lgStrSQL & " use_yn, "
    lgStrSQL = lgStrSQL & " isrt_emp_no, "
    lgStrSQL = lgStrSQL & " updt_emp_no ) "
    lgStrSQL = lgStrSQL & " VALUES(" 
    lgStrSQL = lgStrSQL & FilterVar(Request("txtuser_id"), "''", "S") & ","
    lgStrSQL = lgStrSQL & FilterVar(Request("txtemp_no1"), "''", "S") & ","
    lgStrSQL = lgStrSQL & " " & FilterVar(strPass, "''", "S") & ","
    lgStrSQL = lgStrSQL & FilterVar(Request("txtpro_auth"), "''", "S") & ","
    lgStrSQL = lgStrSQL & FilterVar(Request("txtdept_authv"), "''", "S") & ","
    lgStrSQL = lgStrSQL & FilterVar(Request("txtuse_ynv"), "''", "S") & ","
    lgStrSQL = lgStrSQL & FilterVar(lgKeyStream(0), "''", "S") & ","
    lgStrSQL = lgStrSQL & FilterVar(lgKeyStream(0), "''", "S")
    lgStrSQL = lgStrSQL & ")"
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------

    'Response.Write lgStrSQL

    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
	Call SubHandleError("SC",lgObjConn,lgObjRs,Err)
	
End Sub
'============================================================================================================
' Name : SubBizSaveSingleUpdate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveSingleUpdate()

    
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to update record


    strPass = Request("txtpassword")
'    Response.Write strPass
    strPass = Enc2.Encode(strPass)

'Response.Write strPass
    '--------------------------------------------------------------------------------------------------------
    lgStrSQL = "UPDATE  E11002T"
    lgStrSQL = lgStrSQL & " SET " 
    lgStrSQL = lgStrSQL & " password =  " & FilterVar(strpass , "''", "S") & ","
    lgStrSQL = lgStrSQL & " pro_auth = " & FilterVar(Request("txtpro_auth"), "''", "S") & ","
    lgStrSQL = lgStrSQL & " dept_auth = " & FilterVar(Request("txtdept_authv"), "''", "S") & ","
    lgStrSQL = lgStrSQL & " use_yn = " & FilterVar(Request("txtuse_ynv"), "''", "S") & ","
    lgStrSQL = lgStrSQL & " updt_emp_no = " & FilterVar(lgKeyStream(0), "''", "S")
    lgStrSQL = lgStrSQL & " WHERE emp_no =  " & FilterVar(lgKeyStream(2), "''", "S") & ""
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
'Response.Write lgStrSQL

    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
	Call SubHandleError("SU",lgObjConn,lgObjRs,Err)
    
End Sub

'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data
'============================================================================================================
Sub SubBizSaveMulti()

    Dim arrRowVal
    Dim arrColVal
    Dim iDx

    On Error Resume Next                                                             '☜: Protect system from crashing

    Err.Clear                                                                        '☜: Clear Error status
    
	arrRowVal = Split(Request("txtSpread"), gRowSep)                                 '☜: Split Row    data
	
    For iDx = 1 To lgLngMaxRow
        arrColVal = Split(arrRowVal(iDx-1), gColSep)                                 '☜: Split Column data
        
        Select Case arrColVal(0)
            Case "C"
                    Call SubBizSaveMultiCreate(arrColVal)                            '☜: Create
            Case "U"
                    Call SubBizSaveMultiUpdate(arrColVal)                            '☜: Update
            Case "D"
                    Call SubBizSaveMultiDelete(arrColVal)                            '☜: Delete
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
    On Error Resume Next                                                             '☜: Protect system from crashing

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to create record
    '--------------------------------------------------------------------------------------------------------
    
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
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements(pMode,pCode)
    Dim iSelCount
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
    Select Case pMode 
      Case "R"
             Select Case  lgPrevNext 
                Case ""
                    lgStrSQL = "Select uid, emp_no, password, pro_auth, dept_auth, use_yn," 
                    lgStrSQL = lgStrSQL & " dbo.ufn_H_GetEmpName(emp_no)  as name  "
                    lgStrSQL = lgStrSQL & " From  E11002T "
                    lgStrSQL = lgStrSQL & " WHERE emp_no = " & pCode 	
                Case "P"
                    lgStrSQL = "Select TOP 1 uid, emp_no, password, pro_auth, dept_auth, " 
                    lgStrSQL = lgStrSQL & " dbo.ufn_H_GetEmpName(emp_no)  as name  "
                    lgStrSQL = lgStrSQL & " From  E11002T "
                    lgStrSQL = lgStrSQL & " WHERE emp_no < " & pCode 	
                    lgStrSQL = lgStrSQL & " ORDER BY emp_no DESC "
                Case "N"
                    lgStrSQL = "Select TOP 1 uid, emp_no, password, pro_auth, dept_auth, " 
                    lgStrSQL = lgStrSQL & " dbo.ufn_H_GetEmpName(emp_no)  as name  "
                    lgStrSQL = lgStrSQL & " From  E11002T "
                    lgStrSQL = lgStrSQL & " WHERE emp_no > " & pCode 	
                    lgStrSQL = lgStrSQL & " ORDER BY emp_no ASC "
             End Select
      Case "C"
             lgStrSQL = "Select * " 
             lgStrSQL = lgStrSQL & " From  B_MAJOR "
             lgStrSQL = lgStrSQL & " WHERE MAJOR_CD = " & pCode 	
      Case "U"
             lgStrSQL = "Select * " 
             lgStrSQL = lgStrSQL & " From  B_MAJOR "
             lgStrSQL = lgStrSQL & " WHERE MAJOR_CD = " & pCode 	
      Case "D"
             lgStrSQL = "Select * " 
             lgStrSQL = lgStrSQL & " From  B_MAJOR "
             lgStrSQL = lgStrSQL & " WHERE MAJOR_CD = " & pCode 	
    End Select
	'------ Developer Coding part (End   ) ------------------------------------------------------------------

'Response.Write lgStrSQL
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
					Call DisplayMsgBox("183116", vbInformation, "", "", I_MKSCRIPT)  
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
					Call DisplayMsgBox("183116", vbInformation, "", "", I_MKSCRIPT)  
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
Function GetLicenseInfo()
		Dim objConn
		Set objConn = CreateObject("uniSIMSLic.clsLicense")	
		
		if GetuniERPVer(Request.ServerVariables("APPL_PHYSICAL_PATH")) = "2.7" then
			GetLicenseInfo =  CStr(objConn.GetLicenseTotalUser (gCompany))
		else
			GetLicenseInfo =  CStr(objConn.GetLicenseTotalUser ())
		end if

		Set objConn = Nothing

End Function
function GetuniERPVer(path)

  Dim fso, ver
  Set fso = server.CreateObject("Scripting.FileSystemObject")
  If (fso.FileExists(Path & "uniSystemInfo.ini")) Then
    ver =  "2.7"
  Else
    ver =  "2.5"
  End If
  GetuniERPVer = ver
End Function

%>

<Script Language="VBScript">
    Select Case "<%=lgOpModeCRUD %>"
       Case "UID_M0001"                                                         '☜ : Query
         If Trim("<%=lgErrorStatus%>") = "NO" Then
              With Parent
                 .DBQueryOk        
	          End with
	      Else
              With Parent
                 .DBQueryFail
	          End with
          End If   
       Case "UID_M0002"                                                         '☜ : Save
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DBSaveOk
          Else
             'parent.SubSetErrPos(Trim("<%=lgErrorPos%>"))
          End If   
       Case "UID_M0003"
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DbDeleteOk
          End If   
    End Select    
       
</Script>	
