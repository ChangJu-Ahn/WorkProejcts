<%@ LANGUAGE=VBSCript TRANSACTION=Required%>
<!-- #Include file="../../inc/Adovbs.inc"  -->
<!-- #Include file="../../inc/incServerAdoDb.asp" -->
<!-- #Include file="../../inc/incServer.asp" -->
<!-- #Include file="../../inc/incSvrFuncSims.inc" -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<!-- #Include file="../../inc/incSvrVarSims.inc"  -->

<%
    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status

   	Call HideStatusWnd_uniSIMS
                                                               '¢Ð: Hide Processing message
    '---------------------------------------Common-----------------------------------------------------------
    lgErrorStatus     = "NO"
    lgOpModeCRUD      = Request("txtMode")                                           '¢Ð: Read Operation Mode (CRUD)
    lgPrevNext        = Request("txtPrevNext")
    lgKeyStream       = Split(Request("txtKeyStream"),gColSep)

	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------ 

    Call SubOpenDB(lgObjConn)                                                        '¢Ð: Make a DB Connection

    Select Case lgOpModeCRUD
        Case "UID_M0001"                                                         '¢Ð: Query
             Call SubBizQuery()
        Case "UID_M0002"                                                     '¢Ð: Save,Update
            Call SubBizSave()
        Case CStr(UID_M0003)                                                         '¢Ð: Delete
             Call SubBizDelete()
    End Select
    
    Call SubCloseDB(lgObjConn)                                                       '¢Ð: Close DB Connection

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()
    Dim iKey1
    dim iRet
    Dim strWhere

    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status

	if gProAuth = 0 then
		iKey1 = FilterVar(lgKeyStream(0), "''", "S")
	else	
		iKey1 = FilterVar(lgKeyStream(0), "''", "S")
		iKey1 = iKey1 & " AND internal_cd LIKE  " & FilterVar(lgKeyStream(1) & "%", "''", "S") & ""
	end if
	iKey1 = iKey1 & " AND retire_dt is null"		
	 
    Call SubMakeSQLStatements("R",iKey1)                                       '¢Ð : Make sql statements
'response.write lgKeyStream(0) & "    " & lgStrSQL
'response.end
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then                    'If data not exists

    	If gProAuth = 0 Then
            %>
            <Script Language=vbscript>
                With parent.parent
                    .txtName2.Value = ""
                End With
    			With Parent
    				.FncNew()    
    			End With             
            </Script>       
            <%		
   	
    	Else
		    strWhere = " emp_no=" & FilterVar(lgKeyStream(0), "''", "S")
		    strWhere = strWhere & " AND retire_dt is null"   
    	
    		Call CommonQueryRs(" internal_cd "," HAA010T ",strWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) 
    		lgF0= Replace(lgF0, Chr(11), "")
    		
    		if lgF0="X" or lgF0="" then
    			Call DisplayMsgBox("800048", vbInformation, "", "", I_MKSCRIPT)
            %>
            <Script Language=vbscript>
                With parent.parent
                    .txtName2.Value = ""
                End With
    			With Parent
    				.FncNew()    
    			End With             
            </Script>       
            <%			
    			Response.end
    		end if

    		if inStr(1,ConvSPChars(lgF0),ConvSPChars(lgKeyStream(1)))=0 then
            %>
            <Script Language=vbscript>
                With parent.parent
                    .txtName2.Value = ""
                End With
    			With Parent
    				.FncNew()    
    			End With             
            </Script>       
            <%		
        		Call DisplayMsgBox("800454", vbInformation, "", "", I_MKSCRIPT)
    			Response.end
    		end if
		End If
		
		If lgPrevNext = "" Then
			Call DisplayMsgBox("800048", vbInformation, "", "", I_MKSCRIPT)      '¢Ð : No data is found. 
			Call SetErrorStatus()
		ElseIf lgPrevNext = "P" Then
			Call DisplayMsgBox("900011", vbInformation, "", "", I_MKSCRIPT)      '¢Ð : This is the starting data. 
			lgPrevNext = ""
			Call SubBizQuery()
		ElseIf lgPrevNext = "N" Then
			Call DisplayMsgBox("900012", vbInformation, "", "", I_MKSCRIPT)      '¢Ð : This is the ending data.
			lgPrevNext = ""
			Call SubBizQuery()
		End If
    Else

%>
<Script Language=vbscript>
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	' Set condition area, contents area
	'--------------------------------------------------------------------------------------------------------
	   With parent.parent
		     .txtEmp_no2.Value = "<%=ConvSPChars(lgObjRs("emp_no"))%>"
		     .txtName2.Value = "<%=ConvSPChars(lgObjRs("name"))%>"
	   End With     
	
       With Parent	
            .Frm1.txtEmp_no.Value  = "<%=ConvSPChars(lgObjRs("emp_no"))%>"
            .Frm1.txtName.Value    = "<%=ConvSPChars(lgObjRs("name"))%>"
            .frm1.txtDept_nm.value = "<%=ConvSPChars(lgObjRs("DEPT_NM"))%>"    
            .frm1.txtroll_pstn.value = "<%=lgObjRs("roll_pstn_nm")%>"

            .frm1.txtresent_promote_dt.value = "<%=UniConvDateDbToCompany(lgObjRs("resent_promote_dt"),"")%>"
            .frm1.txtHand_tel_no.value = "<%=ConvSPChars(lgObjRs("Hand_tel_no"))%>"
            .frm1.txtEMail_addr.value = "<%=ConvSPChars(lgObjRs("EMail_addr"))%>"
            .frm1.txtentr_dt.value = "<%=UniConvDateDbToCompany(lgObjRs("entr_dt"),"")%>"
            .frm1.txtgroup_entr_dt.value = "<%=UniConvDateDbToCompany(lgObjRs("group_entr_dt"),"")%>"
            .frm1.txteng_name.value = "<%=ConvSPChars(lgObjRs("eng_name"))%>"
            .frm1.txtbirt.value = "<%=UniConvDateDbToCompany(lgObjRs("birt"),"")%>"
            .frm1.txtmemo_cd.value = "<%=(lgObjRs("memo_cd"))%>"
        
            .frm1.txtmemo_dt.value = "<%=UniConvDateDbToCompany(lgObjRs("memo_dt"),"")%>"

	    	<%	If ConvSPChars(lgObjRs("so_lu_cd")) = "1" Then	%>		'
	    		.frm1.txtso_lu_cd1.Click()
	    	<%	Else  %>
	    		.frm1.txtso_lu_cd2.Click()
	    	<%  End If	%>

            .frm1.txtmarry_cd.value = "<%=ConvSPChars(lgObjRs("marry_cd"))%>"
            .frm1.txthouse_cd.value = "<%=ConvSPChars(lgObjRs("house_cd"))%>"
            .frm1.txthgt.value = "<%=UNINumClientFormat(lgObjRs("hgt"), 1,0)%>"              
            .frm1.txtwgt.value = "<%=UNINumClientFormat(lgObjRs("wgt"), 1,0)%>"  
            .frm1.txteyesgt_left.value = "<%=UNINumClientFormat(lgObjRs("eyesgt_left"), 1,0)%>"
            .frm1.txteyesgt_right.value = "<%=UNINumClientFormat(lgObjRs("eyesgt_right"), 1,0)%>"
  
            .frm1.txtblood_type1.value = "<%=ConvSPChars(lgObjRs("blood_type1"))%>"      
            .frm1.txtblood_type2.value = "<%=ConvSPChars(lgObjRs("blood_type2"))%>"      
            .frm1.txtnat_cd.value = "<%=ConvSPChars(lgObjRs("nat_cd"))%>"           
            .frm1.txtdomi.value = "<%=ConvSPChars(lgObjRs("domi"))%>"             

            .frm1.txtzip_cd.value = "<%=ConvSPChars(lgObjRs("zip_cd"))%>"           
            .frm1.txtaddr.value = "<%=ConvSPChars(lgObjRs("addr"))%>"             
            .frm1.txtcurr_zip_cd.value = "<%=ConvSPChars(lgObjRs("curr_zip_cd"))%>"      
            .frm1.txtcurr_addr.value = "<%=ConvSPChars(lgObjRs("curr_addr"))%>"        
            .frm1.txttel_no.value = "<%=ConvSPChars(lgObjRs("tel_no"))%>"           
            .frm1.txtem_tel_no.value = "<%=ConvSPChars(lgObjRs("em_tel_no"))%>"        

            if "<%=ConvSPChars(lgObjRs("dalt_type"))%>" = "Y" THEN   '»ö¸Í 
                .frm1.txtdalt_type.checked = true
            else
                .frm1.txtdalt_type.checked = false
            end if

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
	
		On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status
 

    lgIntFlgMode = CInt(Request("txtFlgMode"))                                       '¢Ð: Read Operayion Mode (CREATE, UPDATE)

    Select Case lgIntFlgMode
        Case  OPMD_CMODE                                                             '¢Ð : Create
			Call SubBizSaveSingleCreate()  
        
        Case  OPMD_UMODE                                                             '¢Ð : Update
            Call SubBizSaveSingleUpdate()
    
    End Select

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
End Sub    

'============================================================================================================
' Name : SubBizSaveSingleUpdate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveSingleUpdate()
    
    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status
    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to update record
    '--------------------------------------------------------------------------------------------------------
    lgStrSQL = "UPDATE  HAA010T"
    lgStrSQL = lgStrSQL & " SET " 
    lgStrSQL = lgStrSQL & " eng_name = " & FilterVar(Request("txteng_name"), "''", "S") & ","
    if Trim(Request("txtmemo_dt"))="" then
		lgStrSQL = lgStrSQL & " memo_dt = null,"                       
	else 
		lgStrSQL = lgStrSQL & " memo_dt = " & FilterVar(UNIConvDateCompanyToDB(Request("txtmemo_dt"),NULL), "''", "S") & "," ' datetime		
	end if		
	lgStrSQL = lgStrSQL & " memo_cd = " & FilterVar(UCase(Request("txtmemo_cd")), "''", "S") & ","                       	
    lgStrSQL = lgStrSQL & " so_lu_cd = " & FilterVar(UCase(Request("txtso_lu_cdv")), "''", "S") & ","
    lgStrSQL = lgStrSQL & " BIRT = " & FilterVar(UNIConvDateCompanyToDB(Request("txtbirt"),NULL), "''", "S") & ","       ' datetime
    lgStrSQL = lgStrSQL & " MARRY_CD =  " & FilterVar(Request("txtMarry_Cd"), "''", "S") & ","
    If IsEmpty(Request("txtDalt_type")) = true Then
        lgStrSQL = lgStrSQL & " dalt_type = " & FilterVar("N", "''", "S") & ","
    ELSE
        lgStrSQL = lgStrSQL & " dalt_type = " & FilterVar("Y", "''", "S") & ","
    END IF
    lgStrSQL = lgStrSQL & " house_cd = " & FilterVar(UCase(Request("txthouse_cd")), "''", "S") & ","
    lgStrSQL = lgStrSQL & " Hand_tel_no = " & FilterVar(Request("txtHand_tel_no"), "''", "S") & ","
    lgStrSQL = lgStrSQL & " DOMI = " & FilterVar(Request("txtDomi"), "''", "S") & ","
    lgStrSQL = lgStrSQL & " ZIP_CD = " & FilterVar(UCase(Request("txtZip_cd")), "''", "S") & ","
    lgStrSQL = lgStrSQL & " ADDR = " & FilterVar(Request("txtAddr"), "''", "S") & ","
    lgStrSQL = lgStrSQL & " CURR_ZIP_CD = " & FilterVar(UCase(Request("txtCurr_zip_cd")), "''", "S") & ","
    lgStrSQL = lgStrSQL & " CURR_ADDR = " & FilterVar(Request("txtCurr_addr"), "''", "S") & ","
    lgStrSQL = lgStrSQL & " EMAIL_ADDR = " & FilterVar(Request("txtEmail_addr"), "''", "S") & ","
   
    lgStrSQL = lgStrSQL & " TEL_NO = " & FilterVar(Request("txtTel_no"), "''", "S") & ","
    lgStrSQL = lgStrSQL & " EM_TEL_NO = " & FilterVar(Request("txtEm_tel_no"), "''", "S") & ","
    lgStrSQL = lgStrSQL & " HGT = " & UNIConvNum(Request("txtHgt"),0) & ","
    lgStrSQL = lgStrSQL & " WGT = " & UNIConvNum(Request("txtWgt"),0) & ","
    lgStrSQL = lgStrSQL & " BLOOD_TYPE1 = " & FilterVar(UCase(Request("txtBlood_type1")), "''", "S") & ","
    lgStrSQL = lgStrSQL & " BLOOD_TYPE2 = " & FilterVar(UCase(Request("txtBlood_type2")), "''", "S") & ","

    lgStrSQL = lgStrSQL & " EYESGT_RIGHT = " & UNIConvNum(Request("txtEyesgt_right"),0) & ","
    lgStrSQL = lgStrSQL & " EYESGT_LEFT = " & UNIConvNum(Request("txtEyesgt_left"),0) & ""
    lgStrSQL = lgStrSQL & " WHERE emp_no =  " & FilterVar(Request("txtEmp_no"), "''", "S") & ""

    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
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
    On Error Resume Next                                                             '¢Ð: Protect system from crashing

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to create record
    '--------------------------------------------------------------------------------------------------------
    
End Sub
'============================================================================================================
' Name : SubBizSaveMultiUpdate
' Desc : Update Data from Db
'============================================================================================================
Sub SubBizSaveMultiUpdate(arrColVal)

    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status
    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to update record
    '--------------------------------------------------------------------------------------------------------
    
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
                      lgStrSQL = "Select *,dbo.ufn_GetCodeName(" & FilterVar("H0002", "''", "S") & ",roll_pstn) roll_pstn_nm"
                      lgStrSQL = lgStrSQL & " From  HAA010T "
                      lgStrSQL = lgStrSQL & " WHERE emp_no = " & pCode 	
                Case "P"
                      lgStrSQL = "Select TOP 1 *,dbo.ufn_GetCodeName(" & FilterVar("H0002", "''", "S") & ",roll_pstn) roll_pstn_nm " 
                      lgStrSQL = lgStrSQL & " From  HAA010T "
                      lgStrSQL = lgStrSQL & " WHERE emp_no < " & pCode 	
                      lgStrSQL = lgStrSQL & " ORDER BY emp_no DESC "
                Case "N"
                      lgStrSQL = "Select TOP 1 *,dbo.ufn_GetCodeName(" & FilterVar("H0002", "''", "S") & ",roll_pstn) roll_pstn_nm " 
                      lgStrSQL = lgStrSQL & " From  HAA010T "
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
'Response.End
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
    lgErrorStatus     = "YES"                                                         '¢Ð: Set error status
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub
'============================================================================================================
' Name : SubHandleError
' Desc : This Sub handle error
'============================================================================================================
Sub SubHandleError(pOpCode,pConn,pRs,pErr)
    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status

    Select Case pOpCode
        Case "SC"
        Case "SD"
        Case "SR"
        Case "SU"
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
       Case "UID_M0001"                                                         '¢Ð : Query
         If Trim("<%=lgErrorStatus%>") = "NO" Then
              With Parent
                 .DBQueryOk        
	          End with
	      Else
              With Parent
                 .DBQueryFail
	          End with
          End If   
       Case "UID_M0002"                                                         '¢Ð : Save
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DBSaveOk
          Else
       '      Parent.SubSetErrPos(Trim("<%=lgErrorPos%>"))
          End If   
       Case "UID_M0003"
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DbDeleteOk
          End If   
    End Select    
       
</Script>	
