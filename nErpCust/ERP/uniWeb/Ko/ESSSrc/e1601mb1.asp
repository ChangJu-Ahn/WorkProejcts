<%@ LANGUAGE=VBSCript%>
<!-- #Include file="../ESSinc/Adovbs.inc"  -->
<!-- #Include file="../ESSinc/incServerAdoDb.asp" -->
<!-- #Include file="../ESSinc/incServer.asp" -->
<!-- #Include file="../ESSinc/incSvrFuncSims.inc" -->
<!-- #Include file="../ESSinc/lgsvrvariables.inc" -->
<!-- #Include file="../ESSinc/incSvrVarSims.inc"  -->
<%

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    Call HideStatusWnd_uniSIMS

    lgErrorStatus     = "NO"
    lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)
    lgKeyStream       = Split(Request("txtKeyStream"),gColSep)
    lgPrevNext        = Request("txtPrevNext")        

    Call SubOpenDB(lgObjConn)                                                        '☜: Make a DB Connection
    
    Select Case lgOpModeCRUD
        Case "UID_M0001"                                                         '☜: Query
             Call SubBizQuery()
        Case "UID_M0002"                                                     '☜: Save,Update
			 Call SubBizSave() 
        Case "UID_M0003"
             Call SubBizDelete()
    End Select
    
    Call SubCloseDB(lgObjConn)                                                       '☜: Close DB Connection

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()
    Dim iKey1,iRet
    Dim Emp_no,Name,roll_pstn,dept_nm,resent_promote_dt,group_entr_dt,entr_dt
    
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    iRet = EmpBaseDiligAuthCheck(lgKeyStream(0),lgKeyStream(5),lgKeyStream(6),lgKeyStream(7),lgPrevNext,Emp_no,Name,roll_pstn,dept_nm,resent_promote_dt,group_entr_dt,entr_dt)
 
    If iRet = True Then
%>
        <Script Language=vbscript>
            With parent.parent
                .txtEmp_no2.Value = "<%=ConvSPChars(emp_no)%>"
                .txtName2.Value = "<%=ConvSPChars(Name)%>"
            End With          
            With parent.frm1
                .txtEmp_no.Value = "<%=ConvSPChars(emp_no)%>"
                .txtName.Value = "<%=ConvSPChars(Name)%>"
                .txtDept_nm.value = "<%=ConvSPChars(DEPT_NM)%>"    
                .txtroll_pstn.value = "<%=ConvSPChars(roll_pstn)%>"
            End With 
            
        </Script>       
<%
    Else

        if  lgPrevNext = "N" then
            Call DisplayMsgBox("900012", vbInformation, "", "", I_MKSCRIPT)
        elseif lgPrevNext = "P" then

            Call DisplayMsgBox("900011", vbInformation, "", "", I_MKSCRIPT)           
		else 
			%>
			<Script Language=vbscript>
			        With parent.parent
			            .txtName2.Value = ""
			        End With          

			        With parent.frm1
			            .txtEmp_no.Value = ""
			            .txtName.Value = ""
			            .txtDept_nm.value = ""    
			            .txtroll_pstn.value = ""
			        End With 
			       With Parent
						.FncNew()       
			       End With         
			</Script>       
			<%   	    		
			Call DisplayMsgBox("800454", vbInformation, "", "", I_MKSCRIPT)	
        end if

        Call SetErrorStatus()
		Response.End        
    End If

    If  lgKeyStream(1) = "" then 
        lgErrorStatus = "YES"
        exit sub
    End if
    If  lgKeyStream(2) = "" or lgKeyStream(3) = "" then 
        lgErrorStatus = "NO"
        return
    End if
    
    if  lgPrevNext = "N" or lgPrevNext = "P" then
		%>
		<Script Language=vbscript>
		       With Parent
					.FncNew()
		       End With         
		</Script>       
		<%
	    Call SetErrorStatus()
	else 		
	    iKey1 = FilterVar(ConvSPChars(emp_no), "''", "S")
	    iKey1 = iKey1 & " AND dilig_strt_dt =  " & FilterVar( lgKeyStream(2), "''", "S") & ""
	    iKey1 = iKey1 & " AND dilig_cd =  " & FilterVar(lgKeyStream(4), "''", "S") & ""
	    iKey1 = iKey1 & " AND dilig_cd in (select dilig_cd from hca010t where  dilig_type=1)" '잔업이 아닌 근태type
	    iKey1 = iKey1 & "  Order by dilig_strt_dt DESC"
	    
	    Call SubMakeSQLStatements("R",iKey1)                                       '☜ : Make sql statements
'Response.Write lgStrSQL
'Response.end
	    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
	        lgStrPrevKeyIndex = ""
	        Call SetErrorStatus()
	    Else
		%>
		<Script Language=vbscript>
		       With Parent
		            .frm1.txtdilig_strt_dt.value = "<%=UNIDateClientFormat(lgObjRs("dilig_strt_dt"))%>"
		            .frm1.txtdilig_end_dt.value  = "<%=UNIDateClientFormat(lgObjRs("dilig_end_dt"))%>"
		            .frm1.txtdilig_cd.value      = "<%=ConvSPChars(lgObjRs("dilig_cd"))%>"
		            .frm1.txtremark.value        = "<%=ConvSPChars(lgObjRs("remark"))%>"
		            .frm1.txtapp_emp_no.value    = "<%=ConvSPChars(lgObjRs("app_emp_no"))%>"
		            .frm1.txtapp_name.value      = "<%=ConvSPChars(lgObjRs("app_name"))%>"
		       End With         
		</Script>       
		<%     
    
	    End If

	    Call SubCloseRs(lgObjRs)
    end if    
End Sub     

'============================================================================================================
' Name : SubBizSave
' Desc : Date data 
'============================================================================================================
Sub SubBizSave()
	Dim counts
	Dim i
	Dim strInput_emp_no
	Dim strClose_type
	Dim strClose_dt
	
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
 	if  Check_CloseDate() = False  then		'근태마감 여부를  체크한다.
        lgErrorStatus = "YES"
 		Exit Sub
	end if     

    lgIntFlgMode = CInt(Request("txtFlgMode"))                                       '☜: Read Operayion Mode (CREATE, UPDATE)
 
    Select Case lgIntFlgMode
        Case  OPMD_CMODE                                                             '☜ : Create
			if  SubCheckHoliday() = False then
               lgErrorStatus = "YES"
				Exit sub
			end if  

			'기간근태(e11070t)에서 기간에 (중복일자)속했는지를 check 한다 
			Call CommonQueryRs(" isnull(count(emp_no),0) "," e11070t ","  emp_no =  " & FilterVar( lgKeyStream(0), "''", "S") & " AND ((dilig_strt_dt between  " & FilterVar(UNIConvDateCompanyToDB((lgKeyStream(2)),NULL), "''", "S") & " AND  " & FilterVar(UNIConvDateCompanyToDB((lgKeyStream(3)),NULL), "''", "S") & ") OR (dilig_end_dt  between  " & FilterVar(UNIConvDateCompanyToDB((lgKeyStream(2)),NULL), "''", "S") & " AND  " & FilterVar(UNIConvDateCompanyToDB((lgKeyStream(2)),NULL), "''", "S") & ")) and dilig_cd in (select dilig_cd from hca010t where  dilig_type=1)" ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

			If Trim(Replace(lgF0,Chr(11),"")) = 0 then
			Else
		        Call DisplayMsgBox("800067", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
                lgErrorStatus = "YES"
			    Exit sub                                   
			End if
			
			'기간근태(hca050t)에서 기간에 (중복일자)속했는지를 check 한다. 만약 없으면 일일근태(hca060t)에 있는지도 check 한다.
	        Call CommonQueryRs(" isnull(count(emp_no),0) "," hca050t ","  emp_no =  " & FilterVar( lgKeyStream(0), "''", "S") & " AND ((dilig_strt_dt between  " & FilterVar(lgKeyStream(2), "''", "S") & " AND  " & FilterVar(lgKeyStream(3), "''", "S") & ") OR (dilig_end_dt  between  " & FilterVar(lgKeyStream(2), "''", "S") & " AND  " & FilterVar(lgKeyStream(2), "''", "S") & "))" ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
			If Trim(Replace(lgF0,Chr(11),"")) = 0 then
		
			    Call CommonQueryRs(" isnull(count(emp_no),0) "," hca060t ","  emp_no =  " & FilterVar( lgKeyStream(0), "''", "S") & " AND (dilig_dt between  " & FilterVar(lgKeyStream(2), "''", "S") & " AND  " & FilterVar(lgKeyStream(3), "''", "S") & ") and dilig_cd in (select dilig_cd from hca010t where  dilig_type=1)" ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
			    If Trim(Replace(lgF0,Chr(11),"")) > 0 then
			        Call DisplayMsgBox("800067", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
					%>
							<Script Language=vbscript>
							        parent.frm1.txtDilig_STRT_dt.focus()
							</Script>       
					<%  			        
                    lgErrorStatus = "YES"
			        Exit sub                                  
			    End if
			Else
		        Call DisplayMsgBox("800067", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
				%>
				<Script Language=vbscript>
				        parent.frm1.txtDilig_STRT_dt.focus()
				</Script>       
				<%  		        
                lgErrorStatus = "YES"
			    Exit sub                                   
			End if
			Call SubBizSaveSingleCreate()  
        Case  OPMD_UMODE                                                             '☜ : Update
            Call SubBizSaveSingleUpdate()
    End Select
End Sub	
'============================================================================================================
' Name : Check_CloseDate
' Desc : Delete DB data
'============================================================================================================
Function Check_CloseDate()
	Dim strReturn_value,strSQL,IntRetCD
	Dim strCloseDt, strValidDt	

	Check_CloseDate = False
	strReturn_value = "N"
    strSQL = " org_cd = " & FilterVar("1", "''", "S") & " AND pay_gubun = " & FilterVar("Z", "''", "S") & " AND PAY_TYPE = " & FilterVar("#", "''", "S") & " "
    IntRetCD = CommonQueryRs(" close_type, convert(char(10),close_dt,20), emp_no "," hda270t ", strSQL,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

    If  IntRetCd = false then
        strReturn_value = "Y"
    else
		strCloseDt = UniConvDateToYYYYMMDD(Replace(lgF1, Chr(11), ""),gServerDateFormat,"")
		strValidDt = UniConvDateToYYYYMMDD(lgKeyStream(2),gServerDateFormat,"")
'Response.Write ",strCloseDt:" & strCloseDt & ",strValidDt:" & strValidDt

        Select Case Replace(lgF0, Chr(11), "")
            Case "1"    '마감형태 : 정상 
                if strCloseDt <= strValidDt then
                    strReturn_value = "Y"
                else
                    strReturn_value = "N"
                end if

            Case "2"    '마감형태 : 마감 
                if strCloseDt < strValidDt then
                    strReturn_value = "Y"
                else
                    strReturn_value = "N"
                end if
        end Select
    end if

    if  strReturn_value = "N" then

        Call DisplayMsgBox("800291", vbInformation, "", "", I_MKSCRIPT) 
        exit Function
	else 
		Check_CloseDate = True	        
    end if
End Function  
'============================================================================================================
' Name : SubCheckHoliday
' Desc : Check Holiday
'============================================================================================================
Function SubCheckHoliday()
	Dim strFg, strType, strOrgId
	Dim strHoli_type, strHoliday_apply
    Dim strWhere, strDilig_dt, strEnd_Dilig_dt
    Dim iCnt, iHoliday_cnt
    Dim IntRetCD
    
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
	SubCheckHoliday =  False

    Call CommonQueryRs(" top 1 dept_cd "," hba010t a", " a.gazet_dt = (select MAX(gazet_dt) from hba010t where gazet_dt <= getdate()" &_
                                                                    " and emp_no = a.emp_no )" &_
                                                 " and emp_no = " & FilterVar(Request("txtEmp_no"), "''", "S") &_    

                                                 " and a.dept_cd is not null",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    strFg = Trim(Replace(lgF0, Chr(11), ""))
'Response.Write	"***strFg:"   & strFg

    If strFg = "" OR strFg = "X" Then
       Call CommonQueryRs(" dept_cd "," HAA010T "," emp_no = " & FilterVar(Request("txtEmp_no"), "''", "S"),lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	   strFg = Replace(lgF0, Chr(11), "")
    End If
	strWhere = " a.dept_cd =  " & FilterVar(strFg , "''", "S") & "" & _
	           " AND a.org_change_dt = (SELECT MAX(org_change_dt) " &_
	                                   "  FROM b_acct_dept " &_
	                                   " WHERE dept_cd = a.dept_cd " &_
	                                   "   AND org_change_dt <= getdate() ) AND a.cost_cd = b.cost_cd "  
                                    
	Call CommonQueryRs(" b.biz_area_cd "," b_acct_dept a, b_cost_center b ", strWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    strOrgId = Replace(lgF0, Chr(11), "")    
'Response.Write	"***strOrgId:"   & strOrgId    

	If IsNull(strOrgId) or strOrgId = "" or strOrgId = "X" Then
        Call DisplayMsgBox("800486", vbInformation, "", "", I_MKSCRIPT)     'Can not create(Demo code)        
        ObjectContext.SetAbort
        Call SetErrorStatus
		Exit Function
	End If

	strWhere = " a.emp_no = " & FilterVar(Request("txtEmp_no"), "''", "S") &_
	           " AND a.chang_dt = (SELECT MAX(chang_dt) " &_
	                            "  FROM hca040t " &_
	                            " WHERE emp_no = a.emp_no" &_
	                            "   AND chang_dt <= getdate() )" 
  
    Call CommonQueryRs(" a.wk_type "," HCA040T a ", strWhere, lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	strType = Replace(lgF0, Chr(11), "")
'Response.Write	"***strType:"   & strType   	
  	
	If (strType = "X" OR strType = "") Then strType = "0"

    strDilig_dt = UNIConvDate(Request("txtdilig_strt_dt"))
    strEnd_Dilig_dt = UNIConvDate(Request("txtdilig_end_dt"))
    iCnt = 0
    iHoliday_cnt = 0    
    '휴일적용여부가 'N'이고 해당일이 휴일이면 등록 불가 
    Call CommonQueryRs(" holiday_apply "," HCA010T "," dilig_cd = " & FilterVar(UCase(Request("txtdilig_cd")), "''", "S"),lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
'Response.Write	"***where strHoliday_apply:" & " dilig_cd = " & FilterVar(Trim(UCase(Request("txtdilig_cd"))),"''","S")    
    strHoliday_apply = Trim(Replace(lgF0, Chr(11), ""))
'Response.Write	"***strHoliday_apply:"   & strHoliday_apply
'Response.Write	"***strDilig_dt:"   & strDilig_dt 
'Response.Write	"***strEnd_Dilig_dt:"   & strEnd_Dilig_dt   
    '휴일적용여부가 'N'이고 기간내에 근태일이 모두 휴일이면 등록이 안되도록 함. 
	If strHoliday_apply = "N" Then 
        Do While strDilig_dt <= strEnd_Dilig_dt
            iCnt = iCnt +1

            '해당일이 휴일인지 평일인지 가져옴 
	        IntRetCD = CommonQueryRs(" holi_type "," HCA020T "," org_cd =  " & FilterVar(strOrgId , "''", "S") & "" & _
	                                 " and wk_type =  " & FilterVar(strType , "''", "S") & " and date =  " & FilterVar(strDilig_dt , "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
            If IntRetCD = True Then
           
	            strHoli_type = Trim(Replace(lgF0, Chr(11), ""))
'Response.Write	"***strHoli_type :"   & strHoli_type             	            
                If strHoli_type = "" OR strHoli_type = "X" Then
                    Call DisplayMsgBox("800453", vbInformation, "", "", I_MKSCRIPT)   ' 해당사원의 근무칼렌다 정보가 없습니다. 근무칼렌다를 먼저 생성하십시오.

                    ObjectContext.SetAbort
                    Call SetErrorStatus
            		Exit Function
                End If    
            Else
	        	Exit Function
            End If

	        If strHoli_type = "H" Then 
                iHoliday_cnt = iHoliday_cnt +1
            Else
                Exit Do
	        End If

            strDilig_dt = UNIDateAdd("D", 1, strDilig_dt, gAPDateFormat)
        Loop 
'Response.Write "****iCnt:" & iCnt & ",iHoliday_cnt:" & iHoliday_cnt &"*******"
        If iCnt = iHoliday_cnt Then
			Call CommonQueryRs(" DILIG_NM "," HCA010T ","  dilig_cd = " & FilterVar(UCase(Request("txtdilig_cd")), "''", "S"),lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)        
            Call DisplayMsgBox("800505", vbInformation, Replace(lgF0, Chr(11), ""), "", I_MKSCRIPT)           
            ObjectContext.SetAbort
            Call SetErrorStatus
         	Exit Function
	    End If
'Response.Write "***end SubCheckHoliday:" & SubCheckHoliday
  
    End If  
	SubCheckHoliday =  True	     
'Response.Write ",SubCheckHoliday:"	 & SubCheckHoliday
'Response.end	
End Function                     
 
'============================================================================================================
' Name : SubBizDelete
' Desc : Delete DB data
'============================================================================================================
Sub SubBizDelete()

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

	lgStrSQL = "DELETE  E11070T"
	lgStrSQL = lgStrSQL & " WHERE emp_no = "        & FilterVar(Request("txtEmp_no"), "''", "S")
	lgStrSQL = lgStrSQL & "   AND dilig_strt_dt = " & FilterVar(UNIConvDateCompanyToDB(Request("txtdilig_strt_dt"),NULL),"NULL","S")
	lgStrSQL = lgStrSQL & "   AND dilig_cd = "      & FilterVar(lgKeyStream(4), "''", "S")

	lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
    
    End Sub

'============================================================================================================
' Name : SubBizSaveCreate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveSingleCreate()

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    lgStrSQL = "INSERT INTO E11070T("
    lgStrSQL = lgStrSQL & " emp_no, "
    lgStrSQL = lgStrSQL & " dilig_strt_dt, "
    lgStrSQL = lgStrSQL & " dilig_end_dt, "
    lgStrSQL = lgStrSQL & " dilig_cd, "
    lgStrSQL = lgStrSQL & " remark, "
    lgStrSQL = lgStrSQL & " dilig_hh, "
    lgStrSQL = lgStrSQL & " dilig_mm, "    
    lgStrSQL = lgStrSQL & " app_emp_no, "
    lgStrSQL = lgStrSQL & " isrt_emp_no, "
    lgStrSQL = lgStrSQL & " updt_emp_no ) "
    lgStrSQL = lgStrSQL & " VALUES(" 
    lgStrSQL = lgStrSQL & FilterVar(Request("txtEmp_no"), "''", "S") & ","
    lgStrSQL = lgStrSQL & FilterVar(UNIConvDateCompanyToDB(Request("txtdilig_strt_dt"),NULL),"NULL","S")   & "," 
    lgStrSQL = lgStrSQL & FilterVar(UNIConvDateCompanyToDB(Request("txtdilig_end_dt"),NULL),"NULL","S")   & ","
    lgStrSQL = lgStrSQL & FilterVar(Request("txtdilig_cd"), "''", "S") & ","
    lgStrSQL = lgStrSQL & FilterVar(Request("txtremark"), "''", "S") & ","
    lgStrSQL = lgStrSQL & FilterVar(lgKeyStream(8), "''", "S") & ","    
    lgStrSQL = lgStrSQL & FilterVar(lgKeyStream(9), "''", "S")   & ","   
        
    lgStrSQL = lgStrSQL & FilterVar(UCase(Request("txtapp_emp_no")), "''", "S") & ","
    lgStrSQL = lgStrSQL & FilterVar(Request("txtEmp_no"), "''", "S") & ","
    lgStrSQL = lgStrSQL & FilterVar(Request("txtEmp_no"), "''", "S")
    lgStrSQL = lgStrSQL & ")"
'Response.Write lgStrSQL
'Response.End
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

    lgStrSQL = "UPDATE  E11070T"
    lgStrSQL = lgStrSQL & " SET   dilig_end_dt = " & FilterVar(UNIConvDateCompanyToDB(Request("txtdilig_end_dt"),NULL),"NULL","S") & ","
    lgStrSQL = lgStrSQL & "       remark = "        & FilterVar(Request("txtremark"), "''", "S") & ","
    lgStrSQL = lgStrSQL & "       dilig_hh = "        & FilterVar(lgKeyStream(8), "''", "S") & ","    
    lgStrSQL = lgStrSQL & "       dilig_mm = "        & FilterVar(lgKeyStream(9), "''", "S") & ","    
    
    lgStrSQL = lgStrSQL & "       app_emp_no = "    & FilterVar(Request("txtapp_emp_no"), "''", "S") & ","
    lgStrSQL = lgStrSQL & "       updt_emp_no = "   & FilterVar(Request("txtEmp_no"), "''", "S")
    lgStrSQL = lgStrSQL & " WHERE emp_no = "        & FilterVar(Request("txtEmp_no"), "''", "S")
    lgStrSQL = lgStrSQL & "   AND dilig_strt_dt = " & FilterVar(UNIConvDateCompanyToDB(Request("txtdilig_strt_dt"),NULL),"NULL","S")
    lgStrSQL = lgStrSQL & "   AND dilig_cd = "      & FilterVar(lgKeyStream(4), "''", "S")
    
'Response.Write lgStrSQL
'Response.End
    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
	Call SubHandleError("SU",lgObjConn,lgObjRs,Err)
    
End Sub

'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements(pMode,pCode)
    Dim iSelCount

    Select Case pMode 
      Case "R"
             lgStrSQL = "Select top 1 emp_no,dilig_strt_dt,dilig_cd,dilig_end_dt,app_emp_no,remark," 
             lgStrSQL = lgStrSQL & " dbo.ufn_H_GetEmpName(app_emp_no)  as app_name "
             lgStrSQL = lgStrSQL & " From  E11070T "
             lgStrSQL = lgStrSQL & " WHERE emp_no = " & pCode 	
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
			        Call DisplayMsgBox("183116", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
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
			        Call DisplayMsgBox("183116", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
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
