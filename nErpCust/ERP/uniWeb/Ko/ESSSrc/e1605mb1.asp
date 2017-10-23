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

    iRet = EmpBaseDiligAuthCheck(lgKeyStream(0),lgKeyStream(1),lgKeyStream(2),lgKeyStream(7),lgPrevNext,Emp_no,Name,roll_pstn,dept_nm,resent_promote_dt,group_entr_dt,entr_dt)
 
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
	else 
		
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
		Response.End
	end if	
	
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
	    iKey1 = iKey1 & " AND convert(varchar(10),dilig_strt_dt,20) =  " & FilterVar( lgKeyStream(4), "''", "S") & ""
		iKey1 = iKey1 & " AND dilig_cd =  " & FilterVar(lgKeyStream(3), "''", "S") & ""
		iKey1 = iKey1 & " AND dilig_cd in (select dilig_cd from hca010t where  dilig_type=2)"   

	    Call SubMakeSQLStatements("R",iKey1)                                       '☜ : Make sql statements
	    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
	        lgStrPrevKeyIndex = ""
	        Call SetErrorStatus()
	    Else
	%>
	<Script Language=vbscript>
	       With Parent
	            .frm1.txtdilig_cd.value      = "<%=ConvSPChars(lgObjRs("dilig_cd"))%>"
	            .frm1.txtremark.value        = "<%=ConvSPChars(lgObjRs("remark"))%>"
	            .frm1.txtapp_emp_no.value    = "<%=ConvSPChars(lgObjRs("app_emp_no"))%>"
	            .frm1.txtapp_name.value      = "<%=ConvSPChars(lgObjRs("app_name"))%>"
<%

Dim iDate, iHour, iMinute
iDate =  left(lgObjRs("dilig_strt_dt"),10)
iHour =  int(mid(lgObjRs("dilig_strt_dt"),12,2))
iMinute =  int(mid(lgObjRs("dilig_strt_dt"),15,2))

%>	            
	            .frm1.txtFrDilig_dt.value      = "<%=ConvSPChars(iDate)%>"
	            .frm1.txtFrDilig_hour.value    = "<%=ConvSPChars(iHour)%>"
	            .frm1.txtFrDilig_min.value     = "<%=ConvSPChars(iMinute)%>"
<%

iDate =  left(lgObjRs("DILIG_END_DT"),10)
iHour =  int(mid(lgObjRs("DILIG_END_DT"),12,2))
iMinute =  int(mid(lgObjRs("DILIG_END_DT"),15,2))

%>	            
	            .frm1.txtToDilig_dt.value      = "<%=ConvSPChars(iDate)%>"
	            .frm1.txtToDilig_hour.value      = "<%=ConvSPChars(iHour)%>"
	            .frm1.txtToDilig_min.value      = "<%=ConvSPChars(iMinute)%>"

	            .frm1.txtDilig_hour.value      = "<%=ConvSPChars(lgObjRs("dilig_hh"))%>"
	            .frm1.txtDilig_min.value      = "<%=ConvSPChars(lgObjRs("dilig_mm"))%>"
	            call .CalHHMM()
	       End With         
	</Script>       
	<%  

	    End If
    end if    

    Call SubCloseRs(lgObjRs)
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
      
    lgIntFlgMode = CInt(Request("txtFlgMode"))  		'☜: Read Operayion Mode (CREATE, UPDATE)
    Select Case lgIntFlgMode		
        Case  OPMD_CMODE															  '☜ : Create  Mode                                                   
		if  SubCheckHoliday() = False then	'휴일적용근태 여부를 체크한다.
               lgErrorStatus = "YES"
				Exit Sub
			end if   	       
			'기간근태(e11070t)에서 기간에 (중복일자)속했는지를 check 한다 
			'Call CommonQueryRs(" isnull(count(emp_no),0) "," e11070t ","  emp_no =  " & FilterVar( lgKeyStream(0), "''", "S") & " AND dilig_strt_dt = " & FilterVar(UNIConvDateCompanyToDB(lgKeyStream(4),NULL), "''", "S") & " and dilig_cd='" & lgKeyStream(3) & "' and dilig_cd in (select dilig_cd from hca010t where  dilig_type=2)" ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

			'If Trim(Replace(lgF0,Chr(11),""))= 0 then
			'Else
		  '      Call DisplayMsgBox("800067", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
      '          lgErrorStatus = "YES"
			'    Exit sub                                    
			'End if

			'기간근태(hca060t)에서 기간에 (중복일자)속했는지를 check 한다.
	        Call CommonQueryRs(" isnull(count(emp_no),0) "," hca060t ","  emp_no =  " & FilterVar( lgKeyStream(0), "''", "S") & " AND dilig_dt =  " & FilterVar(UNIConvDateCompanyToDB(lgKeyStream(4),NULL), "''", "S") & " and dilig_cd='" & lgKeyStream(3) & "' and dilig_cd in (select dilig_cd from hca010t where  dilig_type=2)" ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
			If Trim(Replace(lgF0,Chr(11),"")) <> 0 then

		        Call DisplayMsgBox("800067", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
			%>
			<Script Language=vbscript>
			        parent.frm1.txtDilig_dt.focus()
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
		strValidDt = UniConvDateToYYYYMMDD(lgKeyStream(4),gServerDateFormat,"")
	
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
' Name : SubBizDelete
' Desc : Delete DB data
'============================================================================================================
Sub SubBizDelete()

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status


'	lgStrSQL = "Select app_yn"
'	lgStrSQL = lgStrSQL & " From  E11070T "
'	lgStrSQL = lgStrSQL & " WHERE emp_no		= " & FilterVar(Request("txtEmp_no"),"''","S") 	
'    lgStrSQL = lgStrSQL & "   AND dilig_strt_dt = " & FilterVar(UNIConvDateCompanyToDB(Request("txtdilig_strt_dt"),NULL),"NULL","S")
'    lgStrSQL = lgStrSQL & "   AND dilig_cd		= " & FilterVar(Request("txtdilig_cd"),"''","S")
'
'	If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
'        Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
'        Call SetErrorStatus()
'	Else

		'''승인완료건에 대한 삭제시도여부 확인 
'			If ConvSPChars(lgObjRs("app_yn")) = "Y" Then
'			Else

				lgStrSQL = "DELETE  E11070T"
				lgStrSQL = lgStrSQL & " WHERE emp_no = "        & FilterVar(Request("txtEmp_no"), "''", "S")
				lgStrSQL = lgStrSQL & "   AND convert(varchar(10),dilig_strt_dt,20) = " & FilterVar(UNIConvDateCompanyToDB(Request("txtFrDilig_dt"),NULL),"NULL","S")
				lgStrSQL = lgStrSQL & "   AND dilig_cd = "      & FilterVar(lgKeyStream(3), "''", "S")

				lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords

'			End If
 '   End If
   
    End Sub

'============================================================================================================
' Name : SubBizSaveCreate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveSingleCreate()

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

			lgStrSQL = "DELETE  E11070T"
			lgStrSQL = lgStrSQL & " WHERE emp_no = "        & FilterVar(lgKeyStream(0), "''", "S")
			lgStrSQL = lgStrSQL & "   AND convert(varchar(10),dilig_strt_dt,20) = " & FilterVar(UNIConvDateCompanyToDB(Request("txtFrDilig_dt"),NULL),"NULL","S")
			lgStrSQL = lgStrSQL & "   AND dilig_cd = "      & FilterVar(Request("txtDilig_cd"), "''", "S")

			lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords

    lgStrSQL = "INSERT INTO E11070T("
    lgStrSQL = lgStrSQL & " emp_no, "
    lgStrSQL = lgStrSQL & " dilig_strt_dt, "
    lgStrSQL = lgStrSQL & " DILIG_END_DT, "    
    lgStrSQL = lgStrSQL & " dilig_cd, "
    lgStrSQL = lgStrSQL & " remark, "
    lgStrSQL = lgStrSQL & " app_emp_no, "
    lgStrSQL = lgStrSQL & " isrt_emp_no, "
    lgStrSQL = lgStrSQL & " updt_emp_no ,dilig_hh,dilig_mm) "
    lgStrSQL = lgStrSQL & " VALUES(" 
    lgStrSQL = lgStrSQL & FilterVar(lgKeyStream(0), "''", "S") & ","

    lgStrSQL = lgStrSQL & FilterVar(Request("txtDilig_STRT_dt"), "''", "S")  & ","
    lgStrSQL = lgStrSQL & FilterVar(Request("txtDilig_END_dt"), "''", "S")  & ","   
    
    lgStrSQL = lgStrSQL & FilterVar(Request("txtDilig_cd"), "''", "S") & ","
    lgStrSQL = lgStrSQL & FilterVar(Request("txtremark"), "''", "S") & ","
    lgStrSQL = lgStrSQL & FilterVar(UCase(Request("txtapp_emp_no")), "''", "S") & ","
    lgStrSQL = lgStrSQL & FilterVar(lgKeyStream(7), "''", "S") & ","
    lgStrSQL = lgStrSQL & FilterVar(lgKeyStream(7), "''", "S") & ","    
    lgStrSQL = lgStrSQL & FilterVar(lgKeyStream(5), "''", "S") & ","    
    lgStrSQL = lgStrSQL & FilterVar(lgKeyStream(6), "''", "S")     
    lgStrSQL = lgStrSQL & ")"

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

			lgStrSQL = "DELETE  E11070T"
			lgStrSQL = lgStrSQL & " WHERE emp_no = "        & FilterVar(Request("txtEmp_no"), "''", "S")
			lgStrSQL = lgStrSQL & "   AND convert(varchar(10),dilig_strt_dt,20) = " & FilterVar(UNIConvDateCompanyToDB(Request("txtFrDilig_dt"),NULL),"NULL","S")
			lgStrSQL = lgStrSQL & "   AND dilig_cd = "      & FilterVar(lgKeyStream(3), "''", "S")

			lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords

    lgStrSQL = "INSERT INTO E11070T("
    lgStrSQL = lgStrSQL & " emp_no, "
    lgStrSQL = lgStrSQL & " dilig_strt_dt, "
    lgStrSQL = lgStrSQL & " DILIG_END_DT, "    
    lgStrSQL = lgStrSQL & " dilig_cd, "
    lgStrSQL = lgStrSQL & " remark, "
    lgStrSQL = lgStrSQL & " app_emp_no, "
    lgStrSQL = lgStrSQL & " isrt_emp_no, "
    lgStrSQL = lgStrSQL & " updt_emp_no ,dilig_hh,dilig_mm) "
    lgStrSQL = lgStrSQL & " VALUES(" 
    lgStrSQL = lgStrSQL & FilterVar(lgKeyStream(0), "''", "S") & ","

    lgStrSQL = lgStrSQL & FilterVar(Request("txtDilig_STRT_dt"), "''", "S")  & ","
    lgStrSQL = lgStrSQL & FilterVar(Request("txtDilig_END_dt"), "''", "S")  & ","   
    
    lgStrSQL = lgStrSQL & FilterVar(lgKeyStream(3), "''", "S") & ","
    lgStrSQL = lgStrSQL & FilterVar(Request("txtremark"), "''", "S") & ","
    lgStrSQL = lgStrSQL & FilterVar(UCase(Request("txtapp_emp_no")), "''", "S") & ","
    lgStrSQL = lgStrSQL & FilterVar(lgKeyStream(7), "''", "S") & ","
    lgStrSQL = lgStrSQL & FilterVar(lgKeyStream(7), "''", "S") & ","    
    lgStrSQL = lgStrSQL & FilterVar(lgKeyStream(5), "''", "S") & ","    
    lgStrSQL = lgStrSQL & FilterVar(lgKeyStream(6), "''", "S")     
    lgStrSQL = lgStrSQL & ")"

    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
	Call SubHandleError("SC",lgObjConn,lgObjRs,Err)
 
End Sub

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
' 입력된일자의 근무당시의 부서코드검색 / 2002.04.08 송봉규 
    Call CommonQueryRs(" top 1 dept_cd "," hba010t a", " a.gazet_dt = (select MAX(gazet_dt) from hba010t where gazet_dt <= getdate()" &_
                                                                    " and emp_no = a.emp_no )" &_
                                                 " and emp_no = " & FilterVar(Request("txtEmp_no"), "''", "S") &_    
                                                 " and a.dept_cd is not null",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    strFg = Trim(Replace(lgF0, Chr(11), ""))
'Response.Write	"***strFg:"   & strFg                                              
'Response.End
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

    strDilig_dt = UNIConvDate(Request("txtFrDilig_dt"))
    iCnt = 0
    iHoliday_cnt = 0    
    '휴일적용여부가 'N'이고 해당일이 휴일이면 등록 불가 2002.11.06 by sbk 
    Call CommonQueryRs(" holiday_apply "," HCA010T "," dilig_cd = " & FilterVar(UCase(Request("txtdilig_cd")), "''", "S"),lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
'Response.Write	"***where strHoliday_apply:" & " dilig_cd = " & FilterVar(Trim(UCase(Request("txtdilig_cd"))),"''","S")    
    strHoliday_apply = Trim(Replace(lgF0, Chr(11), ""))
    
'Response.Write	"***strHoliday_apply:"   & strHoliday_apply

    '휴일적용여부가 'N'이고 기간내에 근태일이 모두 휴일이면 등록이 안되도록 함. 2002.11.08 by sbk
    If strHoliday_apply = "N" Then
            '해당일이 휴일인지 평일인지 가져옴 2002.11.06 by sbk 
	        IntRetCD = CommonQueryRs(" holi_type "," HCA020T "," org_cd =  " & FilterVar(strOrgId , "''", "S") & "" & _
	                                 " and wk_type =  " & FilterVar(strType , "''", "S") & " and date =  " & FilterVar(strDilig_dt , "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)	                                 
            If IntRetCD = True Then

	            strHoli_type = Trim(Replace(lgF0, Chr(11), ""))
'Response.Write " org_cd = '" & strOrgId & "'" & " and wk_type = '" & strType & "' and date = '" & strDilig_dt & "'"	            

'Response.Write	"***strHoli_type :"   & strHoli_type             	            
'Response.End	
                If strHoli_type = "" OR strHoli_type = "X" Then
'*************메세지처리 바꾸기 
                    Call DisplayMsgBox("800453", vbInformation, "", "", I_MKSCRIPT)   ' 해당사원의 근무칼렌다 정보가 없습니다. 근무칼렌다를 먼저 생성하십시오.

                    ObjectContext.SetAbort
                    Call SetErrorStatus
            		Exit Function
				elseif strHoli_type = "H" then
					Call CommonQueryRs(" DILIG_NM "," HCA010T ","  dilig_cd = " & FilterVar(UCase(Request("txtdilig_cd")), "''", "S"),lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)        
					Call DisplayMsgBox("800505", vbInformation, Replace(lgF0, Chr(11), ""), "", I_MKSCRIPT)           
					ObjectContext.SetAbort
					Call SetErrorStatus
         			Exit Function
				End If          			
			End If  				            		
'Response.Write "****iCnt:" & iCnt & ",iHoliday_cnt:" & iHoliday_cnt &"*******"
'Response.Write "***end SubCheckHoliday:" & SubCheckHoliday
  
    End If  
	SubCheckHoliday =  True	      
	
End Function                     

'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements(pMode,pCode)
    Dim iSelCount

    Select Case pMode 
      Case "R"
             Select Case  lgPrevNext 
                Case ""
                      lgStrSQL = "Select top 1 a.emp_no,convert(varchar(20),dilig_strt_dt,20) dilig_strt_dt,dilig_hh,dilig_mm,dilig_cd,app_emp_no,remark," 
                      lgStrSQL = lgStrSQL & " dbo.ufn_H_GetEmpName(app_emp_no)  as app_name,convert(varchar(20),DILIG_END_DT,20) DILIG_END_DT "
                      lgStrSQL = lgStrSQL & " From  E11070T a, haa010t "
                      lgStrSQL = lgStrSQL & " WHERE a.emp_no = " & pCode 	
                      
                Case "P"
                      lgStrSQL = "Select top 1 a.emp_no,convert(varchar(20),dilig_strt_dt,20) dilig_strt_dt,dilig_hh,dilig_mm,dilig_cd,app_emp_no,remark," 
                      lgStrSQL = lgStrSQL & " dbo.ufn_H_GetEmpName(app_emp_no)  as app_name,convert(varchar(20),DILIG_END_DT,20) DILIG_END_DT "
                      lgStrSQL = lgStrSQL & " From  E11070T a, haa010t "
                      lgStrSQL = lgStrSQL & " WHERE a.emp_no = " & pCode 	
                     
                Case "N"
                      lgStrSQL = "Select top 1 a.emp_no,convert(varchar(20),dilig_strt_dt,20) dilig_strt_dt,dilig_hh,dilig_mm,dilig_cd,app_emp_no,remark," 
                      lgStrSQL = lgStrSQL & " dbo.ufn_H_GetEmpName(app_emp_no)  as app_name,convert(varchar(20),DILIG_END_DT,20) DILIG_END_DT "
                      lgStrSQL = lgStrSQL & " From  E11070T a, haa010t "
                      lgStrSQL = lgStrSQL & " WHERE a.emp_no = " & pCode 	
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
