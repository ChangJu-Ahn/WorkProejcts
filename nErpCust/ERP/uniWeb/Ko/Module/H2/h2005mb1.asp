<%@ LANGUAGE=VBSCript TRANSACTION=Required%>

<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/IncServerAdoDb.asp" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->

<%
    call LoadBasisGlobalInf()
	Call loadInfTB19029B( "I", "H","NOCOOKIE","MB")
    Call HideStatusWnd                                                               '☜: Hide Processing message

    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '☜: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)
    lgKeyStream       = Split(Request("txtKeyStream"),gColSep)

    lgPrevNext        = Request("txtPrevNext")                                       '☜: "P"(Prev search) "N"(Next search)

    Call SubOpenDB(lgObjConn)                                                        '☜: Make a DB Connection

    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '☜: Query
             Call SubBizQuery()
        Case CStr(UID_M0002)                                                         '☜: Save,Update
             Call SubBizSave()
        Case CStr(UID_M0003)                                                         '☜: Delete
             Call SubBizDelete()
    End Select

    Call SubCloseDB(lgObjConn)                                                       '☜: Close DB Connection

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()
    Dim iKey1
    Dim strRoll_pstn
    Dim strPay_grd1
	dim strNat_cd
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    ' 성명으로 조회할 경우 
    ' 여러명일 경우 
     iKey1 = FilterVar(lgKeyStream(0), "''", "S")     ' 사번으로조회 

    call CommonQueryRs(" nat_cd "," HAA010T "," EMP_NO =  " & FilterVar(lgKeyStream(0), "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    strNat_cd = Replace(lgF0, Chr(11), "")  ' 주민번호 check를 위해서 

    Call SubMakeSQLStatements("R",iKey1)                                       '☜ : Make sql statements

    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then                    'If data not exists
       If lgPrevNext = "" Then
          Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
          Call SetErrorStatus()
       ElseIf lgPrevNext = "P" Then
          Call DisplayMsgBox("900011", vbInformation, "", "", I_MKSCRIPT)      '☜ : This is the starting data. 
          lgPrevNext = ""
          
          Call SubBizQuery()
       ElseIf lgPrevNext = "N" Then
          Call DisplayMsgBox("900012", vbInformation, "", "", I_MKSCRIPT)      '☜ : This is the ending data.
          lgPrevNext = ""
          Call SubBizQuery()
       End If
       
    Else
%>
<Script Language=vbscript>

    With Parent	
        ' Header (사원기본정보)
        .Frm1.txtEmp_no.Value  = "<%=ConvSPChars(lgObjRs("emp_no"))%>"
        .Frm1.txtName.Value  = "<%=ConvSPChars(lgObjRs("name"))%>"
        .Frm1.txtDept_nm.Value  = "<%=ConvSPChars(lgObjRs("dept_nm"))%>"
        .Frm1.txtRoll_pstn.Value  = "<%=ConvSPChars(lgObjRs("roll_pstn_nm"))%>"
        .Frm1.txtPay_grd.Value  = "<%=ConvSPChars(lgObjRs("pay_grd1_nm"))%>" & "-" & "<%=ConvSPChars(lgObjRs("pay_grd2"))%>"
        .Frm1.txtEntr_dt.text  = "<%=UniConvDateDbToCompany(lgObjRs("entr_dt"),"")%>"
        
        ' 보증보험 
        .Frm1.txtWarnt_insur_nm.Value  = "<%=ConvSPChars(lgObjRs("warnt_insur_nm"))%>"
        .Frm1.txtWarnt_insur_no.Value  = "<%=ConvSPChars(lgObjRs("warnt_insur_no"))%>"
        .Frm1.txtWarnt_insur_comp.Value  = "<%=ConvSPChars(lgObjRs("warnt_insur_comp"))%>"
        .Frm1.txtWarnt_amt.text  = "<%=UNINumClientFormat(lgObjRs("warnt_amt"), ggAmtOfMoney.DecPoint, 0)%>"
        .Frm1.txtWarnt_start.Text  = "<%=UniConvDateDbToCompany(lgObjRs("warnt_start"),"")%>"
        .Frm1.txtWarnt_end.Text  = "<%=UniConvDateDbToCompany(lgObjRs("warnt_end"),"")%>"
         ' 보증인1
        .Frm1.txtWarnt1_name.Value  = "<%=ConvSPChars(lgObjRs("warnt1_name"))%>"
        .Frm1.txtWarnt1_comp_nm.Value  = "<%=ConvSPChars(lgObjRs("warnt1_comp_nm"))%>"
        .Frm1.txtWarnt1_rel.Value  = "<%=ConvSPChars(lgObjRs("warnt1_rel"))%>"
        .Frm1.txtWarnt1_roll_pstn.Value  = "<%=ConvSPChars(lgObjRs("warnt1_roll_pstn"))%>"
    
        if Trim("<%=strNat_cd%>") ="KR" then
			if Trim("<%=lgObjRs("warnt1_res_no")%>") <>"" then
				.frm1.txtWarnt1_res_no.value = "<%=Mid(ConvSPChars(lgObjRs("warnt1_res_no")),1,6) & "-" & Mid(ConvSPChars(lgObjRs("warnt1_res_no")),7,7)%>" '주민번호 
			end if
		else
			.frm1.txtWarnt1_res_no.value = "<%=ConvSPChars(lgObjRs("warnt1_res_no"))%>" '주민번호 
		end if         
        .Frm1.txtwarnt1_incom_tax.text  = "<%=UNINumClientFormat(lgObjRs("warnt1_incom_tax"), ggAmtOfMoney.DecPoint, 0)%>"
        .Frm1.txtWarnt1_start.Text  = "<%=UniConvDateDbToCompany(lgObjRs("warnt1_start"),"")%>"
        .Frm1.txtWarnt1_end.Text  = "<%=UniConvDateDbToCompany(lgObjRs("warnt1_end"),"")%>"
        .Frm1.txtwarnt1_addr.Value  = "<%=ConvSPChars(lgObjRs("warnt1_addr"))%>"
        ' 보증인2
        .Frm1.txtWarnt2_name.Value  = "<%=ConvSPChars(lgObjRs("warnt2_name"))%>"
        .Frm1.txtWarnt2_comp_nm.Value  = "<%=ConvSPChars(lgObjRs("warnt2_comp_nm"))%>"
        .Frm1.txtWarnt2_rel.Value  = "<%=ConvSPChars(lgObjRs("warnt2_rel"))%>"
        .Frm1.txtWarnt2_roll_pstn.Value  = "<%=ConvSPChars(lgObjRs("warnt2_roll_pstn"))%>"
        if Trim("<%=strNat_cd%>") ="KR" then
			if Trim("<%=lgObjRs("warnt2_res_no")%>") <>"" then
				.frm1.txtWarnt2_res_no.value = "<%=Mid(ConvSPChars(lgObjRs("warnt2_res_no")),1,6) & "-" & Mid(ConvSPChars(lgObjRs("warnt2_res_no")),7,7)%>" '주민번호 
			end if
		else
			.frm1.txtWarnt2_res_no.value = "<%=ConvSPChars(lgObjRs("warnt2_res_no"))%>" '주민번호 
		end if          
        .Frm1.txtwarnt2_incom_tax.text  = "<%=UNINumClientFormat(lgObjRs("warnt2_incom_tax"), ggAmtOfMoney.DecPoint, 0)%>"
        .Frm1.txtWarnt2_start.Text  = "<%=UniConvDateDbToCompany(lgObjRs("warnt2_start"),"")%>"
        .Frm1.txtWarnt2_end.Text  = "<%=UniConvDateDbToCompany(lgObjRs("warnt2_end"),"")%>"
        .Frm1.txtwarnt2_addr.Value  = "<%=ConvSPChars(lgObjRs("warnt2_addr"))%>"

    End With          
</Script>       
<%     
    End If

    Call SubCloseRs(lgObjRs)                                                    '☜ : Release RecordSSet

End Sub	
'============================================================================================================
' Name : SubBizQuery
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

    lgStrSQL = "DELETE  HAA040T"
    lgStrSQL = lgStrSQL & " WHERE emp_no = " & FilterVar(lgKeyStream(0), "''", "S")                              ' 사번char(10)

    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
	Call SubHandleError("SD",lgObjConn,lgObjRs,Err)

End Sub

'============================================================================================================
' Name : SubBizSaveCreate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveSingleCreate()

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    lgStrSQL = "INSERT INTO HAA040T("
    lgStrSQL = lgStrSQL & " emp_no, "
    lgStrSQL = lgStrSQL & " warnt_insur_nm, "
    lgStrSQL = lgStrSQL & " warnt_insur_comp, "
    lgStrSQL = lgStrSQL & " warnt_amt, "
    lgStrSQL = lgStrSQL & " warnt_insur_no, "
    lgStrSQL = lgStrSQL & " warnt_start, "
    lgStrSQL = lgStrSQL & " warnt_end, "
    lgStrSQL = lgStrSQL & " warnt1_name, "
    lgStrSQL = lgStrSQL & " warnt1_rel, "
    lgStrSQL = lgStrSQL & " warnt1_roll_pstn, "
    lgStrSQL = lgStrSQL & " warnt1_start, "
    lgStrSQL = lgStrSQL & " warnt1_end, "
    lgStrSQL = lgStrSQL & " warnt1_res_no, "
    lgStrSQL = lgStrSQL & " warnt1_comp_nm, "
    lgStrSQL = lgStrSQL & " warnt1_addr, "
    lgStrSQL = lgStrSQL & " warnt1_incom_tax, "
    lgStrSQL = lgStrSQL & " warnt2_name, "
    lgStrSQL = lgStrSQL & " warnt2_rel, "
    lgStrSQL = lgStrSQL & " warnt2_roll_pstn, "
    lgStrSQL = lgStrSQL & " warnt2_start, "
    lgStrSQL = lgStrSQL & " warnt2_end, "
    lgStrSQL = lgStrSQL & " warnt2_res_no, "
    lgStrSQL = lgStrSQL & " warnt2_comp_nm, "
    lgStrSQL = lgStrSQL & " warnt2_addr, "
    lgStrSQL = lgStrSQL & " warnt2_incom_tax, "

    lgStrSQL = lgStrSQL & " isrt_emp_no, "
    lgStrSQL = lgStrSQL & " isrt_dt, "
    lgStrSQL = lgStrSQL & " updt_emp_no, "
    lgStrSQL = lgStrSQL & " updt_dt ) "

    lgStrSQL = lgStrSQL & " VALUES(" 
    lgStrSQL = lgStrSQL & FilterVar(Request("txtEmp_no"), "''", "S") & ","
    lgStrSQL = lgStrSQL & FilterVar(Request("txtwarnt_insur_nm"), "''", "S") & ","
    lgStrSQL = lgStrSQL & FilterVar(Request("txtwarnt_insur_comp"), "''", "S") & ","
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtwarnt_amt"),0) & ","
    lgStrSQL = lgStrSQL & FilterVar(Request("txtwarnt_insur_no"), "''", "S") & ","
    lgStrSQL = lgStrSQL & FilterVar(UNIConvDateCompanyToDB(Request("txtWarnt_start"),NULL),"NULL","S") & ","
    lgStrSQL = lgStrSQL & FilterVar(UNIConvDateCompanyToDB(Request("txtWarnt_end"),NULL),"NULL","S") & ","

    lgStrSQL = lgStrSQL & FilterVar(Request("txtwarnt1_name"), "''", "S") & ","
    lgStrSQL = lgStrSQL & FilterVar(Request("txtwarnt1_rel"), "''", "S") & ","
    lgStrSQL = lgStrSQL & FilterVar(Request("txtwarnt1_roll_pstn"), "''", "S") & ","
    lgStrSQL = lgStrSQL & FilterVar(UNIConvDateCompanyToDB(Request("txtwarnt1_start"),NULL),"NULL","S") & ","
    lgStrSQL = lgStrSQL & FilterVar(UNIConvDateCompanyToDB(Request("txtwarnt1_end"),NULL),"NULL","S") & ","
    lgStrSQL = lgStrSQL & FilterVar(Request("txtwarnt1_res_no"), "''", "S") & ","
    lgStrSQL = lgStrSQL & FilterVar(Request("txtwarnt1_comp_nm"), "''", "S") & ","
    lgStrSQL = lgStrSQL & FilterVar(Request("txtwarnt1_addr"), "''", "S") & ","
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtwarnt1_incom_tax"),0) & ","

    lgStrSQL = lgStrSQL & FilterVar(Request("txtwarnt2_name"), "''", "S") & ","
    lgStrSQL = lgStrSQL & FilterVar(Request("txtwarnt2_rel"), "''", "S") & ","
    lgStrSQL = lgStrSQL & FilterVar(Request("txtwarnt2_roll_pstn"), "''", "S") & ","
    lgStrSQL = lgStrSQL & FilterVar(UNIConvDateCompanyToDB(Request("txtwarnt2_start"),NULL),"NULL","S") & ","
    lgStrSQL = lgStrSQL & FilterVar(UNIConvDateCompanyToDB(Request("txtwarnt2_end"),NULL),"NULL","S") & ","
    lgStrSQL = lgStrSQL & FilterVar(Request("txtwarnt2_res_no"), "''", "S") & ","
    lgStrSQL = lgStrSQL & FilterVar(Request("txtwarnt2_comp_nm"), "''", "S") & ","
    lgStrSQL = lgStrSQL & FilterVar(Request("txtwarnt2_addr"), "''", "S") & ","
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtwarnt2_incom_tax"),0) & ","

    lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''", "S") & ","
    lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime, "''", "S") & ","
    lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''", "S") & ","
    lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime, "''", "S")


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

    lgStrSQL = "UPDATE  HAA040T"
    lgStrSQL = lgStrSQL & " SET " 
    lgStrSQL = lgStrSQL & " Warnt_insur_nm = " & FilterVar(Request("txtWarnt_insur_nm"), "''", "S") & ","
    lgStrSQL = lgStrSQL & " Warnt_insur_no = " & FilterVar(Request("txtWarnt_insur_no"), "''", "S") & ","
    lgStrSQL = lgStrSQL & " Warnt_insur_comp = " & FilterVar(Request("txtWarnt_insur_comp"), "''", "S") & ","
    lgStrSQL = lgStrSQL & " Warnt_amt = " & UNIConvNum(Request("txtWarnt_amt"),0) & ","
    lgStrSQL = lgStrSQL & " Warnt_start = " & FilterVar(UNIConvDateCompanyToDB(Request("txtWarnt_start"),NULL),"NULL","S") & ","
    lgStrSQL = lgStrSQL & " Warnt_end = " & FilterVar(UNIConvDateCompanyToDB(Request("txtWarnt_end"),NULL),"NULL","S") & ","
    ' 보증인1
    lgStrSQL = lgStrSQL & " Warnt1_name = " & FilterVar(Request("txtWarnt1_name"), "''", "S") & ","
    lgStrSQL = lgStrSQL & " Warnt1_comp_nm = " & FilterVar(Request("txtWarnt1_comp_nm"), "''", "S") & ","
    lgStrSQL = lgStrSQL & " Warnt1_rel = " & FilterVar(Request("txtWarnt1_rel"), "''", "S") & ","
    lgStrSQL = lgStrSQL & " Warnt1_roll_pstn = " & FilterVar(Request("txtWarnt1_roll_pstn"), "''", "S") & ","
    lgStrSQL = lgStrSQL & " Warnt1_res_no = " & FilterVar(Request("txtWarnt1_res_no"), "''", "S") & ","
    lgStrSQL = lgStrSQL & " warnt1_incom_tax = " & UNIConvNum(Request("txtwarnt1_incom_tax"),0) & ","
    lgStrSQL = lgStrSQL & " warnt1_addr = " & FilterVar(Replace(Request("txtwarnt1_addr"),",",""), "''", "S") & ","
    lgStrSQL = lgStrSQL & " Warnt1_start = " & FilterVar(UNIConvDateCompanyToDB(Request("txtWarnt1_start"),NULL),"NULL","S") & ","
    lgStrSQL = lgStrSQL & " Warnt1_end = " & FilterVar(UNIConvDateCompanyToDB(Request("txtWarnt1_end"),NULL),"NULL","S") & ","
    ' 보증인2
    lgStrSQL = lgStrSQL & " Warnt2_name = " & FilterVar(Request("txtWarnt2_name"), "''", "S") & ","
    lgStrSQL = lgStrSQL & " Warnt2_comp_nm = " & FilterVar(Request("txtWarnt2_comp_nm"), "''", "S") & ","
    lgStrSQL = lgStrSQL & " Warnt2_rel = " & FilterVar(Request("txtWarnt2_rel"), "''", "S") & ","
    lgStrSQL = lgStrSQL & " Warnt2_roll_pstn = " & FilterVar(Request("txtWarnt2_roll_pstn"), "''", "S") & ","
    lgStrSQL = lgStrSQL & " Warnt2_res_no = " & FilterVar(Request("txtWarnt2_res_no"), "''", "S") & ","
    lgStrSQL = lgStrSQL & " warnt2_incom_tax = " & UNIConvNum(Request("txtwarnt2_incom_tax"),0) & ","
    lgStrSQL = lgStrSQL & " Warnt2_start = " & FilterVar(UNIConvDateCompanyToDB(Request("txtWarnt2_start"),NULL),"NULL","S") & ","
    lgStrSQL = lgStrSQL & " Warnt2_end = " & FilterVar(UNIConvDateCompanyToDB(Request("txtWarnt2_end"),NULL),"NULL","S") & ","
    lgStrSQL = lgStrSQL & " warnt2_addr = " & FilterVar(Replace(Request("txtwarnt2_addr"),",",""), "''", "S") & ","

    lgStrSQL = lgStrSQL & " updt_emp_no = " & FilterVar(gUsrId, "''", "S") & ","
    lgStrSQL = lgStrSQL & " updt_dt = " & FilterVar(GetSvrDateTime, "''", "S")
    lgStrSQL = lgStrSQL & " WHERE emp_no = " & FilterVar(lgKeyStream(0), "''", "S")                              ' 사번char(10)

    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
	Call SubHandleError("SU",lgObjConn,lgObjRs,Err)
    
End Sub

'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements(pMode,pCode)

    Select Case pMode 
      Case "R"
             Select Case  lgPrevNext 
	                 Case ""
                           lgStrSQL = "Select a.emp_no, a.warnt_insur_nm, a.warnt_insur_no, a.warnt_insur_comp, " 
                           lgStrSQL = lgStrSQL & " a.warnt_amt, a.warnt_start, a.warnt_end, a.warnt1_name, "
                           lgStrSQL = lgStrSQL & " a.warnt1_comp_nm, a.warnt1_rel, a.warnt1_roll_pstn, a.warnt1_res_no, "
                           lgStrSQL = lgStrSQL & " a.warnt1_incom_tax, a.warnt1_start, a.warnt1_end, a.warnt1_addr, "
                           lgStrSQL = lgStrSQL & " a.warnt2_name, a.warnt2_comp_nm, a.warnt2_rel, a.warnt2_roll_pstn, "
                           lgStrSQL = lgStrSQL & " a.warnt2_res_no, a.warnt2_incom_tax, a.warnt2_start, a.warnt2_end, "
                           lgStrSQL = lgStrSQL & " a.warnt2_addr, b.name, b.dept_nm, b.entr_dt, b.pay_grd2, "
                           lgStrSQL = lgStrSQL & " dbo.ufn_GetCodeName(" & FilterVar("H0002", "''", "S") & ", b.roll_pstn) roll_pstn_nm, "
                           lgStrSQL = lgStrSQL & " dbo.ufn_GetCodeName(" & FilterVar("H0001", "''", "S") & ", b.pay_grd1) pay_grd1_nm "
                           lgStrSQL = lgStrSQL & " From  HAA040T a, HAA010t b "
                           lgStrSQL = lgStrSQL & " WHERE a.emp_no = " & pCode 	
                           lgStrSQL = lgStrSQL & "   AND a.emp_no = b.emp_no"

                     Case "P"
                           lgStrSQL = "Select TOP 1 a.emp_no, a.warnt_insur_nm, a.warnt_insur_no, a.warnt_insur_comp, " 
                           lgStrSQL = lgStrSQL & " a.warnt_amt, a.warnt_start, a.warnt_end, a.warnt1_name, "
                           lgStrSQL = lgStrSQL & " a.warnt1_comp_nm, a.warnt1_rel, a.warnt1_roll_pstn, a.warnt1_res_no, "
                           lgStrSQL = lgStrSQL & " a.warnt1_incom_tax, a.warnt1_start, a.warnt1_end, a.warnt1_addr, "
                           lgStrSQL = lgStrSQL & " a.warnt2_name, a.warnt2_comp_nm, a.warnt2_rel, a.warnt2_roll_pstn, "
                           lgStrSQL = lgStrSQL & " a.warnt2_res_no, a.warnt2_incom_tax, a.warnt2_start, a.warnt2_end, "
                           lgStrSQL = lgStrSQL & " a.warnt2_addr, b.name, b.dept_nm, b.entr_dt, b.pay_grd2, "
                           lgStrSQL = lgStrSQL & " dbo.ufn_GetCodeName(" & FilterVar("H0002", "''", "S") & ", b.roll_pstn) roll_pstn_nm, "
                           lgStrSQL = lgStrSQL & " dbo.ufn_GetCodeName(" & FilterVar("H0001", "''", "S") & ", b.pay_grd1) pay_grd1_nm "
                           lgStrSQL = lgStrSQL & " From  HAA040T a, HAA010t b "
                           lgStrSQL = lgStrSQL & " WHERE a.emp_no < " & pCode 	
                           lgStrSQL = lgStrSQL & "   AND a.emp_no = b.emp_no"
                           lgStrSQL = lgStrSQL & " ORDER BY a.emp_no DESC "

                     Case "N"
                           lgStrSQL = "Select TOP 1 a.emp_no, a.warnt_insur_nm, a.warnt_insur_no, a.warnt_insur_comp, " 
                           lgStrSQL = lgStrSQL & " a.warnt_amt, a.warnt_start, a.warnt_end, a.warnt1_name, "
                           lgStrSQL = lgStrSQL & " a.warnt1_comp_nm, a.warnt1_rel, a.warnt1_roll_pstn, a.warnt1_res_no, "
                           lgStrSQL = lgStrSQL & " a.warnt1_incom_tax, a.warnt1_start, a.warnt1_end, a.warnt1_addr, "
                           lgStrSQL = lgStrSQL & " a.warnt2_name, a.warnt2_comp_nm, a.warnt2_rel, a.warnt2_roll_pstn, "
                           lgStrSQL = lgStrSQL & " a.warnt2_res_no, a.warnt2_incom_tax, a.warnt2_start, a.warnt2_end, "
                           lgStrSQL = lgStrSQL & " a.warnt2_addr, b.name, b.dept_nm, b.entr_dt, b.pay_grd2, "
                           lgStrSQL = lgStrSQL & " dbo.ufn_GetCodeName(" & FilterVar("H0002", "''", "S") & ", b.roll_pstn) roll_pstn_nm, "
                           lgStrSQL = lgStrSQL & " dbo.ufn_GetCodeName(" & FilterVar("H0001", "''", "S") & ", b.pay_grd1) pay_grd1_nm "
                           lgStrSQL = lgStrSQL & " From  HAA040T a, HAA010t b "
                           lgStrSQL = lgStrSQL & " WHERE a.emp_no > " & pCode 	
                           lgStrSQL = lgStrSQL & "   AND a.emp_no = b.emp_no"
                           lgStrSQL = lgStrSQL & " ORDER BY a.emp_no ASC "
                           
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
        Case "SC"
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
       Case "<%=UID_M0001%>"
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DBQueryOk        
          End If   
       Case "<%=UID_M0002%>"
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DBSaveOk
          End If   
       Case "<%=UID_M0003%>"
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DbDeleteOk
          End If   
    End Select    
       
</Script>	
