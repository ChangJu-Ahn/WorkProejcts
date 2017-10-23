<%@ LANGUAGE=VBSCript TRANSACTION=Required%>

<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/IncServerAdoDb.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%
    call LoadBasisGlobalInf()
    Call LoadInfTB19029B("I", "H","NOCOOKIE","MB")    
    Call HideStatusWnd                                                               '☜: Hide Processing message

   
    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '☜: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)
    lgKeyStream       = Split(Request("txtKeyStream"),gColSep)
	lgtempflgchk      = Request("temp_flg_chk")
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

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    iKey1 = FilterVar(lgKeyStream(0), "''", "S")     ' 사번으로조회 
    iKey1 = iKey1 & " AND internal_cd LIKE  " & FilterVar(lgKeyStream(1) & "%", "''", "S") & ""
	
    Call SubMakeSQLStatements("R",iKey1)                                   '☜ : Make sql statements
	
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then                    'If data not exists
       If lgPrevNext = "" Then
          Call DisplayMsgBox("800048", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
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

            .Frm1.txtEmp_no1.Value  = "<%=ConvSPChars(lgObjRs("emp_no"))%>"                   'Set condition area
            .Frm1.txtName1.Value  = "<%=ConvSPChars(lgObjRs("name"))%>"

            .Frm1.txtEmp_no.Value  = "<%=ConvSPChars(lgObjRs("emp_no"))%>"                   'Set condition area
            .Frm1.txtName.Value  = "<%=ConvSPChars(lgObjRs("name"))%>"
                
            .frm1.txthanja_name.value = "<%=ConvSPChars(lgObjRs("hanja_name"))%>"         
            .frm1.txteng_name.value = "<%=ConvSPChars(lgObjRs("eng_name"))%>"         
            .frm1.txtcomp_cd.value = "<%=ConvSPChars(lgObjRs("comp_cd"))%>" 
            .frm1.txtcomp_cd_nm.value = "<%=ConvSPChars(lgObjRs("comp_cd_nm"))%>" '회사 
                     
            .frm1.txtsect_cd.value    = "<%=ConvSPChars(lgObjRs("sect_cd"))%>"          
            .frm1.txtsect_cd_nm.value = "<%=ConvSPChars(lgObjRs("sect_cd_nm"))%>"  '근무구역 FuncCodeName(1, "H0035", lgObjRs("sect_cd"))
            .frm1.txtwk_area_cd.value    = "<%=ConvSPChars(lgObjRs("wk_area_cd"))%>"       
            .frm1.txtwk_area_cd_nm.value = "<%=ConvSPChars(lgObjRs("wk_area_cd_nm"))%>" '근무지 FuncCodeName(1, "H0036", lgObjRs("wk_area_cd"))
            .frm1.txtdept_cd.value      = "<%=ConvSPChars(lgObjRs("dept_cd"))%>"          
            .frm1.txtroll_pstn.value    = "<%=ConvSPChars(lgObjRs("roll_pstn"))%>"        
            .frm1.txtRoll_pstn_nm.value = "<%=ConvSPChars(lgObjRs("roll_pstn_nm"))%>" '직위 FuncCodeName(1, "H0002", lgObjRs("Roll_pstn"))
            .frm1.txtocpt_type.value    = "<%=ConvSPChars(lgObjRs("ocpt_type"))%>"
            .frm1.txtocpt_type_nm.value = "<%=ConvSPChars(lgObjRs("ocpt_type_nm"))%>"  'FuncCodeName(1, "H0003", lgObjRs("ocpt_type"))
            .frm1.txtfunc_cd.value    = "<%=ConvSPChars(lgObjRs("func_cd"))%>"          
            .frm1.txtFunc_cd_nm.value = "<%=ConvSPChars(lgObjRs("func_cd_nm"))%>" '직무 FuncCodeName(1, "H0004", lgObjRs("Func_cd"))
            .frm1.txtrole_cd.value    = "<%=ConvSPChars(lgObjRs("role_cd"))%>"          
            .frm1.txtRole_cd_nm.value = "<%=ConvSPChars(lgObjRs("role_cd_nm"))%>" '직책 FuncCodeName(1, "H0026", lgObjRs("Role_cd"))
            .frm1.txtpay_grd1.value    = "<%=ConvSPChars(lgObjRs("pay_grd1"))%>"         
            .frm1.txtpay_grd1_nm.value = "<%=ConvSPChars(lgObjRs("pay_grd1_nm"))%>"  '급호 FuncCodeName(1, "H0001", lgObjRs("pay_grd1"))
            .frm1.txtpay_grd2.value = "<%=ConvSPChars(lgObjRs("pay_grd2"))%>"         

            .frm1.txtbirt.text = "<%=UniConvDateDbToCompany(lgObjRs("birt"),"")%>"

            .frm1.txtmemo_dt.text = "<%=UniConvDateDbToCompany(lgObjRs("memo_dt"),"")%>"

            .frm1.txtmemo_cd.value = "<%=ConvSPChars(lgObjRs("memo_cd"))%>"
            .frm1.txtmemo_cd_nm.value = "<%=ConvSPChars(lgObjRs("memo_cd_nm"))%>" 'FuncCodeName(1, "H0028", lgObjRs("memo_cd"))
            .frm1.txtentr_cd.value = "<%=ConvSPChars(lgObjRs("entr_cd"))%>"
            .frm1.txtentr_cd_nm.value = "<%=ConvSPChars(lgObjRs("entr_cd_nm"))%>" 'FuncCodeName(1, "H0016", lgObjRs("entr_cd"))
            .frm1.txtapp_cd.value = "<%=ConvSPChars(lgObjRs("app_cd"))%>"
            .frm1.txtapp_cd_nm.value = "<%=ConvSPChars(lgObjRs("app_cd_nm"))%>" 'FuncCodeName(1, "H0017", lgObjRs("app_cd"))

            .frm1.txtgroup_entr_dt.text = "<%=UniConvDateDbToCompany(lgObjRs("group_entr_dt"),"")%>"
            .frm1.txtentr_dt.text = "<%=UniConvDateDbToCompany(lgObjRs("entr_dt"),"")%>"

            .frm1.txtretire_dt.text = "<%=UniConvDateDbToCompany(lgObjRs("retire_dt"),"")%>"
            .frm1.txtintern_dt.text = "<%=UniConvDateDbToCompany(lgObjRs("intern_dt"),"")%>"

            .frm1.txtcareer_mm.text = "<%=UNINumClientFormat(lgObjRs("career_mm"), 0, 0)%>"        
            .frm1.txtretire_resn.value = "<%=ConvSPChars(lgObjRs("retire_resn"))%>"      
            .frm1.txtsch_ship.value = "<%=ConvSPChars(lgObjRs("sch_ship"))%>"         
            .frm1.txtrelief_cd.value = "<%=ConvSPChars(lgObjRs("relief_cd"))%>"
            .frm1.txtrelief_grade.value = "<%=ConvSPChars(lgObjRs("relief_grade"))%>"     
            .frm1.txtparia_cd.value = "<%=ConvSPChars(lgObjRs("paria_cd"))%>"         
            .frm1.txtparia_grade.value = "<%=ConvSPChars(lgObjRs("paria_grade"))%>"      
            .frm1.txttalent.value = "<%=ConvSPChars(lgObjRs("talent"))%>"           
            .frm1.txtrelig_cd.value = "<%=ConvSPChars(lgObjRs("relig_cd"))%>"
            .frm1.txtrelig_cd_nm.value = "<%=ConvSPChars(lgObjRs("relig_cd_nm"))%>" '"<%=FuncCodeName(1, "H0018", lgObjRs("relig_cd"))%>"
            .frm1.txtmarry_cd.value = "<%=ConvSPChars(lgObjRs("marry_cd"))%>"         
            .frm1.txtmil_type.value = "<%=ConvSPChars(lgObjRs("mil_type"))%>"         
            .frm1.txtMil_type_nm.value = "<%=ConvSPChars(lgObjRs("mil_type_nm"))%>" '병역구분 FuncCodeName(1, "H0019", lgObjRs("Mil_type"))
            .frm1.txtmil_kind.value = "<%=ConvSPChars(lgObjRs("mil_kind"))%>"
            .frm1.txtmil_kind_nm.value = "<%=ConvSPChars(lgObjRs("mil_kind_nm"))%>" '"<%=FuncCodeName(1, "H0020", lgObjRs("mil_kind"))%>"

            .frm1.txtmil_start.text = "<%=UniConvDateDbToCompany(lgObjRs("mil_start"),"")%>"
            .frm1.txtmil_end.text = "<%=UniConvDateDbToCompany(lgObjRs("mil_end"),"")%>"

            .frm1.txtmil_grade.value = "<%=ConvSPChars(lgObjRs("mil_grade"))%>"        
            .frm1.txtMil_grade_nm.value = "<%=ConvSPChars(lgObjRs("mil_grade_nm"))%>" '병역등급/계급 FuncCodeName(1, "H0021", lgObjRs("Mil_grade"))
            .frm1.txtmil_branch.value = "<%=ConvSPChars(lgObjRs("mil_branch"))%>"       
            .frm1.txtMil_branch_nm.value = "<%=ConvSPChars(lgObjRs("mil_branch_nm"))%>" '병역병과 FuncCodeName(1, "H0022", lgObjRs("Mil_branch"))
            .frm1.txtnomit_name.value = "<%=ConvSPChars(lgObjRs("nomit_name"))%>"       
            .frm1.txtnomit_rel.value = "<%=ConvSPChars(lgObjRs("nomit_rel"))%>"        
            .frm1.txtnomit_comp_nm.value = "<%=ConvSPChars(lgObjRs("nomit_comp_nm"))%>"    
            .frm1.txtnomit_roll_pstn.value = "<%=ConvSPChars(lgObjRs("nomit_roll_pstn"))%>"  
            .frm1.txthgt.text = "<%=UNINumClientFormat(lgObjRs("hgt"),    ggQty.DecPoint,0)%>"              
            .frm1.txtwgt.text = "<%=UNINumClientFormat(lgObjRs("wgt"),    ggQty.DecPoint,0)%>"              
            .frm1.txteyesgt_left.text = "<%=UNINumClientFormat(lgObjRs("eyesgt_left"), 1,0)%>"              
            .frm1.txteyesgt_right.text = "<%=UNINumClientFormat(lgObjRs("eyesgt_right"), 1,0)%>" 
            .frm1.txtblood_type1.value = "<%=ConvSPChars(lgObjRs("blood_type1"))%>"      
            .frm1.txtblood_type2.value = "<%=ConvSPChars(lgObjRs("blood_type2"))%>"      

            .frm1.txtnat_cd.value = "<%=ConvSPChars(lgObjRs("nat_cd"))%>"           
            .frm1.txtnat_cd_nm.value = "<%=ConvSPChars(lgObjRs("nat_cd_nm"))%>" '국적 FuncCodeName(3, "", lgObjRs("nat_cd"))
            
            if .frm1.txtnat_cd.value ="KR" then
				.frm1.txtres_no.value = "<%=Mid(ConvSPChars(lgObjRs("res_no")),1,6) & "-" & Mid(ConvSPChars(lgObjRs("res_no")),7,7)%>" '주민번호 
			else
				.frm1.txtres_no.value = "<%=ConvSPChars(lgObjRs("res_no"))%>" '주민번호 
			end if
            
            .frm1.txtnatv_state.value = "<%=ConvSPChars(lgObjRs("natv_state"))%>"       
            .frm1.txtnatv_state_nm.value = "<%=ConvSPChars(lgObjRs("natv_state_nm"))%>" '출신도 FuncCodeName(1, "H0027", lgObjRs("natv_state"))
            .frm1.txtdomi.value = "<%=ConvSPChars(lgObjRs("domi"))%>"             
            .frm1.txthouse_cd.value = "<%=ConvSPChars(lgObjRs("house_cd"))%>"
            .frm1.txthouse_cd_nm.value = "<%=ConvSPChars(lgObjRs("house_cd_nm"))%>"  '"<%=FuncCodeName(1, "H0015", lgObjRs("house_cd"))%>"
            .frm1.txtzip_cd.value = "<%=ConvSPChars(lgObjRs("zip_cd"))%>"           
            .frm1.txtaddr.value = "<%=ConvSPChars(lgObjRs("addr"))%>"             
            .frm1.txtcurr_zip_cd.value = "<%=ConvSPChars(lgObjRs("curr_zip_cd"))%>"      
            .frm1.txtcurr_addr.value = "<%=ConvSPChars(lgObjRs("curr_addr"))%>"        
            .frm1.txttel_no.value = "<%=ConvSPChars(lgObjRs("tel_no"))%>"           
            .frm1.txtem_tel_no.value = "<%=ConvSPChars(lgObjRs("em_tel_no"))%>"        
            .frm1.txtdir_indir.value = "<%=ConvSPChars(lgObjRs("dir_indir"))%>"
            .frm1.txtdir_indir_nm.value = "<%=ConvSPChars(lgObjRs("dir_indir_nm"))%>"  '"<%=FuncCodeName(1, "H0071", lgObjRs("dir_indir"))%>"

            .frm1.txtrest_month.text = "<%=UNINumClientFormat(lgObjRs("rest_month"), 0, 0)%>"

            if "<%=ConvSPChars(lgObjRs("tech_man"))%>" = "Y" THEN
                .frm1.txtTech_man.checked = true
            else
                .frm1.txtTech_man.checked = false
            end if


            if "<%=ConvSPChars(lgObjRs("dalt_type"))%>" = "Y" THEN   '색맹 
                .frm1.txtdalt_type.checked = true
            else
                .frm1.txtdalt_type.checked = false
            end if

            .frm1.txtresent_promote_dt.text = "<%=UniConvDateDbToCompany(lgObjRs("resent_promote_dt"),"")%>"
            .frm1.txtmil_no.value = "<%=ConvSPChars(lgObjRs("mil_no"))%>"

            .frm1.txtorder_change_dt.text = "<%=UniConvDateDbToCompany(lgObjRs("order_change_dt"),"")%>"

            .frm1.txtDEPT_cd_NM.value = "<%=ConvSPChars(lgObjRs("DEPT_NM"))%>"

	    	<%	If ConvSPChars(lgObjRs("Sex_cd"))=1 Then	%>		' 남여구분 
	    		.frm1.txtSex_cd1.Click()
	    	<%	Else  %>
	    		.frm1.txtSex_cd2.Click()
	    	<%  End If	%>

	    	<%	If ConvSPChars(lgObjRs("so_lu_cd"))=1 Then	%>		'
	    		.frm1.txtso_lu_cd1.Click()
	    	<%	Else  %>
	    		.frm1.txtso_lu_cd2.Click()
	    	<%  End If	%>

            .frm1.txtYear_area_cd.value = "<%=ConvSPChars(lgObjRs("Year_area_cd"))%>"
            .frm1.txtHand_tel_no.value = "<%=ConvSPChars(lgObjRs("Hand_tel_no"))%>"
            .frm1.txtEMail_addr.value = "<%=ConvSPChars(lgObjRs("EMail_addr"))%>"
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

    lgStrSQL = "DELETE  HAA010T"
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

    lgStrSQL = "INSERT INTO HAA010T( emp_no,  name, hanja_name, eng_name, "
    lgStrSQL = lgStrSQL & " comp_cd,  sect_cd,  wk_area_cd,  dept_cd, "
    lgStrSQL = lgStrSQL & " roll_pstn,  ocpt_type,  func_cd,  role_cd, "
    lgStrSQL = lgStrSQL & " pay_grd1,  pay_grd2,  res_no,  birt,  so_lu_cd, "
    lgStrSQL = lgStrSQL & " memo_dt,  memo_cd,  sex_cd,  entr_cd,  app_cd, "
    lgStrSQL = lgStrSQL & " group_entr_dt,  entr_dt, retire_dt, intern_dt, career_mm, "
    lgStrSQL = lgStrSQL & " retire_resn,  sch_ship,  relief_cd,  relief_grade,  paria_cd, "
    lgStrSQL = lgStrSQL & " paria_grade, talent,  relig_cd,  marry_cd, "
    lgStrSQL = lgStrSQL & " mil_type,  mil_kind,  mil_start,  mil_end,  mil_grade,  mil_branch, "
    lgStrSQL = lgStrSQL & " nomit_name, nomit_rel, nomit_comp_nm, nomit_roll_pstn, "
    lgStrSQL = lgStrSQL & " hgt,  wgt,  eyesgt_left,  eyesgt_right, dalt_type, " 
    lgStrSQL = lgStrSQL & " blood_type1,  blood_type2, nat_cd, natv_state,  domi, house_cd, "
    lgStrSQL = lgStrSQL & " zip_cd, addr, curr_zip_cd,  curr_addr, tel_no,  em_tel_no, "
    lgStrSQL = lgStrSQL & " isrt_emp_no, isrt_dt,  updt_emp_no,  updt_dt, dir_indir, "
    lgStrSQL = lgStrSQL & " rest_month, tech_man,  resent_promote_dt,  mil_no,  order_change_dt, "
    lgStrSQL = lgStrSQL & " year_area_cd,  email_addr,  hand_tel_no,  dept_nm, internal_cd ) "

    lgStrSQL = lgStrSQL & " VALUES(" 
    lgStrSQL = lgStrSQL & FilterVar(UCase(Request("txtEmp_no")), "''", "S") & ","                  '사번 char(10)
    lgStrSQL = lgStrSQL & FilterVar(Request("txtname"), "''", "S") & ","                    '성명 char(10)
    lgStrSQL = lgStrSQL & FilterVar(Request("txthanja_name"), "''", "S") & ","              '한자명 char(10)            
    lgStrSQL = lgStrSQL & FilterVar(Request("txteng_name"), "''", "S") & " , "                        '영문명 char(20)
    lgStrSQL = lgStrSQL & FilterVar(UCase(Request("txtcomp_cd")), "''", "S") & ","          '회사코드 char(3)
    lgStrSQL = lgStrSQL & FilterVar(UCase(Request("txtsect_cd")), "''", "S") & ","          '근무구역 char(2)
    lgStrSQL = lgStrSQL & FilterVar(UCase(Request("txtwk_area_cd")), "''", "S") & ","       ' char(3)
    lgStrSQL = lgStrSQL & FilterVar(UCase(Request("txtdept_cd")), "''", "S") & ","          '부서코드 char(7)
    lgStrSQL = lgStrSQL & FilterVar(UCase(Request("txtroll_pstn")), "''", "S") & ","        '직위코드 char(2)
    lgStrSQL = lgStrSQL & FilterVar(UCase(Request("txtocpt_type")), "''", "S") & ","        '직종구분 char(2)
    lgStrSQL = lgStrSQL & FilterVar(UCase(Request("txtfunc_cd")), "''", "S") & ","          '직무코드 char(2)
    lgStrSQL = lgStrSQL & FilterVar(UCase(Request("txtrole_cd")), "''", "S") & ","          '직책코드 char(2)
    lgStrSQL = lgStrSQL & FilterVar(UCase(Request("txtpay_grd1")), "''", "S") & ","         '직급코드 char(2)
    lgStrSQL = lgStrSQL & FilterVar(UCase(Request("txtpay_grd2")), "''", "S") & ","         '호봉 char(2)
    lgStrSQL = lgStrSQL & FilterVar(Request("txtres_no"), "''", "S") & ","                  '주민번호1 char(6)
    lgStrSQL = lgStrSQL & FilterVar(UNIConvDateCompanyToDB(Request("txtbirt"),NULL),"NULL","S") & ","'생년월일datetime
    lgStrSQL = lgStrSQL & FilterVar(UCase(Request("txtso_lu_cd")), "''", "S") & ","         ' char(1)
    lgStrSQL = lgStrSQL & FilterVar(UNIConvDateCompanyToDB(Request("txtmemo_dt"),NULL),"NULL","S") & ","  ' 기념일datetime
    lgStrSQL = lgStrSQL & FilterVar(UCase(Request("txtmemo_cd")), "''", "S") & ","          '기념일구분 char(1)
    lgStrSQL = lgStrSQL & FilterVar(UCase(Request("txtsex_cd")), "''", "S") & ","           ' char(1)
    lgStrSQL = lgStrSQL & FilterVar(UCase(Request("txtentr_cd")), "''", "S") & ","          '입사구분 char(1)
    lgStrSQL = lgStrSQL & FilterVar(UCase(Request("txtapp_cd")), "''", "S") & ","           ' char(1)
    lgStrSQL = lgStrSQL & FilterVar(UNIConvDateCompanyToDB(Request("txtGroup_entr_dt"),NULL),"NULL","S") & "," '그룹입사일    
   
    lgStrSQL = lgStrSQL & FilterVar(UNIConvDateCompanyToDB(Request("txtentr_dt"),NULL),"NULL","S") & ","  '입사일 datetime
    lgStrSQL = lgStrSQL & FilterVar(UNIConvDateCompanyToDB(Request("txtRetire_dt"),NULL),"NULL","S") & ","  '퇴사일 
    lgStrSQL = lgStrSQL & FilterVar(UNIConvDateCompanyToDB(Request("txtIntern_dt"),NULL),"NULL","S") & ","  '수습만료일 
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtcareer_mm"),0) & ","            '인정경력 decimal(16, 0)
    lgStrSQL = lgStrSQL & FilterVar(UCase(Request("txtretire_resn")), "''", "S") & ","    '퇴직사유 char(2)
    lgStrSQL = lgStrSQL & FilterVar(UCase(Request("txtsch_ship")), "''", "S") & ","       ' char(2)
    lgStrSQL = lgStrSQL & FilterVar(UCase(Request("txtrelief_cd")), "''", "S") & ","      ' char(1)
    lgStrSQL = lgStrSQL & FilterVar(UCase(Request("txtrelief_grade")), "''", "S") & ","   ' char(2)
    lgStrSQL = lgStrSQL & FilterVar(UCase(Request("txtparia_cd")), "''", "S") & ","       ' char(1)
    lgStrSQL = lgStrSQL & FilterVar(UCase(Request("txtparia_grade")), "''", "S") & ","    ' char(2)
    lgStrSQL = lgStrSQL & FilterVar(UCase(Request("txttalent")), "''", "S") & ","         ' char(20)
    lgStrSQL = lgStrSQL & FilterVar(UCase(Request("txtrelig_cd")), "''", "S") & ","       ' char(2)
    lgStrSQL = lgStrSQL & FilterVar(UCase(Request("txtmarry_cd")), "''", "S") & ","       ' char(1)
    lgStrSQL = lgStrSQL & FilterVar(UCase(Request("txtmil_type")), "''", "S") & ","       ' char(2)
    lgStrSQL = lgStrSQL & FilterVar(UCase(Request("txtmil_kind")), "''", "S") & ","       ' char(2)
    lgStrSQL = lgStrSQL & FilterVar(UNIConvDateCompanyToDB(Request("txtmil_start"),NULL),"NULL","S") & ","' datetime
    lgStrSQL = lgStrSQL & FilterVar(UNIConvDateCompanyToDB(Request("txtmil_end"),NULL),"NULL","S") & ","  ' datetime
    lgStrSQL = lgStrSQL & FilterVar(UCase(Request("txtmil_grade")), "''", "S") & ","      ' char(2)
    lgStrSQL = lgStrSQL & FilterVar(UCase(Request("txtmil_branch")), "''", "S") & ","     ' char(2)
    lgStrSQL = lgStrSQL & FilterVar(UCase(Request("txtnomit_name")), "''", "S") & ","     ' char(10)
    lgStrSQL = lgStrSQL & FilterVar(UCase(Request("txtnomit_rel")), "''", "S") & ","      ' char(10)
    lgStrSQL = lgStrSQL & FilterVar(UCase(Request("txtnomit_comp_nm")), "''", "S") & ","  ' char(30)
    lgStrSQL = lgStrSQL & FilterVar(UCase(Request("txtnomit_roll_pstn")), "''", "S") & ","' char(10)
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txthgt"),0) & ","                  ' decimal(16, 0)
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtwgt"),0) & ","                  ' decimal(16, 0)
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txteyesgt_left"),0) & ","          ' decimal(16, 0)
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txteyesgt_right"),0) & ","         ' decimal(16, 0)

    If IsEmpty(Request("txtDalt_type")) = true Then                              ' char(1) 색맹??????
        lgStrSQL = lgStrSQL & " " & FilterVar("N", "''", "S") & " ,"
    ELSE
        lgStrSQL = lgStrSQL & " " & FilterVar("Y", "''", "S") & " ,"
    END IF

    lgStrSQL = lgStrSQL & FilterVar(UCase(Request("txtblood_type1")), "''", "S") & ","    ' char(2)
    lgStrSQL = lgStrSQL & FilterVar(UCase(Request("txtblood_type2")), "''", "S") & ","    ' char(1)
    lgStrSQL = lgStrSQL & FilterVar(UCase(Request("txtnat_cd")), "''", "S") & ","         ' char(2)
    lgStrSQL = lgStrSQL & FilterVar(UCase(Request("txtnatv_state")), "''", "S") & ","     ' char(2) 출신도 
    lgStrSQL = lgStrSQL & FilterVar(Request("txtdomi"), "''", "S") & ","           ' char(40) 본적 
    lgStrSQL = lgStrSQL & FilterVar(UCase(Request("txthouse_cd")), "''", "S") & ","       ' char(1)
    lgStrSQL = lgStrSQL & FilterVar(Request("txtzip_cd"), "''", "S") & ","         ' char(6)
    lgStrSQL = lgStrSQL & FilterVar(Request("txtaddr"), "''", "S") & ","           ' char(70)
    lgStrSQL = lgStrSQL & FilterVar(Request("txtcurr_zip_cd"), "''", "S") & ","    ' char(6)
    lgStrSQL = lgStrSQL & FilterVar(Request("txtcurr_addr"), "''", "S") & ","      ' char(70)
    lgStrSQL = lgStrSQL & FilterVar(Request("txttel_no"), "''", "S") & ","         ' char(16)
    lgStrSQL = lgStrSQL & FilterVar(Request("txtem_tel_no"), "''", "S") & ","      ' char(16)
    lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''", "S") & ","                       ' char(10)
    lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime, "''", "S") & ","  ' datetime
    lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''", "S") & ","                       ' char(10)
    lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime, "''", "S") & ","  ' datetime
    lgStrSQL = lgStrSQL & FilterVar(Request("txtdir_indir"), "''", "S") & ","      '직간구분 char(1)
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtrest_month"),0) & ","           ' decimal(16, 0)

    If IsEmpty(Request("txttech_man")) = true Then                               ' char(1)
        lgStrSQL = lgStrSQL & " " & FilterVar("N", "''", "S") & " ,"
    ELSE
        lgStrSQL = lgStrSQL & " " & FilterVar("Y", "''", "S") & " ,"
    END IF

    lgStrSQL = lgStrSQL & FilterVar(UNIConvDateCompanyToDB(Request("txtResent_promote_dt"),NULL),"NULL","S") & "," '최근승급일 
    lgStrSQL = lgStrSQL & FilterVar(Request("txtmil_no"), "''", "S") & ","         ' char(10)
    lgStrSQL = lgStrSQL & FilterVar(UNIConvDateCompanyToDB(Request("txtorder_change_dt"),NULL),"NULL","S") & ","' datetime

    lgStrSQL = lgStrSQL & FilterVar(Request("txtYear_area_cd"), "''", "S") & ","   ' 신고사업장 
    lgStrSQL = lgStrSQL & FilterVar(Request("txtEmail_addr"), "''", "S") & " , "     ' E-Mail
    lgStrSQL = lgStrSQL & FilterVar(Request("txtHand_tel_no"), "''", "S") & ","    ' 핸드폰 
    
    lgStrSQL = lgStrSQL & FilterVar(Request("txtDEPT_cd_nm"), "''", "S") & ","     ' 부서명char(30)
    ' 내부부서코드 
    lgStrSQL = lgStrSQL & FilterVar(FuncCodeName(5, Request("txtdept_cd"), ""), "''", "S")

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
    lgStrSQL = "UPDATE  HAA010T"
    lgStrSQL = lgStrSQL & " SET " 
    lgStrSQL = lgStrSQL & " name = " & FilterVar(Request("txtname"), "''", "S") & ","                               '성명 char(10)
    
    '영문명 char(20)
    lgStrSQL = lgStrSQL & " hanja_name = " & FilterVar(Request("txthanja_name"), "''", "S") & ","                               '성명 char(10)    
    lgStrSQL = lgStrSQL & " eng_name = " & FilterVar(Request("txteng_name"), "''", "S") & " , "

    '회사코드 char(3)
    lgStrSQL = lgStrSQL & " comp_cd = " & FilterVar(UCase(Request("txtcomp_cd")), "''", "S") & ","

    '근무구역 char(2)
    lgStrSQL = lgStrSQL & " sect_cd = " & FilterVar(UCase(Request("txtsect_cd")), "''", "S") & ","
    lgStrSQL = lgStrSQL & " wk_area_cd = " & FilterVar(UCase(Request("txtwk_area_cd")), "''", "S") & ","
    lgStrSQL = lgStrSQL & " dept_cd = " & FilterVar(Request("txtdept_cd"), "''", "S") & ","              '부서코드 char(7)
    lgStrSQL = lgStrSQL & " roll_pstn = " & FilterVar(UCase(Request("txtroll_pstn")), "''", "S") & ","          '직위코드 char(2)
    lgStrSQL = lgStrSQL & " ocpt_type = " & FilterVar(UCase(Request("txtocpt_type")), "''", "S") & ","          '직종구분 char(2)
    lgStrSQL = lgStrSQL & " func_cd = " & FilterVar(UCase(Request("txtfunc_cd")), "''", "S") & ","              '직무코드 char(2)
    lgStrSQL = lgStrSQL & " role_cd = " & FilterVar(UCase(Request("txtrole_cd")), "''", "S") & ","              '직책코드 char(2)
    lgStrSQL = lgStrSQL & " pay_grd1 = " & FilterVar(UCase(Request("txtpay_grd1")), "''", "S") & ","            '직급코드 char(2)
    lgStrSQL = lgStrSQL & " pay_grd2 = " & FilterVar(UCase(Request("txtpay_grd2")), "''", "S") & ","            '호봉 char(2)
    lgStrSQL = lgStrSQL & " res_no = " & FilterVar(Request("txtres_no"), "''", "S") & ","                '주민번호1 char(6)
    lgStrSQL = lgStrSQL & " birt = " & FilterVar(UNIConvDateCompanyToDB(Request("txtbirt"),NULL),"NULL","S") & ","       '생년월일datetime
    lgStrSQL = lgStrSQL & " so_lu_cd = " & FilterVar(UCase(Request("txtso_lu_cd")), "''", "S") & ","            ' char(1)
    lgStrSQL = lgStrSQL & " memo_dt = " & FilterVar(UNIConvDateCompanyToDB(Request("txtmemo_dt"),NULL),"NULL","S") & "," ' datetime

    '기념일구분char(3)
    lgStrSQL = lgStrSQL & " memo_cd = " & FilterVar(UCase(Request("txtmemo_cd")), "''", "S") & ","                       

    '남여구분 char(1)
    lgStrSQL = lgStrSQL & " sex_cd = " & FilterVar(UCase(Request("txtsex_cd")), "''", "S") & ","

    '입사구분 char(1)
    lgStrSQL = lgStrSQL & " entr_cd = " & FilterVar(UCase(Request("txtentr_cd")), "''", "S") & ","
    lgStrSQL = lgStrSQL & " app_cd = " & FilterVar(UCase(Request("txtapp_cd")), "''", "S") & ","                       

    '그룹입사일    
    lgStrSQL = lgStrSQL & " group_entr_dt = " & FilterVar(UNIConvDateCompanyToDB(Request("txtGroup_entr_dt"),NULL),"NULL","S") & ","

    '입사일 datetime
    lgStrSQL = lgStrSQL & " entr_dt = " & FilterVar(UNIConvDateCompanyToDB(Request("txtentr_dt"),NULL),"NULL","S") & ","

    '퇴사일 
    lgStrSQL = lgStrSQL & " retire_dt = " & FilterVar(UNIConvDateCompanyToDB(Request("txtRetire_dt"),NULL),"NULL","S") & ","

    '수습만료일 
    lgStrSQL = lgStrSQL & " intern_dt = " & FilterVar(UNIConvDateCompanyToDB(Request("txtIntern_dt"),NULL),"NULL","S") & ","

    '인정경력 decimal(16, 0)
    lgStrSQL = lgStrSQL & " career_mm = " & UNIConvNum(Request("txtcareer_mm"),0) & ","

    '퇴직사유 char(2)
    lgStrSQL = lgStrSQL & " retire_resn = " & FilterVar(UCase(Request("txtretire_resn")), "''", "S") & ","                       

    '최종학력 char(2)
    lgStrSQL = lgStrSQL & " sch_ship = " & FilterVar(UCase(Request("txtsch_ship")), "''", "S") & ","

    '보훈구분 char(1)
    lgStrSQL = lgStrSQL & " relief_cd = " & FilterVar(UCase(Request("txtrelief_cd")), "''", "S") & ","

    '보훈등급 char(2)
    lgStrSQL = lgStrSQL & " relief_grade = " & FilterVar(UCase(Request("txtrelief_grade")), "''", "S") & ","

    '장애구분 char(1)
    lgStrSQL = lgStrSQL & " paria_cd = " & FilterVar(UCase(Request("txtparia_cd")), "''", "S") & ","

    '장애등급 char(2)
    lgStrSQL = lgStrSQL & " paria_grade = " & FilterVar(UCase(Request("txtparia_grade")), "''", "S") & ","

    '특기 char(20)
    lgStrSQL = lgStrSQL & " talent = " & FilterVar(UCase(Request("txttalent")), "''", "S") & ","

    '종교 char(2)
    lgStrSQL = lgStrSQL & " relig_cd = " & FilterVar(UCase(Request("txtrelig_cd")), "''", "S") & ","

    '결혼여부 char(1)
    lgStrSQL = lgStrSQL & " marry_cd = " & FilterVar(UCase(Request("txtmarry_cd")), "''", "S") & ","

    '병역구분 char(2)
    lgStrSQL = lgStrSQL & " mil_type = " & FilterVar(UCase(Request("txtmil_type")), "''", "S") & ","

    '병역군별 char(2)
    lgStrSQL = lgStrSQL & " mil_kind = " & FilterVar(UCase(Request("txtmil_kind")), "''", "S") & ","
    lgStrSQL = lgStrSQL & " mil_start = " & FilterVar(UNIConvDateCompanyToDB(Request("txtmil_start"),NULL),"NULL","S") & ","    ' datetime
    lgStrSQL = lgStrSQL & " mil_end = " & FilterVar(UNIConvDateCompanyToDB(Request("txtmil_end"),NULL),"NULL","S") & ","        ' datetime

    '병역등급 char(2)
    lgStrSQL = lgStrSQL & " mil_grade = " & FilterVar(UCase(Request("txtmil_grade")), "''", "S") & ","

    '병역병과 char(2)
    lgStrSQL = lgStrSQL & " mil_branch = " & FilterVar(UCase(Request("txtmil_branch")), "''", "S") & ","

    '추천인성명 char(10)
    lgStrSQL = lgStrSQL & " nomit_name = " & FilterVar(Request("txtnomit_name"), "''", "S") & ","

    '추천인관계 char(10)
    lgStrSQL = lgStrSQL & " nomit_rel = " & FilterVar(UCase(Request("txtnomit_rel")), "''", "S") & ","

    '추천인근무지 char(30)
    lgStrSQL = lgStrSQL & " nomit_comp_nm = " & FilterVar(Request("txtnomit_comp_nm"), "''", "S") & ","

    '추천인직위 char(10)
    lgStrSQL = lgStrSQL & " nomit_roll_pstn = " & FilterVar(Request("txtnomit_roll_pstn"), "''", "S") & ","

    '신장 
    lgStrSQL = lgStrSQL & " hgt = " & UNIConvNum(Request("txthgt"),0) & ","

    '몸무계 
    lgStrSQL = lgStrSQL & " wgt = " & UNIConvNum(Request("txtwgt"),0) & ","

    '시력(좌)
    lgStrSQL = lgStrSQL & " eyesgt_left = " & UNIConvNum(Request("txteyesgt_left"),0) & ","

    '시력(우)
    lgStrSQL = lgStrSQL & " eyesgt_right = " & UNIConvNum(Request("txteyesgt_right"),0) & ","

    ' char(1) 색맹??????
    If IsEmpty(Request("txtDalt_type")) = true Then
        lgStrSQL = lgStrSQL & " dalt_type = " & FilterVar("N", "''", "S") & " ,"
    ELSE
        lgStrSQL = lgStrSQL & " dalt_type = " & FilterVar("Y", "''", "S") & " ,"
    END IF

    '혈액형 char(2)
    lgStrSQL = lgStrSQL & " blood_type1 = " & FilterVar(UCase(Request("txtblood_type1")), "''", "S") & ","

    '혈액형 char(1)
    lgStrSQL = lgStrSQL & " blood_type2 = " & FilterVar(UCase(Request("txtblood_type2")), "''", "S") & ","

    '국적 char(2)
    lgStrSQL = lgStrSQL & " nat_cd = " & FilterVar(UCase(Request("txtnat_cd")), "''", "S") & ","

    '출신도 
    lgStrSQL = lgStrSQL & " natv_state = " & FilterVar(UCase(Request("txtnatv_state")), "''", "S") & ","

    '본적 
    lgStrSQL = lgStrSQL & " domi = " & FilterVar(Request("txtdomi"), "''", "S") & ","

    '주거구분 char(1)
    lgStrSQL = lgStrSQL & " house_cd = " & FilterVar(UCase(Request("txthouse_cd")), "''", "S") & ","
    lgStrSQL = lgStrSQL & " zip_cd = " & FilterVar(Request("txtzip_cd"), "''", "S") & ","                           ' char(6)
    lgStrSQL = lgStrSQL & " addr = " & FilterVar(Request("txtaddr"), "''", "S") & ","                               ' char(70)

    '현주소 char(6)
    lgStrSQL = lgStrSQL & " curr_zip_cd = " & FilterVar(Request("txtcurr_zip_cd"), "''", "S") & ","

    '현주소 char(70)
    lgStrSQL = lgStrSQL & " curr_addr = " & FilterVar(Request("txtcurr_addr"), "''", "S") & ","

    '전화번호 char(16)
    lgStrSQL = lgStrSQL & " tel_no = " & FilterVar(Request("txttel_no"), "''", "S") & ","

    '비상전화번호 char(16)
    lgStrSQL = lgStrSQL & " em_tel_no = " & FilterVar(Request("txtem_tel_no"), "''", "S") & ","
    lgStrSQL = lgStrSQL & " isrt_emp_no = " & FilterVar(gUsrId, "''", "S") & ","                                   ' char(10)
    lgStrSQL = lgStrSQL & " isrt_dt = " & FilterVar(GetSvrDateTime, "''", "S") & ","                ' datetime
    lgStrSQL = lgStrSQL & " updt_emp_no = " & FilterVar(gUsrId, "''", "S") & ","                                   ' char(10)
    lgStrSQL = lgStrSQL & " updt_dt = " & FilterVar(GetSvrDateTime, "''", "S") & ","                ' datetime

    '직간구분 char(1)
    lgStrSQL = lgStrSQL & " dir_indir = " & FilterVar(Request("txtdir_indir"), "''", "S") & ","

    '휴직개월 
    lgStrSQL = lgStrSQL & " rest_month = " & UNIConvNum(Request("txtrest_month"),0) & ","

    If IsEmpty(Request("txttech_man")) = true Then     '현장기술인력제공여부 char(1)
        lgStrSQL = lgStrSQL & " tech_man = " & FilterVar("N", "''", "S") & " ,"
    ELSE
        lgStrSQL = lgStrSQL & " tech_man = " & FilterVar("Y", "''", "S") & " ,"
    END IF

    '최근승급일    
    lgStrSQL = lgStrSQL & " resent_promote_dt = " & FilterVar(UNIConvDateCompanyToDB(Request("txtResent_promote_dt"),NULL),"NULL","S") & ","

    '군번 char(10)
    lgStrSQL = lgStrSQL & " mil_no = " & FilterVar(UCase(Request("txtmil_no")), "''", "S") & ","

    '인사변동일 datetime
    lgStrSQL = lgStrSQL & " order_change_dt = " & FilterVar(UNIConvDateCompanyToDB(Request("txtorder_change_dt"),NULL),"NULL","S") & ","

    '신고사업장 
    lgStrSQL = lgStrSQL & " Year_area_cd = " & FilterVar(UCase(Request("txtYear_area_cd")), "''", "S") & ","

    '핸드폰번호 
    lgStrSQL = lgStrSQL & " Hand_tel_no = " & FilterVar(Request("txtHand_tel_no"), "''", "S") & ","

    'E-Mail
    'lgStrSQL = lgStrSQL & " EMail_addr = '" & Request("txtEMail_addr") & "',"
	lgStrSQL = lgStrSQL & " EMail_addr = " & FilterVar(Request("txtEMail_addr"), "''", "S") & " , "
    ' 부서명char(30)
    lgStrSQL = lgStrSQL & " DEPT_NM = " & FilterVar(Request("txtDEPT_cd_nm"), "''", "S") & ","
    
    If Request("txtRetire_dt") ="" Then 
        lgStrSQL = lgStrSQL & " internal_cd = " & FilterVar(FuncCodeName(5, Request("txtdept_cd"), ""), "''", "S")
    Else
		lgStrSQL = lgStrSQL & " internal_cd = " & FilterVar(FuncCodeName(5, Request("txtdept_cd"), UNIConvDateCompanyToDB(Request("txtRetire_dt"),NULL)), "''", "S")
    End If

    lgStrSQL = lgStrSQL & " WHERE emp_no = " & FilterVar(Request("txtEmp_no"), "''", "S")

    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
	Call SubHandleError("SU",lgObjConn,lgObjRs,Err)
	
	If lgtempflgchk = "true" Then	    

        lgStrSQL = "if (select count(*) from HBA010T where emp_no = " & FilterVar(Request("txtEmp_no"), "''", "S") & ") >  0  "
        lgStrSQL = lgStrSQL & " UPDATE  HBA010T"
        lgStrSQL = lgStrSQL & " SET " 
        lgStrSQL = lgStrSQL & " dept_cd      = " & FilterVar(Request("txtdept_cd"), "''", "S") & ","                     '부서코드 char(7)
        lgStrSQL = lgStrSQL & " roll_pstn    = " & FilterVar(UCase(Request("txtroll_pstn")), "''", "S") & ","          '직위코드 char(2)
        lgStrSQL = lgStrSQL & " func_cd      = " & FilterVar(UCase(Request("txtfunc_cd")), "''", "S") & ","              '직무코드 char(2)
        lgStrSQL = lgStrSQL & " role_cd      = " & FilterVar(UCase(Request("txtrole_cd")), "''", "S") & ","              '직책코드 char(2)
        lgStrSQL = lgStrSQL & " pay_grd1     = " & FilterVar(UCase(Request("txtpay_grd1")), "''", "S") & ","            '직급코드 char(2)
        lgStrSQL = lgStrSQL & " pay_grd2     = " & FilterVar(UCase(Request("txtpay_grd2")), "''", "S") & ","            '호봉 char(2)
        lgStrSQL = lgStrSQL & " sect_cd      = " & FilterVar(UCase(Request("txtsect_cd")), "''", "S")
        lgStrSQL = lgStrSQL & " WHERE emp_no = " & FilterVar(Request("txtEmp_no"), "''", "S") 
        lgStrSQL = lgStrSQL & "   AND gazet_dt = (SELECT MAX(gazet_dt) "
        lgStrSQL = lgStrSQL & "                     FROM HBA010T "
        lgStrSQL = lgStrSQL & "                    WHERE emp_no = " & FilterVar(Request("txtEmp_no"), "''", "S") & ") "

        lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
		
		Call SubHandleError("SU",lgObjConn,lgObjRs,Err)
	    
    End if

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
                           lgStrSQL = "Select emp_no, name, hanja_name, eng_name, dept_cd, " 
                           lgStrSQL = lgStrSQL & "comp_cd, dbo.ufn_H_GetCodeName(" & FilterVar("B_COMPANY", "''", "S") & ", comp_cd, '') comp_cd_nm, "
                           lgStrSQL = lgStrSQL & "sect_cd, dbo.ufn_GetCodeName(" & FilterVar("H0035", "''", "S") & ", sect_cd) sect_cd_nm, "
                           lgStrSQL = lgStrSQL & "wk_area_cd, dbo.ufn_GetCodeName(" & FilterVar("H0036", "''", "S") & ", wk_area_cd) wk_area_cd_nm, "
                           lgStrSQL = lgStrSQL & "ocpt_type, dbo.ufn_GetCodeName(" & FilterVar("H0003", "''", "S") & ", ocpt_type) ocpt_type_nm, "
                           lgStrSQL = lgStrSQL & "roll_pstn, dbo.ufn_GetCodeName(" & FilterVar("H0002", "''", "S") & ", roll_pstn) roll_pstn_nm, "
                           lgStrSQL = lgStrSQL & "pay_grd2, birt, memo_dt, so_lu_cd, "
                           lgStrSQL = lgStrSQL & "func_cd, dbo.ufn_GetCodeName(" & FilterVar("H0004", "''", "S") & ", func_cd) func_cd_nm, "
                           lgStrSQL = lgStrSQL & "role_cd, dbo.ufn_GetCodeName(" & FilterVar("H0026", "''", "S") & ", role_cd) role_cd_nm, "
                           lgStrSQL = lgStrSQL & "pay_grd1, dbo.ufn_GetCodeName(" & FilterVar("H0001", "''", "S") & ", pay_grd1) pay_grd1_nm, "
                           lgStrSQL = lgStrSQL & "memo_cd, dbo.ufn_GetCodeName(" & FilterVar("H0028", "''", "S") & ", memo_cd) memo_cd_nm, "
                           lgStrSQL = lgStrSQL & "sex_cd, entr_cd, dbo.ufn_GetCodeName(" & FilterVar("H0016", "''", "S") & ", entr_cd) entr_cd_nm, "
                           lgStrSQL = lgStrSQL & "app_cd, dbo.ufn_GetCodeName(" & FilterVar("H0017", "''", "S") & ", app_cd) app_cd_nm, "
                           lgStrSQL = lgStrSQL & "group_entr_dt, entr_dt, retire_dt, intern_dt, career_mm, "
                           lgStrSQL = lgStrSQL & "retire_resn, sch_ship, relief_grade, "
                           lgStrSQL = lgStrSQL & "relief_cd, " ' dbo.ufn_GetCodeName('H0014', relief_cd) relief_cd_nm, "
                           lgStrSQL = lgStrSQL & "paria_cd, "  ' dbo.ufn_GetCodeName('H0013', paria_cd) paria_cd_nm, "
                           lgStrSQL = lgStrSQL & "paria_grade, talent, marry_cd, mil_start, mil_end, "
                           lgStrSQL = lgStrSQL & "relig_cd, dbo.ufn_GetCodeName(" & FilterVar("H0018", "''", "S") & ", relig_cd) relig_cd_nm, "
                           lgStrSQL = lgStrSQL & "mil_type, dbo.ufn_GetCodeName(" & FilterVar("H0019", "''", "S") & ", mil_type) mil_type_nm, "
                           lgStrSQL = lgStrSQL & "mil_kind, dbo.ufn_GetCodeName(" & FilterVar("H0020", "''", "S") & ", mil_kind) mil_kind_nm, "
                           lgStrSQL = lgStrSQL & "mil_grade, dbo.ufn_GetCodeName(" & FilterVar("H0021", "''", "S") & ", mil_grade) mil_grade_nm, "
                           lgStrSQL = lgStrSQL & "mil_branch, dbo.ufn_GetCodeName(" & FilterVar("H0022", "''", "S") & ", mil_branch) mil_branch_nm, "
                           lgStrSQL = lgStrSQL & "nomit_name, nomit_rel, nomit_comp_nm, "
                           lgStrSQL = lgStrSQL & "nomit_roll_pstn, hgt, wgt, eyesgt_left, eyesgt_right, "
                           lgStrSQL = lgStrSQL & "blood_type1, blood_type2, res_no, domi, "
                           lgStrSQL = lgStrSQL & "natv_state, dbo.ufn_GetCodeName(" & FilterVar("H0027", "''", "S") & ", natv_state) natv_state_nm, "
                           lgStrSQL = lgStrSQL & "house_cd, dbo.ufn_GetCodeName(" & FilterVar("H0015", "''", "S") & ", house_cd) house_cd_nm, "
                           lgStrSQL = lgStrSQL & "nat_cd, dbo.ufn_H_GetCodeName(" & FilterVar("B_COUNTRY", "''", "S") & ", nat_cd, '') nat_cd_nm, "
                           lgStrSQL = lgStrSQL & "zip_cd, addr, curr_zip_cd, curr_addr, tel_no, em_tel_no, "
                           lgStrSQL = lgStrSQL & "dir_indir, dbo.ufn_GetCodeName(" & FilterVar("H0071", "''", "S") & ", dir_indir) dir_indir_nm, "
                           lgStrSQL = lgStrSQL & "rest_month, tech_man, dalt_type, resent_promote_dt, mil_no, "  
                           lgStrSQL = lgStrSQL & "order_change_dt, dept_nm, sex_cd, so_lu_cd, "
                           lgStrSQL = lgStrSQL & "year_area_cd, " 'dbo.ufn_GetCodeName('H0068', year_area_cd) year_area_cd_nm, "
                           lgStrSQL = lgStrSQL & "hand_tel_no, eMail_addr "
                           lgStrSQL = lgStrSQL & " From  HAA010T "
                           lgStrSQL = lgStrSQL & " WHERE emp_no = " & pCode 	

                     Case "P"
                           lgStrSQL = "Select TOP 1 emp_no, name, hanja_name, eng_name, dept_cd, " 
                           lgStrSQL = lgStrSQL & "comp_cd, dbo.ufn_H_GetCodeName(" & FilterVar("B_COMPANY", "''", "S") & ", comp_cd, '') comp_cd_nm, "
                           lgStrSQL = lgStrSQL & "sect_cd, dbo.ufn_GetCodeName(" & FilterVar("H0035", "''", "S") & ", sect_cd) sect_cd_nm, "
                           lgStrSQL = lgStrSQL & "wk_area_cd, dbo.ufn_GetCodeName(" & FilterVar("H0036", "''", "S") & ", wk_area_cd) wk_area_cd_nm, "
                           lgStrSQL = lgStrSQL & "ocpt_type, dbo.ufn_GetCodeName(" & FilterVar("H0003", "''", "S") & ", ocpt_type) ocpt_type_nm, "
                           lgStrSQL = lgStrSQL & "roll_pstn, dbo.ufn_GetCodeName(" & FilterVar("H0002", "''", "S") & ", roll_pstn) roll_pstn_nm, "
                           lgStrSQL = lgStrSQL & "pay_grd2, birt, memo_dt, so_lu_cd, "
                           lgStrSQL = lgStrSQL & "func_cd, dbo.ufn_GetCodeName(" & FilterVar("H0004", "''", "S") & ", func_cd) func_cd_nm, "
                           lgStrSQL = lgStrSQL & "role_cd, dbo.ufn_GetCodeName(" & FilterVar("H0026", "''", "S") & ", role_cd) role_cd_nm, "
                           lgStrSQL = lgStrSQL & "pay_grd1, dbo.ufn_GetCodeName(" & FilterVar("H0001", "''", "S") & ", pay_grd1) pay_grd1_nm, "
                           lgStrSQL = lgStrSQL & "memo_cd, dbo.ufn_GetCodeName(" & FilterVar("H0028", "''", "S") & ", memo_cd) memo_cd_nm, "
                           lgStrSQL = lgStrSQL & "sex_cd, entr_cd, dbo.ufn_GetCodeName(" & FilterVar("H0016", "''", "S") & ", entr_cd) entr_cd_nm, "
                           lgStrSQL = lgStrSQL & "app_cd, dbo.ufn_GetCodeName(" & FilterVar("H0017", "''", "S") & ", app_cd) app_cd_nm, "
                           lgStrSQL = lgStrSQL & "group_entr_dt, entr_dt, retire_dt, intern_dt, career_mm, "
                           lgStrSQL = lgStrSQL & "retire_resn, sch_ship, relief_grade, "
                           lgStrSQL = lgStrSQL & "relief_cd, " ' dbo.ufn_GetCodeName('H0014', relief_cd) relief_cd_nm, "
                           lgStrSQL = lgStrSQL & "paria_cd, "  ' dbo.ufn_GetCodeName('H0013', paria_cd) paria_cd_nm, "
                           lgStrSQL = lgStrSQL & "paria_grade, talent, marry_cd, mil_start, mil_end, "
                           lgStrSQL = lgStrSQL & "relig_cd, dbo.ufn_GetCodeName(" & FilterVar("H0018", "''", "S") & ", relig_cd) relig_cd_nm, "
                           lgStrSQL = lgStrSQL & "mil_type, dbo.ufn_GetCodeName(" & FilterVar("H0019", "''", "S") & ", mil_type) mil_type_nm, "
                           lgStrSQL = lgStrSQL & "mil_kind, dbo.ufn_GetCodeName(" & FilterVar("H0020", "''", "S") & ", mil_kind) mil_kind_nm, "
                           lgStrSQL = lgStrSQL & "mil_grade, dbo.ufn_GetCodeName(" & FilterVar("H0021", "''", "S") & ", mil_grade) mil_grade_nm, "
                           lgStrSQL = lgStrSQL & "mil_branch, dbo.ufn_GetCodeName(" & FilterVar("H0022", "''", "S") & ", mil_branch) mil_branch_nm, "
                           lgStrSQL = lgStrSQL & "nomit_name, nomit_rel, nomit_comp_nm, "
                           lgStrSQL = lgStrSQL & "nomit_roll_pstn, hgt, wgt, eyesgt_left, eyesgt_right, "
                           lgStrSQL = lgStrSQL & "blood_type1, blood_type2, res_no, domi, "
                           lgStrSQL = lgStrSQL & "natv_state, dbo.ufn_GetCodeName(" & FilterVar("H0027", "''", "S") & ", natv_state) natv_state_nm, "
                           lgStrSQL = lgStrSQL & "house_cd, dbo.ufn_GetCodeName(" & FilterVar("H0015", "''", "S") & ", house_cd) house_cd_nm, "
                           lgStrSQL = lgStrSQL & "nat_cd, dbo.ufn_H_GetCodeName(" & FilterVar("B_COUNTRY", "''", "S") & ", nat_cd, '') nat_cd_nm, "
                           lgStrSQL = lgStrSQL & "zip_cd, addr, curr_zip_cd, curr_addr, tel_no, em_tel_no, "
                           lgStrSQL = lgStrSQL & "dir_indir, dbo.ufn_GetCodeName(" & FilterVar("H0071", "''", "S") & ", dir_indir) dir_indir_nm, "
                           lgStrSQL = lgStrSQL & "rest_month, tech_man, dalt_type, resent_promote_dt, mil_no, "  
                           lgStrSQL = lgStrSQL & "order_change_dt, dept_nm, sex_cd, so_lu_cd, "
                           lgStrSQL = lgStrSQL & "year_area_cd, " 'dbo.ufn_GetCodeName('H0068', year_area_cd) year_area_cd_nm, "
                           lgStrSQL = lgStrSQL & "hand_tel_no, eMail_addr "
                           lgStrSQL = lgStrSQL & " From  HAA010T "
                           lgStrSQL = lgStrSQL & " WHERE emp_no < " & pCode 	
                           lgStrSQL = lgStrSQL & " ORDER BY emp_no DESC "
                     Case "N"
                           lgStrSQL = "Select TOP 1 emp_no, name, hanja_name, eng_name, dept_cd, " 
                           lgStrSQL = lgStrSQL & "comp_cd, dbo.ufn_H_GetCodeName(" & FilterVar("B_COMPANY", "''", "S") & ", comp_cd, '') comp_cd_nm, "
                           lgStrSQL = lgStrSQL & "sect_cd, dbo.ufn_GetCodeName(" & FilterVar("H0035", "''", "S") & ", sect_cd) sect_cd_nm, "
                           lgStrSQL = lgStrSQL & "wk_area_cd, dbo.ufn_GetCodeName(" & FilterVar("H0036", "''", "S") & ", wk_area_cd) wk_area_cd_nm, "
                           lgStrSQL = lgStrSQL & "ocpt_type, dbo.ufn_GetCodeName(" & FilterVar("H0003", "''", "S") & ", ocpt_type) ocpt_type_nm, "
                           lgStrSQL = lgStrSQL & "roll_pstn, dbo.ufn_GetCodeName(" & FilterVar("H0002", "''", "S") & ", roll_pstn) roll_pstn_nm, "
                           lgStrSQL = lgStrSQL & "pay_grd2, birt, memo_dt, so_lu_cd, "
                           lgStrSQL = lgStrSQL & "func_cd, dbo.ufn_GetCodeName(" & FilterVar("H0004", "''", "S") & ", func_cd) func_cd_nm, "
                           lgStrSQL = lgStrSQL & "role_cd, dbo.ufn_GetCodeName(" & FilterVar("H0026", "''", "S") & ", role_cd) role_cd_nm, "
                           lgStrSQL = lgStrSQL & "pay_grd1, dbo.ufn_GetCodeName(" & FilterVar("H0001", "''", "S") & ", pay_grd1) pay_grd1_nm, "
                           lgStrSQL = lgStrSQL & "memo_cd, dbo.ufn_GetCodeName(" & FilterVar("H0028", "''", "S") & ", memo_cd) memo_cd_nm, "
                           lgStrSQL = lgStrSQL & "sex_cd, entr_cd, dbo.ufn_GetCodeName(" & FilterVar("H0016", "''", "S") & ", entr_cd) entr_cd_nm, "
                           lgStrSQL = lgStrSQL & "app_cd, dbo.ufn_GetCodeName(" & FilterVar("H0017", "''", "S") & ", app_cd) app_cd_nm, "
                           lgStrSQL = lgStrSQL & "group_entr_dt, entr_dt, retire_dt, intern_dt, career_mm, "
                           lgStrSQL = lgStrSQL & "retire_resn, sch_ship, relief_grade, "
                           lgStrSQL = lgStrSQL & "relief_cd, " ' dbo.ufn_GetCodeName('H0014', relief_cd) relief_cd_nm, "
                           lgStrSQL = lgStrSQL & "paria_cd, "  ' dbo.ufn_GetCodeName('H0013', paria_cd) paria_cd_nm, "
                           lgStrSQL = lgStrSQL & "paria_grade, talent, marry_cd, mil_start, mil_end, "
                           lgStrSQL = lgStrSQL & "relig_cd, dbo.ufn_GetCodeName(" & FilterVar("H0018", "''", "S") & ", relig_cd) relig_cd_nm, "
                           lgStrSQL = lgStrSQL & "mil_type, dbo.ufn_GetCodeName(" & FilterVar("H0019", "''", "S") & ", mil_type) mil_type_nm, "
                           lgStrSQL = lgStrSQL & "mil_kind, dbo.ufn_GetCodeName(" & FilterVar("H0020", "''", "S") & ", mil_kind) mil_kind_nm, "
                           lgStrSQL = lgStrSQL & "mil_grade, dbo.ufn_GetCodeName(" & FilterVar("H0021", "''", "S") & ", mil_grade) mil_grade_nm, "
                           lgStrSQL = lgStrSQL & "mil_branch, dbo.ufn_GetCodeName(" & FilterVar("H0022", "''", "S") & ", mil_branch) mil_branch_nm, "
                           lgStrSQL = lgStrSQL & "nomit_name, nomit_rel, nomit_comp_nm, "
                           lgStrSQL = lgStrSQL & "nomit_roll_pstn, hgt, wgt, eyesgt_left, eyesgt_right, "
                           lgStrSQL = lgStrSQL & "blood_type1, blood_type2, res_no, domi, "
                           lgStrSQL = lgStrSQL & "natv_state, dbo.ufn_GetCodeName(" & FilterVar("H0027", "''", "S") & ", natv_state) natv_state_nm, "
                           lgStrSQL = lgStrSQL & "house_cd, dbo.ufn_GetCodeName(" & FilterVar("H0015", "''", "S") & ", house_cd) house_cd_nm, "
                           lgStrSQL = lgStrSQL & "nat_cd, dbo.ufn_H_GetCodeName(" & FilterVar("B_COUNTRY", "''", "S") & ", nat_cd, '') nat_cd_nm, "
                           lgStrSQL = lgStrSQL & "zip_cd, addr, curr_zip_cd, curr_addr, tel_no, em_tel_no, "
                           lgStrSQL = lgStrSQL & "dir_indir, dbo.ufn_GetCodeName(" & FilterVar("H0071", "''", "S") & ", dir_indir) dir_indir_nm, "
                           lgStrSQL = lgStrSQL & "rest_month, tech_man, dalt_type, resent_promote_dt, mil_no, "  
                           lgStrSQL = lgStrSQL & "order_change_dt, dept_nm, sex_cd, so_lu_cd, "
                           lgStrSQL = lgStrSQL & "year_area_cd, " 'dbo.ufn_GetCodeName('H0068', year_area_cd) year_area_cd_nm, "
                           lgStrSQL = lgStrSQL & "hand_tel_no, eMail_addr "
                           lgStrSQL = lgStrSQL & " From  HAA010T "
                           lgStrSQL = lgStrSQL & " WHERE emp_no > " & pCode 	
                           lgStrSQL = lgStrSQL & " ORDER BY emp_no ASC "
                           
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
