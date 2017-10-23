<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% Option Explicit%>
<% session.CodePage=949 %>

<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<!-- #Include file="../../inc/lgSvrVariables.inc" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->

<%
	Const C_SHEETMAXROWS_D = 100
    Dim lgEmpNo, lgHEmpNo, lgInternal
    Call HideStatusWnd                                                                   '☜: Hide Processing message
        
    Call LoadBasisGlobalInf()
    Call LoadInfTB19029B("I", "H", "NOCOOKIE", "MB")

    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '☜: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)
    lgEmpNo			  = FilterVar(Request("txtEmpNo"), "''", "S")
    lgHEmpNo		  = FilterVar(Request("txtHEmpNo"), "''", "S")    
    lgInternal		  = Request("txtInternal")
    	    
    lgPrevNext        = Request("txtPrevKey")                                       '☜: "P"(Prev search) "N"(Next search)

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

    Call SubMakeSQLStatements("R", lgEmpNo, lgInternal)                                   '☜ : Make sql statements

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

            .frm1.txtEmpNo.Value  = "<%=ConvSPChars(lgObjRs("emp_no"))%>"                   'Set condition area
            .Frm1.txtEmpNm.Value  = "<%=ConvSPChars(lgObjRs("name"))%>"
            .frm1.txtHEmpNo.Value  = "<%=ConvSPChars(lgHObjRs("emp_no"))%>"                   'Set condition area
            .Frm1.txtHEmpNm.Value  = "<%=ConvSPChars(lgHObjRs("name"))%>"       
                            
            .frm1.txtDeptNm.value = "<%=lgObjRs("dept_nm")%>"
            .frm1.cboRollPstn.value = "<%=ConvSPChars(lgObjRs("roll_pstn_nm"))%>"        
            .frm1.cboOcptType.value = "<%=ConvSPChars(lgObjRs("ocpt_type_nm"))%>"
            .frm1.cboFuncCd.value = "<%=ConvSPChars(lgObjRs("func_cd_nm"))%>"          
            .frm1.cboRoleCd.value = "<%=ConvSPChars(lgObjRs("role_cd_nm"))%>"          
            .frm1.cboPay_grd1.value = "<%=ConvSPChars(lgObjRs("pay_grd1_nm"))%>"         
            .frm1.txtPay_grd2.value = "<%=ConvSPChars(lgObjRs("pay_grd2"))%>"         
            .frm1.cboEntrCd.value = "<%=ConvSPChars(lgObjRs("entr_cd_nm"))%>"
			.frm1.txtCareer.value = "<%=ConvSPChars(lgObjRs("career_mm"))%>"
			
            .frm1.txtGroupEntrDt.text = "<%=UniConvDateDbToCompany(lgObjRs("group_entr_dt"),"")%>"
            .frm1.txtEntrDt.text = "<%=UniConvDateDbToCompany(lgObjRs("Entr_Dt"),"")%>"
            .frm1.txtInternDt.text = "<%=UniConvDateDbToCompany(lgObjRs("intern_dt"),"")%>"    

            If ISNULL("<%=Trim(lgObjRs("retire_dt"))%>") OR "<%=Trim(lgObjRs("retire_dt"))%>" = "" Then        
               .frm1.txtRestDt.text = "<%=UniConvDateDbToCompany(lgObjRs("rest_dt"),"")%>"
            Else
               .frm1.txtRestDt.text = "<%=UniConvDateDbToCompany(lgObjRs("retire_dt"),"")%>"
            End If
            
            .frm1.txtInsurGrade.value = "<%=ConvSPChars(lgObjRs("insur_grade"))%>"
            .frm1.cboMedType.value = "<%=ConvSPChars(lgObjRs("med_type"))%>"
            .frm1.txtMedInsurNo.value = "<%=ConvSPChars(lgObjRs("med_insur_no"))%>"
            .frm1.txtSuppcnt.value = "<%=ConvSPChars(lgObjRs("internal_supp_cnt"))%>"
            .frm1.txtMedAcqDt.text = "<%=UniConvDateDbToCompany(lgObjRs("med_acq_dt"),"")%>"            
            .frm1.txtMedLossDt.text = "<%=UniConvDateDbToCompany(lgObjRs("med_loss_dt"),"")%>"
            
            .frm1.cboSpouseAllow.value = "<%=ConvSPChars(lgObjRs("spouse_allow"))%>"
            .frm1.txtSupp.value = "<%=ConvSPChars(lgObjRs("supp_allow"))%>"
            
            .frm1.txtAnutGrade.value = "<%=ConvSPChars(lgObjRs("anut_grade"))%>"
            .frm1.txtAnutNo.value = "<%=ConvSPChars(lgObjRs("anut_no"))%>"
            .frm1.txtAnutAcqDt.text = "<%=UniConvDateDbToCompany(lgObjRs("anut_acq_dt"),"")%>"            
            .frm1.txtAnutLossDt.text = "<%=UniConvDateDbToCompany(lgObjRs("anut_loss_dt"),"")%>"
            
            .frm1.cboPayCd.value = "<%=ConvSPChars(lgObjRs("pay_cd"))%>"
            .frm1.txtAnnualSal.text = "<%=UNINumClientFormat(lgObjRs("annual_sal"), ggAmtOfMoney.DecPoint,0)%>"
            .frm1.txtSalary.text = "<%=UNINumClientFormat(lgObjRs("salary"), ggAmtOfMoney.DecPoint,0)%>"
            .frm1.txtBonusSalary.text = "<%=UNINumClientFormat(lgObjRs("bonus_salary"), ggAmtOfMoney.DecPoint,0)%>"            
            .frm1.cboTaxCd.value = "<%=ConvSPChars(lgObjRs("tax_cd"))%>"
            .frm1.txtBank.value = "<%=ConvSPChars(lgObjRs("bank"))%>"
            .frm1.txtBankNm.value = "<%=ConvSPChars(FuncCodeName(6, "", replace(lgObjRs("bank"),"'","''")))%>"
            .frm1.txtAccntNo.value = "<%=ConvSPChars(lgObjRs("bank_accnt"))%>"
            .frm1.txtBank2.value = "<%=ConvSPChars(lgObjRs("bank2"))%>"
            .frm1.txtBankNm2.value = "<%=ConvSPChars(FuncCodeName(6, "", replace(lgObjRs("bank2"),"'","''")))%>"
            .frm1.txtAccntNo2.value = "<%=ConvSPChars(lgObjRs("bank_accnt2"))%>"
            .frm1.txtBank3.value = "<%=ConvSPChars(lgObjRs("bank3"))%>"
            .frm1.txtBankNm3.value = "<%=ConvSPChars(FuncCodeName(6, "", replace(lgObjRs("bank3"),"'","''")))%>"
            .frm1.txtAccntNo3.value = "<%=ConvSPChars(lgObjRs("bank_accnt3"))%>"
            .frm1.txtBankMaster.value = "<%=ConvSPChars(lgObjRs("BankMaster"))%>"
            .frm1.txtBankMaster2.value = "<%=ConvSPChars(lgObjRs("BankMaster2"))%>"
            .frm1.txtBankMaster3.value = "<%=ConvSPChars(lgObjRs("BankMaster3"))%>"
            
            .frm1.txtSexCd.value = "<%=ConvSPChars(lgObjRs("sex_cd"))%>"
            .frm1.txtResNo.value = "<%=ConvSPChars(lgObjRs("res_no"))%>"
                        
            If "<%=ConvSPChars(lgObjRs("trade_union"))%>" = "Y" Then   '노조 
                .frm1.rdoUnionFlag1.value = "Y"
                .frm1.rdoUnionFlag1.Click()
            Else
                .frm1.rdoUnionFlag2.value = "N"
                .frm1.rdoUnionFlag2.Click()
            End If
            
            If "<%=ConvSPChars(lgObjRs("press_gubun"))%>" = "Y" Then   '기자 
                .frm1.rdoPressFlag1.value = "Y"
                .frm1.rdoPressFlag1.Click()
            Else
                .frm1.rdoPressFlag2.value = "N"
                .frm1.rdoPressFlag2.Click()
            End If
            
            If "<%=ConvSPChars(lgObjRs("oversea_labor_gubun"))%>" = "Y" Then   '기자 
                .frm1.rdoOverseaFlag1.value = "Y"
                .frm1.rdoOverseaFlag1.Click()
            Else
                .frm1.rdoOverseaFlag2.value = "N"
                .frm1.rdoOverseaFlag2.Click()
            End If
            

            If "<%=ConvSPChars(lgObjRs("Bank_Flag"))%>" = "1" Then   '은행사용여부 
                .frm1.rdoBankFlag1.value = "1"
                .frm1.rdoBankFlag1.Click()
            Elseif "<%=ConvSPChars(lgObjRs("Bank_Flag"))%>" = "2" Then
                .frm1.rdoBankFlag2.value = "2"
                .frm1.rdoBankFlag2.Click()
            Else    
                .frm1.rdoBankFlag3.value = "3"
                .frm1.rdoBankFlag3.Click()
            End If
            
            If "<%=ConvSPChars(lgObjRs("res_flag"))%>" = "Y" Then   '거주구분 
                .frm1.rdoResFlag1.value = "Y"
                .frm1.rdoResFlag1.Click()
            Else
                .frm1.rdoResFlag2.value = "N"
                .frm1.rdoResFlag2.Click()
            End If
                        
            If "<%=ConvSPChars(lgObjRs("prov_type"))%>" = "Y" Then
                .frm1.chkPayFlg.checked = true
            Else
                .frm1.chkPayFlg.checked = false
            End If
           
            If "<%=ConvSPChars(lgObjRs("FOREIGN_SEPARATE_TAX_YN"))%>" = "Y" Then '2004 외국인근로자분리과세적용여부 
                .frm1.txtForeign_separate_tax_yn.checked = true
            Else
                .frm1.txtForeign_separate_tax_yn.checked = false
            End If

            If "<%=ConvSPChars(lgObjRs("FOREIGN_NO_TAX_YN"))%>" = "Y" Then '2006 외국인근로자면세적용여부 
                .frm1.txtForeign_no_tax_yn.checked = true
            Else
                .frm1.txtForeign_no_tax_yn.checked = false
            End If

            If "<%=ConvSPChars(lgObjRs("employ_insur"))%>" = "Y" Then
                .frm1.chkEmpInsurFlg.checked = true
            Else
                .frm1.chkEmpInsurFlg.checked = false
            End If
            
            If "<%=ConvSPChars(lgObjRs("year_calcu"))%>" = "Y" Then
                .frm1.chkYearFlg.checked = true
            Else
                .frm1.chkYearFlg.checked = false
            End If
            
            If "<%=ConvSPChars(lgObjRs("retire_give"))%>" = "Y" Then
                .frm1.chkRetireFlg.checked = true
            Else
                .frm1.chkRetireFlg.checked = false
            End If
            
            If "<%=ConvSPChars(lgObjRs("tax_calcu"))%>" = "Y" Then
                .frm1.chkTaxFlg.checked = true
            Else
                .frm1.chkTaxFlg.checked = false
            End If
            
            If "<%=ConvSPChars(lgObjRs("year_mon_give"))%>" = "Y" Then
                .frm1.chkYearTaxFlg.checked = true
            Else
                .frm1.chkYearTaxFlg.checked = false
            End If
                        
            If "<%=ConvSPChars(lgObjRs("spouse"))%>" = "Y" Then
                .frm1.chkSpouseFlg.checked = true
            Else
                .frm1.chkSpouseFlg.checked = false
            End If
            
            If "<%=ConvSPChars(lgObjRs("lady"))%>" = "Y" Then
                .frm1.chkLadyFlg.checked = true
            Else
                .frm1.chkLadyFlg.checked = false
            End If

            .frm1.txtChild.value = "<%=ConvSPChars(lgObjRs("chl_rear"))%>"
            .frm1.txtOld.value = "<%=ConvSPChars(lgObjRs("supp_old_cnt"))%>"
            .frm1.txtYoung.value = "<%=ConvSPChars(lgObjRs("supp_young_cnt"))%>"
            .frm1.txtParia.value = "<%=ConvSPChars(lgObjRs("paria_cnt"))%>"
            .frm1.txtOldCnt1.value = "<%=ConvSPChars(lgObjRs("old_cnt"))%>"		'2004 경로자(65세이상)
            .frm1.txtOldCnt2.value = "<%=ConvSPChars(lgObjRs("old_cnt2"))%>"	'2004 경로자(70세이상)
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

    lgStrSQL = "DELETE  HDF020T"
    lgStrSQL = lgStrSQL & " WHERE emp_no = " & lgEmpNo                              ' 사번char(10)

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
	
End Sub

'============================================================================================================
' Name : SubBizSaveSingleUpdate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveSingleUpdate()
    
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    lgStrSQL = "UPDATE  HDF020T"
    lgStrSQL = lgStrSQL & " SET " 
    
    lgStrSQL = lgStrSQL & " insur_grade = " & FilterVar(Request("txtInsurGrade"), "''", "S") & ","	   '의보등급 
    lgStrSQL = lgStrSQL & " med_type = " & FilterVar(Request("cboMedType"), "''", "S") & ","			   '의보구분 
    lgStrSQL = lgStrSQL & " med_insur_no = " & FilterVar(Request("txtMedInsurNo"), "''", "S") & ","	   '의보번호 
    lgStrSQL = lgStrSQL & " internal_supp_cnt = " & UNIConvNum(Request("txtSuppCnt"),0) & ","          '의보부양자 
    lgStrSQL = lgStrSQL & " spouse_allow = " & FilterVar(Request("cboSpouseAllow"), "''", "S") & ","     '배우자 
    lgStrSQL = lgStrSQL & " supp_allow = " & UNIConvNum(Request("txtSupp"),0) & ","                    '부양자 
    lgStrSQL = lgStrSQL & " anut_grade = " & FilterVar(Request("txtAnutGrade"), "''", "S") & ","         '국민연금등급 
    
    If Request("txtAnutNo") = "" Then
		lgStrSQL = lgStrSQL & " anut_no = " & FilterVar(Request("txtResNo"), "''", "S") & ","            '국민연금번호 
	Else
		lgStrSQL = lgStrSQL & " anut_no = " & FilterVar(Request("txtAnutNo"), "''", "S") & ","           '국민연금번호 
	End If

	lgStrSQL = lgStrSQL & " med_acq_dt = " & FilterVar(UNIConvDateCompanyToDB(Request("txtMedAcqDt"),NULL),"NULL","S") & ","      '의료보험취득일 
	lgStrSQL = lgStrSQL & " med_loss_dt = " & FilterVar(UNIConvDateCompanyToDB(Request("txtMedLossDt"),NULL),"NULL","S") & ","    '의료보험상실일 
	lgStrSQL = lgStrSQL & " anut_acq_dt = " & FilterVar(UNIConvDateCompanyToDB(Request("txtAnutAcqDt"),NULL),"NULL","S") & ","    '국민연금취득일 
	lgStrSQL = lgStrSQL & " anut_loss_dt = " & FilterVar(UNIConvDateCompanyToDB(Request("txtAnutLossDt"),NULL),"NULL","S") & ","  '국민연금상실일 
	
    lgStrSQL = lgStrSQL & " pay_cd = " & FilterVar(Request("cboPayCd"), "''", "S") & ","                 '급여구분 
    lgStrSQL = lgStrSQL & " annual_sal = " & UNIConvNum(Request("txtAnnualSal"),0) & ","               '연봉 
    lgStrSQL = lgStrSQL & " salary = " & UNIConvNum(Request("txtSalary"),0) & ","                      '기본급 
    lgStrSQL = lgStrSQL & " bonus_salary = " & UNIConvNum(Request("txtBonusSalary"),0) & ","           '상여기준금 
    lgStrSQL = lgStrSQL & " tax_cd = " & FilterVar(Request("cboTaxCd"), "''", "S") & ","                 '세액구분 

    lgStrSQL = lgStrSQL & " bank = " & FilterVar(Request("txtBank"), "''", "S") & ","                    '은행 
    lgStrSQL = lgStrSQL & " bank_accnt = " & FilterVar(Request("txtAccntNo"), "''", "S") & ","           '계좌번호 
    lgStrSQL = lgStrSQL & " bank2 = " & FilterVar(Request("txtBank2"), "''", "S") & ","                  '은행2
    lgStrSQL = lgStrSQL & " bank_accnt2 = " & FilterVar(Request("txtAccntNo2"), "''", "S") & ","         '계좌번호2
    lgStrSQL = lgStrSQL & " bank3 = " & FilterVar(Request("txtBank3"), "''", "S") & ","                  '은행3
    lgStrSQL = lgStrSQL & " bank_accnt3 = " & FilterVar(Request("txtAccntNo3"), "''", "S") & ","         '계좌번호3
    lgStrSQL = lgStrSQL & " Bank_Flag = " & FilterVar(Request("rdoBankFlag"), "''", "S") & ","           '은행사용여부 
    
    lgStrSQL = lgStrSQL & " BankMaster = " & FilterVar(Request("txtBankMaster"), "''", "S") & ","        '계좌주 
    lgStrSQL = lgStrSQL & " BankMaster2 = " & FilterVar(Request("txtBankMaster2"), "''", "S") & ","      '계좌주2
    lgStrSQL = lgStrSQL & " BankMaster3 = " & FilterVar(Request("txtBankMaster3"), "''", "S") & ","      '계좌주3
                
    lgStrSQL = lgStrSQL & " chl_rear = " & UNIConvNum(Request("txtChild"),0) & ","                     '자녀양육수 
    lgStrSQL = lgStrSQL & " supp_old_cnt = " & UNIConvNum(Request("txtOld"),0) & ","                   '부양자(대)
    lgStrSQL = lgStrSQL & " supp_young_cnt = " & UNIConvNum(Request("txtYoung"),0) & ","               '부양자(소)
    lgStrSQL = lgStrSQL & " paria_cnt = " & UNIConvNum(Request("txtParia"),0) & ","                    '장애자 
    lgStrSQL = lgStrSQL & " old_cnt = " & UNIConvNum(Request("txtOldCnt1"),0) & ","                     '2004 경로자(65세이상)
    lgStrSQL = lgStrSQL & " old_cnt2 = " & UNIConvNum(Request("txtOldCnt2"),0) & ","                     '2004 경로자(70세이상)

    lgStrSQL = lgStrSQL & " trade_union = " & FilterVar(Request("rdoUnionFlag"), "''", "S") & ","            '노조구분 
    lgStrSQL = lgStrSQL & " press_gubun = " & FilterVar(Request("rdoPressFlag"), "''", "S") & ","            '기자구분 
    lgStrSQL = lgStrSQL & " oversea_labor_gubun = " & FilterVar(Request("rdoOverseaFlag"), "''", "S") & ","  '국외근로자 
    lgStrSQL = lgStrSQL & " res_flag = " & FilterVar(Request("rdoResFlag"), "''", "S") & ","                 '거주구분 
        
    If IsEmpty(Request("chkPayFlg")) = True Then			'임금지급대상 
        lgStrSQL = lgStrSQL & " prov_type = 'N' ,"
    Else
        lgStrSQL = lgStrSQL & " prov_type = 'Y' ,"
    End If
    
    If IsEmpty(Request("txtForeign_separate_tax_yn")) = True Then			'2004 외국인근로자분리과세적용여부 
        lgStrSQL = lgStrSQL & " FOREIGN_SEPARATE_TAX_YN = 'N' ,"
    Else
        lgStrSQL = lgStrSQL & " FOREIGN_SEPARATE_TAX_YN = 'Y' ,"
    End If 
       
    If IsEmpty(Request("txtForeign_no_tax_yn")) = True Then			'2006 외국인근로자면세적용여부 
        lgStrSQL = lgStrSQL & " FOREIGN_NO_TAX_YN = 'N' ,"
    Else
        lgStrSQL = lgStrSQL & " FOREIGN_NO_TAX_YN = 'Y' ,"
    End If 
        
    If IsEmpty(Request("chkEmpInsurFlg")) = True Then		'고용보험 
        lgStrSQL = lgStrSQL & " employ_insur = 'N' ,"
    Else
        lgStrSQL = lgStrSQL & " employ_insur = 'Y' ,"
    End If
    
    If IsEmpty(Request("chkYearFlg")) = True Then			'연월차지급대상 
        lgStrSQL = lgStrSQL & " year_calcu = 'N' ,"
    Else
        lgStrSQL = lgStrSQL & " year_calcu = 'Y' ,"
    End If
    
    If IsEmpty(Request("chkRetireFlg")) = True Then			'퇴직금지급대상 
        lgStrSQL = lgStrSQL & " retire_give = 'N' ,"
    Else
        lgStrSQL = lgStrSQL & " retire_give = 'Y' ,"
    End If
    
    If IsEmpty(Request("chkTaxFlg")) = True Then			'세액계산대상 
        lgStrSQL = lgStrSQL & " tax_calcu = 'N' ,"
    Else
        lgStrSQL = lgStrSQL & " tax_calcu = 'Y' ,"
    End If
    
    If IsEmpty(Request("chkYearTaxFlg")) = True Then		'연말정산신고대상 
        lgStrSQL = lgStrSQL & " year_mon_give = 'N' ,"
    Else
        lgStrSQL = lgStrSQL & " year_mon_give = 'Y' ,"
    End If
    
    If IsEmpty(Request("chkSpouseFlg")) = True Then			'소득공제(배우자)
        lgStrSQL = lgStrSQL & " spouse = 'N' ,"
    Else
        lgStrSQL = lgStrSQL & " spouse = 'Y' ,"
    End If
    
    If IsEmpty(Request("chkLadyFlg")) = True Then			'소득공제(부녀자)
        lgStrSQL = lgStrSQL & " lady = 'N' ,"
    Else
        lgStrSQL = lgStrSQL & " lady = 'Y' ,"
    End If
    
    lgStrSQL = lgStrSQL & " updt_emp_no =  " & FilterVar(gUsrId , "''", "S") & "," 
    lgStrSQL = lgStrSQL & " updt_dt = " & FilterVar(GetSvrDateTime, "''", "S")
    lgStrSQL = lgStrSQL & " WHERE emp_no = " & lgEmpNo

    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
	Call SubHandleError("SU",lgObjConn,lgObjRs,Err)
    
End Sub

'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements(pMode,pCode, pCode1)
    Select Case pMode 
      Case "R"
             Select Case  lgPrevNext 
	                 Case ""
                           lgStrSQL = "Select a.emp_no, a.name, a.dept_cd, b.sex_cd, b.res_no, " 
                           lgStrSQL = lgStrSQL & "a.pay_cd, a.annual_sal, a.salary, a.bonus_salary, "
                           lgStrSQL = lgStrSQL & "a.tax_cd, a.bank, a.bank_accnt, "
                           lgStrSQL = lgStrSQL & "a.bank2, a.bank_accnt2, a.bank3, a.bank_accnt3, a.bank_flag, "
                           lgStrSQL = lgStrSQL & "a.bankmaster, a.bankmaster2, a.bankmaster3, "
                           
                           lgStrSQL = lgStrSQL & "dbo.ufn_GetDeptName(a.dept_cd, GetDate()) dept_nm, "
                           lgStrSQL = lgStrSQL & "dbo.ufn_GetCodeName('H0001', a.pay_grd1) pay_grd1_nm, a.pay_grd2, "
                           lgStrSQL = lgStrSQL & "dbo.ufn_GetCodeName('H0003', a.ocpt_type) ocpt_type_nm, "
                           lgStrSQL = lgStrSQL & "dbo.ufn_GetCodeName('H0002', a.roll_pstn) roll_pstn_nm, "
                           lgStrSQL = lgStrSQL & "dbo.ufn_GetCodeName('H0004', a.func_cd) func_cd_nm, "
                           lgStrSQL = lgStrSQL & "dbo.ufn_GetCodeName('H0026', a.role_cd) role_cd_nm, "
                           lgStrSQL = lgStrSQL & "dbo.ufn_GetCodeName('H0016', a.entr_cd) entr_cd_nm, "
                           lgStrSQL = lgStrSQL & "a.group_entr_dt, a.entr_dt, a.retire_dt, a.intern_dt, a.rest_dt, "
                           lgStrSQL = lgStrSQL & "a.career_mm, a.insur_grade, a.med_type, a.med_insur_no, a.internal_supp_cnt, "
                           lgStrSQL = lgStrSQL & "a.med_acq_dt, a.med_loss_dt, a.spouse_allow, supp_allow, "
                           lgStrSQL = lgStrSQL & "a.anut_grade, a.anut_no, a.anut_acq_dt, anut_loss_dt, "
                           lgStrSQL = lgStrSQL & "a.trade_union, a.press_gubun, a.oversea_labor_gubun, res_flag, FOREIGN_SEPARATE_TAX_YN,FOREIGN_NO_TAX_YN,"
                           lgStrSQL = lgStrSQL & "a.prov_type, a.employ_insur, a.year_calcu, retire_give, "
                           lgStrSQL = lgStrSQL & "a.tax_calcu, a.year_mon_give, a.spouse, a.lady, a.chl_rear, "
                           lgStrSQL = lgStrSQL & "a.supp_old_cnt, a.supp_young_cnt, a.paria_cnt, a.old_cnt, a.old_cnt2"
                           lgStrSQL = lgStrSQL & " From  HDF020T a, HAA010T b "
                           lgStrSQL = lgStrSQL & " WHERE a.emp_no = b.emp_no AND a.emp_no = " & pCode 

                     Case "P"
                           lgStrSQL = "Select a.emp_no, a.name, a.dept_cd, b.sex_cd, b.res_no, " 
                           lgStrSQL = lgStrSQL & "a.pay_cd, a.annual_sal, a.salary, a.bonus_salary, "
                           lgStrSQL = lgStrSQL & "a.tax_cd, a.bank, a.bank_accnt, "
                           lgStrSQL = lgStrSQL & "a.bank2, a.bank_accnt2, a.bank3, a.bank_accnt3, a.bank_flag, "
                           lgStrSQL = lgStrSQL & "a.bankmaster, a.bankmaster2, a.bankmaster3, "
                           
                           lgStrSQL = lgStrSQL & "dbo.ufn_GetDeptName(a.dept_cd, GetDate()) dept_nm, "
                           lgStrSQL = lgStrSQL & "dbo.ufn_GetCodeName('H0001', a.pay_grd1) pay_grd1_nm, a.pay_grd2, "
                           lgStrSQL = lgStrSQL & "dbo.ufn_GetCodeName('H0003', a.ocpt_type) ocpt_type_nm, "
                           lgStrSQL = lgStrSQL & "dbo.ufn_GetCodeName('H0002', a.roll_pstn) roll_pstn_nm, "
                           lgStrSQL = lgStrSQL & "dbo.ufn_GetCodeName('H0004', a.func_cd) func_cd_nm, "
                           lgStrSQL = lgStrSQL & "dbo.ufn_GetCodeName('H0026', a.role_cd) role_cd_nm, "
                           lgStrSQL = lgStrSQL & "dbo.ufn_GetCodeName('H0016', a.entr_cd) entr_cd_nm, "
                           lgStrSQL = lgStrSQL & "a.group_entr_dt, a.entr_dt, a.retire_dt, a.intern_dt, a.rest_dt, "
                           lgStrSQL = lgStrSQL & "a.career_mm, a.insur_grade, a.med_type, a.med_insur_no, a.internal_supp_cnt, "
                           lgStrSQL = lgStrSQL & "a.med_acq_dt, a.med_loss_dt, a.spouse_allow, supp_allow, "
                           lgStrSQL = lgStrSQL & "a.anut_grade, a.anut_no, a.anut_acq_dt, anut_loss_dt, "
                           lgStrSQL = lgStrSQL & "a.trade_union, a.press_gubun, a.oversea_labor_gubun, res_flag, "
                           lgStrSQL = lgStrSQL & "a.prov_type, a.employ_insur, a.year_calcu, retire_give, "
                           lgStrSQL = lgStrSQL & "a.tax_calcu, a.year_mon_give, a.spouse, a.lady, a.chl_rear, "
                           lgStrSQL = lgStrSQL & "a.supp_old_cnt, a.supp_young_cnt, a.paria_cnt, a.old_cnt"
                           lgStrSQL = lgStrSQL & " From  HDF020T a, HAA010T b "
                           lgStrSQL = lgStrSQL & " WHERE a.emp_no = b.emp_no AND a.emp_no < " & pCode 
                           lgStrSQL = lgStrSQL & " and b.internal_cd LIKE  " & FilterVar(pCode1 & "%", "''", "S") & ""
                           lgStrSQL = lgStrSQL & " AND (b.retire_resn is null or "
                           lgStrSQL = lgStrSQL & " b.retire_resn = '' or b.retire_resn = '6') "	
                           lgStrSQL = lgStrSQL & " ORDER BY a.emp_no DESC "
                     Case "N"
                           lgStrSQL = "Select a.emp_no, a.name, a.dept_cd, b.sex_cd, b.res_no, " 
                           lgStrSQL = lgStrSQL & "a.pay_cd, a.annual_sal, a.salary, a.bonus_salary, "
                           lgStrSQL = lgStrSQL & "a.tax_cd, a.bank, a.bank_accnt, "
                           lgStrSQL = lgStrSQL & "a.bank2, a.bank_accnt2, a.bank3, a.bank_accnt3, a.bank_flag, "
                           lgStrSQL = lgStrSQL & "a.bankmaster, a.bankmaster2, a.bankmaster3, "
                           
                           lgStrSQL = lgStrSQL & "dbo.ufn_GetDeptName(a.dept_cd, GetDate()) dept_nm, "
                           lgStrSQL = lgStrSQL & "dbo.ufn_GetCodeName('H0001', a.pay_grd1) pay_grd1_nm, a.pay_grd2, "
                           lgStrSQL = lgStrSQL & "dbo.ufn_GetCodeName('H0003', a.ocpt_type) ocpt_type_nm, "
                           lgStrSQL = lgStrSQL & "dbo.ufn_GetCodeName('H0002', a.roll_pstn) roll_pstn_nm, "
                           lgStrSQL = lgStrSQL & "dbo.ufn_GetCodeName('H0004', a.func_cd) func_cd_nm, "
                           lgStrSQL = lgStrSQL & "dbo.ufn_GetCodeName('H0026', a.role_cd) role_cd_nm, "
                           lgStrSQL = lgStrSQL & "dbo.ufn_GetCodeName('H0016', a.entr_cd) entr_cd_nm, "
                           lgStrSQL = lgStrSQL & "a.group_entr_dt, a.entr_dt, a.retire_dt, a.intern_dt, a.rest_dt, "
                           lgStrSQL = lgStrSQL & "a.career_mm, a.insur_grade, a.med_type, a.med_insur_no, a.internal_supp_cnt, "
                           lgStrSQL = lgStrSQL & "a.med_acq_dt, a.med_loss_dt, a.spouse_allow, supp_allow, "
                           lgStrSQL = lgStrSQL & "a.anut_grade, a.anut_no, a.anut_acq_dt, anut_loss_dt, "
                           lgStrSQL = lgStrSQL & "a.trade_union, a.press_gubun, a.oversea_labor_gubun, res_flag, "
                           lgStrSQL = lgStrSQL & "a.prov_type, a.employ_insur, a.year_calcu, retire_give, "
                           lgStrSQL = lgStrSQL & "a.tax_calcu, a.year_mon_give, a.spouse, a.lady, a.chl_rear, "
                           lgStrSQL = lgStrSQL & "a.supp_old_cnt, a.supp_young_cnt, a.paria_cnt, a.old_cnt"
                           lgStrSQL = lgStrSQL & " From  HDF020T a, HAA010T b "
                           lgStrSQL = lgStrSQL & " WHERE a.emp_no = b.emp_no AND a.emp_no > " & pCode
                           lgStrSQL = lgStrSQL & " and b.internal_cd LIKE  " & FilterVar(pCode1 & "%", "''", "S") & ""
                           lgStrSQL = lgStrSQL & " AND (b.retire_resn is null or "
                           lgStrSQL = lgStrSQL & " b.retire_resn = '' or b.retire_resn = '6') "	
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
    End Select
End Sub

%>

<Script Language="VBScript">
    Select Case "<%=lgOpModeCRUD %>"
       Case "<%=UID_M0001%>"		
          If Trim("<%=lgErrorStatus%>") = "NO" Then			
          	 Parent.frm1.txtHEmpNo.value = Parent.frm1.txtEmpNo.value
			 Parent.frm1.txtHEmpNm.value = Parent.frm1.txtEmpNm.value          	
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
