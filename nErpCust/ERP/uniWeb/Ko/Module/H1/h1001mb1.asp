<%@ LANGUAGE=VBSCript TRANSACTION=Required%>
<% Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/lgSvrVariables.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->

<%    
    Call HideStatusWnd                                                               '☜: Hide Processing message
    Call LoadBasisGlobalInf()  
    Call LoadInfTB19029B("I", "H","NOCOOKIE","MB")

    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '☜: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)
    lgKeyStream       = Split(Request("txtKeyStream"),gColSep)

    Call SubOpenDB(lgObjConn)                                                        '☜: Make a DB Connection

    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)
             Call SubBizQuery()
        Case CStr(UID_M0002)
             Call SubBizSave()
        Case CStr(UID_M0003)
             Call SubBizDelete()
    End Select

    Call SubCloseDB(lgObjConn)

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()
    Dim iKey1

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    iKey1 = FilterVar(lgKeyStream(0), "''", "S")

    Call SubMakeSQLStatements("R",iKey1)                                       '☜ : Make sql statements

    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
          Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
          Call SetErrorStatus()
    Else
%>
<Script Language=vbscript>
       With Parent	
                .Frm1.cboFamily_type.Value  = "<%=ConvSPChars(lgObjRs("family_type"))%>"
                .Frm1.cboSave_script_type.Value  = "<%=ConvSPChars(lgObjRs("Save_script_type"))%>"
                .Frm1.cboIntern_type.Value  = "<%=ConvSPChars(lgObjRs("Intern_type"))%>"
                .Frm1.cboBas_strt_mm.Value  = "<%=ConvSPChars(lgObjRs("bas_strt_mm"))%>"
                .Frm1.txtbas_strt_dd.Text  = "<%=ConvSPChars(lgObjRs("bas_strt_dd"))%>"
                .Frm1.cboBas_end_mm.Value  = "<%=ConvSPChars(lgObjRs("bas_end_mm"))%>"
                .Frm1.txtbas_end_dd.Text  = "<%=ConvSPChars(lgObjRs("bas_end_dd"))%>"

                .Frm1.txtMed_comp_rate.text  = "<%=UNINumClientFormat(lgObjRs("Med_comp_rate"), 3, 0)%>"
                .Frm1.txtmed_prsn_rate.text  = "<%=UNINumClientFormat(lgObjRs("med_prsn_rate"), 3, 0)%>"

                if "<%=ConvSPChars(lgObjRs("med_entr_flag"))%>" = "Y" then
                    .Frm1.txtmed_entr_flag1.checked = true
                else
                    .Frm1.txtmed_entr_flag2.checked = true
                end if

                if "<%=ConvSPChars(lgObjRs("med_retire_flag"))%>" = "Y" then
                    .Frm1.txtmed_retire_flag1.checked = true
                else
                    .Frm1.txtmed_retire_flag2.checked = true
                end if
                if "<%=ConvSPChars(lgObjRs("med_en_re_flag"))%>" = "Y" then
                    .Frm1.txtmed_en_re_flag1.checked = true
                else
                    .Frm1.txtmed_en_re_flag2.checked = true
                end if

				.Frm1.txtanut_comp_rate1.text  = "<%=UNINumClientFormat(lgObjRs("anut_comp_rate1"), ggQty.DecPoint, 0)%>"
                .Frm1.txtanut_comp_rate2.Text  = "<%=UNINumClientFormat(lgObjRs("anut_comp_rate2"), ggQty.DecPoint, 0)%>"
                .Frm1.txtanut_prsn_rate1.Text  = "<%=UNINumClientFormat(lgObjRs("anut_prsn_rate1"), ggQty.DecPoint, 0)%>"
                .Frm1.txtanut_prsn_rate2.Text  = "<%=UNINumClientFormat(lgObjRs("anut_prsn_rate2"), ggQty.DecPoint, 0)%>"
                .Frm1.txtanut_retire_rate1.Text  = "<%=UNINumClientFormat(lgObjRs("anut_retire_rate1"), ggQty.DecPoint, 0)%>"
                .Frm1.txtanut_retire_rate2.Text  = "<%=UNINumClientFormat(lgObjRs("anut_retire_rate2"), ggQty.DecPoint, 0)%>"

                if "<%=ConvSPChars(lgObjRs("anut_entr_flag"))%>" = "Y" then
                    .Frm1.txtanut_entr_flag1.checked = true
                else
                    .Frm1.txtanut_entr_flag2.checked = true
                end if

                if "<%=ConvSPChars(lgObjRs("anut_retire_flag"))%>" = "Y" then
                    .Frm1.txtanut_retire_flag1.checked = true
                else
                    .Frm1.txtanut_retire_flag2.checked = true
                end if

                if "<%=ConvSPChars(lgObjRs("anut_en_re_flag"))%>" = "Y" then
                    .Frm1.txtanut_en_re_flag1.checked = true
                else
                    .Frm1.txtanut_en_re_flag2.checked = true
                end if
                
                if "<%=ConvSPChars(lgObjRs("med_type"))%>" = "1" then
                    .Frm1.txtMed_type1.checked = true
                else
                    .Frm1.txtMed_type2.checked = true
                end if
                
                

				.Frm1.txtemploy_rate.Text  = "<%=UNINumClientFormat(lgObjRs("employ_rate"), ggQty.DecPoint, 0)%>"
                .Frm1.txtre_tax_sub1.Text  = "<%=UNINumClientFormat(lgObjRs("re_tax_sub1"), ggAmtOfMoney.DecPoint, 0)%>"
                .Frm1.txtre_incom_sub.Text  = "<%=UNINumClientFormat(lgObjRs("re_incom_sub"), ggAmtOfMoney.DecPoint, 0)%>"
                .Frm1.txtre_speci_sub.Text  = "<%=UNINumClientFormat(lgObjRs("re_speci_sub"), ggQty.DecPoint, 0)%>"
                .Frm1.txtre_sub_limit.Text  = "<%=UNINumClientFormat(lgObjRs("re_sub_limit"), ggQty.DecPoint, 0)%>"
				.Frm1.txtpay_prov_dd.Text  = "<%=ConvSPChars(lgObjRs("pay_prov_dd"))%>"
                .Frm1.txtpay_bas_dd.Text  = "<%=ConvSPChars(lgObjRs("pay_bas_dd"))%>"
                .Frm1.txtdilig_dd.Text  = "<%=ConvSPChars(lgObjRs("dilig_dd"))%>"
'                .Frm1.txtinsur_dt.Text  = "<%=UNIMonthClientFormat(lgObjRs("insur_dt"))%>"
       End With          
</Script>       
<%     
    End If
    Call SubCloseRs(lgObjRs)
End Sub	
'============================================================================================================
' Name : SubBizQuery
' Desc : Date data 
'============================================================================================================
Sub SubBizSave()
    On Error Resume Next
    Err.Clear
    
    lgIntFlgMode = CInt(Request("txtFlgMode"))                                       '☜: Read Operayion Mode (CREATE, UPDATE)

    Select Case lgIntFlgMode
        Case  OPMD_CMODE                                                             '☜ : Create
              Call SubBizSaveSingleCreate()  
        Case  OPMD_UMODE           
              Call SubBizSaveSingleUpdate()
    End Select
End Sub	
	    
'============================================================================================================
' Name : SubBizDelete
' Desc : Delete DB data
'============================================================================================================
Sub SubBizDelete()
    On Error Resume Next
    Err.Clear

    lgStrSQL = "DELETE  HDA000T"
    lgStrSQL = lgStrSQL & " WHERE        "
    lgStrSQL = lgStrSQL & " COMP_CD   = " & FilterVar(lgKeyStream(0), "''", "S")

    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("SD",lgObjConn,lgObjRs,Err)
End Sub

'============================================================================================================
' Name : SubBizSaveCreate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveSingleCreate(lgObjRs)
    On Error Resume Next
    Err.Clear
End Sub

'============================================================================================================
' Name : SubBizSaveSingleUpdate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveSingleUpdate()
    On Error Resume Next
    Err.Clear

    lgStrSQL =            "UPDATE HDA000T"
    lgStrSQL = lgStrSQL & "   SET " 
    lgStrSQL = lgStrSQL & "       family_type = " & FilterVar(Request("cboFamily_type"), "''", "S") & ","
    lgStrSQL = lgStrSQL & "       Save_script_type = " & FilterVar(Request("cboSave_script_type"), "''", "S") & ","
    lgStrSQL = lgStrSQL & "       Intern_type = " & FilterVar(Request("cboIntern_type"), "''", "S") & ","
    lgStrSQL = lgStrSQL & "       bas_strt_mm = " & FilterVar(Request("cboBas_strt_mm"), "''", "S") & ","
    if  Cint(Request("txtbas_strt_dd")) < 10 then
        lgStrSQL = lgStrSQL & "       bas_strt_dd = " & FilterVar("0" & Request("txtbas_strt_dd"), "''", "S") & ","        
    else
        lgStrSQL = lgStrSQL & "       bas_strt_dd = " & FilterVar(Request("txtbas_strt_dd"),"" & FilterVar("00", "''", "S") & "","S") & ","
    end if
    lgStrSQL = lgStrSQL & "       bas_end_mm = " & FilterVar(Request("cboBas_end_mm"), "''", "S") & ","

    if  Cint(Request("txtbas_end_dd")) < 10 then
        lgStrSQL = lgStrSQL & "       bas_end_dd = " & FilterVar("0" & Request("txtbas_end_dd"), "''", "S") & ","        
    else
        lgStrSQL = lgStrSQL & "       bas_end_dd = " & FilterVar(Request("txtbas_end_dd"),"" & FilterVar("00", "''", "S") & "","S") & ","
    end if
    lgStrSQL = lgStrSQL & "       Med_comp_rate = " & UNIConvNum(Request("txtMed_comp_rate"),0) & ","
    lgStrSQL = lgStrSQL & "       med_prsn_rate = " & UNIConvNum(Request("txtmed_prsn_rate"),0) & ","
    
    lgStrSQL = lgStrSQL & "       med_entr_flag = " & FilterVar(Request("txtmed_entr_flag"), "''", "S") & ","
    lgStrSQL = lgStrSQL & "       med_retire_flag = " & FilterVar(Request("txtmed_retire_flag"), "''", "S") & ","
    lgStrSQL = lgStrSQL & "       med_en_re_flag = " & FilterVar(Request("txtmed_en_re_flag"), "''", "S") & ","

    lgStrSQL = lgStrSQL & "       anut_comp_rate1 = " & UNIConvNum(Request("txtanut_comp_rate1"),0) & ","
    lgStrSQL = lgStrSQL & "       anut_comp_rate2 = " & UNIConvNum(Request("txtanut_comp_rate2"),0) & ","
    lgStrSQL = lgStrSQL & "       anut_prsn_rate1 = " & UNIConvNum(Request("txtanut_prsn_rate1"),0) & ","
    lgStrSQL = lgStrSQL & "       anut_prsn_rate2 = " & UNIConvNum(Request("txtanut_prsn_rate2"),0) & ","
    lgStrSQL = lgStrSQL & "       anut_retire_rate1 = " & UNIConvNum(Request("txtanut_retire_rate1"),0) & ","
    lgStrSQL = lgStrSQL & "       anut_retire_rate2 = " & UNIConvNum(Request("txtanut_retire_rate2"),0) & ","

    lgStrSQL = lgStrSQL & "       anut_entr_flag = " & FilterVar(Request("txtanut_entr_flag"), "''", "S") & ","
    lgStrSQL = lgStrSQL & "       anut_retire_flag = " & FilterVar(Request("txtanut_retire_flag"), "''", "S") & ","
    lgStrSQL = lgStrSQL & "       anut_en_re_flag = " & FilterVar(Request("txtanut_en_re_flag"), "''", "S") & ","

    lgStrSQL = lgStrSQL & "       employ_rate = " & UNIConvNum(Request("txtemploy_rate"),0) & ","
    lgStrSQL = lgStrSQL & "       re_tax_sub1 = " & UNIConvNum(Request("txtre_tax_sub1"),0) & ","
    lgStrSQL = lgStrSQL & "       re_incom_sub = " & UNIConvNum(Request("txtre_incom_sub"),0) & ","
    lgStrSQL = lgStrSQL & "       re_speci_sub = " & UNIConvNum(Request("txtre_speci_sub"),0) & ","
    lgStrSQL = lgStrSQL & "       re_sub_limit = " & UNIConvNum(Request("txtre_sub_limit"),0) & ","

    if  Cint(Request("txtpay_prov_dd")) < 10 then
         lgStrSQL = lgStrSQL & "       pay_prov_dd = " & FilterVar("0" & Request("txtpay_prov_dd"), "''", "S") & ","       
    else
        lgStrSQL = lgStrSQL & "       pay_prov_dd = " & FilterVar(Request("txtpay_prov_dd"),"" & FilterVar("00", "''", "S") & "","S") & ","
    end if

    if  Cint(Request("txtpay_bas_dd")) < 10 then
        lgStrSQL = lgStrSQL & "       pay_bas_dd = " & FilterVar("0" & Request("txtpay_bas_dd"), "''", "S") & ","        
    else
        lgStrSQL = lgStrSQL & "       pay_bas_dd = " & FilterVar(Request("txtpay_bas_dd"),"" & FilterVar("00", "''", "S") & "","S") & ","
    end if

    if  Cint(Request("txtdilig_dd")) < 10 then
        lgStrSQL = lgStrSQL & "       dilig_dd = " & FilterVar("0" & Request("txtdilig_dd"), "''", "S")       
    else
        lgStrSQL = lgStrSQL & "       dilig_dd = " & FilterVar(Request("txtdilig_dd"),"" & FilterVar("00", "''", "S") & "","S") 
    end if
    '2007.02 건강보험기준 추가 lws
		lgStrSQL = lgStrSQL & "       ,MED_TYPE = " & FilterVar( Request("txtMed_type"), "''", "S") 
		
    lgStrSQL = lgStrSQL & " WHERE COMP_CD = " & FilterVar(lgKeyStream(0), "''", "S")

    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("SU",lgObjConn,lgObjRs,Err)
End Sub

'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements(pMode,pCode)
    Select Case pMode 
      Case "R"
            lgStrSQL = "            Select  "
            lgStrSQL = lgStrSQL & " family_type,       Save_script_type,    Intern_type,bas_strt_mm, "
            lgStrSQL = lgStrSQL & " bas_strt_dd,       bas_end_mm,          bas_end_dd, "
            lgStrSQL = lgStrSQL & " Med_comp_rate,     med_prsn_rate,       med_entr_flag, "
            lgStrSQL = lgStrSQL & " med_retire_flag,   med_en_re_flag,      anut_comp_rate1, "
            lgStrSQL = lgStrSQL & " anut_comp_rate2,   anut_prsn_rate1,     anut_prsn_rate2, "
            lgStrSQL = lgStrSQL & " anut_retire_rate1, anut_retire_rate2,   anut_entr_flag, "
            lgStrSQL = lgStrSQL & " anut_retire_flag,  anut_en_re_flag,     employ_rate, "
            lgStrSQL = lgStrSQL & " re_tax_sub1,       re_incom_sub,        re_speci_sub, "
            lgStrSQL = lgStrSQL & " re_sub_limit,      pay_prov_dd,         pay_bas_dd, "
            lgStrSQL = lgStrSQL & " dilig_dd,          insur_dt ,med_type" 
            lgStrSQL = lgStrSQL & "  From HDA000T "
            lgStrSQL = lgStrSQL & " WHERE COMP_CD = " & pCode 	
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
    lgErrorStatus     = "YES"
End Sub

'============================================================================================================
' Name : SubHandleError
' Desc : This Sub handle error
'============================================================================================================
Sub SubHandleError(pOpCode,pConn,pRs,pErr)
    On Error Resume Next
    Err.Clear

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
