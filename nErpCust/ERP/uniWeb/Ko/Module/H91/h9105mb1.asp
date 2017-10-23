<%@ LANGUAGE=VBSCript %>
<% Option Explicit%>

<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<!-- #Include file="../../inc/lgSvrVariables.inc" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<%

    Call HideStatusWnd                                                               '☜: Hide Processing message

    Call LoadBasisGlobalInf()
    Call LoadInfTB19029B("I", "H", "NOCOOKIE", "MB")

	dim lgGetSvrDateTime

    Call HideStatusWnd                                                               '☜: Hide Processing message

    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '☜: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)
    lgKeyStream       = Split(Request("txtKeyStream"),gColSep)

    lgPrevNext        = Request("txtPrevNext")                                       '☜: "P"(Prev search) "N"(Next search)

    Call SubOpenDB(lgObjConn)                                                        '☜: Make a DB Connection

    lgGetSvrDateTime = GetSvrDateTime
    
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
    Dim strWhere
    

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    strWhere =  FilterVar(lgKeyStream(0), "''", "S")
    strWhere =  strWhere  & "  AND YEAR_YY = " & FilterVar(lgKeyStream(1), "''","S")
    
    If Cint(lgKeyStream(2)) = 1 then
    Else
        If Cint(lgKeyStream(3)) <> 2 Then
            strWhere =  strWhere  & "  AND SEQ = 2 "
        End if
    End if
    
    Call SubMakeSQLStatements("R",strWhere)                                       '☜ : Make sql statements

    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then                    'If data not exists
        If lgPrevNext = "" Then
            Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
            If Cint(lgKeyStream(2)) = 1 then
                Call SetErrorStatus()
            Else
                lgKeyStream(2) = 1
                Call SubBizQuery()
            End if
          
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
    
        Do While Not lgObjRs.EOF
%>
<Script Language=vbscript>
        If  <%=ConvSPChars(lgObjRs("SEQ"))= 1 %> Then
            With Parent	
                 .frm1.txtA_comp_nm.value = "<%=ConvSPChars(lgObjRs("A_COMP_NM"))%>"         
                 .frm1.txtA_comp_no.value = "<%=ConvSPChars(lgObjRs("A_COMP_NO"))%>"         
                 .Frm1.txtA_pay_tot_amt.Value  = "<%=UNINumClientFormat(lgObjRs("A_PAY_TOT_AMT"), ggAmtOfMoney.DecPoint,0)%>"
                 .Frm1.txtA_bonus_tot_amt.Value  = "<%=UNINumClientFormat(lgObjRs("A_BONUS_TOT_AMT"), ggAmtOfMoney.DecPoint,0)%>"
                 .Frm1.txtA_after_bonus_amt.Value  = "<%=UNINumClientFormat(lgObjRs("A_AFTER_BONUS_AMT"), ggAmtOfMoney.DecPoint,0)%>"
                 .Frm1.txtA_med_insur_amt.Value  = "<%=UNINumClientFormat(lgObjRs("A_MED_INSUR"), ggAmtOfMoney.DecPoint,0)%>"
                 .frm1.txtA_national_pension_amt.value = "<%=UNINumClientFormat(lgObjRs("A_NATIONAL_PENSION_AMT"), ggAmtOfMoney.DecPoint,0)%>"         
                 .frm1.txtA_save_tax_sub_amt.value = "<%=UNINumClientFormat(lgObjRs("A_SAVE_TAX_SUB"), ggAmtOfMoney.DecPoint,0)%>"        
                 .frm1.txtA_indiv_anu_amt.value = "<%=UNINumClientFormat(lgObjRs("A_INDIV_ANU_AMT"), ggAmtOfMoney.DecPoint,0)%>"      
                 .frm1.txtA_indiv_anu2_amt.value = "<%=UNINumClientFormat(lgObjRs("A_INDIV_ANU2_AMT"), ggAmtOfMoney.DecPoint,0)%>"      
                 .frm1.txtA_income_tax_amt.value = "<%=UNINumClientFormat(lgObjRs("A_INCOME_TAX"), ggAmtOfMoney.DecPoint,0)%>"        
                 .frm1.txtA_res_tax_amt.value = "<%=UNINumClientFormat(lgObjRs("A_RES_TAX"), ggAmtOfMoney.DecPoint,0)%>"        
                 .Frm1.txtA_farm_tax_amt.Value  = "<%=UNINumClientFormat(lgObjRs("A_FARM_TAX"), ggAmtOfMoney.DecPoint,0)%>"
                 .Frm1.txtA_non_tax1_amt.Value = "<%=UNINumClientFormat(lgObjRs("A_NON_TAX1"), ggAmtOfMoney.DecPoint,0)%>"
                 .frm1.txtA_non_tax2_amt.value = "<%=UNINumClientFormat(lgObjRs("A_NON_TAX2"), ggAmtOfMoney.DecPoint,0)%>"         
                 .frm1.txtA_non_tax3_amt.value = "<%=UNINumClientFormat(lgObjRs("A_NON_TAX3"), ggAmtOfMoney.DecPoint,0)%>"          
                 .frm1.txtA_non_tax4_amt.value = "<%=UNINumClientFormat(lgObjRs("A_NON_TAX4"), ggAmtOfMoney.DecPoint,0)%>"        
                 .frm1.txtA_non_tax5_amt.value = "<%=UNINumClientFormat(lgObjRs("A_NON_TAX5"), ggAmtOfMoney.DecPoint,0)%>"      
            End With  
        Else
            With Parent
            
                 .frm1.txtA_comp_nm2.value = "<%=ConvSPChars(lgObjRs("A_COMP_NM"))%>"         
                 .frm1.txtA_comp_no2.value = "<%=ConvSPChars(lgObjRs("A_COMP_NO"))%>"         
                 .Frm1.txtA_pay_tot_amt2.Value  = "<%=UNINumClientFormat(lgObjRs("A_PAY_TOT_AMT"), ggAmtOfMoney.DecPoint,0)%>"
                 .Frm1.txtA_bonus_tot_amt2.Value  = "<%=UNINumClientFormat(lgObjRs("A_BONUS_TOT_AMT"), ggAmtOfMoney.DecPoint,0)%>"
                 .Frm1.txtA_after_bonus_amt2.Value  = "<%=UNINumClientFormat(lgObjRs("A_AFTER_BONUS_AMT"), ggAmtOfMoney.DecPoint,0)%>"
                 .Frm1.txtA_med_insur_amt2.Value  = "<%=UNINumClientFormat(lgObjRs("A_MED_INSUR"), ggAmtOfMoney.DecPoint,0)%>"
                 .frm1.txtA_national_pension_amt2.value = "<%=UNINumClientFormat(lgObjRs("A_NATIONAL_PENSION_AMT"), ggAmtOfMoney.DecPoint,0)%>"         
                 .frm1.txtA_save_tax_sub_amt2.value = "<%=UNINumClientFormat(lgObjRs("A_SAVE_TAX_SUB"), ggAmtOfMoney.DecPoint,0)%>"        
                 .frm1.txtA_indiv_anu_amt2.value = "<%=UNINumClientFormat(lgObjRs("A_INDIV_ANU_AMT"), ggAmtOfMoney.DecPoint,0)%>"      
                 .frm1.txtA_indiv_anu2_amt2.value = "<%=UNINumClientFormat(lgObjRs("A_INDIV_ANU2_AMT"), ggAmtOfMoney.DecPoint,0)%>"      
                 .frm1.txtA_income_tax_amt2.value = "<%=UNINumClientFormat(lgObjRs("A_INCOME_TAX"), ggAmtOfMoney.DecPoint,0)%>"        
                 .frm1.txtA_res_tax_amt2.value = "<%=UNINumClientFormat(lgObjRs("A_RES_TAX"), ggAmtOfMoney.DecPoint,0)%>"        
                 .Frm1.txtA_farm_tax_amt2.Value  = "<%=UNINumClientFormat(lgObjRs("A_FARM_TAX"), ggAmtOfMoney.DecPoint,0)%>"
                 .Frm1.txtA_non_tax1_amt2.Value = "<%=UNINumClientFormat(lgObjRs("A_NON_TAX1"), ggAmtOfMoney.DecPoint,0)%>"
                 .frm1.txtA_non_tax2_amt2.value = "<%=UNINumClientFormat(lgObjRs("A_NON_TAX2"), ggAmtOfMoney.DecPoint,0)%>"         
                 .frm1.txtA_non_tax3_amt2.value = "<%=UNINumClientFormat(lgObjRs("A_NON_TAX3"), ggAmtOfMoney.DecPoint,0)%>"          
                 .frm1.txtA_non_tax4_amt2.value = "<%=UNINumClientFormat(lgObjRs("A_NON_TAX4"), ggAmtOfMoney.DecPoint,0)%>"        
                 .frm1.txtA_non_tax5_amt2.value = "<%=UNINumClientFormat(lgObjRs("A_NON_TAX5"), ggAmtOfMoney.DecPoint,0)%>"      
            End With  
        End If 
               
</Script>       
<%
        lgObjRs.MoveNext
        Loop 

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

    lgStrSQL = "DELETE  HFA040T"
    lgStrSQL = lgStrSQL & " WHERE EMP_NO = " & FilterVar(lgKeyStream(0), "''", "S")
    lgStrSQL =  lgStrSQL  & "  AND YEAR_YY = " & FilterVar(lgKeyStream(1), "''", "S")
    If Cint(lgKeyStream(2)) = 1 then
        lgStrSQL =  lgStrSQL  & "  AND SEQ = 1 "
    Else
        lgStrSQL =  lgStrSQL  & "  AND SEQ = 2 "
    End if
    
    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
	Call SubHandleError("SD",lgObjConn,lgObjRs,Err)

    If lgKeyStream(2) = 1 Then           ' 종전근무지1이 삭제되면 
        Call SubBizDeleteUpdate()
    End if 
End Sub


'============================================================================================================
' Name : SubBizDeleteUpdate
' Desc : Delete DB data
'============================================================================================================
Sub SubBizDeleteUpdate()

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    'Call CommonQueryRs(" count(seq) "," hfa040t ","EMP_NO = " & FilterVar(trim(lgKeyStream(0)),"''", "S") & " and year_yy=" & FilterVar(trim(lgKeyStream(1)),"''", "S") ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    'If Replace(lgF0,Chr(11),"") = 2 Then     ' 두개의 종전근무지를 가지고 있던 사원의 
    '    If lgKeyStream(2) = 1 Then           ' 종전근무지1이 삭제되면 
            lgStrSQL = "UPDATE  HFA040T"     ' 종전근무지2의 SEQ를 2에서 1로 Update한다.
            lgStrSQL = lgStrSQL & " SET " 
            lgStrSQL = lgStrSQL & " SEQ = 1 "
            lgStrSQL =  lgStrSQL & " WHERE EMP_NO = " & FilterVar(lgKeyStream(0), "''", "S")
            lgStrSQL =  lgStrSQL  & "  AND YEAR_YY = " & FilterVar(lgKeyStream(1), "''", "S")
     '   End if 
    'End if 

    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
	Call SubHandleError("SU",lgObjConn,lgObjRs,Err)

End Sub
'============================================================================================================
' Name : SubBizSaveCreate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveSingleCreate()

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

        lgStrSQL = "INSERT INTO HFA040T("
        lgStrSQL = lgStrSQL & " EMP_NO, "  
        lgStrSQL = lgStrSQL & " YEAR_YY, "
        lgStrSQL = lgStrSQL & " A_COMP_NM, " 
        lgStrSQL = lgStrSQL & " A_COMP_NO, "  
        lgStrSQL = lgStrSQL & " A_PAY_TOT_AMT, "  
        lgStrSQL = lgStrSQL & " A_BONUS_TOT_AMT, "    
        lgStrSQL = lgStrSQL & " A_MED_INSUR, " 
        lgStrSQL = lgStrSQL & " A_SAVE_TAX_SUB, "    
        lgStrSQL = lgStrSQL & " A_INCOME_TAX, "  
        lgStrSQL = lgStrSQL & " A_FARM_TAX, "  
        lgStrSQL = lgStrSQL & " A_RES_TAX, "    
        lgStrSQL = lgStrSQL & " A_NON_TAX1, "    
        lgStrSQL = lgStrSQL & " A_NON_TAX2, "   
        lgStrSQL = lgStrSQL & " A_NON_TAX3, "   
        lgStrSQL = lgStrSQL & " A_NON_TAX4, "  
        lgStrSQL = lgStrSQL & " A_NON_TAX5, "    
        lgStrSQL = lgStrSQL & " A_INDIV_ANU_AMT, " 
        lgStrSQL = lgStrSQL & " A_INDIV_ANU2_AMT, " 
        lgStrSQL = lgStrSQL & " A_AFTER_BONUS_AMT, "    
        lgStrSQL = lgStrSQL & " A_NATIONAL_PENSION_AMT, "  
        lgStrSQL = lgStrSQL & " SEQ, " 
        lgStrSQL = lgStrSQL & " ISRT_EMP_NO, "    
        lgStrSQL = lgStrSQL & " ISRT_DT, "  
        lgStrSQL = lgStrSQL & " UPDT_EMP_NO, "  
        lgStrSQL = lgStrSQL & " UPDT_DT) "    
        lgStrSQL = lgStrSQL & " VALUES ("
    If Cint(lgKeyStream(2)) = 1 then
        lgStrSQL = lgStrSQL & FilterVar(lgKeyStream(0), "''", "S")  & ", " 
        lgStrSQL = lgStrSQL & FilterVar(lgKeyStream(1), "''", "S")  & ", " 
        lgStrSQL = lgStrSQL & FilterVar(Request("txtA_comp_nm"), "''", "S")  & ", " 
        lgStrSQL = lgStrSQL & FilterVar(Request("txtA_comp_no"), "''", "S")  & ", " 
        lgStrSQL = lgStrSQL & UNIConvNum(Request("txtA_pay_tot_amt"),0) & ","
        lgStrSQL = lgStrSQL & UNIConvNum(Request("txtA_bonus_tot_amt"),0) & ","
        lgStrSQL = lgStrSQL & UNIConvNum(Request("txtA_med_insur_amt"),0) & ","
        lgStrSQL = lgStrSQL & UNIConvNum(Request("txtA_save_tax_sub_amt"),0) & ","
        lgStrSQL = lgStrSQL & UNIConvNum(Request("txtA_income_tax_amt"),0) & ","
        lgStrSQL = lgStrSQL & UNIConvNum(Request("txtA_farm_tax_amt"),0) & ","
        lgStrSQL = lgStrSQL & UNIConvNum(Request("txtA_res_tax_amt"),0) & ","
        lgStrSQL = lgStrSQL & UNIConvNum(Request("txtA_non_tax1_amt"),0) & ","
        lgStrSQL = lgStrSQL & UNIConvNum(Request("txtA_non_tax2_amt"),0) & ","
        lgStrSQL = lgStrSQL & UNIConvNum(Request("txtA_non_tax3_amt"),0) & ","
        lgStrSQL = lgStrSQL & UNIConvNum(Request("txtA_non_tax4_amt"),0) & ","
        lgStrSQL = lgStrSQL & UNIConvNum(Request("txtA_non_tax5_amt"),0) & ","
        lgStrSQL = lgStrSQL & UNIConvNum(Request("txtA_indiv_anu_amt"),0) & ","
        lgStrSQL = lgStrSQL & UNIConvNum(Request("txtA_indiv_anu2_amt"),0) & ","
        lgStrSQL = lgStrSQL & UNIConvNum(Request("txtA_after_bonus_amt"),0) & ","
        lgStrSQL = lgStrSQL & UNIConvNum(Request("txtA_national_pension_amt"),0) & ","
        lgStrSQL = lgStrSQL & UNIConvNum(lgKeyStream(2),0) & ","
        lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''", "S")  & ","
        lgStrSQL = lgStrSQL & FilterVar(lgGetSvrDateTime,NULL,"S") & ","
        lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''", "S") & ","
        lgStrSQL = lgStrSQL & FilterVar(lgGetSvrDateTime,NULL,"S") & ")"     
        
    Else
        lgStrSQL = lgStrSQL & FilterVar(lgKeyStream(0), "''", "S")  & ", " 
        lgStrSQL = lgStrSQL & FilterVar(lgKeyStream(1), "''", "S")  & ", " 
        lgStrSQL = lgStrSQL & FilterVar(Request("txtA_comp_nm2"), "''", "S")  & ", " 
        lgStrSQL = lgStrSQL & FilterVar(Request("txtA_comp_no2"), "''", "S")  & ", " 
        lgStrSQL = lgStrSQL & UNIConvNum(Request("txtA_pay_tot_amt2"),0) & ","
        lgStrSQL = lgStrSQL & UNIConvNum(Request("txtA_bonus_tot_amt2"),0) & ","
        lgStrSQL = lgStrSQL & UNIConvNum(Request("txtA_med_insur_amt2"),0) & ","
        lgStrSQL = lgStrSQL & UNIConvNum(Request("txtA_save_tax_sub_amt2"),0) & ","
        lgStrSQL = lgStrSQL & UNIConvNum(Request("txtA_income_tax_amt2"),0) & ","
        lgStrSQL = lgStrSQL & UNIConvNum(Request("txtA_farm_tax_amt2"),0) & ","
        lgStrSQL = lgStrSQL & UNIConvNum(Request("txtA_res_tax_amt2"),0) & ","
        lgStrSQL = lgStrSQL & UNIConvNum(Request("txtA_non_tax1_amt2"),0) & ","
        lgStrSQL = lgStrSQL & UNIConvNum(Request("txtA_non_tax2_amt2"),0) & ","
        lgStrSQL = lgStrSQL & UNIConvNum(Request("txtA_non_tax3_amt2"),0) & ","
        lgStrSQL = lgStrSQL & UNIConvNum(Request("txtA_non_tax4_amt2"),0) & ","
        lgStrSQL = lgStrSQL & UNIConvNum(Request("txtA_non_tax5_amt2"),0) & ","
        lgStrSQL = lgStrSQL & UNIConvNum(Request("txtA_indiv_anu_amt2"),0) & ","
        lgStrSQL = lgStrSQL & UNIConvNum(Request("txtA_indiv_anu2_amt2"),0) & ","
        lgStrSQL = lgStrSQL & UNIConvNum(Request("txtA_after_bonus_amt2"),0) & ","
        lgStrSQL = lgStrSQL & UNIConvNum(Request("txtA_national_pension_amt2"),0) & ","
        lgStrSQL = lgStrSQL & UNIConvNum(lgKeyStream(2),0) & ","
        lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''", "S")  & ","
        lgStrSQL = lgStrSQL & FilterVar(lgGetSvrDateTime,NULL,"S") & ","
        lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''", "S") & ","
        lgStrSQL = lgStrSQL & FilterVar(lgGetSvrDateTime,NULL,"S") & ")"
    End if
    
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

    If Cint(lgKeyStream(2)) = 1 then
        lgStrSQL = "UPDATE  HFA040T"
        lgStrSQL = lgStrSQL & " SET " 
        lgStrSQL = lgStrSQL & " A_COMP_NM =  " & FilterVar(Request("txtA_comp_nm"), "''", "S") & ","
        lgStrSQL = lgStrSQL & " A_COMP_NO =  " & FilterVar(Request("txtA_comp_no"), "''", "S") & ","
        lgStrSQL = lgStrSQL & " A_PAY_TOT_AMT = " & UNIConvNum(Request("txtA_pay_tot_amt"),0) & ","
        lgStrSQL = lgStrSQL & " A_BONUS_TOT_AMT = " & UNIConvNum(Request("txtA_bonus_tot_amt"),0) & ","
        lgStrSQL = lgStrSQL & " A_AFTER_BONUS_AMT = " & UNIConvNum(Request("txtA_after_bonus_amt"),0) & ","
        lgStrSQL = lgStrSQL & " A_MED_INSUR = " & UNIConvNum(Request("txtA_med_insur_amt"),0) & ","
        lgStrSQL = lgStrSQL & " A_NATIONAL_PENSION_AMT = " & UNIConvNum(Request("txtA_national_pension_amt"),0) & ","
        lgStrSQL = lgStrSQL & " A_SAVE_TAX_SUB = " & UNIConvNum(Request("txtA_save_tax_sub_amt"),0) & ","
        lgStrSQL = lgStrSQL & " A_INDIV_ANU_AMT = " & UNIConvNum(Request("txtA_indiv_anu_amt"),0) & ","
        lgStrSQL = lgStrSQL & " A_INDIV_ANU2_AMT = " & UNIConvNum(Request("txtA_indiv_anu2_amt"),0) & ","
        lgStrSQL = lgStrSQL & " A_INCOME_TAX = " & UNIConvNum(Request("txtA_income_tax_amt"),0) & ","
        lgStrSQL = lgStrSQL & " A_RES_TAX = " & UNIConvNum(Request("txtA_res_tax_amt"),0) & ","
        lgStrSQL = lgStrSQL & " A_FARM_TAX = " & UNIConvNum(Request("txtA_farm_tax_amt"),0) & ","
        lgStrSQL = lgStrSQL & " A_NON_TAX1 = " & UNIConvNum(Request("txtA_non_tax1_amt"),0) & ","
        lgStrSQL = lgStrSQL & " A_NON_TAX2 = " & UNIConvNum(Request("txtA_non_tax2_amt"),0) & ","
        lgStrSQL = lgStrSQL & " A_NON_TAX3 = " & UNIConvNum(Request("txtA_non_tax3_amt"),0) & ","
        lgStrSQL = lgStrSQL & " A_NON_TAX4 = " & UNIConvNum(Request("txtA_non_tax4_amt"),0) & ","
        lgStrSQL = lgStrSQL & " A_NON_TAX5 = " & UNIConvNum(Request("txtA_non_tax5_amt"),0) & ","
        lgStrSQL = lgStrSQL & " UPDT_EMP_NO = " & FilterVar(gUsrId, "''", "S") & ","                                   ' char(10)
        lgStrSQL = lgStrSQL & " UPDT_DT = " & FilterVar(lgGetSvrDateTime,NULL,"S")                    ' datetime
        lgStrSQL =  lgStrSQL & " WHERE EMP_NO = " & FilterVar(lgKeyStream(0), "''", "S")
        lgStrSQL =  lgStrSQL  & "  AND YEAR_YY = " & FilterVar(lgKeyStream(1), "''", "S")
        lgStrSQL =  lgStrSQL  & "  AND SEQ = 1 "
    Else
        lgStrSQL = "UPDATE  HFA040T"
        lgStrSQL = lgStrSQL & " SET " 
        lgStrSQL = lgStrSQL & " A_COMP_NM =  " & FilterVar(Request("txtA_comp_nm2"), "''", "S") & ","
        lgStrSQL = lgStrSQL & " A_COMP_NO =  " & FilterVar(Request("txtA_comp_no2"), "''", "S") & ","
        lgStrSQL = lgStrSQL & " A_PAY_TOT_AMT = " & UNIConvNum(Request("txtA_pay_tot_amt2"),0) & ","
        lgStrSQL = lgStrSQL & " A_BONUS_TOT_AMT = " & UNIConvNum(Request("txtA_bonus_tot_amt2"),0) & ","
        lgStrSQL = lgStrSQL & " A_AFTER_BONUS_AMT = " & UNIConvNum(Request("txtA_after_bonus_amt2"),0) & ","
        lgStrSQL = lgStrSQL & " A_MED_INSUR = " & UNIConvNum(Request("txtA_med_insur_amt2"),0) & ","
        lgStrSQL = lgStrSQL & " A_NATIONAL_PENSION_AMT = " & UNIConvNum(Request("txtA_national_pension_amt2"),0) & ","
        lgStrSQL = lgStrSQL & " A_SAVE_TAX_SUB = " & UNIConvNum(Request("txtA_save_tax_sub_amt2"),0) & ","
        lgStrSQL = lgStrSQL & " A_INDIV_ANU_AMT = " & UNIConvNum(Request("txtA_indiv_anu_amt2"),0) & ","
        lgStrSQL = lgStrSQL & " A_INDIV_ANU2_AMT = " & UNIConvNum(Request("txtA_indiv_anu2_amt2"),0) & ","
        lgStrSQL = lgStrSQL & " A_INCOME_TAX = " & UNIConvNum(Request("txtA_income_tax_amt2"),0) & ","
        lgStrSQL = lgStrSQL & " A_RES_TAX = " & UNIConvNum(Request("txtA_res_tax_amt2"),0) & ","
        lgStrSQL = lgStrSQL & " A_FARM_TAX = " & UNIConvNum(Request("txtA_farm_tax_amt2"),0) & ","
        lgStrSQL = lgStrSQL & " A_NON_TAX1 = " & UNIConvNum(Request("txtA_non_tax1_amt2"),0) & ","
        lgStrSQL = lgStrSQL & " A_NON_TAX2 = " & UNIConvNum(Request("txtA_non_tax2_amt2"),0) & ","
        lgStrSQL = lgStrSQL & " A_NON_TAX3 = " & UNIConvNum(Request("txtA_non_tax3_amt2"),0) & ","
        lgStrSQL = lgStrSQL & " A_NON_TAX4 = " & UNIConvNum(Request("txtA_non_tax4_amt2"),0) & ","
        lgStrSQL = lgStrSQL & " A_NON_TAX5 = " & UNIConvNum(Request("txtA_non_tax5_amt2"),0) & ","
        lgStrSQL = lgStrSQL & " UPDT_EMP_NO = " & FilterVar(gUsrId, "''", "S") & ","                                   ' char(10)
        lgStrSQL = lgStrSQL & " UPDT_DT = " & FilterVar(lgGetSvrDateTime,NULL,"S")                    ' datetime
        lgStrSQL =  lgStrSQL & " WHERE EMP_NO = " & FilterVar(lgKeyStream(0), "''", "S")
        lgStrSQL =  lgStrSQL  & "  AND YEAR_YY = " & FilterVar(lgKeyStream(1), "''", "S")
        lgStrSQL =  lgStrSQL  & "  AND SEQ = 2 "
    End if
  
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
                           lgStrSQL = " SELECT YEAR_YY, EMP_NO,  A_COMP_NM,  A_COMP_NO, " 
                           lgStrSQL = lgStrSQL & " A_PAY_TOT_AMT,   A_BONUS_TOT_AMT,  A_MED_INSUR, A_SAVE_TAX_SUB,  A_INCOME_TAX, A_FARM_TAX,  "
                           lgStrSQL = lgStrSQL & " A_RES_TAX,  A_NON_TAX1, A_NON_TAX2,  A_NON_TAX3, A_NON_TAX4, A_NON_TAX5, A_INDIV_ANU_AMT,  A_INDIV_ANU2_AMT, "
                           lgStrSQL = lgStrSQL & " A_AFTER_BONUS_AMT,    A_NATIONAL_PENSION_AMT,   ISRT_EMP_NO,   UPDT_EMP_NO, "
                           lgStrSQL = lgStrSQL & " ISRT_DT,   UPDT_DT, SEQ  "
                           lgStrSQL = lgStrSQL & " FROM HFA040T  WHERE EMP_NO =  " & pCode
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
