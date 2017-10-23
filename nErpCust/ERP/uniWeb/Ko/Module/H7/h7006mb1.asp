<%@ LANGUAGE=VBSCript TRANSACTION=Required%>
<% Option Explicit%>

<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<!-- #Include file="../../inc/lgSvrVariables.inc" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->

<% 
	Dim lgStrPrevKey,lgStrPrevKey1
    Dim lgSpreadFlg	
	Const C_SHEETMAXROWS_D = 100
   
    Dim lgSvrDateTime
    
    Call LoadBasisGlobalInf()
    Call LoadInfTB19029B("I", "H", "NOCOOKIE", "MB")
    
    lgSvrDateTime = GetSvrDateTime
    
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    Call HideStatusWnd                                                               '☜: Hide Processing message

    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '☜: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)
    lgKeyStream       = Split(Request("txtKeyStream"),gColSep)

    lgPrevNext        = Request("txtPrevNext")                                       '☜: "P"(Prev search) "N"(Next search)
    lgLngMaxRow       = Request("txtMaxRows")                                        '☜: Read Operation Mode (CRUD)
    lgSpreadFlg       = Request("lgSpreadFlg")

	if lgSpreadFlg = "1" then
		lgStrPrevKey = UNICInt(Trim(Request("lgStrPrevKey")),0)                '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)
	else		
		lgStrPrevKey1 = UNICInt(Trim(Request("lgStrPrevKey1")),0)                '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)
	end if

    Call SubOpenDB(lgObjConn)                                                        '☜: Make a DB Connection


    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '☜: Query
             Call SubBizQuery()
        Case CStr(UID_M0002)                                                         '☜: Save,Update
             Call SubBizSave()
             Call SubBizSaveMulti()
        Case CStr(UID_M0003)                                                         '☜: Delete
  
             Call SubBizDelete()
    End Select
    
    Call SubCloseDB(lgObjConn)                                                       '☜: Close DB Connection

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()

    Dim txtKey
    Dim iLcNo
    Dim strEmp_no
    Dim strname

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    strEmp_no = FilterVar(lgKeyStream(0), "''", "S")
    txtKey = strEmp_no
    txtKey = txtKey & " AND PAY_YYMM = " & FilterVar(Replace(lgKeyStream(1),gComDateType,""), "''", "S")
    txtKey = txtKey & " AND PROV_TYPE = " & FilterVar(lgKeyStream(2), "''", "S")
	txtKey = txtKey & " AND INTERNAL_CD LIKE " & FilterVar(lgKeyStream(3) & "%", "''", "S")

    Call SubMakeSQLStatements("SR",txtKey,"X",C_EQ)                                  '☆: Make sql statements

    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then                 'If data not exists
        strEmp_no = FilterVar(lgKeyStream(0), "''", "S")
       If lgPrevNext = "" Then
          Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)             '☜: No data is found. 
          Call SetErrorStatus()
       ElseIf lgPrevNext = "P" Then
		  if lgSpreadFlg = "1" then       
				Call DisplayMsgBox("900011", vbInformation, "", "", I_MKSCRIPT)             '☜: This is the starting data. 
		  end if
          lgPrevNext = ""
          Call SubBizQuery()
       ElseIf lgPrevNext = "N" Then
		  if lgSpreadFlg = "1" then       
				Call DisplayMsgBox("900012", vbInformation, "", "", I_MKSCRIPT)             '☜: This is the ending data.
		  end if
          lgPrevNext = ""
          Call SubBizQuery()
       End If
    Else
%>
<Script Language=vbscript>

      With Parent.Frm1
        .txtDept_cd.Value       = "<%=ConvSPChars(lgObjRs("dept_NM"))%>"
        .txtOcpt_type.Value     = "<%=ConvSPChars(lgObjRs("ocpt_type_nm"))%>"
        .txtPay_grd1.Value      = "<%=ConvSPChars(lgObjRs("pay_grd1_NM"))%>"
        .txtPay_grd2.Value      = "<%=ConvSPChars(lgObjRs("pay_grd2"))%>"
        .txtPay_cd.Value        = "<%=ConvSPChars(lgObjRs("pay_cd_nm"))%>"
        .txtTax_cd.Value        = "<%=ConvSPChars(lgObjRs("Tax_cd_Nm"))%>"
        .txtExcept_type.Value   = "<%=ConvSPChars(lgObjRs("Except_type_NM"))%>"
        .txtProv_dt.Text        = "<%=UNIConvDateDBToCompany(lgObjRs("Prov_dt"),"")%>"
        .txtDuty_mm.Value       = "<%=ConvSPChars(lgObjRs("duty_mm"))%>"
        .txtSpouse.Value        = "<%=ConvSPChars(lgObjRs("spouse"))%>"
        .txtSupp_cnt.Value      = "<%=ConvSPChars(lgObjRs("supp_cnt"))%>"

        .txtBonus_bas.text     = "<%=UNINumClientFormat(lgObjRs("Bonus_bas"), ggAmtOfMoney.DecPoint, 0)%>"
        .txtBonus_rate.text    = "<%=UNINumClientFormat(lgObjRs("Bonus_rate"), 2, 0)%>"

		.txtAdd_rate.text      = "<%=UNINumClientFormat(lgObjRs("add_rate"), 2, 0)%>"
        .txtMinus1_rate.text   = "<%=UNINumClientFormat(lgObjRs("minus1_rate"), 2, 0)%>"
        .txtMinus2_rate.text   = "<%=UNINumClientFormat(lgObjRs("minus2_rate"), 2, 0)%>"
        .txtMinus_amt.text     = "<%=UNINumClientFormat(lgObjRs("minus_amt"), ggAmtOfMoney.DecPoint, 0)%>"  '2006.05.10 차감근태 발생시 차감금액 
        .txtSplendor_rate.text = "<%=UNINumClientFormat(lgObjRs("Splendor_rate"), 2, 0)%>" '상여장려율 
        .txtSplendor_amt.text  = "<%=UNINumClientFormat(lgObjRs("Splendor_amt"), ggAmtOfMoney.DecPoint, 0)%>"
        .txtprov_rate.text     = "<%=UNINumClientFormat(lgObjRs("Prov_rate"), 2,0)%>"

        .txtsave_fund.value     = "<%=UNINumClientFormat(lgObjRs("save_fund"), ggAmtOfMoney.DecPoint, 0)%>"
        .txtincome_tax.value    = "<%=UNINumClientFormat(lgObjRs("income_tax"), ggAmtOfMoney.DecPoint, 0)%>"
        .txtres_tax.value       = "<%=UNINumClientFormat(lgObjRs("res_tax"), ggAmtOfMoney.DecPoint, 0)%>"
        .txtmed_insur.value     = "<%=UNINumClientFormat(lgObjRs("med_insur"), ggAmtOfMoney.DecPoint, 0)%>"
        .txtanut.value          = "<%=UNINumClientFormat(lgObjRs("anut"), ggAmtOfMoney.DecPoint, 0)%>"
        .txtemp_insur.value     = "<%=UNINumClientFormat(lgObjRs("emp_insur"), ggAmtOfMoney.DecPoint, 0)%>"

        .txtBonus.text         = "<%=UNINumClientFormat(lgObjRs("Bonus"), ggAmtOfMoney.DecPoint, 0)%>"
        .txtSplendor_amt.text  = "<%=UNINumClientFormat(lgObjRs("Splendor_amt"), ggAmtOfMoney.DecPoint, 0)%>"'생산장려금
		.txtProv_tot_amt.text  = "<%=UNINumClientFormat(lgObjRs("Prov_tot_amt"), ggAmtOfMoney.DecPoint, 0)%>"'상여총액
        .txtSub_tot_amt.text   = "<%=UNINumClientFormat(lgObjRs("Sub_tot_amt"), ggAmtOfMoney.DecPoint, 0)%>"'공제총액
		.txtReal_Prov_amt.text = "<%=UNINumClientFormat(lgObjRs("Real_prov_amt"), ggAmtOfMoney.DecPoint, 0)%>"'실지급액

        .txtEmp_no.value        = "<%=ConvSPChars(lgObjRs("emp_no"))%>"
        .txtName.value          = "<%=ConvSPChars(lgObjRs("emp_name"))%>"
      End With          
</Script>       
<%     
        strEmp_no = FilterVar(ConvSPChars(lgObjRs("emp_no")), "''", "S")
		if lgSpreadFlg = "1" then
			Call SubCloseRs(lgObjRs)                                                          '☜ : Release RecordSSet
			txtKey = strEmp_no
			txtKey = txtKey & " AND SUB_YYMM = " & FilterVar(Replace(lgKeyStream(1),gComDateType,""), "''", "S")
			txtKey = txtKey & " AND SUB_TYPE = " & FilterVar(lgKeyStream(2), "''", "S")
			Call SubBizQueryMulti(txtKey)
		else
			txtKey = strEmp_no
			txtKey = txtKey & " AND PAY_YYMM = " & FilterVar(Replace(lgKeyStream(1),gComDateType,""), "''", "S")
			txtKey = txtKey & " AND PROV_TYPE = " & FilterVar(lgKeyStream(2), "''", "S")
			Call SubBizQueryMulti1(txtKey)
		end if
    End If

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
        Case  OPMD_CMODE                                                             '☜: Create
              Call SubBizSaveSingleCreate()
        Case  OPMD_UMODE                                                             '☜: Update
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

    lgStrSQL = "DELETE  B_MAJOR"
    lgStrSQL = lgStrSQL & " WHERE        "
    lgStrSQL = lgStrSQL & " MAJOR_CD   = " & FilterVar(lgKeyStream(0), "''", "S")

    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
	Call SubHandleError("SD",lgObjConn,lgObjRs,Err)

End Sub

'============================================================================================================
' Name : SubBizQueryMulti
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti(pKey1)
    Dim iDx
    Dim iLoopMax

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    Call SubMakeSQLStatements("MR",pKey1,"X",C_EQ)                                   '☆: Make sql statements
    
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        lgStrPrevKey = ""
    Else
        Call SubSkipRs(lgObjRs,C_SHEETMAXROWS_D * lgStrPrevKey)

        lgstrData = ""
        iDx = 1
        
        Do While Not lgObjRs.EOF
             
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("sub_cd"))
            lgstrData = lgstrData & Chr(11) & ""
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("sub_nm"))
           
            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("sub_amt"),   ggAmtOfMoney.DecPoint,0)
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

    Call SubHandleError("MR",lgObjRs,Err)

    Call SubCloseRs(lgObjRs)                                                          '☜: Release RecordSSet

End Sub    

'============================================================================================================
' Name : SubBizQueryMulti1
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti1(pKey1)
    Dim iDx
    Dim iLoopMax
    Dim strSQL

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    Call SubMakeSQLStatements("Mr",pKey1,"X",C_EQ)                                   '☆: Make sql statements
    
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        lgStrPrevKey1 = ""
    Else
        Call SubSkipRs(lgObjRs,C_SHEETMAXROWS_D * lgStrPrevKey1)
        lgstrData1 = ""
        
        iDx = 1
        
        Do While Not lgObjRs.EOF
             
            lgstrData1 = lgstrData1 & Chr(11) & ConvSPChars(lgObjRs("allow_nm"))
            lgstrData1 = lgstrData1 & Chr(11) & UNINumClientFormat(lgObjRs("allow"), ggAmtOfMoney.DecPoint,0)
            lgstrData1 = lgstrData1 & Chr(11) & lgLngMaxRow + iDx
            lgstrData1 = lgstrData1 & Chr(11) & Chr(12)

		    lgObjRs.MoveNext

            iDx =  iDx + 1
            If iDx > C_SHEETMAXROWS_D Then
               lgStrPrevKey1 = lgStrPrevKey1 + 1
               Exit Do
            End If   
              
        Loop 
    End If
    
    If iDx <= C_SHEETMAXROWS_D Then
       lgStrPrevKey1 = ""
    End If   
    Call SubHandleError("Mr",lgObjRs,Err)
    Call SubCloseRs(lgObjRs)                                                          '☜: Release RecordSSet

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
Sub SubBizSaveSingleCreate()
End Sub

'============================================================================================================
' Name : SubBizSaveSingleUpdate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveSingleUpdate()

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    lgStrSQL = "UPDATE  HDF070T"
    lgStrSQL = lgStrSQL & " SET " 
    lgStrSQL = lgStrSQL & " bonus_bas = " & UNIConvNum(Request("txtbonus_bas"),0) & ","
    lgStrSQL = lgStrSQL & " bonus = " & UNIConvNum(Request("txtbonus"),0) & ","
    lgStrSQL = lgStrSQL & " bonus_rate = " & UNIConvNum(Request("txtbonus_rate"),0) & ","
    lgStrSQL = lgStrSQL & " add_rate = " & UNIConvNum(Request("txtadd_rate"),0) & ","
    lgStrSQL = lgStrSQL & " minus1_rate = " & UNIConvNum(Request("txtminus1_rate"),0) & ","
    lgStrSQL = lgStrSQL & " minus2_rate = " & UNIConvNum(Request("txtminus2_rate"),0) & ","
    lgStrSQL = lgStrSQL & " minus_amt = " & UNIConvNum(Request("txtMinus_amt"),0) & ","
    lgStrSQL = lgStrSQL & " splendor_rate = " & UNIConvNum(Request("txtsplendor_rate"),0) & ","
    lgStrSQL = lgStrSQL & " splendor_amt = " & UNIConvNum(Request("txtsplendor_amt"),0) & ","
    lgStrSQL = lgStrSQL & " real_prov_amt = " & UNIConvNum(Request("txtReal_prov_amt"),0) & ","
    lgStrSQL = lgStrSQL & " bonus_tot_amt = " & UNIConvNum(Request("txtProv_tot_amt"),0) & ","
    lgStrSQL = lgStrSQL & " prov_tot_amt = " & UNIConvNum(Request("txtProv_tot_amt"),0) & ","
    lgStrSQL = lgStrSQL & " sub_tot_amt = " & UNIConvNum(Request("txtsub_tot_amt"),0) & ","
    lgStrSQL = lgStrSQL & " prov_rate = " & UNIConvNum(Request("txtProv_rate"),0) & ","
    lgStrSQL = lgStrSQL & " save_fund = " & UNIConvNum(Request("txtsave_fund"),0) & ","
    lgStrSQL = lgStrSQL & " income_tax = " & UNIConvNum(Request("txtincome_tax"),0) & ","
    lgStrSQL = lgStrSQL & " res_tax = " & UNIConvNum(Request("txtres_tax"),0) & ","
    lgStrSQL = lgStrSQL & " med_insur = " & UNIConvNum(Request("txtmed_insur"),0) & ","
    lgStrSQL = lgStrSQL & " anut = " & UNIConvNum(Request("txtanut"),0) & ","
    lgStrSQL = lgStrSQL & " emp_insur = " & UNIConvNum(Request("txtemp_insur"),0)
    lgStrSQL = lgStrSQL & " WHERE           "
    lgStrSQL = lgStrSQL & "       pay_yymm = " & FilterVar(Replace(lgKeyStream(1),gComDateType,""), "''", "S")
    lgStrSQL = lgStrSQL & "   AND prov_type = " & FilterVar(lgKeyStream(2), "''", "S")
    lgStrSQL = lgStrSQL & "   AND emp_no = " & FilterVar(lgKeyStream(0), "''", "S")

    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
	Call SubHandleError("MU",lgObjConn,lgObjRs,Err)

End Sub

'============================================================================================================
' Name : SubBizSaveCreate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveMultiCreate(arrColVal)

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    lgStrSQL = "INSERT INTO hdf060t("
    lgStrSQL = lgStrSQL & " sub_yymm," 
    lgStrSQL = lgStrSQL & " sub_type," 
    lgStrSQL = lgStrSQL & " emp_no," 
    lgStrSQL = lgStrSQL & " sub_cd,"
    lgStrSQL = lgStrSQL & " sub_amt,"  
    lgStrSQL = lgStrSQL & " ISRT_EMP_NO," 
    lgStrSQL = lgStrSQL & " ISRT_DT     ," 
    lgStrSQL = lgStrSQL & " UPDT_EMP_NO ," 
    lgStrSQL = lgStrSQL & " UPDT_DT      )" 
    lgStrSQL = lgStrSQL & " VALUES(" 
    lgStrSQL = lgStrSQL & FilterVar(Replace(lgKeyStream(1),gComDateType,""), "''", "S") & "," 
    lgStrSQL = lgStrSQL & FilterVar(lgKeyStream(2), "''", "S") & "," 
    lgStrSQL = lgStrSQL & FilterVar(lgKeyStream(0), "''", "S") & "," 
    lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(2)), "''", "S") & "," 
    lgStrSQL = lgStrSQL & UNIConvNum(arrColVal(3),0) & "," 
    lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''", "S")                        & "," 
    lgStrSQL = lgStrSQL & FilterVar(lgSvrDateTime,NULL,"S") & "," 
    lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''", "S")                        & "," 
    lgStrSQL = lgStrSQL & FilterVar(lgSvrDateTime,NULL,"S")
    lgStrSQL = lgStrSQL & ")"
    
    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
	Call SubHandleError("MC",lgObjConn,lgObjRs,Err)
	
	lgStrSQL = ""
	If Trim(UCase(arrColVal(2))) = "S03" Then	'고용보험 
        ' 급상여 정산 내역에도 반영한다.
        lgStrSQL = "UPDATE  HDF070T"
        lgStrSQL = lgStrSQL & " SET  emp_insur = " & UNIConvNum(arrColVal(3),0) 
        lgStrSQL = lgStrSQL & " WHERE "
        lgStrSQL = lgStrSQL & "   pay_yymm = " & FilterVar(lgKeyStream(1), "''", "S")
        lgStrSQL = lgStrSQL & "   AND prov_type = " & FilterVar(lgKeyStream(2), "''", "S")
        lgStrSQL = lgStrSQL & "   AND emp_no = " & FilterVar(lgKeyStream(0), "''", "S")

		lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
		Call SubHandleError("MU",lgObjConn,lgObjRs,Err)
	End If
	
	lgStrSQL = ""
	If Trim(UCase(arrColVal(2))) = "S98" Then	'소득세 
        ' 급상여 정산 내역에도 반영한다.
        lgStrSQL = "UPDATE  HDF070T"
        lgStrSQL = lgStrSQL & " SET  income_tax = " & UNIConvNum(arrColVal(3),0) 
        lgStrSQL = lgStrSQL & " WHERE "
        lgStrSQL = lgStrSQL & "   pay_yymm = " & FilterVar(lgKeyStream(1), "''", "S")
        lgStrSQL = lgStrSQL & "   AND prov_type = " & FilterVar(lgKeyStream(2), "''", "S")
        lgStrSQL = lgStrSQL & "   AND emp_no = " & FilterVar(lgKeyStream(0), "''", "S")

		lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
		Call SubHandleError("MU",lgObjConn,lgObjRs,Err)
	End If
	
	lgStrSQL = ""
	If Trim(UCase(arrColVal(2))) = "S99" Then	'주민세 
        ' 급상여 정산 내역에도 반영한다.
        lgStrSQL = "UPDATE  HDF070T"
        lgStrSQL = lgStrSQL & " SET  res_tax = " & UNIConvNum(arrColVal(3),0) 
        lgStrSQL = lgStrSQL & " WHERE "
        lgStrSQL = lgStrSQL & "   pay_yymm = " & FilterVar(lgKeyStream(1), "''", "S")
        lgStrSQL = lgStrSQL & "   AND prov_type = " & FilterVar(lgKeyStream(2), "''", "S")
        lgStrSQL = lgStrSQL & "   AND emp_no = " & FilterVar(lgKeyStream(0), "''", "S")

		lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
		Call SubHandleError("MU",lgObjConn,lgObjRs,Err)
	End If

End Sub
'============================================================================================================
' Name : SubBizSaveMultiUpdate
' Desc : Update Data from Db
'============================================================================================================
Sub SubBizSaveMultiUpdate(arrColVal)

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    lgStrSQL = "UPDATE  hdf060t"
    lgStrSQL = lgStrSQL & " SET " 
    lgStrSQL = lgStrSQL & "      sub_amt = " & UNIConvNum(arrColVal(3),0)
    lgStrSQL = lgStrSQL & " WHERE "
    lgStrSQL = lgStrSQL & "       sub_yymm = " & FilterVar(Replace(lgKeyStream(1),gComDateType,""), "''", "S")
    lgStrSQL = lgStrSQL & "   AND sub_type = " & FilterVar(lgKeyStream(2), "''", "S")
    lgStrSQL = lgStrSQL & "   AND emp_no = " & FilterVar(lgKeyStream(0), "''", "S")
    lgStrSQL = lgStrSQL & "   AND sub_cd = " & FilterVar(UCase(arrColVal(2)), "''", "S")

    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
	Call SubHandleError("MU",lgObjConn,lgObjRs,Err)

	lgStrSQL = ""
	If Trim(UCase(arrColVal(2))) = "S03" Then	'고용보험 
        ' 급상여 정산 내역에도 반영한다.
        lgStrSQL = "UPDATE  HDF070T"
        lgStrSQL = lgStrSQL & " SET  emp_insur = " & UNIConvNum(arrColVal(3),0) 
        lgStrSQL = lgStrSQL & " WHERE "
        lgStrSQL = lgStrSQL & "   pay_yymm = " & FilterVar(lgKeyStream(1), "''", "S")
        lgStrSQL = lgStrSQL & "   AND prov_type = " & FilterVar(lgKeyStream(2), "''", "S")
        lgStrSQL = lgStrSQL & "   AND emp_no = " & FilterVar(lgKeyStream(0), "''", "S")

		lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
		Call SubHandleError("MU",lgObjConn,lgObjRs,Err)
	End If
	
	lgStrSQL = ""
	If Trim(UCase(arrColVal(2))) = "S98" Then	'소득세 
        ' 급상여 정산 내역에도 반영한다.
        lgStrSQL = "UPDATE  HDF070T"
        lgStrSQL = lgStrSQL & " SET  income_tax = " & UNIConvNum(arrColVal(3),0) 
        lgStrSQL = lgStrSQL & " WHERE "
        lgStrSQL = lgStrSQL & "   pay_yymm = " & FilterVar(lgKeyStream(1), "''", "S")
        lgStrSQL = lgStrSQL & "   AND prov_type = " & FilterVar(lgKeyStream(2), "''", "S")
        lgStrSQL = lgStrSQL & "   AND emp_no = " & FilterVar(lgKeyStream(0), "''", "S")

		lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
		Call SubHandleError("MU",lgObjConn,lgObjRs,Err)
	End If
	
	lgStrSQL = ""
	If Trim(UCase(arrColVal(2))) = "S99" Then	'주민세 
        ' 급상여 정산 내역에도 반영한다.
        lgStrSQL = "UPDATE  HDF070T"
        lgStrSQL = lgStrSQL & " SET  res_tax = " & UNIConvNum(arrColVal(3),0) 
        lgStrSQL = lgStrSQL & " WHERE "
        lgStrSQL = lgStrSQL & "   pay_yymm = " & FilterVar(lgKeyStream(1), "''", "S")
        lgStrSQL = lgStrSQL & "   AND prov_type = " & FilterVar(lgKeyStream(2), "''", "S")
        lgStrSQL = lgStrSQL & "   AND emp_no = " & FilterVar(lgKeyStream(0), "''", "S")

		lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
		Call SubHandleError("MU",lgObjConn,lgObjRs,Err)
	End If
	
End Sub


'============================================================================================================
' Name : SubBizSaveMultiDelete
' Desc : Delete Data from Db
'============================================================================================================
Sub SubBizSaveMultiDelete(arrColVal)

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    lgStrSQL = "DELETE  hdf060t"
    lgStrSQL = lgStrSQL & " WHERE        "
    lgStrSQL = lgStrSQL & "       sub_yymm = " & FilterVar(Replace(lgKeyStream(1),gComDateType,""), "''", "S")
    lgStrSQL = lgStrSQL & "   AND sub_type = " & FilterVar(lgKeyStream(2), "''", "S")
    lgStrSQL = lgStrSQL & "   AND emp_no = " & FilterVar(lgKeyStream(0), "''", "S")
    lgStrSQL = lgStrSQL & "   AND sub_cd = " & FilterVar(UCase(arrColVal(2)), "''", "S")

    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
	Call SubHandleError("MD",lgObjConn,lgObjRs,Err)

	lgStrSQL = ""
	If Trim(UCase(arrColVal(2))) = "S03" Then	'고용보험 
        ' 급상여 정산 내역에도 반영한다.
        lgStrSQL = "UPDATE  HDF070T"
        lgStrSQL = lgStrSQL & " SET  emp_insur = 0"
        lgStrSQL = lgStrSQL & " WHERE "
        lgStrSQL = lgStrSQL & "   pay_yymm = " & FilterVar(lgKeyStream(1), "''", "S")
        lgStrSQL = lgStrSQL & "   AND prov_type = " & FilterVar(lgKeyStream(2), "''", "S")
        lgStrSQL = lgStrSQL & "   AND emp_no = " & FilterVar(lgKeyStream(0), "''", "S")

		lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
		Call SubHandleError("MU",lgObjConn,lgObjRs,Err)
	End If
	
	lgStrSQL = ""
	If Trim(UCase(arrColVal(2))) = "S98" Then	'소득세 
        ' 급상여 정산 내역에도 반영한다.
        lgStrSQL = "UPDATE  HDF070T"
        lgStrSQL = lgStrSQL & " SET  income_tax = 0"
        lgStrSQL = lgStrSQL & " WHERE "
        lgStrSQL = lgStrSQL & "   pay_yymm = " & FilterVar(lgKeyStream(1), "''", "S")
        lgStrSQL = lgStrSQL & "   AND prov_type = " & FilterVar(lgKeyStream(2), "''", "S")
        lgStrSQL = lgStrSQL & "   AND emp_no = " & FilterVar(lgKeyStream(0), "''", "S")

		lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
		Call SubHandleError("MU",lgObjConn,lgObjRs,Err)
	End If
	
	lgStrSQL = ""
	If Trim(UCase(arrColVal(2))) = "S99" Then	'주민세 
        ' 급상여 정산 내역에도 반영한다.
        lgStrSQL = "UPDATE  HDF070T"
        lgStrSQL = lgStrSQL & " SET  res_tax = 0"
        lgStrSQL = lgStrSQL & " WHERE "
        lgStrSQL = lgStrSQL & "   pay_yymm = " & FilterVar(lgKeyStream(1), "''", "S")
        lgStrSQL = lgStrSQL & "   AND prov_type = " & FilterVar(lgKeyStream(2), "''", "S")
        lgStrSQL = lgStrSQL & "   AND emp_no = " & FilterVar(lgKeyStream(0), "''", "S")

		lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
		Call SubHandleError("MU",lgObjConn,lgObjRs,Err)
	End If
	
End Sub


'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements(pDataType,pCode,pCode1,pComp)
    Dim iSelCount
    Select Case Mid(pDataType,1,1)
        Case "S"
           Select Case Mid(pDataType,2,1)
               Case "R"
                        Select Case  lgPrevNext 
                             Case ""
                                   lgStrSQL = "Select  dept_NM, dbo.ufn_GetCodeName(" & FilterVar("H0003", "''", "S") & ",ocpt_type) Ocpt_type_Nm, "
                                   lgStrSQL = lgStrSQL & " dbo.Ufn_GetCodeName(" & FilterVar("H0001", "''", "S") & ", pay_grd1) Pay_grd1_NM, pay_grd2, dbo.ufn_GetCodeName(" & FilterVar("H0005", "''", "S") & ",pay_cd) Pay_cd_NM, "
                                   lgStrSQL = lgStrSQL & " dbo.ufn_GetCodeName(" & FilterVar("H0006", "''", "S") & ", tax_cd) Tax_cd_Nm, dbo.ufn_GetCodeName(" & FilterVar("H0054", "''", "S") & ",Except_type) Except_type_NM, " 
                                   lgStrSQL = lgStrSQL & " Prov_dt, duty_mm, spouse, supp_cnt, Bonus_bas, Bonus_rate, add_rate, "
                                   lgStrSQL = lgStrSQL & " minus1_rate, minus2_rate, Splendor_rate, Splendor_amt, Prov_rate, save_fund, minus_amt," '2006.05.10 차감근태 발생시 차감금액 
                                   lgStrSQL = lgStrSQL & " income_tax, res_tax, med_insur, anut, emp_insur, Bonus, Splendor_amt, Prov_tot_amt, "
                                   lgStrSQL = lgStrSQL & " Sub_tot_amt, Real_prov_amt, emp_no, dbo.ufn_H_GetEmpName(emp_no) emp_name  "
                                   lgStrSQL = lgStrSQL & " FROM  hdf070t "
                                   lgStrSQL = lgStrSQL & " WHERE emp_no " & pComp & pCode 
                                   lgStrSQL = lgStrSQL & " Order by emp_no "
                             Case "P"
                                   lgStrSQL = "Select TOP 1  dept_NM, dbo.ufn_GetCodeName(" & FilterVar("H0003", "''", "S") & ",ocpt_type) Ocpt_type_Nm, "
                                   lgStrSQL = lgStrSQL & " dbo.Ufn_GetCodeName(" & FilterVar("H0001", "''", "S") & ", pay_grd1) Pay_grd1_NM, pay_grd2, dbo.ufn_GetCodeName(" & FilterVar("H0005", "''", "S") & ",pay_cd) Pay_cd_NM, "
                                   lgStrSQL = lgStrSQL & " dbo.ufn_GetCodeName(" & FilterVar("H0006", "''", "S") & ", tax_cd) Tax_cd_Nm, dbo.ufn_GetCodeName(" & FilterVar("H0054", "''", "S") & ",Except_type) Except_type_NM, " 
                                   lgStrSQL = lgStrSQL & " Prov_dt, duty_mm, spouse, supp_cnt, Bonus_bas, Bonus_rate, add_rate, "
                                   lgStrSQL = lgStrSQL & " minus1_rate, minus2_rate, Splendor_rate, Splendor_amt, Prov_rate, save_fund,  minus_amt," '2006.05.10 차감근태 발생시 차감금액 
                                   lgStrSQL = lgStrSQL & " income_tax, res_tax, med_insur, anut, emp_insur, Bonus, Splendor_amt, Prov_tot_amt, "
                                   lgStrSQL = lgStrSQL & " Sub_tot_amt, Real_prov_amt, emp_no, dbo.ufn_H_GetEmpName(emp_no) emp_name  "                             
                                   
                                   lgStrSQL = lgStrSQL & " FROM  hdf070t "
                                   lgStrSQL = lgStrSQL & " WHERE emp_no < " & pCode 	
                                   lgStrSQL = lgStrSQL & " Order by emp_no DESC"
                             Case "N"
								   lgStrSQL = "Select TOP 1  dept_NM, dbo.ufn_GetCodeName(" & FilterVar("H0003", "''", "S") & ",ocpt_type) Ocpt_type_Nm, "
                                   lgStrSQL = lgStrSQL & " dbo.Ufn_GetCodeName(" & FilterVar("H0001", "''", "S") & ", pay_grd1) Pay_grd1_NM, pay_grd2, dbo.ufn_GetCodeName(" & FilterVar("H0005", "''", "S") & ",pay_cd) Pay_cd_NM, "
                                   lgStrSQL = lgStrSQL & " dbo.ufn_GetCodeName(" & FilterVar("H0006", "''", "S") & ", tax_cd) Tax_cd_Nm, dbo.ufn_GetCodeName(" & FilterVar("H0054", "''", "S") & ",Except_type) Except_type_NM, " 
                                   lgStrSQL = lgStrSQL & " Prov_dt, duty_mm, spouse, supp_cnt, Bonus_bas, Bonus_rate, add_rate, "
                                   lgStrSQL = lgStrSQL & " minus1_rate, minus2_rate, Splendor_rate, Splendor_amt, Prov_rate, save_fund,  minus_amt," '2006.05.10 차감근태 발생시 차감금액 
                                   lgStrSQL = lgStrSQL & " income_tax, res_tax, med_insur, anut, emp_insur, Bonus, Splendor_amt, Prov_tot_amt, "
                                   lgStrSQL = lgStrSQL & " Sub_tot_amt, Real_prov_amt, emp_no, dbo.ufn_H_GetEmpName(emp_no) emp_name  "                             
                                                                      
                                   lgStrSQL = lgStrSQL & " FROM  hdf070t "
                                   lgStrSQL = lgStrSQL & " WHERE emp_no > " & pCode 	
                                   lgStrSQL = lgStrSQL & " Order by emp_no "
                        End Select
           End Select             
        Case "M"           '                  0               1                 2              3                4                5
           
           Select Case Mid(pDataType,2,1)
               Case "R"
					   iSelCount = C_SHEETMAXROWS_D + C_SHEETMAXROWS_D *  lgStrPrevKey + 1               
               
                       lgStrSQL = "Select  top " &iSelCount & " sub_cd, dbo.ufn_H_GetCodeName(" & FilterVar("HDA010T", "''", "S") & ",sub_cd, '') sub_NM, sub_amt  "
                       lgStrSQL = lgStrSQL & " FROM  HDF060T "
                       lgStrSQL = lgStrSQL & " WHERE emp_no " & pComp & pCode
               Case "r"
					   iSelCount = C_SHEETMAXROWS_D + C_SHEETMAXROWS_D *  lgStrPrevKey1 + 1               
                       
                       lgStrSQL = "Select  top " &iSelCount & " allow_cd, dbo.ufn_H_GetCodeName(" & FilterVar("HDA010T", "''", "S") & ",allow_cd, " & FilterVar("1", "''", "S") & " ) allow_NM, allow  "
                       lgStrSQL = lgStrSQL & " FROM  HDF041T "
                       lgStrSQL = lgStrSQL & " WHERE emp_no " & pComp & pCode 	
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
              With Parent
              
				if .lgSpreadFlg = "1" then              
					.ggoSpread.Source     = .frm1.vspdData
					.ggoSpread.SSShowData "<%=lgstrData%>"                               '☜ : Display data
					.lgStrPrevKey         = "<%=lgStrPrevKey%>"                          '☜ : Next next data tag 
					if .topleftOK then
						.DBQueryOk
					else
						.lgSpreadFlg = "2"							
						.DBQuery
					end if
				else
					.ggoSpread.Source     = .frm1.vspdData1
					.ggoSpread.SSShowData "<%=lgstrData1%>"         
					.lgStrPrevKey1    = "<%=lgStrPrevKey1%>" 
	                .DBQueryOk        
	            end if
	         End with
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
