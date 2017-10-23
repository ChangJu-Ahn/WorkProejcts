<%@ LANGUAGE="VBScript" CODEPAGE=949 TRANSACTION=Required %>
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
	Dim lgStrPrevKey,lgStrPrevKey1
    Dim lgSvrDateTime
    Dim lgTaxFlag, lgLngMaxRow1
    Dim lgSpreadFlg
    Call LoadBasisGlobalInf()
    Call LoadInfTB19029B("I", "H", "NOCOOKIE", "MB")
    
    lgSvrDateTime = GetSvrDateTime

    Call HideStatusWnd                                                               '☜: Hide Processing message

    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '☜: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)
    lgKeyStream       = Split(Request("txtKeyStream"),gColSep)
	lgTaxFlag		  = Request("txtTaxFlag")
    lgPrevNext        = Request("txtPrevNext")                                       '☜: "P"(Prev search) "N"(Next search)
    
    lgLngMaxRow       = Request("txtMaxRows")                                        '☜: Read Operation Mode (CRUD)
    lgLngMaxRow1      = Request("txtMaxRows1")                                       '☜: Read Operation Mode (CRUD)
    lgSpreadFlg       = Request("lgSpreadFlg")
	if lgSpreadFlg = "1" then
		lgStrPrevKey = UNICInt(Trim(Request("lgStrPrevKey")),0)                '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)
	else		
		lgStrPrevKey1 = UNICInt(Trim(Request("lgStrPrevKey1")),0)                '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)
	end if
    
    Call SubCreateCommandObject(lgObjComm)
    Call SubOpenDB(lgObjConn)                                                        '☜: Make a DB Connection

    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '☜: Query
             Call SubBizQuery()
        Case CStr(UID_M0002)                                                         '☜: Save,Update
             Call SubBizSave()        
             Call SubBizSaveMulti()
'             If lgTaxFlag = "Y" Then
				Call SubBatch()
'			 End If
    End Select
    
    Call SubCloseDB(lgObjConn)                                                       '☜: Close DB Connection

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()

    Dim txtKey
    Dim strEmp_no
    Dim strname
    Dim lSum

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    strEmp_no = FilterVar(lgKeyStream(0), "''", "S")
    txtKey = strEmp_no
    txtKey = txtKey & " AND PAY_YYMM = " & FilterVar(lgKeyStream(1), "''", "S")
    txtKey = txtKey & " AND PROV_TYPE = " & FilterVar(lgKeyStream(2), "''", "S")
	txtKey = txtKey & " AND INTERNAL_CD LIKE  " & FilterVar(lgKeyStream(3) & "%", "''", "S") & "" 

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
        
    lSum = CDbl(ConvSPChars(lgObjRs("non_tax1"))) + CDbl(ConvSPChars(lgObjRs("non_tax2"))) + CDbl(ConvSPChars(lgObjRs("non_tax3"))) + CDbl(ConvSPChars(lgObjRs("non_tax4"))) + CDbl(ConvSPChars(lgObjRs("non_tax5"))) + CDbl(ConvSPChars(lgObjRs("non_tax6")))
%>

<Script Language=vbscript>
      With Parent.Frm1
        .txtDept_cd.Value = "<%=ConvSPChars(lgObjRs("dept_nm"))%>"
        .txtOcpt_type.Value = "<%=ConvSPChars(lgObjRs("ocpt_type_nm"))%>"
        .txtPay_grd1.Value = "<%=ConvSPChars(lgObjRs("pay_grd1_nm"))%>"
        .txtPay_grd2.Value = "<%=ConvSPChars(lgObjRs("pay_grd2"))%>"
        .txtPay_cd.Value = "<%=ConvSPChars(lgObjRs("pay_cd_nm"))%>"
        .txtTax_cd.Value = "<%=ConvSPChars(lgObjRs("tax_cd_nm"))%>"
        .txtExcept_type.Value = "<%=ConvSPChars(lgObjRs("Except_type_nm"))%>"
        .txtProv_dt.Text = "<%=UNIConvDateDBToCompany(lgObjRs("Prov_dt"),"")%>"
        .txtDDPay.value = "<%=UNINumClientFormat(lgObjRs("dd_day"), ggAmtOfMoney.DecPoint, 0)%>"
        .txtComDDPay.value = "<%=UNINumClientFormat(lgObjRs("com_dd_day"), ggAmtOfMoney.DecPoint, 0)%>"
        .txtSpouse.Value = "<%=ConvSPChars(lgObjRs("spouse"))%>"
        .txtSupp_cnt.Value = "<%=ConvSPChars(lgObjRs("supp_cnt"))%>"

        .txtMm_holy_crt.value = "<%=UNINumClientFormat(lgObjRs("mm_holy_crt"), ggQty.DecPoint, 0)%>"
        .txtMm_holy_use.value = "<%=UNINumClientFormat(lgObjRs("mm_holy_use"), ggQty.DecPoint, 0)%>"
        .txtMm_holy_prov.value = "<%=UNINumClientFormat(lgObjRs("mm_holy_prov"), ggQty.DecPoint, 0)%>"
        .txtMm_accum.value = "<%=UNINumClientFormat(lgObjRs("mm_accum"), ggQty.DecPoint, 0)%>"
        
        .txtNon_tax1.value = "<%=UNINumClientFormat(lgObjRs("non_tax1"), ggAmtOfMoney.DecPoint, 0)%>"
        .txtNon_tax2.value = "<%=UNINumClientFormat(lgObjRs("non_tax2"), ggAmtOfMoney.DecPoint, 0)%>"
        .txtNon_tax3.value = "<%=UNINumClientFormat(lgObjRs("non_tax3"), ggAmtOfMoney.DecPoint, 0)%>"
        .txtNon_tax4.value = "<%=UNINumClientFormat(lgObjRs("non_tax4"), ggAmtOfMoney.DecPoint, 0)%>"
        .txtNon_tax5.value = "<%=UNINumClientFormat(lgObjRs("non_tax5"), ggAmtOfMoney.DecPoint, 0)%>"
        .txtNon_tax6.value = "<%=UNINumClientFormat(lgObjRs("non_tax6"), ggAmtOfMoney.DecPoint, 0)%>"
        					
        .txtNon_tax_sum.value = "<%=UNINumClientFormat(lSum, ggAmtOfMoney.DecPoint, 0)%>"
        .txtTax_amt.value = "<%=UNINumClientFormat(lgObjRs("tax_amt"), ggAmtOfMoney.DecPoint, 0)%>"
        .txtProv_tot_amt.value = "<%=UNINumClientFormat(lgObjRs("Prov_tot_amt"), ggAmtOfMoney.DecPoint, 0)%>"
        .txtSub_tot_amt.value = "<%=UNINumClientFormat(lgObjRs("Sub_tot_amt"), ggAmtOfMoney.DecPoint, 0)%>"
        .txtReal_Prov_amt.value = "<%=UNINumClientFormat(lgObjRs("Real_prov_amt"), ggAmtOfMoney.DecPoint, 0)%>"
        .txtEtc_non_tax.value = "<%=UNINumClientFormat(lgObjRs("etc_nontax_amt"), ggAmtOfMoney.DecPoint, 0)%>" '2007 다자녀추가공제 
        .txtEmpNo.value = "<%=ConvSPChars(lgObjRs("emp_no"))%>"
        .txtEmpNm.value = "<%=ConvSPChars(lgObjRs("name"))%>"
      End With          
</Script>       

<%     
        strEmp_no = FilterVar(ConvSPChars(lgObjRs("emp_no")), "''", "S")

		Call SubCloseRs(lgObjRs)
		if lgSpreadFlg = "1" then
			If lgErrorStatus <> "YES" Then
				txtKey = strEmp_no
				txtKey = txtKey & " AND a.PAY_YYMM = " & FilterVar(lgKeyStream(1), "''", "S")
				txtKey = txtKey & " AND a.PROV_TYPE = " & FilterVar(lgKeyStream(2), "''", "S")

				Call SubBizQueryMulti1(txtKey)
			End If
		else
		    If lgErrorStatus <> "YES" Then
				txtKey = strEmp_no
				txtKey = txtKey & " AND a.SUB_YYMM = " & FilterVar(lgKeyStream(1), "''", "S")
				txtKey = txtKey & " AND a.SUB_TYPE = " & FilterVar(lgKeyStream(2), "''", "S")
				Call SubBizQueryMulti(txtKey)
			End If
		end if
    End If
End Sub	

'============================================================================================================
' Name : SubBizSave
' Desc : Date data 
'============================================================================================================
Sub SubBizSave()

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
    
    lgIntFlgMode = CInt(Request("txtFlgMode"))                                       '☜: Read Operayion Mode (CREATE, UPDATE)
    Call SubBizSaveSingleUpdate()
End Sub	
	    

'============================================================================================================
' Name : SubBizQueryMulti1
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti1(pKey1)
    Dim iDx

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    Call SubMakeSQLStatements("MT",pKey1,"X",C_EQ)                                   '☆: Make sql statements
    
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        lgStrPrevKey = ""    
		Exit Sub
    Else
        Call SubSkipRs(lgObjRs,C_SHEETMAXROWS_D * lgStrPrevKey)
		lgstrData1 = ""
        
		iDx = 1
        
		Do While Not lgObjRs.EOF
             			                    
			lgstrData1 = lgstrData1 & Chr(11) & ConvSPChars(lgObjRs("allow_cd"))
			lgstrData1 = lgstrData1 & Chr(11) & ""            
			lgstrData1 = lgstrData1 & Chr(11) & ConvSPChars(lgObjRs("allow_nm"))
			lgstrData1 = lgstrData1 & Chr(11) & ConvSPChars(lgObjRs("tax_type"))
			lgstrData1 = lgstrData1 & Chr(11) & UNINumClientFormat(lgObjRs("allow"), ggAmtOfMoney.DecPoint, 0)
			lgstrData1 = lgstrData1 & Chr(11) & lgLngMaxRow + iDx
			lgstrData1 = lgstrData1 & Chr(11) & Chr(12)

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
' Name : SubBizQueryMulti
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti(pKey1)
    Dim iDx

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    Call SubMakeSQLStatements("MR",pKey1,"X",C_EQ)                                   '☆: Make sql statements
    
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        lgStrPrevKey1 = ""    
        Exit Sub
    Else
        Call SubSkipRs(lgObjRs,C_SHEETMAXROWS_D * lgStrPrevKey1)    
        lgstrData = ""
        iDx = 1
        
        Do While Not lgObjRs.EOF
                        
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("sub_cd")  )
            lgstrData = lgstrData & Chr(11) & ""          
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("allow_nm"))
            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("sub_amt"), ggAmtOfMoney.DecPoint, 0)
            lgstrData = lgstrData & Chr(11) & lgLngMaxRow + iDx
            lgstrData = lgstrData & Chr(11) & Chr(12)

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

    Call SubHandleError("MR",lgObjRs,Err)
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
    
    arrRowVal = Split(Request("txtSpread1"), gRowSep)                                 '☜: Split Row    data
	
    For iDx = 1 To lgLngMaxRow1
        arrColVal = Split(arrRowVal(iDx-1), gColSep)                                 '☜: Split Column data
       
        Select Case arrColVal(0)
            Case "C"
                    Call SubBizSaveMultiCreate1(arrColVal)                            '☜: Create
            Case "U"
                    Call SubBizSaveMultiUpdate1(arrColVal)                            '☜: Update
            Case "D"
                    Call SubBizSaveMultiDelete1(arrColVal)                            '☜: Delete
        End Select
        
        If lgErrorStatus    = "YES" Then
           lgErrorPos = lgErrorPos & arrColVal(1) & gColSep
           Exit For
        End If
        
    Next


End Sub    

'============================================================================================================
' Name : SubBizSaveSingleCreate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveSingleCreate()
    Dim txtGlNo

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
    
    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
	Call SubHandleError("MC",lgObjConn,lgObjRs,Err)

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
    lgStrSQL = lgStrSQL & " non_tax1 = " & UNIConvNum(Request("txtNon_tax1"),0) & ","
    lgStrSQL = lgStrSQL & " non_tax2 = " & UNIConvNum(Request("txtNon_tax2"),0) & ","
    lgStrSQL = lgStrSQL & " non_tax3 = " & UNIConvNum(Request("txtNon_tax3"),0) & ","
    lgStrSQL = lgStrSQL & " non_tax4 = " & UNIConvNum(Request("txtNon_tax4"),0) & ","
    lgStrSQL = lgStrSQL & " non_tax6 = " & UNIConvNum(Request("txtNon_tax6"),0) & ","    
    lgStrSQL = lgStrSQL & " tax_amt = " & UNIConvNum(Request("txtTax_amt"),0) & ","
    lgStrSQL = lgStrSQL & " real_prov_amt = " & UNIConvNum(Request("txtReal_prov_amt"),0) & ","
    lgStrSQL = lgStrSQL & " prov_tot_amt = " & UNIConvNum(Request("txtProv_tot_amt"),0) & ","
    lgStrSQL = lgStrSQL & " pay_tot_amt = " & UNIConvNum(Request("txtProv_tot_amt"),0) & ","
    lgStrSQL = lgStrSQL & " sub_tot_amt = " & UNIConvNum(Request("txtsub_tot_amt"),0)
    lgStrSQL = lgStrSQL & " WHERE pay_yymm = " & FilterVar(lgKeyStream(1), "''", "S")
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

    lgStrSQL = "INSERT INTO hdf040t("
    lgStrSQL = lgStrSQL & " pay_yymm," 
    lgStrSQL = lgStrSQL & " prov_type," 
    lgStrSQL = lgStrSQL & " emp_no," 
    lgStrSQL = lgStrSQL & " allow_cd,"
    lgStrSQL = lgStrSQL & " allow,"  
    lgStrSQL = lgStrSQL & " ISRT_EMP_NO," 
    lgStrSQL = lgStrSQL & " ISRT_DT     ," 
    lgStrSQL = lgStrSQL & " UPDT_EMP_NO ," 
    lgStrSQL = lgStrSQL & " UPDT_DT      )" 
    lgStrSQL = lgStrSQL & " VALUES(" 
    lgStrSQL = lgStrSQL & FilterVar(lgKeyStream(1), "''", "S") & "," 
    lgStrSQL = lgStrSQL & FilterVar(lgKeyStream(2), "''", "S") & "," 
    lgStrSQL = lgStrSQL & FilterVar(lgKeyStream(0), "''", "S") & "," 
    lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(2)), "''", "S") & "," 
    lgStrSQL = lgStrSQL & UNIConvNum(arrColVal(3),0) & "," 
    lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''", "S")                        & "," 
    lgStrSQL = lgStrSQL & FilterVar(lgSvrDateTime, "''", "S") & "," 
    lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''", "S")                        & "," 
    lgStrSQL = lgStrSQL & FilterVar(lgSvrDateTime, "''", "S")
    lgStrSQL = lgStrSQL & ")"

    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
	Call SubHandleError("MC",lgObjConn,lgObjRs,Err)

End Sub

'============================================================================================================
' Name : SubBizSaveCreate1
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveMultiCreate1(arrColVal)

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    lgStrSQL = "INSERT INTO hdf060t("
    lgStrSQL = lgStrSQL & " sub_yymm," 
    lgStrSQL = lgStrSQL & " sub_type," 
    lgStrSQL = lgStrSQL & " emp_no," 
    lgStrSQL = lgStrSQL & " sub_cd,"
    lgStrSQL = lgStrSQL & " sub_amt,"
    lgStrSQL = lgStrSQL & " calcu_type,"
    lgStrSQL = lgStrSQL & " ISRT_EMP_NO," 
    lgStrSQL = lgStrSQL & " ISRT_DT     ," 
    lgStrSQL = lgStrSQL & " UPDT_EMP_NO ," 
    lgStrSQL = lgStrSQL & " UPDT_DT      )" 
    lgStrSQL = lgStrSQL & " VALUES (" 
    lgStrSQL = lgStrSQL & FilterVar(lgKeyStream(1), "''", "S") & "," 
    lgStrSQL = lgStrSQL & FilterVar(lgKeyStream(2), "''", "S") & "," 
    lgStrSQL = lgStrSQL & FilterVar(lgKeyStream(0), "''", "S") & "," 
    lgStrSQL = lgStrSQL & FilterVar(UCase(arrColVal(2)), "''", "S") & "," 
    lgStrSQL = lgStrSQL & UNIConvNum(arrColVal(3),0) & ","
    lgStrSQL = lgStrSQL & "" & FilterVar("Y", "''", "S") & " ,"
    lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''", "S")            & "," 
    lgStrSQL = lgStrSQL & FilterVar(lgSvrDateTime, "''", "S")   & "," 
    lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''", "S")            & "," 
    lgStrSQL = lgStrSQL & FilterVar(lgSvrDateTime, "''", "S") 
    lgStrSQL = lgStrSQL & ")"

    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
	Call SubHandleError("MC",lgObjConn,lgObjRs,Err)

	lgStrSQL = ""
	If Trim(UCase(arrColVal(2))) = "S01" or Trim(UCase(arrColVal(2))) = "S02" Then	'의료보험, 국민연금 
		lgStrSQL = "INSERT INTO hdb020t("
		lgStrSQL = lgStrSQL & " pay_yymm,"
		lgStrSQL = lgStrSQL & " insur_type," 
		lgStrSQL = lgStrSQL & " emp_no,"
		lgStrSQL = lgStrSQL & " grade,"
		lgStrSQL = lgStrSQL & " prsn_insur_amt,"
		lgStrSQL = lgStrSQL & " comp_insur_amt,"
		lgStrSQL = lgStrSQL & " ISRT_EMP_NO," 
		lgStrSQL = lgStrSQL & " ISRT_DT     ," 
		lgStrSQL = lgStrSQL & " UPDT_EMP_NO ," 
		lgStrSQL = lgStrSQL & " UPDT_DT      )" 
		lgStrSQL = lgStrSQL & " VALUES (" 
		lgStrSQL = lgStrSQL & FilterVar(lgKeyStream(1), "''", "S") & "," 
		If Trim(UCase(arrColVal(2))) = "S01" Then
			lgStrSQL = lgStrSQL & "" & FilterVar("1", "''", "S") & " ,"
		Else
			lgStrSQL = lgStrSQL & "" & FilterVar("2", "''", "S") & "," 
		End If
		lgStrSQL = lgStrSQL & FilterVar(lgKeyStream(0), "''", "S") & "," 
		lgStrSQL = lgStrSQL & "" & FilterVar("00", "''", "S") & "," 
		lgStrSQL = lgStrSQL & UNIConvNum(arrColVal(3),0) & ","
		lgStrSQL = lgStrSQL & UNIConvNum(arrColVal(3),0) & ","
		lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''", "S")                        & "," 
        lgStrSQL = lgStrSQL & FilterVar(lgSvrDateTime, "''", "S")   & "," 
		lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''", "S")                        & "," 
        lgStrSQL = lgStrSQL & FilterVar(lgSvrDateTime, "''", "S") 
		lgStrSQL = lgStrSQL & ")"

		lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
		Call SubHandleError("MC",lgObjConn,lgObjRs,Err)
	End If

	lgStrSQL = ""
	If Trim(UCase(arrColVal(2))) = "S01"  Then	'의료보험 
        ' 급상여 정산 내역에도 반영한다.
        lgStrSQL = "UPDATE  HDF070T"
        lgStrSQL = lgStrSQL & " SET  med_insur = " & UNIConvNum(arrColVal(3),0) 
        lgStrSQL = lgStrSQL & " WHERE pay_yymm = " & FilterVar(lgKeyStream(1), "''", "S")
        lgStrSQL = lgStrSQL & "   AND prov_type = " & FilterVar(lgKeyStream(2), "''", "S")
        lgStrSQL = lgStrSQL & "   AND emp_no = " & FilterVar(lgKeyStream(0), "''", "S")

		lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
		Call SubHandleError("MU",lgObjConn,lgObjRs,Err)
	End If
	
	lgStrSQL = ""
	If Trim(UCase(arrColVal(2))) = "S02" Then	'국민연금 
        ' 급상여 정산 내역에도 반영한다.
        lgStrSQL = "UPDATE  HDF070T"
        lgStrSQL = lgStrSQL & " SET  anut = " & UNIConvNum(arrColVal(3),0) 
        lgStrSQL = lgStrSQL & " WHERE pay_yymm = " & FilterVar(lgKeyStream(1), "''", "S")
        lgStrSQL = lgStrSQL & "   AND prov_type = " & FilterVar(lgKeyStream(2), "''", "S")
        lgStrSQL = lgStrSQL & "   AND emp_no = " & FilterVar(lgKeyStream(0), "''", "S")

		lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
		Call SubHandleError("MU",lgObjConn,lgObjRs,Err)
	End If
	
	lgStrSQL = ""
	If Trim(UCase(arrColVal(2))) = "S03" Then	'고용보험 
        ' 급상여 정산 내역에도 반영한다.
        lgStrSQL = "UPDATE  HDF070T"
        lgStrSQL = lgStrSQL & " SET  emp_insur = " & UNIConvNum(arrColVal(3),0) 
        lgStrSQL = lgStrSQL & " WHERE pay_yymm = " & FilterVar(lgKeyStream(1), "''", "S")
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
        lgStrSQL = lgStrSQL & " WHERE pay_yymm = " & FilterVar(lgKeyStream(1), "''", "S")
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
        lgStrSQL = lgStrSQL & " WHERE pay_yymm = " & FilterVar(lgKeyStream(1), "''", "S")
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

    lgStrSQL = "UPDATE  hdf040t"
    lgStrSQL = lgStrSQL & " SET allow = " & UNIConvNum(arrColVal(3),0)
    lgStrSQL = lgStrSQL & " WHERE pay_yymm = " & FilterVar(lgKeyStream(1), "''", "S")
    lgStrSQL = lgStrSQL & "   AND prov_type = " & FilterVar(lgKeyStream(2), "''", "S")
    lgStrSQL = lgStrSQL & "   AND emp_no = " & FilterVar(lgKeyStream(0), "''", "S")
    lgStrSQL = lgStrSQL & "   AND allow_cd = " & FilterVar(UCase(arrColVal(2)), "''", "S")

    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
	Call SubHandleError("MU",lgObjConn,lgObjRs,Err)
End Sub

'============================================================================================================
' Name : SubBizSaveMultiUpdate1
' Desc : Update Data from Db
'============================================================================================================
Sub SubBizSaveMultiUpdate1(arrColVal)

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    lgStrSQL = "UPDATE  hdf060t"
    lgStrSQL = lgStrSQL & " SET sub_amt = " & UNIConvNum(arrColVal(3),0)
    lgStrSQL = lgStrSQL & " WHERE sub_yymm = " & FilterVar(lgKeyStream(1), "''", "S")
    lgStrSQL = lgStrSQL & "   AND sub_type = " & FilterVar(lgKeyStream(2), "''", "S")
    lgStrSQL = lgStrSQL & "   AND emp_no = " & FilterVar(lgKeyStream(0), "''", "S")
    lgStrSQL = lgStrSQL & "   AND sub_cd = " & FilterVar(UCase(arrColVal(2)), "''", "S")

    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
	Call SubHandleError("MU",lgObjConn,lgObjRs,Err)
 
	lgStrSQL = ""	
	If Trim(UCase(arrColVal(2))) = "S01" or Trim(UCase(arrColVal(2))) = "S02" Then
		lgStrSQL = "UPDATE  hdb020t"
		lgStrSQL = lgStrSQL & " SET prsn_insur_amt = " & UNIConvNum(arrColVal(3),0) & ","
		lgStrSQL = lgStrSQL & " comp_insur_amt = " & UNIConvNum(arrColVal(3),0)
		lgStrSQL = lgStrSQL & " WHERE pay_yymm = " & FilterVar(lgKeyStream(1), "''", "S")
		lgStrSQL = lgStrSQL & "   AND emp_no = " & FilterVar(lgKeyStream(0), "''", "S")
		If Trim(UCase(arrColVal(2))) = "S01" Then
			lgStrSQL = lgStrSQL & "   AND insur_type = " & FilterVar("1", "''", "S") & " "
		Else
			lgStrSQL = lgStrSQL & "   AND insur_type = " & FilterVar("2", "''", "S") & ""
		End If
	
		lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
		Call SubHandleError("MU",lgObjConn,lgObjRs,Err)
	End If
	
	lgStrSQL = ""
	If Trim(UCase(arrColVal(2))) = "S01"  Then	'의료보험 
        ' 급상여 정산 내역에도 반영한다.
        lgStrSQL = "UPDATE  HDF070T"
        lgStrSQL = lgStrSQL & " SET  med_insur = " & UNIConvNum(arrColVal(3),0) 
        lgStrSQL = lgStrSQL & " WHERE pay_yymm = " & FilterVar(lgKeyStream(1), "''", "S")
        lgStrSQL = lgStrSQL & "   AND prov_type = " & FilterVar(lgKeyStream(2), "''", "S")
        lgStrSQL = lgStrSQL & "   AND emp_no = " & FilterVar(lgKeyStream(0), "''", "S")

		lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
		Call SubHandleError("MU",lgObjConn,lgObjRs,Err)
	End If
	
	lgStrSQL = ""
	If Trim(UCase(arrColVal(2))) = "S02" Then	'국민연금 
        ' 급상여 정산 내역에도 반영한다.
        lgStrSQL = "UPDATE  HDF070T"
        lgStrSQL = lgStrSQL & " SET  anut = " & UNIConvNum(arrColVal(3),0) 
        lgStrSQL = lgStrSQL & " WHERE pay_yymm = " & FilterVar(lgKeyStream(1), "''", "S")
        lgStrSQL = lgStrSQL & "   AND prov_type = " & FilterVar(lgKeyStream(2), "''", "S")
        lgStrSQL = lgStrSQL & "   AND emp_no = " & FilterVar(lgKeyStream(0), "''", "S")

		lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
		Call SubHandleError("MU",lgObjConn,lgObjRs,Err)
	End If
	
	lgStrSQL = ""
	If Trim(UCase(arrColVal(2))) = "S03" Then	'고용보험 
        ' 급상여 정산 내역에도 반영한다.
        lgStrSQL = "UPDATE  HDF070T"
        lgStrSQL = lgStrSQL & " SET  emp_insur = " & UNIConvNum(arrColVal(3),0) 
        lgStrSQL = lgStrSQL & " WHERE pay_yymm = " & FilterVar(lgKeyStream(1), "''", "S")
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
        lgStrSQL = lgStrSQL & " WHERE pay_yymm = " & FilterVar(lgKeyStream(1), "''", "S")
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
        lgStrSQL = lgStrSQL & " WHERE pay_yymm = " & FilterVar(lgKeyStream(1), "''", "S")
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

    lgStrSQL = "DELETE  hdf040t"
    lgStrSQL = lgStrSQL & " WHERE pay_yymm = " & FilterVar(lgKeyStream(1), "''", "S")
    lgStrSQL = lgStrSQL & "   AND prov_type = " & FilterVar(lgKeyStream(2), "''", "S")
    lgStrSQL = lgStrSQL & "   AND emp_no = " & FilterVar(lgKeyStream(0), "''", "S")
    lgStrSQL = lgStrSQL & "   AND allow_cd = " & FilterVar(UCase(arrColVal(2)), "''", "S")

    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
	Call SubHandleError("MD",lgObjConn,lgObjRs,Err)
	
End Sub

'============================================================================================================
' Name : SubBizSaveMultiDelete1
' Desc : Delete Data from Db
'============================================================================================================
Sub SubBizSaveMultiDelete1(arrColVal)

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    lgStrSQL = "DELETE  hdf060t"
    lgStrSQL = lgStrSQL & " WHERE sub_yymm = " & FilterVar(lgKeyStream(1), "''", "S")
    lgStrSQL = lgStrSQL & "   AND sub_type = " & FilterVar(lgKeyStream(2), "''", "S")
    lgStrSQL = lgStrSQL & "   AND emp_no = " & FilterVar(lgKeyStream(0), "''", "S")
    lgStrSQL = lgStrSQL & "   AND sub_cd = " & FilterVar(UCase(arrColVal(2)), "''", "S")

    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
	Call SubHandleError("MD",lgObjConn,lgObjRs,Err)

	lgStrSQL = ""	
	If Trim(UCase(arrColVal(2))) = "S01" or Trim(UCase(arrColVal(2))) = "S02" Then
		lgStrSQL = "Delete  hdb020t"
		lgStrSQL = lgStrSQL & " WHERE pay_yymm = " & FilterVar(lgKeyStream(1), "''", "S")
		lgStrSQL = lgStrSQL & "   AND emp_no = " & FilterVar(lgKeyStream(0), "''", "S")
		If Trim(UCase(arrColVal(2))) = "S01" Then
			lgStrSQL = lgStrSQL & "   AND insur_type = " & FilterVar("1", "''", "S") & " "
		Else
			lgStrSQL = lgStrSQL & "   AND insur_type = " & FilterVar("2", "''", "S") & ""
		End If

		lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
		Call SubHandleError("MU",lgObjConn,lgObjRs,Err)
	End If
	
	lgStrSQL = ""
	If Trim(UCase(arrColVal(2))) = "S01"  Then	'의료보험 
        ' 급상여 정산 내역에도 반영한다.
        lgStrSQL = "UPDATE  HDF070T"
        lgStrSQL = lgStrSQL & " SET  med_insur = 0"
        lgStrSQL = lgStrSQL & " WHERE pay_yymm = " & FilterVar(lgKeyStream(1), "''", "S")
        lgStrSQL = lgStrSQL & "   AND prov_type = " & FilterVar(lgKeyStream(2), "''", "S")
        lgStrSQL = lgStrSQL & "   AND emp_no = " & FilterVar(lgKeyStream(0), "''", "S")

		lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
		Call SubHandleError("MU",lgObjConn,lgObjRs,Err)
	End If
	
	lgStrSQL = ""
	If Trim(UCase(arrColVal(2))) = "S02" Then	'국민연금 
        ' 급상여 정산 내역에도 반영한다.
        lgStrSQL = "UPDATE  HDF070T"
        lgStrSQL = lgStrSQL & " SET  anut = 0"
        lgStrSQL = lgStrSQL & " WHERE pay_yymm = " & FilterVar(lgKeyStream(1), "''", "S")
        lgStrSQL = lgStrSQL & "   AND prov_type = " & FilterVar(lgKeyStream(2), "''", "S")
        lgStrSQL = lgStrSQL & "   AND emp_no = " & FilterVar(lgKeyStream(0), "''", "S")

		lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
		Call SubHandleError("MU",lgObjConn,lgObjRs,Err)
	End If
	
	lgStrSQL = ""
	If Trim(UCase(arrColVal(2))) = "S03" Then	'고용보험 
        ' 급상여 정산 내역에도 반영한다.
        lgStrSQL = "UPDATE  HDF070T"
        lgStrSQL = lgStrSQL & " SET  emp_insur = 0"
        lgStrSQL = lgStrSQL & " WHERE pay_yymm = " & FilterVar(lgKeyStream(1), "''", "S")
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
        lgStrSQL = lgStrSQL & " WHERE pay_yymm = " & FilterVar(lgKeyStream(1), "''", "S")
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
        lgStrSQL = lgStrSQL & " WHERE pay_yymm = " & FilterVar(lgKeyStream(1), "''", "S")
        lgStrSQL = lgStrSQL & "   AND prov_type = " & FilterVar(lgKeyStream(2), "''", "S")
        lgStrSQL = lgStrSQL & "   AND emp_no = " & FilterVar(lgKeyStream(0), "''", "S")

		lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
		Call SubHandleError("MU",lgObjConn,lgObjRs,Err)
	End If
End Sub

Sub SubBatch()		

    On Error Resume Next                                                             '☜: Protect system from crashing

	Dim strEmpNo, strYymm, strProvType, strBasDt, strProvDt, strPaycd
	Dim strMM, strDD, strDt
	Dim strYear, strMonth, strDay
	Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
				
	strEmpNo = lgKeyStream(0)
	strYymm  = lgKeyStream(1)
	strProvType  = lgKeyStream(2)

	Call CommonQueryRs(" CONVERT(VARCHAR(8), a.prov_dt, 112), b.pay_cd "," HDF070T a, HDF020T b"," a.emp_no = " & FilterVar(strEmpNo, "''", "S") & " and a.pay_yymm = " & _
	FilterVar(strYymm, "''", "S") & " and a.prov_type = " & FilterVar(strProvType, "''", "S") & " and a.emp_no=b.emp_no",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	strProvDt = Replace(lgF0, Chr(11), "")
	strPayCd = Replace(lgF1, Chr(11), "")

	Call CommonQueryRs(" pay_bas_mm, pay_bas_dd "," HDA190T "," pay_cd =  " & FilterVar(strPayCd , "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	strMM = Replace(lgF0, Chr(11), "")
	strDD = Replace(lgF1, Chr(11), "")

	strDt = UniConvYYYYMMDDToDate(gDateFormat, Mid(strYymm,1,4), Mid(strYymm,5,2), "01")
	strDt = UniDateAdd("M", CInt(strMM), strDt, gDateFormat)
	If strDD = "00" Then
	    strDt = UniGetLastDay(strDt, gDateFormat)
        Call ExtractDateFrom(strDt, gDateFormat, gComDateType, strYear, strMonth, strDay)
		strBasDt = strYear & Right("0" & strMonth, 2) & Right("0" & strDay, 2)
	Else
        Call ExtractDateFrom(strDt, gDateFormat, gComDateType, strYear, strMonth, strDay)
		strBasDt = strYear & Right("0" & strMonth, 2) & strDD
        
	End If				

    If lgTaxFlag = "Y" Then	
		With lgObjComm
			.CommandText = "usp_hdf210b1"
			.CommandType = adCmdStoredProc
	
			lgObjComm.Parameters.Append lgObjComm.CreateParameter("@usr_id"     ,adVarXChar,adParamInput, 13 , gUsrId)
			lgObjComm.Parameters.Append lgObjComm.CreateParameter("@pay_yymm"   ,adVarXChar,adParamInput, 6  , strYymm)
			lgObjComm.Parameters.Append lgObjComm.CreateParameter("@prov_type"  ,adVarXChar,adParamInput, 1  , strProvType)
			lgObjComm.Parameters.Append lgObjComm.CreateParameter("@bas_dt"     ,adVarXChar,adParamInput, 8  , strBasDt)
			lgObjComm.Parameters.Append lgObjComm.CreateParameter("@prov_dt_s"  ,adVarXChar,adParamInput, 8  , strProvDt)
			lgObjComm.Parameters.Append lgObjComm.CreateParameter("@pay_cd"     ,adVarXChar,adParamInput, 1  , strPayCd)
			lgObjComm.Parameters.Append lgObjComm.CreateParameter("@para_emp_no",adVarXChar,adParamInput, 13 , strEmpNo)
			lgObjComm.Parameters.Append lgObjComm.CreateParameter("@update_flag" ,adVarXChar,adParamInput, 1  , "Y")				
			lgObjComm.Parameters.Append lgObjComm.CreateParameter("@msg_cd"     ,adVarXChar,adParamoutput, 6)
		    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@msg_text"   ,adVarXChar,adParamOutput,60)
		   
		    lgObjComm.Execute ,, adExecuteNoRecords
		    
		End With
	Else
		With lgObjComm
			.CommandText = "usp_h_calc_nontax"
			.CommandType = adCmdStoredProc
	
			lgObjComm.Parameters.Append lgObjComm.CreateParameter("@usr_id"     ,adVarXChar,adParamInput, 13 , gUsrId)
			lgObjComm.Parameters.Append lgObjComm.CreateParameter("@pay_yymm"   ,adVarXChar,adParamInput, 6  , strYymm)
			lgObjComm.Parameters.Append lgObjComm.CreateParameter("@prov_type"  ,adVarXChar,adParamInput, 1  , strProvType)
			lgObjComm.Parameters.Append lgObjComm.CreateParameter("@prov_dt_s"  ,adVarXChar,adParamInput, 8  , strProvDt)
			lgObjComm.Parameters.Append lgObjComm.CreateParameter("@emp_no"     ,adVarXChar,adParamInput, 13 , strEmpNo)
			lgObjComm.Parameters.Append lgObjComm.CreateParameter("@msg_cd"     ,adVarXChar,adParamoutput, 6)
		    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@msg_text"   ,adVarXChar,adParamOutput,60)
		   
		    lgObjComm.Execute ,, adExecuteNoRecords
		    
		End With
    End If
	
    If  Err.number = 0 Then
        IntRetCD = lgObjComm.Parameters("RETURN_VALUE").Value
        if  IntRetCD < 0 then
            strMsg_cd = lgObjComm.Parameters("@msg_text").Value
            call svrmsgbox(strMsg_cd, vbinformation, i_mkscript)
            IntRetCD = -1
            Exit Sub
        else
            IntRetCD = 1
        end if
    Else           
        call svrmsgbox(Err.Description, vbinformation, i_mkscript)
        Call SubHandleError(lgObjComm.ActiveConnection,lgObjRs,Err)
        IntRetCD = -1
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
                                   lgStrSQL =     "Select  dept_nm, " 
                                   lgStrSQL = lgStrSQL & " dbo.ufn_GetCodeName(" & FilterVar("H0003", "''", "S") & ", ocpt_type) ocpt_type_nm, "
                                   lgStrSQL = lgStrSQL & " dbo.ufn_GetCodeName(" & FilterVar("H0001", "''", "S") & ", pay_grd1) pay_grd1_nm, pay_grd2, "
                                   lgStrSQL = lgStrSQL & " dbo.ufn_GetCodeName(" & FilterVar("H0005", "''", "S") & ", pay_cd) pay_cd_nm, "
                                   lgStrSQL = lgStrSQL & " dbo.ufn_GetCodeName(" & FilterVar("H0006", "''", "S") & ", tax_cd) tax_cd_nm, "
                                   lgStrSQL = lgStrSQL & " dbo.ufn_GetCodeName(" & FilterVar("H0054", "''", "S") & ", Except_type) Except_type_nm, "
                                   lgStrSQL = lgStrSQL & " Prov_dt, dd_day, com_dd_day, spouse, supp_cnt, "
                                   lgStrSQL = lgStrSQL & " mm_holy_crt, mm_holy_use, mm_holy_prov, mm_accum, "
                                   lgStrSQL = lgStrSQL & " non_tax1, non_tax2, non_tax3, non_tax4, non_tax5, non_tax6,"
                                   lgStrSQL = lgStrSQL & " tax_amt, Prov_tot_amt, Sub_tot_amt, Real_prov_amt, "
                                   lgStrSQL = lgStrSQL & " emp_no, dbo.ufn_H_GetEmpName(emp_no) name,"
                                   lgStrSQL = lgStrSQL & " isnull(emp_insur,0) + isnull(anut,0) + isnull(med_insur,0) etc_nontax_amt " '2007 국민건강고용 
                                   lgStrSQL = lgStrSQL & " FROM  hdf070t "
                                   lgStrSQL = lgStrSQL & " WHERE emp_no " & pComp & pCode
                             Case "P"
                                   lgStrSQL = "Select TOP 1 " 
                                   lgStrSQL = lgStrSQL & " dept_nm, " 
                                   lgStrSQL = lgStrSQL & " dbo.ufn_GetCodeName(" & FilterVar("H0003", "''", "S") & ", ocpt_type) ocpt_type_nm, "
                                   lgStrSQL = lgStrSQL & " dbo.ufn_GetCodeName(" & FilterVar("H0001", "''", "S") & ", pay_grd1) pay_grd1_nm, pay_grd2, "
                                   lgStrSQL = lgStrSQL & " dbo.ufn_GetCodeName(" & FilterVar("H0005", "''", "S") & ", pay_cd) pay_cd_nm, "
                                   lgStrSQL = lgStrSQL & " dbo.ufn_GetCodeName(" & FilterVar("H0006", "''", "S") & ", tax_cd) tax_cd_nm, "
                                   lgStrSQL = lgStrSQL & " dbo.ufn_GetCodeName(" & FilterVar("H0054", "''", "S") & ", Except_type) Except_type_nm, "
                                   lgStrSQL = lgStrSQL & " Prov_dt, dd_day, com_dd_day, spouse, supp_cnt, "
                                   lgStrSQL = lgStrSQL & " mm_holy_crt, mm_holy_use, mm_holy_prov, mm_accum, "
                                   lgStrSQL = lgStrSQL & " non_tax1, non_tax2, non_tax3, non_tax4, non_tax5,non_tax6, "
                                   lgStrSQL = lgStrSQL & " tax_amt, Prov_tot_amt, Sub_tot_amt, Real_prov_amt, "
                                   lgStrSQL = lgStrSQL & " emp_no, dbo.ufn_H_GetEmpName(emp_no) name,"
                                   lgStrSQL = lgStrSQL & " isnull(emp_insur,0) + isnull(anut,0) + isnull(med_insur,0) etc_nontax_amt " '2007 국민건강고용 
                                   lgStrSQL = lgStrSQL & " FROM  hdf070t "
                                   lgStrSQL = lgStrSQL & " WHERE emp_no < " & pCode 
                                   lgStrSQL = lgStrSQL & " ORDER BY emp_no DESC"
                             Case "N"
                                   lgStrSQL = "Select TOP 1 " 
                                   lgStrSQL = lgStrSQL & " dept_nm, " 
                                   lgStrSQL = lgStrSQL & " dbo.ufn_GetCodeName(" & FilterVar("H0003", "''", "S") & ", ocpt_type) ocpt_type_nm, "
                                   lgStrSQL = lgStrSQL & " dbo.ufn_GetCodeName(" & FilterVar("H0001", "''", "S") & ", pay_grd1) pay_grd1_nm, pay_grd2, "
                                   lgStrSQL = lgStrSQL & " dbo.ufn_GetCodeName(" & FilterVar("H0005", "''", "S") & ", pay_cd) pay_cd_nm, "
                                   lgStrSQL = lgStrSQL & " dbo.ufn_GetCodeName(" & FilterVar("H0006", "''", "S") & ", tax_cd) tax_cd_nm, "
                                   lgStrSQL = lgStrSQL & " dbo.ufn_GetCodeName(" & FilterVar("H0054", "''", "S") & ", Except_type) Except_type_nm, "
                                   lgStrSQL = lgStrSQL & " Prov_dt, dd_day, com_dd_day, spouse, supp_cnt, "
                                   lgStrSQL = lgStrSQL & " mm_holy_crt, mm_holy_use, mm_holy_prov, mm_accum, "
                                   lgStrSQL = lgStrSQL & " non_tax1, non_tax2, non_tax3, non_tax4, non_tax5,non_tax6, "
                                   lgStrSQL = lgStrSQL & " tax_amt, Prov_tot_amt, Sub_tot_amt, Real_prov_amt, "
                                   lgStrSQL = lgStrSQL & " emp_no, dbo.ufn_H_GetEmpName(emp_no) name,"
                                   lgStrSQL = lgStrSQL & " isnull(emp_insur,0) + isnull(anut,0) + isnull(med_insur,0) etc_nontax_amt " '2007 국민건강고용 
                                   lgStrSQL = lgStrSQL & " FROM  hdf070t "
                                   lgStrSQL = lgStrSQL & " WHERE emp_no > " & pCode 	
                                   lgStrSQL = lgStrSQL & " ORDER BY emp_no ASC" 	
                        End Select
           End Select             
        Case "M"                   '                  0               1                 2              3                4                5

           Select Case Mid(pDataType,2,1)
               Case "R"
					   iSelCount = C_SHEETMAXROWS_D + C_SHEETMAXROWS_D *  lgStrPrevKey1 + 1               
					   
                       lgStrSQL = "Select top " &iSelCount & " a.sub_cd, b.allow_nm, a.sub_amt "
                       lgStrSQL = lgStrSQL & " FROM  HDF060T a, HDA010T b "
                       lgStrSQL = lgStrSQL & " WHERE b.pay_cd = " & FilterVar("*", "''", "S") & "  AND b.code_type = " & FilterVar("2", "''", "S") & " "
                       lgStrSQL = lgStrSQL & "   AND a.sub_cd = b.allow_cd AND a.emp_no " & pComp & pCode 	

               Case "T"
					   iSelCount = C_SHEETMAXROWS_D + C_SHEETMAXROWS_D *  lgStrPrevKey + 1               

                       lgStrSQL = "Select top " &iSelCount & " a.allow_cd, b.allow_nm, b.tax_type, a.allow "
                       lgStrSQL = lgStrSQL & " FROM  HDF040T a, HDA010T b "
                       lgStrSQL = lgStrSQL & " WHERE b.pay_cd = " & FilterVar("*", "''", "S") & "  AND b.code_type = " & FilterVar("1", "''", "S") & "  " 
                       lgStrSQL = lgStrSQL & "   AND a.allow_cd = b.allow_cd AND a.emp_no " & pComp & pCode 

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
				if .gSpreadFlg = "1" then
					.ggoSpread.Source     = .frm1.vspdData
					.ggoSpread.SSShowData "<%=lgstrData1%>"                               '☜ : Display data
					.lgStrPrevKey    = "<%=lgStrPrevKey%>"
					if .topleftOK then
						.DBQueryOk
					else
						.gSpreadFlg = "2"						
						.DBQuery
					end if
                else
					.ggoSpread.Source     = .frm1.vspdData1
					.ggoSpread.SSShowData "<%=lgstrData%>"          
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
