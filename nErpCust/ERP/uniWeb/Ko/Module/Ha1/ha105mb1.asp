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
	Const C_SHEETMAXROWS_D = 100
    
    Call LoadBasisGlobalInf()
    Call LoadInfTB19029B("I", "H", "NOCOOKIE", "MB")
	
    Dim lgGetSvrDateTime

    Call HideStatusWnd                                                               '☜: Hide Processing message

    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '☜: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)
    lgKeyStream       = Split(Request("txtKeyStream"),gColSep)

    lgLngMaxRow       = Request("txtMaxRows")                                        '☜: Read Operation Mode (CRUD)
    lgCurrentSpd      = Request("lgCurrentSpd")                                      '☜: "M"(Spread #1) "S"(Spread #2)

    If lgCurrentSpd = "M" Then    
	    lgStrPrevKey = UNICInt(Trim(Request("lgStrPrevKey")),0)                '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)
    elseif lgCurrentSpd = "S" Then    
	    lgStrPrevKey1 = UNICInt(Trim(Request("lgStrPrevKey1")),0)                '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)
	end if
	    
	Dim Sflag    
    Sflag = false

    Call SubOpenDB(lgObjConn)                                                        '☜: Make a DB Connection
	lgGetSvrDateTime = GetSvrDateTime
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
    Dim iKey1
    Dim strWhere
    Dim strRetire_yyyy
    Dim strRetire_yymm    
    Dim strYear,strMonth,strDay

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    strRetire_yymm = UNIConvDateCompanyToDB(lgKeyStream(0),NULL)                         '퇴직년월 
    iKey1 = FilterVar(lgKeyStream(1),"" & FilterVar("%", "''", "S") & "","S")            '사번 
    strWhere = iKey1
    Call ExtractDateFrom(lgKeyStream(0),gDateFormat,gComDateType,strYear,strMonth,strDay)
    strRetire_yyyy = strYear
    strWhere = strWhere & " And ( HGA070T.RETIRE_YY =  " & FilterVar(strRetire_yyyy , "''", "S") & ") "
    strWhere = strWhere & " And ( Convert(varchar(6), HGA070T.RETIRE_DT, 112)) = convert(varchar(6),Convert(DateTime, " & FilterVar(strRetire_yymm, "''", "S") & "), 112) " 
    
    Call SubMakeSQLStatements("SR",strWhere,"X",C_EQ)                                  '☆: Make sql statements

    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then                 'If data not exists
       If lgPrevNext = "" Then
          Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)             '☜: No data is found. 
          Call SetErrorStatus()
       ElseIf lgPrevNext = "P" Then
          Call DisplayMsgBox("900011", vbInformation, "", "", I_MKSCRIPT)             '☜: This is the starting data. 
          lgPrevNext = ""
          Call SubBizQuery()
       ElseIf lgPrevNext = "N" Then
          Call DisplayMsgBox("900012", vbInformation, "", "", I_MKSCRIPT)             '☜: This is the ending data.
          lgPrevNext = ""
          Call SubBizQuery()
       End If
       
    Else
		Sflag = True
%>
<Script Language=vbscript>
      With Parent.Frm1
             .txtEntr_dt.Text               = "<%=UNIConvDateDBToCompany(lgObjRs("entr_dt"),Null)%>"            
             .txtRetire_amt.text           = "<%=UniNumclientFormat(lgObjRs("retire_amt"),ggAmtOfMoney.DecPoint,0)%>"            
             .txtIncome_sub_amt.text       = "<%=UniNumclientFormat(lgObjRs("income_sub"),ggAmtOfMoney.DecPoint,0)%>"            
             .txtDuty_cnt.text             = "<%=UniNumclientFormat(lgObjRs("duty_cnt"),0,0)%>"            
             .txtRetire_dt.Text            = "<%=UNIConvDateDBToCompany(lgObjRs("retire_dt"),Null)%>"
             .txtCorp_insur_amt.text       = "<%=UniNumclientFormat(lgObjRs("corp_insur"),ggAmtOfMoney.DecPoint,0)%>"            
             .txtSpecial_sub_amt.text      = "<%=UniNumclientFormat(lgObjRs("special_sub"),ggAmtOfMoney.DecPoint,0)%>"
             .txtEtc_sub_amt.text          = "<%=UniNumclientFormat(lgObjRs("etc_sub"),ggAmtOfMoney.DecPoint,0)%>"                        
             .txtCalc_tax_amt.text         = "<%=UniNumclientFormat(lgObjRs("calc_tax"),ggAmtOfMoney.DecPoint,0)%>"            
             .txtTot_duty_mm.text          = "<%=UniNumclientFormat(lgObjRs("tot_duty_mm"),0,0)%>"            
             .txtHonor_amt.text            = "<%=UniNumclientFormat(lgObjRs("honor_amt"),ggAmtOfMoney.DecPoint,0)%>"            
             .txtIncome_amt.text           = "<%=UniNumclientFormat(lgObjRs("income_amt"),ggAmtOfMoney.DecPoint,0)%>"            
             .txtAvr_calc_tax_amt.text     = "<%=UniNumclientFormat(lgObjRs("avr_calc_tax"),ggAmtOfMoney.DecPoint,0)%>"            
             .txtPay_avr_amt.text          = "<%=UniNumclientFormat(lgObjRs("pay_avr"),ggAmtOfMoney.DecPoint,0)%>"            
             .txtEtc_amt.text              = "<%=UniNumclientFormat(lgObjRs("etc_amt"),ggAmtOfMoney.DecPoint,0)%>"            
             .txtSub_short_amt.text        = "<%=UniNumclientFormat(lgObjRs("sub_short_amt"),ggAmtOfMoney.DecPoint,0)%>"            
             .txtTax_short_amt.text        = "<%=UniNumclientFormat(lgObjRs("tax_short_amt"),ggAmtOfMoney.DecPoint,0)%>"            
             .txtBonus_avr_amt.text        = "<%=UniNumclientFormat(lgObjRs("bonus_avr"),ggAmtOfMoney.DecPoint,0)%>"            
             .txtIncome_tax_amt.text       = "<%=UniNumclientFormat(lgObjRs("income_tax"),ggAmtOfMoney.DecPoint,0)%>"            
             .txtTax_std_amt.text          = "<%=UniNumclientFormat(lgObjRs("tax_std"),ggAmtOfMoney.DecPoint,0)%>"            
             .txtDeci_income_tax_amt.text  = "<%=UniNumclientFormat(lgObjRs("deci_income_tax"),ggAmtOfMoney.DecPoint,0)%>"            
             .txtYear_avr_amt.text         = "<%=UniNumclientFormat(lgObjRs("year_avr"),ggAmtOfMoney.DecPoint,0)%>"            
             .txtRes_tax_amt.text          = "<%=UniNumclientFormat(lgObjRs("res_tax"),ggAmtOfMoney.DecPoint,0)%>"            
             .txtAvr_tax_std_amt.text      = "<%=UniNumclientFormat(lgObjRs("avr_tax_std"),ggAmtOfMoney.DecPoint,0)%>"            
             .txtDeci_res_tax_amt.text     = "<%=UniNumclientFormat(lgObjRs("deci_res_tax"),ggAmtOfMoney.DecPoint,0)%>"            
             .txtAvr_wages_amt.text        = "<%=UniNumclientFormat(lgObjRs("avr_wages"),ggAmtOfMoney.DecPoint,0)%>"            
             .txtRetire_anu_amt.text       = "<%=UniNumclientFormat(lgObjRs("retire_anu_amt"),ggAmtOfMoney.DecPoint,0)%>"            
             .txtTax_rate.text             = "<%=UniNumclientFormat(lgObjRs("tax_rate"),2,0)%>"            
             .txtTot_prov_amt.text         = "<%=UniNumclientFormat(lgObjRs("tot_prov_amt"),ggAmtOfMoney.DecPoint,0)%>"            
             .txtReal_prov_amt.text        = "<%=UniNumclientFormat(lgObjRs("real_prov_amt"),ggAmtOfMoney.DecPoint,0)%>"            
      End With          
</Script>       
<%     
       Call SubCloseRs(lgObjRs)                                                          '☜ : Release RecordSSet
       Call SubBizQueryMulti()
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
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
End Sub

'============================================================================================================
' Name : SubBizQueryMulti
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()
    Dim iDx
    Dim iKey1
    Dim strWhere_m
    Dim strWhere_s
    Dim strRetire_yymm
    Dim strallow_cd_nm
    Dim strbonus_type_nm
    Dim iSum
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    strRetire_yymm = UNIConvDateCompanyToDB(lgKeyStream(0),NULL)                        '퇴직년월 
    iKey1 = FilterVar(lgKeyStream(1),"" & FilterVar("%", "''", "S") & "","S")            '사번 
    strWhere_m = iKey1                                     '급여조건 
    strWhere_m = strWhere_m & " And convert(varchar(6), hga050t.retire_dt, 112) = convert(varchar(6),Convert(DateTime, " & FilterVar(strRetire_yymm, "''", "S") & "), 112) " 
    strWhere_m = strWhere_m & " GROUP BY HGA050T.PAY_YYMM,HGA050T.ALLOW_CD WITH ROLLUP "
    strWhere_s = iKey1                                     '상여조건 
    strWhere_s = strWhere_s & " And convert(varchar(6), hga060t.retire_dt, 112) = convert(varchar(6),Convert(DateTime, " & FilterVar(strRetire_yymm, "''", "S") & "), 112) " 

    Call SubMakeSQLStatements("MR",strWhere_m,strWhere_s,C_EQ)                              '☆: Make sql statements
    
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        lgStrPrevKey = ""
        lgStrPrevKey1 = ""
    Else
		Select Case lgCurrentSpd
		   Case "M"
 				Call SubSkipRs(lgObjRs,C_SHEETMAXROWS_D * lgStrPrevKey)        
 		   Case "S" 
 				Call SubSkipRs(lgObjRs,C_SHEETMAXROWS_D * lgStrPrevKey1)        
 		end Select
        lgstrData = ""
        iDx = 1
        iSum = 0     '상여금합계를 구함 
        Dim Teststr
        Do While Not lgObjRs.EOF
            Select Case lgCurrentSpd
               Case "M"
                  
                    lgstrData = lgstrData & Chr(11) & UNIMonthClientFormat(Left(lgObjRs("pay_yymm"),4) & gAPDateSeperator & Right(lgObjRs("pay_yymm"),2) & gAPDateSeperator & "01")                    
                    lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("allow_cd"))
                    lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("allow_NM"))
                    lgstrData = lgstrData & Chr(11) & ""
                    lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("allow"), ggAmtOfMoney.DecPoint,0)
               Case "S"
                    lgstrData = lgstrData & Chr(11) & UNIMonthClientFormat(Left(lgObjRs("bonus_yymm"),4) & gAPDateSeperator & Right(lgObjRs("bonus_yymm"),2) & gAPDateSeperator & "01")
                    lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("bonus_type"))
                    lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("bonus_type_NM"))
                    lgstrData = lgstrData & Chr(11) & ""
                    lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("bonus"), ggAmtOfMoney.DecPoint,0)
                    iSum = iSum + CDbl(lgObjRs("bonus"))
            End Select      
            
            lgstrData = lgstrData & Chr(11) & lgLngMaxRow + iDx
            lgstrData = lgstrData & Chr(11) & Chr(12)

		    lgObjRs.MoveNext
            iDx =  iDx + 1
            If iDx > C_SHEETMAXROWS_D Then
				Select Case lgCurrentSpd
				   Case "M"
						lgStrPrevKey = lgStrPrevKey + 1
 				   Case "S"
						lgStrPrevKey1 = lgStrPrevKey1 + 1
 				end Select
				Exit Do
            End If                       
               
        Loop 
    End If
    If iDx <= C_SHEETMAXROWS_D Then
		Select Case lgCurrentSpd
		   Case "M"
				lgStrPrevKey = ""
 		   Case "S"
				if lgstrData <> "" then  
                    lgstrData = lgstrData & Chr(11) & ""
                    lgstrData = lgstrData & Chr(11) & ""
                    lgstrData = lgstrData & Chr(11) & "총합계"
                    lgstrData = lgstrData & Chr(11) & ""
                    lgstrData = lgstrData & Chr(11) & UNINumClientFormat(iSum, ggAmtOfMoney.DecPoint,0)
                    lgstrData = lgstrData & Chr(11) & lgLngMaxRow + iDx
                    lgstrData = lgstrData & Chr(11) & Chr(12)
				end if
				lgStrPrevKey1 = ""				
 		end Select
    End If   
 
    Call SubHandleError("MR",lgObjConn,lgObjRs,Err)
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
' Name : SubBizSaveMultiCreate
' Desc : Save Multi Data
'============================================================================================================
Sub SubBizSaveMultiCreate(arrColVal)
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    If Trim(arrColVal(2)) = "M" Then
        lgStrSQL = "INSERT INTO HGA050T ("
        lgStrSQL = lgStrSQL & " RETIRE_DT     ," 
        lgStrSQL = lgStrSQL & " EMP_NO     ," 
        lgStrSQL = lgStrSQL & " PAY_YYMM     ," 
        lgStrSQL = lgStrSQL & " ALLOW_CD   ," 
        lgStrSQL = lgStrSQL & " ALLOW    ," 
        lgStrSQL = lgStrSQL & " ISRT_EMP_NO," 
        lgStrSQL = lgStrSQL & " ISRT_DT     ," 
        lgStrSQL = lgStrSQL & " UPDT_EMP_NO ," 
        lgStrSQL = lgStrSQL & " UPDT_DT      )" 
        lgStrSQL = lgStrSQL & " VALUES(" 
        lgStrSQL = lgStrSQL & FilterVar(UNIConvDateCompanyToDB(arrColVal(3),NULL),"NULL","S")    & ","
        lgStrSQL = lgStrSQL & FilterVar(arrColVal(4), "''", "S")     & ","
        lgStrSQL = lgStrSQL & FilterVar(arrColVal(5), "''", "S")    & ","
        lgStrSQL = lgStrSQL & FilterVar(arrColVal(6), "''", "S")     & ","
        lgStrSQL = lgStrSQL & Uniconvnum(arrColVal(7),0)     & ","
        lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''", "S")                        & "," 
        lgStrSQL = lgStrSQL & FilterVar(lgGetSvrDateTime,NULL,"S") & "," 
        lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''", "S")                        & "," 
        lgStrSQL = lgStrSQL & FilterVar(lgGetSvrDateTime,NULL,"S")
        lgStrSQL = lgStrSQL & ")"
        
    Else
        lgStrSQL = "INSERT INTO HGA060T ("
        lgStrSQL = lgStrSQL & " RETIRE_DT     ," 
        lgStrSQL = lgStrSQL & " EMP_NO     ," 
        lgStrSQL = lgStrSQL & " BONUS_YYMM     ," 
        lgStrSQL = lgStrSQL & " BONUS_TYPE   ," 
        lgStrSQL = lgStrSQL & " BONUS    ," 
        lgStrSQL = lgStrSQL & " ISRT_EMP_NO," 
        lgStrSQL = lgStrSQL & " ISRT_DT     ," 
        lgStrSQL = lgStrSQL & " UPDT_EMP_NO ," 
        lgStrSQL = lgStrSQL & " UPDT_DT      )" 
        lgStrSQL = lgStrSQL & " VALUES(" 
        lgStrSQL = lgStrSQL & FilterVar(UNIConvDateCompanyToDB(arrColVal(3),NULL),"NULL","S")    & ","
        lgStrSQL = lgStrSQL & FilterVar(arrColVal(4), "''", "S")     & ","
        lgStrSQL = lgStrSQL & FilterVar(arrColVal(5), "''", "S")     & ","
        lgStrSQL = lgStrSQL & FilterVar(arrColVal(6), "''", "S")     & ","
        lgStrSQL = lgStrSQL & UNIConvNum(arrColVal(7),0)     & ","
        lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''", "S")                        & "," 
        lgStrSQL = lgStrSQL & FilterVar(lgGetSvrDateTime,NULL,"S") & "," 
        lgStrSQL = lgStrSQL & FilterVar(gUsrId, "''", "S")                        & "," 
        lgStrSQL = lgStrSQL & FilterVar(lgGetSvrDateTime,NULL,"S")
        lgStrSQL = lgStrSQL & ")"
    End If    
    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
	Call SubHandleError("MC",lgObjConn,lgObjRs,Err)
End Sub

'============================================================================================================
' Name : SubBizSaveMultiUpdate
' Desc : Update Data from Db
'============================================================================================================
Sub SubBizSaveMultiUpdate(arrColVal)

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    If Trim(arrColVal(2)) = "M" Then
        lgStrSQL = "UPDATE  HGA050T"
        lgStrSQL = lgStrSQL & " SET " 
        lgStrSQL = lgStrSQL & " ALLOW            = " &  UNIConvNum(arrColVal(7),0)   & ","
        lgStrSQL = lgStrSQL & " updt_emp_no      = " & FilterVar(gUsrId, "''", "S")   & ","
        lgStrSQL = lgStrSQL & " updt_dt          = " & FilterVar(lgGetSvrDateTime,NULL,"S")
        lgStrSQL = lgStrSQL & " WHERE        "
        lgStrSQL = lgStrSQL & " RETIRE_DT        = " & FilterVar(UNIConvDateCompanyToDB(arrColVal(3),NULL),"NULL","S")
        lgStrSQL = lgStrSQL & " And EMP_NO       = " & FilterVar(UCase(arrColVal(4)), "''", "S")
        lgStrSQL = lgStrSQL & " And PAY_YYMM     = " & FilterVar(arrColVal(5), "''", "S")
        lgStrSQL = lgStrSQL & " And ALLOW_CD     = " & FilterVar(UCase(arrColVal(6)), "''", "S")
    Else
        lgStrSQL = "UPDATE  HGA060T"
        lgStrSQL = lgStrSQL & " SET " 
        lgStrSQL = lgStrSQL & " BONUS            = " &  UNIConvNum(arrColVal(7),0)   & ","
        lgStrSQL = lgStrSQL & " updt_emp_no      = " & FilterVar(gUsrId, "''", "S")   & ","
        lgStrSQL = lgStrSQL & " updt_dt          = " & FilterVar(lgGetSvrDateTime,NULL,"S")
        lgStrSQL = lgStrSQL & " WHERE        "
        lgStrSQL = lgStrSQL & " RETIRE_DT        = " & FilterVar(UNIConvDateCompanyToDB(arrColVal(3),NULL),"NULL","S")
        lgStrSQL = lgStrSQL & " And EMP_NO       = " & FilterVar(UCase(arrColVal(4)), "''", "S")
        lgStrSQL = lgStrSQL & " And BONUS_YYMM   = " & FilterVar(arrColVal(5), "''", "S")
        lgStrSQL = lgStrSQL & " And BONUS_TYPE   = " & FilterVar(UCase(arrColVal(6)), "''", "S")
    End If  
    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
	Call SubHandleError("MU",lgObjConn,lgObjRs,Err)
   
End Sub


'============================================================================================================
' Name : SubBizSaveMultiDelete
' Desc : Delete Data from Db
'============================================================================================================
Sub SubBizSaveMultiDelete(arrColVal)

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    If Trim(arrColVal(2)) = "M" Then
        lgStrSQL = "DELETE  HGA050T"
        lgStrSQL = lgStrSQL & " WHERE        "
        lgStrSQL = lgStrSQL & " RETIRE_DT        = " & FilterVar(UNIConvDateCompanyToDB(arrColVal(3),NULL),"NULL","S")
        lgStrSQL = lgStrSQL & " And EMP_NO       = " & FilterVar(UCase(arrColVal(4)), "''", "S")
        lgStrSQL = lgStrSQL & " And PAY_YYMM     = " & FilterVar(arrColVal(5), "''", "S")
        lgStrSQL = lgStrSQL & " And ALLOW_CD     = " & FilterVar(UCase(arrColVal(6)), "''", "S")
    Else
        lgStrSQL = "DELETE  HGA060T"
        lgStrSQL = lgStrSQL & " WHERE        "
        lgStrSQL = lgStrSQL & " RETIRE_DT        = " & FilterVar(UNIConvDateCompanyToDB(arrColVal(3),NULL),"NULL","S")
        lgStrSQL = lgStrSQL & " And EMP_NO       = " & FilterVar(UCase(arrColVal(4)), "''", "S")
        lgStrSQL = lgStrSQL & " And BONUS_YYMM   = " & FilterVar(arrColVal(5), "''", "S")
        lgStrSQL = lgStrSQL & " And BONUS_TYPE   = " & FilterVar(UCase(arrColVal(6)), "''", "S")
    End If    
    
    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
	Call SubHandleError("MD",lgObjConn,lgObjRs,Err)

End Sub

'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements(pDataType,pCode,pCode1,pComp)
    Dim iSelCount

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    Select Case Mid(pDataType,1,1)
        Case "S"
           iSelCount = C_SHEETMAXROWS_D + C_SHEETMAXROWS_D *  lgStrPrevKey1 + 1        
           Select Case Mid(pDataType,2,1)
               Case "R"
                        lgStrSQL = "Select " 
                        lgStrSQL = lgStrSQL & " hga070t.retire_yy,hga070t.emp_no,hga070t.entr_dt,hga070t.retire_dt,hga070t.duty_mm, "
                        lgStrSQL = lgStrSQL & " hga070t.pay_avr,hga070t.bonus_avr,hga070t.year_avr,hga070t.avr_wages,hga070t.retire_amt, "
                        lgStrSQL = lgStrSQL & " hga070t.corp_insur,hga070t.honor_amt,hga070t.etc_amt,hga070t.income_tax,hga070t.res_tax, "
                        lgStrSQL = lgStrSQL & " hga070t.income_sub,hga070t.special_sub,(hga040t.etc_sub1 + hga040t.etc_sub2 + hga040t.etc_sub3 + hga040t.etc_sub4) etc_sub, hga070t.income_amt,hga070t.sub_short_amt,hga070t.tax_std, "
                        lgStrSQL = lgStrSQL & " hga070t.avr_tax_std,hga070t.tax_rate,hga070t.duty_cnt,hga070t.calc_tax,hga070t.avr_calc_tax, "
                        lgStrSQL = lgStrSQL & " hga070t.tax_short_amt,hga070t.deci_income_tax,hga070t.deci_res_tax,hga070t.retire_yy, "
                        lgStrSQL = lgStrSQL & " hga070t.emp_no,hga070t.duty_yy,hga070t.duty_dd,hga070t.tot_duty_mm,hga070t.tot_duty_dd, "
                        lgStrSQL = lgStrSQL & " hga070t.pay_tot_amt,hga070t.avr_day,hga070t.bonus_tot_amt,hga070t.year_tot_amt,hga070t.yy_retire_amt, "
                        lgStrSQL = lgStrSQL & " hga070t.mm_retire_amt,hga070t.dd_retire_amt,hga070t.tot_prov_amt,hga070t.real_prov_amt, "
                        lgStrSQL = lgStrSQL & " hga070t.isrt_emp_no,hga070t.isrt_dt,hga070t.updt_emp_no,hga070t.updt_dt,hga070t.retire_anu_amt "
                        lgStrSQL = lgStrSQL & " From  HGA070T left outer join HGA040T on hga070t.emp_no = hga040t.emp_no and hga070t.retire_dt = hga040t.retire_dt "
                        lgStrSQL = lgStrSQL & " WHERE hga070t.emp_no " & pComp & pCode 	

           End Select  
                      
	Case "M"
           Select Case Mid(pDataType,2,1)
               Case "R"
                       If lgCurrentSpd = "M" Then
				          iSelCount = C_SHEETMAXROWS_D + C_SHEETMAXROWS_D *  lgStrPrevKey + 1                       
                
                          lgStrSQL = "Select  TOP " & iSelCount
                          lgStrSQL = lgStrSQL & " CASE WHEN (GROUPING(HGA050T.ALLOW_CD) = 1 AND GROUPING(HGA050T.PAY_YYMM) = 1) THEN '' "
                          lgStrSQL = lgStrSQL & "      WHEN (GROUPING(HGA050T.PAY_YYMM) = 0 AND GROUPING(HGA050T.ALLOW_CD) = 1) THEN '' "
                          lgStrSQL = lgStrSQL & "      ELSE ISNULL(HGA050T.PAY_YYMM, '') END AS PAY_YYMM , "
                          lgStrSQL = lgStrSQL & "      ISNULL(HGA050T.ALLOW_CD, '')  ALLOW_CD , "
                          lgStrSQL = lgStrSQL & " CASE WHEN (GROUPING(HGA050T.PAY_YYMM) = 0 AND GROUPING(HGA050T.ALLOW_CD) = 1) THEN " & FilterVar("월소계", "''", "S") & " "
                          lgStrSQL = lgStrSQL & "      WHEN (GROUPING(HGA050T.PAY_YYMM) = 1 AND GROUPING(HGA050T.ALLOW_CD) = 1) THEN '총합계' "
                          lgStrSQL = lgStrSQL & "      ELSE dbo.ufn_H_GetCodeName(" & FilterVar("HDA010T", "''", "S") & ", HGA050T.ALLOW_CD, " & FilterVar("1", "''", "S") & " ) END AS ALLOW_NM , "
                          lgStrSQL = lgStrSQL & " SUM(HGA050T.ALLOW) AS ALLOW "
                          lgStrSQL = lgStrSQL & " From  hga050t "
                          lgStrSQL = lgStrSQL & " Where hga050t.emp_no " & pComp & pCode
                       Else
				          iSelCount = C_SHEETMAXROWS_D + C_SHEETMAXROWS_D *  lgStrPrevKey1 + 1                       

                          lgStrSQL = "Select TOP " & iSelCount
                          lgStrSQL = lgStrSQL & " hga060t.bonus_type, dbo.ufn_GetCodeName(" & FilterVar("H0040", "''", "S") & ", hga060t.bonus_type) bonus_type_NM, hga060t.bonus,hga060t.retire_dt,hga060t.emp_no,hga060t.bonus_yymm, "
                          lgStrSQL = lgStrSQL & " hga060t.isrt_emp_no,hga060t.isrt_dt,hga060t.updt_emp_no,hga060t.updt_dt "
                          lgStrSQL = lgStrSQL & " From  hga060t   "
                          lgStrSQL = lgStrSQL & " Where hga060t.emp_no " & pComp & pCode1                          
                       End If      
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
    On Error Resume Next                                                              '☜: Protect system from crashing
    Err.Clear                                                                         '☜: Clear Error status

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
                 If CheckSYSTEMError(pErr,True) = True Then
                    ObjectContext.SetAbort
                    Call SetErrorStatus
                 Else
                    If CheckSQLError(pConn,True) = True Then
                       ObjectContext.SetAbort
                       Call SetErrorStatus
                    End If
                 End If
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
       Case "<%=UID_M0001%>"                                                         '☜ : Query
			                   
          If Trim("<%=lgErrorStatus%>") = "NO" Then
              With Parent
                If Trim(.lgCurrentSpd) = "M" Then
                   .ggoSpread.Source     = .frm1.vspdData
                   .lgStrPrevKey    = "<%=lgStrPrevKey%>"
					.ggoSpread.SSShowData "<%=lgstrData%>"
					.frm1.Sflag.value = "<%=Sflag%>"                            
					if .topleftOK then
						.DBQueryOk
					else
						.lgCurrentSpd = "S"
						.DBQuery
					end if
				Else
					.ggoSpread.Source     = .frm1.vspdData1
					.lgStrPrevKey1    = "<%=lgStrPrevKey1%>"
					.ggoSpread.SSShowData "<%=lgstrData%>"
					.frm1.Sflag.value = "<%=Sflag%>"                            
					.DBQueryOK
                End If  
	         End with
          End If           
          
       Case "<%=UID_M0002%>"                                                         '☜ : Save
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DBSaveOk
          Else
             Parent.SubSetErrPos(Trim("<%=lgErrorPos%>"))
          End If   
       Case "<%=UID_M0002%>"                                                         '☜ : Delete
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DbDeleteOk
          Else   
          End If   
    End Select    
       
</Script>	
