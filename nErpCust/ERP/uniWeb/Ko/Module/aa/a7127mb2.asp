<%@ LANGUAGE="VBScript" CODEPAGE=949 TRANSACTION=Required %>
<% Option Explicit%>
<% session.CodePage=949 %>

<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/IncSvrMain.asp"  -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc"  -->
<!-- #Include file="../../inc/IncSvrNumber.inc"  -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->

<%

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

	Call LoadBasisGlobalInf()
	Call LoadInfTB19029B("I", "*","NOCOOKIE","MB")
	Call LoadBNumericFormatB("I", "*","NOCOOKIE","MB")

	Dim lgstrDataTotal
	Dim lgStrPrevKey
	Dim lgAsstChgNo
	Dim gIsShowLocal
	Dim iStrCurrency

    Const C_SHEETMAXROWS_D  = 100 

    Call HideStatusWnd                                                               '☜: Hide Processing message
    '---------------------------------------Common-----------------------------------------------------------
    lgErrorStatus		= "NO"
    lgErrorPos			= ""                                                           '☜: Set to space
	gIsShowLocal		= "Y"
    lgOpModeCRUD		= Request("txtMode")                                           '☜: Read Operation Mode (CRUD)
    lgAsstChgNo			= Request("txtAsstChgNo")
    
	' 권한관리 추가
	Dim lgAuthBizAreaCd, lgAuthBizAreaNm			' 사업장
	Dim lgInternalCd, lgDeptCd, lgDeptNm			' 내부부서
	Dim lgSubInternalCd, lgSubDeptCd, lgSubDeptNm	' 내부부서(하위포함)
	Dim lgAuthUsrID, lgAuthUsrNm					' 개인

	Dim lgBizAreaAuthSQL, lgInternalCdAuthSQL, lgSubInternalCdAuthSQL, lgAuthUsrIDAuthSQL    
 
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
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    Call SubBizQuerySingle()
    Call SubBizQuerySingle2()
	Call SubBizQueryMulti()
	Call SubBizQueryMulti2()   
End Sub	
'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery2()
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
End Sub	

'============================================================================================================
' Name : SubBizSave
' Desc : Date data 
'============================================================================================================
Sub SubBizSave()
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
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
' Name : SubBizQuerySingle
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuerySingle()
    Dim iDx
    Dim iKey1
    Dim strWhere
    Dim YYYYMM

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    Call SubMakeSQLStatements("MR",strWhere,"X",C_LIKE)                                 '☆: Make sql statements

    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        lgStrPrevKeyIndex = ""
       
		Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 

		response.end
        Call SetErrorStatus()
    Else
		Response.Write "<Script Language=vbscript>				 " & vbCr
		Response.Write " With parent                                        " & vbCr
		Response.Write "       	parent.frm1.txtAsstChgNo.value		= """ & ConvSPChars(lgObjRs("ASST_CHG_NO")) & """" & vbCr          
		Response.Write "       	parent.frm1.txtAsstChgNo2.value		= """ & ConvSPChars(lgObjRs("ASST_CHG_NO")) & """" & vbCr          

		If  ConvSPChars(lgObjRs("CHG_FG")) = "03" Then
			Response.Write " parent.frm1.Rb_Sold.Checked      = """ & True & """									" & vbCr    
			Response.Write " parent.frm1.txtRadio.value = """ & ConvSPChars(lgObjRs("CHG_FG"))  & """" & vbCr
		Else 
			Response.Write " parent.frm1.Rb_Duse.Checked      = """ & True & """									" & vbCr    
			Response.Write " parent.frm1.txtRadio.value = """ & ConvSPChars(lgObjRs("CHG_FG"))  & """" & vbCr    
		End If	
        iStrCurrency = lgObjRs("DOC_CUR")

		Response.Write "       	parent.frm1.txtChgDt.text			= """ & UNIDateClientFormat(lgObjRs("CHG_DT")) & """" & vbCr          
		Response.Write "       	parent.frm1.txtDeptCd.value			= """ & Trim(ConvSPChars(lgObjRs("FROM_DEPT_CD"))) & """" & vbCr          
		Response.Write "       	parent.frm1.txtDeptNm.value			= """ & ConvSPChars(lgObjRs("DEPT_NM")) & """" & vbCr          
		Response.Write "       	parent.frm1.hORGCHANGEID.value		= """ & ConvSPChars(lgObjRs("FROM_ORG_CHANGE_ID")) & """" & vbCr          
		Response.Write "       	parent.frm1.txtDocCur.value			= """ & ConvSPChars(lgObjRs("DOC_CUR")) & """" & vbCr          
		Response.Write "       	parent.frm1.txtXchRate.value		= """ & UNINumClientFormat(lgObjRs("XCH_RATE"), ggExchRate.DecPoint, 0) & """" & vbCr          
		Response.Write "       	parent.frm1.txtBpCd.value			= """ & ConvSPChars(lgObjRs("BP_CD")) & """" & vbCr          
		Response.Write "       	parent.frm1.txtBpNm.value			= """ & ConvSPChars(lgObjRs("BP_NM")) & """" & vbCr          
		Response.Write "       	parent.frm1.txtChgDesc.value		= """ & ConvSPChars(lgObjRs("ASST_CHG_DESC")) & """" & vbCr          
		Response.Write "       	parent.frm1.txtTempGlNo.value		= """ & ConvSPChars(lgObjRs("TEMP_GL_NO")) & """" & vbCr          
		Response.Write "       	parent.frm1.txtGlNo.value			= """ & ConvSPChars(lgObjRs("GL_NO")) & """" & vbCr          
		Response.Write "       	parent.frm1.txtVatType.value		= """ & ConvSPChars(lgObjRs("TAX_TYPE")) & """" & vbCr          
		Response.Write "       	parent.frm1.txtVatTypeNm.value		= """ & ConvSPChars(lgObjRs("MINOR_NM")) & """" & vbCr          
		Response.Write "       	parent.frm1.txtVatRate.Text			= """ & UNINumClientFormat(lgObjRs("TAX_RATE"), ggExchRate.DecPoint, 0)  & """" & vbCr          
		Response.Write "       	parent.frm1.txtVatAmt.value			= """ & UNIConvNumDBToCompanyByCurrency(lgObjRs("TAX_TOT_AMT"), gCurrency,ggAmtOfMoneyNo, "X" , "X")  & """" & vbCr          
		Response.Write "       	parent.frm1.txtVatLocAmt.Text		= """ & UNIConvNumDBToCompanyByCurrency(lgObjRs("TAX_TOT_LOC_AMT"), gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo, "X") & """" & vbCr          
		Response.Write "       	parent.frm1.txtReportAreaCd.value	= """ & ConvSPChars(lgObjRs("REPORT_BIZ_AREA_CD")) & """" & vbCr          
		Response.Write "       	parent.frm1.txtReportAreaNm.value	= """ & ConvSPChars(lgObjRs("TAX_BIZ_AREA_NM")) & """" & vbCr          
		Response.Write "       	parent.frm1.txtIssuedDt.Text		= """ & UNIDateClientFormat(lgObjRs("ISSUED_DT")) & """" & vbCr          
		Response.Write " End With   " & vbCr
		Response.Write "</Script>				 " & vbCr
    End If

    Call SubHandleError("MR",lgObjConn,lgObjRs,Err)
    Call SubCloseRs(lgObjRs)                                                          '☜: Release RecordSSet
End Sub

'============================================================================================================
' Name : SubBizQuerySingle     
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuerySingle2()
    Dim iDx
    Dim iKey1
    Dim strWhere
    Dim YYYYMM

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    Call SubMakeSQLStatements("MR",strWhere,"5",C_LIKE)                                 '☆: Make sql statements

    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        lgStrPrevKeyIndex = ""
    Else
		Response.Write "<Script Language=vbscript>				 " & vbCr
		Response.Write " With parent                                        " & vbCr
		Response.Write "       	.frm1.txtTotalAmt.Text			= """ & UNIConvNumDBToCompanyByCurrency(lgObjRs("CHG_AMT"),iStrCurrency,ggAmtOfMoneyNo, "X" , "X") & """" & vbCr          	
		Response.Write "       	.frm1.txtTotalLocAmt.Text		= """ & UNIConvNumDBToCompanyByCurrency(lgObjRs("CHG_LOC_AMT"), gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo, "X")  & """" & vbcr
		Response.Write "       	.frm1.txtTotalRcptAmt.value		= """ & UNIConvNumDBToCompanyByCurrency(lgObjRs("PAYM_AMT"),iStrCurrency,ggAmtOfMoneyNo, "X" , "X") & """" & vbCr          
		Response.Write "       	.frm1.txtTotalRcptLocAmt.value	= """ & UNIConvNumDBToCompanyByCurrency(lgObjRs("PAYM_LOC_AMT"), gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo, "X")  & """" & vbCr          
		Response.Write " End With   " & vbCr
		Response.Write "</Script>				 " & vbCr
    End If

    Call SubHandleError("MR",lgObjConn,lgObjRs,Err)
	Call SubCloseRs(lgObjRs)                                                          '☜: Release RecordSSet
End Sub

'============================================================================================================
' Name : SubBizQueryMulti1    두번째 dbqueryok()에서 호출된 두번째 
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()
    Dim iDx
    Dim iKey1
    Dim strWhere
    Dim YYYYMM

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    Call SubMakeSQLStatements("MR",strWhere,"1",C_LIKE)                                 '☆: Make sql statements

    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        lgStrPrevKeyIndex = ""
    Else
        lgstrData = ""
        iDx = 1

        Do While Not lgObjRs.EOF
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("CHG_NO"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ASST_CD"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ASST_NM"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("SUB_NO"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("FROM_DEPT_CD"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("DEPT_NM"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ORG_CHANGE_ID"))
            lgstrData = lgstrData & Chr(11) & UNIDateClientFormat(lgObjRs("REG_DT"))
            lgstrData = lgstrData & Chr(11) & UNIConvNum(lgObjRs("INV_QTY"),0)
            lgstrData = lgstrData & Chr(11) & UNIConvNum(lgObjRs("CHG_QTY"),0)
            lgstrData = lgstrData & Chr(11) & UNINumClientFormat(lgObjRs("SOLD_RATE"), ggExchRate.DecPoint, 0)
            lgstrData = lgstrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(lgObjRs("CHG_AMT"),iStrCurrency,ggAmtOfMoneyNo, "X" , "X")
            lgstrData = lgstrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(lgObjRs("CHG_LOC_AMT"),gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo , "X")
            lgstrData = lgstrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(lgObjRs("DECR_ACQ_LOC_AMT"),gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo , "X")
            lgstrData = lgstrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(lgObjRs("DEPR_TOT_LOC_AMT"),gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo , "X") 'NET_LOC_AMT
            lgstrData = lgstrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(lgObjRs("DEPR_TOT_LOC_AMT_CHG"),gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo , "X")'DECR_ACQ_LOC_AMT
            lgstrData = lgstrData & Chr(11) & "" 'UNIConvNumDBToCompanyByCurrency(lgObjRs("TOT_MNTH_DEPR_AMT"),gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo , "X")'MNTH_DEPR_AMT

            lgstrData = lgstrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(lgObjRs("VAT_AMT"),iStrCurrency,ggAmtOfMoneyNo, "X" , "X")
            lgstrData = lgstrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(lgObjRs("VAT_LOC_AMT"),gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo , "X")
            lgstrData = lgstrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(lgObjRs("DECR_ACQ_LOC_AMT"),gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo , "X")
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ASSET_CHG_DESC"))
			lgstrData = lgstrData & Chr(11) & iDx'lgLngMaxRow + iDx
			lgstrData = lgstrData & Chr(11) & Chr(12)

			lgObjRs.MoveNext
			iDx = iDx + 1
		Loop 
    End If
    
    If iDx <= lgMaxCount Then
		lgStrPrevKeyIndex = ""
    End If   

	Response.Write " <Script Language=vbscript>				" & vbCr
	Response.Write " 	parent.ggoSpread.Source        = parent.frm1.vspdData	      " & vbCr
	Response.Write " 	parent.ggoSpread.SSShowData """ & lgstrData   & """ ,""F""" & vbCr
	Response.Write "    Call parent.ReFormatSpreadCellByCellByCurrency2(Parent.Frm1.vspdData," & 1 & "," & iDx & "," & iStrCurrency & ",parent.C_ChgAmt,   ""A"" ,""I"",""X"",""X"")" & vbCr
	Response.Write " </Script>								" & vbCr

    Call SubCloseRs(lgObjRs)                                                          '☜: Release RecordSSet
End Sub

'============================================================================================================
' Name : SubBizQueryMulti1    두번째 dbqueryok()에서 호출된 두번째 
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti2()
    Dim iDx
    Dim iKey1
    Dim strWhere
    Dim YYYYMM

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

	iDx = 1

    Call SubMakeSQLStatements("MR",strWhere,"10",C_LIKE)                                 '☆: Make sql statements
    
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then

    Else
        lgstrData = ""

        Do While Not lgObjRs.EOF
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("PAYM_TYPE")) ' 입금유형 
            lgstrData = lgstrData & Chr(11) & ""
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("MINOR_NM")) ' 입금유형 
            lgstrData = lgstrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(lgObjRs("PAYM_AMT"),iStrCurrency,ggAmtOfMoneyNo, "X" , "X")
            lgstrData = lgstrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(lgObjRs("PAYM_LOC_AMT"),gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo , "X")
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("AR_AP_NO")) '어음번호 
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("AR_AP_ACCT_CD"))
            lgstrData = lgstrData & Chr(11) & ""
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ACCT_NM"))
            lgstrData = lgstrData & Chr(11) & UNIDateClientFormat(lgObjRs("AR_AP_DUE_DT"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("BANK_CD")) '은행코드 
            lgstrData = lgstrData & Chr(11) & ""
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("BANK_NM")) '은행명 
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("BANK_ACCT_NO")) '계좌번호 
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("NOTE_NO")) '어음번호 
            lgstrData = lgstrData & Chr(11) & ""
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ASST_CHG_ITEM_DESC")) '적요 
			lgstrData = lgstrData & Chr(11) & lgLngMaxRow + iDx
			lgstrData = lgstrData & Chr(11) & Chr(12)

			lgObjRs.MoveNext
			iDx = iDx + 1
		Loop 
    End If

    If iDx <= lgMaxCount Then
		lgStrPrevKeyIndex = ""
    End If   

	Response.Write "<Script Language=vbscript>				" & vbCr
	Response.Write " 	parent.ggoSpread.Source        = parent.frm1.vspdData2	      " & vbCr
	Response.Write " 	parent.ggoSpread.SSShowData """ & lgstrData   & """ ,""F""" & vbCr
	Response.Write  "    Call parent.ReFormatSpreadCellByCellByCurrency2(Parent.Frm1.vspdData2," & 1 & "," & iDx & "," & iStrCurrency & ",parent.C_RcptAmt,   ""A"" ,""I"",""X"",""X"")" & vbCr
	Response.Write " </Script>								" & vbCr

    Call SubCloseRs(lgObjRs)                                                          '☜: Release RecordSSet
End Sub    

'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data 
'============================================================================================================
Sub SubBizSaveMulti()
    On Error Resume Next                                                             '☜: Protect system from crashing
	Err.Clear 
End Sub    

'============================================================================================================
' Name : SubBizSaveMultiCreate
' Desc : Save Multi Data
'============================================================================================================
Sub SubBizSaveMultiCreate(arrColVal)
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
End Sub

'============================================================================================================
' Name : SubBizSaveMultiUpdate
' Desc : Update Data from Db
'============================================================================================================
Sub SubBizSaveMultiUpdate(arrColVal)
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
End Sub


'============================================================================================================
' Name : SubBizSaveMultiDelete
' Desc : Delete Data from Db
'============================================================================================================
Sub SubBizSaveMultiDelete(arrColVal)
	On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
End Sub

'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements(pDataType,pCode,pCode1,pComp)
    Dim iSelCount

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear    

	' 권한관리 추가
	lgAuthBizAreaCd		= Trim(Request("lgAuthBizAreaCd"))
	lgInternalCd		= Trim(Request("lgInternalCd"))
	lgSubInternalCd		= Trim(Request("lgSubInternalCd"))
	lgAuthUsrID		= Trim(Request("lgAuthUsrID"))
	
	Select Case Mid(pDataType,1,1)
        Case "S"
	       Select Case  lgPrevNext 
                  Case " "
                  Case "P"
                  Case "N"
           End Select
        Case "M"
           Select Case Mid(pDataType,2,1)
               Case "R"
				   Select Case pCode1
					   Case "X"
							lgStrSQL = ""
							lgStrSQL = " SELECT A.ASST_CHG_NO, A.CHG_FG, A.CHG_DT, A.FROM_DEPT_CD, B.DEPT_NM , A.FROM_ORG_CHANGE_ID, " & vbcr
							lgStrSQL = lgStrSQL & "	A.DOC_CUR, A.XCH_RATE, A.BP_CD, C.BP_NM ,A.ASST_CHG_DESC,  " & vbcr
							lgStrSQL = lgStrSQL & "	A.TEMP_GL_NO, A.GL_NO, A.TAX_TYPE, A.TAX_RATE, D.MINOR_NM, A.TAX_TOT_AMT, A.TAX_TOT_LOC_AMT,  " & vbcr
							lgStrSQL = lgStrSQL & "	A.REPORT_BIZ_AREA_CD, E.TAX_BIZ_AREA_NM , A.ISSUED_DT  " & vbcr
							lgStrSQL = lgStrSQL & " FROM A_ASSET_CHG_MASTER A, B_ACCT_DEPT B, B_BIZ_PARTNER C, B_MINOR D, B_TAX_BIZ_AREA E " & vbcr
							lgStrSQL = lgStrSQL & " WHERE  A.BP_CD *= C.BP_CD AND " & vbcr
							lgStrSQL = lgStrSQL & "	A.FROM_ORG_CHANGE_ID *= B.ORG_CHANGE_ID AND  " & vbcr
							lgStrSQL = lgStrSQL & "	A.FROM_DEPT_CD *= B.DEPT_CD AND " & vbcr
							lgStrSQL = lgStrSQL & "	D.MAJOR_CD = " & FilterVar("B9001", "''", "S") & "  AND " & vbcr
							lgStrSQL = lgStrSQL & "	A.TAX_TYPE *= D.MINOR_CD AND " & vbcr
							lgStrSQL = lgStrSQL & "	A.REPORT_BIZ_AREA_CD *= E.TAX_BIZ_AREA_CD AND " & vbcr
							lgStrSQL = lgStrSQL & "	A.ASST_CHG_NO =  " & FilterVar(UCase(lgAsstChgNo), "''", "S") & " " & vbcr
							' 권한관리 추가
							If lgAuthBizAreaCd <> "" Then
								lgBizAreaAuthSQL		= " AND A.FROM_BIZ_AREA_CD = " & FilterVar(lgAuthBizAreaCd, "''", "S")
							End If
	
							If lgInternalCd <> "" Then
								lgInternalCdAuthSQL		= " AND A.FROM_INTERNAL_CD = " & FilterVar(lgInternalCd, "''", "S")
							End If
	
							If lgSubInternalCd <> "" Then
								lgSubInternalCdAuthSQL	= " AND A.FROM_INTERNAL_CD LIKE " & FilterVar(lgSubInternalCd & "%", "''", "S")
							End If
	
							If lgAuthUsrID <> "" Then
								lgAuthUsrIDAuthSQL		= " AND A.UPDT_USER_ID = " & FilterVar(lgAuthUsrID, "''", "S")
							End If							
							
							lgStrSQL = lgStrSQL & lgBizAreaAuthSQL & lgInternalCdAuthSQL & lgSubInternalCdAuthSQL & lgAuthUsrIDAuthSQL
					   Case "1"
							lgStrSQL = ""
							lgStrSQL = lgStrSQL & " SELECT	A.CHG_NO,   " & vbcr
							lgStrSQL = lgStrSQL & "	 A.ASST_CD,   " & vbcr
							lgStrSQL = lgStrSQL & "	 C.ASST_NM,   " & vbcr
							lgStrSQL = lgStrSQL & "	 A.FROM_DEPT_CD,   " & vbcr
							lgStrSQL = lgStrSQL & "	 E.DEPT_NM,   " & vbcr
							lgStrSQL = lgStrSQL & "	 E.ORG_CHANGE_ID,   " & vbcr
							lgStrSQL = lgStrSQL & "	 C.REG_DT,   " & vbcr
							lgStrSQL = lgStrSQL & "	 C.INV_QTY,   " & vbcr
							lgStrSQL = lgStrSQL & "	 A.CHG_QTY,   " & vbcr
							lgStrSQL = lgStrSQL & "	 IsNull(A.CHG_AMT, 0) CHG_AMT,   " & vbcr
							lgStrSQL = lgStrSQL & "	 IsNull(A.CHG_LOC_AMT, 0) CHG_LOC_AMT,   " & vbcr
							lgStrSQL = lgStrSQL & "	 IsNull(A.DECR_ACQ_LOC_AMT, 0) DECR_ACQ_LOC_AMT ,   " & vbcr '감소취득금액'
							lgStrSQL = lgStrSQL & "	 IsNull(A.DEPR_TOT_LOC_AMT, 0) DEPR_TOT_LOC_AMT,   " & vbcr '감소상각누계액'
							lgStrSQL = lgStrSQL & "	 IsNull(A.DECR_ACQ_LOC_AMT - A.DEPR_TOT_LOC_AMT, 0) DEPR_TOT_LOC_AMT_CHG,   " & vbcr '자산변동액'
							lgStrSQL = lgStrSQL & "	 A.VAT_RATE,   " & vbcr
							lgStrSQL = lgStrSQL & "	 IsNull(A.VAT_AMT, 0) VAT_AMT,   " & vbcr
							lgStrSQL = lgStrSQL & "	 IsNull(A.VAT_LOC_AMT, 0) VAT_LOC_AMT,   " & vbcr
							lgStrSQL = lgStrSQL & "	 A.ASSET_CHG_DESC,   " & vbcr
							lgStrSQL = lgStrSQL & "	 B.ASST_CHG_NO," & vbcr
							lgStrSQL = lgStrSQL & "	 A.SOLD_RATE," & vbcr
							lgStrSQL = lgStrSQL & "	 A.SUB_NO" & vbcr
							lgStrSQL = lgStrSQL & " FROM	A_ASSET_CHG A," & vbcr
							lgStrSQL = lgStrSQL & "	 A_ASSET_CHG_MASTER B," & vbcr
							lgStrSQL = lgStrSQL & "	 A_ASSET_MASTER C," & vbcr
		'					lgStrSQL = lgStrSQL & "	 A_ASSET_INFORM_OF_DEPT D," & vbcr
							lgStrSQL = lgStrSQL & "	 B_ACCT_DEPT E" & vbcr
							lgStrSQL = lgStrSQL & " WHERE B.ASST_CHG_NO =  " & FilterVar(UCase(lgAsstChgNo), "''", "S") & " " & vbcr
							lgStrSQL = lgStrSQL & "	 AND A.ASST_CHG_NO = B.ASST_CHG_NO" & vbcr
							lgStrSQL = lgStrSQL & "	 AND A.ASST_CD = C.ASST_NO" & vbcr
		'					lgStrSQL = lgStrSQL & "	 AND A.ASST_CD = D.ASST_NO" & vbcr
		'					lgStrSQL = lgStrSQL & "	 AND D.DEPT_CD = A.FROM_DEPT_CD" & vbcr
		'					lgStrSQL = lgStrSQL & "	 AND D.ORG_CHANGE_ID = A.FROM_ORG_CHANGE_ID  " & vbcr
							lgStrSQL = lgStrSQL & "	 AND E.DEPT_CD = A.FROM_DEPT_CD" & vbcr
							lgStrSQL = lgStrSQL & "	 AND E.ORG_CHANGE_ID = A.FROM_ORG_CHANGE_ID" & vbcr
					   Case "5"
							lgStrSQL = ""

							lgStrSQL = lgStrSQL & "	SELECT SUM(CHG_AMT) CHG_AMT,   " & vbcr
							lgStrSQL = lgStrSQL & "		SUM(CHG_LOC_AMT) CHG_LOC_AMT,   " & vbcr
							lgStrSQL = lgStrSQL & "		SUM(PAYM_AMT) PAYM_AMT,   " & vbcr
							lgStrSQL = lgStrSQL & "		SUM(PAYM_LOC_AMT) PAYM_LOC_AMT   " & vbcr
							lgStrSQL = lgStrSQL & "		FROM " & vbcr
							lgStrSQL = lgStrSQL & "		( " & vbcr
							lgStrSQL = lgStrSQL & "		SELECT SUM(A.CHG_AMT) CHG_AMT, SUM(A.CHG_LOC_AMT) CHG_LOC_AMT,0 PAYM_AMT,0  PAYM_LOC_AMT FROM A_ASSET_CHG A " & vbcr
							lgStrSQL = lgStrSQL & "		WHERE A.ASST_CHG_NO =  " & FilterVar(UCase(lgAsstChgNo), "''", "S") & " " & vbcr
							lgStrSQL = lgStrSQL & "		UNION " & vbcr
							lgStrSQL = lgStrSQL & "		SELECT 0 CHG_AMT,0 CHG_LOC_AMT,SUM(B.PAYM_AMT) PAYM_AMT, SUM(B.PAYM_LOC_AMT) PAYM_LOC_AMT FROM A_ASSET_CHG A, A_ASSET_CHG_ITEM B " & vbcr
							lgStrSQL = lgStrSQL & "		WHERE A.CHG_NO = B.CHG_NO AND A.ASST_CHG_NO =  " & FilterVar(UCase(lgAsstChgNo), "''", "S") & " " & vbcr
							lgStrSQL = lgStrSQL & "		) A " & vbcr
					   Case "10"
							lgStrSQL = ""
							lgStrSQL = lgStrSQL & "	  SELECT DISTINCT A.CHG_NO,			" & vbcr
							lgStrSQL = lgStrSQL & "		     A.CHG_SEQ,			" & vbcr
							lgStrSQL = lgStrSQL & "		     A.PAYM_TYPE,		" & vbcr
							lgStrSQL = lgStrSQL & "		     A.PAYM_AMT,		" & vbcr
							lgStrSQL = lgStrSQL & "		     F.MINOR_NM,		" & vbcr
							lgStrSQL = lgStrSQL & "		     A.PAYM_LOC_AMT,	" & vbcr
							lgStrSQL = lgStrSQL & "		     A.NOTE_NO,			" & vbcr
							lgStrSQL = lgStrSQL & "		     A.BANK_CD,			" & vbcr
							lgStrSQL = lgStrSQL & "		     D.BANK_NM,			" & vbcr
							lgStrSQL = lgStrSQL & "		     A.BANK_ACCT_NO,	" & vbcr
							lgStrSQL = lgStrSQL & "		     A.AR_AP_DUE_DT,	" & vbcr
							lgStrSQL = lgStrSQL & "		     A.AR_AP_NO,		" & vbcr
							lgStrSQL = lgStrSQL & "		     A.AR_AP_ACCT_CD,	" & vbcr
							lgStrSQL = lgStrSQL & "		     C.ACCT_NM,	" & vbcr
							lgStrSQL = lgStrSQL & "		     A.ASST_CHG_ITEM_DESC,	" & vbcr
							lgStrSQL = lgStrSQL & "		     B.ASST_CHG_NO		" & vbcr
							lgStrSQL = lgStrSQL & "	  FROM   A_ASSET_CHG_ITEM A,	" & vbcr
							lgStrSQL = lgStrSQL & "		     A_ASSET_CHG_MASTER B,	" & vbcr
							lgStrSQL = lgStrSQL & "		     A_ACCT C,	" & vbcr
							lgStrSQL = lgStrSQL & "		     B_BANK D,	" & vbcr
							lgStrSQL = lgStrSQL & "		     B_BANK_ACCT E,	" & vbcr
							lgStrSQL = lgStrSQL & "		     ( SELECT A.MINOR_CD MINOR_CD, A.MINOR_NM MINOR_NM 	" & vbcr '200309222 jsk
							lgStrSQL = lgStrSQL & "		     FROM B_MINOR A, B_CONFIGURATION B	" & vbcr
							lgStrSQL = lgStrSQL & "		     WHERE (A.MINOR_CD = B.MINOR_CD AND A.MAJOR_CD = B.MAJOR_CD) 	" & vbcr
							lgStrSQL = lgStrSQL & "		     AND (A.MAJOR_CD = " & FilterVar("A1006", "''", "S") & " ) 	" & vbcr
							lgStrSQL = lgStrSQL & "		     AND A.MINOR_CD NOT IN ( " & FilterVar("NP", "''", "S") & " , " & FilterVar("PP", "''", "S") & " , " & FilterVar("AP", "''", "S") & " , " & FilterVar("CP", "''", "S") & "  , " & FilterVar("NE", "''", "S") & " , " & FilterVar("PR", "''", "S") & " ) AND B.SEQ_NO = 4 	" & vbcr
							lgStrSQL = lgStrSQL & "		     UNION ALL " & vbcr
							lgStrSQL = lgStrSQL & "		     SELECT " & FilterVar("AR", "''", "S") & "  MINOR_CD, " & FilterVar("미수금", "''", "S") & "  MINOR_NM ) F	" & vbcr
							lgStrSQL = lgStrSQL & "	  WHERE  A.ASST_CHG_NO = B.ASST_CHG_NO AND	" & vbcr
							lgStrSQL = lgStrSQL & "		    A.AR_AP_ACCT_CD *= C.ACCT_CD AND " & vbcr
							lgStrSQL = lgStrSQL & "		    D.BANK_CD =* A.BANK_CD AND " & vbcr
							lgStrSQL = lgStrSQL & "		    E.BANK_CD =* A.BANK_CD AND " & vbcr
							lgStrSQL = lgStrSQL & "		    A.PAYM_TYPE *= F.MINOR_CD AND " & vbcr '200309222 jsk
							lgStrSQL = lgStrSQL & "		     B.ASST_CHG_NO =  " & FilterVar(UCase(lgAsstChgNo), "''", "S") & " " & vbcr
							lgStrSQL = lgStrSQL & "	  ORDER BY A.CHG_SEQ ASC	" & vbcr
				  End Select             
          End Select             
    End Select

	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub

'============================================================================================================
' Name : CommonOnTransactionCommit
' Desc : This Sub is called by OnTransactionCommit Error handler
'============================================================================================================
Sub CommonOnTransactionCommit()
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub

'============================================================================================================
' Name : CommonOnTransactionAbort
' Desc : This Sub is called by OnTransactionAbort Error handler
'============================================================================================================
Sub CommonOnTransactionAbort()
    lgErrorStatus    = "YES"
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub

'============================================================================================================
' Name : SetErrorStatus
' Desc : This Sub set error status
'============================================================================================================
Sub SetErrorStatus()
    lgErrorStatus     = "YES"                                                         '☜: Set error status

	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
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
                   .ggoSpread.Source     = .frm1.vspdData2
                   .lgStrPrevKeyIndex1    = "<%=lgStrPrevKeyIndex%>"
				   .lgStrPrevKey         = "<%=lgStrPrevKey%>"					
                   .DBQueryOk     

	         End with
          Else
	          With Parent
                If Trim("<%=lgCurrentSpd%>") = "M" Then                   
                   .DBQueryOk        
                Else			
				   '.DBQueryOk2        
                End If  
	         End with

          End If   
 
    End Select    
       
</Script>	
