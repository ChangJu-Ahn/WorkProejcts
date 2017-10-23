<%@ LANGUAGE=VBSCript%>
<%Option Explicit%>

<!-- #Include file="../../inc/IncSvrMAin.asp"  -->
<!-- #Include file="../../inc/IncSvrDate.inc"  -->
<!-- #Include file="../../inc/IncSvrNumber.inc"  -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/incServerAdoDb.asp"  -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<!-- #Include file="../../inc/lgsvrvariables.inc"  -->


<%

    On Error Resume Next
    Err.Clear
    
    Dim lgTaxBizAreaCd
    Dim lgTaxBizAreaBodyCd
    Dim lgTaxBpCdUpFlag
    Dim lgTaxBpCdHistoryUpFlag
	Dim lgStrSQLCreate
    lgTaxBpCdUpFlag = false															 '☜: b_biz_partner에 insert 할것인지 update체크 
    lgTaxBpCdHistoryUpFlag = false
    Call HideStatusWnd                                                               '☜: Hide Processing message
	Call LoadBasisGlobalInf()
    Call LoadInfTB19029B("I","*","NOCOOKIE","MB")   
    Call LoadBNumericFormatB("I", "*","NOCOOKIE","MB") 

    'Dim lgErrorStatus, lgOpModeCRUD, lgObjConn
    '---------------------------------------Common-----------------------------------------------------------
    lgErrorStatus     = "NO"
    lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)
    
    lgTaxBizAreaCd		= FilterVar(UCase(Request("txtTaxBizAreaCd")), "''", "S")
    lgTaxBizAreaBodyCd	= FilterVar(UCase(Request("txtTaxBizAreaCd_Body")), "''", "S")
    


    Call SubOpenDB(lgObjConn)                                                        '☜: Make a DB Connection
	Call SubCreateCommandObject(lgObjComm)

    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '☜: Query
             Call SubBizQuery()
        Case CStr(UID_M0002)                                                         '☜: Save,Update
             Call SubBizSave()
        Case CStr(UID_M0003)                                                         '☜: Delete
             Call SubBizDelete()
    End Select

	Call SubCloseCommandObject(lgObjComm)
    Call SubCloseDB(lgObjConn)                                                       '☜: Close DB Connection


'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()
    Dim iKey1

    On Error Resume Next
    Err.Clear

    Call SubMakeSQLStatements("R", lgTaxBizAreaCd, "")

    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
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

    Else%>
<Script language="VBScript">
	With Parent.frm1 
		.txtTaxBizAreaNm.value = "<%=ConvSPChars(lgObjRs("TAX_BIZ_AREA_NM"))%>"
		.txtTaxBizAreaCd_Body.value = "<%=ConvSPChars(lgObjRs("TAX_BIZ_AREA_CD"))%>"
		.txtTaxBizAreaNm_Body.value = "<%=ConvSPChars(lgObjRs("TAX_BIZ_AREA_NM"))%>"
		.txtTaxBizAreaFullNm.value = "<%=ConvSPChars(lgObjRs("TAX_BIZ_AREA_FULL_NM"))%>"
		.txtTaxBizAreaEngNm.value = "<%=ConvSPChars(lgObjRs("TAX_BIZ_AREA_ENG_NM"))%>"
		.txtOwnRgstNo.value = "<%=ConvSPChars(lgObjRs("OWN_RGST_NO"))%>"
		.txtRepreNm.value = "<%=ConvSPChars(lgObjRs("REPRE_NM"))%>"
		.txtTaxOfficeCd.value = "<%=ConvSPChars(lgObjRs("TAX_OFFICE_CD"))%>"
		.txtTaxOfficeNm.value = "<%=ConvSPChars(lgObjRs("TAX_OFFICE_NM"))%>"
		.txtInd_class.value = "<%=ConvSPChars(lgObjRs("IND_CLASS"))%>"
		.txtInd_class_Nm.value = "<%=ConvSPChars(lgObjRs("IND_CLASS_NM"))%>"
		.txtInd_Type.value = "<%=ConvSPChars(lgObjRs("IND_TYPE"))%>"
		.txtInd_Type_Nm.value = "<%=ConvSPChars(lgObjRs("IND_TYPE_NM"))%>"
		.txtFaxNo.value = "<%=ConvSPChars(lgObjRs("FAX_NO"))%>"  
		.txtTelNo.value = "<%=ConvSPChars(lgObjRs("TEL_NO"))%>"  
		.txtAcctCharge.value = "<%=ConvSPChars(lgObjRs("ACCT_CHARGE"))%>"
		.txtZipCode.value = "<%=ConvSPChars(lgObjRs("ZIP_CODE"))%>"
		.txtIsCharge.value = "<%=ConvSPChars(lgObjRs("IS_CHARGE"))%>"
		.txtAddr1.value = "<%=ConvSPChars(lgObjRs("ADDR1"))%>"
		.txtAddr2.value = "<%=ConvSPChars(lgObjRs("ADDR2"))%>"
		.txtEng1Addr.value = "<%=ConvSPChars(lgObjRs("ADDR1_ENG"))%>"
		.txtEng2Addr.value = "<%=ConvSPChars(lgObjRs("ADDR2_ENG"))%>"
		.txtEng3Addr.value = "<%=ConvSPChars(lgObjRs("ADDR3_ENG"))%>"
	End With                                                       

</Script>
    <%End If
    Call SubCloseRs(lgObjRs)    
                                                    '☜ : Release RecordSSet
End Sub	
'============================================================================================================
' Name : SubBizQuery
' Desc : Date data 
'============================================================================================================
Sub SubBizSave()
    On Error Resume Next
    Err.Clear
    Dim lgIntFlgMode

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

    On Error Resume Next
    Err.Clear

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to update record
    '--------------------------------------------------------------------------------------------------------
    '//b_biz_area_cd
	Call SubMakeSQLStatements("CA", lgTaxBizAreaBodyCd, "")
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = True Then                    'If data not exists
		Call DisplayMsgBox("900020", vbInformation, "", "", I_MKSCRIPT)						'☜ : No data is found. 
        Call SetErrorStatus()
        Exit sub
	End if
	Call SubCloseRs(lgObjRs)																'☜ : Release RecordSSet

	'//a_vat
	Call SubMakeSQLStatements("CV", lgTaxBizAreaBodyCd, "")
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = True Then                    'If data not exists
		Call DisplayMsgBox("900020", vbInformation, "", "", I_MKSCRIPT)						'☜ : No data is found. 
        Call SetErrorStatus()
        Exit sub
	End if
	Call SubCloseRs(lgObjRs)                                                    '☜ : Release RecordSSet

	'//tax_biz_area_cd
    Call SubMakeSQLStatements("D", lgTaxBizAreaBodyCd, "")
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then                    'If data not exists
        '//message code 확인해야함 
        Call DisplayMsgBox("124210", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
        Call SetErrorStatus()

    Else
		Call SubDeleteTalbleCheck()
		lgStrSQL = "DELETE  b_tax_biz_area"
		lgStrSQL = lgStrSQL & " WHERE TAX_BIZ_AREA_CD = " & lgTaxBizAreaBodyCd                              ' 세금신고사업장 
		'---------- Developer Coding part (End  ) ---------------------------------------------------------------
		lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
		Call SubHandleError("SD",lgObjConn,lgObjRs,Err)

	End If

End Sub
'============================================================================================================
' Name : SubBizUpdateAfterDelete
' Desc : Update After Delete DB data When Delete Work is Succeed
'============================================================================================================

Sub SubBizUpdateAfterDelete()

    On Error Resume Next
    Err.Clear
	'---------- Developer Coding part (Start) ---------------------------------------------------------------
	'A developer must define field to update record
	'--------------------------------------------------------------------------------------------------------
	lgStrSQL = "UPDATE  B_BIZ_PARTNER"
	lgStrSQL = lgStrSQL & " SET " 
	lgStrSQL = lgStrSQL & " USAGE_FLAG = " & FilterVar("N", "''", "S") & " ,"
	lgStrSQL = lgStrSQL & " BP_TYPE = " & FilterVar("*", "''", "S") & " "
	lgStrSQL = lgStrSQL & " WHERE BP_RGST_NO = " & FilterVar(Trim(Request("txtOwnRgstNo")), "''","S") 
	lgStrSQL = lgStrSQL & " AND	  BP_TYPE = " & FilterVar("T", "''", "S") & " "
	'---------- Developer Coding part (End  ) ---------------------------------------------------------------
	'Response.write VBTAB & lgStrSQL & VBTAB

	lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords

	Call SubHandleError("DU",lgObjConn,lgObjRs,Err)
	Call SubCloseRs(lgObjRs)
	'//Call ServerMesgBox("vvv", vbCritical, I_MKSCRIPT)			'⊙:                                          '☜ : Release RecordSSet

End Sub



'============================================================================================================
' Name : SubBizSaveSingleCreate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveSingleCreate()
    On Error Resume Next
    Err.Clear
	Dim strTax_biz_area_cd
	Dim strTax_biz_area_nm
	Dim strTax_biz_area_full_nm
	Dim strTax_biz_area_eng_nm
	Dim strCo_cd
	Dim strTax_office_cd
	Dim strOwn_rgst_no
	Dim strRepre_nm
	Dim strInd_class
	Dim strInd_type
	Dim strFax_no
	Dim strTel_no
	Dim strZip_code
	Dim strAddr1
	Dim strAddr2
	Dim strEng1_addr
	Dim strEng2_addr
	Dim strEng3_addr
	Dim strAcct_charge
	Dim strIs_charge
	Dim strInsrt_user_id
	Dim strInsrt_dt
	Dim strUpdt_user_id
	Dim strUpdt_dt
	


	strTax_biz_area_cd		 = FilterVar(UCase(Trim(Request("txtTaxBizAreaCd_Body"))), "''","S")
	strTax_biz_area_nm       = FilterVar(UCase(Trim(Request("txtTaxBizAreaNm_Body"))), "''","S")
	strTax_biz_area_full_nm  = FilterVar(UCase(Trim(Request("txtTaxBizAreaFullNm"))), "''","S")
	strTax_biz_area_eng_nm   = FilterVar(UCase(Trim(Request("txtTaxBizAreaEngNm"))), "''","S")
	strTax_office_cd         = FilterVar(UCase(Trim(Request("txtTaxOfficeCd"))), "''","S")
	strOwn_rgst_no           = FilterVar(Trim(Request("txtOwnRgstNo")), "''","S")
	strRepre_nm              = FilterVar(Trim(Request("txtRepreNm")), "''","S")
	strInd_class             = FilterVar(UCase(Trim(Request("txtInd_class"))), "''","S")
	strInd_type              = FilterVar(UCase(Trim(Request("txtInd_Type"))), "''","S")
	strFax_no                = FilterVar(Trim(Request("txtFaxNo")), "''","S")
	strTel_no                = FilterVar(Trim(Request("txtTelNo")), "''","S")
	strZip_code              = FilterVar(Trim(Request("txtZipCode")), "''","S")
	strAddr1                  = Trim(Request("txtAddr1"))
	strAddr2                  = Trim(Request("txtAddr2"))
	strEng1_addr              = Trim(Request("txtEng1Addr"))
	strEng2_addr              = Trim(Request("txtEng2Addr"))
	strEng3_addr              = Trim(Request("txtEng3Addr"))
	strAcct_charge           = FilterVar(Trim(Request("txtAcctCharge")), "''","S")
	strIs_charge             = FilterVar(Trim(Request("txtIsCharge")), "''","S")
	strInsrt_user_id         = FilterVar(gUsrId, "''","S")
	strInsrt_dt              = FilterVar(GetSvrDateTime, "''","S")
	strUpdt_user_id          = FilterVar(gUsrId, "''","S")
	strUpdt_dt               = FilterVar(GetSvrDateTime, "''","S")
	
	'//co_cd
	Call SubMakeSQLStatements("CC", "", "")
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = True Then          'If data not exists
		strCo_cd = FilterVar(lgObjRs("CO_Cd"), "''","S")
	Else
		strCo_cd = ""
	End if
	Call SubCloseRs(lgObjRs)                                                    '☜ : Release RecordSSet


	'//tax_biz_area_cd
	Call SubMakeSQLStatements("CT", strTax_biz_area_cd, "")
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = True Then                    'If data not exists
		Call DisplayMsgBox("124902", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
        Call SetErrorStatus()
        Exit sub
	End if
	Call SubCloseRs(lgObjRs)                                                    '☜ : Release RecordSSet

	'//bp_cd (For b_biz_partner)
	Call SubMakeSQLStatements("CP", strTax_biz_area_cd, "")
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = True Then        'If data not exists
		Call SubMakeSQLStatements("CD", strTax_biz_area_cd, strOwn_rgst_no)
		If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = True Then        'If data not exists
			lgTaxBpCdUpFlag = True												'☜ : b_biz_partner table not insert but upate 
		Else
			Call DisplayMsgBox("124911", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
			Call SetErrorStatus()
			Exit sub
		End If
	End if
	Call SubCloseRs(lgObjRs)        

	'//bp_cd,valid_from_dt (For b_biz_partner_history)
	Call SubMakeSQLStatements("CH", strTax_biz_area_cd, FilterVar(GetSvrDate, "''","S"))
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = True Then        'If data not exists
			lgTaxBpCdHistoryUpFlag = True									'☜ : b_biz_partner_history table not insert but upate 
	End if
	Call SubCloseRs(lgObjRs)

	'//tax_office_cd
	Call SubMakeSQLStatements("CO", strTax_office_cd, "")
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then                    'If data not exists
		Call DisplayMsgBox("126900", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
        Call SetErrorStatus()
        Exit sub
	End if
	Call SubCloseRs(lgObjRs)                                                    '☜ : Release RecordSSet

	'//ind_class : 업태 b9003
	Call SubMakeSQLStatements("CB", strInd_class, "" & FilterVar("B9003", "''", "S") & " ")
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then                    'If data not exists
		Call DisplayMsgBox("126128", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
        Call SetErrorStatus()
        Exit sub
	End if
	Call SubCloseRs(lgObjRs)                                                    '☜ : Release RecordSSet

	'//ind_type : 업종 b9002
	Call SubMakeSQLStatements("CB", strInd_type, "" & FilterVar("B9002", "''", "S") & " ")
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then          'If data not exists
		Call DisplayMsgBox("126127", vbInformation, "", "", I_MKSCRIPT)			'☜ : No data is found. 
        Call SetErrorStatus()
        Exit sub
	End if
	Call SubCloseRs(lgObjRs)                                                    '☜ : Release RecordSSet

	If Trim(Request("txtZipCode")) <> "" Then
		'//zipCode
		Call SubMakeSQLStatements("CZ", strZip_code, "")
		If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then            'If data not exists
			Call DisplayMsgBox("800016", vbInformation, "", "", I_MKSCRIPT)			'☜ : No data is found. 
		    Call SetErrorStatus()
		    Exit sub
		End if
		Call SubCloseRs(lgObjRs)                                                    '☜ : Release RecordSSet
	End IF

	'//b_tax_biz_area
	lgStrSQL =  		  " Insert into B_Tax_Biz_Area "
	lgStrSQL = lgStrSQL & " ( TAX_BIZ_AREA_CD ,"
	lgStrSQL = lgStrSQL & "TAX_BIZ_AREA_NM     , "
	lgStrSQL = lgStrSQL & "TAX_BIZ_AREA_FULL_NM, "
	lgStrSQL = lgStrSQL & "TAX_BIZ_AREA_ENG_NM , "
	lgStrSQL = lgStrSQL & "CO_CD               , "
	lgStrSQL = lgStrSQL & "TAX_OFFICE_CD       , "
	lgStrSQL = lgStrSQL & "OWN_RGST_NO         , "
	lgStrSQL = lgStrSQL & "REPRE_NM            , "
	lgStrSQL = lgStrSQL & "IND_CLASS           , "
	lgStrSQL = lgStrSQL & "IND_TYPE            , "
	lgStrSQL = lgStrSQL & "FAX_NO              , "
	lgStrSQL = lgStrSQL & "TEL_NO              , "
	lgStrSQL = lgStrSQL & "ZIP_CODE            , "
	lgStrSQL = lgStrSQL & "ADDR                , "
	lgStrSQL = lgStrSQL & "ADDR1                , "
	lgStrSQL = lgStrSQL & "ADDR2                , "
	lgStrSQL = lgStrSQL & "ENG_ADDR            , "
	lgStrSQL = lgStrSQL & "ADDR1_ENG            , "
	lgStrSQL = lgStrSQL & "ADDR2_ENG            , "
	lgStrSQL = lgStrSQL & "ADDR3_ENG            , "
	lgStrSQL = lgStrSQL & "ACCT_CHARGE         , "
	lgStrSQL = lgStrSQL & "IS_CHARGE           , "
	lgStrSQL = lgStrSQL & "INSRT_USER_ID       , "
	lgStrSQL = lgStrSQL & "INSRT_DT            , "
	lgStrSQL = lgStrSQL & "UPDT_USER_ID        , "
	lgStrSQL = lgStrSQL & "UPDT_DT              "
	lgStrSQL = lgStrSQL & " )"  
	lgStrSQL = lgStrSQL & " values ("
	lgStrSQL = lgStrSQL & 		strTax_biz_area_cd			& ","
	lgStrSQL = lgStrSQL & 	 	strTax_biz_area_nm          & ","
	lgStrSQL = lgStrSQL & 	 	strTax_biz_area_full_nm     & ","
	lgStrSQL = lgStrSQL & 	 	strTax_biz_area_eng_nm      & ","
	lgStrSQL = lgStrSQL & 	 	strCo_cd                    & ","
	lgStrSQL = lgStrSQL & 	 	strTax_office_cd            & ","
	lgStrSQL = lgStrSQL & 	 	strOwn_rgst_no              & ","
	lgStrSQL = lgStrSQL & 	 	strRepre_nm                 & ","
	lgStrSQL = lgStrSQL & 	 	strInd_class                & ","
	lgStrSQL = lgStrSQL & 	 	strInd_type                 & ","
	lgStrSQL = lgStrSQL & 	 	strFax_no                   & ","
	lgStrSQL = lgStrSQL & 	 	strTel_no                   & ","
	lgStrSQL = lgStrSQL & 	 	strZip_code                 & ","
	lgStrSQL = lgStrSQL & 	 	" Convert(Nvarchar(128), " & FilterVar(strAddr1 & " " & Trim(strAddr2) , "''", "S") & "),"
	lgStrSQL = lgStrSQL & 	 	" Convert(NVarchar(100), " & FilterVar(strAddr1             , "''", "S") & "),"
	lgStrSQL = lgStrSQL & 	 	" Convert(NVarchar(100), " & FilterVar(strAddr2             , "''", "S") & "),"
	lgStrSQL = lgStrSQL & 	 	" Convert(NVarchar(128), " & FilterVar(strEng1_addr & " " & Trim(strEng2_addr) & " " & Trim(strEng3_addr), "''", "S") & "),"
	lgStrSQL = lgStrSQL & 	 	" Convert(NVarchar(50), " & FilterVar(strEng1_addr         , "''", "S") & "),"
	lgStrSQL = lgStrSQL & 	 	" Convert(NVarchar(50), " & FilterVar(strEng2_addr         , "''", "S") & "),"
	lgStrSQL = lgStrSQL & 	 	" Convert(NVarchar(50), " & FilterVar(strEng3_addr         , "''", "S") & "),"
	lgStrSQL = lgStrSQL & 	 	strAcct_charge              & ","
	lgStrSQL = lgStrSQL & 	 	strIs_charge                & ","
	lgStrSQL = lgStrSQL & 	 	strInsrt_user_id            & ","
	lgStrSQL = lgStrSQL & 	 	strInsrt_dt                 & ","
	lgStrSQL = lgStrSQL & 	 	strUpdt_user_id             & ","
	lgStrSQL = lgStrSQL & 	 	strUpdt_dt                  
	lgStrSQL = lgStrSQL & " )"
	lgStrSQLCreate = ""
	lgStrSQLCreate = lgStrSQL

'	Response.Write vbtab & "B_Tax_Biz_Area = >>> " & lgStrSQL & vbtab
	'lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
	Call SubHandleError("SC",lgObjConn,lgObjRs,Err)

End Sub
'============================================================================================================
' Name : SubBizPartnerSaveSingleCreate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizPartnerSaveSingleCreate()
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
	Dim strTax_biz_area_cd
	Dim strTax_biz_area_nm
	Dim strTax_biz_area_full_nm
	Dim strTax_biz_area_eng_nm
	Dim strCo_cd
	Dim strTax_office_cd
	Dim strOwn_rgst_no
	Dim strRepre_nm
	Dim strInd_class
	Dim strInd_type
	Dim strFax_no
	Dim strTel_no
	Dim strZip_code
	Dim strAddr1
	Dim strAddr2
	Dim strEng1_addr
	Dim strEng2_addr
	Dim strEng3_addr
	Dim strAcct_charge
	Dim strIs_charge
	Dim strInsrt_user_id
	Dim strInsrt_dt
	Dim strUpdt_user_id
	Dim strUpdt_dt
		
	strTax_biz_area_cd		 = FilterVar(UCase(Trim(Request("txtTaxBizAreaCd_Body"))), "''","S")
	strTax_biz_area_nm       = FilterVar(UCase(Trim(Request("txtTaxBizAreaNm_Body"))), "''","S")
	strTax_biz_area_full_nm  = FilterVar(UCase(Trim(Request("txtTaxBizAreaFullNm"))), "''","S")
	strTax_biz_area_eng_nm   = FilterVar(UCase(Trim(Request("txtTaxBizAreaEngNm"))), "''","S")
	strTax_office_cd         = FilterVar(UCase(Trim(Request("txtTaxOfficeCd"))), "''","S")
	strOwn_rgst_no           = FilterVar(Trim(Request("txtOwnRgstNo")), "''","S")
	strRepre_nm              = FilterVar(Trim(Request("txtRepreNm")), "''","S")
	strInd_class             = "Convert(nVarchar(5), " & FilterVar(UCase(Trim(Request("txtInd_class"))), "''","S")  & ")"
	strInd_type              = "Convert(nVarchar(5), " & FilterVar(UCase(Trim(Request("txtInd_Type"))), "''","S")  & ")"
	strFax_no                = FilterVar(Trim(Request("txtFaxNo")), "''","S")
	strTel_no                = FilterVar(Trim(Request("txtTelNo")), "''","S")
	strZip_code              = FilterVar(Trim(Request("txtZipCode")), "''","S")
	strAddr1                  = FilterVar(Trim(Request("txtAddr1")), "","SNM")
	strAddr2                  = FilterVar(Trim(Request("txtAddr2")), "","SNM")
	strEng1_addr              = FilterVar(Trim(Request("txtEng1Addr")), "","SNM")
	strEng2_addr              = FilterVar(Trim(Request("txtEng2Addr")), "","SNM")
	strEng3_addr              = FilterVar(Trim(Request("txtEng3Addr")), "","SNM")
	strAcct_charge           = FilterVar(Trim(Request("txtAcctCharge")), "''","S")
	strIs_charge             = FilterVar(Trim(Request("txtIsCharge")), "''","S")
	strInsrt_user_id         = FilterVar(gUsrId, "''","S")
	strInsrt_dt              = FilterVar(GetSvrDateTime, "''","S")
	strUpdt_user_id          = FilterVar(gUsrId, "''","S")
	strUpdt_dt               = FilterVar(GetSvrDateTime, "''","S")
	'//co_cd

	'//B_Biz_Partner
	If lgTaxBpCdUpFlag = False  Then
		lgStrSQL =  		  " Insert into B_Biz_Partner "
		'//columnList For Insert
		lgStrSQL = lgStrSQL & " ( "
		lgStrSQL = lgStrSQL & " BP_CD           , "
		lgStrSQL = lgStrSQL & " BP_TYPE         , "
		lgStrSQL = lgStrSQL & " BP_GROUP        , "
		lgStrSQL = lgStrSQL & " BP_RGST_NO      , "
		lgStrSQL = lgStrSQL & " BP_FULL_NM      , "
		lgStrSQL = lgStrSQL & " BP_NM           , "
		lgStrSQL = lgStrSQL & " BP_ENG_NM       , "
		lgStrSQL = lgStrSQL & " REPRE_NM        , "
		lgStrSQL = lgStrSQL & " REPRE_RGST_NO   , "
		lgStrSQL = lgStrSQL & " FND_DT          , "
		lgStrSQL = lgStrSQL & " ZIP_CD          , "
		lgStrSQL = lgStrSQL & " ADDR1           , "
		lgStrSQL = lgStrSQL & " ADDR2           , "
		lgStrSQL = lgStrSQL & " ADDR1_ENG       , "
		lgStrSQL = lgStrSQL & " ADDR2_ENG       , "
		lgStrSQL = lgStrSQL & " ADDR3_ENG       , "
		lgStrSQL = lgStrSQL & " IND_TYPE        , "
		lgStrSQL = lgStrSQL & " IND_CLASS       , "
		lgStrSQL = lgStrSQL & " CONTRY_CD       , "
		lgStrSQL = lgStrSQL & " PROVINCE_CD     , "
		lgStrSQL = lgStrSQL & " CURRENCY     	, "
		lgStrSQL = lgStrSQL & " TEL_NO1         , "
		lgStrSQL = lgStrSQL & " TEL_NO2         , "
		lgStrSQL = lgStrSQL & " FAX_NO          , "
		lgStrSQL = lgStrSQL & " HOME_URL        , "
		lgStrSQL = lgStrSQL & " USAGE_FLAG      , "
		lgStrSQL = lgStrSQL & " BP_CONTACT_PT   , "
		lgStrSQL = lgStrSQL & " BIZ_PRSN        , "
		lgStrSQL = lgStrSQL & " BIZ_GRP         , "
		lgStrSQL = lgStrSQL & " BIZ_ORG         , "
		lgStrSQL = lgStrSQL & " DEAL_TYPE       , "
		lgStrSQL = lgStrSQL & " PAY_METH        , "
		lgStrSQL = lgStrSQL & " PAY_DUR         , "
		lgStrSQL = lgStrSQL & " PAY_DAY         , "
		lgStrSQL = lgStrSQL & " PAY_TYPE        , "
		lgStrSQL = lgStrSQL & " CREDIT_MGMT_FLAG, "
		lgStrSQL = lgStrSQL & " CREDIT_ROT_DAY  , "
		lgStrSQL = lgStrSQL & " VAT_INC_FLAG    , "
		lgStrSQL = lgStrSQL & " VAT_TYPE        , "
		lgStrSQL = lgStrSQL & " VAT_RATE        , "
		lgStrSQL = lgStrSQL & " VAT_CALC_TYPE   , "
		lgStrSQL = lgStrSQL & " TRANS_METH      , "
		lgStrSQL = lgStrSQL & " TRANS_LT        , "
		lgStrSQL = lgStrSQL & " SALE_AMT        , "
		lgStrSQL = lgStrSQL & " CAPITAL_AMT     , "
		lgStrSQL = lgStrSQL & " BP_GRADE        , "
		lgStrSQL = lgStrSQL & " COMM_RATE       , "
		lgStrSQL = lgStrSQL & " INSRT_USER_ID   , "
		lgStrSQL = lgStrSQL & " INSRT_DT        , "
		lgStrSQL = lgStrSQL & " UPDT_USER_ID    , "
		lgStrSQL = lgStrSQL & " UPDT_DT         , "
		lgStrSQL = lgStrSQL & " emp_cnt         , "
		lgStrSQL = lgStrSQL & " TAX_BIZ_AREA    , "
		lgStrSQL = lgStrSQL & " PAY_MONTH2      , "
		lgStrSQL = lgStrSQL & " PAY_DAY2        , "
		lgStrSQL = lgStrSQL & " PAY_MONTH3      , "
		lgStrSQL = lgStrSQL & " PAY_DAY3        , "
		lgStrSQL = lgStrSQL & " CLOSE_DAY1_SALES, "
		lgStrSQL = lgStrSQL & " PAY_MONTH1_SALES, "
		lgStrSQL = lgStrSQL & " PAY_DAY1_SALES  , "
		lgStrSQL = lgStrSQL & " CLOSE_DAY2_SALES, "
		lgStrSQL = lgStrSQL & " PAY_MONTH2_SALES, "
		lgStrSQL = lgStrSQL & " PAY_DAY2_SALES  , "
		lgStrSQL = lgStrSQL & " CLOSE_DAY3_SALES, "
		lgStrSQL = lgStrSQL & " PAY_MONTH3_SALES, "
		lgStrSQL = lgStrSQL & " PAY_DAY3_SALES  , "
		lgStrSQL = lgStrSQL & " PAY_TYPE_OUT     "
		lgStrSQL = lgStrSQL & " )"
		'//values For Insert
		lgStrSQL = lgStrSQL & " values ("
		lgStrSQL = lgStrSQL & 	strTax_biz_area_cd & ","       			'// varchar 			BP_CD
		lgStrSQL = lgStrSQL & " " & FilterVar("T", "''", "S") & " , "                                  '// varchar             BP_TYPE
		lgStrSQL = lgStrSQL & " '', "                           		'// varchar             BP_GROUP
		lgStrSQL = lgStrSQL & 	strOwn_rgst_no & ","       				'// varchar             BP_RGST_NO
		lgStrSQL = lgStrSQL & 	strTax_biz_area_full_nm & ","       	'// varchar             BP_FULL_NM
		lgStrSQL = lgStrSQL & 	strTax_biz_area_nm & ","       			'// varchar             BP_NM
		lgStrSQL = lgStrSQL & 	strTax_biz_area_eng_nm & ","       		'// varchar             BP_ENG_NM
		lgStrSQL = lgStrSQL & 	strRepre_nm & ","       				'// varchar             REPRE_NM
		lgStrSQL = lgStrSQL & " '', "                                   '// varchar             REPRE_RGST_NO
		lgStrSQL = lgStrSQL & " '', "                                   '// datetime            FND_DT
		lgStrSQL = lgStrSQL & 	strZip_code & ","       				'// varchar             ZIP_CD
		lgStrSQL = lgStrSQL & 	" Convert(nVarchar(100), " & FilterVar(strAddr1, "''", "S") & "),"       					'// varchar             ADDR1
		lgStrSQL = lgStrSQL & 	" Convert(nVarchar(100), " & FilterVar(strAddr2, "''", "S") & "),"       					'// varchar             ADDR2
		lgStrSQL = lgStrSQL &	" Convert(nVarchar(50), " & FilterVar(strEng1_addr, "''", "S") & "), "                      '// varchar             ADDR1_ENG
		lgStrSQL = lgStrSQL &	" Convert(nVarchar(50), " & FilterVar(strEng2_addr, "''", "S") & "), "                      '// varchar             ADDR2_ENG
		lgStrSQL = lgStrSQL &	" Convert(nVarchar(50), " & FilterVar(strEng3_addr, "''", "S") & "), "                      '// varchar             ADDR3_ENG
		lgStrSQL = lgStrSQL & 	strInd_type  & ","       				'// varchar             IND_TYPE
		lgStrSQL = lgStrSQL & 	strInd_class & ","       				'// varchar             IND_CLASS
		lgStrSQL = lgStrSQL & FilterVar(gCountry, "''","S") & "," 	'// char                CONTRY_CD
		lgStrSQL = lgStrSQL & " '', "                                   '// varchar             PROVINCE_CD
		lgStrSQL = lgStrSQL & FilterVar(gCurrency, "''","S") & ","             
		lgStrSQL = lgStrSQL & 	strTel_no & ","       					'// varchar             TEL_NO1
		lgStrSQL = lgStrSQL & " '', "                                   '// varchar             TEL_NO2
		lgStrSQL = lgStrSQL & 	strFax_no & ","       					'// varchar             FAX_NO
		lgStrSQL = lgStrSQL & " '', "                                   '// varchar             HOME_URL
		lgStrSQL = lgStrSQL & " " & FilterVar("Y", "''", "S") & " , "                                  '// char                USAGE_FLAG
		lgStrSQL = lgStrSQL & " '', "                                   '// varchar             BP_CONTACT_PT
		lgStrSQL = lgStrSQL & " '', "                                   '// varchar             BIZ_PRSN
		lgStrSQL = lgStrSQL & " '', "                                   '// varchar             BIZ_GRP
		lgStrSQL = lgStrSQL & " '', "                                   '// varchar             BIZ_ORG
		lgStrSQL = lgStrSQL & " '', "                                   '// varchar             DEAL_TYPE
		lgStrSQL = lgStrSQL & " '', "                                   '// varchar             PAY_METH
		lgStrSQL = lgStrSQL & " 0, "                                    '// numeric             PAY_DUR
		lgStrSQL = lgStrSQL & " " & FilterVar("0", "''", "S") & " , "                                  '// varchar             PAY_DAY
		lgStrSQL = lgStrSQL & " '', "                                   '// varchar             PAY_TYPE
		lgStrSQL = lgStrSQL & " " & FilterVar("N", "''", "S") & " , "                    				'// char                CREDIT_MGMT_FLAG 
		lgStrSQL = lgStrSQL & " 0, "				                    '// numeric            	CREDIT_ROT_DAY
		lgStrSQL = lgStrSQL & " " & FilterVar("2", "''", "S") & " , "                     				'// char                VAT_INC_FLAG
		lgStrSQL = lgStrSQL & " '', "                     				'// varchar             VAT_TYPE
		lgStrSQL = lgStrSQL & " 0, "                     				'// numeric             VAT_RATE
		lgStrSQL = lgStrSQL & " " & FilterVar("1", "''", "S") & " , "                     				'// char                VAT_CALC_TYPE
		lgStrSQL = lgStrSQL & " '', "                     				'// varchar             TRANS_METH
		lgStrSQL = lgStrSQL & " 0, "                     				'// numeric             TRANS_LT
		lgStrSQL = lgStrSQL & " 0, "                     				'// numeric             SALE_AMT
		lgStrSQL = lgStrSQL & " 0, "                     				'// numeric             CAPITAL_AMT
		lgStrSQL = lgStrSQL & " '' , "                     				'// varchar             BP_GRADE
		lgStrSQL = lgStrSQL & " 0, "                     				'// numeric             COMM_RATE
		lgStrSQL = lgStrSQL & 	strInsrt_user_id  & ","                 '//  			         INSRT_USER_ID
		lgStrSQL = lgStrSQL & 	strInsrt_dt  & ","                      '//      				INSRT_DT
		lgStrSQL = lgStrSQL & 	strUpdt_user_id & ","               	'// 		            UPDT_USER_ID
		lgStrSQL = lgStrSQL & 	strUpdt_dt & ","       					'// datetime            UPDT_DT
		lgStrSQL = lgStrSQL & " 0, "                     				'// numeric             emp_cnt
		lgStrSQL = lgStrSQL &	strTax_biz_area_cd & ","		        '// 
		lgStrSQL = lgStrSQL & " 0, "                     				'// numeric             PAY_MONTH2
		lgStrSQL = lgStrSQL & " '', "                    				'// varchar             PAY_DAY2
		lgStrSQL = lgStrSQL & " 0, "                    				'// numeric             PAY_MONTH3
		lgStrSQL = lgStrSQL & " '', "                     				'// varchar             PAY_DAY3
		lgStrSQL = lgStrSQL & " 0, "                     				'// numeric             CLOSE_DAY1_SALES
		lgStrSQL = lgStrSQL & " 0, "                     				'// numeric             PAY_MONTH1_SALES
		lgStrSQL = lgStrSQL & " '', "                     				'// varchar             PAY_DAY1_SALES
		lgStrSQL = lgStrSQL & " 0,"                     				'// numeric             CLOSE_DAY2_SALES
		lgStrSQL = lgStrSQL & " 0,"                     				'// numeric             PAY_MONTH2_SALES
		lgStrSQL = lgStrSQL & " '', "                     				'// varchar             PAY_DAY2_SALES
		lgStrSQL = lgStrSQL & " 0,"                     				'// numeric             CLOSE_DAY3_SALES
		lgStrSQL = lgStrSQL & " 0,"                     				'// numeric             PAY_MONTH3_SALES
		lgStrSQL = lgStrSQL & " '', "                     				'// varchar             PAY_DAY3_SALES
		lgStrSQL = lgStrSQL & " ''"                      				'// varchar             PAY_TYPE_OUT
		lgStrSQL = lgStrSQL & " ) "
	Else
		lgStrSQL =  " 			Update B_Biz_Partner "
		lgStrSQL = lgStrSQL & " Set					 "	
		lgStrSQL = lgStrSQL & " BP_CD  			 = " & 	strTax_biz_area_cd & ","
		lgStrSQL = lgStrSQL & " BP_TYPE          =  		" & FilterVar("T", "''", "S") & " , "
		lgStrSQL = lgStrSQL & " BP_RGST_NO       = " & 	strOwn_rgst_no & ","
		lgStrSQL = lgStrSQL & " BP_FULL_NM       = " & 	strTax_biz_area_full_nm & ","
		lgStrSQL = lgStrSQL & " BP_NM            = " & 	strTax_biz_area_nm & ","
		lgStrSQL = lgStrSQL & " BP_ENG_NM        = " & 	strTax_biz_area_eng_nm & ","
		lgStrSQL = lgStrSQL & " REPRE_NM         = " & 	strRepre_nm & ","
		lgStrSQL = lgStrSQL & " ZIP_CD           = " & 	strZip_code & ","
		lgStrSQL = lgStrSQL & " ADDR1            =  Convert(nVarchar(100), " & FilterVar(strAddr1, "''", "S") & "),"
		lgStrSQL = lgStrSQL & " ADDR2            =  Convert(nVarchar(100), " & FilterVar(strAddr2, "''", "S") & "),"
		lgStrSQL = lgStrSQL & " ADDR1_ENG        =  Convert(nVarchar(50), " & FilterVar( strEng1_addr, "''", "S") & "),"
		lgStrSQL = lgStrSQL & " ADDR2_ENG        =  Convert(nVarchar(50), " & FilterVar( strEng2_addr, "''", "S") & "),"
		lgStrSQL = lgStrSQL & " ADDR3_ENG        =  Convert(nVarchar(50), " & FilterVar( strEng3_addr, "''", "S") & "),"
		lgStrSQL = lgStrSQL & " IND_TYPE         = " & 	strInd_type  & ","
		lgStrSQL = lgStrSQL & " IND_CLASS        = " & 	strInd_class & ","
		lgStrSQL = lgStrSQL & " USAGE_FLAG       =  		" & FilterVar("Y", "''", "S") & " , "
		lgStrSQL = lgStrSQL & " CONTRY_CD        = " & 	FilterVar(gCountry, "''","S") & ","
		lgStrSQL = lgStrSQL & " CURRENCY     	 = " & 	FilterVar(gCurrency, "''","S") & ","
		lgStrSQL = lgStrSQL & " TEL_NO1          = " & 	strTel_no & ","
		lgStrSQL = lgStrSQL & " FAX_NO           = " & 	strFax_no & ","
		lgStrSQL = lgStrSQL & " UPDT_USER_ID     = " & 	strUpdt_user_id & ","
		lgStrSQL = lgStrSQL & " UPDT_DT          = " & 	strUpdt_dt & ","
		lgStrSQL = lgStrSQL & " TAX_BIZ_AREA     = " & 	strTax_biz_area_cd
		lgStrSQL = lgStrSQL & " WHERE 	BP_CD  			 = " & strTax_biz_area_cd
		lgStrSQL = lgStrSQL & " AND 	USAGE_FLAG       =  " & FilterVar("N", "''", "S") & "   "
		lgStrSQL = lgStrSQL & " AND		BP_TYPE 		 = " & FilterVar("*", "''", "S") & " "	  
	End If	
	'Response.Write vbtab & "B_Biz_Partner = >>> " & lgStrSQL & vbtab
	
	lgStrSQLCreate = lgStrSQLCreate & " " & lgStrSQL
	'lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
	'Response.end
	Call SubHandleError("SCB",lgObjConn,lgObjRs,Err)

End Sub
'============================================================================================================
' Name : SubBizPartnerHistorySaveSingleCreate
' Desc : Query Data from Db
'============================================================================================================

Sub SubBizPartnerHistorySaveSingleCreate()
    On Error Resume Next
    Err.Clear
	Dim strTax_biz_area_cd
	Dim strTax_biz_area_nm
	Dim strTax_biz_area_full_nm
	Dim strTax_biz_area_eng_nm
	Dim strCo_cd
	Dim strTax_office_cd
	Dim strOwn_rgst_no
	Dim strRepre_nm
	Dim strInd_class
	Dim strInd_type
	Dim strFax_no
	Dim strTel_no
	Dim strZip_code
	Dim strAddr,strAddr1,strAddr2,strEng1_addr,strEng2_addr,strEng3_addr
	Dim strEng_addr
	Dim strAcct_charge
	Dim strIs_charge
	Dim strInsrt_user_id
	Dim strInsrt_dt
	Dim strUpdt_user_id
	Dim strUpdt_dt
	Dim strValidFromDt

	strTax_biz_area_cd		 = FilterVar(UCase(Trim(Request("txtTaxBizAreaCd_Body"))), "''","S")
	strValidFromDt			 = FilterVar(GetSvrDate, "''","S")
	strTax_biz_area_nm       = FilterVar(UCase(Trim(Request("txtTaxBizAreaNm_Body"))), "''","S")
	strTax_biz_area_full_nm  = FilterVar(UCase(Trim(Request("txtTaxBizAreaFullNm"))), "''","S")
	strTax_biz_area_eng_nm   = FilterVar(UCase(Trim(Request("txtTaxBizAreaEngNm"))), "''","S")
	strTax_office_cd         = FilterVar(UCase(Trim(Request("txtTaxOfficeCd"))), "''","S")
	strOwn_rgst_no           = FilterVar(Trim(Request("txtOwnRgstNo")), "''","S")
	strRepre_nm              = FilterVar(Trim(Request("txtRepreNm")), "''","S")
	strInd_class             = "Convert(nVarchar(5), " & FilterVar(UCase(Trim(Request("txtInd_class"))), "''","S") & ")"
	strInd_type              = "Convert(nVarchar(5), " & FilterVar(UCase(Trim(Request("txtInd_Type"))), "''","S")  & ")"
	strFax_no                = FilterVar(Trim(Request("txtFaxNo")), "''","S")
	strTel_no                = FilterVar(Trim(Request("txtTelNo")), "''","S")
	strZip_code              = FilterVar(Trim(Request("txtZipCode")), "''","S")
	strAddr1                 = FilterVar(Trim(Request("txtAddr1")), "","SNM")
	strAddr2                 = FilterVar(Trim(Request("txtAddr2")), "","SNM")
	strEng1_addr             = FilterVar(Trim(Request("txtEng1Addr")), "","SNM")
	strEng2_addr             = FilterVar(Trim(Request("txtEng2Addr")), "","SNM")
	strEng3_addr             = FilterVar(Trim(Request("txtEng3Addr")), "","SNM")
	strAcct_charge           = FilterVar(Trim(Request("txtAcctCharge")), "''","S")
	strIs_charge             = FilterVar(Trim(Request("txtIsCharge")), "''","S")
	strInsrt_user_id         = FilterVar(gUsrId, "''","S")
	strInsrt_dt              = FilterVar(GetSvrDateTime, "''","S")
	strUpdt_user_id          = FilterVar(gUsrId, "''","S")
	strUpdt_dt               = FilterVar(GetSvrDateTime, "''","S")
	'//co_cd


	If lgTaxBpCdHistoryUpFlag = false Then
		'//b_tax_biz_area
		lgStrSQL =  		  " Insert into B_Biz_Partner_History"
		lgStrSQL = lgStrSQL & " ( "
		lgStrSQL = lgStrSQL & " BP_CD			  , "
		lgStrSQL = lgStrSQL & " VALID_FROM_DT     , "
		lgStrSQL = lgStrSQL & " BP_RGST_NO        , "
		lgStrSQL = lgStrSQL & " BP_FULL_NM        , "
		lgStrSQL = lgStrSQL & " BP_NM             , "
		lgStrSQL = lgStrSQL & " BP_ENG_NM         , "
		lgStrSQL = lgStrSQL & " REPRE_NM          , "
		lgStrSQL = lgStrSQL & " REPRE_RGST_NO     , "
		lgStrSQL = lgStrSQL & " IND_TYPE          , "
		lgStrSQL = lgStrSQL & " IND_CLASS         , "
		lgStrSQL = lgStrSQL & " CHG_REASON        , "
		lgStrSQL = lgStrSQL & " ZIP_CD          , "
		lgStrSQL = lgStrSQL & " ADDR1           , "
		lgStrSQL = lgStrSQL & " ADDR2           , "
		lgStrSQL = lgStrSQL & " ADDR1_ENG       , "
		lgStrSQL = lgStrSQL & " ADDR2_ENG       , "
		lgStrSQL = lgStrSQL & " ADDR3_ENG       , "
		lgStrSQL = lgStrSQL & " INSRT_USER_ID     , "
		lgStrSQL = lgStrSQL & " INSRT_DT          , "
		lgStrSQL = lgStrSQL & " UPDT_USER_ID      , "
		lgStrSQL = lgStrSQL & " UPDT_DT           , "
		lgStrSQL = lgStrSQL & " EXT1_QTY          , "
		lgStrSQL = lgStrSQL & " EXT2_QTY          , "
		lgStrSQL = lgStrSQL & " EXT1_AMT          , "
		lgStrSQL = lgStrSQL & " EXT2_AMT          , "
		lgStrSQL = lgStrSQL & " EXT1_CD           , "
		lgStrSQL = lgStrSQL & " EXT2_CD            "
		lgStrSQL = lgStrSQL & " )"
		lgStrSQL = lgStrSQL & " values ("
		lgStrSQL = lgStrSQL & 		strTax_biz_area_cd			& ","
		lgStrSQL = lgStrSQL & 	 	strValidFromDt                 & ","
		lgStrSQL = lgStrSQL & 	 	strOwn_rgst_no              & ","
		lgStrSQL = lgStrSQL & 	 	strTax_biz_area_full_nm     & ","
		lgStrSQL = lgStrSQL & 	 	strTax_biz_area_nm          & ","
		lgStrSQL = lgStrSQL & 	 	strTax_biz_area_eng_nm      & ","
		lgStrSQL = lgStrSQL & 	 	strRepre_nm                 & ","
		lgStrSQL = lgStrSQL &		" '', "
		lgStrSQL = lgStrSQL & 	 	strInd_type                 & ","
		lgStrSQL = lgStrSQL & 	 	strInd_class                & ","
		lgStrSQL = lgStrSQL &		" '', "
		lgStrSQL = lgStrSQL & 	strZip_code & ","       				'// varchar             ZIP_CD
		lgStrSQL = lgStrSQL & 	" Convert(nVarchar(100), " & FilterVar(strAddr1, "''", "S") & "),"       					'// varchar             ADDR1
		lgStrSQL = lgStrSQL & 	" Convert(nVarchar(100), " & FilterVar(strAddr2, "''", "S") & "),"       					'// varchar             ADDR2
		lgStrSQL = lgStrSQL &	" Convert(nVarchar(50), " & FilterVar(strEng1_addr, "''", "S") & "), "                      '// varchar             ADDR1_ENG
		lgStrSQL = lgStrSQL &	" Convert(nVarchar(50), " & FilterVar(strEng2_addr, "''", "S") & "), "                      '// varchar             ADDR2_ENG
		lgStrSQL = lgStrSQL &	" Convert(nVarchar(50), " & FilterVar(strEng3_addr, "''", "S") & "), "                      '// varchar             ADDR3_ENG
		lgStrSQL = lgStrSQL & 	 	strInsrt_user_id            & ","
		lgStrSQL = lgStrSQL & 	 	strInsrt_dt                 & ","
		lgStrSQL = lgStrSQL & 	 	strUpdt_user_id             & ","
		lgStrSQL = lgStrSQL & 	 	strUpdt_dt                  & ","
		lgStrSQL = lgStrSQL &		" 0, "
		lgStrSQL = lgStrSQL &		" 0, "
		lgStrSQL = lgStrSQL &		" 0, "
		lgStrSQL = lgStrSQL &		" 0, "
		lgStrSQL = lgStrSQL &		" '', "
		lgStrSQL = lgStrSQL &		" '' "
		lgStrSQL = lgStrSQL &		" )"
	Else
		lgStrSQL =  " 			Update B_Biz_Partner_History "
		lgStrSQL = lgStrSQL & " Set					 "
		lgStrSQL = lgStrSQL & " BP_RGST_NO       = " & 	strOwn_rgst_no & ","
		lgStrSQL = lgStrSQL & " BP_FULL_NM       = " & 	strTax_biz_area_full_nm & ","
		lgStrSQL = lgStrSQL & " BP_NM            = " & 	strTax_biz_area_nm & ","
		lgStrSQL = lgStrSQL & " BP_ENG_NM        = " & 	strTax_biz_area_eng_nm & ","
		lgStrSQL = lgStrSQL & " REPRE_NM         = " & 	strRepre_nm & ","
		lgStrSQL = lgStrSQL & " IND_TYPE         = " & 	strInd_type  & ","
		lgStrSQL = lgStrSQL & " IND_CLASS        = " & 	strInd_class & ","
		lgStrSQL = lgStrSQL & " UPDT_USER_ID     = " & 	strUpdt_user_id & ","
		lgStrSQL = lgStrSQL & " UPDT_DT          = " & 	strUpdt_dt
		lgStrSQL = lgStrSQL & " WHERE 	BP_CD  			 = " & strTax_biz_area_cd
		lgStrSQL = lgStrSQL & " AND 	VALID_FROM_DT       = " & 	strValidFromDt

	End If
	

	lgStrSQLCreate = lgStrSQLCreate & " " & lgStrSQL
	lgObjConn.Execute lgStrSQLCreate,,adCmdText + adExecuteNoRecords
	Call SubHandleError("SCH",lgObjConn,lgObjRs,Err)

End Sub

'============================================================================================================
' Name : SubBizSaveSingleUpdate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveSingleUpdate()
    On Error Resume Next
    Err.Clear

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to update record
    '--------------------------------------------------------------------------------------------------------
	'//zip_code validation
	If Trim(Request("txtZipCode")) <> "" Then
		Call SubMakeSQLStatements("U", FilterVar(Request("txtZipCode"), "''", "S"), "")
		If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then                    'If data not exists
			Call DisplayMsgBox("800016", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
		    Call SetErrorStatus()
		    Exit Sub
		 End If
	End If

	lgStrSQL = "UPDATE  b_tax_biz_area"
	lgStrSQL = lgStrSQL & " SET " 

	lgStrSQL = lgStrSQL & " TEL_NO = " & FilterVar(Request("txtTelNo"), "''", "S") & ","
	lgStrSQL = lgStrSQL & " FAX_NO = " & FilterVar(Request("txtFaxNo"), "''", "S") & ","
	lgStrSQL = lgStrSQL & " ACCT_CHARGE = " & FilterVar(Request("txtAcctCharge"), "''", "S") & ","
	lgStrSQL = lgStrSQL & " TAX_OFFICE_CD = " & FilterVar(Request("txtTaxOfficeCd"), "''", "S") & ","
	lgStrSQL = lgStrSQL & " IS_CHARGE = " & FilterVar(Request("txtIsCharge"), "''", "S") & ","
	lgStrSQL = lgStrSQL & " ZIP_CODE = " & FilterVar(Request("txtZipCode"), "''", "S") & ","
	lgStrSQL = lgStrSQL & " ADDR = Convert(nVarchar(128), " & FilterVar(Trim(Request("txtAddr1")) & " " & Trim(Request("txtAddr2")), "''", "S") & " ),"	
	lgStrSQL = lgStrSQL & " ADDR1 = " & FilterVar(Request("txtAddr1"), "''", "S") & ","
	lgStrSQL = lgStrSQL & " ADDR2 = " & FilterVar(Request("txtAddr2"), "''", "S") & ","
	lgStrSQL = lgStrSQL & " ENG_ADDR =  Convert(nVarchar(128), " & FilterVar(Trim(Request("txtEng1Addr")) & " " & Trim(Request("txtEng2Addr")) & " " & Trim(Request("txtEng3Addr")), "''", "S") & " ),"
	lgStrSQL = lgStrSQL & " ADDR1_ENG = " & FilterVar(Request("txtEng1Addr"), "''", "S") & ","
	lgStrSQL = lgStrSQL & " ADDR2_ENG = " & FilterVar(Request("txtEng2Addr"), "''", "S") & ","
	lgStrSQL = lgStrSQL & " ADDR3_ENG = " & FilterVar(Request("txtEng3Addr"), "''", "S") & ","

	lgStrSQL = lgStrSQL & " UPDT_USER_ID = " & FilterVar(gUsrId, "''","S") & ","   
	lgStrSQL = lgStrSQL & " UPDT_DT = " & FilterVar(GetSvrDateTime, "''", "S")
	lgStrSQL = lgStrSQL & " WHERE TAX_BIZ_AREA_CD = " & lgTaxBizAreaCd

	'---------- Developer Coding part (End  ) ---------------------------------------------------------------
	lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
	Call SubHandleError("SU",lgObjConn,lgObjRs,Err)

	Call SubCloseRs(lgObjRs)                                                    '☜ : Release RecordSSet

End Sub

'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements(pStrMode,pCode, pCode1)
    On Error Resume Next
    Err.Clear
	Dim pMode

	pMode = left(pStrMode,1)

    Select Case pMode 
      Case "R"
            lgStrSQL =  " 		select " 
			lgStrSQL = lgStrSQL & " 	A.TAX_BIZ_AREA_CD,"
			lgStrSQL = lgStrSQL & " 	A.TAX_BIZ_AREA_NM,"
			lgStrSQL = lgStrSQL & " 	A.TAX_BIZ_AREA_FULL_NM,"
			lgStrSQL = lgStrSQL & " 	A.TAX_BIZ_AREA_ENG_NM,"
			lgStrSQL = lgStrSQL & " 	A.CO_CD,"
			lgStrSQL = lgStrSQL & " 	A.TAX_OFFICE_CD,"
			lgStrSQL = lgStrSQL & " 	D.TAX_OFFICE_NM,"
			lgStrSQL = lgStrSQL & " 	A.OWN_RGST_NO,"
			lgStrSQL = lgStrSQL & " 	A.REPRE_NM,"
			lgStrSQL = lgStrSQL & " 	A.IND_CLASS,"
			lgStrSQL = lgStrSQL & " 	B.MINOR_NM AS IND_CLASS_NM,"
			lgStrSQL = lgStrSQL & " 	A.IND_TYPE,"
			lgStrSQL = lgStrSQL & " 	C.MINOR_NM AS IND_TYPE_NM,"
			lgStrSQL = lgStrSQL & " 	A.FAX_NO,"
			lgStrSQL = lgStrSQL & " 	A.TEL_NO,"
			lgStrSQL = lgStrSQL & " 	A.ZIP_CODE,"
			lgStrSQL = lgStrSQL & " 	A.ADDR1,"
			lgStrSQL = lgStrSQL & " 	A.ADDR2,"
			lgStrSQL = lgStrSQL & " 	A.ADDR1_ENG,"
			lgStrSQL = lgStrSQL & " 	A.ADDR2_ENG,"
			lgStrSQL = lgStrSQL & " 	A.ADDR3_ENG,"
			lgStrSQL = lgStrSQL & " 	A.ACCT_CHARGE,"
			lgStrSQL = lgStrSQL & " 	A.IS_CHARGE"
			lgStrSQL = lgStrSQL & " from 	b_tax_biz_area A, "
			lgStrSQL = lgStrSQL & " 	B_MINOR B, B_MINOR C, "
			lgStrSQL = lgStrSQL & " 	B_TAX_OFFICE D"
			lgStrSQL = lgStrSQL & " WHERE  B.MAJOR_CD = " & FilterVar("B9003", "''", "S") & " 	"
			lgStrSQL = lgStrSQL & " AND 	C.MAJOR_CD = " & FilterVar("B9002", "''", "S") & " "
			lgStrSQL = lgStrSQL & " AND	A.IND_CLASS *= B.MINOR_CD"
			lgStrSQL = lgStrSQL & " AND	A.IND_TYPE *= C.MINOR_CD"
			lgStrSQL = lgStrSQL & " AND	A.TAX_OFFICE_CD = D.TAX_OFFICE_CD"
			lgStrSQL = lgStrSQL & " AND	A.TAX_BIZ_AREA_CD = " & pCode

      Case "C"
             Select Case Mid(pStrMode,2,1)
				Case "T"		'//TAX_BIZ_AREA_CD
					lgStrSQL = "Select TAX_BIZ_AREA_CD "
					lgStrSQL = lgStrSQL & " From	b_tax_biz_area "
					lgStrSQL = lgStrSQL & " WHERE	TAX_BIZ_AREA_CD = " & pCode
				Case "P"		'//BP_CD
					lgStrSQL = "Select BP_CD "
					lgStrSQL = lgStrSQL & " From	b_biz_partner "
					lgStrSQL = lgStrSQL & " WHERE	bp_cd = " & pCode
				Case "D"		'//BP_CD
					lgStrSQL = "Select BP_CD "
					lgStrSQL = lgStrSQL & " From	b_biz_partner "
					lgStrSQL = lgStrSQL & " WHERE	bp_cd = " & pCode
					lgStrSQL = lgStrSQL & " AND		USAGE_FLAG = " & FilterVar("N", "''", "S") & " "
					lgStrSQL = lgStrSQL & " AND		BP_TYPE NOT IN (" & FilterVar("C", "''", "S") & " ," & FilterVar("S", "''", "S") & " ," & FilterVar("CS", "''", "S") & " )"
				Case "H"		'//BP_CD
					lgStrSQL = "Select BP_CD "
					lgStrSQL = lgStrSQL & " From	b_biz_partner_history "
					lgStrSQL = lgStrSQL & " WHERE	bp_cd = " & pCode
					lgStrSQL = lgStrSQL & " AND		valid_from_dt = " & pCode1
				Case "A"		'//REPORT_BIZ_AREA_CD
					lgStrSQL = "Select REPORT_BIZ_AREA_CD "
					lgStrSQL = lgStrSQL & " From	b_biz_area "
					lgStrSQL = lgStrSQL & " WHERE	REPORT_BIZ_AREA_CD = " & pCode
				Case "V"		'//REPORT_BIZ_AREA_CD
					lgStrSQL = "Select REPORT_BIZ_AREA_CD "
					lgStrSQL = lgStrSQL & " From	A_VAT "
					lgStrSQL = lgStrSQL & " WHERE	REPORT_BIZ_AREA_CD = " & pCode
				Case "B"		'//b_minor
					lgStrSQL = "Select MINOR_CD, MINOR_NM "
					lgStrSQL = lgStrSQL & " From	B_MINOR "
					lgStrSQL = lgStrSQL & " WHERE	MAJOR_CD = " & pCode1
					lgStrSQL = lgStrSQL & " AND	MINOR_CD = " & pCode
				Case "C"		'//b_minor
					lgStrSQL = "Select CO_CD "
					lgStrSQL = lgStrSQL & " From	B_Company "
				Case "O"		'//TAX_OFFICE_CD
					lgStrSQL = "Select TAX_OFFICE_CD, TAX_OFFICE_NM "
					lgStrSQL = lgStrSQL & " From	B_TAX_OFFICE "
					lgStrSQL = lgStrSQL & " WHERE	TAX_OFFICE_CD = " & pCode
				Case "Z"		'//ZIP_CODE
					lgStrSQL = "Select ZIP_CD, ADDRESS "
					lgStrSQL = lgStrSQL & " From	b_zip_code "
					lgStrSQL = lgStrSQL & " WHERE	ZIP_CD = " & pCode
			End Select 		

      Case "U"					'//ZIP_CODE

       		lgStrSQL = "Select ZIP_CD, ADDRESS "
			lgStrSQL = lgStrSQL & " From	b_zip_code "
			lgStrSQL = lgStrSQL & " WHERE	ZIP_CD = " & pCode
		
      Case "D"					'//TAX_BIZ_AREA_CD
            lgStrSQL = "Select TAX_BIZ_AREA_CD "
			lgStrSQL = lgStrSQL & " From	b_tax_biz_area "
			lgStrSQL = lgStrSQL & " WHERE	TAX_BIZ_AREA_CD = " & pCode
    End Select
'	Response.Write vbcrlf & lgStrSQL & vbcrlf

End Sub

'============================================================================================================
Sub CommonOnTransactionCommit()
End Sub

'============================================================================================================
Sub CommonOnTransactionAbort()
    lgErrorStatus    = "YES"
End Sub

'============================================================================================================
Sub SetErrorStatus()
    lgErrorStatus     = "YES"                                                         '☜: Set error status
End Sub

'============================================================================================================
' Name : SubHandleError
' Desc : This Sub handle error
'============================================================================================================
Sub SubHandleError(pOpCode,pConn,pRs,pErr)
    On Error Resume Next
    Err.Clear

    Select Case pOpCode
        Case "SC"		'//insert
                 If CheckSYSTEMError(pErr,True) = True Then
                    Call DisplayMsgBox("990023", vbInformation, "", "", I_MKSCRIPT)     'Can not create(Demo code)
                    ObjectContext.SetAbort
                    Call SetErrorStatus
                 Else
                    If CheckSQLError(pConn,True) = True Then
                       Call DisplayMsgBox("990023", vbInformation, "", "", I_MKSCRIPT)     'Can not create(Demo code)
                       ObjectContext.SetAbort
                       Call SetErrorStatus
                    Else
						'//성공시 
						Call SubBizPartnerSaveSingleCreate
					End If
                 End If

		Case "SCB"		'//insert B_Biz_Partner
                 If CheckSYSTEMError(pErr,True) = True Then
                    Call DisplayMsgBox("990023", vbInformation, "", "", I_MKSCRIPT)     'Can not create(Demo code)
                    ObjectContext.SetAbort
                    Call SetErrorStatus
                 Else
                    If CheckSQLError(pConn,True) = True Then
                       Call DisplayMsgBox("990023", vbInformation, "", "", I_MKSCRIPT)     'Can not create(Demo code)
                       ObjectContext.SetAbort
                       Call SetErrorStatus
                    Else
						'//성공시 
						Call SubBizPartnerHistorySaveSingleCreate
                    End If
                 End If

         Case "SCH"		'//insert B_Biz_Partner_History
                 If CheckSYSTEMError(pErr,True) = True Then
                    Call DisplayMsgBox("990023", vbInformation, "", "", I_MKSCRIPT)     'Can not create(Demo code)
                    ObjectContext.SetAbort
                    Call SetErrorStatus
                 Else
                    If CheckSQLError(pConn,True) = True Then
                       Call DisplayMsgBox("990023", vbInformation, "", "", I_MKSCRIPT)     'Can not create(Demo code)
                       ObjectContext.SetAbort
                       Call SetErrorStatus
                    Else
						'Call ServerMesgBox("드디어 insert 끝", vbCritical, I_MKSCRIPT)			'⊙:                                          '☜ : Release RecordSSet
				    End If
                 End If
        Case "SD"		'//delete
				If CheckSYSTEMError(pErr,True) = True Then
                    Call DisplayMsgBox("990025", vbInformation, "", "", I_MKSCRIPT)
                    ObjectContext.SetAbort
                    Call SetErrorStatus
                 Else
                    If CheckSQLError(pConn,True) = True Then
                       Call DisplayMsgBox("990025", vbInformation, "", "", I_MKSCRIPT)
                       ObjectContext.SetAbort
                       Call SetErrorStatus
                    Else
						'//성공시 
						Call SubBizUpdateAfterDelete
                    End If
                 End If
        Case "DU"	'//update after delete
				If CheckSYSTEMError(pErr,True) = True Then
                    Call DisplayMsgBox("990025", vbInformation, "", "", I_MKSCRIPT)     'Can not create(Demo code)
                    ObjectContext.SetAbort
                    Call SetErrorStatus
                 Else
                    If CheckSQLError(pConn,True) = True Then
                       Call DisplayMsgBox("990025", vbInformation, "", "", I_MKSCRIPT)     'Can not create(Demo code)
                       ObjectContext.SetAbort
                       Call SetErrorStatus
                    End If
                 End If
        Case "R"
        Case "SU"		'//update
                 If CheckSYSTEMError(pErr,True) = True Then
                    Call DisplayMsgBox("990024", vbInformation, "", "", I_MKSCRIPT)     'Can not create(Demo code)
                    ObjectContext.SetAbort
                    Call SetErrorStatus
                 Else
                    If CheckSQLError(pConn,True) = True Then
                       Call DisplayMsgBox("990024", vbInformation, "", "", I_MKSCRIPT)     'Can not create(Demo code)
                       ObjectContext.SetAbort
                       Call SetErrorStatus
                    End If
                 End If
    End Select

End Sub


Sub SubDeleteTalbleCheck()
    On Error Resume Next
    Err.Clear
	Dim lgStrSql
	Dim i
	Dim lgTableNm
	Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6

	Call CommonQueryRs(" A.NAME "," SYSCOLUMNS B, SYSOBJECTS A "," B.NAME=" & FilterVar("REPORT_BIZ_AREA_CD", "''", "S") & " AND A.ID = B.ID   and A.XTYPE=" & FilterVar("U", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	lgTableNm = split(lgF0,Chr(11))
	For i=0 to Ubound(lgTableNm) - 1
		lgStrSQL = "SELECT report_biz_area_cd FROM " & lgTableNm(i) & " WHERE report_biz_area_cd=" & lgTaxBizAreaBodyCd 
		'response.write lgStrSQL
		If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = True Then                    'If data not exists
			Call DisplayMsgBox("900020", vbInformation, "", "", I_MKSCRIPT)						'☜ : No data is found. 
			response.end
			'Call SetErrorStatus()
		End if
		lgStrSQL = ""
	Next
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
