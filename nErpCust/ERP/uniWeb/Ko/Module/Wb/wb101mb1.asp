<%@ LANGUAGE=VBSCript TRANSACTION=Required%>
<% Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/lgSvrVariables.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../wcm/inc_SvrDebug.asp" -->
<!-- #Include file="../wcm/inc_SvrOperation.asp" -->
<%    
    Call HideStatusWnd                                                               '☜: Hide Processing message
    Call LoadBasisGlobalInf()  
    Call LoadInfTB19029B("I", "H","NOCOOKIE","MB")

	Dim C_SEQ_NO
	Dim C_W_TYPE
	Dim C_W_NAME
	Dim C_W_RGST_NO1
	Dim C_W_MGT_NO
	Dim C_W_RGST_NO
	Dim C_W_RGST_NO2
	Dim C_W_CO_ADDR
	Dim C_W_HOME_ADDR

    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '☜: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)
    lgKeyStream       = Split(Request("txtCo_Cd"),gColSep)

	Call InitSpreadPosVariables
	
    Call SubOpenDB(lgObjConn)                                                        '☜: Make a DB Connection

	'Call CheckVersion(request("txtFISC_YEAR"), request("cboREP_TYPE"))	' 2005-03-11 버전관리기능 추가 
	'PrintLog "lgOpModeCRUD.. : " & lgOpModeCRUD
    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)
             Call SubBizQuery()
        Case CStr(UID_M0002)
             Call SubBizSave()
        Case CStr(UID_M0003)
             Call SubBizDelete()
    End Select

    Call SubCloseDB(lgObjConn)


Sub InitSpreadPosVariables()
    C_SEQ_NO			= 1
    C_W_TYPE			= 2
    C_W_NAME			= 3
    C_W_RGST_NO1		= 4
    C_W_MGT_NO			= 5
    C_W_RGST_NO			= 6
    C_W_RGST_NO2		= 7
    C_W_CO_ADDR			= 8
    C_W_HOME_ADDR		= 9
End Sub

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()
    Dim iKey1, iKey2, iKey3

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    iKey1 = FilterVar(lgKeyStream(0),"''", "S")
    iKey2 = FilterVar(request("txtFISC_YEAR"),"''", "S")
    iKey3 = FilterVar(request("cboREP_TYPE"),"''", "S")

    Call SubMakeSQLStatements("H",iKey1, iKey2, iKey3)                                       '☜ : Make sql statements

    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
          Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
          Call SetErrorStatus()
    Else

%>
<Script Language=vbscript>
       With Parent	
                .Frm1.txtCO_CD_Body.Value  = "<%=ConvSPChars(lgObjRs("CO_CD"))%>"
                .Frm1.txtCO_NM.Value  = "<%=ConvSPChars(lgObjRs("CO_NM"))%>"
                .Frm1.txtCO_ADDR.Value  = "<%=ConvSPChars(lgObjRs("CO_ADDR"))%>"
                .Frm1.txtOWN_RGST_NO.text  = "<%=ConvSPChars(lgObjRs("OWN_RGST_NO"))%>"
                .Frm1.txtLAW_RGST_NO.text  = "<%=ConvSPChars(lgObjRs("LAW_RGST_NO"))%>"
                .Frm1.txtREPRE_NM.Value  = "<%=ConvSPChars(lgObjRs("REPRE_NM"))%>"
                .Frm1.txtREPRE_RGST_NO.text  = "<%=ConvSPChars(lgObjRs("REPRE_RGST_NO"))%>"
                .Frm1.txtTEL_NO.Value  = "<%=ConvSPChars(lgObjRs("TEL_NO"))%>"
                .Frm1.cboCOMP_TYPE1.Value  = "<%=ConvSPChars(lgObjRs("COMP_TYPE1"))%>"
                .Frm1.cboDEBT_MULTIPLE.Value  = "<%=ConvSPChars(lgObjRs("DEBT_MULTIPLE"))%>"
                .Frm1.cboCOMP_TYPE2.Value  = "<%=ConvSPChars(lgObjRs("COMP_TYPE2"))%>"
                .Frm1.txtTAX_OFFICE.Value  = "<%=ConvSPChars(lgObjRs("TAX_OFFICE"))%>"
                .Frm1.txtTAX_OFFICE_NM.Value  = "<%=ConvSPChars(lgObjRs("TAX_OFFICE_NM"))%>"
'                .Frm1.txtFISC_END_DT.Value  = "<%=UNIDateClientFormat(lgObjRs("FISC_END_DT"))%>"
                .Frm1.cboHOLDING_COMP_FLG.Value  = "<%=ConvSPChars(lgObjRs("HOLDING_COMP_FLG"))%>"
                .Frm1.txtIND_CLASS.Value  = "<%=ConvSPChars(lgObjRs("IND_CLASS"))%>"
                .Frm1.txtIND_TYPE.Value  = "<%=ConvSPChars(lgObjRs("IND_TYPE"))%>"
                .Frm1.txtFOUNDATION_DT.text  = "<%=ConvSPChars(lgObjRs("FOUNDATION_DT"))%>"

                .Frm1.txtFISC_YEAR_Body.text  = "<%=ConvSPChars(lgObjRs("FISC_YEAR"))%>"
                .Frm1.cboREP_TYPE_Body.Value  = "<%=ConvSPChars(lgObjRs("REP_TYPE"))%>"
                .Frm1.txtFISC_START_DT.text  = "<%=UNIDateClientFormat(lgObjRs("FISC_START_DT"))%>"
                .Frm1.txtFISC_END_DT.text  = "<%=UNIDateClientFormat(lgObjRs("FISC_END_DT"))%>"
                .Frm1.txtHOME_ANY_START_DT.text  = "<%=UNIDateClientFormat(lgObjRs("HOME_ANY_START_DT"))%>"
                .Frm1.txtHOME_ANY_END_DT.text  = "<%=UNIDateClientFormat(lgObjRs("HOME_ANY_END_DT"))%>"
                .Frm1.txtHOME_TAX_USR_ID.Value  = "<%=ConvSPChars(lgObjRs("HOME_TAX_USR_ID"))%>"
                .Frm1.txtHOME_TAX_EMAIL.Value  = "<%=ConvSPChars(lgObjRs("HOME_TAX_E_MAIL"))%>"
                .Frm1.txtHOME_TAX_MAIN_IND.Value  = "<%=ConvSPChars(lgObjRs("HOME_TAX_MAIN_IND"))%>"
                .Frm1.txtHOME_TAX_MAIN_IND_NM.Value = "<%=ConvSPChars(lgObjRs("HOME_TAX_MAIN_IND_NM"))%>"
                
                .Frm1.txtHOME_FILE_MAKE_DT.TEXT  = "<%=ConvSPChars(lgObjRs("HOME_FILE_MAKE_DT"))%>"
				.Frm1.txtINCOM_DT.TEXT  = "<%=ConvSPChars(lgObjRs("INCOM_DT"))%>"
				
                .Frm1.txtAGENT_NM.Value  = "<%=ConvSPChars(lgObjRs("AGENT_NM"))%>"
                .Frm1.txtRECON_BAN_NO.text  = "<%=ConvSPChars(lgObjRs("RECON_BAN_NO"))%>"
                .Frm1.txtRECON_MGT_NO.text  = "<%=ConvSPChars(lgObjRs("RECON_MGT_NO"))%>"
                .Frm1.txtAGENT_TEL_NO.Value  = "<%=ConvSPChars(lgObjRs("AGENT_TEL_NO"))%>"
				.Frm1.txtAGENT_RGST_NO.text  = "<%=ConvSPChars(lgObjRs("AGENT_RGST_NO"))%>"
				.Frm1.txtREQUEST_DT.Value  = "<%=ConvSPChars(lgObjRs("REQUEST_DT"))%>"
				.Frm1.txtAPPO_NO.TEXT  = "<%=ConvSPChars(lgObjRs("APPO_NO"))%>"
				.Frm1.txtAPPO_DT.TEXT  = "<%=ConvSPChars(lgObjRs("APPO_DT"))%>"
				.Frm1.txtAPPO_DESC.Value  = "<%=ConvSPChars(lgObjRs("APPO_DESC"))%>"
				.Frm1.cboEX_RECON_FLG.Value  = "<%=ConvSPChars(lgObjRs("EX_RECON_FLG"))%>"
				.Frm1.cboEX_54_FLG.Value  = "<%=ConvSPChars(lgObjRs("EX_54_FLG"))%>"
				
                .Frm1.txtBANK_CD.Value  = "<%=ConvSPChars(lgObjRs("BANK_CD"))%>"
                .Frm1.txtBANK_NM.Value  = "<%=ConvSPChars(lgObjRs("BANK_NM"))%>"
                .Frm1.txtBANK_BRANCH.Value  = "<%=ConvSPChars(lgObjRs("BANK_BRANCH"))%>"
                .Frm1.txtBANK_DPST.Value  = "<%=ConvSPChars(lgObjRs("BANK_DPST"))%>"
                .Frm1.txtBANK_ACCT_NO.Value  = "<%=ConvSPChars(lgObjRs("BANK_ACCT_NO"))%>"

				.Frm1.cboSUBMIT_FLG.Value  = "<%=ConvSPChars(lgObjRs("SUBMIT_FLG"))%>"
				.Frm1.cboUSE_FLG.Value  = "<%=ConvSPChars(lgObjRs("USE_FLG"))%>"
				.Frm1.txtREVISION_YM.Value  = "<%=ConvSPChars(lgObjRs("REVISION_YM"))%>"
				
				
       End With          
</Script>       
<%     
		Dim iDx, iStrData, iIntMaxRows
        ' 1번째 그리드 
        Call SubMakeSQLStatements("D",iKey1, iKey2, iKey3) 
        
		If   FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = True Then

		    'Call SubSkipRs(lgObjRs,C_SHEETMAXROWS_D * lgStrPrevKey)
		    lgstrData = ""
		    
		    iDx = 1
		    
		    Do While Not lgObjRs.EOF
				iStrData = iStrData & Chr(11) & ConvSPChars(lgObjRs("SEQ_NO"))
				iStrData = iStrData & Chr(11) & ConvSPChars(lgObjRs("W_TYPE"))			
				iStrData = iStrData & Chr(11) & ConvSPChars(lgObjRs("W_NAME"))	
				iStrData = iStrData & Chr(11) & ConvSPChars(lgObjRs("W_RGST_NO1"))	
				iStrData = iStrData & Chr(11) & ConvSPChars(lgObjRs("W_MGT_NO"))	
				iStrData = iStrData & Chr(11) & ConvSPChars(lgObjRs("W_RGST_NO"))		
				iStrData = iStrData & Chr(11) & ConvSPChars(lgObjRs("W_RGST_NO2"))		
				iStrData = iStrData & Chr(11) & ConvSPChars(lgObjRs("W_CO_ADDR"))		
				iStrData = iStrData & Chr(11) & ConvSPChars(lgObjRs("W_HOME_ADDR"))			 
				iStrData = iStrData & Chr(11) & iIntMaxRows + iDx + 1
				iStrData = iStrData & Chr(11) & Chr(12)
			    lgObjRs.MoveNext

		        iDx =  iDx + 1
           
		    Loop 
		    
			lgObjRs.Close
			Set lgObjRs = Nothing

		End If 

		Response.Write " <Script Language=vbscript>	                        " & vbCr
		Response.Write " With parent                                        " & vbCr
		Response.Write "	.ggoSpread.Source = .frm1.vspdData " & vbCr
		Response.Write "	.ggoSpread.SSShowData """ & iStrData       & """" & vbCr
		Response.Write " End With                                           " & vbCr
		Response.Write " </Script>                                          " & vbCr
		
    End If
    Call SubCloseRs(lgObjRs)
End Sub	
'============================================================================================================
' Name : SubBizQuery
' Desc : Date data 
'============================================================================================================
Sub SubBizSave()
	Dim arrRowVal
    Dim arrColVal, lgLngMaxRow
    Dim iDx , i
    
    On Error Resume Next
    Err.Clear
  
    lgIntFlgMode = CInt(Request("txtFlgMode"))                                       '☜: Read Operayion Mode (CREATE, UPDATE)
 
    Select Case lgIntFlgMode
        Case  OPMD_CMODE                                                             '☜ : Create
              Call SubBizSaveSingleCreate()  
        Case  OPMD_UMODE           
              Call SubBizSaveSingleUpdate()
    End Select

	PrintLog "1번째 그리드. .: " & Request("txtSpread") 

	' --- 1번째 그리드 
	arrRowVal = Split(Request("txtSpread"), gRowSep)                                 '☜: Split Row    data
	lgLngMaxRow = UBound(arrRowVal)

    For iDx = 1 To lgLngMaxRow
        arrColVal = Split(arrRowVal(iDx-1), gColSep)    
        
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
' Name : SubBizDelete
' Desc : Delete DB data
'============================================================================================================
Sub SubBizDelete()
    On Error Resume Next
    Err.Clear

    lgStrSQL = lgStrSQL & "DELETE TB_AGENT_INFO WITH (ROWLOCK) " & vbCrLf
    lgStrSQL = lgStrSQL & " WHERE CO_CD = " &  FilterVar(Request("txtCO_CD_Body"),"''", "S") & vbCrLf
    lgStrSQL = lgStrSQL & " AND FISC_YEAR = " &  FilterVar(Request("txtFISC_YEAR_Body"),"''", "S") & vbCrLf
    lgStrSQL = lgStrSQL & " AND REP_TYPE = " &  FilterVar(Request("cboREP_TYPE_Body"),"''", "S") & vbCrLf
    
    lgStrSQL = lgStrSQL & "DELETE  TB_COMPANY_HISTORY WITH (ROWLOCK) "
    lgStrSQL = lgStrSQL & " WHERE CO_CD = " &  FilterVar(Request("txtCO_CD_Body"),"''", "S") & vbCrLf
    lgStrSQL = lgStrSQL & " AND FISC_YEAR = " &  FilterVar(Request("txtFISC_YEAR_Body"),"''", "S") & vbCrLf
    lgStrSQL = lgStrSQL & " AND REP_TYPE = " &  FilterVar(Request("cboREP_TYPE_Body"),"''", "S") & vbCrLf

	PrintLog "SubBizDelete = " & lgStrSQL
    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("SD",lgObjConn,lgObjRs,Err)
End Sub

'============================================================================================================
' Name : SubBizSaveCreate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveSingleCreate()
    On Error Resume Next
    Err.Clear

	lgStrSQL =            " INSERT INTO TB_COMPANY_HISTORY WITH (ROWLOCK) " & vbCrLf
	lgStrSQL = lgStrSQL & " (CO_CD, FISC_YEAR, REP_TYPE, CO_NM, CO_ADDR "& vbCrLf
	lgStrSQL = lgStrSQL & "	, OWN_RGST_NO, LAW_RGST_NO, REPRE_NM, REPRE_RGST_NO, TEL_NO "& vbCrLf
	lgStrSQL = lgStrSQL & "	, COMP_TYPE1, DEBT_MULTIPLE, COMP_TYPE2, TAX_OFFICE "& vbCrLf
	lgStrSQL = lgStrSQL & "	, HOLDING_COMP_FLG, IND_CLASS, IND_TYPE, FOUNDATION_DT "& vbCrLf
	lgStrSQL = lgStrSQL & "	, HOME_TAX_USR_ID, HOME_TAX_E_MAIL, HOME_TAX_MAIN_IND "& vbCrLf
	lgStrSQL = lgStrSQL & "	, FISC_START_DT, FISC_END_DT "& vbCrLf
	lgStrSQL = lgStrSQL & "	, HOME_ANY_START_DT, HOME_ANY_END_DT, INCOM_DT, HOME_FILE_MAKE_DT "& vbCrLf
	lgStrSQL = lgStrSQL & "	, AGENT_NM, RECON_BAN_NO, RECON_MGT_NO, AGENT_TEL_NO, AGENT_RGST_NO "& vbCrLf
	lgStrSQL = lgStrSQL & "	, REQUEST_DT, APPO_NO, APPO_DT, APPO_DESC "& vbCrLf
	lgStrSQL = lgStrSQL & "	, BANK_CD, BANK_BRANCH, BANK_DPST, BANK_ACCT_NO , EX_RECON_FLG, EX_54_FLG, SUBMIT_FLG, USE_FLG, REVISION_YM"& vbCrLf
	lgStrSQL = lgStrSQL & "	, INSRT_USER_ID, UPDT_USER_ID ) " & vbCrLf
	lgStrSQL = lgStrSQL & " VALUES ( " & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UCASE(Request("txtCO_CD_Body")),"''","S") & ", -- txtCO_CD_Body"& vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(Request("txtFISC_YEAR_Body"),"''","S") & ", -- txtFISC_YEAR_Body"& vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(Request("cboREP_TYPE_Body"),"''","S") & ", -- cboREP_TYPE_Body"& vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(Request("txtCO_NM"),"''","S") & ", -- txtCO_NM"& vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(Request("txtCO_ADDR"),"''","S") & ", -- txtCO_ADDR"& vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(Request("txtOWN_RGST_NO"),"''","S") & ", -- txtOWN_RGST_NO"& vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(Request("txtLAW_RGST_NO"),"''","S") & ", -- txtLAW_RGST_NO"& vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(Request("txtREPRE_NM"),"''","S") & ", -- txtREPRE_NM"& vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(Request("txtREPRE_RGST_NO"),"''","S") & ", -- txtREPRE_RGST_NO"& vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(Request("txtTEL_NO"),"''","S") & ", -- txtTEL_NO"& vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(Request("cboCOMP_TYPE1"),"''","S") & ", -- cboCOMP_TYPE1"& vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(Request("cboDEBT_MULTIPLE"),"''","S") & ", -- cboDEBT_MULTIPLE"& vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(Request("cboCOMP_TYPE2"),"''","S") & ", -- cboCOMP_TYPE2"& vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(Request("txtTAX_OFFICE"),"''","S") & ", -- txtTAX_OFFICE"& vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(Request("cboHOLDING_COMP_FLG"),"''","S") & ", -- cboHOLDING_COMP_FLG"& vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(Request("txtIND_CLASS"),"''","S") & ", -- txtIND_CLASS"& vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(Request("txtIND_TYPE"),"''","S") & ", -- txtIND_TYPE"& vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(Request("txtFOUNDATION_DT"),"''","S") & ", -- txtFOUNDATION_DT"& vbCrLf

	lgStrSQL = lgStrSQL & FilterVar(Request("txtHOME_TAX_USR_ID"),"''","S") & ", -- txtHOME_TAX_USR_ID"& vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(Request("txtHOME_TAX_EMAIL"),"''","S") & ", -- txtHOME_TAX_EMAIL"& vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(Request("txtHOME_TAX_MAIN_IND"),"''","S") & ", -- txtHOME_TAX_MAIN_IND"& vbCrLf

	lgStrSQL = lgStrSQL & FilterVar(Request("txtFISC_START_DT"),"''","S") & ", -- txtFISC_START_DT"& vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(Request("txtFISC_END_DT"),"''","S") & ", -- txtFISC_END_DT"& vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(Request("txtHOME_ANY_START_DT"),"NULL","S") & ", -- txtHOME_ANY_START_DT"& vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(Request("txtHOME_ANY_END_DT"),"NULL","S") & ", -- txtHOME_ANY_END_DT"& vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(Request("txtINCOM_DT"),"NULL","S") & ", -- txtINCOM_DT"& vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(Request("txtHOME_FILE_MAKE_DT"),"NULL","S") & ", -- txtHOME_FILE_MAKE_DT"& vbCrLf
	
	lgStrSQL = lgStrSQL & FilterVar(Request("txtAGENT_NM"),"''","S") & ", -- txtAGENT_NM"& vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(Request("txtRECON_BAN_NO"),"''","S") & ", -- txtRECON_BAN_NO"& vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(Request("txtRECON_MGT_NO"),"''","S") & ", -- txtRECON_MGT_NO"& vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(Request("txtAGENT_TEL_NO"),"''","S") & ", -- txtAGENT_TEL_NO"& vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(Request("txtAGENT_RGST_NO"),"''","S") & ", -- txtAGENT_RGST_NO"& vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(Request("txtREQUEST_DT"),"NULL","S") & ", -- txtREQUEST_DT"& vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(Request("txtAPPO_NO"),"''","S") & ", -- txtAPPO_NO"& vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(Request("txtAPPO_DT"),"NULL","S") & ", -- txtAPPO_DT"& vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(Request("txtAPPO_DESC"),"''","S") & ", -- txtAPPO_DESC"& vbCrLf
	
	lgStrSQL = lgStrSQL & FilterVar(Request("txtBANK_CD"),"''","S") & ", -- txtBANK_CD"& vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(Request("txtBANK_BRANCH"),"''","S") & ", -- txtBANK_BRANCH"& vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(Request("txtBANK_DPST"),"''","S") & ", -- txtBANK_DPST"& vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(Request("txtBANK_ACCT_NO"),"''","S") & ", -- txtBANK_ACCT_NO"& vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(Request("cboEX_RECON_FLG"),"''","S") & ", -- cboEX_RECON_FLG"& vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(Request("cboEX_54_FLG"),"''","S") & ", -- cboEX_54_FLG"& vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(Request("cboSUBMIT_FLG"),"''","S") & ", -- cboSUBMIT_FLG"& vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(Request("cboUSE_FLG"),"''","S") & ", -- cboUSE_FLG"& vbCrLf
	
	lgStrSQL = lgStrSQL & FilterVar(C_REVISION_YM,"''","S") & ","& vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S") & ","& vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S") & ""& vbCrLf
		       
	lgStrSQL = lgStrSQL & "   ) " & vbCrLf

	PrintLog "SubBizSaveSingleCreate_Create = " & lgStrSQL

	lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
 		
	Call SubHandleError("SU",lgObjConn,lgObjRs,Err)
End Sub

'============================================================================================================
' Name : SubBizSaveSingleUpdate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveSingleUpdate()
    On Error Resume Next
    Err.Clear

    lgStrSQL =            "UPDATE TB_COMPANY_HISTORY WITH (ROWLOCK) " & vbCrLf
    lgStrSQL = lgStrSQL & "   SET "  & vbCrLf
    lgStrSQL = lgStrSQL & "       CO_NM = " & FilterVar(Request("txtCO_NM"),"''","S") & "," & vbCrLf
    lgStrSQL = lgStrSQL & "       CO_ADDR = " & FilterVar(Request("txtCO_ADDR"),"''","S") & "," & vbCrLf
    lgStrSQL = lgStrSQL & "       OWN_RGST_NO = " & FilterVar(Request("txtOWN_RGST_NO"),"''","S") & "," & vbCrLf
    lgStrSQL = lgStrSQL & "       LAW_RGST_NO = " & FilterVar(Request("txtLAW_RGST_NO"),"''","S") & "," & vbCrLf
    lgStrSQL = lgStrSQL & "       REPRE_NM = " & FilterVar(Request("txtREPRE_NM"),"''","S") & "," & vbCrLf
    lgStrSQL = lgStrSQL & "       REPRE_RGST_NO = " & FilterVar(Request("txtREPRE_RGST_NO"),"''","S") & "," & vbCrLf
    lgStrSQL = lgStrSQL & "       TEL_NO = " & FilterVar(Request("txtTEL_NO"),"''","S") & "," & vbCrLf
    lgStrSQL = lgStrSQL & "       COMP_TYPE1 = " & FilterVar(Request("cboCOMP_TYPE1"),"''","S") & "," & vbCrLf
    lgStrSQL = lgStrSQL & "       DEBT_MULTIPLE = " & FilterVar(Request("cboDEBT_MULTIPLE"),"''","S") & "," & vbCrLf
    lgStrSQL = lgStrSQL & "       COMP_TYPE2 = " & FilterVar(Request("cboCOMP_TYPE2"),"''","S") & "," & vbCrLf
    lgStrSQL = lgStrSQL & "       TAX_OFFICE = " & FilterVar(Request("txtTAX_OFFICE"),"''","S") & "," & vbCrLf
    lgStrSQL = lgStrSQL & "       HOLDING_COMP_FLG = " & FilterVar(Request("cboHOLDING_COMP_FLG"),"''","S") & "," & vbCrLf
    lgStrSQL = lgStrSQL & "       IND_CLASS = " & FilterVar(Request("txtIND_CLASS"),"''","S") & "," & vbCrLf
    lgStrSQL = lgStrSQL & "       IND_TYPE = " & FilterVar(Request("txtIND_TYPE"),"''","S") & ", " & vbCrLf
    lgStrSQL = lgStrSQL & "       FOUNDATION_DT = " & FilterVar(Request("txtFOUNDATION_DT"),"''","S") & ", " & vbCrLf
    lgStrSQL = lgStrSQL & "       HOME_TAX_USR_ID = " & FilterVar(Request("txtHOME_TAX_USR_ID"),"''","S") & ", " & vbCrLf
    lgStrSQL = lgStrSQL & "       HOME_TAX_E_MAIL = " & FilterVar(Request("txtHOME_TAX_EMAIL"),"''","S") & ", " & vbCrLf
    lgStrSQL = lgStrSQL & "       HOME_TAX_MAIN_IND = " & FilterVar(Request("txtHOME_TAX_MAIN_IND"),"''","S") & ", " & vbCrLf
	
    lgStrSQL = lgStrSQL & "       FISC_START_DT = " & FilterVar(Request("txtFISC_START_DT"),"''","S") & ", " & vbCrLf
    lgStrSQL = lgStrSQL & "       FISC_END_DT = " & FilterVar(Request("txtFISC_END_DT"),"''","S") & ", " & vbCrLf
    lgStrSQL = lgStrSQL & "       HOME_ANY_START_DT = " & FilterVar(Request("txtHOME_ANY_START_DT"),"NULL","S") & ", " & vbCrLf
    lgStrSQL = lgStrSQL & "       HOME_ANY_END_DT = " & FilterVar(Request("txtHOME_ANY_END_DT"),"NULL","S") & ", " & vbCrLf

    lgStrSQL = lgStrSQL & "       HOME_FILE_MAKE_DT = " & FilterVar(Request("txtHOME_FILE_MAKE_DT"),"NULL","S") & ", " & vbCrLf
	lgStrSQL = lgStrSQL & "       INCOM_DT = " & FilterVar(Request("txtINCOM_DT"),"NULL","S") & ", " & vbCrLf
	
	lgStrSQL = lgStrSQL & "       BANK_CD = " & FilterVar(Request("txtBANK_CD"),"''","S") & ", " & vbCrLf
    lgStrSQL = lgStrSQL & "       BANK_BRANCH = " & FilterVar(Request("txtBANK_BRANCH"),"''","S") & ", " & vbCrLf
    lgStrSQL = lgStrSQL & "       BANK_DPST = " & FilterVar(Request("txtBANK_DPST"),"''","S") & ", " & vbCrLf
    lgStrSQL = lgStrSQL & "       BANK_ACCT_NO = " & FilterVar(Request("txtBANK_ACCT_NO"),"''","S") & ", " & vbCrLf
	
    lgStrSQL = lgStrSQL & "       AGENT_NM = " & FilterVar(Request("txtAGENT_NM"),"''","S") & ", " & vbCrLf
    lgStrSQL = lgStrSQL & "       RECON_BAN_NO = " & FilterVar(Request("txtRECON_BAN_NO"),"''","S") & ", " & vbCrLf
    lgStrSQL = lgStrSQL & "       RECON_MGT_NO = " & FilterVar(Request("txtRECON_MGT_NO"),"''","S") & ", " & vbCrLf
    lgStrSQL = lgStrSQL & "       AGENT_TEL_NO = " & FilterVar(Request("txtAGENT_TEL_NO"),"''","S") & ", " & vbCrLf
    lgStrSQL = lgStrSQL & "       AGENT_RGST_NO = " & FilterVar(Request("txtAGENT_RGST_NO"),"''","S") & ", " & vbCrLf
    lgStrSQL = lgStrSQL & "       REQUEST_DT = " & FilterVar(Request("txtREQUEST_DT"),"NULL","S") & ", " & vbCrLf
    lgStrSQL = lgStrSQL & "       APPO_NO = " & FilterVar(Request("txtAPPO_NO"),"''","S") & ", " & vbCrLf
    lgStrSQL = lgStrSQL & "       APPO_DT = " & FilterVar(Request("txtAPPO_DT"),"NULL","S") & ", " & vbCrLf
    lgStrSQL = lgStrSQL & "       APPO_DESC = " & FilterVar(Request("txtAPPO_DESC"),"''","S") & ", " & vbCrLf
    lgStrSQL = lgStrSQL & "       EX_RECON_FLG = " & FilterVar(Request("cboEX_RECON_FLG"),"''","S") & "," & vbCrLf
    lgStrSQL = lgStrSQL & "       EX_54_FLG = " & FilterVar(Request("cboEX_54_FLG"),"''","S") & "," & vbCrLf
    lgStrSQL = lgStrSQL & "       SUBMIT_FLG = " & FilterVar(Request("cboSUBMIT_FLG"),"''","S") & "," & vbCrLf
    lgStrSQL = lgStrSQL & "       USE_FLG = " & FilterVar(Request("cboUSE_FLG"),"''","S") & "," & vbCrLf
    
    lgStrSQL = lgStrSQL & "		  UPDT_USER_ID = " & FilterVar(gUsrId,"''","S") & ", " & vbCrLf
	lgStrSQL = lgStrSQL & "		  UPDT_DT = GETDATE() " & vbCrLf
    lgStrSQL = lgStrSQL & " WHERE CO_CD = " &  FilterVar(Request("txtCO_CD_Body"),"''", "S") & vbCrLf
    lgStrSQL = lgStrSQL & " AND FISC_YEAR = " &  FilterVar(Request("txtFISC_YEAR_Body"),"''", "S") & vbCrLf
    lgStrSQL = lgStrSQL & " AND REP_TYPE = " &  FilterVar(Request("cboREP_TYPE_Body"),"''", "S") & vbCrLf

    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
 
	PrintLog "SubBizSaveSingleUpdate.. : " & lgStrSQL
	Call SubHandleError("SU",lgObjConn,lgObjRs,Err)
End Sub

'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements(pMode,pCode1, pCode2, pCode3)
    Select Case pMode 
      Case "H"

		lgStrSQL =			  " SELECT  " & vbCrLf
        lgStrSQL = lgStrSQL & "a.CO_CD, a.CO_NM, a.CO_ADDR, a.OWN_RGST_NO, a.LAW_RGST_NO, a.REPRE_NM, a.REPRE_RGST_NO " & vbCrLf
        lgStrSQL = lgStrSQL & " , a.TEL_NO, a.COMP_TYPE1, a.DEBT_MULTIPLE, a.COMP_TYPE2, a.TAX_OFFICE,  dbo.ufn_GetCodeName('W1079', a.TAX_OFFICE) as TAX_OFFICE_NM " & vbCrLf
        lgStrSQL = lgStrSQL & " , a.HOLDING_COMP_FLG, a.IND_CLASS, a.IND_TYPE, a.FOUNDATION_DT " & vbcrlf
        lgStrSQL = lgStrSQL & ", A.IND_CLASS, A.IND_TYPE " & vbCrLf
        lgStrSQL = lgStrSQL & " , A.FISC_YEAR, A.REP_TYPE " & vbCrLf
        lgStrSQL = lgStrSQL & " , A.HOME_TAX_USR_ID " & vbCrLf
        lgStrSQL = lgStrSQL & " , A.HOME_TAX_E_MAIL, A.HOME_TAX_MAIN_IND, C.DETAIL_NM HOME_TAX_MAIN_IND_NM, A.EX_RECON_FLG, A.EX_54_FLG, A.SUBMIT_FLG, A.USE_FLG " & vbCrLf
	
        lgStrSQL = lgStrSQL & " , A.FISC_START_DT , A.FISC_END_DT, A.HOME_ANY_START_DT, A.HOME_ANY_END_DT " & vbCrLf
        lgStrSQL = lgStrSQL & " , A.HOME_FILE_MAKE_DT, A.INCOM_DT, A.AGENT_NM, A.RECON_BAN_NO, A.RECON_MGT_NO, A.AGENT_TEL_NO, A.AGENT_RGST_NO, A.REQUEST_DT " & vbCrLf
        lgStrSQL = lgStrSQL & " , A.APPO_NO, A.APPO_DT, A.APPO_DESC, A.REVISION_YM, A.USE_FLG " & vbCrLf
        lgStrSQL = lgStrSQL & " , A.BANK_CD, dbo.ufn_GetCodeName('W1020', A.BANK_CD) as BANK_NM, A.BANK_BRANCH, A.BANK_DPST, A.BANK_ACCT_NO " & vbCrLf
        lgStrSQL = lgStrSQL & " FROM TB_COMPANY_HISTORY A (nolock) " & vbCrLf
		lgStrSQL = lgStrSQL & "		LEFT OUTER JOIN  tb_std_income_rate C (NOLOCK) ON A.HOME_TAX_MAIN_IND = C.STD_INCM_RT_CD" & vbCrLf

        lgStrSQL = lgStrSQL & " WHERE A.CO_CD = " & pCode1 	 & vbCrLf
		lgStrSQL = lgStrSQL & " 	 AND A.FISC_YEAR = " & pCode2 	 & vbCrLf
	    lgStrSQL = lgStrSQL & "      AND A.REP_TYPE = " & pCode3 	&  vbCrLf

	  Case "D"
		lgStrSQL =			  " SELECT  " & vbCrLf
        lgStrSQL = lgStrSQL & " SEQ_NO, W_TYPE, W_NAME, W_RGST_NO1, W_MGT_NO, W_RGST_NO, W_RGST_NO2, W_CO_ADDR, W_HOME_ADDR " & vbCrLf
        lgStrSQL = lgStrSQL & " FROM TB_AGENT_INFO (nolock) " & vbCrLf
		
        lgStrSQL = lgStrSQL & " WHERE CO_CD = " & pCode1 	 & vbCrLf
		lgStrSQL = lgStrSQL & " 	 AND FISC_YEAR = " & pCode2 	 & vbCrLf
	    lgStrSQL = lgStrSQL & "      AND REP_TYPE = " & pCode3 	&  vbCrLf
    End Select
	PrintLog "SubMakeSQLStatements.. : " & lgStrSQL
End Sub

'============================================================================================================
' Name : SubBizSaveCreate
' Desc : 1번째 그리드 저장 
'============================================================================================================
Sub SubBizSaveMultiCreate(arrColVal)
	On Error Resume Next   
	Err.Clear
	dim i
	
	lgStrSQL = "INSERT INTO TB_AGENT_INFO WITH (ROWLOCK) (" & vbCrLf
	lgStrSQL = lgStrSQL & " CO_CD, FISC_YEAR, REP_TYPE, SEQ_NO, W_TYPE,W_NAME, W_RGST_NO1, W_MGT_NO, W_RGST_NO, W_RGST_NO2, W_CO_ADDR, W_HOME_ADDR "   & vbCrLf
	lgStrSQL = lgStrSQL & " , INSRT_USER_ID, UPDT_USER_ID ) VALUES ("  & vbCrLf
	    
	lgStrSQL = lgStrSQL & FilterVar(UCASE(Request("txtCO_CD_Body")),"''","S") & ", -- txtCO_CD_Body"& vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(Request("txtFISC_YEAR_Body"),"''","S") & ", -- txtFISC_YEAR_Body"& vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(Request("cboREP_TYPE_Body"),"''","S") & ", -- cboREP_TYPE_Body"& vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(UNICDbl(arrColVal(C_SEQ_NO), "0"),"1","D")		& "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_W_TYPE))),"''","S")		& "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_W_NAME))),"''","S")		& "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_W_RGST_NO1))),"''","S")		& "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_W_MGT_NO))),"''","S")		& "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_W_RGST_NO))),"''","S")		& "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_W_RGST_NO2))),"''","S")		& "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_W_CO_ADDR))),"''","S")		& "," & vbCrLf
	lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(arrColVal(C_W_HOME_ADDR))),"''","S")		& "," & vbCrLf
	
	'lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,"''","S") & "," 
	lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S")                        & "," 
	'lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,"''","S") & "," 
	lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S")                        
	lgStrSQL = lgStrSQL & ")"

	PrintLog "SubBizSaveMultiCreate = " & lgStrSQL

	lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MC",lgObjConn,lgObjRs,Err)

End Sub

'============================================================================================================
' Name : SubBizSaveMultiUpdate
' Desc : 1번 그리드 업데이트 
'============================================================================================================
Sub SubBizSaveMultiUpdate(arrColVal)
	dim i
	
	On Error Resume Next 
	Err.Clear                                                                        '☜: Clear Error status

    lgStrSQL = "UPDATE  TB_AGENT_INFO WITH (ROWLOCK) "
    
    lgStrSQL = lgStrSQL & " SET " 
    lgStrSQL = lgStrSQL & " W_NAME	= " &  FilterVar(Trim(UCase(arrColVal(C_W_NAME ))),"''","S") & ","
    lgStrSQL = lgStrSQL & " W_RGST_NO1  = " &  FilterVar(Trim(UCase(arrColVal(C_W_RGST_NO1 ))),"''","S") & ","
    lgStrSQL = lgStrSQL & " W_MGT_NO  = " &  FilterVar(Trim(UCase(arrColVal(C_W_MGT_NO ))),"''","S") & ","
    lgStrSQL = lgStrSQL & " W_RGST_NO = " &  FilterVar(Trim(UCase(arrColVal(C_W_RGST_NO))),"''","S") & ","
    lgStrSQL = lgStrSQL & " W_RGST_NO2 = " &  FilterVar(Trim(UCase(arrColVal(C_W_RGST_NO2))),"''","S") & ","
    lgStrSQL = lgStrSQL & " W_CO_ADDR = " &  FilterVar(Trim(arrColVal(C_W_CO_ADDR)),"","S") & ","
    lgStrSQL = lgStrSQL & " W_HOME_ADDR = " &  FilterVar(Trim(arrColVal(C_W_HOME_ADDR)),"","S") & ","
                   
    lgStrSQL = lgStrSQL & " UPDT_DT      = " &  FilterVar(GetSvrDateTime,"''","S") & ","           
    lgStrSQL = lgStrSQL & " UPDT_USER_ID      = " &  FilterVar(gUsrId,"''","S")                  

    lgStrSQL = lgStrSQL & " WHERE CO_CD = " &  FilterVar(Request("txtCO_CD_Body"),"''", "S") & vbCrLf
    lgStrSQL = lgStrSQL & " AND FISC_YEAR = " &  FilterVar(Request("txtFISC_YEAR_Body"),"''", "S") & vbCrLf
    lgStrSQL = lgStrSQL & " AND REP_TYPE = " &  FilterVar(Request("cboREP_TYPE_Body"),"''", "S") & vbCrLf
	lgStrSQL = lgStrSQL & " AND SEQ_NO = " &  FilterVar(UNICDbl(arrColVal(C_SEQ_NO), "0"),"1","D") & vbCrLf

	PrintLog "SubBizSaveMultiUpdate = " & lgStrSQL
    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MU",lgObjConn,lgObjRs,Err)
	
End Sub

'============================================================================================================
' Name : SubBizSaveMultiDelete
' Desc : 1번째 그리드 삭제 
'============================================================================================================
Sub SubBizSaveMultiDelete(arrColVal)
   On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    lgStrSQL = "DELETE  TB_AGENT_INFO WITH (ROWLOCK) "
 	lgStrSQL = lgStrSQL & " WHERE CO_CD = " &  FilterVar(Request("txtCO_CD_Body"),"''", "S") & vbCrLf
    lgStrSQL = lgStrSQL & " AND FISC_YEAR = " &  FilterVar(Request("txtFISC_YEAR_Body"),"''", "S") & vbCrLf
    lgStrSQL = lgStrSQL & " AND REP_TYPE = " &  FilterVar(Request("cboREP_TYPE_Body"),"''", "S") & vbCrLf
	lgStrSQL = lgStrSQL & "		AND SEQ_NO = " & FilterVar(UNICDbl(arrColVal(C_SEQ_NO), "0"),"1","D") & vbCrLf

    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MD",lgObjConn,lgObjRs,Err)

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
        Case "SD"
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
