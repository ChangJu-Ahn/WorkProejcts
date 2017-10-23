
<%@ LANGUAGE=VBSCript %>
<%Option Explicit%>

<%
'======================================================================================================
'*  1. Module Name          : Accounting
'*  2. Function Name        : Treasury - Budget
'*  3. Program ID           : f2102mb1_lk0441
'*  4. Program Name         : 예산정보등록(LKO441) 
'*  5. Program Desc         : Register of Budget
'*  6. Comproxy List        : FU0021, FU0028
'*  7. Modified date(First) : 
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : 
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'*   - 2001.03.09 Song, Mun Gil 예산년월에 Mask 적용 
'=======================================================================================================

'☜ : 항상 서버 사이드 구문의 시작점인 좌꺽쇠(<)% 와 %우꺽쇠(>)는 New Line에 위치하여 
'	  서버 사이드 구문과 클라이언트 사이드 구문의 위치를 가늠할 수 있도록 한다.
'☜ : 아래 HTML 구문은 변경되어서는 안된다. 
%>
<!-- #Include file="../../inc/IncSvrMain.asp"  -->
<!-- #Include file="../../inc/incSvrDate.inc"  -->
<!-- #Include file="../../inc/IncSvrNumber.inc"  -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<% 
	Call LoadBasisGlobalInf() 
    Call LoadInfTB19029B("I", "*","NOCOOKIE","MB")
%>
<%					
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    Dim lgOpModeCRUD
    Call HideStatusWnd                                                               '☜: Hide Processing message
    '---------------------------------------Common-----------------------------------------------------------
	'------ Developer Coding part (Start ) ------------------------------------------------------------------

	' 권한관리 추가 
	Dim lgAuthBizAreaCd, lgAuthBizAreaNm			' 사업장 
	Dim lgInternalCd, lgDeptCd, lgDeptNm			' 내부부서 
	Dim lgSubInternalCd, lgSubDeptCd, lgSubDeptNm	' 내부부서(하위포함)
	Dim lgAuthUsrID, lgAuthUsrNm					' 개인 

	'------ Developer Coding part (End   ) ------------------------------------------------------------------ 
    lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)

	' 권한관리 추가 
	lgAuthBizAreaCd		= Trim(Request("lgAuthBizAreaCd"))
	lgInternalCd		= Trim(Request("lgInternalCd"))
	lgSubInternalCd		= Trim(Request("lgSubInternalCd"))
	lgAuthUsrID			= Trim(Request("lgAuthUsrID"))
    
    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '☜: Query
    
             Call SubBizQueryMulti()

        Case CStr(UID_M0002)                                                         '☜: Save,Update
    
             Call SubBizSaveMulti()

        Case CStr(UID_M0003)                                                         '☜: Delete
             Call SubBizDelete()
    End Select

'============================================================================================================
' Name : SubBizQueryMulti
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()
	On Error Resume Next                                                            '☜: Protect system from crashing
    Err.Clear                                                                       '☜: Clear Error status

	Dim iPAFG210																	' 입력/수정용 ComProxy Dll 사용 변수 
    Dim iStrData
    Dim exportData1
    Dim strDate,strYear,strMonth,strDay   
	
	Const C_EG1_bdg_cd_fr = 0
	Const C_EG1_bdg_nm_fr = 1
	Const C_EG1_bdg_cd_to = 2
	Const C_EG1_bdg_nm_to = 3
	Const C_EG1_dept_cd = 4
	Const C_EG1_dept_nm = 5
	Const C_EG1_bdg_yyyymm_fr=6
	Const C_EG1_bdg_yyyymm_to=7	
	
    Dim exportData2
    Dim iLngRow,iLngCol
    
    Dim iStrNextKey
    Dim iIntMaxRows
    Dim iIntQueryCount
    Dim iIntLoopCount
    
    Dim importArray
	Dim iAcctCd
	Dim iStrPrevBdgCdKey
	Dim iStrPrevBdgYMKey
	Dim iStrPrevDeptCdKey	
 
	Const C_SHEETMAXROWS_D  = 100

	Const C_yyyymm_from		= 0
	Const C_yyyymm_to		= 1
	Const C_dept_org_chg_id = 2
	Const C_dept_cd			= 3
	Const C_acct_cd_from	= 4
	Const C_acct_cd_to		= 5
    Const C_Next_bdg_cd_Key	= 6
    Const C_Next_bdg_ym_Key	= 7
    Const C_Next_dept_cd_Key= 8	
	
	'Const C_Next_Key2	= 8
	Const C_EG2_BDG_CD = 0
	Const C_EG2_BDG_NM = 1
	Const C_EG2_BDG_DT = 2
	Const C_EG2_DEPT_CD = 3
	Const C_EG2_DEPT_NM = 4
	Const C_EG2_ORG_CHG_ID  = 5
	Const C_EG2_ORG_CTRL_FG = 6
	Const C_EG2_BDG_PLAN_AMT  = 7
	Const C_EG2_BDG_AMT  = 8
	Const C_EG2_BDG_GL_AMT  = 9
	Const C_EG2_BDG_TEMP_AMT  = 10
	Const C_EG2_BDG_CTRL_FG  = 16
	Const C_EG2_BDG_CTRL_FG2  = 17
	
    Dim iYymmFr
    Dim iYymmTo

	
   
	'Key 값을 읽어온다	
	
	iStrPrevBdgCdKey	= Trim(Request("lgStrPrevBdgCdKey"))
	iStrPrevBdgYMKey	= Trim(Request("lgStrPrevBdgYMKey"))
	iStrPrevDeptCdKey	= Trim(Request("lgStrPrevDeptCdKey"))
	   
	iIntMaxRows		= Request("txtMaxRows")
    iIntQueryCount	= Request("lgPageNo")	
	
	Redim importArray(C_Next_dept_cd_Key+4)	
	importArray(C_yyyymm_from)		= Request("txtBdgYymmFr")
	importArray(C_yyyymm_to)		= Request("txtBdgYymmTo")
	importArray(C_dept_org_chg_id)	= Request("OrgChangeId")
	importArray(C_dept_cd)			= UCase(Trim(Request("txtDeptCd")))
	importArray(C_acct_cd_from)		= UCase(Trim(Request("txtBdgCdFr")))
	importArray(C_acct_cd_to)		= UCase(Trim(Request("txtBdgCdTo")))
	importArray(C_Next_bdg_cd_Key)	= iStrPrevBdgCdKey
	importArray(C_Next_bdg_ym_Key)	= iStrPrevBdgYMKey
	importArray(C_Next_dept_cd_Key)	= iStrPrevDeptCdKey

	' 권한관리 추가 
	importArray(C_Next_dept_cd_Key+1) = lgAuthBizAreaCd
	importArray(C_Next_dept_cd_Key+2) = lgInternalCd
	importArray(C_Next_dept_cd_Key+3) = lgSubInternalCd
	importArray(C_Next_dept_cd_Key+4) = lgAuthUsrID

	Dim i 
	
	If Len(Trim(iIntQueryCount))  Then                                        '☜ : Chnage Nextkey str into int value
		If Isnumeric(iIntQueryCount) Then
			iIntQueryCount = CInt(iIntQueryCount)          
		End If   
    Else   
		iIntQueryCount = 0
    End If
    
    If CheckSYSTEMError(Err, True) = True Then					
		Exit Sub
    End If  
    
    Set iPAFG210 = Server.CreateObject("PAFG210_LK0441.cFListBdgSvr")
    
	If CheckSYSTEMError(Err, True) = True Then					
		Exit Sub
    End If    
     
    Call iPAFG210.F_LIST_BDG_SVR(gStrGlobalCollection,C_SHEETMAXROWS_D,importArray, exportData1, exportData2)

	If CheckSYSTEMError(Err, True) = True Then					
		Set iPAFG210 = Nothing
		Exit Sub
    End If    
    
    Set IPAFG210 = nothing    
	
    iStrData = ""
    iIntLoopCount = 0	
	'2002/11/15 수정 
	If IsEmpty(exportData2) = False Then
		For iLngRow = 0 To UBound(exportData2, 1)
			iIntLoopCount = iIntLoopCount + 1

			If iIntLoopCount < (C_SHEETMAXROWS_D + 1) Then
                Call ExtractDateFrom(exportData2(iLngRow, C_EG2_BDG_DT), "YYYYMM", "", strYear, strMonth, strDay)
                strDate = UniConvYYYYMMDDToDate(gAPDateFormat, strYear, strMonth, "01")

                iStrData = iStrData & Chr(11) & ConvSPChars(exportData2(iLngRow, C_EG2_BDG_CD)) 'bdg_cd
                iStrData = iStrData & Chr(11) & ""
                iStrData = iStrData & Chr(11) & ConvSPChars(exportData2(iLngRow, C_EG2_BDG_NM)) 'bdg_name
                iStrData = iStrData & Chr(11) & UNIMonthClientFormat(ConvSPChars(strDate)) 'bdg_dt
                iStrData = iStrData & Chr(11) & ConvSPChars(exportData2(iLngRow, C_EG2_DEPT_CD)) 'dept_cd
                iStrData = iStrData & Chr(11) & ""
                iStrData = iStrData & Chr(11) & ConvSPChars(exportData2(iLngRow, C_EG2_DEPT_NM)) 'dept_nm
                iStrData = iStrData & Chr(11) & ConvSPChars(exportData2(iLngRow, C_EG2_ORG_CHG_ID)) 'org_chg_id
				iStrData = iStrData & Chr(11) & ConvSPChars(exportData2(iLngRow, C_EG2_ORG_CTRL_FG)) 'CTRL_FG
                iStrData = iStrData & Chr(11) & UNINumClientFormat(exportData2(iLngRow, C_EG2_BDG_PLAN_AMT), ggAmtOfMoney.DecPoint, 0)
                iStrData = iStrData & Chr(11) & UNINumClientFormat(exportData2(iLngRow, C_EG2_BDG_AMT), ggAmtOfMoney.DecPoint, 0)
                iStrData = iStrData & Chr(11) & UNINumClientFormat(exportData2(iLngRow, C_EG2_BDG_GL_AMT), ggAmtOfMoney.DecPoint, 0)
                iStrData = iStrData & Chr(11) & UNINumClientFormat(exportData2(iLngRow, C_EG2_BDG_TEMP_AMT), ggAmtOfMoney.DecPoint, 0)
                iStrData = iStrData & Chr(11) & ConvSPChars(exportData2(iLngRow, C_EG2_BDG_CTRL_FG)) 'BDG_CTRL_FG
                iStrData = iStrData & Chr(11) & ConvSPChars(exportData2(iLngRow, C_EG2_BDG_CTRL_FG2)) 'BDG_CTRL_FG

                iStrData = iStrData & Chr(11) & CStr(iIntMaxRows + iLngRow + 1) & Chr(11) & Chr(12)
			Else
                iStrPrevBdgCdKey	= exportData2(UBound(exportData2, 1), C_EG2_BDG_CD)
				iStrPrevBdgYMKey	= exportData2(UBound(exportData2, 1), C_EG2_BDG_DT)
				iStrPrevDeptCdKey	= exportData2(UBound(exportData2, 1), C_EG2_DEPT_CD)
	
                iIntQueryCount = iIntQueryCount + 1
				Exit For
			End If
		Next
	End If

	If  iIntLoopCount < (C_SHEETMAXROWS_D + 1) Then
		iStrPrevBdgCdKey = ""
		iStrPrevBdgYMKey = ""
		iStrPrevDeptCdKey = ""
		iIntQueryCount = ""
		    
	End If
	
	
	Response.Write " <Script Language=vbscript>	                         " & vbCr
	Response.Write " With parent                                         " & vbCr
    Response.Write " .ggoSpread.Source = .frm1.vspdData					 " & vbCr 			 
    Response.Write " .ggoSpread.SSShowData """ & iStrData		& """    " & vbCr
    Response.Write " .frm1.txtDeptCd.value = """ & ConvSPChars(Request("txtDeptCd"))   & """    " & vbcr
    Response.Write " .frm1.txtDeptNm.value = """ & exportData1(C_EG1_dept_nm)   & """    " & vbcr
    Response.Write " .frm1.txtBdgNmFr.value = """ & exportData1(C_EG1_bdg_nm_fr)       & """    " & vbCr
    Response.Write " .frm1.txtBdgNmTo.value = """ & exportData1(C_EG1_bdg_nm_to)      & """    " & vbCr
    Response.Write " .lgStrPrevBdgCdKey     = """ & iStrPrevBdgCdKey & """    " & vbCr
    Response.Write " .lgStrPrevBdgYMKey		= """ & iStrPrevBdgYMKey & """    " & vbCr					  
    Response.Write " .lgStrPrevDeptCdKey	= """ & iStrPrevDeptCdKey & """    " & vbCr	
    Response.Write " End With " & vbCr
	Response.Write " </Script> " & vbCr				  
%>        
    <script Language=vbscript >	
	With parent	
		If .frm1.vspdData.MaxRows < C_SHEETMAXROWS_D And lgStrPrevBdgCdKey <> "" and lgStrPrevBdgYMKey <> "" _
			and lgStrPrevDeptCdKey <> "" and lgStrPrevChgSeqKey <> "" Then	 
		Else			 
			.frm1.htxtBdgCdFr.value		= "<%=ConvSPChars(Request("txtBdgCdFr"))%>"
			.frm1.htxtBdgCdto.value		= "<%=ConvSPChars(Request("txtBdgCdto"))%>"
			.frm1.htxtBdgYymmFr.value	= "<%=ConvSPChars(Request("txtBdgYymmFr"))%>"
			.frm1.htxtBdgYymmTo.value	= "<%=ConvSPChars(Request("txtBdgYymmTo"))%>"
			.frm1.htxtDeptCd.value		= "<%=ConvSPChars(Request("txtDeptCd"))%>"
			Call .DbQueryOk
		End If
	End With
	</Script>		
<%
    
End Sub

'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data 
'============================================================================================================
Sub SubBizSaveMulti()

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
  
    Dim iPAFG210		
    Dim importString
    Dim I1_a_data_auth 
    Dim iErrorPosition
 
    Const A750_I1_a_data_auth_data_BizAreaCd = 0
    Const A750_I1_a_data_auth_data_internal_cd = 1
    Const A750_I1_a_data_auth_data_sub_internal_cd = 2
    Const A750_I1_a_data_auth_data_auth_usr_id = 3 
 
  	Redim I1_a_data_auth(3)
	I1_a_data_auth(A750_I1_a_data_auth_data_BizAreaCd)       = Trim(Request("txthAuthBizAreaCd"))
	I1_a_data_auth(A750_I1_a_data_auth_data_internal_cd)     = Trim(Request("txthInternalCd"))
	I1_a_data_auth(A750_I1_a_data_auth_data_sub_internal_cd) = Trim(Request("txthSubInternalCd"))
	I1_a_data_auth(A750_I1_a_data_auth_data_auth_usr_id)     = Trim(Request("txthAuthUsrID"))
 
    Set iPAFG210 = Server.CreateObject("PAFG210_LK0441.cFMngBdgSvr")

    If CheckSYSTEMError(Err, True) = True Then					
		Exit Sub
    End If    

    Call iPAFG210.F_MANAGE_BDG_SVR(gStrGlobalCollection, Trim(Request("txtSpread")),iErrorPosition,I1_a_data_auth)			
	
    If CheckSYSTEMError2(Err, True,iErrorPosition & "행","","","","") = True Then				
		Set iPAFG210 = Nothing
		Exit Sub
    End If    
    
    Set iPAFG210 = Nothing
	
	Response.Write " <Script Language=vbscript> " & vbCr
	Response.Write " parent.DbSaveOk            " & vbCr
    Response.Write "</Script>                   " & vbCr
End Sub    

%>
