
<%@ LANGUAGE=VBSCript %>
<%Option Explicit%>

<%'======================================================================================================
'*  1. Module Name          : Accounting
'*  2. Function Name        : Treasury - Budget
'*  3. Program ID           : f2103mb1
'*  4. Program Name         : 예산추가등록 
'*  5. Program Desc         : Register of Budget Append
'*  6. Comproxy List        : FU0031, FU0038
'*  7. Modified date(First) : 2000.09.22
'*  8. Modified date(Last)  : 2002.06.27
'*  9. Modifier (First)     : You, So Eun
'* 10. Modifier (Last)      : Lee, Hye Young
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

'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 
    Call HideStatusWnd

    Dim iPAFG215										'  ComProxy Dll 사용 변수 

	' 권한관리 추가 
	Dim lgAuthBizAreaCd, lgAuthBizAreaNm			' 사업장 
	Dim lgInternalCd, lgDeptCd, lgDeptNm			' 내부부서 
	Dim lgSubInternalCd, lgSubDeptCd, lgSubDeptNm	' 내부부서(하위포함)
	Dim lgAuthUsrID, lgAuthUsrNm					' 개인 

    Dim lgOpModeCRUD
          
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    Call HideStatusWnd                                                               '☜: Hide Processing message

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
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()
    
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
	
	Dim iLngRow,iLngCol    
    Dim iStrNextKey

	Dim iStrPrevBdgCdKey
	Dim iStrPrevBdgYMKey
	Dim iStrPrevDeptCdKey
	Dim iStrPrevChgSeqKey
	
	Dim iIntMaxRows
	dim lgLngMaxRow
	Dim iIntQueryCount
	Dim iIntLoopCount
	
	Dim importArray
    Dim exportData1
    Dim exportData2
    Dim iStrData    
    
    Dim strDate,strYear,strMonth,strDay
    
	Const C_SHEETMAXROWS_D  = 100
	
	Const C_count           = 0    
	Const C_dept_org_chg_id = 1
	Const C_dept_cd         = 2
	Const C_acct_cd_fr      = 3
	Const C_acct_cd_to      = 4
	Const C_yyyymm_fr       = 5
	Const C_yyyymm_to       = 6
	Const C_chg_fg          = 7
	Const C_chg_seq         = 8
	Const C_max_Rows        = 9
	Const C_Next_bdg_cd_Key = 10
	Const C_Next_bdg_ym_Key = 11
	Const C_Next_dept_cd_Key= 12
	Const C_Next_chg_seq_Key= 13	

    dim iYymmFr
    dim iYymmTo
   
   	
   	iStrPrevBdgCdKey = Trim(Request("lgStrPrevBdgCdKey"))
   	iStrPrevBdgYMKey = Trim(Request("lgStrPrevBdgYMKey"))   
   	iStrPrevDeptCdKey = Trim(Request("lgStrPrevDeptCdKey"))   
   	iStrPrevChgSeqKey = Trim(Request("lgStrPrevChgSeqKey"))   
   	
   	   
	iIntMaxRows     = Request("txtMaxRows")
	iIntQueryCount  = Request("lgPageNo")
	
	'lgLngMaxRow       = CInt(C_SHEETMAXROWS_D)
   
	Redim importArray(C_Next_chg_seq_Key+4)		'Key 값을 읽어온다	

    importArray(C_count)			= "" 'iStrPrevBdgCdKey
    importArray(C_dept_org_chg_id)	= GetGlobalInf("gChangeOrgId")
    importArray(C_dept_cd)			= UCase(Trim(Request("txtDeptCd")))
    importArray(C_acct_cd_fr)		= UCase(Trim(Request("txtBdgCdFr")))
    importArray(C_acct_cd_to)		= UCase(Trim(Request("txtBdgCdTo")))
    importArray(C_yyyymm_fr)		= Request("txtBdgYymmFr")
    importArray(C_yyyymm_to)		= Request("txtBdgYymmTo")
    importArray(C_chg_fg)			= "A"
    importArray(C_chg_seq)			= 0
    importArray(C_max_rows)			= CInt(C_SHEETMAXROWS_D)
    importArray(C_Next_bdg_cd_Key)	= iStrPrevBdgCdKey
    importArray(C_Next_bdg_ym_Key)	= iStrPrevBdgYMKey
    importArray(C_Next_dept_cd_Key)	= iStrPrevDeptCdKey
    importArray(C_Next_chg_seq_Key)	= iStrPrevChgSeqKey
    
	' 권한관리 추가 
	importArray(C_Next_chg_seq_Key+1) = lgAuthBizAreaCd
	importArray(C_Next_chg_seq_Key+2) = lgInternalCd
	importArray(C_Next_chg_seq_Key+3) = lgSubInternalCd
	importArray(C_Next_chg_seq_Key+4) = lgAuthUsrID
    
    If CheckSYSTEMError(Err, True) = True Then					
       Exit Sub
    End If 
    
    Set iPAFG215 = Server.CreateObject("PAFG215.cFListBdgChgSvr")
    
	
	If CheckSYSTEMError(Err, True) = True Then					
       Exit Sub
    End If    
     
    Call iPAFG215.F_LIST_BDG_CHG_SVR(gStrGlobalCollection,importArray, exportData1, exportData2)
	
	
	If CheckSYSTEMError(Err, True) = True Then			
	
       Set iPAFG215 = Nothing
       Exit Sub
    End If    
 
    
    Set IPAFG215 = nothing    
	
    iStrData = ""
    iIntLoopCount = 0	
		


'2002/11/15 수정 
    If IsEmpty(exportData2) = False Then
		For iLngRow = 0 To UBound(exportData2, 1)
			iIntLoopCount = iIntLoopCount + 1
			If iIntLoopCount < (C_SHEETMAXROWS_D + 1) Then
				Call ExtractDateFrom(exportData2(iLngRow,2),"YYYYMM","",strYear,strMonth,strDay)
				StrDate = UniConvYYYYMMDDToDate(gAPDateFormat,strYear,strMonth,"01")
	
	
				iStrData = iStrData & Chr(11) & ConvSPChars(exportData2(iLngRow,0))								
				iStrData = iStrData & Chr(11) & ""
				iStrData = iStrData & Chr(11) & ConvSPChars(exportData2(iLngRow,1))
				iStrData = iStrData & Chr(11) & UNIMonthClientFormat(ConvSPChars(strDate))
				iStrData = iStrData & Chr(11) & ConvSPChars(exportData2(iLngRow,3))
			    iStrData = iStrData & Chr(11) & ""
				iStrData = iStrData & Chr(11) & ConvSPChars(exportData2(iLngRow,4))
				iStrData = iStrData & Chr(11) & ConvSPChars(exportData2(iLngRow,5))
				iStrData = iStrData & Chr(11) & exportData2(iLngRow,6)
				iStrData = iStrData & Chr(11) & UNINumClientFormat(exportData2(iLngRow,7),ggAmtOfMoney.DecPoint, 0)				
				iStrData = iStrData & Chr(11) & ConvSPChars(Trim(exportData2(iLngRow,9)))		'BDG_CTRL_UNIT 11,				
				iStrData = iStrData & Chr(11) & exportData2(iLngRow,13)								'ADD_FG      12				
				iStrData = iStrData & Chr(11) & UNIDateClientFormat(exportData2(iLngRow,8))	'MG_DT
				'~~~2003-02-19 BDG_CHG_DESC 추가 
				iStrData = iStrData & Chr(11) & ConvSPChars(exportData2(iLngRow,24))			'BDG_CHG_DESC				
				iStrData = iStrData & Chr(11) & ConvSPChars(Trim(exportData2(iLngRow,22)))    'BDG_CTRL_FG 22
				iStrData = iStrData & Chr(11) & ConvSPChars(Trim(exportData2(iLngRow,23)))    'GP_BDG_CTRL_FG 23
                iStrData = iStrData & Chr(11) & CStr(iIntMaxRows + iLngRow + 1) & Chr(11) & Chr(12)
			Else
                iStrPrevBdgCdKey = exportData2(UBound(exportData2,1),0)      
				iStrPrevBdgYMKey = exportData2(UBound(exportData2,1),2)
				iStrPrevDeptCdKey = exportData2(UBound(exportData2,1),3)
				iStrPrevChgSeqKey = exportData2(UBound(exportData2,1),6)
                iIntQueryCount = iIntQueryCount + 1
				Exit For
			End If
		Next
    End If

	If iIntLoopCount < (C_SHEETMAXROWS_D + 1) Then
		iStrPrevBdgCdKey = ""
		iStrPrevBdgYMKey = ""
		iStrPrevDeptCdKey = ""
		iStrPrevChgSeqKey = ""
		iIntQueryCount = ""
	End If

	Response.Write " <Script Language=vbscript>	" & vbCr
	Response.write " With parent				" & vbCr
	Response.Write " .ggoSpread.Source = .frm1.vspdData					 " & vbCr 			 
    Response.Write " .ggoSpread.SSShowData """ & iStrData		& """    " & vbCr
   
	Response.Write " .frm1.txtDeptNm.value  = """ & ConvSPChars(exportData1(5)) & """ " & vbCr
	Response.Write " .frm1.txtBdgNmFr.value = """ & ConvSPChars(exportData1(1)) & """ " & vbCr
	Response.Write " .frm1.txtBdgNmTo.value = """ & exportData1(3) & """ " & vbCr
	Response.Write " .lgStrPrevBdgCdKey          = """ & iStrPrevBdgCdKey & """    " & vbCr
	Response.Write " .lgStrPrevBdgYMKey          = """ & iStrPrevBdgYMKey & """    " & vbCr		
	Response.Write " .lgStrPrevDeptCdKey         = """ & iStrPrevDeptCdKey & """    " & vbCr		
	Response.Write " .lgStrPrevChgSeqKey         = """ & iStrPrevChgSeqKey & """    " & vbCr				
	Response.Write " End With " & vbCr
	Response.Write " </Script> " & vbCr

	
%>
	<script Language=vbscript >	
	with parent	
		If .frm1.vspdData.MaxRows < C_SHEETMAXROWS_D And lgStrPrevBdgCdKey <> "" and lgStrPrevBdgYMKey <> "" _
			and lgStrPrevDeptCdKey <> "" and lgStrPrevChgSeqKey <> "" Then	 
	
		Else
			 
			.frm1.htxtBdgCdFr.value		= "<%=ConvSPChars(Request("txtBdgCdFr"))%>"
			.frm1.htxtBdgCdTo.value		= "<%=ConvSPChars(Request("txtBdgCdTo"))%>"
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
  
    Dim iPAFG215		
    Dim importString 
    Dim iErrorPosition
  
    Dim I2_a_data_auth  
    Const A748_I2_a_data_auth_data_BizAreaCd = 0
    Const A748_I2_a_data_auth_data_internal_cd = 1
    Const A748_I2_a_data_auth_data_sub_internal_cd = 2
    Const A748_I2_a_data_auth_data_auth_usr_id = 3 
 
  	Redim I2_a_data_auth(3)
	I2_a_data_auth(A748_I2_a_data_auth_data_BizAreaCd)       = Trim(Request("txthAuthBizAreaCd"))
	I2_a_data_auth(A748_I2_a_data_auth_data_internal_cd)     = Trim(Request("txthInternalCd"))
	I2_a_data_auth(A748_I2_a_data_auth_data_sub_internal_cd) = Trim(Request("txthSubInternalCd"))
	I2_a_data_auth(A748_I2_a_data_auth_data_auth_usr_id)     = Trim(Request("txthAuthUsrID"))
 
    Set iPAFG215 = Server.CreateObject("PAFG215.cFMngBdgChgSvr")

    If CheckSYSTEMError(Err, True) = True Then					
       Exit Sub
    End If    
	
    Call iPAFG215.F_MANAGE_BDG_CHG_SVR(gStrGlobalCollection, Trim(Request("txtSpread")),"A",iErrorPosition,I2_a_data_auth)						
     
    If CheckSYSTEMError2(Err, True,iErrorPosition & "행","","","","") = True Then				
       Set iPAFG215 = Nothing
		Exit Sub
    End If    
    
    Set iPAFG215 = Nothing
	
	Response.Write " <Script Language=vbscript> " & vbCr
	Response.Write " parent.DbSaveOk            " & vbCr
    Response.Write "</Script>                   " & vbCr
    
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



%>


