<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% Option Explicit%>
<% session.CodePage=949 %>

<%
'**********************************************************************************************
'*  1. Module Name          : Accounting
'*  2. Function Name        : Fixed Asset Management
'*  3. Program ID           : a7105mb1(고정자산변동내역-매각/폐기)
'*  4. Program Name         :
'*  5. Program Desc         :
'*  6. Comproxy List        : +B19029LookupNumericFormat
'                             +B25011ManagePlant
'                             +B25011ManagePlant
'                             +B25018ListPlant
'                             +B25019LookUpPlant
'*  7. Modified date(First) : 2000/03/21
'*  8. Modified date(Last)  : 2001/06/02
'*  9. Modifier (First)     : Kim Hee Jung
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'*                          
'**********************************************************************************************
Response.Expires = -1								'☜ : ASP가 캐쉬되지 않도록 한다.
Response.Buffer = True								'☜ : ASP가 버퍼에 저장되지 않고 바로 Client에 내려간다.


'☜ : 항상 서버 사이드 구문의 시작점인 좌꺽쇠(<)% 와 %우꺽쇠(>)는 New Line에 위치하여 
'	  서버 사이드 구문과 클라이언트 사이드 구문의 위치를 가늠할 수 있도록 한다.
'☜ : 아래 HTML 구문은 변경되어서는 안된다. 
%>
<!-- #Include file="../../inc/incSvrMain.asp"  -->
<!-- #Include file="../../inc/incSvrDate.inc"  -->
<!-- #Include file="../../inc/incSvrNumber.inc"  -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<%													'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다
	On Error Resume Next														'☜: 
	Err.Clear	

	Dim LngMaxRow,LngMaxRow2
	Dim lgCurrency
	Dim lgCurrencyAcq
	Dim lgBlnFlgChgValue, lgOpModeCRUD, lgLngMaxRow, lgLngMaxRow2

	' 권한관리 추가
	Dim lgAuthBizAreaCd, lgAuthBizAreaNm			' 사업장
	Dim lgInternalCd, lgDeptCd, lgDeptNm			' 내부부서
	Dim lgSubInternalCd, lgSubDeptCd, lgSubDeptNm	' 내부부서(하위포함)
	Dim lgAuthUsrID, lgAuthUsrNm					' 개인

    Call HideStatusWnd                                                               '☜: Hide Processing message
    
    Call LoadBasisGlobalInf()
	Call LoadInfTB19029B("I","*","NOCOOKIE","MB")
	Call LoadBNumericFormatB("I", "*","NOCOOKIE","MB")
    '---------------------------------------Common-----------------------------------------------------------
'    lgErrorStatus     = "NO"
'    lgErrorPos        = ""                                                           '☜: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)
'    lgKeyStream       = Split(Request("txtKeyStream"),gColSep)
    'Single
'    lgPrevNext        = Request("txtPrevNext")                                       '☜: "P"(Prev search) "N"(Next search)    
    'Multi SpreadSheet
    lgLngMaxRow       = Request("txtMaxRows")                                        '☜: Read Operation Mode (CRUD)
    lgLngMaxRow2       = Request("txtMaxRows")                                        '☜: Read Operation Mode (CRUD)
'    lgMaxCount        = CInt(Request("lgMaxCount"))                                  '☜: Fetch count at a time for VspdData
'    lgStrPrevKeyIndex = UNICInt(Trim(Request("lgStrPrevKeyIndex")),0)                '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)

    Select Case lgOpModeCRUD
        Case CStr(UID_M0002)                                                         '☜: Save,Update
             Call SubBizSave()
        Case CStr(UID_M0003)                                                         '☜: Delete
             Call SubBizDelete()
    End Select

	Response.End    

'============================================================================================================
' Name : SubBizSave
' Desc : Date data 
'============================================================================================================
Sub SubBizSave()
	Dim I1_a_asset_chg_master 

	Dim iPAAG080
	
	Dim E1_a_asset_chg_master
    Const ICommandSent = 0
    Const I1_asst_chg_no = 1
    Const I1_chg_fg = 2
    Const I1_chg_dt = 3
    Const I1_dept_cd = 4
    Const I1_org_change_id = 5
    Const I1_loc_cur = 6
    Const I1_doc_cur = 7
    Const I1_xch_rate = 8
    Const I1_bp_cd = 9
    Const I1_asst_chg_desc = 10
    Const I1_gl_no = 11
    Const I1_temp_gl_no = 12
    Const I1_tax_type_cd = 13
    Const I1_tax_rate = 14
    Const I1_report_biz_area_cd = 15
    Const I1_issued_dt = 16
   
	Const E1_asst_chg_no = 0

    Dim I2_a_data_auth  '--> 파라미터의 순서에 따라 네이밍 변동
    Const C_I2_a_data_auth_data_BizAreaCd = 0
    Const C_I2_a_data_auth_data_internal_cd = 1
    Const C_I2_a_data_auth_data_sub_internal_cd = 2
    Const C_I2_a_data_auth_data_auth_usr_id = 3 
 
  	Redim I2_a_data_auth(3)
	I2_a_data_auth(C_I2_a_data_auth_data_BizAreaCd)       = Trim(Request("txthAuthBizAreaCd"))
	I2_a_data_auth(C_I2_a_data_auth_data_internal_cd)     = Trim(Request("txthInternalCd"))
	I2_a_data_auth(C_I2_a_data_auth_data_sub_internal_cd) = Trim(Request("txthSubInternalCd"))
	I2_a_data_auth(C_I2_a_data_auth_data_auth_usr_id)     = Trim(Request("txthAuthUsrID"))
    
    Redim I1_a_asset_chg_master(I1_issued_dt)

	'***************************************************************
	'                              SAVE
	'***************************************************************									
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status    
	
	Dim lgIntFlgMode

	lgIntFlgMode = CInt(Request("txtFlgMode"))									'☜: 저장시 Create/Update 판별 

	LngMaxRow    = CInt(Request("txtMaxRows"))	
	LngMaxRow2    = CInt(Request("txtMaxRows2"))	
	
'    gChangeOrgId = GetGlobalInf("gChangeOrgId")
    
    gChangeOrgId =Request("hORGCHANGEID")
    
    '-----------------------
    'Data manipulate area
    '-----------------------

    If lgIntFlgMode = OPMD_CMODE Then
		I1_a_asset_chg_master(iCommandSent) = "CREATE"
    ElseIf lgIntFlgMode = OPMD_UMODE Then
		I1_a_asset_chg_master(iCommandSent) = "UPDATE"
    End If
    I1_a_asset_chg_master(I1_asst_chg_no) = UCase(Trim(Request("txtAsstChgNo2")))
    I1_a_asset_chg_master(I1_chg_fg) = UCase(Trim(Request("txtRadio")))

    I1_a_asset_chg_master(I1_chg_dt) = UNIConvDate(Request("txtChgDt")) 
    I1_a_asset_chg_master(I1_dept_cd) = Trim(Request("txtDeptCd"))
    I1_a_asset_chg_master(I1_org_change_id) = gChangeOrgId
    I1_a_asset_chg_master(I1_loc_cur) = gCurrency
    I1_a_asset_chg_master(I1_doc_cur) = UCase(Request("txtDocCur"))
   	if UCase(Request("txtDocCur")) = gCurrency then        
		I1_a_asset_chg_master(I1_xch_rate)  = 1
	else
		I1_a_asset_chg_master(I1_xch_rate)  = UNIConvNum(Request("txtXchRate"),0)        '환율 
	end if			
    I1_a_asset_chg_master(I1_bp_cd) = UCase(Request("txtBpCd")) 
    I1_a_asset_chg_master(I1_asst_chg_desc) = Trim(Request("txtChgDesc"))
    I1_a_asset_chg_master(I1_gl_no) = UCase(Trim(Request("txtGlNo")))
    I1_a_asset_chg_master(I1_temp_gl_no) = UCase(Trim(Request("txtTempGlNo")))
    I1_a_asset_chg_master(I1_tax_type_cd) = UCase(Trim(Request("txtVatType")))
    I1_a_asset_chg_master(I1_tax_rate) = UNIConvNum(Request("txtVatRate"),0)
    I1_a_asset_chg_master(I1_report_biz_area_cd) = UCase(Trim(Request("txtReportAreaCd")))
    I1_a_asset_chg_master(I1_issued_dt) = UNIConvDate(Request("txtIssuedDt"))

	'-----------------------
	'Com Action Area
	'-----------------------

    Set iPAAG080 = Server.CreateObject("PAAG080.cAMngAsChgMas0304Svr") 
    
	If CheckSYSTEMError(Err,True) = True Then
 		Set iPAAG080 =nothing	
       Exit Sub
    End If 

	Call iPAAG080.A_MANAGE_ASSET_CHG_MASTER_0304_SVR( gStrGloBalCollection ,I1_a_asset_chg_master, Request("txtSpread"), _
															Request("txtSpread2"), E1_a_asset_chg_master) 

	'-----------------------
	'DB Error
	'-----------------------
 	If CheckSYSTEMError(Err,True) = True Then
 		Set iPAAG080 =nothing	
       Exit Sub
    End If 

   Set iPAAG080 = Nothing                                                  '☜: Unload 

	Response.Write "<Script Language=vbscript>				 " & vbCr
	Response.Write "With parent						" & vbCr
	Response.Write "  .frm1.txtAsstChgNo.Value=  """ & ConvSPChars(E1_a_asset_chg_master) & 				"""" & vbCr
	Response.Write "	.DbSaveOk " & vbCr  	' 이전 키 값 넘겨줌 , 현재 ComProxy가 제대로 안되 있음 															'☜: 조화가 성공 
	Response.Write "	End With		" & vbCr  
	Response.Write "</Script>		" & vbCr  

End Sub	
	    
'============================================================================================================
' Name : SubBizDelete
' Desc : Delete DB data
'============================================================================================================
Sub SubBizDelete()
	Dim I1_a_asset_chg_master 

	Dim iPAAG080
	
	Dim E1_a_asset_chg_master 	

    Const ICommandSent = 0
    Const I1_asst_chg_no = 1
    Const I1_chg_fg = 2
    Const I1_chg_dt = 3
    Const I1_dept_cd = 4
    Const I1_org_change_id = 5
    Const I1_loc_cur = 6
    Const I1_doc_cur = 7
    Const I1_xch_rate = 8
    Const I1_bp_cd = 9
    Const I1_asst_chg_desc = 10
    Const I1_gl_no = 11
    Const I1_temp_gl_no = 12
    Const I1_tax_type_cd = 13
    Const I1_tax_rate = 14
    Const I1_report_biz_area_cd = 15
    Const I1_issued_dt = 16
   
	Const E1_asst_chg_no = 0

    Dim I2_a_data_auth  '--> 파라미터의 순서에 따라 네이밍 변동
    Const C_I2_a_data_auth_data_BizAreaCd = 0
    Const C_I2_a_data_auth_data_internal_cd = 1
    Const C_I2_a_data_auth_data_sub_internal_cd = 2
    Const C_I2_a_data_auth_data_auth_usr_id = 3 
 
  	Redim I2_a_data_auth(3)
	I2_a_data_auth(C_I2_a_data_auth_data_BizAreaCd)       = Trim(Request("txthAuthBizAreaCd"))
	I2_a_data_auth(C_I2_a_data_auth_data_internal_cd)     = Trim(Request("txthInternalCd"))
	I2_a_data_auth(C_I2_a_data_auth_data_sub_internal_cd) = Trim(Request("txthSubInternalCd"))
	I2_a_data_auth(C_I2_a_data_auth_data_auth_usr_id)     = Trim(Request("txthAuthUsrID"))
    
    Redim I1_a_asset_chg_master(I1_issued_dt)

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to update record
    '--------------------------------------------------------------------------------------------------------
	'***************************************************************
	'                              DELETE
	'***************************************************************
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status    

    If Request("txtAsstChgNo2") = "" Then    	'⊙: 삭제를 위한 값이 들어왔는지 체크
		Call ServerMesgBox("700114", vbInformation, I_MKSCRIPT)			'삭제 조건값이 비어있습니다!
		Response.End 
	End If
	
    I1_a_asset_chg_master(ICommandSent) = "DELETE"
    I1_a_asset_chg_master(I1_asst_chg_no) = UCase(Trim(Request("txtAsstChgNo2")))
    
    Set iPAAG080 = Server.CreateObject("PAAG080.cAMngAsChgMas0304Svr") 
    
	If CheckSYSTEMError(Err,True) = True Then
 		Set iPAAG080 =nothing	
		Exit Sub
    End If 

	Call iPAAG080.A_MANAGE_ASSET_CHG_MASTER_0304_SVR( gStrGloBalCollection ,I1_a_asset_chg_master, "" , _
															"" , E1_a_asset_chg_master ) 
    '-----------------------
    'Com action result check area(OS,internal)
    '-----------------------
 	If CheckSYSTEMError(Err,True) = True Then
 		Set iPAAG080 =nothing	
		Exit Sub
    End If 

    Set iPAAG080 = Nothing                                                   '☜: Unload 
    
	Response.Write "<Script Language=vbscript>				 " & vbCr
	Response.Write "	Call parent.DbDeleteOk()		" & vbCr
	Response.Write "</Script>		" & vbCr 
	
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
End Sub
	
%>
