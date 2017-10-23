<%@ LANGUAGE="VBSCRIPT" %>
<%Option Explicit    %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Prucurement
'*  2. Function Name        : 
'*  3. Program ID           : M2111MB3
'*  4. Program Name         : 구매요청확정등록 
'*  5. Program Desc         : 구매요청확정등록 
'*  6. Component List       : PM2G148.cMLstReleasePurReqS / PM2G141.cMReleasePurReqS
'*  7. Modified date(First) : 2000/04/03
'*  8. Modified date(Last)  : 2003/05/23
'*  9. Modifier (First)     : Min, HJ
'* 10. Modifier (Last)      : Kang Su Hwan
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change" 
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************
-->
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->

<%	

call LoadBasisGlobalInf()
call LoadInfTB19029B("I", "*","NOCOOKIE","MB") 

    Dim lgOpModeCRUD
 
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    Call HideStatusWnd                                                               '☜: Hide Processing message
	'------ Developer Coding part (Start ) ------------------------------------------------------------------

	'------ Developer Coding part (End   ) ------------------------------------------------------------------ 

    lgOpModeCRUD  = Request("txtMode") 						     '☜: Read Operation Mode (CRUD)


'Call ServerMesgBox("CStr(UID_M0002)->" & CStr(UID_M0002) , vbInformation, I_MKSCRIPT)
'Call ServerMesgBox("lgOpModeCRUD->" & lgOpModeCRUD , vbInformation, I_MKSCRIPT)

    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '☜: Query
             Call  SubBizQueryMulti()
        Case CStr(UID_M0002)                                                         '☜: Save,Update
             Call SubBizSaveMulti()
        Case CStr(UID_M0003)                                                         '☜: Delete
    End Select


'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()	
	Dim iMax
	Dim PvArr
	Dim iPM2G148																	'☆ : 조회용 ComProxy Dll 사용 변수 

	Dim StrNextKey		' 다음 값 
	Dim lgStrPrevKey	' 이전 값 
	Dim iLngMaxRow		' 현재 그리드의 최대Row
	Dim iLngRow
	Dim GroupCount
	          
	Const C_SHEETMAXROWS_D  = 100

	' Com+ In/Out Interface Parameter
    Dim I1_b_pur_org_cd			' Query Condition : Purcahse Org. Code
    Dim I2_b_plant_cd			' Query Condition : Plant Code
    Dim I3_b_item_cd			' Query Condition : Item Code
    Dim I4_m_pur_req			' Query Condition : From Date
    Dim I5_m_pur_req			' Query Condition : To Date
    Dim I6_next_pr_no			' Query Condition : Next Query key 
    Dim I7_ief_supplied			' Query Condition : Rlease Flag
    Dim E1_b_plant				' Result Plant Data
    Dim E2_b_pur_org			' Result Org. Data
    Dim E3_next_pr_no			' Result Next Query Key
    Dim EG1_export_group		' Result Data
    Dim E4_b_item				' Result Item Data
	Dim iStrData				' make sheet 
    
	Const M037_I4_req_dt = 0    'I4_m_pur_req
    Const M037_I4_dlvy_dt = 1
    
    Const M037_I5_req_dt = 0    'I5_m_pur_req
    Const M037_I5_dlvy_dt = 1

    Const M037_E1_plant_cd = 0    'E1_b_plant
    Const M037_E1_plant_nm = 1

    Const M037_E2_pur_org = 0    'E2_b_pur_org
    Const M037_E2_pur_org_nm = 1

    Const M037_EG1_E1_plant_cd = 0    'EG1_export_group
    Const M037_EG1_E1_plant_nm = 1
    Const M037_EG1_E2_pr_no = 2    
    Const M037_EG1_E2_req_qty = 3
    Const M037_EG1_E2_req_unit = 4
    Const M037_EG1_E2_req_dt = 5
    Const M037_EG1_E2_req_prsn = 6
    Const M037_EG1_E2_dlvy_dt = 7
    Const M037_EG1_E2_pur_plan_dt = 8
    Const M037_EG1_E2_req_cfm_qty = 9
    Const M037_EG1_E2_pr_type = 10
    Const M037_EG1_E2_procure_type = 11
    Const M037_EG1_E2_pr_sts = 12
    Const M037_EG1_E2_ord_qty = 13
    Const M037_EG1_E2_rcpt_qty = 14
    Const M037_EG1_E2_iv_qty = 15
    Const M037_EG1_E2_req_dept = 16
    Const M037_EG1_E2_sppl_cd = 17
    Const M037_EG1_E2_sl_cd = 18
    Const M037_EG1_E2_pur_org = 19
    Const M037_EG1_E2_pur_grp = 20
    Const M037_EG1_E2_mrp_ord_no = 21
    Const M037_EG1_E2_mrp_run_no = 22
    Const M037_EG1_E2_tracking_no = 23
    Const M037_EG1_E2_base_req_qty = 24
    Const M037_EG1_E2_base_req_unit = 25
    Const M037_EG1_E2_so_no = 26
    Const M037_EG1_E2_so_seq_no = 27
    Const M037_EG1_E2_ext1_cd = 28
    Const M037_EG1_E2_ext1_qty = 29
    Const M037_EG1_E2_ext1_amt = 30
    Const M037_EG1_E2_ext1_rt = 31
    Const M037_EG1_E2_ext2_cd = 32
    Const M037_EG1_E2_ext2_qty = 33
    Const M037_EG1_E2_ext2_amt = 34
    Const M037_EG1_E2_ext2_rt = 35
    Const M037_EG1_E2_ext3_cd = 36
    Const M037_EG1_E2_ext3_qty = 37
    Const M037_EG1_E2_ext3_amt = 38
    Const M037_EG1_E2_ext3_rt = 39
    Const M037_EG1_E3_item_cd = 40   
    Const M037_EG1_E3_item_nm = 41
	Const M037_EG1_E3_item_spec = 42
    Const M037_EG1_E4_req_dept_nm = 43
	
    Const M037_E4_item_cd = 0    'E4_b_item
    Const M037_E4_item_nm = 1

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

	If Len(Trim(Request("txtFrDlvyDt"))) Then
		If UNIConvDate(Request("txtFrDlvyDt")) = "" Then
		    Call DisplayMsgBox("122116", vbInformation, "", "", I_MKSCRIPT)
		    Exit Sub
		End If
	End If
   
	If Len(Trim(Request("txtToDlvyDt"))) Then
		If UNIConvDate(Request("txtToDlvyDt")) = "" Then
		    Call DisplayMsgBox("122116", vbInformation, "", "", I_MKSCRIPT)
		    Exit Sub	
		End If
	End If
    
	If Len(Trim(Request("txtFrReqDt"))) Then
		If UNIConvDate(Request("txtFrReqDt")) = "" Then
		    Call DisplayMsgBox("122116", vbInformation, "", "", I_MKSCRIPT)
		    Exit Sub	
		End If
	End If
    
	If Len(Trim(Request("txtToReqDt"))) Then
		If UNIConvDate(Request("txtToReqDt")) = "" Then
		    Call DisplayMsgBox("122116", vbInformation, "", "", I_MKSCRIPT)
		    Exit Sub	
		End If
	End If
    
	lgStrPrevKey = Request("lgStrPrevKey")
 
'|-Coding Part**************************************************************************** 
    Set iPM2G148 = Server.CreateObject("PM2G148.cMLstReleasePurReqS")    
    
    '-----------------------
    'Com action result check area(OS,internal)
    '-----------------------
   If CheckSYSTEMError(Err,True) = true Then 		
		Set iPM2G148 = Nothing												'☜: ComProxy Unload
		Exit Sub														'☜: 비지니스 로직 처리를 종료함 
	
	End if
	
    '-----------------------
    'Data manipulate  area(import view match)
    '-----------------------
    I1_b_pur_org_cd 			= Trim(Request("txtOrgCd"))
    I2_b_plant_cd 				= Trim(Request("txtPlantCd"))
    I3_b_item_cd 				= Trim(Request("txtItemCd"))
    ReDim I4_m_pur_req(2)		' Query Condition : From Date
    ReDim I5_m_pur_req(2)		' Query Condition : To Date

    if Request("txtFrDlvyDt") = "" then
    	I4_m_pur_req(M037_I4_dlvy_dt) 	= "1900-01-01"
    else
    	I4_m_pur_req(M037_I4_dlvy_dt) 	= UNIConvDate(Request("txtFrDlvyDt"))
    End if 
    if Request("txtToDlvyDt") = "" then
    	I5_m_pur_req(M037_I5_dlvy_dt) 	= "2999-12-31"
    else
    	I5_m_pur_req(M037_I5_dlvy_dt) 	= UNIConvDate(Request("txtToDlvyDt"))
    End if
     
    if Request("txtFrReqDt") = "" then
    	I4_m_pur_req(M037_I4_req_dt) 	= "1900-01-01"
    else
    	I4_m_pur_req(M037_I4_req_dt) 	= UNIConvDate(Request("txtFrReqDt"))
    End if 
    if Request("txtToReqDt") = "" then
    	I5_m_pur_req(M037_I5_req_dt) 	= "2999-12-31"
    else 
    	I5_m_pur_req(M037_I5_req_dt) 	= UNIConvDate(Request("txtToReqDt"))
    End if 

    '|-+--값입력 **************
    I6_next_pr_no 			= Request("lgStrPrevKey")
    I7_ief_supplied			= Request("txtflg")
    
    '-----------------------
    'Com action area
    '-----------------------
    Call iPM2G148.M_LIST_RELEASE_PUR_REQ_SVR(gStrGlobalCollection, Clng(C_SHEETMAXROWS_D), I1_b_pur_org_cd, I2_b_plant_cd, _ 
				I3_b_item_cd, I4_m_pur_req, I5_m_pur_req, I6_next_pr_no, I7_ief_supplied, E1_b_plant, E2_b_pur_org, E3_next_pr_no, _
				EG1_export_group, E4_b_item)					
	'-----------------------
	'Com action result check area(OS,internal)
	'-----------------------
	If CheckSYSTEMError2(Err,True,"","","","","") = true then 		
		Set iPM2G148 = Nothing												'☜: Complus Unload
		Response.Write "<Script Language=vbscript>" & vbCr
		Response.Write "with parent" & vbCr
		Response.Write "	.frm1.txtOrgNm.value = """ & ConvSPChars(E2_b_pur_org(M037_E2_pur_org_nm)) & """" & vbCr
		Response.Write "	.frm1.txtPlantNm.value = """ & ConvSPChars(E1_b_plant(M037_E1_plant_nm))   & """" & vbCr
		Response.Write "	.frm1.txtItemNm.value = """ & ConvSPChars(E4_b_item(M037_E4_item_nm))      & """" & vbCr
		Response.Write "	.frm1.txtOrgCd.focus " & vbCr
		Response.Write "End With "   & vbCr
		Response.Write "</Script>"                  & vbCr
		Exit Sub															'☜: Terminate Biz. Logic
	End if

	Set iPM2G148 = Nothing												'☜: Complus Unload
	
	
	Response.Write "<Script Language=vbscript>" & vbCr
	Response.Write "with parent" & vbCr
	Response.Write "	.frm1.txtOrgNm.value = """ & ConvSPChars(E2_b_pur_org(M037_E2_pur_org_nm)) & """" & vbCr
	Response.Write "	.frm1.txtPlantNm.value = """ & ConvSPChars(E1_b_plant(M037_E1_plant_nm))   & """" & vbCr
	Response.Write "	.frm1.txtItemNm.value = """ & ConvSPChars(E4_b_item(M037_E4_item_nm))      & """" & vbCr
	Response.Write "End With "   & vbCr
    Response.Write "</Script>"                  & vbCr
    	
	iLngMaxRow = CInt(Request("txtMaxRows"))											'Save previous Maxrow                                                
	iMax = UBound(EG1_export_group,1)
	ReDim PvArr(iMax)
    
    For iLngRow = 0 To UBound(EG1_export_group,1) 
		
		StrNextKey = ConvSPChars(E3_next_pr_no)

		If iLngRow >= C_SHEETMAXROWS_D Then
			If ConvSPChars(E3_next_pr_no) = "" Then
				StrNextKey = ConvSPChars(EG1_export_group(iLngRow, M037_EG1_E2_pr_no))
			End If	
			Exit For	
		End If
		
		If UCase(EG1_export_group(iLngRow, M037_EG1_E2_pr_sts)) = "RQ" then
			istrData = istrData & Chr(11) & "0"	
		Else
			istrData = istrData & Chr(11) & "1"	
		End if
		istrData = istrData & Chr(11) & ""
		istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow, M037_EG1_E2_pr_no))
		istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow, M037_EG1_E1_plant_cd))
		istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow, M037_EG1_E1_plant_nm))
        istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow, M037_EG1_E3_item_cd))			
        istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow, M037_EG1_E3_item_nm))		
        istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow, M037_EG1_E3_item_spec))		
        istrData = istrData & Chr(11) & UNINumClientFormat(EG1_export_group(iLngRow, M037_EG1_E2_req_qty), ggQty.DecPoint,0)
        istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow, M037_EG1_E2_req_unit))	
        istrData = istrData & Chr(11) & UNIDateClientFormat(ConvSPChars(EG1_export_group(iLngRow ,M037_EG1_E2_dlvy_dt)))
        istrData = istrData & Chr(11) & UNIDateClientFormat(ConvSPChars(EG1_export_group(iLngRow, M037_EG1_E2_req_dt)))
        istrData = istrData & Chr(11) & ""													    
        istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow, M037_EG1_E2_req_dept))		
        istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow, M037_EG1_E4_req_dept_nm))	
        istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow, M037_EG1_E2_req_prsn))		
        istrData = istrData & Chr(11) & iLngMaxRow + iLngRow                             
        istrData = istrData & Chr(11) & Chr(12)               
		PvArr(iLngRow) = istrData
		istrData=""
    Next      
    istrData = Join(PvArr, "")

    Response.Write "<Script Language=vbscript>" & vbCr
	Response.Write "With parent" & vbCr												'☜: 화면 처리 ASP 를 지칭함 
    Response.Write "	.ggoSpread.Source          =  .frm1.vspdData " & vbCr
    Response.Write "	.ggoSpread.SSShowData        """ & istrData	& """" & vbCr	
    Response.Write "	.frm1.vspdData.Redraw = false " & vbCr
    Response.Write "	.frm1.vspdData.Redraw = True " & vbCr
    Response.Write "	.lgStrPrevKey              = """ & StrNextKey & """" & vbCr  
    Response.Write " .frm1.hdnOrg.value     = """ & ConvSPChars(Request("txtOrgCd")) & """" & vbCr
	Response.Write " .frm1.hdnPlant.value   = """ & ConvSPChars(Request("txtPlantCd"))   & """" & vbCr
	Response.Write " .frm1.hdnItem.value    = """ & ConvSPChars(Request("txtItemCd"))   & """" & vbCr
	Response.Write " .frm1.hdnflg.value     = """ & ConvSPChars(Request("txtflg"))   & """" & vbCr
	Response.Write " .frm1.hdnFrDDt.value   = """ & UNIDateClientFormat(Request("txtFrDlvyDt")) & """" & vbCr
	Response.Write " .frm1.hdnToDDt.value   = """ & UNIDateClientFormat(Request("txtToDlvyDt")) & """" & vbCr
	Response.Write " .frm1.hdnFrRDt.value   = """ & UNIDateClientFormat(Request("txtFrReqDt"))  & """" & vbCr
	Response.Write " .frm1.hdnToRDt.value   = """ & UNIDateClientFormat(Request("txtToReqDt"))  & """" & vbCr
    Response.Write " .DbQueryOk "		    	  & vbCr 
    Response.Write " .frm1.vspdData.focus "		  & vbCr 
    Response.Write "End With" & vbCr
    Response.Write "</Script>" & vbCr   
		
End Sub    
'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data 
'============================================================================================================
Sub SubBizSaveMulti()
	On Error Resume Next
    Err.Clear																		'☜: Protect system from crashing

	Dim iPM2G141																	'☆ : 입력/수정용 ComProxy Dll 사용 변수 
	Dim iErrorPosition
	Dim itxtSpread
    Dim itxtSpreadArr
    Dim itxtSpreadArrCount

    Dim iCUCount
    Dim ii

Call ServerMesgBox("SubBizSaveMulti Start", vbInformation, I_MKSCRIPT)
             
    itxtSpread = ""
             
    iCUCount = Request.Form("txtCUSpread").Count

    itxtSpreadArrCount = -1
             
    ReDim itxtSpreadArr(iCUCount)
             
    For ii = 1 To iCUCount
        itxtSpreadArrCount = itxtSpreadArrCount + 1
        itxtSpreadArr(itxtSpreadArrCount) = Request.Form("txtCUSpread")(ii)
    Next
    itxtSpread = Join(itxtSpreadArr,"")

Call ServerMesgBox("SubBizSaveMulti itxtSpread->" & itxtSpread, vbInformation, I_MKSCRIPT)

    Response.Write "<Script language=vbs> " & vbCr   
    Response.Write "Parent.RemovedivTextArea "      & vbCr   
    Response.Write "</Script> "      & vbCr   

Call ServerMesgBox("SubBizSaveMulti Server.CreateObject before 1" & itxtSpread, vbInformation, I_MKSCRIPT)

    Response.Write "<Script language=vbs> " & vbCr   
    Response.Write "Parent.RemovedivTextArea "      & vbCr   
    Response.Write "</Script> "      & vbCr   

Call ServerMesgBox("SubBizSaveMulti Server.CreateObject before 2" & itxtSpread, vbInformation, I_MKSCRIPT)

    Set iPM2G141 = Server.CreateObject("PM2G141.cMReleasePurReqS")   

Call ServerMesgBox("SubBizSaveMulti Server.CreateObject after" & itxtSpread, vbInformation, I_MKSCRIPT)
 
    '-----------------------
    'Com action result check area(OS,internal)
    '-----------------------
    If CheckSYSTEMError(Err,True) = true then
		Set iPM2G141 = Nothing 		
		Exit Sub
	End If

Call ServerMesgBox("SubBizSaveMulti iPM2G141.M_RELEASE_PUR_REQ_SVR Call Before" & itxtSpread, vbInformation, I_MKSCRIPT)
	
	Call iPM2G141.M_RELEASE_PUR_REQ_SVR(gStrGlobalCollection, itxtSpread,, iErrorPosition)

Call ServerMesgBox("SubBizSaveMulti iPM2G141.M_RELEASE_PUR_REQ_SVR Call after" & itxtSpread, vbInformation, I_MKSCRIPT)

	If CheckSYSTEMError2(Err, True, iErrorPosition & "행","","","","") = True Then
       Set iPM2G141 = Nothing
       call SheetFocus(iErrorPosition,1,I_MKSCRIPT)
       Exit Sub
	End If 
                          
    Set iPM2G141 = Nothing													 '☜: Unload Comproxy       
                                              
	Response.Write "<Script language=vbs> " & vbCr         
    Response.Write " Parent.DbSaveOk "      & vbCr							'☜: 화면 처리 ASP 를 지칭함 
    Response.Write "</Script> "           
        
End Sub    


'==============================================================================
' Function : SheetFocus
' Description : 에러발생시 Spread Sheet에 포커스줌 
'==============================================================================
Function SheetFocus(Byval lRow, Byval lCol, Byval iLoc)
	
	If Trim(lRow) = "" Then Exit Function
	If iLoc = I_INSCRIPT Then
		strHTML = "parent.frm1.vspdData.focus" & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.Row = " & lRow & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.Col = " & lCol & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.Action = 0" & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.SelStart = 0 " & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.SelLength = len(parent.frm1.vspdData.Text) " & vbCrLf
		Response.Write strHTML
	ElseIf iLoc = I_MKSCRIPT Then
		strHTML = "<" & "Script LANGUAGE=VBScript" & ">" & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.focus" & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.Row = " & lRow & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.Col = " & lCol & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.Action = 0" & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.SelStart = 0 " & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.SelLength = len(parent.frm1.vspdData.Text) " & vbCrLf
		strHTML = strHTML & "</" & "Script" & ">" & vbCrLf
		Response.Write strHTML
	End If
End Function
%>
