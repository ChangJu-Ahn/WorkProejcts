<%@ LANGUAGE="VBSCRIPT" %>
<%Option Explicit    %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Prucurement
'*  2. Function Name        : 
'*  3. Program ID           : mc710mb1
'*  4. Program Name         : 납입지시마감취소 
'*  5. Program Desc         : 납입지시마감취소 
'*  6. Component List       : PMCG700.cMListDvlyOrdCls / 
'*  7. Modified date(First) : 2003/03/13
'*  8. Modified date(Last)  : 2003/05/23
'*  9. Modifier (First)     : Kang Su Hwan
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
call LoadBNumericFormatB("I", "*","NOCOOKIE","MB")

    Dim lgOpModeCRUD
	Const PGM_FLG = "CANCEL"
	
    On Error Resume Next                                                            '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    Call HideStatusWnd                                                               '☜: Hide Processing message

    lgOpModeCRUD  = Request("txtMode") 
										                                              '☜: Read Operation Mode (CRUD)
    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '☜: Query
             Call  SubBizQueryMulti()
        Case CStr(UID_M0002)                                                         '☜: Save,Update
             Call SubBizSaveMulti()
    End Select

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()

    On Error Resume Next                                                             '☜: Protect system from crashing
    	
	Dim iPMCG700																	'☆ : 조회용 ComProxy Dll 사용 변수 
	Dim iMax
	Dim PvArr
	Dim StrNextKey		' 다음 값 
	Dim lgStrPrevKey	' 이전 값 
	Dim iLngMaxRow		' 현재 그리드의 최대Row
	Dim iLngRow
	Dim GroupCount      
	Dim arrValue    
	Dim istrData
	 
	Const C_SHEETMAXROWS_D  = 100
	
	Dim I1_select_flag
	Dim I2_m_dlvy_ord
	Const M805_I2_plant_cd = 0
	Const M805_I2_bp_cd = 1
	Const M805_I2_do_start_dt = 2
	Const M805_I2_do_end_dt = 3
	Const M805_I2_item_cd = 4
	
	Dim I3_m_dlvy_ord_next
	Const M805_I3_prodt_order_no = 0
	Const M805_I3_opr_no = 1
	Const M805_I3_seq = 2
	Const M805_I3_sub_seq = 3

	Dim EG_m_dlvy_ord
	Const M805_EG1_prodt_order_no = 0
	Const M805_EG1_opr_no = 1
	Const M805_EG1_seq = 2
	Const M805_EG1_sub_seq = 3
	Const M805_EG1_item_cd = 4
	Const M805_EG1_item_nm = 5
	Const M805_EG1_spec = 6
	Const M805_EG1_req_dt = 7
	Const M805_EG1_basic_unit = 8
	Const M805_EG1_do_qty = 9
	Const M805_EG1_rcpt_qty = 10
	Const M805_EG1_bp_cd = 11
	Const M805_EG1_bp_nm = 12
	Const M805_EG1_tracking_no = 13
	Const M805_EG1_po_no = 14
	Const M805_EG1_po_seq_no = 15
	Const M805_EG1_do_status = 16
	Const M805_EG1_do_stauts_desc = 17
	
	Redim I2_m_dlvy_ord(M805_I2_item_cd)
	Redim I3_m_dlvy_ord_next(M805_I3_sub_seq)
	
	If Len(Trim(Request("txtFrDt"))) Then
		If UNIConvDate(Request("txtFrDt")) = "" Then
		    Call DisplayMsgBox("122116", vbInformation, "", "", I_MKSCRIPT)
		    Call LoadTab("parent.frm1.txtFrDt", 0, I_MKSCRIPT)
		    Exit Sub	
		End If
	End If
	If Len(Trim(Request("txtToDt"))) Then
		If UNIConvDate(Request("txtToDt")) = "" Then
		    Call DisplayMsgBox("122116", vbInformation, "", "", I_MKSCRIPT)
		    Call LoadTab("parent.frm1.txtToDt", 0, I_MKSCRIPT)
		    Exit Sub	
		End If
	End If

	lgStrPrevKey = Request("lgStrPrevKey")
    If lgStrPrevKey <> "" then	
        arrValue = Split(lgStrPrevKey, gColSep)		
		I3_m_dlvy_ord_next(M805_I3_prodt_order_no)	= arrValue(0)
		I3_m_dlvy_ord_next(M805_I3_opr_no)			= arrValue(1)
		I3_m_dlvy_ord_next(M805_I3_seq)				= cInt(arrValue(2))
		I3_m_dlvy_ord_next(M805_I3_sub_seq)			= cInt(arrValue(3))
	else			
		I3_m_dlvy_ord_next(M805_I3_prodt_order_no)	= ""
		I3_m_dlvy_ord_next(M805_I3_opr_no)			= ""
		I3_m_dlvy_ord_next(M805_I3_seq)				= 0
		I3_m_dlvy_ord_next(M805_I3_sub_seq)			= 0
	End If	
  
    Set iPMCG700 = Server.CreateObject("PMCG700.cMListDvlyOrdCls")    

    '-----------------------
    'Com action result check area(OS,internal)
    '-----------------------
    If CheckSYSTEMError(Err,True) = true Then 		
		Set iPMCG700 = Nothing												'☜: ComProxy Unload
		Exit Sub														'☜: 비지니스 로직 처리를 종료함 
	End if
    
    '-----------------------
    'Data manipulate  area(import view match)
    '-----------------------
    I2_m_dlvy_ord(M805_I2_plant_cd) = Trim(Request("txtPlantCd"))
    
    If Trim(Request("txtSupplier")) <> "" Then
    	I2_m_dlvy_ord(M805_I2_bp_cd)				= Trim(Request("txtSupplier"))
    End If
    If Trim(Request("txtFrDt")) = "" Then
    	I2_m_dlvy_ord(M805_I2_do_start_dt)			= "1900-01-01"
    Else
    	I2_m_dlvy_ord(M805_I2_do_start_dt)			= UniConvDate(Request("txtFrDt"))
    End If 
    If Trim(Request("txtToDt")) = "" Then
    	I2_m_dlvy_ord(M805_I2_do_end_dt)			= "2999-12-31"
    Else
    	I2_m_dlvy_ord(M805_I2_do_end_dt)			= UniConvDate(Request("txtToDt"))
    End If
    If Trim(Request("txtItemCd")) <> "" Then
    	I2_m_dlvy_ord(M805_I2_item_cd)				= Trim(Request("txtItemCd"))
    End If

    Call iPMCG700.M_LIST_DLVY_ORD_CLS(gStrGlobalCollection, C_SHEETMAXROWS_D, PGM_FLG,I2_m_dlvy_ord, I3_m_dlvy_ord_next,EG_m_dlvy_ord)				
	
	If CheckSYSTEMError(Err,True) = true Then 		
		Set iPMCG700 = Nothing												'☜: ComProxy Unload
		Response.Write "<Script Language=VBScript>" & vbCr
		Response.Write "	Parent.frm1.txtPlantCd.focus " & vbCr
		Response.Write "	Set Parent.gActiveElement = Parent.document.activeElement    " & vbCr
		Response.Write "</Script>" & vbCr
		Exit Sub														'☜: 비지니스 로직 처리를 종료함 
	End if
	
	iLngMaxRow = Request("txtMaxRows")											'Save previous Maxrow                                                
    GroupCount = UBound(EG_m_dlvy_ord,1) 
	ReDim PvArr(GroupCount)
    
	'-----------------------
	'Result data display area
	'----------------------- 
	For iLngRow = 0 To UBound(EG_m_dlvy_ord,1)
		If  iLngRow < C_SHEETMAXROWS_D  Then
		Else
		   StrNextKey = ConvSPChars(EG_m_dlvy_ord(M805_EG1_prodt_order_no)) & gColSep & ConvSPChars(EG_m_dlvy_ord(M805_EG1_opr_no)) & gColSep & ConvSPChars(EG_m_dlvy_ord(M805_EG1_seq)) & gColSep & ConvSPChars(EG_m_dlvy_ord(M805_EG1_sub_seq))
           Exit For
        End If  
		
		istrData = istrData & Chr(11) & "0"
		istrData = istrData & Chr(11) & ConvSPChars(EG_m_dlvy_ord(iLngRow,M805_EG1_prodt_order_no))
		istrData = istrData & Chr(11) & ConvSPChars(EG_m_dlvy_ord(iLngRow,M805_EG1_item_cd))
		istrData = istrData & Chr(11) & ConvSPChars(EG_m_dlvy_ord(iLngRow,M805_EG1_item_nm))
		istrData = istrData & Chr(11) & ConvSPChars(EG_m_dlvy_ord(iLngRow,M805_EG1_spec))
		istrData = istrData & Chr(11) & UNIDateClientFormat(ConvSPChars(EG_m_dlvy_ord(iLngRow,M805_EG1_req_dt)))
		istrData = istrData & Chr(11) & UNINumClientFormat(EG_m_dlvy_ord(iLngRow,M805_EG1_do_qty),ggQty.DecPoint,0)	
		istrData = istrData & Chr(11) & UNINumClientFormat(EG_m_dlvy_ord(iLngRow,M805_EG1_rcpt_qty),ggQty.DecPoint,0)	
		istrData = istrData & Chr(11) & ConvSPChars(EG_m_dlvy_ord(iLngRow,M805_EG1_basic_unit))
		istrData = istrData & Chr(11) & ConvSPChars(EG_m_dlvy_ord(iLngRow,M805_EG1_bp_cd))
		istrData = istrData & Chr(11) & ConvSPChars(EG_m_dlvy_ord(iLngRow,M805_EG1_bp_nm))
		istrData = istrData & Chr(11) & ConvSPChars(EG_m_dlvy_ord(iLngRow,M805_EG1_tracking_no))
		istrData = istrData & Chr(11) & ConvSPChars(EG_m_dlvy_ord(iLngRow,M805_EG1_opr_no))
		istrData = istrData & Chr(11) & ConvSPChars(EG_m_dlvy_ord(iLngRow,M805_EG1_po_no))
		istrData = istrData & Chr(11) & ConvSPChars(EG_m_dlvy_ord(iLngRow,M805_EG1_po_seq_no))
		istrData = istrData & Chr(11) & ConvSPChars(EG_m_dlvy_ord(iLngRow,M805_EG1_seq))
		istrData = istrData & Chr(11) & ConvSPChars(EG_m_dlvy_ord(iLngRow,M805_EG1_sub_seq))
        istrData = istrData & Chr(11) & iLngMaxRow + iLngRow                             
        istrData = istrData & Chr(11) & Chr(12)       
        
		PvArr(iLngRow) = istrData
		istrData=""
    Next  
    istrData = Join(PvArr, "")
    
    Response.Write "<Script Language=vbscript>" & vbCr
	Response.Write "With parent" & vbCr												'☜: 화면 처리 ASP 를 지칭함 
    Response.Write "	.ggoSpread.Source        =  .frm1.vspdData "         & vbCr
    Response.Write "	.ggoSpread.SSShowData      """ & istrData	    & """" & vbCr	
    Response.Write "	.lgStrPrevKey           = """ & ConvSPChars(StrNextKey)				& """" & vbCr 
    Response.Write " .frm1.hdnPlantCd.value		= """ & ConvSPChars(Request("txtPlantCd"))	& """" & vbCr
	Response.Write " .frm1.hdnSupplier.value	= """ & ConvSPChars(Request("txtSupplier")) & """" & vbCr
	Response.Write " .frm1.hdnFrDt.value		= """ & ConvSPChars(Request("txtFrDt"))		& """" & vbCr
	Response.Write " .frm1.hdnToDt.value		= """ & ConvSPChars(Request("txtToDt"))		& """" & vbCr
	Response.Write " .frm1.hdnItemCd.value		= """ & ConvSPChars(Request("txtItemCd"))	& """" & vbCr
    Response.Write " .DbQueryOk "		    	  & vbCr 
    Response.Write " .frm1.vspdData.focus "		  & vbCr 
    Response.Write "End With" & vbCr
    Response.Write "</Script>" & vbCr


    Set iPMCG700 = Nothing
		
End Sub    
'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data 
'============================================================================================================
Sub SubBizSaveMulti()
	On Error Resume Next
    Err.Clear																		'☜: Protect system from crashing
	
	Dim iPMCG750																	'☆ : 입력/수정용 ComProxy Dll 사용 변수 
    Dim iErrorPosition
	Dim itxtSpread
    Dim itxtSpreadArr
    Dim itxtSpreadArrCount

    Dim iCUCount
    Dim ii
             
    itxtSpread = ""
             
    iCUCount = Request.Form("txtCUSpread").Count
             
    itxtSpreadArrCount = -1
             
    ReDim itxtSpreadArr(iCUCount)
             
    For ii = 1 To iCUCount
        itxtSpreadArrCount = itxtSpreadArrCount + 1
        itxtSpreadArr(itxtSpreadArrCount) = Request.Form("txtCUSpread")(ii)
    Next
    itxtSpread = Join(itxtSpreadArr,"")

    Response.Write "<Script language=vbs> " & vbCr   
    Response.Write "Parent.RemovedivTextArea "      & vbCr   
    Response.Write "</Script> "      & vbCr   
                                                
    Set iPMCG750 = Server.CreateObject("PMCG750.cMManageDlvyOrdCls")    

    '-----------------------
    'Com action result check area(OS,internal)
    '-----------------------
    If CheckSYSTEMError(Err,True) = true Then 		
		Set iPMCG750 = Nothing												'☜: ComProxy Unload
		Exit Sub														'☜: 비지니스 로직 처리를 종료함 
	
	End if
	
	Call iPMCG750.M_MANAGE_DLVY_ORD_CLS(gStrGlobalCollection, PGM_FLG,itxtSpread, iErrorPosition) 
                   
    If CheckSYSTEMError2(Err,True,iErrorPosition & "행:","","","","") = True Then
       Set iPMCG750 = Nothing
	   Call SheetFocus(iErrorPosition, 1, I_MKSCRIPT)
       Exit Sub
	End If

    Set iPMCG750 = Nothing                                                   '☜: Unload Comproxy
    
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
