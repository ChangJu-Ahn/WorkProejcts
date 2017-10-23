<%@ LANGUAGE="VBSCRIPT" %>
<%Option Explicit    %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Prucurement
'*  2. Function Name        : 
'*  3. Program ID           : mc100mb1
'*  4. Program Name         : 납입지시대상선정 
'*  5. Program Desc         : 납입지시대상선정 
'*  6. Component List       : PMCG100.cMListDlvyOrder / PP4C210.cPMngDlvyOrd
'*  7. Modified date(First) : 2003-04-08
'*  8. Modified date(Last)  : 2003/05/23
'*  9. Modifier (First)     : Ahn Jung Je
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
        Case CStr(UID_M0003)                                                         '☜: Delete
    '         Call SubBizDelete()
	
    End Select

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()
    On Error Resume Next                                                             '☜: Protect system from crashing
    	
	Dim PMCG100																	'☆ : 조회용 ComProxy Dll 사용 변수 
	
	Dim StrNextKey		' 다음 값 
	Dim lgStrPrevKey	' 이전 값 
	Dim iLngMaxRow		' 현재 그리드의 최대Row
	Dim iLngRow
	Dim GroupCount      
	Dim arrValue    
	Dim istrData
	Dim PvArr
	 
	Const C_SHEETMAXROWS_D  = 100
	   
    Dim I1_m_Dlvy_ord
		Const M801_I1_plant_cd = 0
		Const M801_I1_start_req_dt = 1
		Const M801_I1_end_req_dt = 2
		Const M801_I1_start_prodt_no = 3
		Const M801_I1_end_prodt_no = 4
		Const M801_I1_item_cd = 5
	ReDim I1_m_Dlvy_ord(5)
	
    Dim I2_m_Dlvy_ord_next
		Const M801_I2_opr_no = 0
		Const M801_I2_seq = 1
		Const M801_I2_sub_seq = 2
    ReDim I2_m_Dlvy_ord_next(2)
       
    Dim EG1_m_Dlvy_ord
		Const M801_EG1_prodt_order_no = 0
		Const M801_EG1_item_cd = 1
		Const M801_EG1_item_nm = 2
		Const M801_EG1_spec = 3
		Const M801_EG1_req_dt = 4
		Const M801_EG1_basic_unit = 5
		Const M801_EG1_req_qty = 6
		Const M801_EG1_base_qty = 7
		Const M801_EG1_Po_req_qty = 8
		Const M801_EG1_po_unit = 9
		Const M801_EG1_po_qty = 10
		Const M801_EG1_tracking_no = 11
		Const M801_EG1_wc_cd = 12
		Const M801_EG1_plan_start_dt = 13
		Const M801_EG1_plan_end_dt = 14
		Const M801_EG1_opr_no = 15
		Const M801_EG1_seq = 16
		Const M801_EG1_sub_seq = 17
		Const M801_EG1_release_dt = 18
    
	lgStrPrevKey = Request("lgStrPrevKey")
    If lgStrPrevKey <> "" then	
        arrValue = Split(lgStrPrevKey, gColSep)		
		I2_m_Dlvy_ord_next(M801_I2_opr_no) = arrValue(0)
		I2_m_Dlvy_ord_next(M801_I2_seq) = arrValue(1)
		I2_m_Dlvy_ord_next(M801_I2_sub_seq) = arrValue(2)
	Else			
		I2_m_Dlvy_ord_next(M801_I2_opr_no) = ""
		I2_m_Dlvy_ord_next(M801_I2_seq) = ""
		I2_m_Dlvy_ord_next(M801_I2_sub_seq) = ""
	End If	
	
    If Request("txtFromReqDt") <> "" Then
    	I1_m_Dlvy_ord(M801_I1_start_req_dt)	= UniConvDate(Request("txtFromReqDt"))
    End If 
    
    If Request("txtToReqDt") <> "" Then
    	I1_m_Dlvy_ord(M801_I1_end_req_dt)	= UniConvDate(Request("txtToReqDt"))
    End If
		
    I1_m_Dlvy_ord(M801_I1_plant_cd)			= Trim(Request("txtPlantCd"))
    I1_m_Dlvy_ord(M801_I1_start_prodt_no)	= Trim(Request("txtProdOrderNo1"))'
    I1_m_Dlvy_ord(M801_I1_end_prodt_no)		= Trim(Request("txtProdOrderNo2"))
    I1_m_Dlvy_ord(M801_I1_item_cd)			= Trim(Request("txtItemCd"))
    
    Set PMCG100 = Server.CreateObject("PMCG100.cMListDlvyOrder")    
	
    If CheckSYSTEMError(Err,True) = true Then 		
		Set PMCG100 = Nothing												'☜: ComProxy Unload
		Exit Sub														'☜: 비지니스 로직 처리를 종료함 
	End if
	
    EG1_m_Dlvy_ord = PMCG100.M_LIST_DLVY_ORDER(gStrGlobalCollection, _
												C_SHEETMAXROWS_D, _
												I1_m_Dlvy_ord, _
												I2_m_Dlvy_ord_next)
		
	If CheckSYSTEMError(Err,True) = true Then 		
		Set PMCG100 = Nothing												'☜: ComProxy Unload
		Response.Write "<Script Language=vbscript>"	& vbCr
		Response.Write "parent.frm1.txtPlantCd.focus"	& vbCr
		Response.Write "</Script>" & vbCr
		Exit Sub														'☜: 비지니스 로직 처리를 종료함 
	End if

	iLngMaxRow = Request("txtMaxRows")											'Save previous Maxrow                                                
    GroupCount = UBound(EG1_m_Dlvy_ord,1) 
    
	ReDim PvArr(GroupCount)
	
	For iLngRow = 0 To UBound(EG1_m_Dlvy_ord,1)
		If  iLngRow < C_SHEETMAXROWS_D  Then
		Else
		   StrNextKey = ConvSPChars(EG1_m_Dlvy_ord(iLngRow,M801_EG1_opr_no)) & gColSep & ConvSPChars(EG1_m_Dlvy_ord(iLngRow,M801_EG1_seq)) & gColSep & ConvSPChars(EG1_m_Dlvy_ord(iLngRow,M801_EG1_sub_seq))
           Exit For
        End If  
		
		istrData = Chr(11) & "0"
		istrData = istrData & Chr(11) & ConvSPChars(EG1_m_Dlvy_ord(iLngRow,M801_EG1_prodt_order_no))
        istrData = istrData & Chr(11) & ConvSPChars(EG1_m_Dlvy_ord(iLngRow,M801_EG1_item_cd))		
        istrData = istrData & Chr(11) & ConvSPChars(EG1_m_Dlvy_ord(iLngRow,M801_EG1_item_nm))		
        istrData = istrData & Chr(11) & ConvSPChars(EG1_m_Dlvy_ord(iLngRow,M801_EG1_spec))			
        istrData = istrData & Chr(11) & UNIDateClientFormat(EG1_m_Dlvy_ord(iLngRow,M801_EG1_req_dt))
        istrData = istrData & Chr(11) & ConvSPChars(EG1_m_Dlvy_ord(iLngRow,M801_EG1_basic_unit))		
        istrData = istrData & Chr(11) & UNINumClientFormat(EG1_m_Dlvy_ord(iLngRow,M801_EG1_req_qty),ggQty.DecPoint,0)
        istrData = istrData & Chr(11) & UNINumClientFormat(EG1_m_Dlvy_ord(iLngRow,M801_EG1_base_qty),ggQty.DecPoint,0)
        istrData = istrData & Chr(11) & UNINumClientFormat( CDbl(EG1_m_Dlvy_ord(iLngRow,M801_EG1_req_qty)) - CDbl(EG1_m_Dlvy_ord(iLngRow,M801_EG1_base_qty)),ggQty.DecPoint,0)
        'istrData = istrData & Chr(11) & UNINumClientFormat(EG1_m_Dlvy_ord(iLngRow,M801_EG1_Po_req_qty),ggQty.DecPoint,0)
        istrData = istrData & Chr(11) & ConvSPChars(EG1_m_Dlvy_ord(iLngRow,M801_EG1_po_unit))
        istrData = istrData & Chr(11) & UNINumClientFormat(EG1_m_Dlvy_ord(iLngRow,M801_EG1_po_qty),ggQty.DecPoint,0)
        istrData = istrData & Chr(11) & ConvSPChars(EG1_m_Dlvy_ord(iLngRow,M801_EG1_tracking_no))			
        istrData = istrData & Chr(11) & ConvSPChars(EG1_m_Dlvy_ord(iLngRow,M801_EG1_wc_cd))	
        istrData = istrData & Chr(11) & UNIDateClientFormat(EG1_m_Dlvy_ord(iLngRow,M801_EG1_plan_start_dt))
        istrData = istrData & Chr(11) & UNIDateClientFormat(EG1_m_Dlvy_ord(iLngRow,M801_EG1_plan_end_dt))
        istrData = istrData & Chr(11) & ConvSPChars(EG1_m_Dlvy_ord(iLngRow,M801_EG1_opr_no))
        istrData = istrData & Chr(11) & ConvSPChars(EG1_m_Dlvy_ord(iLngRow,M801_EG1_seq))
        istrData = istrData & Chr(11) & ConvSPChars(EG1_m_Dlvy_ord(iLngRow,M801_EG1_sub_seq))
        istrData = istrData & Chr(11) & UNIDateClientFormat(EG1_m_Dlvy_ord(iLngRow,M801_EG1_release_dt))
        istrData = istrData & Chr(11) & iLngMaxRow + iLngRow                             
        istrData = istrData & Chr(11) & Chr(12)       
        
        PvArr(iLngRow) = istrData
        
    Next  
    
    istrData = Join(PvArr, "")
    
    Response.Write "<Script Language=vbscript>" & vbCr
	Response.Write "With parent" & vbCr												'☜: 화면 처리 ASP 를 지칭함 
    Response.Write "	.ggoSpread.Source          =  .frm1.vspdData "         & vbCr
    Response.Write "	.ggoSpread.SSShowData        """ & istrData	    & """" & vbCr	
    Response.Write "	.lgStrPrevKey              = """ & ConvSPChars(StrNextKey)   & """" & vbCr 
     
    Response.Write " .frm1.hFromReqDt.value     = """ & ConvSPChars(Request("txtFromReqDt"))     & """" & vbCr
	Response.Write " .frm1.hToReqDt.value     = """ & ConvSPChars(Request("txtToReqDt"))     & """" & vbCr
	Response.Write " .frm1.hItemCd.value = """ & ConvSPChars(Request("txtItemCd")) & """" & vbCr
	Response.Write " .frm1.hProdOrderNo1.value    = """ & ConvSPChars(Request("txtProdOrderNo1"))    & """" & vbCr
	Response.Write " .frm1.hProdOrderNo2.value     = """ & ConvSPChars(Request("txtProdOrderNo2"))     & """" & vbCr
	
    Response.Write " .DbQueryOk "		    	  & vbCr 
    Response.Write " .frm1.vspdData.focus "		  & vbCr 
    Response.Write "End With" & vbCr
    Response.Write "</Script>" & vbCr


    Set PMCG100 = Nothing
		
End Sub    
'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data 
'============================================================================================================
Sub SubBizSaveMulti()
	On Error Resume Next
    Err.Clear																		'☜: Protect system from crashing

	Dim LngMaxRow
	Dim LngRow
	Dim PP4C210																	'☆ : 입력/수정용 ComProxy Dll 사용 변수 
    Dim iErrorPosition, iActRow
    Dim arrVal, arrTemp																'☜: Spread Sheet 의 값을 받을 Array 변수 
    
    Dim I1_module_flag
    Dim I2_plant_cd
    Dim I3_reservation
        
    Const P4PA_IG1_prodt_order_no = 0
    Const P4PA_IG1_opr_no = 1
    Const P4PA_IG1_seq = 2
    Const P4PA_IG1_sub_seq = 3
    Const P4PA_IG1_child_item_cd = 4
    Const P4PA_IG1_tracking_no = 5
    Const P4PA_IG1_req_dt = 6
    Const P4PA_IG1_req_qty = 7
    Const P4PA_IG1_base_unit = 8
    Const P4PA_IG1_wc_cd = 9
    Const P4PA_IG1_plan_start_dt = 10
    Const P4PA_IG1_plan_compt_dt = 11
    Const P4PA_IG1_release_dt = 12
    Const P4PA_IG1_child_item_req_round_flg = 13
    Const P4PA_IG1_error_position = 14

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

	LngMaxRow = CInt(Request("txtMaxRows"))											'☜: 최대 업데이트된 갯수 

	arrTemp = Split(itxtSpread, gRowSep)									'☆: Spread Sheet 내용을 담고 있는 Element명 
	
	Redim I3_reservation(LngMaxRow - 1, 14)

    '1건씩 처리한다 
    For LngRow = 1 To LngMaxRow 
		 
		arrVal = Split(arrTemp(LngRow-1), gColSep)
					
		I3_reservation(LngRow-1,P4PA_IG1_prodt_order_no) = Trim(arrVal(1))
		I3_reservation(LngRow-1,P4PA_IG1_opr_no) = Trim(arrVal(2))
		I3_reservation(LngRow-1,P4PA_IG1_seq) = Trim(arrVal(3))
		I3_reservation(LngRow-1,P4PA_IG1_sub_seq) = Trim(arrVal(4))
		I3_reservation(LngRow-1,P4PA_IG1_child_item_cd) = Trim(arrVal(5))
		I3_reservation(LngRow-1,P4PA_IG1_tracking_no) = Trim(arrVal(6))
		I3_reservation(LngRow-1,P4PA_IG1_req_dt) = Trim(arrVal(7))
		I3_reservation(LngRow-1,P4PA_IG1_req_qty) = Trim(arrVal(8))
		I3_reservation(LngRow-1,P4PA_IG1_base_unit) = Trim(arrVal(9))
		I3_reservation(LngRow-1,P4PA_IG1_wc_cd) = Trim(arrVal(10))
		I3_reservation(LngRow-1,P4PA_IG1_plan_start_dt) = Trim(arrVal(11))
		I3_reservation(LngRow-1,P4PA_IG1_plan_compt_dt) = Trim(arrVal(12))
		I3_reservation(LngRow-1,P4PA_IG1_release_dt) = Trim(arrVal(13))
		I3_reservation(LngRow-1,P4PA_IG1_child_item_req_round_flg) = ""
		I3_reservation(LngRow-1,P4PA_IG1_error_position) = Trim(arrVal(14))
		
	Next
	
	I1_module_flag = "PURCHASE"
	I2_plant_cd = Trim(Request("txtPlantCd"))
	
	Set PP4C210 = Server.CreateObject("PP4C210.cPMngDlvyOrd")    

	If CheckSYSTEMError(Err,True) = true Then 		
		Set PP4C210 = Nothing												'☜: ComProxy Unload
		Exit Sub														'☜: 비지니스 로직 처리를 종료함 
	
	End if
	'납입지시대상 선정된 행수를 메시지 처리함.(2003.07.25)	
	Call PP4C210.P_MANAGE_DLVY_ORD(gStrGlobalCollection, I1_module_flag, I2_plant_cd, _
		I3_reservation,,iErrorPosition, iActRow) 
                   
	If CheckSYSTEMError2(Err,True,iErrorPosition & "행:","","","","") = True Then
	   Set PP4C210 = Nothing
	   Call SheetFocus(iErrorPosition, 1, I_MKSCRIPT)
	   Exit Sub
	End If
		
	Set PP4C210 = Nothing                                                   '☜: Unload Comproxy
	
	Call DisplayMsgBox("17C026", vbInformation, iActRow, "", I_MKSCRIPT)
   
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
