<%@ LANGUAGE="VBSCRIPT" %>
<%Option Explicit    %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Prucurement
'*  2. Function Name        : 
'*  3. Program ID           : mc200mb1
'*  4. Program Name         : 납입지시조정 
'*  5. Program Desc         : 납입지시조정 
'*  6. Component List       : PMCG200.cMListDlvyOrdChg / PMCG250.cMManageDlvyOrdChgSvr
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
    	
	Dim PMCG200																	'☆ : 조회용 ComProxy Dll 사용 변수 
	
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
		Const M802_I1_plant_cd = 0
		Const M802_I1_start_req_dt = 1
		Const M802_I1_end_req_dt = 2
		Const M802_I1_start_prodt_no = 3
		Const M802_I1_end_prodt_no = 4
		Const M802_I1_bp_cd = 5
		Const M802_I1_item_cd = 6
    ReDim I1_m_Dlvy_ord(6)
	
    Dim I2_m_Dlvy_ord_next
		Const M802_I2_opr_no = 0
		Const M802_I2_seq = 1
		Const M802_I2_sub_seq = 2
    ReDim I2_m_Dlvy_ord_next(2)
       
    Dim EG1_m_Dlvy_ord
		Const M802_EG1_prodt_order_no = 0
		Const M802_EG1_item_cd = 1
		Const M802_EG1_item_nm = 2
		Const M802_EG1_spec = 3
		Const M802_EG1_req_dt = 4
		Const M802_EG1_basic_unit = 5
		Const M802_EG1_req_qty = 6
		Const M802_EG1_bp_cd = 7
		Const M802_EG1_bp_nm = 8
		Const M802_EG1_do_qty = 9
		Const M802_EG1_tracking_no = 10
		Const M802_EG1_wc_cd = 11
		Const M802_EG1_plan_start_dt = 12
		Const M802_EG1_plan_end_dt = 13
		Const M802_EG1_opr_no = 14
		Const M802_EG1_po_no = 15
		Const M802_EG1_po_seq_no = 16
		Const M802_EG1_seq = 17
		Const M802_EG1_sub_seq = 18
            
	
    '-----------------------
    'Data manipulate  area(import view match)
    '-----------------------
	lgStrPrevKey = Request("lgStrPrevKey")
    If lgStrPrevKey <> "" then	
        arrValue = Split(lgStrPrevKey, gColSep)		
		I2_m_Dlvy_ord_next(M802_I2_opr_no) = arrValue(0)
		I2_m_Dlvy_ord_next(M802_I2_seq) = arrValue(1)
		I2_m_Dlvy_ord_next(M802_I2_sub_seq) = arrValue(2)
	Else			
		I2_m_Dlvy_ord_next(M802_I2_opr_no) = ""
		I2_m_Dlvy_ord_next(M802_I2_seq) = ""
		I2_m_Dlvy_ord_next(M802_I2_sub_seq) = ""
	End If	

    If Request("txtFromReqDt") <> "" Then
    	I1_m_Dlvy_ord(M802_I1_start_req_dt)			= UniConvDate(Request("txtFromReqDt"))
    End If 
    
    If Request("txtToReqDt") <> "" Then
    	I1_m_Dlvy_ord(M802_I1_end_req_dt)			= UniConvDate(Request("txtToReqDt"))
    End If
		
    I1_m_Dlvy_ord(M802_I1_plant_cd)			= Trim(Request("txtPlantCd"))
    I1_m_Dlvy_ord(M802_I1_start_prodt_no)	= Trim(Request("txtProdOrderNo1"))'
    I1_m_Dlvy_ord(M802_I1_end_prodt_no)		= Trim(Request("txtProdOrderNo2"))
    I1_m_Dlvy_ord(M802_I1_bp_cd)			= Trim(Request("txtSupplier"))	
    I1_m_Dlvy_ord(M802_I1_item_cd)			= Trim(Request("txtItemCd"))


    Set PMCG200 = Server.CreateObject("PMCG200.cMListDlvyOrdChg")    
	
    '-----------------------
    'Com action result check area(OS,internal)
    '-----------------------
    If CheckSYSTEMError(Err,True) = true Then 		
		Set PMCG200 = Nothing												'☜: ComProxy Unload
		Exit Sub														'☜: 비지니스 로직 처리를 종료함 
	End if

    EG1_m_Dlvy_ord = PMCG200.M_LIST_DLVY_ORD_CHG(gStrGlobalCollection, _
												 C_SHEETMAXROWS_D, _
												 I1_m_Dlvy_ord, _
												 I2_m_Dlvy_ord_next)
						

	'-----------------------
	'Com action result check area(DB,internal)
	'-----------------------
	If CheckSYSTEMError(Err,True) = true Then 		
		Set PMCG200 = Nothing												'☜: ComProxy Unload
		Response.Write "<Script Language=vbscript>"	& vbCr
		Response.Write "parent.frm1.txtPlantCd.focus"	& vbCr
		Response.Write "</Script>" & vbCr
		Exit Sub														'☜: 비지니스 로직 처리를 종료함 
	End if
	

	iLngMaxRow = Request("txtMaxRows")											'Save previous Maxrow                                                
    GroupCount = UBound(EG1_m_Dlvy_ord,1) 
       
	'-----------------------
	'Result data display area
	'----------------------- 
	ReDim PvArr(GroupCount)
	
	For iLngRow = 0 To UBound(EG1_m_Dlvy_ord,1)
		If  iLngRow < C_SHEETMAXROWS_D  Then
		Else
		   StrNextKey = ConvSPChars(EG1_m_Dlvy_ord(iLngRow,M802_EG1_opr_no)) & gColSep & ConvSPChars(EG1_m_Dlvy_ord(iLngRow,M802_EG1_seq)) & gColSep & ConvSPChars(EG1_m_Dlvy_ord(iLngRow,M802_EG1_sub_seq))
           Exit For
        End If  
   		   		
		istrData = Chr(11) & ConvSPChars(EG1_m_Dlvy_ord(iLngRow,M802_EG1_prodt_order_no))
        istrData = istrData & Chr(11) & ConvSPChars(EG1_m_Dlvy_ord(iLngRow,M802_EG1_item_cd))		
        istrData = istrData & Chr(11) & ConvSPChars(EG1_m_Dlvy_ord(iLngRow,M802_EG1_item_nm))		
        istrData = istrData & Chr(11) & ConvSPChars(EG1_m_Dlvy_ord(iLngRow,M802_EG1_spec))			
        istrData = istrData & Chr(11) & UNIDateClientFormat(EG1_m_Dlvy_ord(iLngRow,M802_EG1_req_dt))
        istrData = istrData & Chr(11) & ConvSPChars(EG1_m_Dlvy_ord(iLngRow,M802_EG1_basic_unit))		
        istrData = istrData & Chr(11) & UNINumClientFormat(EG1_m_Dlvy_ord(iLngRow,M802_EG1_req_qty),ggQty.DecPoint,0)
        istrData = istrData & Chr(11) & ConvSPChars(EG1_m_Dlvy_ord(iLngRow,M802_EG1_bp_cd))
        istrData = istrData & Chr(11) & ""	
        istrData = istrData & Chr(11) & ConvSPChars(EG1_m_Dlvy_ord(iLngRow,M802_EG1_bp_nm))
        istrData = istrData & Chr(11) & UNINumClientFormat(EG1_m_Dlvy_ord(iLngRow,M802_EG1_do_qty),ggQty.DecPoint,0)
        istrData = istrData & Chr(11) & ConvSPChars(EG1_m_Dlvy_ord(iLngRow,M802_EG1_tracking_no))			
        istrData = istrData & Chr(11) & ConvSPChars(EG1_m_Dlvy_ord(iLngRow,M802_EG1_wc_cd))	
        istrData = istrData & Chr(11) & UNIDateClientFormat(EG1_m_Dlvy_ord(iLngRow,M802_EG1_plan_start_dt))
        istrData = istrData & Chr(11) & UNIDateClientFormat(EG1_m_Dlvy_ord(iLngRow,M802_EG1_plan_end_dt))
        istrData = istrData & Chr(11) & ConvSPChars(EG1_m_Dlvy_ord(iLngRow,M802_EG1_opr_no))
        istrData = istrData & Chr(11) & ConvSPChars(EG1_m_Dlvy_ord(iLngRow,M802_EG1_po_no))
        istrData = istrData & Chr(11) & ConvSPChars(EG1_m_Dlvy_ord(iLngRow,M802_EG1_po_seq_no))
        istrData = istrData & Chr(11) & ConvSPChars(EG1_m_Dlvy_ord(iLngRow,M802_EG1_seq))
        istrData = istrData & Chr(11) & ConvSPChars(EG1_m_Dlvy_ord(iLngRow,M802_EG1_sub_seq))
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
	Response.Write " .frm1.hSupplier.value = """ & ConvSPChars(Request("txtSupplierCd")) & """" & vbCr
	Response.Write " .frm1.hProdOrderNo1.value    = """ & ConvSPChars(Request("txtProdOrderNo1"))    & """" & vbCr
	Response.Write " .frm1.hProdOrderNo2.value     = """ & ConvSPChars(Request("txtProdOrderNo2"))     & """" & vbCr
	
    Response.Write " .DbQueryOk "		    	  & vbCr 
    Response.Write " .frm1.vspdData.focus "		  & vbCr 
    Response.Write "End With" & vbCr
    Response.Write "</Script>" & vbCr


    Set PMCG200 = Nothing
		
End Sub    
'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data 
'============================================================================================================
Sub SubBizSaveMulti()
	On Error Resume Next
    Err.Clear																		'☜: Protect system from crashing
	
	Dim PMCG250																	'☆ : 입력/수정용 ComProxy Dll 사용 변수 
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
    
    Set PMCG250 = Server.CreateObject("PMCG250.cMManageDlvyOrdChgSvr")    

    '-----------------------
    'Com action result check area(OS,internal)
    '-----------------------
    If CheckSYSTEMError(Err,True) = true Then 		
		Set PMCG250 = Nothing												'☜: ComProxy Unload
		Exit Sub														'☜: 비지니스 로직 처리를 종료함 
	End if
	
	Call PMCG250.M_MANAGE_DLVY_ORD_CHG_SVR(gStrGlobalCollection, itxtSpread, iErrorPosition) 
                   
    If CheckSYSTEMError2(Err,True,iErrorPosition & "행:","","","","") = True Then
       Set PMCG250 = Nothing
	   Call SheetFocus(iErrorPosition, 1, I_MKSCRIPT)
       Exit Sub
	End If

    Set PMCG250 = Nothing                                                   '☜: Unload Comproxy
    
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
