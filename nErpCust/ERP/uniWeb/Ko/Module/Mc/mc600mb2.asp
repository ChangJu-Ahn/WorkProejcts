<%@ LANGUAGE="VBSCRIPT" %>
<%Option Explicit    %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Prucurement
'*  2. Function Name        : 
'*  3. Program ID           : mc600mb2
'*  4. Program Name         : 납입지시입고등록 
'*  5. Program Desc         : 납입지시입고등록 
'*  6. Component List       : PMCG650.cMMangeDlvyOrdRcpt
'*  7. Modified date(First) : 2003-02-25
'*  8. Modified date(Last)  : 2003/05/23
'*  9. Modifier (First)     : Ryu Sung Won
'* 10. Modifier (Last)      : Kang Su Hwan
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change" 
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************
-->
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->

<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<%
	On Error Resume Next
	
	Call LoadBasisGlobalInf()
	Call LoadInfTB19029B("*", "M","NOCOOKIE", "MB")
	Call LoadBNumericFormatB("*", "M","NOCOOKIE", "MB")
	Call HideStatusWnd

    Call SubBizSaveMulti()

'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data into Db
'============================================================================================================
Sub subBizSaveMulti()															'☜: 저장 요청을 받음 
     On Error Resume Next 		
    Err.Clear																		'☜: Protect system from crashing

    Dim iPMCG650
    Dim iErrorPosition
    ReDim iErrorPosition(1)
    Dim dErrorPosition
	
	Dim I1_b_biz_partner_bp_cd
	Dim	I2_m_mvmt_type_io_type_cd
	Dim	I3_b_pur_grp_pur_grp
	Dim	I4_m_pur_goods_mvmt
		Const M804_I4_document_no = 0
		Const M804_I4_document_dt = 1
    Redim I4_m_pur_goods_mvmt(M804_I4_document_dt)

    Dim E1_m_pur_goods_mvmt_rcpt_no
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
	
	lgIntFlgMode = CInt(Request("txtFlgMode"))										'☜: 저장시 Create/Update 판별 
		
    Set iPMCG650 = Server.CreateObject("PMCG650.cMMangeDlvyOrdRcpt")    

    If CheckSYSTEMError(Err,True) = True Then Exit Sub
	
    '-----------------------
    'Data manipulate area
    '-----------------------
	If lgIntFlgMode = OPMD_CMODE then
		I1_b_biz_partner_bp_cd                  = UCase(Trim(Request("txtSupplierCd")))
		I2_m_mvmt_type_io_type_cd               = Trim(Request("cboMvmtType"))
		I3_b_pur_grp_pur_grp                    = UCase(Trim(Request("txtGroupCd")))
		I4_m_pur_goods_mvmt(M804_I4_document_no)		= Request("txtMvmtNo1")
		I4_m_pur_goods_mvmt(M804_I4_document_dt)		= UNIConvDate(Request("txtGmDt"))
        'Call ServerMesgBox("Request(txtGmDt) : " & Request("txtGmDt") , vbInformation, I_MKSCRIPT)

		Call iPMCG650.M_CREATE_DLVY_ORDER_RCPT(gStrGlobalCollection, _
											I1_b_biz_partner_bp_cd, _
											I2_m_mvmt_type_io_type_cd, _
											I3_b_pur_grp_pur_grp, _
											I4_m_pur_goods_mvmt, _
											itxtSpread, _
											E1_m_pur_goods_mvmt_rcpt_no, _
											iErrorPosition)
		If CheckSYSTEMError2(Err, True, iErrorPosition(0) & "행","","","","") = True Then
			Call SheetFocus(iErrorPosition, 1, I_MKSCRIPT)
			Set iPM7G421 = Nothing
			Exit Sub
		End If		
	
	Else 
		
		Call iPMCG650.M_DELETE_DLVY_ORDER_RCPT(gStrGlobalCollection, _
											itxtSpread, _
											dErrorPosition)
		If CheckSYSTEMError2(Err, True, dErrorPosition & "행","","","","") = True Then
			Call SheetFocus(iErrorPosition, 1, I_MKSCRIPT)
			Set iPM7G421 = Nothing
			Exit Sub
		End If		

	End if
 	
	If iErrorPosition(1) <> "" Then
		Call DisplayMsgBox("17C024", vbInformation, iErrorPosition(1), "", I_MKSCRIPT)
	End If
   
    Set iPM7G421 = Nothing                                                   '☜: Unload Comproxy  

   	Response.Write "<Script language=vbs> " & vbCr 
	Response.Write "With parent " & vbCr
	
	Response.Write "	If """ & lgIntFlgMode & """ = """ & OPMD_CMODE & """ Then " & vbCr
	Response.Write "		.frm1.txtMvmtNo.Value = """ & UCase(ConvSPChars(E1_m_pur_goods_mvmt_rcpt_no)) & """ " & vbCr
	Response.Write "	End If"				& vbCr	
    
    Response.Write "	.DbSaveOk "      & vbCr						
    
    Response.Write "End With " & vbCr
    Response.Write "</Script> "    
    
End Sub	

%>
<Script Language=vbscript RUNAT=server>
'==============================================================================
' Function : SheetFocus
' Description : 에러발생시 Spread Sheet에 포커스줌 
'==============================================================================
Function SheetFocus(Byval lRow, Byval lCol, Byval iLoc)
	Dim strHTML
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
</Script>
