<%@ LANGUAGE=VBSCript%>
<%Option Explicit    %>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->

<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<%
'**********************************************************************************************
'*  1. Module Name          : Procurement
'*  2. Function Name        : 
'*  3. Program ID           : m4132mb1
'*  4. Program Name         : 예외입고/반품등록 
'*  5. Program Desc         :
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2005/08/22
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Kim Duk Hyun
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'* 14. Business Logic of m4132ma1
'**********************************************************************************************
	
On Error Resume Next
Err.Clear
	
Call LoadBasisGlobalInf()
Call LoadInfTB19029B("I", "*","NOCOOKIE", "MB")
Call LoadBNumericFormatB("I", "*","NOCOOKIE", "MB")
Call HideStatusWnd

lgOpModeCRUD	=	Request("txtMode")	'☜: Read Operation Mode (CRUD)
			
Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '☜: Query
             Call SubBizQueryMulti()
        Case CStr(UID_M0002)
             Call SubBizSaveMulti()
        Case CStr ("changeMvmtType")
			 Call SubchangeMvmtType()
		Case "changeSpplCd" 
			 Call DisplaySupplierNm(Request("txtSupplierCd"))	 
		Case "changeGroupCd" 
			 Call DisplayGroupNm(Request("txtGroupCd"))	 				 
		Case "changeIvTypeCd" 
			 Call DisplayIvTypeNm(Request("txtIvTypeCd"))	 				 
End Select

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()														'☜: 현재 조회/Prev/Next 요청을 받음 
 
	Dim iPM7G438
	
	Dim istrData
	Dim StrNextKey		' 다음 값 
	Dim lgStrPrevKey	' 이전 값 
	Dim iLngMaxRow		' 현재 그리드의 최대Row
	Dim iLngRow
	Dim GroupCount          
	Dim index,Count     ' 저장 후 Return 해줄 값을 넣을때 쓴는 변수 
	 
	Dim TmpBuffer
	Dim iMax
	Dim iIntLoopCount
	Dim iTotalStr
	    
	Dim I1_m_pur_goods_mvmt_no
	Dim I2_m_pur_goods_mvmt_rcpt_no
	 
	Dim EG1_export_group
	Const M_mvmt_rcpt_no	= 0
	Const M_io_type_cd		= 1
	Const M_io_type_nm		= 2
	Const M_mvmt_rcpt_dt	= 3
	Const M_bp_cd			= 4
	Const M_bp_nm			= 5
	Const M_pur_grp			= 6
	Const M_pur_grp_nm		= 7
	
	Const M_plant_cd		= 8
	Const M_plant_nm		= 9
	Const M_item_cd			= 10
	Const M_item_nm			= 11
	Const M_spec			= 12
	Const M_mvmt_rcpt_unit	= 13
	Const M_mvmt_rcpt_qyt	= 14
	Const M_mvmt_qty		= 15
	Const M_mvmt_cur		= 16
	Const M_item_prc		= 17
	Const M_item_doc_amt	= 18
	Const M_mvmt_prc		= 19
	Const M_mvmt_doc_amt	= 20
	Const M_sl_cd			= 21
	Const M_sl_nm			= 22
	Const M_lot_no			= 23
	Const M_lot_sub_no		= 24
	Const M_maker_lot_no	= 25
	Const M_maker_lot_sub_no= 26
	Const M_ret_type		= 27
	Const M_ret_type_nm		= 28
	Const M_tracking_no		= 29
	Const M_gm_no			= 30
	Const M_gm_seq_no		= 31
	Const M_inspect_flg		= 32
	Const M_inspect_sts		= 33
	Const M_mvmt_method		= 34
	Const M_inspect_req_no	= 35
	Const M_inspect_result_no = 36
	Const M_mvmt_no			= 37
	Const M_ref_mvmt_no		= 38
	Const M_procure_type	= 39
	Const M_remark_hdr		= 40
	Const M_remark_dtl		= 41
	
	Dim strGmNo, strTempGlNo, strGlNo 
    
    Const C_SHEETMAXROWS_D  = 100
    
    On Error Resume Next 
	Err.Clear                                                               '☜: Protect system from crashing
    
	lgStrPrevKey = Request("lgStrPrevKey")
  
    Set iPM7G438 = Server.CreateObject("PM7G438.cMListExceptPurRcptS")    

    If CheckSYSTEMError(Err,True) = True Then
		Exit Sub
	End If
	
    I2_m_pur_goods_mvmt_rcpt_no  		= Request("txtGrNo")
    
    if Trim(lgStrPrevKey) <> "" then
		I1_m_pur_goods_mvmt_no  	= lgStrPrevKey
	End if
    
    Call iPM7G438.M_LIST_EXCEPT_PUR_RCPT_SVR(gStrGlobalCollection, C_SHEETMAXROWS_D, I1_m_pur_goods_mvmt_no, I2_m_pur_goods_mvmt_rcpt_no, EG1_export_group)
	
	If CheckSYSTEMError2(Err,True,"","","","","") = true then 
		Set iPM7G438 = Nothing
		Response.Write "<Script Language=VBScript>" & vbCr
		Response.Write "	Parent.FncNew() " & vbCr
		Response.Write "	parent.frm1.txtGrNo.Value 	= """ & ConvSPChars(Request("txtGrNo")) & """" & vbCr
		Response.Write "</Script>" & vbCr
		Exit Sub
	End If

	iLngMaxRow = Request("txtMaxRows")											'Save previous Maxrow                                                
    GroupCount = UBound(EG1_export_group,1)
    
	
	Response.Write "<Script Language=VBScript>" & vbCr
	Response.Write "With parent" & vbCr
    Response.Write ".frm1.txtMvmtType.Value      = """ & ConvSPChars(EG1_export_group(GroupCount,M_io_type_cd)) & """" 				& vbCr
    Response.Write ".frm1.txtMvmtTypeNm.Value    = """ & ConvSPChars(EG1_export_group(GroupCount,M_io_type_nm)) & """" 				& vbCr       	   		
    Response.Write ".frm1.txtGmDt.text           = """ & UNIDateClientFormat(EG1_export_group(GroupCount,M_mvmt_rcpt_dt)) & """" 	& vbCr
    Response.Write ".frm1.txtGroupCd.Value       = """ & ConvSPChars(EG1_export_group(GroupCount,M_pur_grp)) & """"       			& vbCr
    Response.Write ".frm1.txtGroupNm.Value       = """ & ConvSPChars(EG1_export_group(GroupCount,M_pur_grp_nm)) & """"    			& vbCr
    Response.Write ".frm1.txtSupplierCd.Value    = """ & ConvSPChars(EG1_export_group(GroupCount,M_bp_cd)) & """"     				& vbCr
    Response.Write ".frm1.txtSupplierNm.Value    = """ & ConvSPChars(EG1_export_group(GroupCount,M_bp_nm)) & """"     				& vbCr
    Response.Write ".frm1.txtGrNo.Value          = """ & ConvSPChars(I2_m_pur_goods_mvmt_rcpt_no) & """"         					& vbCr
    Response.Write ".frm1.txtRemark.Value        = """ & ConvSPChars(EG1_export_group(GroupCount,M_remark_hdr)) & """"         		& vbCr
	Response.Write "End With" & vbCr
	Response.Write "</Script>"	   & vbCr

     
    iIntLoopCount = 0
	iMax = UBound(EG1_export_group,1)
	ReDim TmpBuffer(iMax)
		
	For iLngRow = 0 To iMax
		If  iLngRow < C_SHEETMAXROWS_D  Then
		Else
		   StrNextKey = ConvSPChars(EG1_export_group(iLngRow, M744_EG1_E11_mvmt_no)) 
           Exit For
        End If  
		
		istrData = ""
        istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow,M_plant_cd))				'입고Seq
        istrData = istrData & Chr(11) & "" 
        istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow,M_plant_nm))
        istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow,M_item_cd))
        istrData = istrData & Chr(11) & "" 
        istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow,M_item_nm))	
        istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow,M_spec))	
        istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow,M_mvmt_rcpt_unit))	
        istrData = istrData & Chr(11) & "" 
                	
        istrData = istrData & Chr(11) & UNINumClientFormat(EG1_export_group(iLngRow,M_mvmt_rcpt_qyt),ggQty.DecPoint,0)
        istrData = istrData & Chr(11) & "" 
        istrData = istrData & Chr(11) & UNINumClientFormat(EG1_export_group(iLngRow,M_mvmt_qty),ggQty.DecPoint,0)
        istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow,M_mvmt_cur))	
        istrData = istrData & Chr(11) & "" 

	    'istrData = istrData & Chr(11) & UNINumClientFormat(EG1_export_group(iLngRow,M_item_prc),ggQty.DecPoint,0)
	    'istrData = istrData & Chr(11) & UNINumClientFormat(EG1_export_group(iLngRow,M_item_doc_amt),ggQty.DecPoint,0)
	    istrData = istrData & Chr(11) & UniConvNumDBToCompanyWithOutChange(EG1_export_group(iLngRow,M_item_prc), 0)
	    istrData = istrData & Chr(11) & UniConvNumDBToCompanyWithOutChange(EG1_export_group(iLngRow,M_item_doc_amt), 0)
		If Trim(ConvSPChars(EG1_export_group(iLngRow,M_procure_type))) <> "P" Then
'	        istrData = istrData & Chr(11) & UNINumClientFormat(EG1_export_group(iLngRow,M_mvmt_prc),ggQty.DecPoint,0)
'	        istrData = istrData & Chr(11) & UNINumClientFormat(EG1_export_group(iLngRow,M_mvmt_doc_amt),ggQty.DecPoint,0)
	        istrData = istrData & Chr(11) & UniConvNumDBToCompanyWithOutChange(EG1_export_group(iLngRow,M_mvmt_prc), 0)
	        istrData = istrData & Chr(11) & UniConvNumDBToCompanyWithOutChange(EG1_export_group(iLngRow,M_mvmt_doc_amt), 0)
		Else
	    	istrData = istrData & Chr(11) & "0"
	    	istrData = istrData & Chr(11) & "0"
		End If

        istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow,M_sl_cd))  
        istrData = istrData & Chr(11) & "" 
        istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow,M_sl_nm))
        istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow,M_lot_no))
        istrData = istrData & Chr(11) & "" 
        istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow,M_lot_sub_no))
        istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow,M_maker_lot_no))
        istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow,M_maker_lot_sub_no))
        istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow,M_ret_type))
        istrData = istrData & Chr(11) & "" 
        istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow,M_ret_type_nm)) 
        istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow,M_remark_dtl)) 
        istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow,M_tracking_no))
        istrData = istrData & Chr(11) & "" 
        istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow,M_gm_no))
        istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow,M_gm_seq_no))
        If ConvSPChars(EG1_export_group(iLngRow,M_inspect_flg)) = "N" Then
        	istrData = istrData & Chr(11) & "0"
        Else
        	istrData = istrData & Chr(11) & "Y"
        End If
        istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow,M_inspect_sts))
        istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow,M_mvmt_method))
        istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow,M_inspect_req_no))
        istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow,M_inspect_result_no))
        istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow,M_mvmt_no))
        istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow,M_ref_mvmt_no))
        istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow,M_procure_type))

        istrData = istrData & Chr(11) & iLngMaxRow + iLngRow
        istrData = istrData & Chr(11) & Chr(12)

        If strGmNo = "" Then
            strGmNo = Trim(EG1_export_group(iLngRow,M_gm_no))
        End if
        
        TmpBuffer(iIntLoopCount) = istrData
        iIntLoopCount = iIntLoopCount + 1
    Next
	
	iTotalStr = Join(TmpBuffer, "")    

	Response.Write "<Script Language=VBScript>" & vbCr
	Response.Write "With parent" & vbCr
	Response.Write "	.ggoSpread.Source 		= .frm1.vspdData" & vbCr
    Response.Write "	.ggoSpread.SSShowData        """ & iTotalStr	& """" & vbCr	   
	Response.Write "	.lgStrPrevKey              = """ & StrNextKey   & """" & vbCr 
    Response.Write "	.frm1.hdnGrNo.value = """ & ConvSPChars(Request("txtGrNo")) & """" & vbCr
    Response.Write "	.frm1.txtGrNo1.value = """ & ConvSPChars(Request("txtGrNo")) & """" & vbCr
    Response.write "	.DbQueryOk " & vbCr
    Response.Write "End With" & vbCr
	Response.Write "</Script>"	   & vbCr
	
	set iPM7G438 = Nothing
	
	'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
	'******* 전표No. 만들기. *********
	'*************************************************************************************************************

    If  strGmNo <> "" Or strGmNo <> Null then  

		Call SubOpenDB(lgObjConn)                                                        '☜: Make a DB Connection
				
		lgStrSQL = "SELECT document_year FROM i_goods_movement_header " 
		lgStrSQL = lgStrSQL & " WHERE item_document_no =  " & FilterVar(strGmNo , "''", "S") & ""		
		If FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False then
			Response.Write "<Script Language=VBScript>" & vbCr
			Response.Write "parent.frm1.hdnGlType.value	=	""B"" " & vbCr
			Response.Write "</Script>"					& vbCr
			
			Call SubCloseRs(lgObjRs)  
			Call SubCloseDB(lgObjConn)                                                        '☜: Make a DB Connection
			Exit Sub
		End if
		'A_GL.Ref_no
		strGmNo	=	strGmNo & "-" & lgObjRs("document_year")
							
		'수정(전표조회 추가)
		Response.Write "<Script Language=VBScript>" & vbCr
		
		lgStrSQL = "select temp_gl_no,gl_no from ufn_a_GetGlNo(  " & FilterVar(strGmNo, "''", "S") & ") " 		
				
		IF FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False then
			Response.Write "parent.frm1.hdnGlType.value	=	""B""	  " & vbCr
			Response.Write "parent.frm1.hdnGlNo.value	=	""""      " & vbCr  
		Else
			strTempGlNo = lgObjRs("temp_gl_no")
			strGlNo = lgObjRs("gl_no")
			
			If ConvSPChars(Trim(strGlNo)) = "" And ConvSPChars(Trim(strTempGlNo)) <> "" Then
				Response.Write "parent.frm1.hdnGlType.value	=	""T""	  " & vbCr
				Response.Write "parent.frm1.hdnGlNo.value	=	""" & lgObjRs("temp_gl_no") & """" & vbCr  
			Else
				Response.Write "parent.frm1.hdnGlType.value	=	""A""	  " & vbCr
				Response.Write "parent.frm1.hdnGlNo.value	=	""" & lgObjRs("gl_no") & """" & vbCr  
			End If
		End If
		Response.Write "</Script>"					& vbCr
		Call SubCloseRs(lgObjRs)
		Call SubCloseDB(lgObjConn)
	Else
		Response.Write "<Script Language=VBScript>" & vbCr
		Response.Write "parent.frm1.hdnGlType.value	=	""B""	  " & vbCr
		Response.Write "</Script>"					& vbCr
	End If
	
End Sub

'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data into Db
'============================================================================================================
Sub SubBizSaveMulti()														'☜: 저장 요청을 받음 

	Dim iPM7G431
	Dim iCommandSent
	Dim iErrorPosition
	
	Dim itxtSpread
	Dim itxtSpreadArr
    Dim itxtSpreadArrCount
    Dim iCUCount
    Dim iDCount
    Dim ii
     
	Dim arrVal, arrTemp, strStatus		
	 
	Dim I1_ief_supplied_select_char	

	Dim I2_m_pur_goods_mvmt
	Const M686_I2_mvmt_rcpt_no = 0
	Const M686_I2_mvmt_rcpt_dt = 1
	Const M686_I2_gm_no = 2
	Const M686_I2_gm_seq_no = 3
	Redim  I2_m_pur_goods_mvmt(M686_I2_gm_seq_no)
	 
	Dim I3_m_mvmt_type_io_type_cd
	Dim I4_b_biz_partner_bp_cd
	Dim I5_b_pur_grp
	Dim I6_remark_hdr

	Dim E5_b_auto_numbering
	Dim iStrSpread
	Dim iLngMaxRow
  
    On Error Resume Next 
    Err.Clear																		'☜: Protect system from crashing

	Call RemovedivTextArea()
	
	If Len(Request("txtGmDt")) Then
		If UNIConvDate(Request("txtGmDt")) = "" Then
		    Call DisplayMsgBox("122116", vbInformation, "", "", I_MKSCRIPT)
		    Exit Sub	
		End If
	End If
	
	itxtSpread = ""
    
    iCUCount = Request.Form("txtCUSpread").Count
    iDCount  = Request.Form("txtDSpread").Count
    
    itxtSpreadArrCount = -1
    
    ReDim itxtSpreadArr(iCUCount + iDCount)
             
    For ii = 1 To iDCount
        itxtSpreadArrCount = itxtSpreadArrCount + 1
        itxtSpreadArr(itxtSpreadArrCount) = Request.Form("txtDSpread")(ii)
    Next
    
    For ii = 1 To iCUCount
        itxtSpreadArrCount = itxtSpreadArrCount + 1
        itxtSpreadArr(itxtSpreadArrCount) = Request.Form("txtCUSpread")(ii)
    Next
   
    itxtSpread = Join(itxtSpreadArr,"")
    
	lgIntFlgMode = CInt(Request("txtFlgMode"))	
	iLngMaxRow = CInt(Request("txtMaxRows"))										'☜: 최대 업데이트된 갯수 
	
	I2_m_pur_goods_mvmt(M686_I2_mvmt_rcpt_no) = Request("txtGrNo1")					' 입출고번호 
    I2_m_pur_goods_mvmt(M686_I2_mvmt_rcpt_dt) = UNIConvDate(Request("txtGmDt"))
    I2_m_pur_goods_mvmt(M686_I2_gm_no)        = ""
    I3_m_mvmt_type_io_type_cd                 = UCase(Request("txtMvmtType"))		' 입출고 유형 
    I4_b_biz_partner_bp_cd                    = Trim(Request("txtSupplierCd"))		' 공급처 
    I5_b_pur_grp                              = Trim(Request("txtGroupCd"))			' 구매그룹 
    I6_remark_hdr							  = Trim(Request("txtRemark"))			' 비고 
    
    arrTemp = Split(itxtSpread, gRowSep)	
    arrVal = Split(arrTemp(0), gColSep)
	    	
    strStatus = arrVal(0)
    
    If strStatus = "C" then
       iCommandSent			= "CREATE"
    Else 
       iCommandSent			= "DELETE"		
    End If

    Set iPM7G431 = Server.CreateObject("PM7G431.cMExceptPurRcpt")    

    If CheckSYSTEMError(Err,True) = True Then
		Exit Sub
	End If

    Call iPM7G431.M_EXCEPT_PUR_RCPT_SVR(gStrGlobalCollection, iCommandSent,I2_m_pur_goods_mvmt, _
                                UCase(I3_m_mvmt_type_io_type_cd), UCase(I4_b_biz_partner_bp_cd), UCase(I5_b_pur_grp), I6_remark_hdr, _
                                itxtSpread, E5_b_auto_numbering, iErrorPosition)
	
	If CheckSYSTEMError2(Err, True, iErrorPosition & "행","","","","") = True Then
	   Set iPM7G431 = Nothing
	   Exit Sub
	End If		

    Set iPM7G431 = Nothing                                                   '☜: Unload Comproxy  
    
   	Response.Write "<Script language=vbs> " & vbCr 
	Response.Write "With parent " & vbCr
	Response.Write "	If """ & lgIntFlgMode & """ = """ & OPMD_CMODE & """ Then " & vbCr
	Response.Write "		.frm1.txtGrNo.value    = """ & UCase(ConvSPChars(E5_b_auto_numbering))  & """" & vbCr'
	Response.Write "	End If"				& vbCr	
    Response.Write " .DbSaveOk "      & vbCr						
    Response.Write "End With " & vbCr
    Response.Write "</Script> "    
  
End Sub

'============================================================================================================
' Name : SubchangeMvmtType
' Desc : 
'============================================================================================================
Sub SubchangeMvmtType()	
  
  Dim iPM1G439
   
  Dim I1_m_mvmt_type_io_type_cd
  Dim I2_ief_supplied_SELECT_CHAR
  
  Dim E1_m_mvmt_type
  Const EA_m_mvmt_type_io_type_cd1 = 0
  Const EA_m_mvmt_type_io_type_nm1 = 1
  Const EA_m_mvmt_type_mvmt_cd1 = 2
  Const EA_m_mvmt_type_rcpt_flg1 = 3
  Const EA_m_mvmt_type_import_flg1 = 4
  Const EA_m_mvmt_type_ret_flg1 = 5
  Const EA_m_mvmt_type_subcontra_flg1 = 6
  Const EA_m_mvmt_type_usage_flg1 = 7
  Const EA_m_mvmt_type_ext1_cd1 = 8
  Const EA_m_mvmt_type_ext2_cd1 = 9
  Const EA_m_mvmt_type_ext3_cd1 = 10
  Const EA_m_mvmt_type_ext4_cd1 = 11
  Const EA_m_mvmt_type_subcontra2_flg1 = 12
  Const EA_m_mvmt_type_child_settle_flg1 = 13

  Dim iErrorPosition
  
    On Error Resume Next
    Err.Clear								    '☜: Protect system	from crashing

    Set	iPM1G439 = CreateObject("PM1G439.cMLookupMvmtTypeS")

    If CheckSYSTEMError(Err,True) = True Then
			Exit Sub
	End If
    
    I1_m_mvmt_type_io_type_cd 	= UCase(Request("txtMvmtType"))
    I2_ief_supplied_SELECT_CHAR = "5"		' 입고등록(1), 반품등록(2), 사급품출고등록(3)

    E1_m_mvmt_type = iPM1G439.M_LOOKUP_MVMT_TYPE_SVR(gStrGlobalCollection,I1_m_mvmt_type_io_type_cd,I2_ief_supplied_SELECT_CHAR)

    If CheckSYSTEMError2(Err, True, "","","","","") = True Then
			Set iPM1G439 = Nothing
			Response.Write "<Script Language=VBScript>" & vbCr
					Response.Write "With parent" & vbCr
					Response.Write ".frm1.txtMvmtType.Value 	= """"	" & vbCr
					Response.Write ".frm1.txtMvmtTypeNm.Value 	= """"	" & vbCr
					Response.Write ".frm1.hdnImportflg.Value 	= """"	" & vbCr
					Response.Write ".frm1.hdnRcptflg.Value 		= """"	" & vbCr
					Response.Write ".frm1.hdnRetflg.Value 		= """"	" & vbCr
					Response.Write ".frm1.hdnSubcontraflg.Value = """"	" & vbCr
					Response.Write ".frm1.hdnSubcontra2flg.Value = """"	" & vbCr
					Response.Write ".frm1.txtMvmtType.focus 	        " & vbCr
					Response.Write "End With"				& vbCr
				Response.Write "</Script>"					& vbCr
			Exit Sub
	End If

	Response.Write "<Script Language=vbscript>" & vbCr
	Response.Write "With parent" & vbCr												'☜: 화면 처리 ASP 를 지칭함 
    Response.Write ".frm1.txtMvmtType.Value 	= """ & ConvSPChars(E1_m_mvmt_type(EA_m_mvmt_type_io_type_cd1)) & """" & vbCr
	Response.Write ".frm1.txtMvmtTypeNm.Value 	= """ & ConvSPChars(E1_m_mvmt_type(EA_m_mvmt_type_io_type_nm1)) & """" & vbCr
	Response.Write ".frm1.hdnImportflg.Value 	= """ & ConvSPChars(E1_m_mvmt_type(EA_m_mvmt_type_import_flg1)) & """" & vbCr
	Response.Write ".frm1.hdnRcptflg.Value 		= """ & ConvSPChars(E1_m_mvmt_type(EA_m_mvmt_type_rcpt_flg1)) & """" & vbCr
	Response.Write ".frm1.hdnRetflg.Value 		= """ & ConvSPChars(E1_m_mvmt_type(EA_m_mvmt_type_ret_flg1)) & """" & vbCr
	Response.Write ".frm1.hdnSubcontraflg.Value = """ & ConvSPChars(E1_m_mvmt_type(EA_m_mvmt_type_child_settle_flg1)) & """" & vbCr
	Response.Write ".frm1.hdnSubcontra2flg.Value = """ & ConvSPChars(E1_m_mvmt_type(EA_m_mvmt_type_subcontra2_flg1)) & """" & vbCr
	
	Response.Write "End With"				& vbCr
	Response.Write "</Script>"					& vbCr

End Sub

'============================================================================================================
' Name : Display CodeName
' Desc : 
'============================================================================================================
Sub DisplaySupplierNm(inCode)
	
	On Error Resume Next						
    Err.Clear   
    
	Call SubOpenDB(lgObjConn)
	
	lgStrSQL = ""
	lgStrSQL = lgStrSQL & " SELECT BP_NM FROM B_BIZ_PARTNER " 
	lgStrSQL = lgStrSQL & " WHERE BP_CD =  " & FilterVar(inCode , "''", "S") & " AND Bp_Type in (" & FilterVar("S", "''", "S") & " ," & FilterVar("CS", "''", "S") & ") AND usage_flag=" & FilterVar("Y", "''", "S") & "  AND  in_out_flag = " & FilterVar("O", "''", "S") & "  "											'사외거래처만"		
	
	If FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X")  then
	
		Response.Write "<Script Language=VBScript>" & vbCr
		Response.Write "With parent.frm1" & vbCr
		Response.Write "	.txtSupplierNm.value	=	""" & lgObjRs("BP_NM") & """ " & vbCr
		Response.Write "End With" & vbCr
		Response.Write "</Script>" & vbCr
			
		Call SubCloseRs(lgObjRs)  
	Else
		Call DisplayMsgBox("179020", vbInformation, "", "", I_MKSCRIPT)
		
		Response.Write "<Script Language=VBScript>" & vbCr
		Response.Write "With parent.frm1" & vbCr
		Response.Write "	.txtSupplierNm.value	=	"""" " & vbCr
		Response.Write "	.txtSupplierCd.focus			 " & vbCr
		Response.Write "End With" & vbCr
		Response.Write "</Script>" & vbCr
		
	End if
	
	Call SubCloseDB(lgObjConn)
End Sub 

'============================================================================================================
' Name : Display GroupName
' Desc : 
'============================================================================================================
Sub DisplayGroupNm(inCode)
	
	On Error Resume Next						
    Err.Clear   
    
	Call SubOpenDB(lgObjConn)
	
	lgStrSQL = ""
	lgStrSQL = lgStrSQL & " SELECT PUR_GRP_NM FROM B_PUR_GRP " 
	lgStrSQL = lgStrSQL & " WHERE PUR_GRP =  " & FilterVar(inCode , "''", "S") & " AND USAGE_FLG=" & FilterVar("Y", "''", "S") & " "		
	
	If FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X")  then
	
		Response.Write "<Script Language=VBScript>" & vbCr
		Response.Write "With parent.frm1" & vbCr
		Response.Write "	.txtGroupNm.value	=	""" & lgObjRs("PUR_GRP_NM") & """ " & vbCr
		Response.Write "End With" & vbCr
		Response.Write "</Script>" & vbCr
			
		Call SubCloseRs(lgObjRs)  
	Else
		Call DisplayMsgBox("125100", vbInformation, "", "", I_MKSCRIPT)
		
		Response.Write "<Script Language=VBScript>" & vbCr
		Response.Write "With parent.frm1" & vbCr
		Response.Write "	.txtGroupNm.value	=	"""" " & vbCr
		Response.Write "	.txtGroupCd.focus			 " & vbCr
		Response.Write "End With" & vbCr
		Response.Write "</Script>" & vbCr
		
	End if
	
	Call SubCloseDB(lgObjConn)
End Sub 

'============================================================================================================
' Name : Display IvTypeName
' Desc : 
'============================================================================================================
Sub DisplayIvTypeNm(inCode)
	
	On Error Resume Next						
    Err.Clear   
    
	Call SubOpenDB(lgObjConn)
	
	lgStrSQL = ""
	lgStrSQL = lgStrSQL & " SELECT IV_TYPE_NM FROM M_IV_TYPE " 
	lgStrSQL = lgStrSQL & " WHERE IV_TYPE_CD =  " & FilterVar(inCode , "''", "S") & " AND IMPORT_FLG = " & FilterVar("N", "''", "S") & " "		
	
	If FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X")  then
	
		Response.Write "<Script Language=VBScript>" & vbCr
		Response.Write "With parent.frm1" & vbCr
		Response.Write "	.txtIvTypeNm.value	=	""" & lgObjRs("IV_TYPE_NM") & """ " & vbCr
		Response.Write "End With" & vbCr
		Response.Write "</Script>" & vbCr
			
		Call SubCloseRs(lgObjRs)  
	Else
		Call DisplayMsgBox("171800", vbInformation, "", "", I_MKSCRIPT)
		
		Response.Write "<Script Language=VBScript>" & vbCr
		Response.Write "With parent.frm1" & vbCr
		Response.Write "	.txtIvTypeNm.value	=	"""" " & vbCr
		Response.Write "	.txtIvTypeCd.focus			 " & vbCr
		Response.Write "End With" & vbCr
		Response.Write "</Script>" & vbCr
		
	End if
	
	Call SubCloseDB(lgObjConn)
End Sub 

'============================================================================================================
' Name : RemovedivTextArea
' Desc : 
'============================================================================================================
Sub RemovedivTextArea()
    On Error Resume Next                                                             
    Err.Clear                                                                        
	
	Response.Write "<Script language=vbs> " & vbCr   
    Response.Write "Parent.RemovedivTextArea "      & vbCr   
    Response.Write "</Script> "      & vbCr   
End Sub
%>
