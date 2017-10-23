<%@ LANGUAGE=VBSCript%>
<%Option Explicit    %>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<%	
call LoadBasisGlobalInf()
call LoadInfTB19029B("I", "*","NOCOOKIE","MB") 
call LoadBNumericFormatB("I","*","NOCOOKIE","MB")
'**********************************************************************************************
'*  1. Module Name          : Procurement
'*  2. Function Name        : 
'*  3. Program ID           : m9111ma1
'*  4. Program Name         :
'*  5. Program Desc         :
'*  6. Comproxy List        : PM9G111(Maint)
'							  PM9G112(확정)
'*  7. Modified date(First) : 2002/12/06
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Oh Chang Won
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'*                           
'*                           
'*                           
'*                           
'* 14. Business Logic of m9111ma1(재고이동요청)
'**********************************************************************************************
    Dim lgOpModeCRUD
    
    Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0                 '☜ : DBAgent Parameter 선언 
    Dim rs1, rs2, rs3, rs4,rs5
	Dim istrData
	Dim iStrPoNo
	Dim StrNextKey		' 다음 값 
	Dim lgStrPrevKey	' 이전 값 
	Dim iLngMaxRow		' 현재 그리드의 최대Row
	Dim iLngRow
	Dim GroupCount  
	Dim lgCurrency        
	Dim index,Count     ' 저장 후 Return 해줄 값을 넣을때 쓴는 변수     
    Dim lgDataExist
    Dim lgPageNo
    Dim lgMaxCount  
    
	Const C_SHEETMAXROWS_D  = 100
 
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    Call HideStatusWnd                                                               '☜: Hide Processing message
	'------ Developer Coding part (Start ) ------------------------------------------------------------------

	'------ Developer Coding part (End   ) ------------------------------------------------------------------ 

    lgOpModeCRUD  = Request("txtMode") 

										                                              '☜: Read Operation Mode (CRUD)
    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '☜: Query
             Call  SubBizQueryMulti()
        Case CStr(UID_M0002)                                                         '☜: Save,Update
             Call SubBizSaveMulti()
        Case CStr(UID_M0003)
             Call SubBizSaveMulti()
		Case "Release","UnRelease"			    '☜: 확정,확정취소 요청을 받음 
			 Call SubRelease()
        Case "LookUpItemByPlant"
             Call LookUpItemByPlant()
    End Select

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

End Sub    
'============================================================================================================
' Name : SubBizSave
' Desc : Save Data 
'============================================================================================================
Sub SubBizSave()
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
End Sub
'============================================================================================================
' Name : SubBizDelete
' Desc : Delete DB data
'============================================================================================================
Sub SubBizDelete()
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
End Sub

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()

    On Error Resume Next

	iStrPoNo = Trim(Request("txtPoNo"))
	lgPageNo       = UNICInt(Trim(Request("lgPageNo")),0)    '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)
	lgMaxCount     = CInt(Request("lgMaxCount"))   
	lgDataExist      = "No"
	iLngMaxRow = CLng(Request("txtMaxRows"))
	lgStrPrevKey = Request("lgStrPrevKey")
	
	Call FixUNISQLData()
	Call QueryData()	
	
	'====================
	'Call PO_DTL List
	'====================

	'-----------------------
	'Result data display area
	'----------------------- 
	Response.Write "<Script Language=vbscript>" & vbCr
	Response.Write "With parent" & vbCr
    
   Response.Write " If .frm1.vspdData.MaxRows < 1 then"						& vbCr
   Response.Write "	    If .frm1.txtRelease.Value = ""Y"" then"				& vbCr
   Response.Write "	        For index = .C_Po_Seq_No to .C_So_Seq_No"			& vbCr
   Response.Write "			     .ggoSpread.SpreadLock index , -1"		& vbCr
   Response.Write "		    Next"					& vbCr
   Response.Write "	    Else"						& vbCr
   Response.Write "		    .SetSpreadLock"			& vbCr
   Response.Write "		End If"						& vbCr
   Response.Write "	End if"							& vbCr
    
    
    Response.Write "	.ggoSpread.Source       = .frm1.vspdData "			& vbCr
    Response.Write "	.ggoSpread.SSShowData     """ & istrData	 & """" & vbCr	
    Response.Write "	.lgPageNo  = """ & lgPageNo   & """" & vbCr  
    Response.Write " .frm1.txthdnPoNo.value		= """ & ConvSPChars(Request("txtPoNo")) & """" & vbCr
	Response.Write "If .frm1.vspdData.MaxRows < .C_SHEETMAXROWS And .lgStrPrevKey <> """" Then " & vbCr
	Response.Write "	.SetSpreadLockAfterQuery " & vbCr
	Response.Write "	.DbQuery " & vbCr	' GroupView 사이즈로 화면 Row수보다 쿼리가 작으면 다시 쿼리함 
	Response.Write "Else " 			& vbCr
    Response.Write " .DbQueryOk "	& vbCr 
	Response.Write "End If "		& vbCr
    Response.Write "End With"		& vbCr
    Response.Write "</Script>"		& vbCr    
		
End Sub    

'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
' Query하기 전에  DB Agent 배열을 이용하여 Query문을 만드는 프로시져 
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()

    Dim strVal
	Dim arrVal(3)
	Redim UNISqlId(2)                                                     '☜: SQL ID 저장을 위한 영역확보 
    Redim UNIValue(2,2)                                                 '⊙: DB-Agent로 전송될 parameter를 위한 변수 
                                                                        '    parameter의 수에 따라 변경함 
    UNISqlId(0) = "m9111ma101" 											' header
    UNISqlId(1) = "m9111ma102"											' detail

   ' UNIValue(0,0) = Trim(lgSelectList)		                            '☜: Select 절에서 Summary    필드 



    UNIValue(0,1) = " " & FilterVar(iStrPoNo, "''", "S") & " "        'header
   
    UNIValue(1,1) = " " & FilterVar(iStrPoNo, "''", "S") & " "  	    'detail

    'UNIValue(0,UBound(UNIValue,2)) = ""
    UNIValue(1,UBound(UNIValue,2)) = "  ORDER BY A.PO_SEQ_NO ASC "			  '	UNISqlId(0)의 마지막 ?에 입력됨	
    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '☜: set ADO read mode
                        '☜: set ADO read mode
End Sub

'----------------------------------------------------------------------------------------------------------
' Query Data
' ADO의 Record Set이용하여 Query를 하고 Record Set을 넘겨서 MakeSpreadSheetData()으로 Spreadsheet에 데이터를 
' 뿌림 
' ADO 객체를 생성할때 prjPublic.dll파일을 이용한다.(상세내용은 vb로 작성된 prjPublic.dll 소스 참조)
'----------------------------------------------------------------------------------------------------------
Sub QueryData()
    Dim lgstrRetMsg                                             '☜ : Record Set Return Message 변수선언 
    Dim lgADF                                                   '☜ : ActiveX Data Factory 지정 변수선언 
    Dim iStr
    
    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3, rs4)

	Set lgADF   = Nothing
	
    iStr = Split(lgstrRetMsg,gColSep)

	If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If 
    
    'If SetConditionData = False Then Exit Sub 
    If  rs0.EOF And rs0.BOF  Then
        Call DisplayMsgBox("173132", vbOKOnly, "재고이동요청", "", I_MKSCRIPT)
        rs0.Close
        Set rs0 = Nothing
        Response.Write "<Script Language=vbscript>" & vbCr
        Response.Write " parent.FncNew "	& vbCr 
        Response.Write "</Script>"		& vbCr    
        Response.end
     End If  
   
    If  rs1.EOF And rs1.BOF  Then
        Call DisplayMsgBox("173132", vbOKOnly, "재고이동요청내역", "", I_MKSCRIPT)
        rs0.Close
        rs1.Close
        Set rs0 = Nothing
        Set rs1 = Nothing
        Response.Write "<Script Language=vbscript>" & vbCr
        Response.Write " parent.FncNew "	& vbCr 
        Response.Write "</Script>"		& vbCr           
        Response.end
    Else    
        Call  MakeHeaderData()
        Call  MakeSpreadSheetData()
    End If  
    
    
   
End Sub

Sub MakeHeaderData()
	
	Dim strPoType       '이동유형 
	Dim strPoTypeNm     '이동유형명 
    Dim strPoDt           '등록일 
    Dim strSupplierCd     '공급창고 
    Dim strSupplierNm     '공급창고명 
    Dim strGroupCd        '구매그룹 
	Dim strGroupNm        '구매그룹명 
    Dim strSuppPrsn       '공급처담당 
    Dim strTel            '긴급연락처 
    Dim strRemark 	      '비고 
    Dim strReleaseflg     '확정여부 
	Dim strClsflg         '마감여부 
	Dim strStosono        '수주번호 
	
	strPoType     = rs0(1)
	strPoTypeNm   = rs0(2)
    strSupplierCd = rs0(3)
    strSupplierNm = rs0(4)
    strGroupCd    = rs0(5)
	strGroupNm    = rs0(6)
    strPoDt       = rs0(7)
    strSuppPrsn   = rs0(8)
    strTel        = rs0(9)
    strReleaseflg = rs0(10)
    strClsflg     = rs0(11)
    strRemark 	  = rs0(12)
    strStosono    = rs0(13)

	Response.Write "<Script Language=vbscript>" & vbCr
	Response.Write "With parent" & vbCr
	Response.Write "if .frm1.vspdData.MaxRows = 0 then " & vbCr
'	Response.Write " 	parent.CurFormatNumericOCX " &vbCr
	Response.Write "	.frm1.txtSupplierCd.value = """ & Trim(UCase(ConvSPChars(strSupplierCd)))              	& """" & vbCr
	Response.Write "	.frm1.txtSupplierNm.value = """ & Trim(UCase(ConvSPChars(strSupplierNm)))              	& """" & vbCr
	Response.Write "	.frm1.txtGroupCd.value    = """ & Trim(UCase(ConvSPChars(strGroupCd)))                	& """" & vbCr
	Response.Write "	.frm1.txtGroupNm.value    = """ & Trim(UCase(ConvSPChars(strGroupNm)))              & """" & vbCr
	Response.Write "	.frm1.txtPoTypeCd.value   = """ & Trim(UCase(ConvSPChars(strPoType)))       & """" & vbCr
	Response.Write "	.frm1.txtPoTypeCdNm.value   = """ & Trim(UCase(ConvSPChars(strPoTypeNm)))       & """" & vbCr
	Response.Write "	.frm1.txtPoNo1.value    = """ & Trim(UCase(ConvSPChars(iStrPoNo)))               & """" & vbCr
	Response.Write "	.frm1.txtPoNo.value       = """ & Trim(UCase(ConvSPChars(iStrPoNo)))               & """" & vbCr
	Response.Write "	.frm1.txtPoDt.text       = """ & UNIDateClientFormat(strPoDt)       	& """" & vbCr

    Response.Write "	.frm1.txtSuppPrsn.Value		= """ & ConvSPChars(Trim(strSuppPrsn)) & """"	& vbCr
	Response.Write "	.frm1.txtTel.value				= """ & ConvSPChars(Trim(strTel)) 	& """"	& vbCr
    Response.Write "	.frm1.txtRemark.value			= """ & ConvSPChars(Trim(strRemark)) 		& """"	& vbCr
    Response.Write "	.frm1.hdnReleaseflg.value = """ & ConvSPChars(strReleaseflg) 			& """" & vbCr
	Response.Write "	.frm1.txtRelease.value = """ & ConvSPChars(strReleaseflg) 			& """" & vbCr
	Response.Write "	.frm1.hdnClsflg.value = """ & ConvSPChars(strClsflg) 			& """" & vbCr
	Response.Write "	.frm1.txtStoSoNo.value = """ & ConvSPChars(strStosono) 			& """" & vbCr
	If ConvSPChars(Trim(strReleaseflg)) = "Y" Then 
		Response.Write "	.frm1.rdoReleaseflg(0).Checked= true" 	& vbCr
	Else
		Response.Write "	.frm1.rdoReleaseflg(1).Checked= true" 	& vbCr
	End If		
	
	Response.Write "If parent.lgIntFlgMode = parent.parent.OPMD_CMODE Then parent.lgIntFlgMode = parent.parent.OPMD_UMODE "	& vbCr
	
'	Response.Write "  parent.CurFormatNumSprSheet "			& vbCr	'화폐별 라운딩 스프레드 포매팅 
	Response.Write " end if   "	& vbCr
	Response.Write " End With "	& vbCr
    Response.Write "</Script>" & vbCr
    
    rs0.Close                                                       '☜: Close recordset object
    Set rs0 = Nothing	       
End Sub


'----------------------------------------------------------------------------------------------------------
'QueryData()에 의해서 Query가 되면 MakeSpreadSheetData()에 의해서 데이터를 스프레드시트에 뿌려주는 프로시져 
'----------------------------------------------------------------------------------------------------------
Sub MakeSpreadSheetData()
    Dim iLoopCount                                                                     
    Dim iRowStr
    Dim ColCnt
    
    lgDataExist    = "Yes"
  
    If CLng(lgPageNo) > 0 Then
       rs1.Move     = CLng(lgMaxCount) * CLng(lgPageNo)                  'lgMaxCount:Max Fetched Count at once , lgStrPrevKeyIndex : Previous PageNo
    End If
   
   iLoopCount = 0
   Do while Not (rs1.EOF Or rs1.BOF)
        iLoopCount =  iLoopCount + 1
        iRowStr = ""
        iRowStr = iRowStr & Chr(11) & ConvSPChars(rs1(1))		'4
        iRowStr = iRowStr & Chr(11) & ConvSPChars(rs1(4))		'5
        iRowStr = iRowStr & Chr(11) & ""						'6								'6
        iRowStr = iRowStr & Chr(11) & ConvSPChars(rs1(5))		'7
        iRowStr = iRowStr & Chr(11) & ConvSPChars(rs1(6))	    '8                                '품목규격 '8	
        
        iRowStr = iRowStr & Chr(11) & UNINumClientFormat(rs1(8),ggQty.DecPoint,0) '9               '9	       
        iRowStr = iRowStr & Chr(11) & ConvSPChars(rs1(9))	    '10
        iRowStr = iRowStr & Chr(11) & ""						'11								'11
        iRowStr = iRowStr & Chr(11) & UNIDateClientFormat(rs1(7))	'12

		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs1(15))			'13
		iRowStr = iRowStr & Chr(11) & ""							'14								'27
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs1(16))	        '15
    
        iRowStr = iRowStr & Chr(11) & UNINumClientFormat(rs1(17),ggExchRate.DecPoint,0)	'16	
        iRowStr = iRowStr & Chr(11) & UNINumClientFormat(rs1(18),ggExchRate.DecPoint,0)	'17
        
        iRowStr = iRowStr & Chr(11) & ConvSPChars(rs1(12))	'18
		iRowStr = iRowStr & Chr(11) & ""		            '19
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs1(11))	'20
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs1(10))	'21
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs1(19))	'22
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs1(20))	'23
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs1(13))  '24
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs1(14))	'25
		
		'iRowStr = iRowStr & Chr(11) & ""		
											    
        iRowStr = iRowStr & Chr(11) & iLngMaxRow + iLoopCount                             

        If iLoopCount - 1 < lgMaxCount Then
           istrData = istrData & iRowStr & Chr(11) & Chr(12)
        Else
           lgPageNo = lgPageNo + 1
           Exit Do
        End If
        rs1.MoveNext
	Loop
    If iLoopCount < lgMaxCount Then                                      '☜: Check if next data exists
       lgPageNo = ""
    End If
    rs1.Close                                                       '☜: Close recordset object
    Set rs1 = Nothing	                                            '☜: Release ADF
End Sub


'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data 
'============================================================================================================
Sub SubBizSaveMulti()
	
	On Error Resume Next
	Err.Clear	
	
	Const M193_I2_po_no = 0										
	Const M193_I2_merg_pur_flg = 1                              
	Const M193_I2_pur_org = 2                                   
	Const M193_I2_pur_biz_area = 3                              
	Const M193_I2_pur_cost_cd = 4                               
	Const M193_I2_po_dt = 5                                     
	Const M193_I2_po_cur = 6                                    
	Const M193_I2_xch_rt = 7                                    
	Const M193_I2_pay_meth = 8                                  
	Const M193_I2_pay_dur = 9                                   
	Const M193_I2_pay_terms_txt = 10                            
	Const M193_I2_pay_type = 11                                 
	Const M193_I2_vat_type = 12                                 
	Const M193_I2_vat_rt = 13                                   
	Const M193_I2_tot_vat_doc_amt = 14                          
	Const M193_I2_tot_vat_loc_amt = 15                          
	Const M193_I2_tot_po_doc_amt = 16                           
	Const M193_I2_tot_po_loc_amt = 17                           
	Const M193_I2_sppl_sales_prsn = 18                          
	Const M193_I2_sppl_tel_no = 19                              
	Const M193_I2_release_flg = 20                              
	Const M193_I2_cls_flg = 21                                  
	Const M193_I2_import_flg = 22                               
	Const M193_I2_lc_flg = 23                                   
	Const M193_I2_bl_flg = 24                                   
	Const M193_I2_cc_flg = 25                                   
	Const M193_I2_rcpt_flg = 26                                 
	Const M193_I2_subcontra_flg = 27                            
	Const M193_I2_ret_flg = 28                                  
	Const M193_I2_iv_flg = 29                                   
	Const M193_I2_rcpt_type = 30                                
	Const M193_I2_issue_type = 31                               
	Const M193_I2_iv_type = 32                                  
	Const M193_I2_sppl_cd = 33                                  
	Const M193_I2_payee_cd = 34                                 
	Const M193_I2_build_cd = 35                                 
	Const M193_I2_remark = 36                                   
	Const M193_I2_manufacturer = 37                             
	Const M193_I2_agent = 38                                    
	Const M193_I2_applicant = 39                                
	Const M193_I2_offer_dt = 40                                 
	Const M193_I2_expiry_dt = 41                                
	Const M193_I2_transport = 42                                
	Const M193_I2_incoterms = 43                                
	Const M193_I2_delivery_plce = 44                            
	Const M193_I2_packing_cond = 45                             
	Const M193_I2_inspect_means = 46                            
	Const M193_I2_dischge_city = 47                             
	Const M193_I2_dischge_port = 48                             
	Const M193_I2_loading_port = 49                             
	Const M193_I2_origin = 50                                   
	Const M193_I2_sending_bank = 51                             
	Const M193_I2_invoice_no = 52                               
	Const M193_I2_fore_dvry_dt = 53                             
	Const M193_I2_shipment = 54                                 
	Const M193_I2_charge_flg = 55                               
	Const M193_I2_tracking_no = 56                              
	Const M193_I2_so_no = 57                                    
	Const M193_I2_inspect_method = 58                           
	Const M193_I2_ext1_cd = 59                                  
	Const M193_I2_ext1_qty = 60                                 
	Const M193_I2_ext1_amt = 61                                 
	Const M193_I2_ext1_rt = 62                                  
	Const M193_I2_ext1_dt = 63                                  
	Const M193_I2_ext2_cd = 64                                  
	Const M193_I2_ext2_qty = 65                                 
	Const M193_I2_ext2_amt = 66                                 
	Const M193_I2_ext2_rt = 67                                  
	Const M193_I2_ext2_dt = 68                                  
	Const M193_I2_ext3_cd = 69                                  
	Const M193_I2_ext3_qty = 70                                 
	Const M193_I2_ext3_amt = 71                                 
	Const M193_I2_ext3_rt = 72                                  
	Const M193_I2_ext3_dt = 73                                  
	Const M193_I2_xch_rate_op = 74                              
	Const M193_I2_vat_inc_flag = 75                             
	Const M193_I2_ref_no = 76                                   
    Const M193_I2_STO_FLG = 77
    Const M193_I2_SO_TYPE = 78
	
	
	Dim iPM9G111
	Dim lgIntFlgMode
	Dim iStrCommandSent
	Dim I1_b_company
	Dim I2_m_config_process
	Dim I3_b_biz_partner
	Dim I4_b_pur_grp
	Dim I5_m_pur_ord_hdr
	
	Redim I5_m_pur_ord_hdr(78)
	
    Const L1_status = 0
    Const L1_seq_no = 1
    Const L1_plant_cd = 2           '공장 
    Const L1_popup1 = 3
    Const L1_plant_nm = 4          '공장명 
    Const L1_item_cd = 5           '품목 
    Const L1_popup2 = 6
    Const L1_item_nm = 7            '품목명 
    Const L1_sppl_spec = 8          '품목규격 
    Const L1_order_qty = 9          '발주수량 
    Const L1_order_unit = 10        '단위 
    Const L1_popup3 = 11
    Const L1_cost = 12              '단가 
    Const L1_cost_con = 13          '단가구분 
    Const L1_cost_con_cd = 14       '단가구분코드 
    Const L1_order_amt = 15         '금액 
    Const L1_io_flg = 16            'VAT포함여부 
    Const L1_io_flg_cd = 17         'VAT포함여부코드 
    Const L1_vat_type = 18          'VAT
    Const L1_popup7 = 19
    Const L1_vat_nm = 20            'VAT명 
    Const L1_vat_rate = 21          'VAT율(%)
    Const L1_vat_amt = 22           'VAT금액 
    Const L1_dlvy_dt = 23           '납기일 
    Const L1_hs_cd = 24             'HS부호 
    Const L1_popup5 = 25
    Const L1_hs_nm = 26             'HS명 
    Const L1_sl_cd = 27             '창고 
    Const L1_popup6 = 28
    Const L1_sl_nm = 29             '창고명 
    Const L1_tracking_no = 30       'Tracking No.
    Const L1_tracking_popup = 31
    Const L1_lot_no = 32            'Lot No.
    Const L1_lot_seq = 33           'Lot No.순번 
    Const L1_ret_cd = 34            '반품유형 
    Const L1_popup8 = 35
    Const L1_ret_nm = 36            '반품유형명 
    Const L1_over = 37              '과부족허용율(+)(%)
    Const L1_under = 38             '과부족허용율(-)(%)
    Const L1_bal_qty = 39           'Bal. Qty.
    Const L1_bal_doc_amt = 40       'Bal. Doc. Amt.
    Const L1_bal_loc_amt = 41       'Bal. Loc. Amt.
    Const L1_ex_rate = 42           'Xch. Rate
    Const L1_pr_no = 43             '구매요청번호 
    Const L1_mvmt_no = 44           '구매입고번호 
    Const L1_po_no = 45             '발주번호 
    Const L1_po_seq_no = 46         '발주SeqNo
    Const L1_maint_seq = 47         'maintseq
    Const L1_so_no = 48
    Const L1_so_seq_no = 49
    Const L1_state_flg = 50
    Const L1_row_num = 51

	Dim LngMaxRow
	Dim iErrorPosition
	Dim iStrSpread
	
	'-------------------
	Dim itxtSpread
    Dim itxtSpreadArr
    Dim itxtSpreadArrCount, ii

    Dim iCUCount
    Dim iDCount
             
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
    '---------------------    
  
    lgIntFlgMode = CInt(Request("txtFlgMode"))
	
	LngMaxRow = CLng(Request("txtMaxRows"))											'☜: 최대 업데이트된 갯수 
             
    Set iPM9G111 = Server.CreateObject("PM9G111.cMMaintSto")    

    If CheckSYSTEMError(Err,True) = true Then 		
		Set iPM9G111 = Nothing												'☜: ComPlus Unload
		Exit Sub														'☜: 비지니스 로직 처리를 종료함 
	End if
	
	If lgIntFlgMode = OPMD_CMODE Then
		iStrCommandSent 							= "CREATE"
    ElseIf lgIntFlgMode = OPMD_UMODE and lgOpModeCRUD = CStr(UID_M0002) Then
    	I5_m_pur_ord_hdr(M193_I2_po_no)				= UCase(Trim(Request("txtPoNo")))
		iStrCommandSent 							= "UPDATE"
    ElseIf lgIntFlgMode = OPMD_UMODE and lgOpModeCRUD = CStr(UID_M0003) Then
        I5_m_pur_ord_hdr(M193_I2_po_no)				= UCase(Trim(Request("txtPoNo")))
		iStrCommandSent 							= "DELETE"
    End If

	I5_m_pur_ord_hdr(M193_I2_po_no)	           = Trim(Request("txtPoNo1"))       '요청번호 
	I2_m_config_process                        = Trim(Request("txtPoTypeCd"))    '이동유형 
	I5_m_pur_ord_hdr(M193_I2_po_dt)            = UNIConvDate(Trim(Request("txtPoDt")))        '등록일 
	I3_b_biz_partner                           = Trim(Request("txtSupplierCd"))  '공급창고 
    I4_b_pur_grp                               = Trim(Request("txtGroupCd"))     '구매그룹 
	I5_m_pur_ord_hdr(M193_I2_sppl_sales_prsn)  = Trim(Request("txtSuppPrsn"))    '공급처담당 
	I5_m_pur_ord_hdr(M193_I2_sppl_tel_no)      = Trim(Request("txtTel"))         '긴급연락처 
	I5_m_pur_ord_hdr(M193_I2_remark)           = Trim(Request("txtRemark"))      '비고 
	I5_m_pur_ord_hdr(M193_I2_release_flg)	   = "N"
	I1_b_company                               = gCurrency

    I5_m_pur_ord_hdr(M193_I2_merg_pur_flg)	   = "N"
    I5_m_pur_ord_hdr(M193_I2_cls_flg)          = "N"
    'I5_m_pur_ord_hdr(M193_I2_release_flg)	   = UCase(Request("hdnrelease"))

  	I5_m_pur_ord_hdr(M193_I2_tot_po_doc_amt)   = "0"
  	I5_m_pur_ord_hdr(M193_I2_tot_po_loc_amt)   = "0"
  	'I5_m_pur_ord_hdr(M193_I2_tot_vat_doc_amt)  = "0"
  	I5_m_pur_ord_hdr(M193_I2_po_cur)		   = gCurrency
  	I5_m_pur_ord_hdr(M193_I2_xch_rt)		   = "1"
    I5_m_pur_ord_hdr(M193_I2_xch_rate_op)      = "*"
    'I5_m_pur_ord_hdr(M193_I2_vat_type)		   = ""
  	'I5_m_pur_ord_hdr(M193_I2_pay_meth)		   = ""
  	'I5_m_pur_ord_hdr(M193_I2_pay_dur)		   = "0"
	
	
  	'I5_m_pur_ord_hdr(M193_I2_vat_rt)		   = ""
  	'I5_m_pur_ord_hdr(M193_I2_pay_terms_txt)	   = ""
  	'I5_m_pur_ord_hdr(M193_I2_pay_type)		   = ""
  	
  	'I5_m_pur_ord_hdr(M193_I2_vat_inc_flag)	   = ""

    'I5_m_pur_ord_hdr(M193_I2_applicant)		   = ""
	
	'iStrSpread = Trim(Request("txtSpread"))

	iStrPoNo = iPM9G111.M_STO(gStrGlobalCollection, _
						    iStrCommandSent, _
						    I1_b_company, _
						    I2_m_config_process, _
						    I3_b_biz_partner, _
						    I4_b_pur_grp, _
						    I5_m_pur_ord_hdr, _
						    LngMaxRow, _
						    gCurrency, _						             	
						    itxtSpread, _
						    iErrorPosition)
  
   If cStr(Err.Description) <> "" and iErrorPosition = "" Then 
      	If CheckSYSTEMError(Err,True) = True Then 
		   Set iPM9G111 = Nothing		                                                 '☜: Unload Comproxy DLL
		   Response.Write "<Script language=vbs> " & vbCr  
			Response.Write " Parent.RemovedivTextArea "      & vbCr
			Response.Write "</Script> "
			Exit Sub
		End If  
   elseif cStr(Err.Description) <> "" and iErrorPosition <> "" Then	
       If CheckSYSTEMError2(Err,True,iErrorPosition & "행:","","","","") = True Then
			Set iPM9G111 = Nothing
	 		Call SheetFocus(iErrorPosition, 2, I_MKSCRIPT)
			Response.Write "<Script language=vbs> " & vbCr  
			Response.Write " Parent.RemovedivTextArea "      & vbCr
			Response.Write "</Script> "
			Exit Sub
       End If
   end if
  
   Set iPM9G111 = Nothing                                                   '☜: Unload Comproxy  

   Response.Write "<Script language=vbs> " & vbCr
   Response.Write "   If  """ & Trim(iStrCommandSent) & """ = ""DELETE"" Then "	& vbCr	         
   Response.Write "       Parent.frm1.txtPoNo.Value = """ & ConvSPChars(Trim(iStrPoNo)) & """" & vbCr							'☜: 화면 처리 ASP 를 지칭함 
   Response.Write "       parent.FncNew1 "	& vbCr 
   Response.Write "   Else   " 					 & vbCr
   Response.Write "       Parent.frm1.txtPoNo.Value = """ & ConvSPChars(Trim(iStrPoNo)) & """" & vbCr							'☜: 화면 처리 ASP 를 지칭함 
   Response.Write "       Parent.DbSaveOk "      & vbCr							'☜: 화면 처리 ASP 를 지칭함 
   Response.Write "   End if   " 					 & vbCr
   Response.Write "</Script> "   & vbCr
       
End Sub    
'============================================================================================================
' Name : SubRelease
' Desc : 발주확정 
'============================================================================================================
Sub SubRelease()

	Dim PM9G112
	Dim strMode,lgIntFlgMode
	Dim txtSpread
    Dim pvCB
	reDim IG1_import_group(0,2)
    Const M155_IG1_I1_select_char = 0 
    Const M155_IG1_I1_count = 1
    Const M155_IG1_I2_po_no = 2

	Dim prErrorPosition 
	Dim E3_m_pur_ord_hdr_po_no
	
    On Error Resume Next
    Err.Clear																		'☜: Protect system from crashing
	
	lgIntFlgMode = CInt(Request("txtFlgMode"))										'☜: 저장시 Create/Update 판별 

	strMode = Request("txtMode")	
												'☜ : 현재 상태를 받음 
	pvCB = "T" '구주인경우 T
    '-----------------------
    'Data manipulate area
    '-----------------------
	txtSpread = "U" & gColSep
    if strMode = "Release" then
		txtSpread = txtSpread & "Y" & gColSep
	else
		txtSpread = txtSpread & "N" & gColSep
	End if

	txtSpread = txtSpread & Trim(Request("txtPoNo")) & gColSep
	txtSpread = txtSpread & Trim(Request("txtStoSoNo")) & gColSep
	txtSpread = txtSpread & "1" & gRowSep

	'⊙: Lookup Pad 동작후 정상적인 데이타 이면, 저장 로직 시작 
	
    Set PM9G112 = Server.CreateObject("PM9G112.cReleaseSto")    

    '-----------------------
    'Com action result check area(OS,internal)
    '-----------------------
    If CheckSYSTEMError(Err,True) = true then
			Set PM9G112 = Nothing 		
			Exit Sub
	End If

	'-----------------------
	'Com Action Area
	'-----------------------
	Call PM9G112.M_RELEASE_STO(gStrGlobalCollection, _
									  , _
									  txtSpread, _
									  pvCB, _
									  prErrorPosition)
	
	If CheckSYSTEMError2(Err,True,"","","","","") = true then 		
		Set M31211 = Nothing												'☜: ComProxy Unload
		Exit Sub															'☜: 비지니스 로직 처리를 종료함 
	 End If

    Set M31211 = Nothing                                                   '☜: Unload Comproxy

	'-----------------------
	'Result data display area
	'----------------------- 

	Response.Write "<Script Language=vbscript>" 					& vbCr
	Response.Write "With parent"									& vbCr	
	Response.Write ".DbSaveOk" & vbCr
	Response.Write "End With"   & vbCr
	Response.Write "</Script>" & vbCr

End Sub

'============================================================================================================
' Name : LookUpItemByPlant
' Desc : 
'============================================================================================================
Sub LookUpItemByPlant()
	
	Dim PB3S106
	Dim BasicUnit, SlCd, SlNm, ItemNm ,ItemSpec,trackingFlg,procurtype
	Dim activeRow
	
	 Const P003_E1_pur_org = 0
     Const P003_E1_pur_org_nm = 1
     Const P003_E1_valid_fr_dt = 2
     Const P003_E1_valid_to_dt = 3
     Const P003_E1_usage_flg = 4

    ' E2_b_item_group
     Const P003_E2_item_group_cd = 0
     Const P003_E2_item_group_nm = 1
     Const P003_E2_leaf_flg = 2

    ' E3_for_issued b_storage_location
     Const P003_E3_sl_cd = 0
     Const P003_E3_sl_type = 1
     Const P003_E3_sl_nm = 2

    ' E4_for_major b_storage_location
     Const P003_E4_sl_cd = 0
     Const P003_E4_sl_type = 1
     Const P003_E4_sl_nm = 2

    ' E5_i_material_valuation
     Const P003_E5_prc_ctrl_indctr = 0
     Const P003_E5_moving_avg_prc = 1
     Const P003_E5_std_prc = 2
     Const P003_E5_prev_std_prc = 3

    ' E6_b_item_by_plant
     Const P003_E6_procur_type = 0
     Const P003_E6_order_unit_mfg = 1
     Const P003_E6_order_lt_mfg = 2
     Const P003_E6_order_lt_pur = 3
     Const P003_E6_order_type = 4
     Const P003_E6_order_rule = 5
     Const P003_E6_req_round_flg = 6
     Const P003_E6_fixed_mrp_qty = 7
     Const P003_E6_min_mrp_qty = 8
     Const P003_E6_max_mrp_qty = 9
     Const P003_E6_round_qty = 10
     Const P003_E6_round_perd = 11
     Const P003_E6_scrap_rate_mfg = 12
     Const P003_E6_ss_qty = 13
     Const P003_E6_prod_env = 14
     Const P003_E6_mps_flg = 15
     Const P003_E6_issue_mthd = 16
     Const P003_E6_mrp_mgr = 17
     Const P003_E6_inv_check_flg = 18
     Const P003_E6_lot_flg = 19
     Const P003_E6_cycle_cnt_perd = 20
     Const P003_E6_inv_mgr = 21
     Const P003_E6_major_sl_cd = 22
     Const P003_E6_abc_flg = 23
     Const P003_E6_mps_mgr = 24
     Const P003_E6_recv_inspec_flg = 25
     Const P003_E6_inspec_lt_mfg = 26
     Const P003_E6_inspec_mgr = 27
     Const P003_E6_valid_from_dt = 28
     Const P003_E6_valid_to_dt = 29
     Const P003_E6_item_acct = 30
     Const P003_E6_single_rout_flg = 31
     Const P003_E6_prod_mgr = 32
     Const P003_E6_issued_sl_cd = 33
     Const P003_E6_issued_unit = 34
     Const P003_E6_order_unit_pur = 35
     Const P003_E6_var_lt = 36
     Const P003_E6_scrap_rate_pur = 37
     Const P003_E6_pur_org = 38
     Const P003_E6_prod_inspec_flg = 39
     Const P003_E6_final_inspec_flg = 40
     Const P003_E6_ship_inspec_flg = 41
     Const P003_E6_inspec_lt_pur = 42
     Const P003_E6_option_flg = 43
     Const P003_E6_over_rcpt_flg = 44
     Const P003_E6_over_rcpt_rate = 45
     Const P003_E6_damper_flg = 46
     Const P003_E6_damper_max = 47
     Const P003_E6_damper_min = 48
     Const P003_E6_reorder_pnt = 49
     Const P003_E6_std_time = 50
     Const P003_E6_add_sel_rule = 51
     Const P003_E6_add_sel_value = 52
     Const P003_E6_add_seq_rule = 53
     Const P003_E6_add_seq_atrid = 54
     Const P003_E6_rem_sel_rule = 55
     Const P003_E6_rem_sel_value = 56
     Const P003_E6_rem_seq_rule = 57
     Const P003_E6_rem_seq_atrid = 58
     Const P003_E6_llc = 59
     Const P003_E6_tracking_flg = 60
     Const P003_E6_valid_flg = 61
     Const P003_E6_work_center = 62
     Const P003_E6_order_from = 63
     Const P003_E6_cal_type = 64
     Const P003_E6_line_no = 65
     Const P003_E6_atp_lt = 66
     Const P003_E6_etc_flg1 = 67
     Const P003_E6_etc_flg2 = 68

    ' E7_b_item
     Const P003_E7_item_cd = 0
     Const P003_E7_item_nm = 1
     Const P003_E7_formal_nm = 2
     Const P003_E7_spec = 3
     Const P003_E7_item_acct = 4
     Const P003_E7_item_class = 5
     Const P003_E7_hs_cd = 6
     Const P003_E7_hs_unit = 7
     Const P003_E7_unit_weight = 8
     Const P003_E7_unit_of_weight = 9
     Const P003_E7_basic_unit = 10
     Const P003_E7_draw_no = 11
     Const P003_E7_item_image_flg = 12
     Const P003_E7_phantom_flg = 13
     Const P003_E7_blanket_pur_flg = 14
     Const P003_E7_base_item_cd = 15
     Const P003_E7_proportion_rate = 16
     Const P003_E7_valid_flg = 17
     Const P003_E7_valid_from_dt = 18
     Const P003_E7_valid_to_dt = 19

    ' E8_b_plant
     Const P003_E8_plant_cd = 0
     Const P003_E8_plant_nm = 1


	On Error Resume Next
	Err.Clear

	Set	PB3S106 = CreateObject("PB3S106.cBLkUpItemByPlt")

    '-----------------------
    'Com action result check area(OS,internal)
    '-----------------------
    If CheckSYSTEMError(Err,True) = true then
		Set PB3S106 = Nothing 		
		Exit Sub
	End If	

    dim I1_select_char 
    dim I2_b_item 
    dim I3_plant_cd 
    dim E1_b_pur_org 
    dim E2_b_item_group 
    dim E3_for_issued_b_storage_location 
    dim E4_for_major_b_storage_location 
    dim E5_i_material_valuation 
    dim E6_b_plant 
    dim E7_b_item 
    dim E8_b_item_by_plant 
    dim prStatusCodePrevNext 
	
	activeRow					= Request("SpreadActiveRow")
	I2_b_item               	= Request("txtItemCd")
	I3_plant_cd             	= Request("txtPlantCd")
	

	Call PB3S106.B_LOOK_UP_ITEM_BY_PLANT_SVR (gStrGlobalCollection, I1_select_char, I3_plant_cd , I2_b_item, _ 
	                                         E1_b_pur_org, E2_b_item_group ,E3_for_issued_b_storage_location, _
	                                         E4_for_major_b_storage_location, E5_i_material_valuation , _
	                                         E6_b_plant, E7_b_item ,E8_b_item_by_plant ,prStatusCodePrevNext)

	If CheckSYSTEMError(Err,True) = true then
		Set PB3S106 = Nothing 		
		Exit Sub
	End If	
	
	
	ItemNm		= ConvSPChars(E7_b_item(P003_E7_item_nm))
	ItemSpec    = ConvSPChars(E7_b_item(P003_E7_spec))
	BasicUnit	= ConvSPChars(E7_b_item(P003_E7_basic_unit))
	SlCd		= ConvSPChars(E3_for_issued_b_storage_location(P003_E3_sl_cd))
	SlNm		= ConvSPChars(E3_for_issued_b_storage_location(P003_E3_sl_nm))
	trackingFlg = ConvSPChars(E8_b_item_by_plant(P003_E6_tracking_flg))
	procurtype  = ConvSPChars(E8_b_item_by_plant(P003_E6_procur_type))
	
	Response.Write "<Script language=vbs> "		& vbCr         
	Response.Write " Dim assignRow "			& vbCr
    Response.Write " Dim IntRetCD "			& vbCr
    Response.Write " With Parent.frm1 "			& vbCr
    'Response.Write "   If  """ & Trim(UCase(procurtype)) & """ <> ""P"" Then " 									& vbCr	
    'Response.Write "       	IntRetCD = Parent.DisplayMsgBox(""179019"",""X"",""X"",""X"") " & vbCr
    'Response.Write "   else "             & vbCr
    Response.Write "       .vspdData.row			= """ & activeRow & """"			& vbCr
    Response.Write "       .vspdData.col			= parent.C_itemNm "			        & vbCr
    Response.Write "       .vspdData.text			= """ & ItemNm & """"				& vbCr
    Response.Write "       .vspdData.col			= parent.C_SpplSpec "			    & vbCr
    Response.Write "       .vspdData.text			= """ & ItemSpec & """"				& vbCr
    Response.Write "       .vspdData.col			= parent.C_OrderUnit "				& vbCr
    Response.Write "       .vspdData.text			= """ & BasicUnit & """"			& vbCr
    Response.Write "       .vspdData.col			= parent.C_SLCd "					& vbCr
    Response.Write "       .vspdData.text			= """ & SlCd & """"					& vbCr
    Response.Write "       .vspdData.col			= parent.C_SLNm "					& vbCr
    Response.Write "       .vspdData.text			= """ & SlNm & """"					& vbCr
	Response.Write "       If  """ & Trim(UCase(trackingFlg)) & """ <> ""Y"" Then " 									& vbCr	
	Response.Write "  	       parent.ggoSpread.spreadlock parent.C_TrackingNo, .vspdData.Row, parent.C_TrackingNoPop, .vspdData.Row " 	& vbCr	
	Response.Write "      	   .vspdData.Col 	= Parent.C_TrackingNo "    														& vbCr	
	Response.Write "  		   .vspdData.text   = ""*""" 																	& vbCr
	Response.Write "       Else   " 																					& vbCr
	Response.Write "   	       parent.ggoSpread.spreadUnlock parent.C_TrackingNo, .vspdData.Row, parent.C_TrackingNoPop, .vspdData.Row   " & vbCr
	Response.Write "   		   parent.ggoSpread.sssetrequired parent.C_TrackingNo,.vspdData.Row, .vspdData.Row   " 						& vbCr
	Response.Write "   		   .vspdData.Col 	= parent.C_TrackingNo   " 						& vbCr
	Response.Write "   		   .vspdData.text = """"   " 						& vbCr
	Response.Write "       End If "             & vbCr
    'Response.Write "   End If " 
    Response.Write " End With "							& vbCr	
    Response.Write "</Script> " 


End Sub


'============================================================================================================
' Name : SubBizSaveCreate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveMultiCreate(arrColVal)
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
    
    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to create record
    '--------------------------------------------------------------------------------------------------------
    
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
End Sub
'============================================================================================================
' Name : SubBizSaveMultiUpdate
' Desc : Update Data from Db
'============================================================================================================
Sub SubBizSaveMultiUpdate(arrColVal)

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to update record
    '--------------------------------------------------------------------------------------------------------
    
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
End Sub
'============================================================================================================
' Name : SubBizSaveMultiDelete
' Desc : Delete Data from Db
'============================================================================================================
Sub SubBizSaveMultiDelete(arrColVal)

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to update record
    '--------------------------------------------------------------------------------------------------------

    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
End Sub
'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements(pDataType,arrColVal)
    Dim iSelCount
    
    On Error Resume Next

	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub
'============================================================================================================
' Name : CommonOnTransactionCommit
' Desc : This Sub is called by OnTransactionCommit Error handler
'============================================================================================================
Sub CommonOnTransactionCommit()
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub

'============================================================================================================
' Name : CommonOnTransactionAbort
' Desc : This Sub is called by OnTransactionAbort Error handler
'============================================================================================================
Sub CommonOnTransactionAbort()
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub

'============================================================================================================
' Name : SetErrorStatus
' Desc : This Sub set error status
'============================================================================================================
Sub SetErrorStatus()
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub
'============================================================================================================
' Name : SubHandleError
' Desc : This Sub handle error
'============================================================================================================
Sub SubHandleError(pOpCode,pConn,pRs,pErr)
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status


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
