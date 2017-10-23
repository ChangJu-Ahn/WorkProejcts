<%@ LANGUAGE=VBSCript%>
<%Option Explicit    %>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../ComASP/LoadInfTb19029.asp" -->
<%
'**********************************************************************************************
'*  1. Module Name          : 영업 
'*  2. Function Name        : 출하관리 
'*  3. Program ID           : S4210MA1
'*  4. Program Name         : 출하요청등록 
'*  5. Program Desc         :
'*  6. Comproxy List        : S42101MaintDnHdrSvr
'*  7. Modified date(First) : 2000/04/17
'*  8. Modified date(Last)  : 2002/06/19
'*  9. Modifier (First)     : Cho song hyon
'* 10. Modifier (Last)      : Lee Myung Wha
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'*                            -2000/04/17 : 화면 Layout & ASP Coding
'*                            -2001/12/19 : Date 표준적용 
'**********************************************************************************************

	Call LoadBasisGlobalInf()
	Call LoadInfTB19029B( "I", "*", "NOCOOKIE", "MB")
	Call LoadBNumericFormatB("I", "*", "NOCOOKIE", "MB")

	Dim lgOpModeCRUD
    
    on Error Resume Next                                                             
    Err.Clear                                                                        

    Call HideStatusWnd                                                               '☜: Hide Processing message
    
    lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)
    
    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '☜: Query
	         Call SubBizQuery()
        Case CStr(UID_M0002)
	         Call SubBizSave()
        Case CStr(UID_M0003)                                                         '☜: Delete
             Call SubBizDelete()
    End Select

'============================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================
Sub SubBizQuery()

	on Error Resume Next
    
    '==================== 조회결과 Index ==================
	Const S5G211_RS_DN_NO = 0                ' 출하번호 
	Const S5G211_RS_SHIP_TO_PARTY = 1        ' 납품처 
	Const S5G211_RS_SHIP_TO_PARTY_NM = 2     ' 납품처명 
	Const S5G211_RS_SALES_GRP = 3            ' 영업그룹 
	Const S5G211_RS_SALES_GRP_NM = 4         ' 영업그룹명 
	Const S5G211_RS_SALES_ORG = 5            ' 영업조직 
	Const S5G211_RS_SALES_ORG_NM = 6         ' 영업조직명 
	Const S5G211_RS_MOV_TYPE = 7             ' 출하형태 
	Const S5G211_RS_MOV_TYPE_NM = 8          ' 출하형태명 
	Const S5G211_RS_DLVY_DT = 9             ' 납기일 
	Const S5G211_RS_PROMISE_DT = 10          ' 출고예정일 
	Const S5G211_RS_ACTUAL_GI_DT = 11        ' 출고일 
	Const S5G211_RS_COST_CD = 12             ' Cost Center
	Const S5G211_RS_BIZ_AREA = 13            ' 사업장 
	Const S5G211_RS_BIZ_AREA_NM = 14         ' 사업장명 
	Const S5G211_RS_TRANS_METH = 15          ' 운송방법 
	Const S5G211_RS_TRANS_METH_NM = 16       ' 운송방법명 
	Const S5G211_RS_GOODS_MV_NO = 17         ' 수불번호(출고번호)
	Const S5G211_RS_CI_FLAG = 18             ' 통관필요여부 
	Const S5G211_RS_POST_FLAG = 19           ' 출고처리여부 
	Const S5G211_RS_SO_TYPE = 20             ' 수주형태 
	Const S5G211_RS_SO_TYPE_NM = 21          ' 수주형태명 
	Const S5G211_RS_SO_NO = 22               ' 수주번호 
	Const S5G211_RS_SHIP_TO_PLCE = 23        ' 납품장송 
	Const S5G211_RS_INSRT_USER_ID = 24       ' 등록자 
	Const S5G211_RS_INSRT_DT = 25            ' 등록일 
	Const S5G211_RS_UPDT_USER_ID = 26        ' 변경자 
	Const S5G211_RS_UPDT_DT = 27             ' 변경일 
	Const S5G211_RS_EXT1_QTY = 28            ' 여유필드(수량)
	Const S5G211_RS_EXT2_QTY = 29            ' 여유필드(수량)
	Const S5G211_RS_EXT3_QTY = 30            ' 여유필드(수량)
	Const S5G211_RS_EXT1_AMT = 31            ' 여유필드(금액)
	Const S5G211_RS_EXT2_AMT = 32            ' 여유필드(금액)
	Const S5G211_RS_EXT3_AMT = 33            ' 여유필드(금액)
	Const S5G211_RS_EXT1_CD = 34             ' 여유필드(Text)
	Const S5G211_RS_EXT2_CD = 35             ' 여유필드(Text)
	Const S5G211_RS_EXT3_CD = 36             ' 여유필드(Text)
	Const S5G211_RS_TEMP_SO_NO = 37          ' 수주번호 
	Const S5G211_RS_VAT_FLAG = 38            ' 세금계산서정보 동시생성여부 
	Const S5G211_RS_AR_FLAG = 39             ' 매출정보 동시생성여부 
	Const S5G211_RS_CUR = 40                 ' 화폐단위 
	Const S5G211_RS_XCHG_RATE = 41           ' 환율 
	Const S5G211_RS_XCHG_RATE_OP = 42        ' 환율연산자 
	Const S5G211_RS_NET_AMT = 43             ' 출고금액 
	Const S5G211_RS_NET_AMT_LOC = 44         ' 출고자국금액 
	Const S5G211_RS_VAT_AMT = 45             ' VAT 금액 
	Const S5G211_RS_VAT_AMT_LOC = 46         ' VAT 자국금액 
	Const S5G211_RS_EXCEPT_DN_FLAG = 47      ' 예외출고여부 
	Const S5G211_RS_REMARK = 48              ' 비고 
	Const S5G211_RS_ARRIVAL_DT = 49          ' 실제납품일 
	Const S5G211_RS_ARRIVAL_TIME = 50        ' 납품시간 
	Const S5G211_RS_STP_INFO_NO = 51         ' 납품처상세정보 번호 
	Const S5G211_RS_ZIP_CD = 52              ' 납품처 우편번호 
	Const S5G211_RS_ADDR1 = 53               ' 납품처 주소 
	Const S5G211_RS_ADDR2 = 54               ' 납품처 주소 
	Const S5G211_RS_ADDR3 = 55               ' 납품처 주소 
	Const S5G211_RS_RECEIVER = 56            ' 인수자명 
	Const S5G211_RS_TEL_NO1 = 57             ' 전화번호1
	Const S5G211_RS_TEL_NO2 = 58             ' 전화번호2
	Const S5G211_RS_TRANS_INFO_NO = 59       ' 운송정보번호 
	Const S5G211_RS_TRANS_CO = 60            ' 운송회사 
	Const S5G211_RS_DRIVER = 61              ' 운전자명 
	Const S5G211_RS_VEHICLE_NO = 62          ' 차량번호 
	Const S5G211_RS_SENDER = 63              ' 인계자명 
	Const S5G211_RS_STO_FLAG = 64            ' STO Flag
	Const S5G211_RS_CASH_DC_AMT = 65         ' 현금할인액 
	Const S5G211_RS_TAX_DC_AMT = 66          ' 세금할인액 
	Const S5G211_RS_TAX_BASE_AMT = 67        ' 세금계산기초금액 
	Const S5G211_RS_CASH_DC_AMT_LOC = 68     ' 현금할인액(자국)
	Const S5G211_RS_TAX_DC_AMT_LOC = 69      ' 세금할인액(자국)
	Const S5G211_RS_TAX_BASE_AMT_LOC = 70    ' 세금계산기초금액(자국)
	Const S5G211_RS_SO_AUTO_FLAG = 71        ' 수주로부터 자동생성여 여부 
	Const S5G211_RS_PLANT_CD = 72            ' 공장 
	Const S5G211_RS_PLANT_NM = 73            ' 공장명 
	Const S5G211_RS_INV_MGR = 74             ' 재고담당 
	Const S5G211_RS_INV_MGR_NM = 75          ' 재고담당자명 
    Const S5G211_RS_CONTRY_CD = 76			 ' 납품처 국가코드 
	Const S5G211_RS_SCM_DO_NO = 80
	Const S5G211_RS_SCM_DO_SEQ = 81
    
    Dim iStrDnNo
	Dim iObjS5G2211
	Dim iArrSDnReqHdr
	
	iStrDnNo = Trim(Request("txtConDnNo"))			' 출하번호 
	
    Set iObjS5G2211 = Server.CreateObject("PS5G221.cLookUpSDnReqHdr")
	
	If CheckSYSTEMError(Err,True) = True Then
       Response.Write "<Script Language=vbscript>"			& vbCr
	   Response.Write "parent.frm1.txtConDnNo.focus"		& vbCr    
	   Response.Write "</Script>"							& vbCr	
       Exit Sub
    End If  
    
    iArrSDnReqHdr = iObjS5G2211.LookUp(gStrGlobalCollection, iStrDnNo, "N")
    
    If CheckSYSTEMError(Err,True) = True Then
       Set iObjS5G2211 = Nothing		                                                 '☜: Unload Comproxy DLL
       Response.Write "<Script Language=vbscript>"			& vbCr
	   Response.Write "parent.frm1.txtConDnNo.focus"		& vbCr    
	   Response.Write "</Script>"							& vbCr	
       Exit Sub
    End If  
    
    Set iObjS5G2211 = Nothing
	
	Response.Write "<Script Language=vbscript>" & vbCr
	Response.Write "With parent.frm1" & vbCr
	
	'--출하번호--
	Response.Write ".txtDnNo.value				= """ & ConvSPChars(iArrSDnReqHdr(S5G211_RS_DN_NO))				& """" & vbcr
	'--출고예정일--
	Response.Write ".txtPlanned_gi_dt.Text		= """ & UNIDateClientFormat(iArrSDnReqHdr(S5G211_RS_PROMISE_DT))	& """" & vbcr
	'--실제출고일--
	Response.Write ".txtActGi_dt.value			= """ & UNIDateClientFormat(iArrSDnReqHdr(S5G211_RS_ACTUAL_GI_DT))	& """" & vbcr
	'--납품장소--
	Response.Write ".txtShip_to_place.value		= """ & ConvSPChars(iArrSDnReqHdr(S5G211_RS_SHIP_TO_PLCE))			& """" & vbcr
	'--운송방법--
	Response.Write ".txtTrans_meth.value		= """ & ConvSPChars(iArrSDnReqHdr(S5G211_RS_TRANS_METH))			& """" & vbcr
	'--운송방법명--
	Response.Write ".txtTrans_meth_nm.value		= """ & ConvSPChars(iArrSDnReqHdr(S5G211_RS_TRANS_METH_NM))		& """" & vbcr
	'--납기일--
	Response.Write ".txtDlvy_dt.value			= """ & UNIDateClientFormat(iArrSDnReqHdr(S5G211_RS_DLVY_DT))		& """" & vbcr
	'--영업그룹--
	Response.Write ".txtSales_Grp.value			= """ & ConvSPChars(iArrSDnReqHdr(S5G211_RS_SALES_GRP))			& """" & vbcr
	'--영업그룹명--
	Response.Write ".txtSales_GrpNm.value		= """ & ConvSPChars(iArrSDnReqHdr(S5G211_RS_SALES_GRP_NM))			& """" & vbcr
	'--출하타입--
	Response.Write ".txtMovType.value			= """ & ConvSPChars(iArrSDnReqHdr(S5G211_RS_MOV_TYPE))				& """" & vbcr
	'--출하타입명--
	Response.Write ".txtMovTypeNm.value			= """ & ConvSPChars(iArrSDnReqHdr(S5G211_RS_MOV_TYPE_NM))			& """" & vbcr
	'--출고번호--
	Response.Write ".txtGoods_mv_no.value		= """ & Trim(ConvSPChars(iArrSDnReqHdr(S5G211_RS_GOODS_MV_NO)))	& """" & vbcr
	'--수주번호--
	Response.Write ".txtSo_no.value				= """ & ConvSPChars(iArrSDnReqHdr(S5G211_RS_SO_NO))				& """" & vbcr
	'--수주번호지정--
	If iArrSDnReqHdr(S5G211_RS_SO_NO) <> "" Then
		Response.Write ".chkSoNo.checked = True" & vbcr
	End If
	'--수주타입--
	Response.Write ".txtSo_Type.value			= """ & ConvSPChars(iArrSDnReqHdr(S5G211_RS_SO_TYPE))				& """" & vbcr
	'--수주타입명--
	Response.Write ".txtSo_TypeNm.value			= """ & ConvSPChars(iArrSDnReqHdr(S5G211_RS_SO_TYPE_NM))			& """" & vbcr
	'--납품처--
	Response.Write ".txtShip_to_party.value		= """ & ConvSPChars(iArrSDnReqHdr(S5G211_RS_SHIP_TO_PARTY))		& """" & vbcr
	'--납품처명--
	Response.Write ".txtShip_to_partyNm.value	= """ & ConvSPChars(iArrSDnReqHdr(S5G211_RS_SHIP_TO_PARTY_NM))		& """" & vbcr
	'--국가코드--
	Response.Write ".txtHCntryCd.value	= """ & ConvSPChars(iArrSDnReqHdr(S5G211_RS_CONTRY_CD))			& """" & vbcr
	'--비고--
	Response.Write ".txtRemark.value	= """ & ConvSPChars(iArrSDnReqHdr(S5G211_RS_REMARK))				& """" & vbcr

	'--비고--
	Response.Write ".txtSCMDONo.value	= """ & ConvSPChars(iArrSDnReqHdr(S5G211_RS_SCM_DO_NO))				& """" & vbcr
	'--비고--
	Response.Write ".txtSCMDONoSeq.value	= """ & ConvSPChars(iArrSDnReqHdr(S5G211_RS_SCM_DO_SEQ))				& """" & vbcr


	'--실제납품일--
	Response.Write ".txtArriv_dt.Text	= """ & UNIDateClientFormat(iArrSDnReqHdr(S5G211_RS_ARRIVAL_DT))	& """" & vbcr
	'--실제납품시간--
	Response.Write ".txtArriv_tm.value	= """ & ConvSPChars(iArrSDnReqHdr(S5G211_RS_ARRIVAL_TIME))			& """" & vbcr
	'--납품처상세정보번호--
	Response.Write ".txtSTP_Inf_No.value	= """ & ConvSPChars(iArrSDnReqHdr(S5G211_RS_STP_INFO_NO))		& """" & vbcr
	'--운송정보번호--
	Response.Write ".txtTrnsp_Inf_No.value	= """ & ConvSPChars(iArrSDnReqHdr(S5G211_RS_TRANS_INFO_NO))	& """" & vbcr
	'--우편번호--
	Response.Write ".txtZIP_cd.value	= """ & ConvSPChars(iArrSDnReqHdr(S5G211_RS_ZIP_CD))			& """" & vbcr
	'--주소1--
	Response.Write ".txtADDR1_Dlv.value	= """ & ConvSPChars(iArrSDnReqHdr(S5G211_RS_ADDR1))			& """" & vbcr
	'--주소2--
	Response.Write ".txtADDR2_Dlv.value	= """ & ConvSPChars(iArrSDnReqHdr(S5G211_RS_ADDR2))			& """" & vbcr
	'--주소3--
	Response.Write ".txtADDR3_Dlv.value	= """ & ConvSPChars(iArrSDnReqHdr(S5G211_RS_ADDR3))			& """" & vbcr
	'--인수자명--
	Response.Write ".txtReceiver.value	= """ & ConvSPChars(iArrSDnReqHdr(S5G211_RS_RECEIVER))			& """" & vbcr
	'--전화번호1--
	Response.Write ".txtTel_No1.value	= """ & ConvSPChars(iArrSDnReqHdr(S5G211_RS_TEL_NO1))			& """" & vbcr
	'--전화번호2--
	Response.Write ".txtTel_No2.value	= """ & ConvSPChars(iArrSDnReqHdr(S5G211_RS_TEL_NO2))			& """" & vbcr
	'--운송회사--
	Response.Write ".txtTransCo.value	= """ & ConvSPChars(iArrSDnReqHdr(S5G211_RS_TRANS_CO))			& """" & vbcr
	'--운전자--
	Response.Write ".txtDriver.value	= """ & ConvSPChars(iArrSDnReqHdr(S5G211_RS_DRIVER))			& """" & vbcr
	'--차량번호--
	Response.Write ".txtVehicleNo.value	= """ & ConvSPChars(iArrSDnReqHdr(S5G211_RS_VEHICLE_NO))		& """" & vbcr
	'--인계자--
	Response.Write ".txtSender.value	= """ & ConvSPChars(iArrSDnReqHdr(S5G211_RS_SENDER))			& """" & vbcr
	'--공장--
	Response.Write ".txtPlantCd.value	= """ & ConvSPChars(iArrSDnReqHdr(S5G211_RS_PLANT_CD))			& """" & vbcr
	Response.Write ".txtPlantNm.value	= """ & ConvSPChars(iArrSDnReqHdr(S5G211_RS_PLANT_NM))			& """" & vbcr
	'--재고담당자--
	Response.Write ".txtInvMgr.value	= """ & ConvSPChars(iArrSDnReqHdr(S5G211_RS_INV_MGR))			& """" & vbcr
	Response.Write ".txtInvMgrNm.value	= """ & ConvSPChars(iArrSDnReqHdr(S5G211_RS_INV_MGR_NM))		& """" & vbcr

	Response.Write "parent.DbQueryOk" & vbCr
	Response.Write "End With"          & vbCr
    Response.Write "</Script>"         & vbCr															'☜: 조회가 성공 
    Response.End
    
End Sub

'============================================
' Name : SubBizSave
' Desc : Date data 
'============================================
Sub SubBizSave()
	on Error Resume Next

    Dim iStrCUDFlag, iStrFixSoNoFlag, iStrSTPInfoNo, iStrTransInfoNo, iStrDnNo
	Dim iStrCrSTPFlag, iStrCrTransFlag, pvCB, lgIntFlgMode
    Dim iObjS5G221
    
    lgIntFlgMode = CInt(Request("txtFlgMode"))									'☜: 저장시 Create/Update 판별 

    If lgIntFlgMode = OPMD_CMODE Then
		iStrCUDFlag = "C"
    ElseIf lgIntFlgMode = OPMD_UMODE Then
		iStrCUDFlag = "U"
	Else
		Call ServerMesgBox("TXTFLGMODE 조건값이 비어있습니다!", vbInformation, I_MKSCRIPT)              
		Response.End 
    End If
    
    ' 출하정보 
    Dim iArrDnHdr

    Const S5G221_DnReqHdr_DN_REQ_NO = 0           '(O)출하번호 
	Const S5G221_DnReqHdr_EXCEPT_DN_FLAG = 1  '(M)예외출고여부(Y/N)	
    Const S5G221_DnReqHdr_SHIP_TO_PARTY = 2   '(M)납품처 
    Const S5G221_DnReqHdr_SALES_GRP = 3       '(M)영업그룹 
    Const S5G221_DnReqHdr_MOV_TYPE = 4        '(M)출하형태 
    Const S5G221_DnReqHdr_DLVY_DT = 5         '(M)납기일 
    Const S5G221_DnReqHdr_PROMISE_DT = 6      '(M)출고예정일 
    Const S5G221_DnReqHdr_TRANS_METH = 7      '(O)운송방법 
    Const S5G221_DnReqHdr_SO_TYPE = 8         '(M)수주형태 
    Const S5G221_DnReqHdr_SO_NO = 9           '(O)S/O번호 
    Const S5G221_DnReqHdr_SHIP_TO_PLCE = 10    '(O)납품장소 
    Const S5G221_DnReqHdr_EXT1_QTY = 11       '(O)여유필드(수량)
    Const S5G221_DnReqHdr_EXT2_QTY = 12       '(O)여유필드(수량)
    Const S5G221_DnReqHdr_EXT3_QTY = 13       '(O)여유필드(수량)
    Const S5G221_DnReqHdr_EXT1_AMT = 14       '(O)여유필드(금액)
    Const S5G221_DnReqHdr_EXT2_AMT = 15       '(O)여유필드(금액)
    Const S5G221_DnReqHdr_EXT3_AMT = 16       '(O)여유필드(금액)
    Const S5G221_DnReqHdr_EXT1_CD = 17        '(O)여유필드(Text)
    Const S5G221_DnReqHdr_EXT2_CD = 18        '(O)여유필드(Text)
    Const S5G221_DnReqHdr_EXT3_CD = 19        '(O)여유필드(Text)
    Const S5G221_DnReqHdr_TEMP_SO_NO = 20     '(O)S/O 번호 
    Const S5G221_DnReqHdr_CUR = 21            '(O)화폐단위 
    Const S5G221_DnReqHdr_XCHG_RATE = 22      '(0)환율 
    Const S5G221_DnReqHdr_XCHG_RATE_OP = 23   '(O)환율연산자 
    Const S5G221_DnReqHdr_REMARK = 24         '(O)비고 
    Const S5G221_DnReqHdr_ARRIVAL_DT = 25     '(O)실제납품일 
    Const S5G221_DnReqHdr_ARRIVAL_TIME = 26   '(O)납품시간 
    Const S5G221_DnReqHdr_SO_AUTO_FLAG = 27   '(O)납품시간 
    Const S5G221_DnReqHdr_PLANT_CD = 28       '(M)공장 
    Const S5G221_DnReqHdr_INV_MGR = 29        '(O)재고담당 

	Redim iArrDnHdr(S5G221_DnReqHdr_INV_MGR)
	  
	' 납품처 상세정보 
    Dim iArrSTPInfo
    
	Const S5G211_STPInfo_STP_INFO_NO = 0
	Const S5G211_STPInfo_SHIP_TO_PARTY = 1
	Const S5G211_STPInfo_ZIP_CD = 2
	Const S5G211_STPInfo_ADDR1 = 3
	Const S5G211_STPInfo_ADDR2 = 4
	Const S5G211_STPInfo_ADDR3 = 5
	Const S5G211_STPInfo_RECEIVER = 6
	Const S5G211_STPInfo_TEL_NO1 = 7
	Const S5G211_STPInfo_TEL_NO2 = 8

    Redim iArrSTPInfo(S5G211_STPInfo_TEL_NO2)

	' 운송상세정보 
    Dim iArrTransInfo
    
	Const S5G211_TransInfo_TRANS_INFO_NO = 0
	Const S5G211_TransInfo_TRANS_CO = 1
	Const S5G211_TransInfo_DRIVER = 2
	Const S5G211_TransInfo_VEHICLE_NO = 3
	Const S5G211_TransInfo_SENDER = 4

    Redim iArrTransInfo(S5G211_TransInfo_SENDER)

    '-----------------------
    'Data manipulate area
    '-----------------------
	iArrDnHdr(S5G221_DnReqHdr_DN_REQ_NO) = UCase(Trim(Request("txtDnNo")))						'(O)출하번호 
	iArrDnHdr(S5G221_DnReqHdr_EXCEPT_DN_FLAG) = "N"										'(M)예외출고여부 
	iArrDnHdr(S5G221_DnReqHdr_SHIP_TO_PARTY) = UCase(Trim(Request("txtShip_to_party")))	'(M)납품처 
	iArrDnHdr(S5G221_DnReqHdr_SALES_GRP) = UCase(Trim(Request("txtSales_Grp")))			'(M)영업그룹 
	iArrDnHdr(S5G221_DnReqHdr_MOV_TYPE) = UCase(Trim(Request("txtMovType")))				'(M)출하형태 
	iArrDnHdr(S5G221_DnReqHdr_DLVY_DT) = UNIConvDate(Request("txtDlvy_dt"))				'(M)납기일 
	iArrDnHdr(S5G221_DnReqHdr_PROMISE_DT) = UNIConvDate(Request("txtPlanned_gi_dt"))		'(M)출고예정일 
	iArrDnHdr(S5G221_DnReqHdr_TRANS_METH) = UCase(Trim(Request("txtTrans_meth")))			'(O)운송방법 
	iArrDnHdr(S5G221_DnReqHdr_SO_TYPE) = UCase(Trim(Request("txtSo_Type")))					'(M)수주형태 
	iArrDnHdr(S5G221_DnReqHdr_SO_NO) = UCase(Trim(Request("txtSo_no")))					'(O)S/O번호 
	iArrDnHdr(S5G221_DnReqHdr_SHIP_TO_PLCE) = Trim(Request("txtShip_to_place"))			'(O)납품장소 
	iArrDnHdr(S5G221_DnReqHdr_EXT1_QTY) = 0												'(O)여유필드(수량)
	iArrDnHdr(S5G221_DnReqHdr_EXT2_QTY) = 0												'(O)여유필드(수량)
	iArrDnHdr(S5G221_DnReqHdr_EXT3_QTY) = 0       											'(O)여유필드(수량)
	iArrDnHdr(S5G221_DnReqHdr_EXT1_AMT) = 0       											'(O)여유필드(금액)
	iArrDnHdr(S5G221_DnReqHdr_EXT2_AMT) = 0       											'(O)여유필드(금액)
	iArrDnHdr(S5G221_DnReqHdr_EXT3_AMT) = 0       											'(O)여유필드(금액)
	iArrDnHdr(S5G221_DnReqHdr_EXT1_CD) = ""												'(O)여유필드(Text)
	iArrDnHdr(S5G221_DnReqHdr_EXT2_CD) = ""		    									'(O)여유필드(Text)
	iArrDnHdr(S5G221_DnReqHdr_EXT3_CD) = ""	        									'(O)여유필드(Text)
	iArrDnHdr(S5G221_DnReqHdr_TEMP_SO_NO) = UCase(Trim(Request("txtTempSoNo")))   			'(O)S/O 번호 
	'iArrDnHdr(S5G221_DnReqHdr_CUR) = ""            										'(O)화폐단위 
	'iArrDnHdr(S5G221_DnReqHdr_XCHG_RATE) = ""      										'(0)환율 
	'iArrDnHdr(S5G221_DnReqHdr_XCHG_RATE_OP) = ""   										'(O)환율연산자 
	iArrDnHdr(S5G221_DnReqHdr_REMARK) = UCase(Trim(Request("txtRemark")))					'(O)비고 
	If Trim(Request("txtArriv_dt")) <> "" then
		iArrDnHdr(S5G221_DnReqHdr_ARRIVAL_DT) = UNIConvDate(Request("txtArriv_dt"))			'(O)실제납품일 
	End If 
	iArrDnHdr(S5G221_DnReqHdr_ARRIVAL_TIME) = Trim(Request("txtArriv_Tm"))					'(O)납품시간 
	iArrDnHdr(S5G221_DnReqHdr_SO_AUTO_FLAG) = "N"											'(O)납품시간 
	iArrDnHdr(S5G221_DnReqHdr_PLANT_CD) = Trim(Request("txtPlantCd"))						'(M)공장 
	iArrDnHdr(S5G221_DnReqHdr_INV_MGR) = Trim(Request("txtInvMgr"))						'(O)재고담당 
	
	' 납품처 상세정보 생성여부 
    If Request("txtlgBlnChgValue1") = "True" Then
		iStrCrSTPFlag = "C"		' 생성 
	Else
		iStrCrSTPFlag = "N"
	End If
	
	iStrSTPInfoNo = UCase(Trim(Request("txtSTP_Inf_No")))			'납품처상세정보번호 
	IF iStrSTPInfoNo = "" Then
		iArrSTPInfo(S5G211_STPInfo_STP_INFO_NO) = ""	
		iArrSTPInfo(S5G211_STPInfo_ZIP_CD) = UCase(Trim(Request("txtZIP_cd")))
		iArrSTPInfo(S5G211_STPInfo_ADDR1) = Trim(Request("txtADDR1_Dlv"))
		iArrSTPInfo(S5G211_STPInfo_ADDR2) = Trim(Request("txtADDR2_Dlv"))
		iArrSTPInfo(S5G211_STPInfo_ADDR3) = Trim(Request("txtADDR3_Dlv"))
		iArrSTPInfo(S5G211_STPInfo_RECEIVER) = Trim(Request("txtReceiver"))
		iArrSTPInfo(S5G211_STPInfo_TEL_NO1) = UCase(Trim(Request("txtTel_No1")))
		iArrSTPInfo(S5G211_STPInfo_TEL_NO2) = UCase(Trim(Request("txtTel_No2")))
		' 2003.09.20 - By Hwang Seongbae
		If Trim(Join(iArrSTPInfo, "")) <> "" Then
			iArrSTPInfo(S5G211_STPInfo_SHIP_TO_PARTY) = UCase(Trim(Request("txtShip_to_party")))
		End If
	Else
		iArrSTPInfo(S5G211_STPInfo_STP_INFO_NO) = iStrSTPInfoNo
	End If
	
	' 운송상세정보 변경여부 
	If Request("txtlgBlnChgValue2") = "True" Then
		iStrCrTransFlag = "C"
	Else
		iStrCrTransFlag = "N"
	End If
	
	iStrTransInfoNo = UCase(Trim(Request("txtTrnsp_Inf_No")))		'운송정보번호 
	IF iStrTransInfoNo = "" Then
		iArrTransInfo(S5G211_TransInfo_TRANS_INFO_NO) = ""
		iArrTransInfo(S5G211_TransInfo_TRANS_CO) = UCase(Trim(Request("txtTransCo")))
		iArrTransInfo(S5G211_TransInfo_DRIVER) = Trim(Request("txtDriver"))
		iArrTransInfo(S5G211_TransInfo_VEHICLE_NO) = UCase(Trim(Request("txtVehicleNo")))
		iArrTransInfo(S5G211_TransInfo_SENDER) = Trim(Request("txtSender"))
	Else
		iArrTransInfo(S5G211_TransInfo_TRANS_INFO_NO) = iStrTransInfoNo
	End If

	' 수주번호 지정여부 
	If Trim(Request("txtChkSoNo")) = "Y" Then
		iStrFixSoNoFlag = "Y"
	ElseIf Trim(Request("txtChkSoNo")) = "N" Then
		iStrFixSoNoFlag = "N"
	End If
	
	'###################
	'2003.01.02 SMJ	
	pvCB = "F"
	'###################
    Set iObjS5G221 = Server.CreateObject("PS5G221.cSDnReqHdrSvr")
    
    If CheckSYSTEMError(Err,True) = True Then
       Set iObjS5G221 = Nothing		                                                 '☜: Unload Comproxy DLL
       Exit Sub
    End If
    
	'Public Function Maintain(ByVal pvCB As String, _
	'                       ByVal pvStrGlobalCollection As String, _
	'                       ByVal pvStrCUDFlag As String, _
	'                       ByVal pvArrDnHdrIn As Variant, _
	'                       Optional ByVal pvArrDnSalesIn As Variant, _
	'                       Optional ByVal pvStrValidityChkFlag As String, _
	'                       Optional ByVal pvStrFixSoNoFlag As String, _
	'                       Optional ByVal pvStrCrSTPFlag As String, _
	'                       Optional ByVal pvArrSTPInfoIn As Variant, _
	'                       Optional ByVal pvStrCrTransFlag As String, _
	'                       Optional ByVal pvArrTransInfoIn As Variant, _
	'                       Optional ByRef prStrDnNoOut As Variant, _
	'                       Optional ByVal pvICustomXML As String, _
	'                       Optional ByRef prOCustomXML As Variant) As Variant

    Call iObjS5G221.Maintain(pvCB, gStrGlobalCollection, iStrCUDFlag, iArrDnHdr, , "Y",_
							iStrFixSoNoFlag, iStrCrSTPFlag, iArrSTPInfo, iStrCrTransFlag, iArrTransInfo, iStrDnNo)
    
	Set iObjS5G221 = Nothing

	If CheckSYSTEMError(Err,True) = True Then
       Set iObjS5G221 = Nothing		                                                 '☜: Unload Comproxy DLL
       Exit Sub
    End If
    
    Response.Write "<Script Language=vbscript>"							& vbCr
	Response.Write "With parent"										& vbCr
	
	If iStrDnNo <> "" then
		Response.Write ".frm1.txtConDnNo.value = """ & ConvSPChars(iStrDnNo)			& """" & vbcr
	End If
	
	Response.Write ".DbSaveOk"											& vbCr
	Response.Write "End With"											& vbCr
    Response.Write "</Script>"											& vbCr
	Response.End
	
End Sub
'============================================
' Name : SubBizDelete
' Desc : Delete DB data
'============================================
Sub SubBizDelete()
    on Error Resume Next

    Dim iStrCUDFlag, pvCB
    Dim iObjS5G221
    Dim iArrDnHdr
    
    ReDim iArrDnHdr(1)
    
    If Request("txtConDnNo") = "" Then										'⊙: 조회를 위한 값이 들어왔는지 체크 
		Call ServerMesgBox("삭제 조건값이 비어있습니다!", vbInformation, I_MKSCRIPT)
		Response.End 
	End If
	
	iStrCUDFlag = "D"
	pvCB = "F"
	
	iArrDnHdr(0) = Trim(Request("txtConDnNo"))
	iArrDnHdr(1) = "N"								' 예외출고여부 

    Set iObjS5G221 = Server.CreateObject("PS5G221.cSDnReqHdrSvr")
    
    If CheckSYSTEMError(Err,True) = True Then
       Set iObjS5G221 = Nothing		                                                 '☜: Unload Comproxy DLL
       Exit Sub
    End If  
    
    Call iObjS5G221.Maintain(pvCB, gStrGlobalCollection, iStrCUDFlag, iArrDnHdr)
	
	Set iObjS5G221 = Nothing
							 
	If CheckSYSTEMError(Err,True) = True Then
       Set iObjS5G221 = Nothing		                                                 '☜: Unload Comproxy DLL
       Exit Sub
    End If 
    
	'-----------------------
	'Result data display area
	'----------------------- 
	Response.Write "<Script Language=vbscript>"							& vbCr
	Response.Write "Call parent.DbDeleteOk()"							& vbCr
	Response.Write "</Script>"											& vbCr
	Response.End		 
	
End Sub
%>
