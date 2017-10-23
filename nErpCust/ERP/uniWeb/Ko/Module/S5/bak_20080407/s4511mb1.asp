<%@ LANGUAGE=VBSCript%>
<%Option Explicit    %>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../ComASP/LoadInfTb19029.asp" -->
<%
'**********************************************************************************************
'*  1. Module Name          : ���� 
'*  2. Function Name        : ���ϰ��� 
'*  3. Program ID           : S4210MA1
'*  4. Program Name         : ���Ͽ�û��� 
'*  5. Program Desc         :
'*  6. Comproxy List        : S42101MaintDnHdrSvr
'*  7. Modified date(First) : 2000/04/17
'*  8. Modified date(Last)  : 2002/06/19
'*  9. Modifier (First)     : Cho song hyon
'* 10. Modifier (Last)      : Lee Myung Wha
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              :
'*                            -2000/04/17 : ȭ�� Layout & ASP Coding
'*                            -2001/12/19 : Date ǥ������ 
'**********************************************************************************************

	Call LoadBasisGlobalInf()
	Call LoadInfTB19029B( "I", "*", "NOCOOKIE", "MB")
	Call LoadBNumericFormatB("I", "*", "NOCOOKIE", "MB")

	Dim lgOpModeCRUD
    
    on Error Resume Next                                                             
    Err.Clear                                                                        

    Call HideStatusWnd                                                               '��: Hide Processing message
    
    lgOpModeCRUD      = Request("txtMode")                                           '��: Read Operation Mode (CRUD)
    
    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '��: Query
	         Call SubBizQuery()
        Case CStr(UID_M0002)
	         Call SubBizSave()
        Case CStr(UID_M0003)                                                         '��: Delete
             Call SubBizDelete()
    End Select

'============================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================
Sub SubBizQuery()

	on Error Resume Next
    
    '==================== ��ȸ��� Index ==================
	Const S5G211_RS_DN_NO = 0                ' ���Ϲ�ȣ 
	Const S5G211_RS_SHIP_TO_PARTY = 1        ' ��ǰó 
	Const S5G211_RS_SHIP_TO_PARTY_NM = 2     ' ��ǰó�� 
	Const S5G211_RS_SALES_GRP = 3            ' �����׷� 
	Const S5G211_RS_SALES_GRP_NM = 4         ' �����׷�� 
	Const S5G211_RS_SALES_ORG = 5            ' �������� 
	Const S5G211_RS_SALES_ORG_NM = 6         ' ���������� 
	Const S5G211_RS_MOV_TYPE = 7             ' �������� 
	Const S5G211_RS_MOV_TYPE_NM = 8          ' �������¸� 
	Const S5G211_RS_DLVY_DT = 9             ' ������ 
	Const S5G211_RS_PROMISE_DT = 10          ' ������� 
	Const S5G211_RS_ACTUAL_GI_DT = 11        ' ����� 
	Const S5G211_RS_COST_CD = 12             ' Cost Center
	Const S5G211_RS_BIZ_AREA = 13            ' ����� 
	Const S5G211_RS_BIZ_AREA_NM = 14         ' ������ 
	Const S5G211_RS_TRANS_METH = 15          ' ��۹�� 
	Const S5G211_RS_TRANS_METH_NM = 16       ' ��۹���� 
	Const S5G211_RS_GOODS_MV_NO = 17         ' ���ҹ�ȣ(����ȣ)
	Const S5G211_RS_CI_FLAG = 18             ' ����ʿ俩�� 
	Const S5G211_RS_POST_FLAG = 19           ' ���ó������ 
	Const S5G211_RS_SO_TYPE = 20             ' �������� 
	Const S5G211_RS_SO_TYPE_NM = 21          ' �������¸� 
	Const S5G211_RS_SO_NO = 22               ' ���ֹ�ȣ 
	Const S5G211_RS_SHIP_TO_PLCE = 23        ' ��ǰ��� 
	Const S5G211_RS_INSRT_USER_ID = 24       ' ����� 
	Const S5G211_RS_INSRT_DT = 25            ' ����� 
	Const S5G211_RS_UPDT_USER_ID = 26        ' ������ 
	Const S5G211_RS_UPDT_DT = 27             ' ������ 
	Const S5G211_RS_EXT1_QTY = 28            ' �����ʵ�(����)
	Const S5G211_RS_EXT2_QTY = 29            ' �����ʵ�(����)
	Const S5G211_RS_EXT3_QTY = 30            ' �����ʵ�(����)
	Const S5G211_RS_EXT1_AMT = 31            ' �����ʵ�(�ݾ�)
	Const S5G211_RS_EXT2_AMT = 32            ' �����ʵ�(�ݾ�)
	Const S5G211_RS_EXT3_AMT = 33            ' �����ʵ�(�ݾ�)
	Const S5G211_RS_EXT1_CD = 34             ' �����ʵ�(Text)
	Const S5G211_RS_EXT2_CD = 35             ' �����ʵ�(Text)
	Const S5G211_RS_EXT3_CD = 36             ' �����ʵ�(Text)
	Const S5G211_RS_TEMP_SO_NO = 37          ' ���ֹ�ȣ 
	Const S5G211_RS_VAT_FLAG = 38            ' ���ݰ�꼭���� ���û������� 
	Const S5G211_RS_AR_FLAG = 39             ' �������� ���û������� 
	Const S5G211_RS_CUR = 40                 ' ȭ����� 
	Const S5G211_RS_XCHG_RATE = 41           ' ȯ�� 
	Const S5G211_RS_XCHG_RATE_OP = 42        ' ȯ�������� 
	Const S5G211_RS_NET_AMT = 43             ' ���ݾ� 
	Const S5G211_RS_NET_AMT_LOC = 44         ' ����ڱ��ݾ� 
	Const S5G211_RS_VAT_AMT = 45             ' VAT �ݾ� 
	Const S5G211_RS_VAT_AMT_LOC = 46         ' VAT �ڱ��ݾ� 
	Const S5G211_RS_EXCEPT_DN_FLAG = 47      ' ��������� 
	Const S5G211_RS_REMARK = 48              ' ��� 
	Const S5G211_RS_ARRIVAL_DT = 49          ' ������ǰ�� 
	Const S5G211_RS_ARRIVAL_TIME = 50        ' ��ǰ�ð� 
	Const S5G211_RS_STP_INFO_NO = 51         ' ��ǰó������ ��ȣ 
	Const S5G211_RS_ZIP_CD = 52              ' ��ǰó �����ȣ 
	Const S5G211_RS_ADDR1 = 53               ' ��ǰó �ּ� 
	Const S5G211_RS_ADDR2 = 54               ' ��ǰó �ּ� 
	Const S5G211_RS_ADDR3 = 55               ' ��ǰó �ּ� 
	Const S5G211_RS_RECEIVER = 56            ' �μ��ڸ� 
	Const S5G211_RS_TEL_NO1 = 57             ' ��ȭ��ȣ1
	Const S5G211_RS_TEL_NO2 = 58             ' ��ȭ��ȣ2
	Const S5G211_RS_TRANS_INFO_NO = 59       ' ���������ȣ 
	Const S5G211_RS_TRANS_CO = 60            ' ���ȸ�� 
	Const S5G211_RS_DRIVER = 61              ' �����ڸ� 
	Const S5G211_RS_VEHICLE_NO = 62          ' ������ȣ 
	Const S5G211_RS_SENDER = 63              ' �ΰ��ڸ� 
	Const S5G211_RS_STO_FLAG = 64            ' STO Flag
	Const S5G211_RS_CASH_DC_AMT = 65         ' �������ξ� 
	Const S5G211_RS_TAX_DC_AMT = 66          ' �������ξ� 
	Const S5G211_RS_TAX_BASE_AMT = 67        ' ���ݰ����ʱݾ� 
	Const S5G211_RS_CASH_DC_AMT_LOC = 68     ' �������ξ�(�ڱ�)
	Const S5G211_RS_TAX_DC_AMT_LOC = 69      ' �������ξ�(�ڱ�)
	Const S5G211_RS_TAX_BASE_AMT_LOC = 70    ' ���ݰ����ʱݾ�(�ڱ�)
	Const S5G211_RS_SO_AUTO_FLAG = 71        ' ���ַκ��� �ڵ������� ���� 
	Const S5G211_RS_PLANT_CD = 72            ' ���� 
	Const S5G211_RS_PLANT_NM = 73            ' ����� 
	Const S5G211_RS_INV_MGR = 74             ' ����� 
	Const S5G211_RS_INV_MGR_NM = 75          ' ������ڸ� 
    Const S5G211_RS_CONTRY_CD = 76			 ' ��ǰó �����ڵ� 
	Const S5G211_RS_SCM_DO_NO = 80
	Const S5G211_RS_SCM_DO_SEQ = 81
    
    Dim iStrDnNo
	Dim iObjS5G2211
	Dim iArrSDnReqHdr
	
	iStrDnNo = Trim(Request("txtConDnNo"))			' ���Ϲ�ȣ 
	
    Set iObjS5G2211 = Server.CreateObject("PS5G221.cLookUpSDnReqHdr")
	
	If CheckSYSTEMError(Err,True) = True Then
       Response.Write "<Script Language=vbscript>"			& vbCr
	   Response.Write "parent.frm1.txtConDnNo.focus"		& vbCr    
	   Response.Write "</Script>"							& vbCr	
       Exit Sub
    End If  
    
    iArrSDnReqHdr = iObjS5G2211.LookUp(gStrGlobalCollection, iStrDnNo, "N")
    
    If CheckSYSTEMError(Err,True) = True Then
       Set iObjS5G2211 = Nothing		                                                 '��: Unload Comproxy DLL
       Response.Write "<Script Language=vbscript>"			& vbCr
	   Response.Write "parent.frm1.txtConDnNo.focus"		& vbCr    
	   Response.Write "</Script>"							& vbCr	
       Exit Sub
    End If  
    
    Set iObjS5G2211 = Nothing
	
	Response.Write "<Script Language=vbscript>" & vbCr
	Response.Write "With parent.frm1" & vbCr
	
	'--���Ϲ�ȣ--
	Response.Write ".txtDnNo.value				= """ & ConvSPChars(iArrSDnReqHdr(S5G211_RS_DN_NO))				& """" & vbcr
	'--�������--
	Response.Write ".txtPlanned_gi_dt.Text		= """ & UNIDateClientFormat(iArrSDnReqHdr(S5G211_RS_PROMISE_DT))	& """" & vbcr
	'--���������--
	Response.Write ".txtActGi_dt.value			= """ & UNIDateClientFormat(iArrSDnReqHdr(S5G211_RS_ACTUAL_GI_DT))	& """" & vbcr
	'--��ǰ���--
	Response.Write ".txtShip_to_place.value		= """ & ConvSPChars(iArrSDnReqHdr(S5G211_RS_SHIP_TO_PLCE))			& """" & vbcr
	'--��۹��--
	Response.Write ".txtTrans_meth.value		= """ & ConvSPChars(iArrSDnReqHdr(S5G211_RS_TRANS_METH))			& """" & vbcr
	'--��۹����--
	Response.Write ".txtTrans_meth_nm.value		= """ & ConvSPChars(iArrSDnReqHdr(S5G211_RS_TRANS_METH_NM))		& """" & vbcr
	'--������--
	Response.Write ".txtDlvy_dt.value			= """ & UNIDateClientFormat(iArrSDnReqHdr(S5G211_RS_DLVY_DT))		& """" & vbcr
	'--�����׷�--
	Response.Write ".txtSales_Grp.value			= """ & ConvSPChars(iArrSDnReqHdr(S5G211_RS_SALES_GRP))			& """" & vbcr
	'--�����׷��--
	Response.Write ".txtSales_GrpNm.value		= """ & ConvSPChars(iArrSDnReqHdr(S5G211_RS_SALES_GRP_NM))			& """" & vbcr
	'--����Ÿ��--
	Response.Write ".txtMovType.value			= """ & ConvSPChars(iArrSDnReqHdr(S5G211_RS_MOV_TYPE))				& """" & vbcr
	'--����Ÿ�Ը�--
	Response.Write ".txtMovTypeNm.value			= """ & ConvSPChars(iArrSDnReqHdr(S5G211_RS_MOV_TYPE_NM))			& """" & vbcr
	'--����ȣ--
	Response.Write ".txtGoods_mv_no.value		= """ & Trim(ConvSPChars(iArrSDnReqHdr(S5G211_RS_GOODS_MV_NO)))	& """" & vbcr
	'--���ֹ�ȣ--
	Response.Write ".txtSo_no.value				= """ & ConvSPChars(iArrSDnReqHdr(S5G211_RS_SO_NO))				& """" & vbcr
	'--���ֹ�ȣ����--
	If iArrSDnReqHdr(S5G211_RS_SO_NO) <> "" Then
		Response.Write ".chkSoNo.checked = True" & vbcr
	End If
	'--����Ÿ��--
	Response.Write ".txtSo_Type.value			= """ & ConvSPChars(iArrSDnReqHdr(S5G211_RS_SO_TYPE))				& """" & vbcr
	'--����Ÿ�Ը�--
	Response.Write ".txtSo_TypeNm.value			= """ & ConvSPChars(iArrSDnReqHdr(S5G211_RS_SO_TYPE_NM))			& """" & vbcr
	'--��ǰó--
	Response.Write ".txtShip_to_party.value		= """ & ConvSPChars(iArrSDnReqHdr(S5G211_RS_SHIP_TO_PARTY))		& """" & vbcr
	'--��ǰó��--
	Response.Write ".txtShip_to_partyNm.value	= """ & ConvSPChars(iArrSDnReqHdr(S5G211_RS_SHIP_TO_PARTY_NM))		& """" & vbcr
	'--�����ڵ�--
	Response.Write ".txtHCntryCd.value	= """ & ConvSPChars(iArrSDnReqHdr(S5G211_RS_CONTRY_CD))			& """" & vbcr
	'--���--
	Response.Write ".txtRemark.value	= """ & ConvSPChars(iArrSDnReqHdr(S5G211_RS_REMARK))				& """" & vbcr

	'--���--
	Response.Write ".txtSCMDONo.value	= """ & ConvSPChars(iArrSDnReqHdr(S5G211_RS_SCM_DO_NO))				& """" & vbcr
	'--���--
	Response.Write ".txtSCMDONoSeq.value	= """ & ConvSPChars(iArrSDnReqHdr(S5G211_RS_SCM_DO_SEQ))				& """" & vbcr


	'--������ǰ��--
	Response.Write ".txtArriv_dt.Text	= """ & UNIDateClientFormat(iArrSDnReqHdr(S5G211_RS_ARRIVAL_DT))	& """" & vbcr
	'--������ǰ�ð�--
	Response.Write ".txtArriv_tm.value	= """ & ConvSPChars(iArrSDnReqHdr(S5G211_RS_ARRIVAL_TIME))			& """" & vbcr
	'--��ǰó��������ȣ--
	Response.Write ".txtSTP_Inf_No.value	= """ & ConvSPChars(iArrSDnReqHdr(S5G211_RS_STP_INFO_NO))		& """" & vbcr
	'--���������ȣ--
	Response.Write ".txtTrnsp_Inf_No.value	= """ & ConvSPChars(iArrSDnReqHdr(S5G211_RS_TRANS_INFO_NO))	& """" & vbcr
	'--�����ȣ--
	Response.Write ".txtZIP_cd.value	= """ & ConvSPChars(iArrSDnReqHdr(S5G211_RS_ZIP_CD))			& """" & vbcr
	'--�ּ�1--
	Response.Write ".txtADDR1_Dlv.value	= """ & ConvSPChars(iArrSDnReqHdr(S5G211_RS_ADDR1))			& """" & vbcr
	'--�ּ�2--
	Response.Write ".txtADDR2_Dlv.value	= """ & ConvSPChars(iArrSDnReqHdr(S5G211_RS_ADDR2))			& """" & vbcr
	'--�ּ�3--
	Response.Write ".txtADDR3_Dlv.value	= """ & ConvSPChars(iArrSDnReqHdr(S5G211_RS_ADDR3))			& """" & vbcr
	'--�μ��ڸ�--
	Response.Write ".txtReceiver.value	= """ & ConvSPChars(iArrSDnReqHdr(S5G211_RS_RECEIVER))			& """" & vbcr
	'--��ȭ��ȣ1--
	Response.Write ".txtTel_No1.value	= """ & ConvSPChars(iArrSDnReqHdr(S5G211_RS_TEL_NO1))			& """" & vbcr
	'--��ȭ��ȣ2--
	Response.Write ".txtTel_No2.value	= """ & ConvSPChars(iArrSDnReqHdr(S5G211_RS_TEL_NO2))			& """" & vbcr
	'--���ȸ��--
	Response.Write ".txtTransCo.value	= """ & ConvSPChars(iArrSDnReqHdr(S5G211_RS_TRANS_CO))			& """" & vbcr
	'--������--
	Response.Write ".txtDriver.value	= """ & ConvSPChars(iArrSDnReqHdr(S5G211_RS_DRIVER))			& """" & vbcr
	'--������ȣ--
	Response.Write ".txtVehicleNo.value	= """ & ConvSPChars(iArrSDnReqHdr(S5G211_RS_VEHICLE_NO))		& """" & vbcr
	'--�ΰ���--
	Response.Write ".txtSender.value	= """ & ConvSPChars(iArrSDnReqHdr(S5G211_RS_SENDER))			& """" & vbcr
	'--����--
	Response.Write ".txtPlantCd.value	= """ & ConvSPChars(iArrSDnReqHdr(S5G211_RS_PLANT_CD))			& """" & vbcr
	Response.Write ".txtPlantNm.value	= """ & ConvSPChars(iArrSDnReqHdr(S5G211_RS_PLANT_NM))			& """" & vbcr
	'--�������--
	Response.Write ".txtInvMgr.value	= """ & ConvSPChars(iArrSDnReqHdr(S5G211_RS_INV_MGR))			& """" & vbcr
	Response.Write ".txtInvMgrNm.value	= """ & ConvSPChars(iArrSDnReqHdr(S5G211_RS_INV_MGR_NM))		& """" & vbcr

	Response.Write "parent.DbQueryOk" & vbCr
	Response.Write "End With"          & vbCr
    Response.Write "</Script>"         & vbCr															'��: ��ȸ�� ���� 
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
    
    lgIntFlgMode = CInt(Request("txtFlgMode"))									'��: ����� Create/Update �Ǻ� 

    If lgIntFlgMode = OPMD_CMODE Then
		iStrCUDFlag = "C"
    ElseIf lgIntFlgMode = OPMD_UMODE Then
		iStrCUDFlag = "U"
	Else
		Call ServerMesgBox("TXTFLGMODE ���ǰ��� ����ֽ��ϴ�!", vbInformation, I_MKSCRIPT)              
		Response.End 
    End If
    
    ' �������� 
    Dim iArrDnHdr

    Const S5G221_DnReqHdr_DN_REQ_NO = 0           '(O)���Ϲ�ȣ 
	Const S5G221_DnReqHdr_EXCEPT_DN_FLAG = 1  '(M)���������(Y/N)	
    Const S5G221_DnReqHdr_SHIP_TO_PARTY = 2   '(M)��ǰó 
    Const S5G221_DnReqHdr_SALES_GRP = 3       '(M)�����׷� 
    Const S5G221_DnReqHdr_MOV_TYPE = 4        '(M)�������� 
    Const S5G221_DnReqHdr_DLVY_DT = 5         '(M)������ 
    Const S5G221_DnReqHdr_PROMISE_DT = 6      '(M)������� 
    Const S5G221_DnReqHdr_TRANS_METH = 7      '(O)��۹�� 
    Const S5G221_DnReqHdr_SO_TYPE = 8         '(M)�������� 
    Const S5G221_DnReqHdr_SO_NO = 9           '(O)S/O��ȣ 
    Const S5G221_DnReqHdr_SHIP_TO_PLCE = 10    '(O)��ǰ��� 
    Const S5G221_DnReqHdr_EXT1_QTY = 11       '(O)�����ʵ�(����)
    Const S5G221_DnReqHdr_EXT2_QTY = 12       '(O)�����ʵ�(����)
    Const S5G221_DnReqHdr_EXT3_QTY = 13       '(O)�����ʵ�(����)
    Const S5G221_DnReqHdr_EXT1_AMT = 14       '(O)�����ʵ�(�ݾ�)
    Const S5G221_DnReqHdr_EXT2_AMT = 15       '(O)�����ʵ�(�ݾ�)
    Const S5G221_DnReqHdr_EXT3_AMT = 16       '(O)�����ʵ�(�ݾ�)
    Const S5G221_DnReqHdr_EXT1_CD = 17        '(O)�����ʵ�(Text)
    Const S5G221_DnReqHdr_EXT2_CD = 18        '(O)�����ʵ�(Text)
    Const S5G221_DnReqHdr_EXT3_CD = 19        '(O)�����ʵ�(Text)
    Const S5G221_DnReqHdr_TEMP_SO_NO = 20     '(O)S/O ��ȣ 
    Const S5G221_DnReqHdr_CUR = 21            '(O)ȭ����� 
    Const S5G221_DnReqHdr_XCHG_RATE = 22      '(0)ȯ�� 
    Const S5G221_DnReqHdr_XCHG_RATE_OP = 23   '(O)ȯ�������� 
    Const S5G221_DnReqHdr_REMARK = 24         '(O)��� 
    Const S5G221_DnReqHdr_ARRIVAL_DT = 25     '(O)������ǰ�� 
    Const S5G221_DnReqHdr_ARRIVAL_TIME = 26   '(O)��ǰ�ð� 
    Const S5G221_DnReqHdr_SO_AUTO_FLAG = 27   '(O)��ǰ�ð� 
    Const S5G221_DnReqHdr_PLANT_CD = 28       '(M)���� 
    Const S5G221_DnReqHdr_INV_MGR = 29        '(O)����� 

	Redim iArrDnHdr(S5G221_DnReqHdr_INV_MGR)
	  
	' ��ǰó ������ 
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

	' ��ۻ����� 
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
	iArrDnHdr(S5G221_DnReqHdr_DN_REQ_NO) = UCase(Trim(Request("txtDnNo")))						'(O)���Ϲ�ȣ 
	iArrDnHdr(S5G221_DnReqHdr_EXCEPT_DN_FLAG) = "N"										'(M)��������� 
	iArrDnHdr(S5G221_DnReqHdr_SHIP_TO_PARTY) = UCase(Trim(Request("txtShip_to_party")))	'(M)��ǰó 
	iArrDnHdr(S5G221_DnReqHdr_SALES_GRP) = UCase(Trim(Request("txtSales_Grp")))			'(M)�����׷� 
	iArrDnHdr(S5G221_DnReqHdr_MOV_TYPE) = UCase(Trim(Request("txtMovType")))				'(M)�������� 
	iArrDnHdr(S5G221_DnReqHdr_DLVY_DT) = UNIConvDate(Request("txtDlvy_dt"))				'(M)������ 
	iArrDnHdr(S5G221_DnReqHdr_PROMISE_DT) = UNIConvDate(Request("txtPlanned_gi_dt"))		'(M)������� 
	iArrDnHdr(S5G221_DnReqHdr_TRANS_METH) = UCase(Trim(Request("txtTrans_meth")))			'(O)��۹�� 
	iArrDnHdr(S5G221_DnReqHdr_SO_TYPE) = UCase(Trim(Request("txtSo_Type")))					'(M)�������� 
	iArrDnHdr(S5G221_DnReqHdr_SO_NO) = UCase(Trim(Request("txtSo_no")))					'(O)S/O��ȣ 
	iArrDnHdr(S5G221_DnReqHdr_SHIP_TO_PLCE) = Trim(Request("txtShip_to_place"))			'(O)��ǰ��� 
	iArrDnHdr(S5G221_DnReqHdr_EXT1_QTY) = 0												'(O)�����ʵ�(����)
	iArrDnHdr(S5G221_DnReqHdr_EXT2_QTY) = 0												'(O)�����ʵ�(����)
	iArrDnHdr(S5G221_DnReqHdr_EXT3_QTY) = 0       											'(O)�����ʵ�(����)
	iArrDnHdr(S5G221_DnReqHdr_EXT1_AMT) = 0       											'(O)�����ʵ�(�ݾ�)
	iArrDnHdr(S5G221_DnReqHdr_EXT2_AMT) = 0       											'(O)�����ʵ�(�ݾ�)
	iArrDnHdr(S5G221_DnReqHdr_EXT3_AMT) = 0       											'(O)�����ʵ�(�ݾ�)
	iArrDnHdr(S5G221_DnReqHdr_EXT1_CD) = ""												'(O)�����ʵ�(Text)
	iArrDnHdr(S5G221_DnReqHdr_EXT2_CD) = ""		    									'(O)�����ʵ�(Text)
	iArrDnHdr(S5G221_DnReqHdr_EXT3_CD) = ""	        									'(O)�����ʵ�(Text)
	iArrDnHdr(S5G221_DnReqHdr_TEMP_SO_NO) = UCase(Trim(Request("txtTempSoNo")))   			'(O)S/O ��ȣ 
	'iArrDnHdr(S5G221_DnReqHdr_CUR) = ""            										'(O)ȭ����� 
	'iArrDnHdr(S5G221_DnReqHdr_XCHG_RATE) = ""      										'(0)ȯ�� 
	'iArrDnHdr(S5G221_DnReqHdr_XCHG_RATE_OP) = ""   										'(O)ȯ�������� 
	iArrDnHdr(S5G221_DnReqHdr_REMARK) = UCase(Trim(Request("txtRemark")))					'(O)��� 
	If Trim(Request("txtArriv_dt")) <> "" then
		iArrDnHdr(S5G221_DnReqHdr_ARRIVAL_DT) = UNIConvDate(Request("txtArriv_dt"))			'(O)������ǰ�� 
	End If 
	iArrDnHdr(S5G221_DnReqHdr_ARRIVAL_TIME) = Trim(Request("txtArriv_Tm"))					'(O)��ǰ�ð� 
	iArrDnHdr(S5G221_DnReqHdr_SO_AUTO_FLAG) = "N"											'(O)��ǰ�ð� 
	iArrDnHdr(S5G221_DnReqHdr_PLANT_CD) = Trim(Request("txtPlantCd"))						'(M)���� 
	iArrDnHdr(S5G221_DnReqHdr_INV_MGR) = Trim(Request("txtInvMgr"))						'(O)����� 
	
	' ��ǰó ������ �������� 
    If Request("txtlgBlnChgValue1") = "True" Then
		iStrCrSTPFlag = "C"		' ���� 
	Else
		iStrCrSTPFlag = "N"
	End If
	
	iStrSTPInfoNo = UCase(Trim(Request("txtSTP_Inf_No")))			'��ǰó��������ȣ 
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
	
	' ��ۻ����� ���濩�� 
	If Request("txtlgBlnChgValue2") = "True" Then
		iStrCrTransFlag = "C"
	Else
		iStrCrTransFlag = "N"
	End If
	
	iStrTransInfoNo = UCase(Trim(Request("txtTrnsp_Inf_No")))		'���������ȣ 
	IF iStrTransInfoNo = "" Then
		iArrTransInfo(S5G211_TransInfo_TRANS_INFO_NO) = ""
		iArrTransInfo(S5G211_TransInfo_TRANS_CO) = UCase(Trim(Request("txtTransCo")))
		iArrTransInfo(S5G211_TransInfo_DRIVER) = Trim(Request("txtDriver"))
		iArrTransInfo(S5G211_TransInfo_VEHICLE_NO) = UCase(Trim(Request("txtVehicleNo")))
		iArrTransInfo(S5G211_TransInfo_SENDER) = Trim(Request("txtSender"))
	Else
		iArrTransInfo(S5G211_TransInfo_TRANS_INFO_NO) = iStrTransInfoNo
	End If

	' ���ֹ�ȣ �������� 
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
       Set iObjS5G221 = Nothing		                                                 '��: Unload Comproxy DLL
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
       Set iObjS5G221 = Nothing		                                                 '��: Unload Comproxy DLL
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
    
    If Request("txtConDnNo") = "" Then										'��: ��ȸ�� ���� ���� ���Դ��� üũ 
		Call ServerMesgBox("���� ���ǰ��� ����ֽ��ϴ�!", vbInformation, I_MKSCRIPT)
		Response.End 
	End If
	
	iStrCUDFlag = "D"
	pvCB = "F"
	
	iArrDnHdr(0) = Trim(Request("txtConDnNo"))
	iArrDnHdr(1) = "N"								' ��������� 

    Set iObjS5G221 = Server.CreateObject("PS5G221.cSDnReqHdrSvr")
    
    If CheckSYSTEMError(Err,True) = True Then
       Set iObjS5G221 = Nothing		                                                 '��: Unload Comproxy DLL
       Exit Sub
    End If  
    
    Call iObjS5G221.Maintain(pvCB, gStrGlobalCollection, iStrCUDFlag, iArrDnHdr)
	
	Set iObjS5G221 = Nothing
							 
	If CheckSYSTEMError(Err,True) = True Then
       Set iObjS5G221 = Nothing		                                                 '��: Unload Comproxy DLL
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
