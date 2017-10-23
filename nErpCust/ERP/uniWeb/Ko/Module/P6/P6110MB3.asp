<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<%'======================================================================================================
'*  1. Module Name          : 금형관리 
'*  2. Function Name        : 
'*  3. Program ID           : P6110Mb3.asp
'*  4. Program Name         : 금형점검계획등록 
'*  5. Program Desc         : 
'*  6. Modified date(First) : 2005-01-25
'*  7. Modified date(Last)  :
'*  8. Modifier (First)     : Lee Sang Ho
'*  9. Modifier (Last)      : 
'* 10. Comment              : 
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'=======================================================================================================
%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgent.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgentVariables.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%													'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 
Call LoadBasisGlobalInf
Call loadInfTB19029B("I", "*", "NOCOOKIE","MB")

Dim ADF										'ActiveX Data Factory 지정 변수선언 
Dim strRetMsg								'Record Set Return Message 변수선언 
Dim UNISqlId, UNIValue, UNILock, UNIFlag	'DBAgent Parameter 선언 
Dim rs0, rs1, rs2, rs3, rs4, rs5			'DBAgent Parameter 선언 
Dim strQryMode

Dim i
Dim iStr

Call HideStatusWnd

strQryMode = Request("lgIntFlgMode")

On Error Resume Next

Err.Clear																	'☜: Protect system from crashing

'=======================================================================================================
'	만약, 선언한 변수가 배열이라면 아래와같은 Fix 된 배열로 Redim 을 해서 넘겨줘야 한다.
'=======================================================================================================
Redim UNISqlId(0)
Redim UNIValue(0, 0)


UNISqlId(0) = "Y6110MB102"
UNIValue(0, 0) = FilterVar(Ucase(request("txtCastCd")),"''","S")

UNILock = DISCONNREAD :	UNIFlag = "1"

Set ADF = Server.CreateObject("prjPublic.cCtlTake")

strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs1)

Set ADF = Nothing
	
iStr = Split(strRetMsg,gColSep)

If iStr(0) <> "0" Then
	Call ServerMesgBox(strRetMsg , vbInformation, I_MKSCRIPT)
End If

	Response.Write "<Script Language=vbscript>" & vbCr
	Response.Write "parent.frm1.txtCastCd1.Value		= """ & ConvSPChars(rs1("CAST_CD")) & """" & vbCr                                       
	Response.Write "parent.frm1.txtCastNm1.Value		= """ & ConvSPChars(rs1("CAST_NM")) & """" & vbCr                               
	Response.Write "parent.frm1.txtCarKind1.Value		= """ & ConvSPChars(rs1("CAR_KIND")) & """" & vbCr                                
	Response.Write "parent.frm1.txtCarKindNm1.Value		= """ & ConvSPChars(rs1("CAR_KIND_NM")) & """" & vbCr                                
	'Response.Write "parent.frm1.txtMfgCd.Value			= """ & ConvSPChars(rs1("MFG_CD")) & """" & vbCr                               
	Response.Write "parent.frm1.txtAsstCd1.Value		= """ & ConvSPChars(rs1("ASST_CD1")) & """" & vbCr                               
	Response.Write "parent.frm1.txtAsstNm1.Value		= """ & ConvSPChars(rs1("ASST_NM1")) & """" & vbCr                                
	Response.Write "parent.frm1.txtAsstCd2.Value		= """ & ConvSPChars(rs1("ASST_CD2")) & """" & vbCr                                    
	Response.Write "parent.frm1.txtAsstNm2.value		= """ & ConvSPChars(rs1("ASST_NM2")) & """" & vbCr   
	'Response.Write "parent.frm1.txtCustomYn.Value		= """ & ConvSPChars(rs1("CUSTOM_YN"))  & """" & vbCr            
	Response.Write "parent.frm1.txtMaker.Value			= """ & ConvSPChars(rs1("MAKER"))  & """" & vbCr              
	Response.Write "parent.frm1.txtMakeDt.Text			= """ & UNIDateClientFormat(rs1("MAKE_DT")) & """" & vbCr
	Response.Write "parent.frm1.txtStrType.Value		= """ & ConvSPChars(rs1("STR_TYPE")) & """" & vbCr                                   
	Response.Write "parent.frm1.txtMatQ.Value			= """ & ConvSPChars(rs1("MAT_Q"))  & """" & vbCr                                   
	Response.Write "parent.frm1.txtProcessType.Value	= """ & ConvSPChars(rs1("PROCESS_TYPE")) & """" & vbCr                                   
	Response.Write "parent.frm1.txtSpec.Value			= """ & ConvSPChars(rs1("SPEC"))  & """" & vbCr                                  
	Response.Write "parent.frm1.txtWeightT.Value		= """ & UNINumClientFormat(rs1("WEIGHT_T"), ggQty.DecPoint, 0)  & """" & vbCr                                  
	Response.Write "parent.frm1.txtSHeight.Value		= """ & UNINumClientFormat(rs1("S_HEIGHT"), ggQty.DecPoint, 0) & """" & vbCr
	Response.Write "parent.frm1.txtDHeight.Value		= """ & UNINumClientFormat(rs1("D_HEIGHT"), ggQty.DecPoint, 0) & """" & vbCr
	Response.Write "parent.frm1.txtFormingP.Value		= """ & ConvSPChars(rs1("FORMING_P")) & """" & vbCr
	Response.Write "parent.frm1.txtCushionPr.Value		= """ & UNINumClientFormat(rs1("CUSHION_PR"), ggQty.DecPoint, 0) & """" & vbCr
	Response.Write "parent.frm1.txtCStroke.Value		= """ & UNINumClientFormat(rs1("C_STROKE"), ggQty.DecPoint, 0) & """" & vbCr
	Response.Write "parent.frm1.txtPurAmt.Value			= """ & UNINumClientFormat(rs1("PUR_AMT"), ggAmtOfMoney.DecPoint, 0) & """" & vbCr
	Response.Write "parent.frm1.txtLifeCycle.Value		= """ & UNINumClientFormat(rs1("LIFE_CYCLE"), ggQty.DecPoint, 0) & """" & vbCr
	Response.Write "parent.frm1.txtCloseDt.Text			= """ & UNIDateClientFormat(rs1("CLOSE_DT")) & """" & vbCr
	Response.Write "parent.frm1.txtUseMachine.Value		= """ & ConvSPChars(rs1("USE_MACHINE"))   & """" & vbCr
	Response.Write "parent.frm1.txtAutoMath.Value		= """ & ConvSPChars(rs1("AUTO_MATH")) & """" & vbCr
	Response.Write "parent.frm1.txtPersonCount.Value	= """ & UNINumClientFormat(rs1("PERSON_COUNT"), ggQty.DecPoint, 0) & """" & vbCr
	Response.Write "parent.frm1.txtModifyDire.Value		= """ & ConvSPChars(rs1("MODIFY_DIRE")) & """" & vbCr
	Response.Write "parent.frm1.txtGuideMath.Value		= """ & ConvSPChars(rs1("GUIDE_MATH")) & """" & vbCr
	Response.Write "parent.frm1.txtLocate.Value			= """ & ConvSPChars(rs1("LOCATE")) & """" & vbCr
	Response.Write "parent.frm1.txtLoading.Value		= """ & ConvSPChars(rs1("LOADING")) & """" & vbCr
	Response.Write "parent.frm1.txtUnLoading.Value		= """ & ConvSPChars(rs1("UNLOADING")) & """" & vbCr
	Response.Write "parent.frm1.txtScrapProcess.Value	= """ & ConvSPChars(rs1("SCRAP_PROCESS")) & """" & vbCr
	Response.Write "parent.frm1.txtCustodyArea.Value	= """ & ConvSPChars(rs1("CUSTODY_AREA")) & """" & vbCr
	Response.Write "parent.frm1.txtCheckEndDt.Text		= """ & UNIDateClientFormat(rs1("CHECK_END_DT")) & """" & vbCr
	Response.Write "parent.frm1.txtRepEndDt.Text		= """ & UNIDateClientFormat(rs1("REP_END_DT")) & """" & vbCr
	'Response.Write "parent.frm1.txtPicFlag.Value		= """ & ConvSPChars(rs1("PIC_FLAG")) & """" & vbCr
	Response.Write "parent.frm1.txtInspPrid.Value		= """ & UNINumClientFormat(rs1("INSP_PRID"), ggQty.DecPoint, 0) & """" & vbCr
	Response.Write "parent.frm1.txtCurAccnt.Value		= """ & UNINumClientFormat(rs1("CUR_ACCNT"), ggQty.DecPoint, 0) & """" & vbCr
	Response.Write "parent.frm1.txtFinCurAccnt.Value	= """ & UNINumClientFormat(rs1("FIN_CUR_ACCNT"), ggQty.DecPoint, 0) & """" & vbCr
	Response.Write "parent.frm1.txtFinAjDt.Text			= """ & UNIDateClientFormat(rs1("FIN_AJ_DT")) & """" & vbCr
	Response.Write "parent.frm1.txtItemCd1.Value		= """ & ConvSPChars(rs1("ITEM_CD_1")) & """" & vbCr
	Response.Write "parent.frm1.txtItemCd2.Value		= """ & ConvSPChars(rs1("ITEM_CD_2")) & """" & vbCr
	Response.Write "parent.frm1.txtItemCd3.Value		= """ & ConvSPChars(rs1("ITEM_CD_3")) & """" & vbCr
	Response.Write "parent.frm1.txtItemCd4.Value		= """ & ConvSPChars(rs1("ITEM_CD_4")) & """" & vbCr
	Response.Write "parent.frm1.txtItemCd5.Value		= """ & ConvSPChars(rs1("ITEM_CD_5")) & """" & vbCr
	Response.Write "parent.frm1.txtItemCd6.Value		= """ & ConvSPChars(rs1("ITEM_CD_6")) & """" & vbCr
	Response.Write "parent.frm1.txtItemCd7.Value		= """ & ConvSPChars(rs1("ITEM_CD_7")) & """" & vbCr
	Response.Write "parent.frm1.txtItemCd8.Value		= """ & ConvSPChars(rs1("ITEM_CD_8")) & """" & vbCr
	Response.Write "parent.frm1.txtItemCd9.Value		= """ & ConvSPChars(rs1("ITEM_CD_9")) & """" & vbCr
	Response.Write "parent.frm1.txtItemCd10.Value		= """ & ConvSPChars(rs1("ITEM_CD_10")) & """" & vbCr
	Response.Write "parent.frm1.txtItemNm1.Value		= """ & ConvSPChars(rs1("ITEM_NM_1")) & """" & vbCr
	Response.Write "parent.frm1.txtItemNm2.Value		= """ & ConvSPChars(rs1("ITEM_NM_2")) & """" & vbCr
	Response.Write "parent.frm1.txtItemNm3.Value		= """ & ConvSPChars(rs1("ITEM_NM_3")) & """" & vbCr
	Response.Write "parent.frm1.txtItemNm4.Value		= """ & ConvSPChars(rs1("ITEM_NM_4")) & """" & vbCr
	Response.Write "parent.frm1.txtItemNm5.Value		= """ & ConvSPChars(rs1("ITEM_NM_5")) & """" & vbCr
	Response.Write "parent.frm1.txtItemNm6.Value		= """ & ConvSPChars(rs1("ITEM_NM_6")) & """" & vbCr
	Response.Write "parent.frm1.txtItemNm7.Value		= """ & ConvSPChars(rs1("ITEM_NM_7")) & """" & vbCr
	Response.Write "parent.frm1.txtItemNm8.Value		= """ & ConvSPChars(rs1("ITEM_NM_8")) & """" & vbCr
	Response.Write "parent.frm1.txtItemNm9.Value		= """ & ConvSPChars(rs1("ITEM_NM_9")) & """" & vbCr
	Response.Write "parent.frm1.txtItemNm10.Value		= """ & ConvSPChars(rs1("ITEM_NM_10")) & """" & vbCr
	Response.Write "parent.frm1.txtPrsUnit.Value		= """ & UNINumClientFormat(rs1("PRS_UNIT"), ggQty.DecPoint, 0)  & """" & vbCr
	Response.Write "parent.frm1.txtSetPlantCd1.Value	= """ & ConvSPChars(rs1("SET_PLANT")) & """" & vbCr
	Response.Write "parent.frm1.txtSetPlantNm1.Value	= """ & ConvSPChars(rs1("SET_PLANT_NM")) & """" & vbCr
	Response.Write "parent.frm1.cboPrsSts.Value			= """ & ConvSPChars(rs1("PRS_STS")) & """" & vbCr
	Response.Write "parent.frm1.cboEmpCd.Value			= """ & ConvSPChars(rs1("EMP_CD")) & """" & vbCr
	Response.Write "parent.frm1.cboUseYn.Value			= """ & ConvSPChars(rs1("USE_YN")) & """" & vbCr
	Response.Write "parent.frm1.txtSetPlace.Value		= """ & ConvSPChars(rs1("SET_PLACE")) & """" & vbCr
	'Response.Write "parent.frm1.cboOprNo.Value			= """ & ConvSPChars(rs1("OPR_NO")) & """" & vbCr
	Response.Write "parent.frm1.txtSetPlaceNm.Value			= """ & ConvSPChars(rs1("SET_PLACE_NM")) & """" & vbCr
	Response.Write "parent.frm1.txtPurCurCd.Value				= """ & ConvSPChars(rs1("PUR_CUR")) & """" & vbCr
	Response.Write "parent.frm1.txtLimitAccnt.Value				= """ & UNINumClientFormat(rs1("LIMIT_ACCNT"), ggQty.DecPoint, 0) & """" & vbCr
	
    Response.Write "parent.DbQueryOk2() "	& vbCr
	Response.Write "</Script>"		& vbCr

	Response.end

%>
