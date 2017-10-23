<%
Option Explicit 
'**********************************************************************************************
'*  1. Module Name          : ACCOUNT
'*  2. Function Name        : 
'*  3. Program ID           : a3104mb2
'*  4. Program Name         : 가수금내역저장 
'*  5. Program Desc         : 가수금내역을 등록,수정 
'*  6. Complus List         : 
'*  7. Modified date(First) : 2000/03/30
'*  8. Modified date(Last)  : 2002/11/13
'*  9. Modifier (First)     : 
'* 10. Modifier (Last)      : Jeong Yong Kyun
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************

'☜ : 항상 서버 사이드 구문의 시작점인 좌꺽쇠(<)% 와 %우꺽쇠(>)는 New Line에 위치하여 
'	  서버 사이드 구문과 클라이언트 사이드 구문의 위치를 가늠할 수 있도록 한다.
'☜ : 아래 HTML 구문은 변경되어서는 안된다. 
%>
<!-- #Include file="../../inc/incSvrMain.asp"  -->
<!-- #Include file="../../inc/incSvrDate.inc"  -->
<!-- #Include file="../../inc/incSvrNumber.inc"  -->
<%																				'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 

	On Error Resume Next														'☜: 

	Call HideStatusWnd()														'☜: 모든 작업 완료후 작업진행중 표시창을 Hide
	Call LoadBasisGlobalInf()	
		
	Dim lgIntFlgMode
	Dim LngMaxRow

	Dim iPARG020
	Dim ImportTransType
	Dim ImportDocCur
	Dim ImportPartnerBpCd
	Dim ImportDeptCd
	Dim ImportARcpt
	Dim iChangeOrgId

	Dim ExportAGl
	Dim ExportRcpt

	Const RcptNo				= 0
	Const RcptDt				= 1
	Const DocCur				= 2
	Const XchRate				= 3
	Const BnkChgAmt				= 4
	Const BnkChgLocAmt			= 5
	Const RcptType				= 6
	Const RcptDesc				= 7
	Const RefNo					= 8
	
	'//기초치를 위한FLAG
	Const Rcpt_Input_Type			= 9		'//기초치구분 
	Const GlFlag				= 10	'//전표구분 
	Const RcptAmt				= 11
	Const RcptLocAmt			= 12
	Const ConfFg				= 13
	Const Project				= 14
	Const GlDt				= 15

	ReDim ImportARcpt(GLDt)
	Redim ImportDeptCd(1)
	
	iChangeOrgId = UCase(Request("hOrgChangeId"))

	lgIntFlgMode				= CInt(Request("txtFlgMode"))					'☜: 저장시 Create/Update 판별 

	ImportTransType				= "AR010"
	ImportDocCur				= Trim(Request("txtDocCur"))
	ImportPartnerBpCd			= Trim(Request("txtBpCd"))
	
	ImportDeptCd(0)				= UCase(Request("hOrgChangeId"))
	ImportDeptCd(1)				= Trim(Request("txtDept"))

	ImportARcpt(RcptNo)			= Trim(Request("txtRcptNo"))
	ImportARcpt(RcptDt)			= UNIConvDate(Request("txtRcptDt"))
	ImportARcpt(DocCur)			= Request("txtDocCur")
	ImportARcpt(XchRate)		= UNIConvNum(Request("txtXchRate"),0)
	ImportARcpt(BnkChgAmt)   	= 0
	ImportARcpt(BnkChgLocAmt)	= 0
	If "" & Trim(Request("txtRcptType")) = "" Then
		ImportARcpt(RcptType)	= "H2"
	Else
		ImportARcpt(RcptType)	= Trim(Request("txtRcptType"))
	End If
	ImportARcpt(RefNo)			= Trim(Request("txtRefNo"))
	ImportARcpt(RcptDesc)		= Request("txtDesc")
	
	'//기초치를 위한FLAG
	ImportARcpt(Rcpt_Input_Type)		= "RT"		'//가수금기초치등록 
	ImportARcpt(GlFlag)			= Trim(Request("txtGlFlag"))
	ImportARcpt(RcptAmt)		= UNIConvNum(Request("txtRcptAmt"),0)
	ImportARcpt(RcptLocAmt)		= UNIConvNum(Request("txtRcptLocAmt"),0)
	ImportARcpt(ConfFg)			= Trim(Request("txtConfFg"))
	ImportARcpt(Project)		= Trim(Request("txtProject"))
	ImportARcpt(GLDt)			= UNIConvDate(Request("htxtGlDt"))
	
	
	Set iPARG020 = Server.CreateObject("PARG020.cAMngRcSvr")

	If CheckSYSTEMError(Err,True) = True Then
		Response.End 
	End If
	
	If lgIntFlgMode = OPMD_CMODE Then
		Call iPARG020.MANAGE_RCPT_SVR(gStrGlobalCollection, "CREATE", ImportTransType, gCurrency, _
		                                       ImportARcpt, ImportPartnerBpCd, ImportDeptCd, ,_
		                                       ExportRcpt,ExportAGl)
	ElseIf lgIntFlgMode = OPMD_UMODE Then
		Call  iPARG020.MANAGE_RCPT_SVR(gStrGlobalCollection, "UPDATE", ImportTransType, gCurrency, _
		                                       ImportARcpt, ImportPartnerBpCd, ImportDeptCd, ,_
		                                       ExportRcpt,ExportAGl)
	End If

	If CheckSYSTEMError(Err,True) = True Then
		Set iPARG020 = Nothing		
		Response.End 
	End If
	    
	Set iPARG020 = Nothing

	Response.Write "<Script Language=vbscript>" & vbcr
	Response.Write " parent.frm1.txtRcptNo.Value	= """ & ConvSPChars(ExportRcpt)	& """" & vbcr
	Response.Write " parent.dbSaveOk          " & vbcr
	Response.Write "</Script>" & vbcr
%>
