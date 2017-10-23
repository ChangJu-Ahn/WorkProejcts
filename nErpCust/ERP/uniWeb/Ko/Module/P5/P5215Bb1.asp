<%@LANGUAGE	=	VBScript%>
<%Option Explicit%>
<!-- #Include	file="../../inc/IncSvrMain.asp"	-->
<!-- #Include	file="../../inc/IncSvrNumber.inc"	-->
<!-- #Include	file="../../inc/IncSvrDate.inc"	-->
<!-- #Include	file="../../inc/IncSvrDBAgent.inc" -->
<!-- #Include	file="../../inc/IncSvrDBAgentVariables.inc"	-->
<!-- #Include	file="../../ComAsp/LoadInfTB19029.asp" -->

<% Call	LoadBasisGlobalInf
'**********************************************************************************************
'*	1. Module	Name			:	Production
'*	2. Function	Name		:
'*	3. Program ID			:	p2210mb2.asp
'*	4. Program Name			:	MPS	老褒 积己 
'*	5. Program Desc			:
'*	6. Comproxy	List		:	PY5G215.cPExecMrpSvr
'*	7. Modified	date(First)	:
'*	8. Modified	date(Last) 	:	2002/06/18
'*	9. Modifier	(First)		:	Lee	Hyun Jae
'* 10. Modifier	(Last)		:	Jung Yu	Kyung
'* 11. Comment				:
'**********************************************************************************************

Call HideStatusWnd

On Error Resume	Next


'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'+ 1.	MPS	老褒积己 
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

	Dim	pPY5G215
	Dim	PY5_Fac_Cast_Parameter
	Dim	WorkDt
	Err.Clear

	Const	PY5_Y5_work_dt = 0
	Const	PY5_Y5_fac_cast_cd = 1


	Redim	PY5_Fac_Cast_Parameter(PY5_Y5_fac_cast_cd)


'		WorkDt = UniConvDateToYYYYMMDD(Request("txtWork_dt"),gDateFormat,"")


	PY5_Fac_Cast_Parameter(PY5_Y5_work_dt) = Request("txtWork_dt")
	PY5_Fac_Cast_Parameter(PY5_Y5_fac_cast_cd) = UCase(Trim(Request("txtFacility_Cd")))

'		Call DisplayMsgBox("x",	vbInformation, "捞惑窍尺"	,	"FASDFADS1111",	I_MKSCRIPT)


	Set	pPY5G215 = Server.CreateObject("PY5G215.cPY5ExecCreatCast")

	If CheckSYSTEMError(Err,True)	=	True Then
		Set	pPY5G215 = Nothing
	End	If

	Call pPY5G215.PY5_ExecCreatCast(gStrGlobalCollection,	PY5_Fac_Cast_Parameter)

	If CheckSYSTEMError(Err,True)	=	True Then
		Set	pPY5G215 = Nothing
	Else
		Call DisplayMsgBox("183114", vbInformation,	"",	"",	I_MKSCRIPT)
	End	If



%>
