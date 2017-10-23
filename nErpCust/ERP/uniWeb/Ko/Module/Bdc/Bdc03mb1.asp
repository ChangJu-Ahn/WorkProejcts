<%@ LANGUAGE=VBSCript %>
<% Option Explicit%>
<!-- #Include file="../../inc/incSvrMain.asp"  -->
<!-- #Include file="../../inc/incSvrDate.inc"  -->
<!-- #Include file="../../inc/incSvrNumber.inc"-->
<%    
	Dim lgOpModeCRUD

	On Error GoTo 0
	Err.Clear

	Call HideStatusWnd

	Call LoadBasisGlobalInf()
	'------ Developer Coding part (Start ) ------------------------------------------------------------------

	'------ Developer Coding part (End   ) ------------------------------------------------------------------ 
	lgOpModeCRUD = Request("txtMode")

	Select Case lgOpModeCRUD
		Case CStr(UID_M0001)
			Call SubBizQueryMulti()
		Case CStr(UID_M0002)
			Call SubBizSaveMulti()
	End Select

'=========================================================================================================
Sub SubBizQueryMulti()
	Dim objBDC003
	Dim istrCode
	Dim lgStrPrevKey

	Dim iLngMaxRow
	Dim iLngRow
	Dim iStrNextKey
	Dim iStrData

	Dim E1_Z_Co_Mnu
	Const C_SHEETMAXROWS_D = 100

	' 컴포넌트에 넘겨줄 파라메터들 
	Const BDC003_PROCESS_ID     = 0
	Const BDC003_USE_FLAG       = 1
	Const BDC003_UNICODE_FLAG   = 2

	' 컴포넌트로 부터 넘겨받을 레코드 쎝 구성 파라메터들...
	Const PROCESS_ID   = 0
	Const PROCESS_NAME = 1
	Const USE_FLAG     = 2
	Const TRAN_FLAG    = 3
	Const RUN_TIME	   = 4
	Const START_ROW    = 5
	Const UPDATE_ID    = 6
	Const UPDATE_DT    = 7

	Redim istrCode(BDC003_UNICODE_FLAG)

	istrCode(BDC003_PROCESS_ID)   = FilterVar(Request("txtProcID"), "", "SNM")
	istrCode(BDC003_USE_FLAG)     = FilterVar(Request("cboUseYN"), "", "SNM")
	istrCode(BDC003_UNICODE_FLAG) = FilterVar("Y", "", "SNM")

	On Error Resume Next
	Err.Clear
	Set objBDC003 = Server.CreateObject("BDC003.clsManager")
	If CheckSYSTEMError(Err,True) = True Then
		Set objBDC003 = Nothing
		Exit Sub
	End If
	On Error Goto 0

	E1_Z_Co_Mnu = objBDC003.GetMasterList(gStrGlobalCollection, C_SHEETMAXROWS_D, istrCode)
	Set objBDC003 = Nothing

	If CheckSYSTEMError(Err,True) = True Then
		Response.Write   "<Script Language=vbscript>"            & vbCr
		Response.Write   "   Parent.frm1.txtProcessID.focus() "  & vbCr
		Response.Write   "</Script>"                             & vbCr
		Exit Sub
	End If

	iLngMaxRow = CLng(Request("txtMaxRows"))

	iStrData = ""
	If E1_Z_Co_Mnu(0,0) <> "" Then
		For iLngRow = 0 To UBound(E1_Z_Co_Mnu, 2)
			iStrData = iStrData & Chr(11) & E1_Z_Co_Mnu(PROCESS_ID, iLngRow)
			iStrData = iStrData & Chr(11) & ConvSPChars(E1_Z_Co_Mnu(PROCESS_NAME, iLngRow))
			iStrData = iStrData & Chr(11) & E1_Z_Co_Mnu(USE_FLAG, iLngRow)
			iStrData = iStrData & Chr(11) & ConvSPChars(E1_Z_Co_Mnu(TRAN_FLAG, iLngRow))
			iStrData = iStrData & Chr(11) & ConvSPChars(E1_Z_Co_Mnu(RUN_TIME, iLngRow))
			iStrData = iStrData & Chr(11) & ConvSPChars(E1_Z_Co_Mnu(START_ROW, iLngRow))
			iStrData = iStrData & Chr(11) & ConvSPChars(E1_Z_Co_Mnu(UPDATE_ID, iLngRow))
			iStrData = iStrData & Chr(11) & ConvSPChars(E1_Z_Co_Mnu(UPDATE_DT, iLngRow))

			iStrData = iStrData & Chr(11) & iLngMaxRow + ConvSPChars(iLngRow)
			iStrData = iStrData & Chr(11) & Chr(12)
		Next
	End If

	Response.Write "<Script Language=vbscript>"                         & vbCr
	Response.Write "With Parent "                                       & vbCr
	Response.Write "    .ggoSpread.Source = .frm1.vspdData "            & vbCr
	Response.Write "    .ggoSpread.SSShowData """ & iStrData & """"     & vbCr
	Response.Write "    .lgStrPrevKey = """ & iStrNextkey & """"        & vbCr    
	Response.Write "    .frm1.vspdData.ReDraw = True "                  & vbCr   
	Response.Write "    .DbQueryOk  "                                   & vbCr    
	Response.Write "End With "                                          & vbCr
	Response.Write "</Script>"                                          & vbCr
End Sub

'=========================================================================================================
Sub SubBizSaveMulti()
	Dim objBDC003
	Dim iErrorPosition
	Dim iStrSpread

	On Error Resume Next
	Err.Clear
	Set objBDC003 = Server.CreateObject("BDC003.clsManager")
	If CheckSYSTEMError(Err, True) = True Then
		Set objBDC003 = Nothing
		Exit Sub
	End If
	On Error Goto 0

	iStrSpread = FilterVar(Request("txtSpread"),"","SNM")

	Call objBDC003.UpdateProcess(gStrGlobalCollection, istrSpread, iErrorPosition)
	Set objBDC003 = Nothing

	If CheckSYSTEMError2(Err, True, iErrorPosition & "행","","","","") = True Then          
		Exit Sub
	End If

	Response.Write "<Script Language=vbscript>"          & vbCr
	Response.Write "Parent.DbSaveOk "                    & vbCr
	Response.Write "</Script>"
End Sub
%>
