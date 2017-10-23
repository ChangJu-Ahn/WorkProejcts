<%@ LANGUAGE=VBSCript %>
<%Option Explicit%>

<!-- #Include file="../../inc/incSvrMain.asp"  -->
<!-- #Include file="../../inc/incSvrDate.inc"  -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<%													'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 
Call LoadBasisGlobalInf()

Call HideStatusWnd

    On Error Resume Next
    Err.Clear
													'☜: 현재 MyBiz.asp 의 진행상태를 나타냄 
Dim dtDate
Dim startIndex
Dim lastDay
Dim i
Dim strTempGl, strGl

   '---------------------------------------Common-----------------------------------------------------------
    lgErrorStatus     = "NO"
    lgErrorPos        = ""                          '☜: Set to space
    lgOpModeCRUD      = Request("txtMode")          '☜: Read Operation Mode (CRUD)
													'value : 1500(일반쿼리), 1501(수정)

    strTempGl		  = Request("htxtTempGl")		'☜: 결의 
    strGl             = Request("htxtGl")			'☜: 회계 

    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)						'☜: Query
             Call SubBizQuery()
        Case CStr(UID_M0002)						'☜: Save,Update
             Call SubBizSave()
    End Select

'    Call SubCloseDB(lgObjConn)						'☜: Close DB Connection

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()

    On Error Resume Next
    Err.Clear

	Dim PABG025Data
	Dim strYYYYMM 
	Dim CalCol

	Dim strDate
    Dim lgarrData		'달력정보 저장 
    Dim strMonth 

	Const A042_EG1_E1_b_calendar_calendar_dt = 0
    Const A042_EG1_E1_b_calendar_day_of_week = 1
    Const A042_EG1_E1_b_calendar_hol_type = 2
    Const A042_EG1_E1_b_calendar_remark = 3
    Const A042_EG1_E1_b_calendar_gl_fg = 4
    Const A042_EG1_E1_b_calendar_temp_gl_fg = 5	

    ReDim lgarrData(31, 5)

	If Request("txtYear") = "" Or Request("txtMonth") = "" Then				'⊙: 조회를 위한 값이 들어왔는지 체크 
		Call DisplayMsgBox("700112", vbOKOnly, "", "", I_MKSCRIPT)			'조회조건값이 비어있습니다.
		Response.End
	End If

    strYYYYMM = Right("0000" & Request("txtYear"), 4)
    strYYYYMM = strYYYYMM & Right("00" & Request("txtMonth"), 2)

 '   Response.Write  "<<" & gStrGlobalCollection & ">>"

	Set PABG025Data = server.CreateObject ("PABG040.cBListCalendar") 

	If CheckSYSTEMError(Err,True) = True Then
		Exit Sub
    End If


	'DEBUG
	'gStrGlobalCollection : connection 정보 저장 변수 
	'Response.Write "gStrGlobalCollection = " & gStrGlobalCollection(incServer.asp)

	lgarrData = PABG025Data.B_READ_CALENDAR(gStrGlobalCollection,strYYYYMM)

	If CheckSYSTEMError(Err,True) = True Then
		Set PABG025 = nothing
		Exit Sub
    End If

	Set PABG025Data = nothing

	strDate = left(lgarrData(0,0),10)	
    dtDate = CDate(strDate)    
    startIndex = WeekDay(dtDate) - 1	

    'DEBUG
	'Response.Write "STRDATE  = " & strdate & "<BR>"
	'Response.Write "MONTH = " & strMonth & "<BR>"
	'Response.Write "STARTINDEX= " & STARTINDEX & "<BR>"

	'화면에 해당월 정보 표시 
	Response.Write "<Script Language=vbscript>  " & vbCr
   	Response.Write " with parent" & vbCr
	Response.Write " .frm1.txtYear.value       = """ & ConvSPChars(Year(dtDate))  & """	        	" & vbcr
	Response.Write " .frm1.txtMonth.value      = """ & ConvSPChars(Month(dtDate)) & """	    	    " & vbcr
    Response.Write " .lgStartIndex             = " & startIndex    & "				" & vbcr
    Response.Write " .document.all.tbTitle.Rows(0).Cells(0).innerText = """ & Year(dtDate) & "." & Month(dtdate) & """" & vbcr    
	Response.Write " End with					" & vbcr
    Response.Write " </Script>                   " & vbCr


	dtDate = UNIDateAdd("M", 1, dtDate, gApDateFormat)
	dtDate = UNIDateAdd("D", -1, dtDate, gApDateFormat)
	lastDay = Day(dtDate)

	'지난달 Display를 위해서....
    dtDate = CDate(strDate)

	'Response.Write "첫날 = " & dtDate &"<BR>"
	'Response.Write "마지막날 = " & LASTDAY &"<BR>"

	Response.Write "<Script Language=vbscript>  " & vbCr
    Response.Write " Parent.lgLastDay =  " & lastDay  & "	    	    " & vbCr
    Response.Write " </Script>                   " & vbCr

	'1일 이전 데이타 클리어 
	Response.Write "<Script Language=vbscript>  " & vbCr
    Response.Write " Dim CalCol  " & vbCr
    Response.write " Dim iIntCount " & vbCr
    Response.Write " iIntCount = 0 " & vbCr    
    Response.Write " For CalCol = " & startIndex & " -1 " & " to 0 Step-1       " & vbCr    
    Response.Write " with parent.frm1 " & vbCr    
    Response.Write " .txtDate(CalCol).value =  CStr( " & Day(DateAdd("d", -1, dtDate))  & " +iIntCount " & ")" & vbCr
	Response.Write " .txtDate(CalCol).className = ""DummyDay""             " & vbCr
    Response.Write " .txtDate(CalCol).disabled  =  true   		           " & vbCr
    Response.Write " .txtT(CalCol).value = """"                            " & vbCr
    Response.Write " .txtT(CalCol).style.cursor = """"                     " & vbCr
    Response.Write " .txtT(CalCol).disabled =  true 		            " & vbCr
    Response.Write " .txtFlgT(CalCol).value = """"				         " & vbCr
    Response.Write " .txtFlgT(CalCol).disabled =  true        		    " & vbCr
    Response.Write " .txtG(CalCol).value = """"                            " & vbCr 
    Response.Write " .txtG(CalCol).style.cursor = """"                     " & vbCr
    Response.Write " .txtG(CalCol).disabled = """"               		     " & vbCr
    Response.Write " .txtFlgG(CalCol).value = """"                         " & vbCr
    Response.Write " .txtFlgG(CalCol).disabled =  true          		 " & vbCr
    Response.Write " .txtDesc(CalCol).value = """"                         " & vbCr 
    Response.Write " .txtDesc(CalCol).title = """"                         " & vbCr
    Response.Write " End with					                         " & vbcr
    Response.Write " iIntCount = iIntCount - 1							 " & vbCr
    Response.Write " Next                                                " & vbCr
    Response.Write " </Script>                                           " & vbCr
    'DEBUG
    'Response.Write "LASTDAY = " & LASTDAY

    '해당 요일에 대한 모든 정보를 보여준다.
    For i = 1 To lastDay
		If lgarrData(i-1,A042_EG1_E1_b_calendar_hol_type) = "H" Then
			'휴일일때 
			Response.Write "<Script Language=vbscript>  " & vbCr	
			Response.Write " Parent.frm1.txtDate(  " & startIndex  & " ).style.color = ""red""     " & vbCr
			Response.Write " </Script>                                           " & vbCr
		Else
			If (startIndex + 1) Mod 7 = 0 Then
				'토요일일때 
				Response.Write "<Script Language=vbscript>  " & vbCr
				Response.Write " Parent.frm1.txtDate( " & startIndex & " ).style.color = ""blue""      " &vbCR
				Response.Write " </Script>										" & vbCr 					
			Else
				'평일 
				Response.Write "<Script Language=vbscript>  " & vbCr
				Response.Write " Parent.frm1.txtDate( " & startIndex & " ).style.color = ""black""       "&vbCR
				Response.Write " </Script>										" & vbCr
			End if
		End if

		if lgarrData(i-1,A042_EG1_E1_b_calendar_temp_gl_fg) = "C" Then
			'결의 
			Response.Write "<Script Language=vbscript>  " & vbCr
			Response.Write " Parent.frm1.txtT( " & startIndex & " ).style.color = ""blue""		" & vbCr
			Response.Write " </Script>										" & vbCr
		Else
			Response.Write "<Script Language=vbscript>  " & vbCr
		  	Response.Write " Parent.frm1.txtT( " & startIndex & " ).style.color = ""silver""		" & vbCr
		  	Response.Write " </Script>										" & vbCr
		End If

		If lgarrData(i-1, A042_EG1_E1_b_calendar_gl_fg) =  "C" Then
			'회계 
			Response.Write "<Script Language=vbscript>  " & vbCr
			Response.Write " Parent.frm1.txtG( " & startIndex & " ).style.color = ""red""			" & vbCr
			Response.Write " </Script>										" & vbCr
		Else
			Response.Write "<Script Language=vbscript>  " & vbCr
			Response.Write " Parent.frm1.txtG(  " & startIndex & " ).style.color = ""silver""    " & vbCr
			Response.Write " </Script>										" & vbCr
		End If

		Response.Write "<Script Language=vbscript>  " & vbCr
		Response.Write " with parent.frm1" & vbCr
		Response.Write " .txtDate( " & startIndex & " ).value = " & i & "           " & vbCr
		Response.Write " .txtDate( " & startIndex & " ).className = ""Day""              " & vbCr
		Response.Write " .txtDate( " & startIndex & " ).disabled = False               "&vbCr

		Response.Write " .txtT( " & startIndex & ").value = """ & strTempGl & """   " & vbCr                          ' T 결의 
		Response.Write " .txtT( " & startIndex & ").style.cursor = ""Hand""           " & vbCr
		Response.Write " .txtT( " & startIndex & ").disabled = False                " & vbCr
		Response.Write " .txtFlgT( " & startIndex & ").value = """ & lgarrData(i-1,A042_EG1_E1_b_calendar_temp_gl_fg) & """   " & vbCr

		Response.Write " .txtFlgT( " & startIndex & ").disabled = False             "&vbCr

		Response.Write " .txtG( " & startIndex & ").value = """ & strGl & """       "&vbCr                           ' G 회계 
		Response.Write " .txtG( " & startIndex & ").style.cursor = ""Hand""           "&vbCr 
		Response.Write " .txtG( " & startIndex & ").disabled = False				"&vbCr 
		Response.Write " .txtFlgG( " & startIndex & ").value = """ & lgarrData(i-1,A042_EG1_E1_b_calendar_gl_fg) & """        "&vbCr
		Response.Write " .txtFlgG( " & startIndex & ").disabled = False				" & vbCr

		Response.Write " .txtDesc( " & startIndex & ").value = """ & ConvSPChars(lgarrData(i-1,A042_EG1_E1_b_calendar_remark)) & """		 "&vbCr 
		Response.Write " .txtDesc( " & startIndex & ").title = """ & ConvSPChars(lgarrData(i-1,A042_EG1_E1_b_calendar_remark)) & """		 "&vbCr
		Response.Write "End with										" & vbcr
		Response.Write " </Script>										" & vbCr
		startIndex = startIndex + 1
	Next

	Response.Write "<Script Language=vbscript>  " & vbCr
	Response.Write " For CalCol = " & startIndex & " to 41                       " & vbCr
	Response.Write " with parent.frm1" & vbCr
	Response.Write " .txtDate(CalCol).value = CStr(CalCol - " & startIndex-1 & ")    " & vbCr
	Response.Write " .txtDate(CalCol).className = ""DummyDay""							 " & vbCr
	Response.Write " .txtDate(CalCol).disabled = True									 " & vbCr

	Response.Write " .txtT(CalCol).value = """"											 " & vbCr
	Response.Write " .txtT(CalCol).style.cursor = """"                                     " & vbCr
	Response.Write " .txtT(CalCol).disabled = True                                       " & vbCr
	Response.Write " .txtFlgT(CalCol).value = """"                                         " & vbCr
	Response.Write " .txtFlgT(CalCol).disabled = True                                    " & vbCr

    Response.Write " .txtG(CalCol).value = """"                                            " & vbCr
	Response.Write " .txtG(CalCol).style.cursor = """"                                     " & vbCr
	Response.Write " .txtG(CalCol).disabled = True                                       " & vbCr
    Response.Write " .txtFlgG(CalCol).value = """"                                         " & vbCr
	Response.Write " .txtFlgG(CalCol).disabled = True                                    " & vbCr

    Response.Write " .txtDesc(CalCol).value = """"                                         " & vbCr
	Response.Write " .txtDesc(CalCol).title = """"                                         " & vbCr
	Response.Write " End with										" & vbcr
	Response.Write " Next											" & vbcr

	Response.Write " Parent.lgNextNo = """"		                    " & vbcr          ' 다음 키 값 넘겨줌 
	Response.Write " Parent.lgPrevNo = """"		                    " & vbcr          ' 이전 키 값 넘겨줌  
    Response.Write " Parent.DbQueryOk			                    " & vbcr
    Response.Write " </Script>										" & vbCr '☜: 조회가 성공 

	Response.End

End Sub

'============================================================================================================
Sub SubBizSave()

	Dim PABG025Data
	Dim lgarrData
	Dim LoopStr
	Dim strYYYYMM
	Dim i

	Const A002_IG1_I1_b_calendar_calendar_dt = 0
    Const A002_IG1_I1_b_calendar_gl_fg = 1
    Const A002_IG1_I1_b_calendar_temp_gl_fg = 2
    LoopStr = Request("txtFlgT").count

	Redim lgarrData(LoopStr,A002_IG1_I1_b_calendar_temp_gl_fg)

    On Error Resume Next
    Err.Clear

	If Request("txtYear") = "" Or Request("txtMonth") = "" Then				'⊙: 조회를 위한 값이 들어왔는지 체크 
		Call DisplayMsgBox("700112", vbOKOnly, "", "", I_MKSCRIPT)			'조회조건값이 비어있습니다.
		Response.End
	End If

    strYYYYMM = Right("0000" & Request("txtYear"), 4)
    strYYYYMM = strYYYYMM & Right("00" & Request("txtMonth"), 2)
	lgIntFlgMode = CInt(Request("txtFlgMode"))								'☜: 저장시 Create/Update 판별 

	Set PABG025Data = Server.CreateObject("PABG040.cBCalCloseGlDtSvr")

	If CheckSYSTEMError(Err,True) = True Then
		Exit Sub
    End If

	For i = 1 To LoopStr
		'현화면의 달력정보를 저장한다.
		dtDate = (Request("txtYear") & "-" & Request("txtMonth") & "-" & i)
		lgarrData(i, A002_IG1_I1_b_calendar_calendar_dt) = dtDate
		lgarrData(i, A002_IG1_I1_b_calendar_temp_gl_fg) = Request("txtFlgT")(i)
		lgarrData(i, A002_IG1_I1_b_calendar_gl_fg) = Request("txtFlgG")(i)
    Next

    Call PABG025Data.B_CALENDAR_CLOSE_GL_DT_SVR(gStrGlobalCollection, strYYYYMM, lgarrData )

    If CheckSYSTEMError(Err,True) = True Then
		Set PB6G1010 = nothing
		Exit Sub
    End If

    Set PABG025Data = nothing

	Response.Write "<Script Language=vbscript>  " & vbCr
	Response.Write " parent.DbSaveOk            " & vbCr
    Response.Write "</Script>                   " & vbCr

End Sub

'============================================================================================================
Sub CommonOnTransactionCommit()
End Sub

'============================================================================================================
Sub CommonOnTransactionAbort()
End Sub

'============================================================================================================
Sub SetErrorStatus()
End Sub
%>
