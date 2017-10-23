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
    Dim objBDC004
    Dim istrCode
    Dim lgStrPrevKey

    Dim iLngMaxRow
    Dim iLngRow
    Dim iStrNextKey
    
    Dim iStrData
    Dim TmpBuffer
    
    Dim E1_Z_Co_Mnu
    Const C_SHEETMAXROWS_D = 100

    ' 컴포넌트에 넘겨줄 파라메터들 
    Const I_PROCESS_ID  = 0
    Const I_JOB_ID      = 1
    Const I_REGISTER_ID = 2
    Const I_FROM_DT     = 3
    Const I_TO_DT       = 4
    Const I_JOB_STATE   = 5
    Const I_RESULT_CD   = 6

    ' 컴포넌트로 부터 넘겨받을 레코드 쎝 구성 파라메터들...
    Const O_PROCESS_ID = 0
    Const O_PROCESS_NM = 1
    Const O_JOB_ID     = 2
    Const O_JOB_TITLE  = 3
    Const O_JOB_STATE  = 4
    Const O_PLAN_TIME  = 5
    Const O_HRESULT    = 6
    Const O_TOTAL      = 7
    Const O_SUCCESS    = 8
    Const O_FAIL	   = 9
    Const O_START_TIME = 10
    Const O_STOP_TIME  = 11

    Redim istrCode(I_RESULT_CD)

    istrCode(I_PROCESS_ID)  = FilterVar(Request("txtProcessID"), "", "SNM")
    If Request("lgIntFlgMode") = CStr(OPMD_UMODE) Then
		istrCode(I_JOB_ID)      = FilterVar(Request("lgStrPrevKey"), "", "SNM")
	Else
		istrCode(I_JOB_ID)      = FilterVar(Request("txtJobID"), "", "SNM")
	End If	
    istrCode(I_REGISTER_ID) = FilterVar(Request("txtRegisterID"), "", "SNM")
    istrCode(I_FROM_DT)     = FilterVar(Request("txtTrnsFrDt"), "", "SNM")
    istrCode(I_TO_DT)       = FilterVar(Request("txtTrnsToDt") & " 23:59", "", "SNM")
    istrCode(I_JOB_STATE)   = FilterVar(Request("cboJobState"), "", "SNM")
    istrCode(I_RESULT_CD)   = FilterVar(Request("txtResultCD"), "", "SNM")

    On Error Resume Next
    Err.Clear
    Set objBDC004 = Server.CreateObject("BDC004.clsJobManager")
    If CheckSYSTEMError(Err,True) = True Then
        Set objBDC004 = Nothing
        Exit Sub
    End If
    On Error Goto 0

    E1_Z_Co_Mnu = objBDC004.GetJobList (gStrGlobalCollection, C_SHEETMAXROWS_D, istrCode)
    Set objBDC004 = Nothing

    If CheckSYSTEMError(Err,True) = True Then
        Response.Write   "<Script Language=vbscript>"            & vbCr
        Response.Write   "   Parent.frm1.txtProcessID.focus() "  & vbCr
        Response.Write   "</Script>"                             & vbCr
        Exit Sub
    End If

    iLngMaxRow = CLng(Request("txtMaxRows"))
    iStrData = ""
    
    If E1_Z_Co_Mnu(0,0) <> "" Then
		If Ubound(E1_Z_Co_Mnu, 2) + 1 <= C_SHEETMAXROWS_D Then
			Redim TmpBuffer(Ubound(E1_Z_Co_Mnu, 2))
		Else
			Redim TmpBuffer(C_SHEETMAXROWS_D)
		End IF		
		
        For iLngRow = 0 To UBound(E1_Z_Co_Mnu, 2)
			If iLngRow + 1 > C_SHEETMAXROWS_D Then
				iStrNextKey = ConvSPChars(E1_Z_Co_Mnu(O_JOB_ID, iLngRow))
				Exit For
			End If	
            iStrData = Chr(11) & ""      'Check Box
            iStrData = iStrData & Chr(11) & ConvSPChars(E1_Z_Co_Mnu(O_PROCESS_ID, iLngRow))
            iStrData = iStrData & Chr(11) & ConvSPChars(E1_Z_Co_Mnu(O_PROCESS_NM, iLngRow))
            iStrData = iStrData & Chr(11) & ConvSPChars(E1_Z_Co_Mnu(O_JOB_ID, iLngRow))
            iStrData = iStrData & Chr(11) & ConvSPChars(E1_Z_Co_Mnu(O_JOB_TITLE, iLngRow))
            
            If ConvSPChars(E1_Z_Co_Mnu(O_JOB_STATE, iLngRow)) = "W" Then
		        iStrData = iStrData & Chr(11) & "대기"
            ElseIf ConvSPChars(E1_Z_Co_Mnu(O_JOB_STATE, iLngRow)) = "R" Then
	            iStrData = iStrData & Chr(11) & "실행"
            ElseIf ConvSPChars(E1_Z_Co_Mnu(O_JOB_STATE, iLngRow)) = "D" Then
	            iStrData = iStrData & Chr(11) & "완료"
	           ElseIf ConvSPChars(E1_Z_Co_Mnu(O_JOB_STATE, iLngRow)) = "C" Then
	            iStrData = iStrData & Chr(11) & "취소"
	        End If
	        
            If ConvSPChars(E1_Z_Co_Mnu(O_HRESULT, iLngRow)) = "S" Then
		        iStrData = iStrData & Chr(11) & "성공"
            ElseIf ConvSPChars(E1_Z_Co_Mnu(O_HRESULT, iLngRow)) = "R" Then
	            iStrData = iStrData & Chr(11) & ""
            ElseIf ConvSPChars(E1_Z_Co_Mnu(O_HRESULT, iLngRow)) = "F" Then
	            iStrData = iStrData & Chr(11) & "실패"
	        Else
	            iStrData = iStrData & Chr(11) & ConvSPChars(E1_Z_Co_Mnu(O_HRESULT, iLngRow))
	        End If

            iStrData = iStrData & Chr(11) & ConvSPChars(E1_Z_Co_Mnu(O_TOTAL, iLngRow))
            iStrData = iStrData & Chr(11) & ConvSPChars(E1_Z_Co_Mnu(O_SUCCESS, iLngRow))
            iStrData = iStrData & Chr(11) & ConvSPChars(E1_Z_Co_Mnu(O_FAIL, iLngRow))
            iStrData = iStrData & Chr(11) & ConvSPChars(E1_Z_Co_Mnu(O_PLAN_TIME, iLngRow))
            iStrData = iStrData & Chr(11) & ConvSPChars(E1_Z_Co_Mnu(O_START_TIME, iLngRow))
            iStrData = iStrData & Chr(11) & ConvSPChars(E1_Z_Co_Mnu(O_STOP_TIME, iLngRow))
            iStrData = iStrData & Chr(11) & iLngMaxRow + ConvSPChars(iLngRow)
            iStrData = iStrData & Chr(11) & Chr(12)
            TmpBuffer(iLngRow) = iStrData
        Next
        
        Response.Write "<Script Language=vbscript>"						& vbCr
		Response.Write "With Parent "									& vbCr
		
		Response.Write "	.frm1.hProcessID.value = """ & Request("txtProcessID") & """" & vbCr
		Response.Write "	.frm1.hRegisterID.value = """ & Request("txtRegisterID") & """" & vbCr
		Response.Write "	.frm1.hJobID.value =		""" & Request("txtJobID") & """" & vbCr
		Response.Write "	.frm1.hTrnsFrDt.value =	""" & Request("txtTrnsFrDt") & """" & vbCr
		Response.Write "	.frm1.hTrnsToDt.value =	""" & Request("txtTrnsToDt") & """" & vbCr
		Response.Write "	.frm1.hJobState.value =	""" & Request("cboJobState") & """" & vbCr
		Response.Write "	.frm1.hResultCD.value =	""" & Request("txtResultCD") & """" & vbCr
		Response.Write "	.frm1.vspdData.ReDraw = False " & vbCr
		Response.Write "    .ggoSpread.Source = .frm1.vspdData "		& vbCr
		Response.Write "    .ggoSpread.SSShowDataByClip """ & Join(TmpBuffer, "") & """"	& vbCr
		Response.Write "	.frm1.vspdData.ReDraw = True " & vbCr
		Response.Write "    .lgStrPrevKey = """ & iStrNextkey & """"    & vbCr    
		Response.Write "    .frm1.vspdData.ReDraw = True "              & vbCr   
		Response.Write "    .DbQueryOk(""" & Request("txtMaxRows") & """)   "                               & vbCr    
		Response.Write "End With "                                      & vbCr
		Response.Write "</Script>"
		
    Else
		Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
		Response.Write   "<Script Language=vbscript>"            & vbCr
        Response.Write   "   Parent.frm1.txtProcessID.focus() "  & vbCr
        Response.Write   "</Script>" 
		Exit Sub    
    End If

End Sub

'=========================================================================================================
Sub SubBizSaveMulti()
    Dim objBDC004
    Dim iErrorPosition
    Dim iStrSpread
  
    'On Error Resume Next
    Err.Clear
    Set objBDC004 = Server.CreateObject("BDC004.clsJobManager")
    If CheckSYSTEMError(Err,True) = True Then
        Set objBDC004 = Nothing
        Exit Sub
    End If
    On Error Goto 0

    iStrSpread = Request("txtSpread")
    Response.Write Request("hAction")
    '  Response.End 
    if Request("hAction") ="D" then 
		Call objBDC004.DeleteJob(gStrGlobalCollection, istrSpread, iErrorPosition)
	else
		
		Call objBDC004.UpdateJob(gStrGlobalCollection, istrSpread, iErrorPosition)
    end if
    
    Set objBDC004 = Nothing
    
    If CheckSYSTEMError2(Err, True, iErrorPosition & "행","","","","") = True Then          
       Exit Sub
    End If

    Response.Write "<Script Language=vbscript>"  & vbCr
    Response.Write "Parent.OpenCanJobOk "            & vbCr
    Response.Write "</Script>"
End Sub
%>
