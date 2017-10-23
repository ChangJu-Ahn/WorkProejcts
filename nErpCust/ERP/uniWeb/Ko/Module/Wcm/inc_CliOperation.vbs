	
' -- 서식정보 
Dim wgConfirmFlg
Dim wgRefDoc
Dim wgStatusFlg
Const C_REVISION_YM = "200703"

' -- 폼 로드시/데이타 조회시 서식정보를 로딩한다.
Function CheckTaxDoc(pCoCd, pFiscYear, pRepType, pPGM_ID)

	call CommonQueryRs(" CONFIRM_FLG, STATUS_FLG"," TB_TAX_DOC_DTL "," CO_CD= '" & pCoCd & "' AND FISC_YEAR='" & pFiscYear & "' AND REP_TYPE='" & pRepType & "' AND PGM_ID='" & pPGM_ID & "'",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
	wgConfirmFlg	= Replace(lgF0,Chr(11),"")
	wgStatusFlg		= Replace(lgF1,Chr(11),"")
	
	If wgConfirmFlg = "1" Then 
		wgConfirmFlg = "Y"
	Else
		wgConfirmFlg = "N"
	End If

	If wgConfirmFlg = "Y" Then
		Call ggoOper.LockField(Document, "Q")
		Call DisplayMsgBox("WC0038", Parent.VB_INFORMATION, "x", "x")
	ElseIf wgConfirmFlg = "" Then
		Call DisplayMsgBox("WC0031", Parent.VB_INFORMATION, "x", "x")
	End If
End Function

' -- 레퍼런스 가져오기할때 서식정보를 로딩한다.
Function GetDocRef(pCoCd, pFiscYear, pRepType, pPGM_ID)

	call CommonQueryRs("REF_DOC"," TB_TAX_DOC "," PGM_ID='" & pPGM_ID & "'",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
	wgRefDoc		= Replace(lgF0,Chr(11),"")
	GetDocRef		= wgRefDoc

End Function

Function RtnQueryVal(strField,strFrom,strWhere)
    Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
    RtnQueryVal = ""
    Call CommonQueryRs(strField,strFrom,strWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    RtnQueryVal = Replace(lgF0,Chr(11),"")
    If RtnQueryVal = "X" Or trim(RtnQueryVal) = "" Or IsNull(RtnQueryVal) Then
        Call DisplayMsgBox("970000", vbInformation, strWhere & strField, "", I_MKSCRIPT)
           
        ObjectContext.SetAbort
        Call SetErrorStatus
	End If
End Function

' -- 윤년 체크 
Function CheckIntercalaryYear(Byval datYear)
	
	If (datYear Mod 4 =0 And datYear Mod 100 > 0) Or datYear Mod 400 = 0 then
		CheckIntercalaryYear = True
	Else
		CheckIntercalaryYear = False
	End If
End Function

' -- 환경코드 콤보 
Sub SetComboX(Byref pCombo, ByVal pCodeArr, ByVal pNameArr, ByVal pSeqNo1Arr, Byval pSeqNo2Arr, ByVal pSeperator)

    Dim iDx, arrCode, arrName, arrSeqNo1, arrSeqNo2, objEl, iLen

    arrCode		= Split(pCodeArr,pSeperator)
    arrName		= Split(pNameArr,pSeperator)
    arrSeqNo1	= Split(pSeqNo1Arr ,pSeperator)
    arrSeqNo2	= Split(pSeqNo2Arr ,pSeperator)
    iLen = UBound(arrCode)
    
    For iDx = 0 To iLen - 1
		Set objEl = document.createElement("OPTION")
		objEl.Value	= arrCode(iDx)
		objEl.Text = arrName(iDx)
		objEl.SetAttribute "VAL", arrSeqNo1(iDx)
		objEl.SetAttribute "VIEW", arrSeqNo2(iDx)
		pcombo.Add objEl
		Set objEl = Nothing
    Next

End Sub


'---셀랙트된 text 색지정 
Function SelectColor(ByVal obj)
    On Error Resume Next
    if obj.tagName =  "INPUT" then
       obj.Style.background   =  "#99ff99"
    else
        obj.BackColor =&H009BF0A2&
    end if
  
End Function