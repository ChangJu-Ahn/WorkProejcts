<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% Option Explicit%>
<% session.CodePage=949 %>

<%
'**********************************************************************************************
'*  1. Module Name          : cost 
'*  2. Function Name        : 회계가공비집계조회 
'*  3. Program ID           : c4214mb1
'*  4. Program Name         : 회계가공비집계조회 
'*  5. Program Desc         :
'*  6. Comproxy List        : ADO
'*  7. Modified date(First) : 2005/09/27
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : choe0tae
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************
%>

<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->

<%

	Call LoadBasisGlobalInf()
    Call HideStatusWnd                                                               '☜: Hide Processing message
	'------ Developer Coding part (Start ) ------------------------------------------------------------------

	'------ Developer Coding part (End   ) ------------------------------------------------------------------ 
	
     Call SubCreateCommandObject(lgObjComm)
     
     Call SubBizQuery()
     
     Call SubCloseCommandObject(lgObjComm)
     
     Response.End 
'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()

	Dim oRs, sTxt, arrRows, iLngRow, iLngCol, iStrData, sNextKey, sRowSeq, iLngRowCnt, iLngColCnt, sGrpTxt
	Dim sCostElmtCd, sFromAcctCd, sToAcctCd, sItemCd, sType, sGrid, sCostCd
	
    'On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

	Const C_MAX_SHEET_COUNT = 100
	' -- 변경해야할 조회조건 (MA에서 보내주는)
	Dim sStartDt, sEndDt
	
	sStartDt	= Request("txtStartDt")	
	sEndDt		= Request("txtEndDt")
	sNextKey	= Request("lgStrPrevKey")
	
	sCostElmtCd	= Request("txtCOST_ELMT_CD")	
	sFromAcctCd	= Request("txtFROM_ACCT_CD")
	sToAcctCd	= Request("txtTO_ACCT_CD")		
	sCostCd		= Request("txtCOST_CD")

	sType		= Request("rdoTYPE")	
	sGrid		= Request("txtGrid")	
		
	If sStartDt = "" And sEndDt = ""  And sCostElmtCd = "" And sFromAcctCd = "" And sGrid = "" Then
		Call DisplayMsgBox("900015", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
		Exit Sub
	End If
	
	If sCostElmtCd = "" Then sCostElmtCd = "%"
	If sToAcctCd = "" Then sToAcctCd = "zzzzzzzzzzzzzz"
	If sCostCd = "" Then sCostCd = "%"
	If Instr(1, sCostCd, "%") = 0 Then sCostCd	= sCostCd & "%"

	'Response.Write "sStartDt=" & sStartDt & vbcrlf
	'Response.Write "sEndDt=" & sEndDt & vbcrlf
	'Response.Write "sNextKey=" & sNextKey & vbcrlf
	
    With lgObjComm
		.CommandTimeout = 0
		
		.CommandText = "dbo.usp_C_C4214MA1_TYPE" & sType & "_" & sGrid		' --  변경해야할 SP 명 
	    .CommandType = adCmdStoredProc

		lgObjComm.Parameters.Append lgObjComm.CreateParameter("RETURN_VALUE",  adInteger,adParamReturnValue)	' -- No 수정 

		' -- 변경해야할 조회조건 파라메타들 
		lgObjComm.Parameters.Append lgObjComm.CreateParameter("@START_DT",	adVarXChar,	adParamInput, 6,Replace(sStartDt, "'", "''"))
		lgObjComm.Parameters.Append lgObjComm.CreateParameter("@END_DT",	adVarXChar,	adParamInput, 6,Replace(sEndDt, "'", "''"))
		
		lgObjComm.Parameters.Append lgObjComm.CreateParameter("@COST_ELMT_CD",	adVarXChar,	adParamInput, 20,Replace(sCostElmtCd, "'", "''"))
		lgObjComm.Parameters.Append lgObjComm.CreateParameter("@COST_CD",	adVarXChar,	adParamInput, 10,Replace(sCostCd, "'", "''"))
		
		lgObjComm.Parameters.Append lgObjComm.CreateParameter("@FROM_ACCT_CD",	adVarXChar,	adParamInput, 20,Replace(sFromAcctCd, "'", "''"))
		lgObjComm.Parameters.Append lgObjComm.CreateParameter("@TO_ACCT_CD",	adVarXChar,	adParamInput, 20,Replace(sToAcctCd, "'", "''"))
		
		lgObjComm.Parameters.Append lgObjComm.CreateParameter("@SHEETCNT",  adSmallInt, adParamInput,, C_MAX_SHEET_COUNT)	
		lgObjComm.Parameters.Append lgObjComm.CreateParameter("@NEXTKEY",	adVarXChar,	adParamInput, 15,Replace(sNextKey, "'", "''"))
		lgObjComm.Parameters.Append lgObjComm.CreateParameter("@DEBUG",  adSmallInt, adParamInput,, 0)	' -- isqlw 에서만 사용하는 디버깅코드 
		    
        Set oRs = lgObjComm.Execute
        
    End With
    'Response.Write "Err=" & Err.Description
    If Instr( Err.Description , "B_MESSAGE") > 0 Then
		If HandleBMessageError(vbObjectError, Err.Description, "", "") = True Then
			Exit Sub
		End If
	Else
		If CheckSYSTEMError(Err, True) = True Then	
			Exit Sub
		End If
	End If
		
    
    ' -- 리턴 레코드셋은 총 2-3종 
    If Not oRs.Eof Then

		If sNextKey = "" Then	' -- 헤더 레코드셋 포함되어 옴 

			' --- 헤더 출력(변경분) ----			
			Dim arrColRow, i, j, ColHeadRowNo
			
			ColHeadRowNo = -1000
			arrColRow = oRs.GetRows()
			iLngRowCnt	= UBound(arrColRow, 2) 
			iLngColCnt	= UBound(arrColRow, 1) 
			
			For j = 0 To iLngColCnt 
				For i = 0 To iLngRowCnt 
					sTxt = sTxt & arrColRow(j, i) & gColSep 
				Next
				sTxt = sTxt & CStr(ColHeadRowNo) & gRowSep
				ColHeadRowNo = ColHeadRowNo + 1
			Next
			
			sTxt	= Replace(sTxt, "%4", "합계")
			sTxt	= Replace(sTxt, "%5", "사업부별 합계")
			sTxt	= Replace(sTxt, "%6", "Cost Center")
			sTxt	= Replace(sTxt, "%7", "공 정")
					
			Set oRs = oRs.NextRecordSet()	' -- 다음(데이타) 레코드셋으로 지정 
			' ----------------------------------------
		End If
		
		' -- 그룹바이 컬럼 정보를 기초로 데이타쉬트를 구성한다.(변경분)
		arrRows		= oRs.GetRows()
		
		Set oRs = oRs.NextRecordSet()	' -- 다음(데이타) 레코드셋으로 지정 
		
		Dim arrDataSet, arrTemp, iLngGrRowCnt, iLngGrColCnt, iMaxCols
		arrDataSet = oRs.GetRows()


		iLngGrRowCnt= UBound(arrRows, 2)				' 그룹바이 행수 
		iLngGrColCnt = UBound(arrRows, 1)				' 그룹바이 열수 
		
		iLngColCnt	= CLng(arrRows(UBound(arrRows, 1), 0))					' 컬럼수 
		
		If sGrid = "A" Then
			iMaxCols	= 7	+ iLngColCnt	' 그룹바이 컬럼행수 
		Else
			iMaxCols	= 7	+ iLngColCnt	' 그룹바이 컬럼행수 
		End If
		
		' -- 데이타셋을 기초로 배열로 재구성한다.
		ReDim arrTemp(iLngColCnt, iLngGrRowCnt)
		
		' -- 데이타셋 행수 
		iLngRowCnt	= UBound(arrDataSet, 2)
		
		For iLngRow = 0 To 	iLngRowCnt
			arrTemp(CLng(arrDataSet(1, iLngRow)), CLng(arrDataSet(0, iLngRow))-1) = arrDataSet(2, iLngRow)	' -- 행, 열, 값(썸)
		Next
		
		Set oRs = Nothing
		
		' ----------------------------------------------------------
		If iLngGrRowCnt = 0 Then 
		
			If sNextKey= ""  Then
				Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
			End If

			Exit Sub
		End If

		sRowSeq = arrRows(UBound(arrRows, 1) -1 , iLngGrRowCnt)		' -- 다은키값 (변경분)

		If iLngGrRowCnt < C_MAX_SHEET_COUNT -1 Then sRowSeq = ""

		' -- 그룹바이 행으로 루핑 
		For iLngRow = 0 To 	iLngGrRowCnt
				
			iStrData = iStrData & Chr(11) & ConvLang(ConvSPChars(Trim(arrRows(0, iLngRow))), arrRows(iLngGrColCnt-2, iLngRow))								' -- PLANT_CD
			iStrData = iStrData & Chr(11) & ConvLang(ConvSPChars(Trim(arrRows(1, iLngRow))), arrRows(iLngGrColCnt-2, iLngRow))	' -- COST_ELMT_CD
			iStrData = iStrData & Chr(11) & ConvLang(ConvSPChars(Trim(arrRows(2, iLngRow))), "0")					' -- COST_ELMT_NM
			iStrData = iStrData & Chr(11) & ConvLang(ConvSPChars(Trim(arrRows(3, iLngRow))), arrRows(iLngGrColCnt-2, iLngRow))	' -- ACCT_CD
			iStrData = iStrData & Chr(11) & ConvLang(ConvSPChars(Trim(arrRows(4, iLngRow))), "0")						' -- ACCT_NM
			iStrData = iStrData & Chr(11) & ConvLang(ConvSPChars(Trim(arrRows(5, iLngRow))), "0")					' -- MINOR_NM

			' -- 데이타그리드 출력(수정분)
			For iLngCol = 1 To iLngColCnt 
				iStrData = iStrData & Chr(11) & arrTemp(iLngCol, iLngRow)
			Next

			If Trim(arrRows(6, iLngRow)) <> "0" Then	' -- 소계행을 구분함 
				sGrpTxt = sGrpTxt & arrRows(6, iLngRow) & gColSep & arrRows(7, iLngRow) & gRowSep		' -- 소계구분|행번호(배열의 위치에 주의)
			End If
				
			'iStrData = iStrData & Chr(11) & arrRows(6, iLngRow)
			iStrData = iStrData & Chr(11) & arrRows(7, iLngRow)
										
			'iStrData = iStrData & Chr(11) & iLngRow+1
			iStrData = iStrData & Chr(12)
		Next
				
		Response.Write " <Script Language=vbscript>	                        " & vbCr
		Response.Write " With parent                                        " & vbCr

		Response.Write "	.frm1.hSTART_DT.value = """ & sStartDt & """" & vbCr 	
		Response.Write "	.frm1.hEND_DT.value = """ & sEndDt & """" & vbCr 	
		Response.Write "	.frm1.hTYPE.value = """ & sType & """" & vbCr 	 	
		Response.Write "	.frm1.hCOST_CD.value = """ & sCostCd & """" & vbCr 	 	
		Response.Write "	.frm1.hCOST_ELMT_CD.value = """ & sCostElmtCd & """" & vbCr 	 	
		Response.Write "	.frm1.hFROM_ACCT_CD.value = """ & sFromAcctCd & """" & vbCr 	 	
		Response.Write "	.frm1.hTO_ACCT_CD.value = """ & sToAcctCd & """" & vbCr 	 	

		If sNextKey = "" Then
			Response.Write  "   Call Parent.InitSpreadSheet(" & iMaxCols & ")" & vbCr
		End If
		
		Response.Write "	.frm1.vspdData.ReDraw = False					" & vbCr 			 
		Response.Write "	.ggoSpread.Source = .frm1.vspdData              " & vbCr 			 
		Response.Write "	.ggoSpread.SSShowData """ & iStrData		       & """" & vbCr
		Response.Write "	.frm1.vspdData.ReDraw = True					" & vbCr 			 
		Response.Write "	.lgStrPrevKey = """ & sRowSeq & """" & vbCr
			
		If sNextKey = "" Then
			Response.Write  "   Call Parent.SetGridHead(""" & sTxt & """)" & vbCr
		End If

		If sGrpTxt <> "" Then
			Response.Write  "   Call Parent.SetQuerySpreadColor(""" & sGrpTxt & """)" & vbCr
		End If
			
		Response.Write "	.frm1.vspdData.MaxCols = """ & (iMaxCols) & """" & vbCr 	 	
		Response.Write  "   Call Parent.DbQueryOk()		" & vbCr
		
		Response.Write " End With                                        " & vbCr
		Response.Write  " </Script>                  " & vbCr
    ElseIf sNextKey = "" Then
       Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
    End If

End Sub	

Function ConvLang(Byval pLang, Byval GroupNo)
	Dim pTmp
	
	If GroupNo <> "0" Then
		pTmp = Replace(pLang , "%1", "합계")
		pTmp = Replace(pTmp , "%2", "원가요소별소계")
		pTmp = Replace(pTmp , "%3", "계정별소계")
	Else
		pTmp = pLang
	End If
	ConvLang = pTmp
End Function


Function ConvLang2(Byval pLang, Byval GroupNo)
	Dim pTmp
	
	If GroupNo <> "0" Then
		pTmp = Replace(pLang , "%1", "합계")
		pTmp = Replace(pTmp , "%2", "관리항목별 소계")
	Else
		pTmp = pLang
	End If
	ConvLang2 = pTmp
End Function
%>

