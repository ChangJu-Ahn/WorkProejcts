<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% Option Explicit%>
<% session.CodePage=949 %>

<%
'**********************************************************************************************
'*  1. Module Name          : 원가 
'*  2. Function Name        : 공정별 배부내역조회 
'*  3. Program ID           : c4211mb1
'*  4. Program Name         : 공정별 배부내역조회 
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
	Dim sToSenderCostCd, sFromSenderCostCd, sFromAcctCd, sToAcctCd, sType, sSendAmt, sRecvAmt, sWCCd
	
    'On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

	' -- 변경해야할 조회조건 (MA에서 보내주는)
	Dim sStartDt, sEndDt
	
	sStartDt		= Request("txtStartDt")	
	sNextKey		= Request("lgStrPrevKey")
	
	sFromSenderCostCd	= Request("txtFROM_SENDER_COST_CD")	
	sToSenderCostCd	= Request("txtTO_SENDER_COST_CD")	
	sFromAcctCd		= Request("txtFROM_ACCT_CD")	
	sToAcctCd		= Request("txtTO_ACCT_CD")
	sWCCd			= Request("txtWC_CD")		
	sType			= Request("rdoTYPE")	
		
	If sStartDt = "" Or sType = "" Then
		Call DisplayMsgBox("900015", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
		Exit Sub
	End If
	
	If sFromSenderCostCd = "" Then sFromSenderCostCd = ""
	If sToSenderCostCd = "" Then sToSenderCostCd = "zzzzzzzzzz"
	If sFromAcctCd = "" Then sFromAcctCd = ""
	If sToAcctCd = "" Then sToAcctCd = "zzzzzzzzzzzzzzzz"
	If sWCCd = "" Then sWCCd = "%"
	
    With lgObjComm
		.CommandTimeout = 0
		
		.CommandText = "dbo.usp_C_C4211MA1_TYPE" & sType		' --  변경해야할 SP 명 
	    .CommandType = adCmdStoredProc

		lgObjComm.Parameters.Append lgObjComm.CreateParameter("RETURN_VALUE",  adInteger,adParamReturnValue)	' -- No 수정 

		' -- 변경해야할 조회조건 파라메타들 
		lgObjComm.Parameters.Append lgObjComm.CreateParameter("@START_DT",	adVarXChar,	adParamInput, 6,Replace(sStartDt, "'", "''"))
		
		lgObjComm.Parameters.Append lgObjComm.CreateParameter("@FROM_SENDER_COST_CD",	adVarXChar,	adParamInput, 10,Replace(sFromSenderCostCd, "'", "''"))
		lgObjComm.Parameters.Append lgObjComm.CreateParameter("@TO_SENDER_COST_CD",	adVarXChar,	adParamInput, 10,Replace(sToSenderCostCd, "'", "''"))
		
		lgObjComm.Parameters.Append lgObjComm.CreateParameter("@WC_CD",	adVarXChar,	adParamInput, 7,Replace(sWCCd, "'", "''"))

		lgObjComm.Parameters.Append lgObjComm.CreateParameter("@FROM_ACCT_CD",	adVarXChar,	adParamInput, 20,Replace(sFromAcctCd, "'", "''"))
		lgObjComm.Parameters.Append lgObjComm.CreateParameter("@TO_ACCT_CD",	adVarXChar,	adParamInput, 20,Replace(sToAcctCd, "'", "''"))
		
		lgObjComm.Parameters.Append lgObjComm.CreateParameter("@SHEETCNT",  adSmallInt, adParamInput,, 100)	
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
    If Not oRs.EOF Then

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
			
			sTxt	= Replace(sTxt, "%CS", "합계")
					
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
		
		If sType = "1" Then
			iMaxCols	= 3	+ iLngColCnt	' 그룹바이 컬럼행수 
		Else
			iMaxCols	= 3	+ iLngColCnt	' 그룹바이 컬럼행수 
		End If
		
		' -- 데이타셋을 기초로 배열로 재구성한다.
		ReDim arrTemp(iLngColCnt, iLngGrRowCnt)
		
		' -- 데이타셋 행수 
		iLngRowCnt	= UBound(arrDataSet, 2)
		
		For iLngRow = 0 To 	iLngRowCnt
			arrTemp(CLng(arrDataSet(1, iLngRow)), CLng(arrDataSet(0, iLngRow))-1) = arrDataSet(2, iLngRow)	' -- 행, 열, 값(썸)
		Next
		
		' -- 배부할 금액/배부받은 금액 
		If sType = "1" And sNextKey= ""  Then
			Set oRs = oRs.NextRecordSet()	' -- 다음(데이타) 레코드셋으로 지정 

			If Not oRs.EOF Then
				sSendAmt = oRs(0)
				sRecvAmt = oRs(1)
			End If
		End If
		
		Set oRs = Nothing

		sRowSeq = arrRows(UBound(arrRows, 1) -1 , iLngGrRowCnt)		' -- 다은키값 (변경분)
		
		If CInt(iLngGrRowCnt) < 100-1 Then
			sRowSeq = ""
		End If

		' -- 그룹바이 행으로 루핑 
		For iLngRow = 0 To 	iLngGrRowCnt
				
			If sType = "1" Then
				iStrData = iStrData & Chr(11) & ConvLang(ConvSPChars(Trim(arrRows(0, iLngRow))), arrRows(iLngGrColCnt-2, iLngRow))	' -- SENDER_COST_CD
				iStrData = iStrData & Chr(11) & ConvLang(ConvSPChars(Trim(arrRows(1, iLngRow))), "0")					' -- SENDER_COST_NM
				iStrData = iStrData & Chr(11) & Trim(arrRows(2, iLngRow))					' -- SEND_AMT

				' -- 데이타그리드 출력(수정분)
				For iLngCol = 1 To iLngColCnt 
					iStrData = iStrData & Chr(11) & arrTemp(iLngCol, iLngRow)
				Next

				If Trim(arrRows(3, iLngRow)) <> "0" Then	' -- 소계행을 구분함 
					sGrpTxt = sGrpTxt & arrRows(3, iLngRow) & gColSep & arrRows(4, iLngRow) & gRowSep		' -- 소계구분|행번호(배열의 위치에 주의)
				End If
				
				iStrData = iStrData & Chr(11) & arrRows(3, iLngRow)
				iStrData = iStrData & Chr(11) & arrRows(4, iLngRow)
				
			Else	 ' -- B 인 경우 
				iStrData = iStrData & Chr(11) & ConvLang2(ConvSPChars(Trim(arrRows(0, iLngRow))), arrRows(iLngGrColCnt-1, iLngRow))	' -- ACCT_CD
				iStrData = iStrData & Chr(11) & ConvLang2(ConvSPChars(Trim(arrRows(1, iLngRow))), "0")					' -- ACCT_NM
				iStrData = iStrData & Chr(11) & Trim(arrRows(2, iLngRow))					' -- RECV_AMT

				' -- 데이타그리드 출력(수정분)
				For iLngCol = 1 To iLngColCnt 
					iStrData = iStrData & Chr(11) & arrTemp(iLngCol, iLngRow)
				Next

				If Trim(arrRows(3, iLngRow)) <> "0" Then	' -- 소계행을 구분함 
					sGrpTxt = sGrpTxt & arrRows(3, iLngRow) & gColSep & arrRows(4, iLngRow) & gRowSep		' -- 소계구분|행번호(배열의 위치에 주의)
				End If
				
				iStrData = iStrData & Chr(11) & arrRows(3, iLngRow)
				iStrData = iStrData & Chr(11) & arrRows(4, iLngRow)

			End If

			iStrData = iStrData & Chr(12)
		Next
				
		Response.Write " <Script Language=vbscript>	                        " & vbCr
		Response.Write " With parent                                        " & vbCr
		
		If sType = "1" Then

			If sNextKey = "" Then
				Response.Write " .frm1.txtSEND_AMT.value =""" & UNINumClientFormat(sSendAmt,ggQty.DecPoint,0) & """" & vbCr
				Response.Write " .frm1.txtRECV_AMT.value =""" & UNINumClientFormat(sRecvAmt,ggQty.DecPoint,0) & """" & vbCr		
			End If
			
			Response.Write "	Call Parent.InitSpreadSheet(" & iMaxCols+1 & ")			" & vbCr 			 
			Response.Write "	.frm1.vspdData.ReDraw = False					" & vbCr 			 
			Response.Write "	.ggoSpread.Source = .frm1.vspdData              " & vbCr 			 
			Response.Write "	.ggoSpread.SSShowData """ & iStrData		       & """" & vbCr
			Response.Write "	.frm1.vspdData.ReDraw = True					" & vbCr 			 
			Response.Write "	.lgStrPrevKey = """ & sRowSeq & """" & vbCr
			
			Response.Write "	.frm1.hSTART_DT.value = """ & sStartDt & """" & vbCr
			Response.Write "	.frm1.hWC_CD.value = """ & sWCCd & """" & vbCr 	
			Response.Write "	.frm1.hFROM_SENDER_COST_CD.value = """ & sFromSenderCostCd & """" & vbCr 	
			Response.Write "	.frm1.hTO_SENDER_COST_CD.value = """ & sToSenderCostCd & """" & vbCr 	
			Response.Write "	.frm1.hFROM_ACCT_CD.value = """ & sFromAcctCd & """" & vbCr 	
			Response.Write "	.frm1.hTO_ACCT_CD.value = """ & sToAcctCd & """" & vbCr 	 	
			
			If sNextKey = "" Then
				Response.Write  "   Call Parent.SetGridHead(""" & sTxt & """)" & vbCr
			End If

			If sGrpTxt <> "" Then
				Response.Write  "   Call Parent.SetQuerySpreadColor(""" & sGrpTxt & """)" & vbCr
			End If
			
			'Response.Write "	.frm1.vspdData.MaxCols = """ & (iMaxCols) & """" & vbCr 	 	
			Response.Write  "   Call Parent.DbQueryOk()		" & vbCr
		Else
			Response.Write "	Call Parent.InitSpreadSheet2(" & iMaxCols+1 & ")			" & vbCr 		
			Response.Write "	.frm1.vspdData2.ReDraw = False					" & vbCr 			 
			Response.Write "	.ggoSpread.Source = .frm1.vspdData2              " & vbCr 			 
			Response.Write "	.ggoSpread.SSShowData """ & iStrData		       & """" & vbCr
			Response.Write "	.frm1.vspdData2.ReDraw = True					" & vbCr 			 
			Response.Write "	.lgStrPrevKey2 = """ & sRowSeq & """" & vbCr

			If sNextKey = "" Then
				Response.Write  "   Call Parent.SetGridHead2(""" & sTxt & """)" & vbCr
			End If

			If sGrpTxt <> "" Then
				Response.Write  "   Call Parent.SetQuerySpreadColor2(""" & sGrpTxt & """)" & vbCr
			End If

			'Response.Write "	.frm1.vspdData2.MaxCols = """ & (iMaxCols) & """" & vbCr 	 	
			Response.Write  "   Call Parent.DbQueryOk2()		" & vbCr
		End If
		
		Response.Write " End With                                        " & vbCr
		Response.Write  " </Script>                  " & vbCr
    ElseIf sType = "1" And sNextKey = "" Then
       Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
    End If

End Sub	

Function ConvLang(Byval pLang, Byval GroupNo)
	Dim pTmp
	
	If GroupNo <> "0" Then
		pTmp = Replace(pLang , "%1", "합계")
		pTmp = Replace(pTmp , "%2", "소계")
	Else
		pTmp = pLang
	End If
	ConvLang = pTmp
End Function


Function ConvLang2(Byval pLang, Byval GroupNo)
	Dim pTmp
	
	If GroupNo <> "0" Then
		pTmp = Replace(pLang , "%1", "합계")
	Else
		pTmp = pLang
	End If
	ConvLang2 = pTmp
End Function
%>

