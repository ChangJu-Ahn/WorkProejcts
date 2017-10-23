<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% Option Explicit%>
<% session.CodePage=949 %>

<%
'**********************************************************************************************
'*  1. Module Name          : 원가 
'*  2. Function Name        : 직과가공비집계조회 
'*  3. Program ID           : c4224mb1
'*  4. Program Name         : 직과가공비집계조회 
'*  5. Program Desc         :
'*  6. Comproxy List        :
'*  7. Modified date(First) : 2005/09/27
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     :choe0tae 
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
	Dim sCostElmtCd, sAcctCd, sItemCd, sType, sGrid, sCostCd, sMinorCd
	
    'On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

	' -- 변경해야할 조회조건 (MA에서 보내주는)
	Dim sStartDt, sEndDt
	
	sStartDt	= Request("txtStartDt")	
	sEndDt		= Request("txtEndDt")
	sNextKey	= Request("lgStrPrevKey")
	
	sCostElmtCd	= Request("txtCOST_ELMT_CD")	
	sCostCd		= Request("txtCOST_CD")	
	sAcctCd		= Request("txtACCT_CD")	
	sMinorCd	= Request("txtMINOR_CD")	
	sType		= Request("rdoTYPE")	
	sGrid		= Request("txtGrid")	
		
	If sStartDt = "" And sEndDt = ""  And sCostElmtCd = "" And sAcctCd = "" And sGrid = "" Then
		Call DisplayMsgBox("900015", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
		Exit Sub
	End If
	
	If sCostElmtCd = "" Then sCostElmtCd = "%"
	If sCostCd = "" Then sCostCd = "%"

	
    With lgObjComm
		.CommandTimeout = 0
		
		.CommandText = "dbo.usp_C_C4224MA1_TYPE" & sType & "_" & sGrid		' --  변경해야할 SP 명 
	    .CommandType = adCmdStoredProc

		lgObjComm.Parameters.Append lgObjComm.CreateParameter("RETURN_VALUE",  adInteger,adParamReturnValue)	' -- No 수정 

		' -- 변경해야할 조회조건 파라메타들 
		lgObjComm.Parameters.Append lgObjComm.CreateParameter("@START_DT",	adVarXChar,	adParamInput, 10,Replace(sStartDt, "'", "''"))
		lgObjComm.Parameters.Append lgObjComm.CreateParameter("@END_DT",	adVarXChar,	adParamInput, 10,Replace(sEndDt, "'", "''"))

		lgObjComm.Parameters.Append lgObjComm.CreateParameter("@COST_CD",	adVarXChar,	adParamInput, 10,Replace(sCostCd, "'", "''"))
		
		If sGrid = "A" Then	' -- 헤더그리드 
			lgObjComm.Parameters.Append lgObjComm.CreateParameter("@COST_ELMT_CD",	adVarXChar,	adParamInput, 20,Replace(sCostElmtCd, "'", "''"))
		Else	' -- 디테일 그리드 
			lgObjComm.Parameters.Append lgObjComm.CreateParameter("@ACCT_CD",	adVarXChar,	adParamInput, 20,Replace(sAcctCd, "'", "''"))
			lgObjComm.Parameters.Append lgObjComm.CreateParameter("@MINOR_CD",	adVarXChar,	adParamInput, 3,Replace(sMinorCd, "'", "''"))
		End If
		
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
		' -- 그룹바이 컬럼 정보를 기초로 데이타쉬트를 구성한다.(변경분)
		arrRows		= oRs.GetRows()
		Set oRs = Nothing
		
		iLngRowCnt= UBound(arrRows, 2)				' 그룹바이 행수 
		iLngColCnt = UBound(arrRows, 1)				' 그룹바이 열수 
		
		If sGrid = "A" Then
			sRowSeq = arrRows(UBound(arrRows, 1) , iLngRowCnt)		' -- 다은키값 (변경분)
		End If
		
		If CInt(sRowSeq) < 100 Then
			sRowSeq = ""
		End If
		' -- 그룹바이 행으로 루핑 
		For iLngRow = 0 To 	iLngRowCnt
				
			If sGrid = "A" Then
				iStrData = iStrData & Chr(11) & ConvLang(ConvSPChars(Trim(arrRows(0, iLngRow))), arrRows(iLngColCnt-2, iLngRow))	' -- cost_cd
				iStrData = iStrData & Chr(11) & ConvLang(ConvSPChars(Trim(arrRows(1, iLngRow))), "0")					' -- cost_nm
				iStrData = iStrData & Chr(11) & ConvLang(ConvSPChars(Trim(arrRows(2, iLngRow))), arrRows(iLngColCnt-2, iLngRow))	' -- COST_ELMT_CD
				iStrData = iStrData & Chr(11) & ConvLang(ConvSPChars(Trim(arrRows(3, iLngRow))), "0")					' -- COST_ELMT_NM
				iStrData = iStrData & Chr(11) & ConvLang(ConvSPChars(Trim(arrRows(4, iLngRow))), arrRows(iLngColCnt-2, iLngRow))	' -- ACCT_CD
				iStrData = iStrData & Chr(11) & ConvLang(ConvSPChars(Trim(arrRows(5, iLngRow))), "0")						' -- ACCT_NM
				iStrData = iStrData & Chr(11) & ConvLang(ConvSPChars(Trim(arrRows(6, iLngRow))), "0")					' -- MINOR_NM
				iStrData = iStrData & Chr(11) & ConvLang(ConvSPChars(Trim(arrRows(7, iLngRow))), "0")					' -- MINOR_NM
				iStrData = iStrData & Chr(11) & ConvLang(ConvSPChars(Trim(arrRows(8, iLngRow))), "0")					' -- MINOR_NM

				' -- 데이타그리드 출력(수정분)
				For iLngCol = 9 To iLngColCnt 
					iStrData = iStrData & Chr(11) & arrRows(iLngCol, iLngRow)	' -- sum_amt
				Next
				If Trim(arrRows(iLngColCnt-1, iLngRow)) <> "0" Then	' -- 소계행을 구분함 
					sGrpTxt = sGrpTxt & arrRows(iLngColCnt-1, iLngRow) & gColSep & arrRows(iLngColCnt, iLngRow) & gRowSep		' -- 소계구분|행번호(배열의 위치에 주의)
				End If			
			Else	 ' -- B 인 경우 
				' -- 데이타그리드 출력(수정분)
				For iLngCol = 0 To iLngColCnt 
					iStrData = iStrData & Chr(11) & ConvSPChars(Trim(arrRows(iLngCol, iLngRow)))
				Next
				iStrData = iStrData & Chr(11) & iLngRow+1
			End If										
			iStrData = iStrData & Chr(12)
		Next
				
		Response.Write " <Script Language=vbscript>	                        " & vbCr
		Response.Write " With parent                                        " & vbCr
		
		If sGrid = "A" Then		
			If sNextKey= ""  Then
				Response.Write "	.InitSpreadSheet(12)					" & vbCr 			 		
			End If
			Response.Write "	.frm1.vspdData.ReDraw = False					" & vbCr 			 
			Response.Write "	.ggoSpread.Source = .frm1.vspdData              " & vbCr 			 
			Response.Write "	.ggoSpread.SSShowData """ & iStrData		       & """" & vbCr
			Response.Write "	.frm1.vspdData.ReDraw = True					" & vbCr 			 
			Response.Write "	.lgStrPrevKey = """ & sRowSeq & """" & vbCr
			Response.Write "	.frm1.hSTART_DT.value = """ & sStartDt & """" & vbCr 	
			Response.Write "	.frm1.hEND_DT.value = """ & sEndDt & """" & vbCr
			Response.Write "	.frm1.hCOST_CD.value = """ & sCostCd & """" & vbCr 	
			Response.Write "	.frm1.hCOST_ELMT_CD.value = """ & sCostElmtCd & """" & vbCr 	 	
			Response.Write "	.frm1.hTYPE.value = """ & sType & """" & vbCr 	 	
			
			If sGrpTxt <> "" Then
				Response.Write  "   Call Parent.SetQuerySpreadColor(""" & sGrpTxt & """)" & vbCr
			End If			
			Response.Write  "   Call Parent.DbQueryOk()		" & vbCr
		Else		
			Response.Write  "   Call Parent.InitSpreadSheet2(" & iLngColCnt+1 & ")		" & vbCr			
			Response.Write "	.frm1.vspdData2.ReDraw = False					" & vbCr 			 
			Response.Write "	.ggoSpread.Source = .frm1.vspdData2              " & vbCr 			 
			Response.Write "	.ggoSpread.SSShowData """ & iStrData		       & """" & vbCr
			Response.Write "	.frm1.vspdData2.ReDraw = True					" & vbCr 	
			Response.Write  "   Call Parent.DbQueryOk2()		" & vbCr
		End If		
		Response.Write " End With                                        " & vbCr
		Response.Write  " </Script>                  " & vbCr
    Else		' ----------------------------------------------------------		
		If sGrid = "A" And sNextKey= ""  Then
			Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
		End If
		
		Response.Write " <Script Language=vbscript>	                        " & vbCr
		Response.Write " With parent                                        " & vbCr			
		If sGrid = "A" Or sNextKey = "" Then
			Response.Write " If Parent.frm1.vspdData.MaxRows > 0 Then Parent.frm1.vspdData.Focus"& vbCr	
		Else
			Response.Write " If Parent.frm1.vspdData2.MaxRows > 0 Then Parent.frm1.vspdData2.Focus"& vbCr	
		End If			
		If sGrid = "A" Then
			Response.Write " Parent.lgEOF1 = True" & vbCr
		Else
			Response.Write " Parent.lgEOF2 = True" & vbCr
		End If			
		Response.Write " End With                                        " & vbCr
		Response.Write  " </Script>                  " & vbCr
		Exit Sub
    End If
	Set oRs = Nothing
End Sub	

Function ConvLang(Byval pLang, Byval GroupNo)
	Dim pTmp
	
	If GroupNo <> "0" Then
		pTmp = Replace(pLang , "%1", "합계")
		pTmp = Replace(pTmp , "%2", "요소소계")
		pTmp = Replace(pTmp , "%3", "계정소계")
	Else
		pTmp = pLang
	End If
	ConvLang = pTmp
End Function


%>

