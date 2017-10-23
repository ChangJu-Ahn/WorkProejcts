<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% Option Explicit%>
<% session.CodePage=949 %>

<%
'**********************************************************************************************
'*  1. Module Name          : ���� 
'*  2. Function Name        : ����������������ȸ 
'*  3. Program ID           : c4224mb2
'*  4. Program Name         : ����������������ȸ 
'*  5. Program Desc         :
'*  6. Comproxy List        : ADO
'*  7. Modified date(First) : 2005/09/27
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : choe0tae 
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
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
    Call HideStatusWnd                                                               '��: Hide Processing message
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
	
    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status

	' -- �����ؾ��� ��ȸ���� (MA���� �����ִ�)
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
		Call DisplayMsgBox("900015", vbInformation, "", "", I_MKSCRIPT)      '�� : No data is found. 
		Exit Sub
	End If
	
	If sCostElmtCd = "" Then sCostElmtCd = "%"
	If sCostCd = "" Then sCostCd = "%"
	
    With lgObjComm
		.CommandTimeout = 0		
		.CommandText = "dbo.usp_C_C4224MA1_TYPE" & sType & "_" & sGrid		' --  �����ؾ��� SP �� 
	    .CommandType = adCmdStoredProc

		lgObjComm.Parameters.Append lgObjComm.CreateParameter("RETURN_VALUE",  adInteger,adParamReturnValue)	' -- No ���� 
		lgObjComm.Parameters.Append lgObjComm.CreateParameter("@START_DT",	adVarXChar,	adParamInput, 10,Replace(sStartDt, "'", "''"))
		lgObjComm.Parameters.Append lgObjComm.CreateParameter("@END_DT",	adVarXChar,	adParamInput, 10,Replace(sEndDt, "'", "''"))
		lgObjComm.Parameters.Append lgObjComm.CreateParameter("@COST_CD",	adVarXChar,	adParamInput, 10,Replace(sCostCd, "'", "''"))
		
		If sGrid = "A" Then	' -- ����׸��� 
			lgObjComm.Parameters.Append lgObjComm.CreateParameter("@COST_ELMT_CD",	adVarXChar,	adParamInput, 20,Replace(sCostElmtCd, "'", "''"))
		Else	' -- ������ �׸��� 
			lgObjComm.Parameters.Append lgObjComm.CreateParameter("@ACCT_CD",	adVarXChar,	adParamInput, 20,Replace(sAcctCd, "'", "''"))
			lgObjComm.Parameters.Append lgObjComm.CreateParameter("@MINOR_CD",	adVarXChar,	adParamInput, 3,Replace(sMinorCd, "'", "''"))
		End If
		lgObjComm.Parameters.Append lgObjComm.CreateParameter("@SHEETCNT",  adSmallInt, adParamInput,, 100)	
		lgObjComm.Parameters.Append lgObjComm.CreateParameter("@NEXTKEY",	adVarXChar,	adParamInput, 15,Replace(sNextKey, "'", "''"))
		lgObjComm.Parameters.Append lgObjComm.CreateParameter("@DEBUG",  adSmallInt, adParamInput,, 0)	' -- isqlw ������ ����ϴ� ������ڵ� 
		    
        Set oRs = lgObjComm.Execute
        
    End With
    
    If Instr( Err.Description , "B_MESSAGE") > 0 Then
		If HandleBMessageError(vbObjectError, Err.Description, "", "") = True Then
			Exit Sub
		End If
	Else
		If CheckSYSTEMError(Err, True) = True Then	
			Exit Sub
		End If
	End If    
    ' -- ���� ���ڵ���� �� 2-3�� 
    If Not oRs.EOF Then
		If sNextKey = "" Then	' -- ��� ���ڵ�� ���ԵǾ� �� 
			' --- ��� ���(�����) ----			
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
			Set oRs = oRs.NextRecordSet()	' -- ����(����Ÿ) ���ڵ������ ���� 
		End If
		' -- �׷���� �÷� ������ ���ʷ� ����Ÿ��Ʈ�� �����Ѵ�.(�����)
		arrRows		= oRs.GetRows()
		
		Set oRs = oRs.NextRecordSet()	' -- ����(����Ÿ) ���ڵ������ ���� 
		
		Dim arrDataSet, arrTemp, iLngGrRowCnt, iLngGrColCnt, iMaxCols
		arrDataSet = oRs.GetRows()

		iLngGrRowCnt= UBound(arrRows, 2)				' �׷���� ��� 
		iLngGrColCnt = UBound(arrRows, 1)				' �׷���� ���� 
		
		iLngColCnt	= CLng(arrRows(UBound(arrRows, 1), 0))					' �÷��� 
		
		If sGrid = "A" Then
			iMaxCols	= 9	+ iLngColCnt + 2	' �׷���� �÷���� 
		Else
			iMaxCols	= iLngGrColCnt + iLngColCnt 
		End If
	' -- ����Ÿ���� ���ʷ� �迭�� �籸���Ѵ�.
		ReDim arrTemp(iLngColCnt, iLngGrRowCnt)		
		' -- ����Ÿ�� ��� 
		iLngRowCnt	= UBound(arrDataSet, 2)
		
		For iLngRow = 0 To 	iLngRowCnt
			arrTemp(CLng(arrDataSet(1, iLngRow)), CLng(arrDataSet(0, iLngRow))-1) = arrDataSet(2, iLngRow)	' -- ��, ��, ��(��)
		Next		
		Set oRs = Nothing		
		sRowSeq = arrRows(UBound(arrRows, 1) -1 , iLngGrRowCnt)		' -- ����Ű�� (�����)

		' -- �׷���� ������ ���� 
		For iLngRow = 0 To 	iLngGrRowCnt
			If sGrid = "A" Then
				iStrData = iStrData & Chr(11) & ConvLang(ConvSPChars(Trim(arrRows(0, iLngRow))), arrRows(iLngGrColCnt-2, iLngRow))	' -- cost_cd
				iStrData = iStrData & Chr(11) & ConvLang(ConvSPChars(Trim(arrRows(1, iLngRow))), "0")					' -- cost_nm
				iStrData = iStrData & Chr(11) & ConvLang(ConvSPChars(Trim(arrRows(2, iLngRow))), arrRows(iLngGrColCnt-2, iLngRow))	' -- COST_ELMT_CD
				iStrData = iStrData & Chr(11) & ConvLang(ConvSPChars(Trim(arrRows(3, iLngRow))), "0")					' -- COST_ELMT_NM
				iStrData = iStrData & Chr(11) & ConvLang(ConvSPChars(Trim(arrRows(4, iLngRow))), arrRows(iLngGrColCnt-2, iLngRow))	' -- ACCT_CD
				iStrData = iStrData & Chr(11) & ConvLang(ConvSPChars(Trim(arrRows(5, iLngRow))), "0")						' -- ACCT_NM
				iStrData = iStrData & Chr(11) & ConvLang(ConvSPChars(Trim(arrRows(6, iLngRow))), "0")					' -- MINOR_NM
				iStrData = iStrData & Chr(11) & ConvLang(ConvSPChars(Trim(arrRows(7, iLngRow))), "0")					' -- MINOR_NM
				iStrData = iStrData & Chr(11) & ConvLang(ConvSPChars(Trim(arrRows(8, iLngRow))), "0")					' -- MINOR_NM

				' -- ����Ÿ�׸��� ���(������)
				For iLngCol = 1 To iLngColCnt 
					iStrData = iStrData & Chr(11) & arrTemp(iLngCol, iLngRow)
				Next
				If Trim(arrRows(9, iLngRow)) <> "0" Then	' -- �Ұ����� ������ 
					sGrpTxt = sGrpTxt & arrRows(9, iLngRow) & gColSep & arrRows(10, iLngRow) & gRowSep		' -- �Ұ豸��|���ȣ(�迭�� ��ġ�� ����)
				End If				
				iStrData = iStrData & Chr(11) & arrRows(9, iLngRow)
				iStrData = iStrData & Chr(11) & arrRows(10, iLngRow)
			Else	 ' -- B �� ��� 
				For iLngCol = 0 To iLngGrColCnt -1
					iStrData = iStrData & Chr(11) & ConvSPChars(Trim(arrRows(iLngCol, iLngRow)))	
				Next
				' -- ����Ÿ�׸��� ���(������)
				For iLngCol = 1 To iLngColCnt 
					iStrData = iStrData & Chr(11) & ConvSPChars(Trim(arrTemp(iLngCol, iLngRow)))
				Next
				iStrData = iStrData & Chr(11) & arrRows(iLngGrColCnt, iLngRow)
			End If
			iStrData = iStrData & Chr(12)
		Next				
		Response.Write " <Script Language=vbscript>	                        " & vbCr
		Response.Write " With parent                                        " & vbCr		
		If sGrid = "A" Then
			Response.Write "	.InitSpreadSheet(" & iMaxCols & ")					" & vbCr 			 		
			Response.Write "	.frm1.vspdData.ReDraw = False					" & vbCr 			 
			Response.Write "	.ggoSpread.Source = .frm1.vspdData              " & vbCr 			 
			Response.Write "	.ggoSpread.SSShowData """ & iStrData		       & """" & vbCr
			Response.Write "	.frm1.vspdData.ReDraw = True					" & vbCr 			 
			Response.Write "	.lgStrPrevKey = """ & sRowSeq & """" & vbCr
			Response.Write "	.frm1.hSTART_DT.value = """ & sStartDt & """" & vbCr 	
			Response.Write "	.frm1.hEND_DT.value = """ & sEndDt & """" & vbCr 	
			Response.Write "	.frm1.hTYPE.value = """ & sType & """" & vbCr 	 				
			If sNextKey = "" Then
				Response.Write  "   Call Parent.SetGridHead(""" & sTxt & """)" & vbCr
			End If
			If sGrpTxt <> "" Then
				Response.Write  "   Call Parent.SetQuerySpreadColor(""" & sGrpTxt & """)" & vbCr
			End If			
			Response.Write "	.frm1.vspdData.MaxCols = """ & (iMaxCols) & """" & vbCr 	 	
			Response.Write  "   Call Parent.DbQueryOk()		" & vbCr
		Else
			Response.Write  "   Call Parent.InitSpreadSheet2(" & iMaxCols & ")		" & vbCr			
			Response.Write "	.frm1.vspdData2.ReDraw = False					" & vbCr 			 
			Response.Write "	.ggoSpread.Source = .frm1.vspdData2              " & vbCr 			 
			Response.Write "	.ggoSpread.SSShowData """ & iStrData		       & """" & vbCr
			Response.Write "	.frm1.vspdData2.ReDraw = True					" & vbCr 			 
			If sNextKey = "" Then
				Response.Write  "   Call Parent.SetGridHead2(" & (iMaxCols) & ", """ & sTxt & """)" & vbCr
			End If
			Response.Write  "   Call Parent.DbQueryOk2()		" & vbCr
		End If		
		Response.Write " End With                                        " & vbCr
		Response.Write  " </Script>                  " & vbCr
    Else
		If sGrid = "A" And sNextKey= ""  Then
			Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '�� : No data is found. 
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

End Sub	

Function ConvLang(Byval pLang, Byval GroupNo)
	Dim pTmp
	
	If GroupNo <> "0" Then
		pTmp = Replace(pLang , "%1", "�հ�")
		pTmp = Replace(pTmp , "%2", "������Һ��Ұ�")
		pTmp = Replace(pTmp , "%3", "�������Ұ�")
	Else
		pTmp = pLang
	End If
	ConvLang = pTmp
End Function


Function ConvLang2(Byval pLang, Byval GroupNo)
	Dim pTmp
	
	If GroupNo <> "0" Then
		pTmp = Replace(pLang , "%1", "�հ�")
		pTmp = Replace(pTmp , "%2", "�����׸� �Ұ�")
	Else
		pTmp = pLang
	End If
	ConvLang2 = pTmp
End Function
%>

