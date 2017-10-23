<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% Option Explicit%>
<% session.CodePage=949 %>

<%
'**********************************************************************************************
'*  1. Module Name          : ����
'*  2. Function Name        : ���������ȸ(����) 
'*  3. Program ID           : c4223mb1_ko441.asp
'*  4. Program Name         : ���������ȸ(����) 
'*  5. Program Desc         :
'*  6. Comproxy List        : ADO
'*  7. Modified date(First) : 2005/09/27
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     :choe0tae 
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
	Dim sPlantCd, sCostCd, sItemCd, sWCCd, sType, sSendAmt, sRecvAmt, arrTmp, TmpBuffer

	' -- ������ �극��ŷ 
	Dim C_SHEET_MAX_CONT 
	
	C_SHEET_MAX_CONT = 500	
    
    Err.Clear                                                                        '��: Clear Error status

	' -- �����ؾ��� ��ȸ���� (MA���� �����ִ�)
	Dim sStartDt, sEndDt
	
	sStartDt		= Request("txtStartDt")	
	sNextKey		= Request("lgStrPrevKey")
	
	sPlantCd		= Request("txtPLANT_CD")	
	sCostCd			= Request("txtCOST_CD")	
	sItemCd			= Request("txtITEM_CD")	
	sWCCd			= Request("txtWC_CD")	
	sType			= Request("rdoTYPE")	
		
	If sStartDt = "" Or sType = "" Then
		Call DisplayMsgBox("900015", vbInformation, "", "", I_MKSCRIPT)      '�� : No data is found. 
		Exit Sub
	End If
	
	If sPlantCd = "" Then sPlantCd = "%"
	If sCostCd = "" Then sCostCd = "%"
	If sItemCd = "" Then sItemCd = "%"
	If sWCCd = "" Then sWCCd = "%"

	' -- sNextKey ������� 
	If Instr(1, sNextKey, "*") > 0 Then
		arrTmp = Split(sNextKey, gColSep)
		sNextKey = arrTmp(0)
		C_SHEET_MAX_CONT = 32000
	End If

'Call ServerMesgBox("sType==>" & sType, vbInformation, I_MKSCRIPT)
	
    With lgObjComm
		.CommandTimeout = 0
		
		.CommandText = "dbo.usp_C_C4223MA1_TYPE1_" & sType & "_Ko441"	' --  �����ؾ��� SP �� 
	    	.CommandType = adCmdStoredProc
				
		lgObjComm.Parameters.Append lgObjComm.CreateParameter("RETURN_VALUE",  adInteger,adParamReturnValue)	' -- No ���� 

		' -- �����ؾ��� ��ȸ���� �Ķ��Ÿ�� 
		lgObjComm.Parameters.Append lgObjComm.CreateParameter("@START_DT",	adVarXChar,	adParamInput, 6,Replace(sStartDt, "'", "''"))
		
		lgObjComm.Parameters.Append lgObjComm.CreateParameter("@PLANT_CD",	adVarXChar,	adParamInput, 4,Replace(sPlantCd, "'", "''"))
		lgObjComm.Parameters.Append lgObjComm.CreateParameter("@COST_CD",	adVarXChar,	adParamInput, 10,Replace(sCostCd, "'", "''"))
		lgObjComm.Parameters.Append lgObjComm.CreateParameter("@ITEM_CD",	adVarXChar,	adParamInput, 20,Replace(sItemCd, "'", "''"))
		lgObjComm.Parameters.Append lgObjComm.CreateParameter("@WC_CD",	adVarXChar,	adParamInput, 20,Replace(sWCCd, "'", "''"))
		
		lgObjComm.Parameters.Append lgObjComm.CreateParameter("@SHEETCNT",  adSmallInt, adParamInput,, C_SHEET_MAX_CONT)	
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

	If oRs is Nothing Then
		Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '�� : No data is found. 
		Exit Sub
	End If 		
    
    ' -- ���� ���ڵ���� �� 2-3�� 
    If Not oRs.EOF Then

		If sNextKey = "" Then	' -- ��� ���ڵ�� ���ԵǾ� �� 

			' --- ��� ���(�����) ----			
			Dim arrColRow, i, j, ColHeadRowNo, sTxt2
			
			ColHeadRowNo = -1000
			arrColRow = oRs.GetRows()
			iLngRowCnt	= UBound(arrColRow, 2) 
			iLngColCnt	= UBound(arrColRow, 1) 

			If sType = "A" Then
				sTxt = sTxt & "�հ�" & gColSep
				sTxt = sTxt & "�հ�" & gColSep
				sTxt2 = sTxt2 & "�������" & gColSep
				sTxt2 = sTxt2 & "�⸻���" & gColSep
				i = 1
			Else
				i = 0
			End If

			For j = i To iLngRowCnt
				sTxt = sTxt & arrColRow(0, j) & gColSep
				sTxt = sTxt & arrColRow(1, j) & gColSep
				sTxt2 = sTxt2 & "�������" & gColSep
				sTxt2 = sTxt2 & "�⸻���" & gColSep
			Next	


			'2009.04.02 kbs change
			'sTxt  = sTxt  & "�������" & gColSep
			'sTxt2 = sTxt2 & "�������" & gColSep

								
			sTxt = sTxt & CStr(ColHeadRowNo) & gRowSep
			ColHeadRowNo = ColHeadRowNo + 1
			sTxt = sTxt & sTxt2 & CStr(ColHeadRowNo) & gRowSep
			
			Set oRs = oRs.NextRecordSet()	' -- ����(����Ÿ) ���ڵ������ ���� 
			' ----------------------------------------
		End If
		
		' -- �׷���� �÷� ������ ���ʷ� ����Ÿ��Ʈ�� �����Ѵ�.(�����)
		arrRows		= oRs.GetRows()
		
		Set oRs = oRs.NextRecordSet()	' -- ����(����Ÿ) ���ڵ������ ���� : ����, �ڽ�Ʈ��Ÿ, �ڽ��ͼ�Ÿ��, �׷��ȣ, Row����, Max Cols
		
		Dim arrDataSet, arrTemp, iLngGrRowCnt, iLngGrColCnt, iMaxCols
		arrDataSet = oRs.GetRows()

		iLngGrRowCnt = UBound(arrRows, 2)				' �׷���� ��� 
		iLngGrColCnt = UBound(arrRows, 1)				' �׷���� ���� 
		
		iLngColCnt	= CLng(arrRows(UBound(arrRows, 1), 0))	* 2		' �÷��� 
		
		If sType = "A" Then
			iMaxCols	= 6	+ iLngColCnt	' �׷���� �÷���� 
		Else
			iMaxCols	= 2	+ iLngColCnt	' �׷���� �÷���� 
		End If
		
		' -- ����Ÿ���� ���ʷ� �迭�� �籸���Ѵ�.
		ReDim arrTemp(iLngColCnt, iLngGrRowCnt)		
				
		' -- ����Ÿ�� ��� 
		iLngRowCnt	= UBound(arrDataSet, 2)
		
		For iLngRow = 0 To 	iLngRowCnt			
			arrTemp((CLng(arrDataSet(1, iLngRow))-1)*2  , CLng(arrDataSet(0, iLngRow))-1) = arrDataSet(2, iLngRow)	' -- ��, ��, ��(��)
			arrTemp((CLng(arrDataSet(1, iLngRow))-1)*2+1, CLng(arrDataSet(0, iLngRow))-1) = arrDataSet(3, iLngRow)	' -- ��, ��, ��(��)
			
			If Err.number <> 0 Then 
				Response.Write "iLngRow=" & iLngRow & ";iLngColCnt=" & iLngColCnt
				Response.End 
			End If
		Next
		
		Set oRs = Nothing
		
		sRowSeq = arrRows(UBound(arrRows, 1) -1 , iLngGrRowCnt)		' -- ����Ű�� (�����)

		' -- ��ȸ Row�� �ִ������ ��ġ�Ҷ��� ��������Ÿ ������ 
		If CInt(sRowSeq) < C_SHEET_MAX_CONT Then
			sRowSeq = ""
		End If

		' -- ���ڿ� ������ �迭�������� �� 
		Redim TmpBuffer(iLngGrRowCnt)
		
		' -- �׷���� ������ ���� 
		For iLngRow = 0 To 	iLngGrRowCnt
			
			iStrData = ""
				
			If sType = "A" Then
			
				iStrData = iStrData & Chr(11) & ConvLang(ConvSPChars(Trim(arrRows(0, iLngRow))), arrRows(iLngGrColCnt-2, iLngRow))	' -- PLANT_CD
				iStrData = iStrData & Chr(11) & ConvLang(ConvSPChars(Trim(arrRows(1, iLngRow))), arrRows(iLngGrColCnt-2, iLngRow))	' -- COST_CD
				iStrData = iStrData & Chr(11) & ConvSPChars(Trim(arrRows(2, iLngRow)))							' -- COST_NM
				iStrData = iStrData & Chr(11) & ConvLang(ConvSPChars(Trim(arrRows(3, iLngRow))), arrRows(iLngGrColCnt-2, iLngRow))	' -- ITEM_CD
				iStrData = iStrData & Chr(11) & ConvSPChars(Trim(arrRows(4, iLngRow)))							' -- ITEM_NM

				' -- ����Ÿ�׸��� ���(������)
				For iLngCol = 0 To iLngColCnt -1
					iStrData = iStrData & Chr(11) & arrTemp(iLngCol, iLngRow)
				Next

				If Trim(arrRows(5, iLngRow)) <> "0" Then	' -- �Ұ����� ������ 
					sGrpTxt = sGrpTxt & arrRows(5, iLngRow) & gColSep & arrRows(6, iLngRow) & gRowSep		' -- �Ұ豸��|���ȣ(�迭�� ��ġ�� ����)
				End If

				
				iStrData = iStrData & Chr(11) & arrRows(5, iLngRow)
				iStrData = iStrData & Chr(11) & arrRows(6, iLngRow)

			Else	 ' -- B �� ��� 
				iStrData = iStrData & Chr(11) & ConvSPChars(Trim(arrRows(0, iLngRow)))		' -- ORDER_NO
				' -- ����Ÿ�׸��� ���(������)
				For iLngCol = 0 To iLngColCnt 
					iStrData = iStrData & Chr(11) & arrTemp(iLngCol, iLngRow)
				Next
				iStrData = iStrData & Chr(11) & arrRows(1, iLngRow)

			End If
			iStrData = iStrData & Chr(12)
			
			TmpBuffer(iLngRow) = iStrData
		Next
		
		iStrData = Join(TmpBuffer, "")
				
		Response.Write " <Script Language=vbscript>	                        " & vbCr
		Response.Write " With parent                                        " & vbCr
		
		If sType = "A" Then

			If sNextKey = "" Then
				Response.Write "	Call Parent.InitSpreadSheet(" & iMaxCols & ")			" & vbCr 			 
			End If
			
			Response.Write "	.frm1.vspdData.ReDraw = False					" & vbCr 			 
			Response.Write "	.ggoSpread.Source = .frm1.vspdData              " & vbCr 			 
			Response.Write "	.ggoSpread.SSShowData """ & iStrData		       & """" & vbCr
			Response.Write "	.frm1.vspdData.ReDraw = True					" & vbCr 	

			If 	C_SHEET_MAX_CONT = 32000 Then
				Response.Write "	.lgStrPrevKey = ""*""" & vbCr
			Else
				Response.Write "	.lgStrPrevKey = """ & sRowSeq & """" & vbCr
			End If

			Response.Write "	.frm1.hSTART_DT.value = """ & sStartDt & """" & vbCr
			Response.Write "	.frm1.hPLANT_CD.value = """ & sPlantCd & """" & vbCr 	 	
			Response.Write "	.frm1.hCOST_CD.value = """ & sCostCd & """" & vbCr 	
			Response.Write "	.frm1.hITEM_CD.value = """ & sItemCd & """" & vbCr
			Response.Write "	.frm1.hWC_CD.value = """ & sWCCd & """" & vbCr
			Response.Write "	.frm1.hTYPE.value = """ & sType & """" & vbCr 	 	
			
			If sNextKey = "" Then
				Response.Write  "   Call Parent.SetGridHead(""" & sTxt & """)" & vbCr
			End If

			If sGrpTxt <> "" Then
				Response.Write  "   Call Parent.SetQuerySpreadColor(""" & sGrpTxt & """)" & vbCr
			End If
			Response.Write  "   Call Parent.DbQueryOk()		" & vbCr
		Else
			Response.Write "	Call Parent.InitSpreadSheet2(" & iMaxCols & ")			" & vbCr 		
			Response.Write "	.frm1.vspdData2.ReDraw = False					" & vbCr 			 
			Response.Write "	.ggoSpread.Source = .frm1.vspdData2              " & vbCr 			 
			Response.Write "	.ggoSpread.SSShowData """ & iStrData		       & """" & vbCr
			Response.Write "	.frm1.vspdData2.ReDraw = True					" & vbCr 			 
			
			Response.Write "	.lgStrPrevKey2 = """ & sRowSeq & """" & vbCr

			If sNextKey = "" Then
				Response.Write  "   Call Parent.SetGridHead2(""" & sTxt & """)" & vbCr
			End If

			Response.Write  "   Call Parent.DbQueryOk2()		" & vbCr
		End If
		
		Response.Write " End With                                        " & vbCr
		Response.Write  " </Script>                  " & vbCr

    ElseIf sType = "A" And sNextKey = "" Then
       Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '�� : No data is found. 
    End If

End Sub	

Function ConvLang(Byval pLang, Byval GroupNo)
	Dim pTmp
	
	If GroupNo <> "0" Then
		pTmp = Replace(pLang , "%1", "�հ�")
		pTmp = Replace(pTmp , "%2", "C/C�Ұ�")
		pTmp = Replace(pTmp , "%3", "ǰ��Ұ�")
	Else
		pTmp = pLang
	End If
	ConvLang = pTmp
End Function

%>

