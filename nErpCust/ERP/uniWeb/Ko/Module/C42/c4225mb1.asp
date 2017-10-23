<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% Option Explicit%>
<% session.CodePage=949 %>

<%
'**********************************************************************************************
'*  1. Module Name          : ���� 
'*  2. Function Name        : ��ο��DATA��ȸ 
'*  3. Program ID           : c4225ma1
'*  4. Program Name         : ��ο��DATA��ȸ 
'*  5. Program Desc         :
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2005/11/24
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
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->

<%

	Call LoadBasisGlobalInf()
	Call loadInfTB19029B("Q", "C", "NOCOOKIE","MB")

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
	Dim sDstbFctrCd, sCostCd, sWcCd, sTypeFlag, arrTmp, TmpBuffer()

	Dim C_SHEET_MAX_CONT 
	
	C_SHEET_MAX_CONT = 1000
	
'    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status

	' -- �����ؾ��� ��ȸ���� (MA���� �����ִ�)
	Dim sStartDt, sEndDt
	
	sStartDt	= Request("txtStartDt")	
	sNextKey	= Request("lgStrPrevKey")
	
	sDstbFctrCd	= Request("txtDSTB_FCTR_CD")	
	sCostCd		= Request("txtCOST_CD")	
	sWcCd		= Request("txtWC_CD")	
	sTypeFlag	= Request("TYPE_FLAG")
		
	If sStartDt = "" And sDstbFctrCd = "" And sCostCd = "" And sWcCd = "" Then
		Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '�� : No data is found. 
		Exit Sub
	End If
	
	If sDstbFctrCd = "" Then sDstbFctrCd = "%"
	If sCostCd = "" Then sCostCd = "%"
	If sWcCd = "" Then sWcCd = "%"

	' -- sNextKey ������� 
	If Instr(1, sNextKey, "*") > 0 Then
		arrTmp = Split(sNextKey, gColSep)
		sNextKey = arrTmp(0)
		C_SHEET_MAX_CONT = 32000
	End If

    With lgObjComm
		.CommandTimeout = 0
		.CommandText = "dbo.usp_C_C4225MA1_TYPE" & sTypeFlag		' --  �����ؾ��� SP �� 
		.CommandType = adCmdStoredProc

		lgObjComm.Parameters.Append lgObjComm.CreateParameter("RETURN_VALUE",  adInteger,adParamReturnValue)	' -- No ���� 

		' -- �����ؾ��� ��ȸ���� �Ķ��Ÿ�� 
		lgObjComm.Parameters.Append lgObjComm.CreateParameter("@START_DT",	adVarXChar,	adParamInput, 6,Replace(sStartDt, "'", "''"))
		
		lgObjComm.Parameters.Append lgObjComm.CreateParameter("@DSTB_FCTR_CD",	adVarXChar,	adParamInput, 3,Replace(sDstbFctrCd, "'", "''"))
		lgObjComm.Parameters.Append lgObjComm.CreateParameter("@COST_CD",	adVarXChar,	adParamInput, 10,Replace(sCostCd, "'", "''"))
		lgObjComm.Parameters.Append lgObjComm.CreateParameter("@WC_CD",	adVarXChar,	adParamInput, 7,Replace(sWcCd, "'", "''"))
		
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
		
    
    ' -- ���� ���ڵ���� �� 2-3�� 
    If Not oRs.EOF Then

		Response.Write " <Script Language=vbscript>	                        " & vbCr
		Response.Write " With parent                                        " & vbCr

		Response.Write "	.frm1.hSTART_DT.value = """ & sStartDt & """" & vbCr 	
		Response.Write "	.frm1.hDSTB_FCTR_CD.value = """ & sDstbFctrCd & """" & vbCr
		Response.Write "	.frm1.hCOST_CD.value = """ & sCostCd & """" & vbCr
		Response.Write "	.frm1.hWC_CD.value = """ & sWcCd & """" & vbCr
		
		If sTypeFlag = "1" Then	' -- All �Ǵ� 1�� �׸��� 
		
			' -- 1�� �׸��� ��ȸ 
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
				' ----------------------------------------
			End If

			' -- �׷���� �÷� ������ ���ʷ� ����Ÿ��Ʈ�� �����Ѵ�.(�����)
			arrRows		= oRs.GetRows()		
			Set oRs = oRs.NextRecordSet()	' -- ����(����Ÿ) ���ڵ������ ���� 
		
			Dim arrDataSet, arrTemp, iLngGrRowCnt, iLngGrColCnt, iMaxCols
			arrDataSet = oRs.GetRows()

			Set oRs = Nothing

			iLngGrRowCnt= UBound(arrRows, 2)				' �׷���� ��� 
			iLngGrColCnt = UBound(arrRows, 1)				' �׷���� ���� 
		
			iLngColCnt	= CLng(arrRows(UBound(arrRows, 1), 0))					' �÷��� 

			iMaxCols	= 4	+ iLngColCnt 	' �׷���� �÷���� 
		
			' -- ����Ÿ���� ���ʷ� �迭�� �籸���Ѵ�.
			ReDim arrTemp(iLngColCnt, iLngGrRowCnt)
		
			' -- ����Ÿ�� ��� 
			iLngRowCnt	= UBound(arrDataSet, 2)
		
			For iLngRow = 0 To 	iLngRowCnt
				arrTemp((CLng(arrDataSet(1, iLngRow))-1), CLng(arrDataSet(0, iLngRow))-1) = arrDataSet(2, iLngRow)	' -- ��, ��, ��(��)
			Next
		
			' ----------------------------------------------------------
			sRowSeq = arrRows(UBound(arrRows, 1) -1 , iLngGrRowCnt)		' -- ����Ű�� (�����)

			' -- ��ȸ Row�� �ִ������ ��ġ�Ҷ��� ��������Ÿ ������ 
			If CInt(sRowSeq) < C_SHEET_MAX_CONT Then
				sRowSeq = ""
			End If

			Redim TmpBuffer(iLngGrRowCnt)
			
			' -- �׷���� ������ ���� 
			For iLngRow = 0 To 	iLngGrRowCnt						
				iStrData = Chr(11) & ConvSPChars(Trim(arrRows(0, iLngRow)))
				iStrData = iStrData & Chr(11) & ConvSPChars(Trim(arrRows(1, iLngRow)))
				iStrData = iStrData & Chr(11) & ConvSPChars(Trim(arrRows(2, iLngRow)))
				' -- ����Ÿ�׸��� ���(������)
				For iLngCol = 0 To iLngColCnt -1
					iStrData = iStrData & Chr(11) & UniConvNumberDBToCompany(arrTemp(iLngCol, iLngRow), ggAmtOfMoney.DecPoint, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit, 0)
				Next
				iStrData = iStrData & Chr(11) & ConvSPChars(Trim(arrRows(3, iLngRow)))	' -- Row_Seq
				iStrData = iStrData & Chr(12)
				
				TmpBuffer(iLngRow) = iStrData
			Next			
			iStrData = Join(TmpBuffer, "")

			If sNextKey = "" Then
				Response.Write  "   Call Parent.InitSpreadSheet(" & iMaxCols & ")" & vbCr
				Response.Write  "   Call Parent.SetGridHead(""" & sTxt & """)" & vbCr
			End If
			Response.Write "	.lgStrPrevKey = """ & sRowSeq & """" & vbCr
			Response.Write "	.frm1.vspdData.ReDraw = False					" & vbCr 			 
			Response.Write "	.ggoSpread.Source = .frm1.vspdData              " & vbCr 			 
			Response.Write "	.ggoSpread.SSShowData """ & iStrData		       & """" & vbCr
			Response.Write "	.frm1.vspdData.ReDraw = True					" & vbCr 				
		End If

		If sTypeFlag = "2" Then	' -- All �Ǵ� 2�� �׸���		
			' -- 1�� �׸��� ��ȸ 
			If sNextKey = "" Then	' -- ��� ���ڵ�� ���ԵǾ� �� 
				' --- ��� ���(�����) ----			
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

			arrDataSet = oRs.GetRows()

			Set oRs = Nothing

			iLngGrRowCnt= UBound(arrRows, 2)				' �׷���� ��� 
			iLngGrColCnt = UBound(arrRows, 1)				' �׷���� ���� 
		
			iLngColCnt	= CLng(arrRows(UBound(arrRows, 1), 0))					' �÷��� 

			iMaxCols	= 6	+ iLngColCnt	' �׷���� �÷���� 
		
			' -- ����Ÿ���� ���ʷ� �迭�� �籸���Ѵ�.
			ReDim arrTemp(iLngColCnt, iLngGrRowCnt)
		
			' -- ����Ÿ�� ��� 
			iLngRowCnt	= UBound(arrDataSet, 2)
		
			For iLngRow = 0 To 	iLngRowCnt
				arrTemp((CLng(arrDataSet(1, iLngRow))-1), CLng(arrDataSet(0, iLngRow))-1) = arrDataSet(2, iLngRow)	' -- ��, ��, ��(��)
			Next
		
			' ----------------------------------------------------------
			sRowSeq = arrRows(UBound(arrRows, 1) -1 , iLngGrRowCnt)		' -- ����Ű�� (�����)

			' -- ��ȸ Row�� �ִ������ ��ġ�Ҷ��� ��������Ÿ ������ 
			If CInt(sRowSeq) < C_SHEET_MAX_CONT Then
				sRowSeq = ""
			End If

			Redim TmpBuffer(iLngGrRowCnt)
			
			' -- �׷���� ������ ���� 
			For iLngRow = 0 To 	iLngGrRowCnt						
				iStrData = Chr(11) & ConvSPChars(Trim(arrRows(0, iLngRow)))
				iStrData = iStrData & Chr(11) & ConvSPChars(Trim(arrRows(1, iLngRow)))
				iStrData = iStrData & Chr(11) & ConvSPChars(Trim(arrRows(2, iLngRow)))
				iStrData = iStrData & Chr(11) & ConvSPChars(Trim(arrRows(3, iLngRow)))
				iStrData = iStrData & Chr(11) & ConvSPChars(Trim(arrRows(4, iLngRow)))
				' -- ����Ÿ�׸��� ���(������)
				For iLngCol = 0 To iLngColCnt -1
					iStrData = iStrData & Chr(11) & arrTemp(iLngCol, iLngRow)
				Next
				iStrData = iStrData & Chr(11) & ConvSPChars(Trim(arrRows(5, iLngRow)))	' -- Row_Seq
				iStrData = iStrData & Chr(12)
				
				TmpBuffer(iLngRow) = iStrData
			Next
			iStrData = Join(TmpBuffer, "")

			If sNextKey = "" Then
				Response.Write  "   Call Parent.InitSpreadSheet2(" & iMaxCols & ")" & vbCr
				Response.Write  "   Call Parent.SetGridHead2(""" & sTxt & """)" & vbCr
			End If
			Response.Write "	.lgStrPrevKey2 = """ & sRowSeq & """" & vbCr
			Response.Write "	.frm1.vspdData2.ReDraw = False					" & vbCr 			 
			Response.Write "	.ggoSpread.Source = .frm1.vspdData2              " & vbCr 			 
			Response.Write "	.ggoSpread.SSShowData """ & iStrData		       & """" & vbCr
			Response.Write "	.frm1.vspdData2.ReDraw = True					" & vbCr 					
		End If
		' ------------------------------------------------------------------
		If sTypeFlag = "3" Then	' -- All �Ǵ� 2�� �׸��� 
			' -- �׷���� �÷� ������ ���ʷ� ����Ÿ��Ʈ�� �����Ѵ�.(�����)
			arrRows		= oRs.GetRows()
		
			Set oRs = Nothing

			iLngGrRowCnt= UBound(arrRows, 2)				' �׷���� ��� 
			iLngGrColCnt = UBound(arrRows, 1)				' �׷���� ����		
			' ----------------------------------------------------------
			sRowSeq = arrRows(UBound(arrRows, 1)  , iLngGrRowCnt)		' -- ����Ű�� (�����)

			' -- ��ȸ Row�� �ִ������ ��ġ�Ҷ��� ��������Ÿ ������ 
			If CInt(sRowSeq) < C_SHEET_MAX_CONT Then
				sRowSeq = ""
			End If

			Redim TmpBuffer(iLngGrRowCnt)
			
			' -- �׷���� ������ ���� 
			For iLngRow = 0 To 	iLngGrRowCnt						
				iStrData = Chr(11) & ConvSPChars(Trim(arrRows(0, iLngRow)))
				iStrData = iStrData & Chr(11) & ConvSPChars(Trim(arrRows(1, iLngRow)))
				iStrData = iStrData & Chr(11) & ConvSPChars(Trim(arrRows(2, iLngRow)))
				iStrData = iStrData & Chr(11) & ConvSPChars(Trim(arrRows(3, iLngRow)))
				iStrData = iStrData & Chr(11) & ConvSPChars(Trim(arrRows(4, iLngRow)))
				iStrData = iStrData & Chr(11) & ConvSPChars(Trim(arrRows(5, iLngRow)))
				iStrData = iStrData & Chr(11) & ConvSPChars(Trim(arrRows(6, iLngRow)))
				iStrData = iStrData & Chr(11) & ConvSPChars(Trim(arrRows(7, iLngRow)))
				iStrData = iStrData & Chr(11) & Trim(arrRows(8, iLngRow))				
				iStrData = iStrData & Chr(11) & ConvSPChars(Trim(arrRows(9, iLngRow)))	' -- Row_Seq

				iStrData = iStrData & Chr(12)				
				TmpBuffer(iLngRow) = iStrData
			Next			
			iStrData = Join(TmpBuffer, "")

			If sNextKey = "" Then
				Response.Write  "   Call Parent.InitSpreadSheet3()" & vbCr
			End If
			Response.Write "	.lgStrPrevKey3 = """ & sRowSeq & """" & vbCr
			Response.Write "	.frm1.vspdData3.ReDraw = False					" & vbCr 			 
			Response.Write "	.ggoSpread.Source = .frm1.vspdData3              " & vbCr 			 
			Response.Write "	.ggoSpread.SSShowData """ & iStrData		       & """" & vbCr
			Response.Write "	.frm1.vspdData3.ReDraw = True					" & vbCr 	
		End If
		' -------------------------------------------------------------------------
		Response.Write  "   Call Parent.DbQueryOk()		" & vbCr
		Response.Write " End With                                        " & vbCr
		Response.Write " </Script>	                        " & vbCr
    
    ElseIf sNextKey = "" Then
       Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '�� : No data is found. 
    End If

End Sub	

%>

