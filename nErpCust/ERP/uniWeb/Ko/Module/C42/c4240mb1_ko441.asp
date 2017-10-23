<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% Option Explicit%>
<% session.CodePage=949 %>

<%
'**********************************************************************************************
'*  1. Module Name          : cost 
'*  2. Function Name        : ����CC��ȸ�谡��������ȸ(S)
'*  3. Program ID           : c4214ma1_ko441.asp
'*  4. Program Name         : ����CC��ȸ�谡��������ȸ(S)
'*  5. Program Desc         :
'*  6. Comproxy List        : ADO
'*  7. Modified date(First) : 2009-08-24
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : han ki hong
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
	Dim sBizAreaCd, sFromAcctCd, sToAcctCd, sItemCd, sType, sGrid, sCostCd
	
    'On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status

	Const C_MAX_COUNT = 500
	' -- �����ؾ��� ��ȸ���� (MA���� �����ִ�)
	Dim sStartDt, sEndDt
			
	sStartDt	= Request("txtStartDt")	
	sEndDt		= Request("txtEndDt")
	sNextKey	= Request("lgStrPrevKey")
	
	sBizAreaCd	= Request("txtBizAreaCd")	
	sFromAcctCd	= Request("txtFromAcctCd")
	sToAcctCd	= Request("txtToAcctCd")		
	sCostCd		= Request("txtCostCd")

	sType		= Request("rdoType")	
	sGrid		= Request("txtGrid")	
		
	If sStartDt = "" And sEndDt = ""  And sBizAreaCd = "" And sFromAcctCd = "" And sGrid = "" Then
		Call DisplayMsgBox("900015", vbInformation, "", "", I_MKSCRIPT)      '�� : No data is found. 
		Exit Sub
	End If
	
	If sBizAreaCd = "" Then sBizAreaCd = "%"
	If sToAcctCd = "" Then sToAcctCd = "zzzzzzzzzzzzzzzzzzz"
	If sCostCd = "" Then sCostCd = "%"
	If Instr(1, sCostCd, "%") = 0 Then sCostCd	= sCostCd & "%"

	'Response.Write "sStartDt=" & sStartDt & vbcrlf
	'Response.Write "sEndDt=" & sEndDt & vbcrlf
	'Response.Write "sNextKey=" & sNextKey & vbcrlf
	
	
    With lgObjComm
		.CommandTimeout = 0
		.CommandText = "dbo.usp_C_C4240MA1_TYPE" & sType & "_" & sGrid		' --  �����ؾ��� SP ��
	    .CommandType = adCmdStoredProc

		lgObjComm.Parameters.Append lgObjComm.CreateParameter("RETURN_VALUE",  adInteger,adParamReturnValue)	' -- No ����

		' -- �����ؾ��� ��ȸ���� �Ķ��Ÿ��
		lgObjComm.Parameters.Append lgObjComm.CreateParameter("@START_DT",	adVarXChar,	adParamInput, 6,Replace(sStartDt, "'", "''"))
		lgObjComm.Parameters.Append lgObjComm.CreateParameter("@END_DT",	adVarXChar,	adParamInput, 6,Replace(sEndDt, "'", "''"))
		
		lgObjComm.Parameters.Append lgObjComm.CreateParameter("@BIZ_AREA_CD",	adVarXChar,	adParamInput, 4,Replace(sBizAreaCd, "'", "''"))
		lgObjComm.Parameters.Append lgObjComm.CreateParameter("@COST_CD",	adVarXChar,	adParamInput, 10,Replace(sCostCd, "'", "''"))
		
		lgObjComm.Parameters.Append lgObjComm.CreateParameter("@FROM_ACCT_CD",	adVarXChar,	adParamInput, 20,Replace(sFromAcctCd, "'", "''"))
		lgObjComm.Parameters.Append lgObjComm.CreateParameter("@TO_ACCT_CD",	adVarXChar,	adParamInput, 20,Replace(sToAcctCd, "'", "''"))
		
		lgObjComm.Parameters.Append lgObjComm.CreateParameter("@SHEETCNT",  adSmallInt, adParamInput,, 500)	
		lgObjComm.Parameters.Append lgObjComm.CreateParameter("@NEXTKEY",	adVarXChar,	adParamInput, 15,Replace(sNextKey, "'", "''"))
		lgObjComm.Parameters.Append lgObjComm.CreateParameter("@DEBUG",  adSmallInt, adParamInput,, 0)	' -- isqlw ������ ����ϴ� ������ڵ�
		    
		
        Set oRs = lgObjComm.Execute
        
        
    End With
    
	
	
	Response.Write "Err=" & Err.Description
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
    If Not oRs.Eof Then

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
			
			sTxt	= Replace(sTxt, "%3", "�հ�")
			sTxt	= Replace(sTxt, "%4", "���庰 �հ�")
			sTxt	= Replace(sTxt, "%5", "Cost Center")
					
			Set oRs = oRs.NextRecordSet()	' -- ����(����Ÿ) ���ڵ������ ����
			' ----------------------------------------
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
			iMaxCols	= 4	+ iLngColCnt	' �׷���� �÷����
		Else
			iMaxCols	= 4	+ iLngColCnt	' �׷���� �÷����
		End If
		
		
	
		' -- ����Ÿ���� ���ʷ� �迭�� �籸���Ѵ�.
		ReDim arrTemp(iLngColCnt, iLngGrRowCnt)
		
		' -- ����Ÿ�� ���
		iLngRowCnt	= UBound(arrDataSet, 2)
		
		For iLngRow = 0 To 	iLngRowCnt
			arrTemp(CLng(arrDataSet(1, iLngRow)), CLng(arrDataSet(0, iLngRow))-1) = arrDataSet(2, iLngRow)	' -- ��, ��, ��(��)
		Next
		
		Set oRs = Nothing
		
		' ----------------------------------------------------------
		If iLngGrRowCnt = 0 Then 
		
			If sNextKey= ""  Then
				Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '�� : No data is found. 
			End If

			Exit Sub
		End If

		sRowSeq = arrRows(UBound(arrRows, 1) -1 , iLngGrRowCnt)		' -- ����Ű�� (�����)

		
		If iLngGrRowCnt < C_MAX_COUNT -1 Then sRowSeq = ""

		' -- �׷���� ������ ����
		For iLngRow = 0 To 	iLngGrRowCnt
				
			iStrData = iStrData & Chr(11) & ConvLang(ConvSPChars(Trim(arrRows(0, iLngRow))), arrRows(iLngGrColCnt-2, iLngRow))								' -- PLANT_CD
			iStrData = iStrData & Chr(11) & ConvLang(ConvSPChars(Trim(arrRows(1, iLngRow))), arrRows(iLngGrColCnt-2, iLngRow))	' -- COST_ELMT_CD
			iStrData = iStrData & Chr(11) & ConvLang(ConvSPChars(Trim(arrRows(2, iLngRow))), "0")					' -- COST_ELMT_NM
			'iStrData = iStrData & Chr(11) & ConvLang(ConvSPChars(Trim(arrRows(3, iLngRow))), arrRows(iLngGrColCnt-2, iLngRow))	' -- ACCT_CD
			'iStrData = iStrData & Chr(11) & ConvLang(ConvSPChars(Trim(arrRows(4, iLngRow))), "0")						' -- ACCT_NM
			'iStrData = iStrData & Chr(11) & ConvLang(ConvSPChars(Trim(arrRows(5, iLngRow))), "0")					' -- MINOR_NM

			' -- ����Ÿ�׸��� ���(������)
			For iLngCol = 1 To iLngColCnt 
				iStrData = iStrData & Chr(11) & arrTemp(iLngCol, iLngRow)
			Next

			If Trim(arrRows(3, iLngRow)) <> "0" Then	' -- �Ұ����� ������
				sGrpTxt = sGrpTxt & arrRows(3, iLngRow) & gColSep & arrRows(4, iLngRow) & gRowSep		' -- �Ұ豸��|���ȣ(�迭�� ��ġ�� ����)
			End If
				
			'iStrData = iStrData & Chr(11) & arrRows(6, iLngRow)
			iStrData = iStrData & Chr(11) & arrRows(4, iLngRow)
										
			'iStrData = iStrData & Chr(11) & iLngRow+1
			iStrData = iStrData & Chr(12)
		Next
				
					
		Response.Write " <Script Language=vbscript>	                        " & vbCr
		Response.Write " With parent                                        " & vbCr

		Response.Write "	.frm1.hStartDt.value = """ & sStartDt & """" & vbCr 	
		Response.Write "	.frm1.hEndDt.value = """ & sEndDt & """" & vbCr 	
		Response.Write "	.frm1.hType.value = """ & sType & """" & vbCr 	 	
		Response.Write "	.frm1.hCostCd.value = """ & sCostCd & """" & vbCr 	 	
		Response.Write "	.frm1.hBizAreaCd.value = """ & sBizAreaCd & """" & vbCr 	 	
		Response.Write "	.frm1.hFromAcctCd.value = """ & sFromAcctCd & """" & vbCr 	 	
		Response.Write "	.frm1.hToAcctCd.value = """ & sToAcctCd & """" & vbCr 	 	

		'Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)
		
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
       Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '�� : No data is found. 
    End If

End Sub	

Function ConvLang(Byval pLang, Byval GroupNo)
	Dim pTmp
	
	If GroupNo <> "0" Then
		pTmp = Replace(pLang , "%1", "�հ�")
		pTmp = Replace(pTmp , "%2", "��")
	Else
		pTmp = pLang
	End If
	ConvLang = pTmp
End Function


Function ConvLang2(Byval pLang, Byval GroupNo)
	Dim pTmp
	
	If GroupNo <> "0" Then
		pTmp = Replace(pLang , "%1", "�հ�")
		pTmp = Replace(pTmp , "%2", "��")
	Else
		pTmp = pLang
	End If
	ConvLang2 = pTmp
End Function
%>
