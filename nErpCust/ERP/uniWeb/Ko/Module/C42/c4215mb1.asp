<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% Option Explicit%>
<% session.CodePage=949 %>

<%
'**********************************************************************************************
'*  1. Module Name          : ���� 
'*  2. Function Name        : ��Һ�����������ȸ 
'*  3. Program ID           : b1256mb1
'*  4. Program Name         : ��Һ�����������ȸ 
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
	Dim sPlantCd, sCostCd, sItemAcct, sItemGroupCd, sItemCd, sType, sGrid, TmpBuffer()
	
    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status

	' -- �����ؾ��� ��ȸ���� (MA���� �����ִ�)
	Dim sStartDt, sEndDt
	
	sStartDt		= Request("txtStartDt")	

	sNextKey		= Request("lgStrPrevKey")
	
	sPlantCd		= Request("txtPLANT_CD")	
	sCostCd			= Request("txtCOST_CD")	
	sItemAcct		= Request("txtITEM_ACCT")	
	sItemGroupCd	= Request("txtITEM_GROUP_CD")
	sItemCd			= Request("txtITEM_CD")		
		
	If sStartDt = "" And (sPlantCd = ""  Or sCostCd = "" Or sItemAcct = "" Or sItemGroupCd = "" Or sItemCd = "") Then
		Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '�� : No data is found. 
		Exit Sub
	End If
	
	If sPlantCd = ""	Then sPlantCd = "%"
	If sCostCd = ""		Then sCostCd = "%"
	If sItemAcct = ""	Then sItemAcct = "%"
	If sItemGroupCd = "" Then sItemGroupCd = "%"
	If sItemCd = ""		Then sItemCd = "%"
	

    With lgObjComm
		.CommandTimeout = 0
		
		.CommandText = "dbo.usp_C_C4215MA1"		' --  �����ؾ��� SP �� 
	    .CommandType = adCmdStoredProc

		lgObjComm.Parameters.Append lgObjComm.CreateParameter("RETURN_VALUE",  adInteger,adParamReturnValue)	' -- No ���� 

		' -- �����ؾ��� ��ȸ���� �Ķ��Ÿ�� 
		lgObjComm.Parameters.Append lgObjComm.CreateParameter("@START_DT",	adVarXChar,	adParamInput, 6,Replace(sStartDt, "'", "''"))
		
		lgObjComm.Parameters.Append lgObjComm.CreateParameter("@PLANT_CD",	adVarXChar,	adParamInput, 4,Replace(sPlantCd, "'", "''"))
		lgObjComm.Parameters.Append lgObjComm.CreateParameter("@COST_CD",	adVarXChar,	adParamInput, 10,Replace(sCostCd, "'", "''"))
		lgObjComm.Parameters.Append lgObjComm.CreateParameter("@ITEM_ACCT",	adVarXChar,	adParamInput, 10,Replace(sItemAcct, "'", "''"))
		lgObjComm.Parameters.Append lgObjComm.CreateParameter("@ITEM_GROUP_CD",	adVarXChar,	adParamInput, 10,Replace(sItemGroupCd, "'", "''"))
		lgObjComm.Parameters.Append lgObjComm.CreateParameter("@ITEM_CD",	adVarXChar,	adParamInput, 18,Replace(sItemCd, "'", "''"))
		
		lgObjComm.Parameters.Append lgObjComm.CreateParameter("@SHEETCNT",  adSmallInt, adParamInput,, 300)	
		lgObjComm.Parameters.Append lgObjComm.CreateParameter("@NEXTKEY",	adVarXChar,	adParamInput, 15,Replace(sNextKey, "'", "''"))
		lgObjComm.Parameters.Append lgObjComm.CreateParameter("@DEBUG",  adSmallInt, adParamInput,, 0)	' -- isqlw ������ ����ϴ� ������ڵ� 
		    
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
		
    
    ' -- ���� ���ڵ���� �� 2-3�� 
    If Not oRs is Nothing Then

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
			
			sTxt	= Replace(sTxt, "%SUM", "�հ�")
					
			Set oRs = oRs.NextRecordSet()	' -- ����(����Ÿ) ���ڵ������ ���� 
			' ----------------------------------------
		End If
		
		' -- �׷���� �÷� ������ ���ʷ� ����Ÿ��Ʈ�� �����Ѵ�.(�����)
		arrRows		= oRs.GetRows()

		Set oRs = oRs.NextRecordSet()	' -- ����(����Ÿ) ���ڵ������ ���� 
		
		Dim arrDataSet, arrTemp, iLngGrRowCnt, iLngGrColCnt, iMaxCols
		arrDataSet = oRs.GetRows()


		iLngGrRowCnt= UBound(arrRows, 2) 			' �׷���� ��� 
		iLngGrColCnt = UBound(arrRows, 1)			' �׷���� ���� 
		
		iLngColCnt	= CLng(arrRows(UBound(arrRows, 1), 0)) 		' �÷���	: (����/�ݾ�/�ܰ� ����)
		
		iMaxCols	= 11	+ iLngColCnt	' �׷���� �÷���� 
		
		' ----------------------------------------------------------
		If iLngGrRowCnt = 0 Then 
		
			If sNextKey= ""  Then
				Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '�� : No data is found. 
			End If
			
			Response.Write " <Script Language=vbscript>	                        " & vbCr
			Response.Write " Parent.lgStrPrevKey = """"" & vbCr
			Response.Write " </Script>	                        " & vbCr	
			Exit Sub
		End If

		' -- ����Ÿ���� ���ʷ� �迭�� �籸���Ѵ�.
		ReDim arrTemp(iLngColCnt, (iLngGrRowCnt+1) * 2)	

		
		' -- ����Ÿ�� ��� 
		iLngRowCnt	= UBound(arrDataSet, 2)
		
		For iLngRow = 0 To 	iLngRowCnt ' 0 - 49
			arrTemp((CLng(arrDataSet(1, iLngRow))-1), (CLng(arrDataSet(0, iLngRow))-1)*2) = arrDataSet(2, iLngRow)	' -- ��, ��, ��(��)
			arrTemp((CLng(arrDataSet(1, iLngRow))-1), (CLng(arrDataSet(0, iLngRow))-1)*2+1) = arrDataSet(3, iLngRow)	' -- ��, ��, ��(��)
		Next
		
		Set oRs = Nothing
		
		' ----------------------------------------------------------
				
		sRowSeq = arrRows(UBound(arrRows, 1)-1 , iLngGrRowCnt)
		
		Redim TmpBuffer(iLngGrRowCnt)

		' -- �׷���� ������ ���� 
		For iLngRow = 0 To 	iLngGrRowCnt    ' -- ������ٲ� 
			iStrData = ""
			For i = 0 To 1		
				iStrData = iStrData & Chr(11) & ConvLang(ConvSPChars(Trim(arrRows(0, iLngRow))), arrRows(iLngGrColCnt-2, iLngRow))	' -- PLANT_CD
				iStrData = iStrData & Chr(11) & ConvLang(ConvSPChars(Trim(arrRows(1, iLngRow))), arrRows(iLngGrColCnt-2, iLngRow))	' -- COST_CD
				iStrData = iStrData & Chr(11) & ConvLang(ConvSPChars(Trim(arrRows(2, iLngRow))), "0")								' -- COST_NM
				
				iStrData = iStrData & Chr(11) & ConvLang(ConvSPChars(Trim(arrRows(4, iLngRow))), arrRows(iLngGrColCnt-2, iLngRow))								' -- ITEM_ACCT_NM
				iStrData = iStrData & Chr(11) & ConvLang(ConvSPChars(Trim(arrRows(5, iLngRow))), arrRows(iLngGrColCnt-2, iLngRow))	' -- ITEM_GROUP_CD
				iStrData = iStrData & Chr(11) & ConvLang(ConvSPChars(Trim(arrRows(6, iLngRow))), "0")								' -- ITEM_GROUP_NM
				iStrData = iStrData & Chr(11) & ConvLang(ConvSPChars(Trim(arrRows(7, iLngRow))), arrRows(iLngGrColCnt-2, iLngRow))	' -- ITEM_CD
				iStrData = iStrData & Chr(11) & ConvLang(ConvSPChars(Trim(arrRows(8, iLngRow))), "0")								' -- ITEM_NM
				iStrData = iStrData & Chr(11) & arrRows(9, iLngRow)
				iStrData = iStrData & Chr(11) & arrRows(10, iLngRow)
			
				' -- ����Ÿ�׸��� ���(������)
				For iLngCol = 0 To iLngColCnt -1
					iStrData = iStrData & Chr(11) & arrTemp(iLngCol, (iLngRow * 2) + i)
				Next

				If Trim(arrRows(11, iLngRow)) <> "0" Then	' -- �Ұ����� ������ 
					sGrpTxt = sGrpTxt & arrRows(11, iLngRow) & gColSep & arrRows(12, iLngRow) & gRowSep		' -- �Ұ豸��|���ȣ(�迭�� ��ġ�� ����)
				End If
				
				iStrData = iStrData & Chr(11) & arrRows(12, iLngRow)	' -- GROUP_NM
				
				' ----------------------------------------------------------
						
				
				iStrData = iStrData & Chr(11) & Chr(12)
			Next
			TmpBuffer(iLngRow) = iStrData
		Next
				
		iStrData = Join(TmpBuffer, "")
		
		Response.Write " <Script Language=vbscript>	                        " & vbCr
		Response.Write " With parent                                        " & vbCr

		If sNextKey = "" Then
			Response.Write "	.frm1.hSTART_DT.value = """ & sStartDt & """" & vbCr 	
			Response.Write "	.frm1.hPLANT_CD.value = """ & sPlantCd & """" & vbCr 	
			Response.Write "	.frm1.hCOST_CD.value = """ & sCostCd & """" & vbCr
			Response.Write "	.frm1.hITEM_ACCT.value = """ & sItemAcct & """" & vbCr
			Response.Write "	.frm1.hITEM_GROUP_CD.value = """ & sItemGroupCd & """" & vbCr
			Response.Write "	.frm1.hITEM_CD.value = """ & sItemCd & """" & vbCr
			
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
		
		If sNextKey = "" Then
			Response.Write "	Call .AddSpanSpecialColumn(1," & iLngGrRowCnt & ") " & vbCr
		Else
			Response.Write "	Call .AddSpanSpecialColumn(" & CLng(sNextKey)+1 & ", " & sNextKey & "+" & iLngGrRowCnt & ") " & vbCr
		End If
		
		Response.Write "	.frm1.vspdData.MaxCols = """ & (iMaxCols-1) & """" & vbCr 	 	
		Response.Write  "   Call Parent.DbQueryOk()		" & vbCr
		Response.Write " End With                                        " & vbCr
		Response.Write  " </Script>                  " & vbCr
    Else
       Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '�� : No data is found. 
    End If

End Sub	

Function ConvLang(Byval pLang, Byval GroupNo)
	Dim pTmp
	
	If GroupNo <> "0" Then
		pTmp = Replace(pLang , "%1", "�հ�")
		pTmp = Replace(pTmp , "%2", "�۾����üҰ�")
		pTmp = Replace(pTmp , "%3", "ǰ������Ұ�")
		pTmp = Replace(pTmp , "%4", "ǰ��׷�Ұ�")
		pTmp = Replace(pTmp , "%5", "ǰ��Ұ�")
	Else
		pTmp = pLang
	End If
	ConvLang = pTmp
End Function
%>