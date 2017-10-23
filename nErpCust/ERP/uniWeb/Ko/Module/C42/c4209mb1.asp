<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% Option Explicit%>
<% session.CodePage=949 %>

<%
'**********************************************************************************************
'*  1. Module Name          : ���� 
'*  2. Function Name        : ����������������������ȸ 
'*  3. Program ID           : b1256mb1
'*  4. Program Name         : ����������������������ȸ 
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
	Dim sPlantCd, sItemAcct, sItemCd, sType, arrTmp

	' -- ������ �극��ŷ 
	Dim C_SHEET_MAX_CONT 
	
	C_SHEET_MAX_CONT = 100
	
    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status

	' -- �����ؾ��� ��ȸ���� (MA���� �����ִ�)
	Dim sStartDt, sEndDt
	
	sStartDt	= Request("txtStartDt")	
	sEndDt		= Request("txtEndDt")
	sNextKey	= Request("lgStrPrevKey")
	
	sPlantCd	= Request("txtPLANT_CD")	
	sItemAcct	= Request("txtITEM_ACCT")	
		
	If sStartDt = "" And sEndDt = ""  And sPlantCd = "" And sItemAcct = "" Then
		Call DisplayMsgBox("900015", vbInformation, "", "", I_MKSCRIPT)      '�� : No data is found. 
		Exit Sub
	End If
	
	If sPlantCd = "" Then sPlantCd = "%"
	If sItemAcct = "" Then sItemAcct = "%"

	' -- sNextKey ������� 
	If Instr(1, sNextKey, "*") > 0 Then
		arrTmp = Split(sNextKey, gColSep)
		sNextKey = arrTmp(0)
		C_SHEET_MAX_CONT = 32000
	End If
	
    With lgObjComm
		.CommandTimeout = 0
		
		.CommandText = "dbo.usp_C_C4209MA1" 		' --  �����ؾ��� SP �� 
	    .CommandType = adCmdStoredProc

		lgObjComm.Parameters.Append lgObjComm.CreateParameter("RETURN_VALUE",  adInteger,adParamReturnValue)	' -- No ���� 

		' -- �����ؾ��� ��ȸ���� �Ķ��Ÿ�� 
		lgObjComm.Parameters.Append lgObjComm.CreateParameter("@START_DT",	adVarXChar,	adParamInput, 10,Replace(sStartDt, "'", "''"))
		lgObjComm.Parameters.Append lgObjComm.CreateParameter("@END_DT",	adVarXChar,	adParamInput, 10,Replace(sEndDt, "'", "''"))
		'lgObjComm.Parameters.Append lgObjComm.CreateParameter("@PLANT_CD",	adVarXChar,	adParamInput, 4,Replace(sPlantCd, "'", "''"))
		lgObjComm.Parameters.Append lgObjComm.CreateParameter("@ITEM_ACCT",	adVarXChar,	adParamInput, 2,Replace(sItemAcct, "'", "''"))
		
		lgObjComm.Parameters.Append lgObjComm.CreateParameter("@SHEETCNT",  adSmallInt, adParamInput,, C_SHEET_MAX_CONT)	
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
		
    
    If Not oRs is Nothing Then
		
		' -- �׷���� �÷� ������ ���ʷ� ����Ÿ��Ʈ�� �����Ѵ�.(�����)
		arrRows		= oRs.GetRows()

		Set oRs = oRs.NextRecordSet()	' -- ����(����Ÿ) ���ڵ������ ���� 
		
		Dim arrDataSet, arrTemp, iLngGrRowCnt, iLngGrColCnt, iMaxCols
		arrDataSet = oRs.GetRows()


		iLngGrRowCnt= UBound(arrRows, 2)				' �׷���� ��� 
		iLngGrColCnt = UBound(arrRows, 1)				' �׷���� ���� 
		
		iLngColCnt	= CLng(arrRows(UBound(arrRows, 1), 0)) * 3		' �÷���	: (����/�ݾ�/�ܰ� ����)
		
		iMaxCols	= 7	+ iLngColCnt	' �׷���� �÷���� 
		
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
		ReDim arrTemp(iLngColCnt, iLngGrRowCnt)
		
		' -- ����Ÿ�� ��� 
		iLngRowCnt	= UBound(arrDataSet, 2)
		
		For iLngRow = 0 To 	iLngRowCnt
			arrTemp((CLng(arrDataSet(1, iLngRow))-1)*3, CLng(arrDataSet(0, iLngRow))-1) = arrDataSet(2, iLngRow)	' -- ��, ��, ��(��)
			arrTemp((CLng(arrDataSet(1, iLngRow))-1)*3+1, CLng(arrDataSet(0, iLngRow))-1) = arrDataSet(3, iLngRow)	' -- ��, ��, ��(��)
			arrTemp((CLng(arrDataSet(1, iLngRow))-1)*3+2, CLng(arrDataSet(0, iLngRow))-1) = arrDataSet(4, iLngRow)	' -- ��, ��, ��(��)
		Next
		
		Set oRs = Nothing
		
		' ----------------------------------------------------------
				
		sRowSeq = arrRows(UBound(arrRows, 1) -1, iLngGrRowCnt)

		' -- ��ȸ Row�� �ִ������ ��ġ�Ҷ��� ��������Ÿ ������ 
		If CInt(sRowSeq) < C_SHEET_MAX_CONT Then
			sRowSeq = ""
		End If

		' -- �׷���� ������ ���� 
		For iLngRow = 0 To 	iLngGrRowCnt    ' -- ������ٲ� 
				
			iStrData = iStrData & Chr(11) & ""								' -- PLANT_CD
			iStrData = iStrData & Chr(11) & ConvLang(ConvSPChars(Trim(arrRows(0, iLngRow))), arrRows(iLngGrColCnt-2, iLngRow))	' -- ITEM_ACCT
			iStrData = iStrData & Chr(11) & ConvLang(ConvSPChars(Trim(arrRows(1, iLngRow))), "0")					' -- MINOR_NM
			iStrData = iStrData & Chr(11) & ConvLang(ConvSPChars(Trim(arrRows(2, iLngRow))), arrRows(iLngGrColCnt-2, iLngRow))	' -- ITEM_CD
			iStrData = iStrData & Chr(11) & ConvLang(ConvSPChars(Trim(arrRows(3, iLngRow))), "0")					' -- ITEM_NM

			' -- ����Ÿ�׸��� ���(������)
			For iLngCol = 0 To iLngColCnt -1
				iStrData = iStrData & Chr(11) & arrTemp(iLngCol, iLngRow)
			Next

			If Trim(arrRows(4, iLngRow)) <> "0" Then	' -- �Ұ����� ������ 
				sGrpTxt = sGrpTxt & arrRows(4, iLngRow) & gColSep & arrRows(5, iLngRow) & gRowSep		' -- �Ұ豸��|���ȣ(�迭�� ��ġ�� ����)
			End If
				
			iStrData = iStrData & Chr(11) & arrRows(4, iLngRow)
			iStrData = iStrData & Chr(11) & arrRows(5, iLngRow)
			' ----------------------------------------------------------
					
			'iStrData = iStrData & Chr(11) & iLngRow+1
			iStrData = iStrData & Chr(11) & Chr(12)
		Next
				
		Response.Write " <Script Language=vbscript>	                        " & vbCr
		Response.Write " With parent                                        " & vbCr
		Response.Write "	.frm1.hSTART_DT.value = """ & sStartDt & """" & vbCr 	
		Response.Write "	.frm1.hEND_DT.value = """ & sEndDt & """" & vbCr 	
		Response.Write "	.frm1.hITEM_ACCT.value = """ & sPlantCd & """" & vbCr
		
		Response.Write "	.frm1.vspdData.ReDraw = False					" & vbCr 			 
		Response.Write "	.ggoSpread.Source = .frm1.vspdData              " & vbCr 			 
		Response.Write "	.ggoSpread.SSShowData """ & iStrData		       & """" & vbCr
		Response.Write "	.frm1.vspdData.MaxCols = """ & (iMaxCols-1) & """" & vbCr 	 	
		Response.Write "	.frm1.vspdData.ReDraw = True					" & vbCr 			 

		If 	C_SHEET_MAX_CONT = 100 Then
			Response.Write "	.lgStrPrevKey = """ & sRowSeq & """" & vbCr
		Else
			Response.Write "	.lgStrPrevKey = ""*""" & vbCr
		End If

		If sNextKey = "" And sType = "2" Then
			Response.Write  "   Call Parent.SetGridHead(""" & sTxt & """)" & vbCr
		End If
		If sGrpTxt <> "" Then
			Response.Write  "   Call Parent.SetQuerySpreadColor(""" & sGrpTxt & """)" & vbCr
		End If
		
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
		pTmp = Replace(pTmp , "%2", "ǰ������Ұ�")
	Else
		pTmp = Replace(pLang , "%T1", "���� ���")
		pTmp = Replace(pTmp , "%T2", "�⸻ ���")
	End If
	ConvLang = pTmp
End Function
%>

