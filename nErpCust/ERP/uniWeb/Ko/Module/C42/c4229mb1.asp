<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% Option Explicit%>
<% session.CodePage=949 %>

<%
'======================================================================================================
'*  1. Module Name          : Cost
'*  2. Function Name        :CC�������м�(����) 
'*  3. Program ID           : c4229mb1.asp
'*  4. Program Name         : CC�������м�(����)
'*  5. Program Desc         : CC�������м�(����)
'*  6. Modified date(First) : 2005-11-18
'*  7. Modified date(Last)  : 2005-11-18
'*  8. Modifier (First)     : HJO
'*  9. Modifier (Last)      : 
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(��) means that "Do not change"
'======================================================================================================
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
	Dim strFlag
	Dim sCostCd,sWcCd, sEmpNo,tmpKey
	Dim sStartDt,sFrame,sEndDt
	
	sStartDt	= Request("txtFrom_YYYYMM")
	sEndDt	= Request("txtTo_YYYYMM")				
	
	sCostCd	= Request("txtCost_cd")
	
	
	sFrame=request("txtFrame")

	If sCostCd = "" Then sCostCd = "%"

	'------ Developer Coding part (End   ) ------------------------------------------------------------------ 
	
     Call SubCreateCommandObject(lgObjComm)
     
     Select case  sFrame
		case 1 
			Call SubBizQueryA()
		case 2
			Call SubBizQueryB()
		case 3
			Call SubBizQueryC()
		case 4
			Call SubBizQueryD()
		case 5
			Call SubBizQueryE()     
     End Select
     
     Call SubCloseCommandObject(lgObjComm)
     
     Response.End 


'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryA()

	Dim oRs, sTxt, arrRows, iLngRow, iLngCol, iStrData, sNextKey, sRowSeq, iLngRowCnt, iLngColCnt, sGrpTxt

	Dim tmpC1,sTxt2
	
    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status

	' -- �����ؾ��� ��ȸ���� (MA���� �����ִ�)

	sNextKey	= Request("lgStrPrevKey")	
	
	
    With lgObjComm
		.CommandTimeout = 0
		
		.CommandText = "dbo.usp_C_C4229MA1_T1"		' --  �����ؾ��� SP �� 
	    .CommandType = adCmdStoredProc

		lgObjComm.Parameters.Append lgObjComm.CreateParameter("RETURN_VALUE",  adInteger,adParamReturnValue)	' -- No ���� 

		' -- �����ؾ��� ��ȸ���� �Ķ��Ÿ�� 
		lgObjComm.Parameters.Append lgObjComm.CreateParameter("@From_YYYYMM",	adVarXChar,	adParamInput, 6,Replace(sStartDt, "'", "''"))
		lgObjComm.Parameters.Append lgObjComm.CreateParameter("@To_YYYYMM",	adVarXChar,	adParamInput, 6,Replace(sEndDt, "'", "''"))				
		lgObjComm.Parameters.Append lgObjComm.CreateParameter("@Cost_CD",	adVarXChar,	adParamInput, 10,Replace(sCostCd, "'", "''"))
		lgObjComm.Parameters.Append lgObjComm.CreateParameter("@SHEETCNT",  adSmallInt, adParamInput,, 100)	
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

	If oRs.EoF and oRs.Bof and sNextKey="" then
		Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '�� : No data is found. 
		oRs.Close
		Set oRs = Nothing
		Exit Sub
	End If
	
	If sNextKey ="" Then 
		tmpKey=0
	Else
		tmpKey=Clng(sNextKey)
	End IF
	
    If Not oRs is nothing Then
		If sNextKey="" then 
			' --- ��� ���(�����) ----			
			Dim arrColRow, i, j, ColHeadRowNo
			
			ColHeadRowNo = -1000
			arrColRow = oRs.GetRows()
			iLngRowCnt	= UBound(arrColRow, 2) 
			iLngColCnt	= UBound(arrColRow, 1) 
						
			For i = 0 To   iLngRowCnt 
					sTxt = sTxt & arrColRow(0, i) & gColSep
					sTxt = sTxt & arrColRow(0, i) & gColSep 
					sTxt = sTxt & arrColRow(0, i) & gColSep
					sTxt = sTxt & arrColRow(0, i) & gColSep 
					sTxt = sTxt & arrColRow(0, i) & gColSep
					sTxt = sTxt & arrColRow(0, i) & gColSep 
					sTxt2 = sTxt2 & "��������" & gColSep 
					sTxt2 = sTxt2 & "��������" & gColSep 
					sTxt2 = sTxt2 &  "����" & gColSep 
					sTxt2 = sTxt2 &  "�빫��" & gColSep
					sTxt2 = sTxt2 &  "���" & gColSep 
					sTxt2 = sTxt2 &  "���ְ�����" & gColSep  
				Next
				sTxt = sTxt & CStr(ColHeadRowNo) & gRowSep
				ColHeadRowNo = ColHeadRowNo + 1
				sTxt = sTxt & sTxt2 & CStr(ColHeadRowNo) & gRowSep

				' --- �� ȭ�麰�� �����ؾߵ� ġȯ ���ڿ���..
				sTxt	= Replace(sTxt, "%99", "�հ�")			
						
			Set oRs = oRs.NextRecordSet()	' -- ����(����Ÿ) ���ڵ������ ���� 
			' ----------------------------------------
		End If
		
		' -- �׷���� �÷� ������ ���ʷ� ����Ÿ��Ʈ�� �����Ѵ�.(�����)
		arrRows		= oRs.GetRows()
		
		Set oRs = oRs.NextRecordSet()	' -- ����(����Ÿ) ���ڵ������ ���� 
		
		Dim arrDataSet, arrTemp, iLngGrRowCnt, iLngGrColCnt, iMaxCols
		arrDataSet = oRs.GetRows()

		Set oRs = Nothing		'
		
		iLngGrRowCnt= UBound(arrRows, 2)				' �׷���� ��� 
		iLngGrColCnt = UBound(arrRows, 1)				' �׷���� ���� 
	
		iLngColCnt	= CLng(arrRows(UBound(arrRows, 1), 0))	*6				' �÷��� 
		
		iMaxCols	= 3	+ iLngColCnt	' �׷���� �÷���� 
		
		' -- ����Ÿ���� ���ʷ� �迭�� �籸���Ѵ�.
		ReDim arrTemp(iLngColCnt, iLngGrRowCnt)

		' -- ����Ÿ�� ��� 
		iLngRowCnt	= UBound(arrDataSet, 2)
	
		For iLngRow = 0 To 	iLngRowCnt 

			arrTemp(CLng(arrDataSet(1, iLngRow)-1)*6, CLng(arrDataSet(0, iLngRow))-1-tmpKey) = arrDataSet(2, iLngRow)	' -- ��, ��, ��(��)
			arrTemp(CLng(arrDataSet(1, iLngRow)-1)*6+1, CLng(arrDataSet(0, iLngRow))-1-tmpKey) = arrDataSet(3, iLngRow)	' -- ��, ��, ��(��)
			arrTemp(CLng(arrDataSet(1, iLngRow)-1)*6+2, CLng(arrDataSet(0, iLngRow))-1-tmpKey) = arrDataSet(4, iLngRow)	' -- ��, ��, ��(��)
			arrTemp(CLng(arrDataSet(1, iLngRow)-1)*6+3, CLng(arrDataSet(0, iLngRow))-1-tmpKey) = arrDataSet(5, iLngRow)	' -- ��, ��, ��(��)
			arrTemp(CLng(arrDataSet(1, iLngRow)-1)*6+4, CLng(arrDataSet(0, iLngRow))-1-tmpKey) = arrDataSet(6, iLngRow)	' -- ��, ��, ��(��)
			arrTemp(CLng(arrDataSet(1, iLngRow)-1)*6+5, CLng(arrDataSet(0, iLngRow))-1-tmpKey) = arrDataSet(7, iLngRow)	' -- ��, ��, ��(��)
		Next
		
		Set oRs = Nothing
		
		' ----------------------------------------------------------
		If iLngGrRowCnt ="" Then 
			If  sNextKey= ""  Then				
				Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '�� : No data is found. 
				'Response.Write " If Parent.frm1.vspdData.MaxRows > 0 Then Parent.frm1.vspdData.Focus"& vbCr	
			End If

			Response.Write " <Script Language=vbscript>	                        " & vbCr
			Response.Write " With parent                                        " & vbCr
			Response.Write  "	.lgStrPrevKey=""""						 " & vbCr
			Response.Write " End With                                        " & vbCr
			Response.Write  " </Script>                  " & vbCr
			sRowSeq=""
			Exit Sub
		End If

		sRowSeq = arrRows(UBound(arrRows, 1) -1 , iLngGrRowCnt)		' --����Ű�� (�����)		
		Redim TmpBuffer(iLngGrRowCnt)		
		' -- �׷���� ������ ���� 
		For iLngRow = 0 To 	iLngGrRowCnt			

					iStrData = iStrData & Chr(11) & ConvSPChars(replace(arrRows(0, iLngRow),"%1","�հ�"))
					iStrData = iStrData & Chr(11) & ConvSPChars(replace(arrRows(1, iLngRow),"%1","�հ�"))
					
					' -- ����Ÿ�׸��� ���(������)
					For iLngCol =0 To iLngColCnt   '''? 1~
						iStrData = iStrData & Chr(11) & arrTemp(iLngCol, iLngRow)
					Next
					If Trim(arrRows(1, iLngRow)) = "%1" Then	' -- �Ұ����� ������ 
						sGrpTxt = sGrpTxt & arrRows(1, iLngRow) & gColSep & 1 & gColSep & iLngRow+1 & gRowSep		' -- �Ұ豸��|���ȣ(�迭�� ��ġ�� ����)
				
					End If

					iStrData = iStrData & Chr(11) & arrRows(3, iLngRow)

					iStrData = iStrData & Chr(12)	
					TmpBuffer(iLngRow) = iStrData
					iStrData=""
	Next
		iStrData = Join(TmpBuffer, "")
		
	Response.Write " <Script Language=vbscript>	                        " & vbCr
	Response.Write " With parent                                        " & vbCr

	If sNextKey = "" Then
			Response.Write  "   Call Parent.InitSpreadSheet(" & sFrame & "," &  iMaxCols & ")" & vbCr
			Response.Write  "   Call Parent.SetGridHead(""" & sTxt & """)" & vbCr
			
	End If
	Response.Write "	.frm1.vspdData.ReDraw = False					" & vbCr 			 
	Response.Write "	.ggoSpread.Source = .frm1.vspdData              " & vbCr 			 
	Response.Write "	.ggoSpread.SSShowData """ & iStrData		       & """" & vbCr
	Response.Write "	.frm1.vspdData.MaxCols = """ & (iMaxCols) & """" & vbCr 	 		
	Response.Write "	.frm1.vspdData.ReDraw = True					" & vbCr 
	If sGrpTxt <> "" Then
		Response.Write  "   Call Parent.SetQuerySpreadColor(""" & sGrpTxt & """)" & vbCr
	End If			 
	Response.Write "	.lgStrPrevKey = """ & sRowSeq & """" & vbCr 		
	Response.Write "	.frm1.hYYYYMM.value=""" & sStartDt & """" & vbcr
	Response.Write "	.frm1.hYYYYMM2.value=""" & sEndDt & """" & vbcr
	Response.Write "	.frm1.hCost_cd.value=""" & sCostCd & """" & vbcr


	Response.Write  "   Call Parent.DbQueryOk()		" & vbCr
	Response.Write " End With                                        " & vbCr
	Response.Write  " </Script>                  " & vbCr
   End If
       
End Sub	
'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryB()

	Dim oRs, sTxt, arrRows, iLngRow, iLngCol, iStrData, sNextKey, sRowSeq, iLngRowCnt, iLngColCnt, sGrpTxt

	Dim tmpC1,sTxt2
	
    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status

	' -- �����ؾ��� ��ȸ���� (MA���� �����ִ�)

	sNextKey	= Request("lgStrPrevKey")	
	
	
    With lgObjComm
		.CommandTimeout = 0
		
		.CommandText = "dbo.usp_C_C4229MA1_T2"		' --  �����ؾ��� SP �� 
	    .CommandType = adCmdStoredProc

		lgObjComm.Parameters.Append lgObjComm.CreateParameter("RETURN_VALUE",  adInteger,adParamReturnValue)	' -- No ���� 

		' -- �����ؾ��� ��ȸ���� �Ķ��Ÿ�� 
		lgObjComm.Parameters.Append lgObjComm.CreateParameter("@From_YYYYMM",	adVarXChar,	adParamInput, 6,Replace(sStartDt, "'", "''"))
		lgObjComm.Parameters.Append lgObjComm.CreateParameter("@To_YYYYMM",	adVarXChar,	adParamInput, 6,Replace(sEndDt, "'", "''"))				
		lgObjComm.Parameters.Append lgObjComm.CreateParameter("@Cost_CD",	adVarXChar,	adParamInput, 10,Replace(sCostCd, "'", "''"))
		lgObjComm.Parameters.Append lgObjComm.CreateParameter("@SHEETCNT",  adSmallInt, adParamInput,, 50)	
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

	If oRs.EoF and oRs.Bof THEN
		IF sNextKey="" then
			Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '�� : No data is found. 
			oRs.Close
			Set oRs = Nothing
			Exit Sub
		ELSE
			Response.Write " <Script Language=vbscript>	                        " & vbCr
			Response.Write " With parent                                        " & vbCr
			Response.Write  "	.lgStrPrevKey=""""						 " & vbCr
			Response.Write " End With                                        " & vbCr
			Response.Write  " </Script>                  " & vbCr
			oRs.Close
			Set oRs = Nothing
			Exit Sub
		END IF
	End If
	
	If sNextKey ="" Then 
		tmpKey=0
	Else
		tmpKey=Clng(sNextKey)
	End IF
	
    If Not oRs is nothing Then
		If sNextKey="" then 
			' --- ��� ���(�����) ----			
			Dim arrColRow, i, j, ColHeadRowNo
			
			ColHeadRowNo = -1000
			arrColRow = oRs.GetRows()
			iLngRowCnt	= UBound(arrColRow, 2) 
			iLngColCnt	= UBound(arrColRow, 1) 
						
			For j = 0 To iLngColCnt 
				For i = 0 To   iLngRowCnt 
					sTxt = sTxt & arrColRow(j, i) & gColSep
				Next
				sTxt = sTxt & CStr(ColHeadRowNo) & gRowSep
				ColHeadRowNo = ColHeadRowNo + 1
			Next			
				' --- �� ȭ�麰�� �����ؾߵ� ġȯ ���ڿ���..
				sTxt	= Replace(sTxt, "%99", "�հ�")									
			Set oRs = oRs.NextRecordSet()	' -- ����(����Ÿ) ���ڵ������ ���� 
			' ----------------------------------------
		End If
		
		' -- �׷���� �÷� ������ ���ʷ� ����Ÿ��Ʈ�� �����Ѵ�.(�����)
		arrRows		= oRs.GetRows()
		
		Set oRs = oRs.NextRecordSet()	' -- ����(����Ÿ) ���ڵ������ ���� 
		
		Dim arrDataSet, arrTemp, iLngGrRowCnt, iLngGrColCnt, iMaxCols,arrTempG
		
		arrDataSet = oRs.GetRows()

		Set oRs = Nothing		'
		
		iLngGrRowCnt= UBound(arrRows, 2)				' �׷���� ��� 
		iLngGrColCnt = UBound(arrRows, 1)				' �׷���� ���� 
	
		iLngColCnt	= CLng(arrRows(UBound(arrRows, 1), 0))					' �÷���	
		
		iMaxCols	= 5	+ iLngColCnt	' �׷���� �÷���� 
		' -- ����Ÿ���� ���ʷ� �迭�� �籸���Ѵ�.
		ReDim arrTemp(iLngColCnt, (iLngGrRowCnt+1)*3)
		ReDim arrTempG(0, (iLngGrRowCnt+1)*3)

		' -- ����Ÿ�� ��� 
		iLngRowCnt	= UBound(arrDataSet, 2)
	
		For iLngRow = 0 To 	iLngRowCnt 
			arrTempG(0, (CLng(arrDataSet(0, iLngRow))-1-tmpKey)*3) = "�հ�"' -- ��, ��, ��(��)
			arrTempG(0, (CLng(arrDataSet(0, iLngRow))-1-tmpKey)*3+1) = "����"	' -- ��, ��, ��(��)
			arrTempG(0, (CLng(arrDataSet(0, iLngRow))-1-tmpKey)*3+2) = "����"	' -- ��, ��, ��(��)
		
			arrTemp(CLng(arrDataSet(1, iLngRow)-1), (CLng(arrDataSet(0, iLngRow))-1-tmpKey)*3) = arrDataSet(2, iLngRow)	' -- ��, ��, ��(��)
			arrTemp(CLng(arrDataSet(1, iLngRow)-1), (CLng(arrDataSet(0, iLngRow))-1-tmpKey)*3+1) = arrDataSet(3, iLngRow)	' -- ��, ��, ��(��)
			arrTemp(CLng(arrDataSet(1, iLngRow)-1), (CLng(arrDataSet(0, iLngRow))-1-tmpKey)*3+2) = arrDataSet(4, iLngRow)	' -- ��, ��, ��(��)
		
		Next
		
		Set oRs = Nothing

		' ----------------------------------------------------------
'		If iLngGrRowCnt = 0 Then 
'			If  sNextKey= ""  Then				
'				Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '�� : No data is found. 
'				
'			End If
'
'			Response.Write " <Script Language=vbscript>	                        " & vbCr
'			Response.Write " With parent                                        " & vbCr
'			Response.Write  "	.lgStrPrevKey=""""						 " & vbCr
'			Response.Write " End With                                        " & vbCr
'			Response.Write  " </Script>                  " & vbCr
'			sRowSeq=""
'			Exit Sub
'		End If

		sRowSeq = arrRows(UBound(arrRows, 1) -1 , iLngGrRowCnt)		' --����Ű�� (�����)
		
		
		Redim TmpBuffer(iLngGrRowCnt)		
		' -- �׷���� ������ ���� 
		For iLngRow = 0 To 	iLngGrRowCnt		
		
				If Trim(arrRows(0, iLngRow)) = "%1" Then	' -- �Ұ����� ������ 
					sGrpTxt = sGrpTxt & arrRows(0, iLngRow) & gColSep & 1 & gColSep & arrRows(5, iLngRow)*3-2 & gRowSep		' -- �Ұ豸��|���ȣ(�迭�� ��ġ�� ����)
				End If
				If Trim(arrRows(2, iLngRow)) = "%2" Then	' -- �Ұ����� ������ 
					sGrpTxt = sGrpTxt & arrRows(2, iLngRow) & gColSep & 3 & gColSep & arrRows(5, iLngRow)*3-2  & gRowSep		' -- �Ұ豸��|���ȣ(�迭�� ��ġ�� ����)
				End If
						
				For i = 0 To 2	
					iStrData = iStrData & Chr(11) & ConvSPChars(replace(arrRows(0, iLngRow),"%1","�հ�"))
					iStrData = iStrData & Chr(11) & ConvSPChars(replace(arrRows(1, iLngRow),"%1","�հ�"))
					iStrData = iStrData & Chr(11) & ConvSPChars(replace(arrRows(2, iLngRow),"%2","�����Ұ�"))
					iStrData = iStrData & Chr(11) & ConvSPChars(replace(arrRows(3, iLngRow),"%2","�����Ұ�"))
					iStrData = iStrData & Chr(11) & arrTempG(0, (iLngRow * 3) + i)

					' -- ����Ÿ�׸��� ���(������)
					For iLngCol = 0 To iLngColCnt 	
						If arrTemp(iLngCol, (iLngRow * 3) + i)<>"" then 					
						iStrData = iStrData & Chr(11) & arrTemp(iLngCol, (iLngRow * 3) + i)					
						Else
						iStrData = iStrData & Chr(11) & 0					
						End IF
					Next
					iStrData = iStrData & Chr(11) & arrRows(5, iLngRow)
					iStrData = iStrData & Chr(12)	
					
				NEXT
				TmpBuffer(iLngRow) = iStrData
				iStrData=""		
		Next
		iStrData = Join(TmpBuffer, "")

		
	Response.Write " <Script Language=vbscript>	                        " & vbCr
	Response.Write " With parent                                        " & vbCr

	If sNextKey = "" Then
			
			Response.Write  "   Call Parent.InitSpreadSheet(" & sFrame & "," &  iMaxCols & ")" & vbCr
			Response.Write  "   Call Parent.SetGridHead(""" & sTxt & """)" & vbCr
	End If
	Response.Write "	.frm1.vspdData2.ReDraw = False					" & vbCr 			 
	Response.Write "	.ggoSpread.Source = .frm1.vspdData2              " & vbCr 			 
	Response.Write "	.ggoSpread.SSShowData """ & iStrData		       & """" & vbCr
	Response.Write "	.frm1.vspdData2.MaxCols = """ & (iMaxCols) & """" & vbCr 	 		
	Response.Write "	.frm1.vspdData2.ReDraw = True					" & vbCr 
	If sGrpTxt <> "" Then
		Response.Write  "   Call Parent.SetQuerySpreadColor(""" & sGrpTxt & """)" & vbCr
		
	End If			 
	Response.Write "	.lgStrPrevKey = """ & sRowSeq & """" & vbCr 		
	Response.Write "	.frm1.hYYYYMM.value=""" & sStartDt & """" & vbcr
	Response.Write "	.frm1.hYYYYMM2.value=""" & sEndDt & """" & vbcr
	Response.Write "	.frm1.hCost_cd.value=""" & sCostCd & """" & vbcr


	Response.Write  "   Call Parent.DbQueryOk()		" & vbCr
	Response.Write " End With                                        " & vbCr
	Response.Write  " </Script>                  " & vbCr
   End If
       

End Sub

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryC()

	Dim oRs, sTxt, arrRows, iLngRow, iLngCol, iStrData, sNextKey, sRowSeq, iLngRowCnt, iLngColCnt, sGrpTxt

	Dim tmpC1,sTxt2
	
    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status

	' -- �����ؾ��� ��ȸ���� (MA���� �����ִ�)

	sNextKey	= Request("lgStrPrevKey")	
	
	
    With lgObjComm
		.CommandTimeout = 0
		
		.CommandText = "dbo.usp_C_C4229MA1_T3"		' --  �����ؾ��� SP �� 
	    .CommandType = adCmdStoredProc

		lgObjComm.Parameters.Append lgObjComm.CreateParameter("RETURN_VALUE",  adInteger,adParamReturnValue)	' -- No ���� 

		' -- �����ؾ��� ��ȸ���� �Ķ��Ÿ�� 
		lgObjComm.Parameters.Append lgObjComm.CreateParameter("@From_YYYYMM",	adVarXChar,	adParamInput, 6,Replace(sStartDt, "'", "''"))
		lgObjComm.Parameters.Append lgObjComm.CreateParameter("@To_YYYYMM",	adVarXChar,	adParamInput, 6,Replace(sEndDt, "'", "''"))				
		lgObjComm.Parameters.Append lgObjComm.CreateParameter("@Cost_CD",	adVarXChar,	adParamInput, 10,Replace(sCostCd, "'", "''"))
		lgObjComm.Parameters.Append lgObjComm.CreateParameter("@SHEETCNT",  adSmallInt, adParamInput,, 100)	
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

	If oRs.EoF and oRs.Bof and sNextKey="" then
		Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '�� : No data is found. 
		oRs.Close
		Set oRs = Nothing
		Exit Sub
	End If
	
	If sNextKey ="" Then 
		tmpKey=0
	Else
		tmpKey=Clng(sNextKey)
	End IF
	
    If Not oRs is nothing Then
		If sNextKey="" then 
			' --- ��� ���(�����) ----			
			Dim arrColRow, i, j, ColHeadRowNo
			
			ColHeadRowNo = -1000
			arrColRow = oRs.GetRows()
			iLngRowCnt	= UBound(arrColRow, 2) 
			iLngColCnt	= UBound(arrColRow, 1) 
			
			For i = 0 To   iLngRowCnt 
					sTxt = sTxt & arrColRow(0, i) & gColSep
					sTxt = sTxt & arrColRow(0, i) & gColSep 
					sTxt = sTxt & arrColRow(0, i) & gColSep

					sTxt2 = sTxt2 & "���귮*�Ⱓ��������ǰ�" & gColSep 
					sTxt2 = sTxt2 & "��������" & gColSep 
					sTxt2 = sTxt2 &  "����������(%)" & gColSep 
 
			Next
				sTxt = sTxt & CStr(ColHeadRowNo) & gRowSep
				ColHeadRowNo = ColHeadRowNo + 1
				sTxt = sTxt & sTxt2 & CStr(ColHeadRowNo) & gRowSep

				' --- �� ȭ�麰�� �����ؾߵ� ġȯ ���ڿ���..
				sTxt	= Replace(sTxt, "%99", "�����հ�")			
						
			Set oRs = oRs.NextRecordSet()	' -- ����(����Ÿ) ���ڵ������ ���� 
			' ----------------------------------------
		End If
		
		' -- �׷���� �÷� ������ ���ʷ� ����Ÿ��Ʈ�� �����Ѵ�.(�����)
		arrRows		= oRs.GetRows()
		
		Set oRs = oRs.NextRecordSet()	' -- ����(����Ÿ) ���ڵ������ ���� 
		
		Dim arrDataSet, arrTemp, iLngGrRowCnt, iLngGrColCnt, iMaxCols
		arrDataSet = oRs.GetRows()

		Set oRs = Nothing		'
		
		iLngGrRowCnt= UBound(arrRows, 2)				' �׷���� ��� 
		iLngGrColCnt = UBound(arrRows, 1)				' �׷���� ���� 
	
		iLngColCnt	= CLng(arrRows(UBound(arrRows, 1), 0))	*3				' �÷���		
		iMaxCols	= 3	+ iLngColCnt	' �׷���� �÷���� 
			
		' -- ����Ÿ���� ���ʷ� �迭�� �籸���Ѵ�.
		ReDim arrTemp(iLngColCnt, iLngGrRowCnt)

		' -- ����Ÿ�� ��� 
		iLngRowCnt	= UBound(arrDataSet, 2)
		Response.Write iLngGrRowCnt & "=iLngGrRowCnt<br>" & iLngColCnt & "=>COL" 
	
		For iLngRow = 0 To 	iLngRowCnt 																		 
			arrTemp(CLng(arrDataSet(1, iLngRow)-1)*3, CLng(arrDataSet(0, iLngRow))-1-tmpKey) =UniConvNumberDBToCompany( arrDataSet(2, iLngRow),ggAmtOfMoney.DecPoint, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit, 0)	' -- ��, ��, ��(��)
			arrTemp(CLng(arrDataSet(1, iLngRow)-1)*3+1, CLng(arrDataSet(0, iLngRow))-1-tmpKey) = UniConvNumberDBToCompany(arrDataSet(3, iLngRow),ggAmtOfMoney.DecPoint, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit, 0)	' -- ��, ��, ��(��)
			arrTemp(CLng(arrDataSet(1, iLngRow)-1)*3+2, CLng(arrDataSet(0, iLngRow))-1-tmpKey) = UniConvNumberDBToCompany(arrDataSet(4, iLngRow)	,ggExchRate.DecPoint, ggExchRate.RndPolicy, ggExchRate.RndUnit, 0)	' -- ��, ��, ��(��)
		Next														
		
		Set oRs = Nothing
		
		' ----------------------------------------------------------
		If iLngGrRowCnt ="" Then 
			If  sNextKey= ""  Then				
				Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '�� : No data is found. 
			End If

			Response.Write " <Script Language=vbscript>	                        " & vbCr
			Response.Write " With parent                                        " & vbCr
			Response.Write  "	.lgStrPrevKey=""""						 " & vbCr
			Response.Write " End With                                        " & vbCr
			Response.Write  " </Script>                  " & vbCr
			sRowSeq=""
			Exit Sub
		End If

		sRowSeq = arrRows(UBound(arrRows, 1) -1 , iLngGrRowCnt)		' --����Ű�� (�����)
		
		Redim TmpBuffer(iLngGrRowCnt)		
				
		' -- �׷���� ������ ���� 
		For iLngRow = 0 To 	iLngGrRowCnt			

					iStrData = iStrData & Chr(11) & ConvSPChars(replace(arrRows(0, iLngRow),"%1","�հ�"))
					iStrData = iStrData & Chr(11) & ConvSPChars(replace(arrRows(1, iLngRow),"%1","�հ�"))
					
					' -- ����Ÿ�׸��� ���(������)
					For iLngCol =0 To iLngColCnt   
						iStrData = iStrData & Chr(11) & arrTemp(iLngCol, iLngRow)

					Next
					If Trim(arrRows(1, iLngRow)) = "%1" Then	' -- �Ұ����� ������ 
						sGrpTxt = sGrpTxt & arrRows(1, iLngRow) & gColSep & 1 & gColSep &  arrRows(3, iLngRow) & gRowSep		' -- �Ұ豸��|���ȣ(�迭�� ��ġ�� ����)
				
					End If

					iStrData = iStrData & Chr(11) & arrRows(3, iLngRow)

					iStrData = iStrData & Chr(12)	
					TmpBuffer(iLngRow) = iStrData
					iStrData=""
		Next
		iStrData = Join(TmpBuffer, "")
		
	
	Response.Write " <Script Language=vbscript>	                        " & vbCr
	Response.Write " With parent                                        " & vbCr

	If sNextKey = "" Then
			Response.Write  "   Call Parent.InitSpreadSheet(" & sFrame & "," &  iMaxCols & ")" & vbCr
			Response.Write  "   Call Parent.SetGridHead(""" & sTxt & """)" & vbCr
			
	End If
	Response.Write "	.frm1.vspdData3.ReDraw = False					" & vbCr 			 
	Response.Write "	.ggoSpread.Source = .frm1.vspdData3              " & vbCr 			 
	Response.Write "	.ggoSpread.SSShowData """ & iStrData		       & """" & vbCr
	Response.Write "	.frm1.vspdData3.MaxCols = """ & (iMaxCols) & """" & vbCr 	 		
	Response.Write "	.frm1.vspdData3.ReDraw = True					" & vbCr 
	If sGrpTxt <> "" Then
		Response.Write  "   Call Parent.SetQuerySpreadColor(""" & sGrpTxt & """)" & vbCr
	End If			 
	Response.Write "	.lgStrPrevKey = """ & sRowSeq & """" & vbCr 		
	Response.Write "	.frm1.hYYYYMM.value=""" & sStartDt & """" & vbcr
	Response.Write "	.frm1.hYYYYMM2.value=""" & sEndDt & """" & vbcr
	Response.Write "	.frm1.hCost_cd.value=""" & sCostCd & """" & vbcr


	Response.Write  "   Call Parent.DbQueryOk()		" & vbCr
	Response.Write " End With                                        " & vbCr
	Response.Write  " </Script>                  " & vbCr
   End If
       
End Sub	


'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryD()

	Dim oRs, sTxt, arrRows, iLngRow, iLngCol, iStrData, sNextKey, sRowSeq, iLngRowCnt, iLngColCnt, sGrpTxt

	Dim tmpC1,sTxt2
	
    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status

	' -- �����ؾ��� ��ȸ���� (MA���� �����ִ�)

	sNextKey	= Request("lgStrPrevKey")	
	
	
    With lgObjComm
		.CommandTimeout = 0
		
		.CommandText = "dbo.usp_C_C4229MA1_T4"		' --  �����ؾ��� SP �� 
	    .CommandType = adCmdStoredProc

		lgObjComm.Parameters.Append lgObjComm.CreateParameter("RETURN_VALUE",  adInteger,adParamReturnValue)	' -- No ���� 

		' -- �����ؾ��� ��ȸ���� �Ķ��Ÿ�� 
		lgObjComm.Parameters.Append lgObjComm.CreateParameter("@From_YYYYMM",	adVarXChar,	adParamInput, 6,Replace(sStartDt, "'", "''"))
		lgObjComm.Parameters.Append lgObjComm.CreateParameter("@To_YYYYMM",	adVarXChar,	adParamInput, 6,Replace(sEndDt, "'", "''"))				
		lgObjComm.Parameters.Append lgObjComm.CreateParameter("@Cost_CD",	adVarXChar,	adParamInput, 10,Replace(sCostCd, "'", "''"))
		lgObjComm.Parameters.Append lgObjComm.CreateParameter("@SHEETCNT",  adSmallInt, adParamInput,, 100)	
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

	If oRs.EoF and oRs.Bof and sNextKey="" then
		Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '�� : No data is found. 
		oRs.Close
		Set oRs = Nothing
		Exit Sub
	End If
	
	If sNextKey ="" Then 
		tmpKey=0
	Else
		tmpKey=Clng(sNextKey)
	End IF
	
    If Not oRs is nothing Then
		If sNextKey="" then 
			' --- ��� ���(�����) ----			
			Dim arrColRow, i, j, ColHeadRowNo
			
			ColHeadRowNo = -1000
			arrColRow = oRs.GetRows()
			iLngRowCnt	= UBound(arrColRow, 2) 
			iLngColCnt	= UBound(arrColRow, 1) 
			
			
			For j=0 to iLngColCnt
				For i = 0 To   iLngRowCnt 
					sTxt = sTxt & arrColRow(j, i) & gColSep
				Next
				sTxt = sTxt & CStr(ColHeadRowNo) & gRowSep
				ColHeadRowNo = ColHeadRowNo + 1
			Next
				' --- �� ȭ�麰�� �����ؾߵ� ġȯ ���ڿ���..
				sTxt	= Replace(sTxt, "%99", "�����հ�")			
						
			Set oRs = oRs.NextRecordSet()	' -- ����(����Ÿ) ���ڵ������ ���� 
			' ----------------------------------------
		End If
		
		' -- �׷���� �÷� ������ ���ʷ� ����Ÿ��Ʈ�� �����Ѵ�.(�����)
		arrRows		= oRs.GetRows()
		
		Set oRs = oRs.NextRecordSet()	' -- ����(����Ÿ) ���ڵ������ ���� 
		
		Dim arrDataSet, arrTemp, iLngGrRowCnt, iLngGrColCnt, iMaxCols
		arrDataSet = oRs.GetRows()

		Set oRs = Nothing		'
		
		iLngGrRowCnt= UBound(arrRows, 2)				' �׷���� ��� 
		iLngGrColCnt = UBound(arrRows, 1)				' �׷���� ���� 
	
		iLngColCnt	= CLng(arrRows(UBound(arrRows, 1), 0))			' �÷��� 
		
		
		iMaxCols	= 3	+ iLngColCnt	' �׷���� �÷���� 
		
	
		' -- ����Ÿ���� ���ʷ� �迭�� �籸���Ѵ�.
		ReDim arrTemp(iLngColCnt, iLngGrRowCnt)

		' -- ����Ÿ�� ��� 
		iLngRowCnt	= UBound(arrDataSet, 2)
			
		For iLngRow = 0 To 	iLngRowCnt 																		 
			arrTemp(CLng(arrDataSet(1, iLngRow)-1), CLng(arrDataSet(0, iLngRow))-1-tmpKey) =UniConvNumberDBToCompany( arrDataSet(2, iLngRow),ggqty.DecPoint, ggqty.RndPolicy, ggqty.RndUnit, 0)	' -- ��, ��, ��(��)
		Next														
		
		Set oRs = Nothing
		
		' ----------------------------------------------------------
		If iLngGrRowCnt ="" Then 
			If  sNextKey= ""  Then				
				Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '�� : No data is found. 
			End If

			Response.Write " <Script Language=vbscript>	                        " & vbCr
			Response.Write " With parent                                        " & vbCr
			Response.Write  "	.lgStrPrevKey=""""						 " & vbCr
			Response.Write " End With                                        " & vbCr
			Response.Write  " </Script>                  " & vbCr
			sRowSeq=""
			Exit Sub
		End If

		sRowSeq = arrRows(UBound(arrRows, 1) -1 , iLngGrRowCnt)		' --����Ű�� (�����)
		
		Redim TmpBuffer(iLngGrRowCnt)		
		
		' -- �׷���� ������ ���� 
		For iLngRow = 0 To 	iLngGrRowCnt			

					iStrData = iStrData & Chr(11) & ConvSPChars(replace(arrRows(0, iLngRow),"%1","�հ�"))
					iStrData = iStrData & Chr(11) & ConvSPChars(replace(arrRows(1, iLngRow),"%1","�հ�"))
					
					' -- ����Ÿ�׸��� ���(������)
					For iLngCol =0 To iLngColCnt   
						iStrData = iStrData & Chr(11) & arrTemp(iLngCol, iLngRow)
					Next
					If Trim(arrRows(1, iLngRow)) = "%1" Then	' -- �Ұ����� ������ 
						sGrpTxt = sGrpTxt & arrRows(1, iLngRow) & gColSep & 1 & gColSep & arrRows(3, iLngRow) & gRowSep		' -- �Ұ豸��|���ȣ(�迭�� ��ġ�� ����)
					End If
					iStrData = iStrData & Chr(11) & arrRows(3, iLngRow)
					iStrData = iStrData & Chr(12)	
					
					TmpBuffer(iLngRow) = iStrData
					iStrData=""					
		Next
		iStrData = Join(TmpBuffer, "")
					
		Response.Write " <Script Language=vbscript>	                        " & vbCr
		Response.Write " With parent                                        " & vbCr

		If sNextKey = "" Then
				Response.Write  "   Call Parent.InitSpreadSheet(" & sFrame & "," &  iMaxCols & ")" & vbCr
				Response.Write  "   Call Parent.SetGridHead(""" & sTxt & """)" & vbCr
				
		End If
		Response.Write "	.frm1.vspdData4.ReDraw = False					" & vbCr 			 
		Response.Write "	.ggoSpread.Source = .frm1.vspdData4             " & vbCr 			 
		Response.Write "	.ggoSpread.SSShowData """ & iStrData		       & """" & vbCr
		Response.Write "	.frm1.vspdData4.MaxCols = """ & (iMaxCols) & """" & vbCr 	 		
		Response.Write "	.frm1.vspdData4.ReDraw = True					" & vbCr 
		If sGrpTxt <> "" Then
			Response.Write  "   Call Parent.SetQuerySpreadColor(""" & sGrpTxt & """)" & vbCr
		End If			 
		'Response.Write "		msgbox 1 " & vbcr 
		Response.Write "	.lgStrPrevKey = """ & sRowSeq & """" & vbCr 		
		Response.Write "	.frm1.hYYYYMM.value=""" & sStartDt & """" & vbcr
		Response.Write "	.frm1.hYYYYMM2.value=""" & sEndDt & """" & vbcr
		Response.Write "	.frm1.hCost_cd.value=""" & sCostCd & """" & vbcr


		Response.Write  "   Call Parent.DbQueryOk()		" & vbCr
		Response.Write " End With                                        " & vbCr
		Response.Write  " </Script>                  " & vbCr
   End If

       
End Sub	

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryE()

	Dim oRs, sTxt, arrRows, iLngRow, iLngCol, iStrData, sNextKey, sRowSeq, iLngRowCnt, iLngColCnt, sGrpTxt

	Dim tmpC1,sTxt2
	
    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status

	' -- �����ؾ��� ��ȸ���� (MA���� �����ִ�)

	sNextKey	= Request("lgStrPrevKey")	
	
	
    With lgObjComm
		.CommandTimeout = 0
		
		.CommandText = "dbo.usp_C_C4229MA1_T5"		' --  �����ؾ��� SP �� 
	    .CommandType = adCmdStoredProc

		lgObjComm.Parameters.Append lgObjComm.CreateParameter("RETURN_VALUE",  adInteger,adParamReturnValue)	' -- No ���� 

		' -- �����ؾ��� ��ȸ���� �Ķ��Ÿ�� 
		lgObjComm.Parameters.Append lgObjComm.CreateParameter("@From_YYYYMM",	adVarXChar,	adParamInput, 6,Replace(sStartDt, "'", "''"))
		lgObjComm.Parameters.Append lgObjComm.CreateParameter("@To_YYYYMM",	adVarXChar,	adParamInput, 6,Replace(sEndDt, "'", "''"))				
		lgObjComm.Parameters.Append lgObjComm.CreateParameter("@Cost_CD",	adVarXChar,	adParamInput, 10,Replace(sCostCd, "'", "''"))
		lgObjComm.Parameters.Append lgObjComm.CreateParameter("@SHEETCNT",  adSmallInt, adParamInput,, 100)	
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

	If oRs.EoF and oRs.Bof and sNextKey="" then
		Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '�� : No data is found. 
		oRs.Close
		Set oRs = Nothing
		Exit Sub
	End If
	
	If sNextKey ="" Then 
		tmpKey=0
	Else
		tmpKey=Clng(sNextKey)
	End IF
	
    If Not oRs is nothing Then
		If sNextKey="" then 
			' --- ��� ���(�����) ----			
			Dim arrColRow, i, j, ColHeadRowNo
			
			ColHeadRowNo = -1000
			arrColRow = oRs.GetRows()
			iLngRowCnt	= UBound(arrColRow, 2) 
			iLngColCnt	= UBound(arrColRow, 1) 			
	
				For i = 0 To   iLngRowCnt 
					sTxt = sTxt & arrColRow(0, i) & gColSep
					sTxt = sTxt & arrColRow(0, i) & gColSep
					sTxt = sTxt & arrColRow(0, i) & gColSep
					sTxt2 = sTxt2 & "��������" & gColSep 
					sTxt2 = sTxt2 & "��ο��DATA" & gColSep 
					sTxt2 = sTxt2 &  "����(%)" & gColSep 
 
				Next
				sTxt = sTxt & CStr(ColHeadRowNo) & gRowSep
				ColHeadRowNo = ColHeadRowNo + 1
				sTxt = sTxt & sTxt2 & CStr(ColHeadRowNo) & gRowSep				

			Set oRs = oRs.NextRecordSet()	' -- ����(����Ÿ) ���ڵ������ ���� 
			' ----------------------------------------
		End If
		
		' -- �׷���� �÷� ������ ���ʷ� ����Ÿ��Ʈ�� �����Ѵ�.(�����)
		arrRows		= oRs.GetRows()
		
		Set oRs = oRs.NextRecordSet()	' -- ����(����Ÿ) ���ڵ������ ���� 
		
		Dim arrDataSet, arrTemp, iLngGrRowCnt, iLngGrColCnt, iMaxCols
		arrDataSet = oRs.GetRows()

		Set oRs = Nothing		'
		
		iLngGrRowCnt= UBound(arrRows, 2)				' �׷���� ��� 
		iLngGrColCnt = UBound(arrRows, 1)				' �׷���� ���� 
	
		iLngColCnt	= CLng(arrRows(UBound(arrRows, 1), 0))	*3		' �÷���			
		iMaxCols	= 5	+ iLngColCnt	' �׷���� �÷���� 
	
		' -- ����Ÿ���� ���ʷ� �迭�� �籸���Ѵ�.
		ReDim arrTemp(iLngColCnt, iLngGrRowCnt)

		' -- ����Ÿ�� ��� 
		iLngRowCnt	= UBound(arrDataSet, 2)		
	
		For iLngRow = 0 To 	iLngRowCnt 																		 
			arrTemp(CLng(arrDataSet(1, iLngRow)-1)*3, CLng(arrDataSet(0, iLngRow))-1-tmpKey) = arrDataSet(2, iLngRow)
			arrTemp(CLng(arrDataSet(1, iLngRow)-1)*3+1, CLng(arrDataSet(0, iLngRow))-1-tmpKey) = arrDataSet(3, iLngRow)
			arrTemp(CLng(arrDataSet(1, iLngRow)-1)*3+2, CLng(arrDataSet(0, iLngRow))-1-tmpKey) = arrDataSet(4, iLngRow)
		Next														
		
		Set oRs = Nothing
		
		' ----------------------------------------------------------
		If iLngGrRowCnt ="" Then 
			If  sNextKey= ""  Then				
				Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '�� : No data is found. 
				'Response.Write " If Parent.frm1.vspdData.MaxRows > 0 Then Parent.frm1.vspdData.Focus"& vbCr	
			End If

			Response.Write " <Script Language=vbscript>	                        " & vbCr
			Response.Write " With parent                                        " & vbCr
			Response.Write  "	.lgStrPrevKey=""""						 " & vbCr
			Response.Write " End With                                        " & vbCr
			Response.Write  " </Script>                  " & vbCr
			sRowSeq=""
			Exit Sub
		End If

		sRowSeq = arrRows(UBound(arrRows, 1) -1 , iLngGrRowCnt)		' --����Ű�� (�����)		
		Redim TmpBuffer(iLngGrRowCnt)		
		' -- �׷���� ������ ���� 
		For iLngRow = 0 To 	iLngGrRowCnt			

					iStrData = iStrData & Chr(11) & ConvSPChars(replace(arrRows(0, iLngRow),"%1","�հ�"))
					iStrData = iStrData & Chr(11) & ConvSPChars(replace(arrRows(1, iLngRow),"%1","�հ�"))
					
					iStrData = iStrData & Chr(11) & ConvSPChars(replace(arrRows(2, iLngRow),"%2","��ο�ҼҰ�"))
					iStrData = iStrData & Chr(11) & ConvSPChars(replace(arrRows(3, iLngRow),"%2","��ο�ҼҰ�"))
					
					' -- ����Ÿ�׸��� ���(������)
					For iLngCol =0 To iLngColCnt   
						iStrData = iStrData & Chr(11) & UniConvNumberDBToCompany(arrTemp(iLngCol, iLngRow),ggExchRate.DecPoint, ggExchRate.RndPolicy, ggExchRate.RndUnit, 0)	
					Next
					If Trim(arrRows(0, iLngRow)) = "%1" Then	' -- �Ұ����� ������ 
						sGrpTxt = sGrpTxt & arrRows(0, iLngRow) & gColSep & 1 & gColSep & arrRows(5, iLngRow) & gRowSep		' -- �Ұ豸��|���ȣ(�迭�� ��ġ�� ����)
				
					End If
					If Trim(arrRows(2, iLngRow)) = "%2" Then	' -- �Ұ����� ������ 
						sGrpTxt = sGrpTxt & arrRows(2, iLngRow) & gColSep & 3 & gColSep & arrRows(5, iLngRow) & gRowSep		' -- �Ұ豸��|���ȣ(�迭�� ��ġ�� ����)
				
					End If					
					iStrData = iStrData & Chr(11) & arrRows(5, iLngRow)
					iStrData = iStrData & Chr(12)	
					
					TmpBuffer(iLngRow) = iStrData
					iStrData=""
					
		Next
		iStrData = Join(TmpBuffer, "")
				
	Response.Write " <Script Language=vbscript>	                        " & vbCr
	Response.Write " With parent                                        " & vbCr

	If sNextKey = "" Then
			Response.Write  "   Call Parent.InitSpreadSheet(" & sFrame & "," &  iMaxCols & ")" & vbCr
			Response.Write  "   Call Parent.SetGridHead(""" & sTxt & """)" & vbCr
			
	End If
	Response.Write "	.frm1.vspdData5.ReDraw = False					" & vbCr 			 
	Response.Write "	.ggoSpread.Source = .frm1.vspdData5             " & vbCr 			 
	Response.Write "	.ggoSpread.SSShowData """ & iStrData		       & """" & vbCr
	Response.Write "	.frm1.vspdData5.MaxCols = """ & (iMaxCols) & """" & vbCr 	 		
	Response.Write "	.frm1.vspdData5.ReDraw = True					" & vbCr 
	If sGrpTxt <> "" Then
		Response.Write  "   Call Parent.SetQuerySpreadColor(""" & sGrpTxt & """)" & vbCr
	End If			 
	'Response.Write "		msgbox 1 " & vbcr 
	Response.Write "	.lgStrPrevKey = """ & sRowSeq & """" & vbCr 		
	Response.Write "	.frm1.hYYYYMM.value=""" & sStartDt & """" & vbcr
	Response.Write "	.frm1.hYYYYMM2.value=""" & sEndDt & """" & vbcr
	Response.Write "	.frm1.hCost_cd.value=""" & sCostCd & """" & vbcr


	Response.Write  "   Call Parent.DbQueryOk()		" & vbCr
	Response.Write " End With                                        " & vbCr
	Response.Write  " </Script>                  " & vbCr
   End If

       
End Sub	
%>

