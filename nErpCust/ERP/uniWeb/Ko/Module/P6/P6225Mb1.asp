 <%@ LANGUAGE=VBSCript%>
<%Option Explicit    %>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgent.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgentVariables.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%
Call LoadBasisGlobalInf
Call LoadInfTB19029B("I",	"*", "NOCOOKIE","MB")

'**********************************************************************************************
'*  1. Module Name          : Procurement
'*  2. Function Name        :
'*  3. Program ID           : P6225ma1
'*  4. Program Name         :
'*  5. Program Desc         :
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2005/01/23
'*  8. Modified date(Last)  :
'*  9. Modifier (First)     : LEE SANG HO
'* 10. Modifier (Last)      :
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"

Dim lgOpModeCRUD

Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0                 '�� : DBAgent Parameter ���� 
Dim rs1, rs2, rs3, rs4,rs5
Dim istrData1
Dim istrData2
Dim istrData3
Dim iStrPoNo
Dim StrNextKey		' ���� �� 
Dim lgStrPrevKey	' ���� �� 
Dim iLngMaxRow1		' ���� �׸����� �ִ�Row
Dim iLngMaxRow2		' ���� �׸����� �ִ�Row
Dim iLngMaxRow3		' ���� �׸����� �ִ�Row
Dim iLngRow
Dim GroupCount
Dim lgCurrency
Dim index,Count     ' ���� �� Return ���� ���� ������ ���� ���� 
Dim lgDataExist
Dim lgPageNo_A
Dim lgPageNo_B
Dim lgPageNo_C
Dim lgMaxCount
Dim strFlag

	Const C_SHEETMAXROWS_D  = 100000

    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status

    Call HideStatusWnd                                                               '��: Hide Processing message
	'------ Developer Coding part (Start ) ------------------------------------------------------------------

	'------ Developer Coding part (End   ) ------------------------------------------------------------------

    lgOpModeCRUD  = Request("txtMode")

										                                              '��: Read Operation Mode (CRUD)
    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '��: Query
             Call  SubBizQueryMulti()
        Case CStr(UID_M0002)                                                         '��: Save,Update
             Call SubBizSaveMulti()
        Case CStr(UID_M0003)
             Call SubBizSaveMulti()
    End Select

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()
    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status

End Sub
'============================================================================================================
' Name : SubBizSave
' Desc : Save Data
'============================================================================================================
Sub SubBizSave()
    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status
End Sub
'============================================================================================================
' Name : SubBizDelete
' Desc : Delete DB data
'============================================================================================================
Sub SubBizDelete()
    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status
End Sub

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()

    On Error Resume Next

	lgPageNo_A       = UNICInt(Trim(Request("lgPageNo_A")),0)    '��: "0"(First),"1"(Second),"2"(Third),"3"(...)
	lgPageNo_B       = UNICInt(Trim(Request("lgPageNo_B")),0)    '��: "0"(First),"1"(Second),"2"(Third),"3"(...)
	lgPageNo_C       = UNICInt(Trim(Request("lgPageNo_C")),0)    '��: "0"(First),"1"(Second),"2"(Third),"3"(...)
    lgMaxCount       = C_SHEETMAXROWS_D                           '�� : �ѹ��� �����ü� �ִ� ����Ÿ �Ǽ� 
	lgDataExist      = "No"
	iLngMaxRow1	     = CDbl(lgMaxCount) * CDbl(lgPageNo_A) + 1
	iLngMaxRow2	     = CDbl(lgMaxCount) * CDbl(lgPageNo_B) + 1
	iLngMaxRow3	     = CDbl(lgMaxCount) * CDbl(lgPageNo_C) + 1

	lgStrPrevKey = Request("lgStrPrevKey")


	Call FixUNISQLData()
	Call QueryData()

	'====================
	'Call PO_DTL List
	'====================
	
	'-----------------------
	'Result data display area
	'-----------------------

	if Request("txtType") = "A" Then							'�� : ������ �˻� 

		Response.Write "<Script Language=vbscript>" & vbCr
		Response.Write "With parent" & vbCr
		Response.Write "	If .frm1.vspdData.MaxRows < 1 then"						& vbCr
		Response.Write "	End if"							& vbCr
		
		
		Response.Write "	.ggoSpread.Source       = .frm1.vspdData "			& vbCr
		Response.Write "	.ggoSpread.SSShowData     """ & istrData1	 & """" & vbCr
		Response.Write "	.lgPageNo_A  = """ & lgPageNo_A   & """" & vbCr
		
		Response.Write " .DbQueryOk "	& vbCr
		Response.Write "End With"		& vbCr
		Response.Write "</Script>"		& vbCr
	Elseif Request("txtType") = "B" Then
		Response.Write "<Script Language=vbscript>" & vbCr
		Response.Write "With parent" & vbCr
		Response.Write "	If .frm1.vspdData1.MaxRows < 1 then"						& vbCr
		Response.Write "	End if"							& vbCr
		Response.Write "	.ggoSpread.Source       = .frm1.vspdData1 "			& vbCr
		Response.Write "	.ggoSpread.SSShowData     """ & istrData2	 & """" & vbCr
		Response.Write "	.lgPageNo_B  = """ & lgPageNo_B   & """" & vbCr
		Response.Write " .DbDtlQuery2 "	& vbCr
		Response.Write "End With"		& vbCr
		Response.Write "</Script>"		& vbCr
	Elseif Request("txtType") = "C" Then
		Response.Write "<Script Language=vbscript>" & vbCr
		Response.Write "With parent" & vbCr
		Response.Write "	If .frm1.vspdData2.MaxRows < 1 then"						& vbCr
		Response.Write "	End if"							& vbCr
		Response.Write "	.ggoSpread.Source       = .frm1.vspdData2 "			& vbCr
		Response.Write "	.ggoSpread.SSShowData     """ & istrData3	 & """" & vbCr
		Response.Write "	.lgPageNo_C  = """ & lgPageNo_C   & """" & vbCr
		
' 		Response.Write " .DbDtlQueryOk2 "	& vbCr
		Response.Write "End With"		& vbCr
		Response.Write "</Script>"		& vbCr
	End if
End Sub

'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
' Query�ϱ� ����  DB Agent �迭�� �̿��Ͽ� Query���� ����� ���ν��� 
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()

	Dim strCarKind, strCast_Cd
	Dim strWork_Dt, strPlantCd

	Redim UNISqlId(2)                                                     '��: SQL ID ������ ���� ����Ȯ�� 
	Redim UNIValue(2, 3)

	UNISqlId(0) = "Y6000Y600"
	UNISqlId(1) = "Y6110Y6AA"
	UNISqlId(2) = "Y6110Y641"

	IF Request("txtWork_Dt") = "" Then
	   strWork_Dt = "1999-01-01"
	ELSE
	   strWork_Dt = FilterVar(Ucase(Trim(Request("txtWork_Dt"))),"''","S")
	END IF
	IF Request("txtCast_Cd") = "" Then
	   strCast_Cd = "'%'"
	ELSE
	   strCast_Cd = FilterVar(Ucase(Trim(Request("txtCast_Cd"))),"''","S")
	END IF
	IF Request("txtPlantCd") = "" Then
	   strPlantCd = "'%'"
	ELSE
	   strPlantCd = FilterVar(Ucase(Trim(Request("txtPlantCd"))),"''","S")
	END IF
	IF Request("txtCarKind") = "" Then
	   strCarKind = "'%'"
	ELSE
	   strCarKind = FilterVar(Ucase(Trim(Request("txtCarKind"))),"''","S")
	END IF


	UNIValue(0, 0) = FilterVar(Ucase(Trim(Request("txtPlantCd"))),"''","S")

	UNIValue(1, 0) = FilterVar(Ucase(Trim(Request("txtCast_Cd"))),"''","S")

	UNIValue(2, 0) = strWork_Dt
	UNIValue(2, 1) = strPlantCd
	UNIValue(2, 2) = strCarKind
	UNIValue(2, 3) = strCast_Cd

	UNILock = DISCONNREAD :	UNIFlag = "1"


End Sub

'----------------------------------------------------------------------------------------------------------
' Query Data
' ADO�� Record Set�̿��Ͽ� Query�� �ϰ� Record Set�� �Ѱܼ� MakeSpreadSheetData1()���� Spreadsheet�� �����͸� 
' �Ѹ� 
' ADO ��ü�� �����Ҷ� prjPublic.dll������ �̿��Ѵ�.(�󼼳����� vb�� �ۼ��� prjPublic.dll �ҽ� ����)
'----------------------------------------------------------------------------------------------------------
Sub QueryData()
    Dim lgstrRetMsg                                             '�� : Record Set Return Message �������� 
    Dim lgADF                                                   '�� : ActiveX Data Factory ���� �������� 
    Dim iStr
	Dim strCarKind, strCast_Cd
	Dim strWork_Dt, strPlantCd


	IF Request("txtWork_Dt") = "" Then
	   strWork_Dt = "1900-01-01"
	ELSE
	   strWork_Dt = FilterVar(Ucase(Trim(Request("txtWork_Dt"))),"''","S")
	END IF
	IF Request("txtCast_Cd") = "" Then
	   strCast_Cd = "'%'"
	ELSE
	   strCast_Cd = FilterVar(Ucase(Trim(Request("txtCast_Cd"))),"''","S")
	END IF
	IF Request("txtPlantCd") = "" Then
	   strPlantCd = "'%'"
	ELSE
	   strPlantCd = FilterVar(Ucase(Trim(Request("txtPlantCd"))),"''","S")
	END IF
	IF Request("txtCarKind") = "" Then
	   strCarKind = "'%'"
	ELSE
	   strCarKind = FilterVar(Ucase(Trim(Request("txtCarKind"))),"''","S")
	END IF


	if Request("txtType") = "A" Then							'�� : ������ �˻� 

	    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
	    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2)
	
		Set lgADF   = Nothing
	
	    iStr = Split(lgstrRetMsg,gColSep)
	
		If iStr(0) <> "0" Then
	        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
	    End If


		IF Request("txtCAST_Cd") <> "" Then
		    If  rs1.EOF And rs1.BOF  Then
				strFlag = "ERROR_CAST"
		        Response.Write "<Script Language=vbscript>" & vbCr
				Response.Write "parent.frm1.txtCastCd.value = """ & "" & """" & vbCr
				Response.Write "parent.frm1.txtCastNm.value = """ & "" & """" & vbCr
		        Response.Write "</Script>"		& vbCr
		        Response.end
			Else
		        Response.Write "<Script Language=vbscript>" & vbCr
				Response.Write "parent.frm1.txtCastNm.value = """ & ConvSPChars(rs1("CAST_NM")) & """" & vbCr
		        Response.Write "</Script>"		& vbCr
			End If
		End if

	    rs0.Close
	    Set rs0 = Nothing
	    rs1.Close
	    Set rs1 = Nothing

	    If  rs2.EOF And rs2.BOF  Then
			Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
	        Response.Write "<Script Language=vbscript>" & vbCr
	        Response.Write "</Script>"		& vbCr
	        Response.end
	    Else
	        Call  MakeSpreadSheetData1()
	    End If
	Elseif Request("txtType") = "B" Then

		Redim UNISqlId(1)                                                     '��: SQL ID ������ ���� ����Ȯ�� 
		Redim UNIValue(1, 1)
		UNIValue(0, 0) = strWork_Dt
		UNIValue(0, 1) = FilterVar(Ucase(Trim(Request("txtCast_Cd"))),"''","S")

		UNISqlId(0) = "Y6110Y651"

	    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
	    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)
	
		Set lgADF   = Nothing
	
	    iStr = Split(lgstrRetMsg,gColSep)
	
		If iStr(0) <> "0" Then
	        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
	    End If

	    If  rs0.EOF And rs0.BOF  Then
' 			Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
' 	        Response.Write "<Script Language=vbscript>" & vbCr
' 	        Response.Write "</Script>"		& vbCr
' 	        Response.end
	    Else
	        Call  MakeSpreadSheetData2()
	    End If
	Elseif Request("txtType") = "C" Then

		Redim UNISqlId(1)                                                     '��: SQL ID ������ ���� ����Ȯ�� 
		Redim UNIValue(1, 1)
		UNIValue(0, 0) = strWork_Dt
		UNIValue(0, 1) = FilterVar(Ucase(Trim(Request("txtCast_Cd"))),"''","S")

		UNISqlId(0) = "Y6110Y661"

	    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
	    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)
	
		Set lgADF   = Nothing
	
	    iStr = Split(lgstrRetMsg,gColSep)
	
		If iStr(0) <> "0" Then
	        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
	    End If

	    If  rs0.EOF And rs0.BOF  Then
' 			Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
' 	        Response.Write "<Script Language=vbscript>" & vbCr
' 	        Response.Write "</Script>"		& vbCr
' 	        Response.end
	    Else
	        Call  MakeSpreadSheetData3()
	    End If
	End If

'     Call DisplayMsgBox("x", vbInformation, "�̻��ϳ�", "FASDFADS1111", I_MKSCRIPT)


End Sub


'----------------------------------------------------------------------------------------------------------
'QueryData()�� ���ؼ� Query�� �Ǹ� MakeSpreadSheetData1()�� ���ؼ� �����͸� ���������Ʈ�� �ѷ��ִ� ���ν��� 
'----------------------------------------------------------------------------------------------------------
Sub MakeSpreadSheetData1()

    Dim iLoopCount
    Dim iRowStr
    Dim ColCnt
    lgDataExist    = "Yes"
    If CLng(lgPageNo_A) > 0 Then
       rs2.Move     = CLng(lgMaxCount) * CLng(lgPageNo_A)                  'lgMaxCount:Max Fetched Count at once , lgStrPrevKeyIndex : Previous PageNo
    End If

	iLoopCount = 0
	Do while Not (rs2.EOF Or rs2.BOF)
        iLoopCount =  iLoopCount + 1
        iRowStr = ""
        iRowStr = iRowStr & Chr(11) & ConvSPChars(rs2("FAC_CAST_CD"))
        iRowStr = iRowStr & Chr(11) & ""
        iRowStr = iRowStr & Chr(11) & ConvSPChars(rs2("CAST_NM"))
        iRowStr = iRowStr & Chr(11) & ConvSPChars(rs2("PLANT_NM"))
        iRowStr = iRowStr & Chr(11) & ConvSPChars(rs2("CAR_KIND_NM"))
        iRowStr = iRowStr & Chr(11) & UNIDateClientFormat(rs2("WORK_DT")) 
        iRowStr = iRowStr & Chr(11) & ConvSPChars(rs2("INSP_TEXT"))
        iRowStr = iRowStr & Chr(11) & UNINumClientFormat(rs2("INSP_HOUR"),ggQty.DecPoint,0)
        iRowStr = iRowStr & Chr(11) & UNINumClientFormat(rs2("INSP_MIN"),ggQty.DecPoint,0)
        iRowStr = iRowStr & Chr(11) & ConvSPChars(rs2("REQ_DEPT"))
        iRowStr = iRowStr & Chr(11) & ""
        iRowStr = iRowStr & Chr(11) & ConvSPChars(rs2("REQ_DEPT_NM"))
        iRowStr = iRowStr & Chr(11) & ConvSPChars(rs2("INSP_DEPT"))
        iRowStr = iRowStr & Chr(11) & ""
        iRowStr = iRowStr & Chr(11) & ConvSPChars(rs2("INSP_DEPT_NM"))
        iRowStr = iRowStr & Chr(11) & UNINumClientFormat(rs2("INSP_EMP_QTY"),ggQty.DecPoint,0)
        iRowStr = iRowStr & Chr(11) & UNINumClientFormat(rs2("PAYROLL"),ggQty.DecPoint,0)
        iRowStr = iRowStr & Chr(11) & UNINumClientFormat(rs2("MATL_COST"),ggQty.DecPoint,0)
        iRowStr = iRowStr & Chr(11) & ""
        iRowStr = iRowStr & Chr(11) & ConvSPChars(rs2("INSP_FLAG"))
        iRowStr = iRowStr & Chr(11) & iLngMaxRow1 + iLoopCount

        If iLoopCount - 1 < lgMaxCount Then
           istrData1 = istrData1 & iRowStr & Chr(11) & Chr(12)
        Else
           lgPageNo_A = lgPageNo_A + 1
           Exit Do
        End If
        rs2.MoveNext
	Loop

    If iLoopCount <= lgMaxCount Then                                      '��: Check if next data exists
       lgPageNo_A = ""
    End If
    rs2.Close                                                       '��: Close recordset object
    Set rs2 = Nothing	                                            '��: Release ADF
End Sub

'----------------------------------------------------------------------------------------------------------
'QueryData()�� ���ؼ� Query�� �Ǹ� MakeSpreadSheetData2()�� ���ؼ� �����͸� ���������Ʈ�� �ѷ��ִ� ���ν��� 
'----------------------------------------------------------------------------------------------------------
Sub MakeSpreadSheetData2()

    Dim iLoopCount
    Dim iRowStr
    Dim ColCnt

    lgDataExist    = "Yes"
    If CLng(lgPageNo_B) > 0 Then
       rs0.Move     = CLng(lgMaxCount) * CLng(lgPageNo_B)                  'lgMaxCount:Max Fetched Count at once , lgStrPrevKeyIndex : Previous PageNo
    End If

	iLoopCount = 0
	Do while Not (rs0.EOF Or rs0.BOF)

        iLoopCount =  iLoopCount + 1
   
        iRowStr = ""
        iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0("SEQ"))
        iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0("ZINSP_PART"))
        iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0("ZINSP_PART_nm"))
        iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0("INSP_PART"))
        iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0("INSP_PART_nm"))
        iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0("INSP_METH"))
        iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0("INSP_METH_nm"))
        iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0("INSP_DECISION"))
        iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0("INSP_DECISION_nm"))
        iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0("ST_GO_GUBUN"))
        iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0("ST_GO_GUBUN_nm"))
        iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0("SURY_ASSY"))
        iRowStr = iRowStr & Chr(11) & ""
        iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0("ITEM_NM"))
        iRowStr = iRowStr & Chr(11) & UNINumClientFormat(rs0("S_QTY"),ggQty.DecPoint,0)
        iRowStr = iRowStr & Chr(11) & UNINumClientFormat(rs0("PRICE"),ggQty.DecPoint,0)
        iRowStr = iRowStr & Chr(11) & UNINumClientFormat(rs0("SURY_AMT"),ggQty.DecPoint,0)
        iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0("SURI_TYPE"))
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0("SURI_TYPE_NM"))
        iRowStr = iRowStr & Chr(11) & iLngMaxRow2 + iLoopCount

        If iLoopCount - 1 < lgMaxCount Then
           istrData2 = istrData2 & iRowStr & Chr(11) & Chr(12)
        Else
           lgPageNo_B = lgPageNo_B + 1
           Exit Do
        End If

        rs0.MoveNext
	Loop


    If iLoopCount <= lgMaxCount Then                                      '��: Check if next data exists
       lgPageNo_B = ""
    End If
    rs0.Close                                                       '��: Close recordset object
    Set rs0 = Nothing	                                            '��: Release ADF
End Sub

'----------------------------------------------------------------------------------------------------------
'QueryData()�� ���ؼ� Query�� �Ǹ� MakeSpreadSheetData3()�� ���ؼ� �����͸� ���������Ʈ�� �ѷ��ִ� ���ν��� 
'----------------------------------------------------------------------------------------------------------
Sub MakeSpreadSheetData3()
    Dim iLoopCount
    Dim iRowStr
    Dim ColCnt

    lgDataExist    = "Yes"


    If CLng(lgPageNo_C) > 0 Then
       rs0.Move     = CLng(lgMaxCount) * CLng(lgPageNo_C)                  'lgMaxCount:Max Fetched Count at once , lgStrPrevKeyIndex : Previous PageNo
    End If

	iLoopCount = 0
	Do while Not (rs0.EOF Or rs0.BOF)
        iLoopCount =  iLoopCount + 1
        iRowStr = ""
        iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0("SEQ"))
        iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0("INSP_EMP_GB"))
        iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0("INSP_EMP_GB_NM"))
        iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0("INSP_EMP_CD"))
        iRowStr = iRowStr & Chr(11) & ""
        iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0("Emp_nm"))
        iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0("CUST_CD"))
        iRowStr = iRowStr & Chr(11) & ""
        iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0("CUST_NM"))
        iRowStr = iRowStr & Chr(11) & UNINumClientFormat(rs0("INSP_HOUR"),ggQty.DecPoint,0)
        iRowStr = iRowStr & Chr(11) & UNINumClientFormat(rs0("INSP_MIN"),ggQty.DecPoint,0)
        iRowStr = iRowStr & Chr(11) & UNINumClientFormat(rs0("PAYROLL"),ggQty.DecPoint,0)

        iRowStr = iRowStr & Chr(11) & iLngMaxRow3 + iLoopCount

        If iLoopCount - 1 < lgMaxCount Then
           istrData3 = istrData3 & iRowStr & Chr(11) & Chr(12)
        Else
           lgPageNo_C = lgPageNo_C + 1
           Exit Do
        End If
        rs0.MoveNext
	Loop

    If iLoopCount <= lgMaxCount Then                                      '��: Check if next data exists
       lgPageNo_C = ""
    End If
    rs0.Close                                                       '��: Close recordset object
    Set rs0 = Nothing	                                            '��: Release ADF
End Sub

'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data
'============================================================================================================
Sub SubBizSaveMulti()

	On Error Resume Next
	Err.Clear

	Dim pPY6G225		'�� pS13111
	Dim iErrorPosition

	On Error Resume Next                                                                 '��: Protect system from crashing
	Err.Clear																			 '��: Clear Error status

	Set pPY6G225 = Server.CreateObject("PY6G225.CsF_Cast_PlanMultiSvr")

	If CheckSYSTEMError(Err,True) = true then
		Exit Sub
	End If

	Dim reqtxtSpread
	reqtxtSpread = Request("txtSpread")
	Call pPY6G225.PY6_MAINT_Y_FAC_CAST_MULTI_SVR(gStrGlobalCollection, trim(reqtxtSpread), iErrorPosition)

	If CheckSYSTEMError2(Err, True, iErrorPosition & "��","","","","") = True Then
	   Set pPY6G225 = Nothing
	   Exit Sub
	End If

	Set pPY6G225 = Nothing

	Response.Write "<Script Language=vbscript>" & vbCr
	Response.Write "Parent.DBSaveOK "           & vbCr
	Response.Write "</Script>"                  & vbCr
End Sub


'============================================================================================================
' Name : SubBizSaveCreate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveMultiCreate(arrColVal)
On Error Resume Next                                                             '��: Protect system from crashing
Err.Clear                                                                        '��: Clear Error status

'---------- Developer Coding part (Start) ---------------------------------------------------------------
'A developer must define field to create record
'--------------------------------------------------------------------------------------------------------

'---------- Developer Coding part (End  ) ---------------------------------------------------------------
End Sub
'============================================================================================================
' Name : SubBizSaveMultiUpdate
' Desc : Update Data from Db
'============================================================================================================
Sub SubBizSaveMultiUpdate(arrColVal)

On Error Resume Next                                                             '��: Protect system from crashing
Err.Clear                                                                        '��: Clear Error status

'---------- Developer Coding part (Start) ---------------------------------------------------------------
'A developer must define field to update record
'--------------------------------------------------------------------------------------------------------

'---------- Developer Coding part (End  ) ---------------------------------------------------------------
End Sub
'============================================================================================================
' Name : SubBizSaveMultiDelete
' Desc : Delete Data from Db
'============================================================================================================
Sub SubBizSaveMultiDelete(arrColVal)

On Error Resume Next                                                             '��: Protect system from crashing
Err.Clear                                                                        '��: Clear Error status

'---------- Developer Coding part (Start) ---------------------------------------------------------------
'A developer must define field to update record
'--------------------------------------------------------------------------------------------------------

'---------- Developer Coding part (End  ) ---------------------------------------------------------------
End Sub
'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements(pDataType,arrColVal)
Dim iSelCount

On Error Resume Next

'------ Developer Coding part (Start ) ------------------------------------------------------------------
'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub
'============================================================================================================
' Name : CommonOnTransactionCommit
' Desc : This Sub is called by OnTransactionCommit Error handler
'============================================================================================================
Sub CommonOnTransactionCommit()
'------ Developer Coding part (Start ) ------------------------------------------------------------------
'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub

'============================================================================================================
' Name : CommonOnTransactionAbort
' Desc : This Sub is called by OnTransactionAbort Error handler
'============================================================================================================
Sub CommonOnTransactionAbort()
'------ Developer Coding part (Start ) ------------------------------------------------------------------
'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub

'============================================================================================================
' Name : SetErrorStatus
' Desc : This Sub set error status
'============================================================================================================
Sub SetErrorStatus()
'------ Developer Coding part (Start ) ------------------------------------------------------------------
'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub
'============================================================================================================
' Name : SubHandleError
' Desc : This Sub handle error
'============================================================================================================
Sub SubHandleError(pOpCode,pConn,pRs,pErr)
On Error Resume Next                                                             '��: Protect system from crashing
Err.Clear                                                                        '��: Clear Error status


End Sub
'==============================================================================
' Function : SheetFocus
' Description : �����߻��� Spread Sheet�� ��Ŀ���� 
'==============================================================================
Function SheetFocus(Byval lRow, Byval lCol, Byval iLoc)

If Trim(lRow) = "" Then Exit Function
If iLoc = I_INSCRIPT Then
	strHTML = "parent.frm1.vspdData1.focus" & vbCrLf
	strHTML = strHTML & "parent.frm1.vspdData1.Row = " & lRow & vbCrLf
	strHTML = strHTML & "parent.frm1.vspdData1.Col = " & lCol & vbCrLf
	strHTML = strHTML & "parent.frm1.vspdData1.Action = 0" & vbCrLf
	strHTML = strHTML & "parent.frm1.vspdData1.SelStart = 0 " & vbCrLf
	strHTML = strHTML & "parent.frm1.vspdData1.SelLength = len(parent.frm1.vspdData1.Text) " & vbCrLf
	Response.Write strHTML
ElseIf iLoc = I_MKSCRIPT Then
	strHTML = "<" & "Script LANGUAGE=VBScript" & ">" & vbCrLf
	strHTML = strHTML & "parent.frm1.vspdData1.focus" & vbCrLf
	strHTML = strHTML & "parent.frm1.vspdData1.Row = " & lRow & vbCrLf
	strHTML = strHTML & "parent.frm1.vspdData1.Col = " & lCol & vbCrLf
	strHTML = strHTML & "parent.frm1.vspdData1.Action = 0" & vbCrLf
	strHTML = strHTML & "parent.frm1.vspdData1.SelStart = 0 " & vbCrLf
	strHTML = strHTML & "parent.frm1.vspdData1.SelLength = len(parent.frm1.vspdData1.Text) " & vbCrLf
	strHTML = strHTML & "</" & "Script" & ">" & vbCrLf
	Response.Write strHTML
End If
End Function

%>
