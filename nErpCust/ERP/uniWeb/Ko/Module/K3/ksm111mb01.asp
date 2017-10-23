<%@ LANGUAGE="VBSCRIPT" %>
<%Option Explicit    %>
<!--
'**********************************************************************************************
'*  1. Module Name          : MC
'*  2. Function Name        :
'*  3. Program ID           : sm111mb01
'*  4. Program Name         : ��Ƽ���۴ϼ��ֵ�� 
'*  5. Program Desc         : ��Ƽ���۴ϼ��ֵ��-��Ƽ 
'*  6. Component List       :
'*  7. Modified date(First) : 2005/01/24
'*  8. Modified date(Last)  :
'*  9. Modifier (First)     : Sim Hae Young
'* 10. Modifier (Last)      :
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              :
'**********************************************************************************************
-->
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<%
call LoadBasisGlobalInf()
call LoadInfTB19029B("I", "*","NOCOOKIE","MB")
call LoadBNumericFormatB("I","*","NOCOOKIE","MB")
	Dim lgOpModeCRUD

	Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0                 '�� : DBAgent Parameter ���� 
	Dim rs1, rs2, rs3, rs4,rs5
	Dim istrData
	Dim iStrPrNo
	Dim StrNextKey		' ���� �� 
	Dim lgStrPrevKey	' ���� �� 
	Dim iLngMaxRow		' ���� �׸����� �ִ�Row
	Dim iLngRow
	Dim GroupCount
	Dim lgCurrency
	Dim index,Count     ' ���� �� Return ���� ���� ������ ���� ���� 
	Dim lgDataExist
	Dim lgPageNo
	Dim sRow
	Dim lglngHiddenRows
	Dim MaxRow2
	Dim MaxCount

	Const C_SHEETMAXROWS_D  = 100

	On Error Resume Next                                                             '��: Protect system from crashing
	Err.Clear                                                                        '��: Clear Error status

	Call HideStatusWnd                                                               '��: Hide Processing message
	'------ Developer Coding part (Start ) ------------------------------------------------------------------

	'------ Developer Coding part (End   ) ------------------------------------------------------------------

	lgOpModeCRUD  = Request("txtMode")

	'response.write lgOpModeCRUD & lgOpModeCRUD &"<br>"
	'response.write UID_M0001 & UID_M0001 &"<br>"

	'response.end

	Select Case lgOpModeCRUD
	Case CStr(UID_M0001)
	     Call  SubBizQueryMulti()
	End Select


'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()

    On Error Resume Next
	lgPageNo       = UNICInt(Trim(Request("lgPageNo1")),0)    '��: "0"(First),"1"(Second),"2"(Third),"3"(...)
	lgStrPrevKeyM  = UNICInt(Trim(Request("lgStrPrevKeyM")),0)
	lgDataExist    = "No"
	iLngMaxRow     = CLng(Request("txtMaxRows"))
	lgStrPrevKey   = Request("lgStrPrevKey")
	sRow           = CLng(Request("lRow"))
	lglngHiddenRows = CLng(Request("lglngHiddenRows"))

	Call FixUNISQLData()
	Call QueryData()

End Sub

'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
' Query�ϱ� ����  DB Agent �迭�� �̿��Ͽ� Query���� ����� ���ν��� 
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()
	Dim strVal
	Redim UNISqlId(0)                                                     '��: SQL ID ������ ���� ����Ȯ�� 
	Redim UNIValue(0,2)                                                 '��: DB-Agent�� ���۵� parameter�� ���� ���� 
	                                                                '    parameter�� ���� ���� ������ 
	UNISqlId(0) = "SM111MA102" 											' header


	UNIValue(0,0) = " '"& FilterVar(Trim(UCase(Request("strPO_COMPANY"))), " " , "SNM") & "' "
	UNIValue(0,1) = " '"& FilterVar(Trim(UCase(Request("strSO_COMPANY"))), " " , "SNM") & "' "
	UNIValue(0,2) = " '"& FilterVar(Trim(UCase(Request("strPO_NO"))), " " , "SNM") & "' "

	'--------------- ������ coding part(�������,End)------------------------------------------------------
	UNILock = DISCONNREAD :	UNIFlag = "1"                                 '��: set ADO read mode
End Sub

'----------------------------------------------------------------------------------------------------------
' Query Data
' ADO�� Record Set�̿��Ͽ� Query�� �ϰ� Record Set�� �Ѱܼ� MakeSpreadSheetData()���� Spreadsheet�� �����͸� 
' �Ѹ� 
' ADO ��ü�� �����Ҷ� prjPublic.dll������ �̿��Ѵ�.(�󼼳����� vb�� �ۼ��� prjPublic.dll �ҽ� ����)
'----------------------------------------------------------------------------------------------------------
Sub QueryData()
	Dim lgstrRetMsg                                             '�� : Record Set Return Message �������� 
	Dim lgADF                                                   '�� : ActiveX Data Factory ���� �������� 
	Dim iStr

	Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")

	lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3, rs4)

	Set lgADF   = Nothing


	'//Response.write "rs0" & rs0(0) & "<br>"
	'//Response.End

	iStr = Split(lgstrRetMsg,gColSep)

	If iStr(0) <> "0" Then
		Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
	End If

	Dim FalsechkFlg

	FalsechkFlg = False

	If  rs0.EOF And rs0.BOF And FalsechkFlg =  False Then
		'Call DisplayMsgBox("172400", vbOKOnly, iStrPrNo, "", I_MKSCRIPT)
		rs0.Close
		Set rs0 = Nothing
		Response.End
	Else
		Call  MakeSpreadSheetData()
	End If





	Response.Write "<Script Language=vbscript>" & vbCr
	Response.Write "With parent" & vbCr
	Response.Write "	.ggoSpread.Source       = .frm1.vspdData2 "			& vbCr
	Response.Write "	.ggoSpread.SSShowData     """ & istrData	 & """" & vbCr
	Response.Write "	.lgPageNo1		=  """ & lgPageNo	 & """" & vbCr
    	Response.Write "    .DbQueryOk2(" & MaxCount & ")" & vbCr
	Response.Write "End With"		& vbCr
	Response.Write "</Script>"		& vbCr



End Sub


'----------------------------------------------------------------------------------------------------------
'QueryData()�� ���ؼ� Query�� �Ǹ� MakeSpreadSheetData()�� ���ؼ� �����͸� ���������Ʈ�� �ѷ��ִ� ���ν��� 
'----------------------------------------------------------------------------------------------------------
Sub MakeSpreadSheetData()
	Const C_SHEETMAXROWS_D  = 100
	Dim iLoopCount
	Dim iRowStr
	Dim ColCnt
        DIM i


	Const C_SO_ITEM_CD	=0
	Const C_SO_ITEM_NM	=1
	Const C_SO_ITEM_SPEC	=2
	Const C_ITEM_CD		=3
	Const C_BP_ITEM_NM	=4
	Const C_BP_ITEM_SPEC	=5
	Const C_PO_QTY		=6
	Const C_PO_UNIT		=7
	Const C_PO_PRC		=8
	Const C_PO_DOC_AMT	=9
	Const C_DLVY_DT		=10
	Const C_VAT_DOC_AMT	=11
	Const C_PO_VAT_RATE	=12
	Const C_PO_VAT_TYPE	=13
	Const C_PO_VAT_TYPE_NM	=14
	Const C_PO_SEQ_NO	=15


	lgDataExist    = "Yes"
	'----- ���ڵ�� Į�� ���� ----------
	'-----------------------------------
	iLngMaxRow=0
	iLoopCount = 0
        i = 0

	If CLng(lgPageNo) > 0 Then
		rs0.Move     	= CLng(C_SHEETMAXROWS_D) * CLng(lgPageNo)                  'C_SHEETMAXROWS_D:Max Fetched Count at once , lgStrPrevKeyIndex : Previous PageNo
		MaxRow2     	= CLng(C_SHEETMAXROWS_D) * CLng(lgPageNo)                  'C_SHEETMAXROWS_D:Max Fetched Count at once , lgStrPrevKeyIndex : Previous PageNo
	End If

   	Do while Not (rs0.EOF Or rs0.BOF)
	        iLoopCount =  iLoopCount + 1
	        iRowStr = ""

	        iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(C_SO_ITEM_CD))
	        iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(C_SO_ITEM_NM))
	        iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(C_SO_ITEM_SPEC))
	        iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(C_ITEM_CD))
	        iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(C_BP_ITEM_NM))
	        iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(C_BP_ITEM_SPEC))
	        iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(C_PO_QTY))
	        iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(C_PO_UNIT))
	        iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(C_PO_PRC))
	        iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(C_PO_DOC_AMT))
	        iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(C_DLVY_DT))
	        iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(C_VAT_DOC_AMT))
	        iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(C_PO_VAT_RATE))
	        iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(C_PO_VAT_TYPE))
	        iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(C_PO_VAT_TYPE_NM))
	        iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(C_PO_SEQ_NO))
	        iRowStr = iRowStr & Chr(11) & sRow
	        iRowStr = iRowStr & Chr(11) & Trim(ConvSPChars(MaxRow2 + iLoopCount))

		iRowStr = iRowStr & Chr(11) & iLngMaxRow + iLoopCount

	        If iLoopCount - 1 < C_SHEETMAXROWS_D Then
	           istrData = istrData & iRowStr & Chr(11) & Chr(12)
	        Else
		   istrData = ""
	           lgPageNo = lgPageNo + 1

	           Exit Do
	        End If
	        rs0.MoveNext
	        i = i + 1
   	Loop


    If iLoopCount-1 < C_SHEETMAXROWS_D Then                                      '��: Check if next data exists
		lgPageNo = ""
       'lgStrPrevKeyM = ""
    End If

    MaxRow2 = MaxRow2 + iLoopCount
    MaxCount = iLoopCount
    rs0.Close                                                       '��: Close recordset object
    Set rs0 = Nothing	                                            '��: Release ADF

End Sub

%>