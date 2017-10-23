<%Option Explicit%>
<!--'======================================================================================================
'*  1. Module Name          : COSTING
'*  2. Function Name        : ������������ 
'*  3. Program ID           : c3605mB1
'*  4. Program Name         : ����������������ȸ 
'*  5. Program Desc         : ����������������ȸ 
'*  6. Comproxy List        : +
'*  7. Modified date(First) : 2002/03/25
'*  8. Modified date(Last)  : 2002/03/25
'*  9. Modifier (First)     : Jang Yoon Ki
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'======================================================================================================= -->
<!-- #Include file="../../inc/IncSvrMain.asp"  -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<!-- #Include file="../../inc/IncSvrNumber.inc"  -->
<!-- #Include file="../../inc/incSvrDBAgent.inc"  -->
<!-- #Include file="../../inc/incSvrDBAgentVariables.inc"  -->

<%													'�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 

On Error Resume Next								'��: 

Call LoadBasisGlobalInf() 
Call LoadInfTB19029B("I","*", "NOCOOKIE", "MB")       

Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2							'�� : DBAgent Parameter ���� 
Dim lgstrData																'�� : data for spreadsheet data
Dim lgStrPrevKey                                                           '�� : ���� �� 
Dim lgMaxCount                                                             '�� : �ѹ��� �����ü� �ִ� ����Ÿ �Ǽ� 
Dim lgTailList                                                             '�� : Orderby���� ���� field ����Ʈ 
Dim lgSelectList
Dim lgSelectListDT
Dim lgDataExist
Dim lgPageNo

'--------------- ������ coding part(��������,Start)--------------------------------------------------------
Dim txtYyyyMm
Dim txtCostCd
Dim txtCostNm
Dim txtBasSum
Dim txtIssueSum
Dim txtRcptSum
Dim txtBalSum
Dim SetFocusFlag
'--------------- ������ coding part(��������,End)----------------------------------------------------------

	Call HideStatusWnd

    lgPageNo       = Trim(Request("lgPageNo"))                  '��: "0"(First),"1"(Second),"2"(Third),"3"(...)
    lgMaxCount     = Trim(Request("lgMaxCount"))                           '�� : �ѹ��� �����ü� �ִ� ����Ÿ �Ǽ� 
    lgSelectList   = Request("lgSelectList")                               '�� : select ����� 
    lgSelectListDT = Split(Request("lgSelectListDT"), gColSep)             '�� : �� �ʵ��� ����Ÿ Ÿ�� 
    lgTailList     = Request("lgTailList")                                 '�� : Orderby value
    lgDataExist    = "No"
 
	txtYyyyMm = Trim(Request("txtYyyyMm"))
	txtCostCd = Trim(Request("txtCostCd"))

    Call FixUNISQLData()
    Call QueryData()

'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
Sub MakeSpreadSheetData()
    Dim  RecordCnt
    Dim  ColCnt
    Dim  iLoopCount
    Dim  iRowStr
    
    lgDataExist    = "Yes"
    lgstrData      = ""

    If UniConvNumStringToDouble(lgPageNo,0) > 0 Then
       rs0.Move     = UniConvNumStringToDouble(lgMaxCount,0) * UniConvNumStringToDouble(lgPageNo,0)                  'lgMaxCount:Max Fetched Count at once , lgStrPrevKeyIndex : Previous PageNo
    End If

    iLoopCount = -1
    
    Do while Not (rs0.EOF Or rs0.BOF)
        iLoopCount =  iLoopCount + 1
        iRowStr = ""
		For ColCnt = 0 To UBound(lgSelectListDT) - 1 
            iRowStr = iRowStr & Chr(11) & FormatRsString(lgSelectListDT(ColCnt),rs0(ColCnt))
		Next
 
        If  iLoopCount < UniConvNumStringToDouble(lgMaxCount,0) Then
            lgstrData      = lgstrData      & iRowStr & Chr(11) & Chr(12)
        Else
            lgPageNo = Cstr(UniConvNumStringToDouble(lgPageNo,0) + 1)
            Exit Do
        End If
        rs0.MoveNext
                
	Loop
		
    If  iLoopCount < UniConvNumStringToDouble(lgMaxCount,0) Then                                            '��: Check if next data exists
        lgPageNo = ""                                                  '��: ���� ����Ÿ ����.
    End If
  	
	rs0.Close
    Set rs0 = Nothing
    
End Sub

'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()
    Redim UNISqlId(2)                                                     '��: SQL ID ������ ���� ����Ȯ�� 
    '--------------- ������ coding part(�������,Start)----------------------------------------------------

	Dim strWhere
    Redim UNIValue(2,2)                                                  '��: DB-Agent�� ���۵� parameter�� ���� ���� 

    UNISqlId(0) = "C3605MA101"	'spread sheet
    UNISqlId(1) = "commonqry"	'name
    UNISqlId(2) = "C3605MA102"	'sum

    '--------------- ������ coding part(�������,End)------------------------------------------------------
    UNIValue(0,0) = lgSelectList                                          '��: Select list
    '--------------- ������ coding part(�������,Start)----------------------------------------------------

    strWhere = " and A.YYYYMM = " & FilterVar(txtYyyyMm ,"''"       ,"S")
    if txtCostCd <> "" then
		strWhere = strWhere & " and A.COST_CD >= " & FilterVar(txtCostCd   , "''", "S")
	end if
	
	UNIValue(0,1)  = strWhere

	UNIValue(1,0) = "select cost_nm from b_cost_center Where cost_cd= " & FilterVar(txtCostCd, "''", "S") & " "

	UNIValue(2,0)  = strWhere

    
    '--------------- ������ coding part(�������,End)------------------------------------------------------
    UNIValue(0,UBound(UNIValue,2)) = UCase(Trim(lgTailList))
    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '��: set ADO read mode
 
End Sub
'----------------------------------------------------------------------------------------------------------
' Query Data
'--------------------------------------------------------------------------------------------------------
Sub QueryData()

    Dim lgstrRetMsg                                                            '�� : Record Set Return Message �������� 
    Dim iStr
    Dim lgADF
                                                                      '�� : ActiveX Data Factory ���� �������� 

    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2)
    
    Set lgADF = Nothing                                                    '��: ActiveX Data Factory Object Nothing
    
    iStr = Split(lgstrRetMsg,gColSep)   
   
    'If iStr(0) <> "0" Then
    '    Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    'End If 
   
	If txtCostCd <> "" Then
		If Not (rs1.EOF OR rs1.BOF) Then
			txtCostNm = Trim(rs1("Cost_Nm"))
		'Else
			
			SetFocusFlag = 1
		'	Call DisplayMsgBox("124400", vbOKOnly, "", "", I_MKSCRIPT)		'No Data Found!!
		'	txtCostNm = ""
		'	rs1.Close
		'	Set rs1 = Nothing 
		'	Exit Sub
		End IF
		rs1.Close
		Set rs1 = Nothing 
	End If	
	

    If  rs0.EOF And rs0.BOF Then
		SetFocusFlag = 2
		Call DisplayMsgBox("233700", vbOKOnly, "", "", I_MKSCRIPT)		'No Data Found!!
        rs0.Close
        Set rs0 = Nothing
        Exit Sub
    Else    
        Call  MakeSpreadSheetData()

	    'rs2
		If Not (rs2.EOF OR rs2.BOF) Then
			txtBasSum = rs2("TOT_PRNT_BAS_AMT")
			txtIssueSum = rs2("TOT_PRNT_ISSUE_AMT")
			txtRcptSum = rs2("TOT_PRNT_RCPT_AMT")
			txtBalSum = rs2("TOT_PRNT_BAL_AMT")
		Else
			txtBasSum = 0
			txtIssueSum = 0
			txtRcptSum = 0
			txtBalSum = 0
		End IF
		rs2.Close
		Set rs2 = Nothing
    End If
    
End Sub

%>

<Script Language=vbscript>

With Parent
	.frm1.txtCostNm.value				= "<%=ConvSPChars(txtCostNm)%>"

	If "<%=lgDataExist%>" = "Yes" Then

       'Set condition data to hidden area
		If "<%=lgPageNo%>" = "1" Then   ' "1" means that this query is first and next data exists
			.Frm1.hYyyyMm.Value	  = "<%=txtYyyyMm%>"                'For Next Search
			.Frm1.hCostCd.Value	  = .Frm1.txtCostCd.Value                  'For Next Search			
		End If

						'rs1 �� �ޱ� �˾����� ���ϰ� �׳� �Է������� ���־��ֱ� 
		.frm1.txtBasSum.text				= "<%=UNINumClientFormat(txtBasSum,ggAmtOfMoney.Decpoint,0)%>"				'���ʱݾ� �հ� 
		
		.frm1.txtIssueSum.text				= "<%=UNINumClientFormat(txtIssueSum,ggAmtOfMoney.Decpoint,0)%>"			'���ݾ� �հ� 
		
		.frm1.txtRcptSum.text				= "<%=UNINumClientFormat(txtRcptSum,ggAmtOfMoney.Decpoint,0)%>"				'�԰�ݾ� �հ� 
		
		.frm1.txtBalSum.text				= "<%=UNINumClientFormat(txtBalSum,ggAmtOfMoney.Decpoint,0)%>"				'�⸻�ݾ� �հ� 
		
         
       'Show multi spreadsheet data from this line			
		
		.ggoSpread.Source  = Parent.frm1.vspdData
		.ggoSpread.SSShowData "<%=lgstrData%>"                  '�� : Display data
		.lgPageNo      =  "<%=lgPageNo%>"               '�� : Next next data tag
		.DbQueryOk
	Else
		.Frm1.txtCostCd.Focus				
    End If
    
End With
</Script>
