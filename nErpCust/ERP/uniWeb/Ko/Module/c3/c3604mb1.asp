<%Option Explicit%>
<!--'======================================================================================================
'*  1. Module Name          : COSTING
'*  2. Function Name        : ������������ 
'*  3. Program ID           : c3604mB1
'*  4. Program Name         : ȸ�谡���� ���� ��ȸ 
'*  5. Program Desc         : ȸ�谡���� ���� ��ȸ 
'*  6. Comproxy List        : +
'*  7. Modified date(First) : 2002/03/25
'*  8. Modified date(Last)  : 2002/03/25
'*  9. Modifier (First)     : Eun Hee, Hwang
'* 10. Modifier (Last)      : Eun Hee, Hwang
'* 11. Comment              :
'======================================================================================================= -->

<!-- #Include file="../../inc/IncSvrMain.asp"  -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<!-- #Include file="../../inc/IncSvrNumber.inc"  -->
<!-- #Include file="../../inc/incSvrDBAgent.inc"  -->
<!-- #Include file="../../inc/incSvrDBAgentVariables.inc"  -->

<%                                                                         '�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 


On Error Resume Next

Call LoadBasisGlobalInf() 
Call LoadInfTB19029B("I","*", "NOCOOKIE", "MB")       

Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2                             '�� : DBAgent Parameter ���� 
Dim lgstrData                                                              '�� : data for spreadsheet data
Dim lgStrPrevKey                                                           '�� : ���� �� 
Dim lgMaxCount                                                             '�� : �ѹ��� �����ü� �ִ� ����Ÿ �Ǽ� 
Dim lgTailList                                                             '�� : Orderby���� ���� field ����Ʈ 
Dim lgSelectList
Dim lgSelectListDT
Dim lgDataExist
Dim lgPageNo
'--------------- ������ coding part(��������,Start)--------------------------------------------------------
DIM COST_NM
DIM SUM,SUM1,SUM2


 
'--------------- ������ coding part(��������,End)----------------------------------------------------------
  
    Call HideStatusWnd 


    lgPageNo       = UNICInt(Trim(Request("lgPageNo")),0)                  '��: "0"(First),"1"(Second),"2"(Third),"3"(...)
'   lgMaxCount     = CInt(Request("lgMaxCount"))                           '�� : �ѹ��� �����ü� �ִ� ����Ÿ �Ǽ� 
    lgSelectList   = Request("lgSelectList")                               '�� : select ����� 
    lgSelectListDT = Split(Request("lgSelectListDT"), gColSep)             '�� : �� �ʵ��� ����Ÿ Ÿ�� 
    lgTailList     = Request("lgTailList")                                 '�� : Orderby value
    lgDataExist    = "No"
 
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
    
  	Const C_SHEETMAXROWS_D  = 100  
    lgMaxCount = CInt(C_SHEETMAXROWS_D)                     '��: Max fetched data at a time          

    If UniConvNumStringToDouble(lgPageNo,0) > 0 Then
       rs0.Move     = UniConvNumStringToDouble(lgMaxCount,0) * UniConvNumStringToDouble(lgPageNo,0)                  'lgMaxCount:Max Fetched Count at once , lgStrPrevKeyIndex : Previous PageNo
    End If

    iLoopCount = -1
     
    'rs0�� ���� ��� 
    Do while Not (rs0.EOF Or rs0.BOF)
        iLoopCount =  iLoopCount + 1
        iRowStr = ""
		For ColCnt = 0 To UBound(lgSelectListDT) - 1 
            iRowStr = iRowStr & Chr(11) & FormatRsString(lgSelectListDT(ColCnt),rs0(ColCnt))
		Next
 
        If  iLoopCount < lgMaxCount Then
            lgstrData      = lgstrData      & iRowStr & Chr(11) & Chr(12)
        Else
            lgPageNo = lgPageNo + 1
            Exit Do
        End If
        rs0.MoveNext
	Loop

    If  iLoopCount < lgMaxCount Then                                            '��: Check if next data exists
        lgPageNo = ""                                                  '��: ���� ����Ÿ ����.
    End If
  	
	rs0.Close
    Set rs0 = Nothing 
    
      
End Sub
'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()
	DIm strWhere
	
    Redim UNISqlId(2)                                                     '��: SQL ID ������ ���� ����Ȯ�� 
    '--------------- ������ coding part(�������,Start)----------------------------------------------------

    Redim UNIValue(2,6)                                                  '��: DB-Agent�� ���۵� parameter�� ���� ���� 

    UNISqlId(0) = "C3604MA01"
    UNISqlId(1) = "C3604MA02"					'SUM
    UNISqlId(2) = "COMMONQRY"					'COST_NM

    '--------------- ������ coding part(�������,End)------------------------------------------------------
    UNIValue(0,0) = lgSelectList                                          '��: Select list
    '--------------- ������ coding part(�������,Start)----------------------------------------------------
	
	strWhere = " AND A.YYYYMM = "
	strWhere = strWhere & FilterVar(Request("txtYyyymm"),"''"       ,"S")
	
	IF Request("txtCostCd")<>"" then
		strWhere = strWhere & " AND A.COST_CD = "
		strWhere = strWhere & FilterVar(Request("txtCostCd"), "''", "S")  
	End If
	
	UNIValue(0,1)  = strWhere
	UNIValue(0,2)  = strWhere
	UNIValue(0,3)  = strWhere
	UNIValue(0,4)  = strWhere
	UNIValue(0,5)  = strWhere
	
	'rs1�� ���� Value�� setting(�Ѱ�)
	UNIValue(1,0)  = strWhere
	UNIValue(1,1)  = strWhere
	UNIValue(1,2)  = strWhere
	UNIValue(1,3)  = strWhere
	UNIValue(1,4)  = strWhere
	
	'rs2�� ���� Value�� setting(�ڽ�Ʈ��Ÿ��)
	UNIValue(2,0)  = "SELECT COST_NM FROM B_COST_CENTER WHERE COST_CD= " & FilterVar(Request("txtCostCd"), "''", "S") & " "
	
    '--------------- ������ coding part(�������,End)------------------------------------------------------
    UNIValue(0,UBound(UNIValue,2)) = UCase(Trim(lgTailList))
    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '��: set ADO read mode
 
End Sub
'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
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
    '    Call ServerMesgBox(iStr(1), vbInformation, I_MKSCRIPT)
    'End If 
    
      'rs1�� ���� ��� 
    IF NOT (rs1.EOF or rs1.BOF) then
		SUM = rs1(0)				' SUM(A.AMT)
		SUM1 = rs1(1)
		Sum2 = rs1(2)
	ELSE
		SUM = 0
		SUM1 = 0
		SUM2 = 0
	End if
	
    rs1.Close
    Set rs1 = Nothing 
    
    'rs2�� ���� ��� 
	IF Request("txtCostCd")<>"" then    
		IF NOT (rs2.EOF or rs2.BOF) then
		    COST_NM = rs2("COST_NM")
		ELSE
			Call DisplayMsgBox("124400", vbOKOnly, "", "", I_MKSCRIPT)		'No Data Found!!
		%>
			<Script Language=vbscript>
				parent.frm1.txtCostCd.focus
			</Script>
		<%
		    rs2.Close
		    Set rs2 = Nothing 
		    Exit Sub
		END IF
		rs2.Close
		Set rs2 = Nothing
	End If
 
 
    If  rs0.EOF And rs0.BOF Then
		Call DisplayMsgBox("233300", vbOKOnly, "", "", I_MKSCRIPT)		'No Data Found!!
	%>
		<Script Language=vbscript>
			parent.frm1.txtYyyymm.focus
		</Script>
	<%
        rs0.Close
        Set rs0 = Nothing 
        Exit Sub
    Else    

        Call  MakeSpreadSheetData()
    End If
  
End Sub

%>

<Script Language=vbscript>
	With Parent
		If "<%=lgDataExist%>" = "Yes" Then
		   'Set condition data to hidden area
		   If "<%=lgPageNo%>" = "1" Then   ' "1" means that this query is first and next data exists
				.frm1.hYyyymm.value = "<%=Request("txtYyyymm")%>"
				.frm1.hCostCd.value = "<%=Request("txtCostCd")%>"
		   End If
		   'Show multi spreadsheet data from this line
		   .ggoSpread.Source  = Parent.frm1.vspdData
		   .ggoSpread.SSShowData "<%=lgstrData%>"                  '�� : Display data
		   .lgPageNo      =  "<%=lgPageNo%>"  '�� : Next next data tag
		   .frm1.txtSum.text = "<%=UniNumClientFormat(SUM,ggAmtOfMoney.Decpoint,0)%>" 
		   .frm1.txtMfcSum.text = "<%=UniNumClientFormat(SUM1,ggAmtOfMoney.Decpoint,0)%>" 
		   .frm1.txtDirSum.text = "<%=UniNumClientFormat(SUM2,ggAmtOfMoney.Decpoint,0)%>" 
		   .DbQueryOk
		End If   
		
		.frm1.txtCostNm.value = "<%=ConvSPChars(COST_NM)%>" 
		
	End With
</Script>	



</Script>	



