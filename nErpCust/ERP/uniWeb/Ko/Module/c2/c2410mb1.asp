<%Option Explicit%>
<%
'======================================================================================================
'*  1. Module Name          : COSTING
'*  2. Function Name        : ǥ�ؿ������� 
'*  3. Program ID           : c2410mb1
'*  4. Program Name         : ǥ�ؿ��� ������ ����ٰ���ȸ 
'*  5. Program Desc         : ���庰 ǥ�ذ��� ������ ���� ����ٰŸ� �Ѵ�.
'*  6. Comproxy List        : +
'*  7. Modified date(First) : 2000/09/02
'*  8. Modified date(Last)  : 2002/03/25
'*  9. Modifier (First)     : Ig Sung, Cho
'* 10. Modifier (Last)      : Jang Yoon Ki
'* 11. Comment              :
'=======================================================================================================
Response.Expires = -1								'�� : ASP�� ĳ������ �ʵ��� �Ѵ�.
Response.Buffer = True								'�� : ASP�� ���ۿ� ������� �ʰ� �ٷ� Client�� ��������.
%>
<!-- #Include file="../../inc/IncSvrMain.asp"  -->
<!-- #Include file="../../inc/IncSvrNumber.inc"  -->
<!-- #Include file="../../inc/incSvrDBAgent.inc"  -->
<!-- #Include file="../../inc/incSvrDBAgentVariables.inc"  -->

<%													'�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 

On Error Resume Next								'��: 
Call LoadBasisGlobalInf() 

Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2						'�� : DBAgent Parameter ���� 
Dim lgstrData																'�� : data for spreadsheet data
Dim lgStrPrevKey                                                           '�� : ���� �� 
Dim lgMaxCount                                                             '�� : �ѹ��� �����ü� �ִ� ����Ÿ �Ǽ� 
Dim lgTailList                                                             '�� : Orderby���� ���� field ����Ʈ 
Dim lgSelectList
Dim lgSelectListDT
Dim lgDataExist
Dim lgPageNo

'--------------- ������ coding part(��������,Start)--------------------------------------------------------
Dim txtPlantCd
Dim txtPlantNm
Dim txtItemCd
Dim txtItemNm
Dim SetFocusFlag
'--------------- ������ coding part(��������,End)----------------------------------------------------------

	Call HideStatusWnd

    lgPageNo       = Trim(Request("lgPageNo"))                  '��: "0"(First),"1"(Second),"2"(Third),"3"(...)
'    lgMaxCount     = Trim(Request("lgMaxCount"))                           '�� : �ѹ��� �����ü� �ִ� ����Ÿ �Ǽ� 
    lgSelectList   = Request("lgSelectList")                               '�� : select ����� 
    lgSelectListDT = Split(Request("lgSelectListDT"), gColSep)             '�� : �� �ʵ��� ����Ÿ Ÿ�� 
    lgTailList     = Request("lgTailList")                                 '�� : Orderby value
    lgDataExist    = "No"
 
	txtPlantCd = Trim(Request("txtPlantCd"))
	txtItemCd = Trim(Request("txtItemCd"))
	
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
    
    Const C_SHEETMAXROWS_D  = 100 
	
	lgMaxCount = CInt(C_SHEETMAXROWS_D)                     '��: Max fetched data at a time
    
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

    UNISqlId(0) = "c2410ma1"
    UNISQLID(1) = "commonqry"
    UNISQLID(2) = "commonqry"

    '--------------- ������ coding part(�������,End)------------------------------------------------------
    UNIValue(0,0) = lgSelectList                                          '��: Select list
    '--------------- ������ coding part(�������,Start)----------------------------------------------------
	strWhere = "AND   B.DI_FLAG = " & FilterVar("D", "''", "S") & " "
	strWhere = strWhere & "AND   C.DI_FLAG = " & FilterVar("I", "''", "S") & " "
    strWhere = strWhere & " and A.PLANT_CD = " & FilterVar(txtPlantCd ,"''"       ,"S")
    If txtItemCd <> "" then
		strWhere = strWhere & " and A.ITEM_CD >= " & FilterVar(txtItemCd   , "''", "S")
	end if

	UNIValue(0,1)  = strWhere

	UNIValue(1,0) = "SELECT PLANT_NM FROM B_PLANT WHERE PLANT_CD =  " & FilterVar(txtPlantCd, "''", "S") & " "

	UNIValue(2,0) = "SELECT A.Item_nm from b_item A, b_item_by_plant B WHERE A.item_cd = B.item_cd AND A.item_cd = " & FilterVar(txtItemCd   , "''", "S")

    
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
    Dim lgADF                                                                  '�� : ActiveX Data Factory ���� �������� 

    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2)
    
    Set lgADF = Nothing                                                    '��: ActiveX Data Factory Object Nothing
    
    iStr = Split(lgstrRetMsg,gColSep)
    
    'If iStr(0) <> "0" Then
    '    Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    'End If    
    
    'rs1
	If txtPlantCd <> "" Then
		If Not (rs1.EOF OR rs1.BOF) Then
			txtPlantNm = Trim(rs1("Plant_Nm"))
		Else
			txtPlantNm = ""
			Call DisplayMsgBox("125000", vbOKOnly, "", "", I_MKSCRIPT)		'No Data Found!!
	        rs1.Close
		    Set rs1 = Nothing
			Exit sub
		End IF
		rs1.Close
		Set rs1 = Nothing
	End If
    
    'rs2
    If txtItemCd <> "" Then
		If Not (rs2.EOF OR rs2.BOF) Then
			txtItemNm = Trim(rs2("Item_Nm"))
		Else
			txtItemNm = ""
		End IF
		rs2.Close
		Set rs2 = Nothing
	End If
    
    If  rs0.EOF And rs0.BOF Then
		Call DisplayMsgBox("232300", vbOKOnly, "", "", I_MKSCRIPT)		'No Data Found!!
        rs0.Close
        Set rs0 = Nothing
        Exit Sub
    Else
        Call MakeSpreadSheetData()
    End If
   
    
End Sub
%>

<Script Language=vbscript>

With Parent
		.frm1.txtPlantNm.value = "<%=ConvSPChars(txtPlantNm)%>"			'rs1 �� �ޱ� �˾����� ���ϰ� �׳� �Է������� ���־��ֱ� 
		.frm1.txtItemNm.value = "<%=ConvSPChars(txtItemNm)%>"			'rs2 �� �ޱ� �˾����� ���ϰ� �׳� �Է������� ���־��ֱ� 

	If "<%=lgDataExist%>" = "Yes" Then
				
       'Set condition data to hidden area
		If "<%=lgPageNo%>" = "1" Then   ' "1" means that this query is first and next data exists
			.Frm1.htxtPlantCd.Value			= .Frm1.txtPlantCd.Value                  'For Next Search
			.Frm1.htxtItemCd.Value			= .Frm1.txtItemCd.Value
		End If
       
       'Show multi spreadsheet data from this line

		.ggoSpread.Source  = Parent.frm1.vspdData
		.ggoSpread.SSShowData "<%=lgstrData%>"                  '�� : Display data
		.lgPageNo      =  "<%=lgPageNo%>"               '�� : Next next data tag
		.DbQueryOk
	Else		
		.Frm1.txtPlantCd.Focus		
		
    End If
    
End With

</Script>
