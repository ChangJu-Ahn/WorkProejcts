<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrDBAgent.inc" -->
<!-- #Include file="../../inc/incSvrDBAgentVariables.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%             

On Error Resume Next
Err.Clear

Call LoadBasisGlobalInf()
Call LoadInfTB19029B("*", "M", "NOCOOKIE", "RB")   'ggQty.DecPoint Setting...

Call HideStatusWnd 

Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0 , rs1, rs2                        '�� : DBAgent Parameter ���� 
Dim lgstrData                                                              '�� : data for spreadsheet data
Dim lgStrPrevKey                                                           '�� : ���� �� 
Dim lgMaxCount                                                             '�� : �ѹ��� �����ü� �ִ� ����Ÿ �Ǽ� 
Dim lgTailList                                                             '�� : Orderby���� ���� field ����Ʈ 
Dim lgSelectList
Dim lgSelectListDT
Dim lgDataExist
Dim lgPageNo
'--------------- ������ coding part(��������,Start)--------------------------------------------------------
Dim strFrPrDt	                                                           
Dim strToPrDt
Dim strFrPrNo	                                                           
Dim strToPrNo
Dim strBpCd
DIm strPrrcptType

Dim iPrevEndRow
Dim iEndRow	
	                                                           '�� : ������ 
Dim strCond

Dim strMsgCd
Dim strMsg1
' ���Ѱ��� �߰� 
Dim lgAuthBizAreaCd	' ����� 
Dim lgInternalCd	' ���κμ� 
Dim lgSubInternalCd	' ���κμ�(��������)
Dim lgAuthUsrID		' ���� 
'--------------- ������ coding part(��������,End)----------------------------------------------------------
  



    lgPageNo       = UNICInt(Trim(Request("lgPageNo")),0)                  '��: "0"(First),"1"(Second),"2"(Third),"3"(...)
    lgMaxCount     = CInt(Request("lgMaxCount"))                           '�� : �ѹ��� �����ü� �ִ� ����Ÿ �Ǽ� 
    lgSelectList   = Request("lgSelectList")                               '�� : select ����� 
    lgSelectListDT = Split(Request("lgSelectListDT"), gColSep)             '�� : �� �ʵ��� ����Ÿ Ÿ�� 
    lgTailList     = Request("lgTailList")                                 '�� : Orderby value
    lgDataExist    = "No"
    iPrevEndRow = 0
    iEndRow = 0
        
    strFrPrNo	   = UCase(Trim(Request("txtFrPrNo")))
    strToPrNo      = UCase(Trim(Request("txtToPrNo")))
    strFrPrDt      = UNIConvDate(Trim(Request("txtFrPrDt")))
    strToPrDt      = UNIConvDate(Trim(Request("txtToPrDt")))
    strBpCd		   = UCase(Trim(Request("txtBpCd")))
	strPrrcptType  = UCase(Trim(Request("txtPrrcptType")))

	' ���Ѱ��� �߰� 
	lgAuthBizAreaCd		= Trim(Request("lgAuthBizAreaCd"))
	lgInternalCd		= Trim(Request("lgInternalCd"))
	lgSubInternalCd		= Trim(Request("lgSubInternalCd"))
	lgAuthUsrID			= Trim(Request("lgAuthUsrID"))
		   
    Call TrimData()
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

    If CDbl(lgPageNo) > 0 Then
		iPrevEndRow = CDbl(lgMaxCount) * CDbl(lgPageNo)    
       rs0.Move= iPrevEndRow                   'lgMaxCount:Max Fetched Count at once , lgStrPrevKeyIndex : Previous PageNo
    End If

    iLoopCount = -1
    
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
        iEndRow = iPrevEndRow + iLoopCount + 1
    Else
        iEndRow = iPrevEndRow + iLoopCount
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

    Redim UNIValue(2,2)                                                  '��: DB-Agent�� ���۵� parameter�� ���� ���� 

    UNISqlId(0) = "A7101RA1"
	UNISqlId(1) = "ABPNM"
	UNISqlId(2) = "Commonqry"

    '--------------- ������ coding part(�������,End)------------------------------------------------------
    UNIValue(0,0) = lgSelectList                                          '��: Select list
    '--------------- ������ coding part(�������,Start)----------------------------------------------------

    UNIValue(0,1) = strCond
    
    UNIValue(1,0) = UCase(" " & FilterVar(strBpCd, "''", "S") & " " )    
    UNIValue(2,0) = "select jnl_nm from A_JNL_ITEM where jnl_cd =  " & FilterVar(strPrrcptType, "''", "S") & " " 

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
    
    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If    
        
	If (rs1.EOF And rs1.BOF) Then
		If strMsgCd = "" And strBpCd <> "" Then
			strMsgCd = "970000"		'Not Found
			strMsg1 = Request("txtBPCd_Alt")
		End If
%>
		<Script Language=vbScript>
		With parent
			.Frm1.txtBpNm.value = ""
		End With
		</Script>
<%
	Else
%>
		<Script Language=vbScript>
		With parent
			.frm1.txtBPCd.value = "<%=Trim(rs1(0))%>"
			.frm1.txtBPNm.value = "<%=Trim(rs1(1))%>"
		End With
		</Script>
<%
    End If
    rs1.Close
	Set rs1 = Nothing 
        
    If  Trim(strMsgCd) <> "" Then
		Call DisplayMsgBox(strMsgCd, vbOKOnly, strMsg1, "", I_MKSCRIPT)
		Response.End													'��: �����Ͻ� ���� ó���� ������ 
    End If
 
	If strPrrcptType <> "" Then
		If Not (rs2.EOF OR rs2.BOF) Then
%>
		<Script Language=vbScript>
			With parent
				
				.frm1.txtPrrcptTypeNm.value = "<%=Trim(rs2(0))%>"
			End With
		</Script>
<%		
		Else
%>
		<Script Language=vbScript>
			With parent
				.frm1.txtPrrcptTypeNm.value = ""
			End With
		</Script>
<%		
			Call DisplayMsgBox("141800", vbOKOnly, "", "", I_MKSCRIPT)		'No Data Found!!
	        rs2.Close
		    Set rs2 = Nothing
			Exit sub
		End IF
		rs2.Close
		Set rs2 = Nothing
	End If
	
	
    If  rs0.EOF And rs0.BOF Then
		Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)		'No Data Found!!
        rs0.Close
        Set rs0 = Nothing
        Exit Sub
    Else    
        Call  MakeSpreadSheetData()
    End If

End Sub

'----------------------------------------------------------------------------------------------------------
' Set default value or preset value
'----------------------------------------------------------------------------------------------------------
Sub TrimData()
 
    '--------------- ������ coding part(�������,Start)----------------------------------------------------

     If strFrPrNo <> "" Then strCond = strCond & " and A.PRRCPT_NO >= " & Parent.FilterVar(strFrPrNo, "''", "S")

     If strToPrNo <> "" Then strCond = strCond & " and A.PRRCPT_NO <= " & Parent.FilterVar(strToPrNo, "''", "S")

     If strToPrDt <> "" Then strCond = strCond & " and A.PRRCPT_DT <=  " & FilterVar(strToPrDt , "''", "S") & ""

     If strFrPrDt <> "" Then strCond = strCond & " and A.PRRCPT_DT >=  " & FilterVar(strFrPrDt , "''", "S") & ""

     If strBpCd <> "" Then strCond = strCond & " and A.BP_CD =  " & FilterVar(strBpCd, "''", "S") & " "

     If strPrrcptType <> "" Then strCond = strCond & "  and a.prrcpt_type = " & FilterVar(strPrrcptType , "''", "S")

     strCond = strCond & " and A.PRRCPT_FG = " & FilterVar("PC", "''", "S") & ""

	' ���Ѱ��� �߰� 
	If lgAuthBizAreaCd <> "" Then
		strCond		= strCond & " AND A.BIZ_AREA_CD = " & FilterVar(lgAuthBizAreaCd, "''", "S")
	End If
	
	If lgInternalCd <> "" Then
		strCond		= strCond & " AND A.INTERNAL_CD = " & FilterVar(lgInternalCd, "''", "S")
	End If
	
	If lgSubInternalCd <> "" Then
		strCond		= strCond & " AND A.INTERNAL_CD LIKE " & FilterVar(lgSubInternalCd & "%", "''", "S")
	End If
	
	If lgAuthUsrID <> "" Then
		strCond		= strCond & " AND A.INSRT_USER_ID = " & FilterVar(lgAuthUsrID, "''", "S")
	End If     

    '--------------- ������ coding part(�������,End)------------------------------------------------------

End Sub

%>

<Script Language=vbscript>
    If "<%=lgDataExist%>" = "Yes" Then

       'Set condition data to hidden area
       If "<%=lgPageNo%>" = "1" Then   ' "1" means that this query is first and next data exists
			Parent.Frm1.htxtFrPrDt.Value    = Parent.Frm1.txtFrPrDt.Text
			Parent.Frm1.htxtToPrDt.Value    = Parent.Frm1.txtToPrDt.Text
			Parent.Frm1.htxtFrPrNo.Value	= Parent.Frm1.txtFrPrNo.Value
			Parent.Frm1.htxtToPrNo.Value	= Parent.Frm1.txtToPrNo.Value
			Parent.Frm1.htxtBPCd.Value		= Parent.Frm1.txtBPCd.Value
			Parent.Frm1.htxtPrrcptType.Value= Parent.Frm1.txtPrrcptType.Value
       End If
       'Show multi spreadsheet data from this line
			Parent.Parent.ggoSpread.Source  = Parent.frm1.vspdData
			Parent.frm1.vspdData.Redraw = False
			Parent.ggoSpread.SSShowData "<%=lgstrData%>", "F"                    '�� : Display data
			Call Parent.ReFormatSpreadCellByCellByCurrency(Parent.Frm1.vspdData,<%=iPrevEndRow+1%>,<%=iEndRow%>,Parent.GetKeyPos("A",2),Parent.GetKeyPos("A",3),"A", "Q" ,"X","X")
			Call Parent.ReFormatSpreadCellByCellByCurrency(Parent.Frm1.vspdData,<%=iPrevEndRow+1%>,<%=iEndRow%>,Parent.GetKeyPos("A",2),Parent.GetKeyPos("A",4),"A", "Q" ,"X","X")
			Parent.frm1.vspdData.Redraw = True
			Parent.lgPageNo      =  "<%=lgPageNo%>"							'�� : Next next data tag

			Parent.DbQueryOk
    End If   

</Script>	

