
<%@ LANGUAGE=VBSCript %>

<%Option Explicit%>
<!-- #Include file="../inc/incSvrMain.asp" -->
<!-- #Include file="../inc/incSvrDate.inc" -->
<!-- #Include file="../inc/incSvrNumber.inc" -->
<!-- #Include file="../inc/incSvrDBAgent.inc" -->
<!-- #Include file="../inc/incSvrDBAgentVariables.inc" -->
<!-- #Include file="../ComAsp/LoadInfTB19029.asp" -->
<% 

Err.Clear
On Error Resume Next


Const C_SHEETMAXROWS_D  = 30                                   '��: Server���� �ѹ��� fetch�� �ִ� ����Ÿ �Ǽ� 

Dim lgADF                                                                  '�� : ActiveX Data Factory ���� �������� 
Dim lgstrRetMsg                                                            '�� : Record Set Return Message �������� 
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1                         '�� : DBAgent Parameter ���� 
Dim lgstrData                                                              '�� : data for spreadsheet data
Dim lgStrPrevKey                                                           '�� : ���� �� 
Dim lgMaxCount                                                             '�� : �ѹ��� �����ü� �ִ� ����Ÿ �Ǽ� 
Dim lgTailList                                                             '�� : Orderby���� ���� field ����Ʈ 
Dim lgSelectList
Dim lgSelectListDT
Dim lgPageNo
'Dim lgtxtMaxRows
Dim iLoopCount, iEndRow

'--------------- ������ coding part(��������,Start)--------------------------------------------------------
Dim strCond
Dim strtempGlNo, strRefNo
Dim strMsgCd, strMsg1, strMsg2

'--------------- ������ coding part(��������,End)----------------------------------------------------------
	Call LoadBasisGlobalInf()
	Call loadInfTB19029B("Q", "A","NOCOOKIE","RB")
  
    Call HideStatusWnd 

    lgStrPrevKey   = Request("lgStrPrevKey")                               '�� : Next key flag
    lgPageNo       = UNICInt(Trim(Request("lgPageNo")),0)                  '��: "0"(First),"1"(Second),"2"(Third),"3"(...)
    lgMaxCount     = C_SHEETMAXROWS_D										   '�� : �ѹ��� �����ü� �ִ� ����Ÿ �Ǽ� 
    lgSelectList   = Request("lgSelectList")                               '�� : select ����� 
    lgSelectListDT = Split(Request("lgSelectListDT"), gColSep)             '�� : �� �ʵ��� ����Ÿ Ÿ�� 
    lgTailList     = Request("lgTailList")                                 '�� : Orderby value
	
    Call TrimData()
    Call FixUNISQLData()
    Call QueryData()
    
'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------

Sub MakeSpreadSheetData()
    Dim  RecordCnt
    Dim  ColCnt
    Dim  iCnt
    Dim  iRCnt
    Dim  iStr
    Dim iPrevEndRow

    iCnt = 0
    lgstrData = ""

	If Len(Trim(lgPageNo)) Then                                        '�� : Chnage Nextkey str into int value
       If Isnumeric(lgPageNo) Then          
          lgPageNo = CInt(lgPageNo)
       End If   
    Else   
       lgPageNo = 0
    End If   
    
    If Len(Trim(lgStrPrevKey)) Then                                        '�� : Chnage Nextkey str into int value
       If Isnumeric(lgStrPrevKey) Then
          iCnt = CInt(lgStrPrevKey)
       End If   
    End If

	iLoopCount = 0

    For iRCnt = 1 to iCnt  *  lgMaxCount                                   '�� : Discard previous data
		iLoopCount = iRCnt
        rs0.MoveNext
    Next

    rs0.PageSize     = lgMaxCount
	rs0.AbsolutePage = lgPageNo + 1
    iRCnt = -1
	iEndRow = 0
    
    Do while Not (rs0.EOF Or rs0.BOF)
        iRCnt =  iRCnt + 1
		iEndRow = iEndRow + 1
        iStr = ""
		For ColCnt = 0 To UBound(lgSelectListDT) - 1 
            iStr = iStr & Chr(11) & FormatRsString(lgSelectListDT(ColCnt),rs0(ColCnt))
		Next
 
        If  iRCnt < lgMaxCount Then
            lgstrData      = lgstrData      & iStr & Chr(11) & Chr(12)
        Else
            iCnt = iCnt + 1        
            lgStrPrevKey = CStr(iCnt)
            lgPageNo = lgPageNo + 1
            Exit Do
        End If
        rs0.MoveNext
	Loop

    If  iRCnt < lgMaxCount Then                                            '��: Check if next data exists
        lgStrPrevKey = ""                  
        lgPageNo = ""                                                 '��: ���� ����Ÿ ����.
    End If
  	
'	rs0.Close
'	Set rs0 = Nothing 
'	Set lgADF = Nothing                                                    '��: ActiveX Data Factory Object Nothing
End Sub
'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()
    Redim UNISqlId(0)                                                     '��: SQL ID ������ ���� ����Ȯ�� 
    '--------------- ������ coding part(�������,Start)----------------------------------------------------

    Redim UNIValue(0,2)

    UNISqlId(0) = "A5130RA101"

    '--------------- ������ coding part(�������,End)------------------------------------------------------
    UNIValue(0,0) = lgSelectList                                          '��: Select list
    '--------------- ������ coding part(�������,Start)----------------------------------------------------

    UNIValue(0,1) = strCond
    
    '--------------- ������ coding part(�������,End)------------------------------------------------------
    UNIValue(0,UBound(UNIValue,2)) = UCase(Trim(lgTailList))
    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '��: set ADO read mode
 
End Sub
'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
Sub QueryData()
    Dim iStr

    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)
    
    iStr = Split(lgstrRetMsg,gColSep)
    
    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If    
        
    If  rs0.EOF And rs0.BOF Then
		strMsgCd = "900014"
    Else    
        Call  MakeSpreadSheetData()
    End If

	rs0.Close
	Set rs0 = Nothing 
	Set lgADF = Nothing                                                    '��: ActiveX Data Factory Object Nothing

	If strMsgCd <> "" Then
		Call DisplayMsgBox(strMsgCd, vbInformation, strMsg1, strMsg2, I_MKSCRIPT)
		Response.End 
	End If
	
End Sub

'----------------------------------------------------------------------------------------------------------
' Set default value or preset value
'----------------------------------------------------------------------------------------------------------
Sub TrimData()
 
    '--------------- ������ coding part(�������,Start)----------------------------------------------------
    strtempGlNo  = FilterVar(Request("txttempGlNo"),"","S")
    strRefNo	 = FilterVar(Request("txtRefNo"),"","S")

	strCond = ""
	
	If strtempGlNo = "" Then 
		strCond = strCond & " ref_no = " & strRefNo & " "
		strMsg1 = Request("txtRefNo_Alt")
	Else
		strCond = strCond & " A.temp_gl_no = " & strtempGlNo & " "
		strMsg1 = Request("txttempGlNo_Alt")
	End If
	
	'==============================================
	'Gl No, Ref No�� Gl Header �б� 
	'==============================================
    Redim UNISqlId(0)                                                     '��: SQL ID ������ ���� ����Ȯ�� 
    Redim UNIValue(0,0)

    UNISqlId(0)   = "A5130RA102KO441"
    UNIValue(0,0) = strCond
    
    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '��: set ADO read mode

    Dim iStr

    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs1)
       
    iStr = Split(lgstrRetMsg,gColSep)
    
    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If    
        
    If  rs1.EOF And rs1.BOF Then		
		rs1.Close
		Set rs1 = Nothing
		Set lgADF = Nothing

		strMsgCd = "114100"
		Call DisplayMsgBox(strMsgCd, vbOKOnly, "", "", I_MKSCRIPT)
		Response.Write "<Script Language=vbScript>"	 & vbcr
		Response.Write "Parent.CancelClick()	  "	 & vbcr	
		Response.Write "</Script>				  "	 & vbcr													'��: �����Ͻ� ���� ó���� ������ 
    Else
		strTempGlNo = Trim(rs1(0))
		strRefNo = Trim(rs1(2))
		
%>
		<Script Language=vbScript>
			
			With parent
				.frm1.txttempGlNo.value    = "<%=Trim(rs1(0))%>"
				.frm1.txttempGlDt.value    = "<%=UNIDateClientFormat(rs1(1))%>"
				If IsNull("<%=ConvSPChars(rs1(2))%>") = True Then 
					.frm1.txtRefNo.value   = ""
				Else
					.frm1.txtRefNo.value   = "<%=ConvSPChars(Trim(rs1(2)))%>"
				End If
				.frm1.txtCrAmt.text   = "<%=UNINumClientFormat(rs1(3), ggAmtOfMoney.DecPoint, 0)%>"
				.frm1.txtCrLocAmt.text= "<%=UNINumClientFormat(rs1(4), ggAmtOfMoney.DecPoint, 0)%>"
				.frm1.txtDrAmt.text   = "<%=UNINumClientFormat(rs1(5), ggAmtOfMoney.DecPoint, 0)%>"
				.frm1.txtDrLocAmt.text= "<%=UNINumClientFormat(rs1(6), ggAmtOfMoney.DecPoint, 0)%>"
				.frm1.txtGlInputType.value= "<%=ConvSPChars(Trim(rs1(10)))%>"
				.frm1.txtGlInputTypeNm.value= "<%=ConvSPChars(Trim(rs1(11)))%>"
				.frm1.txtTempGlDesc.value	="<%=ConvSPChars(Trim(rs1(9)))%>"
				.frm1.txtConfFgNm.value	="<%=ConvSPChars(Trim(rs1(12)))%>"
				.frm1.hHqBrchNo.value = "<%=ConvSPChars(Trim(rs1(13)))%>" 
				
				If IsNull("<%=ConvSPChars(rs1(7))%>") = True Then
					.frm1.txtDeptCd.value = ""
					.frm1.txtDeptNm.value = ""
				Else
					.frm1.txtDeptCd.value		= "<%=ConvSPChars(Trim(rs1(7)))%>"
					.frm1.txtDeptNm.value		= "<%=ConvSPChars(Trim(rs1(8)))%>"
				End If


				'.frm1.txtConfirmEmp.value = "<%=ConvSPChars(Trim(rs1(14)))%>" 	'�׷���� ��ǥ������
				.frm1.txtConfirmEmp.value = "<%=ConvSPChars(Trim(rs1(15)))%>" 	'��ǥ ����������
				
			End With
		</Script>
<%
    End If

	rs1.Close
	Set rs1 = Nothing
    '--------------- ������ coding part(�������,End)------------------------------------------------------

End Sub

%>

<Script Language=vbscript>

    With parent
         .ggoSpread.Source     = .frm1.vspdData 
         .frm1.vspdData.Redraw = False
         .ggoSpread.SSShowData "<%=lgstrData%>","F"                            '��: Display data 
		Call .ReFormatSpreadCellByCellByCurrency(.frm1.vspdData,<%=iLoopCount%>,<%=iLoopCount + iEndRow%>,.GetKeyPos("A",5),.GetKeyPos("A",3),"A", "Q" ,"X","X")
         .lgPageNo_A		   =  "<%=lgPageNo%>"
         .lgStrPrevKey         = "<%=lgStrPrevKey%>"                      
                  
         Call .DbQueryOk(1)
         .frm1.vspdData.Redraw = True
	End with
</Script>	

<%
	Response.End 
%>