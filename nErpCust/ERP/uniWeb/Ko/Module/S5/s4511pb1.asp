<%
'**********************************************************************************************
'*  1. Module Name          : ���� 
'*  2. Function Name        : ���Ͽ�û���� 
'*  3. Program ID           : S4511PA1
'*  4. Program Name         : ���Ͽ�û������ȣ �˾� 
'*  5. Program Desc         : ADO Query
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2000/12/09
'*  8. Modified date(Last)  : 2001/12/19
'*  9. Modifier (First)     : Byun Jee Hyun
'* 10. Modifier (Last)      : Kim Hyungsuk
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              :
'*                            2000/12/09
'*                            2001/12/19	Dateǥ������ 
'**********************************************************************************************
%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDBAgent.inc" -->
<%                                                                         '�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 

On Error Resume Next

Dim lgADF                                                                  '�� : ActiveX Data Factory ���� �������� 
Dim lgstrRetMsg                                                            '�� : Record Set Return Message �������� 
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0                              '�� : DBAgent Parameter ���� 
Dim lgstrData                                                              '�� : data for spreadsheet data
Dim lgStrPrevKey                                                           '�� : ���� �� 
Dim lgMaxCount                                                             '�� : �ѹ��� �����ü� �ִ� ����Ÿ �Ǽ� 
Dim lgTailList                                                             '�� : Orderby���� ���� field ����Ʈ 
Dim lgSelectList
Dim lgSelectListDT
Dim lgDataExist
Dim lgPageNo

'--------------- ������ coding part(��������,Start)--------------------------------------------------------
Dim strPoType	                                                           '�� : �������� 
Dim strPoFrDt	                                                           '�� : ������ 
Dim strPoToDt	                                                           '�� :
Dim strSpplCd	                                                           '�� : ����ó 
Dim strPurGrpCd	                                                           '�� : ���ű׷� 
Dim strItemCd	                                                           '�� : ǰ�� 
Dim strTrackNo	                                                           '�� : Tracking No
Dim BlankchkFlg
'----------------------- �߰��� �κ� ----------------------------------------------------------------------
Dim arrRsVal(5)								'* : ȭ�鿡 ��ȸ�ؿ� Name�� ��Ƴ��� ���� ���� Array
'----------------------------------------------------------------------------------------------------------
'--------------- ������ coding part(��������,End)----------------------------------------------------------
  
	Call LoadBasisGlobalInf()
    Call HideStatusWnd 

	lgPageNo       = UNICInt(Trim(Request("lgPageNo")),0)    '��: "0"(First),"1"(Second),"2"(Third),"3"(...)
    lgStrPrevKey   = Request("lgStrPrevKey")                               '�� : Next key flag
    lgMaxCount     = 30							                           '�� : �ѹ��� �����ü� �ִ� ����Ÿ �Ǽ� 
    lgSelectList   = Request("lgSelectList")                               '�� : select ����� 
    lgSelectListDT = Split(Request("lgSelectListDT"), gColSep)             '�� : �� �ʵ��� ����Ÿ Ÿ�� 
    lgTailList     = Request("lgTailList")                                 '�� : Orderby value
    lgDataExist    = "No"

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

    iCnt = 0
    lgstrData = ""
    lgDataExist    = "Yes"
	
	If CLng(lgPageNo) > 0 Then
       rs0.Move     = CLng(lgMaxCount) * CLng(lgPageNo)                  'lgMaxCount:Max Fetched Count at once , lgStrPrevKeyIndex : Previous PageNo
    End If
    
    For iRCnt = 1 to iCnt  *  lgMaxCount                                   '�� : Discard previous data
        rs0.MoveNext
    Next

    iRCnt = -1
    
    Do while Not (rs0.EOF Or rs0.BOF)
        iRCnt =  iRCnt + 1
        iStr = ""
		For ColCnt = 0 To UBound(lgSelectListDT) - 1 
            Select Case  lgSelectListDT(ColCnt)
               Case "DD"   '��¥ 
					iStr = iStr & Chr(11) & UNIDateClientFormat(rs0(ColCnt))
               Case "F2"  ' �ݾ� 
                           iStr = iStr & Chr(11) & UNINumClientFormat(rs0(ColCnt), ggAmtOfMoney.DecPoint, 0)
               Case "F3"  '���� 
                           iStr = iStr & Chr(11) & UNINumClientFormat(rs0(ColCnt), ggQty.DecPoint       , 0)
               Case "F4"  '�ܰ� 
                           iStr = iStr & Chr(11) & UNINumClientFormat(rs0(ColCnt), ggUnitCost.DecPoint  , 0)
               Case "F5"   'ȯ�� 
                           iStr = iStr & Chr(11) & UNINumClientFormat(rs0(ColCnt), ggExchRate.DecPoint  , 0)
               Case Else
                    iStr = iStr & Chr(11) & ConvSPChars(rs0(ColCnt))
            End Select
		Next
 
        If  iRCnt < lgMaxCount Then
            lgstrData      = lgstrData      & iStr & Chr(11) & Chr(12)
        Else
            iCnt = iCnt + 1
'            lgStrPrevKey = CStr(iCnt)
            lgPageNo = lgPageNo + 1
            Exit Do
        End If
        rs0.MoveNext
	Loop

    If  iRCnt < lgMaxCount Then                                            '��: Check if next data exists
'        lgStrPrevKey = ""
        lgPageNo = ""                                                  '��: ���� ����Ÿ ����.
    End If
  	
	rs0.Close
    Set rs0 = Nothing 
    Set lgADF = Nothing                                                    '��: ActiveX Data Factory Object Nothing
End Sub
'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()
    Dim strVal
    Dim arrVal(2)															
    Dim MajorCd
    Redim UNISqlId(3)                                                     '��: SQL ID ������ ���� ����Ȯ�� 
    '--------------- ������ coding part(�������,Start)----------------------------------------------------

    Redim UNIValue(3,2)

	If Len(Request("txtCCNO")) Then
	    UNISqlId(0) = "S4511PA102"									'* : ������ ��ȸ�� ���� SQL�� 
	Else
	    UNISqlId(0) = "S4511PA101"									'* : ������ ��ȸ�� ���� SQL�� 
	End If

    UNISqlId(1) = "S0000QA002"									'* : ������ ��ȸ���Ǻθ��� Name �� �������� SQL ���� ���� 
    UNISqlId(2) = "S0000QA005"
    UNISqlId(3) = "S0000QA000"
 
    '--------------- ������ coding part(�������,End)------------------------------------------------------
    UNIValue(0,0) = lgSelectList                                          '��: Select list
    '--------------- ������ coding part(�������,Start)----------------------------------------------------
	strVal = " "

	If Len(Request("txtBpCd")) Then
		strVal = " AND SHIP_TO_PARTY = " & FilterVar(Request("txtBpCd"), "''", "S") & " "		
	Else
		strVal = ""
	End If
	arrVal(0) = FilterVar(Trim(Request("txtBpCd")), "", "S")

	If Len(Request("txtSalesGroup")) Then		
		strVal = strVal & " AND SALES_GRP = " & FilterVar(Request("txtSalesGroup"), "''", "S") & " "		
	End If	
	arrVal(1) = FilterVar(Trim(Request("txtSalesGroup")), "", "S")
	
 	If Len(Request("txtMovType")) Then		
		strVal = strVal & " AND MOV_TYPE = " & FilterVar(Request("txtMovType"), "''", "S") & " "		
		MajorCd = "I0001"
	End If
	arrVal(2) = FilterVar(Trim(Request("txtMovType")), "", "S")		
    
	If Trim(Request("txtRadio")) = "Y" Then
		strVal = strVal & " AND POST_FLAG =" & FilterVar("Y", "''", "S") & " "
	ElseIf Trim(Request("txtRadio")) = "N" Then
		strVal = strVal & " AND POST_FLAG =" & FilterVar("N", "''", "S") & " "
	End If			
		
	If Len(Request("txtDnReqNo")) Then
		strVal = strVal & " AND DN_REQ_NO >= " & FilterVar(Request("txtDnReqNo"), "''", "S") & " "		
	End If	

	If Len(Request("txtSCMDoNo")) Then
		strVal = strVal & " AND SCM_DO_NO = " & FilterVar(Request("txtSCMDoNo"), "''", "S") & " "		
	End If	


	If Len(Request("txtSONO")) Then
		strVal = strVal & " AND SO_NO = " & FilterVar(Request("txtSONO"), "''", "S") & " "		
	End If	

	If Len(Request("txtCCNo")) Then
		strVal = strVal & " AND CC_NO = " & FilterVar(Request("txtCCNo"), "''", "S") & " "		
	End If	

    If Len(Request("txtDlvyFrDt")) Then
		strVal = strVal & " AND DLVY_DT >= " & FilterVar(UNIConvDate(Request("txtDlvyFrDt")), "''", "S") & ""		
	End If		
	
	If Len(Request("txtDlvyToDt")) Then
		strVal = strVal & " AND DLVY_DT <= " & FilterVar(UNIConvDate(Request("txtDlvyToDt")), "''", "S") & ""		
	End If

    If Len(Request("txtDNFrDt")) Then
		strVal = strVal & " AND PROMISE_DT >= " & FilterVar(UNIConvDate(Request("txtDNFrDt")), "''", "S") & ""		
	End If		
	
	If Len(Request("txtDNToDt")) Then
		strVal = strVal & " AND PROMISE_DT <= " & FilterVar(UNIConvDate(Request("txtDNToDt")), "''", "S") & ""		
	End If
	

	If Trim(Request("txtExceptFlag")) = "Y" Then
		strVal = strVal & " AND EXCEPT_DN_FLAG = " & FilterVar("Y", "''", "S") & " "
	ElseIf Trim(Request("txtExceptFlag")) = "N" Then
		strVal = strVal & " AND EXCEPT_DN_FLAG <> " & FilterVar("Y", "''", "S") & " "
	End If	

    UNIValue(0,1) = strVal   '---������ 
    UNIValue(1,0) = arrVal(0)  
    UNIValue(2,0) = arrVal(1)  
    UNIValue(3,0) = FilterVar(MajorCd , "''", "S") 
    UNIValue(3,1) = arrVal(2)  
   
'================================================   
   
    '--------------- ������ coding part(�������,End)------------------------------------------------------
    UNIValue(0,UBound(UNIValue,2)) = UCase(Trim(lgTailList))
    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '��: set ADO read mode
 
End Sub
'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
Sub QueryData()
    Dim iStr
	BlankchkFlg = False
	
    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3) '* : Record Set �� ���� ���� 
	Set lgADF   = Nothing
	
    iStr = Split(lgstrRetMsg,gColSep)

    If  rs1.EOF And rs1.BOF Then
        rs1.Close
        Set rs1 = Nothing

		If Len(Request("txtBpCd")) And BlankchkFlg = False Then
		   Call DisplayMsgBox("970000", vbInformation, "��ǰó", "", I_MKSCRIPT)	'��: you must release this line if you change msg into code
	       BlankchkFlg = True	
%>
<Script Language=vbscript>
			parent.frm1.txtBpCd.focus
</Script>	       	
<%	       
		End If	
    Else    
		arrRsVal(0) = rs1(0)
		arrRsVal(1) = rs1(1)
        rs1.Close
        Set rs1 = Nothing
    End If
    
    If  rs2.EOF And rs2.BOF Then
        rs2.Close
        Set rs2 = Nothing

	 	If Len(Request("txtSalesGroup")) And BlankchkFlg = False Then
		   Call DisplayMsgBox("970000", vbInformation, "�����׷�", "", I_MKSCRIPT)	'��: you must release this line if you change msg into code
	       BlankchkFlg = True	
%>
<Script Language=vbscript>
			parent.frm1.txtSalesGroup.focus
</Script>	       	
<%	       
		End If	
    Else    
		arrRsVal(2) = rs2(0)
		arrRsVal(3) = rs2(1)
        rs2.Close
        Set rs2 = Nothing
    End If
    
    If  rs3.EOF And rs3.BOF Then
        rs3.Close
        Set rs3 = Nothing

	 	If Len(Request("txtMovType")) And BlankchkFlg = False Then
		   Call DisplayMsgBox("970000", vbInformation, "��������", "", I_MKSCRIPT)	'��: you must release this line if you change msg into code		   
	       BlankchkFlg = True	
%>
<Script Language=vbscript>
			parent.frm1.txtMovType.focus
</Script>	       	
<%	       
		End If	
    Else    
		arrRsVal(4) = rs3(0)
		arrRsVal(5) = rs3(1)
        rs3.Close
        Set rs3 = Nothing
    End If

    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If    
        
    If BlankchkFlg = False Then	
		If  rs0.EOF And rs0.BOF Then
		    Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
		    rs0.Close
		    Set rs0 = Nothing
%>
<Script Language=vbscript>
			parent.frm1.txtBpCd.focus
</Script>	       	
<%		    
    
			' �� ��ġ�� �ִ� Response.End �� �����Ͽ��� ��. Client �ܿ��� Name�� ��� �ѷ��� �Ŀ� Response.End �� �����.
		Else    
		    Call  MakeSpreadSheetData()
		End If
	End If	
End Sub

'----------------------------------------------------------------------------------------------------------
' Set default value or preset value
'----------------------------------------------------------------------------------------------------------
Sub TrimData()
End Sub


%>
<Script Language=vbscript>
    With parent
        .frm1.txtBpNm.value				=  "<%=ConvSPChars(arrRsVal(1))%>" 	
  		.frm1.txtSalesGroupNm.value		=  "<%=ConvSPChars(arrRsVal(3))%>" 	
  		.frm1.txtMovTypeNm.value		=  "<%=ConvSPChars(arrRsVal(5))%>" 	
        
        If "<%=lgDataExist%>" = "Yes" Then
			'Set condition data to hidden area
			If "<%=lgPageNo%>" = "1" Then   ' "1" means that this query is first and next data exists
				.frm1.HBpCd.value		= "<%=ConvSPChars(Request("txtBpCd"))%>"
				.frm1.HSalesGroup.value	= "<%=ConvSPChars(Request("txtSalesGroup"))%>"
				.frm1.HMovType.value	= "<%=ConvSPChars(Request("txtMovType"))%>"
				.frm1.HRadio.value		= "<%=Request("txtRadio")%>"
				.frm1.HDnReqNo.value		= "<%=ConvSPChars(Request("txtDnReqNo"))%>"
				.frm1.HSCMDoNo.value		= "<%=ConvSPChars(Request("txtSCMDoNo"))%>"

				.frm1.HSONO.value		= "<%=ConvSPChars(Request("txtSONO"))%>"
				.frm1.HCCNo.value		= "<%=ConvSPChars(Request("txtCCNO"))%>"
				.frm1.HDlvyFrDt.value	= "<%=Request("txtDlvyFrDt")%>"
				.frm1.HDlvyToDt.value	= "<%=Request("txtDlvyToDt")%>"
				.frm1.HDNFrDt.value		= "<%=Request("txtDNFrDt")%>"		
				.frm1.HDNToDt.value		= "<%=Request("txtDNToDt")%>"				
			End If    
		
			'Show multi spreadsheet data from this line
			.ggoSpread.Source    = .frm1.vspdData 
			.ggoSpread.SSShowData "<%=lgstrData%>"                            '��: Display data 
			.lgPageNo			 = "<%=lgPageNo%>"							  '��: Next next data tag
'			.lgStrPrevKey		 = "<%=lgStrPrevKey%>"                       '��: set next data tag
  			.DbQueryOk
		End IF      
	End with
</Script>	
<%
    Response.End													'��: �����Ͻ� ���� ó���� ������ 
%>
