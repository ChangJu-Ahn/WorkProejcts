<%'======================================================
'*  1. Module Name          : ���� 
'*  2. Function Name        : ����ä�ǰ��� 
'*  3. Program ID           : S5111GA1
'*  4. Program Name         : ����ä������ 
'*  5. Program Desc         : ADO Query
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2000/12/09
'*  8. Modified date(Last)  : 2001/12/18
'*  9. Modifier (First)     : Byun Jee Hyun
'* 10. Modifier (Last)      : Kim Hyungsuk
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              :
'*                            2000/12/09
'*                            2001/12/18	Dateǥ������ 
'=======================================================
%>
<!-- #Include file="../../inc/incSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDBAgent.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
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
'--------------- ������ coding part(��������,Start)--------------------------------------------------------
Dim strPoType	                                                           '�� : �������� 
Dim strPoFrDt	                                                           '�� : ������ 
Dim strPoToDt	                                                           '�� :
Dim strSpplCd	                                                           '�� : ����ó 
Dim strPurGrpCd	                                                           '�� : ���ű׷� 
Dim strItemCd	                                                           '�� : ǰ�� 
Dim strTrackNo	                                                           '�� : Tracking No
Dim arrRsVal(7)
Dim lgDataExist
Dim FalsechkFlg
'--------------- ������ coding part(��������,End)----------------------------------------------------------
  
    Call HideStatusWnd 
    Call LoadBasisGlobalInf()
	Call LoadInfTB19029B("Q", "S", "NOCOOKIE", "QB") 


    lgStrPrevKey   = Request("lgStrPrevKey")                               '�� : Next key flag
    lgMaxCount     = 100							                           '�� : �ѹ��� �����ü� �ִ� ����Ÿ �Ǽ� 
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

    iCnt = 0
    lgstrData = ""
    
    If Len(Trim(lgStrPrevKey)) Then                                        '�� : Chnage Nextkey str into int value
       If Isnumeric(lgStrPrevKey) Then
          iCnt = CInt(lgStrPrevKey)
       End If   
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
            lgStrPrevKey = CStr(iCnt)
            Exit Do
        End If
        rs0.MoveNext
	Loop

    If  iRCnt < lgMaxCount Then                                            '��: Check if next data exists
        lgStrPrevKey = ""                                                  '��: ���� ����Ÿ ����.
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
    Dim arrVal(3)
    Redim UNISqlId(4)                                                     '��: SQL ID ������ ���� ����Ȯ�� 
    '--------------- ������ coding part(�������,Start)----------------------------------------------------

    Redim UNIValue(4,2)
    lgDataExist = "Yes" 
    UNISqlId(0) = "S5111GA101"
	UNISqlId(1) = "S0000qA002"
	UNISqlId(2) = "S0000qA011"
	UNISqlId(3) = "S0000qA005"
	UNISqlId(4) = "S0000qA001"
    '--------------- ������ coding part(�������,End)------------------------------------------------------
    UNIValue(0,0) = lgSelectList                                          '��: Select list
    '--------------- ������ coding part(�������,Start)----------------------------------------------------

	strVal = " "
	
	If Len(Request("txtconBp_cd")) Then
		strVal = "AND A.SOLD_TO_PARTY = " & FilterVar(Request("txtconBp_cd"), "''", "S") & " "		
	Else
		strVal = ""
	End If
	arrVal(0) = FilterVar(Trim(Request("txtconBp_cd")),"","S")

	If Len(Request("txtBillType")) Then
		strVal = strVal & " AND A.BILL_TYPE = " & FilterVar(Request("txtBillType"), "''", "S") & " "			
	End If		
	arrVal(1) = FilterVar(Trim(Request("txtBillType")),"","S")
		   
 	If Len(Request("txtSalesGroup")) Then
		strVal = strVal & " AND A.SALES_GRP = " & FilterVar(Request("txtSalesGroup"), "''", "S") & " "				
	End If		
	arrVal(2) = FilterVar(Trim(Request("txtSalesGroup")),"","S")
    
	If Len(Request("txtItem_cd")) Then
		strVal = strVal & " AND B.ITEM_CD = " & FilterVar(Request("txtItem_cd"), "''", "S") & " "				
	End If		
	arrVal(3) = FilterVar(Trim(Request("txtItem_cd")),"","S")
    
    If Len(Request("txtBillFrDt")) Then
		strVal = strVal & " AND A.BILL_DT >= " & FilterVar(UNIConvDate(Request("txtBillFrDt")), "''", "S") & ""		
	End If		
	
	If Len(Request("txtBillToDt")) Then
		strVal = strVal & " AND A.BILL_DT <= " & FilterVar(UNIConvDate(Request("txtBillToDt")), "''", "S") & ""		
	End If

    If Request("txtRadio") = "Y" Then
		strVal = strVal & "AND A.POST_FLAG = " & FilterVar("Y", "''", "S") & " "
	ElseIf Request("txtRadio") = "N" Then
		strVal = strVal & "AND A.POST_FLAG = " & FilterVar("N", "''", "S") & " "			
	End If
	If Len(Request("gBizArea")) Then
		strVal = strVal & " AND A.BIZ_AREA = " & FilterVar(Request("gBizArea"), "''", "S") & " "			
	End If

	If Len(Request("gPlant")) Then
		strVal = strVal & " AND B.PLANT_CD = " & FilterVar(Request("gPlant"), "''", "S") & " "			
	End If

	If Len(Request("gSalesGrp")) Then
		strVal = strVal & " AND A.SALES_GRP = " & FilterVar(Request("gSalesGrp"), "''", "S") & " "			
	End If

	If Len(Request("gSalesOrg")) Then
		strVal = strVal & " AND A.SALES_ORG = " & FilterVar(Request("gSalesOrg"), "''", "S") & " "			
	End If

    UNIValue(0,1) = strVal   '---������ 
    UNIValue(1,0) = arrVal(0)
    UNIValue(2,0) = arrVal(1)
    UNIValue(3,0) = arrVal(2)
    UNIValue(4,0) = arrVal(3)
    
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
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3, rs4)

    FalsechkFlg = False
    If  rs1.EOF And rs1.BOF Then
        rs1.Close
        Set rs1 = Nothing

		If Len(Request("txtconBp_cd")) And FalsechkFlg = False Then
		   Call DisplayMsgBox("970000", vbInformation, "�ֹ�ó", "", I_MKSCRIPT)	'��: you must release this line if you change msg into code
			%>
			<Script Language=vbscript>
			    parent.frm1.txtconBp_cd.focus
			</Script>	
			<%
			FalsechkFlg = True
	       
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
		If  FalsechkFlg = False Then
			If Len(Request("txtBillType")) And FalsechkFlg = False Then
				Call DisplayMsgBox("970000", vbInformation, "����ä������", "", I_MKSCRIPT)	'��: you must release this line if you change msg into code
				%>
				<Script Language=vbscript>
				    parent.frm1.txtBillType.focus
				</Script>	
				<%
			   FalsechkFlg = True
			End If	
		End If
    Else    
		arrRsVal(2) = rs2(0)
		arrRsVal(3) = rs2(1)
        rs2.Close
        Set rs2 = Nothing
    End If

    If  rs4.EOF And rs4.BOF Then
        rs4.Close
        Set rs4 = Nothing
		If  FalsechkFlg = False Then
			If Len(Request("txtItem_cd")) And FalsechkFlg = False Then
				Call DisplayMsgBox("970000", vbInformation, "ǰ��", "", I_MKSCRIPT)	'��: you must release this line if you change msg into code
				%>
				<Script Language=vbscript>
				    parent.frm1.txtItem_cd.focus
				</Script>	
				<%
			   FalsechkFlg = True
	       		
			End If	
		End If
    Else    
		arrRsVal(6) = rs4(0)
		arrRsVal(7) = rs4(1)
        rs4.Close
        Set rs4 = Nothing
    End If

    If  rs3.EOF And rs3.BOF Then
        rs3.Close
        Set rs3 = Nothing
		If  FalsechkFlg = False Then
			If Len(Request("txtSalesGroup")) And FalsechkFlg = False Then
				Call DisplayMsgBox("970000", vbInformation, "�����׷�", "", I_MKSCRIPT)	'��: you must release this line if you change msg into code
				%>
				<Script Language=vbscript>
				    parent.frm1.txtSalesGroup.focus
				</Script>	
				<%
			   FalsechkFlg = True
	       End If	
		End If	

    Else    
		arrRsVal(4) = rs3(0)
		arrRsVal(5) = rs3(1)
        rs3.Close
        Set rs3 = Nothing
    End If
  

    iStr = Split(lgstrRetMsg,gColSep)
    
    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If    
        
    If  rs0.EOF And rs0.BOF And FalsechkFlg = False Then
        Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
        rs0.Close
        Set rs0 = Nothing
        %>
			<Script Language=vbscript>
			    parent.frm1.txtconBp_cd.focus
			</Script>	
		<%
    Else    
        Call  MakeSpreadSheetData()
    End If

End Sub

'----------------------------------------------------------------------------------------------------------
' Set default value or preset value
'----------------------------------------------------------------------------------------------------------
Sub TrimData()
End Sub


%>

<Script Language=vbscript>
If "<%=lgDataExist%>" = "Yes" Then
    With parent
        .ggoSpread.Source    = .frm1.vspdData 
        .ggoSpread.SSShowDataByClip "<%=lgstrData%>"                            '��: Display data 
        .lgStrPrevKey        =  "<%=lgStrPrevKey%>"                       '��: set next data tag
        .frm1.txtconBp_Nm.value		= "<%=ConvSPChars(arrRsVal(1))%>" 
        .frm1.txtBillTypeNm.value	= "<%=ConvSPChars(arrRsVal(3))%>"
        .frm1.txtSalesGroupNm.value	= "<%=ConvSPChars(arrRsVal(5))%>" 
        .frm1.txtItem_Nm.value		= "<%=ConvSPChars(arrRsVal(7))%>" 
        .frm1.txtHconBp_cd.value	= "<%=ConvSPChars(Request("txtconBp_cd"))%>"
		.frm1.txtHBillType.value	= "<%=ConvSPChars(Request("txtBillType"))%>"
		.frm1.txtHSalesGroup.value	= "<%=ConvSPChars(Request("txtSalesGroup"))%>"
		.frm1.txtHItem_cd.value		= "<%=ConvSPChars(Request("txtItem_cd"))%>"
		.frm1.txtHBillFrDt.value	= "<%=Request("txtBillFrDt")%>"
		.frm1.txtHBillToDt.value	= "<%=Request("txtBillToDt")%>"
        .frm1.txtHRadio.value	= "<%=Request("txtRadio")%>"
        If "<%=FalsechkFlg%>" = False Then
        .DbQueryOk
        End If
	End with
End if
</Script>	

<%
    Response.End													'��: �����Ͻ� ���� ó���� ������ 
%>
