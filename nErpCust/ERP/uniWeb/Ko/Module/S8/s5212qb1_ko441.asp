<%'======================================================================================================
'*  1. Module Name          : Sales
'*  2. Function Name        : 
'*  3. Program ID           : S5212QB1
'*  4. Program Name         : B/L����ȸ 
'*  5. Program Desc         : B/L����ȸ 
'*  6. Modified date(First) : 2000/11/01
'*  7. Modified date(Last)  : 2000/11/01
'*  8. Modifier (First)     : KimTaeHyun
'*  9. Modifier (Last)      : KimTaeHyun
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(��) means that "Do not change"
'=======================================================================================================
%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDBAgent.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%                                                                         '�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 

On Error Resume Next

Dim lgADF                                                                  '�� : ActiveX Data Factory ���� �������� 
Dim lgstrRetMsg                                                            '�� : Record Set Return Message �������� 
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0,rs1, rs2, rs3                         '�� : DBAgent Parameter ���� 
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

Dim iFrPoint
iFrPoint=0
'--------------- ������ coding part(��������,End)----------------------------------------------------------
 
    Call HideStatusWnd 
	Call LoadBasisGlobalInf()
	Call LoadInfTB19029B("Q", "S", "NOCOOKIE", "QB")
	Call LoadBNumericFormatB("Q", "S", "NOCOOKIE", "QB")

    lgStrPrevKey   = Request("lgStrPrevKey")                               '�� : Next key flag
    lgMaxCount     = CInt(100)                           '�� : �ѹ��� �����ü� �ִ� ����Ÿ �Ǽ� 
    lgSelectList   = Request("lgSelectList")                               '�� : select ����� 
    lgSelectListDT = Split(Request("lgSelectListDT"), gColSep)             '�� : �� �ʵ��� ����Ÿ Ÿ�� 
    lgTailList     = Request("lgTailList")                                 '�� : Orderby value

    Call TrimData()
    Call FixUNISQLData()
    Call QueryData()
'----------------------------------------------------------------------------------------------------------
Private Sub MakeSpreadSheetData()
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
        iFrPoint = iCnt  *  lgMaxCount
    Next

    iRCnt = -1
    
    Do while Not (rs0.EOF Or rs0.BOF)
        iRCnt =  iRCnt + 1
        iStr = ""
        
		For ColCnt = 0 To UBound(lgSelectListDT) - 1 
            iStr = iStr & Chr(11) & FormatRsString(lgSelectListDT(ColCnt),rs0(ColCnt))
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
Private Sub FixUNISQLData()

    Dim strVal
    Dim arrVal(4)
    Redim UNISqlId(5)                                                     '��: SQL ID ������ ���� ����Ȯ�� 
    '--------------- ������ coding part(�������,Start)----------------------------------------------------

    Redim UNIValue(5,2)

    UNISqlId(0) = "S5212QA101"
    UNISqlId(1) = "S0000QA002"
    UNISqlId(2) = "S0000QA005"
    UNISqlId(3) = "S0000QA001"
    '--------------- ������ coding part(�������,End)------------------------------------------------------
    UNIValue(0,0) = lgSelectList                                          '��: Select list
    '--------------- ������ coding part(�������,Start)----------------------------------------------------

	strVal = " "
	
	If Len(Request("txtconBp_cd")) Then
		strVal = "AND A.BILL_TO_PARTY = " & FilterVar(Request("txtconBp_cd"), "''", "S") & " "		
	Else
		strVal = ""
	End If
	arrVal(0) = FilterVar(Request("txtconBp_cd"),"","S")

	If Len(Request("txtSalesGroup")) Then
		strVal = strVal & " AND A.SALES_GRP = " & FilterVar(Request("txtSalesGroup"), "''", "S") & " "			
	End If	
	arrVal(1) = FilterVar( Request("txtSalesGroup"),"","S")	
    
 	If Len(Request("txtItem_cd")) Then
		strVal = strVal & " AND B.ITEM_CD = " & FilterVar(Request("txtItem_cd"), "''", "S") & " "				
	End If	
	arrVal(2) = FilterVar(Request("txtItem_cd"),"","S")
	
	If Len(Request("txtBLNo")) Then
		strVal = strVal & " AND A.BILL_NO = " & FilterVar(Request("txtBLNo"), "''", "S") & " "		
	End If		
    
    If Len(Request("txtBLFrDt")) Then
		strVal = strVal & " AND A.BILL_DT >= " & FilterVar(UNIConvDate(Trim(Request("txtBLFrDt"))), "''", "S") & ""		
	End If		
	
	If Len(Request("txtBLToDt")) Then
		strVal = strVal & " AND A.BILL_DT <= " & FilterVar(UNIConvDate(Trim(Request("txtBLToDt"))), "''", "S") & ""		
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

    '--------------- ������ coding part(�������,End)------------------------------------------------------
    UNIValue(0,UBound(UNIValue,2)) = UCase(Trim(lgTailList))
    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '��: set ADO read mode
 
End Sub
'----------------------------------------------------------------------------------------------------------
Private Sub QueryData()
    Dim iStr
	Dim FalsechkFlg 

    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3)
    
    If  rs2.EOF And rs2.BOF Then
        rs2.Close
        Set rs2 = Nothing
   			If Len(Trim(Request("txtSalesGroup"))) Then
			   Call DisplayMsgBox("970000", vbInformation, "�����׷�", "", I_MKSCRIPT)	'��: you must release this line if you change msg into code
				%>
				<Script Language=vbscript>
				    parent.frm1.txtSalesGroup.focus
				</Script>	
				<%
			   FalsechkFlg = True	
			End If	
    Else    
		arrRsVal(2) = rs2(0)
		arrRsVal(3) = rs2(1)
        rs2.Close
        Set rs2 = Nothing
    End If
   
    
    If  rs1.EOF And rs1.BOF Then
        rs1.Close
        Set rs1 = Nothing
		If FalsechkFlg = False Then
   			If Len(Trim(Request("txtconBp_cd"))) Then
				Call DisplayMsgBox("970000", vbInformation, "������", "", I_MKSCRIPT)	'��: you must release this line if you change msg into code
				%>
				<Script Language=vbscript>
				    parent.frm1.txtconBp_cd.focus
				</Script>	
				<%
				FalsechkFlg = True	
			End If	
		End if
   Else    
		arrRsVal(0) = rs1(0)
		arrRsVal(1) = rs1(1)
        rs1.Close
        Set rs1 = Nothing
    End If
    

    If  rs3.EOF And rs3.BOF Then
        rs3.Close
        Set rs3 = Nothing

		If FalsechkFlg = False Then
   			If Len(Trim(Request("txtItem_cd"))) Then
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
    Else    
        Call  MakeSpreadSheetData()
    End If
End Sub
'----------------------------------------------------------------------------------------------------------
Private Sub TrimData()
End Sub

%>

<Script Language=vbscript>
    With parent
        .ggoSpread.Source    = .frm1.vspdData 
        .frm1.vspdData.Redraw = False
		.ggoSpread.SSShowDataByClip "<%=lgstrData%>","F"          '�� : Display data
		Call .ReFormatSpreadCellByCellByCurrency(.frm1.vspdData,"<%=iFrPoint+1%>",.frm1.vspddata.maxrows,.GetKeyPos("A",3),.GetKeyPos("A",2),"C","Q","X","X")
		Call .ReFormatSpreadCellByCellByCurrency(.frm1.vspdData,"<%=iFrPoint+1%>",.frm1.vspddata.maxrows,.GetKeyPos("A",3),.GetKeyPos("A",4),"A","Q","X","X")

        .frm1.txtconBp_Nm.value		= "<%=ConvSPChars(arrRsVal(1))%>" 
        .frm1.txtSalesGroupNm.value	= "<%=ConvSPChars(arrRsVal(3))%>" 
        .frm1.txtItem_Nm.value		= "<%=ConvSPChars(arrRsVal(5))%>" 
        .lgStrPrevKey        =  "<%=ConvSPChars(lgStrPrevKey)%>"                    '��: set next data tag
        .DbQueryOk
        .frm1.vspdData.Redraw = True                
	End with
</Script>	
<%
    Response.End													'��: �����Ͻ� ���� ó���� ������ 
%>
