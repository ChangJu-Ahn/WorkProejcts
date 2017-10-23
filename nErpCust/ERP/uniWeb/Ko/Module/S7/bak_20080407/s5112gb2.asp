<%'======================================================
'*  1. Module Name          : ���� 
'*  2. Function Name        : ����ä�ǰ��� 
'*  3. Program ID           : S5112GA2
'*  4. Program Name         : ����ä�Ǽ�����ȸ 
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
Dim FalsechkFlg
Dim arrRsVal(7)
'--------------- ������ coding part(��������,End)----------------------------------------------------------
  
    Call HideStatusWnd 
    Call LoadBasisGlobalInf()
	Call LoadInfTB19029B("Q", "S", "NOCOOKIE", "QB") 


    lgStrPrevKey   = Request("lgStrPrevKey")                               '�� : Next key flag
    lgMaxCount     = 100							                       '�� : �ѹ��� �����ü� �ִ� ����Ÿ �Ǽ� 
    lgSelectList   = Request("lgSelectList")                               '�� : select ����� 
    lgSelectListDT = Split(Request("lgSelectListDT"), gColSep)             '�� : �� �ʵ��� ����Ÿ Ÿ�� 
    lgTailList     = Request("lgTailList")                                 '�� : Orderby value

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
    Redim UNISqlId(3)                                                     '��: SQL ID ������ ���� ����Ȯ�� 
    '--------------- ������ coding part(�������,Start)----------------------------------------------------

    Redim UNIValue(4,2)

    UNISqlId(0) = "S5112GA201"
    UNISqlId(1) = "s0000qa002"
    UNISqlId(2) = "s0000qa011"
    UNISqlId(3) = "S0000QA005"

    '--------------- ������ coding part(�������,End)------------------------------------------------------
    UNIValue(0,0) = lgSelectList                                          '��: Select list
    '--------------- ������ coding part(�������,Start)----------------------------------------------------

	strVal = " "

	If Len(Request("txtSoldToParty")) Then
		strVal = "AND A.SOLD_TO_PARTY = " & FilterVar(Request("txtSoldToParty"), "''", "S") & " "		
	Else
		strVal = ""
	End If
	arrVal(0) = FilterVar(Request("txtSoldToParty"),"","S")

	If Len(Request("txtBillType")) Then
		strVal = strVal & " AND A.BILL_TYPE = " & FilterVar(Request("txtBillType"), "''", "S") & " "			
	End If		
	arrVal(1) = FilterVar(Request("txtBillType"),"","S")
		   
 	If Len(Request("txtSalesGrp")) Then
		strVal = strVal & " AND A.SALES_GRP = " & FilterVar(Request("txtSalesGrp"), "''", "S") & " "			
	End If		
	arrVal(2) = FilterVar(Request("txtSalesGrp"),"","S")
    
    If Len(Request("txtBillFrDt")) Then
		strVal = strVal & " AND A.BILL_DT >= " & FilterVar(UNIConvDate(Request("txtBillFrDt")), "''", "S") & ""		
	End If		
	
	If Len(Request("txtBillToDt")) Then
		strVal = strVal & " AND A.BILL_DT <= " & FilterVar(UNIConvDate(Request("txtBillToDt")), "''", "S") & ""		
	End If

    If Request("txtRadio1") = "Y" Then
		strVal = strVal & "AND A.POST_FLAG = " & FilterVar("Y", "''", "S") & " "
	ElseIf Request("txtRadio1") = "N" Then
		strVal = strVal & "AND A.POST_FLAG = " & FilterVar("N", "''", "S") & " "			
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
' Query Data
'----------------------------------------------------------------------------------------------------------
Sub QueryData()
    Dim iStr

	FalsechkFlg = False 
    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3)

    FalsechkFlg = False
    If  rs1.EOF And rs1.BOF Then
        rs1.Close
        Set rs1 = Nothing

		If Len(Request("txtSoldToParty")) And FalsechkFlg = False Then
		   Call DisplayMsgBox("970000", vbInformation, "�ֹ�ó", "", I_MKSCRIPT)	'��: you must release this line if you change msg into code
		   Response.Write "<Script language=vbs>  " & vbCr
           Response.Write "   Parent.frm1.txtSoldToParty.focus " & vbCr    
           Response.Write "</Script>      " & vbCr    
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

		If Len(Request("txtBillType")) And FalsechkFlg = False Then
		   Call DisplayMsgBox("970000", vbInformation, "����ä������", "", I_MKSCRIPT)	'��: you must release this line if you change msg into code
		   Response.Write "<Script language=vbs>  " & vbCr
           Response.Write "   Parent.frm1.txtBillType.focus " & vbCr    
           Response.Write "</Script>      " & vbCr    
	       FalsechkFlg = True
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

		If Len(Request("txtSalesGrp")) And FalsechkFlg = False Then
		   Call DisplayMsgBox("970000", vbInformation, "�����׷�", "", I_MKSCRIPT)	'��: you must release this line if you change msg into code
		   Response.Write "<Script language=vbs>  " & vbCr
           Response.Write "   Parent.frm1.txtSalesGrp.focus " & vbCr    
           Response.Write "</Script>      " & vbCr    
	       FalsechkFlg = True
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
		Response.Write "<Script language=vbs>  " & vbCr
        Response.Write "parent.DbQueryOk " & vbCr    
        Response.Write "</Script>      " & vbCr    
    Else    
        Call  MakeSpreadSheetData()
    End If
   
 End Sub

%>
<Script Language=vbscript>
    With parent
        .ggoSpread.Source    = .frm1.vspdData 
        .ggoSpread.SSShowDataByClip "<%=lgstrData%>"                            '��: Display data 
        .frm1.txtSoldToPartyNm.value		= "<%=ConvSPChars(arrRsVal(1))%>" 
        .frm1.txtBillTypeNm.value	= "<%=ConvSPChars(arrRsVal(3))%>" 
        .frm1.txtSalesGrpNm.value	= "<%=ConvSPChars(arrRsVal(5))%>" 
        .frm1.txtHSoldToParty.value	= "<%=ConvSPChars(Request("txtSoldToParty"))%>"
		.frm1.txtHBillType.value	= "<%=ConvSPChars(Request("txtBillType"))%>"
		.frm1.txtHSalesGrp.value	= "<%=ConvSPChars(Request("txtSalesGrp"))%>"
		.frm1.txtHBillFrDt.value	= "<%=Request("txtBillFrDt")%>"
		.frm1.txtHRadio1.value		= "<%=Request("txtRadio1")%>"
		.frm1.txtHBillToDt.value	= "<%=Request("txtBillToDt")%>"                
        .lgStrPrevKey        =  "<%=lgStrPrevKey%>"                       '��: set next data tag
        .DbQueryOk
	End with
</Script>	
<%
    Response.End													'��: �����Ͻ� ���� ó���� ������ 
%>
