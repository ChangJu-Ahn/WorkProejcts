<%'======================================================
'*  1. Module Name          : ���� 
'*  2. Function Name        : ������� 
'*  3. Program ID           : s5311pb1
'*  4. Program Name         : ���ݰ�꼭������ȣ Popup
'*  5. Program Desc         : 
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2002/05/10
'*  8. Modified date(Last)  : 2002/05/10
'*  9. Modifier (First)     : Hwangseongbae
'* 10. Modifier (Last)      : Hwangseongbae
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              :
'=======================================================
%>
<!-- #Include file="../../inc/incSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDBAgent.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%
On Error Resume Next														'��: 
                                                                         
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0 , rs1, rs2, rs3			   '�� : DBAgent Parameter ���� 
Dim lgStrData                                                 '�� : Spread sheet�� ������ ����Ÿ�� ���� ���� 
Dim lgMaxCount                                                '�� : Spread sheet �� visible row �� 
Dim lgTailList                                                '�� : Orderby���� ���� field ����Ʈ 
Dim lgSelectList
Dim lgSelectListDT
Dim lgDataExist
Dim lgPageNo

	Call LoadBasisGlobalInf()
	Call LoadInfTB19029B("Q", "S", "NOCOOKIE", "PB")
	Call LoadBNumericFormatB("Q", "S", "NOCOOKIE", "PB")
'	Call LoadBNumericFormatB("Q", "S", "NOCOOKIE", "PB")
	 Call HideStatusWnd 

	lgPageNo       = UNICInt(Trim(Request("lgPageNo")),0)    '��: "0"(First),"1"(Second),"2"(Third),"3"(...)
	lgMaxCount     = 30							             '�� : �ѹ��� �����ü� �ִ� ����Ÿ �Ǽ� 
	lgSelectList   = Request("lgSelectList")                               '�� : select ����� 
	lgTailList     = Request("lgTailList")                                 '�� : Orderby value
	lgSelectListDT = Split(Request("lgSelectListDT"), gColSep)             '�� : �� �ʵ��� ����Ÿ Ÿ�� 
	lgDataExist      = "No"
	    
    Call FixUNISQLData()									 '�� : DB-Agent�� ���� parameter ����Ÿ set

    Call QueryData()										 '�� : DB-Agent�� ���� ADO query

'----------------------------------------------------------------------------------------------------------
' Make srpread sheet data
'----------------------------------------------------------------------------------------------------------
Sub MakeSpreadSheetData()

    Dim iLoopCount                                                                     
    Dim iRowStr
    Dim ColCnt
    
    lgDataExist    = "Yes"
    lgstrData      = ""
  
    If CLng(lgPageNo) > 0 Then
       rs0.Move     = CLng(lgMaxCount) * CLng(lgPageNo)                  'lgMaxCount:Max Fetched Count at once , lgStrPrevKeyIndex : Previous PageNo
    End If
    
    iLoopCount = -1
    
   Do while Not (rs0.EOF Or rs0.BOF)
   
        iLoopCount =  iLoopCount + 1
        iRowStr = ""
        
		For ColCnt = 0 To UBound(lgSelectListDT) - 1 
            iRowStr = iRowStr & Chr(11) & FormatRsString(lgSelectListDT(ColCnt),rs0(ColCnt))
		Next
 
        If iLoopCount < lgMaxCount Then
           lgstrData = lgstrData & iRowStr & Chr(11) & Chr(12)
        Else
           lgPageNo = lgPageNo + 1
           Exit Do
        End If
        
        rs0.MoveNext
	Loop

    If iLoopCount < lgMaxCount Then                                 '��: Check if next data exists
       lgPageNo = ""
    End If
    rs0.Close                                                       '��: Close recordset object
    Set rs0 = Nothing	                                            '��: Release ADF

End Sub

'----------------------------------------------------------------------------------------------------------
' Name : SetConditionData
' Desc : set value in condition area
'----------------------------------------------------------------------------------------------------------
Sub SetConditionData()
End Sub

'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()

    Redim UNISqlId(0)                                                     '��: SQL ID ������ ���� ����Ȯ�� 
'--------------- ������ coding part(�������,Start)----------------------------------------------------
    Redim UNIValue(0,9)

    UNISqlId(0) = "S5311PA101KO441"
'--------------- ������ coding part(�������,End)------------------------------------------------------

    UNIValue(0,0) = Trim(lgSelectList)                                      '��: Select list
    
'--------------- ������ coding part(�������,Start)----------------------------------------------------
    If Len(Trim(Request("txtBillToParty"))) Then
		UNIValue(0,1) = " " & FilterVar(Request("txtBillToParty"), "''", "S") & ""                      '�ֹ�ó 
	Else
		UNIValue(0,1) = "Null"
	End If
    If Len(Trim(Request("txtSalesGrp"))) Then
		UNIValue(0,2) = " " & FilterVar(Request("txtSalesGrp"), "''", "S") & ""                      '�ֹ�ó 
	Else
		UNIValue(0,2) = "Null"
	End If

    If Len(Trim(Request("txtVatType"))) Then
		UNIValue(0,3) = " " & FilterVar(Request("txtVatType"), "''", "S") & ""                      '�ֹ�ó 
	Else
		UNIValue(0,3) = "Null"
	End If
    If Len(Trim(Request("txtTaxBizArea"))) Then
		UNIValue(0,4) = " " & FilterVar(Request("txtTaxBizArea"), "''", "S") & ""                      '�ֹ�ó 
	Else
		UNIValue(0,4) = "Null"
	End If
    If Len(Trim(Request("txtFromDt"))) Then
		UNIValue(0,5) = " " & FilterVar(UNIConvDate(Request("txtFromDt")), "''", "S") & ""                             '������ 
	Else
		UNIValue(0,5) = "Null"
	End If

    If Len(Trim(Request("txtToDt"))) Then
		UNIValue(0,6) = " " & FilterVar(UNIConvDate(Request("txtToDt")), "''", "S") & ""									'������ 
	Else
		UNIValue(0,6) = "Null"
	End If
	
    If Len(Trim(Request("txtPostFlag"))) Then
		UNIValue(0,7) = " " & FilterVar(Request("txtPostFlag"), "''", "S") & ""                      '�ֹ�ó 
	Else
		UNIValue(0,7) = "Null"
	End If

	If Len(Request("gBizArea")) Then
		UNIValue(0,8) = UNIValue(0,8) & " AND A.BIZ_AREA_CD = " & FilterVar(Request("gBizArea"), "''", "S") & " "			
	End If

	If Len(Request("gSalesGrp")) Then
		UNIValue(0,8) = UNIValue(0,8) & " AND A.SALES_GRP = " & FilterVar(Request("gSalesGrp"), "''", "S") & " "			
	End If

	If Len(Request("gSalesOrg")) Then
		UNIValue(0,8) = UNIValue(0,8) & " AND A.SALES_ORG = " & FilterVar(Request("gSalesOrg"), "''", "S") & " "			
	End If

'--------------- ������ coding part(�������,End)------------------------------------------------------

    UNIValue(0,UBound(UNIValue,2)) = UCase(Trim(lgTailList))
    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '��: set ADO read mode
 
End Sub

'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
Sub QueryData()

    Dim lgstrRetMsg                                             '�� : Record Set Return Message �������� 
    Dim lgADF                                                   '�� : ActiveX Data Factory ���� �������� 
    Dim iStr

    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")

    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3)

	Set lgADF   = Nothing
	
    iStr = Split(lgstrRetMsg,gColSep)
	If iStr(0) <> "0" Then

        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If 

    If  rs0.EOF And rs0.BOF And FalsechkFlg =  False Then
        rs0.Close
        Set rs0 = Nothing
        Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
        %>
		<Script Language=vbscript>
		parent.frm1.txtBillToParty.focus
		</Script>	
        <%
    Else    
        Call  MakeSpreadSheetData()
        Call  SetConditionData()
    End If  
End Sub

%>

<Script Language=vbscript>
    With parent.frm1
		
		If "<%=lgDataExist%>" = "Yes" Then
			'Set condition data to hidden area
			.txtHBillToParty.value	= "<%=ConvSPChars(Request("txtBillToParty"))%>"
			.txtHFromDt.value		= "<%=Request("txtFromDT")%>"
			.txtHToDt.value			= "<%=Request("txtToDT")%>"
			.txtHSalesGrp.value		= "<%=ConvSPChars(Request("txtSalesGrp"))%>"
			.txtHVatType.value		= "<%=ConvSPChars(Request("txtVatType"))%>"
			.txtHTaxBizArea.value	= "<%=ConvSPChars(Request("txtTaxBizArea"))%>"
			.txtHPostFlag.value		= "<%=Request("txtHPostFlag")%>"

			'Show multi spreadsheet data from this line
			parent.ggoSpread.Source		= .vspdData
			.vspdData.Redraw = False 
			parent.ggoSpread.SSShowDataByClip "<%=lgstrData%>","F"                      '��: Display data 
			Call parent.ReFormatSpreadCellByCellByCurrency(.vspdData,-1,-1,parent.GetKeyPos("A",2),parent.GetKeyPos("A",3),"A","Q","X","X") 
			Call parent.ReFormatSpreadCellByCellByCurrency(.vspdData,-1,-1,parent.GetKeyPos("A",2),parent.GetKeyPos("A",4),"A","Q","X","X") 
			parent.lgPageNo				=  "<%=lgPageNo%>"						  '��: Next next data tag
			parent.DbQueryOk
			.vspdData.Redraw = True
		End If
	End with
</Script>	
