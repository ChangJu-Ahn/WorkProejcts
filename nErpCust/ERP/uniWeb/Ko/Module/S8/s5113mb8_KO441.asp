<%'======================================================================================================
'*  1. Module Name          : Sales
'*  2. Function Name        : B/L���� 
'*  3. Program ID           : S5113MB8
'*  4. Program Name         : B/L��Ȳ��ȸ 
'*  5. Program Desc         : B/L��Ȳ��ȸ 
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2002/08/02
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Hwangseongbae
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              :
'=======================================================================================================
%>
<!-- #Include file="../../inc/incSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDBAgent.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%
On Error Resume Next                                                                         
                                                                         
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0 , rs1, rs2, rs3			   '�� : DBAgent Parameter ���� 
Dim lgStrData                                                 '�� : Spread sheet�� ������ ����Ÿ�� ���� ���� 
Dim lgMaxCount                                                '�� : Spread sheet �� visible row �� 
Dim lgTailList                                                '�� : Orderby���� ���� field ����Ʈ 
Dim lgSelectList
Dim lgSelectListDT
Dim lgDataExist
Dim lgPageNo

Dim iFrPoint
iFrPoint=0

    Call HideStatusWnd 
    Call LoadBasisGlobalInf()
	Call LoadInfTB19029B("Q", "S", "NOCOOKIE", "QB")
	Call LoadBNumericFormatB("Q", "S", "NOCOOKIE", "QB")

	lgPageNo       = UNICInt(Trim(Request("lgPageNo")),0)   '��: "0"(First),"1"(Second),"2"(Third),"3"(...)
	lgMaxCount     = CInt(100)								'�� : �ѹ��� �����ü� �ִ� ����Ÿ �Ǽ� 
	lgSelectList   = Request("lgSelectList")                               '�� : select ����� 
	lgTailList     = Request("lgTailList")                                 '�� : Orderby value
	lgSelectListDT = Split(Request("lgSelectListDT"), gColSep)             '�� : �� �ʵ��� ����Ÿ Ÿ�� 
	lgDataExist      = "No"
	    
    Call FixUNISQLData()									 '�� : DB-Agent�� ���� parameter ����Ÿ set
    Call QueryData()										 '�� : DB-Agent�� ���� ADO query
'----------------------------------------------------------------------------------------------------------
Sub MakeSpreadSheetData()

    Dim iLoopCount                                                                     
    Dim iRowStr
    Dim ColCnt
    
    lgDataExist    = "Yes"
    lgstrData      = ""
  
    If CLng(lgPageNo) > 0 Then
       rs0.Move     = CLng(lgMaxCount) * CLng(lgPageNo)                  'lgMaxCount:Max Fetched Count at once , lgStrPrevKeyIndex : Previous PageNo
       iFrPoint		= CLng(lgMaxCount) * CLng(lgPageNo)
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
Sub SetConditionData()
End Sub
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()
	Dim iStrWhere1

    Redim UNISqlId(0)                                                     '��: SQL ID ������ ���� ����Ȯ�� 
'--------------- ������ coding part(�������,Start)----------------------------------------------------
    Redim UNIValue(0,2)

    UNISqlId(0) = "S5113MA801"
'--------------- ������ coding part(�������,End)------------------------------------------------------

    UNIValue(0,0) = Trim(lgSelectList)                                      '��: Select list
    
'--------------- ������ coding part(�������,Start)----------------------------------------------------
	iStrWhere1 = ""
	
    If Len(Trim(Request("txtApplicant"))) Then
		iStrWhere1 = iStrWhere1 & " AND BH.sold_to_party =  " & FilterVar(Request("txtApplicant"), "''", "S") & ""		'������ 
	End If

    If Len(Trim(Request("txtFromDt"))) Then
		iStrWhere1 = iStrWhere1 & " AND BH.bill_dt >=  " & FilterVar(UNIConvDate(Request("txtFromDt")), "''", "S") & ""						'������ 
	End If

    If Len(Trim(Request("txtToDt"))) Then
		iStrWhere1 = iStrWhere1 & " AND BH.bill_dt <= '" &UNIConvDate(Request("txtToDt")) & "'"							'������ 
	End If

    If Len(Trim(Request("txtForwarder"))) Then
		iStrWhere1 = iStrWhere1 & " AND BI.Forwarder =  " & FilterVar(Request("txtForwarder"), "''", "S") & ""				'����ä������ 
	End If

    If Len(Trim(Request("txtSalesGrp"))) Then
		iStrWhere1 = iStrWhere1 & " AND BH.sales_grp =  " & FilterVar(Request("txtSalesGrp"), "''", "S") & ""				'�����׷� 
	End If

    If Len(Trim(Request("txtPostFiFlag"))) Then
		iStrWhere1 = iStrWhere1 & " AND BH.post_flag =  " & FilterVar(Request("txtPostFiFlag"), "''", "S") & ""				'Ȯ������ 
	End If

	If Len(Request("gBizArea")) Then
		iStrWhere1 = iStrWhere1 & " AND BH.BIZ_AREA = " & FilterVar(Request("gBizArea"), "''", "S") & " "			
	End If

	If Len(Request("gSalesGrp")) Then
		iStrWhere1 = iStrWhere1 & " AND BH.SALES_GRP = " & FilterVar(Request("gSalesGrp"), "''", "S") & " "			
	End If

	If Len(Request("gSalesOrg")) Then
		iStrWhere1 = iStrWhere1 & " AND BH.SALES_ORG = " & FilterVar(Request("gSalesOrg"), "''", "S") & " "			
	End If
	
	UNIValue(0,1) = iStrWhere1
'--------------- ������ coding part(�������,End)------------------------------------------------------

    UNIValue(0,UBound(UNIValue,2)) = UCase(Trim(lgTailList))
    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '��: set ADO read mode
 
End Sub
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
		Call parent.SetFocusToDocument("M")		
		parent.frm1.txtFromDt.focus
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

			'Show multi spreadsheet data from this line
			Parent.ggoSpread.Source		= .vspdData 
			.vspdData.Redraw = False
			Parent.ggoSpread.SSShowDataByClip "<%=lgstrData%>","F"						'�� : Display data
			Call Parent.ReFormatSpreadCellByCellByCurrency(.vspdData,"<%=iFrPoint+1%>",.vspdData.MaxRows,Parent.GetKeyPos("A",5),Parent.GetKeyPos("A",6),"A","Q","X","X")

			parent.lgPageNo				=  "<%=lgPageNo%>"							'��: Next next data tag
			.txtHApplicant.value	= "<%=ConvSPChars(Request("txtApplicant"))%>"
			.txtHFromDt.value		= "<%=Request("txtFromDT")%>"
			.txtHToDt.value			= "<%=Request("txtToDT")%>"
			.txtHForwarder.value	= "<%=ConvSPChars(Request("txtForwarder"))%>"
			.txtHSalesGrp.value		= "<%=ConvSPChars(Request("txtSalesGrp"))%>"
			.txtHPostfiFlag.value	= "<%=Request("txtPostFiFlag")%>"
			parent.DbQueryOk
			.vspdData.Redraw = True        
		End If
	End with
</Script>	
