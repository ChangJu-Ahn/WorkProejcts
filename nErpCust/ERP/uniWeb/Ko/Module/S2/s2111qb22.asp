<%
'**********************************************************************************************
'*  1. Module Name          : Sales
'*  2. Function Name        : 
'*  3. Program ID           : S2111QB22
'*  4. Program Name         : ������ ǰ���ǸŰ�ȹ������ȸ 
'*  5. Program Desc         : ������ ǰ���ǸŰ�ȹ������ȸ 
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 
'*  8. Modified date(Last)  : 2001/12/19
'*  9. Modifier (First)     : Cho Song Hyun
'* 10. Modifier (Last)      : sonbumyeol
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              :
'**********************************************************************************************
%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<%                                                                         '�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 

On Error Resume Next

Call LoadBasisGlobalInf()

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
'--------------- ������ coding part(��������,End)----------------------------------------------------------
  
    Call HideStatusWnd 


    lgStrPrevKey   = Request("lgStrPrevKey_B")                               '�� : Next key flag
    lgMaxCount     = CInt(100)                           '�� : �ѹ��� �����ü� �ִ� ����Ÿ �Ǽ� 
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
    Redim UNISqlId(0)                                                     '��: SQL ID ������ ���� ����Ȯ�� 
    '--------------- ������ coding part(�������,Start)----------------------------------------------------

    Redim UNIValue(0,3)

    UNISqlId(0) = "S2111QA202"

    '--------------- ������ coding part(�������,End)------------------------------------------------------
    UNIValue(0,0) = lgSelectList                                          '��: Select list
    '--------------- ������ coding part(�������,Start)----------------------------------------------------

	strVal = " "

	Select Case CStr(Request("txtConPlanTypeCd"))
	Case CStr(1)					'�����ΰ�� 
		strVal = "SELECT SD.ITEM_CD ITEM_CD, SUM(SD.NET_AMT_LOC) NET_AMT_LOC,MONTH(SH.SO_DT) SO_MONTH"
		strVal = strVal + " FROM S_SO_DTL SD, S_SO_HDR SH"
		strVal = strVal + " WHERE SD.SO_NO = SH.SO_NO AND SH.CFM_FLAG = " & FilterVar("Y", "''", "S") & " "

		If CStr(Request("txtConDealTypeCd")) = CStr(1) Then
			strVal = strVal + " AND SH.EXPORT_FLAG = " & FilterVar("N", "''", "S") & " "
		ElseIf CStr(Request("txtConDealTypeCd")) = CStr(2) Then
			strVal = strVal + " AND SH.EXPORT_FLAG = " & FilterVar("Y", "''", "S") & " "
		Else
			strVal = strVal
		End If

		If Len(Request("txtConSalesOrg")) Then
			strVal = strVal + " AND SH.SALES_ORG = " & FilterVar(UCase(Request("txtConSalesOrg")), "''", "S") & " "
		Else
			strVal = strVal + ""
		End If
	
		If Len(Request("txtConSpYear")) Then
			strVal = strVal + " AND YEAR(SH.SO_DT) = " & FilterVar(Request("txtConSpYear"), "''", "S") & ""
		Else
			strVal = strVal + ""
		End If

		If Len(Request("txtItemCd")) Then
			strVal = strVal + " AND SD.ITEM_CD = " & FilterVar(UCase(Request("txtItemCd")), "''", "S") & " "
		Else
			strVal = strVal + ""
		End If

		strVal = strVal + " GROUP BY SD.ITEM_CD,MONTH(SH.SO_DT)"

	Case CStr(2)					'�����ΰ�� 
		strVal = "SELECT SD.ITEM_CD ITEM_CD, SUM(SD.BILL_AMT_LOC) NET_AMT_LOC,MONTH(SH.BILL_DT) SO_MONTH"
		strVal = strVal + " FROM S_BILL_DTL SD, S_BILL_HDR SH"
		strVal = strVal + " WHERE SD.BILL_NO = SH.BILL_NO"

		If CStr(Request("txtConDealTypeCd")) = CStr(1) Then
			strVal = strVal + " AND SH.BL_FLAG = " & FilterVar("N", "''", "S") & " "
		ElseIf CStr(Request("txtConDealTypeCd")) = CStr(2) Then
			strVal = strVal + " AND SH.BL_FLAG = " & FilterVar("Y", "''", "S") & " "
		Else
			strVal = strVal
		End If

		If Len(Request("txtConSalesOrg")) Then
			strVal = strVal + " AND SH.SALES_ORG = " & FilterVar(UCase(Request("txtConSalesOrg")), "''", "S") & " "
		Else
			strVal = strVal + ""
		End If
	
		If Len(Request("txtConSpYear")) Then
			strVal = strVal + " AND YEAR(SH.BILL_DT) = " & FilterVar(Request("txtConSpYear"), "''", "S") & ""
		Else
			strVal = strVal + ""
		End If

		If Len(Request("txtItemCd")) Then
			strVal = strVal + " AND SD.ITEM_CD = " & FilterVar(UCase(Request("txtItemCd")), "''", "S") & " "
		Else
			strVal = strVal + ""
		End If

		strVal = strVal + " GROUP BY SD.ITEM_CD,MONTH(SH.BILL_DT)"
	
	End Select	

    UNIValue(0,1) = strVal   


	strVal = " "

	strVal = " AND A.ORG_GRP_FLAG=" & FilterVar("O", "''", "S") & " "

	If Len(Request("txtConSalesOrg")) Then
		strVal = strVal + " AND A.SALES_ORG = " & FilterVar(UCase(Request("txtConSalesOrg")), "''", "S") & " "
	Else
		strVal = strVal + ""
	End If
	
	If Len(Request("txtConSpYear")) Then
		strVal = strVal + " AND A.SP_YEAR = " & FilterVar(Request("txtConSpYear"), "''", "S") & ""
	Else
		strVal = strVal + ""
	End If
	
	If Len(Request("txtConPlanTypeCd")) Then
		strVal = strVal + " AND A.PLAN_FLAG = " & FilterVar(UCase(Request("txtConPlanTypeCd")), "''", "S") & " "
	Else
		strVal = strVal + ""
	End If
	
	If Len(Request("txtConDealTypeCd")) Then
		strVal = strVal + " AND A.EXPORT_FLAG = " & FilterVar(UCase(Request("txtConDealTypeCd")), "''", "S") & " "
	Else
		strVal = strVal + ""
	End If

	If Len(Request("txtItemCd")) Then
		strVal = strVal + " AND A.ITEM_CD = " & FilterVar(UCase(Request("txtItemCd")), "''", "S") & " "
	Else
		strVal = strVal + ""
	End If

    UNIValue(0,2) = strVal   
    
    '--------------- ������ coding part(�������,End)------------------------------------------------------
'==    UNIValue(0,UBound(UNIValue,2)) = UCase(Trim(lgTailList))
    UNIValue(0,UBound(UNIValue,2)) = " " + UCase(Trim(lgTailList)) 
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
        Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
        rs0.Close
        Set rs0 = Nothing
        Response.End													'��: �����Ͻ� ���� ó���� ������ 
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
    With parent
         .ggoSpread.Source    = .frm1.vspdData2 
         .ggoSpread.SSShowDataByClip "<%=lgstrData%>"                            '��: Display data 
         .lgStrPrevKey_B      =  "<%=lgStrPrevKey%>"                       '��: set next data tag
         .DbQueryOk("B")
	End with
</Script>	

