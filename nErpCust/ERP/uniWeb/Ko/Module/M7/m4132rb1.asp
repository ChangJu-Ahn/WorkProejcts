<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% Option Explicit%>
<% session.CodePage=949 %>

<!--
'**********************************************************************************************
'*  1. Module Name          : Prucurement
'*  2. Function Name        : 
'*  3. Program ID           : M3113RB1
'*  4. Program Name         : ��ǰ������� 
'*  5. Program Desc         : ��ǰ������� 
'*  6. Component List       : 
'*  7. Modified date(First) : 2002/03/21	
'*  8. Modified date(Last)  : 2003/05/22
'*  9. Modifier (First)     : Oh chang won
'* 10. Modifier (Last)      : Kang Su Hwan
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(��) means that "Do not change" 
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              :
'**********************************************************************************************
-->
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
Dim BlankchkFlg
'----------------------- �߰��� �κ� ----------------------------------------------------------------------
Dim arrRsVal(5)								'* : ȭ�鿡 ��ȸ�ؿ� Name�� ��Ƴ��� ���� ���� Array
'----------------------------------------------------------------------------------------------------------
'--------------- ������ coding part(��������,End)----------------------------------------------------------
  
    Call HideStatusWnd 
	Call LoadBasisGlobalInf()
	Call LoadInfTB19029B("I", "*", "NOCOOKIE", "PB")
	Call LoadBNumericFormatB("I", "*", "NOCOOKIE", "PB")

    lgStrPrevKey   = Request("lgStrPrevKey")                               '�� : Next key flag
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
	Const C_SHEETMAXROWS_D  = 100            
    Dim  RecordCnt
    Dim  ColCnt
    Dim  iCnt
    Dim  iRCnt
    Dim  iStr
	Dim PvArr

    iCnt = 0
    lgstrData = ""

    If Len(Trim(lgStrPrevKey)) Then                                        '�� : Chnage Nextkey str into int value
       If Isnumeric(lgStrPrevKey) Then
          iCnt = CInt(lgStrPrevKey)
       End If   
    End If   

    For iRCnt = 1 to iCnt  *  C_SHEETMAXROWS_D                                   '�� : Discard previous data
        rs0.MoveNext
    Next

    iRCnt = -1
    
   ReDim PvArr(C_SHEETMAXROWS_D - 1)
    Do while Not (rs0.EOF Or rs0.BOF)
        iRCnt =  iRCnt + 1
        iStr = ""
		For ColCnt = 0 To UBound(lgSelectListDT) - 1 
            Select Case  lgSelectListDT(ColCnt)
               Case "DD"  '��¥ 
					iStr = iStr & Chr(11) & UNIDateClientFormat(rs0(ColCnt))
               Case "F2"  '�ݾ� 
                    iStr = iStr & Chr(11) & UniConvNumberDBToCompany(rs0(ColCnt), ggAmtOfMoney.DecPoint, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit, 0)
               Case "F3"  '���� 
                    iStr = iStr & Chr(11) & UNINumClientFormat(rs0(ColCnt), ggQty.DecPoint       , 0)
               Case "F4"  '�ܰ� 
                    iStr = iStr & Chr(11) & UniConvNumberDBToCompany(rs0(ColCnt), ggUnitCost.DecPoint, ggUnitCost.RndPolicy, ggUnitCost.RndUnit, 0)
               Case "F5"  'ȯ�� 
                    iStr = iStr & Chr(11) & UNINumClientFormat(rs0(ColCnt), ggExchRate.DecPoint  , 0)
               Case "F2D" '�ŷ��ݾ� 
                    iStr = iStr & Chr(11) & UNIConvNumDBToCompanyByCurrency(rs0(ColCnt), lgCurrency, ggAmtOfMoneyNo,"X","X")
               Case "F4D" '�ŷ��ܰ� 
                    iStr = iStr & Chr(11) & UNIConvNumDBToCompanyByCurrency(rs0(ColCnt), lgCurrency, ggUnitCostNo,"X","X")
               Case Else
                    iStr = iStr & Chr(11) & ConvSPChars(rs0(ColCnt)) 
            End Select
		Next
 
        If  iRCnt < C_SHEETMAXROWS_D Then
            lgstrData      = lgstrData      & iStr & Chr(11) & Chr(12)
        Else
            iCnt = iCnt + 1
            lgStrPrevKey = CStr(iCnt)
            Exit Do
        End If

        PvArr(iRCnt) = lgstrData
        lgstrData=""
        rs0.MoveNext
	Loop
    lgstrData = Join(PvArr,"")

    If  iRCnt < C_SHEETMAXROWS_D Then                                            '��: Check if next data exists
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
    Redim UNISqlId(2)                                                     '��: SQL ID ������ ���� ����Ȯ�� 
    '--------------- ������ coding part(�������,Start)----------------------------------------------------

    Redim UNIValue(1,2)

    UNISqlId(0) = "M4132RA101"									'* : ������ ��ȸ�� ���� SQL�� ���� 
 
    '--------------- ������ coding part(�������,End)------------------------------------------------------
    UNIValue(0,0) = lgSelectList                                          '��: Select list
    '--------------- ������ coding part(�������,Start)----------------------------------------------------


	If Trim(Request("txtMvmtNo")) <> "" Then
		strVal = " AND A.MVMT_RCPT_NO >= " & FilterVar(Trim(UCase(Request("txtMvmtNo"))), " " , "S") & "  AND A.MVMT_RCPT_NO <=  " & FilterVar(Trim(UCase(Request("txtMvmtNo"))), " " , "S") & " "
	Else
		strVal = ""
	End If

  	If Len(Trim(Request("txtFrIvDt"))) Then
		strVal = strVal & " AND A.MVMT_RCPT_DT >= " & FilterVar(UNIConvDate(Request("txtFrIvDt")), "''", "S") & ""		
	End If		
	
	If Len(Trim(Request("txtToIvDt"))) Then
		strVal = strVal & " AND A.MVMT_RCPT_DT <= " & FilterVar(UNIConvDate(Request("txtToIvDt")), "''", "S") & ""		
	End If

    If Trim(Request("txtSupplierCd")) <> "" Then
        strVal = strVal & " AND A.BP_CD = " & FilterVar(Trim(UCase(Request("txtSupplierCd"))), " " , "S") & " "		
    ELSE
        strVal = strVal & " AND A.BP_CD = " & FilterVar("zzzzzzzzz", "''", "S") & ""		
    End If

    If Trim(Request("txtSubcontra2flg")) <> "" Then
		strVal = strVal & " AND G.SUBCONTRA2_FLG = " & FilterVar(Trim(UCase(Request("txtSubcontra2flg"))), "''", "S") & " "			' ���ְ������� 
    End If

    UNIValue(0,1) = strVal 

'================================================================================================================   
    '--------------- ������ coding part(�������,End)------------------------------------------------------
    'UNIValue(0,UBound(UNIValue,2)) = " " & UCase(Trim(lgTailList))
    UNIValue(0,2) = " ORDER BY A.MVMT_RCPT_NO DESC "
    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '��: set ADO read mode
 
End Sub
'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
Sub QueryData()
    Dim iStr
	BlankchkFlg = False
	
    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0) '* : Record Set �� ���� ���� 
    
    iStr = Split(lgstrRetMsg,gColSep)

    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If    
        
    If BlankchkFlg = False Then	
		If  rs0.EOF And rs0.BOF Then
		    Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
		    rs0.Close
		    Set rs0 = Nothing
    
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
        .ggoSpread.Source    	= .frm1.vspdData 
        .ggoSpread.SSShowData "<%=lgstrData%>"                            	'��: Display data 
        .lgStrPrevKey		    =  "<%=lgStrPrevKey%>"                      '��: set next data tag
  		.frm1.txtMvmtNo.value	=  "<%=ConvSPChars(Request("txtMvmtNo"))%>" 	
        .DbQueryOk
	End with
</Script>	
<%
    Response.End													'��: �����Ͻ� ���� ó���� ������ 
%>

