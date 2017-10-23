<%'======================================================
'*  1. Module Name          : ���� 
'*  2. Function Name        : ������� 
'*  3. Program ID           : s3113ma9
'*  4. Program Name         : �̸���ä����Ȳ��ȸ 
'*  5. Program Desc         : 
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2002/05/30
'*  8. Modified date(Last)  : 2002/05/30
'*  9. Modifier (First)     : Hwang Seongbae
'* 10. Modifier (Last)      : Hwang Seongbae
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
On Error Resume Next                                                                         
                                                                         
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0 , rs1, rs2, rs3			   '�� : DBAgent Parameter ���� 
Dim lgStrData                                                 '�� : Spread sheet�� ������ ����Ÿ�� ���� ���� 
Dim lgMaxCount                                                '�� : Spread sheet �� visible row �� 
Dim lgTailList                                                '�� : Orderby���� ���� field ����Ʈ 
Dim lgSelectList
Dim lgSelectListDT
Dim lgDataExist
Dim lgPageNo

Call HideStatusWnd 
Call LoadBasisGlobalInf()
Call LoadInfTB19029B("Q", "S", "NOCOOKIE", "QB")  

lgPageNo       = UNICInt(Trim(Request("lgPageNo")),0)    '��: "0"(First),"1"(Second),"2"(Third),"3"(...)
lgMaxCount     = 100							             '�� : �ѹ��� �����ü� �ִ� ����Ÿ �Ǽ� 
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
	Response.Write "<Script Language=vbscript>" & vbCr
	Response.Write "With parent.frm1" & vbCr
	Response.Write ".txtHFromDt.value = """ & Request("txtFromDT") & """" & vbCr

	Response.Write ".txtHToDt.value	= """ & Request("txtToDT") & """" & vbCr
	Response.Write ".txtHPlantCd.value	= """ & ConvSPChars(Request("txtPlantCd")) & """" & vbCr
	Response.Write ".txtHSalesGrp.value	= """ & ConvSPChars(Request("txtSalesGrp")) & """" & vbCr
	Response.Write ".txtHSoldToParty.value	= """ & ConvSPChars(Request("txtSoldToParty")) & """" & vbCr

	Response.Write ".txtHSoNo.value	= """ & ConvSPChars(Request("txtSoNo")) & """" & vbCr
	Response.Write ".txtHItemCd.value	= """ & ConvSPChars(Request("txtItemCd")) & """" & vbCr
	Response.Write "End with" & vbCr
	Response.Write "</Script>" & vbCr
End Sub

'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()

    Redim UNISqlId(0)                                                     '��: SQL ID ������ ���� ����Ȯ�� 
'--------------- ������ coding part(�������,Start)----------------------------------------------------
    Redim UNIValue(0,8)

    UNISqlId(0) = "S3113MA901"
'--------------- ������ coding part(�������,End)------------------------------------------------------

    UNIValue(0,0) = Trim(lgSelectList)                                      '��: Select list
    
'--------------- ������ coding part(�������,Start)----------------------------------------------------
    If Len(Trim(Request("txtFromDt"))) Then
		UNIValue(0,1) = " " & FilterVar(UNIConvDate(Request("txtFromDt")), "''", "S") & ""                           '������ 
	Else
		UNIValue(0,1) = "Null"
	End If

    If Len(Trim(Request("txtFromDt"))) Then
		UNIValue(0,2) = " " & FilterVar(UNIConvDate(Request("txtToDt")), "''", "S") & ""									'������ 
	Else
		UNIValue(0,2) = "Null"
	End If

	If Len(Trim(Request("txtSoldToParty"))) Then
		UNIValue(0,3) = " " & FilterVar(Request("txtSoldToParty"), "''", "S") & ""		'�ֹ�ó 
    Else
		UNIValue(0,3) = "NULL"
   	End If

	If Len(Trim(Request("txtSoNo"))) Then
		UNIValue(0,4) = " " & FilterVar(Request("txtSoNo"), "''", "S") & ""				'S/O��ȣ 
    Else
		UNIValue(0,4) = "NULL"
   	End If

	If Len(Trim(Request("txtSalesGrp"))) Then
		UNIValue(0,5) = " " & FilterVar(Request("txtSalesGrp"), "''", "S") & ""			'�����׷� 
    Else
		UNIValue(0,5) = "NULL"
   	End If

	If Len(Trim(Request("txtPlantCd"))) Then
		UNIValue(0,6) = " " & FilterVar(Request("txtPlantCd"), "''", "S") & ""			'���� 
    Else
		UNIValue(0,6) = "NULL"
   	End If
	
	If Len(Trim(Request("txtItemCd"))) Then
		UNIValue(0,7) = " " & FilterVar(Request("txtItemCd"), "''", "S") & ""				'ǰ�� 
    Else
		UNIValue(0,7) = "NULL"
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
		Call parent.DbQueryOk
		</Script>	
        <%
    Else    
        Call MakeSpreadSheetData()
		If lgPageNo = "1" Then Call SetConditionData()
        Call WriteResult()
    End If  
End Sub

' ��ȸ ����� Display�ϴ� Script �ۼ� 
Sub WriteResult()
	Response.Write "<Script Language=vbscript>" & vbCr
	Response.Write "With parent.frm1" & vbCr
	Response.Write "Parent.ggoSpread.Source	= .vspdData" & vbCr
	Response.Write "parent.ggoSpread.SSShowDataByClip """ & lgstrData & """" & vbCr
	Response.Write "parent.lgPageNo	= """ & lgPageNo & """" & vbCr
	Response.Write "parent.DbQueryOk" & vbCr
	Response.Write "End with" & vbCr
	Response.Write "</Script>" & vbCr
End Sub

%>