
<%@ LANGUAGE=VBSCript %>
<%Option Explicit%>
<!-- #Include file="../../inc/incSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDBAgent.inc" -->
<!-- #Include file="../../inc/incSvrDBAgentVariables.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<% 

Err.Clear
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
Dim strfrgldt	                                                           
Dim strtogldt
Dim strfrglno	                                                           
Dim strtoglno
Dim strdeptcd
	                                                           '�� : ������ 
Dim strCond
Dim strDeptNm
'--------------- ������ coding part(��������,End)----------------------------------------------------------
  
Call HideStatusWnd 
Call LoadBasisGlobalInf()
Call loadInfTB19029B("Q", "A","NOCOOKIE","QB")
'Call LoadBNumericFormatB("Q", "A","NOCOOKIE","QB")

lgStrPrevKey   = Request("lgStrPrevKey")                               '�� : Next key flag
lgMaxCount     = CInt(Request("lgMaxCount"))                           '�� : �ѹ��� �����ü� �ִ� ����Ÿ �Ǽ� 
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
    
    strDeptNm = rs0(1)
    
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
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()
    Redim UNISqlId(0)                                                     '��: SQL ID ������ ���� ����Ȯ�� 
    '--------------- ������ coding part(�������,Start)----------------------------------------------------

    Redim UNIValue(0,2)

    UNISqlId(0) = "A5106RA101"

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
        Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)
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
 
    '--------------- ������ coding part(�������,Start)----------------------------------------------------
     strfrgldt     = UNIConvDate(Request("txtfrtempgldt"))
     strtogldt     = UNIConvDate(Request("txttotempgldt"))
     strfrglno	   = Request("txtfrtempglno")
     strtoglno     = Request("txttotempglno")
     strdeptcd     = Request("txtdeptcd")
     
     
     If strfrgldt <> "" Then
		strCond = strCond & " and a.gl_dt >=  " & FilterVar(strfrgldt , "''", "S") & ""
     End If
     
     If strtogldt <> "" Then
		strCond = strCond & " and a.gl_dt <=  " & FilterVar(strtogldt , "''", "S") & ""
     End If
     
     If strfrglno <> "" Then
		strCond = strCond & " and a.gl_no >=  " & FilterVar(strfrglno , "''", "S") & ""
     End If
     
     If strtoglno <> "" Then
		strCond = strCond & " and a.gl_no <=  " & FilterVar(strtoglno , "''", "S") & ""
     End If
     
     If strdeptcd <> "" Then
		strCond = strCond & " and a.dept_cd =  " & FilterVar(strdeptcd , "''", "S") & ""
     End If
     
     iF Request("lgAuthorityFlag") = "Y" then      '���Ѱ��� �߰� 
		strCond = strCond & " and EXISTS ( SELECT 1 FROM z_usr_authority_value S WHERE a.dept_cd = S.code_value and S.usr_id =  " & FilterVar(gUsrId , "''", "S") & " AND S.module_cd = " & FilterVar("A", "''", "S") & "  )  "   '���Ѱ��� �߰� 
	 end if  '���Ѱ��� �߰� 
     
    '--------------- ������ coding part(�������,End)------------------------------------------------------

End Sub

%>

<Script Language=vbscript>
    With parent
		 IF Trim(.frm1.txtDeptCd.value) <> "" Then
			.frm1.txtDeptNm.Value = "<%=ConvSPChars(strDeptNm)%>"
		 ElseIF Trim(.frm1.txtDeptcd.value) = "" Then	
			.frm1.txtDeptNm.Value = ""
		 END IF	         
		 
         .ggoSpread.Source    = .frm1.vspdData 
         .ggoSpread.SSShowData "<%=ConvSPChars(lgstrData)%>"                            '��: Display data 
         .lgStrPrevKey        =  "<%=ConvSPChars(lgStrPrevKey)%>"                       '��: set next data tag
         .DbQueryOk
	End with
</Script>	

