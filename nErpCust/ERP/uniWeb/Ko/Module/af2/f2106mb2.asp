
<%@ LANGUAGE=VBSCript %>
<%Option Explicit%>

<%'======================================================================================================
'*  1. Module Name          : Accounting
'*  2. Function Name        : Treasury - Budget
'*  3. Program ID           : f2106mb2
'*  4. Program Name         : ���������ȸ 
'*  5. Program Desc         : Query of Budget Result
'*  6. Comproxy List        : ADO
'*  7. Modified date(First) : 2001.02.12
'*  8. Modified date(Last)  : 2001.03.05
'*  9. Modifier (First)     : Song, Mun Gil
'* 10. Modifier (Last)      : Song, Mun Gil
'* 11. Comment              :
'*   - 2001.03.21  Song,Mun Gil  �μ��ڵ�, �����ڵ� ���� Check
'=======================================================================================================


%>

<!-- #Include file="../../inc/IncSvrMain.asp"  -->
<!-- #Include file="../../inc/incSvrDate.inc"  -->
<!-- #Include file="../../inc/incSvrDBAgent.inc"  -->
<!-- #Include file="../../inc/incSvrDBAgentVariables.inc"  -->

<!-- #Include file="../../inc/IncSvrNumber.inc"  -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<% 
	Call LoadBasisGlobalInf()     
    Call LoadInfTB19029B("Q", "A","NOCOOKIE","QB")
    Call LoadBNumericFormatB("Q", "A","NOCOOKIE","QB")
    
%><%                                                                       '�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 

'On Error Resume Next

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
Dim strBdgCd
Dim strBdgYymm
Dim strDeptCd, strOrgChangeId , strInternalCd
Dim strColYymm, strDateType
Dim strWhere

' ���Ѱ��� �߰� 
Dim lgAuthBizAreaCd, lgAuthBizAreaNm			' ����� 
Dim lgInternalCd, lgDeptCd, lgDeptNm			' ���κμ�		
Dim lgSubInternalCd, lgSubDeptCd, lgSubDeptNm	' ���κμ�(��������)				
Dim lgAuthUsrID, lgAuthUsrNm					' ���� 

Dim lgBizAreaAuthSQL, lgInternalCdAuthSQL, lgSubInternalCdAuthSQL, lgAuthUsrIDAuthSQL					



'--------------- ������ coding part(��������,End)----------------------------------------------------------
  
    Call HideStatusWnd 
	
	Const C_SHEETMAXROWS_D = 100
    
    lgStrPrevKey   = Request("lgStrPrevKey")                               '�� : Next key flag
    lgMaxCount     = C_SHEETMAXROWS_D										'�� : �ѹ��� �����ü� �ִ� ����Ÿ �Ǽ� 
    lgSelectList   = Request("lgSelectList")                               '�� : select ����� 
    lgSelectListDT = Split(Request("lgSelectListDT"), gColSep)             '�� : �� �ʵ��� ����Ÿ Ÿ�� 
    lgTailList     = Request("lgTailList")                                 '�� : Orderby value


	' ���Ѱ��� �߰� 
	lgAuthBizAreaCd		= Trim(Request("lgAuthBizAreaCd"))		
	lgInternalCd		= Trim(Request("lgInternalCd"))	
	lgSubInternalCd		= Trim(Request("lgSubInternalCd"))	
	lgAuthUsrID			= Trim(Request("lgAuthUsrID"))

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
End Sub
'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()

    Redim UNISqlId(0)                                                     '��: SQL ID ������ ���� ����Ȯ�� 
    '--------------- ������ coding part(�������,Start)----------------------------------------------------

    Redim UNIValue(0,2)

    UNISqlId(0) = "F2106MA102"

    '--------------- ������ coding part(�������,End)------------------------------------------------------
    UNIValue(0,0) = lgSelectList                                          '��: Select list
    '--------------- ������ coding part(�������,Start)----------------------------------------------------
	UNIValue(0,1) = UCase(Trim(strWhere))

'	UNIValue(0,1) = FilterVar(UCase(strBdgCd), "''", "S")
'	UNIValue(0,2) = FilterVar(UCase(strBdgYymm), "''", "S")
'	UNIValue(0,3) = FilterVar(UCase(strInternalCd), "''", "S")    
'	UNIValue(0,3) = FilterVar(UCase(strDeptCd), "''", "S")
'	UNIValue(0,4) = FilterVar(UCase(strOrgChangeId), "''", "S")
    
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
'		If strMsgCd = "" Then strMsgCd = "900014"		
'		rs0.Close
'		Set rs0 = Nothing
'		Set lgADF = Nothing
'		Response.End													'��: �����Ͻ� ���� ó���� ������ 
    Else    
        Call  MakeSpreadSheetData()                
    End If    
    
	rs0.Close
	Set rs0 = Nothing 
	Set lgADF = Nothing                                                    '��: ActiveX Data Factory Object Nothing
	
End Sub

'----------------------------------------------------------------------------------------------------------
' Set default value or preset value
'----------------------------------------------------------------------------------------------------------
Sub TrimData()
 
    '--------------- ������ coding part(�������,Start)----------------------------------------------------
	strBdgCd		= Request("txtBdgCd")
	strBdgYymm		= Request("txtBdgYymm")
	strDeptCd		= Request("txtDeptCd")
	strOrgChangeId	= Request("txtOrgChangeId")
	strInternalCd   = Request("txtinternalCd")
	strColYymm		= Request("txtColYymm")
	strDateType		= Request("txtDateType")
    '--------------- ������ coding part(�������,End)------------------------------------------------------

	' AND ((DEPT_CTRL_FLAG = 'N' AND INTERNAL_CD = @INTERNAL_CD) OR
'	             (DEPT_CTRL_FLAG = 'Y' AND INTERNAL_CD LIKE SUBSTRING(@INTERNAL_CD, 1, Len(@INTERNAL_CD)-1)+'%'))

	strWhere = ""
	strWhere	= strWhere & " where A.bdg_cd =  " & FilterVar(strBdgCd, "''", "S") & " and A.bdg_yyyymm =  " & FilterVar(strBdgYymm, "''", "S") _
						   & " and A.internal_cd = " & FilterVar(strInternalCd, "''", "S") & " "	
    
	' ���Ѱ��� �߰� 
	If lgAuthBizAreaCd <> "" Then			
		lgBizAreaAuthSQL		= " AND C.BIZ_AREA_CD = " & FilterVar(lgAuthBizAreaCd, "''", "S")  		
	End If			

	If lgInternalCd <> "" Then			
		lgInternalCdAuthSQL		= " AND A.INTERNAL_CD = " & FilterVar(lgInternalCd, "''", "S")  		
	End If			

	If lgSubInternalCd <> "" Then	
		lgSubInternalCdAuthSQL	= " AND A.INTERNAL_CD LIKE " & FilterVar(lgSubInternalCd & "%", "''", "S")  
	End If	

	If lgAuthUsrID <> "" Then	
		lgAuthUsrIDAuthSQL		= " AND A.INSRT_USER_ID = " & FilterVar(lgAuthUsrID, "''", "S")  
	End If	

	strWhere	= strWhere	& lgBizAreaAuthSQL & lgInternalCdAuthSQL & lgSubInternalCdAuthSQL & lgAuthUsrIDAuthSQL	

End Sub

%>

<Script Language=vbscript>

    With parent
         .ggoSpread.Source		= .frm1.vspdData2 
         .ggoSpread.SSShowData "<%=lgstrData%>"                         '��: Display data 
         .lgStrPrevKey_B		= "<%=ConvSPChars(lgStrPrevKey)%>"                       '��: set next data tag
	End with
</Script>	


