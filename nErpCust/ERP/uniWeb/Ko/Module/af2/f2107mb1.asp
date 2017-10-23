
<%@ LANGUAGE=VBSCript %>
<%Option Explicit%>

<%
'======================================================================================================
'*  1. Module Name          : Accounting
'*  2. Function Name        : Treasury - Budget
'*  3. Program ID           : f2107mb1
'*  4. Program Name         : �������������ȸ 
'*  5. Program Desc         : Query of Budget Result by Monthly
'*  6. Comproxy List        : ADO
'*  7. Modified date(First) : 
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Oh, Soo Min 
'* 10. Modifier (Last)      : 
'* 11. Comment              :2003/06/12 lgMaxCount ������ ����(MA�� C_SHEETMAXROWS_D ����)
'=======================================================================================================
													'�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 



'�� : �׻� ���� ���̵� ������ �������� �²���(<)% �� %�첩��(>)�� New Line�� ��ġ�Ͽ� 
'	  ���� ���̵� ������ Ŭ���̾�Ʈ ���̵� ������ ��ġ�� ������ �� �ֵ��� �Ѵ�.
'�� : �Ʒ� HTML ������ ����Ǿ�� �ȵȴ�. 
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
%>	
<%                                                          '�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 

On Error Resume Next

Dim lgADF                                                   '�� : ActiveX Data Factory ���� �������� 
Dim lgstrRetMsg                                             '�� : Record Set Return Message �������� 
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1          '�� : DBAgent Parameter ���� 
Dim lgStrData                                               '�� : Spread sheet�� ������ ����Ÿ�� ���� ���� 
Dim lgStrPrevKey                                            '�� : ���� �� 
Dim lgMaxCount                                              '�� : Spread sheet �� visible row �� 
Dim lgTailList
Dim lgSelectList
Dim lgSelectListDT

'--------------- ������ coding part(��������,Start)----------------------------------------------------
Dim strBdgYear
Dim strDeptCd, strInternalCd
Dim strWhere
Dim strMsgCd, strMsg1, strMsg2


' ���Ѱ��� �߰� 
Dim lgAuthBizAreaCd, lgAuthBizAreaNm			' ����� 
Dim lgInternalCd, lgDeptCd, lgDeptNm			' ���κμ�		
Dim lgSubInternalCd, lgSubDeptCd, lgSubDeptNm	' ���κμ�(��������)				
Dim lgAuthUsrID, lgAuthUsrNm					' ���� 

Dim lgBizAreaAuthSQL, lgInternalCdAuthSQL, lgSubInternalCdAuthSQL, lgAuthUsrIDAuthSQL					


'--------------- ������ coding part(��������,End)------------------------------------------------------

	Call HideStatusWnd 

	'-----------------------------------------------------------------------
	'�ʵ尹���� ���Ƽ� RunMyBizASP ��� ExecMyBizASP ������...	
	'-----------------------------------------------------------------------
	lgStrPrevKey     = Request("txtPrevKey")
	lgMaxCount       = 100
	lgSelectList     = Request("txtSelectList")
	lgTailList       = Request("txtTailList")
	lgSelectListDT   = Split(Request("txtSelectListDT"), gColSep)


	strBdgYear = Request("txtBdgYear")
	strDeptCd  = UCase(Trim(Request("txtDeptCd")))


	' ���Ѱ��� �߰� 
	lgAuthBizAreaCd		= Trim(Request("txtAuthBizAreaCd"))	
	lgInternalCd		= Trim(Request("txtInternalCd"))	
	lgSubInternalCd		= Trim(Request("txtSubInternalCd"))	
	lgAuthUsrID			= Trim(Request("txtAuthUsrID"))


     Call  TrimData()                                                     '�� : Parent�� ������ ����Ÿ ���� 
     Call  FixUNISQLData()                                                '�� : DB-Agent�� ���� parameter ����Ÿ set
     call  QueryData()                                                    '�� : DB-Agent�� ���� ADO query

'----------------------------------------------------------------------------------------------------------
' Make srpread sheet data
'----------------------------------------------------------------------------------------------------------
Sub MakeSpreadSheetData()

    Dim iCnt
    Dim iRCnt                                                                     
    Dim strTmpBuffer                                                              
    Dim iStr
    Dim ColCnt
     
    iCnt = 0
    lgstrData = ""
   
    If Len(Trim(lgStrPrevKey)) Then                                              '�� : Chnage str into int
       If Isnumeric(lgStrPrevKey) Then
          iCnt = CInt(lgStrPrevKey)
       End If   
    End If   

    For iRCnt = 1 to iCnt  *  lgMaxCount                                         '�� : Discard previous data
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

    If  iRCnt < lgMaxCount Then                                     '��: Check if next data exists
        lgStrPrevKey = ""
    End If

'	rs0.Close                                                       '��: Close recordset object
'	Set rs0 = Nothing	                                            '��: Release ADF
'	Set lgADF = Nothing                                             '��: Release ADF

End Sub

'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()

    Redim UNISqlId(1)                                                     '��: SQL ID ������ ���� ����Ȯ�� 
'--------------- ������ coding part(�������,Start)----------------------------------------------------
    Redim UNIValue(1,2)                                                  '��: DB-Agent�� ���۵� parameter�� ���� ���� 
                                                                          '    parameter�� ���� ���� ������ 
     UNISqlId(0) = "F2107MA101"
     UNISqlId(1) = "F2107MA103"	'�μ��ڵ� 
'--------------- ������ coding part(�������,End)------------------------------------------------------

     UNIValue(0,0) = Trim(lgSelectList)		                              '��: Select ������ Summary    �ʵ� 

'--------------- ������ coding part(�������,Start)----------------------------------------------------

     UNIValue(0,1)  = strWhere
     
     UNIValue(1,0)  = FilterVar(strDeptCd, "''", "S")
     UNIValue(1,1)  = FilterVar(Request("hOrgChangeId"), "''", "S")
'--------------- ������ coding part(�������,End)----------------------------------------------------
     
     UNIValue(0,UBound(UNIValue,2)) = Trim(lgTailList)	'---Order By ���� 

     UNILock = DISCONNREAD :	UNIFlag = "1"                                 '��: set ADO read mode
 
End Sub

'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
Sub QueryData()
    Dim iStr

    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1)
    
    iStr = Split(lgstrRetMsg,gColSep)
    
    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If    
    
	If (rs1.EOF And rs1.BOF) Then
		If strMsgCd = "" And strDeptCd <> "" Then
			strMsgCd = "970000"
			strMsg1 = Request("txtDeptCd_Alt")
		End If
	Else
%>
	<Script Language=vbScript>
		With parent.frm1
			.txtDeptCd.value = "<%=ConvSPChars(strDeptCd)%>"
			.txtDeptNm.value = "<%=ConvSPChars(Trim(rs1(1)))%>"
		End With
	</Script>
<%
    End If
    
	rs1.Close
	Set rs1 = Nothing

    If  rs0.EOF And rs0.BOF Then
		If strMsgCd = "" Then strMsgCd = "900014"
'		Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)		'No Data Found.
'		rs0.Close
'		Set rs0 = Nothing
'		Set lgADF = Nothing
'		Response.End													'��: �����Ͻ� ���� ó���� ������ 
    Else    
        Call  MakeSpreadSheetData()
    End If
    
	rs0.Close                                                       '��: Close recordset object
	Set rs0 = Nothing	                                            '��: Release ADF
	Set lgADF = Nothing                                             '��: Release ADF

	If strMsgCd <> "" Then
		Call DisplayMsgBox(strMsgCd, vbOKOnly, strMsg1, strMsg2, I_MKSCRIPT)
		Response.End 
	End If
	
End Sub

'----------------------------------------------------------------------------------------------------------
' Set default value or preset value
'----------------------------------------------------------------------------------------------------------
Sub TrimData()

'--------------- ������ coding part(�������,Start)----------------------------------------------------
	strBdgYear = Request("txtBdgYear")
'	strDeptCd  = UCase(Trim(FilterVar(Request("txtDeptCd"),"","S")))
	strDeptCd  = UCase(Trim(Request("txtDeptCd")))
	strWhere = ""
	strWhere = strWhere & " and A.bdg_yyyymm between  " & FilterVar(strBdgYear & "01", "''", "S") & " and  " & FilterVar(strBdgYear & "12", "''", "S") & " "
	

	If strDeptCd <> "" Then
		strInternalCd = fnGetInternalCd()
		strWhere = strWhere & " and A.internal_cd =  " & FilterVar(strInternalCd , "''", "S") & " "
	End If


	' ���Ѱ��� �߰� 
	If lgAuthBizAreaCd <> "" Then			
		lgBizAreaAuthSQL		= " AND A.BIZ_AREA_CD = " & FilterVar(lgAuthBizAreaCd, "''", "S")  		
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



'--------------- ������ coding part(�������,End)------------------------------------------------------

End Sub

'----------------------------------------------------------------------------------------------------------
' Trim string and set string to space if string length is zero
'----------------------------------------------------------------------------------------------------------
'Function FilterVar(Byval str,Byval strALT)
 '    Dim strL
  '   strL = UCASE(Trim(str))
   '  If Len(strL) Then
    '    FilterVar = "'" & strL  & "'"
    ' Else
     '   FilterVar = strALT   
     'End If
'End Function

'--------------------------------------------
'���κμ��ڵ� select
'--------------------------------------------
Function fnGetInternalCd()
    Redim UNISqlId(0)                                                     '��: SQL ID ������ ���� ����Ȯ�� 
    Redim UNIValue(0,1)

    UNISqlId(0) = "F2107MA102"

    UNIValue(0,0) = FilterVar(strDeptCd, "''", "S") 
    UNIValue(0,1) = FilterVar(Request("hOrgChangeId"), "''", "S") 'GetGlobalInf("gChangeOrgId")
    
    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '��: set ADO read mode
	
    Dim iStr

    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)
    
    iStr = Split(lgstrRetMsg,gColSep)
    
    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If    
        
    If  rs0.EOF And rs0.BOF Then
        rs0.Close
        Set rs0 = Nothing

		strMsgCd = "970000"
		strMsg1 = Request("txtDeptCd_Alt")
		Call DisplayMsgBox(strMsgCd, vbOKOnly, strMsg1, strMsg2, I_MKSCRIPT)

        fnGetInternalCd = ""

        Response.End													'��: �����Ͻ� ���� ó���� ������ 
    Else    
        fnGetInternalCd = rs0(0)
    End If
End Function

%>

<Script Language=vbscript>
    
    With Parent
         .ggoSpread.Source  = .frm1.vspdData
         .ggoSpread.SSShowData "<%=lgstrData%>"                  '�� : Display data
         .lgStrPrevKey      =  "<%=ConvSPChars(lgStrPrevKey)%>"               '�� : Next next data tag
         Call .DbQueryOk         
	End with
</Script>	

<%
	Response.End													'��: �����Ͻ� ���� ó���� ������ 
%>

