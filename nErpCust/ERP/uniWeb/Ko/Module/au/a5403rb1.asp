<%@ LANGUAGE=VBSCript %>
<%Option Explicit%>


<%'======================================================================================================
'*  1. Module Name          : Basic Architect
'*  2. Function Name        : ADO Template (Save)
'*  3. Program ID           :
'*  4. Program Name         :
'*  5. Program Desc         :
'*  6. Modified date(First) : 2000/11/01
'*  7. Modified date(Last)  : 2000/11/01
'*  8. Modifier (First)     : KimTaeHyun
'*  9. Modifier (Last)      : KimTaeHyun
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(��) means that "Do not change"
'=======================================================================================================

Response.Expires = -1                                                       '�� : ASP�� ĳ������ �ʵ��� �Ѵ�.
Response.Buffer = True                                                     '�� : ASP�� ���ۿ� ������� �ʰ� �ٷ� Client�� ��������.
%>

<!-- #Include file="../../inc/incSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDBAgent.inc" -->
<!-- #Include file="../../inc/incSvrDBAgentVariables.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%                                                                         '�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 
'On Error Resume Next
'Err.Clear

Call LoadBasisGlobalInf()
Call loadInfTB19029B("Q", "A","NOCOOKIE","RB")
Call HideStatusWnd 

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
Dim strfrtempgldt	                                                           
Dim strtotempgldt
Dim strfrtempglno	                                                           
Dim strtotempglno
Dim strdeptcd

Dim strCond
Dim	strDeptNm

' ���Ѱ��� �߰� 
Dim lgAuthBizAreaCd, lgAuthBizAreaNm			' ����� 
Dim lgInternalCd, lgDeptCd, lgDeptNm			' ���κμ� 
Dim lgSubInternalCd, lgSubDeptCd, lgSubDeptNm	' ���κμ�(��������)
Dim lgAuthUsrID, lgAuthUsrNm					' ���� 

Dim lgBizAreaAuthSQL, lgInternalCdAuthSQL, lgSubInternalCdAuthSQL, lgAuthUsrIDAuthSQL

'--------------- ������ coding part(��������,End)----------------------------------------------------------
  
    lgStrPrevKey   = Request("lgStrPrevKey")                               '�� : Next key flag
'   lgMaxCount     = CInt(Request("lgMaxCount"))                           '�� : �ѹ��� �����ü� �ִ� ����Ÿ �Ǽ� 
    lgSelectList   = Request("lgSelectList")                               '�� : select ����� 
    lgSelectListDT = Split(Request("lgSelectListDT"), gColSep)             '�� : �� �ʵ��� ����Ÿ Ÿ�� 
    lgTailList     = Request("lgTailList")                                 '�� : Orderby value

    Const C_SHEETMAXROWS_D  = 30 
   
    lgMaxCount = CInt(C_SHEETMAXROWS_D)                     '��: Max fetched data at a time

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
				    iStr = iStr & Chr(11) & rs0(ColCnt) 
            End Select
		Next
 
        If iRCnt < lgMaxCount Then
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

    UNISqlId(0) = "A5403RA101"

    '--------------- ������ coding part(�������,End)------------------------------------------------------
    UNIValue(0,0) = lgSelectList                                          '��: Select list
    '--------------- ������ coding part(�������,Start)----------------------------------------------------

    UNIValue(0,1) = strCond
    'UNIValue(0,2) = UCase(Trim(strtotempgldt))
    
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
    If Request("txtfrtempgldt") <> "" Then 
		strfrtempgldt = UNIConvDate(Request("txtfrtempgldt"))
    End If
    
    If Request("txttotempgldt") <> "" Then
		strtotempgldt = UNIConvDate(Request("txttotempgldt"))
    End If
     
    strfrtempglno = Request("txtfrtempglno")
    strtotempglno = Request("txttotempglno")
    strdeptcd     = Request("txtdeptcd")

	' ���Ѱ��� �߰� 
	lgAuthBizAreaCd	= Trim(Request("lgAuthBizAreaCd"))
	lgInternalCd	= Trim(Request("lgInternalCd"))
	lgSubInternalCd	= Trim(Request("lgSubInternalCd"))
	lgAuthUsrID		= Trim(Request("lgAuthUsrID"))   

    If strfrtempgldt <> "" Then
		strCond = strCond & " and a.temp_gl_dt >=  " & FilterVar(strfrtempgldt , "''", "S") & ""
    End If
     
    If strtotempgldt <> "" Then
		strCond = strCond & " and a.temp_gl_dt <=  " & FilterVar(strtotempgldt , "''", "S") & ""
    End If
     
    If strfrtempglno <> "" Then
		strCond = strCond & " and a.temp_gl_no >=  " & FilterVar(strfrtempglno , "''", "S") & ""
    End If
     
    If strtotempglno <> "" Then
		strCond = strCond & " and a.temp_gl_no <=  " & FilterVar(strtotempglno , "''", "S") & ""
    End If
     
    If strdeptcd <> "" Then
		strCond = strCond & " and a.dept_cd =  " & FilterVar(strdeptcd , "''", "S") & ""
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
		lgAuthUsrIDAuthSQL		= " AND A.UPDT_USER_ID = " & FilterVar(lgAuthUsrID, "''", "S")
	End If

	strCond = strCond & lgBizAreaAuthSQL & lgInternalCdAuthSQL & lgSubInternalCdAuthSQL & lgAuthUsrIDAuthSQL
     
    iF Request("lgAuthorityFlag") = "Y" then      '���Ѱ��� �߰� 
		strCond = strCond & " and EXISTS ( SELECT 1 FROM z_usr_authority_value S WHERE a.dept_cd = S.code_value and S.usr_id =  " & FilterVar(gUsrID , "''", "S") & " AND S.module_cd = " & FilterVar("A", "''", "S") & "  )  "   '���Ѱ��� �߰� 
	end if      '���Ѱ��� �߰� 
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
