<%'======================================================================================================
'*  1. Module Name          : Sales
'*  2. Function Name        : �������� 
'*  3. Program ID           : 
'*  4. Program Name         : �ŷ�ó�˾� 
'*  5. Program Desc         : �ŷ�ó������ �ŷ�ó�˾� 
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2000/12/09
'*  8. Modified date(Last)  : 2002/04/23
'*  9. Modifier (First)     : 
'* 10. Modifier (Last)      :
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              : 2000/12/09
'*                            2001/12/18  Date ǥ������ 
'*							  2002/04/12 ADO ��ȯ 
'=======================================================================================================
%>
<!-- #Include file="../../inc/incSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDBAgent.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%

Call LoadBasisGlobalInf()
Call LoadInfTB19029B("I","*","NOCOOKIE","QB")
On Error Resume Next
                                                                         
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0      '�� : DBAgent Parameter ���� 
Dim lgStrData                                                 '�� : Spread sheet�� ������ ����Ÿ�� ���� ���� 
Dim lgMaxCount                                                '�� : Spread sheet �� visible row �� 
Dim lgTailList                                                '�� : Orderby���� ���� field ����Ʈ 
Dim lgSelectList
Dim lgSelectListDT
Dim lgDataExist
Dim lgPageNo


Dim strFrDt
Dim strToDt
Dim strBpCd
Dim strBpNm
Dim strOwnRgstN

Dim strCond
Dim BlankchkFlg 											  ' ���ű׷�� 

    Call HideStatusWnd 

	lgPageNo       = UNICInt(Trim(Request("lgPageNo")),0)    '��: "0"(First),"1"(Second),"2"(Third),"3"(...)
	lgSelectList   = Request("lgSelectList")                               '�� : select ����� 
	lgTailList     = Request("lgTailList")                                 '�� : Orderby value
	lgSelectListDT = Split(Request("lgSelectListDT"), gColSep)             '�� : �� �ʵ��� ����Ÿ Ÿ�� 
	lgDataExist      = "No"
    
    Call TrimData()	    
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
	
	Const C_SHEETMAXROWS_D  = 30  
    
    
    If CLng(lgPageNo) > 0 Then
       rs0.Move     = CLng(C_SHEETMAXROWS_D) * CLng(lgPageNo)                  'lgMaxCount:Max Fetched Count at once , lgStrPrevKeyIndex : Previous PageNo
    End If
    
    iLoopCount = -1
    
   Do while Not (rs0.EOF Or rs0.BOF)
   
        iLoopCount =  iLoopCount + 1
        iRowStr = ""
        
		For ColCnt = 0 To UBound(lgSelectListDT) - 1 
            iRowStr = iRowStr & Chr(11) & FormatRsString(lgSelectListDT(ColCnt),rs0(ColCnt))
		Next
 
        If iLoopCount < C_SHEETMAXROWS_D Then
           lgstrData = lgstrData & iRowStr & Chr(11) & Chr(12)
        Else
           lgPageNo = lgPageNo + 1
           Exit Do
        End If
        
        rs0.MoveNext
	Loop

    If iLoopCount < C_SHEETMAXROWS_D Then                                 '��: Check if next data exists
       lgPageNo = ""
    End If
    rs0.Close                                                       '��: Close recordset object
    Set rs0 = Nothing	                                            '��: Release ADF

End Sub



'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()

    Redim UNISqlId(1)                                                     '��: SQL ID ������ ���� ����Ȯ�� 
'--------------- ������ coding part(�������,Start)----------------------------------------------------
    Redim UNIValue(1,4)

    UNISqlId(0) = "BpPopUpBiz"
'--------------- ������ coding part(�������,End)------------------------------------------------------
    
    UNIValue(0,1) = Trim(lgSelectList)                                      '��: Select list
    
'--------------- ������ coding part(�������,Start)----------------------------------------------------
	
	
	IF UCase(Trim(Request("lgTableNm"))) = "A_OPEN_AR" THEN
		UNIValue(0,0) = "DISTINCT "
		UNIValue(0,2) = ", A_OPEN_AR C" 
	ELSEIF UCase(Trim(Request("lgTableNm"))) = "A_OPEN_AP" THEN
		UNIValue(0,0) = "DISTINCT "
		UNIValue(0,2) = ", A_OPEN_AP C"
	ELSE
		UNIValue(0,0) = ""
		UNIValue(0,2) = ""
	END IF
  
	UNIValue(0,3) = strCond 
 '--------------- ������ coding part(�������,End)------------------------------------------------------

    UNIValue(0,UBound(UNIValue,2)) = UCase(Trim(lgTailList))					  '��: ǥ�������� �Է� 
    UNILock = DISCONNREAD :	UNIFlag = "1"										  '��: set ADO read mode
 
End Sub

'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
Sub QueryData()

    Dim lgstrRetMsg                                             '�� : Record Set Return Message �������� 
    Dim lgADF                                                   '�� : ActiveX Data Factory ���� �������� 
    Dim iStr
    BlankchkFlg = False
    
    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)

	Set lgADF   = Nothing
	
    iStr = Split(lgstrRetMsg,gColSep)
        

	If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If 
         
    If  rs0.EOF And rs0.BOF And BlankchkFlg  =  False Then
        Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
        rs0.Close
        Set rs0 = Nothing
    Else    
        Call  MakeSpreadSheetData()
    End If  
End Sub
'----------------------------------------------------------------------------------------------------------
' Set default value or preset value
'----------------------------------------------------------------------------------------------------------
Sub  TrimData()
 
    '--------------- ������ coding part(�������,Start)----------------------------------------------------
	
     strFrDt     = UCase(Trim(UNIConvDate(Request("txtFrDt"))))
     strToDt     = UCase(Trim(UNIConvDate(Request("txtToDt"))))
     
     strBpCd	 = UCase(Trim(Request("txtBp_cd")))
     strBpNm     = UCase(Trim(Request("txtBp_nm")))
     strOwnRgstN = UCase(Trim(Request("txtOwnRgstN")))
     
     
    strCond = " "
	
	IF UCase(Trim(Request("lgTableNm"))) = "A_OPEN_AR" THEN
		strCond = strCond & " AND  A.BP_CD=C.PAY_BP_CD AND C.CONF_FG = " & FilterVar("C", "''", "S") & "  AND C.AR_STS=" & FilterVar("O", "''", "S") & "  AND C.BAL_AMT <> 0 AND C.GL_NO <>'' "
		IF Request("txtFrDt") <>"" THEN 	strCond = strCond & " AND C.AR_DT >= " & FilterVar(strFrDt, "''", "S") & ""
		IF Request("txtToDt") <>"" THEN	strCond = strCond & " AND C.AR_DT <= " & FilterVar(strToDt, "''", "S") & ""

		
	ELSEIF UCase(Trim(Request("lgTableNm"))) = "A_OPEN_AP" THEN
		strCond = strCond & " AND  A.BP_CD=C.PAY_BP_CD AND C.CONF_FG = " & FilterVar("C", "''", "S") & "  AND C.AP_STS=" & FilterVar("O", "''", "S") & "  AND C.BAL_AMT <> 0 AND C.GL_NO <>'' "
		IF Request("txtFrDt") <>"" THEN 	strCond = strCond & " AND C.AP_DT >= " & FilterVar(strFrDt, "''", "S") & ""
		IF Request("txtToDt") <>"" THEN	strCond = strCond & " AND C.AP_DT <= " & FilterVar(strToDt, "''", "S") & ""
		
	END IF
	
	
	If strBpCd <> "" Then	strCond = strCond & "AND A.BP_CD LIKE  " & FilterVar("%" & strBpCd & "%", "''", "S") & ""	
	
	If strBpNm <> "" Then	strCond = strCond & " AND A.BP_NM LIKE  " & FilterVar("%" & strBpNm & "%", "''", "S") & ""				
		
	If Trim(Request("txtRadio2")) = "C" Or Trim(Request("txtRadio2")) = "S" Then
		strCond = strCond & " AND A.BP_TYPE LIKE  " & FilterVar("%" & Trim(Request("txtRadio2")) & "%", "''", "S") & ""		
	End If
	
	If Trim(Request("txtRadio3")) = "Y" Or Trim(Request("txtRadio3")) = "N" Then
		strCond = strCond & " AND A.USAGE_FLAG = " & FilterVar(Request("txtRadio3"), "''", "S") & ""		
	End If   	
	
	If strOwnRgstN <> "" Then 	strCond = strCond & " AND A.BP_RGST_NO LIKE  " & FilterVar(strOwnRgstN & "%", "''", "S") & ""

    '--------------- ������ coding part(�������,End)------------------------------------------------------

End Sub

%>
<Script Language=vbscript>
    With parent
'		
		If "<%=lgDataExist%>" = "Yes" Then
			'Set condition data to hidden area
			If "<%=lgPageNo%>" = "1" Then           ' "1" means that this query is first and next data exists
				.frm1.HBp_cd.value	= "<%=ConvSPChars(Request("txtBp_cd"))%>"
				.frm1.HBp_nm.value	= "<%=ConvSPChars(Request("txtBp_nm"))%>"
			
				.frm1.HRadio2.value	= "<%=Request("txtRadio2")%>"
				.frm1.HRadio3.value	= "<%=Request("txtRadio3")%>"					
				.frm1.HOwn_Rgst_N.value	= "<%=ConvSPChars(Request("txtOwnRgstN"))%>"					
			End If    
			'Show multi spreadsheet data from this line
			       
			.ggoSpread.Source    = .frm1.vspdData 
			.ggoSpread.SSShowData "<%=lgstrData%>"                  '��: Display data 																					
			.lgPageNo			 =  "<%=lgPageNo%>"				    '��: Next next data tag
			.DbQueryOk
		End If
	End with
</Script>	
