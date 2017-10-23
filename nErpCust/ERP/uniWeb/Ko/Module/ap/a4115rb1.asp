<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrDBAgent.inc" -->
<!-- #Include file="../../inc/incSvrDBAgentVariables.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%																			'�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 

'On Error Resume Next
Err.Clear 

Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1							'�� : DBAgent Parameter ���� 
Dim lgstrData																'�� : data for spreadsheet data
Dim lgStrPrevKey															'�� : ���� �� 
Dim lgMaxCount																'�� : �ѹ��� �����ü� �ִ� ����Ÿ �Ǽ� 
Dim lgTailList																'�� : Orderby���� ���� field ����Ʈ 
Dim lgSelectList
Dim lgSelectListDT
Dim lgDataExist
Dim lgPageNo
Dim iPrevEndRow
Dim iEndRow	

'--------------- ������ coding part(��������,Start)--------------------------------------------------------
Dim LngRow
Dim GroupCount    
Dim strVal

Dim lgADF																	'�� : ActiveX Data Factory ���� �������� 
Dim lgstrRetMsg																'�� : Record Set Return Message �������� 

Dim strDeptCd
Dim strCond
Dim iChangeOrgID

' ���Ѱ��� �߰� 
Dim lgAuthBizAreaCd, lgAuthBizAreaNm			' ����� 
Dim lgInternalCd, lgDeptCd, lgDeptNm			' ���κμ� 
Dim lgSubInternalCd, lgSubDeptCd, lgSubDeptNm	' ���κμ�(��������)
Dim lgAuthUsrID, lgAuthUsrNm					' ���� 

Dim lgBizAreaAuthSQL, lgInternalCdAuthSQL, lgSubInternalCdAuthSQL, lgAuthUsrIDAuthSQL

    Call HideStatusWnd() 
	Call LoadBasisGlobalInf()    
	Call LoadInfTB19029B("*", "A", "NOCOOKIE", "RB")   

    lgPageNo       = UNICInt(Trim(Request("lgPageNo")),0)					'��: "0"(First),"1"(Second),"2"(Third),"3"(...)
    lgSelectList   = Request("lgSelectList")								'�� : select ����� 
    lgSelectListDT = Split(Request("lgSelectListDT"), gColSep)				'�� : �� �ʵ��� ����Ÿ Ÿ�� 
    lgTailList     = Request("lgTailList")									'�� : Orderby value
    lgDataExist    = "No"
    iPrevEndRow = 0
    iEndRow = 0
    
    iChangeOrgID = UCase(Trim(Request("txtOrgChangeId")))
    
    Call TrimData()
    Call FixUNISQLData()
    Call QueryData()

'===========================================================================================================
' Query Data
'===========================================================================================================
Sub MakeSpreadSheetData()
    Dim  RecordCnt
    Dim  ColCnt
    Dim  iLoopCount
    Dim  iRowStr
    
    lgDataExist = "Yes"
    lgstrData   = ""


	Const C_SHEETMAXROWS_D = 30
    
    If CInt(lgPageNo) > 0 Then
		iPrevEndRow =  C_SHEETMAXROWS_D * CDbl(lgPageNo)    
       rs0.Move= iPrevEndRow                   'lgMaxCount:Max Fetched Count at once , lgStrPrevKeyIndex : Previous PageNo

    End If

    iLoopCount = -1
    
    Do While Not (rs0.EOF Or rs0.BOF)
        iLoopCount =  iLoopCount + 1
        iRowStr = ""
		For ColCnt = 0 To UBound(lgSelectListDT) - 1 
            iRowStr = iRowStr & Chr(11) & FormatRsString(lgSelectListDT(ColCnt),rs0(ColCnt))
		Next
 
        If  iLoopCount < C_SHEETMAXROWS_D Then
            lgstrData = lgstrData & iRowStr & Chr(11) & Chr(12)
        Else
            lgPageNo = lgPageNo + 1
            Exit Do
        End If
        rs0.MoveNext
	Loop

    If  iLoopCount < C_SHEETMAXROWS_D Then                                            '��: Check if next data exists
        lgPageNo = ""                                                  '��: ���� ����Ÿ ����.
        iEndRow = iPrevEndRow + iLoopCount + 1
    Else
        iEndRow = iPrevEndRow + iLoopCount
    End If
    	
	rs0.Close
    Set rs0 = Nothing 
End Sub


'===========================================================================================================
' Set DB Agent arg
'===========================================================================================================
Sub FixUNISQLData()
	Redim UNISqlId(2)															'��: SQL ID ������ ���� ����Ȯ�� 
	Redim UNIValue(2,2)
    
	UNISqlId(0) = "A4115RA101"
	UNISqlId(1) = "ADEPTNM"
	
    UNIValue(0,0) = lgSelectList												'��: Select list
    UNIValue(0,1) = strCond
    
    UNIValue(1,0) = FilterVar(strDeptCd, "''", "S")
    UNIValue(1,1) = FilterVar(UCase(iChangeOrgID), "''", "S")   
    
    UNIValue(0,UBound(UNIValue,2)) = UCase(Trim(lgTailList))
    UNILock = DISCONNREAD :	UNIFlag = "1"										'��: set ADO read mode
End Sub

'===========================================================================================================
' Query Data
'===========================================================================================================
Sub QueryData()
    Dim iStr
    Dim strMsg
    
    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1)
    Set lgADF = Nothing
    
    iStr = Split(lgstrRetMsg,gColSep)
  
    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
        Response.End
    End If
	
	If (rs1.EOF And rs1.BOF) Then
		If strDeptCd <> "" Then
			strMsg = Request("txtDeptCd_Alt")
			Call DisplayMsgBox("970000", vbOKOnly, strMsg, "", I_MKSCRIPT)
			Response.End
%>
			<Script Language=vbScript>
			With parent
				.frm1.txtDeptCd.value = ""
				.frm1.txtDeptNm.value = ""
			End With
			</Script>
<%		
		End If
    Else
%>
		<Script Language=vbScript>
		With parent
			.frm1.txtDeptCd.value = "<%=ConvSPChars(Trim(rs1(0)))%>"
			.frm1.txtDeptNm.value = "<%=ConvSPChars(Trim(rs1(1)))%>"
		End With
		</Script>
<%
    End If
    
	rs1.Close
	Set rs1 = Nothing   
  
    If rs0.EOF And rs0.BOF Then
		Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
		rs0.Close:		Set rs0 = Nothing
		Response.End											'��: �����Ͻ� ���� ó���� ������ 
	Else
		Call  MakeSpreadSheetData()
    End If				
    
    Set rs0 = Nothing
End Sub

'===========================================================================================================
' Set default value or preset value
'===========================================================================================================
Sub  TrimData()
	Dim strFrAllcDt,strToAllcDt,strFrAllcNo,strToAllcNo
	
     strFrAllcDt = UCase(Trim(UNIConvDate(Request("txtFrAllcDt"))))
     strToAllcDt = UCase(Trim(UNIConvDate(Request("txtToAllcDt"))))
     strFrAllcNo = UCase(Trim(Request("txtFrAllcNo")))
     strToAllcNo = UCase(Trim(Request("txtToAllcNo")))
     strDeptCd	 = UCase(Trim(Request("txtDeptCd")))

	' ���Ѱ��� �߰� 
	lgAuthBizAreaCd		= Trim(Request("lgAuthBizAreaCd"))
	lgInternalCd		= Trim(Request("lgInternalCd"))
	lgSubInternalCd		= Trim(Request("lgSubInternalCd"))
	lgAuthUsrID		= Trim(Request("lgAuthUsrID"))

     strCond = " and A.PAYM_DT >= " & FilterVar(strFrAllcDt, "''", "S") & " and A.PAYM_DT <= " & FilterVar(strToAllcDt, "''", "S")
    
     If strFrAllcNo <> "" Then strCond = strCond & " and A.ref_no >= " & FilterVar(strFrAllcNo, "''", "S")
     If strToAllcNo <> "" Then strCond = strCond & " and A.ref_no <= " & FilterVar(strToAllcNo, "''", "S")
     If strDeptCd <> "" Then strCond = strCond & " and A.dept_cd = " & FilterVar(strDeptCd, "''", "S")
     
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
	
	' ���Ѱ��� �߰� 
	strCond = strCond & lgBizAreaAuthSQL & lgInternalCdAuthSQL & lgSubInternalCdAuthSQL & lgAuthUsrIDAuthSQL     
End Sub

%>
<Script Language=vbscript>
    If "<%=lgDataExist%>" = "Yes" Then

       'Set condition data to hidden area
       If "<%=lgPageNo%>" <= "1" Then   ' "1" means that this query is first and next data exists
          parent.Frm1.htxtFrAllcDt.Value  = Parent.Frm1.txtFrAllcDt.Text
          Parent.Frm1.htxtToAllcDt.Value  = Parent.Frm1.txtToAllcDt.Text
          Parent.Frm1.htxtFrAllcNo.Value  = Parent.Frm1.txtFrAllcNo.Value
          Parent.Frm1.htxtToAllcNo.Value  = Parent.Frm1.txtToAllcNo.Value
          Parent.Frm1.htxtDeptCd.Value	 = Parent.Frm1.txtDeptCd.Value
       End If
       
       'Show multi spreadsheet data from this line
       
       Parent.ggoSpread.Source  = Parent.frm1.vspdData
       Parent.frm1.vspdData.Redraw = False
       Parent.ggoSpread.SSShowData "<%=lgstrData%>", "F"                    '�� : Display data
       Call Parent.ReFormatSpreadCellByCellByCurrency(Parent.Frm1.vspdData,<%=iPrevEndRow+1%>,<%=iEndRow%>,Parent.GetKeyPos("A",3),Parent.GetKeyPos("A",2),"A", "Q" ,"X","X")
       Parent.frm1.vspdData.Redraw = True
       Parent.lgPageNo      =  "<%=lgPageNo%>"               '�� : Next next data tag
       Parent.DbQueryOk
    End If   
</Script>
	
