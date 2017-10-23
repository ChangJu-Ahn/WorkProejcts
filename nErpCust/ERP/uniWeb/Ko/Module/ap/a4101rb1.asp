
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrDBAgent.inc" -->
<!-- #Include file="../../inc/incSvrDBAgentVariables.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%             

On Error Resume Next
Err.Clear

Call LoadBasisGlobalInf()
Call LoadInfTB19029B("Q", "A", "NOCOOKIE", "RB")   'ggQty.DecPoint Setting...

Call HideStatusWnd 

Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2                         '�� : DBAgent Parameter ���� 
Dim lgstrData                                                              '�� : data for spreadsheet data
Dim lgStrPrevKey                                                           '�� : ���� �� 
Dim lgMaxCount                                                             '�� : �ѹ��� �����ü� �ִ� ����Ÿ �Ǽ� 
Dim lgTailList                                                             '�� : Orderby���� ���� field ����Ʈ 
Dim lgSelectList
Dim lgSelectListDT
Dim lgDataExist
Dim lgPageNo


'--------------- ������ coding part(��������,Start)--------------------------------------------------------
	Dim strFrApDt	                                                           
	Dim strToApDt
	Dim strFrApNo	                                                           
	Dim strToApNo
	Dim strdeptcd
	DIm strDealBpCd
	
	Dim strCond

	Dim strMsgCd
	Dim strMsg1
    Dim iPrevEndRow
    Dim iEndRow

' ���Ѱ��� �߰� 
Dim lgAuthBizAreaCd, lgAuthBizAreaNm			' ����� 
Dim lgInternalCd, lgDeptCd, lgDeptNm			' ���κμ� 
Dim lgSubInternalCd, lgSubDeptCd, lgSubDeptNm	' ���κμ�(��������)
Dim lgAuthUsrID, lgAuthUsrNm					' ���� 

Dim lgBizAreaAuthSQL, lgInternalCdAuthSQL, lgSubInternalCdAuthSQL, lgAuthUsrIDAuthSQL


'--------------- ������ coding part(��������,End)----------------------------------------------------------
  
	lgPageNo		= UNICInt(Trim(Request("lgPageNo")),0)                  '��: "0"(First),"1"(Second),"2"(Third),"3"(...)
	lgMaxCount		= CInt(Request("lgMaxCount"))                           '�� : �ѹ��� �����ü� �ִ� ����Ÿ �Ǽ� 
	lgSelectList	= Request("lgSelectList")                               '�� : select ����� 
	lgSelectListDT	= Split(Request("lgSelectListDT"), gColSep)             '�� : �� �ʵ��� ����Ÿ Ÿ�� 
	lgTailList		= Request("lgTailList")                                 '�� : Orderby value
	lgDataExist		= "No"
	iPrevEndRow		= 0
	iEndRow			= 0

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
    Dim  iLoopCount
    Dim  iRowStr

    lgDataExist    = "Yes"
    lgstrData      = ""
    iPrevEndRow = 0

    If CDbl(lgPageNo) > 0 Then
		iPrevEndRow = CDbl(lgMaxCount) * CDbl(lgPageNo)    
       rs0.Move= iPrevEndRow                   'lgMaxCount:Max Fetched Count at once , lgStrPrevKeyIndex : Previous PageNo
    End If

    iLoopCount = -1
    
    Do while Not (rs0.EOF Or rs0.BOF)
        iLoopCount =  iLoopCount + 1
        iRowStr = ""
		For ColCnt = 0 To UBound(lgSelectListDT) - 1 
            iRowStr = iRowStr & Chr(11) & FormatRsString(lgSelectListDT(ColCnt),rs0(ColCnt))
		Next
 
        If  iLoopCount < lgMaxCount Then
            lgstrData      = lgstrData      & iRowStr & Chr(11) & Chr(12)
        Else
            lgPageNo = lgPageNo + 1
            Exit Do
        End If
        rs0.MoveNext
	Loop

    If  iLoopCount < lgMaxCount Then                                            '��: Check if next data exists
        lgPageNo = ""                                                  '��: ���� ����Ÿ ����.
        iEndRow = iPrevEndRow + iLoopCount + 1
    Else
        iEndRow = iPrevEndRow + iLoopCount
    End If
  	
  	
	rs0.Close
    Set rs0 = Nothing 
End Sub
'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()

    Redim UNISqlId(2)                                                     '��: SQL ID ������ ���� ����Ȯ�� 
    '--------------- ������ coding part(�������,Start)----------------------------------------------------

    Redim UNIValue(2,2)                                                  '��: DB-Agent�� ���۵� parameter�� ���� ���� 

    UNISqlId(0) = "A4101RA101"
    UNISqlId(1) = "ADEPTNM"
	UNISqlId(2) = "ABPNM"
    '--------------- ������ coding part(�������,End)------------------------------------------------------
    UNIValue(0,0) = lgSelectList                                          '��: Select list
    '--------------- ������ coding part(�������,Start)----------------------------------------------------

	UNIValue(0,1) = strCond
    UNIValue(1,0) = FilterVar(strDeptCd, "''", "S")
    UNIValue(1,1) = FilterVar(UCase(Request("txtOrgChangeId")), "''", "S")
    UNIValue(2,0) = FilterVar(strDealBpCd, "''", "S")
    
    '--------------- ������ coding part(�������,End)------------------------------------------------------
    UNIValue(0,UBound(UNIValue,2)) = UCase(Trim(lgTailList))
    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '��: set ADO read mode

End Sub
'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
Sub QueryData()

    Dim lgstrRetMsg                                                            '�� : Record Set Return Message �������� 
    Dim iStr
    Dim lgADF                                                                  '�� : ActiveX Data Factory ���� �������� 

    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2)
    
    Set lgADF = Nothing                                                    '��: ActiveX Data Factory Object Nothing
    
    iStr = Split(lgstrRetMsg,gColSep)
    
    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If    

	If (rs1.EOF And rs1.BOF) Then
		If strMsgCd = "" And strDeptCd <> "" Then
			strMsgCd = "970000"		'Not Found
			strMsg1 = Request("txtDeptCd_Alt")
%>
		<Script Language=vbScript>
		With parent
			.frm1.txtDeptNm.value = ""
		End With
		</Script>
<%			
		End If
    Else
%>
		<Script Language=vbScript>
		With parent
			.frm1.txtDeptCd.value = "<%=Trim(rs1(0))%>"
			.frm1.txtDeptNm.value = "<%=Trim(rs1(1))%>"
		End With
		</Script>
<%
    End If
    
	Set rs1 = Nothing 

    If  "" & Trim(strMsgCd) <> "" Then
		Call DisplayMsgBox("970000", vbOKOnly, strMsg1, "", I_MKSCRIPT)
        Response.End													'��: �����Ͻ� ���� ó���� ������ 
    End If
    
    If strDealBpCd <> "" Then
		If Not (rs2.EOF OR rs2.BOF) Then
%>
		<Script Language=vbScript>
			With parent
				.frm1.txtDealBpNm.value = "<%=Trim(rs2(1))%>"
			End With
		</Script>
<%		
		Else
%>
		<Script Language=vbScript>
			With parent
				.frm1.txtDealBpNm.value = ""
			End With
		</Script>
<%		
			Call DisplayMsgBox("126100", vbOKOnly, "", "", I_MKSCRIPT)		'No Data Found!!
	        rs2.Close
		    Set rs2 = Nothing
			Exit sub
		End IF
	rs2.Close
	Set rs2 = Nothing
	End If
    
    If  rs0.EOF And rs0.BOF Then
		Call DisplayMsgBox("990007", vbOKOnly, "", "", I_MKSCRIPT)
        rs0.Close
        Set rs0 = Nothing
        Exit Sub
    Else    
        Call  MakeSpreadSheetData()
    End If
End Sub

'----------------------------------------------------------------------------------------------------------
' Set default value or preset value
'----------------------------------------------------------------------------------------------------------
Sub  TrimData()
 
    '--------------- ������ coding part(�������,Start)----------------------------------------------------
     strFrApDt     = UCase(Trim(UNIConvDate(Request("txtFrApDt"))))
     strToApDt     = UCase(Trim(UNIConvDate(Request("txtToApDt"))))
     strFrApNo	   = UCase(Trim(Request("txtFrApNo")))
     strToApNo     = UCase(Trim(Request("txtToApNo")))
     strdeptcd     = UCase(Trim(Request("txtdeptcd")))
     strDealBpCd   = UCase(Trim(Request("txtDealBpCd")))
     
     strCond  = " and A.AP_TYPE = " & FilterVar("NP", "''", "S") & " "
     
     If strFrApDt <> "" Then strCond = strCond & " and A.AP_DT >= " & FilterVar(strFrApDt ,null,"S") 

     If strToApDt <> "" Then strCond = strCond & " and A.AP_DT <= " & FilterVar(strToApDt ,null,"S") 
     
     If strFrApNo <> "" Then strCond = strCond & " and A.AP_NO >= " & FilterVar(strFrApNo ,null,"S") 
     
     If strToApNo <> "" Then strCond = strCond & " and A.AP_NO <= " & FilterVar(strToApNo ,null,"S") 

     If strdeptcd <> "" Then strCond = strCond & " and A.dept_cd = " & FilterVar(strdeptcd ,null,"S") 
   
     If strDealBpCd <> "" Then	strCond = strCond & " and A.deal_bp_cd = " & FilterVar(strDealBpCd ,null,"S") 


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

	strCond		=	strCond & lgBizAreaAuthSQL & lgInternalCdAuthSQL & lgSubInternalCdAuthSQL & lgAuthUsrIDAuthSQL


    '--------------- ������ coding part(�������,End)------------------------------------------------------

End Sub

%>

<Script Language=vbscript>
    If "<%=lgDataExist%>" = "Yes" Then

       'Set condition data to hidden area
       If "<%=lgPageNo%>" = "1" Then   ' "1" means that this query is first and next data exists
          Parent.Frm1.htxtFrApDt.Value			= Parent.Frm1.txtFrApDt.Text
          Parent.Frm1.htxtToApDt.Value			= Parent.Frm1.txtToApDt.Text
          Parent.Frm1.htxtFrApNo.Value			= Parent.Frm1.txtFrApNo.Value
          Parent.Frm1.htxtToApNo.Value			= Parent.Frm1.txtToApNo.Value
          Parent.Frm1.htxtdeptcd.Value			= Parent.Frm1.txtdeptcd.Value
          Parent.Frm1.htxtDealBpCd.Value		= Parent.Frm1.txtDealBpCd.Value
       End If
       
       'Show multi spreadsheet data from this line
       
       Parent.ggoSpread.Source  = Parent.frm1.vspdData
       Parent.frm1.vspdData.Redraw = False
       Parent.ggoSpread.SSShowData "<%=lgstrData%>", "F"                    '�� : Display data

       Call Parent.ReFormatSpreadCellByCellByCurrency(Parent.Frm1.vspdData,<%=iPrevEndRow+1%>,<%=iEndRow%>,Parent.GetKeyPos("A",5),Parent.GetKeyPos("A",6),"A", "Q" ,"X","X")
       Call Parent.ReFormatSpreadCellByCellByCurrency(Parent.Frm1.vspdData,<%=iPrevEndRow+1%>,<%=iEndRow%>,Parent.GetKeyPos("A",5),Parent.GetKeyPos("A",7),"A", "Q" ,"X","X")
       Call Parent.ReFormatSpreadCellByCellByCurrency(Parent.Frm1.vspdData,<%=iPrevEndRow+1%>,<%=iEndRow%>,Parent.GetKeyPos("A",5),Parent.GetKeyPos("A",14),"A", "Q" ,"X","X")
       Parent.frm1.vspdData.Redraw = True

       Parent.lgPageNo      =  "<%=lgPageNo%>"               '�� : Next next data tag
       Parent.DbQueryOk
    End If   

</Script>	

