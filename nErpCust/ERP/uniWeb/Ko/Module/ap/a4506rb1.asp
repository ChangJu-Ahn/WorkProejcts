
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
Call LoadInfTB19029B("*", "A", "NOCOOKIE", "RB")   'ggQty.DecPoint Setting...

Call HideStatusWnd 

Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2                        '�� : DBAgent Parameter ���� 
Dim lgstrData                                                              '�� : data for spreadsheet data
Dim lgStrPrevKey                                                           '�� : ���� �� 
Dim lgMaxCount                                                             '�� : �ѹ��� �����ü� �ִ� ����Ÿ �Ǽ� 
Dim lgTailList                                                             '�� : Orderby���� ���� field ����Ʈ 
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

Dim lgADF                                                                  '�� : ActiveX Data Factory ���� �������� 
Dim lgstrRetMsg                                                            '�� : Record Set Return Message �������� 

Dim strAdjustDtFr	                                                           
Dim strAdjustDtTo
Dim strAdjustNoFr
Dim strAdjustNoTo
Dim strApNoFr
Dim strApNoTo
Dim strBpCd

Dim strCond
	
'--------------- ������ coding part(��������,End)----------------------------------------------------------
  
    

    lgPageNo       = UNICInt(Trim(Request("lgPageNo")),0)                  '��: "0"(First),"1"(Second),"2"(Third),"3"(...)
    lgMaxCount     = CInt(Request("lgMaxCount"))                           '�� : �ѹ��� �����ü� �ִ� ����Ÿ �Ǽ� 
    lgSelectList   = Request("lgSelectList")                               '�� : select ����� 
    lgSelectListDT = Split(Request("lgSelectListDT"), gColSep)             '�� : �� �ʵ��� ����Ÿ Ÿ�� 
    lgTailList     = Request("lgTailList")                                 '�� : Orderby value
    lgDataExist    = "No"
    iPrevEndRow = 0
    iEndRow = 0
    
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

   Redim UNISqlId(1)                                                     '��: SQL ID ������ ���� ����Ȯ�� 
    '--------------- ������ coding part(�������,Start)----------------------------------------------------

    UNISqlId(0) = "A4506RA101"
    UNISqlId(1) = "COMMONQRY"
    
    Redim UNIValue(1,2)

    '--------------- ������ coding part(�������,End)------------------------------------------------------
    UNIValue(0,0) = lgSelectList                                          '��: Select list
    '--------------- ������ coding part(�������,Start)----------------------------------------------------

	UNIValue(0,1) = strCond
	UNIValue(1,0) = "SELECT BP_NM FROM B_BIZ_PARTNER WHERE BP_CD =  " & FilterVar(UCase(Request("txtBpCd")), "''", "S") & " "
     
    '--------------- ������ coding part(�������,End)------------------------------------------------------
    UNIValue(0,UBound(UNIValue,2)) = UCase(Trim(lgTailList))
    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '��: set ADO read mode

End Sub



'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
Sub QueryData()
    Dim iStr
    Dim strMsg
    Dim strMsg1
    Dim strMsgCd
    Dim strMsgCd1
    
    strMsg = Trim(Request("txtBpcd_Alt"))
    strMsg1 = Trim(Request("txtBizCd_Alt"))
    
    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2)
    Set lgADF = Nothing
    
    iStr = Split(lgstrRetMsg,gColSep)
  
    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
        Response.End
    End If
	
	IF NOT (rs1.EOF or rs1.BOF) then
		Response.Write " <Script Language=vbscript>	                        " & vbCr
		Response.Write " With parent.frm1   " & vbCr
		Response.Write " 	.txtBpNm.value		=	""" & ConvSPChars(rs1(0))  & """"	& vbCr 
		Response.Write " End With													" & vbCr
		Response.Write " </Script>													" & vbCr       

	ELSE
		if Trim(Request("txtBpCd")) <> "" Then
			strMsgCd = "970000"
			Response.Write " <Script Language=vbscript>								" & vbCr
			Response.Write " With parent.frm1										" & vbCr
			Response.Write " 	.txtBpNm.value		=	"""""						  & vbCr 
			Response.Write " End With												" & vbCr
			Response.Write " </Script>												" & vbCr       

		Else 
			Response.Write " <Script Language=vbscript>								" & vbCr
			Response.Write " With parent.frm1										" & vbCr
			Response.Write " 	.txtBpNm.value		=	"""""						  & vbCr 
			Response.Write " End With												" & vbCr
			Response.Write " </Script>												" & vbCr       		
		End if
	End if
    rs1.Close
    Set rs1 = Nothing 
    

	
	If  "" & Trim(strMsgCd) <> "" Then
		Call DisplayMsgBox("970000", vbOKOnly, strMsg, "", I_MKSCRIPT)
        Response.End													'��: �����Ͻ� ���� ó���� ������ 
    End If
    
  
	
    If rs0.EOF And rs0.BOF Then
		Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
		rs0.Close:		Set rs0 = Nothing
		Response.End													'��: �����Ͻ� ���� ó���� ������ 
	Else
		Call  MakeSpreadSheetData()
    End If				
    
    Set rs0 = Nothing
End Sub
'----------------------------------------------------------------------------------------------------------
' Set default value or preset value
'----------------------------------------------------------------------------------------------------------
Sub  TrimData()
	
    '--------------- ������ coding part(�������,Start)----------------------------------------------------
     strAdjustDtFr     = UCase(Trim(UNIConvDate(Request("txtAdJustDtFr"))))
     strAdjustDtTo     = UCase(Trim(UNIConvDate(Request("txtAdJustDtTo"))))
     strAdjustNoFr	   = FilterVar(UCase(Trim(Request("txtAdjustNoFr"))),"","S")
     strAdjustNoTo	   = FilterVar(UCase(Trim(Request("txtAdjustNoTo"))),"","S")
     strApNoFr	   = FilterVar(UCase(Trim(Request("txtApNoFr"))),"","S")
     strApNoTo	   = FilterVar(UCase(Trim(Request("txtApNoTo"))),"","S")
     strBpCd    = FilterVar(UCase(Trim(Request("txtBpCd"))),"","S")
     
          
     
     If strAdjustDtFr <> "" Then
		strCond = strCond & " and A.ADJUST_DT >=  " & FilterVar(strAdjustDtFr , "''", "S") & ""
     End If
     
     If strAdjustDtTo <> "" Then
		strCond = strCond & " and A.ADJUST_DT <=  " & FilterVar(strAdjustDtTo , "''", "S") & ""
     End If
     
   
     If strBpCd <> "" Then
		strCond = strCond & " and B.DEAL_BP_CD =  " & FilterVar(strBpCd , "''", "S") & ""
     End If
     
     If strAdjustNoFr <> "" Then
		strCond = strCond & " and A.ADJUST_NO >=  " & FilterVar(strAdjustNoFr , "''", "S") & ""
     End If
     
     If strAdjustNoTo <> "" Then
		strCond = strCond & " and A.ADJUST_NO <=  " & FilterVar(strAdjustNoTo , "''", "S") & ""
     End If
     
     If strApNoFr <> "" Then
		strCond = strCond & " and A.AP_NO >=  " & FilterVar(strApNoFr , "''", "S") & ""
     End If
     
     If strApNoTo <> "" Then
		strCond = strCond & " and A.AP_NO <=  " & FilterVar(strApNoTo , "''", "S") & ""
     End If
     
    '--------------- ������ coding part(�������,End)------------------------------------------------------

End Sub

%>
<Script Language=vbscript>
    If "<%=lgDataExist%>" = "Yes" Then

       'Set condition data to hidden area
       If "<%=lgPageNo%>" = "1" Then   ' "1" means that this query is first and next data exists
          Parent.Frm1.htxtBpCd.Value			= Parent.Frm1.txtBpCd.Value
          Parent.Frm1.htxtAdJustDtFr.Value		= Parent.Frm1.txtAdJustDtFr.Text
          Parent.Frm1.htxtAdJustDtTo.Value		= Parent.Frm1.txtAdJustDtTo.Text
		  Parent.Frm1.htxtAdjustNoFr.Value		= Parent.Frm1.txtAdjustNoFr.Value
          Parent.Frm1.htxtAdjustNoTo.Value		= Parent.Frm1.txtAdjustNoTo.Value
          Parent.Frm1.htxtApNoFr.Value			= Parent.Frm1.txtApNoFr.Value
          Parent.Frm1.htxtApNoTo.Value			= Parent.Frm1.txtApNoTo.Value
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
	
