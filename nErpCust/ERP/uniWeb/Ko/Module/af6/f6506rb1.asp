
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

Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1                      '�� : DBAgent Parameter ���� 
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

Dim strSttlDtFr	                                                           
Dim strSttlDtTo
Dim strSttlmentNoFr
Dim strSttlmentNoTo
Dim strPrpaymNoFr
Dim strPrpaymNoTo
Dim strBpCd

Dim strCond

' ���Ѱ��� �߰� 
Dim lgAuthBizAreaCd	' ����� 
Dim lgInternalCd	' ���κμ� 
Dim lgSubInternalCd	' ���κμ�(��������)
Dim lgAuthUsrID		' ����	
'--------------- ������ coding part(��������,End)----------------------------------------------------------
  
    

    lgPageNo       = UNICInt(Trim(Request("lgPageNo")),0)                  '��: "0"(First),"1"(Second),"2"(Third),"3"(...)
    lgMaxCount     = CInt(Request("lgMaxCount"))                           '�� : �ѹ��� �����ü� �ִ� ����Ÿ �Ǽ� 
    lgSelectList   = Request("lgSelectList")                               '�� : select ����� 
    lgSelectListDT = Split(Request("lgSelectListDT"), gColSep)             '�� : �� �ʵ��� ����Ÿ Ÿ�� 
    lgTailList     = Request("lgTailList")                                 '�� : Orderby value
    lgDataExist    = "No"
    iPrevEndRow = 0
    iEndRow = 0

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

    UNISqlId(0) = "F6506RA101"
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
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1)
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
     strSttlDtFr     = UCase(Trim(UNIConvDate(Request("txtSttlDtFr"))))
     strSttlDtTo     = UCase(Trim(UNIConvDate(Request("txtSttlDtTo"))))
     strSttlmentNoFr = UCase(Trim(Request("txtSttlmentNoFr")))
     strSttlmentNoTo = UCase(Trim(Request("txtSttlmentNoTo")))
     strPrpaymNoFr	 = UCase(Trim(Request("txtPrpaymNoFr")))
     strPrpaymNoTo	 = UCase(Trim(Request("txtPrpaymNoTo")))
     strBpCd		 = UCase(Trim(Request("txtBpCd")))
     
     If strSttlDtFr <> "" Then
		strCond = strCond & " and A.STTL_DT >=  " & FilterVar(strSttlDtFr , "''", "S") & ""
     End If
     
     If strSttlDtTo <> "" Then
		strCond = strCond & " and A.STTL_DT <=  " & FilterVar(strSttlDtTo , "''", "S") & ""
     End If
     
   
     If strBpCd <> "" Then
		strCond = strCond & " and B.BP_CD =  " & FilterVar(strBpCd , "''", "S") & ""
     End If
     
     If strSttlmentNoFr <> "" Then
		strCond = strCond & " and A.STTLMENT_NO >=  " & FilterVar(strSttlmentNoFr , "''", "S") & ""
     End If
     
     If strSttlmentNoTo <> "" Then
		strCond = strCond & " and A.STTLMENT_NO <=  " & FilterVar(strSttlmentNoTo , "''", "S") & ""
     End If
     
     If strPrpaymNoFr <> "" Then
		strCond = strCond & " and A.PRPAYM_NO >=  " & FilterVar(strPrpaymNoFr , "''", "S") & ""
     End If
     
     If strPrpaymNoTo <> "" Then
		strCond = strCond & " and A.PRPAYM_NO <=  " & FilterVar(strPrpaymNoTo , "''", "S") & ""
     End If

	' ���Ѱ��� �߰� 
	If lgAuthBizAreaCd <> "" Then
		strCond		= strCond & " AND b.BIZ_AREA_CD = " & FilterVar(lgAuthBizAreaCd, "''", "S")
	End If
	
	If lgInternalCd <> "" Then
		strCond		= strCond & " AND b.INTERNAL_CD = " & FilterVar(lgInternalCd, "''", "S")
	End If
	
	If lgSubInternalCd <> "" Then
		strCond		= strCond & " AND b.INTERNAL_CD LIKE " & FilterVar(lgSubInternalCd & "%", "''", "S")
	End If
	
	If lgAuthUsrID <> "" Then
		strCond		= strCond & " AND b.INSRT_USER_ID = " & FilterVar(lgAuthUsrID, "''", "S")
	End If  
	     
    '--------------- ������ coding part(�������,End)------------------------------------------------------

End Sub

%>
<Script Language=vbscript>
    If "<%=lgDataExist%>" = "Yes" Then

       'Set condition data to hidden area
       If "<%=lgPageNo%>" = "1" Then   ' "1" means that this query is first and next data exists
          Parent.Frm1.htxtBpCd.Value			= Parent.Frm1.txtBpCd.Value
          Parent.Frm1.htxtSttlDtFr.Value		= Parent.Frm1.txtSttlDtFr.Text
          Parent.Frm1.htxtSttlDtTo.Value		= Parent.Frm1.txtSttlDtTo.Text
		  Parent.Frm1.htxtSttlmentNoFr.Value		= Parent.Frm1.txtSttlmentNoFr.Value
          Parent.Frm1.htxtSttlmentNoTo.Value		= Parent.Frm1.txtSttlmentNoTo.Value
          Parent.Frm1.htxtPrpaymNoFr.Value			= Parent.Frm1.txtPrpaymNoFr.Value
          Parent.Frm1.htxtPrpaymNoTo.Value			= Parent.Frm1.txtPrpaymNoTo.Value
       End If
       
       'Show multi spreadsheet data from this line
       
       Parent.ggoSpread.Source  = Parent.frm1.vspdData
		Parent.frm1.vspdData.Redraw = False
		Parent.ggoSpread.SSShowData "<%=lgstrData%>", "F"                    '�� : Display data
		Call Parent.ReFormatSpreadCellByCellByCurrency(Parent.Frm1.vspdData,<%=iPrevEndRow+1%>,<%=iEndRow%>,Parent.GetKeyPos("A",2),Parent.GetKeyPos("A",3),"A", "Q" ,"X","X")
		Parent.frm1.vspdData.Redraw = True
       Parent.lgPageNo      =  "<%=lgPageNo%>"               '�� : Next next data tag
       Parent.DbQueryOk
    End If   

</Script>
	
