
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
Call LoadInfTB19029B("I", "A", "NOCOOKIE", "RB")   'ggQty.DecPoint Setting...

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
Dim AAmt, PAmt  

'--------------- ������ coding part(��������,Start)--------------------------------------------------------
Dim LngRow
Dim GroupCount    
Dim strVal

Dim lgADF                                                                  '�� : ActiveX Data Factory ���� �������� 
Dim lgstrRetMsg                                                            '�� : Record Set Return Message �������� 

Dim strFrApDt	                                                           
Dim strToApDt
Dim strDocCur                                                          
Dim strPayBpCd
Dim strBizCd
Dim strAllcDt

Dim strCond
Dim iPrevEndRow
Dim iEndRow	
'--------------- ������ coding part(��������,End)----------------------------------------------------------
 Dim BP_NM
 Dim BIZ_AREA_NM 
 
 ' ���Ѱ��� �߰� 
Dim lgAuthBizAreaCd, lgAuthBizAreaNm			' ����� 
Dim lgInternalCd, lgDeptCd, lgDeptNm			' ���κμ� 
Dim lgSubInternalCd, lgSubDeptCd, lgSubDeptNm	' ���κμ�(��������)
Dim lgAuthUsrID, lgAuthUsrNm					' ���� 

Dim lgBizAreaAuthSQL, lgInternalCdAuthSQL, lgSubInternalCdAuthSQL, lgAuthUsrIDAuthSQL

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

Const ConDate = "1899/12/30"

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
	
	If AAmt <> 0 Then

		Do While Not (Rs0.EOF Or Rs0.BOF)
			PAmt = AAmt 
			AAmt = AAmt - UNIConvNum(Rs0(6) ,0)

		    iRowStr = ""
			For ColCnt = 0 To UBound(lgSelectListDT) - 1
				If ColCnt = 6  Then '�����ݾ� ���� 
					If AAmt > 0 Then 
						iRowStr = iRowStr & Chr(11) & FormatRsString(lgSelectListDT(ColCnt),rs0(ColCnt))
					Else
						iRowStr = iRowStr & Chr(11) & FormatRsString(lgSelectListDT(ColCnt),PAmt)
					End If
				Else					
					iRowStr = iRowStr & Chr(11) & FormatRsString(lgSelectListDT(ColCnt),rs0(ColCnt))
				End If		
			Next
 
			lgstrData = lgstrData & iRowStr & Chr(11) & Chr(12)
			iEndRow = iLoopCount
			
			iLoopCount = iLoopCount + 1

			If AAmt <= 0 Then 
				Exit Do
			End If	
			        
			rs0.MoveNext
		Loop
	Else
	
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
	
	End If
	
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

    UNISqlId(0) = "a4105ra101"
	UNISqlId(1) = "COMMONQRY"
    UNISqlId(2) = "COMMONQRY"
    
    Redim UNIValue(2,2)

    '--------------- ������ coding part(�������,End)------------------------------------------------------
    UNIValue(0,0) = lgSelectList                                          '��: Select list
    '--------------- ������ coding part(�������,Start)----------------------------------------------------

	UNIValue(0,1) = strCond  
	UNIValue(1,0) = "SELECT BP_NM FROM B_BIZ_PARTNER WHERE BP_CD =  " & FilterVar(UCase(Request("txtBpCd")), "''", "S") & " "
    UNIValue(2,0) = "SELECT BIZ_AREA_NM FROM B_BIZ_AREA WHERE BIZ_AREA_CD =  " & FilterVar(UCase(Request("txtBizCd")), "''", "S") & " "             
                   
    
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
		BP_NM = rs1(0)
%>
		<Script Language=vbScript>
		With parent
			.Frm1.txtBpNm.Value  = "<%=ConvSPChars(BP_NM)%>"   
		End With
		</Script>
<%			
	ELSE
		if Trim(Request("txtBpCd")) <> "" Then
			strMsgCd = "970000"
%>
		<Script Language=vbScript>
		With parent
			.Frm1.txtBpNm.Value  = ""   
		End With
		</Script>
<%	
		Else 
%>
		<Script Language=vbScript>
		With parent
			.Frm1.txtBpNm.Value  = ""   
		End With
		</Script>
<%			
		End if
	End if
    rs1.Close
    Set rs1 = Nothing 
    
    'rs2�� ���� ��� 
    IF NOT (rs2.EOF or rs2.BOF) then
	    BIZ_AREA_NM = rs2(0)
%>
		<Script Language=vbScript>
		With parent
			.Frm1.txtBizNm.Value = "<%=ConvSPChars(BIZ_AREA_NM)%>"   
		End With
		</Script>
<%			    
	ELSE
		if Trim(Request("txtBizCd")) <> "" Then
			strMsgCd1 = "970000"
%>
		<Script Language=vbScript>
		With parent
			.Frm1.txtBizNm.Value = ""   
		End With
		</Script>
<%	
		Else
%>
		<Script Language=vbScript>
		With parent
			.Frm1.txtBizNm.Value = ""   
		End With
		</Script>
<%		
		End if
    END IF
    rs2.Close
    Set rs2 = Nothing
	
	If  "" & Trim(strMsgCd) <> "" Then
		Call DisplayMsgBox("970000", vbOKOnly, strMsg, "", I_MKSCRIPT)
        Response.End													'��: �����Ͻ� ���� ó���� ������ 
    End If
    
    If  "" & Trim(strMsgCd1) <> "" Then
		Call DisplayMsgBox("970000", vbOKOnly, strMsg1, "", I_MKSCRIPT)
        Response.End													'��: �����Ͻ� ���� ó���� ������ 
    End If
	
    If rs0.EOF And rs0.BOF Then
		Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
		rs0.Close:		Set rs0 = Nothing
		Response.End		
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
    AAmt = Request("txtRcptAmt")
    strFrApDt     = UCase(Trim(UNIConvDate(Request("txtApDt"))))
    strToApDt     = UCase(Trim(UNIConvDate(Request("txtToApDt"))))
    strDocCur	   = UCase(Trim(Request("txtDocCur")))
    strPayBpCd    = UCase(Trim(Request("txtBpCd")))
    strBizCd	   = UCase(Trim(Request("txtBizCd")))
	strDealBpCd   = UCase(Trim(Request("txtDealBpCd")))
	strApNo	   = UCase(Trim(Request("txtApNo")))
    strAllcDt		= UNIConvDate(Trim(Request("txtAllcDt")))
     
	' ���Ѱ��� �߰� 
	lgAuthBizAreaCd		= Trim(Request("lgAuthBizAreaCd"))
	lgInternalCd		= Trim(Request("lgInternalCd"))
	lgSubInternalCd		= Trim(Request("lgSubInternalCd"))
	lgAuthUsrID		= Trim(Request("lgAuthUsrID"))     

    If strFrApDt <> "" Then	:		strCond = strCond & " and A.AP_DT >=  " & FilterVar(strFrApDt , "''", "S") & ""
          
    If strToApDt <> "" Then	:		strCond = strCond & " and A.AP_DT <=  " & FilterVar(strToApDt , "''", "S") & ""
          
    If strDocCur <> "" Then	:		strCond = strCond & " and A.doc_cur =  " & FilterVar(strDocCur , "''", "S") & ""
          
    If strPayBpCd <> "" Then	:		strCond = strCond & " and A.pay_bp_cd =  " & FilterVar(strPayBpCd , "''", "S") & ""
          
    If strBizCd <> "" Then		:		strCond = strCond & " and A.biz_area_cd =  " & FilterVar(strBizCd , "''", "S") & ""
          
	If "" & strDealBpCd <> "" Then			:		strCond = strCond & " AND A.deal_bp_cd =  " & FilterVar(strDealBpCd , "''", "S") & "" 
		
	If "" & strApNo <> "" Then				:		strCond = strCond & " AND A.ap_no =  " & FilterVar(strApNo , "''", "S") & "" 
	      
    strCond = strCond & " AND A.bal_amt <> 0 and a.gl_no <> '' " 
    strCond = strCond & " AND A.ap_dt <=  " & FilterVar(strAllcDt , "''", "S") & ""

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
    '--------------- ������ coding part(�������,End)------------------------------------------------------

End Sub


%>
<Script Language=vbscript>
    If "<%=lgDataExist%>" = "Yes" Then

       'Set condition data to hidden area
       If "<%=lgPageNo%>" = "1" Then   ' "1" means that this query is first and next data exists
          parent.Frm1.htxtBizCd.Value		= Parent.Frm1.txtBizCd.Value
          Parent.Frm1.htxtBpCd.Value        = Parent.Frm1.txtBpCd.Value
          Parent.Frm1.htxtApDt.Value		= Parent.Frm1.txtApDt.Text
          Parent.Frm1.htxtToApDt.Value		= Parent.Frm1.txtToApDt.Text
          Parent.Frm1.htxtDocCur.Value		= Parent.Frm1.txtDocCur.Value
		  Parent.Frm1.htxtDealBpCd.Value	= Parent.Frm1.txtDealBpCd.value
		  Parent.Frm1.htxtApNo.Value		= Parent.Frm1.txtApNo.value          
       End If
       
       'Show multi spreadsheet data from this line
       
       Parent.ggoSpread.Source  = Parent.frm1.vspdData
       Parent.frm1.vspdData.Redraw = False
       Parent.ggoSpread.SSShowData "<%=lgstrData%>", "F"                    '�� : Display data
       Call Parent.ReFormatSpreadCellByCellByCurrency(Parent.Frm1.vspdData,<%=iPrevEndRow+1%>,<%=iEndRow%>,Parent.GetKeyPos("A",12),Parent.GetKeyPos("A",4),"A", "I" ,"X","X")
       Call Parent.ReFormatSpreadCellByCellByCurrency(Parent.Frm1.vspdData,<%=iPrevEndRow+1%>,<%=iEndRow%>,Parent.GetKeyPos("A",12),Parent.GetKeyPos("A",5),"A", "I" ,"X","X")
       Call Parent.ReFormatSpreadCellByCellByCurrency(Parent.Frm1.vspdData,<%=iPrevEndRow+1%>,<%=iEndRow%>,Parent.GetKeyPos("A",12),Parent.GetKeyPos("A",6),"A", "I" ,"X","X")
       Call Parent.ReFormatSpreadCellByCellByCurrency(Parent.Frm1.vspdData,<%=iPrevEndRow+1%>,<%=iEndRow%>,Parent.GetKeyPos("A",12),Parent.GetKeyPos("A",7),"A", "I" ,"X","X")
       Parent.frm1.vspdData.Redraw = True
       Parent.lgPageNo      =  "<%=lgPageNo%>"               '�� : Next next data tag
       Parent.DbQueryOk
    End If   
</Script>
	
