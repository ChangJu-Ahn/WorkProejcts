<%
'**********************************************************************************************
'*  1. Module��          : ȸ�� 
'*  2. Function��        : A_RECEIPT
'*  3. Program ID        : f5116ma
'*  4. Program �̸�      : ����ī���ϰ�ó��(tab1)
'*  5. Program ����      : ����ī���ϰ�ó��(tab1)
'*  6. Comproxy ����Ʈ   : f5116ma
'*  7. ���� �ۼ������   : 2002/10/14
'*  8. ���� ���������   : 
'*  9. ���� �ۼ���       : ������ 
'* 10. ���� �ۼ���       : 
'* 11. ��ü comment      :
'* 12. ���� Coding Guide : this mark(��) means that "Do not change"
'*                         this mark(��) Means that "may  change"
'*                         this mark(��) Means that "must change"
'* 13. History           :
'**********************************************************************************************




'�� : �׻� ���� ���̵� ������ �������� �²���(<)% �� %�첩��(>)�� New Line�� ��ġ�Ͽ� 
'	  ���� ���̵� ������ Ŭ���̾�Ʈ ���̵� ������ ��ġ�� ������ �� �ֵ��� �Ѵ�.
'�� : �Ʒ� HTML ������ ����Ǿ�� �ȵȴ�. 
%>

<!-- #Include file="../../inc/IncServer.asp"  -->

<%					

'�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 

'On Error Resume Next			' ��: 

Dim lgADF                                                                  '�� : ActiveX Data Factory ���� �������� 
Dim lgstrRetMsg                                                            '�� : Record Set Return Message �������� 
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3, rs4					           '�� : DBAgent Parameter ���� 


Call HideStatusWnd															'��: ��� �۾� �Ϸ��� �۾������� ǥ��â�� Hide

'Dim LngMaxRow																' ���� �׸����� �ִ�Row
Dim LngRow

Dim arrVal, arrTemp																'��: Spread Sheet �� ���� ���� Array ���� 
Dim strStatus																	'��: Sheet �� ���� Row�� ���� (Create/Update/Delete)
Dim	lGrpCnt																		'��: Group Count

Dim strMode						'��: ���� MyBiz.asp �� ������¸� ��Ÿ�� 

Dim StrNextKeyNoteNo			' NoteNO ���� �� 
Dim StrNextKeyGlNo				' GLNO ���� �� 
Dim lgStrPrevKeyNoteNo			' Note NO ���� �� 
Dim lgStrPrevKeyGlNo

Dim strNoteFg, strNoteSts, strDueDtStart, strDueDtEnd, strBankCd 
Dim strWhere0
Dim strMsgCd, strMsg1, strMsg2
Dim strBizAreaCd
Dim strBizAreaNm
Dim strBizAreaCd1
Dim strBizAreaNm1

Dim lgAuthBizAreaCd, lgAuthBizAreaNm			' ����� 
Dim lgInternalCd, lgDeptCd, lgDeptNm			' ���κμ� 
Dim lgSubInternalCd, lgSubDeptCd, lgSubDeptNm	' ���κμ�(��������)
Dim lgAuthUsrID, lgAuthUsrNm					' ���� 

Const GroupCount = 30
	
	strMode = Request("txtMode")	'�� : ���� ���¸� ���� 

    lgPageNo       = UNICInt(Trim(Request("lgPageNo")),0)                  '��: "0"(First),"1"(Second),"2"(Third),"3"(...)

	lgStrPrevKeyNoteNo = "" & UCase(Trim(Request("lgStrPrevKeyNoteNo")))
	lgStrPrevKeyGlNo = "" & UCase(Trim(Request("lgStrPrevKeyGlNo")))
		
Call TrimData()
Call FixUNISQLDATA()
Call QueryData()

'----------------------------------------------------------------------------------------------------------
' Function FixUNISQLData()
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()
    Dim intI
    Redim UNISqlId(4)                                                     '��: SQL ID ������ ���� ����Ȯ�� 
		    '--------------- ������ coding part(�������,Start)----------------------------------------------------

    UNISqlId(0) = "F5116MA101"
    UNISqlId(1) = "ABPNM"
    UNISqlId(2) = "ACARDCONM"
    UNISqlId(3) = "A_GETBIZ"
    UNISqlId(4) = "A_GETBIZ"

    Redim UNIValue(4,1)
	
	UNIValue(0,0) = strWhere0
	UNIValue(1,0) = Filtervar(strBpCd, "''", "S")
	UNIValue(2,0) = Filtervar(strCardCoCd, "''", "S")
	UNIValue(3,0) = FilterVar(strBizAreaCd, "''", "S")
	UNIValue(4,0) = FilterVar(strBizAreaCd1, "''", "S")	
		
    UNILock = DISCONNREAD :	UNIFlag = "1"								'��: set ADO read mode
		 
End Sub

'----------------------------------------------------------------------------------------------------------
' Function QueryData()
'----------------------------------------------------------------------------------------------------------
Sub QueryData()
    Dim iStr
		    
    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3, rs4)
    
    iStr = Split(lgstrRetMsg,gColSep)
		  
    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
        Response.End
    End If
	
	If rs1.EOF And rs1.BOF Then
		If strMsgCd = "" And strBankCd <> "" Then			
			strMsgCd = "970000"
			strMsg1 = Request("txtBankCd_Alt")
			If strMsgCd <> "" Then
				Call DisplayMsgBox(strMsgCd, vbInformation, strMsg1, strMsg2, I_MKSCRIPT)
				Response.End 
			End If	
		End If
	Else
%>
		<Script Language=vbScript>
		With parent.frm1
			.txtBankCd.value = "<%=strBizAreaCd%>"
			.txtBankNm.value = "<%=ConvSPChars(Trim(rs1(1)))%>"
		End With
		</Script>
<%
	End If
	
	'rs3
    If Not( rs3.EOF OR rs3.BOF) Then
   		strBizAreaCd = Trim(rs3(0))
		strBizAreaNm = Trim(rs3(1))
	Else
		strBizAreaCd = ""
		strBizAreaNm = ""
		
    End IF
%>
		<Script Language=vbScript>
		With parent.frm1
			.txtfromBizAreaCd.value = "<%=strBizAreaCd%>"
			.txtfromBizAreaNm.value = "<%=strBizAreaNm%>"
		End With
		</Script>
<%
    
    rs3.Close
    Set rs3 = Nothing
    
    ' rs4
    If Not( rs4.EOF OR rs4.BOF) Then
   		strBizAreaCd1 = Trim(rs4(0))
		strBizAreaNm1 = Trim(rs4(1))
	Else
		strBizAreaCd1 = ""
		strBizAreaNm1 = ""
		
    End IF
%>
		<Script Language=vbScript>
		With parent.frm1
			.txttoBizAreaCd.value = "<%=strBizAreaCd1%>"
			.txttoBizAreaNm.value = "<%=strBizAreaNm1%>"
		End With
		</Script>
<%    
    rs4.Close
    Set rs4 = Nothing
	
    If rs0.EOF And rs0.BOF Then
		Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
		rs0.Close:		Set rs0 = Nothing
		Set lgADF = Nothing
		Response.End													'��: �����Ͻ� ���� ó���� ������ 
	Else
		Call  MakeSpreadSheetData()
    End If				
	
    Call ReleaseObj()
End Sub


'----------------------------------------------------------------------------------------------------------
' Set default value or preset value
'----------------------------------------------------------------------------------------------------------
Sub TrimData()  

	strNoteFg = "CP"																'ī�屸�� 
'	strNoteSts = Request("cboNoteSts")									'ī����� 
'	strDueDtStart = UNIConvDate(Request("txtDueDtStart"))			'���۸����� 
	strDueDtEnd = UNIConvDate(Request("txtDueDtEnd"))				'���Ḹ����	
	strFrNoteNo = Request("txtFrNoteNo")									'����ī���ȣ 
	strToNoteNo = Request("txtToNoteNo")								'����ī���ȣ 
	strBpCd = UCase(Request("txtBpCd"))									'�ŷ�ó�ڵ� 
	strCardCoCd = UCase(Request("txtCardCoCd"))						'ī����ڵ� 
	strBizAreaCd	= Trim(UCase(Request("txtBizAreaCd")))            '�����From
	strBizAreaCd1	= Trim(UCase(Request("txtBizAreaCd1")))            '�����To

	lgAuthBizAreaCd		= Trim(Request("lgAuthBizAreaCd"))
	lgInternalCd		= Trim(Request("lgInternalCd"))
	lgSubInternalCd		= Trim(Request("lgSubInternalCd"))
	lgAuthUsrID		= Trim(Request("lgAuthUsrID"))
		    		
	strWhere0 = ""
	strWhere0 = strWhere0 & " and A.NOTE_FG =  " & FilterVar(strNoteFg , "''", "S") & " "	
	strWhere0 = strWhere0 & " and  A.DUE_DT <=  " & FilterVar(strDueDtEnd , "''", "S") & " "
	
'	If strNoteSts <> "" Then 
'		strWhere0 = strWhere0 & " and A.NOTE_STS = '" & strNoteSts & "' "				'�������� 
'		strWhere0 = strWhere0 & " and D.MINOR_CD = '" & strNoteSts & "' "	
'	Else
		strWhere0 = strWhere0 & " and (A.NOTE_STS = " & FilterVar("OC", "''", "S") & "  OR A.NOTE_STS = " & FilterVar("DC", "''", "S") & " ) "		
'	End If
	
	If strFrNoteNo <> "" Then 
		strWhere0 = strWhere0 & " and A.NOTE_NO >= " & Filtervar(strFrNoteNo	, "''", "S")		
	End If
	
	If strToNoteNo <> "" Then 
		strWhere0 = strWhere0 & " and A.NOTE_NO <= " & Filtervar(strToNoteNo	, "''", "S")		
	End If
	
	If strBpCd <> "" Then 
		strWhere0 = strWhere0 & " and A.BP_CD = " & Filtervar(strBpCd	, "''", "S")
		strWhere0 = strWhere0 & " and C.BP_CD = " & Filtervar(strBpCd	, "''", "S")				
	End If
	
	If strCardCoCd <> "" Then 
		strWhere0 = strWhere0 & " and A.CARD_CO_CD = " & Filtervar(strCardCoCd	, "''", "S")
		strWhere0 = strWhere0 & " and E.CARD_CO_CD = " & Filtervar(strCardCoCd	, "''", "S")				
	End If
	
	if strBizAreaCd <> "" then
		strWhere0 = strWhere0 & " AND A.BIZ_AREA_CD >= " & FilterVar(strBizAreaCd , "''", "S") 
	else
		strWhere0 = strWhere0 & " AND A.BIZ_AREA_CD >= " & FilterVar("0", "''", "S") & " "
	end if
	
	if strBizAreaCd1 <> "" then
		strWhere0 = strWhere0 & " AND A.BIZ_AREA_CD <= " & FilterVar(strBizAreaCd1 , "''", "S") 
	else
		strWhere0 = strWhere0 & " AND A.BIZ_AREA_CD <= " & FilterVar("ZZZZZZZZZZ", "''", "S") & " "
	end if

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
	
	strWhere0 = strWhere0 & lgBizAreaAuthSQL & lgInternalCdAuthSQL & lgSubInternalCdAuthSQL & lgAuthUsrIDAuthSQL	

	If lgStrPrevKeyNoteNo <> "" Then strWhere0 = strWhere0 & " and A.NOTE_NO >= " & Filtervar(lgStrPrevKeyNoteNo	, "''", "S")	

End Sub

'----------------------------------------------------------------------------------------------------------
' Set MakeSpreadSheetData
'----------------------------------------------------------------------------------------------------------
Sub MakeSpreadSheetData()
	Dim intLoopCnt
%>
<Script Language=vbscript>
Option Explicit

	Dim LngMaxRow       
	Dim strData
	
	Const C_SHEETMAXROWS_D = 30
	
	LngMaxRow = parent.frm1.vspdData.MaxRows										'Save previous Maxrow                                         	
<%
	
	If rs0.recordcount > GroupCount Then
		intLoopCnt = GroupCount
	Else
		intLoopCnt = rs0.recordcount
	End If
			
	For LngRow = 1 To intLoopCnt
%>		
		strData = strData & Chr(11) & 0
		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("NOTE_NO"))%>"
		strData = strData & Chr(11) & "<%=UNINumClientFormat(rs0("NOTE_AMT"), ggAmtOfMoney.DecPoint, 0)%>"
		strData = strData & Chr(11) & "<%=UNIDateClientFormat(rs0("DUE_DT"))%>"
		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("MINOR_NM"))%>"				
		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("BP_CD"))%>" 
		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("BP_NM"))%>"
		<%if rs0("CARD_CO_CD") <> "" then%>
		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("CARD_CO_CD"))%>"
		<%else%>
		strData = strData & Chr(11) & ""
		<%end if%>
		<%if rs0("CARD_CO_NM") <> "" then%>
		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("CARD_CO_NM"))%>"
		<%else%>
		strData = strData & Chr(11) & ""
		<%end if%>
		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("DEPT_CD"))%>"
		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("DEPT_NM"))%>"
		strData = strData & Chr(11) & ""    
		strData = strData & Chr(11) & LngMaxRow + <%=LngRow%> 
		strData = strData & Chr(11) & Chr(12)			
<%      				
		rs0.MoveNext
    Next
		    
    If Not rs0.EOF Then
%>    
		parent.lgStrPrevKeyNoteNo = "<%=ConvSPChars(rs0("NOTE_NO"))%>"
		parent.lgStrPrevKeyGlNo = ""
<%	Else	%>
		parent.lgStrPrevKeyNoteNo = ""
		parent.lgStrPrevKeyGlNo = ""
<%	End If	%>
		
	With parent
		.ggoSpread.Source = .frm1.vspdData
		.ggoSpread.SSShowData strData
				
		If .frm1.vspdData.MaxRows < C_SHEETMAXROWS_D And .lgStrPrevKeyNoteNo <> "" Then
			.DbQuery					
		Else
			.frm1.hProcFg.value			= "<%=ConvSPChars(Request("cboProcFg"))%>"
			.frm1.hNoteFg.value			= "CP"
			.frm1.hDueDtEnd.value		= "<%=Request("txtDueDtEnd")%>"						
			.frm1.hBpCd1.value			= "<%=ConvSPChars(Request("txtBpCd"))%>"
			.frm1.hCardCoCd1.value		= "<%=ConvSPChars(Request("txtCardCoCd"))%>"
			.frm1.hfromtxtBizAreaCd.value		= "<%=strBizAreaCd%>"							
			.frm1.htotxtBizAreaCd.value			= "<%=strBizAreaCd1%>"
											
			.DbQueryOK
		End If
		
	End With
					
</script>
<%      
End Sub
		
Sub ReleaseObj()			
	Set rs0 = Nothing
	Set lgADF = Nothing                                                    '��: ActiveX Data Factory Object Nothing
End Sub			
		
%>		
