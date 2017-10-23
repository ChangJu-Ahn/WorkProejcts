<%
'**********************************************************************************************
'*  1. Module��          : ȸ�� 
'*  2. Function��        : A_RECEIPT
'*  3. Program ID        : f5116ma
'*  4. Program �̸�      : ����ī���ϰ�ó��(tab2)
'*  5. Program ����      : ����ī���ϰ�ó��(tab2)
'*  6. Comproxy ����Ʈ   : f5116ma
'*  7. ���� �ۼ������   : 2002/10/15
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

On Error Resume Next			' ��: 

Dim lgADF                                                                  '�� : ActiveX Data Factory ���� �������� 
Dim lgstrRetMsg                                                            '�� : Record Set Return Message �������� 
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2			           '�� : DBAgent Parameter ���� 


Call HideStatusWnd															'��: ��� �۾� �Ϸ��� �۾������� ǥ��â�� Hide

Dim LngMaxRow			' ���� �׸����� �ִ�Row
Dim LngRow

Dim arrVal, arrTemp																'��: Spread Sheet �� ���� ���� Array ���� 
Dim strStatus																	'��: Sheet �� ���� Row�� ���� (Create/Update/Delete)
Dim	lGrpCnt																		'��: Group Count
Dim ColSep, RowSep 

Dim strMode						'��: ���� MyBiz.asp �� ������¸� ��Ÿ�� 

Dim StrNextKeyNoteNo			' NoteNO ���� �� 
Dim StrNextKeyGlNo				' GLNO ���� �� 
Dim lgStrPrevKeyNoteNo			' Note NO ���� �� 
Dim lgStrPrevKeyGlNo
Dim lgStrPrevKeyTempGlNo
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

'Call GetGlobalVar

    lgPageNo       = UNICInt(Trim(Request("lgPageNo")),0)                  '��: "0"(First),"1"(Second),"2"(Third),"3"(...)

	lgStrPrevKeyNoteNo = "" & UCase(Trim(Request("lgStrPrevKeyNoteNo")))
	lgStrPrevKeyGlNo = "" & UCase(Trim(Request("lgStrPrevKeyGlNo")))

	strBizAreaCd	= Trim(UCase(Request("txtBizAreaCd")))            '�����From
	strBizAreaCd1	= Trim(UCase(Request("txtBizAreaCd1")))            '�����To
	
Call FixUNISQLDATA()

Call QueryData()


Sub FixUNISQLData()
    Dim intI
    Redim UNISqlId(2)                                                     '��: SQL ID ������ ���� ����Ȯ�� 
		    '--------------- ������ coding part(�������,Start)----------------------------------------------------

	lgAuthBizAreaCd		= Trim(Request("lgAuthBizAreaCd"))
	lgInternalCd		= Trim(Request("lgInternalCd"))
	lgSubInternalCd		= Trim(Request("lgSubInternalCd"))
	lgAuthUsrID			= Trim(Request("lgAuthUsrID"))

    UNISqlId(0) = "F5116MA102"
	UNISqlId(1) = "A_GETBIZ"
    UNISqlId(2) = "A_GETBIZ"
    
    Redim UNIValue(2,9)

	UNIValue(0,0) = FilterVar(UNIConvDate(Request("txtStsDtStart")), "''", "S")   
	UNIValue(0,1) = FilterVar(UNIConvDate(Request("txtStsDtEnd")), "''", "S")    
	
	If UCase(Trim(Request("txtBpCd"))) = "" Then
		UNIValue(0,2) = Filtervar("%"	, "", "S")
	Else
		UNIValue(0,2) = Filtervar(UCase(Request("txtBpCd"))	, "", "S")
	End If
	
	If UCase(Trim(Request("txtCardCoCd"))) = "" Then
		UNIValue(0,3) = ""
	Else
		UNIValue(0,3) = "AND A.CARD_CO_CD LIKE " & " " & FilterVar(UCase(Trim(Request("txtCardCoCd")))	, "''", "S") & " "
	End If	
	
	if strBizAreaCd <> "" then
		UNIValue(0,3) = UNIValue(0,3) & " AND A.BIZ_AREA_CD >= "	& FilterVar(strBizAreaCd , "''", "S") 
	else
		UNIValue(0,3) = UNIValue(0,3) & " AND A.BIZ_AREA_CD >= " & FilterVar("0", "''", "S") & " "
	end if
	
	if strBizAreaCd1 <> "" then
		UNIValue(0,3) = UNIValue(0,3) & " AND A.BIZ_AREA_CD <= "	& FilterVar(strBizAreaCd1 , "''", "S") 
	else
		UNIValue(0,3) = UNIValue(0,3) & " AND A.BIZ_AREA_CD <= " & FilterVar("ZZZZZZZZZZ", "''", "S") & " "
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
	
	UNIValue(0,3) = UNIValue(0,3) & lgBizAreaAuthSQL & lgInternalCdAuthSQL & lgSubInternalCdAuthSQL & lgAuthUsrIDAuthSQL	

	If lgStrPrevKeyNoteNo = "" Then
		UNIValue(0,4) = Filtervar("_"	, "", "S")
	Else 
		UNIValue(0,4) = Filtervar(lgStrPrevKeyNoteNo	, "''", "S")
	End If
	
	If lgStrPrevKeyGlNo = "" Then
		UNIValue(0,5) = Filtervar("_"	, "", "S")
		UNIValue(0,6) = Filtervar("_"	, "", "S")
	Else
		UNIValue(0,5) = Filtervar(lgStrPrevKeyGlNo	, "", "S")
		UNIValue(0,6) = Filtervar(lgStrPrevKeyGlNo	, "", "S")
	End If

	If lgStrPrevKeyNoteNo = "" Then
		UNIValue(0,7) = Filtervar("_"	, "", "S")
	Else 
		UNIValue(0,7) = Filtervar(lgStrPrevKeyNoteNo	, "''", "S")
	End If
	
	If lgStrPrevKeyTempGlNo = "" Then
		UNIValue(0,8) = Filtervar("_"	, "", "S")
		UNIValue(0,9) = Filtervar("_"	, "", "S")
	Else
		UNIValue(0,8) = Filtervar(lgStrPrevKeyTempGlNo	, "", "S")
		UNIValue(0,9) = Filtervar(lgStrPrevKeyTempGlNo	, "", "S")
	End If	
	
	UNIValue(1,0)  = FilterVar(strBizAreaCd, "''", "S")
	UNIValue(2,0)  = FilterVar(strBizAreaCd1, "''", "S")
		
    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '��: set ADO read mode
		 
End Sub

Sub QueryData()
    Dim iStr
		    
    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2)
		    
    iStr = Split(lgstrRetMsg,gColSep)
		  
    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
        Response.End
    End If
	
	
	'rs1
    If Not( rs1.EOF OR rs1.BOF) Then
   		strBizAreaCd = Trim(rs1(0))
		strBizAreaNm = Trim(rs1(1))
	Else
		strBizAreaCd = ""
		strBizAreaNm = ""
		
    End IF
%>
		<Script Language=vbScript>
		With parent.frm1
			.txtfromBizAreaCd1.value = "<%=strBizAreaCd%>"
			.txtfromBizAreaNm1.value = "<%=strBizAreaNm%>"
		End With
		</Script>
<%    
    rs1.Close
    Set rs1 = Nothing
    
    ' rs2
    If Not( rs2.EOF OR rs2.BOF) Then
   		strBizAreaCd1 = Trim(rs2(0))
		strBizAreaNm1 = Trim(rs2(1))
	Else
		strBizAreaCd1 = ""
		strBizAreaNm1 = ""
		
    End IF
%>
		<Script Language=vbScript>
		With parent.frm1
			.txttoBizAreaCd1.value = "<%=strBizAreaCd1%>"
			.txttoBizAreaNm1.value = "<%=strBizAreaNm1%>"
		End With
		</Script>
<%        
    rs2.Close
    Set rs2 = Nothing
    
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
		
Sub MakeSpreadSheetData()
	Dim intLoopCnt
%>
<Script Language=vbscript>
Option Explicit

	Dim LngMaxRow       
	Dim strData
	
	Const C_SHEETMAXROWS_D = 30
	
	LngMaxRow = parent.frm1.vspdData2.MaxRows										'Save previous Maxrow                                                

<%
	If rs0.recordcount > GroupCount Then
		intLoopCnt = GroupCount
	Else
		intLoopCnt = rs0.recordcount
	End If
			
	For LngRow = 1 To intLoopCnt
'	For LngRow = 1 To rs0.recordcount	
%>
		strData = strData & Chr(11) & 0
		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("NOTE_NO"))%>"
		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("TEMP_GL_NO"))%>"
		strData = strData & Chr(11) & "<%=UNIDateClientFormat(rs0("TEMP_GL_DT"))%>"
		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("GL_NO"))%>"
		strData = strData & Chr(11) & "<%=UNIDateClientFormat(rs0("GL_DT"))%>"
		strData = strData & Chr(11) & "<%=UNINumClientFormat(rs0("NOTE_AMT"), ggAmtOfMoney.DecPoint, 0)%>"
		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("BP_CD"))%>"
		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("BP_NM"))%>"
		If "<%=ConvSPChars(rs0("CARD_CO_CD"))%>" <> "" Then
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("CARD_CO_CD"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("CARD_CO_NM"))%>"		
		Else 
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("BANK_CD"))%>"
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("BANK_NM"))%>"		
		End If 
		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("DEPT_CD"))%>"
		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("DEPT_NM"))%>"
		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("NOTE_ITEM_DESC"))%>"
		strData = strData & Chr(11) & LngMaxRow + <%=LngRow%>
		strData = strData & Chr(11) & Chr(12)

<%      
		rs0.MoveNext
    Next
		    
    If Not rs0.EOF Then
%>    
		parent.lgStrPrevKeyNoteNo = "<%=ConvSPChars(rs0("NOTE_NO"))%>"
		parent.lgStrPrevKeyGlNo = "<%=ConvSPChars(rs0("GL_NO"))%>"		
<%	Else	%>
		parent.lgStrPrevKeyNoteNo = ""
		parent.lgStrPrevKeyGlNo = ""
<%	End If	%>
		
	With parent
		.ggoSpread.Source = .frm1.vspdData2
		.ggoSpread.SSShowData strData
				
		If .frm1.vspdData2.MaxRows < C_SHEETMAXROWS_D And .lgStrPrevKeyNoteNo <> "" And .lgStrPrevKeyGlNo <> "" Then
			.DbQuery					
		Else
			.frm1.hProcFg.value		= "<%=ConvSPChars(Request("cboProcFg"))%>"
			.frm1.hStsDtStart.value	= "<%=Request("txtStsDtStart")%>"
			.frm1.hStsDtEnd.value	= "<%=Request("txtStsDtEnd")%>"				
			.frm1.hBpCd2.value		= "<%=ConvSPChars(Request("txtBpCd2"))%>"
			.frm1.hCardCoCd2.value	= "<%=ConvSPChars(Request("txtCardCoCd2"))%>"
			.frm1.hfromtxtBizAreaCd1.value		= "<%=strBizAreaCd%>"							
			.frm1.htotxtBizAreaCd1.value		= "<%=strBizAreaCd1%>"
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
