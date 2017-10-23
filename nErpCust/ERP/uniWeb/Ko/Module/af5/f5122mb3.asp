<%
'**********************************************************************************************
'*  1. Module��          : ȸ�� 
'*  2. Function��        : �������� 
'*  3. Program ID        : f5122mb3
'*  4. Program �̸�      : ���������̵�ó�� 
'*  5. Program ����      : ���������̵�ó�� 
'*  6. Comproxy ����Ʈ   : 
'*  7. ���� �ۼ������   : 2004/04/07
'*  8. ���� ���������   : 
'*  9. ���� �ۼ���       : ������ 
'* 10. ���� �ۼ���       : 
'* 11. ��ü comment      :
'* 12. ���� Coding Guide : this mark(��) means that "Do not change"
'*                         this mark(��) Means that "may  change"
'*                         this mark(��) Means that "must change"
'* 13. History           :
'*                         -2000/10/16 : ..........
'**********************************************************************************************
'�� : �׻� ���� ���̵� ������ �������� �²���(<)% �� %�첩��(>)�� New Line�� ��ġ�Ͽ� 
'	  ���� ���̵� ������ Ŭ���̾�Ʈ ���̵� ������ ��ġ�� ������ �� �ֵ��� �Ѵ�.
'�� : �Ʒ� HTML ������ ����Ǿ�� �ȵȴ�. 
%>

<!-- #Include file="../../inc/IncServer.asp"  -->

<%					

'�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 

On Error Resume Next														' ��: 
Err.Clear 

Dim lgADF																	'�� : ActiveX Data Factory ���� �������� 
Dim lgstrRetMsg																'�� : Record Set Return Message �������� 
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0								'�� : DBAgent Parameter ���� 


Call HideStatusWnd															'��: ��� �۾� �Ϸ��� �۾������� ǥ��â�� Hide

'Dim LngMaxRow																' ���� �׸����� �ִ�Row
Dim LngRow

Dim arrVal, arrTemp															'��: Spread Sheet �� ���� ���� Array ���� 
Dim strStatus																'��: Sheet �� ���� Row�� ���� (Create/Update/Delete)
Dim	lGrpCnt																	'��: Group Count

Dim strMode																	'��: ���� MyBiz.asp �� ������¸� ��Ÿ�� 

Dim StrNextKeyNoteNo														' NoteNO ���� �� 
Dim StrNextKeyGlNo															' GLNO ���� �� 
Dim lgStrPrevKeyNoteNo														' Note NO ���� �� 
Dim lgStrPrevKeyGlNo

Dim strNoteNo,strFrBizCd 
Dim strWhere0
Dim strMsgCd, strMsg1, strMsg2

Dim lgAuthBizAreaCd, lgAuthBizAreaNm			' ����� 
Dim lgInternalCd, lgDeptCd, lgDeptNm			' ���κμ� 
Dim lgSubInternalCd, lgSubDeptCd, lgSubDeptNm	' ���κμ�(��������)
Dim lgAuthUsrID, lgAuthUsrNm					' ���� 

Const GroupCount = 30

	strMode = Request("txtMode")											'�� : ���� ���¸� ���� 
    lgPageNo       = UNICInt(Trim(Request("lgPageNo")),0)					'��: "0"(First),"1"(Second),"2"(Third),"3"(...)
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
    Redim UNISqlId(0)														'��: SQL ID ������ ���� ����Ȯ�� 
		    '--------------- ������ coding part(�������,Start)----------------------------------------------------
                   
    UNISqlId(0) = "F5122MA102"
    
    Redim UNIValue(0,0)
	
	UNIValue(0,0) = strWhere0
		
    UNILock = DISCONNREAD :	UNIFlag = "1"									'��: set ADO read mode
 
End Sub

'----------------------------------------------------------------------------------------------------------
' Function QueryData()
'----------------------------------------------------------------------------------------------------------
Sub QueryData()
    Dim iStr
		    
    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1)
    
    iStr = Split(lgstrRetMsg,gColSep)
	  
    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
        Response.End
    End If
	
    If rs0.EOF And rs0.BOF Then
		Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
		rs0.Close:		Set rs0 = Nothing
		Set lgADF = Nothing
		Response.End														'��: �����Ͻ� ���� ó���� ������ 
	Else
		Call  MakeSpreadSheetData()
    End If				
	
    Call ReleaseObj()
End Sub
'----------------------------------------------------------------------------------------------------------
' Set default value or preset value
'----------------------------------------------------------------------------------------------------------
Sub TrimData()  

	strFromDt	 = UNIConvDate(Request("txtFrGlDt"))						'����ȸ���� 
	strToDt		 = UNIConvDate(Request("txtToGlDt"))						'����ȸ���� 
	
	lgAuthBizAreaCd		= Trim(Request("lgAuthBizAreaCd"))
	lgInternalCd		= Trim(Request("lgInternalCd"))
	lgSubInternalCd		= Trim(Request("lgSubInternalCd"))
	lgAuthUsrID		= Trim(Request("lgAuthUsrID"))	
			    		
	'''''''''''''''''''''KHJ -s  '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	strWhere0 = ""		
	strWhere0 = strWhere0 & " JOIN f_note b ON b.note_no = a.note_no "
    strWhere0 = strWhere0 & " JOIN b_acct_dept c ON c.org_change_id = a.org_change_id and c.dept_cd = a.dept_cd " 
	strWhere0 = strWhere0 & " JOIN b_acct_dept d ON d.org_change_id = b.org_change_id and d.dept_cd = b.dept_cd "
	strWhere0 = strWhere0 & " JOIN b_biz_partner e ON e.bp_cd = b.bp_cd "
	strWhere0 = strWhere0 & " LEFT OUTER JOIN a_temp_gl f ON f.temp_gl_no = a.temp_gl_no "
	strWhere0 = strWhere0 & " LEFT OUTER JOIN a_gl g ON g.gl_no = a.gl_no "
	strWhere0 = strWhere0 & " JOIN (select note_no,max(seq) seq from f_note_item where note_sts = " & FilterVar("MF","''","S") & " group by note_no) h on a.note_no=h.note_no and a.seq=h.seq "
	strWhere0 = strWhere0 & " JOIN (select note_no,gl_no,temp_gl_no,seq from f_note_item where note_sts = " & FilterVar("MV","''","S") & " ) i on a.note_no=i.note_no "
	strWhere0 = strWhere0 & " join (select note_no,max(seq) seq from f_note_item where note_sts = " & FilterVar("MV","''","S") & " group by note_no ) j on i.note_no=j.note_no and i.seq=j.seq "
	strWhere0 = strWhere0 & " where a.sts_dt between  " & FilterVar(strFromDt, "''", "S") & " and  " & FilterVar(strToDt, "''", "S") & "" 
	strWhere0 = strWhere0 & "  and  a.note_sts = " & FilterVar("MF", "''", "S") & "  " 
	strWhere0 = strWhere0 & "  and  b.note_sts = " & FilterVar("MV", "''", "S") & "  " 
	strWhere0 = strWhere0 & "  and  a.note_fg  = " & FilterVar("D1", "''", "S") & "  "
	
	' ���Ѱ��� �߰� 
	If lgAuthBizAreaCd <> "" Then
		lgBizAreaAuthSQL		= " AND b.BIZ_AREA_CD = " & FilterVar(lgAuthBizAreaCd, "''", "S")
	End If
	
	If lgInternalCd <> "" Then
		lgInternalCdAuthSQL		= " AND b.INTERNAL_CD = " & FilterVar(lgInternalCd, "''", "S")
	End If
	
	If lgSubInternalCd <> "" Then
		lgSubInternalCdAuthSQL	= " AND b.INTERNAL_CD LIKE " & FilterVar(lgSubInternalCd & "%", "''", "S")
	End If
	
	If lgAuthUsrID <> "" Then
		lgAuthUsrIDAuthSQL		= " AND b.UPDT_USER_ID = " & FilterVar(lgAuthUsrID, "''", "S")
	End If
	
	strWhere0 = strWhere0 & lgBizAreaAuthSQL & lgInternalCdAuthSQL & lgSubInternalCdAuthSQL & lgAuthUsrIDAuthSQL	
	
	If lgStrPrevKeyNoteNo <> "" Then strWhere0 = strWhere0 & " and A.NOTE_NO >= " & Filtervar(lgStrPrevKeyNoteNo	, "''", "S")
    
   	strWhere0 = strWhere0 & " Order by 1, 3 "
	
	'''''''''''''''''''''KHJ -e '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

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
	Const C_SHEETMAXROWS_D = 100
	
	LngMaxRow = parent.frm1.vspdData2.MaxRows										'Save previous Maxrow                                         	
<%
	
	If rs0.recordcount > GroupCount Then
		intLoopCnt = GroupCount
	Else
		intLoopCnt = rs0.recordcount
	End If
			
	For LngRow = 1 To intLoopCnt
%>		
		strData = strData & Chr(11) & 0
		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("TO_DEPT_CD"))%>"
		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("TO_DEPT_NM"))%>"
		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("GL_NO"))%>" 
		strData = strData & Chr(11) & "<%=UNIDateClientFormat(rs0("GL_DT"))%>"
		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("TEMP_GL_NO"))%>" 
		strData = strData & Chr(11) & "<%=UNIDateClientFormat(rs0("TEMP_GL_DT"))%>"		
		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("NOTE_NO"))%>"
		strData = strData & Chr(11) & "<%=UNINumClientFormat(rs0("AMT"), ggAmtOfMoney.DecPoint, 0)%>"	
		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("FR_DEPT_CD"))%>"  '�̵��μ� 
		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("FR_DEPT_NM"))%>"  '�̵��μ��� 
		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("BP_CD"))%>" 
		strData = strData & Chr(11) & "<%=ConvSPChars(rs0("BP_NM"))%>" 
		strData = strData & Chr(11) & "<%=UNIDateClientFormat(rs0("ISSUE_DT"))%>"
		strData = strData & Chr(11) & "<%=UNIDateClientFormat(rs0("DUE_DT"))%>"		
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
		.ggoSpread.Source = .frm1.vspdData2
		.ggoSpread.SSShowData strData

		If .frm1.vspdData2.MaxRows < C_SHEETMAXROWS_D And .lgStrPrevKeyNoteNo <> "" Then
			.DbQuery					
		Else
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
		


