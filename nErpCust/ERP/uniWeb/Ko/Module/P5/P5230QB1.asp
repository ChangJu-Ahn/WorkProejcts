<%@LANGUAGE = VBScript%>
<% Option Explicit%>
<% 
'======================================================================================================
'*  1. Module Name          : p5
'*  2. Function Name        : �������������ȸ(HB)
'*  3. Program ID           : XP010OA_KO321
'*  4. Program Name         : �������������ȸ(HB)
'*  5. Program Desc         : 
'*  6. Comproxy List        :
'*  7. Modified date(First) : 2001/04/18
'*  8. Modified date(Last)  : 2005/07/20
'*  9. Modifier (First)     : Joo Young Hoon
'* 10. Modifier (Last)      : Joo Young Hoon
'* 11. Comment              :
'************************************************************************************************************************

'�� : �׻� ���� ���̵� ������ �������� �²���(<)% �� %�첩��(>)�� New Line�� ��ġ�Ͽ� 
'	  ���� ���̵� ������ Ŭ���̾�Ʈ ���̵� ������ ��ġ�� ������ �� �ֵ��� �Ѵ�.
'�� : �Ʒ� HTML ������ ����Ǿ�� �ȵȴ�.
%> 

<!-- #Include file="../../inc/IncServer.asp" -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->	
<!-- #Include file="../../inc/adovbs.inc" -->	
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->

	
<%													'�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 

On Error Resume Next														'��: 

    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '��: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '��: Read Operation Mode (CRUD)
    lgKeyStream       = Split(Request("txtKeyStream"),gColSep)

    lgLngMaxRow       = Request("txtMaxRows")                                        '��: Read Operation Mode (CRUD)
    lgMaxCount        = CInt(Request("lgMaxCount"))                                  '��: Fetch count at a time for VspdData
    lgStrPrevKeyIndex = UNICInt(Trim(Request("lgStrPrevKeyIndex")),0)                '��: "0"(First),"1"(Second),"2"(Third),"3"(...)



Call HideStatusWnd

Dim strMode		
Dim StrNextKey							' ���� �� 
Dim LngMaxRow							' ���� �׸����� �ִ�Row
Dim LngRow
Dim intGroupCount															'��: ���� MyBiz.asp �� ������¸� ��Ÿ�� 
Dim StrNext1
Dim StrNext2
Dim StrSeq
Dim lsSEQ

strMode = Request("txtMode")												'�� : ���� ���¸� ���� 

Call SubOpenDB(lgObjConn)                                                        '��: Make a DB Connection

Select Case strMode

Case CStr(UID_M0001)														'��: ���� ��ȸ/Prev/Next ��û�� ���� 
    
    Dim lsBP_CD
    Dim lstxtBP_CD    
    Dim SumQty
	Dim TNum
	Dim lsReqdlvyFromDt
	Dim lsReqdlvyToDt
	Dim seltype, txtCastCd


    Err.Clear                                                                '��: Protect system from crashing

	lgMaxCount=100
	iSelCount = lgMaxCount + lgMaxCount *  lgStrPrevKeyIndex + 1
	 
	lsReqdlvyFromDt = UNIConvDate(Request("txtReqdlvyFromDt"))
	lsReqdlvyToDt = UNIConvDate(Request("txtReqdlvyToDT"))	
	seltype = trim(Request("seltype"))
	txtCastCd = trim(Request("txtCastCd"))
	
	
	'����� ��ȸ 
	If Trim(txtCastCd) <> "" Then
		lgStrSQL = "SELECT FACILITY_NM FROM Y_FACILITY (NOLOCK) WHERE FACILITY_CD = " & Filtervar(txtCastCd, "''", "S")      
		If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
			Response.Write " <Script Language=""VBScript""> " & vbCrLf
			Response.Write " Parent.frm1.txtCastNm.value =  """"" & vbCrLf
			Response.Write " </Script> " & vbCrLf	
		Else
			Response.Write " <Script Language=""VBScript""> " & vbCrLf
			Response.Write " Parent.frm1.txtCastNm.value =  """ & ConvSPChars(lgObjRs("FACILITY_NM")) & """" & vbCrLf
			Response.Write " </Script> " & vbCrLf
		End if   
	Else
		Response.Write " <Script Language=""VBScript""> " & vbCrLf
		Response.Write " Parent.frm1.txtCastNm.value =  """"" & vbCrLf
		Response.Write " </Script> " & vbCrLf	   
	End If
	    
	'���� ��ȸ 
    lgStrSQL = " SELECT A.FAC_CAST_CD AS AB,B.FACILITY_NM AS BC, dbo.ufn_GetCodeName('Z410', FACILITY_ACCNT) AS CD, "                    
    lgStrSQL = lgStrSQL & " A.WORK_DT AS DE  ,dbo.ufn_GetCodeName('Z425', D.ZINSP_PART) AS EF, "  
    lgStrSQL = lgStrSQL & " A.INSP_TEXT AS FG,G.BP_NM AS GH, dbo.ufn_H_GetEmpName(F.INSP_EMP_CD) AS HI  "  
    lgStrSQL = lgStrSQL & " FROM Y_FAC_CAST_PLAN AS A,Y_FACILITY AS B,Y_FAC_CAST_CHECK AS D, Y_FAC_CAST_REPAIR AS F,B_BIZ_PARTNER AS G  "  
    lgStrSQL = lgStrSQL & " WHERE A.FAC_CAST_CD = B.FACILITY_CD "  
    lgStrSQL = lgStrSQL & " AND A.GUBUN_CD='10' "  
    lgStrSQL = lgStrSQL & " AND A.PLAN_GUBUN='2' "
    lgStrSQL = lgStrSQL & " AND A.GUBUN_CD=D.GUBUN_CD  AND A.FAC_CAST_CD=D.FAC_CAST_CD  "    
    lgStrSQL = lgStrSQL & " AND A.WORK_DT=D.WORK_DT  AND A.PLAN_GUBUN=D.PLAN_GUBUN   "  
    lgStrSQL = lgStrSQL & " AND A.FAC_CAST_CD=F.FAC_CAST_CD  AND A.WORK_DT=F.WORK_DT "  
    lgStrSQL = lgStrSQL & " AND A.PLAN_GUBUN=F.PLAN_GUBUN AND F.CUST_CD *= G.BP_CD   "   
       
	if Trim(lsReqdlvyFromDt) <> "" Then 
		lgStrSQL = lgStrSQL & " and a.work_dt >= '" & lsReqdlvyFromDt & "' "  
	end if

	if Trim(lsReqdlvyToDt) <> "" Then 
		lgStrSQL = lgStrSQL & " and a.work_dt <=  '" & lsReqdlvyToDt & "' " 		
	end if
	
	if Trim(seltype) <> "" Then 
		lgStrSQL = lgStrSQL & " and facility_accnt = '" & seltype & "' "  
	end if

	if Trim(txtCastCd) <> "" Then 
		lgStrSQL = lgStrSQL & " and a.fac_cast_cd = '" & txtCastCd & "' " 		
	end if
		
    lgStrSQL = lgStrSQL & " order by ab"
    
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
		
        lgStrPrevKeyIndex = ""
        Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '�� : No data is found. 
    Else 
		
        Do While Not lgObjRs.EOF
			
					lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ab")) 
					lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("bc")) 
					lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("cd")) 
					lgstrData = lgstrData & Chr(11) & UNIDateClientFormat(lgObjRs("de"))
					lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ef")) 
					lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("fg")) 
					lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("gh"))
					lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("hi"))		
					lgstrData = lgstrData & Chr(11) & Chr(12)

		
	'------ Developer Coding part (End   ) ------------------------------------------------------------------

			
		    lgObjRs.MoveNext
             
        
        Loop 

    End if 

    Call SubCloseRs(lgObjRs)                                                              '��: Release RecordSSet

%>
    
<Script Language="VBScript">

             With Parent
                .ggoSpread.Source     = .frm1.vspdData
                .ggoSpread.SSShowData "<%=lgstrData%>"

                .DBQueryOk                  
	         End with
       
</Script>	

<%
	Response.End																				'��: Process End
%>
<Script Language=vbscript>
	With parent																			
		.fncQuery
	End With
</Script>
<%					
	Call SubCloseDB(lgObjConn)                                                       '��: Close DB Connection
	Response.End																				 '��: Process End

End Select


%>
