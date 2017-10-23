<%@LANGUAGE = VBScript%>
<% Option Explicit%>
<%
'***********************************************************************************************************************
'*  1. Module Name          : ������� 
'*  2. Function Name        : �����������(HB)
'*  3. Program ID           : XP011OA_KO321
'*  4. Program Name         : �����������(HB)
'*  5. Program Desc         : 
'*  6. Comproxy List        :
'*  7. Modified date(First) : 2001/04/18
'*  8. Modified date(Last)  : 2005/07/29
'*  9. Modifier (First)     : Joo Young Hoon
'* 10. Modifier (Last)      : Joo Young Hoon
'* 11. Comment      
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
    
	Dim TNum
	Dim lsReqdlvyFromDt
	Dim lsReqdlvyToDt
	Dim txtFacilityCd, txtFacilityNM
	Dim JFlag, seltype
	
	JFlag=false
	

    Err.Clear                                                                '��: Protect system from crashing

	lgMaxCount=100
	iSelCount = lgMaxCount + lgMaxCount *  lgStrPrevKeyIndex + 1
	 
	lsReqdlvyFromDt = trim(Request("txtReqdlvyFromDt"))
	lsReqdlvyToDt = trim(Request("txtReqdlvyToDT"))
	txtFacilityCd = trim(Request("txtFacilityCd"))
	txtFacilityNM = trim(Request("txtFacilityNM"))
	seltype = trim(Request("seltype"))

		
    lgStrSQL = " SELECT FACILITY_CD AS AB,FACILITY_NM AS BC, dbo.ufn_GetCodeName('Y6002', A.EMP_NO) AS CD, SET_DT AS DE,PROD_CO AS EF,PROD_AMT AS FG,PM_DT AS GH        "                    
    lgStrSQL = lgStrSQL & " FROM Y_FACILITY A "  
          
          
	if Trim(lsReqdlvyFromDt) <> "" Then 
		lgStrSQL = lgStrSQL & " where a.set_dt >= '" & lsReqdlvyFromDt & "' "  
		JFlag=true
	end if


	if Trim(lsReqdlvyToDt) <> "" Then
		if JFlag then
			lgStrSQL = lgStrSQL & " and a.set_dt <= '" & lsReqdlvyToDt & "' "	
		else
			lgStrSQL = lgStrSQL & " where a.set_dt <= '" & lsReqdlvyToDt & "' "
			JFlag=true
		end if
	end if
		
	if Trim(txtFacilityCd) <> "" Then
		if JFlag then
			lgStrSQL = lgStrSQL & " and   a.facility_Cd = '" & txtFacilityCd & "' " 
		else
			lgStrSQL = lgStrSQL & " where   a.facility_Cd = '" & txtFacilityCd & "' " 
			JFlag=true
		end if
	end if
	
	
	if Trim(txtFacilityNM) <> "" Then
		if JFlag then
			lgStrSQL = lgStrSQL & " and  a.facility_NM like '%" & txtFacilityNM & "%' " 		
		else
			lgStrSQL = lgStrSQL & " where  a.facility_NM like '%" & txtFacilityNM & "%' " 
			JFlag=true
		end if
	end if  
	
	
	if Trim(seltype) <> "" Then
		if JFlag then
			lgStrSQL = lgStrSQL & " and  a.emp_no='" & seltype & "' " 		
		else
			lgStrSQL = lgStrSQL & " where  a.emp_no='" & seltype & "' " 
			JFlag=true
		end if
	end if  
	   
    lgStrSQL = lgStrSQL & " order by ab "	

    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
		
        lgStrPrevKeyIndex = ""
        Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '�� : No data is found. 
        Call SetErrorStatus()
    else 
		
		TNum=0		
		
		
		
        Do While Not lgObjRs.EOF
						
			lgstrData = lgstrData & Chr(11) & ConvSpChars(lgObjRs("ab")) 
			lgstrData = lgstrData & Chr(11) & ConvSpChars(lgObjRs("bc"))
			lgstrData = lgstrData & Chr(11) & ConvSpChars(lgObjRs("cd")) 
			lgstrData = lgstrData & Chr(11) & UNIDateClientFormat(lgObjRs("de")) 
			lgstrData = lgstrData & Chr(11) & ConvSpChars(lgObjRs("ef")) 
			lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs("fg"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
			lgstrData = lgstrData & Chr(11) & UNIDateClientFormat(lgObjRs("gh"))	
			lgstrData = lgstrData & Chr(11) & Chr(12)					
		
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
			
		    lgObjRs.MoveNext
             
        Loop 

    End if 
 	
 	Call SubHandleError("MR",lgObjConn,lgObjRs,Err)
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
	Call SubCloseDB(lgObjConn)                                                       '��: Close DB Connection			
	Response.End																				 '��: Process End

End Select


%>
