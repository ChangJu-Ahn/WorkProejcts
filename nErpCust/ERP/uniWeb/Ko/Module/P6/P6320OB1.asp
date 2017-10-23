<%
'***********************************************************************************************************************
'*  1. Module Name          : XA
'*  2. Function Name        : ���ں�ASŬ��������(HB)
'*  3. Program ID           : XA005MB_KO321
'*  4. Program Name         : ���ں�ASŬ��������(HB)
'*  5. Program Desc         :
'*  6. Comproxy List        : +S31111MaintSoHdrSvr
'*  7. Modified date(First) : 2005/05/03
'*  8. Modified date(Last)  : 2005/06/02
'*  9. Modifier (First)     : Yoo Myung Sik
'* 10. Modifier (Last)      : Joo Young Hoon
'* 11. Comment              :
'************************************************************************************************************************

Response.Expires = -1								'�� : ASP�� ĳ������ �ʵ��� �Ѵ�.
Response.Buffer = True								'�� : ASP�� ���ۿ� ������� �ʰ� �ٷ� Client�� ��������.

'�� : �׻� ���� ���̵� ������ �������� �²���(<)% �� %�첩��(>)�� New Line�� ��ġ�Ͽ� 
'	  ���� ���̵� ������ Ŭ���̾�Ʈ ���̵� ������ ��ġ�� ������ �� �ֵ��� �Ѵ�.
'�� : �Ʒ� HTML ������ ����Ǿ�� �ȵȴ�. 

%>
<!-- #Include file="../../inc/IncServer.asp" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->		


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
	Dim txtCoWorker
	Dim selProcessType
	Dim PreAs_Dt
	Dim NextAs_Dt,JFlag
	
	JFlag=false
	

    Err.Clear                                                                '��: Protect system from crashing

	lgMaxCount=100
	iSelCount = lgMaxCount + lgMaxCount *  lgStrPrevKeyIndex + 1
	 
	lsReqdlvyFromDt = trim(Request("txtReqdlvyFromDt"))
	lsReqdlvyToDt = trim(Request("txtReqdlvyToDT"))
	txtCastCd = trim(Request("txtCastCd"))
	txtCastNM = trim(Request("txtCastNM"))
	seltype = trim(Request("seltype"))
	txtCustArea = trim(Request("txtCustArea"))
	txtFormalNm = trim(Request("txtFormalNm"))
	
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
    lgStrSQL = " SELECT distinct CAST_CD AS AB ,CAST_NM AS BC ,a.ITEM_CD_1 AS CD,ITEM_NM AS de ,FOrMAL_NM as ef,minor_nm as fg ,MAKE_DT AS gh         "                    
    lgStrSQL = lgStrSQL & " ,MAKER AS hi ,PRS_UNIT AS ij ,A.SPEC AS jk ,MAT_Q AS kl  ,CUR_ACCNT AS lm ,CUSTODY_AREA AS mn ,CLOSE_DT AS op   "  
    lgStrSQL = lgStrSQL & "   From Y_CAST A left outer join B_ITEM C   "  
    lgStrSQL = lgStrSQL & "  on	A.ITEM_CD_1  =  C.ITEM_CD "  
    lgStrSQL = lgStrSQL & "  left outer join B_MINOR B "  
    lgStrSQL = lgStrSQL & " on     A.EMP_CD = B.MINOR_CD and b.major_cd='y6002' "  
          
	if Trim(lsReqdlvyFromDt) <> "" Then 
		lgStrSQL = lgStrSQL & " where a.make_dt >= '" & lsReqdlvyFromDt & "' "  
		JFlag=true
	end if

	if Trim(lsReqdlvyToDt) <> "" Then
		if JFlag then
			lgStrSQL = lgStrSQL & " and a.make_dt <= '" & lsReqdlvyToDt & "' " 		
		else
			lgStrSQL = lgStrSQL & " where a.make_dt <= '" & lsReqdlvyToDt & "' " 
			JFlag=true
		end if
	end if
		
	if Trim(txtCastCd) <> "" Then
		if JFlag then
			lgStrSQL = lgStrSQL & " and   a.Cast_Cd = '" & txtCastCd & "' " 
		else
			lgStrSQL = lgStrSQL & " where   a.Cast_Cd = '" & txtCastCd & "' " 
			JFlag=true
		end if
	end if
	
	if Trim(txtCastNM) <> "" Then
		if JFlag then
			lgStrSQL = lgStrSQL & " and  a.CAST_NM like '%" & txtCastNM & "%' " 		
		else
			lgStrSQL = lgStrSQL & " where  a.CAST_NM like '%" & txtCastNM & "%' " 
			JFlag=true
		end if
	end if  
	
	if Trim(seltype) <> "" Then
		if JFlag then
			lgStrSQL = lgStrSQL & " and  a.emp_cd='" & seltype & "' " 		
		else
			lgStrSQL = lgStrSQL & " where  a.emp_cd='" & seltype & "' " 
			JFlag=true
		end if
	end if  
	
	if Trim(txtCustArea) <> "" Then
		if JFlag then
			lgStrSQL = lgStrSQL & " and  a.CUSTODY_AREA like '%" & txtCustArea & "%' " 		
		else
			lgStrSQL = lgStrSQL & " where  a.CUSTODY_AREA like '%" & txtCustArea & "%' " 
			JFlag=true
		end if
	end if 	
	if Trim(txtFormalNm) <> "" Then
		if JFlag then
			lgStrSQL = lgStrSQL & " and  formal_nm like '%" & txtFormalNm & "%' " 		
		else
			lgStrSQL = lgStrSQL & " where  formal_nm like '%" & txtFormalNm & "%' " 			
		end if
	end if 
	
   
    lgStrSQL = lgStrSQL & " order by ab"
     

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
					lgstrData = lgstrData & Chr(11) & ConvSpChars(lgObjRs("de")) 
					lgstrData = lgstrData & Chr(11) & ConvSpChars(lgObjRs("ef")) 
					lgstrData = lgstrData & Chr(11) & ConvSpChars(lgObjRs("fg")) 
					lgstrData = lgstrData & Chr(11) & UNIDateClientFormat(lgObjRs("gh")) 
					lgstrData = lgstrData & Chr(11) & ConvSpChars(lgObjRs("hi")) 
					lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs("ij"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
					lgstrData = lgstrData & Chr(11) & ConvSpChars(lgObjRs("jk"))	
					lgstrData = lgstrData & Chr(11) & ConvSpChars(lgObjRs("kl"))					
					lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs("lm"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
					lgstrData = lgstrData & Chr(11) & ConvSpChars(lgObjRs("mn"))
					lgstrData = lgstrData & Chr(11) & UNIDateClientFormat(lgObjRs("op"))
					
				
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
	Response.End																				'��: Process End


End Select

%>
