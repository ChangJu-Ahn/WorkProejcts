<%@LANGUAGE = VBScript%>
<%'======================================================================================================
'*  1. Module Name          : Production
'*  2. Function Name        : Multi Sample
'*  3. Program ID           : p1502mb9
'*  4. Program Name         : p1502mb9
'*  5. Program Desc         : �ڿ��׷���ȸ 
'*  6. Comproxy List        :
'*  7. Modified date(First) : 2001/11/27
'*  8. Modified date(Last)  : 2003/01/28
'*  9. Modifier (First)     : Jung Yu Kyung
'* 10. Modifier (Last)      : Park Hyun Soo
'* 11. Comment              :
'=======================================================================================================%>

<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/lgSvrVariables.inc" -->

<!-- #Include file="../../inc/adoVbs.inc" -->
<!-- #Include file="../../inc/incServerAdoDb.asp" -->

<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%
Call LoadBasisGlobalInf
Call LoadInfTB19029B("Q", "P", "NOCOOKIE", "MB")

Dim IntRetCD
Dim strPlantCd
Dim strResourceGroupCd
Dim strResourceCd

Call HideStatusWnd                                                               '��: Hide Processing message

On Error Resume Next                                                             '��: Protect system from crashing
Err.Clear                                                                        '��: Clear Error status


    '---------------------------------------Common-----------------------------------------------------------
    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '��: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '��: Read Operation Mode (CRUD)
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	
	strPlantCd = FilterVar(Trim(Request("txtPlantCd"))	, "''", "S")

	
	'------ Developer Coding part (End   ) ------------------------------------------------------------------ 

    Call SubOpenDB(lgObjConn)                                                        '��: Make a DB Connection
	
    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)													'��: ��üQuery
			strResourceGroupCd = FilterVar(Trim(Request("txtResourceGroupCd"))	, "''", "S")
		
			Call SubBizQueryMulti()
			Call SubBizQuery("RG")
			
        Case CStr(UID_M0002)								        							'��: Header Query
			strResourceGroupCd = FilterVar(Trim(Request("txtResourceGroupCd"))	, "''", "S")
			Call SubBizQuery("RG")
			
        Case CStr(UID_M0003)	
			strResourceCd = FilterVar(Trim(Request("txtResourceCd"))	, "''", "S")
			Call SubBizQuery("R")
			
    End Select
    
    Call SubCloseDB(lgObjConn)                                                       '��: Close DB Connection

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery(pOpCode)

	On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status
	
	Select Case pOpCode
		Case "RG"
			'--------------
			'�ڿ��׷� ��ȸ		
			'--------------	
			lgStrSQL = ""
			Call SubMakeSQLStatements("RG",strPlantCd,strResourceGroupCd)           '�� : Make sql statements
			
			If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then                    'If data not exists
		    
				IntRetCD = -1
				Call DisplayMsgBox("181704", vbInformation, "", "", I_MKSCRIPT)      '�� : No data is found. 
				Call SetErrorStatus()
%>
				<Script Language=vbscript>
					'------ Developer Coding part (Start ) ------------------------------------------------------------------
					' Set condition area, contents area
					'--------------------------------------------------------------------------------------------------------
					With Parent	
						.Frm1.txtResourceGroupNm.Value  = ""                   'Set condition area
						.Frm1.txtResourceGroupCd.Focus()
				    End With          
					'------ Developer Coding part (End   ) ------------------------------------------------------------------
				</Script>       
<%				Response.End
			
		    Else
				IntRetCD = 1
%>
				<Script Language=vbscript>
					'------ Developer Coding part (Start ) ------------------------------------------------------------------
					' Set condition area, contents area
					'--------------------------------------------------------------------------------------------------------
					With Parent	
						.Frm1.txtResourceGroupNm.Value			= "<%=ConvSPChars(lgObjRs("description"))%>"                   'Set condition area
				    End With          
					'------ Developer Coding part (End   ) ------------------------------------------------------------------
				</Script>       
<%			
			End If
		
			Call SubCloseRs(lgObjRs) 
			
		Case "R"																	'��: header ��ȸ ��� 
			'--------------
			'�ڿ� ��ȸ		
			'--------------	
			lgStrSQL = ""
			
		    Call SubMakeSQLStatements("R",strPlantCd,strResourceCd)           '�� : Make sql statements
 
		    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then                    'If data not exists
				
				IntRetCD = -1
				
				Call DisplayMsgBox("181600", vbInformation, "", "", I_MKSCRIPT)      '�� : No data is found. 
				Call SetErrorStatus()
				Response.End 
		    Else
				IntRetCD = 1
%>
				<Script Language=vbscript>
					'------ Developer Coding part (Start ) ------------------------------------------------------------------
					' Set condition area, contents area
					'--------------------------------------------------------------------------------------------------------
					With Parent	
						.frm1.txtResourceCd.value		= "<%=ConvSPChars(UCase(lgObjRs("resource_cd")))%>"
						.frm1.txtResourceNm.value		= "<%=ConvSPChars(lgObjRs("description"))%>"
						.frm1.txtResourceGroupCd2.value	= "<%=ConvSPChars(UCase(lgObjRs("resource_group_cd")))%>"
						.frm1.txtResourceGroupNm2.value	= "<%=ConvSPChars(lgObjRs("rg_nm"))%>"
						.frm1.txtResourceType.value		= "<%=ConvSPChars(lgObjRs("RType_CodeName"))%>"
						.frm1.txtNoOfResource.value		= "<%=lgObjRs("No_Of_Resource")%>"
						.frm1.txtEfficiency.value		= "<%=lgObjRs("Efficiency")%>"
						.frm1.txtUtilization.value		= "<%=lgObjRs("Utilization")%>"
						
						If "<%=ConvSPChars(lgObjRs("Run_Rccp"))%>" = "Y" Then
							.frm1.rdoRunRccp1.checked = True
						Else
							.frm1.rdoRunRccp2.checked = True
						End If
						
						If "<%=ConvSPChars(lgObjRs("Run_Crp"))%>" = "Y" Then
							.frm1.rdoRunCrp1.checked = True
						Else
							.frm1.rdoRunCrp2.checked = True
						End If
						
						.frm1.txtOverloadTol.value		= "<%=lgObjRs("Overload_Tol")%>"
						.frm1.txtResourceEa.value		= "<%=udf_UniConvNumberDBToCompany(lgObjRs("rsc_base_qty"),ggQty.DecPoint,ggQty.RndPolicy, ggQty.RndUnit, 0)%>"
						.frm1.txtResourceUnitCd.value	= "<%=ConvSPChars(UCase(lgObjRs("rsc_base_unit")))%>"
						.frm1.txtMfgCost.text			= "<%=udf_UniConvNumberDBToCompany(lgObjRs("mfg_cost"),ggUnitCost.DecPoint,ggUnitCost.RndPolicy, ggUnitCost.RndUnit, 0)%>"
						.frm1.txtResourceEa1.text		= "<%=udf_UniConvNumberDBToCompany(lgObjRs("rsc_base_qty"),ggQty.DecPoint,ggQty.RndPolicy, ggQty.RndUnit, 0)%>"
						.frm1.txtResourceUnitCd1.value	= "<%=ConvSPChars(lgObjRs("rsc_base_unit"))%>"
						.frm1.txtCurCd.value			= "<%=ConvSPChars(UCase(lgObjRs("cur_cd")))%>" 
						.frm1.txtValidFromDt.value		= "<%=UNIDateClientFormat(lgObjRs("valid_from_Dt"))%>"
						.frm1.txtValidToDt.value		= "<%=UNIDateClientFormat(lgObjRs("valid_To_Dt"))%>"
						
				    End With          
					'------ Developer Coding part (End   ) ------------------------------------------------------------------
				</Script>       
<%     
		    End If		    
		    
		    Call SubCloseRs(lgObjRs) 
		    
	End Select
    
End Sub    
'============================================================================================================
' Name : SubBizQueryMulti
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()
	
	On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status
    
'    Dim PrntKey
	Dim NodX
	Dim Node		
	Dim i    

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    '--------------
	'���� üũ		
	'--------------	
	lgStrSQL = ""
	Call SubMakeSQLStatements("P_CK",strPlantCd,"")           '�� : Make sql statements
			
	If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then                    'If data not exists
 		    
		IntRetCD = -1
		Call DisplayMsgBox("125000", vbInformation, "", "", I_MKSCRIPT)      '�� : No data is found. 
		Call SetErrorStatus()
%>
		<Script Language=vbscript>
			'------ Developer Coding part (Start ) ------------------------------------------------------------------
			' Set condition area, contents area
			'--------------------------------------------------------------------------------------------------------
			With Parent	
				.Frm1.txtPlantNm.Value  = ""                   'Set condition area
				.Frm1.txtPlantCd.Focus()
		    End With          
			'------ Developer Coding part (End   ) ------------------------------------------------------------------
		</Script>       
<%				Response.End
			
	Else
		IntRetCD = 1
%>
		<Script Language=vbscript>
			'------ Developer Coding part (Start ) ------------------------------------------------------------------
			' Set condition area, contents area
			'--------------------------------------------------------------------------------------------------------
			With Parent	
				.Frm1.txtPlantNm.Value			= "<%=ConvSPChars(lgObjRs("plant_nm"))%>"                   'Set condition area
				
		    End With          
			'------ Developer Coding part (End   ) ------------------------------------------------------------------
		</Script>       
<%			
	End If
		
	Call SubCloseRs(lgObjRs) 
    
    '--------------
	'�ڿ��׷� üũ		
	'--------------	
	Dim lgBlnResourceGroup
	lgBlnResourceGroup = False
	lgStrSQL = ""
	Call SubMakeSQLStatements("RG",strPlantCd,strResourceGroupCd)           '�� : Make sql statements
			
	If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then                    'If data not exists
 		    
		IntRetCD = -1
		Call DisplayMsgBox("181704", vbInformation, "", "", I_MKSCRIPT)      '�� : No data is found. 
		Call SetErrorStatus()
%>
		<Script Language=vbscript>
			'------ Developer Coding part (Start ) ------------------------------------------------------------------
			' Set condition area, contents area
			'--------------------------------------------------------------------------------------------------------
			With Parent	
				.Frm1.txtResourceGroupNm.Value  = ""                   'Set condition area
				.Frm1.txtResourceGroupCd.Focus()
		    End With          
			'------ Developer Coding part (End   ) ------------------------------------------------------------------
		</Script>       
<%				Response.End
			
	Else
		IntRetCD = 1
%>
		<Script Language=vbscript>
			'------ Developer Coding part (Start ) ------------------------------------------------------------------
			' Set condition area, contents area
			'--------------------------------------------------------------------------------------------------------
			With Parent	
				.Frm1.txtResourceGroupNm.Value			= "<%=ConvSPChars(lgObjRs("Description"))%>"                   'Set condition area
				
		    End With
			'------ Developer Coding part (End   ) ------------------------------------------------------------------
		</Script>
<%		lgBlnResourceGroup = True
	End If
		
	Call SubCloseRs(lgObjRs)
	
	'------------------------
    'Treeview ��ȸ	
    '------------------------
    '===========================================================================
	' TreeView ���� : Ű����(ù��°,����°����) ���ڰ� ���� ������ ���´�.
	' ��ġ����      : ������ �����ڿ� �����Ͽ� Ű���� �����. "A"�� ÷����.
	'===========================================================================
    Call SubMakeSQLStatements("M",strPlantCd,strResourceGroupCd)                                   '�� : Make sql statements
    
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
		If lgBlnResourceGroup Then
%>    
		<Script Language=vbscript>
			With parent.frm1.uniTree1
				Set NodX = .Nodes.Add(,,"A" & "<%=ConvSPChars(UCase(Trim(Request("txtResourceGroupCd"))))%>","<%=ConvSPChars(UCase(Trim(Request("txtResourceGroupCd"))))%>",parent.C_GROUP, parent.C_GROUP)      
					NodX.Expanded = True
				Set NodX = Nothing
			End With
		</Script>
<%
		End If
		Call SetErrorStatus()		
    Else

%>
	<Script Language=vbscript>
		With parent.frm1.uniTree1
			Set NodX = .Nodes.Add(,,"A" & "<%=ConvSPChars(UCase(Trim(Request("txtResourceGroupCd"))))%>","<%=ConvSPChars(UCase(Trim(Request("txtResourceGroupCd"))))%>",parent.C_GROUP, parent.C_GROUP)      
				NodX.Expanded = True
			Set NodX = Nothing
		
			.MousePointer = 11														'��: ���콺 ����Ʈ ��ȭ 
			.Indentation = 50														'��: �θ�Ʈ���� �ڽ�Ʈ�� ������ ���� 

<%
			Do While Not lgObjRs.EOF
%>
				Set Node = .Nodes.Add("A" & "<%=ConvSPChars(UCase(Trim(Request("txtResourceGroupCd"))))%>", parent.tvwChild, "A" & "<%=ConvSPChars(Trim(lgObjRs("resource_cd")))%>" , "<%=ConvSPChars(Trim(lgObjRs("resource_cd")))%>" , parent.C_PROD, parent.C_PROD)
					Node.Expanded = True
<%	
				lgObjRs.MoveNext
				
			Loop
%>
			.MousePointer = 1
			Set Node = Nothing
		End With
	</Script>
<%

    End If

	Call SubHandleError("MR",lgObjConn,lgObjRs,Err)
    Call SubCloseRs(lgObjRs)                                                          '��: Release RecordSSet
    
End Sub    

'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements(pDataType,pCode,pCode1)
    Dim iSelCount
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
    Select Case pDataType
		
		Case "M"
			lgStrSQL = "SELECT * FROM p_resource "
			lgStrSQL = lgStrSQL & " WHERE plant_cd = " & pCode
			lgStrSQL = lgStrSQL & " AND resource_group_cd = " & pCode1
			
		Case "R"
			lgStrSQL = " Select a.*, e.cur_cd, b.description sl_nm, c.description sq_nm , d.description rg_nm,"
			lgStrSQL = lgStrSQL & " (select Minor_Nm from B_Minor where Major_cd=" & FilterVar("p1502", "''", "S") & " and Minor_cd = a.Resource_Type) as RType_CodeName "
			lgStrSQL = lgStrSQL & " From p_resource a, p_aps_rule_detail b,  p_aps_rule_detail c, p_resource_group d, b_plant e "
			lgStrSQL = lgStrSQL & " WHERE a.selection_rule *= b.rule_type and a.sequence_rule *= c.rule_type "
			lgStrSQL = lgStrSQL & " AND b.rule_type_cd = " & FilterVar("RSSLRL", "''", "S") & " and c.rule_type_Cd= " & FilterVar("RSSQRL", "''", "S") & " "
			lgStrSQL = lgStrSQL & " AND a.resource_group_cd = d.resource_group_cd "
			lgStrSQL = lgStrSQL & " AND a.plant_cd = " & pCode
			lgStrSQL = lgStrSQL & " AND e.plant_cd = " & pCode
			lgStrSQL = lgStrSQL & " AND a.resource_cd >= " & pCode1
		Case "RG"
			lgStrSQL = "SELECT * FROM p_resource_group a, b_plant b "
			lgStrSQL = lgStrSQL & " WHERE a.plant_cd = b.plant_cd and a.plant_Cd = " & pCode
			lgStrSQL = lgStrSQL & " AND a.resource_group_cd = " & pCode1
		Case "P_CK"
			lgStrSQL = "SELECT * FROM b_plant where plant_cd = " & pCode 
			
    End Select

	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub
'============================================================================================================
' Name : CommonOnTransactionCommit
' Desc : This Sub is called by OnTransactionCommit Error handler
'============================================================================================================
Sub CommonOnTransactionCommit()
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub

'============================================================================================================
' Name : CommonOnTransactionAbort
' Desc : This Sub is called by OnTransactionAbort Error handler
'============================================================================================================
Sub CommonOnTransactionAbort()
    lgErrorStatus    = "YES"
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub

'============================================================================================================
' Name : SetErrorStatus
' Desc : This Sub set error status
'============================================================================================================
Sub SetErrorStatus()
    lgErrorStatus     = "YES"                                                         '��: Set error status
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub
'============================================================================================================
' Name : SubHandleError
' Desc : This Sub handle error
'============================================================================================================
Sub SubHandleError(pOpCode,pConn,pRs,pErr)
    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status

    Select Case pOpCode
        Case "MC"
            If CheckSYSTEMError(pErr,True) = True Then
               ObjectContext.SetAbort
               Call SetErrorStatus
            Else
               If CheckSQLError(pConn,True) = True Then
                  ObjectContext.SetAbort
                  Call SetErrorStatus
               End If
            End If
        Case "MD"
        Case "MR"
        Case "MU"
            If CheckSYSTEMError(pErr,True) = True Then
               ObjectContext.SetAbort
               Call SetErrorStatus
            Else
               If CheckSQLError(pConn,True) = True Then
                  ObjectContext.SetAbort
                  Call SetErrorStatus
               End If
            End If
        Case "MB"
			ObjectContext.SetAbort
            Call SetErrorStatus        
    End Select
End Sub


'==============================================================================
' ����� ���� ���� �Լ� 
'==============================================================================
'==============================================================================
' Function Name : udf_UniConvNumberDBToCompany
' Function Desc : �ִ밪�� ���� udf_UniConvNumberDBToCompany�Լ��� ����ϸ� 
'				�ݿø� ��å�� ���� �ִ밪�� �Ѿ� ���� ���� ���� �ϱ� ���� �Լ� 
' 
'==============================================================================
Function udf_UniConvNumberDBToCompany(ByVal pNum,ByVal pDecPoint,ByVal pRndPolicy, ByVal pRndUnit, ByVal pDefault)

	Dim rtnNum
	
	Const maxNum	= 99999999999.9999	'�ִ밪 (�ʵ��� �Ӽ��� ���� ���� ����)
	Const maxDecPnt = 4					'�Ҽ��� ���� �ִ��ڸ��� (�ʵ��� �Ӽ��� ���� ���� ����,�ý��� ������������ ���밡���� �ִ��ڸ���)
	
	rtnNum = UniConvNumberDBToCompany(pNum, pDecPoint, pRndPolicy, pRndUnit, pDefault)
	
	If rtnNum > UniConvNumberDBToCompany(maxNum,pDecPoint, pRndPolicy, pRndUnit,pDefault) Then	'�ִ밪���� ū ���϶� ���� 
		If pDecPoint <> maxDecPnt Then							'�Ҽ��� ���� �ִ밪�� �ƴҶ��� ���� 
			rtnNum = int(cdbl(pNum) * cdbl(10 ^ pDecPoint))
			rtnNum = rtnNum * cdbl(pRndUnit) * 10
		End if
		udf_UniConvNumberDBToCompany = UniConvNumberDBToCompany(rtnNum,pDecPoint, pRndPolicy, pRndUnit,pDefault)
	Else
		udf_UniConvNumberDBToCompany = rtnNum
	End If
	
End Function
%>

<Script Language="VBScript">
    Select Case "<%=lgOpModeCRUD %>"
       Case "<%=UID_M0001%>"                                                         '�� : Query
          If Trim("<%=lgErrorStatus%>") = "NO" And <%=IntRetCd%> <> -1 Then
              With Parent
                .DBQueryOk()        
	         End with
          End If   
      
    End Select    
       
</Script>	
