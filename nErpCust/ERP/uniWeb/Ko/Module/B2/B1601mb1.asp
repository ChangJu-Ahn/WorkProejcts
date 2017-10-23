<%
'**********************************************************************************************
'*  1. Module Name          : Basis Architect
'*  2. Function Name        : Master Data(Country)
'*  3. Program ID           : B1601mb1.asp
'*  4. Program Name         : B1601mb1.asp
'*  5. Program Desc         :
'*  6. Comproxy List        : +B16011ControlCountry
'                             +B16018ListCountry
'*  7. Modified date(First) : 2000/09/07
'*  8. Modified date(Last)  : 2002/12/12
'*  9. Modifier (First)     : Hwang Jeong-won
'* 10. Modifier (Last)      : Sim Hae Young
'* 11. Comment              :
'**********************************************************************************************

%>
<% Option Explicit %>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<%													'�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 
Call HideStatusWnd															'��: ��� �۾� �Ϸ��� �۾������� ǥ��â�� Hide

On Error Resume Next									'��: 
Err.Clear												'��: Protect system from crashing

Dim strMode	
Dim strSpread																'��: ���� MyBiz.asp �� ������¸� ��Ÿ�� 

Dim StrNextKey		' ���� �� 
Dim lgStrPrevKey	' ���� �� 
Dim LngMaxRow		' ���� �׸����� �ִ�Row
Dim LngRow
Dim GroupCount          

Call LoadBasisGlobalInf()

strMode = Request("txtMode")												'�� : ���� ���¸� ���� 
strSpread = Trim(Request("txtSpread"))

Select Case strMode

Case CStr(UID_M0001)														'��: ���� ��ȸ/Prev/Next ��û�� ���� 

    Dim I1_b_country
    Const B352_I1_country_cd = 0 
    Const B352_I1_country_nm = 1

    Const B352_EG1_E1_country_cd = 0 
    Const B352_EG1_E1_country_nm = 1
    Const B352_EG1_E1_region_cd = 2
    Const B352_EG1_E1_minor_nm = 3   
    Const B352_EG1_E1_dial_cd = 4

	Dim ObjPB2G131
	Dim Export_Array
    
    ReDim I1_b_country(B352_I1_country_nm)
%>
<Script Language=vbscript>
	parent.frm1.txtCountryNm.value = "<%=ConvSPChars(LookUpCountry(Request("txtCountryCd")))%>"
</Script>
<%  
    I1_b_country(B352_I1_country_cd) = Request("txtCountryCd")

    Set ObjPB2G131 = server.CreateObject ("PB2G131.cBListCountry")    
    on error resume next
    Err.Clear 
    Export_Array = ObjPB2G131.B_LIST_COUNTRY(gStrGlobalCollection,I1_b_country)
    Set ObjPB2G131 = nothing

    If CheckSYSTEMError(Err,True) = True Then      
        Response.End 
    End If
    on error goto 0
    
%>
<Script Language=vbscript>
	Dim LngLastRow      
	Dim LngMaxRow       
	Dim LngRow          
	Dim strTemp
	Dim strData
	
	With parent																	'��: ȭ�� ó�� ASP �� ��Ī�� 

		LngMaxRow = 0
<%      
        GroupCount = Ubound(Export_Array,1)
	    For LngRow = 0 To GroupCount
%>        
            strData = strData & Chr(11) & "<%=ConvSPChars(Trim(Export_Array(LngRow,B352_EG1_E1_country_cd)))%>"
            strData = strData & Chr(11) & "<%=ConvSPChars(Trim(Export_Array(LngRow,B352_EG1_E1_country_nm)))%>"
            strData = strData & Chr(11) & "<%=ConvSPChars(Trim(Export_Array(LngRow,B352_EG1_E1_region_cd)))%>"
            
            strData = strData & Chr(11) & "<%=ConvSPChars(Trim(Export_Array(LngRow,B352_EG1_E1_minor_nm)))%>"
            strData = strData & Chr(11) & "<%=ConvSPChars(Trim(Export_Array(LngRow,B352_EG1_E1_dial_cd)))%>"
            
            strData = strData & Chr(11) & LngMaxRow + <%=LngRow%> + 1
            strData = strData & Chr(11) & Chr(12)
<%      		
        Next
%>    		
		.ggoSpread.Source = .frm1.vspdData 
		.ggoSpread.SSShowData strData

		.frm1.hCountryCd.value = "<%=ConvSPChars(Request("txtCountryCd"))%>"		
		.DbQueryOk
	End With
</Script>	
<%    
Case CStr(UID_M0002)																'��: ���� ��û�� ���� 
									
	If Request("txtMaxRows") = "" Then
		Call DisplayMsgBox("700117", vbInformation, "", "", I_MKSCRIPT)
		Response.End 
	End If
	
    Dim Obj2PB2G131
    Dim iErrorPosition

    Set Obj2PB2G131 = server.CreateObject ("PB2G131.cBControlCountry")    
    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear        
    Call Obj2PB2G131.B_CONTROL_COUNTRY (gStrGlobalCollection,strSpread,iErrorPosition)
    Set Obj2PB2G131 = nothing

    If CheckSYSTEMError2(Err,True,iErrorPosition & "��","","","","") = True Then 
        Response.End 
    End If
    on error goto 0                                                             
    
%>
<Script Language=vbscript>
	With parent																		'��: ȭ�� ó�� ASP �� ��Ī�� 
		'window.status = "���� ����"
		.DbSaveOk
	End With
</Script>
<%					
End Select
%>



<%
'==============================================================================
' Function : LookUp...
' Description : ����� Lookup
'==============================================================================
Function LookUpCountry(Byval strCode)
    Const B250_I1_country_cd = 0
    Const B250_I1_country_nm = 1

    Const B250_E1_country_cd = 0
    Const B250_E1_country_nm = 1
    Const B250_E1_region_cd = 2
    Const B250_E1_dial_cd = 3

	Dim ObjPB0C003
	Dim I1_b_country
	Dim E1_b_country
	
    ReDim I1_b_country(B250_I1_country_nm)
    ReDim E1_b_country(B250_E1_dial_cd)

    I1_b_country(B250_I1_country_cd) = strCode
    I1_b_country(B250_I1_country_nm) = ""

    Set ObjPB0C003 = server.CreateObject ("PB0C003.CB0C003")    
    
    On Error Resume Next
    Err.Clear                                                                            '��: Clear Error status
    E1_b_country = ObjPB0C003.B_SELECT_COUNTRY(gStrGlobalCollection, I1_b_country)
    Set ObjPB0C003 = nothing    
    
    If Err.number <> 0 and inStr(Err.Description ,"121300") > 0 then
	LookUpCountry = ""
    Else
        If CheckSYSTEMError(Err,True) = True Then                                              
        	Exit Function
	    End If
        on error goto 0
        	LookUpCountry  = E1_b_country(B250_E1_country_nm)
    End If
    
End Function
%>

