<% 
'**********************************************************************************************
'*  1. Module Name          : Basis Architect
'*  2. Function Name        : Master Data(Unit of Measure)
'*  3. Program ID           : B1f03mb1.asp
'*  4. Program Name         : B1f03mb1.asp
'*  5. Program Desc         :
'*  6. Comproxy List        : 
'                             +B1F031ControlUnitOfMeasure
'                             +B1F038ListUnitOfMeasure
'*  7. Modified date(First) : 2000/09/05
'*  8. Modified date(Last)  : 2002/06/19
'*  9. Modifier (First)     : Hwang Jeong-won
'* 10. Modifier (Last)      : Sim Hae Young
'* 11. Comment              : 
'**********************************************************************************************
%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<%                           												'�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 
Call HideStatusWnd
On Error Resume Next														'��: 

Dim comB1F031																	'�� : �Է�/������ ComProxy Dll ��� ���� 
Dim comB1F038																'�� : ��ȸ�� ComProxy Dll ��� ���� 

Dim strMode																	'��: ���� MyBiz.asp �� ������¸� ��Ÿ�� 
Dim strSpread

Dim StrNextKey		' ���� ��(Unit)
Dim StrNextKey2		' ���� ��(Dimension)
Dim lgStrPrevKey	' ���� ��(Unit)
Dim lgStrPrevKey2	' ���� ��(Dimension)
Dim LngMaxRow		' ���� �׸����� �ִ�Row
Dim LngRow
Dim GroupCount      

Call LoadBasisGlobalInf()    

strMode = Request("txtMode")												'�� : ���� ���¸� ���� 
strSpread = Trim(Request("txtSpread"))

Select Case strMode

Case CStr(UID_M0001)														'��: ���� ��ȸ/Prev/Next ��û�� ���� 

    Dim I1_b_unit_of_measure
    
    Const B440_I1_unit = 0
    Const B440_I1_dimension = 1
 
    Const B440_EG1_E1_minor_nm = 0
    Const B440_EG1_E2_unit = 1
    Const B440_EG1_E2_unit_nm = 2
    Const B440_EG1_E2_dimension = 3
    Const B440_EG1_E2_symbol = 4
    
	Dim ObjPB2G091
    ReDim I1_b_unit_of_measure(B440_I1_dimension)	
	Dim Export_Array
    
%>
<Script Language=vbscript>
	parent.frm1.txtUnitNm.value = "<%=ConvSPChars(LookUpUnit(Request("txtUnit")))%>"
</Script>
<%  

    I1_b_unit_of_measure(B440_I1_unit) = Request("txtUnit")
    I1_b_unit_of_measure(B440_I1_dimension) = Request("txtDimensionCd")

    Set ObjPB2G091 = server.CreateObject("PB2G091.cBListUnitOfMeasure")    

    on error resume next
    Err.Clear 
    Export_Array = ObjPB2G091.B_LIST_UNIT_OF_MEASURE(gStrGlobalCollection,I1_b_unit_of_measure)
    Set ObjPB2G091 = nothing

    If CheckSYSTEMError(Err,True) = True Then                               
		Response.End														'��: �����Ͻ� ���� ó���� ������ 
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
		strData = strData & Chr(11) & "<%=ConvSPChars(Trim(Export_Array(LngRow,B440_EG1_E1_minor_nm )))%>"     
		strData = strData & Chr(11) & "<%=ConvSPChars(Trim(Export_Array(LngRow,B440_EG1_E2_dimension )))%>"  
		strData = strData & Chr(11) & "<%=ConvSPChars(Trim(Export_Array(LngRow,B440_EG1_E2_unit )))%>" 
		strData = strData & Chr(11) & "<%=ConvSPChars(Trim(Export_Array(LngRow,B440_EG1_E2_unit_nm )))%>"      
		strData = strData & Chr(11) & "<%=ConvSPChars(Trim(Export_Array(LngRow,B440_EG1_E2_symbol )))%>"
		strData = strData & Chr(11) & LngMaxRow + <%=LngRow%> + 1
		strData = strData & Chr(11) & Chr(12)
<%
    Next
%>      
		.ggoSpread.Source = .frm1.vspdData 
		.ggoSpread.SSShowData strData
		.frm1.hDimension.value = "<%=Request("cboDimension")%>"
		.frm1.hUnit.value = "<%=ConvSPChars(Request("txtUnit"))%>"			
	
		.DbQueryOk	
		
	End With
</Script>	
<%    
    
Case CStr(UID_M0002)																'��: ���� ��û�� ���� 
    Dim Obj2PB2G091
    Dim iErrorPosition

    Set Obj2PB2G091 = server.CreateObject ("PB2G091.cBCtrlUnitOfMeasure")    
    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear        
    
    Call Obj2PB2G091.B_CONTROL_UNIT_OF_MEASURE (gStrGlobalCollection,strSpread,iErrorPosition)
    Set Obj2PB2G091 = nothing

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
Function LookUpUnit(Byval strCode)
    Const B267_I1_unit = 0
    Const B267_I1_unit_nm = 1
    Const B267_I1_dimension = 2

    Const B267_E1_unit = 0
    Const B267_E1_unit_nm = 1
    Const B267_E1_dimension = 2
    Const B267_E1_symbol = 3
    
	Dim ObjPB0C006		
	Dim I1_b_unit_of_measure
	Dim E1_b_unit_of_measure
	
    ReDim I1_b_unit_of_measure(B267_I1_dimension)
    ReDim E1_b_unit_of_measure(B267_E1_dimension)
    
    I1_b_unit_of_measure(B267_I1_unit) = strCode
    
    Set ObjPB0C006 = server.CreateObject ("PB0C006.CB0C006")    
    
    On Error Resume Next
    Err.Clear                                                                            '��: Clear Error status
    E1_b_unit_of_measure = ObjPB0C006.B_SELECT_UNIT_OF_MEASURE (gStrGlobalCollection,I1_b_unit_of_measure)
    Set ObjPB0C006 = nothing    
    
    If Err.number <> 0 and inStr(Err.Description ,"124000") > 0 then
  	LookUpUnit = ""
    Else
        If CheckSYSTEMError(Err,True) = True Then                                              
        	Exit Function
	    End If
        on error goto 0
        LookUpUnit = E1_b_unit_of_measure(B267_E1_unit_nm)
    End If


End Function
%>
