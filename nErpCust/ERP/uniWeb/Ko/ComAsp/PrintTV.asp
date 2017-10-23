<STYLE TYPE="text/css">
TD
{
    FONT-SIZE: 9pt;
    CURSOR: default
}
		
</STYLE>

<SCRIPT LANGUAGE="VBScript">

Option Explicit                              '¢Ð: indicates that All variables must be declared in advance


Dim arrReturn
Dim arrParent
Dim arrParam
dim i 
dim j					

	 '------ Set Parameters from Parent ASP ------ 
Sub Window_OnLoad()
    Dim strHTML
    Dim iPlusMark
    Dim iDX
    Dim iSplit
    
    
	arrParent = Window.DialogArguments
	arrParam = arrParent(0)
	
    ReDim iPlusMark(UBound(arrParam,2))
    
    For iDx = 0 To UBound(arrParam,2)
      iPlusMark(iDx) = "+"
    Next   
	
	
    strHTML = "<TABLE border=0 cellspacing=0 cellpadding=0>"
      For i = 0 To UBound(arrParam,1)
          strHTML = strHTML &  "<TR>"
          For j = 0 To UBound(arrParam,2)
              If Mid(arrParam(i,j),1,1) = "C" Then
                 iPlusMark(j) = "+"              
              End If
              If Mid(arrParam(i,j),1,1) = "E" Then
                 iPlusMark(j) = "+"              
              End If
          
              If j > 0 Then
                 strHTML = strHTML &  "<TD>" & iPlusMark(j) & "</TD>"
              End If
              
              If Mid(arrParam(i,j),1,1) = "C" Then
                 iPlusMark(j) = "|"              
              End If
              
              If Mid(arrParam(i,j),1,1) = "E" Then
                 iPlusMark(j) = ""              
              End If
              
              If Trim(arrParam(i,j)) = "" Then
                 strHTML = strHTML &  "<TD>&nbsp;</TD>"
              Else
                 strHTML = strHTML &  "<TD>"  & arrParam(i,j)  & "</TD>"
              End If   
          Next
          strHTML = strHTML &  "</TR>"
      Next
      
    strHTML = strHTML &  "</TABLE>"
    
    document.all("divRawXML").innerHTML = strHTML
End Sub 	

</SCRIPT>

<BODY>
<DIV id=divRawXML></DIV>
<table>
<tr>
  <td> <INPUT TYPE=BUTTON VALUE ="PRINT" onClick="print()">
</tr>
<table>
</BODY>
