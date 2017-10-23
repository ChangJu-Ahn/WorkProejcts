function InvokePop(fname) {
    val = document.getElementById(fname).value;    
    
    // to handle in IE 7.0           
    if (window.showModalDialog) {        
        
        retVal = window.showModalDialog("../../../pop_item_cd.aspx?Control1=" + fname + "&ControlVal=" + val, 'Show Popup Window', "dialogHeight:360px,dialogWidth:360px,resizable:yes,center:yes,");
        document.getElementById(fname).value = retVal;
    }
    // to handle in Firefox
    else {
        retVal = window.open("pop_item_cd.aspx?Control1=" + fname + "&ControlVal=" + val, 'Show Popup Window', 'height=450,width=550,resizable=yes,modal=yes');
        retVal.focus();
    }
}

function ReverseString() {
    var originalString = document.getElementById('tb_pop_item_cd').value;
    var reversedString = Reverse(originalString);
    RetrieveControl();
    // to handle in IE 7.0
    if (window.showModalDialog) {
        window.returnValue = reversedString;
        window.close();
    }
    // to handle in Firefox
    else {
        if ((window.opener != null) && (!window.opener.closed)) {
            // Access the control.        
            window.opener.document.getElementById(ctr[1]).value = reversedString;
        }
        window.close();
    }
}

function Reverse(str) {
    var revStr = "";
    for (i = str.length - 1; i > -1; i--) {
        revStr += str.substr(i, 1);
    }
    return revStr;
}

