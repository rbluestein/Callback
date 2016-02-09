var BrowserIsIE = 0
var BrowserIsChrome = 0
var BrowserIsFirefox = 0
var BrowserName

function JSLibraryTest() {
	alert("In JSLibrary.js")
}

function Left(str, n){
	if (n <= 0)
	    return "";
	else if (n > String(str).length)
	    return str;
	else
	    return String(str).substring(0,n);
}
function Right(str, n){
    if (n <= 0)
       return "";
    else if (n > String(str).length)
       return str;
    else {
       var iLen = String(str).length;
       return String(str).substring(iLen, iLen - n);
    }
}

function AllowAlphaNumericOnly(txtbox)
{
		var ToTest = txtbox.value
		var re = /[^0-9A-Za-z]/
		if (re.test(ToTest)) {
			txtbox.value = txtbox.value.substring(0, txtbox.value.length-1)
		}
}

function KeypressNumeric(txtbox) {
    if (event.keyCode < 48 || event.keyCode > 57) {
        return false
    }
}
                   
function NumericOnly(txtbox) {
    var OldValue
    var NewValue = ""
    var Char
    var re = /[^0-9]/    
    
    if (re.test(txtbox.value)) {
        OldValue = txtbox.value                        
        re = /[0-9]/
        
        for (i=0; i<OldValue.length; i++) {
            Char = OldValue.substring(i , i+1)
            if (re.test(Char)) {
                NewValue = NewValue + parseInt(Char, 10)
            }
        } 
        txtbox.value = NewValue                                     
    }                       
} 

function AllowIntegerOnly(txtbox) {
		var ToTest = txtbox.value
		var re = /[^0-9]/
		if (re.test(ToTest)) {
			txtbox.value = txtbox.value.substring(0, txtbox.value.length-1)
		}
}


function OLDAllowIntegerOnly(txtbox)  {
	var i
	var ValidChars = "0123456789";
	var Text
	var Char
	var IsValid=true

	Text = txtbox.value
	for (i=0;i<Text.length;i++)  {
		Char = Text.charAt(i)
		if (ValidChars.indexOf(Char) == -1)  {
			IsValid = false
		}	
	}
	if (IsValid == false)  {
		txtbox.value = txtbox.value.substring(0, txtbox.value.length-1)
	}				
}


function IsNumericValueOnly(vIn) {
	var re = /[^0-9]/
	if ((re.test(vIn))  || (vIn == ""))   {
		return 0
	} else {
		return 1	
	}						
}



function IsValidChar(txtbox, AllowDecimal)  {
	var ValidChars
	var IsValid=true
	var Char
	var Text
	var DecPos=0
	var HasDec = 0
	var i
	var results
	
	if (AllowDecimal == 1) {
		ValidChars = ".0123456789";
	} else {
		ValidChars = "0123456789";
	}
	
	Text = txtbox.value
	for (i=0;i<Text.length && IsValid;i++)  {
		Char = Text.charAt(i)
		if (Char == "$" && i > 0)  {
			IsValid = false
		}	
		if (Char == "-" && i > 0)  {
			IsValid = false
		}
		if (Char == ".")  {
			if (HasDec == 1)  {
				IsValid = false
			} else {
				HasDec = 1
				DecPos = i
			}
		}
		if (DecPos > 0)  {
			if (i > (DecPos + 2))  {
				IsValid = false
			}
		}
		if (ValidChars.indexOf(Char) == -1)  {
			IsValid = false
		}		
	}
	if (IsValid == false)  {
		txtbox.value = txtbox.value.substring(0,i-1)
		results = 0
	} else {
		results = 1
	}
	return results
}


function IsValidCharMoney(txtbox, RequireDecimal, MaxLength)  {
	var ValidChars
	var IsValid=true
	var Char
	var Text
	var DecPos= -1
	var HasDec = 0
	var i
	var results
	
	if (RequireDecimal == 1) {
		ValidChars = ".0123456789";
	} else {
		ValidChars = "0123456789";
	}
	
	Text = txtbox.value
	for (i=0;i<Text.length && IsValid;i++)  {
		
		// Get the character for test
		Char = Text.charAt(i)
		
		// Prohibit any invalid character		
		if (ValidChars.indexOf(Char) == -1)  {
			IsValid = false
		}			
		
		// Prohibit 0 in first position	
		//if (IsValid == true)  {	
		//	if (i == 0 && Char == "0")  {
		//		IsValid = false
		//	}			
		//}

		// Record the decimal position. Do not allow a second decimal
		if (IsValid == true)  {
			if (Char == ".")  {
				if (HasDec == 1)  {
					IsValid = false
				} else {
					HasDec = 1
					DecPos = i
				}
			}
		}
		
		// Provide for decimal places
		if (IsValid == true)  {
			if (((i+1) == (MaxLength - 2)) && (HasDec == 0) && (Char != "."))  {
				IsValid = false
			}
		}		
		
		
		// Enforce the number of decimal places
		if (IsValid == true)  {
			if (DecPos > -1)  {
				if (i > (DecPos + 2))  {
					IsValid = false
				}
			}
		}

	}
	if (IsValid == false)  {
		txtbox.value = txtbox.value.substring(0,i-1)
		results = 0
	} else {
		results = 1
	}
	return results
}


function ORIGINAL_IsValidChar(txtbox)  {
	var ValidChars = ".,0123456789";
	var IsValid=true
	var Char
	var Text
	var DecPos=0
	var HasDec = 0
	var i
	var results
	
	Text = txtbox.value
	for (i=0;i<Text.length && IsValid;i++)  {
		Char = Text.charAt(i)
		if (Char == "$" && i > 0)  {
			IsValid = false
		}	
		if (Char == "-" && i > 0)  {
			IsValid = false
		}
		if (Char == ".")  {
			if (HasDec == 1)  {
				IsValid = false
			} else {
				HasDec = 1
				DecPos = i
			}
		}
		if (DecPos > 0)  {
			if (i > (DecPos + 2))  {
				IsValid = false
			}
		}
		if (ValidChars.indexOf(Char) == -1)  {
			IsValid = false
		}		
	}
	if (IsValid == false)  {
		txtbox.value = txtbox.value.substring(0,i-1)
		results = 0
	} else {
		results = 1
	}
	return results
}

function fnReplace(vIn, vChar1, vChar2)  {
	var curchar
	var newStr = ""
	var i
	for (i=0;i<vIn.length;i++)  {
		curchar = vIn.charAt(i)
		if (curchar == vChar1)  {
			newStr = newStr + vChar2
		} else {
			newStr = newStr + curchar
		}
	}
		return newStr
}

function roundNumber(number,decimal_places)  {
	if (!decimal_places) return Math.round(number);
	if (number == 0)  {
		var decimals = "";
		for (var i=0;i<decimal_points;i++)  decimals += "0";
		return "0." + decimals;
	}	
	var exponent = Math.pow(10,decimal_places)
	var num = Math.round((number * exponent)).toString();
	return num.slice(0,-1*decimal_places) + "." + num.slice(-1*decimal_places)	
}

function FormatMoney(Input)  {
	var Output
	var Box

	Input = Input.toString()
	if (Input.indexOf(".") == -1) {
		Input = Input + "."
	}
	Box = Input.split(".")
	if (Box[1].length == 0) {
		Box[1] = Box[1] + "00"
	}
	else if (Box[1].length == 1)  {
		Box[1] = Box[1] + "0"
	}	
	else if (Box[1].length > 2)  {
		Box[1] = Box[1].substring(0, 2)
	}
	Output = Box[0] + "." + Box[1]
	return Output
}


function IsNumeric(sText)  {
   var ValidChars = "0123456789.";
   var IsNumber;
   var Char;
   
   if (sText.length == 0) {
        IsNumber = false;
   } else {

           IsNumber=true;
           for (i = 0; i <= sText.length && IsNumber == true; i++)  { 
              Char = sText.charAt(i); 
              if (ValidChars.indexOf(Char) == -1)  {
                 IsNumber = false;
                 break;
               }
            }
  
    }
    
   return IsNumber;
 }
 
 
 
 
function NumericOnlyValue(Value) {
    var NewText = ""
    var Char
    var re = /[^0-9]/    
    
    // ___ Exit if all of the characters are numeric
    if (re.test(Value)) {
    } else {
          return Value
    }
    
    // ___ Build the return string
    re = /[0-9]/
    for (i=0; i<=Value.length; i++) {
        Char = Value.substring(i , i+1)
        if (re.test(Char)) {
            NewText = NewText + parseInt(Char, 10)
        }
    }
    return NewText
}  


// TRIM FUNCTIONS  
function trim(stringToTrim) {
    return stringToTrim.replace(/^\s+|\s+$/g,"");
}
function ltrim(stringToTrim) {
    return stringToTrim.replace(/^\s+/,"");
}
function rtrim(stringToTrim) {
    return stringToTrim.replace(/\s+$/,"");
}	

//function EnterKeySubmit(e) {
//    var KeyID
//    if (BrowserIsIE == 1) {
//        KeyID = window.event.keyCode
//    } else {
//        KeyID = e.charCode
//    }
//    if (KeyID == 13) {
//        //fnSubmit(Action)
//        document.getElementsByTagName("form")[0].submit()
//    }
//}

//function DoBrowser() {
//    var BrowserName = document.getElementById("hdBrowserName").value
//    switch (BrowserName) {
//        case "IE":
//            BrowserIsIE = 1
//            BrowserName = "IE"
//            break;
//        case "Firefox":
//            BrowserIsFirefox = 1
//            BrowserName = "Firefox"
//            break;
//        case "Chrome":
//            BrowserIsChrome = 1
//            BrowserName = "Chrome"
//            break;
//        default:
//            fnMeKickedOffBrowserUnsupported(BrowserName)
//            break;
//    }
//}