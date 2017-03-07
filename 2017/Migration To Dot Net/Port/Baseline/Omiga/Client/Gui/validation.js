/*====================================================================
Function: Err
Purpose:  Custom object constructor
Inputs:   None
Returns:  undefined
Changes:  10/10/05  MahaT	New Message for Min & Max (MAR95)
		  10/11/05 	MAR273	disable BACKSPACE to navigate to prev page 
		  29/11/05	SD		MAR614 Allow backspace for PostCode, for class 'msgTxtUpper' 
		  30/11/06	AS		EP1251 Internet Explorer process is leaking memory when using Omiga
							Removed memory leaks due to function closures creating circular
							references between DOM and Javascript.
							See http://msdn.microsoft.com/library/default.asp?url=/library/en-us/IETechCol/dnwebgen/ie_leak_patterns.asp
							and http://www.codeproject.com/jscript/leakpatterns.asp.
====================================================================*/

function Err(){
	this.clear=clear;
	this.add=add;
	this.raise=raise;

	// Define the working object model
	this.clear();

	//////////////////////////////////////////////////////////////////////
	//Method:   Err.clear
	//Purpose:  Clear values from Error object
	//Inputs:   none
	//Returns:  undefined
	//////////////////////////////////////////////////////////////////////
	function clear()
	{
		this.source=new Object;
		this.type=new Object;
		this.format=new String;
	}

	//////////////////////////////////////////////////////////////////////
	//Method:   Err.add
	//Purpose:  Adds error to Error object
	//Inputs:   oSource - source element object
	//          vType   - integer value of error type (or custom string)
	//          sFormat - optional date format
	//Returns:  undefined
	//////////////////////////////////////////////////////////////////////
	function add(oSource,vType,sFormat)
	{
		this.source=oSource;
		this.type=vType;
		this.format=sFormat;
	}

	//////////////////////////////////////////////////////////////////////
	//Method:   Err.raise
	//Purpose:  Gives visual warning to user about all errors contained in
	//          the Error object
	//Inputs:   none
	//Returns:  undefined
	//////////////////////////////////////////////////////////////////////
	function raise()
	{
		var oElement=this.source;
		var sLang;
		var sNym=oElement.getAttribute("nym");
		// if type is not a number, it must be a custom error message
		var sMsg=(typeof this.type=="string")?this.type:oElement.getAttribute("msg");

//		oElement.paint();
		if(oElement.select)
			oElement.select();
		if(sMsg)
			alert(sMsg);
		else{
			// Walk through object hierarchy to find applicable language
			var oParent=oElement;
			sLang=oParent.getAttribute("lang").substring(0,2).toLowerCase();
			while(!sLang || !_validation.messages[sLang]){
				oParent=oParent.parentElement;
				if(oParent)
					sLang=oParent.getAttribute("lang").substring(0,2).toLowerCase();
				else
					// Default language is English
					sLang="en";
			}
			sMsg=_validation.messages[sLang][this.type];
			
			if(oElement.getAttribute("min"))
			{
				sMsg += ", minimum value: " + oElement.getAttribute("min");
			}
			if(oElement.getAttribute("max"))
			{
				sMsg += ", maximum value: " + oElement.getAttribute("max");
			}

			alert(((sNym)?sNym+": ":"")+sMsg+((this.format)
				?" "+this.format.reformat(sLang,this.type):""));
		}

		// Perform onvalidatefocus event handler for invalid field
		if(oElement.onvalidatefocus)
			oElement.onvalidatefocus();

		// Give invalid field focus
		oElement.focus();
		// Clear the Err object
		this.clear();
	}
}

/*====================================================================
Function: Validation
Purpose:  Custom object constructor.
Inputs:   None
Returns:  undefined
Changes: 08/05/00  IW	Changed Invalid Date message (SYS0198)
		 08/06/00  MC	Added MinMax validation for AMOUNT fields (SYS0866)
		 27/07/2006 MH	EP1040 Added date validation to stop years < 1753 being added
====================================================================*/
function Validation(){
	// Define global constants for calls to error message arrays
	this.REQUIRED = 0;
	this.INTEGER  = 1;
	this.FLOAT    = 2;
	this.DATE     = 3;
	this.AMOUNT   = 4;
	this.MASK     = 5;
	this.PHONE	  = 6;
	// EP1040
	this.YEARERROR = 7;
	var validateYear=null;
	
	// Create error message dictionary
	this.messages = new Array;

	// Prototype the date tokens for each language
	Array.prototype.MM = new String;
	Array.prototype.DD = new String;
	Array.prototype.YYYY = new String;

	//English
	this.messages["en"]=new Array(
		"Please enter a value",
		"Please enter a valid integer",
		"Please enter a valid floating point",
		"The Date you have entered is invalid. Please enter a valid date in the format ",
		"Please enter a valid monetary amount",
		"Please enter a value in the form of ",
		"Invalid character/s. Please enter numeric characters only",
		"The year entered is outside the expected range. Please correct");
		with(this.messages["en"]){
			MM="MM";
			DD="DD";
			YYYY="YYYY";
		}
	

	// AS 30/11/06 EP1251
	this.setDefault=setDefault;
	this.IntMinMaxTest=IntMinMaxTest;
	this.FloatMinMaxTest=FloatMinMaxTest;
	this.isDate=isDate;
	this.isNum=isNum;
	this.setup=setup;


	///////////////////////////////////////////////////////////////////////////
	// AS 30/11/06 EP1251 Start. Remove memory leaks due to function closures.

	///////////////////////////////////////////////////////////////////////////
	//Method:   Validation.setDefault
	//Purpose:  Set value for variable v if v is zero, empty string or
	//          undefined
	//Inputs:   v - variable (passed by value)
	//          d - default value
	//Returns:  v or d
	///////////////////////////////////////////////////////////////////////////
	function setDefault(v, d)
	{
		return (v)?v:d;
	}
	
	///////////////////////////////////////////////////////////////////////////
	//Method:   Validation.IntMinMaxTest
	//Purpose:  Check that integer value is a within a given range
	//Inputs:   oElement - form element
	//          sFormat  - string format
	//Returns:  boolean
	///////////////////////////////////////////////////////////////////////////
	function IntMinMaxTest(v)
	{
		var retval = true;
		if(v.getAttribute("min"))
		{
			retval = parseInt(v.value) >= parseInt(v.getAttribute("min")) && parseInt(v.value) <= parseInt(v.getAttribute("max"));
		}
		return retval;
	}
	
	///////////////////////////////////////////////////////////////////////////
	//Method:   Validation.FloatMinMaxTest
	//Purpose:  Check that integer value is a within a given range
	//Inputs:   oElement - form element
	//          sFormat  - string format
	//Returns:  boolean
	///////////////////////////////////////////////////////////////////////////
	function FloatMinMaxTest(v)
	{
		var retval = true;
		if(v.getAttribute("min"))
		{
			retval = parseFloat(v.value) >= parseFloat(v.getAttribute("min")) && parseFloat(v.value) <= parseFloat(v.getAttribute("max"));
		}
		return retval;
	}
	
	///////////////////////////////////////////////////////////////////////////
	//Method:   Validation.isDate
	//Purpose:  Check that value is a date of the correct format
	//Inputs:   oElement - form element
	//          sFormat  - string format
	//Returns:  boolean
	//Changes:  EP1040 Added code to populate variable validateYear
	///////////////////////////////////////////////////////////////////////////
	function isDate(oElement,sFormat)
	{
		var sDate=oElement.value;
		var aDaysInMonth=new Array(31,28,31,30,31,30,31,31,30,31,30,31);

		// Fetch the date separator from the user's input
		var sSepDate=sDate.charAt(sDate.search(/\D/));
		// Fetch the date separator from the format
		var sSepFormat=sFormat.charAt(sFormat.search(/[^MDY]/i));
		// Compare separators
		if (sSepDate!=sSepFormat)
			return false;

		// Fetch the three pieces of the date from the user's input and the format
		var aValueMDY=sDate.split(sSepDate,3);
		var aFormatMDY=sFormat.split(sSepFormat,3);
		var iMonth,iDay,iYear;

		if (!aValueMDY[0] || !aValueMDY[1] || !aValueMDY[2])
			return false;
		// Validate that all pieces of the date are numbers
		if (!_validation.isNum(aValueMDY[0])
			||!_validation.isNum(aValueMDY[1])
			||!_validation.isNum(aValueMDY[2]))
			return false;

		// Assign day, month, year based on format
		switch (aFormatMDY[0].toUpperCase()){
			case "YYYY" :
				iYear=aValueMDY[0];
				break;
			case "DD" :
				iDay=aValueMDY[0];
				break;
			case "MM" :
				iMonth=aValueMDY[0];
				break;
			default :
				return false;
		}
		switch (aFormatMDY[1].toUpperCase()){
			case "YYYY" :
				iYear=aValueMDY[1];
				break;
			case "MM" :
				iMonth=aValueMDY[1];
				break;
			case "DD" :
				iDay=aValueMDY[1];
				break;
			default :
				return false;
		}
		switch(aFormatMDY[2].toUpperCase()){
			case "MM" :
				iMonth=aValueMDY[2];
				break;
			case "DD" :
				iDay=aValueMDY[2];
				break;
			case "YYYY" :
				iYear=aValueMDY[2];
				break;
			default :
				return false;
		}

		// Require 2 digit month and day
		if(iDay.length!=2)
			return false;
		if(iMonth.length!=2)
			return false;


		// Require 4 digit year
		if(oElement.form.getAttribute("year4")!=null && iYear.length!=4)
			return false;
		// Process pivot date and update field
		var iPivot=_validation.setDefault(oElement.getAttribute("pivot"),
			oElement.form.getAttribute("pivot"));
		if(iPivot && iPivot.length==2 && iYear.length==2){
			iYear=((iYear>iPivot)?19:20).toString()+iYear;
			var sValue=aFormatMDY.join(sSepFormat).replace(/MM/i,iMonth);
			sValue=sValue.replace(/DD/i,iDay).replace(/YYYY/i,iYear);
			oElement.value=sValue;
		}

		// Check for leap year
		var iDaysInMonth=(iMonth!=2)?aDaysInMonth[iMonth-1]:
			((iYear%4==0 && iYear%100!=0 || iYear % 400==0)?29:28);
		//EP1040
		validateYear=iYear;
		
		return (iDay!=null && iMonth!=null && iYear!=null
				&& iMonth<13 && iMonth>0 && iDay>0 && iDay<=iDaysInMonth);
	}
	
	///////////////////////////////////////////////////////////////////////////
	//Method:   Validation.isNum
	//Purpose:  Check that parameter is a number
	//Inputs:   v - string value
	//Returns:  boolean
	///////////////////////////////////////////////////////////////////////////
	function isNum(v)
	{
		return (v.toString() && !/\D/.test(v));
	}

	///////////////////////////////////////////////////////////////////////////
	//Method:   Validation.setup
	//Purpose:  Set up methods and event handlers for all forms and elements
	//Inputs:   none
	//Returns:  undefined
	///////////////////////////////////////////////////////////////////////////
	function setup()
	{
		// Fan through forms on page to perform initializations
		var i,iForms=document.forms.length;
		for(i=0; i<iForms; i++){
			var oForm=document.forms[i];
			if(!oForm.bProcessed){
				///////////////////////////////////////////////
				//Method:   Form.markRequired
				//Purpose:  Mark all required fields for a form
				//Inputs:   none
				//Returns:  undefined
				///////////////////////////////////////////////
				
				// AS 30/11/06 EP1251
				oForm.markRequired=markRequired;
				
				var sValidateWhen=oForm.getAttribute("validate");
				if (sValidateWhen!=null){
					//
					// Capture and replace onreset and onsubmit event handlers
					//
					oForm.fSubmit=oForm.onsubmit;
					oForm.fReset=oForm.onreset;

					// Create new event handlers
					
					// AS 30/11/06 EP1251
					oForm.onsubmit=onSubmit;
					
					// AS 30/11/06 EP1251
					oForm.onreset=onReset;
				}
				oForm.bProcessed=true;
			}
			// Create Input methods
			var j, iElements=oForm.elements.length;
			for(j=0; j<iElements; j++){
				var oElement=oForm.elements[j];
				if(!oElement.bProcessed) {

					// All event handlers are presumed to be strings/functions
					// at parse-time and assigned only as functions at run-time.

					// Create custom onvalidate event handlers
					var vOnValidate=oElement.getAttribute("onvalidate");
					if(vOnValidate){
						if(typeof vOnValidate!="function")
							oElement.onvalidate=new Function(vOnValidate);
						else
							oElement.onvalidate=vOnValidate;
					}
					// Create custom handler for onvalidatefocus event
					var vOnValidateFocus=oElement.getAttribute("onvalidatefocus");
					if(vOnValidateFocus){
						if(typeof vOnValidateFocus!="function")
							oElement.onvalidatefocus=new Function(vOnValidateFocus);
						else
							oElement.onvalidatefocus=vOnValidateFocus;
					}
					// Custom onmark event handler
					var vOnMark=oElement.getAttribute("onmark");
					if(vOnMark){
						if(typeof vOnMark!="function")
							oElement.onmark=new Function(vOnMark);
						else
							oElement.onmark=vOnMark;
					}
					// Custom onkeypress filtering for text fields
					if(oElement.onkeypress)
						oElement.fKeypress=oElement.onkeypress;
						
					// AS 30/11/06 EP1251
					oElement.onkeypress=onKeyPress;
					
					// Custom onchange validation
					if(sValidateWhen=="onchange") {
						// Capture and replace onchange event handlers
						if(oElement.onchange)
							oElement.fChange=oElement.onchange;
							
						// AS 30/11/06 EP1251
						oElement.onchange=onChange;
					}

					// AS 30/11/06 EP1251
					oElement.paint=paint;				
					oElement.restore=restore;
					oElement.valid=valid;

					this.bProcessed=true;
				}
			}
		}
	}

	function onKeyPress()
	{
		if(this.fKeypress && this.fKeypress()==false)
			event.returnValue=false;
		m_bIsChanged = true;
		var sFilter=this.getAttribute("filter");
		if(!sFilter) sFilter = "[-.,0-9a-zA-Z'/& ]";
		if(this.getAttribute("wildcard"))
			sFilter = sFilter.substr(0,1) + "*" + sFilter.substr(1);
		if(sFilter){
			var sKey=String.fromCharCode(event.keyCode);
			var re=new RegExp(sFilter);
			// Do not filter out ENTER!
			if(sKey!="\r" && !re.test(sKey))
				event.returnValue=false;
			if(this.getAttribute("upper"))
				sKey = sKey.toUpperCase();
			event.keyCode=sKey.charCodeAt(0);

		}
	}
	
	function onChange()
	{
		this.restore();
		if(!this.valid()){
			_err.raise();
			event.returnValue=false;
		}
		if(this.fChange && this.fChange()==false){
			event.returnValue=false;
		}
		FlagChange(true);
	}
	
	////////////////////////////////////
	//Method:   Input.paint
	//Purpose:  Change style of element
	//Inputs:   none
	//Returns:  undefined
	////////////////////////////////////
	function paint()
	{
		var sColor = _validation.setDefault(
			this.getAttribute("invalidColor"),
			this.form.getAttribute("invalidColor") );
		if (!sColor){
			// Paint element by changing class
			this.setAttribute("oldClass", this.className);
			this.className = "invalid";
		}else{
			// Paint element by changing color directly
			this.setAttribute("bg", this.style.backgroundColor);
			this.style.backgroundColor = sColor;
		}
	}
	
	/////////////////////////////////////////////
	//Method:   Input.restore
	//Purpose:  Restore element to original style
	//Inputs:   none
	//Returns:  undefined
	/////////////////////////////////////////////
	function restore() 
	{
		var sBG=this.getAttribute("bg");
		if (sBG!=null) {
			// Revert to previous background color
			this.style.backgroundColor = sBG;
			this.removeAttribute("bg");
		}else{
			var sOldClass=this.getAttribute("oldClass");
			if (sOldClass!=null){
				// Revert to previous class
				this.className=sOldClass;
				this.removeAttribute("oldClass");
			}
		}
	}

	/////////////////////////////////////////////
	//Method:   Input.valid
	//Purpose:  Validate an element based on the
	//          attributes provided in the HTML text
	//Inputs:   none
	//Returns:  boolean
	/////////////////////////////////////////////
	function valid()
	{
		var sType=this.type;
		
		// APS 24/02/00 - We do not want to validate against hideen fields
		if(this.style.getAttribute("visibility")=="hidden"){							
			return true;
		}
		
		if(sType=="text" || sType=="textarea" || sType=="file"){
			// Trim leading and trailing spaces
			if(this.form.getAttribute("notrim")==null)
				this.value = this.value.trim();
			// Remove any Server Side Include text
			if(this.form.getAttribute("ssi")==null){
				while (this.value.search("<!-"+"-#") > -1)
					this.value = this.value.replace("<"+"!--#", "<"+"!--");
			}
		}
		// REQUIRED
		if(this.getAttribute("required")!=null && this.disabled == false && !this.value){
			_err.add(this, _validation.REQUIRED, null);
			return false;
		}
		// FLOAT
		var sFloatDelimiter=this.getAttribute("float");
		var bSigned=this.getAttribute("signed")!=null;
		if (sFloatDelimiter!=null && this.value){
			// Assign default value to delimiter
			sFloatDelimiter=(sFloatDelimiter==",")?",":"\\.";
			var re=new RegExp("^("+((bSigned)?"[\\-\\+]?":"")+"(\\d*"+sFloatDelimiter+"?\\d+)|(\\d+"+sFloatDelimiter+"?\\d*))$");
			if (!re.test(this.value)){
				_err.add(this, _validation.FLOAT, null);
				return false;
			}
		}
		// AMOUNT
		var sAmtDelimiter = this.getAttribute("amount");
		if (sAmtDelimiter!=null && this.value){
			// Assign default value to delimiter
			sAmtDelimiter=(sAmtDelimiter==",")?",":"\\.";
			var re = new RegExp("^"+((bSigned)?"[\\-\\+]?":"")+"\\d+("+sAmtDelimiter+"\\d{2})?$");
			if(!re.test(this.value)|| !_validation.FloatMinMaxTest(this)){
				_err.add(this, _validation.AMOUNT, null);
				return false;
			}
		}
		// INTEGER
		if (this.getAttribute("integer")!=null && this.value){
			var re=new RegExp("^"+((bSigned)?"[\\-\\+]?":"")+"\\d+$");
			if (!re.test(this.value) || !_validation.IntMinMaxTest(this)){
				_err.add(this, _validation.INTEGER, null);
				return false;
			}
		}
		// PHONE  ASu BMIDS00106 - Start
		var sFormat=this.getAttribute("phone");
		if (sFormat!=null && this.value)
		{
			//var re=new RegExp(sFormat);
			if (isNaN(this.value)== true)
			{
				_err.add(this, _validation.PHONE, null);
				return false;
			}
		}						
		// ASu BMIDS00106 - End 
		// DATE
		var sFormat=this.getAttribute("date");
		if (sFormat!=null && this.value) {
			// Set default date format
			sFormat = _validation.setDefault(sFormat, "MM/DD/YYYY");
			if (!_validation.isDate(this,sFormat)){
				_err.add(this, _validation.DATE, sFormat.toUpperCase());
				return false;
			}
			//EP1040
			if (validateYear<1753){
				_err.add(this, _validation.YEARERROR, null);
				return false;
			}
		}
		// MASK
		var sMask=this.getAttribute("mask");
		if(sMask && this.value){
			var sPattern=sMask.replace(
				/(\$|\^|\*|\(|\)|\+|\.|\?|\\|\{|\}|\||\[|\])/g,"\\$1");
			sPattern=sPattern.replace(/9/g ,"\\d");
			sPattern=sPattern.replace(/x/ig,".");
			sPattern=sPattern.replace(/z/ig,"\\d?");
			sPattern=sPattern.replace(/a/ig,"[A-Za-z]");
			var re=new RegExp("^"+sPattern+"$");
			if(!re.test(this.value)){
				_err.add(this, _validation.MASK, sMask);
				return false;
			}
		}
		// REGEXP
		var sRegexp=this.getAttribute("regexp");
		if(sRegexp && this.value){
			var re=new RegExp(sRegexp);
			if(!re.test(this.value)){
				_err.add(this, _validation.MASK, sRegexp);
				return false;
			}
		}
		// AND
		var sAnd=this.getAttribute("and");
		if(sAnd && this.value){
			var aAnd = sAnd.split(/,/);
			var i, iFields=aAnd.length;
			// Require each element in the list if this element is valued
			for(i=0; i<iFields; i++){
				var oNewElement=this.form.elements[aAnd[i]];
				if(oNewElement && oNewElement.value.trim()==""){
					_err.add(oNewElement, _validation.REQUIRED, null);
					return false;
				}
			}
		}
		// OR
		var sOr=this.getAttribute("or");
		if(sOr && this.value==""){
			var aOr=sOr.split(/,/);
			var i, iFields=aOr.length;
			var oNewElement, bAccum=false;
			for(i=0; i<iFields; i++){
				oNewElement=this.form.elements[aOr[i]];
				if(oNewElement)
					bAccum = bAccum || oNewElement.value.trim();
			}
			if(!bAccum){
				_err.add(this, _validation.REQUIRED, null);
				return false;
			}
		}
		// NOSPACE
		if (this.getAttribute("nospace")!=null)
			this.value=this.value.replace(/\s/g,"");
		// UPPERCASE
		if (this.getAttribute("uppercase")!=null)
			this.value=this.value.toUpperCase();
		// LOWERCASE
		if (this.getAttribute("lowercase")!=null)
			this.value=this.value.toLowerCase();
		// Capitalize - SYS5729 Added by TW 28/10/2002
		if (this.getAttribute("capitalize")!=null)
		{
			var words = this.value.split(" ");
			var i, iFields = words.length;
			for(i=0; i<iFields; i++)
			{
				words[i] = words[i].substring(0, 1).toUpperCase() + words[i].substring(1);
			}
			this.value=words.join(" ");
		}
		// Perform onvalidate event handler
		if(this.onvalidate && this.onvalidate()==false)
			return false;

		return true;
	}

	function markRequired()
	{		
		var i, iElements=this.elements.length;
		var sMarkHTML, sMarkWhere;
		for(i=0; i<iElements; i++){
			var oElement=this.elements[i];
			// TW SYS5729 Added to ensure textTransform is carried through into field
			var elementAttribute=oElement.getAttribute("style").textTransform;
			if(elementAttribute!="")
			{
				oElement.setAttribute(elementAttribute, true);
			}
			// Perform onmark event handler
			if(oElement.onmark && oElement.onmark()==false)
				continue;
			if(oElement.getAttribute("required")!=null){
				sMarkHTML=this.getAttribute("insert");
				sMarkWhere=this.getAttribute("mark");
				if(sMarkHTML){
					switch(sMarkWhere.toLowerCase()){
						case "before" :
							sMarkWhere="beforeBegin";
							break;
						default :
							sMarkWhere="afterEnd";
					}
					oElement.insertAdjacentHTML(sMarkWhere,sMarkHTML);
				}else{
					var sClassName=oElement.className;
					if(sClassName!="required"){
						oElement.setAttribute("nonreqClass",oElement.className);
						oElement.parentElement.parentElement.style.color="red";
//									oElement.parentElement.parentElement.style.visibility = "hidden";
//									oElement.parentElement.parentElement.style.visibility = "visible";
					}else{
						oElement.className=_validation.setDefault(oElement.getAttribute("nonreqClass"),oElement.className);
						oElement.removeAttribute("nonreqClass");
					}
				}
			}
		}
	}

	function onSubmit()
	{
		var i, oElement, iElements=this.elements.length;
		event.returnValue=true;
		// Restore all elements to original style
		for (i=0; i<iElements; i++)
			this.elements[i].restore();
		// Validate individual elements
		for(i=0;i<iElements;i++){
			oElement=this.elements[i];
			// Perform default validation for element
			if (!oElement.valid()){
				_err.raise();
				event.returnValue=false;
				return false;
			}
		}

		// Perform original onsubmit event handler
		if (this.fSubmit && this.fSubmit()==false){
			event.returnValue=false;
			return false;
		}

		// Insert default values just before submit
		var vDefault;
		for(i=0;i<iElements;i++){
			oElement=this.elements[i];
			vDefault=oElement.getAttribute("default");
			if(vDefault && !oElement.value)
				oElement.value=vDefault;
		}
		return true;
	}
	
	function onReset()
	{
		var i, iElements=this.elements.length;
		for (i=0; i<iElements; i++)
			this.elements[i].restore();
		// Perform original event handler if present
		if (this.fReset && this.fReset()==false)
				event.returnValue=false;
	}
	// AS 30/11/06 EP1251 End
	///////////////////////////////////////////////////////////////////////////
}

/*====================================================================
Function: Validation_Init
Purpose:  Sets up validation routines and captures event handlers that
          exist on the form already for the onsubmit and onreset events.
          Creates methods on FORMs and INPUTs.
====================================================================*/
function Validation_Init(){
	// Limit use of script to valid environments
	if("".replace && document.body && document.body.getAttribute){

		String.prototype.trim=trim;
		String.prototype.reformat=reformat;

		// Form setup
		if(document.forms){
			// Create custom objects
			_validation=new Validation;
			_err=new Err;

			// Process forms and elements
			_validation.setup();

			var i, iForms=document.forms.length;
			for(i=0; i<iForms; i++){
				var oForm=document.forms[i];
				if(oForm.getAttribute("mark")!=null)
					oForm.markRequired();
			}
		}
	}
	
	///////////////////////////////////////////////////////////////
	//Method:   String.trim
	//Purpose:  Removing leading and trailing spaces
	//Inputs:   none
	//Returns:  string
	///////////////////////////////////////////////////////////////
	function trim()
	{
		return this.replace(/^\s+/,"").replace(/\s+$/,"");
	}

	///////////////////////////////////////////////////////////////
	//Method:   String.reformat
	//Purpose:  Translate the date format into the correct language
	//Inputs:   sLang   - language of error message to display
	//					iType   - type of failed validation
	//Returns:  string
	////////////////////////////////////////////////////////////////
	function reformat(sLang,iType)
	{
		var sString=this.valueOf();
		if (iType==_validation.DATE && _validation.messages[sLang]) {
			sString=sString.replace(/MM/,_validation.messages[sLang].MM);
			sString=sString.replace(/DD/,_validation.messages[sLang].DD);
			sString=sString.replace(/YYYY/,_validation.messages[sLang].YYYY);
		}
		return sString;
	}
}

var m_bIsChanged=false;

function IsChanged()
{
	return m_bIsChanged;
}
function FlagChange(bIsChange)
{
	if(bIsChange)
		m_bIsChanged = true;
}

// Removed by TW 26/7/2002 - This routine was being executed twice in most forms
//Validation_Init();

// Added by TW 15/10/2002 - SYS5588 To inhibit mouse right-click
function document.oncontextmenu() { return false; }

// Added by TW 22/10/2002 - SYS5043
document.onkeydown = function()
{
// This function inhibits the use of function keys F2 to F12 and 
// any alt + key combination outside the range alt + A to Alt + Z

	var x = window.event.keyCode;
	var retVal = true;

	if(window.event.altKey)
		{
		if ((x < 65) || (x > 90)) 
			{
// Disallow Alt keys outside the range Alt + A to Alt + Z
			retVal = false;
			}
		}
	else
		{
		
		// MAR273 disable BACKSPACE to navigate to prev page
		if (window.event && x==8)
			{
			var id = window.event.srcElement.id; 
			var cl = window.event.srcElement.className;
			//MAR614, added msgTxtUpper class here
			//if (cl != 'msgTxt' && cl != 'msgInput' ) 
			if (cl != 'msgTxt' && cl != 'msgInput' && cl != 'msgTxtUpper' ) 
				{
// Disallow Backspace
				retVal = false;
// Override the key code to prevent the browser handling the key anyway
				window.event.keyCode = 500 + (x - 111);
				}
			}
			
		if ((x > 112) && (x < 124))
			{
// Disallow Function Keys F2 to F12
			retVal = false;

// Override the key code to prevent the browser handling the key anyway
			window.event.keyCode = 500 + (x - 111);
			}
		}
      	window.event.returnValue = retVal;
}

// Added by TW 4/11/2002 to 
// 	1. Provide a common trim function for JScript in ASPs
//	2. Provide a correction for SYS5759 (Coding error in CR044)
//	3. Correct a coding error in DC225 

/*********************************************************************
Method:   String.trim
Purpose:  Removing leading and trailing spaces
Inputs:   none
Returns:  string
*********************************************************************/
String.prototype.trim=function (){
	return this.replace(/^\s+/,"").replace(/\s+$/,"");
}
