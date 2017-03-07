// Flash Detection / Redirect (cookie variant)  v1.1.0
// http://www.dithered.com/javascript/flash_cookie/index.html
// code by Chris Nott (chris@NOSPAMdithered.com - remove NOSPAM)


var dontKnow = false;
var flashVersion = 0;

// Retrieve flash cookie
var cookieStart = document.cookie.indexOf('flash');
if (cookieStart != -1) {
	var cookieEnd = document.cookie.indexOf(';', cookieStart);
	if (cookieEnd == -1) cookieEnd = document.cookie.length;
	flashVersion = document.cookie.substring(cookieStart + 6, cookieEnd); 
}

// If the cookie doesn't exist...
else {
	
	// use flash_detect.js to return the Flash version
	flashVersion = getFlashVersion();
	
	// write the version information to a cookie
	document.cookie = 'flash=' + flashVersion;
}

// For the situation where we can't detect, set the values of the reference variables
if (flashVersion == flashVersion_DONTKNOW) {
	flashVersion = 0;
	dontKnow = true;
}
