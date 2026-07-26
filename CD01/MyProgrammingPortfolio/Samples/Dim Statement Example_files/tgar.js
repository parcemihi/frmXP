var ToolBar_Supported = false;
var Frame_Supported   = false;
var DoInstrumentation = false;
var doImage = doImage;
var TType = TType;

if (navigator.userAgent.indexOf("MSIE")    != -1 && 
	navigator.userAgent.indexOf("Windows") != -1 && 
	navigator.appVersion.substring(0,1) > 3)
{
	ToolBar_Supported = true;
}

if(doImage == null)
{
	var a= new Array();
	a[0] = prepTrackingString(window.location.hostname,7);
	if (TType == null)
	{	
		a[1] = prepTrackingString('PV',8);
	}
	else
	{
		a[1] = prepTrackingString(TType,8);
	}
	a[2] = prepTrackingString(window.location.pathname,0);
	if(window.location.search != '')
	{
		a[a.length] = prepTrackingString(window.location.search,1);
	}
	if( '' != window.document.referrer)
	{
		a[a.length] = prepTrackingString(window.document.referrer,5);
	}
	
	if (navigator.userAgent.indexOf("SunOS") == -1 && navigator.userAgent.indexOf("Linux") == -1)
	{
		buildIMG(a);
	}
}	

function buildIMG(pArr)
{
	var TG = '<LAYER visibility="hide"><div style="display:none;"><IMG src="' + location.protocol + '//c.microsoft.com/trans_pixel.asp?';
	for(var i=0; i<pArr.length; i++)
	{
		if(0 == i)
		{
			TG +=  pArr[i];
		}
		else
		{
			TG += '&' + pArr[i];
		}
	}
	TG +='" height="0" width="0" hspace="0" vspace="0" Border="0"></div></layer>';

	document.write(TG);
}
function prepTrackingString(ts, type)
{
	var rArray;
	var rString;
	var pName = '';
	if (0 == type)
	{
		pName = 'p=';
		rString = ts.substring(1);
		rArray = rString.split('/');
	}
	if (1 == type)
	{
		pName = 'qs=';
		rString = ts.substring(1);
		rArray = rString.split('&');		
	}
	if (2 == type)
	{
		pName = 'f=';
		rString = escape(ts);
		return pName + rString;
	}
	if (3 == type)
	{
		pName = 'tPage=';
		rString = escape(ts);
		return pName+rString;
	}
	if (4 == type)
	{
		pName = 'sPage=';
		rString = escape(ts);
		return pName + rString;
	}
	if (5 == type)
	{
		pName = 'r=';
		rString = escape(ts);
		return pName + rString;
	}
	if (6 == type)
	{
		pName = 'MSID=';
		rString = escape(ts);
		return pName + rString;
	}
	if (7 == type)
	{
		pName = 'source=';
		rString = ts.toLowerCase();
		if(rString.indexOf("microsoft.com") != -1)
		{
			rString = rString.substring(0,rString.indexOf("microsoft.com"));
			if('' == rString)
			{
				rString = "www";
			}	
			else
			{
				rString = rString.substring(0,rString.length -1);
			}
		}
		return pName + rString;
	}
	if (8 == type)
	{
		pName = 'TYPE=';
		rString = escape(ts);
		return pName + rString;
	}
	rString = '';
	if(null != rArray)
	{
		if(0 == type)
		{
			for( j=0; j < rArray.length - 1; j++)
			{	
				rString += rArray[j] + '_';  
			} 
		}
		else
		{
			for( j=0; j < rArray.length  ; j++)
			{
				rString += rArray[j] + '_';  
			} 
		}
	}
	rString = rString.substring(0, rString.length - 1);  	 	
	return pName + rString;
}