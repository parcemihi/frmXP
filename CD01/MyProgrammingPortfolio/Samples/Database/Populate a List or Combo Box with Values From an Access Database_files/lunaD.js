function w(p1) { document.writeln(p1); }

browserType = (navigator.appName.substring(0,8) + parseFloat(navigator.appVersion)).toLowerCase();
browserClass = navigator.appName.substring(0,8) + parseFloat(navigator.appVersion.substring(0,1));
var confirmed = 1;
var vpn = "000195";

var rand = Math.random() * 10000000000000000;
var j = '<a href="http://ad.doubleclick.net/jump';
var i = '<img src="http://ad.doubleclick.net/ad';
var f = '<iframe src="http://ad.doubleclick.net/adi';

var sprinks_link1 = 'http://sprinks.about.com?PM=999';
var sprinks_title1 = 'Be the first Sponsored Links advertiser on this site...';
var sprinks_descr1 = 'Start driving targeted traffic to your site with Sprinks, About\'s Sponsored Links program.';
var sprinks_displink1 = 'http://sprinks.about.com/';
var sprinks_cpc1 = '$0.00';
var sprinks_link2 = 'http://sprinks.about.com?PM=999';
var sprinks_title2 = 'Be a Sponsored Links advertiser on this site...';
var sprinks_descr2 = 'Start driving targeted traffic to your site with Sprinks, About\'s Sponsored Links program.';
var sprinks_displink2 = 'http://sprinks.about.com/';
var sprinks_cpc2 = '$0.00';
var texisplaceholder="";
var sprinks_title1="Find Practically Anything on Earth at eBay!";
var sprinks_link1="http://sprinks-clicks.about.com/?ref=vpn000195&bid=1389379&lnk=http://cgi.ebay.com/aw%2Dcgi/eBayISAPI.dll%3FRedirectEnter%26partner%3D14289%26loc%3Dhttp://pages.ebay.com/computer%2Dindex.html";
var sprinks_descr1="Ebay is the world's largest online trading site where people buy, sell or trade anything on earth. There are more than 4 million items available in more than 4,000 categories! Discover the fun!";
var sprinks_displink1="http://cgi.ebay.com/aw-cgi/eBayISAPI.dll";
var sprinks_cpc1="$0.30";
var sprinks_title2="E-Labeling LLC: Visual Basic label programming";
var sprinks_link2="http://sprinks-clicks.about.com/?ref=vpn000195&bid=1216123&lnk=http://www.elabeling.com/";
var sprinks_descr2="Online sales of Zebra and Eltron direct thermal and thermal transfer tag, receipt and label printers, barcode labels, thermal transfer wax and resin ribbons and bar code labeling design software";
var sprinks_displink2="http://www.elabeling.com/";
var sprinks_cpc2="$0.17";



var chan='compute';
var guidesite='visualbasic';
var n_guidesite='visualbasic';
var nlink1='<A HREF="http://home.about.com/compute/newsletters.htm?IAM=vpn000195_3" TARGET="_top"><FONT COLOR="#FFFFFF" FACE="Verdana, Geneva, Helvetica" SIZE=1><B>Free Newsletters</B></font></a>';
var nlink2='<A HREF="http://howto.about.com/compute/index.htm?IAM=vpn000195_4" TARGET="_top"><FONT COLOR="#FFFFFF" FACE="Verdana, Geneva, Helvetica" SIZE=1><B>How To\'s</B></font></a>';
var nlink3='<A HREF="http://home.about.com/compute/index.htm?IAM=vpn000195_5" TARGET="_top"><FONT COLOR="#FFFFFF" FACE="Verdana, Geneva, Helvetica" SIZE=1><B>About Computing</B></font></a>';
var nlink4='<A HREF="http://home.about.com/index.htm?IAM=vpn000195_6" TARGET="_top"><FONT COLOR="#FFFFFF" FACE="Verdana, Geneva, Helvetica" SIZE="1"><B>About&reg;</B></font></a>';
var tgt='_top';
var top='<SCRIPT>navbar();<\/SCRIPT>';
var topad1='<SCRIPT>adslot1();<\/SCRIPT>';
var topad2='<SCRIPT>pw();<\/SCRIPT>';
var column='<SCRIPT>sqslot1(); vmslot1();<\/SCRIPT>';
var botspr='<SCRIPT>sprinksbottom();<\/SCRIPT>';
var botad1='<SCRIPT>adslot2();<\/SCRIPT>';
var botad2='<SCRIPT>pw();<\/SCRIPT>';
var bottom='<SCRIPT>popslot1();<\/SCRIPT>';
var sprinksboxtitle='Visual Basic';
var aux = '/vpn'+vpn+'.about.com/'+chan+'_'+guidesite+';svc=;site='+guidesite+';kw=;chan='+chan+';syn=vpn'+vpn+';pos=';
function adslot1() {

 w('<table border="0" cellpadding="0" cellspacing="0"><tr><td>');
 	   	if (browserType.indexOf("netscape4.0")!= -1 || browserType.indexOf("webtv")!= -1) {
			w(j+aux+'slot1;sz=468x60;ord=1'+rand+'" target="'+tgt+'">');
			w(i+aux+'slot1;sz=468x60;ord=1'+rand+'" height=60 width=468 border=0></a>');
		}
		else{
			w('<nolayer>');
			w(f+aux+'slot1;sz=468x60;ord=1'+rand+'" target="'+tgt+'" height=60 width=468 border=0 marginwidth=0 marginheight=0 hspace=0 vspace=0 frameborder=0 scrolling=no></a>');
			if (browserClass != 'Netscape2') {
				w(j+aux+'slot1;sz=468x60;ord=1'+rand+'" target="'+tgt+'">');
				w(i+aux+'slot1;sz=468x60;ord=1'+rand+'" height=60 width=468 border=0></a>');
			}
			w('</iframe>');
			w('</nolayer>');
			w('<ilayer id="adslot1" visibility=hide width=468 height=60></ilayer>');
		}

w('</td></tr>');
w('<tr><td>');
w('<DIV ALIGN="center"><FONT FACE="verdana, geneva, helvetica" SIZE="1" COLOR="#999999" STYLE="font: Verdana; font-size: 9px">please support our sponsors</font></DIV>');
w('</td></tr></table>');

}
cookiePath = ";Path=/";
var interValue;
function dartCall() {
	if (browserType.indexOf("netscape4.0")!= -1 || browserType.indexOf("webtv")!= -1) {
		w(j+aux+'popslot1;sz=1x1;ord=1'+rand+'" target="'+tgt+'">');
		w(i+aux+'popslot1;sz=1x1;ord=1'+rand+'" height=1 width=1 border=0></a>');
	}
	else {
			w('<nolayer>');
			w(f+aux+'popslot1;sz=1x1;ord=1'+rand+'" target="'+tgt+'" height=1 width=1 border=0 marginwidth=0 marginheight=0 hspace=0 vspace=0 frameborder=0 scrolling=no></a>');
			if (browserClass != "Netscape2") {
				w(j+aux+'popslot1;sz=1x1;ord=1'+rand+'" target="'+tgt+'">');
				w(i+aux+'popslot1;sz=1x1;ord=1'+rand+'" height=1 width=1 border=0></a>');
			}
			w('</iframe>');
			w('</nolayer>');
			w('<ilayer id="popslot1" visibility=hide width=1 height=1></ilayer>');
	}
}
			
			
function checkCookie() {
	if ((parseInt(navigator.appVersion) < 4 )||(navigator.appName.substring(0,8).indexOf("webtv")!= -1 )){
			return;
	}
		
	cookieExpDays = 1;		
		var expDate = new Date();
		expDate.setTime(expDate.getTime() + (86400000 * cookieExpDays));
		cookieDate = ";Expires=" + expDate.toGMTString();
	
		if(document.cookie) {
			var bigCookie = document.cookie;
			var interPos = bigCookie.indexOf("pixelpopslot=");
			
			if (interPos != -1) {
				var interStart = interPos + 13;
				var interEnd = bigCookie.indexOf(";", interStart);
				if (interEnd == -1) {
					interEnd = bigCookie.length;
				}
				var interValue = bigCookie.substring(interStart, interEnd);
				interCount = parseInt(interValue);
				
				if (interCount <= 1) {
					interCount++;
					document.cookie="pixelpopslot="+interCount+";"+cookiePath+cookieDate;
					dartCall();
				} 
				else {
					return;
				}
			} 
			else {
				document.cookie="pixelpopslot=1;"+cookiePath+cookieDate;
			}
		}
	}
	
function popslot1() {
	checkCookie();
}

function adslot2()
{
w('<table border="0" cellpadding="0" cellspacing="0"><tr><td>');
				
if (browserType.indexOf("netscape4.0")!= -1 || browserType.indexOf("webtv")!= -1) {
	w(j+aux+'slot1;sz=468x60;ord=1'+rand+'" target="'+tgt+'">');
	w(i+aux+'slot1;sz=468x60;ord=1'+rand+'" height=60 width=468 border=0></a>');
		}
	
	else{
			w('<nolayer>');
			w(f+aux+'slot1;sz=468x60;ord=1'+rand+'" target="'+tgt+'" height=60 width=468 border=0 marginwidth=0 marginheight=0 hspace=0 vspace=0 frameborder=0 scrolling=no></a>');
			if (browserClass != 'Netscape2') {
				w(j+aux+'slot1;sz=468x60;ord=1'+rand+'" target="'+tgt+'">');
				w(i+aux+'slot1;sz=468x60;ord=1'+rand+'" height=60 width=468 border=0></a>');
			}
			w('</iframe>');
			w('</nolayer>');
			w('<ilayer id="adslot2" visibility=hide width=468 height=60></ilayer>');
		}

w('</td></tr>');
w('<tr><td>');
w('<DIV ALIGN="center"><FONT FACE="verdana, geneva, helvetica" SIZE="1" COLOR="#999999" STYLE="font: Verdana; font-size: 9px">please support our sponsors</font></DIV>');
w('</td></tr></table>');

}

function sprinksbottom()
{
w('<table border=0 width=90% cellpadding=1 cellspacing=0 bgcolor="#999999">');
w('<tr><td><FONT FACE="Verdana, Geneva, Helvetica" size="2" COLOR="#FFFFFF"><B>Sponsored Links About&reg; ' + sprinksboxtitle + '</B><IMG SRC="http://pixel3.about.com/mp/pixel/pixel.cgi?partner=vpn' + vpn + '_sp" width=1 height=1 border=0></td></tr>');
w('<tr><td align="center">');
w('<table border=0 width=100% cellpadding=5 cellspacing=0 bgcolor="#ffffff">');
w('<tr><td><a href="' + sprinks_link1 + '" target="'+tgt+'"><font face="arial, geneva, helvetica" size="2" color="#330066"><B>' + sprinks_title1 + '</B></font></a><BR><font face="arial, geneva, helvetica" size="2" color="#000000">' + sprinks_descr1 + '</font><BR><font face="arial, geneva, helvetica" color="#999999" size="2">' + sprinks_displink1 + '</font>&nbsp;&nbsp;<font face="arial, geneva, helvetica" size="1" color="#0066CC">(Listing fee:'+sprinks_cpc1+')</font></td></tr>');
w('<tr><td><a href="' + sprinks_link2 + '" target="'+tgt+'"><font face="arial, geneva, helvetica" size="2" color="#330066"><B>' + sprinks_title2 + '</B></font></a><BR><font face="arial, geneva, helvetica" size="2" color="#000000">' + sprinks_descr2 + '</font><BR><font face="arial, geneva, helvetica" color="#999999" size="2">' + sprinks_displink2 + '</font>&nbsp;&nbsp;<font face="arial, geneva, helvetica" size="1" color="#0066CC">(Listing fee:'+sprinks_cpc2+')</font></td></tr>');
w('<tr><td><div align="right"><a href="http://sprinks.about.com/" target="'+tgt+'"><font face="arial, geneva, helvetica" size="2" color="#330066"><B>Buy a Link Now!</B></font></a></div></td></tr>');
w('</td></tr></table>');
w('</td></tr></table>');
}
function sqslot1() {
w('<table border="0" cellpadding="0" cellspacing="0"><tr><td>');
if (browserType.indexOf("netscape4.0")!= -1 || browserType.indexOf("webtv")!= -1) {
			w(j+aux+'sqslot1;sz=125x125;ord=1'+rand+'" target="'+tgt+'">');
			w(i+aux+'sqslot1;sz=125x125;ord=1'+rand+'" height=125 width=125 border=0></a>');
		}
		else{
			w('<nolayer>');
			w(f+aux+'sqslot1;sz=125x125;ord=1'+rand+'" target="'+tgt+'" height=125 width=125 border=0 marginwidth=0 marginheight=0 hspace=0 vspace=0 frameborder=0 scrolling=no></a>');
			if (browserClass != 'Netscape2') {
				w(j+aux+'sqslot1;sz=125x125;ord=1'+rand+'" target="'+tgt+'">');
				w(i+aux+'sqslot1;sz=125x125;ord=1'+rand+'" height=125 width=125 border=0></a>');
			}
			w('</iframe>');
			w('</nolayer>');
			w('<ilayer id="sqslot1" visibility=hide width=125 height=125></ilayer>');
		}


w('</td></tr></table>');

}
function vmslot1() {
	w('<BR>');
	w('<table border="0" cellpadding="0" cellspacing="0"><tr><td>');
		if (browserType.indexOf("netscape4.0")!= -1 || browserType.indexOf("webtv")!= -1) {
			w(j+aux+'vmslot1;sz=120x600;ord=1'+rand+'" target="'+tgt+'">');
			w(i+aux+'vmslot1;sz=120x600;ord=1'+rand+'" height=600 width=120 border=0></a>');
		}
		else{
			w('<nolayer>');
			w(f+aux+'vmslot1;sz=120x600;ord=1'+rand+'" target="'+tgt+'" height=600 width=120 border=0 marginwidth=0 marginheight=0 hspace=0 vspace=0 frameborder=0 scrolling=no></a>');
			if (browserClass != 'Netscape2') {
				w(j+aux+'vmslot1;sz=120x600;ord=1'+rand+'" target="'+tgt+'">');
				w(i+aux+'vmslot1;sz=120x600;ord=1'+rand+'" height=600 width=120 border=0></a>');
			}
			w('</iframe>');
		    w('</nolayer>');
			w('<ilayer id="vmslot1" visibility=hide width=120 height=600></ilayer>');
		}
	
	w('</td></tr></table>');
}
	

function pw()
{
w('<A HREF="http://clicks.about.com/137/zd_luna2.asp?rURL=http://chkpt.zdnet.com/chkpt/hud00014adszdcamera/www.zdnet.com/feeds/cgi/framer/hud0001400/www.zdnet.com/special/filters/sc/camera/" TARGET="'+tgt+'"><IMG SRC="http://a1028.g.akamai.net/6/1028/968/444516c765e1c0/images.about.com/partners/vpn/partnerbox/zdnet/zd_digicam.gif" WIDTH=150 HEIGHT=60 BORDER=0 ALT="Find the right digital camera!"></a>');
}





	var color="#CC0000";
	function navbar() 
{
w('<STYLE type="text/css">'); 
w('<!--');
w('table.zzzzlunatable { background-color: #ffffff; }');
w('tr.zzzzlunatoprow { background-color: #ffffff; }');
w('tr.zzzzlunabottomrow { background-color:'+color+'; }');
w('td.lunabarlogo { text-align: left; }');
w('td.zzzzgoluna { color:#000000;	font: 12px Verdana, Geneva, Arial, Helvetica, sans-serif; font-weight:bold;	text-align: right; }');
w('td.zzzzgoluna a:hover { color:#000000; font-weight: bold; background-color:#eeeeee }');
w('td.zzzzrecommended { color:#ffffff; font: 10px Verdana, Geneva, Arial, Helvetica, sans-serif;  font-weight:bold; text-align: left; }');
w('td.zzzzrecommended b { color:#ffffff; font: 10px Verdana, Geneva, Arial, Helvetica, sans-serif;  font-weight:bold; }');
w('td.zzzzrecommendedwithlinks { color:#ffffff; font: 10px Verdana, Geneva, Arial, Helvetica, sans-serif; font-weight:bold; text-align: right; }');
w('td.zzzzrecommendedwithlinks a { color:#ffffff; font: 10px Verdana, Geneva, Arial, Helvetica, sans-serif; text-decoration:underline;  }');
w('td.zzzzrecommendedwithlinks a:hover { color:#ffffff;  font: 10px Verdana, Geneva, Arial, Helvetica, sans-serif; text-decoration:none; }');
w('td.zzzzrecommendedwithlinks a:visited { color:#ff0000; font: 10px Verdana, Geneva, Arial, Helvetica, sans-serif; }');
w('td.zzzzrecommendedwithlinks b {  color:#ffffff; font: 10px Verdana, Geneva, Arial, Helvetica, sans-serif; font-weight:bold; }');
w('// -->');
w('</STYLE>');

w('<TABLE WIDTH=100% BORDER=0 CELLSPACING=0 CELLPADDING=0 CLASS="zzzzlunatable" BGCOLOR="#FFFFFF">');
w('<FORM NAME="zzzzaboutsearchzzzz" METHOD="get" ACTION="http://partners.about.com/vpn/search/searchx.htm" TARGET="'+tgt+'">');
w('<TR CLASS="zzzzlunatoprow" BGCOLOR="#FFFFFF">');
w('<TD WIDTH=24% CLASS="lunabarlogo"><A HREF="http://'+n_guidesite+'.about.com/mbody.htm?IAM=vpn'+vpn+'_1" TARGET="'+tgt+'"><IMG SRC="http://a1028.g.akamai.net/6/1028/968/444516c765e1c0/images.about.com/partners/vpn/brandbar/logo_lunabar_'+n_guidesite+'.gif" WIDTH=200 HEIGHT=20 BORDER=0></A></TD>');
w('<TD WIDTH=75% HEIGHT=25 CLASS="zzzzgoluna"><DIV ALIGN="right"><FONT FACE="Verdana, Geneva, Helvetica" SIZE="2">Search About&reg;</FONT> <INPUT TYPE="text" NAME="terms" SIZE="15"></TD>');
w('<TD width=1%><TABLE BORDER=0 CELLPADDING=2 CELLSPACING=2><TR><TD CLASS="zzzzgoluna"><A HREF="javascript:document.zzzzaboutsearchzzzz.submit()"><FONT FACE="Verdana, Geneva, Helvetica" COLOR="#000000" SIZE=2><B>Go</B></FONT></A></TD></TR></TABLE></TD>');
w('</TR>');
w('<TR CLASS="zzzzlunabottomrow" BGCOLOR="'+color+'" HEIGHT=15>');
w('<TD CLASS="zzzzrecommended"><FONT FACE="Verdana, Geneva, Helvetica" SIZE="1" COLOR="#FFFFFF"><B>Recommended Site</B></FONT></TD>');
w('<TD CLASS="zzzzrecommendedwithlinks" colspan=2><DIV ALIGN="right"><FONT FACE="Verdana, Geneva, Helvetica" SIZE="1" COLOR="#FFFFFF"><B>'+nlink1+' | '+nlink2+' | '+nlink3+' | '+nlink4+'</B></FONT></DIV></TD>');
w('<TR><TD COLSPAN=3 HEIGHT=1 BGCOLOR=#000000><IMG SRC="http://a1028.g.akamai.net/6/1028/968/444516c765e1c0/images.about.com/all/bullets/dot_clea.gif" width=1 height=1></TD></TR>');
w('<INPUT TYPE="hidden" NAME="SUName" VALUE="'+n_guidesite+'">');
w('<INPUT TYPE="hidden" NAME="IAM" VALUE="vpn'+vpn+'_2">');
w('</FORM>');
w('</TABLE>');
}


function ar() {
}

function vlslot1() {
}

function tabletop()
{
w('<TABLE CELLSPACING=2 CELLPADDING=2 BORDER=0>');
w('<TR><TD COLSPAN=2>');
w(top);
w('</TD></TR>');
w('<TR VALIGN=top><TD COLSPAN=2 WIDTH=100%>');
w('<DIV ALIGN="center">');
w('<TABLE BORDER=0>');
w('<TR VALIGN=top><TD WIDTH=468>');
w(topad1);
w('</TD><TD WIDTH=15><IMG SRC="http://images.about.com/all/bullets/dot_clea.gif" WIDTH=15 HEIGHT=1></TD><TD>');
w(topad2);
w('</TD></TR></TABLE></DIV></TD></TR><TR VALIGN=top><TD WIDTH=90%>');
}

function tablebottom()
{
w('</TD><TD ALIGN="right"><TABLE BORDER=0><TR><TD><DIV ALIGN="center">');
w(column);
w('</DIV></TD></TR></TABLE></TD></TR><TR VALIGN=top><TD COLSPAN=2 ALIGN="center">');
w(botspr);
w('</TD></TR><TR VALIGN=top><TD COLSPAN=2 ALIGN=center><DIV ALIGN="center">');
w('<TABLE BORDER=0><TR VALIGN=top><TD WIDTH=468>');
w(botad1);
w('</TD><TD WIDTH=15><IMG SRC="http://images.about.com/all/bullets/dot_clea.gif" WIDTH=15 HEIGHT=1></TD><TD>');
w(botad2);
w('</TD></TR></TABLE></DIV></TD></TR><TR><TD COLSPAN=2>');
w(bottom);
w('</TD></TR></TABLE>');
adclose();
w('<IMG SRC="http://pixel3.about.com/mp/pixel/pixel.cgi?partner=vpn' + vpn + '_ld" WIDTH=1 HEIGHT=1>');
}

function tabletop2()
{
w('<TABLE CELLSPACING=2 CELLPADDING=2 BORDER=0>');
w('<TR><TD COLSPAN=2>');
w(top);
w('</TD></TR>');
w('<TR VALIGN=top><TD WIDTH=100%>');
w('<DIV ALIGN="center">');
w('<TABLE BORDER=0>');
w('<TR VALIGN=top><TD WIDTH=468>');
w(topad1);
w('</TD><TD WIDTH=15><IMG SRC="http://images.about.com/all/bullets/dot_clea.gif" WIDTH=15 HEIGHT=1></TD><TD>');
w(topad2);
w('</TD></TR></TABLE></DIV></TD></TR><TR VALIGN=top><TD WIDTH=90%>');
}

function tablebottom2()
{
w('</TD></TR><TR VALIGN=top><TD COLSPAN=2 ALIGN="center">');
w(botspr);
w('</TD></TR><TR VALIGN=top><TD ALIGN=center><DIV ALIGN="center">');
w('<TABLE BORDER=0><TR VALIGN=top><TD WIDTH=468>');
w(botad1);
w('</TD><TD WIDTH=15><IMG SRC="http://images.about.com/all/bullets/dot_clea.gif" WIDTH=15 HEIGHT=1></TD><TD>');
w(botad2);
w('</TD></TR></TABLE></DIV></TD></TR><TR><TD COLSPAN=2>');
w(bottom);
w('</TD></TR></TABLE>');
adclose();
w('<IMG SRC="http://pixel3.about.com/mp/pixel/pixel.cgi?partner=vpn' + vpn + '_ld" WIDTH=1 HEIGHT=1>');
}

function adclose() {
  if (browserType.indexOf("netscape4.0")== -1 || browserType.indexOf("webtv")== -1) {
	if (document.adslot1) {
		w('<layer clip="470,62" src="http://ad.doubleclick.net/adl'+aux+'slot1;sz=468x60;ord=1'+rand+'" target="'+tgt+'" visibility="hide" onload="moveToAbsolute(adslot1.pageX,adslot1.pageY); visibility=\'show\';"></layer>');
	}
	if (document.vlslot1) {
		w('<layer clip="150,802" src="http://ad.doubleclick.net/adl'+aux+'vlslot1;sz=148x800;ord='+rand+'" target="'+tgt+'" visibility="hide" onload="moveToAbsolute(vlslot1.pageX,vlslot1.pageY); visibility=\'show\';"></layer>');
	}
	if (document.popslot1) {
		w('<layer clip="3,3" src="http://ad.doubleclick.net/adl'+aux+'popslot1;sz=1x1;ord=1'+rand+'" target="'+tgt+'" visibility="hide" onload="moveToAbsolute(popslot1.pageX,popslot1.pageY); visibility=\'show\';"></layer>');
	}
	if (document.adslot2) {
		w('<layer clip="470,62" src="http://ad.doubleclick.net/adl'+aux+'slot1;sz=468x60;ord=1'+rand+'" target="'+tgt+'" visibility="hide" onload="moveToAbsolute(adslot2.pageX,adslot2.pageY); visibility=\'show\';"></layer>');
	}
	if (document.capslot1) {
		w('<layer clip="122,62" src="http://ad.doubleclick.net/adl'+aux+'capslot1;sz=120x60;ord='+rand+'" target="'+tgt+'" visibility="hide" onload="moveToAbsolute(capslot1.pageX,capslot1.pageY); visibility=\'show\';"></layer>');
	}
	if (document.capslot2) {
		w('<layer clip="122,62" src="http://ad.doubleclick.net/adl'+aux+'capslot1;sz=120x60;ord='+rand+'" target="'+tgt+'" visibility="hide" onload="moveToAbsolute(capslot1.pageX,capslot2.pageY); visibility=\'show\';"></layer>');
	}
	 if (document.vslot3) {
      		w('<layer clip="142,802" src="http://ad.doubleclick.net/adl'+aux+'vslot3;sz=140x800;ord=1'+rand+'" target="'+tgt+'" visibility="hide" onload="moveToAbsolute(vslot3.pageX,vslot3.pageY); visibility=\'show\';"></layer>');
	}
	if (document.vmslot1) {
			w('<layer clip="122,602" src="http://ad.doubleclick.net/adl'+aux+'vmslot1;sz=120x600;ord=1'+rand+'" target="'+tgt+'" visibility="hide" onload="moveToAbsolute(vmslot1.pageX,vmslot1.pageY); visibility=\'show\';"></layer>');
	}
	if (document.sqslot1) {
			w('<layer clip="127,127" src="http://ad.doubleclick.net/adl'+aux+'sqslot1;sz=125x125;ord=1'+rand+'" target="'+tgt+'" visibility="hide" onload="moveToAbsolute(sqslot1.pageX,sqslot1.pageY); visibility=\'show\';"></layer>');
	}	
 } 
}
