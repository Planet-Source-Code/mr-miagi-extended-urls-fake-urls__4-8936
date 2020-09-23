<div align="center">

## Extended Urls / Fake Urls


</div>

### Description

:How to extend your urls dynamically using a Mysql Database without having file structures for every directory in your Url's, this code will improve search engine status of your website and much more.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Mr\.Miagi](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/mr-miagi.md)
**Level**          |Intermediate
**User Rating**    |4.3 (13 globes from 3 users)
**Compatibility**  |ASP \(Active Server Pages\)
**Category**       |[Data Structures](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/data-structures__4-8.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/mr-miagi-extended-urls-fake-urls__4-8936/archive/master.zip)





### Source Code

<div class=Section1>
<p class=MsoNormal><span style='font-size:20.0pt'>404<span class=GramE>.asp</span><o:p></o:p></span></p>
<p class=MsoNormal><span class=GramE><span style='font-size:14.0pt'>this</span></span><span
style='font-size:14.0pt'> page should be put as your custom error page for 404
http errors.<o:p></o:p></span></p>
<table class=MsoTableGrid border=1 cellspacing=0 cellpadding=0
 style='background:#B3B3B3;border-collapse:collapse;border:none;mso-border-alt:
 solid windowtext .5pt;mso-yfti-tbllook:480;mso-padding-alt:0in 5.4pt 0in 5.4pt;
 mso-border-insideh:.5pt solid windowtext;mso-border-insidev:.5pt solid windowtext'>
 <tr style='mso-yfti-irow:0;mso-yfti-lastrow:yes'>
 <td width=590 valign=top style='width:6.15in;border:solid windowtext 1.0pt;
 mso-border-alt:solid windowtext .5pt;padding:0in 5.4pt 0in 5.4pt'>
 <p class=MsoNormal><span style='font-size:11.0pt;font-family:Garamond'>&lt;%option
 <span class=GramE>explicit%</span>&gt;<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:11.0pt;font-family:Garamond'>&lt;%<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:11.0pt;font-family:Garamond'>dim
 refer, <span class=SpellE>indexc</span>, <span class=SpellE>strpage,strpagename</span>,
 <span class=SpellE>dbcon</span>, connection, <span class=SpellE>sql</span>, <span
 class=SpellE>dbrs</span>, <span class=SpellE>realpage</span><o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:11.0pt;font-family:Garamond'>refer=<span
 class=SpellE>Request.ServerVariables</span>(&quot;QUERY_STRING&quot;)<o:p></o:p></span></p>
 <p class=MsoNormal><span class=SpellE><span style='font-size:11.0pt;
 font-family:Garamond'>indexc</span></span><span style='font-size:11.0pt;
 font-family:Garamond'>=<span class=SpellE>instrrev</span>(refer,&quot;/&quot;)<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:11.0pt;font-family:Garamond'>'++++++++++++++++++++++++++++++++++++++++++++++++++<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:11.0pt;font-family:Garamond'>'+
 getting the number of &quot;/&quot; in refer<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:11.0pt;font-family:Garamond'>'+ and
 stripping them off<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:11.0pt;font-family:Garamond'>'+
 then removing the .<span class=SpellE>htm</span> or .html from them<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:11.0pt;font-family:Garamond'>'++++++++++++++++++++++++++++++++++++++++++++++++++<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:11.0pt;font-family:Garamond'>if (<span
 class=SpellE>indexc</span>&gt;0) then<o:p></o:p></span></p>
 <p class=MsoNormal><span class=SpellE><span style='font-size:11.0pt;
 font-family:Garamond'>strpage</span></span><span style='font-size:11.0pt;
 font-family:Garamond'> = Right(<span class=SpellE>refer,Len</span>(refer)-<span
 class=SpellE>indexc</span>)<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:11.0pt;font-family:Garamond'>refer
 = Left(refer,indexc-1)<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:11.0pt;font-family:Garamond'>end if<o:p></o:p></span></p>
 <p class=MsoNormal><span class=SpellE><span style='font-size:11.0pt;
 font-family:Garamond'>strpage</span></span><span style='font-size:11.0pt;
 font-family:Garamond'>=replace(strpage,&quot;.html&quot;,&quot;&quot;)<o:p></o:p></span></p>
 <p class=MsoNormal><span class=SpellE><span style='font-size:11.0pt;
 font-family:Garamond'>strpage</span></span><span style='font-size:11.0pt;
 font-family:Garamond'>=replace(strpage,&quot;.htm&quot;,&quot;&quot;)<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:11.0pt;font-family:Garamond'>'++++++++++++++++++++++++++++++++++++++++++++++++++<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:11.0pt;font-family:Garamond'>'+ <span
 class=GramE>connection</span> strings...<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:11.0pt;font-family:Garamond'>'+ <span
 class=SpellE>sql</span>= all fields where the page name is the requested <span
 class=SpellE>url</span><o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:11.0pt;font-family:Garamond'>'+
 that we stripped above<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:11.0pt;font-family:Garamond'>'++++++++++++++++++++++++++++++++++++++++++++++++++<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:11.0pt;font-family:Garamond'>set <span
 class=SpellE>dbcon</span> = <span class=SpellE>server.createobject</span>(&quot;<span
 class=SpellE>adodb.connection</span>&quot;)<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:11.0pt;font-family:Garamond'>connection
 = &quot;Driver={<span class=SpellE>Mysql</span>}; Server=69.56.199.234;
 Database=; UID=; PWD=;&quot;<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:11.0pt;font-family:Garamond'>SQL =
 &quot;SELECT * FROM `pages` WHERE <span class=SpellE>PageName</span>='&quot;&amp;<span
 class=SpellE>strpage</span>&amp;&quot;'&quot;<o:p></o:p></span></p>
 <p class=MsoNormal><span class=SpellE><span style='font-size:11.0pt;
 font-family:Garamond'>dbcon.open</span></span><span style='font-size:11.0pt;
 font-family:Garamond'>(connection)<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:11.0pt;font-family:Garamond'>set <span
 class=SpellE>dbrs</span> = <span class=SpellE>server.CreateObject</span>(&quot;<span
 class=SpellE>adodb.recordset</span>&quot;)<o:p></o:p></span></p>
 <p class=MsoNormal><span class=SpellE><span style='font-size:11.0pt;
 font-family:Garamond'>dbrs.open</span></span><span style='font-size:11.0pt;
 font-family:Garamond'> <span class=SpellE>sql</span>, <span class=SpellE>dbcon</span><o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:11.0pt;font-family:Garamond'>if not
 <span class=SpellE>dbrs.eof</span> then<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:11.0pt;font-family:Garamond'>'++++++++++++++++++++++++++++++++++++++++++++++++++<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:11.0pt;font-family:Garamond'>'+
 setting our variables<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:11.0pt;font-family:Garamond'>'+
 setting the session of the <span class=SpellE>pageID</span> since we cannot<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:11.0pt;font-family:Garamond'>'+
 transfer regular variables in <span class=SpellE>server.transfer</span><o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:11.0pt;font-family:Garamond'>'++++++++++++++++++++++++++++++++++++++++++++++++++<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:11.0pt;font-family:Garamond'>set <span
 class=SpellE>realpage</span>=<span class=SpellE>dbrs.fields</span>(&quot;<span
 class=SpellE>realpage</span>&quot;)<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:11.0pt;font-family:Garamond'>session(&quot;<span
 class=SpellE>PageID</span>&quot;)=<span class=SpellE>dbrs.fields</span>(&quot;<span
 class=SpellE>PageID</span>&quot;)<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:11.0pt;font-family:Garamond'>'++++++++++++++++++++++++++++++++++++++++++++++++++<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:11.0pt;font-family:Garamond'>'+
 Sending the user a page, but not redirecting them<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:11.0pt;font-family:Garamond'>'+ (<span
 class=SpellE>realpage</span>) should be a field in our database with the page
 name. EX. &quot;News&quot;<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:11.0pt;font-family:Garamond'>'+
 then we redirect them to our main page if it is a REAL 404 error<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:11.0pt;font-family:Garamond'>'++++++++++++++++++++++++++++++++++++++++++++++++++<o:p></o:p></span></p>
 <p class=MsoNormal><span class=SpellE><span style='font-size:11.0pt;
 font-family:Garamond'>server.transfer</span></span><span style='font-size:
 11.0pt;font-family:Garamond'>(&quot;/asp/&quot;&amp;<span class=SpellE>realpage&amp;&quot;.asp</span>&quot;)<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:11.0pt;font-family:Garamond'>else<o:p></o:p></span></p>
 <p class=MsoNormal><span class=SpellE><span style='font-size:11.0pt;
 font-family:Garamond'>response.redirect</span></span><span style='font-size:
 11.0pt;font-family:Garamond'>(&quot;index.html&quot;)<o:p></o:p></span></p>
 <p class=MsoNormal><span class=SpellE><span style='font-size:11.0pt;
 font-family:Garamond'>dbrs.close</span></span><span style='font-size:11.0pt;
 font-family:Garamond'><o:p></o:p></span></p>
 <p class=MsoNormal><span class=SpellE><span style='font-size:11.0pt;
 font-family:Garamond'>dbcon.close</span></span><span style='font-size:11.0pt;
 font-family:Garamond'><o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:11.0pt;font-family:Garamond'>end if<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:11.0pt;font-family:Garamond'>%&gt;</span></p>
 </td>
 </tr>
</table>
<p class=MsoNormal><o:p>&nbsp;</o:p></p>
<p class=MsoNormal><span class=SpellE><span style='font-size:20.0pt'>Wwwroot/asp/urlgen.asp</span></span><span
style='font-size:20.0pt'><o:p></o:p></span></p>
<p class=MsoNormal><span class=GramE><span style='font-size:14.0pt'>this</span></span><span
style='font-size:14.0pt'> page contains the script which generates the fake <span
class=SpellE>urls</span><o:p></o:p></span></p>
<table class=MsoTableGrid border=1 cellspacing=0 cellpadding=0
 style='background:#333333;border-collapse:collapse;border:none;mso-border-alt:
 solid windowtext .5pt;mso-yfti-tbllook:480;mso-padding-alt:0in 5.4pt 0in 5.4pt;
 mso-border-insideh:.5pt solid windowtext;mso-border-insidev:.5pt solid windowtext'>
 <tr style='mso-yfti-irow:0;mso-yfti-lastrow:yes'>
 <td width=590 valign=top style='width:6.15in;border:solid windowtext 1.0pt;
 mso-border-alt:solid windowtext .5pt;padding:0in 5.4pt 0in 5.4pt'>
 <p class=MsoNormal><span style='font-size:11.0pt;font-family:Garamond;
 color:white'>&lt;% <o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:11.0pt;font-family:Garamond;
 color:white'>'++++++++++++++++++++++++++++++++++++++++++++++++++<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:11.0pt;font-family:Garamond;
 color:white'>'+ this page writes our very false <span class=SpellE>urls</span>
 and puts <o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:11.0pt;font-family:Garamond;
 color:white'>'+ them into links<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:11.0pt;font-family:Garamond;
 color:white'>'++++++++++++++++++++++++++++++++++++++++++++++++++<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:11.0pt;font-family:Garamond;
 color:white'>'++++++++++++++++++++++++++++++++++++++++++++++++++<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:11.0pt;font-family:Garamond;
 color:white'>'+ this is our domain, change it to your domain<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:11.0pt;font-family:Garamond;
 color:white'>'+ domain format must include http://'YOUR DOMAIN'/<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:11.0pt;font-family:Garamond;
 color:white'>'++++++++++++++++++++++++++++++++++++++++++++++++++<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:11.0pt;font-family:Garamond;
 color:white'>const <span class=SpellE>url</span>=&quot;http://halonation.com/&quot;<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:11.0pt;font-family:Garamond;
 color:white'>'++++++++++++++++++++++++++++++++++++++++++++++++++<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:11.0pt;font-family:Garamond;
 color:white'>'+ our function for creating <span class=SpellE>urls</span><o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:11.0pt;font-family:Garamond;
 color:white'>'+ writes a link with the db <span class=SpellE>urls</span>,
 page directories<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:11.0pt;font-family:Garamond;
 color:white'>'+ and <span class=SpellE>pagename</span>+ a .<span
 class=SpellE>htm</span><o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:11.0pt;font-family:Garamond;
 color:white'>'++++++++++++++++++++++++++++++++++++++++++++++++++<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:11.0pt;font-family:Garamond;
 color:white'>Function <span class=SpellE>CreateUrl</span>(ID)<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:11.0pt;font-family:Garamond;
 color:white'>set <span class=SpellE>dbcon</span> = <span class=SpellE>server.createobject</span>(&quot;<span
 class=SpellE>adodb.connection</span>&quot;)<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:11.0pt;font-family:Garamond;
 color:white'>connection = &quot;Driver={<span class=SpellE>Mysql</span>};
 Server=69.56.199.234; Database=; UID=; PWD=;&quot;<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:11.0pt;font-family:Garamond;
 color:white'>SQL = &quot;SELECT * FROM `pages` WHERE <span class=SpellE>PageID</span>='&quot;&amp;ID&amp;&quot;'&quot;<o:p></o:p></span></p>
 <p class=MsoNormal><span class=SpellE><span style='font-size:11.0pt;
 font-family:Garamond;color:white'>dbcon.open</span></span><span
 style='font-size:11.0pt;font-family:Garamond;color:white'>(connection)<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:11.0pt;font-family:Garamond;
 color:white'>set <span class=SpellE>dbrs</span> = <span class=SpellE>server.CreateObject</span>(&quot;<span
 class=SpellE>adodb.recordset</span>&quot;)<o:p></o:p></span></p>
 <p class=MsoNormal><span class=SpellE><span style='font-size:11.0pt;
 font-family:Garamond;color:white'>dbrs.open</span></span><span
 style='font-size:11.0pt;font-family:Garamond;color:white'> <span
 class=SpellE>sql</span>, <span class=SpellE>dbcon</span><o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:11.0pt;font-family:Garamond;
 color:white'>if not <span class=SpellE>dbrs.eof</span> then<o:p></o:p></span></p>
 <p class=MsoNormal><span class=SpellE><span style='font-size:11.0pt;
 font-family:Garamond;color:white'>pagename</span></span><span
 style='font-size:11.0pt;font-family:Garamond;color:white'>=<span
 class=SpellE>dbrs.fields</span>(&quot;<span class=SpellE>pagename</span>&quot;)<o:p></o:p></span></p>
 <p class=MsoNormal><span class=SpellE><span style='font-size:11.0pt;
 font-family:Garamond;color:white'>pagedir</span></span><span
 style='font-size:11.0pt;font-family:Garamond;color:white'>=<span
 class=SpellE>dbrs.fields</span>(&quot;<span class=SpellE>pagedir</span>&quot;)<o:p></o:p></span></p>
 <p class=MsoNormal><span class=SpellE><span style='font-size:11.0pt;
 font-family:Garamond;color:white'>CreateUrl</span></span><span
 style='font-size:11.0pt;font-family:Garamond;color:white'>=url&amp;pagedir&amp;pagename&amp;&quot;.htm&quot;<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:11.0pt;font-family:Garamond;
 color:white'>end if<o:p></o:p></span></p>
 <p class=MsoNormal><span class=SpellE><span style='font-size:11.0pt;
 font-family:Garamond;color:white'>dbrs.close</span></span><span
 style='font-size:11.0pt;font-family:Garamond;color:white'><o:p></o:p></span></p>
 <p class=MsoNormal><span class=SpellE><span style='font-size:11.0pt;
 font-family:Garamond;color:white'>dbcon.close</span></span><span
 style='font-size:11.0pt;font-family:Garamond;color:white'><o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:11.0pt;font-family:Garamond;
 color:white'>End Function<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:11.0pt;font-family:Garamond;
 color:white'>'++++++++++++++++++++++++++++++++++++++++++++++++++<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:11.0pt;font-family:Garamond;
 color:white'>'+ our function for making the link name<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:11.0pt;font-family:Garamond;
 color:white'>'+ connects, gets <span class=SpellE>linkname</span> from the
 DB, and returns it<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:11.0pt;font-family:Garamond;
 color:white'>'++++++++++++++++++++++++++++++++++++++++++++++++++<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:11.0pt;font-family:Garamond;
 color:white'>Function <span class=SpellE>CreateName</span>(ID)<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:11.0pt;font-family:Garamond;
 color:white'>set <span class=SpellE>dbcon</span> = <span class=SpellE>server.createobject</span>(&quot;<span
 class=SpellE>adodb.connection</span>&quot;)<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:11.0pt;font-family:Garamond;
 color:white'>connection = &quot;Driver={<span class=SpellE>Mysql</span>};
 Server=69.56.199.234; Database=; UID=; PWD=;&quot;<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:11.0pt;font-family:Garamond;
 color:white'>SQL = &quot;SELECT `<span class=SpellE>linktext</span>` FROM
 `pages` WHERE <span class=SpellE>PageID</span>='&quot;&amp;ID&amp;&quot;'&quot;<o:p></o:p></span></p>
 <p class=MsoNormal><span class=SpellE><span style='font-size:11.0pt;
 font-family:Garamond;color:white'>dbcon.open</span></span><span
 style='font-size:11.0pt;font-family:Garamond;color:white'>(connection)<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:11.0pt;font-family:Garamond;
 color:white'>set <span class=SpellE>dbrs</span> = <span class=SpellE>server.CreateObject</span>(&quot;<span
 class=SpellE>adodb.recordset</span>&quot;)<o:p></o:p></span></p>
 <p class=MsoNormal><span class=SpellE><span style='font-size:11.0pt;
 font-family:Garamond;color:white'>dbrs.open</span></span><span
 style='font-size:11.0pt;font-family:Garamond;color:white'> <span
 class=SpellE>sql</span>, <span class=SpellE>dbcon</span><o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:11.0pt;font-family:Garamond;
 color:white'>if not <span class=SpellE>dbrs.eof</span> then<o:p></o:p></span></p>
 <p class=MsoNormal><span class=SpellE><span style='font-size:11.0pt;
 font-family:Garamond;color:white'>CreateName</span></span><span
 style='font-size:11.0pt;font-family:Garamond;color:white'>=<span
 class=SpellE>dbrs.fields</span>(&quot;<span class=SpellE>linktext</span>&quot;)<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:11.0pt;font-family:Garamond;
 color:white'>end if<o:p></o:p></span></p>
 <p class=MsoNormal><span class=SpellE><span style='font-size:11.0pt;
 font-family:Garamond;color:white'>dbrs.close</span></span><span
 style='font-size:11.0pt;font-family:Garamond;color:white'><o:p></o:p></span></p>
 <p class=MsoNormal><span class=SpellE><span style='font-size:11.0pt;
 font-family:Garamond;color:white'>dbcon.close</span></span><span
 style='font-size:11.0pt;font-family:Garamond;color:white'><o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:11.0pt;font-family:Garamond;
 color:white'>End Function<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:11.0pt;font-family:Garamond;
 color:white'>'++++++++++++++++++++++++++++++++++++++++++++++++++<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:11.0pt;font-family:Garamond;
 color:white'>'+ connection strings<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:11.0pt;font-family:Garamond;
 color:white'>'+ <span class=SpellE>rowcount</span> gets the number of rows
 total, for our Loop<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:11.0pt;font-family:Garamond;
 color:white'>'++++++++++++++++++++++++++++++++++++++++++++++++++<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:11.0pt;font-family:Garamond;
 color:white'>set <span class=SpellE>dbcon</span> = <span class=SpellE>server.createobject</span>(&quot;<span
 class=SpellE>adodb.connection</span>&quot;)<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:11.0pt;font-family:Garamond;
 color:white'>connection = &quot;Driver={<span class=SpellE>Mysql</span>};
 Server=69.56.199.234; Database=; UID=; PWD=;&quot;<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:11.0pt;font-family:Garamond;
 color:white'>SQL = &quot;SELECT COUNT(*) as Count FROM `pages`&quot;<o:p></o:p></span></p>
 <p class=MsoNormal><span class=SpellE><span style='font-size:11.0pt;
 font-family:Garamond;color:white'>dbcon.open</span></span><span
 style='font-size:11.0pt;font-family:Garamond;color:white'>(connection)<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:11.0pt;font-family:Garamond;
 color:white'>set <span class=SpellE>dbrs</span> = <span class=SpellE>server.CreateObject</span>(&quot;<span
 class=SpellE>adodb.recordset</span>&quot;)<o:p></o:p></span></p>
 <p class=MsoNormal><span class=SpellE><span style='font-size:11.0pt;
 font-family:Garamond;color:white'>dbrs.open</span></span><span
 style='font-size:11.0pt;font-family:Garamond;color:white'> <span
 class=SpellE>sql</span>, <span class=SpellE>dbcon</span><o:p></o:p></span></p>
 <p class=MsoNormal><span class=SpellE><span style='font-size:11.0pt;
 font-family:Garamond;color:white'>RowCount</span></span><span
 style='font-size:11.0pt;font-family:Garamond;color:white'>=<span
 class=SpellE>dbrs</span>(&quot;Count&quot;)<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:11.0pt;font-family:Garamond;
 color:white'>'++++++++++++++++++++++++++++++++++++++++++++++++++<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:11.0pt;font-family:Garamond;
 color:white'>'+ we loop Loopvar1 to <span class=SpellE>Rowcount</span> to
 write every link<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:11.0pt;font-family:Garamond;
 color:white'>'+ inside the DB, you might want to change this <o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:11.0pt;font-family:Garamond;
 color:white'>'+ depending on which page it is displayed on<span class=GramE>.,</span>
 <o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:11.0pt;font-family:Garamond;
 color:white'>'+ just remove <span class=SpellE>rowcount</span>=<span
 class=SpellE>dbrs</span>(&quot;count&quot;) and replace with<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:11.0pt;font-family:Garamond;
 color:white'>'+ <span class=SpellE>RowCount</span>='YOUR COUNT<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:11.0pt;font-family:Garamond;
 color:white'>'++++++++++++++++++++++++++++++++++++++++++++++++++<o:p></o:p></span></p>
 <p class=MsoNormal><span class=GramE><span style='font-size:11.0pt;
 font-family:Garamond;color:white'>for</span></span><span style='font-size:
 11.0pt;font-family:Garamond;color:white'> LoopVar1=0 to <span class=SpellE>RowCount</span>%&gt;<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:11.0pt;font-family:Garamond;
 color:white'>&lt;a <span class=SpellE>href</span>=&lt;%=<span class=SpellE>CreateUrl</span>(LoopVar1)%&gt;&gt;&lt;%=<span
 class=SpellE>CreateName</span>(LoopVar1)%&gt;&lt;/a&gt;<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:11.0pt;font-family:Garamond;
 color:white'>&lt;<span class=SpellE>br</span>&gt;<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:11.0pt;font-family:Garamond;
 color:white'>&lt;%next<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:11.0pt;font-family:Garamond;
 color:white'>'++++++++++++++++++++++++++++++++++++++++++++++++++<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:11.0pt;font-family:Garamond;
 color:white'>'+ disposing of garbage<span class=GramE>..</span><o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:11.0pt;font-family:Garamond;
 color:white'>'+ <o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:11.0pt;font-family:Garamond;
 color:white'>'++++++++++++++++++++++++++++++++++++++++++++++++++<o:p></o:p></span></p>
 <p class=MsoNormal><span style='font-size:11.0pt;font-family:Garamond;
 color:white'>set <span class=SpellE>dbrs</span>=nothing<o:p></o:p></span></p>
 <p class=MsoNormal><span class=GramE><span style='font-size:11.0pt;
 font-family:Garamond;color:white'>set</span></span><span style='font-size:
 11.0pt;font-family:Garamond;color:white'> <span class=SpellE>dbcon</span>=nothing%&gt;</span><span
 style='font-size:11.0pt;font-family:Garamond'><o:p></o:p></span></p>
 </td>
 </tr>
</table>
<p class=MsoNormal><span style='font-size:20.0pt'><o:p>&nbsp;</o:p></span></p>
<p class=MsoNormal><span style='font-size:20.0pt'><o:p>&nbsp;</o:p></span></p>
<p class=MsoNormal><o:p>&nbsp;</o:p></p>
<p class=MsoNormal>Use This SQL <span class=GramE>To</span> create your table
correctly.</p>
<table class=MsoTableGrid border=1 cellspacing=0 cellpadding=0
 style='background:#0C0C0C;border-collapse:collapse;border:none;mso-border-alt:
 solid windowtext .5pt;mso-yfti-tbllook:480;mso-padding-alt:0in 5.4pt 0in 5.4pt;
 mso-border-insideh:.5pt solid windowtext;mso-border-insidev:.5pt solid windowtext'>
 <tr style='mso-yfti-irow:0;mso-yfti-lastrow:yes'>
 <td width=590 valign=top style='width:6.15in;border:solid windowtext 1.0pt;
 mso-border-alt:solid windowtext .5pt;padding:0in 5.4pt 0in 5.4pt'>
 <p class=MsoNormal><span style='color:white'>CREATE TABLE `pages` (<o:p></o:p></span></p>
 <p class=MsoNormal><span style='color:white'><span style='mso-spacerun:yes'> 
 </span>`<span class=SpellE>PageID</span>` <span class=SpellE>tinyint</span>(4)
 NOT NULL <span class=SpellE>auto_increment</span>,<o:p></o:p></span></p>
 <p class=MsoNormal><span style='color:white'><span style='mso-spacerun:yes'> 
 </span>`<span class=SpellE>PageName</span>` <span class=SpellE>tinytext</span>,<o:p></o:p></span></p>
 <p class=MsoNormal><span style='color:white'><span style='mso-spacerun:yes'> 
 </span>`<span class=SpellE>PageDir</span>` text,<o:p></o:p></span></p>
 <p class=MsoNormal><span style='color:white'><span style='mso-spacerun:yes'> 
 </span>`<span class=SpellE>RealPage</span>` <span class=SpellE>tinytext</span>,<o:p></o:p></span></p>
 <p class=MsoNormal><span style='color:white'><span style='mso-spacerun:yes'> 
 </span>`<span class=SpellE>LinkText</span>` <span class=SpellE>tinytext</span>,<o:p></o:p></span></p>
 <p class=MsoNormal><span style='color:white'><span style='mso-spacerun:yes'> 
 </span>PRIMARY KEY<span style='mso-spacerun:yes'>  </span>(`<span
 class=SpellE>PageID</span>`)<o:p></o:p></span></p>
 <p class=MsoNormal><span style='color:white'>) TYPE=<span class=SpellE>MyISAM</span>;</span></p>
 </td>
 </tr>
</table>
<p class=MsoNormal><o:p>&nbsp;</o:p></p>
<p class=MsoNormal>To Conclude, write a &lt;!--#include virtual=”./asp/<span
class=SpellE>urlgen.asp</span>”--&gt; in the body of your content page inside
root/asp/ where you want your links generated. And always- CUSTOMIZE</p>
<p class=MsoNormal>Summary<span class=GramE>:</span><br>
Extended <span class=SpellE>Urls</span> help with marketing, page rank on
search engines, and much more. Without having to really make all the
directories physically. I have a running example of this code at <a
href="http://www.halonation.com/asp/news.asp">www.halonation.com/asp/news.asp</a>,
or any <span class=SpellE>url</span> you can think of that isn’t on my web
server- <span style='font-family:Wingdings;mso-ascii-font-family:"Times New Roman";
mso-hansi-font-family:"Times New Roman";mso-char-type:symbol;mso-symbol-font-family:
Wingdings'><span style='mso-char-type:symbol;mso-symbol-font-family:Wingdings'>J</span></span><br>
On A side note, if anyone finds out a way to include char(63), &amp;#63;, or ? <span
class=GramE>into</span> the <span class=SpellE>server.transfer</span> without
getting a <span class=SpellE>server.mappath</span> error, please tell me,
because that is one of the things that would make this script Very Good.</p>
<p class=MsoNormal><o:p>&nbsp;</o:p></p>
<p class=MsoNormal>This Code has been provided free grant of use
non-commercially</p>
<p class=MsoNormal><span class=SpellE><span style='font-family:Garamond'>Copywrite</span></span><span
style='font-family:Garamond'> synapse design 2003-2004<o:p></o:p></span></p>
</div>

