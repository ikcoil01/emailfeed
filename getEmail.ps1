Set-Location $PSScriptRoot
$olFolderInbox = 6
$outlook = new-object -com outlook.application;
$mapi = $outlook.GetNameSpace("MAPI");
$inbox = $mapi.GetDefaultFolder($olFolderInbox)
$itemNumber=0
$maxMailCount=$inbox.items.Count
New-Item -ItemType Directory -Path "$PSScriptRoot\mailbox"
foreach($item in $inbox.items){
    $itemNumber=$itemNumber+1
    $item.HTMLBody > "$PSScriptRoot\mailbox\page$itemNumber.html"
}
"<script>var maxMailCount=$maxMailCount</script>" > $PSScriptRoot\index.html
$index='&lt;style&gt;
iframe {
-moz-transform: scale(.9, .9);
-webkit-transform: scale(.9, .9);
-o-transform: scale(.9, .9);
-ms-transform: scale(.9, .9);
transform: scale(.9, .9);
-moz-transform-origin: top left;
-webkit-transform-origin: top left;
-o-transform-origin: top left;
-ms-transform-origin: top left;
transform-origin: top left;
}
&lt;/style&gt;
&lt;iframe src=&quot;./mailbox/page1.html&quot; title=&quot;description&quot; width=&quot;100%&quot; height=&quot;50%&quot;&gt;&lt;/iframe&gt;
&lt;iframe src=&quot;./mailbox/page2.html&quot; title=&quot;description&quot; width=&quot;100%&quot; height=&quot;50%&quot;&gt;&lt;/iframe&gt;
&lt;iframe src=&quot;./mailbox/page3.html&quot; title=&quot;description&quot; width=&quot;100%&quot; height=&quot;50%&quot;&gt;&lt;/iframe&gt;
&lt;iframe src=&quot;./mailbox/page4.html&quot; title=&quot;description&quot; width=&quot;100%&quot; height=&quot;50%&quot;&gt;&lt;/iframe&gt;

&lt;script&gt;
var mailcount=4;
window.onscroll = function(ev) {
if ((window.innerHeight + window.scrollY) &gt;= document.body.offsetHeight) {
  mailcount=mailcount+1;
  if(maxMailCount == mailcount){
    alert(&quot;Read all mail&quot;)
  }
  var iframe = document.createElement(&quot;iframe&quot;);
  iframe.src = &quot;./mailbox/page&quot;+mailcount+&quot;.html&quot;;
  iframe.width=&quot;100%&quot;;
  iframe.height=height=&quot;50%&quot;;
  document.body.appendChild(iframe);
}
};
&lt;/script&gt;'
[System.Web.HttpUtility]::HtmlDecode("$index") >> $PSScriptRoot\index.html
