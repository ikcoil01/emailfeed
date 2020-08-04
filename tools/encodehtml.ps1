Add-Type -AssemblyName System.Web
[System.Web.HttpUtility]::HtmlEncode('<style>
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
</style>
<iframe src="./mailbox/page1.html" title="description" width="100%" height="50%"></iframe>
<iframe src="./mailbox/page2.html" title="description" width="100%" height="50%"></iframe>
<iframe src="./mailbox/page3.html" title="description" width="100%" height="50%"></iframe>
<iframe src="./mailbox/page4.html" title="description" width="100%" height="50%"></iframe>

<script>
var mailcount=4;
window.onscroll = function(ev) {
if ((window.innerHeight + window.scrollY) >= document.body.offsetHeight) {
  mailcount=mailcount+1;
  if(maxMailCount == mailcount){
    alert("Read all mail")
  }
  var iframe = document.createElement("iframe");
  iframe.src = "./mailbox/page"+mailcount+".html";
  iframe.width="100%";
  iframe.height=height="50%";
  document.body.appendChild(iframe);
}
};
</script>')