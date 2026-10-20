const {chromium}=require('/opt/node22/lib/node_modules/playwright');
(async()=>{const b=await chromium.launch();const p=await b.newPage({viewport:{width:1024,height:1536},deviceScaleFactor:2});
await p.goto('file://'+process.argv[2]);await p.waitForTimeout(1500);await p.screenshot({path:process.argv[3]});await b.close();})();
