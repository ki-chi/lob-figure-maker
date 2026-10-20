const {test}=require('node:test');
const assert=require('node:assert/strict');
const fs=require('node:fs');
const path=require('node:path');
const vm=require('node:vm');
const source=fs.readFileSync(path.join(__dirname,'../index.html'),'utf8').match(/<script>([\s\S]*?)<\/script>/)[1];

// A small DOM substitute lets us test actual application handlers without dependencies.
function app() {
 const registry=new Map(),downloads=[],clipboard=[],urls=[];
 class Element {
  constructor(tag='div'){this.tag=tag;this.children=[];this.attrs={};this.dataset={};this.value='';this.textContent='';this.events={};this.classList={toggle(){}};}
  set id(id){this._id=id;registry.set(id,this);}get id(){return this._id;}
  append(...children){for(const child of children){child.parentElement=this;this.children.push(child);}}
  replaceChildren(...children){this.children=[];this.append(...children);}
  setAttribute(key,value){this.attrs[key]=value;}getAttribute(key){return this.attrs[key];}
  addEventListener(key,fn){this.events[key]=fn;}focus(){}remove(){}
  click(){if(this.tag==='a')downloads.push(this);return this.events.click?.({target:this});}
  getContext(){return {font:'',measureText(text){return {width:text.length*17};},drawImage(){}};}
  toBlob(callback){callback(new Blob(['mock png'],{type:'image/png'}));}
 }
 for(const id of source.matchAll(/\$\('([^']+)'\)/g))if(!registry.has(id[1])){const e=new Element();e.id=id[1];}
 const settingKeys=['askTitle','bidTitle','askColor','askHighlightColor','bidColor','bidHighlightColor','fontSize','barHeight','rowGap','barWidth','background'];
 for(const id of settingKeys){const e=new Element();e.id=id;e.parentElement={firstChild:{textContent:id}};}
 function inputs(){return registry.get('rows').children.flatMap(tr=>tr.children.flatMap(td=>td.children)).filter(e=>e.tag==='input');}
 const labelInputs=()=>registry.get('label-rows').children.flatMap(tr=>tr.children.flatMap(td=>td.children)).filter(e=>e.tag==='input');
 const document={createElement:tag=>new Element(tag),getElementById:id=>registry.get(id),body:new Element(),querySelectorAll:selector=>selector==='input[data-field]'?[...inputs(),...labelInputs()]:['save-project','save-svg','save-png','copy-image'].map(id=>registry.get(id)),querySelector:()=>null};
 const context=vm.createContext({document,structuredClone,Blob,console,setTimeout:fn=>{fn();},URL:{createObjectURL(blob){urls.push(blob);return 'blob:test';},revokeObjectURL(){}},navigator:{clipboard:{async write(items){clipboard.push(items);await items[0].values['image/png'];}}},ClipboardItem:class{constructor(values){this.values=values;}},Image:class{set src(value){queueMicrotask(()=>this.onload());}}});
 vm.runInContext(source,context);
 return {registry,downloads,clipboard,urls,run:code=>vm.runInContext(code,context),inputs,async event(id,type,event={}){await registry.get(id).events[type](event);}};
}
test('sample matches the reference and highlights only the outer end',()=>{
 const a=app();assert.equal(a.run('valid'),true);assert.equal(a.run('validateRows(rows).parsed.length'),8);
 const svg=a.run('currentSVG');assert.equal((svg.match(/<rect /g)||[]).length,9);assert.match(svg,/>100<\/text>/);assert.match(svg,/fill="#2c638d"/);
 assert.match(svg,/width="21\.428571428571427"/);assert.match(svg,/>400<\/text>/);
});
test('numbers, missing prices, duplicates and highlight totals are validated',()=>{
 const a=app();for(const value of ['1,000','-1.5','.25','0'])assert.equal(a.run(`Number.isFinite(numeric(${JSON.stringify(value)}))`),true);
 for(const value of ['1,2','Infinity','1e5','abc','9007199254740992'])assert.equal(a.run(`Number.isNaN(numeric(${JSON.stringify(value)}))`),true);
 a.run("rows=[{price:'10',ask:'-2',askHighlight:'1',bid:'1',bidHighlight:'2'},{price:'10',ask:'',askHighlight:'',bid:'',bidHighlight:''}];update()");
 assert.equal(a.run('valid'),false);assert.equal(a.registry.get('save-svg').disabled,true);
 a.run("rows=[{price:'',ask:'1',askHighlight:'',bid:'',bidHighlight:''}];update()");assert.equal(a.run('valid'),false);
 a.run("rows=[emptyRow()];update()");assert.equal(a.run('valid'),false);
});
test('prices sort descending, price-only rows survive, and SVG escapes headings',()=>{
 const a=app();a.run("rows=[{price:'-0.5',ask:'',askHighlight:'',bid:'',bidHighlight:''},{price:'10.5',ask:'1000000000',askHighlight:'10',bid:'',bidHighlight:''}];settings.askTitle='<script>& \"test\"';update()");
 const svg=a.run('currentSVG');assert.equal(a.run('valid'),true);assert.match(svg,/&lt;script&gt;&amp;/);assert.ok(svg.indexOf('>10.5</text>')<svg.indexOf('>-0.5</text>'));
 assert.equal((svg.match(/<rect /g)||[]).length,3);
 a.run("settings.background='transparent';update()");assert.doesNotMatch(a.run('currentSVG'),/fill="#ffffff"/);
 a.run("rows=[{price:'1',ask:'',askHighlight:'',bid:'2',bidHighlight:'2'}];update()");assert.equal(a.run('valid'),true);
});
test('Excel paste appends rows, validates values and refuses column overflow atomically',async()=>{
 const a=app();const input=a.inputs()[0];
 await a.event('rows','paste',{target:input,clipboardData:{getData:()=> '105\t100\t10\t20\t\t\t\r\n104\t\t\t\t300\t25\t50\r\n'},preventDefault(){}});
 assert.equal(a.run('rows[0].price'),'105');assert.equal(a.run('rows[1].bidHighlight'),'50');
 const before=a.run('JSON.stringify(rows)');await a.event('rows','paste',{target:a.inputs()[6],clipboardData:{getData:()=> '1\t2'},preventDefault(){}});
 assert.equal(a.run('JSON.stringify(rows)'),before);assert.match(a.registry.get('status').textContent,/Too many columns/);
 await a.event('add-row','click');assert.equal(a.run('rows.length'),9);
 const last=a.registry.get('rows').children.at(-1).children.at(-1).children[0];last.click();assert.equal(a.run('rows.length'),8);
});
test('settings reject bad bounds and retain a usable SVG after correction',()=>{
 const a=app();a.run("settings.fontSize='';update()");assert.equal(a.run('valid'),false);
 a.run("settings.fontSize='72';settings.barWidth=600;settings.askTitle='Long heading'.repeat(40);update()");assert.equal(a.run('valid'),true);assert.ok(a.run('size.width')>2000);
});
test('fractional quantities use the same full-width scale and all-zero rows are valid',()=>{
 const a=app();a.run("rows=[{price:'1',ask:'.5',askHighlight:'.1',bid:'.25',bidHighlight:''}];update()");
 assert.match(a.run('currentSVG'),/width="120" height="28" fill="#ff0000"/);
 assert.match(a.run('currentSVG'),/width="75" height="28" fill="#00b3df"/);
 a.run("rows=[{price:'0',ask:'0',askHighlight:'0',bid:'0',bidHighlight:'0'}];update()");assert.equal(a.run('valid'),true);assert.doesNotMatch(a.run('currentSVG'),/NaN/);
});
test('inner and outer regions face the right direction on both sides',()=>{
 const a=app();a.run("rows=[{price:'100',ask:'100',askInnerHighlight:'20',askHighlight:'30',bid:'100',bidInnerHighlight:'20',bidHighlight:'30'}];update()");
 assert.equal(a.run('valid'),true);
 const rects=[...a.run('currentSVG').matchAll(/<rect x="([\d.e+-]+)" y="([\d.e+-]+)" width="([\d.e+-]+)" height="([\d.e+-]+)" fill="(#[a-f0-9]+)"\/>/g)].map(m=>({x:+m[1],width:+m[3],color:m[5]}));
 const find=color=>rects.find(rect=>rect.color===color);
 const askInner=find('#ff8080'),askOuter=find('#b91c1c'),askNormal=find('#ff0000');
 const bidInner=find('#80d9ef'),bidOuter=find('#2c638d'),bidNormal=find('#00b3df');
 assert.equal(askInner.width,30);assert.equal(askOuter.width,45);assert.equal(askNormal.width,75);
 assert.equal(askOuter.x+askOuter.width,askNormal.x);assert.equal(askNormal.x+askNormal.width,askInner.x);
 assert.equal(bidInner.x+bidInner.width,bidNormal.x);assert.equal(bidNormal.x+bidNormal.width,bidOuter.x);
 a.run("rows[0].askInnerHighlight='80';update()");assert.equal(a.run('valid'),false);
 assert.equal(a.run("validateRows(rows).errors.has('0:askInnerHighlight')"),true);assert.equal(a.run("validateRows(rows).errors.has('0:askHighlight')"),true);
});
test('dotted white regions use selected colors and survive transparent export and JSON reload',async()=>{
 const a=app();a.run("rows=[{price:'100',ask:'100',askInnerHighlight:'20',askHighlight:'30',bid:'100',bidInnerHighlight:'20',bidHighlight:'30'}];settings.background='transparent';for(const [part] of PAINTS)settings[part+'Style']='dotted';update()");
 const svg=a.run('currentSVG');assert.equal((svg.match(/stroke-dasharray="1 4"/g)||[]).length,6);
 assert.equal((svg.match(/fill="#ffffff"/g)||[]).length,6);assert.match(svg,/stroke="#ff8080"/);assert.match(svg,/stroke="#2c638d"/);
 assert.equal(a.run("validateSettings({...settings,bidStyle:'bad'})"),'Invalid fill style.');
 await a.event('save-project','click');const json=await a.urls.at(-1).text();a.run('rows=structuredClone(SAMPLE);settings={...DEFAULTS};update()');
 await a.event('file-input','change',{target:{files:[{size:json.length,text:async()=>json}],value:'test'}});
 assert.equal(a.run('settings.askInnerHighlightStyle'),'dotted');assert.equal(a.run('rows[0].bidInnerHighlight'),'20');assert.equal(a.run('currentSVG'),svg);
 if(process.env.REGIONS_TEST_OUTPUT)fs.writeFileSync(process.env.REGIONS_TEST_OUTPUT,svg);
});
test('version 1 projects preserve their colors and outer highlights',async()=>{
 const a=app();const json=JSON.stringify({version:1,rows:[{price:'100',ask:'50',askHighlight:'10',bid:'80',bidHighlight:'20'}],settings:{askHighlightColor:'#a52e36'}});
 await a.event('file-input','change',{target:{files:[{size:json.length,text:async()=>json}],value:'test'}});
 assert.equal(a.run('valid'),true);assert.equal(a.run('rows[0].askInnerHighlight'),'');assert.equal(a.run('rows[0].askHighlight'),'10');
 assert.equal(a.run('settings.askHighlightColor'),'#a52e36');assert.equal(a.run('settings.bidInnerHighlightStyle'),'solid');
});
test('SVG and JSON download the real payload and JSON round trips',async()=>{
 const a=app();await a.event('save-svg','click');assert.equal(a.downloads[0].download,'order-book.svg');assert.match(await a.urls.at(-1).text(),/<svg/);
 await a.event('save-project','click');const json=await a.urls.at(-1).text();assert.equal(JSON.parse(json).version,3);
 a.run("rows=[emptyRow()];update()");await a.event('file-input','change',{target:{files:[{size:json.length,text:async()=>json}],value:'test'}});
 assert.equal(a.run('rows.length'),8);assert.equal(a.run('valid'),true);
 const before=a.run('JSON.stringify(rows)');await a.event('file-input','change',{target:{files:[{size:2,text:async()=>'{'}],value:'test'}});
 assert.equal(a.run('JSON.stringify(rows)'),before);assert.match(a.registry.get('status').textContent,/Could not open project/);
});
test('PNG/copy pipelines, missing permissions and oversized raster fallback',async()=>{
 const a=app();await a.event('save-png','click');assert.equal(a.downloads[0].download,'order-book.png');assert.equal(a.urls.at(-1).type,'image/png');
 await a.event('copy-image','click');assert.equal(a.clipboard.length,1);assert.match(a.registry.get('status').textContent,/Image copied/);
 a.run('navigator.clipboard.write=async()=>{throw new Error("denied")}');await a.event('copy-image','click');assert.match(a.registry.get('status').textContent,/Save PNG/);assert.equal(a.registry.get('save-png').disabled,false);
 a.run('navigator.clipboard=undefined');await a.event('copy-image','click');assert.match(a.registry.get('status').textContent,/unavailable/);
 await assert.rejects(a.run("pngBlob(currentSVG,20000,20000)"),/too large/);
});
test('SVG output has positive geometry and bounded text for edge cases',()=>{
 const a=app();for(const expr of ["settings.fontSize=12;settings.barHeight=8;settings.rowGap=0", "settings.fontSize=72;settings.barHeight=100;settings.rowGap=80", "settings.bidTitle='Buy\\norders\\nheading';settings.askTitle='Sell'"]){
  a.run(expr+';update()');const svg=a.run('currentSVG');const width=a.run('size.width'),height=a.run('size.height');
  for(const match of svg.matchAll(/<text x="([\d.e+-]+)" y="([\d.e+-]+)"/g)){assert.ok(+match[1]>0&&+match[1]<width);assert.ok(+match[2]>0&&+match[2]<height);}
  assert.doesNotMatch(svg,/NaN|undefined/);
 }
 if(process.env.SVG_TEST_OUTPUT)fs.writeFileSync(process.env.SVG_TEST_OUTPUT,a.run('buildSVG(validateRows(SAMPLE).parsed,DEFAULTS).svg'));
});
test('custom quantity labels replace numeric text without changing bar geometry',()=>{
 const a=app(),before=a.run('currentSVG'),beforeRects=[...before.matchAll(/<rect[^>]+>/g)].map(m=>m[0]);
 a.run("rows[0].askLabel='ABC';rows[0].askLabelColor='#d97706';rows[4].bidLabel='XYZ';rows[4].bidLabelColor='#7c3aed';update()");
 const svg=a.run('currentSVG');assert.match(svg,/text-anchor="end" fill="#d97706">ABC<\/text>/);assert.match(svg,/text-anchor="start" fill="#7c3aed">XYZ<\/text>/);
 assert.doesNotMatch(svg,/>300<\/text>|>400<\/text>/);assert.deepEqual([...svg.matchAll(/<rect[^>]+>/g)].map(m=>m[0]),beforeRects);
 a.run("rows[0].askLabel='   ';update()");assert.match(a.run('currentSVG'),/fill="#d97706">300<\/text>/);
});
test('long labels expand the figure and escape markup, including label-only rows',()=>{
 const a=app(),originalWidth=a.run('size.width');
 a.run("rows[0].askLabel='<tag>& '.repeat(20);rows[3].bidLabel='No orders';update()");
 assert.equal(a.run('valid'),true);assert.ok(a.run('size.width')>originalWidth);assert.match(a.run('currentSVG'),/&lt;tag&gt;&amp;/);assert.match(a.run('currentSVG'),/>No orders<\/text>/);
 assert.doesNotMatch(a.run('currentSVG'),/<tag>/);
 a.run("rows[0].askLabelColor='red';update()");assert.equal(a.run('valid'),false);
 a.run("rows[0].askLabelColor='#111111';rows[0].askLabel='a'.repeat(201);update()");assert.equal(a.run('valid'),false);
});
test('quantity label controls edit individual rows and persist through project round trips',async()=>{
 const a=app();const controls=a.registry.get('label-rows').children[0].children;
 const label=controls[1].children[0],color=controls[2].children[0];label.value='Custom quantity';color.value='#123456';
 await a.event('label-rows','input',{target:label});await a.event('label-rows','input',{target:color});
 assert.equal(a.run('rows[0].askLabel'),'Custom quantity');assert.match(a.run('currentSVG'),/fill="#123456">Custom quantity/);
 await a.event('save-project','click');const json=await a.urls.at(-1).text();const before=a.run('currentSVG');
 a.run('rows=structuredClone(SAMPLE);renderTable();update()');await a.event('file-input','change',{target:{files:[{size:json.length,text:async()=>json}],value:'test'}});
 assert.equal(a.run('currentSVG'),before);assert.equal(a.run('rows[0].askLabelColor'),'#123456');
 const legacy=JSON.parse(json);legacy.version=2;for(const row of legacy.rows)for(const key of ['askLabel','askLabelColor','bidLabel','bidLabelColor'])delete row[key];
 const text=JSON.stringify(legacy);await a.event('file-input','change',{target:{files:[{size:text.length,text:async()=>text}],value:'test'}});
 assert.equal(a.run('valid'),true);assert.equal(a.run('rows[0].askLabel'),'');assert.equal(a.run('rows[0].askLabelColor'),'#111111');
});
