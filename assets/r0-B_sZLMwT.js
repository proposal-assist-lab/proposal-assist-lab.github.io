import{J as pt}from"./jszip.min-CG32BIOO.js";const ht=4*1024*1024;function ft(){return new Promise((t,e)=>Office.context.document.getFileAsync(Office.FileType.Compressed,{sliceSize:ht},n=>n.status===Office.AsyncResultStatus.Succeeded?t(n.value):e(n.error)))}function mt(t,e){return new Promise((n,s)=>t.getSliceAsync(e,o=>o.status===Office.AsyncResultStatus.Succeeded?n(o.value):s(o.error)))}async function vt(t){const e=performance.now(),n=await ft();try{const s=[];let o=0;for(let i=0;i<n.sliceCount;i++){const c=await mt(n,i),g=new Uint8Array(c.data);s.push(g),o+=g.byteLength,t==null||t(i+1,n.sliceCount)}const a=new Uint8Array(o);let r=0;for(const i of s)a.set(i,r),r+=i.byteLength;return{bytes:a,fetchMs:performance.now()-e,sliceCount:n.sliceCount}}finally{n.closeAsync()}}function gt(t){let e="";for(let s=0;s<t.length;s+=32768)e+=String.fromCharCode(...t.subarray(s,s+32768));return btoa(e)}const N="http://schemas.openxmlformats.org/presentationml/2006/main",Y="http://schemas.openxmlformats.org/drawingml/2006/main",$t="http://schemas.openxmlformats.org/officeDocument/2006/relationships",wt="http://schemas.openxmlformats.org/package/2006/relationships";async function Q(t){const e=typeof t=="string"?new TextEncoder().encode(t):t,n=await crypto.subtle.digest("SHA-256",e);return Array.from(new Uint8Array(n),s=>s.toString(16).padStart(2,"0")).join("")}function tt(t,e){if(e.startsWith("/"))return e.slice(1);const n=(t+e).split("/"),s=[];for(const o of n)o===".."?s.pop():o!=="."&&o!==""&&s.push(o);return s.join("/")}function F(t){const e=t.lastIndexOf("/");return e<0?"":t.slice(0,e+1)}function yt(t){return`${F(t)}_rels/${t.slice(F(t).length)}.rels`}function et(t){const e=new Map;if(!t)return e;for(const n of Array.from(t.getElementsByTagNameNS(wt,"Relationship")))e.set(n.getAttribute("Id")??"",{type:n.getAttribute("Type")??"",target:n.getAttribute("Target")??""});return e}function bt(t){return Array.from(t.getElementsByTagNameNS(Y,"p")).map(n=>{let s="";for(const o of Array.from(n.children))o.localName==="r"||o.localName==="fld"?s+=Array.from(o.getElementsByTagNameNS(Y,"t")).map(a=>a.textContent??"").join(""):o.localName==="br"&&(s+="\v");return s}).join("\r")}function St(t){for(const e of Array.from(t.getElementsByTagNameNS(N,"sp"))){const n=e.getElementsByTagNameNS(N,"ph")[0],s=n==null?void 0:n.getAttribute("type");if(s==="title"||s==="ctrTitle"){const o=e.getElementsByTagNameNS(N,"txBody")[0];return o?bt(o):""}}return""}const Pt=new Set(["nvGrpSpPr","grpSpPr","extLst"]);function At(t){const e=t.getElementsByTagNameNS(N,"spTree")[0];if(!e)return[];const n=[];for(const s of Array.from(e.children)){if(Pt.has(s.localName))continue;const o=s.getElementsByTagNameNS(N,"cNvPr")[0];o&&n.push({id:o.getAttribute("id")??"",name:o.getAttribute("name")??""})}return n}async function Ct(t){var b;const e=performance.now(),n=await pt.loadAsync(t),s=new DOMParser,o=async l=>{var v;return await((v=n.file(l))==null?void 0:v.async("string"))??null},a=async l=>{const v=await o(l);return v==null?null:s.parseFromString(v,"application/xml")},r=await a("ppt/presentation.xml");if(!r)throw new Error("ppt/presentation.xml が見つかりません");const i=et(await a("ppt/_rels/presentation.xml.rels")),c=new Map,g=[],O=Array.from(r.getElementsByTagNameNS(N,"sldId"));for(const[l,v]of O.entries()){const k=v.getAttribute("id")??"",U=v.getAttributeNS($t,"id")??"",x=i.get(U);if(!x)continue;const M=tt("ppt/",x.target),$=await o(M)??"",C=s.parseFromString($,"application/xml");let W="";const ut=et(await a(yt(M)));for(const Z of ut.values()){if(!Z.type.endsWith("/slideLayout"))continue;const B=tt(F(M),Z.target);if(!c.has(B)){const _=await a(B);c.set(B,((b=_==null?void 0:_.getElementsByTagNameNS(N,"cSld")[0])==null?void 0:b.getAttribute("name"))??"")}W=c.get(B)??""}g.push({index:l,sldId:k,path:M,hash:"",layoutName:W,title:St(C),topShapes:At(C),allShapeCount:C.getElementsByTagNameNS(N,"cNvPr").length,xml:$})}const d=performance.now()-e,f=performance.now();for(const l of g)l.hash=await Q(l.xml);const u=new Map;for(const l of Object.values(n.files))l.dir||u.set(l.name,await Q(await l.async("uint8array")));const y=performance.now()-f;return{byteLength:t.byteLength,slides:g,partHashes:u,parseMs:d,hashMs:y}}function kt(t,e){const n=[];let s=0;const o=Math.min(t.slides.length,e.slides.length);for(let c=0;c<o;c++)t.slides[c].hash===e.slides[c].hash&&t.slides[c].sldId===e.slides[c].sldId?s++:n.push({index:c,path:e.slides[c].path});const a=[],r=[],i=[];for(const[c,g]of e.partHashes)t.partHashes.has(c)?t.partHashes.get(c)!==g&&a.push(c):r.push(c);for(const c of t.partHashes.keys())e.partHashes.has(c)||i.push(c);return{slideSame:s,slideChanged:n,slideCountChanged:t.slides.length!==e.slides.length,partsChanged:a,partsAdded:r,partsRemoved:i}}const ot=[{id:"V1",title:"バージョンと PowerPointApi 1.8",fallback:"表セルの置換とF-05を延期。表は検出のみ"},{id:"V2",title:"getFileAsync（未保存反映・取得時間・ハッシュ一致）",fallback:"①保存を促す／Office.jsのみ構成　③正規化処理を追加"},{id:"V3",title:"XMLのID ↔ Office.js の ID 対応",fallback:"順序・名前・位置で照合／Office.jsのみ構成"},{id:"V4",title:"垂直タブの表現と getSubstring の書式保持",fallback:"全体を書き戻して書式を再適用"},{id:"V5",title:"createPresentation と ダウンロード",fallback:"新しいウィンドウで開く方式のみ"},{id:"V6",title:"document.settings の保存と再読込",fallback:"スライドのタグで代替"},{id:"V7",title:"github.io と api.anthropic.com への接続",fallback:"配信元の見直し（Azure SWA など）"},{id:"V8",title:"共有フォルダカタログの登録",fallback:"サイドロードで運用"},{id:"V9",title:"本番CSPでCSP違反が出ないか",fallback:"違反ログから必要ドメインをCSPに追加"},{id:"V10",title:"開発環境（git / Node.js / Claude Code / localhost証明書）",fallback:"開発用ブランチを別Pagesに配信"}];function at(){return`${Office.context.partitionKey??""}r0.results`}function Nt(){try{return JSON.parse(localStorage.getItem(at())??"{}")}catch{return{}}}function it(t){try{localStorage.setItem(at(),JSON.stringify(t))}catch{}}function Ot(t,e){const n=[`## R0 検証結果（${new Date().toLocaleString("ja-JP")}）`,"",e,"","| # | 検証項目 | 結果 | 詳細 |","| --- | --- | --- | --- |"];for(const s of ot){const o=t[s.id],a=((o==null?void 0:o.detail)??"").replace(/\|/g,"｜").replace(/\r?\n/g,"<br>");n.push(`| ${s.id} | ${s.title} | ${(o==null?void 0:o.status)??"未実施"} | ${a} |`)}return n.join(`
`)}var st;const T=((st=document.querySelector('meta[name="csp-label"]'))==null?void 0:st.content)??"?",Mt="https://proposal-assist-lab.github.io/",ct="r0.v6";let P={},L=null;const p={},S={},m=t=>document.getElementById(t),E=t=>`${(t/1e3).toFixed(2)}秒`,X=t=>`${(t/1024/1024).toFixed(2)}MB`,K=t=>t.replace(/[&<>"]/g,e=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;"})[e]);function h(t,e,n=!1){const s=m(t);s.textContent=n?`${s.textContent}
${e}`:e}function q(t){if(t&&typeof t=="object"){const e=t;return`エラー: ${e.name??""} ${e.code??""} ${e.message??""}${e.debugInfo?`
${JSON.stringify(e.debugInfo)}`:""}`}return`エラー: ${String(t)}`}function A(t,e,n){const s=m(t);s.addEventListener("click",async()=>{s.disabled=!0,h(e,"実行中…");try{await n()}catch(o){h(e,q(o),!0)}finally{s.disabled=!1}})}function w(t,e,n){P[t]={status:e,detail:n,updatedAt:new Date().toISOString()},it(P),G()}function V(t){return Office.context.requirements.isSetSupported("PowerPointApi",t)}async function R(t){const e=await vt((s,o)=>h(t,`取得中… ${s}/${o} スライス`));h(t,`解析中…（${X(e.bytes.byteLength)}）`);const n=await Ct(e.bytes);return{...e,deck:n}}function Et(){m("app").innerHTML=`
  <header>
    <h1>提案書アシスト R0検証</h1>
    <div class="meta">CSP: <b class="csp-${T}">${T==="prod"?"本番用":"開発用"}</b> ／ ${K(location.origin)}</div>
  </header>

  <section>
    <h2>結果一覧</h2>
    <table id="summary"></table>
    <div class="row">
      <button id="copy-md">結果をMarkdownでコピー</button>
      <button id="reset" class="secondary">結果をリセット</button>
    </div>
    <textarea id="md" rows="6" readonly placeholder="ここに結果のMarkdownが出ます。Claude Codeに貼り付けてください"></textarea>
  </section>

  <section>
    <h2>V1 バージョンと要求セット</h2>
    <p class="note">起動時に自動判定します。更新チャネルは［ファイル］→［アカウント］→［Microsoft 365 とOffice の更新プログラム］付近の表示で確認し、詳細に追記してください。</p>
    <button id="v1-run">再判定</button>
    <pre id="v1-out"></pre>
  </section>

  <section>
    <h2>V2 getFileAsync</h2>
    <p class="note">新テンプレートの実資料（約12MB）の<b>複製</b>を開いて実行してください。</p>
    <div class="step"><b>② 取得時間</b>：資料を取得・解析し、比較の基準として保存します。</div>
    <button id="v2-fetch">取得＋解析（基準を保存）</button>
    <div class="step"><b>③ ハッシュ一致</b>：資料を<u>変更せずに</u>再取得し、基準と比較します。</div>
    <button id="v2-compare">再取得して比較</button>
    <div class="step"><b>① 未保存の編集</b>：選択中のスライドに目印のテキストボックスを追加し（保存しない）、取得したファイルに含まれるかを調べた後、目印を削除します。</div>
    <button id="v2-unsaved">未保存反映テスト（自動）</button>
    <details><summary>手動で確認する場合</summary>
      <p class="note">スライドに任意の文字列を入力し、保存せずに下で検索します。</p>
      <div class="row"><input id="v2-marker" placeholder="検索する文字列" /><button id="v2-manual">検索</button></div>
    </details>
    <pre id="v2-out"></pre>
  </section>

  <section>
    <h2>V3 ID の対応付け</h2>
    <p class="note">XMLの sldId / cNvPr id と、Office.js の slide.id / shape.id を全スライドで突き合わせます。</p>
    <button id="v3-run">対応付けを確認</button>
    <pre id="v3-out"></pre>
  </section>

  <section>
    <h2>V4 垂直タブと getSubstring</h2>
    <p class="note">本文スライドのタイトル（例：「1. ご提案の趣旨⏎1-1. 弊社の理解」）を<b>コピーしたスライド上で</b>選択して実行してください。</p>
    <button id="v4-inspect">選択図形のテキストを調べる</button>
    <div class="row">
      <label>開始 <input id="v4-start" type="number" min="0" value="0" class="num" /></label>
      <label>文字数 <input id="v4-len" type="number" min="1" value="1" class="num" /></label>
      <label>置換後 <input id="v4-repl" value="9-9. " class="short" /></label>
    </div>
    <button id="v4-replace" class="warn">番号部分を置換（選択図形を書き換えます）</button>
    <pre id="v4-out"></pre>
  </section>

  <section>
    <h2>V5 スナップショットの復元とダウンロード</h2>
    <p class="note">直近に取得した資料（未取得なら今取得）を使います。</p>
    <div class="row">
      <button id="v5-open">新しいウィンドウで開く</button>
      <button id="v5-dl">バックアップをダウンロード</button>
    </div>
    <div class="row">ダウンロード結果：
      <button id="v5-dl-ok" class="secondary">保存できた</button>
      <button id="v5-dl-ng" class="secondary">できなかった</button>
    </div>
    <pre id="v5-out"></pre>
  </section>

  <section>
    <h2>V6 document.settings</h2>
    <p class="note">書き込み後、資料を<b>保存→閉じる→開き直す→このパネルを開く</b>と自動判定します。</p>
    <button id="v6-write">書き込んで保存</button>
    <pre id="v6-out"></pre>
  </section>

  <section>
    <h2>V7 外部への接続</h2>
    <div class="row"><input id="v7-pages" value="${Mt}" /></div>
    <button id="v7-run">接続を確認</button>
    <pre id="v7-out"></pre>
  </section>

  <section>
    <h2>V8 共有フォルダカタログ（手動）</h2>
    <ol class="note">
      <li>共有フォルダ（例：C:\\AddinCatalog）を作り、［プロパティ］→［共有］で共有する</li>
      <li>manifest.prod.xml をそのフォルダに置く</li>
      <li>PowerPoint：［ファイル］→［オプション］→［トラストセンター］→［トラストセンターの設定］→［信頼できるアドインカタログ］に <code>\\\\PC名\\共有名</code> を追加し「メニューに表示する」をオン</li>
      <li>PowerPointを再起動 →［ホーム］→［アドイン］→［その他のアドイン］→［共有フォルダー］に表示されれば OK</li>
    </ol>
    <p class="note">結果は上の結果一覧で選んでください。</p>
  </section>

  <section>
    <h2>V9 CSP 違反</h2>
    <p class="note">本番用CSPは <code>npm run build</code> → <code>npm run preview</code> で確認できます（同じ localhost:3000 で配信）。V2〜V7 を一通り操作した後に判定してください。</p>
    <button id="v9-judge">判定</button>
    <pre id="v9-out"></pre>
  </section>

  <section>
    <h2>V10 開発環境（手動）</h2>
    <p class="note">このパネルが localhost から表示できていれば、証明書と読み込みは OK です。git・Claude Code の導入状況とあわせて、結果一覧で選んでください。</p>
  </section>
  `}function G(){const t=m("summary");t.innerHTML="<tr><th>#</th><th>項目</th><th>結果</th></tr>"+ot.map(e=>{const n=P[e.id],s=(n==null?void 0:n.status)??"未実施",o=["未実施","OK","NG","要確認"].map(a=>`<option${a===s?" selected":""}>${a}</option>`).join("");return`<tr class="st-${s}">
        <td>${e.id}</td>
        <td>${K(e.title)}
          <textarea class="detail" data-id="${e.id}" rows="1" placeholder="詳細・メモ">${K((n==null?void 0:n.detail)??"")}</textarea>
          ${s==="NG"?`<div class="fallback">→ ${K(e.fallback)}</div>`:""}</td>
        <td><select data-id="${e.id}">${o}</select></td>
      </tr>`}).join(""),t.querySelectorAll("select").forEach(e=>e.addEventListener("change",()=>{var s;const n=e.dataset.id;w(n,e.value,((s=P[n])==null?void 0:s.detail)??"")})),t.querySelectorAll("textarea.detail").forEach(e=>e.addEventListener("change",()=>{var s;const n=e.dataset.id;w(n,((s=P[n])==null?void 0:s.status)??"未実施",e.value)}))}function jt(){const t=Office.context.diagnostics;return`環境：${t.platform} / PowerPoint ${t.version} / CSP ${T} / ${location.origin}`}function nt(){var i,c;const t=Office.context.diagnostics,e=["1.1","1.2","1.3","1.4","1.5","1.6","1.7","1.8","1.9","1.10"],n=e.filter(V),s=V("1.8"),o=[`ホスト: ${t.host}  プラットフォーム: ${t.platform}`,`バージョン: ${t.version}`,`PowerPointApi: ${e.map(g=>`${g}${V(g)?"○":"×"}`).join(" ")}`,`1.8: ${s?"対応":"非対応"}`].join(`
`);h("v1-out",o);const r=(c=(((i=P.V1)==null?void 0:i.detail)??"").match(/チャネル[:：].*/))==null?void 0:c[0];w("V1",s?"OK":"NG",`${t.platform} ${t.version} / 最大 PowerPointApi ${n.at(-1)??"なし"}${r?` / ${r}`:" / チャネル：未記入"}`)}function D(){const t=[];p.unsaved!==void 0&&t.push(`①未保存反映 ${p.unsaved?"含まれる":"含まれない"}`),p.fetchMs!==void 0&&t.push(`②取得 ${E(p.fetchMs)}（${X(p.bytes??0)}）`),p.hashEqual!==void 0&&t.push(`③スライドXML ${p.hashEqual?"一致":"不一致あり"}`);const e=p.unsaved!==void 0&&p.fetchMs!==void 0&&p.hashEqual!==void 0;let n="要確認";p.unsaved===!1?n="NG":e&&(n=p.hashEqual&&p.fetchMs<=1e4?"OK":"要確認"),w("V2",n,t.join(" / ")+(e?"":"（未完了の項目あり）"))}async function Lt(){const t=await R("v2-out");L={deck:t.deck,bytes:t.bytes,fetchMs:t.fetchMs},p.fetchMs=t.fetchMs,p.bytes=t.bytes.byteLength;const e=t.deck,n=[`サイズ: ${X(t.bytes.byteLength)}（${t.sliceCount} スライス）`,`取得: ${E(t.fetchMs)}  解析: ${E(e.parseMs)}  ハッシュ: ${E(e.hashMs)}`,`合計: ${E(t.fetchMs+e.parseMs+e.hashMs)}（目標：取得と解析で15秒以内／再取得は10秒以内）`,`スライド: ${e.slides.length}枚  パーツ: ${e.partHashes.size}`,"","先頭10枚（sldId / レイアウト / タイトル）:",...e.slides.slice(0,10).map(s=>`  ${s.index+1}. [${s.sldId}] ${s.layoutName} ／ ${s.title.replace(/\v/g,"⏎").replace(/\r/g,"¶")}`)];h("v2-out",n.join(`
`)),D()}async function xt(){if(!L)throw new Error("先に「取得＋解析（基準を保存）」を実行してください");const t=await R("v2-out"),e=kt(L.deck,t.deck);p.hashEqual=e.slideChanged.length===0&&!e.slideCountChanged,p.fetchMs=Math.max(p.fetchMs??0,t.fetchMs);const n=[`再取得: ${E(t.fetchMs)}`,`スライドXML: 一致 ${e.slideSame} ／ 不一致 ${e.slideChanged.length}${e.slideCountChanged?"（枚数が変化）":""}`,...e.slideChanged.slice(0,20).map(s=>`  × ${s.index+1}枚目 ${s.path}`),`その他のパーツ: 変化 ${e.partsChanged.length} / 追加 ${e.partsAdded.length} / 削除 ${e.partsRemoved.length}`,...[...e.partsChanged,...e.partsAdded.map(s=>`${s}（追加）`),...e.partsRemoved.map(s=>`${s}（削除）`)].slice(0,30).map(s=>`  ・${s}`),"",p.hashEqual?"→ 変更しなければスライドXMLは一致します（正規化は不要）":"→ 不一致あり。変動している要素を調べ、正規化処理が必要です"];h("v2-out",n.join(`
`)),D()}async function rt(t){return(await R("v2-out")).deck.slides.filter(n=>n.xml.includes(t)).map(n=>n.index+1)}async function Vt(){if(!V("1.5"))throw new Error("PowerPointApi 1.5 未対応のため自動テストできません。手動で確認してください");const t=`R0MARK${Date.now()}`;let e="",n="";await PowerPoint.run(async o=>{const a=o.presentation.getSelectedSlides();a.load("items/id"),await o.sync();const r=a.items[0]??o.presentation.slides.getItemAt(0);r.load("id");const i=r.shapes.addTextBox(t,{left:10,top:10,width:200,height:30});i.load("id"),await o.sync(),e=i.id,n=r.id});let s=[];try{s=await rt(t)}finally{await PowerPoint.run(async o=>{o.presentation.slides.getItem(n).shapes.getItem(e).delete(),await o.sync()})}p.unsaved=s.length>0,h("v2-out",[`目印「${t}」をスライド ${n} に追加 → 取得 → 削除しました`,s.length?`→ 取得したファイルの ${s.join(", ")} 枚目に含まれていました（未保存の編集も取得される）`:"→ 取得したファイルに含まれていません（保存しないと反映されない）"].join(`
`),!0),D()}async function It(){const t=m("v2-marker").value.trim();if(!t)throw new Error("検索する文字列を入力してください");const e=await rt(t);p.unsaved=e.length>0,h("v2-out",e.length?`「${t}」は ${e.join(", ")} 枚目に含まれていました`:`「${t}」は見つかりませんでした`,!0),D()}async function Tt(){const e=(await R("v3-out")).deck.slides;h("v3-out","Office.js 側を読み込み中…");const n=[];await PowerPoint.run(async d=>{const f=d.presentation.slides;f.load("items/id"),await d.sync();const u=f.items.map(y=>{const b=y.shapes;return b.load("items/id,items/name"),b});await d.sync(),f.items.forEach((y,b)=>n.push({id:y.id,shapes:u[b].items.map(l=>({id:l.id,name:l.name}))}))});let s=0,o=0,a=0,r=0;const i=[],c=Math.max(e.length,n.length);for(let d=0;d<c;d++){const f=e[d],u=n[d];if(!f||!u){i.push(`${d+1}枚目: 片方にしか存在しない`);continue}u.id.split("#")[0]===f.sldId?s++:i.push(`${d+1}枚目: sldId=${f.sldId} / slide.id=${u.id}`);const b=new Set(f.topShapes.map(l=>l.id));o+=u.shapes.length,u.shapes.forEach((l,v)=>{var k;b.has(l.id)?a++:i.length<30&&i.push(`${d+1}枚目: shape.id=${l.id}「${l.name}」がXMLの直下図形にない`),((k=f.topShapes[v])==null?void 0:k.id)===l.id&&r++}),f.topShapes.length!==u.shapes.length&&i.length<30&&i.push(`${d+1}枚目: 直下図形の数 XML=${f.topShapes.length} / Office.js=${u.shapes.length}`)}const g=e.slice(0,2).map((d,f)=>{const u=n[f];return[`  ${f+1}枚目 sldId=${d.sldId}  slide.id=${u==null?void 0:u.id}`,`    XML   : ${d.topShapes.map(y=>y.id).join(", ")}`,`    Office: ${u==null?void 0:u.shapes.map(y=>y.id).join(", ")}`].join(`
`)}),O=s===c&&a===o;h("v3-out",[`スライド: ${s}/${c} 一致（slide.id の「#」より前と sldId を比較）`,`図形: ID一致 ${a}/${o}　順序も一致 ${r}/${o}`,"","例:",...g,i.length?`
不一致（先頭30件）:
${i.slice(0,30).map(d=>`  ・${d}`).join(`
`)}`:""].join(`
`)),w("V3",O?"OK":s===c?"要確認":"NG",`スライド ${s}/${c}、図形ID ${a}/${o}、順序 ${r}/${o}`)}const I={11:"VT",13:"CR",10:"LF",9:"TAB",8232:"LS",8233:"PS"};function H(t){return Array.from(t).map(e=>{const n=e.codePointAt(0);return I[n]||n<32?`⟨${I[n]??"U+"+n.toString(16).padStart(4,"0")}⟩`:e}).join("")}const Rt=/^[0-9０-９]+(?:[-‐－―][0-9０-９]+)*[.．]\s?/;function Bt(t){const e=t.split(new RegExp("(?<=[\\v\\r\\n\\u2028\\u2029])"));let n=0;const s=[];for(const o of e){const a=o.match(Rt);a&&s.push({start:n,len:a[0].length}),n+=o.length}return s.at(-1)??null}async function j(t,e,n,s){const o=[];for(let a=n;a<n+s;a++){const r=e.getSubstring(a,1).font;r.load("name,size,bold,italic,underline,color"),o.push(r)}return await t.sync(),o.map(a=>`${a.name}|${a.size}|${a.bold?"B":""}${a.italic?"I":""}|${a.underline}|${a.color}`)}async function lt(t){const e=t.presentation.getSelectedShapes();if(e.load("items/id,items/name"),await t.sync(),!e.items.length)throw new Error("図形（タイトル）を1つ選択してください");return e.items[0]}async function Kt(){await PowerPoint.run(async t=>{const e=await lt(t),n=e.textFrame.textRange;n.load("text"),await t.sync();const s=n.text,o=new Set(Array.from(s).map(i=>i.codePointAt(0)).filter(i=>I[i]||i<32)),a=Bt(s);a&&(m("v4-start").value=String(a.start),m("v4-len").value=String(a.len)),h("v4-out",[`図形: ${e.name}（id=${e.id}）  文字数: ${s.length}`,`テキスト: ${H(s)}`,`制御文字: ${o.size?[...o].map(i=>`${I[i]??"?"}(U+${i.toString(16).padStart(4,"0")})`).join(", "):"なし"}`,a?`番号部分を検出: 開始=${a.start} 文字数=${a.len}「${s.substr(a.start,a.len)}」`:"番号部分は検出できませんでした"].join(`
`));const r=o.has(11)?"VT(U+000B)":o.size?[...o].map(i=>I[i]??i).join(","):"改行なし";w("V4","要確認",`改行の表現: ${r}（getSubstring 未確認）`)})}async function Ht(){const t=Number(m("v4-start").value),e=Number(m("v4-len").value),n=m("v4-repl").value;await PowerPoint.run(async s=>{var x,M;const a=(await lt(s)).textFrame.textRange;a.load("text"),await s.sync();const r=a.text;if(t+e>r.length)throw new Error("範囲が文字数を超えています");const i=Math.min(r.length-t-e,300),c=await j(s,a,0,Math.min(t,300)),g=await j(s,a,t,e),O=await j(s,a,t+e,i);a.getSubstring(t,e).text=n,await s.sync(),a.load("text"),await s.sync();const d=a.text,f=await j(s,a,0,Math.min(t,300)),u=await j(s,a,t,n.length),y=await j(s,a,t+n.length,i),b=r.slice(0,t)+n+r.slice(t+e),l=d===b,v=[];O.forEach(($,C)=>$!==y[C]&&v.push(C)),f.forEach(($,C)=>$!==c[C]&&v.push(-1-C));const k=v.length===0;h("v4-out",[`変更前: ${H(r)}`,`変更後: ${H(d)}`,`テキスト: ${l?"期待どおり":`想定と異なる（期待: ${H(b)}）`}`,`前後の書式: ${k?`保持された（${c.length+O.length}文字を1文字ずつ比較）`:`変化あり（${v.length}文字）`}`,...v.slice(0,10).map($=>$>=0?`  後ろ${$}文字目: ${O[$]} → ${y[$]}`:`  前${-1-$}文字目: ${c[-1-$]} → ${f[-1-$]}`),`置換前の番号部分の書式: ${[...new Set(g)].join(" / ")}`,`置換後の番号部分の書式: ${[...new Set(u)].join(" / ")}`,"","※ 元に戻すには Ctrl+Z"].join(`
`));const U=(((M=(x=P.V4)==null?void 0:x.detail.match(/改行の表現: [^（]*/))==null?void 0:M[0])??"改行の表現: 未確認").trim();w("V4",l&&k?"OK":"NG",`${U} / getSubstring置換: テキスト${l?"○":"×"} 書式保持${k?"○":"×"}`)})}async function dt(t){if(L)return L.bytes;const e=await R(t);return L={deck:e.deck,bytes:e.bytes,fetchMs:e.fetchMs},e.bytes}function z(){const t=[];S.open!==void 0&&t.push(`createPresentation ${S.open?"○":"×"}`),S.download!==void 0&&t.push(`ダウンロード ${S.download?"○":"×（新しいウィンドウ方式で運用）"}`);const e=S.open===!1?"NG":S.open&&S.download!==void 0&&S.download?"OK":"要確認";w("V5",e,t.join(" / "))}async function Dt(){if(!V("1.1"))throw new Error("PowerPointApi 1.1 未対応");const t=await dt("v5-out"),e=performance.now();try{await PowerPoint.createPresentation(gt(t)),S.open=!0,h("v5-out",`新しいウィンドウで開きました（${E(performance.now()-e)}）。枚数・タイトル・書式が元と同じか確認してください。`)}catch(n){throw S.open=!1,n}finally{z()}}async function Ut(){const t=await dt("v5-out"),e=new Blob([t],{type:"application/vnd.openxmlformats-officedocument.presentationml.presentation"}),n=URL.createObjectURL(e),s=new Date().toISOString().replace(/[-:]/g,"").replace("T","-").slice(0,15),o=document.createElement("a");o.href=n,o.download=`snapshot-${s}.pptx`,document.body.appendChild(o),o.click(),o.remove(),setTimeout(()=>URL.revokeObjectURL(n),6e4),h("v5-out",`snapshot-${s}.pptx のダウンロードを試みました。保存ダイアログが出たか・ファイルができたかを確認し、下のボタンで結果を記録してください。`)}function _t(){const t=Office.context.document.settings.get(ct);t!=null&&t.writtenAt?(h("v6-out",`資料に保存された値を読み込めました（書き込み日時: ${t.writtenAt}）`),sessionStorage.getItem("r0.v6.wroteThisSession")!=="1"&&w("V6","OK",`再読込で値を確認（書き込み ${t.writtenAt}）`)):h("v6-out","資料に保存された値はありません")}async function Ft(){const t=new Date().toLocaleString("ja-JP");Office.context.document.settings.set(ct,{writtenAt:t}),await new Promise((e,n)=>Office.context.document.settings.saveAsync(s=>s.status===Office.AsyncResultStatus.Succeeded?e():n(s.error)));try{sessionStorage.setItem("r0.v6.wroteThisSession","1")}catch{}h("v6-out",`書き込みました（${t}）。資料を保存→閉じる→開き直す→このパネルを開いてください。`),w("V6","要確認",`書き込み済み（${t}）。再読込待ち`)}async function qt(){const t=[];let e=!1;try{const o=await fetch("https://api.anthropic.com/v1/messages",{method:"POST",headers:{"content-type":"application/json","anthropic-version":"2023-06-01","anthropic-dangerous-direct-browser-access":"true"},body:"{}"});e=!0,t.push(`api.anthropic.com: 到達 OK（HTTP ${o.status}。キー未指定のため 401 は正常）`)}catch(o){t.push(`api.anthropic.com: 到達できず（ネットワーク／CORS／CSP）${q(o)}`)}let n=!1;const s=m("v7-pages").value.trim();try{const o=await fetch(s,{cache:"no-store"});n=!0,t.push(`${s}: 到達 OK（HTTP ${o.status}${o.status===404?"。Pagesが未公開の可能性。到達自体はできている":""}）`)}catch(o){try{await fetch(s,{mode:"no-cors",cache:"no-store"}),n=!0,t.push(`${s}: 到達 OK（応答の中身は読めない。Pagesが未公開の可能性。到達自体はできている）`)}catch{t.push(`${s}: 到達できず ${q(o)}`)}}h("v7-out",t.join(`
`)),w("V7",e&&n?"OK":"NG",t.join(" / "))}function J(){const t=window.__cspViolations??[];h("v9-out",[`現在のCSP: ${T==="prod"?"本番用":"開発用"}`,`違反: ${t.length}件`,...t.slice(0,30).map(e=>`  ・${e.directive} ← ${e.blockedURI} ${e.source}`)].join(`
`))}function Gt(){J();const t=window.__cspViolations??[];if(T!=="prod"){w("V9","要確認",`開発用CSPで違反${t.length}件。本番用（npm run build → npm run preview）で再判定が必要`);return}const e=[...new Set(t.map(n=>`${n.directive}:${n.blockedURI}`))];w("V9",t.length?"NG":"OK",t.length?`違反 ${t.length}件: ${e.join(", ")}`:"本番用CSPで違反なし")}Office.onReady(t=>{Et(),t.host!==Office.HostType.PowerPoint&&m("app").insertAdjacentHTML("afterbegin",'<p class="warn-box">PowerPoint 以外で開かれています</p>'),P=Nt(),G(),m("copy-md").addEventListener("click",async()=>{const e=Ot(P,jt()),n=m("md");n.value=e;try{await navigator.clipboard.writeText(e)}catch{n.select(),document.execCommand("copy")}}),m("reset").addEventListener("click",()=>{P={},it(P),G()}),m("v1-run").addEventListener("click",nt),A("v2-fetch","v2-out",Lt),A("v2-compare","v2-out",xt),A("v2-unsaved","v2-out",Vt),A("v2-manual","v2-out",It),A("v3-run","v3-out",Tt),A("v4-inspect","v4-out",Kt),A("v4-replace","v4-out",Ht),A("v5-open","v5-out",Dt),A("v5-dl","v5-out",Ut),m("v5-dl-ok").addEventListener("click",()=>(S.download=!0,z())),m("v5-dl-ng").addEventListener("click",()=>(S.download=!1,z())),A("v6-write","v6-out",Ft),A("v7-run","v7-out",qt),m("v9-judge").addEventListener("click",Gt),window.addEventListener("csp-violation",J),nt(),_t(),J(),location.hostname==="localhost"&&!P.V10&&w("V10","要確認","localhost からの読み込み・証明書は OK。git / Claude Code の導入は手動で確認")});
