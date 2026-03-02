"use strict";
// ════════════════════════════════════════════════════════════════════════════
// matcher.js — Matcher Engine: Web Worker + Similarity Algorithms + UI
// ════════════════════════════════════════════════════════════════════════════

// ════════════════════════════════════════════════════════════════════════════
// MATCHER ENGINE — Web Worker
// ════════════════════════════════════════════════════════════════════════════
const MATCHER_WORKER_SRC = `
"use strict";
const TL=[['ight','айт'],['tion','шн'],['ough','оф'],['sch','ш'],['tch','ч'],['all','ол'],['ing','инг'],['igh','ай'],['ull','ул'],['oor','ур'],['alk','ок'],['awn','он'],['sh','ш'],['ch','ч'],['zh','ж'],['kh','х'],['ph','ф'],['th','т'],['wh','в'],['ck','к'],['qu','кв'],['ts','ц'],['tz','ц'],['oo','у'],['ee','и'],['ea','и'],['ui','у'],['ew','ю'],['aw','о'],['ow','оу'],['oi','ой'],['oy','ой'],['ai','ей'],['ay','ей'],['au','о'],['ou','у'],['bb','бб'],['cc','кк'],['dd','дд'],['ff','фф'],['gg','гг'],['ll','лл'],['mm','мм'],['nn','нн'],['pp','пп'],['rr','рр'],['ss','сс'],['tt','тт'],['zz','цц'],['a','а'],['b','б'],['c','к'],['d','д'],['e','э'],['f','ф'],['g','г'],['h','х'],['i','и'],['j','дж'],['k','к'],['l','л'],['m','м'],['n','н'],['o','о'],['p','п'],['q','к'],['r','р'],['s','с'],['t','т'],['u','у'],['v','в'],['w','в'],['x','кс'],['y','й'],['z','з']];
const WORD_END=[[/блес$/,'блс'],[/тлес$/,'тлс'],[/плес$/,'плс'],[/лес$/,'лс'],[/([бвгджзклмнпрстфхцчшщ])ес$/,'$1с'],[/([бвгджзклмнпрстфхцчшщ])е$/,'$1']];
const STOP=new Set(['и','или','в','на','с','по','для','к','от','из','за','не','как','это','то','the','a','an','of','for','with','and','or','in','on','at','to','by']);
const UNIT_CANON=new Map([['г','г'],['г.','г'],['гр','г'],['гр.','г'],['грамм','г'],['граммов','г'],['грамма','г'],['g','г'],['gr','г'],['gramm','г'],['gram','г'],['мг','мг'],['мг.','мг'],['миллиграмм','мг'],['миллиграммов','мг'],['mg','мг'],['кг','кг'],['кг.','кг'],['килограмм','кг'],['килограммов','кг'],['кило','кг'],['kg','кг'],['мл','мл'],['мл.','мл'],['миллилитр','мл'],['миллилитров','мл'],['ml','мл'],['л','л'],['л.','л'],['литр','л'],['литров','л'],['литра','л'],['l','л'],['lt','л'],['ltr','л'],['шт','шт'],['шт.','шт'],['штука','шт'],['штук','шт'],['штуки','шт'],['pcs','шт'],['pc','шт'],['pcs.','шт']]);
const UNIT_CONV={'мг':{base:'г',factor:0.001},'г':{base:'г',factor:1},'кг':{base:'г',factor:1000},'мл':{base:'мл',factor:1},'л':{base:'мл',factor:1000}};
const ABBR_DICT=new Map([['уп','упаковка'],['упак','упаковка'],['упк','упаковка'],['нб','набор'],['нбр','набор'],['кор','коробка'],['ком','комплект'],['компл','комплект'],['кмп','комплект']]);
function normalizeUnits(s){return s.replace(/(\\d+(?:[.,]\\d+)?)\\s*([а-яёa-zA-Z]{1,12}\\.?)/gi,(m,num,unitStr)=>{const uk=unitStr.toLowerCase().replace(/\\.$/,'');const canon=UNIT_CANON.get(uk);if(!canon)return m;const conv=UNIT_CONV[canon];if(!conv)return num+canon+' ';const val=Math.round(parseFloat(num.replace(',','.'))*conv.factor*100000)/100000;return val+conv.base+' ';});}
function expandAbbr(tokens){const res=[];for(const t of tokens){const exp=ABBR_DICT.get(t);if(exp){for(const w of exp.split(' ')){if(w.length>1&&!STOP.has(w))res.push(w);}}else{res.push(t);}}return res;}
function translitWord(w){let s=w.toLowerCase(),out='',i=0;while(i<s.length){let hit=false;for(const[lat,cyr]of TL){if(s.startsWith(lat,i)){out+=cyr;i+=lat.length;hit=true;break;}}if(!hit){out+=s[i];i++;}}for(const[re,rep]of WORD_END)out=out.replace(re,rep);return out;}
const _VOW=new Set('аеёиоуыэюя'.split(''));const _isV=c=>_VOW.has(c);
function _rv(w){for(let i=0;i<w.length;i++)if(_isV(w[i]))return i+1;return w.length;}
function _r1(w){for(let i=1;i<w.length;i++)if(!_isV(w[i])&&_isV(w[i-1]))return i+1;return w.length;}
function _r2(w,r1){for(let i=r1+1;i<w.length;i++)if(!_isV(w[i])&&_isV(w[i-1]))return i+1;return w.length;}
function _strip(w,ss,f){const s=[...ss].sort((a,b)=>b.length-a.length);for(const x of s)if(w.endsWith(x)&&(w.length-x.length)>=f)return w.slice(0,-x.length);return null;}
function stemRu(word){if(!word||word.length<=2)return word;if(!/[а-яё]/.test(word))return word;const rv=_rv(word),r1=_r1(word),r2=_r2(word,r1);let w=word,r;r=_strip(w,['ившись','ивши','ывшись','ывши','ив','ыв'],rv);if(r!=null){w=r;}else{const pa=new Set(['а','я']);r=null;const pvf=['авшись','явшись','авши','явши','ав','яв'];for(const x of pvf.sort((a,b)=>b.length-a.length)){if(w.endsWith(x)&&pa.has(w[w.length-x.length-1]||'')){r=w.slice(0,-x.length);break;}}if(r!=null)w=r;}r=_strip(w,['ся','сь'],rv);if(r!=null)w=r;const ADJ=['ими','ыми','ией','ий','ый','ой','ей','ем','им','ым','ом','его','ого','ему','ому','ую','юю','ая','яя','ое','ее'];let adj=false;for(const s of [...ADJ].sort((a,b)=>b.length-a.length)){if(w.endsWith(s)&&(w.length-s.length)>=rv){w=w.slice(0,-s.length);adj=true;break;}}if(!adj){const VA=['ала','яла','али','яли','ало','яло','ана','яна','ает','яет','ают','яют','аешь','яешь','ай','яй','ал','ял','ать','ять'];const VF=['ила','ыла','ена','ейте','уйте','ите','или','ыли','ей','уй','ил','ыл','им','ым','ен','ило','ыло','ено','ует','уют','ит','ыт','ишь','ышь','ую','ю'];const pa=new Set(['а','я']);r=null;for(const x of VA.sort((a,b)=>b.length-a.length)){if(w.endsWith(x)&&(w.length-x.length)>=rv&&pa.has(w[w.length-x.length-1]||'')){r=w.slice(0,-x.length);break;}}if(r==null)r=_strip(w,VF,rv);if(r!=null){w=r;}else{const N=['иями','ями','ами','ией','ием','иям','ев','ов','ие','ье','еи','ии','ей','ой','ий','ям','ем','ам','ом','ях','ах','е','и','й','о','у','а','ь','ю','я'];r=_strip(w,N,rv);if(r!=null)w=r;}}if(w.endsWith('и')&&(w.length-1)>=rv)w=w.slice(0,-1);r=_strip(w,['ость','ост'],r2);if(r!=null)w=r;if(w.endsWith('нн'))w=w.slice(0,-1);if(w.endsWith('ь')&&(w.length-1)>=rv)w=w.slice(0,-1);return w.length>=2?w:word;}
// PKG_ABBR: раскрываем аббревиатуры упаковки ДО того как normalize() снесёт символы /\.
// Это позволяет bsynMap корректно сматчить "ст/б" и "стеклобанка" — оба станут одним токеном.
const PKG_ABBR=new Map([['ст/б','стеклобанка'],['с/б','стеклобанка'],['стб','стеклобанка'],['м/у','мягкаяупаковка'],['м/уп','мягкаяупаковка'],['ж/б','жестяная'],['д/п','дойпак'],['с/я','саше'],['ф/п','саше'],['пл/б','пластик'],['б/к','безконверта']]);
function pkgExpand(s){return s.split(' ').map(function(t){return PKG_ABBR.get(t)||t;}).join(' ');}
function preNorm(raw){let s=normalizeUnits(String(raw||''));s=s.replace(/([а-яёa-zA-Z])-([а-яёa-zA-Z])/gi,'$1 $2');s=s.toLowerCase();return pkgExpand(s);}
function normalize(raw){if(!raw)return '';let s=preNorm(raw);s=s.replace(/[^\\wа-яё0-9\\s]/gi,' ').replace(/[a-z]+/gi,m=>translitWord(m)).replace(/\\s+/g,' ').trim();const toks=s.split(' ').filter(w=>{if(STOP.has(w))return false;if(/^\\d+$/.test(w))return true;return w.length>=2||/^[а-яёa-z0-9]+$/i.test(w);});return expandAbbr(toks).map(t=>stemRu(t)).join(' ');}
// Применить словарь брендов: заменить синонимы на канонические формы
// applyBrandNorm: заменяет синонимы (в т.ч. мультитокенные) на канонические формы.
// Приоритет поиска: trigram → bigram → unigram.
// Это позволяет корректно обрабатывать бренды из 2-3 слов ("kit kat", "alpen gold", "carte noire").
function applyBrandNorm(norm,bsynMap){
  if(!bsynMap||!bsynMap.size)return norm;
  const toks=norm.split(' ');
  const out=[];
  let i=0;
  while(i<toks.length){
    let matched=false;
    // Trigram (3 токена): "carte du noir" и т.п.
    if(!matched&&i+2<toks.length){
      const key=toks[i]+' '+toks[i+1]+' '+toks[i+2];
      const mapped=bsynMap.get(key);
      if(mapped){for(const t of mapped.split(' '))out.push(t);i+=3;matched=true;}
    }
    // Bigram (2 токена): "kit kat", "alpen gold", "earl grey" и т.п.
    if(!matched&&i+1<toks.length){
      const key=toks[i]+' '+toks[i+1];
      const mapped=bsynMap.get(key);
      if(mapped){for(const t of mapped.split(' '))out.push(t);i+=2;matched=true;}
    }
    // Unigram (1 токен)
    if(!matched){out.push(bsynMap.get(toks[i])||toks[i]);i++;}
  }
  return out.join(' ');
}
function trigrams(s){const st=new Set(),p='#'+s+'#';for(let i=0;i<p.length-2;i++)st.add(p.slice(i,i+3));return st;}
function triSim(a,b){if(!a||!b)return 0;const ta=trigrams(a),tb=trigrams(b);let n=0;for(const g of ta)if(tb.has(g))n++;return n*2/(ta.size+tb.size);}
function lcsLen(a,b){const A=a.split(' ').slice(0,80),B=b.split(' ').slice(0,80);let prev=new Uint8Array(B.length+1),curr=new Uint8Array(B.length+1);for(let i=0;i<A.length;i++){for(let j=0;j<B.length;j++)curr[j+1]=A[i]===B[j]?prev[j]+1:Math.max(curr[j],prev[j+1]);[prev,curr]=[curr,prev];curr.fill(0);}return prev[B.length];}
function lcsSim(a,b){if(!a||!b)return(a===b)?1:0;const wa=Math.min(a.split(' ').length,80),wb=Math.min(b.split(' ').length,80);return 2*lcsLen(a,b)/(wa+wb);}

// extractWeightNums: извлекает ТОЛЬКО числа при единицах веса/объёма (г, мл, кг, л, мг).
// Числа вроде "30шт" или "24бл" — логистика, не идентификатор товара — сюда НЕ попадают.
// Это устраняет ложный штраф 0.82 при сравнении "30шт/20бл" vs "30шт/24бл" одного товара.
function extractWeightNums(s){var r=[];s.replace(/(\\d+(?:[.,]\\d+)?)\\s*(?:г|мл|кг|л|мг)/g,function(m,n){r.push(parseFloat(n.replace(',','.')));});return r;}
function buildIDF(items){const df=new Map();for(const it of items){const seen=new Set(it._norm.split(' ').filter(t=>t.length>0));for(const t of seen)df.set(t,(df.get(t)||0)+1);}const N=items.length,idf=new Map();for(const[t,freq]of df)idf.set(t,Math.log((N+1)/(freq+1))+1);return idf;}
function wTokenSim(a,b,idf){const A=a.split(' '),B=b.split(' ');let wA=0,wB=0;const mA=new Map(),mB=new Map();for(const t of A){const w=idf.get(t)||1;mA.set(t,(mA.get(t)||0)+w);wA+=w;}for(const t of B){const w=idf.get(t)||1;mB.set(t,(mB.get(t)||0)+w);wB+=w;}if(!wA||!wB)return 0;let inter=0;for(const[t,wa]of mA)if(mB.has(t))inter+=Math.min(wa,mB.get(t));return 2*inter/(wA+wB);}
function buildInvIdx(items){const idx=new Map();items.forEach((it,i)=>{for(const tok of it._norm.split(' ')){if(tok.length<1)continue;if(!idx.has(tok))idx.set(tok,[]);idx.get(tok).push(i);}});return idx;}
function getCandidates(norm,idx,limit,idf){const toks=norm.split(' ').filter(t=>t.length>0);const scores=new Map();for(const t of toks){const w=idf?(idf.get(t)||1):1;for(const id of(idx.get(t)||[]))scores.set(id,(scores.get(id)||0)+w);}const ranked=[...scores.entries()].sort((a,b)=>b[1]-a[1]).slice(0,limit).map(([id])=>id);if(idf){const IDF_RARE=3.0;const rareToks=toks.filter(t=>(idf.get(t)||0)>IDF_RARE).slice(0,6);const rankedSet=new Set(ranked);for(const t of rareToks){for(const id of(idx.get(t)||[]).slice(0,25)){if(!rankedSet.has(id)){ranked.push(id);rankedSet.add(id);}}}}return ranked;}
function calcSim(name1,name2,idf,bsynMap,bantMap){const pre1=preNorm(name1),pre2=preNorm(name2);const nums1=extractWeightNums(pre1),nums2=extractWeightNums(pre2);let n1=normalize(name1),n2=normalize(name2);if(!n1||!n2)return 0;
// Применяем словарь брендов/упаковок и детектируем совпадения через синонимы.
// Синоним сработал, если строка изменилась И новый токен теперь встречается в другой строке.
// Важно: bigram-замена меняет количество токенов, поэтому сравниваем множества, а не индексы.
let synBonus=0;
if(bsynMap&&bsynMap.size){
  const n1b=n1,n2b=n2;
  n1=applyBrandNorm(n1,bsynMap);n2=applyBrandNorm(n2,bsynMap);
  if(n1!==n1b||n2!==n2b){
    // Хотя бы одна строка изменилась — проверяем появились ли новые общие токены
    const setAfter1=new Set(n1.split(' ')),setAfter2=new Set(n2.split(' '));
    const setBefore1=new Set(n1b.split(' ')),setBefore2=new Set(n2b.split(' '));
    // Новые токены в n1 (которых не было до замены) и они есть в n2 — синоним сработал
    for(const t of setAfter1){if(!setBefore1.has(t)&&setAfter2.has(t)){synBonus=0.20;break;}}
    // Новые токены в n2 и они есть в n1
    if(!synBonus){for(const t of setAfter2){if(!setBefore2.has(t)&&setAfter1.has(t)){synBonus=0.20;break;}}}
  }
}
// Антонимы брендов — немедленный возврат 0.
// Проверяем и одиночные токены, и биграммы (для мультитокенных canonical типа "kit kat", "alpen gold").
if(bantMap&&bantMap.size){
  const toks1=n1.split(' '),toks2=n2.split(' ');
  // Собираем все антонимные множества из обеих строк (unigram + bigram ключи)
  function _getBantSets(toks){
    const sets=[];
    for(const t of toks){if(bantMap.has(t))sets.push(bantMap.get(t));}
    for(let j=0;j<toks.length-1;j++){const bi=toks[j]+' '+toks[j+1];if(bantMap.has(bi))sets.push(bantMap.get(bi));}
    return sets;
  }
  // Строим расширенный набор токенов для проверки (unigram + bigram)
  function _tokSet(toks){
    const s=new Set(toks);
    for(let j=0;j<toks.length-1;j++)s.add(toks[j]+' '+toks[j+1]);
    return s;
  }
  const anti1=_getBantSets(toks1),set2=_tokSet(toks2);
  for(const anti of anti1){for(const t of set2){if(anti.has(t))return 0;}}
  const anti2=_getBantSets(toks2),set1=_tokSet(toks1);
  for(const anti of anti2){for(const t of set1){if(anti.has(t))return 0;}}
}
// Числовой штраф — СМЯГЧЁН: числа могут обозначать не только вес/объём,
// но и номер паллета, количество в блоке, артикул и т.д.
// Строгий штраф применяется только когда числа явно несовместимы.
// Если синоним уже дал бонус — numFactor не может упасть ниже 0.88.
let numFactor=1.0,numBonus=0;if(nums1.length>0&&nums2.length>0){const s1=new Set(nums1.map(n=>n.toFixed(5))),s2=new Set(nums2.map(n=>n.toFixed(5)));const common=[...s1].filter(x=>s2.has(x)).length;const total=new Set([...s1,...s2]).size;const ratio=common/total;if(ratio===0)numFactor=0.82;else if(ratio<0.5)numFactor=0.90;else numBonus=ratio*0.12;if(synBonus>0)numFactor=Math.max(numFactor,0.88);}
const wc1=n1.split(' ').length,wc2=n2.split(' ').length;const lenRatio=Math.min(wc1,wc2)/Math.max(wc1,wc2);const lenPenalty=lenRatio<0.33?0.6:lenRatio<0.5?0.82:1.0;
// n1s/n2s вычисляем ПЕРЕД tri, чтобы triSim тоже брал максимум по отсортированному варианту.
// Это фиксит "Мыло туалетное" vs "Туалетное мыло" — оба дают одинаковую строку после sort().
const n1s=n1.split(' ').sort().join(' ');const n2s=n2.split(' ').sort().join(' ');const tri=Math.max(triSim(n1,n2),triSim(n1s,n2s));const lcss=Math.max(lcsSim(n1,n2),lcsSim(n1s,n2s));const wTok=idf?wTokenSim(n1,n2,idf):lcss;const len=(n1.split(' ').length+n2.split(' ').length)/2;const wTri=len<=3?0.25:0.35,wLcs=len<=3?0.20:0.25,wTok_=len<=3?0.55:0.40;
// synBonus добавляется ПОСЛЕ numFactor (числа не должны уничтожать синонимное совпадение)
let score=(tri*wTri+lcss*wLcs+wTok*wTok_+numBonus)*numFactor*lenPenalty+synBonus;
const fw1=n1.split(' ')[0],fw2=n2.split(' ')[0];if(fw1&&fw2&&fw1.length>2&&fw2.length>2){const fw1Idf=idf?(idf.get(fw1)||1):1,fw2Idf=idf?(idf.get(fw2)||1):1;const isRare1=fw1Idf>1.5,isRare2=fw2Idf>1.5;
// ⚙ СНИЖЕН бонус: первое слово часто бывает названием категории/типа («молоко», «сок», «масло»),
// совпадение само по себе мало о чём говорит. Бонус 0.04 вместо 0.10, штраф за несовпадение сохранён.
if(fw1===fw2&&isRare1)score=Math.min(1,score+0.04);else if(fw1!==fw2&&isRare1&&isRare2)score*=0.82;}
let _fs=Math.min(100,Math.round(score*100));
// 100% только если оригинальные названия буквально одинаковы (без учёта регистра и пробелов).
// Нормализатор слишком агрессивен — разные названия могут стать идентичными после стемминга.
// Это не означает 100% сходство с точки зрения пользователя.
if(_fs>=100){const _o1=String(name1||'').toLowerCase().replace(/\s+/g,' ').trim();
const _o2=String(name2||'').toLowerCase().replace(/\s+/g,' ').trim();
if(_o1!==_o2)_fs=99;}
return _fs;}
const BC_COLS_W=['штрихкод','штрих-код','barcode','шк','ean','код'];
const NAME_COLS_W=['название','наименование','name','товар','продукт','наим'];
function findCol(data,variants){if(!data?.length)return null;const cols=Object.keys(data[0]);return cols.find(c=>variants.some(v=>c.toLowerCase().includes(v)))??cols[0];}
self.onmessage=function({data}){
  if(data.type!=='run')return;
  const{db,priceFiles,brandDB}=data;
  const activePairs=[],knownPairs=[];

  // Build brand synonym/antonym maps for calcSim
  // Pass 1: build bsynMap (synonym->canonical) and canon2syns (canonical->all variants).
  // Все синонимы (в т.ч. мультитокенные) хранятся как ключи в bsynMap — строки с пробелами.
  // applyBrandNorm теперь поддерживает bigram/trigram lookup и найдёт их.
  // Antonyms are NOT touched here — otherwise bsynMap.set(an,an) corrupts synonym mappings
  const bsynMap=new Map(),bantMap=new Map();
  const canon2syns=new Map();
  for(const[canon,val]of Object.entries(brandDB||{})){
    const cNorm=normalize(canon);if(!cNorm)continue;
    if(!bsynMap.has(cNorm))bsynMap.set(cNorm,cNorm);
    // synSet содержит ВСЕ нормализованные варианты (canonical + все синонимы),
    // включая мультитокенные — они нужны bantMap для корректной симметрии.
    const synSet=new Set([cNorm]);
    for(const s of(val.synonyms||[])){const sn=normalize(s);if(sn){if(!bsynMap.has(sn))bsynMap.set(sn,cNorm);synSet.add(sn);}}
    canon2syns.set(cNorm,synSet);
  }
  // Диагностика отключена — мультитокенные canonical без якоря не блокируют работу
  // (bantMap с bigram lookup корректно обрабатывает их через applyBrandNorm)
  for(const[cNorm,syns]of canon2syns){
    if(cNorm.includes(' ')){
      // No-op: multi-token canonicals are handled via bigram bantMap lookup
    }
  }
  // Pass 2: build bantMap with full synonym expansion and symmetry.
  // bantMap хранит ключи и как однотокенные ("нескаф"), так и мультитокенные ("карт нуар").
  // calcSim проверяет bantMap с поддержкой bigram — мультитокенные ключи будут найдены.
  // Симметрия: если A антоним B — bantMap[B] тоже получит все варианты A (автоматически).
  for(const[canon,val]of Object.entries(brandDB||{})){
    const cNorm=normalize(canon);if(!cNorm)continue;
    if(!(val.antonyms||[]).length)continue;
    const antiCanons=new Set();
    for(const a of(val.antonyms||[])){const an=normalize(a);if(an)antiCanons.add(bsynMap.get(an)||an);}
    const antiSet=bantMap.get(cNorm)||new Set();
    for(const ac of antiCanons){const syns=canon2syns.get(ac);if(syns){for(const s of syns)antiSet.add(s);}else antiSet.add(ac);}
    if(antiSet.size)bantMap.set(cNorm,antiSet);
    // Симметрия: добавляем все варианты cNorm в bantMap каждого антонима
    const mySyns=canon2syns.get(cNorm)||new Set([cNorm]);
    for(const ac of antiCanons){
      if(!bantMap.has(ac))bantMap.set(ac,new Set());
      for(const s of mySyns)bantMap.get(ac).add(s);
    }
  }

  // Build DB lookup: bc -> canonical key, canonical key -> display name
  const bc2key=new Map(),bc2name=new Map();
  for(const[key,val]of Object.entries(db)){
    const name=Array.isArray(val)?(val[0]||String(key)):'';
    bc2key.set(String(key),String(key));bc2name.set(String(key),name);
    if(Array.isArray(val))val.slice(1).forEach(s=>{s=String(s).trim();if(s){bc2key.set(s,String(key));bc2name.set(s,name);}});
  }

  // Build items from price files — deduplicate by (fi, bc): keep first name per bc per file
  const items=[];
  const seenBcPerFile=new Map();
  for(let fi=0;fi<priceFiles.length;fi++){
    const f=priceFiles[fi];
    const bcC=findCol(f.data,BC_COLS_W),nmC=findCol(f.data,NAME_COLS_W);
    if(!bcC||!nmC)continue;
    for(const row of f.data){
      const bc=String(row[bcC]||'').trim(),nm=String(row[nmC]||'').trim();
      if(!bc||!nm)continue;
      const fk=fi+'\x00'+bc;
      if(seenBcPerFile.has(fk))continue;
      seenBcPerFile.set(fk,true);
      items.push({bc,name:nm,fi,file:f.name,_norm:normalize(nm)});
    }
  }

  if(!items.length){self.postMessage({type:'done',activePairs:[],knownPairs:[]});return;}

  // Build jsonItems from DB for index (only entries with a name)
  const jsonItems=Object.entries(db).map(([key,val])=>{
    const name=Array.isArray(val)?(val[0]||String(key)):String(key);
    return {bc:String(key),name,fi:-1,file:'JSON',_norm:normalize(name)};
  }).filter(it=>it._norm.length>0);

  // allItems for indexing and IDF: price items + json items
  const allItems=[...items,...jsonItems];
  const invIdx=buildInvIdx(allItems);
  const idf=buildIDF(allItems);

  // ── ЛИМИТ КАНДИДАТОВ ─────────────────────────────────────────────────────
  // Сколько кандидатов проверять через calcSim для каждого товара.
  // Больше = лучше охват совпадений, но медленнее.
  // Минимум 200, максимум без ограничений (allItems.length).
  // Можно изменить число 200 или убрать Math.min для полного перебора.
  const effectiveLimit = Math.min(200, allItems.length);
  // ─────────────────────────────────────────────────────────────────────────

  // seen pairs by sorted bc pair to avoid duplicates
  const seenPairs=new Set();

  for(let i=0;i<items.length;i++){
    if(i%200===0)self.postMessage({type:'progress',pct:Math.round(i/items.length*100)});
    const a=items[i];
    const aCanon=bc2key.get(a.bc);
    const aInDB=!!aCanon;
    const aKey=aCanon||null;
    // Use DB name for matching if available (canonical)
    const matchNorm=(aInDB&&bc2name.get(a.bc))?normalize(bc2name.get(a.bc)):a._norm;

    for(const id of getCandidates(matchNorm,invIdx,effectiveLimit,idf)){
      const b=allItems[id];
      if(!b)continue;
      if(b.bc===a.bc)continue;
      // Skip same file (fi=-1 JSON items are always allowed as targets)
      if(b.fi>=0&&b.fi===a.fi)continue;
      // Skip if already in same DB group
      const bCanon=bc2key.get(b.bc);
      if(aInDB&&bCanon&&aKey===bCanon)continue;
      // Skip already-committed pair
      const pk=a.bc<b.bc?a.bc+'\x01'+b.bc:b.bc+'\x01'+a.bc;
      if(seenPairs.has(pk))continue;

      const sim=calcSim(a.name,b.name,idf,bsynMap,bantMap);
      if(sim<52)continue;

      seenPairs.add(pk);

      const bInDB=!!bCanon;
      const pair={
        bc1:a.bc,name1:a.name,file1:a.file,
        bc2:b.bc,name2:b.name,file2:b.file,
        sim,aInDB,bInDB,
        aKey:aKey||a.bc,bKey:bCanon||b.bc
      };
      if(aInDB&&bInDB&&aKey===bCanon)knownPairs.push(pair);
      else activePairs.push(pair);
    }
  }

  activePairs.sort((a,b)=>b.sim-a.sim);
  knownPairs.sort((a,b)=>b.sim-a.sim);
  self.postMessage({type:'progress',pct:100});
  self.postMessage({type:'done',activePairs,knownPairs,allItems:items});
};
`;

// ════════════════════════════════════════════════════════════════════════════
// MATCHER STATE
// ════════════════════════════════════════════════════════════════════════════
let _matchActivePairs = [];
let _matchKnownPairs  = [];
let _matchAllItems    = []; // все товары из всех прайсов (для поиска без пары)
let _matchCurrentView = 'all';
let _matchHideKnown   = false;
let _matchWorker      = null;
let _matchWorkerUrl   = null;
let _matchPending     = null; // { pair, idx, view }
let _matchBgResult    = null; // buffered result while tab was hidden

// ── Выбор файлов для матчинга ──────────────────────────────────────────────
// Set имён файлов, которые ОТКЛЮЧЕНЫ (не участвуют в матчинге)
let _matcherDisabledFiles = new Set();

// Рендерит чипы файлов в панели матчера
function matcherFileChipsRender() {
  const panel = document.getElementById('matcherFilesPanel');
  const wrap  = document.getElementById('matcherFileChips');
  if (!panel || !wrap) return;

  const hasFiles = typeof allFilesData !== 'undefined' && allFilesData.length > 0;
  const hasJson = typeof jeDB !== 'undefined' && Object.keys(jeDB).length > 0;

  if (!hasFiles && !hasJson) {
    panel.style.display = 'none';
    return;
  }

  panel.style.display = 'flex';

  // Update JSON row
  if (typeof window._matcherUpdateJsonInfo === 'function') window._matcherUpdateJsonInfo();

  if (!hasFiles) {
    wrap.innerHTML = '';
    return;
  }

  wrap.innerHTML = allFilesData.map(f => {
    const off = _matcherDisabledFiles.has(f.fileName);
    const label = f.fileName.length > 35 ? f.fileName.slice(0, 33) + '…' : f.fileName;
    const safeTitle = (f.fileName + (off ? ' (отключён)' : ' — включён в матчинг'))
      .replace(/"/g, '&quot;');
    const safeLabel = label.replace(/</g, '&lt;').replace(/>/g, '&gt;');
    // data-mf-name вместо onclick — безопасно для любых имён файлов
    return `<span class="mf-chip${off ? ' mf-off' : ''}" data-mf-name="${encodeURIComponent(f.fileName)}" title="${safeTitle}"><span class="mf-chip-icon">${off ? '✕' : '✓'}</span>${safeLabel}</span>`;
  }).join('');
}

// Переключает файл вкл/выкл в матчинге
function matcherToggleFile(fileName) {
  if (_matcherDisabledFiles.has(fileName)) {
    _matcherDisabledFiles.delete(fileName);
  } else {
    // Нельзя отключить все файлы — должен остаться хотя бы один
    const enabledCount = (typeof allFilesData !== 'undefined' ? allFilesData.length : 0)
      - _matcherDisabledFiles.size - 1;
    if (enabledCount < 1) {
      showToast('Должен быть включён хотя бы один файл', 'warn');
      return;
    }
    _matcherDisabledFiles.add(fileName);
  }
  matcherFileChipsRender();
  // Немедленно обновляем таблицу — скрываем/показываем пары с этим файлом
  if (typeof renderMatcherTable === 'function') renderMatcherTable();
}

// Event delegation для чипов файлов — один обработчик вместо inline onclick на каждом чипе
document.addEventListener('click', function(e) {
  const chip = e.target.closest('[data-mf-name]');
  if (!chip) return;
  matcherToggleFile(decodeURIComponent(chip.dataset.mfName));
});


// Apply buffered matcher result when tab becomes visible again
document.addEventListener('visibilitychange', function() {
  if (!document.hidden && _matchBgResult) {
    const { activePairs, knownPairs, allItems, btn } = _matchBgResult;
    _matchBgResult = null;
    _matchActivePairs = activePairs;
    _matchKnownPairs  = knownPairs;
    _matchAllItems    = allItems || [];
    btn.disabled = false; btn.textContent = '▶ Запустить матчинг';
    document.getElementById('matcherProgress').style.display = 'none';
    updateMatcherStats();
    setMatchView('all');
    document.getElementById('matcherStats').style.display = 'flex';
    document.getElementById('matcherSearchInp').disabled = false;
    const _msr3=document.getElementById('matcherSearchRow');if(_msr3)_msr3.style.display='';
    document.getElementById('matcherHideKnownBtn').style.display = '';
    showToast('Матчинг завершён в фоне: ' + _matchActivePairs.length + ' пар найдено', 'ok');
  }
});

function getMatcherDB() {
  // Check if JSON is disabled via the matcher toggle
  const chk = document.getElementById('matcherJsonEnabled');
  if (chk && !chk.checked) return {};
  // jeDB is the single source of truth — barcodeAliasMap is derived from it
  return jeDB;
}

function getMatcherPriceFiles() {
  // Collect from allFilesData, excluding files disabled by user via chip panel
  if (typeof allFilesData === 'undefined' || !allFilesData.length) return [];
  return allFilesData
    .filter(f => !_matcherDisabledFiles.has(f.fileName))
    .map(f => ({ name: f.fileName, data: f.data }));
}

function runMatcher() {
  const files = getMatcherPriceFiles();
  if (!files.length) {
    showToast('Сначала загрузите прайсы на вкладке «Мониторинг»', 'warn');
    return;
  }
  const btn = document.getElementById('matcherRunBtn');
  btn.disabled = true; btn.textContent = '⏳ Анализ...';

  document.getElementById('matcherProgress').style.display = '';
  document.getElementById('matcherProgressLbl').textContent = 'Анализирую прайсы...';
  document.getElementById('matcherProgressFill').style.width = '0%';
  document.getElementById('matcherStats').style.display = 'none';
  document.getElementById('matcherTableWrap').style.display = 'none';
  document.getElementById('matcherEmpty').style.display = 'none';

  _matchActivePairs = []; _matchKnownPairs = [];

  if (_matchWorker) _matchWorker.terminate();
  if (_matchWorkerUrl) URL.revokeObjectURL(_matchWorkerUrl);
  const blob = new Blob([MATCHER_WORKER_SRC], { type: 'application/javascript' });
  _matchWorkerUrl = URL.createObjectURL(blob);
  _matchWorker = new Worker(_matchWorkerUrl);

  _matchWorker.onmessage = function({ data }) {
    if (data.type === 'progress') {
      if (!document.hidden) {
        document.getElementById('matcherProgressFill').style.width = data.pct + '%';
      }
    } else if (data.type === 'done') {
      if (document.hidden) {
        // Tab is hidden — buffer result, apply on visibilitychange
        _matchBgResult = { activePairs: data.activePairs, knownPairs: data.knownPairs, allItems: data.allItems || [], btn };
      } else {
        _matchActivePairs = data.activePairs;
        _matchKnownPairs  = data.knownPairs;
        _matchAllItems    = data.allItems || [];
        btn.disabled = false; btn.textContent = '▶ Запустить матчинг';
        document.getElementById('matcherProgress').style.display = 'none';
        updateMatcherStats();
        setMatchView('all');
        document.getElementById('matcherStats').style.display = 'flex';
        document.getElementById('matcherSearchInp').disabled = false;
        const _msr2=document.getElementById('matcherSearchRow');if(_msr2)_msr2.style.display='';
        document.getElementById('matcherHideKnownBtn').style.display = '';
        showToast('Матчинг завершён: ' + _matchActivePairs.length + ' пар найдено', 'ok');
      }
    }
  };
  _matchWorker.onerror = err => {
    console.error('Matcher worker error:', err);
    btn.disabled = false; btn.textContent = '▶ Запустить матчинг';
    document.getElementById('matcherProgress').style.display = 'none';
    showToast('Ошибка матчинга: ' + err.message, 'err');
  };

  const db = getMatcherDB();
  const brandDB = typeof _brandDB !== 'undefined' ? _brandDB : {};
  _matchWorker.postMessage({ type: 'run', db, priceFiles: files, brandDB });
}

function updateMatcherStats() {
  const active = _matchHideKnown
    ? _matchActivePairs.filter(p => !(p.aInDB || p.bInDB))
    : _matchActivePairs;
  document.getElementById('ms-all').textContent   = active.length;
  document.getElementById('ms-high').textContent  = active.filter(p => p.sim >= 80).length;
  document.getElementById('ms-mid').textContent   = active.filter(p => p.sim >= 52 && p.sim < 80).length;
    const badge = document.getElementById('matcherBadge');
  if (active.length > 0) { badge.textContent = active.length + ' пар'; badge.style.display = ''; }
  else badge.style.display = 'none';
}

function setMatchView(v) {
  _matchCurrentView = v;
  document.querySelectorAll('.mstat').forEach(s => s.classList.toggle('active', s.dataset.mv === v));
  renderMatcherTable();
}

function getMatchViewList() {
  if (_matchCurrentView === 'high') return _matchActivePairs.filter(p => p.sim >= 80);
  if (_matchCurrentView === 'mid')  return _matchActivePairs.filter(p => p.sim >= 52 && p.sim < 80);
  return _matchActivePairs;
}

let _matchRenderedPairs = [];

function renderMatcherTable() {
  const q = (document.getElementById('matcherSearchInp').value || '').toLowerCase().trim();
  const wrap = document.getElementById('matcherTableWrap');
  const empty = document.getElementById('matcherEmpty');

    // ── Normal pair list view ────────────────────────────────────────
  let list = getMatchViewList().slice();
  if (_matchHideKnown && (_matchCurrentView === 'all' || _matchCurrentView === 'high' || _matchCurrentView === 'mid')) {
    list = list.filter(r => !(r.aInDB || r.bInDB));
  }
  // Скрываем пары, у которых хотя бы один файл отключён через чипы
  if (_matcherDisabledFiles.size > 0) {
    list = list.filter(r =>
      !_matcherDisabledFiles.has(r.file1) && !_matcherDisabledFiles.has(r.file2)
    );
  }
  if (q) {
    list = list.filter(r =>
      (r.name1||'').toLowerCase().includes(q) || (r.name2||'').toLowerCase().includes(q) ||
      (r.bc1||'').includes(q) || (r.bc2||'').includes(q) ||
      (r.file1||'').toLowerCase().includes(q) || (r.file2||'').toLowerCase().includes(q)
    );
  }

  _matchRenderedPairs = list;

  if (!list.length) {
    wrap.style.display = 'none'; empty.style.display = '';
    empty.querySelector('h3').textContent = 'Нет совпадений';
    empty.querySelector('p').textContent = q ? 'Попробуйте изменить поисковый запрос' : 'Нажмите «Запустить матчинг» для поиска похожих товаров';
    return;
  }
  wrap.style.display = ''; empty.style.display = 'none';

  // ── Virtual scroll for pairs ─────────────────────────────────────
  const PAIR_H = 74; // px per pair (2 rows)
  const OVERSCAN = 8; // extra pairs to render above/below viewport
  const COL_SPAN = 5; // colspan for spacer rows (7 cols total, but first 2 are rowspan)

  function _mvsRenderMatcherRows() {
    const scrollTop = wrap.scrollTop;
    const viewH = wrap.clientHeight || 500;
    const total = list.length;
    const start = Math.max(0, Math.floor(scrollTop / PAIR_H) - OVERSCAN);
    const end   = Math.min(total, Math.ceil((scrollTop + viewH) / PAIR_H) + OVERSCAN);
    const topPad = start * PAIR_H;
    const botPad = Math.max(0, (total - end)) * PAIR_H;
    const view = _matchCurrentView;
    let html = '';
    if (topPad > 0) html += `<tr class="mvs-spacer-row" style="height:${topPad}px"><td colspan="7"></td></tr>`;
    for (let i = start; i < end; i++) {
      const r = list[i];
      const sc = r.sim;
      const cls = sc >= 85 ? 'm-score-hi' : sc >= 60 ? 'm-score-mid' : 'm-score-lo';
      const rowAttr = ` data-mrow="${i}" data-mview="${view}" style="cursor:pointer"`;
      let tag, btnHtml;
      tag = r.aInDB && r.bInDB && r.aKey !== r.bKey
        ? '<span class="m-tag m-tag-mrg" title="Объединить группы">🔀</span>'
        : r.aInDB || r.bInDB
          ? '<span class="m-tag m-tag-syn">синоним</span>'
          : '<span class="m-tag m-tag-new">новое</span>';
      btnHtml = `<button class="m-ibtn" data-openm="${i}" data-mview="${view}" title="Добавить в базу штрихкодов">+</button>`;
      html += `<tr class="mp-a"${rowAttr}>
        <td rowspan="2" style="text-align:center;color:#999;font-size:11px;vertical-align:middle;width:32px;">${i+1}</td>
        <td rowspan="2" class="${cls}" style="text-align:center;vertical-align:middle;width:46px;">${sc}%</td>
        <td><span class="src-lbl">${esc(r.file1)}</span></td>
        <td>${esc(r.name1)}</td>
        <td style="font-family:monospace;font-size:11px;">${esc(r.bc1)}</td>
        <td rowspan="2" style="text-align:center;vertical-align:middle;width:60px;">${tag}</td>
        <td rowspan="2" style="vertical-align:middle;text-align:center;width:38px;">${btnHtml}</td>
      </tr><tr class="mp-b"${rowAttr}>
        <td><span class="src-lbl">${esc(r.file2)}</span></td>
        <td>${esc(r.name2)}</td>
        <td style="font-family:monospace;font-size:11px;">${esc(r.bc2)}</td>
      </tr>`;
    }
    if (botPad > 0) html += `<tr class="mvs-spacer-row" style="height:${botPad}px"><td colspan="7"></td></tr>`;
    document.getElementById('matcherTbody').innerHTML = html;
  }

  // Attach scroll handler once
  if (!wrap._mvsScrollAttached) {
    wrap._mvsScrollAttached = true;
    wrap.addEventListener('scroll', function() {
      if (!wrap._mvsTicking) {
        wrap._mvsTicking = true;
        requestAnimationFrame(function() { _mvsRenderMatcherRows(); wrap._mvsTicking = false; });
      }
    }, { passive: true });
  }
  // Store renderer for re-use on scroll
  wrap._mvsRender = _mvsRenderMatcherRows;
  wrap.scrollTop = 0;
  _mvsRenderMatcherRows();
}

// Matcher table click delegation
document.getElementById('matcherTbody').addEventListener('click', function(e) {
  const openBtn = e.target.closest('[data-openm]');
  if (openBtn) {
    e.stopPropagation();
    openMatchModal(+openBtn.dataset.openm, openBtn.dataset.mview);
    return;
  }
  const row = e.target.closest('tr[data-mrow]');
  if (row && !e.target.closest('button')) {
    openMatchModal(+row.dataset.mrow, row.dataset.mview);
  }
});

document.getElementById('matcherSearchInp').addEventListener('input', renderMatcherTable);

document.querySelectorAll('.mstat[data-mv]').forEach(s =>
  s.addEventListener('click', () => setMatchView(s.dataset.mv)));

document.getElementById('matcherRunBtn').addEventListener('click', runMatcher);

document.getElementById('matcherHideKnownBtn').addEventListener('click', function() {
  _matchHideKnown = !_matchHideKnown;
  this.textContent = _matchHideKnown ? '👁 Показать все пары' : '🔗 Скрыть уже связанные';
  this.style.background = _matchHideKnown ? '#fff3cd' : '';
  this.style.borderColor = _matchHideKnown ? '#ffc107' : '';
  this.style.color = _matchHideKnown ? '#7d5a00' : '';
  updateMatcherStats();
  renderMatcherTable();
});


// ════════════════════════════════════════════════════════════════════════════
// MATCH MODAL — tab switching + brand tab
// ════════════════════════════════════════════════════════════════════════════
function mcSwitchTab(tab) {
  const isSyn = (tab === 'syn');
  document.getElementById('mcPaneSyn').style.display   = isSyn ? '' : 'none';
  document.getElementById('mcPaneBrand').style.display = isSyn ? 'none' : '';
  const mcOkBtn = document.getElementById('mcOkBtn');
  if (mcOkBtn) mcOkBtn.style.display = isSyn ? '' : 'none';
  const tSyn = document.getElementById('mcTabSyn');
  const tBrand = document.getElementById('mcTabBrand');
  if (tSyn)   { tSyn.style.borderBottomColor   = isSyn  ? '#217346' : 'transparent'; tSyn.style.color   = isSyn  ? '#217346' : '#888'; tSyn.style.fontWeight = isSyn ? '700' : '400'; }
  if (tBrand) { tBrand.style.borderBottomColor = !isSyn ? '#217346' : 'transparent'; tBrand.style.color = !isSyn ? '#217346' : '#888'; tBrand.style.fontWeight = !isSyn ? '700' : '400'; }
}

function mcbFillFromPair(name1, name2) {
  function extractBrandWords(name) {
    if (!name) return [];
    return name.toLowerCase()
      .replace(/[^a-zA-Zа-яёА-ЯЁ0-9\s]/g, ' ')
      .split(/\s+/)
      .filter(function(w) { return w.length >= 3 && !/^\d+$/.test(w); })
      .slice(0, 5);
  }
  let words1 = extractBrandWords(name1);
  let words2 = extractBrandWords(name2);
  let all = [];
  let seen = {};
  words1.concat(words2).forEach(function(w) { if (!seen[w]) { seen[w] = true; all.push(w); } });

  const elSugR = document.getElementById('mcbSuggestRow');
  const elSugT = document.getElementById('mcbSuggestTags');
  const elCanon = document.getElementById('mcbCanon');
  const elSyns = document.getElementById('mcbSyns');

  if (all.length && elSugT) {
    elSugR.style.display = '';
    elSugT.innerHTML = all.map(function(w) {
      let safe = w.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
      return '<button class="mqb-tag" data-w="' + safe + '" style="display:inline-block;margin:2px 3px;padding:2px 8px;background:#fff;border:1px solid #217346;border-radius:10px;font-size:11px;cursor:pointer;color:#155724;">' + safe + '</button>';
    }).join('');
    elSugT.querySelectorAll('.mqb-tag').forEach(function(btn) {
      btn.addEventListener('click', function() {
        if (elCanon) elCanon.value = this.dataset.w;
      });
    });
  } else if (elSugR) {
    elSugR.style.display = 'none';
  }
  if (elCanon && !elCanon.value && words1.length) elCanon.value = words1[0];
  if (elSyns  && !elSyns.value  && words2.length) elSyns.value  = words2.join(', ');
}

document.addEventListener('DOMContentLoaded', function() {
  const mcbSaveBtn = document.getElementById('mcbSaveBtn');
  if (mcbSaveBtn) {
    mcbSaveBtn.addEventListener('click', function() {
      const elCanon = document.getElementById('mcbCanon');
      const elSyns = document.getElementById('mcbSyns');
      const elAnti = document.getElementById('mcbAnti');
      const elStatus = document.getElementById('mcbStatus');
      let canon = brandNormKey(elCanon.value);
      if (!canon) { elStatus.textContent = '⚠ Введите канонический бренд'; return; }
      let syns = (elSyns.value || '').split(',').map(function(s) { return brandNormKey(s); }).filter(Boolean);
      let anti = (elAnti.value || '').split(',').map(function(s) { return brandNormKey(s); }).filter(Boolean);

      const check = brandCheckConflicts(canon, syns, anti, null);
      if (check.conflicts.length) {
        elStatus.innerHTML = `<span style="color:#c0392b;">⚠ ${check.conflicts[0]}</span>`;
        return;
      }

      if (_brandDB[canon]) {
        let ex = _brandDB[canon];
        const mergedSyns = Array.from(new Set((ex.synonyms || []).concat(syns)));
        const mergedAnti = Array.from(new Set((ex.antonyms || []).concat(anti)));
        // Проверяем пересечения в объединённых данных
        const mergedConflict = mergedSyns.filter(s => mergedAnti.includes(s));
        if (mergedConflict.length) {
          elStatus.innerHTML = `<span style="color:#c0392b;">⚠ Противоречие: «${mergedConflict[0]}» в синонимах и антонимах</span>`;
          return;
        }
        _brandDB[canon] = { synonyms: mergedSyns, antonyms: mergedAnti };
      } else {
        _brandDB[canon] = { synonyms: syns, antonyms: anti };
      }
      brandRender();
      brandMarkUnsaved();
      showToast('Бренд «' + canon + '» сохранён', 'ok');
      elStatus.textContent = '✓ Сохранён';
      setTimeout(function() { elStatus.textContent = ''; }, 2500);
      elCanon.value = ''; elSyns.value = ''; elAnti.value = '';
      const sugR = document.getElementById('mcbSuggestRow');
      if (sugR) sugR.style.display = 'none';
    });
  }
});

// ════════════════════════════════════════════════════════════════════════════
// MATCH MODAL
// ════════════════════════════════════════════════════════════════════════════
function openMatchModal(i, view) {
  // i is an index into _matchRenderedPairs (the currently displayed list)
  const r = _matchRenderedPairs[i];
  if (!r) return;
  // Refresh aInDB/bInDB/aKey/bKey from current jeDB state
  const bc2key = new Map();
  for (const [key, val] of Object.entries(jeDB)) {
    bc2key.set(key, key);
    if (Array.isArray(val)) val.slice(1).forEach(s => { s = String(s).trim(); if (s) bc2key.set(s, key); });
  }
  r.aInDB = bc2key.has(r.bc1); r.aKey = bc2key.get(r.bc1);
  r.bInDB = bc2key.has(r.bc2); r.bKey = bc2key.get(r.bc2);

  // Store a live reference to the pair object (not a copy)
  _matchPending = { pair: r, renderedIdx: i, view };

  const sc = r.sim;
  const col = sc >= 85 ? '#155724' : sc >= 60 ? '#7d5a00' : '#721c24';
  document.getElementById('mcScore').innerHTML = `<span style="color:${col}">${sc}%</span>`;
  document.getElementById('mc-src1').textContent  = r.file1;
  document.getElementById('mc-name1').textContent = r.name1;
  document.getElementById('mc-bc1').textContent   = r.bc1;
  document.getElementById('mc-src2').textContent  = r.file2;
  document.getElementById('mc-name2').textContent = r.name2;
  document.getElementById('mc-bc2').textContent   = r.bc2;

  const isReadOnly = r.aInDB && r.bInDB && r.aKey === r.bKey;
  let action = '';
  if (isReadOnly) action = `Уже в одной группе (главный ШК: «${r.aKey}»)`;
  else if (r.aInDB && r.bInDB && r.aKey !== r.bKey) action = `Объединить группы «${r.aKey}» и «${r.bKey}» → один синоним`;
  else if (!r.aInDB && r.bInDB) action = `Добавить «${r.bc1}» как синоним к группе «${r.bKey}»`;
  else if (r.aInDB && !r.bInDB) action = `Добавить «${r.bc2}» как синоним к группе «${r.aKey}»`;
  else action = `Создать новую группу: главный «${r.bc1}», синоним «${r.bc2}»`;

  document.getElementById('mcAction').textContent = action;
  document.getElementById('mcOkBtn').style.display = isReadOnly ? 'none' : '';
  document.getElementById('mcOkBtn').textContent = 'Подтвердить';
  // Reset to synonym tab, clear brand fields
  mcSwitchTab('syn');
  ['mcbCanon','mcbSyns','mcbAnti'].forEach(function(id) {
    const el = document.getElementById(id); if (el) el.value = '';
  });
  const mcbStatus = document.getElementById('mcbStatus'); if (mcbStatus) mcbStatus.textContent = '';
  const sugR = document.getElementById('mcbSuggestRow'); if (sugR) sugR.style.display = 'none';
  // Pre-fill brand suggestions
  mcbFillFromPair(r.name1 || '', r.name2 || '');
  document.getElementById('matchConfirmModal').style.display = 'flex';
}

function closeMatchModal() {
  document.getElementById('matchConfirmModal').style.display = 'none';
  _matchPending = null;
}

function applyMatchPair(r) {
  // Always work with fresh aInDB/bInDB state from current jeDB
  const bc2key = new Map();
  for (const [key, val] of Object.entries(jeDB)) {
    bc2key.set(String(key), String(key));
    if (Array.isArray(val)) val.slice(1).forEach(s => { s = String(s).trim(); if (s) bc2key.set(s, String(key)); });
  }
  const bc1 = String(r.bc1).trim(), bc2 = String(r.bc2).trim();
  const name1 = String(r.name1 || bc1), name2 = String(r.name2 || bc2);
  const aInDB = bc2key.has(bc1), bInDB = bc2key.has(bc2);
  const aKey = bc2key.get(bc1), bKey = bc2key.get(bc2);

  jeDBSaveHistory();

  if (aInDB && bInDB && aKey !== bKey) {
    // Merge two groups into aKey, keeping aKey as main
    const a1 = Array.isArray(jeDB[aKey]) ? jeDB[aKey] : [String(aKey)];
    const a2 = Array.isArray(jeDB[bKey]) ? jeDB[bKey] : [String(bKey)];
    const synSet = new Set();
    // All synonyms from both groups + bKey itself becomes a synonym
    a1.slice(1).forEach(s => { s = String(s).trim(); if (s && s !== aKey) synSet.add(s); });
    synSet.add(bKey);
    a2.slice(1).forEach(s => { s = String(s).trim(); if (s && s !== aKey) synSet.add(s); });
    jeDB[aKey] = [a1[0] || a2[0] || aKey, ...synSet];
    delete jeDB[bKey];
  } else if (!aInDB && bInDB) {
    // bc1 is new → add as synonym to bKey group
    const arr = jeDB[bKey];
    if (Array.isArray(arr)) {
      if (!arr.map(s=>String(s).trim()).slice(1).includes(bc1)) arr.push(bc1);
    } else {
      jeDB[bKey] = [name2, bc1];
    }
  } else if (aInDB && !bInDB) {
    // bc2 is new → add as synonym to aKey group
    const arr = jeDB[aKey];
    if (Array.isArray(arr)) {
      if (!arr.map(s=>String(s).trim()).slice(1).includes(bc2)) arr.push(bc2);
    } else {
      jeDB[aKey] = [name1, bc2];
    }
  } else {
    // Both new → create new group with bc1 as main key
    if (!jeDB[bc1]) {
      jeDB[bc1] = [name1, bc2];
    } else {
      // bc1 already a main key (shouldn't happen after re-check, but be safe)
      const arr = jeDB[bc1];
      if (!arr.map(s=>String(s).trim()).slice(1).includes(bc2)) arr.push(bc2);
    }
  }
  jeDBNotifyChange();
}

function confirmMatchAction() {
  if (!_matchPending) return;
  const { pair, view } = _matchPending;
  if (!pair) return;

  // Re-check current state before applying
  const bc2key = new Map();
  for (const [key, val] of Object.entries(jeDB)) {
    bc2key.set(key, key);
    if (Array.isArray(val)) val.slice(1).forEach(s => { s = String(s).trim(); if (s) bc2key.set(s, key); });
  }
  pair.aInDB = bc2key.has(pair.bc1); pair.aKey = bc2key.get(pair.bc1);
  pair.bInDB = bc2key.has(pair.bc2); pair.bKey = bc2key.get(pair.bc2);

  if (pair.aInDB && pair.bInDB && pair.aKey === pair.bKey) {
    // Already in same group, nothing to do
    closeMatchModal();
    showToast('Эти товары уже в одной группе', 'info');
    return;
  }

  applyMatchPair(pair);
  jeRenderEditor(true); // preserve scroll in synonym editor

  // Remove from _matchActivePairs by object reference
  const srcIdx = _matchActivePairs.indexOf(pair);
  if (srcIdx !== -1) _matchActivePairs.splice(srcIdx, 1);

  updateMatcherStats();
  clearTimeout(rebuildBarcodeAliasFromJeDB._t);
  rebuildBarcodeAliasFromJeDB._t = setTimeout(function() { if (typeof allFilesData !== 'undefined' && allFilesData.length > 0) { processData(); renderTable(true); updateUI(); } }, 120);
  closeMatchModal();
  renderMatcherTable();
  showToast('Добавлено в базу синонимов', 'ok');
}

function updateMatchPairTags() {
  const bc2key = new Map();
  for (const [key, val] of Object.entries(jeDB)) {
    bc2key.set(String(key), String(key));
    if (Array.isArray(val)) val.slice(1).forEach(s => { s = String(s).trim(); if (s) bc2key.set(s, String(key)); });
  }
  // Update active pairs: ones that are now "known" move to knownPairs
  const still = [], newKnown = [];
  for (const r of _matchActivePairs) {
    r.aInDB = bc2key.has(String(r.bc1)); r.aKey = bc2key.get(String(r.bc1));
    r.bInDB = bc2key.has(String(r.bc2)); r.bKey = bc2key.get(String(r.bc2));
    if (r.aInDB && r.bInDB && r.aKey === r.bKey) newKnown.push(r); else still.push(r);
  }
  _matchActivePairs.length = 0; _matchActivePairs.push(...still);
  // Add newly known to existing knownPairs (don't overwrite)
  _matchKnownPairs.push(...newKnown);
  // Also update tags on existing knownPairs
  for (const r of _matchKnownPairs) {
    r.aInDB = bc2key.has(String(r.bc1)); r.aKey = bc2key.get(String(r.bc1));
    r.bInDB = bc2key.has(String(r.bc2)); r.bKey = bc2key.get(String(r.bc2));
  }
  updateMatcherStats();
  if (document.querySelector('.nav-tab[data-pane="matcher"].active')) renderMatcherTable();
}

