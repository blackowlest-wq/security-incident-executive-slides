const fs=require('node:fs');
const path=require('node:path');
const PptxGenJS=require('pptxgenjs');
const D=__dirname;
const old=JSON.parse(fs.readFileSync(path.join(D,'previous_notes.json'),'utf8'));
const sources=JSON.parse(fs.readFileSync(path.join(D,'sources.json'),'utf8'));
const src=Object.fromEntries(sources.map(a=>[a.id,a]));
const p=new PptxGenJS();p.layout='LAYOUT_WIDE';
p.title='セキュリティインシデントと経営の備え（図解改訂版）';p.author='経営会議資料';p.lang='ja-JP';
p.theme={headFontFace:'Yu Gothic',bodyFontFace:'Yu Gothic',lang:'ja-JP'};
const C={ink:'13263D',gray:'526174',red:'AE2937',teal:'006B68',line:'D4DCE4',pale:'EDF3F6',white:'FFFFFF'};
const F='Yu Gothic';let n=0;const manifest=[];
function txt(s,v,x,y,w,h,size=24,o={}){s.addText(v,{x,y,w,h,margin:0,fontFace:F,fontSize:size,color:C.ink,valign:'mid',breakLine:false,...o});manifest[n-1].texts.push({text:v,x,y,w,h,size});}
function box(s,x,y,w,h,fill=C.pale){s.addShape(p.ShapeType.rect,{x,y,w,h,line:{color:fill,transparency:100},fill:{color:fill}});}
function line(s,x1,y1,x2,y2,color=C.line,arrow=false,width=1.5){s.addShape(p.ShapeType.line,{x:x1,y:y1,w:x2-x1,h:y2-y1,line:{color,width,...(arrow?{endArrowType:'triangle'}:{})}});}
function foot(s,ids=[],note=''){
 const gap=11.8/Math.max(ids.length,1);
 ids.forEach((id,i)=>txt(s,`${id} ${id==='S13'?'国家公安委員会ほか':src[id].org.split('（')[0]}`,.6+i*gap,7.02,gap-.12,.2,10,{color:C.gray,hyperlink:{url:src[id].url}}));
 if(!ids.length&&note)txt(s,note,.6,7.02,11.6,.2,10,{color:C.gray});
 txt(s,String(n).padStart(2,'0'),12.2,7.02,.5,.2,10,{color:C.gray,align:'right'});
}
function slide(title,ids=[],note=''){
 const s=p.addSlide();n++;manifest.push({slide:n,title,texts:[],sources:ids});s.background={color:C.white};
 txt(s,title,.6,.42,12.1,.75,31,{bold:true});line(s,.6,1.45,12.7,1.45,C.line,false,.7);foot(s,ids,note);return s;
}
function notes(s,key,oldnum,ids=[],detail=''){
 const previous=oldnum?old[String(oldnum)].replace(/^要点：[^\n]*\n?/,''):'';
 const references=ids.map(id=>`${id} ${src[id].org}\n${src[id].title}\n${src[id].date}\n${src[id].loc}\n${src[id].url}\n${src[id].claims}`).join('\n\n');
 s.addNotes(`要点：${key}\n\n${detail}\n\n${previous}\n\n出典・補足：\n${references}\n\n調査基準日：2026年10月22日（日本時間）`);
 manifest[n-1].oldSlide=oldnum||null;
 manifest[n-1].sources=[...new Set([...ids,...(previous.match(/S\d{2}/g)||[])])];
}
function stages(s,items,{x=.7,y=2.5,w=11.9,h=1.1,size=26,gap=.5}={}){
 const bw=(w-gap*(items.length-1))/items.length;
 items.forEach((it,i)=>{const xx=x+i*(bw+gap);box(s,xx,y,bw,h,it.fill||C.pale);txt(s,it.title,xx+.12,y+.12,bw-.24,h-.24,size,{bold:true,align:'center',color:it.color||C.ink});if(it.sub)txt(s,it.sub,xx,y+h+.35,bw,.9,21,{align:'center',color:C.gray});if(i<items.length-1)line(s,xx+bw+.07,y+h/2,xx+bw+gap-.07,y+h/2,C.gray,true);});
}
function chart(s,title,labels,values,{x,y,w,h,max,unit,colors=[C.ink],labelSize=16}={}){
 txt(s,title,x,y-.5,w,.4,23,{bold:true});
 s.addChart(p.ChartType.bar,[{name:title,labels,values}],{x,y,w,h,barDir:'col',barGrouping:'clustered',catAxisMultiLevelLabels:false,catAxisLabelFontFace:F,catAxisLabelFontSize:labelSize,catAxisLabelColor:C.gray,catAxisMajorTickMark:'none',catAxisLineColor:C.line,catAxisLineShow:true,valAxisMinVal:0,valAxisMaxVal:max,valAxisMajorUnit:max/4,valAxisLabelFontFace:F,valAxisLabelFontSize:12,valAxisLabelColor:C.gray,valAxisLineShow:false,valAxisMajorTickMark:'none',valGridLine:{color:C.line,width:.6},showLegend:false,showTitle:false,showValue:true,dataLabelPosition:'outEnd',dataLabelFontFace:F,dataLabelFontSize:17,dataLabelColor:C.ink,dataLabelFormatCode:'#,##0',chartColors:colors,showBorder:false,showShadow:false,chartArea:{fill:{color:C.white},border:{color:C.white,pt:0}},plotArea:{fill:{color:C.white},border:{color:C.white,pt:0}}});
 txt(s,unit,x,y+h+.12,w,.28,12,{color:C.gray});manifest[n-1].charts=(manifest[n-1].charts||[]).concat({labels,values});
}
// 1 — cover
{
 const s=p.addSlide();n++;manifest.push({slide:n,title:'セキュリティインシデントと経営の備え',texts:[],sources:[]});s.background={color:C.white};
 txt(s,'セキュリティインシデントと\n経営の備え',.75,1.55,11.9,1.9,44,{bold:true});
 txt(s,'国内事例から考える、事業継続と改善投資',.8,4.0,11.9,.65,24,{color:C.gray});
 txt(s,'2026年10月22日',.8,6.15,8,.4,17,{color:C.gray});foot(s);notes(s,'重要業務の継続と復旧を目的に、改善投資を検討する。',1);
}
// 2 — replaces original slide 3. Do not conflate report and recognized offence counts.
{
 const s=slide('不正アクセスなど、幅広いインシデントが続く',['S13','S14']);
 chart(s,'不正アクセスの認知',['2021','2022','2023','2024','2025'],[1516,2200,6312,5358,7190],{x:.65,y:2.28,w:5.9,h:3.7,max:8000,unit:'警察庁等／件・暦年'});
 chart(s,'幅広いインシデントの報告',['2025年\n4～6月','2026年\n4～6月'],[14558,14056],{x:7.02,y:2.28,w:5.65,h:3.7,max:16000,unit:'JPCERT/CC／件・四半期',colors:[C.teal],labelSize:16});
 txt(s,'フィッシング・Web改ざん・マルウェアなど',7.05,6.35,5.55,.32,15,{color:C.teal,bold:true});
 txt(s,'※異なる集計のため合算不可。JPCERT/CCの報告には国内外・同一事象の複数報告を含む。',.67,6.7,12,.24,11.5,{color:C.gray});
 notes(s,'不正アクセスが近年増加する一方、幅広いインシデント報告も継続している。',null,['S13','S14'],'左は不正アクセス禁止法の犯罪構成要件に該当する行為の認知件数。2024年に減少したことも省略しない。被害企業数ではない。右は国内外からWebフォーム・メールで寄せられた報告の総数で、被害総数・攻撃回数・重複除去済みのインシデント件数とは異なる。2026年4～6月は14,056件で前年同期14,558件より減少している。すべての種類が単調増加しているとは主張しない。異なる母集団と期間のため、左右の値を合算・直接比較しない。2026年4～6月の重複をまとめたインシデント件数は9,305件だが、報告件数のグラフへ混在させない。');
}
// 3 — KADOKAWA: one business outcome; secondary figures in speaker notes.
{
 const s=slide('KADOKAWA：サービス停止が収益に波及',['S02','S03']);
 stages(s,[{title:'動画サービス\n停止'},{title:'出版にも\n影響'},{title:'営業利益\n減少',fill:'F8ECEE',color:C.red}],{y:2.0,h:1.2});
 txt(s,'▲47億円',.7,4.0,11.9,1.12,66,{bold:true,color:C.red,align:'center'});
 txt(s,'2025年3月期の営業利益への影響（会社試算）',.7,5.4,11.9,.6,23,{color:C.gray,align:'center'});
 notes(s,'サービス停止は、直接の復旧費用を超えて収益に波及する。',5,['S02','S03','S04'],'本編では営業利益への影響47億円に焦点を絞る。売上高影響83億円、特別損失24億円、個人情報254,241人、日付差58日と原因推測の詳細は以下の補足として保持する。これらを単純加算しない。');
}
// 4 — ASKUL: a single exception opens an access path.
{
 const s=slide('アスクル：認証の例外が侵入口になった',['S05']);
 stages(s,[{title:'委託先の\n管理者ID',sub:'多要素認証が未適用',fill:'F8ECEE',color:C.red},{title:'社内へ\n不正アクセス'},{title:'物流システム\n被害'}],{y:2.45,h:1.35});
 txt(s,'委託先を含め、認証の例外を残さない',.7,5.5,11.9,.65,30,{bold:true,color:C.teal,align:'center'});
 notes(s,'委託先を含めた認証の例外管理が必要。',6,['S05'],'図は公表されたアカウント不正利用と物流システム被害の関係を簡略化したもの。認証情報の漏洩手段は不明であり、原因を完全に究明したとは断定しない。監視未整備、バックアップ暗号化と会社再発防止策の詳細はノートに保持する。下段は作成者の提案。');
}
// 5 — ASKUL: duration only, no cost figures on the same page.
{
 const s=slide('段階的に再開しても、復旧は長期化する',['S05','S06']);
 txt(s,'108日',.75,1.85,11.8,1.1,64,{bold:true,color:C.red,align:'center'});
 txt(s,'全物流拠点の出荷再開までの日付差',.75,3.0,11.8,.5,23,{color:C.gray,align:'center'});
 stages(s,[{title:'10/19\n受注・出荷停止',sub:'2025年'},{title:'10/29\n出荷トライアル'},{title:'2/4\n全拠点で出荷再開',sub:'2026年'}],{y:4.05,h:1.15,size:23});
 txt(s,'アスクルの公表日から計算。108日間の全面停止ではない。間隔は日数に比例しない。',.72,6.45,11.9,.35,13,{color:C.gray});
 notes(s,'部分再開しても、事業を戻すまでに長い時間がかかる。',7,['S05','S06','S07'],'表示の108日は2025-10-19から2026-02-04の日付差で、初日を含めない。旧ページの費用51.08億円・12.69億円はノートへ移し、復旧期間だけを本編のメッセージとした。');
}
// 6 — ASAHI: one financial effect, conceptual components, no invented proportions.
{
 const s=slide('被害は、事業利益の減少と追加費用に及ぶ',['S09']);
 txt(s,'アサヒグループ',.7,1.85,11.9,.45,24,{align:'center',bold:true});
 txt(s,'約400億円',.7,2.55,11.9,1.2,66,{color:C.red,bold:true,align:'center'});
 txt(s,'2025年の業績影響（会社算定）',.7,3.93,11.9,.45,23,{color:C.gray,align:'center'});
 line(s,6.65,4.65,6.65,5.05,C.gray);line(s,3.2,5.05,10.12,5.05,C.gray);
 line(s,3.2,5.05,3.2,5.38,C.gray);line(s,10.12,5.05,10.12,5.38,C.gray);
 txt(s,'事業利益の減少',.7,5.55,5.0,.55,28,{bold:true,align:'center'});
 txt(s,'「その他」の費用',7.55,5.55,5.0,.55,28,{bold:true,align:'center'});
 txt(s,'被害総額・特別損失の額ではない。図の幅は金額の割合を表さない。',.72,6.5,11.9,.3,13,{color:C.gray});
 notes(s,'直接費用に加え、事業の収益にも影響する。',8,['S08','S09'],'会社が説明する「事業利益の減少額とその他費用」の合計約400億円を示す。左右の幅で内訳比率を表していない。受注再開・物流正常化の経過と会社再発防止策はノートに保持する。');
}
// 7 — attack stages matched to controls.
{
 const s=slide('事業停止を抑えるには、侵入後の対策も必要',['S01','S05']);
 stages(s,[{title:'侵入'},{title:'被害拡大'},{title:'暗号化・窃取'},{title:'業務停止',fill:'F8ECEE',color:C.red}],{y:2.3,h:1.1,size:26,gap:.4});
 const labels=[['認証・更新',1.85],['監視・権限分離',5.65],['バックアップ・復旧',9.8]];
 labels.forEach(([v,x])=>{line(s,x,3.85,x,4.55,C.teal,true,2);txt(s,v,x-1.65,4.92,3.3,.85,24,{bold:true,color:C.teal,align:'center'});});
 txt(s,'事例から整理した一般的な流れと対策。自社の被害診断ではない。',.72,6.5,11.9,.3,13,{color:C.gray});
 notes(s,'予防だけでなく、侵入後の拡大抑制と復旧を整備する。',9,['S01','S05','S08'],'各段階を一般化した概念図。すべての企業で全段階・同一の順序を確認したと主張するものではない。バックアップで窃取や漏えいまで防げるという意味ではない。');
}
// 8 — PHP: one action, details editable.
{
 const s=slide('PHPの重大脆弱性には、基盤更新が必要',['S10','S11','S12']);
 txt(s,'2026年10月21日公表',.7,1.83,11.9,.4,20,{color:C.gray});
 box(s,.75,2.55,4.05,1.4,'F8ECEE');txt(s,'CVE-2026-109618\nCVE-2026-109619',.9,2.7,3.75,1.1,23,{bold:true,color:C.red,align:'center'});
 line(s,5.05,3.25,6.15,3.25,C.gray,true,2);
 txt(s,'認証前のコード実行\nのおそれ',6.5,2.7,5.8,1.15,30,{bold:true});
 txt(s,'各系列の修正リリース',.7,4.65,11.9,.45,21,{color:C.gray});
 txt(s,'8.2.35    8.3.36    8.4.27    8.5.12',.7,5.3,11.9,.75,32,{bold:true,color:C.teal});
 txt(s,'成立条件は構成・アプリケーションに依存。ベンダーの修正バックポートも確認する。',.72,6.5,11.9,.3,13,{color:C.gray});
 notes(s,'該当するPHP基盤を特定し、速やかに更新する。',10,['S10','S11','S12'],'本編では緊急更新という行動に絞り、OOM破損・解放後メモリ参照・方式ごとの実証範囲の詳細はノートへ移す。国内3事例と本CVEとの関連は未確認であり、図もその関係を示さない。');
}
// 9 — completion evidence.
{
 const s=slide('対処完了は、稼働環境で確認する',['S10']);
 stages(s,[{title:'対象を特定'},{title:'修正を適用'},{title:'再起動'},{title:'業務を検証',fill:'E7F2F0',color:C.teal}],{y:2.65,h:1.3,size:25,gap:.4});
 line(s,1.4,4.6,11.95,4.6,C.teal);line(s,1.4,4.6,1.4,4.25,C.teal);line(s,11.95,4.6,11.95,4.25,C.teal);
 txt(s,'稼働版・再起動・業務テストの証拠を残す',.7,5.25,11.9,.85,29,{color:C.teal,bold:true,align:'center'});
 notes(s,'更新したという申告だけでなく、稼働環境で完了を確認する。',11,['S10','S11','S12'],'ベンダー修正やディストリビューションのバックポートも確認する。WAF等の限定的緩和を修正と同等に扱わない。更新が困難なら公開制限を責任者が事業影響と合わせて判断する。');
}
// 10 — staged investment. Tiered scope is a proposal, not measured effectiveness.
{
 const s=slide('投資は、重要業務から段階的に広げる',[],'投資の優先順位は提案。自社の現状確認後に確定');
 const x=[.75,4.95,9.15],ys=[4.2,3.2,2.2];
 const heads=['緊急是正','検知・復旧','全社へ拡張'];
 const subs=['公開資産・認証','重要業務を対象','委託先も含む'];
 for(let i=0;i<3;i++){box(s,x[i],ys[i],3.4,1.15,i===1?'E7F2F0':C.pale);txt(s,heads[i],x[i]+.15,ys[i]+.15,3.1,.75,29,{bold:true,color:i===1?C.teal:C.ink,align:'center'});txt(s,subs[i],x[i],ys[i]+1.42,3.4,.48,21,{align:'center',color:C.gray});if(i<2){line(s,x[i]+3.62,ys[i]+.58,x[i+1]-.17,ys[i]+.58,C.gray,true);}}
 notes(s,'重要業務の検知・復旧を推奨範囲とし、検証して全社へ広げる。',12,['S01','S05','S08'],'図の段差・面積は予算額や効果率を示さない。緊急是正だけでは監視・復旧のリスクが残る。最低限、推奨、拡張の費用項目・残存リスクはノートに保持する。');
}
// 11 — budget decision: loss side versus implementation cost side.
{
 const s=slide('予算は、抑える損失と実装費を照合して決める',[],'予算額・発生確率・ROIは未算定');
 txt(s,'止まったときの損失',.7,2.0,5.25,.6,28,{bold:true,color:C.red,align:'center'});
 txt(s,'失う限界利益\n＋追加の対応費\n－重複・回収の調整',.8,3.05,5.05,1.95,29,{align:'center'});
 txt(s,'対策を続ける費用',7.38,2.0,5.25,.6,28,{bold:true,color:C.teal,align:'center'});
 txt(s,'初期費用\n＋運用費・社内工数',7.48,3.3,5.05,1.45,29,{align:'center'});
 line(s,6.67,2.05,6.67,5.0,C.line);
 line(s,3.3,5.35,3.3,5.65,C.gray);line(s,10.0,5.35,10.0,5.65,C.gray);line(s,3.3,5.65,10.0,5.65,C.gray);
 txt(s,'自社の業務・見積で予算審議',.75,6.0,11.8,.6,28,{bold:true,align:'center'});
 notes(s,'自社の停止シナリオの損失と、実装・継続運用の費用を照合する。',13,[],'図は判断に必要な項目を整理した概念図で、公的な統一算式ではない。損失見込みの全額が投資によって防げるとはみなさない。具体的な対策の削減効果、停止損失、予算は自社データで評価する。');
}
// 12 — 90 day outcome.
{
 const s=slide('90日で、重要業務を戻せるか実測する',[],'日数・実行範囲は提案。PHPの緊急対応は直ちに実施');
 line(s,1.1,3.1,12.1,3.1,C.gray,false,2);
 const x=[1.0,5.25,9.5];
 [['30日','把握・見積'],['60日','是正・準備'],['90日','復旧演習']].forEach((a,i)=>{box(s,x[i],2.25,2.8,1.1,i===2?'E7F2F0':C.pale);txt(s,a[0],x[i],2.45,2.8,.65,36,{bold:true,align:'center',color:i===2?C.teal:C.ink});txt(s,a[1],x[i]-.2,3.95,3.2,.7,27,{align:'center',bold:true});});
 txt(s,'事業責任者とIT責任者が、復旧の結果を共同で確認',.75,5.6,11.8,.7,27,{align:'center',color:C.teal,bold:true});
 notes(s,'90日以内に重要業務の復旧・代替運用を実測する。',14,[],'30・60・90日の提案ロードマップ。緊急更新を30日後まで待つ計画ではない。責任者と成果物の細目、予算決裁条件はノートに保持する。');
}
// 13 — evidence-oriented oversight.
{
 const s=slide('運用と復旧の証拠で、成果を確認する',[],'成果指標は提案。自社の現状値・達成状況は未測定');
 stages(s,[{title:'対象を網羅',sub:'対象台帳'},{title:'即時に対応',sub:'訓練記録'},{title:'業務を復旧',sub:'復旧の実測'}],{y:2.3,h:1.25});
 line(s,2.6,4.85,10.75,4.85,C.teal);line(s,6.67,4.85,6.67,5.35,C.teal,true,2);
 txt(s,'経営へ報告',.75,5.75,11.8,.65,31,{bold:true,color:C.teal,align:'center'});
 notes(s,'導入数だけでなく、適用範囲・対応・復旧の証拠で投資成果を見る。',15,[],'対象台帳を分母に固定し未確認資産も併記する。MFA適用100％・PHP未対処0件は提案であり自社実績ではない。復旧の目標は事業部門が決め、演習実測と比較する。詳細な指標定義はノートに保持する。');
}
// 14 — one ask with its decision cadence.
{
 const s=slide('重要業務の継続と復旧を目的に、投資着手',[],'決議案。予算金額は自社評価・見積後に確定');
 txt(s,'本日の承認',.7,1.95,11.9,.7,37,{bold:true,color:C.teal,align:'center'});
 stages(s,[{title:'着手・責任者\nを決定',fill:'E7F2F0',color:C.teal},{title:'30日以内\n投資案を審議'},{title:'90日以内\n復旧結果を確認'}],{y:3.25,h:1.45,size:26});
 notes(s,'本日、責任者と期限を決めて改善投資に着手する。',16,[],'予算の具体額は自社の見積・決裁に基づき決定する。30日以内に費用・成果・残存リスクをそろえて審議し、90日以内の復旧実測で追加投資の要否を判断する。緊急保守の権限、経営スポンサー・CISO相当・事業責任者の役割はノートに保持する。');
}
// 15-18 — references only. Dense definitions move to handout and notes.
const groups=[sources.slice(0,4),sources.slice(4,8),sources.slice(8,11),sources.slice(11)];
groups.forEach((group,idx)=>{
 const s=slide(`出典 ${idx+1}／4`,[],'完全URLは該当ページのノートにも記載');
 const step=group.length===4?1.23:1.56;
 group.forEach((a,i)=>{const y=1.85+i*step;txt(s,`${a.id}  ${a.org}`, .68,y,12,.25,13,{bold:true});txt(s,a.title,.68,y+.33,12,.33,12.5,{bold:true});txt(s,`${a.date} ／ ${a.loc}`,.68,y+.76,12,.23,10,{color:C.gray});txt(s,a.url,.68,y+1.02,12,.19,9.5,{color:C.teal,hyperlink:{url:a.url}});});
 notes(s,'重要事実を一次資料へさかのぼって確認する。',null,group.map(a=>a.id));
});
fs.writeFileSync(path.join(D,'sources_all.json'),JSON.stringify(sources,null,2),'utf8');
fs.writeFileSync(path.join(D,'manifest.json'),JSON.stringify(manifest,null,2),'utf8');
p.writeFile({fileName:path.join(D,'candidate.pptx')});

