const pptxgen = require('pptxgenjs');
const R = require('../nps/datos2.json');

const P = { azul:'2B0B6B', azul2:'4A1D96', azul3:'7C3AED', lila:'A78BFA',
  lilaLt:'EDE9FE', bg:'F8F5FF', w:'FFFFFF', ln:'DDD6FE', txt:'6B7280',
  dk:'1E1235', ok:'16A34A', wn:'D97706', bad:'DC2626' };
const F = { h:'Bookman Old Style', b:'Calibri' };
const pres = new pptxgen();
pres.layout='LAYOUT_WIDE';
pres.title='Indicador NPS — Junio a agosto 2026';
const W=13.3, H=7.5, M=0.7;
const pc  = n => n==null ? '—' : Number(n).toFixed(1).replace('.',',')+'%';
const mil = n => Number(n).toLocaleString('es-CO');
const MES = R.meses;

function fondo(s,osc){ s.background={ color: osc?P.azul:P.bg }; }
function cab(s,t,sub,conChip){
  s.addText(t,{x:M,y:0.44,w:W-2*M-(conChip?1.5:0),h:0.56,isTextBox:true,margin:0,
    fontFace:F.h,fontSize:26,bold:true,color:P.azul});
  if(sub) s.addText(sub,{x:M,y:1.02,w:W-2*M-(conChip?1.5:0),h:0.32,isTextBox:true,margin:0,
    fontFace:F.b,fontSize:12,color:P.txt});
}
function fuente(s,t){
  s.addText(t,{x:M,y:H-0.5,w:W-2*M,h:0.28,isTextBox:true,margin:0,
    fontFace:F.b,fontSize:9,color:P.txt});
}
// Etiqueta de skill, el separador visual de todo el informe
function chip(s,x,y,skill){
  const c = skill==='Chat' ? P.bad : P.azul2;
  s.addShape(pres.ShapeType.roundRect,{x,y,w:1.15,h:0.3,rectRadius:0.15,
    fill:{color:c},line:{width:0}});
  s.addText(skill.toUpperCase(),{x,y:y+0.02,w:1.15,h:0.26,isTextBox:true,margin:0,
    align:'center',fontFace:F.b,fontSize:9.5,bold:true,color:P.w,charSpacing:1.2});
}
function tabla(s,x,y,anchos,cabs,filas,fmt){
  let cx=x;
  cabs.forEach((c,i)=>{
    s.addShape(pres.ShapeType.rect,{x:cx,y,w:anchos[i],h:0.36,
      fill:{color:P.azul2},line:{width:0}});
    s.addText(c,{x:cx+0.07,y:y+0.03,w:anchos[i]-0.14,h:0.3,isTextBox:true,margin:0,
      align:i===0?'left':'center',fontFace:F.b,fontSize:8.5,bold:true,color:P.w,charSpacing:0.6});
    cx+=anchos[i];
  });
  let cy=y+0.36;
  filas.forEach((f,r)=>{
    s.addShape(pres.ShapeType.rect,{x,y:cy,w:anchos.reduce((a,b)=>a+b,0),h:0.34,
      fill:{color: r%2?P.w:'FCFAFF'},line:{color:P.ln,width:0.5}});
    let bx=x;
    f.forEach((v,i)=>{
      const col = fmt ? fmt(r,i,v) : P.dk;
      s.addText(String(v),{x:bx+0.07,y:cy+0.03,w:anchos[i]-0.14,h:0.28,isTextBox:true,
        margin:0,align:i===0?'left':'center',fontFace:i===0?F.b:F.h,
        fontSize:i===0?9.5:10,bold:i!==0,color:col});
      bx+=anchos[i];
    });
    cy+=0.34;
  });
  return cy;
}

/* ── 1. PORTADA ───────────────────────────────────────────── */
{
  const s=pres.addSlide(); fondo(s,true);
  s.addShape(pres.ShapeType.ellipse,{x:10.1,y:-1.8,w:5.6,h:5.6,
    fill:{color:P.azul2,transparency:60},line:{width:0}});
  s.addText('INFORME DE INDICADOR',{x:M,y:2.3,w:9,h:0.3,isTextBox:true,margin:0,
    fontFace:F.b,fontSize:12,bold:true,color:P.lila,charSpacing:3});
  s.addText('Net Promoter Score',{x:M,y:2.64,w:9.4,h:0.85,isTextBox:true,margin:0,
    fontFace:F.h,fontSize:42,bold:true,color:P.w});
  s.addText('Junio, julio y agosto de 2026',{x:M,y:3.54,w:9.4,h:0.45,isTextBox:true,margin:0,
    fontFace:F.h,fontSize:20,color:P.lila});
  s.addText('Comparativa mensual, resultado por tipificación y agentes con menor calificación. Inbound y Chat se presentan por separado.',
    {x:M,y:4.16,w:8.3,h:0.62,isTextBox:true,margin:0,fontFace:F.b,fontSize:13,
     color:'C4B5FD',lineSpacingMultiple:1.25});
  s.addText(`${mil(R.global.enc)} encuestas  ·  Inbound ${mil(R.resumen.Inbound.enc)}  ·  Chat ${mil(R.resumen.Chat.enc)}`,
    {x:M,y:5.2,w:9,h:0.3,isTextBox:true,margin:0,fontFace:F.b,fontSize:12,
     bold:true,color:P.lila,charSpacing:0.8});
}

/* ── 2. RESULTADO POR SKILL ───────────────────────────────── */
{
  const s=pres.addSlide(); fondo(s);
  cab(s,'Resultado del período por skill','NPS = porcentaje de promotores menos porcentaje de detractores');
  const gap=0.4, w=(W-2*M-gap)/2;
  ['Inbound','Chat'].forEach((sk,i)=>{
    const r=R.resumen[sk], x=M+i*(w+gap), esChat=sk==='Chat';
    s.addShape(pres.ShapeType.roundRect,{x,y:1.5,w,h:2.05,rectRadius:0.09,
      fill:{color:P.w},line:{color:esChat?'FECACA':P.ln,width:esChat?1.5:1}});
    chip(s,x+0.3,1.72,sk);
    s.addText(pc(r.nps),{x:x+0.3,y:2.06,w:w-0.6,h:0.82,isTextBox:true,margin:0,
      fontFace:F.h,fontSize:44,bold:true,color:esChat?P.bad:P.azul2});
    s.addText(`${mil(r.enc)} encuestas   ·   ${mil(r.det)} detractores   ·   ${pc(r.pdet)} de insatisfacción`,
      {x:x+0.3,y:2.9,w:w-0.6,h:0.34,isTextBox:true,margin:0,
       fontFace:F.b,fontSize:10.5,color:P.dk});
    s.addText(`Promotores ${mil(r.prom)}  ·  Neutros ${mil(r.neu)}`,
      {x:x+0.3,y:3.22,w:w-0.6,h:0.26,isTextBox:true,margin:0,
       fontFace:F.b,fontSize:10,color:P.txt});
  });
  // Tabla mensual, un bloque por skill
  let y=3.78;
  ['Inbound','Chat'].forEach((sk,i)=>{
    const r=R.resumen[sk], x=M+i*(w+gap);
    tabla(s,x,y,[w-3.3,1.15,1.05,1.1],
      ['MES','ENCUESTAS','DETRAC.','NPS'],
      r.mes.map(m=>[m.m,mil(m.enc),mil(m.det),pc(m.nps)]),
      (row,col)=> col===3 ? (r.mes[row].nps < r.nps-2 ? P.bad : (r.mes[row].nps > r.nps+2 ? P.ok : P.dk)) : P.dk);
  });
  s.addShape(pres.ShapeType.roundRect,{x:M,y:5.44,w:W-2*M,h:0.8,rectRadius:0.09,
    fill:{color:P.lilaLt},line:{color:P.ln,width:1}});
  s.addText([
    {text:'Los dos skills no son comparables entre sí. ',options:{bold:true,color:P.azul}},
    {text:`Inbound cierra en ${pc(R.resumen.Inbound.nps)} y Chat en ${pc(R.resumen.Chat.nps)}: 33 puntos de diferencia sostenidos en los tres meses. Todo el informe los presenta por separado.`,
     options:{color:P.dk}}
  ],{x:M+0.3,y:5.6,w:W-2*M-0.6,h:0.5,isTextBox:true,margin:0,fontFace:F.b,fontSize:11.5});
  fuente(s,'Fuente: hoja BBDD Encuestas de los consolidados mensuales de junio, julio y agosto de 2026.');
}

/* ── 3. COMPARATIVA MENSUAL ───────────────────────────────── */
{
  const s=pres.addSlide(); fondo(s);
  cab(s,'Comparativa mensual','Evolución del NPS de cada skill en el período');
  s.addChart(pres.ChartType.line,[
    {name:'Inbound',labels:MES,values:R.resumen.Inbound.mes.map(m=>m.nps)},
    {name:'Chat',   labels:MES,values:R.resumen.Chat.mes.map(m=>m.nps)}
  ],{ x:M,y:1.5,w:7.6,h:4.3, chartColors:[P.azul2,P.bad], lineDataSymbolSize:9,
      lineSize:3, showLegend:true, legendPos:'b', legendFontFace:F.b, legendFontSize:11,
      legendColor:P.dk, showValue:true, dataLabelFontFace:F.b, dataLabelFontSize:11,
      dataLabelColor:P.dk, dataLabelFormatCode:'0.0"%"', dataLabelPosition:'t',
      valAxisMinVal:0, valAxisMaxVal:85, showTitle:false,
      catAxisLabelColor:P.txt, catAxisLabelFontFace:F.b, catAxisLabelFontSize:12,
      valAxisLabelColor:P.txt, valAxisLabelFontFace:F.b, valAxisLabelFontSize:10,
      valGridLine:{color:P.ln,size:1}, catGridLine:{style:'none'} });
  const bx=8.6, bw=W-M-bx;
  s.addText('VARIACIÓN JULIO A AGOSTO',{x:bx,y:1.56,w:bw,h:0.26,isTextBox:true,margin:0,
    fontFace:F.b,fontSize:9.5,bold:true,color:P.azul3,charSpacing:1.2});
  let yy=1.92;
  ['Inbound','Chat'].forEach(sk=>{
    const r=R.resumen[sk], dif=+(r.mes[2].nps-r.mes[1].nps).toFixed(1);
    s.addShape(pres.ShapeType.roundRect,{x:bx,y:yy,w:bw,h:1.5,rectRadius:0.09,
      fill:{color:P.w},line:{color:P.ln,width:1}});
    chip(s,bx+0.25,yy+0.2,sk);
    s.addText((dif>0?'+':'')+String(dif).replace('.',',')+' pts',
      {x:bx+0.25,y:yy+0.56,w:bw-0.5,h:0.48,isTextBox:true,margin:0,
       fontFace:F.h,fontSize:26,bold:true,color:dif<-2?P.bad:(dif>0?P.ok:P.wn)});
    s.addText(`${pc(r.mes[1].nps)} → ${pc(r.mes[2].nps)}`,
      {x:bx+0.25,y:yy+1.06,w:bw-0.5,h:0.28,isTextBox:true,margin:0,
       fontFace:F.b,fontSize:11,color:P.txt});
    yy+=1.66;
  });
  s.addShape(pres.ShapeType.roundRect,{x:bx,y:yy,w:bw,h:0.98,rectRadius:0.09,
    fill:{color:'FEF2F2'},line:{color:'FECACA',width:1.5}});
  s.addText('La caída de agosto es de Chat. Inbound se mantiene dentro de su rango habitual.',
    {x:bx+0.25,y:yy+0.16,w:bw-0.5,h:0.7,isTextBox:true,margin:0,
     fontFace:F.b,fontSize:10.5,color:'991B1B',lineSpacingMultiple:1.2});
  fuente(s,'Agosto corresponde al corte hasta el día 20; los demás meses son completos.');
}

/* ── 4-5. TIPIFICACIÓN POR SKILL ──────────────────────────── */
['Inbound','Chat'].forEach(sk=>{
  const T=R.tipif[sk], base=R.resumen[sk].nps;
  const s=pres.addSlide(); fondo(s);
  cab(s,`Tipificación que más afecta — ${sk}`,
    `Procesos ordenados por cantidad de detractores aportados. NPS del skill: ${pc(base)}`, true);
  chip(s,W-M-1.15,0.5,sk);

  const fin = tabla(s,M,1.5,[3.6,1.1,1.05,1.15,0.85,0.85,0.95,0.9,1.0],
    ['PROCESO','ENCUESTAS','DETRAC.','% DEL TOTAL','JUNIO','JULIO','AGOSTO','TASA','NPS'],
    T.proc_vol.slice(0,7).map(x=>[x.t.length>34?x.t.slice(0,33)+'…':x.t,
      mil(x.enc),mil(x.det),pc(x.pct),mil(x.mes[0]),mil(x.mes[1]),mil(x.mes[2]),
      pc(x.pdet),pc(x.nps)]),
    (r,c)=>{ const x=T.proc_vol[r];
      if(c===8) return x.nps<base-8?P.bad:(x.nps>base+5?P.ok:P.dk);
      if(c===7) return x.pdet>=(sk==='Chat'?28:13)?P.bad:P.dk;
      return P.dk; });

  s.addText('Columnas Junio, Julio y Agosto: detractores aportados en cada mes.',
    {x:M,y:fin+0.1,w:6,h:0.26,isTextBox:true,margin:0,fontFace:F.b,fontSize:9.5,
     italic:true,color:P.txt});

  // Subprocesos con peor NPS
  const y2=fin+0.5;
  s.addText('SUBPROCESOS CON MENOR NPS',{x:M,y:y2,w:6,h:0.26,isTextBox:true,margin:0,
    fontFace:F.b,fontSize:10,bold:true,color:P.azul3,charSpacing:1.2});
  const sub=T.sub_nps.slice(0,4);
  const cw=(W-2*M-3*0.28)/4;
  sub.forEach((x,i)=>{
    const px=M+i*(cw+0.28);
    s.addShape(pres.ShapeType.roundRect,{x:px,y:y2+0.32,w:cw,h:1.16,rectRadius:0.09,
      fill:{color:P.w},line:{color:P.ln,width:1}});
    s.addText(pc(x.nps),{x:px+0.22,y:y2+0.44,w:cw-0.44,h:0.42,isTextBox:true,margin:0,
      fontFace:F.h,fontSize:21,bold:true,color:P.bad});
    s.addText(x.t.length>44?x.t.slice(0,43)+'…':x.t,
      {x:px+0.22,y:y2+0.86,w:cw-0.44,h:0.36,isTextBox:true,margin:0,
       fontFace:F.b,fontSize:9.5,bold:true,color:P.dk});
    s.addText(mil(x.enc)+' encuestas',{x:px+0.22,y:y2+1.2,w:cw-0.44,h:0.22,
      isTextBox:true,margin:0,fontFace:F.b,fontSize:9,color:P.txt});
  });
  fuente(s, sk==='Chat'
    ? 'Mínimo 40 encuestas por proceso y por subproceso, dado el menor volumen del skill.'
    : 'Mínimo 120 encuestas por proceso y 100 por subproceso para el cálculo de NPS.');
});

/* ── 6. MACROPROCESO ──────────────────────────────────────── */
{
  const s=pres.addSlide(); fondo(s);
  cab(s,'Detractores por macroproceso',
    'Los dos skills sobre la misma clasificación. El macroproceso «Banco Agrario de Colombia» agrupa la operación general.');
  const gap=0.4, w=(W-2*M-gap)/2;
  ['Inbound','Chat'].forEach((sk,i)=>{
    const x=M+i*(w+gap), datos=R.tipif[sk].macro.slice(0,6);
    chip(s,x,1.5,sk);
    tabla(s,x,1.94,[w-2.55,0.95,0.85,0.75],
      ['MACROPROCESO','ENCUEST.','DETRAC.','NPS'],
      datos.map(d=>[d.t.length>28?d.t.slice(0,27)+'…':d.t,mil(d.enc),mil(d.det),pc(d.nps)]),
      (r,c)=> c===3 ? (datos[r].nps<R.resumen[sk].nps-8?P.bad:P.dk) : P.dk);
  });
  s.addShape(pres.ShapeType.roundRect,{x:M,y:4.6,w:W-2*M,h:1.4,rectRadius:0.09,
    fill:{color:P.w},line:{color:P.ln,width:1}});
  s.addText('Observación sobre esta clasificación',{x:M+0.3,y:4.76,w:W-2*M-0.6,h:0.26,
    isTextBox:true,margin:0,fontFace:F.b,fontSize:10.5,bold:true,color:P.azul3});
  s.addText('El 92% de las encuestas del período se clasifica como «Banco Agrario de Colombia», por lo que el macroproceso no permite distinguir causas. El análisis útil está en el proceso y el subproceso, que sí discriminan. Se recomienda revisar el criterio de tipificación en este nivel.',
    {x:M+0.3,y:5.04,w:W-2*M-0.6,h:0.8,isTextBox:true,margin:0,
     fontFace:F.b,fontSize:11,color:P.dk,lineSpacingMultiple:1.2});
  fuente(s,'Macroproceso, proceso y subproceso corresponden a la tipificación registrada en la gestión.');
}

/* ── 7-8. AGENTES POR SKILL ───────────────────────────────── */
['Inbound','Chat'].forEach(sk=>{
  const A=R.agentes[sk];
  const s=pres.addSlide(); fondo(s);
  cab(s,`Agentes con menor NPS — ${sk}`,
    `${A.n} agentes con 50 o más encuestas en el período. NPS del skill: ${pc(A.base)}`, true);
  chip(s,W-M-1.15,0.5,sk);
  const fin = tabla(s,M,1.5,[4.9,1.35,1.35,1.35,1.25,1.7],
    ['AGENTE','ENCUESTAS','DETRACTORES','% DETRAC.','NPS','DIFERENCIA'],
    A.top.map(x=>[x.a.length>42?x.a.slice(0,41)+'…':x.a,
      mil(x.enc),mil(x.det),pc(x.pdet),pc(x.nps),
      (x.br>0?'+':'')+String(x.br).replace('.',',')+' pts']),
    (r,c)=>{ const x=A.top[r];
      if(c===5) return x.br<=-10?P.bad:P.wn;
      if(c===4) return x.br<=-10?P.bad:P.dk;
      return P.dk; });
  s.addShape(pres.ShapeType.roundRect,{x:M,y:fin+0.24,w:W-2*M,h:1.0,rectRadius:0.09,
    fill:{color:P.lilaLt},line:{color:P.ln,width:1}});
  s.addText([
    {text:'Criterio.  ',options:{bold:true,color:P.azul}},
    {text:`Cada agente se compara contra el NPS de su propio skill (${pc(A.base)}), no contra el total. La columna Diferencia indica cuántos puntos está por encima o por debajo de ese valor. En ${sk} hay ${A.debajo} agentes de ${A.n} por debajo del promedio del skill; se listan los diez con menor NPS. Se exige un mínimo de 50 encuestas para evitar que un puñado de respuestas determine la posición.`,
     options:{color:P.dk}}
  ],{x:M+0.3,y:fin+0.4,w:W-2*M-0.6,h:0.72,isTextBox:true,margin:0,
     fontFace:F.b,fontSize:10.5,lineSpacingMultiple:1.2});
  fuente(s,'No se utiliza el campo Supervisor: en la base del período figura sin asignar en la totalidad de los registros.');
});

/* ── 9. CONCLUSIONES ──────────────────────────────────────── */
{
  const s=pres.addSlide(); fondo(s,true);
  s.addShape(pres.ShapeType.ellipse,{x:10.6,y:-2.1,w:5.6,h:5.6,
    fill:{color:P.azul2,transparency:62},line:{width:0}});
  s.addText('CONCLUSIONES',{x:M,y:0.68,w:8,h:0.3,isTextBox:true,margin:0,
    fontFace:F.b,fontSize:11.5,bold:true,color:P.lila,charSpacing:3});
  s.addText('Hallazgos del período',{x:M,y:1.0,w:9,h:0.62,isTextBox:true,margin:0,
    fontFace:F.h,fontSize:30,bold:true,color:P.w});
  const items=[
    ['Chat concentra el deterioro',
     `Chat cierra el período en ${pc(R.resumen.Chat.nps)} frente a ${pc(R.resumen.Inbound.nps)} de Inbound, y cayó ${String(Math.abs(R.resumen.Chat.mes[2].nps-R.resumen.Chat.mes[1].nps).toFixed(1)).replace('.',',')} puntos entre julio y agosto. Inbound se mantuvo estable.`],
    ['Una tipificación explica la mayor parte',
     `«Soporte y acompañamiento en la Banca» aporta ${mil(R.tipif.Chat.sub_vol[0].det)} detractores en Chat sobre ${mil(R.tipif.Chat.sub_vol[0].enc)} encuestas, con NPS de ${pc(R.tipif.Chat.sub_vol[0].nps)}.`],
    ['Banca Persona Natural es el proceso de mayor impacto',
     `Aporta ${pc(R.tipif.Inbound.proc_vol[0].pct)} de los detractores de Inbound y ${pc(R.tipif.Chat.proc_vol[0].pct)} de los de Chat, por volumen de gestión.`],
    ['La brecha entre agentes es mayor en Chat',
     `En Chat el rango va de ${pc(R.agentes.Chat.top[0].nps)} a valores cercanos al promedio; en Inbound la dispersión es menor. ${R.agentes.Chat.debajo} de ${R.agentes.Chat.n} agentes de Chat están bajo el promedio de su skill.`]
  ];
  let y=1.94;
  items.forEach((it,i)=>{
    s.addShape(pres.ShapeType.ellipse,{x:M,y,w:0.34,h:0.34,
      fill:{color:P.lila},line:{width:0}});
    s.addText(String(i+1),{x:M,y:y+0.02,w:0.34,h:0.3,isTextBox:true,margin:0,
      align:'center',fontFace:F.b,fontSize:12,bold:true,color:P.azul});
    s.addText(it[0],{x:M+0.54,y:y-0.03,w:9.2,h:0.3,isTextBox:true,margin:0,
      fontFace:F.b,fontSize:14,bold:true,color:P.w});
    s.addText(it[1],{x:M+0.54,y:y+0.28,w:9.2,h:0.6,isTextBox:true,margin:0,
      fontFace:F.b,fontSize:11,color:'C4B5FD',lineSpacingMultiple:1.2});
    y+=1.06;
  });
  s.addShape(pres.ShapeType.roundRect,{x:M,y:6.36,w:W-2*M,h:0.6,rectRadius:0.09,
    fill:{color:P.azul2},line:{width:0}});
  s.addText(`${mil(R.global.enc)} encuestas · junio a agosto de 2026 · Inbound ${mil(R.resumen.Inbound.enc)} · Chat ${mil(R.resumen.Chat.enc)}`,
    {x:M+0.3,y:6.49,w:W-2*M-0.6,h:0.34,isTextBox:true,margin:0,
     fontFace:F.b,fontSize:10.5,color:'C4B5FD'});
}

/* ── 10. NOTA METODOLÓGICA ────────────────────────────────── */
{
  const s=pres.addSlide(); fondo(s);
  cab(s,'Nota metodológica','Alcance de las cifras y diferencia con las hojas de resumen del archivo');
  const gap=0.35, w=(W-2*M-gap)/2;
  const bl=[
    ['Fuente de los datos',
     'Hoja «BBDD Encuestas» de los consolidados de junio, julio y agosto de 2026: 40.738 registros individuales de encuesta.'],
    ['Diferencia con las demás hojas',
     'Las hojas de resumen del archivo (Detallado NPS Agente, Detallado NPS Proceso, Profundización) provienen de una tabla dinámica cuya caché contiene 123.031 registros de 2025 y 2026. No corresponden al período de este informe y por eso sus cifras difieren.'],
    ['Cálculo del indicador',
     'NPS = porcentaje de promotores menos porcentaje de detractores sobre las encuestas con clasificación. Los neutros integran la base pero no suman ni restan.'],
    ['Escalas por skill',
     'Inbound califica de 0 a 9 (promotor 8-9) y Chat de 0 a 10 (promotor 9-10). Una misma nota no equivale a la misma categoría en los dos skills, por lo que no se comparan notas entre skills ni se emplea la nota cruda.']
  ];
  bl.forEach((b,i)=>{
    const x=M+(i%2)*(w+gap), y=1.5+Math.floor(i/2)*2.15;
    s.addShape(pres.ShapeType.roundRect,{x,y,w,h:1.92,rectRadius:0.09,
      fill:{color:P.w},line:{color:P.ln,width:1}});
    s.addText(b[0].toUpperCase(),{x:x+0.3,y:y+0.22,w:w-0.6,h:0.26,isTextBox:true,
      margin:0,fontFace:F.b,fontSize:10,bold:true,color:P.azul3,charSpacing:1.6});
    s.addText(b[1],{x:x+0.3,y:y+0.54,w:w-0.6,h:1.24,isTextBox:true,margin:0,
      fontFace:F.b,fontSize:11,color:P.txt,lineSpacingMultiple:1.25});
  });
  s.addText('Mínimos aplicados: 50 encuestas por agente; 120 encuestas por proceso y 100 por subproceso en Inbound; 40 en Chat, dado su menor volumen.',
    {x:M,y:5.9,w:W-2*M,h:0.4,isTextBox:true,margin:0,
     fontFace:F.b,fontSize:10.5,color:P.dk});
  fuente(s,'Elaborado por la supervisión de Skill Inbound · agosto de 2026');
}

pres.writeFile({fileName:'NPS_Junio_Agosto_2026.pptx'}).then(f=>console.log('generado:',f));
