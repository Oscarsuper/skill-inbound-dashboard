const pptxgen = require('pptxgenjs');
const D = require('../nps/datos.json');

const P = {
  azul:'2B0B6B', azul2:'4A1D96', azul3:'7C3AED', lila:'A78BFA',
  lilaLt:'EDE9FE', bg:'F8F5FF', w:'FFFFFF', ln:'DDD6FE',
  txt:'6B7280', dk:'1E1235',
  ok:'16A34A', wn:'D97706', bad:'DC2626'
};
const F = { h:'Bookman Old Style', b:'Calibri' };

const pres = new pptxgen();
pres.layout = 'LAYOUT_WIDE';                 // 13.3 x 7.5
pres.author = 'Supervisión Skill Inbound';
pres.title  = 'NPS Banco Agrario — Junio a Agosto 2026';

const W = 13.3, H = 7.5, M = 0.7;

// En español el separador decimal es la coma: 67.2 -> "67,2"
const dec = n => String(n).replace('.', ',');
// Siempre un decimal: 65 debe leerse «65,0%», no «65%»
const pc  = n => Number(n).toFixed(1).replace('.', ',') + '%';
const mil = n => Number(n).toLocaleString('es-CO');

/* ── piezas reutilizables ───────────────────────────────────── */
function fondo(s, oscuro){
  s.background = { color: oscuro ? P.azul : P.bg };
}
function titulo(s, t, sub){
  s.addText(t, { x:M, y:0.42, w:W-2*M, h:0.62, isTextBox:true, margin:0,
    fontFace:F.h, fontSize:30, bold:true, color:P.azul });
  if(sub) s.addText(sub, { x:M, y:1.06, w:W-2*M, h:0.36, isTextBox:true, margin:0,
    fontFace:F.b, fontSize:13, color:P.txt });
}
function pie(s, txt){
  s.addText(txt, { x:M, y:H-0.52, w:W-2*M, h:0.3, isTextBox:true, margin:0,
    fontFace:F.b, fontSize:9.5, color:P.txt, italic:true });
}
// Tarjeta de dato grande
function stat(s, x, y, w, valor, etiqueta, color, nota){
  s.addShape(pres.ShapeType.roundRect, { x, y, w, h:1.62, rectRadius:0.09,
    fill:{ color:P.w }, line:{ color:P.ln, width:1 },
    shadow:{ type:'outer', color:'4A1D96', opacity:0.10, blur:8, offset:2, angle:90 } });
  s.addText(valor, { x, y:y+0.16, w, h:0.72, isTextBox:true, margin:0, align:'center',
    fontFace:F.h, fontSize:38, bold:true, color:color||P.azul2 });
  s.addText(etiqueta, { x, y:y+0.92, w, h:0.28, isTextBox:true, margin:0, align:'center',
    fontFace:F.b, fontSize:11, bold:true, color:P.dk });
  if(nota) s.addText(nota, { x, y:y+1.19, w, h:0.3, isTextBox:true, margin:0, align:'center',
    fontFace:F.b, fontSize:9.5, color:P.txt });
}
// Punto de una lista con círculo numerado
function punto(s, x, y, w, n, titulo_, cuerpo){
  s.addShape(pres.ShapeType.ellipse, { x, y, w:0.34, h:0.34,
    fill:{ color:P.azul3 }, line:{ color:P.azul3, width:0 } });
  s.addText(String(n), { x, y:y+0.02, w:0.34, h:0.3, isTextBox:true, margin:0,
    align:'center', fontFace:F.b, fontSize:12, bold:true, color:P.w });
  s.addText(titulo_, { x:x+0.5, y:y-0.03, w:w-0.5, h:0.3, isTextBox:true, margin:0,
    fontFace:F.b, fontSize:13.5, bold:true, color:P.dk });
  s.addText(cuerpo, { x:x+0.5, y:y+0.27, w:w-0.5, h:0.62, isTextBox:true, margin:0,
    fontFace:F.b, fontSize:11.5, color:P.txt, lineSpacingMultiple:1.15 });
}
const opcChart = {
  showLegend:false, showTitle:false,
  catAxisLabelColor:P.txt, catAxisLabelFontFace:F.b, catAxisLabelFontSize:10,
  valAxisLabelColor:P.txt, valAxisLabelFontFace:F.b, valAxisLabelFontSize:10,
  valGridLine:{ color:P.ln, size:1 }, catGridLine:{ style:'none' },
  dataLabelFontFace:F.b, dataLabelFontSize:10, dataLabelColor:P.dk,
  showValue:true
};

/* ── 1. PORTADA ─────────────────────────────────────────────── */
{
  const s = pres.addSlide(); fondo(s, true);
  s.addShape(pres.ShapeType.ellipse, { x:9.6, y:-1.9, w:6.2, h:6.2,
    fill:{ color:P.azul2, transparency:55 }, line:{ width:0 } });
  s.addShape(pres.ShapeType.ellipse, { x:11.2, y:3.9, w:4.4, h:4.4,
    fill:{ color:P.azul3, transparency:70 }, line:{ width:0 } });
  s.addText('ANÁLISIS DE INDICADOR', { x:M, y:2.05, w:9, h:0.34, isTextBox:true, margin:0,
    fontFace:F.b, fontSize:13, bold:true, color:P.lila, charSpacing:3 });
  s.addText('NPS Banco Agrario', { x:M, y:2.42, w:9.4, h:0.95, isTextBox:true, margin:0,
    fontFace:F.h, fontSize:50, bold:true, color:P.w });
  s.addText('Junio · Julio · Agosto 2026', { x:M, y:3.42, w:9.4, h:0.5, isTextBox:true, margin:0,
    fontFace:F.h, fontSize:24, color:P.lila });
  s.addText('Motivo de las malas calificaciones por macroproceso y proceso,\nagentes con menor NPS y comparativa mes a mes.',
    { x:M, y:4.12, w:8.2, h:0.8, isTextBox:true, margin:0,
      fontFace:F.b, fontSize:14, color:'C4B5FD', lineSpacingMultiple:1.3 });
  s.addText(`${D.total.Encuestas.toLocaleString('es-CO')} encuestas analizadas  ·  ${D.nagentes} agentes evaluados`,
    { x:M, y:5.28, w:9, h:0.34, isTextBox:true, margin:0,
      fontFace:F.b, fontSize:12, bold:true, color:P.lila, charSpacing:1 });
  s.addNotes('Base: hoja BBDD Encuestas de los tres consolidados mensuales. 40.738 encuestas con clasificación NPS.');
}

/* ── 2. RESUMEN ─────────────────────────────────────────────── */
{
  const s = pres.addSlide(); fondo(s);
  titulo(s, 'Resumen de los tres meses', 'Junio a agosto de 2026 · base completa de encuestas');
  const gap=0.28, w=(W-2*M-3*gap)/4;
  stat(s, M,               1.68, w, pc(D.total.NPS),
       'NPS del período', P.azul2, 'Promotores menos detractores');
  stat(s, M+(w+gap),       1.68, w, D.total.Encuestas.toLocaleString('es-CO'),
       'Encuestas', P.azul2, 'Voz y chat');
  stat(s, M+2*(w+gap),     1.68, w, D.total.Detractores.toLocaleString('es-CO'),
       'Detractores', P.bad, pc(D.total['%Detractores'])+' del total');
  stat(s, M+3*(w+gap),     1.68, w, D.ndebajo+' de '+D.nagentes,
       'Agentes bajo el promedio', P.wn, 'Con 50 o más encuestas');

  s.addShape(pres.ShapeType.roundRect, { x:M, y:3.62, w:W-2*M, h:2.62, rectRadius:0.1,
    fill:{ color:P.w }, line:{ color:P.ln, width:1 } });
  s.addText('Lo que hay que saber', { x:M+0.35, y:3.82, w:5, h:0.3, isTextBox:true, margin:0,
    fontFace:F.b, fontSize:12, bold:true, color:P.azul3, charSpacing:1.5 });
  const cw=(W-2*M-0.7-0.5)/2;
  punto(s, M+0.35, 4.24, cw, 1, 'Chat arrastra el indicador',
    `Chat cierra en ${pc(D.npsc.Chat)} y voz en ${pc(D.npsc.Voz)}: 33 puntos de diferencia con la misma operación.`);
  punto(s, M+0.35, 5.22, cw, 2, 'Un solo subproceso explica el 28%',
    `«Soporte y acompañamiento en la Banca» aporta ${D.vol[0].det.toLocaleString('es-CO')} de los ${D.total.Detractores.toLocaleString('es-CO')} detractores.`);
  punto(s, M+0.35+cw+0.5, 4.24, cw, 3, 'Agosto rompió la tendencia',
    'Julio venía mejorando (68,6%) y agosto cayó a 65,0%. La caída es de chat, no de voz.');
  punto(s, M+0.35+cw+0.5, 5.22, cw, 4, 'El ranking absoluto es injusto',
    `${D.debajo_chat} de los ${D.ndebajo} agentes bajo el promedio son de chat, siendo solo ${D.agentes_chat} de ${D.nagentes}.`);
  pie(s, 'NPS = % promotores − % detractores, sobre encuestas con clasificación.');
}

/* ── 3. MES A MES ───────────────────────────────────────────── */
{
  const s = pres.addSlide(); fondo(s);
  titulo(s, 'Comparativa mes a mes', 'El total se sostiene, pero la brecha entre canales se abre');
  s.addChart(pres.ChartType.bar, [
    { name:'Voz',   labels:D.meses.map(m=>m.mes), values:D.meses.map(m=>m.voz) },
    { name:'Chat',  labels:D.meses.map(m=>m.mes), values:D.meses.map(m=>m.chat) },
    { name:'Total', labels:D.meses.map(m=>m.mes), values:D.meses.map(m=>m.nps) }
  ], Object.assign({}, opcChart, {
    x:M, y:1.62, w:7.9, h:4.05,
    barDir:'col', barGrouping:'clustered',
    chartColors:[P.azul2, P.bad, P.lila],
    valAxisMinVal:0, valAxisMaxVal:80,
    dataLabelPosition:'outEnd', dataLabelFormatCode:'0.0"%"',
    showLegend:true, legendPos:'b', legendFontFace:F.b, legendFontSize:11, legendColor:P.dk
  }));
  const bx=8.86, bw=W-M-bx;
  s.addShape(pres.ShapeType.roundRect, { x:bx, y:1.62, w:bw, h:4.05, rectRadius:0.1,
    fill:{ color:P.w }, line:{ color:P.ln, width:1 } });
  s.addText('Detalle por mes', { x:bx+0.28, y:1.84, w:bw-0.56, h:0.28, isTextBox:true, margin:0,
    fontFace:F.b, fontSize:12, bold:true, color:P.azul3, charSpacing:1.2 });
  let yy=2.24;
  D.meses.forEach(m=>{
    s.addText(m.mes, { x:bx+0.28, y:yy, w:1.36, h:0.28, isTextBox:true, margin:0,
      fontFace:F.b, fontSize:14, bold:true, color:P.dk });
    s.addText(pc(m.nps), { x:bx+1.7, y:yy, w:bw-1.98, h:0.28, isTextBox:true, margin:0,
      align:'right', fontFace:F.h, fontSize:16, bold:true,
      color: m.nps>=68 ? P.ok : (m.nps>=66 ? P.wn : P.bad) });
    s.addText(`${m.enc.toLocaleString('es-CO')} encuestas · ${m.det.toLocaleString('es-CO')} detractores`,
      { x:bx+0.28, y:yy+0.29, w:bw-0.56, h:0.24, isTextBox:true, margin:0,
        fontFace:F.b, fontSize:10, color:P.txt });
    s.addText(`Voz ${pc(m.voz)}  ·  Chat ${pc(m.chat)}`,
      { x:bx+0.28, y:yy+0.52, w:bw-0.56, h:0.24, isTextBox:true, margin:0,
        fontFace:F.b, fontSize:10, bold:true, color:P.azul2 });
    yy += 1.12;
  });
  s.addText('Chat cayó 9,5 puntos de julio a agosto. Voz apenas 1,7.',
    { x:bx+0.28, y:5.24, w:bw-0.56, h:0.34, isTextBox:true, margin:0,
      fontFace:F.b, fontSize:10.5, italic:true, color:P.bad });
  pie(s, 'Junio 14.616 encuestas · Julio 15.933 · Agosto 10.189.');
  s.addNotes('Agosto tiene menos encuestas porque el corte va al día 20.');
}

/* ── 4. LA BRECHA DE CANAL ──────────────────────────────────── */
{
  const s = pres.addSlide(); fondo(s);
  titulo(s, 'Voz y chat no juegan el mismo partido',
    'Misma operación, 33 puntos de diferencia — es el hallazgo más grande del período');
  const gap=0.4, w=(W-2*M-gap)/2;
  D.canal.forEach((c,i)=>{
    const x=M+i*(w+gap), esChat=c.canal==='Chat';
    s.addShape(pres.ShapeType.roundRect, { x, y:1.66, w, h:2.5, rectRadius:0.1,
      fill:{ color: esChat ? 'FEF2F2' : 'F0FDF4' },
      line:{ color: esChat ? 'FECACA' : 'BBF7D0', width:1.5 } });
    s.addText(c.canal.toUpperCase(), { x:x+0.34, y:1.9, w:w-0.68, h:0.3, isTextBox:true, margin:0,
      fontFace:F.b, fontSize:13, bold:true, color: esChat?P.bad:P.ok, charSpacing:2.5 });
    s.addText(pc(c.nps), { x:x+0.34, y:2.24, w:w-0.68, h:0.95, isTextBox:true, margin:0,
      fontFace:F.h, fontSize:52, bold:true, color: esChat?P.bad:P.ok });
    s.addText(`${c.enc.toLocaleString('es-CO')} encuestas   ·   ${c.det.toLocaleString('es-CO')} detractores (${pc(c.pdet)})`,
      { x:x+0.34, y:3.32, w:w-0.68, h:0.3, isTextBox:true, margin:0,
        fontFace:F.b, fontSize:12, color:P.dk });
    s.addText(esChat ? 'Uno de cada cuatro clientes atendidos por chat queda insatisfecho.'
                     : 'Uno de cada nueve. Voz sostiene el indicador de toda la operación.',
      { x:x+0.34, y:3.64, w:w-0.68, h:0.4, isTextBox:true, margin:0,
        fontFace:F.b, fontSize:11, italic:true, color:P.txt });
  });
  s.addShape(pres.ShapeType.roundRect, { x:M, y:4.36, w:W-2*M, h:1.86, rectRadius:0.1,
    fill:{ color:P.w }, line:{ color:P.ln, width:1 } });
  s.addText('Dónde se rompe el chat', { x:M+0.34, y:4.56, w:6, h:0.3, isTextBox:true, margin:0,
    fontFace:F.b, fontSize:12, bold:true, color:P.azul3, charSpacing:1.2 });
  let xx=M+0.34;
  const cw2=(W-2*M-0.68-0.9)/4;
  D.chat_peor.slice(0,4).forEach(c=>{
    s.addText(pc(c.nps), { x:xx, y:4.92, w:cw2, h:0.44, isTextBox:true, margin:0,
      fontFace:F.h, fontSize:24, bold:true, color:P.bad });
    s.addText(c.s, { x:xx, y:5.36, w:cw2, h:0.5, isTextBox:true, margin:0,
      fontFace:F.b, fontSize:10.5, bold:true, color:P.dk });
    s.addText(c.enc.toLocaleString('es-CO')+' encuestas', { x:xx, y:5.86, w:cw2, h:0.24,
      isTextBox:true, margin:0, fontFace:F.b, fontSize:9.5, color:P.txt });
    xx += cw2+0.3;
  });
  pie(s, 'NPS de cada subproceso dentro del canal chat, mínimo 80 encuestas.');
}


/* ── 5. MACROPROCESO › PROCESO — VOLUMEN ────────────────────── */
{
  const s = pres.addSlide(); fondo(s);
  titulo(s, 'Macroproceso y proceso que más afectan',
    'Por volumen: dónde se concentran los detractores. Arreglar arriba es lo que más mueve el indicador.');
  const filas = D.mp_vol.slice(0,8).map(r=>({
    p: r['MACROPROCESO › PROCESO'].replace('Banco Agrario de Colombia  ›  ','').trim(),
    det: r.Detrac, pct: r['% det'], enc: r.Enc, tasa: r.Tasa, nps: r.NPS
  }));
  s.addChart(pres.ChartType.bar, [
    { name:'Detractores', labels:filas.map(f=>f.p), values:filas.map(f=>f.det) }
  ], Object.assign({}, opcChart, {
    x:M, y:1.68, w:6.5, h:4.3, barDir:'bar',
    chartColors:[P.azul2], valAxisMinVal:0,
    dataLabelPosition:'outEnd', catAxisLabelFontSize:9.5, barGapWidthPct:45
  }));
  const bx=7.5, bw=W-M-bx;
  s.addText('% de todos los detractores', { x:bx, y:1.68, w:bw, h:0.28, isTextBox:true, margin:0,
    fontFace:F.b, fontSize:11, bold:true, color:P.azul3, charSpacing:1 });
  let yy=2.06;
  filas.forEach(f=>{
    const ancho=(bw-1.55)*(f.pct/filas[0].pct);
    s.addShape(pres.ShapeType.roundRect, { x:bx, y:yy+0.05, w:Math.max(ancho,0.06), h:0.2,
      rectRadius:0.05, fill:{ color: f.tasa>=17 ? P.bad : P.lila }, line:{ width:0 } });
    s.addText(pc(f.pct), { x:bx+bw-1.5, y:yy-0.02, w:0.7, h:0.28, isTextBox:true, margin:0,
      align:'right', fontFace:F.h, fontSize:12, bold:true, color:P.dk });
    s.addText(pc(f.tasa), { x:bx+bw-0.72, y:yy-0.02, w:0.72, h:0.28, isTextBox:true, margin:0,
      align:'right', fontFace:F.b, fontSize:10.5, bold:true,
      color: f.tasa>=17 ? P.bad : P.txt });
    yy += 0.53;
  });
  s.addText('Barra: peso sobre el total de detractores.   Cifra gris: qué tanto de ese proceso califica mal.',
    { x:bx, y:yy+0.06, w:bw, h:0.5, isTextBox:true, margin:0,
      fontFace:F.b, fontSize:9.5, italic:true, color:P.txt, lineSpacingMultiple:1.2 });
  pie(s, 'El macroproceso "Banco Agrario de Colombia" agrupa el 92% de las encuestas, por eso se muestra el proceso.');
  s.addNotes('Banca Pnatural aporta 2.089 detractores, el 38,7% del total del período.');
}

/* ── 6. MACROPROCESO › PROCESO — PEOR NPS ───────────────────── */
{
  const s = pres.addSlide(); fondo(s);
  titulo(s, 'Dónde peor nos califican',
    'Por tasa: el NPS de cada macroproceso y proceso, de peor a mejor. Mínimo 150 encuestas.');
  const f = D.mp_tasa.slice(0,9);
  const enc = f.map(r=>r.mp.split('  ›  '));
  s.addChart(pres.ChartType.bar, [
    { name:'NPS', labels:enc.map(e=>e[1]||e[0]), values:f.map(r=>r.nps) }
  ], Object.assign({}, opcChart, {
    x:M, y:1.7, w:7.5, h:4.4, barDir:'bar',
    chartColors:[P.bad], valAxisMinVal:0, valAxisMaxVal:80,
    dataLabelPosition:'outEnd', dataLabelFormatCode:'0.0"%"',
    catAxisLabelFontSize:9.5, barGapWidthPct:45
  }));
  const bx=8.45, bw=W-M-bx;
  s.addShape(pres.ShapeType.roundRect, { x:bx, y:1.7, w:bw, h:2.3, rectRadius:0.1,
    fill:{ color:P.w }, line:{ color:P.ln, width:1 } });
  s.addText('Referencia', { x:bx+0.28, y:1.9, w:bw-0.56, h:0.28, isTextBox:true, margin:0,
    fontFace:F.b, fontSize:11, bold:true, color:P.azul3, charSpacing:1.2 });
  s.addText(pc(D.total.NPS), { x:bx+0.28, y:2.2, w:bw-0.56, h:0.7, isTextBox:true, margin:0,
    fontFace:F.h, fontSize:34, bold:true, color:P.azul2 });
  s.addText('NPS general del período.\nTodo lo del gráfico está por debajo.',
    { x:bx+0.28, y:2.94, w:bw-0.56, h:0.7, isTextBox:true, margin:0,
      fontFace:F.b, fontSize:11, color:P.txt, lineSpacingMultiple:1.2 });
  s.addShape(pres.ShapeType.roundRect, { x:bx, y:4.2, w:bw, h:1.9, rectRadius:0.1,
    fill:{ color:'FEF2F2' }, line:{ color:'FECACA', width:1.5 } });
  s.addText('Los dos peores', { x:bx+0.28, y:4.4, w:bw-0.56, h:0.28, isTextBox:true, margin:0,
    fontFace:F.b, fontSize:11, bold:true, color:P.bad, charSpacing:1.2 });
  s.addText([
    { text:'Cancelación no efectiva', options:{ bold:true, fontSize:12, color:P.dk, breakLine:true } },
    { text:'NPS 41,0% · 178 encuestas · 23,0% detractores', options:{ fontSize:10, color:P.txt, breakLine:true } },
    { text:' ', options:{ fontSize:5, breakLine:true } },
    { text:'Aplicativos Banco', options:{ bold:true, fontSize:12, color:P.dk, breakLine:true } },
    { text:'NPS 46,9% · 196 encuestas · indisponibilidad de canales', options:{ fontSize:10, color:P.txt } }
  ], { x:bx+0.28, y:4.7, w:bw-0.56, h:1.3, isTextBox:true, margin:0,
       fontFace:F.b, lineSpacingMultiple:1.15 });
  pie(s, 'Se exige un mínimo de 150 encuestas para que el dato no dependa de unas pocas respuestas.');
}

/* ── 7. SUBPROCESOS ─────────────────────────────────────────── */
{
  const s = pres.addSlide(); fondo(s);
  titulo(s, 'El detalle fino: subprocesos',
    'Los ocho que más detractores aportan en los tres meses');
  const cab=['SUBPROCESO','ENCUESTAS','DETRACTORES','% DEL TOTAL','TASA','NPS'];
  const anchos=[5.5,1.35,1.5,1.35,1.05,1.15];
  let x0=M;
  cab.forEach((c,i)=>{
    s.addShape(pres.ShapeType.rect, { x:x0, y:1.66, w:anchos[i], h:0.42,
      fill:{ color:P.azul2 }, line:{ width:0 } });
    s.addText(c, { x:x0+0.1, y:1.7, w:anchos[i]-0.2, h:0.34, isTextBox:true, margin:0,
      align: i===0?'left':'center', fontFace:F.b, fontSize:9, bold:true,
      color:P.w, charSpacing:0.8 });
    x0 += anchos[i];
  });
  let y=2.08;
  D.vol.slice(0,8).forEach((r,i)=>{
    const alto=0.46;
    s.addShape(pres.ShapeType.rect, { x:M, y, w:W-2*M, h:alto,
      fill:{ color: i%2 ? P.w : 'FCFAFF' }, line:{ color:P.ln, width:0.5 } });
    const vals=[r.sub, r.enc.toLocaleString('es-CO'), r.det.toLocaleString('es-CO'),
                pc(r.pct), pc(r.tasa), pc(r.nps)];
    let x=M;
    vals.forEach((v,j)=>{
      const rojo = (j===4 && r.tasa>=16) || (j===5 && r.nps<62);
      s.addText(v, { x:x+0.1, y:y+0.05, w:anchos[j]-0.2, h:0.36, isTextBox:true, margin:0,
        align: j===0?'left':'center', fontFace: j===0?F.b:F.h,
        fontSize: j===0?11:11.5, bold: j!==0,
        color: rojo ? P.bad : (j===0 ? P.dk : P.dk) });
      x += anchos[j];
    });
    y += alto;
  });
  s.addText('«Soporte y acompañamiento en la Banca» concentra 1.497 detractores: más que los siete siguientes juntos.',
    { x:M, y:y+0.18, w:W-2*M, h:0.34, isTextBox:true, margin:0,
      fontFace:F.b, fontSize:12, bold:true, italic:true, color:P.azul2 });
  pie(s, 'Tasa = porcentaje de encuestas de ese subproceso que califican como detractor.');
}

/* ── 8. AGENTES: RANKING ABSOLUTO ───────────────────────────── */
{
  const s = pres.addSlide(); fondo(s);
  titulo(s, 'Agentes con peor NPS',
    `Consolidado de los tres meses · ${D.nagentes} agentes con 50 o más encuestas`);
  const cab=['#','AGENTE','CANAL','ENCUESTAS','DETRACTORES','% DETRAC.','NPS'];
  const an=[0.5,4.85,1.05,1.35,1.5,1.3,1.35];
  let x0=M;
  cab.forEach((c,i)=>{
    s.addShape(pres.ShapeType.rect, { x:x0, y:1.62, w:an[i], h:0.4,
      fill:{ color:P.azul2 }, line:{ width:0 } });
    s.addText(c, { x:x0+0.08, y:1.66, w:an[i]-0.16, h:0.32, isTextBox:true, margin:0,
      align: i===1?'left':'center', fontFace:F.b, fontSize:8.5, bold:true,
      color:P.w, charSpacing:0.7 });
    x0 += an[i];
  });
  let y=2.02;
  D.abs.slice(0,10).forEach((r,i)=>{
    const alto=0.4, top3=i<3;
    s.addShape(pres.ShapeType.rect, { x:M, y, w:W-2*M, h:alto,
      fill:{ color: top3 ? 'FEF2F2' : (i%2 ? P.w : 'FCFAFF') },
      line:{ color:P.ln, width:0.5 } });
    const vals=[String(i+1), r.a, r.c, r.enc.toLocaleString('es-CO'),
                r.det.toLocaleString('es-CO'), pc(r.pdet), pc(r.nps)];
    let x=M;
    vals.forEach((v,j)=>{
      s.addText(v, { x:x+0.08, y:y+0.04, w:an[j]-0.16, h:0.32, isTextBox:true, margin:0,
        align: j===1?'left':'center', fontFace: j===1?F.b:F.h,
        fontSize: j===1?10.5:11, bold: j!==1,
        color: j===6 ? (r.nps<35?P.bad:P.wn) : (j===5 && r.pdet>=28 ? P.bad : P.dk) });
      x += an[j];
    });
    y += alto;
  });
  s.addShape(pres.ShapeType.roundRect, { x:M, y:y+0.16, w:W-2*M, h:0.86, rectRadius:0.09,
    fill:{ color:'FEF3C7' }, line:{ color:'FDE68A', width:1.5 } });
  s.addText([
    { text:'Ojo con la lectura.  ', options:{ bold:true, color:'92400E' } },
    { text:`${D.debajo_chat} de los ${D.ndebajo} agentes bajo el promedio son de chat, cuando chat solo tiene ${D.agentes_chat} de los ${D.nagentes} agentes. No atienden peor: el canal parte 33 puntos abajo. La comparación justa está en la lámina siguiente.`,
      options:{ color:'92400E' } }
  ], { x:M+0.3, y:y+0.28, w:W-2*M-0.6, h:0.62, isTextBox:true, margin:0,
       fontFace:F.b, fontSize:11, lineSpacingMultiple:1.2 });
  pie(s, `Se listan los 10 peores. En total ${D.ndebajo} agentes están por debajo del NPS general (${pc(D.total.NPS)}).`);
  s.addNotes('El archivo CONSOLIDADO_agentes.csv trae los 106 agentes con la marca de quién está bajo el promedio.');
}

/* ── 9. AGENTES CONTRA SU PROPIO CANAL ──────────────────────── */
{
  const s = pres.addSlide(); fondo(s);
  titulo(s, 'La comparación justa: cada agente contra su canal',
    `Voz se compara con ${pc(D.npsc.Voz)} y chat con ${pc(D.npsc.Chat)} — así se ve quién de verdad está por debajo`);
  const f=D.justo.slice(0,10);
  const cab=['AGENTE','CANAL','ENCUESTAS','SU NPS','META DEL CANAL','DIFERENCIA'];
  const an=[4.9,1.1,1.4,1.35,1.85,1.8];
  let x0=M;
  cab.forEach((c,i)=>{
    s.addShape(pres.ShapeType.rect, { x:x0, y:1.66, w:an[i], h:0.4,
      fill:{ color:P.azul2 }, line:{ width:0 } });
    s.addText(c, { x:x0+0.08, y:1.7, w:an[i]-0.16, h:0.32, isTextBox:true, margin:0,
      align: i===0?'left':'center', fontFace:F.b, fontSize:8.5, bold:true,
      color:P.w, charSpacing:0.7 });
    x0 += an[i];
  });
  let y=2.06;
  f.forEach((r,i)=>{
    const alto=0.42;
    s.addShape(pres.ShapeType.rect, { x:M, y, w:W-2*M, h:alto,
      fill:{ color: r.br<=-15 ? 'FEF2F2' : (i%2 ? P.w : 'FCFAFF') },
      line:{ color:P.ln, width:0.5 } });
    const meta = r.c==='Chat' ? D.npsc.Chat : D.npsc.Voz;
    const vals=[r.a, r.c, r.enc.toLocaleString('es-CO'), pc(r.nps), pc(meta),
                (r.br>0?'+':'')+dec(r.br)+' pts'];
    let x=M;
    vals.forEach((v,j)=>{
      s.addText(v, { x:x+0.08, y:y+0.05, w:an[j]-0.16, h:0.32, isTextBox:true, margin:0,
        align: j===0?'left':'center', fontFace: j===0?F.b:F.h,
        fontSize: j===0?10.5:11.5, bold: j!==0,
        color: j===5 ? (r.br<=-15?P.bad:P.wn) : P.dk });
      x += an[j];
    });
    y += alto;
  });
  const gap=0.35, cw=(W-2*M-gap)/2;
  s.addShape(pres.ShapeType.roundRect, { x:M, y:y+0.2, w:cw, h:1.06, rectRadius:0.09,
    fill:{ color:P.w }, line:{ color:P.ln, width:1 } });
  s.addText('Quiénes cambian de lugar', { x:M+0.28, y:y+0.34, w:cw-0.56, h:0.26,
    isTextBox:true, margin:0, fontFace:F.b, fontSize:11, bold:true, color:P.azul3 });
  s.addText('Daniel Felipe Castiblanco sube al 3.º puesto: su 46,6% en voz es peor que un 30% en chat.',
    { x:M+0.28, y:y+0.6, w:cw-0.56, h:0.58, isTextBox:true, margin:0,
      fontFace:F.b, fontSize:10.5, color:P.txt, lineSpacingMultiple:1.15 });
  s.addShape(pres.ShapeType.roundRect, { x:M+cw+gap, y:y+0.2, w:cw, h:1.06, rectRadius:0.09,
    fill:{ color:P.w }, line:{ color:P.ln, width:1 } });
  s.addText('Quiénes se confirman', { x:M+cw+gap+0.28, y:y+0.34, w:cw-0.56, h:0.26,
    isTextBox:true, margin:0, fontFace:F.b, fontSize:11, bold:true, color:P.azul3 });
  s.addText('Bejarano y Moreno siguen de primeras: −38,6 y −29,8 puntos contra su propio canal.',
    { x:M+cw+gap+0.28, y:y+0.6, w:cw-0.56, h:0.58, isTextBox:true, margin:0,
      fontFace:F.b, fontSize:10.5, color:P.txt, lineSpacingMultiple:1.15 });
  pie(s, 'Diferencia = NPS del agente menos el NPS de su canal. Ordenado de la brecha más negativa.');
}

/* ── 10. EN QUÉ FALLA CADA UNO ──────────────────────────────── */
{
  const s = pres.addSlide(); fondo(s);
  titulo(s, 'En qué se le van las malas notas a cada uno',
    'Subproceso donde se concentran los detractores de los cinco agentes con menor NPS');
  const gap=0.3, w=(W-2*M-2*gap)/3;
  D.causas.slice(0,3).forEach((c,i)=>{
    const x=M+i*(w+gap);
    s.addShape(pres.ShapeType.roundRect, { x, y:1.7, w, h:2.15, rectRadius:0.1,
      fill:{ color:P.w }, line:{ color:P.ln, width:1 } });
    s.addText(c.a, { x:x+0.26, y:1.9, w:w-0.52, h:0.56, isTextBox:true, margin:0,
      fontFace:F.b, fontSize:11, bold:true, color:P.dk });
    s.addText(c.tot+' detractores', { x:x+0.26, y:2.44, w:w-0.52, h:0.26, isTextBox:true,
      margin:0, fontFace:F.b, fontSize:10, color:P.bad, bold:true });
    let yy=2.76;
    c.top.forEach(t=>{
      s.addText(pc(t.p), { x:x+0.26, y:yy, w:0.72, h:0.26, isTextBox:true, margin:0,
        fontFace:F.h, fontSize:12, bold:true, color:P.azul2 });
      s.addText(t.s, { x:x+1.0, y:yy, w:w-1.26, h:0.34, isTextBox:true, margin:0,
        fontFace:F.b, fontSize:9.5, color:P.txt });
      yy += 0.37;
    });
  });
  D.causas.slice(3,5).forEach((c,i)=>{
    const w2=(W-2*M-gap)/2, x=M+i*(w2+gap);
    s.addShape(pres.ShapeType.roundRect, { x, y:4.02, w:w2, h:1.62, rectRadius:0.1,
      fill:{ color:P.w }, line:{ color:P.ln, width:1 } });
    s.addText(c.a, { x:x+0.26, y:4.2, w:w2-0.52, h:0.3, isTextBox:true, margin:0,
      fontFace:F.b, fontSize:11.5, bold:true, color:P.dk });
    s.addText(c.tot+' detractores', { x:x+0.26, y:4.48, w:w2-0.52, h:0.26, isTextBox:true,
      margin:0, fontFace:F.b, fontSize:10, bold:true, color:P.bad });
    let yy=4.78;
    c.top.forEach(t=>{
      s.addText(pc(t.p), { x:x+0.26, y:yy, w:0.72, h:0.26, isTextBox:true, margin:0,
        fontFace:F.h, fontSize:12, bold:true, color:P.azul2 });
      s.addText(t.s, { x:x+1.0, y:yy, w:w2-1.26, h:0.28, isTextBox:true, margin:0,
        fontFace:F.b, fontSize:9.5, color:P.txt });
      yy += 0.3;
    });
  });
  s.addText('Cuatro de los cinco fallan en el mismo sitio: «Soporte y acompañamiento en la Banca». No es un problema de persona, es del proceso.',
    { x:M, y:5.82, w:W-2*M, h:0.4, isTextBox:true, margin:0,
      fontFace:F.b, fontSize:12, bold:true, italic:true, color:P.azul2 });
  pie(s, 'Porcentaje sobre los detractores propios de cada agente.');
}

/* ── 11. CIERRE ─────────────────────────────────────────────── */
{
  const s = pres.addSlide(); fondo(s, true);
  s.addShape(pres.ShapeType.ellipse, { x:10.4, y:-2.2, w:6, h:6,
    fill:{ color:P.azul2, transparency:60 }, line:{ width:0 } });
  s.addText('CONCLUSIONES', { x:M, y:0.72, w:8, h:0.34, isTextBox:true, margin:0,
    fontFace:F.b, fontSize:12, bold:true, color:P.lila, charSpacing:3 });
  s.addText('Dónde poner el foco', { x:M, y:1.06, w:9, h:0.7, isTextBox:true, margin:0,
    fontFace:F.h, fontSize:34, bold:true, color:P.w });
  const items=[
    ['Chat, antes que agentes',
     `El canal está 33 puntos bajo voz y explica la caída de agosto. Con 6.887 encuestas concentra ${D.canal[1].det.toLocaleString('es-CO')} detractores.`],
    ['«Soporte y acompañamiento en la Banca»',
     `1.497 detractores, el 27,7% del total. En chat ese subproceso cierra en 22,7% de NPS sobre 2.113 encuestas.`],
    ['Cancelación no efectiva y Aplicativos Banco',
     'NPS 41,0% y 46,9%. Volumen bajo pero tasa de insatisfacción alta: se corrigen rápido y sin costo.'],
    ['Acompañar a 5 agentes, no a 29',
     'Contra su propio canal, solo cinco quedan más de 15 puntos abajo. Los demás están dentro del rango del canal.']
  ];
  let y=2.06;
  items.forEach((it,i)=>{
    s.addShape(pres.ShapeType.ellipse, { x:M, y, w:0.38, h:0.38,
      fill:{ color:P.lila }, line:{ width:0 } });
    s.addText(String(i+1), { x:M, y:y+0.03, w:0.38, h:0.32, isTextBox:true, margin:0,
      align:'center', fontFace:F.b, fontSize:13, bold:true, color:P.azul });
    s.addText(it[0], { x:M+0.58, y:y-0.03, w:8.9, h:0.32, isTextBox:true, margin:0,
      fontFace:F.b, fontSize:15, bold:true, color:P.w });
    s.addText(it[1], { x:M+0.58, y:y+0.3, w:8.9, h:0.62, isTextBox:true, margin:0,
      fontFace:F.b, fontSize:11.5, color:'C4B5FD', lineSpacingMultiple:1.2 });
    y += 1.08;
  });
  s.addShape(pres.ShapeType.roundRect, { x:M, y:6.42, w:W-2*M, h:0.62, rectRadius:0.09,
    fill:{ color:P.azul2 }, line:{ width:0 } });
  s.addText(`Base: ${D.total.Encuestas.toLocaleString('es-CO')} encuestas · junio a agosto de 2026 · hoja BBDD Encuestas de los tres consolidados mensuales`,
    { x:M+0.3, y:6.55, w:W-2*M-0.6, h:0.36, isTextBox:true, margin:0,
      fontFace:F.b, fontSize:10.5, color:'C4B5FD' });
}

/* ── 12. NOTA METODOLÓGICA ──────────────────────────────────── */
{
  const s = pres.addSlide(); fondo(s);
  titulo(s, 'Cómo se calculó', 'Para que cualquiera pueda reproducir estos números');
  const gap=0.35, w=(W-2*M-gap)/2;
  const bloques=[
    ['Fuente', 'Hoja «BBDD Encuestas» de los tres consolidados mensuales. Es la base cruda de respuestas; las demás hojas del archivo son resúmenes calculados sobre ella.'],
    ['Fórmula', 'NPS = % promotores − % detractores, sobre las encuestas con clasificación. Los neutros cuentan en la base pero no suman ni restan.'],
    ['Dos escalas', 'Voz califica de 0 a 9 (promotor 8–9) y chat de 0 a 10 (promotor 9–10). Un 8 es promotor en voz y neutro en chat, por eso se usa la clasificación del archivo y no la nota cruda.'],
    ['Mínimos', 'Agentes: 50 encuestas en los tres meses (106 de 141 califican). Procesos: 150 encuestas. Subprocesos: 100. Sin ese piso, un puñado de respuestas malas distorsiona el ranking.']
  ];
  bloques.forEach((b,i)=>{
    const x=M+(i%2)*(w+gap), y=1.7+Math.floor(i/2)*2.05;
    s.addShape(pres.ShapeType.roundRect, { x, y, w, h:1.82, rectRadius:0.1,
      fill:{ color:P.w }, line:{ color:P.ln, width:1 } });
    s.addText(b[0].toUpperCase(), { x:x+0.3, y:y+0.22, w:w-0.6, h:0.28, isTextBox:true,
      margin:0, fontFace:F.b, fontSize:11, bold:true, color:P.azul3, charSpacing:2 });
    s.addText(b[1], { x:x+0.3, y:y+0.56, w:w-0.6, h:1.1, isTextBox:true, margin:0,
      fontFace:F.b, fontSize:11.5, color:P.txt, lineSpacingMultiple:1.25 });
  });
  s.addText('Se entregan además los archivos con el detalle completo: los 106 agentes, el motivo por macroproceso, proceso y subproceso, y el evolutivo mensual.',
    { x:M, y:5.94, w:W-2*M, h:0.4, isTextBox:true, margin:0,
      fontFace:F.b, fontSize:11.5, italic:true, color:P.dk });
  pie(s, 'Preparado por la supervisión de Skill Inbound · agosto de 2026');
}

pres.writeFile({ fileName: 'NPS_Banco_Agrario_Jun_Ago_2026.pptx' })
  .then(f=>console.log('generado:', f));
