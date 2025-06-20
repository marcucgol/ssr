const fs   = require('fs');
const path = require('path');
const XLSX = require('xlsx');

function loadCandles(txtPath) {
  return fs.readFileSync(txtPath, 'utf8')
    .split(/\r?\n/)
    .filter(l => l.trim())
    .map(line => {
      const parts = line.split(',');
      if (parts.length < 7 || !/^\d{8}$/.test(parts[0]) || !/^\d{6}$/.test(parts[1]))
        return null;
      const [datePart, timePart, openRaw, highRaw, lowRaw, closeRaw, volRaw] = parts;
      const open   = parseFloat(openRaw);
      const high   = parseFloat(highRaw);
      const low    = parseFloat(lowRaw);
      const close  = parseFloat(closeRaw);
      const volume = parseFloat(volRaw);
      if ([open,high,low,close,volume].some(v=>isNaN(v))) return null;
      const iso = `${datePart.slice(0,4)}-${datePart.slice(4,6)}-${datePart.slice(6,8)}T`
                + `${timePart.slice(0,2)}:${timePart.slice(2,4)}:00Z`;
      return { iso, open, high, low, close, volume };
    })
    .filter(x => x);
}

(async () => {
  const EXCEL_IN  = 'processed_data.xlsx';
  const EXCEL_OUT = 'processed_data_with_entries.xlsx';
  const TXT_DIR   = __dirname;
  const PRECISION = 1000000000;
  const round5   = v => Math.round(v * PRECISION) / PRECISION;

  const candlesMap = {};
  fs.readdirSync(TXT_DIR).filter(f => f.endsWith('.txt'))
    .forEach(fn => {
      const inst = path.basename(fn, '.txt');
      candlesMap[inst] = loadCandles(path.join(TXT_DIR, fn));
      console.log(`Loaded ${candlesMap[inst].length} bars for ${inst}`);
    });

  const wb     = XLSX.readFile(EXCEL_IN);
  const sheet  = wb.Sheets[wb.SheetNames[0]];
  const rawHdr = XLSX.utils.sheet_to_json(sheet, { header:1 })[0];
  const data   = XLSX.utils.sheet_to_json(sheet, { defval:'', raw:true });

  const getPct5 = (a,b) => (a>0 && b>0) ? round5(Math.abs((b-a)/a*100)) : '';

  data.forEach(row => {
    const bars    = candlesMap[row.Instrument] || [];
    const btcBars = candlesMap['BTCUSDT'] || [];

    let Entry = '', EntryTime = '', ClosePrice = '', baseIdx = -1, triggerIdx = -1, EntryBar = '';
    if (bars.length) {
      const sig = new Date(row.Time); sig.setSeconds(0,0);
      const iso0 = sig.toISOString().slice(0,19)+'Z';
      const p6 = parseFloat(row['Pivot\u00a06']||row['Pivot 6']||0);
      const p5 = parseFloat(row['Pivot\u00a05']||row['Pivot 5']||0);
      const thr = p6||p5;
      baseIdx = bars.findIndex(b=>b.iso===iso0);
      if (thr>0 && baseIdx>=0) {
        for (let i=0;i<3;i++){
          const b=bars[baseIdx+i]; if(!b)break;
          const hit = (row['Point Type']==='LONG'  && b.close>thr)
                   || (row['Point Type']==='SHORT' && b.close<thr);
          if(hit){
            Entry= 'YES';
            EntryBar = i + 1;
            EntryTime= b.iso;
            ClosePrice= b.close;
            triggerIdx= baseIdx+i;
            break;
          }
        }
        if(!Entry) Entry='NO';
      } else Entry='NO';
    } else Entry='NO';

    const p1 = parseFloat(row['Pivot\u00a01']||row['Pivot 1']||0);
    const main6 = parseFloat(row['Pivot\u00a06']||row['Pivot 6']||row['Pivot\u00a05']||row['Pivot 5']||0);
    const p4 = parseFloat(row['Pivot\u00a04']||row['Pivot 4']||0);
    const PivotDiffPct = getPct5(p1, main6);
    const Diff4_6Pct   = getPct5(p4, main6);

    let EntryPrice='', StopPrice='', ProfitPct='', ProfitPrice='', ActualProfitPct='';
    if(Entry==='YES'&&triggerIdx>=0){
      const nb = bars[triggerIdx+1];
      if(nb){
        EntryPrice = nb.open;
        const stopPct   = round5(Diff4_6Pct * 1);
        const profitPct = round5(Diff4_6Pct * 2);
        ProfitPct = profitPct;

        if(row['Point Type']==='LONG'){
          StopPrice   = round5(EntryPrice * (1 - stopPct/100));
          ProfitPrice = round5(EntryPrice * (1 + profitPct/100));
        } else {
          StopPrice   = round5(EntryPrice * (1 + stopPct/100));
          ProfitPrice = round5(EntryPrice * (1 - profitPct/100));
        }

        let exit=null;
        for(let i=triggerIdx+1;i<bars.length;i++){
          const b=bars[i];
          if(row['Point Type']==='LONG'){
            if(b.high>=ProfitPrice){exit=ProfitPrice;break;}
            if(b.low <=StopPrice)  {exit=StopPrice;  break;}
          } else {
            if(b.low <=ProfitPrice){exit=ProfitPrice;break;}
            if(b.high>=StopPrice)  {exit=StopPrice;  break;}
          }
        }
        if(exit===null) exit=bars[bars.length-1].close;
        ActualProfitPct = round5(
          ((row['Point Type']==='LONG'
             ? (exit-EntryPrice)/EntryPrice
             : (EntryPrice-exit)/EntryPrice
           )*100)
        );
      }
    }

    let PrevExtremum='';
    if(baseIdx>0){
      const pr=bars[baseIdx-1];
      PrevExtremum = row['Point Type']==='LONG'? pr.low: pr.high;
    }

    let SignalVolRatio='', EntryVolRatio='';
    if(baseIdx>0){
      const now=bars[baseIdx].volume, pr=bars[baseIdx-1].volume;
      SignalVolRatio = pr>0? round5(now/pr): '';
    }
    if(triggerIdx>0){
      const now=bars[triggerIdx].volume, pr=bars[triggerIdx-1].volume;
      EntryVolRatio  = pr>0? round5(now/pr): '';
    }

    let BTC1mPct='', BTC5mPct='', BTC10mPct='', BTC15mPct='', BTC30mPct='';
    if(btcBars.length){
      const sig = new Date(row.Time); sig.setSeconds(0,0);
      const iso0 = sig.toISOString().slice(0,19)+'Z';
      const idx = btcBars.findIndex(b=>b.iso===iso0);
      if(idx>=0){
        const nowClose = btcBars[idx].close;
        const calc = (now, prev) => (prev>0? round5((now-prev)/prev*100): '');
        BTC1mPct  = idx>=1  ? calc(nowClose, btcBars[idx-1].close)  : '';
        BTC5mPct  = idx>=5  ? calc(nowClose, btcBars[idx-5].close)  : '';
        BTC10mPct = idx>=10 ? calc(nowClose, btcBars[idx-10].close) : '';
        BTC15mPct = idx>=15 ? calc(nowClose, btcBars[idx-15].close) : '';
        BTC30mPct = idx>=30 ? calc(nowClose, btcBars[idx-30].close) : '';
      }
    }

    Object.assign(row, {
      Entry, EntryTime, ClosePrice, EntryBar,
      PivotDiffPct, Diff4_6Pct,
      PrevExtremum, SignalVolRatio, EntryVolRatio,
      EntryPrice, StopPrice, ProfitPct, ProfitPrice, ActualProfitPct,
      BTC1mPct, BTC5mPct, BTC10mPct, BTC15mPct, BTC30mPct
    });      
  });

  const newHdr = [
    ...rawHdr,
    'Entry','EntryTime','EntryBar','ClosePrice',
    'PivotDiffPct','Diff4_6Pct','PrevExtremum',
    'SignalVolRatio','EntryVolRatio',
    'EntryPrice','StopPrice','ProfitPct','ProfitPrice','ActualProfitPct',
    '=(G2-F2)/(F2-E2)*-1','=(H2-G2)/(G2-F2)*-1','=\u0418(J2=""; K2<0,0001; L2>=0,8;L2<=1,999; M2>=0,6; M2<=1,999; U2>1,2; X2>=0,6; X2<5; O2>200)',
    'BTC1mPct','BTC5mPct','BTC10mPct','BTC15mPct','BTC30mPct'
  ];

  const out = XLSX.utils.json_to_sheet(data, { header:newHdr, skipHeader:false });
  wb.Sheets[wb.SheetNames[0]] = out;
  XLSX.writeFile(wb, EXCEL_OUT);
  console.log(`\u2714 ${EXCEL_OUT} \u0433\u043e\u0442\u043e\u0432`);
})();
