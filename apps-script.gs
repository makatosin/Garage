/**
 * Garage · collegamento al Foglio Google
 *
 * Come si installa (una volta sola):
 * 1. Nel foglio: Estensioni › Apps Script. Cancella tutto e incolla questo file. Salva.
 * 2. In alto scegli la funzione "inizia" e premi Esegui. Autorizza (è il tuo account).
 *    Crea il foglio "Sistema" con il codice di accesso.
 * 3. Distribuisci › Nuova distribuzione › tipo "App web".
 *    Esegui come: Me. Chi ha accesso: Chiunque. Distribuisci e copia l'URL (finisce con /exec).
 * 4. Nell'app Garage: Altro › Foglio Google › Collega. Incolla URL e codice.
 *
 * I fogli "Rifornimenti", "Manutenzione" e "Veicoli" sono solo da leggere:
 * l'app li riscrive a ogni salvataggio. I dati veri stanno nel foglio "Dati".
 */

const PEZZO = 45000; // caratteri per cella (il limite di Google è 50.000)

function inizia(){
  const ss = SpreadsheetApp.getActive();
  codice_(ss);
  foglio_(ss, 'Dati');
  ss.setActiveSheet(ss.getSheetByName('Sistema'));
  Logger.log('Pronto. Il codice di accesso è nel foglio "Sistema".');
}

function doGet(){ return out_({ok:true, app:'garage'}); }

function doPost(e){
  let req;
  try { req = JSON.parse(e.postData.contents); } catch (x){ return out_({ok:false, error:'richiesta non valida'}); }
  const ss = SpreadsheetApp.getActive();
  if (!req.token || req.token !== codice_(ss)) return out_({ok:false, error:'codice di accesso sbagliato'});
  const lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try {
    const cur = leggi_(ss);
    if (req.action === 'load') return out_({ok:true, data:cur});
    if (req.action === 'save'){
      const p = req.payload || {};
      if (p.rev !== cur.rev) return out_({ok:false, conflict:true, error:'dati cambiati nel frattempo'});
      const rev = cur.rev + 1;
      scrivi_(ss, p.state, rev);
      try { tabelle_(ss, p.state); } catch (x){}
      return out_({ok:true, data:{rev:rev}});
    }
    return out_({ok:false, error:'azione sconosciuta'});
  } finally { lock.releaseLock(); }
}

function out_(o){ return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

function foglio_(ss, nome){ return ss.getSheetByName(nome) || ss.insertSheet(nome); }

function codice_(ss){
  const sh = foglio_(ss, 'Sistema');
  let c = String(sh.getRange('B1').getValue() || '').trim();
  if (!c){
    c = Utilities.getUuid().replace(/-/g, '').slice(0, 12);
    sh.getRange('A1:B1').setValues([['Codice di accesso', c]]);
    sh.getRange('A3').setValue('Serve per collegare l\'app Garage a questo foglio. Non condividerlo.');
    sh.setColumnWidth(1, 160); sh.setColumnWidth(2, 160);
  }
  return c;
}

function leggi_(ss){
  const sh = foglio_(ss, 'Dati');
  const n = sh.getLastRow();
  if (n < 2) return {rev:0, state:null};
  const vals = sh.getRange(1, 1, n, 1).getValues().map(r => String(r[0]));
  const rev = Number(vals[0]) || 0;
  const txt = vals.slice(1).join('');
  let state = null;
  try { state = txt ? JSON.parse(txt) : null; } catch (x){ state = null; }
  return {rev:rev, state:state};
}

function scrivi_(ss, state, rev){
  const sh = foglio_(ss, 'Dati');
  const txt = JSON.stringify(state || {});
  const righe = [[String(rev)]];
  for (let i = 0; i < txt.length; i += PEZZO) righe.push([txt.slice(i, i + PEZZO)]);
  sh.clear();
  sh.getRange(1, 1, righe.length, 1).setNumberFormat('@').setValues(righe);
}

/* ---------- fogli leggibili ---------- */
function tabella_(ss, nome, intest, righe){
  const sh = foglio_(ss, nome);
  sh.clear();
  const tutto = [intest].concat(righe);
  sh.getRange(1, 1, tutto.length, intest.length).setValues(tutto);
  sh.getRange(1, 1, 1, intest.length).setFontWeight('bold');
  sh.setFrozenRows(1);
}
function d_(iso){ if (!iso) return ''; const p = String(iso).split('-').map(Number); return new Date(p[0], p[1] - 1, p[2]); }
function n_(x){ return x == null || x === '' || !isFinite(x) ? '' : Number(x); }

function tabelle_(ss, s){
  s = s || {};
  const veic = s.veicoli || [];
  const nome = id => (veic.find(v => v.id === id) || {}).nome || '';
  tabella_(ss, 'Veicoli', ['Nome', 'Modello', 'Targa', 'Carburante', 'Ultima lettura km', 'Data lettura'],
    veic.map(v => [v.nome || '', v.modello || '', v.targa || '', v.carburante || '', v.lettura ? n_(v.lettura.km) : '', v.lettura ? d_(v.lettura.data) : '']));
  const rif = (s.rif || []).slice().sort((a, b) => String(b.data).localeCompare(String(a.data)));
  tabella_(ss, 'Rifornimenti', ['Data', 'Veicolo', 'Carburante', 'Km', 'Litri', '€/litro', 'Totale €', 'Pieno', 'Luogo', 'Nota'],
    rif.map(r => [d_(r.data), nome(r.v), r.carb || '', n_(r.km), n_(r.litri), n_(r.prezzo), n_(r.totale), r.pieno === false ? 'no' : 'sì', r.luogo || '', r.nota || '']));
  const sp = (s.spese || []).slice().sort((a, b) => String(b.data || '9999').localeCompare(String(a.data || '9999')));
  tabella_(ss, 'Manutenzione', ['Tipo', 'Veicolo', 'Cosa', 'Dettagli', 'Stato', 'Data', 'Km', 'Costo €', 'Scadenza data', 'Scadenza km', 'Ripeti ogni mesi', 'Ripeti ogni km', 'Nota'],
    sp.map(x => [x.tipo === 'documenti' ? 'Documento' : 'Manutenzione', nome(x.v), x.titolo || '', x.dettagli || '', x.fatto === false ? 'Da fare' : 'Fatto',
      d_(x.data), n_(x.km), n_(x.importo) || '', x.scad ? d_(x.scad.data) : '', x.scad ? n_(x.scad.km) : '', x.rip ? n_(x.rip.mesi) : '', x.rip ? n_(x.rip.km) : '', x.nota || '']));
}
