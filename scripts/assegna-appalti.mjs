// DR Portal — Assegnazione APPALTO ai dipendenti (una tantum, 08/2026).
// -----------------------------------------------------------------------------
// Scrive la colonna "Appalto" sulla lista Dipendenti passando dal gateway del
// portale (stesse credenziali di scripts/import-dipendenti.mjs). L'elenco
// nome→appalto è INCORPORATO qui sotto (fornito dall'ufficio il 07/08/2026).
//
// - I nomi si agganciano ai record esistenti a token, in ordine libero
//   (GUERMANDI GIANLUCA == GIANLUCA GUERMANDI); gli ambigui vengono SALTATI
//   e riportati a video.
// - I nomi SENZA scheda vengono CREATI con il minimo indispensabile
//   (Nome, Cognome, Appalto, Attivo=Sì, Visibile=Sì): niente Codice/PIN,
//   quindi non possono accedere finché l'ufficio non completa la scheda.
//
// USO
//   node scripts/assegna-appalti.mjs --dry-run    (prova: non scrive nulla)
//   node scripts/assegna-appalti.mjs              (scrive)
//
// CREDENZIALI (env, dai secret del progetto Lovable):
//   LOVABLE_API_KEY=...  MICROSOFT_SHAREPOINT_API_KEY=...
// PREREQUISITO: colonna "Appalto" (testo) creata sulla lista Dipendenti.
// -----------------------------------------------------------------------------

const GATEWAY = "https://connector-gateway.lovable.dev/microsoft_sharepoint";
const HOST = "drlogisticaroma.sharepoint.com";
const SITE_PATH = "DRPORTAL";

const LOVABLE_KEY = process.env.LOVABLE_API_KEY;
const SP_KEY = process.env.MICROSOFT_SHAREPOINT_API_KEY;
const dryRun = process.argv.includes("--dry-run");

// nome completo (come fornito) → appalto
const DATI = [
  ["GIANLUCA GUERMANDI", "CEVA"],
  ["BRUNO VECCHIO", "CEVA"],
  ["MARCO MOLINARI SIGHINOLFI", "CEVA"],
  ["RICCARDO CARLONE", "IMILE"],
  ["VASILE CRISTINEL BUCA", "POSTA DOC DISTRIBUZIONE COLLEFERRO"],
  ["GIANLUCA CARATELLI", "POSTA DOC DISTRIBUZIONE COLLEFERRO"],
  ["SARA CELLITTI", "POSTA DOC DISTRIBUZIONE COLLEFERRO"],
  ["LORENZO D'ALESSANDRO", "POSTA DOC DISTRIBUZIONE COLLEFERRO"],
  ["FEDERICA DE SANTIS", "POSTA DOC DISTRIBUZIONE COLLEFERRO"],
  ["FEDERICO DI FATTA", "POSTA DOC DISTRIBUZIONE COLLEFERRO"],
  ["ALESSIA DI PINTO", "POSTA DOC DISTRIBUZIONE COLLEFERRO"],
  ["LUANA GIULIANI", "POSTA DOC DISTRIBUZIONE COLLEFERRO"],
  ["LAURA LAURETI", "POSTA DOC DISTRIBUZIONE COLLEFERRO"],
  ["FRANCESCO LUNEDI", "POSTA DOC DISTRIBUZIONE COLLEFERRO"],
  ["ANDREA MARI", "POSTA DOC DISTRIBUZIONE COLLEFERRO"],
  ["ALESSIA MORTARI", "POSTA DOC DISTRIBUZIONE COLLEFERRO"],
  ["MAURIZIO PELLEGRINI", "POSTA DOC DISTRIBUZIONE COLLEFERRO"],
  ["MASSIMILIANO PIERRO", "POSTA DOC DISTRIBUZIONE COLLEFERRO"],
  ["GIACOMO MAGGI", "POSTA DOC DISTRIBUZIONE FIANO ROMANO"],
  ["FABIO BARBA", "POSTA DOC DISTRIBUZIONE FIANO ROMANO"],
  ["KLODIAN PASHKAJ", "POSTA DOC DISTRIBUZIONE FIANO ROMANO"],
  ["JESSICA GORGAJ", "POSTA DOC MAGAZZINO FIANO ROMANO"],
  ["LUCA SIMEONI", "POSTA DOC MAGAZZINO FIANO ROMANO"],
  ["VALERIO TRIONFERA", "POSTA DOC MAGAZZINO FIANO ROMANO"],
  ["ANGELO PESCARA", "POSTA DOC MAGAZZINO FIANO ROMANO"],
  ["ANNA DONATIELLO", "POSTA DOC MAGAZZINO FIANO ROMANO"],
  ["DAVIDE URBANO", "POSTA DOC MAGAZZINO FIANO ROMANO"],
  ["ROMINA SIMEI", "POSTA DOC MAGAZZINO FIANO ROMANO"],
  ["KATRINA MATIJA", "POSTA DOC MAGAZZINO FIANO ROMANO"],
  ["SALI PASHKAJ", "POSTA DOC MAGAZZINO FIANO ROMANO"],
  ["PAVLINA NDOJA", "POSTA DOC MAGAZZINO FIANO ROMANO"],
  ["ANNA MARIA TEMPERINI", "POSTA DOC MAGAZZINO FIANO ROMANO"],
  ["MASSIMO FIORONI", "POSTA DOC MAGAZZINO FIANO ROMANO"],
  ["MATILDE DANIELI", "POSTA DOC MAGAZZINO FIANO ROMANO"],
  ["NICOLAE ANDREI NEBUNESCU", "POSTA DOC MAGAZZINO FIANO ROMANO"],
  ["ARDIANA MARKU", "POSTA DOC MAGAZZINO FIANO ROMANO"],
  ["TAMARA PACIONI", "POSTA DOC MAGAZZINO FIANO ROMANO"],
  ["RIGERTO MALSHI", "POSTA DOC MAGAZZINO FIANO ROMANO"],
  ["ALMA NDOJA", "POSTA DOC MAGAZZINO FIANO ROMANO"],
  ["KATIA STRAZZIERI", "POSTA DOC MAGAZZINO FIANO ROMANO"],
  ["FEDERICA EVANGELISTI", "POSTA DOC MAGAZZINO FIANO ROMANO"],
  ["YOUSRA JARMOUNI", "POSTA DOC MAGAZZINO FIANO ROMANO"],
  ["DIEGO SCARLATO", "POSTA DOC MAGAZZINO FIANO ROMANO"],
  ["LORENZO MEI", "POSTA DOC MAGAZZINO FIANO ROMANO"],
  ["MICHELE MADONIA", "POSTA DOC MAGAZZINO FIANO ROMANO"],
  ["VERA VERGINELLI", "POSTA DOC MAGAZZINO FIANO ROMANO"],
  ["ALESSANDRO IBRAHIMI", "POSTA DOC MAGAZZINO FIANO ROMANO"],
  ["ALESSIA SERPIETRI", "POSTA DOC MAGAZZINO FIANO ROMANO"],
  ["SABRINA GUIDARELLI", "UFFICIO FIANO ROMANO"],
  ["TAZJANA ZEKAJ", "UFFICIO FIANO ROMANO"],
  ["SIMONE RUSSO", "UFFICIO FIANO ROMANO"],
  ["LUCREZIA PRATESI", "UFFICIO FIANO ROMANO"],
  ["SARA MONTI", "UFFICIO SAN GIULIANO MILANESE"],
  ["ELISABETTA NOTARO", "UFFICIO SAN GIULIANO MILANESE"],
  ["GIULIA SPERA", "UFFICIO SAN GIULIANO MILANESE"],
  ["JOYSIE AQUINO", "UFFICIO SAN GIULIANO MILANESE"],
  ["DIEGO GABELLI SERVENTI", "UFFICIO SAN GIULIANO MILANESE"],
  ["MARCO MARELLI", "UFFICIO SAN GIULIANO MILANESE"],
  ["ARMANDO BLADIMIR GONZALEZ DEPAZ", "UNIVEX MILANO"],
  ["SEGUNDO JAVIER CARDENAS ONTANEDA", "UNIVEX MILANO"],
  ["CRISTIAN MESSI", "UNIVEX MILANO"],
  ["JORGE ROMAN VELARDES", "UNIVEX MILANO"],
  ["CHANDANA KUMARA WIJESINGHE", "UNIVEX MILANO"],
  ["AMRO AHMED MOHAMED KAF", "UNIVEX MILANO"],
  ["CHRISTIAN ANDRES VARGAS VALDEZ", "UNIVEX MILANO"],
  ["ORONZO CANDELA", "UNIVEX MILANO"],
  ["GIANLUCA MARTEL MINAYA", "UNIVEX MILANO"],
  ["HENRY PAUL ALVAREZ EQUILLAS", "UNIVEX MILANO"],
  ["CARMELO NIGRELLI", "UNIVEX MILANO"],
  ["MANUEL PAUL ALFARO CRISOLOGO", "UNIVEX MILANO"],
  ["ABDELHALIM MAHMOUD ABDELHADI MOHAMED ZINA", "UNIVEX MILANO"],
  ["RUGGIERO RIGLIETTI", "UNIVEX MILANO"],
  ["FREDDY LANDER ARTEAGA CORREA", "UNIVEX MILANO"],
  ["MAURO AGOLINI", "UNIVEX MILANO"],
  ["ADELIO MARTINOTTI", "Number 1"],
  ["ALFREDO ALEJANDRO ISAAC MARTINEZ MEZA", "Number 1"],
  ["EDISON MARCELO CHICO", "Number 1"],
  ["GIORGIO LAMBERTO", "UNIVEX SAVONA"],
  ["MAURIZIO BRASCA", "UNIVEX SAVONA"],
  ["CHRISTIAN CANNIZZARO", "UNIVEX SAVONA"],
  ["GIOVANNI GALLINA", "UNIVEX SAVONA"],
  ["FILIPPO CAMBIASO", "UNIVEX SAVONA"],
  ["ERMANNO OLIVIERI", "UNIVEX SAVONA"],
  ["GIANVITO MARIANO", "UNIVEX SAVONA"],
  ["PIETRO DRAGO", "UNIVEX SAVONA"],
  ["MOUSSA KONDE", "UNIVEX SAVONA"],
  ["DOMENICO CASOLA", "UNIVEX SAVONA"],
  ["DAVIDE MARSILI", "UNIVEX SAVONA"],
  ["ALEX MOSCATO", "UNIVEX SAVONA"],
  ["LUCA OTTONELLO", "UNIVEX SAVONA"],
  ["MARTA BARONI", "UNIVEX SAVONA"],
  ["ANTONIO MARINO", "UNIVEX SAVONA"],
  ["GIANFRANCO BOBOCCA", "UNIVEX SAVONA"],
  ["STEFANO RIBOLA", "UNIVEX SAVONA"],
  ["MARCO GARZOGLIO", "UNIVEX SAVONA"],
  ["LORIS LIAM OTTONELLO", "UNIVEX SAVONA"],
  ["FAUSTO PARODI", "UNIVEX SAVONA"],
  ["AZEDDINE BASTI", "UNIVEX SAVONA"],
  ["LORIS ACCA", "ZINGALI"],
  ["GEREMIA ALTROCCHI", "ZINGALI"],
  ["MAURIZIO ALDO BALLARINO", "ZINGALI"],
  ["ELIAS PATRICIO JARAMILLO RUCANO", "ZINGALI"],
  ["EMAD RASHAD MOHAMED KHAFAGY", "ZINGALI"],
  ["PAUL CHENGO MWAGANDI", "ZINGALI"],
  ["MASSIMILIANO PARMA", "ZINGALI"],
  ["MATTEO VITULANO", "ZINGALI"],
  ["SIMONE AUSILIO", "ZINGALI"],
  ["AZZEDDINE MOUSLIM", "ZINGALI"],
  ["EFE JOSEPHINE IBHAFIDON", "ZINGALI"],
  ["FRANCESCO MARIA BELLINI", "ZINGALI"],
  ["EGLINA SPEDIACCI", "ZINGALI"],
  ["MAURO GAGGIANESI", "ZINGALI"],
  ["GIOVANNI BENANTI", "ZINGALI"],
  ["ANTONELLO DI MURO", "ZINGALI"],
];

function fail(msg) {
  console.error(`\n❌ ${msg}\n`);
  process.exit(1);
}
if (!LOVABLE_KEY || !SP_KEY) fail("Servono le env LOVABLE_API_KEY e MICROSOFT_SHAREPOINT_API_KEY.");

async function gj(path, init = {}) {
  const headers = {
    Authorization: `Bearer ${LOVABLE_KEY}`,
    "X-Connection-Api-Key": SP_KEY,
    ...(init.body ? { "Content-Type": "application/json" } : {}),
  };
  const res = await fetch(`${GATEWAY}${path}`, { ...init, headers });
  if (!res.ok) {
    const body = await res.text().catch(() => "");
    throw new Error(
      `${init.method ?? "GET"} ${path.split("?")[0]} → ${res.status} ${body.slice(0, 300)}`,
    );
  }
  return res.json();
}

// Normalizzazione a token: minuscole, niente accenti/punteggiatura, ordine libero.
const tokens = (s) =>
  (s ?? "")
    .normalize("NFD")
    .replace(/[̀-ͯ]/g, "")
    .toLowerCase()
    .replace(/[^a-z0-9 ]/g, " ")
    .split(/\s+/)
    .filter((x) => x.length >= 2)
    .sort();
const chiave = (s) => tokens(s).join(" ");

async function main() {
  console.log(`\n📋 Appalti da assegnare: ${DATI.length}${dryRun ? "  (DRY-RUN)" : ""}`);
  const site = await gj(`/sites/${HOST}:/sites/${SITE_PATH}`);
  const listsRes = await gj(`/sites/${site.id}/lists?$select=id,name,displayName`);
  const dip = (listsRes.value ?? []).find((l) => /dipendent/i.test(l.displayName || l.name || ""));
  if (!dip) fail('Lista "Dipendenti" non trovata.');
  const colsRes = await gj(
    `/sites/${site.id}/lists/${dip.id}/columns?$select=name,displayName,readOnly,hidden`,
  );
  const byDisplay = new Map();
  for (const c of colsRes.value ?? []) {
    if (c.hidden || c.readOnly) continue;
    if (c.displayName) byDisplay.set(c.displayName.toLowerCase(), c.name);
  }
  const COL = {
    nome: byDisplay.get("nome"),
    cognome: byDisplay.get("cognome"),
    appalto: byDisplay.get("appalto"),
    attivo: byDisplay.get("attivo"),
    visibile: byDisplay.get("visibile"),
  };
  if (!COL.appalto) fail('Colonna "Appalto" assente sulla lista Dipendenti: crearla (testo) prima.');
  if (!COL.nome || !COL.cognome) fail("Colonne Nome/Cognome non trovate.");

  // Tutti i dipendenti esistenti (paginato).
  const esistenti = [];
  let path = `/sites/${site.id}/lists/${dip.id}/items?expand=fields&$top=999`;
  while (path) {
    const res = await gj(path);
    esistenti.push(...(res.value ?? []));
    const next = res["@odata.nextLink"];
    path = next ? new URL(next).pathname.replace(/^\/(?:v1\.0|beta)/, "") + new URL(next).search : null;
  }
  console.log(`👥 Dipendenti in lista: ${esistenti.length}`);

  const perChiave = new Map();
  for (const it of esistenti) {
    const f = it.fields ?? {};
    const nomeC = `${f[COL.nome] ?? ""} ${f[COL.cognome] ?? ""}`;
    const k = chiave(nomeC);
    if (!k) continue;
    if (perChiave.has(k)) perChiave.set(k, "AMBIGUO");
    else perChiave.set(k, it);
  }

  let aggiornati = 0;
  let invariati = 0;
  let creati = 0;
  const ambigui = [];
  for (const [nomeCompleto, appalto] of DATI) {
    const k = chiave(nomeCompleto);
    const hit = perChiave.get(k);
    if (hit === "AMBIGUO") {
      ambigui.push(nomeCompleto);
      continue;
    }
    if (hit) {
      const attuale = String((hit.fields ?? {})[COL.appalto] ?? "").trim();
      if (attuale === appalto) {
        invariati++;
        continue;
      }
      console.log(`  ✏️  ${nomeCompleto}  →  ${appalto}${attuale ? `  (era: ${attuale})` : ""}`);
      if (!dryRun)
        await gj(`/sites/${site.id}/lists/${dip.id}/items/${hit.id}/fields`, {
          method: "PATCH",
          body: JSON.stringify({ [COL.appalto]: appalto }),
        });
      aggiornati++;
    } else {
      // Scheda MINIMA: nome = primo token, cognome = resto. Niente
      // Codice/PIN: la persona esiste per i salari, non può accedere.
      const parti = nomeCompleto.trim().split(/\s+/);
      const nome = parti[0];
      const cognome = parti.slice(1).join(" ") || parti[0];
      console.log(`  ➕ NUOVA scheda: ${nome} / ${cognome}  →  ${appalto}`);
      if (!dryRun) {
        const fields = {
          [COL.nome]: nome,
          [COL.cognome]: cognome,
          [COL.appalto]: appalto,
        };
        if (COL.attivo) fields[COL.attivo] = true;
        if (COL.visibile) fields[COL.visibile] = true;
        await gj(`/sites/${site.id}/lists/${dip.id}/items`, {
          method: "POST",
          body: JSON.stringify({ fields }),
        });
      }
      creati++;
    }
  }

  console.log(`\n✅ Fatto${dryRun ? " (nessuna scrittura: dry-run)" : ""}.`);
  console.log(`   aggiornati: ${aggiornati} · nuove schede: ${creati} · già a posto: ${invariati}`);
  if (ambigui.length) {
    console.log(`   ⚠️ ambigui (saltati, da fare a mano): ${ambigui.join(", ")}`);
  }
}

main().catch((e) => fail(e.message));
