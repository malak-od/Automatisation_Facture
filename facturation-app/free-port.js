#!/usr/bin/env node
// ============================================================================
//  Libere le port de dev (4000 par defaut) avant de lancer le serveur : tue
//  toute instance node.exe deja en ecoute dessus (ex. serveur oublie d'une
//  session precedente), pour que l'app demarre TOUJOURS sur le meme port en
//  dev plutot que de glisser silencieusement sur 4001+ (cf. fallback dans
//  server.js, utile en prod mais deroutant en dev -- on peut se retrouver a
//  tester contre une vieille instance sans le savoir).
// ============================================================================
const { execSync } = require('child_process');

const PORT = Number(process.env.PORT) || 4000;

function freePortWindows(port) {
  let out;
  try {
    out = execSync(`netstat -ano -p TCP | findstr :${port}`, { encoding: 'utf8' });
  } catch (e) {
    return; // rien sur ce port
  }
  const pids = new Set();
  for (const line of out.split(/\r?\n/)) {
    const m = /LISTENING\s+(\d+)\s*$/.exec(line.trim());
    if (m) pids.add(m[1]);
  }
  for (const pid of pids) {
    try {
      execSync(`taskkill /PID ${pid} /F`, { stdio: 'ignore' });
      console.log(`Port ${port} libéré (arrêt du processus ${pid}).`);
    } catch (e) {
      console.warn(`Impossible d'arrêter le processus ${pid} sur le port ${port} :`, e.message);
    }
  }
}

if (process.platform === 'win32') {
  freePortWindows(PORT);
} else {
  try {
    const pids = execSync(`lsof -ti tcp:${PORT}`, { encoding: 'utf8' }).trim();
    if (pids) {
      execSync(`kill -9 ${pids.split(/\s+/).join(' ')}`);
      console.log(`Port ${PORT} libéré.`);
    }
  } catch (e) {
    // rien sur ce port
  }
}
