const fs = require('fs');

const MAIN_PATH = 'MainApp_RecursosHumanosHTML.html';
const MARKER = '// ════════════════════════ MÓDULOS — rebuild_main.js ════════════════════════ //';

const ORDEN = [
  'Mod_DashboardRH',
  'Mod_Plantilla',
  'Mod_Bajas',
  'Mod_Vacantes',
  'Mod_Vacaciones',
  'Mod_Autorizar',
  'Mod_Aplicar',
  'Mod_Consulta',
  'Mod_MiUsuario',
  'Mod_Usuarios',
  'Mod_Actas',
  'Mod_ConsultaActas',
  'Mod_Ausentismo',
  'Mod_ConsultaAusentismo',
  'Mod_Retardos',
  'Mod_ConsultaRetardos',
  'Mod_PermisosEmp',
  'Mod_ConsultaPermisos',
  'Mod_NominaSemanal',
  'Mod_NominaQuincenal',
  'Mod_TiempoExtra',
  'Mod_ModuloProximo',
  'Mod_Tabulador',
  'Mod_Capacitacion',
  'Mod_PerfilesPuestos',
  'Mod_AjustesCat',
];

const main = fs.readFileSync(MAIN_PATH, 'utf8');
const markerIdx = main.indexOf(MARKER);
if (markerIdx === -1) { console.error('ERROR: Marcador no encontrado.'); process.exit(1); }

const closeTag = '\n</script>';
const closeIdx = main.indexOf(closeTag, markerIdx);
if (closeIdx === -1) { console.error('ERROR: No se encontró </script> después del marcador.'); process.exit(1); }

const before = main.slice(0, markerIdx);
const after = main.slice(closeIdx);

let jsUnido = '';
for (const mod of ORDEN) {
  const filePath = mod + '.html';
  if (!fs.existsSync(filePath)) { console.warn(`  ⚠ No encontrado: ${filePath}`); continue; }
  let contenido = fs.readFileSync(filePath, 'utf8');
  if (!contenido.endsWith('\n')) contenido += '\n';
  jsUnido += contenido;
  console.log(`  + ${filePath} (${contenido.split('\n').length - 1} líneas)`);
}

const result = before + MARKER + '\n' + jsUnido + after;
fs.writeFileSync(MAIN_PATH, result, 'utf8');
console.log(`\nMainApp reconstruido: ${result.split('\n').length} líneas.`);
