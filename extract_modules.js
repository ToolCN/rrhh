const fs = require('fs');

const content = fs.readFileSync('MainApp_RecursosHumanosHTML.html', 'utf8');
const lines = content.split('\n');

const modulos = [
  { nombre: 'Mod_DashboardRH',         inicio: 867,  fin: 1062 },
  { nombre: 'Mod_Plantilla',           inicio: 1063, fin: 1581 },
  { nombre: 'Mod_Bajas',               inicio: 1582, fin: 1628 },
  { nombre: 'Mod_Vacantes',            inicio: 1629, fin: 2030 },
  { nombre: 'Mod_Vacaciones',          inicio: 2031, fin: 2296 },
  { nombre: 'Mod_Autorizar',           inicio: 2297, fin: 2396 },
  { nombre: 'Mod_Aplicar',             inicio: 2397, fin: 2492 },
  { nombre: 'Mod_Consulta',            inicio: 2493, fin: 2847 },
  { nombre: 'Mod_MiUsuario',           inicio: 2848, fin: 2878 },
  { nombre: 'Mod_Usuarios',            inicio: 2879, fin: 3060 },
  { nombre: 'Mod_Actas',               inicio: 3061, fin: 3396 },
  { nombre: 'Mod_ConsultaActas',       inicio: 3397, fin: 3862 },
  { nombre: 'Mod_Ausentismo',          inicio: 3863, fin: 4002 },
  { nombre: 'Mod_ConsultaAusentismo',  inicio: 4003, fin: 4931 },
  { nombre: 'Mod_Retardos',            inicio: 4932, fin: 5036 },
  { nombre: 'Mod_ConsultaRetardos',    inicio: 5037, fin: 5198 },
  { nombre: 'Mod_PermisosEmp',         inicio: 5199, fin: 5374 },
  { nombre: 'Mod_ConsultaPermisos',    inicio: 5375, fin: 5869 },
  { nombre: 'Mod_NominaSemanal',       inicio: 5870, fin: 5870 },
  { nombre: 'Mod_NominaQuincenal',     inicio: 5871, fin: 6314 },
  { nombre: 'Mod_TiempoExtra',         inicio: 6315, fin: 6620 },
  { nombre: 'Mod_ModuloProximo',       inicio: 6621, fin: 6641 },
  { nombre: 'Mod_Tabulador',           inicio: 6642, fin: 7121 },
  { nombre: 'Mod_Capacitacion',        inicio: 7122, fin: 8134 },
  { nombre: 'Mod_PerfilesPuestos',     inicio: 8135, fin: 8238 },
  { nombre: 'Mod_AjustesCat',          inicio: 8239, fin: lines.length },
];

modulos.forEach(mod => {
  const contenido = lines.slice(mod.inicio - 1, mod.fin).join('\n');
  fs.writeFileSync(mod.nombre + '.html', contenido, 'utf8');
  console.log(`✓ ${mod.nombre}.html — ${mod.fin - mod.inicio + 1} líneas`);
});

console.log('\nMódulos creados: ' + modulos.length);
