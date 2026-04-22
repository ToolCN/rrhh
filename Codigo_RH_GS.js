// ═══════════════════════════════════════════════════════════════════
//  CODIGO_RH_GS — Google Apps Script SEPARADO para Recursos Humanos (NUEVO) OK bien Prueba
//  Vinculado a la misma hoja de cálculo del GS original
// ═══════════════════════════════════════════════════════════════════

var ID_HOJA_CALCULO = "1RKi09zpQ3KMa_JLUINYJysDOFRi3tM2M2a8JW8Qy7gk";
var ID_HOJA_RH     = "1maI-y2eV6ZxQoV4frYT6eRkynstQErVbYWvmHz7sDUk"; // Libro RecursosHumanos

// ─── PUNTO DE ENTRADA ────────────────────────────────────────────────────────────
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('MainApp_RecursosHumanosHTML')
    .setTitle('Recursos Humanos')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// =================================================================================
// ══════════════════  GESTIÓN DE USUARIOS  ════════════════════════════════════════
// =================================================================================
// Orden de columnas en hoja USUARIOS:
//   A=NOMBRE | B=PASSWORD | C=ROL | D=FOTO | E=AREA | F=PERMISOS
// La función _getIdxUsuarios busca por nombre Y tiene fallback por posición,
// así el código funciona aunque el encabezado tenga acento o espacios.
// =================================================================================

function _getIdxUsuarios(headers) {
  var find = function(nombres, fallback) {
    for (var n = 0; n < nombres.length; n++) {
      var idx = headers.indexOf(nombres[n]);
      if (idx > -1) return idx;
    }
    return fallback;
  };
  return {
    NOMBRE:   find(["NOMBRE","NAME"],                    0),
    PASSWORD: find(["PASSWORD","PASS","CONTRASEÑA"],     1),
    ROL:      find(["ROL","ROLE"],                       2),
    FOTO:     find(["FOTO","PHOTO","IMAGEN"],             3),
    AREA:     find(["AREA","ÁREA","DEPARTAMENTO"],        4),
    PERMISOS: find(["PERMISOS","PERMISO","PERMISOS_RH"],  5)  // columna F
  };
}

function getUsuariosLogin() {
  var ss    = SpreadsheetApp.openById(ID_HOJA_CALCULO);
  var sheet = ss.getSheetByName("USUARIOS");
  if (!sheet) return [];
  var data    = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  var headers = data[0].map(function(h){ return String(h).toUpperCase().trim(); });
  var IDX     = _getIdxUsuarios(headers);
  var lista   = [];
  for (var i = 1; i < data.length; i++) {
    var nombre = String(data[i][IDX.NOMBRE] || "").trim();
    if (!nombre) continue;
    lista.push({
      nombre:   nombre,
      password: String(data[i][IDX.PASSWORD] || ""),
      permisos: String(data[i][IDX.PERMISOS] || "")
    });
  }
  return lista;
}

function getUsuariosAdmin() {
  var ss    = SpreadsheetApp.openById(ID_HOJA_CALCULO);
  var sheet = ss.getSheetByName("USUARIOS");
  if (!sheet) return [];
  var data    = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  var headers = data[0].map(function(h){ return String(h).toUpperCase().trim(); });
  var IDX     = _getIdxUsuarios(headers);
  var lista   = [];
  for (var i = 1; i < data.length; i++) {
    var nombre = String(data[i][IDX.NOMBRE] || "").trim();
    if (!nombre) continue;
    lista.push({
      nombre:   nombre,
      password: String(data[i][IDX.PASSWORD] || ""),
      permisos: String(data[i][IDX.PERMISOS] || ""),
      rol:      String(data[i][IDX.ROL]      || ""),
      foto:     String(data[i][IDX.FOTO]     || "")
    });
  }
  return lista;
}

function guardarUsuario(datos) {
  var ss    = SpreadsheetApp.openById(ID_HOJA_CALCULO);
  var sheet = ss.getSheetByName("USUARIOS");
  if (!sheet) return "Error: No existe la hoja USUARIOS";
  var data    = sheet.getDataRange().getValues();
  var headers = data[0].map(function(h){ return String(h).toUpperCase().trim(); });
  var IDX     = _getIdxUsuarios(headers);
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][IDX.NOMBRE]).toUpperCase().trim() === datos.nombre.toUpperCase().trim()) {
      return "Ya existe un usuario con ese nombre.";
    }
  }
  var numCols = headers.length;
  var fila    = [];
  for (var c = 0; c < numCols; c++) fila.push("");
  fila[IDX.NOMBRE]   = datos.nombre.toUpperCase().trim();
  fila[IDX.PASSWORD] = datos.password || "";
  fila[IDX.PERMISOS] = datos.permisos || "";
  fila[IDX.ROL]      = datos.rol      || "";
  fila[IDX.FOTO]     = datos.foto     || "";
  sheet.appendRow(fila);
  return "OK";
}

function actualizarUsuario(datos) {
  var ss    = SpreadsheetApp.openById(ID_HOJA_CALCULO);
  var sheet = ss.getSheetByName("USUARIOS");
  if (!sheet) return "Error: No existe la hoja USUARIOS";
  var data    = sheet.getDataRange().getValues();
  var headers = data[0].map(function(h){ return String(h).toUpperCase().trim(); });
  var IDX     = _getIdxUsuarios(headers);
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][IDX.NOMBRE]).toUpperCase().trim() === datos.nombre.toUpperCase().trim()) {
      sheet.getRange(i + 1, IDX.PASSWORD + 1).setValue(datos.password || "");
      sheet.getRange(i + 1, IDX.PERMISOS + 1).setValue(datos.permisos || "");
      return "OK";
    }
  }
  return "Usuario no encontrado.";
}

function eliminarUsuario(nombre) {
  var ss    = SpreadsheetApp.openById(ID_HOJA_CALCULO);
  var sheet = ss.getSheetByName("USUARIOS");
  if (!sheet) return "Error: No existe la hoja USUARIOS";
  var data    = sheet.getDataRange().getValues();
  var headers = data[0].map(function(h){ return String(h).toUpperCase().trim(); });
  var IDX     = _getIdxUsuarios(headers);
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][IDX.NOMBRE]).toUpperCase().trim() === nombre.toUpperCase().trim()) {
      sheet.deleteRow(i + 1);
      return "OK";
    }
  }
  return "Usuario no encontrado.";
}

function actualizarPasswordUsuario(nombre, nuevaPassword) {
  var ss    = SpreadsheetApp.openById(ID_HOJA_CALCULO);
  var sheet = ss.getSheetByName("USUARIOS");
  if (!sheet) return "Error: No existe la hoja USUARIOS";
  var data    = sheet.getDataRange().getValues();
  var headers = data[0].map(function(h){ return String(h).toUpperCase().trim(); });
  var IDX     = _getIdxUsuarios(headers);
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][IDX.NOMBRE]).toUpperCase().trim() === nombre.toUpperCase().trim()) {
      sheet.getRange(i + 1, IDX.PASSWORD + 1).setValue(nuevaPassword);
      return "OK";
    }
  }
  return "Usuario no encontrado.";
}

// =================================================================================
// ══════════════════  DASHBOARD RH  ═══════════════════════════════════════════════
// =================================================================================

function obtenerDashboardRH(mes, anio) {
  var ss       = SpreadsheetApp.openById(ID_HOJA_CALCULO);
  var sheetOps = ss.getSheetByName("OPERADORES");
  var sheetSols= ss.getSheetByName("SOLICITUDES");
  var hoy      = new Date();
  var mesDash  = (mes  !== undefined && mes  !== null) ? Number(mes)  : hoy.getMonth();
  var anioDash = (anio !== undefined && anio !== null) ? Number(anio) : hoy.getFullYear();

  // ── OPERADORES ──────────────────────────────────────────────────────────────
  var totalActivos = 0, totalVacantes = 0;
  var porAreaActivo = {}, porAreaVacante = {};
  var semaforo = { verde_fuerte:[], verde_claro:[], amarillo:[], naranja:[] };
  var bajasDelMes = 0;

  // Para rotación: contar activos al inicio del mes (aproximado = activos + bajas del mes)
  var dataOps = sheetOps ? sheetOps.getDataRange().getValues() : [];
  var headOps = dataOps.length > 0 ? dataOps[0].map(function(h){ return String(h).toUpperCase().trim(); }) : [];
  var iEstado = headOps.indexOf("ESTADO"); if (iEstado===-1) iEstado=9;
  var iArea   = headOps.indexOf("AREA");   if (iArea===-1)   iArea=4;
  var iPuesto = headOps.indexOf("PUESTO"); if (iPuesto===-1) iPuesto=5;
  var iIngreso= headOps.indexOf("FECHA_INGRESO"); if (iIngreso===-1) iIngreso=headOps.indexOf("INGRESO"); if (iIngreso===-1) iIngreso=8;
  var iFechaBaja=headOps.indexOf("FECHA_BAJA"); if (iFechaBaja===-1) iFechaBaja=-1;
  // Columnas vacaciones (índice base 0): O=14 totalDisp, T=19 diasTomados, U=20 saldoAnt
  var iDispAnt = 20; // saldo del periodo anterior
  var iDiasAct = 18; // proporcionales periodo actual

  for (var i = 1; i < dataOps.length; i++) {
    var est = String(dataOps[i][iEstado]||"").toUpperCase().trim();
    var area= String(dataOps[i][iArea]  ||"Sin área").trim();
    if (est === "ACTIVO" || est === "INCAPACIDAD") {
      totalActivos++;
      porAreaActivo[area] = (porAreaActivo[area]||0) + 1;

      // Semáforo de vacaciones (basado en DÍAS DEL PERIODO ACTUAL solamente)
      var fIngreso = dataOps[i][iIngreso] instanceof Date ? dataOps[i][iIngreso] : new Date(dataOps[i][iIngreso]);
      var diasAnt  = Number(dataOps[i][iDispAnt] || 0);  // días del periodo anterior disponibles
      var diasProp = Number(dataOps[i][iDiasAct] || 0);  // proporcionales del periodo actual
      var nombre   = String(dataOps[i][1]||"Sin nombre");

      if (!isNaN(fIngreso.getTime())) {
        var aniosCump = hoy.getFullYear() - fIngreso.getFullYear();
        var anivEsteAnio = new Date(hoy.getFullYear(), fIngreso.getMonth(), fIngreso.getDate());
        if (hoy < anivEsteAnio) aniosCump--;
        var ultimoAniv = new Date(fIngreso.getFullYear() + aniosCump, fIngreso.getMonth(), fIngreso.getDate());
        var diasDesdeAniv = Math.floor((hoy - ultimoAniv) / 86400000);
        var mesesDesdeAniv = diasDesdeAniv / 30.44;

        if (diasAnt <= 0) {
          semaforo.verde_fuerte.push({nombre:nombre, area:area, diasPend:0, meses:mesesDesdeAniv.toFixed(1)});
        } else if (mesesDesdeAniv <= 6) {
          semaforo.verde_claro.push({nombre:nombre, area:area, diasPend:diasAnt, meses:mesesDesdeAniv.toFixed(1)});
        } else if (mesesDesdeAniv <= 10) {
          semaforo.amarillo.push({nombre:nombre, area:area, diasPend:diasAnt, meses:mesesDesdeAniv.toFixed(1)});
        } else {
          semaforo.naranja.push({nombre:nombre, area:area, diasPend:diasAnt, meses:mesesDesdeAniv.toFixed(1)});
        }
      }
    }
    if (est === "VACANTE") {
      totalVacantes++;
      porAreaVacante[area] = (porAreaVacante[area]||0) + 1;
    }
    if (est === "BAJA" && iFechaBaja > -1) {
      var fb = dataOps[i][iFechaBaja];
      if (fb instanceof Date && fb.getFullYear()===anioDash && fb.getMonth()===mesDash) bajasDelMes++;
    }
  }

  // Rotación = bajas del mes / ((activos inicio + activos fin) / 2) * 100
  var activosInicioAprox = totalActivos + bajasDelMes;
  var rotacion = (activosInicioAprox + totalActivos) > 0
    ? ((bajasDelMes / ((activosInicioAprox + totalActivos) / 2)) * 100).toFixed(1)
    : "0.0";

  // ── SOLICITUDES DEL MES ────────────────────────────────────────────────────
  var dataSols = sheetSols ? sheetSols.getDataRange().getValues() : [];
  var headSol  = dataSols.length > 0 ? dataSols[0].map(function(h){ return String(h).toUpperCase().trim(); }) : [];
  var iEstSol  = headSol.indexOf("ESTADO");   if (iEstSol===-1)   iEstSol=6;
  var iAreaSol = headSol.indexOf("AREA");     // puede no existir en solicitudes
  var iFechaSol= headSol.indexOf("FECHA");    if (iFechaSol===-1) iFechaSol=2;
  var iDiasSol = headSol.indexOf("DIAS");     if (iDiasSol===-1)  iDiasSol=5;
  var iIdOp    = headSol.indexOf("ID_OPERADOR"); if (iIdOp===-1)  iIdOp=1;

  var vacMes = { solicitado:0, autorizado:0, aplicado:0, cancelado:0 };
  var vacAnteriorPendiente = 0; // días de solicitudes de meses ANTERIORES no aplicadas

  for (var j = 1; j < dataSols.length; j++) {
    var est2  = String(dataSols[j][iEstSol]||"").toUpperCase().trim();
    var fSol  = dataSols[j][iFechaSol];
    var dias2 = Number(dataSols[j][iDiasSol]||0);
    var esMes = fSol instanceof Date && fSol.getMonth()===mesDash && fSol.getFullYear()===anioDash;
    var esAnterior = fSol instanceof Date && (fSol.getFullYear()<anioDash || (fSol.getFullYear()===anioDash && fSol.getMonth()<mesDash));

    if (esMes) {
      if (est2==="SOLICITADO")  vacMes.solicitado  += dias2;
      if (est2==="AUTORIZADO")  vacMes.autorizado  += dias2;
      if (est2==="APLICADO")    vacMes.aplicado    += dias2;
      if (est2==="CANCELADO")   vacMes.cancelado   += dias2;
    }
    // Solicitudes anteriores no aplicadas aún
    if (esAnterior && (est2==="SOLICITADO"||est2==="AUTORIZADO")) {
      vacAnteriorPendiente += dias2;
    }
  }

  // Vacaciones por área del mes (para gráfica)
  var vacPorArea = {};
  for (var k = 1; k < dataSols.length; k++) {
    var est3  = String(dataSols[k][iEstSol]||"").toUpperCase().trim();
    var fSol3 = dataSols[k][iFechaSol];
    var dias3 = Number(dataSols[k][iDiasSol]||0);
    var esMes3= fSol3 instanceof Date && fSol3.getMonth()===mesDash && fSol3.getFullYear()===anioDash;
    if (!esMes3 || est3==="CANCELADO") continue;
    // Buscar área del operador
    var idOp3 = String(dataSols[k][iIdOp]||"");
    var areaOp= "Sin área";
    for (var m2=1; m2<dataOps.length; m2++){
      if (String(dataOps[m2][0]||"")===idOp3){ areaOp=String(dataOps[m2][iArea]||"Sin área"); break; }
    }
    if (!vacPorArea[areaOp]) vacPorArea[areaOp]={solicitado:0,autorizado:0,aplicado:0};
    if (est3==="SOLICITADO")  vacPorArea[areaOp].solicitado+=dias3;
    if (est3==="AUTORIZADO")  vacPorArea[areaOp].autorizado+=dias3;
    if (est3==="APLICADO")    vacPorArea[areaOp].aplicado  +=dias3;
  }

  return {
    totalActivos:    totalActivos,
    totalVacantes:   totalVacantes,
    porAreaActivo:   porAreaActivo,
    porAreaVacante:  porAreaVacante,
    bajasDelMes:     bajasDelMes,
    rotacion:        rotacion,
    vacMes:          vacMes,
    vacAnteriorPendiente: vacAnteriorPendiente,
    vacPorArea:      vacPorArea,
    semaforo:        semaforo,
    mesDash:         mesDash,
    anioDash:        anioDash
  };
}

// =================================================================================
// ══════════════════  PLANTILLA / OPERADORES (movidas del GS original)  ═══════════
// =================================================================================

function obtenerDatosRH() {
  var ss = SpreadsheetApp.openById(ID_HOJA_CALCULO);
  var sheetOps = ss.getSheetByName("OPERADORES");
  var dataOps = sheetOps.getDataRange().getValues();
  var headersOps = dataOps[0].map(function(h){ return String(h).toUpperCase().trim(); });

  var getIdx = function(n) { return headersOps.indexOf(n); };
  var IDX = {
    ID: getIdx("ID"), NOMBRE: getIdx("NOMBRE"), PROCESOS: getIdx("PROCESOS"),
    FOTO: getIdx("FOTO"), AREA: getIdx("AREA"), PUESTO: getIdx("PUESTO"),
    CLAVE: getIdx("CLAVE"), TIPO: getIdx("TIPO"), INGRESO: getIdx("FECHA_INGRESO"),
    ESTADO: getIdx("ESTADO"), FISCAL: getIdx("SAL_DIA_FIS"), INTEGRADO: getIdx("SAL_DIA_INT"),
    BAJA: getIdx("FECHA_BAJA"), CELULAR: getIdx("CELULAR"), BONOS: getIdx("BONOS")
  };
  if (IDX.INGRESO   === -1) IDX.INGRESO   = getIdx("INGRESO");
  if (IDX.PROCESOS  === -1) IDX.PROCESOS  = 2;

  var operadores = [];
  var areasMap = {};
  var listaProcesosUnicos = [];

  for (var i = 1; i < dataOps.length; i++) {
    var fila = dataOps[i];
    var id = String(fila[IDX.ID]);
    if (!id) continue;

    operadores.push({
      id:          id,
      nombre:      fila[IDX.NOMBRE],
      clave:       fila[IDX.CLAVE],
      celular:     fila[IDX.CELULAR],
      foto:        fila[IDX.FOTO],
      area:        fila[IDX.AREA],
      puesto:      fila[IDX.PUESTO],
      procesos:    fila[IDX.PROCESOS],
      fiscal:      Number(fila[IDX.FISCAL])    || 0,
      integrado:   Number(fila[IDX.INTEGRADO]) || 0,
      bonos:       Number(fila[IDX.BONOS])     || 0,
      estado:      fila[IDX.ESTADO] || "ACTIVO",
      tipoNomina:  fila[IDX.TIPO],
      ingreso:     formatearFechaInput(fila[IDX.INGRESO]),
      fechaBaja:   formatearFechaInput(fila[IDX.BAJA])
    });

    var area = String(fila[IDX.AREA]).toUpperCase().trim();
    var puesto = String(fila[IDX.PUESTO]).toUpperCase().trim();
    if (area) {
      if (!areasMap[area]) areasMap[area] = [];
      if (puesto && areasMap[area].indexOf(puesto) === -1) areasMap[area].push(puesto);
    }

    var celdaProcesos = String(fila[IDX.PROCESOS]);
    if (celdaProcesos) {
      celdaProcesos.split(",").forEach(function(p) {
        var limpio = p.toUpperCase().trim();
        if (limpio && listaProcesosUnicos.indexOf(limpio) === -1) listaProcesosUnicos.push(limpio);
      });
    }
  }

  listaProcesosUnicos.sort();

  // ── Complementar procesos con ESTANDARES col C ────────────────────────────
  var sheetEst = ss.getSheetByName("ESTANDARES");
  if (sheetEst) {
    var dataEst = sheetEst.getDataRange().getValues();
    var hEst    = dataEst[0].map(function(h){ return String(h).toUpperCase().trim(); });
    var iProcEst = hEst.indexOf("PROCESO"); if (iProcEst===-1) iProcEst=2;
    for (var ie=1; ie<dataEst.length; ie++){
      var pe=String(dataEst[ie][iProcEst]||"").toUpperCase().trim();
      if(pe && listaProcesosUnicos.indexOf(pe)===-1) listaProcesosUnicos.push(pe);
    }
    listaProcesosUnicos.sort();
  }

  return { lista: operadores, catalogoAreas: areasMap, catalogoProcesos: listaProcesosUnicos };
}

function guardarLoteOperadores(listaCambios) {
  var ss = SpreadsheetApp.openById(ID_HOJA_CALCULO);
  var sheet = ss.getSheetByName("OPERADORES");
  var data = sheet.getDataRange().getValues();
  var headers = data[0].map(function(h){ return String(h).toUpperCase().trim(); });

  // Búsqueda exacta primero, luego parcial como fallback
  var getIdx = function(nombres) {
    var lista = Array.isArray(nombres) ? nombres : [nombres];
    // Paso 1: coincidencia exacta
    for (var k = 0; k < headers.length; k++) {
      for (var n = 0; n < lista.length; n++) {
        if (headers[k] === lista[n]) return k;
      }
    }
    // Paso 2: coincidencia parcial
    for (var k2 = 0; k2 < headers.length; k2++) {
      for (var n2 = 0; n2 < lista.length; n2++) {
        if (headers[k2].includes(lista[n2])) return k2;
      }
    }
    return -1;
  };

  var IDX = {
    ID:         getIdx("ID"),
    NOMBRE:     getIdx("NOMBRE"),
    PROCESOS:   getIdx("PROCESOS"),
    FOTO:       getIdx("FOTO"),
    AREA:       getIdx("AREA"),
    PUESTO:     getIdx("PUESTO"),
    CLAVE:      getIdx("CLAVE"),
    TIPO:       getIdx("TIPO"),
    INGRESO:    getIdx(["FECHA_INGRESO","INGRESO"]),
    ESTADO:     getIdx("ESTADO"),
    FISCAL:     getIdx(["SAL_DIA_FIS","FISCAL"]),
    INTEGRADO:  getIdx(["SAL_DIA_INT","INTEGRADO"]),
    BAJA:       getIdx(["FECHA_BAJA","BAJA"]),
    CELULAR:    getIdx("CELULAR"),
    BONOS:      getIdx("BONOS")
  };

  var mapaFilas = {};
  for (var i = 1; i < data.length; i++) {
    var idFila = String(data[i][IDX.ID]);
    if (idFila) mapaFilas[idFila] = i + 1;
  }

  var actualizados = 0;
  listaCambios.forEach(function(datos) {
    var filaDestino = (datos.id && mapaFilas[datos.id]) ? mapaFilas[datos.id] : -1;
    if (filaDestino === -1) {
      filaDestino = sheet.getLastRow() + 1;
      if (!datos.id) datos.id = Utilities.getUuid();
    }

    var setVal = function(idx, val) {
      if (idx > -1) {
        sheet.getRange(filaDestino, idx + 1).setValue(val);
      }
    };

    setVal(IDX.ID,         datos.id);
    setVal(IDX.NOMBRE,     datos.nombre);
    setVal(IDX.CLAVE,      datos.clave);
    setVal(IDX.CELULAR,    datos.celular);
    setVal(IDX.FOTO,       datos.foto);
    setVal(IDX.AREA,       datos.area);
    setVal(IDX.PUESTO,     datos.puesto);
    setVal(IDX.PROCESOS,   datos.procesos);
    setVal(IDX.FISCAL,     typeof datos.fiscal    === 'number' ? datos.fiscal    : (parseFloat(datos.fiscal)    || 0));
    setVal(IDX.INTEGRADO,  typeof datos.integrado === 'number' ? datos.integrado : (parseFloat(datos.integrado) || 0));
    setVal(IDX.BONOS,      typeof datos.bonos     === 'number' ? datos.bonos     : (parseFloat(datos.bonos)     || 0));
    setVal(IDX.ESTADO,     datos.estado);
    setVal(IDX.TIPO,       datos.tipoNomina);
    setVal(IDX.BAJA,       datos.fechaBaja);
    if (IDX.INGRESO > -1 && datos.ingreso) {
      var partes = String(datos.ingreso).split("-");
      if (partes.length === 3) {
        sheet.getRange(filaDestino, IDX.INGRESO + 1).setValue(new Date(partes[0], partes[1]-1, partes[2]));
      } else {
        sheet.getRange(filaDestino, IDX.INGRESO + 1).setValue(datos.ingreso);
      }
    }
    actualizados++;

    if (datos.estado === "BAJA") {
      crearVacanteAutomatica(datos.puesto || "Sin puesto", datos.area || "Sin área");
    }
  });

  return "✅ Proceso completado. " + actualizados + " registros actualizados.";
}

function guardarOperador(datos) {
  var ss = SpreadsheetApp.openById(ID_HOJA_CALCULO);
  var sheet = ss.getSheetByName("OPERADORES");
  var data = sheet.getDataRange().getValues();
  var headers = data[0];

  var getIdx = function(n) {
    for (var k = 0; k < headers.length; k++)
      if (String(headers[k]).toUpperCase().trim().includes(n)) return k;
    return -1;
  };

  var IDX = {
    ID: getIdx("ID"), NOMBRE: getIdx("NOMBRE"), PROCESOS: getIdx("PROCESOS"),
    FOTO: getIdx("FOTO"), AREA: getIdx("AREA"), PUESTO: getIdx("PUESTO"),
    CLAVE: getIdx("CLAVE"), TIPO: getIdx("TIPO"), INGRESO: getIdx("INGRESO"),
    ESTADO: getIdx("ESTADO"), FISCAL: getIdx("FISCAL"), INTEGRADO: getIdx("INTEGRADO"),
    BAJA: getIdx("BAJA"), CELULAR: getIdx("CELULAR"), BONOS: getIdx("BONOS")
  };
  if (IDX.INGRESO   === -1) IDX.INGRESO   = getIdx("FECHA_INGRESO");
  if (IDX.FISCAL    === -1) IDX.FISCAL    = getIdx("SAL_DIA_FIS");
  if (IDX.INTEGRADO === -1) IDX.INTEGRADO = getIdx("SAL_DIA_INT");
  if (IDX.BAJA      === -1) IDX.BAJA      = getIdx("FECHA_BAJA");

  var filaDestino = -1;
  if (datos.id) {
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][IDX.ID]) === String(datos.id)) { filaDestino = i + 1; break; }
    }
  }
  if (filaDestino === -1) { datos.id = Utilities.getUuid(); filaDestino = sheet.getLastRow() + 1; }

  var setVal = function(idx, val) { if (idx > -1) sheet.getRange(filaDestino, idx + 1).setValue(val); };

  setVal(IDX.ID,        datos.id);
  setVal(IDX.NOMBRE,    datos.nombre);
  setVal(IDX.CLAVE,     datos.clave);
  setVal(IDX.CELULAR,   datos.celular);
  setVal(IDX.FOTO,      datos.foto);
  setVal(IDX.AREA,      datos.area);
  setVal(IDX.PUESTO,    datos.puesto);
  setVal(IDX.PROCESOS,  datos.procesos);
  setVal(IDX.FISCAL,    datos.fiscal);
  setVal(IDX.INTEGRADO, datos.integrado);
  setVal(IDX.BONOS,     datos.bonos);
  setVal(IDX.ESTADO,    datos.estado);
  setVal(IDX.TIPO,      datos.tipoNomina);
  setVal(IDX.INGRESO,   datos.ingreso);
  setVal(IDX.BAJA,      datos.fechaBaja);

  return "✅ Guardado correctamente.";
}

// Helper de fecha (necesario para obtenerDatosRH y guardarOperador)
function formatearFechaInput(fechaRaw) {
  if (!fechaRaw || fechaRaw === "") return "";
  try {
    var d = new Date(fechaRaw);
    var year  = d.getFullYear();
    var month = ("0" + (d.getMonth() + 1)).slice(-2);
    var day   = ("0" + d.getDate()).slice(-2);
    return year + "-" + month + "-" + day;
  } catch(e) { return ""; }
}

// =================================================================================
// ══════════════════  VACACIONES (movidas del GS original)  ═══════════════════════
// =================================================================================

function obtenerDatosVacaciones() {
  actualizarSaldosYAniversarios();
  var ss = SpreadsheetApp.openById(ID_HOJA_CALCULO);
  var sheet = ss.getSheetByName("OPERADORES");
  var data = sheet.getDataRange().getValues();
  var lista = [];
  var areas = {};
  var hoy = new Date();

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][9]).trim().toUpperCase() !== "ACTIVO") continue;
    lista.push({
      id:              String(data[i][0]),
      nombre:          String(data[i][1]),
      foto:            String(data[i][3]),
      area:            String(data[i][4]),
      puesto:          String(data[i][5]),
      clave:           String(data[i][6]),
      ingreso:         data[i][8] instanceof Date ? Utilities.formatDate(data[i][8], "GMT-6", "dd/MM/yyyy") : String(data[i][8]),
      totalDisponible: Math.floor(Number(data[i][14]) || 0),
      saldoInicial:    Number(data[i][15]) || 0,
      dispAnt:         Math.floor(Number(data[i][20]) || 0),
      proporcionales:  Math.floor(Number(data[i][18]) || 0),
      diasTomados:     Math.floor(Number(data[i][19]) || 0),
      autoriza:        String(data[i][22] || ""),
      anioAnt:         hoy.getFullYear() - 1,
      anioAct:         hoy.getFullYear()
    });
    if (data[i][4]) areas[String(data[i][4]).trim()] = true;
  }
  return { lista: lista, catalogoAreas: areas };
}

function actualizarSaldosYAniversarios() {
  var ss    = SpreadsheetApp.openById(ID_HOJA_CALCULO);
  var shOp  = ss.getSheetByName("OPERADORES");
  var shSol = ss.getSheetByName("SOLICITUDES");
  var shTab = ss.getSheetByName("TABLA_VACACIONES");

  if (!shOp || !shSol || !shTab) {
    console.error("actualizarSaldosYAniversarios: falta hoja OPERADORES, SOLICITUDES o TABLA_VACACIONES en el libro ID_HOJA_CALCULO.");
    return;
  }

  var dataOp  = shOp.getDataRange().getValues();
  var dataSol = shSol.getDataRange().getValues();
  var dataTab = shTab.getDataRange().getValues();

  var hoy = new Date();
  // Fecha de hoy sin hora para comparaciones limpias
  var hoyStr = hoy.getFullYear() + '-' +
               String(hoy.getMonth()+1).padStart(2,'0') + '-' +
               String(hoy.getDate()).padStart(2,'0');

  var dicVac = {};
  var maxAnioEnTabla = 0;
  for (var t = 1; t < dataTab.length; t++) {
    var anio = parseInt(dataTab[t][0]);
    var dias = parseInt(dataTab[t][1]);
    if (!isNaN(anio)) {
      dicVac[anio] = dias;
      if (anio > maxAnioEnTabla) maxAnioEnTabla = anio;
    }
  }

  for (var i = 1; i < dataOp.length; i++) {
    if (String(dataOp[i][9]).trim().toUpperCase() !== "ACTIVO") continue;

    var idOp     = String(dataOp[i][0]).trim();
    var fIngreso = (dataOp[i][8] instanceof Date) ? dataOp[i][8] : new Date(dataOp[i][8]);
    if (isNaN(fIngreso.getTime())) continue;

    // ── Calcular años cumplidos y fechas de aniversario ──
    var aniosCumplidos      = hoy.getFullYear() - fIngreso.getFullYear();
    var anivEsteAnio        = new Date(hoy.getFullYear(), fIngreso.getMonth(), fIngreso.getDate());
    if (hoy < anivEsteAnio) aniosCumplidos--;

    var ultimoAniv   = new Date(fIngreso.getFullYear() + aniosCumplidos,     fIngreso.getMonth(), fIngreso.getDate());
    var proxAniv     = new Date(fIngreso.getFullYear() + aniosCumplidos + 1,  fIngreso.getMonth(), fIngreso.getDate());
    var fraccionAnio = (hoy - ultimoAniv) / (proxAniv - ultimoAniv);

    // Fecha del último aniversario como string YYYY-MM-DD para comparar con Col Y
    var ultimoAnivStr = ultimoAniv.getFullYear() + '-' +
                        String(ultimoAniv.getMonth()+1).padStart(2,'0') + '-' +
                        String(ultimoAniv.getDate()).padStart(2,'0');

    // Col Y [24]: ULTIMO_PROCESO — guarda la fecha en que se hizo la última renovación
    // Normalizar a YYYY-MM-DD aunque GAS lo haya interpretado como Date
    var rawY = dataOp[i][24];
    var ultimoProceso = "";
    if (rawY instanceof Date && !isNaN(rawY.getTime())) {
      ultimoProceso = rawY.getFullYear() + '-' +
                      String(rawY.getMonth()+1).padStart(2,'0') + '-' +
                      String(rawY.getDate()).padStart(2,'0');
    } else {
      ultimoProceso = String(rawY || "").trim();
    }

    // ── Sumar días APLICADO + CONSIDERAR=SI para este operador ──
    var sumaDiasSi = 0;
    for (var j = 1; j < dataSol.length; j++) {
      var idSolOp    = String(dataSol[j][1]).trim();
      var estadoJ    = String(dataSol[j][6]  || "").trim().toUpperCase();
      var considerarJ = String(dataSol[j][12] || "").trim().toUpperCase();
      if (idSolOp === idOp && estadoJ === "APLICADO" && considerarJ === "SI") {
        sumaDiasSi += Number(dataSol[j][5] || 0);
      }
    }

    // ── BLOQUE A: Renovación de periodo (solo si cumplió aniversario Y no se ha procesado aún) ──
    // Se considera "no procesado" cuando ultimoProceso NO contiene la fecha del último aniversario.
    var yaProcesoEsteAniversario = (ultimoProceso === ultimoAnivStr);

    if (hoy >= anivEsteAnio && aniosCumplidos >= 1 && !yaProcesoEsteAniversario) {
      // XX: leer Col U [20] ANTES de modificarla — saldo real del periodo que cierra
      // Para primer aniversario (aniosCumplidos=1) no había periodo anterior: siempre 0
      var XX_diasAnterior        = aniosCumplidos === 1 ? 0 : Number(dataOp[i][20] || 0);
      var YY_tomados             = sumaDiasSi;
      var ZZ_nuevos              = dicVac[aniosCumplidos] || dicVac[maxAnioEnTabla];
      // Periodo que CIERRA: año en que inició (un año antes del aniversario de hoy)
      // Periodo que ABRE:   año en que cumple aniversario hoy
      var anioAniversarioAnterior = ultimoAniv.getFullYear() - 1;
      var anioNuevoPeriodo        = ultimoAniv.getFullYear();

      // Marcar en SOLICITUDES: CONSIDERAR=NO, NOTA=Periodo AAAA
      for (var k = 1; k < dataSol.length; k++) {
        var idSolK      = String(dataSol[k][1]).trim();
        var estadoK     = String(dataSol[k][6]  || "").trim().toUpperCase();
        var considerarK = String(dataSol[k][12] || "").trim().toUpperCase();
        if (idSolK === idOp && estadoK === "APLICADO" && considerarK === "SI") {
          dataSol[k][12] = "NO";
          dataSol[k][13] = "Periodo " + anioAniversarioAnterior;
        }
      }

      var saldoPendiente = XX_diasAnterior - YY_tomados; // puede ser negativo
      var nuevoArranque  = saldoPendiente + ZZ_nuevos;

      // Respaldar Col W en Col X antes de sobreescribir
      var auditAnterior = String(dataOp[i][22] || "").trim();
      if (auditAnterior !== "") dataOp[i][23] = auditAnterior;

      dataOp[i][20] = nuevoArranque;  // Col U: DISP_PER_ANT
      dataOp[i][15] = nuevoArranque;  // Col P: SALDO_INICIAL actualizado al nuevo arranque
      dataOp[i][24] = ultimoAnivStr;  // Col Y: ULTIMO_PROCESO — marcar que este aniversario ya se procesó
      dataOp[i][22] = "Tenía " + XX_diasAnterior + " días (" + anioAniversarioAnterior + ")" +
                      " ha tomado " + YY_tomados + " días." +
                      " Tiene " + saldoPendiente + " días + " + ZZ_nuevos +
                      " días (" + anioNuevoPeriodo + ")= " + nuevoArranque + " días (inicio)";

      // Resetear sumaDiasSi porque ya se procesaron en la renovación
      sumaDiasSi = 0;
    }

    // ── BLOQUE B: Siempre actualizar días tomados, proporcionales y total ──
    // sumaDiasSi aquí es 0 si hubo renovación, o los días actuales si no hubo renovación
    var diasSiguienteAnio = dicVac[aniosCumplidos + 1] || dicVac[maxAnioEnTabla];
    var propCalculado     = Math.floor(diasSiguienteAnio * fraccionAnio);

    var saldoInicial = Number(dataOp[i][15]) || 0;          // Col P: arranque fijo del periodo
    var restanteAnt  = Math.max(0, saldoInicial - sumaDiasSi); // Col U = Col P - Col T, mínimo 0
    // Si Col T superó Col P, los días extra ya están comiendo de proporcionales
    var propConsumidos = Math.max(0, sumaDiasSi - saldoInicial);
    var propRestantes  = Math.max(0, propCalculado - propConsumidos);
    dataOp[i][18] = propCalculado;          // Col S: proporcionales totales del periodo (no baja)
    dataOp[i][19] = sumaDiasSi;             // Col T: días tomados acumulados
    dataOp[i][20] = restanteAnt;            // Col U: restante del periodo anterior (mínimo 0)
    dataOp[i][21] = propRestantes;          // Col V: proporcionales restantes después de consumo
    dataOp[i][14] = restanteAnt + propRestantes; // Col O: total disponible real
  }

  shOp.getRange(1, 1, dataOp.length,  dataOp[0].length).setValues(dataOp);
  shSol.getRange(1, 1, dataSol.length, dataSol[0].length).setValues(dataSol);
  console.log("Sincronización masiva completada.");
}

function obtenerSolicitudesEstado(estadoFiltro) {
  var ss        = SpreadsheetApp.openById(ID_HOJA_CALCULO);
  var sheetSols = ss.getSheetByName("SOLICITUDES");
  var dataSols  = sheetSols.getDataRange().getValues();
  var ops       = obtenerDatosVacaciones().lista;
  var lista     = [];

  for (var i = 1; i < dataSols.length; i++) {
    if (dataSols[i][6] === estadoFiltro) {
      var op = ops.find(function(x){ return String(x.id) === String(dataSols[i][1]); }) || {};
      lista.push({
        idSol:      dataSols[i][0],
        nombre:     op.nombre  || "Desconocido",
        clave:      op.clave   || "",
        area:       op.area    || "",
        foto:       op.foto    || "",
        autorizados: op.autoriza || "",
        inicio:     dataSols[i][3] instanceof Date ? Utilities.formatDate(dataSols[i][3], "GMT", "yyyy-MM-dd") : dataSols[i][3],
        fin:        dataSols[i][4] instanceof Date ? Utilities.formatDate(dataSols[i][4], "GMT", "yyyy-MM-dd") : dataSols[i][4],
        dias:       dataSols[i][5],
        estado:     dataSols[i][6],
        obs:        dataSols[i][9],
        folio:      dataSols[i][11]
      });
    }
  }
  return lista;
}

function procesarEstadoSolicitud(idSol, nuevoEstado, datosEditados) {
  var ss        = SpreadsheetApp.openById(ID_HOJA_CALCULO);
  var sheetSols = ss.getSheetByName("SOLICITUDES");
  var dataSols  = sheetSols.getDataRange().getValues();

  for (var i = 1; i < dataSols.length; i++) {
    if (String(dataSols[i][0]) === String(idSol)) {
      var fila = i + 1;
      sheetSols.getRange(fila, 4).setValue(datosEditados.inicio || dataSols[i][3]);
      sheetSols.getRange(fila, 5).setValue(datosEditados.fin    || dataSols[i][4]);
      sheetSols.getRange(fila, 6).setValue(datosEditados.dias   || dataSols[i][5]);
      sheetSols.getRange(fila, 7).setValue(nuevoEstado);
      if (nuevoEstado === "APLICADO") {
        ejecutarDescuentoSaldos(dataSols[i][1], Number(datosEditados.dias || dataSols[i][5]));
      }
      return "OK";
    }
  }
  return "No encontrado";
}

function calcularDiasLaborales(inicio, fin) {
  var d1 = new Date(inicio + "T00:00:00");
  var d2 = new Date(fin    + "T00:00:00");
  var count = 0;
  while (d1 <= d2) {
    if (d1.getDay() !== 0) count++;  // excluye domingos
    d1.setDate(d1.getDate() + 1);
  }
  return count;
}

function guardarSolicitudVacaciones(datos) {
  var ss    = SpreadsheetApp.openById(ID_HOJA_CALCULO);
  var shOp  = ss.getSheetByName("OPERADORES");
  var shSol = ss.getSheetByName("SOLICITUDES");

  var dataOp = shOp.getDataRange().getValues();
  var histP = 0, histS = 0, histT = 0;

  for (var i = 1; i < dataOp.length; i++) {
    if (String(dataOp[i][0]) === String(datos.idOperador)) {
      histP = dataOp[i][20]; // Col U OPERADORES → Col O SOLICITUDES: DISP_PER_ANT
      histS = dataOp[i][18]; // Col S OPERADORES → Col P SOLICITUDES: DISP_PER_ACT (proporcionales)
      histT = dataOp[i][19]; // Col T OPERADORES → Col Q SOLICITUDES: DIAS_TOMADOS
      break;
    }
  }

  var ultimoFolio = shSol.getLastRow() > 1 ? Number(shSol.getRange(shSol.getLastRow(), 12).getValue()) : 0;
  var nuevoFolio  = Utilities.formatString('%04d', (ultimoFolio + 1));

  shSol.appendRow([
    Utilities.getUuid(), datos.idOperador, new Date(),
    datos.inicio, datos.fin, datos.dias,
    "SOLICITADO", "", "", datos.observaciones, "", nuevoFolio, "SI", "",
    histP, histS, histT
  ]);

  return { success: true, folio: nuevoFolio };
}

function obtenerHistorialPeriodo(idOperador) {
  var ss   = SpreadsheetApp.openById(ID_HOJA_CALCULO);
  var data = ss.getSheetByName("SOLICITUDES").getDataRange().getValues();
  var historial = [];

  for (var i = 1; i < data.length; i++) {
    var considerar = String(data[i][12] || '').toUpperCase().trim();
    if (String(data[i][1]) === String(idOperador) && data[i][6] === "APLICADO" && considerar === "SI") {
      historial.push({
        inicio:     data[i][3] instanceof Date ? Utilities.formatDate(data[i][3], "GMT", "dd/MM/yyyy") : String(data[i][3]),
        fin:        data[i][4] instanceof Date ? Utilities.formatDate(data[i][4], "GMT", "dd/MM/yyyy") : String(data[i][4]),
        dias:       data[i][5],
        considerar: considerar
      });
    }
  }
  return historial;
}

function sincronizarVacacionesMasivo() {
  actualizarSaldosYAniversarios();
  return "Sincronización completa.";
}

function procesarLoteSolicitudes(listaIds, nuevoEstado) {
  var ss        = SpreadsheetApp.openById(ID_HOJA_CALCULO);
  var sheetSols = ss.getSheetByName("SOLICITUDES");
  var dataSols  = sheetSols.getDataRange().getValues();

  listaIds.forEach(function(idSol) {
    for (var i = 1; i < dataSols.length; i++) {
      if (String(dataSols[i][0]) === String(idSol)) {
        sheetSols.getRange(i + 1, 7).setValue(nuevoEstado);
        if (nuevoEstado === "APLICADO") {
          ejecutarDescuentoSaldos(dataSols[i][1], Number(dataSols[i][5]));
        }
        break;
      }
    }
  });
  return "OK";
}

function ejecutarDescuentoSaldos(idOp, dias) {
  var ss        = SpreadsheetApp.openById(ID_HOJA_CALCULO);
  var sheetOps  = ss.getSheetByName("OPERADORES");
  var dataOps   = sheetOps.getDataRange().getValues();

  for (var k = 1; k < dataOps.length; k++) {
    if (String(dataOps[k][0]) === String(idOp)) {
      var fila    = k + 1;
      // Solo acumulamos Col T (días tomados). Col U, V y O los recalcula
      // actualizarSaldosYAniversarios en cada ejecución del trigger diario.
      var tomados = Number(dataOps[k][19] || 0); // Col T: acumulado días tomados
      sheetOps.getRange(fila, 20).setValue(tomados + dias); // Col T: sumar días de esta solicitud
      // Col U (21), Col V (22), Col O (15): NO se tocan — los recalcula el trigger
      break;
    }
  }
}

// =================================================================================
// ══════════════════  CONSULTA INTEGRAL (ConsultaVacacionesHTML)  ═════════════════
// =================================================================================

function obtenerDatosConsulta() {
  var ss        = SpreadsheetApp.openById(ID_HOJA_CALCULO);
  var sheetOps  = ss.getSheetByName("OPERADORES");
  var sheetSols = ss.getSheetByName("SOLICITUDES");
  var dataOps   = sheetOps.getDataRange().getValues();
  var dataSols  = sheetSols.getDataRange().getValues();
  var lista     = [];

  for (var i = 1; i < dataOps.length; i++) {
    if (String(dataOps[i][9]).trim().toUpperCase() !== "ACTIVO") continue;
    var idOp = String(dataOps[i][0]);

    var susSols = dataSols
      .filter(function(r){ return String(r[1]) === idOp; })
      .reverse()
      .slice(0, 5)
      .map(function(r){
        return {
          idSol:      r[0],
          inicio:     r[3] instanceof Date ? Utilities.formatDate(r[3], "GMT-6", "yyyy-MM-dd") : String(r[3]),
          fin:        r[4] instanceof Date ? Utilities.formatDate(r[4], "GMT-6", "yyyy-MM-dd") : String(r[4]),
          dias:       r[5],
          estado:     r[6],
          obs:        String(r[9]  || ''),
          folio:      r[11],
          considerar: String(r[12] || ''),
          nota:       String(r[13] || ''),
          histP:      r[14] || 0,
          histS:      r[15] || 0,
          histT:      r[16] || 0
        };
      });

    lista.push({
      id:              idOp,
      nombre:          String(dataOps[i][1]),
      foto:            String(dataOps[i][3]),
      area:            String(dataOps[i][4]),
      puesto:          String(dataOps[i][5]),
      clave:           String(dataOps[i][6]),
      ingreso:         dataOps[i][8] instanceof Date ? Utilities.formatDate(dataOps[i][8], "GMT-6", "dd/MM/yyyy") : String(dataOps[i][8]),
      totalDisponible: Math.floor(Number(dataOps[i][14]) || 0),
      saldoInicial:    Math.floor(Number(dataOps[i][15]) || 0),
      dispAnt:         Math.floor(Number(dataOps[i][20]) || 0),
      proporcionales:  Math.floor(Number(dataOps[i][18]) || 0),
      diasTomados:     Math.floor(Number(dataOps[i][19]) || 0),
      solicitudesRecientes: susSols
    });
  }
  return lista;
}

function obtenerSolicitudesPaginadas(idOp, offset) {
  var ss       = SpreadsheetApp.openById(ID_HOJA_CALCULO);
  var sheetSols = ss.getSheetByName("SOLICITUDES");
  var dataSols  = sheetSols.getDataRange().getValues();

  var todas = dataSols
    .filter(function(r){ return String(r[1]) === String(idOp); })
    .reverse();

  return todas
    .slice(offset, offset + 5)
    .map(function(r){
      return {
        idSol:      r[0],
        inicio:     r[3] instanceof Date ? Utilities.formatDate(r[3], "GMT-6", "yyyy-MM-dd") : String(r[3]),
        fin:        r[4] instanceof Date ? Utilities.formatDate(r[4], "GMT-6", "yyyy-MM-dd") : String(r[4]),
        dias:       r[5],
        estado:     r[6],
        obs:        String(r[9]  || ''),
        folio:      r[11],
        considerar: String(r[12] || ''),
        nota:       String(r[13] || ''),
        histP:      r[14] || 0,
        histS:      r[15] || 0,
        histT:      r[16] || 0,
        hayMas:     (offset + 5) < todas.length
      };
    });
}

// =================================================================================
// ══════════════════  VACANTES  ════════════════════════════════════════════════════
// =================================================================================

// ── VACANTES: Leer/crear/cerrar en tabla OPERADORES (ESTADO = "VACANTE") ──

function obtenerVacantes() {
  var ss    = SpreadsheetApp.openById(ID_HOJA_CALCULO);
  var sheet = ss.getSheetByName("OPERADORES");
  if (!sheet) return [];
  var data    = sheet.getDataRange().getValues();
  var headers = data[0].map(function(h){ return String(h).toUpperCase().trim(); });
  var iID      = headers.indexOf("ID");
  var iPuesto  = headers.indexOf("PUESTO");
  var iArea    = headers.indexOf("AREA");
  var iEstado  = headers.indexOf("ESTADO");       if (iEstado  === -1) iEstado  = 9;
  var iIngreso = headers.indexOf("FECHA_INGRESO"); if (iIngreso === -1) iIngreso = headers.indexOf("INGRESO");
  var iFiscal  = headers.indexOf("SAL_DIA_FIS");  if (iFiscal  === -1) iFiscal  = headers.indexOf("FISCAL");
  var iInteg   = headers.indexOf("SAL_DIA_INT");  if (iInteg   === -1) iInteg   = headers.indexOf("INTEGRADO");
  var lista   = [];
  for (var i = 1; i < data.length; i++) {
    var est = String(data[i][iEstado] || "").toUpperCase().trim();
    if (est !== "VACANTE") continue;
    lista.push({
      id:         String(data[i][iID]     || i),
      puesto:     String(data[i][iPuesto] || "Sin puesto"),
      area:       String(data[i][iArea]   || "Sin área"),
      motivo:     "",
      fiscal:     Number(iFiscal  > -1 ? data[i][iFiscal]  : 0) || 0,
      integrado:  Number(iInteg   > -1 ? data[i][iInteg]   : 0) || 0,
      fecha:      data[i][iIngreso] instanceof Date
        ? Utilities.formatDate(data[i][iIngreso], "GMT-6", "dd/MM/yyyy")
        : String(data[i][iIngreso] || "")
    });
  }
  return lista;
}

function crearVacante(datos) {
  var ss    = SpreadsheetApp.openById(ID_HOJA_CALCULO);
  var sheet = ss.getSheetByName("OPERADORES");
  if (!sheet) return "ERROR: No existe hoja OPERADORES";
  var data    = sheet.getDataRange().getValues();
  var headers = data[0].map(function(h){ return String(h).toUpperCase().trim(); });
  var iID     = headers.indexOf("ID");
  var iNombre = headers.indexOf("NOMBRE");
  var iPuesto = headers.indexOf("PUESTO");
  var iArea   = headers.indexOf("AREA");
  var iEstado = headers.indexOf("ESTADO"); if (iEstado === -1) iEstado = 9;
  var iIngreso= headers.indexOf("FECHA_INGRESO"); if (iIngreso===-1) iIngreso=headers.indexOf("INGRESO");
  var newRow  = [];
  for (var c = 0; c < headers.length; c++) newRow.push("");
  newRow[iID]     = Utilities.getUuid();
  newRow[iNombre] = "VACANTE - " + (datos.puesto || "");
  newRow[iPuesto] = datos.puesto || "";
  newRow[iArea]   = datos.area   || "";
  newRow[iEstado] = "VACANTE";
  newRow[iIngreso]= "";
  sheet.appendRow(newRow);
  return "OK";
}

function editarVacante(datos) {
  var ss    = SpreadsheetApp.openById(ID_HOJA_CALCULO);
  var sheet = ss.getSheetByName("OPERADORES");
  if (!sheet) return "ERROR";
  var data    = sheet.getDataRange().getValues();
  var headers = data[0].map(function(h){ return String(h).toUpperCase().trim(); });
  var iID     = headers.indexOf("ID");
  var iNombre = headers.indexOf("NOMBRE");
  var iPuesto = headers.indexOf("PUESTO");
  var iArea   = headers.indexOf("AREA");
  var iFiscal = headers.indexOf("SAL_DIA_FIS"); if (iFiscal===-1) iFiscal=headers.indexOf("FISCAL");
  var iInteg  = headers.indexOf("SAL_DIA_INT"); if (iInteg ===-1) iInteg =headers.indexOf("INTEGRADO");
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][iID]) !== String(datos.id)) continue;
    var row = i + 1;
    if (iPuesto>-1) sheet.getRange(row,iPuesto+1).setValue(datos.puesto||"" );
    if (iArea  >-1) sheet.getRange(row,iArea  +1).setValue(datos.area  ||"" );
    if (iNombre>-1) sheet.getRange(row,iNombre+1).setValue("VACANTE - "+(datos.puesto||"" ));
    if (iFiscal>-1) sheet.getRange(row,iFiscal+1).setValue(datos.fiscal   ||0);
    if (iInteg >-1) sheet.getRange(row,iInteg +1).setValue(datos.integrado||0);
    return "OK";
  }
  return "No encontrado";
}

function cerrarVacante(datos) {
  var ss    = SpreadsheetApp.openById(ID_HOJA_CALCULO);
  var sheet = ss.getSheetByName("OPERADORES");
  if (!sheet) return "ERROR";
  var data    = sheet.getDataRange().getValues();
  var headers = data[0].map(function(h){ return String(h).toUpperCase().trim(); });
  var iID      = headers.indexOf("ID");
  var iNombre  = headers.indexOf("NOMBRE");
  var iPuesto  = headers.indexOf("PUESTO");
  var iArea    = headers.indexOf("AREA");
  var iEstado  = headers.indexOf("ESTADO");       if (iEstado ===-1) iEstado =9;
  var iClave   = headers.indexOf("CLAVE");
  var iCelular = headers.indexOf("CELULAR");
  var iIngreso = headers.indexOf("FECHA_INGRESO"); if (iIngreso===-1) iIngreso=headers.indexOf("INGRESO");
  var iFiscal  = headers.indexOf("SAL_DIA_FIS");   if (iFiscal ===-1) iFiscal =headers.indexOf("FISCAL");
  var iInteg   = headers.indexOf("SAL_DIA_INT");   if (iInteg  ===-1) iInteg  =headers.indexOf("INTEGRADO");
  var iTipo    = headers.indexOf("TIPO");
  var iBonos   = headers.indexOf("BONOS");
  var iFoto    = headers.indexOf("FOTO");
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][iID]) !== String(datos.id)) continue;
    var row = i + 1;
    var sv = function(idx,val){ if(idx>-1) sheet.getRange(row,idx+1).setValue(val); };
    sv(iNombre,  datos.nombre   ||"" );
    sv(iClave,   datos.clave    ||"" );
    sv(iPuesto,  datos.puesto   ||"" );
    sv(iArea,    datos.area     ||"" );
    sv(iEstado,  "ACTIVO");
    sv(iCelular, datos.celular  ||"" );
    sv(iFiscal,  datos.fiscal   ||0  );
    sv(iInteg,   datos.integrado||0  );
    sv(iTipo,    datos.tipoNomina||"" );
    sv(iBonos,   datos.bonos    ||0  );
    sv(iFoto,    datos.foto     ||"" );
    if (iIngreso>-1 && datos.ingreso) {
      var p=datos.ingreso.split("-");
      if(p.length===3) sheet.getRange(row,iIngreso+1).setValue(new Date(p[0],p[1]-1,p[2]));
    }
    return "OK";
  }
  return "No encontrado";
}

function crearVacanteAutomatica(puesto, area) {
  crearVacante({ puesto: puesto || "Sin puesto", area: area || "Sin área" });
}

// ── Catálogos para módulo Vacantes: áreas de USUARIOS + procesos de ESTANDARES ──
function obtenerCatalogosVacantes() {
  var ssRH  = SpreadsheetApp.openById(ID_HOJA_RH);      // TABULADOR, ESTANDARES
  var ssCal = SpreadsheetApp.openById(ID_HOJA_CALCULO);  // OPERADORES

  // ── 1. Áreas únicas desde OPERADORES (libro ID_HOJA_CALCULO) col E ──
  var areas = [];
  var sheetO = ssCal.getSheetByName("OPERADORES");
  if (sheetO) {
    var dataO = sheetO.getDataRange().getValues();
    for (var i = 1; i < dataO.length; i++) {
      var a = String(dataO[i][4] || "").toUpperCase().trim();
      if (a && areas.indexOf(a) === -1) areas.push(a);
    }
    areas.sort();
  }

  // ── 2. Procesos desde ESTANDARES (libro ID_HOJA_RH) col C ──
  var procesos = [];
  var sheetE = ssRH.getSheetByName("ESTANDARES");
  if (!sheetE) sheetE = ssCal.getSheetByName("ESTANDARES"); // fallback
  if (sheetE) {
    var dataE = sheetE.getDataRange().getValues();
    for (var j = 1; j < dataE.length; j++) {
      var p = String(dataE[j][2] || "").toUpperCase().trim();
      if (p && procesos.indexOf(p) === -1) procesos.push(p);
    }
    procesos.sort();
  }

  // ── 3. Tabulador desde hoja TABULADOR (libro ID_HOJA_RH) ──
  // Columnas: A(0)=ID  B(1)=DIARIO  C(2)=DIA_INT  D(3)=SEMANAL
  //           E(4)=CLAVO  F(5)=VARILLA  G(6)=BARRA  H(7)=TORNILLO
  //           I(8)=TREFILADO  J(9)=TRAT_TERM  K(10)=COLATADO
  //           L(11)=MANTTO  M(12)=TALLER  N(13)=INDIRECTOS

  // Helper para limpiar valores tipo "$279.00" o "$2,384.00" → número
  function limpiarNum(v) {
    if (typeof v === 'number') return v;
    return parseFloat(String(v).replace(/[$,\s]/g, '')) || 0;
  }

  // Columnas del TABULADOR: E(4)=CLAVO F(5)=VARILLA G(6)=BARRA H(7)=TORNILLO
  // I(8)=TREFILADO/DECAPADO J(9)=TRAT_TERM K(10)=COLATADO L(11)=MANTTO M(12)=TALLER N(13)=INDIRECTOS
  var AREA_A_COL = {
    'CLAVO':                   4,
    'VARILLA':                 5,
    'BARRA':                   6,
    'ESTAMPADO':               6,
    'ESTIRADO':                6,
    'TORNILLO':                7,
    'LAVADO':                  7,
    'TORNILLO CEVA':           7,
    'TORNILLO FIDE-PIJAS':     7,
    'TORNILLO FIDE PIJAS':     7,
    'TORNILLO HANREZ 1':       7,
    'TORNILLO HANREZ1':        7,
    'TORNILLO HANREZ 2':       7,
    'TORNILLO HANREZ2':        7,
    'TORNILLO NBM I':          7,
    'TORNILLO NBM 1':          7,
    'TORNILLO NBM II':         7,
    'TORNILLO NBM 2':          7,
    'TORNILLO PIJAS':          7,
    'TORNILLO PIJAS TOOL':     7,
    'DECAPADO':                8,
    'TREFILADO':               8,
    'RECOCIDO':                9,
    'TEMPLE Y REVENIDO':       9,
    'TRAT_TERMICO':            9,
    'TRAT TERMICO':            9,
    'TRATAMIENTO TERMICO':     9,
    'COLATADO':                10,
    'MANTTO':                  11,
    'MANTENIMIENTO':           11,
    'TALLER':                  12,
    'TALLER MECANIZADO':       12,
    'TALLER MECANICO':         12,
    'ADMINISTRATIVO':          13,
    'CALIDAD':                 13,
    'EMBARQUES':               13,
    'GENERALES':               13,
    'INDIRECTOS':              13,
    'SUPERVISION':             13,
    'MONTACARGAS':             13,
    'LIMPIEZA':                13
  };

  // Precalcula lista de puestos por índice de columna (1 sola pasada sobre el sheet)
  function buildListaParaCol(dataT, colIdx) {
    var lista = [];
    for (var r = 1; r < dataT.length; r++) {
      var celda = String(dataT[r][colIdx] || '').trim();
      if (!celda) continue;
      var diario  = limpiarNum(dataT[r][1]);
      var diaInt  = limpiarNum(dataT[r][2]);
      var semanal = limpiarNum(dataT[r][3]);
      celda.split(',').map(function(x){ return x.trim(); }).filter(Boolean)
        .forEach(function(puesto) {
          if (!lista.some(function(x){ return x.puesto === puesto; })) {
            lista.push({ puesto: puesto, diario: diario, diaInt: diaInt, semanal: semanal });
          }
        });
    }
    return lista;
  }

  var tabulador = {};
  var sheetT = ssRH.getSheetByName('TABULADOR');
  if (sheetT) {
    var dataT = sheetT.getDataRange().getValues();

    // Cache por columna para no releer el sheet múltiples veces
    var cacheCol = {};
    function listaDeCol(c) {
      if (!cacheCol[c]) cacheCol[c] = buildListaParaCol(dataT, c);
      return cacheCol[c];
    }

    // 1. Indexar por todos los keys del mapa (nombres canónicos y aliases)
    Object.keys(AREA_A_COL).forEach(function(areaKey) {
      tabulador[areaKey] = listaDeCol(AREA_A_COL[areaKey]);
    });

    // 2. Indexar también por cada área real de OPERADORES para garantizar match exacto
    areas.forEach(function(areaReal) {
      var k = areaReal.toUpperCase().trim();
      if (tabulador[k]) return; // ya cubierto por el mapa
      // Buscar columna por coincidencia parcial como fallback
      var colIdx = -1;
      var keys = Object.keys(AREA_A_COL);
      for (var i = 0; i < keys.length; i++) {
        var mk = keys[i].toUpperCase().trim();
        if (k.indexOf(mk) !== -1 || mk.indexOf(k) !== -1) { colIdx = AREA_A_COL[keys[i]]; break; }
      }
      tabulador[k] = colIdx >= 0 ? listaDeCol(colIdx) : [];
    });
  }

  return { areas: areas, procesos: procesos, tabulador: tabulador };
}
// =================================================================================
// ══════════════════  SANCIONES — ACTAS ADMINISTRATIVAS  ══════════════════════════
// Libro: RecursosHumanos (ID_HOJA_RH)  |  Hoja: ACTAS_ADMIN
// Columnas: ID|FOLIO|FECHA|EMPLEADO|NUM_EMPLEADO|DEPARTAMENTO|PUESTO|TIPO_FALTA
//           |DESCRIPCION|ARTICULO_REGLAMENTO|SANCION|TESTIGO_1|TESTIGO_2
//           |AUTORIZADO_POR|FIRMA_EMPLEADO|ESTATUS|OBSERVACIONES|FECHA_CREACION
// =================================================================================

function obtenerCatalogosSanciones() {
  var ssOp  = SpreadsheetApp.openById(ID_HOJA_CALCULO);
  var ssRH  = SpreadsheetApp.openById(ID_HOJA_RH);
  var shOp  = ssOp.getSheetByName("OPERADORES");
  var shAct = ssRH.getSheetByName("ACTAS_ADMIN");

  var empleados = [];
  if (shOp) {
    var dataOp  = shOp.getDataRange().getValues();
    var headOp  = dataOp[0].map(function(h){ return String(h).toUpperCase().trim(); });
    var iID     = headOp.indexOf("ID");           if (iID     === -1) iID     = 0;
    var iNom    = headOp.indexOf("NOMBRE");       if (iNom    === -1) iNom    = 1;
    var iArea   = headOp.indexOf("AREA");         if (iArea   === -1) iArea   = 4;
    var iPuesto = headOp.indexOf("PUESTO");       if (iPuesto === -1) iPuesto = 5;
    var iClave  = headOp.indexOf("CLAVE");        if (iClave  === -1) iClave  = 6;
    var iEst    = headOp.indexOf("ESTADO");       if (iEst    === -1) iEst    = 9;
    for (var i = 1; i < dataOp.length; i++) {
      var est = String(dataOp[i][iEst] || "").toUpperCase().trim();
      if (est !== "ACTIVO" && est !== "INCAPACIDAD") continue;
      empleados.push({
        id:          String(dataOp[i][iID]     || ""),
        nombre:      String(dataOp[i][iNom]    || ""),
        numEmpleado: String(dataOp[i][iClave]  || ""),
        area:        String(dataOp[i][iArea]   || ""),
        puesto:      String(dataOp[i][iPuesto] || "")
      });
    }
  }

  var conteoActas = {};
  if (shAct && shAct.getLastRow() > 1) {
    var dataAct = shAct.getDataRange().getValues();
    var headAct = dataAct[0].map(function(h){ return String(h).toUpperCase().trim(); });
    var iEmpAct = headAct.indexOf("EMPLEADO"); if (iEmpAct === -1) iEmpAct = 3;
    for (var j = 1; j < dataAct.length; j++) {
      var emp = String(dataAct[j][iEmpAct] || "").trim();
      if (!emp) continue;
      conteoActas[emp] = (conteoActas[emp] || 0) + 1;
    }
  }
  return { empleados: empleados, conteoActas: conteoActas };
}

function guardarActaAdministrativa(datos) {
  var ssRH  = SpreadsheetApp.openById(ID_HOJA_RH);
  var sheet = ssRH.getSheetByName("ACTAS_ADMIN");
  if (!sheet) return "Error: No existe la hoja ACTAS_ADMIN en el libro RecursosHumanos.";

  // Generar folio
  var lastRow = sheet.getLastRow();
  var anio    = new Date().getFullYear();
  var consec  = ("000" + lastRow).slice(-4);
  var folio   = "ACTA-" + anio + "-" + consec;

  // Leer encabezados reales de la hoja para mapear correctamente
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
                     .map(function(h){ return String(h).toUpperCase().trim(); });

  var col = function(nombre) {
    var idx = headers.indexOf(nombre);
    return idx; // -1 si no existe
  };

  // Construir fila con el tamaño exacto de columnas de la hoja
  var numCols = headers.length;
  var fila = [];
  for (var c = 0; c < numCols; c++) fila.push("");

  var set = function(nombre, val) { var i = col(nombre); if (i > -1) fila[i] = val; };

  set("ID",                   Utilities.getUuid());
  set("FOLIO",                folio);
  set("FECHA",                datos.fecha || Utilities.formatDate(new Date(), "GMT-6", "dd/MM/yyyy"));
  set("EMPLEADO",             datos.empleado      || "");
  set("NUM_EMPLEADO",         datos.numEmpleado   || "");
  set("DEPARTAMENTO",         datos.departamento  || "");
  set("PUESTO",               datos.puesto        || "");
  set("TIPO_FALTA",           datos.tipoFalta     || "");
  set("DESCRIPCION",          datos.descripcion   || "");
  set("ARTICULO_REGLAMENTO",  datos.articulo      || "");
  set("SANCION",              datos.sancion       || "");
  set("TESTIGO_1",            datos.testigo1      || "");
  set("TESTIGO_2",            datos.testigo2      || "");
  set("AUTORIZADO_POR",       datos.autorizadoPor || "");
  set("FIRMA_EMPLEADO",       "PENDIENTE");
  set("ESTATUS",              datos.estatus       || "ACTIVA");
  set("OBSERVACIONES",        datos.observaciones || "");
  set("FECHA_CREACION",       Utilities.formatDate(new Date(), "GMT-6", "dd/MM/yyyy HH:mm"));

  sheet.appendRow(fila);
  return folio;
}

function obtenerActasPorEmpleado(nombreEmpleado) {
  var ssRH  = SpreadsheetApp.openById(ID_HOJA_RH);
  var sheet = ssRH.getSheetByName("ACTAS_ADMIN");
  if (!sheet || sheet.getLastRow() < 2) return [];
  var data    = sheet.getDataRange().getValues();
  var headers = data[0].map(function(h){ return String(h).toUpperCase().trim(); });
  var iEmp   = headers.indexOf("EMPLEADO");   if (iEmp   === -1) iEmp   = 3;
  var iFolio = headers.indexOf("FOLIO");      if (iFolio === -1) iFolio = 1;
  var iFecha = headers.indexOf("FECHA");      if (iFecha === -1) iFecha = 2;
  var iTipo  = headers.indexOf("TIPO_FALTA"); if (iTipo  === -1) iTipo  = 7;
  var iSanc  = headers.indexOf("SANCION");    if (iSanc  === -1) iSanc  = 10;
  var iEst   = headers.indexOf("ESTATUS");    if (iEst   === -1) iEst   = 15;
  var lista  = [];
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][iEmp] || "").toUpperCase().trim() !== nombreEmpleado.toUpperCase().trim()) continue;
    lista.push({
      folio:     String(data[i][iFolio] || ""),
      fecha:     _fmtFechaActa(data[i][iFecha]),
      tipoFalta: String(data[i][iTipo]  || ""),
      sancion:   String(data[i][iSanc]  || ""),
      estatus:   String(data[i][iEst]   || "")
    });
  }
  return lista;
}

function _fmtHora(v){
  if(!v) return '';
  if(v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'HH:mm');
  var s = String(v).trim();
  if(/^\d{1,2}:\d{2}(:\d{2})?$/.test(s)) return s.substring(0,5);
  if(/^\d{1,2}:\d{2}\s*(AM|PM)$/i.test(s)) return s;
  return s;
}
function _fmtFechaActa(v) {
  if (!v || v === "") return "";
  if (v instanceof Date) {
    try { return Utilities.formatDate(v, "GMT-6", "dd/MM/yyyy"); }
    catch(e) { return ("0"+v.getDate()).slice(-2)+"/"+("0"+(v.getMonth()+1)).slice(-2)+"/"+v.getFullYear(); }
  }
  var s = String(v);
  if (s.indexOf("GMT") > -1 || s.indexOf("00:00:00") > -1 || s.indexOf("T00:00") > -1) {
    try { var d = new Date(s); if (!isNaN(d.getTime())) return Utilities.formatDate(d, "GMT-6", "dd/MM/yyyy"); }
    catch(e) {}
  }
  return s;
}

function obtenerTodasLasActas(filtros) {
  var ssRH  = SpreadsheetApp.openById(ID_HOJA_RH);
  var sheet = ssRH.getSheetByName("ACTAS_ADMIN");
  if (!sheet || sheet.getLastRow() < 2) return [];
  var data    = sheet.getDataRange().getValues();
  var headers = data[0].map(function(h){ return String(h).toUpperCase().trim(); });
  // Mapeo dinámico — funciona sin importar cuántas columnas tenga la hoja
  var col = function(nombre) {
    var idx = headers.indexOf(nombre);
    return idx; // devuelve -1 si no existe, lo manejamos en el push
  };
  var IDX = {
    ID:    col("ID"),   FOLIO: col("FOLIO"),  FECHA: col("FECHA"),
    EMP:   col("EMPLEADO"), NUM: col("NUM_EMPLEADO"), DEPT: col("DEPARTAMENTO"),
    PUES:  col("PUESTO"),   TIPO: col("TIPO_FALTA"),  DESC: col("DESCRIPCION"),
    ART:   col("ARTICULO_REGLAMENTO"), SANC: col("SANCION"),
    TEST1: col("TESTIGO_1"), TEST2: col("TESTIGO_2"),
    AUTH:  col("AUTORIZADO_POR"), FIRMA: col("FIRMA_EMPLEADO"),
    EST:   col("ESTATUS"), OBS: col("OBSERVACIONES"), URL: col("URL")
  };
  var f = filtros || {};
  var lista = [];
  for (var i = 1; i < data.length; i++) {
    var emp  = String(data[i][IDX.EMP]  || "").toUpperCase().trim();
    var est  = String(data[i][IDX.EST]  || "").toUpperCase().trim();
    var tipo = String(data[i][IDX.TIPO] || "").toUpperCase().trim();
    if (f.empleado  && !emp.includes(f.empleado.toUpperCase()))  continue;
    if (f.estatus   && f.estatus  !== "TODOS" && est  !== f.estatus.toUpperCase())  continue;
    if (f.tipoFalta && f.tipoFalta !== "TODOS" && !tipo.includes(f.tipoFalta.toUpperCase())) continue;
    // Filtro por mes/año
    if (f.mes || f.anio) {
      var rawFecha = data[i][IDX.FECHA];
      var dFecha = rawFecha instanceof Date ? rawFecha : new Date(rawFecha);
      if (isNaN(dFecha.getTime())) {
        // Intentar parsear formato dd/MM/yyyy
        var partes = String(rawFecha||"").split("/");
        if(partes.length===3) dFecha = new Date(partes[2], partes[1]-1, partes[0]);
      }
      if (!isNaN(dFecha.getTime())) {
        if (f.mes  && (dFecha.getMonth()+1) !== Number(f.mes))  continue;
        if (f.anio && dFecha.getFullYear()  !== Number(f.anio)) continue;
      }
    }
    var _g = function(idx) { return idx > -1 ? data[i][idx] : ""; };
    lista.push({
      id: String(_g(IDX.ID)||""), folio: String(_g(IDX.FOLIO)||""),
      fecha: _fmtFechaActa(_g(IDX.FECHA)), empleado: String(_g(IDX.EMP)||""),
      numEmpleado: String(_g(IDX.NUM)||""), departamento: String(_g(IDX.DEPT)||""),
      puesto: String(_g(IDX.PUES)||""), tipoFalta: String(_g(IDX.TIPO)||""),
      descripcion: String(_g(IDX.DESC)||""), articulo: String(_g(IDX.ART)||""),
      sancion: String(_g(IDX.SANC)||""), testigo1: String(_g(IDX.TEST1)||""),
      testigo2: String(_g(IDX.TEST2)||""), autorizadoPor: String(_g(IDX.AUTH)||""),
      firmaEmpleado: String(_g(IDX.FIRMA)||""), estatus: String(_g(IDX.EST)||""),
      observaciones: String(_g(IDX.OBS)||""),
      url: String(_g(IDX.URL)||"")
    });
  }
  lista.reverse();
  return lista;
}

function subirPdfActa(payload) {
  // payload: { folio, nombre, base64 }
  try {
    var carpeta = DriveApp.getRootFolder();
    var nombre  = 'ACTA_' + payload.folio + '_' + payload.nombre;
    // Eliminar archivo anterior con mismo folio si existe
    var iter = carpeta.getFilesByName(nombre);
    while(iter.hasNext()) iter.next().setTrashed(true);
    // Crear nuevo archivo
    var decoded = Utilities.base64Decode(payload.base64);
    var blob    = Utilities.newBlob(decoded, 'application/pdf', nombre);
    var file    = carpeta.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    var url = 'https://drive.google.com/file/d/' + file.getId() + '/view';
    // Guardar URL en columna S (URL) de ACTAS_ADMIN
    var ssRH  = SpreadsheetApp.openById(ID_HOJA_RH);
    var sheet = ssRH.getSheetByName("ACTAS_ADMIN");
    if(sheet && sheet.getLastRow() > 1){
      var data    = sheet.getDataRange().getValues();
      var headers = data[0].map(function(h){ return String(h).toUpperCase().trim(); });
      var iFolio  = headers.indexOf("FOLIO");
      var iUrl    = headers.indexOf("URL");
      if(iFolio > -1 && iUrl > -1){
        for(var i = 1; i < data.length; i++){
          if(String(data[i][iFolio]||"") === String(payload.folio)){
            sheet.getRange(i+1, iUrl+1).setValue(url);
            break;
          }
        }
      }
    }
    return url;
  } catch(e) {
    return 'ERROR: ' + e.toString();
  }
}

function actualizarEstatusActa(folio, nuevoEstatus) {
  var ssRH  = SpreadsheetApp.openById(ID_HOJA_RH);
  var sheet = ssRH.getSheetByName("ACTAS_ADMIN");
  if (!sheet) return "Error: Hoja no encontrada.";
  var data    = sheet.getDataRange().getValues();
  var headers = data[0].map(function(h){ return String(h).toUpperCase().trim(); });
  var iFolio  = headers.indexOf("FOLIO");   if (iFolio  === -1) iFolio  = 1;
  var iEst    = headers.indexOf("ESTATUS"); if (iEst    === -1) iEst    = 15;
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][iFolio] || "") === folio) {
      sheet.getRange(i + 1, iEst + 1).setValue(nuevoEstatus);
      return "OK";
    }
  }
  return "Acta no encontrada.";
}

function actualizarActaAdministrativa(datos) {
  var ssRH  = SpreadsheetApp.openById(ID_HOJA_RH);
  var sheet = ssRH.getSheetByName("ACTAS_ADMIN");
  if (!sheet) return "Error: No existe la hoja ACTAS_ADMIN.";
  var data    = sheet.getDataRange().getValues();
  var headers = data[0].map(function(h){ return String(h).toUpperCase().trim(); });

  var col = function(nombre) { return headers.indexOf(nombre); };
  var iFolio = col("FOLIO"); if (iFolio === -1) iFolio = 1;

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][iFolio] || "") !== String(datos.folio)) continue;
    var row = i + 1;
    var set = function(nombre, val) {
      var idx = col(nombre);
      if (idx > -1) sheet.getRange(row, idx + 1).setValue(val);
    };
    set("FECHA",               datos.fecha        || "");
    set("DEPARTAMENTO",        datos.departamento || "");
    set("PUESTO",              datos.puesto       || "");
    set("TIPO_FALTA",          datos.tipoFalta    || "");
    set("DESCRIPCION",         datos.descripcion  || "");
    set("ARTICULO_REGLAMENTO", datos.articulo     || "");
    set("SANCION",             datos.sancion      || "");
    set("AUTORIZADO_POR",      datos.autorizadoPor|| "");
    set("TESTIGO_1",           datos.testigo1     || "");
    set("TESTIGO_2",           datos.testigo2     || "");
    set("OBSERVACIONES",       datos.observaciones|| "");
    return "OK";
  }
  return "Acta no encontrada con folio: " + datos.folio;
}

// =================================================================================
// ══  UTILIDAD: Recrear tabla ACTAS_ADMIN con columnas correctas (18 columnas)  ══
// Ejecuta esta función UNA VEZ desde Apps Script si la tabla actual tiene 14 cols.
// IMPORTANTE: Borrará las filas existentes y creará los encabezados correctos.
// =================================================================================
function recrearTablaActasAdmin() {
  var ss    = SpreadsheetApp.openById(ID_HOJA_RH);
  var sheet = ss.getSheetByName("ACTAS_ADMIN");
  if (!sheet) {
    sheet = ss.insertSheet("ACTAS_ADMIN");
    Logger.log("Hoja ACTAS_ADMIN creada nueva.");
  }

  // Limpiar todo
  sheet.clearContents();

  // Encabezados correctos — exactamente 18 columnas
  var headers = [
    "ID", "FOLIO", "FECHA", "EMPLEADO", "NUM_EMPLEADO",
    "DEPARTAMENTO", "PUESTO", "TIPO_FALTA", "DESCRIPCION",
    "ARTICULO_REGLAMENTO", "SANCION", "TESTIGO_1", "TESTIGO_2",
    "AUTORIZADO_POR", "FIRMA_EMPLEADO", "ESTATUS",
    "OBSERVACIONES", "FECHA_CREACION"
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // Formato encabezado
  var hRange = sheet.getRange(1, 1, 1, headers.length);
  hRange.setBackground("#1a237e");
  hRange.setFontColor("white");
  hRange.setFontWeight("bold");
  hRange.setFontSize(10);
  sheet.setFrozenRows(1);

  // Ancho de columnas
  var anchos = [120,160,100,200,120,160,160,200,300,220,200,160,160,160,140,100,220,160];
  for (var c = 0; c < anchos.length; c++) {
    sheet.setColumnWidth(c + 1, anchos[c]);
  }

  SpreadsheetApp.flush();
  Logger.log("✅ ACTAS_ADMIN recreada con 18 columnas correctas.");
  try {
    SpreadsheetApp.getUi().alert("✅ Tabla ACTAS_ADMIN recreada correctamente con 18 columnas.");
  } catch(e) {}
}


// =================================================================================
// ══════════════════  INCIDENCIAS  ════════════════════════════════════════════════
// Libro: RecursosHumanos (ID_HOJA_RH)
// Tablas: AUSENTISMO | RETARDOS | PERMISOS_EMPLEADOS
// NOTA: Las tablas necesitan recrearse con columnas adicionales.
//       Ejecuta recrearTablasIncidencias() UNA VEZ antes de usar el módulo.
// Columnas AUSENTISMO (12):
//   ID | FOLIO | FECHA | EMPLEADO | NUM_EMPLEADO | DEPARTAMENTO | PUESTO
//   | TURNO | MOTIVO | JUSTIFICADO | DOCUMENTO_SOPORTE | OBSERVACIONES
// Columnas RETARDOS (13):
//   ID | FOLIO | FECHA | EMPLEADO | NUM_EMPLEADO | DEPARTAMENTO | PUESTO
//   | TURNO | HORA_REAL | HORA_PROGRAMADA | MINUTOS_RETARDO | JUSTIFICADO | OBSERVACIONES
// Columnas PERMISOS_EMPLEADOS (15):
//   ID | FOLIO | FECHA_SOLICITUD | EMPLEADO | NUM_EMPLEADO | DEPARTAMENTO | PUESTO
//   | TIPO_PERMISO | FECHA_INICIO | FECHA_FIN | DIAS_SOLICITADOS
//   | MOTIVO | AUTORIZADO_POR | ESTATUS | OBSERVACIONES
// =================================================================================

// ── Helper compartido para mapear columnas por nombre ──────────────────────────
function _colIdx(headers, nombre) {
  var idx = headers.indexOf(nombre.toUpperCase().trim());
  return idx; // -1 si no existe
}

// ── Catálogos comunes para los 3 sub-módulos ───────────────────────────────────
function obtenerCatalogosIncidencias() {
  var ssOp = SpreadsheetApp.openById(ID_HOJA_CALCULO);
  var shOp = ssOp.getSheetByName("OPERADORES");
  var empleados = [];
  if (shOp) {
    var dataOp = shOp.getDataRange().getValues();
    var hOp    = dataOp[0].map(function(h){ return String(h).toUpperCase().trim(); });
    var iID    = hOp.indexOf("ID");           if (iID    < 0) iID    = 0;
    var iNom   = hOp.indexOf("NOMBRE");       if (iNom   < 0) iNom   = 1;
    var iArea  = hOp.indexOf("AREA");         if (iArea  < 0) iArea  = 4;
    var iPues  = hOp.indexOf("PUESTO");       if (iPues  < 0) iPues  = 5;
    var iClave = hOp.indexOf("CLAVE");        if (iClave < 0) iClave = 6;
    var iEst   = hOp.indexOf("ESTADO");       if (iEst   < 0) iEst   = 9;
    for (var i = 1; i < dataOp.length; i++) {
      var est = String(dataOp[i][iEst] || "").toUpperCase().trim();
      if (est !== "ACTIVO" && est !== "INCAPACIDAD") continue;
      empleados.push({
        nombre:      String(dataOp[i][iNom]   || ""),
        numEmpleado: String(dataOp[i][iClave] || ""),
        area:        String(dataOp[i][iArea]  || ""),
        puesto:      String(dataOp[i][iPues]  || "")
      });
    }
  }
  return { empleados: empleados };
}

// ══════════════════════════════════════════════════════════════════════
//  AUSENTISMO
// ══════════════════════════════════════════════════════════════════════
function guardarAusentismo(datos) {
  var ssRH  = SpreadsheetApp.openById(ID_HOJA_RH);
  var sheet = ssRH.getSheetByName("AUSENTISMO");
  if (!sheet) return "Error: Hoja AUSENTISMO no existe. Ejecuta recrearTablasIncidencias()";
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
                     .map(function(h){ return String(h).toUpperCase().trim(); });
  var lastRow = sheet.getLastRow();
  var anio    = new Date().getFullYear();
  var folio   = "AUS-" + anio + "-" + ("000" + lastRow).slice(-4);
  var numCols = headers.length;
  var fila    = []; for (var c = 0; c < numCols; c++) fila.push("");
  var set = function(col, val) { var i = _colIdx(headers, col); if (i > -1) fila[i] = val; };
  set("ID",                Utilities.getUuid());
  set("FOLIO",             folio);
  set("FECHA",             datos.fecha        || "");
  set("EMPLEADO",          datos.empleado     || "");
  set("NUM_EMPLEADO",      datos.numEmpleado  || "");
  set("DEPARTAMENTO",      datos.departamento || "");
  set("PUESTO",            datos.puesto       || "");
  set("TURNO",             datos.turno        || "");
  set("MOTIVO",            datos.motivo       || "");
  set("JUSTIFICADO",       datos.justificado  || "NO");
  set("DOCUMENTO_SOPORTE", datos.documento    || "");
  set("OBSERVACIONES",     datos.observaciones|| "");
  sheet.appendRow(fila);
  return folio;
}

function obtenerAusentismo(filtros) {
  var ssRH  = SpreadsheetApp.openById(ID_HOJA_RH);
  var sheet = ssRH.getSheetByName("AUSENTISMO");
  if (!sheet || sheet.getLastRow() < 2) return [];
  var data    = sheet.getDataRange().getValues();
  var headers = data[0].map(function(h){ return String(h).toUpperCase().trim(); });
  var f = filtros || {};
  var filtrarMes  = (f.mes  !== undefined && f.mes  !== null);
  var filtrarAnio = (f.anio !== undefined && f.anio !== null);
  var lista = [];
  for (var i = 1; i < data.length; i++) {
    var _g = function(col) { var idx = _colIdx(headers, col); return idx > -1 ? data[i][idx] : ""; };
    if (filtrarMes || filtrarAnio) {
      var rawFecha = _g("FECHA");
      var fechaStr = rawFecha instanceof Date
        ? _fmtFechaActa(rawFecha)
        : String(rawFecha || "");
      var partes = fechaStr.split("/");
      // Formato esperado: dd/MM/yyyy
      if (partes.length === 3) {
        if (filtrarAnio && Number(partes[2]) !== Number(f.anio)) continue;
        if (filtrarMes  && (Number(partes[1]) - 1) !== Number(f.mes)) continue;
      }
    }
    if (f.empleado && !String(_g("EMPLEADO")||"").toUpperCase().includes(f.empleado.toUpperCase())) continue;
    if (f.justificado && f.justificado !== "TODOS" && String(_g("JUSTIFICADO")||"").toUpperCase() !== f.justificado.toUpperCase()) continue;
    lista.push({
      folio:        String(_g("FOLIO")            || ""),
      fecha:        _fmtFechaActa(_g("FECHA")),
      empleado:     String(_g("EMPLEADO")         || ""),
      numEmpleado:  String(_g("NUM_EMPLEADO")      || ""),
      departamento: String(_g("DEPARTAMENTO")      || ""),
      puesto:       String(_g("PUESTO")            || ""),
      turno:        String(_g("TURNO")             || ""),
      motivo:       String(_g("MOTIVO")            || ""),
      justificado:  String(_g("JUSTIFICADO")       || "NO"),
      documento:    String(_g("DOCUMENTO_SOPORTE") || ""),
      observaciones:String(_g("OBSERVACIONES")     || "")
    });
  }
  lista.reverse();
  return lista;
}

function eliminarIncidencia(folio, hoja) {
  var ssRH  = SpreadsheetApp.openById(ID_HOJA_RH);
  var sheet = ssRH.getSheetByName(hoja);
  if (!sheet) return "Hoja no encontrada";
  var data    = sheet.getDataRange().getValues();
  var headers = data[0].map(function(h){ return String(h).toUpperCase().trim(); });
  var iFolio  = _colIdx(headers, "FOLIO"); if (iFolio < 0) iFolio = 1;
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][iFolio] || "") === folio) {
      sheet.deleteRow(i + 1);
      return "OK";
    }
  }
  return "Registro no encontrado";
}

// ══════════════════════════════════════════════════════════════════════
//  RETARDOS
// ══════════════════════════════════════════════════════════════════════
function guardarRetardo(datos) {
  var ssRH  = SpreadsheetApp.openById(ID_HOJA_RH);
  var sheet = ssRH.getSheetByName("RETARDOS");
  if (!sheet) return "Error: Hoja RETARDOS no existe. Ejecuta recrearTablasIncidencias()";
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
                     .map(function(h){ return String(h).toUpperCase().trim(); });
  var lastRow = sheet.getLastRow();
  var anio    = new Date().getFullYear();
  var folio   = "RET-" + anio + "-" + ("000" + lastRow).slice(-4);
  var numCols = headers.length;
  var fila    = []; for (var c = 0; c < numCols; c++) fila.push("");
  var set = function(col, val) { var i = _colIdx(headers, col); if (i > -1) fila[i] = val; };
  // Calcular minutos automáticamente si vienen las horas
  var minutos = Number(datos.minutosRetardo || 0);
  if (!minutos && datos.horaReal && datos.horaProgramada) {
    var toMin = function(t) {
      var p = String(t).split(":"); return Number(p[0] || 0) * 60 + Number(p[1] || 0);
    };
    minutos = Math.max(0, toMin(datos.horaReal) - toMin(datos.horaProgramada));
  }
  set("ID",               Utilities.getUuid());
  set("FOLIO",            folio);
  set("FECHA",            datos.fecha          || "");
  set("EMPLEADO",         datos.empleado       || "");
  set("NUM_EMPLEADO",     datos.numEmpleado    || "");
  set("DEPARTAMENTO",     datos.departamento   || "");
  set("PUESTO",           datos.puesto         || "");
  set("TURNO",            datos.turno          || "");
  set("HORA_REAL",        datos.horaReal       || "");
  set("HORA_PROGRAMADA",  datos.horaProgramada || "");
  set("MINUTOS_RETARDO",  minutos);
  set("JUSTIFICADO",      datos.justificado    || "NO");
  set("OBSERVACIONES",    datos.observaciones  || "");
  sheet.appendRow(fila);
  return folio;
}

function obtenerRetardos(filtros) {
  var ssRH  = SpreadsheetApp.openById(ID_HOJA_RH);
  var sheet = ssRH.getSheetByName("RETARDOS");
  if (!sheet || sheet.getLastRow() < 2) return [];
  var data    = sheet.getDataRange().getValues();
  var headers = data[0].map(function(h){ return String(h).toUpperCase().trim(); });
  var f = filtros || {};
  var filtrarMes  = (f.mes  !== undefined && f.mes  !== null);
  var filtrarAnio = (f.anio !== undefined && f.anio !== null);
  var lista = [];
  for (var i = 1; i < data.length; i++) {
    var _g = function(col) { var idx = _colIdx(headers, col); return idx > -1 ? data[i][idx] : ""; };
    if (filtrarMes || filtrarAnio) {
      var rawFecha2 = _g("FECHA");
      var fechaStr2 = rawFecha2 instanceof Date
        ? _fmtFechaActa(rawFecha2)
        : String(rawFecha2 || "");
      var partes2 = fechaStr2.split("/");
      if (partes2.length === 3) {
        if (filtrarAnio && Number(partes2[2]) !== Number(f.anio)) continue;
        if (filtrarMes  && (Number(partes2[1]) - 1) !== Number(f.mes)) continue;
      }
    }
    if (f.empleado && !String(_g("EMPLEADO")||"").toUpperCase().includes(f.empleado.toUpperCase())) continue;
    if (f.justificado && f.justificado !== "TODOS" && String(_g("JUSTIFICADO")||"").toUpperCase() !== f.justificado.toUpperCase()) continue;
    lista.push({
      folio:          String(_g("FOLIO")           || ""),
      fecha:          _fmtFechaActa(_g("FECHA")),
      empleado:       String(_g("EMPLEADO")        || ""),
      numEmpleado:    String(_g("NUM_EMPLEADO")     || ""),
      departamento:   String(_g("DEPARTAMENTO")    || ""),
      puesto:         String(_g("PUESTO")          || ""),
      turno:          String(_g("TURNO")           || ""),
      horaReal:       _fmtHora(_g("HORA_REAL")),
      horaProgramada: _fmtHora(_g("HORA_PROGRAMADA")),
      minutosRetardo: String(_g("MINUTOS_RETARDO") || "0"),
      justificado:    String(_g("JUSTIFICADO")     || "NO"),
      observaciones:  String(_g("OBSERVACIONES")   || "")
    });
  }
  lista.reverse();
  return lista;
}

// ══════════════════════════════════════════════════════════════════════
//  PERMISOS DE EMPLEADOS
// ══════════════════════════════════════════════════════════════════════
function guardarPermisoEmpleado(datos) {
  var ssRH  = SpreadsheetApp.openById(ID_HOJA_RH);
  var sheet = ssRH.getSheetByName("PERMISOS_EMPLEADOS");
  if (!sheet) return "Error: Hoja PERMISOS_EMPLEADOS no existe. Ejecuta recrearTablasIncidencias()";
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
                     .map(function(h){ return String(h).toUpperCase().trim(); });
  var lastRow = sheet.getLastRow();
  var anio    = new Date().getFullYear();
  var folio   = "PER-" + anio + "-" + ("000" + lastRow).slice(-4);
  var numCols = headers.length;
  var fila    = []; for (var c = 0; c < numCols; c++) fila.push("");
  var set = function(col, val) { var i = _colIdx(headers, col); if (i > -1) fila[i] = val; };
  // Calcular días automáticamente si vienen las fechas
  var dias = Number(datos.diasSolicitados || 0);
  if (!dias && datos.fechaInicio && datos.fechaFin) {
    try {
      var d1 = new Date(datos.fechaInicio + "T00:00:00");
      var d2 = new Date(datos.fechaFin    + "T00:00:00");
      dias   = Math.max(1, Math.round((d2 - d1) / 86400000) + 1);
    } catch(e) {}
  }
  set("ID",               Utilities.getUuid());
  set("FOLIO",            folio);
  set("FECHA_SOLICITUD",  datos.fechaSolicitud || Utilities.formatDate(new Date(), "GMT-6", "dd/MM/yyyy"));
  set("EMPLEADO",         datos.empleado       || "");
  set("NUM_EMPLEADO",     datos.numEmpleado    || "");
  set("DEPARTAMENTO",     datos.departamento   || "");
  set("PUESTO",           datos.puesto         || "");
  set("TIPO_PERMISO",      datos.tipoPermiso     || "");
  set("HORA_SALIDA",       datos.horaSalida      || "");
  set("HORA_REGRESO",      datos.horaRegreso     || "");
  set("FECHA_INICIO",      datos.fechaInicio     || "");
  set("FECHA_FIN",         datos.fechaFin        || "");
  set("DIAS_SOLICITADOS",  dias);
  set("MOTIVO",            datos.motivo          || "");
  set("AUTORIZADO_POR",    datos.autorizadoPor   || "");
  set("ESTATUS",           datos.estatus         || "PENDIENTE");
  set("FECHA_PAGO_TIEMPO", datos.fechaPagoTiempo || "");
  set("HORA_INICIO_PAGO",  datos.horaInicioPago  || "");
  set("HORA_FIN_PAGO",     datos.horaFinPago     || "");
  set("OBSERVACIONES",     datos.observaciones   || "");
  sheet.appendRow(fila);
  return folio;
}

function obtenerPermisosEmpleados(filtros) {
  var ssRH  = SpreadsheetApp.openById(ID_HOJA_RH);
  var sheet = ssRH.getSheetByName("PERMISOS_EMPLEADOS");
  if (!sheet || sheet.getLastRow() < 2) return [];
  var data    = sheet.getDataRange().getValues();
  var headers = data[0].map(function(h){ return String(h).toUpperCase().trim(); });
  var f = filtros || {};
  var filtrarMes  = (f.mes  !== undefined && f.mes  !== null);
  var filtrarAnio = (f.anio !== undefined && f.anio !== null);
  var lista = [];
  for (var i = 1; i < data.length; i++) {
    var _g = function(col) { var idx = _colIdx(headers, col); return idx > -1 ? data[i][idx] : ""; };
    var emp  = String(_g("EMPLEADO") || "").toUpperCase();
    var est  = String(_g("ESTATUS")  || "").toUpperCase();
    if (filtrarMes || filtrarAnio) {
      var rawSol = _g("FECHA_SOLICITUD");
      var fSol = rawSol instanceof Date
        ? _fmtFechaActa(rawSol)
        : String(rawSol || "");
      var pSol = fSol.split("/");
      if (pSol.length === 3) {
        if (filtrarAnio && Number(pSol[2]) !== Number(f.anio)) continue;
        if (filtrarMes  && (Number(pSol[1]) - 1) !== Number(f.mes)) continue;
      }
    }
    if (f.empleado && !emp.includes(f.empleado.toUpperCase())) continue;
    if (f.estatus  && f.estatus !== "TODOS" && est !== f.estatus.toUpperCase()) continue;
    lista.push({
      folio:           String(_g("FOLIO")             || ""),
      fechaSolicitud:  _fmtFechaActa(_g("FECHA_SOLICITUD")),
      empleado:        String(_g("EMPLEADO")          || ""),
      numEmpleado:     String(_g("NUM_EMPLEADO")       || ""),
      departamento:    String(_g("DEPARTAMENTO")       || ""),
      puesto:          String(_g("PUESTO")             || ""),
      tipoPermiso:     String(_g("TIPO_PERMISO")       || ""),
      horaSalida:      _fmtHora(_g("HORA_SALIDA")),
      horaRegreso:     _fmtHora(_g("HORA_REGRESO")),
      fechaInicio:     _fmtFechaActa(_g("FECHA_INICIO")),
      fechaFin:        _fmtFechaActa(_g("FECHA_FIN")),
      diasSolicitados: String(_g("DIAS_SOLICITADOS")   || ""),
      motivo:          String(_g("MOTIVO")             || ""),
      autorizadoPor:   String(_g("AUTORIZADO_POR")     || ""),
      estatus:         String(_g("ESTATUS")            || ""),
      fechaPagoTiempo: _fmtFechaActa(_g("FECHA_PAGO_TIEMPO")),
      horaInicioPago:  String(_g("HORA_INICIO_PAGO")   || ""),
      horaFinPago:     String(_g("HORA_FIN_PAGO")      || ""),
      observaciones:   String(_g("OBSERVACIONES")      || "")
    });
  }
  lista.reverse();
  return lista;
}

function actualizarEstatusPermiso(folio, nuevoEstatus) {
  var ssRH  = SpreadsheetApp.openById(ID_HOJA_RH);
  var sheet = ssRH.getSheetByName("PERMISOS_EMPLEADOS");
  if (!sheet) return "Hoja no encontrada";
  var data    = sheet.getDataRange().getValues();
  var headers = data[0].map(function(h){ return String(h).toUpperCase().trim(); });
  var iFolio  = _colIdx(headers, "FOLIO");   if (iFolio  < 0) iFolio  = 1;
  var iEst    = _colIdx(headers, "ESTATUS"); if (iEst    < 0) iEst    = 13;
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][iFolio] || "") === folio) {
      sheet.getRange(i + 1, iEst + 1).setValue(nuevoEstatus);
      return "OK";
    }
  }
  return "No encontrado";
}

// Editar cualquier registro de incidencia (Ausentismo, Retardos, Permisos) por folio
function editarRegistroIncidencia(payload) {
  // payload: { hoja: "AUSENTISMO"|"RETARDOS"|"PERMISOS_EMPLEADOS", campos: { folio, ...camposAEditar } }
  var ssRH  = SpreadsheetApp.openById(ID_HOJA_RH);
  var sheet = ssRH.getSheetByName(payload.hoja);
  if (!sheet) return "Error: Hoja " + payload.hoja + " no encontrada";
  var data    = sheet.getDataRange().getValues();
  var headers = data[0].map(function(h){ return String(h).toUpperCase().trim(); });
  var folio   = String(payload.campos.folio || "").trim();
  var iFolio  = _colIdx(headers, "FOLIO"); if (iFolio < 0) iFolio = 1;
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][iFolio] || "").trim() !== folio) continue;
    // Actualizar sólo los campos que vienen en payload.campos (excepto folio y empleado)
    var c = payload.campos;
    var set = function(col, val) {
      var idx = _colIdx(headers, col);
      if (idx > -1 && val !== undefined && val !== null) sheet.getRange(i + 1, idx + 1).setValue(val);
    };
    if (c.fecha)          set("FECHA",            c.fecha);
    if (c.turno)          set("TURNO",            c.turno);
    if (c.motivo)         set("MOTIVO",           c.motivo);
    if (c.justificado)    set("JUSTIFICADO",      c.justificado);
    if (c.documento !== undefined) set("DOCUMENTO_SOPORTE", c.documento);
    if (c.horaProgramada) set("HORA_PROGRAMADA",  c.horaProgramada);
    if (c.horaReal)       set("HORA_REAL",        c.horaReal);
    if (c.minutosRetardo !== undefined) set("MINUTOS_RETARDO", Number(c.minutosRetardo)||0);
    if (c.tipoPermiso)    set("TIPO_PERMISO",     c.tipoPermiso);
    if (c.horaSalida !== undefined)  set("HORA_SALIDA",  c.horaSalida);
    if (c.horaRegreso !== undefined) set("HORA_REGRESO", c.horaRegreso);
    if (c.fechaInicio)    set("FECHA_INICIO",     c.fechaInicio);
    if (c.fechaFin)       set("FECHA_FIN",        c.fechaFin);
    if (c.diasSolicitados !== undefined) set("DIAS_SOLICITADOS", Number(c.diasSolicitados)||1);
    if (c.autorizadoPor !== undefined)   set("AUTORIZADO_POR",   c.autorizadoPor);
    if (c.observaciones !== undefined)   set("OBSERVACIONES",    c.observaciones);
    return "OK";
  }
  return "No encontrado: " + folio;
}

// Aprobar permiso tipo PAGO_TIEMPO guardando también fecha/horas de pago
function aprobarPermisoConPagoTiempo(datos) {
  // datos: { folio, nuevoEstatus, fechaPagoTiempo, horaInicioPago, horaFinPago }
  var ssRH  = SpreadsheetApp.openById(ID_HOJA_RH);
  var sheet = ssRH.getSheetByName("PERMISOS_EMPLEADOS");
  if (!sheet) return "Hoja no encontrada";
  var data    = sheet.getDataRange().getValues();
  var headers = data[0].map(function(h){ return String(h).toUpperCase().trim(); });
  var iFolio  = _colIdx(headers, "FOLIO");            if (iFolio < 0) iFolio = 1;
  var iEst    = _colIdx(headers, "ESTATUS");
  var iFPago  = _colIdx(headers, "FECHA_PAGO_TIEMPO");
  var iHIni   = _colIdx(headers, "HORA_INICIO_PAGO");
  var iHFin   = _colIdx(headers, "HORA_FIN_PAGO");
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][iFolio] || "") !== datos.folio) continue;
    var row = i + 1;
    if (iEst   > -1) sheet.getRange(row, iEst   + 1).setValue(datos.nuevoEstatus || "APROBADO");
    if (iFPago > -1) sheet.getRange(row, iFPago + 1).setValue(datos.fechaPagoTiempo || "");
    if (iHIni  > -1) sheet.getRange(row, iHIni  + 1).setValue(datos.horaInicioPago  || "");
    if (iHFin  > -1) sheet.getRange(row, iHFin  + 1).setValue(datos.horaFinPago     || "");
    return "OK";
  }
  return "Folio no encontrado: " + datos.folio;
}

// =================================================================================
// ══  UTILIDAD: Recrear tablas Incidencias con columnas correctas  ══════════════
// Ejecuta recrearTablasIncidencias() UNA VEZ si las tablas tienen columnas antiguas.
// =================================================================================
function recrearTablasIncidencias() {
  var ss = SpreadsheetApp.openById(ID_HOJA_RH);

  var tablas = [
    {
      nombre: "AUSENTISMO",
      headers: ["ID","FOLIO","FECHA","EMPLEADO","NUM_EMPLEADO","DEPARTAMENTO",
                "PUESTO","TURNO","MOTIVO","JUSTIFICADO","DOCUMENTO_SOPORTE","OBSERVACIONES"],
      anchos:  [120,160,100,200,120,160,160,100,200,100,160,200]
    },
    {
      nombre: "RETARDOS",
      headers: ["ID","FOLIO","FECHA","EMPLEADO","NUM_EMPLEADO","DEPARTAMENTO",
                "PUESTO","TURNO","HORA_REAL","HORA_PROGRAMADA","MINUTOS_RETARDO",
                "JUSTIFICADO","OBSERVACIONES"],
      anchos:  [120,160,100,200,120,160,160,100,120,140,120,100,200]
    },
    {
      nombre: "PERMISOS_EMPLEADOS",
      headers: ["ID","FOLIO","FECHA_SOLICITUD","EMPLEADO","NUM_EMPLEADO","DEPARTAMENTO",
                "PUESTO","TIPO_PERMISO","HORA_SALIDA","HORA_REGRESO",
                "FECHA_INICIO","FECHA_FIN","DIAS_SOLICITADOS",
                "MOTIVO","AUTORIZADO_POR","ESTATUS",
                "FECHA_PAGO_TIEMPO","HORA_INICIO_PAGO","HORA_FIN_PAGO",
                "OBSERVACIONES"],
      anchos:  [120,160,110,200,120,160,160,160,100,100,100,100,100,200,160,100,110,110,110,200]
    }
  ];

  var creadas = [], omitidas = [];

  tablas.forEach(function(t) {
    var sheet = ss.getSheetByName(t.nombre);
    var tieneDatos = sheet && sheet.getLastRow() > 1;

    if (tieneDatos) {
      // Ya tiene datos — solo agregar columnas faltantes sin borrar nada
      var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
                         .map(function(h){ return String(h).toUpperCase().trim(); });
      var agregadas = 0;
      t.headers.forEach(function(h) {
        if (headers.indexOf(h) === -1) {
          var newCol = sheet.getLastColumn() + 1;
          sheet.getRange(1, newCol).setValue(h)
               .setBackground("#1a237e").setFontColor("white")
               .setFontWeight("bold").setFontSize(10);
          sheet.setColumnWidth(newCol, 140);
          agregadas++;
        }
      });
      omitidas.push(t.nombre + (agregadas > 0 ? " (+" + agregadas + " cols)" : " (sin cambios)"));
      return;
    }

    // Sin datos → recrear limpia
    if (!sheet) sheet = ss.insertSheet(t.nombre);
    sheet.clearContents();
    sheet.getRange(1, 1, 1, t.headers.length).setValues([t.headers]);
    var hRange = sheet.getRange(1, 1, 1, t.headers.length);
    hRange.setBackground("#1a237e").setFontColor("white")
          .setFontWeight("bold").setFontSize(10);
    sheet.setFrozenRows(1);
    for (var c = 0; c < t.anchos.length; c++) {
      sheet.setColumnWidth(c + 1, t.anchos[c]);
    }
    creadas.push(t.nombre);
  });

  var msg = "✅ Proceso completado.\n";
  if (creadas.length)   msg += "Recreadas: " + creadas.join(", ")   + "\n";
  if (omitidas.length)  msg += "Con datos (solo se agregaron cols faltantes): " + omitidas.join(", ");
  Logger.log(msg);
  try { SpreadsheetApp.getUi().alert(msg); } catch(e) {}
}

// ── Registrar pago de tiempo en permiso existente ──────────────────────────
function registrarPagoTiempo(datos) {
  // datos: { folio, fechaPagoTiempo, horaInicioPago, horaFinPago }
  var ssRH  = SpreadsheetApp.openById(ID_HOJA_RH);
  var sheet = ssRH.getSheetByName("PERMISOS_EMPLEADOS");
  if (!sheet) return "Hoja no encontrada";
  var data    = sheet.getDataRange().getValues();
  var headers = data[0].map(function(h){ return String(h).toUpperCase().trim(); });
  var iFolio  = _colIdx(headers, "FOLIO"); if (iFolio < 0) iFolio = 1;
  var iEst    = _colIdx(headers, "ESTATUS");
  var iFPago  = _colIdx(headers, "FECHA_PAGO_TIEMPO");
  var iHIni   = _colIdx(headers, "HORA_INICIO_PAGO");
  var iHFin   = _colIdx(headers, "HORA_FIN_PAGO");
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][iFolio] || "") !== datos.folio) continue;
    var row = i + 1;
    if (iFPago > -1) sheet.getRange(row, iFPago + 1).setValue(datos.fechaPagoTiempo || "");
    if (iHIni  > -1) sheet.getRange(row, iHIni  + 1).setValue(datos.horaInicioPago  || "");
    if (iHFin  > -1) sheet.getRange(row, iHFin  + 1).setValue(datos.horaFinPago     || "");
    if (iEst   > -1) sheet.getRange(row, iEst   + 1).setValue("PAGADO");
    return "OK";
  }
  return "Folio no encontrado: " + datos.folio;
}

// =================================================================================
// ══════════════════  NÓMINA  ═════════════════════════════════════════════════════
// Libro: RecursosHumanos (ID_HOJA_RH)
// Tablas: NOMINA_SEMANAL | NOMINA_QUINCENAL | TIEMPO_EXTRA
//
// NOMINA_SEMANAL / NOMINA_QUINCENAL (13 columnas):
//   ID | FECHA | SEMANA | MES | CLAVE | NOMBRE | AREA
//   | SAL_DIA | DIAS_LAB | TIEMPO_EXTRA | FISCAL | CAFI | NETO
//
// TIEMPO_EXTRA (11 columnas):
//   ID | FECHA | SEMANA | MES | CLAVE | NOMBRE | AREA
//   | SAL_DIA | HORAS_DOBLES | HORAS_TRIPLES | IMPORTE_TOTAL
// =================================================================================

// ── Helper: buscar área de empleado por clave en OPERADORES ───────────
function _buscarAreaEmpleado(clave) {
  if (!clave || String(clave).trim() === "") return "SIN AREA";
  var ss    = SpreadsheetApp.openById(ID_HOJA_CALCULO);
  var shOp  = ss.getSheetByName("OPERADORES");
  if (!shOp) return "SIN AREA";
  var data  = shOp.getDataRange().getValues();
  var hOp   = data[0].map(function(h){ return String(h).toUpperCase().trim(); });
  var iClave= hOp.indexOf("CLAVE"); if (iClave < 0) iClave = 6;
  var iArea = hOp.indexOf("AREA");  if (iArea  < 0) iArea  = 4;
  var claveStr = String(clave).trim().toUpperCase();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][iClave] || "").trim().toUpperCase() === claveStr) {
      return String(data[i][iArea] || "SIN AREA").trim() || "SIN AREA";
    }
  }
  return "SIN AREA";
}

// ── Helper: parsear fecha desde string o número de Excel ─────────────
function _parseFechaNomina(v) {
  if (!v || v === "") return "";
  if (v instanceof Date) {
    try { return Utilities.formatDate(v, "GMT-6", "dd/MM/yyyy"); } catch(e) {}
    return ("0"+v.getDate()).slice(-2)+"/"+("0"+(v.getMonth()+1)).slice(-2)+"/"+v.getFullYear();
  }
  var s = String(v).trim();
  // Número serial de Excel (días desde 1900-01-01)
  var n = Number(s);
  if (!isNaN(n) && n > 40000 && n < 60000) {
    var d = new Date((n - 25569) * 86400000);
    return ("0"+d.getUTCDate()).slice(-2)+"/"+("0"+(d.getUTCMonth()+1)).slice(-2)+"/"+d.getUTCFullYear();
  }
  if (s.indexOf("GMT") > -1 || s.indexOf("T00:00") > -1) {
    var dd = new Date(s);
    if (!isNaN(dd.getTime())) {
      try { return Utilities.formatDate(dd, "GMT-6", "dd/MM/yyyy"); } catch(e) {}
    }
  }
  return s;
}

// ══════════════════════════════════════════════════════════════════════
//  CARGA MASIVA — NÓMINA SEMANAL / QUINCENAL
//  Recibe: { tipo: "SEMANAL"|"QUINCENAL", filas: [{FECHA,SEMANA,MES,CLAVE,NOMBRE,SAL_DIA,...}] }
// ══════════════════════════════════════════════════════════════════════
function cargarNominaMasiva(payload) {
  var tipo  = (payload.tipo  || "SEMANAL").toUpperCase();
  var filas = payload.filas  || [];
  if (!filas.length) return { ok: 0, errores: 0, msg: "Sin filas para cargar." };

  var nombreHoja = tipo === "QUINCENAL" ? "NOMINA_QUINCENAL" : "NOMINA_SEMANAL";
  var ssRH  = SpreadsheetApp.openById(ID_HOJA_RH);
  var sheet = ssRH.getSheetByName(nombreHoja);
  if (!sheet) return { ok: 0, errores: 0, msg: "Error: Hoja " + nombreHoja + " no existe. Ejecuta recrearTablasNomina()" };

  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
                     .map(function(h){ return String(h).toUpperCase().trim(); });
  var col = function(n) { return headers.indexOf(n.toUpperCase()); };
  var numCols = headers.length;

  var ok = 0, errores = 0;
  var nuevasFilas = [];

  filas.forEach(function(r) {
    try {
      var clave = String(r.CLAVE || r.clave || "").trim();
      var nombre= String(r.NOMBRE|| r.nombre|| "").trim();
      if (!clave && !nombre) { errores++; return; }
      var area  = _buscarAreaEmpleado(clave);
      var fila  = []; for (var c = 0; c < numCols; c++) fila.push("");
      var set = function(cn, val) { var i = col(cn); if (i > -1) fila[i] = val; };
      set("ID",           Utilities.getUuid());
      set("FECHA",        _parseFechaNomina(r.FECHA  || r.fecha  || ""));
      set("SEMANA",       r.SEMANA  || r.semana  || "");
      set("MES",          r.MES     || r.mes      || "");
      set("CLAVE",        clave);
      set("NOMBRE",       nombre);
      set("AREA",         area);
      set("SAL_DIA",      r.SAL_DIA || r.sal_dia  || 0);
      set("DIAS_LAB",     r.DIAS_LAB|| r.dias_lab || 0);
      set("TIEMPO_EXTRA", r.TIEMPO_EXTRA || r.tiempo_extra || 0);
      set("FISCAL",       r.FISCAL  || r.fiscal   || 0);
      set("CAFI",         r.CAFI    || r.cafi      || 0);
      set("NETO",         r.NETO    || r.neto      || 0);
      nuevasFilas.push(fila);
      ok++;
    } catch(e) { errores++; }
  });

  if (nuevasFilas.length > 0) {
    var startRow = sheet.getLastRow() + 1;
    sheet.getRange(startRow, 1, nuevasFilas.length, numCols).setValues(nuevasFilas);
  }

  return { ok: ok, errores: errores, msg: "✅ " + ok + " filas cargadas" + (errores ? " | ⚠️ " + errores + " con error" : "") };
}

// ══════════════════════════════════════════════════════════════════════
//  CONSULTA — NÓMINA SEMANAL / QUINCENAL (con rango de fechas)
// ══════════════════════════════════════════════════════════════════════
function consultarNomina(payload) {
  // payload: { tipo:"SEMANAL"|"QUINCENAL", fechaInicio:"dd/MM/yyyy", fechaFin:"dd/MM/yyyy" }
  var tipo       = (payload.tipo || "SEMANAL").toUpperCase();
  var nombreHoja = tipo === "QUINCENAL" ? "NOMINA_QUINCENAL" : "NOMINA_SEMANAL";
  var ssRH  = SpreadsheetApp.openById(ID_HOJA_RH);
  var sheet = ssRH.getSheetByName(nombreHoja);
  if (!sheet || sheet.getLastRow() < 2) return [];
  var data    = sheet.getDataRange().getValues();
  var headers = data[0].map(function(h){ return String(h).toUpperCase().trim(); });
  var col = function(n) { var i = headers.indexOf(n); return i; };

  // Parsear rango de fechas del filtro
  var parseFiltro = function(s) {
    if (!s) return null;
    var p = s.split("/");
    if (p.length === 3) return new Date(Number(p[2]), Number(p[1])-1, Number(p[0]));
    return null;
  };
  var dIni = parseFiltro(payload.fechaInicio);
  var dFin = parseFiltro(payload.fechaFin);

  var lista = [];
  for (var i = 1; i < data.length; i++) {
    var _g = function(cn) { var idx = col(cn); return idx > -1 ? data[i][idx] : ""; };
    // Filtrar por fecha si se especificó rango
    if (dIni || dFin) {
      var fStr = String(_g("FECHA") || "");
      var p2   = fStr.split("/");
      if (p2.length === 3) {
        var fRow = new Date(Number(p2[2]), Number(p2[1])-1, Number(p2[0]));
        if (dIni && fRow < dIni) continue;
        if (dFin && fRow > dFin) continue;
      }
    }
    lista.push({
      rowIndex:    i + 1, // fila real en la hoja (para edición)
      id:          String(_g("ID")           || ""),
      fecha:       _parseFechaNomina(_g("FECHA")),
      semana:      String(_g("SEMANA")       || ""),
      mes:         String(_g("MES")          || ""),
      clave:       String(_g("CLAVE")        || ""),
      nombre:      String(_g("NOMBRE")       || ""),
      area:        String(_g("AREA")         || ""),
      salDia:      Number(_g("SAL_DIA")      || 0),
      diasLab:     Number(_g("DIAS_LAB")     || 0),
      tiempoExtra: Number(_g("TIEMPO_EXTRA") || 0),
      fiscal:      Number(_g("FISCAL")       || 0),
      cafi:        Number(_g("CAFI")         || 0),
      neto:        Number(_g("NETO")         || 0)
    });
  }
  return lista;
}

// ══════════════════════════════════════════════════════════════════════
//  EDICIÓN DE FILA EN NÓMINA (actualiza una fila por rowIndex)
// ══════════════════════════════════════════════════════════════════════
function editarFilaNomina(payload) {
  // payload: { tipo, rowIndex, campos:{fecha,semana,mes,clave,nombre,area,salDia,...} }
  var tipo       = (payload.tipo || "SEMANAL").toUpperCase();
  var nombreHoja = tipo === "QUINCENAL" ? "NOMINA_QUINCENAL" : "NOMINA_SEMANAL";
  var ssRH  = SpreadsheetApp.openById(ID_HOJA_RH);
  var sheet = ssRH.getSheetByName(nombreHoja);
  if (!sheet) return "Hoja no encontrada";
  var row = Number(payload.rowIndex);
  if (!row || row < 2) return "Fila inválida";
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
                     .map(function(h){ return String(h).toUpperCase().trim(); });
  var set = function(cn, val) {
    var i = headers.indexOf(cn.toUpperCase());
    if (i > -1) sheet.getRange(row, i + 1).setValue(val);
  };
  var c = payload.campos || {};
  set("FECHA",        c.fecha        || "");
  set("SEMANA",       c.semana       || "");
  set("MES",          c.mes          || "");
  set("CLAVE",        c.clave        || "");
  set("NOMBRE",       c.nombre       || "");
  set("AREA",         c.area         || "");
  set("SAL_DIA",      Number(c.salDia      || 0));
  set("DIAS_LAB",     Number(c.diasLab     || 0));
  set("TIEMPO_EXTRA", Number(c.tiempoExtra || 0));
  set("FISCAL",       Number(c.fiscal      || 0));
  set("CAFI",         Number(c.cafi        || 0));
  set("NETO",         Number(c.neto        || 0));
  return "OK";
}

// ══════════════════════════════════════════════════════════════════════
//  ELIMINAR FILA EN NÓMINA
// ══════════════════════════════════════════════════════════════════════
function eliminarFilaNomina(payload) {
  var tipo       = (payload.tipo || "SEMANAL").toUpperCase();
  var nombreHoja = tipo === "QUINCENAL" ? "NOMINA_QUINCENAL" : "NOMINA_SEMANAL";
  var ssRH  = SpreadsheetApp.openById(ID_HOJA_RH);
  var sheet = ssRH.getSheetByName(nombreHoja);
  if (!sheet) return "Hoja no encontrada";
  var row = Number(payload.rowIndex);
  if (!row || row < 2) return "Fila inválida";
  sheet.deleteRow(row);
  return "OK";
}

// ══════════════════════════════════════════════════════════════════════
//  CARGA MASIVA — TIEMPO EXTRA
// ══════════════════════════════════════════════════════════════════════
function cargarTiempoExtraMasivo(filas) {
  if (!filas || !filas.length) return { ok: 0, errores: 0, msg: "Sin filas para cargar." };
  var ssRH  = SpreadsheetApp.openById(ID_HOJA_RH);
  var sheet = ssRH.getSheetByName("TIEMPO_EXTRA");
  if (!sheet) return { ok: 0, errores: 0, msg: "Error: Hoja TIEMPO_EXTRA no existe. Ejecuta recrearTablasNomina()" };
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
                     .map(function(h){ return String(h).toUpperCase().trim(); });
  var col = function(n) { return headers.indexOf(n.toUpperCase()); };
  var numCols = headers.length;
  var ok = 0, errores = 0;
  var nuevasFilas = [];

  filas.forEach(function(r) {
    try {
      var clave = String(r.CLAVE || r.clave || "").trim();
      var nombre= String(r.NOMBRE|| r.nombre|| "").trim();
      if (!clave && !nombre) { errores++; return; }
      var area  = _buscarAreaEmpleado(clave);
      var fila  = []; for (var c = 0; c < numCols; c++) fila.push("");
      var set = function(cn, val) { var i = col(cn); if (i > -1) fila[i] = val; };
      var hD   = Number(r.HORAS_DOBLES   || r.horas_dobles   || 0);
      var hT   = Number(r.HORAS_TRIPLES  || r.horas_triples  || 0);
      var sDay = Number(r.SAL_DIA || r.sal_dia || 0);
      // Importe = (salDia/8)*2*hDobles + (salDia/8)*3*hTriples
      var importe = sDay > 0 ? ((sDay / 8) * 2 * hD) + ((sDay / 8) * 3 * hT) : 0;
      set("ID",             Utilities.getUuid());
      set("FECHA",          _parseFechaNomina(r.FECHA  || r.fecha  || ""));
      set("SEMANA",         r.SEMANA  || r.semana  || "");
      set("MES",            r.MES     || r.mes      || "");
      set("CLAVE",          clave);
      set("NOMBRE",         nombre);
      set("AREA",           area);
      set("SAL_DIA",        sDay);
      set("HORAS_DOBLES",   hD);
      set("HORAS_TRIPLES",  hT);
      set("IMPORTE_TOTAL",  Math.round(importe * 100) / 100);
      nuevasFilas.push(fila);
      ok++;
    } catch(e) { errores++; }
  });

  if (nuevasFilas.length > 0) {
    var startRow = sheet.getLastRow() + 1;
    sheet.getRange(startRow, 1, nuevasFilas.length, numCols).setValues(nuevasFilas);
  }

  return { ok: ok, errores: errores, msg: "✅ " + ok + " filas cargadas" + (errores ? " | ⚠️ " + errores + " con error" : "") };
}

// ══════════════════════════════════════════════════════════════════════
//  CONSULTA — TIEMPO EXTRA
// ══════════════════════════════════════════════════════════════════════
function consultarTiempoExtra(payload) {
  var ssRH  = SpreadsheetApp.openById(ID_HOJA_RH);
  var sheet = ssRH.getSheetByName("TIEMPO_EXTRA");
  if (!sheet || sheet.getLastRow() < 2) return [];
  var data    = sheet.getDataRange().getValues();
  var headers = data[0].map(function(h){ return String(h).toUpperCase().trim(); });
  var col = function(n) { return headers.indexOf(n); };
  var parseFiltro = function(s) {
    if (!s) return null;
    var p = s.split("/");
    if (p.length === 3) return new Date(Number(p[2]), Number(p[1])-1, Number(p[0]));
    return null;
  };
  var dIni = parseFiltro(payload && payload.fechaInicio);
  var dFin = parseFiltro(payload && payload.fechaFin);
  var lista = [];
  for (var i = 1; i < data.length; i++) {
    var _g = function(cn) { var idx = col(cn); return idx > -1 ? data[i][idx] : ""; };
    if (dIni || dFin) {
      var fStr = String(_g("FECHA") || "");
      var p2   = fStr.split("/");
      if (p2.length === 3) {
        var fRow = new Date(Number(p2[2]), Number(p2[1])-1, Number(p2[0]));
        if (dIni && fRow < dIni) continue;
        if (dFin && fRow > dFin) continue;
      }
    }
    lista.push({
      rowIndex:      i + 1,
      id:            String(_g("ID")             || ""),
      fecha:         _parseFechaNomina(_g("FECHA")),
      semana:        String(_g("SEMANA")         || ""),
      mes:           String(_g("MES")            || ""),
      clave:         String(_g("CLAVE")          || ""),
      nombre:        String(_g("NOMBRE")         || ""),
      area:          String(_g("AREA")           || ""),
      salDia:        Number(_g("SAL_DIA")        || 0),
      horasDobles:   Number(_g("HORAS_DOBLES")   || 0),
      horasTriples:  Number(_g("HORAS_TRIPLES")  || 0),
      importeTotal:  Number(_g("IMPORTE_TOTAL")  || 0)
    });
  }
  return lista;
}

// ══════════════════════════════════════════════════════════════════════
//  EDICIÓN DE FILA EN TIEMPO EXTRA
// ══════════════════════════════════════════════════════════════════════
function editarFilaTiempoExtra(payload) {
  var ssRH  = SpreadsheetApp.openById(ID_HOJA_RH);
  var sheet = ssRH.getSheetByName("TIEMPO_EXTRA");
  if (!sheet) return "Hoja no encontrada";
  var row = Number(payload.rowIndex);
  if (!row || row < 2) return "Fila inválida";
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
                     .map(function(h){ return String(h).toUpperCase().trim(); });
  var set = function(cn, val) {
    var i = headers.indexOf(cn.toUpperCase());
    if (i > -1) sheet.getRange(row, i + 1).setValue(val);
  };
  var c = payload.campos || {};
  var hD  = Number(c.horasDobles  || 0);
  var hT  = Number(c.horasTriples || 0);
  var sD  = Number(c.salDia       || 0);
  var imp = sD > 0 ? ((sD/8)*2*hD) + ((sD/8)*3*hT) : Number(c.importeTotal || 0);
  set("FECHA",         c.fecha       || "");
  set("SEMANA",        c.semana      || "");
  set("MES",           c.mes         || "");
  set("CLAVE",         c.clave       || "");
  set("NOMBRE",        c.nombre      || "");
  set("AREA",          c.area        || "");
  set("SAL_DIA",       sD);
  set("HORAS_DOBLES",  hD);
  set("HORAS_TRIPLES", hT);
  set("IMPORTE_TOTAL", Math.round(imp * 100) / 100);
  return "OK";
}

function eliminarFilaTiempoExtra(rowIndex) {
  var ssRH  = SpreadsheetApp.openById(ID_HOJA_RH);
  var sheet = ssRH.getSheetByName("TIEMPO_EXTRA");
  if (!sheet) return "Hoja no encontrada";
  var row = Number(rowIndex);
  if (!row || row < 2) return "Fila inválida";
  sheet.deleteRow(row);
  return "OK";
}

// =================================================================================
// ══  UTILIDAD: Recrear tablas de Nómina  ══════════════════════════════════════════
// Ejecuta recrearTablasNomina() UNA VEZ.
// Sin datos → crea limpia. Con datos → solo agrega columnas faltantes.
// =================================================================================
function recrearTablasNomina() {
  var ss = SpreadsheetApp.openById(ID_HOJA_RH);

  var tablas = [
    {
      nombre:  "NOMINA_SEMANAL",
      headers: ["ID","FECHA","SEMANA","MES","CLAVE","NOMBRE","AREA",
                "SAL_DIA","DIAS_LAB","TIEMPO_EXTRA","FISCAL","CAFI","NETO"],
      anchos:  [120,100,80,80,100,200,160,100,90,110,100,100,100]
    },
    {
      nombre:  "NOMINA_QUINCENAL",
      headers: ["ID","FECHA","SEMANA","MES","CLAVE","NOMBRE","AREA",
                "SAL_DIA","DIAS_LAB","TIEMPO_EXTRA","FISCAL","CAFI","NETO"],
      anchos:  [120,100,80,80,100,200,160,100,90,110,100,100,100]
    },
    {
      nombre:  "TIEMPO_EXTRA",
      headers: ["ID","FECHA","SEMANA","MES","CLAVE","NOMBRE","AREA",
                "SAL_DIA","HORAS_DOBLES","HORAS_TRIPLES","IMPORTE_TOTAL"],
      anchos:  [120,100,80,80,100,200,160,100,110,110,120]
    }
  ];

  var creadas = [], conDatos = [];

  tablas.forEach(function(t) {
    var sheet     = ss.getSheetByName(t.nombre);
    var hayDatos  = sheet && sheet.getLastRow() > 1;

    if (hayDatos) {
      var hdrs = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0]
                      .map(function(h){ return String(h).toUpperCase().trim(); });
      var agg = 0;
      t.headers.forEach(function(h) {
        if (hdrs.indexOf(h) === -1) {
          var nc = sheet.getLastColumn() + 1;
          sheet.getRange(1, nc).setValue(h)
               .setBackground("#1a237e").setFontColor("white")
               .setFontWeight("bold").setFontSize(10);
          sheet.setColumnWidth(nc, 120);
          agg++;
        }
      });
      conDatos.push(t.nombre + (agg > 0 ? " (+" + agg + " cols)" : " (sin cambios)"));
      return;
    }

    if (!sheet) sheet = ss.insertSheet(t.nombre);
    sheet.clearContents();
    sheet.getRange(1, 1, 1, t.headers.length).setValues([t.headers]);
    var hr = sheet.getRange(1, 1, 1, t.headers.length);
    hr.setBackground("#1a237e").setFontColor("white").setFontWeight("bold").setFontSize(10);
    sheet.setFrozenRows(1);
    for (var c = 0; c < t.anchos.length; c++) sheet.setColumnWidth(c+1, t.anchos[c]);
    creadas.push(t.nombre + " (" + t.headers.length + " cols)");
  });

  var msg = "✅ recrearTablasNomina completado\n";
  if (creadas.length)  msg += "Creadas: "   + creadas.join(", ")  + "\n";
  if (conDatos.length) msg += "Con datos: " + conDatos.join(", ") + "\n";
  Logger.log(msg);
  try { SpreadsheetApp.getUi().alert(msg); } catch(e) {}
}
// ══════════════════════════════════════════════════════════════════════
//  TABULADOR DE SALARIOS
// ══════════════════════════════════════════════════════════════════════
function obtenerTabulador() {
  var ss    = SpreadsheetApp.openById(ID_HOJA_RH);
  var sheet = ss.getSheetByName('TABULADOR');
  if (!sheet) return { categorias: [], datos: [] };

  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return { categorias: [], datos: [] };

  // Fila 1 (índice 0) = encabezados. Columnas E-N = índices 4-13
  var headers = data[0];
  var categorias = [];
  for (var c = 4; c <= 13; c++) {
    var h = String(headers[c] || '').trim();
    if (h) categorias.push(h);
  }

  function _fmtSal(v) {
    if (typeof v === 'number') {
      var n = Math.round(v);
      return '$' + n.toString().replace(/\B(?=(\d{3})+(?!\d))/g, ',');
    }
    return String(v || '').trim();
  }

  var datos = [];
  for (var ci = 0; ci < categorias.length; ci++) {
    var colIdx  = ci + 4;
    var escalones = [];
    for (var r = 1; r < data.length; r++) {
      var celda = String(data[r][colIdx] || '').trim();
      if (!celda) continue;
      var salario = _fmtSal(data[r][3]);
      var nombres = celda.split(',').map(function(x){ return x.trim(); }).filter(Boolean);
      if (!nombres.length) continue;
      var items   = nombres.map(function(n){ return { nombre: n, salario: salario }; });
      var esPlus  = celda.indexOf('+') !== -1;
      if (!esPlus) {
        // Nuevo escalón base
        escalones.push({ base: items, plus: [] });
      } else {
        if (escalones.length > 0) {
          // Agregar al último escalón base como plus
          escalones[escalones.length - 1].plus =
            escalones[escalones.length - 1].plus.concat(items);
        } else {
          // Sin base previa, crear uno con los items como base
          escalones.push({ base: items, plus: [] });
        }
      }
    }
    datos.push({ categoria: categorias[ci], escalones: escalones });
  }

  return { categorias: categorias, datos: datos };
}

// ── Leer filas editables del Tabulador (Col A ID, B DIARIO, C DIA_INT, D SEMANAL) ──
function obtenerFilasTabulador() {
  var ss    = SpreadsheetApp.openById(ID_HOJA_RH);
  var sheet = ss.getSheetByName('TABULADOR');
  if (!sheet || sheet.getLastRow() < 2) return [];
  var data    = sheet.getDataRange().getValues();
  var headers = data[0].map(function(h){ return String(h).toUpperCase().trim(); });
  var iID  = headers.indexOf('ID');      if(iID  < 0) iID  = 0;
  var iD   = headers.indexOf('DIARIO');  if(iD   < 0) iD   = 1;
  var iDI  = headers.indexOf('DIA_INT'); if(iDI  < 0) iDI  = 2;
  var iS   = headers.indexOf('SEMANAL'); if(iS   < 0) iS   = 3;
  function _limpiar(v){
    if(typeof v === 'number') return v;
    return parseFloat(String(v).replace(/[$,\s]/g,'')) || 0;
  }
  function _fmt(v){
    var n = _limpiar(v);
    return n > 0 ? '$'+n.toFixed(2).replace(/\B(?=(\d{3})+(?!\d))/g,',') : '';
  }
  var filas = [];
  for(var r = 1; r < data.length; r++){
    var id = String(data[r][iID] || '').trim();
    var diario  = _limpiar(data[r][iD]);
    var diaInt  = _limpiar(data[r][iDI]);
    var semanal = _limpiar(data[r][iS]);
    if(!diario && !semanal) continue; // fila vacía
    filas.push({ rowNum: r+1, id: id, diario: diario, diaInt: diaInt, semanal: semanal,
      diarioFmt: _fmt(data[r][iD]), diaIntFmt: _fmt(data[r][iDI]), semanalFmt: _fmt(data[r][iS]) });
  }
  return filas;
}

// ── Guardar filas editadas del Tabulador ──
function actualizarFilasTabulador(filas) {
  var ss    = SpreadsheetApp.openById(ID_HOJA_RH);
  var sheet = ss.getSheetByName('TABULADOR');
  if (!sheet) return 'Error: hoja TABULADOR no encontrada';
  var data    = sheet.getDataRange().getValues();
  var headers = data[0].map(function(h){ return String(h).toUpperCase().trim(); });
  var iD  = headers.indexOf('DIARIO');  if(iD  < 0) iD  = 1;
  var iDI = headers.indexOf('DIA_INT'); if(iDI < 0) iDI = 2;
  var iS  = headers.indexOf('SEMANAL'); if(iS  < 0) iS  = 3;
  var actualizados = 0;
  filas.forEach(function(f){
    var rowNum = Number(f.rowNum);
    if(rowNum < 2 || rowNum > sheet.getLastRow()) return;
    sheet.getRange(rowNum, iD  + 1).setValue(Number(f.diario)  || 0);
    sheet.getRange(rowNum, iDI + 1).setValue(Number(f.diaInt)  || 0);
    sheet.getRange(rowNum, iS  + 1).setValue(Number(f.semanal) || 0);
    actualizados++;
  });
  SpreadsheetApp.flush();
  return 'OK:'+actualizados;
}

// ══════════════════════════════════════════════════════════════════════
//  CAPACITACIÓN — Matriz de Habilidades
// ══════════════════════════════════════════════════════════════════════

var CAP_COLS = ['SEG','CAL','PRO','MAT','LUB','OPE','HER','OPT','COD'];
var CAP_VALS = [0, 25, 50, 75, 100]; // únicos valores permitidos

// ── Inicializar hoja CAPACITACION (crear si no existe, agregar puestos nuevos) ──
function inicializarCapacitacion() {
  var ss     = SpreadsheetApp.openById(ID_HOJA_RH);
  var sheetT = ss.getSheetByName('TABULADOR');
  if (!sheetT) return 'ERROR: No se encontró la hoja TABULADOR';

  var dataT   = sheetT.getDataRange().getValues();
  var headers = dataT[0];

  // Leer áreas (columnas E-N = índices 4-13)
  var areas = [];
  for (var c = 4; c <= 13; c++) {
    var h = String(headers[c] || '').trim();
    if (h) areas.push({ nombre: h, colIdx: c });
  }

  // Recolectar puestos base (sin +) por área
  var filas = []; // {area, puesto}
  areas.forEach(function(area) {
    for (var r = 1; r < dataT.length; r++) {
      var celda = String(dataT[r][area.colIdx] || '').trim();
      if (!celda || celda.indexOf('+') !== -1) continue;
      var puestos = celda.split(',').map(function(x){ return x.trim(); }).filter(Boolean);
      puestos.forEach(function(p) {
        var existe = filas.some(function(f){ return f.area === area.nombre && f.puesto === p; });
        if (!existe) filas.push({ area: area.nombre, puesto: p });
      });
    }
  });

  // Crear o abrir hoja CAPACITACION
  var sheetC = ss.getSheetByName('CAPACITACION');
  if (!sheetC) {
    sheetC = ss.insertSheet('CAPACITACION');
  }

  // Escribir encabezado si la hoja está vacía
  var existentes = {};
  var dataC = sheetC.getDataRange().getValues();
  var encabezado = ['AREA', 'PUESTO'].concat(CAP_COLS);

  if (dataC.length < 1 || String(dataC[0][0]).trim() !== 'AREA') {
    sheetC.clearContents();
    sheetC.appendRow(encabezado);
    dataC = sheetC.getDataRange().getValues();
  }

  // Indexar puestos existentes en la hoja
  for (var i = 1; i < dataC.length; i++) {
    var key = String(dataC[i][0]).trim() + '|' + String(dataC[i][1]).trim();
    existentes[key] = true;
  }

  // Agregar filas faltantes con 0 en todas las categorías
  var nuevos = 0;
  filas.forEach(function(f) {
    var key = f.area + '|' + f.puesto;
    if (!existentes[key]) {
      sheetC.appendRow([f.area, f.puesto, 0, 0, 0, 0, 0, 0, 0, 0, 0]);
      nuevos++;
    }
  });

  return 'OK: ' + nuevos + ' puestos agregados. Total: ' + filas.length;
}

// ── Leer datos de capacitación para el frontend ──
function obtenerCapacitacion() {
  var ss = SpreadsheetApp.openById(ID_HOJA_RH);

  // Primero asegurar que la hoja exista y tenga todos los puestos del tabulador
  inicializarCapacitacion();

  var sheetC = ss.getSheetByName('CAPACITACION');
  if (!sheetC) return { areas: [], datos: {} };

  var data = sheetC.getDataRange().getValues();
  if (data.length < 2) return { areas: [], datos: {} };

  // Leer encabezado de categorías (columnas C en adelante = índice 2+)
  var cats = [];
  for (var c = 2; c < data[0].length; c++) {
    var h = String(data[0][c] || '').trim();
    if (h) cats.push(h);
  }

  // Agrupar por área
  var areas = [];
  var datos = {}; // { 'AREA': [ {puesto, SEG:0, CAL:25, ...}, ... ] }

  for (var r = 1; r < data.length; r++) {
    var area   = String(data[r][0] || '').trim();
    var puesto = String(data[r][1] || '').trim();
    if (!area || !puesto) continue;

    if (areas.indexOf(area) === -1) areas.push(area);
    if (!datos[area]) datos[area] = [];

    var entry = { puesto: puesto, salSem: '' };
    cats.forEach(function(cat, ci) {
      var raw = data[r][ci + 2];
      var val = parseInt(raw, 10);
      entry[cat] = isNaN(val) ? 0 : val;
    });
    datos[area].push(entry);
  }

  // ── Cruzar salario semanal desde hoja TABULADOR ───────────────────
  var sheetT = ss.getSheetByName('TABULADOR');
  if (sheetT) {
    var dataT = sheetT.getDataRange().getValues();
    function _fmtSal(v) {
      if (typeof v === 'number') {
        var n = Math.round(v);
        return '$' + n.toString().replace(/\B(?=(\d{3})+(?!\d))/g, ',');
      }
      return String(v || '').trim();
    }
    var salMap = {};
    for (var tr = 1; tr < dataT.length; tr++) {
      var salVal = dataT[tr][3]; // columna D
      if (!salVal) continue;
      for (var tc = 4; tc <= 13; tc++) {
        var celda = String(dataT[tr][tc] || '').trim();
        if (!celda || celda.indexOf('+') !== -1) continue;
        var partes = celda.split(',');
        for (var pi = 0; pi < partes.length; pi++) {
          var nombre = partes[pi].trim();
          if (nombre) salMap[nombre.toUpperCase()] = _fmtSal(salVal);
        }
      }
    }
    var areaKeys = Object.keys(datos);
    for (var ai = 0; ai < areaKeys.length; ai++) {
      var filas = datos[areaKeys[ai]];
      for (var fi = 0; fi < filas.length; fi++) {
        filas[fi].salSem = salMap[filas[fi].puesto.toUpperCase()] || '—';
      }
    }
  }

  return { areas: areas, cats: cats, datos: datos };
}

// ── Guardar un cambio de nivel ──
function guardarCapacitacion(area, puesto, cat, valor) {
  var ss     = SpreadsheetApp.openById(ID_HOJA_RH);
  var sheetC = ss.getSheetByName('CAPACITACION');
  if (!sheetC) return 'ERROR: hoja no encontrada';

  // Validar valor
  var valNum = Number(valor);
  var valido = false;
  for (var vi = 0; vi < CAP_VALS.length; vi++) {
    if (CAP_VALS[vi] === valNum) { valido = true; break; }
  }
  if (!valido) return 'ERROR: valor inválido';

  var data    = sheetC.getDataRange().getValues();
  var headers = data[0];

  // Encontrar columna de la categoría
  var colIdx = -1;
  for (var c = 2; c < headers.length; c++) {
    if (String(headers[c]).trim() === cat) { colIdx = c + 1; break; } // +1 para 1-based
  }
  if (colIdx === -1) return 'ERROR: categoría no encontrada';

  // Encontrar fila del puesto en el área
  for (var r = 1; r < data.length; r++) {
    if (String(data[r][0]).trim() === area && String(data[r][1]).trim() === puesto) {
      sheetC.getRange(r + 1, colIdx).setValue(valNum);
      return 'OK';
    }
  }
  return 'ERROR: puesto no encontrado';
}

// ══════════════════════════════════════════════════════════════════════
//  PERFIL DE PUESTOS
// ══════════════════════════════════════════════════════════════════════

var PERFIL_CATS = ['SEG','CAL','PRO','MAT','LUB','OPE','HER','OPT','COD'];

// ── Inicializar hoja PERFILES_PUESTOS ─────────────────────────────────
function inicializarPerfilesPuestos() {
  var ss = SpreadsheetApp.openById(ID_HOJA_RH);
  var sheet = ss.getSheetByName('PERFILES_PUESTOS');
  if (!sheet) {
    sheet = ss.insertSheet('PERFILES_PUESTOS');
    var header = ['AREA','PUESTO','OBJETIVO_PUESTO','CAT','OBJETIVO_CAT','METODO','URL_ADJUNTO'];
    for (var t = 1; t <= 20; t++) header.push('TAREA_' + t);
    sheet.appendRow(header);
  }
  return 'OK';
}

// ── Obtener perfil de un puesto — crea filas si no existen ────────────
function obtenerPerfilPuesto(area, puesto) {
  try {
  inicializarPerfilesPuestos();
  var ss    = SpreadsheetApp.openById(ID_HOJA_RH);
  var sheet = ss.getSheetByName('PERFILES_PUESTOS');

  // ── 1. Leer salarios del TABULADOR ───────────────────────────────
  var salarios  = [];
  var salPuesto = '';
  var _fmt = function(v) {
    var n;
    if (typeof v === 'number') {
      n = Math.round(v);
    } else {
      var limpio = String(v || '').replace(/[$,\s]/g, '').replace(/\..*$/, '');
      n = parseInt(limpio, 10);
    }
    if (isNaN(n) || n <= 0) return '';
    return '$' + n.toString().replace(/\B(?=(\d{3})+(?!\d))/g, ',');
  };
  var sheetT = ss.getSheetByName('TABULADOR');
  if (sheetT) {
    var dataT  = sheetT.getDataRange().getValues();
    var salSet = {};
    for (var tr = 1; tr < dataT.length; tr++) {
      var salV = dataT[tr][3];
      if (!salV) continue;
      var salStr = _fmt(salV);
      if (!salStr) continue;
      var salNum = typeof salV === 'number' ? Math.round(salV) : parseInt(String(salV).replace(/[$,\s]/g,'').replace(/\..*$/,''), 10);
      if (!salSet[salNum]) salSet[salNum] = salStr;
      for (var tc = 4; tc <= 13; tc++) {
        var celda = String(dataT[tr][tc] || '').trim();
        if (!celda || celda.indexOf('+') !== -1) continue;
        var partes = celda.split(',');
        for (var pi = 0; pi < partes.length; pi++) {
          if (partes[pi].trim().toUpperCase() === puesto.toUpperCase()) {
            salPuesto = salStr;
          }
        }
      }
    }
    var salNums = Object.keys(salSet).map(Number).filter(function(n){ return !isNaN(n) && n > 0; }).sort(function(a, b) { return a - b; });
    salarios = salNums.map(function(n) { return { num: n, str: salSet[n] }; });
  }

  // ── 2. Leer filas existentes en PERFILES_PUESTOS ─────────────────
  var data           = sheet.getDataRange().getValues();
  var objetivoPuesto = '';
  var cats           = {};
  var encontrado     = false;

  for (var r = 1; r < data.length; r++) {
    if (String(data[r][0]).trim().toUpperCase() !== area.toUpperCase()) continue;
    if (String(data[r][1]).trim().toUpperCase() !== puesto.toUpperCase()) continue;
    encontrado = true;
    if (!objetivoPuesto) objetivoPuesto = String(data[r][2] || '').trim();
    var cat = String(data[r][3] || '').trim();
    if (!cat) continue;
    var tareas = [];
    for (var tc2 = 7; tc2 < data[r].length; tc2++) {
      var tv = String(data[r][tc2] || '').trim();
      if (tv) tareas.push(tv);
    }
    cats[cat] = {
      objetivo:    String(data[r][4] || '').trim(),
      metodo:      String(data[r][5] || '').trim(),
      urlAdjunto:  String(data[r][6] || '').trim(),
      tareas:      tareas
    };
  }

  // ── 3. Si no existe, crear una fila por categoría ────────────────
  if (!encontrado) {
    for (var ci = 0; ci < PERFIL_CATS.length; ci++) {
      var emptyRow = [area, puesto, '', PERFIL_CATS[ci], '', '', ''];
      for (var ti = 0; ti < 20; ti++) emptyRow.push('');
      sheet.appendRow(emptyRow);
    }
  }

  // ── 4. Asegurar que todas las categorías existan en el objeto ────
  for (var ci2 = 0; ci2 < PERFIL_CATS.length; ci2++) {
    if (!cats[PERFIL_CATS[ci2]]) {
      cats[PERFIL_CATS[ci2]] = { objetivo: '', metodo: '', tareas: [] };
    }
  }

  return {
    area:           area,
    puesto:         puesto,
    objetivoPuesto: objetivoPuesto,
    salarios:       salarios,
    salPuesto:      salPuesto,
    cats:           cats
  };
  } catch(e) {
    Logger.log('ERROR obtenerPerfilPuesto: ' + e.toString());
    throw e;
  }
}

// ── Guardar perfil (objetivo puesto + una categoría completa) ─────────
function guardarPerfilPuesto(area, puesto, objetivoPuesto, cat, objetivoCat, metodo, tareas, urlAdjunto) {
  inicializarPerfilesPuestos();
  var ss    = SpreadsheetApp.openById(ID_HOJA_RH);
  var sheet = ss.getSheetByName('PERFILES_PUESTOS');
  var data  = sheet.getDataRange().getValues();
  var head  = data[0].map(function(h){ return String(h).trim(); });

  for (var r = 1; r < data.length; r++) {
    if (String(data[r][0]).trim().toUpperCase() === area.toUpperCase() &&
        String(data[r][1]).trim().toUpperCase() === puesto.toUpperCase()) {
      sheet.getRange(r + 1, 3).setValue(objetivoPuesto);
    }
  }

  var filaExistente = -1;
  for (var r2 = 1; r2 < data.length; r2++) {
    if (String(data[r2][0]).trim().toUpperCase() === area.toUpperCase() &&
        String(data[r2][1]).trim().toUpperCase() === puesto.toUpperCase() &&
        String(data[r2][3]).trim() === cat) {
      filaExistente = r2 + 1;
      break;
    }
  }

  // Si no se pasa urlAdjunto nueva, preservar la existente
  var urlFinal = urlAdjunto || '';
  if (!urlFinal && filaExistente > 0) {
    var adjIdx = head.indexOf('URL_ADJUNTO');
    if (adjIdx >= 0) urlFinal = String(data[filaExistente-1][adjIdx]||'').trim();
  }

  var tareasArr = Array.isArray(tareas) ? tareas : [];
  var rowData   = [area, puesto, objetivoPuesto, cat, objetivoCat, metodo, urlFinal];
  for (var ti = 0; ti < 20; ti++) rowData.push(tareasArr[ti] || '');

  if (filaExistente > 0) {
    sheet.getRange(filaExistente, 1, 1, rowData.length).setValues([rowData]);
  } else {
    sheet.appendRow(rowData);
  }
  return 'OK';
}

// ── Guardar URL adjunto de una categoría del perfil ───────────────────
function guardarAdjuntoPerfil(area, puesto, cat, urlAdjunto) {
  inicializarPerfilesPuestos();
  var ss    = SpreadsheetApp.openById(ID_HOJA_RH);
  var sheet = ss.getSheetByName('PERFILES_PUESTOS');
  var data  = sheet.getDataRange().getValues();
  var head  = data[0].map(function(h){ return String(h).trim(); });
  var adjIdx = head.indexOf('URL_ADJUNTO');
  if (adjIdx < 0) return 'ERROR: columna URL_ADJUNTO no encontrada';
  for (var r = 1; r < data.length; r++) {
    if (String(data[r][0]).trim().toUpperCase() === area.toUpperCase() &&
        String(data[r][1]).trim().toUpperCase() === puesto.toUpperCase() &&
        String(data[r][3]).trim() === cat) {
      sheet.getRange(r + 1, adjIdx + 1).setValue(urlAdjunto);
      return 'OK';
    }
  }
  return 'ERROR: fila no encontrada';
}

// ── Obtener todos los perfiles para el módulo Perfil de Puestos ───────
function obtenerTodosPerfiles() {
  inicializarPerfilesPuestos();
  var ss    = SpreadsheetApp.openById(ID_HOJA_RH);
  var sheet = ss.getSheetByName('PERFILES_PUESTOS');
  var data  = sheet.getDataRange().getValues();

  var perfiles = {};
  for (var r = 1; r < data.length; r++) {
    var area   = String(data[r][0] || '').trim();
    var puesto = String(data[r][1] || '').trim();
    if (!area || !puesto) continue;
    var key = area + '|' + puesto;
    if (!perfiles[key]) perfiles[key] = { area: area, puesto: puesto, objetivo: String(data[r][2] || '').trim(), cats: {} };
    var cat = String(data[r][3] || '').trim();
    if (cat) {
      var tareas = [];
      for (var tc = 6; tc < data[r].length; tc++) {
        var tv = String(data[r][tc] || '').trim();
        if (tv) tareas.push(tv);
      }
      perfiles[key].cats[cat] = { objetivo: String(data[r][4] || '').trim(), metodo: String(data[r][5] || '').trim(), tareas: tareas };
    }
  }

  var areas = {};
  var keys = Object.keys(perfiles);
  for (var ki = 0; ki < keys.length; ki++) {
    var p = perfiles[keys[ki]];
    if (!areas[p.area]) areas[p.area] = [];
    areas[p.area].push(p);
  }
  return areas;
}

// =================================================================================
// ══════════════════  EVALUACIONES  ═══════════════════════════════════════════════
// =================================================================================
// Hoja EVALUACIONES en libro ID_HOJA_RH
// Columnas: ID | CLAVE | FECHA | AREA | PUESTO | NOMBRE | COMENTARIOS | ASCENSO | COMENTARIOS_EVALUADO | CATS_JSON

var EVAL_CATS_LIST = ['SEG','CAL','PRO','MAT','LUB','OPE','HER','OPT','COD'];

function _inicializarEvaluaciones() {
  var ss    = SpreadsheetApp.openById(ID_HOJA_RH);
  var sheet = ss.getSheetByName('EVALUACIONES');
  if (!sheet) {
    sheet = ss.insertSheet('EVALUACIONES');
    sheet.appendRow(['ID','CLAVE','FECHA','AREA','PUESTO','NOMBRE','COMENTARIOS','ASCENSO','COMENTARIOS_EVALUADO','CATS_JSON']);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// ── Plantilla: lista de empleados activos + su historial de evaluaciones ──
function obtenerPlantillaEvaluaciones() {
  // 1. Leer OPERADORES activos
  var ss    = SpreadsheetApp.openById(ID_HOJA_CALCULO);
  var shOps = ss.getSheetByName('OPERADORES');
  var dataOps = shOps.getDataRange().getValues();
  var hd   = dataOps[0].map(function(h){ return String(h).toUpperCase().trim(); });
  var iClave   = hd.indexOf('CLAVE');
  var iNombre  = hd.indexOf('NOMBRE');
  var iArea    = hd.indexOf('AREA');
  var iPuesto  = hd.indexOf('PUESTO');
  var iEstado  = hd.indexOf('ESTADO');
  var iIngreso = hd.indexOf('FECHA_INGRESO'); if (iIngreso === -1) iIngreso = hd.indexOf('INGRESO');
  var iFiscal  = hd.indexOf('SAL_DIA_FIS');   if (iFiscal  === -1) iFiscal  = hd.indexOf('FISCAL');
  var iInteg   = hd.indexOf('SAL_DIA_INT');   if (iInteg   === -1) iInteg   = hd.indexOf('INTEGRADO');

  var empMap = {}; // clave → emp object
  for (var i = 1; i < dataOps.length; i++) {
    var fila   = dataOps[i];
    var estado = String(fila[iEstado] || '').toUpperCase().trim();
    if (estado === 'BAJA') continue;
    var clave  = String(fila[iClave]  || '').trim(); if (!clave) continue;
    var salNum = Number(fila[iFiscal]) || Number(fila[iInteg]) || 0;
    var salStr = salNum > 0 ? '$' + salNum.toFixed(2).replace(/\B(?=(\d{3})+(?!\d))/g, ',') : '';
    var ingreso = '';
    var ingVal = fila[iIngreso];
    if (ingVal instanceof Date) {
      ingreso = Utilities.formatDate(ingVal, 'GMT-6', 'dd/MM/yyyy');
    } else if (ingVal) {
      ingreso = String(ingVal);
    }
    empMap[clave] = {
      clave:   clave,
      nombre:  String(fila[iNombre]  || '').trim(),
      area:    String(fila[iArea]    || '').trim(),
      puesto:  String(fila[iPuesto]  || '').trim(),
      salDia:  salStr,
      ingreso: ingreso,
      historial: []
    };
  }

  // 2. Leer historial de EVALUACIONES
  var shEval = _inicializarEvaluaciones();
  var dataEv = shEval.getDataRange().getValues();
  for (var e = 1; e < dataEv.length; e++) {
    var id      = String(dataEv[e][0] || '').trim();
    var claveEv = String(dataEv[e][1] || '').trim();
    var fecha   = String(dataEv[e][2] || '').trim();
    var ascenso = String(dataEv[e][7] || '').trim();
    if (!id || !claveEv) continue;
    if (empMap[claveEv]) {
      empMap[claveEv].historial.push({ id: id, fecha: fecha, ascenso: ascenso });
    }
  }

  // Ordenar historial más reciente primero (IDs son timestamp-based)
  var empleados = [];
  var claves = Object.keys(empMap);
  for (var ci = 0; ci < claves.length; ci++) {
    var emp = empMap[claves[ci]];
    emp.historial.sort(function(a, b){ return b.id > a.id ? 1 : -1; });
    empleados.push(emp);
  }
  empleados.sort(function(a, b){
    if (a.area < b.area) return -1;
    if (a.area > b.area) return 1;
    if (a.nombre < b.nombre) return -1;
    return a.nombre > b.nombre ? 1 : 0;
  });

  return { empleados: empleados, cats: EVAL_CATS_LIST };
}

// ── Detalle de una evaluación existente ───────────────────────────────
function obtenerDetalleEvaluacion(idEval) {
  var shEval = _inicializarEvaluaciones();
  var ssRH   = SpreadsheetApp.openById(ID_HOJA_RH);
  var shPerf = ssRH.getSheetByName('PERFILES_PUESTOS');
  var data   = shEval.getDataRange().getValues();
  var result = null;

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(idEval).trim()) {
      var catsJson = String(data[i][9] || '{}');
      var cats = {};
      try { cats = JSON.parse(catsJson); } catch(ex) { cats = {}; }
      result = {
        id:                   String(data[i][0]).trim(),
        clave:                String(data[i][1]).trim(),
        fecha:                String(data[i][2]).trim(),
        area:                 String(data[i][3]).trim(),
        puesto:               String(data[i][4]).trim(),
        nombre:               String(data[i][5]).trim(),
        comentarios:          String(data[i][6] || '').trim(),
        ascenso:              String(data[i][7] || '').trim(),
        comentariosEvaluado:  String(data[i][8] || '').trim(),
        cats:                 cats
      };
      break;
    }
  }

  if (!result) return null;

  // Agregar tareas y adjunto del perfil para saber qué evaluar
  if (shPerf && result.puesto) {
    var dataP = shPerf.getDataRange().getValues();
    var headP = dataP[0].map(function(h){ return String(h).trim(); });
    var adjIdxP = headP.indexOf('URL_ADJUNTO');
    EVAL_CATS_LIST.forEach(function(cat){
      if (!result.cats[cat]) result.cats[cat] = { evalFinal:'', urlAdjunto:'', tareas:[] };
      var tareasP = [];
      var urlP    = '';
      for (var rp = 1; rp < dataP.length; rp++) {
        if (String(dataP[rp][1]||'').trim().toUpperCase() !== result.puesto.toUpperCase()) continue;
        if (String(dataP[rp][3]||'').trim() !== cat) continue;
        if (adjIdxP >= 0) urlP = String(dataP[rp][adjIdxP]||'').trim();
        for (var tp = 7; tp < dataP[rp].length; tp++) {
          var tv = String(dataP[rp][tp]||'').trim();
          if (tv) tareasP.push(tv);
        }
        break;
      }
      result.cats[cat].tareasDesc = tareasP;
      result.cats[cat].urlAdjunto = result.cats[cat].urlAdjunto || urlP;
    });
  }

  return result;
}

// ── Guardar / actualizar evaluación ───────────────────────────────────
function guardarEvaluacion(evalObj) {
  var shEval = _inicializarEvaluaciones();
  var data   = shEval.getDataRange().getValues();
  var id     = String(evalObj.id || ('EVAL-' + new Date().getTime())).trim();
  var cats   = evalObj.cats || {};
  var catsJson = JSON.stringify(cats);
  var row = [
    id,
    String(evalObj.clave    || '').trim(),
    String(evalObj.fecha    || '').trim(),
    String(evalObj.area     || '').trim(),
    String(evalObj.puesto   || '').trim(),
    String(evalObj.nombre   || '').trim(),
    String(evalObj.comentarios          || '').trim(),
    String(evalObj.ascenso              || '').trim(),
    String(evalObj.comentariosEvaluado  || '').trim(),
    catsJson
  ];

  // Buscar fila existente
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === id) {
      shEval.getRange(i + 1, 1, 1, row.length).setValues([row]);
      return id;
    }
  }
  // Nueva fila
  shEval.appendRow(row);
  return id;
}

// ── Subir adjunto a Drive ─────────────────────────────────────────────
function subirAdjuntoEvaluacion(b64, filename, mimetype) {
  var ss      = SpreadsheetApp.openById(ID_HOJA_RH);
  var ssFname = ss.getName();
  var folders = DriveApp.getFoldersByName('RH_Evaluaciones_Adjuntos');
  var folder;
  if (folders.hasNext()) {
    folder = folders.next();
  } else {
    // Intentar guardar junto al archivo del spreadsheet
    var ssFile   = DriveApp.getFileById(ID_HOJA_RH);
    var parents  = ssFile.getParents();
    var parent   = parents.hasNext() ? parents.next() : DriveApp.getRootFolder();
    folder = parent.createFolder('RH_Evaluaciones_Adjuntos');
  }
  var decoded = Utilities.base64Decode(b64);
  var blob    = Utilities.newBlob(decoded, mimetype, filename);
  var file    = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return file.getUrl();
}

// ══════════════════════════════════════════════════════════
//  TELEGRAM: NOTIFICACIÓN DIARIA SOLICITUDES PENDIENTES
// ══════════════════════════════════════════════════════════

function enviarNotificacionTelegramPendientes() {
  var TELEGRAM_TOKEN = '7947767393:AAFmZUcSTnV5gvP6u_UsBcSHlz-0s9x1kSQ';
  var CHAT_ID = '625827165';
  var DEPLOY_URL = 'https://script.google.com/macros/s/AKfycbwVRAgJccVKU6TgYD3NN2KZjX1zwxwIYUuG9FiVTmJRPoW4SkIThAPbH5w408oBb73l/exec';

  var pendientes = obtenerSolicitudesEstado('SOLICITADO');
  var n = pendientes.length;

  if (n === 0) return;

  var urlAutorizar = DEPLOY_URL + '?modulo=AUTORIZAR';

  var texto = '🔔 *Solicitudes de Vacaciones Pendientes*\n\n'
    + 'Tienes *' + n + ' solicitud' + (n === 1 ? '' : 'es') + ' pendiente' + (n === 1 ? '' : 's') + '* por autorizar.\n\n'
    + 'Toca el botón para revisar y autorizar desde tu celular:';

  var payload = {
    chat_id: CHAT_ID,
    text: texto,
    parse_mode: 'Markdown',
    reply_markup: {
      inline_keyboard: [[
        {
          text: '✅ Autorizar Solicitudes (' + n + ')',
          url: urlAutorizar
        }
      ]]
    }
  };

  var options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  var url = 'https://api.telegram.org/bot' + TELEGRAM_TOKEN + '/sendMessage';
  var response = UrlFetchApp.fetch(url, options);
  var result = JSON.parse(response.getContentText());

  if (!result.ok) {
    console.error('Error Telegram: ' + JSON.stringify(result));
  } else {
    console.log('Notificación Telegram enviada. Pendientes: ' + n);
  }
}

function crearTriggerDiarioTelegram() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'enviarNotificacionTelegramPendientes') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  ScriptApp.newTrigger('enviarNotificacionTelegramPendientes')
    .timeBased()
    .atHour(14)
    .everyDays(1)
    .create();
  console.log('Trigger diario creado correctamente. Se ejecutará cada día a las 8 AM hora México.');
}

// =================================================================================
// ══════════════════  TRIGGER DIARIO — VACACIONES  ════════════════════════════════
// =================================================================================

/**
 * Crea un trigger que ejecuta actualizarSaldosYAniversarios() cada día entre 1 y 2 AM.
 * EJECUTAR SOLO UNA VEZ manualmente desde el editor de Apps Script.
 * Si ya existe un trigger con ese nombre lo elimina primero para evitar duplicados.
 */
function crearTriggerDiarioVacaciones() {
  // Eliminar triggers anteriores de esta misma función para no duplicar
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'actualizarSaldosYAniversarios') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  // Crear el trigger diario entre 1 y 2 AM
  ScriptApp.newTrigger('actualizarSaldosYAniversarios')
    .timeBased()
    .everyDays(1)
    .atHour(1)
    .create();
  Logger.log('Trigger diario de vacaciones creado correctamente.');
}

/**
 * Elimina el trigger diario de vacaciones (usar si necesitas desactivarlo).
 */
function eliminarTriggerDiarioVacaciones() {
  var triggers = ScriptApp.getProjectTriggers();
  var eliminados = 0;
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'actualizarSaldosYAniversarios') {
      ScriptApp.deleteTrigger(triggers[i]);
      eliminados++;
    }
  }
  Logger.log('Triggers eliminados: ' + eliminados);
}
