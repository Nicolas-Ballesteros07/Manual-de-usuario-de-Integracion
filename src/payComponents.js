import { dividirNotas } from "./dividirNotas";

/**
 * Reglas Concepto Excel -> Pay Component.
 *
 * Los códigos salen de la lista del manual (secciones 7.2 y 7.3).
 * Solo SUELDO (001) está confirmado (lo dijo Johana); el resto son
 * PROPUESTAS pendientes de validar con Ricardo. Para corregir una,
 * basta con cambiar el "codigo" de su línea o borrarla.
 *
 * fuente:
 *  - "salarioContrato": toma la columna "Salario del contrato" (una sola vez por empleado/mes)
 *  - "devengo": suma la columna "Devengo" de todas las filas del concepto en el mes
 *
 * Varios conceptos pueden compartir el mismo código: sus valores se suman
 * (o, con "salarioContrato", se toma el salario una sola vez).
 *
 * No se reportan (no hay código equivalente en el manual): deducciones,
 * prestaciones legales, incapacidades, licencias y los Sodexo repetidos
 * 020, 612 y 924.
 */
export const REGLAS_CONCEPTOS = {
  "001": { codigo: "9GBA", fuente: "salarioContrato" }, // SUELDO (confirmado)
  "002": { codigo: "9GBA", fuente: "salarioContrato" }, // SALARIO INTEGRAL (propuesto)
  "006": { codigo: "9ASA", fuente: "devengo" }, // COMISIONES SALARIO INTEGRAL (propuesto)
  "007": { codigo: "9ASA", fuente: "devengo" }, // COMISIONES (propuesto)
  "016": { codigo: "9OVI", fuente: "devengo" }, // PREMIO MBO -INTEGRAL (propuesto)
  "019": { codigo: "9OVI", fuente: "devengo" }, // BONIFICACION ANUAL X SALAR (propuesto)
  "021": { codigo: "9OVI", fuente: "devengo" }, // BONO SALARIAL MESES (propuesto)
  "027": { codigo: "9OVI", fuente: "devengo" }, // PREMIO MBO - SALARIO ORDINARIO (propuesto)
  "516": { codigo: "9OVI", fuente: "devengo" }, // PREMIO MERA (propuesto)
  "200": { codigo: "9ROT", fuente: "devengo" }, // ALIMENTACION (propuesto)
  "201": { codigo: "9MOA", fuente: "devengo" }, // MEDIOS DE TRANSPORTE CUOTA (propuesto)
  "202": { codigo: "9MOA", fuente: "devengo" }, // MEDIOS DE TRANSPORTE (propuesto)
  "923": { codigo: "SODEXO", fuente: "devengo" }, // SODEXO BONOS ALIMENTACION (propuesto)
};

/**
 * Orden de los códigos dentro de la cadena. Se fija a mano porque JavaScript
 * ordena primero las claves numéricas ("200", "923") y eso desordenaba la
 * cadena. Sigue el ejemplo del manual: 9GBA-.../SODEXO-.../9ROT-...
 */
const ORDEN_CODIGOS = ["9GBA", "SODEXO", "9ROT", "9ASA", "9OVI", "9MOA"];

const COLUMNAS_REQUERIDAS = [
  "Tercero",
  "Descripción",
  "Contrato",
  "Concepto",
  "Devengo",
  "Salario del contrato",
  "Fecha Movto.",
];

/* ---------- 1. Hoja leída con JSZip -> arreglo de objetos ---------- */

export function hojaAObjetos(hoja) {
  if (!hoja || hoja.filas.length < 2) {
    throw new Error("La hoja no tiene encabezado y datos.");
  }

  const [encabezado, ...datos] = hoja.filas;

  const columnas = Object.entries(encabezado.celdas).map(([ref, nombre]) => ({
    letra: ref.replace(/\d+/g, ""),
    nombre: String(nombre).trim(),
  }));

  const faltantes = COLUMNAS_REQUERIDAS.filter(
    (requerida) => !columnas.some((c) => c.nombre === requerida)
  );

  if (faltantes.length > 0) {
    throw new Error(`Faltan columnas en el Excel: ${faltantes.join(", ")}.`);
  }

  return datos.map((fila) => {
    const registro = {};

    for (const { letra, nombre } of columnas) {
      registro[nombre] = fila.celdas[`${letra}${fila.numeroFila}`] ?? "";
    }

    return registro;
  });
}

/* ---------- 2. Fechas: el XML guarda las fechas como número serial ---------- */

export function fechaAMes(valor) {
  const texto = String(valor ?? "").trim();

  if (/^\d{4}-\d{2}/.test(texto)) return texto.slice(0, 7);

  if (/^\d+(\.\d+)?$/.test(texto)) {
    const fecha = new Date(Date.UTC(1899, 11, 30) + Number(texto) * 86400000);
    return fecha.toISOString().slice(0, 7);
  }

  return "";
}

export function listarMeses(registros) {
  return Array.from(
    new Set(registros.map((r) => fechaAMes(r["Fecha Movto."])).filter(Boolean))
  ).sort();
}

/* ---------- 3. Construcción de la cadena por empleado/contrato ---------- */

export function construirPayComponents(
  registros,
  mes,
  reglas = REGLAS_CONCEPTOS
) {
  const grupos = new Map();

  for (const registro of registros) {
    if (fechaAMes(registro["Fecha Movto."]) !== mes) continue;

    const regla = reglas[String(registro["Concepto"]).trim()];
    if (!regla) continue;

    const clave = `${registro["Tercero"]}|${registro["Contrato"]}`;

    if (!grupos.has(clave)) {
      grupos.set(clave, {
        tercero: String(registro["Tercero"]).trim(),
        nombre: registro["Descripción"],
        contrato: String(registro["Contrato"]).trim(),
        valores: {},
        codigosSalario: new Set(),
      });
    }

    const grupo = grupos.get(clave);

    if (regla.fuente === "salarioContrato") {
      grupo.valores[regla.codigo] = Number(registro["Salario del contrato"]);
      grupo.codigosSalario.add(regla.codigo);
    } else {
      grupo.valores[regla.codigo] =
        (grupo.valores[regla.codigo] ?? 0) + Number(registro["Devengo"]);
    }
  }

  const empleados = [];
  const errores = [];
  const advertencias = [];

  // Sin repetidos y con el orden fijo de ORDEN_CODIGOS; un código de las
  // reglas que no esté en la lista va al final, para que nunca se pierda.
  const codigosEnReglas = Array.from(
    new Set(Object.values(reglas).map((r) => r.codigo))
  );
  const ordenCodigos = [
    ...ORDEN_CODIGOS.filter((c) => codigosEnReglas.includes(c)),
    ...codigosEnReglas.filter((c) => !ORDEN_CODIGOS.includes(c)),
  ];

  for (const grupo of grupos.values()) {
    const conceptos = ordenCodigos
      .filter((codigo) => Number.isFinite(grupo.valores[codigo]))
      .map((codigo) => [codigo, Math.round(grupo.valores[codigo])])
      .filter(([, valor]) => valor > 0)
      .map(([codigo, valor]) => `${codigo}-${valor}`);

    // Advertencia: el concepto de salario existe pero su valor es 0, vacío
    // o inválido, así que ese código no se reporta (no es un error: el
    // resto de la cadena sigue siendo válido).
    const codigosSinSalario = Array.from(grupo.codigosSalario).filter(
      (codigo) => !(grupo.valores[codigo] > 0)
    );

    for (const codigo of codigosSinSalario) {
      advertencias.push({
        tercero: grupo.tercero,
        nombre: grupo.nombre,
        contrato: grupo.contrato,
        codigo,
        enTXT: conceptos.length > 0,
        mensaje:
          conceptos.length > 0
            ? `Salario del contrato en 0: no se reporta ${codigo}.`
            : `Salario del contrato en 0: no se reporta ${codigo} y el empleado no aparece en el TXT.`,
      });
    }

    if (conceptos.length === 0) continue;

    const cadena = conceptos.join("/");

    try {
      empleados.push({
        tercero: grupo.tercero,
        nombre: grupo.nombre,
        contrato: grupo.contrato,
        cadena,
        campos: dividirNotas(cadena),
      });
    } catch (error) {
      errores.push({
        tercero: grupo.tercero,
        nombre: grupo.nombre,
        contrato: grupo.contrato,
        mensaje: error.message,
      });
    }
  }

  return { empleados, errores, advertencias };
}

/* ---------- 4. Filas listas para generarTXT ---------- */

export function filasParaTXT(empleados, { incluirEncabezado = true } = {}) {
  const encabezado = [
    "tercero",
    "contrato",
    "notas",
    "txt_adc1",
    "txt_adc2",
    "txt_adc3",
    "txt_adc4",
  ];

  const filas = empleados.map((e) => [
    e.tercero,
    e.contrato,
    e.campos.notas,
    e.campos.txt_adc1,
    e.campos.txt_adc2,
    e.campos.txt_adc3,
    e.campos.txt_adc4,
  ]);

  return incluirEncabezado ? [encabezado, ...filas] : filas;
}