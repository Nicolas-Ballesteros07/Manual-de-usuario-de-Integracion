import { dividirNotas } from "./dividirNotas";

/**
 * Reglas Concepto Excel -> Pay Component.
 * Por ahora solo SUELDO (001) está confirmado (lo dijo Johana).
 * Cuando Ricardo confirme los demás, se agregan aquí sin tocar el resto.
 *
 * fuente:
 *  - "salarioContrato": toma la columna "Salario del contrato" (una sola vez por empleado/mes)
 *  - "devengo": suma la columna "Devengo" de todas las filas del concepto en el mes
 */
export const REGLAS_CONCEPTOS = {
  "001": { codigo: "9GBA", fuente: "salarioContrato" },
  // "923": { codigo: "SODEXO", fuente: "devengo" },   // pendiente de confirmar
  // "007": { codigo: "9OVI", fuente: "devengo" },     // pendiente de confirmar
};

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
      });
    }

    const grupo = grupos.get(clave);

    if (regla.fuente === "salarioContrato") {
      grupo.valores[regla.codigo] = Number(registro["Salario del contrato"]);
    } else {
      grupo.valores[regla.codigo] =
        (grupo.valores[regla.codigo] ?? 0) + Number(registro["Devengo"]);
    }
  }

  const empleados = [];
  const errores = [];
  const ordenCodigos = Object.values(reglas).map((r) => r.codigo);

  for (const grupo of grupos.values()) {
    const conceptos = ordenCodigos
      .filter((codigo) => Number.isFinite(grupo.valores[codigo]))
      .map((codigo) => [codigo, Math.round(grupo.valores[codigo])])
      .filter(([, valor]) => valor > 0)
      .map(([codigo, valor]) => `${codigo}-${valor}`);

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

  return { empleados, errores };
}

/* ---------- 4. Filas listas para generarTXT ---------- */

export function filasParaTXT(empleados) {
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

  return [encabezado, ...filas];
}