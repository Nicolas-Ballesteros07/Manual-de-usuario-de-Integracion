
const CAPACIDADES = [
  { campo: "notas", maximo: 255 },
  { campo: "txt_adc1", maximo: 50 },
  { campo: "txt_adc2", maximo: 50 },
  { campo: "txt_adc3", maximo: 50 },
  { campo: "txt_adc4", maximo: 50 },
];

export function dividirNotas(cadena) {
  if (typeof cadena !== "string" || !cadena.trim()) {
    throw new Error("La cadena de Pay Components está vacía.");
  }

  const conceptos = cadena.split("/");

  if (conceptos.some((concepto) => !concepto.trim())) {
    throw new Error("Hay conceptos vacíos en la cadena.");
  }

  for (const concepto of conceptos) {
    if (!/^[A-Z0-9]+--?\d+$/.test(concepto)) {
      throw new Error(
        `El concepto "${concepto}" no cumple el formato CODIGO-VALOR.`
      );
    }
  }

  const resultado = {
    notas: "",
    txt_adc1: "",
    txt_adc2: "",
    txt_adc3: "",
    txt_adc4: "",
  };

  let indiceCampo = 0;

  for (const concepto of conceptos) {
    let asignado = false;

    while (indiceCampo < CAPACIDADES.length) {
      const configuracion = CAPACIDADES[indiceCampo];
      const campoActual = resultado[configuracion.campo];

      const candidato = campoActual
        ? `${campoActual}/${concepto}`
        : concepto;

      if (candidato.length <= configuracion.maximo) {
        resultado[configuracion.campo] = candidato;
        asignado = true;
        break;
      }

      indiceCampo++;
    }

    if (!asignado) {
      throw new Error(
        `No hay espacio suficiente para almacenar el concepto "${concepto}".`
      );
    }
  }

  return resultado;
}
