function main(workbook: ExcelScript.Workbook) {

    let hojaObjetivo = workbook.getActiveWorksheet();
    let filMax = hojaObjetivo.getUsedRange().getRowCount();    
    let valores = hojaObjetivo.getRangeByIndexes(0, 0, filMax, 2).getValues() as string[][];
    
  console.log(valores);

    const valoresNuevos = [
        ["Valor-A", "NuevoValor-A"],
        ["Valor-B", "NuevoValor-B"],
        ["Valor-C", "NuevoValor-C"],
        ["Valor-D", "NuevoValor-D"],
    ];

    for (let i = 1; i < filMax; i++) {
        let valoresCambiar = valores[i][0];
        let valorPorDefecto = "Otro";

        for (let j = 0; j < valoresNuevos.length; j++) {
            if (valoresCambiar === valoresNuevos[j][0]) {
                valorPorDefecto = valoresNuevos[j][1];
                break;
            }
        }

        hojaObjetivo.getCell(i, 1).setValue(valorPorDefecto);
    };

    let valoresActualizados = hojaObjetivo.getRangeByIndexes(0, 0, filMax, 2).getValues() as string[][];
    console.log(valoresActualizados);
}
