function main(workbook: ExcelScript.Workbook) {

    let hojaObjetivo = workbook.getActiveWorksheet();
    let filMax = hojaObjetivo.getUsedRange().getRowCount();
    let valores = hojaObjetivo.getRangeByIndexes(1, 0, filMax, 2).getValues() as String[][];
    console.log(valores);

    for (let valor = 1; valor < filMax; valor++) {
        let valorNuevo = String(valores[valor][0]);
        switch (valorNuevo) {
            case "Valor-A":
                hojaObjetivo.getCell(valor + 1, 1).setValue("NuevoValor-A");
                break;
            case "Valor-B":
                hojaObjetivo.getCell(valor + 1, 1).setValue("NuevoValor-B");
                break;
            default:
                hojaObjetivo.getCell(valor + 1, 1).setValue("No hay referencia");
        }
    };

    let valoresActualizados = hojaObjetivo.getRangeByIndexes(1, 0, filMax, 2).getValues() as String[][];
    console.log(valoresActualizados);

}
