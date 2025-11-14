function main(workbook: ExcelScript.Workbook) {
    // 
    const hoja = workbook.getActiveWorksheet();
    const filMax = hoja.getUsedRange().getRowCount();
    const contenidoTotal = hoja.getUsedRange().getValues();
    const fila1 = contenidoTotal[0];

    for (let fila = 1; fila < filMax; fila++){
        let fila2 = contenidoTotal[fila];
        let resultado = [fila1, fila2];

        console.log(String(resultado[0]) + "\n" + String(resultado[1]));

    }

}
