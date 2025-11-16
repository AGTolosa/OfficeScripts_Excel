function main(workbook: ExcelScript.Workbook) {
    const hoja = workbook.getActiveWorksheet();
    const filMax = hoja.getUsedRange().getRowCount();

    function limpiarDNIs(columna: number) {
        let columnaObjetivo = hoja.getRangeByIndexes(0, columna - 1, filMax, 1).getValues() as string[][];

        for (let celda = 1; celda < filMax; celda++) {
            let valor = String(columnaObjetivo[celda][0]);
            hoja.getRangeByIndexes(celda, columna - 1, 1, 1).setValue(valor.replace(/[- ]/g, ""));
        };
    };

    limpiarDNIs(20);

}
