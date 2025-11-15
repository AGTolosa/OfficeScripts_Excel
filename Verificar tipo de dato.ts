function main(workbook: ExcelScript.Workbook) {
    
    const hoja = workbook.getActiveWorksheet();
    const filMax = hoja.getUsedRange().getRowCount();

    function verificarTipoDato(columna: number){
        let columnaObjetivo = hoja.getRangeByIndexes(0, columna - 1, filMax, 1).getValues();

        for (let celda = 1; celda < filMax; celda++){
            if (!/^\d+$/.test(String(columnaObjetivo[celda][0]))){
                hoja.getCell(celda, columna - 1).getFormat().getFill().setColor("pink");
            }
        }
    }

    verificarTipoDato(1)
}
