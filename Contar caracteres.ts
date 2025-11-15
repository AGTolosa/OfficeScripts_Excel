function main(workbook: ExcelScript.Workbook) {
    
    const hoja = workbook.getActiveWorksheet();
    const contenido = hoja.getUsedRange().getValues();
    const colMax = hoja.getUsedRange().getColumnCount();
    const filMax = hoja.getUsedRange().getRowCount();

    let maxCaracteres = 0
    let celdaObjetivo = "";
    
    for (let columna = 0; columna < colMax; columna++){
        for (let fila = 0; fila < filMax; fila++){
            let celda = String(contenido[columna][fila]);
            let numCaracteres = celda.length;
            
            if (numCaracteres > maxCaracteres){
                maxCaracteres = numCaracteres;
                // .getAddress() muestra el nombre de la hoja + ! + coordenadas de la celda.
                celdaObjetivo = hoja.getCell(columna, fila).getAddress();
            }
        }
    }

    console.log('Maximo:  ' + maxCaracteres + ' en la celda: ' + celdaObjetivo)
}
