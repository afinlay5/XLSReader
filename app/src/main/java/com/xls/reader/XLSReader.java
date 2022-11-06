/*
 * Demo app for reading xls,xlsx.
 */
package com.xls.reader;

import static java.lang.System.out;

import java.util.List;
import java.util.ArrayList;
import java.util.Map;
import java.util.HashMap;
import java.io.IOException;
import java.io.InputStream;
import org.apache.poi.ss.usermodel.WorkbookFactory;
public class XLSReader {
	public static void main (String[] args) throws IOException {
		readExcel(XLSReader.class.getResourceAsStream("/xls/DietPlan.xlsx"))
                .entrySet()
                .stream()
                .filter(e -> !e.getValue().isEmpty())
                .forEach(e-> out.printf("k:[%s], v:%s\n",e.getKey()+1,e.getValue()));
	}
	
	public static Map<Integer, List<String>> readExcel(InputStream excelFileIS) throws IOException {
 
        var data = new HashMap<Integer, List<String>>();

        var workbook = WorkbookFactory.create(excelFileIS);
        var sheet = workbook.getSheetAt(3);
        int rowMax = 50; //sheet.getPhysicalNumberOfRows();

        for (int rowCt = 0; rowCt < rowMax; rowCt++) {
            data.put(rowCt, new ArrayList<>());
            var row = sheet.getRow(rowCt);
            if (row == null)
                continue;
            int maxPhysicalNumCells = row.getPhysicalNumberOfCells();
            if (maxPhysicalNumCells>0)
                out.printf("ROW %s has %s cell(s)\n",row.getRowNum()+1, row.getPhysicalNumberOfCells());
            for (int colPos = 0; colPos < maxPhysicalNumCells; colPos++) {
                var cell = row.getCell(colPos);
                
                final String value;
                if (cell != null) {
                    var CELL_TYPE = cell.getCellType();
                    switch (CELL_TYPE) {
                        case BOOLEAN:
                            value = String.valueOf(cell.getBooleanCellValue());
                            break;
                        case FORMULA:
                            value = String.format("FORMULA:[%s]",cell.getCellFormula());
                            break;
                        case NUMERIC:
                            value = String.valueOf(cell.getNumericCellValue());
                            break;
                        case STRING:
                            value = cell.getStringCellValue();
                            break;

                        case ERROR:
                            value = String.valueOf(cell.getErrorCellValue());
                            break;
                        case BLANK:
                            value = "BLANK";
                            break;
                        case _NONE:
                            value = "NONE";
                            break;
                        default:
                            value = "DEFAULT";
                            break;
                    }
                    data.get(rowCt).add(value);
                }
                else
                    out.printf("Ran into a null cell at index [row:%s][col:%s] \n", rowCt+1,colPos+1);

            }
        }
        out.println();
        return data;
    }
}