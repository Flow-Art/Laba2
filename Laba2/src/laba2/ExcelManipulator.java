/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package laba2;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.FileSystems;
import java.util.Collection;
import java.util.Collections;
import java.util.HashMap;
import java.util.Map;
import java.util.logging.Level;
import java.util.logging.Logger;
import javafx.scene.control.Cell;
import org.apache.commons.math3.stat.descriptive.DescriptiveStatistics;
import org.apache.poi.sl.draw.geom.Path;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author Admin
 */
public class ExcelManipulator {
    public ExcelManipulator() {
    
};
    HashMap<String, double[]> MyExport = new HashMap();
    
    public void export() throws FileNotFoundException, IOException {
        
        java.nio.file.Path file_path = FileSystems.getDefault().getPath("ДЗ2.xlsx");
        XSSFWorkbook MyBook = new XSSFWorkbook(new FileInputStream(file_path.toString()));
        XSSFSheet MySheet = MyBook.getSheet("Вариант 10");
        int rowCount = MySheet.getPhysicalNumberOfRows();
        XSSFRow headers = MySheet.getRow(0);
        for (int i=0; i<headers.getPhysicalNumberOfCells() ; i++) {
            XSSFCell header = headers.getCell(i);
            String ColName = header.getStringCellValue();
            double[] values = new double[rowCount-1];
            int k = 0;
            for (int j=1; j<rowCount; j++) {
                values[k] = MySheet.getRow(j).getCell(i).getNumericCellValue();
                k++;
            }
            MyExport.put(ColName, values);
        }
      
    }
    public void result() throws FileNotFoundException, IOException, FileWasNotImportedException {
        if (MyExport.isEmpty()) {
            throw new FileWasNotImportedException();
        }
              
        XSSFWorkbook MyBook = new XSSFWorkbook();
        XSSFSheet MySheet = MyBook.createSheet("Вариант 10");
        Row row1 = MySheet.createRow(0);
        Row row2 = MySheet.createRow(1);
        org.apache.poi.ss.usermodel.Cell dx = row2.createCell(0);
        dx.setCellValue("Оценка дисперсии выборки = ");
        MySheet.autoSizeColumn(0);
        int i=1;
        for (Map.Entry<String, double[]> pair:MyExport.entrySet()) {
            DescriptiveStatistics descriptiveStatistics = new DescriptiveStatistics();
            double[] vals = pair.getValue();
            for (double v : vals) {
                descriptiveStatistics.addValue(v);
            }
            double disp = descriptiveStatistics.getVariance();
            org.apache.poi.ss.usermodel.Cell header = row1.createCell(i);
            header.setCellValue(pair.getKey());
            org.apache.poi.ss.usermodel.Cell dispcell = row2.createCell(i);
            dispcell.setCellValue(disp);
            i++;
        }
        
        try {
            MyBook.write(new FileOutputStream("Расчёты.xlsx"));
        } catch (IOException ex) {
            Logger.getLogger(ExcelManipulator.class.getName()).log(Level.SEVERE, null, ex);
        }
        MyBook.close();
       
    }
}
