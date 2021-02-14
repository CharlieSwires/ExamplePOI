package com.charlie.ExamplePOI;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.Calendar;
import java.util.Iterator;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

import com.itextpdf.text.Document;
import com.itextpdf.text.Phrase;
import com.itextpdf.text.pdf.PdfPCell;
import com.itextpdf.text.pdf.PdfPTable;
import com.itextpdf.text.pdf.PdfWriter;

public class Main{  
        public static void main(String[] args) throws Exception{

                FileInputStream input_document = new FileInputStream(new File("C:\\Users\\charl\\eclipse-workspace\\ExamplePOI\\excel_to_pdf.xls"));
                // Read workbook into HSSFWorkbook
                HSSFWorkbook my_xls_workbook = new HSSFWorkbook(input_document); 
                // Read worksheet into HSSFSheet
                HSSFSheet my_worksheet = my_xls_workbook.getSheetAt(0); 
                // To iterate over the rows
                Iterator<Row> rowIterator = my_worksheet.iterator();
                //We will create output PDF document objects at this point
                Document iText_xls_2_pdf = new Document();
                PdfWriter.getInstance(iText_xls_2_pdf, new FileOutputStream("Excel2PDF_Output.pdf"));
                iText_xls_2_pdf.open();
                //we have two columns in the Excel sheet, so we create a PDF table with two columns
                //Note: There are ways to make this dynamic in nature, if you want to.
                PdfPTable my_table = new PdfPTable(4);
                //We will use the object below to dynamically add new data to the table
                PdfPCell table_cell;
                //Loop through rows.
                while(rowIterator.hasNext()) {
                        Row row = rowIterator.next(); 
                        Iterator<Cell> cellIterator = row.cellIterator();
                                while(cellIterator.hasNext()) {
                                    Calendar instance;
                                        Cell cell = cellIterator.next();//Fetch CELL
                                        switch(cell.getCellType().toString()) { //Identify CELL type
                                                //you need to add more code here based on
                                                //your requirement / transformations
                                        case "STRING":
                                            System.out.println("STRING");

                                            //Push the data from Excel to PDF Cell
                                             table_cell=new PdfPCell(new Phrase(cell.getStringCellValue()));
                                             //feel free to move the code below to suit to your needs
                                             my_table.addCell(table_cell);
                                            break;
                                        case "NUMERIC":
                                            System.out.println("NUMERIC");
                                            instance = Calendar.getInstance();
                                            instance.setTime(cell.getDateCellValue());
                                            if (instance.get(Calendar.YEAR) == 2021) {
                                            table_cell=new PdfPCell(new Phrase(""+instance.get(Calendar.DATE)+"/"+(1+instance.get(Calendar.MONTH))+"/"+instance.get(Calendar.YEAR)));
                                            }else {
                                              //Push the data from Excel to PDF Cell
                                              table_cell=new PdfPCell(new Phrase(""+cell.getNumericCellValue()));
                                              //feel free to move the code below to suit to your needs
                                                
                                            }
                                             my_table.addCell(table_cell);
                                            break;
                                        case "DATE":
                                            System.out.println("DATE");
                                            //Push the data from Excel to PDF Cell
                                            instance = Calendar.getInstance();
                                            instance.setTime(cell.getDateCellValue());
                                            table_cell=new PdfPCell(new Phrase(""+instance.get(Calendar.DATE)+"/"+instance.get(Calendar.MONTH)+"/"+instance.get(Calendar.YEAR)));
                                             //feel free to move the code below to suit to your needs
                                             my_table.addCell(table_cell);
                                            break;
                                        }
                                        //next line
                                }

                }
                //Finally add the table to PDF document
                iText_xls_2_pdf.add(my_table);                       
                iText_xls_2_pdf.close();                
                //we created our pdf file..
                input_document.close(); //close xls
        }
}

