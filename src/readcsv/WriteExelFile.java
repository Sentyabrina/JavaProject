package readcsv;

import javafx.stage.FileChooser;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblBorders;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STBorder;

import java.io.*;
import java.util.ArrayList;
/**
 *@author Katerina Perepechenko 535st group
 *@version 1.0
 *класс для записи  CSV-файла
 */

public class WriteExelFile {
    @SuppressWarnings("deprecation")
    public static void main(String args[]) throws IOException, InvalidFormatException {
        /* open CSV-file. */
       /* ArrayList<ArrayList<String>> allRowAndColData = null;
        ArrayList<String> oneRowData = null;
        String fName = "C:\\Users\\Katerina\\IdeaProjects\\JavaProject\\src\\readcsv\\inputInfo.csv";
        String currentLine;
        FileInputStream fis = new FileInputStream(fName);
        DataInputStream myInput = new DataInputStream(fis);
        int i = 0;
        allRowAndColData = new ArrayList<ArrayList<String>>();
        while ((currentLine = myInput.readLine()) != null) {
            oneRowData = new ArrayList<String>();
            String oneRowArray[] = currentLine.split(";");
            for (int j = 0; j < oneRowArray.length; j++) {
                oneRowData.add(oneRowArray[j]);
            }
            allRowAndColData.add(oneRowData);
            System.out.println();
            i++;

        }
        /*-------------------------------------*/
        /*open docx file*/
        //FileInputStream f = new FileInputStream("C:\\Users\\Katerina\\IdeaProjects\\JavaProject\\src\\readcsv\\template.docx");
       /*XWPFDocument doc=new XWPFDocument();

        /* create table */
     /*   XWPFTable table = doc.createTable(0,0);


        /* save doc file*/
    /*    FileChooser fileChooser = new FileChooser();

        fileChooser.setTitle("Сохранить как ...");
        FileChooser.ExtensionFilter extFilter = new FileChooser.ExtensionFilter("Документ Word (*.docx)", "*.docx");
        fileChooser.getExtensionFilters().add(extFilter);
        fileChooser.setInitialFileName("111.docx");
        File userDirectory = new File("C:\\Users\\Katerina\\IdeaProjects\\JavaProject\\src\\readcsv\\111.doc");
        fileChooser.setInitialDirectory(userDirectory);

        try {

            for (i = 0; i < allRowAndColData.size(); i++) {
                ArrayList<?> ardata = (ArrayList<?>) allRowAndColData.get(i);
                XWPFTableRow row = table.createRow();
                for (int k = 0; k < ardata.size(); k++) {
                    System.out.print(ardata.get(k));
                    XWPFTableCell cell = row.createCell();
                    cell.setText(ardata.get(k).toString());

                }
                System.out.println();

            }

            /* write docx file */
    /*        OutputStream outputStream = new FileOutputStream(new File(userDirectory.getAbsolutePath()));
            doc.write(outputStream);
            outputStream.flush();
            outputStream.close();
        } catch (Exception ex) {
        }*/

    ReadCSV read = new ReadCSV();
    WriteDOCX write = new WriteDOCX();
    write.writeDoc(read.csv());
    }
}