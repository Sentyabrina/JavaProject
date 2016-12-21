package readcsv;

import javafx.stage.FileChooser;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTRow;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTbl;

import java.io.*;
import java.util.ArrayList;

/**
 * Created by Katerina on 21.12.2016 at 23:56.
 */
public class WriteDOCX {
    public void writeDoc(ArrayList csvFile) throws IOException, InvalidFormatException {
        //XWPFDocument doc = new XWPFDocument();
        FileInputStream fis = new FileInputStream("C:\\Users\\Katerina\\IdeaProjects\\JavaProject\\src\\readcsv\\template.docx");
        XWPFDocument doc = new XWPFDocument(OPCPackage.open(fis));

        /* create table */
        XWPFTable table = doc.getTables().get(0);
        CTTbl ctTbl = CTTbl.Factory.newInstance(); // Create a new CTTbl for the new table

        /* save doc file*/
        FileChooser fileChooser = new FileChooser();

        fileChooser.setTitle("Сохранить как ...");
        FileChooser.ExtensionFilter extFilter = new FileChooser.ExtensionFilter("Документ Word (*.docx)", "*.docx");
        fileChooser.getExtensionFilters().add(extFilter);
        fileChooser.setInitialFileName("111.docx");
        File userDirectory = new File("C:\\Users\\Katerina\\IdeaProjects\\JavaProject\\src\\readcsv\\111.doc");
        fileChooser.setInitialDirectory(userDirectory);
        int i = 0;
        try {

            for (i = 1; i < csvFile.size(); i++) {
                ArrayList<?> ardata = (ArrayList<?>) csvFile.get(i);
                XWPFTableRow row1 = table.createRow();
                for (int k = 0; k < ardata.size(); k++) {
                    //System.out.print(ardata.get(k));
                    XWPFTableCell cell = row1.createCell();
                    cell.setText(ardata.get(k).toString());
                }
              //  table.removeRow(i);
                System.out.println();

            }
            System.out.println("Done");

            /* write docx file */
            OutputStream outputStream = new FileOutputStream(new File(userDirectory.getAbsolutePath()));
            doc.write(outputStream);
            outputStream.flush();
            outputStream.close();
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}