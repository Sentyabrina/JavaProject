package readcsv;

import java.io.DataInputStream;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;

/**
 * Created by Katerina on 21.12.2016 at 23:48.
 */
public class ReadCSV {
    public ArrayList csv() throws IOException {
        ArrayList<ArrayList<String>> allRowAndColData = null;
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
        return allRowAndColData;
    }

}
