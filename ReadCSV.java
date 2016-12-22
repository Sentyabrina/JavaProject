package readcsv;

import java.io.DataInputStream;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;

/**
 * @author Katerina
 * @version 1.1
*/
public class ReadCSV {
    /** * Класс для чтения информации с CSV файла
     * считывает построчно, разбивая информацию по ячейкам
     * разделитель ячейки - ";"
     * @exception IOException - исключение, если файл не был создан
     * @return  информацию в строках и столбцах таблицы
     */
    public ArrayList csv() throws IOException {
        ArrayList<ArrayList<String>> allRowAndColData = null;
        ArrayList<String> oneRowData = null;
        /** указание путя для выбора файла */
        String fName = "C:\\Users\\Katerina\\IdeaProjects\\JavaProject\\src\\readcsv\\inputInfo.csv";
        String currentLine;
        /**открытие потока для работы с файлом */
        FileInputStream fis = new FileInputStream(fName);
        DataInputStream myInput = new DataInputStream(fis);
        int i = 0;
        /** переносим данные  в массив */
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
