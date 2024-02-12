import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import org.apache.poi.POIXMLException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.*;

public class TextStatistics {

    public static List<String> EXCEL_HEADERS = Arrays.asList("Шаблон", "Количество слов");

    public static void main(String []args) throws IOException, POIXMLException, FileNotFoundException {

        if (args.length != 3) {
            System.out.println("ERROR: Expected 3 params to execute, but got " + args.length);
            return;
        }

        //Retrieving args
        Path textFile = Paths.get(args[0]);
        String patternsFile = Paths.get(args[1]).toString();
        String finalFile = Paths.get(args[2]).toString();

        //Lowering case in order to match given patterns
        String text = new String(Files.readAllBytes(textFile)).toLowerCase();

        if (text.length() == 0) {
            System.out.println("WARN: File with text is empty");
        }

        //Reading patterns from Excel File
        XSSFWorkbook wb = new XSSFWorkbook(patternsFile);
        XSSFSheet sheet = wb.getSheetAt(0);
        Iterator<Row> rowIter = sheet.rowIterator();

        if (!rowIter.hasNext()) {
            System.out.println("WARN: File with patterns is empty");
        }

        List<String> patterns = new ArrayList<>();

        //Putting patterns into patterns list
        while (rowIter.hasNext()) {
            Row row = rowIter.next();
            String cellValue = "";
            if (row.getCell(0).getCellType() == 0) {
                cellValue = String.valueOf( (int) row.getCell(0).getNumericCellValue());
            }
            else cellValue = row.getCell(0).toString();
            patterns.add(cellValue);
        }

        //Preparing final map to write then into Excel file
        Map<String, Integer> finalMap = new HashMap<>();
        for (String pattern : patterns) {
            finalMap.put(pattern, 0);
        }

        //Regexp to match chars and the number of expected occurrences
        String chars = "[а-яa-z]\\d.*?";
        //Regexp to match words in given text
        String words = "[а-яa-z-\\d]+";

        Pattern p = Pattern.compile(words);
        Matcher m = p.matcher(text);

        //Iterate each word
        while (m.find()) {
            //Iterate each pattern
            for (String textPattern : patterns) {
                Pattern patternChars = Pattern.compile(chars);
                Matcher matcherChars = patternChars.matcher(textPattern);

                //If a pattern is a type of chars and the number of occurrences in a word, iterate over it
                boolean matchesChars = false;
                while (matcherChars.find()) {
                    matchesChars = true;
                    String ch = matcherChars.group().substring(0, 1);
                    Integer count = Integer.parseInt(matcherChars.group().substring(1));

                    //Regexp to match a word with a char and a number of its occurrences
                    String str = String.format("^(?:[^%s\\r\\n]*%s[^%s\\r\\n]*){%d}$", ch, ch, ch, count);
                    Pattern patternWord = Pattern.compile(str);
                    Matcher matcherWord = patternWord.matcher(m.group());

                    //If at least 1 char doesn't match a word, the whole pattern doesn't match as well
                    if (!matcherWord.matches()) {
                        matchesChars = false;
                        break;
                    }
                }

                //In a case when second type of pattern found - sequence of chars
                //Regexp to match a word with a sequence of chars
                String str = String.format(".*?%s.*?", textPattern);
                Pattern patternSeq = Pattern.compile(str);
                Matcher matcherSeq = patternSeq.matcher(m.group());

                boolean matchesSeq = matcherSeq.matches();

                //Updating values inside a final map
                if (matchesChars || matchesSeq) {
                    if (finalMap.get(textPattern) != null) {
                        int count = finalMap.get(textPattern);
                        finalMap.put(textPattern, ++count);
                    } else finalMap.put(textPattern, 1);
                }
            }
        }

        //Creating final Excel Workbook
        Workbook finalWb = new SXSSFWorkbook();
        Sheet finalSheet = finalWb.createSheet("Шаблоны-Появления");

        //Creating style for a header row
        XSSFFont font = (XSSFFont) finalWb.createFont();
        font.setBold(true);
        XSSFCellStyle style = (XSSFCellStyle) finalWb.createCellStyle();
        style.setFont(font);
        style.setAlignment(CellStyle.ALIGN_CENTER);

        //Creating style for cells with values
        XSSFFont fontCells = (XSSFFont) finalWb.createFont();
        XSSFCellStyle styleCells = (XSSFCellStyle) finalWb.createCellStyle();
        styleCells.setFont(fontCells);
        styleCells.setAlignment(CellStyle.ALIGN_CENTER);

        //Putting header row into Excel
        int rowNum = 0;
        Row headerRow = finalSheet.createRow(rowNum++);
        for (int i = 0; i < EXCEL_HEADERS.size(); i++) {
            Cell cell =  headerRow.createCell(i);
            cell.setCellStyle(style);
            cell.setCellValue(EXCEL_HEADERS.get(i));
        }

        //Putting values from final map into Excel
        for (String key : finalMap.keySet()) {
            Row row = finalSheet.createRow(rowNum++);

            Cell cellPattern = row.createCell(0);
            cellPattern.setCellStyle(styleCells);
            cellPattern.setCellValue(key);

            Cell cellSummary = row.createCell(1);
            cellSummary.setCellStyle(styleCells);
            cellSummary.setCellValue((double) finalMap.get(key));
        }

        //Autosize only after data was added
        finalSheet.autoSizeColumn(0);
        finalSheet.autoSizeColumn(1);

        //Writing in a file from args
        File file = new File(finalFile);
        finalWb.write(new FileOutputStream(file, false));
        finalWb.close();
    }
}
