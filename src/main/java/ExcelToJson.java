import java.io.File;
import java.io.FileInputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Collection;
import java.util.Iterator;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.stream.Collectors;

import model.Student;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.fasterxml.jackson.core.JsonProcessingException;
import com.fasterxml.jackson.databind.ObjectMapper;

public class ExcelToJson {

    private static List<Student> readExcelFile(String filePath) {
        try {
            FileInputStream excelFile = new FileInputStream(new File(filePath));
            Workbook workbook = new XSSFWorkbook(excelFile);

            Sheet sheet = workbook.getSheetAt(0);
            Iterator<Row> rows = sheet.iterator();
            List<Student> lstStudents = new ArrayList<Student>();


            int rowNumber = 0;
            while (rows.hasNext()) {
                Row currentRow = rows.next();

                // skip header
                if (rowNumber == 0) {
                    rowNumber++;
                    continue;
                }

                Iterator<Cell> cellsInRow = currentRow.iterator();

                Student student = new Student();

                int cellIndex = 0;
                while (cellsInRow.hasNext()) {
                    Cell currentCell = cellsInRow.next();

                    if (cellIndex == 0) {
                        student.setName(String.valueOf(currentCell.getStringCellValue()));
                    } else if (cellIndex == 1) {
                        student.setAge((int) currentCell.getNumericCellValue());
                    } else if (cellIndex == 2) {
                        student.setMarks((int) currentCell.getNumericCellValue());
                    }
                    cellIndex++;
                }
                lstStudents.add(student);
            }
            return lstStudents;

        } catch (IOException e) {
            throw new RuntimeException("message = " + e.getMessage());
        }
    }

    static <T> Collection<List<T>> partitionBasedOnSize(List<T> inputList, int size) {
        final AtomicInteger counter = new AtomicInteger(0);
        return inputList.stream().collect(Collectors.groupingBy(s -> counter.getAndIncrement() / size)).values();
    }

    public static Map<String, Map<String, List<Student>>> pagingData(List<Student> students) {
        Map<String, Map<String, List<Student>>> actualStudent = new LinkedHashMap<>();
        Collection<List<Student>> c = partitionBasedOnSize(students, 1);
        System.out.println(c);
        int i = 1;
        Map<String, List<Student>> studentMap = new LinkedHashMap<>();
        for (List<Student> list : c) {
            studentMap.put("student" + i, list);
            i++;
        }
        actualStudent.put("Details", studentMap);
        return actualStudent;
    }

    private static String convertObjects2JsonString(Map<String, Map<String, List<Student>>> students) {
        ObjectMapper mapper = new ObjectMapper();
        String jsonString = "";
        try {
            jsonString = mapper.writerWithDefaultPrettyPrinter().writeValueAsString(students);
        } catch (JsonProcessingException e) {
            e.printStackTrace();
        }

        return jsonString;
    }

    public static void main(String[] args) {
        Map<String, Map<String, List<Student>>> students = pagingData(readExcelFile("student_list.xlsx"));
        try (FileWriter file = new FileWriter("students.json")) {
            file.write(convertObjects2JsonString(students));
            file.flush();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

}