import java.io.FileOutputStream;
import java.util.*;

import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Main {
    public static void main(String[] args) {
        Map<String, List<String[]>> timetable = new HashMap<>();
        List<String> subjects = Arrays.asList("Math", "Science", "English", "History");
        List<String> teachers = Arrays.asList("Teacher A", "Teacher B", "Teacher C", "Teacher D");
        List<String> classrooms = Arrays.asList("Classroom 1", "Classroom 2", "Classroom 3");
        Scanner sc = new Scanner(System.in);
        boolean repeat = true;
        while (repeat) {
            System.out.print("Enter group name: ");
            String groupName = sc.nextLine();
            System.out.print("Enter number of pairs: ");
            int numPairs = sc.nextInt();
            sc.nextLine();
            if (numPairs == 0) {
                System.out.println("No pairs were added for " + groupName);
                continue;
            }
            int[] subjectNumbers = new int[numPairs];
            int[] classroomNumbers = new int[numPairs];
            int[] teacherNumbers = new int[numPairs];
            List<String[]> pairs = new ArrayList<>();
            for (int i = 0; i < numPairs; i++) {
                System.out.println("\nSelect subjects: ");
                for (int j = 0; j < subjects.size(); j++) {
                    System.out.println((j + 1) + ": " + subjects.get(j));
                }
                System.out.print("Enter subject number: ");
                subjectNumbers[i] = sc.nextInt();
                sc.nextLine();
                while (subjectNumbers[i] < 1 || subjectNumbers[i] > subjects.size()) {
                    System.out.println("Invalid subject number. Please enter a valid number: ");
                    subjectNumbers[i] = sc.nextInt();
                    sc.nextLine();
                }

                System.out.println("\nSelect classrooms: ");
                for (int j = 0; j < classrooms.size(); j++) {
                    System.out.println((j + 1) + ": " + classrooms.get(j));
                }
                System.out.print("Enter classroom number: ");
                classroomNumbers[i] = sc.nextInt();
                sc.nextLine();
                while (classroomNumbers[i] < 1 || classroomNumbers[i] > classrooms.size()) {
                    System.out.println("Invalid classroom number. Please enter a valid number: ");
                    classroomNumbers[i] = sc.nextInt();
                    sc.nextLine();
                }

                System.out.println("\nSelect teachers: ");
                for (int j = 0; j < teachers.size(); j++) {
                    System.out.println((j + 1) + ": " + teachers.get(j));
                }
                System.out.print("Enter teacher number: ");
                teacherNumbers[i] = sc.nextInt();
                sc.nextLine();
                while (teacherNumbers[i] < 1 || teacherNumbers[i] > teachers.size()) {
                    System.out.println("Invalid teacher number. Please enter a valid number: ");
                    teacherNumbers[i] = sc.nextInt();
                    sc.nextLine();
                }
                String[] pair = {String.valueOf(i + 1), subjects.get(subjectNumbers[i] - 1),
                        classrooms.get(classroomNumbers[i] - 1),
                        teachers.get(teacherNumbers[i] - 1)};
                pairs.add(pair);
            }
            timetable.put(groupName, pairs);
            System.out.print("Do you want to enter more groups? (yes/no): ");
            String answer = sc.nextLine();
            if (!answer.equalsIgnoreCase("yes")) {
                repeat = false;
            }
        }

        String[] header = {"Group", "Pair", "Subject", "Classroom", "Teacher"};
        try (Workbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("Timetable");

            Font headerFont = workbook.createFont();
            headerFont.setBold(true);
            headerFont.setColor(IndexedColors.BLUE.getIndex());

            CellStyle headerCellStyle = workbook.createCellStyle();
            headerCellStyle.setFont(headerFont);
            headerCellStyle.setVerticalAlignment(VerticalAlignment.CENTER); // added
            headerCellStyle.setAlignment(HorizontalAlignment.CENTER); // added
            headerCellStyle.setBorderTop(BorderStyle.THIN); // added
            headerCellStyle.setBorderBottom(BorderStyle.THIN); // added
            headerCellStyle.setBorderLeft(BorderStyle.THIN); // added
            headerCellStyle.setBorderRight(BorderStyle.THIN); // added

            CellStyle cellStyle = workbook.createCellStyle();
            cellStyle.setBorderTop(BorderStyle.THIN); // added
            cellStyle.setBorderBottom(BorderStyle.THIN); // added
            cellStyle.setBorderLeft(BorderStyle.THIN); // added
            cellStyle.setBorderRight(BorderStyle.THIN); // added

            Row headerRow = sheet.createRow(0);
            for (int i = 0; i < header.length; i++) {
                Cell cell = headerRow.createCell(i);
                cell.setCellValue(header[i]);
                if (i == 0) {
                    cell.setCellStyle(headerCellStyle); // set the style for the "Group" column
                } else {
                    cell.setCellStyle(headerCellStyle); // set the same style for other columns
                }
            }

            int rowNum = 1;
            for (Map.Entry<String, List<String[]>> entry : timetable.entrySet()) {
                int numPairs = entry.getValue().size();
                int startRow = rowNum;
                int endRow = rowNum + numPairs - 1;
                Row groupRow = sheet.getRow(startRow);
                if (groupRow == null) {
                    groupRow = sheet.createRow(startRow);
                }
                Cell groupCell = groupRow.createCell(0);
                groupCell.setCellValue(entry.getKey());
                groupCell.setCellStyle(headerCellStyle); // set the style for the "Group" column
                int colNum = 1;
                for (String[] data : entry.getValue()) {
                    Row dataRow = sheet.getRow(rowNum);
                    if (dataRow == null) {
                        dataRow = sheet.createRow(rowNum);
                    }
                    for (int i = 0; i < data.length; i++) {
                        Cell cell = dataRow.createCell(colNum++);
                        cell.setCellValue(data[i]);
                        cell.setCellStyle(cellStyle); // set the style for the data cells
                    }
                    colNum = 1;
                    rowNum++;
                }
                if (numPairs > 1) {
                    sheet.addMergedRegion(new CellRangeAddress(startRow, endRow, 0, 0));
                }
            }

        for (int i = 0; i < header.length; i++) {
                sheet.autoSizeColumn(i);
            }

            FileOutputStream fileOut = new FileOutputStream("timetable.xlsx");
            workbook.write(fileOut);
            fileOut.close();
            workbook.close();
            System.out.println("Timetable written to 'timetable.xlsx'");
        } catch (Exception e) {
            System.out.println("Error writing to file: " + e.getMessage());
        }
    }
}
