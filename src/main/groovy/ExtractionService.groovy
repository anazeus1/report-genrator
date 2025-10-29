import org.apache.logging.log4j.LogManager
import org.apache.logging.log4j.Logger
import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.Workbook
import org.apache.poi.ss.usermodel.WorkbookFactory
import org.apache.poi.xssf.usermodel.XSSFCell
import org.apache.poi.xssf.usermodel.XSSFRow
import org.apache.poi.xssf.usermodel.XSSFSheet

import java.nio.file.Paths
import java.time.ZoneId
import java.time.LocalDate
import java.time.format.DateTimeFormatter
/**
 * Company:   PSI Software SE
 * @author:  Anas Gharbi
 */

class ExtractionService {
    XSSFSheet einsatzPlanSheet
    Workbook templateWorkbook
    private static final Logger logger = LogManager.getLogger(ExtractionService.class);

    /**
     * extract einsatzPlan sheet from Wartungsliste.xlsx
     *
     * @param filename Wartungsliste workbook fileName
     */
    def extractEinsatzPlanSheet(String fileName) {
        if(fileName.isEmpty()){
            fileName ="./data/Wartungsliste_2025.xlsx"
        }
        InputStream inp = new FileInputStream(fileName)
        def wb = WorkbookFactory.create(inp)
        einsatzPlanSheet = wb.getSheetAt(1)
    }

    /**
     * extract the template workbook from template.xlsx
     *
     * @param filename template workbook fileName
     */
    def extractTemplateSheets(String fileName) {
        if(fileName.isEmpty()){
            fileName ="./data/template.xlsx"
        }
        InputStream inp = new FileInputStream(fileName)
        templateWorkbook = WorkbookFactory.create(inp)
    }

    private static LocalDate getLocalDateFromCell(XSSFCell dateCell) {
        LocalDate localDate

        try {
            Date date = dateCell.getDateCellValue()
            localDate = date.toInstant()
                    .atZone(ZoneId.systemDefault())
                    .toLocalDate()
        }
        catch (Exception e) {
            localDate = null
        }
        return localDate
    }

    /**
     * generate excel file for each employee that have hours in the given month
     * and the last day of the previous month
     *
     * @param givenDate a date formatted as yyyy-MM-dd that contains the month and year given in the ui
     */
    def generateFilesByDate(LocalDate givenDate) {
        def generated = []

        for (int i = 4; i < einsatzPlanSheet.size(); i++) {
            def row = einsatzPlanSheet.getRow(i)
            def dateCell = row.getCell(0)
            def localDate = getLocalDateFromCell(dateCell)
            //includes current month and also last day in last month
            boolean sameMonthAndYear = (localDate != null) && (givenDate.getYear() == localDate.getYear())
                    && ((localDate.getDayOfYear() >= givenDate.getDayOfYear() - 1) &&
                    (localDate.getDayOfYear() < givenDate.getDayOfYear() + givenDate.lengthOfMonth()))

            if (sameMonthAndYear) {
                def formatter = DateTimeFormatter.ofPattern("MMMM-yyyy")
                def monthYear = givenDate.format(formatter)

                for (int j = 7; j < row.size(); j++) {
                    if (row.getCell(j).toString() == "w") {
                        def teamMember = getCell(einsatzPlanSheet, 0, j + 1)
                        def fileName = "${teamMember}-${monthYear}"
                        def acronym = getCell(einsatzPlanSheet, 5, j)

                        if (fileName in generated) {
                            setWorkTime(fileName, localDate, acronym)
                        } else {
                            generateFile(fileName, teamMember, givenDate, acronym)
                            setWorkTime(fileName, localDate, acronym)
                            generated.add(fileName)
                            logger.info("${fileName} is generated")
                        }
                    }
                }

            }
        }
        logger.info("${generated.size()} Excel files generated ")

    }

    private static void setWorkTime(String fileName, LocalDate date, acronym) {
        try (InputStream inp = new FileInputStream("output/${fileName}.xlsx")) {
            def wb = WorkbookFactory.create(inp)
            XSSFSheet sheet = wb.getSheetAt(0)
            String nextMonth = date.plusMonths(1).getMonth()

            //write hours from last month
            if (fileName.contains(nextMonth)) {
                setCell(sheet, 16, 4, acronym)
                setCell(sheet, 16, 5, 0)
                setCell(sheet, 16, 6, 8.5 / 24)
            }

            //find the row with the same date
            else {
                def rowIndex = 0
                for (int i = 0; i < sheet.size(); i++) {
                    LocalDate localDate = getLocalDateFromCell(sheet.getRow(i).getCell(0))
                    if (localDate != null && (localDate.getDayOfMonth() == date.getDayOfMonth())) {
                        rowIndex = i
                    }
                }

                //set working hours (from 16:30 till 24:00 and from 00:00 till 8:30 the next day)
                setCell(sheet, rowIndex + 1, 4, acronym)
                setCell(sheet, rowIndex, 4, acronym)
                setCell(sheet, rowIndex, 5, 16.5 / 24)
                setCell(sheet, rowIndex, 6, 1)
                setCell(sheet, rowIndex + 1, 5, 0)
                setCell(sheet, rowIndex + 1, 6, 8.5 / 24)
            }
            try (OutputStream fileOut = new FileOutputStream("output/${fileName}.xlsx")) {
                wb.write(fileOut)
                wb.close()
            }
        }
    }

    /**
     * generate xlsx file for a specific teamMember for a specific month
     * and have the same sheets of the template workbook
     *
     * @param fileName the generated file name
     * @param teamMember the first number
     * @param date that contains the month and year given in the ui
     * @param acronym teamMemberAcronym
     */
    def generateFile(String fileName, String teamMember, LocalDate date, String acronym) {
        XSSFSheet sheet = templateWorkbook.getSheetAt(0)

        setCell(sheet, 1, 9, teamMember)
        setCell(sheet, 2, 6, date)
        setCell(sheet, 2, 9, acronym)

        def outputDir = Paths.get("output").toFile()

        if (!outputDir.exists()) {
            outputDir.mkdirs()
        }
        try (OutputStream fileOut = new FileOutputStream("output/${fileName}.xlsx")) {
            templateWorkbook.write(fileOut)
        }
    }

    private static void setCell(XSSFSheet sheet, int rowIndex, int cellIndex, value) {
        XSSFRow row = sheet.getRow(rowIndex)

        if (row == null) {
            row = sheet.createRow(rowIndex)
        }
        Cell cell = row.getCell(cellIndex)
        if (cell == null) {
            cell = row.createCell(cellIndex)
        }
        cell.setCellValue(value)

    }

    private static String getCell(XSSFSheet sheet, int rowindex, int cellIndex) {
        XSSFRow row = sheet.getRow(rowindex)
        Cell cell = row.getCell(cellIndex)

        return cell.getStringCellValue()
    }
}
