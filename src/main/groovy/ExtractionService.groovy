import org.apache.logging.log4j.LogManager
import org.apache.logging.log4j.Logger
import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.CellStyle
import org.apache.poi.ss.usermodel.CreationHelper
import org.apache.poi.ss.usermodel.WorkbookFactory
import org.apache.poi.xssf.usermodel.XSSFCell
import org.apache.poi.xssf.usermodel.XSSFRow
import org.apache.poi.xssf.usermodel.XSSFSheet
import org.apache.poi.xssf.usermodel.XSSFWorkbook

import java.time.ZoneId
import java.time.LocalDate

/**
 * Company:   PSI Software SE
 * @author: Anas Gharbi
 */

class ExtractionService {
    XSSFSheet BHSSaleSheet
    XSSFSheet MOBSaleSheet
    XSSFSheet BHSActSheet
    XSSFSheet MOBActSheet
    XSSFSheet templateSheet
    XSSFWorkbook templateWorkbook
    XSSFWorkbook SaleWorkbook
    String templateFileName
    Map<String, XSSFSheet> agentSheets = new HashMap<>()
    Map<String, String> AgentList = new HashMap<String, String>()

    CellStyle dateStyle
    CellStyle activeStyle
    CellStyle waitStyle
    private static final Logger logger = LogManager.getLogger(ExtractionService.class)

    /**
     * extract the report workbook
     *
     * @param workbook fileName
     */
    def extractReportWorkbook(String fileName) {
        if (fileName.isEmpty()) {
            fileName = "./data/report.xlsx"
        }
        InputStream inp = new FileInputStream(fileName)
        SaleWorkbook = WorkbookFactory.create(inp)
        BHSSaleSheet = SaleWorkbook.getSheetAt(0)
        BHSActSheet = SaleWorkbook.getSheetAt(1)
        MOBSaleSheet = SaleWorkbook.getSheetAt(2)
        MOBActSheet = SaleWorkbook.getSheetAt(3)
    }

    /**
     * extract the template workbook from template.xlsx
     * and the team list from first sheet
     * @param filename template workbook fileName
     */
    def extractTemplateSheets(String fileName) {
        if (fileName.isEmpty()) {
            fileName = "./data/template.xlsx"
        }
        InputStream inp = new FileInputStream(fileName)
        templateWorkbook = WorkbookFactory.create(inp)
        templateWorkbook.forEach { sheet -> agentSheets[sheet.getSheetName()] = sheet }
        def teamListSheet = templateWorkbook.getSheetAt(0)
        templateSheet = templateWorkbook.getSheetAt(1)
        for (int i = 1; i < teamListSheet.size(); i++) {
            XSSFRow row = teamListSheet.getRow(i)
            String agentId = row.getCell(0)
            String agentName = row.getCell(1)
            /*if(agentName in agentSheets.keySet()){
                createSheet(agentName,fileName)
            }*/
            agentList[agentId] = agentName
        }

        //get styles
        CreationHelper createHelper = templateWorkbook.getCreationHelper();
        dateStyle = templateWorkbook.createCellStyle();
        dateStyle.setDataFormat(
                createHelper.createDataFormat().getFormat("dd/MM/yyyy"));

        activeStyle = templateSheet.getRow(1).getCell(3).getCellStyle()
        waitStyle = templateSheet.getRow(1).getCell(5).getCellStyle()
        templateFileName = fileName

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
    def updateTeamMemberSheets(LocalDate givenDate) {


        for (int i = 1; i < BHSSaleSheet.size(); i++) {
            def row = BHSSaleSheet.getRow(i)
            def dateCell = row.getCell(9)
            def localDate = getLocalDateFromCell(dateCell)
            //includes current month and also last day in last month
            boolean isAfterOrEqualToGivenDate = (localDate != null) && (localDate.isEqual(givenDate) || localDate.isAfter(givenDate))

            if (isAfterOrEqualToGivenDate) {

                def agentId = getCell(BHSSaleSheet, i, 0)
                def agentName = AgentList[agentId]
                def orderNumber = getCell(BHSSaleSheet, i, 6)
                def segment = getCell(BHSSaleSheet, i, 7)

                if (agentId in agentList.keySet()) {
                    addSegment(localDate,agentName, orderNumber, segment)
                }
            }
        }
        for (int i = 1; i < MOBSaleSheet.size(); i++) {
            def row = MOBSaleSheet.getRow(i)
            def dateCell = row.getCell(3)
            def localDate = getLocalDateFromCell(dateCell)
            //includes current month and also last day in last month
            boolean isAfterOrEqualToGivenDate = (localDate != null) && (localDate.isEqual(givenDate) || localDate.isAfter(givenDate))

            if (isAfterOrEqualToGivenDate) {
                def agentId = getCell(MOBSaleSheet, i, 1)
                def agentName = AgentList[agentId]
                def orderNumber = getCell(MOBSaleSheet, i, 0)
                if (agentId in agentList.keySet()) {
                    addSegment(localDate, agentName, orderNumber, "MOB")
                }
            }
        }
    }


    private void addSegment(LocalDate saleDate, String agentName, String orderNumber, String segment) {
        XSSFSheet sheet = agentSheets[agentName]
        if (sheet == null)
            return
        int lastRow = sheet.getLastRowNum()
        int rowId
        boolean orderNumberExists = false
        //see if order is initilised
        for (int i = 0; i <= lastRow; i++) {
            def orderNumberOld = getCell(sheet, i, 4)
            println(orderNumber + "pld" + orderNumberOld)
            if (orderNumber.equals(orderNumberOld)) {
                println("equla")
                rowId = i
                orderNumberExists = true
            }
        }
        if (!orderNumberExists) {
            rowId = lastRow + 1
            sheet.createRow(rowId)
            //set order number
            setCell(sheet, rowId, 4, orderNumber)

            //set order date cell
            def dateCell= sheet.getRow(rowId).createCell(2)
            dateCell.setCellValue(saleDate)
            dateCell.setCellStyle(dateStyle)
        }
        if (segment == "Internet") {
            setCell(sheet, rowId, 5, 1);
        } else if (segment.contains("TV")) {

            setCell(sheet, rowId, 6, 1);
        } else if (segment == "HP") {
            setCell(sheet, rowId, 7, 1);
        }
        else if(segment =="MOB"){
            setCell(sheet,rowId,8,1)
        }
    }

    public void saveFile() {
        def newFile = new File("template_output.xlsx")
        def out = new FileOutputStream(newFile)
        templateWorkbook.write(out)
        out.close()
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
    def createSheet(String teamMember, String fileName) {
        def newSheet = templateWorkbook.cloneSheet(1)
        newSheet.setSheetName(1, "CopiedSheet") // rename


        setCell(newSheet, 0, 0, teamMember)
        try (OutputStream fileOut = new FileOutputStream(fileName)) {
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
        if (row == null) return null //row doest exist
        Cell cell = row.getCell(cellIndex)
        if (cell == null) return null  // cell doesn't exist
        return cell.getStringCellValue()
    }

}
