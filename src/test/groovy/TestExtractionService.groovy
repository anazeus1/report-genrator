
import org.junit.Test

import java.time.LocalDate


class  TestExtractionService {
    ExtractionService extractionService = new ExtractionService()

    @Test
    void testExtractReport() {
        extractionService.extractReportWorkbook("C:\\workspace\\personal project\\tracker-generator\\src\\test\\groovy\\October Report Jusqu'a 27-10.xlsx")
    }

    @Test
    void testExtractTemplate(){
        extractionService.extractTemplateSheets("C:\\workspace\\personal project\\tracker-generator\\src\\test\\groovy\\Template.xlsx")
        print(extractionService.getAgentList())

    }

    @Test
    void updateTeamMemberSheet() {
        extractionService.extractTemplateSheets("C:\\workspace\\personal project\\tracker-generator\\src\\test\\groovy\\Template.xlsx")
        extractionService.extractReportWorkbook("C:\\workspace\\personal project\\tracker-generator\\src\\test\\groovy\\October Report Jusqu'a 27-10.xlsx")

        extractionService.updateTeamMemberSheets(LocalDate.of(2025,10,1))
        extractionService.saveFile()
    }
}

