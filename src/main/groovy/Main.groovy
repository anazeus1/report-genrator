import javax.swing.JFrame
import javax.swing.JLabel
import javax.swing.JTextField
import javax.swing.JButton
import javax.swing.JOptionPane
import javax.swing.JPanel
import java.awt.FlowLayout
import java.awt.GridLayout
import java.time.LocalDate
import java.time.format.DateTimeFormatter
import java.time.format.DateTimeParseException

/**
 * Company: PSI Software SE
 * @author: Anas Gharbi
 */

class Main {
    static void main(String[] args) {
        ExtractionService extractionService = new ExtractionService()
        //ui components
        JFrame frame = new JFrame("Tracker Report Generator")
        JPanel panel = new JPanel(new GridLayout(3, 1))

        JPanel row1 = new JPanel(new FlowLayout(FlowLayout.LEFT))
        JLabel wartungsListeLabel = new JLabel("Pleas give the Report path.")
        JTextField wartungsListeField = new JTextField(50)
        JButton confirmWartungslisteButton = new JButton("Confirm")
        row1.add(wartungsListeLabel)
        row1.add(wartungsListeField)
        row1.add(confirmWartungslisteButton)

        JPanel row2 = new JPanel(new FlowLayout(FlowLayout.LEFT))
        JLabel templatePathLabel = new JLabel("Please give the Tracker Templat path.")
        JTextField templatePathField = new JTextField(50)
        JButton confirmTemplateButton = new JButton("Confirm")
        row2.add(templatePathLabel)
        row2.add(templatePathField)
        row2.add(confirmTemplateButton)

        JPanel row3 = new JPanel(new FlowLayout(FlowLayout.LEFT))
        JLabel dateLabel = new JLabel("Give the Start Date")
        JTextField dateField = new JTextField(10)
        JButton generateButton = new JButton("Generate")
        row3.add(dateLabel)
        row3.add(dateField)
        row3.add(generateButton)

        confirmWartungslisteButton.addActionListener { l ->
            try {
                extractionService.extractReportWorkbook(wartungsListeField.getText())
            }
            catch (FileNotFoundException e) {
                JOptionPane.showMessageDialog(null,
                        "Report file not Found.",
                        "Error",
                        JOptionPane.ERROR_MESSAGE)
            }
        }

        confirmTemplateButton.addActionListener { l ->
            {
                try {
                    extractionService.extractTemplateSheets(templatePathField.getText())
                }
                catch (FileNotFoundException e) {
                    JOptionPane.showMessageDialog(null,
                            "Template file not Found.",
                            "Error",
                            JOptionPane.ERROR_MESSAGE)
                }
            }
        }
        generateButton.addActionListener(l -> {
            String date = dateField.getText()
            LocalDate ym = getDate(date)
            if (extractionService.teamListSheet == null) {
                JOptionPane.showMessageDialog(null,
                        "The Report file is not extracted yet."+
                                " Please make sure you clicked confirm.",
                        "Fehler",
                        JOptionPane.ERROR_MESSAGE)
            } else if (extractionService.templateWorkbook == null) {
                JOptionPane.showMessageDialog(null,
                        "The Template file is not extracted yet."+
                                " Please make sure you clicked confirm.",
                        "Error",
                        JOptionPane.ERROR_MESSAGE)

            } else if (ym != null) {
                extractionService.updateTeamMemberSheets(ym)
                JOptionPane.showMessageDialog(null,
                        "Excel Sheets are being generated",
                        "Succes!",
                        JOptionPane.INFORMATION_MESSAGE)
            }
        })

        panel.add(row1)
        panel.add(row2)
        panel.add(row3)

        frame.setContentPane(panel)
        frame.setLocationRelativeTo(null)
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE)
        frame.pack()
        frame.setVisible(true)
    }

    static LocalDate getDate(String input) {
        try {
            def fullInput = "01-${input}"
            DateTimeFormatter formatter = DateTimeFormatter.ofPattern("dd-MM-yyyy")
            return LocalDate.parse(fullInput, formatter)
        } catch (DateTimeParseException e) {
            JOptionPane.showMessageDialog(null,
                    "Ungültiges Datumsformat. Bitte verwenden Sie MM-JJJJ (z. B. 06-2025).",
                    "Fehler",
                    JOptionPane.ERROR_MESSAGE)
            return null
        }
    }
}
