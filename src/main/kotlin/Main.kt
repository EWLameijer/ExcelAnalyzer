import org.apache.poi.ss.usermodel.*
import org.apache.poi.xssf.usermodel.XSSFSheet
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.FileInputStream
import java.time.LocalDate
import java.time.format.DateTimeFormatter
import java.time.format.DateTimeParseException

fun main() {
    val contents = getSheetContents("GR50.xlsx", "Blad1")
    val schedule = extractSchedule(contents)
    display(schedule)
}

private fun display(schedule: MutableMap<LocalDate, String>) {
    val dateFormatter = DateTimeFormatter.ofPattern("EEE dd LLL")
    // show all dates, getting rid of dummy Jan 1 2000
    for (date in schedule.keys.sorted().drop(1)) {
        val dateAsString = date.format(dateFormatter)
        println("$dateAsString: ${schedule[date]}")
    }
}

private fun extractSchedule(contents: MutableMap<Coordinate, String?>): MutableMap<LocalDate, String> {
    var dateColumn = 'A'
    var lastDate: LocalDate = LocalDate.of(2000, 1, 1)
    val schedule = mutableMapOf<LocalDate, String>()

    for (row in 1..contents.keys.maxOf { it.row }) {
        for (column in 'A'..'Z') {
            val chosenCoordinate = Coordinate(row, column)
            val cellContents = contents[chosenCoordinate]
            val date = dateOrNull(cellContents)
            if (date != null) {
                dateColumn = column
                lastDate = date
            } else if (column == dateColumn + 4 && cellContents!!.isNotBlank() && cellContents.trim() != "KLAS VRIJ!") {
                schedule[lastDate] = cellContents
            }
        }
    }
    return schedule
}

data class Coordinate(val row: Int, val column: Char)

private fun getSheetContents(filename: String, sheetName: String): MutableMap<Coordinate, String?> {
    val fis = FileInputStream(filename)
    val workbook = XSSFWorkbook(fis)
    val sheet: XSSFSheet = workbook.getSheet(sheetName)
    val iterator: Iterator<Row> = sheet.rowIterator()
    val contents = mutableMapOf<Coordinate, String?>();
    var rowNo = 0
    while (iterator.hasNext()) {
        val row = iterator.next()
        rowNo++;
        for (i in 0 until row.physicalNumberOfCells) {
            val cell = row.getCell(i)
            contents[Coordinate(rowNo, 'A' + i)] = cell?.toString();
        }
    }
    return contents
}

private fun dateOrNull(cellContents: String?): LocalDate? =
    try {
        LocalDate.parse(cellContents, DateTimeFormatter.ofPattern("dd-LLL-yyyy"))
    } catch (e: DateTimeParseException) {
        null
    }