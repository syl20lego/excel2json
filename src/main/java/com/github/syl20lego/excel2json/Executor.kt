package com.github.syl20lego.excel2json

import com.fasterxml.jackson.core.util.DefaultPrettyPrinter
import com.fasterxml.jackson.databind.ObjectMapper
import com.fasterxml.jackson.databind.SerializationFeature
import org.apache.poi.ss.usermodel.*

import java.awt.*
import java.io.*
import java.util.*

data class FileInfo(val directory: String?, val file: String?)

private fun open(): FileInfo {
    val dialog = FileDialog(null as Frame?, "Select File to Open")
    dialog.mode = FileDialog.LOAD
    dialog.setFilenameFilter { dir, name -> name.endsWith(".xlsx") || name.endsWith(".xls") }
    dialog.isVisible = true
    return FileInfo(dialog.directory, dialog.file)
}

private fun cellValue(cell: Cell?): String {

    cell ?: return ""
    when (cell.cellTypeEnum) {
        CellType.STRING -> return cell.richStringCellValue.string
        CellType.FORMULA -> return cell.cellFormula
        CellType.BLANK -> return ""
        CellType.NUMERIC -> return if (DateUtil.isCellDateFormatted(cell)) {
            val dataFormatter = DataFormatter()
            dataFormatter.formatCellValue(cell)
        } else {
            if (cell.numericCellValue % 1 == 0.0) {
                Integer.valueOf(java.lang.Double.valueOf(cell.numericCellValue).toInt())!!.toString()
            } else {
                java.lang.Double.valueOf(cell.numericCellValue).toString()
            }
        }
        else -> return ""
    }
}

private fun writeJson(directory: String?, file: String, sheets: LinkedHashMap<String, List<*>>) {
    val jsonFile = file.substring(0, file.lastIndexOf(".")) + ".json"
    val mapper = ObjectMapper()
    mapper.enable(SerializationFeature.INDENT_OUTPUT)
    val writer = mapper.writer(DefaultPrettyPrinter())
    val fileWriter = BufferedWriter(OutputStreamWriter(FileOutputStream(directory + jsonFile), "UTF-8"))
    writer.writeValue(fileWriter, sheets)
}

fun isNotEmptyRow(row: Row): Boolean {
    for (cellNum in row.firstCellNum until row.lastCellNum) {
        val cell = row.getCell(cellNum)
        if (cell != null &&  cell.cellTypeEnum != CellType.BLANK && cell.toString().isNotEmpty()) {
            return true
        }
    }
    return false
}

object Executor {

    @JvmStatic
    fun main(args: Array<String>) {
        val (directory, file) = open()
        file ?: return

        val sheets = LinkedHashMap<String, List<*>>()
        WorkbookFactory.create(File(directory + file), null, true)
                .use {
                    it.forEach { sheet ->
                        val list = ArrayList<Map<String, Any>>()
                        sheets[sheet.sheetName] = list
                        val header = ArrayList<String>()
                        sheet.forEach { row ->
                            val outputRow = LinkedHashMap<String, Any>()
                            if (header.isEmpty()) {
                                row.forEach { cell -> header.add(cell.toString()) }
                            } else if (isNotEmptyRow(row)){
                                for (i in 0 until header.size) {
                                    outputRow[header[i]] = cellValue(row.getCell(i))
                                }
                            }
                            if (!outputRow.isEmpty()) {
                                list.add(outputRow)
                            }
                        }
                    }
                }
        writeJson(directory, file, sheets)
        println("Done")
    }
}
