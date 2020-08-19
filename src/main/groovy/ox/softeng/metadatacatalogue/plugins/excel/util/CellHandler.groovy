package ox.softeng.metadatacatalogue.plugins.excel.util

import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.DataFormatter
import org.apache.poi.ss.usermodel.Row
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.util.CellRangeAddress
import org.apache.poi.ss.util.CellUtil

/**
 * @since 14/03/2018
 */
trait CellHandler {

    DataFormatter getDataFormatter() {
        new DataFormatter()
    }

    @SuppressWarnings('GroovyAssignabilityCheck')
    void addCellToRow(Row row, int cellNum, Object content, boolean wrapText = false) {
        Cell cell = row.createCell(cellNum)
        cell.setCellValue(content)
        if (wrapText) {
            CellUtil.setCellStyleProperty(cell, CellUtil.WRAP_TEXT, wrapText)
            cell.cellStyle.setShrinkToFit(!wrapText)
        }
    }

    boolean cellHasContent(Row row, Integer columnIndex) {
        getCellValueAsString(row, columnIndex)
    }

    Cell findCell(Row row, String content) {
        row.find {getCellValueAsString(it) == content} as Cell
    }

    Cell findCell(Row row, String content, int column) {
        findCell(row, content, column, column)
    }

    Cell findCell(Row row, String content, int firstColumn, int lastColumn) {
        row.find {it.columnIndex >= firstColumn && it.columnIndex <= lastColumn && getCellValueAsString(it) == content} as Cell
    }

    String getCellValueAsString(Cell cell) {
        cell ? dataFormatter.formatCellValue(cell).replaceAll(/’/, '\'').replaceAll(/—/, '-').trim() : ''
    }

    String getCellValueAsString(Row row, Integer column) {
        row && column != null ? getCellValueAsString(row.getCell(column)) : ''
    }

    Integer getCellValueAsInteger(Cell cell, Integer defaultValue) {
        String cellValue = getCellValueAsString(cell)
        cellValue ? cellValue.toInteger() : defaultValue
    }

    CellRangeAddress getMergeRegion(Cell cell) {
        getMergeRegion(cell?.sheet, cell.rowIndex, cell.columnIndex)
    }

    CellRangeAddress getMergeRegion(Sheet sheet, int rowIndex, int columnIndex) {
        sheet?.mergedRegions?.find {it.isInRange(rowIndex, columnIndex)}
    }

    Cell buildHeaderCell(Row headerRow, int cellStyleColumn, String cellValue, boolean createMergeRegion = true) {
        buildHeaderCell(headerRow, headerRow.lastCellNum, cellStyleColumn, cellValue, createMergeRegion)
    }

    Cell buildHeaderCell(Row headerRow, int columnIndex, int cellStyleColumn, String cellValue, boolean createMergeRegion = true) {
        Cell headerCell = headerRow.createCell(columnIndex)
        headerCell.setCellValue(cellValue)

        Cell copyCell = headerRow.getCell(cellStyleColumn)

        CellRangeAddress mergedRegion = getMergeRegion(copyCell)

        if (createMergeRegion && mergedRegion) {
            CellRangeAddress newRegion = mergedRegion.copy()
            newRegion.firstColumn = headerCell.columnIndex
            newRegion.lastColumn = headerCell.columnIndex

            for (int r = newRegion.firstRow; r <= newRegion.lastRow; r++) {
                Row mr = headerRow.sheet.getRow(r)
                Cell hc = CellUtil.getCell(mr, headerCell.columnIndex)

                hc.setCellStyle(copyCell.cellStyle)
                if (r == newRegion.lastRow) {
                    Cell cc = mr.getCell(cellStyleColumn)
                    CellUtil.setCellStyleProperty(hc, CellUtil.BORDER_BOTTOM, cc.cellStyle.borderBottomEnum)
                }
            }
            headerRow.sheet.addMergedRegion(newRegion)
        } else {
            headerCell.setCellStyle(copyCell.cellStyle)
        }

        headerCell
    }

    String extractKey(String entry, int maxKeySize) {
        String[] comps = entry.trim().split(' ')
        if (!comps) return null
        String key = comps[0].trim()
        if (!key) return null
        if (maxKeySize == -1) return key

        if (key.length() <= maxKeySize) return key

        try {
            Integer.parseInt(key)
            return key
        } catch (NumberFormatException ignored) {
        }
        if (key.startsWith('+') || key.startsWith('-')) return key

        int i = firstIndexOfUppercaseLetter(entry)
        i == -1 ? entry : entry.charAt(i).toString()
    }

    String extractValue(String entry, String key) {
        String[] comps = entry.trim().split(' ')

        // First entry is key, rest is value
        if (comps[0].trim() == key) {
            StringBuilder sb = new StringBuilder()
            for (int j = 1; j < comps.length; j++) {
                sb.append(comps[j]).append(' ')
            }
            return sb.toString().trim().toLowerCase().capitalize()
        }

        // Key is capital letter inside entry, so fix the case of the entry
        entry.trim().toLowerCase().capitalize()
    }

    int firstIndexOfUppercaseLetter(String str) {
        for (int i = 1; i < str.length(); i++) {
            if (Character.isUpperCase(str.charAt(i))) {
                return i
            }
        }
        Character.isUpperCase(str.charAt(0)) ? 0 : -1
    }

    Map<String, String> getEnumValues(String format, int discardLines, int maxKeySize) {
        Map<String, String> enumMap = [:]
        String[] parts = format.split('\n')
        for (int i = discardLines; i < parts.length; i++) {
            if (parts[i]) {
                String key = extractKey(parts[i], maxKeySize)
                if (key) {
                    String value = extractValue(parts[i], key)
                    if (value) enumMap[key] = value
                }
            }
        }
        enumMap.findAll {it.key}
    }
}