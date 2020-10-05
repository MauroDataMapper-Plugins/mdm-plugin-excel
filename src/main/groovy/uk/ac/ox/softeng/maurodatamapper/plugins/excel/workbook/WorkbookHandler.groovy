/*
 * Copyright 2020 University of Oxford
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 *
 * SPDX-License-Identifier: Apache-2.0
 */
package uk.ac.ox.softeng.maurodatamapper.plugins.excel.workbook

import uk.ac.ox.softeng.maurodatamapper.api.exception.ApiException
import uk.ac.ox.softeng.maurodatamapper.api.exception.ApiInternalException
import uk.ac.ox.softeng.maurodatamapper.plugins.excel.datarow.CellHandler
import uk.ac.ox.softeng.maurodatamapper.plugins.excel.datarow.StandardDataRow

import org.apache.poi.EncryptedDocumentException
import org.apache.poi.openxml4j.exceptions.InvalidFormatException
import org.apache.poi.openxml4j.util.ZipSecureFile
import org.apache.poi.ss.usermodel.Row
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.usermodel.Workbook
import org.apache.poi.ss.usermodel.WorkbookFactory
import org.apache.poi.ss.util.CellRangeAddress

import groovy.util.logging.Slf4j

@Slf4j
// @CompileStatic
trait WorkbookHandler extends CellHandler {

    Workbook loadWorkbookFromFilename(String filename) throws ApiException {
        loadWorkbookFromInputStream(filename, getClass().classLoader.getResourceAsStream(filename))
    }

    Workbook loadWorkbookFromInputStream(String filename, InputStream inputStream) throws ApiException {
        if (!inputStream) throw new ApiInternalException('EFS01', "No inputstream for ${filename}")
        try {
            ZipSecureFile.setMinInflateRatio(0)
            WorkbookFactory.create(inputStream)
        } catch (EncryptedDocumentException ignored) {
            throw new ApiInternalException('EFS02', "Excel file ${filename} could not be read as it is encrypted")
        } catch (InvalidFormatException ignored) {
            throw new ApiInternalException('EFS03', "Excel file ${filename} could not be read as it is not a valid format")
        } catch (IOException ex) {
            throw new ApiInternalException('EFS04', "Excel file ${filename} could not be read", ex)
        }
    }

    public <K extends StandardDataRow> List<K> loadDataRows(Workbook workbook, Class<K> dataRowClass, String filename, String sheetName,
                                                            int numberOfHeaderRows, int idCellNumber) throws ApiException {
        Sheet sheet = workbook.getSheet(sheetName)
        if (!sheet) throw new ApiInternalException('EFS05', "No sheet named [${sheetName}] present in [${filename}]")

        List<CellRangeAddress> mergedRegions = sheet.mergedRegions.findAll {
            it.firstColumn == idCellNumber && it.lastRow != it.firstRow
        }.sort { it.firstRow }
        log.trace('{} merged regions found, reduced down to {} useable regions', sheet.mergedRegions.size(), mergedRegions.size())

        List<K> dataRows = []
        Iterator<Row> rowIterator = sheet.rowIterator()
        numberOfHeaderRows.times { rowIterator.next() }

        while (rowIterator.hasNext()) {
            Row row = rowIterator.next()
            String cellValue = getCellValue(row.getCell(idCellNumber))
            if (!cellValue) continue

            log.trace('Examining row {} at {}', cellValue, row.rowNum)
            K dataRow = dataRowClass.getDeclaredConstructor().newInstance().tap { setAndInitialise row }
            dataRows << dataRow

            Integer lastRowIndexOfMerge = mergedRegions.find { it.firstRow == row.rowNum }?.lastRow
            if (!lastRowIndexOfMerge) continue

            log.trace('Merged region for {} from {} to {}(inclusive)', cellValue, row.rowNum, lastRowIndexOfMerge)
            dataRow.addToMergedContentRows(row)
            while (rowIterator.hasNext() && row.rowNum < lastRowIndexOfMerge) {
                row = rowIterator.next()
                dataRow.addToMergedContentRows(row)
            }
            log.trace('Data row has {} extra content', dataRow.mergedContentRows.size())
        }

        log.debug('Loaded {} data rows from sheet [{}] in [{}]', dataRows.size(), sheet.sheetName, filename)
        dataRows
    }

    void closeWorkbook(Workbook workbook) {
        try {
            workbook?.close()
        } catch (IOException ignored) {
            // Ignored
        }
    }
}
