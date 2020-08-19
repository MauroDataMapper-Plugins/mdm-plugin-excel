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
package uk.ac.ox.softeng.maurodatamapper.plugins.excel

import org.apache.poi.ss.usermodel.Workbook
import org.springframework.beans.factory.annotation.Autowired
import org.springframework.core.GenericTypeResolver
import ox.softeng.metadatacatalogue.core.api.exception.ApiException
import ox.softeng.metadatacatalogue.core.catalogue.linkable.component.DataClassService
import ox.softeng.metadatacatalogue.core.catalogue.linkable.component.DataElementService
import ox.softeng.metadatacatalogue.core.catalogue.linkable.component.datatype.EnumerationTypeService
import ox.softeng.metadatacatalogue.core.catalogue.linkable.component.datatype.PrimitiveTypeService
import ox.softeng.metadatacatalogue.core.catalogue.linkable.datamodel.DataModel
import ox.softeng.metadatacatalogue.core.catalogue.linkable.datamodel.DataModelService
import ox.softeng.metadatacatalogue.core.spi.dataloader.DataModelDataLoaderPlugin
import ox.softeng.metadatacatalogue.core.user.CatalogueUser

import uk.ac.ox.softeng.maurodatamapper.plugins.excel.util.WorkbookHandler

/**
 * @since 15/02/2018
 */
abstract class AbstractExcelDataLoader<K extends uk.ac.ox.softeng.maurodatamapper.plugins.excel.row.StandardDataRow> implements DataModelDataLoaderPlugin, WorkbookHandler {

    @Autowired
    DataModelService dataModelService

    @Autowired
    PrimitiveTypeService primitiveTypeService

    @Autowired
    DataClassService dataClassService

    @Autowired
    EnumerationTypeService enumerationTypeService

    @Autowired
    DataElementService dataElementService

    abstract String getExcelFilename()

    abstract List<String> getExcelSheetNames()

    abstract String getDataModelName()

    abstract int getNumberOfHeaderRows()

    abstract int getIdCellNumber()

    abstract DataModel loadSheetDataRowsIntoDataModel(String sheetName, List<K> dataRows, DataModel dataModel, CatalogueUser catalogueUser)

    @Override
    List<DataModel> importData(CatalogueUser catalogueUser) throws ApiException {
        logger.info('Importing {} Model', name)

        Workbook workbook = null
        try {

            workbook = loadWorkbookFromFilename(getExcelFilename())

            DataModel dataModel = dataModelService.createDataModel(catalogueUser, getDataModelName(), getDescription(), getAuthor(),
                                                                   getOrganisation())
            Class<K> dataRowClass = (Class<K>) GenericTypeResolver.resolveTypeArgument(getClass(), AbstractExcelDataLoader)
            // Load each sheet
            excelSheetNames.each {sheetName ->
                logger.debug('Loading sheet {} from file {}', sheetName, getExcelFilename())
                List<K> dataRows = loadDataRowsFromSheet(workbook, dataRowClass, sheetName)
                if (!dataRows) return
                loadSheetDataRowsIntoDataModel(sheetName, dataRows, dataModel, catalogueUser)
            }

            dataModel.dataClasses ? [dataModel] : []

        } finally {
            closeWorkbook workbook
        }
    }

    List<K> loadDataRowsFromSheet(Workbook workbook, Class<K> dataRowClass, String sheetName) {
        loadDataRows(workbook, dataRowClass, getExcelFilename(), sheetName, getNumberOfHeaderRows(),
                     getIdCellNumber())
    }
}
