/*
 * Copyright 2020-2022 University of Oxford and Health and Social Care Information Centre, also known as NHS Digital
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
package uk.ac.ox.softeng.maurodatamapper.plugins.excel.datamodel.provider.exporter

import uk.ac.ox.softeng.maurodatamapper.core.provider.importer.parameter.FileParameter
import uk.ac.ox.softeng.maurodatamapper.datamodel.DataModel
import uk.ac.ox.softeng.maurodatamapper.datamodel.item.DataClass
import uk.ac.ox.softeng.maurodatamapper.datamodel.provider.exporter.DataModelExporterProviderService
import uk.ac.ox.softeng.maurodatamapper.datamodel.provider.importer.DataModelImporterProviderService
import uk.ac.ox.softeng.maurodatamapper.datamodel.provider.importer.DataModelJsonImporterService
import uk.ac.ox.softeng.maurodatamapper.datamodel.provider.importer.DataModelXmlImporterService
import uk.ac.ox.softeng.maurodatamapper.datamodel.provider.importer.parameter.DataModelFileImporterProviderServiceParameters
import uk.ac.ox.softeng.maurodatamapper.plugins.excel.BaseExcelDataModelImporterExporterProviderServiceSpec
import uk.ac.ox.softeng.maurodatamapper.plugins.excel.datamodel.provider.importer.ExcelDataModelImporterProviderService
import uk.ac.ox.softeng.maurodatamapper.util.Utils

import grails.gorm.transactions.Rollback
import grails.testing.mixin.integration.Integration
import groovy.util.logging.Slf4j
import spock.lang.Ignore

import java.nio.file.Files
import java.nio.file.Path
import java.nio.file.Paths

import static org.junit.Assert.assertEquals
import static org.junit.Assert.assertNotNull

@Slf4j
@Integration
@Rollback
class ExcelDataModelExporterProviderServiceSpec extends BaseExcelDataModelImporterExporterProviderServiceSpec {

    DataModelXmlImporterService dataModelXmlImporterService
    ExcelDataModelExporterProviderService excelDataModelExporterProviderService
    ExcelDataModelImporterProviderService excelDataModelImporterProviderService
    DataModelJsonImporterService dataModelJsonImporterService

    @Override
    DataModelImporterProviderService getDataModelImporterProviderService() {
        excelDataModelImporterProviderService
    }

    @Override
    DataModelExporterProviderService getDataModelExporterProviderService() {
        excelDataModelExporterProviderService
    }

    def 'testSimpleExport'() {
        given:
        setupDomainData()

        when:
        DataModel dataModel = importThenExportWorkbook('simpleImport.xlsx', 'simpleImport_export.xlsx') as DataModel

        then:
        verifySimpleDataModel dataModel
        verifySimpleDataModelContent dataModel
    }

    def 'testSimpleExportWithComplexMetadata'() {
        given:
        setupDomainData()

        when:
        DataModel dataModel = importThenExportWorkbook('simpleImportComplexMetadata.xlsx', 'simpleImportComplexMetadata_export.xlsx') as DataModel

        then:
        verifySimpleDataModelWithComplexMetadata dataModel
        verifySimpleDataModelWithComplexMetadataContent dataModel
    }

    def 'testMultipleDataModelExport'() {
        given:
        setupDomainData()

        when:
        List<DataModel> dataModels = importThenExportWorkbook('multiDataModelImport.xlsx', 'multiDataModelImport_export.xlsx') as List<DataModel>
        DataModel simpleDataModel = findByLabel(dataModels, 'test')

        then:
        verifySimpleDataModel simpleDataModel
        verifySimpleDataModelContent simpleDataModel

        when:
        DataModel dataFlowDataModel = findByLabel(dataModels, 'Another Model')

        then:
        verifyDataFlowDataModel dataFlowDataModel
        verifyDataFlowDataModelContent dataFlowDataModel

        when:
        DataModel complexDataModel = findByLabel(dataModels, 'complex.xsd')

        then:
        verifyComplexDataModel complexDataModel
        verifyComplexDataModelContent complexDataModel
    }

    def 'testDefaultMultiplicitiesIfUnset'() {
        given:
        setupDomainData()

        when:
        // Import test Data Model from XML file
        Path importFilepath = Paths.get(IMPORT_FILEPATH, 'unsetMultiplicities.xml')
        DataModelFileImporterProviderServiceParameters importParameters = new DataModelFileImporterProviderServiceParameters().tap {
            importFile = new FileParameter(importFilepath.toString(), 'text/xml', Files.readAllBytes(importFilepath))
        }
        DataModel dataModel = dataModelXmlImporterService.importModel(admin, importParameters)

        // Export Data Model as Excel file
        ByteArrayOutputStream exportedDataModel = excelDataModelExporterProviderService.exportDataModel(admin, dataModel, [:])
        Path exportFilepath = Paths.get(EXPORT_FILEPATH, 'unsetMultiplicities_export.xlsx')
        Files.write(exportFilepath, exportedDataModel.toByteArray())

        // Import Data Model from Excel file to test exporting
        DataModel testDataModel = importAndValidateModel(createImportParameters(exportFilepath))

        DataClass testDataClass = testDataModel.dataClasses.find { it.label == 'Data Class with Unset Multiplicities' }

        then:
        assertNotNull "DataClass ${testDataClass.label} must exist", testDataClass
        assertEquals "DataClass ${testDataClass.label} minMultiplicity must default to 0", 0, testDataClass.minMultiplicity
        assertEquals "DataClass ${testDataClass.label} maxMultiplicity must default to 0", 0, testDataClass.maxMultiplicity
    }

    @Ignore('too long to run')
    void 'test import and export of badgernet to check for speed'(){

        given:
        setupDomainData()

        Path importFilepath = Paths.get(IMPORT_FILEPATH, 'badgernet.json')
        DataModelFileImporterProviderServiceParameters importParameters = new DataModelFileImporterProviderServiceParameters().tap {
            importFile = new FileParameter(importFilepath.toString(), 'text/json', Files.readAllBytes(importFilepath))
        }
        DataModel dataModel = dataModelJsonImporterService.importDomain(admin, importParameters)
        validateAndSave(dataModel)

        when:
        long start = System.currentTimeMillis()
        Path exportPath = exportDataModel(dataModel,'badgernet.xlsx')
        log.warn('Took {}', Utils.timeTaken(start))

        then:
        Files.exists(exportPath)
    }
}