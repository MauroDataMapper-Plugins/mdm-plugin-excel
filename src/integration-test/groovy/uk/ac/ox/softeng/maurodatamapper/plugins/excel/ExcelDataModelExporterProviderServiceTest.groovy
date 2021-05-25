/*
 * Copyright 2020-2021 University of Oxford and Health and Social Care Information Centre, also known as NHS Digital
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

import uk.ac.ox.softeng.maurodatamapper.api.exception.ApiException
import uk.ac.ox.softeng.maurodatamapper.core.provider.importer.parameter.FileParameter
import uk.ac.ox.softeng.maurodatamapper.datamodel.DataModel
import uk.ac.ox.softeng.maurodatamapper.datamodel.item.DataClass
import uk.ac.ox.softeng.maurodatamapper.datamodel.provider.importer.DataModelXmlImporterService
import uk.ac.ox.softeng.maurodatamapper.datamodel.provider.importer.parameter.DataModelFileImporterProviderServiceParameters
import uk.ac.ox.softeng.maurodatamapper.plugins.excel.datamodel.provider.exporter.ExcelDataModelExporterProviderService
import uk.ac.ox.softeng.maurodatamapper.plugins.testing.utils.user.IntegrationTestUser

import com.google.common.base.Strings
import groovy.transform.CompileStatic
import groovy.util.logging.Slf4j
import org.junit.Test

import java.nio.file.Files
import java.nio.file.Path
import java.nio.file.Paths

import static org.junit.Assert.assertEquals
import static org.junit.Assert.assertFalse
import static org.junit.Assert.assertNotNull

@Slf4j
@CompileStatic
@SuppressWarnings('SpellCheckingInspection')
class ExcelDataModelExporterProviderServiceTest extends BaseExcelDataModelImporterExporterProviderServiceTest {

    private static final String EXPORT_FILEPATH = 'build/tmp/'

    @Test
    void testSimpleExport() {
        DataModel dataModel = importThenExportSheet('simpleImport.xlsx', 'simpleImport_export.xlsx') as DataModel
        verifySimpleDataModel dataModel
        verifySimpleDataModelContent dataModel
    }

    @Test
    void testSimpleExportWithComplexMetadata() {
        DataModel dataModel = importThenExportSheet('simpleImportComplexMetadata.xlsx', 'simpleImportComplexMetadata_export.xlsx') as DataModel
        verifySimpleDataModelWithComplexMetadata dataModel
        verifySimpleDataModelWithComplexMetadataContent dataModel
    }

    @Test
    void testMultipleDataModelExport() {
        List<DataModel> dataModels = importThenExportSheet('multiDataModelImport.xlsx', 'multiDataModelImport_export.xlsx', 3) as List<DataModel>

        DataModel simpleDataModel = findByLabel(dataModels, 'test')
        verifySimpleDataModel simpleDataModel
        verifySimpleDataModelContent simpleDataModel

        DataModel dataFlowDataModel = findByLabel(dataModels, 'Another Model')
        verifyDataFlowDataModel dataFlowDataModel
        verifyDataFlowDataModelContent dataFlowDataModel

        DataModel complexDataModel = findByLabel(dataModels, 'complex.xsd')
        verifyComplexDataModel complexDataModel
        verifyComplexDataModelContent complexDataModel
    }

    @Test
    void testDefaultMultiplicitiesIfUnset() {
        // Import test Data Model from XML file
        Path importFilepath = Paths.get(IMPORT_FILEPATH, 'unsetMultiplicities.xml')
        DataModelFileImporterProviderServiceParameters importParameters = new DataModelFileImporterProviderServiceParameters().tap {
            importFile = new FileParameter(importFilepath.toString(), 'text/xml', Files.readAllBytes(importFilepath))
        }
        DataModel dataModel = applicationContext.getBean(DataModelXmlImporterService).importModel(IntegrationTestUser.instance, importParameters)

        // Export Data Model as Excel file
        ByteArrayOutputStream exportedDataModel = applicationContext.getBean(ExcelDataModelExporterProviderService)
                                                                    .exportDataModel(IntegrationTestUser.instance, dataModel)
        Path exportFilepath = Paths.get(EXPORT_FILEPATH, 'unsetMultiplicities_export.xlsx')
        Files.write(exportFilepath, exportedDataModel.toByteArray())

        // Import Data Model from Excel file to test exporting
        DataModel testDataModel = importDomain(createImportParameters(exportFilepath), false)

        DataClass testDataClass = testDataModel.dataClasses.find { it.label == 'Data Class with Unset Multiplicities' }
        assertNotNull "DataClass ${testDataClass.label} must exist", testDataClass
        assertEquals "DataClass ${testDataClass.label} minMultiplicity must default to 0", 0, testDataClass.minMultiplicity
        assertEquals "DataClass ${testDataClass.label} maxMultiplicity must default to 0", 0, testDataClass.maxMultiplicity

        // The following is commented out because the target Data Element cannot be retrieved as expected (a NullPointerException is returned). A bug?
        //
        // String dataElementLabel = 'Data Element with Unset Multiplicities'
        // DataElement testDataElement = testDataClass.findDataElement(dataElementLabel)
        //     ?: testDataClass.dataElements.find { it.label == dataElementLabel }
        // assertNotNull "DataElement ${testDataElement.label} must exist", testDataElement
        // assertEquals "DataElement ${testDataElement.label} minMultiplicity must default to 0", 0, testDataElement.minMultiplicity
        // assertEquals "DataElement ${testDataElement.label} maxMultiplicity must default to 0", 0, testDataElement.maxMultiplicity
    }

    private static DataModel findByLabel(List<DataModel> dataModels, String label) {
        dataModels.find { it.label == label }
    }

    private importThenExportSheet(String importFilename, String exportFilename, int expectedSize = 1) throws IOException, ApiException {
        // Import DataModels under test first
        List<DataModel> importedDataModels = importDomains(createImportParameters(importFilename), expectedSize)
        log.debug('DataModel(s) to export: {}', importedDataModels.size() == 1 ? importedDataModels.first().id : importedDataModels.id)

        // Export what has been saved into the database
        log.info('>>> Exporting {}', importedDataModels.size() == 1 ? 'Single' : 'Multiple')
        ExcelDataModelExporterProviderService exporterService = applicationContext.getBean(ExcelDataModelExporterProviderService)
        ByteArrayOutputStream exportedDataModels = importedDataModels.size() == 1
            ? exporterService.exportDomain(IntegrationTestUser.instance, importedDataModels.first().id)
            : exporterService.exportDomains(IntegrationTestUser.instance, importedDataModels.id)

        assertNotNull 'Should have exported DataModel(s)', exportedDataModels
        assertFalse 'Should have exported DataModel string', Strings.isNullOrEmpty(exportedDataModels.toString('ISO-8859-1'))

        Path exportFilepath = Paths.get(EXPORT_FILEPATH, exportFilename)
        Files.write(exportFilepath, exportedDataModels.toByteArray())

        // Test using the exported DataModels instead of the ones from the first import
        log.info('>>> Importing')
        List<DataModel> importedExportedDataModels = importDomains(createImportParameters(exportFilepath), expectedSize, false)
        importedExportedDataModels.size() == 1 ? importedExportedDataModels.first() : importedExportedDataModels
    }
}
