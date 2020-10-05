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

import uk.ac.ox.softeng.maurodatamapper.api.exception.ApiException
import uk.ac.ox.softeng.maurodatamapper.datamodel.DataModel
import uk.ac.ox.softeng.maurodatamapper.plugins.testing.utils.user.IntegrationTestUser

import com.google.common.base.Strings
import org.junit.Before
import org.junit.Test

import groovy.transform.CompileStatic
import groovy.util.logging.Slf4j

import java.nio.file.Files
import java.nio.file.Path
import java.nio.file.Paths

import static org.junit.Assert.assertFalse
import static org.junit.Assert.assertNotNull

@Slf4j
@CompileStatic
@SuppressWarnings('SpellCheckingInspection')
class ExcelDataModelExporterProviderServiceTest extends BaseExcelDataModelImporterExporterProviderServiceTest {

    private static final String CHARSET = 'ISO-8859-1'
    private static final String EXPORT_FILEPATH = 'build/tmp/'

    @Before
    void disableDataModelSavingOnCreate() {
        importerInstance.saveDataModelsOnCreate = false
    }

    @Test
    void testSimpleExport() {
        DataModel importedDataModel = testExportViaImport('simpleImport.xlsx', 'simpleImport_export.xlsx').first()
        verifySimpleDataModel importedDataModel
        verifySimpleDataModelContent importedDataModel
    }

    @Test
    void testSimpleExportWithComplexMetadata() {
        DataModel importedDataModel = testExportViaImport('simpleImportComplexMetadata.xlsx', 'simpleImportComplexMetadata_export.xlsx').first()
        verifySimpleDataModel importedDataModel, 4
        verifyMetadata importedDataModel, 'ox.softeng.database|dialect', 'test'
        verifySimpleDataModelContentWithComplexMetadata importedDataModel
    }

    @Test
    void testMultipleDataModelExport() {
        List<DataModel> importedDataModels = testExportViaImport('multiDataModelImport.xlsx', 'multiDataModelImport_export.xlsx', 3)
        verifySimpleDataModel importedDataModels.find { it.label == 'test' }
        verifySimpleDataModelContent importedDataModels.find { it.label == 'test' }
    }

    private List<DataModel> testExportViaImport(String importFilename, String exportFilename, int expectedSize = 1) throws IOException, ApiException {
        // Import DataModels under test first
        List<DataModel> importedDataModels = importDomains(createImportParameters(importFilename), expectedSize)
        log.debug('DataModel(s) to export: {}', importedDataModels.size() == 1 ? importedDataModels.first().id : importedDataModels.id)

        // Export what has been saved into the database
        log.info('>>> Exporting {}', importedDataModels.size() == 1 ? 'Single' : 'Multiple')
        ExcelDataModelExporterProviderService exporterService = applicationContext.getBean(ExcelDataModelExporterProviderService)
        ByteArrayOutputStream exportedDataModels = importedDataModels.size() == 1
            ? exporterService.exportDomain(IntegrationTestUser.instance, importedDataModels.first().id)
            : exporterService.exportDomains(IntegrationTestUser.instance, importedDataModels.id)

        assertNotNull 'Should have exported model(s)', exportedDataModels
        assertFalse 'Should have exported model string', Strings.isNullOrEmpty(exportedDataModels.toString(CHARSET))

        Path exportFilepath = Paths.get(EXPORT_FILEPATH, exportFilename)
        Files.write(exportFilepath, exportedDataModels.toByteArray())

        // Test using the exported DataModels instead of the ones from the first import
        log.info('>>> Importing')
        importDomains(createImportParameters(exportFilepath), expectedSize, false)
    }
}
