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
package uk.ac.ox.softeng.maurodatamapper.plugins.excel.simple

import uk.ac.ox.softeng.maurodatamapper.core.importer.ImporterService
import uk.ac.ox.softeng.maurodatamapper.core.rest.transport.importer.ImportParameterGroup
import uk.ac.ox.softeng.maurodatamapper.datamodel.DataModel

import groovy.transform.CompileStatic
import org.junit.Ignore
import org.junit.Test

import static org.junit.Assert.assertEquals
import static org.junit.Assert.assertTrue

@CompileStatic
@SuppressWarnings('SpellCheckingInspection')
class ExcelSimpleDataModelImporterProviderServiceTest extends BaseExcelSimpleDataModelImporterExporterProviderServiceTest {

    @Ignore
    void testSheetKey() {
        Map<String, String> results = [
            'My Model'                  : 'MM',
            'GEL: CancerSchema-v1.0.1'  : 'GC101',
            '[GEL: CancerSchema-v1.0.1]': 'GC101'
        ].each { assertEquals ExcelSimpleDataModelExporterProviderService.createSheetKey(it.key), it.value }
    }

    @Test
    void testImportParameters() {
        List<ImportParameterGroup> importParameters = applicationContext.getBean(ImporterService).describeImporterParams(importerInstance)

        assertEquals 'Number of parameter groups', 4, importParameters.size()
        assertEquals 'Number of source parameters', 1, importParameters[0].size()
        assertEquals 'Number of standard parameters', 5, importParameters[1].size()

        assertTrue 'Has Finalised parameter', importParameters[1].any {it.name == 'finalised'}
        assertTrue 'Has ImportAsNewDocumentationVersion parameter', importParameters[1].any {it.name == 'importAsNewDocumentationVersion'}
        assertTrue 'Has FolderId parameter', importParameters[1].any {it.name == 'folderId'}
        assertTrue 'Has ImportFile parameter', importParameters[0].first().name == 'importFile'
    }

    @Test
    void testSimpleImport() {
        DataModel dataModel = importSheet('simpleImport.simple.xlsx') as DataModel
        verifySimpleDataModel dataModel
        verifySimpleDataModelContent dataModel
    }

    @Test
    void testSimpleImportWithComplexMetadata() {
        DataModel dataModel = importSheet('simpleImportComplexMetadata.simple.xlsx') as DataModel
        verifySimpleDataModelWithComplexMetadata dataModel
        verifySimpleDataModelWithComplexMetadataContent dataModel
    }

    @Test
    void testMultipleDataModelImport() {

        List<DataModel> dataModels = importSheet('multiDataModelImport.simple.xlsx', 3) as List<DataModel>
        assertEquals 'Number of DataModels imported', 3, dataModels.size()

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

    private static DataModel findByLabel(List<DataModel> dataModels, String label) {
        DataModel.get(dataModels.find { it.label == label }.id)
    }

    private importSheet(String sheetFilename, int dataModelCount = 1) {
        if (dataModelCount == 1) return DataModel.get(importDomain(createImportParameters(sheetFilename)).id)
        importDomains(createImportParameters(sheetFilename), dataModelCount)
    }
}
