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

import uk.ac.ox.softeng.maurodatamapper.core.importer.ImporterService
import uk.ac.ox.softeng.maurodatamapper.core.rest.transport.importer.ImportParameterGroup
import uk.ac.ox.softeng.maurodatamapper.datamodel.DataModel
import uk.ac.ox.softeng.maurodatamapper.datamodel.DataModelType
import uk.ac.ox.softeng.maurodatamapper.datamodel.item.DataClass

import org.junit.Test

import groovy.transform.CompileStatic

import static org.junit.Assert.assertEquals
import static org.junit.Assert.assertNotNull
import static org.junit.Assert.assertTrue

@CompileStatic
@SuppressWarnings('SpellCheckingInspection')
class ExcelDataModelImporterProviderServiceTest extends BaseExcelDataModelImporterExporterProviderServiceTest {

    @Test
    void testImportParametersAreCorrect() {
        List<ImportParameterGroup> importParameters = applicationContext.getBean(ImporterService).describeImporterParams(importerInstance)

        assertEquals 'Number of parameter groups', 2, importParameters.size()
        assertEquals 'Number of standard parameters', 4, importParameters[0].size()
        assertEquals 'Number of source parameters', 1, importParameters[1].size()

        assertTrue 'Has Finalised parameter', importParameters[0].any { it.name == 'finalised' }
        assertTrue 'Has ImportAsNewDocumentationVersion parameter', importParameters[0].any { it.name == 'importAsNewDocumentationVersion' }
        assertTrue 'Has FolderId parameter', importParameters[0].any { it.name == 'folderId' }
        assertTrue 'Has ImportFile parameter', importParameters[1].first().name == 'importFile'
    }

    @Test
    void testSimpleImport() {
        DataModel importedDataModel = importDomain(createImportParameters('simpleImport.xlsx'))
        verifySimpleDataModel DataModel.get(importedDataModel.id)
        verifySimpleDataModelContent DataModel.get(importedDataModel.id)
    }

    @Test
    void testSimpleImportWithComplexMetadata() {
        DataModel importedDataModel = DataModel.get(importDomain(createImportParameters('simpleImportComplexMetadata.xlsx')).id)
        verifySimpleDataModel importedDataModel, 4
        verifyMetadata importedDataModel, 'ox.softeng.database|dialect', 'test'
        verifySimpleDataModelContentWithComplexMetadata importedDataModel
    }

    @Test
    void testMultipleDataModelImport() {
        importerInstance.saveDataModelsOnCreate = false

        List<DataModel> importedDataModels = importDomains(createImportParameters('multiDataModelImport.xlsx'), 3)
        assertEquals 'Number of DataModels imported', 3, importedDataModels.size()

        verifySimpleDataModel DataModel.get(importedDataModels.find { it.label == 'test' }?.id)
        verifySimpleDataModelContent DataModel.get(importedDataModels.find { it.label == 'test' }?.id)

        DataModel dataModel = DataModel.get(importedDataModels.find { it.label == 'Another Model' }?.id)
        assertNotNull 'Second model must exist', dataModel

        assertEquals 'DataModel name', 'Another Model', dataModel.label
        assertEquals 'DataModel description',
                     'this is another model which isn\'t overly complicated, and is designed for testing dataflows', dataModel.description
        assertEquals 'DataModel author', 'tester', dataModel.author
        assertEquals 'DataModel organisation', 'Oxford', dataModel.organisation
        assertEquals 'DataModel type', DataModelType.DATA_ASSET.toString(), dataModel.modelType

        assertEquals 'DataModel Metadata count', 2, dataModel.metadata.size()
        assertEquals 'Number of DataTypes', 6, dataModel.dataTypes.size()
        assertEquals 'Number of DataClasses', 3, dataModel.dataClasses.size()

        verifyMetadata dataModel, 'reviewed', 'no'
        verifyMetadata dataModel, 'distributed', 'no'

        verifyDataType dataModel, 'text', 'unlimited text'
        verifyDataType dataModel, 'string', 'just a plain old string'
        verifyDataType dataModel, 'int', 'a number'
        verifyDataType dataModel, 'table', 'a random thing'
        verifyDataType dataModel, 'child', null
        verifyEnumeration dataModel, 'yesno', 'an eumeration', [y: 'yes', n: 'no']

        DataClass top = verifyDataClass(dataModel, 'top', 3, 'tops description', 1, 1, ['extra info': 'some extra info'])
        verifyDataElement top, 'info', 'info description', 'string', 1, 1, ['different info': 'info']
        verifyDataElement top, 'another', null, 'int', 0, 1, ['extra info': 'some extra info', 'different info': 'info']
        verifyDataElement top, 'allinfo', 'all the info', 'text'

        DataClass child = verifyDataClass(top, 'child', 3, 'child description', 0)
        verifyDataElement child, 'info', 'childs info', 'string', 1, 1, ['extra info': 'some extra info']
        verifyDataElement child, 'does it work', 'I don\'t know', 'yesno', 1, -1
        verifyDataElement child, 'wrong order', 'should be fine', 'table'

        DataClass brother = verifyDataClass(top, 'brother', 3)
        verifyDataElement brother, 'info', 'brothers info', 'string', 0
        verifyDataElement brother, 'sibling', 'reference to the other child', 'child', 1, -1, ['extra info': 'some extra info']
        verifyDataElement brother, 'twin_sibling', 'reference to the other child', 'child', 0, -1

        dataModel = DataModel.get(importedDataModels.find { it.label == 'complex.xsd' }?.id)
        assertNotNull 'Third model must exist', dataModel
    }
}
