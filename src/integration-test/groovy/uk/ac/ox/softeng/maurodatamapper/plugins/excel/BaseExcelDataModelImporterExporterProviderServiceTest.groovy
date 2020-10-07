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

import uk.ac.ox.softeng.maurodatamapper.core.facet.Metadata
import uk.ac.ox.softeng.maurodatamapper.core.model.CatalogueItem
import uk.ac.ox.softeng.maurodatamapper.core.provider.importer.parameter.FileParameter
import uk.ac.ox.softeng.maurodatamapper.datamodel.DataModel
import uk.ac.ox.softeng.maurodatamapper.datamodel.DataModelService
import uk.ac.ox.softeng.maurodatamapper.datamodel.DataModelType
import uk.ac.ox.softeng.maurodatamapper.datamodel.item.DataClass
import uk.ac.ox.softeng.maurodatamapper.datamodel.item.DataElement
import uk.ac.ox.softeng.maurodatamapper.datamodel.item.datatype.DataType
import uk.ac.ox.softeng.maurodatamapper.datamodel.item.datatype.EnumerationType
import uk.ac.ox.softeng.maurodatamapper.datamodel.item.datatype.ReferenceType
import uk.ac.ox.softeng.maurodatamapper.plugins.testing.utils.BaseImportPluginTest

import java.nio.file.Files
import java.nio.file.Path
import java.nio.file.Paths

import static org.junit.Assert.assertEquals
import static org.junit.Assert.assertNotNull
import static org.junit.Assert.fail

abstract class BaseExcelDataModelImporterExporterProviderServiceTest
    extends BaseImportPluginTest<DataModel, ExcelDataModelFileImporterProviderServiceParameters, ExcelDataModelImporterProviderService> {

    private static final String EXCEL_FILETYPE = 'application/vnd.ms-excel'
    private static final String IMPORT_FILEPATH = 'src/integration-test/resources/'

    DataModelService dataModelService = applicationContext.getBean(DataModelService)

    @Override
    DataModel saveDomain(DataModel domain) {
        dataModelService.saveWithBatching(domain)
    }

    protected ExcelDataModelFileImporterProviderServiceParameters createImportParameters(String importFilename) throws IOException {
        Path importFilepath = Paths.get(IMPORT_FILEPATH, importFilename)
        if (!Files.exists(importFilepath)) fail("File ${importFilename} cannot be found")
        createImportParameters(importFilepath)
    }

    protected ExcelDataModelFileImporterProviderServiceParameters createImportParameters(Path importFilepath) throws IOException {
        new ExcelDataModelFileImporterProviderServiceParameters(finalised: false).tap {
            importFile = new FileParameter(importFilepath.toString(), EXCEL_FILETYPE, Files.readAllBytes(importFilepath))
        }
    }

    protected void verifySimpleDataModel(DataModel dataModel, int metadataCount = 3) {
        assertNotNull 'Simple DataModel must exist', dataModel

        assertEquals 'DataModel Name', 'test', dataModel.label
        assertEquals 'DataModel Description', 'this is a simple test with only 1 DataModel', dataModel.description
        assertEquals 'DataModel Author', 'tester', dataModel.author
        assertEquals 'DataModel Organisation', 'Oxford', dataModel.organisation
        assertEquals 'DataModel Type', DataModelType.DATA_STANDARD.toString(), dataModel.modelType

        assertEquals 'DataModel Metadata count', metadataCount, dataModel.metadata.size()
        assertEquals 'DataModel number of DataTypes', 6, dataModel.dataTypes.size()
        assertEquals 'DataModel number of DataClasses', 3, dataModel.dataClasses.size()

        verifyMetadata dataModel, 'reviewed', 'yes'
        verifyMetadata dataModel, 'approved', 'yes'
        verifyMetadata dataModel, 'distributed', 'no'

        verifyDataType dataModel, 'int', 'a number'
        verifyDataType dataModel, 'string', 'just a plain old string'
        verifyDataType dataModel, 'table', 'a random thing'

        verifyReferenceType dataModel, 'child', null, 'top|child'
        verifyEnumerationType dataModel, 'yesno', 'an eumeration', [y: 'yes', n: 'no']
        verifyEnumerationType dataModel, 'possibly', null, ['0': 'lazy', '1': 'not lazy', '2': 'very lazy']
    }

    protected void verifySimpleDataModelWithComplexMetadata(DataModel dataModel) {
        verifySimpleDataModel dataModel, 4
        verifyMetadata dataModel, 'ox.softeng.database|dialect', 'test'
    }

    protected void verifyDataFlowDataModel(DataModel dataModel) {
        assertNotNull 'Second DataModel must exist', dataModel

        assertEquals 'DataModel Name', 'Another Model', dataModel.label
        assertEquals 'DataModel Description',
                     'this is another model which isn\'t overly complicated, and is designed for testing dataflows', dataModel.description
        assertEquals 'DataModel Author', 'tester', dataModel.author
        assertEquals 'DataModel Organisation', 'Oxford', dataModel.organisation
        assertEquals 'DataModel Type', DataModelType.DATA_ASSET.toString(), dataModel.modelType

        assertEquals 'DataModel Metadata count', 2, dataModel.metadata.size()
        assertEquals 'DataModel number of DataTypes', 6, dataModel.dataTypes.size()
        assertEquals 'DataModel number of DataClasses', 3, dataModel.dataClasses.size()

        verifyMetadata dataModel, 'reviewed', 'no'
        verifyMetadata dataModel, 'distributed', 'no'

        verifyDataType dataModel, 'int', 'a number'
        verifyDataType dataModel, 'string', 'just a plain old string'
        verifyDataType dataModel, 'text', 'unlimited text'
        verifyDataType dataModel, 'table', 'a random thing'

        verifyReferenceType dataModel, 'child', null, 'top|child'
        verifyEnumerationType dataModel, 'yesno', 'an eumeration', [y: 'yes', n: 'no']
    }

    protected void verifySimpleDataModelContent(DataModel dataModel) {
        DataClass top = verifyDataClass(dataModel, 'top', 3, 'tops description', 1, 1, ['extra info': 'some extra info'])
        verifyDataElement top, 'info', 'info description', 'string', 1, 1, ['different info': 'info']
        verifyDataElement top, 'another', null, 'int', 0, 1, ['extra info': 'some extra info', 'different info': 'info']
        verifyDataElement top, 'more info', 'why not', 'string', 0

        DataClass child = verifyDataClass(top, 'child', 4, 'child description', 0)
        verifyDataElement child, 'info', 'childs info', 'string', 1, 1, ['extra info': 'some extra info']
        verifyDataElement child, 'does it work', 'I don\'t know', 'yesno', 0, -1
        verifyDataElement child, 'lazy', 'where the merging only happens in the id col', 'possibly', 1, 1, ['different info': 'info']
        verifyDataElement child, 'wrong order', 'should be fine', 'table'

        DataClass brother = verifyDataClass(top, 'brother', 2)
        verifyDataElement brother, 'info', 'brothers info', 'string', 0
        verifyDataElement brother, 'sibling', 'reference to the other child', 'child', 1, -1, ['extra info': 'some extra info']
    }

    protected void verifySimpleDataModelWithComplexMetadataContent(DataModel dataModel) {
        DataClass top = verifyDataClass(dataModel, 'top', 3, 'tops description', 1, 1, [
            'extra info'          : 'some extra info',
            'ox.softeng.test|blob': 'hello'])
        verifyDataElement top, 'info', 'info description', 'string', 1, 1, [
            'different info'    : 'info',
            'ox.softeng.xsd|min': '1',
            'ox.softeng.xsd|max': '20']
        verifyDataElement top, 'another', null, 'int', 0, 1, [
            'extra info'    : 'some extra info',
            'different info': 'info']
        verifyDataElement top, 'more info', 'why not', 'string', 0

        DataClass child = verifyDataClass(top, 'child', 4, 'child description', 0)
        verifyDataElement child, 'info', 'childs info', 'string', 1, 1, [
            'extra info'          : 'some extra info',
            'ox.softeng.xsd|min'  : '0',
            'ox.softeng.xsd|max'  : '10',
            'ox.softeng.test|blob': 'wobble']
        verifyDataElement child, 'does it work', 'I don\'t know', 'yesno', 0, -1, ['ox.softeng.xsd|choice': 'a']
        verifyDataElement child, 'lazy', 'where the merging only happens in the id col', 'possibly', 1, 1, [
            'different info'       : 'info',
            'ox.softeng.xsd|choice': 'a']
        verifyDataElement child, 'wrong order', 'should be fine', 'table'

        DataClass brother = verifyDataClass(top, 'brother', 2)
        verifyDataElement brother, 'info', 'brothers info', 'string', 0, 1, [
            'ox.softeng.xsd|min': '1',
            'ox.softeng.xsd|max': '255']
        verifyDataElement brother, 'sibling', 'reference to the other child', 'child', 1, -1, ['extra info': 'some extra info']
    }

    protected void verifyDataFlowDataModelContent(DataModel dataModel) {
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
    }

    private void verifyMetadata(CatalogueItem catalogueItem, String fullKeyName, String value) {
        String[] keyNameParts = fullKeyName.split(/\|/)
        String namespace = keyNameParts.size() > 1 ? keyNameParts[0] : importerInstance.namespace
        String key = keyNameParts[-1]

        Metadata metadata = catalogueItem.id
            ? Metadata.findByCatalogueItemIdAndNamespaceAndKey(catalogueItem.id, namespace, key)
            : Metadata.findByNamespaceAndKeyAndValue(namespace, key, value)

        assertNotNull "${catalogueItem.label} Metadata ${fullKeyName} must exist", metadata
        assertEquals "${catalogueItem.label} Metadata ${fullKeyName} value", value, metadata.value
    }

    private DataClass verifyDataClass(CatalogueItem catalogueItem, String label, int elementCount, String description = null,
                                      int min = 1, int max = 1, Map<String, String> metadata = [:]) {
        DataClass dataClass = catalogueItem instanceof DataModel
            ? catalogueItem.dataClasses.find { it.label == label && !it.parentDataClass }
            : (catalogueItem as DataClass).dataClasses.find { it.label == label }

        assertNotNull "DataClass ${label} must exist", dataClass
        assertEquals "DataClass ${label} Description", description, dataClass.description
        assertEquals "DataClass ${label} Minimum Multiplicity", min, dataClass.minMultiplicity
        assertEquals "DataClass ${label} Maximum Multiplicity", max, dataClass.maxMultiplicity
        assertEquals "DataClass ${label} number of DataElements", elementCount, dataClass.dataElements.size()

        if (dataClass.id) assertEquals "DataClass ${label} Metadata count", metadata.size(), Metadata.countByCatalogueItemId(dataClass.id)
        metadata.each { String key, String value -> verifyMetadata(dataClass, key, value) }

        dataClass
    }

    private void verifyDataElement(DataClass parent, String label, String description, String dataType, int min = 1, int max = 1,
                                   Map<String, String> metadata = [:]) {
        DataElement dataElement = parent.dataElements.find { it.label == label }

        assertNotNull "DataElement ${label} must exist", dataElement
        assertEquals "DataElement ${label} Description", description, dataElement.description
        assertEquals "DataElement ${label} Minimum Multiplicity", min, dataElement.minMultiplicity
        assertEquals "DataElement ${label} Maximum Multiplicity", max, dataElement.maxMultiplicity
        assertEquals "DataElement ${label} DataType", dataType, dataElement.dataType.label

        if (dataElement.id) assertEquals "DataElement ${label} Metadata count", metadata.size(), Metadata.countByCatalogueItemId(dataElement.id)
        metadata.each { String key, String value -> verifyMetadata(dataElement, key, value) }
    }

    private DataType verifyDataType(DataModel dataModel, String label, String description) {
        DataType dataType = dataModel.dataTypes.find { it.label == label }
        assertNotNull "DataType ${label} must exist", dataType
        assertEquals "DataType ${label} Description", dataType.description, description
        dataType
    }

    private void verifyReferenceType(DataModel dataModel, String label, String description, String referenceClassPath) {
        ReferenceType referenceType = verifyDataType(dataModel, label, description) as ReferenceType
        assertEquals "DataType ${referenceType.label} Reference to DataClass Path",
                     referenceClassPath, dataModelService.dataClassService.buildPath(referenceType.referenceClass)
    }

    private void verifyEnumerationType(DataModel dataModel, String label, String description, Map<String, String> enumerationValues) {
        EnumerationType enumerationType = verifyDataType(dataModel, label, description) as EnumerationType
        assertEquals "DataType ${enumerationType.label} number of Enumerations", enumerationValues.size(), enumerationType.enumerationValues.size()
        enumerationValues.each { String key, String value ->
            assertEquals "DataType ${enumerationType.label} Enumeration ${key}", value, enumerationType.findEnumerationValueByKey(key).value
        }
    }
}
