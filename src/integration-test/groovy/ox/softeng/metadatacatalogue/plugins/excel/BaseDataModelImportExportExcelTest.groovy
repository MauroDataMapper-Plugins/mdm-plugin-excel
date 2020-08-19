package ox.softeng.metadatacatalogue.plugins.excel

import ox.softeng.metadatacatalogue.core.catalogue.linkable.component.DataClass
import ox.softeng.metadatacatalogue.core.catalogue.linkable.component.DataElement
import ox.softeng.metadatacatalogue.core.catalogue.linkable.component.datatype.DataType
import ox.softeng.metadatacatalogue.core.catalogue.linkable.component.datatype.EnumerationType
import ox.softeng.metadatacatalogue.core.catalogue.linkable.datamodel.DataModel
import ox.softeng.metadatacatalogue.core.catalogue.linkable.datamodel.DataModelService
import ox.softeng.metadatacatalogue.core.spi.importer.parameter.FileParameter
import ox.softeng.metadatacatalogue.core.type.catalogue.DataModelType
import ox.softeng.metadatacatalogue.plugins.excel.parameter.ExcelFileImporterParameters

import java.nio.file.Files
import java.nio.file.Path

import static org.junit.Assert.assertEquals
import static org.junit.Assert.assertNotNull

/**
 * @since 01/03/2018
 */
abstract class BaseDataModelImportExportExcelTest
    extends BaseImportExportExcelTest<DataModel, ExcelFileImporterParameters, ExcelDataModelImporterService> {

    @Override
    DataModel saveDomain(DataModel domain) {
        getBean(DataModelService).saveWithBatching(domain)
    }

    protected void verifySimpleDataModel(DataModel dataModel, int mdCount = 3) {
        assertNotNull('Simple model must exist', dataModel)

        assertEquals('DataModel name', 'test', dataModel.label)
        assertEquals('DataModel description', 'this is a simple test with only 1 DataModel', dataModel.description)
        assertEquals('DataModel author', 'tester', dataModel.author)
        assertEquals('DataModel organisation', 'Oxford', dataModel.organisation)
        assertEquals('DataModel type', DataModelType.DATA_STANDARD, dataModel.type)

        assertEquals('DataModel Metadata count', mdCount, dataModel.metadata.size())

        checkMetadata dataModel, 'reviewed', 'yes'
        checkMetadata dataModel, 'approved', 'yes'
        checkMetadata dataModel, 'distributed', 'no'

        assertEquals('Number of DataTypes', 6, dataModel.dataTypes.size())
        assertEquals('Number of DataClasses', 3, dataModel.dataClasses.size())

        checkDataType dataModel, 'string', 'just a plain old string'
        checkDataType dataModel, 'int', 'a number'
        checkDataType dataModel, 'table', 'a random thing'
        checkEnumeration dataModel, 'yesno', 'an eumeration', [y: 'yes', n: 'no']
        checkEnumeration dataModel, 'possibly', null, ['0': 'lazy', '1': 'not lazy', '2': 'very lazy']
        checkDataType dataModel, 'child', null

    }

    protected void verifySimpleDataModelContent(DataModel dataModel) {

        DataClass top = checkDataClass(dataModel, 'top', 3, 'tops description', 1, 1, ['extra info': 'some extra info'])
        checkDataElement top, 'info', 'info description', 'string', 1, 1, ['different info': 'info']
        checkDataElement top, 'another', null, 'int', 0, 1, ['extra info': 'some extra info', 'different info': 'info']
        checkDataElement top, 'more info', 'why not', 'string', 0

        DataClass child = checkDataClass(top, 'child', 4, 'child description', 0)
        checkDataElement child, 'info', 'childs info', 'string', 1, 1, ['extra info': 'some extra info']
        checkDataElement child, 'does it work', 'I don\'t know', 'yesno', 0, -1
        checkDataElement child, 'lazy', 'where the merging only happens in the id col', 'possibly', 1, 1, ['different info': 'info']
        checkDataElement child, 'wrong order', 'should be fine', 'table'

        DataClass brother = checkDataClass(top, 'brother', 2)
        checkDataElement brother, 'info', 'brothers info', 'string', 0
        checkDataElement brother, 'sibling', 'reference to the other child', 'child', 1, -1, ['extra info': 'some extra info']
    }

    protected void checkDataElement(DataClass parent, String label, String description, String dataType, int min = 1, int max = 1,
                                    Map<String, String> metadata = [:]) {
        DataElement dataElement = parent.childDataElements.find {it.label == label}
        assertNotNull("${label} dataelement exists", dataElement)
        assertEquals("${label} dataelement description", description, dataElement.description)
        assertEquals("${label} dataelement datatype", dataType, dataElement.dataType.label)
        assertEquals("${label} dataelement min multiplicity", min, dataElement.minMultiplicity)
        assertEquals("${label} dataelement max multiplicity", max, dataElement.maxMultiplicity)
        assertEquals("${label} dataelement Metadata count", metadata.size(), dataElement.metadata?.size() ?: 0)
        metadata.each {k, v ->
            checkMetadata(dataElement, k, v)
        }
    }

    protected DataClass checkDataClass(DataModel dataModel, String label, int elementCount, String description = null, int min = 1, int max = 1,
                                       Map<String, String> metadata = [:]) {
        DataClass dataClass = dataModel.dataClasses.find {it.label == label && !it.parentDataClass}
        checkDataClassData(dataClass, label, elementCount, description, min, max, metadata)
    }

    protected DataClass checkDataClass(DataClass parent, String label, int elementCount, String description = null, int min = 1, int max = 1,
                                       Map<String, String> metadata = [:]) {
        DataClass dataClass = parent.childDataClasses.find {it.label == label}
        checkDataClassData(dataClass, label, elementCount, description, min, max, metadata)
    }

    protected DataClass checkDataClassData(DataClass dataClass, String label, int elementCount, String description = null, int min = 1, int max = 1,
                                           Map<String, String> metadata = [:]) {
        assertNotNull("${label} dataclass exists", dataClass)
        assertEquals("${label} dataclass has number of elements", elementCount, dataClass.childDataElements.size())
        assertEquals("${label} dataclass description", description, dataClass.description)
        assertEquals("${label} dataclass min multiplicity", min, dataClass.minMultiplicity)
        assertEquals("${label} dataclass max multiplicity", max, dataClass.maxMultiplicity)
        assertEquals("${label} dataclass Metadata count", metadata.size(), dataClass.metadata?.size() ?: 0)
        metadata.each {k, v ->
            checkMetadata(dataClass, k, v)
        }
        dataClass
    }

    protected void checkEnumeration(DataModel dataModel, String label, String description, Map<String, String> enumValues) {
        EnumerationType enumerationType = checkDataType(dataModel, label, description) as EnumerationType
        assertEquals("${enumerationType.label} number of enumerations", enumValues.size(), enumerationType.enumerationValues.size())
        enumValues.each {k, v ->
            assertEquals("${enumerationType.label} enumeration ${k}", v, enumerationType.findEnumerationValueByKey(k).value)
        }
    }

    protected DataType checkDataType(DataModel dataModel, String label, String description) {
        DataType dataType = dataModel.dataTypes.find {it.label == label}
        assertNotNull("${label} datatype exists", dataType)
        assertEquals("${label} datatype description", dataType.description, description)
        dataType
    }

    protected ExcelFileImporterParameters createImportParameters(Path path) throws IOException {
        ExcelFileImporterParameters params = new ExcelFileImporterParameters(finalised: false)
        FileParameter file = new FileParameter(path.toString(), 'application/vnd.ms-excel', Files.readAllBytes(path))

        params.setImportFile(file)
        params
    }
}
