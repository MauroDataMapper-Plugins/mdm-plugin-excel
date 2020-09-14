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
import uk.ac.ox.softeng.maurodatamapper.core.provider.importer.ImporterProviderService
import uk.ac.ox.softeng.maurodatamapper.core.provider.importer.parameter.ImporterProviderServiceParameters
import uk.ac.ox.softeng.maurodatamapper.plugins.testing.utils.BaseImportPluginTest

import org.grails.datastore.gorm.GormEntity

import java.nio.file.Files
import java.nio.file.Path
import java.nio.file.Paths

import static org.junit.Assert.assertEquals
import static org.junit.Assert.assertNotNull
import static org.junit.Assert.fail

/**
 * @since 09/03/2018
 */
abstract class BaseImporterExporterProviderServiceTest<D extends GormEntity, P extends ImporterProviderServiceParameters, T extends ImporterProviderService<D, P>>
    extends BaseImportPluginTest<D, P, T> {

    protected void checkMetadata(CatalogueItem catalogueItem, String key, String value) {

        String[] parts = key.split(/\|/)
        String ns = importerInstance.namespace
        String k = parts[0]
        if (parts.size() != 1) {
            ns = parts[0]
            k = parts[1]
        }

        Metadata md = catalogueItem.id ? Metadata.findByCatalogueItemIdAndNamespaceAndKey(catalogueItem.id, ns, k)
                                       : Metadata.findByNamespaceAndKey(ns, k)
        assertNotNull("${catalogueItem.label} Metadata ${key} exists", md)
        assertEquals("${catalogueItem.label} Metadata ${key} value", value, md.value)
    }

    protected P createImportParameters(String filename) throws IOException {
        Path p = Paths.get('src/integration-test/resources/' + filename)
        if (!Files.exists(p)) {
            fail("File ${filename} cannot be found")
        }
        createImportParameters(p)
    }

    abstract protected P createImportParameters(Path path) throws IOException
}
