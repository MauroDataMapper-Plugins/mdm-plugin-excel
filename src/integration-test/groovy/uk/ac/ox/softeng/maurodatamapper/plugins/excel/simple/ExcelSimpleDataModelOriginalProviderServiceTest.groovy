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
package uk.ac.ox.softeng.maurodatamapper.plugins.excel.simple

import org.junit.Test

import static org.junit.Assert.assertEquals

class ExcelSimpleDataModelOriginalProviderServiceTest {
    @Test
    void testSheetKey() {
        Map<String, String> results = [
            'My Model'                  : 'MM',
            'GEL: CancerSchema-v1.0.1'  : 'GC101',
            '[GEL: CancerSchema-v1.0.1]': 'GC101'
        ].each { assertEquals ExcelSimpleDataModelExporterProviderService.createSheetKey(it.key), it.value }
    }
}
