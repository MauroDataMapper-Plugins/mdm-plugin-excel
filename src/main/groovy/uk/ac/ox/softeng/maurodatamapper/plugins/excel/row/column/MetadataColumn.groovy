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
package uk.ac.ox.softeng.maurodatamapper.plugins.excel.row.column

/**
 * @since 13/03/2018
 */
class MetadataColumn {

    String namespace
    String key
    String value

    boolean equals(o) {
        if (this.is(o)) return true
        if (getClass() != o.class) return false

        MetadataColumn that = (MetadataColumn) o

        if (key != that.key) return false
        if (namespace != that.namespace) return false

        return true
    }

    int hashCode() {
        int result
        result = (namespace != null ? namespace.hashCode() : 0)
        result = 31 * result + (key != null ? key.hashCode() : 0)
        return result
    }

    String toString() {
        "${namespace}/${key}"
    }
}
