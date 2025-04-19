/*
 * Copyright 2015-2016 the original author or authors.
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *      http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */
package org.hellojavaer.poi.excel.utils.write;

import org.apache.poi.common.usermodel.HyperlinkType;
import org.hellojavaer.poi.excel.utils.ExcelUtils;

import java.io.Serializable;
import java.util.LinkedHashMap;
import java.util.Map;
import java.util.concurrent.ConcurrentHashMap;

/**
 * Config the mapping between excel column(by index) to Object field(by name).
 * 
 * @author <a href="mailto:hellojavaer@gmail.com">zoukaiming</a>
 */
public class ExcelWriteFieldMapping implements Serializable {

    private static final long                                          serialVersionUID = 1L;

    private Map<String, Map<Integer, ExcelWriteFieldMappingAttribute>> fieldMapping     = new LinkedHashMap<String, Map<Integer, ExcelWriteFieldMappingAttribute>>();

    public ExcelWriteFieldMappingAttribute put(String colIndex, String fieldName) {
        if (colIndex == null) {
            throw new IllegalArgumentException("colIndex must not be null");
        }
        if (fieldName == null) {
            throw new IllegalArgumentException("fieldName must not be null");
        }
        
        Map<Integer, ExcelWriteFieldMappingAttribute> map = fieldMapping.get(fieldName);
        if (map == null) {
            synchronized (fieldMapping) {
                if (fieldMapping.get(colIndex) == null) {
                    map = new ConcurrentHashMap<Integer, ExcelWriteFieldMappingAttribute>();
                    fieldMapping.put(fieldName, map);
                }
            }
        }
        ExcelWriteFieldMappingAttribute attribute = new ExcelWriteFieldMappingAttribute();
        map.put(ExcelUtils.convertColCharIndexToIntIndex(colIndex), attribute);
        return attribute;
    }

    public Map<String, Map<Integer, ExcelWriteFieldMappingAttribute>> export() {
        return fieldMapping;
    }

    public class ExcelWriteFieldMappingAttribute implements Serializable {

        private static final long          serialVersionUID = 1L;
        @SuppressWarnings("rawtypes")
        private ExcelWriteCellProcessor    cellProcessor;
        private ExcelWriteCellValueMapping valueMapping;
        private String                     head;
        private String                     linkField;
        private HyperlinkType linkType;

        @SuppressWarnings("rawtypes")
        public ExcelWriteFieldMappingAttribute setCellProcessor(ExcelWriteCellProcessor cellProcessor) {
            this.cellProcessor = cellProcessor;
            return this;
        }

        public ExcelWriteFieldMappingAttribute setValueMapping(ExcelWriteCellValueMapping valueMapping) {
            this.valueMapping = valueMapping;
            return this;
        }

        public ExcelWriteFieldMappingAttribute setHead(String head) {
            this.head = head;
            return this;
        }

        public ExcelWriteFieldMappingAttribute setLinkField(String linkField) {
            this.linkField = linkField;
            this.linkType = HyperlinkType.URL;
            return this;
        }

        public ExcelWriteFieldMappingAttribute setLink(String linkFieldName, HyperlinkType linkType) {
            this.linkField = linkFieldName;
            this.linkType = linkType;
            return this;
        }

        @SuppressWarnings("rawtypes")
        public ExcelWriteCellProcessor getCellProcessor() {
            return cellProcessor;
        }

        public ExcelWriteCellValueMapping getValueMapping() {
            return valueMapping;
        }

        public String getHead() {
            return head;
        }

        public String getLinkField() {
            return linkField;
        }

        public HyperlinkType getLinkType() {
            return linkType;
        }

    }
}
