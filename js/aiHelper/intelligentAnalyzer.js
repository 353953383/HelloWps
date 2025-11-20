/**
 * 智能分析器 - 分析表格数据并提供智能建议
 * 负责从表格数据中提取有用信息，帮助AI生成更准确的公式
 */

class IntelligentAnalyzer {
    constructor() {
        this.dataCache = new Map();
        this.analysisHistory = new Map();
    }

    /**
     * 分析当前状态（为AI接口兼容性提供）
     * @returns {Object} 分析结果
     */
    async analyzeCurrentState() {
        try {
            global.debugLog(`🔍 [智能分析器] analyzeCurrentState 开始分析当前状态...`);
            
            // 创建默认的请求数据
            const defaultRequestData = {
                currentSheet: '当前工作表',
                currentCell: 'A1',
                description: '自动分析当前状态',
                sheetData: {
                    headers: [],
                    data: [],
                    totalRows: 0
                }
            };
            
            // 调用现有的分析工作表数据方法
            const result = this.analyzeWorksheetData(defaultRequestData);
            
            global.debugLog(`✅ [智能分析器] analyzeCurrentState 分析完成:`, result);
            
            return result;
        } catch (error) {
            global.debugLog(`❌ [智能分析器] analyzeCurrentState 分析失败: ${error.message}`, error.stack);
            
            // 返回默认分析结果
            return this.getDefaultAnalysis();
        }
    }

    /**
     * 分析当前工作表数据
     * @param {Object} requestData - 请求数据
     * @returns {Object} 分析结果
     */
    analyzeWorksheetData(requestData) {
        try {
            // 记录输入参数
            global.debugLog(`[智能分析器] 开始分析工作表数据:`, {
                sheetName: requestData.currentSheet || 'unknown',
                targetCell: requestData.currentCell || 'unknown',
                hasDescription: !!requestData.description
            });

            // 检查必需数据
            if (!requestData.sheetData || !requestData.sheetData.headers) {
                global.debugLog(`[智能分析器] 缺少表头数据`);
                return this.getDefaultAnalysis();
            }

            // 提取工作表信息
            const sheetInfo = this.extractSheetInfo(requestData);
            
            // 分析数据模式
            const dataPatterns = this.analyzeDataPatterns(requestData);
            
            // 生成智能建议
            const suggestions = this.generateSuggestions(requestData, sheetInfo, dataPatterns);

            const analysisResult = {
                sheetInfo,
                dataPatterns,
                suggestions,
                confidence: this.calculateConfidence(requestData, dataPatterns),
                metadata: {
                    analysisTime: Date.now(),
                    dataSource: 'current_worksheet',
                    version: '1.0.0'
                }
            };

            global.debugLog(`[智能分析器] 分析完成:`, analysisResult);

            return analysisResult;

        } catch (error) {
            global.debugLog(`[智能分析器] 分析失败: ${error.message}`, error.stack);
            return this.getDefaultAnalysis();
        }
    }

    /**
     * 提取工作表信息
     * @param {Object} requestData - 请求数据
     * @returns {Object} 工作表信息
     */
    extractSheetInfo(requestData) {
        const sheetName = requestData.currentSheet || '工作表1';
        const headers = requestData.sheetData?.headers || [];
        const currentCell = requestData.currentCell || 'A1';
        const currentColumn = this.extractColumnFromCell(currentCell);

        return {
            name: sheetName,
            headers: headers,
            headerCount: headers.length,
            currentCell,
            currentColumn,
            totalRows: requestData.sheetData?.totalRows || 0,
            lastColumn: headers.length > 0 ? this.columnIndexToLetter(headers.length) : 'A'
        };
    }

    /**
     * 分析数据模式
     * @param {Object} requestData - 请求数据
     * @returns {Object} 数据模式
     */
    analyzeDataPatterns(requestData) {
        const patterns = {
            numericColumns: [],
            dateColumns: [],
            textColumns: [],
            emptyColumns: [],
            dataTypes: {},
            relationships: {}
        };

        try {
            if (requestData.sheetData && requestData.sheetData.data) {
                const data = requestData.sheetData.data;
                
                // 分析数据类型
                patterns.dataTypes = this.detectDataTypes(data, requestData.sheetData.headers);
                
                // 识别空列
                patterns.emptyColumns = this.findEmptyColumns(data, requestData.sheetData.headers);
                
                // 分析数值列
                patterns.numericColumns = this.findNumericColumns(data, requestData.sheetData.headers);
                
                // 分析日期列
                patterns.dateColumns = this.findDateColumns(data, requestData.sheetData.headers);
                
                // 分析文本列
                patterns.textColumns = this.findTextColumns(data, requestData.sheetData.headers);
                
            } else {
                // 如果没有数据，仅基于表头推断
                patterns.dataTypes = this.inferTypesFromHeaders(requestData.sheetData?.headers || []);
            }

        } catch (error) {
            global.debugLog(`[智能分析器] 数据模式分析失败: ${error.message}`);
        }

        return patterns;
    }

    /**
     * 生成智能建议
     * @param {Object} requestData - 请求数据
     * @param {Object} sheetInfo - 工作表信息
     * @param {Object} dataPatterns - 数据模式
     * @returns {Array} 建议列表
     */
    generateSuggestions(requestData, sheetInfo, dataPatterns) {
        const suggestions = [];
        
        try {
            // 基于description生成建议
            if (requestData.description) {
                const descSuggestions = this.analyzeDescription(requestData.description, sheetInfo, dataPatterns);
                suggestions.push(...descSuggestions);
            }

            // 基于数据结构生成建议
            const structureSuggestions = this.analyzeStructure(sheetInfo, dataPatterns);
            suggestions.push(...structureSuggestions);

            // 基于数据模式生成建议
            const patternSuggestions = this.analyzePatterns(dataPatterns);
            suggestions.push(...patternSuggestions);

        } catch (error) {
            global.debugLog(`[智能分析器] 建议生成失败: ${error.message}`);
        }

        return suggestions.length > 0 ? suggestions : this.getDefaultSuggestions();
    }

    /**
     * 基于description分析
     * @param {string} description - 描述
     * @param {Object} sheetInfo - 工作表信息
     * @param {Object} dataPatterns - 数据模式
     * @returns {Array} 建议
     */
    analyzeDescription(description, sheetInfo, dataPatterns) {
        const suggestions = [];
        const desc = description.toLowerCase();

        // 数值计算相关
        if (desc.includes('求和') || desc.includes('sum')) {
            suggestions.push({
                type: 'sum',
                formula: 'SUM',
                description: '数值求和计算',
                confidence: 0.9,
                parameters: this.findNumericColumnsForFormula(sheetInfo, dataPatterns)
            });
        }

        if (desc.includes('平均') || desc.includes('average')) {
            suggestions.push({
                type: 'average',
                formula: 'AVERAGE',
                description: '数值平均值计算',
                confidence: 0.9,
                parameters: this.findNumericColumnsForFormula(sheetInfo, dataPatterns)
            });
        }

        if (desc.includes('计数') || desc.includes('count')) {
            suggestions.push({
                type: 'count',
                formula: 'COUNT',
                description: '数值计数',
                confidence: 0.8,
                parameters: this.findNumericColumnsForFormula(sheetInfo, dataPatterns)
            });
        }

        // 日期相关
        if (desc.includes('日期') || desc.includes('时间') || desc.includes('date') || desc.includes('time')) {
            suggestions.push({
                type: 'date_calculation',
                formula: 'DATE',
                description: '日期计算',
                confidence: 0.7,
                parameters: this.findDateColumnsForFormula(sheetInfo, dataPatterns)
            });
        }

        // 查找相关列
        if (desc.includes('查找') || desc.includes('搜索') || desc.includes('vlookup') || desc.includes('lookup')) {
            suggestions.push({
                type: 'lookup',
                formula: 'VLOOKUP',
                description: '垂直查找',
                confidence: 0.8,
                parameters: this.findLookupColumns(sheetInfo, dataPatterns)
            });
        }

        return suggestions;
    }

    /**
     * 基于结构分析
     * @param {Object} sheetInfo - 工作表信息
     * @param {Object} dataPatterns - 数据模式
     * @returns {Array} 建议
     */
    analyzeStructure(sheetInfo, dataPatterns) {
        const suggestions = [];

        // 如果有完整的表头结构
        if (sheetInfo.headers && sheetInfo.headers.length > 0) {
            suggestions.push({
                type: 'header_based',
                formula: 'INDEX-MATCH',
                description: '基于表头的索引匹配',
                confidence: 0.7,
                parameters: sheetInfo.headers
            });
        }

        return suggestions;
    }

    /**
     * 基于模式分析
     * @param {Object} dataPatterns - 数据模式
     * @returns {Array} 建议
     */
    analyzePatterns(dataPatterns) {
        const suggestions = [];

        // 基于数据类型模式
        if (dataPatterns.numericColumns.length > 0) {
            suggestions.push({
                type: 'numeric_aggregation',
                formula: 'SUM',
                description: '数值聚合计算',
                confidence: 0.6,
                parameters: dataPatterns.numericColumns
            });
        }

        if (dataPatterns.dateColumns.length > 0) {
            suggestions.push({
                type: 'date_analysis',
                formula: 'DATEDIF',
                description: '日期差值计算',
                confidence: 0.6,
                parameters: dataPatterns.dateColumns
            });
        }

        return suggestions;
    }

    /**
     * 检测数据类型
     * @param {Array} data - 数据
     * @param {Array} headers - 表头
     * @returns {Object} 数据类型
     */
    detectDataTypes(data, headers) {
        const types = {};

        headers.forEach((header, index) => {
            const columnData = data.map(row => row[index]).filter(val => val !== null && val !== undefined && val !== '');
            
            if (columnData.length === 0) {
                types[header] = 'empty';
            } else {
                types[header] = this.detectColumnType(columnData);
            }
        });

        return types;
    }

    /**
     * 检测列类型
     * @param {Array} columnData - 列数据
     * @returns {string} 数据类型
     */
    detectColumnType(columnData) {
        let numericCount = 0;
        let dateCount = 0;
        let textCount = 0;

        columnData.forEach(value => {
            const strValue = String(value).trim();
            
            if (this.isNumeric(strValue)) {
                numericCount++;
            } else if (this.isDate(strValue)) {
                dateCount++;
            } else {
                textCount++;
            }
        });

        const total = columnData.length;
        const numericRatio = numericCount / total;
        const dateRatio = dateCount / total;

        if (numericRatio > 0.8) return 'numeric';
        if (dateRatio > 0.8) return 'date';
        if (textRatio > 0.8) return 'text';
        return 'mixed';
    }

    /**
     * 查找空列
     * @param {Array} data - 数据
     * @param {Array} headers - 表头
     * @returns {Array} 空列索引
     */
    findEmptyColumns(data, headers) {
        const emptyColumns = [];
        
        headers.forEach((header, index) => {
            const columnData = data.map(row => row[index]).filter(val => val !== null && val !== undefined);
            if (columnData.length === 0) {
                emptyColumns.push(index);
            }
        });

        return emptyColumns;
    }

    /**
     * 查找数值列
     * @param {Array} data - 数据
     * @param {Array} headers - 表头
     * @returns {Array} 数值列索引
     */
    findNumericColumns(data, headers) {
        const numericColumns = [];
        
        headers.forEach((header, index) => {
            const columnData = data.map(row => row[index]).filter(val => val !== null && val !== undefined);
            if (columnData.length > 0 && columnData.every(val => this.isNumeric(String(val)))) {
                numericColumns.push(index);
            }
        });

        return numericColumns;
    }

    /**
     * 查找日期列
     * @param {Array} data - 数据
     * @param {Array} headers - 表头
     * @returns {Array} 日期列索引
     */
    findDateColumns(data, headers) {
        const dateColumns = [];
        
        headers.forEach((header, index) => {
            const columnData = data.map(row => row[index]).filter(val => val !== null && val !== undefined);
            if (columnData.length > 0 && columnData.every(val => this.isDate(String(val)))) {
                dateColumns.push(index);
            }
        });

        return dateColumns;
    }

    /**
     * 查找文本列
     * @param {Array} data - 数据
     * @param {Array} headers - 表头
     * @returns {Array} 文本列索引
     */
    findTextColumns(data, headers) {
        const textColumns = [];
        
        headers.forEach((header, index) => {
            const columnData = data.map(row => row[index]).filter(val => val !== null && val !== undefined);
            if (columnData.length > 0 && columnData.every(val => !this.isNumeric(String(val)) && !this.isDate(String(val)))) {
                textColumns.push(index);
            }
        });

        return textColumns;
    }

    /**
     * 从表头推断类型
     * @param {Array} headers - 表头
     * @returns {Object} 数据类型
     */
    inferTypesFromHeaders(headers) {
        const types = {};
        
        headers.forEach(header => {
            const lowerHeader = header.toLowerCase();
            
            if (lowerHeader.includes('金额') || lowerHeader.includes('价格') || lowerHeader.includes('数量') || lowerHeader.includes('数') || 
                lowerHeader.includes('amount') || lowerHeader.includes('price') || lowerHeader.includes('quantity') || lowerHeader.includes('number')) {
                types[header] = 'numeric';
            } else if (lowerHeader.includes('日期') || lowerHeader.includes('时间') || lowerHeader.includes('date') || lowerHeader.includes('time')) {
                types[header] = 'date';
            } else if (lowerHeader.includes('名称') || lowerHeader.includes('描述') || lowerHeader.includes('备注') || 
                       lowerHeader.includes('name') || lowerHeader.includes('description') || lowerHeader.includes('remark')) {
                types[header] = 'text';
            } else {
                types[header] = 'unknown';
            }
        });

        return types;
    }

    /**
     * 计算置信度
     * @param {Object} requestData - 请求数据
     * @param {Object} dataPatterns - 数据模式
     * @returns {number} 置信度
     */
    calculateConfidence(requestData, dataPatterns) {
        let confidence = 0.5; // 基础置信度

        // description 存在性
        if (requestData.description && requestData.description.trim() !== '') {
            confidence += 0.2;
        }

        // 数据完整性
        if (requestData.sheetData && requestData.sheetData.headers && requestData.sheetData.headers.length > 0) {
            confidence += 0.2;
        }

        // 数据模式匹配
        if (Object.keys(dataPatterns.dataTypes).length > 0) {
            confidence += 0.1;
        }

        return Math.min(confidence, 1.0);
    }

    /**
     * 获取默认分析结果
     * @returns {Object} 默认分析
     */
    getDefaultAnalysis() {
        return {
            sheetInfo: {
                name: '默认工作表',
                headers: [],
                headerCount: 0,
                currentCell: 'A1',
                currentColumn: 'A',
                totalRows: 0,
                lastColumn: 'A'
            },
            dataPatterns: {
                numericColumns: [],
                dateColumns: [],
                textColumns: [],
                emptyColumns: [],
                dataTypes: {},
                relationships: {}
            },
            suggestions: this.getDefaultSuggestions(),
            confidence: 0.3,
            metadata: {
                analysisTime: Date.now(),
                dataSource: 'default',
                version: '1.0.0'
            }
        };
    }

    /**
     * 获取默认建议
     * @returns {Array} 默认建议
     */
    getDefaultSuggestions() {
        return [
            {
                type: 'basic_calculation',
                formula: 'SUM',
                description: '基础求和计算',
                confidence: 0.5,
                parameters: ['A:A']
            },
            {
                type: 'basic_lookup',
                formula: 'VLOOKUP',
                description: '基础查找功能',
                confidence: 0.4,
                parameters: ['lookup_value', 'table_array', 'col_index']
            }
        ];
    }

    /**
     * 查找公式的数值列
     * @param {Object} sheetInfo - 工作表信息
     * @param {Object} dataPatterns - 数据模式
     * @returns {Array} 数值列
     */
    findNumericColumnsForFormula(sheetInfo, dataPatterns) {
        if (!dataPatterns.numericColumns || dataPatterns.numericColumns.length === 0) {
            return ['A:A']; // 默认返回第一列
        }
        
        return dataPatterns.numericColumns.map(index => 
            this.columnIndexToLetter(index + 1) + ':' + this.columnIndexToLetter(index + 1)
        );
    }

    /**
     * 查找公式的日期列
     * @param {Object} sheetInfo - 工作表信息
     * @param {Object} dataPatterns - 数据模式
     * @returns {Array} 日期列
     */
    findDateColumnsForFormula(sheetInfo, dataPatterns) {
        if (!dataPatterns.dateColumns || dataPatterns.dateColumns.length === 0) {
            return ['B:B']; // 默认返回第二列
        }
        
        return dataPatterns.dateColumns.map(index => 
            this.columnIndexToLetter(index + 1) + ':' + this.columnIndexToLetter(index + 1)
        );
    }

    /**
     * 查找公式的查找列
     * @param {Object} sheetInfo - 工作表信息
     * @param {Object} dataPatterns - 数据模式
     * @returns {Array} 查找列
     */
    findLookupColumns(sheetInfo, dataPatterns) {
        const result = [];
        
        if (dataPatterns.textColumns && dataPatterns.textColumns.length > 0) {
            result.push('查找值列: ' + this.columnIndexToLetter(dataPatterns.textColumns[0] + 1));
        }
        
        if (dataPatterns.dataTypes && Object.keys(dataPatterns.dataTypes).length > 0) {
            result.push('数据表范围');
        }
        
        return result.length > 0 ? result : ['查找值', '数据表', '列索引'];
    }

    // 工具方法
    isNumeric(value) {
        return !isNaN(parseFloat(value)) && isFinite(value);
    }

    isDate(value) {
        const dateRegex = /^\d{1,2}[-/]\d{1,2}[-/]\d{2,4}$|^\d{4}[-/]\d{1,2}[-/]\d{1,2}$/;
        return dateRegex.test(value) && !isNaN(new Date(value).getTime());
    }

    extractColumnFromCell(cellAddress) {
        const match = cellAddress.match(/^([A-Z]+)/);
        return match ? match[1] : 'A';
    }

    columnIndexToLetter(index) {
        let temp = '';
        let columnIndex = index;
        
        while (columnIndex > 0) {
            const modulo = (columnIndex - 1) % 26;
            temp = String.fromCharCode(65 + modulo) + temp;
            columnIndex = Math.floor((columnIndex - modulo) / 26);
        }
        
        return temp;
    }

    /**
     * 缓存分析结果
     * @param {string} key - 缓存键
     * @param {Object} result - 分析结果
     */
    cacheAnalysis(key, result) {
        this.dataCache.set(key, {
            result,
            timestamp: Date.now()
        });
    }

    /**
     * 获取缓存的分析结果
     * @param {string} key - 缓存键
     * @param {number} maxAge - 最大年龄（毫秒）
     * @returns {Object|null} 分析结果
     */
    getCachedAnalysis(key, maxAge = 300000) { // 5分钟默认缓存
        const cached = this.dataCache.get(key);
        if (cached && Date.now() - cached.timestamp < maxAge) {
            return cached.result;
        }
        return null;
    }

    /**
     * 清理过期缓存
     */
    cleanupCache() {
        const now = Date.now();
        const maxAge = 300000; // 5分钟
        
        for (const [key, cached] of this.dataCache.entries()) {
            if (now - cached.timestamp > maxAge) {
                this.dataCache.delete(key);
            }
        }
    }

    /**
     * 获取分析统计信息
     * @returns {Object} 统计信息
     */
    getStats() {
        return {
            cacheSize: this.dataCache.size,
            analysisHistorySize: this.analysisHistory.size,
            uptime: Date.now() - (global.startTime || Date.now())
        };
    }
}

// 如果在浏览器环境中使用
if (typeof window !== 'undefined') {
    window.IntelligentAnalyzer = IntelligentAnalyzer;
}

// 如果在Node.js环境中使用
if (typeof module !== 'undefined' && module.exports) {
    module.exports = IntelligentAnalyzer;
}