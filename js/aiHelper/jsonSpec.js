/**
 * JSON格式规范配置
 * 定义AI接口的输入输出JSON格式标准
 */

const AIJsonSpec = {
    // 请求数据格式规范
    request: {
        description: {
            type: 'string',
            required: true,
            description: '用户对公式需求的文字描述',
            example: '计算销售总额和平均销售额',
            maxLength: 500
        },
        referenceType: {
            type: 'string',
            required: true,
            enum: ['current', 'worksheet', 'workbook'],
            description: '引用数据类型',
            description_map: {
                'current': '当前工作表',
                'worksheet': '跨工作表',
                'workbook': '跨工作簿'
            },
            example: 'workbook'
        },
        currentCell: {
            type: 'object',
            required: true,
            description: '当前操作的单元格信息',
            properties: {
                cellAddress: {
                    type: 'string',
                    description: '单元格地址',
                    example: 'C5'
                },
                row: {
                    type: 'number',
                    description: '行号',
                    example: 5
                },
                column: {
                    type: 'number',
                    description: '列号',
                    example: 3
                },
                worksheet: {
                    type: 'string',
                    description: '工作表名称',
                    example: '销售数据'
                }
            }
        },
        selectedWorkbooks: {
            type: 'array',
            required: false,
            description: '选中的工作簿信息',
            items: {
                type: 'object',
                properties: {
                    name: {
                        type: 'string',
                        description: '工作簿名称',
                        example: '销售报表.xlsx'
                    },
                    worksheets: {
                        type: 'array',
                        items: {
                            type: 'object',
                            properties: {
                                name: {
                                    type: 'string',
                                    description: '工作表名称',
                                    example: '汇总数据'
                                },
                                usedRange: {
                                    type: 'object',
                                    properties: {
                                        rows: {
                                            type: 'number',
                                            description: '使用行数',
                                            example: 100
                                        },
                                        columns: {
                                            type: 'number',
                                            description: '使用列数',
                                            example: 10
                                        }
                                    }
                                },
                                headers: {
                                    type: 'array',
                                    items: {
                                        type: 'object',
                                        properties: {
                                            column: {
                                                type: 'number',
                                                description: '列号',
                                                example: 1
                                            },
                                            value: {
                                                type: 'string',
                                                description: '列标题',
                                                example: '产品名称'
                                            }
                                        }
                                    },
                                    description: '表头信息'
                                }
                            }
                        }
                    }
                }
            }
        },
        selectedWorksheets: {
            type: 'array',
            required: false,
            description: '选中的工作表信息（当引用类型为跨工作表时）',
            items: {
                type: 'object',
                properties: {
                    name: {
                        type: 'string',
                        description: '工作表名称',
                        example: '数据源'
                    },
                    range: {
                        type: 'string',
                        description: '选择范围',
                        example: 'A1:Z100'
                    }
                }
            }
        },
        fillOptions: {
            type: 'object',
            required: true,
            description: '填充选项',
            properties: {
                right: {
                    type: 'boolean',
                    description: '向右填充',
                    example: false
                },
                down: {
                    type: 'boolean',
                    description: '向下填充',
                    example: true
                }
            }
        },
        headers: {
            type: 'array',
            required: false,
            description: '发现的表头信息',
            items: {
                type: 'object',
                properties: {
                    workbook: {
                        type: 'string',
                        description: '工作簿名称',
                        example: '销售报表.xlsx'
                    },
                    worksheet: {
                        type: 'string',
                        description: '工作表名称',
                        example: '销售数据'
                    },
                    headers: {
                        type: 'array',
                        items: {
                            type: 'object',
                            properties: {
                                column: {
                                    type: 'number',
                                    description: '列号',
                                    example: 1
                                },
                                value: {
                                    type: 'string',
                                    description: '列标题',
                                    example: '产品名称'
                                },
                                dataType: {
                                    type: 'string',
                                    enum: ['text', 'number', 'date', 'currency', 'percentage'],
                                    description: '数据类型',
                                    example: 'number'
                                }
                            }
                        }
                    }
                }
            }
        }
    },
    
    // 响应数据格式规范
    response: {
        formulas: {
            type: 'array',
            required: true,
            description: 'AI生成的公式建议列表',
            items: {
                type: 'object',
                properties: {
                    title: {
                        type: 'string',
                        required: true,
                        description: '公式名称或描述',
                        example: '销售总额求和',
                        maxLength: 100
                    },
                    formula: {
                        type: 'string',
                        required: true,
                        description: '完整的Excel公式',
                        example: '=SUM(B2:B100)',
                        maxLength: 500
                    },
                    explanation: {
                        type: 'string',
                        required: true,
                        description: '公式详细说明',
                        example: '对B2到B100单元格进行求和，计算所有产品的销售总额',
                        maxLength: 300
                    },
                    confidence: {
                        type: 'number',
                        required: true,
                        minimum: 1,
                        maximum: 100,
                        description: '公式准确性置信度（1-100）',
                        example: 95
                    },
                    applicable_ranges: {
                        type: 'array',
                        required: false,
                        description: '公式适用范围的说明',
                        items: {
                            type: 'string'
                        },
                        example: ['数值列', '连续数据区域']
                    },
                    required_functions: {
                        type: 'array',
                        required: false,
                        description: '公式使用的函数列表',
                        items: {
                            type: 'string'
                        },
                        example: ['SUM', 'IF', 'VLOOKUP']
                    },
                    example: {
                        type: 'string',
                        required: false,
                        description: '使用示例',
                        example: '=SUM(B2:B10) 计算B2到B10的总和'
                    }
                }
            },
            minItems: 1,
            maxItems: 5
        },
        data_analysis: {
            type: 'object',
            required: false,
            description: '数据分析结果',
            properties: {
                headers_found: {
                    type: 'array',
                    required: false,
                    description: '发现的表头列表',
                    items: {
                        type: 'string'
                    },
                    example: ['产品名称', '销售数量', '单价', '总价']
                },
                data_types: {
                    type: 'array',
                    required: false,
                    description: '数据类型分析结果',
                    items: {
                        type: 'string'
                    },
                    example: ['文本', '数字', '货币', '百分比']
                },
                recommendations: {
                    type: 'array',
                    required: false,
                    description: '使用建议',
                    items: {
                        type: 'string'
                    },
                    example: [
                        '数据连续无空值，适合使用聚合函数',
                        '建议使用绝对引用锁定关键单元格'
                    ]
                }
            }
        },
        alternative_formulas: {
            type: 'array',
            required: false,
            description: '替代方案',
            items: {
                type: 'object',
                properties: {
                    description: {
                        type: 'string',
                        required: true,
                        description: '替代方案描述',
                        example: '使用数组公式实现相同功能'
                    },
                    formula: {
                        type: 'string',
                        required: true,
                        description: '替代公式',
                        example: '=SUMPRODUCT(B2:B100, C2:C100)'
                    },
                    pros: {
                        type: 'array',
                        required: false,
                        description: '优点列表',
                        items: {
                            type: 'string'
                        },
                        example: ['性能更好', '支持复杂条件']
                    },
                    cons: {
                        type: 'array',
                        required: false,
                        description: '缺点列表',
                        items: {
                            type: 'string'
                        },
                        example: ['公式较复杂', '兼容性要求较高']
                    }
                }
            }
        },
        metadata: {
            type: 'object',
            required: false,
            description: '响应元数据',
            properties: {
                model: {
                    type: 'string',
                    description: '使用的AI模型名称',
                    example: 'abab6.5s-chat'
                },
                tokens_used: {
                    type: 'number',
                    description: '消耗的令牌数量',
                    example: 150
                },
                timestamp: {
                    type: 'string',
                    format: 'date-time',
                    description: '响应时间戳',
                    example: '2024-01-15T10:30:00Z'
                },
                request_id: {
                    type: 'string',
                    description: '请求唯一标识',
                    example: 'req_123456789'
                }
            }
        }
    },
    
    // 错误响应格式
    error_response: {
        error: {
            type: 'boolean',
            required: true,
            description: '是否发生错误',
            example: true
        },
        error_message: {
            type: 'string',
            required: true,
            description: '错误描述',
            example: '输入数据不足以生成准确的公式'
        },
        fallback_response: {
            type: 'object',
            required: false,
            description: '备用响应内容',
            properties: {
                formulas: {
                    type: 'array',
                    description: '基础公式建议'
                },
                data_analysis: {
                    type: 'object',
                    description: '基础数据分析'
                },
                alternative_formulas: {
                    type: 'array',
                    description: '替代方案'
                }
            }
        },
        metadata: {
            type: 'object',
            required: false,
            description: '元数据',
            properties: {
                model: {
                    type: 'string',
                    description: 'AI模型名称'
                },
                timestamp: {
                    type: 'string',
                    format: 'date-time',
                    description: '响应时间戳'
                }
            }
        }
    },
    
    // 系统提示词模板
    system_prompt_template: {
        role: 'system',
        content: `你是一个专业的Excel公式专家助手。你的任务是根据用户的需求和提供的数据信息，生成精确的Excel公式。

你的回答必须严格按照以下JSON格式返回，不要包含任何其他文本：

{formulas: [{title: 公式名称, formula: 完整公式, explanation: 详细说明, confidence: 95, applicable_ranges: [适用范围], required_functions: [使用函数], example: 使用示例}], data_analysis: {headers_found: [表头列表], data_types: [数据类型], recommendations: [建议]}, alternative_formulas: [{description: 方案描述, formula: 替代公式, pros: [优点], cons: [缺点]}]}

重要规则：
1. 公式必须使用正确的Excel语法
2. 跨工作簿引用使用 '[工作簿名]工作表名!单元格引用' 格式
3. 跨工作表引用使用 '工作表名!单元格引用' 格式
4. 考虑引用锁定方式（相对/绝对引用）
5. 填充时公式引用需要相应调整
6. confidence值应在70-99之间
7. 如数据不足，请说明并提供一般建议`
    }
};

// 示例数据
const AIJsonExamples = {
    // 请求示例
    request_example: {
        description: "计算每个产品的销售总额，需要单价乘以数量",
        referenceType: "current",
        currentCell: {
            cellAddress: "D2",
            row: 2,
            column: 4,
            worksheet: "销售数据"
        },
        selectedWorkbooks: [],
        selectedWorksheets: [],
        fillOptions: {
            right: false,
            down: true
        },
        headers: [
            {
                workbook: "销售数据.xlsx",
                worksheet: "销售数据",
                headers: [
                    { column: 1, value: "产品名称", dataType: "text" },
                    { column: 2, value: "销售数量", dataType: "number" },
                    { column: 3, value: "单价", dataType: "currency" }
                ]
            }
        ]
    },
    
    // 响应示例
    response_example: {
        formulas: [
            {
                title: "销售总额计算",
                formula: "=B2*C2",
                explanation: "将当前行的销售数量(B列)乘以单价(C列)，得出该产品的销售总额",
                confidence: 98,
                applicable_ranges: ["数值列", "连续数据"],
                required_functions: ["基础乘法"],
                example: "在D2中输入=B2*C2，拖拽填充到D3:D10"
            },
            {
                title: "绝对引用版本",
                formula: "=$B2*$C$2",
                explanation: "使用混合引用，数量列相对引用，单价列绝对引用",
                confidence: 85,
                applicable_ranges: ["需要固定单价的情况"],
                required_functions: ["混合引用"],
                example: "适用于单价固定的情况"
            }
        ],
        data_analysis: {
            headers_found: ["产品名称", "销售数量", "单价", "销售总额"],
            data_types: ["文本", "数字", "货币", "计算结果"],
            recommendations: [
                "数据连续无空值，适合使用填充功能",
                "建议在填充前验证第一个公式的正确性"
            ]
        },
        alternative_formulas: [
            {
                description: "使用PRODUCT函数",
                formula: "=PRODUCT(B2,C2)",
                pros: ["函数名称更清晰", "易于理解"],
                cons: ["可能影响兼容性"]
            }
        ],
        metadata: {
            model: "abab6.5s-chat",
            tokens_used: 234,
            timestamp: "2024-01-15T10:30:00Z",
            request_id: "req_123456789"
        }
    },
    
    // 跨工作簿请求示例
    cross_workbook_request: {
        description: "汇总所有工作表的产品销售额",
        referenceType: "workbook",
        currentCell: {
            cellAddress: "汇总.C2",
            row: 2,
            column: 3,
            worksheet: "汇总"
        },
        selectedWorkbooks: [
            {
                name: "产品A销售.xlsx",
                worksheets: [
                    {
                        name: "Sheet1",
                        usedRange: { rows: 50, columns: 5 },
                        headers: [
                            { column: 1, value: "产品名称" },
                            { column: 2, value: "数量" },
                            { column: 3, value: "单价" }
                        ]
                    }
                ]
            },
            {
                name: "产品B销售.xlsx", 
                worksheets: [
                    {
                        name: "数据",
                        usedRange: { rows: 30, columns: 5 },
                        headers: [
                            { column: 1, value: "产品名称" },
                            { column: 2, value: "数量" },
                            { column: 3, value: "单价" }
                        ]
                    }
                ]
            }
        ],
        selectedWorksheets: [],
        fillOptions: {
            right: false,
            down: true
        },
        headers: []
    }
};

// 验证函数
const AIJsonValidator = {
    /**
     * 验证请求数据格式
     */
    validateRequest(data) {
        const errors = [];
        
        if (!data.description || typeof data.description !== 'string') {
            errors.push('description是必需的字符串字段');
        }
        
        if (!data.referenceType || !['current', 'worksheet', 'workbook'].includes(data.referenceType)) {
            errors.push('referenceType必须是current/worksheet/workbook之一');
        }
        
        if (!data.currentCell || typeof data.currentCell !== 'object') {
            errors.push('currentCell是必需的object字段');
        }
        
        if (!data.fillOptions || typeof data.fillOptions !== 'object') {
            errors.push('fillOptions是必需的object字段');
        }
        
        if (data.fillOptions && (typeof data.fillOptions.right !== 'boolean' || typeof data.fillOptions.down !== 'boolean')) {
            errors.push('fillOptions的right和down必须是boolean类型');
        }
        
        return {
            isValid: errors.length === 0,
            errors: errors
        };
    },
    
    /**
     * 验证响应数据格式
     */
    validateResponse(data) {
        const errors = [];
        
        if (!data.formulas || !Array.isArray(data.formulas)) {
            errors.push('formulas是必需的array字段');
        } else {
            data.formulas.forEach((formula, index) => {
                if (!formula.title || !formula.formula || !formula.explanation) {
                    errors.push(`第${index + 1}个公式缺少必要字段`);
                }
                
                if (formula.confidence && (typeof formula.confidence !== 'number' || formula.confidence < 0 || formula.confidence > 100)) {
                    errors.push(`第${index + 1}个公式的confidence值必须在0-100之间`);
                }
            });
        }
        
        return {
            isValid: errors.length === 0,
            errors: errors
        };
    }
};

// 导出规范和工具
window.AIJsonSpec = AIJsonSpec;
window.AIJsonExamples = AIJsonExamples;
window.AIJsonValidator = AIJsonValidator;