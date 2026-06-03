import express from 'express';
import multer from 'multer';
import { read, utils } from 'xlsx';
import { fileURLToPath } from 'url';
import { dirname, join } from 'path';

const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);

const app = express();

const storage = multer.diskStorage({
  destination: (req, file, cb) => {
    cb(null, 'uploads/');
  },
  filename: (req, file, cb) => {
    cb(null, `${Date.now()}-${file.originalname}`);
  }
});

const upload = multer({
  storage,
  limits: { fileSize: 10 * 1024 * 1024 }
});

const SHEETS_CONFIG = {
  '一': {
    name: 'BOM 物料清单',
    requiredHeaders: ['导入结果', '导入类型', '工程物料', '说明', '物料类型', '物料组', '工程物料版本', '工程物料版本说明', '物料信号', '项目经理', '生效时间', '单位', '材料', '尺寸', '标准', '重量单位'],
    requiredFields: ['导入类型', '工程物料', '说明', '物料类型', '物料组', '工程物料版本', '物料信号', '项目经理', '生效时间', '单位', '重量单位'],
    forbiddenFields: ['工程物料版本说明', '材料', '尺寸', '标准', '导入结果'],
    allowedValues: {
      '导入类型': ['I', 'D', 'U'],
      '物料类型': ['1', '2', '3'],
      '物料组': ['101101', '101201', '101301', '101501', '101601', '101602', '102001', '102003', '102004', '901001'],
      '物料信号': ['EM', 'GM', 'DS', 'ME', 'FU'],
      '项目经理': ['JA097', 'JA129', 'JA166', 'JA167', 'JA177', 'JA225', 'JB012', 'JC099', 'JD099', 'JE368', 'JK098', 'HL268', 'HN227', 'HN046'],
      '单位': ['PCS', 'MM'],
      '重量单位': ['KG', 'G']
    },
    materialTypeGroups: {
      '1': ['101101', '101201', '101301', '101501', '101601', '101602', '901001', '211101'],
      '2': ['102001', '102004'],
      '3': ['102002', '102003']
    }
  },
  '二': {
    name: '物料通用数据和客户料号导入',
    requiredHeaders: ['导入结果', '导入类型', '物料名称', '物料类型 (1 采购 2 自制 3 成本)', '物料组', '单位集', '单位', '重量单位', '原材料', '标准', '大小', '产品类型', '产品分类', '产品大类', '制造商', '物料信号', '商品代码', '项目经理', '批次控制', '项目编号', '项目系列', '保质期/呆滞期', '计划数据默认仓库', '计划数据计划员', '订货数据仓库', '订货数据计划员', '订货数据车间计划员', '工艺路线代码', '备注 1', '备注 2', '备注 3', '物料代码系统', '客户代码', '业务伙伴物料代码', '客户物料开票名称'],
    requiredFields: ['导入类型', '物料号', '物料名称', '物料类型 (1 采购 2 自制 3 成本)', '物料组', '单位集', '单位', '重量单位', '产品类型', '产品分类', '物料信号', '商品代码', '项目经理', '批次控制', '保质期/呆滞期', '计划数据默认仓库', '计划数据计划员', '订货数据仓库', '订货数据计划员', '订货数据车间计划员'],
    forbiddenFields: ['原材料', '标准', '大小', '备注 1', '备注 2', '备注 3', '客户物料开票名称', '导入结果', '项目编号', '项目系列'],
    allowedValues: {
      '导入类型': ['I', 'D', 'U'],
      '物料类型 (1 采购 2 自制 3 成本)': ['1', '2', '3'],
      '物料组': ['101101', '101201', '101301', '101501', '101601', '101602', '102001', '102003', '102004', '901001', '211101'],
      '单位': ['PCS', 'KG', 'G', 'MM'],
      '重量单位': ['KG', 'G'],
      '物料信号': ['EM', 'GM', 'DS', 'ME', 'FU'],
      '项目经理': ['JA097', 'JA129', 'JA166', 'JA167', 'JA177', 'JA225', 'JB012', 'JC099', 'JD099', 'JE368', 'JK098', 'HL268', 'HN227', 'HN046'],
      '单位集': ['A01'],
      '计划数据计划员': ['JA153', 'JA303', 'JA081', 'JA097', 'JB023', 'JA202', 'HG031'],
      '订货数据计划员': ['JA153', 'JA303', 'JA081', 'JA097', 'JB023', 'JA202', 'HG031']
    }
  },
  '三': {
    name: '工程版本 BOM 审核及制造 BOM 导入',
    requiredHeaders: ['导入结果', '导入类型', '工程物料', '版本', '类型', '组件', '净数量', '废品率', '是否虚拟', '子件仓库', '工序', '生效时间', '失效时间'],
    requiredFields: ['导入类型', '工程物料', '版本', '类型', '组件', '净数量', '是否虚拟', '子件仓库', '工序', '生效时间', '失效时间'],
    forbiddenFields: ['导入结果'],
    allowedValues: {
      '导入类型': ['I', 'D', 'U']
    }
  },
  '四': {
    name: '物料工艺流程导入',
    requiredHeaders: ['导入结果', '导入类型', '制造物料', '工艺流程', '工艺流程说明', '工序', '任务', '工作中心', '机器', '计数点', '生产周期 (分钟)', '生效日期', '失效日期', '任务分类'],
    requiredFields: ['导入类型', '制造物料', '工艺流程', '工艺流程说明', '工序', '任务', '工作中心', '机器', '计数点', '生产周期 (分钟)', '生效日期', '失效日期', '任务分类'],
    forbiddenFields: ['导入结果'],
    allowedValues: {
      '导入类型': ['I', 'D', 'U']
    }
  },
  '五': {
    name: '物料扩展属性导入',
    requiredHeaders: ['导入结果', '导入类型', '物料', '物料开票名称', '产品行业', '行业细分', 'HS 编码', 'SPM', 'PIN', '模穴数', '复制模模穴数', '注塑周期', '复制注塑周期', '步距', '密度', '料宽', '电镀产速（m/min）', '产品塑胶重量', '料头重量', '复制模料头重量', '回料百分比', '模号', '复制模号', '用金量', '理论边角料', '机台效率', '产品理论重量', '模具原因', '产品应用', '项目名称', '麦途物料号', '电镀方式', '电镀类型'],
    requiredFields: ['导入类型', '物料', '产品行业', '行业细分', '产品理论重量', '产品应用', '项目名称'],
    forbiddenFields: ['物料开票名称', 'HS 编码', '密度', '料宽', '用金量', '机台效率', '模具原因', '麦途物料号', '导入结果'],
    allowedValues: {
      '导入类型': ['I', 'D', 'U']
    }
  }
};

app.use(express.json());
app.use(express.static(join(__dirname, 'dist')));

function findHeaderRow(data, requiredHeaders) {
  for (let i = 0; i < Math.min(data.length, 10); i++) {
    const row = data[i].map(h => String(h).trim());
    const matchCount = requiredHeaders.filter(h => row.includes(h)).length;
    if (matchCount >= requiredHeaders.length * 0.8) {
      return { index: i, headers: row };
    }
  }
  return null;
}

function normalizeHeader(header) {
  return String(header).trim().replace(/\s+/g, '');
}

function validateSheetData(bodyData, headers, config, startRowNumber, configKey) {
  const errors = [];
  
  const normalizedHeaders = headers.map(h => normalizeHeader(h));

  bodyData.forEach((row, rowIndex) => {
    const rowNumber = rowIndex + startRowNumber;

    const isRowEmpty = row.every(cell => cell === null || cell === undefined || String(cell).trim() === '');
    if (isRowEmpty) {
      return;
    }

    // 小数位数校验：所有数字最多 6 位小数
    row.forEach((cell, cellIndex) => {
      if (cell !== null && cell !== undefined && cell !== '') {
        const cellStr = String(cell).trim();
        if (/^-?\d+\.\d+$/.test(cellStr)) {
          const decimalPart = cellStr.split('.')[1];
          if (decimalPart && decimalPart.length > 6) {
            const headerName = headers[cellIndex] || `列${cellIndex + 1}`;
            errors.push({
              row: rowNumber,
              field: headerName,
              value: cellStr,
              error: '小数位数最多 6 位'
            });
          }
        }
      }
    });

    config.requiredFields.forEach(field => {
      const normalizedField = normalizeHeader(field);
      const index = normalizedHeaders.findIndex(h => h === normalizedField);
      const value = index !== -1 ? row[index] : undefined;
      
      if (index !== -1) {
        if (value === null || value === undefined || String(value).trim() === '') {
          const isSheetTwo = configKey === '二';
          const materialNoIndex = normalizedHeaders.findIndex(h => h === normalizeHeader('物料号'));
          const materialNo = materialNoIndex !== -1 ? (row[materialNoIndex] || '') : '';
          const materialNoStr = String(materialNo).trim();
          
          const hasWX = materialNoStr.toUpperCase().includes('WX');
          const isExpiryField = field === '保质期/呆滞期';
          
          if (isSheetTwo && isExpiryField && hasWX) {
            // 跳过此字段，不报错
          } else {
            errors.push({
              row: rowNumber,
              field,
              value: '',
              error: '必填项为空'
            });
          }
        }
      }
    });
    
    // 禁止填写字段校验
    if (config.forbiddenFields) {
      config.forbiddenFields.forEach(field => {
        const normalizedField = normalizeHeader(field);
        const index = normalizedHeaders.findIndex(h => h === normalizedField);
        const value = index !== -1 ? row[index] : undefined;
        
        if (index !== -1 && value !== null && value !== undefined && String(value).trim() !== '') {
          errors.push({
            row: rowNumber,
            field,
            value: String(value).trim(),
            error: '此字段不需要填写'
          });
        }
      });
    }
    
    // 枚举值校验
    if (config.allowedValues) {
      Object.keys(config.allowedValues).forEach(field => {
        const normalizedField = normalizeHeader(field);
        const fieldIndex = normalizedHeaders.findIndex(h => h === normalizedField);
        if (fieldIndex !== -1) {
          const value = row[fieldIndex];
          if (value && String(value).trim() !== '') {
            const valueStr = String(value).trim();
            if (!config.allowedValues[field].includes(valueStr)) {
              errors.push({
                row: rowNumber,
                field,
                value: valueStr,
                error: `无效值，只能填写：${config.allowedValues[field].join('、')}`
              });
            }
          }
        }
      });
    }

    // 表一特殊规则：导入类型与物料信号的匹配
    if (configKey === '一') {
      const importTypeIndex = normalizedHeaders.findIndex(h => h === normalizeHeader('导入类型'));
      const materialSignalIndex = normalizedHeaders.findIndex(h => h === normalizeHeader('物料信号'));
      
      if (importTypeIndex !== -1 && materialSignalIndex !== -1) {
        const importType = row[importTypeIndex] ? String(row[importTypeIndex]).trim() : '';
        const materialSignal = row[materialSignalIndex] ? String(row[materialSignalIndex]).trim() : '';
        
        if (importType === 'I' && materialSignal && materialSignal !== 'GM') {
          errors.push({
            row: rowNumber,
            field: '物料信号',
            value: materialSignal,
            error: '导入类型为 I 时，物料信号必须为 GM'
          });
        }
        
        if (materialSignal === 'GM' && importType && importType !== 'I') {
          errors.push({
            row: rowNumber,
            field: '导入类型',
            value: importType,
            error: '物料信号为 GM 时，导入类型必须为 I'
          });
        }
      }
    }

    // 表二特殊规则：计划数据默认仓库与订货数据仓库必须一致
    if (configKey === '二') {
      const defaultWarehouseIndex = normalizedHeaders.findIndex(h => h === normalizeHeader('计划数据默认仓库'));
      const orderWarehouseIndex = normalizedHeaders.findIndex(h => h === normalizeHeader('订货数据仓库'));
      
      if (defaultWarehouseIndex !== -1 && orderWarehouseIndex !== -1) {
        const defaultWarehouse = row[defaultWarehouseIndex] ? String(row[defaultWarehouseIndex]).trim() : '';
        const orderWarehouse = row[orderWarehouseIndex] ? String(row[orderWarehouseIndex]).trim() : '';
        
        if (defaultWarehouse && orderWarehouse && defaultWarehouse !== orderWarehouse) {
          errors.push({
            row: rowNumber,
            field: '计划数据默认仓库/订货数据仓库',
            value: `计划=${defaultWarehouse}, 订货=${orderWarehouse}`,
            error: '两个仓库字段必须一致'
          });
        }
      }
    }

    if (config.materialTypeGroups) {
      const typeField = Object.keys(config.allowedValues || {}).find(f => f.includes('物料类型'));
      if (typeField) {
        const normalizedType = normalizeHeader(typeField);
        const normalizedGroup = normalizeHeader('物料组');
        const typeIndex = normalizedHeaders.findIndex(h => h === normalizedType);
        const groupIndex = normalizedHeaders.findIndex(h => h === normalizedGroup);
        if (typeIndex !== -1 && groupIndex !== -1) {
          const materialType = row[typeIndex];
          const materialGroup = row[groupIndex];
          if (materialType && String(materialType).trim() !== '' && materialGroup && String(materialGroup).trim() !== '') {
            const typeStr = String(materialType).trim();
            const groupStr = String(materialGroup).trim();
            if (config.materialTypeGroups[typeStr] && !config.materialTypeGroups[typeStr].includes(groupStr)) {
              errors.push({
                row: rowNumber,
                field: '物料类型/物料组',
                value: `类型${typeStr} - 组${groupStr}`,
                error: `物料类型 ${typeStr} 对应的物料组不能是 ${groupStr}，应为：${config.materialTypeGroups[typeStr].join('、')}`
              });
            }
          }
        }
      }
    }
  });

  return errors;
}

function validateCrossSheets(sheetOne, sheetTwo) {
  const errors = [];
  
  const oneHeaders = sheetOne.headers.map(h => normalizeHeader(h));
  const twoHeaders = sheetTwo.headers.map(h => normalizeHeader(h));
  
  const oneEngIndex = oneHeaders.findIndex(h => h === normalizeHeader('工程物料'));
  const oneDescIndex = oneHeaders.findIndex(h => h === normalizeHeader('说明'));
  const oneTypeIndex = oneHeaders.findIndex(h => h === normalizeHeader('物料类型'));
  const oneGroupIndex = oneHeaders.findIndex(h => h === normalizeHeader('物料组'));
  const oneUnitIndex = oneHeaders.findIndex(h => h === normalizeHeader('单位'));
  const oneWeightUnitIndex = oneHeaders.findIndex(h => h === normalizeHeader('重量单位'));
  const oneImportTypeIndex = oneHeaders.findIndex(h => h === normalizeHeader('导入类型'));
  const oneEffectDateIndex = oneHeaders.findIndex(h => h === normalizeHeader('生效时间'));
  const oneProjectMgrIndex = oneHeaders.findIndex(h => h === normalizeHeader('项目经理'));

  const twoMaterialIndex = twoHeaders.findIndex(h => h === normalizeHeader('物料名称'));
  const twoDescIndex = twoHeaders.findIndex(h => h === normalizeHeader('说明'));
  const twoTypeIndex = twoHeaders.findIndex(h => h === normalizeHeader('物料类型 (1 采购 2 自制 3 成本)'));
  const twoGroupIndex = twoHeaders.findIndex(h => h === normalizeHeader('物料组'));
  const twoUnitIndex = twoHeaders.findIndex(h => h === normalizeHeader('单位'));
  const twoWeightUnitIndex = twoHeaders.findIndex(h => h === normalizeHeader('重量单位'));
  const twoImportTypeIndex = twoHeaders.findIndex(h => h === normalizeHeader('导入类型'));
  const twoEffectDateIndex = twoHeaders.findIndex(h => h === normalizeHeader('生效时间'));
  const twoProjectMgrIndex = twoHeaders.findIndex(h => h === normalizeHeader('项目经理'));

  if (oneEngIndex === -1 || twoMaterialIndex === -1) return errors;

  const engMaterialMap = new Map();
  sheetOne.data.forEach((row, idx) => {
    const engMaterial = row[oneEngIndex];
    if (engMaterial && String(engMaterial).trim() !== '') {
      engMaterialMap.set(String(engMaterial).trim(), {
        row: idx + sheetOne.headerRowIndex + 2,
        desc: row[oneDescIndex] || '',
        type: row[oneTypeIndex] || '',
        group: row[oneGroupIndex] || '',
        unit: row[oneUnitIndex] || '',
        weightUnit: row[oneWeightUnitIndex] || '',
        importType: oneImportTypeIndex !== -1 ? (row[oneImportTypeIndex] || '') : '',
        effectDate: oneEffectDateIndex !== -1 ? (row[oneEffectDateIndex] || '') : '',
        projectMgr: oneProjectMgrIndex !== -1 ? (row[oneProjectMgrIndex] || '') : ''
      });
    }
  });

  sheetTwo.data.forEach((row, idx) => {
    const rowNumber = idx + sheetTwo.headerRowIndex + 2;
    const materialName = row[twoMaterialIndex];
    
    if (!materialName || String(materialName).trim() === '') return;
    
    const materialNameStr = String(materialName).trim();
    
    if (engMaterialMap.has(materialNameStr)) {
      const oneData = engMaterialMap.get(materialNameStr);
      
      const twoUnit = row[twoUnitIndex] || '';
      const twoWeightUnit = row[twoWeightUnitIndex] || '';
      const twoDesc = row[twoDescIndex] || '';
      const twoType = row[twoTypeIndex] || '';
      const twoGroup = row[twoGroupIndex] || '';

      if (twoUnit && twoUnit !== oneData.unit) {
        errors.push({
          row: rowNumber,
          field: '单位',
          value: twoUnit,
          error: `与表一不一致，应为：${oneData.unit}（表一工程物料${materialNameStr}的单位为${oneData.unit}）`
        });
      }

      if (twoWeightUnit && twoWeightUnit !== oneData.weightUnit) {
        errors.push({
          row: rowNumber,
          field: '重量单位',
          value: twoWeightUnit,
          error: `与表一不一致，应为：${oneData.weightUnit}（表一工程物料${materialNameStr}的重量单位为${oneData.weightUnit}）`
        });
      }

      if (twoDesc && twoDesc !== oneData.desc) {
        errors.push({
          row: rowNumber,
          field: '说明/物料名称',
          value: twoDesc,
          error: `与表一不一致，应为：${oneData.desc}（表一工程物料${materialNameStr}的说明为${oneData.desc}）`
        });
      }

      if (twoType && twoType !== oneData.type) {
        errors.push({
          row: rowNumber,
          field: '物料类型',
          value: twoType,
          error: `与表一不一致，应为：${oneData.type}（表一工程物料${materialNameStr}的物料类型为${oneData.type}）`
        });
      }

      if (twoGroup && twoGroup !== oneData.group) {
        errors.push({
          row: rowNumber,
          field: '物料组',
          value: twoGroup,
          error: `与表一不一致，应为：${oneData.group}（表一工程物料${materialNameStr}的物料组为${oneData.group}）`
        });
      }

      const oneDescStr = oneData.desc || '';
      const twoMaterialNameStr = twoMaterialName || '';
      if (oneDescStr && twoMaterialNameStr) {
        if (oneDescStr.includes('  ') || twoMaterialNameStr.includes('  ')) {
          errors.push({
            row: rowNumber,
            field: '说明/物料名称',
            value: `表一="${oneDescStr}", 表二="${twoMaterialNameStr}"`,
            error: '说明与物料名称中不允许存在连续空格'
          });
        }
      }

      const twoImportType = twoImportTypeIndex !== -1 ? (row[twoImportTypeIndex] || '') : '';
      const twoEffectDate = twoEffectDateIndex !== -1 ? (row[twoEffectDateIndex] || '') : '';
      const twoProjectMgr = twoProjectMgrIndex !== -1 ? (row[twoProjectMgrIndex] || '') : '';
      
      const twoImportTypeStr = String(twoImportType).trim();
      const twoEffectDateStr = String(twoEffectDate).trim();
      const twoProjectMgrStr = String(twoProjectMgr).trim();
      
      if (twoImportTypeStr && oneData.importType && twoImportTypeStr !== oneData.importType) {
        errors.push({
          row: rowNumber,
          field: '导入类型',
          value: twoImportTypeStr,
          error: `与表一不一致，应为：${oneData.importType}`
        });
      }
      
      if (twoEffectDateStr && oneData.effectDate && twoEffectDateStr !== oneData.effectDate) {
        errors.push({
          row: rowNumber,
          field: '生效时间',
          value: twoEffectDateStr,
          error: `与表一不一致，应为：${oneData.effectDate}`
        });
      }
      
      if (twoProjectMgrStr && oneData.projectMgr && twoProjectMgrStr !== oneData.projectMgr) {
        errors.push({
          row: rowNumber,
          field: '项目经理',
          value: twoProjectMgrStr,
          error: `与表一不一致，应为：${oneData.projectMgr}`
        });
      }
    }
  });

  return errors;
}

function validateSheetThreeToSheetFive(sheetThree, sheetFive) {
  const errors = [];
  if (!sheetThree || !sheetFive) return errors;

  const threeNormHeaders = sheetThree.headers.map(h => normalizeHeader(h));
  const fiveNormHeaders = sheetFive.headers.map(h => normalizeHeader(h));

  const threeEngIndex = threeNormHeaders.findIndex(h => h === normalizeHeader('工程物料'));
  const threeCompIndex = threeNormHeaders.findIndex(h => h === normalizeHeader('组件'));
  const fiveMatIndex = fiveNormHeaders.findIndex(h => h === normalizeHeader('物料'));
  const fiveRecycleIndex = fiveNormHeaders.findIndex(h => h === normalizeHeader('回料百分比'));

  if (threeEngIndex === -1 || fiveMatIndex === -1 || threeCompIndex === -1 || fiveRecycleIndex === -1) {
    return errors;
  }

  const engMatToNeedsRecycle = new Map();
  sheetThree.data.forEach(row => {
    const eng = row[threeEngIndex];
    const comp = row[threeCompIndex];
    if (eng && String(eng).trim() !== '') {
      const engStr = String(eng).trim();
      const compStr = comp ? String(comp).trim() : '';
      if (compStr.includes('R1') || compStr.includes('R2')) {
        engMatToNeedsRecycle.set(engStr, true);
      }
    }
  });

  sheetFive.data.forEach((row, idx) => {
    const mat = row[fiveMatIndex];
    if (mat && String(mat).trim() !== '') {
      const matStr = String(mat).trim();
      if (engMatToNeedsRecycle.has(matStr)) {
        const recycle = row[fiveRecycleIndex];
        if (!recycle || String(recycle).trim() === '') {
          errors.push({
            row: idx + sheetFive.headerRowIndex + 2,
            field: '回料百分比',
            value: '',
            error: '表三中存在对应组件以R1或R2结尾，此字段必须填写'
          });
        }
      }
    }
  });

  return errors;
}

function validateSheetThreeSpecial(data, headers, startRowNumber) {
  const errors = [];
  const normalizedHeaders = headers.map(h => normalizeHeader(h));
  
  const componentIndex = normalizedHeaders.findIndex(h => h === normalizeHeader('组件'));
  const isVirtualIndex = normalizedHeaders.findIndex(h => h === normalizeHeader('是否虚拟'));

  data.forEach((row, rowIndex) => {
    const rowNumber = rowIndex + startRowNumber;
    
    const isRowEmpty = row.every(cell => cell === null || cell === undefined || String(cell).trim() === '');
    if (isRowEmpty) return;
    
    if (componentIndex !== -1 && isVirtualIndex !== -1) {
      const component = row[componentIndex];
      const isVirtual = row[isVirtualIndex];
      
      if (!component) return;
      
      const componentStr = String(component).trim();
      const isVirtualStr = isVirtual ? String(isVirtual).trim() : '';
      
      if (componentStr === 'PPM000100' || componentStr === 'PPM000300') {
        if (isVirtualStr !== '1') {
          errors.push({
            row: rowNumber,
            field: '是否虚拟',
            value: isVirtualStr,
            error: `组件为${componentStr}时，是否虚拟必须为 1`
          });
        }
      } else if (isVirtualStr && isVirtualStr !== '2') {
        errors.push({
          row: rowNumber,
          field: '是否虚拟',
          value: isVirtualStr,
          error: `组件为${componentStr}时，是否虚拟必须为 2`
        });
      }
    }

    // 废品率校验
    const scrapRateIndex = normalizedHeaders.findIndex(h => h === normalizeHeader('废品率'));
    
    if (scrapRateIndex !== -1) {
      const scrapRate = row[scrapRateIndex];
      const scrapRateStr = scrapRate !== null && scrapRate !== undefined ? String(scrapRate).trim() : '';
      
      // Identify relevant columns
      const compIdx = normalizedHeaders.findIndex(h => h.includes(normalizeHeader('组件')));
      const engIdx = normalizedHeaders.findIndex(h => h.includes(normalizeHeader('工程物料')));

      const compStr = compIdx !== -1 ? String(row[compIdx] || '').trim() : '';
      const engStr = engIdx !== -1 ? String(row[engIdx] || '').trim() : '';

      // Specific Materials List
      const specificMaterials = ['ABS', 'COC', 'HTNFR', 'LCP', 'MPPO', 'LDS', 'PA', 'PBT', 'PC', 'PCABS', 'PP', 'PPA', 'PPO', 'PPS', 'TPX', 'TPE'];
      const forbiddenPpmCodes = ['PPM000100', 'PPM000201', 'PPM000300', 'PPM000401'];

      // Priority 1: Forbidden (Highest)
      // Component contains 'RR', 'FL', or 'FLD'
      const isForbidden = compStr.includes('RR') || compStr.includes('FL') || compStr.includes('FLD') || forbiddenPpmCodes.includes(compStr);

      // Priority 2: Material Specific (3.5) - Higher than P/M/S ending
      // Component contains specific materials
      const isSpecificMat = specificMaterials.some(m => compStr.includes(m));

      // Priority 3: P or M + digit ending (2.5)
      // Checks Engineering Material, falling back to Component if empty
      const targetPmStr = engStr || compStr;
      const pmMatch = targetPmStr.match(/([PM])\d$/);
      const isEndingPM = !!pmMatch;

      // Priority 4: Standard Material Code Pattern (1.5) - Metal Alloy Grades
      // Format: Prefix(letters) + Number + Hyphen + Suffix(number or letter+number)
      // Examples: SUS301-055, C5191-032, CUSN6-S22, QSN65-025, T2-S03
      const hasHyphen = compStr.includes('-') || engStr.includes('-');
      const metalGradePattern = /^[A-Za-z]+\d+(-[A-Za-z]?\d+)+$/;
      const isMetalGrade = metalGradePattern.test(compStr) || metalGradePattern.test(engStr);
      const isStandardCode = hasHyphen && isMetalGrade;

      // Rule 1: Forbidden - scrap rate must be empty
      if (isForbidden) {
         if (scrapRateStr) {
           errors.push({
             row: rowNumber,
             field: '废品率',
             value: scrapRateStr,
             error: '包含禁用废品率的代码，不需要填写'
           });
         }
      }
      // Rule 2: Specific Materials -> 3.5
      else if (isSpecificMat) {
         if (scrapRateStr !== '3.5') {
           errors.push({
             row: rowNumber,
             field: '废品率',
             value: scrapRateStr || '',
             error: '组件包含指定材料，废品率必须填写 3.5'
           });
         }
      }
      // Rule 3: P/M Ending -> 2.5
      else if (isEndingPM) {
         if (scrapRateStr !== '2.5') {
           errors.push({
             row: rowNumber,
             field: '废品率',
             value: scrapRateStr || '',
             error: `物料倒数两位为${pmMatch[1]}+数字，废品率必须填写 2.5`
           });
         }
       }
       // Rule 4: Metal Grade (Standard Material Code) -> 1.5
       else if (isStandardCode) {
          if (scrapRateStr !== '1.5') {
            errors.push({
              row: rowNumber,
              field: '废品率',
              value: scrapRateStr || '',
              error: '标准物料编号（金属材料牌号），废品率必须填写 1.5'
            });
          }
       }
     }
    });

  return errors;
}

function validateSheetTwoInternal(data, headers, startRowNumber) {
  const errors = [];
  const normalizedHeaders = headers.map(h => normalizeHeader(h));
  
  const planPlannerIndex = normalizedHeaders.findIndex(h => h === normalizeHeader('计划数据计划员'));
  const orderPlannerIndex = normalizedHeaders.findIndex(h => h === normalizeHeader('订货数据计划员'));

  data.forEach((row, rowIndex) => {
    const rowNumber = rowIndex + startRowNumber;
    
    const isRowEmpty = row.every(cell => cell === null || cell === undefined || String(cell).trim() === '');
    if (isRowEmpty) return;
    
    if (planPlannerIndex !== -1 && orderPlannerIndex !== -1) {
      const planPlanner = row[planPlannerIndex];
      const orderPlanner = row[orderPlannerIndex];
      
      const planPlannerStr = planPlanner ? String(planPlanner).trim() : '';
      const orderPlannerStr = orderPlanner ? String(orderPlanner).trim() : '';
      
      if (planPlannerStr && orderPlannerStr && planPlannerStr !== orderPlannerStr) {
        errors.push({
          row: rowNumber,
          field: '计划数据计划员/订货数据计划员',
          value: `计划=${planPlannerStr}, 订货=${orderPlannerStr}`,
          error: `两个字段填写不一致`
        });
      }
    }
  });

  return errors;
}

function validateSheetFourSpecial(data, headers, startRowNumber) {
  const errors = [];
  const normalizedHeaders = headers.map(h => normalizeHeader(h));
  
  const mfgIndex = normalizedHeaders.findIndex(h => h === normalizeHeader('制造物料'));
  const processIndex = normalizedHeaders.findIndex(h => h === normalizeHeader('工艺流程说明'));
  const taskCategoryIndex = normalizedHeaders.findIndex(h => h === normalizeHeader('任务分类'));
  
  const processRules = [
    { code: 'S', keywords: ['冲压'] },
    { code: 'P', keywords: ['电镀'] },
    { code: 'M', keywords: ['注塑', '成型'] },
    { code: 'A', keywords: ['组装'] },
    { code: 'H', keywords: ['热处理'] },
    { code: 'J', keywords: ['压接'] },
    { code: 'C', keywords: ['裁切'] },
    { code: 'W', keywords: ['焊接'] },
    { code: 'R', keywords: ['铆接'] },
    { code: 'E', keywords: ['蚀刻'] },
    { code: 'B', keywords: ['折弯'] },
    { code: 'L', keywords: ['剥金'] },
    { code: 'Q', keywords: ['清洗'] },
    { code: 'G', keywords: ['刮料', '异型材加工'] },
    { code: 'Z', keywords: ['包装'] },
    { code: 'N', keywords: ['其他'] },
    { code: 'D', keywords: ['去毛刺'] },
    { code: 'BL', keywords: ['拌料'] },
    { code: 'SX', keywords: ['筛选'] },
    { code: 'FC', keywords: ['全检'] }
  ];
  
  data.forEach((row, rowIndex) => {
    const rowNumber = rowIndex + startRowNumber;
    
    const isRowEmpty = row.every(cell => cell === null || cell === undefined || String(cell).trim() === '');
    if (isRowEmpty) return;
    
    const mfg = mfgIndex !== -1 ? (row[mfgIndex] || '') : '';
    const process = processIndex !== -1 ? (row[processIndex] || '') : '';
    const taskCategory = taskCategoryIndex !== -1 ? (row[taskCategoryIndex] || '') : '';
    
    const mfgStr = String(mfg).trim();
    const processStr = String(process).trim();
    const taskCategoryStr = String(taskCategory).trim();
    
    if (mfgStr && processIndex !== -1 && mfgStr.length >= 2) {
      // 规则生效条件：制造物料最后一位必须是阿拉伯数字
      const lastChar = mfgStr.slice(-1);
      const isLastDigit = /^[0-9]$/.test(lastChar);

      if (isLastDigit) {
        // 根据倒数第二位的特征码匹配规则
        // 匹配所有可能的特征码（处理代码重叠情况，例如 'C' 和 'FC'）
        const mfgUpper = mfgStr.toUpperCase();
        const matches = processRules.filter(rule => rule.code && mfgUpper.endsWith(rule.code + lastChar));
        
        if (matches.length > 0) {
          // 优先匹配更长的特征码（例如 'FC' 优先于 'C'）
          matches.sort((a, b) => b.code.length - a.code.length);
          const matchedRule = matches[0];

          const hasKeyword = matchedRule.keywords.some(kw => processStr.includes(kw));
          if (!hasKeyword) {
            errors.push({
              row: rowNumber,
              field: '工艺流程说明',
              value: processStr,
              error: `制造物料包含特征码"${matchedRule.code}"，工艺流程说明应包含"${matchedRule.keywords.join('或')}"`
            });
          }
        }
      }
    }
    
    // 工作中心为 PL0GXT 时，任务分类必须包含"电镀 GXT"字符
    const workCenterIndex = normalizedHeaders.findIndex(h => h === normalizeHeader('工作中心'));
    if (workCenterIndex !== -1 && taskCategoryIndex !== -1) {
      const workCenter = row[workCenterIndex] ? String(row[workCenterIndex]).trim() : '';
      const taskCategoryRaw = row[taskCategoryIndex] ? String(row[taskCategoryIndex]).trim() : '';
      
      if (workCenter === 'PL0GXT' && taskCategoryRaw) {
        const hasPlating = taskCategoryRaw.includes('电镀');
        const hasGXT = taskCategoryRaw.toUpperCase().includes('GXT');
        
        if (!hasPlating || !hasGXT) {
          errors.push({
            row: rowNumber,
            field: '任务分类',
            value: taskCategoryRaw,
            error: '工作中心为 PL0GXT 时，任务分类必须包含"电镀 GXT"字符'
          });
        }
      }
    }
  });

  return errors;
}

function validateCrossSheetsSPM(sheetFour, sheetFive) {
  const errors = [];
  
  if (!sheetFour || !sheetFive) {
    return errors;
  }
  
  // 校验表四"工艺流程"固定为 A10
  const normalizedHeaders = sheetFour.headers.map(normalizeHeader);
  const fourProcessA10Index = normalizedHeaders.findIndex(h => h === normalizeHeader('工艺流程'));
  if (fourProcessA10Index !== -1) {
    sheetFour.data.forEach((row, idx) => {
      const rowNumber = idx + sheetFour.headerRowIndex + 2;
      const isRowEmpty = row.every(cell => cell === null || cell === undefined || String(cell).trim() === '');
      if (isRowEmpty) return;

      const processVal = row[fourProcessA10Index];
      const processStr = processVal ? String(processVal).trim() : '';
      
      if (processStr && processStr !== 'A10') {
        errors.push({
          row: rowNumber,
          field: '工艺流程',
          value: processStr,
          error: '工艺流程固定为 A10'
        });
      }
    });
  }
  
  const fourMfgIndex = sheetFour.headers.map(h => normalizeHeader(h)).findIndex(h => h === normalizeHeader('制造物料'));
  const fourProcessIndex = sheetFour.headers.map(h => normalizeHeader(h)).findIndex(h => h === normalizeHeader('工艺流程说明'));
  const fourCycleIndex = sheetFour.headers.map(h => normalizeHeader(h)).findIndex(h => h === normalizeHeader('生产周期 (分钟)'));
  
  const fiveMaterialIndex = sheetFive.headers.map(h => normalizeHeader(h)).findIndex(h => h === normalizeHeader('物料'));
  const fiveSPMIndex = sheetFive.headers.map(h => normalizeHeader(h)).findIndex(h => h === normalizeHeader('SPM'));
  const fivePlatingSpeedIndex = sheetFive.headers.findIndex(h => normalizeHeader(h).includes(normalizeHeader('电镀产速')));
  const fiveMoldCavityIndex = sheetFive.headers.map(h => normalizeHeader(h)).findIndex(h => h === normalizeHeader('模穴数'));
  const fiveInjectionCycleIndex = sheetFive.headers.map(h => normalizeHeader(h)).findIndex(h => h === normalizeHeader('注塑周期'));
  const fiveStepIndex = sheetFive.headers.map(h => normalizeHeader(h)).findIndex(h => h === normalizeHeader('步距'));
  
  if (fourMfgIndex === -1 || fiveMaterialIndex === -1 || fiveSPMIndex === -1) {
    return errors;
  }
  
  const fourDataMap = new Map();
  sheetFour.data.forEach((row, idx) => {
    const mfg = row[fourMfgIndex];
    if (mfg && String(mfg).trim() !== '') {
      const mfgStr = String(mfg).trim();
      fourDataMap.set(mfgStr, {
        row: idx + sheetFour.headerRowIndex + 2,
        process: fourProcessIndex !== -1 ? (row[fourProcessIndex] || '') : '',
        cycle: fourCycleIndex !== -1 ? (row[fourCycleIndex] || '') : ''
      });
    }
  });
  
  let matchCount = 0;
  let platingCount = 0;
  
  sheetFive.data.forEach((row, idx) => {
    const rowNumber = idx + sheetFive.headerRowIndex + 2;
    const material = row[fiveMaterialIndex];
    const spm = row[fiveSPMIndex];
    const fivePlatingSpeed = fivePlatingSpeedIndex !== -1 ? (row[fivePlatingSpeedIndex] || '') : '';
    
    if (!material || String(material).trim() === '') return;
    const materialStr = String(material).trim();
    
    const fourData = fourDataMap.get(materialStr);
    if (!fourData) return;
    
    matchCount++;
    
    // 检查 SPM 是否为空
    const spmRaw = (spm !== null && spm !== undefined) ? String(spm).trim() : '';
    
    const process = fourData.process || '';
    const cycleStr = fourData.cycle || '';
    const cycleNum = parseFloat(cycleStr);
    const fourPlatingSpeed = fivePlatingSpeed;
    
    const isStamping = process.includes('冲压');
    const isMolding = process.includes('成型') || process.includes('注塑');
    const isPlating = process.includes('电镀');
    const hasSpecialProcess = isStamping || isMolding || isPlating;
    
    // 冲压工艺：SPM 必须填写
    if (isStamping && !spmRaw) {
      errors.push({
        row: rowNumber,
        field: 'SPM',
        value: '',
        error: '工艺流程包含"冲压"，SPM 必须填写'
      });
      return;
    }
    
    // 成型/注塑、电镀工艺：SPM 不能填写
    if ((isMolding || isPlating) && spmRaw) {
      errors.push({
        row: rowNumber,
        field: 'SPM',
        value: spmRaw,
        error: `工艺流程包含"${isMolding ? '成型/注塑' : '电镀'}"，SPM 不能填写`
      });
      return;
    }
    
    // 冲压工艺的 SPM 公式校验：(生产周期/1.2)*SPM=1
    if (isStamping) {
      const spmNum = parseFloat(spmRaw);
      if (isNaN(spmNum)) {
        return;
      }
      
      if (isPlating) platingCount++;
      
      if (cycleNum && !isNaN(cycleNum)) {
        const result = (cycleNum / 1.2) * spmNum;
        if (Math.abs(result - 1) > 0.01) {
          errors.push({
            row: rowNumber,
            field: 'SPM',
            value: spmNum.toString(),
            error: `工艺流程包含"冲压"，应满足公式：(生产周期 (${cycleNum})/1.2)*SPM=1，当前计算结果：${result.toFixed(4)}`
          });
        }
      }
    }
    
    // 成型/注塑工艺校验
    if (isMolding) {
      const fiveMoldCavity = fiveMoldCavityIndex !== -1 ? (row[fiveMoldCavityIndex] || '') : '';
      const fiveInjectionCycle = fiveInjectionCycleIndex !== -1 ? (row[fiveInjectionCycleIndex] || '') : '';
      
      // 模穴数和注塑周期必须填写
      if (!fiveMoldCavity || !fiveInjectionCycle) {
        errors.push({
          row: rowNumber,
          field: '模穴数/注塑周期',
          value: `模穴数=${fiveMoldCavity}, 注塑周期=${fiveInjectionCycle}`,
          error: '工艺流程包含"成型/注塑"，模穴数和注塑周期必须填写'
        });
      } else {
        const moldCavityNum = parseFloat(fiveMoldCavity);
        const injectionCycleNum = parseFloat(fiveInjectionCycle);
        
        if (!isNaN(moldCavityNum) && !isNaN(injectionCycleNum)) {
          const expectedCycle = (injectionCycleNum / moldCavityNum) / 60 * 1.1;
          if (cycleNum && !isNaN(cycleNum) && Math.abs(cycleNum - expectedCycle) > 0.01) {
            errors.push({
              row: rowNumber,
              field: '生产周期 (分钟)',
              value: cycleNum.toString(),
              error: `工艺流程包含"成型/注塑"，应满足公式：(注塑周期 (${injectionCycleNum})/模穴数 (${moldCavityNum}))/60*1.1=${expectedCycle.toFixed(4)}，当前表四生产周期：${cycleNum.toFixed(4)}`
            });
          }
        }
      }
    }
    // 电镀工艺校验
    else if (isPlating) {
      const fivePlatingSpeedVal = (fivePlatingSpeed !== null && fivePlatingSpeed !== undefined) ? String(fivePlatingSpeed).trim() : '';
      const fiveStepVal = fiveStepIndex !== -1 ? (row[fiveStepIndex] || '') : '';
      const fiveStepNum = parseFloat(String(fiveStepVal).trim());
      
      if (!fivePlatingSpeedVal) {
        errors.push({
          row: rowNumber,
          field: '电镀产速（m/min）',
          value: '',
          error: '工艺流程包含"电镀"，电镀产速必须填写'
        });
      } else if (!fiveStepVal || isNaN(fiveStepNum) || fiveStepNum === 0) {
        errors.push({
          row: rowNumber,
          field: '步距',
          value: String(fiveStepVal).trim(),
          error: '工艺流程包含"电镀"，步距必须填写且为有效数字'
        });
      } else {
        const platingSpeedNum = parseFloat(fivePlatingSpeedVal);
        if (!isNaN(platingSpeedNum) && platingSpeedNum !== 0) {
          const expectedCycle = (1 / (platingSpeedNum * 1000 / fiveStepNum)) * 1.05;
          
          if (cycleNum && !isNaN(cycleNum) && Math.abs(cycleNum - expectedCycle) > expectedCycle * 0.1) {
            errors.push({
              row: rowNumber,
              field: '生产周期 (分钟)',
              value: cycleNum.toString(),
              error: `工艺流程包含"电镀"，表四生产周期应等于 [1/（电镀产速 (${platingSpeedNum})×1000/步距 (${fiveStepNum})）]×1.05=${expectedCycle.toFixed(4)}，当前值：${cycleNum.toFixed(4)}`
            });
          }
        }
      }
    }
  });
  
  return errors;
}

function validateCrossSheetsTypeConsistency(sheetOne, sheetTwo) {
  const errors = [];
  // 表一
  const oneEngIndex = sheetOne.headers.map(h => normalizeHeader(h)).findIndex(h => h === normalizeHeader('工程物料'));
  const oneTypeIndex = sheetOne.headers.map(h => normalizeHeader(h)).findIndex(h => h.includes(normalizeHeader('物料类型')));
  if (oneEngIndex === -1 || oneTypeIndex === -1) return errors;

  const twoMaterialIndex = sheetTwo ? sheetTwo.headers.map(h => normalizeHeader(h)).findIndex(h => h === normalizeHeader('物料号')) : -1;
  const twoTypeIndex = sheetTwo ? sheetTwo.headers.map(h => normalizeHeader(h)).findIndex(h => h.includes(normalizeHeader('物料类型'))) : -1;
  
  const oneTypeMap = new Map();
  sheetOne.data.forEach((row, idx) => {
    const eng = row[oneEngIndex];
    const type = row[oneTypeIndex];
    if (eng && String(eng).trim() !== '') {
      oneTypeMap.set(String(eng).trim(), {
        row: idx + sheetOne.headerRowIndex + 2,
        type: type ? String(type).trim() : ''
      });
    }
  });

  const twoTypeMap = new Map();
  if (sheetTwo && twoMaterialIndex !== -1 && twoTypeIndex !== -1) {
    sheetTwo.data.forEach((row, idx) => {
      const material = row[twoMaterialIndex];
      const type = row[twoTypeIndex];
      if (material && String(material).trim() !== '') {
        twoTypeMap.set(String(material).trim(), {
          row: idx + sheetTwo.headerRowIndex + 2,
          type: type ? String(type).trim() : ''
        });
      }
    });
  }

  oneTypeMap.forEach((oneData, material) => {
    const twoData = twoTypeMap.get(material);
    if (twoData) {
      if (oneData.type !== twoData.type && oneData.type !== '' && twoData.type !== '') {
        errors.push({
          row: oneData.row,
          field: '物料类型',
          value: oneData.type,
          error: `与表二不一致：表二物料 ${material} 类型为 ${twoData.type}`
        });
      }
    }
  });

  return errors;
}

function validateCrossSheetsExistence(sheetOne, sheetTwo, sheetThree, sheetFour, sheetFive) {
  const errors = [];
  
  const oneEngIndex = sheetOne.headers.map(h => normalizeHeader(h)).findIndex(h => h === normalizeHeader('工程物料'));
  if (oneEngIndex === -1) return errors;
  
  const twoMaterialIndex = sheetTwo ? sheetTwo.headers.map(h => normalizeHeader(h)).findIndex(h => h === normalizeHeader('物料号')) : -1;
  const threeEngIndex = sheetThree ? sheetThree.headers.map(h => normalizeHeader(h)).findIndex(h => h === normalizeHeader('工程物料')) : -1;
  const fourMfgIndex = sheetFour ? sheetFour.headers.map(h => normalizeHeader(h)).findIndex(h => h === normalizeHeader('制造物料')) : -1;
  const fiveMaterialIndex = sheetFive ? sheetFive.headers.map(h => normalizeHeader(h)).findIndex(h => h === normalizeHeader('物料')) : -1;
  
  const oneProjectMgrIndex = sheetOne.headers.map(h => normalizeHeader(h)).findIndex(h => h === normalizeHeader('项目经理'));
  const twoProjectMgrIndex = sheetTwo ? sheetTwo.headers.map(h => normalizeHeader(h)).findIndex(h => h === normalizeHeader('项目经理')) : -1;
  
  const engMaterialSet = new Set();
  const engMaterialMap = new Map();
  
  sheetOne.data.forEach((row, idx) => {
    const engMaterial = row[oneEngIndex];
    if (engMaterial && String(engMaterial).trim() !== '') {
      const engStr = String(engMaterial).trim();
      engMaterialSet.add(engStr);
      engMaterialMap.set(engStr, {
        row: idx + sheetOne.headerRowIndex + 2,
        projectMgr: oneProjectMgrIndex !== -1 ? (row[oneProjectMgrIndex] || '') : ''
      });
    }
  });
  
  if (sheetThree && threeEngIndex !== -1) {
    const threeEngSet = new Set();
    sheetThree.data.forEach(row => {
      const eng = row[threeEngIndex];
      if (eng && String(eng).trim() !== '') {
        threeEngSet.add(String(eng).trim());
      }
    });
    
    engMaterialSet.forEach(eng => {
      if (!threeEngSet.has(eng)) {
        const oneData = engMaterialMap.get(eng);
        errors.push({
          row: oneData.row,
          field: '工程物料',
          value: eng,
          error: `此工程物料未出现在表三（工程版本 BOM 审核及制造 BOM 导入）的工程物料列中`
        });
      }
    });
  }
  
  if (sheetFour && fourMfgIndex !== -1) {
    const fourMfgSet = new Set();
    sheetFour.data.forEach(row => {
      const mfg = row[fourMfgIndex];
      if (mfg && String(mfg).trim() !== '') {
        fourMfgSet.add(String(mfg).trim());
      }
    });
    
    engMaterialSet.forEach(eng => {
      if (!fourMfgSet.has(eng)) {
        const oneData = engMaterialMap.get(eng);
        errors.push({
          row: oneData.row,
          field: '工程物料',
          value: eng,
          error: `此工程物料未出现在表四（物料工艺流程导入）的制造物料列中`
        });
      }
    });
  }
  
  if (sheetFive && fiveMaterialIndex !== -1) {
    const fiveMaterialSet = new Set();
    sheetFive.data.forEach(row => {
      const material = row[fiveMaterialIndex];
      if (material && String(material).trim() !== '') {
        fiveMaterialSet.add(String(material).trim());
      }
    });
    
    engMaterialSet.forEach(eng => {
      if (!fiveMaterialSet.has(eng)) {
        const oneData = engMaterialMap.get(eng);
        errors.push({
          row: oneData.row,
          field: '工程物料',
          value: eng,
          error: `此工程物料未出现在表五（物料扩展属性导入）的物料列中`
        });
      }
    });
  }
  
  if (sheetTwo && twoMaterialIndex !== -1 && twoProjectMgrIndex !== -1) {
    const twoMaterialMap = new Map();
    sheetTwo.data.forEach(row => {
      const material = row[twoMaterialIndex];
      const projectMgr = row[twoProjectMgrIndex];
      if (material && String(material).trim() !== '') {
        twoMaterialMap.set(String(material).trim(), projectMgr ? String(projectMgr).trim() : '');
      }
    });
    
    engMaterialSet.forEach(eng => {
      const oneData = engMaterialMap.get(eng);
      const twoProjectMgr = twoMaterialMap.get(eng);
      
      if (twoProjectMgr !== undefined) {
        const oneProjectMgrStr = oneData.projectMgr ? String(oneData.projectMgr).trim() : '';
        const twoProjectMgrStr = twoProjectMgr ? String(twoProjectMgr).trim() : '';
        
        if (oneProjectMgrStr !== twoProjectMgrStr) {
          errors.push({
            row: oneData.row,
            field: '项目经理',
            value: oneData.projectMgr,
            error: `与表二不一致，表二相同物料号的项目经理为：${twoProjectMgr}`
          });
        }
      }
    });
  }
  
  return errors;
}

function validateSheetFiveSpecial(data, headers, startRowNumber) {
  const errors = [];
  const normalizedHeaders = headers.map(h => normalizeHeader(h));

  const materialIndex = normalizedHeaders.findIndex(h => h === normalizeHeader('物料'));
  const spmIndex = normalizedHeaders.findIndex(h => h === normalizeHeader('SPM'));
  const pinIndex = normalizedHeaders.findIndex(h => h === normalizeHeader('PIN'));
  const moldCavityIndex = normalizedHeaders.findIndex(h => h === normalizeHeader('模穴数'));
  const injectionCycleIndex = normalizedHeaders.findIndex(h => h === normalizeHeader('注塑周期'));
  const platingSpeedIndex = normalizedHeaders.findIndex(h => h === normalizeHeader('电镀产速（m/min）'));
  const stepIndex = normalizedHeaders.findIndex(h => h === normalizeHeader('步距'));
  const moldNoIndex = normalizedHeaders.findIndex(h => h === normalizeHeader('模号'));
  const plasticWeightIndex = normalizedHeaders.findIndex(h => h === normalizeHeader('产品塑胶重量'));
  const sprueWeightIndex = normalizedHeaders.findIndex(h => h === normalizeHeader('料头重量'));
  const recycleIndex = normalizedHeaders.findIndex(h => h === normalizeHeader('回料百分比'));
  const replMoldCavityIndex = normalizedHeaders.findIndex(h => h === normalizeHeader('复制模模穴数'));
  const replInjectionCycleIndex = normalizedHeaders.findIndex(h => h === normalizeHeader('复制注塑周期'));
  const replSprueWeightIndex = normalizedHeaders.findIndex(h => h === normalizeHeader('复制模料头重量'));
  const replMoldNoIndex = normalizedHeaders.findIndex(h => h === normalizeHeader('复制模号'));
  const goldAmountIndex = normalizedHeaders.findIndex(h => h === normalizeHeader('用金量'));
  const scrapIndex = normalizedHeaders.findIndex(h => h === normalizeHeader('理论边角料'));
  const platingMethodIndex = normalizedHeaders.findIndex(h => h === normalizeHeader('电镀方式'));
  const platingTypeIndex = normalizedHeaders.findIndex(h => h === normalizeHeader('电镀类型'));

  // 新增字段：复制模回料百分比
  const replRecycleIndex = normalizedHeaders.findIndex(h => h === normalizeHeader('复制模回料百分比'));
  const replPlasticWeightIndex = normalizedHeaders.findIndex(h => h === normalizeHeader('复制模产品塑胶重量'));
  const categoryIndex = normalizedHeaders.findIndex(h => h === normalizeHeader('分类'));

  // 校验规则：倒数第二位为特征码，且最后一位必须是阿拉伯数字
  const getSuffixRule = (str) => {
    if (str.length < 2) return null;
    const lastChar = str.slice(-1);
    if (!/^[0-9]$/.test(lastChar)) return null;
    const secondLastChar = str.slice(-2, -1).toUpperCase();
    if (secondLastChar === 'S' || secondLastChar === 'M' || secondLastChar === 'P') {
      return secondLastChar;
    }
    return null;
  };

  data.forEach((row, rowIndex) => {
    const rowNumber = rowIndex + startRowNumber;

    const isRowEmpty = row.every(cell => cell === null || cell === undefined || String(cell).trim() === '');
    if (isRowEmpty) return;

    const material = materialIndex !== -1 ? (row[materialIndex] || '') : '';
    const materialStr = String(material).trim();
    if (!materialStr) return;

    const suffixType = getSuffixRule(materialStr);
    if (!suffixType) return;

    let forbiddenFields = [];
    let requiredFields = [];
    let moldConstraint = null;

    if (suffixType === 'S') {
      // 冲压类（S）
      forbiddenFields = [
        { name: '模穴数', index: moldCavityIndex },
        { name: '注塑周期', index: injectionCycleIndex },
        { name: '电镀产速（m/min）', index: platingSpeedIndex },
        { name: '回料百分比', index: recycleIndex },
        { name: '复制模回料百分比', index: replRecycleIndex },
        { name: '复制注塑周期', index: replInjectionCycleIndex },
        { name: '复制模模穴数', index: replMoldCavityIndex },
        { name: '复制模产品塑胶重量', index: replPlasticWeightIndex },
        { name: '复制模料头重量', index: replSprueWeightIndex },
        { name: '复制模号', index: replMoldNoIndex },
        { name: '产品塑胶重量', index: plasticWeightIndex },
        { name: '料头重量', index: sprueWeightIndex },
        { name: '用金量', index: goldAmountIndex },
        { name: '分类', index: categoryIndex }
      ];
      requiredFields = [
        { name: 'SPM', index: spmIndex },
        { name: 'PIN', index: pinIndex },
        { name: '模号', index: moldNoIndex },
        { name: '理论边角料', index: scrapIndex }
      ];
      moldConstraint = 'CM';
    } else if (suffixType === 'M') {
      // 注塑类（M）
      forbiddenFields = [
        { name: 'SPM', index: spmIndex },
        { name: 'PIN', index: pinIndex },
        { name: '电镀产速（m/min）', index: platingSpeedIndex },
        { name: '理论边角料', index: scrapIndex },
        { name: '用金量', index: goldAmountIndex },
        { name: '分类', index: categoryIndex }
      ];
      requiredFields = [
        { name: '模穴数', index: moldCavityIndex },
        { name: '注塑周期', index: injectionCycleIndex },
        { name: '产品塑胶重量', index: plasticWeightIndex },
        { name: '料头重量', index: sprueWeightIndex },
        { name: '模号', index: moldNoIndex }
      ];
      moldConstraint = 'SM';
    } else if (suffixType === 'P') {
      // 电镀类（P）
      forbiddenFields = [
        { name: 'SPM', index: spmIndex },
        { name: 'PIN', index: pinIndex },
        { name: '模穴数', index: moldCavityIndex },
        { name: '复制模模穴数', index: replMoldCavityIndex },
        { name: '注塑周期', index: injectionCycleIndex },
        { name: '复制注塑周期', index: replInjectionCycleIndex },
        { name: '复制模产品塑胶重量', index: replPlasticWeightIndex },
        { name: '复制模料头重量', index: replSprueWeightIndex },
        { name: '复制模回料百分比', index: replRecycleIndex },
        { name: '复制模号', index: replMoldNoIndex },
        { name: '产品塑胶重量', index: plasticWeightIndex },
        { name: '料头重量', index: sprueWeightIndex },
        { name: '回料百分比', index: recycleIndex },
        { name: '模号', index: moldNoIndex },
        { name: '用金量', index: goldAmountIndex },
        { name: '理论边角料', index: scrapIndex }
      ];
      requiredFields = [
        { name: '电镀产速（m/min）', index: platingSpeedIndex },
        { name: '电镀方式', index: platingMethodIndex },
        { name: '电镀类型', index: platingTypeIndex }
      ];
      moldConstraint = null;
    }

    // 禁止字段校验
    forbiddenFields.forEach(field => {
      if (field.index !== -1) {
        const value = row[field.index];
        if (value !== null && value !== undefined && String(value).trim() !== '') {
          errors.push({
            row: rowNumber,
            field: field.name,
            value: String(value).trim(),
            error: `物料倒数第二位为"${suffixType}"且最后一位为数字，此字段禁止填写`
          });
        }
      }
    });

    // 必填字段校验
    requiredFields.forEach(field => {
      if (field.index !== -1) {
        const value = row[field.index];
        if (value === null || value === undefined || String(value).trim() === '') {
          errors.push({
            row: rowNumber,
            field: field.name,
            value: '',
            error: `物料倒数第二位为"${suffixType}"且最后一位为数字，此字段必须填写`
          });
        }
      }
    });

    // 表五特殊联动校验
    // 规则1：SPM 和 步距联动（SPM 有值则步距必填）
    const spmRaw = spmIndex !== -1 ? (row[spmIndex] !== null && row[spmIndex] !== undefined ? String(row[spmIndex]).trim() : '') : '';
    const stepRaw = stepIndex !== -1 ? (row[stepIndex] !== null && row[stepIndex] !== undefined ? String(row[stepIndex]).trim() : '') : '';
    if (spmRaw && !stepRaw) {
      errors.push({
        row: rowNumber,
        field: '步距',
        value: '',
        error: 'SPM 已填写值，步距必须填写'
      });
    }

    // 规则2：电镀方式与电镀类型双向联动
    const methodRaw = platingMethodIndex !== -1 ? (row[platingMethodIndex] !== null && row[platingMethodIndex] !== undefined ? String(row[platingMethodIndex]).trim() : '') : '';
    const typeRaw = platingTypeIndex !== -1 ? (row[platingTypeIndex] !== null && row[platingTypeIndex] !== undefined ? String(row[platingTypeIndex]).trim() : '') : '';
    if ((methodRaw && !typeRaw) || (!methodRaw && typeRaw)) {
      errors.push({
        row: rowNumber,
        field: '电镀方式/电镀类型',
        value: `电镀方式=${methodRaw}, 电镀类型=${typeRaw}`,
        error: '电镀方式与电镀类型必须同时填写或同时留空'
      });
    }

    // 模号包含字符约束
    if (moldConstraint && moldNoIndex !== -1) {
      const moldNo = row[moldNoIndex];
      const moldNoStr = moldNo ? String(moldNo).trim() : '';
      if (!moldNoStr.includes(moldConstraint)) {
        errors.push({
          row: rowNumber,
          field: '模号',
          value: moldNoStr,
          error: `物料倒数第二位为"${suffixType}"且最后一位为数字，模号必须包含"${moldConstraint}"字符`
        });
      }
    }
  });

  return errors;
}

app.post('/api/upload', upload.single('file'), async (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ error: '未上传文件' });
    }

    const workbook = read(req.file.path, { type: 'file' });
    const allSheetResults = [];
    const sheetDataMap = new Map();

    workbook.SheetNames.forEach((sheetName, sheetIndex) => {
      const sheet = workbook.Sheets[sheetName];
      const data = utils.sheet_to_json(sheet, { header: 1 });
      
      if (data.length === 0) {
        return;
      }

      let matchedConfig = null;
      let sheetLabel = sheetName;

      for (const [key, config] of Object.entries(SHEETS_CONFIG)) {
        const headerResult = findHeaderRow(data, config.requiredHeaders);
        if (headerResult) {
          matchedConfig = { key, config, headerResult };
          sheetLabel = `(${key}) ${config.name}`;
          break;
        }
      }

      if (!matchedConfig) {
        allSheetResults.push({
          name: sheetName,
          rows: data.length,
          headers: [],
          data: [],
          errors: [{ row: 1, field: '表头', value: '', error: '未识别到此 sheet 的类型，无法匹配任何已知模板' }]
        });
        return;
      }

      const { config, headerResult } = matchedConfig;
      const { index: headerRowIndex, headers } = headerResult;
      
      const normalizedHeaders = headers.map(normalizeHeader);
      const normalizedRequired = config.requiredHeaders.map(normalizeHeader);
      
      const missingHeaders = config.requiredHeaders.filter((h, idx) => {
        const normalized = normalizedRequired[idx];
        return !normalizedHeaders.includes(normalized);
      });
      
      const sheetErrors = [];
      
      if (missingHeaders.length > 0) {
        sheetErrors.push({
          row: headerRowIndex + 1,
          field: '表头',
          value: '',
          error: `表头缺失字段：${missingHeaders.join(', ')}`
        });
      }

      const bodyData = data.slice(headerRowIndex + 1);
      const errors = validateSheetData(bodyData, headers, config, headerRowIndex + 2, matchedConfig.key);
      sheetErrors.push(...errors);

      const sheetData = {
        name: sheetLabel,
        key: matchedConfig.key,
        rows: bodyData.length,
        headers,
        data: bodyData.map(row => row.map(cell => cell || '')),
        errors: sheetErrors,
        headerRowIndex
      };

      
      
      allSheetResults.push(sheetData);
      sheetDataMap.set(matchedConfig.key, sheetData);
    });

    if (allSheetResults.length === 0) {
      return res.status(400).json({ 
        error: '文件为空或没有可识别的 sheet'
      });
    }

    const sheetOne = sheetDataMap.get('一');
    const sheetTwo = sheetDataMap.get('二');
    const sheetThree = sheetDataMap.get('三');
    const sheetFour = sheetDataMap.get('四');
    const sheetFive = sheetDataMap.get('五');

    // 各表关键字段重复校验
    const duplicateCheckConfigs = [
      { sheet: sheetOne, fieldName: '工程物料', errorSheet: sheetOne },
      { sheet: sheetTwo, fieldName: '物料号', errorSheet: sheetTwo },
      { sheet: sheetThree, fieldName: '工程物料', errorSheet: sheetThree },
      { sheet: sheetFour, fieldName: '制造物料', errorSheet: sheetFour },
      { sheet: sheetFive, fieldName: '物料', errorSheet: sheetFive }
    ];

    duplicateCheckConfigs.forEach(({ sheet, fieldName, errorSheet }) => {
      if (!sheet) return;
      const normalizedHeaders = sheet.headers.map(h => normalizeHeader(h));
      const fieldIndex = normalizedHeaders.findIndex(h => h === normalizeHeader(fieldName));
      if (fieldIndex === -1) return;

      const seen = new Map();
      sheet.data.forEach((row, idx) => {
        const val = row[fieldIndex];
        if (val === null || val === undefined || String(val).trim() === '') return;
        const valStr = String(val).trim();
        if (seen.has(valStr)) {
          seen.get(valStr).push(idx + sheet.headerRowIndex + 2);
        } else {
          seen.set(valStr, [idx + sheet.headerRowIndex + 2]);
        }
      });

      seen.forEach((rows, val) => {
        if (rows.length > 1) {
          rows.forEach(rowNum => {
            errorSheet.errors.push({
              row: rowNum,
              field: fieldName,
              value: val,
              error: `${fieldName}存在重复填写`
            });
          });
        }
      });
    });

    // 表二内部校验：计划数据计划员与订货数据计划员一致
    if (sheetTwo) {
      const sheetTwoErrors = validateSheetTwoInternal(sheetTwo.data, sheetTwo.headers, sheetTwo.headerRowIndex + 2);
      sheetTwo.errors.push(...sheetTwoErrors);
    }

    // 表三特殊校验：组件与是否虚拟的关系
    if (sheetThree) {
      const sheetThreeErrors = validateSheetThreeSpecial(sheetThree.data, sheetThree.headers, sheetThree.headerRowIndex + 2);
      sheetThree.errors.push(...sheetThreeErrors);
    }

    // 跨表校验：表一工程物料必须出现在表二、三、四、五
    if (sheetOne) {
      const crossTypeErrors = validateCrossSheetsTypeConsistency(sheetOne, sheetTwo);
      if (crossTypeErrors.length > 0) {
        sheetOne.errors.push(...crossTypeErrors);
      }

      const crossSheetErrors = validateCrossSheetsExistence(sheetOne, sheetTwo, sheetThree, sheetFour, sheetFive);
      if (crossSheetErrors.length > 0) {
        sheetOne.errors.push(...crossSheetErrors);
      }
    }

    if (sheetOne && sheetTwo) {
      const crossErrors = validateCrossSheets(sheetOne, sheetTwo);
      if (crossErrors.length > 0) {
        sheetTwo.errors.push(...crossErrors);
      }
    }

    // 表四特殊校验：制造物料后两位与工艺流程说明的匹配规则
    if (sheetFour) {
      const sheetFourErrors = validateSheetFourSpecial(sheetFour.data, sheetFour.headers, sheetFour.headerRowIndex + 2);
      sheetFour.errors.push(...sheetFourErrors);
    }

    // 表五特殊校验
    if (sheetFive) {
      const sheetFiveErrors = validateSheetFiveSpecial(sheetFive.data, sheetFive.headers, sheetFive.headerRowIndex + 2);
      sheetFive.errors.push(...sheetFiveErrors);
    }

    // 表四表五跨表 SPM 校验
    if (sheetFour && sheetFive) {
      const crossSPMErrors = validateCrossSheetsSPM(sheetFour, sheetFive);
      if (crossSPMErrors.length > 0) {
        sheetFour.errors.push(...crossSPMErrors);
      }
    }

    // 表三与表五跨表校验：组件后两位为R1/R2时必填回料百分比
    if (sheetThree && sheetFive) {
      const sheetThreeToFiveErrors = validateSheetThreeToSheetFive(sheetThree, sheetFive);
      if (sheetThreeToFiveErrors.length > 0) {
        sheetFive.errors.push(...sheetThreeToFiveErrors);
      }
    }

    const totalErrors = allSheetResults.reduce((sum, s) => sum + s.errors.length, 0);
    const affectedSheets = allSheetResults.filter(s => s.errors.length > 0).length;

    res.json({
      success: true,
      sheets: allSheetResults,
      summary: {
        totalErrors,
        affectedSheets,
        totalSheets: allSheetResults.length
      }
    });

  } catch (error) {
    console.error('Upload error:', error);
    res.status(500).json({ error: '文件处理失败：' + error.message });
  }
});

app.get('*', (req, res) => {
  res.sendFile(join(__dirname, 'dist', 'index.html'));
});

const PORT = process.env.PORT || 3001;

app.listen(PORT, () => {
  console.log(`Server running on http://localhost:${PORT}`);
});
