(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.ReconcileProfiles = factory();
})(globalThis, function () {
  'use strict';

  const material = ['料品', '物料代码', '零件代码', '物料编码'];
  const materialCode = ['物料代码', '料品', '零件代码', '物料编码'];
  const bom = ['BOM号', 'BOM编号', 'BOM'];
  const task = ['任务代码'];

  const general = {
    id: 'general',
    label: '通用',
    headerStrategy: 'first-nonempty',
    keyPresets: []
  };

  const manufacturing = {
    id: 'manufacturing',
    label: '生产物料',
    headerStrategy: 'keywords',
    headerKeywords: [
      '序号', '料品', '物料代码', '零件代码', '集团项目号', '任务代码', '设计批号',
      '受益项目', '项目号', '项目编号', '物料编码', '物料名称', '规格型号', '图号',
      '数量', '单号', '行号', '编码', '代码', '名称', 'ID', 'BOM', '批次', '批号'
    ],
    strongKeywords: ['序号', '料品', '物料代码', '零件代码', '集团项目号', '任务代码', '设计批号', 'BOM', '受益项目'],
    keyPresets: [
      { label: '设计批号 + 任务代码 + 料品', groups: [['设计批号'], task, material] },
      { label: '任务代码 + BOM号 + 物料代码', groups: [task, bom, materialCode] },
      { label: '受益项目 + 料品', groups: [['受益项目号', '受益项目', '集团项目号', '项目编号', '项目号'], material] },
      { label: '任务代码 + 料品', groups: [task, material] },
      { label: 'BOM号 + 物料代码', groups: [bom, materialCode] },
      { label: '料品', groups: [material] },
      { label: '序号', groups: [['序号']] }
    ]
  };

  return { general, manufacturing };
});
