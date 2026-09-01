# 稳定 Excel 渲染输入

AI 完成全文提取和校验后，把结果整理成以下 JSON，再交给`../scripts/build_project_snapshot.py`一次性生成工作簿。JSON 是内部中间数据，不是交付物。

```json
{
  "overview": {
    "project_name": "",
    "project_number": "",
    "one_sentence_summary": "",
    "purchaser": "",
    "procurement_agent": "",
    "procurement_method": "",
    "award_count": "",
    "bid_deadline": "",
    "opening_location": "",
    "source": ""
  },
  "qualifications": [
    {
      "question": "",
      "check_material": "",
      "user_choice": "",
      "consequence": "",
      "source": "",
      "notes": ""
    }
  ],
  "project_requirements": [
    {
      "component": "",
      "requirement": "",
      "recipient": "",
      "time": "",
      "location": "",
      "scale": "",
      "standard_or_deliverable": "",
      "acceptance": "",
      "source": ""
    }
  ],
  "contract_comparison": {
    "contract_present": true,
    "comparison_completed": true,
    "contract_source": "",
    "template_note": "",
    "template_note_source": "",
    "items": [
      {
        "subject": "",
        "requirement_text": "",
        "contract_text": "",
        "result": "一致或明确冲突或疑似冲突或合同未覆盖或项目需求未覆盖或待确认",
        "impact": "",
        "confirmation_question": "",
        "requirement_source": "",
        "contract_source": "",
        "manual_status": "待确认"
      }
    ]
  },
  "quote_restrictions": [
    {
      "scope": "",
      "restriction_type": "",
      "requirement": "",
      "tax_requirement": "",
      "quote_carrier": "",
      "consequence": "",
      "source": ""
    }
  ]
}
```

规则：

- 所有数组只能包含对象；没有记录时使用空数组。
- 正式渲染时`user_choice`必须为空字符串；`是/否`只允许用户在交付后的工作簿中选择。含预填值的输入会被渲染器拒绝。
- 不得把报名、购标或文件获取材料放入`qualifications`。
- `contract_comparison`必须存在。没有合同模板时，`contract_present`和`comparison_completed`均为`false`，`items`为空，`contract_source`写“全文未发现合同模板”。
- 有合同模板时，`contract_present`和`comparison_completed`必须为`true`，`contract_source`必须注明位置，`items`至少有一项；否则渲染器拒绝生成 Excel。
- 有合同模板时，每个`items`记录都必须包含事项、双方原文、双方位置和规定的核对结果。未覆盖的一方使用“未说明”，不得留空。
- 未发现合同模板适用说明时，`template_note`写“未说明”，`template_note_source`写明已经全文检索但未发现；找到说明时两个字段都必须填写。
- 正式渲染时`manual_status`必须为`待确认`；已确认状态只允许用户在交付后的工作簿中选择。含预填确认结果的输入会被渲染器拒绝。
- 一致项必须保留在内部`items`中作为完成比对的证据，但不会展开到 Excel；只有非`一致`项才显示异常明细。
