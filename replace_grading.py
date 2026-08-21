# -*- coding: utf-8 -*-
import os

workspace = r'C:\Users\32891\.qclaw\workspace'
file_path = os.path.join(workspace, '审稿意见_2026_5_批次.md')
new_grading_path = os.path.join(workspace, 'grading_new.txt')

with open(file_path, 'r', encoding='utf-8') as f:
    content = f.read()

with open(new_grading_path, 'r', encoding='utf-8') as f:
    new_grading = f.read()

start_marker = '\n## 按核心期刊标准的分级评定'
end_marker = '\n## 审稿总结'

start_idx = content.find(start_marker)
end_idx = content.find(end_marker)

if start_idx == -1:
    print('ERROR: start_marker not found')
elif end_idx == -1:
    print('ERROR: end_marker not found')
else:
    print('Found start at %d, end at %d' % (start_idx, end_idx))
    new_content = content[:start_idx] + '\n' + new_grading + '\n' + content[end_idx:]
    with open(file_path, 'w', encoding='utf-8') as f:
        f.write(new_content)
    print('SUCCESS: File updated.')
