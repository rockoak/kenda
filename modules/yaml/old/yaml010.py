import os
import re
import sys
from pathlib import Path

# 添加 lib 目錄到路徑
sys.path.insert(0, str(Path(__file__).parent.parent.parent / 'lib'))
import mainlib

# Define input and output directories（使用專案根目錄）
project_root = mainlib.get_project_root()
input_dir = project_root / 'input_yamls'
output_dir = project_root / 'output_yamls'

# Create output directory if it doesn't exist
os.makedirs(output_dir, exist_ok=True)

# Check if input directory exists
if not input_dir.exists():
    print(f"錯誤：找不到輸入目錄 {input_dir}")
    print(f"請確認 {input_dir} 目錄存在並包含 .yaml 檔案")
    sys.exit(1)

# Process each .yaml file in the input directory
for filename in os.listdir(input_dir):
    if filename.endswith('.yaml'):
        input_path = os.path.join(input_dir, filename)
        output_path = os.path.join(output_dir, filename)

        with open(input_path, 'r', encoding='utf-8') as infile:
            lines = infile.readlines()

        # Process lines according to the specified rules
        new_lines = lines[:]
        for i in range(len(lines)):
            line = lines[i]
            if 'unit: "mm"' in line:
                new_lines[i] = line.replace('unit: "mm"', 'unit: "m"')
                if i >= 2:
                    mid_line = lines[i - 2]
                    match = re.search(r'mid:\s*(\"?\d+\.?\d*\"?)', mid_line)
                    if match:
                        original = match.group(1)
                        has_quotes = original.startswith('"') and original.endswith('"')
                        number = float(original.strip('"')) / 1000
                        new_value = f'"{number}"' if has_quotes else f'{number}'
                        new_lines[i - 2] = re.sub(r'mid:\s*(\"?\d+\.?\d*\"?)', f'mid: {new_value}', mid_line)

        # Write the modified content to the output file
        with open(output_path, 'w', encoding='utf-8') as outfile:
            outfile.writelines(new_lines)
