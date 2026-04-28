import os
import yaml
import json
from datetime import datetime
from excel_handler import ExcelHandler

CONFIG_PATH = os.path.join(os.path.dirname(__file__), '../config/fields.yaml')
TEMPLATE_PATH = '/Users/shawnpxf/Downloads/QC_SCM_GOODS_L1_IMPORT.xlsx'
OUTPUT_DIR = '/Users/shawnpxf/Documents/trae_projects/upload2.0/openclaw-excel-skill/output'

class ExcelSkill:
    def __init__(self):
        self.config = self.load_config()
        self.handler = None
        self.data = {}
        self.defaults = {}
    
    def load_config(self):
        with open(CONFIG_PATH, 'r', encoding='utf-8') as f:
            return yaml.safe_load(f)
    
    def init_handler(self):
        self.handler = ExcelHandler(TEMPLATE_PATH, self.config)
        self.handler.load_template()
        self.defaults = self.handler.get_default_values()
    
    def ask_field(self, field):
        name = field['name']
        required = field.get('required', False)
        description = field.get('description', '')
        field_type = field.get('type', 'text')
        options = field.get('options', [])
        default_value = field.get('default', self.defaults.get(name, ''))
        
        prompt = f"请输入{name}"
        if description:
            prompt += f"（{description}）"
        if required:
            prompt += " *必填"
        
        if default_value:
            prompt += f"\n[默认: {default_value}]"
            prompt += "\n直接回车使用默认值，或输入新值："
        else:
            prompt += "\n请输入值："
        
        if field_type == 'select' and options:
            options_str = "\n".join([f"{i+1}. {opt}" for i, opt in enumerate(options)])
            prompt += f"\n选项：\n{options_str}\n请输入序号或直接输入值："
        
        while True:
            user_input = input(prompt + "\n").strip()
            
            if not user_input:
                if required:
                    if default_value:
                        print(f"使用默认值: {default_value}")
                        return default_value
                    print("此字段为必填项，请输入内容")
                    continue
                return default_value if default_value else None
            
            if field_type == 'select' and options:
                if user_input.isdigit():
                    idx = int(user_input) - 1
                    if 0 <= idx < len(options):
                        return options[idx]
                if user_input in options:
                    return user_input
                print(f"输入不在选项列表中，请重新输入")
                continue
            
            if field_type == 'float':
                try:
                    return float(user_input)
                except ValueError:
                    print("请输入有效的数字")
                    continue
            
            if field_type == 'int':
                try:
                    return int(user_input)
                except ValueError:
                    print("请输入有效的整数")
                    continue
            
            if field_type == 'bool':
                if user_input in ['是', 'true', 'True', '1']:
                    return '是'
                elif user_input in ['否', 'false', 'False', '0']:
                    return '否'
                print("请输入'是'或'否'")
                continue
            
            return user_input
    
    def collect_data(self):
        fields = self.config['fields']
        self.data = {}
        
        print("=" * 50)
        print("欢迎使用商品数据Excel填写工具")
        print("=" * 50)
        
        if self.defaults:
            print(f"\n提示：检测到模板默认值，将自动填充未指定的字段")
        
        for section, section_fields in fields.items():
            print(f"\n--- {self.get_section_name(section)} ---")
            for field in section_fields:
                value = self.ask_field(field)
                if value is not None:
                    self.data[field['name']] = value
        
        return self.data
    
    def get_section_name(self, section):
        names = {
            'basic': '基本信息',
            'specification': '规格信息',
            'business': '业务设置',
            'print': '打印设置'
        }
        return names.get(section, section)
    
    def generate_output_path(self):
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        filename = f'商品数据_{timestamp}.xlsx'
        return os.path.join(OUTPUT_DIR, filename)
    
    def execute(self):
        self.init_handler()
        self.collect_data()
        
        if self.data:
            self.handler.write_data(self.data)
            output_path = self.generate_output_path()
            self.handler.save(output_path)
            print(f"\n✓ Excel文件已成功保存：")
            print(f"  {output_path}")
            return output_path
        else:
            print("未收集到任何数据，操作已取消")
            return None
    
    def add_more_items(self):
        while True:
            choice = input("\n是否继续添加商品？(是/否) ").strip()
            if choice in ['是', 'yes', 'y']:
                self.data = {}
                self.collect_data()
                if self.data:
                    self.handler.write_data(self.data)
                    print("✓ 商品已添加")
            elif choice in ['否', 'no', 'n']:
                output_path = self.generate_output_path()
                self.handler.save(output_path)
                print(f"\n✓ Excel文件已成功保存：")
                print(f"  {output_path}")
                return output_path
            else:
                print("请输入'是'或'否'")

def main():
    skill = ExcelSkill()
    output_path = skill.execute()
    
    if output_path:
        skill.add_more_items()

if __name__ == '__main__':
    main()