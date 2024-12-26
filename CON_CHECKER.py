import pandas as pd
import re


def parse_lp_file(lp_filename):
    constraints = []
    with open(lp_filename, 'r', encoding='utf-8') as file:
        lines = file.readlines()

    in_subject_to = False
    current_constraint = None

    for line in lines:
        # 去除注释和多余的空白字符
        line = line.split('\\')[0].strip()
        if not line:
            continue

        # 检测不同的部分
        if re.match(r'^(Minimize|Maximize)$', line, re.IGNORECASE):
            in_subject_to = False
            continue
        elif re.match(r'^Subject\s+To$', line, re.IGNORECASE) or re.match(r'^SubjectTo$', line, re.IGNORECASE):
            in_subject_to = True
            continue
        elif re.match(r'^(Bounds|End)$', line, re.IGNORECASE):
            in_subject_to = False
            continue

        if in_subject_to:
            # 处理跨多行的约束
            if ':' in line:
                # 新的约束
                constraint_name, expression = line.split(':', 1)
                current_constraint = {
                    'name': constraint_name.strip(),
                    'expression': expression.strip()
                }
                constraints.append(current_constraint)
            else:
                # 约束的延续
                if current_constraint:
                    current_constraint['expression'] += ' ' + line.strip()

    # 解析每个约束表达式
    parsed_constraints = []
    for constraint in constraints:
        name = constraint['name']
        expr = constraint['expression']

        # 将表达式分为左侧和右侧
        match = re.match(r'(.+?)(<=|>=|=)(.+)', expr)
        if not match:
            print(f"无法解析约束：{name}")
            continue
        lhs, sense, rhs = match.groups()
        lhs = lhs.strip()
        sense = sense.strip()
        rhs = rhs.strip()

        # 解析左侧的各项
        # 处理像 "2 x[t0,p0,0]" 或 "-delta[0]" 这样的项
        terms = re.finditer(r'([+-]?\s*\d*\.?\d*)\s*([a-zA-Z_][\w\[\],]*)', lhs)
        coefficients = {}
        for term in terms:
            coef_str, var = term.groups()
            coef_str = coef_str.replace(' ', '')
            if coef_str in ['', '+']:
                coef = 1.0
            elif coef_str == '-':
                coef = -1.0
            else:
                try:
                    coef = float(coef_str)
                except ValueError:
                    print(f"约束 '{name}' 中的系数 '{coef_str}' 无效")
                    coef = 1.0
            coefficients[var] = coefficients.get(var, 0) + coef

        # 解析右侧的值
        try:
            rhs_value = float(rhs)
        except ValueError:
            print(f"约束 '{name}' 中的右侧值 '{rhs}' 无效")
            rhs_value = 0.0

        parsed_constraints.append({
            'name': name,
            'coefficients': coefficients,
            'sense': sense,
            'rhs': rhs_value
        })

    return parsed_constraints


def read_initial_solution(excel_filename):
    df = pd.read_excel(excel_filename)
    if 'K' not in df.columns or 'V' not in df.columns:
        raise ValueError("Excel文件必须包含 'K' 和 'V' 两列。")
    solution = dict(zip(df['K'], df['V']))
    return solution


def evaluate_constraints(constraints, solution):
    all_satisfied = True
    for constraint in constraints:
        lhs_value = 0.0
        for var, coef in constraint['coefficients'].items():
            var_value = solution.get(var, 0.0)
            lhs_value += coef * var_value
        sense = constraint['sense']
        rhs = constraint['rhs']
        satisfied = False
        if sense == '=':
            satisfied = abs(lhs_value - rhs) < 1e-5
        elif sense == '<=':
            satisfied = lhs_value <= rhs + 1e-5
        elif sense == '>=':
            satisfied = lhs_value >= rhs - 1e-5
        else:
            print(f"约束 '{constraint['name']}' 中存在未知的关系符号 '{sense}'")

        if not satisfied:
            all_satisfied = False
            print(f"约束 '{constraint['name']}' 未被满足：")
            print(constraint)
            print(f"  左侧值 = {lhs_value}, 关系符号 = '{sense}', 右侧值 = {rhs}")
        else:
            print(f"约束 '{constraint['name']}' 被满足。")

    if all_satisfied:
        print("\n所有约束都被初始解满足。")
    else:
        print("\n部分约束未被初始解满足。")


def main():
    excel_filename = 'sol.xlsx'  # Replace with your actual Excel file name
    lp_filename = 'model162.lp'  # Replace with your actual LP file name

    print("正在解析 LP 文件...")
    constraints = parse_lp_file(lp_filename)
    print(f"解析到的约束总数：{len(constraints)}\n")

    print("正在读取 Excel 中的初始解...")
    solution = read_initial_solution(excel_filename)
    print(f"初始解中的变量总数：{len(solution)}\n")

    print("正在评估约束条件...")
    evaluate_constraints(constraints, solution)


if __name__ == "__main__":
    main()
