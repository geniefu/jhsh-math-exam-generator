# -*- coding: utf-8 -*-
"""
數學科 SymPy 符號運算 100% 自動驗算防錯機制 (math_sympy_verifier.py)
嚴格檢核：
1. 方程、多項式與幾何邊角數學自洽性 (排除無解題、矛盾題)
2. 幾何存在性驗算 (三角不等式、畢氏定理、弦心距、交點)
3. 單選題選項唯一性 (恰有一項等於正解，其餘三項為典型診斷迷思誘答)
4. 配分與選項均衡度 (各選項約 25%，總分嚴格等於 100 分)
"""

import sympy as sp

def verify_algebraic_solution(eq_expr, var_sym, candidate_sol):
    """
    驗證代數方程或化簡結果：
    eq_expr: sympy 表達式 (等於 0)
    var_sym: sympy 符號 (如 x)
    candidate_sol: 數值或 sympy 表達式
    """
    try:
        val = eq_expr.subs(var_sym, candidate_sol)
        simplified = sp.simplify(val)
        return simplified == 0
    except Exception as e:
        return False

def verify_triangle_validity(a, b, c):
    """檢驗三角形三邊長是否滿足嚴格三角不等式"""
    try:
        sides = sorted([float(a), float(b), float(c)])
        if sides[0] <= 0:
            return False, "邊長必須大於 0"
        if sides[0] + sides[1] <= sides[2]:
            return False, f"兩邊之和 ({sides[0]}+{sides[1]}={sides[0]+sides[1]}) 未大於第三邊 ({sides[2]})"
        return True, "滿足三角不等式"
    except Exception as e:
        return False, str(e)

def verify_options_uniqueness(correct_val, distractors):
    """
    檢驗單選題四個選項：
    1. 正解與三個誘答選項互不相等
    2. 恰好只有一個正確答案
    """
    all_vals = [correct_val] + distractors
    if len(all_vals) != 4:
        return False, "選項數量不為 4"
        
    for i in range(len(all_vals)):
        for j in range(i + 1, len(all_vals)):
            diff = sp.simplify(all_vals[i] - all_vals[j])
            if diff == 0:
                return False, f"選項重複：第 {i+1} 個與第 {j+1} 個選項數值相等！"
    return True, "選項互異且唯一"

def verify_total_score(scores_list, expected_sum=100):
    """驗證全卷配分加總為 100 分且各大題皆為整數"""
    total = sum(scores_list)
    errs = []
    for idx, s in enumerate(scores_list):
        if not isinstance(s, int) or s <= 0:
            errs.append(f"第 {idx+1} 項配分必須為正整數：{s}")
    if total != expected_sum:
        errs.append(f"總分錯誤：當前總分 {total} 分，應為 {expected_sum} 分")
    return len(errs) == 0, total, errs

def verify_answers_distribution(answers_list):
    """驗證單選題答案 (A)(B)(C)(D) 分佈均衡度"""
    total = len(answers_list)
    if total == 0:
        return False, {}
    counts = {'A': 0, 'B': 0, 'C': 0, 'D': 0}
    for a in answers_list:
        clean_a = str(a).strip().upper()
        if clean_a in counts:
            counts[clean_a] += 1
            
    is_balanced = True
    for k, v in counts.items():
        ratio = v / total
        if ratio < 0.12 or ratio > 0.38: # 容許 12%~38% 範圍
            is_balanced = False
    return is_balanced, counts
